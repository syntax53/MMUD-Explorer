Attribute VB_Name = "modReachability"
Option Explicit

' Monster reachability by character level, for the Monsters tab filter
' "Hide monsters I can't reach at my level" (frmMonsterFilters).
'
' For every monster: in which level bands can a room it appears in be reached
' from Bank of Godfrey (1/297)? Only LEVEL gates count; keys, items, cash, quests,
' class/race/alignment are treated as met. The answer depends only on the database,
' so it's computed once per database (on first use) and kept until another loads.
'
' Port of tools/mudroute (build.py + reachability.py); design and the game
' mechanics behind each edge type are in docs/room-routing.md. In dev mode a CSV
' (_reachability.csv) is written in the same format as
'   python tools/mudroute/reachability.py --mdb <db> --no-overrides --csv <file>
' so the two can be diffed line by line.
'
' Edges (from -> to, with a level range):
'   exits (Rooms.N..D, "(Level: a to b)", b=0 means no upper limit)
'   room commands (Rooms.CMD textblock lines), room spells (Rooms.Spell),
'   NPC conversations (Monsters.GreetTXT keyword lines, from the NPC's rooms),
'   teleport spells that only work where a room item is ("roomitem N")
' Textblock chains follow text/bare-number jumps, LinkTo, cast -> spell chains
' (140/141 teleport, 151 EndCast, 148 TextBlock). A 'random' block whose rolls go
' to more than one place is random (only usable when all destinations are in one
' strongly connected area); with more than 20 places it's dropped.
'
' Routes (ReachRoute, for the detail pane) use RouteSearch rather than the BFS: fewest coin-costing
' steps first (toll exits, chains with 'price N'), then fewest steps; optionally ignoring level gates.
' When a shorter route costs coin and the coin-avoiding one is more than ROUTE_DETOUR_MAX times its length,
' the shorter one is used; the other is offered as the alternate route (ReachRoute bAltRoute).
' Monster rooms carry a category (RCAT_*: Room < Group(lair) < resolved summon < Group) so a route
' prefers a fixed spawn room. Reference fields cut off at 250 chars end in '+' ("Room 14/62+" is
' really 14/62xx); such truncated references are skipped.

Private Const START_MAP As Long = 1
Private Const START_ROOM As Long = 297
Private Const LEVEL_CAP As Long = 255     'a max of 255/999/9999 means "no upper limit"
Private Const NO_MAX As Long = 9999
Private Const MAX_CHAIN_DEPTH As Long = 10
Private Const MAX_RANDOM_DESTS As Long = 20
Private Const MAX_RESOLVE_DEPTH As Long = 8
Private Const MAX_ABILS As Long = 20

Private Const ABIL_TELEPORT_ROOM As Long = 140
Private Const ABIL_TELEPORT_MAP As Long = 141
Private Const ABIL_TEXTBLOCK As Long = 148
Private Const ABIL_ENDCAST As Long = 151

Private Const OP_NONE As Long = 0
Private Const OP_TELEPORT As Long = 1
Private Const OP_TEXT As Long = 2
Private Const OP_RANDOM As Long = 3
Private Const OP_CAST As Long = 4
Private Const OP_MINLEVEL As Long = 5
Private Const OP_MAXLEVEL As Long = 6
Private Const OP_ROOMITEM As Long = 7
Private Const OP_OTHER As Long = 8

Private Const MON_NOT_IN_GAME As Byte = 1
Private Const MON_NO_LOCATION As Byte = 2
Private Const MON_LOCATED As Byte = 3

Public Enum ReachStatus
    reachUnknown = 0    'not computed, unknown monster, or monster not located / not in game (callers show nothing)
    reachYes = 1
    reachNo = 2         'unreachable; gate info filled in
    reachNoGate = 3     'unreachable but no level gate found on any path (no known path)
End Enum

'Step kind flags (bitmask) for route steps
Public Const REACH_STEP_ACTION As Long = 1    'not a plain exit: typed command / NPC talk / room spell / item spell
Public Const REACH_STEP_LEVEL As Long = 2     'edge has a level gate (min > 1 or max < no-limit)
Public Const REACH_STEP_RESTRICT As Long = 4  'exit with a non-level restriction (Key/Item/Toll/Door/Hidden/Text/...)
Public Const REACH_STEP_MAPCHANGE As Long = 8 'destination room is on a different map than the previous room
Public Const REACH_STEP_RANDOM As Long = 16   'random-destination teleport group
Public Const REACH_STEP_COIN As Long = 32     'step costs coin: exit toll or a command/spell chain 'price N'

'edge kinds (eKind)
Private Const EK_EXIT As Byte = 1
Private Const EK_COMMAND As Byte = 2
Private Const EK_ROOMSPELL As Byte = 3
Private Const EK_NPC As Byte = 4
Private Const EK_ITEMSPELL As Byte = 5
Private Const NO_DIR As Byte = 255

Private Const NUM_POIS As Long = 6        'route sources, see ResolvePOIs

'monster room categories (mRooms/mBase values): lower = preferred route destination
Private Const RCAT_ROOM As Byte = 1       '[Summoned By] "Room m/r"
Private Const RCAT_LAIR As Byte = 2       '[Summoned By] "Group(lair): m/r"
Private Const RCAT_RESOLVED As Byte = 3   'resolved from a summoning textblock/spell/monster
Private Const RCAT_GROUP As Byte = 4      '[Summoned By] "Group: m/r"

Private Const ROUTE_COIN_COST As Long = 1000000   'RouteSearch: one coin-costing step outweighs any step count
Private Const ROUTE_INF As Long = 2147483647
Private Const ROUTE_DETOUR_MAX As Long = 3        'ReachRoute: a coin-free route over 3x a shorter coin route's steps loses to it

Private Type tTeleOut
    Room As Long
    Map As Long           '0 = same map as the room the chain started in
    MinLvl As Long
    MaxLvl As Long
    IsRandom As Boolean
    RoomItems As String   '",id,id," -- 'roomitem' requirements seen on the way
    LineNo As Long        'top-level textblock line that produced it
    Cmd As String         'top-level command / greet keyword that produced it ("" for spells)
    Price As Long         'highest 'price N' seen on the way (0 = free)
End Type

Private sReachDatabase As String          'database file the results belong to ("" = none)
Private sReachFailedFor As String         'don't retry a failed computation for the same database
Private bReachComputing As Boolean        'ReachEnsureComputed is running (its DoEvents can re-enter via FilterMonsters)
Private bBuildError As Boolean            'a build step's error handler ran: the results are not usable
Private nDroppedRandom As Long

'rooms (1-based index)
Private dRoomIdx As Dictionary            '"map/room" -> index
Private nRooms As Long
Private rMap() As Long, rRoom() As Long, rCMD() As Long, rSpell() As Long, rNPC() As Long
Private rExit() As String                 '(room, 0..9) raw exit fields
Private rPlaced() As String
Private sDirs As Variant                  'exit direction names for Rooms columns 0..9 (kept for route labels)

'textblocks
Private dTB As Dictionary                 'number -> index
Private tbAction() As String, tbLink() As Long, tbCalled() As String

'spells
Private dSpell As Dictionary              'number -> index
Private spNum() As Long, spAb() As Long, spVal() As Long, spReqLevel() As Long, spCastedBy() As String
Private nSpells As Long

'items: rooms holding each item (Rooms.Placed, Items.[Obtained From] "Room m/r")
Private dItemRooms As Dictionary          'item number -> Collection of room indexes

'monsters
Private dMonIdx As Dictionary             'number -> index
Private nMons As Long
Private mNum() As Long, mName() As String, mSummonedBy() As String, mGreet() As Long
Private mStatus() As Byte
Private mRooms() As Dictionary            'final rooms (index -> RCAT_* category Byte)
Private mBase() As Dictionary             'rooms from [Summoned By] Room/Group entries only (index -> RCAT_*)
Private mBaseRaw() As Long                'count of those entries, including rooms not in Rooms

'edges
Private nEdges As Long
Private eFrom() As Long, eTo() As Long, eMin() As Long, eMax() As Long, eGroup() As Long
Private eKind() As Byte                   'EK_* (for ReachRoute)
Private eDir() As Byte                    'exit direction 0..9, NO_DIR for non-exits
Private eLabel() As String                'exit: non-level restriction text ("" = plain); command: command text;
                                          'room spell: spell number; NPC: "npc|keyword";
                                          'item spell: "spell#|item#,item#" (numbers, so the UI can name them)
Private eCoin() As Long                   'coin the step costs: exit "(Toll: N)" / chain 'price N' (0 = free)
Private nGroups As Long

'chain-walk output buffer and per-root cache
Private tOut() As tTeleOut, nOut As Long
Private tPool() As tTeleOut, nPool As Long
Private dChainCache As Dictionary         'key -> "start,count" into tPool

'traversal (kept after computing, for ReachMonsterInfo's gate tracing)
Private nStartIdx As Long                 'Bank of Godfrey room index (0 = not in this database)
Private adjStart() As Long, adjEdge() As Long     'CSR over non-random edges, keyed on eFrom
Private grpOK() As Boolean                'random group usable at all (destinations in one SCC)
Private grpSrc() As Long                  'random group's source room index
Private bRevAdjBuilt As Boolean
Private radjStart() As Long, radjEdge() As Long   'CSR over ALL edges, keyed on eTo (built lazily)
Private bGrpAdjBuilt As Boolean
Private gadjStart() As Long, gadjEdge() As Long   'CSR over random-group edges, keyed on eFrom (lazily, RouteSearch)
Private nVisCacheLevel As Long            'VisitAtLevel cache: level (0 = empty) and its visited() result
Private visCache() As Long
Private nPOIIdx() As Long                 '(0..NUM_POIS-1) route source room index (0 = not in this database)
Private sPOIName() As String              'route source names; index 0 = Bank of Godfrey (= nStartIdx)

'results
Private nBP As Long
Private bpLevel() As Long
Private mHit() As Byte                    '(monster, breakpoint) 1 = reachable


'=============================================================== public

Public Sub ReachInvalidate()
sReachDatabase = ""
sReachFailedFor = ""
End Sub

'Drop the results and compute them again now (shows the progress bar).
Public Sub ReachRecompute()
Call ReachInvalidate
Call ReachEnsureComputed
End Sub

Public Function ReachIsComputed() As Boolean
ReachIsComputed = (sReachDatabase <> "" And sReachDatabase = sCurrentDatabaseFile)
End Function

'True if the monster should be shown at this level: reachable now, or its location
'is unknown (never hide what we can't place), or it's not in the game (other
'filters decide). False only when it's located but no room of it is reachable.
Public Function ReachIsMonsterOK(ByVal nMonsterNum As Long, ByVal nLevel As Long) As Boolean
Dim nIdx As Long, i As Long
On Error GoTo error:

ReachIsMonsterOK = True
If Not ReachIsComputed() Then Call ReachEnsureComputed
If Not ReachIsComputed() Then Exit Function
If Not dMonIdx.Exists(nMonsterNum) Then Exit Function
nIdx = dMonIdx(nMonsterNum)
If mStatus(nIdx) <> MON_LOCATED Then Exit Function

If nLevel < 1 Then nLevel = 1
i = 0
Do While i + 1 < nBP
    If bpLevel(i + 1) > nLevel Then Exit Do
    i = i + 1
Loop
ReachIsMonsterOK = (mHit(nIdx, i) = 1)

Exit Function
error:
Call HandleError("ReachIsMonsterOK")
ReachIsMonsterOK = True
End Function

'Reachability of one monster at a level, plus the closest level gate that blocks it.
'bComputeIfNeeded: False = never trigger the (progress-bar) computation; return reachUnknown if not computed.
'On reachNo: nGateMap/nGateRoom = the reachable room holding the gated exit/command; nDestMap/nDestRoom = the room
'it leads to; nMinLvl/nMaxLvl = the edge's level range (nMaxLvl = 0 when no upper limit) -- for a random-teleport
'group the range every edge of the group allows; bRandomGroup = True when the gate is a random-teleport group that
'is unusable at any level (scattered destinations) rather than a level range.
'On reachNo/reachNoGate: nReachableAtLvl = the lowest level above nLevel at which the monster IS reachable (from the
'precomputed level bands; 0 = none). Nested gates make this higher than nMinLvl: the gate reported is only the
'nearest one, so "needs Level 30+" may still leave a Level 60+ gate behind it.
'The gate is found by a reverse BFS from the monster's rooms over all edges (levels ignored), through rooms
'not reachable at this level, stopping at the first edge whose source room IS reachable (fewest hops).
Public Function ReachMonsterInfo(ByVal nMonsterNum As Long, ByVal nLevel As Long, ByVal bComputeIfNeeded As Boolean, _
    ByRef nGateMap As Long, ByRef nGateRoom As Long, ByRef nDestMap As Long, ByRef nDestRoom As Long, _
    ByRef nMinLvl As Long, ByRef nMaxLvl As Long, ByRef bRandomGroup As Boolean, ByRef nReachableAtLvl As Long) As ReachStatus
Dim nIdx As Long, i As Long, r As Long, nGate As Long, g As Long, e As Long, nLo As Long, nHi As Long
Dim visited() As Long, vKey As Variant
On Error GoTo error:

ReachMonsterInfo = reachUnknown
nGateMap = 0: nGateRoom = 0: nDestMap = 0: nDestRoom = 0
nMinLvl = 0: nMaxLvl = 0: bRandomGroup = False: nReachableAtLvl = 0

If Not ReachIsComputed() And bComputeIfNeeded Then Call ReachEnsureComputed
If Not ReachIsComputed() Then Exit Function
If Not dMonIdx.Exists(nMonsterNum) Then Exit Function
nIdx = dMonIdx(nMonsterNum)
If mStatus(nIdx) <> MON_LOCATED Then Exit Function
If nStartIdx = 0 Or nRooms < 1 Then Exit Function

If nLevel < 1 Then nLevel = 1
i = 0
Do While i + 1 < nBP
    If bpLevel(i + 1) > nLevel Then Exit Do
    i = i + 1
Loop
If mHit(nIdx, i) = 1 Then
    ReachMonsterInfo = reachYes
    Exit Function
End If

'first higher band that reaches it (bpLevel is sorted ascending)
For i = i + 1 To nBP - 1
    If mHit(nIdx, i) = 1 Then nReachableAtLvl = bpLevel(i): Exit For
Next i

Call VisitAtLevel(nLevel, visited)

'none of the monster's rooms is reachable, or it would have been reachYes
For Each vKey In mRooms(nIdx).Keys
    r = CLng(vKey)
    If r >= 1 And r <= nRooms Then
        If visited(r) = 1 Then   'shouldn't happen (mHit says unreachable); trust the BFS
            ReachMonsterInfo = reachYes
            nReachableAtLvl = 0
            Exit Function
        End If
    End If
Next vKey

nGate = FindGateEdge(nIdx, visited)
If nGate = 0 Then
    ReachMonsterInfo = reachNoGate
    Exit Function
End If

nGateMap = rMap(eFrom(nGate)): nGateRoom = rRoom(eFrom(nGate))
nDestMap = rMap(eTo(nGate)): nDestRoom = rRoom(eTo(nGate))
nLo = eMin(nGate): nHi = eMax(nGate)
g = eGroup(nGate)
If g > 0 Then
    bRandomGroup = Not grpOK(g)
    'a usable group is level-gated as a whole: it needs a level every one of its edges allows
    For e = 1 To nEdges
        If eGroup(e) = g Then
            If eMin(e) > nLo Then nLo = eMin(e)
            If eMax(e) < nHi Then nHi = eMax(e)
        End If
    Next e
End If
nMinLvl = nLo
If nHi >= NO_MAX Or nHi >= LEVEL_CAP Then nMaxLvl = 0 Else nMaxLvl = nHi
ReachMonsterInfo = reachNo

Exit Function
error:
Call HandleError("ReachMonsterInfo")
ReachMonsterInfo = reachUnknown
End Function

'Route from the nearest usable source to the monster's best spawn room (bToGate=False) or to the
'reachable room holding the gate that blocks it (bToGate=True). Step i (1..nSteps) = room reached after
'action i; index 0 = the source room. sStepLabel(i) = player-facing action leading INTO step i.
'Returns False when not computed / unknown monster / no route. Never triggers the computation.
'Sources are Bank of Godfrey plus the other points of interest (ResolvePOIs) that are reachable from it
'at this level. Two RouteSearches are seeded with all of them (Godfrey first, so it wins ties): one finds the
'paths with the fewest coin-costing steps (toll / price), then the fewest steps; the other just the fewest
'steps. Destination: among the monster's reachable rooms, the lowest category (RCAT_*: Room, Group(lair),
'resolved summon, Group), then the lowest search cost, then the lowest room index (see RouteDest).
'The coin-avoiding route is used unless the shortest route is shorter with more coin-costing steps and the
'coin-avoiding one is more than ROUTE_DETOUR_MAX times its length. When the two differ like that, the one not
'used is the alternate route: nAltSteps/nAltCoinSteps/sAltSourceName describe it (nAltSteps = 0 = none).
'bAltRoute=True returns the alternate instead (the alt outputs then describe the default route), or False when
'there is none. nCoinSteps = coin-costing steps on the returned route.
'bIgnoreLevel: the searches ignore level gates (the sources are still chosen at the character's level, see
'below); bToGate is then ignored as well (the route goes to the monster's rooms). Labels still show each
'step's level range, so a gated step is visible in the route.
'bToGate=True also returns False when the monster is reachable at this level (nothing blocks it); the
'gate is the same one ReachMonsterInfo reports. nSteps = 0 with True = the destination is the source.
'Output arrays are always ReDim'd (0 To nSteps); on False they are (0 To 0) and nSteps = 0.
Public Function ReachRoute(ByVal nMonsterNum As Long, ByVal nLevel As Long, ByVal bToGate As Boolean, _
    ByVal bIgnoreLevel As Boolean, ByVal bAltRoute As Boolean, ByRef nSourceMap As Long, ByRef nSourceRoom As Long, _
    ByRef sSourceName As String, ByRef nStepMap() As Long, ByRef nStepRoom() As Long, ByRef nStepFlags() As Long, _
    ByRef sStepLabel() As String, ByRef nSteps As Long, ByRef nCoinSteps As Long, ByRef nAltSteps As Long, _
    ByRef nAltCoinSteps As Long, ByRef sAltSourceName As String) As Boolean
Dim nIdx As Long, p As Long, r As Long, e As Long
Dim visG() As Long, visF() As Long, parF() As Long, distF() As Long, costF() As Long
Dim visS() As Long, parS() As Long, distS() As Long, costS() As Long
Dim nSeeds() As Long, nSeedCount As Long, vKey As Variant
Dim nDestF As Long, nDestS As Long, nF As Long, nS As Long, nCoinF As Long, nCoinS As Long
Dim nSrcF As Long, nSrcS As Long, bHasAlt As Boolean, bUseShort As Boolean
On Error GoTo error:

nSourceMap = 0: nSourceRoom = 0: sSourceName = "": nSteps = 0
nCoinSteps = 0: nAltSteps = 0: nAltCoinSteps = 0: sAltSourceName = ""
ReDim nStepMap(0 To 0): ReDim nStepRoom(0 To 0): ReDim nStepFlags(0 To 0): ReDim sStepLabel(0 To 0)

If Not ReachIsComputed() Then Exit Function
If Not dMonIdx.Exists(nMonsterNum) Then Exit Function
nIdx = dMonIdx(nMonsterNum)
If mStatus(nIdx) <> MON_LOCATED Then Exit Function
If nStartIdx = 0 Or nRooms < 1 Then Exit Function
If nLevel < 1 Then nLevel = 1

'sources: Bank of Godfrey first, then every other point of interest reachable from it at this level.
'Even when bIgnoreLevel the sources are chosen at the character's level: the route must start somewhere the
'character can actually get to, so the level gate shows up as a step instead of being skipped by starting
'from a point of interest beyond it.
Call VisitAtLevel(nLevel, visG)
ReDim nSeeds(1 To NUM_POIS)
nSeedCount = 1: nSeeds(1) = nStartIdx
For p = 1 To NUM_POIS - 1
    r = nPOIIdx(p)
    If r >= 1 And r <= nRooms Then
        If visG(r) = 1 Then nSeedCount = nSeedCount + 1: nSeeds(nSeedCount) = r
    End If
Next p
'F = fewest coin-costing steps, then fewest steps; S = fewest steps (both reach the same rooms)
Call RouteSearch(nSeeds, nSeedCount, nLevel, bIgnoreLevel, True, visF, parF, distF, costF)
Call RouteSearch(nSeeds, nSeedCount, nLevel, bIgnoreLevel, False, visS, parS, distS, costS)

If bToGate And Not bIgnoreLevel Then
    For Each vKey In mRooms(nIdx).Keys
        r = CLng(vKey)
        If r >= 1 And r <= nRooms Then
            If visG(r) = 1 Then Exit Function   'reachable at this level: no gate
        End If
    Next vKey
    e = FindGateEdge(nIdx, visG)
    If e = 0 Then Exit Function
    nDestF = eFrom(e): nDestS = nDestF
Else
    nDestF = RouteDest(nIdx, visF, costF)
    nDestS = RouteDest(nIdx, visS, costS)
End If
If nDestF = 0 Or nDestS = 0 Then Exit Function
If visF(nDestF) = 0 Or visS(nDestS) = 0 Then Exit Function

nF = RoutePathInfo(nDestF, parF, nCoinF, nSrcF)
nS = RoutePathInfo(nDestS, parS, nCoinS, nSrcS)
If nF < 0 Or nS < 0 Then Exit Function

'the coin-avoiding route, unless a shorter coin-costing one exists and it's more than ROUTE_DETOUR_MAX times longer
bHasAlt = (nCoinS > nCoinF And nS < nF)
bUseShort = False
If bHasAlt Then bUseShort = (nF > ROUTE_DETOUR_MAX * nS)
If bAltRoute Then
    If Not bHasAlt Then Exit Function
    bUseShort = Not bUseShort
End If

If bUseShort Then
    Call RouteSteps(nDestS, nS, parS, nSourceMap, nSourceRoom, nStepMap, nStepRoom, nStepFlags, sStepLabel)
    sSourceName = POIName(nSrcS): nSteps = nS: nCoinSteps = nCoinS
    If bHasAlt Then nAltSteps = nF: nAltCoinSteps = nCoinF: sAltSourceName = POIName(nSrcF)
Else
    Call RouteSteps(nDestF, nF, parF, nSourceMap, nSourceRoom, nStepMap, nStepRoom, nStepFlags, sStepLabel)
    sSourceName = POIName(nSrcF): nSteps = nF: nCoinSteps = nCoinF
    If bHasAlt Then nAltSteps = nS: nAltCoinSteps = nCoinS: sAltSourceName = POIName(nSrcS)
End If
ReachRoute = True

Exit Function
error:
Call HandleError("ReachRoute")
nSourceMap = 0: nSourceRoom = 0: sSourceName = "": nSteps = 0
nCoinSteps = 0: nAltSteps = 0: nAltCoinSteps = 0: sAltSourceName = ""
ReDim nStepMap(0 To 0): ReDim nStepRoom(0 To 0): ReDim nStepFlags(0 To 0): ReDim sStepLabel(0 To 0)
ReachRoute = False
End Function

Public Sub ReachEnsureComputed()
Dim t As Single, bWasEnabled As Boolean
bWasEnabled = True
On Error GoTo error:

If ReachIsComputed() Then Exit Sub
If bReachComputing Then Exit Sub    're-entered from a DoEvents below (e.g. the filter form's Apply)
If DB Is Nothing Then Exit Sub
If sReachFailedFor <> "" And sReachFailedFor = sCurrentDatabaseFile Then Exit Sub
bReachComputing = True: bBuildError = False
t = Timer

bWasEnabled = frmMain.Enabled
frmMain.Enabled = False
Load frmProgressBar
Call frmProgressBar.SetRange(100)
frmProgressBar.ProgressBar.Value = 1
frmProgressBar.lblCaption.Caption = "Mapping monster reachability..."
Set frmProgressBar.objFormOwner = frmMain
frmProgressBar.Show vbModeless, frmMain
DoEvents

nOut = 0: nPool = 0: nEdges = 0: nGroups = 0: nDroppedRandom = 0
bRevAdjBuilt = False: bGrpAdjBuilt = False: nStartIdx = 0: nVisCacheLevel = 0
ReDim nPOIIdx(0 To NUM_POIS - 1): ReDim sPOIName(0 To NUM_POIS - 1)
Erase radjStart, radjEdge, gadjStart, gadjEdge, visCache
Set dChainCache = New Dictionary

Call LoadTextblocks:  Call SetProgress(10)
Call LoadSpells:      Call SetProgress(15)
Call LoadRooms:       Call ResolvePOIs: Call SetProgress(30)
Call LoadItemRooms
Call LoadMonsters:    Call SetProgress(35)
Call BuildEdges:      Call SetProgress(70)
Call ResolveMonsterRooms: Call SetProgress(80)
Call ComputeReachability: Call SetProgress(100)
'each build step handles its own errors (HandleError) and returns; partial data must not count as computed
If bBuildError Then GoTo failed:

sReachDatabase = sCurrentDatabaseFile
If DEVELOPMENT_MODE_RT Then
    Call DebugLogPrint("Reachability: " & nRooms & " rooms, " & nEdges & " edges, " & nGroups & _
        " random groups, " & nDroppedRandom & " scatter blocks dropped, " & nBP & " level bands, " & _
        Format(Timer - t, "0.0") & "s")
    Call ReachWriteDebugCSV(sGlobalWorkingDirectory & "\_reachability.csv")
End If

'ReachIsMonsterOK only needs dMonIdx, mStatus, mHit, bpLevel; ReachMonsterInfo's gate tracing
'also needs the rooms (rMap/rRoom), edges (e*, nEdges), traversal arrays (adjStart/adjEdge, grpOK,
'grpSrc, nStartIdx) and mRooms; ReachRoute also needs the edge metadata (eKind/eDir/eLabel/eCoin),
'the room categories (mRooms values), the direction names (sDirs) and the route sources
'(nPOIIdx/sPOIName); RouteSearch builds its group adjacency (gadjStart/gadjEdge) lazily. Free the rest
'of the build data (labels hold spell/item numbers, so no spell names are needed).
Set dChainCache = Nothing
Set dTB = Nothing: Set dSpell = Nothing: Set dItemRooms = Nothing: Set dRoomIdx = Nothing
Erase tOut, tPool, rExit, rPlaced, rCMD, rSpell, rNPC
Erase tbAction, tbLink, tbCalled, spNum, spAb, spVal, spReqLevel, spCastedBy
Erase mBase, mBaseRaw, mSummonedBy, mGreet
nOut = 0: nPool = 0
GoTo done:

failed:
sReachDatabase = ""
sReachFailedFor = sCurrentDatabaseFile
GoTo done:
error:
Call HandleError("ReachEnsureComputed")
sReachDatabase = ""
sReachFailedFor = sCurrentDatabaseFile
done:
On Error Resume Next
're-enable frmMain before closing the progress bar: if the active window closes while its owner is still
'disabled, Windows activates another application's window and MME ends up behind it
frmMain.Enabled = bWasEnabled
If FormIsLoaded("frmProgressBar") Then
    If bWasEnabled Then Call EnsureAppForeground(frmMain)
    frmProgressBar.Hide
    DoEvents
    Unload frmProgressBar
End If
bReachComputing = False
End Sub

'Same format as reachability.py --csv: number,name,status,levels (sorted by number).
Public Sub ReachWriteDebugCSV(ByVal sFile As String)
Dim nFile As Integer, i As Long, sStatus As String, sLevels As String
On Error GoTo error:

nFile = FreeFile
Open sFile For Output As #nFile
Print #nFile, "number,name,status,levels"
For i = 1 To nMons
    sLevels = ""
    Select Case mStatus(i)
        Case MON_NOT_IN_GAME: sStatus = "not-in-game"
        Case MON_NO_LOCATION: sStatus = "no-location"
        Case Else
            sLevels = BandsText(i)
            sStatus = IIf(sLevels = "", "unreachable", "reachable")
    End Select
    Print #nFile, mNum(i) & "," & CsvField(mName(i)) & "," & sStatus & "," & CsvField(sLevels)
Next i
Close #nFile
Exit Sub

error:
Call HandleError("ReachWriteDebugCSV")
On Error Resume Next
Close #nFile
End Sub


'=============================================================== loading

Private Sub SetProgress(ByVal nPct As Long)
On Error Resume Next
If nPct >= frmProgressBar.ProgressBar.Max Then nPct = frmProgressBar.ProgressBar.Max - 1
frmProgressBar.ProgressBar.Value = nPct
DoEvents
End Sub

Private Function FieldExists(rs As DAO.Recordset, ByVal sName As String) As Boolean
Dim f As DAO.Field
On Error GoTo nope:
Set f = rs.Fields(sName)
FieldExists = True
Exit Function
nope:
FieldExists = False
End Function

Private Function FldStr(rs As DAO.Recordset, ByVal sName As String) As String
On Error GoTo nope:
If Not IsNull(rs.Fields(sName).Value) Then FldStr = CStr(rs.Fields(sName).Value)
nope:
End Function

Private Function FldLng(rs As DAO.Recordset, ByVal sName As String) As Long
On Error GoTo nope:
If Not IsNull(rs.Fields(sName).Value) Then FldLng = CLng(val(rs.Fields(sName).Value))
nope:
End Function

Private Sub LoadTextblocks()
Dim rs As DAO.Recordset, n As Long, nCount As Long, bLink As Boolean, bCalled As Boolean
On Error GoTo error:

Set dTB = New Dictionary
Set rs = DB.OpenRecordset("SELECT * FROM TBInfo", dbOpenSnapshot, dbForwardOnly)
bLink = FieldExists(rs, "LinkTo")
bCalled = FieldExists(rs, "Called From")
ReDim tbAction(1 To 1000): ReDim tbLink(1 To 1000): ReDim tbCalled(1 To 1000)
Do Until rs.EOF
    nCount = nCount + 1
    If nCount > UBound(tbAction) Then
        ReDim Preserve tbAction(1 To nCount * 2)
        ReDim Preserve tbLink(1 To nCount * 2)
        ReDim Preserve tbCalled(1 To nCount * 2)
    End If
    n = FldLng(rs, "Number")
    tbAction(nCount) = Replace(FldStr(rs, "Action"), vbCr, "")
    If bLink Then tbLink(nCount) = FldLng(rs, "LinkTo")
    If bCalled Then tbCalled(nCount) = FldStr(rs, "Called From")
    If Not dTB.Exists(n) Then dTB.Add n, nCount
    rs.MoveNext
Loop
rs.Close
Exit Sub
error:
Call HandleError("Reach.LoadTextblocks")
bBuildError = True
End Sub

Private Sub LoadSpells()
Dim rs As DAO.Recordset, n As Long, j As Long, bCasted As Boolean, bReq As Boolean
Dim fAb(0 To MAX_ABILS - 1) As Boolean
On Error GoTo error:

Set dSpell = New Dictionary
Set rs = DB.OpenRecordset("SELECT * FROM Spells", dbOpenSnapshot, dbForwardOnly)
bCasted = FieldExists(rs, "Casted By")
bReq = FieldExists(rs, "ReqLevel")
For j = 0 To MAX_ABILS - 1
    fAb(j) = FieldExists(rs, "Abil-" & j)
Next j
nSpells = 0
ReDim spNum(1 To 500): ReDim spAb(1 To 500, 0 To MAX_ABILS - 1): ReDim spVal(1 To 500, 0 To MAX_ABILS - 1)
ReDim spReqLevel(1 To 500): ReDim spCastedBy(1 To 500)
Do Until rs.EOF
    nSpells = nSpells + 1
    If nSpells > UBound(spNum) Then
        ReDim Preserve spNum(1 To nSpells * 2)
        ReDim Preserve spReqLevel(1 To nSpells * 2)
        ReDim Preserve spCastedBy(1 To nSpells * 2)
        Call GrowAbils(nSpells * 2)
    End If
    n = FldLng(rs, "Number")
    spNum(nSpells) = n
    If bReq Then spReqLevel(nSpells) = FldLng(rs, "ReqLevel")
    If bCasted Then spCastedBy(nSpells) = FldStr(rs, "Casted By")
    For j = 0 To MAX_ABILS - 1
        If fAb(j) Then
            spAb(nSpells, j) = FldLng(rs, "Abil-" & j)
            spVal(nSpells, j) = FldLng(rs, "AbilVal-" & j)
        End If
    Next j
    If Not dSpell.Exists(n) Then dSpell.Add n, nSpells
    rs.MoveNext
Loop
rs.Close
Exit Sub
error:
Call HandleError("Reach.LoadSpells")
bBuildError = True
End Sub

'ReDim Preserve can only grow the last dimension, so copy the 2-D ability arrays.
Private Sub GrowAbils(ByVal nNew As Long)
Dim a() As Long, v() As Long, i As Long, j As Long
ReDim a(1 To nNew, 0 To MAX_ABILS - 1): ReDim v(1 To nNew, 0 To MAX_ABILS - 1)
For i = 1 To UBound(spAb, 1)
    For j = 0 To MAX_ABILS - 1
        a(i, j) = spAb(i, j): v(i, j) = spVal(i, j)
    Next j
Next i
spAb = a: spVal = v
End Sub

Private Sub LoadRooms()
Dim rs As DAO.Recordset, j As Long, sKey As String
Dim fMap As DAO.Field, fRoom As DAO.Field, fCMD As DAO.Field, fSpell As DAO.Field, fNPC As DAO.Field
Dim fExit(0 To 9) As DAO.Field, bSpell As Boolean, bNPC As Boolean, bPlaced As Boolean
On Error GoTo error:

sDirs = Array("N", "S", "E", "W", "NE", "NW", "SE", "SW", "U", "D")
Set dRoomIdx = New Dictionary
Set rs = DB.OpenRecordset("SELECT * FROM Rooms", dbOpenSnapshot, dbForwardOnly)
Set fMap = rs.Fields("Map Number"): Set fRoom = rs.Fields("Room Number")
Set fCMD = rs.Fields("CMD")
bSpell = FieldExists(rs, "Spell"): bNPC = FieldExists(rs, "NPC"): bPlaced = FieldExists(rs, "Placed")
If bSpell Then Set fSpell = rs.Fields("Spell")
If bNPC Then Set fNPC = rs.Fields("NPC")
For j = 0 To 9
    Set fExit(j) = rs.Fields(sDirs(j))
Next j

nRooms = 0
ReDim rMap(1 To 5000): ReDim rRoom(1 To 5000): ReDim rCMD(1 To 5000): ReDim rSpell(1 To 5000)
ReDim rNPC(1 To 5000): ReDim rPlaced(1 To 5000): ReDim rExit(0 To 9, 1 To 5000)
Do Until rs.EOF
    nRooms = nRooms + 1
    If nRooms > UBound(rMap) Then
        ReDim Preserve rMap(1 To nRooms * 2): ReDim Preserve rRoom(1 To nRooms * 2)
        ReDim Preserve rCMD(1 To nRooms * 2): ReDim Preserve rSpell(1 To nRooms * 2)
        ReDim Preserve rNPC(1 To nRooms * 2): ReDim Preserve rPlaced(1 To nRooms * 2)
        ReDim Preserve rExit(0 To 9, 1 To nRooms * 2)
    End If
    rMap(nRooms) = Nz0(fMap.Value): rRoom(nRooms) = Nz0(fRoom.Value)
    rCMD(nRooms) = Nz0(fCMD.Value)
    If bSpell Then rSpell(nRooms) = Nz0(fSpell.Value)
    If bNPC Then rNPC(nRooms) = Nz0(fNPC.Value)
    If bPlaced Then rPlaced(nRooms) = FldStr(rs, "Placed")
    For j = 0 To 9
        If Not IsNull(fExit(j).Value) Then rExit(j, nRooms) = CStr(fExit(j).Value) Else rExit(j, nRooms) = ""
    Next j
    sKey = rMap(nRooms) & "/" & rRoom(nRooms)
    If Not dRoomIdx.Exists(sKey) Then dRoomIdx.Add sKey, nRooms
    rs.MoveNext
Loop
rs.Close
Exit Sub
error:
Call HandleError("Reach.LoadRooms")
bBuildError = True
End Sub

'Route sources for ReachRoute, in priority order (ties go to the earlier one). Index 0 must be Bank of
'Godfrey (START_MAP/START_ROOM), the reachability start, so nPOIIdx(0) = nStartIdx. Needs dRoomIdx.
Private Sub ResolvePOIs()
ReDim nPOIIdx(0 To NUM_POIS - 1): ReDim sPOIName(0 To NUM_POIS - 1)
Call SetPOI(0, "Bank of Godfrey", START_MAP, START_ROOM)
Call SetPOI(1, "Bank of Arlysia", 17, 2435)
Call SetPOI(2, "Bank of Khazarad", 6, 1334)
Call SetPOI(3, "The Lost City", 16, 426)
Call SetPOI(4, "Bank of Albion", 14, 1)
Call SetPOI(5, "Port Blackwater", 14, 10)
End Sub

Private Sub SetPOI(ByVal i As Long, ByVal sName As String, ByVal nMap As Long, ByVal nRoom As Long)
sPOIName(i) = sName
nPOIIdx(i) = RoomIndex(nMap, nRoom)
End Sub

Private Function Nz0(ByVal v As Variant) As Long
If IsNull(v) Then Nz0 = 0 Else Nz0 = CLng(val(v))
End Function

Private Function RoomIndex(ByVal nMap As Long, ByVal nRoom As Long) As Long
Dim sKey As String
sKey = nMap & "/" & nRoom
If dRoomIdx.Exists(sKey) Then RoomIndex = dRoomIdx(sKey)
End Function

Private Sub AddItemRoom(ByVal nItem As Long, ByVal nRoomIdx As Long)
Dim c As Collection
If nRoomIdx = 0 Then Exit Sub
If Not dItemRooms.Exists(nItem) Then
    Set c = New Collection
    dItemRooms.Add nItem, c
End If
dItemRooms(nItem).Add nRoomIdx
End Sub

Private Sub LoadItemRooms()
Dim rs As DAO.Recordset, i As Long, sNums As Variant, j As Long, n As Long, sTxt As String
Dim nPos As Long, nM As Long, nR As Long
On Error GoTo error:

Set dItemRooms = New Dictionary
For i = 1 To nRooms
    If rPlaced(i) <> "" Then
        sNums = NumbersIn(rPlaced(i))
        For j = 0 To UBound(sNums)
            Call AddItemRoom(CLng(sNums(j)), i)
        Next j
    End If
Next i

Set rs = DB.OpenRecordset("SELECT * FROM Items", dbOpenSnapshot, dbForwardOnly)
If FieldExists(rs, "Obtained From") Then
    Do Until rs.EOF
        sTxt = FldStr(rs, "Obtained From")
        n = FldLng(rs, "Number")
        nPos = 1
        Do While NextMapRoom(sTxt, "Room ", nPos, nM, nR)
            Call AddItemRoom(n, RoomIndex(nM, nR))
        Loop
        rs.MoveNext
    Loop
End If
rs.Close
Exit Sub
error:
Call HandleError("Reach.LoadItemRooms")
bBuildError = True
End Sub

Private Sub LoadMonsters()
Dim rs As DAO.Recordset, n As Long, bSB As Boolean, bGreet As Boolean, bInGame As Boolean
Dim sSB As String, nPos As Long, nM As Long, nR As Long, nIdx As Long, sPrefix As Variant, k As Long
Dim sCat As Variant, nCat As Byte
On Error GoTo error:

Set dMonIdx = New Dictionary
Set rs = DB.OpenRecordset("SELECT * FROM Monsters ORDER BY Number", dbOpenSnapshot, dbForwardOnly)
bSB = FieldExists(rs, "Summoned By"): bGreet = FieldExists(rs, "GreetTXT"): bInGame = FieldExists(rs, "In Game")
sPrefix = Array("Group: ", "Group(lair): ", "Room ")
sCat = Array(RCAT_GROUP, RCAT_LAIR, RCAT_ROOM)   'a room listed more than once keeps the lowest
nMons = 0
ReDim mNum(1 To 500): ReDim mName(1 To 500): ReDim mSummonedBy(1 To 500): ReDim mGreet(1 To 500)
ReDim mStatus(1 To 500): ReDim mRooms(1 To 500): ReDim mBase(1 To 500): ReDim mBaseRaw(1 To 500)
Do Until rs.EOF
    nMons = nMons + 1
    If nMons > UBound(mNum) Then
        ReDim Preserve mNum(1 To nMons * 2): ReDim Preserve mName(1 To nMons * 2)
        ReDim Preserve mSummonedBy(1 To nMons * 2): ReDim Preserve mGreet(1 To nMons * 2)
        ReDim Preserve mStatus(1 To nMons * 2): ReDim Preserve mRooms(1 To nMons * 2)
        ReDim Preserve mBase(1 To nMons * 2): ReDim Preserve mBaseRaw(1 To nMons * 2)
    End If
    n = FldLng(rs, "Number")
    mNum(nMons) = n
    mName(nMons) = FldStr(rs, "Name")
    If bSB Then mSummonedBy(nMons) = FldStr(rs, "Summoned By")
    If bGreet Then mGreet(nMons) = FldLng(rs, "GreetTXT")
    mStatus(nMons) = MON_NO_LOCATION
    If bInGame Then
        If FldLng(rs, "In Game") = 0 Then mStatus(nMons) = MON_NOT_IN_GAME
    End If
    Set mBase(nMons) = New Dictionary
    For k = 0 To 2
        nPos = 1: nCat = CByte(sCat(k))
        Do While NextMapRoom(mSummonedBy(nMons), CStr(sPrefix(k)), nPos, nM, nR)
            mBaseRaw(nMons) = mBaseRaw(nMons) + 1
            nIdx = RoomIndex(nM, nR)
            If nIdx > 0 Then
                If Not mBase(nMons).Exists(nIdx) Then
                    mBase(nMons).Add nIdx, nCat
                ElseIf mBase(nMons)(nIdx) > nCat Then
                    mBase(nMons)(nIdx) = nCat
                End If
            End If
        Loop
    Next k
    If Not dMonIdx.Exists(n) Then dMonIdx.Add n, nMons
    rs.MoveNext
Loop
rs.Close
Exit Sub
error:
Call HandleError("Reach.LoadMonsters")
bBuildError = True
End Sub


'=============================================================== text helpers

Private Function TrimWS(ByVal s As String) As String
Dim a As Long, b As Long, c As String
a = 1: b = Len(s)
Do While a <= b
    c = mid$(s, a, 1)
    If c <> " " And c <> vbTab And c <> vbCr And c <> vbLf Then Exit Do
    a = a + 1
Loop
Do While b >= a
    c = mid$(s, b, 1)
    If c <> " " And c <> vbTab And c <> vbCr And c <> vbLf Then Exit Do
    b = b - 1
Loop
If b >= a Then TrimWS = mid$(s, a, b - a + 1)
End Function

Private Function IsDigits(ByVal s As String) As Boolean
Dim i As Long
If Len(s) = 0 Or Len(s) > 9 Then Exit Function
For i = 1 To Len(s)
    If mid$(s, i, 1) < "0" Or mid$(s, i, 1) > "9" Then Exit Function
Next i
IsDigits = True
End Function

'Python int(): optional sign, then digits.
Private Function IsInt(ByVal s As String) As Boolean
If Len(s) > 1 And (Left$(s, 1) = "-" Or Left$(s, 1) = "+") Then s = mid$(s, 2)
IsInt = IsDigits(s)
End Function

Private Function SplitWords(ByVal s As String) As Variant
s = TrimWS(Replace(s, vbTab, " "))
Do While InStr(1, s, "  ") > 0
    s = Replace(s, "  ", " ")
Loop
SplitWords = Split(s, " ")
End Function

'All integers in a string ("1419,<nul>" -> [1419]).
Private Function NumbersIn(ByVal s As String) As Variant
Dim i As Long, sCur As String, sAll As String, c As String
For i = 1 To Len(s)
    c = mid$(s, i, 1)
    If c >= "0" And c <= "9" Then
        sCur = sCur & c
    ElseIf sCur <> "" Then
        If Len(sCur) <= 9 Then sAll = sAll & sCur & " "
        sCur = ""
    End If
Next i
If sCur <> "" And Len(sCur) <= 9 Then sAll = sAll & sCur & " "
NumbersIn = Split(TrimWS(sAll), " ")
If TrimWS(sAll) = "" Then NumbersIn = Split("", " ")
End Function

'Finds the next "<prefix><map>/<room>" at or after nPos. Advances nPos. An entry followed by "+" was
'cut off (the field is truncated at 250 chars: "Room 14/62+" is really 14/62xx) and is skipped.
Private Function NextMapRoom(ByVal s As String, ByVal sPrefix As String, ByRef nPos As Long, _
    ByRef nMap As Long, ByRef nRoom As Long) As Boolean
Dim x As Long, i As Long, sM As String, sR As String
Do
    x = InStr(nPos, s, sPrefix)
    If x = 0 Then Exit Function
    nPos = x + Len(sPrefix)
    i = nPos: sM = "": sR = ""
    Do While i <= Len(s)
        If mid$(s, i, 1) < "0" Or mid$(s, i, 1) > "9" Then Exit Do
        sM = sM & mid$(s, i, 1): i = i + 1
    Loop
    If sM <> "" And mid$(s, i, 1) = "/" Then
        i = i + 1
        Do While i <= Len(s)
            If mid$(s, i, 1) < "0" Or mid$(s, i, 1) > "9" Then Exit Do
            sR = sR & mid$(s, i, 1): i = i + 1
        Loop
        If sR <> "" And Len(sM) <= 9 And Len(sR) <= 9 And mid$(s, i, 1) <> "+" Then
            nMap = CLng(sM): nRoom = CLng(sR): nPos = i
            NextMapRoom = True
            Exit Function
        End If
    End If
Loop
End Function

Private Function CsvField(ByVal s As String) As String
If InStr(1, s, ",") > 0 Or InStr(1, s, """") > 0 Then
    CsvField = """" & Replace(s, """", """""") & """"
Else
    CsvField = s
End If
End Function


'=============================================================== textblock ops

'One textblock op -> kind (OP_*), with its numbers in v1/v2 (v2 = 0 when absent).
Private Function OpKind(ByVal sOp As String, ByRef v1 As Long, ByRef v2 As Long) As Long
Dim w As Variant, k As String, n As Long
v1 = 0: v2 = 0
w = SplitWords(sOp)
n = UBound(w)
If n < 0 Then Exit Function
k = LCase$(w(0))
If k = "" Then Exit Function

If n = 0 And IsDigits(k) Then v1 = CLng(k): OpKind = OP_TEXT: Exit Function   '"keyword:N"
If n >= 1 Then
    If IsInt(w(1)) Then
        v1 = CLng(w(1))
        Select Case k
            Case "teleport"
                If n >= 2 Then
                    If IsInt(w(2)) Then v2 = CLng(w(2))
                End If
                OpKind = OP_TELEPORT: Exit Function
            Case "text": OpKind = OP_TEXT: Exit Function
            Case "random": OpKind = OP_RANDOM: Exit Function
            Case "cast": OpKind = OP_CAST: Exit Function
            Case "minlevel": OpKind = OP_MINLEVEL: Exit Function
            Case "maxlevel": OpKind = OP_MAXLEVEL: Exit Function
            Case "roomitem": OpKind = OP_ROOMITEM: Exit Function
            Case "price", "checkitem", "takeitem", "class", "race", "checkability", "testability"
                OpKind = OP_OTHER: Exit Function
        End Select
        v1 = 0
    End If
    If k = "testskill" And n >= 2 Then OpKind = OP_OTHER: Exit Function
End If
Select Case k
    Case "goodaligned", "evilaligned", "nomonsters", "needmonster", "failroomitem", "failability"
        OpKind = OP_OTHER
End Select
End Function

'"cmd:op:op" -> command + ops. In a block reached by a jump/LinkTo (bCalled) the first
'segment is often already an op; then bHasCmd = False and it's the first op.
Private Sub SplitLine(ByVal sLine As String, ByVal bCalled As Boolean, ByRef sCmd As String, _
    ByRef bHasCmd As Boolean, ByRef sOps() As String, ByRef nOps As Long)
Dim parts As Variant, i As Long, s As String, v1 As Long, v2 As Long
parts = Split(TrimWS(sLine), ":")
nOps = 0
ReDim sOps(0 To UBound(parts) + 1)
sCmd = TrimWS(parts(0)): bHasCmd = True
If bCalled And sCmd <> "" Then
    If OpKind(sCmd, v1, v2) <> OP_NONE And Not IsDigits(sCmd) Then
        sOps(0) = sCmd: nOps = 1
        sCmd = "": bHasCmd = False
    End If
End If
For i = 1 To UBound(parts)
    s = TrimWS(parts(i))
    If s <> "" Then sOps(nOps) = s: nOps = nOps + 1
Next i
End Sub

'Top-level command line worth following: has a typed command, not a random roll number.
Private Function IsCommandLine(ByVal sCmd As String) As Boolean
Dim names As Variant, i As Long
If IsDigits(sCmd) Then Exit Function
names = Split(Replace(sCmd, "*", ""), "|")
For i = 0 To UBound(names)
    If TrimWS(names(i)) <> "" Then IsCommandLine = True: Exit Function
Next i
End Function

Private Function TBIndex(ByVal n As Long) As Long
If dTB.Exists(n) Then TBIndex = dTB(n)
End Function

Private Function SpellIndex(ByVal n As Long) As Long
If dSpell.Exists(n) Then SpellIndex = dSpell(n)
End Function

Private Sub AddOut(ByVal nRoom As Long, ByVal nMap As Long, ByVal nMin As Long, ByVal nMax As Long, _
    ByVal sItems As String, ByVal nPrice As Long, ByVal nLine As Long)
If nOut = 0 Then ReDim tOut(0 To 255)
If nOut > UBound(tOut) Then ReDim Preserve tOut(0 To nOut * 2)
With tOut(nOut)
    .Room = nRoom: .Map = nMap: .MinLvl = nMin: .MaxLvl = nMax
    .IsRandom = False: .RoomItems = sItems: .LineNo = nLine: .Cmd = "": .Price = nPrice
End With
nOut = nOut + 1
End Sub

Private Sub FromOps(sOps() As String, ByVal nOps As Long, ByVal nMin As Long, ByVal nMax As Long, _
    ByVal sItems As String, ByVal nPrice As Long, ByVal nDepth As Long, ByVal sSeen As String, _
    ByVal sCmd As String, ByVal bHasCmd As Boolean, ByVal nLine As Long)
Dim i As Long, k As Long, v1 As Long, v2 As Long, nStart As Long, a As Long, b As Long, nDistinct As Long
Dim bDup As Boolean, w As Variant
For i = 0 To nOps - 1
    k = OpKind(sOps(i), v1, v2)
    Select Case k
        Case OP_MINLEVEL: If v1 > nMin Then nMin = v1
        Case OP_MAXLEVEL: If v1 > 0 And v1 < nMax Then nMax = v1   '"maxlevel 0" = no upper limit, as for exits
        Case OP_ROOMITEM: sItems = sItems & v1 & ","
        Case OP_TELEPORT: If v1 <> 0 Then Call AddOut(v1, v2, nMin, nMax, sItems, nPrice, nLine)
        Case OP_TEXT: Call FromTB(v1, nMin, nMax, sItems, nPrice, nDepth + 1, sSeen, sCmd, bHasCmd, nLine)
        Case OP_CAST: Call FromSpell(v1, nMin, nMax, sItems, nPrice, nDepth + 1, sSeen, nLine)
        Case OP_OTHER
            If v1 > nPrice Then       'OpKind puts the amount of 'price N' in v1 (other keywords too)
                w = SplitWords(sOps(i))
                If LCase$(w(0)) = "price" Then nPrice = v1
            End If
        Case OP_RANDOM
            'Only truly random if the rolls lead to different places (the ships roll
            'random but every outcome casts the same voyage).
            nStart = nOut
            Call FromTB(v1, nMin, nMax, sItems, nPrice, nDepth + 1, sSeen, "", False, nLine)
            nDistinct = 0
            For a = nStart To nOut - 1
                bDup = False
                For b = nStart To a - 1
                    If tOut(b).Room = tOut(a).Room And tOut(b).Map = tOut(a).Map Then bDup = True: Exit For
                Next b
                If Not bDup Then nDistinct = nDistinct + 1
            Next a
            If nDistinct > MAX_RANDOM_DESTS Then
                nOut = nStart                  'scatter teleport: never guarantees arrival
                nDroppedRandom = nDroppedRandom + 1
            ElseIf nDistinct > 1 Then
                For a = nStart To nOut - 1
                    tOut(a).IsRandom = True
                Next a
            End If
    End Select
Next i
End Sub

Private Sub FromTB(ByVal n As Long, ByVal nMin As Long, ByVal nMax As Long, ByVal sItems As String, _
    ByVal nPrice As Long, ByVal nDepth As Long, ByVal sSeen As String, ByVal sCmd As String, _
    ByVal bHasCmd As Boolean, ByVal nLine As Long)
Dim nIdx As Long, lines As Variant, i As Long, bFilter As Boolean
Dim sLineCmd As String, bLineHasCmd As Boolean, sOps() As String, nOps As Long

If nDepth > MAX_CHAIN_DEPTH Then Exit Sub
If InStr(1, sSeen, ",t" & n & ",") > 0 Then Exit Sub
sSeen = sSeen & "t" & n & ","
nIdx = TBIndex(n)
If nIdx = 0 Then Exit Sub

lines = Split(tbAction(nIdx), vbLf)
If bHasCmd Then
    'A 'text N' jump runs the called block for the same command when it has one.
    For i = 0 To UBound(lines)
        If TrimWS(lines(i)) <> "" Then
            Call SplitLine(lines(i), True, sLineCmd, bLineHasCmd, sOps, nOps)
            If bLineHasCmd And sLineCmd = sCmd Then bFilter = True: Exit For
        End If
    Next i
End If
For i = 0 To UBound(lines)
    If TrimWS(lines(i)) <> "" Then
        Call SplitLine(lines(i), True, sLineCmd, bLineHasCmd, sOps, nOps)
        If Not bFilter Or (bLineHasCmd And sLineCmd = sCmd) Then
            Call FromOps(sOps, nOps, nMin, nMax, sItems, nPrice, nDepth, sSeen, sCmd, bHasCmd, nLine)
        End If
    End If
Next i
If tbLink(nIdx) <> 0 Then Call FromTB(tbLink(nIdx), nMin, nMax, sItems, nPrice, nDepth + 1, sSeen, sCmd, bHasCmd, nLine)
End Sub

Private Sub FromSpell(ByVal s As Long, ByVal nMin As Long, ByVal nMax As Long, ByVal sItems As String, _
    ByVal nPrice As Long, ByVal nDepth As Long, ByVal sSeen As String, ByVal nLine As Long)
Dim nIdx As Long, j As Long, nRoom As Long, nMap As Long, bRoom As Boolean, bMap As Boolean

If nDepth > MAX_CHAIN_DEPTH Then Exit Sub
If InStr(1, sSeen, ",s" & s & ",") > 0 Then Exit Sub
sSeen = sSeen & "s" & s & ","
nIdx = SpellIndex(s)
If nIdx = 0 Then Exit Sub

For j = 0 To MAX_ABILS - 1
    If spAb(nIdx, j) = ABIL_TELEPORT_ROOM And Not bRoom Then nRoom = spVal(nIdx, j): bRoom = True
    If spAb(nIdx, j) = ABIL_TELEPORT_MAP And Not bMap Then nMap = spVal(nIdx, j): bMap = True
Next j
If nRoom <> 0 Then Call AddOut(nRoom, nMap, nMin, nMax, sItems, nPrice, nLine)
For j = 0 To MAX_ABILS - 1
    If spVal(nIdx, j) <> 0 Then
        If spAb(nIdx, j) = ABIL_ENDCAST Then
            Call FromSpell(spVal(nIdx, j), nMin, nMax, sItems, nPrice, nDepth + 1, sSeen, nLine)
        ElseIf spAb(nIdx, j) = ABIL_TEXTBLOCK Then
            Call FromTB(spVal(nIdx, j), nMin, nMax, sItems, nPrice, nDepth + 1, sSeen, "", False, nLine)
        End If
    End If
Next j
End Sub

'Chain outputs for a root (room spell / room command block / NPC greet block), computed
'once and reused for every room that shares it. Returns start/count into tPool.
Private Sub ChainOutputs(ByVal sKind As String, ByVal n As Long, ByRef nStart As Long, ByRef nCount As Long)
Dim sKey As String, v As Variant, nIdx As Long, lines As Variant, i As Long
Dim sCmd As String, bHasCmd As Boolean, sOps() As String, nOps As Long, nLine As Long
Dim nFirst As Long, k As Long

sKey = sKind & n
If dChainCache.Exists(sKey) Then
    v = Split(dChainCache(sKey), ",")
    nStart = CLng(v(0)): nCount = CLng(v(1))
    Exit Sub
End If

nOut = 0
Select Case sKind
    Case "S"
        Call FromSpell(n, 0, NO_MAX, ",", 0, 0, ",", 0)
    Case "C", "G"   'room command block / NPC greet block: top-level command lines
        nIdx = TBIndex(n)
        If nIdx > 0 Then
            lines = Split(tbAction(nIdx), vbLf)
            For i = 0 To UBound(lines)
                If TrimWS(lines(i)) <> "" Then
                    nLine = nLine + 1
                    Call SplitLine(lines(i), False, sCmd, bHasCmd, sOps, nOps)
                    If IsCommandLine(sCmd) Then
                        nFirst = nOut
                        If sKind = "C" Then
                            Call FromOps(sOps, nOps, 0, NO_MAX, ",", 0, 0, ",", sCmd, True, nLine)
                        Else
                            Call FromOps(sOps, nOps, 0, NO_MAX, ",", 0, 0, ",t" & n & ",", "", False, nLine)
                        End If
                        For k = nFirst To nOut - 1   'for route labels (FromOps' random rollback is already done)
                            tOut(k).Cmd = sCmd
                        Next k
                    End If
                End If
            Next i
        End If
End Select

nStart = nPool: nCount = nOut
If nPool + nOut > 0 Then
    If nPool = 0 Then ReDim tPool(0 To 1023)
    Do While nPool + nOut - 1 > UBound(tPool)
        ReDim Preserve tPool(0 To UBound(tPool) * 2 + 1)
    Loop
    For i = 0 To nOut - 1
        tPool(nPool + i) = tOut(i)
    Next i
    nPool = nPool + nOut
End If
dChainCache.Add sKey, nStart & "," & nCount
End Sub


'=============================================================== edges

Private Sub AddEdge(ByVal nFrom As Long, ByVal nTo As Long, ByVal nMin As Long, ByVal nMax As Long, _
    ByVal nGroup As Long, ByVal nKind As Byte, ByVal nDir As Byte, ByVal sLabel As String, ByVal nCoin As Long)
If nFrom = 0 Or nTo = 0 Then Exit Sub
If nEdges = 0 Then
    ReDim eFrom(1 To 65536): ReDim eTo(1 To 65536): ReDim eMin(1 To 65536)
    ReDim eMax(1 To 65536): ReDim eGroup(1 To 65536)
    ReDim eKind(1 To 65536): ReDim eDir(1 To 65536): ReDim eLabel(1 To 65536)
    ReDim eCoin(1 To 65536)
End If
nEdges = nEdges + 1
If nEdges > UBound(eFrom) Then
    ReDim Preserve eFrom(1 To nEdges * 2): ReDim Preserve eTo(1 To nEdges * 2)
    ReDim Preserve eMin(1 To nEdges * 2): ReDim Preserve eMax(1 To nEdges * 2)
    ReDim Preserve eGroup(1 To nEdges * 2)
    ReDim Preserve eKind(1 To nEdges * 2): ReDim Preserve eDir(1 To nEdges * 2)
    ReDim Preserve eLabel(1 To nEdges * 2): ReDim Preserve eCoin(1 To nEdges * 2)
End If
eFrom(nEdges) = nFrom: eTo(nEdges) = nTo: eMin(nEdges) = nMin: eMax(nEdges) = nMax: eGroup(nEdges) = nGroup
eKind(nEdges) = nKind: eDir(nEdges) = nDir: eCoin(nEdges) = nCoin
If LenB(sLabel) > 0 Then eLabel(nEdges) = sLabel   'plain exits stay unallocated
End Sub

'Edges from one room for chain outputs tPool(nStart..): random outputs get a group id
'per (room, top-level line). Edge labels: room command = the output's command; NPC greet =
'sLabelPrefix (NPC name) & "|" & keyword; room spell = sLabelPrefix (spell number). Coin = the
'output's 'price'.
Private Sub AddChainEdges(ByVal nFromIdx As Long, ByVal nStart As Long, ByVal nCount As Long, _
    ByVal bSkipSelf As Boolean, ByVal nKind As Byte, ByVal sLabelPrefix As String)
Dim i As Long, nTo As Long, nMap As Long, nGroup As Long, nLastLine As Long, sLabel As String
nLastLine = -1
For i = nStart To nStart + nCount - 1
    nMap = tPool(i).Map
    If nMap = 0 Then nMap = rMap(nFromIdx)
    nTo = RoomIndex(nMap, tPool(i).Room)
    If nTo > 0 And Not (bSkipSelf And nTo = nFromIdx) Then
        nGroup = 0
        If tPool(i).IsRandom Then
            If tPool(i).LineNo <> nLastLine Then nGroups = nGroups + 1: nLastLine = tPool(i).LineNo
            nGroup = nGroups
        End If
        Select Case nKind
            Case EK_COMMAND: sLabel = tPool(i).Cmd
            Case EK_NPC: sLabel = sLabelPrefix & "|" & tPool(i).Cmd
            Case Else: sLabel = sLabelPrefix
        End Select
        Call AddEdge(nFromIdx, nTo, tPool(i).MinLvl, tPool(i).MaxLvl, nGroup, nKind, NO_DIR, sLabel, tPool(i).Price)
    End If
Next i
End Sub

'"(Toll: N)" in an exit's restriction text -> N (0 when none; "(Toll: 0)" is free).
Private Function ExitToll(ByVal sRest As String) As Long
Dim x As Long, y As Long, i As Long, sBody As String, sNum As String
If LenB(sRest) = 0 Then Exit Function
x = InStr(1, sRest, "(")
Do While x > 0
    y = InStr(x + 1, sRest, ")")
    If y = 0 Then Exit Do
    sBody = TrimWS(mid$(sRest, x + 1, y - x - 1))
    If Left$(sBody, 6) = "Toll: " Then
        sBody = TrimWS(mid$(sBody, 7)): sNum = "": i = 1
        Do While i <= Len(sBody)
            If mid$(sBody, i, 1) < "0" Or mid$(sBody, i, 1) > "9" Then Exit Do
            sNum = sNum & mid$(sBody, i, 1): i = i + 1
        Loop
        If sNum <> "" And Len(sNum) <= 9 Then
            If CLng(sNum) > ExitToll Then ExitToll = CLng(sNum)
        End If
    End If
    x = InStr(y + 1, sRest, "(")
Loop
End Function

'"(Level: a to b)" in an exit's restriction text; b = 0 (or b < a) means no upper limit.
Private Sub ExitLevel(ByVal sRest As String, ByRef nMin As Long, ByRef nMax As Long)
Dim x As Long, y As Long, sBody As String, w As Variant, lo As Long, hi As Long
nMin = 0: nMax = NO_MAX
x = InStr(1, sRest, "(")
Do While x > 0
    y = InStr(x + 1, sRest, ")")
    If y = 0 Then Exit Do
    sBody = TrimWS(mid$(sRest, x + 1, y - x - 1))
    If Left$(sBody, 7) = "Level: " Then
        w = Split(mid$(sBody, 8), " to ")
        If UBound(w) = 1 Then
            If IsDigits(w(0)) And IsDigits(w(1)) Then
                lo = CLng(w(0)): hi = CLng(w(1))
                If hi = 0 Or hi < lo Then hi = NO_MAX
                If lo > nMin Then nMin = lo
                If hi < nMax Then nMax = hi
            End If
        End If
    End If
    x = InStr(y + 1, sRest, "(")
Loop
End Sub

'Non-level part of an exit's restriction text: its "(...)" groups other than "(Level: ...)", e.g.
'"(Key: 1234)"; "" for a plain or level-only exit. Text with no parentheses is returned as is.
Private Function ExitRestriction(ByVal sRest As String) As String
Dim x As Long, y As Long, sBody As String, s As String
If LenB(sRest) = 0 Then Exit Function
x = InStr(1, sRest, "(")
If x = 0 Then ExitRestriction = sRest: Exit Function
Do While x > 0
    y = InStr(x + 1, sRest, ")")
    If y = 0 Then Exit Do
    sBody = TrimWS(mid$(sRest, x + 1, y - x - 1))
    If Left$(sBody, 7) <> "Level: " Then s = s & IIf(s = "", "", " ") & mid$(sRest, x, y - x + 1)
    x = InStr(y + 1, sRest, "(")
Loop
ExitRestriction = s
End Function

'"<map>/<room> (restriction)" -> True with the parts; False for 0/blank/Action fields.
Private Function ParseExit(ByVal s As String, ByRef nMap As Long, ByRef nRoom As Long, ByRef sRest As String) As Boolean
Dim i As Long, sM As String, sR As String
s = TrimWS(s)
If s = "" Or s = "0" Or Left$(s, 6) = "Action" Then Exit Function
i = 1
Do While i <= Len(s)
    If mid$(s, i, 1) < "0" Or mid$(s, i, 1) > "9" Then Exit Do
    sM = sM & mid$(s, i, 1): i = i + 1
Loop
Do While mid$(s, i, 1) = " ": i = i + 1: Loop
If sM = "" Or mid$(s, i, 1) <> "/" Then Exit Function
i = i + 1
Do While mid$(s, i, 1) = " ": i = i + 1: Loop
Do While i <= Len(s)
    If mid$(s, i, 1) < "0" Or mid$(s, i, 1) > "9" Then Exit Do
    sR = sR & mid$(s, i, 1): i = i + 1
Loop
If sR = "" Or Len(sM) > 9 Or Len(sR) > 9 Then Exit Function
nMap = CLng(sM): nRoom = CLng(sR): sRest = TrimWS(mid$(s, i))
ParseExit = True
End Function

Private Sub BuildEdges()
Dim i As Long, j As Long, nMap As Long, nRoom As Long, sRest As String, nMin As Long, nMax As Long
Dim nStart As Long, nCount As Long, m As Long, vKey As Variant, dWhere As Dictionary
Dim s As Long, nTo As Long, sNeed As Variant, k As Long, dHold As Dictionary, c As Collection, vR As Variant
Dim dNPCRooms As Dictionary, sItemList As String
On Error GoTo error:

'exits, room spells, room commands
For i = 1 To nRooms
    For j = 0 To 9
        If ParseExit(rExit(j, i), nMap, nRoom, sRest) Then
            Call ExitLevel(sRest, nMin, nMax)
            Call AddEdge(i, RoomIndex(nMap, nRoom), nMin, nMax, 0, EK_EXIT, CByte(j), ExitRestriction(sRest), _
                ExitToll(sRest))
        End If
    Next j
    If rSpell(i) <> 0 Then
        Call ChainOutputs("S", rSpell(i), nStart, nCount)
        Call AddChainEdges(i, nStart, nCount, True, EK_ROOMSPELL, CStr(rSpell(i)))
    End If
    If rCMD(i) <> 0 Then
        Call ChainOutputs("C", rCMD(i), nStart, nCount)
        Call AddChainEdges(i, nStart, nCount, False, EK_COMMAND, "")
    End If
    If i Mod 2000 = 0 Then Call SetProgress(35 + (i * 25) \ nRooms)
Next i

'NPC conversations: 'ask <npc> <keyword>' from every room the NPC is in
Set dNPCRooms = New Dictionary
For i = 1 To nRooms
    If rNPC(i) <> 0 Then
        If Not dNPCRooms.Exists(rNPC(i)) Then dNPCRooms.Add rNPC(i), New Collection
        dNPCRooms(rNPC(i)).Add i
    End If
Next i
For m = 1 To nMons
    If mGreet(m) <> 0 Then
        Set dWhere = New Dictionary
        For Each vKey In mBase(m).Keys
            dWhere(vKey) = True
        Next vKey
        If dNPCRooms.Exists(mNum(m)) Then
            For Each vR In dNPCRooms(mNum(m))
                dWhere(vR) = True
            Next vR
        End If
        If dWhere.Count > 0 Then
            Call ChainOutputs("G", mGreet(m), nStart, nCount)
            If nCount > 0 Then
                For Each vKey In dWhere.Keys
                    Call AddChainEdges(CLng(vKey), nStart, nCount, False, EK_NPC, mName(m))
                Next vKey
            End If
        End If
    End If
Next m

'Teleport spells that need a room item ('roomitem N') only work where that item is,
'e.g. potion of levitation at the waterfall (3/1) -> 9/1009.
For s = 1 To nSpells
    Call ChainOutputs("S", spNum(s), nStart, nCount)
    For k = nStart To nStart + nCount - 1
        If tPool(k).Map <> 0 Then
            nTo = RoomIndex(tPool(k).Map, tPool(k).Room)
            If nTo > 0 Then
                If tPool(k).RoomItems <> "," Then
                    nMin = tPool(k).MinLvl: nMax = tPool(k).MaxLvl
                    If spReqLevel(s) > nMin Then nMin = spReqLevel(s)
                    sNeed = NumbersIn(tPool(k).RoomItems)
                    Set dHold = Nothing
                    sItemList = ""
                    For j = 0 To UBound(sNeed)
                        sItemList = sItemList & IIf(j > 0, ",", "") & sNeed(j)
                        Set dWhere = New Dictionary
                        If dItemRooms.Exists(CLng(sNeed(j))) Then
                            Set c = dItemRooms(CLng(sNeed(j)))
                            For Each vR In c
                                If dHold Is Nothing Then
                                    dWhere(vR) = True
                                ElseIf dHold.Exists(vR) Then
                                    dWhere(vR) = True
                                End If
                            Next vR
                        End If
                        Set dHold = dWhere
                    Next j
                    If Not dHold Is Nothing Then
                        For Each vR In dHold.Keys
                            Call AddEdge(CLng(vR), nTo, nMin, nMax, 0, EK_ITEMSPELL, NO_DIR, spNum(s) & "|" & sItemList, tPool(k).Price)
                        Next vR
                    End If
                End If
                Exit For   'first teleport with an explicit map, like build.py
            End If
        End If
    Next k
Next s
Exit Sub
error:
Call HandleError("Reach.BuildEdges")
bBuildError = True
End Sub


'=============================================================== monster locations

'"Monster #14, Room 17/1773, Textblock(rndm) #3354(13%)" -> Collection of "R:idx"/"M:n"/"T:n"/"S:n"/"I:n"
'(a number followed by "+" was cut off by the field's 250-char limit and is skipped)
Private Function ParseRefs(ByVal s As String) As Collection
Dim c As New Collection, nPos As Long, nM As Long, nR As Long, nIdx As Long, k As Long
Dim sPre As Variant, sTag As Variant, x As Long, i As Long, sNum As String
sPre = Array("Monster #", "Textblock #", "Textblock(rndm) #", "Spell #", "Item #")
sTag = Array("M:", "T:", "T:", "S:", "I:")
nPos = 1
Do While NextMapRoom(s, "Room ", nPos, nM, nR)
    nIdx = RoomIndex(nM, nR)
    If nIdx > 0 Then c.Add "R:" & nIdx
Loop
For k = 0 To 4
    x = InStr(1, s, sPre(k))
    Do While x > 0
        i = x + Len(sPre(k)): sNum = ""
        Do While i <= Len(s)
            If mid$(s, i, 1) < "0" Or mid$(s, i, 1) > "9" Then Exit Do
            sNum = sNum & mid$(s, i, 1): i = i + 1
        Loop
        If sNum <> "" And Len(sNum) <= 9 And mid$(s, i, 1) <> "+" Then c.Add sTag(k) & sNum
        x = InStr(i, s, sPre(k))
    Loop
Next k
Set ParseRefs = c
End Function

Private Sub Resolve(ByVal sRef As String, ByVal nDepth As Long, ByVal sSeen As String, dOut As Dictionary)
Dim sKind As String, n As Long, nIdx As Long, vKey As Variant, c As Collection, v As Variant
If nDepth > MAX_RESOLVE_DEPTH Then Exit Sub
If InStr(1, sSeen, "|" & sRef & "|") > 0 Then Exit Sub
sSeen = sSeen & sRef & "|"
sKind = Left$(sRef, 1): n = CLng(mid$(sRef, 3))
Select Case sKind
    Case "R"
        dOut(n) = True
        Exit Sub
    Case "I"
        Exit Sub      'summoned by using an item: anywhere, no room to add
    Case "M"
        If dMonIdx.Exists(n) Then
            nIdx = dMonIdx(n)
            If mBaseRaw(nIdx) > 0 Then   'has its own Room/Group entries (even ones not in Rooms)
                For Each vKey In mBase(nIdx).Keys
                    dOut(vKey) = True
                Next vKey
                Exit Sub
            End If
            Set c = SummonRefs(nIdx)
        Else
            Exit Sub
        End If
    Case "T"
        nIdx = TBIndex(n)
        If nIdx = 0 Then Exit Sub
        Set c = ParseRefs(tbCalled(nIdx))
    Case "S"
        nIdx = SpellIndex(n)
        If nIdx = 0 Then Exit Sub
        Set c = ParseRefs(spCastedBy(nIdx))
    Case Else
        Exit Sub
End Select
For Each v In c
    Call Resolve(CStr(v), nDepth + 1, sSeen, dOut)
Next v
End Sub

'A monster's own summon sources: plain "Textblock #n" and "Spell #n" in [Summoned By]
'(not "Textblock(rndm) #", matching build.py).
Private Function SummonRefs(ByVal nMon As Long) As Collection
Dim c As New Collection, s As String, k As Long, x As Long, i As Long, sNum As String
Dim sPre As Variant, sTag As Variant
sPre = Array("Textblock #", "Spell #")
sTag = Array("T:", "S:")
s = mSummonedBy(nMon)
For k = 0 To 1
    x = InStr(1, s, sPre(k))
    Do While x > 0
        i = x + Len(sPre(k)): sNum = ""
        Do While i <= Len(s)
            If mid$(s, i, 1) < "0" Or mid$(s, i, 1) > "9" Then Exit Do
            sNum = sNum & mid$(s, i, 1): i = i + 1
        Loop
        If sNum <> "" And Len(sNum) <= 9 And mid$(s, i, 1) <> "+" Then c.Add sTag(k) & sNum
        x = InStr(i, s, sPre(k))
    Loop
Next k
Set SummonRefs = c
End Function

'(room, map) if a line of this textblock teleports before 'summon <monster>' (Zanthus).
Private Function SummonAfterTeleport(ByVal sAction As String, ByVal nMon As Long, ByRef nRoom As Long, _
    ByRef nMap As Long) As Boolean
Dim lines As Variant, ops As Variant, i As Long, j As Long, w As Variant, bTP As Boolean
lines = Split(sAction, vbLf)
For i = 0 To UBound(lines)
    bTP = False
    ops = Split(lines(i), ":")
    For j = 0 To UBound(ops)
        w = SplitWords(ops(j))
        If UBound(w) >= 1 Then
            If LCase$(w(0)) = "teleport" And IsDigits(w(1)) Then
                nRoom = CLng(w(1)): nMap = 0: bTP = True
                If UBound(w) >= 2 Then
                    If IsDigits(w(2)) Then nMap = CLng(w(2))
                End If
            ElseIf LCase$(w(0)) = "summon" And w(1) = CStr(nMon) And bTP Then
                SummonAfterTeleport = True
                Exit Function
            End If
        End If
    Next j
Next i
End Function

Private Sub ResolveMonsterRooms()
Dim m As Long, v As Variant, dFound As Dictionary, dR As Dictionary, vKey As Variant
Dim nTBIdx As Long, nRoom As Long, nMap As Long, nTo As Long, nCallerMap As Long
On Error GoTo error:

For m = 1 To nMons
    Set mRooms(m) = New Dictionary
    For Each vKey In mBase(m).Keys
        mRooms(m)(vKey) = mBase(m)(vKey)        'RCAT_ROOM / RCAT_LAIR / RCAT_GROUP
    Next vKey
    Set dFound = New Dictionary
    For Each v In SummonRefs(m)
        Set dR = New Dictionary
        Call Resolve(CStr(v), 0, "|", dR)
        nTBIdx = 0
        If Left$(v, 2) = "T:" Then nTBIdx = TBIndex(CLng(mid$(v, 3)))
        If nTBIdx > 0 Then
            If SummonAfterTeleport(tbAction(nTBIdx), mNum(m), nRoom, nMap) Then
                'Teleported first, then summoned: it appears at the destination.
                If dR.Count = 0 Then
                    If nMap <> 0 Then
                        nTo = RoomIndex(nMap, nRoom)
                        If nTo > 0 Then dFound(nTo) = True
                    End If
                Else
                    For Each vKey In dR.Keys
                        nCallerMap = nMap
                        If nCallerMap = 0 Then nCallerMap = rMap(CLng(vKey))
                        nTo = RoomIndex(nCallerMap, nRoom)
                        If nTo > 0 Then dFound(nTo) = True
                    Next vKey
                End If
                GoTo next_ref:
            End If
        End If
        For Each vKey In dR.Keys
            dFound(vKey) = True
        Next vKey
next_ref:
    Next v
    For Each vKey In dFound.Keys
        If Not mRooms(m).Exists(vKey) Then
            mRooms(m).Add vKey, RCAT_RESOLVED
        ElseIf mRooms(m)(vKey) > RCAT_RESOLVED Then
            mRooms(m)(vKey) = RCAT_RESOLVED     'a resolved summon room beats a plain Group entry
        End If
    Next vKey
    If mStatus(m) <> MON_NOT_IN_GAME Then
        'Room entries that aren't in the Rooms table still count as a location (just an
        'unreachable one), as in build.py.
        If mRooms(m).Count > 0 Or mBaseRaw(m) > 0 Then mStatus(m) = MON_LOCATED Else mStatus(m) = MON_NO_LOCATION
    End If
Next m
Exit Sub
error:
Call HandleError("Reach.ResolveMonsterRooms")
bBuildError = True
End Sub


'=============================================================== reachability

Private Sub ComputeReachability()
Dim i As Long, j As Long, nPos() As Long
Dim comp() As Long, grpComp() As Long
Dim dBP As Dictionary, v As Variant, nTmp As Long, b As Long
Dim visited() As Long, e As Long, g As Long
Dim m As Long, vKey As Variant
On Error GoTo error:

nStartIdx = 0
'adjacency (CSR) over non-random edges
ReDim adjStart(1 To nRooms + 1): ReDim nPos(1 To nRooms + 1)
For e = 1 To nEdges
    If eGroup(e) = 0 Then adjStart(eFrom(e)) = adjStart(eFrom(e)) + 1
Next e
j = 1
For i = 1 To nRooms
    nTmp = adjStart(i): adjStart(i) = j: nPos(i) = j: j = j + nTmp
Next i
adjStart(nRooms + 1) = j
ReDim adjEdge(1 To IIf(j > 1, j - 1, 1))
For e = 1 To nEdges
    If eGroup(e) = 0 Then adjEdge(nPos(eFrom(e))) = e: nPos(eFrom(e)) = nPos(eFrom(e)) + 1
Next e

'random groups: usable only if every destination is in one strongly connected area
comp = SCCIds(adjStart, adjEdge)
ReDim grpOK(0 To nGroups): ReDim grpComp(0 To nGroups): ReDim grpSrc(0 To nGroups)
For g = 1 To nGroups
    grpOK(g) = True: grpComp(g) = -1
Next g
For e = 1 To nEdges
    g = eGroup(e)
    If g > 0 Then
        grpSrc(g) = eFrom(e)
        If grpComp(g) = -1 Then
            grpComp(g) = comp(eTo(e))
        ElseIf grpComp(g) <> comp(eTo(e)) Then
            grpOK(g) = False
        End If
    End If
Next e

'level breakpoints (where the set of usable edges can change)
Set dBP = New Dictionary
dBP(CLng(1)) = True   'keys must all be Long: Dictionary treats Integer 1 and Long 1 as different
'(random-group edges included: VisitFromSeeds/RouteSearch test their level ranges too)
For e = 1 To nEdges
    If eMin(e) >= 1 And eMin(e) < LEVEL_CAP Then dBP(eMin(e)) = True
    If eMax(e) >= 0 And eMax(e) < LEVEL_CAP Then dBP(eMax(e) + 1) = True
Next e
nBP = 0
ReDim bpLevel(0 To dBP.Count - 1)
For Each v In dBP.Keys
    If v >= 1 And v <= LEVEL_CAP Then bpLevel(nBP) = v: nBP = nBP + 1
Next v
For i = 1 To nBP - 1     'insertion sort; a few dozen values
    nTmp = bpLevel(i): j = i - 1
    Do While j >= 0
        If bpLevel(j) <= nTmp Then Exit Do
        bpLevel(j + 1) = bpLevel(j): j = j - 1
    Loop
    bpLevel(j + 1) = nTmp
Next i

ReDim mHit(1 To nMons, 0 To nBP - 1)
nStartIdx = RoomIndex(START_MAP, START_ROOM)
If nStartIdx = 0 Then
    'No Bank of Godfrey in this database: don't hide anything.
    For m = 1 To nMons
        If mStatus(m) = MON_LOCATED Then mStatus(m) = MON_NO_LOCATION
    Next m
    Exit Sub
End If

For b = 0 To nBP - 1
    Call VisitAtLevel(bpLevel(b), visited)

    For m = 1 To nMons
        If mStatus(m) = MON_LOCATED Then
            For Each vKey In mRooms(m).Keys
                If visited(CLng(vKey)) = 1 Then mHit(m, b) = 1: Exit For
            Next vKey
        End If
    Next m
    Call SetProgress(80 + (b * 19) \ nBP)
Next b
Exit Sub
error:
Call HandleError("Reach.ComputeReachability")
bBuildError = True
End Sub

'Rooms reachable from Bank of Godfrey at level L, following only edges whose level range
'includes L plus random groups usable at L: visited(1..nRooms) = 1 if reachable.
'Uses adjStart/adjEdge, grpOK, grpSrc, nStartIdx from ComputeReachability.
Private Sub VisitAtLevel(ByVal L As Long, visited() As Long)
Dim nSeeds() As Long, parentEdge() As Long, dist() As Long, i As Long

If nStartIdx = 0 Then
    ReDim visited(1 To nRooms)   'all 0
    Exit Sub
End If
If nVisCacheLevel > 0 And nVisCacheLevel = L Then
    ReDim visited(1 To nRooms)
    For i = 1 To nRooms
        visited(i) = visCache(i)
    Next i
    Exit Sub
End If
ReDim nSeeds(1 To 1)
nSeeds(1) = nStartIdx
Call VisitFromSeeds(nSeeds, 1, L, visited, parentEdge, dist)
'the detail pane asks for the same level several times per monster (gate, routes); keep the last answer
ReDim visCache(1 To nRooms)
For i = 1 To nRooms
    visCache(i) = visited(i)
Next i
nVisCacheLevel = L
End Sub

'Rooms reachable at level L from any of nSeeds(1..nSeedCount), following only edges whose level range
'includes L plus random groups usable at L: visited(1..nRooms) = 1 if reachable, parentEdge() = the edge
'it was first reached by (0 at a seed), dist() = steps from its seed. Seeds are queued in order, so
'equal-distance ties go to the earlier seed. Uses adjStart/adjEdge, grpOK, grpSrc from ComputeReachability.
Private Sub VisitFromSeeds(nSeeds() As Long, ByVal nSeedCount As Long, ByVal L As Long, visited() As Long, _
    parentEdge() As Long, dist() As Long)
Dim queue() As Long, qHead As Long, qTail As Long, i As Long, j As Long
Dim u As Long, e As Long, bChanged As Boolean, g As Long, bLvlOK() As Boolean

ReDim visited(1 To nRooms): ReDim parentEdge(1 To nRooms): ReDim dist(1 To nRooms)   'all 0
If nSeedCount < 1 Then Exit Sub
ReDim queue(1 To nRooms)
ReDim bLvlOK(0 To nGroups)
For g = 1 To nGroups
    bLvlOK(g) = grpOK(g)
Next g
For e = 1 To nEdges
    g = eGroup(e)
    If g > 0 Then
        If L < eMin(e) Or L > eMax(e) Then bLvlOK(g) = False
    End If
Next e

qHead = 1: qTail = 0
For i = 1 To nSeedCount
    u = nSeeds(i)
    If u >= 1 And u <= nRooms Then
        If visited(u) = 0 Then visited(u) = 1: qTail = qTail + 1: queue(qTail) = u
    End If
Next i
Do
    Do While qHead <= qTail
        u = queue(qHead): qHead = qHead + 1
        For j = adjStart(u) To adjStart(u + 1) - 1
            e = adjEdge(j)
            If visited(eTo(e)) = 0 Then
                If L >= eMin(e) And L <= eMax(e) Then
                    visited(eTo(e)) = 1: parentEdge(eTo(e)) = e: dist(eTo(e)) = dist(u) + 1
                    qTail = qTail + 1: queue(qTail) = eTo(e)
                End If
            End If
        Next j
    Loop
    bChanged = False
    For e = 1 To nEdges
        g = eGroup(e)
        If g > 0 Then
            If bLvlOK(g) Then
                If visited(grpSrc(g)) = 1 And visited(eTo(e)) = 0 Then
                    visited(eTo(e)) = 1: parentEdge(eTo(e)) = e: dist(eTo(e)) = dist(grpSrc(g)) + 1
                    qTail = qTail + 1: queue(qTail) = eTo(e)
                    bChanged = True
                End If
            End If
        End If
    Next e
Loop While bChanged
End Sub

'Cheapest paths from nSeeds(1..nSeedCount) for ReachRoute: Dijkstra over rooms keyed on steps, plus
'ROUTE_COIN_COST per coin-costing step when bAvoidCoin (fewest toll/price steps first, then fewest steps).
'Edges are usable when bIgnoreLevel or their level range includes L; random groups when grpOK and (bIgnoreLevel
'or every edge of the group allows L). visited(room) = 1 when reached, parentEdge() = edge of the chosen path
'(0 at a seed), dist() = its step count, cost() = its key (ROUTE_INF when not reached). Seeds start at 0 in
'order and equal keys go to the earlier-pushed entry, so the earlier seed (Godfrey) wins ties. Uses
'adjStart/adjEdge, grpOK from ComputeReachability.
Private Sub RouteSearch(nSeeds() As Long, ByVal nSeedCount As Long, ByVal L As Long, ByVal bIgnoreLevel As Boolean, _
    ByVal bAvoidCoin As Boolean, visited() As Long, parentEdge() As Long, dist() As Long, cost() As Long)
Dim bLvlOK() As Boolean, g As Long, e As Long, i As Long, j As Long, nPass As Long
Dim hKey() As Long, hSeq() As Long, hNode() As Long, nHeap As Long, nSeq As Long
Dim u As Long, v As Long, nKey As Long, nNew As Long

ReDim visited(1 To nRooms): ReDim parentEdge(1 To nRooms): ReDim dist(1 To nRooms)   'all 0
ReDim cost(1 To nRooms)
For i = 1 To nRooms
    cost(i) = ROUTE_INF
Next i
If nSeedCount < 1 Then Exit Sub
If Not bGrpAdjBuilt Then Call BuildGroupAdjacency
ReDim bLvlOK(0 To nGroups)
For g = 1 To nGroups
    bLvlOK(g) = grpOK(g)
Next g
If Not bIgnoreLevel Then
    For e = 1 To nEdges
        g = eGroup(e)
        If g > 0 Then
            If L < eMin(e) Or L > eMax(e) Then bLvlOK(g) = False
        End If
    Next e
End If

ReDim hKey(1 To 1024): ReDim hSeq(1 To 1024): ReDim hNode(1 To 1024)
For i = 1 To nSeedCount
    u = nSeeds(i)
    If u >= 1 And u <= nRooms Then
        If cost(u) > 0 Then
            cost(u) = 0
            Call HeapPush(hKey, hSeq, hNode, nHeap, nSeq, 0, u)
        End If
    End If
Next i
Do While nHeap > 0
    Call HeapPop(hKey, hSeq, hNode, nHeap, nKey, u)
    If visited(u) = 0 And nKey = cost(u) Then
        visited(u) = 1
        'pass 1: plain edges (CSR on eFrom); pass 2: random-group edges leaving u
        For nPass = 1 To 2
            If nPass = 1 Then
                i = adjStart(u): j = adjStart(u + 1) - 1
            Else
                i = gadjStart(u): j = gadjStart(u + 1) - 1
            End If
            Do While i <= j
                If nPass = 1 Then e = adjEdge(i) Else e = gadjEdge(i)
                i = i + 1
                v = eTo(e)
                If visited(v) = 0 Then
                    If nPass = 1 Then
                        If Not bIgnoreLevel Then
                            If L < eMin(e) Or L > eMax(e) Then GoTo next_edge:
                        End If
                    ElseIf Not bLvlOK(eGroup(e)) Then
                        GoTo next_edge:
                    End If
                    nNew = nKey + 1
                    If bAvoidCoin And eCoin(e) > 0 And nNew < ROUTE_INF - ROUTE_COIN_COST Then nNew = nNew + ROUTE_COIN_COST
                    If nNew < cost(v) Then
                        cost(v) = nNew: parentEdge(v) = e: dist(v) = dist(u) + 1
                        Call HeapPush(hKey, hSeq, hNode, nHeap, nSeq, nNew, v)
                    End If
                End If
next_edge:
            Loop
        Next nPass
    End If
Loop
End Sub

'ReachRoute destination among monster nIdx's rooms reached by a RouteSearch: the lowest category (RCAT_*),
'then the lowest cost (so a coin-free room beats a closer coin-costing one when the search avoided coin),
'then the lowest room index. 0 = none of its rooms reached.
Private Function RouteDest(ByVal nIdx As Long, visited() As Long, cost() As Long) As Long
Dim r As Long, nCat As Long, nDest As Long, nDestCat As Long, vKey As Variant

For Each vKey In mRooms(nIdx).Keys
    r = CLng(vKey)
    If r >= 1 And r <= nRooms Then
        If visited(r) = 1 Then
            nCat = CLng(mRooms(nIdx)(vKey))
            If nDest = 0 Then
                nDest = r: nDestCat = nCat
            ElseIf nCat < nDestCat Then
                nDest = r: nDestCat = nCat
            ElseIf nCat = nDestCat And (cost(r) < cost(nDest) Or (cost(r) = cost(nDest) And r < nDest)) Then
                nDest = r: nDestCat = nCat
            End If
        End If
    End If
Next vKey
RouteDest = nDest
End Function

'Steps on the RouteSearch path to nDest (parentEdge = 0 at a seed), or -1 if the links don't lead back to a
'seed. nCoinSteps = coin-costing steps on it; nSrc = the seed room it starts from.
Private Function RoutePathInfo(ByVal nDest As Long, parentEdge() As Long, ByRef nCoinSteps As Long, _
    ByRef nSrc As Long) As Long
Dim n As Long, r As Long, e As Long

nCoinSteps = 0: nSrc = 0
r = nDest
Do While parentEdge(r) <> 0
    n = n + 1
    If n > nRooms Then RoutePathInfo = -1: Exit Function   'safety; the parent links form a tree
    e = parentEdge(r)
    If eCoin(e) > 0 Then nCoinSteps = nCoinSteps + 1
    r = eFrom(e)
Loop
nSrc = r
RoutePathInfo = n
End Function

'Fills ReachRoute's source and step outputs for the n-step path (from RoutePathInfo) to nDest along parentEdge.
Private Sub RouteSteps(ByVal nDest As Long, ByVal n As Long, parentEdge() As Long, ByRef nSourceMap As Long, _
    ByRef nSourceRoom As Long, nStepMap() As Long, nStepRoom() As Long, nStepFlags() As Long, sStepLabel() As String)
Dim i As Long, r As Long, e As Long

ReDim nStepMap(0 To n): ReDim nStepRoom(0 To n): ReDim nStepFlags(0 To n): ReDim sStepLabel(0 To n)
'walk back from the destination, filling the steps from the last one
r = nDest
For i = n To 1 Step -1
    e = parentEdge(r)
    nStepMap(i) = rMap(eTo(e)): nStepRoom(i) = rRoom(eTo(e))
    nStepFlags(i) = StepFlags(e)
    sStepLabel(i) = StepLabel(e)
    r = eFrom(e)
Next i
nStepMap(0) = rMap(r): nStepRoom(0) = rRoom(r)
nSourceMap = rMap(r): nSourceRoom = rRoom(r)
End Sub

'Name of the route source (point of interest) at room index r; "" if it isn't one.
Private Function POIName(ByVal r As Long) As String
Dim p As Long
For p = 0 To NUM_POIS - 1
    If nPOIIdx(p) = r Then POIName = sPOIName(p): Exit Function
Next p
End Function

'Binary min-heap on (key, seq): equal keys pop in push order.
Private Sub HeapPush(hKey() As Long, hSeq() As Long, hNode() As Long, ByRef nHeap As Long, ByRef nSeq As Long, _
    ByVal nKey As Long, ByVal nNode As Long)
Dim i As Long, p As Long
nHeap = nHeap + 1: nSeq = nSeq + 1
If nHeap > UBound(hKey) Then
    ReDim Preserve hKey(1 To nHeap * 2): ReDim Preserve hSeq(1 To nHeap * 2): ReDim Preserve hNode(1 To nHeap * 2)
End If
i = nHeap
Do While i > 1
    p = i \ 2
    If hKey(p) < nKey Or (hKey(p) = nKey And hSeq(p) < nSeq) Then Exit Do
    hKey(i) = hKey(p): hSeq(i) = hSeq(p): hNode(i) = hNode(p)
    i = p
Loop
hKey(i) = nKey: hSeq(i) = nSeq: hNode(i) = nNode
End Sub

Private Sub HeapPop(hKey() As Long, hSeq() As Long, hNode() As Long, ByRef nHeap As Long, ByRef nKey As Long, _
    ByRef nNode As Long)
Dim i As Long, c As Long, k As Long, s As Long, n As Long
nKey = hKey(1): nNode = hNode(1)
k = hKey(nHeap): s = hSeq(nHeap): n = hNode(nHeap)
nHeap = nHeap - 1
If nHeap = 0 Then Exit Sub
i = 1
Do
    c = i * 2
    If c > nHeap Then Exit Do
    If c < nHeap Then
        If hKey(c + 1) < hKey(c) Or (hKey(c + 1) = hKey(c) And hSeq(c + 1) < hSeq(c)) Then c = c + 1
    End If
    If k < hKey(c) Or (k = hKey(c) And s < hSeq(c)) Then Exit Do
    hKey(i) = hKey(c): hSeq(i) = hSeq(c): hNode(i) = hNode(c)
    i = c
Loop
hKey(i) = k: hSeq(i) = s: hNode(i) = n
End Sub

'Adjacency (CSR keyed on eFrom = the group's source room) over random-group edges, for RouteSearch.
Private Sub BuildGroupAdjacency()
Dim i As Long, j As Long, e As Long, nTmp As Long, nPos() As Long

ReDim gadjStart(1 To nRooms + 1): ReDim nPos(1 To nRooms + 1)
For e = 1 To nEdges
    If eGroup(e) > 0 Then gadjStart(eFrom(e)) = gadjStart(eFrom(e)) + 1
Next e
j = 1
For i = 1 To nRooms
    nTmp = gadjStart(i): gadjStart(i) = j: nPos(i) = j: j = j + nTmp
Next i
gadjStart(nRooms + 1) = j
ReDim gadjEdge(1 To IIf(j > 1, j - 1, 1))
For e = 1 To nEdges
    If eGroup(e) > 0 Then gadjEdge(nPos(eFrom(e))) = e: nPos(eFrom(e)) = nPos(eFrom(e)) + 1
Next e
bGrpAdjBuilt = True
End Sub

'The level gate blocking monster nIdx, given the rooms reachable at some level (visited): a reverse BFS
'from the monster's rooms over all edges (levels ignored), through rooms not reachable at this level,
'stopping at the first edge whose source room IS reachable (fewest hops). 0 = none (no known path).
'Callers make sure none of the monster's rooms is reachable first.
Private Function FindGateEdge(ByVal nIdx As Long, visited() As Long) As Long
Dim seen() As Long, queue() As Long, qHead As Long, qTail As Long
Dim u As Long, j As Long, e As Long, r As Long, vKey As Variant

If Not bRevAdjBuilt Then Call BuildReverseAdjacency
ReDim seen(1 To nRooms): ReDim queue(1 To nRooms)
qHead = 1: qTail = 0
For Each vKey In mRooms(nIdx).Keys
    r = CLng(vKey)
    If r >= 1 And r <= nRooms Then
        If seen(r) = 0 Then seen(r) = 1: qTail = qTail + 1: queue(qTail) = r
    End If
Next vKey

Do While qHead <= qTail
    u = queue(qHead): qHead = qHead + 1
    For j = radjStart(u) To radjStart(u + 1) - 1
        e = radjEdge(j)
        If visited(eFrom(e)) = 1 Then
            FindGateEdge = e
            Exit Function
        ElseIf seen(eFrom(e)) = 0 Then
            seen(eFrom(e)) = 1: qTail = qTail + 1: queue(qTail) = eFrom(e)
        End If
    Next j
Loop
End Function

'Edge has a level range (a max of LEVEL_CAP or more means no upper limit).
Private Function EdgeHasLevel(ByVal e As Long) As Boolean
EdgeHasLevel = (eMin(e) > 1 Or eMax(e) < LEVEL_CAP)
End Function

'REACH_STEP_* flags for a route step taken over edge e.
Private Function StepFlags(ByVal e As Long) As Long
Dim f As Long
If eKind(e) <> EK_EXIT Then f = f Or REACH_STEP_ACTION
If EdgeHasLevel(e) Then f = f Or REACH_STEP_LEVEL
If eKind(e) = EK_EXIT And LenB(eLabel(e)) > 0 Then f = f Or REACH_STEP_RESTRICT
If rMap(eFrom(e)) <> rMap(eTo(e)) Then f = f Or REACH_STEP_MAPCHANGE
If eGroup(e) > 0 Then f = f Or REACH_STEP_RANDOM
If eCoin(e) > 0 Then f = f Or REACH_STEP_COIN
StepFlags = f
End Function

'First typed form of a textblock command ("*" wildcards dropped, "a|b" alternatives -> "a").
Private Function CmdText(ByVal sCmd As String) As String
Dim names As Variant, i As Long
names = Split(Replace(sCmd, "*", ""), "|")
For i = 0 To UBound(names)
    If TrimWS(names(i)) <> "" Then CmdText = TrimWS(names(i)): Exit Function
Next i
CmdText = TrimWS(sCmd)
End Function

'Player-facing action for a route step taken over edge e. Spells and items appear as "spell #N" /
'"item #N" so the caller can put names in.
Private Function StepLabel(ByVal e As Long) As String
Dim s As String, x As Long, nLo As Long, v As Variant, i As Long

Select Case eKind(e)
    Case EK_EXIT
        If eDir(e) <= 9 And IsArray(sDirs) Then s = "go " & sDirs(eDir(e)) Else s = "go ?"
        If LenB(eLabel(e)) > 0 Then s = s & " " & eLabel(e)
    Case EK_COMMAND
        s = "type """ & CmdText(eLabel(e)) & """"
    Case EK_ROOMSPELL
        s = "room spell #" & eLabel(e)
    Case EK_NPC
        x = InStr(1, eLabel(e), "|")
        If x > 0 Then
            s = "ask " & Left$(eLabel(e), x - 1) & " " & CmdText(mid$(eLabel(e), x + 1))
        Else
            s = "ask " & eLabel(e)
        End If
    Case EK_ITEMSPELL
        x = InStr(1, eLabel(e), "|")
        If x > 0 Then
            v = Split(mid$(eLabel(e), x + 1), ",")
            s = "cast spell #" & Left$(eLabel(e), x - 1) & " here (needs "
            For i = 0 To UBound(v)
                s = s & IIf(i > 0, ", ", "") & "item #" & v(i)
            Next i
            s = s & ")"
        Else
            s = "cast spell #" & eLabel(e) & " here"
        End If
End Select
'exits already show "(Toll: N)" in their restriction text
If eKind(e) <> EK_EXIT And eCoin(e) > 0 Then s = s & " (price " & Format$(eCoin(e), "#,##0") & ")"

If EdgeHasLevel(e) Then
    nLo = eMin(e)
    If nLo < 1 Then nLo = 1
    If eMax(e) >= LEVEL_CAP Then
        s = s & " (Level " & nLo & "+)"
    Else
        s = s & " (Level " & nLo & " to " & eMax(e) & ")"
    End If
End If
If eGroup(e) > 0 Then s = s & " (random destination)"
StepLabel = s
End Function

'Reverse adjacency (CSR keyed on eTo) over ALL edges, random groups included, for gate tracing.
Private Sub BuildReverseAdjacency()
Dim i As Long, j As Long, e As Long, nTmp As Long, nPos() As Long

ReDim radjStart(1 To nRooms + 1): ReDim nPos(1 To nRooms + 1)
For e = 1 To nEdges
    radjStart(eTo(e)) = radjStart(eTo(e)) + 1
Next e
j = 1
For i = 1 To nRooms
    nTmp = radjStart(i): radjStart(i) = j: nPos(i) = j: j = j + nTmp
Next i
radjStart(nRooms + 1) = j
ReDim radjEdge(1 To IIf(j > 1, j - 1, 1))
For e = 1 To nEdges
    radjEdge(nPos(eTo(e))) = e: nPos(eTo(e)) = nPos(eTo(e)) + 1
Next e
bRevAdjBuilt = True
End Sub

'Strongly connected component id per room (iterative Tarjan), ignoring level gates.
Private Function SCCIds(adjStart() As Long, adjEdge() As Long) As Long()
Dim idx() As Long, low() As Long, onStk() As Boolean, comp() As Long
Dim stk() As Long, nStk As Long, wNode() As Long, wPos() As Long, nW As Long
Dim nCounter As Long, nComp As Long, root As Long, u As Long, p As Long, v As Long, w As Long
Dim bAdvanced As Boolean

ReDim idx(1 To nRooms): ReDim low(1 To nRooms): ReDim onStk(1 To nRooms): ReDim comp(1 To nRooms)
ReDim stk(1 To nRooms): ReDim wNode(1 To nRooms): ReDim wPos(1 To nRooms)
nCounter = 1
For root = 1 To nRooms
    If idx(root) = 0 Then
        idx(root) = nCounter: low(root) = nCounter: nCounter = nCounter + 1
        nStk = nStk + 1: stk(nStk) = root: onStk(root) = True
        nW = 1: wNode(1) = root: wPos(1) = adjStart(root)
        Do While nW > 0
            u = wNode(nW): p = wPos(nW): bAdvanced = False
            Do While p < adjStart(u + 1)
                v = eTo(adjEdge(p)): p = p + 1
                If idx(v) = 0 Then
                    wPos(nW) = p
                    idx(v) = nCounter: low(v) = nCounter: nCounter = nCounter + 1
                    nStk = nStk + 1: stk(nStk) = v: onStk(v) = True
                    nW = nW + 1: wNode(nW) = v: wPos(nW) = adjStart(v)
                    bAdvanced = True
                    Exit Do
                ElseIf onStk(v) Then
                    If idx(v) < low(u) Then low(u) = idx(v)
                End If
            Loop
            If Not bAdvanced Then
                nW = nW - 1
                If nW > 0 Then
                    If low(u) < low(wNode(nW)) Then low(wNode(nW)) = low(u)
                End If
                If low(u) = idx(u) Then
                    Do
                        w = stk(nStk): nStk = nStk - 1
                        onStk(w) = False: comp(w) = nComp
                    Loop While w <> u
                    nComp = nComp + 1
                End If
            End If
        Loop
    End If
Next root
SCCIds = comp
End Function

'"1-5,50+" style bands for one monster, from its per-breakpoint hits.
Private Function BandsText(ByVal m As Long) As String
Dim b As Long, s As String, nLo As Long
For b = 0 To nBP - 1
    If mHit(m, b) = 1 Then
        If b = 0 Then
            nLo = bpLevel(b)
        ElseIf mHit(m, b - 1) = 0 Then
            nLo = bpLevel(b)
        End If
        If b = nBP - 1 Then
            s = s & IIf(s = "", "", ",") & nLo & "+"
        ElseIf mHit(m, b + 1) = 0 Then
            s = s & IIf(s = "", "", ",") & nLo & "-" & (bpLevel(b + 1) - 1)
        End If
    End If
Next b
BandsText = s
End Function
