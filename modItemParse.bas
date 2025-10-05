Attribute VB_Name = "modItemParse"
Option Explicit
Option Base 0

'==============================================================================
' Module: modItemParse.bas
'
' Purpose
' -------
' Parse raw MajorMUD-style game text (inventory, keys, ground “notice” lines)
' into structured arrays, de-duplicate and consolidate counts, and optionally
' populate the Item Manager ListView with enriched, DB-backed rows (including
' Encumbrance, Worn/Type, “Usable”, best Shop and Value). This module is meant
' to be the single source of truth for text?items?ListView flow.
'
' What this module does
' ---------------------
' • Text parsing
'   - Detects and extracts three sections from raw text:
'       • Inventory:    lines after "You are carrying"
'       • Keys:         lines after "You have the following keys" (or ignores "You have no keys")
'       • Ground items: spans beginning "You notice ... here."
'   - Handles wrapped and inline headers (e.g., inventory and keys on the same line).
'   - Ignores bracketed client inserts "[...]" and known coin/currency tokens.
'   - For “Ground”, aggregates by room: MAX per room, then SUM across rooms
'     (avoids double-counting repeated searches of the same room).
'   - De-duplicates “Equipped” entries (case-insensitive).
'   - Consolidates duplicate names and attaches counts: “name (N)”.  Keys are
'     also singularized (e.g., “golden idols (2)”?“golden idol (2)”).
'
' • Public output shape
'   - ItemParseResult with four String() lists: sEquipped, sInventory, sKeys, sGround
'     Values are normalized to “name” or “name (N)”.  “Equipped” retains “Item (Slot)”
'     during initial parse and is deduped; consumers usually strip the slot text.
'
' • DB lookup & ListView population
'   - Matches parsed names to tabItems via exact name (GetItemsByExactNameArr).
'   - Skips items where [Gettable] = 0.
'   - Chooses a representative shop and value using EvaluateBestPriceForHit:
'       • Prefers the cheapest BUY among shops that buy; tie?lowest shop #.
'       • If no BUY, chooses SELL-only; tie?lowest shop #.
'       • Returns friendly “buy/sell” or “(sell) X” text and the chosen shop #.
'   - Adds rows to the Item Manager ListView with AddOneRow, setting columns:
'       [1] Number (text), [2] Name, [3] Flag, [4] QTY, [5] Source,
'       [6] Enc, [7] Type, [8] Worn, [9] Usable, [10] Value, [11] Shop.
'     Also stamps ListSubItems(9).Tag (Value col) with a numeric SELL copper
'     amount (or "0") for stable numeric sorting.
'
' • Convenience helpers
'   - PopulateItemManagerFromParsed: prompts for EQUIPPED/KEYS import and can
'     clear non-flagged existing rows before import.
'   - GetBestShopNumForItem: returns the chosen shop # for a given item record,
'     mirroring the EvaluateBestPriceForHit decision logic.
'   - LV_AddRowByItemNumber: adds one row by item record #, optionally forcing
'     a shop, and setting Flag/QTY as provided.
'
' Key types (public)
' ------------------
' • ItemParseResult: four String() lists (Equipped, Inventory, Keys, Ground)
' • ItemMatch: fields pulled from tabItems for downstream decisions (Number,
'   Name, ItemType, Worn, WeaponType, Encum, ObtainedFrom, Gettable)
' • ShopToken: parsed shop reference from [Obtained From] with “sell-only” flag
'
' External dependencies (project-level)
' -------------------------------------
' • DAO/Access: global Recordset tabItems (indexed by "pkItems" on [Number]).
' • Pricing/shops:
'     GetItemValue(ByVal nItem As Long, ByVal nCharm As Integer, _
'                  ByVal reserved As Long, ByVal shop As Long, ByVal bSellOnly As Boolean) _
'         As tItemValue
'         ' must expose: nCopperBuy, nCopperSell, sFriendlyBuyShort, sFriendlySellShort
'
'     GetShopRoomNames(ByVal shopNum As Long, Optional ByVal unused, _
'                      Optional ByVal bHideRecordNumbers As Boolean) As String
'
' • Item metadata:
'     GetItemType(ByVal nItemType As Long) As String
'     GetWornType(ByVal nWorn As Long) As String
'     GetWeaponType(ByVal nWeaponType As Long) As String
'
' • ListView helpers:
'     LV_AssignRowSeqIfMissing(ByRef lv As ListView, ByRef li As ListItem)
'     ParseActionAndQty(ByVal sFlag As String, ByRef sAction As String, ByRef nQty As Long) ' in modListViewExt
'
' • UI/Forms:
'     frmMain.lvItemManager (ListView target)
'     frmMain.txtCharStats(5).Tag  ' Charm (integer string) used in pricing
'     frmMain.TestGlobalFilter(ByVal nItemNumber As Long) As Boolean  ' “Usable” Yes/No
'
' • Error logging:
'     HandleError(ByVal where As String)  ' expected to exist in the project
'
' Typical usage
' -------------
'   Dim parsed As ItemParseResult
'   parsed = ParseGameTextInventory(rawText)
'   PopulateItemManagerFromParsed parsed, frmMain.lvItemManager
'
' Maintenance tips
' ----------------
' • If you add/rename ListView columns, update AddOneRow’s assignments and any
'   sort-tagging relying on specific subitem indices (Value is SubItem(9)).
' • When changing coin names or adding new cash tokens, update IsCashItem.
' • If you add new movement verbs (e.g., “up”, “down”), extend IsMovementCommand.
'==============================================================================

' ==============================
' Public data shapes
' ==============================
Public Type ItemParseResult
    sEquipped()  As String  ' "Item (Slot)" (deduped later)
    sInventory() As String  ' "name" or "name (N)" consolidated later
    sKeys()      As String  ' keys list items (normalized to singular base name; consolidated)
    sGround()    As String  ' ground items (consolidated across room groups)
End Type

' what we read out of tabItems for a name-exact match
Public Type ItemMatch
    Number       As Long
    name         As String
    ItemType     As Long
    Worn         As Long
    WeaponType   As Long   ' (kept)
    encum        As Long   ' NEW: for Enc column
    ObtainedFrom As String
    Gettable     As Long   ' NEW: 0/1 gate
End Type


Public Type ShopToken
    ShopNumber As Long
    bNObuy     As Boolean  ' True means shop does not sell to player (sell-only)
    bNoSell    As Boolean  ' (unused here, future-safe)
End Type

' ======================================================================
' Room-based ground aggregation (no double-count on repeated searches)
' ======================================================================
Private Type RoomAgg
    Key    As String ' room name || exits (both lower)
    names() As String
    counts() As Long
    used   As Long
End Type

' ==============================
' Constants (headers / markers)
' ==============================
Private Const HDR_INV   As String = "You are carrying"
Private Const HDR_KEYS  As String = "You have the following keys"
Private Const HDR_NOKEY As String = "You have no keys"
Private Const HDR_NOTC  As String = "You notice"

' ==============================
' Public entry
' ==============================
Public Function ParseGameTextInventory(ByVal sInput As String) As ItemParseResult
On Error GoTo fail
    Dim tOut As ItemParseResult
    Dim sNorm As String
    Dim vLines() As String
    Dim i As Long, sLine As String

    sNorm = Replace$(sInput, vbCr, vbNullString)   ' normalize to LF only
    vLines = Split(sNorm, vbLf)
    
    ' init zero-length arrays (one empty cell we’ll reuse for the first add)
    ReDim tOut.sEquipped(0)
    ReDim tOut.sInventory(0)
    ReDim tOut.sKeys(0)
    ReDim tOut.sGround(0)

    ' ---------- Pass 1: scan lines & collect blobs ----------
    i = 0
    Do While i <= UBound(vLines)
        sLine = Trim$(vLines(i))

        If StartsWithCI(sLine, HDR_INV) Then
            Dim blobI As String
            blobI = CollectInventoryBlob(vLines, i)
            ParseInventoryBlob blobI, tOut

        ElseIf StartsWithCI(sLine, HDR_KEYS) Or StartsWithCI(sLine, HDR_NOKEY) Then
            Dim blobK As String
            blobK = CollectKeysBlob(vLines, i)
            ParseKeysBlob blobK, tOut

        ElseIf StartsWithCI(sLine, HDR_NOTC) Then
            Dim blobN As String
            blobN = CollectNoticeBlob(vLines, i)
            ParseNoticeBlob blobN, tOut
        End If

        i = i + 1
    Loop

    ' ---------- Pass 1b: rebuild GROUND by room (dedupe repeated searches) ----------
    ConsolidateGroundByRoom sNorm, tOut

    ' ---------- Pass 2: consolidate lists ----------
    ConsolidateList tOut.sInventory, False
    ConsolidateList tOut.sGround, False
    ConsolidateList tOut.sKeys, True        ' normalize key names ("... keys" -> "... key"; "golden idol (2)")

    ' ---------- Pass 3: dedupe equipped ----------
    DedupeEquipped tOut.sEquipped

    ParseGameTextInventory = tOut
    Exit Function

fail:
    HandleError "ParseGameTextInventory"
    ParseGameTextInventory = tOut
End Function


' ======================================================================
' Inventory blob collection/parsing
' ======================================================================
Private Function CollectInventoryBlob(ByRef vLines() As String, ByRef i As Long) As String
On Error GoTo fail
    Dim s As String, j As Long, sLine As String
    Dim pos As Long, leftPart As String

    s = vbNullString
    j = i
    Do While j <= UBound(vLines)
        sLine = Trim$(vLines(j))

        If ShouldStopSection(sLine, (j <> i)) Then Exit Do

        ' mid-line header split (e.g., "... stormmetal greataxe You have the following keys: ...")
        pos = FindInlineBoundaryPos(sLine)
        If pos > 1 Then
            leftPart = Trim$(Left$(sLine, pos - 1))
            If LenB(leftPart) > 0 Then
                If LenB(s) > 0 Then s = s & " "
                s = s & leftPart
            End If
            vLines(j) = Mid$(sLine, pos)   ' let main loop see the header on this same line next time
            Exit Do
        End If

        If LenB(s) > 0 Then s = s & " "
        s = s & sLine
        j = j + 1
    Loop

    i = j - 1
    CollectInventoryBlob = s
    Exit Function
fail:
    HandleError "CollectInventoryBlob"
    CollectInventoryBlob = s
End Function

Private Sub ParseInventoryBlob(ByVal sBlob As String, ByRef tOut As ItemParseResult)
On Error GoTo fail
    Dim s As String
    Dim toks() As String, k As Long, it As String

    s = sBlob
    If StartsWithCI(s, HDR_INV) Then s = Trim$(Mid$(s, Len(HDR_INV) + 1))

    s = SqueezeSpaces(Replace$(s, vbLf, " "))
    SplitToArray s, toks

    If (Not Not toks) <> 0 Then
        For k = LBound(toks) To UBound(toks)
            it = Trim$(toks(k))
            If LenB(it) = 0 Then GoTo nxt
            If IsBracketLine(it) Then GoTo nxt
            If IsCashItem(it) Then GoTo nxt

            If IsEquippedItem(it) Then
                AddString tOut.sEquipped, it
            Else
                AddString tOut.sInventory, it
            End If
nxt:
        Next k
    End If
    Exit Sub
fail:
    HandleError "ParseInventoryBlob"
End Sub


' ======================================================================
' Keys blob collection/parsing
' ======================================================================
Private Function CollectKeysBlob(ByRef vLines() As String, ByRef i As Long) As String
On Error GoTo fail
    Dim s As String, j As Long, sLine As String
    Dim pos As Long, leftPart As String

    s = vbNullString
    j = i
    Do While j <= UBound(vLines)
        sLine = Trim$(vLines(j))
        
        If j = i And StartsWithCI(sLine, "you have no keys") Then
            ' consume that one line and return it for ParseKeysBlob to ignore
            s = sLine
            j = j + 1
            Exit Do
        End If

        If ShouldStopSection(sLine, (j <> i)) Then Exit Do

        pos = FindInlineBoundaryPos(sLine)
        If pos > 1 Then
            leftPart = Trim$(Left$(sLine, pos - 1))
            If LenB(leftPart) > 0 Then
                If LenB(s) > 0 Then s = s & " "
                s = s & leftPart
            End If
            vLines(j) = Mid$(sLine, pos)
            Exit Do
        End If

        If LenB(s) > 0 Then s = s & " "
        s = s & sLine

        If Right$(sLine, 1) = "." Then
            j = j + 1
            Exit Do
        End If

        j = j + 1
    Loop

    i = j - 1
    CollectKeysBlob = s
    Exit Function
fail:
    HandleError "CollectKeysBlob"
    CollectKeysBlob = s
End Function

Private Sub ParseKeysBlob(ByVal sBlob As String, ByRef tOut As ItemParseResult)
On Error GoTo fail
    Dim s As String
    Dim toks() As String, k As Long, it As String

    s = sBlob

    If StartsWithCI(s, HDR_NOKEY) Then
        Exit Sub
    ElseIf StartsWithCI(s, HDR_KEYS) Then
        s = Trim$(Mid$(s, Len(HDR_KEYS) + 1))
        If Left$(s, 1) = ":" Then s = Trim$(Mid$(s, 2))
    End If

    If InStr(1, LCase$(s), "no keys", vbTextCompare) > 0 Then Exit Sub

    s = RTrim$(s)
    If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)

    s = SqueezeSpaces(Replace$(s, vbLf, " "))

    SplitToArray s, toks
    If (Not Not toks) <> 0 Then
        For k = LBound(toks) To UBound(toks)
            it = Trim$(toks(k))
            If LenB(it) > 0 Then AddString tOut.sKeys, it
        Next k
    End If
    Exit Sub
fail:
    HandleError "ParseKeysBlob"
End Sub


' ======================================================================
' Ground blob collection/parsing (line-wise)
' ======================================================================
Private Function CollectNoticeBlob(ByRef vLines() As String, ByRef i As Long) As String
On Error GoTo fail
    Dim s As String, j As Long, sLine As String
    Dim pos As Long, leftPart As String

    s = vbNullString
    j = i
    Do While j <= UBound(vLines)
        sLine = Trim$(vLines(j))

        If IsBracketLine(sLine) Then j = j + 1: GoTo cont

        If ShouldStopSection(sLine, (j <> i)) Then Exit Do

        pos = FindInlineBoundaryPos(sLine)
        If pos > 1 Then
            leftPart = Trim$(Left$(sLine, pos - 1))
            If LenB(leftPart) > 0 Then
                If LenB(s) > 0 Then s = s & " "
                s = s & leftPart
            End If
            vLines(j) = Mid$(sLine, pos)
            Exit Do
        End If

        If LenB(s) > 0 Then s = s & " "
        s = s & sLine

        If LCase$(Right$(sLine, 5)) = "here." Then
            j = j + 1
            Exit Do
        End If

        j = j + 1
cont:
    Loop

    i = j - 1
    CollectNoticeBlob = s
    Exit Function
fail:
    HandleError "CollectNoticeBlob"
    CollectNoticeBlob = s
End Function

Private Sub ParseNoticeBlob(ByVal sBlob As String, ByRef tOut As ItemParseResult)
On Error GoTo fail
    Dim s As String
    Dim toks() As String, k As Long, it As String

    s = sBlob
    If StartsWithCI(s, HDR_NOTC) Then s = Trim$(Mid$(s, Len(HDR_NOTC) + 1))
    If LCase$(Right$(s, 5)) = "here." Then s = Left$(s, Len(s) - 5)

    ' remove bracket inserts and newlines before splitting
    s = StripBracketedChunks(s)
    s = SqueezeSpaces(Replace$(s, vbLf, " "))

    SplitToArray s, toks
    If (Not Not toks) <> 0 Then
        For k = LBound(toks) To UBound(toks)
            it = Trim$(toks(k))
            If LenB(it) = 0 Then GoTo nxt
            If IsBracketLine(it) Then GoTo nxt
            If IsCashItem(it) Then GoTo nxt
            AddString tOut.sGround, it
nxt:
        Next k
    End If
    Exit Sub
fail:
    HandleError "ParseNoticeBlob"
End Sub


Private Sub ConsolidateGroundByRoom(ByVal sAll As String, ByRef tOut As ItemParseResult)
On Error GoTo fail
    Dim v() As String, i As Long, line As String
    Dim curRoomName As String, curExits As String, curKey As String
    Dim pending As String
    Dim rooms() As RoomAgg, roomUsed As Long

    sAll = Replace$(sAll, vbCr, vbNullString)
    v = Split(sAll, vbLf)

    curRoomName = "": curExits = "": curKey = "": pending = ""

    For i = 0 To UBound(v)
        line = Trim$(v(i))
        If LenB(line) = 0 Then GoTo cont

        ' movement commands => start of a new room likely; flush pending to unknown
        If IsMovementCommand(line) Then
            If LenB(pending) > 0 Then
                Dim mvKey As String
                If LenB(curExits) > 0 Then
                    mvKey = LCase$(Trim$(curRoomName) & "||" & Trim$(curExits))
                Else
                    mvKey = LCase$("__unknown||__unknown")
                End If
                AddNoticeToRoom rooms, roomUsed, mvKey, pending
                pending = ""
            End If
            curRoomName = "": curExits = "": curKey = ""
            GoTo cont
        End If


        ' candidate room name
        If Not StartsWithCI(line, HDR_NOTC) _
        And Not StartsWithCI(line, "also here:") _
        And InStr(1, LCase$(line), "obvious exits:", vbTextCompare) = 0 _
        And Left$(line, 1) <> "[" _
        And Not StartsWithCI(line, "name:") _
        And Not StartsWithCI(line, "race:") _
        And Not StartsWithCI(line, "class:") _
        And Not StartsWithCI(line, HDR_INV) _
        And Not StartsWithCI(line, HDR_KEYS) _
        And Not StartsWithCI(line, HDR_NOKEY) Then
            curRoomName = line
        End If

        If StartsWithCI(line, HDR_NOTC) Then
            ' collect full "You notice ... here." span (may wrap)
            Dim span As String
            span = line
            Do While i < UBound(v)
                If LCase$(Right$(Trim$(v(i)), 5)) = "here." Then Exit Do
                i = i + 1
                span = span & " " & Trim$(v(i))
            Loop
            pending = SqueezeSpaces(StripBracketedChunks(span))
            GoTo cont
        End If

        If InStr(1, LCase$(line), "obvious exits:", vbTextCompare) > 0 Then
            curExits = line
            curKey = LCase$(Trim$(curRoomName) & "||" & Trim$(curExits))
            If LenB(pending) > 0 Then
                AddNoticeToRoom rooms, roomUsed, curKey, pending
                pending = ""
            End If
        End If

cont:
    Next i

    ' final flush: prefer current room key if we have exits, else unknown
    If LenB(pending) > 0 Then
        Dim endKey As String
        If LenB(curExits) > 0 Then
            endKey = LCase$(Trim$(curRoomName) & "||" & Trim$(curExits))
        Else
            endKey = LCase$("__unknown||__unknown")
        End If
        AddNoticeToRoom rooms, roomUsed, endKey, pending
        pending = ""
    End If

    ' merge per-room (MAX within room), then SUM across rooms
    Dim totN() As String, totC() As Long, totU As Long
    Dim r As Long, n As Long

    For r = 0 To roomUsed - 1
        For n = 0 To rooms(r).used - 1
            AccumulateSum totN, totC, totU, rooms(r).names(n), rooms(r).counts(n)
        Next n
    Next r

    ' write back to tOut.sGround as "name (N)"
    If totU = 0 Then
        ReDim tOut.sGround(0)
        tOut.sGround(0) = vbNullString
    Else
        ReDim tOut.sGround(totU - 1)
        For n = 0 To totU - 1
            If totC(n) > 1 Then
                tOut.sGround(n) = totN(n) & " (" & CStr(totC(n)) & ")"
            Else
                tOut.sGround(n) = totN(n)
            End If
        Next n
    End If
    Exit Sub
fail:
    HandleError "ConsolidateGroundByRoom"
End Sub

Private Sub AddNoticeToRoom(ByRef rooms() As RoomAgg, ByRef used As Long, ByVal Key As String, ByVal span As String)
    Dim idx As Long
    idx = GetOrCreateRoomIndex(rooms, used, Key)

    ' strip header/tail & split
    If StartsWithCI(span, HDR_NOTC) Then span = Trim$(Mid$(span, Len(HDR_NOTC) + 1))
    If LCase$(Right$(span, 5)) = "here." Then span = Left$(span, Len(span) - 5)
    span = SqueezeSpaces(Replace$(StripBracketedChunks(span), vbLf, " "))

    Dim toks() As String, i As Long, nm As String, qty As Long, t As String
    SplitToArray span, toks
    If (Not Not toks) = 0 Then Exit Sub

    For i = LBound(toks) To UBound(toks)
        t = Trim$(toks(i))
        If LenB(t) = 0 Then GoTo nxt
        If IsCashItem(t) Then GoTo nxt
        ParseNameAndQty t, nm, qty
        nm = LCase$(nm)
        If qty < 1 Then qty = 1
        AccumulateMax rooms(idx).names, rooms(idx).counts, rooms(idx).used, nm, qty
nxt:
    Next i
End Sub

Private Function GetOrCreateRoomIndex(ByRef rooms() As RoomAgg, ByRef used As Long, ByVal Key As String) As Long
    Dim i As Long
    For i = 0 To used - 1
        If rooms(i).Key = Key Then GetOrCreateRoomIndex = i: Exit Function
    Next
    If used = 0 Then
        ReDim rooms(0)
    Else
        ReDim Preserve rooms(used)
    End If
    rooms(used).Key = Key
    rooms(used).used = 0
    GetOrCreateRoomIndex = used
    used = used + 1
End Function

Private Sub AccumulateMax(ByRef names() As String, ByRef counts() As Long, ByRef used As Long, ByVal nm As String, ByVal qty As Long)
    Dim i As Long
    For i = 0 To used - 1
        If names(i) = nm Then
            If qty > counts(i) Then counts(i) = qty
            Exit Sub
        End If
    Next
    If used = 0 Then
        ReDim names(0): ReDim counts(0)
    Else
        ReDim Preserve names(used): ReDim Preserve counts(used)
    End If
    names(used) = nm
    counts(used) = qty
    used = used + 1
End Sub

Private Sub AccumulateSum(ByRef names() As String, ByRef counts() As Long, ByRef used As Long, ByVal nm As String, ByVal qty As Long)
    Dim i As Long
    For i = 0 To used - 1
        If names(i) = nm Then counts(i) = counts(i) + qty: Exit Sub
    Next
    If used = 0 Then
        ReDim names(0): ReDim counts(0)
    Else
        ReDim Preserve names(used): ReDim Preserve counts(used)
    End If
    names(used) = nm
    counts(used) = qty
    used = used + 1
End Sub

Private Function IsMovementCommand(ByVal line As String) As Boolean
    Dim s As String
    s = LCase$(Trim$(line))
    ' allow prompts like "[HP=...]:n"
    Dim p As Long
    p = InStr(s, "]:")
    If p > 0 Then s = Mid$(s, p + 2)
    s = Trim$(s)

    Select Case s
        Case "n", "s", "e", "w", "nw", "ne", "sw", "se", _
             "north", "south", "east", "west", _
             "northwest", "northeast", "southwest", "southeast"
            IsMovementCommand = True
        Case Else
            IsMovementCommand = False
    End Select
End Function


' ======================================================================
' Section boundary detection
' ======================================================================
Private Function IsSectionBoundary(ByVal sLine As String) As Boolean
    Dim sL As String
    sL = LCase$(Trim$(sLine))

    If LenB(sL) = 0 Then IsSectionBoundary = True: Exit Function

    If StartsWithCI(sL, LCase$(HDR_INV)) Then IsSectionBoundary = True: Exit Function
    If StartsWithCI(sL, LCase$(HDR_KEYS)) Then IsSectionBoundary = True: Exit Function
    If StartsWithCI(sL, LCase$(HDR_NOKEY)) Then IsSectionBoundary = True: Exit Function
    If StartsWithCI(sL, LCase$(HDR_NOTC)) Then IsSectionBoundary = True: Exit Function

    If StartsWithCI(sL, "wealth:") Then IsSectionBoundary = True: Exit Function
    If StartsWithCI(sL, "encumbrance:") Then IsSectionBoundary = True: Exit Function

    If StartsWithCI(sL, "name:") Then IsSectionBoundary = True: Exit Function
    If StartsWithCI(sL, "race:") Then IsSectionBoundary = True: Exit Function
    If StartsWithCI(sL, "class:") Then IsSectionBoundary = True: Exit Function

    If Left$(sL, 1) = "[" Then IsSectionBoundary = True: Exit Function
    If InStr(1, sL, "obvious exits:", vbTextCompare) > 0 Then IsSectionBoundary = True: Exit Function
    If StartsWithCI(sL, "also here:") Then IsSectionBoundary = True: Exit Function
    If StartsWithCI(sL, "you have no keys") Then IsSectionBoundary = True: Exit Function
    
    If InStr(1, sL, "*combat", vbTextCompare) > 0 Then IsSectionBoundary = True: Exit Function
End Function

Private Function ShouldStopSection(ByVal sLine As String, ByVal bPastFirst As Boolean) As Boolean
    ShouldStopSection = (bPastFirst And IsSectionBoundary(sLine))
End Function

' return earliest position (>1) where any known header appears mid-line
Private Function FindInlineBoundaryPos(ByVal sLine As String) As Long
    Dim sL As String, p As Long, best As Long
    sL = LCase$(sLine): best = 0

    p = InStr(1, sL, LCase$(HDR_KEYS)):  If p > 1 Then If best = 0 Or p < best Then best = p
    p = InStr(1, sL, LCase$(HDR_NOKEY)): If p > 1 Then If best = 0 Or p < best Then best = p
    p = InStr(1, sL, LCase$(HDR_INV)):   If p > 1 Then If best = 0 Or p < best Then best = p
    p = InStr(1, sL, LCase$(HDR_NOTC)):  If p > 1 Then If best = 0 Or p < best Then best = p
    p = InStr(1, sL, "wealth:"):         If p > 1 Then If best = 0 Or p < best Then best = p
    p = InStr(1, sL, "encumbrance:"):    If p > 1 Then If best = 0 Or p < best Then best = p
    p = InStr(1, sL, "you have no keys")
    
    If p > 1 Then If best = 0 Or p < best Then best = p

    FindInlineBoundaryPos = best
End Function

Private Function StartsWithCI(ByVal s As String, ByVal prefix As String) As Boolean
    StartsWithCI = (LCase$(Left$(s, Len(prefix))) = LCase$(prefix))
End Function


' ======================================================================
' Consolidation helpers (counts, plural/singular, formatting)
' ======================================================================
Private Sub ConsolidateList(ByRef a() As String, ByVal bIsKeysList As Boolean)
On Error GoTo fail
    Dim keys() As String, counts() As Long, nKeys As Long
    Dim i As Long, nm As String, cnt As Long, idx As Long

    If Not SafeArrayHasData(a) Then Exit Sub

    nKeys = 0
    For i = LBound(a) To UBound(a)
        If LenB(a(i)) > 0 Then
            ParseCountAndName a(i), nm, cnt, bIsKeysList
            idx = FindKeyIndex(keys, nKeys, nm)
            If idx < 0 Then
                If nKeys = 0 Then
                    ReDim keys(0): ReDim counts(0)
                Else
                    ReDim Preserve keys(nKeys): ReDim Preserve counts(nKeys)
                End If
                keys(nKeys) = nm
                counts(nKeys) = cnt
                nKeys = nKeys + 1
            Else
                counts(idx) = counts(idx) + cnt
            End If
        End If
    Next i

    If nKeys = 0 Then
        ReDim a(0): a(0) = vbNullString
        Exit Sub
    End If

    ReDim a(nKeys - 1)
    For i = 0 To nKeys - 1
        If counts(i) > 1 Then
            a(i) = keys(i) & " (" & CStr(counts(i)) & ")"
        Else
            a(i) = keys(i)
        End If
    Next i
    Exit Sub
fail:
    HandleError "ConsolidateList"
End Sub

' Robust count parser:
' - Removes trailing "(digits)" groups (keeps the LAST as the explicit count)
' - Also supports leading counts ("2 black star keys")
' - For keys lists, normalizes "... keys" -> "... key" and singularizes base name
Private Sub ParseCountAndName(ByVal token As String, ByRef nameOut As String, ByRef countOut As Long, ByVal bIsKeysList As Boolean)
    Dim s As String, p As Long, firstWord As String
    Dim lastTrailing As Long, pOpen As Long, pClose As Long, inside As String

    s = Trim$(token)

    ' strip trailing punctuation
    Do While Len(s) > 0 And (Right$(s, 1) = "." Or Right$(s, 1) = ";")
        s = Left$(s, Len(s) - 1)
        s = RTrim$(s)
    Loop

    ' collect trailing "(digits)" groups; keep the LAST as the explicit count
    lastTrailing = 0
    Do
        pClose = InStrRev(s, ")")
        If pClose <= 0 Then Exit Do
        pOpen = InStrRev(s, "(", pClose)
        If pOpen <= 0 Or pOpen >= pClose Then Exit Do
        inside = Mid$(s, pOpen + 1, pClose - pOpen - 1)
        If IsNumeric(Trim$(inside)) Then
            lastTrailing = CLng(val(inside))
            s = RTrim$(Left$(s, pOpen - 1))
        Else
            Exit Do
        End If
        s = RTrim$(s)
    Loop

    ' leading count?
    p = InStr(1, s, " ")
    If p > 0 Then
        firstWord = Left$(s, p - 1)
        If IsNumeric(firstWord) Then
            countOut = CLng(val(firstWord))
            nameOut = Trim$(Mid$(s, p + 1))
        Else
            countOut = 1
            nameOut = s
        End If
    Else
        countOut = 1
        nameOut = s
    End If

    nameOut = SqueezeSpaces(Trim$(nameOut))

    ' if there was NO explicit leading count, but we found trailing "(N)", use trailing
    If countOut = 1 And lastTrailing > 1 Then countOut = lastTrailing

    If bIsKeysList Then
        If LCase$(Right$(nameOut, 5)) = " keys" Then nameOut = Left$(nameOut, Len(nameOut) - 1)
        If LCase$(Right$(nameOut, 6)) = " idols" Then nameOut = Left$(nameOut, Len(nameOut) - 1) ' example: "2 golden idols" -> "golden idol"
        If countOut > 1 Then nameOut = SingularizeSimple(nameOut)
    End If

    nameOut = LCase$(nameOut)
End Sub

Private Function SingularizeSimple(ByVal s As String) As String
    Dim L As Long
    s = Trim$(s): L = Len(s)
    If L = 0 Then SingularizeSimple = s: Exit Function

    If L >= 3 And Right$(s, 3) = "ies" Then
        SingularizeSimple = Left$(s, L - 3) & "y"
        Exit Function
    End If

    If L >= 2 And Right$(s, 2) = "es" Then
        If Right$(s, 4) = "sses" Or Right$(s, 4) = "ches" Or Right$(s, 4) = "shes" Or Right$(s, 4) = "xes" Then
            SingularizeSimple = Left$(s, L - 2)
            Exit Function
        End If
    End If

    If L >= 2 And Right$(s, 1) = "s" And Right$(s, 2) <> "ss" Then
        SingularizeSimple = Left$(s, L - 1)
    Else
        SingularizeSimple = s
    End If
End Function

Private Function FindKeyIndex(ByRef keys() As String, ByVal used As Long, ByVal nm As String) As Long
    Dim i As Long
    For i = 0 To used - 1
        If keys(i) = nm Then FindKeyIndex = i: Exit Function
    Next
    FindKeyIndex = -1
End Function

Private Sub DedupeEquipped(ByRef a() As String)
On Error GoTo fail
    Dim seen() As String, out() As String
    Dim used As Long, oUsed As Long
    Dim i As Long, keyL As String

    If Not SafeArrayHasData(a) Then Exit Sub
    used = 0: oUsed = 0
    For i = LBound(a) To UBound(a)
        keyL = LCase$(Trim$(a(i)))
        If LenB(keyL) = 0 Then GoTo nxt
        If IndexOf(seen, used, keyL) < 0 Then
            If used = 0 Then ReDim seen(0) Else ReDim Preserve seen(used)
            seen(used) = keyL: used = used + 1

            If oUsed = 0 Then ReDim out(0) Else ReDim Preserve out(oUsed)
            out(oUsed) = Trim$(a(i)): oUsed = oUsed + 1
        End If
nxt:
    Next i
    If oUsed = 0 Then
        ReDim a(0): a(0) = vbNullString
    Else
        a = out
    End If
    Exit Sub
fail:
    HandleError "DedupeEquipped"
End Sub

Private Function IndexOf(ByRef arr() As String, ByVal used As Long, ByVal keyL As String) As Long
    Dim i As Long
    For i = 0 To used - 1
        If arr(i) = keyL Then IndexOf = i: Exit Function
    Next
    IndexOf = -1
End Function


' ======================================================================
' Populate ListView (QTY column + best shop via GetItemValue)
' ======================================================================
Public Sub PopulateItemManagerFromParsed(ByRef tItems As ItemParseResult, ByRef lvReferencedLV As ListView)
On Error GoTo fail
    Dim doEquipped As Boolean, doKeys As Boolean
    Dim haveEquipped As Boolean, haveKeys As Boolean, haveInv As Boolean, haveGround As Boolean
    Dim totalToAdd As Long, x As Long

    haveEquipped = HasAnyContent(tItems.sEquipped)
    haveKeys = HasAnyContent(tItems.sKeys)
    haveInv = HasAnyContent(tItems.sInventory)
    haveGround = HasAnyContent(tItems.sGround)

    doEquipped = True
    doKeys = True

    If haveEquipped Then
        If MsgBox("Import EQUIPPED items into Item Manager?", vbQuestion Or vbYesNo, "Import Equipped?") = vbNo Then doEquipped = False
    End If
    If haveKeys Then
        If MsgBox("Import KEYS (from inventory) into Item Manager?", vbQuestion Or vbYesNo, "Import Keys?") = vbNo Then doKeys = False
    End If

    totalToAdd = 0
    If doEquipped Then totalToAdd = totalToAdd + ArrayCount(tItems.sEquipped)
    If doKeys Then totalToAdd = totalToAdd + ArrayCount(tItems.sKeys)
    If haveGround Then totalToAdd = totalToAdd + ArrayCount(tItems.sGround)
    If haveInv Then totalToAdd = totalToAdd + ArrayCount(tItems.sInventory)

    If totalToAdd > 0 And lvReferencedLV.ListItems.Count > 0 Then
        If MsgBox("Clear the Item Manager of NON-FLAGGED items first?", vbQuestion + vbYesNo + vbDefaultButton2, "Clear List?") = vbYes Then
            'lvReferencedLV.ListItems.clear
            For x = lvReferencedLV.ListItems.Count To 1 Step -1
                If lvReferencedLV.ListItems(x).ListSubItems.Count >= 2 Then
                    If Len(Trim(lvReferencedLV.ListItems(x).ListSubItems(2).Text)) = 0 Then
                    'If InStr(1, lvReferencedLV.ListItems(x).ListSubItems(2).Text, "CARRIED", vbTextCompare) = 0 _
                        'And InStr(1, lvReferencedLV.ListItems(x).ListSubItems(2).Text, "STASH", vbTextCompare) = 0 Then
                        
                        lvReferencedLV.ListItems.Remove x
                    End If
                Else
                    lvReferencedLV.ListItems.Remove x
                End If
            Next x
        End If
    End If

    If doEquipped And haveEquipped Then AddSectionItems tItems.sEquipped, "Equipped", lvReferencedLV, True, False
    If doKeys And haveKeys Then AddSectionItems tItems.sKeys, "Inventory", lvReferencedLV, False, True       ' bIsKeySection:=True
    If haveGround Then AddSectionItems tItems.sGround, "Ground", lvReferencedLV, False, False
    If haveInv Then AddSectionItems tItems.sInventory, "Inventory", lvReferencedLV, False, False

    Exit Sub
fail:
    MsgBox "PopulateItemManagerFromParsed error: " & Err.Description, vbExclamation
End Sub

Private Function ArrayCount(ByRef a() As String) As Long
    If SafeArrayHasData(a) Then ArrayCount = UBound(a) - LBound(a) + 1 Else ArrayCount = 0
End Function

' sectionName: "Equipped", "Inventory", "Ground"
Private Sub AddSectionItems(ByRef arrItems() As String, ByVal sectionName As String, _
                            ByRef lv As ListView, ByVal isEquippedSection As Boolean, _
                            ByVal bIsKeySection As Boolean)
    On Error GoTo fail
    Dim i As Long, baseName As String, qty As Long
    Dim hits() As ItemMatch, hitCount As Long
    Dim nCharm As Integer
    
    If (Not Not arrItems) = 0 Then Exit Sub
    nCharm = val(frmMain.txtCharStats(5).Tag)
    For i = LBound(arrItems) To UBound(arrItems)
        If isEquippedSection Then
            baseName = NormalizeEquippedName(arrItems(i))
            qty = 1
        Else
            ParseNameAndQty arrItems(i), baseName, qty
            If Len(baseName) = 0 Then GoTo nextI
        End If
        
        Erase hits
        hitCount = GetItemsByExactNameArr(baseName, hits)
        If hitCount = 0 Then GoTo nextI

        ' choose best single DB record for this name (and skip non-gettable)
        Dim h As Long, bestH As Long, haveBest As Boolean
        Dim bestSort As Double, sortVal As Double, dummyShop As String, dummyVal As String, dummyMore As Long, dummySellOnly As Boolean

        haveBest = False
        For h = 0 To hitCount - 1
            If hits(h).Gettable = 0 Then GoTo nextH

            ' score by cheapest BUY if any shop buys; else cheapest SELL
            Dim dummyChosenShop As Long
            sortVal = EvaluateBestPriceForHit(hits(h), nCharm, dummyShop, dummyVal, dummyMore, dummySellOnly, dummyChosenShop)
            If sortVal >= 0 Then
                If Not haveBest Or sortVal < bestSort Or (sortVal = bestSort And hits(h).Number < hits(bestH).Number) Then
                    bestSort = sortVal
                    bestH = h
                    haveBest = True
                End If
            End If
nextH:
        Next h
        If Not haveBest Then GoTo nextI

        ' add a single row, carrying qty to the QTY column
        AddListViewRowsForItem hits(bestH), sectionName, lv, qty, bIsKeySection

nextI:
    Next i
    Exit Sub
fail:
    MsgBox "AddSectionItems error: " & Err.Description, vbExclamation
End Sub


' === DB lookup ===
' Now also requires [Gettable] <> 0
Private Function GetItemsByExactNameArr(ByVal exactName As String, ByRef hits() As ItemMatch) As Long
    On Error GoTo fail
    Dim rs As DAO.Recordset
    Dim cnt As Long
    Dim itm As ItemMatch

    If tabItems Is Nothing Then Exit Function

    Set rs = tabItems.Clone
    If (rs.BOF And rs.EOF) Then rs.Close: Exit Function
    rs.MoveFirst

    Do While Not rs.EOF
        If NzLong(rs.Fields("In Game").Value) <> 0 Then
            If StrComp(Trim$(NzStr(rs.Fields("Name").Value)), Trim$(exactName), vbTextCompare) = 0 Then
                itm.Number = NzLong(rs.Fields("Number").Value)
                itm.name = NzStr(rs.Fields("Name").Value)
                itm.ItemType = NzLong(rs.Fields("ItemType").Value)
                itm.Worn = NzLong(rs.Fields("Worn").Value)
                itm.WeaponType = NzLong(rs.Fields("WeaponType").Value)
                itm.encum = NzLong(rs.Fields("Encum").Value)                      ' <- NEW
                itm.ObtainedFrom = NzStr(rs.Fields("Obtained From").Value)
                itm.Gettable = NzLong(rs.Fields("Gettable").Value)                ' <- NEW

                If cnt = 0 Then
                    ReDim hits(0)
                Else
                    ReDim Preserve hits(cnt)
                End If
                hits(cnt) = itm
                cnt = cnt + 1
            End If
        End If
        rs.MoveNext
    Loop

    rs.Close
    GetItemsByExactNameArr = cnt
    Exit Function
fail:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    GetItemsByExactNameArr = 0
End Function


' === Row creation (best shop using GetItemValue + tie-breaks) ===
Private Sub AddListViewRowsForItem(ByRef hit As ItemMatch, ByRef sectionName As String, _
                                   ByRef lv As ListView, ByVal qty As Long, _
                                   ByVal bIsKeySection As Boolean)
    On Error GoTo fail

    Dim bestShopName As String, valueCell As String
    Dim moreShops As Long, sellOnly As Boolean
    Dim selectMetric As Double
    Dim chosenShopNum As Long
    Dim nCharm As Integer
    Dim sortTagCopper As Double

    nCharm = val(frmMain.txtCharStats(5).Tag)

    ' pick the display shop/value and get the chosen shop #
    selectMetric = EvaluateBestPriceForHit(hit, nCharm, bestShopName, valueCell, moreShops, sellOnly, chosenShopNum)

    ' suffix "(+N more shops)"
    If Len(bestShopName) > 0 And moreShops > 0 Then
        bestShopName = bestShopName & " +" & CStr(moreShops) & " more"
    End If

    ' ALWAYS tag with nCopperSell (0 when no shop/value)
    If chosenShopNum > 0 Then
        Dim tv As tItemValue
        tv = GetItemValue(hit.Number, nCharm, 0, chosenShopNum, sellOnly)
        sortTagCopper = tv.nCopperSell
        If sortTagCopper < 0 Then sortTagCopper = 0
    Else
        sortTagCopper = 0
    End If
    
    AddOneRow lv, hit, sectionName, hit.encum, qty, bestShopName, valueCell, sortTagCopper, bIsKeySection
    Exit Sub
fail:
    MsgBox "AddListViewRowsForItem error: " & Err.Description, vbExclamation
End Sub


' Returns the numeric price used for selection (>=0) or -1 if no shops usable.
' BUY is preferred; ties -> lower shop #. If no BUY exists, pick best SELL (ties -> lower shop #).
' Also returns the chosen shop number via chosenShopNum so the caller can tag with nCopperSell.
' chooser: prefers BUY when available; else SELL-only; ties -> lower shop #
' Returns the numeric tag to use for sorting (ALWAYS nCopperSell, or 0)
Private Function EvaluateBestPriceForHit( _
    ByRef hit As ItemMatch, _
    ByVal nCharm As Integer, _
    ByRef bestShopName As String, _
    ByRef valueCell As String, _
    ByRef moreShops As Long, _
    ByRef sellOnly As Boolean, _
    ByRef chosenShopNum As Long) As Double

    Dim shops() As ShopToken
    Dim shopCnt As Long
    Dim i As Long
    Dim itemVal As tItemValue

    Dim haveBuy As Boolean, haveSell As Boolean
    Dim bestBuy As Double: bestBuy = -1#
    Dim bestBuyShop As Long

    Dim bestSellVal As Double
    Dim bestSellShop As Long

    ' Initialize ByRef outs
    bestShopName = ""
    valueCell = ""
    moreShops = 0
    sellOnly = False
    chosenShopNum = 0
    EvaluateBestPriceForHit = 0#

    ' Gather candidate shops
    shopCnt = ExtractShopsFromObtainedFrom(hit.ObtainedFrom, shops)
    If shopCnt <= 0 Then
        bestShopName = "none"
        ' no shops: still return 0 so value sorts correctly
        Exit Function
    End If

    moreShops = shopCnt - 1

    ' Evaluate candidates
    For i = 0 To shopCnt - 1
        itemVal = GetItemValue(hit.Number, nCharm, 0, shops(i).ShopNumber, shops(i).bNObuy)

        ' BUY candidate only if the shop allows buying and has a buy price
        If (Not shops(i).bNObuy) And itemVal.nCopperBuy > 0 Then
            If (Not haveBuy) _
               Or (itemVal.nCopperBuy < bestBuy) _
               Or (itemVal.nCopperBuy = bestBuy And shops(i).ShopNumber < bestBuyShop) Then
                haveBuy = True
                bestBuy = itemVal.nCopperBuy
                bestBuyShop = shops(i).ShopNumber
            End If
        End If

        ' SELL candidate: keep the lowest record number as the representative
        If itemVal.nCopperSell > 0 Then
            If (Not haveSell) Or (shops(i).ShopNumber < bestSellShop) Then
                haveSell = True
                bestSellShop = shops(i).ShopNumber
                bestSellVal = itemVal.nCopperSell
            End If
        End If
    Next i

    ' Prefer cheapest BUY if available
    If haveBuy Then
        chosenShopNum = bestBuyShop
        bestShopName = GetShopRoomNames(bestBuyShop, , bHideRecordNumbers)

        itemVal = GetItemValue(hit.Number, nCharm, 0, bestBuyShop, False)
        valueCell = itemVal.sFriendlyBuyShort & " / " & itemVal.sFriendlySellShort
        sellOnly = False

        ' Always return SELL value for sorting
        EvaluateBestPriceForHit = itemVal.nCopperSell
        Exit Function
    End If

    ' Fallback: SELL-ONLY
    If haveSell Then
        chosenShopNum = bestSellShop
        bestShopName = GetShopRoomNames(bestSellShop, , bHideRecordNumbers)

        itemVal = GetItemValue(hit.Number, nCharm, 0, bestSellShop, True)
        valueCell = "(sell) " & itemVal.sFriendlySellShort
        sellOnly = True

        EvaluateBestPriceForHit = itemVal.nCopperSell
        Exit Function
    End If

    ' No usable prices — show as none, sort as 0
    bestShopName = "none"
    valueCell = ""
    EvaluateBestPriceForHit = 0#
End Function


Private Sub AddOneRow(ByRef lv As ListView, ByRef hit As ItemMatch, ByVal sectionName As String, _
                      ByVal encum As Long, ByVal qty As Long, _
                      ByVal shopCell As String, ByVal valueCell As String, _
                      ByVal sortCopper As Double, ByVal bIsKey As Boolean, _
                      Optional ByVal sFlag As String)

    Dim oLI As ListItem
    Dim wornText As String
    Dim usableText As String
    Dim wornTag As Integer
    
    ' Worn text:
    If bIsKey Then
        wornText = "Key"
    Else
        Select Case hit.ItemType
            Case 0:  wornText = GetWornType(hit.Worn): wornTag = hit.Worn               ' Armour
            Case 1:  wornText = GetWeaponType(hit.WeaponType): wornTag = hit.WeaponType ' Weapon
            Case Else: wornText = "Nowhere"
        End Select
    End If

    ' Usable: Yes/No
    If frmMain.TestGlobalFilter(hit.Number) Then
        usableText = "Yes"
    Else
        usableText = "No"
    End If
    
    Dim sFlagBase As String, tmpQty As Long
    Call ParseActionAndQty(NzStr(sFlag), sFlagBase, tmpQty) ' lives in modListViewExt
    
    Set oLI = lv.ListItems.Add()
    Call LV_AssignRowSeqIfMissing(lv, oLI)
    oLI.Text = CStr(hit.Number)                                  ' Col 1: Number

    oLI.ListSubItems.Add 1, "Name", hit.name                     ' Col 2
    oLI.ListSubItems.Add 2, "Flag", sFlagBase & IIf(tmpQty > 1, " x" & CStr(tmpQty), "")
    oLI.ListSubItems.Add 3, "QTY", CStr(qty)                     ' Col 4
    oLI.ListSubItems.Add 4, "Source", sectionName                ' Col 5
    oLI.ListSubItems.Add 5, "Enc", CStr(encum)                   ' Col 6
    oLI.ListSubItems.Add 6, "Type", GetItemType(hit.ItemType)    ' Col 7
    oLI.ListSubItems.Add 7, "Worn", wornText                     ' Col 8
    oLI.ListSubItems.Add 8, "Usable", usableText                 ' Col 9
    oLI.ListSubItems.Add 9, "Value", valueCell                   ' Col 10
    oLI.ListSubItems.Add 10, "Shop", shopCell                     ' Col 11
    
    oLI.ListSubItems(7).Tag = wornTag
    
    ' set numeric sort tag on the Value column; 0 when blank
    If LenB(valueCell) = 0 Or sortCopper <= 0 Then
        oLI.ListSubItems(9).Tag = "0"
    Else
        oLI.ListSubItems(9).Tag = CStr(IIf(sortCopper > 0, sortCopper, 0))
    End If
End Sub



' ======================================================================
' Tokenization & small helpers
' ======================================================================
Private Function HasAnyContent(ByRef a() As String) As Boolean
    HasAnyContent = SafeArrayHasData(a)
End Function


' "Name (N)" -> base name + qty; if none, qty=1
Private Sub ParseNameAndQty(ByVal raw As String, ByRef baseName As String, ByRef qty As Long)
    Dim s As String, pOpen As Long, pClose As Long, inside As String, lastTrailing As Long
    s = Trim$(raw)
    baseName = s
    qty = 1

    ' Robustly strip ALL trailing "(digits)" groups; keep LAST as qty
    lastTrailing = 0
    Do
        pClose = InStrRev(s, ")")
        If pClose <= 0 Then Exit Do
        pOpen = InStrRev(s, "(", pClose)
        If pOpen <= 0 Or pOpen >= pClose Or pClose <> Len(s) Then Exit Do
        inside = Mid$(s, pOpen + 1, pClose - pOpen - 1)
        If IsNumeric(Trim$(inside)) Then
            lastTrailing = CLng(val(inside))
            s = RTrim$(Left$(s, pOpen - 1))
        Else
            Exit Do
        End If
        s = RTrim$(s)
    Loop

    If lastTrailing > 0 Then
        qty = lastTrailing
        baseName = s
    Else
        baseName = s
        qty = 1
    End If
End Sub

' Equipped lines: "Item Name (Slot)" -> just the name
Private Function NormalizeEquippedName(ByVal raw As String) As String
    Dim s As String, pOpen As Long, pClose As Long
    s = Trim$(raw)
    pClose = InStrRev(s, ")")
    If pClose > 0 Then
        pOpen = InStrRev(s, "(", pClose)
        If pOpen > 0 And pClose = Len(s) Then
            NormalizeEquippedName = Trim$(Left$(s, pOpen - 1))
            Exit Function
        End If
    End If
    NormalizeEquippedName = s
End Function

' Splits by comma into String() without Variants
Private Sub SplitToArray(ByVal s As String, ByRef aOut() As String)
On Error GoTo fail
    Dim startPos As Long, commaPos As Long, tok As String
    Dim n As Long
    Erase aOut: n = -1: startPos = 1
    Do
        commaPos = InStr(startPos, s, ",")
        If commaPos = 0 Then
            tok = Trim$(Mid$(s, startPos))
            If LenB(tok) > 0 Then
                n = n + 1
                If n = 0 Then ReDim aOut(0) Else ReDim Preserve aOut(n)
                aOut(n) = tok
            End If
            Exit Do
        Else
            tok = Trim$(Mid$(s, startPos, commaPos - startPos))
            If LenB(tok) > 0 Then
                n = n + 1
                If n = 0 Then ReDim aOut(0) Else ReDim Preserve aOut(n)
                aOut(n) = tok
            End If
            startPos = commaPos + 1
        End If
    Loop
    Exit Sub
fail:
    HandleError "SplitToArray"
End Sub

Private Function SqueezeSpaces(ByVal s As String) As String
    Do While InStr(1, s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop
    SqueezeSpaces = s
End Function

Private Function SafeArrayHasData(ByRef a() As String) As Boolean
    If Not IsArrayInitialized(a) Then Exit Function
    Dim i As Long
    For i = LBound(a) To UBound(a)
        If LenB(Trim$(a(i))) > 0 Then SafeArrayHasData = True: Exit Function
    Next
End Function

Private Sub AddString(ByRef a() As String, ByVal s As String)
    Dim n As Long

    ' If array isn’t initialized at all, initialize and write first value.
    If Not IsArrayInitialized(a) Then
        ReDim a(0)
        a(0) = s
        Exit Sub
    End If

    ' If it’s a single empty slot, reuse it.
    If LBound(a) = 0 And UBound(a) = 0 And LenB(a(0)) = 0 Then
        a(0) = s
        Exit Sub
    End If

    ' Normal append.
    n = UBound(a) + 1
    ReDim Preserve a(n)
    a(n) = s
End Sub

Private Function IsArrayInitialized(ByRef a() As String) As Boolean
    On Error Resume Next
    Dim lb As Long, ub As Long
    lb = LBound(a): ub = UBound(a)
    If Err.Number = 0 Then IsArrayInitialized = True
    Err.clear
End Function


' Remove bracketed client inserts inside a span
Private Function StripBracketedChunks(ByVal s As String) As String
    Dim res As String, ch As String * 1, i As Long, depth As Long
    res = vbNullString: depth = 0
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = "[" Then
            depth = depth + 1
        ElseIf ch = "]" Then
            If depth > 0 Then depth = depth - 1
        Else
            If depth = 0 Then res = res & ch
        End If
    Next i
    StripBracketedChunks = res
End Function

' Filters
Private Function IsCashItem(ByVal sItem As String) As Boolean
    Dim sL As String
    sL = " " & LCase$(Trim$(sItem)) & " "
    If InStr(sL, " gold crown ") > 0 Or InStr(sL, " gold crowns ") > 0 Then IsCashItem = True: Exit Function
    If InStr(sL, " silver noble ") > 0 Or InStr(sL, " silver nobles ") > 0 Then IsCashItem = True: Exit Function
    If InStr(sL, " copper farthing ") > 0 Or InStr(sL, " copper farthings ") > 0 Then IsCashItem = True: Exit Function
    If InStr(sL, " platinum piece ") > 0 Or InStr(sL, " platinum pieces ") > 0 Then IsCashItem = True: Exit Function
    If InStr(sL, " runic coin ") > 0 Or InStr(sL, " runic coins ") > 0 Then IsCashItem = True: Exit Function
End Function

Private Function IsEquippedItem(ByVal sItem As String) As Boolean
    Dim p As Long
    If Right$(sItem, 1) <> ")" Then Exit Function
    p = InStrRev(sItem, " (")
    If p > 0 And p < Len(sItem) Then IsEquippedItem = True
End Function

Private Function IsBracketLine(ByVal sLine As String) As Boolean
    Dim t As String
    t = Trim$(sLine)
    If Len(t) >= 2 Then
        IsBracketLine = (Left$(t, 1) = "[" And Right$(t, 1) = "]")
    Else
        IsBracketLine = False
    End If
End Function

' Shop parsing (from "Obtained From")
' Recognizes shop tokens in Obtained From and fills shops() accordingly.
' Returns count. Flags:
'   bNObuy = True  => shop does NOT buy (i.e., SELL-ONLY to the shop)
'   bNObuy = False => shop buys (normal buy/sell)
Private Function ExtractShopsFromObtainedFrom(ByVal obtained As String, ByRef shops() As ShopToken) As Long
    Dim parts() As String, i As Long, t As String
    Dim tok As ShopToken, cnt As Long
    Dim hasSellFlag As Boolean
    Dim n As Long

    Erase shops
    ExtractShopsFromObtainedFrom = 0
    If Len(Trim$(obtained)) = 0 Then Exit Function

    parts = Split(obtained, ",")
    For i = LBound(parts) To UBound(parts)
        t = LCase$(SqueezeSpaces(Trim$(parts(i))))
        If Left$(t, 4) = "shop" Then
            ' Detect (sell) anywhere in the token
            hasSellFlag = (InStr(1, t, "(sell)", vbTextCompare) > 0)

            ' Pull first number after "shop"
            n = ExtractFirstNumber(t)
            If n > 0 Then
                tok.ShopNumber = n
                tok.bNObuy = hasSellFlag     ' if (sell), mark as no-buy (sell-only)
                tok.bNoSell = False          ' keep simple; your value function handles sell pricing

                If cnt = 0 Then
                    ReDim shops(0)
                Else
                    ReDim Preserve shops(cnt)
                End If
                shops(cnt) = tok
                cnt = cnt + 1
            End If
        End If
    Next i

    ExtractShopsFromObtainedFrom = cnt
End Function


Private Function ExtractFirstNumber(ByVal s As String) As Long
    Dim i As Long, digits As String, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        ElseIf Len(digits) > 0 Then
            Exit For
        End If
    Next i
    If Len(digits) > 0 Then ExtractFirstNumber = CLng(digits)
End Function

' Null helpers
Private Function NzStr(ByVal v As Variant) As String
    If IsNull(v) Then NzStr = "" Else NzStr = CStr(v)
End Function

Private Function NzLong(ByVal v As Variant) As Long
    If IsNull(v) Or v = "" Then NzLong = 0 Else NzLong = CLng(v)
End Function

' Minimal error handler (keeps module self-contained)
'Private Sub HandleError(ByVal where As String)
'    Debug.Print where & " error: "; Err.Number & " - " & Err.Description
'    Err.clear
'End Sub



'------------------------------------------------------------------------------
' GetBestShopNumForItem
'   Wrapper to get the chosen shop number for a given item number using the
'   same selection logic as EvaluateBestPriceForHit.
'
'   Returns:
'     Long  -> chosen shop number (0 if none)
'
'   Optional ByRef outs mirror EvaluateBestPriceForHit for convenience:
'     bSellOnly        -> True if result came from SELL-only fallback
'     nSortSellCopper  -> numeric SELL copper value used for sorting (or 0)
'     sBestShopName    -> friendly shop name (per GetShopRoomNames)
'     sValueCell       -> "buy/sell" or "(sell) X" like the original
'
'   Notes:
'     - Uses tabItems to fetch the item and build a minimal ItemMatch.
'     - Defaults nCharm to 0 to keep the signature “item-only”.
'     - Returns 0 if item not found or no usable shop per current logic.
'------------------------------------------------------------------------------
Public Function GetBestShopNumForItem( _
    ByVal nItemNum As Long, _
    Optional ByVal nCharm As Integer = 0, _
    Optional ByRef bSellOnly As Boolean, _
    Optional ByRef nSortSellCopper As Double, _
    Optional ByRef sBestShopName As String, _
    Optional ByRef sValueCell As String) As Long
On Error GoTo error:

    Dim hit As ItemMatch
    Dim sName As String
    Dim sVal As String
    Dim lMore As Long
    Dim bSell As Boolean
    Dim lChosen As Long
    Dim dSortTag As Double

    ' Defaults
    bSellOnly = False
    nSortSellCopper = 0#
    sBestShopName = ""
    sValueCell = ""
    GetBestShopNumForItem = 0

    ' Quick guards
    If nItemNum = 0 Then Exit Function
    If tabItems Is Nothing Then Exit Function
    If tabItems.RecordCount = 0 Then Exit Function

    ' Seek the item by record number (mirrors GetItemName pattern)
    On Error GoTo seek2:
    If NzLong(tabItems.Fields("Number").Value) = nItemNum Then GoTo have_row
    GoTo seekit:

seek2:
    Resume seekit:

seekit:
    On Error GoTo error:
    tabItems.Index = "pkItems"
    tabItems.Seek "=", nItemNum
    If tabItems.NoMatch Then
        tabItems.MoveFirst
        Exit Function
    End If

have_row:
    ' Build the minimal ItemMatch required by EvaluateBestPriceForHit
    hit.Number = NzLong(tabItems.Fields("Number").Value)
    hit.name = NzStr(tabItems.Fields("Name").Value)
    hit.ItemType = NzLong(tabItems.Fields("ItemType").Value)
    hit.Worn = NzLong(tabItems.Fields("Worn").Value)
    hit.WeaponType = NzLong(tabItems.Fields("WeaponType").Value)
    hit.encum = NzLong(tabItems.Fields("Encum").Value)
    hit.ObtainedFrom = NzStr(tabItems.Fields("Obtained From").Value)
    hit.Gettable = NzLong(tabItems.Fields("Gettable").Value)

    ' Defer to the authoritative chooser so results always match
    dSortTag = EvaluateBestPriceForHit(hit, nCharm, sName, sVal, lMore, bSell, lChosen)

    ' Populate outs
    sBestShopName = sName
    sValueCell = sVal
    bSellOnly = bSell
    nSortSellCopper = dSortTag

    ' Return chosen shop number (0 if none)
    GetBestShopNumForItem = lChosen

out:
    Exit Function
error:
    Call HandleError("GetBestShopNumForItem")
    Resume out:
End Function


' Adds a single row to the Item Manager for a known item record number.
' Populates all columns consistently with AddSectionItems/AddListViewRowsForItem/AddOneRow,
' then sets the Flag column to the supplied value ("CARRIED" or "STASH"), and QTY to nQTY.
'
' Inputs:
'   nItemNumber  - [tabItems].[Number]
'   sFlag        - "CARRIED" or "STASH" (placed in Flag column)
'   nQTY         - quantity for the QTY column
'
' Notes/Assumptions:
'   • Uses frmMain.lvItemManager as the target ListView.
'   • Honors your current “best shop/value” chooser via EvaluateBestPriceForHit.
'   • Numeric sort tag for Value column matches your existing behavior (nCopperSell or 0).
'   • Skips items not found or not gettable ([Gettable]=0), matching your other import paths.
Public Sub LV_AddRowByItemNumber(ByVal nItemNumber As Long, Optional ByVal sSource As String, _
    Optional ByVal sFlag As String, Optional ByVal nQTY As Integer, _
    Optional ByVal nForceShop As Long)
On Error GoTo error:

    Dim lv As ListView
    Set lv = frmMain.lvItemManager   ' Assumption: this is the Item Manager LV you want to add to

    If nItemNumber = 0 Then Exit Sub
    If tabItems Is Nothing Then Exit Sub

    ' -------- Lookup the item (mirrors your seek pattern and the GetBestShopNumForItem internals) --------
    Dim hit As ItemMatch

    On Error GoTo seek2:
    If NzLong(tabItems.Fields("Number").Value) = nItemNumber Then GoTo have_row
    GoTo seekit:

seek2:
    Resume seekit:

seekit:
    On Error GoTo error:
    tabItems.Index = "pkItems"
    tabItems.Seek "=", nItemNumber
    If tabItems.NoMatch Then
        tabItems.MoveFirst
        Exit Sub
    End If

have_row:
    If nQTY < 0 Then nQTY = 1
    If nQTY > 9999 Then nQTY = 9999
    
    ' Fill the ItemMatch with all columns AddOneRow depends upon
    hit.Number = NzLong(tabItems.Fields("Number").Value)
    hit.name = NzStr(tabItems.Fields("Name").Value)
    hit.ItemType = NzLong(tabItems.Fields("ItemType").Value)
    hit.Worn = NzLong(tabItems.Fields("Worn").Value)
    hit.WeaponType = NzLong(tabItems.Fields("WeaponType").Value)
    hit.encum = NzLong(tabItems.Fields("Encum").Value)
    hit.ObtainedFrom = NzStr(tabItems.Fields("Obtained From").Value)
    hit.Gettable = NzLong(tabItems.Fields("Gettable").Value)

    ' Respect your prior filter: skip non-gettable items
    If hit.Gettable = 0 Then Exit Sub

    ' -------- Choose best Shop/Value exactly like AddListViewRowsForItem --------
    Dim nCharm As Integer
    Dim bestShopName As String
    Dim valueCell As String
    'Dim moreShops As Long
    Dim sellOnly As Boolean
    Dim chosenShopNum As Long
    Dim sortTagCopper As Double

    nCharm = val(frmMain.txtCharStats(5).Tag)
    
    ' This function internally scores using EvaluateBestPriceForHit to stay consistent
    chosenShopNum = GetBestShopNumForItem( _
                        nItemNumber, _
                        nCharm, _
                        sellOnly, _
                        sortTagCopper, _
                        bestShopName, _
                        valueCell)
    
    If nForceShop > 0 And chosenShopNum <> nForceShop Then
        chosenShopNum = nForceShop
        bestShopName = GetShopRoomNames(nForceShop, , bHideRecordNumbers)
    End If
    
    ' ALWAYS tag the Value column with SELL copper (0 when none), matching your existing behavior
    If chosenShopNum > 0 Then
        Dim tv As tItemValue
        tv = GetItemValue(hit.Number, nCharm, 0, chosenShopNum, sellOnly)
        sortTagCopper = tv.nCopperSell
        If sortTagCopper < 0 Then sortTagCopper = 0
    Else
        sortTagCopper = 0
    End If

    ' -------- Add the row using the same builder you already use --------
    ' Source column: follow your import convention—default to "Inventory"
    Call AddOneRow(lv, hit, sSource, hit.encum, nQTY, bestShopName, valueCell, sortTagCopper, (hit.ItemType = 7), sFlag)

'    ' Stamp the Flag column on the newly added row
'    Dim oLI As ListItem
'    Set oLI = lv.ListItems(lv.ListItems.Count)
'    If Not (oLI Is Nothing) Then
'        If oLI.ListSubItems.Count >= 2 Then
'            oLI.ListSubItems(2).Text = sFlag
'        End If
'    End If

    Exit Sub

error:
    Call HandleError("LV_AddRowByItemNumber")
End Sub


