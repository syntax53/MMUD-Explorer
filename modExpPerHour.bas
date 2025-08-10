Attribute VB_Name = "modExpPerHour"
#Const DEVELOPMENT_MODE = 0 'TURN OFF BEFORE RELEASE
Option Explicit
Option Base 0

Private Const SEC_PER_ROUND      As Double = 5#
Private Const SEC_PER_REST_TICK  As Double = 20#
Private Const SEC_PER_REGEN_TICK As Double = 30#
Private Const SEC_PER_MEDI_TICK  As Double = 10#
Private Const SECS_ROOM_BASE     As Double = 1.2
Private Const SECS_ROOM_HEAVY    As Double = 1.8
Private Const HEAVY_ENCUM_PCT    As Double = 67#

Global bGlobalExpHrKnobsByChar As Boolean
Global nGlobalExpHrModel As eCalcExpModel

Global nGlobal_cephA_DMG            As Double
Global nGlobal_cephA_Mana           As Double
Global nGlobal_cephA_MoveRecover    As Double
Global nGlobal_cephA_Move           As Double
Global nGlobal_cephA_ClusterMx      As Integer

Public nGlobal_cephB_XP     As Double
Public nGlobal_cephB_DMG    As Double
Public nGlobal_cephB_Mana   As Double
Public nGlobal_cephB_Move   As Double

Private Const cephB_LOGISTIC_CAP      As Double = 700#
Private Const cephB_LOGISTIC_DENOM    As Double = 0.5
Private Const cephB_MIN_LOOP          As Double = 22#
Private Const cephB_TF_LOG_COEF       As Double = 0.15
Private Const cephB_TF_SMALL_BUMP     As Double = 0.7
Private Const cephB_TF_SCARCITY_COEF  As Double = 0.15
Private Const cephB_LAIR_OVERHEAD_R   As Double = 1

Public Type tExpPerHourInfo
    nExpPerHour As Double
    nHitpointRecovery As Double
    nManaRecovery As Double
    nTimeRecovering As Double
    nOverkill As Double
    nMove As Double
    nRTC As Double
    nAttackTime As Double
    nSlowdownTime As Double
    nRoamTime As Double
    sHitpointRecovery As String
    sManaRecovery As String
    sTimeRecovering As String
    sRTCText As String
    sMoveText As String
End Type

Public Enum eCalcExpModel
    default = 0
    average = 1
    modelA = 2
    modelB = 3
End Enum


Public Function CalcExpPerHour( _
    Optional ByVal nExp As Currency, Optional ByVal nRegenTime As Double, Optional ByVal nNumMobs As Integer, _
    Optional ByVal nTotalLairs As Long = -1, Optional ByVal nPossSpawns As Long, Optional ByVal nRTK As Double, _
    Optional ByVal nCharDMG As Double, Optional ByVal nCharHP As Long, Optional ByVal nCharHPRegen As Long, _
    Optional ByVal nMobDmg As Double, Optional ByVal nMobHP As Long, Optional ByVal nMobHPRegen As Long, _
    Optional ByVal nDamageThreshold As Long, Optional ByVal nSpellCost As Integer, _
    Optional ByVal nSpellOverhead As Double, Optional ByVal nCharMana As Long, _
    Optional ByVal nCharMPRegen As Long, Optional ByVal nMeditateRate As Long, _
    Optional ByVal nAvgWalk As Double, Optional ByVal nEncumPct As Integer, _
    Optional ByVal eModel As eCalcExpModel = 0) As tExpPerHourInfo

'Function input details...
'nExp = Exp per kill/clear ((nExp / nNumMobs) = per mob exp)
'nRegenTime = Regen time of each lair or single monster
'nNumMobs = Number of mobs in each lair (i.e. number of mobs represented in nExp, nMobDmg, nMobHP, and nMobHP)
'nTotalLairs = Total number of lairs that spawn the monster
'nPossSpawns = Total number of rooms around the lairs in the same group/index
'nRTK = Rounds to kill a SINGLE MONSTER (e.g. nRTK * nNumMobs = nRTC [rounds to clear lair])
'nCharDMG = Character/party output damage against monster
'nCharHP = Character/party total hitpoints
'nCharHPRegen = Character/party hitpoint regen
'nMobDmg = Monster damage against character/party ((nMobDmg / nNumMobs) = per mob dmg)
'nMobHP = Monster total hitpoints ((nMobHP / nNumMobs) = per mob hp)
'nMobHPRegen = Monster hp regen ((nMobHPRegen / nNumMobs) = per mob regen)
'nDamageThreshold = Damage threshold where a player will need to eventually recover hitpoints.
'   Meaning, if a value of 10 was specified, then the player should be able to sustain 10 damage per round without every needing to rest.
'nSpellCost = Cost of the main spell attack, if applicable
'nSpellOverhead = Cost of any per-round spell upkeep from bless spells and/or healing spells
'nCharMana = Character total mana
'nCharMPRegen = Character mana regen rate
'nMeditateRate = Character meditate rate, if applicable
'nAvgWalk = Average walking room distance from lair to lair
'nEncumPCT = Weight of character (>= 67 is heavy)

Dim tRetA As tExpPerHourInfo, tRetB As tExpPerHourInfo, eDefault As eCalcExpModel
Dim tRet As tExpPerHourInfo, bMovementLimited As Boolean

eDefault = nGlobalExpHrModel
If eDefault = default Then eDefault = average
If eModel = default Then eModel = eDefault

If eModel = modelA Or eModel = average Then
    tRetA = ceph_ModelA( _
        nExp, nRegenTime, nNumMobs, nTotalLairs, nPossSpawns, nRTK, _
        nCharDMG, nCharHP, nCharHPRegen, nMobDmg, nMobHP, nMobHPRegen, _
        nDamageThreshold, nSpellCost, nSpellOverhead, nCharMana, nCharMPRegen, nMeditateRate, nAvgWalk, nEncumPct)
    If tRetA.nMove < 0 Then
        bMovementLimited = True
        tRetA.nMove = tRetA.nMove * -1
    End If
End If

If eModel = modelB Or eModel = average Then
    tRetB = ceph_ModelB( _
        nExp, nRegenTime, nNumMobs, nTotalLairs, nPossSpawns, nRTK, _
        nCharDMG, nCharHP, nCharHPRegen, nMobDmg, nMobHP, nMobHPRegen, _
        nDamageThreshold, nSpellCost, nSpellOverhead, nCharMana, nCharMPRegen, nMeditateRate, nAvgWalk, nEncumPct)
    If tRetB.nMove < 0 Then
        bMovementLimited = True
        tRetB.nMove = tRetB.nMove * -1
    End If
End If

If eModel = average Then
    tRet.nExpPerHour = Round((tRetA.nExpPerHour + tRetB.nExpPerHour) / 2)
    tRet.nHitpointRecovery = Round((tRetA.nHitpointRecovery + tRetB.nHitpointRecovery) / 2, 2)
    tRet.nManaRecovery = Round((tRetA.nManaRecovery + tRetB.nManaRecovery) / 2, 2)
    tRet.nTimeRecovering = Round((tRetA.nTimeRecovering + tRetB.nTimeRecovering) / 2, 2)
    tRet.nOverkill = Round((tRetA.nOverkill + tRetB.nOverkill) / 2, 2)
    tRet.nMove = Round((tRetA.nMove + tRetB.nMove) / 2, 2)
    tRet.nRTC = Round((tRetA.nRTC + tRetB.nRTC) / 2, 2)
    tRet.nRoamTime = Round((tRetA.nRoamTime + tRetB.nRoamTime) / 2, 2)
    tRet.nSlowdownTime = Round((tRetA.nSlowdownTime + tRetB.nSlowdownTime) / 2, 2)
    tRet.nAttackTime = Round((tRetA.nAttackTime + tRetB.nAttackTime) / 2, 2)
    tRet.nRTC = Round((tRetA.nRTC + tRetB.nRTC) / 2, 2)
ElseIf eModel = modelA Then
    tRet = tRetA
ElseIf eModel = modelB Then
    tRet = tRetB
End If
    
If tRet.nAttackTime > 0 And tRet.nAttackTime < 1 Then tRet.sRTCText = Round(tRet.nAttackTime * 100) & "% time spent attacking"
If tRet.nSlowdownTime > 0.01 And tRet.nSlowdownTime < 1 Then tRet.sRTCText = AutoAppend(tRet.sRTCText, Round(tRet.nSlowdownTime * 100) & "% slower kill speed")
If tRet.nOverkill > 0.01 And nCharDMG < 9999999 Then tRet.sRTCText = AutoAppend(tRet.sRTCText, Round(tRet.nOverkill * 100) & "% wasted overkill")

If tRet.nTimeRecovering > 0.01 Then tRet.sTimeRecovering = Round(tRet.nTimeRecovering * 100) & "% time spent recovering"
If tRet.nHitpointRecovery > 0.01 Then tRet.sHitpointRecovery = Round(tRet.nHitpointRecovery * 100) & "% reduction due to HP recovery"
If tRet.nManaRecovery > 0.01 Then tRet.sManaRecovery = Round(tRet.nManaRecovery * 100) & "% reduction due to mana recovery"

If tRet.nMove > 0.01 Then tRet.sMoveText = Round(tRet.nMove * 100) & "% time spent moving"
If tRet.nRoamTime > 0.04 Then tRet.sMoveText = AutoAppend(tRet.sMoveText, Round(tRet.nRoamTime * 100) & "% time lost due to insufficient lairs")
If bMovementLimited Then tRet.sMoveText = AutoAppend(tRet.sMoveText, "(cluster detected: movement limited)", " ")

CalcExpPerHour = tRet

End Function

Public Sub RunAllSimulations()
    Dim result As tExpPerHourInfo
    Dim summaries() As String, observations() As Double
    Dim i As Integer, nTest As Integer, nMaxDesc As Integer
    Dim nAvg(1 To 4) As Double, nAvgCount(1 To 4) As Integer
    
    Dim nTotalObs As Integer
    nTotalObs = 17
    ReDim observations(1 To nTotalObs, 1 To 4)
    ReDim summaries(1 To nTotalObs)
    
    'exp
    observations(1, 1) = 7174000
    observations(2, 1) = 909000
    observations(3, 1) = 890000
    observations(4, 1) = 325000
    observations(5, 1) = 1181000
    observations(6, 1) = 20000
    observations(7, 1) = 52000
    observations(8, 1) = 34000
    observations(9, 1) = 75000
    observations(10, 1) = 82000
    observations(11, 1) = 847000
    observations(12, 1) = 633000
    observations(13, 1) = 409000
    observations(14, 1) = 1942000
    observations(15, 1) = 1818000
    observations(16, 1) = 987000
    observations(17, 1) = 2231000
    'rest
    observations(1, 2) = 0
    observations(2, 2) = 0
    observations(3, 2) = 0
    observations(4, 2) = 0
    observations(5, 2) = 50 / 100
    observations(6, 2) = 32 / 100
    observations(7, 2) = 1 / 100
    observations(8, 2) = 0
    observations(9, 2) = 0
    observations(10, 2) = 4 / 100
    observations(11, 2) = 2 / 100
    observations(12, 2) = 1 / 100
    observations(13, 2) = 37 / 100
    observations(14, 2) = 29 / 100
    observations(15, 2) = 31 / 100
    observations(16, 2) = 0
    observations(17, 2) = 67 / 100
    'mana
    observations(1, 3) = 0
    observations(2, 3) = 16 / 100
    observations(3, 3) = 39 / 100
    observations(4, 3) = 29 / 100
    observations(5, 3) = 0
    observations(6, 3) = 43 / 100
    observations(7, 3) = 52 / 100
    observations(8, 3) = 75 / 100
    observations(9, 3) = 0
    observations(10, 3) = 0
    observations(11, 3) = 0
    observations(12, 3) = 0
    observations(13, 3) = 0
    observations(14, 3) = 0
    observations(15, 3) = 0
    observations(16, 3) = 0
    observations(17, 3) = 0
    'move
    observations(1, 4) = 17 / 100
    observations(2, 4) = 55 / 100
    observations(3, 4) = 36 / 100
    observations(4, 4) = 11 / 100
    observations(5, 4) = 7 / 100
    observations(6, 4) = 5 / 100
    observations(7, 4) = 10 / 100
    observations(8, 4) = 5 / 100
    observations(9, 4) = 13 / 100
    observations(10, 4) = 7 / 100
    observations(11, 4) = 16 / 100
    observations(12, 4) = 9 / 100
    observations(13, 4) = 4 / 100
    observations(14, 4) = 34 / 100
    observations(15, 4) = 31 / 100
    observations(16, 4) = 62 / 100
    observations(17, 4) = 3 / 100
    
    DebugLogPrint "=== Running All CalcExpPerHour Simulations ==="
    
    ' Helper: RunSim simulates one run and returns summary line
    Dim simIndex As Integer
    simIndex = 0
    
    'frmMain.mnuLairLimitMovement.Checked = False
    
    ' --- Simulation 1 ---
    simIndex = simIndex + 1
    result = CalcExpPerHour(54125, 3, 3, 48, 85, 1.5, 619, 1952, 158, 51, 1800, 0, 50, 0, 0, 0, 0, 0, 1.8, 0)
    summaries(simIndex) = "SIM" & simIndex & ": 135 Cleric/manscorpions/physical/50 heal: " & FormatResult(result, observations, simIndex)
    DebugLogPrint summaries(simIndex)
    If result.nExpPerHour > 0 Then nAvg(1) = nAvg(1) + (1 - (observations(simIndex, 1) / result.nExpPerHour)) Else nAvg(1) = nAvg(1) + 1
    If result.nHitpointRecovery <> 0 Or observations(simIndex, 2) <> 0 Then nAvg(2) = nAvg(2) + result.nHitpointRecovery - observations(simIndex, 2): nAvgCount(2) = nAvgCount(2) + 1
        If result.nManaRecovery <> 0 Or observations(simIndex, 3) <> 0 Then nAvg(3) = nAvg(3) + result.nManaRecovery - observations(simIndex, 3): nAvgCount(3) = nAvgCount(3) + 1
                If result.nMove <> 0 Or observations(simIndex, 4) <> 0 Then nAvg(4) = nAvg(4) + result.nMove - observations(simIndex, 4): nAvgCount(4) = nAvgCount(4) + 1
    DebugLogPrint "-----------------------------------------------"
    
    ' --- Simulation 2 ---
    simIndex = simIndex + 1
    result = CalcExpPerHour(3171, 3, 1, 35, 100, 1, 564, 891, 67, 4, 238, 0, 0, 30, 1, 623, 63, 42, 2.9, 0)
    summaries(simIndex) = "SIM" & simIndex & ": 81 Priest/stone elementals/srip+MEDITATE/no heal: " & FormatResult(result, observations, simIndex)
    DebugLogPrint summaries(simIndex)
    If result.nExpPerHour > 0 Then nAvg(1) = nAvg(1) + (1 - (observations(simIndex, 1) / result.nExpPerHour)) Else nAvg(1) = nAvg(1) + 1
    If result.nHitpointRecovery <> 0 Or observations(simIndex, 2) <> 0 Then nAvg(2) = nAvg(2) + result.nHitpointRecovery - observations(simIndex, 2): nAvgCount(2) = nAvgCount(2) + 1
        If result.nManaRecovery <> 0 Or observations(simIndex, 3) <> 0 Then nAvg(3) = nAvg(3) + result.nManaRecovery - observations(simIndex, 3): nAvgCount(3) = nAvgCount(3) + 1
                If result.nMove <> 0 Or observations(simIndex, 4) <> 0 Then nAvg(4) = nAvg(4) + result.nMove - observations(simIndex, 4): nAvgCount(4) = nAvgCount(4) + 1
    DebugLogPrint "-----------------------------------------------"
    
    ' --- Simulation 3 ---
    simIndex = simIndex + 1
    result = CalcExpPerHour(3171, 3, 1, 35, 100, 1, 564, 891, 67, 4, 238, 0, 0, 30, 1, 623, 63, 0, 2.9, 0)
    summaries(simIndex) = "SIM" & simIndex & ": 81 Priest/stone elementals/srip/no heal: " & FormatResult(result, observations, simIndex)
    DebugLogPrint summaries(simIndex)
    If result.nExpPerHour > 0 Then nAvg(1) = nAvg(1) + (1 - (observations(simIndex, 1) / result.nExpPerHour)) Else nAvg(1) = nAvg(1) + 1
    If result.nHitpointRecovery <> 0 Or observations(simIndex, 2) <> 0 Then nAvg(2) = nAvg(2) + result.nHitpointRecovery - observations(simIndex, 2): nAvgCount(2) = nAvgCount(2) + 1
        If result.nManaRecovery <> 0 Or observations(simIndex, 3) <> 0 Then nAvg(3) = nAvg(3) + result.nManaRecovery - observations(simIndex, 3): nAvgCount(3) = nAvgCount(3) + 1
                If result.nMove <> 0 Or observations(simIndex, 4) <> 0 Then nAvg(4) = nAvg(4) + result.nMove - observations(simIndex, 4): nAvgCount(4) = nAvgCount(4) + 1
    DebugLogPrint "-----------------------------------------------"
    
    ' --- Simulation 4 ---
    'frmMain.mnuLairLimitMovement.Checked = True
    simIndex = simIndex + 1
    result = CalcExpPerHour(2071, 2, 3, 13, 1223, 1, 255, 891, 67, 1, 339, 0, 0, 16, 1, 623, 63, 0, 1.3, 0)
    summaries(simIndex) = "SIM" & simIndex & ": 81 Priest/gnolls+LIMIT_MOVE/fury/no heal: " & FormatResult(result, observations, simIndex)
    'frmMain.mnuLairLimitMovement.Checked = False
    DebugLogPrint summaries(simIndex)
    If result.nExpPerHour > 0 Then nAvg(1) = nAvg(1) + (1 - (observations(simIndex, 1) / result.nExpPerHour)) Else nAvg(1) = nAvg(1) + 1
    If result.nHitpointRecovery <> 0 Or observations(simIndex, 2) <> 0 Then nAvg(2) = nAvg(2) + result.nHitpointRecovery - observations(simIndex, 2): nAvgCount(2) = nAvgCount(2) + 1
        If result.nManaRecovery <> 0 Or observations(simIndex, 3) <> 0 Then nAvg(3) = nAvg(3) + result.nManaRecovery - observations(simIndex, 3): nAvgCount(3) = nAvgCount(3) + 1
                If result.nMove <> 0 Or observations(simIndex, 4) <> 0 Then nAvg(4) = nAvg(4) + result.nMove - observations(simIndex, 4): nAvgCount(4) = nAvgCount(4) + 1
    DebugLogPrint "-----------------------------------------------"
    
    ' --- Simulation 5 ---
    simIndex = simIndex + 1
    result = CalcExpPerHour(19650, 3, 3, 20, 26, 1.5, 232, 891, 67, 40, 876, 0, 0, 16, 1, 623, 63, 0, 1.3, 0)
    summaries(simIndex) = "SIM" & simIndex & ": 81 Priest/white dragons/fury/no heal: " & FormatResult(result, observations, simIndex)
    DebugLogPrint summaries(simIndex)
    If result.nExpPerHour > 0 Then nAvg(1) = nAvg(1) + (1 - (observations(simIndex, 1) / result.nExpPerHour)) Else nAvg(1) = nAvg(1) + 1
    If result.nHitpointRecovery <> 0 Or observations(simIndex, 2) <> 0 Then nAvg(2) = nAvg(2) + result.nHitpointRecovery - observations(simIndex, 2): nAvgCount(2) = nAvgCount(2) + 1
        If result.nManaRecovery <> 0 Or observations(simIndex, 3) <> 0 Then nAvg(3) = nAvg(3) + result.nManaRecovery - observations(simIndex, 3): nAvgCount(3) = nAvgCount(3) + 1
                If result.nMove <> 0 Or observations(simIndex, 4) <> 0 Then nAvg(4) = nAvg(4) + result.nMove - observations(simIndex, 4): nAvgCount(4) = nAvgCount(4) + 1
    DebugLogPrint "-----------------------------------------------"
    
    ' --- Simulation 6 ---
    simIndex = simIndex + 1
    result = CalcExpPerHour(857, 5, 2, 14, 239, 2.5, 34, 182, 11, 14, 166, 0, 0, 5, 0, 54, 6, 0, 12.5, 0)
    summaries(simIndex) = "SIM" & simIndex & ": 12 Gypsy/orc shaman/lbol/no heal: " & FormatResult(result, observations, simIndex)
    DebugLogPrint summaries(simIndex)
    If result.nExpPerHour > 0 Then nAvg(1) = nAvg(1) + (1 - (observations(simIndex, 1) / result.nExpPerHour)) Else nAvg(1) = nAvg(1) + 1
    If result.nHitpointRecovery <> 0 Or observations(simIndex, 2) <> 0 Then nAvg(2) = nAvg(2) + result.nHitpointRecovery - observations(simIndex, 2): nAvgCount(2) = nAvgCount(2) + 1
        If result.nManaRecovery <> 0 Or observations(simIndex, 3) <> 0 Then nAvg(3) = nAvg(3) + result.nManaRecovery - observations(simIndex, 3): nAvgCount(3) = nAvgCount(3) + 1
                If result.nMove <> 0 Or observations(simIndex, 4) <> 0 Then nAvg(4) = nAvg(4) + result.nMove - observations(simIndex, 4): nAvgCount(4) = nAvgCount(4) + 1
    DebugLogPrint "-----------------------------------------------"
    
    ' --- Simulation 7 ---
    simIndex = simIndex + 1
    result = CalcExpPerHour(690, 3, 3, 33, 83, 1, 59, 198, 15, 2, 120, 0, 0, 8, 0, 131, 11, 9, 2.4, 0)
    summaries(simIndex) = "SIM" & simIndex & ": 20 Druid/kobolds/acid+MEDITATE/no heal: " & FormatResult(result, observations, simIndex)
    DebugLogPrint summaries(simIndex)
    If result.nExpPerHour > 0 Then nAvg(1) = nAvg(1) + (1 - (observations(simIndex, 1) / result.nExpPerHour)) Else nAvg(1) = nAvg(1) + 1
    If result.nHitpointRecovery <> 0 Or observations(simIndex, 2) <> 0 Then nAvg(2) = nAvg(2) + result.nHitpointRecovery - observations(simIndex, 2): nAvgCount(2) = nAvgCount(2) + 1
        If result.nManaRecovery <> 0 Or observations(simIndex, 3) <> 0 Then nAvg(3) = nAvg(3) + result.nManaRecovery - observations(simIndex, 3): nAvgCount(3) = nAvgCount(3) + 1
                If result.nMove <> 0 Or observations(simIndex, 4) <> 0 Then nAvg(4) = nAvg(4) + result.nMove - observations(simIndex, 4): nAvgCount(4) = nAvgCount(4) + 1
    DebugLogPrint "-----------------------------------------------"
    
    ' --- Simulation 8 ---
    simIndex = simIndex + 1
    result = CalcExpPerHour(690, 3, 3, 33, 83, 1, 59, 198, 15, 2, 120, 0, 0, 8, 0, 131, 11, 0, 2.4, 0)
    summaries(simIndex) = "SIM" & simIndex & ": 20 Druid/kobolds/acid/no heal: " & FormatResult(result, observations, simIndex)
    DebugLogPrint summaries(simIndex)
    If result.nExpPerHour > 0 Then nAvg(1) = nAvg(1) + (1 - (observations(simIndex, 1) / result.nExpPerHour)) Else nAvg(1) = nAvg(1) + 1
    If result.nHitpointRecovery <> 0 Or observations(simIndex, 2) <> 0 Then nAvg(2) = nAvg(2) + result.nHitpointRecovery - observations(simIndex, 2): nAvgCount(2) = nAvgCount(2) + 1
        If result.nManaRecovery <> 0 Or observations(simIndex, 3) <> 0 Then nAvg(3) = nAvg(3) + result.nManaRecovery - observations(simIndex, 3): nAvgCount(3) = nAvgCount(3) + 1
                If result.nMove <> 0 Or observations(simIndex, 4) <> 0 Then nAvg(4) = nAvg(4) + result.nMove - observations(simIndex, 4): nAvgCount(4) = nAvgCount(4) + 1
    DebugLogPrint "-----------------------------------------------"
    
    ' --- Simulation 9 ---
    simIndex = simIndex + 1
    result = CalcExpPerHour(690, 3, 3, 33, 83, 1.5, 41, 198, 15, 2, 120, 0, 10, 0, 0, 0, 0, 0, 2.4, 0)
    summaries(simIndex) = "SIM" & simIndex & ": 20 Druid/kobolds/physical/10 heal: " & FormatResult(result, observations, simIndex)
    DebugLogPrint summaries(simIndex)
    If result.nExpPerHour > 0 Then nAvg(1) = nAvg(1) + (1 - (observations(simIndex, 1) / result.nExpPerHour)) Else nAvg(1) = nAvg(1) + 1
    If result.nHitpointRecovery <> 0 Or observations(simIndex, 2) <> 0 Then nAvg(2) = nAvg(2) + result.nHitpointRecovery - observations(simIndex, 2): nAvgCount(2) = nAvgCount(2) + 1
        If result.nManaRecovery <> 0 Or observations(simIndex, 3) <> 0 Then nAvg(3) = nAvg(3) + result.nManaRecovery - observations(simIndex, 3): nAvgCount(3) = nAvgCount(3) + 1
                If result.nMove <> 0 Or observations(simIndex, 4) <> 0 Then nAvg(4) = nAvg(4) + result.nMove - observations(simIndex, 4): nAvgCount(4) = nAvgCount(4) + 1
    DebugLogPrint "-----------------------------------------------"
    
    ' --- Simulation 10 ---
    simIndex = simIndex + 1
    result = CalcExpPerHour(500, 0, 0, -1, 0, 0, 56, 198, 15, 8, 170, 25, 10, 0, 0, 0, 0, 0, 0, 0)
    summaries(simIndex) = "SIM" & simIndex & ": 20 Druid/slime beast/physical/10 heal: " & FormatResult(result, observations, simIndex)
    DebugLogPrint summaries(simIndex)
    If result.nExpPerHour > 0 Then nAvg(1) = nAvg(1) + (1 - (observations(simIndex, 1) / result.nExpPerHour)) Else nAvg(1) = nAvg(1) + 1
    If result.nHitpointRecovery <> 0 Or observations(simIndex, 2) <> 0 Then nAvg(2) = nAvg(2) + result.nHitpointRecovery - observations(simIndex, 2): nAvgCount(2) = nAvgCount(2) + 1
        If result.nManaRecovery <> 0 Or observations(simIndex, 3) <> 0 Then nAvg(3) = nAvg(3) + result.nManaRecovery - observations(simIndex, 3): nAvgCount(3) = nAvgCount(3) + 1
                If result.nMove <> 0 Or observations(simIndex, 4) <> 0 Then nAvg(4) = nAvg(4) + result.nMove - observations(simIndex, 4): nAvgCount(4) = nAvgCount(4) + 1
    DebugLogPrint "-----------------------------------------------"
    
    ' --- Simulation 11 ---
    simIndex = simIndex + 1
    result = CalcExpPerHour(8875, 3, 4, 13, 85, 1.5, 249, 585, 37, 5, 877, 0, 15, 0, 0, 0, 0, 0, 6.2, 0)
    summaries(simIndex) = "SIM" & simIndex & ": 46 Paladin/orc captains/bash/15 heal: " & FormatResult(result, observations, simIndex)
    DebugLogPrint summaries(simIndex)
    If result.nExpPerHour > 0 Then nAvg(1) = nAvg(1) + (1 - (observations(simIndex, 1) / result.nExpPerHour)) Else nAvg(1) = nAvg(1) + 1
    If result.nHitpointRecovery <> 0 Or observations(simIndex, 2) <> 0 Then nAvg(2) = nAvg(2) + result.nHitpointRecovery - observations(simIndex, 2): nAvgCount(2) = nAvgCount(2) + 1
        If result.nManaRecovery <> 0 Or observations(simIndex, 3) <> 0 Then nAvg(3) = nAvg(3) + result.nManaRecovery - observations(simIndex, 3): nAvgCount(3) = nAvgCount(3) + 1
                If result.nMove <> 0 Or observations(simIndex, 4) <> 0 Then nAvg(4) = nAvg(4) + result.nMove - observations(simIndex, 4): nAvgCount(4) = nAvgCount(4) + 1
    DebugLogPrint "-----------------------------------------------"
    
    ' --- Simulation 12 ---
    simIndex = simIndex + 1
    result = CalcExpPerHour(8875, 3, 4, 13, 85, 2, 145, 585, 37, 7, 877, 0, 15, 0, 0, 0, 0, 0, 6.2, 0)
    summaries(simIndex) = "SIM" & simIndex & ": 46 Paladin/orc captains/physical/15 heal: " & FormatResult(result, observations, simIndex)
    DebugLogPrint summaries(simIndex)
    If result.nExpPerHour > 0 Then nAvg(1) = nAvg(1) + (1 - (observations(simIndex, 1) / result.nExpPerHour)) Else nAvg(1) = nAvg(1) + 1
    If result.nHitpointRecovery <> 0 Or observations(simIndex, 2) <> 0 Then nAvg(2) = nAvg(2) + result.nHitpointRecovery - observations(simIndex, 2): nAvgCount(2) = nAvgCount(2) + 1
        If result.nManaRecovery <> 0 Or observations(simIndex, 3) <> 0 Then nAvg(3) = nAvg(3) + result.nManaRecovery - observations(simIndex, 3): nAvgCount(3) = nAvgCount(3) + 1
                If result.nMove <> 0 Or observations(simIndex, 4) <> 0 Then nAvg(4) = nAvg(4) + result.nMove - observations(simIndex, 4): nAvgCount(4) = nAvgCount(4) + 1
    DebugLogPrint "-----------------------------------------------"
    
    ' --- Simulation 13 ---
    simIndex = simIndex + 1
    result = CalcExpPerHour(8875, 3, 4, 13, 85, 2, 145, 585, 37, 7, 877, 0, 0, 0, 0, 0, 0, 0, 6.2, 0)
    summaries(simIndex) = "SIM" & simIndex & ": 46 Paladin/orc captains/physical/no heal: " & FormatResult(result, observations, simIndex)
    DebugLogPrint summaries(simIndex)
    If result.nExpPerHour > 0 Then nAvg(1) = nAvg(1) + (1 - (observations(simIndex, 1) / result.nExpPerHour)) Else nAvg(1) = nAvg(1) + 1
    If result.nHitpointRecovery <> 0 Or observations(simIndex, 2) <> 0 Then nAvg(2) = nAvg(2) + result.nHitpointRecovery - observations(simIndex, 2): nAvgCount(2) = nAvgCount(2) + 1
        If result.nManaRecovery <> 0 Or observations(simIndex, 3) <> 0 Then nAvg(3) = nAvg(3) + result.nManaRecovery - observations(simIndex, 3): nAvgCount(3) = nAvgCount(3) + 1
                If result.nMove <> 0 Or observations(simIndex, 4) <> 0 Then nAvg(4) = nAvg(4) + result.nMove - observations(simIndex, 4): nAvgCount(4) = nAvgCount(4) + 1
    DebugLogPrint "-----------------------------------------------"
    
    ' --- Simulation 14 ---
    simIndex = simIndex + 1
    result = CalcExpPerHour(20236, 3, 3, 7, 53, 1, 566, 1175, 81, 24, 928, 0, 0, 0, 0, 0, 0, 0, 10.4, 0)
    summaries(simIndex) = "SIM" & simIndex & ": 75 Warrior/white dragons/bash/no heal: " & FormatResult(result, observations, simIndex)
    DebugLogPrint summaries(simIndex)
    If result.nExpPerHour > 0 Then nAvg(1) = nAvg(1) + (1 - (observations(simIndex, 1) / result.nExpPerHour)) Else nAvg(1) = nAvg(1) + 1
    If result.nHitpointRecovery <> 0 Or observations(simIndex, 2) <> 0 Then nAvg(2) = nAvg(2) + result.nHitpointRecovery - observations(simIndex, 2): nAvgCount(2) = nAvgCount(2) + 1
        If result.nManaRecovery <> 0 Or observations(simIndex, 3) <> 0 Then nAvg(3) = nAvg(3) + result.nManaRecovery - observations(simIndex, 3): nAvgCount(3) = nAvgCount(3) + 1
                If result.nMove <> 0 Or observations(simIndex, 4) <> 0 Then nAvg(4) = nAvg(4) + result.nMove - observations(simIndex, 4): nAvgCount(4) = nAvgCount(4) + 1
    DebugLogPrint "-----------------------------------------------"
    
    ' --- Simulation 15 ---
    simIndex = simIndex + 1
    result = CalcExpPerHour(20236, 3, 3, 7, 53, 1.1, 522, 1175, 81, 24, 928, 0, 0, 0, 0, 0, 0, 0, 10.4, 0)
    summaries(simIndex) = "SIM" & simIndex & ": 75 Warrior/white dragons/physical/no heal: " & FormatResult(result, observations, simIndex)
    DebugLogPrint summaries(simIndex)
    If result.nExpPerHour > 0 Then nAvg(1) = nAvg(1) + (1 - (observations(simIndex, 1) / result.nExpPerHour)) Else nAvg(1) = nAvg(1) + 1
    If result.nHitpointRecovery <> 0 Or observations(simIndex, 2) <> 0 Then nAvg(2) = nAvg(2) + result.nHitpointRecovery - observations(simIndex, 2): nAvgCount(2) = nAvgCount(2) + 1
        If result.nManaRecovery <> 0 Or observations(simIndex, 3) <> 0 Then nAvg(3) = nAvg(3) + result.nManaRecovery - observations(simIndex, 3): nAvgCount(3) = nAvgCount(3) + 1
                If result.nMove <> 0 Or observations(simIndex, 4) <> 0 Then nAvg(4) = nAvg(4) + result.nMove - observations(simIndex, 4): nAvgCount(4) = nAvgCount(4) + 1
    DebugLogPrint "-----------------------------------------------"
    
    ' --- Simulation 16 ---
    simIndex = simIndex + 1
    result = CalcExpPerHour(3171, 3, 1, 35, 100, 1, 448, 1175, 81, 1, 238, 0, 0, 0, 0, 0, 0, 0, 2.9, 0)
    summaries(simIndex) = "SIM" & simIndex & ": 75 Warrior/stone elementals/physical/no heal: " & FormatResult(result, observations, simIndex)
    DebugLogPrint summaries(simIndex)
    If result.nExpPerHour > 0 Then nAvg(1) = nAvg(1) + (1 - (observations(simIndex, 1) / result.nExpPerHour)) Else nAvg(1) = nAvg(1) + 1
    If result.nHitpointRecovery <> 0 Or observations(simIndex, 2) <> 0 Then nAvg(2) = nAvg(2) + result.nHitpointRecovery - observations(simIndex, 2): nAvgCount(2) = nAvgCount(2) + 1
        If result.nManaRecovery <> 0 Or observations(simIndex, 3) <> 0 Then nAvg(3) = nAvg(3) + result.nManaRecovery - observations(simIndex, 3): nAvgCount(3) = nAvgCount(3) + 1
                If result.nMove <> 0 Or observations(simIndex, 4) <> 0 Then nAvg(4) = nAvg(4) + result.nMove - observations(simIndex, 4): nAvgCount(4) = nAvgCount(4) + 1
    DebugLogPrint "-----------------------------------------------"
    
    ' --- Simulation 17 ---
    simIndex = simIndex + 1
    result = CalcExpPerHour(49500, 5, 3, 25, 50, 1.5, 447, 1175, 81, 104, 1620, 0, 0, 0, 0, 0, 0, 0, 2, 0)
    summaries(simIndex) = "SIM" & simIndex & ": 75 Warrior/stone giants/physical/no heal: " & FormatResult(result, observations, simIndex)
    DebugLogPrint summaries(simIndex)
    If result.nExpPerHour > 0 Then nAvg(1) = nAvg(1) + (1 - (observations(simIndex, 1) / result.nExpPerHour)) Else nAvg(1) = nAvg(1) + 1
    If result.nHitpointRecovery <> 0 Or observations(simIndex, 2) <> 0 Then nAvg(2) = nAvg(2) + result.nHitpointRecovery - observations(simIndex, 2): nAvgCount(2) = nAvgCount(2) + 1
        If result.nManaRecovery <> 0 Or observations(simIndex, 3) <> 0 Then nAvg(3) = nAvg(3) + result.nManaRecovery - observations(simIndex, 3): nAvgCount(3) = nAvgCount(3) + 1
                If result.nMove <> 0 Or observations(simIndex, 4) <> 0 Then nAvg(4) = nAvg(4) + result.nMove - observations(simIndex, 4): nAvgCount(4) = nAvgCount(4) + 1
    DebugLogPrint "-----------------------------------------------"
    
    DebugLogPrint vbCrLf & String(100, "=")
    DebugLogPrint "Simulation Summary:"
    For i = LBound(summaries) To UBound(summaries)
        nTest = InStr(1, summaries(i), vbTab)
        If nTest > nMaxDesc Then nMaxDesc = nTest
    Next
    For i = LBound(summaries) To UBound(summaries)
        nTest = InStr(1, summaries(i), vbTab)
        nTest = nMaxDesc - nTest
        DebugLogPrint Replace(summaries(i), vbTab & "Exp/hr:", String(nTest, " ") & vbTab & "Exp/hr:")
    Next i
    If nAvgCount(2) = 0 Then nAvgCount(2) = 1
    If nAvgCount(3) = 0 Then nAvgCount(3) = 1
    If nAvgCount(4) = 0 Then nAvgCount(4) = 1
    DebugLogPrint "Avg Exp Diff: " & IIf(nAvg(1) > 0, "+", "") & Format((nAvg(1) / UBound(summaries)) * 100, "0.0") & "%"
    DebugLogPrint "Avg Rest Diff: " & IIf(nAvg(2) > 0, "+", "") & Format((nAvg(2) / nAvgCount(2)) * 100, "0.0") & "%"
    DebugLogPrint "Avg Mana Diff: " & IIf(nAvg(3) > 0, "+", "") & Format((nAvg(3) / nAvgCount(3)) * 100, "0.0") & "%"
    DebugLogPrint "Avg Move Diff: " & IIf(nAvg(4) > 0, "+", "") & Format((nAvg(4) / nAvgCount(4)) * 100, "0.0") & "%"
    DebugLogPrint String(100, "=")
End Sub

Private Function FormatResult(ByRef r As tExpPerHourInfo, ByRef o() As Double, Index As Integer, Optional ByVal nFormat As Integer) As String
    Dim diff(1 To 4) As Double
    If r.nExpPerHour > 0 Then
        diff(1) = 1 - (o(Index, 1) / r.nExpPerHour)
    Else
        diff(1) = 1
    End If
    diff(2) = r.nHitpointRecovery - o(Index, 2)
    diff(3) = r.nManaRecovery - o(Index, 3)
    diff(4) = r.nMove - o(Index, 4)
    
    If nFormat = 0 Then
        FormatResult = _
            vbTab & "Exp/hr: " & Format(r.nExpPerHour / 1000, "0") & "k est VS " & Format(o(Index, 1) / 1000, "0") & IIf(diff(1) > 0, "k obs (+", "k obs (") & Format(diff(1) * 100, "0.0") & "%)" & _
            vbTab & " Rest%: " & Format(r.nHitpointRecovery * 100, "0.0") & "% est VS " & Format(o(Index, 2) * 100, "0.0") & IIf(diff(2) > 0, "% obs (+", "% obs (") & Format(diff(2) * 100, "0.0") & "%)" & _
            vbTab & " Mana%: " & Format(r.nManaRecovery * 100, "0.0") & "% est VS " & Format(o(Index, 3) * 100, "0.0") & IIf(diff(3) > 0, "% obs (+", "% obs (") & Format(diff(3) * 100, "0.0") & "%)" & _
            vbTab & " Move: " & Format(r.nMove * 100, "0.0") & "% est VS " & Format(o(Index, 4) * 100, "0.0") & IIf(diff(4) > 0, "% obs (+", "% obs (") & Format(diff(4) * 100, "0.0") & "%)"
    Else
        FormatResult = _
            vbTab & "Exp/hr: " & Format(r.nExpPerHour / 1000, "0.0") & "k" & _
            vbTab & " Overkill: " & Format(r.nOverkill * 100, "0.00") & "%" & _
            vbTab & " Rest%: " & Format(r.nHitpointRecovery * 100, "0.00") & "%" & _
            vbTab & " Mana%: " & Format(r.nManaRecovery * 100, "0.00") & "%" & _
            vbTab & " Move: " & Format(r.nMove * 100, "0.00") & "%"
    End If
End Function

'==============================================================================
'  Exp/Hour – Model A (ceph_ModelA) – Overview & Calibration Notes
'  Version: v4.1   Date: 2025-08-09
'------------------------------------------------------------------------------
'  PURPOSE
'    Estimate effective EXP/hour (EPH) for lair-style zones by modeling:
'      • Attack time (rounds-to-kill across all mobs in a pull/“room”)
'      • Recovery time (HP + Mana), with movement overlap credits
'      • Movement time between lairs from spawn density + patrol routing
'      • Spawn-gating waits when the loop outruns respawns
'
'  TOP-LEVEL FLOW (per room)
'    1) Setup & validation
'       - Derive nRTC (room rounds) and killTimeSec = nRTC * SEC_PER_ROUND.
'       - Overkill heuristic (overshootFrac) for wasted damage on the last hit.
'
'    2) HP recovery demand (cephA_CalcHPRecoveryRounds v4.0)
'       - During combat, passive 30s ticks apply; after combat, resting heals
'         combine REST (every 20s) + passive (every 30s) into a per-second rate.
'       - Convert net damage ? rest rounds; apply q-elasticity:
'           q = incoming dmg/round ÷ rest heal/round while resting
'           • Boost when q = 0.9 (cap ~1.9×).  • Damp when q > 0.9 (floor 0.4).
'       - Scale by nLocalDmgScaleFactor and nGlobal_cephA_DMG (currently 1).
'
'    3) Mana recovery demand (pool model)
'       - costRoom = (nSpellCost * nRTK * nNumMobs) + (nSpellOverhead * nRTC)
'       - regenRoom = in-combat passive MP regen; drainRoom = cost - regen.
'       - roomsPerPool = nCharMana / drainRoom; tRestAvg ˜ refill time / roomsPerPool.
'       - nManaRecoveryTimeSec = tRestAvg * nGlobal_cephA_Mana (currently 1).
'
'    4) Combine HP + Mana demand (no overlap yet)
'       - recoveryDemandFrac = HP + MP - (HP × MP)
'       - recoveryDemandTime converts that fraction into seconds vs. kill time.
'
'    5) Movement model (density- and route-aware)
'       Inputs: nTotalLairs, nPossSpawns, nRegenTime, nAvgWalk, encumbrance.
'       - roomsRaw = nPossSpawns + nTotalLairs
'       - Scale for respawn: compute effectiveLairs & roomsScaled (bounded by loop
'         throughput). densityP = effectiveLairs / roomsScaled. pTravel = nTotalLairs / roomsRaw.
'       - Density-referenced “effective secs/room”:
'           scaleFactor = 1 + ((1/densityP - 1) / (1/nRoomDensityCoef - 1)) * (targetFactor-1)
'           where targetFactor = 2.5 / nSecsPerRoom   (nRoomDensityCoef = 0.25)
'           SecsPerRoomEff = nSecsPerRoom * scaleFactor
'       - Two movement estimates (use the larger; you can’t move less than the loop):
'           • Spawn-based:   moveSpawnBased = ((1-densityP)/densityP) * SecsPerRoomEff * nMoveBias
'           • Route-based:   moveRouteBased = ((roomsRaw/nTotalLairs)-1) * nSecsPerRoom * nRouteBiasLocal * nGlobal_cephA_Move
'             (guards: minimum when very dense; slight uplift when 0.20=pTravel=0.30)
'         Defaults: nGlobalMoveBias=0.85, nGlobalRoomRouteBias=1.00.
'
'    6) Walk-credit ? Mana only (caster-friendly, density-damped)
'       - regenWalkRaw = mpPerSec_regen * moveBaseSec.
'       - For sparse maps (pTravel < 0.25) damp to 40–90% (walkScale), then cap:
'           micro-rooms cap at 70% of killTime MP regen; otherwise cap at 40%.
'       - regenRoom += regenWalk; recompute drain ? roomsPerPool ? nManaRecoveryTimeSec.
'
'    7) Overlap credits (movement overlapping recovery)
'       - Constant overlap: recoveryCreditSec = nGlobal_cephA_MoveRecover * recoveryDemandFrac * moveBaseSec
'       - Split credit HP/MP in proportion to their remaining times; optionally
'         convert any residual HP-rest into mana-equivalent if meditating.
'       - Recompute HP/MP fractions from adjusted times (no micro-rest floor here).
'
'    8) Spawn gating
'       - If timePerClear < spawnInterval, wait fillerSec to the gate; split that
'         wait into movement vs. standing via pTravel (0.20–1.00). Standing time
'         grants extra recovery credit; only the movement share adds to move time.
'       - Record roam fraction (nRoamTime) as fillerSec/spawnInterval.
'
'    9) Final assembly
'       - timePerClear = kill + remaining recovery + movement (+ gating)
'       - effClearsPerHour = 3600 / timePerClear; EPH = nExp * effClearsPerHour
'       - Output fractions + diagnostics: Attack, Move, HP, MP, Recover, Slowdown,
'         Overkill, and RoamTime.
'
'  GLOBAL KNOBS
'    nGlobal_cephA_DMG            = 1.00
'    nGlobal_cephA_Mana           = 1.00
'    nGlobal_cephA_MoveRecover     = 0.85
'    nRoomDensityCoef            = 0.25
'    nGlobal_cephA_Move             = 1.00
'
'  TICKS / CONSTANTS USED
'    SEC_PER_ROUND        (combat round, 5s typical)
'    SEC_PER_REGEN_TICK   (passive tick, 30s)
'    SEC_PER_REST_TICK    (rest tick, 20s)
'    SEC_PER_MEDI_TICK    (meditation tick, 10s)
'    SECS_ROOM_BASE / SECS_ROOM_HEAVY (encumbrance-aware travel pacing)
'
'  CALIBRATION INTENT
'    • Favor slight optimism on EPH rather than under-estimating.
'    • “White Dragons”: raise EPH mainly by trimming movement without lowering
'      HP rest%. The new movement model + strong overlap credit achieved this.
'    • “Stone Elementals”: micro-rooms stay stable (little walk credit; minimal rest).
'
'  SAFE TUNERS FOR FUTURE PASSES
'    • Movement optimism: nMoveBias, nGlobal_cephA_Move,
'      density ref (nRoomDensityCoef), and the pTravel windows.
'    • Overlap feel: nGlobal_cephA_MoveRecover (currently 0.85).
'    • Caster feel: walkScale curve & walk caps in the Walk-Credit section.
'
'  DEPENDENCIES
'    IsMobKillable(), DebugLogPrint(), formatters F1/F2/F3/F6(), Pct(), HandleError()
'    cephA_CalcHPRecoveryRounds() – HP rounds/seconds + q-elasticity.
'==============================================================================

Private Function ceph_ModelA( _
    Optional ByVal nExp As Currency, Optional ByVal nRegenTime As Double, Optional ByVal nNumMobs As Integer, _
    Optional ByVal nTotalLairs As Long = -1, Optional ByVal nPossSpawns As Long, Optional ByVal nRTK As Double, _
    Optional ByVal nCharDMG As Double, Optional ByVal nCharHP As Long, Optional ByVal nCharHPRegen As Long, _
    Optional ByVal nMobDmg As Double, Optional ByVal nMobHP As Long, Optional ByVal nMobHPRegen As Long, _
    Optional ByVal nDamageThreshold As Long, Optional ByVal nSpellCost As Integer, _
    Optional ByVal nSpellOverhead As Double, Optional ByVal nCharMana As Long, _
    Optional ByVal nCharMPRegen As Long, Optional ByVal nMeditateRate As Long, _
    Optional ByVal nAvgWalk As Double, Optional ByVal nEncumPct As Integer) As tExpPerHourInfo

On Error GoTo error

'------------------------------------------------------------------
'  -- local variables -------------------
'------------------------------------------------------------------
Dim nSecsPerRoom         As Double
Dim nHitpointRecovery    As Double
Dim nHitpointRecoveryTimeSec    As Double
Dim nManaRecovery    As Double
Dim nManaRecoveryTimeSec    As Double
Dim roundsHitpoints      As Double
Dim nRTC                 As Double
Dim killTimeSec          As Double
Dim recoveryTimeSec      As Double
Dim moveTimeSec          As Double
Dim timeLoss             As Double
Dim timePerClearSec      As Double
Dim effClearsPerHour     As Double
Dim attackFrac           As Double
Dim recoverFrac          As Double
Dim hitpointFrac         As Double
Dim manaFrac             As Double
Dim moveFrac             As Double
Dim maxRooms             As Double
Dim effectiveLairs       As Double
Dim slowdownFrac         As Double
Dim totalDamage          As Double
Dim effectiveMobHP       As Double
Dim overshootFrac        As Double

Dim recoveryDemandFrac  As Double
Dim recoveryDemandTime  As Double
Dim recoveryCreditSec   As Double
Dim roomsRaw            As Double
Dim roomsScaled         As Double
Dim densityP            As Double
Dim targetFactor        As Double
Dim scaleFactor         As Double
Dim SecsPerRoomEff      As Double
Dim moveBaseSec         As Double
Dim fillerSec           As Double
Dim spawnInterval       As Double
Dim pTravel             As Double
Dim moveSpawnBased      As Double
Dim moveRouteBased      As Double
Dim nRouteBiasLocal     As Double
Dim nMoveBias     As Double
Dim bLimitMovement As Boolean
Dim nRoomDensityCoef As Double

'------------------------------------------------------------------
'  -- fast bail-outs ----------------------------------------------
'------------------------------------------------------------------
If nExp = 0 Then Exit Function
If nTotalLairs = 0 And nRegenTime = 0 Then Exit Function
If Not IsMobKillable(nCharDMG, nCharHP, nMobDmg, nMobHP, nCharHPRegen, nMobHPRegen) Then
    ceph_ModelA.nExpPerHour = -1
    ceph_ModelA.nHitpointRecovery = 1
    ceph_ModelA.nTimeRecovering = 1
    Exit Function
End If

'------------------------------------------------------------------
'  -- globals / tuners --------------------------------------------
'------------------------------------------------------------------
nMoveBias = 0.85
nRouteBiasLocal = 0.98
nRoomDensityCoef = 0.25
If nGlobal_cephA_Move > 0 And nGlobal_cephA_Move <> 1 Then nRoomDensityCoef = nRoomDensityCoef * nGlobal_cephA_Move

If nAvgWalk > 0 And nAvgWalk <= 2 And nTotalLairs > 0 And nPossSpawns > nTotalLairs Then
    'cluster detection (i.e. gnoll encampment)
    If nPossSpawns / nTotalLairs >= nGlobal_cephA_ClusterMx Then bLimitMovement = True
End If

If nEncumPct >= HEAVY_ENCUM_PCT Then       ' heavy
    nSecsPerRoom = SECS_ROOM_HEAVY          ' 100 rooms / 180 s
Else
    nSecsPerRoom = SECS_ROOM_BASE          ' 100 rooms / 120 s
End If

'------------------------------------------------------------------
'  -- DEBUG: raw inputs -------------------------------------------
'------------------------------------------------------------------
#If DEVELOPMENT_MODE Then
    If bDebugExpPerHour Then
        DebugLogPrint "DBG_IN ------------- ceph_ModelA -------------"
        DebugLogPrint "  nExp=" & nExp & "; nRegenTime=" & nRegenTime & "; nNumMobs=" & nNumMobs & _
                    "; nTotalLairs=" & nTotalLairs & "; nPossSpawns=" & nPossSpawns & "; nRTK=" & nRTK
        DebugLogPrint "  nCharDMG=" & nCharDMG & "; nCharHP=" & nCharHP & "; nCharHPRegen=" & nCharHPRegen & _
                    "; nMobDmg=" & nMobDmg & "; nMobHP=" & nMobHP & "; nMobHPRegen=" & nMobHPRegen
        DebugLogPrint "  nDamageThreshold=" & nDamageThreshold & "; nSpellCost=" & nSpellCost & _
                    "; nSpellOverhead=" & nSpellOverhead & "; nCharMana=" & nCharMana
        DebugLogPrint "  nCharMPRegen=" & nCharMPRegen & "; nMeditateRate=" & nMeditateRate & _
                    "; nAvgWalk=" & nAvgWalk & "; nEncumPct=" & nEncumPct
    End If
#End If


'------------------------------------------------------------------
'  -- validation ---------------------------------------------------
'------------------------------------------------------------------
If nCharHP < 1 Then nCharHP = 1
If nMobHP < 1 Then nMobHP = 1
If nCharHPRegen < 1 Then nCharHPRegen = 1
If nRegenTime < 0 Then nRegenTime = 0
If nRegenTime > 60 Then nRegenTime = 60

If nCharDMG > 0 And nCharDMG < nMobHP And nRTK = 0 Then
    nRTK = nMobHP / nCharDMG
    If nRTK > 1 Then nRTK = -Int(-(nRTK * 2)) / 2 'round up to the nearest 0.5
End If
If nRTK < 1 Then nRTK = 1

If nNumMobs < 1 Then nNumMobs = 1
nRTC = nRTK * nNumMobs

'------------------------------------------------------------------
'  -- NPC / boss shortcut -----------------------------------------
'------------------------------------------------------------------
If nTotalLairs <= 0 And nRegenTime > 0 Then
    effClearsPerHour = 1 / nRegenTime
    ceph_ModelA.nExpPerHour = Round(nExp * effClearsPerHour)
    Exit Function
End If

If nRegenTime > 0 Then nRegenTime = nRegenTime + 0.25 'in reality, because lair regen happens by time of day, it's actually "regen time + however many seconds left in the minute when the last mob was killed"

'------------------------------------------------------------------
'  -- attack time & over-damage -----------------------------------
'------------------------------------------------------------------
killTimeSec = nRTC * SEC_PER_ROUND
If nCharDMG > 0 Then
    totalDamage = nRTK * nCharDMG
    If nNumMobs > 1 Then
        effectiveMobHP = nMobHP / nNumMobs
    Else
        effectiveMobHP = nMobHP
    End If
    If totalDamage > effectiveMobHP Then
        overshootFrac = ((totalDamage - effectiveMobHP) / totalDamage) * 0.8
    ElseIf effectiveMobHP > 0 And (nRTK <> Fix(nRTK)) Then
        overshootFrac = (Abs(nRTK - Fix(nRTK)) * nCharDMG) / effectiveMobHP
    End If
    If nNumMobs > 1 Then totalDamage = totalDamage * nNumMobs
End If

'------------------------------------------------------------------
'  -- HP recovery --------------------------------------------------
'------------------------------------------------------------------
Dim qRatio As Double           ' ratio of HP healed / dmg taken in one round
Dim nLocalDmgScaleFactor As Double

qRatio = 1#
If nMobDmg > (nDamageThreshold / 2#) Then
    qRatio = nCharHPRegen / (nMobDmg - (nDamageThreshold / 2#))
    If qRatio < 0# Then qRatio = 0#
End If

Select Case qRatio
    Case Is < 0.3
        nLocalDmgScaleFactor = 1.2    ' still need a bit of slack
    Case 0.3 To 0.7
        ' slides 1.20 <> 1.05  over q = 0.30?0.70
        nLocalDmgScaleFactor = 1.05 + (0.375 * (0.7 - qRatio))
    Case Else
        nLocalDmgScaleFactor = 1#     ' near one-to-one regen/dmg
End Select

If nDamageThreshold > 0 And nDamageThreshold < nMobDmg Then
    roundsHitpoints = cephA_CalcHPRecoveryRounds(nMobDmg - nDamageThreshold, nCharDMG, nMobHP, nCharHPRegen, nNumMobs, nRTC)
ElseIf nDamageThreshold = 0 And nMobDmg > 0 Then
    roundsHitpoints = cephA_CalcHPRecoveryRounds(nMobDmg - (nCharHPRegen / 18), nCharDMG, nMobHP, nCharHPRegen, nNumMobs, nRTC)
End If
If roundsHitpoints < 0 Then roundsHitpoints = 0

' Direct time from rounds (apply scale on the physical quantity)
Dim R_HP_adj As Double
R_HP_adj = roundsHitpoints * nLocalDmgScaleFactor * nGlobal_cephA_DMG
nHitpointRecoveryTimeSec = R_HP_adj * SEC_PER_ROUND
If nHitpointRecoveryTimeSec < 0# Then nHitpointRecoveryTimeSec = 0#

' Fraction for demand math / UI is derived from time (no special-case)
If (killTimeSec + nHitpointRecoveryTimeSec) > 0# Then
    nHitpointRecovery = nHitpointRecoveryTimeSec / (killTimeSec + nHitpointRecoveryTimeSec)
Else
    nHitpointRecovery = 0#
End If
If nHitpointRecovery > 1# Then nHitpointRecovery = 1#
If nHitpointRecovery < 0# Then nHitpointRecovery = 0#

#If DEVELOPMENT_MODE Then
    If bDebugExpPerHour Then
        DebugLogPrint "HPDBG --- After HP rounds?time (pre-overlap) ---"
        DebugLogPrint "  roundsHitpoints=" & F6(roundsHitpoints) & _
                    "; qRatio=" & F3(qRatio) & _
                    "; nLocalDmgScaleFactor=" & F3(nLocalDmgScaleFactor) & _
                    "; nGlobalDmgScaleFactor=" & F3(nGlobal_cephA_DMG) & _
                    "; nHitpointRecoveryTimeSec=" & F1(nHitpointRecoveryTimeSec) & "s" & _
                    "; killTimeSec=" & F1(killTimeSec) & "s" & _
                    "; HPfrac=" & Pct(nHitpointRecovery)
    End If
#End If

'------------------------------------------------------------------
'  -- Mana recovery (per-room pool model) -------------------------
'
' Terminology
'   ? Room  = one complete pull (all mobs present)
'   ? nRTK     rounds-to-kill a single mob
'   ? nRTC     rounds in the whole room  (= nRTK ?nNumMobs)
'   ? killTimeSec = nRTC ?5 s
'------------------------------------------------------------------
nManaRecovery = 0#
nManaRecoveryTimeSec = 0#

' 1)  Per-second regeneration rates
Dim mpPerSec_regen    As Double
Dim mpPerSec_meditate As Double
mpPerSec_regen = nCharMPRegen / SEC_PER_REGEN_TICK                            ' 30-s tick spread
mpPerSec_meditate = mpPerSec_regen + (nMeditateRate / SEC_PER_MEDI_TICK)     ' add 10-s ticks

' 2)  Mana **spent per room**
Dim costRoom  As Double
costRoom = (nSpellCost * nRTK * nNumMobs) + (nSpellOverhead * nRTC)

' 3)  Mana regenerated *during* that room
Dim regenRoom As Double
regenRoom = mpPerSec_regen * killTimeSec

' 4)  Net drain that must be refilled by meditating
Dim drainRoom As Double
drainRoom = costRoom - regenRoom
If drainRoom < 0# Then drainRoom = 0#

' 5)  How many rooms can we clear before the pool is empty?
Dim roomsPerPool As Double
If drainRoom = 0# Or nCharMana = 0 Then
    roomsPerPool = 1E+30          ' effectively infinite
Else
    roomsPerPool = nCharMana / drainRoom
End If

' 6)  Time (s) to refill 95 % of the pool
Dim refillTarget As Double, tRefill As Double, tRestAvg As Double
refillTarget = 0.95 * nCharMana
If mpPerSec_meditate > 0# Then
    tRefill = refillTarget / mpPerSec_meditate
Else
    tRefill = 0#
End If

' 7)  Average rest time **per room**
tRestAvg = tRefill / roomsPerPool

' 8)  Apply global optimism/pessimism knob
nManaRecoveryTimeSec = tRestAvg * nGlobal_cephA_Mana
If nManaRecoveryTimeSec < 0# Then nManaRecoveryTimeSec = 0#

' 9)  Convert to fractional demand
If (killTimeSec + nManaRecoveryTimeSec) > 0# Then
    nManaRecovery = nManaRecoveryTimeSec / (killTimeSec + nManaRecoveryTimeSec)
Else
    nManaRecovery = 0#
End If
If nManaRecovery > 1# Then nManaRecovery = 1#

#If DEVELOPMENT_MODE Then
    If bDebugExpPerHour Then
        DebugLogPrint "MPDBG --- pool model ---"
        DebugLogPrint "  costRoom=" & F6(costRoom) & "; regenRoom=" & F6(regenRoom) _
                    & "; drainRoom=" & F6(drainRoom)
        DebugLogPrint "  roomsPerPool=" & F6(roomsPerPool) _
                    & "; refillTarget=" & F6(refillTarget)
        DebugLogPrint "  tRefill=" & F6(tRefill) & "; tRestAvg=" & F6(tRestAvg)
        DebugLogPrint "  => nManaRecoveryTimeSec=" & F6(nManaRecoveryTimeSec)
    End If
#End If

If nManaRecovery = 1# Then
    nManaRecoveryTimeSec = killTimeSec * 2#
ElseIf nManaRecovery > 0# And nManaRecoveryTimeSec = 0# Then
    nManaRecoveryTimeSec = killTimeSec * (nManaRecovery / (1# - nManaRecovery))
End If
If nManaRecoveryTimeSec < 0# Then nManaRecoveryTimeSec = 0#

#If DEVELOPMENT_MODE Then
    If bDebugExpPerHour Then
        DebugLogPrint "MPDBG --- Mana demand calc ---"
        DebugLogPrint "  nManaRecoveryFrac=" & Pct(nManaRecovery)
        DebugLogPrint "  nManaRecoveryTimeSec=" & F1(nManaRecoveryTimeSec) & "s"
    End If
#End If

'------------------------------------------------------------------
'  -- Combine HP & Mana demand (no movement overlap yet) ----------
'------------------------------------------------------------------
recoveryDemandFrac = nHitpointRecovery + nManaRecovery - (nHitpointRecovery * nManaRecovery)
If recoveryDemandFrac < 0 Then recoveryDemandFrac = 0
If recoveryDemandFrac > 1 Then recoveryDemandFrac = 1

' Rest time required if NO movement overlapped any recovery:
If recoveryDemandFrac > 0 And recoveryDemandFrac < 1 Then
    recoveryDemandTime = killTimeSec * (recoveryDemandFrac / (1# - recoveryDemandFrac))
ElseIf recoveryDemandFrac >= 1 Then
    recoveryDemandTime = 3600# ' cap; pathological, will be clipped by spawn later
Else
    recoveryDemandTime = 0
End If

#If DEVELOPMENT_MODE Then
    If bDebugExpPerHour Then
        DebugLogPrint "HPDBG --- Demand ---"
        DebugLogPrint "  nManaRecoveryTimeSec(pre)=" & F1(nManaRecoveryTimeSec) & "s"
        DebugLogPrint "  recoveryDemandFrac=" & Pct(recoveryDemandFrac) & _
                    "; recoveryDemandTime=" & F1(recoveryDemandTime) & "s"
    End If
#End If

'------------------------------------------------------------------
'  -- Movement model ----------------------------------------------
'------------------------------------------------------------------
If nTotalLairs > 0 And bLimitMovement = False Then

    roomsRaw = nPossSpawns + nTotalLairs
    If nRegenTime > 0 Then
        effectiveLairs = (60# * nRegenTime) / 5#
        If effectiveLairs > nTotalLairs Then
            roomsScaled = roomsRaw * (nTotalLairs / effectiveLairs)
            effectiveLairs = nTotalLairs
        Else
            roomsScaled = roomsRaw
        End If
        maxRooms = effectiveLairs * (60# / nRegenTime)
    Else
        effectiveLairs = nTotalLairs
        roomsScaled = roomsRaw
        maxRooms = 720# '3600/5 = 1 lair every 5 seconds for an hour
    End If

    If roomsScaled > maxRooms Then roomsScaled = maxRooms

    ' Density of lairs among walkable rooms after scaling
    If roomsScaled <= 0 Then
        densityP = 1
    Else
        densityP = effectiveLairs / roomsScaled
        If densityP < 0.01 Then densityP = 0.01
        If densityP > 1 Then densityP = 1
    End If

    pTravel = nTotalLairs / roomsRaw                  ' map density (true patrol)
    If pTravel < 0.0001 Then pTravel = 0.0001
    If pTravel > 1 Then pTravel = 1

    If pTravel < 0.1 Then
        nRouteBiasLocal = 0.7 + (3 * pTravel)        ' = 1.0 at p=0.10, 0.70 at p -> 0
    ElseIf pTravel < 0.18 Then
        If densityP > 0.5 Then
            nRouteBiasLocal = 1.08
        ElseIf densityP >= 0.25 And densityP <= 0.4 Then
            nRouteBiasLocal = 0.85
        Else
            nRouteBiasLocal = 1.02
        End If
    End If

    ' Density-aware effective seconds per room.
    targetFactor = 2.5 / nSecsPerRoom
    If densityP <> 0 And nRoomDensityCoef <> 0 And nRoomDensityCoef <> 1 Then
        scaleFactor = 1 + ((1# / densityP - 1) / (1 / nRoomDensityCoef - 1)) * (targetFactor - 1)
        If scaleFactor < 1 Then scaleFactor = 1
    ElseIf nRoomDensityCoef = 0 Then
        scaleFactor = 0.00001
    Else
        scaleFactor = targetFactor
    End If
    SecsPerRoomEff = nSecsPerRoom * scaleFactor

    ' 1) Spawn-based
    If densityP > 0 Then moveSpawnBased = ((1# - densityP) / densityP) * SecsPerRoomEff * nMoveBias

    ' 2) Route-based: rooms per lair on the loop, minus the lair room itself
    moveRouteBased = ((roomsRaw / nTotalLairs) - 1#) * nSecsPerRoom * nRouteBiasLocal * nGlobal_cephA_Move
    If densityP > 0.8 And moveRouteBased < 2 * nSecsPerRoom Then
        moveRouteBased = 2 * nSecsPerRoom
    End If
    If pTravel >= 0.2 And pTravel <= 0.3 Then
        moveRouteBased = moveRouteBased * (1 + 0.25 * (pTravel - 0.2) / 0.1)
    End If
    If moveRouteBased < 0 Then moveRouteBased = 0
    
    ' Use the larger ? you can?t realistically move less than the loop implies
    If moveRouteBased > moveSpawnBased Then
        moveBaseSec = moveRouteBased
    Else
        moveBaseSec = moveSpawnBased
    End If
ElseIf bLimitMovement Then
    moveBaseSec = nSecsPerRoom * nAvgWalk
Else
    'minimal step to next opportunity
    moveBaseSec = nSecsPerRoom
End If

#If DEVELOPMENT_MODE Then
    If bDebugExpPerHour Then
        DebugLogPrint "HPDBG --- Movement model ---"
        DebugLogPrint "  roomsRaw=" & roomsRaw & "; roomsScaled=" & F3(roomsScaled) & _
                    "; effectiveLairs=" & F3(effectiveLairs)
        DebugLogPrint "  densityP=" & F6(densityP) & " (" & Pct(densityP) & "); pTravel=" & F6(pTravel) & " (" & Pct(pTravel) & ")"
        DebugLogPrint "  scaleFactor=" & F6(scaleFactor) & "; SecsPerRoomEff=" & F3(SecsPerRoomEff)
        DebugLogPrint "  moveSpawnBased=" & F3(moveSpawnBased) & "; moveRouteBased=" & F3(moveRouteBased) & _
                    "; moveBaseSec=" & F3(moveBaseSec)
    End If
#End If

'------------------------------------------------------------------
'  -- Walk-credit rolled back into mana model ---------------------
'------------------------------------------------------------------
Dim regenWalkRaw  As Double   ' value before density scaling
Dim regenWalk     As Double   ' final, possibly damped, credit
Dim walkScale     As Double
Dim walkCap       As Double
regenWalkRaw = mpPerSec_regen * moveBaseSec   ' mana-tick during travel
walkScale = 1#

If bLimitMovement Then
    regenWalk = regenWalkRaw * 0.25
Else
    ' Density dampener: 0.40 <> 0.90 for pTravel = 0 <> 0.25
    If pTravel < 0.25 Then
        walkScale = 0.4 + (2 * pTravel)
        If walkScale > 0.9 Then walkScale = 0.9 ' upper clamp at 90 % of the raw value
    End If
    regenWalk = regenWalkRaw * walkScale          ' final credit
End If

If killTimeSec <= 6# Then
    walkCap = 0.7 * killTimeSec * mpPerSec_regen    'cap for micro-rooms
Else
    walkCap = 0.4 * killTimeSec * mpPerSec_regen
End If
If regenWalk > walkCap Then regenWalk = walkCap

regenRoom = regenRoom + regenWalk             ' adjust in-room regen
drainRoom = costRoom - regenRoom              ' refresh downstream value
If drainRoom < 0# Then drainRoom = 0#

If drainRoom = 0# Then
    roomsPerPool = 1E+30
Else
    roomsPerPool = nCharMana / drainRoom
    If roomsPerPool < 1# Then roomsPerPool = 1#
End If

tRestAvg = tRefill / roomsPerPool
nManaRecoveryTimeSec = tRestAvg * nGlobal_cephA_Mana
If nManaRecoveryTimeSec < 0# Then nManaRecoveryTimeSec = 0#

'   Update the fraction as well:
If (killTimeSec + nManaRecoveryTimeSec) > 0# Then
    nManaRecovery = nManaRecoveryTimeSec / (killTimeSec + nManaRecoveryTimeSec)
Else
    nManaRecovery = 0#
End If
If nManaRecovery > 1# Then nManaRecovery = 1#

#If DEVELOPMENT_MODE Then
    If bDebugExpPerHour Then
        DebugLogPrint "MPDBG --- Walk-Credit ---"
        DebugLogPrint "  regenWalk=" & F2(regenWalk)
        DebugLogPrint "  regenRoom=" & F2(regenRoom)
        DebugLogPrint "  drainRoom=" & F2(drainRoom)
        DebugLogPrint "  roomsPerPool=" & F2(roomsPerPool)
        DebugLogPrint "  tRestAvg=" & F2(tRestAvg)
        DebugLogPrint "  nManaRecoveryTimeSec=" & F2(nManaRecoveryTimeSec)
        DebugLogPrint "  nManaRecovery=" & F2(nManaRecovery)
    End If
#End If

'------------------------------------------------------------------
'  -- Overlap credits ---------------------------------------------
'------------------------------------------------------------------
Dim T_HP0 As Double, T_M0 As Double
Dim moveCredHP As Double, moveCredMP As Double
Dim T_HP1 As Double, T_M1 As Double, T_M2 As Double
Dim restManaCredit As Double, restAsManaEq As Double

' Baselines (before overlap credits)
T_HP0 = nHitpointRecoveryTimeSec
T_M0 = nManaRecoveryTimeSec

#If DEVELOPMENT_MODE Then
    If bDebugExpPerHour Then
        DebugLogPrint "HPDBG --- Overlap (start) ---"
        DebugLogPrint "  T_HP0=" & F1(nHitpointRecoveryTimeSec) & "s; T_M0=" & F1(nManaRecoveryTimeSec) & "s" & _
                    "; moveBaseSec=" & F1(moveBaseSec) & "s"
    End If
#End If

' 1) Movement overlap: split proportionally between HP and Mana demand
recoveryCreditSec = nGlobal_cephA_MoveRecover * recoveryDemandFrac * moveBaseSec
If recoveryCreditSec > (T_HP0 + T_M0) Then recoveryCreditSec = (T_HP0 + T_M0)

If (T_HP0 + T_M0) > 0# Then
    moveCredHP = recoveryCreditSec * (T_HP0 / (T_HP0 + T_M0))
Else
    moveCredHP = 0#
End If
moveCredMP = recoveryCreditSec - moveCredHP

T_HP1 = T_HP0 - moveCredHP: If T_HP1 < 0# Then T_HP1 = 0#
T_M1 = T_M0 - moveCredMP: If T_M1 < 0# Then T_M1 = 0#

If mpPerSec_meditate > 0# Then
    restAsManaEq = T_HP1 * (mpPerSec_regen / mpPerSec_meditate)
Else
    restAsManaEq = 0#
End If

#If DEVELOPMENT_MODE Then
    If bDebugExpPerHour Then
        DebugLogPrint "HPDBG --- Overlap (after move) ---"
        DebugLogPrint "  recoveryCreditSec=" & F1(recoveryCreditSec) & "s" & _
                    "; moveCredHP=" & F1(moveCredHP) & "s; moveCredMP=" & F1(moveCredMP) & "s"
        DebugLogPrint "  T_HP1=" & F1(T_HP1) & "s; T_M1=" & F1(T_M1) & "s"
        Dim pureMediTime As Double
        If mpPerSec_meditate > 0# Then pureMediTime = (T_HP1 * mpPerSec_regen) / mpPerSec_meditate
        DebugLogPrint "MPDBG --- convert HP-rest -> MP-meditate eq"
        DebugLogPrint "  T_HP1=" & F1(T_HP1) & "s; restAsManaEq=" & F1(restAsManaEq) & _
                    "s; pureMediTime=" & F1(pureMediTime) & "s"
    End If
#End If

restManaCredit = restAsManaEq
If restManaCredit > T_M1 Then restManaCredit = T_M1

T_M2 = T_M1 - restManaCredit
If T_M2 < 0# Then T_M2 = 0#

#If DEVELOPMENT_MODE Then
    If bDebugExpPerHour Then
        DebugLogPrint "HPDBG --- Overlap (rest?mana) ---"
        DebugLogPrint "  mpPerSec_regen=" & F6(mpPerSec_regen) & _
                    "; mpPerSec_meditate=" & F6(mpPerSec_meditate) & _
                    "; restAsManaEq=" & F1(restAsManaEq) & "s"
        DebugLogPrint "  restManaCredit=" & F1(restManaCredit) & "s; T_M2=" & F1(T_M2) & "s"
    End If
#End If

' Final recovery breakdown and total sequential time
nHitpointRecoveryTimeSec = T_HP1
nManaRecoveryTimeSec = T_M2
recoveryTimeSec = T_HP1 + T_M2
recoveryDemandTime = recoveryTimeSec

If nHitpointRecoveryTimeSec < 0# Then nHitpointRecoveryTimeSec = 0#
If nManaRecoveryTimeSec < 0# Then nManaRecoveryTimeSec = 0#
If recoveryTimeSec < 0 Then recoveryTimeSec = 0

'recalculate if nHitpointRecoveryTimeSec adjusted
If (killTimeSec + nHitpointRecoveryTimeSec) > 0# Then
    nHitpointRecovery = nHitpointRecoveryTimeSec / (killTimeSec + nHitpointRecoveryTimeSec)
Else
    nHitpointRecovery = 0#
End If
If nHitpointRecovery > 1# Then nHitpointRecovery = 1#

'recalculate if nManaRecoveryTimeSec adjusted
If (killTimeSec + nManaRecoveryTimeSec) > 0# Then
    nManaRecovery = nManaRecoveryTimeSec / (killTimeSec + nManaRecoveryTimeSec)
Else
    nManaRecovery = 0#
End If
If nManaRecovery > 1# Then nManaRecovery = 1#

'------------------------------------------------------------------
'  -- Totals pre-spawn gate ---------------------------------------
'------------------------------------------------------------------
moveTimeSec = moveBaseSec
timePerClearSec = killTimeSec + recoveryTimeSec + moveTimeSec

#If DEVELOPMENT_MODE Then
    If bDebugExpPerHour Then
        DebugLogPrint "HPDBG --- Totals (pre-spawn) ---"
        DebugLogPrint "  recoveryTimeSec=" & F1(recoveryTimeSec) & "s; moveTimeSec=" & F1(moveTimeSec) & "s; killTimeSec=" & F1(killTimeSec) & "s"
        DebugLogPrint "  timePerClear(pre)=" & F1(timePerClearSec) & "s"
    End If
#End If

'------------------------------------------------------------------
'  -- Spawn-gating / filler wait ----------------------------------
'------------------------------------------------------------------
spawnInterval = 0
If nTotalLairs > 0 And nRegenTime > 0 Then spawnInterval = (nRegenTime * 60#) / nTotalLairs

fillerSec = 0
If timePerClearSec > 0 And spawnInterval > timePerClearSec Then
    fillerSec = spawnInterval - timePerClearSec

     ' Portion of wait that is spent moving vs standing still, based on density:
    Dim fillerToMoveFrac As Double, fillerMove As Double, fillerStand As Double
    ' More sparse => more walking. Bias floor at 0.2 so we never assign 0.
    'fillerToMoveFrac = 0.2 + 0.8 * (1# - densityP)    ' densityP in [0,1]
    fillerToMoveFrac = 0.2 + 0.8 * (1# - pTravel)
    If fillerToMoveFrac < 0# Then fillerToMoveFrac = 0#
    If fillerToMoveFrac > 1# Then fillerToMoveFrac = 1#

    fillerMove = fillerSec * fillerToMoveFrac
    fillerStand = fillerSec - fillerMove

    ' Standing still also recovers HP/MP
    Dim extraRestCredit As Double
    extraRestCredit = fillerStand * recoveryDemandFrac
    ' Apply extra credit, bounded:
    recoveryCreditSec = recoveryCreditSec + extraRestCredit
    If recoveryCreditSec > recoveryDemandTime Then recoveryCreditSec = recoveryDemandTime
    recoveryTimeSec = recoveryDemandTime - recoveryCreditSec
    If recoveryTimeSec < 0 Then recoveryTimeSec = 0

    ' Add only the moving share to move time:
    moveTimeSec = moveBaseSec + fillerMove
    ' Final gated clear time:
    timePerClearSec = spawnInterval
    
    timeLoss = fillerSec / spawnInterval
    'timeLoss = Round((fillerSec / spawnInterval) * 100)
    'If timeLoss >= 1 Then sLairText = timeLoss & "% time lost due to insufficient lairs"

#If DEVELOPMENT_MODE Then
        If bDebugExpPerHour Then
            DebugLogPrint "HPDBG --- Spawn gating ---"
            DebugLogPrint "  spawnInterval=" & F1(spawnInterval) & "s; fillerSec=" & F1(fillerSec) & "s; fillerToMoveFrac=" & F3(fillerToMoveFrac)
            DebugLogPrint "  fillerMove=" & F1(fillerMove) & "s; fillerStand=" & F1(fillerStand) & "s"
            DebugLogPrint "  timePerClear(gated)=" & F1(timePerClearSec) & "s"
        End If
#End If
Else
    moveTimeSec = moveBaseSec
    timeLoss = 0
End If

'------------------------------------------------------------------
'  -- Final EPH & fractions ---------------------------------------
'------------------------------------------------------------------
If timePerClearSec > 0 Then effClearsPerHour = 3600# / timePerClearSec

' Fractions for output
If timePerClearSec > 0 Then
    attackFrac = killTimeSec / timePerClearSec
    recoverFrac = recoveryTimeSec / timePerClearSec
    hitpointFrac = nHitpointRecoveryTimeSec / timePerClearSec
    manaFrac = nManaRecoveryTimeSec / timePerClearSec
    moveFrac = moveTimeSec / timePerClearSec
    If nRTC > 1 Then slowdownFrac = (killTimeSec - (SEC_PER_ROUND * nNumMobs)) / timePerClearSec
Else
    attackFrac = 0: recoverFrac = 0: moveFrac = 0: slowdownFrac = 0
End If

ceph_ModelA.nExpPerHour = Round(nExp * effClearsPerHour)
ceph_ModelA.nHitpointRecovery = hitpointFrac
ceph_ModelA.nManaRecovery = manaFrac
ceph_ModelA.nTimeRecovering = recoverFrac
ceph_ModelA.nOverkill = overshootFrac
ceph_ModelA.nMove = moveFrac
ceph_ModelA.nAttackTime = attackFrac
ceph_ModelA.nSlowdownTime = slowdownFrac
ceph_ModelA.nRoamTime = timeLoss

If bLimitMovement Then ceph_ModelA.nMove = ceph_ModelA.nMove * -1

#If DEVELOPMENT_MODE Then
    If bDebugExpPerHour Then
        DebugLogPrint "HPDBG --- Fractions & EPH ---"
        DebugLogPrint "  attackFrac=" & Pct(attackFrac) & "; hitpointFrac=" & Pct(hitpointFrac) & _
                    "; manaFrac=" & Pct(manaFrac) & "; moveFrac=" & Pct(moveFrac) & "; recoverFrac=" & Pct(recoverFrac)
        DebugLogPrint "  effClearsPerHour=" & F3(effClearsPerHour) & "; ExpPerHour=" & ceph_ModelA.nExpPerHour
    End If
#End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("ceph_ModelA")
Resume out:
End Function

'======================================================================
'  cephA_CalcHPRecoveryRounds v4.0  (2025-08-04)
'======================================================================
Private Function cephA_CalcHPRecoveryRounds(ByVal nDmgIN As Double, ByVal nDmgOut As Double, _
    ByVal nMobHP As Double, ByVal nRestHP As Double, _
    Optional ByVal nMobs As Integer = 0, Optional ByVal nRTC As Double) As Double

On Error GoTo error:

If nDmgIN <= 0# Then Exit Function

Dim r As Double                    ' attack rounds
Dim combatSecs As Double
Dim dmgInTotal As Double
Dim passivePerTick As Double
Dim passiveHealCombat As Double
Dim netDmg As Double
Dim restHealPerSec As Double
Dim restRounds As Double
Dim restRounds_full As Double
Dim restTimeContinuous As Double

' Elasticity variables
Dim restHealPerRound As Double
Dim q As Double, g As Double

If nMobs < 1 Then nMobs = 1
If nRestHP < 1 Then nRestHP = 1

' Determine attack rounds if not supplied
If nRTC = 0# And nDmgOut > 0# Then
    If nDmgOut >= nMobHP Then
        r = 1#
    Else
        r = nMobHP / nDmgOut
    End If
    If nMobs > 1 Then r = r * nMobs
Else
    r = nRTC
End If
If r < 1# Then r = 1#

combatSecs = r * SEC_PER_ROUND

' Total incoming damage during combat
dmgInTotal = r * nDmgIN

' Passive ticks: (nRestHP/3) every 30s, regardless of state
passivePerTick = nRestHP / 3#
passiveHealCombat = (combatSecs / SEC_PER_REGEN_TICK) * passivePerTick

' Net damage that must be recovered after combat
netDmg = dmgInTotal - passiveHealCombat
If netDmg < 0# Then netDmg = 0#

' While RESTING: rest tick every 20s (nRestHP) + passive every 30s (nRestHP/3)
restHealPerSec = (nRestHP / SEC_PER_REST_TICK) + (passivePerTick / SEC_PER_REGEN_TICK)   ' exact average while resting

' Continuous rest seconds needed (no discretization)
If restHealPerSec > 0# Then
    restTimeContinuous = netDmg / restHealPerSec
Else
    restTimeContinuous = 0#
End If

' Rounds approximation used by caller (?s later)
If restHealPerSec > 0# Then
    restRounds = netDmg / (restHealPerSec * SEC_PER_ROUND)
Else
    restRounds = 0#
End If
If restRounds < 0# Then restRounds = 0#
restRounds_full = restRounds   ' keep pre-elastic value for diagnostics

' ---- Elastic correction based on damage-vs-rest ratio q ----
' q = incoming damage per round / rest heal per round (while resting)
restHealPerRound = restHealPerSec * SEC_PER_ROUND
If restHealPerRound > 0# Then
    q = nDmgIN / restHealPerRound
Else
    q = 0#
End If

' Piecewise, single-parameter-free shape:
' - For q <= 0.9 ? boost up to about 1.9?(strongest when q is very small).
' - For q > 0.9  ? damp smoothly; floor at 0.55 to avoid collapsing rest entirely.
If q <= 0# Then
    g = 1#
ElseIf q <= 0.9 Then
    g = (0.9 / q)
    If g > 1.9 Then g = 1.9
Else
    'g = 1# / (1# + 0.6 * (q - 0.9))
    'If g < 0.55 Then g = 0.55
    g = 1 / (1 + 0.45 * (q - 0.9))     ' gentler damping
    If g < 0.4 Then g = 0.4            ' lower floor
End If

restRounds = restRounds * g
If restRounds < 0# Then restRounds = 0#

' ---------- Debug ----------

#If DEVELOPMENT_MODE Then
    If bDebugExpPerHour Then
        DebugLogPrint "HPDBG --- cephA_CalcHPRecoveryRounds ---"
        DebugLogPrint "  Inputs: nDmgIN=" & F6(nDmgIN) & _
                    "; nDmgOut=" & F6(nDmgOut) & _
                    "; nMobHP=" & F6(nMobHP) & _
                    "; nRestHP=" & F6(nRestHP) & _
                    "; nMobs=" & nMobs & _
                    "; nRTC(in)=" & F6(nRTC)
        DebugLogPrint "  Attack: R=" & F6(r) & " rounds; combatSecs=" & F1(combatSecs)
        DebugLogPrint "  Damage: dmgInTotal=" & F1(dmgInTotal) & _
                    "; passiveHealCombat=" & F1(passiveHealCombat) & _
                    "; netDmg=" & F1(netDmg)
        DebugLogPrint "  Rest rate: restHealPerSec=" & F6(restHealPerSec)
        DebugLogPrint "  Rest(full): restTimeContinuous=" & F1(restTimeContinuous) & "s; restRounds=" & F6(restRounds_full) & _
                    " (~" & F1(restRounds_full * SEC_PER_ROUND) & "s)"
        DebugLogPrint "HPDBG --- q-elasticity ---"
        DebugLogPrint "  restHealPerRound=" & F6(restHealPerRound) & _
                    "; q=" & F6(q) & "; g=" & F6(g)
        DebugLogPrint "  restRounds(final)=" & F6(restRounds) & _
                    " (~" & F1(restRounds * SEC_PER_ROUND) & "s)"
    End If
#End If
' ---------------------------

cephA_CalcHPRecoveryRounds = restRounds

out:
On Error Resume Next
Exit Function
error:
Call HandleError("cephA_CalcHPRecoveryRounds")
Resume out:
End Function

'==============================================================================
'  Exp/Hour – Model B (ceph_ModelB) – Overview & Calibration Notes
'  Version: v9.5   Date: 2025-08-09
'------------------------------------------------------------------------------
'  PURPOSE
'    Estimate effective EXP/hour (EPH) for lair-style zones using a smoothed,
'    band-aware model that avoids hard cliffs:
'      • Attack time from effective rounds-to-kill (effRTK) with overkill time.
'      • Recovery time split into HP-rest and Mana/meditation (with relabeling
'        when “no meditate” and light damage).
'      • Movement time from density-, route-, and chain-size–aware travel.
'      • Spawn-gating awareness when the loop outruns respawns.
'
'  TOP-LEVEL FLOW (per loop)
'    1) Setup & validation
'       - Validate inputs; early out if not killable (IsMobKillable).
'       - Density proxy: densGuess = possSpawns/totalLairs (falls back to nAvgWalk).
'       - effRTK:
'           • Casters (spellCost>0): effRTK = clamp(min(nRTK*0.78, RoundUp(nRTK)), 0.74, 8).
'           • Melee when nRTK<2: blend a 10% reduction via SmoothStep(1.2–1.6); floor at 1.
'           • Otherwise RoundUp(nRTK).
'         nRTC = effRTK * nNumMobs.
'
'    2) Kill time & overkill
'       - Base killSecsPerLair = effRTK * nNumMobs * SEC_PER_ROUND.
'       - Overkill inflation ok = cephB_CalcOverkill(dmg, mobHP, isSpell) with caps
'         (spells =1.06, melee =1.18) and near-one-shot blending that tightens the cap
'         around effRTK˜1; melee gets an extra ~3.5% shave near one-shots.
'       - Global chain cut (big chains, short walk) and a mid-band trim for casters
'         (28–40 lairs, walk˜2.4–3.3, density˜2–4) are applied smoothly.
'
'    3) Travel loop seconds (cephB_CalcTravelLoopSecs)
'       Inputs: nAvgWalk, nTotalLairs, nPossSpawns, encumbrance.
'       - Per-room seconds from encumbrance (SECS_ROOM_BASE/HEAVY).
'       - tf multiplier: mild log growth + “small bump”; scarcity scales time when
'         walk >> density; lairOverhead scales with chain size.
'       - Route bands:
'           • Big-chain/short-walk trims (weakened on huge chains).
'           • Junction complexity: add up to +2.5s/lair and +10% tf in a narrow
'             band (30–38 lairs, walk˜2.6–3.2).
'           • Sparse and mid-walk/mid-density trims with smooth fades.
'           • Very sparse × big-chain easing.
'       - Post-blend easing for the same junction band, then:
'         walkLoopSecs *= nGlobal_cephB_Move   ' user knob (default 1.00).
'
'    4) HP recovery demand
'       - Damage per round: hpLossPerRound = max(0, nMobDmg - nDamageThreshold),
'         scaled by nGlobal_cephB_DMG (default 1.00).
'       - Passive HP always ticks (30s) with a small coefficient that grows when
'         per-mob hits are low or walks are long.
'       - Buffer: % of max HP with gates for heavier hits, long-walks, and tiny-chain
'         long routes.
'       - tickBoost vs. minBoost curves set rest intensity; restRateBoost blends
'         per-mob intensity, adds help for “pack” damage, and gives a tiny lift to
'         melee bruisers on short walks. Cap at ~2.07×.
'       - If deficit>0, compute pulse-style restTickHP; restSecs = need ÷ (rest rate).
'
'    5) Mana / meditate demand
'       - totalRounds = effRTK * nNumMobs * nTotalLairs.
'       - manaCostLoop = totalRounds * (nSpellCost + nSpellOverhead),
'         scaled by nGlobal_cephB_Mana (default 1.00).
'       - In-combat regen fraction inCombatMPFrac is density- & walk-aware with caps
'         (caster/no-meditate band trims in 28–40 lairs, walk˜2.4–3.2).
'       - manaRegenSecs = walkLoopSecs + restSecs + inCombatMPFrac * killSecsAll.
'       - poolCredit defaults to 10% of mana; adjusts by band, especially smaller for
'         no-meditate casters in the mid-band (to pull Move% back down).
'       - medNeeded -> medSecs via meditate rate when present, else via passive ticks.
'       - If medSecs=0 and no-meditate, optionally re-label a portion of rest as mana
'         (long-walk+dense or very light incoming damage).
'
'    6) Final assembly & fractions
'       - loopSecs = kill + walk + restSecs + medSecs (loop floor cephB_MIN_LOOP applied
'         before passive regen calc).
'       - EPH: xpPerCycle = nExp * nTotalLairs; cyclesPerHour = 3600/loopSecs;
'         r.nExpPerHour = xpPerCycle * cyclesPerHour * nGlobal_cephB_XP (XP knob).
'       - Fractions reported: Attack, Move, HP (rest), Mana (med), Recover (=HP+MP),
'         Slowdown (effRTK vs. nRTK), Overkill, and RoamTime (spawn gate).
'
'  GLOBAL KNOBS (user-tunable; defaults = 1.00)
'    nGlobal_cephB_XP     ' multiplies final EXP/hr
'    nGlobal_cephB_DMG    ' scales incoming damage per round (affects Rest%)
'    nGlobal_cephB_Mana   ' scales mana cost per loop (affects Mana%)
'    nGlobal_cephB_Move   ' scales travel seconds (affects Move%)
'
'  TICKS / CONSTANTS USED
'    SEC_PER_ROUND        ' combat round (5s typical)
'    SEC_PER_REGEN_TICK   ' passive regen tick (30s)
'    SEC_PER_REST_TICK    ' rest tick (20s)
'    SEC_PER_MEDI_TICK    ' meditate tick (10s)
'    SECS_ROOM_BASE / SECS_ROOM_HEAVY   ' encumbrance-aware travel pacing
'    cephB_LOGISTIC_CAP / cephB_LOGISTIC_DENOM   ' overkill logistic
'    cephB_MIN_LOOP       ' minimum loop envelope for passive regen math
'    cephB_TF_LOG_COEF, cephB_TF_SMALL_BUMP, cephB_TF_SCARCITY_COEF, cephB_LAIR_OVERHEAD_R
'
'  CALIBRATION INTENT
'    • Use S-curves and band weights to create gentle transitions across chains,
'      walk distances, and densities—no hard thresholds.
'    • Keep big-chain (~30–40) with ~3-room walks realistic by adding route
'      complexity instead of globally inflating movement.
'    • Preserve micro-loop caster feel: cap in-combat MP regen; reduce pool credit
'      in the caster mid-band when not meditating so Movement% doesn’t get too low.
'    • Favor mild optimism on EPH while keeping Rest% and Move% within observed
'      ranges across mid-band and sparse-edge cases.
'
'  SAFE TUNERS FOR FUTURE PASSES
'    • User knobs: nGlobal_cephB_XP, nGlobal_cephB_DMG, nGlobal_cephB_Mana, nGlobal_cephB_Move.
'    • Travel feel: cephB_TF_LOG_COEF, cephB_TF_SCARCITY_COEF, cephB_LAIR_OVERHEAD_R,
'      junctionSec scale and the wCX band shape.
'    • Caster feel: no-meditate band trim strength, mpFracHi cap, and poolCredit band.
'    • Rest feel: minBoost bands, restPulseK range, and restRateBoost cap.
'    • Overkill flavor: cephB_LOGISTIC_* and near-one-shot blend caps.
'
'  DEPENDENCIES
'    IsMobKillable(), DebugLogPrint(), HandleError()
'    Helpers: cephB_SmoothStep(), cephB_Lerp(), cephB_MulBlend(), cephB_BandWeight(),
'             cephB_CalcOverkill(), cephB_CalcTravelLoopSecs(),
'             ClampDbl(), SafeDiv(), RoundUp(), MaxDbl(), MinDbl().

'============================ MAIN ==================================
Private Function ceph_ModelB( _
        Optional ByVal nExp As Currency = 0@, _
        Optional ByVal nRegenTime As Double = 0#, _
        Optional ByVal nNumMobs As Integer = 1, _
        Optional ByVal nTotalLairs As Long = -1, _
        Optional ByVal nPossSpawns As Long = 0, _
        Optional ByVal nRTK As Double = 1#, _
        Optional ByVal nCharDMG As Double = 0#, _
        Optional ByVal nCharHP As Long = 0, _
        Optional ByVal nCharHPRegen As Long = 0, _
        Optional ByVal nMobDmg As Double = 0#, _
        Optional ByVal nMobHP As Long = 0, _
        Optional ByVal nMobHPRegen As Long = 0, _
        Optional ByVal nDamageThreshold As Long = 0, _
        Optional ByVal nSpellCost As Integer = 0, _
        Optional ByVal nSpellOverhead As Double = 0#, _
        Optional ByVal nCharMana As Integer = 0, _
        Optional ByVal nCharMPRegen As Long = 0, _
        Optional ByVal nMeditateRate As Long = 0, _
        Optional ByVal nAvgWalk As Double = 0#, _
        Optional ByVal nEncumPct As Integer = 0) As tExpPerHourInfo

On Error GoTo error:

    Dim r As tExpPerHourInfo
    If nRTK <= 0# Then nRTK = 1#
    If nNumMobs <= 0 Then nNumMobs = 1
    If nExp = 0 Then Exit Function

'------------------------------------------------------------------
'  -- DEBUG: raw inputs -------------------------------------------
'------------------------------------------------------------------
    cephB_DebugLog " ------------- ceph_ModelB INPUTS -------------"
    cephB_DebugLog "  nExp=" & nExp & "; nRegenTime=" & nRegenTime & "; nNumMobs=" & nNumMobs & _
                "; nTotalLairs=" & nTotalLairs & "; nPossSpawns=" & nPossSpawns & "; nRTK=" & nRTK
    cephB_DebugLog "  nCharDMG=" & nCharDMG & "; nCharHP=" & nCharHP & "; nCharHPRegen=" & nCharHPRegen & _
                "; nMobDmg=" & nMobDmg & "; nMobHP=" & nMobHP & "; nMobHPRegen=" & nMobHPRegen
    cephB_DebugLog "  nDamageThreshold=" & nDamageThreshold & "; nSpellCost=" & nSpellCost & _
                "; nSpellOverhead=" & nSpellOverhead & "; nCharMana=" & nCharMana
    cephB_DebugLog "  nCharMPRegen=" & nCharMPRegen & "; nMeditateRate=" & nMeditateRate & _
                "; nAvgWalk=" & nAvgWalk & "; nEncumPct=" & nEncumPct

    If Not IsMobKillable(nCharDMG, nCharHP, nMobDmg, nMobHP, nCharHPRegen, nMobHPRegen) Then
        ceph_ModelB.nExpPerHour = -1
        ceph_ModelB.nHitpointRecovery = 1
        ceph_ModelB.nTimeRecovering = 1
        cephB_DebugLog "IsMobKillable", 0
        Exit Function
    End If
    
    '------------------------------------------------------------------
    '  -- NPC / boss shortcut -----------------------------------------
    '------------------------------------------------------------------
    If nTotalLairs <= 0 And nRegenTime > 0 Then
        ceph_ModelB.nExpPerHour = Round(nExp * (1 / nRegenTime))
        Exit Function
    End If
    If nTotalLairs = 0 And nRegenTime = 0 Then Exit Function
    If nTotalLairs <= 0 Then nTotalLairs = 1
    
    Dim densGuess As Double
    densGuess = cephB_CalcDensity(nTotalLairs, nPossSpawns, nAvgWalk)

    '----- effective RTK -----
    Dim effRTK As Double
    If nSpellCost > 0 Then
        effRTK = MaxDbl(MinDbl(nRTK * 0.78, RoundUp(nRTK)), 0.74)
    ElseIf nRTK < 2# Then
        Dim tMelee As Double
        tMelee = cephB_SmoothStep(1.2, 1.6, nRTK)
        effRTK = nRTK * (1# - 0.1 * tMelee)
        effRTK = MaxDbl(effRTK, 1#)
    Else
        effRTK = RoundUp(nRTK)
    End If
    r.nRTC = effRTK * nNumMobs
        
    Dim killSecsPerLair As Double
    killSecsPerLair = effRTK * nNumMobs * SEC_PER_ROUND
    
    '-- Over-kill time inflation -----------------------------------------
    Dim overkillFactor As Double
    overkillFactor = cephB_CalcOverkill(nCharDMG, nMobHP, (nSpellCost > 0))
    
    ' --- One-shotty blend: when effRTK is ~1, dial down the melee overkill cap ---
    Dim okCap As Double
    okCap = IIf(nSpellCost > 0, 1.06, 1.18)
    
    ' tOne=1 at effRTK<=1.05, fades to 0 by ~1.20
    Dim tOne As Double
    tOne = 1# - cephB_SmoothStep(1.05, 1.2, effRTK)
    
    ' Near-one-shot cap to blend toward
    Dim capNearOne As Double
    capNearOne = IIf(nSpellCost > 0, 1.02, 1.06)
    
    ' Blend and clamp
    okCap = cephB_Lerp(okCap, capNearOne, tOne)
    overkillFactor = MinDbl(overkillFactor, okCap)

    killSecsPerLair = killSecsPerLair * overkillFactor
    
    ' Near-one-shot melee: shave a few % off kill time when RTK ~1 (don’t alter spells)
    If nSpellCost = 0 Then
        Dim tNear1 As Double
        tNear1 = 1# - cephB_SmoothStep(1#, 1.15, nRTK)
        Dim nearOneCut As Double
        nearOneCut = cephB_Lerp(1#, 0.965, tNear1)
        killSecsPerLair = killSecsPerLair * nearOneCut
    End If
    
    ' Smoothed global chain cut (>=40 lairs & <=2.5 walk previously)
    Dim wChainL As Double: wChainL = cephB_SmoothStep(32#, 44#, nTotalLairs)
    Dim wChainW As Double: wChainW = 1# - cephB_SmoothStep(2.5, 3.1, nAvgWalk)
    Dim wChain  As Double: wChain = wChainL * wChainW
    killSecsPerLair = cephB_MulBlend(killSecsPerLair, 0.97, wChain)
    cephB_DebugLog "chainCut_w", wChain
    
    ' Targeted mid-band spell/no-meditate trim with blend
    Dim wMBL As Double: wMBL = cephB_BandWeight(nTotalLairs, 28#, 40#, 4#)
    Dim wMBW As Double: wMBW = cephB_SmoothStep(2.4, 2.6, nAvgWalk) * (1# - cephB_SmoothStep(3.3, 3.7, nAvgWalk))
    Dim wMBD As Double: wMBD = cephB_SmoothStep(1.7, 2#, densGuess) * (1# - cephB_SmoothStep(4#, 4.6, densGuess))
    Dim wMB  As Double: wMB = wMBL * wMBW * wMBD * IIf((nSpellCost > 0 Or nSpellOverhead > 0) And nMeditateRate = 0, 1#, 0#)
    killSecsPerLair = cephB_MulBlend(killSecsPerLair, 0.95, wMB)
    cephB_DebugLog "chainCut_midband_w", wMB

    r.nOverkill = overkillFactor - 1#
    cephB_DebugLog "Overkill", r.nOverkill
    cephB_DebugLog "effRTK", effRTK
    cephB_DebugLog "okFactor", overkillFactor
    cephB_DebugLog "killSecs_lair", killSecsPerLair
    cephB_DebugLog "killSecs_all", killSecsPerLair * nTotalLairs

    Dim walkLoopSecs As Double
    walkLoopSecs = cephB_CalcTravelLoopSecs(nAvgWalk, nTotalLairs, nPossSpawns, nEncumPct)
    cephB_DebugLog "walkLoopSecs_base", walkLoopSecs

    ' Ease off the global travel cut on huge chains with modest walk
    'Dim wHuge As Double
    'wHuge = cephB_SmoothStep(30#, 45#, nTotalLairs) * cephB_SmoothStep(2.4, 3.6, nAvgWalk)
    ' Ease off the global travel cut when the big-chain ~3-walk band is active
    Dim wLx As Double, wWx As Double, wCX As Double
    wLx = cephB_BandWeight(nTotalLairs, 30#, 38#, 3#)
    wWx = cephB_SmoothStep(2.6, 3.2, nAvgWalk) * (1# - cephB_SmoothStep(3.6, 3.9, nAvgWalk))
    wCX = wLx * wWx
    
    Dim cutFactor As Double
    cutFactor = cephB_Lerp(0.96, 1#, wCX)
    walkLoopSecs = walkLoopSecs * cutFactor
    
    'MOVEMENT KNOB
    walkLoopSecs = walkLoopSecs * nGlobal_cephB_Move
    cephB_DebugLog "kMove", nGlobal_cephB_Move

    cephB_DebugLog "walkLoopSecs", walkLoopSecs
    
    Dim regenWindow As Double
    regenWindow = nRegenTime * 60#: cephB_DebugLog "regenWindow", regenWindow

    Dim loopSecs As Double
    loopSecs = killSecsPerLair * nTotalLairs + walkLoopSecs
    If loopSecs < cephB_MIN_LOOP Then loopSecs = cephB_MIN_LOOP
    
    '===== HP / Rest =====
    Dim hpLossPerRound As Double
    hpLossPerRound = MaxDbl(0#, nMobDmg - nDamageThreshold)
    
    'HP KNOB
    hpLossPerRound = hpLossPerRound * nGlobal_cephB_DMG
    cephB_DebugLog "kRest", nGlobal_cephB_DMG
    
    ' Per-mob intensity only for gating/smoothing
    Dim hLair As Double: hLair = hpLossPerRound
    Dim hPerMob As Double: hPerMob = SafeDiv(hLair, MaxDbl(1#, nNumMobs))
    cephB_DebugLog "hPerMob", hPerMob

    Dim hpLossPerLoop As Double
    hpLossPerLoop = hpLossPerRound * effRTK * nNumMobs * nTotalLairs
    cephB_DebugLog "hpLossPerLoop", hpLossPerLoop
    
    '- Passive regen that is always ticking
    Dim passiveHP As Double
    Dim passiveCoef As Double: passiveCoef = 0.08
    Dim wLowH As Double: wLowH = 1# - cephB_SmoothStep(10#, 18#, hPerMob)
    Dim wLongWalk As Double: wLongWalk = cephB_SmoothStep(8#, 12#, nAvgWalk)
    passiveCoef = passiveCoef + 0.02 * MaxDbl(wLowH, wLongWalk)
    passiveHP = (nCharHPRegen * passiveCoef) * SafeDiv(loopSecs, SEC_PER_REGEN_TICK)

    Dim hpBuffer As Double
    Dim hGateBuf As Double: hGateBuf = cephB_SmoothStep(24#, 36#, hPerMob)
    Dim wTinyLong As Double
    wTinyLong = cephB_BandWeight(nTotalLairs, 5#, 9#, 1#) * cephB_SmoothStep(8#, 12#, nAvgWalk)
    hpBuffer = nCharHP * (0.04 + 0.015 * hGateBuf + 0.015 * wLongWalk + 0.01 * wTinyLong)
    
    '- First pass: figure out if we *need* to rest
    Dim deficit As Double
    deficit = hpLossPerLoop - passiveHP - hpBuffer
    
    Dim restTickHP As Double
    Dim dmgPerRound As Double: dmgPerRound = MaxDbl(1#, nMobDmg - nDamageThreshold)
    
    ' Continuous minBoost based on damage bands
    Dim h As Double: h = hpLossPerRound
    Dim minBoost As Double
    minBoost = 2.1 _
             + 0.2 * cephB_SmoothStep(1#, 4#, h) _
             + 0.2 * cephB_SmoothStep(10#, 15#, h) _
             + 0.3 * cephB_SmoothStep(25#, 35#, h)

    ' Heavy-rest loop blend (tiny chain + long walk + big hits)
    Dim wHeavy As Double
    wHeavy = cephB_BandWeight(nTotalLairs, 8#, 16#, 4#) * cephB_SmoothStep(5#, 7#, nAvgWalk) * cephB_SmoothStep(10#, 16#, h)
    minBoost = cephB_Lerp(minBoost, MaxDbl(minBoost, 2.8), wHeavy)   ' was 2.5
    cephB_DebugLog "minBoost", minBoost

    Dim tickBoost As Double
    If nCharHPRegen = 0 Then
        tickBoost = 1#
    Else
        tickBoost = ClampDbl(dmgPerRound / nCharHPRegen, 1#, 8#)
        If tickBoost < minBoost Then tickBoost = minBoost
    End If
    
    ' Boost actual resting rate for high incoming damage, gated by per-mob intensity.
    Dim restRateBoost As Double
    Dim hGate As Double: hGate = cephB_SmoothStep(12#, 24#, hPerMob)
    restRateBoost = 1# + 0.54 * (tickBoost - 1#) * hGate
    ' tiny allowance for truly heavy per-mob hits
    If hPerMob >= 30# Then restRateBoost = restRateBoost * 1.005
    ' Extra help if total per-lair damage is moderate/high but per-mob is modest
    Dim hPackGate As Double: hPackGate = cephB_SmoothStep(32#, 60#, hLair)  ' hLair = hpLossPerRound
    restRateBoost = restRateBoost + 0.18 * (tickBoost - 1#) * hPackGate * (1# - cephB_SmoothStep(12#, 24#, hPerMob))
    
    ' Short-walk, big-hit bruiser: micro lift (SIM17 territory)
    If nSpellCost = 0 Then
        Dim wBruiser As Double
        wBruiser = cephB_SmoothStep(28#, 36#, hPerMob) * (1# - cephB_SmoothStep(2.2, 3#, nAvgWalk))
        restRateBoost = restRateBoost * (1# + 0.015 * wBruiser)    ' up to +1.5%
    End If
    
    restRateBoost = ClampDbl(restRateBoost, 1#, 2.07)
    cephB_DebugLog "restRateBoost", restRateBoost
    
    If deficit > 0 Then
        Dim restPulseK As Double: restPulseK = 0.35
        If nSpellCost > 0 Then
            Dim kChain As Double: kChain = cephB_SmoothStep(20#, 36#, nTotalLairs)
            Dim kShort As Double: kShort = 1# - cephB_SmoothStep(1.6, 2.2, nAvgWalk)
            restPulseK = restPulseK + 0.1 * MaxDbl(kChain, kShort)    ' casters: up to 0.45
        Else
            ' NEW: melee bruiser pulse (short-walk + high per-mob hits)
            Dim hGateHi As Double: hGateHi = cephB_SmoothStep(18#, 32#, hPerMob)
            Dim wShortWalk As Double: wShortWalk = 1# - cephB_SmoothStep(2.2, 3.3, nAvgWalk)
            restPulseK = restPulseK + 0.05 * (hGateHi * wShortWalk)   ' melee: up to 0.40
        End If
        restTickHP = nCharHPRegen * restPulseK * MaxDbl(0#, tickBoost - 1#)
    End If
    
    '- Final regen total and rest calculation
    Dim regenHP As Double
    regenHP = passiveHP + restTickHP
    
    Dim restNeeded As Double
    restNeeded = MaxDbl(0#, hpLossPerLoop - regenHP - hpBuffer)
    
    Dim restSecs As Double
    If nCharHPRegen > 0 And restRateBoost > 0 Then
        restSecs = restNeeded * SEC_PER_REST_TICK / (nCharHPRegen * restRateBoost)
    End If
    cephB_DebugLog "restSecs", restSecs
    cephB_DebugLog "hpLossPerRound", hpLossPerRound
    cephB_DebugLog "hpLossPerLoop", hpLossPerLoop
    cephB_DebugLog "restRateBoost", restRateBoost
    cephB_DebugLog "passiveHP", passiveHP
    cephB_DebugLog "deficit", deficit
    cephB_DebugLog "tickBoost", tickBoost
    cephB_DebugLog "restTickHP", restTickHP
    cephB_DebugLog "regenHP", regenHP
    cephB_DebugLog "restNeeded", restNeeded
    
    '===== Mana / Meditate =====
    Dim medSecs As Double
    Dim restSecsDisp As Double: restSecsDisp = restSecs
    Dim medSecsDisp  As Double: medSecsDisp = 0#

    If (nSpellCost > 0 Or nSpellOverhead > 0) Then
        Dim manaCostLoop As Double
        Dim totalRounds  As Double
        Dim manaGain     As Double
        Dim killSecsAll  As Double

        totalRounds = effRTK * nNumMobs * nTotalLairs
        manaCostLoop = totalRounds * (nSpellCost + nSpellOverhead)
        
        'MANA KNOB
        manaCostLoop = manaCostLoop * nGlobal_cephB_Mana
        cephB_DebugLog "kMana", nGlobal_cephB_Mana

        killSecsAll = killSecsPerLair * nTotalLairs
        cephB_DebugLog "killSecsAll", killSecsAll

        Dim inCombatMPFrac As Double
        If nMeditateRate > 0 Then
            inCombatMPFrac = 0.26
            If nTotalLairs >= 28 And nAvgWalk <= 3.5 Then inCombatMPFrac = inCombatMPFrac + 0.02
        Else
            cephB_DebugLog "dens_guess", densGuess

            inCombatMPFrac = 0.31 - 0.035 * nAvgWalk
            inCombatMPFrac = inCombatMPFrac _
                + 0.04 * (1# - cephB_SmoothStep(2#, 2.6, nAvgWalk)) _
                + 0.015 * (1# - cephB_SmoothStep(1.6, 1.9, nAvgWalk)) _
                + 0.01 * cephB_SmoothStep(30#, 50#, densGuess) _
                + 0.01 * cephB_SmoothStep(70#, 90#, densGuess) * (1# - cephB_SmoothStep(1.4, 1.8, nAvgWalk)) _
                + 0.01 * IIf(nSpellCost > 0, 1#, 0#) _
                + 0.01 * IIf(nSpellCost > 0 And nMeditateRate = 0, 1#, 0#) _
                          * cephB_SmoothStep(2.4, 2.6, nAvgWalk) * (1# - cephB_SmoothStep(3.5, 3.9, nAvgWalk)) _
                          * cephB_SmoothStep(1.7, 2#, densGuess) * (1# - cephB_SmoothStep(4#, 4.6, densGuess))

            Dim mpFracHi As Double: mpFracHi = 0.34
            If nSpellCost > 0 And nMeditateRate = 0 Then
                mpFracHi = 0.36
                If densGuess >= 60# And nAvgWalk <= 1.6 Then mpFracHi = 0.38
            End If

            ' Extra micro-bump + cap lift for ultra-dense, very short walk
            If densGuess >= 80# And nAvgWalk <= 1.4 Then
                inCombatMPFrac = inCombatMPFrac + 0.005
                If nSpellCost > 0 And nMeditateRate = 0 Then
                    mpFracHi = MaxDbl(mpFracHi, 0.385)
                End If
                cephB_DebugLog "mpFrac_ultradense_shortwalk_bump2", inCombatMPFrac
            End If

            ' Mid-band weight (same shape as travel band)
            Dim wMBn As Double
            wMBn = cephB_BandWeight(nTotalLairs, 28#, 40#, 4#) _
                 * cephB_SmoothStep(2.4, 2.6, nAvgWalk) * (1# - cephB_SmoothStep(3.3, 3.7, nAvgWalk)) _
                 * cephB_SmoothStep(1.7, 2#, densGuess) * (1# - cephB_SmoothStep(4#, 4.6, densGuess))
            
            ' Tiny in-combat regen bump in the band to offset travel trims
            inCombatMPFrac = inCombatMPFrac + 0.005 * wMBn
            
            ' Allow a hair more headroom in the same band
            If wMBn > 0# Then mpFracHi = MaxDbl(mpFracHi, 0.37)
            
            cephB_DebugLog "mpFrac_preClamp", inCombatMPFrac
            inCombatMPFrac = ClampDbl(inCombatMPFrac, 0.1, mpFracHi)
            
            If nMeditateRate = 0 Then
                Dim wNoMedBand As Double
                wNoMedBand = cephB_BandWeight(nTotalLairs, 28#, 40#, 4#) _
                           * cephB_SmoothStep(2.4, 3.2, nAvgWalk) * (1# - cephB_SmoothStep(3.4, 3.9, nAvgWalk))
            
                ' stronger than the earlier -0.02: SIM3 needs a real nudge
                inCombatMPFrac = MaxDbl(0.1, inCombatMPFrac - 0.035 * wNoMedBand)
            End If
        End If
        cephB_DebugLog "inCombatMPFrac", inCombatMPFrac

        Dim manaRegenSecs As Double
        manaRegenSecs = walkLoopSecs + restSecs + inCombatMPFrac * killSecsAll
        cephB_DebugLog "manaRegenSecs", manaRegenSecs

        manaGain = nCharMPRegen * SafeDiv(manaRegenSecs, SEC_PER_REGEN_TICK)
        cephB_DebugLog "manaGain", manaGain

        Dim poolCredit As Double: poolCredit = nCharMana * 0.1
        If nSpellCost > 0 And nMeditateRate = 0 Then
            If densGuess >= 60# And nAvgWalk <= 1.6 Then
                poolCredit = nCharMana * 0.16            ' ultra-dense micro-loops
            ElseIf nAvgWalk >= 2.5 And nAvgWalk <= 3.5 And densGuess >= 2# And densGuess <= 4# Then
                Dim wMBn_pc As Double
                wMBn_pc = cephB_BandWeight(nTotalLairs, 28#, 40#, 4#) _
                        * cephB_SmoothStep(2.4, 2.6, nAvgWalk) * (1# - cephB_SmoothStep(3.3, 3.7, nAvgWalk)) _
                        * cephB_SmoothStep(1.7, 2#, densGuess) * (1# - cephB_SmoothStep(4#, 4.6, densGuess))
            
                ' No-med in this band: smaller pool so more med time (pulling Move% back down)
                poolCredit = nCharMana * cephB_Lerp(0.1, 0.13, wMBn_pc)
            End If
        End If

        Dim medNeeded As Double
        medNeeded = MaxDbl(0#, manaCostLoop - manaGain - poolCredit)
        cephB_DebugLog "medNeeded", medNeeded

        If nMeditateRate > 0 And medNeeded >= nMeditateRate / 2# Then
            medSecs = (medNeeded / nMeditateRate) * SEC_PER_MEDI_TICK
        ElseIf nMeditateRate = 0 And nCharMPRegen > 0 Then
            medSecs = (medNeeded / nCharMPRegen) * SEC_PER_REGEN_TICK
        Else
            medSecs = 0#
        End If
        cephB_DebugLog "medSecs", medSecs

        medSecsDisp = medSecs
        If medSecs = 0# And nMeditateRate = 0 Then
            Dim relabelCapPct As Double: relabelCapPct = 0#
        
            If nAvgWalk >= 8# And densGuess >= 12# Then
                relabelCapPct = 0.55           ' lets ~half of rest show as "mana"
            ElseIf hpLossPerRound <= 8# Then
                relabelCapPct = 0.35
            End If
        
            If relabelCapPct > 0# Then
                Dim manaRegenSecsNoRest As Double
                manaRegenSecsNoRest = walkLoopSecs + inCombatMPFrac * killSecsAll
        
                Dim manaGainNoRest As Double
                manaGainNoRest = nCharMPRegen * SafeDiv(manaRegenSecsNoRest, SEC_PER_REGEN_TICK)
        
                Dim medNeededNoRest As Double
                medNeededNoRest = MaxDbl(0#, manaCostLoop - manaGainNoRest - poolCredit)
        
                If medNeededNoRest > 0# And nCharMPRegen > 0# Then
                    Dim relabel As Double
                    relabel = (medNeededNoRest / nCharMPRegen) * SEC_PER_REGEN_TICK
                    relabel = MinDbl(relabel, restSecsDisp * relabelCapPct)
        
                    medSecsDisp = medSecsDisp + relabel
                    restSecsDisp = restSecsDisp - relabel
        
                    cephB_DebugLog "relabel_cap_pct", relabelCapPct
                    cephB_DebugLog "relabel_medSecs", medSecsDisp
                    cephB_DebugLog "relabel_restSecs", restSecsDisp
                End If
            End If
        End If
    Else
        ' No mana consumer -> skip mana model entirely
        cephB_DebugLog "mana_skip", 1#
        medSecs = 0#
        medSecsDisp = 0#
        ' (restSecsDisp stays as restSecs)
    End If
    
    '===== Final loop time =====
    loopSecs = loopSecs + restSecs + medSecs
    cephB_DebugLog "loopSecs", loopSecs

    Dim xpPerCycle As Double
    xpPerCycle = nExp * nTotalLairs
    cephB_DebugLog "xpPerCycle", xpPerCycle
    
    Dim cyclesPerHour As Double: cyclesPerHour = SafeDiv(3600#, loopSecs)
    cephB_DebugLog "cyclesPerHour", cyclesPerHour
    
    Dim killShare As Double: killShare = SafeDiv(killSecsPerLair * nTotalLairs, loopSecs)
    Dim walkShare As Double: walkShare = SafeDiv(walkLoopSecs, loopSecs)
    Dim restShare As Double: restShare = SafeDiv(restSecsDisp, loopSecs)
    Dim medShare  As Double: medShare = SafeDiv(medSecsDisp, loopSecs)
    cephB_DebugLog "killShare", killShare
    cephB_DebugLog "walkShare", walkShare
    cephB_DebugLog "restShare ", restShare
    cephB_DebugLog "medShare", medShare

    '===== Pack (ModelB-style, like ModelA strings) =====================
    r.nExpPerHour = xpPerCycle * cyclesPerHour
    
    'EXP KNOB
    r.nExpPerHour = r.nExpPerHour * nGlobal_cephB_XP
    cephB_DebugLog "kXP", nGlobal_cephB_XP

    r.nHitpointRecovery = SafeDiv(restSecsDisp, loopSecs)
    r.nManaRecovery = SafeDiv(medSecsDisp, loopSecs)
    r.nTimeRecovering = r.nHitpointRecovery + r.nManaRecovery
    r.nMove = SafeDiv(walkLoopSecs, loopSecs)
    If r.nMove > 1# Then r.nMove = 1#
    r.nOverkill = MaxDbl(0#, overkillFactor - 1#)
    
    ' Fractions for text (mirrors ModelA semantics)
    Dim attackFrac As Double:    attackFrac = SafeDiv(killSecsPerLair * nTotalLairs, loopSecs)
    Dim slowdownFrac As Double:  slowdownFrac = IIf(nRTK > 0#, MaxDbl(0#, SafeDiv(effRTK, nRTK) - 1#), 0#)
    Dim overshootFrac As Double: overshootFrac = r.nOverkill
    Dim recoverFrac As Double:   recoverFrac = r.nTimeRecovering
    Dim hitpointFrac As Double:  hitpointFrac = r.nHitpointRecovery
    Dim manaFrac As Double:      manaFrac = r.nManaRecovery
    Dim moveFrac As Double:      moveFrac = r.nMove

    ' --- Insufficient-lairs (respawn gating) display metric ---
    Dim roamShare As Double, respawnGated As Double
    If regenWindow > 0# And loopSecs < regenWindow Then
        roamShare = ClampDbl((regenWindow - loopSecs) / regenWindow, 0#, 1#)
        respawnGated = 1#
    Else
        roamShare = 0#
        respawnGated = 0#
    End If
    cephB_DebugLog "respawnGated", respawnGated
    cephB_DebugLog "roamShare", roamShare
    
    r.nExpPerHour = r.nExpPerHour
    r.nHitpointRecovery = hitpointFrac
    r.nManaRecovery = manaFrac
    r.nTimeRecovering = recoverFrac
    r.nMove = moveFrac
    r.nOverkill = overshootFrac
    r.nSlowdownTime = slowdownFrac
    r.nAttackTime = attackFrac
    r.nRoamTime = roamShare

    ceph_ModelB = r

out:
On Error Resume Next
Exit Function
error:
Call HandleError("ceph_ModelB")
Resume out:
End Function

'=========================== CalcExpPerHour – Model B UTILITIES ===============================
Private Sub cephB_DebugLog(ByVal lbl As String, Optional ByVal val As Double = -99999)
#If DEVELOPMENT_MODE Then
    If bDebugExpPerHour Then
        If val <> -99999 Then
            DebugLogPrint "EPH-DBG " & lbl & "=" & Format$(val, "0.########")
        Else
            DebugLogPrint "EPH-DBG " & lbl
        End If
    End If
#End If
End Sub

'------------- smoothing helpers -------------
Private Function cephB_Saturate(ByVal x As Double) As Double
    If x <= 0# Then
        cephB_Saturate = 0#
    ElseIf x >= 1# Then
        cephB_Saturate = 1#
    Else
        cephB_Saturate = x
    End If
End Function

' SmoothStep(edge0, edge1, x): 0?1 with eased S-curve
Private Function cephB_SmoothStep(ByVal edge0 As Double, ByVal edge1 As Double, ByVal x As Double) As Double
    If edge0 = edge1 Then
        cephB_SmoothStep = IIf(x >= edge1, 1#, 0#)
        Exit Function
    End If
    Dim t As Double: t = cephB_Saturate((x - edge0) / (edge1 - edge0))
    cephB_SmoothStep = t * t * (3# - 2# * t)
End Function

' Lerp(a,b,t): a + (b-a)*t
Private Function cephB_Lerp(ByVal a As Double, ByVal b As Double, ByVal t As Double) As Double
    cephB_Lerp = a + (b - a) * t
End Function

' Multiply-by-factor with blend t (t=0 -> *1, t=1 -> *factor)
Private Function cephB_MulBlend(ByVal cur As Double, ByVal factor As Double, ByVal t As Double) As Double
    cephB_MulBlend = cur * cephB_Lerp(1#, factor, cephB_Saturate(t))
End Function

' “Band” weight for lo..hi with soft fades on both sides
Private Function cephB_BandWeight(ByVal x As Double, ByVal lo As Double, ByVal hi As Double, Optional ByVal fade As Double = 2#) As Double
    Dim wIn As Double:  wIn = cephB_SmoothStep(lo - fade, lo, x)
    Dim wOut As Double: wOut = 1# - cephB_SmoothStep(hi, hi + fade, x)
    cephB_BandWeight = cephB_Saturate(wIn * wOut)
End Function

'---------------- over-kill factor -------------
Private Function cephB_CalcOverkill(ByVal dmg As Double, ByVal mobHP As Long, ByVal isSpell As Boolean) As Double
On Error GoTo error:

    If mobHP <= 0 Then cephB_CalcOverkill = 1#: Exit Function
    Dim raw As Double: raw = (mobHP - dmg) / (cephB_LOGISTIC_DENOM * mobHP)
    raw = ClampDbl(raw, -cephB_LOGISTIC_CAP, cephB_LOGISTIC_CAP)

    Dim mult As Double: mult = IIf(isSpell, 1.35, 1#)
    Dim ok As Double: ok = 1# + 1# / (1# + Exp(raw * mult))
    If isSpell Then
        cephB_CalcOverkill = ClampDbl(ok, 1#, 1.06)
    Else
        cephB_CalcOverkill = ClampDbl(ok, 1#, 1.18)
    End If


out:
On Error Resume Next
Exit Function
error:
Call HandleError("cephB_CalcOverkill")
Resume out:
End Function


'---------------- travel helpers ----------------
Private Function cephB_CalcDensity(ByVal totalLairs As Long, ByVal possSpawns As Long, ByVal avgWalk As Double) As Double
    ' Rooms per lair; fallback to avgWalk if data is missing
    If totalLairs > 0 And possSpawns > 0 Then
        cephB_CalcDensity = SafeDiv(possSpawns, totalLairs, avgWalk)
    Else
        cephB_CalcDensity = avgWalk
    End If
End Function


Private Function cephB_CalcTravelLoopSecs(ByVal avgWalk As Double, ByVal totalLairs As Long, _
                                    ByVal possSpawns As Long, ByVal encPct As Integer) As Double
On Error GoTo error:

Dim secPerRoom As Double, dens As Double, scarcity As Double
Dim tf As Double, lairOverhead As Double, baseRooms As Double, damp As Double, aw As Double

' Encumbrance -> per-room time
secPerRoom = IIf(encPct >= HEAVY_ENCUM_PCT, SECS_ROOM_HEAVY, SECS_ROOM_BASE)

' Rooms per lair (fallback to avgWalk if missing inputs)
dens = cephB_CalcDensity(totalLairs, possSpawns, avgWalk)
cephB_DebugLog "dens", dens

' Scarcity: more time if avgWalk >> density (rooms/lair)
If dens >= 5# Then
    scarcity = 1# + (cephB_TF_SCARCITY_COEF - 0.03) * SafeDiv(avgWalk, MaxDbl(1#, dens))
Else
    scarcity = 1# + cephB_TF_SCARCITY_COEF * SafeDiv(avgWalk, MaxDbl(1#, dens))
End If

' Short-walk bump + mild log growth
tf = 1# + cephB_TF_LOG_COEF * Log(1# + avgWalk) + cephB_TF_SMALL_BUMP / (1# + avgWalk)
If avgWalk <= 1.6 And dens >= 30# Then
    tf = tf * 0.93
    cephB_DebugLog "tf_microcut", tf
End If

' Base overheads
lairOverhead = cephB_LAIR_OVERHEAD_R * secPerRoom
baseRooms = avgWalk * secPerRoom

' Scale overheads for lair count
Dim overheadScale As Double
overheadScale = 0.6 + 0.4 * MinDbl(1#, 20# / MaxDbl(1#, totalLairs))
Dim scaleUp As Double
scaleUp = 0.06 * cephB_SmoothStep(30#, 45#, totalLairs) * cephB_SmoothStep(2.4, 3.6, avgWalk)
overheadScale = overheadScale + scaleUp
lairOverhead = lairOverhead * overheadScale
cephB_DebugLog "overheadScale", overheadScale

' ---- Smoothed micro-route shaves ----
Dim wShort As Double:  wShort = 1# - cephB_SmoothStep(1.6, 2.2, avgWalk)
Dim wDense As Double:  wDense = cephB_SmoothStep(50#, 70#, dens)
Dim wUD    As Double:  wUD = wShort * wDense
tf = cephB_MulBlend(tf, 0.91, wUD)
lairOverhead = cephB_MulBlend(lairOverhead, 0.91, wUD)
cephB_DebugLog "wUD", wUD

Dim wShort2 As Double: wShort2 = 1# - cephB_SmoothStep(1.4, 1.9, avgWalk)
Dim wDense2 As Double: wDense2 = cephB_SmoothStep(70#, 90#, dens)
Dim wUD2    As Double: wUD2 = wShort2 * wDense2
tf = cephB_MulBlend(tf, 0.97, wUD2)
lairOverhead = cephB_MulBlend(lairOverhead, 0.96, wUD2)
cephB_DebugLog "wUD2", wUD2

' Only damp LONG routes: no penalty until ~6 rooms
aw = MaxDbl(0#, avgWalk - 5#)
damp = 1# / (1# + 0.12 * aw ^ 1.4)

' --- Route band tweaks ---
If totalLairs >= 12 And totalLairs <= 16 And avgWalk >= 5# Then
    ' Keep this one discrete (tiny population + explicit target)
    tf = tf * 0.75
    damp = damp * 0.75
    lairOverhead = lairOverhead * 0.8
    If avgWalk >= 6# Then
        tf = tf * 0.97
        damp = damp * 0.92
        lairOverhead = lairOverhead * 0.94
        cephB_DebugLog "midcut_longwalk2", tf
    End If
    If totalLairs <= 13 And avgWalk >= 6# Then
        tf = tf * 0.97
        lairOverhead = lairOverhead * 0.96
        cephB_DebugLog "midcut_tinychain", tf
    End If
    cephB_DebugLog "tf_midcut", tf
    cephB_DebugLog "damp_midcut", damp
    cephB_DebugLog "lairOverhead_midcut", lairOverhead

Else
    ' Smoothed big-chain / short-walk regime
    Dim wBig  As Double: wBig = cephB_SmoothStep(24#, 34#, totalLairs)
    Dim wShort3 As Double: wShort3 = 1# - cephB_SmoothStep(3.3, 4.2, avgWalk)
    Dim wLowWalk As Double: wLowWalk = wBig * wShort3
    ' On very large chains (30–45 lairs) with modest walk, reduce trim strength
    Dim wHuge As Double
    wHuge = cephB_SmoothStep(30#, 45#, totalLairs) * cephB_SmoothStep(2.4, 3.6, avgWalk)
    
    ' Up to 50% weaker trim on huge chains; trims themselves less aggressive
    Dim wLL As Double: wLL = wLowWalk * (1# - 0.5 * wHuge)
    tf = cephB_MulBlend(tf, 0.94, wLL)
    lairOverhead = cephB_MulBlend(lairOverhead, 0.96, wLL)
    cephB_DebugLog "wLowWalk", wLowWalk
    
    ' --- NEW: route complexity for big-chains with ~3-room walks (SIM2/16) ---
    Dim wLx As Double, wWx As Double, wCX As Double
    wLx = cephB_BandWeight(totalLairs, 30#, 38#, 3#)
    wWx = cephB_SmoothStep(2.6, 3.2, avgWalk) * (1# - cephB_SmoothStep(3.6, 3.9, avgWalk))
    wCX = wLx * wWx
    
    ' Add banded “junction” seconds per lair (raw seconds, not scaled by secPerRoom)
    Dim junctionSec As Double
    junctionSec = 2.5 * wCX                      ' up to +2.5s per lair at peak
    lairOverhead = lairOverhead + junctionSec
    
    ' A little extra pathing inefficiency on the loop multiplier
    tf = cephB_MulBlend(tf, 1.1, wCX)            ' up to +10% on TF in this narrow band
    
    cephB_DebugLog "wCX", wCX
    cephB_DebugLog "junctionSec", junctionSec
    
    ' Scarcity easing within low-walk
    Dim ratio As Double: ratio = SafeDiv(avgWalk, MaxDbl(1#, dens))
    Dim scCoef As Double
    scCoef = (cephB_TF_SCARCITY_COEF - 0.02 * wLowWalk) + 0.03 * wHuge  ' add back a bit on huge chains
    scarcity = 1# + scCoef * ratio

    ' Sparse chains inside (dens<~6) with soft blend
    Dim wSparse As Double: wSparse = 1# - cephB_SmoothStep(5#, 7#, dens)
    Dim wSparseBand As Double: wSparseBand = cephB_BandWeight(totalLairs, 28#, 40#, 4#) * wShort3 * wSparse
    tf = cephB_MulBlend(tf, 0.91, wSparseBand)
    lairOverhead = cephB_MulBlend(lairOverhead, 0.94, wSparseBand)
    cephB_DebugLog "wSparseBand", wSparseBand

    ' Mid-walk / mid-density band trim (2.6–3.3 walk, 2–4 dens) smoothed
    Dim wWalkBand As Double: wWalkBand = cephB_SmoothStep(2.4, 2.6, avgWalk) * (1# - cephB_SmoothStep(3.3, 3.7, avgWalk))
    Dim wDensBand As Double: wDensBand = cephB_SmoothStep(1.7, 2#, dens) * (1# - cephB_SmoothStep(4#, 4.6, dens))
    Dim wBand As Double:     wBand = wBig * wWalkBand * wDensBand
    tf = cephB_MulBlend(tf, 0.93, wBand)
    lairOverhead = cephB_MulBlend(lairOverhead, 0.88, wBand)
    Dim scEase As Double: scEase = (cephB_TF_SCARCITY_COEF - 0.12)
    scarcity = 1# + cephB_Lerp(cephB_TF_SCARCITY_COEF, scEase, wBand) * ratio
    cephB_DebugLog "wBand", wBand
End If
' --- end band tweaks ---

' Very sparse & big chain smoothing
Dim wVerySparse As Double
wVerySparse = (1# - cephB_SmoothStep(1.6, 2.4, dens)) * cephB_SmoothStep(36#, 44#, totalLairs)
tf = cephB_MulBlend(tf, 0.985, wVerySparse)
scarcity = cephB_Lerp(scarcity, scarcity * 0.97, wVerySparse)
cephB_DebugLog "wVerySparse", wVerySparse

cephB_DebugLog "secPerRoom", secPerRoom
cephB_DebugLog "lairOverhead", lairOverhead
cephB_DebugLog "dens", dens
cephB_DebugLog "scarcity", scarcity
cephB_DebugLog "tf", tf
cephB_DebugLog "damp", damp

cephB_CalcTravelLoopSecs = (totalLairs * (baseRooms + lairOverhead)) * tf * scarcity * damp

out:
On Error Resume Next
Exit Function
error:
Call HandleError("cephB_CalcTravelLoopSecs")
Resume out:
End Function

'============================ END ===================================
