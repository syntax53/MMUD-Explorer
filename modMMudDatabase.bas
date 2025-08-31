Attribute VB_Name = "modMMudDatabase"
Option Explicit
Option Base 0

Global UseExpMulti As Boolean

Public nSpellNest As Integer

Public DB As Database
Public tabItems As Recordset
Public tabClasses As Recordset
Public tabRaces As Recordset
Public tabSpells As Recordset
Public tabInfo As Recordset
Public tabMonsters As Recordset
Public tabShops As Recordset
Public tabRooms As Recordset
Public tabTBInfo As Recordset
Public tabTempRS As Recordset
Public tabLairs As Recordset

Public nMonsterDamageVsDefault() As Currency
Public nMonsterDamageVsChar() As Currency
Public nMonsterDamageVsParty() As Currency
Public nMonsterPossy() As Currency
Public nMonsterSpawnChance() As Currency
Public bQuickSpell As Boolean

Public nCharDamageVsMonster() As Currency
Public nCharMinDamageVsMonster() As Currency
Public nCharSurpriseDamageVsMonster() As Currency
Public sCharDamageVsMonsterConfig As String

Public Type MonAttackSimReturn
    nAverageDamage As Currency
    nMaxDamage As Currency
End Type

Public Type SpellMinMaxDur
    nMin As Currency
    nMax As Currency
    nDur As Currency
    sMin As String
    sMax As String
    sDur As String
    bNoHeader As Boolean
End Type

Public Type typItemCostDetail
    Cost As Long
    Coin As Long
End Type

Public Enum enmMagicEnum
    None = 0
    Mage = 1
    Priest = 2
    Druid = 3
    Bard = 4
    Kai = 5
End Enum

Public Enum MVDatType
    item = 0
    Spell = 1
    Monster = 2
    Shop = 3
    Class = 4
    Race = 5
    Text = 6
    Room = 7
End Enum

Public dictLairInfo As Dictionary
Public Type LairInfoType
    sGroupIndex As String
    sMobList As String
    nMobs As Currency
    nMaxRegen As Currency
    nAvgExp As Currency
    nAvgDmg As Currency 'avg dmg/mob/round (single mob, alone)
    nAvgHP As Long
    nAvgAC As Integer
    nAvgDR As Integer
    nAvgMR As Integer
    nAvgDodge As Integer
    nTotalLairs As Long
    nAvgWalk As Currency
    nAvgDelay As Integer
    nDamageMitigated As Long
    nDamageOut As Long
    nMinDamageOut As Long
    nSurpriseDamageOut As Long
    nPossSpawns As Long
    sGlobalAttackConfig As String
    nAvgDmgLair As Currency 'avg dmg/round taken to clear lair of all mobs
    nRTK As Double 'rounds to kill each mob
    nRTC As Double 'rounds to clear the lair
    nMagicLVL As Integer
    nMaxMagicLVL As Integer
    nSpellImmuLVL As Integer
    nMaxSpellImmuLVL As Integer
    nNumUndeads As Integer
    nNumAntiMagic As Integer
End Type
Dim colLairs() As LairInfoType


Public Function GetLairAveragesFromLocs(ByVal sLoc As String) As LairInfoType
On Error GoTo error:
Dim sGroupIndex As String, iLair As Integer, nLairs As Long, nMaxRegen As Currency
Dim sRegexPattern As String, tMatches() As RegexMatches, tLairInfo As LairInfoType
Dim tmp_nAvgDmg As Currency, tmp_nAvgExp As Currency, tmp_nAvgHP As Currency, tmp_nAvgDodge As Long
Dim tmp_nMaxRegen As Currency, tmp_nAvgDmgLair As Currency, tmp_nAvgDelay As Integer
Dim tmp_sMobList As String, tmp_nAvgAC As Long, tmp_nAvgDR As Long, tmp_nAvgMR As Long, tmp_nAvgMitigation As Currency
Dim tmp_nRTC As Double, tmp_nRTK As Double, tmp_nAvgDamageOut As Currency, tmp_nAvgMobs As Double
Dim tmp_nAvgWalk() As Double, tmp_nSurpriseDamageOut As Currency, tmp_nMinDmgOut As Double
Dim tmp_nMaxMagicLVL As Integer, tmp_nMagicLVL As Double, tmp_nMaxSpellImmuLVL As Integer, tmp_nSpellImmuLVL As Double
Dim tmp_nAvgNumUndeads As Double, tmp_nAvgNumAntiMagic As Double, nDmgOut() As Currency

GetLairAveragesFromLocs.nPossSpawns = InstrCount(tabMonsters.Fields("Summoned By"), "Group:")

If nNMRVer < 1.83 Then Exit Function
sRegexPattern = "\[([\d\-]+)\]\[(\d+)\]Group\(lair\): (\d+)\/(\d+)"

tMatches() = RegExpFindv2(sLoc, sRegexPattern)
If UBound(tMatches()) > 0 Or Len(tMatches(0).sFullMatch) > 0 Then
    
    nLairs = UBound(tMatches()) + 1
    ReDim tmp_nAvgWalk(UBound(tMatches()))
    For iLair = 0 To UBound(tMatches())
        
        '[7-8-9][6]Group(lair): 1/2345
        sGroupIndex = tMatches(iLair).sSubMatches(0)
        nMaxRegen = val(tMatches(iLair).sSubMatches(1))
        If nMaxRegen > 0 Then
            
            tLairInfo = GetLairInfo(sGroupIndex, nMaxRegen)
            If tLairInfo.nMobs > 0 Then
                tmp_nAvgMobs = tmp_nAvgMobs + tLairInfo.nMobs
                tmp_nAvgExp = tmp_nAvgExp + (tLairInfo.nAvgExp * tLairInfo.nMaxRegen)
                tmp_nAvgHP = tmp_nAvgHP + (tLairInfo.nAvgHP * tLairInfo.nMaxRegen)
                tmp_nAvgDmg = tmp_nAvgDmg + tLairInfo.nAvgDmg
                tmp_nAvgDmgLair = tmp_nAvgDmgLair + tLairInfo.nAvgDmgLair
                tmp_nRTC = tmp_nRTC + tLairInfo.nRTC
                tmp_nRTK = tmp_nRTK + tLairInfo.nRTK
                tmp_nAvgAC = tmp_nAvgAC + tLairInfo.nAvgAC
                tmp_nAvgDR = tmp_nAvgDR + tLairInfo.nAvgDR
                tmp_nAvgMR = tmp_nAvgMR + tLairInfo.nAvgMR
                tmp_nAvgDodge = tmp_nAvgDodge + tLairInfo.nAvgDodge
                tmp_nAvgDamageOut = tmp_nAvgDamageOut + tLairInfo.nDamageOut
                tmp_nMinDmgOut = tmp_nMinDmgOut + tLairInfo.nMinDamageOut
                tmp_nSurpriseDamageOut = tmp_nSurpriseDamageOut + tLairInfo.nSurpriseDamageOut
                tmp_nAvgMitigation = tmp_nAvgMitigation + tLairInfo.nDamageMitigated
                
                If tLairInfo.nMagicLVL > tmp_nMagicLVL Then tmp_nMagicLVL = tLairInfo.nMagicLVL
                If tLairInfo.nMaxMagicLVL > tmp_nMaxMagicLVL Then tmp_nMaxMagicLVL = tLairInfo.nMaxMagicLVL
                
                If tLairInfo.nSpellImmuLVL > tmp_nSpellImmuLVL Then tmp_nSpellImmuLVL = tLairInfo.nSpellImmuLVL
                If tLairInfo.nMaxSpellImmuLVL > tmp_nMaxSpellImmuLVL Then tmp_nMaxSpellImmuLVL = tLairInfo.nMaxSpellImmuLVL
                
                tmp_nAvgNumUndeads = tmp_nAvgNumUndeads + tLairInfo.nNumUndeads
                tmp_nAvgNumAntiMagic = tmp_nAvgNumAntiMagic + tLairInfo.nNumAntiMagic
                
                tmp_nMaxRegen = tmp_nMaxRegen + tLairInfo.nMaxRegen
                tmp_nAvgDelay = tmp_nAvgDelay + tLairInfo.nAvgDelay
                tmp_nAvgWalk(iLair) = tLairInfo.nAvgWalk
                
                tmp_sMobList = AutoAppend(tmp_sMobList, tLairInfo.sMobList, ",")
            End If
            
        End If
    Next iLair
    
    GetLairAveragesFromLocs.nAvgDmg = Round(tmp_nAvgDmg / nLairs)
    GetLairAveragesFromLocs.nAvgDmgLair = Round(tmp_nAvgDmgLair / nLairs)
    GetLairAveragesFromLocs.nRTC = Round(tmp_nRTC / nLairs, 1)
    GetLairAveragesFromLocs.nRTK = Round(tmp_nRTK / nLairs, 1)
    GetLairAveragesFromLocs.nAvgExp = Round(tmp_nAvgExp / nLairs)
    GetLairAveragesFromLocs.nAvgHP = Round(tmp_nAvgHP / nLairs)
    GetLairAveragesFromLocs.nAvgAC = Round(tmp_nAvgAC / nLairs)
    GetLairAveragesFromLocs.nAvgDR = Round(tmp_nAvgDR / nLairs)
    GetLairAveragesFromLocs.nAvgMR = Round(tmp_nAvgMR / nLairs)
    GetLairAveragesFromLocs.nAvgDodge = Round(tmp_nAvgDodge / nLairs)
    GetLairAveragesFromLocs.nDamageMitigated = Round(tmp_nAvgMitigation / nLairs)
    GetLairAveragesFromLocs.nMobs = (tmp_nAvgMobs / nLairs)
    GetLairAveragesFromLocs.nMaxRegen = Round(tmp_nMaxRegen / nLairs, 1)
    GetLairAveragesFromLocs.nAvgDelay = Round(tmp_nAvgDelay / nLairs, 1)
    
    GetLairAveragesFromLocs.nMagicLVL = tmp_nMagicLVL 'RoundUp(tmp_nMagicLVL / nLairs)
    GetLairAveragesFromLocs.nMaxMagicLVL = tmp_nMaxMagicLVL
    'If GetLairAveragesFromLocs.nMagicLVL >= (tmp_nMaxMagicLVL / 2) Then GetLairAveragesFromLocs.nMagicLVL = tmp_nMaxMagicLVL
    
    GetLairAveragesFromLocs.nSpellImmuLVL = tmp_nSpellImmuLVL 'RoundUp(tmp_nSpellImmuLVL / nLairs)
    GetLairAveragesFromLocs.nMaxSpellImmuLVL = tmp_nMaxSpellImmuLVL
    'If GetLairAveragesFromLocs.nSpellImmuLVL >= (tmp_nMaxSpellImmuLVL / 2) Then GetLairAveragesFromLocs.nSpellImmuLVL = tmp_nMaxSpellImmuLVL
    
    GetLairAveragesFromLocs.nNumUndeads = RoundUp(tmp_nAvgNumUndeads / nLairs)
    GetLairAveragesFromLocs.nNumAntiMagic = RoundUp(tmp_nAvgNumAntiMagic / nLairs)
       
    Call RemoveOutliers(tmp_nAvgWalk)
    GetLairAveragesFromLocs.nAvgWalk = Round(CalcAverageNonZero(tmp_nAvgWalk), 1)
    GetLairAveragesFromLocs.nTotalLairs = nLairs
    
    If GetLairAveragesFromLocs.nMaxRegen < 1 Then GetLairAveragesFromLocs.nMaxRegen = 1
    
    GetLairAveragesFromLocs.nDamageOut = Round(tmp_nAvgDamageOut / nLairs)
    GetLairAveragesFromLocs.nMinDamageOut = Round(tmp_nMinDmgOut / nLairs)
    GetLairAveragesFromLocs.nSurpriseDamageOut = Round(tmp_nSurpriseDamageOut / nLairs)
    GetLairAveragesFromLocs.nPossSpawns = GetLairAveragesFromLocs.nPossSpawns + nLairs
    GetLairAveragesFromLocs.sGroupIndex = sLoc
    GetLairAveragesFromLocs.sGlobalAttackConfig = sGlobalAttackConfig
    GetLairAveragesFromLocs.sMobList = RemoveDuplicateNumbersFromString(tmp_sMobList)
    
    If GetLairAveragesFromLocs.nSpellImmuLVL > 0 Or GetLairAveragesFromLocs.nMagicLVL > 0 Or GetLairAveragesFromLocs.nNumUndeads > 0 Then
        nDmgOut = GetDamageOutput(0, GetLairAveragesFromLocs.nAvgAC, GetLairAveragesFromLocs.nAvgDR, GetLairAveragesFromLocs.nAvgMR, GetLairAveragesFromLocs.nAvgDodge, _
                        IIf(GetLairAveragesFromLocs.nNumAntiMagic >= (GetLairAveragesFromLocs.nMobs / 2), True, False), True, 100, _
                        GetLairAveragesFromLocs.nSpellImmuLVL, GetLairAveragesFromLocs.nMagicLVL, IIf(GetLairAveragesFromLocs.nNumUndeads >= (GetLairAveragesFromLocs.nMobs * LAIR_UNDEAD_RATIO), True, False))
        If nDmgOut(0) = -9998 Then GetLairAveragesFromLocs.nDamageOut = 0
        If nDmgOut(1) = -9998 Then GetLairAveragesFromLocs.nMinDamageOut = 0
        If nDmgOut(2) = -9998 Then GetLairAveragesFromLocs.nSurpriseDamageOut = 0
    End If
End If
out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetLairAveragesFromLocs")
Resume out:
End Function

Public Function GetLairInfoIndex(sGroupIndex As String) As Long
On Error GoTo error:
If Len(sGroupIndex) < 1 Then Exit Function

If dictLairInfo.Exists(sGroupIndex) Then
    GetLairInfoIndex = val(dictLairInfo.item(sGroupIndex))
Else
    GetLairInfoIndex = UBound(colLairs()) + 1
    ReDim Preserve colLairs(GetLairInfoIndex)
    dictLairInfo.Add sGroupIndex, GetLairInfoIndex
    colLairs(GetLairInfoIndex).sGroupIndex = sGroupIndex
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetLairInfoIndex")
Resume out:
End Function

Public Function GetLairInfo(ByVal sGroupIndex As String, Optional ByVal nMaxRegen As Integer) As LairInfoType
On Error GoTo error:
Dim x As Long, sArr() As String, nDamageOut As Long, nParty As Integer
Dim avgAlive As Double, nRTK As Double, nRTC As Double
Dim nDmgOut() As Currency, nMinDamageOut As Long
Dim nMinDmgPct As Double, nSurpriseDamageOut As Long, tCombatInfo As tCombatRoundInfo
If Len(sGroupIndex) < 5 Then Exit Function

If nMaxRegen = 0 Then
    sArr() = Split(sGroupIndex, "-", , vbTextCompare)
    If UBound(sArr()) < 3 Then Exit Function
    nMaxRegen = val(sArr(3))
End If

x = GetLairInfoIndex(sGroupIndex)

GetLairInfo.sGroupIndex = colLairs(x).sGroupIndex
GetLairInfo.sMobList = colLairs(x).sMobList
GetLairInfo.nMobs = colLairs(x).nMobs
GetLairInfo.nAvgExp = colLairs(x).nAvgExp
GetLairInfo.nAvgDmg = colLairs(x).nAvgDmg
GetLairInfo.nAvgHP = colLairs(x).nAvgHP
GetLairInfo.nAvgAC = colLairs(x).nAvgAC
GetLairInfo.nAvgDR = colLairs(x).nAvgDR
GetLairInfo.nAvgMR = colLairs(x).nAvgMR
GetLairInfo.nAvgDodge = colLairs(x).nAvgDodge
GetLairInfo.nDamageOut = colLairs(x).nDamageOut
GetLairInfo.nMinDamageOut = colLairs(x).nMinDamageOut
GetLairInfo.nSurpriseDamageOut = colLairs(x).nSurpriseDamageOut
GetLairInfo.sGlobalAttackConfig = colLairs(x).sGlobalAttackConfig
GetLairInfo.nMaxRegen = nMaxRegen
GetLairInfo.nAvgDelay = colLairs(x).nAvgDelay
GetLairInfo.nAvgWalk = colLairs(x).nAvgWalk
GetLairInfo.nTotalLairs = colLairs(x).nTotalLairs
GetLairInfo.nMagicLVL = colLairs(x).nMagicLVL
GetLairInfo.nMaxMagicLVL = colLairs(x).nMaxMagicLVL
GetLairInfo.nSpellImmuLVL = colLairs(x).nSpellImmuLVL
GetLairInfo.nMaxSpellImmuLVL = colLairs(x).nMaxSpellImmuLVL
GetLairInfo.nNumUndeads = colLairs(x).nNumUndeads
GetLairInfo.nNumAntiMagic = colLairs(x).nNumAntiMagic

GetLairInfo.nRTK = 1
GetLairInfo.nRTC = nMaxRegen
'GetLairInfo.nScriptValue = colLairs(x).nScriptValue
'GetLairInfo.nRestRate = colLairs(x).nRestRate
'GetLairInfo.nManaRecoveryRate = colLairs(x).nManaRecoveryRate
GetLairInfo.nDamageMitigated = 0

nRTK = 1
nRTC = nMaxRegen
nDamageOut = -9999
nMinDamageOut = -9999
nSurpriseDamageOut = -9999

If Len(GetLairInfo.sMobList) > 0 And Not bStartup Then
    nParty = 1
    If frmMain.optMonsterFilter(1).Value = True Then nParty = val(frmMain.txtMonsterLairFilter(0).Text)
    If nParty < 1 Then nParty = 1
    If nParty > 6 Then nParty = 6
    
    If nParty = 1 And Len(GetLairInfo.sGlobalAttackConfig) > 1 And GetLairInfo.sGlobalAttackConfig = sGlobalAttackConfig Then
        nDamageOut = GetLairInfo.nDamageOut
        nMinDamageOut = GetLairInfo.nMinDamageOut
        nSurpriseDamageOut = GetLairInfo.nSurpriseDamageOut
    Else
        nDmgOut = GetDamageOutput(0, GetLairInfo.nAvgAC, GetLairInfo.nAvgDR, GetLairInfo.nAvgMR, GetLairInfo.nAvgDodge, _
                        IIf(GetLairInfo.nNumAntiMagic >= (GetLairInfo.nMobs / 2), True, False), True, 100, _
                        GetLairInfo.nSpellImmuLVL, GetLairInfo.nMagicLVL, IIf(GetLairInfo.nNumUndeads >= (GetLairInfo.nMobs * LAIR_UNDEAD_RATIO), True, False))
        nDamageOut = nDmgOut(0)
        nMinDamageOut = nDmgOut(1)
        nSurpriseDamageOut = nDmgOut(2)
        If nDamageOut > -9999 Or nSurpriseDamageOut > -9999 Then
            GetLairInfo.nDamageOut = IIf(nDamageOut > -9990, nDamageOut, 0)
            GetLairInfo.nMinDamageOut = IIf(nMinDamageOut > -9990, nMinDamageOut, 0)
            GetLairInfo.nSurpriseDamageOut = IIf(nSurpriseDamageOut > -9990, nSurpriseDamageOut, 0)
            If nParty = 1 Then
                GetLairInfo.sGlobalAttackConfig = sGlobalAttackConfig
                Call SetLairInfo(GetLairInfo)
            End If
        Else
            nDamageOut = 9999999
        End If
    End If
    If nDamageOut <= -9990 Then nDamageOut = 0
    If nMinDamageOut <= -9990 Then nMinDamageOut = 0
    If nSurpriseDamageOut <= -9990 Then nSurpriseDamageOut = 0
    
    If frmMain.chkGlobalFilter.Value = 1 Or nParty > 1 Then 'vs char or vs party
        'GetLairInfo.nAvgDmg = 0
        sArr() = Split(GetLairInfo.sMobList, ",")
        For x = 0 To UBound(sArr())
            If val(sArr(x)) <= UBound(nMonsterDamageVsChar()) Then
                If nParty > 1 And nMonsterDamageVsParty(val(sArr(x))) >= 0 Then 'vs party
                    GetLairInfo.nDamageMitigated = GetLairInfo.nDamageMitigated + nMonsterDamageVsParty(val(sArr(x)))
                ElseIf nParty = 1 And frmMain.chkGlobalFilter.Value = 1 And nMonsterDamageVsChar(val(sArr(x))) >= 0 Then
                    GetLairInfo.nDamageMitigated = GetLairInfo.nDamageMitigated + nMonsterDamageVsChar(val(sArr(x)))
                ElseIf nMonsterDamageVsDefault(val(sArr(x))) >= 0 Then
                    GetLairInfo.nDamageMitigated = GetLairInfo.nDamageMitigated + nMonsterDamageVsDefault(val(sArr(x)))
                Else
                    GetLairInfo.nDamageMitigated = GetLairInfo.nDamageMitigated + GetMonsterAvgDmgFromDB(val(sArr(x)))
                End If
            End If
        Next x
        GetLairInfo.nDamageMitigated = Round(GetLairInfo.nDamageMitigated / (UBound(sArr()) + 1), 1)
    End If
    
    If GetLairInfo.nAvgDmg > 0 And GetLairInfo.nDamageMitigated <> GetLairInfo.nAvgDmg Then
        GetLairInfo.nDamageMitigated = GetLairInfo.nAvgDmg - GetLairInfo.nDamageMitigated
        GetLairInfo.nAvgDmg = GetLairInfo.nAvgDmg - GetLairInfo.nDamageMitigated
    Else
        GetLairInfo.nDamageMitigated = 0
    End If
    GetLairInfo.nAvgDmgLair = GetLairInfo.nAvgDmg

'/patch 2025.08.25
    If nDamageOut + nSurpriseDamageOut > 0 Then
        tCombatInfo = CalcCombatRounds(nDamageOut, GetLairInfo.nAvgHP, GetLairInfo.nAvgDmgLair, , , 1, , nSurpriseDamageOut, nMinDamageOut)
        nRTK = tCombatInfo.nRTK
        If nRTK < 1 Then nRTK = 1
        GetLairInfo.nRTK = nRTK
        If nRTK > 1 Then GetLairInfo.nAvgDmgLair = Round(GetLairInfo.nAvgDmgLair * nRTK, 1)
    End If
    
'    If nDamageOut > 0 And (nDamageOut < GetLairInfo.nAvgHP Or (nMinDamageOut > -9990 And nMinDamageOut < GetLairInfo.nAvgHP)) Then
'        nRTK = Round(GetLairInfo.nAvgHP / nDamageOut, 2)
'        If nRTK < 1 Then nRTK = 1
'
'        If nRTK = 1 And nMinDamageOut < nDamageOut And nMinDamageOut > -999 And nMinDamageOut < GetLairInfo.nAvgHP Then
'            nMinDmgPct = (GetLairInfo.nAvgHP - nMinDamageOut) / (nDamageOut - nMinDamageOut)
'            If nMinDmgPct >= 0.5 Then nRTK = 1.5
'        End If
'
'        If nRTK > 1 Then
'            nRTK = -Int(-(nRTK * 2)) / 2 'round up to nearest 0.5
'            'if the character/party damage output is less than the lair's mob's average HPs, increase their damage output
'            'e.g. if it takes 2 rounds to kill each mob, then their damage would be x2
'            GetLairInfo.nAvgDmgLair = Round(GetLairInfo.nAvgDmgLair * nRTK, 1)
'            'this damage increase is to account for per-mob in the lair
'            GetLairInfo.nRTK = nRTK
'        End If
'    End If
'/patch 2025.08.25

    If GetLairInfo.nMaxRegen > 1 And GetLairInfo.nAvgDmgLair > 0 Then
        'unless rooming or attacking different mobs, >1 mobs = more than one round to kill, even if damage out > all mob HP combined
        'this is to simulate increased damage from those extra rounds, but with less mobs per round as rounds progress
        avgAlive = (GetLairInfo.nMaxRegen + 1) / (2 * GetLairInfo.nMaxRegen)
        'avgAlive will product a decminal value. thus, the below divide will actually increase the damage
        GetLairInfo.nAvgDmgLair = Round(GetLairInfo.nAvgDmgLair / avgAlive, 1)
    End If
    
    If nDamageOut + nSurpriseDamageOut > 0 Then
        If GetLairInfo.nMaxRegen > 1 Then
            nRTC = nRTK * GetLairInfo.nMaxRegen
        Else
            nRTC = nRTK
        End If
        GetLairInfo.nRTC = nRTC
    End If
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetLairInfo")
Resume out:
End Function

Public Sub SetLairInfo(tUpdatedLairInfo As LairInfoType) ', Optional bSVspecified As Boolean = False
On Error GoTo error:
Dim x As Long ', sArr() As String, i As Integer

If Len(tUpdatedLairInfo.sGroupIndex) < 5 Then Exit Sub
x = GetLairInfoIndex(tUpdatedLairInfo.sGroupIndex)

'averages are for a single mob, averaged from all of the mobs in the list/index
'the max regen is not taken into account (e.g. multiply .nMaxRegen by .nAvgExp for total average exp for the lair)
colLairs(x).sMobList = tUpdatedLairInfo.sMobList
colLairs(x).nMobs = tUpdatedLairInfo.nMobs
colLairs(x).nAvgExp = tUpdatedLairInfo.nAvgExp
colLairs(x).nAvgDmg = tUpdatedLairInfo.nAvgDmg
colLairs(x).nAvgHP = tUpdatedLairInfo.nAvgHP
colLairs(x).nAvgAC = tUpdatedLairInfo.nAvgAC
colLairs(x).nAvgDR = tUpdatedLairInfo.nAvgDR
colLairs(x).nAvgMR = tUpdatedLairInfo.nAvgMR
colLairs(x).nAvgDodge = tUpdatedLairInfo.nAvgDodge
colLairs(x).nAvgDelay = tUpdatedLairInfo.nAvgDelay
colLairs(x).nAvgWalk = tUpdatedLairInfo.nAvgWalk
colLairs(x).nTotalLairs = tUpdatedLairInfo.nTotalLairs
colLairs(x).nMagicLVL = tUpdatedLairInfo.nMagicLVL
colLairs(x).nMaxMagicLVL = tUpdatedLairInfo.nMaxMagicLVL
colLairs(x).nSpellImmuLVL = tUpdatedLairInfo.nSpellImmuLVL
colLairs(x).nMaxSpellImmuLVL = tUpdatedLairInfo.nMaxSpellImmuLVL
colLairs(x).nNumUndeads = tUpdatedLairInfo.nNumUndeads
colLairs(x).nNumAntiMagic = tUpdatedLairInfo.nNumAntiMagic

If Not tUpdatedLairInfo.sGlobalAttackConfig = "" Then
    colLairs(x).nDamageOut = tUpdatedLairInfo.nDamageOut
    colLairs(x).nMinDamageOut = tUpdatedLairInfo.nMinDamageOut
    colLairs(x).nSurpriseDamageOut = tUpdatedLairInfo.nSurpriseDamageOut
    colLairs(x).sGlobalAttackConfig = tUpdatedLairInfo.sGlobalAttackConfig
End If
If tUpdatedLairInfo.nMaxRegen > 0 Then colLairs(x).nMaxRegen = tUpdatedLairInfo.nMaxRegen

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("SetLairInfo")
Resume out:
End Sub

Public Function CalcExpNeededByRaceClass(ByVal nLevel As Long, ByVal nClass As Long, ByVal nRace As Long) As Currency
Dim nClassExp As Integer, nRaceExp As Integer, nExp As Currency, nChart As Long

On Error GoTo error:

If nClass > 0 Then
    tabClasses.Index = "pkClasses"
    tabClasses.Seek "=", nClass
    If tabClasses.NoMatch = True Then
        nClassExp = 0
        tabClasses.MoveFirst
    Else
        nClassExp = tabClasses.Fields("ExpTable") + 100
    End If
End If

If nRace > 0 Then
    tabRaces.Index = "pkRaces"
    tabRaces.Seek "=", nRace
    If tabRaces.NoMatch = True Then
        nRaceExp = 0
        tabRaces.MoveFirst
    Else
        nRaceExp = tabRaces.Fields("ExpTable")
    End If
End If

nChart = nClassExp + nRaceExp
nExp = CalcExpNeeded(nLevel, nChart)
CalcExpNeededByRaceClass = nExp

Exit Function
error:
Call HandleError("CalcExpNeededByRaceClass")

End Function
Public Function OpenTables(sFile As String) As Boolean
On Error GoTo error:
Dim nMaxMon As Long

UseExpMulti = False

'Set WS = DAO.CreateWorkspace("MMUD_Explorer_WS", "MMUD_Explorer", False, dbUseJet)
Set DB = OpenDatabase(sFile, False, True)

sCurrentDatabaseFile = sFile

Set tabItems = DB.OpenRecordset("Items")
Set tabClasses = DB.OpenRecordset("Classes")
Set tabRaces = DB.OpenRecordset("Races")
Set tabSpells = DB.OpenRecordset("Spells")
Set tabMonsters = DB.OpenRecordset("Monsters")
Set tabShops = DB.OpenRecordset("Shops")
Set tabRooms = DB.OpenRecordset("Rooms")
Set tabInfo = DB.OpenRecordset("Info")
Set tabTBInfo = DB.OpenRecordset("TBInfo")

Call TestMonExpMulti

If tabMonsters.RecordCount > 0 Then
    tabMonsters.MoveLast
    nMaxMon = tabMonsters.Fields("Number")
    ReDim nMonsterDamageVsChar(nMaxMon)
    ReDim nMonsterPossy(nMaxMon)
    ReDim nMonsterSpawnChance(nMaxMon)
    ReDim nCharDamageVsMonster(nMaxMon)
    ReDim nCharMinDamageVsMonster(nMaxMon)
    ReDim nCharSurpriseDamageVsMonster(nMaxMon)
    ReDim nMonsterDamageVsDefault(nMaxMon)
    ReDim nMonsterDamageVsParty(nMaxMon)
End If

OpenTables = True

Exit Function
error:
Call HandleError("OpenDatabase")
'Resume Next
End Function

'==============================================================================
' LoadLairInfo – majority (mode) for Magic / Spell Immunity levels
'  • nMagicLVL / nSpellImmuLVL = most common level across mobs in the lair
'  • Includes 0 (“no level”) in the competition
'  • Tie-breaker prefers the HIGHER level
'  • nMaxMagicLVL / nMaxSpellImmuLVL still track maxima
'==============================================================================

Public Sub LoadLairInfo()
On Error GoTo error:

    '---------------------------
    ' Declarations (all at top)
    '---------------------------
    Dim tLairInfo As LairInfoType
    Dim sGroupIndex As String
    Dim sArr() As String
    Dim x As Integer
    Dim y As Integer
    Dim nTemp As Long

    Dim dictMagicCounts As Scripting.Dictionary
    Dim dictSpellCounts As Scripting.Dictionary

    Dim lvlMag As Long
    Dim lvlImm As Long
    Dim hadMagField As Boolean
    Dim hadImmField As Boolean

    ' Zero-able per-lair accumulators/flags
    Dim zNumUndeads As Long
    Dim zNumAntiMagic As Long
    Dim zMagicLVL As Long
    Dim zMaxMagicLVL As Long
    Dim zSpellImmuLVL As Long
    Dim zMaxSpellImmuLVL As Long

    '---------------------------
    ' Setup overall structures
    '---------------------------
    Set dictLairInfo = Nothing
    Set dictLairInfo = New Dictionary
    dictLairInfo.CompareMode = vbTextCompare
    ReDim colLairs(0)

    If tabLairs Is Nothing Then Set tabLairs = DB.OpenRecordset("Lairs")
    If tabLairs.RecordCount = 0 Then Exit Sub

    tabLairs.MoveFirst
    Do While Not tabLairs.EOF

        '---------------------------
        ' Reset / zero per-lair vars
        '---------------------------
        sGroupIndex = tabLairs.Fields("GroupIndex")
        tLairInfo = GetLairInfo("")
        tLairInfo.sGroupIndex = sGroupIndex

        tLairInfo.sMobList = tabLairs.Fields("MobList")
        tLairInfo.nMobs = tabLairs.Fields("Mobs")
        tLairInfo.nAvgDelay = tabLairs.Fields("AvgDelay")
        tLairInfo.nAvgExp = tabLairs.Fields("AvgExp")
        tLairInfo.nAvgDmg = tabLairs.Fields("AvgDmg")
        tLairInfo.nAvgHP = tabLairs.Fields("AvgHP")
        tLairInfo.nAvgAC = tabLairs.Fields("AvgAC")
        tLairInfo.nAvgDR = tabLairs.Fields("AvgDR")
        tLairInfo.nAvgMR = tabLairs.Fields("AvgMR")
        tLairInfo.nAvgDodge = tabLairs.Fields("AvgDodge")
        tLairInfo.nAvgWalk = tabLairs.Fields("AvgWalk")
        tLairInfo.nTotalLairs = tabLairs.Fields("TotalLairs")

        ' Explicit zeroing of counters that we (re)compute
        zNumUndeads = 0
        zNumAntiMagic = 0
        zMagicLVL = 0
        zMaxMagicLVL = 0
        zSpellImmuLVL = 0
        zMaxSpellImmuLVL = 0

        ' fresh per-lair dictionaries
        Set dictMagicCounts = New Scripting.Dictionary
        Set dictSpellCounts = New Scripting.Dictionary

        If tLairInfo.nMobs > 0 And tabMonsters.RecordCount > 0 Then
            sArr = Split(tLairInfo.sMobList, ",")

            For x = 0 To UBound(sArr)
                nTemp = val(sArr(x))
                If nTemp > 0 Then
                    tabMonsters.Index = "pkMonsters"
                    tabMonsters.Seek "=", nTemp
                    If tabMonsters.NoMatch = False Then

                        ' track undead count
                        If tabMonsters.Fields("Undead") = 1 Then
                            zNumUndeads = zNumUndeads + 1
                        End If

                        ' per-mob flags to detect missing ability fields (so we can count 0)
                        hadMagField = False
                        hadImmField = False

                        ' scan ability slots
                        For y = 0 To 9
                            If Not tabMonsters.Fields("Abil-" & y) = 0 Then
                                Select Case tabMonsters.Fields("Abil-" & y)

                                    Case 28      ' magical level
                                        hadMagField = True
                                        lvlMag = CLng(tabMonsters.Fields("AbilVal-" & y)) ' may be 0+
                                        Call LoadLairInfo_BumpCount(dictMagicCounts, lvlMag)
                                        If lvlMag > zMaxMagicLVL Then zMaxMagicLVL = lvlMag

                                    Case 51      ' anti-magic flag
                                        zNumAntiMagic = zNumAntiMagic + 1

                                    Case 139     ' spell immunity level
                                        hadImmField = True
                                        lvlImm = CLng(tabMonsters.Fields("AbilVal-" & y)) ' may be 0+
                                        Call LoadLairInfo_BumpCount(dictSpellCounts, lvlImm)
                                        If lvlImm > zMaxSpellImmuLVL Then zMaxSpellImmuLVL = lvlImm

                                End Select
                            End If
                        Next y

                        ' If no magical ability field was present at all, that mob is level 0
                        If Not hadMagField Then
                            Call LoadLairInfo_BumpCount(dictMagicCounts, 0&)
                        End If
                        ' If no spell-immune field was present at all, that mob is level 0
                        If Not hadImmField Then
                            Call LoadLairInfo_BumpCount(dictSpellCounts, 0&)
                        End If

                    End If
                End If
            Next x

            tabMonsters.MoveFirst

            ' Majority (mode); includes 0 and prefers HIGHER level on ties
            If dictMagicCounts.Count > 0 Then
                zMagicLVL = LoadLairInfo_ModeFromCounts(dictMagicCounts, 0&)
            Else
                zMagicLVL = 0&
            End If

            If dictSpellCounts.Count > 0 Then
                zSpellImmuLVL = LoadLairInfo_ModeFromCounts(dictSpellCounts, 0&)
            Else
                zSpellImmuLVL = 0&
            End If

        End If

        ' write back the zeroed/derived fields
        tLairInfo.nNumUndeads = zNumUndeads
        tLairInfo.nNumAntiMagic = zNumAntiMagic

        tLairInfo.nMagicLVL = zMagicLVL
        tLairInfo.nMaxMagicLVL = zMaxMagicLVL

        tLairInfo.nSpellImmuLVL = zSpellImmuLVL
        tLairInfo.nMaxSpellImmuLVL = zMaxSpellImmuLVL

        ' store
        Call SetLairInfo(tLairInfo)

        ' next lair
        tabLairs.MoveNext

        '------------------------------------------
        ' CLEAR / zero between lairs (defensive)
        '------------------------------------------
        Set dictMagicCounts = Nothing
        Set dictSpellCounts = Nothing
        zNumUndeads = 0: zNumAntiMagic = 0
        zMagicLVL = 0: zMaxMagicLVL = 0
        zSpellImmuLVL = 0: zMaxSpellImmuLVL = 0

    Loop

out:
    On Error Resume Next
    Exit Sub

error:
    Call HandleError("LoadLairInfo")
    Resume out

End Sub


' Returns the key (level) with the highest count in a Dictionary(Long->Long).
' Includes 0 as a valid competitor.
' Tie-breaker: prefers the HIGHER level.
Private Function LoadLairInfo_ModeFromCounts(ByVal dict As Scripting.Dictionary, Optional ByVal defaultLevel As Long = 0) As Long
    Dim k As Variant
    Dim bestLevel As Long
    Dim bestCount As Long
    Dim curLevel As Long
    Dim curCount As Long

    bestLevel = defaultLevel
    bestCount = -1

    For Each k In dict.Keys
        curLevel = CLng(k)
        curCount = CLng(dict(k))
        If curCount > bestCount Then
            bestCount = curCount
            bestLevel = curLevel
        ElseIf curCount = bestCount Then
            ' Prefer HIGHER level on ties
            If curLevel > bestLevel Then bestLevel = curLevel
        End If
    Next k

    LoadLairInfo_ModeFromCounts = bestLevel
End Function

' Increment a numeric bucket in a Dictionary(Long->Long). Creates the key if missing.
Private Sub LoadLairInfo_BumpCount(ByVal dict As Scripting.Dictionary, ByVal level As Long)
    If dict.Exists(level) Then
        dict(level) = CLng(dict(level)) + 1&
    Else
        dict.Add level, 1&
    End If
End Sub

Public Sub CalculateAverageLairs()
Dim sGroupIndex As String, sRoomKey As String ', nMapRoom As RoomExitType
Dim iLair As Integer, nLairs As Long, nMaxRegen As Currency, nMobsTotal As Currency, nSpawnChance As Currency
Dim sRegexPattern As String, tMatches() As RegexMatches, tLairInfo As LairInfoType
On Error GoTo error:

Set tabTempRS = DB.OpenRecordset( _
    "SELECT [Number],[Summoned By] FROM Monsters WHERE [Summoned By] Like ""*(lair)*""", dbOpenSnapshot)

sRegexPattern = "Group\(lair\): (\d+)\/(\d+)"
If nNMRVer >= 1.82 Then sRegexPattern = "\[(\d+)\]" & sRegexPattern
If nNMRVer >= 1.83 Then sRegexPattern = "\[([\d\-]+)\]" & sRegexPattern

If Not tabTempRS.EOF Then
    tabTempRS.MoveFirst

    Do While Not tabTempRS.EOF
        nMobsTotal = 0
        nLairs = 0
        nSpawnChance = 0
        
        tMatches() = RegExpFindv2(tabTempRS.Fields("Summoned By"), sRegexPattern)
        If UBound(tMatches()) > 0 Or Len(tMatches(0).sFullMatch) > 0 Then
            nLairs = UBound(tMatches()) + 1
            
            For iLair = 0 To UBound(tMatches())
                nMaxRegen = 0
                sGroupIndex = "0-0-0"
                
                If nNMRVer >= 1.83 Then
                    '[7-8-9][6]Group(lair): 1/2345
                    sGroupIndex = tMatches(iLair).sSubMatches(0)
                    nMaxRegen = val(tMatches(iLair).sSubMatches(1))
                    sRoomKey = tMatches(iLair).sSubMatches(2) & "/" & tMatches(iLair).sSubMatches(3)
                    If nMaxRegen > 0 Then
                        tLairInfo = GetLairInfo(sGroupIndex, nMaxRegen)
                        If tLairInfo.nMobs > 0 Then
                            nSpawnChance = nSpawnChance + Round(1 - (1 - (1 / tLairInfo.nMobs)) ^ nMaxRegen, 2)
                            '1 - (1 - (x / y)) ^ z
                            '(x / y) == (1) of (y) totalmobs
                            'z == maxregen (chance to spawn)
                        End If
                    End If
                ElseIf nNMRVer >= 1.82 Then
                    '[6]Group(lair): 1/2345
                    nMaxRegen = val(tMatches(iLair).sSubMatches(0))
                    sRoomKey = tMatches(iLair).sSubMatches(1) & "/" & tMatches(iLair).sSubMatches(2)
                Else
                    'Group(lair): 1/2345
                    nMaxRegen = 1
                    sRoomKey = tMatches(iLair).sSubMatches(0) & "/" & tMatches(iLair).sSubMatches(1)
                End If
                
                nMobsTotal = nMobsTotal + nMaxRegen
            Next iLair
            
            If nMobsTotal > 0 Then
                nMonsterPossy(tabTempRS.Fields("Number")) = Round(nMobsTotal / nLairs, 1)
            End If
            
            If nNMRVer >= 1.83 Then
                nMonsterSpawnChance(tabTempRS.Fields("Number")) = Round(nSpawnChance / nLairs, 2)
            End If
        End If
        
        tabTempRS.MoveNext
    Loop
    
    tabTempRS.MoveLast
    tabTempRS.Close
    Set tabTempRS = Nothing
End If

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("CalculateAverageLairs")
Resume out:
End Sub

Private Sub TestMonExpMulti()
On Error GoTo nomulti:
Dim sTest As String

'this just tests the exp multiplier field. if it exists, it wont error out

sTest = tabMonsters.Fields("ExpMulti").Value

UseExpMulti = True

Exit Sub
nomulti:
Err.clear
End Sub

Public Sub CloseDatabases()
On Error Resume Next

tabItems.Close
tabSpells.Close
tabRaces.Close
tabClasses.Close
tabInfo.Close
tabMonsters.Close
tabShops.Close
tabRooms.Close
tabTBInfo.Close

If nNMRVer >= 1.83 Then
    tabLairs.Close
End If

Set tabRooms = Nothing
Set tabMonsters = Nothing
Set tabShops = Nothing
Set tabItems = Nothing
Set tabSpells = Nothing
Set tabRaces = Nothing
Set tabClasses = Nothing
Set tabInfo = Nothing
Set tabTBInfo = Nothing
Set tabLairs = Nothing

DB.Close
'WS.Close

Set DB = Nothing
'Set WS = Nothing

End Sub


Public Function GetShopName(ByVal nNum As Long, Optional ByVal bNoNumber As Boolean) As String
On Error GoTo error:

If nNum = 0 Then GetShopName = "None": Exit Function
GetShopName = nNum
If tabShops.RecordCount = 0 Then Exit Function

On Error GoTo seek2:
If tabShops.Fields("Number") = nNum Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabShops.Index = "pkShops"
tabShops.Seek "=", nNum
If tabShops.NoMatch = True Then
    tabShops.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:

GetShopName = tabShops.Fields("Name")
If Not bNoNumber Then GetShopName = GetShopName & "(" & nNum & ")"

out:
Exit Function
error:
Call HandleError("GetShopName")
Resume out:
End Function

Public Function GetItemShopRegenPCT(ByVal nShopNum As Long, ByVal nItemNum As Long) As Currency
Dim nRegenTimeMultiplier As Currency, x As Integer
On Error GoTo error:

GetItemShopRegenPCT = 0
If nItemNum < 1 Or nShopNum < 1 Then Exit Function

tabShops.Index = "pkShops"
tabShops.Seek "=", nShopNum
If tabShops.NoMatch = True Then
    tabShops.MoveFirst
    Exit Function
End If

If tabShops.Fields("ShopType") = 8 Then Exit Function
    
For x = 0 To 19
    If tabShops.Fields("Item-" & x) = nItemNum And tabShops.Fields("Max-" & x) > 0 Then
    
        If tabShops.Fields("Time-" & x) > 0 And tabShops.Fields("%-" & x) > 0 And tabShops.Fields("Amount-" & x) > 0 Then
            nRegenTimeMultiplier = 1440 / tabShops.Fields("Time-" & x)
            GetItemShopRegenPCT = GetItemShopRegenPCT + (tabShops.Fields("Amount-" & x) * nRegenTimeMultiplier * (tabShops.Fields("%-" & x) / 100))
        Else
            'stock only, we'll give it a 1% chance
            GetItemShopRegenPCT = GetItemShopRegenPCT + 1
        End If
    End If
Next

out:
On Error Resume Next
If GetItemShopRegenPCT > 99 Then GetItemShopRegenPCT = 99
GetItemShopRegenPCT = Round(GetItemShopRegenPCT, 2)
Exit Function
error:
Call HandleError("GetItemShopRegenPCT")
Resume out:
End Function

Public Function GetSpellName(ByVal nNum As Long, Optional ByVal bNoNumber As Boolean) As String
On Error GoTo error:

If nNum = 0 Then GetSpellName = "None": Exit Function
GetSpellName = nNum
If tabSpells.RecordCount = 0 Then Exit Function

On Error GoTo seek2:
If tabSpells.Fields("Number") = nNum Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabSpells.Index = "pkSpells"
tabSpells.Seek "=", nNum
If tabSpells.NoMatch = True Then
    tabSpells.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:
GetSpellName = tabSpells.Fields("Name")
If Not bNoNumber Then GetSpellName = GetSpellName & "(" & nNum & ")"

out:
Exit Function
error:
Call HandleError("GetSpellName")
Resume out:
End Function

Public Function GetSpellManaCost(ByVal nNum As Long) As Long
On Error GoTo error:

If nNum = 0 Then Exit Function
If tabSpells.RecordCount = 0 Then Exit Function

On Error GoTo seek2:
If tabSpells.Fields("Number") = nNum Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabSpells.Index = "pkSpells"
tabSpells.Seek "=", nNum
If tabSpells.NoMatch = True Then
    tabSpells.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:

GetSpellManaCost = tabSpells.Fields("ManaCost")
If tabSpells.Fields("EnergyCost") > 0 And tabSpells.Fields("EnergyCost") <= 500 Then
    GetSpellManaCost = GetSpellManaCost * Fix(1000 / tabSpells.Fields("EnergyCost"))
End If

out:
Exit Function
error:
Call HandleError("GetSpellManaCost")
Resume out:
End Function

Public Function GetSpellShort(ByVal nNum As Long) As String
On Error GoTo error:
GetSpellShort = "n/a"

If nNum = 0 Then Exit Function
If tabSpells.RecordCount = 0 Then Exit Function

On Error GoTo seek2:
If tabSpells.Fields("Number") = nNum Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabSpells.Index = "pkSpells"
tabSpells.Seek "=", nNum
If tabSpells.NoMatch = True Then
    tabSpells.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:
GetSpellShort = tabSpells.Fields("Short")

out:
Exit Function
error:
Call HandleError("GetSpellShort")
Resume out:
End Function

Public Function GetSpellByShort(ByVal sFindShort As String, Optional ByVal nLearnableByClass As Long) As Long
On Error GoTo error:
Dim nMagery As Integer, nMageryLevel As Integer

If nLearnableByClass > 0 Then
    nMagery = GetClassMagery(nLearnableByClass)
    If nMagery = 0 Then Exit Function
    nMageryLevel = GetClassMageryLVL(nLearnableByClass)
End If

sFindShort = Trim(sFindShort)
If sFindShort = "" Then Exit Function

If tabSpells.RecordCount = 0 Then Exit Function
tabSpells.MoveFirst
Do Until tabSpells.EOF

    If Trim(tabSpells.Fields("Short")) <> sFindShort Then GoTo skip_spell:
    
    If tabSpells.Fields("Learnable") = 0 And Len(tabSpells.Fields("Learned From")) <= 1 And Len(tabSpells.Fields("Casted By")) <= 1 _
        And (tabSpells.Fields("Magery") <> 5 Or (tabSpells.Fields("Magery") = 5 And tabSpells.Fields("ReqLevel") < 1)) Then
        If nNMRVer >= 1.8 Then
            If Len(tabSpells.Fields("Classes")) <= 1 Then GoTo skip_spell:
        Else
            GoTo skip_spell:
        End If
    End If
    
    If nMagery > 0 And Not nMagery = tabSpells.Fields("Magery") Then
        If tabSpells.Fields("Learnable") > 0 _
            And tabSpells.Fields("Magery") = 0 _
            And nNMRVer >= 1.7 Then
            
            If nLearnableByClass = 0 _
                Or tabSpells.Fields("Classes") = "(*)" _
                Or InStr(1, tabSpells.Fields("Classes"), _
                    "(" & nLearnableByClass & ")", vbTextCompare) > 0 Then
                GoTo skip_magery_check:
            Else
                GoTo skip_spell:
            End If
        Else
            GoTo skip_spell:
        End If
    End If
    
    If Not nMagery = 0 Then
        If nMageryLevel < tabSpells.Fields("MageryLVL") Then GoTo skip_spell:
    End If

    'magery 5 is kai
    If Not nMagery = 5 And tabSpells.Fields("Learnable") = 0 Then GoTo skip_spell:
    
skip_magery_check:

    GetSpellByShort = tabSpells.Fields("Number")
    Exit Do
    
skip_spell:
    tabSpells.MoveNext
Loop

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetSpellByShort")
Resume out:
End Function
Public Function GetRaceHPBonus(ByVal nNum As Long) As Integer
On Error GoTo error:

If nNum = 0 Then GetRaceHPBonus = 0: Exit Function
If tabRaces.RecordCount = 0 Then GetRaceHPBonus = 0: Exit Function

tabRaces.Index = "pkRaces"
tabRaces.Seek "=", nNum
If tabRaces.NoMatch = True Then
    GetRaceHPBonus = 0
    tabRaces.MoveFirst
Else
    GetRaceHPBonus = tabRaces.Fields("HPPerLVL")
End If

Exit Function
error:
Call HandleError("GetRaceHPBonus")
GetRaceHPBonus = 0
End Function

Public Function GetClassMaxHP(ByVal nNum As Long) As Integer
On Error GoTo error:

If nNum = 0 Then GetClassMaxHP = 0: Exit Function
If tabClasses.RecordCount = 0 Then GetClassMaxHP = 0: Exit Function

tabClasses.Index = "pkClasses"
tabClasses.Seek "=", nNum
If tabClasses.NoMatch = True Then
    GetClassMaxHP = 0
    tabClasses.MoveFirst
Else
    GetClassMaxHP = tabClasses.Fields("MinHits") + tabClasses.Fields("MaxHits")
End If

Exit Function
error:
Call HandleError("GetClassMaxHP")
GetClassMaxHP = 0
End Function

Public Function GetClassMinHP(ByVal nNum As Long) As Integer
On Error GoTo error:

If nNum = 0 Then GetClassMinHP = 0: Exit Function
If tabClasses.RecordCount = 0 Then GetClassMinHP = 0: Exit Function

tabClasses.Index = "pkClasses"
tabClasses.Seek "=", nNum
If tabClasses.NoMatch = True Then
    GetClassMinHP = 0
    tabClasses.MoveFirst
Else
    GetClassMinHP = tabClasses.Fields("MinHits")
End If

Exit Function
error:
Call HandleError("GetClassMinHP")
GetClassMinHP = 0
End Function

Public Function GetClassName(ByVal nNum As Long) As String
On Error GoTo error:

If nNum = 0 Then GetClassName = "None": Exit Function
If tabClasses.RecordCount = 0 Then GetClassName = nNum: Exit Function

tabClasses.Index = "pkClasses"
tabClasses.Seek "=", nNum
If tabClasses.NoMatch = True Then
    GetClassName = nNum
    tabClasses.MoveFirst
Else
    GetClassName = tabClasses.Fields("Name")
End If

Exit Function
error:
Call HandleError("GetClassName")
GetClassName = nNum
End Function

Public Function GetClassMageryLVL(ByVal nNum As Long) As Integer

If nNum = 0 Then GetClassMageryLVL = 0: Exit Function
If tabClasses.RecordCount = 0 Then GetClassMageryLVL = 0: Exit Function

tabClasses.Index = "pkClasses"
tabClasses.Seek "=", nNum
If tabClasses.NoMatch = True Then
    GetClassMageryLVL = 0
    tabClasses.MoveFirst
Else
    GetClassMageryLVL = tabClasses.Fields("MageryLVL")
End If

Exit Function
error:
Call HandleError("GetClassMageryLVL")
GetClassMageryLVL = 0
End Function

Public Function GetClassMagery(ByVal nNum As Long) As enmMagicEnum

If nNum = 0 Then GetClassMagery = None: Exit Function
If tabClasses.RecordCount = 0 Then GetClassMagery = None: Exit Function

tabClasses.Index = "pkClasses"
tabClasses.Seek "=", nNum
If tabClasses.NoMatch = True Then
    GetClassMagery = None
    tabClasses.MoveFirst
Else
    Select Case tabClasses.Fields("MageryType")
        Case 1:
            GetClassMagery = Mage
        Case 2:
            GetClassMagery = Priest
        Case 3:
            GetClassMagery = Druid
        Case 4:
            GetClassMagery = Bard
        Case 5:
            GetClassMagery = Kai
        Case Else:
            GetClassMagery = None
    End Select
End If

Exit Function
error:
Call HandleError("GetClassMagery")
GetClassMagery = None
End Function

Public Function GetClassCombat(ByVal nNum As Long) As Integer
On Error GoTo error:
GetClassCombat = 1
If nNum = 0 Then Exit Function
If tabClasses.RecordCount = 0 Then Exit Function

On Error GoTo seek2:
If tabClasses.Fields("Number") = nNum Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabClasses.Index = "pkClasses"
tabClasses.Seek "=", nNum
If tabClasses.NoMatch = True Then
    tabClasses.MoveFirst
    Exit Function
End If

ready:
GetClassCombat = tabClasses.Fields("CombatLVL") - 2

Exit Function
error:
Call HandleError("GetClassCombat")
GetClassCombat = 1
End Function

Public Function GetRaceName(ByVal nNum As Long) As String
On Error GoTo error:

If nNum = 0 Then GetRaceName = "None": Exit Function
If tabRaces.RecordCount = 0 Then GetRaceName = nNum: Exit Function

tabRaces.Index = "pkRaces"
tabRaces.Seek "=", nNum
If tabRaces.NoMatch = True Then
    GetRaceName = nNum
    tabRaces.MoveFirst
Else
    GetRaceName = tabRaces.Fields("Name")
End If

Exit Function
error:
Call HandleError("GetRaceName")
GetRaceName = nNum
End Function

Public Function GetRaceCP(ByVal nNum As Long) As Integer
On Error GoTo error:

If nNum = 0 Then GetRaceCP = 100: Exit Function
If tabRaces.RecordCount = 0 Then GetRaceCP = 100: Exit Function

tabRaces.Index = "pkRaces"
tabRaces.Seek "=", nNum
If tabRaces.NoMatch = True Then
    GetRaceCP = 100
    tabRaces.MoveFirst
Else
    GetRaceCP = tabRaces.Fields("BaseCP")
End If

Exit Function
error:
Call HandleError("GetRaceCP")
GetRaceCP = 100
End Function

Public Function GetRaceStealth(Optional ByVal nNum As Long) As Boolean
Dim x As Integer
On Error GoTo error:

If tabRaces.RecordCount = 0 Then Exit Function
If nNum = 0 Then
    If frmMain.chkGlobalFilter.Value = 1 And frmMain.cmbGlobalClass(0).ListCount > 0 Then nNum = frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex)
    If nNum <= 0 Then Exit Function
End If

tabRaces.Index = "pkRaces"
tabRaces.Seek "=", nNum
If tabRaces.NoMatch = True Then
    tabRaces.MoveFirst
    Exit Function
End If

For x = 0 To 9
    If tabRaces.Fields("Abil-" & x) = 102 Then
        GetRaceStealth = True
        Exit For
    End If
Next x

Exit Function
error:
Call HandleError("GetRaceStealth")
End Function

Public Function GetClassStealth(Optional ByVal nNum As Long) As Boolean
Dim x As Integer
On Error GoTo error:

If tabClasses.RecordCount = 0 Then Exit Function
If nNum = 0 Then
    If frmMain.chkGlobalFilter.Value = 1 And frmMain.cmbGlobalRace(0).ListCount > 0 Then nNum = frmMain.cmbGlobalRace(0).ItemData(frmMain.cmbGlobalRace(0).ListIndex)
    If nNum = 0 Then Exit Function
End If

tabClasses.Index = "pkClasses"
tabClasses.Seek "=", nNum
If tabClasses.NoMatch = True Then
    tabClasses.MoveFirst
    Exit Function
End If

For x = 0 To 9
    If tabClasses.Fields("Abil-" & x) = 103 Then
        GetClassStealth = True
        Exit For
    End If
Next x

Exit Function
error:
Call HandleError("GetClassStealth")
End Function

Public Function GetMultiMonsterNames(ByVal sNumbers As String, ByVal HideNumber As Boolean) As String
Dim x As Long, y As Long
On Error GoTo error:

If sNumbers = "" Then GetMultiMonsterNames = "None": Exit Function
If tabMonsters.RecordCount = 0 Then Exit Function

tabMonsters.Index = "pkMonsters"
x = 0
Do While Not InStr(x + 1, sNumbers, ",") = 0
    y = InStr(x + 1, sNumbers, ",")
    
    tabMonsters.Seek "=", val(Mid(sNumbers, x + 1, y - x - 1))
    If tabMonsters.NoMatch = False Then
        GetMultiMonsterNames = GetMultiMonsterNames & IIf(GetMultiMonsterNames = "", "", ", ") _
            & tabMonsters.Fields("Name")
            
        If Not HideNumber Then
            GetMultiMonsterNames = GetMultiMonsterNames & "(" & tabMonsters.Fields("Number") & ")"
        End If
    End If
    x = y
Loop

Exit Function
error:
Call HandleError("GetMultiMonsterNames")
GetMultiMonsterNames = sNumbers
End Function
Public Function GetMonsterName(ByVal nNum As Long, ByVal bNoNumber As Boolean) As String
On Error GoTo error:
GetMonsterName = nNum

If nNum = 0 Then GetMonsterName = "None": Exit Function
GetMonsterName = nNum
If tabMonsters.RecordCount = 0 Then Exit Function

On Error GoTo seek2:
If tabMonsters.Fields("Number") = nNum Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabMonsters.Index = "pkMonsters"
tabMonsters.Seek "=", nNum
If tabMonsters.NoMatch = True Then
    tabMonsters.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:
GetMonsterName = tabMonsters.Fields("Name")
If Not bNoNumber Then GetMonsterName = GetMonsterName & "(" & nNum & ")"

Exit Function
error:
Call HandleError("GetMonsterName")
End Function

Public Function GetMonsterAvgDmgFromDB(ByVal nNum As Long) As Long
On Error GoTo error:
Dim nLocalMonsterDamage As MonAttackSimReturn

If nNum = 0 Then Exit Function
If tabMonsters.RecordCount = 0 Then Exit Function

If nNMRVer >= 1.8 Then
    tabMonsters.Index = "pkMonsters"
    tabMonsters.Seek "=", nNum
    If tabMonsters.NoMatch = True Then
        tabMonsters.MoveFirst
        Exit Function
    End If
    GetMonsterAvgDmgFromDB = tabMonsters.Fields("AvgDmg")
Else
    nLocalMonsterDamage = CalculateMonsterAvgDmg(nNum, nGlobalMonsterSimRounds)
    GetMonsterAvgDmgFromDB = nLocalMonsterDamage.nAverageDamage
End If

nMonsterDamageVsDefault(nNum) = GetMonsterAvgDmgFromDB

Exit Function
error:
Call HandleError("GetMonsterAvgDmgFromDB")
End Function

Public Function GetRoomName(Optional ByVal sMapRoom As String, Optional ByVal nMap As Long, _
    Optional ByVal nRoom As Long, Optional bNoRoomNumber As Boolean) As String
On Error GoTo error:
Dim tExit As RoomExitType, sName As String

If sMapRoom = "" Then
    tExit.Map = nMap
    tExit.Room = nRoom
Else
    tExit = ExtractMapRoom(sMapRoom)
End If

If tExit.Map = 0 Or tExit.Room = 0 Then GetRoomName = "?": Exit Function

On Error GoTo seek2:
If tabRooms.Fields("Map Number") = tExit.Map And tabRooms.Fields("Room Number") = tExit.Room Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabRooms.Index = "idxRooms"
tabRooms.Seek "=", tExit.Map, tExit.Room
If tabRooms.NoMatch = True Then
    GetRoomName = tExit.Map & "/" & tExit.Room
    Exit Function
End If

ready:
On Error GoTo error:
sName = tabRooms.Fields("Name")
If sName = "" Then sName = "(no name)"
If Not bNoRoomNumber Then sName = sName & " (" & tExit.Map & "/" & tExit.Room & ")"
GetRoomName = sName

out:
Exit Function
error:
Call HandleError("GetRoomName")
Resume out:
End Function

Public Function GetRoomCMDTB(Optional ByVal sMapRoom As String, Optional ByVal nMap As Long, Optional ByVal nRoom As Long) As Long
On Error GoTo error:
Dim tExit As RoomExitType

If sMapRoom = "" Then
    tExit.Map = nMap
    tExit.Room = nRoom
Else
    tExit = ExtractMapRoom(sMapRoom)
End If

If tExit.Map = 0 Or tExit.Room = 0 Then GetRoomCMDTB = 0: Exit Function

On Error GoTo seek2:
If tabRooms.Fields("Map Number") = tExit.Map And tabRooms.Fields("Room Number") = tExit.Room Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabRooms.Index = "idxRooms"
tabRooms.Seek "=", tExit.Map, tExit.Room
If tabRooms.NoMatch = True Then
    GetRoomCMDTB = 0
    tabRooms.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:
GetRoomCMDTB = tabRooms.Fields("CMD")

out:
Exit Function
error:
Call HandleError("GetRoomCMDTB")
Resume out:
End Function

Public Function GetItemLimit(ByVal nItemNumber As Long) As Integer
On Error GoTo error:

GetItemLimit = -1

If nItemNumber = 0 Then Exit Function

On Error GoTo seek2:
If tabItems.Fields("Number") = nItemNumber Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabItems.Index = "pkItems"
tabItems.Seek "=", nItemNumber
If tabItems.NoMatch = True Then
    tabItems.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:
GetItemLimit = tabItems.Fields("Limit")

out:
Exit Function
error:
Call HandleError("GetItemLimit")
GetItemLimit = -1
Resume out:
End Function

Public Function RaceHasAbility(ByVal nRace As Long, ByVal nAbility As Integer) As Integer
Dim x As Integer
On Error GoTo error:

'-31337 = does not have

RaceHasAbility = -31337
If nAbility <= 0 Or nRace <= 0 Then Exit Function

On Error GoTo seek2:
If tabRaces.Fields("Number") = nRace Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabRaces.Index = "pkRaces"
tabRaces.Seek "=", nRace
If tabRaces.NoMatch Then
    tabRaces.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:
For x = 0 To 9
    If tabRaces.Fields("Abil-" & x) = nAbility Then
        RaceHasAbility = tabRaces.Fields("AbilVal-" & x)
        Exit Function
    End If
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("RaceHasAbility")
Resume out:
End Function

Public Function ClassHasAbility(ByVal nClass As Long, ByVal nAbility As Integer) As Integer
Dim x As Integer
On Error GoTo error:

'-31337 = does not have

ClassHasAbility = -31337
If nAbility <= 0 Or nClass <= 0 Then Exit Function

On Error GoTo seek2:
If tabClasses.Fields("Number") = nClass Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabClasses.Index = "pkClasses"
tabClasses.Seek "=", nClass
If tabClasses.NoMatch Then
    tabClasses.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:
For x = 0 To 9
    If tabClasses.Fields("Abil-" & x) = nAbility Then
        ClassHasAbility = tabClasses.Fields("AbilVal-" & x)
        Exit Function
    End If
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("ClassHasAbility")
Resume out:
End Function

Public Function ItemHasAbility(ByVal nItemNumber As Long, ByVal nAbility As Integer) As Integer
Dim x As Integer
On Error GoTo error:

'-31337 = does not have

ItemHasAbility = -31337
If nAbility <= 0 Or nItemNumber <= 0 Then Exit Function

On Error GoTo seek2:
If tabItems.Fields("Number") = nItemNumber Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabItems.Index = "pkItems"
tabItems.Seek "=", nItemNumber
If tabItems.NoMatch Then
    tabItems.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:
For x = 0 To 9
    If tabItems.Fields("Abil-" & x) = nAbility Then
        ItemHasAbility = tabItems.Fields("AbilVal-" & x)
        Exit Function
    End If
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("ItemHasAbility")
Resume out:
End Function

Public Function ItemIsChest(ByVal nItemNumber As Long) As Boolean
On Error GoTo error:

On Error GoTo seek2:
If tabItems.Fields("Number") = nItemNumber Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabItems.Index = "pkItems"
tabItems.Seek "=", nItemNumber
If tabItems.NoMatch Then
    tabItems.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:
If tabItems.Fields("ItemType") = 8 Then ItemIsChest = True

out:
On Error Resume Next
Exit Function
error:
Call HandleError("ItemIsChest")
Resume out:
End Function

Public Function GetItemCost(ByVal nNum As Long, Optional ByVal MarkUp As Integer) As typItemCostDetail
On Error GoTo error:

If nNum = 0 Or tabItems.RecordCount = 0 Then
    GetItemCost.Cost = 0
    GetItemCost.Coin = 0
    Exit Function
End If

On Error GoTo seek2:
If tabItems.Fields("Number") = nNum Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabItems.Index = "pkItems"
tabItems.Seek "=", nNum
If tabItems.NoMatch = True Then
    GetItemCost.Cost = 0
    GetItemCost.Coin = 0
    tabItems.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:
If tabItems.Fields("Price") = 0 Then
    GetItemCost.Cost = 0
    GetItemCost.Coin = 0
Else
    If MarkUp > 0 Then
        GetItemCost.Cost = tabItems.Fields("Price") + Fix(tabItems.Fields("Price") * (MarkUp / 100))
    Else
        GetItemCost.Cost = tabItems.Fields("Price")
    End If
    
    GetItemCost.Coin = tabItems.Fields("Currency")
End If


Exit Function
error:
HandleError
GetItemCost.Cost = 0
GetItemCost.Coin = 0
End Function

Public Function GetItemWeight(ByVal nNum As Long) As Long
On Error GoTo error:

If nNum = 0 Then Exit Function

On Error GoTo seek2:
If tabItems.Fields("Number") = nNum Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabItems.Index = "pkItems"
tabItems.Seek "=", nNum
If tabItems.NoMatch = True Then
    tabItems.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:
GetItemWeight = tabItems.Fields("Encum")

Exit Function
error:
Call HandleError("GetItemWeight")
End Function

Public Function GetItemUses(ByVal nNum As Long) As Long
On Error GoTo error:

If nNum = 0 Then Exit Function

On Error GoTo seek2:
If tabItems.Fields("Number") = nNum Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabItems.Index = "pkItems"
tabItems.Seek "=", nNum
If tabItems.NoMatch = True Then
    tabItems.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:
GetItemUses = tabItems.Fields("UseCount")

Exit Function
error:
Call HandleError("GetItemUses")
End Function


'Public Function GetItemCostType(ByVal nNum As Long) As Integer
'
'On Error GoTo Error:
'
'If nNum = 0 Then GetItemCostType = 0: Exit Function
'If tabItems.RecordCount = 0 Then GetItemCostType = 0: Exit Function
'
'If Not tabItems.Fields("Number") = nNum Then
'    tabItems.Index = "pkItems"
'    tabItems.Seek "=", nNum
'    If tabItems.NoMatch = True Then
'        GetItemCostType = 0
'        Exit Function
'    End If
'End If
'
'GetItemCostType = tabItems.Fields("Currency")
'
'Exit Function
'Error:
'Call HandleError("GetItemCostType")
'
'End Function

Public Function GetItemName(ByVal nNum As Long, Optional ByVal bNoNumber As Boolean) As String
On Error GoTo error:

If nNum = 0 Then GetItemName = "None": Exit Function
GetItemName = nNum
If tabItems.RecordCount = 0 Then Exit Function

On Error GoTo seek2:
If tabItems.Fields("Number") = nNum Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabItems.Index = "pkItems"
tabItems.Seek "=", nNum
If tabItems.NoMatch = True Then
    tabItems.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:
GetItemName = tabItems.Fields("Name")
If Not bNoNumber Then GetItemName = GetItemName & "(" & nNum & ")"

out:
Exit Function
error:
Call HandleError("GetItemName")
Resume out:
End Function

Public Function ItemIsWeapon(ByVal nNum As Long) As Boolean
On Error GoTo error:

If nNum = 0 Then Exit Function
If tabItems.RecordCount = 0 Then Exit Function

On Error GoTo seek2:
If tabItems.Fields("Number") = nNum Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabItems.Index = "pkItems"
tabItems.Seek "=", nNum
If tabItems.NoMatch = True Then
    tabItems.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:
If tabItems.Fields("ItemType") = 1 Then ItemIsWeapon = True

out:
Exit Function
error:
Call HandleError("ItemIsWeapon")
Resume out:
End Function

Public Function GetCurrentSpellMinMax(Optional ByRef bUseLevel As Boolean, Optional ByVal nLevel As Integer, Optional ByRef bNoHeader As Boolean) As SpellMinMaxDur
Dim nMin As Currency, nMinIncr As Currency, nMinLVLs As Currency
Dim nMax As Currency, nMaxIncr As Currency, nMaxLVLs As Currency
Dim nDur As Currency, nDurIncr As Currency, nDurLVLs As Currency
Dim sMin As String, sMax As String, sDur As String
On Error GoTo error:

If tabSpells Is Nothing Then Exit Function
If tabSpells.EOF Then Exit Function

nMin = tabSpells.Fields("MinBase")
nMinIncr = tabSpells.Fields("MinInc")
nMinLVLs = tabSpells.Fields("MinIncLVLs")

nMax = tabSpells.Fields("MaxBase")
nMaxIncr = tabSpells.Fields("MaxInc")
nMaxLVLs = tabSpells.Fields("MaxIncLVLs")

nDur = tabSpells.Fields("Dur")
nDurIncr = tabSpells.Fields("DurInc")
nDurLVLs = tabSpells.Fields("DurIncLVLs")

If bUseLevel Then
    If (nMinIncr = 0 Or nMinLVLs = 0) And (nMaxIncr = 0 Or nMaxLVLs = 0) And _
        (nDurIncr = 0 Or nDurLVLs = 0) Then bUseLevel = False
End If

If tabSpells.Fields("Cap") = 0 And tabSpells.Fields("ReqLevel") = 0 And bUseLevel = False Then
    sDur = nDur
    sMax = nMax
    sMin = nMin
Else
    'if there is an amount specified in the ability value, dont use the spells min and max
    'If Not tabSpells.Fields("Ability Value 0") = 0 Then
    '    sMin = tabSpells.Fields("Ability Value 0")
    '    sMax = tabSpells.Fields("Ability Value 0")
    '    GoTo CalcDur:
    'End If
    
    'figure out mins and maxs...
    If nMinLVLs = 0 Or nMinIncr = 0 Then
        sMin = nMin
    Else
        If bUseLevel = True Then
            nMin = nMin + Fix((nMinIncr / nMinLVLs) * nLevel)
            nMin = Fix(nMin)
            sMin = nMin
        Else
            bNoHeader = True
            sMin = nMin & "+(" & Round(nMinIncr / nMinLVLs, 2) & "*lvl)"
        End If
    End If
    
    If nMaxLVLs = 0 Or nMaxIncr = 0 Then
        sMax = nMax
    Else
        If bUseLevel = True Then
            nMax = nMax + Fix((nMaxIncr / nMaxLVLs) * nLevel)
            nMax = Fix(nMax)
            sMax = nMax
        Else
            bNoHeader = True
            sMax = nMax & "+(" & Round(nMaxIncr / nMaxLVLs, 2) & "*lvl)"
        End If
    End If
    
'CalcDur:
    If nDurLVLs = 0 Or nDurIncr = 0 Then
        sDur = nDur
    Else
        If bUseLevel = True Then
            nDur = nDur + Fix((nDurIncr / nDurLVLs) * nLevel)
            nDur = Fix(nDur)
            sDur = nDur
        Else
            sDur = nDur & "+(" & Round(nDurIncr / nDurLVLs, 2) & "*lvl)"
        End If
    End If
End If

out:
On Error Resume Next
GetCurrentSpellMinMax.nMin = nMin
GetCurrentSpellMinMax.nMax = nMax
GetCurrentSpellMinMax.nDur = nDur
GetCurrentSpellMinMax.sMin = sMin
GetCurrentSpellMinMax.sMax = sMax
GetCurrentSpellMinMax.sDur = sDur
Exit Function
error:
Call HandleError("GetSpellMinMax")
Resume out:
End Function
Public Function PullSpellEQ(ByVal bCalcLevel As Boolean, Optional ByVal nLevel As Integer, _
    Optional ByVal nSpell As Long, Optional ByRef LV As ListView, Optional bMinMaxDamageOnly As Boolean = False, _
    Optional bForMonster As Boolean, Optional ByVal bPercentColumn As Boolean, Optional ByVal bIsNested As Boolean, _
    Optional ByVal bNoShowLevel As Boolean) As String
Dim oLI As ListItem, sTemp As String
Dim sMin As String, sMax As String, sDur As String, sDetail As String
Dim nMin As Currency, nMax As Currency, nDur As Currency, tSpellMinMaxDur As SpellMinMaxDur
Dim sMinHeader As String, sMaxHeader As String, sRemoves As String, bUseLevel As Boolean
Dim y As Long, nAbilValue As Long, x As Integer, bNoHeader As Boolean, nMap As Long
Dim bDoesDamage As Boolean, sEndCastPercent As String, sEndONE As String, sEndTWO As String

On Error GoTo error:

nSpellNest = nSpellNest + 1

If tabSpells.RecordCount = 0 Then PullSpellEQ = "(No Spell Records)": GoTo out:
If nSpellNest > 19 Then PullSpellEQ = " ... to infinity and beyond?": GoTo out:

If bQuickSpell And nSpellNest > 1 Then
    PullSpellEQ = "(click)"
    GoTo out:
End If

'base + ((how_much_incr / lvls_for_incr) * level)

If Not nSpell = 0 Then
    tabSpells.Index = "pkSpells"
    tabSpells.Seek "=", nSpell
    If tabSpells.NoMatch = True Then
        PullSpellEQ = "?"
        tabSpells.MoveFirst
        GoTo out:
    End If
Else
    nSpell = tabSpells.Fields("Number")
End If

bUseLevel = bCalcLevel
If bUseLevel Then
    'use the value in the global filter for level if a level hasn't been specified
    If nLevel = 0 And frmMain.chkGlobalFilter.Value = 1 Then
        nLevel = val(frmMain.txtGlobalLevel(0).Text)
    End If
    
    'make the level less if it's above the level cap, and more if it's below the required, except for monster attacks
    If Not bForMonster Then
        If nLevel > tabSpells.Fields("Cap") And tabSpells.Fields("Cap") > 0 Then nLevel = tabSpells.Fields("Cap")
        If nLevel < tabSpells.Fields("ReqLevel") Then nLevel = tabSpells.Fields("ReqLevel")
    End If
    If nLevel < 1 Then nLevel = tabSpells.Fields("ReqLevel")
    
    If nLevel = 0 Then bUseLevel = False
End If

tSpellMinMaxDur = GetCurrentSpellMinMax(bUseLevel, nLevel, bNoHeader)

nMin = tSpellMinMaxDur.nMin
nMax = tSpellMinMaxDur.nMax
nDur = tSpellMinMaxDur.nDur
sMin = tSpellMinMaxDur.sMin
sMax = tSpellMinMaxDur.sMax
sDur = tSpellMinMaxDur.sDur

For x = 0 To 9
    If Not tabSpells.Fields("Abil-" & x) = 0 Then
    
        Select Case tabSpells.Fields("Abil-" & x)
            Case 1, 8, 17, 18, 19:
                bDoesDamage = True
                If bMinMaxDamageOnly Then Exit For
        End Select
        
        sMinHeader = ""
        sMaxHeader = ""
        nAbilValue = tabSpells.Fields("AbilVal-" & x)
        'If nAbilValue = 0 And nMin = nMax Then nAbilValue = nMin
        
        If tabSpells.Fields("Abil-" & x) = 122 Then 'RemoveSpell
            If bQuickSpell Then
                If sRemoves = "" Then sRemoves = "click"
            Else
                If Not sRemoves = "" Then sRemoves = sRemoves & ", "
                sRemoves = sRemoves & GetSpellName(nAbilValue, bHideRecordNumbers)
            End If
        ElseIf tabSpells.Fields("Abil-" & x) = 137 Then
            'shock -- ingnore it (it's just the message)
        Else
            'If Not sDetail = "" Then sDetail = sDetail & ", "
            
            If nAbilValue = 0 Then
                Select Case tabSpells.Fields("Abil-" & x)
                    Case 140: 'teleport
                        sDetail = AutoAppend(sDetail, GetAbilityStats(tabSpells.Fields("Abil-" & x), , IIf(LV Is Nothing, Nothing, LV), , bPercentColumn) _
                            & " " & IIf(sMin = sMax, sMin, sMin & " to " & sMax))
                        If Not LV Is Nothing Then
                            nMap = 0
                            For y = 0 To 9
                                If tabSpells.Fields("Abil-" & y) = 141 Then 'tele map
                                    nMap = tabSpells.Fields("AbilVal-" & y)
                                End If
                            Next y
                            
                            If nMap > 0 Then
                                For y = val(sMin) To val(sMax)
                                    If bPercentColumn Then
                                        Set oLI = LV.ListItems.Add()
                                        oLI.Text = ""
                                        oLI.ListSubItems.Add , , "Teleport: " & GetRoomName(, nMap, y, False)
                                        oLI.ListSubItems(1).Tag = nMap & "/" & y
                                    Else
                                        Set oLI = LV.ListItems.Add(, , "Teleport: " & GetRoomName(, nMap, y, False))
                                        oLI.Tag = nMap & "/" & y
                                    End If
                                    Set oLI = Nothing
                                Next y
                            End If
                        End If
                    Case 148: 'textblock
                        sDetail = AutoAppend(sDetail, GetAbilityStats(tabSpells.Fields("Abil-" & x), , LV, , bPercentColumn) _
                            & " " & IIf(sMin = sMax, sMin, sMin & " to " & sMax))
                        If Not LV Is Nothing Then
                            For y = val(sMin) To val(sMax)
                                If bPercentColumn Then
                                    Set oLI = LV.ListItems.Add()
                                    oLI.Text = ""
                                    oLI.ListSubItems.Add , , "Execute: Textblock " & y
                                    oLI.ListSubItems(1).Tag = y
                                Else
                                    Set oLI = LV.ListItems.Add(, , "Execute: Textblock " & y)
                                    oLI.Tag = y
                                End If
                                Set oLI = Nothing
                            Next y
                        End If
                    Case 164: 'endcast %
                        sEndCastPercent = nAbilValue & "% "
                    Case 151: 'endcast
                        If bQuickSpell Then
                            If nMax > nMin Then
                                sEndONE = AutoAppend(sEndONE, sEndCastPercent & "End cast " & nMin & " to " & nMax)
                            Else
                                sEndONE = AutoAppend(sEndONE, sEndCastPercent & "End cast " & nMin)
                            End If
                        Else
                            If nMin >= nMax Then
                                sEndONE = AutoAppend(sEndONE, sEndCastPercent & "EndCast [" & GetSpellName(nMin, bHideRecordNumbers) & ", " & PullSpellEQ(bCalcLevel, nLevel, nMin, LV, , , bPercentColumn) & "]")
                            Else
                                sEndONE = AutoAppend(sEndONE, sEndCastPercent & "EndCast [{" & GetSpellName(nMin, bHideRecordNumbers) & ", " & PullSpellEQ(bCalcLevel, nLevel, nMin, LV, , , bPercentColumn) & "}")
                                For y = nMin + 1 To nMax
                                    sEndONE = sEndONE & " OR {" & GetSpellName(y, bHideRecordNumbers) & ", " & PullSpellEQ(bCalcLevel, nLevel, y, LV, , , bPercentColumn) & "}"
                                Next y
                                sEndONE = sEndONE & "]"
                            End If
                        End If
                        
'                    Case 124: 'negateabil
'                        If sMin >= sMax Then
'                            sDetail = sDetail & "NegateAbility " & GetAbilityName(sMin)
'                        Else
'                            sDetail = sDetail & "NegateAbilities{" & GetAbilityName(sMin)
'                            For y = sMin + 1 To sMax
'                                sDetail = sDetail & " OR " & GetAbilityName(y)
'                            Next y
'                            sDetail = sDetail & "}"
'                        End If
                    Case 12: 'summon
                        If bQuickSpell Then
                            sDetail = AutoAppend(sDetail, "Summon")
                        Else
                            If nMin >= nMax Then
                                sTemp = GetMonsterName(nMin, bHideRecordNumbers)
                                sDetail = AutoAppend(sDetail, "Summon " & sTemp)
                                If Not LV Is Nothing Then
                                    Set oLI = LV.ListItems.Add()
                                    If bPercentColumn Then
                                        oLI.Text = ""
                                        oLI.ListSubItems.Add , , "Summon: " & sTemp
                                        oLI.ListSubItems(1).Tag = nMin
                                    Else
                                        oLI.Text = "Summon: " & sTemp
                                        oLI.Tag = nMin
                                    End If
                                    Set oLI = Nothing
                                End If
                            Else
                                sTemp = GetMonsterName(nMin, bHideRecordNumbers)
                                sDetail = AutoAppend(sDetail, "Summons{" & sTemp)
                                If Not LV Is Nothing Then
                                    Set oLI = LV.ListItems.Add()
                                    If bPercentColumn Then
                                        oLI.Text = ""
                                        oLI.ListSubItems.Add , , "Summon: " & sTemp
                                        oLI.ListSubItems(1).Tag = nMin
                                    Else
                                        oLI.Text = "Summon: " & sTemp
                                        oLI.Tag = nMin
                                    End If
                                    Set oLI = Nothing
                                End If
                                
                                For y = nMin + 1 To nMax
                                    sTemp = GetMonsterName(y, bHideRecordNumbers)
                                    sDetail = sDetail & " OR " & sTemp
                                    If Not LV Is Nothing Then
                                        Set oLI = LV.ListItems.Add()
                                        If bPercentColumn Then
                                            oLI.Text = ""
                                            oLI.ListSubItems.Add , , "Summon: " & sTemp
                                            oLI.ListSubItems(1).Tag = y
                                        Else
                                            oLI.Text = "Summon: " & sTemp
                                            oLI.Tag = y
                                        End If
                                        Set oLI = Nothing
                                    End If
                                Next y
                                sDetail = sDetail & "}"
                            End If
                        End If
                    Case 23, 51, 52, 80, 97, 98, 100, 108 To 113, 119, 138, 144, 178:
                        '23 - effectsundead
                        '51: 'anti magic
                        '52: 'evil in combat
                        '80: 'effects animal
                        '97-98 - good/evil only
                        '100: 'loyal
                        '108: 'effects living
                        '109 To 113: 'nonliving, notgood, notevil, neutral, not neutral
                        '112 - neut only
                        '119: 'del@main
                        '138: 'roomvis
                        '144: 'non magic spell
                        '178: shadowform
                        sEndTWO = AutoAppend(sEndTWO, GetAbilityStats(tabSpells.Fields("Abil-" & x)))
                    Case 7: 'DR
                        If Not bNoHeader Then
                            If val(sMin) > 0 Then sMinHeader = "+"
                            If val(sMax) > 0 Then sMaxHeader = "+"
                        End If
                        
                        If bUseLevel Then
                            sDetail = AutoAppend(sDetail, GetAbilityStats(tabSpells.Fields("Abil-" & x), , LV, , bPercentColumn) _
                                & " " & IIf(nMin = nMax, sMinHeader & (nMin / 10), sMinHeader & (nMin / 10) & " to " & sMaxHeader & (nMax / 10)))
                        Else
                            sDetail = AutoAppend(sDetail, GetAbilityStats(tabSpells.Fields("Abil-" & x), , LV, , bPercentColumn) _
                                & " " & IIf(sMin = sMax, sMinHeader & sMin, sMinHeader & sMin & " to " & sMaxHeader & sMax))
                        End If
                    Case Else:

                        If Not bNoHeader Then
                            Select Case tabSpells.Fields("Abil-" & x)
                                Case 1, 8, 17, 18, 19, 140, 141, 148:
                                'damage, drain, damage(on armr), poison, heal, teleport room, teleport map, textblocks
                                ' *** ALSO ADD THESE TO GetAbilityStats ***
                                Case Else:
                                    If val(sMin) > 0 Then sMinHeader = "+"
                                    If val(sMax) > 0 Then sMaxHeader = "+"
                            End Select
                        End If
                        
                        'sDetail = sDetail & GetAbilityStats(tabSpells.Fields("Abil-" & x), , IIf(LV Is Nothing, Nothing, LV)) _
                            & " " & IIf(sMin = sMax, sMinHeader & sMin, sMinHeader & sMin & " to " & sMaxHeader & sMax)
                        sDetail = AutoAppend(sDetail, GetAbilityStats(tabSpells.Fields("Abil-" & x), , LV, bCalcLevel, bPercentColumn) _
                            & " " & IIf(sMin = sMax, sMinHeader & sMin, sMinHeader & sMin & " to " & sMaxHeader & sMax))
                        
                End Select
                
            Else 'abilval <> 0
                
                Select Case tabSpells.Fields("Abil-" & x)
                    Case 148: 'textblock
                        sDetail = AutoAppend(sDetail, GetAbilityStats(tabSpells.Fields("Abil-" & x), nAbilValue, IIf(LV Is Nothing, Nothing, LV), , bPercentColumn))
                        If Not LV Is Nothing Then
                            If bPercentColumn Then
                                Set oLI = LV.ListItems.Add()
                                oLI.Text = ""
                                oLI.ListSubItems.Add , , "Execute: Textblock " & nAbilValue
                                oLI.ListSubItems(1).Tag = nAbilValue
                            Else
                                Set oLI = LV.ListItems.Add(, , "Execute: Textblock " & nAbilValue)
                                oLI.Tag = nAbilValue
                            End If
                            Set oLI = Nothing
                        End If
                    Case 12: 'summon
                        If bQuickSpell Then
                            sDetail = AutoAppend(sDetail, "Summon")
                        Else
                            sTemp = GetMonsterName(nAbilValue, bHideRecordNumbers)
                            sDetail = AutoAppend(sDetail, "Summon " & sTemp)
                            If Not LV Is Nothing Then
                                Set oLI = LV.ListItems.Add()
                                If bPercentColumn Then
                                    oLI.Text = ""
                                    oLI.ListSubItems.Add , , "Summon: " & sTemp
                                    oLI.ListSubItems(1).Tag = nAbilValue
                                Else
                                    oLI.Text = "Summon: " & sTemp
                                    oLI.Tag = nAbilValue
                                End If
                                Set oLI = Nothing
                            End If
                        End If
                        
                    Case 140: 'teleport
                        sDetail = AutoAppend(sDetail, GetAbilityStats(tabSpells.Fields("Abil-" & x), nAbilValue, IIf(LV Is Nothing, Nothing, LV), , bPercentColumn))
                        If Not LV Is Nothing Then
                            nMap = 0
                            For y = 0 To 9
                                If tabSpells.Fields("Abil-" & y) = 141 Then
                                    nMap = tabSpells.Fields("AbilVal-" & y)
                                End If
                            Next y
                            
                            If nMap > 0 Then
                                If bPercentColumn Then
                                    Set oLI = LV.ListItems.Add()
                                    oLI.Text = ""
                                    oLI.ListSubItems.Add , , "Teleport: " & GetRoomName(, nMap, nAbilValue, False)
                                    oLI.ListSubItems(1).Tag = nMap & "/" & nAbilValue
                                Else
                                    Set oLI = LV.ListItems.Add(, , "Teleport: " & GetRoomName(, nMap, nAbilValue, False))
                                    oLI.Tag = nMap & "/" & nAbilValue
                                End If
                                Set oLI = Nothing
                            End If
                        End If
                    'Case 178: 'shadowform
                        'sDetail = sDetail & GetAbilityStats(tabSpells.Fields("Abil-" & x))
                    Case 164: 'endcast %
                        sEndCastPercent = nAbilValue & "% "
                    Case 151: 'endcast
                        sEndONE = AutoAppend(sEndONE, sEndCastPercent & GetAbilityStats(tabSpells.Fields("Abil-" & x), nAbilValue, LV, bCalcLevel, bPercentColumn))
                        
                    Case 23, 51, 52, 80, 97, 98, 100, 108 To 113, 119, 138, 144, 178:
                        '23 - effectsundead
                        '51: 'anti magic
                        '52: 'evil in combat
                        '80: 'effects animal
                        '97-98 - good/evil only
                        '100: 'loyal
                        '108: 'effects living
                        '109 To 113: 'nonliving, notgood, notevil, neutral, not neutral
                        '112 - neut only
                        '119: 'del@main
                        '138: 'roomvis
                        '144: 'non magic spell
                        '178: shadowform
                        sEndTWO = AutoAppend(sEndTWO, GetAbilityStats(tabSpells.Fields("Abil-" & x), nAbilValue))
                        
                    Case Else:
                        sDetail = AutoAppend(sDetail, GetAbilityStats(tabSpells.Fields("Abil-" & x), nAbilValue, LV, bCalcLevel, bPercentColumn))
                        
                End Select
            End If
            
            If Right(sDetail, 2) = ", " Then sDetail = Left(sDetail, Len(sDetail) - 2)
        End If
        
        'reposition in case the ability function changed it
        If Not tabSpells.Fields("Number") = nSpell Then tabSpells.Seek "=", nSpell
    End If
Next x

If bMinMaxDamageOnly Then
    If bDoesDamage Then
        PullSpellEQ = sMin & ":" & sMax & IIf(nDur > 0, ":" & sDur, "")
    Else
        PullSpellEQ = "0:0:0"
    End If
    GoTo out:
End If

If sDetail = "" And sRemoves = "" And sEndONE = "" And sEndTWO = "" Then
    PullSpellEQ = "(No EQ)"
    GoTo out:
End If

PullSpellEQ = sDetail
If Len(sEndONE) > 0 Then PullSpellEQ = AutoAppend(PullSpellEQ, sEndONE)

If Not bIsNested And tabSpells.Fields("EnergyCost") > 0 And tabSpells.Fields("EnergyCost") <= 500 Then
    PullSpellEQ = PullSpellEQ & " x" & Fix(1000 / tabSpells.Fields("EnergyCost"))
    If bQuickSpell Then
        PullSpellEQ = PullSpellEQ & "/rnd"
    Else
        PullSpellEQ = PullSpellEQ & " times/round"
    End If
End If

If Len(sEndTWO) > 0 Then PullSpellEQ = AutoAppend(PullSpellEQ, sEndTWO)

If bUseLevel = True And Not bNoShowLevel Then
    If tabSpells.Fields("Cap") > 0 Or tabSpells.Fields("ReqLevel") > 0 Then
        PullSpellEQ = "(@lvl " & nLevel & "): " & PullSpellEQ
    End If
End If

'If Not sDetail = "" Then
'    PullSpellEQ = PullSpellEQ & ", " & sDetail
'End If

If Not sDur = "0" Then
    If Not PullSpellEQ = "" Then PullSpellEQ = PullSpellEQ & " "
    PullSpellEQ = PullSpellEQ & "for " & sDur & " rounds"
End If

If bQuickSpell Then GoTo out:

If Not sRemoves = "" Then
    If Not PullSpellEQ = "" Then PullSpellEQ = PullSpellEQ & " -- "
    PullSpellEQ = PullSpellEQ & "RemovesSpells(" & sRemoves & ")"
End If

out:
On Error Resume Next
nSpellNest = nSpellNest - 1
Exit Function

error:
Call HandleError("PullSpellEQ")
Resume out:
End Function

Public Function GetSpellMinDamage(ByVal nSpellNumber As Long, Optional ByVal nCastLevel As Integer, Optional nEnergyRem As Integer, _
    Optional bForMonster As Boolean, Optional bHealsInstead As Boolean) As Long
Dim bDoesDamage As Boolean, x As Integer, nEndCast As Long
On Error GoTo error:

GetSpellMinDamage = 0
If nSpellNumber = 0 Then Exit Function

On Error GoTo seek2:
If tabSpells.Fields("Number") = nSpellNumber Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabSpells.Index = "pkSpells"
tabSpells.Seek "=", nSpellNumber
If tabSpells.NoMatch = True Then
    tabSpells.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:
For x = 0 To 9
    Select Case tabSpells.Fields("Abil-" & x)
        Case 1, 8, 17, 18: '1-dmg, 8-drain, 17-dmg-mr, 18=heals
            If tabSpells.Fields("Abil-" & x) = 18 Or (tabSpells.Fields("Abil-" & x) = 8 And bHealsInstead) Then
                If Not bHealsInstead Then GoTo skip:
            Else
                If bHealsInstead Then GoTo skip:
            End If
            bDoesDamage = True
            If tabSpells.Fields("AbilVal-" & x) <> 0 Then
                GetSpellMinDamage = tabSpells.Fields("AbilVal-" & x)
            End If
        Case 151:
            nEndCast = tabSpells.Fields("AbilVal-" & x)
    End Select
skip:
Next x
If GetSpellMinDamage <> 0 Then GoTo multi_calc:
If bDoesDamage = False Then Exit Function

If Not bForMonster Or nCastLevel = 0 Then
    If nCastLevel > tabSpells.Fields("Cap") And tabSpells.Fields("Cap") > 0 Then nCastLevel = tabSpells.Fields("Cap")
    If nCastLevel < tabSpells.Fields("ReqLevel") Then nCastLevel = tabSpells.Fields("ReqLevel")
End If

If tabSpells.Fields("MinIncLVLs") = 0 Or nCastLevel < 1 Then
    GetSpellMinDamage = tabSpells.Fields("MinBase")
Else
    GetSpellMinDamage = tabSpells.Fields("MinBase") + Fix((tabSpells.Fields("MinInc") / tabSpells.Fields("MinIncLVLs")) * nCastLevel)
End If

multi_calc:
If bForMonster Then Exit Function

If nEnergyRem = 0 Then nEnergyRem = 1000
nEnergyRem = nEnergyRem - tabSpells.Fields("EnergyCost")
If nEnergyRem < 1 Then nEnergyRem = 1

If nEnergyRem >= 143 And tabSpells.Fields("EnergyCost") >= 143 Then
    If nEndCast = 0 Then
        If tabSpells.Fields("EnergyCost") <= 500 Then
            GetSpellMinDamage = GetSpellMinDamage + (GetSpellMinDamage * Fix(nEnergyRem / tabSpells.Fields("EnergyCost")))
        End If
    Else
        GetSpellMinDamage = GetSpellMinDamage + GetSpellMinDamage(nEndCast, nCastLevel, nEnergyRem, bForMonster)
    End If
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetSpellMinDamage")
Resume out:
End Function

Public Function GetSpellMaxDamage(ByVal nSpellNumber As Long, Optional ByVal nCastLevel As Integer, Optional nEnergyRem As Integer, _
    Optional bForMonster As Boolean, Optional bHealsInstead As Boolean) As Long
Dim bDoesDamage As Boolean, x As Integer, nEndCast As Long
On Error GoTo error:

GetSpellMaxDamage = 0
If nSpellNumber = 0 Then Exit Function

On Error GoTo seek2:
If tabSpells.Fields("Number") = nSpellNumber Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabSpells.Index = "pkSpells"
tabSpells.Seek "=", nSpellNumber
If tabSpells.NoMatch = True Then
    tabSpells.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:
For x = 0 To 9
    Select Case tabSpells.Fields("Abil-" & x)
        Case 1, 8, 17, 18: 'dmg/drain/dmg-mr, 18=heals
            If tabSpells.Fields("Abil-" & x) = 18 Or (tabSpells.Fields("Abil-" & x) = 8 And bHealsInstead) Then
                If Not bHealsInstead Then GoTo skip:
            Else
                If bHealsInstead Then GoTo skip:
            End If
            bDoesDamage = True
            If tabSpells.Fields("AbilVal-" & x) <> 0 Then
                GetSpellMaxDamage = tabSpells.Fields("AbilVal-" & x)
            End If
        Case 151:
            nEndCast = tabSpells.Fields("AbilVal-" & x)
    End Select
skip:
Next x
If GetSpellMaxDamage <> 0 Then GoTo multi_calc:
If bDoesDamage = False Then Exit Function

If Not bForMonster Or nCastLevel = 0 Then
    If nCastLevel > tabSpells.Fields("Cap") And tabSpells.Fields("Cap") > 0 Then nCastLevel = tabSpells.Fields("Cap")
    If nCastLevel < tabSpells.Fields("ReqLevel") Then nCastLevel = tabSpells.Fields("ReqLevel")
End If

If tabSpells.Fields("MaxIncLVLs") = 0 Or nCastLevel < 1 Then
    GetSpellMaxDamage = tabSpells.Fields("MaxBase")
Else
    GetSpellMaxDamage = tabSpells.Fields("MaxBase") + Fix((tabSpells.Fields("MaxInc") / tabSpells.Fields("MaxIncLVLs")) * nCastLevel)
End If

multi_calc:
If bForMonster Then Exit Function

If nEnergyRem = 0 Then nEnergyRem = 1000
nEnergyRem = nEnergyRem - tabSpells.Fields("EnergyCost")
If nEnergyRem < 1 Then nEnergyRem = 1

If nEnergyRem >= 143 And tabSpells.Fields("EnergyCost") >= 143 Then
    If nEndCast = 0 Then
        If tabSpells.Fields("EnergyCost") <= 500 Then
            GetSpellMaxDamage = GetSpellMaxDamage + (GetSpellMaxDamage * Fix(nEnergyRem / tabSpells.Fields("EnergyCost")))
        End If
    Else
        GetSpellMaxDamage = GetSpellMaxDamage + GetSpellMaxDamage(nEndCast, nCastLevel, nEnergyRem, bForMonster)
    End If
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetSpellMaxDamage")
Resume out:
End Function

Public Function GetSpellDuration(ByVal nSpellNumber As Long, Optional ByVal nCastLevel As Integer, Optional bForMonster As Boolean) As Long
On Error GoTo error:

GetSpellDuration = 0
If nSpellNumber = 0 Then Exit Function

On Error GoTo seek2:
If tabSpells.Fields("Number") = nSpellNumber Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabSpells.Index = "pkSpells"
tabSpells.Seek "=", nSpellNumber
If tabSpells.NoMatch = True Then
    tabSpells.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:
If Not bForMonster Or nCastLevel = 0 Then
    If nCastLevel > tabSpells.Fields("Cap") And tabSpells.Fields("Cap") > 0 Then nCastLevel = tabSpells.Fields("Cap")
    If nCastLevel < tabSpells.Fields("ReqLevel") Then nCastLevel = tabSpells.Fields("ReqLevel")
End If

If tabSpells.Fields("DurIncLVLs") = 0 Or nCastLevel < 1 Then
    GetSpellDuration = tabSpells.Fields("Dur")
Else
    GetSpellDuration = tabSpells.Fields("Dur") + Fix((tabSpells.Fields("DurInc") / tabSpells.Fields("DurIncLVLs")) * nCastLevel)
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetSpellDuration")
Resume out:
End Function

Public Function SpellIsAreaAttack(ByVal nSpellNumber As Long) As Boolean
On Error GoTo error:

SpellIsAreaAttack = False
On Error GoTo seek2:
If tabSpells.Fields("Number") = nSpellNumber Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabSpells.Index = "pkSpells"
tabSpells.Seek "=", nSpellNumber
If tabSpells.NoMatch Then
    tabSpells.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:
Select Case tabSpells.Fields("Targets")
    Case 9, 11, 12:
        SpellIsAreaAttack = True
End Select

out:
On Error Resume Next
Exit Function
error:
Call HandleError("SpellIsAreaAttack")
Resume out:
End Function

Public Function SpellHasAbility(ByVal nSpellNumber As Long, ByVal nAbility As Integer) As Integer
Dim x As Integer
On Error GoTo error:

'-1 = does not have
'>=0 = value of ability

SpellHasAbility = -1
If nAbility <= 0 Or nSpellNumber <= 0 Then Exit Function

On Error GoTo seek2:
If tabSpells.Fields("Number") = nSpellNumber Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabSpells.Index = "pkSpells"
tabSpells.Seek "=", nSpellNumber
If tabSpells.NoMatch Then
    tabSpells.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:
For x = 0 To 9
    If tabSpells.Fields("Abil-" & x) = nAbility Then
        SpellHasAbility = tabSpells.Fields("AbilVal-" & x)
        Exit Function
    End If
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("SpellHasAbility")
Resume out:
End Function

Public Function GetTextblockAction(ByVal nTextblockNumber As Long) As String
On Error GoTo error:

If nTextblockNumber = 0 Then
    GetTextblockAction = "none": Exit Function
End If

tabTBInfo.Index = "pkTBInfo"
tabTBInfo.Seek "=", nTextblockNumber
If tabTBInfo.NoMatch Then
    GetTextblockAction = "none"
    Exit Function
End If
GetTextblockAction = tabTBInfo.Fields("Action")

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetTextblockAction")
Resume out:
End Function

Public Function GetTextblockCMDS(ByVal nTextblockNumber As Long, Optional ByVal nMaxLength As Integer) As String
Dim x1 As Integer, x2 As Integer, sDecrypted As String

If nTextblockNumber = 0 Then GetTextblockCMDS = "none": Exit Function

tabTBInfo.Index = "pkTBInfo"
tabTBInfo.Seek "=", nTextblockNumber
If tabTBInfo.NoMatch Then
    GetTextblockCMDS = "Textblock " & nTextblockNumber & " not found."
    Exit Function
End If
    
sDecrypted = tabTBInfo.Fields("Action")
If sDecrypted = "" Then
'    If tabTBInfo.Fields("LinkTo") > 0 Then
'        tabTBInfo.Index = "pkTBInfo"
'        tabTBInfo.Seek "=", tabTBInfo.Fields("LinkTo")
'        If tabTBInfo.NoMatch Then
'            GetTextblockCMDS = "none"
'            Exit Function
'        End If
'    Else
        GetTextblockCMDS = "none"
        Exit Function
'    End If
End If

x1 = 1
x1 = InStr(x1, sDecrypted, ":")
If x1 = 0 Then GetTextblockCMDS = "none": Exit Function

GetTextblockCMDS = Mid(sDecrypted, 1, x1 - 1)

x1 = x1 + 1
Do While x1 < Len(sDecrypted)
    x1 = InStr(x1, sDecrypted, Chr(10)) + 1
    If x1 = 1 Then GoTo done:
    
    x2 = InStr(x1, sDecrypted, ":")
    If x2 = 0 Then GoTo done:
    GetTextblockCMDS = GetTextblockCMDS & ", " & Mid(sDecrypted, x1, x2 - x1)
    
    x1 = x2 + 1
Loop


done:
GetTextblockCMDS = Replace(GetTextblockCMDS, "*", "")
GetTextblockCMDS = Replace(GetTextblockCMDS, "|", " OR ")

If nMaxLength > 0 And Len(GetTextblockCMDS) > nMaxLength Then
    GetTextblockCMDS = Left(GetTextblockCMDS, nMaxLength - 1) & "+"
End If

End Function

Public Function GetTextblockTrigger(ByVal nTextblockNumber As Long, ByVal nValue As Long) As String
Dim x1 As Integer
Dim z As Integer, sCommand As String

On Error GoTo error:

If nTextblockNumber = 0 Then GetTextblockTrigger = "none": Exit Function

tabTBInfo.Index = "pkTBInfo"
tabTBInfo.Seek "=", nTextblockNumber
If tabTBInfo.NoMatch Then
    GetTextblockTrigger = "Textblock " & nTextblockNumber & " not found."
    Exit Function
End If

If tabTBInfo.Fields("LinkTo") = nValue Then
    GetTextblockTrigger = "[dialog link]"
    Exit Function
End If

If InStr(1, tabTBInfo.Fields("Action"), "random " & nValue) > 0 Then
    GetTextblockTrigger = "[random " & nValue & "]"
ElseIf InStr(1, tabTBInfo.Fields("Action"), "text " & nValue) > 0 Then
    GetTextblockTrigger = "[text " & nValue & "]"
ElseIf InStr(1, tabTBInfo.Fields("Action"), ":" & nValue) > 0 Then
    
    sCommand = Left(tabTBInfo.Fields("Action"), InStr(1, tabTBInfo.Fields("Action"), ":" & nValue) - 1)
    
    x1 = 1
    Do While x1 < Len(sCommand)
        z = InStr(x1, sCommand, Chr(10))
        If z = 0 Then
            z = x1
            Exit Do
        Else
            x1 = z + 1
        End If
    Loop
    
    sCommand = Right(sCommand, Len(sCommand) - z + 1)
    
    GetTextblockTrigger = "[ask monster " & sCommand & "]"
    
End If

Exit Function
error:
Call HandleError("GetTextblockTrigger")
End Function

Public Function GetTextblockCMDLine(ByVal sCommand As String, Optional ByVal sTextblockData As String, _
    Optional ByVal nTextblockNumber As Long) As String
Dim x1 As Integer, y As Integer
On Error GoTo error:

If nTextblockNumber = 0 And sTextblockData = "" Then
    GetTextblockCMDLine = "unknown": Exit Function
End If

If sTextblockData = "" Then
    tabTBInfo.Index = "pkTBInfo"
    tabTBInfo.Seek "=", nTextblockNumber
    If tabTBInfo.NoMatch Then
        GetTextblockCMDLine = "unknown"
        Exit Function
    End If
    sTextblockData = tabTBInfo.Fields("Action")
End If

'z = slook number
'x1 = end of last command
'x2 = beginning of new command
'slook execution text being searched
'nvalue value of execution
'y = temp position on linebreak

If Not Right(sCommand, 1) = ":" Then sCommand = sCommand & ":"

x1 = InStr(1, sTextblockData, sCommand) 'position x1 at command
If x1 = 0 Then GetTextblockCMDLine = "none": Exit Function
x1 = x1 + Len(sCommand)

y = InStr(x1, sTextblockData, Chr(10))
If y = 0 Then y = Len(sTextblockData)

GetTextblockCMDLine = Mid(sTextblockData, x1, y - x1)
GetTextblockCMDLine = Replace(GetTextblockCMDLine, "*", "")
GetTextblockCMDLine = Replace(GetTextblockCMDLine, "|", " OR ")

Exit Function
error:
Call HandleError("GetTextblockCMDLine")
End Function

Public Function GetTextblockCMDText(ByVal sCommand As String, Optional ByVal sTextblockData As String, _
    Optional ByVal nTextblockNumber As Long) As String
Dim x1 As Integer, sLine As String
On Error GoTo error:

If nTextblockNumber = 0 And sTextblockData = "" Then
    GetTextblockCMDText = "": Exit Function
End If

If sTextblockData = "" Then
    tabTBInfo.Index = "pkTBInfo"
    tabTBInfo.Seek "=", nTextblockNumber
    If tabTBInfo.NoMatch Then
        GetTextblockCMDText = ""
        Exit Function
    End If
    sTextblockData = tabTBInfo.Fields("Action")
End If


x1 = InStr(1, sTextblockData, sCommand) 'position x1 at command
If x1 = 0 Then GetTextblockCMDText = "": Exit Function

sLine = Mid(sTextblockData, 1, x1)

Do While InStr(1, sLine, Chr(10)) > 0
    sLine = Mid(sLine, InStr(1, sLine, Chr(10)) + 1)
Loop
If InStr(1, sLine, ":") > 0 Then
    sLine = Left(sLine, InStr(1, sLine, ":") - 1)
End If

GetTextblockCMDText = sLine
GetTextblockCMDText = Replace(GetTextblockCMDText, "*", "")
GetTextblockCMDText = Replace(GetTextblockCMDText, "|", " OR ")
    
Exit Function
error:
Call HandleError("GetTextblockCMDText")
End Function


Public Sub GetChestItems(ByRef nChestArray() As Currency, ByVal nTBNumber As Long, _
    ByRef nNest As Long, Optional ByVal nPercentMod As Currency)
Dim sData As String, nDataPos As Long, x As Long, y As Long
Dim nPer1 As Long, nPer2 As Long, sLine As String, nValue As Long, nPercent As Currency
Dim nItemArray() As Currency
On Error GoTo error:

tabTBInfo.Index = "pkTBinfo"
tabTBInfo.Seek "=", nTBNumber
If tabTBInfo.NoMatch Then Exit Sub
sData = LCase(tabTBInfo.Fields("Action"))
If sData = Chr(0) Then Exit Sub
nDataPos = 1

If nNest > 5 Then Exit Sub
nNest = nNest + 1

If nPercentMod <= 0 Then nPercentMod = 1
ReDim nItemArray(1 To 2, 0) '1=number, 2=percent

'first we collect all the items and total their %'s
Do While nDataPos < Len(sData)
    x = InStr(nDataPos, sData, ":")
    If x > nDataPos Then
        nPer1 = val(Mid(sData, nDataPos, x - nDataPos))
        nPercent = (nPer1 - nPer2) / 100
        nPer2 = nPer1
        
        nDataPos = x + 1
        'nNest = nNest + 1
        
        x = InStr(nDataPos, sData, Chr(10))
        If x <= 0 Then x = Len(sData)
        sLine = LCase(Mid(sData, nDataPos, x - nDataPos))
        nDataPos = x
        
        y = 1
check_give_again:
        y = InStr(y, sLine, "giveitem ")
        If y > 0 Then
            nValue = ExtractValueFromString(Mid(sLine, y), "giveitem ")
            
            For x = 0 To UBound(nItemArray(), 2)
                If nItemArray(1, x) = nValue Then
                    nItemArray(2, x) = nItemArray(2, x) + nPercent
                    x = -1
                    Exit For
                End If
            Next x
            If x >= 0 Then
                x = UBound(nItemArray(), 2) + 1
                ReDim Preserve nItemArray(1 To 2, x)
                nItemArray(1, x) = nValue
                nItemArray(2, x) = nPercent
            End If
            
            y = y + 1
            GoTo check_give_again:
        End If
        
        y = 1
check_random_again:
        y = InStr(y, sLine, "random ")
        If y > 0 Then
            
            nValue = ExtractValueFromString(Mid(sLine, y), "random ")
            If nValue > 0 Then
                Call GetChestItems(nChestArray(), nValue, nNest, (nPercent * nPercentMod))
            End If
            
            y = y + 1
            GoTo check_random_again:
        End If
        
'''''        If InStr(1, sLine, "giveitem ") > 0 Then
'''''            nValue = ExtractValueFromString(sLine, "item ")
'''''
'''''            For x = 0 To UBound(nItemArray(), 2)
'''''                If nItemArray(1, x) = nValue Then
'''''                    nItemArray(2, x) = nItemArray(2, x) + nPercent
'''''                    x = -1
'''''                    Exit For
'''''                End If
'''''            Next x
'''''            If x >= 0 Then
'''''                x = UBound(nItemArray(), 2) + 1
'''''                ReDim Preserve nItemArray(1 To 2, x)
'''''                nItemArray(1, x) = nValue
'''''                nItemArray(2, x) = nPercent
'''''            End If
'''''
'''''        ElseIf InStr(1, sLine, "random ") > 0 Then
'''''            nValue = ExtractValueFromString(sLine, "random ")
'''''            If nValue > 0 Then
'''''                Call GetChestItems(nChestArray(), nValue, nNest, (nPercent * nPercentMod))
'''''            End If
'''''        End If
    Else
        nDataPos = nDataPos + 1
    End If
Loop

'then we put the collected items into the chest array
'...this is actually sort of unecessary i found out afterwards
For y = 0 To UBound(nItemArray(), 2)
    If nItemArray(1, y) > 0 Then
        nPercent = nItemArray(2, y)
        
        For x = 0 To UBound(nChestArray(), 2)
            If nChestArray(1, x) = nItemArray(1, y) Then
                'If nChestArray(3, x) = 0 Then nChestArray(3, x) = 1
                nChestArray(2, x) = nChestArray(2, x) + _
                    (nChestArray(3, x) * nPercent * nPercentMod)
                nChestArray(3, x) = nChestArray(3, x) * (1 - nPercent)
                x = -1
                Exit For
            End If
        Next x
        If x >= 0 Then
            x = UBound(nChestArray(), 2) + 1
            ReDim Preserve nChestArray(1 To 3, x)
            nChestArray(1, x) = nItemArray(1, y)
            nChestArray(2, x) = nPercent * nPercentMod
            nChestArray(3, x) = 1 - nChestArray(2, x)
        End If
        
    End If
Next y

nNest = nNest - 1

Erase nItemArray()

Exit Sub
error:
Call HandleError("GetChestItems-#" & nTBNumber)
Erase nItemArray()
End Sub

Public Function CalculateMonsterItemBonuses(nMonster As Long, nAbilities As Variant) As Integer
Dim x As Integer, y As Integer, nTest As Integer
On Error GoTo error:

If Not IsDimmed(nAbilities) Then Exit Function

On Error GoTo seek2:
If tabMonsters.Fields("Number") = nMonster Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabMonsters.Index = "pkMonsters"
tabMonsters.Seek "=", nMonster
If tabMonsters.NoMatch = True Then
    tabMonsters.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:
If tabMonsters.Fields("Weapon") > 0 Then
    If GetItemLimit(tabMonsters.Fields("Weapon")) = 0 Then
        For y = LBound(nAbilities) To UBound(nAbilities)
            nTest = ItemHasAbility(tabMonsters.Fields("Weapon"), nAbilities(y))
            If nTest <> -31337 Then
                CalculateMonsterItemBonuses = CalculateMonsterItemBonuses + nTest
            End If
        Next y
    End If
End If

For x = 0 To 9
    If tabMonsters.Fields("DropItem-" & x) > 0 Then
        If GetItemLimit(tabMonsters.Fields("DropItem-" & x)) = 0 Then
            For y = LBound(nAbilities) To UBound(nAbilities)
                nTest = ItemHasAbility(tabMonsters.Fields("DropItem-" & x), nAbilities(y))
                If nTest <> -31337 Then
                    If tabMonsters.Fields("DropItem%-" & x) > 100 Then tabMonsters.Fields("DropItem%-" & x) = 100
                    CalculateMonsterItemBonuses = CalculateMonsterItemBonuses + (nTest * (tabMonsters.Fields("DropItem%-" & x) / 100))
                End If
            Next y
        End If
    End If
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalculateMonsterItemBonuses")
Resume out:
End Function

Public Sub PopulateMonsterDataToAttackSim(ByVal nMonster As Long, ByRef clsMonAtkSim As clsMonsterAttackSim)
On Error GoTo error:
Dim x As Integer, y As Integer
Dim nPercent As Integer, sTemp As String, nTest As Integer
Dim nDamageArr As Variant, nAccyArr As Variant
Dim nItemDamageBonus As Integer, nItemAccyBonus As Integer

On Error GoTo seek2:
If tabMonsters.Fields("Number") = nMonster Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabMonsters.Index = "pkMonsters"
tabMonsters.Seek "=", nMonster
If tabMonsters.NoMatch = True Then
    tabMonsters.MoveFirst
    Exit Sub
End If

ready:
On Error GoTo error:
If nNMRVer >= 1.71 Then
    clsMonAtkSim.nEnergyPerRound = tabMonsters.Fields("Energy")
Else
    clsMonAtkSim.nEnergyPerRound = 1000
End If

nDamageArr = Array(4) '4=max damage
nAccyArr = Array(22, 105, 106) '22, 105, 106 = accuracy

nItemDamageBonus = CalculateMonsterItemBonuses(nMonster, nDamageArr)
nItemAccyBonus = CalculateMonsterItemBonuses(nMonster, nAccyArr)

For x = 0 To 4
    sTemp = ""
    If tabMonsters.Fields("AttType-" & x) > 0 And tabMonsters.Fields("AttType-" & x) < 4 Then
        If nNMRVer >= 1.8 Then
            sTemp = tabMonsters.Fields("AttName-" & x)
        Else
            If tabMonsters.Fields("AttType-" & x) = 2 And tabMonsters.Fields("AttAcc-" & x) > 0 Then
                sTemp = GetSpellName(tabMonsters.Fields("AttAcc-" & x), True)
            End If
            If sTemp = "" Or sTemp = "None" Then sTemp = "Attack " & (x + 1)
        End If
        clsMonAtkSim.sAtkName(x) = Trim(sTemp)
        clsMonAtkSim.nAtkType(x) = tabMonsters.Fields("AttType-" & x)
        clsMonAtkSim.nAtkEnergy(x) = tabMonsters.Fields("AttEnergy-" & x)
        clsMonAtkSim.nAtkChance(x) = tabMonsters.Fields("Att%-" & x)
        
        If tabMonsters.Fields("AttType-" & x) = 2 Then 'spell
            
            tabSpells.Index = "pkSpells"
            tabSpells.Seek "=", tabMonsters.Fields("AttAcc-" & x)
            If tabSpells.NoMatch = True Then
                tabSpells.MoveFirst
                GoTo next_attack_slot:
            Else
                If tabSpells.Fields("Targets") = 12 Then
                    If GetSpellDuration(tabMonsters.Fields("AttAcc-" & x), tabMonsters.Fields("AttMax-" & x), True) = 0 Then
                        nTest = SpellHasAbility(tabMonsters.Fields("AttAcc-" & x), 1) '1=damage
                        If nTest > -1 Then
                            'MsgBox "Attack #" & (x + 1) & " (" & txtAtkName(x).Text & ") has an area attack spell in a regular attack slot using ability 1 (damage) instead of 17 (damage-MR). " _
                                & "This is an error and MMUD will not cast this.  Area attack spells must use ability 17 (or possibly 8-drain?).  The min/max damage and energy cost has been zero'd out for the sim to reflect the game.", vbExclamation
                            clsMonAtkSim.nAtkDuration(x) = 0
                            clsMonAtkSim.nAtkMin(x) = 0
                            clsMonAtkSim.nAtkMax(x) = 0
                            clsMonAtkSim.nAtkEnergy(x) = 0
                            GoTo next_attack_slot:
                        End If
                    End If
                End If
                
                If nNMRVer >= 1.8 Then clsMonAtkSim.nAtkResist(x) = tabSpells.Fields("TypeOfResists")
                
                clsMonAtkSim.nAtkDuration(x) = GetSpellDuration(tabMonsters.Fields("AttAcc-" & x), tabMonsters.Fields("AttMax-" & x), True)
                clsMonAtkSim.nAtkMin(x) = 0
                clsMonAtkSim.nAtkMax(x) = 0
                
                nTest = SpellHasAbility(tabMonsters.Fields("AttAcc-" & x), 1) '1=damage
                If nTest >= 0 Then
                    clsMonAtkSim.nAtkMRdmgResist(x) = 0 'NO MR resist
                    If nTest > 0 Then
                        clsMonAtkSim.nAtkMin(x) = nTest
                        clsMonAtkSim.nAtkMax(x) = nTest
                    Else
                        clsMonAtkSim.nAtkMin(x) = GetSpellMinDamage(tabMonsters.Fields("AttAcc-" & x), tabMonsters.Fields("AttMax-" & x), -1, True)
                        clsMonAtkSim.nAtkMax(x) = GetSpellMaxDamage(tabMonsters.Fields("AttAcc-" & x), tabMonsters.Fields("AttMax-" & x), -1, True)
                    End If
                End If
                
                nTest = SpellHasAbility(tabMonsters.Fields("AttAcc-" & x), 17) '17=damage
                If nTest >= 0 Then
                    clsMonAtkSim.nAtkMRdmgResist(x) = 1 'MR resist
                    If nTest > 0 Then
                        clsMonAtkSim.nAtkMin(x) = nTest
                        clsMonAtkSim.nAtkMax(x) = nTest
                    Else
                        clsMonAtkSim.nAtkMin(x) = GetSpellMinDamage(tabMonsters.Fields("AttAcc-" & x), tabMonsters.Fields("AttMax-" & x), -1, True)
                        clsMonAtkSim.nAtkMax(x) = GetSpellMaxDamage(tabMonsters.Fields("AttAcc-" & x), tabMonsters.Fields("AttMax-" & x), -1, True)
                    End If
                End If
                
                nTest = SpellHasAbility(tabMonsters.Fields("AttAcc-" & x), 8) '8=drain
                If nTest >= 0 Then
                    clsMonAtkSim.nAtkMRdmgResist(x) = 0 'NO MR resist
                    If nTest > 0 Then
                        clsMonAtkSim.nAtkMin(x) = nTest
                        clsMonAtkSim.nAtkMax(x) = nTest
                    Else
                        clsMonAtkSim.nAtkMin(x) = GetSpellMinDamage(tabMonsters.Fields("AttAcc-" & x), tabMonsters.Fields("AttMax-" & x), -1, True)
                        clsMonAtkSim.nAtkMax(x) = GetSpellMaxDamage(tabMonsters.Fields("AttAcc-" & x), tabMonsters.Fields("AttMax-" & x), -1, True)
                    End If
                End If
            End If
            clsMonAtkSim.nAtkSuccess(x) = tabMonsters.Fields("AttMin-" & x)
        Else
            clsMonAtkSim.nAtkMin(x) = tabMonsters.Fields("AttMin-" & x) + nItemDamageBonus
            clsMonAtkSim.nAtkMax(x) = tabMonsters.Fields("AttMax-" & x) + nItemDamageBonus
            clsMonAtkSim.nAtkSuccess(x) = tabMonsters.Fields("AttAcc-" & x) + nItemAccyBonus
            If tabMonsters.Fields("AttHitSpell-" & x) > 0 Then
                
                tabSpells.Index = "pkSpells"
                tabSpells.Seek "=", tabMonsters.Fields("AttHitSpell-" & x)
                If tabSpells.NoMatch = True Then
                    tabSpells.MoveFirst
                    GoTo next_attack_slot:
                Else
                    If nNMRVer >= 1.8 Then clsMonAtkSim.nAtkResist(x) = tabSpells.Fields("TypeOfResists")
                    clsMonAtkSim.sAtkHitSpellName(x) = tabSpells.Fields("Name")
                    clsMonAtkSim.nAtkDuration(x) = GetSpellDuration(tabMonsters.Fields("AttHitSpell-" & x))
                    
                    If SpellHasAbility(tabMonsters.Fields("AttHitSpell-" & x), 1) >= 0 Then
                        clsMonAtkSim.nAtkMRdmgResist(x) = 0
                        clsMonAtkSim.nAtkHitSpellMin(x) = GetSpellMinDamage(tabMonsters.Fields("AttHitSpell-" & x))
                        clsMonAtkSim.nAtkHitSpellMax(x) = GetSpellMaxDamage(tabMonsters.Fields("AttHitSpell-" & x))
                        
                    ElseIf SpellHasAbility(tabMonsters.Fields("AttHitSpell-" & x), 17) >= 0 Then
                        clsMonAtkSim.nAtkMRdmgResist(x) = 1
                        clsMonAtkSim.nAtkHitSpellMin(x) = GetSpellMinDamage(tabMonsters.Fields("AttHitSpell-" & x))
                        clsMonAtkSim.nAtkHitSpellMax(x) = GetSpellMaxDamage(tabMonsters.Fields("AttHitSpell-" & x))
                        
                    Else
                        clsMonAtkSim.nAtkHitSpellMin(x) = 0
                        clsMonAtkSim.nAtkHitSpellMax(x) = 0
                    End If
                End If
            End If
        End If
    End If
next_attack_slot:
Next x

nPercent = 0
For x = 0 To 4
    If tabMonsters.Fields("MidSpell-" & x) > 0 Then
        tabSpells.Index = "pkSpells"
        tabSpells.Seek "=", tabMonsters.Fields("MidSpell-" & x)
        If tabSpells.NoMatch = True Then
            tabSpells.MoveFirst
            'GoTo next_attack_slot:
        Else
            clsMonAtkSim.sBetweenRoundName(x) = tabSpells.Fields("Name")
            If nNMRVer >= 1.8 Then clsMonAtkSim.nBetweenRoundResistType(x) = tabSpells.Fields("TypeOfResists")
            clsMonAtkSim.nBetweenRoundChance(x) = tabMonsters.Fields("MidSpell%-" & x)
            clsMonAtkSim.nBetweenRoundDuration(x) = GetSpellDuration(tabMonsters.Fields("MidSpell-" & x), tabMonsters.Fields("MidSpellLVL-" & x), True)
            
            nTest = SpellHasAbility(tabMonsters.Fields("MidSpell-" & x), 1) '1=damage
            If nTest >= 0 Then
                clsMonAtkSim.nBetweenRoundResistDmgMR(x) = 0 'NO MR resist
                If nTest > 0 Then
                    clsMonAtkSim.nBetweenRoundMin(x) = nTest
                    clsMonAtkSim.nBetweenRoundMax(x) = nTest
                Else
                    clsMonAtkSim.nBetweenRoundMin(x) = GetSpellMinDamage(tabMonsters.Fields("MidSpell-" & x), tabMonsters.Fields("MidSpellLVL-" & x), -1, True)
                    clsMonAtkSim.nBetweenRoundMax(x) = GetSpellMaxDamage(tabMonsters.Fields("MidSpell-" & x), tabMonsters.Fields("MidSpellLVL-" & x), -1, True)
                End If
            End If
            
            nTest = SpellHasAbility(tabMonsters.Fields("MidSpell-" & x), 17) '17=damage-mr
            If nTest >= 0 Then
                clsMonAtkSim.nBetweenRoundResistDmgMR(x) = 1 'MR resist
                If nTest > 0 Then
                    clsMonAtkSim.nBetweenRoundMin(x) = nTest
                    clsMonAtkSim.nBetweenRoundMax(x) = nTest
                Else
                    clsMonAtkSim.nBetweenRoundMin(x) = GetSpellMinDamage(tabMonsters.Fields("MidSpell-" & x), tabMonsters.Fields("MidSpellLVL-" & x), -1, True)
                    clsMonAtkSim.nBetweenRoundMax(x) = GetSpellMaxDamage(tabMonsters.Fields("MidSpell-" & x), tabMonsters.Fields("MidSpellLVL-" & x), -1, True)
                End If
            End If
            
            nTest = SpellHasAbility(tabMonsters.Fields("MidSpell-" & x), 8) '8=drain
            If nTest >= 0 Then
                clsMonAtkSim.nBetweenRoundResistDmgMR(x) = 0 'NO MR resist
                If nTest > 0 Then
                    clsMonAtkSim.nBetweenRoundMin(x) = nTest
                    clsMonAtkSim.nBetweenRoundMax(x) = nTest
                Else
                    clsMonAtkSim.nBetweenRoundMin(x) = GetSpellMinDamage(tabMonsters.Fields("MidSpell-" & x), tabMonsters.Fields("MidSpellLVL-" & x), -1, True)
                    clsMonAtkSim.nBetweenRoundMax(x) = GetSpellMaxDamage(tabMonsters.Fields("MidSpell-" & x), tabMonsters.Fields("MidSpellLVL-" & x), -1, True)
                End If
            End If
        End If
    End If
Next x

For x = 0 To 4
    If Len(clsMonAtkSim.sAtkName(x)) > 0 Then
        For y = 0 To 4
            If y <> x And clsMonAtkSim.sAtkName(x) = clsMonAtkSim.sAtkName(y) Then
                clsMonAtkSim.sAtkName(x) = Trim(clsMonAtkSim.sAtkName(x)) & (x + 1)
                clsMonAtkSim.sAtkName(y) = Trim(clsMonAtkSim.sAtkName(y)) & (y + 1)
            End If
        Next y
    End If
Next x

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("PopulateMonsterDataToAttackSim")
Resume out:
End Sub

Public Function CalculateMonsterAvgDmg(ByVal nMonster As Long, Optional nNumRounds As Long = 500) As MonAttackSimReturn
On Error GoTo error:

Call clsMonAtkSim.ResetValues
clsMonAtkSim.bUseCPU = True
clsMonAtkSim.nCombatLogMaxRounds = 0
clsMonAtkSim.nNumberOfRounds = nNumRounds 'IIf(nNumRounds <> 0, nNumRounds, 500)
clsMonAtkSim.nUserMR = 50
clsMonAtkSim.bGreaterMUD = bGreaterMUD
clsMonAtkSim.bDynamicCalc = False
clsMonAtkSim.nDynamicCalcDifference = 0.001

Call PopulateMonsterDataToAttackSim(nMonster, clsMonAtkSim)

If clsMonAtkSim.nNumberOfRounds > 0 Then clsMonAtkSim.RunSim

CalculateMonsterAvgDmg.nAverageDamage = clsMonAtkSim.nAverageDamage
CalculateMonsterAvgDmg.nMaxDamage = clsMonAtkSim.GetMaxDamage

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalculateMonsterAvgDmg")
Resume out:
End Function

Public Function IsDimmed(arr As Variant) As Boolean
On Error GoTo ReturnFalse
  IsDimmed = UBound(arr) >= LBound(arr)
ReturnFalse:
End Function

Public Function TextBlockHasTeleport(ByVal nTextblock As Long, ByVal nFindRoom As Long, Optional ByVal nFindMap As Long, Optional ByVal bStrict As Boolean) As Boolean
'bStrict true == nFindMap should be > 0 and map MUST match. a missing map specified will result in false.
'bStrict false == only room must match. however, if nFindMap is specified and the textblock does specify the map and it doesn't match, then result = false
On Error GoTo error:
Dim sData As String, nDataPos As Long, sLine As String, sChar As String, nRoom As Long, nMap As Long
Dim x As Integer, y As Integer

If nTextblock <= 0 Then Exit Function

tabTBInfo.Index = "pkTBInfo"
tabTBInfo.Seek "=", nTextblock
If tabTBInfo.NoMatch Then
    tabTBInfo.MoveFirst
    Exit Function
End If

sData = tabTBInfo.Fields("Action")
nDataPos = 1

Do While nDataPos < Len(sData)
    x = InStr(nDataPos, sData, Chr(10))
    If x = 0 Then x = Len(sData)
    sLine = Mid(sData, nDataPos, x - nDataPos)
    nDataPos = x + 1
    
    x = InStr(1, sLine, "teleport ")
    If x > 0 Then
        y = x + Len("teleport ")
        x = y
        
        Do While y <= Len(sLine)
            sChar = Mid(sLine, y, 1)
            Select Case sChar
                Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
                Case " ":
                    If y > x And nRoom = 0 Then
                        nRoom = val(Mid(sLine, x, y - x))
                        x = y + 1
                    Else
                        nMap = val(Mid(sLine, x, y - x))
                        Exit Do
                    End If
                Case Else:
                    If y > x And nRoom = 0 Then
                        nRoom = val(Mid(sLine, x, y - x))
                        Exit Do
                    Else
                        nMap = val(Mid(sLine, x, y - x))
                        Exit Do
                    End If
                    Exit Do
            End Select
            y = y + 1
        Loop
        
        If nRoom = nFindRoom Then
            If nFindMap > 0 And nMap > 0 And nMap <> nFindMap Then GoTo skip:
            If bStrict And nMap <> nFindMap Then GoTo skip:
            TextBlockHasTeleport = True
            Exit Function
        End If
skip:
        nRoom = 0
        nMap = 0
    End If
Loop

        
out:
On Error Resume Next
Exit Function
error:
Call HandleError("TextBlockHasTeleport")
Resume out:
End Function
