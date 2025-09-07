Attribute VB_Name = "modMMudFunc"
Option Explicit
Option Base 0

Public Const STOCK_HIT_MIN As Integer = 8#
Public Const STOCK_HIT_CAP As Integer = 99#
Public Const GMUD_HIT_MIN As Integer = 2#
Public Const GMUD_HIT_CAP As Integer = 98#

Public bGreaterMUD As Boolean
Public Const STOCK_DODGE_CAP As Integer = 95#
Public Const GMUD_DODGE_SOFTCAP As Integer = 45#
Public Const GMUD_DODGE_CAP As Integer = 98#

'alignment = max value of alignment, i.e. anything <= -201 is saint, then <= -51 is good, etc
'Saint = -201.0f;
'Good = -51.0f;
'Neutral = 29.99f;
'Seedy = 39.99f;
'Outlaw = 79.99f;
'Criminal = 119.99f;
'Villain = 299.99f;
'FIEND = 500.0f;
Public Enum eEvilPoints
    e0_Saint = -201#
    e1_Good = -51#
    e2_Neutral = 29.99
    e3_Seedy = 39.99
    e4_Outlaw = 79.99
    e5_Criminal = 119.99
    e6_Villian = 299.99
    e7_FIEND = 500#
End Enum

Public Type RoomExitType
    Map As Long
    Room As Long
    ExitType As String
End Type

Public Enum eAttackTypeMUD
    a0_none = 0
    a1_Punch = 1
    a2_Kick = 2
    a3_Jumpkick = 3
    a4_Surprise = 4
    a5_Normal = 5
    a6_Bash = 6
    a7_Smash = 7
End Enum

Public Type tCharacterProfile
    bIsLoadedCharacter As Boolean
    nParty As Integer
    nHP As Double
    nHPRegen As Double
    nDamageThreshold As Double
    nSpellAttackCost As Double
    nSpellOverhead As Double
    nMaxMana As Double
    nSpellcasting As Integer
    nManaRegen As Double
    nMeditateRate As Double
    nEncumPCT As Double
    nAccuracy As Double
    nLevel As Long
    nClass As Long
    nRace As Long
    nCombat As Integer
    nSTR As Integer
    nAGI As Integer
    nCHA As Integer
    nWis As Integer
    nINT As Integer
    nHEA As Integer
    nCrit As Integer
    nDodge As Integer
    nDodgeCap As Integer
    nPlusMaxDamage As Integer
    nPlusMinDamage As Integer
    nPlusBSaccy As Integer
    nPlusBSmindmg As Integer
    nPlusBSmaxdmg As Integer
    nMAPlusSkill(1 To 3) As Integer
    nMAPlusAccy(1 To 3) As Integer
    nMAPlusDmg(1 To 3) As Integer
    nStealth As Integer
    bClassStealth As Boolean
    bRaceStealth As Boolean
End Type

Public Type tCombatRoundInfo
    nRTK As Double
    nRTD As Double
    sRTK As String
    sRTD As String
    nSuccess As Integer
    sSuccess As String
End Type

Public Type tSpellCastValues
    nMinCast As Long
    nMaxCast As Long
    nAvgCast As Long
    nNumCasts As Double
    nCastChance As Integer
    nAvgRoundDmg As Long
    nAvgRoundHeals As Long
    nDuration As Integer
    nDamageResisted As Long
    nFullResistChance As Integer
    nManaCost As Integer
    nOOM As Integer
    bDoesHeal As Boolean
    bDoesDamage As Boolean
    sAvgRound As String
    sLVLincreases As String
    sMMA As String
    sSpellName As String
    nCastLevel As Integer
End Type

Public Type tAttackDamage
    nMinDmg As Long
    nMaxDmg As Long
    nHitChance As Integer
    nDodgeChance As Integer
    nCritChance As Integer
    nQnDBonus As Integer
    nAccy As Integer
    nAvgHit As Long
    nAvgCrit As Long
    nMaxCrit As Long
    nAvgExtraHit As Long
    nAvgExtraSwing As Long
    nSwings As Double
    nRoundPhysical As Long
    nRoundTotal As Long
    sAttackDesc As String
    sAttackDetail As String
End Type

Public Function CalcCombatRounds(Optional ByVal nDamageOut As Long = -9999, Optional ByVal nMobHealth As Long, _
    Optional ByVal nMobDamage As Long = -1, Optional ByVal nCharHealth As Long, Optional ByVal nMobHPRegen As Long, _
    Optional ByVal nNumMobs As Double = 1, Optional ByVal nOverrideRTK As Double, _
    Optional ByVal nSurpriseDamageOut As Double = -9999, Optional ByVal nMinDamageOut As Long = -9999) As tCombatRoundInfo
On Error GoTo error:
Dim nTest As Double, nMobHP As Long, nMinDmgPct As Double

If nNumMobs < 1 Then nNumMobs = 1
If nMobDamage > 0 And nCharHealth > 0 Then CalcCombatRounds.nRTD = Round(nCharHealth / nMobDamage, 1)  'was (nMobDamage / nNumMobs) 2025.08.18
If nOverrideRTK > 0 Then CalcCombatRounds.nRTK = nOverrideRTK * nNumMobs

If nDamageOut > 0 And nMobHealth > 1 Then
    nMobHP = nMobHealth
    If nNumMobs > 1 Then nMobHP = nMobHP / nNumMobs
    
    If nOverrideRTK < 1 Then
        CalcCombatRounds.nRTK = Round(nMobHP / nDamageOut, 2)
        CalcCombatRounds.nRTK = -Int(-(CalcCombatRounds.nRTK * 2)) / 2 'round up to nearest 0.5
        CalcCombatRounds.nRTK = CalcCombatRounds.nRTK * nNumMobs
    End If
    
    If CalcCombatRounds.nRTK >= 16 And nMobHPRegen > 0 Then
        '16 is 90% of 18, arbitrarily chosen as where there begins to be a chance for the mob to regen its hp
        '18 is the number of rounds in 90 seconds
        '90 seconds is the number of rounds to mob hp regen
        'thus we are adding hp if it takes that long to kill mob
        nTest = 1
        Do While (CalcCombatRounds.nRTK - (nTest * 18)) / 18 >= 0.9
            nTest = nTest + 1
        Loop
        nMobHealth = nMobHealth + (nTest * nMobHPRegen)
        CalcCombatRounds.nRTK = Round(nMobHealth / nDamageOut, 2)
    End If
    
    If CalcCombatRounds.nRTK = 1 And nMinDamageOut >= 0 And nMinDamageOut < nDamageOut And nMinDamageOut < nMobHP Then
        nMinDmgPct = (nMobHP - nMinDamageOut) / (nDamageOut - nMinDamageOut)
        If nMinDmgPct >= 0.5 Then CalcCombatRounds.nRTK = 1.5
    End If
End If

' ===== Surprise opener credit (first target replaces the first normal round) =====
If nSurpriseDamageOut > 0# And nMobHealth > 1# And CalcCombatRounds.nRTK > 0# Then
    Dim hpPerMob As Double, rtkSingleNormal As Double, rtkSingleSurp As Double
    Dim regenPerRound As Double, regenRatio As Double, regenAtten As Double
    Dim packFade As Double, fadeGate As Double
    Dim deltaFirst As Double, adj As Double

    ' Per-mob HP (same split as elsewhere)
    hpPerMob = ccr_SafeDiv(nMobHealth, nNumMobs, nMobHealth)

    ' Baseline per-mob RTK (your 0.5-step rule)
    If nOverrideRTK > 0# Then
        rtkSingleNormal = nOverrideRTK           ' override is already per-single-mob
    Else
        rtkSingleNormal = ccr_SafeDiv(hpPerMob, ccr_Max(1#, nDamageOut), 1#)
        rtkSingleNormal = -Int(-(rtkSingleNormal * 2#)) / 2#   ' round up to 0.5
    End If

    ' If we OPEN with Surprise instead of a normal swing:
    ' 1 surprise round + rounded-up remainder at normal DPR
    rtkSingleSurp = 1# + ccr_SafeDiv(ccr_Max(0#, hpPerMob - nSurpriseDamageOut), ccr_Max(1#, nDamageOut), 0#)
    rtkSingleSurp = -Int(-(rtkSingleSurp * 2#)) / 2#           ' round up to 0.5

    ' Positive => Surprise path is better (fewer rounds). Negative => worse.
    deltaFirst = rtkSingleNormal - rtkSingleSurp

    If deltaFirst <> 0# Then
        ' Regen attenuation: big regen vs DPR shrinks the effect either way
        regenPerRound = nMobHPRegen / 6#
        regenRatio = ccr_SafeDiv(regenPerRound, ccr_Max(1#, nDamageOut), 0#)
        regenAtten = 1# - 0.45 * ccr_SmoothStep(0#, 0.6, regenRatio)   ' 0.55..1.00

        ' Pack fade (credit/penalty gets a bit messier as pack size rises)
        packFade = 1# / Sqr(ccr_Max(1#, nNumMobs))                      ' 1, ~0.71, ~0.58, …
        fadeGate = ccr_SmoothStep(3#, 8#, nNumMobs)

        ' Apply symmetrically to savings and penalties
        adj = deltaFirst * regenAtten * ccr_Lerp(1#, packFade, fadeGate)

        ' Reduce RTK if adj>0 (surprise better), or increase if adj<0 (surprise worse)
        CalcCombatRounds.nRTK = ccr_Max(nNumMobs, CalcCombatRounds.nRTK - adj)
    End If
End If
' ===== end surprise opener credit =====

If nNumMobs > 1 And CalcCombatRounds.nRTK > 0 And CalcCombatRounds.nRTK < nNumMobs Then CalcCombatRounds.nRTK = nNumMobs
    
If CalcCombatRounds.nRTK > 0 And CalcCombatRounds.nRTK < 1 Then CalcCombatRounds.nRTK = 1

If nMobHealth > 1# And (CalcCombatRounds.nRTK < 1 Or CalcCombatRounds.nRTK > 200) Then
    CalcCombatRounds.sRTK = "<infinitely attacking>"
ElseIf CalcCombatRounds.nRTK > 0 Then
    CalcCombatRounds.sRTK = Round(CalcCombatRounds.nRTK, 1) & IIf(nNumMobs > 1, " RTC", " RTK")
End If

If CalcCombatRounds.nRTD > 0 And CalcCombatRounds.nRTD < 200 Then
    CalcCombatRounds.sRTD = "vs " & CalcCombatRounds.nRTD & " RTD"
ElseIf (CalcCombatRounds.nRTD = 0 Or CalcCombatRounds.nRTD >= 200) And nMobDamage >= 0 And nCharHealth > 0 Then
    CalcCombatRounds.sRTD = "vs <unfazed by damage>"
End If

If Len(CalcCombatRounds.sRTK & CalcCombatRounds.sRTD) > 0 Then
    If CalcCombatRounds.nRTD > 0 And CalcCombatRounds.nRTK >= 1 Then
        CalcCombatRounds.nSuccess = Round((CalcCombatRounds.nRTD ^ 2) / ((CalcCombatRounds.nRTK ^ 2) + (CalcCombatRounds.nRTD ^ 2)) * 100)
        
        If CalcCombatRounds.nSuccess >= 95 Then
            CalcCombatRounds.sSuccess = " - certain success"
        ElseIf CalcCombatRounds.nSuccess >= 5 Then
            CalcCombatRounds.sSuccess = " - " & CalcCombatRounds.nSuccess & "% chance of success"
        Else
            CalcCombatRounds.sSuccess = " - certain failure"
        End If
    End If
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalcCombatRounds")
Resume out:
End Function

'--------------------- local helpers (collision-safe) ----------------------
Private Function ccr_Saturate(ByVal x As Double) As Double
    If x <= 0# Then
        ccr_Saturate = 0#
    ElseIf x >= 1# Then
        ccr_Saturate = 1#
    Else
        ccr_Saturate = x
    End If
End Function

Private Function ccr_SmoothStep(ByVal edge0 As Double, ByVal edge1 As Double, ByVal x As Double) As Double
    If edge0 = edge1 Then
        ccr_SmoothStep = IIf(x >= edge1, 1#, 0#)
        Exit Function
    End If
    Dim t As Double: t = ccr_Saturate((x - edge0) / (edge1 - edge0))
    ccr_SmoothStep = t * t * (3# - 2# * t)
End Function

Private Function ccr_Lerp(ByVal a As Double, ByVal b As Double, ByVal t As Double) As Double
    ccr_Lerp = a + (b - a) * t
End Function

Private Function ccr_SafeDiv(ByVal n As Double, ByVal d As Double, Optional ByVal def As Double = 0#) As Double
    If d = 0# Then ccr_SafeDiv = def Else ccr_SafeDiv = n / d
End Function

Private Function ccr_Min(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then ccr_Min = a Else ccr_Min = b
End Function

Private Function ccr_Max(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then ccr_Max = a Else ccr_Max = b
End Function

Public Function CalcExpNeeded(ByVal startlevel As Long, ByVal exptable As Long) As Currency
'FROM: https://www.mudinfo.net/viewtopic.php?p=7703
On Error GoTo error:
Dim nModifiers() As Integer, i As Long, j As Currency, k As Currency, exp_multiplier As Long, exp_divisor As Long, Ret() As Currency
Dim lastexp As Currency, startexp As Currency, running_exp_tabulation As Currency, billions_tabulator As Currency
Dim potential_new_exp As Currency, ALTERNATE_NEW_EXP As Currency
Dim MAX_UINT As Double, numlevels As Integer, num_divides As Integer

MAX_UINT = 4294967295#
numlevels = 1

ReDim Ret(startlevel To (startlevel + numlevels - 1))

For i = 1 To (startlevel + numlevels - 1)
    startexp = lastexp
    If i = 1 Then
        running_exp_tabulation = 0
    ElseIf i = 2 Then
        running_exp_tabulation = exptable * 10
    Else
        If i <= 26 Then 'levels 1-26
            nModifiers() = GetExpModifiers(i)
            exp_multiplier = nModifiers(0)
            exp_divisor = nModifiers(1)
        ElseIf i <= 55 Then 'levels 27-55
            exp_multiplier = 115
            exp_divisor = 100
        ElseIf i <= 58 Then 'levels 56-58
            exp_multiplier = 109
            exp_divisor = 100
        Else 'levels 59+
            exp_multiplier = 108
            exp_divisor = 100
        End If
        
        If i = 97 Then
            'Debug.Print i
        End If
        
        If exp_multiplier = 0 Or exp_divisor = 0 Then
            potential_new_exp = 0
        Else
            potential_new_exp = running_exp_tabulation * exp_multiplier
        End If
        
        If potential_new_exp > MAX_UINT Then 'UINT ROLLOVER #1
            num_divides = 0
            Do While potential_new_exp > MAX_UINT
                running_exp_tabulation = Fix(running_exp_tabulation / 100)
                potential_new_exp = running_exp_tabulation * exp_multiplier
                num_divides = num_divides + 1
            Loop
            If num_divides > 1 Then
                ALTERNATE_NEW_EXP = Fix((running_exp_tabulation * exp_multiplier * 100) / exp_divisor)
            Else
                ALTERNATE_NEW_EXP = Fix(potential_new_exp / exp_divisor)
            End If
            Do While num_divides > 0
                ALTERNATE_NEW_EXP = ALTERNATE_NEW_EXP * 100
                num_divides = num_divides - 1
            Loop
        Else
            ALTERNATE_NEW_EXP = Fix(potential_new_exp / exp_divisor)
        End If
        
        j = (1000000 * exp_multiplier * billions_tabulator)
        Do While j > MAX_UINT
            j = j - MAX_UINT - 1 'UINT ROLLOVER #2
        Loop
        Do While j >= 1000000000
            j = j - 1000000000
            billions_tabulator = billions_tabulator + 1
        Loop
        
        k = (j + ALTERNATE_NEW_EXP)
        Do While k >= 1000000000
            k = k - 1000000000
            billions_tabulator = billions_tabulator + 1
        Loop
        
        running_exp_tabulation = k
    End If
    
    lastexp = running_exp_tabulation + (billions_tabulator * 1000000000)
    
    If i >= startlevel Then
        Ret(i) = lastexp
    End If
Next i

CalcExpNeeded = Ret(startlevel)

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalcExpNeeded")
Resume out:
End Function

Private Function GetExpModifiers(ByVal nLevel As Integer) As Integer()
Dim Ret(1) As Integer
Ret(0) = 0
Ret(1) = 0

Select Case nLevel
    Case 3:
        Ret(0) = 40
        Ret(1) = 20
        'return [40, 20];
    Case 4, 5:
        Ret(0) = 44
        Ret(1) = 24
        'return [44, 24];
    Case 6, 7:
        Ret(0) = 48
        Ret(1) = 28
        'return [48, 28];
    Case 8, 9:
        Ret(0) = 52
        Ret(1) = 32
        'return [52, 32];
    Case 10, 11:
        Ret(0) = 56
        Ret(1) = 36
        'return [56, 36];
    Case 12, 13:
        Ret(0) = 60
        Ret(1) = 40
        'return [60, 40];
    Case 14, 15:
        Ret(0) = 65
        Ret(1) = 45
        'return [65, 45];
    Case 16, 17:
        Ret(0) = 70
        Ret(1) = 50
        'return [70, 50];
    Case 18:
        Ret(0) = 75
        Ret(1) = 55
        'return [75, 55];
    Case Else:
        If nLevel <= 26 Then
            Ret(0) = 50
            Ret(1) = 40
            'return [50, 40];
        Else
            Ret(0) = 23
            Ret(1) = 20
            'return [23, 20];
        End If

End Select

GetExpModifiers = Ret

'function GetExpModifiers($iLevel) {
'    switch ($iLevel) {
'        Case 3:
'            return [40, 20];
'        Case 4:
'        Case 5:
'            return [44, 24];
'        Case 6:
'        Case 7:
'            return [48, 28];
'        Case 8:
'        Case 9:
'            return [52, 32];
'        Case 10:
'        Case 11:
'            return [56, 36];
'        Case 12:
'        Case 13:
'            return [60, 40];
'        Case 14:
'        Case 15:
'            return [65, 45];
'        Case 16:
'        Case 17:
'            return [70, 50];
'        Case 18:
'            return [75, 55];
'default:
'            if ($iLevel <= 26) {
'                return [50, 40];
'            } else {
'                return [23, 20];
'            }
'    }
'}
End Function

Public Function CalculateSpellCast(ByVal nSpellNum As Long, Optional ByRef nCastLVL As Long, Optional ByVal nSpellcasting As Long, _
    Optional ByVal nVSMR As Long, Optional ByVal bVSAntiMagic As Boolean, _
    Optional ByVal nMaxMana As Long, Optional ByVal nManaRegenRate As Long, Optional ByVal bMeditate As Boolean, _
    Optional ByVal nSpellOverhead As Long) As tSpellCastValues
On Error GoTo error:
Dim x As Integer, y As Integer, tSpellMinMaxDur As SpellMinMaxDur, nDamage As Long, nHeals As Long
Dim nMinCast As Long, nMaxCast As Long, nSpellAvgCast As Long, nSpellDuration As Long, nFullResistChance As Integer
Dim nCastChance As Integer, bDamageMinusMR As Boolean, nCasts As Double ', nRoundTotal As Long
Dim sAvgRound As String, bLVLspecified As Boolean, sLVLincreases As String, sMMA As String
Dim nTemp As Long, nTemp2 As Long, sTemp As String, sTemp2 As String, sCastLVL As String, sAbil As String

If nSpellNum = 0 Then Exit Function

On Error GoTo seekit:
If tabSpells.Fields("Number") = nSpellNum Then GoTo ready:
GoTo seekit2:

seekit:
Resume seekit2:
seekit2:
On Error GoTo error:
tabSpells.Index = "pkSpells"
tabSpells.Seek "=", nSpellNum
If tabSpells.NoMatch = True Then
    tabSpells.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:
CalculateSpellCast.sSpellName = tabSpells.Fields("Name")

If nCastLVL <= 0 Then
    If nCastLVL < tabSpells.Fields("Cap") Then nCastLVL = tabSpells.Fields("Cap")
Else
    bLVLspecified = True
End If
If nCastLVL < tabSpells.Fields("ReqLevel") Then nCastLVL = tabSpells.Fields("ReqLevel")
If nCastLVL > tabSpells.Fields("Cap") And tabSpells.Fields("Cap") > 0 Then nCastLVL = tabSpells.Fields("Cap")
CalculateSpellCast.nCastLevel = nCastLVL

tSpellMinMaxDur = GetCurrentSpellMinMax(IIf(nCastLVL > 0, True, False), nCastLVL)

nMinCast = tSpellMinMaxDur.nMin
nMaxCast = tSpellMinMaxDur.nMax
nSpellDuration = tSpellMinMaxDur.nDur
If nSpellDuration < 1 Then nSpellDuration = 1
nSpellAvgCast = Round((nMinCast + nMaxCast) / 2)

If Not tabSpells.Fields("Number") = nSpellNum Then tabSpells.Seek "=", nSpellNum

If nSpellcasting > 0 And tabSpells.Fields("Diff") < 200 Then
    nCastChance = GetSpellCastChance(tabSpells.Fields("Diff"), nSpellcasting, IIf(tabSpells.Fields("Magery") = 5, True, False))
Else
    nCastChance = 100
End If

For x = 0 To 9
    Select Case tabSpells.Fields("Abil-" & x)
        Case 1: 'dmg
            CalculateSpellCast.bDoesDamage = True
            If tabSpells.Fields("AbilVal-" & x) <> 0 Then
                nDamage = nDamage + tabSpells.Fields("AbilVal-" & x)
            Else
                nDamage = nDamage + nSpellAvgCast
            End If
            
        Case 8: 'drain
            CalculateSpellCast.bDoesDamage = True
            CalculateSpellCast.bDoesHeal = True
            If tabSpells.Fields("AbilVal-" & x) <> 0 Then
                nDamage = tabSpells.Fields("AbilVal-" & x)
                nHeals = tabSpells.Fields("AbilVal-" & x)
            Else
                nDamage = nDamage + nSpellAvgCast
                nHeals = nHeals + nSpellAvgCast
            End If
            
        Case 17: 'dmg-mr
            CalculateSpellCast.bDoesDamage = True
            If tabSpells.Fields("AbilVal-" & x) = 0 Then
                bDamageMinusMR = True
                nTemp = nSpellAvgCast
            Else
                nTemp = tabSpells.Fields("AbilVal-" & x)
            End If
            
            If nVSMR Then
                nTemp2 = CalculateResistDamage(nTemp, nVSMR, tabSpells.Fields("TypeOfResists"), True, False, bVSAntiMagic, 0)
                CalculateSpellCast.nDamageResisted = CalculateSpellCast.nDamageResisted + (nTemp - nTemp2)
                nDamage = nDamage + nTemp2
            Else
                nDamage = nDamage + nTemp
            End If
            
        Case 18: 'healing
            CalculateSpellCast.bDoesHeal = True
            If tabSpells.Fields("AbilVal-" & x) <> 0 Then
                nHeals = tabSpells.Fields("AbilVal-" & x)
            Else
                nHeals = nHeals + nSpellAvgCast
            End If
            
    End Select
Next x

If nVSMR > 0 Then
    If bDamageMinusMR Then
        nMinCast = CalculateResistDamage(nMinCast, nVSMR, tabSpells.Fields("TypeOfResists"), bDamageMinusMR, False, bVSAntiMagic, 0)
        nMaxCast = CalculateResistDamage(nMaxCast, nVSMR, tabSpells.Fields("TypeOfResists"), bDamageMinusMR, False, bVSAntiMagic, 0)
        nSpellAvgCast = Round((nMinCast + nMaxCast) / 2)
    End If
    
    If tabSpells.Fields("TypeOfResists") = 2 Or (tabSpells.Fields("TypeOfResists") = 1 And bVSAntiMagic) Then
        nFullResistChance = Fix(nVSMR / 2)
        If nFullResistChance > 98 Then nFullResistChance = 98
    End If
End If

If tabSpells.Fields("EnergyCost") > 0 And tabSpells.Fields("EnergyCost") <= 500 Then
    nCasts = Fix(1000 / tabSpells.Fields("EnergyCost"))
Else
    nCasts = 1
End If

CalculateSpellCast.nMinCast = nMinCast 'Fix(nMinCast / nCasts)
CalculateSpellCast.nMaxCast = nMaxCast 'Fix(nMaxCast / nCasts)
CalculateSpellCast.nAvgCast = nSpellAvgCast 'Fix(nSpellAvgCast / nCasts)
CalculateSpellCast.nNumCasts = nCasts
CalculateSpellCast.nManaCost = tabSpells.Fields("ManaCost") * nCasts
CalculateSpellCast.nCastChance = nCastChance
CalculateSpellCast.nAvgRoundDmg = Round(((nDamage * nCasts) * (nCastChance / 100#)) * (1# - (nFullResistChance / 100#)))
CalculateSpellCast.nAvgRoundHeals = Round(((nHeals * nCasts) * (nCastChance / 100#)) * (1# - (nFullResistChance / 100#)))
CalculateSpellCast.nDuration = nSpellDuration
CalculateSpellCast.nFullResistChance = nFullResistChance

If CalculateSpellCast.nManaCost > 0 And CalculateSpellCast.nManaCost <= nMaxMana Then  'and (CalculateSpellCast.bDoesDamage Or CalculateSpellCast.bDoesHeal)
    CalculateSpellCast.nOOM = CalcRoundsToOOM(CalculateSpellCast.nManaCost + nSpellOverhead, nMaxMana, nManaRegenRate, nCastChance, nSpellDuration)
End If

If CalculateSpellCast.nDamageResisted > 0 Then
    If nDamage = 0 Then
        CalculateSpellCast.nDamageResisted = 100
    Else
        nTemp = CalculateSpellCast.nDamageResisted
        CalculateSpellCast.nDamageResisted = Round((nTemp / (nDamage + nTemp)) * 100)
    End If
End If

'===========================

If CalculateSpellCast.bDoesDamage Or CalculateSpellCast.bDoesHeal Then
    If Not bLVLspecified And nCastLVL > 0 Then sCastLVL = "(@lvl " & nCastLVL & ") "
    If CalculateSpellCast.bDoesDamage And CalculateSpellCast.bDoesHeal Then
        sAvgRound = sCastLVL & IIf(nSpellDuration > 1, nSpellAvgCast, CalculateSpellCast.nAvgRoundDmg) & " damage + " & CalculateSpellCast.nAvgRoundHeals & " heals/round"
    ElseIf CalculateSpellCast.bDoesDamage Then
        sAvgRound = sCastLVL & IIf(nSpellDuration > 1, nSpellAvgCast, CalculateSpellCast.nAvgRoundDmg) & " damage/round"
    ElseIf CalculateSpellCast.bDoesHeal Then
        sAvgRound = sCastLVL & CalculateSpellCast.nAvgRoundHeals & " healing/round"
    End If
    
    If nSpellDuration > 1 Then
        sAvgRound = sAvgRound & " for " & nSpellDuration & " rounds (" & ((CalculateSpellCast.nAvgRoundDmg + CalculateSpellCast.nAvgRoundHeals) * nSpellDuration) & " total)"
        If CalculateSpellCast.nDamageResisted > 0 Then sAvgRound = sAvgRound & " after " & CalculateSpellCast.nDamageResisted & "% damage resisted"
        sTemp = ""
        If bLVLspecified And nCastChance < 100 Then sTemp = AutoAppend(sTemp, (100 - nCastChance) & "% chance to fail cast", " and ")
        If CalculateSpellCast.nFullResistChance > 0 Then sTemp = AutoAppend(sTemp, CalculateSpellCast.nFullResistChance & "% chance to fully-resist", " and ")
        If Not sTemp = "" Then sAvgRound = sAvgRound & ", not including " & sTemp
    Else
        If bLVLspecified And nCastChance < 100 Then sAvgRound = sAvgRound & " @ " & nCastChance & "% chance to cast"
        If CalculateSpellCast.nDamageResisted > 0 Then sAvgRound = sAvgRound & ", " & CalculateSpellCast.nDamageResisted & "% damage resisted"
        If CalculateSpellCast.nFullResistChance > 0 Then sAvgRound = sAvgRound & ", " & CalculateSpellCast.nFullResistChance & "% chance to fully-resist"
    End If
End If

If CalculateSpellCast.nMinCast > 0 And (CalculateSpellCast.nMinCast <> CalculateSpellCast.nMaxCast Or CalculateSpellCast.nMaxCast <> nSpellAvgCast) Then
    If Not bLVLspecified And nCastLVL > 0 Then sCastLVL = " (@lvl " & nCastLVL & ")"
    sMMA = "Min/Max/Avg Cast" & sCastLVL & ": " & CalculateSpellCast.nMinCast & "/" & CalculateSpellCast.nMaxCast & "/" & nSpellAvgCast
    If CalculateSpellCast.nNumCasts > 1 Then sMMA = sMMA & " x" & CalculateSpellCast.nNumCasts & "/round"
    If bLVLspecified And nSpellDuration = 1 Then
        If CalculateSpellCast.nFullResistChance > 0 And nCastChance < 100 Then
            sMMA = sMMA & " (before full resist & cast % reductions)"
        ElseIf CalculateSpellCast.nFullResistChance > 0 Then
            sMMA = sMMA & " (before full resist reduction)"
        ElseIf nCastChance < 100 Then
            sMMA = sMMA '& " (before cast % reduction)"
        End If
    End If
End If

If (tabSpells.Fields("Cap") = 0 Or tabSpells.Fields("Cap") > tabSpells.Fields("ReqLevel")) _
    And ((tabSpells.Fields("MinInc") > 0 And tabSpells.Fields("MinIncLVLs") > 0) _
        Or (tabSpells.Fields("MaxInc") > 0 And tabSpells.Fields("MaxIncLVLs") > 0) _
        Or (tabSpells.Fields("DurInc") > 0 And tabSpells.Fields("DurIncLVLs") > 0)) Then

    sTemp = ""
    sTemp2 = ""
    y = 0
    For x = 0 To 9
        If tabSpells.Fields("Abil-" & x) > 0 Then
            Select Case tabSpells.Fields("Abil-" & x)
                Case 23, 51, 52, 80, 97, 98, 100, 108 To 113, 115, 119, 122, 138, 144, 151, 164, 178:
                    'ignore:
                    '23 - effectsundead
                    '51: 'anti magic
                    '52: 'evil in combat
                    '80: 'effects animal
                    '97-98 - good/evil only
                    '100: 'loyal
                    '108: 'effects living
                    '109 To 113: 'nonliving, notgood, notevil, neutral, not neutral
                    '112 - neut only
                    '115 - descmsg
                    '119: 'del@main
                    '122: removespell
                    '138: 'roomvis
                    '144: 'non magic spell
                    '151,164: endcast, endcast%
                    '178: shadowform
                Case 137:
                    '137-shock... really ignore
                Case Else:
                    sAbil = GetAbilityStats(tabSpells.Fields("Abil-" & x), 0, , False)
                    If Len(sAbil) > 0 Then
                        y = y + 1
                        If tabSpells.Fields("AbilVal-" & x) = 0 Then
                            sTemp = AutoAppend(sTemp, sAbil)
                        Else
                            sTemp2 = AutoAppend(sTemp2, sAbil)
                        End If
                    End If
            End Select
        End If
    Next x
    If Not sTemp = "" Then
        If Not sTemp2 = "" Then sTemp = sTemp & " (not " & sTemp2 & ")"
        sTemp2 = ""
        tSpellMinMaxDur = GetCurrentSpellMinMax(False)
        If CStr(tSpellMinMaxDur.nDur) <> tSpellMinMaxDur.sDur Then sTemp2 = AutoAppend(sTemp2, "Duration: " & tSpellMinMaxDur.sDur)
        If CStr(tSpellMinMaxDur.nMin) <> tSpellMinMaxDur.sMin Then sTemp2 = AutoAppend(sTemp2, "Min: " & tSpellMinMaxDur.sMin)
        If CStr(tSpellMinMaxDur.nMax) <> tSpellMinMaxDur.sMax Then sTemp2 = AutoAppend(sTemp2, "Max: " & tSpellMinMaxDur.sMax)
        If Not sTemp2 = "" Then
            sLVLincreases = "LVL Increases: " & sTemp2
            If y > 1 Then sLVLincreases = sLVLincreases & " for: " & sTemp
        End If
    End If
End If

If Len(sAvgRound) > 0 Then CalculateSpellCast.sAvgRound = sAvgRound
If Len(sMMA) > 0 Then CalculateSpellCast.sMMA = sMMA
If Len(sLVLincreases) > 0 Then CalculateSpellCast.sLVLincreases = sLVLincreases

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalculateSpellCast")
Resume out:
End Function

Public Function CalculateAttack(tCharStats As tCharacterProfile, ByVal nAttackTypeMUD As eAttackTypeMUD, Optional ByVal nWeaponNumber As Long, _
    Optional ByVal bAbil68Slow As Boolean, Optional ByVal nSpeedAdj As Integer = 100, Optional ByVal nVSAC As Long, Optional ByVal nVSDR As Long, _
    Optional ByVal nVSDodge As Long, Optional ByRef sCasts As String = "", Optional ByVal bForceCalc As Boolean, _
    Optional ByVal nSpecifyDamage As Double = -1, Optional ByVal nSpecifyAccy As Double = -1, _
    Optional ByVal nSecondaryDefense As Long) As tAttackDamage
On Error GoTo error:
Dim x As Integer, nAvgHit As Currency, nPlusMaxDamage As Integer, nCritChance As Integer, nAvgCrit As Long
Dim nPercent As Double, nDurDamage As Currency, nDurCount As Integer, nTemp As Long, nPlusMinDamage As Integer
Dim tMatches() As RegexMatches, sRegexPattern As String, sAttackDetail As String, nTemp2 As Long
Dim sArr() As String, iMatch As Integer, nExtraTMP As Currency, nExtraAvgSwing As Currency, nCount As Integer, nExtraPCT As Double
Dim nEncum As Currency, nEnergy As Long, nCombat As Currency, nQnDBonus As Currency, nSwings As Double, nExtraAvgHit As Currency
Dim nMinCrit As Long, nMaxCrit As Long, nStrReq As Integer, nAttackAccuracy As Currency, nPercent2 As Double
Dim nDmgMin As Long, nDmgMax As Long, nAttackSpeed As Integer, nMAPlusAccy(1 To 3) As Long, nMAPlusDmg(1 To 3) As Long, nMAPlusSkill(1 To 3) As Integer
Dim nLevel As Integer, nStrength As Integer, nAgility As Integer, nPlusBSaccy As Integer, nPlusBSmindmg As Integer, nPlusBSmaxdmg As Integer
Dim nStealth As Integer, bClassStealth As Boolean, bRaceStealth As Boolean, nHitChance As Currency
Dim tStatIndex As TypeGetEquip, tRet As tAttackDamage, accTemp As Long, nDefense() As Long
Dim nPreRollMinModifier As Double, nPreRollMaxModifier As Double, nDamageMultiplierMin As Double, nDamageMultiplierMax As Double

nPreRollMinModifier = 1
nPreRollMaxModifier = 1
nDamageMultiplierMin = 1
nDamageMultiplierMax = 1

'nAttackTypeMUD:
'1-punch, 2-kick, 3-jumpkick
'4-surprise, 5-normal, 6-bash, 7-smash
If nSpecifyDamage >= 0 Then
    tRet.sAttackDesc = "Manual"
    nDmgMin = nSpecifyDamage
    nDmgMax = nSpecifyDamage
    If nDmgMin < 0 Then nDmgMin = 0
    If nDmgMax > 9999999 Then nDmgMax = 9999999
    If nDmgMin > nDmgMax Then nDmgMin = nDmgMax
    nSwings = 1
    GoTo calc_damage:
End If
If nAttackTypeMUD <= 0 Then nAttackTypeMUD = a5_Normal
If nWeaponNumber = 0 And nAttackTypeMUD > a5_Normal Then Exit Function 'bash/smash

If tCharStats.nLevel = 0 Then
    nLevel = 255
    nCombat = 3
    nStrength = 255
    nAgility = 255
    nStealth = 255
    If nAttackTypeMUD >= a1_Punch And nAttackTypeMUD <= a3_Jumpkick Then nMAPlusSkill(nAttackTypeMUD) = 1
    nAttackAccuracy = 999
    nPlusBSaccy = 999
    bClassStealth = True
Else
    nLevel = tCharStats.nLevel
    nCombat = tCharStats.nCombat
    nStrength = tCharStats.nSTR
    nAgility = tCharStats.nAGI
    nPlusMinDamage = tCharStats.nPlusMinDamage
    nPlusMaxDamage = tCharStats.nPlusMaxDamage
    nStealth = tCharStats.nStealth
    nCritChance = tCharStats.nCrit
    If tCharStats.bIsLoadedCharacter Then nCritChance = nCritChance - nGlobalCharQnDbonus
    nMAPlusSkill(1) = tCharStats.nMAPlusSkill(1)
    nMAPlusAccy(1) = tCharStats.nMAPlusAccy(1)
    nMAPlusDmg(1) = tCharStats.nMAPlusDmg(1)
    nMAPlusSkill(2) = tCharStats.nMAPlusSkill(2)
    nMAPlusAccy(2) = tCharStats.nMAPlusAccy(2)
    nMAPlusDmg(2) = tCharStats.nMAPlusDmg(2)
    nMAPlusSkill(3) = tCharStats.nMAPlusSkill(3)
    nMAPlusAccy(3) = tCharStats.nMAPlusAccy(3)
    nMAPlusDmg(3) = tCharStats.nMAPlusDmg(3)
    nAttackAccuracy = tCharStats.nAccuracy
    nPlusBSaccy = tCharStats.nPlusBSaccy
    nPlusBSmindmg = tCharStats.nPlusBSmindmg
    nPlusBSmaxdmg = tCharStats.nPlusBSmaxdmg
    nEncum = tCharStats.nEncumPCT
    
    If nCombat = 0 And tCharStats.nClass > 0 Then nCombat = GetClassCombat(tCharStats.nClass)
    
    bClassStealth = tCharStats.bClassStealth
    bRaceStealth = tCharStats.bRaceStealth
    If Not bClassStealth And tCharStats.nClass > 0 Then bClassStealth = GetClassStealth(tCharStats.nClass)
    If Not bRaceStealth And tCharStats.nRace > 0 Then bRaceStealth = GetRaceStealth(tCharStats.nRace)
        
    If bClassStealth = False And bForceCalc = True Then
        nStealth = CalculateStealth(nLevel, nAgility, tCharStats.nINT, tCharStats.nCHA, False, True, nStealth)
    ElseIf nStealth = 0 And (bClassStealth Or bRaceStealth) Then
        nStealth = CalculateStealth(nLevel, nAgility, tCharStats.nINT, tCharStats.nCHA, bClassStealth, bRaceStealth)
    End If
    
    'force calc punch/kick/jumpkick:
    If bForceCalc And nAttackTypeMUD >= a1_Punch And nAttackTypeMUD <= a3_Jumpkick Then
        If nMAPlusSkill(nAttackTypeMUD) < 1 Then nMAPlusSkill(nAttackTypeMUD) = 1
    End If
End If

If nWeaponNumber = 0 Then GoTo non_weapon_attack:

On Error GoTo seek2:
If tabItems.Fields("Number") = nWeaponNumber Then GoTo item_ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabItems.Index = "pkItems"
tabItems.Seek "=", nWeaponNumber
If tabItems.NoMatch = True Then
    tabItems.MoveFirst
    Exit Function
End If

item_ready:
On Error GoTo error:

If tCharStats.bIsLoadedCharacter And nWeaponNumber > 0 And nWeaponNumber <> nGlobalCharWeaponNumber(0) Then
    'current weapon is different than this weapon...
    If tabItems.Fields("WeaponType") = 1 Or tabItems.Fields("WeaponType") = 3 Then
        '+this weapon is two-handed...
        If nGlobalCharWeaponNumber(1) > 0 Then
            '+off-hand currently equipped. subtract those stats too...
            nAttackAccuracy = nAttackAccuracy - nGlobalCharWeaponAccy(1)
            nCritChance = nCritChance - nGlobalCharWeaponCrit(1)
            nPlusMaxDamage = nPlusMaxDamage - nGlobalCharWeaponMaxDmg(1)
            nPlusBSaccy = nPlusBSaccy - nGlobalCharWeaponBSaccy(1)
            nPlusBSmindmg = nPlusBSmindmg - nGlobalCharWeaponBSmindmg(1)
            nPlusBSmaxdmg = nPlusBSmaxdmg - nGlobalCharWeaponBSmaxdmg(1)
            nStealth = nStealth - nGlobalCharWeaponStealth(1)
            Select Case nAttackTypeMUD
                Case 1: 'Punch
                    nMAPlusSkill(1) = nMAPlusSkill(1) - nGlobalCharWeaponPunchSkill(1)
                    nMAPlusAccy(1) = nMAPlusAccy(1) - nGlobalCharWeaponPunchAccy(1)
                    nMAPlusDmg(1) = nMAPlusDmg(1) - nGlobalCharWeaponPunchDmg(1)
                Case 2: 'Kick
                    nMAPlusSkill(2) = nMAPlusSkill(2) - nGlobalCharWeaponKickSkill(1)
                    nMAPlusAccy(2) = nMAPlusAccy(2) - nGlobalCharWeaponKickAccy(1)
                    nMAPlusDmg(2) = nMAPlusDmg(2) - nGlobalCharWeaponKickDmg(1)
                Case 3: 'Jumpkick
                    nMAPlusSkill(3) = nMAPlusSkill(3) - nGlobalCharWeaponJkSkill(1)
                    nMAPlusAccy(3) = nMAPlusAccy(3) - nGlobalCharWeaponJkAccy(1)
                    nMAPlusDmg(3) = nMAPlusDmg(3) - nGlobalCharWeaponJkDmg(1)
            End Select
        End If
    End If
    
    'now add in current item's stats...
    If nAttackTypeMUD > a3_Jumpkick Then
        'weapon accuracy does not count towards mystic attacks
        nAttackAccuracy = nAttackAccuracy + tabItems.Fields("Accy")
    End If
    
    For x = 0 To 19
        If tabItems.Fields("Abil-" & x) > 0 And tabItems.Fields("AbilVal-" & x) <> 0 Then
            
            tStatIndex = InvenGetEquipInfo(tabItems.Fields("Abil-" & x), tabItems.Fields("AbilVal-" & x))
            If Not tabItems.Fields("Number") = nWeaponNumber Then tabItems.Seek "=", nWeaponNumber
            
            If tStatIndex.nEquip > 0 Then
                Select Case tStatIndex.nEquip
                    Case 7: nCritChance = nCritChance + tabItems.Fields("AbilVal-" & x)
                    Case 11: nPlusMaxDamage = nPlusMaxDamage + tabItems.Fields("AbilVal-" & x)
                    Case 13: nPlusBSaccy = nPlusBSaccy + tabItems.Fields("AbilVal-" & x)
                    Case 14: nPlusBSmindmg = nPlusBSmindmg + tabItems.Fields("AbilVal-" & x)
                    Case 15: nPlusBSmaxdmg = nPlusBSmaxdmg + tabItems.Fields("AbilVal-" & x)
                    Case 37: If nAttackTypeMUD = a1_Punch Then nMAPlusSkill(1) = nMAPlusSkill(1) + tabItems.Fields("AbilVal-" & x) 'pu
                    Case 40: If nAttackTypeMUD = a1_Punch Then nMAPlusAccy(1) = nMAPlusAccy(1) + tabItems.Fields("AbilVal-" & x) 'pu
                    Case 34: If nAttackTypeMUD = a1_Punch Then nMAPlusDmg(1) = nMAPlusDmg(1) + tabItems.Fields("AbilVal-" & x) 'pu
                    Case 38: If nAttackTypeMUD = a2_Kick Then nMAPlusSkill(2) = nMAPlusSkill(2) + tabItems.Fields("AbilVal-" & x) 'kick
                    Case 41: If nAttackTypeMUD = a2_Kick Then nMAPlusAccy(2) = nMAPlusAccy(2) + tabItems.Fields("AbilVal-" & x) 'kick
                    Case 35: If nAttackTypeMUD = a2_Kick Then nMAPlusDmg(2) = nMAPlusDmg(2) + tabItems.Fields("AbilVal-" & x) 'kick
                    Case 39: If nAttackTypeMUD = a3_Jumpkick Then nMAPlusSkill(3) = nMAPlusSkill(3) + tabItems.Fields("AbilVal-" & x) 'jk
                    Case 42: If nAttackTypeMUD = a3_Jumpkick Then nMAPlusAccy(3) = nMAPlusAccy(3) + tabItems.Fields("AbilVal-" & x) 'jk
                    Case 36: If nAttackTypeMUD = a3_Jumpkick Then nMAPlusDmg(3) = nMAPlusDmg(3) + tabItems.Fields("AbilVal-" & x) 'jk
                    Case 19: nStealth = nStealth + tabItems.Fields("AbilVal-" & x)
                End Select
            End If
        End If
    Next x
End If

If nAttackTypeMUD <= a3_Jumpkick Then GoTo non_weapon_attack:

tRet.sAttackDesc = tabItems.Fields("Name")
nStrReq = tabItems.Fields("StrReq")
nDmgMin = tabItems.Fields("Min")
nDmgMax = tabItems.Fields("Max")
nAttackSpeed = tabItems.Fields("Speed")
If bAbil68Slow Then nAttackSpeed = Fix((nAttackSpeed * 3) / 2)

GoTo calc_energy:

non_weapon_attack:
If nAttackTypeMUD <= a3_Jumpkick Then
    If nMAPlusSkill(nAttackTypeMUD) <= 0 Then Exit Function
End If
tRet.sAttackDesc = "Punch"

Select Case nAttackTypeMUD
    Case 1: 'Punch
        nAttackSpeed = 1150
        If bAbil68Slow Then nAttackSpeed = 1750
    Case 2: 'Kick
        nAttackSpeed = 1400
        If bAbil68Slow Then nAttackSpeed = 2000
    Case 3: 'Jumpkick
        If bGreaterMUD Then
            nAttackSpeed = 2900
            If bAbil68Slow Then nAttackSpeed = 4045
            nAttackSpeed = 1900
            If bAbil68Slow Then nAttackSpeed = 2650
        End If
    Case 4, 5: 'surprise/normal Punch
        nAttackSpeed = 1200
        If bAbil68Slow Then nAttackSpeed = 1800
    Case Else:
        Exit Function
End Select

If nAttackTypeMUD < a4_Surprise Then
    
    If bGreaterMUD Then
        If nLevel < 20 Then
            nTemp = (nLevel / 8) + 2
        Else
            nTemp = (nLevel / 6)
            If nTemp < 5 Then nTemp = 5
        End If
        nDmgMin = nTemp + nMAPlusSkill(nAttackTypeMUD)
        
        nTemp = 0
        Select Case nAttackTypeMUD
            Case 1: 'Punch
                If nLevel < 20 Then
                    nTemp = ((nLevel + 3) / 4) + 6
                Else
                    nTemp = (nLevel / 4)
                    If nTemp < 12 Then nTemp = 12
                End If
            Case 2: 'Kick
                If nLevel < 20 Then
                    nTemp = (nLevel / 5) + 7
                Else
                    nTemp = (nLevel / 4)
                    If nTemp < 10 Then nTemp = 10
                End If
            Case 3: 'Jumpkick
                If nLevel < 20 Then
                    nTemp = (nLevel / 6) + 7
                Else
                    nTemp = (nLevel / 4)
                    If nTemp < 10 Then nTemp = 10
                End If
        End Select
        nDmgMax = nTemp + nMAPlusSkill(nAttackTypeMUD)
    Else
        nTemp = nLevel
        If nTemp > 20 Then nTemp = 20
        
        nDmgMin = nMAPlusSkill(nAttackTypeMUD) * nTemp
        If nDmgMin < 0 Then nDmgMin = nDmgMin + 7 'it's in the dll... not sure why as this would only happen is skill was < 0, but just in case.
        nDmgMin = Fix(nDmgMin / 8) + 2
        
        Select Case nAttackTypeMUD
            Case 1: 'Punch
                nDmgMax = nMAPlusSkill(nAttackTypeMUD) * (nTemp + 3)
                If nDmgMax < 0 Then nDmgMax = nDmgMax + 3 'it's in the dll... not sure why as this would only happen is skill was < 0, but just in case.
                nDmgMax = Fix(nDmgMax / 4) + 6
            Case 2: 'Kick
                nDmgMax = nMAPlusSkill(nAttackTypeMUD) * nTemp
                nDmgMax = Fix(nDmgMax / 6) + 7
            Case 3: 'Jumpkick
                nDmgMax = nMAPlusSkill(nAttackTypeMUD) * nTemp
                nDmgMax = Fix(nDmgMax / 6) + 8
        End Select
    End If
    
Else 'attacking without +punch or without a weapon
    nDmgMin = 1
    nDmgMax = 4
End If

calc_energy:
If nAttackTypeMUD = a4_Surprise Or nAttackTypeMUD = a7_Smash Then 'backstab, smash
    nEnergy = 1000
    nSwings = 1
Else
    nEnergy = CalcEnergyUsed(nCombat, nLevel, nAttackSpeed, nAgility, nStrength, nEncum, nStrReq, nSpeedAdj, IIf(nAttackTypeMUD = a4_Surprise, True, False))
End If

If tCharStats.bIsLoadedCharacter And nStrength >= nStrReq And Not nAttackTypeMUD = a4_Surprise And Not nAttackTypeMUD = a6_Bash And Not nAttackTypeMUD = a7_Smash Then
    nQnDBonus = CalcQuickAndDeadlyBonus(nAgility, nEnergy, nEncum)
    nCritChance = nCritChance + nQnDBonus
End If
If nCritChance > 40 Then
    If bGreaterMUD Then
        If nCritChance > 65 Then nCritChance = 65
    Else
        nCritChance = (40 + Fix((nCritChance - 40) / 3)) 'diminishing returns
        If nCritChance > 99 Then nCritChance = 99
    End If
End If

If nAttackTypeMUD = a6_Bash Then nEnergy = nEnergy * 2 'bash
If nEnergy < 200 Then nEnergy = 200
If nEnergy > 1000 Then nEnergy = 1000
nSwings = Round((1000 / nEnergy), 4)
If nSwings > 5 Then nSwings = 5

nDmgMin = nDmgMin + nPlusMinDamage
nDmgMax = nDmgMax + nPlusMaxDamage
If nDmgMin > nDmgMax Then nDmgMin = nDmgMax
If nDmgMin < 0 Then nDmgMin = 0
If nDmgMax < 0 Then nDmgMax = 0

If nAttackTypeMUD < a4_Surprise Then
    nAttackAccuracy = nAttackAccuracy + nMAPlusAccy(nAttackTypeMUD)
    nDmgMin = nDmgMin + nMAPlusDmg(nAttackTypeMUD)
    nDmgMax = nDmgMax + nMAPlusDmg(nAttackTypeMUD)
    If nAttackTypeMUD = a2_Kick Then 'kick
        If bGreaterMUD Then
            nDamageMultiplierMin = 1.33
            nDamageMultiplierMax = 1.33
            nAttackAccuracy = nAttackAccuracy - 10
        Else
            nPreRollMinModifier = 1.33
            nPreRollMaxModifier = 1.33
        End If
        tRet.sAttackDesc = "Kick"
    
    ElseIf nAttackTypeMUD = a3_Jumpkick Then 'jk
        If bGreaterMUD Then
            nDamageMultiplierMin = 1.66
            nDamageMultiplierMax = 1.66
             nAttackAccuracy = nAttackAccuracy - 15
        Else
            nPreRollMinModifier = 1.66
            nPreRollMaxModifier = 1.66
        End If
        tRet.sAttackDesc = "JumpKick"
    End If
    
ElseIf nAttackTypeMUD = a4_Surprise Then 'surprise
    If tRet.sAttackDesc = "Punch" Then
        tRet.sAttackDesc = "surprise punch"
    Else
        tRet.sAttackDesc = "backstab with " & tRet.sAttackDesc
    End If
    
    nCritChance = 0
    nQnDBonus = 0
    
    nTemp = (nLevel * 2) + Fix(nStealth / 10)
    nDmgMin = (nDmgMin * 2) + nTemp + nPlusBSmindmg
    nDmgMax = (nDmgMax * 2) + nTemp + nPlusBSmaxdmg
    
    If Not bClassStealth Then
        nDmgMin = Fix((nDmgMin * 75) / 100)
        nDmgMax = Fix((nDmgMax * 75) / 100)
    End If
    
    If bClassStealth Or Not bGreaterMUD Then
        nDmgMin = Fix(((nLevel + 100) * nDmgMin) / 100)
        nDmgMax = Fix(((nLevel + 100) * nDmgMax) / 100)
    End If
    
    nAttackAccuracy = CalculateBackstabAccuracy(nStealth, nAgility, nPlusBSaccy, bClassStealth, _
        IIf(tCharStats.bIsLoadedCharacter, nGlobalCharAccyAbils + nGlobalCharAccyOther, 0), nLevel, nStrength, nStrReq)
    
ElseIf nAttackTypeMUD = a6_Bash Then 'bash
    nCritChance = 0
    nQnDBonus = 0
    nPreRollMinModifier = 1.1
    nPreRollMaxModifier = 1.1
    If bGreaterMUD Then
        nDamageMultiplierMin = 2.5
        nDamageMultiplierMax = 3
    Else
        nDamageMultiplierMin = 3
        nDamageMultiplierMax = 3
    End If
    nAttackAccuracy = nAttackAccuracy - 15
    tRet.sAttackDesc = "bash with " & tRet.sAttackDesc
ElseIf nAttackTypeMUD = a7_Smash Then 'smash
    nCritChance = 0
    nQnDBonus = 0
    nPreRollMinModifier = 1.2
    nPreRollMaxModifier = 1.2
    nDamageMultiplierMin = 5
    nDamageMultiplierMax = 5
    nAttackAccuracy = nAttackAccuracy - 25
    tRet.sAttackDesc = "smash with " & tRet.sAttackDesc
End If

'to be implemented:

'If PARTY_FRONTRANK Then
'    if NOT ATTACK_TYPE = 3 then CHAR_ACCY += 5 'NOT jumpkick
'ElseIf PARTY_BACKRANK Then
'    if NOT ATTACK_TYPE = 3 then CHAR_ACCY -= 10 'NOT jumpkick
'End If

'ACCY_PENALTY += MONSTER_ABILITY104 'Abil 104 = DefenseModifier

calc_damage:
If nAttackAccuracy = 0 And nSpecifyAccy < 0 And tCharStats.bIsLoadedCharacter Then
    nAttackAccuracy = val(frmMain.lblInvenCharStat(10).Tag)
ElseIf nSpecifyAccy >= 0 Then
    nAttackAccuracy = nSpecifyAccy
End If
If nAttackAccuracy < 8 Then nAttackAccuracy = 8

nHitChance = 100

'//before switching to CalculateAttackDefense
'If nVSAC > 0 Then
'    accTemp = (nAttackAccuracy * nAttackAccuracy) \ 140
'    If accTemp < 1 Then accTemp = 1
'
'    If nAttackTypeMUD = a4_Surprise Then 'surprise
'        If bGreaterMUD Then
'            '(Backstab ACC)(Backstab ACC) / ((((AC/4)+BS Defense)(((AC/4)+BS Defense)/140)
'            nHitChance = 100 - ((((nVSAC \ 4) + nSecondaryDefense) * ((nVSAC \ 4) + nSecondaryDefense)) \ accTemp)
'        Else
'            nHitChance = 100 - nAttackAccuracy - nVSAC
'        End If
'    Else
'        'SuccessChance = Round(1 - (((m_nUserAC * m_nUserAC) / 100) / ((nAttack_AdjSuccessChance * nAttack_AdjSuccessChance) / 140)), 2) * 100
'        'nHitChance = Round(1 - (((nVSAC * nVSAC) / 100) / ((nAttackAccuracy * nAttackAccuracy) / 140)), 2) * 100
'        If nSecondaryDefense > 0 Then
'            nDefense = ((nVSAC * 10) + nSecondaryDefense) \ 10
'        Else
'            nDefense = nVSAC
'        End If
'        nHitChance = 100 - ((nDefense * nDefense) \ accTemp)
'    End If
'End If
'
'If bGreaterMUD Then
'    If nHitChance < 2 Then nHitChance = 2
'Else
'    If nHitChance < 8 Then nHitChance = 8
'End If
'If nHitChance > 98 Then nHitChance = 98
'
'If nVSDodge < 0 And nVSAC > 0 And Not bGreaterMUD Then 'i'm assuming this doesn't exist in gmud
'    'the dll provides for a x% chance for AC to be ignored if dodge is negative
'    'i.e.: (-dodge+100) = chance to ignore AC check and have a 99% hit chance
'    'so, if dodge was -10, there would be a 10% chance to ignore AC
'    'i'm simulating this by just reducing the hitchance at scale.
'    nPercent = ((nVSDodge + 100) / 100)  'chance for 99% hit
'    nPercent2 = 1 - nPercent 'chance for regular hit chance
'    nHitChance = (99 * nPercent) + (nHitChance * nPercent2)
'    If nHitChance < 8 Then nHitChance = 8
'
'ElseIf nVSDodge > 0 And nAttackAccuracy > 0 Then
'    If bGreaterMUD Then
''        If nAttackTypeMUD = a4_Surprise Then
''            '(Backstab ACC)(Backstab ACC) / (((Dodge))((Dodge))/140)
''            accTemp = (nVSDodge * nVSDodge) \ 140
''            If accTemp < 1 Then accTemp = 1
''            nPercent = (nAttackAccuracy * nAttackAccuracy) \ accTemp
''        Else
'            '((dodge * dodge)) / Math.Max((((accuracy * accuracy) / 14) / 10), 1)
'            accTemp = (nAttackAccuracy * nAttackAccuracy) \ 140
'            If accTemp < 1 Then accTemp = 1
'            nPercent = (nVSDodge * nVSDodge) \ accTemp
'            If nPercent > GMUD_DODGE_SOFTCAP Then
'                nPercent = GMUD_DODGE_SOFTCAP + GmudDiminishingReturns(nPercent - GMUD_DODGE_SOFTCAP, 4#)
'            End If
''        End If
'        If nPercent > 98 Then nPercent = 98
'    Else
'        accTemp = Fix(nAttackAccuracy \ 8)
'        If accTemp < 1 Then accTemp = 1
'        nPercent = Fix((nVSDodge * 10) \ accTemp)
'        If nPercent > 95 Then nPercent = 95
'        If nAttackTypeMUD = a4_Surprise Then nPercent = Fix(nPercent / 5) 'backstab
'    End If
'    tRet.nDodgeChance = nPercent
'    nPercent = (nPercent / 100) '% chance to dodge
'    nHitChance = (nHitChance * (1 - nPercent))
'End If


'//AS IT WAS 2025.08.12:
'If nVSAC > 0 Then
'    If nAttackTypeMUD = 4 Then 'surprise
'        nHitChance = nAttackAccuracy - nVSAC
'    Else
'        'SuccessChance = Round(1 - (((m_nUserAC * m_nUserAC) / 100) / ((nAttack_AdjSuccessChance * nAttack_AdjSuccessChance) / 140)), 2) * 100
'        nHitChance = Round(1 - (((nVSAC * nVSAC) / 100) / ((nAttackAccuracy * nAttackAccuracy) / 140)), 2) * 100
'    End If
'    If nHitChance < 9 Then nHitChance = 9
'    If nHitChance > 99 Then nHitChance = 99
'End If
'If nVSDodge < 0 And nVSAC > 0 Then
'    'the dll provides for a x% chance for AC to be ignored if dodge is negative
'    'i.e.: (-dodge+100) = chance to ignore AC check and have a 99% hit chance
'    'so, if dodge was -10, there would be a 10% chance to ignore AC
'    'i'm simulating this by just reducing the hitchance at scale.
'    nPercent = ((nVSDodge + 100) / 100)  'chance for 99% hit
'    nPercent2 = 1 - nPercent 'chance for regular hit chance
'    nHitChance = (99 * nPercent) + (nHitChance * nPercent2)
'    If nHitChance < 9 Then nHitChance = 9
'ElseIf nVSDodge > 0 And nAttackAccuracy > 0 Then
'    nPercent = Fix((nVSDodge * 10) / Fix(nAttackAccuracy / 8))
'    If nPercent > 95 Then nPercent = 95
'    If nAttackTypeMUD = 4 Then nPercent = Fix(nPercent / 5) 'backstab
'    CalculateAttack.nDodgeChance = nPercent
'    nPercent = (nPercent / 100) 'chance to dodge
'    nHitChance = (nHitChance * (1 - nPercent))
'    If nHitChance < 9 Then nHitChance = 9
'End If

If nVSAC > 0 Or nVSDodge > 0 Then
    nDefense = CalculateAttackDefense(nAttackAccuracy, nVSAC, nVSDodge, nSecondaryDefense, 0, 0, 0, 0, 0, False, False, _
                    IIf(nAttackTypeMUD = a4_Surprise, True, False), False, GetDodgeCap(-1))
                    
    nHitChance = nDefense(0)
    If nDefense(1) > 0 Then
        tRet.nDodgeChance = nDefense(1)
        nHitChance = (nHitChance * (1 - (nDefense(1) / 100)))
    End If
End If

If Not bGreaterMUD And nVSDodge < 0 And nVSAC > 0 Then
    'the dll provides for a x% chance for AC to be ignored if dodge is negative
    'i.e.: (-dodge+100) = chance to ignore AC check and have a 99% hit chance
    'so, if dodge was -10, there would be a 10% chance to ignore AC
    'i'm simulating this by just reducing the hitchance at scale.
    nPercent = ((nVSDodge + 100) / 100)  'chance for 99% hit
    nPercent2 = 1 - nPercent 'chance for regular hit chance
    nHitChance = (99 * nPercent) + (nHitChance * nPercent2)
    If nHitChance < STOCK_HIT_MIN Then nHitChance = STOCK_HIT_MIN
End If

nHitChance = nHitChance / 100

If nPreRollMinModifier > 1 Then nDmgMin = Fix(nDmgMin * nPreRollMinModifier)
If nPreRollMaxModifier > 1 Then nDmgMax = Fix(nDmgMax * nPreRollMaxModifier)

If nCritChance > 0 Then
    nMinCrit = nDmgMax * 2
    nMaxCrit = nDmgMax * 4
    If nMinCrit > nMaxCrit Then nMaxCrit = nMinCrit
    nAvgCrit = Round((nMinCrit + nMaxCrit) / 2) - nVSDR
    nMinCrit = nMinCrit - nVSDR
    nMaxCrit = nMaxCrit - nVSDR
    If nAvgCrit < 0 Then nAvgCrit = 0
    If nMinCrit < 0 Then nMinCrit = 0
    If nMaxCrit < 0 Then nMaxCrit = 0
End If

If bGreaterMUD Then
    nDmgMin = Fix(nDmgMin * nDamageMultiplierMin) - nVSDR
    nDmgMax = Fix(nDmgMax * nDamageMultiplierMax) - nVSDR
Else
    nDmgMin = Fix((nDmgMin - nVSDR) * nDamageMultiplierMin)
    nDmgMax = Fix((nDmgMax - nVSDR) * nDamageMultiplierMax)
End If

If nDmgMin < 0 Then nDmgMin = 0
If nDmgMax < 0 Then nDmgMax = 0
nAvgHit = Round((nDmgMin + nDmgMax) / 2)

If Len(sCasts) = 0 And nWeaponNumber > 0 And nAttackTypeMUD > a3_Jumpkick Then
    For x = 0 To 19
        Select Case tabItems.Fields("Abil-" & x)
            Case 43: 'casts spell
                sCasts = AutoAppend(sCasts, "[" & GetSpellName(tabItems.Fields("AbilVal-" & x), bHideRecordNumbers) _
                    & ", " & PullSpellEQ(True, 0, tabItems.Fields("AbilVal-" & x), , , , True), "|")
                If Not nPercent = 0 Then
                    sCasts = sCasts & ", " & nPercent & "%]"
                Else
                    sCasts = sCasts & "]"
                End If
                
            Case 114: '%spell
                nPercent = tabItems.Fields("AbilVal-" & x)
              
        End Select
    Next x
End If

If Len(sCasts) > 0 And nWeaponNumber > 0 And nAttackTypeMUD > a3_Jumpkick Then
    'this is matching against:
    '[fire burns(979), Damage 5 to 15, 100%] -- would produce 1 full match and 3 submatches for the 5, 15, and 100
    'or: [lacerate(985), Damage 3 to 12, AffectsLivingOnly, for 10 rounds, 100%] -- this will produce the same matching as above with the 10 rounds being ignored
    'or: [fire burns(979), Damage 5 to 15, 25%], [ice freezes(978), Damage 5 to 15, 25%] -- would produce 2 full matches, each with 3 submatches
    'or: [{rocks shred(977), Damage 5 to 15} OR {ice freezes(978), Damage 5 to 15} OR {fire burns(979), Damage 5 to 15} OR {acid sears(980), Damage 5 to 15} OR {lightning shocks(981), Damage 5 to 15}], 100%]
    '     ...which would produce only 1 full match with all of the damage numbers and final percentage as submatches
    sRegexPattern = "(?:(?:Damage(?:\(-MR\))?|DrainLife) (-?\d+) to (-?\d+)[^\]]*, (\d+)%|\[(?:{[^\[\{\}\]]+, (?:Damage(?:\(-MR\))?|DrainLife) (-?\d+) to (-?\d+)[^\]\}]*(?:} OR ))(?:{[^\[\{\}\]]+, (?:Damage(?:\(-MR\))?|DrainLife) (-?\d+) to (-?\d+)[^\]\}]*(?:} OR )?)?(?:{[^\[\{\}\]]+, (?:Damage(?:\(-MR\))?|DrainLife) (-?\d+) to (-?\d+)[^\]\}]*(?:} OR )?)?(?:{[^\[\{\}\]]+, (?:Damage(?:\(-MR\))?|DrainLife) (-?\d+) to (-?\d+)[^\]\}]*(?:} OR )?)?(?:{[^\[\{\}\]]+, (?:Damage(?:\(-MR\))?|DrainLife) (-?\d+) to (-?\d+)[^\]\}]*(?:} OR )?)?(?:{[^\[\{\}\]]+, (?:Damage(?:\(-MR\))?|DrainLife) (-?\d+) to (-?\d+)[^\]\}]*(?:} OR )?)?}], (\d+)%)"
    tMatches() = RegExpFindv2(sCasts, sRegexPattern, False, False, False)
    If UBound(tMatches()) = 0 And Len(tMatches(0).sFullMatch) = 0 Then GoTo done_extra:
       
'    If tabItems.Fields("Number") = 2577 Then
'        Debug.Print tabItems.Fields("Number")
'    End If
    
    For iMatch = 0 To UBound(tMatches())
        If UBound(tMatches(iMatch).sSubMatches()) < 2 Then GoTo skip_match:
        
        If InStr(1, tMatches(iMatch).sFullMatch, "} or {", vbTextCompare) > 0 Then
            'multiple spells with equal chance
            sArr() = Split(tMatches(iMatch).sFullMatch, "} or {", , vbTextCompare)
        Else
            ReDim sArr(0)
            sArr(0) = tMatches(iMatch).sFullMatch
        End If
        
        nExtraTMP = 0
        nCount = 0
        nDurDamage = 0
        nDurCount = 0
        For x = 0 To UBound(tMatches(iMatch).sSubMatches()) - 1
            nTemp = x - Fix((x + 1) / 2) 'index that refers to the full text string of the match for these two damage values
            If UBound(sArr()) >= nTemp Then
                If InStr(1, sArr(nTemp), ", for", vbTextCompare) > 0 And InStr(1, sArr(nTemp), "rounds", vbTextCompare) > 0 Then
                    nDurDamage = nDurDamage + Abs(val(tMatches(iMatch).sSubMatches(x)))
                    nDurCount = nDurCount + 1
                    nCount = nCount + 1
                    x = x + 1 'get the next number
                    If UBound(tMatches(iMatch).sSubMatches()) >= (x + 1) Then 'plus another because there should also be the percentage at the end
                        nDurDamage = nDurDamage + Abs(val(tMatches(iMatch).sSubMatches(x)))
                        nDurCount = nDurCount + 1
                        nCount = nCount + 1 'still counting here because its presence would reduce the chance of casting the other spells in the group, thereby reducing their overall effect on the average damage
                    End If
                    GoTo skip_submatch:
                End If
            End If
            nExtraTMP = nExtraTMP + Abs(val(tMatches(iMatch).sSubMatches(x)))
            nCount = nCount + 1
skip_submatch:
        Next x
        
        If nCount > 0 Then nExtraTMP = Round(nExtraTMP / nCount, 2)
        nExtraAvgHit = Fix(nExtraAvgHit + nExtraTMP)
        nExtraPCT = Round(val(tMatches(iMatch).sSubMatches(UBound(tMatches(iMatch).sSubMatches()))) / 100, 2)
        nExtraTMP = Round(nExtraTMP * nExtraPCT, 2)
        
        'dividing durection by SWINGS so it actually counts only once when it multiplies by SWINGS later (e.g. we're adding one tick of the duration damage to the total per-round damage)
        If nDurCount > 0 Then nExtraTMP = nExtraTMP + Round(((nDurDamage / nDurCount) * nExtraPCT) / nSwings, 2)
        
        nExtraAvgSwing = Fix(nExtraAvgSwing + nExtraTMP)
skip_match:
    Next iMatch
    
    If UBound(tMatches()) > 0 Then nExtraAvgHit = Round(nExtraAvgHit / (UBound(tMatches()) + 1))
    nExtraAvgSwing = Round(nExtraAvgSwing)
End If
done_extra:

tRet.nMinDmg = nDmgMin
tRet.nMaxDmg = nDmgMax
tRet.nAvgHit = nAvgHit
tRet.nAvgCrit = nAvgCrit
tRet.nMaxCrit = nMaxCrit
tRet.nAvgExtraHit = nExtraAvgHit
tRet.nAvgExtraSwing = nExtraAvgSwing
tRet.nCritChance = nCritChance
tRet.nQnDBonus = nQnDBonus
tRet.nSwings = nSwings
tRet.nAccy = nAttackAccuracy

nPercent = (nCritChance / 100) 'chance to crit
tRet.nRoundPhysical = (((1 - nPercent) * nAvgHit) + (nPercent * nAvgCrit)) * nSwings * nHitChance
tRet.nRoundTotal = tRet.nRoundPhysical + (nExtraAvgSwing * nSwings * nHitChance)
tRet.nHitChance = Round(nHitChance * 100)

If nSwings > 0 And (nAvgHit + nAvgCrit) > 0 Then
    If nAttackTypeMUD = a4_Surprise Then
        sAttackDetail = "Backstab: " & tRet.nRoundTotal & " avg @ " & tRet.nHitChance & "% hit "
        If nDmgMin < nAvgHit Or nDmgMax > nAvgHit Or nExtraAvgHit <> nExtraAvgSwing Then
            
            nTemp = nDmgMin
            nTemp2 = nDmgMax
            If nExtraAvgHit > 0 Then
                If nExtraAvgHit = nExtraAvgSwing Then nTemp = nTemp + nExtraAvgHit
                nTemp2 = nTemp2 + nExtraAvgHit
            End If
            sAttackDetail = sAttackDetail & "(Min/Avg/Max: " & nTemp
            sAttackDetail = sAttackDetail & "/" & ((nDmgMin + nDmgMax + nExtraAvgSwing + nExtraAvgSwing) \ 2)
            sAttackDetail = sAttackDetail & "/" & nTemp2 & ")"
'            If nExtraAvgHit > 0 Then
'                sAttackDetail = sAttackDetail & "(Min/Max: " & IIf(nExtraAvgHit <> nExtraAvgSwing, nDmgMin, nDmgMin + nExtraAvgHit) & "/" & (nDmgMax + nExtraAvgHit) & ")"
'            ElseIf nDmgMin <> nDmgMax Then
'                sAttackDetail = sAttackDetail & "(Min/Max: " & nDmgMin & "/" & nDmgMax & ")"
'            End If
        End If
    Else
        sAttackDetail = "Swings: " & Round(nSwings, 1) & ", Avg Hit: " & nAvgHit
        If nAvgCrit > 0 Then
            sAttackDetail = AutoAppend(sAttackDetail, "Avg/Max Crit: " & nAvgCrit & "/" & nMaxCrit)
            If nCritChance > 0 Then sAttackDetail = sAttackDetail & " (" & nCritChance & "%)"
        End If
        If tRet.nHitChance > 0 Then sAttackDetail = AutoAppend(sAttackDetail, "Hit: " & tRet.nHitChance & "%")
    End If
End If
tRet.sAttackDetail = sAttackDetail

CalculateAttack = tRet

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalculateAttack")
Resume out:
End Function

Public Function GmudDiminishingReturns(ByVal nValue As Double, ByVal nScale As Double) As Double
On Error GoTo error:
Dim mult As Double
Dim triNum As Double
Dim isNeg As Boolean

If nScale <= 0# Then
    GmudDiminishingReturns = nValue
    Exit Function
End If

isNeg = (nValue < 0#)
If isNeg Then nValue = -nValue

mult = nValue / nScale
triNum = (Sqr(8# * mult + 1#) - 1#) / 2#

If isNeg Then
    GmudDiminishingReturns = -triNum * nScale
Else
    GmudDiminishingReturns = triNum * nScale
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GmudDiminishingReturns")
Resume out:
End Function

Public Function GetDodgeCap(Optional ByVal nClass As Long) As Long
On Error GoTo error:

If Not bGreaterMUD Then
    GetDodgeCap = STOCK_DODGE_CAP
    Exit Function
End If

GetDodgeCap = GMUD_DODGE_SOFTCAP

If nClass < 0 Then Exit Function 'input -1 to force skip class lookup
If nClass < 1 Then
    If frmMain.chkGlobalFilter.Value = 1 Then
        nClass = frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex)
    End If
End If
If nClass < 1 Then Exit Function

If GetClassMagery(nClass) = Kai Then 'mystic
    GetDodgeCap = GetDodgeCap + 10
ElseIf GetClassArmourType(nClass) = 2 Then 'ninja
    GetDodgeCap = GetDodgeCap + 10
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetDodgeCap")
Resume out: End Function

Public Function CalculateAttackDefense(ByVal nAccy As Long, ByVal nAC As Long, ByRef nDodge As Long, Optional ByRef nSecondaryDef As Long, _
    Optional ByVal nProtEv As Long, Optional ByVal nProtGd As Long, Optional ByVal nPerception As Long, Optional ByVal nVileWard As Long, Optional ByVal eEvil As eEvilPoints, _
    Optional ByVal bShadow As Boolean, Optional ByVal bSeeHidden As Boolean, Optional ByVal bBackstab As Boolean, Optional ByVal bVsPlayer As Boolean, _
    Optional ByVal nDodgeCap As Long) As Long()
On Error GoTo error:
Dim nHitChance As Currency, nTotalHitPercent As Currency, nDefense As Long, nShadow As Integer
Dim dimReturns As Currency, nDodgeChance As Currency, nTemp As Long
Dim sPrint As String, accTemp As Long, dodgeTemp As Long, nReturn() As Long

'nSecondaryDef = BS Defense for backstabs

'0=nHitChance
'1=nDodgeChance
ReDim nReturn(1)
CalculateAttackDefense = nReturn

If nAccy > 9999 Then nAccy = 9999: If nAccy < 1 Then nAccy = 1
If nAC > 9999 Then nAC = 9999: If nAC < 0 Then nAC = 0
If nDodge > 9999 Then nDodge = 9999: If nDodge < -999 Then nDodge = 0
If nProtEv > 9999 Then nProtEv = 9999: If nProtEv < 0 Then nProtEv = 0
If nPerception > 9999 Then nPerception = 9999: If nPerception < 0 Then nPerception = 0
If nSecondaryDef > 9999 Then nSecondaryDef = 9999: If nSecondaryDef < 0 Then nSecondaryDef = 0
If nVileWard > 9999 Then nVileWard = 9999: If nVileWard < 0 Then nVileWard = 0
If eEvil > e7_FIEND Then eEvil = e7_FIEND: If eEvil < e0_Saint Then eEvil = e0_Saint
If nDodgeCap < 1 Then nDodgeCap = GetDodgeCap(-1)

'common accuracy value calculation for most things (exceptions: stock backstab)
accTemp = (nAccy * nAccy) \ 140
If accTemp < 1 Then accTemp = 1

'GET HIT CHANCE
If nAC + nDefense <= 0 Then
    nHitChance = 100
Else
    If bBackstab Then '[BACKSTAB]
        If bGreaterMUD Then '[BACKSTAB+GREATERMUD]
            If bVsPlayer Then '[BACKSTAB+GREATERMUD+PLAYER]
                
                If nVileWard > 0 And eEvil > 0 Then
                    If eEvil <= e3_Seedy Then
                        nVileWard = 0
                    ElseIf eEvil <= e5_Criminal Then
                        nVileWard = nVileWard \ 2
                    End If
                    nVileWard = nVileWard \ 10
                End If
                
                       '(ac + prev + (int)(inTarget.Perception*0.8) + ward) / 2 + shadow;
                nDefense = nAC + nProtEv + Fix(nPerception * 0.8) + nVileWard
                If bShadow Then nShadow = 10
                nDefense = (nDefense \ 2) + nShadow
                nSecondaryDef = nDefense
                
            Else '[BACKSTAB+GREATERMUD+MOB]
                
                '(((AC/4)+BS Defense)(((AC/4)+BS Defense)
                If bSeeHidden Then
                    nDefense = nAC + nSecondaryDef
                Else
                    nDefense = (nAC \ 4) + nSecondaryDef
                End If
            
            End If
            
            If nDefense < 0 Then nDefense = 0
            If nDefense > 9999 Then nDefense = 9999
            nHitChance = 100 - ((nDefense * nDefense) \ accTemp)
            
        Else '[BACKSTAB+STOCK]
            If bVsPlayer Then
                nDefense = (nAC + nPerception) \ 2
            Else
                nDefense = (nAC \ 4) + nSecondaryDef
            End If
            
            'TECHNICALLY THE STOCK DLL HAS THIS:
            'ARRAYofSTUFF[1] = ((int64_t)AC_FROM_WORNITEMS / 10 + (int32_t)*(uint16_t*)((char*)USER_RECORD + 0x5f8)) / 2
            'and then... ARRAYofSTUFF[1] += CUMLTV_AC_FROM_ABILITY2 >> 2;
            'Which means the AC in the above calculation should be for worn items only
            'Then we should be adding ([AC from abil 2] / 4) to nDefense after that calculation
            'e.g. it should be: ((AC_WORN+nPerception)\2)+(AC_ABIL2/4)
            
            nHitChance = nAccy - nDefense 'need to add +AC_BLUR HERE
            
        End If
        
    Else 'NORMAL ATTACK
        
        '((AC*AC)/100)/((ACCY*ACCY)/140)=fail %
        'nAccy = Round((((nAC * nAC) / 100) / ((nAccy * nAccy) / 140)), 2) * 100
        If bGreaterMUD Then
            'implement pop-up questionaire on hitcalc like we did for backstab?
            nTemp = nProtEv + nProtGd + (nVileWard + IIf(bShadow, 100, 0) \ 10)
            If nSecondaryDef < nTemp Then nSecondaryDef = nTemp
        End If
        nDefense = nAC + nSecondaryDef 'need to add +(AC_BLUR\2) HERE
        nHitChance = 100 - ((nDefense * nDefense) \ accTemp)
    
    End If
End If

If bGreaterMUD Then
    If nHitChance < GMUD_HIT_MIN Then nHitChance = GMUD_HIT_MIN '2
    If nHitChance > GMUD_HIT_CAP Then nHitChance = GMUD_HIT_CAP '98
Else
    If nHitChance < STOCK_HIT_MIN Then nHitChance = STOCK_HIT_MIN '8
    If nHitChance > STOCK_HIT_CAP Then nHitChance = STOCK_HIT_CAP '99
End If

'GET DODGE CHANCE
If bGreaterMUD Then

    If (nDodge > 0 Or (nPerception > 0 And bBackstab And bVsPlayer)) Then
        If bBackstab And bVsPlayer Then 'bs AND vs player
            dodgeTemp = (nDodge + (nPerception \ 2)) \ 2
            If bSeeHidden Then
                If nDodge - 9 > dodgeTemp Then dodgeTemp = nDodge
            End If
            nDodge = dodgeTemp
        End If
        
        '((dodge * dodge)) / Math.Max((((accuracy * accuracy) / 14) / 10), 1)
        nDodgeChance = (nDodge * nDodge) \ accTemp
        If nDodgeChance > nDodgeCap Then
            nDodgeChance = nDodgeCap + GmudDiminishingReturns(nDodgeChance - nDodgeCap, 4#)
        End If
    End If

ElseIf nDodge > 0 Then

    accTemp = nAccy \ 8
    If accTemp < 1 Then accTemp = 1
    nDodgeChance = Fix((nDodge * 10) \ accTemp)
    If nDodgeChance > nDodgeCap Then nDodgeChance = nDodgeCap
    If bBackstab Then nDodgeChance = Fix(nDodgeChance \ 5)  'backstab
    
End If

If nDodgeChance < 0 Then nDodgeChance = 0
If bGreaterMUD Then
    If nDodgeChance > GMUD_DODGE_CAP Then nDodgeChance = GMUD_DODGE_CAP '98
Else
    If nDodgeChance > STOCK_DODGE_CAP Then nDodgeChance = STOCK_DODGE_CAP '95
End If

nReturn(0) = nHitChance
nReturn(1) = nDodgeChance

CalculateAttackDefense = nReturn

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalculateAttackDefense")
Resume out:
End Function

Public Function CalculateBackstabAccuracy(ByVal nStealth As Integer, ByVal nAgility As Integer, ByVal nPlusBSaccy As Integer, _
    ByVal bClassStealth As Boolean, ByVal nPlusNormalAccy As Integer, _
    Optional ByVal nLevel As Integer, Optional ByVal nStrength As Integer, Optional ByVal nStrReq As Integer) As Long
On Error GoTo error:
Dim nAccy As Long

If bGreaterMUD Then
    nAccy = ((nStealth / 3) + ((nAgility - 50) + nLevel) / 2) + 15 + nPlusBSaccy
    If nStrength < nStrReq Then nAccy = nAccy - 15
Else
    nAccy = Fix((nStealth + nAgility) / 2) + Fix(nPlusBSaccy / 2)
End If

If bClassStealth Then 'has classstealth
    nAccy = nAccy + 5
Else 'has racestealth only
    nAccy = nAccy - 15
End If
nAccy = nAccy + nPlusNormalAccy

CalculateBackstabAccuracy = nAccy

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalculateBackstabAccuracy")
Resume out: End Function

Public Function CalculateAccuracy(Optional ByVal nClass As Integer, Optional ByVal nLevel As Integer, _
    Optional ByVal nSTR As Integer, Optional ByVal nAGI As Integer, _
    Optional ByVal nINT As Integer, Optional ByVal nCHA As Integer, _
    Optional ByVal nAccyWorn As Integer, Optional ByVal nPlusAccy As Integer, _
    Optional ByVal nEncumPCT As Integer, Optional ByRef sReturnText As String, _
    Optional ByVal nAttackTypeMUD As eAttackTypeMUD) As Long
On Error GoTo error:
Dim nTemp As Double, nCombatLevel As Integer, nAccyCalc As Integer, nEncumBonus As Double, nBaseAccy As Double

'note: in CalcCharacterStats, nAccyWorn is purposely being passed as -1 because it is adding nAccyWorn after this calculation
If nAccyWorn = 0 And (bGreaterMUD = False Or nPlusAccy = 0) Then
    nAccyWorn = 1
    sReturnText = AutoAppend(sReturnText, "Pity Accy (" & nAccyWorn & ")", vbCrLf)
End If
If nEncumPCT = 0 Then nEncumPCT = 1
If nAccyWorn < 0 Then nAccyWorn = 0
'If bGreaterMUD Then nAccyCalc = 2

If nEncumPCT < 33 Then
    nEncumBonus = 15 - Fix(nEncumPCT / 10)
    sReturnText = AutoAppend(sReturnText, "Encum (" & nEncumBonus & ")", vbCrLf)
    nAccyCalc = nAccyCalc + nEncumBonus
End If

If Not bGreaterMUD Then
    nTemp = nAccyCalc
    nAccyCalc = Fix(nAccyCalc / 2) * 2 '...how it is in the dll
    If nTemp <> nAccyCalc Then sReturnText = AutoAppend(sReturnText, "odd number penalty (" & Round(nAccyCalc - nTemp) & ")", vbCrLf)
End If

If nLevel > 0 Then
    nBaseAccy = Fix(Sqr(nLevel))
    Do While ((nBaseAccy + 1) * (nBaseAccy + 1)) <= nLevel And bGreaterMUD = False
        nBaseAccy = nBaseAccy + 1
    Loop
End If

If nClass > 0 Then
    nCombatLevel = GetClassCombat(nClass) + 2 'GetClassCombat subtracts 2 from the value in the database
    nBaseAccy = nBaseAccy * (nCombatLevel - 1)
    nBaseAccy = (nBaseAccy + (nCombatLevel * 2) + Fix(nLevel / 2) - 2) * 2
End If
If nBaseAccy > 0 Then sReturnText = AutoAppend(sReturnText, "Combat+Level (" & nBaseAccy & ")", vbCrLf)
nAccyCalc = nAccyCalc + nBaseAccy

If nSTR > 0 And (bGreaterMUD = False Or nAttackTypeMUD = a6_Bash Or nAttackTypeMUD = a7_Smash) Then
    nTemp = Fix((nSTR - 50) / 3)
    If nTemp <> 0 Then 'str
        sReturnText = AutoAppend(sReturnText, "Strength (" & nTemp & ")", vbCrLf)
        If bGreaterMUD And (nAttackTypeMUD = a6_Bash Or nAttackTypeMUD = a7_Smash) Then
            sReturnText = sReturnText & " *" & IIf(nAttackTypeMUD = a7_Smash, "smash", "bash")
        End If
        nAccyCalc = nAccyCalc + nTemp
    End If
End If

If nAGI > 0 Then
    If (bGreaterMUD = False Or nAttackTypeMUD = a6_Bash Or nAttackTypeMUD = a7_Smash) Then
        nTemp = Fix((nAGI - 50) / 6)
    Else
        nTemp = Fix((nAGI - 50) / 3)
    End If
    If nTemp <> 0 Then
        sReturnText = AutoAppend(sReturnText, "Agility (" & nTemp & ")", vbCrLf)
        If bGreaterMUD And (nAttackTypeMUD = a6_Bash Or nAttackTypeMUD = a7_Smash) Then
            sReturnText = sReturnText & " *" & IIf(nAttackTypeMUD = a7_Smash, "smash", "bash")
        End If
        nAccyCalc = nAccyCalc + nTemp
    End If
End If

If nINT > 0 And bGreaterMUD And nAttackTypeMUD <> a6_Bash And nAttackTypeMUD <> a7_Smash Then
    nTemp = Fix((nINT - 50) / 6)
    If nTemp <> 0 Then
        nAccyCalc = nAccyCalc + nTemp
        sReturnText = AutoAppend(sReturnText, "Intellect (" & nTemp & ")", vbCrLf)
    End If
End If

If nCHA > 0 And bGreaterMUD And nAttackTypeMUD <> a6_Bash And nAttackTypeMUD <> a7_Smash Then
    nTemp = Fix((nCHA - 50) / 10)
    If nTemp <> 0 Then
        nAccyCalc = nAccyCalc + nTemp
        sReturnText = AutoAppend(sReturnText, "Charm (" & nTemp & ")", vbCrLf)
    End If
End If

CalculateAccuracy = nAccyCalc + nAccyWorn + nPlusAccy
If (bGreaterMUD And nAttackTypeMUD = a7_Smash) Then
    nTemp = Fix((CalculateAccuracy * 3) / 2)
    sReturnText = AutoAppend(sReturnText, "Smash Bonus (" & (nTemp - CalculateAccuracy) & ")", vbCrLf)
    CalculateAccuracy = nTemp
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalculateAccuracy")
Resume out:
End Function

Public Function GetAbilityStats(ByVal nNum As Integer, Optional ByVal nValue As Integer, _
    Optional ByRef LV As ListView, Optional ByVal bCalcSpellLevel As Boolean = True, _
    Optional ByVal bPercentColumn As Boolean) As String
Dim sHeader As String, oLI As ListItem, sTemp As String, sArr() As String, x As Integer
Dim sTextblockCasts As String
On Error GoTo error:

GetAbilityStats = GetAbilityName(nNum)
If GetAbilityStats = "" Then Exit Function

sTemp = ""
If nNum = 148 And nValue > 0 Then '148-execute tb
    sTemp = GetTextblockAction(nValue)
    If InStr(1, sTemp, "cast ", vbTextCompare) = 0 Then
        GoTo skip_textblock_spells_only:
    Else
        sArr() = Split(sTemp, ":")
        For x = 0 To UBound(sArr())
            If Left(sArr(x), 5) = "cast " Then
                sTemp = PullSpellEQ(False, , (val(Mid(sArr(x), 6))))
                If Not sTextblockCasts = "" Then sTextblockCasts = sTextblockCasts & ", "
                sTextblockCasts = sTextblockCasts & sTemp
            Else
                GoTo skip_textblock_spells_only:
            End If
        Next x
    End If
End If
If Not sTextblockCasts = "" Then
    If InStr(1, sTextblockCasts, "(click)", vbTextCompare) > 0 Then sTextblockCasts = "(click)"
    GetAbilityStats = sTextblockCasts
    Exit Function
End If

skip_textblock_spells_only:
If Not nValue = 0 Then

    If nValue < 0 Then
        sHeader = " "
    Else
        sHeader = " +"
    End If
    
    Select Case nNum
        Case 7: 'DR
            GetAbilityStats = GetAbilityStats & sHeader & (nValue / 10)
        Case 42, 122, 160: 'learnspell, removesspell, givetempspell
            sTemp = GetSpellName(nValue, bHideRecordNumbers)
            GetAbilityStats = GetAbilityStats & " (" & sTemp & ")"
            If Not LV Is Nothing Then
                If bPercentColumn Then
                    Set oLI = LV.ListItems.Add()
                    oLI.Text = ""
                    oLI.ListSubItems.Add , , "Spell: " & sTemp
                    oLI.ListSubItems(1).Tag = nValue
                Else
                    Set oLI = LV.ListItems.Add()
                    oLI.Text = "Spell: " & sTemp
                    oLI.Tag = nValue
                End If
            End If
        Case 43, 153: 'castsp, killspell
            GetAbilityStats = GetAbilityStats & " [" & GetSpellName(nValue, bHideRecordNumbers) & ", " & PullSpellEQ(bCalcSpellLevel, 0, nValue, IIf(LV Is Nothing, Nothing, LV), , , bPercentColumn, True) & "]"
        Case 73, 124: 'dispell magic, negateabil
            GetAbilityStats = GetAbilityStats & " (" & GetAbilityName(nValue) & ")"
        Case 151: 'endcast
            GetAbilityStats = GetAbilityStats & " [" & GetSpellName(nValue, bHideRecordNumbers) & ", " & PullSpellEQ(bCalcSpellLevel, 0, nValue, IIf(LV Is Nothing, Nothing, LV), , , bPercentColumn, True) & "]"
        Case 59: 'class ok
            GetAbilityStats = GetAbilityStats & " " & GetClassName(nValue)
        Case 146, 12: 'mon guards, summon
            GetAbilityStats = GetAbilityStats & " " & GetMonsterName(nValue, bHideRecordNumbers)
        Case 1, 8, 17, 18, 19, 140, 141, 148:
            'NO HEADERS, damage, drain, damage(on armr), poison, heal, teleport room, teleport map, textblocks
            ' *** ALSO ADD THESE TO PullSpellEQ ***
            GetAbilityStats = GetAbilityStats & " " & nValue
        Case 178: 'no action
            '178-shadowform: value is just the message
        Case 185: 'noattack / bad attack
            GetAbilityStats = GetAbilityStats & " " & GetItemName(nValue, bHideRecordNumbers)
        Case Else:
            GetAbilityStats = GetAbilityStats & sHeader & nValue
    End Select
    
End If

Set oLI = Nothing
Exit Function

error:
Call HandleError("GetAbilityStats")
Set oLI = Nothing
End Function

Public Function ExtractTextCommand(ByVal sWholeString As String) As String
On Error GoTo error:
Dim x As Long, sCommand As String, sChar As String

x = InStr(1, sWholeString, " ") + 1
If x = 1 Then
    ExtractTextCommand = sWholeString
    Exit Function
End If

Do While x < Len(sWholeString)
    sChar = Mid(sWholeString, x, 1)
    If sChar = "," Then
        If Not sCommand = "" Then Exit Do
    End If
    sCommand = sCommand & sChar
    x = x + 1
Loop

If sCommand = "" Then
    ExtractTextCommand = sWholeString
    Exit Function
End If

ExtractTextCommand = sCommand

Exit Function

error:
Call HandleError("ExtractTextCommand")
ExtractTextCommand = sWholeString
End Function
Public Function ExtractMapRoom(ByVal sExit As String) As RoomExitType
Dim x As Integer, y As Integer, i As Integer

On Error GoTo error:

ExtractMapRoom.Map = 0
ExtractMapRoom.Room = 0
ExtractMapRoom.ExitType = 0

x = InStr(1, sExit, "/")
Do While x - 1 > 0 'gets where the map number starts
    Select Case Mid(sExit, x - 1, 1)
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0":
            i = x - 1
        Case Else:
            Exit Do
    End Select
    x = x - 1
Loop

'For i = 1 To Len(sExit) - 1 'gets where the first number is
'    Select Case Mid(sExit, i, 1)
'        Case "1", "2", "3", "4", "5", "6", "7", "8", "9": Exit For
'    End Select
'Next

x = InStr(1, sExit, "/")
If x = 0 Then Exit Function
If x = Len(sExit) Then Exit Function

ExtractMapRoom.Map = val(Mid(sExit, i, x - 1))

y = InStr(x, sExit, " ")
If y = 0 Then
    ExtractMapRoom.Room = val(Mid(sExit, x + 1))
Else
    ExtractMapRoom.Room = val(Mid(sExit, x + 1, y - 1))
    ExtractMapRoom.ExitType = Mid(sExit, y + 1)
End If

Exit Function

error:
Call HandleError("ExtractMapRoom")

End Function

Public Function CalcDodge(Optional ByVal nCharLevel As Integer, Optional ByVal nAgility As Integer, Optional ByVal nCharm As Integer, Optional ByVal nPlusDodge As Integer, _
    Optional ByVal nCurrentEncum As Integer = 0, Optional ByVal nMaxEncum As Integer = -1, Optional ByVal nClass As Long, _
    Optional ByVal bRawValues As Boolean) As Integer
On Error GoTo error:
Dim nDodge As Integer, nEncumPCT As Integer, nTemp As Integer, nSoftCap As Long

nDodge = Fix(nCharLevel / 5)
nDodge = nDodge + Fix((nCharm - 50) / 5)
nDodge = nDodge + Fix((nAgility - 50) / 3)
nDodge = nDodge + nPlusDodge '[cumulative dodge from: abilities + auras + race + class + items]

If nMaxEncum > 0 Then
    nEncumPCT = Fix((nCurrentEncum / nMaxEncum) * 100)
    If nEncumPCT < 33 Then
        nDodge = nDodge + 10 - Fix(nEncumPCT / 10)
    End If
End If

If nDodge < 0 Then nDodge = 0

If Not bRawValues Then
    If bGreaterMUD Then
        nSoftCap = GetDodgeCap(nClass)
        If nDodge > nSoftCap And nSoftCap > 0 Then nDodge = nSoftCap + GmudDiminishingReturns(nDodge - nSoftCap, 4#)
        If nDodge > 98 Then nDodge = 98
    Else
        If nDodge > 95 Then nDodge = 95
    End If
End If

CalcDodge = nDodge

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalcDodge")
Resume out:
End Function

Public Function CalcEncum(ByVal nStrength As Integer, Optional ByVal nEncumBonus As Integer) As Long
On Error GoTo error:

If nStrength < 0 Then CalcEncum = 0: Exit Function

If LCase(Right(frmMain.lblDatVer.Caption, 6)) = "v1.11i" Then
    CalcEncum = nStrength * 48
Else
    If nStrength < 101 Then
        CalcEncum = nStrength * 48
    Else
        CalcEncum = 4800 + ((CLng(nStrength) - 100) * 84)
    End If
End If

If nEncumBonus > 0 Then
    CalcEncum = CalcEncum + (CalcEncum * (nEncumBonus / 100))
End If

CalcEncum = Round(CalcEncum, 0)

Exit Function

error:
Call HandleError("CalcEncum")
End Function
Public Function GetSpellAttackType(ByVal nAttackTypeMUD As Integer) As String

On Error GoTo error:

Select Case nAttackTypeMUD
    Case 0: GetSpellAttackType = "Cold"
    Case 1: GetSpellAttackType = "Hot"
    Case 2: GetSpellAttackType = "Stone"
    Case 3: GetSpellAttackType = "Lightning"
    Case 4: GetSpellAttackType = "Normal"
    Case 5: GetSpellAttackType = "Water"
    Case 6: GetSpellAttackType = "Poison"
    Case Else: GetSpellAttackType = nAttackTypeMUD
End Select

Exit Function

error:
Call HandleError("GetSpellAttackType")

End Function

Public Sub MudviewLookup(DatType As MVDatType, ByVal nNum As Long)
Dim sSuffix As String

'   Item = 0
'    Spell = 1
'    Monster = 2
'    Shop = 3
'    Class = 4
'    Race = 5

On Error GoTo error:

Select Case DatType
    Case 0: 'item
        sSuffix = "items.php?disp=1&id=" & nNum
    Case 1: 'spell
        sSuffix = "spells.php?disp=1&id=" & nNum
    Case 2: 'monster
        sSuffix = "monsters.php?disp=1&id=" & nNum
    Case 3: 'shop
        sSuffix = "shops.php?disp=1&id=" & nNum
    Case 4: 'class
        sSuffix = "classes.php"
    Case 5: 'race
        sSuffix = "races.php"
End Select

Call ShellExecute(0&, "open", "http://mudview.mudinfo.net/" & sSuffix, vbNullString, vbNullString, vbNormalFocus)

Exit Sub

error:
Call HandleError("MudviewLookup")

End Sub




Public Function GetArmourType(ByVal nNum As Integer) As String
On Error GoTo error:

Select Case nNum
    Case 0: GetArmourType = "Natural"
    Case 1: GetArmourType = "Silk"      '"Robes"
    Case 2: GetArmourType = "Ninja"     '"Padded"
    '"Soft Leather","Soft Studded","Rigid Leather","Rigid Studded"
    'Case 3: GetArmourType = "Leather(S)"
    'Case 4: GetArmourType = "Leather(SS)"
    'Case 5: GetArmourType = "Leather(R)"
    'Case 6: GetArmourType = "Leather(RS)"
    Case 3 To 6: GetArmourType = "Leather"
    Case 7: GetArmourType = "Chainmail"
    Case 8: GetArmourType = "Scalemail"
    Case 9: GetArmourType = "Platemail"
    Case Else: GetArmourType = "Unknown (" & nNum & ")"
End Select

Exit Function

error:
Call HandleError("GetArmourType")
End Function

Public Function GetWeaponType(ByVal nNum As Integer) As String
On Error GoTo error:

Select Case nNum
    Case 0: GetWeaponType = "1H Blunt"
    Case 1: GetWeaponType = "2H Blunt"
    Case 2: GetWeaponType = "1H Sharp"
    Case 3: GetWeaponType = "2H Sharp"
    Case Else: GetWeaponType = "Unknown (" & nNum & ")"
End Select

Exit Function

error:
Call HandleError("GetWeaponType")
End Function

Public Function GetClassWeaponType(ByVal nNum As Integer) As String
On Error GoTo error:

Select Case nNum
    Case 0: GetClassWeaponType = "1H Blunt"
    Case 1: GetClassWeaponType = "2H Blunt"
    Case 2: GetClassWeaponType = "1H Sharp"
    Case 3: GetClassWeaponType = "2H Sharp"
    Case 4: GetClassWeaponType = "Any 1H"
    Case 5: GetClassWeaponType = "Any 2H"
    Case 6: GetClassWeaponType = "Any Sharp"
    Case 7: GetClassWeaponType = "Any Blunt"
    Case 8: GetClassWeaponType = "Any Weapon"
    Case 9: GetClassWeaponType = "Staff"
    Case Else: GetClassWeaponType = "Unknown (" & nNum & ")"
End Select

Exit Function

error:
Call HandleError("GetClassWeaponType")
End Function

Public Function GetWornType(ByVal nNum As Integer) As String
On Error GoTo error:

Select Case nNum
    Case 0: GetWornType = "Nowhere"
    Case 1: GetWornType = "Everywhere"
    Case 2: GetWornType = "Head"
    Case 3: GetWornType = "Hands"
    Case 4: GetWornType = "Finger"
    Case 5: GetWornType = "Feet"
    Case 6: GetWornType = "Arms"
    Case 7: GetWornType = "Back"
    Case 8: GetWornType = "Neck"
    Case 9: GetWornType = "Legs"
    Case 10: GetWornType = "Waist"
    Case 11: GetWornType = "Torso"
    Case 12: GetWornType = "Off-Hand"
    Case 13: GetWornType = "Finger"
    Case 14: GetWornType = "Wrist"
    Case 15: GetWornType = "Ears"
    Case 16: GetWornType = "Worn"
    Case 17: GetWornType = "Wrist"
    Case 18: GetWornType = "Eyes"
    Case 19: GetWornType = "Face"
    Case Else: GetWornType = "Unknown (" & nNum & ")"
End Select

Exit Function

error:
Call HandleError("GetWornType")
End Function

Public Function GetItemType(ByVal ItemType As Integer) As String
On Error GoTo error:

Select Case ItemType
    Case 0: GetItemType = "Armour"
    Case 1: GetItemType = "Weapon"
    Case 2: GetItemType = "Projectile"
    Case 3: GetItemType = "Sign"
    Case 4: GetItemType = "Food"
    Case 5: GetItemType = "Drink"
    Case 6: GetItemType = "Light"
    Case 7: GetItemType = "Key"
    Case 8: GetItemType = "Container"
    Case 9: GetItemType = "Scroll"
    Case 10: GetItemType = "Special"
    Case Else: GetItemType = ItemType
End Select

Exit Function

error:
Call HandleError("GetItemType")
End Function

'Public Function GetItemCostType(ByVal CostType As Integer) As String
'On Error GoTo Error:
'
'Select Case CostType
'    Case 0: GetItemCostType = "Copper"
'    Case 1: GetItemCostType = "Silver"
'    Case 2: GetItemCostType = "Gold"
'    Case 3: GetItemCostType = "Platinum"
'    Case 4: GetItemCostType = "Runic"
'    Case Else: GetItemCostType = CostType
'End Select
'
'Exit Function
'
'Error:
'Call HandleError("GetItemCostType")
'End Function

Public Function GetCostType(ByVal nNum As Integer) As String
On Error GoTo error:

Select Case nNum
    Case 0: GetCostType = "Copper"
    Case 1: GetCostType = "Silver"
    Case 2: GetCostType = "Gold"
    Case 3: GetCostType = "Platinum"
    Case 4: GetCostType = "Runic"
    Case Else: GetCostType = "Unknown (" & nNum & ")"
End Select

Exit Function

error:
Call HandleError("GetCostType")
End Function

Public Function GetSpellTargets(ByVal nNum As Integer) As String
On Error GoTo error:

Select Case nNum
    Case 0: GetSpellTargets = "User"
    Case 1: GetSpellTargets = "Self"
    Case 2: GetSpellTargets = "Self or User"
    Case 3: GetSpellTargets = "Divided Area (not self)"
    Case 4: GetSpellTargets = "Monster"
    Case 5: GetSpellTargets = "Divided Area (incl self)"
    Case 6: GetSpellTargets = "Any"
    Case 7: GetSpellTargets = "Item"
    Case 8: GetSpellTargets = "Monster or User"
    Case 9: GetSpellTargets = "Divided Attack Area"
    Case 10: GetSpellTargets = "Divided Party Area"
    Case 11: GetSpellTargets = "Full Area"
    Case 12: GetSpellTargets = "Full Attack Area"
    Case 13: GetSpellTargets = "Full Party Area"
    Case Else: GetSpellTargets = "Unknown (" & nNum & ")"
End Select

Exit Function

error:
Call HandleError("GetSpellTargets")

End Function

Public Function GetShopType(ByVal nNum As Long) As String
On Error GoTo error:

Select Case nNum
    Case 0: GetShopType = "General"
    Case 1: GetShopType = "Weapons"
    Case 2: GetShopType = "Armour"
    Case 3: GetShopType = "Items"
    Case 4: GetShopType = "Spells"
    Case 5: GetShopType = "Hospital"
    Case 6: GetShopType = "Tavern"
    Case 7: GetShopType = "Bank"
    Case 8: GetShopType = "Training"
    Case 9: GetShopType = "Inn"
    Case 10: GetShopType = "Specific"
    Case 11: GetShopType = "Gang Shop"
    Case 12: GetShopType = "Deed Shop"
    Case Else: GetShopType = "Unknown (" & nNum & ")"
End Select

Exit Function

error:
Call HandleError("GetShopType")
End Function

Public Function GetMonAttackType(ByVal nNum As Integer) As String
On Error GoTo error:

Select Case nNum
    Case 0: GetMonAttackType = "None"
    Case 1: GetMonAttackType = "Normal"
    Case 2: GetMonAttackType = "Spell"
    Case 3: GetMonAttackType = "Rob"
    Case Else: GetMonAttackType = "Unknown (" & nNum & ")"
End Select

Exit Function

error:
Call HandleError("GetMonAttackType")
End Function

Public Function GetMonType(ByVal nNum As Integer) As String
On Error GoTo error:

Select Case nNum
    Case 0: GetMonType = "Solo"
    Case 1: GetMonType = "Leader"
    Case 2: GetMonType = "Follower"
    Case 3: GetMonType = "Stationary"
    Case Else: GetMonType = "Unknown (" & nNum & ")"
End Select

Exit Function

error:
Call HandleError("GetMonType")
End Function

Public Function GetMonAlignment(ByVal nNum As Integer) As String
On Error GoTo error:

Select Case nNum
    Case 0: GetMonAlignment = "Good"
    Case 1: GetMonAlignment = "Evil"
    Case 2: GetMonAlignment = "Chaotic Evil"
    Case 3: GetMonAlignment = "Neutral"
    Case 4: GetMonAlignment = "Lawful Good"
    Case 5: GetMonAlignment = "Neutral Evil"
    Case 6: GetMonAlignment = "Lawful Evil"
    Case Else: GetMonAlignment = "Unknown (" & nNum & ")"
End Select

Exit Function

error:
Call HandleError("GetMonAlignment")
End Function

Public Function GetMagery(ByVal nNum As Integer, Optional ByVal nLevel As Integer) As String
On Error GoTo error:

Select Case nNum
    Case 0: GetMagery = "None"
    Case 1: GetMagery = "Mage"
    Case 2: GetMagery = "Priest"
    Case 3: GetMagery = "Druid"
    Case 4: GetMagery = "Bard"
    Case 5: GetMagery = "Kai"
    Case Else: GetMagery = "Unknown (" & nNum & ")"
End Select

If Not nNum = 0 Then
    GetMagery = GetMagery & "-" & nLevel
End If

Exit Function

error:
Call HandleError("GetMagery")

End Function

Public Function TestPasteChar(ByVal sTestChar As String) As Boolean
On Error GoTo error:

TestPasteChar = True

Select Case LCase(sTestChar)
    Case "a":
    Case "e":
    Case "i":
    Case "o":
    Case "u":
    Case "y":
    
    Case "b":
    Case "c":
    Case "d":
    Case "f":
    Case "g":
    Case "h":
    Case "j":
    Case "k":
    Case "l":
    Case "m":
    Case "n":
    Case "p":
    Case "q":
    Case "r":
    Case "s":
    Case "t":
    Case "v":
    Case "w":
    Case "x":
    Case "z":
    
    Case "(":
    Case ")":
    
    Case "-":
    Case "_":
    Case ",":
    Case ":":
    Case " ":
    Case "'":
    Case """":
    Case ".":
    Case "`":
    
    Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0":
    
    Case Else: TestPasteChar = False
End Select

Exit Function
error:
Call HandleError("TestPasteChar")
End Function
Public Function TestAlphaChar(ByVal sTestChar As String) As Boolean
On Error GoTo error:

TestAlphaChar = True

Select Case LCase(sTestChar)
    Case "a":
    Case "e":
    Case "i":
    Case "o":
    Case "u":
    Case "y":
    
    Case "b":
    Case "c":
    Case "d":
    Case "f":
    Case "g":
    Case "h":
    Case "j":
    Case "k":
    Case "l":
    Case "m":
    Case "n":
    Case "p":
    Case "q":
    Case "r":
    Case "s":
    Case "t":
    Case "v":
    Case "w":
    Case "x":
    Case "z":
    
    Case Else: TestAlphaChar = False
End Select

Exit Function
error:
Call HandleError("TestAlphaChar")
End Function

Public Function GetAbilityList() As String()
On Error GoTo error:
Dim sArr() As String, x As Integer, nMax As Integer

If bGreaterMUD Then
    nMax = 1120
Else
    nMax = 200
End If

ReDim sArr(nMax)
For x = 1 To nMax
    sArr(x) = GetAbilityName(x, True)
    If sArr(x) = "" Or sArr(x) = "Ability " & x Then
        If x <= 200 Then
            sArr(x) = "[Ability " & x & "]"
        Else
            sArr(x) = ""
        End If
    ElseIf Not bHideRecordNumbers Then
        sArr(x) = sArr(x) & " (" & x & ")"
    End If
Next x

out:
On Error Resume Next
GetAbilityList = sArr
Exit Function
error:
Call HandleError("GetAbilityList")
Resume out:
End Function

Public Function GetSpellCastChance(Optional ByVal nDifficulty As Integer, Optional ByVal nSpellcasting As Integer, _
                                    Optional ByVal bKai As Boolean, Optional ByVal nSpell As Long) As Integer
On Error GoTo error:
Dim nCastChance As Integer

If nDifficulty = 0 And nSpell = 0 And nSpellcasting = 0 Then Exit Function
If nDifficulty <> 0 Or nSpell = 0 Then GoTo ready2:
If tabSpells.RecordCount = 0 Then Exit Function

On Error GoTo seek2:
If tabSpells.Fields("Number") = nSpell Then GoTo ready1:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabSpells.Index = "pkSpells"
tabSpells.Seek "=", nSpell
If tabSpells.NoMatch = True Then
    tabSpells.MoveFirst
    Exit Function
End If

ready1:
On Error GoTo error:
nDifficulty = tabSpells.Fields("Diff")
If tabSpells.Fields("Magery") = 5 Then bKai = True

ready2:
If nSpellcasting > 0 And nDifficulty < 200 Then
    nCastChance = nSpellcasting + nDifficulty
    If nCastChance < 0 Then nCastChance = 0
    If bKai Then 'kai
        If nCastChance > 100 Then nCastChance = 100
    Else
        If nCastChance > 98 Then nCastChance = 98
    End If
Else
    nCastChance = 100
End If

GetSpellCastChance = nCastChance
out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetSpellCastChance")
Resume out:
End Function
Public Function GetAbilityName(ByVal nNum As Integer, Optional ByVal bForceAll As Boolean) As String

Select Case nNum
    Case 0: GetAbilityName = "None"
    Case 1: GetAbilityName = "Damage"
    Case 2: GetAbilityName = "AC"
    Case 3: GetAbilityName = "Rcol"
    Case 4: GetAbilityName = "MaxDamage"
    Case 5: GetAbilityName = "Rfir"
    Case 6: GetAbilityName = "Enslave"
    Case 7: GetAbilityName = "DR"
    Case 8: GetAbilityName = "DrainLife"
    Case 9: GetAbilityName = "Shadow"
    Case 10: GetAbilityName = "AC(Blur)"
    Case 11: GetAbilityName = "EnergyLevel"
    Case 12: GetAbilityName = "Summon"
    Case 13: GetAbilityName = "Illu"
    Case 14: GetAbilityName = "RoomIllu"
    Case 15: GetAbilityName = "Alterhunger"
    Case 16: GetAbilityName = "Alterthirst"
    Case 17: GetAbilityName = "Damage(-MR)"
    Case 18: GetAbilityName = "Heal"
    Case 19: GetAbilityName = "Poison"
    Case 20: GetAbilityName = "CurePoison"
    Case 21: GetAbilityName = "ImmuPoison"
    Case 22: GetAbilityName = "Accuracy"
    Case 23: GetAbilityName = "AffectsUndeadOnly"
    Case 24: GetAbilityName = "ProtEvil"
    Case 25: GetAbilityName = "ProtGood"
    Case 26: GetAbilityName = "DetectMagic"
    Case 27: GetAbilityName = "Stealth"
    Case 28: GetAbilityName = "Magical"
    Case 29: GetAbilityName = "Punch"
    Case 30: GetAbilityName = "Kick"
    Case 31: GetAbilityName = "Bash"
    Case 32: GetAbilityName = "Smash"
    Case 33: GetAbilityName = "Killblow"
    Case 34: GetAbilityName = "Dodge"
    Case 35: GetAbilityName = "JumpKick"
    Case 36: GetAbilityName = "M.R."
    Case 37: GetAbilityName = "Picklocks"
    Case 38: GetAbilityName = "Tracking"
    Case 39: GetAbilityName = "Thievery"
    Case 40: GetAbilityName = "FindTraps"
    Case 41: GetAbilityName = "DisarmTraps"
    Case 42: GetAbilityName = "LearnSp"
    Case 43: GetAbilityName = "CastsSp"
    Case 44: GetAbilityName = "Intel"
    Case 45: GetAbilityName = "Wisdom"
    Case 46: GetAbilityName = "Strength"
    Case 47: GetAbilityName = "Health"
    Case 48: GetAbilityName = "Agility"
    Case 49: GetAbilityName = "Charm"
    Case 50: GetAbilityName = "MageBaneQuest"
    Case 51: GetAbilityName = "AntiMagic"
    Case 52: GetAbilityName = "EvilInCombat"
    Case 53: GetAbilityName = "BlindingLight"
    Case 54: GetAbilityName = "IlluTarget"
    Case 55: GetAbilityName = "AlterGeneralLightDuration"
    Case 56: GetAbilityName = "RechargeItem"
    Case 57: GetAbilityName = "SeeHidden"
    Case 58: GetAbilityName = "Crits"
    Case 59: GetAbilityName = "ClassOk"
    Case 60: GetAbilityName = "Fear"
    Case 61: GetAbilityName = "AffectExit"
    Case 62: GetAbilityName = "EvilChance"
    Case 63: GetAbilityName = "Experience"
    Case 64: GetAbilityName = "AddCP"
    Case 65: GetAbilityName = "ResistStone"
    Case 66: GetAbilityName = "Rlit"
    Case 67: GetAbilityName = "Quickness"
    Case 68: GetAbilityName = "Slowness"
    Case 69: GetAbilityName = "MaxMana"
    Case 70: GetAbilityName = "S.C."
    Case 71: GetAbilityName = "Confusion"
    Case 72: GetAbilityName = "DamageShield"
    Case 73: GetAbilityName = "Dispell"
    Case 74: GetAbilityName = "HoldPerson"
    Case 75: GetAbilityName = "Paralyze"
    Case 76: GetAbilityName = "Mute"
    Case 77: GetAbilityName = "Percep"
    Case 78: GetAbilityName = "Animal"
    Case 79: GetAbilityName = "MageBind"
    Case 80: GetAbilityName = "AffectsAnimalsOnly"
    Case 81: GetAbilityName = "Freedom"
    Case 82: GetAbilityName = "Cursed"
    Case 83: GetAbilityName = "CURSED"
    Case 84: GetAbilityName = "Rcrs"
    Case 85: GetAbilityName = "Shatter"
    Case 86: GetAbilityName = "Quality"
    Case 87: GetAbilityName = "Speed"
    Case 88: GetAbilityName = "Alter HP"
    Case 89: GetAbilityName = "PunchAcc"
    Case 90: GetAbilityName = "KickAcc"
    Case 91: GetAbilityName = "JumpKAcc"
    Case 92: GetAbilityName = "PunchDmg"
    Case 93: GetAbilityName = "KickDmg"
    Case 94: GetAbilityName = "JumpKDmg"
    Case 95: GetAbilityName = "Slay"
    Case 96: GetAbilityName = "Encum"
    Case 97: GetAbilityName = "GoodOnly"
    Case 98: GetAbilityName = "EvilOnly"
    Case 99: GetAbilityName = "AlterDRpercent"
    Case 100: GetAbilityName = "LoyalItem"
    Case 101:
        If Not bForceAll Then
            Exit Function
        Else
            GetAbilityName = "ConfuseMsg"
        End If
    Case 102: GetAbilityName = "RaceStealth"
    Case 103: GetAbilityName = "ClassStealth"
    Case 104: GetAbilityName = "DefenseModifier"
    Case 105: GetAbilityName = "Accuracy2"
    Case 106: GetAbilityName = "Accuracy3"
    Case 107: GetAbilityName = "BlindUser"
    Case 108: GetAbilityName = "AffectsLivingOnly"
    Case 109: GetAbilityName = "NonLiving"
    Case 110: GetAbilityName = "NotGood"
    Case 111: GetAbilityName = "NotEvil"
    Case 112: GetAbilityName = "NeutralOnly"
    Case 113: GetAbilityName = "NotNeutral"
    Case 114: GetAbilityName = "%Spell"
    Case 115:
        If Not bForceAll Then
            Exit Function
        Else
            GetAbilityName = "DescMsg"
        End If
    Case 116: GetAbilityName = "BSAccu"
    Case 117: GetAbilityName = "BsMinDmg"
    Case 118: GetAbilityName = "BsMaxDmg"
    Case 119: GetAbilityName = "Del@Maint"
    Case 120:
        If Not bForceAll Then
            Exit Function
        Else
            GetAbilityName = "StartMsg"
        End If
    Case 121: GetAbilityName = "Recharge"
    Case 122: GetAbilityName = "RemovesSpell"
    Case 123: GetAbilityName = "HPRegen"
    Case 124: GetAbilityName = "NegateAbility"
    Case 125: GetAbilityName = "IceSorcQuest"
    Case 126: GetAbilityName = "GoodQuest"
    Case 127: GetAbilityName = "NeutralQuest"
    Case 128: GetAbilityName = "EvilQuest"
    Case 129: GetAbilityName = "DarkDruidQuest"
    Case 130: GetAbilityName = "BloodChampQuest"
    Case 131: GetAbilityName = "SheDragonQuest"
    Case 132: GetAbilityName = "WereratQuest"
    Case 133: GetAbilityName = "PhoenixQuest"
    Case 134: GetAbilityName = "DaoLordQuest"
    Case 135: GetAbilityName = "MinLevel"
    Case 136: GetAbilityName = "MaxLevel"
    Case 137: GetAbilityName = "Shock"
    Case 138: GetAbilityName = "RoomVisible"
    Case 139: GetAbilityName = "SpellImmu"
    Case 140: GetAbilityName = "TeleportRoom"
    Case 141: GetAbilityName = "TeleportMap"
    Case 142: GetAbilityName = "HitMagic"
    Case 143: GetAbilityName = "ClearItem"
    Case 144:
        If Not bForceAll Then
            Exit Function
        Else
            GetAbilityName = "NonMagicalSpell"
        End If
    Case 145: GetAbilityName = "ManaRgn"
    Case 146: GetAbilityName = "MonsGuards"
    Case 147: GetAbilityName = "ResistWater"
    Case 148: GetAbilityName = "TextBlock" '1'1'1'1
    Case 149: GetAbilityName = "Remove@Maint"
    Case 150: GetAbilityName = "HealMana"
    Case 151: GetAbilityName = "EndCast"
    Case 152: GetAbilityName = "Rune"
    Case 153: GetAbilityName = "KillSpell"
    Case 154: GetAbilityName = "Visible@Maint"
    Case 155:
        If Not bForceAll Then
            Exit Function
        Else
            GetAbilityName = "DeathText"
        End If
    Case 156: GetAbilityName = "QuestItem"
    Case 157: GetAbilityName = "ScatterItems"
    Case 158: GetAbilityName = "ReqToHit"
    Case 159: GetAbilityName = "KaiBind"
    Case 160: GetAbilityName = "GiveTempSpell"
    Case 161: GetAbilityName = "OpenDoor"
    Case 162: GetAbilityName = "Lore"
    Case 163: GetAbilityName = "SpellComponent"
    Case 164: GetAbilityName = "EndCast%"
    Case 165: GetAbilityName = "AlterSpDmg"
    Case 166: GetAbilityName = "AlterSpLength"
    Case 167: GetAbilityName = "UnEquipItem"
    Case 168: GetAbilityName = "EquipItem"
    Case 169: GetAbilityName = "CannotWearLocation"
    Case 170: GetAbilityName = "Sleep"
    Case 171: GetAbilityName = "Invisibility"
    Case 172: GetAbilityName = "SeeInvisible"
    Case 173: GetAbilityName = "Scry"
    Case 174: GetAbilityName = "StealMana"
    Case 175: GetAbilityName = "StealHPtoMP"
    Case 176: GetAbilityName = "StealMPtoHP"
    Case 177: GetAbilityName = "SpellColours"
    Case 178: GetAbilityName = "Shadowform"
    Case 179: GetAbilityName = "FindTrapsValue"
    Case 180: GetAbilityName = "PickLocksValue"
    Case 181: GetAbilityName = "GHouseDeed"
    Case 182: GetAbilityName = "GHouseTax"
    Case 183: GetAbilityName = "GHouseItem"
    Case 184: GetAbilityName = "GShopItem"
    Case 185: GetAbilityName = "NoAttack"
    Case 186: GetAbilityName = "PerStealth"
    Case 187: GetAbilityName = "Meditate"
    Case Else:
        If bGreaterMUD Then
            Select Case nNum
                Case 188: GetAbilityName = "Unique Pool"
                Case 189: GetAbilityName = "Witchy Badges"
                Case 190: GetAbilityName = "No Stock"
                Case 200: GetAbilityName = "Mandos Quest"
                Case 201: GetAbilityName = "Volums Quest"
                Case 202: GetAbilityName = "Cartographer's Quest"
                Case 203: GetAbilityName = "Loremaster's Quest"
                Case 204: GetAbilityName = "Guildmaster's Bounty Quest"
                Case 205: GetAbilityName = "Darkbane Quest"
                Case 206: GetAbilityName = "Grizzled Ranger"
                Case 207: GetAbilityName = "Amazon Huntress"
                Case 208: GetAbilityName = "Conquest 1"
                Case 209: GetAbilityName = "Conquest 2"
                Case 210: GetAbilityName = "Tarl's Quest"
                Case 211: GetAbilityName = "Tal'kiran passa"
                Case 212: GetAbilityName = "Trendel Quest"
                Case 213: GetAbilityName = "Luca Prodigioourtesan Quest"
                Case 1001: GetAbilityName = "GrantThievery"
                Case 1002: GetAbilityName = "GrantTraps"
                Case 1003: GetAbilityName = "GrantPicklocks"
                Case 1004: GetAbilityName = "GrantTracking"
                Case 1100: GetAbilityName = "AntiMagicNotOK"
                Case 1101: GetAbilityName = "MeetsReqToHit"
                Case 1103: GetAbilityName = "Shadow Rest"
                Case 1104: GetAbilityName = "AlterSpellHeal"
                Case 1105: GetAbilityName = "AlterSpells"
                Case 1106: GetAbilityName = "AlterSpellBuffs"
                Case 1107: GetAbilityName = "NoAutoLearn"
                Case 1108: GetAbilityName = "NotForPVP"
                Case 1109: GetAbilityName = "Enchant"
                Case 1110: GetAbilityName = "BSDR"
                Case 1111: GetAbilityName = "Absorb"
                Case 1112: GetAbilityName = "Patrol"
                Case 1113: GetAbilityName = "Vile Ward"
                Case 1114: GetAbilityName = "Cast on Kill"
                Case 1115: GetAbilityName = "NoFirstKill Drop"
                Case 1116:
                    If Not bForceAll Then
                        Exit Function
                    Else
                        GetAbilityName = "AccountVerified"
                    End If
                Case 1117: GetAbilityName = "Not Sellable"
                Case 1118: GetAbilityName = "NoRandomRegen"
                Case 1119: GetAbilityName = "Del@Ganghouse"
                Case Else: GetAbilityName = "Ability " & nNum
            End Select
        Else
            GetAbilityName = "Ability " & nNum
        End If
End Select

End Function

Public Function CalcMoneyRequiredToTrain(ByVal nLevel As Currency, _
    ByVal nMarkup As Currency) As Currency
'{ Calculates the copper farthings needed to train for a specific level }
' function  CalcMoneyRequiredToTrain(Level, Markup: integer): longword;
' begin
'   Result := (longword((Level * 5) * (Markup + 100)) div 100) * 10;
' end;
On Error GoTo error:

CalcMoneyRequiredToTrain = Fix((nLevel * 5) * (nMarkup + 100) / 100) * 10

Exit Function
error:
Call HandleError("CalcMoneyRequiredToTrain")
End Function

Public Function CalcRestingRate(ByVal nLevel As Long, ByVal nHealth As Long, _
    Optional ByVal nHPRegenPercent As Long, Optional ByVal bResting As Boolean) As Long
'{ Calculates HP regen for a given Level, Health, HPRegen and Resting state }
'function  CalcHPRegen(Level, HEA: integer; HPRegen: integer = 0; Resting: boolean = False): integer;
'begin
'  Result := (((Level + 20) * HEA) div 750);
'  If (Result < 1) Then
'    Result := 1;
'  If (Resting) Then
'    Result := Result * 3;
'
'  Result := ((HPRegen + 100) * Result) div 100;
'end;

'resting rate ticks every 20 seconds / 4 rounds
'non-resting rate ticks every 30 seconds / 6 rounds

Dim nHPRegen As Long
On Error GoTo error:

nHPRegen = Fix(((nLevel + 20) * nHealth) / 750)
If nHPRegen < 1 Then nHPRegen = 1

If bResting Then nHPRegen = nHPRegen * 3

CalcRestingRate = Fix(((nHPRegenPercent + 100) * nHPRegen) / 100)

Exit Function
error:
If Err.Number = 6 Then
    CalcRestingRate = -1
Else
    Call HandleError("CalcRestingRate")
End If

End Function

Public Function CalcBSDamage(ByVal nLevel As Integer, ByVal nStealth As Integer, _
    ByVal nDMG As Integer, ByVal nBsDmgMod As Integer, ByVal bClassStealth As Boolean) As Long
'const
'  ST_RACE_STEALTH  = 1;
'  ST_CLASS_STEALTH = 2;
'  ST_BOTH_STEALTH  = ST_RACE_STEALTH or ST_CLASS_STEALTH;
'
'{ Calculates Backstab damage for a given Level, Stealth, Dmg, BsDmgMod and
'  StealthType }
'function  CalcBSDamage(Level, Stealth, Dmg, BsDmgMod, StealthType: integer): integer;
'begin
'  Result := ((Level * 2) + (Stealth div 10)) + (Dmg * 2) + BsDmgMod;
'  If (StealthType = ST_RACE_STEALTH) Then
'    Result := (Result * 75) div 100;
'
'  Result := ((Level + 100) * Result) div 100;
'end;
On Error GoTo error:

'Debug.Print ""
'Debug.Print "Debug-Level: " & nLevel
'Debug.Print "Debug-Stealth: " & nStealth
'Debug.Print "Debug-DMG: " & nDMG
'Debug.Print "Debug-BsDmgMod: " & nBsDmgMod
'Debug.Print "Debug-ClassStealth: " & IIf(bClassStealth, "True", "False")
'Debug.Print ""

CalcBSDamage = (nLevel * 2) + Fix(nStealth / 10) + (nDMG * 2) + nBsDmgMod
If Not bClassStealth Then CalcBSDamage = Fix((CalcBSDamage * 75) / 100)
CalcBSDamage = Fix(((nLevel + 100) * CalcBSDamage) / 100)

out:
Exit Function
error:
Call HandleError("CalcBSDamage")
Resume out:
End Function


Public Function CalcManaRegen(ByVal nLevel As Long, ByVal nINT As Long, ByVal nWil As Long, _
    ByVal nCHA As Long, ByVal nMagicLVL As Long, ByVal nMagicType As enmMagicEnum, _
    Optional ByVal nMPRegen As Long, Optional ByVal bMeditating As Boolean) As Currency
On Error GoTo error:
' { Calculates mana regen from a given Level, INT, WIL, CHA, MagicLevel,
'   MagicType and optional MPRegen }
' function  CalcMPRegen(Level, INT, WIL, CHA, MagicLevel: integer; MagicType: TMagicType; MPRegen: integer = 0): integer;
' begin
'   If (MagicType <> mtKai) Then begin
'     case MagicType of
'       mtMage: Result := INT;
'     mtPriest: Result := WIL;
'      mtDruid: Result := (INT + WIL) div 2;
'       mtBard: Result := CHA;
'       Else
'         Result := 0;
'     end;
'     Result := (((Level + 20) * Result) * (MagicLevel + 2)) div 1650;
'   end else begin
'     Result := 1;    // Mystics are always 1
'   end;
'   Result := ((MPRegen + 100) * Result) div 100;
' end;

Select Case nMagicType
    Case 0: 'none
        Exit Function
    Case 1: 'mage
        CalcManaRegen = nINT
    Case 2: 'priest
        CalcManaRegen = nWil
    Case 3: 'druid
        CalcManaRegen = Fix((nINT + nWil) / 2)
    Case 4: 'bard
        CalcManaRegen = nCHA
    Case 5: 'kai
        CalcManaRegen = Fix(((nMPRegen + 100) * 1) / 100)
        Exit Function
    Case Else:
        Exit Function
End Select

CalcManaRegen = Fix((((nLevel + 20) * CalcManaRegen) * (nMagicLVL + 2)) / 1650)
If bMeditating Then Exit Function
CalcManaRegen = Fix(((nMPRegen + 100) * CalcManaRegen) / 100)

Exit Function
error:
Call HandleError("CalcManaRegen")
End Function

Public Function CalcMR(ByVal nINT As Integer, ByVal nWis As Integer, Optional ByVal nModifiers As Integer) As Long
On Error GoTo error:

CalcMR = Fix((nINT + (nWis * 3)) / 4) + nModifiers

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalcMR")
Resume out:
End Function

Public Function CalcMaxHP(ByVal nRandom As Long, ByVal nLevel As Long, _
    ByVal nHealth As Long, ByVal nMinHPPerLevel As Long) As Long
'{ Calculates MaxHP for a given Random value (*), Level, HEA and MinHPPerLevel
'
'   HOW THIS WORKS--
'
'     At character creation, the player is given their maximum "range" HP roll.
'     This range is what is encoded into the MajorMUD classes database.  For
'     example, on a Druid, the range is 3, so at level 1, the 'Random' portion
'     as listed below is set to 3.  Each time the player trains, a random number
'     is generated between 0 and "range"-- going back to Druids again, this means
'     the RNG could return 0, 1, 2, or 3, and this value would be summed with
'     the already existing 'Random' value.  To determine the maximum possible HP
'     for a given level then, you take the "range" and multiply it by the level
'     of the char you want the max for, then pass that result as the 'Random'
'     value for the function below.  To determine the minimum possible HP for a
'     given level, you take the "range" and pass it as the 'Random' value for the
'     function below.  You do NOT multiply it because for minimum rolls you'd
'     have received all 0's-- the only reason you pass the "range" is for the
'     reason stated above: at level 1, MajorMUD gives you the maximum "range"
'     roll your class can get.
'
'     Penalties or bonuses such as HP per level (Halfling, Half-Ogre), or Ability
'     modifications such as +HP on sunstone wristbands are figured *after* this
'     formula is applied. }
' function  CalcHP(Random, Level, HEA, MinHPPerLevel: integer): integer;
' begin
'   Result := ((HEA div 2) + Level * MinHPPerLevel) + (((HEA - 50) * Level) div 16) + Random;
' end;

On Error GoTo error:

CalcMaxHP = (Fix(nHealth / 2) + nLevel * nMinHPPerLevel) _
    + Fix(((nHealth - 50) * nLevel) / 16) + nRandom

Exit Function

error:
Call HandleError("CalcMaxHP")

End Function

Public Function CalcMaxMana(ByVal nLevel As Long, ByVal nMagicLevel As Long) As Long
' { Calculates the maximum Mana for a given Level and MagicLevel }
' function  CalcMP(Level, MagicLevel: integer): integer;
' begin
'   Result := ((MagicLevel * Level) * 2) + 6;
' end;
On Error GoTo error:

CalcMaxMana = ((nMagicLevel * nLevel) * 2) + 6

Exit Function

error:
Call HandleError("CalcMaxMana")
End Function

Public Function CalcSpellCasting(ByVal nLevel As Long, ByVal nINT As Long, ByVal nWil As Long, _
    ByVal nCHA As Long, ByVal nMagicLVL As Long, ByVal nMagicType As enmMagicEnum) As Long
' { Calculates SC from a given Level, MagicLevel, INT, WIL, CHA and MagicType }
' function  CalcSC(Level, MagicLevel, INT, WIL, CHA: integer; MagicType: TMagicType): integer;
' begin
'   case MagicType of
'     mtMage: Result := (((INT * 3) + WIL) div 6) + (Level * 2) + (MagicLevel * 5);
'   mtPriest: Result := (((WIL * 3) + INT) div 6) + (Level * 2) + (MagicLevel * 5);
'    mtDruid: Result := ((WIL + INT) div 3) + (Level * 2) + (MagicLevel * 5);
'     mtBard: Result := (((CHA * 3) + WIL) div 6) + (Level * 2) + (MagicLevel * 5);
'      mtKai: Result := 500 + (Level * 2) + (MagicLevel * 5);
'     Else
'       Result := 0;
'   end;
' end;
On Error GoTo error:

Select Case nMagicType
    Case 0: 'none
        Exit Function
    Case 1: 'mage
        CalcSpellCasting = Fix(((nINT * 3) + nWil) / 6) + (nLevel * 2) + (nMagicLVL * 5)
    Case 2: 'priest
        CalcSpellCasting = Fix(((nWil * 3) + nINT) / 6) + (nLevel * 2) + (nMagicLVL * 5)
    Case 3: 'druid
        CalcSpellCasting = Fix((nWil + nINT) / 3) + (nLevel * 2) + (nMagicLVL * 5)
    Case 4: 'bard
        CalcSpellCasting = Fix(((nCHA * 3) + nWil) / 6) + (nLevel * 2) + (nMagicLVL * 5)
    Case 5: 'kai
        CalcSpellCasting = 500 + (nLevel * 2) + (nMagicLVL * 5)
    Case Else:
        Exit Function
End Select


Exit Function

error:
Call HandleError("CalcSpellCasting")
End Function

Public Function GetEncumPercents(ByVal nTotalEncum As Long) As String
Dim x As Double
On Error GoTo error:
'& "/" & nTotalEncum
If Not nTotalEncum = 0 Then
    GetEncumPercents = "Light @ " & Fix(nTotalEncum * 0.17) + 1 & "/" & nTotalEncum & vbCrLf _
            & "Medium @ " & Fix(nTotalEncum * 0.34) + 1 & "/" & nTotalEncum & vbCrLf _
            & "Heavy @ " & Fix(nTotalEncum * 0.67) + 1 & "/" & nTotalEncum
    
    GetEncumPercents = GetEncumPercents & vbCrLf
    
    For x = 0.1 To 0.9 Step 0.1
        GetEncumPercents = GetEncumPercents & vbCrLf & (x * 100) & "% @ " & Fix(nTotalEncum * x) + 1 '& "/" & nTotalEncum
    Next x
Else
    GetEncumPercents = ""
End If

Exit Function

error:
Call HandleError("GetEncumPercents")

End Function

Public Function CalcPicklocks(ByVal nLevel As Long, ByVal nAGL As Long, ByVal nINT As Long) As Long
' { Calculates Picklocks for a given Level, Agility and Intellect }
' function  CalcPicklocks(Level, AGL, INT: integer): integer;
' begin
'   If (Level <= 15) Then
'     Result := Level * 2
'   Else
'     Result := (((Level - 15) div 2) + 15) * 2;
'
'   Result := (((Result * 5) + (AGL + INT)) * 2) div 7;
' end;
If nLevel <= 15 Then
    CalcPicklocks = nLevel * 2
Else
    CalcPicklocks = (Fix((nLevel - 15) / 2) + 15) * 2
End If

CalcPicklocks = Fix((((CalcPicklocks * 5) + (nAGL + nINT)) * 2) / 7)
End Function

Function CalcCPLevel(ByVal nLevel As Long) As Long
'{ Calculates the CP for a level }
' function  CPLevel(Level: integer): integer;
' Var
'   I: integer;
' begin
'   Result := 0;
'   for I := 2 to Level do begin
'     Result := Result + (((I - 1) div 10) * 5) + 10;
'   end;
' end;
Dim i As Long

For i = 1 To nLevel - 1
    CalcCPLevel = CalcCPLevel + (Fix(i / 10) * 5) + 10
Next i

End Function

Public Function CalcTrueAverage(ByVal nSwings As Double, ByVal nHitP As Double, ByVal nHitA As Long, _
    ByVal nCritP As Double, ByVal nCritA As Long, ByVal nExtraP As Double, ByVal nExtraA As Long) As Double
On Error GoTo error:

If nSwings <= 0 Then CalcTrueAverage = -1: Exit Function
If nSwings > 5 Then nSwings = 5

nHitP = nHitP / 100
nCritP = nCritP / 100
nExtraP = nExtraP / 100
'((HIT% * HITAVE) + (CRIT% * CRITAVE) + (HIT% * EXTRA% * EXTRAAVE) + (CRIT% * EXTRA% * EXTRAAVE)) * SWINGS
'CalcTrueAverage = Round(((nHitP * nHitA) + (nCritP * nCritA) + (nHitP * nExtraP * nExtraA) + (nCritP * nExtraP * nExtraA)) * nSwings, 2)
CalcTrueAverage = Round(((nHitP * nHitA) + (nCritP * nCritA) + ((nHitP + nCritP) * nExtraP * nExtraA)) * nSwings, 2)

Exit Function
error:
Call HandleError("CalcTrueAverage")

End Function

Public Function GetQuickAndDeadlyBonus(ByVal nItemNum As Long) As Integer
On Error GoTo error:
Dim nEnergy As Currency, nCombat As Integer, nEncum As Currency

If nItemNum = 0 Or tabItems.RecordCount = 0 Then Exit Function
If frmMain.chkGlobalFilter.Value = 0 Then Exit Function
If frmMain.cmbGlobalClass(0).ListIndex < 0 Then Exit Function
If frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex) < 1 Then Exit Function

On Error GoTo seek2:
If tabItems.Fields("Number") = nItemNum Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabItems.Index = "pkItems"
tabItems.Seek "=", nItemNum
If tabItems.NoMatch = True Then
    tabItems.MoveFirst
    Exit Function
End If

ready:
If tabItems.Fields("StrReq") > val(frmMain.txtCharStats(0).Text) Then Exit Function

nCombat = GetClassCombat(frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex))
nEncum = CalcEncumbrancePercent(val(frmMain.lblInvenCharStat(0).Caption), val(frmMain.lblInvenCharStat(1).Caption))
nEnergy = CalcEnergyUsedWithEncum(nCombat, val(frmMain.txtGlobalLevel(0).Text), tabItems.Fields("Speed"), _
    val(frmMain.txtCharStats(3).Text), val(frmMain.txtCharStats(0).Text), nEncum, tabItems.Fields("StrReq"))

GetQuickAndDeadlyBonus = CalcQuickAndDeadlyBonus(val(frmMain.txtCharStats(3).Text), nEnergy, nEncum)

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetQuickAndDeadlyBonus")
Resume out:
End Function

Public Function CalcQuickAndDeadlyBonus(ByVal nAGL As Currency, ByVal nEU As Currency, _
    ByVal nEncum As Currency) As Currency
On Error GoTo error:
Dim gmudMultiplier As Integer, gmudEnergyRemain As Integer

' { Calculates the critical hit chance bonus for being quick and deadly with a
'   weapon for a previously calculated energy use }
' function  CalcQuickAndDeadlyBonus(AGL, EU, Encumbrance: integer): integer;
' begin
'   Result := 0;
'   If (EU >= 200) Or (Encumbrance > 66) Then
'     Exit;
'
'   Result := 200 - EU;
'   Result := Result + ((AGL - 50) div 10);
'
' //  Result := ((200 - EU) + ((AGL - 50) div 10));
'   If (Result > 20) Then
'     Result := 20;
'
'   If (Encumbrance >= 33) Then
'     Result := Result div 2;
' end;

CalcQuickAndDeadlyBonus = 0
If (nEU >= 200) Or (nEncum > 66 And Not bGreaterMUD) Then Exit Function

If bGreaterMUD Then
    gmudMultiplier = 50
    gmudEnergyRemain = 1000 - (nEU * 5)
    CalcQuickAndDeadlyBonus = Fix(gmudEnergyRemain / gmudMultiplier)
Else
    CalcQuickAndDeadlyBonus = (200 - nEU) + Fix((nAGL - 50) / 10)
    If (CalcQuickAndDeadlyBonus > 20) Then CalcQuickAndDeadlyBonus = 20
    If (nEncum >= 33) Then CalcQuickAndDeadlyBonus = Fix(CalcQuickAndDeadlyBonus / 2)
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalcQuickAndDeadlyBonus")
Resume out:
End Function

Public Function CalcEncumbrancePercent(ByVal nCurrent As Currency, ByVal nMax As Currency) As Currency
'{ Calculates the encumbrance percentage used for calculating Q&D bonuses and
'  energy used }
'function  CalcEncumbrancePercent(Current, Maximum: integer): integer; begin
'  Result := (Current * 100) div Maximum; end;

If nMax <= 0 Then nMax = 1

CalcEncumbrancePercent = Fix((nCurrent * 100) / nMax)

End Function

Public Function AdjustSpeedForSlowness(ByVal nSpeed As Currency) As Currency
'{ Adjusts the Speed of a weapon for the case where a player has the Slowness
'  flag on them }
'function  AdjustSpeedForSlowness(Speed: integer): integer; begin
'  Result := (Speed * 3) div 2;
'end;

AdjustSpeedForSlowness = Fix((nSpeed * 3) / 2)

End Function

Public Function CalculateStealth(ByVal nLevel As Integer, ByVal nAgility As Integer, ByVal nIntellect As Integer, ByVal nCharm As Integer, _
    ByVal bClassStealth As Boolean, ByVal bRaceStealth As Boolean, Optional ByVal nPlusStealth As Integer) As Integer
On Error GoTo error:
Dim nStealth As Integer

If Not bRaceStealth And Not bClassStealth Then Exit Function

If nLevel <= 15 Then
    nStealth = nLevel * 2
Else
    nStealth = Fix(((nLevel - 15) * 2) / 2) + 30
End If
nStealth = nStealth + 20

nStealth = nStealth + Fix(nAgility / 4)
nStealth = nStealth + Fix(nIntellect / 8)
nStealth = nStealth + Fix(nCharm / 6)

If bRaceStealth And bClassStealth Then
    nStealth = nStealth + 10
ElseIf bRaceStealth Then 'implies Not bClassStealth
    nStealth = nStealth - 15
End If

nStealth = nStealth + nPlusStealth

out:
On Error Resume Next
CalculateStealth = nStealth
Exit Function
error:
Call HandleError("CalculateStealth")
Resume out:
End Function

Public Function CalcEnergyUsed(ByVal nCombat As Currency, ByVal nLevel As Currency, _
    ByVal nAttackSpeed As Currency, ByVal nAGL As Currency, Optional ByVal nSTR As Currency = 0, _
    Optional ByVal nEncum As Currency = -1, Optional ByVal nItemSTR As Currency = 0, _
    Optional ByVal nSpeedAdj As Currency = 0, Optional ByVal bIsBackstab As Boolean) As Currency
'{ Calculates the energy used for a given Combat rating, Level, Speed, AGL, STR,
'  and ItemSTR }
'function  CalcEnergyUsed(Combat, Level, Speed, AGL: integer; STR: integer = 0; ItemSTR: integer = 0): longword; begin
'  Result := longword(Speed * 1000) div (longword((((Level * (Combat + 2)) + 45) * (AGL + 150)) * 1500) div 9000);
'  If (STR < ItemSTR) Then
'    Result := longword(longword((longword(ItemSTR - STR) * 3) + 200) *
'Result) div 200; end;

'CalcEnergyUsed = Fix((nAttackSpeed * 1000) / Fix(((((nLevel * (nCombat + 2)) + 45) * (nAGL + 150)) * 1500) / 9000))

'simplified 2025.03.16:
CalcEnergyUsed = Fix((nAttackSpeed * 1000) / Fix((((nLevel * (nCombat + 2)) + 45) * (nAGL + 150)) / 6))

If nSTR > 0 And nSTR < nItemSTR Then
    CalcEnergyUsed = Fix(((((nItemSTR - nSTR) * 3) + 200) * CalcEnergyUsed) / 200)
End If

If nSpeedAdj > 0 And nSpeedAdj <> 100 And bIsBackstab = False Then
    CalcEnergyUsed = Fix((CalcEnergyUsed * nSpeedAdj) / 100)
End If

If nEncum >= 0 Then CalcEnergyUsed = Fix((CalcEnergyUsed * (Fix(nEncum / 2) + 75)) / 100)

End Function

Public Function CalcEnergyUsedWithEncum(ByVal nCombat As Currency, ByVal nLevel As Currency, _
    ByVal nSpeed As Currency, ByVal nAGL As Currency, ByVal nSTR As Currency, ByVal nEncum As Currency, _
    Optional ByVal nItemSTR As Currency = 0) As Currency
'{ Calculates the energy used for a given Combat rating, Level, Speed, AGL, STR,
'  Encumbrance, and ItemSTR }
'function  CalcEnergyUsedWithEncum(Combat, Level, Speed, AGL, STR: integer;
'Encumbrance: integer; ItemSTR: integer = 0): integer; begin
'  Result := CalcEnergyUsed(Combat, Level, Speed, AGL, STR, ItemSTR);
'  Result := (Result * ((Encumbrance div 2) + 75)) div 100; end;
    
CalcEnergyUsedWithEncum = CalcEnergyUsed(nCombat, nLevel, nSpeed, nAGL, nSTR, , nItemSTR)
CalcEnergyUsedWithEncum = Fix((CalcEnergyUsedWithEncum * (Fix(nEncum / 2) + 75)) / 100)

End Function



Public Function AdjustEnergyUsedWithSpeed(ByVal nEU As Currency, ByVal nSpeed As Currency) As Currency
'{ Adjusts a previously calculated energy use with a specified Speed amount }
'
'function  AdjustEnergyUsedWithSpeed(EU, Speed: integer): integer; begin
'  Result := (EU * Speed) div 100;
'end;

AdjustEnergyUsedWithSpeed = Fix((nEU * nSpeed) / 100)

End Function

Public Function AdjustEnergyUsedWithEncum(ByVal nEU As Currency, ByVal nEncum As Currency) As Currency
'{ Adjusts a previously calculated energy use with a specified Encumbrance
'  amount }
'function  AdjustEnergyUsedWithEncum(EU, Encumbrance: longword): longword; begin
'  Result := (EU * ((Encumbrance div 2) + 75)) div 100; end;

AdjustEnergyUsedWithEncum = Fix((nEU * (Fix(nEncum / 2) + 75)) / 100)

End Function

Public Function CalculateResistDamage(ByVal nDamage As Currency, ByVal nVSMagicResist As Long, Optional ByVal nSpellResistType As Integer = 2, _
    Optional ByVal bDamageResistable As Boolean = True, Optional ByVal bIncludeTotalResist As Boolean, _
    Optional ByVal bVSAntiMagic As Boolean, Optional ByVal nBonusResist As Long) As Currency
On Error GoTo error:
'nSpellResistType: 0-never, 1-antimagic only, 2-everyone
'nBonusResist = matching resist like rfir, rcol, etc
'bDamageResistable = false:ability1 (damage) or true:ability17 (Damage-MR), basically
Dim nDamageResist As Double, nTotalResist As Double

If nVSMagicResist <= 0 Then nVSMagicResist = 1

If nBonusResist > 0 Then
    nDamage = Fix(((100 - nBonusResist) * nDamage) / 100)
End If

If bDamageResistable Then
    If bVSAntiMagic Then
        nDamageResist = Fix(nVSMagicResist / 2)
        If nDamageResist > 75 Then nDamageResist = 75
    ElseIf nVSMagicResist > 51 Then
        nDamageResist = Fix((nVSMagicResist - 50) / 2)
        If nDamageResist > 50 Then nDamageResist = 50
    End If
    
    If nDamageResist > 0 Then
        nDamage = nDamage * (1 - (nDamageResist / 100))
    ElseIf Not bVSAntiMagic And nVSMagicResist < 50 Then '+damage for mr < 50
        nDamage = nDamage + ((nDamage * (50 - nVSMagicResist)) / 100)
    End If
End If

If bIncludeTotalResist And nVSMagicResist > 1 And (nSpellResistType = 2 Or (bVSAntiMagic And nSpellResistType = 1)) Then
    nTotalResist = Fix(nVSMagicResist / 2)
    If nTotalResist > 98 Then nTotalResist = 98
    nDamage = nDamage * (1 - (nTotalResist / 100))
End If

CalculateResistDamage = Round(nDamage)

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalculateResistDamage")
Resume out:
End Function

