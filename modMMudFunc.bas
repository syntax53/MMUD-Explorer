Attribute VB_Name = "modMMudFunc"
Option Explicit
Option Base 0

Public Type RoomExitType
    Map As Long
    Room As Long
    ExitType As String
End Type

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
                sTemp = PullSpellEQ(False, , (Val(Mid(sArr(x), 6))))
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
            GetAbilityStats = GetAbilityStats & " [" & GetSpellName(nValue, bHideRecordNumbers) & ", " & PullSpellEQ(bCalcSpellLevel, 0, nValue, IIf(LV Is Nothing, Nothing, LV), , , bPercentColumn) & "]"
        Case 73, 124: 'dispell magic, negateabil
            GetAbilityStats = GetAbilityStats & " (" & GetAbilityName(nValue) & ")"
        Case 151: 'endcast
            GetAbilityStats = GetAbilityStats & " [" & GetSpellName(nValue, bHideRecordNumbers) & ", " & PullSpellEQ(bCalcSpellLevel, 0, nValue, IIf(LV Is Nothing, Nothing, LV), , , bPercentColumn) & "]"
        Case 59: 'class ok
            GetAbilityStats = GetAbilityStats & " " & GetClassName(nValue)
        Case 146, 12: 'mon guards, summon
            GetAbilityStats = GetAbilityStats & " " & GetMonsterName(nValue, bHideRecordNumbers)
        Case 1, 8, 17, 18, 19, 140, 141, 148:
            'NO HEADERS, damage, drain, damage(on armr), poison, heal, teleport room, teleport map, textblocks
            ' *** ALSO ADD THESE TO PullSpellEQ ***
            GetAbilityStats = GetAbilityStats & " " & nValue
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
Dim x As Integer, Y As Integer, i As Integer

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

ExtractMapRoom.Map = Val(Mid(sExit, i, x - 1))

Y = InStr(x, sExit, " ")
If Y = 0 Then
    ExtractMapRoom.Room = Val(Mid(sExit, x + 1))
Else
    ExtractMapRoom.Room = Val(Mid(sExit, x + 1, Y - 1))
    ExtractMapRoom.ExitType = Mid(sExit, Y + 1)
End If

Exit Function

error:
Call HandleError("ExtractMapRoom")

End Function

Public Function CalcDodge(Optional ByVal nCharLevel As Integer, Optional ByVal nAgility As Integer, Optional ByVal nCharm As Integer, Optional ByVal nPlusDodge As Integer, _
    Optional ByVal nCurrentEncum As Integer = 0, Optional ByVal nMaxEncum As Integer = -1) As Integer
On Error GoTo error:
Dim nDodge As Integer, nEncumPct As Integer

nDodge = Fix(nCharLevel / 5)
nDodge = nDodge + Fix((nCharm - 50) / 5)
nDodge = nDodge + Fix((nAgility - 50) / 3)
nDodge = nDodge + nPlusDodge '[cumulative dodge from: abilities + auras + race + class + items]

If nMaxEncum > 0 Then
    nEncumPct = Fix((nCurrentEncum / nMaxEncum) * 100)
    If nEncumPct < 33 Then
        nDodge = nDodge + 10 - Fix(nEncumPct / 10)
    End If
End If

If nDodge < 0 Then nDodge = 0
If nDodge > 95 Then nDodge = 95

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
Public Function GetSpellAttackType(ByVal nAttackType As Integer) As String

On Error GoTo error:

Select Case nAttackType
    Case 0: GetSpellAttackType = "Cold"
    Case 1: GetSpellAttackType = "Hot"
    Case 2: GetSpellAttackType = "Stone"
    Case 3: GetSpellAttackType = "Lightning"
    Case 4: GetSpellAttackType = "Normal"
    Case 5: GetSpellAttackType = "Water"
    Case 6: GetSpellAttackType = "Poison"
    Case Else: GetSpellAttackType = nAttackType
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

Public Function GetAbilityList() As Variant()
On Error GoTo error:
Dim sArr() As Variant, x As Integer

ReDim sArr(200)
For x = 1 To 200
    sArr(x) = GetAbilityName(x, True)
    If sArr(x) = "" Or sArr(x) = "Ability " & x Then
        sArr(x) = "[Ability " & x & "]"
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
    Case 185: GetAbilityName = "BadAttk"
    Case 186: GetAbilityName = "PerStealth"
    Case 187: GetAbilityName = "Meditate"
    Case Else: GetAbilityName = "Ability " & nNum
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

CalcBSDamage = (nLevel * 2) + Fix(nStealth / 10) + (nDMG * 2) + nBsDmgMod '+ 20
If Not bClassStealth Then CalcBSDamage = Fix((CalcBSDamage * 75) / 100)
CalcBSDamage = Fix(((nLevel + 100) * CalcBSDamage) / 100)

out:
Exit Function
error:
Call HandleError("CalcBSDamage")
Resume out:
End Function


Public Function CalcManaRegen(ByVal nLevel As Long, ByVal nInt As Long, ByVal nWil As Long, _
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
        CalcManaRegen = nInt
    Case 2: 'priest
        CalcManaRegen = nWil
    Case 3: 'druid
        CalcManaRegen = Fix((nInt + nWil) / 2)
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

Public Function CalcMR(ByVal nInt As Integer, ByVal nWis As Integer, Optional ByVal nModifiers As Integer) As Long
On Error GoTo error:

CalcMR = Fix((nInt + (nWis * 3)) / 4) + nModifiers

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

Public Function CalcSpellCasting(ByVal nLevel As Long, ByVal nInt As Long, ByVal nWil As Long, _
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
        CalcSpellCasting = Fix(((nInt * 3) + nWil) / 6) + (nLevel * 2) + (nMagicLVL * 5)
    Case 2: 'priest
        CalcSpellCasting = Fix(((nWil * 3) + nInt) / 6) + (nLevel * 2) + (nMagicLVL * 5)
    Case 3: 'druid
        CalcSpellCasting = Fix((nWil + nInt) / 3) + (nLevel * 2) + (nMagicLVL * 5)
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

Public Function CalcPicklocks(ByVal nLevel As Long, ByVal nAGL As Long, ByVal nInt As Long) As Long
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

CalcPicklocks = Fix((((CalcPicklocks * 5) + (nAGL + nInt)) * 2) / 7)
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
nCombat = GetClassCombat(frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex))
nEncum = CalcEncumbrancePercent(Val(frmMain.lblInvenCharStat(0).Caption), Val(frmMain.lblInvenCharStat(1).Caption))
nEnergy = CalcEnergyUsedWithEncum(nCombat, Val(frmMain.txtGlobalLevel(0).Text), tabItems.Fields("Speed"), _
    Val(frmMain.txtCharStats(3).Text), Val(frmMain.txtCharStats(0).Text), nEncum, tabItems.Fields("StrReq"))

If Not Val(frmMain.txtCharStats(0).Text) < tabItems.Fields("StrReq") Then
    GetQuickAndDeadlyBonus = CalcQuickAndDeadlyBonus(Val(frmMain.txtCharStats(3).Text), nEnergy, nEncum)
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetQuickAndDeadlyBonus")
Resume out:
End Function

Public Function CalcQuickAndDeadlyBonus(ByVal nAGL As Currency, ByVal nEU As Currency, _
    ByVal nEncum As Currency) As Currency
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

If (nEU >= 200) Or (nEncum > 66) Then Exit Function

CalcQuickAndDeadlyBonus = (200 - nEU) + Fix((nAGL - 50) / 10)

If (CalcQuickAndDeadlyBonus > 20) Then CalcQuickAndDeadlyBonus = 20
If (nEncum >= 33) Then CalcQuickAndDeadlyBonus = Fix(CalcQuickAndDeadlyBonus / 2)

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

Public Function CalcEnergyUsed(ByVal nCombat As Currency, ByVal nLevel As Currency, _
    ByVal nSpeed As Currency, ByVal nAGL As Currency, Optional ByVal nSTR As Currency = 0, _
    Optional ByVal nItemSTR As Currency = 0) As Currency
'{ Calculates the energy used for a given Combat rating, Level, Speed, AGL, STR,
'  and ItemSTR }
'function  CalcEnergyUsed(Combat, Level, Speed, AGL: integer; STR: integer = 0; ItemSTR: integer = 0): longword; begin
'  Result := longword(Speed * 1000) div (longword((((Level * (Combat + 2)) + 45) * (AGL + 150)) * 1500) div 9000);
'  If (STR < ItemSTR) Then
'    Result := longword(longword((longword(ItemSTR - STR) * 3) + 200) *
'Result) div 200; end;

CalcEnergyUsed = Fix((nSpeed * 1000) / Fix(((((nLevel * (nCombat + 2)) + 45) * (nAGL + 150)) * 1500) / 9000))

If (nSTR < nItemSTR) Then
    CalcEnergyUsed = Fix(((((nItemSTR - nSTR) * 3) + 200) * CalcEnergyUsed) / 200)
End If

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
    
CalcEnergyUsedWithEncum = CalcEnergyUsed(nCombat, nLevel, nSpeed, nAGL, nSTR, nItemSTR)
CalcEnergyUsedWithEncum = Fix((CalcEnergyUsedWithEncum * Fix(Fix(nEncum / 2) + 75)) / 100)

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

