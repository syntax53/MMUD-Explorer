Attribute VB_Name = "modMain"
Option Explicit
Option Base 0

Global bHideRecordNumbers As Boolean
Global bOnlyInGame As Boolean
Global nNMRVer As Double
Global sCurrentDatabaseFile As String
'Global bOnlyLearnable As Boolean

Public Enum QBColorCode
    Black = 0
    Blue = 1
    Green = 2
    Cyan = 3
    Red = 4
    Magenta = 5
    Yellow = 6
    White = 7
    Grey = 8
    BrightBlue = 9
    BrightGreen = 10
    BrightCyan = 11
    BrightRed = 12
    BrightMagenta = 13
    BrightYellow = 14
    BrightWhite = 15
End Enum

Public Enum eExpandBy
    Percent50 = 0
    Percent75 = 1
    DoubleWidth = 2
    TripleWidth = 3
    QuadWidth = 4
    NoExpand = 5
End Enum
Public Enum eExpandType
    WidthOnly = 0
    HeightOnly = 1
    HeightAndWidth = 2
End Enum

Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_GETITEMHEIGHT = &H154
Private Const CB_SETDROPPEDWIDTH = &H160
Private Const CB_GETDROPPEDWIDTH = &H15F
Private Const CB_SETDROPPEDCONTROLRECT = &H160
Private Const DT_CALCRECT = &H400

Public bPromptSave As Boolean
Public bCancelTerminate As Boolean
Public bAppTerminating As Boolean
Public sRecentFiles(1 To 5, 1 To 2) As String '1=shown, 2=filename
Public sRecentDBs(1 To 5, 1 To 2) As String '1=shown, 2=filename
Public nEquippedItem(0 To 18) As Long
Public bLegit As Boolean

Public Type POINTAPI
   x As Long
   y As Long
End Type

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Public Declare Function MoveWindow Lib "user32" _
  (ByVal hWnd As Long, _
   ByVal x As Long, ByVal y As Long, _
   ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal bRepaint As Long) As Long

Public Declare Function GetWindowRect Lib "user32" _
  (ByVal hWnd As Long, _
   lpRect As RECT) As Long

Public Declare Function ScreenToClient Lib "user32" _
  (ByVal hWnd As Long, _
   lpPoint As POINTAPI) As Long

Private Declare Function SendMessageLong Lib "user32" Alias _
        "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function DrawText Lib "user32" Alias _
    "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, _
    ByVal nCount As Long, lpRect As RECT, ByVal wFormat _
    As Long) As Long
    
Public Declare Function CalcExpNeeded Lib "lltmmudxp" (ByVal Level As Long, ByVal Chart As Long) As Currency

Public Sub ExpandCombo(ByRef Combo As ComboBox, ByVal ExpandType As eExpandType, _
    ByVal ExpandBy As eExpandBy, Optional ByVal hFrame As Long)

    Dim lRet As Long
    Dim pt As POINTAPI
    Dim rc As RECT
    Dim lComboWidth As Long
    Dim lNewHeight As Long
    Dim lItemHeight As Long
    
    If ExpandType <> HeightOnly Then
        lComboWidth = (Combo.Width / Screen.TwipsPerPixelX)
        Select Case ExpandBy
            Case 0:
                lComboWidth = lComboWidth + (lComboWidth * 0.5)
            Case 1:
                lComboWidth = lComboWidth + (lComboWidth * 0.75)
            Case 2:
                lComboWidth = lComboWidth * 2
            Case 3:
                lComboWidth = lComboWidth * 3
            Case 4:
                lComboWidth = lComboWidth * 4
        End Select
        lRet = SendMessage(Combo.hWnd, CB_SETDROPPEDCONTROLRECT, lComboWidth, 0)
        
    End If
    
    If ExpandType <> WidthOnly Then
        lComboWidth = Combo.Width / Screen.TwipsPerPixelX
        lItemHeight = SendMessage(Combo.hWnd, CB_GETITEMHEIGHT, 0, 0)
        Select Case ExpandBy
            Case 1:
                'lComboWidth = lComboWidth + (lComboWidth * 0.75)
                lNewHeight = lItemHeight * 16
            Case 2:
                'lComboWidth = lComboWidth * 2
                lNewHeight = lItemHeight * 18
            Case 3:
                'lComboWidth = lComboWidth * 3
                lNewHeight = lItemHeight * 26
            Case 4:
                'lComboWidth = lComboWidth * 4
                lNewHeight = lItemHeight * 32
            Case Else:
                lNewHeight = lItemHeight * 14
                'lComboWidth = lComboWidth + (lComboWidth * 0.5)
        End Select
        Call GetWindowRect(Combo.hWnd, rc)
        pt.x = rc.Left
        pt.y = rc.Top
        Call ScreenToClient(hFrame, pt)
        Call MoveWindow(Combo.hWnd, pt.x, pt.y, lComboWidth, lNewHeight, True)
    End If
    
End Sub

Public Function AutoSizeDropDownWidth(Combo As Object) As Boolean
'**************************************************************
'PURPOSE: Automatically size the combo box drop down width
'         based on the width of the longest item in the combo box

'PARAMETERS: Combo - ComboBox to size

'RETURNS: True if successful, false otherwise

'ASSUMPTIONS: 1. Form's Scale Mode is vbTwips, which is why
'                conversion from twips to pixels are made.
'                API functions require units in pixels
'
'             2. Combo Box's parent is a form or other
'                container that support the hDC property

'EXAMPLE: AutoSizeDropDownWidth Combo1
'****************************************************************
Dim lRet As Long
Dim lCurrentWidth As Single
Dim rectCboText As RECT
Dim lParentHDC As Long
Dim lListCount As Long
Dim lCtr As Long
Dim lTempWidth As Long
Dim lWidth As Long
Dim sSavedFont As String
Dim sngSavedSize As Single
Dim bSavedBold As Boolean
Dim bSavedItalic As Boolean
Dim bSavedUnderline As Boolean
Dim bFontSaved As Boolean

On Error GoTo ErrorHandler

If Not TypeOf Combo Is ComboBox Then Exit Function
lParentHDC = Combo.Parent.hdc
If lParentHDC = 0 Then Exit Function
lListCount = Combo.ListCount
If lListCount = 0 Then Exit Function


'Change font of parent to combo box's font
'Save first so it can be reverted when finished
'this is necessary for drawtext API Function
'which is used to determine longest string in combo box
With Combo.Parent

    sSavedFont = .FontName
    sngSavedSize = .FontSize
    bSavedBold = .FontBold
    bSavedItalic = .FontItalic
    bSavedUnderline = .FontUnderline
    
    .FontName = Combo.FontName
    .FontSize = Combo.FontSize
    .FontBold = Combo.FontBold
    .FontItalic = Combo.FontItalic
    .FontUnderline = Combo.FontItalic

End With

bFontSaved = True

'Get the width of the largest item
For lCtr = 0 To lListCount
   DrawText lParentHDC, Combo.List(lCtr), -1, rectCboText, _
        DT_CALCRECT
   'adjust the number added (20 in this case to
   'achieve desired right margin
   lTempWidth = rectCboText.Right - rectCboText.Left + 20

   If (lTempWidth > lWidth) Then
      lWidth = lTempWidth
   End If
Next
 
lCurrentWidth = SendMessageLong(Combo.hWnd, CB_GETDROPPEDWIDTH, _
    0, 0)

If lCurrentWidth > lWidth Then 'current drop-down width is
'                               sufficient

    AutoSizeDropDownWidth = True
    GoTo ErrorHandler
    Exit Function
End If
 
'don't allow drop-down width to
'exceed screen.width
 
   If lWidth > Screen.Width \ Screen.TwipsPerPixelX - 20 Then _
    lWidth = Screen.Width \ Screen.TwipsPerPixelX - 20

lRet = SendMessageLong(Combo.hWnd, CB_SETDROPPEDWIDTH, lWidth, 0)

AutoSizeDropDownWidth = lRet > 0
ErrorHandler:
On Error Resume Next
If bFontSaved Then
'restore parent's font settings
  With Combo.Parent
    .FontName = sSavedFont
    .FontSize = sngSavedSize
    .FontUnderline = bSavedUnderline
    .FontBold = bSavedBold
    .FontItalic = bSavedItalic
 End With
End If
End Function

Public Function PullItemDetail(DetailTB As TextBox, LocationLV As ListView)
Dim sStr As String, sAbil As String, x As Integer, sCasts As String, nPercent As Integer
Dim sNegate As String, sClasses As String, sRaces As String, sClassOk As String, sCost As String
Dim sUses As String, sGetDrop As String, oLI As ListItem, nNumber As Long

'sStr = ClipNull(tabItems.Fields("Name")) & " (" & tabItems.Fields("Number") & ")"

On Error GoTo Error:

nNumber = tabItems.Fields("Number")

If tabItems.Fields("UseCount") > 0 Then
    sUses = tabItems.Fields("UseCount")
    If tabItems.Fields("Retain After Uses") = 1 Then
        sUses = sUses & "x/day"
    Else
        sUses = sUses & " (destroys after uses)"
    End If
End If

If tabItems.Fields("Gettable") = 0 Then
    sGetDrop = "Not Getable"
End If

If tabItems.Fields("Not Droppable") = 1 Then
    If Not sGetDrop = "" Then sGetDrop = sGetDrop & ", "
    sGetDrop = sGetDrop & "Not Droppable"
End If

If tabItems.Fields("Destroy On Death") = 1 Then
    If Not sGetDrop = "" Then sGetDrop = sGetDrop & ", "
    sGetDrop = sGetDrop & "Destroys On Death"
End If

For x = 0 To 9
    If tabItems.Fields("ClassRest-" & x) <> 0 Then
        If sClasses <> "" Then sClasses = sClasses & ", "
        sClasses = sClasses & GetClassName(tabItems.Fields("ClassRest-" & x))
    End If
Next

For x = 0 To 9
    If tabItems.Fields("RaceRest-" & x) <> 0 Then
        If sRaces <> "" Then sRaces = sRaces & ", "
        sRaces = sRaces & GetRaceName(tabItems.Fields("RaceRest-" & x))
    End If
Next

For x = 0 To 9
    If tabItems.Fields("NegateSpell-" & x) <> 0 Then
        If sNegate <> "" Then sNegate = sNegate & ", "
        sNegate = sNegate & GetSpellName(tabItems.Fields("NegateSpell-" & x), bHideRecordNumbers)
    End If
Next

Call GetLocations(tabItems.Fields("Obtained From"), LocationLV)
If nNMRVer >= 1.7 Then Call GetLocations(tabItems.Fields("References"), LocationLV, True)

If Not tabItems.Fields("Number") = nNumber Then
    tabItems.Index = "pkItems"
    tabItems.Seek "=", nNumber
End If
        
For x = 0 To 19
    If tabItems.Fields("Abil-" & x) <> 0 Then
        Select Case tabItems.Fields("Abil-" & x)
            Case 116: '116-bsacc
                If Not DetailTB.name = "txtWeaponCompareDetail" And _
                    Not DetailTB.name = "txtWeaponDetail" Then
                    
                    If sAbil <> "" Then sAbil = sAbil & ", "
                    sAbil = sAbil & GetAbilityStats(tabItems.Fields("Abil-" & x), tabItems.Fields("AbilVal-" & x), frmMain.lvOtherItemLoc)
                    If Right(sAbil, 2) = ", " Then sAbil = Left(sAbil, Len(sAbil) - 2)
                End If
            Case 22, 105, 106, 135:  '22-acc, 105-acc, 106-acc, 135-minlvl
                If Not DetailTB.name = "txtWeaponCompareDetail" And _
                    Not DetailTB.name = "txtWeaponDetail" And _
                    Not DetailTB.name = "txtArmourCompareDetail" And _
                    Not DetailTB.name = "txtArmourDetail" Then
                    
                    If sAbil <> "" Then sAbil = sAbil & ", "
                    sAbil = sAbil & GetAbilityStats(tabItems.Fields("Abil-" & x), tabItems.Fields("AbilVal-" & x), frmMain.lvOtherItemLoc)
                    If Right(sAbil, 2) = ", " Then sAbil = Left(sAbil, Len(sAbil) - 2)
                End If
            Case 59: 'class ok
                If sClassOk <> "" Then sClassOk = sClassOk & ", "
                sClassOk = sClassOk & GetClassName(tabItems.Fields("AbilVal-" & x))
            
            Case 43: 'casts spell
                If sCasts <> "" Then sCasts = sCasts & ", "
                'nSpellNest = 0 'make sure this doesn't nest too deep
                sCasts = sCasts & "[" & GetSpellName(tabItems.Fields("AbilVal-" & x), bHideRecordNumbers) _
                    & ", " & PullSpellEQ(True, 0, tabItems.Fields("AbilVal-" & x), frmMain.lvOtherItemLoc)
                If Not nPercent = 0 Then
                    sCasts = sCasts & ", " & nPercent & "%]"
                Else
                    sCasts = sCasts & "]"
                End If
                
                Set oLI = LocationLV.ListItems.Add
                oLI.Text = "Casts: " & GetSpellName(tabItems.Fields("AbilVal-" & x), bHideRecordNumbers)
                oLI.Tag = tabItems.Fields("AbilVal-" & x)
            
            Case 114: '%spell
                nPercent = tabItems.Fields("AbilVal-" & x)
            
            Case Else:
                If sAbil <> "" Then sAbil = sAbil & ", "
                sAbil = sAbil & GetAbilityStats(tabItems.Fields("Abil-" & x), tabItems.Fields("AbilVal-" & x), frmMain.lvOtherItemLoc)
                If Right(sAbil, 2) = ", " Then sAbil = Left(sAbil, Len(sAbil) - 2)
                
        End Select
        
    End If
Next


If Not sAbil = "" Then
    If Not sStr = "" Then sStr = sStr & " -- "
    sStr = sStr & "Abilities: " & sAbil
End If
If Not sUses = "" Then
    If Not sStr = "" Then sStr = sStr & " -- "
    sStr = sStr & "Uses: " & sUses
End If
If Not sCasts = "" Then
    If Not sStr = "" Then sStr = sStr & " -- "
    sStr = sStr & "Casts: " & sCasts
End If
If Not sNegate = "" Then
    If Not sStr = "" Then sStr = sStr & " -- "
    sStr = sStr & "Negates: " & sNegate
End If
If Not sClassOk = "" Then
    If Not sStr = "" Then sStr = sStr & " -- "
    sStr = sStr & "ClassOK: " & sClassOk
End If
If Not sClasses = "" Then
    If Not sStr = "" Then sStr = sStr & " -- "
    sStr = sStr & "Classes: " & sClasses
End If
If Not sRaces = "" Then
    If Not sStr = "" Then sStr = sStr & " -- "
    sStr = sStr & "Races: " & sRaces
End If
If Not sGetDrop = "" Then
    If Not sStr = "" Then sStr = sStr & " -- "
    sStr = sStr & sGetDrop
End If

DetailTB.Text = sStr
Set oLI = Nothing
Exit Function

Error:
Call HandleError
Set oLI = Nothing
End Function

Public Function PullClassDetail(nClassNum As Long, DetailTB As TextBox)
Dim sAbil As String, x As Integer

On Error GoTo Error:

tabClasses.Index = "pkClasses"
tabClasses.Seek "=", nClassNum
If tabClasses.NoMatch = True Then
    DetailTB.Text = "Class not found"
    Exit Function
End If

'sStr = ClipNull(tabClasses.Fields("Name")) & " (" & tabClasses.Fields("Number") & ")"

'sStr = "Exp: " & (Val(tabClasses.Fields("ExpTable")) + 100) & "%"
'sStr = sStr & " -- Weapon: " & GetClassWeaponType(tabClasses.Fields("WeaponType"))
'sStr = sStr & " -- Armour: " & GetArmourType(tabClasses.Fields("ArmourType"))
'sStr = sStr & " -- Magic: " & GetMagery(tabClasses.Fields("MageryType"), tabClasses.Fields("MageryLVL"))
'sStr = sStr & " -- Combat: " & (tabClasses.Fields("CombatLVL") - 2)
'sStr = sStr & " -- HP: " & tabClasses.Fields("MinHits") & "-" & (tabClasses.Fields("MinHits") + tabClasses.Fields("MaxHits"))

For x = 0 To 9
    If Not tabClasses.Fields("Abil-" & x) = 0 Then
        Select Case tabClasses.Fields("Abil-" & x)
            Case 0:
            Case 59: 'classok (dont want it changing the class on us)
                
            Case Else:
                If sAbil <> "" Then sAbil = sAbil & ", "
                sAbil = sAbil & GetAbilityStats(tabClasses.Fields("Abil-" & x), tabClasses.Fields("AbilVal-" & x))
                If Right(sAbil, 2) = ", " Then sAbil = Left(sAbil, Len(sAbil) - 2)
                
        End Select
    End If
Next

DetailTB.Text = sAbil

Exit Function

Error:
Call HandleError

End Function
Public Function PullRaceDetail(nRaceNum As Long, DetailTB As TextBox)
Dim sAbil As String, x As Integer

On Error GoTo Error:

tabRaces.Index = "pkRaces"
tabRaces.Seek "=", nRaceNum
If tabRaces.NoMatch = True Then
    DetailTB.Text = "Race not found"
    Exit Function
End If

'sStr = ClipNull(tabRaces.Fields("Name")) & " (" & tabRaces.Fields("Number") & ")"
'sStr = "Exp: " & tabRaces.Fields("ExpTable") & "%, HP Bonus: " & tabRaces.Fields("HPPerLVL")
'
'sStr = sStr & vbCrLf _
'    & "Str: " & tabRaces.Fields("mSTR") & "-" & tabRaces.Fields("xSTR") & " ... " & "Agi: " & tabRaces.Fields("mAGL") & "-" & tabRaces.Fields("xAGL") & vbCrLf _
'    & "Int: " & tabRaces.Fields("mINT") & "-" & tabRaces.Fields("xINT") & " ... " & "Hea: " & tabRaces.Fields("mHEA") & "-" & tabRaces.Fields("xHEA") & vbCrLf _
'    & "Wis: " & tabRaces.Fields("mWIL") & "-" & tabRaces.Fields("xWIL") & " ... " & "Cha: " & tabRaces.Fields("mCHM") & "-" & tabRaces.Fields("xCHM")

For x = 0 To 9
    If Not tabRaces.Fields("Abil-" & x) = 0 Then
        Select Case tabRaces.Fields("Abil-" & x)
            Case 0:
            Case Else:
                If sAbil <> "" Then sAbil = sAbil & ", "
                sAbil = sAbil & GetAbilityStats(tabRaces.Fields("Abil-" & x), tabRaces.Fields("AbilVal-" & x))
                If Right(sAbil, 2) = ", " Then sAbil = Left(sAbil, Len(sAbil) - 2)
        End Select
    End If
Next

DetailTB.Text = sAbil

Exit Function

Error:
Call HandleError

End Function
Public Function PullMonsterDetail(nMonsterNum As Long, DetailLV As ListView) ', DetailTB As TextBox)
Dim sAbil As String, x As Integer, y As Integer
Dim sCash As String, nPercent As Integer, sMonGuards As String
Dim oLI As ListItem, nExp As Currency, nMonsterDamage() As Long, nMonsterEnergy As Long

On Error GoTo Error:

DetailLV.ListItems.clear

tabMonsters.Index = "pkMonsters"
tabMonsters.Seek "=", nMonsterNum
If tabMonsters.NoMatch = True Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Monster not found"
    'DetailTB.Text = "Monster not found"
    Set oLI = Nothing
    Exit Function
End If

Set oLI = DetailLV.ListItems.Add()
oLI.Text = "Name"
oLI.ListSubItems.Add (1), "Detail", tabMonsters.Fields("Name") & " (" & tabMonsters.Fields("Number") & ")"

Set oLI = DetailLV.ListItems.Add()
oLI.Text = "Experience"
If UseExpMulti Then
    nExp = tabMonsters.Fields("EXP") * tabMonsters.Fields("ExpMulti")
Else
    nExp = tabMonsters.Fields("EXP")
End If
oLI.ListSubItems.Add (1), "Detail", IIf(nExp > 0, Format(nExp, "#,#"), 0)

If Not tabMonsters.Fields("RegenTime") = 0 Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Regen Time"
    oLI.ListSubItems.Add (1), "Detail", tabMonsters.Fields("RegenTime") & " hour(s)"
End If

Set oLI = DetailLV.ListItems.Add()
oLI.Text = "Type"
oLI.ListSubItems.Add (1), "Detail", GetMonType(tabMonsters.Fields("Type"))

Set oLI = DetailLV.ListItems.Add()
oLI.Text = "Alignment"
oLI.ListSubItems.Add (1), "Detail", GetMonAlignment(tabMonsters.Fields("Align"))

If tabMonsters.Fields("Undead") = 1 Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Undead"
    oLI.ListSubItems.Add (1), "Detail", "Yes"
End If

Set oLI = DetailLV.ListItems.Add()
oLI.Text = "HPs"
oLI.ForeColor = RGB(204, 0, 0)
oLI.ListSubItems.Add (1), "Detail", tabMonsters.Fields("HP") & " (Regens: " & tabMonsters.Fields("HPRegen") & " HPs/click)"
oLI.ListSubItems(1).ForeColor = RGB(204, 0, 0)

Set oLI = DetailLV.ListItems.Add()
oLI.Text = "AC/DR"
oLI.ListSubItems.Add (1), "Detail", tabMonsters.Fields("ArmourClass") & "/" & tabMonsters.Fields("DamageResist")

Set oLI = DetailLV.ListItems.Add()
oLI.Text = "MR"
oLI.ListSubItems.Add (1), "Detail", tabMonsters.Fields("MagicRes")

Set oLI = DetailLV.ListItems.Add()
oLI.Text = "Follow %"
oLI.ListSubItems.Add (1), "Detail", tabMonsters.Fields("Follow%")

Set oLI = DetailLV.ListItems.Add()
oLI.Text = "Charm LVL"
oLI.ListSubItems.Add (1), "Detail", tabMonsters.Fields("CharmLVL")

'cash
If Not tabMonsters.Fields("R") = 0 Then
    sCash = sCash & tabMonsters.Fields("R") & " Runic"
End If
If Not tabMonsters.Fields("P") = 0 Then
    If Not sCash = "" Then sCash = sCash & ", "
    sCash = sCash & tabMonsters.Fields("P") & " Plat"
End If
If Not tabMonsters.Fields("G") = 0 Then
    If Not sCash = "" Then sCash = sCash & ", "
    sCash = sCash & tabMonsters.Fields("G") & " Gold"
End If
If Not tabMonsters.Fields("S") = 0 Then
    If Not sCash = "" Then sCash = sCash & ", "
    sCash = sCash & tabMonsters.Fields("S") & " Silver"
End If
If Not tabMonsters.Fields("C") = 0 Then
    If Not sCash = "" Then sCash = sCash & ", "
    sCash = sCash & tabMonsters.Fields("C") & " Copper"
End If
If Not sCash = "" Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Cash (up to)"
    oLI.ListSubItems.Add (1), "Detail", sCash
End If

If Not tabMonsters.Fields("Weapon") = 0 Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Weapon"
    oLI.Tag = "Item"
    oLI.ListSubItems.Add (1), "Detail", GetItemName(tabMonsters.Fields("Weapon"), bHideRecordNumbers)
    oLI.ListSubItems(1).Tag = tabMonsters.Fields("Weapon")
End If

If tabMonsters.Fields("CreateSpell") > 0 Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Create Spell:"
    oLI.Tag = "Spell"
    'oLI.ForeColor = &HC00000
    'nSpellNest = 0 'make sure this doesn't nest too deep
    oLI.ListSubItems.Add (1), "Detail", "[" & GetSpellName(tabMonsters.Fields("CreateSpell"), bHideRecordNumbers) _
        & ", " & PullSpellEQ(False, , tabMonsters.Fields("CreateSpell")) & "]"
    If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
    oLI.ListSubItems(1).Tag = tabMonsters.Fields("CreateSpell")
    'oLI.ListSubItems(1).ForeColor = &HC00000
End If

If tabMonsters.Fields("DeathSpell") > 0 Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Death Spell"
    oLI.Tag = "Spell"
    'oLI.ForeColor = &HC00000
    'nSpellNest = 0 'make sure this doesn't nest too deep
    oLI.ListSubItems.Add (1), "Detail", "[" & GetSpellName(tabMonsters.Fields("DeathSpell"), bHideRecordNumbers) _
        & ", " & PullSpellEQ(False, , tabMonsters.Fields("DeathSpell")) & "]"
    If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
    oLI.ListSubItems(1).Tag = tabMonsters.Fields("DeathSpell")
    'oLI.ListSubItems(1).ForeColor = &HC00000
End If

If tabMonsters.Fields("GreetTXT") > 0 Then
    tabTBInfo.Index = "pkTBInfo"
    tabTBInfo.Seek "=", tabMonsters.Fields("GreetTXT")
    If Not tabTBInfo.NoMatch Then
        If Not LCase(Left(tabTBInfo.Fields("Action"), 4)) = "heh:" _
            And Not LCase(Left(tabTBInfo.Fields("Action"), 8)) = "nothing:" _
            And Not LCase(Left(tabTBInfo.Fields("Action"), 5)) = "yada:" _
            And Not LCase(Left(tabTBInfo.Fields("Action"), 5)) = "hehe:" _
            And Not LCase(Left(tabTBInfo.Fields("Action"), 5)) = "shit:" Then
            
            Set oLI = DetailLV.ListItems.Add()
            oLI.Tag = "greet_text"
            oLI.Text = "Greet Commands:"
            oLI.ForeColor = &H8000&
            oLI.ListSubItems.Add (1), "Detail", "Textblock " & tabMonsters.Fields("GreetTXT")
            oLI.ListSubItems(1).Tag = tabMonsters.Fields("GreetTXT")
            oLI.ListSubItems(1).ForeColor = &H8000&
        End If
    End If
End If

For x = 0 To 9 'abilities
    If tabMonsters.Fields("Abil-" & x) > 0 And Not tabMonsters.Fields("Abil-" & x) = 146 Then
        If sAbil <> "" Then sAbil = sAbil & ", "
        sAbil = sAbil & GetAbilityStats(tabMonsters.Fields("Abil-" & x), tabMonsters.Fields("AbilVal-" & x))
        If Right(sAbil, 2) = ", " Then sAbil = Left(sAbil, Len(sAbil) - 2)
    End If
Next x
If Not sAbil = "" Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Abilities: "
    oLI.ForeColor = &HC00000
    'oLI.ForeColor = &HFF00FF
    'oLI.Bold = True
    oLI.ListSubItems.Add (1), "Detail", sAbil
    oLI.ListSubItems(1).ForeColor = &HC00000
    'oLI.ListSubItems(1).ForeColor = &HFF00FF
    'oLI.Height = oLI.Height * 2
End If

Set oLI = Nothing
For x = 0 To 9 'mon guards
    If tabMonsters.Fields("Abil-" & x) = 146 And tabMonsters.Fields("AbilVal-" & x) > 0 Then
        If oLI Is Nothing Then
            Set oLI = DetailLV.ListItems.Add()
            oLI.Text = ""
            
            Set oLI = DetailLV.ListItems.Add()
            oLI.Text = "Guarded by: "
            oLI.ForeColor = &H800080
            oLI.Tag = "Monster"
        Else
            Set oLI = DetailLV.ListItems.Add()
            oLI.Text = ""
            oLI.Tag = "Monster"
        End If
        
        oLI.ListSubItems.Add (1), "Detail", GetMonsterName(tabMonsters.Fields("AbilVal-" & x), bHideRecordNumbers)
        oLI.ListSubItems(1).ForeColor = &H800080
        
        tabMonsters.Seek "=", nMonsterNum
        oLI.ListSubItems(1).Tag = tabMonsters.Fields("AbilVal-" & x)
    End If
Next

Set oLI = DetailLV.ListItems.Add()
oLI.Text = ""

y = 0
For x = 0 To 9 'item drops
    If Not tabMonsters.Fields("DropItem-" & x) = 0 Then
        y = y + 1
        Set oLI = DetailLV.ListItems.Add()
        If y = 1 Then
            oLI.Text = "Item Drops"
        Else
            oLI.Text = ""
        End If
        oLI.Tag = "Item"
        
        oLI.ListSubItems.Add (1), "Detail", y & ". " & GetItemName(tabMonsters.Fields("DropItem-" & x), bHideRecordNumbers) _
            & " (" & tabMonsters.Fields("DropItem%-" & x) & "%)"
        oLI.ListSubItems(1).Tag = tabMonsters.Fields("DropItem-" & x)
    End If
Next
If y > 0 Then 'add blank line if there were entried added
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
End If

nMonsterDamage() = GetMonsterDamagePerRound(tabMonsters.Fields("Number"))
If nMonsterDamage(1) > 0 Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Dmg/Round **"
    If nMonsterDamage(0) < nMonsterDamage(1) Then
        oLI.ListSubItems.Add (1), "Detail", "AVG: " & nMonsterDamage(0) & ", Max: " & nMonsterDamage(1)
    Else
        oLI.ListSubItems.Add (1), "Detail", "AVG: " & nMonsterDamage(0)
    End If
    oLI.ListSubItems(1).Text = oLI.ListSubItems(1).Text & " (-DR/-MR)  ** Rough average. Use simulator for accurate dmg."
    
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
End If

If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum

nPercent = 0
y = 0
For x = 0 To 4 'between round spells
    If Not tabMonsters.Fields("MidSpell-" & x) = 0 Then
        
        y = y + 1
        Set oLI = DetailLV.ListItems.Add()
        If y = 1 Then
            oLI.Text = "Between Rounds"
        Else
            oLI.Text = ""
        End If
        oLI.Tag = "Spell"
        
        nPercent = tabMonsters.Fields("MidSpell%-" & x) - nPercent
        'nSpellNest = 0 'make sure this doesn't nest too deep
        oLI.ListSubItems.Add (1), "Detail", "(" & nPercent & "%) [" & _
            GetSpellName(tabMonsters.Fields("MidSpell-" & x), bHideRecordNumbers) _
            & ", " & PullSpellEQ(True, tabMonsters.Fields("MidSpellLVL-" & x), _
            tabMonsters.Fields("MidSpell-" & x), , , True) & "]"
        If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
        oLI.ListSubItems(1).Tag = tabMonsters.Fields("MidSpell-" & x)
        nPercent = tabMonsters.Fields("MidSpell%-" & x)
    End If
Next
If y > 0 Then 'add blank line if there was entried added
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
End If

nMonsterEnergy = 1000
If nNMRVer >= 1.71 Then
    If Not tabMonsters.Fields("Energy") = 0 Then
        nMonsterEnergy = tabMonsters.Fields("Energy")
        
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "Energy/Round"
        oLI.ListSubItems.Add (1), "Detail", nMonsterEnergy
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = ""
    End If
End If

nPercent = 0
y = 0
For x = 0 To 4 'attacks
    If tabMonsters.Fields("AttType-" & x) > 0 And tabMonsters.Fields("AttType-" & x) <= 3 Then
        y = y + 1
        Set oLI = DetailLV.ListItems.Add()
        
        nPercent = tabMonsters.Fields("Att%-" & x) - nPercent
        
        oLI.Text = "Attack " & y & " (" & nPercent & "%)"
        oLI.ListSubItems.Add (1), "Detail", GetMonAttackType(tabMonsters.Fields("AttType-" & x))
        
        nPercent = tabMonsters.Fields("Att%-" & x)
        
        Select Case tabMonsters.Fields("AttType-" & x)
            Case 1, 3: 'normal, rob
                
                Set oLI = DetailLV.ListItems.Add()
                oLI.Text = ""
                oLI.ListSubItems.Add (1), "Detail", "Min-Max: " & tabMonsters.Fields("AttMin-" & x) & "-" & tabMonsters.Fields("AttMax-" & x)
                
                Set oLI = DetailLV.ListItems.Add()
                oLI.Text = ""
                oLI.ListSubItems.Add (1), "Detail", "Accuracy: " & tabMonsters.Fields("AttAcc-" & x)
                
                If tabMonsters.Fields("AttEnergy-" & x) > 0 Then
                    Set oLI = DetailLV.ListItems.Add()
                    oLI.Text = ""
                    oLI.ListSubItems.Add (1), "Detail", "Energy: " & tabMonsters.Fields("AttEnergy-" & x) _
                        & " (Max " & Fix(nMonsterEnergy / tabMonsters.Fields("AttEnergy-" & x)) & "x/round)"
                Else
                    Set oLI = DetailLV.ListItems.Add()
                    oLI.Text = ""
                    oLI.ListSubItems.Add (1), "Detail", "Energy: " & tabMonsters.Fields("AttEnergy-" & x)
                End If
                
                If Not tabMonsters.Fields("AttHitSpell-" & x) = 0 Then
                    Set oLI = DetailLV.ListItems.Add()
                    oLI.Text = ""
                    oLI.Tag = "Spell"
                    'nSpellNest = 0 'make sure this doesn't nest too deep
                    oLI.ListSubItems.Add (1), "Detail", "Hit Spell: [" & _
                        GetSpellName(tabMonsters.Fields("AttHitSpell-" & x), bHideRecordNumbers) _
                        & ", " & PullSpellEQ(False, , tabMonsters.Fields("AttHitSpell-" & x)) & "]"
                    If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
                    oLI.ListSubItems(1).Tag = tabMonsters.Fields("AttHitSpell-" & x)
                End If
                
            Case 2: 'spell
                
                Set oLI = DetailLV.ListItems.Add()
                oLI.Text = ""
                oLI.Tag = "Spell"
                'nSpellNest = 0 'make sure this doesn't nest too deep
                oLI.ListSubItems.Add (1), "Detail", "Spell: [" & _
                    GetSpellName(tabMonsters.Fields("AttAcc-" & x), bHideRecordNumbers) _
                    & ", " & PullSpellEQ(True, tabMonsters.Fields("AttMax-" & x), tabMonsters.Fields("AttAcc-" & x), , , True) & "]"
                If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
                oLI.ListSubItems(1).Tag = tabMonsters.Fields("AttAcc-" & x)
                
                Set oLI = DetailLV.ListItems.Add()
                oLI.Text = ""
                oLI.ListSubItems.Add (1), "Detail", "Cast %: " & tabMonsters.Fields("AttMin-" & x)
                
                If tabMonsters.Fields("AttEnergy-" & x) > 0 Then
                    Set oLI = DetailLV.ListItems.Add()
                    oLI.Text = ""
                    oLI.ListSubItems.Add (1), "Detail", "Energy: " & tabMonsters.Fields("AttEnergy-" & x) _
                        & " (Max " & Fix(nMonsterEnergy / tabMonsters.Fields("AttEnergy-" & x)) & "x/round)"
                Else
                    Set oLI = DetailLV.ListItems.Add()
                    oLI.Text = ""
                    oLI.ListSubItems.Add (1), "Detail", "Energy: " & tabMonsters.Fields("AttEnergy-" & x)
                End If
                
                If Not tabMonsters.Fields("AttHitSpell-" & x) = 0 Then
                    Set oLI = DetailLV.ListItems.Add()
                    oLI.Text = ""
                    oLI.Tag = "Spell"
                    'nSpellNest = 0 'make sure this doesn't nest too deep
                    oLI.ListSubItems.Add (1), "Detail", "Hit Spell: [" & _
                        GetSpellName(tabMonsters.Fields("AttHitSpell-" & x), bHideRecordNumbers) _
                        & ", " & PullSpellEQ(False, , tabMonsters.Fields("AttHitSpell-" & x)) & "]"
                    If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
                    oLI.ListSubItems(1).Tag = tabMonsters.Fields("AttHitSpell-" & x)
                End If

        End Select
        
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = ""
    End If
Next


'For x = 0 To 9 'abilities
'    If Not tabMonsters.Fields("Abil-" & x) = 0 Then
'        Select Case tabMonsters.Fields("Abil-" & x)
'            Case 0:
'            Case 146: 'mon guards
'                If sMonGuards <> "" Then sMonGuards = sMonGuards & ", "
'                sMonGuards = sMonGuards & GetMonsterName(tabMonsters.Fields("AbilVal-" & x), bHideRecordNumbers)
'                tabMonsters.Seek "=", nMonsterNum
'
'            Case Else:
'                If sAbil <> "" Then sAbil = sAbil & ", "
'                sAbil = sAbil & GetAbilityStats(tabMonsters.Fields("Abil-" & x), tabMonsters.Fields("AbilVal-" & x))
'                If Right(sAbil, 2) = ", " Then sAbil = Left(sAbil, Len(sAbil) - 2)
'        End Select
'    End If
'Next
'
'DetailTB.Text = sAbil
'If Not sMonGuards = "" Then
'    sMonGuards = "Guarded By: " & sMonGuards
'    If sAbil = "" Then
'        DetailTB.Text = sMonGuards
'    Else
'        DetailTB.Text = sAbil & ", " & sMonGuards
'    End If
'End If

If frmMain.chkMonstersNoRegenLookUp.Value = 0 Then
    If Len(tabMonsters.Fields("Summoned By")) > 4 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = ""
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "Regens ..."
        Call frmMain.LookUpMonsterRegen(nMonsterNum, False, DetailLV)
    End If
End If

Set oLI = Nothing

Exit Function

Error:
Call HandleError
End Function

Public Function GetMonsterDamagePerRound(nMonsterNum As Long) As Long()
Dim x As Integer, y As Integer, z As Integer, nRound As Integer
Dim nMonEnergy As Long, nPercent As Integer, nReturn() As Long
Dim nMax As Long, nAVG As Long, nBetweenMax As Long, nBetweenAVG As Long
Dim sSpellEQ As String, sArr() As String, nTempVal1 As Currency, nTempVal2 As Currency
Dim nMonsterEnergy As Long, nRemainingEnergy As Currency

Dim nAtkMin(4) As Long, nAtkMax(4) As Long, nAtkType(4) As Integer
Dim nAtkAvgDmg(4) As Currency, nAtkAvgDmg_ADJ(4) As Currency
Dim nAtkEnergy(4) As Long, nAtkEnergy_ADJ(4) As Long
Dim nAtk_PCT(4) As Currency, nAtkSuccess_PCT(4) As Currency
Dim nAtk_HitSpellAvgDmg(4) As Currency, nAtk_HitSpellMaxDmg(4) As Long

On Error GoTo Error:

If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum

nPercent = 0
For x = 0 To 4 'between round spells
    If Not tabMonsters.Fields("MidSpell-" & x) = 0 Then
        
        nPercent = tabMonsters.Fields("MidSpell%-" & x) - nPercent
        
        sSpellEQ = PullSpellEQ(True, tabMonsters.Fields("MidSpellLVL-" & x), tabMonsters.Fields("MidSpell-" & x), , True, True)
        If InStr(1, sSpellEQ, ":") > 0 Then
            sArr() = Split(sSpellEQ, ":")
            If UBound(sArr()) >= 1 Then
                nTempVal1 = Val(sArr(0))
                nTempVal2 = Val(sArr(1))
                If UBound(sArr()) >= 2 Then
                    If Val(sArr(2)) > 0 Then
                        'normalize spell round (3sec) to combat round (5sec)
                        nTempVal1 = (nTempVal1 * 1.666) '/ Val(sArr(2))
                        nTempVal2 = (nTempVal2 * 1.666) '/ Val(sArr(2))
                    End If
                End If
                If nTempVal2 > nBetweenMax Then nBetweenMax = nTempVal2
                nBetweenAVG = nBetweenAVG + (((nTempVal1 + nTempVal2) / 2) * (nPercent / 100))
            End If
        End If
        
        If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
        nPercent = tabMonsters.Fields("MidSpell%-" & x)
    End If
Next

nPercent = 0
For x = 0 To 4
    nAtkType(x) = tabMonsters.Fields("AttType-" & x)
    nAtkEnergy(x) = tabMonsters.Fields("AttEnergy-" & x)
    
    If nAtkType(x) > 0 And tabMonsters.Fields("Att%-" & x) > 0 Then
        nAtk_PCT(x) = tabMonsters.Fields("Att%-" & x) - nPercent
        nPercent = tabMonsters.Fields("Att%-" & x)
        nAtk_PCT(x) = nAtk_PCT(x) / 100
        
        nAtk_HitSpellMaxDmg(x) = 0
        nAtk_HitSpellAvgDmg(x) = 0
        If Not tabMonsters.Fields("AttHitSpell-" & x) = 0 Then
            sSpellEQ = PullSpellEQ(False, , tabMonsters.Fields("AttHitSpell-" & x), , True, True)
            If InStr(1, sSpellEQ, ":") > 0 Then
                sArr() = Split(sSpellEQ, ":")
                If UBound(sArr()) >= 1 Then
                    nTempVal1 = Val(sArr(0))
                    nTempVal2 = Val(sArr(1))
                    
                    nAtk_HitSpellAvgDmg(x) = (nTempVal1 + nTempVal2) / 2
                    nAtk_HitSpellMaxDmg(x) = nTempVal2
                End If
            End If
        End If
        
        If nAtkType(x) = 2 Then 'spell
        
            nAtkSuccess_PCT(x) = tabMonsters.Fields("AttMin-" & x) / 100
            If nAtkSuccess_PCT(x) > 1 Then nAtkSuccess_PCT(x) = 1
            nAtkEnergy_ADJ(x) = (nAtkEnergy(x) * nAtkSuccess_PCT(x)) + ((nAtkEnergy(x) / 2) * (1 - nAtkSuccess_PCT(x)))
            
            sSpellEQ = PullSpellEQ(True, tabMonsters.Fields("AttMax-" & x), tabMonsters.Fields("AttAcc-" & x), , True, True)
            If InStr(1, sSpellEQ, ":") > 0 Then
                sArr() = Split(sSpellEQ, ":")
                If UBound(sArr()) >= 1 Then
                    nTempVal1 = Val(sArr(0))
                    nTempVal2 = Val(sArr(1))
                    If UBound(sArr()) >= 2 Then
                        If Val(sArr(2)) > 0 Then
                            'normalize spell rounds (3sec) to combat rounds (5sec)
                            nTempVal1 = (nTempVal1 * 1.666) '/ Val(sArr(2))
                            nTempVal2 = (nTempVal2 * 1.666) '/ Val(sArr(2))
                        End If
                    End If
                    nAtkMin(x) = nTempVal1
                    nAtkMax(x) = nTempVal2
                    
                    nAtkAvgDmg(x) = ((nAtkMin(x) + nAtkMax(x)) / 2) + nAtk_HitSpellAvgDmg(x)
                    nAtkAvgDmg_ADJ(x) = nAtkAvgDmg(x) * nAtkSuccess_PCT(x)
                End If
            End If
            If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
        Else
            
            nAtkSuccess_PCT(x) = tabMonsters.Fields("AttAcc-" & x) / 100
            
            'PUT AC ADJUSTMENT HERE
            nAtkSuccess_PCT(x) = 0.98
            If nAtkSuccess_PCT(x) > 0.98 Then nAtkSuccess_PCT(x) = 0.98
            
            
            nAtkEnergy_ADJ(x) = nAtkEnergy(x)
            
            nAtkMin(x) = tabMonsters.Fields("AttMin-" & x)
            nAtkMax(x) = tabMonsters.Fields("AttMax-" & x)
            
            nAtkAvgDmg(x) = ((nAtkMin(x) + nAtkMax(x)) / 2) + nAtk_HitSpellAvgDmg(x)
            nAtkAvgDmg_ADJ(x) = nAtkAvgDmg(x) * nAtkSuccess_PCT(x)
            
        End If
    End If
Next x

nRemainingEnergy = 0
If nNMRVer >= 1.71 Then
    nMonsterEnergy = tabMonsters.Fields("Energy")
Else
    nMonsterEnergy = 1000
End If

'Dim nAtkMin(4) As Long, nAtkMax(4) As Long, nAtkType(4) As Integer
'Dim nAtkAvgDmg(4) As Long, nAtkAvgDmg_ADJ(4) As Long
'Dim nAtkEnergy(4) As Long, nAtkEnergy_ADJ(4) As Long
'Dim nAtk_PCT(4) As Currency, nAtkSuccess_PCT(4) As Currency
Dim nAttackAttempt As Integer, nStartingEnergy As Long
Dim nChanceAtNotHavingEnergy_Attempt(4) As Currency, nChanceAtNotHavingEnergy_PREV(4) As Currency
Dim nEnergyUsageThatWouldCauseDeficit As Currency, nPreviousAttemptStartingEnergy As Long
Dim n_PREV_Atk_PCT(4) As Currency, nAtk_PCT_For_Attempt(4) As Currency, nDamageCaused As Currency
Dim nMaxRoundsToSim As Integer, nMaxEnergy As Long

nDamageCaused = 0
nRemainingEnergy = 0
nMaxRoundsToSim = 4
nMaxEnergy = nMonsterEnergy
For nRound = 1 To nMaxRoundsToSim
    'reset n_PREV_Atk_PCT
    For x = 0 To 4
        nChanceAtNotHavingEnergy_PREV(x) = 0
        n_PREV_Atk_PCT(x) = nAtk_PCT(x)
    Next x
    
    'add monster energy to left over energy
    nRemainingEnergy = nMonsterEnergy + nRemainingEnergy
    If nRemainingEnergy > nMaxEnergy Then nMaxEnergy = nRemainingEnergy
    nPreviousAttemptStartingEnergy = nRemainingEnergy
    
    '6 monster attack attempts
    For nAttackAttempt = 1 To 6
    
        nStartingEnergy = nRemainingEnergy
        For x = 0 To 4
            If nAtkAvgDmg_ADJ(x) > 0 Then
                nChanceAtNotHavingEnergy_Attempt(x) = 0
                If nAttackAttempt > 1 Then
                    nEnergyUsageThatWouldCauseDeficit = nPreviousAttemptStartingEnergy - nAtkEnergy(x)
                    For z = 0 To 4
                        If nAtkEnergy_ADJ(z) >= nEnergyUsageThatWouldCauseDeficit Then
                            nChanceAtNotHavingEnergy_Attempt(x) = nChanceAtNotHavingEnergy_Attempt(x) + n_PREV_Atk_PCT(z)
                        End If
                    Next z
                ElseIf nAtkEnergy(x) > nStartingEnergy Then
                    nChanceAtNotHavingEnergy_Attempt(x) = 1
                End If
                nChanceAtNotHavingEnergy_Attempt(x) = nChanceAtNotHavingEnergy_Attempt(x) + nChanceAtNotHavingEnergy_PREV(x)
                If nChanceAtNotHavingEnergy_Attempt(x) > 1 Then nChanceAtNotHavingEnergy_Attempt(x) = 1
                nAtk_PCT_For_Attempt(x) = (1 - nChanceAtNotHavingEnergy_Attempt(x)) * nAtk_PCT(x)
                
                nDamageCaused = nDamageCaused + (nAtkAvgDmg_ADJ(x) * nAtk_PCT_For_Attempt(x))
                nRemainingEnergy = nRemainingEnergy - (nAtkEnergy_ADJ(x) * nAtk_PCT_For_Attempt(x))
            End If
        Next x
        
        For x = 0 To 4
            nChanceAtNotHavingEnergy_PREV(x) = nChanceAtNotHavingEnergy_Attempt(x)
            n_PREV_Atk_PCT(x) = nAtk_PCT_For_Attempt(x)
        Next x
        
        nPreviousAttemptStartingEnergy = nStartingEnergy
        
    Next nAttackAttempt
    
Next nRound

Dim nMaxNumAttacks As Integer

nAVG = nDamageCaused / nMaxRoundsToSim
For x = 0 To 4
    If nAtkEnergy(x) > 0 Then
        nMaxNumAttacks = (nMaxEnergy / nAtkEnergy(x))
        If nMaxNumAttacks > 6 Then nMaxNumAttacks = 6
        If nMaxNumAttacks * (nAtkMax(x) + nAtk_HitSpellMaxDmg(x)) > nMax Then nMax = nMaxNumAttacks * (nAtkMax(x) + nAtk_HitSpellMaxDmg(x))
    End If
Next x
out:
On Error Resume Next

ReDim nReturn(0 To 1)
nReturn(0) = nAVG + nBetweenAVG
nReturn(1) = nMax + nBetweenMax

GetMonsterDamagePerRound = nReturn

Exit Function
Error:
Call HandleError("GetMonsterDamagePerRound")
Resume out:
End Function

Public Function PullShopDetail(nShopNum As Long, DetailLV As ListView, _
    DetailTB As TextBox, lvAssigned As ListView, ByVal nCharm As Integer, ByVal bShowSell As Boolean)
Dim sStr As String, x As Integer, nRegenTime As Integer, sRegenTime As String
Dim oLI As ListItem, tCostType As typItemCostDetail, nCopper As Currency, sCopper As String
Dim nCharmMod As Double, sCharmMod As String
Dim nReducedCoin As Currency, sReducedCoin As String

On Error GoTo Error:

'Call LockWindowUpdate(DetailLV.hwnd)

DetailLV.ListItems.clear

tabShops.Index = "pkShops"
tabShops.Seek "=", nShopNum
If tabShops.NoMatch = True Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Shop not found"
    DetailTB.Text = "Shop not found"
    Set oLI = Nothing
    Exit Function
End If


If nCharm > 0 Or bShowSell Then
    If bShowSell And Not tabShops.Fields("ShopType") = 8 Then
        nCharmMod = Fix(nCharm / 2) + 25
        
        sCharmMod = nCharmMod & "% cost (pre-markup)"
    Else
        nCharmMod = 1 - ((Fix(nCharm / 5) - 10) / 100)
        If nCharmMod > 1 Then
            sCharmMod = Abs(1 - CCur(nCharmMod)) * 100 & "% Markup"
        ElseIf nCharmMod < 1 Then
            sCharmMod = Val(1 - CCur(nCharmMod)) * 100 & "% Discount"
        Else
            sCharmMod = "Retail Value"
        End If
    End If
End If

    
If tabShops.Fields("ShopType") = 8 Then 'Case 8: GetShopType = "Training"
    If Not DetailLV.ColumnHeaders.Count = 2 Then
        DetailLV.ColumnHeaders.clear
        DetailLV.ColumnHeaders.Add 1, "Level", "LVL", 1000, lvwColumnLeft
        DetailLV.ColumnHeaders.Add 2, "Cost", "Cost", 4000, lvwColumnLeft
    End If
    
    frmMain.chkShopShowCharm(0).Enabled = False
    frmMain.chkShopShowCharm(1).Enabled = False
    
    frmMain.lblCharmMod.Caption = ""
    
    For x = tabShops.Fields("MinLVL") To tabShops.Fields("MaxLVL")
        sReducedCoin = ""
        nReducedCoin = 0
        
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = x
        nCopper = CalcMoneyRequiredToTrain(x - 1, tabShops.Fields("Markup%")) '* nCharmMod
        If nCopper < 0 Then nCopper = 0
        
        nCopper = Round(nCopper)
        
        sReducedCoin = "Copper"
        If nCopper >= 10000000 Then
            nReducedCoin = nCopper / 1000000
            sReducedCoin = "Runic"
        ElseIf nCopper >= 100000 Then
            nReducedCoin = nCopper / 10000
            sReducedCoin = "Platinum"
        ElseIf nCopper >= 1000 Then
            nReducedCoin = nCopper / 100
            sReducedCoin = "Gold"
        ElseIf nCopper >= 100 Then
            nReducedCoin = nCopper / 10
            sReducedCoin = "Silver"
        End If
        If nReducedCoin > 0 Then nReducedCoin = Round(nReducedCoin, 2)
        
        sCopper = Format(nCopper, "##,##0.00")
        If Right(sCopper, 3) = ".00" Then sCopper = Left(sCopper, Len(sCopper) - 3)
        
        If nReducedCoin = 0 Then
            oLI.ListSubItems.Add (1), "Cost", IIf(nCopper <= 0, "Free", sCopper & " Copper")
        Else
            sStr = Format(nReducedCoin, "##,##0.00")
            If Right(sStr, 3) = ".00" Then sStr = Left(sStr, Len(sStr) - 3)
            oLI.ListSubItems.Add (1), "Cost", Format(nCopper, "#,#") & " copper (" & sStr & " " & sReducedCoin & ")"
        End If
        
        oLI.ListSubItems(1).Tag = nCopper
    Next x
    
Else
    frmMain.chkShopShowCharm(0).Enabled = True
    frmMain.chkShopShowCharm(1).Enabled = True
    
    frmMain.lblCharmMod.Caption = sCharmMod
    
    If Not DetailLV.ColumnHeaders.Count = 5 Then
        DetailLV.ColumnHeaders.clear
        DetailLV.ColumnHeaders.Add 1, "Number", "#", 0, lvwColumnLeft
        DetailLV.ColumnHeaders.Add 2, "Name", "Name", 2000, lvwColumnCenter
        DetailLV.ColumnHeaders.Add 3, "Max", "Max", 550, lvwColumnCenter
        DetailLV.ColumnHeaders.Add 4, "Regen", "Regen", 1650, lvwColumnCenter
        'DetailLV.ColumnHeaders.Add 5, "Rgn%", "Rgn%", 600, lvwColumnCenter
        'DetailLV.ColumnHeaders.Add 6, "Rgn#", "Rgn#", 600, lvwColumnCenter
        DetailLV.ColumnHeaders.Add 5, "Cost", "Cost", 5000, lvwColumnLeft
    End If
    
    For x = 0 To 19
        sReducedCoin = ""
        nReducedCoin = 0
        
        If tabShops.Fields("Item-" & x) = 0 Then GoTo skip:
        
        Set oLI = DetailLV.ListItems.Add()
        
        oLI.Text = tabShops.Fields("Item-" & x)
        oLI.ListSubItems.Add (1), "Name", GetItemName(tabShops.Fields("Item-" & x), True)
        oLI.ListSubItems.Add (2), "Max", tabShops.Fields("Max-" & x)
        
        sRegenTime = ""
        If tabShops.Fields("Time-" & x) > 0 And tabShops.Fields("%-" & x) > 0 _
            And tabShops.Fields("Amount-" & x) > 0 Then
            
            nRegenTime = tabShops.Fields("Time-" & x)
            If nRegenTime >= (60 * 24) Then 'days
                sRegenTime = sRegenTime & Int(nRegenTime / (60 * 24)) & "d"
                nRegenTime = nRegenTime - (Int(nRegenTime / (60 * 24)) * (60 * 24))
            End If
            If nRegenTime >= 60 Then 'hours
                sRegenTime = sRegenTime & Int(nRegenTime / 60) & "h"
                nRegenTime = nRegenTime - (Int(nRegenTime / 60) * 60)
            End If
            If nRegenTime > 0 Then 'minutes
                sRegenTime = sRegenTime & nRegenTime & "m"
            End If
            
            sRegenTime = tabShops.Fields("%-" & x) & "% for " & tabShops.Fields("Amount-" & x) _
                & " per " & sRegenTime
        Else
            sRegenTime = "no regen"
        End If
        
        oLI.ListSubItems.Add (3), "Time", sRegenTime
'        oLI.ListSubItems.Add (4), "Rgn %", tabShops.Fields("%-" & x)
'        oLI.ListSubItems.Add (5), "Rgn #", tabShops.Fields("Amount-" & x)
        
        If bShowSell Then
            tCostType = GetItemCost(tabShops.Fields("Item-" & x), 0)
        Else
            tCostType = GetItemCost(tabShops.Fields("Item-" & x), tabShops.Fields("Markup%"))
        End If
        
        Select Case tCostType.Coin
            Case 0: 'GetCostType = "Copper"
                nCopper = Val(tCostType.Cost)
            Case 1: 'GetCostType = "Silver"
                nCopper = Val(tCostType.Cost) * 10
            Case 2: 'GetCostType = "Gold"
                nCopper = Val(tCostType.Cost) * 100
            Case 3: 'GetCostType = "Platinum"
                nCopper = Val(tCostType.Cost) * 10000
            Case 4: 'GetCostType = "Runic"
                nCopper = Val(tCostType.Cost) * 1000000
            Case Else:
                nCopper = tCostType.Cost
        End Select
        
        If nCharm > 0 Or bShowSell Then
            If bShowSell Then
                nCopper = nCharmMod * nCopper
                Do While nCopper > 4294967295# 'for the overflow bug
                    nCopper = nCopper - 4294967295#
                Loop
                nCopper = Fix(nCopper / 100)
                
                'nCopper = nCopper * nCharmMod
            Else
                nCopper = (nCharmMod * nCopper)
                Do While nCopper > 4294967295# 'for the overflow bug
                    nCopper = nCopper - 4294967295#
                Loop
                'nCopper = nCopper * nCharmMod
            End If
            If nCopper <= 0 Then nCopper = 0
        End If
        
        nCopper = Round(nCopper)
        
        sReducedCoin = "Copper"
        If nCopper >= 10000000 Then
            nReducedCoin = nCopper / 1000000
            sReducedCoin = "Runic"
        ElseIf nCopper >= 100000 Then
            nReducedCoin = nCopper / 10000
            sReducedCoin = "Platinum"
        ElseIf nCopper >= 1000 Then
            nReducedCoin = nCopper / 100
            sReducedCoin = "Gold"
        ElseIf nCopper >= 100 Then
            nReducedCoin = nCopper / 10
            sReducedCoin = "Silver"
        End If
        If nReducedCoin > 0 Then nReducedCoin = Round(nReducedCoin, 2)
        
        sCopper = Format(nCopper, "##,##0.00")
        If Right(sCopper, 3) = ".00" Then sCopper = Left(sCopper, Len(sCopper) - 3)
        
        If nReducedCoin = 0 Then
            oLI.ListSubItems.Add (4), "Cost", IIf(nCopper <= 0, "Free", sCopper & " Copper")
        Else
            sStr = Format(nReducedCoin, "##,##0.00")
            If Right(sStr, 3) = ".00" Then sStr = Left(sStr, Len(sStr) - 3)
            oLI.ListSubItems.Add (4), "Cost", Format(nCopper, "#,#") & " copper (" & sStr & " " & sReducedCoin & ")"
        End If
        oLI.ListSubItems(4).Tag = nCopper
skip:
    Next
End If


sStr = "Levels: " & tabShops.Fields("MinLVL") & " to " & tabShops.Fields("MaxLVL")
sStr = sStr & " -- Markup: " & tabShops.Fields("Markup%") & "%"

If Not tabShops.Fields("ClassRest") = 0 Then
    sStr = sStr & " -- Class: " & GetClassName(tabShops.Fields("ClassRest"))
End If
    
DetailTB.Text = sStr

Call GetLocations(tabShops.Fields("Assigned To"), lvAssigned)

Set oLI = Nothing

'Call LockWindowUpdate(0&)
Exit Function

Error:
Call HandleError
'Call LockWindowUpdate(0&)
End Function


Public Function PullSpellDetail(nSpellNum As Long, DetailTB As TextBox, LocationLV As ListView)
Dim sStr As String
Dim sRemoves As String, sArr() As String, x As Integer

On Error GoTo Error:

tabSpells.Index = "pkSpells"
tabSpells.Seek "=", nSpellNum
If tabSpells.NoMatch = True Then
    DetailTB.Text = "spell not found"
    Exit Function
End If

'sStr = ClipNull(tabSpells.Fields("Name")) & " (" & tabSpells.Fields("Number") & ") -- " & GetSpellTargets(tabSpells.Fields("Targets"))
sStr = "Target: " & GetSpellTargets(tabSpells.Fields("Targets"))

sStr = sStr & " -- Attack Type: " & GetSpellAttackType(tabSpells.Fields("AttType"))

'For x = 0 To 9
'    If Not tabSpells.Fields("Abil-" & x) = 0 Then
'        Select Case tabSpells.Fields("Abil-" & x)
'            Case 122: 'removes spell
'                If sRemoves <> "" Then sRemoves = sRemoves & ", "
'                sRemoves = sRemoves & GetSpellName(tabSpells.Fields("AbilVal-" & x))
'
'            Case Else:
'                If Not x = 0 Then
'                    If sAbil <> "" Then sAbil = sAbil & ", "
'                    sAbil = sAbil & GetAbilityStats(tabSpells.Fields("Abil-" & x), tabSpells.Fields("AbilVal-" & x))
'                    If Right(sAbil, 2) = ", " Then sAbil = Left(sAbil, Len(sAbil) - 2)
'                End If
'
'        End Select
'        'reposition in case the ability function changed it
'        If Not tabSpells.Fields("Number") = nSpellNum Then tabSpells.Seek "=", nSpellNum
'    End If
'Next

If Not tabSpells.Fields("Cap") = 0 Then
    sStr = sStr & " -- LVL Gain Cap: " & tabSpells.Fields("Cap")
End If

'nSpellNest = 0 'make sure this doesn't nest too deep
'sStr = sStr & " -- " & PullSpellEQ(False)

LocationLV.ListItems.clear

'nSpellNest = 0
If frmMain.chkGlobalFilter.Value = 1 Then
    sStr = sStr & " -- " & PullSpellEQ(True, Val(frmMain.txtGlobalLevel(0).Text), , LocationLV)
Else
    sStr = sStr & " -- " & PullSpellEQ(False, , , LocationLV)
End If

If Not tabSpells.Fields("Number") = nSpellNum Then tabSpells.Seek "=", nSpellNum

If tabSpells.Fields("EnergyCost") > 0 And tabSpells.Fields("EnergyCost") <= 500 Then
    sStr = sStr & ", x" & Fix(1000 / tabSpells.Fields("EnergyCost")) & " times/round"
End If

If nNMRVer >= 1.7 Then
    If Len(tabSpells.Fields("Classes")) > 2 And Not tabSpells.Fields("Classes") = "(*)" Then
        
        sStr = sStr & " -- Class Restricted (via learning method): "
        
        sArr() = StringOfNumbersToArray(tabSpells.Fields("Classes"))
        For x = 0 To UBound(sArr())
            If x > 0 Then sStr = sStr & ", "
            sStr = sStr & GetClassName(sArr(x))
        Next x
    End If
End If

If Not sRemoves = "" Then sStr = sStr & " -- Removes: " & sRemoves
'If Not sAbil = "" Then sStr = sStr & " -- Abilities: " & sAbil

'If tabSpells.Fields("Learnable") = 1 Then
'    sStr = sStr & " -- Learned From: " & GetLocations_STR(tabSpells.Fields("Learned From"))
'End If

DetailTB.Text = sStr
'Call GetLocations(tabSpells.Fields("Casted By"), lvCasted)

Call GetLocations(tabSpells.Fields("Learned From"), LocationLV, True, "(learn) ")
If Not tabSpells.Fields("Number") = nSpellNum Then tabSpells.Seek "=", nSpellNum
Call GetLocations(tabSpells.Fields("Casted By"), LocationLV, True)

Exit Function

Error:
Call HandleError
End Function

Public Function Get_Enc_Ratio(nENC As Long, nVal1 As Long, Optional nVal2 As Long) As Currency
Dim nTotal As Long

nTotal = nVal1 + nVal2

If nTotal > 0 Then
    If nENC < 1 Then
        Get_Enc_Ratio = nTotal '* 10
    Else
        Get_Enc_Ratio = Round(nTotal / nENC, 4) * 100 'Round((nTotal / 10) / nENC, 5) * 1000
    End If
Else
    Get_Enc_Ratio = 0
End If

End Function
Public Sub AddArmour2LV(LV As ListView, Optional AddToInven As Boolean, Optional nAbility As Integer)
Dim oLI As ListItem, x As Integer, sName As String, nAbilityVal As Integer

sName = tabItems.Fields("Name")
If sName = "" Then GoTo skip:

Set oLI = LV.ListItems.Add()
oLI.Text = tabItems.Fields("Number")

oLI.ListSubItems.Add (1), "Name", tabItems.Fields("Name")
oLI.ListSubItems.Add (2), "Worn", GetWornType(tabItems.Fields("Worn"))
oLI.ListSubItems.Add (3), "Armr Type", GetArmourType(tabItems.Fields("ArmourType"))
oLI.ListSubItems.Add (4), "Level", 0
oLI.ListSubItems.Add (5), "Enc", tabItems.Fields("Encum")
oLI.ListSubItems.Add (6), "AC", (tabItems.Fields("ArmourClass") / 10) & "/" & (tabItems.Fields("DamageResist") / 10)
oLI.ListSubItems(6).Tag = tabItems.Fields("ArmourClass") + tabItems.Fields("DamageResist")

oLI.ListSubItems.Add (7), "Acc", tabItems.Fields("Accy")
oLI.ListSubItems.Add (8), "Crits", 0
oLI.ListSubItems.Add (9), "Limit", tabItems.Fields("Limit")

oLI.ListSubItems.Add (10), "AC/Enc", _
    Get_Enc_Ratio(tabItems.Fields("Encum"), tabItems.Fields("ArmourClass"), tabItems.Fields("DamageResist"))
    
For x = 0 To 19
    Select Case tabItems.Fields("Abil-" & x)
        Case nAbility:
            nAbilityVal = tabItems.Fields("AbilVal-" & x)
            
        Case 58: ' crits
            oLI.ListSubItems(8).Text = tabItems.Fields("AbilVal-" & x)
            
        Case 135: 'min level
            oLI.ListSubItems(4).Text = tabItems.Fields("AbilVal-" & x)
        
        Case 22, 105, 106: 'acc
            oLI.ListSubItems(7).Text = Val(oLI.ListSubItems(7).Text) + tabItems.Fields("AbilVal-" & x)
    End Select
Next x

If nAbility > 0 Then
    oLI.ListSubItems.Add (11), "Ability", nAbilityVal
End If

If AddToInven Then Call frmMain.InvenAddEquip(tabItems.Fields("Number"), sName, tabItems.Fields("ItemType"), tabItems.Fields("Worn"))
    
skip:
Set oLI = Nothing
End Sub
Public Sub AddOtherItem2LV(LV As ListView)
Dim oLI As ListItem

If tabItems.Fields("Name") = "" Then GoTo skip:

Set oLI = LV.ListItems.Add()
oLI.Text = tabItems.Fields("Number")

oLI.ListSubItems.Add (1), "Name", tabItems.Fields("Name")
oLI.ListSubItems.Add (2), "Type", GetItemType(tabItems.Fields("ItemType"))
oLI.ListSubItems.Add (3), "Enc", tabItems.Fields("Encum")
oLI.ListSubItems.Add (4), "Limit", tabItems.Fields("Limit")

'For x = 0 To 9
'    If Not tabItems.Fields("Abil-" & x).Value = 0 Then
'        If sAbil <> "" Then sAbil = sAbil & ", "
'        sAbil = sAbil & GetAbilityStats(tabItems.Fields("Abil-" & x), tabItems.Fields("AbilVal-" & x))
'        If Right(sAbil, 2) = ", " Then sAbil = Left(sAbil, Len(sAbil) - 2)
'    End If
'Next

'oLI.ListSubItems.Add (5), "Abils", sAbil

skip:
Set oLI = Nothing
End Sub
Public Sub AddWeapon2LV(LV As ListView, Optional AddToInven As Boolean, Optional nAbility As Integer)
Dim oLI As ListItem, x As Integer, sName As String, nSpeed As Integer, nDMG As Integer, nAbilityVal As Integer

sName = tabItems.Fields("Name")
If sName = "" Then GoTo skip:

Set oLI = LV.ListItems.Add()
oLI.Text = tabItems.Fields("Number")

oLI.ListSubItems.Add (1), "Name", tabItems.Fields("Name")
oLI.ListSubItems.Add (2), "Wepn Type", GetWeaponType(tabItems.Fields("WeaponType"))
oLI.ListSubItems.Add (3), "Min Dmg", tabItems.Fields("Min")
oLI.ListSubItems.Add (4), "Max Dmg", tabItems.Fields("Max")
oLI.ListSubItems.Add (5), "Speed", tabItems.Fields("Speed")
oLI.ListSubItems.Add (6), "Level", 0
oLI.ListSubItems.Add (7), "Str", tabItems.Fields("StrReq")
oLI.ListSubItems.Add (8), "Enc", tabItems.Fields("Encum")
oLI.ListSubItems.Add (9), "AC", RoundUp(tabItems.Fields("ArmourClass") / 10) & "/" & (tabItems.Fields("DamageResist") / 10)
oLI.ListSubItems.Add (10), "Acc", tabItems.Fields("Accy")

oLI.ListSubItems.Add (11), "BS Acc", "No"
oLI.ListSubItems.Add (12), "Crits", 0

For x = 0 To 19
    Select Case tabItems.Fields("Abil-" & x)
        Case nAbility:
            nAbilityVal = tabItems.Fields("AbilVal-" & x)
            
        Case 58: 'crits
            oLI.ListSubItems(12).Text = tabItems.Fields("AbilVal-" & x)
            
        Case 22, 105, 106: 'acc
            oLI.ListSubItems(10).Text = Val(oLI.ListSubItems(9).Text) + tabItems.Fields("AbilVal-" & x)
        
        Case 135: 'min level
            oLI.ListSubItems(6).Text = tabItems.Fields("AbilVal-" & x)
            
        Case 116: 'bs accu
            oLI.ListSubItems(11).Text = tabItems.Fields("AbilVal-" & x)
    End Select
Next x

oLI.ListSubItems.Add (13), "Limit", tabItems.Fields("Limit")

nSpeed = tabItems.Fields("Speed")
nDMG = tabItems.Fields("Min") + tabItems.Fields("Max")

If nSpeed > 0 And nDMG > 0 Then
    oLI.ListSubItems.Add (14), "Dmg/Spd", Round(nDMG / nSpeed, 4) * 1000
Else
    oLI.ListSubItems.Add (14), "Dmg/Spd", 0
End If

nDMG = ((tabItems.Fields("Min") + tabItems.Fields("Max")) / 2) * 5
If nDMG > 0 Then
    oLI.ListSubItems.Add (15), "Dmg*5", nDMG
Else
    oLI.ListSubItems.Add (15), "Dmg*5", 0
End If

If nAbility > 0 Then
    oLI.ListSubItems.Add (16), "Ability", nAbilityVal
End If

If AddToInven Then Call frmMain.InvenAddEquip(tabItems.Fields("Number"), sName, tabItems.Fields("ItemType"), tabItems.Fields("Worn"))

skip:
Set oLI = Nothing
End Sub

Public Sub AddSpell2LV(LV As ListView, Optional ByVal AddBless As Boolean)
Dim oLI As ListItem, sName As String, x As Integer, nSpell As Long
Dim nSpellDamage As Currency

    nSpell = tabSpells.Fields("Number")
    sName = tabSpells.Fields("Name")
    If sName = "" Then GoTo skip:
    If Left(sName, 1) = "1" Then GoTo skip:
    If Left(LCase(sName), 3) = "sdf" Then GoTo skip:
    
    Set oLI = LV.ListItems.Add()
    oLI.Text = nSpell
    
    oLI.ListSubItems.Add (1), "Name", sName
    oLI.ListSubItems.Add (2), "Short", tabSpells.Fields("Short")
    oLI.ListSubItems.Add (3), "Magery", GetMagery(tabSpells.Fields("Magery"), tabSpells.Fields("MageryLVL"))
    oLI.ListSubItems.Add (4), "Level", tabSpells.Fields("ReqLevel")
    oLI.ListSubItems.Add (5), "Mana", tabSpells.Fields("ManaCost")
    oLI.ListSubItems.Add (6), "Diff", tabSpells.Fields("Diff")
    
'    If nSpell = 207 Then
'        Debug.Print nSpell
'    End If
    
    If tabSpells.Fields("Learnable") = 1 Then
        If frmMain.chkGlobalFilter.Value = 1 And Val(frmMain.txtGlobalLevel(1).Text) > 0 Then
            nSpellDamage = GetSpellMinDamage(nSpell, Val(frmMain.txtGlobalLevel(1).Text))
            nSpellDamage = nSpellDamage + GetSpellMaxDamage(nSpell, Val(frmMain.txtGlobalLevel(1).Text))
        Else
            nSpellDamage = GetSpellMinDamage(nSpell)
            nSpellDamage = nSpellDamage + GetSpellMaxDamage(nSpell)
        End If
        
        If Not tabSpells.Fields("Number") = nSpell Then
            tabSpells.Index = "pkSpells"
            tabSpells.Seek "=", nSpell
            If tabSpells.NoMatch = True Then Exit Sub
        End If
    
        nSpellDamage = nSpellDamage / 2
        If tabSpells.Fields("ManaCost") > 0 Then
            If tabSpells.Fields("EnergyCost") >= 200 And tabSpells.Fields("EnergyCost") <= 500 Then
                nSpellDamage = Round(nSpellDamage / (tabSpells.Fields("ManaCost") * Round(1000 / tabSpells.Fields("EnergyCost"))), 1)
            Else
                nSpellDamage = Round(nSpellDamage / tabSpells.Fields("ManaCost"), 1)
            End If
        End If
    Else
        nSpellDamage = 0
    End If
    
    oLI.ListSubItems.Add (7), "Dmg/M", nSpellDamage
    
    bQuickSpell = True
    oLI.ListSubItems.Add (8), "Detail", PullSpellEQ(False, , nSpell)
    bQuickSpell = False
    
    If Not tabSpells.Fields("Number") = nSpell Then tabSpells.Seek "=", nSpell
    
    If AddBless Then
        If Not tabSpells.Fields("Learnable") = 0 Then
                        
            Select Case tabSpells.Fields("Targets")
                Case 0: GoTo skip: 'GetSpellTargets = "User"
                Case 1: 'GetSpellTargets = "Self"
                Case 2: 'GetSpellTargets = "Self or User"
                Case 3: GoTo skip: 'GetSpellTargets = "Divided Area (not self)"
                Case 4: GoTo skip: 'GetSpellTargets = "Monster"
                Case 5: 'GetSpellTargets = "Divided Area (incl self)"
                Case 6: GoTo skip: 'GetSpellTargets = "Any"
                Case 7: GoTo skip: 'GetSpellTargets = "Item"
                Case 8: GoTo skip: 'GetSpellTargets = "Monster or User"
                Case 9: GoTo skip: 'GetSpellTargets = "Divided Attack Area"
                Case 10: GoTo skip: ' GetSpellTargets = "Divided Party Area"
                Case 11: GoTo skip: ' GetSpellTargets = "Full Area"
                Case 12: GoTo skip: ' GetSpellTargets = "Full Attack Area"
                Case 13: 'GetSpellTargets = "Full Party Area"
                Case Else: GoTo skip: 'GetSpellTargets = "Unknown (" & nNum & ")"
            End Select

            If tabSpells.Fields("Dur") > 0 Then
                'nothing
            ElseIf tabSpells.Fields("DurInc") > 0 And _
                tabSpells.Fields("DurIncLVLs") > 0 Then
                'nothing
            Else
                GoTo skip:
            End If
            
            For x = 0 To 9
                frmMain.cmbCharBless(x).AddItem sName & IIf(bHideRecordNumbers, "", " (" & nSpell & ")")
                frmMain.cmbCharBless(x).ItemData(frmMain.cmbCharBless(x).NewIndex) = nSpell
            Next x
        End If
    End If
    
skip:
Set oLI = Nothing
End Sub

Public Sub AddRace2LV(LV As ListView)
Dim oLI As ListItem, x As Integer, sAbil As String
    
    If tabRaces.Fields("Name") = "" Then GoTo skip:
    
    Set oLI = LV.ListItems.Add()
    oLI.Text = tabRaces.Fields("Number")
    
    oLI.ListSubItems.Add (1), "Name", tabRaces.Fields("Name")
    oLI.ListSubItems.Add (2), "Exp%", tabRaces.Fields("ExpTable") & "%"
    oLI.ListSubItems.Add (3), "HP", tabRaces.Fields("HPPerLVL")
    oLI.ListSubItems.Add (4), "Str", tabRaces.Fields("mSTR") & "-" & tabRaces.Fields("xSTR")
    oLI.ListSubItems.Add (5), "Int", tabRaces.Fields("mINT") & "-" & tabRaces.Fields("xINT")
    oLI.ListSubItems.Add (6), "Wis", tabRaces.Fields("mWIL") & "-" & tabRaces.Fields("xWIL")
    oLI.ListSubItems.Add (7), "Agi", tabRaces.Fields("mAGL") & "-" & tabRaces.Fields("xAGL")
    oLI.ListSubItems.Add (8), "Hea", tabRaces.Fields("mHEA") & "-" & tabRaces.Fields("xHEA")
    oLI.ListSubItems.Add (9), "Cha", tabRaces.Fields("mCHM") & "-" & tabRaces.Fields("xCHM")
    
    For x = 0 To 9
        Select Case tabRaces.Fields("Abil-" & x)
            Case 0:
            Case Else:
                If sAbil <> "" Then sAbil = sAbil & ", "
                sAbil = sAbil & GetAbilityStats(tabRaces.Fields("Abil-" & x), tabRaces.Fields("AbilVal-" & x))
                If Right(sAbil, 2) = ", " Then sAbil = Left(sAbil, Len(sAbil) - 2)
        End Select
    Next

    oLI.ListSubItems.Add (10), "Abilities", sAbil
    
skip:
Set oLI = Nothing
End Sub

Public Sub AddMonster2LV(LV As ListView)
Dim oLI As ListItem, sName As String, x As Integer, nMagicLVL As Integer, nExp As Currency
Dim nAvgDMG As Long, nExpDmgHP As Currency, nIndex As Integer, nPossSpawns As Long, sArr() As String
Dim nScriptValue As Currency

sName = tabMonsters.Fields("Name")
If sName = "" Or Left(sName, 3) = "sdf" Then GoTo skip:

Set oLI = LV.ListItems.Add()
oLI.Text = tabMonsters.Fields("Number")

nIndex = 1

oLI.ListSubItems.Add (nIndex), "Name", sName
nIndex = nIndex + 1

oLI.ListSubItems.Add (nIndex), "RGN", tabMonsters.Fields("RegenTime")
oLI.ListSubItems(nIndex).Tag = tabMonsters.Fields("RegenTime")
nIndex = nIndex + 1

If UseExpMulti Then
    nExp = tabMonsters.Fields("EXP") * tabMonsters.Fields("ExpMulti")
Else
    nExp = tabMonsters.Fields("EXP")
End If
oLI.ListSubItems.Add (nIndex), "Exp", IIf(nExp > 0, Format(nExp, "#,#"), 0)
oLI.ListSubItems(nIndex).Tag = nExp
nIndex = nIndex + 1

oLI.ListSubItems.Add (nIndex), "HP", IIf(nExp > 0, Format(tabMonsters.Fields("HP"), "#,#"), 0)
oLI.ListSubItems(nIndex).Tag = tabMonsters.Fields("HP")
nIndex = nIndex + 1

nPossSpawns = 0
If Len(tabMonsters.Fields("Summoned By")) > 4 Then
    sArr() = Split(tabMonsters.Fields("Summoned By"), ",")
    For x = 0 To UBound(sArr())
        If InStr(1, sArr(x), "(lair)", vbTextCompare) Then nPossSpawns = nPossSpawns + 1
    Next x
End If

If nNMRVer >= 1.8 Then
    nAvgDMG = tabMonsters.Fields("AvgDmg")
End If
oLI.ListSubItems.Add (nIndex), "Damage", IIf(nAvgDMG > 0, Format(nAvgDMG, "#,#"), 0)
oLI.ListSubItems(nIndex).Tag = nAvgDMG
nIndex = nIndex + 1

If nAvgDMG > 0 Or tabMonsters.Fields("HP") > 0 Then
    nExpDmgHP = Round(nExp / (nAvgDMG + tabMonsters.Fields("HP")), 2) * 100
    oLI.ListSubItems.Add (nIndex), "Exp/(Dmg+HP)", IIf(nExpDmgHP > 0, Format(nExpDmgHP, "#,#"), 0)
    oLI.ListSubItems(nIndex).Tag = nExpDmgHP
Else
    nExpDmgHP = 0
    oLI.ListSubItems.Add (nIndex), "Exp/(Dmg+HP)", IIf(nExp > 0, Format(nExp, "#,#"), 0)
    oLI.ListSubItems(nIndex).Tag = nExp
End If
nIndex = nIndex + 1

oLI.ListSubItems.Add (nIndex), "Lairs", nPossSpawns
oLI.ListSubItems(nIndex).Tag = nPossSpawns
nIndex = nIndex + 1

If tabMonsters.Fields("RegenTime") > 0 Or nPossSpawns < 1 Then
    nScriptValue = 0
Else
    nScriptValue = nExpDmgHP * nPossSpawns
End If

If nScriptValue > 1000000000 Then
    oLI.ListSubItems.Add (nIndex), "Script Value", Format((nScriptValue / 1000000), "#,#M")
ElseIf nScriptValue > 1000000 Then
    oLI.ListSubItems.Add (nIndex), "Script Value", Format((nScriptValue / 1000), "#,#K")
Else
    oLI.ListSubItems.Add (nIndex), "Script Value", IIf(nScriptValue > 0, Format(nScriptValue, "#,#"), "0")
End If
oLI.ListSubItems(nIndex).Tag = nScriptValue
nIndex = nIndex + 1


skip:
Set oLI = Nothing
End Sub

Public Sub AddShop2LV(LV As ListView)
Dim oLI As ListItem, sName As String
    
    sName = tabShops.Fields("Name")
    If sName = "" Or Left(LCase(sName), 3) = "sdf" Then GoTo skip:
    If sName = "Leave this blank" Then GoTo skip:
    
    Set oLI = LV.ListItems.Add()
    oLI.Text = tabShops.Fields("Number")
    
    oLI.ListSubItems.Add (1), "Name", sName
    oLI.ListSubItems.Add (2), "Type", GetShopType(tabShops.Fields("ShopType"))
    
skip:
Set oLI = Nothing
End Sub

Public Sub AddClass2LV(LV As ListView)
Dim oLI As ListItem, x As Integer, sAbil As String
    
    If tabClasses.Fields("Name") = "" Then GoTo skip:
    
    Set oLI = LV.ListItems.Add()
    oLI.Text = tabClasses.Fields("Number")
    
    oLI.ListSubItems.Add (1), "Name", tabClasses.Fields("Name")
    oLI.ListSubItems.Add (2), "Exp%", (Val(tabClasses.Fields("ExpTable")) + 100) & "%"
    oLI.ListSubItems.Add (3), "Weapon", GetClassWeaponType(tabClasses.Fields("WeaponType"))
    oLI.ListSubItems.Add (4), "Armour", GetArmourType(tabClasses.Fields("ArmourType"))
    oLI.ListSubItems.Add (5), "Magic", GetMagery(tabClasses.Fields("MageryType"), tabClasses.Fields("MageryLVL"))
    oLI.ListSubItems.Add (6), "Cmbt", tabClasses.Fields("CombatLVL") - 2
    oLI.ListSubItems.Add (7), "HP", tabClasses.Fields("MinHits") & "-" & (tabClasses.Fields("MinHits") + tabClasses.Fields("MaxHits"))
    
    For x = 0 To 9
        Select Case tabClasses.Fields("Abil-" & x)
            Case 0:
            Case 59: 'class ok
            Case Else:
                If sAbil <> "" Then sAbil = sAbil & ", "
                sAbil = sAbil & GetAbilityStats(tabClasses.Fields("Abil-" & x), tabClasses.Fields("AbilVal-" & x))
                If Right(sAbil, 2) = ", " Then sAbil = Left(sAbil, Len(sAbil) - 2)
        End Select
    Next

    oLI.ListSubItems.Add (8), "Abilities", sAbil
    
skip:
Set oLI = Nothing
End Sub



Public Sub RaceColorCode(LV As ListView)
Dim oLI As ListItem, x As Integer
Dim Stat(1 To 6, 1 To 2) As Integer, Min(1 To 6) As Integer, Max(1 To 6) As Integer, nRaces As Integer
'1-6 = str, int, wis, agi, hea, cha

'stat, 1 = min
'stat, 2 = max
'min, total min
'max, total max
'then get avg

For Each oLI In LV.ListItems
    
    tabRaces.Index = "pkRaces"
    tabRaces.Seek "=", Val(oLI.Text)
    If tabRaces.NoMatch = False Then
        
        nRaces = nRaces + 1
        Stat(1, 1) = Val(tabRaces.Fields("mSTR"))
        Stat(2, 1) = Val(tabRaces.Fields("mINT"))
        Stat(3, 1) = Val(tabRaces.Fields("mWIL"))
        Stat(4, 1) = Val(tabRaces.Fields("mAGL"))
        Stat(5, 1) = Val(tabRaces.Fields("mHEA"))
        Stat(6, 1) = Val(tabRaces.Fields("mCHM"))
        Stat(1, 2) = Val(tabRaces.Fields("xSTR"))
        Stat(2, 2) = Val(tabRaces.Fields("xINT"))
        Stat(3, 2) = Val(tabRaces.Fields("xWIL"))
        Stat(4, 2) = Val(tabRaces.Fields("xAGL"))
        Stat(5, 2) = Val(tabRaces.Fields("xHEA"))
        Stat(6, 2) = Val(tabRaces.Fields("xCHM"))
        
        For x = 1 To 6
            Min(x) = Min(x) + Stat(x, 1)
            Max(x) = Max(x) + Stat(x, 2)
        Next x
        
    End If
Next oLI

If nRaces = 0 Then Exit Sub

For x = 1 To 6
    Stat(x, 1) = Min(x) / nRaces 'avg min
    Stat(x, 2) = Max(x) / nRaces 'avg max
    Stat(x, 1) = Stat(x, 1) + Stat(x, 2) 'avg total
Next

For Each oLI In LV.ListItems
    
    tabRaces.Index = "pkRaces"
    tabRaces.Seek "=", Val(oLI.Text)
    If tabRaces.NoMatch = False Then
        
        
        If Val(tabRaces.Fields("mSTR")) + Val(tabRaces.Fields("xSTR")) < Stat(1, 1) - 20 Then oLI.ListSubItems(4).ForeColor = &H80&
        If Val(tabRaces.Fields("mINT")) + Val(tabRaces.Fields("xINT")) < Stat(2, 1) - 20 Then oLI.ListSubItems(5).ForeColor = &H80&
        If Val(tabRaces.Fields("mWIL")) + Val(tabRaces.Fields("xWIL")) < Stat(3, 1) - 20 Then oLI.ListSubItems(6).ForeColor = &H80&
        If Val(tabRaces.Fields("mAGL")) + Val(tabRaces.Fields("xAGL")) < Stat(4, 1) - 20 Then oLI.ListSubItems(7).ForeColor = &H80&
        If Val(tabRaces.Fields("mHEA")) + Val(tabRaces.Fields("xHEA")) < Stat(5, 1) - 20 Then oLI.ListSubItems(8).ForeColor = &H80&
        If Val(tabRaces.Fields("mCHM")) + Val(tabRaces.Fields("xCHM")) < Stat(6, 1) - 20 Then oLI.ListSubItems(9).ForeColor = &H80&
        
        If Val(tabRaces.Fields("mSTR")) + Val(tabRaces.Fields("xSTR")) > Stat(1, 1) + 20 Then oLI.ListSubItems(4).ForeColor = &H8000&
        If Val(tabRaces.Fields("mINT")) + Val(tabRaces.Fields("xINT")) > Stat(2, 1) + 20 Then oLI.ListSubItems(5).ForeColor = &H8000&
        If Val(tabRaces.Fields("mWIL")) + Val(tabRaces.Fields("xWIL")) > Stat(3, 1) + 20 Then oLI.ListSubItems(6).ForeColor = &H8000&
        If Val(tabRaces.Fields("mAGL")) + Val(tabRaces.Fields("xAGL")) > Stat(4, 1) + 20 Then oLI.ListSubItems(7).ForeColor = &H8000&
        If Val(tabRaces.Fields("mHEA")) + Val(tabRaces.Fields("xHEA")) > Stat(5, 1) + 20 Then oLI.ListSubItems(8).ForeColor = &H8000&
        If Val(tabRaces.Fields("mCHM")) + Val(tabRaces.Fields("xCHM")) > Stat(6, 1) + 20 Then oLI.ListSubItems(9).ForeColor = &H8000&
        
    End If
Next

Set oLI = Nothing
End Sub

Public Function SearchLV(ByVal KeyCode As Integer, oLVW As ListView, oTXT As TextBox) As Boolean
Dim i As Long, SearchStart As Long, SelectText As String, bSearchAgain As Boolean
Dim nCIndex As Integer

On Error GoTo Error:

If oLVW.ListItems.Count < 1 Then Exit Function

If KeyCode = vbKeyUp Then Exit Function
If KeyCode = vbKeyLeft Then Exit Function
If KeyCode = vbKeyBack Then Exit Function
If KeyCode = vbKeyDown Then oLVW.SetFocus
If KeyCode = vbKeyRight Then bSearchAgain = True
If KeyCode = vbKeyShift Then Exit Function 'shift
If KeyCode = vbKeyControl Then Exit Function 'control
If KeyCode = 18 Then Exit Function 'alt
If KeyCode = vbKeyTab Then Exit Function 'alt

If oTXT.Text = "" Then Exit Function
SelectText = oTXT.Text

If bSearchAgain = True Then 'searching for next?
    SearchStart = oLVW.SelectedItem.Index + 1
Else
    SearchStart = 1
End If

If Not SearchStart + 1 <= oLVW.ListItems.Count Then Exit Function 'if it's the last item in the list

'LockWindowUpdate oLVW.hwnd

nCIndex = oLVW.SelectedItem.Index

For i = SearchStart To oLVW.ListItems.Count
    If Not InStr(1, LCase(oLVW.ListItems(i).ListSubItems(1).Text), LCase(SelectText)) = 0 Then
        If Not i = nCIndex Then
            SearchLV = True
            nCIndex = i
        End If
        Exit For
    End If
Next
        
If SearchLV Then
    For i = 1 To oLVW.ListItems.Count
        oLVW.ListItems(i).Selected = False
    Next
    Set oLVW.SelectedItem = oLVW.ListItems(nCIndex)
    oLVW.SelectedItem.EnsureVisible
End If

'LockWindowUpdate 0&

Exit Function
Error:
Call HandleError
'LockWindowUpdate 0&
End Function

Public Sub CopyWholeLVtoClipboard(LV As ListView, Optional ByVal UsePeriods As Boolean)
On Error GoTo Error:
Dim oLI As ListItem, oLSI As ListSubItem, oCH As ColumnHeader
Dim str As String, x As Integer, sSpacer As String, nLongText() As Integer
    
str = ""
sSpacer = IIf(UsePeriods, ".", " ")

ReDim nLongText(0 To LV.ColumnHeaders.Count - 1)

'find longest text(s)
For Each oLI In LV.ListItems
    If Len(oLI.Text) > nLongText(0) Then nLongText(0) = Len(oLI.Text)
    x = 1
    For Each oLSI In oLI.ListSubItems
        If Len(oLSI.Text) > nLongText(x) Then nLongText(x) = Len(oLSI.Text)
        x = x + 1
    Next
Next
'put on 3 spaces
For x = 0 To UBound(nLongText())
    nLongText(x) = nLongText(x) + 3
Next

x = 0
For Each oCH In LV.ColumnHeaders
    str = str & oCH.Text
    str = str & " " & String(nLongText(x) - Len(oCH.Text), " ") & " "
    x = x + 1
Next

str = str & vbCrLf

For Each oLI In LV.ListItems
    str = str & oLI.Text
    str = str & " " & String(nLongText(0) - Len(oLI.Text), sSpacer) & " "
    
    x = 1
    For Each oLSI In oLI.ListSubItems
        str = str & oLSI.Text
        If Not x = oLI.ListSubItems.Count Then str = str & " " & String(nLongText(x) - Len(oLSI.Text), sSpacer) & " "
        x = x + 1
    Next
    str = str & vbCrLf
Next

If Not str = "" Then
    Clipboard.clear
    Clipboard.SetText str
End If

Set oLI = Nothing
Set oLSI = Nothing
Set oCH = Nothing

Exit Sub
Error:
Call HandleError("CopyWholeLVtoClip")
Set oLI = Nothing
Set oLSI = Nothing
Set oCH = Nothing
End Sub
Public Sub CopyLVLinetoClipboard(LV As ListView, Optional DetailTB As TextBox, _
    Optional LocationLV As ListView, Optional ByVal nExcludeColumn As Integer = -1, Optional bNameOnly As Boolean = False)
On Error GoTo Error:
Dim oLI As ListItem, oLI2 As ListItem, oCH As ColumnHeader
Dim str As String, x As Integer, nCount As Integer

If LV.ListItems.Count < 1 Then Exit Sub

nCount = 1
For Each oLI In LV.ListItems
    If oLI.Selected Then
        If nCount > 100 Then GoTo done:
        If nCount > 1 Then
            If bNameOnly Then
                str = str & ", "
            Else
                str = str & vbCrLf & vbCrLf
            End If
        End If
        
        x = 0
        For Each oCH In LV.ColumnHeaders
            If Not x = nExcludeColumn Then
                If bNameOnly Then
                    If oCH.Text = "Name" Then
                        If x = 0 Then
                            str = str & oLI.Text
                        Else
                            str = str & oLI.SubItems(x)
                        End If
                    End If
                Else
                    If Not x = 0 Then str = str & ", "
                    
                    str = str & oCH.Text & ": "
                    'If Len(oCH.Text) <= 9 Then str = str & String(10 - Len(oCH.Text), " ")
                    
                    If x = 0 Then
                        str = str & oLI.Text
                    Else
                        str = str & oLI.SubItems(x)
                    End If
                End If
            End If
            x = x + 1
        Next oCH
        
        Select Case LV.name
            Case "lvWeapons":
                Call frmMain.lvWeapons_ItemClick(oLI)
            Case "lvArmour":
                Call frmMain.lvArmour_ItemClick(oLI)
            Case "lvSpells":
                Call frmMain.lvSpells_ItemClick(oLI)
            Case "lvWeaponCompare":
                Call frmMain.lvWeaponCompare_ItemClick(oLI)
            Case "lvArmourCompare":
                Call frmMain.lvArmourCompare_ItemClick(oLI)
            Case "lvSpellCompare":
                Call frmMain.lvSpellCompare_ItemClick(oLI)
        End Select
        
        If Not bNameOnly And Not DetailTB Is Nothing Then
            If Not DetailTB.Text = "" Then
                str = str & vbCrLf & ">> " & DetailTB.Text
            End If
        End If
        
        If Not bNameOnly And Not LocationLV Is Nothing Then
            If LocationLV.ListItems.Count > 0 Then
                str = str & vbCrLf & ">> " & LocationLV.ColumnHeaders(1).Text & ": "
                x = 1
                For Each oLI2 In LocationLV.ListItems
                    If x > 1 Then str = str & ", "
                    str = str & oLI2.Text
                    x = x + 1
                Next oLI2
                Set oLI2 = Nothing
            End If
        End If
        nCount = nCount + 1
    End If
    Set oLI = Nothing
    DoEvents
Next oLI

done:
If Not str = "" Then
    Clipboard.clear
    Clipboard.SetText str
End If

out:
Set oLI = Nothing
Set oLI2 = Nothing
Set oCH = Nothing

Exit Sub
Error:
Call HandleError("CopyLVLinetoClip")
Resume out:
End Sub

Public Sub GetLocations(ByVal sLoc As String, LV As ListView, _
    Optional bDontClear As Boolean, Optional ByVal sHeader As String, _
    Optional ByVal nAuxValue As Long, Optional ByVal bTwoColumns As Boolean, _
    Optional ByVal bDontSort As Boolean)
On Error GoTo Error:
Dim sLook As String, sChar As String, sTest As String, oLI As ListItem, sPercent As String
Dim x As Integer, y1 As Integer, y2 As Integer, z As Integer, nValue As Long, x2 As Integer
Dim sLocation As String

If Not bDontClear Then LV.ListItems.clear
If bDontSort Then LV.Sorted = False

If Len(sLoc) < 5 Then Exit Sub

sTest = LCase(sLoc)

For z = 1 To 12
    
    x = 1
    Select Case z
        Case 1: sLook = "room "
        Case 2: sLook = "monster #"
        Case 3: sLook = "textblock #"
        Case 4: sLook = "textblock(rndm) #"
        Case 5: sLook = "item #"
        Case 6: sLook = "spell #"
        Case 7: sLook = "shop #"
        Case 8: sLook = "shop(sell) #"
        Case 9: sLook = "shop(nogen) #"
        Case 10: sLook = "group(lair): "
        Case 11: sLook = "group: "
        Case 12: sLook = "npc #"
    End Select

checknext:
    sPercent = ""
    If Not InStr(x, sTest, sLook) = 0 Then
        
        x = InStr(x, sTest, sLook) 'sets x to the position of the string we're looking for
        
'        If z = 10 Then
'            y1 = x + 1
'            GoTo nonumber:
'        End If
        
        y1 = x + Len(sLook) 'len of string searching (to position y1 at first number)
        y2 = 0
nextnumber:
        sChar = Mid(sTest, y1 + y2, 1)
        Select Case sChar
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "/":
                If Not y1 + y2 - 1 = Len(sTest) Then
                    y2 = y2 + 1
                    GoTo nextnumber:
                End If
            Case "+": 'end of string
                Exit Sub
            Case "(": 'precent
                x2 = InStr(y1 + y2, sTest, ")")
                If Not x2 = 0 Then
                    sPercent = " " & Mid(sTest, y1 + y2, x2 - y1 - y2 + 1)
                End If
            Case Else:
        End Select
        
        If y2 = 0 Then
            'if there were no numbers after the string
            x = y1
            GoTo checknext:
        End If
        
        If Not z = 1 Or z = 10 Or z = 11 Then 'not room or group
            nValue = Val(Mid(sTest, y1, y2))
        End If

nonumber:
        Select Case z
            Case 1: '"room "
                sLocation = "Room: "
                Set oLI = LV.ListItems.Add()
                If bTwoColumns Then
                    oLI.Text = IIf(sHeader = "", sLocation, sHeader)
                    oLI.ListSubItems.Add 1, , GetRoomName(Mid(sTest, y1, y2), , , bHideRecordNumbers) & sPercent
                    oLI.Tag = "Room"
                Else
                    oLI.Text = IIf(sHeader = "", sLocation, sHeader) _
                        & GetRoomName(Mid(sTest, y1, y2), , , bHideRecordNumbers) & sPercent
                End If
                
                If nAuxValue > 0 Then
                    If bTwoColumns Then
                        oLI.ListSubItems(1).Tag = nAuxValue
                    Else
                        oLI.Tag = nAuxValue
                    End If
                Else
                    If bTwoColumns Then
                        oLI.ListSubItems(1).Tag = Mid(sTest, y1, y2)
                    Else
                        oLI.Tag = Mid(sTest, y1, y2)
                    End If
                End If
                
            Case 2: '"monster #"
                sLocation = "Monster: "
                Set oLI = LV.ListItems.Add()
                If bTwoColumns Then
                    oLI.Text = sLocation & sHeader
                    oLI.ListSubItems.Add 1, , GetMonsterName(nValue, bHideRecordNumbers) & sPercent
                    oLI.Tag = "monster"
                    oLI.ListSubItems(1).Tag = nValue
                Else
                    oLI.Text = sLocation & sHeader & GetMonsterName(nValue, bHideRecordNumbers) & sPercent
                    oLI.Tag = nValue
                End If
                
            Case 3: '"textblock #"
                sLocation = "Textblock "
                Set oLI = LV.ListItems.Add()
                If bTwoColumns Then
                    oLI.Text = sLocation & sHeader
                    oLI.ListSubItems.Add 1, , nValue & sPercent
                    oLI.Tag = "textblock"
                    oLI.ListSubItems(1).Tag = nValue
                Else
                    oLI.Text = sLocation & sHeader & nValue & sPercent
                    oLI.Tag = nValue
                End If
                
            Case 4: '"textblock(rndm) #"
                sLocation = "Textblock "
                Set oLI = LV.ListItems.Add()
                If bTwoColumns Then
                    oLI.Text = sLocation & sHeader
                    oLI.ListSubItems.Add 1, , nValue & sPercent '& " (random)" & sPercent
                    oLI.Tag = "textblock"
                    oLI.ListSubItems(1).Tag = nValue
                Else
                    oLI.Text = sLocation & sHeader & nValue & sPercent '& " (random)" & sPercent
                    oLI.Tag = nValue
                End If
                
            Case 5: '"item #"
                sLocation = "Item: "
                Set oLI = LV.ListItems.Add()
                If bTwoColumns Then
                    oLI.Text = sLocation & sHeader
                    oLI.ListSubItems.Add 1, , GetItemName(nValue, bHideRecordNumbers) & sPercent
                    oLI.Tag = "item"
                    oLI.ListSubItems(1).Tag = nValue
                Else
                    oLI.Text = sLocation & sHeader & GetItemName(nValue, bHideRecordNumbers) & sPercent
                    oLI.Tag = nValue
                End If
                
            Case 6: '"spell #"
                sLocation = "Spell: "
                Set oLI = LV.ListItems.Add()
                If bTwoColumns Then
                    oLI.Text = sLocation & sHeader
                    oLI.ListSubItems.Add 1, , GetSpellName(nValue, bHideRecordNumbers) & sPercent
                    oLI.Tag = "spell"
                    oLI.ListSubItems(1).Tag = nValue
                Else
                    oLI.Text = sLocation & sHeader & GetSpellName(nValue, bHideRecordNumbers) & sPercent
                    oLI.Tag = nValue
                End If
                
            Case 7: '"shop #"
                Call GetLocations(GetShopLocation(nValue), LV, True, "Shop: ", nValue)
                'Set oLI = LV.ListItems.Add()
                'oLI.Text = "Shop: " & GetShopName(nValue) & sPercent
                'oLI.Tag = nValue
            Case 8: '"shop(sell) #"
                Call GetLocations(GetShopLocation(nValue), LV, True, "Shop (sell): ", nValue)
'                Set oLI = LV.ListItems.Add()
'                oLI.Text = "Shop: " & GetShopName(nValue) & " (sell only)" & sPercent
'                oLI.Tag = nValue
            Case 9: '"shop(nogen) #"
                Call GetLocations(GetShopLocation(nValue), LV, True, "Shop (nogen): ", nValue)
'                Set oLI = LV.ListItems.Add()
'                oLI.Text = "Shop: " & GetShopName(nValue) & " (wont regen)" & sPercent
'                oLI.Tag = nValue
            Case 10: 'group (lair)
                sLocation = "Group(Lair): "
                Set oLI = LV.ListItems.Add()
                If bTwoColumns Then
                    oLI.Text = sLocation
                    oLI.ListSubItems.Add 1, , GetRoomName(Mid(sTest, y1, y2), , , bHideRecordNumbers) & sPercent
                    oLI.Tag = "Room"
                    oLI.ListSubItems(1).Tag = Mid(sTest, y1, y2)
                Else
                    oLI.Text = sLocation & GetRoomName(Mid(sTest, y1, y2), , , bHideRecordNumbers) & sPercent
                    oLI.Tag = Mid(sTest, y1, y2)
                End If
                
            Case 11: 'group
                sLocation = "Group: "
                Set oLI = LV.ListItems.Add()
                If bTwoColumns Then
                    oLI.Text = sLocation
                    oLI.ListSubItems.Add 1, , GetRoomName(Mid(sTest, y1, y2), , , bHideRecordNumbers) & sPercent
                    oLI.Tag = "Room"
                    oLI.ListSubItems(1).Tag = Mid(sTest, y1, y2)
                Else
                    oLI.Text = sLocation & GetRoomName(Mid(sTest, y1, y2), , , bHideRecordNumbers) & sPercent
                    oLI.Tag = Mid(sTest, y1, y2)
                End If
            
            Case 12: '"NPC #"
                sLocation = "NPC: "
                Set oLI = LV.ListItems.Add()
                If bTwoColumns Then
                    oLI.Text = sLocation & sHeader
                    oLI.ListSubItems.Add 1, , GetMonsterName(nValue, bHideRecordNumbers) & sPercent
                    oLI.Tag = "monster"
                    oLI.ListSubItems(1).Tag = nValue
                Else
                    oLI.Text = sLocation & sHeader & GetMonsterName(nValue, bHideRecordNumbers) & sPercent
                    oLI.Tag = nValue
                End If
        End Select
        
        x = y1
        GoTo checknext:
    End If
Next z

If LV.ListItems.Count > 1 And Not bDontSort Then
    Call SortListView(LV, 1, ldtstring, True)
    LV.Sorted = False
End If

If LV.ListItems.Count > 1 Then
    If Right(sLoc, 2) = "+" & Chr(0) Then
        Set oLI = LV.ListItems.Add(LV.ListItems.Count + 1)
        If bTwoColumns Then
            oLI.ListSubItems.Add 1, , "... plus more."
        Else
            oLI.Text = "... plus more."
        End If
        oLI.Tag = 0
    End If
End If

Set oLI = Nothing
Exit Sub

Error:
HandleError
Set oLI = Nothing

End Sub

Private Function GetShopLocation(ByVal nNum As Long) As String

tabShops.Index = "pkShops"
tabShops.Seek "=", nNum
If tabShops.NoMatch Then
    GetShopLocation = ""
    Exit Function
End If

GetShopLocation = tabShops.Fields("Assigned To")
End Function

'Public Function GetLocations_STR(ByVal sLoc As String) As String
'Dim sLook As String, sChar As String, sTest As String, sSuffix As String
'Dim x As Integer, y1 As Integer, y2 As Integer, z As Integer, nValue As Long, x2 As Integer
'
'If Len(sLoc) < 5 Then
'    GetLocations_STR = "None."
'    Exit Function
'End If
'
'sTest = LCase(sLoc)
'
'For z = 1 To 10
'
'    x = 1
'    Select Case z
'        Case 1: sLook = "room "
'        Case 2: sLook = "monster #"
'        Case 3: sLook = "textblock #"
'        Case 4: sLook = "textblock(rndm) #"
'        Case 5: sLook = "item #"
'        Case 6: sLook = "spell #"
'        Case 7: sLook = "shop #"
'        Case 8: sLook = "shop(sell) #"
'        Case 9: sLook = "shop(nogen) #"
'        Case 10: sLook = "group "
'    End Select
'
'checknext:
'    sSuffix = ""
'    If Not InStr(x, sTest, sLook) = 0 Then
'
'
'        x = InStr(x, sTest, sLook) 'sets x to the position of the string we're looking for
'
'        If z = 10 Then
'            y1 = x + 1
'            GoTo nonumber:
'        End If
'
'        y1 = x + Len(sLook) 'len of string searching (to position y1 at first number)
'        y2 = 0
'nextnumber:
'        sChar = Mid(sTest, y1 + y2, 1)
'        Select Case sChar
'            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "/":
'                If Not y1 + y2 - 1 = Len(sTest) Then
'                    y2 = y2 + 1
'                    GoTo nextnumber:
'                End If
'            Case "+": 'end of string
'                Exit Function
'            Case "(": 'precent
'                x2 = InStr(y1 + y2, sTest, ")")
'                If Not x2 = 0 Then
'                    sSuffix = " " & Mid(sTest, y1 + y2, x2 - y1 - y2 + 1)
'                End If
'            Case Else:
'        End Select
'
'        If y2 = 0 Then
'            'if there were no numbers after the string
'            x = y1
'            GoTo checknext:
'        End If
'
'        If Not z = 1 Then 'not room or group
'            nValue = Val(Mid(sTest, y1, y2))
'        End If
'
'nonumber:
'        If Not GetLocations_STR = "" Then GetLocations_STR = GetLocations_STR & ", "
'
'        Select Case z
'            Case 1: '"room "
'                GetLocations_STR = GetLocations_STR & "Room: " & GetRoomName(Mid(sTest, y1, y2)) & sSuffix
'
'            Case 2: '"monster #"
'                GetLocations_STR = GetLocations_STR & "Monster: " & GetMonsterName(nValue, True) & sSuffix
'
'            Case 3: '"textblock #"
'                GetLocations_STR = GetLocations_STR & "Textblock " & nValue & sSuffix
'
'            Case 4: '"textblock(rndm) #"
'                GetLocations_STR = GetLocations_STR & "Textblock " & nValue & " (random)" & sSuffix
'
'            Case 5: '"item #"
'                GetLocations_STR = GetLocations_STR & "Item: " & GetItemName(nValue) & sSuffix
'
'            Case 6: '"spell #"
'                GetLocations_STR = GetLocations_STR & "Spell: " & GetSpellName(nValue) & sSuffix
'
'            Case 7: '"shop #"
'                GetLocations_STR = GetLocations_STR & "Shop: " & GetShopName(nValue) & sSuffix
'
'            Case 8: '"shop(sell) #"
'                GetLocations_STR = GetLocations_STR & "Shop: " & GetShopName(nValue) & " (sell only)" & sSuffix
'
'            Case 9: '"shop(nogen) #"
'                GetLocations_STR = GetLocations_STR & "Shop: " & GetShopName(nValue) & " (wont regen)" & sSuffix
'
'            Case 10: 'group
'                GetLocations_STR = GetLocations_STR & "Regular room regeneration via group."
'
'        End Select
'
'        x = y1
'        GoTo checknext:
'    End If
'Next z
'
'If GetLocations_STR = "" Then GetLocations_STR = "None."
'
'End Function

Public Sub SelectAll(ByRef TB As TextBox)
TB.SelStart = 0
TB.SelLength = Len(TB.Text)
End Sub

Public Function GetEquipCaption(ByVal nIndex As Integer) As String

Select Case nIndex
    Case 0: GetEquipCaption = "Head"
    Case 1: GetEquipCaption = "Ears"
    Case 2: GetEquipCaption = "Neck"
    Case 3: GetEquipCaption = "Back"
    Case 4: GetEquipCaption = "Torso"
    Case 5: GetEquipCaption = "Arms"
    Case 6: GetEquipCaption = "Wrist"
    Case 7: GetEquipCaption = "Wrist"
    Case 8: GetEquipCaption = "Hands"
    Case 9: GetEquipCaption = "Finger"
    Case 10: GetEquipCaption = "Finger"
    Case 11: GetEquipCaption = "Waist"
    Case 12: GetEquipCaption = "Legs"
    Case 13: GetEquipCaption = "Feet"
    Case 14: GetEquipCaption = "Worn"
    Case 15: GetEquipCaption = "Off-Hand"
    Case 16: GetEquipCaption = "Weapon Hand"
    Case 17: GetEquipCaption = "Wrist"
    Case 18: GetEquipCaption = "Eyes"
    Case 19: GetEquipCaption = "Face"
End Select

End Function

Public Sub AppReload(ByVal bNewSettings As Boolean)

frmMain.bDontCallTerminate = True
Call AppTerminate
DoEvents

bAppTerminating = False
If bNewSettings Then Call CreateSettings
DoEvents

Load frmMain
frmMain.bDontCallTerminate = False
End Sub

Public Sub AppTerminate()
bAppTerminating = True
bCancelTerminate = False

Call UnloadForms("none")
DoEvents

If bCancelTerminate Then
    bAppTerminating = False
    Exit Sub
End If

Call CloseDatabases
DoEvents

'End
End Sub

Public Function RegCreateKeyPath(ByVal enmHKEY As hkey, ByVal strKeyPath As String) As Integer
'****************************************************************************
' By Syntax53
' Create a path of keys
' Inputs: HKEY, KeyPath
' Return: Error code, 0=no error
'****************************************************************************
On Error GoTo Error:
Dim cReg As clsRegistryRoutines
Dim x As Long, y As Long, KeyArray() As String

Set cReg = New clsRegistryRoutines

cReg.hkey = enmHKEY

x = InStr(1, strKeyPath, "\")
If x = 0 Then
    cReg.KeyRoot = ""
    cReg.Subkey = strKeyPath
    If Not cReg.KeyExists Then Call cReg.CreateKey(strKeyPath)
    GoTo quit:
End If

KeyArray() = Split(strKeyPath, "\", , vbTextCompare)
'ReDim KeyArray(0)
'KeyArray(0) = Mid(strKeyPath, 1, x - 1)
'y = y + 1
'
'Do While InStr(x + 1, strKeyPath, "\") > 0
'    ReDim Preserve KeyArray(0 To y)
'    KeyArray(y) = Mid(strKeyPath, x + 1, InStr(x + 1, strKeyPath, "\") - x - 1)
'    x = InStr(x + 1, strKeyPath, "\")
'    y = y + 1
'Loop
'
'ReDim Preserve KeyArray(0 To y)
'KeyArray(y) = Mid(strKeyPath, x + 1)

For x = 0 To UBound(KeyArray())
    cReg.KeyRoot = ""
    For y = 0 To (x - 1)
        cReg.KeyRoot = cReg.KeyRoot & KeyArray(y)
        If Not y = (x - 1) Then cReg.KeyRoot = cReg.KeyRoot & "\"
    Next
    cReg.Subkey = KeyArray(x)
    If Not cReg.KeyExists Then Call cReg.CreateKey(cReg.Subkey)
Next x

quit:
Set cReg = Nothing
Erase KeyArray()
Exit Function

Error:
RegCreateKeyPath = Err.Number
Resume quit:
End Function

Public Function StringOfNumbersToArray(sNumberString As String) As String()
Dim x As Long, sRet() As String
On Error GoTo Error:

If InStr(1, sNumberString, ",", vbTextCompare) = 0 Then
    ReDim sRet(0)
    sRet(0) = Val(Replace(Replace(sNumberString, "(", "", 1, -1, vbTextCompare), ")", "", 1, -1, vbTextCompare))
Else
    sRet = Split(sNumberString, ",")
    For x = 0 To UBound(sRet())
        sRet(x) = Val(Replace(Replace(sRet(x), "(", "", 1, -1, vbTextCompare), ")", "", 1, -1, vbTextCompare))
    Next x
End If

StringOfNumbersToArray = sRet

out:
On Error Resume Next
Exit Function
Error:
Call HandleError("NumberStringToArray")
Resume out:
End Function

' Omit pvarMirror, plngLeft & plngRight; they are used internally during recursion
Public Sub MergeSort1(ByRef pvarArray As Variant, Optional pvarMirror As Variant, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngMid As Long
    Dim l As Long
    Dim R As Long
    Dim O As Long
    Dim varSwap As Variant
 
    If plngRight = 0 Then
        plngLeft = LBound(pvarArray)
        plngRight = UBound(pvarArray)
        ReDim pvarMirror(plngLeft To plngRight)
    End If
    lngMid = plngRight - plngLeft
    Select Case lngMid
        Case 0
        Case 1
            If pvarArray(plngLeft) > pvarArray(plngRight) Then
                varSwap = pvarArray(plngLeft)
                pvarArray(plngLeft) = pvarArray(plngRight)
                pvarArray(plngRight) = varSwap
            End If
        Case Else
            lngMid = lngMid \ 2 + plngLeft
            MergeSort1 pvarArray, pvarMirror, plngLeft, lngMid
            MergeSort1 pvarArray, pvarMirror, lngMid + 1, plngRight
            ' Merge the resulting halves
            l = plngLeft ' start of first (left) half
            R = lngMid + 1 ' start of second (right) half
            O = plngLeft ' start of output (mirror array)
            Do
                If pvarArray(R) < pvarArray(l) Then
                    pvarMirror(O) = pvarArray(R)
                    R = R + 1
                    If R > plngRight Then
                        For l = l To lngMid
                            O = O + 1
                            pvarMirror(O) = pvarArray(l)
                        Next
                        Exit Do
                    End If
                Else
                    pvarMirror(O) = pvarArray(l)
                    l = l + 1
                    If l > lngMid Then
                        For R = R To plngRight
                            O = O + 1
                            pvarMirror(O) = pvarArray(R)
                        Next
                        Exit Do
                    End If
                End If
                O = O + 1
            Loop
            For O = plngLeft To plngRight
                pvarArray(O) = pvarMirror(O)
            Next
    End Select
End Sub
