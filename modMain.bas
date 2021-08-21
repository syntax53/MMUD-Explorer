Attribute VB_Name = "modMain"
Option Explicit
Option Base 0

Global bHideRecordNumbers As Boolean
Global bOnlyInGame As Boolean
Global nNMRVer As Double
Global sCurrentDatabaseFile As String
'Global bOnlyLearnable As Boolean

Global nLastItemSortCol As Integer

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
Public nEquippedItem(0 To 19) As Long
Public nLearnedSpells(0 To 99) As Long
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
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Public Declare Function MoveWindow Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal x As Long, ByVal y As Long, _
   ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal bRepaint As Long) As Long

Public Declare Function GetWindowRect Lib "user32" _
  (ByVal hwnd As Long, _
   lpRect As RECT) As Long

Public Declare Function ScreenToClient Lib "user32" _
  (ByVal hwnd As Long, _
   lpPoint As POINTAPI) As Long

Private Declare Function SendMessageLong Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function DrawText Lib "user32" Alias _
    "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, _
    ByVal nCount As Long, lpRect As RECT, ByVal wFormat _
    As Long) As Long
    
Public Declare Function CalcExpNeeded Lib "lltmmudxp" (ByVal Level As Long, ByVal Chart As Long) As Currency

Public Sub LearnOrUnlearnSpell(nSpell As Long)
On Error GoTo error:

If in_long_arr(ByVal nSpell, nLearnedSpells()) Then
    Call UnLearnSpell(nSpell)
Else
    Call LearnSpell(nSpell)
End If

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("LearnOrUnlearnSpell")
Resume out:
End Sub

Public Sub LearnSpell(nSpell As Long)
Dim x As Integer
On Error GoTo error:

For x = 0 To 99
    If nLearnedSpells(x) = 0 Then
        nLearnedSpells(x) = nSpell
        Exit For
    End If
Next x

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("LearnSpell")
Resume out:
End Sub

Public Sub UnLearnSpell(nSpell As Long)
Dim x As Integer
On Error GoTo error:

For x = 0 To 99
    If nLearnedSpells(x) = nSpell Then
        nLearnedSpells(x) = 0
    End If
Next x

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("UnLearnSpell")
Resume out:
End Sub

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
        lRet = SendMessage(Combo.hwnd, CB_SETDROPPEDCONTROLRECT, lComboWidth, 0)
        
    End If
    
    If ExpandType <> WidthOnly Then
        lComboWidth = Combo.Width / Screen.TwipsPerPixelX
        lItemHeight = SendMessage(Combo.hwnd, CB_GETITEMHEIGHT, 0, 0)
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
        Call GetWindowRect(Combo.hwnd, rc)
        pt.x = rc.Left
        pt.y = rc.Top
        Call ScreenToClient(hFrame, pt)
        Call MoveWindow(Combo.hwnd, pt.x, pt.y, lComboWidth, lNewHeight, True)
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
 
lCurrentWidth = SendMessageLong(Combo.hwnd, CB_GETDROPPEDWIDTH, _
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

lRet = SendMessageLong(Combo.hwnd, CB_SETDROPPEDWIDTH, lWidth, 0)

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

Public Sub PullItemDetail(DetailTB As TextBox, LocationLV As ListView)
Dim sStr As String, sAbil As String, x As Integer, sCasts As String, nPercent As Integer
Dim sNegate As String, sClasses As String, sRaces As String, sClassOk As String
Dim sUses As String, sGetDrop As String, oLI As ListItem, nNumber As Long
Dim y As Integer, z As Integer, bCompareWeapon As Boolean, bCompareArmor As Boolean
Dim nInvenSlot1 As Integer, nInvenSlot2 As Integer
Dim sCompareText1 As String, sCompareText2 As String
Dim tabItems1 As Recordset, tabItems2 As Recordset
Dim sTemp1 As String, sTemp2 As String, sTemp3 As String, sArr() As String
Dim bFlag1 As Boolean, bFlag2 As Boolean, sOP As String
Dim nClassRestrictions(0 To 2, 0 To 9) As Long
Dim nRaceRestrictions(0 To 2, 0 To 9) As Long
Dim nNegateSpells(0 To 2, 0 To 9) As Long
Dim nAbils(0 To 2, 0 To 19, 0 To 2) As Long, sAbilText(0 To 2, 0 To 19) As String
Dim nReturnValue As Long, nMatchReturnValue As Long
Dim sClassOk1 As String, sClassOk2 As String
Dim sCastSp1 As String, sCastSp2 As String
Dim bCastSpFlag(0 To 2) As Boolean
Dim nPct(0 To 2) As Integer

nInvenSlot1 = -1
nInvenSlot2 = -1

'sStr = ClipNull(tabItems.Fields("Name")) & " (" & tabItems.Fields("Number") & ")"

On Error GoTo error:

nNumber = tabItems.Fields("Number")

Select Case tabItems.Fields("ItemType")
    Case 0: 'armour
        If tabItems.Fields("Worn") <= 0 Then
            'nada
        Else
            If tabItems.Fields("Worn") <= UBound(nEquippedItem) Then
                Select Case tabItems.Fields("Worn")
                    Case 0: '"Nowhere"
                    Case 1: '"Everywhere"
                        nInvenSlot1 = 19
                    Case 2: '"Head"
                        nInvenSlot1 = 0
                    Case 3: '"Hands"
                        nInvenSlot1 = 8
                    Case 4, 13: '"Finger"
                        If nEquippedItem(9) > 0 Then
                            nInvenSlot1 = 9
                            If nEquippedItem(10) > 0 Then nInvenSlot2 = 10
                        ElseIf nEquippedItem(10) > 0 Then
                            nInvenSlot1 = 10
                        End If
                    Case 5: '"Feet"
                        nInvenSlot1 = 13
                    Case 6: '"Arms"
                        nInvenSlot1 = 5
                    Case 7: '"Back"
                        nInvenSlot1 = 3
                    Case 8: '"Neck"
                        nInvenSlot1 = 2
                    Case 9: '"Legs"
                        nInvenSlot1 = 12
                    Case 10: '"Waist"
                        nInvenSlot1 = 11
                    Case 11: '"Torso"
                        nInvenSlot1 = 4
                    Case 12: '"Off-Hand"
                        nInvenSlot1 = 15
                    Case 14: '"Wrist"
                        If nEquippedItem(6) > 0 Then
                            nInvenSlot1 = 6
                            If nEquippedItem(7) > 0 Then nInvenSlot2 = 7
                        ElseIf nEquippedItem(7) > 0 Then
                            nInvenSlot1 = 7
                        End If
                    Case 15: '"Ears"
                        nInvenSlot1 = 1
                    Case 16: '"Worn"
                        nInvenSlot1 = 14
                    Case 18: '"Eyes"
                        nInvenSlot1 = 17
                    Case 19: '"Face"
                        nInvenSlot1 = 18
                    Case Else:
                End Select
                
                If nInvenSlot1 >= 0 Then
                    If nEquippedItem(nInvenSlot1) > 0 Then
                        bCompareArmor = True
                    Else
                        nInvenSlot1 = -1
                    End If
                End If
                
                If nInvenSlot2 >= 0 Then
                    If nEquippedItem(nInvenSlot2) > 0 Then
                        bCompareArmor = True
                    Else
                        nInvenSlot2 = -1
                    End If
                End If
                
            End If
        End If
        
    Case 1: 'weapons
        nInvenSlot1 = 16
        If nEquippedItem(nInvenSlot1) > 0 Then
            bCompareWeapon = True
        End If
        
    Case Else: 'other
        'nada
        
End Select

If Not bCompareWeapon And Not bCompareArmor Then
    nInvenSlot1 = -1
    nInvenSlot2 = -1
End If

If nInvenSlot1 >= 0 Then
    If nEquippedItem(nInvenSlot1) = nNumber Then nInvenSlot1 = -1
End If
If nInvenSlot2 >= 0 Then
    If nEquippedItem(nInvenSlot2) = nNumber Then nInvenSlot2 = -1
End If

If nInvenSlot2 >= 0 And nInvenSlot1 < 0 Then
    nInvenSlot1 = nInvenSlot2
    nInvenSlot2 = -1
End If

If nInvenSlot1 < 0 And nInvenSlot2 < 0 Then
    bCompareWeapon = False
    bCompareArmor = False
End If

If bCompareWeapon Or bCompareArmor And nInvenSlot1 >= 0 Then
    Set tabItems1 = DB.OpenRecordset("Items")
    tabItems1.Index = "pkItems"
    tabItems1.Seek "=", nEquippedItem(nInvenSlot1)
    If tabItems1.NoMatch = True Then
        If nInvenSlot2 < 0 Then
            bCompareWeapon = False
            bCompareArmor = False
            nInvenSlot1 = -1
            nInvenSlot2 = -1
            tabItems1.Close
            Set tabItems1 = Nothing
        End If
    End If
    
    If nInvenSlot2 >= 0 Then
        Set tabItems2 = DB.OpenRecordset("Items")
        tabItems2.Index = "pkItems"
        tabItems2.Seek "=", nEquippedItem(nInvenSlot2)
        If tabItems2.NoMatch = True Then
            If nInvenSlot1 < 0 Then
                bCompareWeapon = False
                bCompareArmor = False
                nInvenSlot1 = -1
                nInvenSlot2 = -1
                tabItems2.Close
                Set tabItems2 = Nothing
            End If
        End If
    End If
End If

'#################
If tabItems.Fields("UseCount") > 0 Then
    sUses = tabItems.Fields("UseCount")
    If tabItems.Fields("Retain After Uses") = 1 Then
        sUses = sUses & " start/max"
    Else
        sUses = sUses & " (destroys after uses)"
    End If
End If

If nInvenSlot1 >= 0 Then
    If tabItems1.Fields("UseCount") > 0 Then
        sTemp1 = tabItems1.Fields("UseCount")
        If tabItems1.Fields("Retain After Uses") = 1 Then
            sTemp1 = "Uses: " & sTemp1 & " start/max"
        Else
            sTemp1 = "Uses: " & sTemp1 & " (destroys after uses)"
        End If
    End If
    If sTemp1 <> sUses Then sCompareText1 = AutoAppend(sCompareText1, sTemp1)
End If

If nInvenSlot2 >= 0 Then
    If tabItems2.Fields("UseCount") > 0 Then
        sTemp2 = tabItems2.Fields("UseCount")
        If tabItems2.Fields("Retain After Uses") = 1 Then
            sTemp2 = "Uses: " & sTemp2 & " start/max"
        Else
            sTemp2 = "Uses: " & sTemp2 & " (destroys after uses)"
        End If
    End If
    If sTemp2 <> sUses Then sCompareText2 = AutoAppend(sCompareText2, sTemp2)
End If


'#################
If tabItems.Fields("Gettable") = 0 Then sGetDrop = AutoAppend(sGetDrop, "Not Getable")
If nInvenSlot1 >= 0 Then
    If tabItems.Fields("Gettable") <> tabItems1.Fields("Gettable") Then
        If tabItems1.Fields("Gettable") = 0 Then
            sCompareText1 = AutoAppend(sCompareText1, "+Getable")
        Else
            sCompareText1 = AutoAppend(sCompareText1, "+Not Getable")
        End If
    End If
End If
If nInvenSlot2 >= 0 Then
    If tabItems.Fields("Gettable") <> tabItems2.Fields("Gettable") Then
        If tabItems2.Fields("Gettable") = 0 Then
            sCompareText2 = AutoAppend(sCompareText2, "+Getable")
        Else
            sCompareText2 = AutoAppend(sCompareText2, "+Not Getable")
        End If
    End If
End If

'#################
If tabItems.Fields("Not Droppable") = 1 Then sGetDrop = AutoAppend(sGetDrop, "Not Droppable")
If nInvenSlot1 >= 0 Then
    If tabItems.Fields("Not Droppable") <> tabItems1.Fields("Not Droppable") Then
        If tabItems1.Fields("Not Droppable") = 1 Then
            sCompareText1 = AutoAppend(sCompareText1, "+Droppable")
        Else
            sCompareText1 = AutoAppend(sCompareText1, "+Not Droppable")
        End If
    End If
End If
If nInvenSlot2 >= 0 Then
    If tabItems.Fields("Not Droppable") <> tabItems2.Fields("Not Droppable") Then
        If tabItems2.Fields("Not Droppable") = 1 Then
            sCompareText2 = AutoAppend(sCompareText2, "+Droppable")
        Else
            sCompareText2 = AutoAppend(sCompareText2, "+Not Droppable")
        End If
    End If
End If

'#################
If tabItems.Fields("Destroy On Death") = 1 Then sGetDrop = AutoAppend(sGetDrop, "Destroys On Death")
If nInvenSlot1 >= 0 Then
    If tabItems.Fields("Destroy On Death") <> tabItems1.Fields("Destroy On Death") Then
        If tabItems1.Fields("Destroy On Death") = 1 Then
            sCompareText1 = AutoAppend(sCompareText1, "-Destroy On Death")
        Else
            sCompareText1 = AutoAppend(sCompareText1, "+Destroy On Death")
        End If
    End If
End If
If nInvenSlot2 >= 0 Then
    If tabItems.Fields("Destroy On Death") <> tabItems2.Fields("Destroy On Death") Then
        If tabItems2.Fields("Destroy On Death") = 1 Then
            sCompareText2 = AutoAppend(sCompareText2, "-Destroy On Death")
        Else
            sCompareText2 = AutoAppend(sCompareText2, "+Destroy On Death")
        End If
    End If
End If

'#################

For x = 0 To 19
    
    If tabItems.Fields("Abil-" & x) > 0 Then
        nAbils(0, x, 0) = tabItems.Fields("Abil-" & x)
        'If getindex_array_long_3d(nAbils, nAbils(0, x, 0), nReturnValue, 2, 0, , , x, 0) Then
        '    nAbils(0, nReturnValue, 1) = nAbils(0, nReturnValue, 1) + tabItems.Fields("AbilVal-" & x)
        '    nAbils(0, x, 0) = 0
        'Else
            nAbils(0, x, 1) = tabItems.Fields("AbilVal-" & x)
        'End If
    End If
    
    If nInvenSlot1 >= 0 Then
        If tabItems1.Fields("Abil-" & x) > 0 Then
            nAbils(1, x, 0) = tabItems1.Fields("Abil-" & x)
            'If getindex_array_long_3d(nAbils, nAbils(1, x, 0), nReturnValue, 2, 1, , , x, 0) Then
            '    nAbils(1, nReturnValue, 1) = nAbils(1, nReturnValue, 1) + tabItems1.Fields("AbilVal-" & x)
            '    nAbils(1, x, 0) = 0
            'Else
                nAbils(1, x, 1) = tabItems1.Fields("AbilVal-" & x)
            'End If
        End If
    End If
    
    If nInvenSlot2 >= 0 Then
        If tabItems2.Fields("Abil-" & x) > 0 Then
            nAbils(2, x, 0) = tabItems2.Fields("Abil-" & x)
            'If getindex_array_long_3d(nAbils, nAbils(2, x, 0), nReturnValue, 2, 2, , , x, 0) Then
            '    nAbils(2, nReturnValue, 1) = nAbils(2, nReturnValue, 1) + tabItems2.Fields("AbilVal-" & x)
            '    nAbils(2, x, 0) = 0
            'Else
                nAbils(2, x, 1) = tabItems2.Fields("AbilVal-" & x)
            'End If
        End If
    End If
Next

For x = 0 To 19
    If nAbils(0, x, 0) > 0 Then
        Select Case nAbils(0, x, 0)
            Case 116: '116-bsacc
                If Not DetailTB.name = "txtWeaponCompareDetail" And _
                    Not DetailTB.name = "txtWeaponDetail" Then
                    
                    sTemp1 = GetAbilityStats(nAbils(0, x, 0), nAbils(0, x, 1), LocationLV, , True)
                    sAbilText(0, x) = sTemp1
                    sAbil = AutoAppend(sAbil, sTemp1)
                End If
                
            Case 22, 105, 106, 135:  '22-acc, 105-acc, 106-acc, 135-minlvl
                If Not DetailTB.name = "txtWeaponCompareDetail" And _
                    Not DetailTB.name = "txtWeaponDetail" And _
                    Not DetailTB.name = "txtArmourCompareDetail" And _
                    Not DetailTB.name = "txtArmourDetail" Then
    
                    sTemp1 = GetAbilityStats(nAbils(0, x, 0), nAbils(0, x, 1), LocationLV, , True)
                    sAbilText(0, x) = sTemp1
                    sAbil = AutoAppend(sAbil, sTemp1)
                End If
                
            Case 59: 'class ok
                sTemp1 = GetClassName(nAbils(0, x, 1))
                sAbilText(0, x) = sTemp1
                sClassOk = AutoAppend(sClassOk, sTemp1)
                
            Case 43: 'casts spell
                'nSpellNest = 0 'make sure this doesn't nest too deep
                sCasts = AutoAppend(sCasts, "[" & GetSpellName(nAbils(0, x, 1), bHideRecordNumbers) _
                    & ", " & PullSpellEQ(True, 0, nAbils(0, x, 1), LocationLV, , , True))
                If Not nPercent = 0 Then
                    sCasts = sCasts & ", " & nPercent & "%]"
                Else
                    sCasts = sCasts & "]"
                End If
                sAbilText(0, x) = sCasts
                
                Set oLI = LocationLV.ListItems.Add
                oLI.Text = ""
                oLI.ListSubItems.Add 1, , "Casts: " & GetSpellName(nAbils(0, x, 1), bHideRecordNumbers)
                oLI.ListSubItems(1).Tag = nAbils(0, x, 1)
                
            Case 114: '%spell
                nPercent = nAbils(0, x, 1)
                
            Case Else:
                sTemp1 = GetAbilityStats(nAbils(0, x, 0), nAbils(0, x, 1), LocationLV, , True)
                sAbilText(0, x) = sTemp1
                sAbil = AutoAppend(sAbil, sTemp1)
                
        End Select
    End If
Next x

If nInvenSlot1 >= 0 Then
    sTemp1 = ""
    sTemp2 = ""
    bCastSpFlag(0) = False
    bCastSpFlag(1) = False
    bCastSpFlag(2) = False
    bFlag1 = False
    bFlag2 = False
    For y = 0 To 2
        
        nPct(0) = 0
        nPct(1) = 0
        nPct(2) = 0
        For x = 0 To 19
            
            nMatchReturnValue = -32000
            
            If nAbils(y, x, 0) > 0 Then
            
                If nAbils(y, x, 0) = 59 Then nMatchReturnValue = nAbils(y, x, 1) 'classok
                If nAbils(y, x, 0) = 43 Then
                    nMatchReturnValue = nAbils(y, x, 1) 'casts spell
                    bCastSpFlag(y) = True
                End If
                If nAbils(y, x, 0) = 114 Then nPct(y) = nAbils(y, x, 1) '%spell
                
                If nAbils(y, x, 0) = 117 Then
                    Debug.Print 1
                End If
                
                If y = 0 Then
                    
                    If Not getval_array_long_3d(nAbils, nAbils(y, x, 0), nReturnValue, 1, nMatchReturnValue, 1, , , , 0) Then
                        
                        sTemp3 = GetAbilDiffText(nAbils(y, x, 0), nAbils(y, x, 1), 0, sAbilText(y, x), nPct(y))
                        If Len(sTemp3) > 0 Then
                            If nAbils(y, x, 0) = 59 Then 'classok
                                sClassOk1 = AutoAppend(sClassOk1, "+" & sTemp3)
                            ElseIf nAbils(y, x, 0) = 43 Then 'casts spell
                                sCastSp1 = AutoAppend(sCastSp1, "+" & sTemp3)
                            Else
                                bFlag1 = True
                                sTemp1 = AutoAppend(sTemp1, IIf(nAbils(y, x, 1) = 0, "+", "") & sTemp3)
                            End If
                        End If
                        
                    ElseIf nReturnValue <> nAbils(y, x, 1) Then
                        
                        sTemp3 = GetAbilDiffText(nAbils(y, x, 0), nAbils(y, x, 1), nReturnValue, sAbilText(y, x), nPct(y))
                        If Len(sTemp3) > 0 Then
                            bFlag1 = True
                            sTemp1 = AutoAppend(sTemp1, sTemp3)
                        End If
                        
                    End If
                    
                    If nInvenSlot2 >= 0 Then
                    
                        If Not getval_array_long_3d(nAbils, nAbils(y, x, 0), nReturnValue, 1, nMatchReturnValue, 2, , , , 0) Then
                        
                            sTemp3 = GetAbilDiffText(nAbils(y, x, 0), nAbils(y, x, 1), 0, sAbilText(y, x), nPct(y))
                            If Len(sTemp3) > 0 Then
                                If nAbils(y, x, 0) = 59 Then 'classok
                                    sClassOk2 = AutoAppend(sClassOk2, "+" & sTemp3)
                                ElseIf nAbils(y, x, 0) = 43 Then 'casts spell
                                    sCastSp2 = AutoAppend(sCastSp2, "+" & sTemp3)
                                Else
                                    bFlag2 = True
                                    sTemp2 = AutoAppend(sTemp2, IIf(nAbils(y, x, 1) = 0, "+", "") & sTemp3)
                                End If
                            End If
                            
                        ElseIf nReturnValue <> nAbils(y, x, 1) Then
                        
                            sTemp3 = GetAbilDiffText(nAbils(y, x, 0), nAbils(y, x, 1), nReturnValue, sAbilText(y, x), nPct(y))
                            If Len(sTemp3) > 0 Then
                                bFlag2 = True
                                sTemp2 = AutoAppend(sTemp2, sTemp3)
                            End If
                            
                        End If
                        
                    End If
                Else
                    If Not getval_array_long_3d(nAbils, nAbils(y, x, 0), nReturnValue, 1, nMatchReturnValue, 0, , , , 0) Then
                        
                        sTemp3 = GetAbilDiffText(nAbils(y, x, 0), nAbils(y, x, 1), 0, , nPct(y), True)
                        If Len(sTemp3) > 0 Then
                            If nAbils(y, x, 0) = 59 Then 'classok
                                If y = 1 Then
                                    sClassOk1 = AutoAppend(sClassOk1, "-" & sTemp3)
                                Else
                                    sClassOk2 = AutoAppend(sClassOk2, "-" & sTemp3)
                                End If
                            ElseIf nAbils(y, x, 0) = 43 Then 'casts spell
                                If y = 1 Then
                                    sCastSp1 = AutoAppend(sCastSp1, "-" & sTemp3)
                                Else
                                    sCastSp2 = AutoAppend(sCastSp2, "-" & sTemp3)
                                End If
                            Else
                                If y = 1 Then
                                    bFlag1 = True
                                    sTemp1 = AutoAppend(sTemp1, IIf(nAbils(y, x, 1) = 0, "-", "") & sTemp3)
                                Else
                                    bFlag2 = True
                                    sTemp2 = AutoAppend(sTemp2, IIf(nAbils(y, x, 1) = 0, "-", "") & sTemp3)
                                End If
                            End If
                        End If
                        
                    End If
                End If
            End If
        Next x
    Next y
    
    If Len(sTemp1) > 0 Then
        sTemp3 = "Abilities"
        If Len(sAbil) = 0 And bFlag1 Then
            sTemp3 = sTemp3 & " [none]: " & sTemp1
        Else
            sTemp3 = sTemp3 & ": " & sTemp1
        End If
        sCompareText1 = AutoAppend(sCompareText1, sTemp3, " -- ")
    End If

    If Len(sTemp2) > 0 Then
        sTemp3 = "Abilities"
        If Len(sAbil) = 0 And bFlag2 Then
            sTemp3 = sTemp3 & " [none]: " & sTemp2
        Else
            sTemp3 = sTemp3 & ": " & sTemp2
        End If
        sCompareText2 = AutoAppend(sCompareText2, sTemp3, " -- ")
    End If
    
    If Len(sCastSp1) > 0 Then
        sTemp3 = "Casts"
        If Len(sCasts) = 0 And bCastSpFlag(1) Then
            sTemp3 = sTemp3 & " [none]: " & sCastSp1
        Else
            sTemp3 = sTemp3 & ": " & sCastSp1
        End If
        sCompareText1 = AutoAppend(sCompareText1, sTemp3, " -- ")
    End If

    If Len(sCastSp2) > 0 Then
        sTemp3 = "Casts"
        If Len(sCasts) = 0 And bCastSpFlag(2) Then
            sTemp3 = sTemp3 & " [none]: " & sCastSp2
        Else
            sTemp3 = sTemp3 & ": " & sCastSp2
        End If
        sCompareText2 = AutoAppend(sCompareText2, sTemp3, " -- ")
    End If
    
    If Len(sClassOk1) > 0 Then
        sTemp3 = "ClassOk"
        sTemp3 = sTemp3 & ": " & sClassOk1
        sCompareText1 = AutoAppend(sCompareText1, sTemp3, " -- ")
    End If

    If Len(sClassOk2) > 0 Then
        sTemp3 = "ClassOk"
        sTemp3 = sTemp3 & ": " & sClassOk2
        sCompareText2 = AutoAppend(sCompareText2, sTemp3, " -- ")
    End If
End If

'#################

For x = 0 To 9
    If tabItems.Fields("ClassRest-" & x) <> 0 Then
        sClasses = AutoAppend(sClasses, GetClassName(tabItems.Fields("ClassRest-" & x)))
        nClassRestrictions(0, x) = tabItems.Fields("ClassRest-" & x)
    End If
    If nInvenSlot1 >= 0 Then
        If tabItems1.Fields("ClassRest-" & x) <> 0 Then nClassRestrictions(1, x) = tabItems1.Fields("ClassRest-" & x)
    End If
    If nInvenSlot2 >= 0 Then
        If tabItems2.Fields("ClassRest-" & x) <> 0 Then nClassRestrictions(2, x) = tabItems2.Fields("ClassRest-" & x)
    End If
Next

If nInvenSlot1 >= 0 Then
    sTemp1 = ""
    sTemp2 = ""
    bFlag1 = False
    bFlag2 = False
    For y = 0 To 2
        For x = 0 To 9
            If nClassRestrictions(y, x) > 0 Then
                If y = 0 Then
                    If Not in_array_long_md(nClassRestrictions, nClassRestrictions(y, x), 1) Then
                        sTemp1 = AutoAppend(sTemp1, "+" & GetClassName(nClassRestrictions(y, x)))
                    End If
                    If nInvenSlot2 >= 0 Then
                        If Not in_array_long_md(nClassRestrictions, nClassRestrictions(y, x), 2) Then
                            sTemp2 = AutoAppend(sTemp2, "+" & GetClassName(nClassRestrictions(y, x)))
                        End If
                    End If
                Else
                    If y = 1 Then 'mark that there are values found
                        bFlag1 = True
                    Else
                        bFlag2 = True
                    End If
                    If Not in_array_long_md(nClassRestrictions, nClassRestrictions(y, x), 0) Then
                        If y = 1 Then
                            sTemp1 = AutoAppend(sTemp1, "-" & GetClassName(nClassRestrictions(y, x)))
                        Else
                            sTemp2 = AutoAppend(sTemp2, "-" & GetClassName(nClassRestrictions(y, x)))
                        End If
                    End If
                End If
            End If
        Next x
    Next y
    
    If Len(sTemp1) > 0 Then
        sTemp3 = "Classes"
        If Len(sClasses) = 0 Then
            sTemp3 = sTemp3 & " [not-restricted]"
        ElseIf Not bFlag1 Then
            sTemp3 = sTemp3 & " [+restricted]" & ": " & sTemp1
        Else
            sTemp3 = sTemp3 & ": " & sTemp1
        End If
        sCompareText1 = AutoAppend(sCompareText1, sTemp3, " -- ")
    End If

    If Len(sTemp2) > 0 Then
        sTemp3 = "Classes"
        If Len(sClasses) = 0 Then
            sTemp3 = sTemp3 & " [not-restricted]"
        ElseIf Not bFlag2 Then
            sTemp3 = sTemp3 & " [+restricted]" & ": " & sTemp2
        Else
            sTemp3 = sTemp3 & ": " & sTemp2
        End If
        sCompareText2 = AutoAppend(sCompareText2, sTemp3, " -- ")
    End If
End If


'#################

For x = 0 To 9
    If tabItems.Fields("RaceRest-" & x) <> 0 Then
        sRaces = AutoAppend(sRaces, GetRaceName(tabItems.Fields("RaceRest-" & x)))
        nRaceRestrictions(0, x) = tabItems.Fields("RaceRest-" & x)
    End If
    If nInvenSlot1 >= 0 Then
        If tabItems1.Fields("RaceRest-" & x) <> 0 Then nRaceRestrictions(1, x) = tabItems1.Fields("RaceRest-" & x)
    End If
    If nInvenSlot2 >= 0 Then
        If tabItems2.Fields("RaceRest-" & x) <> 0 Then nRaceRestrictions(2, x) = tabItems2.Fields("RaceRest-" & x)
    End If
Next

If nInvenSlot1 >= 0 Then
    sTemp1 = ""
    sTemp2 = ""
    bFlag1 = False
    bFlag2 = False
    For y = 0 To 2
        For x = 0 To 9
            If nRaceRestrictions(y, x) > 0 Then
                If y = 0 Then
                    If Not in_array_long_md(nRaceRestrictions, nRaceRestrictions(y, x), 1) Then
                        sTemp1 = AutoAppend(sTemp1, "+" & GetRaceName(nRaceRestrictions(y, x)))
                    End If
                    If nInvenSlot2 >= 0 Then
                        If Not in_array_long_md(nRaceRestrictions, nRaceRestrictions(y, x), 2) Then
                            sTemp2 = AutoAppend(sTemp2, "+" & GetRaceName(nRaceRestrictions(y, x)))
                        End If
                    End If
                Else
                    If y = 1 Then 'mark that there are values found
                        bFlag1 = True
                    Else
                        bFlag2 = True
                    End If
                    If Not in_array_long_md(nRaceRestrictions, nRaceRestrictions(y, x), 0) Then
                        If y = 1 Then
                            sTemp1 = AutoAppend(sTemp1, "-" & GetRaceName(nRaceRestrictions(y, x)))
                        Else
                            sTemp2 = AutoAppend(sTemp2, "-" & GetRaceName(nRaceRestrictions(y, x)))
                        End If
                    End If
                End If
            End If
        Next x
    Next y
    
    If Len(sTemp1) > 0 Then
        sTemp3 = "Races"
        If Len(sRaces) = 0 Then
            sTemp3 = sTemp3 & " [not-restricted]"
        ElseIf Not bFlag1 Then
            sTemp3 = sTemp3 & " [+restricted]" & ": " & sTemp1
        Else
            sTemp3 = sTemp3 & ": " & sTemp1
        End If
        sCompareText1 = AutoAppend(sCompareText1, sTemp3, " -- ")
    End If

    If Len(sTemp2) > 0 Then
        sTemp3 = "Races"
        If Len(sRaces) = 0 Then
            sTemp3 = sTemp3 & " [not-restricted]"
        ElseIf Not bFlag2 Then
            sTemp3 = sTemp3 & " [+restricted]" & ": " & sTemp2
        Else
            sTemp3 = sTemp3 & ": " & sTemp2
        End If
        sCompareText2 = AutoAppend(sCompareText2, sTemp3, " -- ")
    End If
End If

'#################

For x = 0 To 9
    If tabItems.Fields("NegateSpell-" & x) <> 0 Then
        sNegate = AutoAppend(sNegate, GetSpellName(tabItems.Fields("NegateSpell-" & x)))
        nNegateSpells(0, x) = tabItems.Fields("NegateSpell-" & x)
    End If
    If nInvenSlot1 >= 0 Then
        If tabItems1.Fields("NegateSpell-" & x) <> 0 Then nNegateSpells(1, x) = tabItems1.Fields("NegateSpell-" & x)
    End If
    If nInvenSlot2 >= 0 Then
        If tabItems2.Fields("NegateSpell-" & x) <> 0 Then nNegateSpells(2, x) = tabItems2.Fields("NegateSpell-" & x)
    End If
Next

If nInvenSlot1 >= 0 Then
    sTemp1 = ""
    sTemp2 = ""
    bFlag1 = False
    bFlag2 = False
    For y = 0 To 2
        For x = 0 To 9
            If nNegateSpells(y, x) > 0 Then
                If y = 0 Then
                    If Not in_array_long_md(nNegateSpells, nNegateSpells(y, x), 1) Then
                        sTemp1 = AutoAppend(sTemp1, "+" & GetSpellName(nNegateSpells(y, x)))
                    End If
                    If nInvenSlot2 >= 0 Then
                        If Not in_array_long_md(nNegateSpells, nNegateSpells(y, x), 2) Then
                            sTemp2 = AutoAppend(sTemp2, "+" & GetSpellName(nNegateSpells(y, x)))
                        End If
                    End If
                Else
                    If y = 1 Then 'mark that there are values found
                        bFlag1 = True
                    Else
                        bFlag2 = True
                    End If
                    If Not in_array_long_md(nNegateSpells, nNegateSpells(y, x), 0) Then
                        If y = 1 Then
                            sTemp1 = AutoAppend(sTemp1, "-" & GetSpellName(nNegateSpells(y, x)))
                        Else
                            sTemp2 = AutoAppend(sTemp2, "-" & GetSpellName(nNegateSpells(y, x)))
                        End If
                    End If
                End If
            End If
        Next x
    Next y
    
    If Len(sTemp1) > 0 Then
        sTemp3 = "Negate"
        If Len(sNegate) = 0 And bFlag1 Then
            sTemp3 = sTemp3 & " [none]: " & sTemp1
        Else
            sTemp3 = sTemp3 & ": " & sTemp1
        End If
        sCompareText1 = AutoAppend(sCompareText1, sTemp3, " -- ")
    End If

    If Len(sTemp2) > 0 Then
        sTemp3 = "Negate"
        If Len(sNegate) = 0 And bFlag2 Then
            sTemp3 = sTemp3 & " [none]: " & sTemp2
        Else
            sTemp3 = sTemp3 & ": " & sTemp2
        End If
        sCompareText2 = AutoAppend(sCompareText2, sTemp3, " -- ")
    End If
    
End If

'#################

Call GetLocations(tabItems.Fields("Obtained From"), LocationLV, , , nNumber, , , True)
If nNMRVer >= 1.7 Then Call GetLocations(tabItems.Fields("References"), LocationLV, True, , , , , True)

If Not tabItems.Fields("Number") = nNumber Then
    tabItems.Index = "pkItems"
    tabItems.Seek "=", nNumber
End If

'####################

If Not sGetDrop = "" Then
    sStr = AutoAppend(sStr, sGetDrop, " -- ")
End If
If Not sUses = "" Then
    sStr = AutoAppend(sStr, "Uses: " & sUses, " -- ")
End If
If Not sAbil = "" Then
    sStr = AutoAppend(sStr, "Abilities: " & sAbil, " -- ")
End If
If Not sCasts = "" Then
    sStr = AutoAppend(sStr, "Casts: " & sCasts, " -- ")
End If
If Not sClassOk = "" Then
    sStr = AutoAppend(sStr, "ClassOK: " & sClassOk, " -- ")
End If
If Not sClasses = "" Then
    sStr = AutoAppend(sStr, "Classes: " & sClasses, " -- ")
End If
If Not sRaces = "" Then
    sStr = AutoAppend(sStr, "Races: " & sRaces, " -- ")
End If
If Not sNegate = "" Then
    sStr = AutoAppend(sStr, "Negates: " & sNegate, " -- ")
End If

If Len(sCompareText1) > 0 Then
    If bCompareWeapon Then sCompareText1 = AutoPrepend(sCompareText1, CompareWeapons(tabItems, tabItems1))
    If bCompareArmor Then sCompareText1 = AutoPrepend(sCompareText1, CompareArmor(tabItems, tabItems1))
    sStr = AutoAppend(sStr, "Compared to " & tabItems1.Fields("Name") & ": " & sCompareText1, vbCrLf & vbCrLf)
End If
If Len(sCompareText2) > 0 Then
    If bCompareWeapon Then sCompareText2 = AutoPrepend(sCompareText2, CompareWeapons(tabItems, tabItems2))
    If bCompareArmor Then sCompareText2 = AutoPrepend(sCompareText2, CompareArmor(tabItems, tabItems2))
    sStr = AutoAppend(sStr, "Compared to " & tabItems2.Fields("Name") & ": " & sCompareText2, vbCrLf & vbCrLf)
End If

DetailTB.Text = sStr

If LocationLV.ListItems.Count > 0 Then
    If nLastItemSortCol > LocationLV.ColumnHeaders.Count Then nLastItemSortCol = 1
    If nLastItemSortCol = 1 Then
        Call SortListViewByTag(LocationLV, 1, ldtnumber, False)
    Else
        Call SortListView(LocationLV, nLastItemSortCol, ldtstring, False)
    End If
End If

out:
On Error Resume Next
Set oLI = Nothing
tabItems2.Close
Set tabItems2 = Nothing
tabItems1.Close
Set tabItems1 = Nothing
Exit Sub

error:
Call HandleError("PullItemDetail")
Resume out:
End Sub

Private Function CompareArmor(tabItem1 As Recordset, tabItem2 As Recordset) As String
Dim nTemp As Long, nTemp2 As Double
On Error GoTo error:

If tabItem1.Fields("Worn") <> tabItem2.Fields("Worn") Then
    CompareArmor = AutoAppend(CompareArmor, GetWornType(tabItem2.Fields("Worn")) & " -> " & GetWornType(tabItem1.Fields("Worn")))
End If

If tabItem1.Fields("ArmourType") <> tabItem2.Fields("ArmourType") Then
    CompareArmor = AutoAppend(CompareArmor, GetArmourType(tabItem2.Fields("ArmourType")) & " -> " & GetArmourType(tabItem1.Fields("ArmourType")))
End If

If tabItem1.Fields("Encum") <> tabItem2.Fields("Encum") Then
    nTemp = tabItem1.Fields("Encum") - tabItem2.Fields("Encum")
    If nTemp <> 0 Then CompareArmor = AutoAppend(CompareArmor, "Encum: " & IIf(nTemp > 0, "+", "") & nTemp)
End If

If tabItem1.Fields("Accy") <> tabItem2.Fields("Accy") Then
    nTemp = tabItem1.Fields("Accy") - tabItem2.Fields("Accy")
    If nTemp <> 0 Then CompareArmor = AutoAppend(CompareArmor, "WepAccy: " & IIf(nTemp > 0, "+", "") & nTemp)
End If

If tabItem1.Fields("ArmourClass") <> tabItem2.Fields("ArmourClass") Or tabItem1.Fields("DamageResist") <> tabItem2.Fields("DamageResist") Then
    nTemp = RoundUp(tabItem1.Fields("ArmourClass") / 10) - RoundUp(tabItem2.Fields("ArmourClass") / 10)
    nTemp2 = (tabItem1.Fields("DamageResist") / 10) - (tabItem2.Fields("DamageResist") / 10)
    If nTemp <> 0 Or nTemp2 <> 0 Then CompareArmor = AutoAppend(CompareArmor, "AC: " & IIf(nTemp > 0, "+", "") & nTemp & "/" & IIf(nTemp2 > 0, "+", "") & nTemp2)
End If

'If tabItem1.Fields("Encum") <> tabItem2.Fields("Encum") Or tabItem1.Fields("ArmourClass") <> tabItem2.Fields("ArmourClass") Or tabItem1.Fields("DamageResist") <> tabItem2.Fields("DamageResist") Then
'    nTemp = Get_Enc_Ratio(tabItem1.Fields("Encum"), tabItem1.Fields("ArmourClass"), tabItem1.Fields("DamageResist")) - Get_Enc_Ratio(tabItem2.Fields("Encum"), tabItem2.Fields("ArmourClass"), tabItem2.Fields("DamageResist"))
'    If nTemp <> 0 Then CompareArmor = AutoAppend(CompareArmor, "AC/Enc: " & IIf(nTemp > 0, "+", "") & nTemp)
'End If

If tabItem1.Fields("Limit") <> tabItem2.Fields("Limit") Then
    nTemp = tabItem1.Fields("Limit") - tabItem2.Fields("Limit")
    If nTemp <> 0 Then CompareArmor = AutoAppend(CompareArmor, "Limit: " & IIf(nTemp > 0, "+", "") & nTemp)
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CompareArmor")
Resume out:
End Function


Private Function CompareWeapons(tabItem1 As Recordset, tabItem2 As Recordset) As String
Dim nTemp As Long, nTemp2 As Double
On Error GoTo error:

If tabItem1.Fields("WeaponType") <> tabItem2.Fields("WeaponType") Then
    CompareWeapons = AutoAppend(CompareWeapons, GetWeaponType(tabItem2.Fields("WeaponType")) & " -> " & GetWeaponType(tabItem1.Fields("WeaponType")))
End If

If tabItem1.Fields("Min") <> tabItem2.Fields("Min") Or tabItem1.Fields("Max") <> tabItem2.Fields("Max") Then
    nTemp = Round(tabItem1.Fields("Min") + tabItem1.Fields("Max") / 2, 0) - Round(tabItem2.Fields("Min") + tabItem2.Fields("Max") / 2, 0)
    If nTemp <> 0 Then CompareWeapons = AutoAppend(CompareWeapons, "Damage: " & IIf(nTemp > 0, "+", "") & nTemp)
End If

If tabItem1.Fields("Speed") <> tabItem2.Fields("Speed") Then
    nTemp = tabItem1.Fields("Speed") - tabItem2.Fields("Speed")
    If nTemp <> 0 Then CompareWeapons = AutoAppend(CompareWeapons, "Speed: " & IIf(nTemp > 0, "+", "") & nTemp)
End If

If tabItem1.Fields("Encum") <> tabItem2.Fields("Encum") Then
    nTemp = tabItem1.Fields("Encum") - tabItem2.Fields("Encum")
    If nTemp <> 0 Then CompareWeapons = AutoAppend(CompareWeapons, "Encum: " & IIf(nTemp > 0, "+", "") & nTemp)
End If

If tabItem1.Fields("Accy") <> tabItem2.Fields("Accy") Then
    nTemp = tabItem1.Fields("Accy") - tabItem2.Fields("Accy")
    If nTemp <> 0 Then CompareWeapons = AutoAppend(CompareWeapons, "WepAccy: " & IIf(nTemp > 0, "+", "") & nTemp)
End If

If tabItem1.Fields("ArmourClass") <> tabItem2.Fields("ArmourClass") Or tabItem1.Fields("DamageResist") <> tabItem2.Fields("DamageResist") Then
    nTemp = RoundUp(tabItem1.Fields("ArmourClass") / 10) - RoundUp(tabItem2.Fields("ArmourClass") / 10)
    nTemp2 = (tabItem1.Fields("DamageResist") / 10) - (tabItem2.Fields("DamageResist") / 10)
    If nTemp <> 0 Or nTemp2 <> 0 Then CompareWeapons = AutoAppend(CompareWeapons, "AC: " & IIf(nTemp > 0, "+", "") & nTemp & "/" & IIf(nTemp2 > 0, "+", "") & nTemp2)
End If

If tabItem1.Fields("StrReq") <> tabItem2.Fields("StrReq") Then
    nTemp = tabItem1.Fields("StrReq") - tabItem2.Fields("StrReq")
    If nTemp <> 0 Then CompareWeapons = AutoAppend(CompareWeapons, "StrReq: " & IIf(nTemp > 0, "+", "") & nTemp & " (" & tabItem1.Fields("StrReq") & ")")
End If

If tabItem1.Fields("Limit") <> tabItem2.Fields("Limit") Then
    nTemp = tabItem1.Fields("Limit") - tabItem2.Fields("Limit")
    If nTemp <> 0 Then CompareWeapons = AutoAppend(CompareWeapons, "Limit: " & IIf(nTemp > 0, "+", "") & nTemp & " (" & tabItem1.Fields("Limit") & ")")
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CompareWeapons")
Resume out:
End Function

Public Sub PullClassDetail(nClassNum As Long, DetailTB As TextBox)

On Error GoTo error:

Dim sAbil As String, x As Integer

tabClasses.Index = "pkClasses"
tabClasses.Seek "=", nClassNum
If tabClasses.NoMatch = True Then
    DetailTB.Text = "Class not found"
    Exit Sub
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

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("PullClassDetail")
Resume out:
End Sub
Public Sub PullRaceDetail(nRaceNum As Long, DetailTB As TextBox)

On Error GoTo error:

Dim sAbil As String, x As Integer


tabRaces.Index = "pkRaces"
tabRaces.Seek "=", nRaceNum
If tabRaces.NoMatch = True Then
    DetailTB.Text = "Race not found"
    Exit Sub
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



out:
On Error Resume Next
Exit Sub
error:
Call HandleError("PullRaceDetail")
Resume out:
End Sub
Public Sub PullMonsterDetail(nMonsterNum As Long, DetailLV As ListView)
On Error GoTo error:

Dim sAbil As String, x As Integer, y As Integer
Dim sCash As String, nCash As Currency, nPercent As Integer, nTest As Long
Dim oLI As ListItem, nExp As Currency, nMonsterDamage() As Variant, nMonsterEnergy As Long
Dim sReducedCoin As String, nReducedCoin As Currency

DetailLV.ListItems.clear

tabMonsters.Index = "pkMonsters"
tabMonsters.Seek "=", nMonsterNum
If tabMonsters.NoMatch = True Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Monster not found"
    'DetailTB.Text = "Monster not found"
    Set oLI = Nothing
    Exit Sub
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
nCash = nCash + (tabMonsters.Fields("R") * 1000000)
nCash = nCash + (tabMonsters.Fields("P") * 10000)
nCash = nCash + (tabMonsters.Fields("G") * 100)
nCash = nCash + (tabMonsters.Fields("S") * 10)
nCash = nCash + tabMonsters.Fields("C")

sReducedCoin = "Copper"
nReducedCoin = nCash
If nReducedCoin > 0 Then
    If nCash >= 10000000 Then
        nReducedCoin = nCash / 1000000
        sReducedCoin = "Runic"
    ElseIf nCash >= 100000 Then
        nReducedCoin = nCash / 10000
        sReducedCoin = "Platinum"
    ElseIf nCash >= 1000 Then
        nReducedCoin = nCash / 100
        sReducedCoin = "Gold"
    ElseIf nCash >= 100 Then
        nReducedCoin = nCash / 10
        sReducedCoin = "Silver"
    End If
    
    If Not sReducedCoin = "Copper" Then
        nReducedCoin = Round(nReducedCoin, 2)
    End If
    
    sCash = Format(nReducedCoin, "##,##0.00")
    If Right(sCash, 3) = ".00" Then sCash = Left(sCash, Len(sCash) - 3)

    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Cash (up to)"
    oLI.ListSubItems.Add (1), "Detail", sCash & " " & sReducedCoin
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

nMonsterDamage = CalculateMonsterAvgDmg(tabMonsters.Fields("Number"), 500) 'GetMonsterDamagePerRound(tabMonsters.Fields("Number"))
If nMonsterDamage(1) > 0 Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Dmg/Round *"
    oLI.ForeColor = RGB(204, 0, 0)
    
    If nMonsterDamage(0) < nMonsterDamage(1) Then
        oLI.ListSubItems.Add (1), "Detail", "AVG: " & nMonsterDamage(0) & ", Max: " & nMonsterDamage(1)
    Else
        oLI.ListSubItems.Add (1), "Detail", "AVG: " & nMonsterDamage(0)
    End If
    oLI.ListSubItems(1).Text = oLI.ListSubItems(1).Text & "   * quick 500 round sim, before character defenses"
    oLI.ListSubItems(1).ForeColor = RGB(204, 0, 0)
    
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
    If tabMonsters.Fields("AttType-" & x) > 0 And tabMonsters.Fields("AttType-" & x) <= 3 And tabMonsters.Fields("Att%-" & x) > 0 Then
        y = y + 1
        Set oLI = DetailLV.ListItems.Add()
        
        nPercent = tabMonsters.Fields("Att%-" & x) - nPercent
        If nPercent < 0 Then nPercent = 0
        
        If nNMRVer >= 1.8 Then
            If Round(tabMonsters.Fields("AttTrue%-" & x)) = 0 Then
                oLI.Text = "(" & nPercent & "%) "
            Else
                oLI.Text = "(" & Round(tabMonsters.Fields("AttTrue%-" & x)) & "%) "
            End If
            
            If Len(Trim(tabMonsters.Fields("AttName-" & x))) = 0 Then
                oLI.Text = oLI.Text & "Attack " & y
            Else
                oLI.Text = oLI.Text & Trim(tabMonsters.Fields("AttName-" & x))
            End If
        Else
            oLI.Text = "Attack " & y & " (" & nPercent & "%)"
        End If
        'oLI.ListSubItems.Add (1), "Detail", GetMonAttackType(tabMonsters.Fields("AttType-" & x))
        
        nPercent = tabMonsters.Fields("Att%-" & x)
        
        Select Case tabMonsters.Fields("AttType-" & x)
            Case 1, 3: 'normal, rob
                
                'Set oLI = DetailLV.ListItems.Add()
                'oLI.Text = ""
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
                
                'Set oLI = DetailLV.ListItems.Add()
                'oLI.Text = ""
                oLI.Tag = "Spell"
                'nSpellNest = 0 'make sure this doesn't nest too deep
                oLI.ListSubItems.Add (1), "Detail", "Spell: [" & _
                    GetSpellName(tabMonsters.Fields("AttAcc-" & x), bHideRecordNumbers) _
                    & ", " & PullSpellEQ(True, tabMonsters.Fields("AttMax-" & x), tabMonsters.Fields("AttAcc-" & x), , , True) & "]"
                If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
                oLI.ListSubItems(1).Tag = tabMonsters.Fields("AttAcc-" & x)
                
                If SpellIsAreaAttack(tabMonsters.Fields("AttAcc-" & x)) Then
                    Set oLI = DetailLV.ListItems.Add()
                    oLI.Text = ""
                    oLI.ListSubItems.Add (1), "Detail", "Target: " & GetSpellTargets(tabSpells.Fields("Targets"))
                End If
                
                Set oLI = DetailLV.ListItems.Add()
                oLI.Text = ""
                oLI.ListSubItems.Add (1), "Detail", "Success %: " & tabMonsters.Fields("AttMin-" & x)
                
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
                
                If SpellIsAreaAttack(tabMonsters.Fields("AttAcc-" & x)) Then
                    If GetSpellDuration(tabMonsters.Fields("AttAcc-" & x), tabMonsters.Fields("AttMax-" & x), True) = 0 Then
                        nTest = SpellHasAbility(tabMonsters.Fields("AttAcc-" & x), 1) '1=damage
                        If nTest > -1 Then
                            Set oLI = DetailLV.ListItems.Add()
                            oLI.Text = ""
                            oLI.ListSubItems.Add (1), "Detail", "NOTE: This is an invalid spell and will not be cast. Area attack spells from regular monster casts must use ability 17 (Damage-MR)."
                            oLI.ListSubItems(1).Bold = True
                        End If
                    End If
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



out:
Set oLI = Nothing
On Error Resume Next
Exit Sub
error:
Call HandleError("PullMonsterDetail")
Resume out:
End Sub

'Public Function GetMonsterDamagePerRound(nMonsterNum As Long) As Long()
'Dim x As Integer, y As Integer, z As Integer, nRound As Integer
'Dim nMonEnergy As Long, nPercent As Integer, nReturn() As Long
'Dim nMax As Long, nAVG As Long, nBetweenMax As Long, nBetweenAVG As Long
'Dim sSpellEQ As String, sArr() As String, nTempVal1 As Currency, nTempVal2 As Currency
'Dim nMonsterEnergy As Long, nRemainingEnergy As Currency
'
'Dim nAtkMin(4) As Long, nAtkMax(4) As Long, nAtkType(4) As Integer
'Dim nAtkAvgDmg(4) As Currency, nAtkAvgDmg_ADJ(4) As Currency
'Dim nAtkEnergy(4) As Long, nAtkEnergy_ADJ(4) As Long
'Dim nAtk_PCT(4) As Currency, nAtkSuccess_PCT(4) As Currency
'Dim nAtk_HitSpellAvgDmg(4) As Currency, nAtk_HitSpellMaxDmg(4) As Long
'
'On Error GoTo error:
'
'If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
'
'nPercent = 0
'For x = 0 To 4 'between round spells
'    If Not tabMonsters.Fields("MidSpell-" & x) = 0 Then
'
'        nPercent = tabMonsters.Fields("MidSpell%-" & x) - nPercent
'
'        sSpellEQ = PullSpellEQ(True, tabMonsters.Fields("MidSpellLVL-" & x), tabMonsters.Fields("MidSpell-" & x), , True, True)
'        If InStr(1, sSpellEQ, ":") > 0 Then
'            sArr() = Split(sSpellEQ, ":")
'            If UBound(sArr()) >= 1 Then
'                nTempVal1 = Val(sArr(0))
'                nTempVal2 = Val(sArr(1))
'                If UBound(sArr()) >= 2 Then
'                    If Val(sArr(2)) > 0 Then
'                        'normalize spell round (3sec) to combat round (5sec)
'                        nTempVal1 = (nTempVal1 * 1.666) '/ Val(sArr(2))
'                        nTempVal2 = (nTempVal2 * 1.666) '/ Val(sArr(2))
'                    End If
'                End If
'                If nTempVal2 > nBetweenMax Then nBetweenMax = nTempVal2
'                nBetweenAVG = nBetweenAVG + (((nTempVal1 + nTempVal2) / 2) * (nPercent / 100))
'            End If
'        End If
'
'        If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
'        nPercent = tabMonsters.Fields("MidSpell%-" & x)
'    End If
'Next
'
'nPercent = 0
'For x = 0 To 4
'    nAtkType(x) = tabMonsters.Fields("AttType-" & x)
'    nAtkEnergy(x) = tabMonsters.Fields("AttEnergy-" & x)
'
'    If nAtkType(x) > 0 And tabMonsters.Fields("Att%-" & x) > 0 Then
'        nAtk_PCT(x) = tabMonsters.Fields("Att%-" & x) - nPercent
'        nPercent = tabMonsters.Fields("Att%-" & x)
'        nAtk_PCT(x) = nAtk_PCT(x) / 100
'
'        nAtk_HitSpellMaxDmg(x) = 0
'        nAtk_HitSpellAvgDmg(x) = 0
'        If Not tabMonsters.Fields("AttHitSpell-" & x) = 0 Then
'            sSpellEQ = PullSpellEQ(False, , tabMonsters.Fields("AttHitSpell-" & x), , True, True)
'            If InStr(1, sSpellEQ, ":") > 0 Then
'                sArr() = Split(sSpellEQ, ":")
'                If UBound(sArr()) >= 1 Then
'                    nTempVal1 = Val(sArr(0))
'                    nTempVal2 = Val(sArr(1))
'
'                    nAtk_HitSpellAvgDmg(x) = (nTempVal1 + nTempVal2) / 2
'                    nAtk_HitSpellMaxDmg(x) = nTempVal2
'                End If
'            End If
'        End If
'
'        If nAtkType(x) = 2 Then 'spell
'
'            nAtkSuccess_PCT(x) = tabMonsters.Fields("AttMin-" & x) / 100
'            If nAtkSuccess_PCT(x) > 1 Then nAtkSuccess_PCT(x) = 1
'            nAtkEnergy_ADJ(x) = (nAtkEnergy(x) * nAtkSuccess_PCT(x)) + ((nAtkEnergy(x) / 2) * (1 - nAtkSuccess_PCT(x)))
'
'            sSpellEQ = PullSpellEQ(True, tabMonsters.Fields("AttMax-" & x), tabMonsters.Fields("AttAcc-" & x), , True, True)
'            If InStr(1, sSpellEQ, ":") > 0 Then
'                sArr() = Split(sSpellEQ, ":")
'                If UBound(sArr()) >= 1 Then
'                    nTempVal1 = Val(sArr(0))
'                    nTempVal2 = Val(sArr(1))
'                    If UBound(sArr()) >= 2 Then
'                        If Val(sArr(2)) > 0 Then
'                            'normalize spell rounds (3sec) to combat rounds (5sec)
'                            nTempVal1 = (nTempVal1 * 1.666) '/ Val(sArr(2))
'                            nTempVal2 = (nTempVal2 * 1.666) '/ Val(sArr(2))
'                        End If
'                    End If
'                    nAtkMin(x) = nTempVal1
'                    nAtkMax(x) = nTempVal2
'
'                    nAtkAvgDmg(x) = ((nAtkMin(x) + nAtkMax(x)) / 2) + nAtk_HitSpellAvgDmg(x)
'                    nAtkAvgDmg_ADJ(x) = nAtkAvgDmg(x) * nAtkSuccess_PCT(x)
'                End If
'            End If
'            If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
'        Else
'
'            nAtkSuccess_PCT(x) = tabMonsters.Fields("AttAcc-" & x) / 100
'
'            'PUT AC ADJUSTMENT HERE
'            nAtkSuccess_PCT(x) = 0.98
'            If nAtkSuccess_PCT(x) > 0.98 Then nAtkSuccess_PCT(x) = 0.98
'
'
'            nAtkEnergy_ADJ(x) = nAtkEnergy(x)
'
'            nAtkMin(x) = tabMonsters.Fields("AttMin-" & x)
'            nAtkMax(x) = tabMonsters.Fields("AttMax-" & x)
'
'            nAtkAvgDmg(x) = ((nAtkMin(x) + nAtkMax(x)) / 2) + nAtk_HitSpellAvgDmg(x)
'            nAtkAvgDmg_ADJ(x) = nAtkAvgDmg(x) * nAtkSuccess_PCT(x)
'
'        End If
'    End If
'Next x
'
'nRemainingEnergy = 0
'If nNMRVer >= 1.71 Then
'    nMonsterEnergy = tabMonsters.Fields("Energy")
'Else
'    nMonsterEnergy = 1000
'End If
'
''Dim nAtkMin(4) As Long, nAtkMax(4) As Long, nAtkType(4) As Integer
''Dim nAtkAvgDmg(4) As Long, nAtkAvgDmg_ADJ(4) As Long
''Dim nAtkEnergy(4) As Long, nAtkEnergy_ADJ(4) As Long
''Dim nAtk_PCT(4) As Currency, nAtkSuccess_PCT(4) As Currency
'Dim nAttackAttempt As Integer, nStartingEnergy As Long
'Dim nChanceAtNotHavingEnergy_Attempt(4) As Currency, nChanceAtNotHavingEnergy_PREV(4) As Currency
'Dim nEnergyUsageThatWouldCauseDeficit As Currency, nPreviousAttemptStartingEnergy As Long
'Dim n_PREV_Atk_PCT(4) As Currency, nAtk_PCT_For_Attempt(4) As Currency, nDamageCaused As Currency
'Dim nMaxRoundsToSim As Integer, nMaxEnergy As Long
'
'nDamageCaused = 0
'nRemainingEnergy = 0
'nMaxRoundsToSim = 4
'nMaxEnergy = nMonsterEnergy
'For nRound = 1 To nMaxRoundsToSim
'    'reset n_PREV_Atk_PCT
'    For x = 0 To 4
'        nChanceAtNotHavingEnergy_PREV(x) = 0
'        n_PREV_Atk_PCT(x) = nAtk_PCT(x)
'    Next x
'
'    'add monster energy to left over energy
'    nRemainingEnergy = nMonsterEnergy + nRemainingEnergy
'    If nRemainingEnergy > nMaxEnergy Then nMaxEnergy = nRemainingEnergy
'    nPreviousAttemptStartingEnergy = nRemainingEnergy
'
'    '6 monster attack attempts
'    For nAttackAttempt = 1 To 6
'
'        nStartingEnergy = nRemainingEnergy
'        For x = 0 To 4
'            If nAtkAvgDmg_ADJ(x) > 0 Then
'                nChanceAtNotHavingEnergy_Attempt(x) = 0
'                If nAttackAttempt > 1 Then
'                    nEnergyUsageThatWouldCauseDeficit = nPreviousAttemptStartingEnergy - nAtkEnergy(x)
'                    For z = 0 To 4
'                        If nAtkEnergy_ADJ(z) >= nEnergyUsageThatWouldCauseDeficit Then
'                            nChanceAtNotHavingEnergy_Attempt(x) = nChanceAtNotHavingEnergy_Attempt(x) + n_PREV_Atk_PCT(z)
'                        End If
'                    Next z
'                ElseIf nAtkEnergy(x) > nStartingEnergy Then
'                    nChanceAtNotHavingEnergy_Attempt(x) = 1
'                End If
'                nChanceAtNotHavingEnergy_Attempt(x) = nChanceAtNotHavingEnergy_Attempt(x) + nChanceAtNotHavingEnergy_PREV(x)
'                If nChanceAtNotHavingEnergy_Attempt(x) > 1 Then nChanceAtNotHavingEnergy_Attempt(x) = 1
'                nAtk_PCT_For_Attempt(x) = (1 - nChanceAtNotHavingEnergy_Attempt(x)) * nAtk_PCT(x)
'
'                nDamageCaused = nDamageCaused + (nAtkAvgDmg_ADJ(x) * nAtk_PCT_For_Attempt(x))
'                nRemainingEnergy = nRemainingEnergy - (nAtkEnergy_ADJ(x) * nAtk_PCT_For_Attempt(x))
'            End If
'        Next x
'
'        For x = 0 To 4
'            nChanceAtNotHavingEnergy_PREV(x) = nChanceAtNotHavingEnergy_Attempt(x)
'            n_PREV_Atk_PCT(x) = nAtk_PCT_For_Attempt(x)
'        Next x
'
'        nPreviousAttemptStartingEnergy = nStartingEnergy
'
'    Next nAttackAttempt
'
'Next nRound
'
'Dim nMaxNumAttacks As Integer
'
'nAVG = nDamageCaused / nMaxRoundsToSim
'For x = 0 To 4
'    If nAtkEnergy(x) > 0 Then
'        nMaxNumAttacks = (nMaxEnergy / nAtkEnergy(x))
'        If nMaxNumAttacks > 6 Then nMaxNumAttacks = 6
'        If nMaxNumAttacks * (nAtkMax(x) + nAtk_HitSpellMaxDmg(x)) > nMax Then nMax = nMaxNumAttacks * (nAtkMax(x) + nAtk_HitSpellMaxDmg(x))
'    End If
'Next x
'out:
'On Error Resume Next
'
'ReDim nReturn(0 To 1)
'nReturn(0) = nAVG + nBetweenAVG
'nReturn(1) = nMax + nBetweenMax
'
'GetMonsterDamagePerRound = nReturn
'
'Exit Function
'error:
'Call HandleError("GetMonsterDamagePerRound")
'Resume out:
'End Function

Public Sub PullShopDetail(nShopNum As Long, DetailLV As ListView, _
    DetailTB As TextBox, lvAssigned As ListView, ByVal nCharm As Integer, ByVal bShowSell As Boolean)

On Error GoTo error:

Dim sStr As String, x As Integer, nRegenTime As Integer, sRegenTime As String
Dim oLI As ListItem, tCostType As typItemCostDetail, nCopper As Currency, sCopper As String
Dim nCharmMod As Double, sCharmMod As String
Dim nReducedCoin As Currency, sReducedCoin As String

DetailLV.ListItems.clear

tabShops.Index = "pkShops"
tabShops.Seek "=", nShopNum
If tabShops.NoMatch = True Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Shop not found"
    DetailTB.Text = "Shop not found"
    Set oLI = Nothing
    Exit Sub
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


out:
On Error Resume Next
Set oLI = Nothing
Exit Sub
error:
Call HandleError("PullShopDetail")
Resume out:
End Sub


Public Sub PullSpellDetail(nSpellNum As Long, DetailTB As TextBox, LocationLV As ListView)

On Error GoTo error:

Dim sStr As String
Dim sRemoves As String, sArr() As String, x As Integer


tabSpells.Index = "pkSpells"
tabSpells.Seek "=", nSpellNum
If tabSpells.NoMatch = True Then
    DetailTB.Text = "spell not found"
    Exit Sub
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

If nNMRVer >= 1.8 Then
    If tabSpells.Fields("TypeOfResists") = 1 Then
        sStr = sStr & " -- Resistable by Anti-Magic Only"
    ElseIf tabSpells.Fields("TypeOfResists") = 2 Then
        sStr = sStr & " -- Resistable by Everyone"
    End If
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



out:
On Error Resume Next
Exit Sub
error:
Call HandleError("PullSpellDetail")
Resume out:
End Sub

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

On Error GoTo error:

Dim oLI As ListItem, x As Integer, sName As String, nAbilityVal As Integer

sName = tabItems.Fields("Name")
If sName = "" Then GoTo skip:

Set oLI = LV.ListItems.Add()
oLI.Text = tabItems.Fields("Number")

oLI.ListSubItems.Add (1), "Name", tabItems.Fields("Name")
oLI.ListSubItems.Add (2), "Worn", GetWornType(tabItems.Fields("Worn"))
oLI.ListSubItems(2).Tag = tabItems.Fields("Worn")
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
        Case 58: ' crits
            oLI.ListSubItems(8).Text = tabItems.Fields("AbilVal-" & x)
            
        Case 135: 'min level
            oLI.ListSubItems(4).Text = tabItems.Fields("AbilVal-" & x)
        
        Case 22, 105, 106: 'acc
            oLI.ListSubItems(7).Text = Val(oLI.ListSubItems(7).Text) + tabItems.Fields("AbilVal-" & x)
    End Select
    
    If nAbility > 0 And tabItems.Fields("Abil-" & x) = nAbility Then
        nAbilityVal = tabItems.Fields("AbilVal-" & x)
    End If
Next x

If nAbility > 0 Then
    oLI.ListSubItems.Add (11), "Ability", nAbilityVal
End If

If AddToInven Then Call frmMain.InvenAddEquip(tabItems.Fields("Number"), sName, tabItems.Fields("ItemType"), tabItems.Fields("Worn"))
    
skip:
Set oLI = Nothing


out:
On Error Resume Next
Exit Sub
error:
Call HandleError("AddArmour2LV")
Resume out:
End Sub
Public Sub AddOtherItem2LV(LV As ListView)

On Error GoTo error:

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


out:
On Error Resume Next
Exit Sub
error:
Call HandleError("AddOtherItem2LV")
Resume out:
End Sub
Public Sub AddWeapon2LV(LV As ListView, Optional AddToInven As Boolean, Optional nAbility As Integer)

On Error GoTo error:

Dim oLI As ListItem, x As Integer, sName As String, nSpeed As Integer, nDMG As Currency, nAbilityVal As Integer

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
oLI.ListSubItems.Add (10), "Acc", 0 'tabItems.Fields("Accy")

oLI.ListSubItems.Add (11), "BS Acc", "No"
oLI.ListSubItems.Add (12), "Crits", 0

For x = 0 To 19
    Select Case tabItems.Fields("Abil-" & x)
        Case 58: 'crits
            oLI.ListSubItems(12).Text = tabItems.Fields("AbilVal-" & x)
            
        Case 22, 105, 106: 'acc
            oLI.ListSubItems(10).Text = tabItems.Fields("AbilVal-" & x)
        
        Case 135: 'min level
            oLI.ListSubItems(6).Text = tabItems.Fields("AbilVal-" & x)
            
        Case 116: 'bs accu
            oLI.ListSubItems(11).Text = tabItems.Fields("AbilVal-" & x)
    End Select
    
    If nAbility > 0 And tabItems.Fields("Abil-" & x) = nAbility Then
        nAbilityVal = tabItems.Fields("AbilVal-" & x)
    End If
Next x

oLI.ListSubItems(10).Text = Val(oLI.ListSubItems(10).Text) + tabItems.Fields("Accy")

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


out:
On Error Resume Next
Exit Sub
error:
Call HandleError("AddWeapon2LV")
Resume out:
End Sub

Public Sub AddSpell2LV(LV As ListView, Optional ByVal AddBless As Boolean)

On Error GoTo error:

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
        oLI.ListSubItems.Add (7), "Dmg", Round(nSpellDamage)
        
        If nSpellDamage > 0 Then
            If tabSpells.Fields("ManaCost") > 0 Then
                If tabSpells.Fields("EnergyCost") >= 200 And tabSpells.Fields("EnergyCost") <= 500 Then
                    nSpellDamage = Round(nSpellDamage / (tabSpells.Fields("ManaCost") * Round(1000 / tabSpells.Fields("EnergyCost"))), 1)
                Else
                    nSpellDamage = Round(nSpellDamage / tabSpells.Fields("ManaCost"), 1)
                End If
            End If
        End If
    Else
        oLI.ListSubItems.Add (7), "Dmg", 0
        nSpellDamage = 0
    End If
    
    oLI.ListSubItems.Add (8), "Dmg/M", nSpellDamage
    
    bQuickSpell = True
    If LV.name = "lvSpellBook" And FormIsLoaded("frmSpellBook") And frmMain.chkGlobalFilter.Value = 1 Then
        If Val(frmSpellBook.txtLevel) > 0 Then
            oLI.ListSubItems.Add (9), "Detail", PullSpellEQ(True, Val(frmSpellBook.txtLevel), nSpell)
        Else
            oLI.ListSubItems.Add (9), "Detail", PullSpellEQ(False, , nSpell)
        End If
    Else
        oLI.ListSubItems.Add (9), "Detail", PullSpellEQ(False, , nSpell)
    End If
    bQuickSpell = False
    
    If Not tabSpells.Fields("Number") = nSpell Then tabSpells.Seek "=", nSpell
    
    If AddBless Then
        If tabSpells.Fields("Learnable") = 1 Or Len(tabSpells.Fields("Learned From")) > 0 _
            Or (tabSpells.Fields("Magery") = 5 And tabSpells.Fields("ReqLevel") > 0) Then
                        
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
                frmMain.cmbCharBless(x).AddItem sName & " (" & nSpell & ")"
                frmMain.cmbCharBless(x).ItemData(frmMain.cmbCharBless(x).NewIndex) = nSpell
            Next x
        End If
    End If
    
skip:
Set oLI = Nothing


out:
On Error Resume Next
Exit Sub
error:
Call HandleError("AddSpell2LV")
Resume out:
End Sub

Public Sub AddRace2LV(LV As ListView)

On Error GoTo error:

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


out:
On Error Resume Next
Exit Sub
error:
Call HandleError("AddRace2LV")
Resume out:
End Sub

Public Sub AddMonster2LV(LV As ListView)

On Error GoTo error:

Dim oLI As ListItem, sName As String, nExp As Currency, x As Integer
Dim nAvgDMG As Long, nExpDmgHP As Currency, nIndex As Integer, nMagicLVL As Integer
Dim nScriptValue As Currency, nLairPCT As Currency, nPossSpawns As Long, sPossSpawns As String

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
nLairPCT = nPossSpawns
If InStr(1, tabMonsters.Fields("Summoned By"), "(lair)", vbTextCompare) > 0 Then
    nPossSpawns = InstrCount(tabMonsters.Fields("Summoned By"), "(lair)")
    sPossSpawns = nPossSpawns
    nLairPCT = nPossSpawns
    If nAverageLairs > 0 Then
        nLairPCT = Round(IIf(nPossSpawns > 100, 100, nPossSpawns) / nAverageLairs, 2)
        sPossSpawns = nPossSpawns & " (" & (nLairPCT * 100) & "%)"
    End If
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

oLI.ListSubItems.Add (nIndex), "Lairs", sPossSpawns
oLI.ListSubItems(nIndex).Tag = nPossSpawns
nIndex = nIndex + 1

If tabMonsters.Fields("RegenTime") > 0 Or nLairPCT <= 0 Then
    nScriptValue = 0
Else
    nScriptValue = nExpDmgHP * nLairPCT
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

For x = 0 To 9 'abilities
    If Not tabMonsters.Fields("Abil-" & x) = 0 Then
        Select Case tabMonsters.Fields("Abil-" & x)
            Case 28: 'magical
                nMagicLVL = tabMonsters.Fields("AbilVal-" & x)
                Exit For
        End Select
    End If
Next

oLI.ListSubItems.Add (nIndex), "Mag.", nMagicLVL
oLI.ListSubItems(nIndex).Tag = nMagicLVL
nIndex = nIndex + 1

skip:
Set oLI = Nothing


out:
On Error Resume Next
Exit Sub
error:
Call HandleError("AddMonster2LV")
Resume out:
End Sub

Public Sub AddShop2LV(LV As ListView)

On Error GoTo error:

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


out:
On Error Resume Next
Exit Sub
error:
Call HandleError("AddShop2LV")
Resume out:
End Sub

Public Sub AddClass2LV(LV As ListView)

On Error GoTo error:

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


out:
On Error Resume Next
Exit Sub
error:
Call HandleError("AddClass2LV")
Resume out:
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

On Error GoTo error:

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

Exit Function
error:
Call HandleError
End Function

Public Sub CopyWholeLVtoClipboard(LV As ListView, Optional ByVal UsePeriods As Boolean)
On Error GoTo error:
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
error:
Call HandleError("CopyWholeLVtoClip")
Set oLI = Nothing
Set oLSI = Nothing
Set oCH = Nothing
End Sub
Public Sub CopyLVLinetoClipboard(LV As ListView, Optional DetailTB As TextBox, _
    Optional LocationLV As ListView, Optional ByVal nExcludeColumn As Integer = -1, Optional bNameOnly As Boolean = False)
On Error GoTo error:
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
                    If LV.name = "lvMapLoc" And x = 0 Then
                        If InStr(1, oLI.Text, ":", vbTextCompare) > 0 Then
                            str = str & Trim(Mid(oLI.Text, InStr(1, oLI.Text, ":", vbTextCompare) + 1, 999))
                        Else
                            str = str & oLI.Text
                        End If
                    ElseIf oCH.Text = "Name" Then
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
error:
Call HandleError("CopyLVLinetoClip")
Resume out:
End Sub

Public Sub GetLocations(ByVal sLoc As String, LV As ListView, _
    Optional bDontClear As Boolean, Optional ByVal sHeader As String, _
    Optional ByVal nAuxValue As Long, Optional ByVal bTwoColumns As Boolean, _
    Optional ByVal bDontSort As Boolean, Optional ByVal bPercentColumn As Boolean, _
    Optional ByVal sFooter As String)
On Error GoTo error:
Dim sLook As String, sChar As String, sTest As String, oLI As ListItem, sPercent As String
Dim x As Integer, y1 As Integer, y2 As Integer, z As Integer, nValue As Long, x2 As Integer
Dim sLocation As String, nPercent As Currency, nPercent2 As Currency, sTemp As String
Dim sDisplayFooter As String

sDisplayFooter = sFooter

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
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "/", ".":
                If Not y1 + y2 - 1 = Len(sTest) Then
                    y2 = y2 + 1
                    GoTo nextnumber:
                End If
            Case "+": 'end of string
                Exit Sub
            Case "(": 'percent
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
        If bPercentColumn Then
            nPercent = ExtractNumbersFromString(sPercent)
            
            If InStr(1, sFooter, "%)", vbTextCompare) > 0 Then
                nPercent2 = ExtractNumbersFromString(Mid(sFooter, InStr(1, sFooter, "(", vbTextCompare)))
                If nPercent2 > 0 Then
                    sDisplayFooter = Replace(sFooter, "(" & nPercent2 & "%)", "")
                    If nPercent > 0 Then
                        nPercent = (nPercent * nPercent2) / 100
                    Else
                        nPercent = nPercent2
                    End If
                End If
            End If
            
            If nPercent > 1 Then
                nPercent = Round(nPercent)
            Else
                nPercent = Round(nPercent, 2)
            End If
            'sPercent = " (" & nPercent & "%)"
        End If
        
        Select Case z
            Case 1: '"room "
                sLocation = "Room: "
                If Not sHeader = "" Then
                    If InStr(1, sHeader, ":", vbTextCompare) = 0 Then
                        sLocation = sLocation & sHeader
                    Else
                        sLocation = sHeader
                    End If
                End If
                
                Set oLI = LV.ListItems.Add()
                If bPercentColumn Then
                    oLI.Text = ""
                    If nPercent > 0 Then oLI.Text = nPercent & "%"
                    oLI.Tag = nPercent
                    oLI.ListSubItems.Add 1, , sLocation & GetRoomName(Mid(sTest, y1, y2), , , bHideRecordNumbers) & sDisplayFooter
                    
                ElseIf bTwoColumns Then
                    oLI.Text = sLocation
                    oLI.ListSubItems.Add 1, , GetRoomName(Mid(sTest, y1, y2), , , bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = "Room"
                Else
                    oLI.Text = sLocation & GetRoomName(Mid(sTest, y1, y2), , , bHideRecordNumbers) & sPercent & sDisplayFooter
                End If
                
                If nAuxValue > 0 Then
                    If bTwoColumns Or bPercentColumn Then
                        oLI.ListSubItems(1).Tag = nAuxValue
                    Else
                        oLI.Tag = nAuxValue
                    End If
                Else
                    If bTwoColumns Or bPercentColumn Then
                        oLI.ListSubItems(1).Tag = Mid(sTest, y1, y2)
                    Else
                        oLI.Tag = Mid(sTest, y1, y2)
                    End If
                End If
                
            Case 2: '"monster #"
                sLocation = "Monster: "
                Set oLI = LV.ListItems.Add()
                If bPercentColumn Then
                    oLI.Text = ""
                    If nPercent > 0 Then oLI.Text = nPercent & "%"
                    oLI.Tag = nPercent
                    oLI.ListSubItems.Add 1, , sLocation & sHeader & GetMonsterName(nValue, bHideRecordNumbers) & sDisplayFooter
                    oLI.ListSubItems(1).Tag = nValue
                ElseIf bTwoColumns Then
                    oLI.Text = sLocation & sHeader
                    oLI.ListSubItems.Add 1, , GetMonsterName(nValue, bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = "monster"
                    oLI.ListSubItems(1).Tag = nValue
                Else
                    oLI.Text = sLocation & sHeader & GetMonsterName(nValue, bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = nValue
                End If
                
            Case 3: '"textblock #"
                sLocation = "Textblock "
                Set oLI = LV.ListItems.Add()
                If bPercentColumn Then
                    oLI.Text = ""
                    If nPercent > 0 Then oLI.Text = nPercent & "%"
                    oLI.Tag = nPercent
                    oLI.ListSubItems.Add 1, , sLocation & sHeader & nValue & sDisplayFooter
                    oLI.ListSubItems(1).Tag = nValue
                    
                ElseIf bTwoColumns Then
                    oLI.Text = sLocation & sHeader
                    oLI.ListSubItems.Add 1, , nValue & sPercent & sDisplayFooter
                    oLI.Tag = "textblock"
                    oLI.ListSubItems(1).Tag = nValue
                Else
                    oLI.Text = sLocation & sHeader & nValue & sPercent & sDisplayFooter
                    oLI.Tag = nValue
                End If
                
            Case 4: '"textblock(rndm) #"
                sLocation = "Textblock "
                Set oLI = LV.ListItems.Add()
                If bPercentColumn Then
                    oLI.Text = ""
                    If nPercent > 0 Then oLI.Text = nPercent & "%"
                    oLI.Tag = nPercent
                    oLI.ListSubItems.Add 1, , sLocation & sHeader & nValue & sDisplayFooter
                    oLI.ListSubItems(1).Tag = nValue
                ElseIf bTwoColumns Then
                    oLI.Text = sLocation & sHeader
                    oLI.ListSubItems.Add 1, , nValue & sPercent & sDisplayFooter
                    oLI.Tag = "textblock"
                    oLI.ListSubItems(1).Tag = nValue
                Else
                    oLI.Text = sLocation & sHeader & nValue & sPercent & sDisplayFooter
                    oLI.Tag = nValue
                End If
                
            Case 5: '"item #"
                sLocation = "Item: "
                Set oLI = LV.ListItems.Add()
                If bPercentColumn Then
                    oLI.Text = ""
                    If nPercent > 0 Then oLI.Text = nPercent & "%"
                    oLI.Tag = nPercent
                    oLI.ListSubItems.Add 1, , sLocation & sHeader & GetItemName(nValue, bHideRecordNumbers) & sDisplayFooter
                    oLI.ListSubItems(1).Tag = nValue
                ElseIf bTwoColumns Then
                    oLI.Text = sLocation & sHeader
                    oLI.ListSubItems.Add 1, , GetItemName(nValue, bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = "item"
                    oLI.ListSubItems(1).Tag = nValue
                Else
                    oLI.Text = sLocation & sHeader & GetItemName(nValue, bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = nValue
                End If
                
                If ItemIsChest(nValue) And sHeader = "" And sFooter = "" Then
                    Call GetLocations(tabItems.Fields("Obtained From"), LV, True, , , , True, bPercentColumn, " -> " & tabItems.Fields("Name") & sPercent)
                End If
                
            Case 6: '"spell #"
                sLocation = "Spell: "
                Set oLI = LV.ListItems.Add()
                If bPercentColumn Then
                    oLI.Text = ""
                    If nPercent > 0 Then oLI.Text = nPercent & "%"
                    oLI.Tag = nPercent
                    oLI.ListSubItems.Add 1, , sLocation & sHeader & GetSpellName(nValue, bHideRecordNumbers) & sDisplayFooter
                    oLI.ListSubItems(1).Tag = nValue
                ElseIf bTwoColumns Then
                    oLI.Text = sLocation & sHeader
                    oLI.ListSubItems.Add 1, , GetSpellName(nValue, bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = "spell"
                    oLI.ListSubItems(1).Tag = nValue
                Else
                    oLI.Text = sLocation & sHeader & GetSpellName(nValue, bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = nValue
                End If
                
            Case 7: '"shop #"
                If bPercentColumn And nAuxValue > 0 Then
                    nPercent = GetItemShopRegenPCT(nValue, nAuxValue)
                    If nPercent > 0 Then
                        sTemp = GetShopLocation(nValue)
                        sTemp = Join(Split(sTemp, ","), "(" & nPercent & "%),")
                        If Not Right(sTemp, 2) = "%)" Then sTemp = sTemp & "(" & nPercent & "%)"
                        Call GetLocations(sTemp, LV, True, "Shop: ", nValue, , , bPercentColumn)
                    Else
                        Call GetLocations(GetShopLocation(nValue), LV, True, "Shop: ", nValue, , , bPercentColumn)
                    End If
                Else
                    Call GetLocations(GetShopLocation(nValue), LV, True, "Shop: ", nValue)
                End If
                'Set oLI = LV.ListItems.Add()
                'oLI.Text = "Shop: " & GetShopName(nValue) & sPercent
                'oLI.Tag = nValue
            Case 8: '"shop(sell) #"
                Call GetLocations(GetShopLocation(nValue), LV, True, "Shop (sell): ", nValue, , , bPercentColumn)
'                Set oLI = LV.ListItems.Add()
'                oLI.Text = "Shop: " & GetShopName(nValue) & " (sell only)" & sPercent
'                oLI.Tag = nValue
            Case 9: '"shop(nogen) #"
                Call GetLocations(GetShopLocation(nValue), LV, True, "Shop (nogen): ", nValue, , , bPercentColumn)
'                Set oLI = LV.ListItems.Add()
'                oLI.Text = "Shop: " & GetShopName(nValue) & " (wont regen)" & sPercent
'                oLI.Tag = nValue
            Case 10: 'group (lair)
                sLocation = "Group(Lair): "
                Set oLI = LV.ListItems.Add()
                If bPercentColumn Then
                    oLI.Text = ""
                    If nPercent > 0 Then oLI.Text = nPercent & "%"
                    oLI.Tag = nPercent
                    oLI.ListSubItems.Add 1, , sLocation & GetRoomName(Mid(sTest, y1, y2), , , bHideRecordNumbers) & sDisplayFooter
                    oLI.ListSubItems(1).Tag = Mid(sTest, y1, y2)
                ElseIf bTwoColumns Then
                    oLI.Text = sLocation
                    oLI.ListSubItems.Add 1, , GetRoomName(Mid(sTest, y1, y2), , , bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = "Room"
                    oLI.ListSubItems(1).Tag = Mid(sTest, y1, y2)
                Else
                    oLI.Text = sLocation & GetRoomName(Mid(sTest, y1, y2), , , bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = Mid(sTest, y1, y2)
                End If
                
            Case 11: 'group
                sLocation = "Group: "
                Set oLI = LV.ListItems.Add()
                If bPercentColumn Then
                    oLI.Text = ""
                    If nPercent > 0 Then oLI.Text = nPercent & "%"
                    oLI.Tag = nPercent
                    oLI.ListSubItems.Add 1, , sLocation & GetRoomName(Mid(sTest, y1, y2), , , bHideRecordNumbers) & sDisplayFooter
                    oLI.ListSubItems(1).Tag = Mid(sTest, y1, y2)
                ElseIf bTwoColumns Then
                    oLI.Text = sLocation
                    oLI.ListSubItems.Add 1, , GetRoomName(Mid(sTest, y1, y2), , , bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = "Room"
                    oLI.ListSubItems(1).Tag = Mid(sTest, y1, y2)
                Else
                    oLI.Text = sLocation & GetRoomName(Mid(sTest, y1, y2), , , bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = Mid(sTest, y1, y2)
                End If
            
            Case 12: '"NPC #"
                sLocation = "NPC: "
                Set oLI = LV.ListItems.Add()
                If bPercentColumn Then
                    oLI.Text = ""
                    If nPercent > 0 Then oLI.Text = nPercent & "%"
                    oLI.Tag = nPercent
                    oLI.ListSubItems.Add 1, , sLocation & sHeader & GetMonsterName(nValue, bHideRecordNumbers) & sDisplayFooter
                    oLI.ListSubItems(1).Tag = nValue
                ElseIf bTwoColumns Then
                    oLI.Text = sLocation & sHeader
                    oLI.ListSubItems.Add 1, , GetMonsterName(nValue, bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = "monster"
                    oLI.ListSubItems(1).Tag = nValue
                Else
                    oLI.Text = sLocation & sHeader & GetMonsterName(nValue, bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = nValue
                End If
        End Select
        
        x = y1
        GoTo checknext:
    End If
Next z

If LV.ListItems.Count > 1 And Not bDontSort And Not bPercentColumn Then
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

error:
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
    Case 17: GetEquipCaption = "Eyes"
    Case 18: GetEquipCaption = "Face"
    Case 19: GetEquipCaption = "Worn"
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
On Error GoTo error:
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

error:
RegCreateKeyPath = Err.Number
Resume quit:
End Function

Public Function StringOfNumbersToArray(sNumberString As String) As String()
Dim x As Long, sRet() As String
On Error GoTo error:

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
error:
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

Public Sub ColorListviewRow(LV As ListView, RowNbr As Long, RowColor As OLE_COLOR, Optional bAndBold As Boolean)

On Error GoTo error:

'***************************************************************************
'Purpose: Color a ListView Row
'Inputs : lv - The ListView
'         RowNbr - The index of the row to be colored
'         RowColor - The color to color it
'Outputs: None
'***************************************************************************
    
Dim itmX As ListItem
Dim lvSI As ListSubItem
Dim intIndex As Integer


Set itmX = LV.ListItems(RowNbr)
itmX.ForeColor = RowColor
If bAndBold Then
    itmX.Bold = True
Else
    itmX.Bold = False
End If
For intIndex = 1 To itmX.ListSubItems.Count
    Set lvSI = itmX.ListSubItems(intIndex)
    lvSI.ForeColor = RowColor
    If bAndBold Then
        lvSI.Bold = True
    Else
        lvSI.Bold = False
    End If
    DoEvents
Next
'lv.ListItems(2).Selected = True
'lv.ListItems(1).Selected = True

Set itmX = Nothing
Set lvSI = Nothing
   

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("ColorListviewRow")
Resume out:
End Sub

Public Function InstrCount(StringToSearch As String, StringToFind As String) As Long
    If Len(StringToFind) > 0 Then
        InstrCount = UBound(Split(StringToSearch, StringToFind))
    End If
End Function

Public Function Get_MegaMUD_ExitsCode(ByVal nMapNum As Long, ByVal nRoomNum As Long) As String
Dim RoomExit As RoomExitType, x As Integer, sRoomName As String
Dim sExits(9) As String, nExitVal(9) As Integer, nExitPosition(9) As Integer
Dim nExitsCalculated(1 To 5) As Integer
Dim bDoor As Boolean
On Error GoTo error:

sExits(0) = "N"
sExits(1) = "S"
sExits(2) = "E"
sExits(3) = "W"
sExits(4) = "NE"
sExits(5) = "NW"
sExits(6) = "SE"
sExits(7) = "SW"
sExits(8) = "U"
sExits(9) = "D"

nExitVal(0) = 1 '"N"
nExitVal(1) = 4 '"S"
nExitVal(2) = 1 '"E"
nExitVal(3) = 4 '"W"
nExitVal(4) = 1 '"NE"
nExitVal(5) = 4 '"NW"
nExitVal(6) = 1 '"SE"
nExitVal(7) = 4 '"SW"
nExitVal(8) = 1 '"U"
nExitVal(9) = 4 '"D"

nExitPosition(0) = 5 '"N"
nExitPosition(1) = 5 '"S"
nExitPosition(2) = 4 '"E"
nExitPosition(3) = 4 '"W"
nExitPosition(4) = 3 '"NE"
nExitPosition(5) = 3 '"NW"
nExitPosition(6) = 2 '"SE"
nExitPosition(7) = 2 '"SW"
nExitPosition(8) = 1 '"U"
nExitPosition(9) = 1 '"D"

tabRooms.Index = "idxRooms"
tabRooms.Seek "=", nMapNum, nRoomNum
If tabRooms.NoMatch Then GoTo out:

sRoomName = tabRooms.Fields("Name")

For x = 0 To 9
    bDoor = False
    If Not Val(tabRooms.Fields(sExits(x))) = 0 Then
        RoomExit = ExtractMapRoom(tabRooms.Fields(sExits(x)))
        If RoomExit.Map > 0 And RoomExit.Room > 0 Then
            If Len(RoomExit.ExitType) > 2 Then
                Select Case Left(RoomExit.ExitType, 5)
                    Case "(Key:": bDoor = True
                    Case "(Item":
                    Case "(Toll":
                    Case "(Hidd": GoTo exit_not_seen:
                    Case "(Door": bDoor = True
                    Case "(Trap":
                    Case "(Text": GoTo exit_not_seen:
                    Case "(Gate": bDoor = True
                    Case "Actio": GoTo exit_not_seen:
                    Case "(Clas":
                    Case "(Race":
                    Case "(Leve":
                    Case "(Time":
                    Case "(Tick":
                    Case "(Max ":
                    Case "(Bloc":
                    Case "(Alig":
                    Case "(Dela":
                    Case "(Cast":
                    Case "(Abil":
                    Case "(Spel":
                End Select
            End If
            nExitsCalculated(nExitPosition(x)) = nExitsCalculated(nExitPosition(x)) + (nExitVal(x) * IIf(bDoor, 2, 1))
        End If
    End If
exit_not_seen:
Next x

For x = 1 To 5
    Get_MegaMUD_ExitsCode = Get_MegaMUD_ExitsCode & Hex(nExitsCalculated(x))
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("Get_MegaMUD_ExitsCode")
End Function

    
Public Function Get_MegaMUD_RoomHash(ByVal sRoomName As String, Optional ByVal nMapNum As Long, Optional ByVal nRoomNum As Long) As String
Dim x As Integer, nValue As Long
On Error GoTo error:

Get_MegaMUD_RoomHash = "FFF"

If nMapNum > 0 And nRoomNum > 0 Then
    tabRooms.Index = "idxRooms"
    tabRooms.Seek "=", nMapNum, nRoomNum
    If tabRooms.NoMatch Then GoTo out:
    sRoomName = tabRooms.Fields("Name")
End If

'// Calculate the checksum of the room name
'dwCheckSum = 0;
'for (i = 0; pRec->szName[i]; i++)
'    dwCheckSum += (DWORD) ((int)(pRec->szName[i] * (i+1)));
'dwCheckSum = dwCheckSum << 20;
    
x = 1
Do While x <= Len(sRoomName)
    nValue = nValue + (x * Asc(Mid(sRoomName, x, 1)))
    x = x + 1
Loop

Get_MegaMUD_RoomHash = Hex(nValue)
Get_MegaMUD_RoomHash = Right(Get_MegaMUD_RoomHash, 3)
If Len(Get_MegaMUD_RoomHash) < 3 Then Get_MegaMUD_RoomHash = String(3 - Len(Get_MegaMUD_RoomHash), "0") & Get_MegaMUD_RoomHash

out:
On Error Resume Next
Exit Function
error:
Call HandleError("Get_MegaMUD_RoomHash")
Resume out:
End Function

Function in_array_long_md(ByRef SearchArray() As Long, ByVal nFindValue As Long, _
    Optional ByVal nIndexDimension1 As Integer = -32000, Optional ByVal nNotIndexDimension1 As Integer = -32000, _
    Optional ByVal nIndexDimension2 As Integer = -32000, Optional ByVal nNotIndexDimension2 As Integer = -32000) As Boolean
Dim x As Long, y As Long
On Error GoTo error:

For x = LBound(SearchArray(), 1) To UBound(SearchArray(), 1)
    If nIndexDimension1 <> -32000 And x <> nIndexDimension1 Then GoTo skipx:
    If nNotIndexDimension1 <> -32000 And x = nNotIndexDimension1 Then GoTo skipx:
    For y = LBound(SearchArray(), 2) To UBound(SearchArray(), 2)
        If nIndexDimension2 <> -32000 And y <> nIndexDimension2 Then GoTo skipy:
        If nNotIndexDimension2 <> -32000 And y = nNotIndexDimension2 Then GoTo skipy:
        If SearchArray(x, y) = nFindValue Then
            in_array_long_md = True
            Exit Function
        End If
skipy:
    Next y
skipx:
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("in_array_long_md")
Resume out:
End Function

Function in_array_long_3d(ByRef SearchArray() As Long, ByVal nFindValue As Long, _
    Optional ByVal nIndexDimension1 As Integer = -32000, Optional ByVal nNotIndexDimension1 As Integer = -32000, _
    Optional ByVal nIndexDimension2 As Integer = -32000, Optional ByVal nNotIndexDimension2 As Integer = -32000, _
    Optional ByVal nIndexDimension3 As Integer = -32000, Optional ByVal nNotIndexDimension3 As Integer = -32000) As Boolean
Dim x As Long, y As Long, z As Long
On Error GoTo error:

For x = LBound(SearchArray(), 1) To UBound(SearchArray(), 1)
    If nIndexDimension1 <> -32000 And x <> nIndexDimension1 Then GoTo skipx:
    If nNotIndexDimension1 <> -32000 And x = nNotIndexDimension1 Then GoTo skipx:
    For y = LBound(SearchArray(), 2) To UBound(SearchArray(), 2)
        If nIndexDimension2 <> -32000 And y <> nIndexDimension2 Then GoTo skipy:
        If nNotIndexDimension2 <> -32000 And y = nNotIndexDimension2 Then GoTo skipy:
        For z = LBound(SearchArray(), 3) To UBound(SearchArray(), 3)
            If nIndexDimension3 <> -32000 And y <> nIndexDimension3 Then GoTo skipz:
            If nNotIndexDimension3 <> -32000 And y = nNotIndexDimension3 Then GoTo skipz:
            If SearchArray(x, y, z) = nFindValue Then
                in_array_long_3d = True
                Exit Function
            End If
skipz:
        Next z
skipy:
    Next y
skipx:
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("in_array_long_3d")
Resume out:
End Function

Function getval_array_long_3d(ByRef SearchArray() As Long, ByVal nFindValue As Long, _
    ByRef nReturnValueByRef As Long, Optional ByVal nReturnIndex = 1, Optional ByVal nMatchReturnValue As Integer = -32000, _
    Optional ByVal nIndexDimension1 As Integer = -32000, Optional ByVal nNotIndexDimension1 As Integer = -32000, _
    Optional ByVal nIndexDimension2 As Integer = -32000, Optional ByVal nNotIndexDimension2 As Integer = -32000, _
    Optional ByVal nIndexDimension3 As Integer = -32000, Optional ByVal nNotIndexDimension3 As Integer = -32000) As Boolean
Dim x As Long, y As Long, z As Long
On Error GoTo error:
nReturnValueByRef = 0

For x = UBound(SearchArray(), 1) To LBound(SearchArray(), 1) Step -1
    If nIndexDimension1 <> -32000 And x <> nIndexDimension1 Then GoTo skipx:
    If nNotIndexDimension1 <> -32000 And x = nNotIndexDimension1 Then GoTo skipx:
    
    For y = UBound(SearchArray(), 2) To LBound(SearchArray(), 2) Step -1
        If nIndexDimension2 <> -32000 And y <> nIndexDimension2 Then GoTo skipy:
        If nNotIndexDimension2 <> -32000 And y = nNotIndexDimension2 Then GoTo skipy:
        
        For z = UBound(SearchArray(), 3) To LBound(SearchArray(), 3) Step -1
            If nIndexDimension3 <> -32000 And z <> nIndexDimension3 Then GoTo skipz:
            If nNotIndexDimension3 <> -32000 And z = nNotIndexDimension3 Then GoTo skipz:
            
            If SearchArray(x, y, z) = nFindValue Then
                
                If nMatchReturnValue <> -32000 And SearchArray(x, y, nReturnIndex) <> nMatchReturnValue Then GoTo skipz:
                
                nReturnValueByRef = SearchArray(x, y, nReturnIndex)
                getval_array_long_3d = True
                Exit Function
            End If
skipz:
        Next z
skipy:
    Next y
skipx:
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("getval_array_long_3d")
Resume out:
End Function

Function getindex_array_long_3d(ByRef SearchArray() As Long, ByVal nFindValue As Long, _
    ByRef nReturnIndexByRef As Long, ByVal nReturnDimension As Long, _
    Optional ByVal nIndexDimension1 As Integer = -32000, Optional ByVal nNotIndexDimension1 As Integer = -32000, _
    Optional ByVal nIndexDimension2 As Integer = -32000, Optional ByVal nNotIndexDimension2 As Integer = -32000, _
    Optional ByVal nIndexDimension3 As Integer = -32000, Optional ByVal nNotIndexDimension3 As Integer = -32000) As Boolean
Dim x As Long, y As Long, z As Long
On Error GoTo error:
nReturnIndexByRef = 0

For x = LBound(SearchArray(), 1) To UBound(SearchArray(), 1)
    If nIndexDimension1 <> -32000 And x <> nIndexDimension1 Then GoTo skipx:
    If nNotIndexDimension1 <> -32000 And x = nNotIndexDimension1 Then GoTo skipx:
    
    For y = LBound(SearchArray(), 2) To UBound(SearchArray(), 2)
        If nIndexDimension2 <> -32000 And y <> nIndexDimension2 Then GoTo skipy:
        If nNotIndexDimension2 <> -32000 And y = nNotIndexDimension2 Then GoTo skipy:
        
        For z = LBound(SearchArray(), 3) To UBound(SearchArray(), 3)
            If nIndexDimension3 <> -32000 And z <> nIndexDimension3 Then GoTo skipz:
            If nNotIndexDimension3 <> -32000 And z = nNotIndexDimension3 Then GoTo skipz:
            
            If SearchArray(x, y, z) = nFindValue Then
                
                'If nMatchReturnValue <> -32000 And SearchArray(x, y, nReturnIndex) <> nMatchReturnValue Then GoTo skipz:
                
                Select Case nReturnDimension
                    Case 1:
                        nReturnIndexByRef = x
                        getindex_array_long_3d = True
                    Case 2:
                        nReturnIndexByRef = y
                        getindex_array_long_3d = True
                    Case 3:
                        nReturnIndexByRef = z
                        getindex_array_long_3d = True
                End Select
                
                Exit Function
            End If
skipz:
        Next z
skipy:
    Next y
skipx:
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("getindex_array_long_3d")
Resume out:
End Function

Function GetAbilDiffText(nAbilNumber, ByVal nValue1 As Long, ByVal nValue2 As Long, Optional ByVal sCurrentText As String, _
    Optional ByVal nPercent As Integer, Optional ByVal bOppositeMath As Boolean) As String
On Error GoTo error:

Dim nValue As Long
If bOppositeMath Then
    nValue = nValue2 - nValue1
Else
    nValue = nValue1 - nValue2
End If

'59 and 43 currently coded so that nValue1 and nValue2 would always be equal when reaching this function

Select Case nAbilNumber
    'Case 135: 'minlvl
    '    GetAbilDiffText = "Min LVL: " & nValue1
    '    If nValue2 > 0 Then GetAbilDiffText = GetAbilDiffText & " vs " & nValue2
    Case 59: 'class ok
        GetAbilDiffText = GetClassName(nValue1)
        
    Case 43: 'casts spell
        If nValue1 > 0 Then
            GetAbilDiffText = "[" & GetSpellName(nValue1, bHideRecordNumbers) & ", " & PullSpellEQ(True, 0, nValue1)
            If Not nPercent = 0 Then
                GetAbilDiffText = GetAbilDiffText & ", " & nPercent & "%]"
            Else
                GetAbilDiffText = GetAbilDiffText & "]"
            End If
        End If
        'GetAbilDiffText = GetSpellName(nValue1, bHideRecordNumbers)
        
    Case 114: '%spell
                
    Case Else:
        GetAbilDiffText = GetAbilityStats(nAbilNumber, nValue)
        If nAbilNumber = 135 Then GetAbilDiffText = GetAbilDiffText & " (" & IIf(bOppositeMath, nValue2, nValue1) & ")"
End Select

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetAbilDiffText")
Resume out:
End Function

Function ExtractRoomActions(ByVal sExit As String) As String
On Error GoTo error:

ExtractRoomActions = sExit
If InStr(1, ExtractRoomActions, ":", vbTextCompare) > 0 Then
    ExtractRoomActions = Trim(Mid(ExtractRoomActions, InStr(1, ExtractRoomActions, ":", vbTextCompare) + 1, 999))
End If


out:
On Error Resume Next
Exit Function
error:
Call HandleError("ExtractRoomActions")
Resume out:
End Function
