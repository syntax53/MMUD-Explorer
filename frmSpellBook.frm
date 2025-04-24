VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmSpellBook 
   Caption         =   "Spell Book"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5175
   Icon            =   "frmSpellBook.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   5175
   Begin VB.Timer timWindowMove 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox txtLevel 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1140
      MaxLength       =   3
      TabIndex        =   11
      Text            =   "999"
      Top             =   780
      Width           =   555
   End
   Begin VB.ComboBox cmbAlignment 
      Height          =   315
      Left            =   3900
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   300
      Width           =   1155
   End
   Begin VB.CommandButton cmdListSpells 
      Caption         =   "&List Spells"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3900
      TabIndex        =   8
      Top             =   720
      Width           =   1155
   End
   Begin VB.ComboBox cmbSpellMagery 
      Height          =   315
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   300
      Width           =   1035
   End
   Begin VB.ComboBox cmbSpellMageryLevel 
      Height          =   315
      ItemData        =   "frmSpellBook.frx":0CCA
      Left            =   2940
      List            =   "frmSpellBook.frx":0CCC
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   300
      Width           =   795
   End
   Begin VB.ComboBox cmbClass 
      Height          =   315
      Left            =   60
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   300
      Width           =   1635
   End
   Begin VB.CommandButton cmdPasteChar 
      Caption         =   "&Paste Character"
      Height          =   375
      Left            =   1860
      TabIndex        =   0
      Top             =   720
      Width           =   1875
   End
   Begin MSComctlLib.ListView lvSpellBook 
      Height          =   4095
      Left            =   60
      TabIndex        =   2
      Top             =   1200
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label4 
      Caption         =   "Char LVL:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   915
   End
   Begin VB.Label lblLabelArray 
      Caption         =   "Align:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   3900
      TabIndex        =   10
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   60
      Width           =   1635
   End
   Begin VB.Label lblLabelArray 
      Caption         =   "Magery:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   1860
      TabIndex        =   6
      Top             =   60
      Width           =   735
   End
   Begin VB.Label lblLabelArray 
      Caption         =   "Level:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   2940
      TabIndex        =   5
      Top             =   60
      Width           =   735
   End
   Begin VB.Menu mnuSpellsPopUp 
      Caption         =   "SpellPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuSpellsPopUpItem 
         Caption         =   "Add to Compare"
         Index           =   0
      End
      Begin VB.Menu mnuSpellsPopUpItem 
         Caption         =   "Copy Details to Clipboard"
         Index           =   1
      End
      Begin VB.Menu mnuSpellsPopUpItem 
         Caption         =   "Copy Name(s) to Clipboard"
         Index           =   2
      End
      Begin VB.Menu mnuSpellsPopUpItem 
         Caption         =   "What casts this spell?"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmSpellBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Dim tWindowSize As WindowSizeRestrictions
Dim bSortOrderAsc As Boolean
Dim oLastColumnSorted As ColumnHeader
Dim nLastSpellSort As Integer
Dim bKeepSortOrder As Boolean
Dim nMagicLVL As Integer
Dim nMagery As Integer

Public nLastPosTop As Long
Public nLastPosLeft As Long
Public nLastPosMoved As Long
Public nLastPosMonitor As Long

Public nLastTimerTop As Long
Public nLastTimerLeft As Long

Private Sub cmdListSpells_Click()
On Error GoTo error:
Dim oLI As ListItem, x As Integer, nAlign As Integer, nNotAlign As Integer
Dim bFiltered As Boolean, bHasAbility As Boolean

If tabSpells.RecordCount = 0 Then Exit Sub

lvSpellBook.ListItems.clear
DoEvents

'97, good
'98, evil
'112, neutral
'110, not good
'111, not evil
'113, not neutral

'0 "User"
'1 "Self"
'2 "Self or User"
'3 "Divided Area (not self)"
'4 "Monster"
'5 "Divided Area (incl self)"
'6 "Any"
'7 "Item"
'8 "Monster or User"
'9 "Divided Attack Area"
'10 "Divided Party Area"
'11 "Full Area"
'12 "Full Attack Area"
'13 "Full Party Area"

tabSpells.MoveFirst
Do Until tabSpells.EOF
    bHasAbility = False
    nAlign = 0
    nNotAlign = 0
    
'    If tabSpells.Fields("Name") = "form of the crane" Then
'        Debug.Print 1
'    End If
    
    If bOnlyInGame Then
        'tabSpells.Fields("Magery") = 5 = kai
        If tabSpells.Fields("Learnable") = 0 And Len(tabSpells.Fields("Learned From")) <= 1 And Len(tabSpells.Fields("Casted By")) <= 1 _
            And ( _
                    tabSpells.Fields("Magery") <> 5 _
                    Or (tabSpells.Fields("Magery") = 5 And tabSpells.Fields("ReqLevel") < 1) _
                    Or (tabSpells.Fields("Magery") = 5 And bDisableKaiAutolearn) _
                ) Then
                
                If nNMRVer >= 1.8 Then
                    If Len(tabSpells.Fields("Classes")) <= 1 Then GoTo skip:
                Else
                    GoTo skip:
                End If
        End If
    End If
    
    If Not cmbSpellMagery.ListIndex = 0 Then
        If Not cmbSpellMagery.ListIndex = tabSpells.Fields("Magery") Then
            If tabSpells.Fields("Learnable") > 0 _
                And tabSpells.Fields("Magery") = 0 _
                And nNMRVer >= 1.7 Then
                
                If tabSpells.Fields("Classes") = "(*)" _
                    Or InStr(1, tabSpells.Fields("Classes"), _
                        "(" & cmbClass.ItemData(cmbClass.ListIndex) & ")", vbTextCompare) > 0 Then
                    GoTo skip_magery_check:
                Else
                    GoTo skip:
                End If
            Else
                GoTo skip:
            End If
        End If
    End If
    
    If Not cmbSpellMageryLevel.ListIndex = 0 Then
        If cmbSpellMageryLevel.ListIndex < tabSpells.Fields("MageryLVL") Then GoTo skip:
    End If

    'magery 5 is kai
    If Not cmbSpellMagery.ListIndex = 5 And tabSpells.Fields("Learnable") = 0 Then GoTo skip:
    If cmbSpellMagery.ListIndex = 5 And bDisableKaiAutolearn And tabSpells.Fields("Learnable") = 0 Then GoTo skip:
    
skip_magery_check:
    
    If nNMRVer >= 1.7 And cmbClass.ListIndex > 0 Then
        If Len(tabSpells.Fields("Classes")) > 2 And Not tabSpells.Fields("Classes") = "(*)" Then
            If Not InStr(1, tabSpells.Fields("Classes"), _
                "(" & cmbClass.ItemData(cmbClass.ListIndex) & ")", vbTextCompare) > 0 Then GoTo skip:
        End If
    End If
    
    If Val(txtLevel.Text) < tabSpells.Fields("ReqLevel") Then GoTo skip:
    
    For x = 0 To 9
        Select Case tabSpells.Fields("Abil-" & x)
            Case 0:
                
            Case 97, 98, 112: 'good/evil/neutral abils
                nAlign = tabSpells.Fields("Abil-" & x)
                Select Case cmbAlignment.ListIndex
                    Case 0:
                    Case 1: 'good
                        If Not nAlign = 97 Then GoTo skip:
                    Case 2: 'netural
                        If Not nAlign = 112 Then GoTo skip:
                    Case 3: 'evil
                        If Not nAlign = 98 Then GoTo skip:
                End Select
        
            Case 110, 111, 113: 'notgood/notevil/notneutral abils
                nNotAlign = tabSpells.Fields("Abil-" & x)
                Select Case cmbAlignment.ListIndex
                    Case 0:
                    Case 1: 'good
                        If nNotAlign = 110 Then GoTo skip:
                    Case 2: 'netural
                        If nNotAlign = 113 Then GoTo skip:
                    Case 3: 'evil
                        If nNotAlign = 111 Then GoTo skip:
                End Select

        End Select
    Next x
    
    Call AddSpell2LV(lvSpellBook)

GoTo MoveNext:
skip:
bFiltered = True
MoveNext:
    tabSpells.MoveNext
    'DoEvents
Loop

For Each oLI In lvSpellBook.ListItems
    oLI.Selected = False
Next

bKeepSortOrder = True
Call lvSpellBook_ColumnClick(lvSpellBook.ColumnHeaders(5))

If lvSpellBook.ListItems.Count >= 1 Then Call lvSpellBook_ItemClick(lvSpellBook.ListItems(1))

lvSpellBook.Refresh
DoEvents
out:
On Error Resume Next
Call frmMain.RefreshLearnedSpellColors_byLV(lvSpellBook)
Exit Sub
error:
Call HandleError("cmdListSpells_Click")
Resume out:

End Sub

Private Sub Form_Load()
On Error GoTo error:
Dim x As Integer, sSectionName As String, nTemp As Long

tWindowSize.twpMinWidth = 5175
tWindowSize.twpMinHeight = 5355
Call SubclassFormMinMaxSize(Me, tWindowSize)

'Me.Width = Val(ReadINI("Settings", "SpellbookWidth", , 5415))
'Me.Height = Val(ReadINI("Settings", "SpellbookHeight", , 8040))
Call ResizeForm(Me, Val(ReadINI("Settings", "SpellbookWidth", , 5400)), Val(ReadINI("Settings", "SpellbookHeight", , 7875)))

nLastSpellSort = 2
nMagicLVL = 3
nMagery = 0

cmbSpellMagery.clear
cmbSpellMagery.AddItem "Any", 0
cmbSpellMagery.AddItem "Mage", 1
cmbSpellMagery.AddItem "Priest", 2
cmbSpellMagery.AddItem "Druid", 3
cmbSpellMagery.AddItem "Bard", 4
cmbSpellMagery.AddItem "Kai", 5
cmbSpellMagery.ListIndex = 0

cmbSpellMageryLevel.clear
cmbSpellMageryLevel.AddItem "Any", 0
cmbSpellMageryLevel.AddItem "1", 1
cmbSpellMageryLevel.AddItem "2", 2
cmbSpellMageryLevel.AddItem "3", 3
cmbSpellMageryLevel.ListIndex = 0

cmbAlignment.clear
cmbAlignment.AddItem "Any"
cmbAlignment.AddItem "Good"
cmbAlignment.AddItem "Neutral"
cmbAlignment.AddItem "Evil"
cmbAlignment.ListIndex = 0

cmbClass.clear
If Not tabClasses.RecordCount = 0 Then
    tabClasses.MoveFirst
    Do While Not tabClasses.EOF
        cmbClass.AddItem tabClasses.Fields("Name")
        cmbClass.ItemData(cmbClass.NewIndex) = tabClasses.Fields("Number")
        tabClasses.MoveNext
    Loop
End If
cmbClass.AddItem "none", 0

Call CopyGlobalChar

If cmbClass.ListIndex < 0 Then cmbClass.ListIndex = 0

'Level Mana Short Spell Name

lvSpellBook.ColumnHeaders.clear
lvSpellBook.ColumnHeaders.Add 1, "Number", "#", 600, lvwColumnLeft
lvSpellBook.ColumnHeaders.Add 2, "Name", "Name", 2000, lvwColumnCenter
lvSpellBook.ColumnHeaders.Add 3, "Short", "Short", 650, lvwColumnCenter
lvSpellBook.ColumnHeaders.Add 4, "Magery", "Magery", 900, lvwColumnCenter
lvSpellBook.ColumnHeaders.Add 5, "LVL", "LVL", 500, lvwColumnCenter
lvSpellBook.ColumnHeaders.Add 6, "Mana", "Mana", 650, lvwColumnCenter
lvSpellBook.ColumnHeaders.Add 7, "Diff", "Diff", 500, lvwColumnCenter
lvSpellBook.ColumnHeaders.Add 8, "Dmg", "Dmg", 700, lvwColumnCenter
lvSpellBook.ColumnHeaders.Add 9, "Dmg/M", "Dmg/M", 900, lvwColumnCenter
lvSpellBook.ColumnHeaders.Add 10, "Detail", "Detail", 8000, lvwColumnLeft


sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ")
'txtLevel.Text = ReadINI(sSectionName, "ExpCalcStartLevel")
'txtEndLVL.Text = ReadINI(sSectionName, "ExpCalcEndLevel")
'If Val(txtEndLVL.Text) < 10 Then txtEndLVL.Text = 255

If cmbClass.ListIndex > 0 Then Call cmdListSpells_Click

nTemp = Val(ReadINI("Settings", "SpellbookTop"))
If nTemp = 0 Then
    If frmMain.WindowState = vbMinimized Then
        nTemp = (Screen.Height - Me.Height) / 2
    Else
        nTemp = frmMain.Top + ((frmMain.Height - Me.Height) / 2)
    End If
End If
Me.Top = nTemp

nTemp = Val(ReadINI("Settings", "SpellbookLeft"))
If nTemp = 0 Then
    If frmMain.WindowState = vbMinimized Then
        nTemp = (Screen.Width - Me.Width) / 2
    Else
        nTemp = frmMain.Left + ((frmMain.Width - Me.Width) / 2)
    End If
End If
Me.Left = nTemp

timWindowMove.Enabled = True

Exit Sub
error:
Call HandleError("LoadExpCalc")
Resume Next
End Sub

Private Sub CopyGlobalChar()
Dim x As Integer

If frmMain.chkGlobalFilter.Value = 1 Then
    If frmMain.cmbGlobalAlignment.ListIndex > 0 Then cmbAlignment.ListIndex = frmMain.cmbGlobalAlignment.ListIndex
    If Val(frmMain.txtGlobalLevel(1).Text) > 0 Then txtLevel.Text = Val(frmMain.txtGlobalLevel(1).Text)
    
    If frmMain.cmbGlobalClass(1).ListIndex > 0 Then
        For x = 0 To frmMain.cmbGlobalClass(1).ListCount - 1
            If frmMain.cmbGlobalClass(1).ItemData(frmMain.cmbGlobalClass(0).ListIndex) = cmbClass.ItemData(x) Then
                cmbClass.ListIndex = x
                Exit For
            End If
        Next
    End If
End If

End Sub

Private Sub cmbClass_Click()
On Error GoTo error:

If cmbClass.ItemData(cmbClass.ListIndex) < 1 Then GoTo out:

tabClasses.Index = "pkClasses"
tabClasses.Seek "=", cmbClass.ItemData(cmbClass.ListIndex)
If tabClasses.NoMatch = True Then
    MsgBox "Class not found.", vbInformation + vbOKOnly
    GoTo out:
End If

nMagicLVL = tabClasses.Fields("MageryLVL")
If nMagicLVL > 3 Then nMagicLVL = 3

nMagery = tabClasses.Fields("MageryType")

If nMagicLVL < cmbSpellMageryLevel.ListCount Then cmbSpellMageryLevel.ListIndex = nMagicLVL
If nMagery < cmbSpellMagery.ListCount Then cmbSpellMagery.ListIndex = nMagery

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("cmbClass_Click")
Resume out:

End Sub


Private Sub cmdPasteChar_Click()
Call frmMain.PasteCharacter
Call CopyGlobalChar
Call cmdListSpells_Click
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = vbMinimized Then Exit Sub

lvSpellBook.Width = Me.Width - 350
lvSpellBook.Height = Me.Height - TITLEBAR_OFFSET - 1825
'CheckPosition Me

'Debug.Print Me.Height
'Debug.Print Me.Width

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo error:

If Not Me.WindowState = vbMinimized And Not Me.WindowState = vbMaximized Then
    Call WriteINI("Settings", "SpellbookTop", Me.Top)
    Call WriteINI("Settings", "SpellbookLeft", Me.Left)
    Call WriteINI("Settings", "SpellbookWidth", Me.Width)
    Call WriteINI("Settings", "SpellbookHeight", Me.Height)
End If

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("Form_Unload")
Resume out:
End Sub

Private Sub lvSpellBook_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim bSort As Boolean, nSort As ListDataType

On Error GoTo error:

nLastSpellSort = ColumnHeader.Index
If bKeepSortOrder Then
    bSort = IIf(lvSpellBook.SortOrder = lvwDescending, False, True)
    bKeepSortOrder = False
Else
    If oLastColumnSorted Is ColumnHeader Then
        If bSortOrderAsc = True Then
            bSortOrderAsc = False
        Else
            bSortOrderAsc = True
        End If
    End If
    bSort = bSortOrderAsc
    Set oLastColumnSorted = ColumnHeader
End If

If ColumnHeader.Index = 2 Or ColumnHeader.Index = 3 Or ColumnHeader.Index = 4 Then
    nSort = ldtstring
Else
    nSort = ldtnumber
End If

SortListView lvSpellBook, ColumnHeader.Index, nSort, bSort

Exit Sub

error:
Call HandleError("lvSpellBook_ColumnClick")
End Sub

Private Sub lvSpellBook_DblClick()
If lvSpellBook.ListItems.Count = 0 Then Exit Sub
If lvSpellBook.SelectedItem Is Nothing Then Exit Sub
Call frmMain.GotoSpell(Val(lvSpellBook.SelectedItem.Text))
End Sub

Private Sub lvSpellBook_ItemClick(ByVal item As MSComctlLib.ListItem)

Set lvSpellBook.SelectedItem = item
'Call PullSpellDetail(Val(Item.Text), txtSpellDetail, lvSpellLoc)

item.Selected = True
item.EnsureVisible

End Sub

Private Sub lvSpellBook_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then Call frmMain.PopUpSpellsMenu(lvSpellBook)
End Sub


Private Sub timWindowMove_Timer()
Call MonitorFormTimer(Me)
End Sub

Private Sub txtLevel_GotFocus()
Call SelectAll(txtLevel)
End Sub
