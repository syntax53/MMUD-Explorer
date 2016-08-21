VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExpCalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exp Calculator"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   Icon            =   "frmExpCalc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   4215
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtEndLVL 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3600
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "255"
      Top             =   720
      Width           =   555
   End
   Begin VB.TextBox txtStartLVL 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2760
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "2"
      Top             =   720
      Width           =   555
   End
   Begin VB.ComboBox cmbRace 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   300
      Width           =   1515
   End
   Begin VB.ComboBox cmbClass 
      Height          =   315
      Left            =   60
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   300
      Width           =   1635
   End
   Begin VB.TextBox txtCalcEXPTable 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdCalcExp 
      Caption         =   "&Calc."
      Height          =   555
      Left            =   3420
      TabIndex        =   2
      Top             =   60
      Width           =   735
   End
   Begin MSComctlLib.ListView lvCalcExp 
      Height          =   3555
      Left            =   60
      TabIndex        =   8
      Top             =   1020
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   6271
      View            =   3
      LabelEdit       =   1
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
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "to"
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
      Left            =   3375
      TabIndex        =   11
      Top             =   780
      Width           =   195
   End
   Begin VB.Label Label4 
      Caption         =   "LVL Range:"
      Height          =   195
      Left            =   1800
      TabIndex        =   10
      Top             =   780
      Width           =   915
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Race"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   60
      Width           =   1515
   End
   Begin VB.Label Label1 
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
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1635
   End
   Begin VB.Label Label39 
      Caption         =   "Exp Table %:"
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   780
      Width           =   1035
   End
End
Attribute VB_Name = "frmExpCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Sub Form_Load()
Dim x As Integer, sSectionName As String
On Error GoTo Error:

cmbClass.clear
If Not tabClasses.RecordCount = 0 Then
    tabClasses.MoveFirst
    Do While Not tabClasses.EOF
        cmbClass.AddItem tabClasses.Fields("Name")
        cmbClass.ItemData(cmbClass.NewIndex) = tabClasses.Fields("Number")
        tabClasses.MoveNext
    Loop
End If
cmbClass.AddItem "custom", 0

If frmMain.cmbGlobalClass(0).ListIndex > 0 Then
    For x = 0 To frmMain.cmbGlobalClass(0).ListCount - 1
        If frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex) = cmbClass.ItemData(x) Then
            cmbClass.ListIndex = x
            Exit For
        End If
    Next
End If

cmbRace.clear
If Not tabRaces.RecordCount = 0 Then
    tabRaces.MoveFirst
    Do While Not tabRaces.EOF
        cmbRace.AddItem tabRaces.Fields("Name")
        cmbRace.ItemData(cmbRace.NewIndex) = tabRaces.Fields("Number")
        tabRaces.MoveNext
    Loop
End If
cmbRace.AddItem "custom", 0

If frmMain.cmbGlobalRace(0).ListIndex > 0 Then
    For x = 0 To frmMain.cmbGlobalRace(0).ListCount - 1
        If frmMain.cmbGlobalRace(0).ItemData(frmMain.cmbGlobalRace(0).ListIndex) = cmbRace.ItemData(x) Then
            cmbRace.ListIndex = x
            Exit For
        End If
    Next
End If

If cmbRace.ListIndex < 0 Then cmbRace.ListIndex = 0
If cmbClass.ListIndex < 0 Then cmbClass.ListIndex = 0

sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ")
txtStartLVL.Text = ReadINI(sSectionName, "ExpCalcStartLevel")
txtEndLVL.Text = ReadINI(sSectionName, "ExpCalcEndLevel")
If Val(txtEndLVL.Text) < 10 Then txtEndLVL.Text = 255

Exit Sub
Error:
Call HandleError("LoadExpCalc")
Resume Next
End Sub

Private Sub CalcExp()
Dim nClassExp As Integer, nRaceExp As Integer

On Error GoTo Error:

If cmbClass.ListIndex > 0 Then
    tabClasses.Index = "pkClasses"
    tabClasses.Seek "=", cmbClass.ItemData(cmbClass.ListIndex)
    If tabClasses.NoMatch = True Then
        nClassExp = 0
    Else
        nClassExp = tabClasses.Fields("ExpTable") + 100
    End If
End If

If cmbRace.ListIndex > 0 Then
    tabRaces.Index = "pkRaces"
    tabRaces.Seek "=", cmbRace.ItemData(cmbRace.ListIndex)
    If tabRaces.NoMatch = True Then
        nRaceExp = 0
    Else
        nRaceExp = tabRaces.Fields("ExpTable")
    End If
End If

txtCalcEXPTable.Text = nClassExp + nRaceExp

Exit Sub
Error:
Call HandleError
End Sub

Private Sub cmbClass_Click()
Call CalcExp
End Sub

Private Sub cmbRace_Click()
Call CalcExp
End Sub

Private Sub cmdCalcExp_Click()
Dim sExp As String, nExp As Currency, x As Long
Dim oLI As ListItem, nLastExp As Currency

On Error GoTo Error:

lvCalcExp.ListItems.clear
lvCalcExp.ColumnHeaders.clear
lvCalcExp.ColumnHeaders.Add , , "LVL", 500
lvCalcExp.ColumnHeaders.Add , , "Experience", 1600
lvCalcExp.ColumnHeaders.Add , , "Needed", 1400

If Val(txtStartLVL.Text) < 2 Then
    txtStartLVL.Text = 2
ElseIf Val(txtStartLVL.Text) > 500 Then
    txtStartLVL.Text = 500
End If

If Val(txtEndLVL.Text) < 10 Then
    txtEndLVL.Text = 10
ElseIf Val(txtEndLVL.Text) > 999 Then
    txtEndLVL.Text = 999
End If

For x = Val(txtStartLVL.Text) To Val(txtEndLVL.Text)
    nExp = CalcExpNeeded(x, CLng(txtCalcEXPTable.Text))
    sExp = CStr(nExp * 10000)
    
    Set oLI = lvCalcExp.ListItems.Add()
    oLI.Text = x
    oLI.SubItems(1) = PutCommas(sExp)
    oLI.SubItems(2) = PutCommas(Val(sExp) - nLastExp)

    nLastExp = Val(sExp)
    Set oLI = Nothing
Next

Exit Sub

Error:
Call HandleError

End Sub



Private Sub Form_Unload(Cancel As Integer)
Dim sSectionName As String
sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ")
Call WriteINI(sSectionName, "ExpCalcStartLevel", Val(txtStartLVL.Text))
Call WriteINI(sSectionName, "ExpCalcEndLevel", Val(txtEndLVL.Text))
End Sub

Private Sub lvCalcExp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then Call frmMain.PopUpAuxMenu(lvCalcExp)
End Sub

Private Sub txtCalcEXPTable_GotFocus()
Call SelectAll(txtCalcEXPTable)
End Sub

Private Sub txtCalcEXPTable_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtCalcEXPTable_KeyUp(KeyCode As Integer, Shift As Integer)
Dim sStr As String, nPos As Integer, nSel As Integer

On Error GoTo Error:

nPos = txtCalcEXPTable.SelStart
sStr = txtCalcEXPTable.Text
nSel = txtCalcEXPTable.SelLength

cmbClass.ListIndex = 0
cmbRace.ListIndex = 0
txtCalcEXPTable.Text = sStr
txtCalcEXPTable.SelStart = nPos
txtCalcEXPTable.SelLength = nSel

Exit Sub

Error:
Call HandleError

End Sub

Private Sub txtEndLVL_GotFocus()
Call SelectAll(txtEndLVL)
End Sub

Private Sub txtEndLVL_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtStartLVL_GotFocus()
Call SelectAll(txtStartLVL)
End Sub

Private Sub txtStartLVL_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub
