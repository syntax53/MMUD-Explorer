VERSION 5.00
Begin VB.Form frmPopUpOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MMUD Explorer"
   ClientHeight    =   4290
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   7035
   ControlBox      =   0   'False
   Icon            =   "frmPopUpOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7035
   Begin VB.Frame fraChooseAttack 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3855
      Left            =   60
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   6915
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "With the updates I have coming out in the  "
         Height          =   3495
         Left            =   240
         TabIndex        =   4
         Top             =   180
         Width           =   6435
         Begin VB.OptionButton optAttackType 
            Caption         =   "None (Assume 1-shot Everything)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   360
            TabIndex        =   6
            Top             =   360
            Value           =   -1  'True
            Width           =   4335
         End
         Begin VB.TextBox txtAttackMag 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   300
            Left            =   5340
            TabIndex        =   22
            Text            =   "999"
            Top             =   3060
            Width           =   735
         End
         Begin VB.TextBox txtAttackPhys 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   300
            Left            =   3840
            TabIndex        =   20
            Text            =   "999"
            Top             =   3060
            Width           =   735
         End
         Begin VB.ComboBox cmbAttackMA 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmPopUpOptions.frx":0CCA
            Left            =   2700
            List            =   "frmPopUpOptions.frx":0CDA
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   2630
            Width           =   1515
         End
         Begin VB.TextBox txtAttackSpellLevel 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   300
            Left            =   4710
            TabIndex        =   15
            Text            =   "999"
            Top             =   2220
            Width           =   615
         End
         Begin VB.ComboBox cmbAttackSpell 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            ItemData        =   "frmPopUpOptions.frx":0D00
            Left            =   660
            List            =   "frmPopUpOptions.frx":0D02
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   2220
            Width           =   3555
         End
         Begin VB.ComboBox cmbAttackSpell 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            ItemData        =   "frmPopUpOptions.frx":0D04
            Left            =   660
            List            =   "frmPopUpOptions.frx":0D06
            Sorted          =   -1  'True
            TabIndex        =   11
            Top             =   1500
            Width           =   2295
         End
         Begin VB.CheckBox chkSmashing 
            Caption         =   "Smashing"
            Enabled         =   0   'False
            Height          =   240
            Left            =   3960
            TabIndex        =   9
            Top             =   810
            Width           =   1095
         End
         Begin VB.CheckBox chkBashing 
            Caption         =   "Bashing"
            Enabled         =   0   'False
            Height          =   240
            Left            =   2880
            TabIndex        =   8
            Top             =   810
            Width           =   975
         End
         Begin VB.OptionButton optAttackType 
            Caption         =   "Enter Values Manually:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   360
            TabIndex        =   18
            Top             =   3105
            Width           =   2775
         End
         Begin VB.OptionButton optAttackType 
            Caption         =   "Martial Arts Attack: "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   360
            TabIndex        =   16
            Top             =   2640
            Width           =   2295
         End
         Begin VB.OptionButton optAttackType 
            Caption         =   "Any Spell @ Specified Level:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   360
            TabIndex        =   12
            Top             =   1920
            Width           =   3975
         End
         Begin VB.OptionButton optAttackType 
            Caption         =   "A Learned Spell @ Current Level:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   360
            TabIndex        =   10
            Top             =   1200
            Width           =   4515
         End
         Begin VB.OptionButton optAttackType 
            Caption         =   "Equipped Weapon"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   360
            TabIndex        =   7
            Top             =   780
            Width           =   2415
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Magical:"
            Height          =   195
            Index           =   9
            Left            =   4620
            TabIndex        =   21
            Top             =   3120
            Width           =   675
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Physical:"
            Height          =   195
            Index           =   8
            Left            =   3060
            TabIndex        =   19
            Top             =   3120
            Width           =   735
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Caption         =   "@"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   4260
            TabIndex        =   14
            Top             =   2220
            Width           =   375
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Caption         =   "Choose Attack"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   315
            Index           =   7
            Left            =   120
            TabIndex        =   5
            Top             =   0
            Width           =   6195
         End
      End
   End
   Begin VB.Frame fraRoomFind 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3855
      Left            =   60
      TabIndex        =   23
      Top             =   360
      Visible         =   0   'False
      Width           =   6915
      Begin VB.Frame fraMISC 
         BorderStyle     =   0  'None
         Caption         =   "With the updates I have coming out in the  "
         Height          =   3495
         Left            =   240
         TabIndex        =   24
         Top             =   180
         Width           =   6435
         Begin VB.OptionButton optRoomFindMatch 
            Caption         =   "Exact Match"
            Height          =   240
            Index           =   1
            Left            =   4200
            TabIndex        =   29
            Top             =   720
            Width           =   1515
         End
         Begin VB.OptionButton optRoomFindMatch 
            Caption         =   "Partial Match"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   2520
            TabIndex        =   28
            Top             =   720
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.CommandButton cmdRoomFindDir 
            Caption         =   "D"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   9
            Left            =   4200
            MaskColor       =   &H80000016&
            Style           =   1  'Graphical
            TabIndex        =   41
            Tag             =   "0"
            Top             =   2700
            Width           =   615
         End
         Begin VB.CommandButton cmdRoomFindDir 
            Caption         =   "U"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   8
            Left            =   3000
            MaskColor       =   &H80000016&
            Style           =   1  'Graphical
            TabIndex        =   40
            Tag             =   "0"
            Top             =   2700
            Width           =   615
         End
         Begin VB.CommandButton cmdRoomFindDir 
            Caption         =   "SE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   6
            Left            =   4200
            MaskColor       =   &H80000016&
            Style           =   1  'Graphical
            TabIndex        =   39
            Tag             =   "0"
            Top             =   2220
            Width           =   615
         End
         Begin VB.CommandButton cmdRoomFindDir 
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   3600
            MaskColor       =   &H80000016&
            Style           =   1  'Graphical
            TabIndex        =   38
            Tag             =   "0"
            Top             =   2220
            Width           =   615
         End
         Begin VB.CommandButton cmdRoomFindDir 
            Caption         =   "SW"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   3000
            MaskColor       =   &H80000016&
            Style           =   1  'Graphical
            TabIndex        =   37
            Tag             =   "0"
            Top             =   2220
            Width           =   615
         End
         Begin VB.CommandButton cmdRoomFindDir 
            Caption         =   "E"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   4200
            MaskColor       =   &H80000016&
            Style           =   1  'Graphical
            TabIndex        =   36
            Tag             =   "0"
            Top             =   1740
            Width           =   615
         End
         Begin VB.CommandButton cmdRoomFindDir 
            Caption         =   "W"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   3000
            MaskColor       =   &H80000016&
            Style           =   1  'Graphical
            TabIndex        =   35
            Tag             =   "0"
            Top             =   1740
            Width           =   615
         End
         Begin VB.CommandButton cmdRoomFindDir 
            Caption         =   "NE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   4200
            MaskColor       =   &H80000016&
            Style           =   1  'Graphical
            TabIndex        =   33
            Tag             =   "0"
            Top             =   1260
            Width           =   615
         End
         Begin VB.CommandButton cmdRoomFindDir 
            Caption         =   "N"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   3600
            MaskColor       =   &H80000016&
            Style           =   1  'Graphical
            TabIndex        =   32
            Tag             =   "0"
            Top             =   1260
            Width           =   615
         End
         Begin VB.CommandButton cmdRoomFindDir 
            Caption         =   "NW"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   3000
            MaskColor       =   &H80000016&
            Style           =   1  'Graphical
            TabIndex        =   31
            Tag             =   "0"
            Top             =   1260
            Width           =   615
         End
         Begin VB.TextBox txtRoomName 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2520
            TabIndex        =   26
            Top             =   300
            Width           =   3015
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Do not include invisible, hidden, or activated exits"
            Height          =   495
            Index           =   3
            Left            =   240
            TabIndex        =   34
            Top             =   1620
            Width           =   2115
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Obvious Exits:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   660
            TabIndex        =   30
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "3 or more letters"
            Height          =   195
            Index           =   1
            Left            =   540
            TabIndex        =   27
            Top             =   600
            Width           =   1755
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Room Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   780
            TabIndex        =   25
            Top             =   300
            Width           =   1575
         End
      End
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "&Paste from Clipboard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2340
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5460
      TabIndex        =   2
      Top             =   0
      Width           =   1515
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Co&ntinue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox txtText 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   60
      MaxLength       =   10000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   42
      Top             =   360
      Visible         =   0   'False
      Width           =   6945
   End
End
Attribute VB_Name = "frmPopUpOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Dim tWindowSize As WindowSizeRestrictions

Private Sub chkBashing_Click()
If chkBashing.Value = 1 Then chkSmashing.Value = 0
If chkSmashing.Value = 1 Then chkBashing.Value = 0
End Sub

Private Sub chkSmashing_Click()
If chkSmashing.Value = 1 Then chkBashing.Value = 0
If chkBashing.Value = 1 Then chkSmashing.Value = 0
End Sub

Private Sub cmbAttackSpell_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = AutoComplete(cmbAttackSpell(Index), KeyAscii, False)
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
'Me.bPasteParty = False
'txtText.Visible = True
'fraRoomFind.Visible = False
'cmdPaste.Enabled = True
Me.Hide
End Sub

Private Sub cmdContinue_Click()
On Error GoTo error:


out:
On Error Resume Next
Me.Tag = "1"
'txtText.Visible = True
'fraRoomFind.Visible = False
'cmdPaste.Enabled = True
'txtText.SetFocus
Me.Hide
Exit Sub
error:
Call HandleError("cmdContinue_Click")
Resume out:
End Sub

Private Sub cmdPaste_Click()
Dim nYesNo As Integer

If Not Clipboard.GetText = "" Then
    If Not txtText.Text = "" And fraRoomFind.Visible = False Then
        nYesNo = MsgBox("Clear paste area first?", vbYesNo + vbDefaultButton1 + vbQuestion, "Clear?")
    Else
        nYesNo = vbYes
    End If
    
    If nYesNo = vbYes Then
        txtText.Text = Clipboard.GetText
    Else
        txtText.Text = Clipboard.GetText & vbCrLf & vbCrLf & txtText.Text
    End If
    
    'txtText.Visible = True
    'fraRoomFind.Visible = False
End If

End Sub

Public Sub SetupChooseAttack()
On Error GoTo error:
Dim x As Integer, y As Integer

fraChooseAttack.Visible = True
fraRoomFind.Visible = False

If cmbAttackSpell(1).ListCount < 5 Then
    Call RefreshAttackSpells
End If

If nCurrentAttackSpellNum > 0 And nCurrentAttackSpellLVL > 0 Then
    For y = 2 To 3
        For x = 0 To cmbAttackSpell(y - 2).ListCount - 1
            If cmbAttackSpell(y - 2).ItemData(x) = nCurrentAttackSpellNum Then
                cmbAttackSpell(y - 2).ListIndex = x
                Exit For
            End If
        Next x
    Next y
End If

If nCurrentAttackSpellLVL > 0 Then
    txtAttackSpellLevel.Text = nCurrentAttackSpellLVL
ElseIf frmMain.chkGlobalFilter.Value = 1 And Val(frmMain.txtGlobalLevel(0).Text) > 0 Then
    txtAttackSpellLevel.Text = Val(frmMain.txtGlobalLevel(0).Text)
End If

If nCurrentAttackMA > 0 And nCurrentAttackMA <= 3 Then cmbAttackMA.ListIndex = nCurrentAttackMA

If nCurrentAttackManualPhys > 0 Or nCurrentAttackManualMag > 0 Then
    txtAttackPhys.Text = nCurrentAttackManualPhys
    txtAttackMag.Text = nCurrentAttackManualMag
End If

If nCurrentAttackType = 6 Then chkBashing.Value = 1: chkSmashing.Value = 0
If nCurrentAttackType = 7 Then chkBashing.Value = 0: chkSmashing.Value = 1

If nCurrentAttackType = 1 Or nCurrentAttackType > 5 Then
    If optAttackType(1).Value = True Then
        Call optAttackType_Click(1)
    Else
        optAttackType(1).Value = True
    End If
Else
    If optAttackType(nCurrentAttackType).Value = True Then
        Call optAttackType_Click(nCurrentAttackType)
    Else
        optAttackType(nCurrentAttackType).Value = True
    End If
End If

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("SetupChooseAttack")
Resume out:

End Sub

Public Sub ResetRoomFind()
On Error GoTo error:
Dim x As Integer

fraChooseAttack.Visible = False

For x = 0 To cmdRoomFindDir.Count - 1
    cmdRoomFindDir(x).BackColor = &H8000000F
    cmdRoomFindDir(x).Tag = 0
Next x

Call optRoomFindMatch_Click(0)
txtRoomName.Text = ""

Me.Caption = "Find Room with Exits"
fraRoomFind.Visible = True

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("ResetRoomFind")
Resume out:

End Sub

Private Sub cmdRoomFindDir_Click(Index As Integer)
On Error GoTo error:

If cmdRoomFindDir(Index).Tag = "0" Then
    cmdRoomFindDir(Index).BackColor = &HC000&
    cmdRoomFindDir(Index).Tag = 1
Else
    cmdRoomFindDir(Index).BackColor = &H8000000F
    cmdRoomFindDir(Index).Tag = 0
End If

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("cmdRoomFindDir_Click")
Resume out:
End Sub

Private Sub Form_Load()
On Error GoTo error:

Call SetWindowLong(Me.hWnd, GWL_HWNDPARENT, 0)

'stop windows from resizing fixed-size windows when changing dpi
If bDPIAwareMode Then Call SubclassFormMinMaxSize(Me, tWindowSize, True)

If frmMain.WindowState = vbMinimized Then
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
Else
    Me.Left = frmMain.Left + ((frmMain.Width - Me.Width) / 2)
    Me.Top = frmMain.Top + ((frmMain.Height - Me.Height) / 2)
End If

cmbAttackMA.ListIndex = 0

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("Form_Load")
Resume out:
End Sub

Public Sub RefreshAttackSpells()
On Error GoTo error:
Dim x As Integer, bHasDmg As Boolean

cmbAttackSpell(0).clear
cmbAttackSpell(1).clear
If Not tabSpells.RecordCount = 0 Then
    tabSpells.MoveFirst
    Do While Not tabSpells.EOF
        
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
        
        If Len(tabSpells.Fields("Short")) > 1 Then
            bHasDmg = False
            For x = 0 To 9 Or bHasDmg
                Select Case tabSpells.Fields("Abil-" & x)
                    Case 1, 8, 17: '1-dmg, 8-drain, 17-dmg-mr
                        bHasDmg = True
                End Select
            Next x
            
            If bHasDmg Then
                If in_long_arr(tabSpells.Fields("Number"), nLearnedSpells()) Then
                    cmbAttackSpell(0).AddItem tabSpells.Fields("Name") & IIf(bHideRecordNumbers, "", " (" & tabSpells.Fields("Number") & ")")
                    cmbAttackSpell(0).ItemData(cmbAttackSpell(0).NewIndex) = tabSpells.Fields("Number")
                End If
                cmbAttackSpell(1).AddItem tabSpells.Fields("Name") & " (" & tabSpells.Fields("Number") & ") - LVL " & tabSpells.Fields("ReqLevel") & " " & GetMagery(tabSpells.Fields("Magery"), tabSpells.Fields("MageryLVL"))
                cmbAttackSpell(1).ItemData(cmbAttackSpell(1).NewIndex) = tabSpells.Fields("Number")
            End If
        End If
skip:
        tabSpells.MoveNext
    Loop
End If
cmbAttackSpell(0).AddItem "Select...", 0
cmbAttackSpell(1).AddItem "Select...", 0
cmbAttackSpell(0).ListIndex = 0
cmbAttackSpell(1).ListIndex = 0

'Call ExpandCombo(cmbAttackSpell(0), HeightOnly, NoExpand, Frame2.hwnd)
'Call ExpandCombo(cmbAttackSpell(1), HeightOnly, NoExpand, Frame2.hwnd)

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("RefreshAttackSpells")
Resume out:
End Sub
Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = vbMinimized Then Exit Sub

txtText.Width = Me.Width - 400
txtText.Height = Me.Height - TITLEBAR_OFFSET - 1000

End Sub

Private Sub optAttackType_Click(Index As Integer)
Dim x As Integer, nSelected As Integer

For x = 0 To 5
    If optAttackType(x).Value = True Then
        nSelected = x
    End If
Next x
If Not optAttackType(nSelected).Value = True Then optAttackType(nSelected).Value = True

Select Case nSelected
    Case 1: 'wep
        chkBashing.Enabled = True
        chkSmashing.Enabled = True
    Case 2: 'learned spell
        cmbAttackSpell(0).Enabled = True
    Case 3: 'any spell
        cmbAttackSpell(1).Enabled = True
        txtAttackSpellLevel.Enabled = True
        lblLabels(5).Enabled = True
    Case 4: 'ma
        cmbAttackMA.Enabled = True
    Case 5: 'manual
        txtAttackPhys.Enabled = True
        txtAttackMag.Enabled = True
        lblLabels(8).Enabled = True
        lblLabels(9).Enabled = True
End Select

If optAttackType(1).Value = False Then
    chkBashing.Enabled = False
    chkSmashing.Enabled = False
End If

If optAttackType(2).Value = False Then
    cmbAttackSpell(0).Enabled = False
End If

If optAttackType(3).Value = False Then
    cmbAttackSpell(1).Enabled = False
    txtAttackSpellLevel.Enabled = False
    lblLabels(5).Enabled = False
End If

If optAttackType(4).Value = False Then cmbAttackMA.Enabled = False

If optAttackType(5).Value = False Then
    txtAttackPhys.Enabled = False
    txtAttackMag.Enabled = False
    lblLabels(8).Enabled = False
    lblLabels(9).Enabled = False
End If

End Sub

Private Sub optRoomFindMatch_Click(Index As Integer)
optRoomFindMatch(Index).Value = True
optRoomFindMatch(Index).FontBold = True
If Index = 0 Then
    optRoomFindMatch(1).FontBold = False
ElseIf Index = 1 Then
    optRoomFindMatch(0).FontBold = False
End If
End Sub

Private Sub txtAttackMag_GotFocus()
Call SelectAll(txtAttackMag)
End Sub

Private Sub txtAttackMag_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtAttackPhys_GotFocus()
Call SelectAll(txtAttackPhys)
End Sub

Private Sub txtAttackPhys_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtAttackSpellLevel_GotFocus()
Call SelectAll(txtAttackSpellLevel)
End Sub

Private Sub txtAttackSpellLevel_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtRoomName_GotFocus()
Call SelectAll(txtRoomName)
End Sub
