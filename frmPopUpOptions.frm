VERSION 5.00
Object = "{20D5284F-7B23-4F0A-B8B1-6C9D18B64F1C}#1.0#0"; "exlimiter.ocx"
Begin VB.Form frmPopUpOptions 
   Caption         =   "MMUD Explorer"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7020
   ControlBox      =   0   'False
   Icon            =   "frmPopUpOptions.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   7020
   Begin VB.Frame fraRoomFind 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3855
      Left            =   60
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   6915
      Begin VB.Frame fraMISC 
         BorderStyle     =   0  'None
         Caption         =   "With the updates I have coming out in the  "
         Height          =   3495
         Left            =   240
         TabIndex        =   5
         Top             =   180
         Width           =   6435
         Begin VB.OptionButton optRoomFindMatch 
            Caption         =   "Exact Match"
            Height          =   240
            Index           =   1
            Left            =   4200
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   12
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
            TabIndex        =   13
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
            TabIndex        =   14
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
            TabIndex        =   15
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
            TabIndex        =   16
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
            TabIndex        =   17
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
            TabIndex        =   18
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
            TabIndex        =   19
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
            TabIndex        =   20
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
            TabIndex        =   21
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
            TabIndex        =   7
            Top             =   300
            Width           =   3015
         End
         Begin VB.CommandButton cmdPasteQ 
            BackColor       =   &H00FFC0FF&
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Do not include invisible, hidden, or activated exits"
            Height          =   495
            Index           =   3
            Left            =   240
            TabIndex        =   11
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
            TabIndex        =   10
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "3 or more letters"
            Height          =   195
            Index           =   1
            Left            =   540
            TabIndex        =   9
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
            TabIndex        =   8
            Top             =   300
            Width           =   1575
         End
      End
   End
   Begin exlimiter.EL EL1 
      Left            =   4560
      Top             =   3180
      _ExtentX        =   1270
      _ExtentY        =   1270
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
      TabIndex        =   2
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
      TabIndex        =   3
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
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox txtText 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   60
      MaxLength       =   10000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
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

Private Sub cmdPasteQ_Click(Index As Integer)
cmdContinue.SetFocus
Select Case Index
    Case 0:
        
End Select
End Sub

Public Sub ResetRoomFind()
On Error GoTo error:
Dim x As Integer

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

With EL1
    .CenterOnLoad = True
    .FormInQuestion = Me
    .MinWidth = 500
    .MinHeight = 340
    .EnableLimiter = True
End With

If frmMain.WindowState = vbMinimized Then
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
Else
    Me.Left = frmMain.Left + ((frmMain.Width - Me.Width) / 2)
    Me.Top = frmMain.Top + ((frmMain.Height - Me.Height) / 2)
End If
    
out:
On Error Resume Next
Exit Sub
error:
Call HandleError("Form_Load")
Resume out:
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = vbMinimized Then Exit Sub

txtText.Width = Me.Width - 400
txtText.Height = Me.Height - TITLEBAR_OFFSET - 1000

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

Private Sub txtRoomName_GotFocus()
Call SelectAll(txtRoomName)
End Sub
