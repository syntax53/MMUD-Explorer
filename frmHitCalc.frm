VERSION 5.00
Begin VB.Form frmHitCalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hit Calculator"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6240
   Icon            =   "frmHitCalc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timButtonPress 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5640
      Top             =   120
   End
   Begin VB.CommandButton cmdCharHitCalc 
      Caption         =   "-"
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
      Index           =   7
      Left            =   180
      TabIndex        =   34
      Top             =   4980
      Width           =   315
   End
   Begin VB.CommandButton cmdCharHitCalc 
      Caption         =   "+"
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
      Index           =   6
      Left            =   1680
      TabIndex        =   33
      Top             =   4980
      Width           =   315
   End
   Begin VB.ComboBox cmbMonsterList 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1260
      Sorted          =   -1  'True
      TabIndex        =   31
      Top             =   3000
      Width           =   4815
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   435
      Left            =   180
      TabIndex        =   26
      Top             =   1350
      Width           =   5895
      Begin VB.OptionButton optDefender 
         Caption         =   "Manual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   4140
         TabIndex        =   29
         Top             =   60
         Width           =   1230
      End
      Begin VB.OptionButton optDefender 
         Caption         =   "Monster"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   2520
         TabIndex        =   28
         Top             =   60
         Width           =   1290
      End
      Begin VB.OptionButton optDefender 
         Caption         =   "Current Character"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   60
         TabIndex        =   27
         Top             =   60
         Value           =   -1  'True
         Width           =   2250
      End
   End
   Begin VB.TextBox txtHitCalc 
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
      Index           =   1
      Left            =   2700
      MaxLength       =   4
      TabIndex        =   21
      Text            =   "0"
      Top             =   4200
      Width           =   795
   End
   Begin VB.TextBox txtHitCalc 
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
      Index           =   0
      Left            =   540
      MaxLength       =   4
      TabIndex        =   20
      Text            =   "9999"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox txtHitCalc 
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
      Index           =   2
      Left            =   4560
      MaxLength       =   5
      TabIndex        =   19
      Text            =   "0"
      ToolTipText     =   "Dodge value, not dodge %"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox txtHitCalc 
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
      Index           =   3
      Left            =   540
      MaxLength       =   5
      TabIndex        =   18
      Text            =   "0"
      Top             =   4980
      Width           =   1095
   End
   Begin VB.CommandButton cmdCharButtons 
      Caption         =   "Refresh Stats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   5
      Left            =   540
      TabIndex        =   17
      ToolTipText     =   "Reset to current char's stats"
      Top             =   5580
      Width           =   1095
   End
   Begin VB.CommandButton cmdCharButtons 
      Caption         =   "?"
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
      Index           =   6
      Left            =   60
      TabIndex        =   16
      ToolTipText     =   "Reset to current char's stats"
      Top             =   5940
      Width           =   315
   End
   Begin VB.CommandButton cmdCharHitCalc 
      Caption         =   "+"
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
      Index           =   1
      Left            =   1680
      TabIndex        =   15
      Top             =   4200
      Width           =   315
   End
   Begin VB.CommandButton cmdCharHitCalc 
      Caption         =   "-"
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
      Index           =   0
      Left            =   180
      TabIndex        =   14
      Top             =   4200
      Width           =   315
   End
   Begin VB.CommandButton cmdCharHitCalc 
      Caption         =   "+"
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
      Index           =   5
      Left            =   5700
      TabIndex        =   13
      Top             =   4200
      Width           =   315
   End
   Begin VB.CommandButton cmdCharHitCalc 
      Caption         =   "-"
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
      Index           =   4
      Left            =   4200
      TabIndex        =   12
      Top             =   4200
      Width           =   315
   End
   Begin VB.CommandButton cmdCharHitCalc 
      Caption         =   "+"
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
      Index           =   3
      Left            =   3600
      TabIndex        =   11
      Top             =   4200
      Width           =   315
   End
   Begin VB.CommandButton cmdCharHitCalc 
      Caption         =   "-"
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
      Index           =   2
      Left            =   2340
      TabIndex        =   10
      Top             =   4200
      Width           =   315
   End
   Begin VB.Timer timWindowMove 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   435
      Left            =   180
      TabIndex        =   4
      Top             =   2160
      Width           =   5895
      Begin VB.OptionButton optDefender 
         Caption         =   "vs Current Char."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   60
         Value           =   -1  'True
         Width           =   2130
      End
      Begin VB.OptionButton optDefender 
         Caption         =   "vs Mob"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   2460
         TabIndex        =   6
         Top             =   60
         Width           =   1290
      End
      Begin VB.OptionButton optDefender 
         Caption         =   "vs Player"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3900
         TabIndex        =   7
         Top             =   60
         Width           =   1650
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5955
      Begin VB.OptionButton optHitCalcType 
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   2520
         TabIndex        =   1
         Top             =   30
         Width           =   1350
      End
      Begin VB.OptionButton optHitCalcType 
         Caption         =   "Backstab"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   4200
         TabIndex        =   2
         Top             =   30
         Width           =   1530
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Calculate For:"
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
         Left            =   360
         TabIndex        =   9
         Top             =   60
         Width           =   1455
      End
   End
   Begin VB.Label lblLabelArray 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2700
      TabIndex        =   36
      Top             =   6000
      Width           =   2835
   End
   Begin VB.Label lblLabelArray 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Index           =   35
      Left            =   2700
      TabIndex        =   35
      Top             =   4800
      Width           =   2835
   End
   Begin VB.Label lblLabelArray 
      AutoSize        =   -1  'True
      Caption         =   "Monster:"
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
      Left            =   180
      TabIndex        =   32
      Top             =   3060
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   4
      X1              =   120
      X2              =   6120
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   120
      X2              =   2160
      Y1              =   1290
      Y2              =   1290
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   3960
      X2              =   6060
      Y1              =   1290
      Y2              =   1290
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Attacker"
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
      Left            =   2460
      TabIndex        =   30
      Top             =   1140
      Width           =   1230
   End
   Begin VB.Label lblLabelArray 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "vs AC:"
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
      Index           =   32
      Left            =   2790
      TabIndex        =   25
      Top             =   3900
      Width           =   690
   End
   Begin VB.Label lblLabelArray 
      Alignment       =   2  'Center
      Caption         =   "Attacker Accuracy:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   34
      Left            =   570
      TabIndex        =   24
      Top             =   3660
      Width           =   1050
   End
   Begin VB.Label lblLabelArray 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "vs Dodge:"
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
      Index           =   62
      Left            =   4575
      TabIndex        =   23
      Top             =   3900
      Width           =   1080
   End
   Begin VB.Label lblLabelArray 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "+2nd D:"
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
      Index           =   63
      Left            =   720
      TabIndex        =   22
      Top             =   4680
      Width           =   810
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Hit Calculator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5970
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Defender"
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
      Left            =   2460
      TabIndex        =   3
      Top             =   1950
      Width           =   1230
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   3960
      X2              =   6060
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   2160
      Y1              =   2100
      Y2              =   2100
   End
End
Attribute VB_Name = "frmHitCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Public nLastPosTop As Long
Public nLastPosLeft As Long
Public nLastPosMoved As Long
Public nLastPosMonitor As Long

Public nLastTimerTop As Long
Public nLastTimerLeft As Long

Dim bMouseDown As Boolean
Dim tWindowSize As WindowSizeProperties


Private Sub cmbMonsterList_KeyPress(KeyAscii As Integer)
KeyAscii = AutoComplete(cmbMonsterList, KeyAscii, False)
End Sub

Private Sub cmdCharHitCalc_Click(Index As Integer)
If Not bMouseDown Then Call ModifyHitCalc(Index)
timButtonPress.Enabled = False
bMouseDown = False
End Sub

Private Sub cmdCharHitCalc_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

bMouseDown = True

Do While bMouseDown
    timButtonPress.Enabled = True
    Call ModifyHitCalc(Index)
    Do While timButtonPress.Enabled
        DoEvents
    Loop
Loop

End Sub

Private Sub ModifyHitCalc(ByVal Index As Integer)
On Error GoTo error:


out:
On Error Resume Next
Exit Sub
error:
Call HandleError("ModifyHitCalc")
Resume out:
End Sub

Private Sub cmdCharHitCalc_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
bMouseDown = False
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

timWindowMove.Enabled = True

Call LoadMonsters

Exit Sub
error:
Call HandleError("frmHitCalc_Load")
Resume Next
End Sub

Private Sub LoadMonsters()
Dim sName As String
On Error GoTo error:

cmbMonsterList.clear

tabMonsters.MoveFirst
Do While Not tabMonsters.EOF
    If bOnlyInGame And tabMonsters.Fields("In Game") = 0 Then GoTo MoveNext:
    
    sName = tabMonsters.Fields("Name")
    If sName = "" Or Left(sName, 3) = "sdf" Then GoTo MoveNext:
    sName = sName & " (" & tabMonsters.Fields("Number") & ")"
    
    cmbMonsterList.AddItem sName
    cmbMonsterList.ItemData(cmbMonsterList.NewIndex) = tabMonsters.Fields("Number")
    
MoveNext:
    tabMonsters.MoveNext
Loop
tabMonsters.MoveFirst

If cmbMonsterList.ListCount = 0 Then Exit Sub

cmbMonsterList.ListIndex = 0
Call AutoSizeDropDownWidth(cmbMonsterList)
Call ExpandCombo(cmbMonsterList, HeightOnly, DoubleWidth, Me.hWnd)
cmbMonsterList.SelLength = 0

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("LoadMonsters")
Resume out:
End Sub

Private Sub timWindowMove_Timer()
Call MonitorFormTimer(Me)
End Sub


Private Sub timButtonPress_Timer()
If bAppReallyTerminating Or bAppTerminating Then Exit Sub
timButtonPress.Enabled = False
End Sub
