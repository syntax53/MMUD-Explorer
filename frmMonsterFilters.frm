VERSION 5.00
Begin VB.Form frmMonsterFilters 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extra Monster Filters"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6345
   Icon            =   "frmMonsterFilters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQ 
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
      Height          =   375
      Left            =   3420
      TabIndex        =   51
      Top             =   5760
      Width           =   375
   End
   Begin VB.Timer timWindowMove 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5760
      Top             =   4920
   End
   Begin VB.Frame fraOther 
      Height          =   1455
      Left            =   2880
      TabIndex        =   50
      Top             =   2160
      Width           =   3375
      Begin VB.CheckBox chkIsUndead 
         Caption         =   "Is Undead"
         Height          =   255
         Left            =   1560
         TabIndex        =   21
         Top             =   240
         Width           =   1155
      End
      Begin VB.CheckBox chkNonHostile_vEvil 
         Caption         =   "Non-Hostile VS Evil"
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   600
         Width           =   1755
      End
      Begin VB.CheckBox chkIsNonHostile_vNG 
         Caption         =   "Non-Hostile VS Neutral/Good"
         Height          =   435
         Left            =   1560
         TabIndex        =   23
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkAtkNoPoison 
         Caption         =   "No Poison"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Only Undead"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkAtkNoConfusion 
         Caption         =   "No Confusion"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Only Undead"
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox chkAtkNoFear 
         Caption         =   "No Fear"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Only Undead"
         Top             =   960
         Width           =   1035
      End
   End
   Begin VB.Frame fraAbils 
      Height          =   1635
      Left            =   1140
      TabIndex        =   48
      Top             =   3720
      Width           =   3975
      Begin VB.CommandButton cmdAbilClear 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3300
         TabIndex        =   52
         Top             =   120
         Width           =   555
      End
      Begin VB.ComboBox cmbAbilities 
         Height          =   315
         Index           =   0
         ItemData        =   "frmMonsterFilters.frx":09AA
         Left            =   120
         List            =   "frmMonsterFilters.frx":09AC
         Sorted          =   -1  'True
         TabIndex        =   24
         Text            =   "cmbAbilities"
         Top             =   480
         Width           =   2475
      End
      Begin VB.ComboBox cmbAbilities 
         Height          =   315
         Index           =   1
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   27
         Text            =   "cmbAbilities"
         Top             =   840
         Width           =   2475
      End
      Begin VB.ComboBox cmbAbilities 
         Height          =   315
         Index           =   2
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   30
         Text            =   "cmbAbilities"
         Top             =   1200
         Width           =   2475
      End
      Begin VB.TextBox txtAbilityVal 
         Height          =   315
         Index           =   0
         Left            =   3300
         MaxLength       =   4
         TabIndex        =   26
         Text            =   "0"
         Top             =   480
         Width           =   555
      End
      Begin VB.ComboBox cmbAbilityOp 
         Height          =   315
         Index           =   0
         ItemData        =   "frmMonsterFilters.frx":09AE
         Left            =   2640
         List            =   "frmMonsterFilters.frx":09B8
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtAbilityVal 
         Height          =   315
         Index           =   1
         Left            =   3300
         MaxLength       =   4
         TabIndex        =   29
         Text            =   "0"
         Top             =   840
         Width           =   555
      End
      Begin VB.ComboBox cmbAbilityOp 
         Height          =   315
         Index           =   1
         ItemData        =   "frmMonsterFilters.frx":09C4
         Left            =   2640
         List            =   "frmMonsterFilters.frx":09CE
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtAbilityVal 
         Height          =   315
         Index           =   2
         Left            =   3300
         MaxLength       =   4
         TabIndex        =   32
         Text            =   "0"
         Top             =   1200
         Width           =   555
      End
      Begin VB.ComboBox cmbAbilityOp 
         Height          =   315
         Index           =   2
         ItemData        =   "frmMonsterFilters.frx":09DA
         Left            =   2640
         List            =   "frmMonsterFilters.frx":09E4
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblLabelArray 
         Alignment       =   2  'Center
         Caption         =   "Ability Filters (all AND'ed)"
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
         Left            =   120
         TabIndex        =   49
         Top             =   180
         Width           =   3675
      End
   End
   Begin VB.TextBox txtNumMobsLTE 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   5280
      MaxLength       =   4
      TabIndex        =   11
      Text            =   "9999"
      Top             =   1620
      Width           =   855
   End
   Begin VB.TextBox txtNumMobsGTE 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   4260
      MaxLength       =   4
      TabIndex        =   10
      Text            =   "0"
      Top             =   1620
      Width           =   855
   End
   Begin VB.TextBox txtNumLairs 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   5280
      MaxLength       =   4
      TabIndex        =   6
      Text            =   "0"
      Top             =   540
      Width           =   855
   End
   Begin VB.TextBox txtDodge 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2220
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "9999"
      Top             =   540
      Width           =   855
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "Reset"
      Height          =   615
      Index           =   1
      Left            =   3960
      TabIndex        =   35
      Top             =   5520
      Width           =   1035
   End
   Begin VB.CommandButton cmdExec 
      Cancel          =   -1  'True
      Caption         =   "Cancel +Close"
      Height          =   615
      Index           =   0
      Left            =   5160
      TabIndex        =   0
      Top             =   5520
      Width           =   1035
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "Save +Apply"
      Height          =   615
      Index           =   3
      Left            =   120
      TabIndex        =   33
      Top             =   5520
      Width           =   915
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "Save +Close"
      Height          =   615
      Index           =   2
      Left            =   1200
      TabIndex        =   34
      Top             =   5520
      Width           =   915
   End
   Begin VB.TextBox txtAtkAccuracyMax 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   8
      Text            =   "9999"
      Top             =   1620
      Width           =   855
   End
   Begin VB.TextBox txtAtkAccuracyMaj 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   180
      MaxLength       =   4
      TabIndex        =   7
      Text            =   "9999"
      Top             =   1620
      Width           =   855
   End
   Begin VB.TextBox txtLairEXP 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2640
      MaxLength       =   9
      TabIndex        =   9
      Text            =   "0"
      Top             =   1620
      Width           =   1035
   End
   Begin VB.TextBox txtGameLimit 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   4260
      MaxLength       =   4
      TabIndex        =   5
      Text            =   "9999"
      ToolTipText     =   "(Game limit is different than regen time)"
      Top             =   540
      Width           =   855
   End
   Begin VB.TextBox txtMR 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3240
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "9999"
      Top             =   540
      Width           =   855
   End
   Begin VB.TextBox txtDR 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   2
      Text            =   "9999"
      Top             =   540
      Width           =   855
   End
   Begin VB.TextBox txtAC 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   180
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "9999"
      Top             =   540
      Width           =   855
   End
   Begin VB.Frame fraCash 
      Caption         =   "Drops Coin"
      Height          =   1455
      Left            =   120
      TabIndex        =   36
      Top             =   2160
      Width           =   2655
      Begin VB.OptionButton optCash 
         Caption         =   "Runic"
         Height          =   315
         Index           =   5
         Left            =   1320
         TabIndex        =   17
         Top             =   1020
         Width           =   1155
      End
      Begin VB.OptionButton optCash 
         Caption         =   "Platinum+"
         Height          =   315
         Index           =   4
         Left            =   1320
         TabIndex        =   15
         Top             =   660
         Width           =   1155
      End
      Begin VB.OptionButton optCash 
         Caption         =   "Gold+"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   1020
         Width           =   975
      End
      Begin VB.OptionButton optCash 
         Caption         =   "Silver+"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   660
         Width           =   975
      End
      Begin VB.OptionButton optCash 
         Caption         =   "Copper+"
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   13
         Top             =   300
         Width           =   1155
      End
      Begin VB.OptionButton optCash 
         Caption         =   "No Filter"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Label lblLabelArray 
      Alignment       =   2  'Center
      Caption         =   "Mobs/ Lair <="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   11
      Left            =   5280
      TabIndex        =   47
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblLabelArray 
      Alignment       =   2  'Center
      Caption         =   "Mobs/ Lair >="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   10
      Left            =   4260
      TabIndex        =   46
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblLabelArray 
      Alignment       =   2  'Center
      Caption         =   "# Lairs >="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   5220
      TabIndex        =   45
      Top             =   300
      Width           =   975
   End
   Begin VB.Label lblLabelArray 
      Alignment       =   2  'Center
      Caption         =   "Dodge <="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   8
      Left            =   2280
      TabIndex        =   44
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblLabelArray 
      Alignment       =   2  'Center
      Caption         =   "MAX Attack ACCY <="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   1200
      TabIndex        =   43
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label lblLabelArray 
      Alignment       =   2  'Center
      Caption         =   "Majority Attack ACCY <="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   180
      TabIndex        =   42
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label lblLairEXP 
      Alignment       =   2  'Center
      Caption         =   "Lair EXP >="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2580
      TabIndex        =   41
      Top             =   1380
      Width           =   1155
   End
   Begin VB.Label lblLabelArray 
      Alignment       =   2  'Center
      Caption         =   "Game Limit <="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   4
      Left            =   4260
      TabIndex        =   40
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblLabelArray 
      Alignment       =   2  'Center
      Caption         =   "Magic Res <="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   3
      Left            =   3240
      TabIndex        =   39
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblLabelArray 
      Alignment       =   2  'Center
      Caption         =   "DR <="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   38
      Top             =   300
      Width           =   855
   End
   Begin VB.Label lblLabelArray 
      Alignment       =   2  'Center
      Caption         =   "AC <="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   37
      Top             =   300
      Width           =   855
   End
End
Attribute VB_Name = "frmMonsterFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Dim tWindowSize As WindowSizeProperties

Dim bMouseDown As Boolean
Dim ntimButtonPressCount As Long

Public nLastPosTop As Long
Public nLastPosLeft As Long
Public nLastPosMoved As Long
Public nLastPosMonitor As Long

Public nLastTimerTop As Long
Public nLastTimerLeft As Long

Private Sub cmbAbilities_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = AutoComplete(cmbAbilities(Index), KeyAscii, False)
End Sub

Private Sub cmdAbilClear_Click()
Dim y As Integer
For y = 0 To 2
    cmbAbilities(y).ListIndex = 0
    cmbAbilityOp(y).ListIndex = 0
    txtAbilityVal(y).Text = 0
Next y
End Sub

Private Sub cmdExec_Click(Index As Integer)
On Error GoTo error:
Dim x As Integer

Select Case Index
    Case 0: 'cancel+close
        Me.Tag = "0"
        GoTo unload_frm:
        
    Case 1: 'reset
        optCash(0).Value = True: Call optCash_Click(0)
        txtAC.Text = 9999
        txtDR.Text = 9999
        txtDodge.Text = 9999
        txtMR.Text = 9999
        txtGameLimit.Text = 9999
        txtLairEXP.Text = 0
        txtAtkAccuracyMaj.Text = 9999
        txtAtkAccuracyMax.Text = 9999
        txtNumLairs.Text = 0
        txtNumMobsLTE.Text = 9999
        txtNumMobsGTE.Text = 0
        
        chkIsUndead.Value = 0
        chkNonHostile_vEvil.Value = 0
        chkIsNonHostile_vNG.Value = 0
        chkAtkNoPoison.Value = 0
        chkAtkNoConfusion.Value = 0
        chkAtkNoFear.Value = 0
        
        For x = 0 To 2
            cmbAbilities(x).ListIndex = 0
            cmbAbilityOp(x).ListIndex = 0
            txtAbilityVal(x).Text = 0
        Next x
        
    Case 2: 'save+close
        Me.Tag = "1"
        GoTo save:
    Case 3: 'save+apply
        Me.Tag = "2"
        GoTo save:
End Select

out:
On Error Resume Next
Exit Sub

save:
If val(txtAC.Text) > 9999 Then txtAC.Text = 9999
If val(txtDR.Text) > 9999 Then txtDR.Text = 9999
If val(txtDodge.Text) > 9999 Then txtDodge.Text = 9999
If val(txtMR.Text) > 9999 Then txtMR.Text = 9999
If val(txtGameLimit.Text) > 9999 Then txtGameLimit.Text = 9999
If val(txtLairEXP.Text) > 999999999 Then txtLairEXP.Text = 999999999
If val(txtAtkAccuracyMaj.Text) > 9999 Then txtAtkAccuracyMaj.Text = 9999
If val(txtAtkAccuracyMax.Text) > 9999 Then txtAtkAccuracyMax.Text = 9999
If val(txtNumLairs.Text) > 9999 Then txtNumLairs.Text = 9999
If val(txtNumMobsLTE.Text) > 9999 Then txtNumMobsLTE.Text = 9999
If val(txtNumMobsGTE.Text) > 9999 Then txtNumMobsGTE.Text = 9999

If val(txtAC.Text) < 0 Then txtAC.Text = 9999
If val(txtDR.Text) < 0 Then txtDR.Text = 9999
If val(txtDodge.Text) < 0 Then txtDodge.Text = 9999
If val(txtMR.Text) < 0 Then txtMR.Text = 9999
If val(txtGameLimit.Text) < 0 Then txtGameLimit.Text = 9999
If val(txtLairEXP.Text) < 0 Then txtLairEXP.Text = 0
If val(txtAtkAccuracyMaj.Text) < 0 Then txtAtkAccuracyMaj.Text = 9999
If val(txtAtkAccuracyMax.Text) < 0 Then txtAtkAccuracyMax.Text = 9999
If val(txtNumLairs.Text) < 0 Then txtNumLairs.Text = 0
If val(txtNumMobsLTE.Text) < 0 Then txtNumMobsLTE.Text = 9999
If val(txtNumMobsGTE.Text) < 0 Then txtNumMobsGTE.Text = 0

filter_Monster_nArmourClass = val(txtAC.Text)
filter_Monster_nDamageResist = val(txtDR.Text)
filter_Monster_nMagicRes = val(txtMR.Text)
filter_Monster_nGameLimit = val(txtGameLimit.Text)
filter_Monster_nAvgLairExp = val(txtLairEXP.Text)
filter_Monster_nAtkAccuracyMaj = val(txtAtkAccuracyMaj.Text)
filter_Monster_nAtkAccuracyMax = val(txtAtkAccuracyMax.Text)
filter_Monster_nDodge = val(txtDodge.Text)
filter_Monster_nNumLairs = val(txtNumLairs.Text)
filter_Monster_nNumMobsLTE = val(txtNumMobsLTE.Text)
filter_Monster_nNumMobsGTE = val(txtNumMobsGTE.Text)

filter_Monster_bDropsCash = False
filter_Monster_bDropsR = False
filter_Monster_bDropsP = False
filter_Monster_bDropsG = False
filter_Monster_bDropsS = False
If optCash(1).Value = True Then
    filter_Monster_bDropsCash = True
ElseIf optCash(2).Value = True Then
    filter_Monster_bDropsS = True
ElseIf optCash(3).Value = True Then
    filter_Monster_bDropsG = True
ElseIf optCash(4).Value = True Then
    filter_Monster_bDropsP = True
ElseIf optCash(5).Value = True Then
    filter_Monster_bDropsR = True
End If

filter_Monster_bIsUndead = IIf(chkIsUndead.Value = 1, True, False)
filter_Monster_bIsNonHostile_vEvil = IIf(chkNonHostile_vEvil.Value = 1, True, False)
filter_Monster_bIsNonHostile_vNG = IIf(chkIsNonHostile_vNG.Value = 1, True, False)
filter_Monster_bAtkNoPoison = IIf(chkAtkNoPoison.Value = 1, True, False)
filter_Monster_bAtkNoConfusion = IIf(chkAtkNoConfusion.Value = 1, True, False)
filter_Monster_bAtkNoFear = IIf(chkAtkNoFear.Value = 1, True, False)

For x = 0 To 2
    If cmbAbilities(x).ListIndex > 0 Then
        filter_Monster_nAbilities(x, 0) = cmbAbilities(x).ItemData(cmbAbilities(x).ListIndex)
        filter_Monster_nAbilities(x, 1) = cmbAbilityOp(x).ListIndex
        
        If val(txtAbilityVal(x).Text) < 0 Then txtAbilityVal(x).Text = 0
        If val(txtAbilityVal(x).Text) > 9999 Then txtAbilityVal(x).Text = 9999
        filter_Monster_nAbilities(x, 2) = val(txtAbilityVal(x).Text)
    Else
        filter_Monster_nAbilities(x, 0) = 0
        filter_Monster_nAbilities(x, 1) = 0
        filter_Monster_nAbilities(x, 2) = 0
    End If
Next x

unload_frm:
On Error Resume Next
Me.Hide
Exit Sub

error:
Call HandleError("cmdExec_Click")
Resume out:
End Sub

Private Sub cmdQ_Click()
MsgBox "Click the headers to reset individual boxes.", vbInformation
End Sub

Private Sub Form_Load()
On Error GoTo error:
Dim y As Integer, x As Integer, sAbilityList() As String

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

Me.Tag = "0"

sAbilityList = GetAbilityList()

For y = 0 To 2
    cmbAbilities(y).clear
    For x = 1 To UBound(sAbilityList())
        If Len(sAbilityList(x)) > 0 Then
            cmbAbilities(y).AddItem sAbilityList(x)
            cmbAbilities(y).ItemData(cmbAbilities(y).NewIndex) = x
        End If
    Next x
    cmbAbilities(y).AddItem "Any", 0
    cmbAbilities(y).ItemData(cmbAbilities(y).NewIndex) = 0
    Call AutoSizeDropDownWidth(cmbAbilities(y))
    Call ExpandCombo(cmbAbilities(y), HeightOnly, DoubleWidth, fraAbils.hWnd)
    cmbAbilities(y).ListIndex = 0
    cmbAbilityOp(y).ListIndex = 0
Next y

cmdExec(0).Caption = "Cancel" & vbCrLf & "+ Close"
cmdExec(1).Caption = "Reset"
cmdExec(2).Caption = "Save" & vbCrLf & "+ Close"
cmdExec(3).Caption = "Save" & vbCrLf & "+ Apply"

txtAC.Text = filter_Monster_nArmourClass
txtDR.Text = filter_Monster_nDamageResist
txtMR.Text = filter_Monster_nMagicRes
txtGameLimit.Text = filter_Monster_nGameLimit
txtLairEXP.Text = filter_Monster_nAvgLairExp
txtAtkAccuracyMaj.Text = filter_Monster_nAtkAccuracyMaj
txtAtkAccuracyMax.Text = filter_Monster_nAtkAccuracyMax
txtDodge.Text = filter_Monster_nDodge
txtNumLairs.Text = filter_Monster_nNumLairs
txtNumMobsLTE.Text = filter_Monster_nNumMobsLTE
txtNumMobsGTE.Text = filter_Monster_nNumMobsGTE

If filter_Monster_bDropsCash Then
    optCash(1).Value = True
ElseIf filter_Monster_bDropsS Then
    optCash(2).Value = True
ElseIf filter_Monster_bDropsG Then
    optCash(3).Value = True
ElseIf filter_Monster_bDropsP Then
    optCash(4).Value = True
ElseIf filter_Monster_bDropsR Then
    optCash(5).Value = True
Else
    optCash(0).Value = True
End If
Call optCash_Click(0)

chkIsUndead.Value = IIf(filter_Monster_bIsUndead, 1, 0)
chkNonHostile_vEvil.Value = IIf(filter_Monster_bIsNonHostile_vEvil, 1, 0)
chkIsNonHostile_vNG.Value = IIf(filter_Monster_bIsNonHostile_vNG, 1, 0)
chkAtkNoPoison.Value = IIf(filter_Monster_bAtkNoPoison, 1, 0)
chkAtkNoConfusion.Value = IIf(filter_Monster_bAtkNoConfusion, 1, 0)
chkAtkNoFear.Value = IIf(filter_Monster_bAtkNoFear, 1, 0)

For x = 0 To 2
    If filter_Monster_nAbilities(x, 0) > 0 Then
        For y = 0 To cmbAbilities(x).ListCount - 1
            If cmbAbilities(x).ItemData(y) = filter_Monster_nAbilities(x, 0) Then
                cmbAbilities(x).ListIndex = y
                GoTo abil_found:
            End If
        Next y
        cmbAbilities(x).ListIndex = 0
        cmbAbilityOp(x).ListIndex = 0
        txtAbilityVal(x).Text = 0
        GoTo abil_notfound:
abil_found:
        If filter_Monster_nAbilities(x, 1) = 1 Then
            cmbAbilityOp(x).ListIndex = 1
        Else
            cmbAbilityOp(x).ListIndex = 0
        End If
        txtAbilityVal(x).Text = filter_Monster_nAbilities(x, 2)
abil_notfound:
    Else
        cmbAbilities(x).ListIndex = 0
        cmbAbilityOp(x).ListIndex = 0
        txtAbilityVal(x).Text = 0
    End If
Next x

If nNMRVer >= 1.83 Then
    lblLairEXP.Enabled = True
    txtLairEXP.Enabled = True
    txtLairEXP.Locked = False
Else
    lblLairEXP.Enabled = False
    txtLairEXP.Enabled = False
    txtLairEXP.Locked = True
    txtLairEXP.Text = 0
    filter_Monster_nAvgLairExp = 0
End If

timWindowMove.Enabled = True

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("Form_Load")
Resume out:
End Sub


Private Sub lblLabelArray_Click(Index As Integer)
Select Case Index
    Case 1: txtAC.Text = 9999
    Case 2: txtDR.Text = 9999
    Case 8: txtDodge.Text = 9999
    Case 3: txtMR.Text = 9999
    Case 4: txtGameLimit.Text = 9999
    Case 9: txtNumLairs.Text = 0
    Case 6: txtAtkAccuracyMaj.Text = 9999
    Case 7: txtAtkAccuracyMax.Text = 9999
    Case 10: txtNumMobsGTE.Text = 0
    Case 11: txtNumMobsLTE.Text = 9999
End Select
End Sub

Private Sub lblLairEXP_Click()
txtLairEXP.Text = 0
End Sub

Private Sub optCash_Click(Index As Integer)
If optCash(0).Value = True Then
    fraCash.FontBold = False
Else
    fraCash.FontBold = True
End If
End Sub

Private Sub timWindowMove_Timer()
Call MonitorFormTimer(Me)
End Sub

Private Sub txtAbilityVal_GotFocus(Index As Integer)
Call SelectAll(txtAbilityVal(Index))
End Sub

Private Sub txtAbilityVal_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtAC_GotFocus()
Call SelectAll(txtAC)
End Sub

Private Sub txtAC_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtAtkAccuracyMaj_GotFocus()
Call SelectAll(txtAtkAccuracyMaj)
End Sub

Private Sub txtAtkAccuracyMax_GotFocus()
Call SelectAll(txtAtkAccuracyMax)
End Sub

Private Sub txtDodge_GotFocus()
Call SelectAll(txtDodge)
End Sub

Private Sub txtDodge_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtDR_GotFocus()
Call SelectAll(txtDR)
End Sub

Private Sub txtDR_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtGameLimit_GotFocus()
Call SelectAll(txtGameLimit)
End Sub

Private Sub txtGameLimit_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtLairEXP_GotFocus()
Call SelectAll(txtLairEXP)
End Sub

Private Sub txtLairEXP_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtMR_GotFocus()
Call SelectAll(txtMR)
End Sub

Private Sub txtMR_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtNumLairs_GotFocus()
Call SelectAll(txtNumLairs)
End Sub

Private Sub txtNumLairs_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtNumMobsGTE_GotFocus()
Call SelectAll(txtNumMobsGTE)
End Sub

Private Sub txtNumMobsGTE_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtNumMobsLTE_GotFocus()
Call SelectAll(txtNumMobsLTE)
End Sub

Private Sub txtNumMobsLTE_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub
