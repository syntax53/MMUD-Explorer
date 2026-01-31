VERSION 5.00
Begin VB.Form frmMonsterFilters 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extra Monster Filters"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6345
   Icon            =   "frmMonsterFilters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBSDef 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2220
      MaxLength       =   4
      TabIndex        =   11
      Text            =   "9999"
      Top             =   2160
      Width           =   855
   End
   Begin VB.OptionButton optEnabled 
      Caption         =   "Enabled"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   3780
      TabIndex        =   2
      Top             =   120
      Width           =   1395
   End
   Begin VB.OptionButton optEnabled 
      Caption         =   "Disabled"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   420
      Index           =   0
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CheckBox chkShowAll 
      Caption         =   "Show All Monsters, even if they don't match filter (will be greyed out)"
      Height          =   315
      Left            =   480
      TabIndex        =   36
      Top             =   6060
      Width           =   5355
   End
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
      TabIndex        =   55
      Top             =   6780
      Width           =   375
   End
   Begin VB.Timer timWindowMove 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5760
      Top             =   5460
   End
   Begin VB.Frame fraOther 
      Height          =   1455
      Left            =   2880
      TabIndex        =   54
      Top             =   2700
      Width           =   3375
      Begin VB.CheckBox chkIsUndead 
         Caption         =   "Is Undead"
         Height          =   255
         Left            =   1560
         TabIndex        =   24
         Top             =   240
         Width           =   1155
      End
      Begin VB.CheckBox chkNonHostile_vEvil 
         Caption         =   "Non-Hostile VS Evil"
         Height          =   255
         Left            =   1560
         TabIndex        =   25
         Top             =   600
         Width           =   1755
      End
      Begin VB.CheckBox chkIsNonHostile_vNG 
         Caption         =   "Non-Hostile VS Neutral/Good"
         Height          =   435
         Left            =   1560
         TabIndex        =   26
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkAtkNoPoison 
         Caption         =   "No Poison"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Only Undead"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkAtkNoConfusion 
         Caption         =   "No Confusion"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Only Undead"
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox chkAtkNoFear 
         Caption         =   "No Fear"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   "Only Undead"
         Top             =   960
         Width           =   1035
      End
   End
   Begin VB.Frame fraAbils 
      Height          =   1635
      Left            =   1140
      TabIndex        =   52
      Top             =   4320
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
         TabIndex        =   56
         Top             =   120
         Width           =   555
      End
      Begin VB.ComboBox cmbAbilities 
         Height          =   315
         Index           =   0
         ItemData        =   "frmMonsterFilters.frx":0CCA
         Left            =   120
         List            =   "frmMonsterFilters.frx":0CCC
         Sorted          =   -1  'True
         TabIndex        =   27
         Text            =   "cmbAbilities"
         Top             =   480
         Width           =   2475
      End
      Begin VB.ComboBox cmbAbilities 
         Height          =   315
         Index           =   1
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   30
         Text            =   "cmbAbilities"
         Top             =   840
         Width           =   2475
      End
      Begin VB.ComboBox cmbAbilities 
         Height          =   315
         Index           =   2
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   33
         Text            =   "cmbAbilities"
         Top             =   1200
         Width           =   2475
      End
      Begin VB.TextBox txtAbilityVal 
         Height          =   315
         Index           =   0
         Left            =   3300
         MaxLength       =   4
         TabIndex        =   29
         Text            =   "0"
         Top             =   480
         Width           =   555
      End
      Begin VB.ComboBox cmbAbilityOp 
         Height          =   315
         Index           =   0
         ItemData        =   "frmMonsterFilters.frx":0CCE
         Left            =   2640
         List            =   "frmMonsterFilters.frx":0CD8
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtAbilityVal 
         Height          =   315
         Index           =   1
         Left            =   3300
         MaxLength       =   4
         TabIndex        =   32
         Text            =   "0"
         Top             =   840
         Width           =   555
      End
      Begin VB.ComboBox cmbAbilityOp 
         Height          =   315
         Index           =   1
         ItemData        =   "frmMonsterFilters.frx":0CE4
         Left            =   2640
         List            =   "frmMonsterFilters.frx":0CEE
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtAbilityVal 
         Height          =   315
         Index           =   2
         Left            =   3300
         MaxLength       =   4
         TabIndex        =   35
         Text            =   "0"
         Top             =   1200
         Width           =   555
      End
      Begin VB.ComboBox cmbAbilityOp 
         Height          =   315
         Index           =   2
         ItemData        =   "frmMonsterFilters.frx":0CFA
         Left            =   2640
         List            =   "frmMonsterFilters.frx":0D04
         Style           =   2  'Dropdown List
         TabIndex        =   34
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
         TabIndex        =   53
         Top             =   180
         Width           =   3675
      End
   End
   Begin VB.TextBox txtNumMobsLTE 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   5280
      MaxLength       =   4
      TabIndex        =   14
      Text            =   "9999"
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtNumMobsGTE 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   4260
      MaxLength       =   4
      TabIndex        =   13
      Text            =   "0"
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtNumLairs 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   5280
      MaxLength       =   4
      TabIndex        =   8
      Text            =   "0"
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtDodge 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2220
      MaxLength       =   4
      TabIndex        =   5
      Text            =   "9999"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "Reset"
      Height          =   615
      Index           =   1
      Left            =   3960
      TabIndex        =   39
      Top             =   6540
      Width           =   1035
   End
   Begin VB.CommandButton cmdExec 
      Cancel          =   -1  'True
      Caption         =   "Cancel +Close"
      Height          =   615
      Index           =   0
      Left            =   5160
      TabIndex        =   0
      Top             =   6540
      Width           =   1035
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "Save +Apply"
      Height          =   615
      Index           =   3
      Left            =   120
      TabIndex        =   37
      Top             =   6540
      Width           =   915
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "Save +Close"
      Height          =   615
      Index           =   2
      Left            =   1200
      TabIndex        =   38
      Top             =   6540
      Width           =   915
   End
   Begin VB.TextBox txtAtkAccuracyMax 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   10
      Text            =   "9999"
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtAtkAccuracyMaj 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   180
      MaxLength       =   4
      TabIndex        =   9
      Text            =   "9999"
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtLairEXP 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3240
      MaxLength       =   9
      TabIndex        =   12
      Text            =   "0"
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtGameLimit 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   4260
      MaxLength       =   4
      TabIndex        =   7
      Text            =   "9999"
      ToolTipText     =   "(Game limit is different than regen time)"
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtMR 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3240
      MaxLength       =   4
      TabIndex        =   6
      Text            =   "9999"
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtDR 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "9999"
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtAC 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   180
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "9999"
      Top             =   1080
      Width           =   855
   End
   Begin VB.Frame fraCash 
      Caption         =   "Drops Coin"
      Height          =   1455
      Left            =   120
      TabIndex        =   40
      Top             =   2700
      Width           =   2655
      Begin VB.OptionButton optCash 
         Caption         =   "Runic"
         Height          =   315
         Index           =   5
         Left            =   1320
         TabIndex        =   20
         Top             =   1020
         Width           =   1155
      End
      Begin VB.OptionButton optCash 
         Caption         =   "Platinum+"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   1020
         Width           =   1095
      End
      Begin VB.OptionButton optCash 
         Caption         =   "Gold+"
         Height          =   315
         Index           =   3
         Left            =   1320
         TabIndex        =   18
         Top             =   660
         Width           =   1155
      End
      Begin VB.OptionButton optCash 
         Caption         =   "Silver+"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   660
         Width           =   975
      End
      Begin VB.OptionButton optCash 
         Caption         =   "Copper+"
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   16
         Top             =   300
         Width           =   1155
      End
      Begin VB.OptionButton optCash 
         Caption         =   "No Filter"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Label lblLabelArray 
      Alignment       =   2  'Center
      Caption         =   "BS Def <="
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
      Left            =   2280
      TabIndex        =   58
      Top             =   1740
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Toggle Filter:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   555
      TabIndex        =   57
      Top             =   160
      Width           =   1515
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
      Index           =   5
      Left            =   5280
      TabIndex        =   51
      Top             =   1740
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
      TabIndex        =   50
      Top             =   1740
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
      TabIndex        =   49
      Top             =   840
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
      TabIndex        =   48
      Top             =   660
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
      TabIndex        =   47
      Top             =   1560
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
      TabIndex        =   46
      Top             =   1560
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
      Height          =   435
      Left            =   3300
      TabIndex        =   45
      Top             =   1740
      Width           =   795
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
      TabIndex        =   44
      Top             =   660
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
      TabIndex        =   43
      Top             =   660
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
      TabIndex        =   42
      Top             =   840
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
      TabIndex        =   41
      Top             =   840
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
Dim bSkipCheckForceEnable As Boolean

Public nLastPosTop As Long
Public nLastPosLeft As Long
Public nLastPosMoved As Long
Public nLastPosMonitor As Long

Public nLastTimerTop As Long
Public nLastTimerLeft As Long

Private Sub chkAtkNoConfusion_Click()
Call CheckForceEnable
End Sub

Private Sub chkAtkNoFear_Click()
Call CheckForceEnable
End Sub

Private Sub chkAtkNoPoison_Click()
Call CheckForceEnable
End Sub

Private Sub chkIsNonHostile_vNG_Click()
Call CheckForceEnable
End Sub

Private Sub chkIsUndead_Click()
Call CheckForceEnable
End Sub

Private Sub chkNonHostile_vEvil_Click()
Call CheckForceEnable
End Sub

Private Sub cmbAbilities_Click(Index As Integer)
Call CheckForceEnable
End Sub

Private Sub cmbAbilities_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = AutoComplete(cmbAbilities(Index), KeyAscii, False)
End Sub

Private Sub cmdAbilClear_Click()
On Error GoTo error:
Dim y As Integer
bSkipCheckForceEnable = True

For y = 0 To 2
    cmbAbilities(y).ListIndex = 0
    cmbAbilityOp(y).ListIndex = 0
    txtAbilityVal(y).Text = 0
Next y

out:
On Error Resume Next
bSkipCheckForceEnable = False
Call CheckForceEnable

Exit Sub
error:
Call HandleError("cmdAbilClear_Click")
Resume out:
End Sub

Private Sub cmdExec_Click(Index As Integer)
On Error GoTo error:
Dim x As Integer

Select Case Index
    Case 0: 'cancel+close
        Me.Tag = "0"
        GoTo hide_frm:
        
    Case 1: 'reset
        bSkipCheckForceEnable = True
        optCash(0).Value = True: Call optCash_Click(0)
        optEnabled(0).Value = True: Call optEnabled_Click(0)
        txtAC.Text = 9999
        txtDR.Text = 9999
        txtDodge.Text = 9999
        txtMR.Text = 9999
        txtBSDef.Text = 9999
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
        chkShowAll.Value = 0
        
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
bSkipCheckForceEnable = False
Exit Sub

save:
If val(txtAC.Text) > 9999 Then txtAC.Text = 9999
If val(txtDR.Text) > 9999 Then txtDR.Text = 9999
If val(txtDodge.Text) > 9999 Then txtDodge.Text = 9999
If val(txtMR.Text) > 9999 Then txtMR.Text = 9999
If val(txtBSDef.Text) > 9999 Then txtBSDef.Text = 9999
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
If val(txtBSDef.Text) < 0 Then txtBSDef.Text = 9999
If val(txtGameLimit.Text) < 0 Then txtGameLimit.Text = 9999
If val(txtLairEXP.Text) < 0 Then txtLairEXP.Text = 0
If val(txtAtkAccuracyMaj.Text) < 0 Then txtAtkAccuracyMaj.Text = 9999
If val(txtAtkAccuracyMax.Text) < 0 Then txtAtkAccuracyMax.Text = 9999
If val(txtNumLairs.Text) < 0 Then txtNumLairs.Text = 0
If val(txtNumMobsLTE.Text) < 0 Then txtNumMobsLTE.Text = 9999
If val(txtNumMobsGTE.Text) < 0 Then txtNumMobsGTE.Text = 0

If optEnabled(1).Value = True Then
    filter_Monster_bExtrasEnabled = True
Else
    filter_Monster_bExtrasEnabled = False
End If

filter_Monster_nArmourClass = val(txtAC.Text)
filter_Monster_nDamageResist = val(txtDR.Text)
filter_Monster_nMagicRes = val(txtMR.Text)
filter_Monster_nBSDef = val(txtBSDef.Text)
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
filter_Monster_bShowAll = IIf(chkShowAll.Value = 1, True, False)

For x = 0 To 2
    If cmbAbilities(x).ListIndex > 0 Then
        filter_Monster_nAbilities(x, 0) = cmbAbilities(x).ItemData(cmbAbilities(x).ListIndex)
        filter_Monster_nAbilities(x, 1) = cmbAbilityOp(x).ListIndex
        
        If val(txtAbilityVal(x).Text) < -9999 Then txtAbilityVal(x).Text = 0
        If val(txtAbilityVal(x).Text) > 9999 Then txtAbilityVal(x).Text = 9999
        filter_Monster_nAbilities(x, 2) = val(txtAbilityVal(x).Text)
    Else
        filter_Monster_nAbilities(x, 0) = 0
        filter_Monster_nAbilities(x, 1) = 0
        filter_Monster_nAbilities(x, 2) = 0
    End If
Next x

If Me.Tag = "2" Then GoTo no_hide: 'save+apply

hide_frm:
Me.Hide
no_hide:
On Error Resume Next
bSkipCheckForceEnable = False
Call frmMain.MonsterFilterFormAction
Exit Sub

error:
Call HandleError("cmdExec_Click")
Resume out:
End Sub

Private Sub cmdQ_Click()
MsgBox "Filters on this pop-up are not active until saved.  Click the headers to reset individual boxes.  Everything but the ability filters are saved to the character file.", vbInformation
End Sub

Private Sub Form_Load()
On Error GoTo error:
Dim y As Integer, x As Integer, sAbilityList() As String

Call SetWindowLong(Me.hWnd, GWL_HWNDPARENT, 0)

'stop windows from resizing fixed-size windows when changing dpi
If bDPIAwareMode Then Call SubclassFormMinMaxSize(Me, tWindowSize, True)

bSkipCheckForceEnable = True

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

If filter_Monster_bExtrasEnabled Then
    optEnabled(1).Value = True
Else
    optEnabled(0).Value = True
End If
Call optEnabled_Click(0)

txtAC.Text = filter_Monster_nArmourClass
txtDR.Text = filter_Monster_nDamageResist
txtMR.Text = filter_Monster_nMagicRes
txtBSDef.Text = filter_Monster_nBSDef
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
chkShowAll.Value = IIf(filter_Monster_bShowAll, 1, 0)

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
bSkipCheckForceEnable = False
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
    Case 5: txtNumMobsLTE.Text = 9999
    Case 11: txtBSDef.Text = 9999
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
Call CheckForceEnable
End Sub

Private Sub optEnabled_Click(Index As Integer)
If optEnabled(0).Value = True Then
    optEnabled(0).ForeColor = &HC0&       'red
    optEnabled(0).FontBold = True
    optEnabled(1).ForeColor = &H80000012  'black
    optEnabled(1).FontBold = False
Else
    optEnabled(0).ForeColor = &H80000012  'black
    optEnabled(0).FontBold = False
    optEnabled(1).ForeColor = &H8000&     'green
    optEnabled(1).FontBold = True
End If
On Error Resume Next
'cmdExec(3).SetFocus
End Sub

Private Sub timWindowMove_Timer()
Call MonitorFormTimer(Me)
End Sub

Private Sub txtAbilityVal_Change(Index As Integer)
Call CheckForceEnable
End Sub

Private Sub txtAbilityVal_GotFocus(Index As Integer)
Call SelectAll(txtAbilityVal(Index))
End Sub

Private Sub txtAbilityVal_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub CheckForceEnable()
On Error GoTo error:
Dim x As Integer

If bSkipCheckForceEnable Then Exit Sub
bSkipCheckForceEnable = True

If val(txtAC.Text) <> 9999 Then GoTo active:
If val(txtDR.Text) <> 9999 Then GoTo active:
If val(txtMR.Text) <> 9999 Then GoTo active:
If val(txtBSDef.Text) <> 9999 Then GoTo active:
If val(txtGameLimit.Text) <> 9999 Then GoTo active:
If val(txtLairEXP.Text) <> 0 Then GoTo active:
If val(txtAtkAccuracyMaj.Text) <> 9999 Then GoTo active:
If val(txtAtkAccuracyMax.Text) <> 9999 Then GoTo active:
If val(txtDodge.Text) <> 9999 Then GoTo active:
If val(txtNumLairs.Text) <> 0 Then GoTo active:
If val(txtNumMobsLTE.Text) <> 9999 Then GoTo active:
If val(txtNumMobsGTE.Text) <> 0 Then GoTo active:

If chkIsUndead.Value = 1 Then GoTo active:
If chkNonHostile_vEvil.Value = 1 Then GoTo active:
If chkIsNonHostile_vNG.Value = 1 Then GoTo active:

If Not optCash(0).Value = True Then GoTo active:
If chkAtkNoPoison.Value = 1 Then GoTo active:
If chkAtkNoConfusion.Value = 1 Then GoTo active:
If chkAtkNoFear.Value = 1 Then GoTo active:

For x = 0 To 2
    If cmbAbilities(x).ListIndex > 0 Then GoTo active:
Next x

optEnabled(0).Value = True
GoTo out:

active:
optEnabled(1).Value = True

out:
On Error Resume Next
bSkipCheckForceEnable = False
Exit Sub
error:
Call HandleError("CheckForceEnable")
Resume out:
End Sub

Private Sub txtAC_Change()
Call CheckForceEnable
End Sub

Private Sub txtAC_GotFocus()
Call SelectAll(txtAC)
End Sub

Private Sub txtAC_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtAtkAccuracyMaj_Change()
Call CheckForceEnable
End Sub

Private Sub txtAtkAccuracyMaj_GotFocus()
Call SelectAll(txtAtkAccuracyMaj)
End Sub

Private Sub txtAtkAccuracyMax_Change()
Call CheckForceEnable
End Sub

Private Sub txtAtkAccuracyMax_GotFocus()
Call SelectAll(txtAtkAccuracyMax)
End Sub

Private Sub txtBSDef_Change()
Call CheckForceEnable
End Sub

Private Sub txtBSDef_GotFocus()
Call SelectAll(txtBSDef)
End Sub

Private Sub txtBSDef_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtDodge_Change()
Call CheckForceEnable
End Sub

Private Sub txtDodge_GotFocus()
Call SelectAll(txtDodge)
End Sub

Private Sub txtDodge_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtDR_Change()
Call CheckForceEnable
End Sub

Private Sub txtDR_GotFocus()
Call SelectAll(txtDR)
End Sub

Private Sub txtDR_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtGameLimit_Change()
Call CheckForceEnable
End Sub

Private Sub txtGameLimit_GotFocus()
Call SelectAll(txtGameLimit)
End Sub

Private Sub txtGameLimit_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtLairEXP_Change()
Call CheckForceEnable
End Sub

Private Sub txtLairEXP_GotFocus()
Call SelectAll(txtLairEXP)
End Sub

Private Sub txtLairEXP_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtMR_Change()
Call CheckForceEnable
End Sub

Private Sub txtMR_GotFocus()
Call SelectAll(txtMR)
End Sub

Private Sub txtMR_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtNumLairs_Change()
Call CheckForceEnable
End Sub

Private Sub txtNumLairs_GotFocus()
Call SelectAll(txtNumLairs)
End Sub

Private Sub txtNumLairs_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtNumMobsGTE_Change()
Call CheckForceEnable
End Sub

Private Sub txtNumMobsGTE_GotFocus()
Call SelectAll(txtNumMobsGTE)
End Sub

Private Sub txtNumMobsGTE_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtNumMobsLTE_Change()
Call CheckForceEnable
End Sub

Private Sub txtNumMobsLTE_GotFocus()
Call SelectAll(txtNumMobsLTE)
End Sub

Private Sub txtNumMobsLTE_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub
