VERSION 5.00
Begin VB.Form frmHitCalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hit Calculator"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6480
   Icon            =   "frmHitCalc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timButtonPress 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3300
      Top             =   0
   End
   Begin VB.Timer timWindowMove 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2640
      Top             =   0
   End
   Begin VB.Frame Frame4 
      Height          =   6735
      Left            =   60
      TabIndex        =   22
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton cmdMobGoto 
         Caption         =   ">"
         Height          =   315
         Left            =   5400
         TabIndex        =   50
         Top             =   2640
         Width           =   255
      End
      Begin VB.CommandButton cmdRefreshMonster 
         Height          =   315
         Index           =   1
         Left            =   300
         Picture         =   "frmHitCalc.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Refresh Current Set"
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton cmdRefreshMonster 
         Height          =   315
         Index           =   0
         Left            =   5760
         Picture         =   "frmHitCalc.frx":0F1F
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Refresh Monster Only"
         Top             =   2646
         Width           =   375
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
         Left            =   5820
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox chkShadow 
         Caption         =   "Shadow"
         Height          =   195
         Left            =   600
         TabIndex        =   15
         Top             =   5100
         Width           =   1335
      End
      Begin VB.CheckBox chkSeeHidden 
         Caption         =   "See Hidden"
         Height          =   195
         Left            =   4680
         TabIndex        =   19
         Top             =   5100
         Width           =   1335
      End
      Begin VB.ComboBox cmbEvil 
         Height          =   315
         ItemData        =   "frmHitCalc.frx":1174
         Left            =   2460
         List            =   "frmHitCalc.frx":1176
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   5040
         Width           =   1515
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
         Index           =   6
         Left            =   240
         TabIndex        =   48
         Top             =   4560
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
         Index           =   7
         Left            =   1740
         TabIndex        =   47
         Top             =   4560
         Width           =   315
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
         Left            =   600
         MaxLength       =   5
         TabIndex        =   14
         Text            =   "0"
         Top             =   4560
         Width           =   1095
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
         Index           =   8
         Left            =   2280
         TabIndex        =   45
         Top             =   4560
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
         Index           =   9
         Left            =   3780
         TabIndex        =   44
         Top             =   4560
         Width           =   315
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
         Index           =   4
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   16
         Text            =   "0"
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   495
         Left            =   240
         TabIndex        =   33
         Top             =   180
         Width           =   5895
         Begin VB.OptionButton optHitCalcType 
            Caption         =   "Backstab"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   3600
            TabIndex        =   1
            Top             =   120
            Width           =   1530
         End
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
            Height          =   300
            Index           =   0
            Left            =   1140
            TabIndex        =   0
            Top             =   150
            Value           =   -1  'True
            Width           =   1350
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   435
         Left            =   240
         TabIndex        =   32
         Top             =   1800
         Width           =   6015
         Begin VB.OptionButton optDefender 
            Caption         =   "Manual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   4680
            TabIndex        =   8
            Top             =   120
            Width           =   1230
         End
         Begin VB.OptionButton optDefender 
            Caption         =   "vs Player"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   3060
            TabIndex        =   7
            Top             =   120
            Width           =   1470
         End
         Begin VB.OptionButton optDefender 
            Caption         =   "vs Mob"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   1620
            TabIndex        =   6
            Top             =   120
            Width           =   1230
         End
         Begin VB.OptionButton optDefender 
            Caption         =   "vs Char."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   5
            Top             =   120
            Value           =   -1  'True
            Width           =   1410
         End
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
         Left            =   2280
         TabIndex        =   31
         Top             =   3600
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
         Left            =   3780
         TabIndex        =   30
         Top             =   3600
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
         Left            =   4320
         TabIndex        =   29
         Top             =   3600
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
         Left            =   5820
         TabIndex        =   28
         Top             =   3600
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
         Left            =   240
         TabIndex        =   27
         Top             =   3600
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
         Left            =   1740
         TabIndex        =   26
         Top             =   3600
         Width           =   315
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
         Index           =   5
         Left            =   4680
         MaxLength       =   5
         TabIndex        =   18
         Text            =   "0"
         Top             =   4560
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
         Left            =   4680
         MaxLength       =   5
         TabIndex        =   13
         Text            =   "10"
         ToolTipText     =   "Dodge value, not dodge %"
         Top             =   3600
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
         Index           =   0
         Left            =   600
         MaxLength       =   4
         TabIndex        =   11
         Text            =   "100"
         Top             =   3600
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
         Index           =   1
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   12
         Text            =   "0"
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   990
         Width           =   6015
         Begin VB.OptionButton optAttacker 
            Caption         =   "Current Char."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   2
            Top             =   120
            Value           =   -1  'True
            Width           =   1890
         End
         Begin VB.OptionButton optAttacker 
            Caption         =   "Monster"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   2340
            TabIndex        =   3
            Top             =   120
            Width           =   1290
         End
         Begin VB.OptionButton optAttacker 
            Caption         =   "Manual "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   3960
            TabIndex        =   4
            Top             =   120
            Width           =   1710
         End
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
         Left            =   1320
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2640
         Width           =   4035
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
         Index           =   11
         Left            =   5820
         TabIndex        =   24
         Top             =   4560
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
         Index           =   10
         Left            =   4320
         TabIndex        =   23
         Top             =   4560
         Width           =   315
      End
      Begin VB.Label lblPROT 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "vs Prot. Evil"
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
         Left            =   540
         TabIndex        =   49
         Top             =   4260
         Width           =   1230
      End
      Begin VB.Label lblVW 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "vs Vile Ward"
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
         Left            =   2445
         TabIndex        =   46
         Top             =   4260
         Width           =   1320
      End
      Begin VB.Label lblSubLeft 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Left Text"
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   240
         TabIndex        =   43
         Top             =   5640
         Width           =   1635
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   180
         X2              =   2220
         Y1              =   1740
         Y2              =   1740
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   4020
         X2              =   6120
         Y1              =   1740
         Y2              =   1740
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
         Left            =   2520
         TabIndex        =   42
         Top             =   1590
         Width           =   1230
      End
      Begin VB.Label lblBSPercep 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "vs Perception"
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
         Left            =   4515
         TabIndex        =   41
         Top             =   4260
         Width           =   1440
      End
      Begin VB.Label lblDodge 
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
         Left            =   4695
         TabIndex        =   40
         Top             =   3300
         Width           =   1080
      End
      Begin VB.Label lblACC 
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
         Height          =   480
         Left            =   360
         TabIndex        =   39
         Top             =   3060
         Width           =   1590
      End
      Begin VB.Label lblAC 
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
         Left            =   2730
         TabIndex        =   38
         Top             =   3300
         Width           =   810
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
         Left            =   2520
         TabIndex        =   37
         Top             =   780
         Width           =   1230
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   4020
         X2              =   6120
         Y1              =   930
         Y2              =   930
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   180
         X2              =   2220
         Y1              =   930
         Y2              =   930
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         Index           =   4
         X1              =   180
         X2              =   6180
         Y1              =   2400
         Y2              =   2400
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
         Left            =   240
         TabIndex        =   36
         Top             =   2700
         Width           =   900
      End
      Begin VB.Label lblMain 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hit Rate: %%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   2100
         TabIndex        =   35
         Top             =   5640
         Width           =   2175
      End
      Begin VB.Label lblSubRight 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Right Text"
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   4500
         TabIndex        =   34
         Top             =   5640
         Width           =   1635
      End
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

Dim nMonsterAccy As Long
Dim nMonsterAC As Long
Dim nMonsterDodge As Long
Dim nMonsterBSdef As Long
Dim bMonsterEvil As Boolean
Dim bMonsterSeeHidden As Boolean

Dim nDefenderAC As Long
Dim nDefenderDodge As Long
Dim nDefenderPerception As Long
Dim nDefenderProtEvil As Long
Dim nDefenderVileWard As Long
Dim nDefenderEvilness As Long
Dim bDefenderShadow As Boolean
'Dim nDefenderClass As Long

Dim nCharVileWard As Long
Dim nCharPerception As Long
Dim nCharEvilness As Long
Dim bCharShadow As Boolean

Dim ntimButtonPressCount As Long
Dim nPlayer
Dim bDontRefresh As Boolean
Dim bMouseDown As Boolean
Dim tWindowSize As WindowSizeProperties

Public Sub DoHitCalc()
On Error GoTo error:
Dim nAC As Long, nAccy As Long, nHitChance As Currency, nTotalHitPercent As Currency
Dim nDodge As Long, nDodgeChance As Currency, sArr() As String
Dim sPrint As String, nTemp As Long ', nShadow As Long
Dim nProtEv As Long, nPerception As Long, nVileWard As Long, eEvil As eEvilPoints 'nSecondaryDef As Long,
Dim bShadow As Boolean, bSeeHidden As Boolean, bBackstab As Boolean, bVSplayer As Boolean
Dim nClass As Integer, nDefense() As Long, nBSdefense As Long, bManual As Boolean
Dim nMinHit As Integer, nMaxHit As Integer

If bDontRefresh Then Exit Sub

If optHitCalcType(1).Value = True Then bBackstab = True
If optDefender(0).Value = True Or optDefender(2).Value = True Then bVSplayer = True
If optDefender(3).Value = True Then bManual = True

nAccy = Fix(val(txtHitCalc(0).Text))
nAC = Fix(val(txtHitCalc(1).Text))
nDodge = Fix(val(txtHitCalc(2).Text))

If optDefender(2).Value = True Then
    nDefenderAC = nAC
    nDefenderDodge = nDodge
End If

If txtHitCalc(3).Enabled Then
    nProtEv = Fix(val(txtHitCalc(3).Text))
    If optDefender(2).Value = True Then nDefenderProtEvil = nProtEv
End If
If txtHitCalc(4).Enabled Then
    nVileWard = Fix(val(txtHitCalc(4).Text))
    If optDefender(0).Value = True Then nCharVileWard = nVileWard
    If optDefender(2).Value = True Then nDefenderVileWard = nVileWard
End If

If txtHitCalc(5).Enabled Then
    'we can only get away with this because this two stats are never used in combination
    'they really could be combined into one var, but it didn't seem worth the confusion
    nPerception = Fix(val(txtHitCalc(5).Text))
    If optDefender(0).Value = True Then nCharPerception = nPerception
    If optDefender(2).Value = True Then nDefenderPerception = nPerception
    nBSdefense = Fix(val(txtHitCalc(5).Text))
End If

If nAccy > 9999 Then nAccy = 9999: If nAccy < 1 Then nAccy = 1
If nAC > 9999 Then nAC = 9999: If nAC < 0 Then nAC = 0
If nDodge > 9999 Then nDodge = 9999: If nDodge < -999 Then nDodge = 0
If nPerception > 9999 Then nPerception = 9999: If nPerception < 0 Then nPerception = 0
If nBSdefense > 9999 Then nBSdefense = 9999: If nBSdefense < 0 Then nBSdefense = 0

If bGreaterMUD And cmbEvil.Enabled = True Then 'cmbEvil will be disabled when vileward does not count
    Select Case cmbEvil.ListIndex
        Case 1: eEvil = e5_Criminal
        Case 2: eEvil = e7_FIEND
    End Select
    If optDefender(0).Value = True Then nCharEvilness = cmbEvil.ListIndex
    If optDefender(2).Value = True Then nDefenderEvilness = cmbEvil.ListIndex
End If

If chkShadow.Enabled Then
    If chkShadow.Value = 1 Then bShadow = True
    If optDefender(0).Value = True Then bCharShadow = bShadow
    If optDefender(2).Value = True Then bDefenderShadow = bShadow
End If

If chkSeeHidden.Enabled Then
    If chkSeeHidden.Value = 1 Then bSeeHidden = True
End If

If optDefender(0).Value = True Then nClass = frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex)

nDefense = CalculateAttackDefense(nAccy, nAC, nDodge, nBSdefense, nProtEv, 0, nPerception, _
    nVileWard, eEvil, bShadow, bSeeHidden, bBackstab, bVSplayer, nClass)

nHitChance = nDefense(0)
nDodgeChance = nDefense(1)

prin:
sPrint = "Hit: " & nHitChance & "%"
If nDodgeChance > 0 Then
    sPrint = sPrint & vbCrLf & "Dodge: " & Round(nDodgeChance) & "%"
    nTotalHitPercent = Round(nHitChance - ((nHitChance * (nDodgeChance / 100))))
    sPrint = sPrint & vbCrLf & "Overall Hit: " & nTotalHitPercent & "%"
Else
    sPrint = sPrint & vbCrLf & "Dodge: 0%"
    sPrint = sPrint & vbCrLf & "Overall Hit: " & nHitChance & "%"
End If

lblMain.Caption = sPrint

nMinHit = GetHitMin(nClass)
nMaxHit = GetHitCap()

lblSubRight.Caption = "Hit Min-Cap:" & vbCrLf
If nMinHit <> GetHitMin() Then
    lblSubRight.Caption = lblSubRight.Caption & "*" & nMinHit & "%*"
    lblSubRight.FontBold = True
Else
    lblSubRight.Caption = lblSubRight.Caption & nMinHit & "%"
    lblSubRight.FontBold = False
End If
lblSubRight.Caption = lblSubRight.Caption & " - " & nMaxHit & "%"

If bGreaterMUD Then
    lblSubRight.Caption = lblSubRight.Caption & vbCrLf & "Dodge DR-Cap:" & vbCrLf & GetDodgeCap(, True) & "% - " & GetDodgeCap() & "%"
Else
    lblSubRight.Caption = lblSubRight.Caption & vbCrLf & "Dodge Cap:" & vbCrLf & GetDodgeCap() & "%"
End If

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("DoHitCalc")
Resume out:
End Sub

Private Sub chkSeeHidden_Click()
Call DoHitCalc
End Sub

Private Sub chkShadow_Click()
Call DoHitCalc
End Sub

Private Sub cmbEvil_Click()
Call DoHitCalc
End Sub

Private Sub cmbMonsterList_Click()
On Error GoTo error:

If cmbMonsterList.ListCount < 1 Then Exit Sub
If cmbMonsterList.ListIndex < 0 Then Exit Sub
If cmbMonsterList.ItemData(cmbMonsterList.ListIndex) < 1 Then Exit Sub
If GetMonsterData(cmbMonsterList.ItemData(cmbMonsterList.ListIndex)) Then
    Call SetHitCalcVals(False, True)
'    If optAttacker(1).Value = True Then 'attacking - accy
'        txtHitCalc(0).Text = nMonsterAccy
'    End If
'    If optDefender(1).Value = True Then 'defending - ac + dodge
'        txtHitCalc(1).Text = nMonsterAC
'        txtHitCalc(2).Text = nMonsterDodge
'        If optHitCalcType(1).Value = True Then 'backstab - bs defense
'            txtHitCalc(3).Text = nMonsterBSdef
'        End If
'    End If
End If

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("cmbMonsterList_Click")
Resume out:
End Sub

Private Sub cmbMonsterList_KeyPress(KeyAscii As Integer)
'KeyAscii = AutoComplete(cmbMonsterList, KeyAscii, False)
End Sub

Public Sub SetHitCalcVals(Optional ByVal bCharacterOnly As Boolean, Optional ByVal bMonsterOnly As Boolean)
On Error GoTo error:
Dim nBSWep As Long, nBSAccyAdj As Integer, nNormAccyAdj As Integer
Dim bCharAttack As Boolean, bMobAttack As Boolean, bManualAttack As Boolean
Dim bVSchar As Boolean, bVSmob As Boolean, bVSplayer As Boolean, bVSmanual As Boolean
Dim bNormal As Boolean, bBackstab As Boolean

lblSubLeft.Caption = ""
lblSubRight.Caption = ""

If optHitCalcType(0).Value = True Then
    bNormal = True
ElseIf optHitCalcType(1).Value = True Then
    bBackstab = True
Else
    Exit Sub
End If

If optAttacker(0).Value = True Then
    bCharAttack = True
ElseIf optAttacker(1).Value = True Then
    bMobAttack = True
ElseIf optAttacker(2).Value = True Then
    bManualAttack = True
Else
    Exit Sub
End If

If optDefender(0).Value = True Then
    bVSchar = True
ElseIf optDefender(1).Value = True Then
    bVSmob = True
ElseIf optDefender(2).Value = True Then
    bVSplayer = True
ElseIf optDefender(3).Value = True Then
    bVSmanual = True
Else
    Exit Sub
End If

bDontRefresh = True

If bGreaterMUD And lblVW.Visible = False Then
    lblVW.Visible = True
    txtHitCalc(4).Visible = True
    cmbEvil.Visible = True
    cmdCharHitCalc(8).Visible = True
    cmdCharHitCalc(9).Visible = True
ElseIf Not bGreaterMUD And lblVW.Visible = True Then
    lblVW.Visible = False
    txtHitCalc(4).Visible = False
    cmbEvil.Visible = False
    cmdCharHitCalc(8).Visible = False
    cmdCharHitCalc(9).Visible = False
End If

'ATTACKER / ACCY
If bBackstab And bCharAttack And Not bMonsterOnly Then 'bs + current character
    
    If bGlobalAttackBackstab And nGlobalAttackBackstabWeapon > 0 Then
        nBSWep = nGlobalAttackBackstabWeapon
    ElseIf Not bGlobalAttackBackstab Or nGlobalAttackBackstabWeapon = 0 Then
        nBSWep = nGlobalCharWeaponNumber(0)
    End If
    
    'not currently accounting for removal of shield if new bs weapon is two hander...
    If nBSWep <> nGlobalCharWeaponNumber(0) Then
        If nBSWep > 0 Then
            nNormAccyAdj = ItemHasAbility(nBSWep, 22)
            If nNormAccyAdj < 0 Then nNormAccyAdj = 0
            nBSAccyAdj = ItemHasAbility(nBSWep, 116)
            If nBSAccyAdj < 0 Then nBSAccyAdj = 0
        End If
        nNormAccyAdj = nNormAccyAdj - nGlobalCharWeaponAccy(0)
        nBSAccyAdj = nBSAccyAdj - nGlobalCharWeaponBSaccy(0)
    End If
    
    txtHitCalc(0).Text = CalculateBackstabAccuracy(val(frmMain.lblInvenCharStat(19).Tag), val(frmMain.txtCharStats(3).Tag), _
        val(frmMain.lblInvenCharStat(13).Tag) + nBSAccyAdj, _
        GetClassStealth(frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex)), _
        nGlobalCharAccyAbils + nGlobalCharAccyOther + nNormAccyAdj, val(frmMain.txtGlobalLevel(0).Text), val(frmMain.txtCharStats(0).Tag), GetItemStrReq(nBSWep))
    
    lblSubLeft.Caption = "BS Accuracy Calculated for Char."
    
ElseIf bNormal And bCharAttack And Not bMonsterOnly Then 'normal + current character
    
    txtHitCalc(0).Text = val(frmMain.lblInvenCharStat(10).Tag)  'acc
    
    If nGlobalAttackTypeMME = a6_PhysBash Then
        txtHitCalc(0).Text = val(txtHitCalc(0).Text) + nGlobalAttackAccyAdj
        lblSubLeft.Caption = AutoAppend(lblSubLeft.Caption, "Bash " & nGlobalAttackAccyAdj & " acc", vbCrLf)
    ElseIf nGlobalAttackTypeMME = a7_PhysSmash Then
        txtHitCalc(0).Text = val(txtHitCalc(0).Text) + nGlobalAttackAccyAdj
        lblSubLeft.Caption = AutoAppend(lblSubLeft.Caption, "Smash " & nGlobalAttackAccyAdj & " acc", vbCrLf)
    ElseIf nGlobalAttackTypeMME = a4_MartialArts Then
        If nGlobalAttackMA = 2 Then 'kick
            txtHitCalc(0).Text = val(txtHitCalc(0).Text) + nGlobalAttackAccyAdj
            lblSubLeft.Caption = AutoAppend(lblSubLeft.Caption, "Kick " & nGlobalAttackAccyAdj & " acc", vbCrLf)
        ElseIf nGlobalAttackMA = 3 Then 'jk
            txtHitCalc(0).Text = val(txtHitCalc(0).Text) + nGlobalAttackAccyAdj
            lblSubLeft.Caption = AutoAppend(lblSubLeft.Caption, "Jumpkick " & nGlobalAttackAccyAdj & " acc", vbCrLf)
        End If
    End If

ElseIf bNormal And bMobAttack And Not bCharacterOnly Then 'normal + monster

    txtHitCalc(0).Text = nMonsterAccy
    
End If

'DEFENDER - AC, DODGE, SECONDARY DEFENSE
If bVSmob And Not bCharacterOnly Then 'monster
    
    txtHitCalc(1).Text = nMonsterAC
    txtHitCalc(2).Text = nMonsterDodge
    txtHitCalc(3).Text = 0
    txtHitCalc(4).Text = 0
    lblBSPercep.Caption = "vs BS Defense"
    txtHitCalc(5).Text = nMonsterBSdef
    chkShadow.Value = 0
    cmbEvil.ListIndex = 0
    If bMonsterSeeHidden Then
        chkSeeHidden.Value = 1
    Else
        chkSeeHidden.Value = 0
    End If
    
ElseIf bVSchar And Not bMonsterOnly Then

    txtHitCalc(1).Text = Fix(val(frmMain.lblInvenCharStat(2).Tag)) 'ac
    txtHitCalc(2).Text = val(frmMain.lblInvenCharStat(8).Tag) 'dodge
    txtHitCalc(3).Text = val(frmMain.lblInvenCharStat(20).Tag) 'prot. evil
    txtHitCalc(4).Text = nCharVileWard
    txtHitCalc(5).Text = nCharPerception
    
    If bBackstab Then
        lblBSPercep.Caption = "vs Perception"
    Else
        lblBSPercep.Caption = "N/A"
    End If
    
    If bCharShadow Then
        chkShadow.Value = 1
    Else
        chkShadow.Value = 0
    End If
    
    Select Case nCharEvilness
        Case 1: cmbEvil.ListIndex = 1
        Case 2: cmbEvil.ListIndex = 2
        Case Else: cmbEvil.ListIndex = 0
    End Select
    
    chkSeeHidden.Value = 0
    
ElseIf bVSplayer And Not bMonsterOnly Then

    txtHitCalc(1).Text = nDefenderAC
    txtHitCalc(2).Text = nDefenderDodge
    txtHitCalc(3).Text = nDefenderProtEvil
    txtHitCalc(4).Text = nDefenderVileWard
    txtHitCalc(5).Text = nDefenderPerception
    
    If bBackstab Then
        lblBSPercep.Caption = "vs Perception"
    Else
        lblBSPercep.Caption = "N/A"
    End If
    
    If bDefenderShadow Then
        chkShadow.Value = 1
    Else
        chkShadow.Value = 0
    End If
    
    Select Case nDefenderEvilness
        Case 1: cmbEvil.ListIndex = 1
        Case 2: cmbEvil.ListIndex = 2
        Case Else: cmbEvil.ListIndex = 0
    End Select
    
    chkSeeHidden.Value = 0

ElseIf bCharacterOnly Or bMonsterOnly Then
    'do nothing below, dump out
    
ElseIf bVSmanual And bBackstab Then
    lblBSPercep.Caption = "vs BS Defense"

Else
    lblBSPercep.Caption = "N/A"
    
End If

If bMobAttack Or bVSmob Then
    cmbMonsterList.Enabled = True
Else
    cmbMonsterList.Enabled = False
End If

If bVSmob Or (bBackstab And Not bGreaterMUD) Then
    
    lblPROT.Enabled = False
    txtHitCalc(3).Enabled = False
    cmdCharHitCalc(6).Enabled = False
    cmdCharHitCalc(7).Enabled = False
    
    lblVW.Enabled = False
    txtHitCalc(4).Enabled = False
    cmdCharHitCalc(8).Enabled = False
    cmdCharHitCalc(9).Enabled = False
    
    If bBackstab Then
        lblBSPercep.Enabled = True
        txtHitCalc(5).Enabled = True
        cmdCharHitCalc(10).Enabled = True
        cmdCharHitCalc(11).Enabled = True
        chkSeeHidden.Enabled = True
    Else
        lblBSPercep.Enabled = False
        txtHitCalc(5).Enabled = False
        cmdCharHitCalc(10).Enabled = False
        cmdCharHitCalc(11).Enabled = False
        chkSeeHidden.Enabled = False
    End If
    chkShadow.Enabled = False
    cmbEvil.Enabled = False
    
Else 'vs char, player, or manual
    
    If bMobAttack And Not bMonsterEvil Then
        lblPROT.Enabled = False
        txtHitCalc(3).Enabled = False
        cmdCharHitCalc(6).Enabled = False
        cmdCharHitCalc(7).Enabled = False
    Else
        lblPROT.Enabled = True
        txtHitCalc(3).Enabled = True
        cmdCharHitCalc(6).Enabled = True
        cmdCharHitCalc(7).Enabled = True
    End If
    chkShadow.Enabled = True
    
    If bGreaterMUD And (bVSmanual Or Not bMobAttack Or (bMobAttack And bMonsterEvil)) Then
        lblVW.Enabled = True
        txtHitCalc(4).Enabled = True
        cmdCharHitCalc(8).Enabled = True
        cmdCharHitCalc(9).Enabled = True
        cmbEvil.Enabled = True
    Else
        lblVW.Enabled = False
        txtHitCalc(4).Enabled = False
        cmdCharHitCalc(8).Enabled = False
        cmdCharHitCalc(9).Enabled = False
        cmbEvil.Enabled = False
    End If
    
    If bBackstab And bVSmanual Then
        lblBSPercep.Enabled = True
        txtHitCalc(5).Enabled = True
        cmdCharHitCalc(10).Enabled = True
        cmdCharHitCalc(11).Enabled = True
        chkSeeHidden.Enabled = True
    Else
        lblBSPercep.Enabled = False
        txtHitCalc(5).Enabled = False
        cmdCharHitCalc(10).Enabled = False
        cmdCharHitCalc(11).Enabled = False
        chkSeeHidden.Enabled = False
    End If
    
End If

If bMobAttack And Not bMonsterEvil Then
    lblSubLeft.Caption = AutoAppend(lblSubLeft.Caption, "Mob Not Evil", ", ")
End If

out:
On Error Resume Next
bDontRefresh = False
Call DoHitCalc
Exit Sub
error:
Call HandleError("SetHitCalcVals")
Resume out:
End Sub

Private Function GetVileWardValue(ByVal nDefenderVileWard As Long, ByVal nDefenderEvilness As Long) As Long
If nDefenderVileWard > 0 Then
    If nDefenderEvilness >= 2 Then
        GetVileWardValue = nDefenderVileWard
    ElseIf nDefenderEvilness >= 1 Then
        GetVileWardValue = nDefenderVileWard \ 2
    Else
        GetVileWardValue = 0
    End If
End If
End Function

Private Sub cmdCharHitCalc_Click(Index As Integer)
If Not bMouseDown Then Call ModifyHitCalcValues(Index)
timButtonPress.Enabled = False
bMouseDown = False
End Sub

Private Sub cmdCharHitCalc_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

bMouseDown = True
ntimButtonPressCount = 1

Do While bMouseDown And ntimButtonPressCount < 101
    ntimButtonPressCount = ntimButtonPressCount + 1
    timButtonPress.Enabled = True
    Call ModifyHitCalcValues(Index)
    Do While timButtonPress.Enabled And Not bAppTerminating
        DoEvents
    Loop
Loop
bMouseDown = False

End Sub

Private Sub ModifyHitCalcValues(ByVal Index As Integer)
On Error GoTo error:

Select Case Index
    Case 0: 'acc -
        txtHitCalc(0).Text = Fix(val(txtHitCalc(0).Text)) - 1
    Case 1: 'acc +
        txtHitCalc(0).Text = Fix(val(txtHitCalc(0).Text)) + 1
    Case 2: 'ac -
        txtHitCalc(1).Text = Fix(val(txtHitCalc(1).Text)) - 1
    Case 3: 'ac +
        txtHitCalc(1).Text = Fix(val(txtHitCalc(1).Text)) + 1
    Case 4: 'dg -
        txtHitCalc(2).Text = Fix(val(txtHitCalc(2).Text)) - 1
    Case 5: 'dg +
        txtHitCalc(2).Text = Fix(val(txtHitCalc(2).Text)) + 1
    Case 6: 'prev -
        txtHitCalc(3).Text = Fix(val(txtHitCalc(3).Text)) - 1
    Case 7: 'rev +
        txtHitCalc(3).Text = Fix(val(txtHitCalc(3).Text)) + 1
    Case 8: 'vile -
        txtHitCalc(4).Text = Fix(val(txtHitCalc(4).Text)) - 1
    Case 9: 'vile +
        txtHitCalc(4).Text = Fix(val(txtHitCalc(4).Text)) + 1
    Case 10: 'precep/bs -
        txtHitCalc(5).Text = Fix(val(txtHitCalc(5).Text)) - 1
    Case 11: 'percep/bs +
        txtHitCalc(5).Text = Fix(val(txtHitCalc(5).Text)) + 1
End Select

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("ModifyHitCalcValues")
Resume out:
End Sub

Private Sub cmdCharHitCalc_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
bMouseDown = False
End Sub

Private Sub cmdMobGoto_Click()
If cmbMonsterList.ListCount < 1 Then Exit Sub
If cmbMonsterList.ListIndex < 0 Then Exit Sub
If cmbMonsterList.ItemData(cmbMonsterList.ListIndex) < 1 Then Exit Sub
Call frmMain.GotoMonster(cmbMonsterList.ItemData(cmbMonsterList.ListIndex))
On Error Resume Next
frmMain.SetFocus
End Sub

Private Sub cmdQ_Click()
Dim sTemp As String
sTemp = "Having 'current char' selected will continually update the stats in the calcualtor to match the current character's stats in MME." _
    & " The only difference between 'vs Char' and 'vs Player' and the manual options is they won't have their stats changed when MME changes."
'If bGreaterMUD Then sTemp = sTemp & vbCrLf & vbCrLf & "Note: Vile Ward is only considered vs Evil Monsters."

MsgBox sTemp, vbInformation

End Sub

Private Sub cmdRefreshMonster_Click(Index As Integer)
If Index = 0 Then
    Call cmbMonsterList_Click
Else
    Call SetHitCalcVals
End If
End Sub

Private Sub Form_Load()
On Error GoTo error:

If bAppTerminating Then Exit Sub

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

cmbEvil.AddItem "None", 0
cmbEvil.AddItem "Criminal (50%)", 1
cmbEvil.AddItem "Fiend/Villain", 2
cmbEvil.ListIndex = 0

Call LoadMonsters
Call SetHitCalcVals

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
'Call ExpandCombo(cmbMonsterList, HeightOnly, DoubleWidth, Me.hWnd)
'cmbMonsterList.SelLength = 0
Call cmbMonsterList_Click

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("LoadMonsters")
Resume out:
End Sub

Public Function GotoMonster(ByVal nMonster As Long) As Boolean
On Error GoTo error:
Dim x As Integer

If nMonster < 1 Then Exit Function

For x = 0 To cmbMonsterList.ListCount - 1
    If cmbMonsterList.ItemData(x) = nMonster Then
        cmbMonsterList.ListIndex = x
        GotoMonster = True
        Exit Function
    End If
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GotoMonster")
Resume out:
End Function

Private Function GetMonsterData(ByVal nMonster As Long) As Boolean
On Error GoTo error:
Dim x As Integer

On Error GoTo seek2:
If tabMonsters.Fields("Number") = nMonster Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabMonsters.Index = "pkMonsters"
tabMonsters.Seek "=", nMonster
If tabMonsters.NoMatch = True Then
    tabMonsters.MoveFirst
    Exit Function
End If

ready:
nMonsterAccy = 0
nMonsterDodge = 0
nMonsterAC = tabMonsters.Fields("ArmourClass")
If nNMRVer >= 1.83 Then
    nMonsterBSdef = tabMonsters.Fields("BSDefense")
Else
    nMonsterBSdef = 0
End If

bMonsterSeeHidden = False
For x = 0 To 9 'abilities
    If tabMonsters.Fields("Abil-" & x) = 34 And tabMonsters.Fields("AbilVal-" & x) > 0 Then 'dodge
        nMonsterDodge = tabMonsters.Fields("AbilVal-" & x)
    ElseIf tabMonsters.Fields("Abil-" & x) = 57 Then
        bMonsterSeeHidden = True
    End If
Next x

For x = 0 To 4 'attacks
    If tabMonsters.Fields("AttType-" & x) > 0 And tabMonsters.Fields("AttType-" & x) <= 3 And tabMonsters.Fields("Att%-" & x) > 0 Then
        Select Case tabMonsters.Fields("AttType-" & x)
            Case 1, 3  ' melee (normal/rob)
                If tabMonsters.Fields("AttAcc-" & x) > nMonsterAccy Then nMonsterAccy = tabMonsters.Fields("AttAcc-" & x)
                
        End Select
    End If
Next x

Select Case tabMonsters.Fields("Align")
    Case 1, 2, 5, 6: bMonsterEvil = True
    Case Else: bMonsterEvil = False
End Select

GetMonsterData = True

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetMonsterData")
Resume out:
End Function


Private Sub Form_Unload(Cancel As Integer)
If Not bAppTerminating Then frmMain.SetFocus
End Sub

Private Sub optAttacker_Click(Index As Integer)
Dim x As Integer
For x = 0 To 2
    If optAttacker(x).Value = True Then
        optAttacker(x).FontBold = True
    Else
        optAttacker(x).FontBold = False
    End If
Next x
Call SetHitCalcVals
End Sub

Private Sub optDefender_Click(Index As Integer)
Dim x As Integer
For x = 0 To 3
    If optDefender(x).Value = True Then
        optDefender(x).FontBold = True
    Else
        optDefender(x).FontBold = False
    End If
Next x
Call SetHitCalcVals
End Sub

Private Sub optHitCalcType_Click(Index As Integer)
Dim x As Integer

For x = 0 To 1
    If optHitCalcType(x).Value = True Then
        optHitCalcType(x).FontBold = True
    Else
        optHitCalcType(x).FontBold = False
    End If
Next x

If optHitCalcType(1).Value = True Then 'if backstab
    If optAttacker(1).Value = True Then optAttacker(2).Value = True 'no monster attacker
    optAttacker(1).Enabled = False
Else
    optAttacker(1).Enabled = True
End If

Call SetHitCalcVals

End Sub

Private Sub timWindowMove_Timer()
Call MonitorFormTimer(Me)
End Sub


Private Sub timButtonPress_Timer()
timButtonPress.Enabled = False
If bAppReallyTerminating Or bAppTerminating Then Exit Sub
End Sub

Private Sub txtHitCalc_Change(Index As Integer)
Call DoHitCalc
End Sub

Private Sub txtHitCalc_GotFocus(Index As Integer)
Call SelectAll(txtHitCalc(Index))
End Sub

Private Sub txtHitCalc_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii, False)
End Sub
