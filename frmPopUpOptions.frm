VERSION 5.00
Begin VB.Form frmPopUpOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MMUD Explorer"
   ClientHeight    =   4335
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   7065
   ControlBox      =   0   'False
   Icon            =   "frmPopUpOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7065
   Begin VB.Timer timRefreshHeals 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1800
      Top             =   60
   End
   Begin VB.Timer timWindowMove 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   60
   End
   Begin VB.Frame fraChooseAttack 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
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
         Begin VB.TextBox txtAttackManualMagic 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   300
            Left            =   4800
            TabIndex        =   20
            Text            =   "9999"
            ToolTipText     =   "vs 50 MR"
            Top             =   3060
            Width           =   675
         End
         Begin VB.CheckBox chkMeditate 
            Caption         =   "Use Meditate"
            Enabled         =   0   'False
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
            Left            =   4440
            TabIndex        =   12
            Top             =   1620
            Width           =   1815
         End
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
         Begin VB.TextBox txtAttackManual 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   300
            Left            =   3540
            MaxLength       =   10
            TabIndex        =   19
            Text            =   "9999"
            ToolTipText     =   "vs 0 AC, DR, and Dodge"
            Top             =   3060
            Width           =   675
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
            Left            =   4800
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
            TabIndex        =   14
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
            Caption         =   "Enter Damage Manually:"
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
            ToolTipText     =   "(Mob defenses will not be factored)"
            Top             =   3105
            Width           =   2955
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
            TabIndex        =   13
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
            Caption         =   "Phys"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   195
            Index           =   8
            Left            =   4260
            TabIndex        =   63
            Top             =   3120
            Width           =   435
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblLabels 
            Caption         =   "Spell"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   195
            Index           =   9
            Left            =   5520
            TabIndex        =   62
            Top             =   3120
            Width           =   555
            WordWrap        =   -1  'True
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
            Left            =   4380
            TabIndex        =   21
            Top             =   2220
            Width           =   375
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Caption         =   "Choose Attack"
            BeginProperty Font 
               Name            =   "Arial"
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
      TabIndex        =   22
      Top             =   360
      Visible         =   0   'False
      Width           =   6915
      Begin VB.Frame fraMISC 
         BorderStyle     =   0  'None
         Caption         =   "With the updates I have coming out in the  "
         Height          =   3495
         Left            =   240
         TabIndex        =   23
         Top             =   180
         Width           =   6435
         Begin VB.OptionButton optRoomFindMatch 
            Caption         =   "Exact Match"
            Height          =   240
            Index           =   1
            Left            =   4200
            TabIndex        =   28
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
            TabIndex        =   27
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
            TabIndex        =   40
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
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
            TabIndex        =   35
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
            TabIndex        =   34
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
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
            TabIndex        =   25
            Top             =   300
            Width           =   3015
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Do not include invisible, hidden, or activated exits"
            Height          =   495
            Index           =   3
            Left            =   240
            TabIndex        =   33
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
            TabIndex        =   29
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "3 or more letters"
            Height          =   195
            Index           =   1
            Left            =   540
            TabIndex        =   26
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
            TabIndex        =   24
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
      TabIndex        =   41
      Top             =   360
      Visible         =   0   'False
      Width           =   6945
   End
   Begin VB.Frame fraChooseHealing 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   60
      TabIndex        =   42
      Top             =   360
      Visible         =   0   'False
      Width           =   6915
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "With the updates I have coming out in the  "
         Height          =   3495
         Left            =   240
         TabIndex        =   43
         Top             =   180
         Width           =   6435
         Begin VB.CommandButton cmdHealHelp 
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
            Left            =   3960
            TabIndex        =   46
            Top             =   840
            Width           =   315
         End
         Begin VB.OptionButton optHealingType 
            Caption         =   "Base on current char HP Regen"
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
            Left            =   240
            TabIndex        =   45
            Top             =   900
            Width           =   3735
         End
         Begin VB.TextBox txtHealingCastNumRounds 
            Alignment       =   2  'Center
            Enabled         =   0   'False
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
            Left            =   5760
            TabIndex        =   50
            Text            =   "2"
            Top             =   2040
            Width           =   555
         End
         Begin VB.CheckBox chkMeditate 
            Caption         =   "Use Meditate"
            Enabled         =   0   'False
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
            Index           =   1
            Left            =   4620
            TabIndex        =   49
            Top             =   1200
            Width           =   1815
         End
         Begin VB.OptionButton optHealingType 
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
            Left            =   240
            TabIndex        =   47
            Top             =   1320
            Width           =   3855
         End
         Begin VB.OptionButton optHealingType 
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
            Left            =   240
            TabIndex        =   51
            Top             =   2160
            Width           =   3375
         End
         Begin VB.OptionButton optHealingType 
            Caption         =   "Enter Healing Manually:"
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
            Left            =   240
            TabIndex        =   54
            ToolTipText     =   "(Mob defenses will not be factored)"
            Top             =   3045
            Width           =   2835
         End
         Begin VB.ComboBox cmbHealingSpell 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            ItemData        =   "frmPopUpOptions.frx":0D08
            Left            =   540
            List            =   "frmPopUpOptions.frx":0D0A
            Sorted          =   -1  'True
            TabIndex        =   48
            Top             =   1620
            Width           =   2595
         End
         Begin VB.ComboBox cmbHealingSpell 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            ItemData        =   "frmPopUpOptions.frx":0D0C
            Left            =   540
            List            =   "frmPopUpOptions.frx":0D0E
            Sorted          =   -1  'True
            TabIndex        =   52
            Top             =   2460
            Width           =   2595
         End
         Begin VB.TextBox txtHealingSpellLVL 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   300
            Left            =   3720
            TabIndex        =   53
            Text            =   "999"
            Top             =   2460
            Width           =   555
         End
         Begin VB.TextBox txtHealingManual 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   300
            Left            =   3240
            TabIndex        =   55
            Text            =   "999"
            Top             =   3015
            Width           =   795
         End
         Begin VB.OptionButton optHealingType 
            Caption         =   "Invincible"
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
            TabIndex        =   44
            Top             =   480
            Value           =   -1  'True
            Width           =   4815
         End
         Begin VB.Label lblHealHEALSPerRound 
            Alignment       =   2  'Center
            Caption         =   "## heal/rnd"
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
            Left            =   4500
            TabIndex        =   61
            Top             =   2520
            Width           =   1875
         End
         Begin VB.Label lblHealMANAPerRound 
            Alignment       =   2  'Center
            Caption         =   "## mana/rnd"
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
            Left            =   4500
            TabIndex        =   60
            Top             =   2820
            Width           =   1875
         End
         Begin VB.Label Label1 
            Caption         =   "Cast heal every X rounds:"
            Height          =   375
            Index           =   1
            Left            =   4560
            TabIndex        =   59
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Line Line1 
            X1              =   4410
            X2              =   4410
            Y1              =   3000
            Y2              =   1200
         End
         Begin VB.Label Label1 
            Caption         =   "Cast Frequency:"
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
            Index           =   0
            Left            =   4560
            TabIndex        =   58
            Top             =   1680
            Width           =   1755
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Caption         =   "Choose Healing"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   6
            Left            =   120
            TabIndex        =   57
            Top             =   0
            Width           =   6195
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
            Index           =   4
            Left            =   3240
            TabIndex        =   56
            Top             =   2460
            Width           =   375
         End
      End
   End
End
Attribute VB_Name = "frmPopUpOptions"
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

Dim tWindowSize As WindowSizeProperties

Private Sub chkBashing_Click()
If chkBashing.Value = 1 Then chkSmashing.Value = 0
If chkSmashing.Value = 1 Then chkBashing.Value = 0
End Sub


Private Sub chkMeditate_Click(Index As Integer)
If Index = 0 Then
    If chkMeditate(0).Value <> chkMeditate(1).Value Then chkMeditate(1).Value = chkMeditate(0).Value
Else
    If chkMeditate(1).Value <> chkMeditate(0).Value Then chkMeditate(0).Value = chkMeditate(1).Value
End If
End Sub

Private Sub chkSmashing_Click()
If chkSmashing.Value = 1 Then chkBashing.Value = 0
If chkBashing.Value = 1 Then chkSmashing.Value = 0
End Sub

Private Sub cmbAttackSpell_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = AutoComplete(cmbAttackSpell(Index), KeyAscii, False)
End Sub

Private Sub cmbHealingSpell_Click(Index As Integer)
timRefreshHeals.Enabled = False
timRefreshHeals.Enabled = True
End Sub

Private Sub cmbHealingSpell_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = AutoComplete(cmbHealingSpell(Index), KeyAscii, False)
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

If Me.Tag = "attack" And nNMRVer >= 1.83 Then
    fraChooseHealing.Visible = True
    fraChooseAttack.Visible = False
    Me.Tag = "-1"
    Exit Sub
End If

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

Private Sub cmdHealHelp_Click()
MsgBox "Char HP Regen will also be added to the spell and manual options (in the background when calculating).", vbInformation
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
Dim x As Integer, Y As Integer

fraChooseAttack.Visible = True
fraChooseHealing.Visible = False
fraRoomFind.Visible = False

Call RefreshSpells

If nCurrentAttackSpellNum > 0 Then
    For Y = 1 To 2
        For x = 0 To cmbAttackSpell(Y - 1).ListCount - 1
            If cmbAttackSpell(Y - 1).ItemData(x) = nCurrentAttackSpellNum Then
                cmbAttackSpell(Y - 1).ListIndex = x
                Exit For
            End If
        Next x
    Next Y
End If

If nCurrentAttackSpellLVL > 0 Then
    txtAttackSpellLevel.Text = nCurrentAttackSpellLVL
ElseIf frmMain.chkGlobalFilter.Value = 1 And val(frmMain.txtGlobalLevel(0).Text) > 0 Then
    txtAttackSpellLevel.Text = val(frmMain.txtGlobalLevel(0).Text)
End If

If nCurrentAttackMA > 0 And nCurrentAttackMA <= 3 Then cmbAttackMA.ListIndex = nCurrentAttackMA

txtAttackManual.Text = nCurrentAttackManual
txtAttackManualMagic.Text = nCurrentAttackManualMag

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

If bCurrentAttackUseMeditate Then
    chkMeditate(0).Value = 1
Else
    chkMeditate(0).Value = 0
End If

If nCurrentAttackHealSpellNum > 0 Then
    For Y = 1 To 2
        For x = 0 To cmbHealingSpell(Y - 1).ListCount - 1
            If cmbHealingSpell(Y - 1).ItemData(x) = nCurrentAttackHealSpellNum Then
                cmbHealingSpell(Y - 1).ListIndex = x
                Exit For
            End If
        Next x
    Next Y
End If

If nCurrentAttackHealSpellLVL > 0 Then
    txtHealingSpellLVL.Text = nCurrentAttackHealSpellLVL
ElseIf frmMain.chkGlobalFilter.Value = 1 And val(frmMain.txtGlobalLevel(0).Text) > 0 Then
    txtHealingSpellLVL.Text = val(frmMain.txtGlobalLevel(0).Text)
End If

txtHealingCastNumRounds.Text = nCurrentAttackHealRounds
txtHealingManual.Text = nCurrentAttackHealManual

optHealingType(nCurrentAttackHealType).Value = True
Call optHealingType_Click(nCurrentAttackHealType)

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

If nNMRVer < 1.83 Then
    chkMeditate(0).Visible = False
    chkMeditate(1).Visible = False
End If

cmbAttackMA.ListIndex = 0

timWindowMove.Enabled = True
timRefreshHeals.Enabled = True

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("Form_Load")
Resume out:
End Sub

Private Sub RefreshSpells()
On Error GoTo error:
Dim x As Integer, bHasDmg As Boolean, bHasHeal As Boolean

cmbAttackSpell(0).clear
cmbAttackSpell(1).clear
cmbHealingSpell(0).clear
cmbHealingSpell(1).clear
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
            bHasDmg = False: bHasHeal = False
            For x = 0 To 9
                Select Case tabSpells.Fields("Abil-" & x)
                    Case 1, 8, 17: '1-dmg, 8-drain, 17-dmg-mr
                        bHasDmg = True
                    Case 8, 18: '8-drain, 18-heal
                        bHasHeal = True
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
            
            If bHasHeal Then
                If in_long_arr(tabSpells.Fields("Number"), nLearnedSpells()) Then
                    cmbHealingSpell(0).AddItem tabSpells.Fields("Name") & IIf(bHideRecordNumbers, "", " (" & tabSpells.Fields("Number") & ")")
                    cmbHealingSpell(0).ItemData(cmbHealingSpell(0).NewIndex) = tabSpells.Fields("Number")
                End If
                cmbHealingSpell(1).AddItem tabSpells.Fields("Name") & " (" & tabSpells.Fields("Number") & ") - LVL " & tabSpells.Fields("ReqLevel") & " " & GetMagery(tabSpells.Fields("Magery"), tabSpells.Fields("MageryLVL"))
                cmbHealingSpell(1).ItemData(cmbHealingSpell(1).NewIndex) = tabSpells.Fields("Number")
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

cmbHealingSpell(0).AddItem "Select...", 0
cmbHealingSpell(1).AddItem "Select...", 0
cmbHealingSpell(0).ListIndex = 0
cmbHealingSpell(1).ListIndex = 0

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

txtText.Width = Me.ScaleWidth - txtText.Left - 50
txtText.Height = Me.ScaleHeight - txtText.Top - 50

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
        txtAttackManual.Enabled = True
        txtAttackManual.Enabled = True
        txtAttackManualMagic.Enabled = True
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
    chkMeditate(0).Enabled = True
End If

If optAttackType(2).Value = True Or optAttackType(3).Value = True Then
    chkMeditate(0).Enabled = True
Else
    chkMeditate(0).Enabled = False
End If

If optAttackType(4).Value = False Then cmbAttackMA.Enabled = False

If optAttackType(5).Value = False Then
    txtAttackManual.Enabled = False
    txtAttackManualMagic.Enabled = False
    lblLabels(8).Enabled = False
    lblLabels(9).Enabled = False
End If

End Sub

Private Sub optHealingType_Click(Index As Integer)
On Error GoTo error:
Dim x As Integer, nSelected As Integer

For x = 0 To 4
    If optHealingType(x).Value = True Then
        nSelected = x
    End If
Next x
If Not optHealingType(nSelected).Value = True Then optHealingType(nSelected).Value = True

Select Case nSelected
    Case 0, 1: 'infinite/none
        cmbHealingSpell(0).Enabled = False
        cmbHealingSpell(1).Enabled = False
        txtHealingSpellLVL.Enabled = False
        chkMeditate(1).Enabled = False
        txtHealingCastNumRounds.Enabled = False
        'lblHealHEALSPerRound.Enabled = False
        'lblHealMANAPerRound.Enabled = False
        txtHealingManual.Enabled = False
    Case 2, 3: 'spell
        cmbHealingSpell(0).Enabled = True
        cmbHealingSpell(1).Enabled = True
        txtHealingSpellLVL.Enabled = True
        chkMeditate(1).Enabled = True
        txtHealingCastNumRounds.Enabled = True
        'lblHealHEALSPerRound.Enabled = True
        'lblHealMANAPerRound.Enabled = True
        txtHealingManual.Enabled = False
    Case 4: 'manual
        cmbHealingSpell(0).Enabled = False
        cmbHealingSpell(1).Enabled = False
        txtHealingSpellLVL.Enabled = False
        chkMeditate(1).Enabled = False
        txtHealingCastNumRounds.Enabled = False
        'lblHealHEALSPerRound.Enabled = False
        'lblHealMANAPerRound.Enabled = False
        txtHealingManual.Enabled = True
End Select

timRefreshHeals.Enabled = False
timRefreshHeals.Enabled = True

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("optHealingType_Click")
Resume out:
End Sub

Private Sub RefreshHealingStats()
Dim nLocalHealType As Integer, nLocalHealCost As Double, nLocalHealSpellNum As Long
Dim nLocalHealSpellLVL As Long, bLocalUseMeditate As Boolean, nLocalHealRounds As Integer
Dim nLocalHealManual As Long, nLocalHealValue As Long
Dim tHealSpell As tSpellCastValues, x As Integer, nCharHeal As Double
On Error GoTo error:

nCharHeal = val(frmMain.lblCharRestRate.Tag) / 18

For x = 0 To 4
    If optHealingType(x).Value = True Then
        Select Case x
            Case 0, 1: 'infinite/none
                nLocalHealType = x
                'nLocalHealCost = 0
            Case 2, 3: 'spell/any
                'nLocalHealCost = 0
                nLocalHealSpellNum = 0
                If cmbHealingSpell(x - 2).ListIndex > 0 Then
                    If cmbHealingSpell(x - 2).ItemData(cmbHealingSpell(x - 2).ListIndex) > 0 Then
                        nLocalHealSpellNum = cmbHealingSpell(x - 2).ItemData(cmbHealingSpell(x - 2).ListIndex)
                        If x = 3 Then 'any spell
                            nLocalHealSpellLVL = val(txtHealingSpellLVL.Text)
                        End If
                    Else
                        GoTo out_heal:
                    End If
                Else
                    GoTo out_heal:
                End If
                If nLocalHealSpellNum > 0 Then
                    nLocalHealType = x
                    If chkMeditate(1).Value = 1 Then
                        bLocalUseMeditate = True
                    Else
                        bLocalUseMeditate = False
                    End If
                    
                    If nLocalHealSpellLVL < 0 Then nLocalHealSpellLVL = 0
                    If nLocalHealSpellLVL > 9999 Then nLocalHealSpellLVL = 9999
                    
                    nLocalHealRounds = val(txtHealingCastNumRounds.Text)
                    If nLocalHealRounds < 1 Then nLocalHealRounds = 1
                    If nLocalHealRounds > 50 Then nLocalHealRounds = 50
                    
                    'nLocalHealCost = GetSpellManaCost(nLocalHealSpellNum)
                    'nLocalHealCost = Round(nLocalHealCost / nLocalHealRounds, 1)
                    'If nLocalHealCost < 0.25 Then nLocalHealCost = 0
                    'If nLocalHealCost > 9999 Then nLocalHealCost = 9999
                End If
                
            Case 4: 'manual
                nLocalHealType = x
                'nLocalHealCost = 0
                nLocalHealManual = val(txtHealingManual.Text)
                If nLocalHealManual < 0 Then nLocalHealManual = 0
                If nLocalHealManual > 99999 Then nLocalHealManual = 99999
        End Select
    End If
Next x
out_heal:

Select Case nLocalHealType
    Case 0: 'infinite
        nLocalHealValue = 99999
    Case 1: 'base
        nLocalHealValue = nCharHeal
    Case 2, 3: 'spell
        If nLocalHealSpellNum > 0 Then
            If nLocalHealSpellLVL < 0 Then nLocalHealSpellLVL = 0
            If nLocalHealSpellLVL > 9999 Then nLocalHealSpellLVL = 9999
            If nLocalHealRounds < 1 Then nLocalHealRounds = 1
            If nLocalHealRounds > 50 Then nLocalHealRounds = 50
            
            tHealSpell = CalculateSpellCast(nLocalHealSpellNum, IIf(nLocalHealType = 3, nLocalHealSpellLVL, val(frmMain.txtGlobalLevel(0).Text)), _
                            val(frmMain.lblCharSC.Tag), , , val(frmMain.lblCharMaxMana.Tag), val(frmMain.lblCharManaRate.Tag) - val(frmMain.lblCharBless.Caption))
            nLocalHealCost = Round(tHealSpell.nManaCost / nLocalHealRounds, 2)
            nLocalHealValue = Round((tHealSpell.nAvgCast / nLocalHealRounds), 2)
        End If
    Case 4: 'manual
        nLocalHealValue = nLocalHealManual
End Select

If nLocalHealCost < 0.25 Then nLocalHealCost = 0
If nLocalHealCost > 9999 Then nLocalHealCost = 9999
If nLocalHealValue < 0 Then nLocalHealValue = 0
If nLocalHealValue > 99999 Then nLocalHealValue = 99999

lblHealHEALSPerRound.Caption = nLocalHealValue & " heal/rnd"
lblHealMANAPerRound.Caption = nLocalHealCost & " mana/rnd"

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("RefreshHealingStats")
Resume out:
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

Private Sub timRefreshHeals_Timer()
timRefreshHeals.Enabled = False
Call RefreshHealingStats
End Sub

Private Sub timWindowMove_Timer()
Call MonitorFormTimer(Me)
End Sub

'Private Sub txtAttackMag_GotFocus()
'Call SelectAll(txtAttackMag)
'End Sub
'
'Private Sub txtAttackMag_KeyPress(KeyAscii As Integer)
'KeyAscii = NumberKeysOnly(KeyAscii)
'End Sub

Private Sub txtAttackManual_GotFocus()
Call SelectAll(txtAttackManual)
End Sub

Private Sub txtAttackManual_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtAttackManualMagic_GotFocus()
Call SelectAll(txtAttackManualMagic)
End Sub

Private Sub txtAttackManualMagic_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtAttackSpellLevel_GotFocus()
Call SelectAll(txtAttackSpellLevel)
End Sub

Private Sub txtAttackSpellLevel_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtHealingCastNumRounds_Change()
timRefreshHeals.Enabled = False
timRefreshHeals.Enabled = True
End Sub

Private Sub txtHealingCastNumRounds_GotFocus()
Call SelectAll(txtHealingCastNumRounds)
End Sub

Private Sub txtHealingCastNumRounds_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtHealingManual_Change()
timRefreshHeals.Enabled = False
timRefreshHeals.Enabled = True
End Sub

Private Sub txtHealingManual_GotFocus()
Call SelectAll(txtHealingManual)
End Sub

Private Sub txtHealingManual_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtHealingSpellLVL_Change()
timRefreshHeals.Enabled = False
timRefreshHeals.Enabled = True
End Sub

Private Sub txtHealingSpellLVL_GotFocus()
Call SelectAll(txtHealingSpellLVL)
End Sub

Private Sub txtHealingSpellLVL_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtRoomName_GotFocus()
Call SelectAll(txtRoomName)
End Sub
