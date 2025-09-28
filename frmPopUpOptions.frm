VERSION 5.00
Begin VB.Form frmPopUpOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MMUD Explorer"
   ClientHeight    =   5220
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   8625
   ControlBox      =   0   'False
   Icon            =   "frmPopUpOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timRefreshSpellStats 
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
      Left            =   3000
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
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
      Left            =   7020
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
   Begin VB.Frame fraChooseAttack 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   4755
      Left            =   60
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   8475
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "With the updates I have coming out in the  "
         Height          =   4395
         Left            =   180
         TabIndex        =   4
         Top             =   180
         Width           =   8115
         Begin VB.ComboBox cmbBackstabWeapon 
            Height          =   315
            ItemData        =   "frmPopUpOptions.frx":0CCA
            Left            =   4800
            List            =   "frmPopUpOptions.frx":0CCC
            Sorted          =   -1  'True
            TabIndex        =   11
            Top             =   1260
            Width           =   3075
         End
         Begin VB.CheckBox chkBackstab 
            Caption         =   "+Backstab"
            Enabled         =   0   'False
            Height          =   240
            Left            =   4800
            TabIndex        =   10
            Top             =   990
            Width           =   1275
         End
         Begin VB.CommandButton cmdHelp 
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
            Index           =   0
            Left            =   7560
            TabIndex        =   71
            Top             =   120
            Width           =   375
         End
         Begin VB.TextBox txtAttackManualMagic 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   300
            Left            =   4860
            MaxLength       =   5
            TabIndex        =   16
            Text            =   "9999"
            ToolTipText     =   "vs 50 MR"
            Top             =   1860
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
            Left            =   360
            TabIndex        =   22
            Top             =   3960
            Width           =   1935
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
            Top             =   540
            Value           =   -1  'True
            Width           =   4335
         End
         Begin VB.TextBox txtAttackManual 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   300
            Left            =   3540
            MaxLength       =   5
            TabIndex        =   15
            Text            =   "9999"
            ToolTipText     =   "vs 0 AC, DR, and Dodge"
            Top             =   1860
            Width           =   675
         End
         Begin VB.ComboBox cmbAttackMA 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmPopUpOptions.frx":0CCE
            Left            =   2700
            List            =   "frmPopUpOptions.frx":0CDE
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1425
            Width           =   1515
         End
         Begin VB.TextBox txtAttackSpellLevel 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   300
            Left            =   4860
            MaxLength       =   3
            TabIndex        =   21
            Text            =   "999"
            Top             =   3540
            Width           =   615
         End
         Begin VB.ComboBox cmbAttackSpell 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            ItemData        =   "frmPopUpOptions.frx":0D04
            Left            =   660
            List            =   "frmPopUpOptions.frx":0D06
            Sorted          =   -1  'True
            TabIndex        =   20
            Top             =   3540
            Width           =   3555
         End
         Begin VB.ComboBox cmbAttackSpell 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            ItemData        =   "frmPopUpOptions.frx":0D08
            Left            =   660
            List            =   "frmPopUpOptions.frx":0D0A
            Sorted          =   -1  'True
            TabIndex        =   18
            Top             =   2760
            Width           =   3555
         End
         Begin VB.CheckBox chkSmashing 
            Caption         =   "Smash"
            Enabled         =   0   'False
            Height          =   240
            Left            =   3780
            TabIndex        =   9
            Top             =   990
            Width           =   855
         End
         Begin VB.CheckBox chkBashing 
            Caption         =   "Bash"
            Enabled         =   0   'False
            Height          =   240
            Left            =   2880
            TabIndex        =   8
            Top             =   990
            Width           =   735
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
            TabIndex        =   14
            ToolTipText     =   "(Mob defenses will not be factored)"
            Top             =   1905
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
            TabIndex        =   12
            Top             =   1440
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
            TabIndex        =   19
            Top             =   3240
            Width           =   3495
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
            TabIndex        =   17
            Top             =   2460
            Width           =   3915
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
            Top             =   960
            Width           =   2355
         End
         Begin VB.Label lblAttackSpellDMGPerRound 
            Alignment       =   2  'Center
            Caption         =   "## dmg/rnd"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   255
            Left            =   5760
            TabIndex        =   70
            Top             =   2700
            Width           =   1935
         End
         Begin VB.Line Line2 
            X1              =   360
            X2              =   7800
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Label lblAttackManaOOM 
            Alignment       =   2  'Center
            Caption         =   "## rnds to oom"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5700
            TabIndex        =   69
            Top             =   3060
            Width           =   2055
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
            TabIndex        =   67
            Top             =   1920
            Width           =   495
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
            Left            =   5640
            TabIndex        =   66
            Top             =   1920
            Width           =   615
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
            TabIndex        =   23
            Top             =   3540
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
            Top             =   60
            Width           =   7875
         End
      End
   End
   Begin VB.Frame fraRoomFind 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4755
      Left            =   60
      TabIndex        =   24
      Top             =   360
      Visible         =   0   'False
      Width           =   8475
      Begin VB.Frame fraMISC 
         BorderStyle     =   0  'None
         Caption         =   "With the updates I have coming out in the  "
         Height          =   4155
         Left            =   300
         TabIndex        =   25
         Top             =   300
         Width           =   7875
         Begin VB.OptionButton optRoomFindMatch 
            Caption         =   "Exact Match"
            Height          =   240
            Index           =   1
            Left            =   4500
            TabIndex        =   30
            Top             =   900
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
            Left            =   2820
            TabIndex        =   29
            Top             =   900
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
            Left            =   4500
            MaskColor       =   &H80000016&
            Style           =   1  'Graphical
            TabIndex        =   42
            Tag             =   "0"
            Top             =   3000
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
            Left            =   3300
            MaskColor       =   &H80000016&
            Style           =   1  'Graphical
            TabIndex        =   41
            Tag             =   "0"
            Top             =   3000
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
            Left            =   4500
            MaskColor       =   &H80000016&
            Style           =   1  'Graphical
            TabIndex        =   40
            Tag             =   "0"
            Top             =   2520
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
            Left            =   3900
            MaskColor       =   &H80000016&
            Style           =   1  'Graphical
            TabIndex        =   39
            Tag             =   "0"
            Top             =   2520
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
            Left            =   3300
            MaskColor       =   &H80000016&
            Style           =   1  'Graphical
            TabIndex        =   38
            Tag             =   "0"
            Top             =   2520
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
            Left            =   4500
            MaskColor       =   &H80000016&
            Style           =   1  'Graphical
            TabIndex        =   37
            Tag             =   "0"
            Top             =   2040
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
            Left            =   3300
            MaskColor       =   &H80000016&
            Style           =   1  'Graphical
            TabIndex        =   36
            Tag             =   "0"
            Top             =   2040
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
            Left            =   4500
            MaskColor       =   &H80000016&
            Style           =   1  'Graphical
            TabIndex        =   34
            Tag             =   "0"
            Top             =   1560
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
            Left            =   3900
            MaskColor       =   &H80000016&
            Style           =   1  'Graphical
            TabIndex        =   33
            Tag             =   "0"
            Top             =   1560
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
            Left            =   3300
            MaskColor       =   &H80000016&
            Style           =   1  'Graphical
            TabIndex        =   32
            Tag             =   "0"
            Top             =   1560
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
            Left            =   2820
            TabIndex        =   27
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Do not include invisible, hidden, or activated exits"
            Height          =   495
            Index           =   3
            Left            =   540
            TabIndex        =   35
            Top             =   1920
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
            Left            =   960
            TabIndex        =   31
            Top             =   1620
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "3 or more letters"
            Height          =   195
            Index           =   1
            Left            =   840
            TabIndex        =   28
            Top             =   780
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
            Left            =   1080
            TabIndex        =   26
            Top             =   480
            Width           =   1575
         End
      End
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
      Height          =   4755
      Left            =   60
      MaxLength       =   10000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   43
      Top             =   360
      Visible         =   0   'False
      Width           =   8505
   End
   Begin VB.Frame fraChooseHealing 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Caption         =   "s"
      Height          =   4755
      Left            =   60
      TabIndex        =   44
      Top             =   360
      Visible         =   0   'False
      Width           =   8475
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "With the updates I have coming out in the  "
         Height          =   4335
         Left            =   240
         TabIndex        =   45
         Top             =   180
         Width           =   7995
         Begin VB.CommandButton cmdHealRoundsMod 
            Caption         =   "+"
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
            Index           =   1
            Left            =   6780
            TabIndex        =   58
            Top             =   2760
            Width           =   315
         End
         Begin VB.CommandButton cmdHealRoundsMod 
            Caption         =   "-"
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
            Left            =   5760
            TabIndex        =   56
            Top             =   2760
            Width           =   315
         End
         Begin VB.CommandButton cmdHelp 
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
            Index           =   1
            Left            =   7440
            TabIndex        =   59
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton optHealingType 
            Caption         =   "Base on current char passive HP regen"
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
            Left            =   720
            TabIndex        =   47
            Top             =   960
            Width           =   4935
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
            Left            =   6120
            MaxLength       =   3
            TabIndex        =   57
            Text            =   "##"
            Top             =   2760
            Width           =   615
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
            Left            =   720
            TabIndex        =   55
            Top             =   3780
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
            Left            =   720
            TabIndex        =   50
            Top             =   2160
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
            Left            =   720
            TabIndex        =   52
            Top             =   3000
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
            Left            =   720
            TabIndex        =   48
            ToolTipText     =   "(Mob defenses will not be factored)"
            Top             =   1485
            Width           =   2895
         End
         Begin VB.ComboBox cmbHealingSpell 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            ItemData        =   "frmPopUpOptions.frx":0D0C
            Left            =   1020
            List            =   "frmPopUpOptions.frx":0D0E
            Sorted          =   -1  'True
            TabIndex        =   51
            Top             =   2520
            Width           =   2835
         End
         Begin VB.ComboBox cmbHealingSpell 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            ItemData        =   "frmPopUpOptions.frx":0D10
            Left            =   1020
            List            =   "frmPopUpOptions.frx":0D12
            Sorted          =   -1  'True
            TabIndex        =   53
            Top             =   3300
            Width           =   2835
         End
         Begin VB.TextBox txtHealingSpellLVL 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   300
            Left            =   4560
            MaxLength       =   3
            TabIndex        =   54
            Text            =   "999"
            Top             =   3300
            Width           =   615
         End
         Begin VB.TextBox txtHealingManual 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   300
            Left            =   3720
            MaxLength       =   5
            TabIndex        =   49
            Text            =   "999"
            Top             =   1455
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
            Left            =   720
            TabIndex        =   46
            Top             =   480
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.Line Line1 
            X1              =   720
            X2              =   7440
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Label lblHealManaOOM 
            Alignment       =   2  'Center
            Caption         =   "## rnds to oom"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5460
            TabIndex        =   68
            Top             =   3780
            Width           =   1875
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
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   5460
            TabIndex        =   65
            Top             =   3180
            Width           =   1875
         End
         Begin VB.Label lblHealMANAPerRound 
            Alignment       =   2  'Center
            Caption         =   "## mana/rnd"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5460
            TabIndex        =   64
            Top             =   3480
            Width           =   1875
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Cast heal every # rounds:"
            Height          =   195
            Index           =   1
            Left            =   5520
            TabIndex        =   63
            Top             =   2460
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Cast Frequency"
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
            Index           =   0
            Left            =   5520
            TabIndex        =   62
            Top             =   2160
            Width           =   1755
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
            Left            =   3960
            TabIndex        =   60
            Top             =   3300
            Width           =   495
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
            TabIndex        =   61
            Top             =   120
            Width           =   7755
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

Private bLocalDontRefresh As Boolean
'Private nLocalCurrentAttackSpellNum As Long
'Private nLocalCurrentAttackSpellCost As Long
'Private nLocalCurrentHealSpellNum As Long
'Private nLocalCurrentHealSpellCost As Long

Dim tWindowSize As WindowSizeProperties

Private Sub chkBackstab_Click()
If chkBackstab.Enabled = True And chkBackstab.Value = 1 Then
    cmbBackstabWeapon.Enabled = True
    chkBackstab.FontBold = True
Else
    cmbBackstabWeapon.Enabled = False
    If chkBackstab.Value = 0 Then chkBackstab.FontBold = False
End If
End Sub

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
timRefreshSpellStats.Enabled = False
timRefreshSpellStats.Enabled = True
End Sub

Private Sub chkSmashing_Click()
If chkSmashing.Value = 1 Then chkBashing.Value = 0
If chkBashing.Value = 1 Then chkSmashing.Value = 0
End Sub

Private Sub cmbAttackSpell_Click(Index As Integer)
timRefreshSpellStats.Enabled = False
timRefreshSpellStats.Enabled = True
End Sub

Private Sub cmbAttackSpell_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = AutoComplete(cmbAttackSpell(Index), KeyAscii, False)
End Sub

Private Sub cmbBackstabWeapon_KeyPress(KeyAscii As Integer)
KeyAscii = AutoComplete(cmbBackstabWeapon, KeyAscii, False)
End Sub

Private Sub cmbHealingSpell_Click(Index As Integer)
timRefreshSpellStats.Enabled = False
timRefreshSpellStats.Enabled = True
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
    Call RefreshSpellStats
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

Private Sub cmdHelp_Click(Index As Integer)
If Index = 0 Then 'attack
    MsgBox "Choosing [none/one-shot] will assume each mob can be killed instantly. However, " _
        & "it will not one-shot an entire room. i.e. 3 mobs would take 3 rounds and exp would be calculated " _
        & "based on that rate. Room/Area attack calculations planned for a future release." _
        & vbCrLf & vbCrLf & "Implemented attack restrictions: MagicLVL, SpellImmuLVL, AffectsUndead/Living/Animal" _
        & vbCrLf & vbCrLf & "NOT yet Impemented: Healing from Drains, Elemental Resistances" _
        & vbCrLf & vbCrLf & "Manual damage should be entered as if it where versus 0/0/0 ac/dr/dodge and 50 MR." _
        & vbCrLf & vbCrLf & "* in oom = current bless spells factored" _
                & vbCrLf & "+ in oom = current heal spell factored", vbInformation
                
ElseIf Index = 1 Then 'heal
    MsgBox "Choosing [invincible] will result in no recovery time due to hitpoints." _
        & vbCrLf & vbCrLf & "[Base on current char HP regen] will use the current character's " _
                & "resting rate to determine a per-round healing rate and will be used for time spent resting. " _
                & "These values will also be used when the spell and manual options are chosen, but only reflected in " _
                & "background calculations." _
        & vbCrLf & vbCrLf & "* in oom = current bless spells factored" _
                & vbCrLf & "+ in oom = current attack spell factored", vbInformation
End If
End Sub

Private Sub cmdHealRoundsMod_Click(Index As Integer)
On Error GoTo error:

If Index = 0 Then
    txtHealingCastNumRounds.Text = val(txtHealingCastNumRounds.Text) - 1
ElseIf Index = 1 Then
    txtHealingCastNumRounds.Text = val(txtHealingCastNumRounds.Text) + 1
End If
If val(txtHealingCastNumRounds.Text) < 1 Then txtHealingCastNumRounds.Text = 1
If val(txtHealingCastNumRounds.Text) > 99 Then txtHealingCastNumRounds.Text = 99

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("cmdHealRoundsMod_Click")
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

Public Sub SetupChooseAttack(Optional ByVal nGotoBackstab As Long, Optional ByVal nGotoSpell As Long)
On Error GoTo error:
Dim x As Integer, y As Integer, bGotoHeal As Boolean
Dim nFindSpell As Long, nFindItem As Long

fraChooseAttack.Visible = True
fraChooseHealing.Visible = False
fraRoomFind.Visible = False

bLocalDontRefresh = True

If cmbAttackSpell(1).ListCount < 10 Or cmbHealingSpell(1).ListCount < 10 Then Call RefreshSpells
If cmbBackstabWeapon.ListCount < 10 Then Call RefreshItems

If nGotoSpell > 0 Then If SpellHasAbility(nGotoSpell, 18) >= 0 Then bGotoHeal = True

If nGlobalAttackSpellNum > 0 Or (nGotoSpell > 0 And bGotoHeal = False) Then
    If (nGotoSpell > 0 And bGotoHeal = False) Then
        nFindSpell = nGotoSpell
    Else
        nFindSpell = nGlobalAttackSpellNum
    End If
    For y = 1 To 2
        For x = 0 To cmbAttackSpell(y - 1).ListCount - 1
            If cmbAttackSpell(y - 1).ItemData(x) = nFindSpell Then
                cmbAttackSpell(y - 1).ListIndex = x
                Exit For
            End If
        Next x
    Next y
End If

If nGlobalAttackSpellLVL > 0 Then
    txtAttackSpellLevel.Text = nGlobalAttackSpellLVL
ElseIf frmMain.chkGlobalFilter.Value = 1 And val(frmMain.txtGlobalLevel(0).Text) > 0 Then
    txtAttackSpellLevel.Text = val(frmMain.txtGlobalLevel(0).Text)
End If

If nGlobalAttackMA > 0 And nGlobalAttackMA <= 3 Then
    cmbAttackMA.ListIndex = nGlobalAttackMA
Else
    cmbAttackMA.ListIndex = 0
End If

If bGlobalAttackBackstab Then
    chkBackstab.Value = 1
Else
    chkBackstab.Value = 0
End If
If nGlobalAttackBackstabWeapon > 0 Or nGotoBackstab > 0 Then
    If nGotoBackstab > 0 Then
        nFindItem = nGotoBackstab
    Else
        nFindItem = nGlobalAttackBackstabWeapon
    End If
    For x = 0 To cmbBackstabWeapon.ListCount - 1
        If cmbBackstabWeapon.ItemData(x) = nFindItem Then
            cmbBackstabWeapon.ListIndex = x
            Exit For
        End If
    Next x
    If Not cmbBackstabWeapon.ListIndex = x Then cmbBackstabWeapon.ListIndex = 0
Else
    cmbBackstabWeapon.ListIndex = 0
End If

txtAttackManual.Text = nGlobalAttackManualP
txtAttackManualMagic.Text = nGlobalAttackManualM

If nGlobalAttackTypeMME = a6_PhysBash Then chkBashing.Value = 1: chkSmashing.Value = 0
If nGlobalAttackTypeMME = a7_PhysSmash Then chkBashing.Value = 0: chkSmashing.Value = 1

If (nGotoSpell > 0 And bGotoHeal = False) And cmbAttackSpell(0).ItemData(cmbAttackSpell(0).ListIndex) = nGotoSpell Then
    optAttackType(2).Value = True
    Call optAttackType_Click(2)

ElseIf (nGotoSpell > 0 And bGotoHeal = False) And cmbAttackSpell(1).ItemData(cmbAttackSpell(1).ListIndex) = nGotoSpell Then
    optAttackType(3).Value = True
    Call optAttackType_Click(3)

ElseIf nGlobalAttackTypeMME = a1_PhysAttack Or nGlobalAttackTypeMME > a5_Manual Then
    If optAttackType(1).Value = True Then
        Call optAttackType_Click(1)
    Else
        optAttackType(1).Value = True
    End If
Else
    If optAttackType(nGlobalAttackTypeMME).Value = True Then
        Call optAttackType_Click(val(nGlobalAttackTypeMME))
    Else
        optAttackType(nGlobalAttackTypeMME).Value = True
    End If
End If

If nGotoBackstab > 0 And cmbBackstabWeapon.ItemData(cmbBackstabWeapon.ListIndex) = nGotoBackstab Then chkBackstab.Value = 1

If bGlobalAttackUseMeditate Then
    chkMeditate(0).Value = 1
Else
    chkMeditate(0).Value = 0
End If

If nGlobalAttackHealSpellNum > 0 Or (nGotoSpell > 0 And bGotoHeal = True) Then
    If (nGotoSpell > 0 And bGotoHeal = True) Then
        nFindSpell = nGotoSpell
    Else
        nFindSpell = nGlobalAttackHealSpellNum
    End If
    For y = 1 To 2
        For x = 0 To cmbHealingSpell(y - 1).ListCount - 1
            If cmbHealingSpell(y - 1).ItemData(x) = nFindSpell Then
                cmbHealingSpell(y - 1).ListIndex = x
                Exit For
            End If
        Next x
    Next y
End If

If nGlobalAttackHealSpellLVL > 0 Then
    txtHealingSpellLVL.Text = nGlobalAttackHealSpellLVL
ElseIf frmMain.chkGlobalFilter.Value = 1 And val(frmMain.txtGlobalLevel(0).Text) > 0 Then
    txtHealingSpellLVL.Text = val(frmMain.txtGlobalLevel(0).Text)
End If

txtHealingCastNumRounds.Text = nGlobalAttackHealRounds
txtHealingManual.Text = nGlobalAttackHealManual

If (nGotoSpell > 0 And bGotoHeal = True) And cmbHealingSpell(0).ItemData(cmbHealingSpell(0).ListIndex) = nGotoSpell Then
    optHealingType(2).Value = True
    Call optHealingType_Click(2)

ElseIf (nGotoSpell > 0 And bGotoHeal = True) And cmbHealingSpell(1).ItemData(cmbHealingSpell(1).ListIndex) = nGotoSpell Then
    optHealingType(3).Value = True
    Call optHealingType_Click(3)
Else
    optHealingType(nGlobalAttackHealType).Value = True
    Call optHealingType_Click(nGlobalAttackHealType)
End If

Call RefreshSpellStats

out:
On Error Resume Next
timRefreshSpellStats.Enabled = False
bLocalDontRefresh = False
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

timWindowMove.Enabled = True

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
    tabSpells.MoveFirst
End If

If cmbAttackSpell(0).ListCount > 0 Then
    cmbAttackSpell(0).ListIndex = 0
    Call AutoSizeDropDownWidth(cmbAttackSpell(0))
    Call ExpandCombo(cmbAttackSpell(0), HeightOnly, DoubleWidth, Frame2.hWnd)
    cmbAttackSpell(0).SelLength = 0
End If
If cmbAttackSpell(1).ListCount > 0 Then
    cmbAttackSpell(1).ListIndex = 0
    Call AutoSizeDropDownWidth(cmbAttackSpell(1))
    Call ExpandCombo(cmbAttackSpell(1), HeightOnly, DoubleWidth, Frame2.hWnd)
    cmbAttackSpell(1).SelLength = 0
End If
If cmbHealingSpell(0).ListCount > 0 Then
    cmbHealingSpell(0).ListIndex = 0
    Call AutoSizeDropDownWidth(cmbHealingSpell(0))
    Call ExpandCombo(cmbHealingSpell(0), HeightOnly, DoubleWidth, Frame2.hWnd)
    cmbHealingSpell(0).SelLength = 0
End If
If cmbHealingSpell(1).ListCount > 0 Then
    cmbHealingSpell(1).ListIndex = 0
    Call AutoSizeDropDownWidth(cmbHealingSpell(1))
    Call ExpandCombo(cmbHealingSpell(1), HeightOnly, DoubleWidth, Frame2.hWnd)
    cmbHealingSpell(1).SelLength = 0
End If

cmbAttackSpell(0).AddItem "Select...", 0
cmbAttackSpell(1).AddItem "Select...", 0
cmbAttackSpell(0).ListIndex = 0
cmbAttackSpell(1).ListIndex = 0

cmbHealingSpell(0).AddItem "Select...", 0
cmbHealingSpell(1).AddItem "Select...", 0
cmbHealingSpell(0).ListIndex = 0
cmbHealingSpell(1).ListIndex = 0

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("RefreshAttackSpells")
Resume out:
End Sub

Private Sub RefreshItems()
On Error GoTo error:
Dim x As Integer, bHasBS As Boolean

cmbBackstabWeapon.clear
If Not tabItems.RecordCount = 0 Then
    tabItems.MoveFirst
    Do Until tabItems.EOF
        bHasBS = False
        If bOnlyInGame And tabItems.Fields("In Game") = 0 Then GoTo skip:
        If tabItems.Fields("ItemType") = 1 Then
            For x = 0 To 19
                If tabItems.Fields("Abil-" & x) = 116 Then 'bs accu
                    bHasBS = True
                    Exit For
                End If
            Next x
            If bHasBS Then
                cmbBackstabWeapon.AddItem (tabItems.Fields("Name") & " (" & tabItems.Fields("Number") & ")")
                cmbBackstabWeapon.ItemData(cmbBackstabWeapon.NewIndex) = tabItems.Fields("Number")
            End If
        End If
skip:
        tabItems.MoveNext
    Loop
    tabItems.MoveFirst
End If

If cmbBackstabWeapon.ListCount > 0 Then
    cmbBackstabWeapon.ListIndex = 0
    Call AutoSizeDropDownWidth(cmbBackstabWeapon)
    Call ExpandCombo(cmbBackstabWeapon, HeightOnly, DoubleWidth, Frame2.hWnd)
    cmbBackstabWeapon.SelLength = 0
End If

cmbBackstabWeapon.AddItem "[No Weapon (Punch)]", 0
cmbBackstabWeapon.ItemData(0) = -1

cmbBackstabWeapon.AddItem "[Equipped Weapon]", 0
cmbBackstabWeapon.ListIndex = 0

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("RefreshAttackItems")
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

If nSelected <> 2 Then 'not learned spell
    cmbAttackSpell(0).Enabled = False
End If

If nSelected <> 3 Then 'not any spell
    cmbAttackSpell(1).Enabled = False
    txtAttackSpellLevel.Enabled = False
    lblLabels(5).Enabled = False
End If

If nSelected <> 2 And nSelected <> 3 Then 'not mana
    lblAttackManaOOM.Visible = False
    lblAttackSpellDMGPerRound.Visible = False
    chkMeditate(0).Enabled = False
End If

If nSelected <> 4 Then cmbAttackMA.Enabled = False 'not MA

If nSelected <> 5 Then 'not manual
    txtAttackManual.Enabled = False
    txtAttackManualMagic.Enabled = False
    lblLabels(8).Enabled = False
    lblLabels(9).Enabled = False
End If

If nSelected <> 1 Then
    chkBashing.Enabled = False
    chkSmashing.Enabled = False
End If

If nSelected > 0 Then 'enable bs
    chkBackstab.Enabled = True
    If chkBackstab.Value = 1 Then
        cmbBackstabWeapon.Enabled = True
    Else
        cmbBackstabWeapon.Enabled = False
    End If
Else 'disable bs
    chkBackstab.Enabled = False
    cmbBackstabWeapon.Enabled = False
End If

Select Case nSelected
    Case 1: 'wep
        chkBashing.Enabled = True
        chkSmashing.Enabled = True
    Case 2: 'learned spell
        cmbAttackSpell(0).Enabled = True
        lblAttackManaOOM.Visible = True
        lblAttackSpellDMGPerRound.Visible = True
        chkMeditate(0).Enabled = True
    Case 3: 'any spell
        cmbAttackSpell(1).Enabled = True
        txtAttackSpellLevel.Enabled = True
        lblLabels(5).Enabled = True
        lblAttackManaOOM.Visible = True
        lblAttackSpellDMGPerRound.Visible = True
        chkMeditate(0).Enabled = True
    Case 4: 'ma
        cmbAttackMA.Enabled = True
    Case 5: 'manual
        txtAttackManual.Enabled = True
        txtAttackManual.Enabled = True
        txtAttackManualMagic.Enabled = True
        lblLabels(8).Enabled = True
        lblLabels(9).Enabled = True
End Select

timRefreshSpellStats.Enabled = False
timRefreshSpellStats.Enabled = True

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
        cmdHealRoundsMod(0).Enabled = False
        cmdHealRoundsMod(1).Enabled = False
        'lblHealHEALSPerRound.Enabled = False
        'lblHealMANAPerRound.Enabled = False
        txtHealingManual.Enabled = False
    Case 2, 3: 'spell
        cmbHealingSpell(0).Enabled = True
        cmbHealingSpell(1).Enabled = True
        txtHealingSpellLVL.Enabled = True
        chkMeditate(1).Enabled = True
        txtHealingCastNumRounds.Enabled = True
        cmdHealRoundsMod(0).Enabled = True
        cmdHealRoundsMod(1).Enabled = True
        'lblHealHEALSPerRound.Enabled = True
        'lblHealMANAPerRound.Enabled = True
        txtHealingManual.Enabled = False
    Case 4: 'manual
        cmbHealingSpell(0).Enabled = False
        cmbHealingSpell(1).Enabled = False
        txtHealingSpellLVL.Enabled = False
        chkMeditate(1).Enabled = False
        txtHealingCastNumRounds.Enabled = False
        cmdHealRoundsMod(0).Enabled = False
        cmdHealRoundsMod(1).Enabled = False
        'lblHealHEALSPerRound.Enabled = False
        'lblHealMANAPerRound.Enabled = False
        txtHealingManual.Enabled = True
End Select

timRefreshSpellStats.Enabled = False
timRefreshSpellStats.Enabled = True

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("optHealingType_Click")
Resume out:
End Sub

Private Sub RefreshSpellStats()
Dim nLocalHealSpellLVL As Integer, nLocalHealCost As Double, nLocalHealSpellNum As Long
Dim nLocalHealType As Long, bLocalUseMeditate As Boolean, nLocalHealRounds As Integer
Dim nLocalHealManual As Long, nLocalHealValue As Long, nLocalHealRoundsOOM As Integer
Dim tHealSpell As tSpellCastValues, tAttackSpell As tSpellCastValues
Dim x As Integer, nCharHeal As Double, nLocalAttackRoundsOOM As Integer
Dim nLocalAttackSpellLVL As Long, nLocalAttackSpellCost As Double, nLocalAttackSpellNum As Long
Dim nLocalAttackType As Integer, nLocalAttackSpellManual As Long, nLocalAttackSpellValue As Long
Dim tChar As tCharacterProfile
On Error GoTo error:

For x = 0 To 5
    If optAttackType(x).Value = True Then
        Select Case x
            Case 0, 1, 4: 'one-shot, weapon, martial-arts
                nLocalAttackType = x
            Case 2, 3: 'spell/any
                nLocalAttackSpellNum = 0
                If cmbAttackSpell(x - 2).ListIndex > 0 Then
                    If cmbAttackSpell(x - 2).ItemData(cmbAttackSpell(x - 2).ListIndex) > 0 Then
                        nLocalAttackSpellNum = cmbAttackSpell(x - 2).ItemData(cmbAttackSpell(x - 2).ListIndex)
                        nLocalAttackSpellCost = GetSpellManaCost(nLocalAttackSpellNum)
                        If x = 3 Then 'any spell
                            nLocalAttackSpellLVL = val(txtAttackSpellLevel.Text)
                            If nLocalAttackSpellLVL < 0 Then nLocalAttackSpellLVL = 0
                            If nLocalAttackSpellLVL > 9999 Then nLocalAttackSpellLVL = 9999
                        End If
                    Else
                        GoTo out_attack:
                    End If
                Else
                    GoTo out_attack:
                End If
                If nLocalAttackSpellNum > 0 Then
                    nLocalAttackType = x
                    If chkMeditate(1).Value = 1 Then
                        bLocalUseMeditate = True
                    Else
                        bLocalUseMeditate = False
                    End If
                End If
            Case 5: 'manual
                nLocalAttackType = x
                nLocalAttackSpellManual = val(txtAttackManualMagic.Text)
        End Select
    End If
Next x
out_attack:

nCharHeal = val(frmMain.lblCharRestRate.Tag) / 18
nLocalHealRounds = 1
For x = 0 To 4
    If optHealingType(x).Value = True Then
        Select Case x
            Case 0, 1: 'infinite/none
                nLocalHealType = x
            Case 2, 3: 'spell/any
                nLocalHealSpellNum = 0
                If cmbHealingSpell(x - 2).ListIndex > 0 Then
                    If cmbHealingSpell(x - 2).ItemData(cmbHealingSpell(x - 2).ListIndex) > 0 Then
                        nLocalHealSpellNum = cmbHealingSpell(x - 2).ItemData(cmbHealingSpell(x - 2).ListIndex)
                        If x = 3 Then 'any spell
                            nLocalHealSpellLVL = val(txtHealingSpellLVL.Text)
                            If nLocalHealSpellLVL < 0 Then nLocalHealSpellLVL = 0
                            If nLocalHealSpellLVL > 9999 Then nLocalHealSpellLVL = 9999
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
                    
                    nLocalHealRounds = val(txtHealingCastNumRounds.Text)
                    If nLocalHealRounds < 1 Then nLocalHealRounds = 1
                    If nLocalHealRounds > 50 Then nLocalHealRounds = 50
                End If
                
            Case 4: 'manual
                nLocalHealType = x
                nLocalHealManual = val(txtHealingManual.Text)
        End Select
    End If
Next x
out_heal:

tChar.nMaxMana = val(frmMain.lblCharMaxMana.Tag)
tChar.nManaRegen = val(frmMain.lblCharManaRate.Tag)
tChar.nSpellcasting = val(frmMain.lblCharSC.Tag)
tChar.nSpellDmgBonus = val(frmMain.lblInvenCharStat(33).Tag)

'calc heals
Select Case nLocalHealType
    Case 0: 'infinite
        nLocalHealValue = 99999
    Case 1: 'base
        nLocalHealValue = nCharHeal
    Case 2, 3: 'spell
        If nLocalHealSpellNum > 0 Then
            tChar.nLevel = IIf(nLocalHealType = 3, nLocalHealSpellLVL, val(frmMain.txtGlobalLevel(0).Text))
            tChar.nSpellOverhead = nLocalAttackSpellCost + (val(frmMain.lblCharBless.Caption) / 6)
            
            tHealSpell = CalculateSpellCast(tChar, nLocalHealSpellNum, tChar.nLevel)
            
            nLocalHealValue = Round((tHealSpell.nAvgCast / nLocalHealRounds), 2)
            nLocalHealCost = Round(tHealSpell.nManaCost / nLocalHealRounds, 2)
            If nLocalHealCost < 0.1 Then nLocalHealCost = 0
            If nLocalHealCost > 9999 Then nLocalHealCost = 9999
            If nLocalHealCost > 0 Then
                nLocalHealRoundsOOM = CalcRoundsToOOM(nLocalHealCost + nLocalAttackSpellCost + (val(frmMain.lblCharBless.Caption) / 6), _
                                    val(frmMain.lblCharMaxMana.Tag), val(frmMain.lblCharManaRate.Tag), tHealSpell.nCastChance, tHealSpell.nDuration)
            End If
        End If
    Case 4: 'manual
        nLocalHealValue = nLocalHealManual
End Select
If nLocalHealValue < 0 Then nLocalHealValue = 0
If nLocalHealValue > 99999 Then nLocalHealValue = 99999

'calc spell ttack
Select Case nLocalAttackType
    Case 2, 3: 'spell
        If nLocalAttackSpellNum > 0 Then
            tChar.nLevel = IIf(nLocalAttackType = 3, nLocalAttackSpellLVL, val(frmMain.txtGlobalLevel(0).Text))
            tChar.nSpellOverhead = nLocalHealCost + (val(frmMain.lblCharBless.Caption) / 6)
            
            tAttackSpell = CalculateSpellCast(tChar, nLocalAttackSpellNum, tChar.nLevel)
            
            nLocalAttackSpellValue = tAttackSpell.nAvgCast
            nLocalAttackRoundsOOM = tAttackSpell.nOOM
        End If
    Case 5: 'manual
        nLocalAttackSpellValue = nLocalAttackSpellManual
End Select
If nLocalAttackSpellValue < 0 Then nLocalAttackSpellManual = 0
If nLocalAttackSpellValue > 99999 Then nLocalAttackSpellManual = 99999

'heal captions
lblHealHEALSPerRound.Caption = nLocalHealValue & " heal/rnd"
lblHealMANAPerRound.Caption = nLocalHealCost & " mana/rnd"
If nLocalHealRoundsOOM > 0 Then
    lblHealManaOOM.Caption = nLocalHealRoundsOOM & " rnds to oom"
    If val(frmMain.lblCharBless.Caption) > 0 Then lblHealManaOOM.Caption = lblHealManaOOM.Caption & "*"
    If nLocalAttackSpellCost > 0 Then lblHealManaOOM.Caption = lblHealManaOOM.Caption & "+"
ElseIf nLocalAttackRoundsOOM > 0 Then
    lblHealManaOOM.Caption = nLocalAttackRoundsOOM & " rnds to oom"
    If val(frmMain.lblCharBless.Caption) > 0 Then lblHealManaOOM.Caption = lblHealManaOOM.Caption & "*"
    lblHealManaOOM.Caption = lblHealManaOOM.Caption & "+"
Else
    lblHealManaOOM.Caption = ""
End If

'spell attack captions
lblAttackSpellDMGPerRound.Caption = nLocalAttackSpellValue & " dmg/rnd"
If nLocalAttackRoundsOOM > 0 Then
    lblAttackManaOOM.Caption = nLocalAttackRoundsOOM & " rnds to oom"
    If val(frmMain.lblCharBless.Caption) > 0 Then lblAttackManaOOM.Caption = lblAttackManaOOM.Caption & "*"
    If nLocalHealCost > 0 Then lblAttackManaOOM.Caption = lblAttackManaOOM.Caption & "+"
Else
    lblAttackManaOOM.Caption = ""
End If

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

Private Sub timRefreshSpellStats_Timer()
timRefreshSpellStats.Enabled = False
If bLocalDontRefresh Then Exit Sub
Call RefreshSpellStats
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

Private Sub txtAttackManualMagic_Change()
timRefreshSpellStats.Enabled = False
timRefreshSpellStats.Enabled = True
End Sub

Private Sub txtAttackManualMagic_GotFocus()
Call SelectAll(txtAttackManualMagic)
End Sub

Private Sub txtAttackManualMagic_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtAttackSpellLevel_Change()
timRefreshSpellStats.Enabled = False
timRefreshSpellStats.Enabled = True
End Sub

Private Sub txtAttackSpellLevel_GotFocus()
Call SelectAll(txtAttackSpellLevel)
End Sub

Private Sub txtAttackSpellLevel_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtHealingCastNumRounds_Change()
timRefreshSpellStats.Enabled = False
timRefreshSpellStats.Enabled = True
End Sub

Private Sub txtHealingCastNumRounds_GotFocus()
Call SelectAll(txtHealingCastNumRounds)
End Sub

Private Sub txtHealingCastNumRounds_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtHealingManual_Change()
timRefreshSpellStats.Enabled = False
timRefreshSpellStats.Enabled = True
End Sub

Private Sub txtHealingManual_GotFocus()
Call SelectAll(txtHealingManual)
End Sub

Private Sub txtHealingManual_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtHealingSpellLVL_Change()
timRefreshSpellStats.Enabled = False
timRefreshSpellStats.Enabled = True
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
