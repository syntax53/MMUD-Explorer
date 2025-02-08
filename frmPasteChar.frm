VERSION 5.00
Object = "{20D5284F-7B23-4F0A-B8B1-6C9D18B64F1C}#1.0#0"; "exlimiter.ocx"
Begin VB.Form frmPasteChar 
   Caption         =   "Paste Characters/Equipment/Spells"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10260
   ControlBox      =   0   'False
   Icon            =   "frmPasteChar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   10260
   Begin VB.Frame fraPasteParty 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3855
      Left            =   60
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   9855
      Begin VB.Frame fraMISC 
         BorderStyle     =   0  'None
         Caption         =   "With the updates I have coming out in the  "
         Height          =   3495
         Left            =   240
         TabIndex        =   5
         Top             =   180
         Width           =   9375
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
            TabIndex        =   121
            Top             =   120
            Width           =   315
         End
         Begin VB.TextBox txtPastePartyDMG 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   6900
            MaxLength       =   6
            TabIndex        =   120
            Top             =   3000
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyDMG 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   6900
            MaxLength       =   6
            TabIndex        =   119
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyDMG 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   6900
            MaxLength       =   6
            TabIndex        =   118
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyDMG 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   6900
            MaxLength       =   6
            TabIndex        =   117
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyDMG 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   6900
            MaxLength       =   6
            TabIndex        =   116
            Top             =   1560
            Width           =   735
         End
         Begin VB.CommandButton cmdPasteQ 
            Caption         =   "DMG"
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
            Left            =   6900
            TabIndex        =   115
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyDMG 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   6900
            MaxLength       =   6
            TabIndex        =   114
            Top             =   420
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyDMG 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   6900
            MaxLength       =   6
            TabIndex        =   113
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyHeals 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   6060
            MaxLength       =   6
            TabIndex        =   112
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyHeals 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   6060
            MaxLength       =   6
            TabIndex        =   111
            Top             =   420
            Width           =   735
         End
         Begin VB.CommandButton cmdPasteQ 
            Caption         =   "Heals"
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
            Left            =   6060
            TabIndex        =   110
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyHeals 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   6060
            MaxLength       =   6
            TabIndex        =   109
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyHeals 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   6060
            MaxLength       =   6
            TabIndex        =   108
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyHeals 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   6060
            MaxLength       =   6
            TabIndex        =   107
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyHeals 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   6060
            MaxLength       =   6
            TabIndex        =   106
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyHeals 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   6060
            MaxLength       =   6
            TabIndex        =   105
            Top             =   3000
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyRestHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   5280
            MaxLength       =   6
            TabIndex        =   104
            Top             =   3000
            Width           =   615
         End
         Begin VB.TextBox txtPastePartyRestHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   5280
            MaxLength       =   6
            TabIndex        =   103
            Top             =   2640
            Width           =   615
         End
         Begin VB.TextBox txtPastePartyRestHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   5280
            MaxLength       =   6
            TabIndex        =   102
            Top             =   2280
            Width           =   615
         End
         Begin VB.TextBox txtPastePartyRestHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   5280
            MaxLength       =   6
            TabIndex        =   101
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtPastePartyRestHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   5280
            MaxLength       =   6
            TabIndex        =   100
            Top             =   1560
            Width           =   615
         End
         Begin VB.CommandButton cmdPasteQ 
            Caption         =   "Regen - Rest"
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
            Left            =   4560
            TabIndex        =   99
            Top             =   840
            Width           =   1395
         End
         Begin VB.TextBox txtPastePartyRestHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   5280
            MaxLength       =   6
            TabIndex        =   98
            Top             =   420
            Width           =   615
         End
         Begin VB.TextBox txtPastePartyRestHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   5280
            MaxLength       =   6
            TabIndex        =   97
            Top             =   1200
            Width           =   615
         End
         Begin VB.CommandButton cmdPasteQ 
            Caption         =   "Attack Last"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   8520
            TabIndex        =   96
            Top             =   700
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyPartyTotal 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   720
            MaxLength       =   6
            TabIndex        =   94
            Top             =   420
            Width           =   555
         End
         Begin VB.TextBox txtPastePartyRegenHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   4620
            MaxLength       =   6
            TabIndex        =   55
            Top             =   3000
            Width           =   555
         End
         Begin VB.OptionButton optPastyPartyAtkLast 
            Caption         =   "Option1"
            Height          =   195
            Index           =   6
            Left            =   8760
            TabIndex        =   57
            Top             =   3060
            Width           =   195
         End
         Begin VB.CheckBox chkPastePartyAM 
            Caption         =   "Check1"
            Height          =   195
            Index           =   6
            Left            =   7920
            TabIndex        =   56
            ToolTipText     =   "Anti-Magic"
            Top             =   3060
            Width           =   195
         End
         Begin VB.TextBox txtPastePartyHitpoints 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   3840
            MaxLength       =   6
            TabIndex        =   54
            Top             =   3000
            Width           =   675
         End
         Begin VB.TextBox txtPastePartyDodge 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   3240
            MaxLength       =   6
            TabIndex        =   53
            Top             =   3000
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyMR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   2640
            MaxLength       =   6
            TabIndex        =   52
            Top             =   3000
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   2040
            MaxLength       =   6
            TabIndex        =   51
            Top             =   3000
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyAC 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   50
            Top             =   3000
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyName 
            BackColor       =   &H80000016&
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   360
            TabIndex        =   80
            Top             =   3000
            Width           =   915
         End
         Begin VB.TextBox txtPastePartyRegenHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   4620
            MaxLength       =   6
            TabIndex        =   47
            Top             =   2640
            Width           =   555
         End
         Begin VB.OptionButton optPastyPartyAtkLast 
            Caption         =   "Option1"
            Height          =   195
            Index           =   5
            Left            =   8760
            TabIndex        =   49
            Top             =   2700
            Width           =   195
         End
         Begin VB.CheckBox chkPastePartyAM 
            Caption         =   "Check1"
            Height          =   195
            Index           =   5
            Left            =   7920
            TabIndex        =   48
            ToolTipText     =   "Anti-Magic"
            Top             =   2700
            Width           =   195
         End
         Begin VB.TextBox txtPastePartyHitpoints 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   3840
            MaxLength       =   6
            TabIndex        =   46
            Top             =   2640
            Width           =   675
         End
         Begin VB.TextBox txtPastePartyDodge 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   3240
            MaxLength       =   6
            TabIndex        =   45
            Top             =   2640
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyMR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   2640
            MaxLength       =   6
            TabIndex        =   44
            Top             =   2640
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   2040
            MaxLength       =   6
            TabIndex        =   43
            Top             =   2640
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyAC 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   42
            Top             =   2640
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyName 
            BackColor       =   &H80000016&
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   360
            TabIndex        =   78
            Top             =   2640
            Width           =   915
         End
         Begin VB.TextBox txtPastePartyRegenHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   4620
            MaxLength       =   6
            TabIndex        =   39
            Top             =   2280
            Width           =   555
         End
         Begin VB.OptionButton optPastyPartyAtkLast 
            Caption         =   "Option1"
            Height          =   195
            Index           =   4
            Left            =   8760
            TabIndex        =   41
            Top             =   2340
            Width           =   195
         End
         Begin VB.CheckBox chkPastePartyAM 
            Caption         =   "Check1"
            Height          =   195
            Index           =   4
            Left            =   7920
            TabIndex        =   40
            ToolTipText     =   "Anti-Magic"
            Top             =   2340
            Width           =   195
         End
         Begin VB.TextBox txtPastePartyHitpoints 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   3840
            MaxLength       =   6
            TabIndex        =   38
            Top             =   2280
            Width           =   675
         End
         Begin VB.TextBox txtPastePartyDodge 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   3240
            MaxLength       =   6
            TabIndex        =   37
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyMR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   2640
            MaxLength       =   6
            TabIndex        =   36
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   2040
            MaxLength       =   6
            TabIndex        =   35
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyAC 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   34
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyName 
            BackColor       =   &H80000016&
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   360
            TabIndex        =   76
            Top             =   2280
            Width           =   915
         End
         Begin VB.TextBox txtPastePartyRegenHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   4620
            MaxLength       =   6
            TabIndex        =   31
            Top             =   1920
            Width           =   555
         End
         Begin VB.OptionButton optPastyPartyAtkLast 
            Caption         =   "Option1"
            Height          =   195
            Index           =   3
            Left            =   8760
            TabIndex        =   33
            Top             =   1980
            Width           =   195
         End
         Begin VB.CheckBox chkPastePartyAM 
            Caption         =   "Check1"
            Height          =   195
            Index           =   3
            Left            =   7920
            TabIndex        =   32
            ToolTipText     =   "Anti-Magic"
            Top             =   1980
            Width           =   195
         End
         Begin VB.TextBox txtPastePartyHitpoints 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   3840
            MaxLength       =   6
            TabIndex        =   30
            Top             =   1920
            Width           =   675
         End
         Begin VB.TextBox txtPastePartyDodge 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   3240
            MaxLength       =   6
            TabIndex        =   29
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyMR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   2640
            MaxLength       =   6
            TabIndex        =   28
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   2040
            MaxLength       =   6
            TabIndex        =   27
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyAC 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   26
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyName 
            BackColor       =   &H80000016&
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   360
            TabIndex        =   74
            Top             =   1920
            Width           =   915
         End
         Begin VB.TextBox txtPastePartyRegenHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   4620
            MaxLength       =   6
            TabIndex        =   23
            Top             =   1560
            Width           =   555
         End
         Begin VB.OptionButton optPastyPartyAtkLast 
            Caption         =   "Option1"
            Height          =   195
            Index           =   2
            Left            =   8760
            TabIndex        =   25
            Top             =   1620
            Width           =   195
         End
         Begin VB.CheckBox chkPastePartyAM 
            Caption         =   "Check1"
            Height          =   195
            Index           =   2
            Left            =   7920
            TabIndex        =   24
            ToolTipText     =   "Anti-Magic"
            Top             =   1620
            Width           =   195
         End
         Begin VB.TextBox txtPastePartyHitpoints 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   3840
            MaxLength       =   6
            TabIndex        =   22
            Top             =   1560
            Width           =   675
         End
         Begin VB.TextBox txtPastePartyDodge 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   3240
            MaxLength       =   6
            TabIndex        =   21
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyMR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   2640
            MaxLength       =   6
            TabIndex        =   20
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   2040
            MaxLength       =   6
            TabIndex        =   19
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyAC 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   18
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyName 
            BackColor       =   &H80000016&
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   360
            TabIndex        =   72
            Top             =   1560
            Width           =   915
         End
         Begin VB.OptionButton optPastyPartyAtkLast 
            Caption         =   "Option1"
            Height          =   195
            Index           =   0
            Left            =   8760
            TabIndex        =   65
            Top             =   480
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.TextBox txtPastePartyAMTotal 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   7740
            MaxLength       =   6
            TabIndex        =   64
            Top             =   420
            Width           =   555
         End
         Begin VB.TextBox txtPastePartyRegenHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   4620
            MaxLength       =   6
            TabIndex        =   63
            Top             =   420
            Width           =   555
         End
         Begin VB.TextBox txtPastePartyHitpoints 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   3840
            MaxLength       =   6
            TabIndex        =   62
            Top             =   420
            Width           =   675
         End
         Begin VB.TextBox txtPastePartyDodge 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   3240
            MaxLength       =   6
            TabIndex        =   61
            Top             =   420
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyMR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   2640
            MaxLength       =   6
            TabIndex        =   60
            Top             =   420
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   2040
            MaxLength       =   6
            TabIndex        =   59
            Top             =   420
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyAC 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   58
            Top             =   420
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyRegenHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   4620
            MaxLength       =   6
            TabIndex        =   14
            Top             =   1200
            Width           =   555
         End
         Begin VB.OptionButton optPastyPartyAtkLast 
            Caption         =   "Option1"
            Height          =   195
            Index           =   1
            Left            =   8760
            TabIndex        =   17
            Top             =   1260
            Width           =   195
         End
         Begin VB.TextBox txtPastePartyName 
            BackColor       =   &H80000016&
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   360
            TabIndex        =   7
            Top             =   1200
            Width           =   915
         End
         Begin VB.TextBox txtPastePartyAC 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   6
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   2040
            MaxLength       =   6
            TabIndex        =   8
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyMR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   2640
            MaxLength       =   6
            TabIndex        =   9
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDodge 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   3240
            MaxLength       =   6
            TabIndex        =   11
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyHitpoints 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   3840
            MaxLength       =   6
            TabIndex        =   13
            Top             =   1200
            Width           =   675
         End
         Begin VB.CheckBox chkPastePartyAM 
            Caption         =   "Check1"
            Height          =   195
            Index           =   1
            Left            =   7920
            TabIndex        =   16
            ToolTipText     =   "Anti-Magic"
            Top             =   1260
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   1  'Right Justify
            Caption         =   "#"
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
            Index           =   28
            Left            =   480
            TabIndex        =   95
            Top             =   480
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   1  'Right Justify
            Caption         =   "6"
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
            Index           =   27
            Left            =   8280
            TabIndex        =   93
            Top             =   3060
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   1  'Right Justify
            Caption         =   "5"
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
            Index           =   26
            Left            =   8280
            TabIndex        =   92
            Top             =   2700
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   1  'Right Justify
            Caption         =   "4"
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
            Index           =   25
            Left            =   8280
            TabIndex        =   91
            Top             =   2340
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   1  'Right Justify
            Caption         =   "3"
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
            Index           =   24
            Left            =   8280
            TabIndex        =   90
            Top             =   1980
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   1  'Right Justify
            Caption         =   "2"
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
            Index           =   23
            Left            =   8280
            TabIndex        =   89
            Top             =   1620
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
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
            Index           =   22
            Left            =   8280
            TabIndex        =   88
            Top             =   1260
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   1  'Right Justify
            Caption         =   "6"
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
            Index           =   21
            Left            =   120
            TabIndex        =   87
            Top             =   3060
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   1  'Right Justify
            Caption         =   "5"
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
            Index           =   20
            Left            =   120
            TabIndex        =   86
            Top             =   2700
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   1  'Right Justify
            Caption         =   "4"
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
            Index           =   19
            Left            =   120
            TabIndex        =   85
            Top             =   2340
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   1  'Right Justify
            Caption         =   "3"
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
            Index           =   18
            Left            =   120
            TabIndex        =   84
            Top             =   1980
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   1  'Right Justify
            Caption         =   "2"
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
            Index           =   17
            Left            =   120
            TabIndex        =   83
            Top             =   1620
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
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
            Index           =   16
            Left            =   120
            TabIndex        =   82
            Top             =   1260
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   2  'Center
            Caption         =   "/"
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
            Index           =   15
            Left            =   1860
            TabIndex        =   81
            Top             =   3060
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   2  'Center
            Caption         =   "/"
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
            TabIndex        =   79
            Top             =   2700
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   2  'Center
            Caption         =   "/"
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
            Left            =   1860
            TabIndex        =   77
            Top             =   2340
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   2  'Center
            Caption         =   "/"
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
            Index           =   12
            Left            =   1860
            TabIndex        =   75
            Top             =   1980
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   2  'Center
            Caption         =   "/"
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
            Index           =   11
            Left            =   1860
            TabIndex        =   73
            Top             =   1620
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   2  'Center
            Caption         =   "none"
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
            Index           =   10
            Left            =   8520
            TabIndex        =   71
            Top             =   240
            Width           =   675
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   2  'Center
            Caption         =   "[=========================  VALUES TO BE SAVED  =========================]"
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
            Index           =   9
            Left            =   600
            TabIndex        =   70
            Top             =   120
            Width           =   7815
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   2  'Center
            Caption         =   "/"
            Enabled         =   0   'False
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
            Index           =   8
            Left            =   1860
            TabIndex        =   69
            Top             =   480
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   2  'Center
            Caption         =   "/"
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
            Index           =   6
            Left            =   1860
            TabIndex        =   68
            Top             =   1260
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   2  'Center
            Caption         =   "AC  /  DR"
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
            Left            =   1500
            TabIndex        =   67
            Top             =   900
            Width           =   975
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   2  'Center
            Caption         =   "MR"
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
            Index           =   1
            Left            =   2700
            TabIndex        =   66
            Top             =   900
            Width           =   375
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   2  'Center
            Caption         =   "Dodge"
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
            Index           =   2
            Left            =   3180
            TabIndex        =   15
            Top             =   900
            Width           =   615
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   2  'Center
            Caption         =   "HP"
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
            Index           =   3
            Left            =   3900
            TabIndex        =   12
            Top             =   900
            Width           =   555
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   2  'Center
            Caption         =   "Anti Magic"
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
            Left            =   7740
            TabIndex        =   10
            Top             =   765
            Width           =   555
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
      Left            =   1560
      TabIndex        =   2
      Top             =   0
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
      Left            =   4020
      TabIndex        =   3
      Top             =   0
      Width           =   1275
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
      Width           =   1215
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
      Height          =   4215
      Left            =   60
      MaxLength       =   10000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   10185
   End
End
Attribute VB_Name = "frmPasteChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bPasteParty As Boolean
Dim bHoldPartyRefresh As Boolean

Private Enum PartyCalc
    all = 0
    ac = 1
    DR = 2
    MR = 3
    Dodge = 4
    HP = 5
    Regen = 6
    Rest = 7
    Heal = 8
    Dmg = 9
    AM = 10
End Enum


Private Sub chkPastePartyAM_Click(Index As Integer)
Call CalculateAverageParty(AM)
End Sub



Private Sub cmdCancel_Click()
On Error Resume Next
Me.bPasteParty = False
txtText.Visible = True
fraPasteParty.Visible = False
cmdPaste.Enabled = True
Me.Hide
End Sub

Private Sub cmdContinue_Click()
On Error GoTo error:
Dim nUpdateHeals As Integer, nHealing As Long

If Me.bPasteParty = True Then
    If fraPasteParty.Visible = False Then
        Call ParsePasteParty
        Exit Sub
    Else
        If (Len(Trim(txtPastePartyRegenHP(0).Text)) > 0 Or Len(Trim(txtPastePartyHeals(0).Text)) > 0) Then
            nUpdateHeals = 1
            nHealing = Round(Val(txtPastePartyRegenHP(0).Text) / 6) + Val(txtPastePartyHeals(0).Text)
            
            If nHealing > 0 And (Len(Trim(txtPastePartyRegenHP(0).Text)) > 0 Or Len(Trim(txtPastePartyHeals(0).Text)) > 0) And _
                (Len(Trim(txtPastePartyRegenHP(0).Text)) = 0 Or Len(Trim(txtPastePartyHeals(0).Text)) = 0) Then
                'one or the other specified, but not both
                
                If Val(frmMain.txtMonsterDamage.Text) > (nHealing * 1.1) Then
                    nUpdateHeals = MsgBox("Either regen rate or healing spells not specified. Update healing/DMG IN field?", vbQuestion + vbDefaultButton3 + vbYesNoCancel)
                    If nUpdateHeals = vbCancel Then Exit Sub
                    If nUpdateHeals = vbYes Then
                        nUpdateHeals = 1
                    Else
                        nUpdateHeals = 0
                    End If
                End If
            End If
            
            If nUpdateHeals = 1 And nHealing > 0 Then
                frmMain.txtMonsterDamage.Text = nHealing
            End If
        End If
        
        If Len(Trim(txtPastePartyAC(0).Text)) > 0 Then frmMain.txtMonsterLairFilter(1).Text = Trim(txtPastePartyAC(0).Text)
        If Len(Trim(txtPastePartyDR(0).Text)) > 0 Then frmMain.txtMonsterLairFilter(2).Text = Trim(txtPastePartyDR(0).Text)
        If Len(Trim(txtPastePartyMR(0).Text)) > 0 Then frmMain.txtMonsterLairFilter(3).Text = Trim(txtPastePartyMR(0).Text)
        If Len(Trim(txtPastePartyDodge(0).Text)) > 0 Then frmMain.txtMonsterLairFilter(4).Text = Trim(txtPastePartyDodge(0).Text)
        If Len(Trim(txtPastePartyHitpoints(0).Text)) > 0 Then frmMain.txtMonsterLairFilter(5).Text = Trim(txtPastePartyHitpoints(0).Text)
        If Len(Trim(txtPastePartyRestHP(0).Text)) > 0 Then frmMain.txtMonsterLairFilter(7).Text = Trim(txtPastePartyRestHP(0).Text)
        If Len(Trim(txtPastePartyAMTotal.Text)) > 0 Then frmMain.txtMonsterLairFilter(6).Text = Trim(txtPastePartyAMTotal.Text)
        If Len(Trim(txtPastePartyPartyTotal.Text)) > 0 Then frmMain.txtMonsterLairFilter(0).Text = Trim(txtPastePartyPartyTotal.Text)
        If Len(Trim(txtPastePartyDMG(0).Text)) > 0 Then frmMain.txtMonsterDamageOUT.Text = Trim(txtPastePartyDMG(0).Text)
        
    End If
End If

out:
On Error Resume Next
Me.Tag = "1"
Me.bPasteParty = False
txtText.Visible = True
fraPasteParty.Visible = False
cmdPaste.Enabled = True
txtText.SetFocus
Me.Hide
Exit Sub
error:
Call HandleError("cmdContinue_Click")
Resume out:
End Sub

Private Sub ParsePasteParty()
On Error GoTo error:
'Dim sSearch As String, sChar As String
', bResult As Boolean, nTries As Integer
'
'Dim nEncum As Long, nStat As String, sName As String
'Dim sCharFile As String, sSectionName As String, nResult As Integer, nYesNo As Integer
Dim tMatches() As RegexMatches, sRegexPattern As String, sSubMatches() As String, sSubValues() As String
Dim sName(6) As String, nMR(6) As Integer, nAC(6) As Integer, nDR(6) As Integer
Dim sRaceName(6) As String, sClassName(6) As String, nClass(6) As Integer, nRace(6) As Integer
Dim nCurrentEnc(6) As Long, nMaxEnc(6) As Long, nHitPoints(6) As Long
Dim nLevel(6) As Integer, nAgility(6) As Integer, nHealth(6) As Integer, nCharm(6) As Integer
Dim x As Integer, x2 As Integer, y As Integer, iMatch As Integer, sPastedText As String
Dim sWorn(1 To 6, 0 To 1) As String, sText As String, iChar As Integer
Dim bItemsFound As Boolean, sEquipLoc(1 To 6, 0 To 19) As String, nItemNum As Long
Dim nPlusRegen(6) As Integer, nPlusDodge(6) As Integer, nTemp As Long

sPastedText = frmPasteChar.txtText.Text
If Len(sPastedText) < 10 Then GoTo canceled:

'Name: Kratos                           Lives/CP:      9/20
'Race: Half-Ogre   Exp: 5281866477      Perception:     75
'Class: Witchunter     Level: 75            Stealth:         0
'Hits:  1402/1447  Armour Class:  82/30 Thievery:        0
'                                       Traps:           0
'                                       Picklocks:       0
'Strength:  170    Agility: 80          Tracking:        0
'Intellect: 80     Health:  170         Martial Arts:   18
'Willpower: 90     Charm:   80          MagicRes:      107
'
'Encumbrance: 5986/10680 - Medium [56%]
'
'Name: Syntax BlackVail                 Lives/CP:      9/4
'Race: Dark-Elf    Exp: 5198070801      Perception:    104
'Class: Ranger     Level: 67            Stealth:        52
'Hits:   815/815   Armour Class:  51/5  Thievery:        0
'Mana:   136/140   Spellcasting: 216    Traps:           0
'                                       Picklocks:       0
'Strength:  110    Agility: 140         Tracking:      256
'Intellect: 110    Health:  110         Martial Arts:   78
'Willpower: 98     Charm:   90          MagicRes:      106
'
'Encumbrance: 2669/5640 - Medium [47%]
'
'Name: Buster Brown                     Lives/CP:      9/2
'Race: Dwarf       Exp: 5096647782      Perception:    111
'Class: Priest     Level: 72            Stealth:         0
'Hits:   714/807   Armour Class:  21/0  Thievery:        0
'Mana: * 312/579   Spellcasting: 290    Traps:           0
'                                       Picklocks:       0
'Strength:  80     Agility: 110         Tracking:        0
'Intellect: 101    Health:  140         Martial Arts:   61
'Willpower: 140    Charm:   105         MagicRes:      143
'
'Encumbrance: 1576/4608 - Medium [34%]

sRegexPattern = "(?:(Armour Class|Hits|Encumbrance):\s*\*?\s*(-?\d+)\/(\d+)|(MagicRes|Level|Agility|Charm|Health):\s*\*?\s*(\d+)|(Name|Race|Class):\s*([^\s:]+(?:\s[^\s:]+)?))"
tMatches() = RegExpFindv2(sPastedText, sRegexPattern, False, True, False)
If UBound(tMatches()) = 0 And Len(tMatches(0).sFullMatch) = 0 Then
    If Val(txtPastePartyPartyTotal.Text) = 0 Then
        MsgBox "No matching data.", vbOKOnly + vbExclamation, "Paste Party"
        GoTo canceled:
    Else
        MsgBox "No data, returning to previous screen.", vbOKOnly + vbInformation, "Paste Party"
        fraPasteParty.Visible = True
        txtText.Visible = False
        GoTo out:
    End If
End If

bHoldPartyRefresh = True

'optPastyPartyAtkLast(0).Value = True
For iChar = 0 To 6
    If iChar > 0 Then txtPastePartyName(iChar).Text = ""
    If iChar > 0 Then chkPastePartyAM(iChar).Value = 0
    txtPastePartyAC(iChar).Text = ""
    txtPastePartyDR(iChar).Text = ""
    txtPastePartyMR(iChar).Text = ""
    txtPastePartyDodge(iChar).Text = ""
    txtPastePartyHitpoints(iChar).Text = ""
    txtPastePartyRestHP(iChar).Text = ""
    txtPastePartyRegenHP(iChar).Text = ""
    'txtPastePartyHeals(iChar).Text = ""
    'txtPastePartyDMG(iChar).Text = ""
Next iChar
txtPastePartyAMTotal.Text = ""
txtPastePartyPartyTotal.Text = ""
fraPasteParty.Visible = True
txtText.Visible = False
'cmdPaste.Enabled = False
DoEvents

For iMatch = 0 To UBound(tMatches())
    If UBound(tMatches(iMatch).sSubMatches()) = 0 Then GoTo skip_match
    
    Select Case tMatches(iMatch).sSubMatches(0)
    
        Case "Name":
            If Val(sName(0)) >= 6 Then GoTo skip_match
            sName(Val(sName(0)) + 1) = Trim(tMatches(iMatch).sSubMatches(1))
            sName(0) = Val(sName(0)) + 1
        Case "Class":
            If Val(sClassName(0)) >= 6 Then GoTo skip_match
            sClassName(Val(sClassName(0)) + 1) = Trim(tMatches(iMatch).sSubMatches(1))
            sClassName(0) = Val(sClassName(0)) + 1
        Case "Race":
            If Val(sRaceName(0)) >= 6 Then GoTo skip_match
            sRaceName(Val(sRaceName(0)) + 1) = Trim(tMatches(iMatch).sSubMatches(1))
            sRaceName(0) = Val(sRaceName(0)) + 1
        
        Case "MagicRes":
            If nMR(0) >= 6 Then GoTo skip_match
            nMR(nMR(0) + 1) = Trim(tMatches(iMatch).sSubMatches(1))
            nMR(0) = nMR(0) + 1
        Case "Level":
            If nLevel(0) >= 6 Then GoTo skip_match
            nLevel(nLevel(0) + 1) = Trim(tMatches(iMatch).sSubMatches(1))
            nLevel(0) = nLevel(0) + 1
        Case "Agility":
            If nAgility(0) >= 6 Then GoTo skip_match
            nAgility(nAgility(0) + 1) = Trim(tMatches(iMatch).sSubMatches(1))
            nAgility(0) = nAgility(0) + 1
        Case "Health":
            If nHealth(0) >= 6 Then GoTo skip_match
            nHealth(nHealth(0) + 1) = Trim(tMatches(iMatch).sSubMatches(1))
            nHealth(0) = nHealth(0) + 1
        Case "Charm":
            If nCharm(0) >= 6 Then GoTo skip_match
            nCharm(nCharm(0) + 1) = Trim(tMatches(iMatch).sSubMatches(1))
            nCharm(0) = nCharm(0) + 1
        
        Case "Hits":
            If nHitPoints(0) >= 6 Then GoTo skip_match
            nHitPoints(nHitPoints(0) + 1) = Trim(tMatches(iMatch).sSubMatches(2))
            nHitPoints(0) = nHitPoints(0) + 1
        
        Case "Armour Class":
            If nAC(0) >= 6 Then GoTo skip_match
            If nDR(0) >= 6 Then GoTo skip_match
            nAC(nAC(0) + 1) = Trim(tMatches(iMatch).sSubMatches(1)): nAC(0) = nAC(0) + 1
            nDR(nDR(0) + 1) = Trim(tMatches(iMatch).sSubMatches(2)): nDR(0) = nDR(0) + 1
        
        Case "Encumbrance":
            If nCurrentEnc(0) >= 6 Then GoTo skip_match
            If nMaxEnc(0) >= 6 Then GoTo skip_match
            nCurrentEnc(nCurrentEnc(0) + 1) = Trim(tMatches(iMatch).sSubMatches(1)): nCurrentEnc(0) = nCurrentEnc(0) + 1
            nMaxEnc(nMaxEnc(0) + 1) = Trim(tMatches(iMatch).sSubMatches(2)): nMaxEnc(0) = nMaxEnc(0) + 1
        
    End Select
    
skip_match:
Next iMatch


'adapted from frmmain.pastecharacter
x = 1
y = 1
x2 = -1
iChar = 0
Do Until x + y > Len(sPastedText) + 1
    
    sChar = Mid(sPastedText, x + y - 1, 1)
    
    bResult = TestPasteChar(sChar)
    If bResult = False Then GoTo next_y:
    
    sText = RemoveCharacter(sText & sChar, " ")
    
    If InStr(1, LCase(sText), "equippedwith:") > 0 Then
        iChar = iChar + 1
        GoTo clear:
    ElseIf InStr(1, LCase(sText), "arecarrying") > 0 Then
        iChar = iChar + 1
        GoTo clear:
    End If
    
    Select Case sChar
        Case ",":
            GoTo clear:
        Case "(":
            x2 = Len(sText)
        Case ")":
            If x2 = -1 Then GoTo clear:
            
            If iChar >= 1 And iChar <= 6 Then
                Select Case UCase(Mid(sText, x2 + 1, Len(sText) - x2 - 1))
                    Case "HEAD": sEquipLoc(iChar, 0) = Left(sText, x2 - 1)
                    Case "EARS": sEquipLoc(iChar, 1) = Left(sText, x2 - 1)
                    Case "EYES": sEquipLoc(iChar, 17) = Left(sText, x2 - 1)
                    Case "FACE": sEquipLoc(iChar, 18) = Left(sText, x2 - 1)
                    Case "NECK": sEquipLoc(iChar, 2) = Left(sText, x2 - 1)
                    Case "BACK": sEquipLoc(iChar, 3) = Left(sText, x2 - 1)
                    Case "TORSO": sEquipLoc(iChar, 4) = Left(sText, x2 - 1)
                    Case "ARMS": sEquipLoc(iChar, 5) = Left(sText, x2 - 1)
                    Case "WRIST":
                        If Not sEquipLoc(iChar, 6) = "" Then
                            If sEquipLoc(iChar, 7) = "" Then
                                sEquipLoc(iChar, 7) = Left(sText, x2 - 1)
                            End If
                        Else
                            sEquipLoc(iChar, 6) = Left(sText, x2 - 1)
                        End If
                    Case "WAIST": sEquipLoc(iChar, 11) = Left(sText, x2 - 1)
                    Case "FINGER":
                        If Not sEquipLoc(iChar, 9) = "" Then
                            If sEquipLoc(iChar, 10) = "" Then
                                sEquipLoc(iChar, 10) = Left(sText, x2 - 1)
                            End If
                        Else
                            sEquipLoc(iChar, 9) = Left(sText, x2 - 1)
                        End If
                    Case "HANDS": sEquipLoc(iChar, 8) = Left(sText, x2 - 1)
                    Case "LEGS": sEquipLoc(iChar, 12) = Left(sText, x2 - 1)
                    Case "FEET": sEquipLoc(iChar, 13) = Left(sText, x2 - 1)
                    Case "WORN":
                        If Not sWorn(iChar, 0) = "" Then
                            If sWorn(iChar, 1) = "" Then
                                sWorn(iChar, 1) = Left(sText, x2 - 1)
                            End If
                        Else
                            sWorn(iChar, 0) = Left(sText, x2 - 1)
                        End If
                        
                    Case "OFF-HAND": sEquipLoc(iChar, 15) = Left(sText, x2 - 1)
                    Case "WEAPONHAND": sEquipLoc(iChar, 16) = Left(sText, x2 - 1)
                    Case "TWOHANDED": sEquipLoc(iChar, 16) = Left(sText, x2 - 1)
                End Select
            End If
            
            GoTo clear:
    End Select

GoTo next_y:

clear:
sText = ""
x = x + y
y = 0
x2 = -1

next_y:
    y = y + 1
Loop

For iChar = 1 To 6
    If sWorn(iChar, 0) <> "" Or sWorn(iChar, 1) <> "" Then bItemsFound = True
    If Not bItemsFound Then
        For x = 0 To UBound(sEquipLoc(), 2)
            If sEquipLoc(iChar, x) <> "" Then
                bItemsFound = True
                Exit For
            End If
        Next x
    End If
Next iChar

If bItemsFound = False Then GoTo calcit:

tabItems.MoveFirst
DoEvents
Do Until tabItems.EOF
        
    If bOnlyInGame And tabItems.Fields("In Game") = 0 Then GoTo skip:
    
    sText = RemoveCharacter(tabItems.Fields("Name"), " ")
    If Len(Trim(sText)) = 0 Then GoTo skip:
    
    For iChar = 1 To 6
        For x = 0 To UBound(sEquipLoc(), 2)
            
            If (x = 14 Or x = 19) And (sText = sWorn(iChar, 0) Or sText = sWorn(iChar, 1)) Then
                If tabItems.Fields("Worn") = 1 Then
                    sEquipLoc(iChar, 19) = sText
                ElseIf tabItems.Fields("Worn") = 16 Then
                    sEquipLoc(iChar, 14) = sText
                End If
            End If
            
            If sText = sEquipLoc(iChar, x) Then
                If x = 7 And Not bInvenUse2ndWrist Then GoTo skip:
                
                For y = 0 To 19
                    If tabItems.Fields("Abil-" & y) > 0 And tabItems.Fields("AbilVal-" & y) <> 0 Then
                        Select Case tabItems.Fields("Abil-" & y)
                            Case 34: 'dodge
                                nPlusDodge(iChar) = nPlusDodge(iChar) + tabItems.Fields("AbilVal-" & y)
                            Case 123: 'hpregen
                                nPlusRegen(iChar) = nPlusRegen(iChar) + tabItems.Fields("AbilVal-" & y)
                        End Select
                    End If
                Next y
            End If

        Next x
    Next iChar
skip:
    tabItems.MoveNext
Loop
tabItems.MoveFirst

calcit:
For iChar = 1 To 6
    
    If Val(sName(0)) >= iChar Then txtPastePartyName(iChar).Text = sName(iChar)
    If nAC(0) >= iChar Then txtPastePartyAC(iChar).Text = nAC(iChar)
    If nDR(0) >= iChar Then txtPastePartyDR(iChar).Text = nDR(iChar)
    If nMR(0) >= iChar Then txtPastePartyMR(iChar).Text = nMR(iChar)
    If nHitPoints(0) >= iChar Then txtPastePartyHitpoints(iChar).Text = nHitPoints(iChar)
    
    If Val(sClassName(0)) >= iChar Then
        If frmMain.cmbGlobalClass(0).ListCount > 0 Then
            For y = 0 To frmMain.cmbGlobalClass(0).ListCount - 1
                If frmMain.cmbGlobalClass(0).List(y) = sClassName(iChar) Then
                    nClass(iChar) = frmMain.cmbGlobalClass(0).ItemData(y)
                End If
            Next
        End If
    End If
    
    If Val(sRaceName(0)) >= iChar Then
        If frmMain.cmbGlobalRace(0).ListCount > 0 Then
            For y = 0 To frmMain.cmbGlobalRace(0).ListCount - 1
                If frmMain.cmbGlobalRace(0).List(y) = sRaceName(iChar) Then
                    nRace(iChar) = frmMain.cmbGlobalRace(0).ItemData(y)
                End If
            Next
        End If
    End If
    
    If nClass(iChar) > 0 Then
        If ClassHasAbility(nClass(iChar), 51) >= 0 Then chkPastePartyAM(iChar).Value = 1
        nTemp = ClassHasAbility(nClass(iChar), 34) 'dodge
        If nTemp <> -31337 Then nPlusDodge(iChar) = nPlusDodge(iChar) + nTemp
        nTemp = ClassHasAbility(nClass(iChar), 123) 'hpregen
        If nTemp <> -31337 Then nPlusRegen(iChar) = nPlusRegen(iChar) + nTemp
    End If
    
    If nRace(iChar) > 0 Then
        If RaceHasAbility(nRace(iChar), 51) >= 0 Then chkPastePartyAM(iChar).Value = 1
        nTemp = RaceHasAbility(nRace(iChar), 34) 'dodge
        If nTemp <> -31337 Then nPlusDodge(iChar) = nPlusDodge(iChar) + nTemp
        nTemp = RaceHasAbility(nRace(iChar), 123) 'hpregen
        If nTemp <> -31337 Then nPlusRegen(iChar) = nPlusRegen(iChar) + nTemp
    End If
    
    If nLevel(iChar) > 0 And nAgility(iChar) > 0 And nCharm(iChar) > 0 Then
        txtPastePartyDodge(iChar).Text = CalcDodge(nLevel(iChar), nAgility(iChar), nCharm(iChar), nPlusDodge(iChar), nCurrentEnc(iChar), nMaxEnc(iChar))
    End If
    
    If nLevel(iChar) > 0 And nHealth(iChar) > 0 Then
        txtPastePartyRegenHP(iChar).Text = CalcRestingRate(nLevel(iChar), nHealth(iChar), nPlusRegen(iChar), False)
        txtPastePartyRestHP(iChar).Text = CalcRestingRate(nLevel(iChar), nHealth(iChar), nPlusRegen(iChar), True)
    End If
Next iChar

out:
On Error Resume Next
tabItems.MoveFirst
bHoldPartyRefresh = False
Call CalculateAverageParty
Exit Sub
canceled:
On Error Resume Next
Call cmdCancel_Click
Exit Sub
error:
Call HandleError("ParsePasteParty")
Resume out:
End Sub

Private Sub CalculateAverageParty(Optional ByVal nWhat As PartyCalc = 0)
On Error GoTo error:
Dim x As Integer, nCount As Integer, nTotal As Long, nPartySize As Integer, bAtkLast As Boolean

If bHoldPartyRefresh Then Exit Sub

Select Case nWhat
    '(0-all and 1-ac will happen anyway)
    Case 2: 'dr
        GoTo dr_only:
    Case 3: 'mr
        GoTo mr_only:
    Case 4: 'dodge
        GoTo dodge_only:
    Case 5: 'hp
        GoTo hp_only:
    Case 6: 'regen
        GoTo regen_only:
    Case 7: 'rest
        GoTo rest_only:
    Case 8: 'heal
        GoTo heal_only:
    Case 9: 'dmg
        GoTo dmg_only:
    Case 10: 'anti-magic
        GoTo am_only:
End Select

nTotal = 0
nCount = 0
For x = 1 To 6
    If Len(Trim(txtPastePartyAC(x).Text)) > 0 Then
        If optPastyPartyAtkLast(x).Value Then bAtkLast = True
        nTotal = nTotal + Val(txtPastePartyAC(x).Text) + IIf(optPastyPartyAtkLast(x).Value, Val(txtPastePartyAC(x).Text), 0)
        nCount = nCount + 1
    End If
Next x
If nCount > 0 Then txtPastePartyAC(0).Text = Round(nTotal / (nCount + IIf(bAtkLast, 1, 0)))
If nCount > nPartySize Then nPartySize = nCount

If nWhat > 0 Then GoTo out
dr_only:

nTotal = 0
nCount = 0
For x = 1 To 6
    If Len(Trim(txtPastePartyDR(x).Text)) > 0 Then
        If optPastyPartyAtkLast(x).Value Then bAtkLast = True
        nTotal = nTotal + Val(txtPastePartyDR(x).Text) + IIf(optPastyPartyAtkLast(x).Value, Val(txtPastePartyDR(x).Text), 0)
        nCount = nCount + 1
    End If
Next x
If nCount > 0 Then txtPastePartyDR(0).Text = Round(nTotal / (nCount + IIf(bAtkLast, 1, 0)))
If nCount > nPartySize Then nPartySize = nCount

If nWhat > 0 Then GoTo out
mr_only:

nTotal = 0
nCount = 0
For x = 1 To 6
    If Len(Trim(txtPastePartyMR(x).Text)) > 0 Then
        If optPastyPartyAtkLast(x).Value Then bAtkLast = True
        nTotal = nTotal + Val(txtPastePartyMR(x).Text) + IIf(optPastyPartyAtkLast(x).Value, Val(txtPastePartyMR(x).Text), 0)
        nCount = nCount + 1
    End If
Next x
If nCount > 0 Then txtPastePartyMR(0).Text = Round(nTotal / (nCount + IIf(bAtkLast, 1, 0)))
If nCount > nPartySize Then nPartySize = nCount

If nWhat > 0 Then GoTo out
dodge_only:

nTotal = 0
nCount = 0
For x = 1 To 6
    If Len(Trim(txtPastePartyDodge(x).Text)) > 0 Then
        If optPastyPartyAtkLast(x).Value Then bAtkLast = True
        nTotal = nTotal + Val(txtPastePartyDodge(x).Text) + IIf(optPastyPartyAtkLast(x).Value, Val(txtPastePartyDodge(x).Text), 0)
        nCount = nCount + 1
    End If
Next x
If nCount > 0 Then txtPastePartyDodge(0).Text = Round(nTotal / (nCount + IIf(bAtkLast, 1, 0)))
If nCount > nPartySize Then nPartySize = nCount

If nWhat > 0 Then GoTo out
hp_only:

nTotal = 0
nCount = 0
For x = 1 To 6
    If Len(Trim(txtPastePartyHitpoints(x).Text)) > 0 Then
        nTotal = nTotal + Val(txtPastePartyHitpoints(x).Text)
        nCount = nCount + 1
    End If
Next x
If nCount > 0 Then txtPastePartyHitpoints(0).Text = Round(nTotal / nCount)
If nCount > nPartySize Then nPartySize = nCount

If nWhat > 0 Then GoTo out
regen_only:

nTotal = 0
nCount = 0
For x = 1 To 6
    If Len(Trim(txtPastePartyRegenHP(x).Text)) > 0 Then
        nTotal = nTotal + Val(txtPastePartyRegenHP(x).Text) + IIf(optPastyPartyAtkLast(x).Value, Val(txtPastePartyRegenHP(x).Text), 0)
        nCount = nCount + 1
    End If
Next x
If nCount > 0 Then txtPastePartyRegenHP(0).Text = Round(nTotal / (nCount + IIf(bAtkLast, 1, 0)))
If nCount > nPartySize Then nPartySize = nCount

'If Trim(txtPastePartyHeals(0).Text) = "" And Len(Trim(txtPastePartyRegenHP(0).Text)) > 0 Then
'    txtPastePartyHeals(0).Text = Val(frmMain.txtMonsterDamage.Text) - Round(Val(txtPastePartyRegenHP(0).Text) / 6)
'End If

If nWhat > 0 Then GoTo out
rest_only:

nTotal = 0
nCount = 0
For x = 1 To 6
    If Len(Trim(txtPastePartyRestHP(x).Text)) > 0 Then
        nTotal = nTotal + Val(txtPastePartyRestHP(x).Text)
        nCount = nCount + 1
    End If
Next x
If nCount > 0 Then txtPastePartyRestHP(0).Text = Round(nTotal / nCount)
If nCount > nPartySize Then nPartySize = nCount

If nWhat > 0 Then GoTo out
heal_only:

nTotal = 0
For x = 1 To 6
    If Len(Trim(txtPastePartyHeals(x).Text)) > 0 Then
        nTotal = nTotal + Val(txtPastePartyHeals(x).Text)
    End If
Next x
If nTotal > 0 Then txtPastePartyHeals(0).Text = nTotal

If nWhat > 0 Then GoTo out
dmg_only:

nTotal = 0
nCount = 0
For x = 1 To 6
    If Len(Trim(txtPastePartyDMG(x).Text)) > 0 Then
        nTotal = nTotal + Val(txtPastePartyDMG(x).Text)
        nCount = nCount + 1
    End If
Next x
If nCount > 0 Then txtPastePartyDMG(0).Text = Round(nTotal / nCount)

If nWhat > 0 Then GoTo out
am_only:

nCount = 0
For x = 1 To 6
    If chkPastePartyAM(x).Value = 1 Then nCount = nCount + 1
Next x
txtPastePartyAMTotal.Text = nCount

If nWhat > 0 Then GoTo out

txtPastePartyPartyTotal.Text = nPartySize

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("CalculateAverageParty")
Resume out:
End Sub

Private Sub cmdPaste_Click()
Dim nYesNo As Integer

If fraPasteParty.Visible Then
    nYesNo = MsgBox("Are you sure?", vbYesNo + vbDefaultButton2 + vbQuestion, "Sure?")
    If Not nYesNo = vbYes Then Exit Sub
End If

If Not Clipboard.GetText = "" Then
    If Not txtText.Text = "" And fraPasteParty.Visible = False Then
        nYesNo = MsgBox("Clear paste area first?", vbYesNo + vbDefaultButton1 + vbQuestion, "Clear?")
    Else
        nYesNo = vbYes
    End If
    
    If nYesNo = vbYes Then
        txtText.Text = Clipboard.GetText
    Else
        txtText.Text = Clipboard.GetText & vbCrLf & vbCrLf & txtText.Text
    End If
    
    txtText.Visible = True
    fraPasteParty.Visible = False
End If

End Sub




Private Sub cmdPasteQ_Click(Index As Integer)
cmdContinue.SetFocus
Select Case Index
    Case 0: 'regen / rest
        MsgBox "The left value is the character's calculated non-resting hp regen rate.  " _
            & "This occurs every 30 seconds so we will divide that by 6 to get a per round [5 second] rate and then factor that into the sustainable damage/healing per round." _
            & vbCrLf & vbCrLf & "The right value is the hp regen rate while resting.  This occurs every 20 seconds when resting and is factored into the calculation for time spent resting.", vbInformation
        
    Case 1: 'heals
        MsgBox "Enter each character's sustainable healing (to self or party) per round from spells.  Note that sustainable here means without requiring to rest." _
            & vbCrLf & vbCrLf & "So, for instance, if you could cast MAHE once every 2 rounds without requiring to rest, take your average MAHE and divide it by 2." _
            & vbCrLf & vbCrLf & "[non-resting regen rate / 6] is added to the saved output.", vbInformation
        
    Case 2: 'attack last
        MsgBox "Selecting one character with attack last will cause that character's defenses to be counted an extra time in the average.", vbInformation
        
    Case 3: 'dmg out
        MsgBox "Enter each character's average per-round damage output.", vbInformation
        
    Case 4: 'general help
        MsgBox "Only fields that have a value are considered in the averages. " _
            & "Likewise, only the fields in the saved area with actual values (a 0 is a value, a blank field is not) will be written back to MME when you click continue. " _
            & "This would allow you to only paste/populate some stats without altering others." _
            & vbCrLf & vbCrLf & "Be sure to click the header buttons for additional information.", vbInformation
        
End Select
End Sub

Private Sub Form_Load()
On Error GoTo error:

cmdPasteQ(2).Caption = "Attack" & vbCrLf & "Last"
lblLabelArray(4).Caption = "Anti" & vbCrLf & "Magic"

With EL1
    .CenterOnLoad = True
    .FormInQuestion = Me
    .MinWidth = 700
    .MinHeight = 350
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

Private Sub optPastyPartyAtkLast_Click(Index As Integer)
Call CalculateAverageParty
End Sub

Private Sub txtPastePartyAC_Change(Index As Integer)
Call CalculateAverageParty(ac)
End Sub

Private Sub txtPastePartyAC_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyAC(Index))
End Sub

Private Sub txtPastePartyAC_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPastePartyAMTotal_GotFocus()
Call SelectAll(txtPastePartyAMTotal)
End Sub

Private Sub txtPastePartyAMTotal_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPastePartyDMG_Change(Index As Integer)
Call CalculateAverageParty(Dmg)
End Sub

Private Sub txtPastePartyDMG_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyDMG(Index))
End Sub

Private Sub txtPastePartyDMG_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPastePartyDodge_Change(Index As Integer)
Call CalculateAverageParty(Dodge)
End Sub

Private Sub txtPastePartyDodge_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyDodge(Index))
End Sub

Private Sub txtPastePartyDodge_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPastePartyDR_Change(Index As Integer)
Call CalculateAverageParty(DR)
End Sub

Private Sub txtPastePartyDR_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyDR(Index))
End Sub

Private Sub txtPastePartyDR_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPastePartyHeals_Change(Index As Integer)
Call CalculateAverageParty(Heal)
End Sub

Private Sub txtPastePartyHeals_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyHeals(Index))
End Sub

Private Sub txtPastePartyHeals_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPastePartyHitpoints_Change(Index As Integer)
Call CalculateAverageParty(HP)
End Sub

Private Sub txtPastePartyHitpoints_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyHitpoints(Index))
End Sub

Private Sub txtPastePartyHitpoints_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPastePartyMR_Change(Index As Integer)
Call CalculateAverageParty(MR)
End Sub

Private Sub txtPastePartyMR_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyMR(Index))
End Sub

Private Sub txtPastePartyMR_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPastePartyName_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyName(Index))
End Sub

Private Sub txtPastePartyPartyTotal_GotFocus()
Call SelectAll(txtPastePartyPartyTotal)
End Sub

Private Sub txtPastePartyRegenHP_Change(Index As Integer)
Call CalculateAverageParty(Regen)
End Sub

Private Sub txtPastePartyRegenHP_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyRegenHP(Index))
End Sub

Private Sub txtPastePartyRegenHP_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPastePartyRestHP_Change(Index As Integer)
Call CalculateAverageParty(Rest)
End Sub

Private Sub txtPastePartyRestHP_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyRestHP(Index))
End Sub

Private Sub txtPastePartyRestHP_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

