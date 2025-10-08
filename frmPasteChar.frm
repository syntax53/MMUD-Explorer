VERSION 5.00
Begin VB.Form frmPasteChar 
   Caption         =   "Paste Characters/Equipment/Spells"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13260
   ControlBox      =   0   'False
   Icon            =   "frmPasteChar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   13260
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timWindowMove 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5580
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
      Height          =   4095
      Left            =   60
      MaxLength       =   32000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   13185
   End
   Begin VB.Frame fraPasteParty 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3855
      Left            =   60
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   12915
      Begin VB.Frame fraMISC 
         BorderStyle     =   0  'None
         Caption         =   "With the updates I have coming out in the  "
         Height          =   3495
         Left            =   240
         TabIndex        =   5
         Top             =   180
         Width           =   12435
         Begin VB.CommandButton cmdPasteQ 
            Caption         =   "Swings"
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
            Left            =   8160
            TabIndex        =   157
            Top             =   840
            Width           =   795
         End
         Begin VB.TextBox txtPastePartySwings 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   8220
            MaxLength       =   3
            TabIndex        =   156
            Top             =   420
            Width           =   675
         End
         Begin VB.TextBox txtPastePartySwings 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   8220
            MaxLength       =   3
            TabIndex        =   155
            Top             =   1200
            Width           =   675
         End
         Begin VB.TextBox txtPastePartySwings 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   8220
            MaxLength       =   3
            TabIndex        =   154
            Top             =   1560
            Width           =   675
         End
         Begin VB.TextBox txtPastePartySwings 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   8220
            MaxLength       =   3
            TabIndex        =   153
            Top             =   1920
            Width           =   675
         End
         Begin VB.TextBox txtPastePartySwings 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   8220
            MaxLength       =   3
            TabIndex        =   152
            Top             =   2280
            Width           =   675
         End
         Begin VB.TextBox txtPastePartySwings 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   8220
            MaxLength       =   3
            TabIndex        =   151
            Top             =   2640
            Width           =   675
         End
         Begin VB.TextBox txtPastePartySwings 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   8220
            MaxLength       =   3
            TabIndex        =   150
            Top             =   3000
            Width           =   675
         End
         Begin VB.CommandButton cmdPasteMegaDmg 
            Height          =   255
            Index           =   12
            Left            =   10470
            TabIndex        =   149
            Top             =   3000
            Width           =   135
         End
         Begin VB.CommandButton cmdPasteMegaDmg 
            Height          =   255
            Index           =   11
            Left            =   10470
            TabIndex        =   148
            Top             =   2640
            Width           =   135
         End
         Begin VB.CommandButton cmdPasteMegaDmg 
            Height          =   255
            Index           =   10
            Left            =   10470
            TabIndex        =   147
            Top             =   2280
            Width           =   135
         End
         Begin VB.CommandButton cmdPasteMegaDmg 
            Height          =   255
            Index           =   9
            Left            =   10470
            TabIndex        =   146
            Top             =   1920
            Width           =   135
         End
         Begin VB.CommandButton cmdPasteMegaDmg 
            Height          =   255
            Index           =   8
            Left            =   10470
            TabIndex        =   145
            Top             =   1560
            Width           =   135
         End
         Begin VB.CommandButton cmdPasteMegaDmg 
            Height          =   255
            Index           =   7
            Left            =   10470
            TabIndex        =   144
            Top             =   1200
            Width           =   135
         End
         Begin VB.CommandButton cmdPasteQ 
            Caption         =   "M.DMG"
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
            Left            =   9720
            TabIndex        =   100
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtPastePartySpellDMG 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   9720
            MaxLength       =   6
            TabIndex        =   96
            Top             =   3000
            Width           =   735
         End
         Begin VB.TextBox txtPastePartySpellDMG 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   9720
            MaxLength       =   6
            TabIndex        =   83
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox txtPastePartySpellDMG 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   9720
            MaxLength       =   6
            TabIndex        =   70
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txtPastePartySpellDMG 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   9720
            MaxLength       =   6
            TabIndex        =   57
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox txtPastePartySpellDMG 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   9720
            MaxLength       =   6
            TabIndex        =   44
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtPastePartySpellDMG 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   9720
            MaxLength       =   6
            TabIndex        =   31
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox txtPastePartySpellDMG 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   9720
            MaxLength       =   6
            TabIndex        =   18
            Top             =   420
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyACCY 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   9060
            MaxLength       =   6
            TabIndex        =   95
            Top             =   3000
            Width           =   555
         End
         Begin VB.TextBox txtPastePartyACCY 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   9060
            MaxLength       =   6
            TabIndex        =   82
            Top             =   2640
            Width           =   555
         End
         Begin VB.TextBox txtPastePartyACCY 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   9060
            MaxLength       =   6
            TabIndex        =   69
            Top             =   2280
            Width           =   555
         End
         Begin VB.TextBox txtPastePartyACCY 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   9060
            MaxLength       =   6
            TabIndex        =   56
            Top             =   1920
            Width           =   555
         End
         Begin VB.TextBox txtPastePartyACCY 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   9060
            MaxLength       =   6
            TabIndex        =   43
            Top             =   1560
            Width           =   555
         End
         Begin VB.TextBox txtPastePartyACCY 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   9060
            MaxLength       =   6
            TabIndex        =   30
            Top             =   1200
            Width           =   555
         End
         Begin VB.TextBox txtPastePartyACCY 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   9060
            MaxLength       =   6
            TabIndex        =   17
            Top             =   420
            Width           =   555
         End
         Begin VB.CommandButton cmdPasteMegaDmg 
            Height          =   255
            Index           =   6
            Left            =   7950
            TabIndex        =   142
            Top             =   3000
            Width           =   135
         End
         Begin VB.CommandButton cmdPasteMegaDmg 
            Height          =   255
            Index           =   5
            Left            =   7950
            TabIndex        =   141
            Top             =   2640
            Width           =   135
         End
         Begin VB.CommandButton cmdPasteMegaDmg 
            Height          =   255
            Index           =   4
            Left            =   7950
            TabIndex        =   140
            Top             =   2280
            Width           =   135
         End
         Begin VB.CommandButton cmdPasteMegaDmg 
            Height          =   255
            Index           =   3
            Left            =   7950
            TabIndex        =   139
            Top             =   1920
            Width           =   135
         End
         Begin VB.CommandButton cmdPasteMegaDmg 
            Height          =   255
            Index           =   2
            Left            =   7950
            TabIndex        =   138
            Top             =   1560
            Width           =   135
         End
         Begin VB.CommandButton cmdPasteMegaDmg 
            Height          =   255
            Index           =   1
            Left            =   7950
            TabIndex        =   137
            Top             =   1200
            Width           =   135
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
            Width           =   315
         End
         Begin VB.TextBox txtPastePartyDMG 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   7200
            MaxLength       =   6
            TabIndex        =   94
            Top             =   3000
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyDMG 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   7200
            MaxLength       =   6
            TabIndex        =   81
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyDMG 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   7200
            MaxLength       =   6
            TabIndex        =   68
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyDMG 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   7200
            MaxLength       =   6
            TabIndex        =   55
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyDMG 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   7200
            MaxLength       =   6
            TabIndex        =   42
            Top             =   1560
            Width           =   735
         End
         Begin VB.CommandButton cmdPasteQ 
            Caption         =   "P.DMG"
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
            Left            =   7200
            TabIndex        =   99
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtPastePartyDMG 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   7200
            MaxLength       =   6
            TabIndex        =   16
            Top             =   420
            Width           =   855
         End
         Begin VB.TextBox txtPastePartyDMG 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   7200
            MaxLength       =   6
            TabIndex        =   29
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyHeals 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   6420
            MaxLength       =   6
            TabIndex        =   28
            Top             =   1200
            Width           =   675
         End
         Begin VB.TextBox txtPastePartyHeals 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   6420
            MaxLength       =   6
            TabIndex        =   15
            Top             =   420
            Width           =   675
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
            Left            =   6420
            TabIndex        =   98
            Top             =   840
            Width           =   675
         End
         Begin VB.TextBox txtPastePartyHeals 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   6420
            MaxLength       =   6
            TabIndex        =   41
            Top             =   1560
            Width           =   675
         End
         Begin VB.TextBox txtPastePartyHeals 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   6420
            MaxLength       =   6
            TabIndex        =   54
            Top             =   1920
            Width           =   675
         End
         Begin VB.TextBox txtPastePartyHeals 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   6420
            MaxLength       =   6
            TabIndex        =   67
            Top             =   2280
            Width           =   675
         End
         Begin VB.TextBox txtPastePartyHeals 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   6420
            MaxLength       =   6
            TabIndex        =   80
            Top             =   2640
            Width           =   675
         End
         Begin VB.TextBox txtPastePartyHeals 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   6420
            MaxLength       =   6
            TabIndex        =   93
            Top             =   3000
            Width           =   675
         End
         Begin VB.TextBox txtPastePartyRestHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   5640
            MaxLength       =   6
            TabIndex        =   92
            Top             =   3000
            Width           =   615
         End
         Begin VB.TextBox txtPastePartyRestHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   5640
            MaxLength       =   6
            TabIndex        =   79
            Top             =   2640
            Width           =   615
         End
         Begin VB.TextBox txtPastePartyRestHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   5640
            MaxLength       =   6
            TabIndex        =   66
            Top             =   2280
            Width           =   615
         End
         Begin VB.TextBox txtPastePartyRestHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   5640
            MaxLength       =   6
            TabIndex        =   53
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtPastePartyRestHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   5640
            MaxLength       =   6
            TabIndex        =   40
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
            Left            =   4920
            TabIndex        =   97
            Top             =   840
            Width           =   1395
         End
         Begin VB.TextBox txtPastePartyRestHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   5640
            MaxLength       =   6
            TabIndex        =   14
            Top             =   420
            Width           =   615
         End
         Begin VB.TextBox txtPastePartyRestHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   5640
            MaxLength       =   6
            TabIndex        =   27
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
            Left            =   11460
            TabIndex        =   101
            Top             =   705
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyPartyTotal 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1080
            MaxLength       =   6
            TabIndex        =   7
            Top             =   420
            Width           =   555
         End
         Begin VB.TextBox txtPastePartyRegenHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   4980
            MaxLength       =   6
            TabIndex        =   91
            Top             =   3000
            Width           =   555
         End
         Begin VB.OptionButton optPastyPartyAtkLast 
            DownPicture     =   "frmPasteChar.frx":0CCA
            Height          =   195
            Index           =   6
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   103
            Top             =   3060
            Width           =   435
         End
         Begin VB.CheckBox chkPastePartyAM 
            Caption         =   "Check1"
            Height          =   195
            Index           =   6
            Left            =   10905
            TabIndex        =   102
            ToolTipText     =   "Anti-Magic"
            Top             =   3060
            Width           =   195
         End
         Begin VB.TextBox txtPastePartyHitpoints 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   4200
            MaxLength       =   6
            TabIndex        =   90
            Top             =   3000
            Width           =   675
         End
         Begin VB.TextBox txtPastePartyDodge 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   3600
            MaxLength       =   6
            TabIndex        =   89
            Top             =   3000
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyMR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   3000
            MaxLength       =   6
            TabIndex        =   88
            Top             =   3000
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   2400
            MaxLength       =   6
            TabIndex        =   87
            Top             =   3000
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyAC 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   1740
            MaxLength       =   6
            TabIndex        =   86
            Top             =   3000
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyName 
            BackColor       =   &H80000016&
            Height          =   285
            Index           =   6
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   122
            Top             =   3000
            Width           =   1275
         End
         Begin VB.TextBox txtPastePartyRegenHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   4980
            MaxLength       =   6
            TabIndex        =   78
            Top             =   2640
            Width           =   555
         End
         Begin VB.OptionButton optPastyPartyAtkLast 
            DownPicture     =   "frmPasteChar.frx":47A4
            Height          =   195
            Index           =   5
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   85
            Top             =   2700
            Width           =   435
         End
         Begin VB.CheckBox chkPastePartyAM 
            Caption         =   "Check1"
            Height          =   195
            Index           =   5
            Left            =   10905
            TabIndex        =   84
            ToolTipText     =   "Anti-Magic"
            Top             =   2700
            Width           =   195
         End
         Begin VB.TextBox txtPastePartyHitpoints 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   4200
            MaxLength       =   6
            TabIndex        =   77
            Top             =   2640
            Width           =   675
         End
         Begin VB.TextBox txtPastePartyDodge 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   3600
            MaxLength       =   6
            TabIndex        =   76
            Top             =   2640
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyMR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   3000
            MaxLength       =   6
            TabIndex        =   75
            Top             =   2640
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   2400
            MaxLength       =   6
            TabIndex        =   74
            Top             =   2640
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyAC 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   1740
            MaxLength       =   6
            TabIndex        =   73
            Top             =   2640
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyName 
            BackColor       =   &H80000016&
            Height          =   285
            Index           =   5
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   120
            Top             =   2640
            Width           =   1275
         End
         Begin VB.TextBox txtPastePartyRegenHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   4980
            MaxLength       =   6
            TabIndex        =   65
            Top             =   2280
            Width           =   555
         End
         Begin VB.OptionButton optPastyPartyAtkLast 
            DownPicture     =   "frmPasteChar.frx":827E
            Height          =   195
            Index           =   4
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   72
            Top             =   2340
            Width           =   435
         End
         Begin VB.CheckBox chkPastePartyAM 
            Caption         =   "Check1"
            Height          =   195
            Index           =   4
            Left            =   10905
            TabIndex        =   71
            ToolTipText     =   "Anti-Magic"
            Top             =   2340
            Width           =   195
         End
         Begin VB.TextBox txtPastePartyHitpoints 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   4200
            MaxLength       =   6
            TabIndex        =   64
            Top             =   2280
            Width           =   675
         End
         Begin VB.TextBox txtPastePartyDodge 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   3600
            MaxLength       =   6
            TabIndex        =   63
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyMR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   3000
            MaxLength       =   6
            TabIndex        =   62
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   2400
            MaxLength       =   6
            TabIndex        =   61
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyAC 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   1740
            MaxLength       =   6
            TabIndex        =   60
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyName 
            BackColor       =   &H80000016&
            Height          =   285
            Index           =   4
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   118
            Top             =   2280
            Width           =   1275
         End
         Begin VB.TextBox txtPastePartyRegenHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   4980
            MaxLength       =   6
            TabIndex        =   52
            Top             =   1920
            Width           =   555
         End
         Begin VB.OptionButton optPastyPartyAtkLast 
            DownPicture     =   "frmPasteChar.frx":BD58
            Height          =   195
            Index           =   3
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   1980
            Width           =   435
         End
         Begin VB.CheckBox chkPastePartyAM 
            Caption         =   "Check1"
            Height          =   195
            Index           =   3
            Left            =   10905
            TabIndex        =   58
            ToolTipText     =   "Anti-Magic"
            Top             =   1980
            Width           =   195
         End
         Begin VB.TextBox txtPastePartyHitpoints 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   4200
            MaxLength       =   6
            TabIndex        =   51
            Top             =   1920
            Width           =   675
         End
         Begin VB.TextBox txtPastePartyDodge 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   3600
            MaxLength       =   6
            TabIndex        =   50
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyMR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   3000
            MaxLength       =   6
            TabIndex        =   49
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   2400
            MaxLength       =   6
            TabIndex        =   48
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyAC 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   1740
            MaxLength       =   6
            TabIndex        =   47
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyName 
            BackColor       =   &H80000016&
            Height          =   285
            Index           =   3
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   116
            Top             =   1920
            Width           =   1275
         End
         Begin VB.TextBox txtPastePartyRegenHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   4980
            MaxLength       =   6
            TabIndex        =   39
            Top             =   1560
            Width           =   555
         End
         Begin VB.OptionButton optPastyPartyAtkLast 
            DownPicture     =   "frmPasteChar.frx":F832
            Height          =   195
            Index           =   2
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   1620
            Width           =   435
         End
         Begin VB.CheckBox chkPastePartyAM 
            Caption         =   "Check1"
            Height          =   195
            Index           =   2
            Left            =   10905
            TabIndex        =   45
            ToolTipText     =   "Anti-Magic"
            Top             =   1620
            Width           =   195
         End
         Begin VB.TextBox txtPastePartyHitpoints 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   4200
            MaxLength       =   6
            TabIndex        =   38
            Top             =   1560
            Width           =   675
         End
         Begin VB.TextBox txtPastePartyDodge 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   3600
            MaxLength       =   6
            TabIndex        =   37
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyMR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   3000
            MaxLength       =   6
            TabIndex        =   36
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   2400
            MaxLength       =   6
            TabIndex        =   35
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyAC 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   1740
            MaxLength       =   6
            TabIndex        =   34
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyName 
            BackColor       =   &H80000016&
            Height          =   285
            Index           =   2
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   114
            Top             =   1560
            Width           =   1275
         End
         Begin VB.OptionButton optPastyPartyAtkLast 
            DownPicture     =   "frmPasteChar.frx":1330C
            Height          =   195
            Index           =   0
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   480
            Value           =   -1  'True
            Width           =   435
         End
         Begin VB.TextBox txtPastePartyAMTotal 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   10740
            MaxLength       =   6
            TabIndex        =   19
            Top             =   420
            Width           =   555
         End
         Begin VB.TextBox txtPastePartyRegenHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   4980
            MaxLength       =   6
            TabIndex        =   13
            Top             =   420
            Width           =   555
         End
         Begin VB.TextBox txtPastePartyHitpoints 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   4200
            MaxLength       =   6
            TabIndex        =   12
            Top             =   420
            Width           =   675
         End
         Begin VB.TextBox txtPastePartyDodge 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   3600
            MaxLength       =   6
            TabIndex        =   11
            Top             =   420
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyMR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   3000
            MaxLength       =   6
            TabIndex        =   10
            Top             =   420
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   2400
            MaxLength       =   6
            TabIndex        =   9
            Top             =   420
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyAC 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   1740
            MaxLength       =   6
            TabIndex        =   8
            Top             =   420
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyRegenHP 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   4980
            MaxLength       =   6
            TabIndex        =   26
            Top             =   1200
            Width           =   555
         End
         Begin VB.OptionButton optPastyPartyAtkLast 
            DownPicture     =   "frmPasteChar.frx":16DE6
            Height          =   195
            Index           =   1
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   1260
            Width           =   435
         End
         Begin VB.TextBox txtPastePartyName 
            BackColor       =   &H80000016&
            Height          =   285
            Index           =   1
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   104
            Top             =   1200
            Width           =   1275
         End
         Begin VB.TextBox txtPastePartyAC 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   1740
            MaxLength       =   6
            TabIndex        =   21
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   2400
            MaxLength       =   6
            TabIndex        =   22
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyMR 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   3000
            MaxLength       =   6
            TabIndex        =   23
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDodge 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   3600
            MaxLength       =   6
            TabIndex        =   24
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyHitpoints 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   4200
            MaxLength       =   6
            TabIndex        =   25
            Top             =   1200
            Width           =   675
         End
         Begin VB.CheckBox chkPastePartyAM 
            Caption         =   "Check1"
            Height          =   195
            Index           =   1
            Left            =   10905
            TabIndex        =   32
            ToolTipText     =   "Anti-Magic"
            Top             =   1260
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   2  'Center
            Caption         =   "ACCY"
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
            Index           =   5
            Left            =   9060
            TabIndex        =   143
            Top             =   900
            Width           =   555
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
            Left            =   840
            TabIndex        =   136
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
            Left            =   11220
            TabIndex        =   135
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
            Left            =   11220
            TabIndex        =   134
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
            Left            =   11220
            TabIndex        =   133
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
            Left            =   11220
            TabIndex        =   132
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
            Left            =   11220
            TabIndex        =   131
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
            Left            =   11220
            TabIndex        =   130
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
            TabIndex        =   129
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
            TabIndex        =   128
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
            TabIndex        =   127
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
            TabIndex        =   126
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
            TabIndex        =   125
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
            TabIndex        =   124
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
            Left            =   2220
            TabIndex        =   123
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
            Left            =   2220
            TabIndex        =   121
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
            Left            =   2220
            TabIndex        =   119
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
            Left            =   2220
            TabIndex        =   117
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
            Left            =   2220
            TabIndex        =   115
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
            Left            =   11580
            TabIndex        =   113
            Top             =   240
            Width           =   555
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   2  'Center
            Caption         =   "[=======================================  VALUES TO BE SAVED  ======================================]"
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
            Left            =   900
            TabIndex        =   112
            Top             =   120
            Width           =   10575
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
            Left            =   2220
            TabIndex        =   111
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
            Left            =   2220
            TabIndex        =   110
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
            Left            =   1860
            TabIndex        =   109
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
            Left            =   3060
            TabIndex        =   108
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
            Left            =   3540
            TabIndex        =   107
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
            Left            =   4260
            TabIndex        =   106
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
            Left            =   10740
            TabIndex        =   105
            Top             =   765
            Width           =   555
         End
      End
   End
End
Attribute VB_Name = "frmPasteChar"
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

Public bPasteParty As Boolean
Dim bHoldPartyRefresh As Boolean
Dim tWindowSize As WindowSizeProperties
Private Enum PartyCalc
    All = 0
    ac = 1
    DR = 2
    MR = 3
    Dodge = 4
    HP = 5
    Regen = 6
    Rest = 7
    heal = 8
    dmg = 9
    AM = 10
    accy = 11
    SPdmg = 12
    Swings = 13
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
Dim nUpdateHeals As Integer, nHealing As Long, sSectionName As String, sCharFile As String
Dim sTemp As String

sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ")
sCharFile = ReadINI(sSectionName, "LastCharFile")
If Len(sSessionLastCharFile) > 0 Then sCharFile = sSessionLastCharFile
If Not FileExists(sCharFile) Then
    sCharFile = ""
    sSessionLastCharFile = ""
End If

If Me.bPasteParty = True Then
    If fraPasteParty.Visible = False Then
        Call ParsePasteParty
        Exit Sub
    Else
        If val(txtPastePartyPartyTotal.Text) > 1 Then
            frmMain.txtMonsterLairFilter(0).Text = val(txtPastePartyPartyTotal.Text)
            DoEvents
        
            If (Len(Trim(txtPastePartyRegenHP(0).Text)) > 0 Or Len(Trim(txtPastePartyHeals(0).Text)) > 0) Then
                nUpdateHeals = 1
                If nNMRVer < 1.83 And Len(Trim(txtPastePartyHeals(0).Text)) = 0 Then
                    nHealing = Round(val(txtPastePartyRegenHP(0).Text) / 2) + Round(val(txtPastePartyRestHP(0).Text) / 3)
                Else
                    nHealing = Round(val(txtPastePartyRegenHP(0).Text) / 6) + val(txtPastePartyHeals(0).Text)
                End If
                
                If nHealing > 0 And val(frmMain.txtMonsterDamage.Text) <> 99999 And (Len(Trim(txtPastePartyRegenHP(0).Text)) > 0 Or Len(Trim(txtPastePartyHeals(0).Text)) > 0) And _
                    (Len(Trim(txtPastePartyRegenHP(0).Text)) = 0 Or Len(Trim(txtPastePartyHeals(0).Text)) = 0) And (val(frmMain.txtMonsterDamage.Text) <> nHealing) Then
                    'one or the other specified, but not both
                    
                    'If Val(frmMain.txtMonsterDamage.Text) > (nHealing * 1.1) And Val(frmMain.txtMonsterDamage.Text) <> 99999 Then
                        sTemp = "Current value of " & val(frmMain.txtMonsterDamage.Text) & " would be overwritten to " & nHealing
                        If nNMRVer < 1.83 Then
                            nUpdateHeals = MsgBox("Update the [DMG <=] field? " & vbCrLf & vbCrLf & "Either regen rate or healing spells not specified. These two fields are normally computed together to update the [DMG <=] field." _
                                & vbCrLf & vbCrLf & "This is your sustainable damage IN/healing amount before requiring to rest. " _
                                & "This is an older database and therefore does not scale this field with more damage. Instead, the filter will simply exclude mobs that deal more damage than this. The in-combat resting rate has been adjusted to compensate for this limitation." _
                                & vbCrLf & vbCrLf & sTemp, vbQuestion + vbDefaultButton3 + vbYesNoCancel, "Update the [DMG <=] field?")
                        Else
                            nUpdateHeals = MsgBox("Update the [HEALS] field? " & vbCrLf & vbCrLf & "Either regen rate or healing spells not specified. These two fields are normally computed together to update the [HEALS] field." _
                                & vbCrLf & vbCrLf & "This is your sustainable damage IN/healing amount before requiring to rest. " _
                                & "If you have no additional healing, you probably want to answer yes." _
                                & vbCrLf & vbCrLf & sTemp, vbQuestion + vbDefaultButton3 + vbYesNoCancel, "Update the [HEALS] field?")
                        End If
                        If nUpdateHeals = vbCancel Then Exit Sub
                        If nUpdateHeals = vbYes Then
                            nUpdateHeals = 1
                        Else
                            nUpdateHeals = 0
                        End If
                    'End If
                End If
                
                If nUpdateHeals = 1 And nHealing > 0 Then frmMain.txtMonsterDamage.Text = nHealing
            End If
            
            If Len(Trim(txtPastePartyAC(0).Text)) > 0 Then frmMain.txtMonsterLairFilter(1).Text = Trim(txtPastePartyAC(0).Text)
            If Len(Trim(txtPastePartyDR(0).Text)) > 0 Then frmMain.txtMonsterLairFilter(2).Text = Trim(txtPastePartyDR(0).Text)
            If Len(Trim(txtPastePartyMR(0).Text)) > 0 Then frmMain.txtMonsterLairFilter(3).Text = Trim(txtPastePartyMR(0).Text)
            If Len(Trim(txtPastePartyDodge(0).Text)) > 0 Then frmMain.txtMonsterLairFilter(4).Text = Trim(txtPastePartyDodge(0).Text)
            If Len(Trim(txtPastePartyHitpoints(0).Text)) > 0 Then frmMain.txtMonsterLairFilter(5).Text = Trim(txtPastePartyHitpoints(0).Text)
            If Len(Trim(txtPastePartyAMTotal.Text)) > 0 Then frmMain.txtMonsterLairFilter(6).Text = Trim(txtPastePartyAMTotal.Text)
            If Len(Trim(txtPastePartyRestHP(0).Text)) > 0 Then frmMain.txtMonsterLairFilter(7).Text = Trim(txtPastePartyRestHP(0).Text)
            If Len(Trim(txtPastePartyACCY(0).Text)) > 0 Then frmMain.txtMonsterLairFilter(8).Text = Trim(txtPastePartyACCY(0).Text)
            If Len(Trim(txtPastePartySwings(0).Text)) > 0 Then frmMain.txtMonsterLairFilter(9).Text = Trim(txtPastePartySwings(0).Text)
            
            If Len(Trim(txtPastePartyDMG(0).Text)) > 0 Then
                frmMain.txtMonsterDamageOUT(0).Text = Trim(txtPastePartyDMG(0).Text)
                If Len(Trim(txtPastePartySpellDMG(0).Text)) = 0 And val(frmMain.txtMonsterDamageOUT(1).Text) >= 9999 Then frmMain.txtMonsterDamageOUT(1).Text = 0
            End If
            If Len(Trim(txtPastePartySpellDMG(0).Text)) > 0 Then
                frmMain.txtMonsterDamageOUT(1).Text = Trim(txtPastePartySpellDMG(0).Text)
                If Len(Trim(txtPastePartyDMG(0).Text)) = 0 And val(frmMain.txtMonsterDamageOUT(0).Text) >= 9999 Then frmMain.txtMonsterDamageOUT(0).Text = 0
            End If
        Else
            MsgBox "Data only updated when party size > 1.", vbInformation
        End If
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
Dim tMatches() As RegexMatches, sRegexPattern As String, nSpellcasting(6) As Integer, nMaxMana(6) As Integer
Dim sName(6) As String, nMR(6) As Integer, nAC(6) As Integer, nDR(6) As Integer
Dim sRaceName(6) As String, sClassName(6) As String, nClass(6) As Integer, nRace(6) As Integer
Dim nCurrentEnc(6) As Long, nMaxEnc(6) As Long, nHitPoints(6) As Long, bResult As Boolean
Dim nLevel(6) As Integer, nAgility(6) As Integer, nHealth(6) As Integer, nCharm(6) As Integer
Dim nStrength(6) As Integer, nIntellect(6) As Integer, nAccyWorn(6) As Double, nAccyAbil(6) As Integer
Dim x As Integer, x2 As Integer, y As Integer, iMatch As Integer, sPastedText As String
Dim sWorn(1 To 6, 0 To 1) As String, sText As String, iChar As Integer, sChar As String
Dim bItemsFound As Boolean, sEquipLoc(1 To 6, 0 To 19) As String, nWillpower(6) As Integer
Dim nPlusRegen(6) As Integer, nPlusDodge(6) As Integer, nTemp As Long, sFindAtkLast As String
Dim nPhysDamage(6) As Long, nSwings(6) As Double, tAttack As tAttackDamage, nSpellDamage(6) As Long
Dim tCharacter(6) As tCharacterProfile, nEnergy As Integer, nAttackTypeMUD As eAttackTypeMUD, nCombat As Integer
Dim nWeaponNum(6) As Long, nWeaponSpeed(6) As Long, nWeaponSTR(6) As Long
Dim nAttackSpellNum As Long, sSpellAttackShort As String, tSpellcast As tSpellCastValues

sPastedText = frmPasteChar.txtText.Text
If Len(sPastedText) < 10 Then GoTo canceled:
'
'Name: Kratos                           Lives/CP:      9/495
'Race: Half-Ogre   Exp: 13072715761     Perception:     75
'Class: Warrior    Level: 85            Stealth:         0
'Hits:  1563/1614  Armour Class:  84/30 Thievery:        0
'                                       Traps:           0
'                                       Picklocks:       0
'Strength:  170    Agility: 80          Tracking:        0
'Intellect: 80     Health:  170         Martial Arts:   21
'Willpower: 90     Charm:   80          MagicRes:      107
'You are carrying 28 gold crowns, starsteel plate gauntlets (Hands), golden
'belt (Waist), visored greathelm (Head), plate boots (Feet), carved ivory mask
'(Ears), platinum ring (Finger), green fighter's eyeglasses (Eyes), cat's-eye
'pendant (Neck), crimson cloak (Back), starsteel plate leggings (Legs), crimson
'bracers (Arms), phoenix feather (Worn), petrified stone corselet (Torso),
'nexus spear (Weapon Hand), griffon shield, king crab claw, white gold ring,
'severed earth dragon claw, rope and grapple, mine pass, 3 waterskin, climbing
'harness, lionskin belt, stormmetal greataxe
'You have the following keys:  helmet key, magma key, 2 dragon keys, ancient
'obsidian key, 2 golden idols, skeleton key, steel key, 2 black star keys,
'green metal key, large iron key, glass key, 2 black serpent keys.
'Wealth: 2800 copper farthings
'Encumbrance: 6473/10680 - Medium [60%]
'
'Name: Buster Brown                     Lives/CP:      9/190
'Race: Dwarf       Exp: 13275472086     Perception:    116
'Class: Priest     Level: 81            Stealth:         0
'Hits:   892/892   Armour Class:  50/3  Thievery:        0
'Mana: * 633/633   Spellcasting: 310    Traps:           0
'                                       Picklocks:       0
'Strength:  110    Agility: 110         Tracking:        0
'Intellect: 110    Health:  140         Martial Arts:   70
'Willpower: 140    Charm:   105         MagicRes:      145
'You feel protected!
'You are carrying 1 platinum piece, 5 gold crowns, white gold ring (Finger),
'cloth pants (Legs), white satin gloves (Hands), silver bracers (Arms), silver
'hood (Head), crimson cloak (Back), amethyst pendant (Neck), skull mask (Ears),
'shimmering white robes (Torso), severed head of Goru-Nezar (Worn), astral
'slippers (Feet), moonstone ring (Finger), golden chalice (Off-Hand), lionskin
'belt (Waist), amber sceptre (Weapon Hand), silverbark canoe, mine pass, golden
'chalice, rope and grapple, 5 waterskin, magma amulet, climbing harness, torch,
'phoenix feather
'You have the following keys:  sapphire key, 2 iron keys, basalt key, black
'serpent key, 2 black star keys.
'Wealth: 10500 copper farthings
'Encumbrance: 1567/6768 - Light [23%]
'
'Name: Happy Gilmore                    Lives/CP:      9/1
'Race: Kang        Exp: 31465429        Perception:     73
'Class: Druid      Level: 20            Stealth:         0
'Hits:   227/227   Armour Class:  46/5  Thievery:        0
'Mana: *  12/131   Spellcasting: 140    Traps:           0
'                                       Picklocks:       0
'Strength:  60     Agility: 61          Tracking:        0
'Intellect: 70     Health:  80          Martial Arts:   34
'Willpower: 80     Charm:   50          MagicRes:       92
'You are carrying silvery skullcap (Head), carved ivory mask (Ears), Crest of
'Silvermere (Neck), crimson cloak (Back), ogre-skin shirt (Torso), platinum
'bracers (Arms), fingerbone bracelet (Wrist), silver bracelet (Wrist), hard
'leather gauntlets (Hands), platinum ring (Finger), white gold ring (Finger),
'golden belt (Waist), rigid leather pants (Legs), hardened leather boots
'(Feet), darkwood staff (Weapon Hand)
'You have the following keys:  bone key, black star key.
'Wealth: 0 copper farthings
'Encumbrance: 1535/2880 - Medium [53%]

sRegexPattern = "(?:(Armour Class|Hits|Mana|Encumbrance):\s*\*?\s*(-?\d+)\/(\d+)|(MagicRes|Level|Spellcasting|Strength|Agility|Willpower|Charm|Intellect|Health):\s*\*?\s*(\d+)|(Name|Race|Class):\s*([^\s:]+(?:\s[^\s:]+)?))"
tMatches() = RegExpFindv2(sPastedText, sRegexPattern, False, True, False)
If UBound(tMatches()) = 0 And Len(tMatches(0).sFullMatch) = 0 Then
    If val(txtPastePartyPartyTotal.Text) = 0 Then
        x = MsgBox("No matching data pasted. Continue to party screen anyway?", vbYesNo + vbDefaultButton2 + vbQuestion, "Paste Party")
        If x <> vbYes Then Exit Sub
    Else
        MsgBox "No data, returning to previous screen.", vbOKOnly + vbInformation, "Paste Party"
    End If
    fraPasteParty.Visible = True
    txtText.Visible = False
    GoTo out:
End If

bHoldPartyRefresh = True

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
    txtPastePartyACCY(iChar).Text = ""
Next iChar

txtPastePartyAMTotal.Text = ""
txtPastePartyPartyTotal.Text = ""
DoEvents

For iMatch = 0 To UBound(tMatches())
    If UBound(tMatches(iMatch).sSubMatches()) = 0 Then GoTo skip_match
    
    Select Case tMatches(iMatch).sSubMatches(0)
    
        Case "Name":
            If val(sName(0)) >= 6 Then GoTo skip_match
            sName(val(sName(0)) + 1) = Trim(tMatches(iMatch).sSubMatches(1))
            sName(0) = val(sName(0)) + 1
        Case "Class":
            If val(sClassName(0)) >= 6 Then GoTo skip_match
            sClassName(val(sClassName(0)) + 1) = Trim(tMatches(iMatch).sSubMatches(1))
            sClassName(0) = val(sClassName(0)) + 1
        Case "Race":
            If val(sRaceName(0)) >= 6 Then GoTo skip_match
            sRaceName(val(sRaceName(0)) + 1) = Trim(tMatches(iMatch).sSubMatches(1))
            sRaceName(0) = val(sRaceName(0)) + 1
        
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
        Case "Strength":
            If nStrength(0) >= 6 Then GoTo skip_match
            nStrength(nStrength(0) + 1) = Trim(tMatches(iMatch).sSubMatches(1))
            nStrength(0) = nStrength(0) + 1
        Case "Intellect":
            If nIntellect(0) >= 6 Then GoTo skip_match
            nIntellect(nIntellect(0) + 1) = Trim(tMatches(iMatch).sSubMatches(1))
            nIntellect(0) = nIntellect(0) + 1
        Case "Health":
            If nHealth(0) >= 6 Then GoTo skip_match
            nHealth(nHealth(0) + 1) = Trim(tMatches(iMatch).sSubMatches(1))
            nHealth(0) = nHealth(0) + 1
        Case "Charm":
            If nCharm(0) >= 6 Then GoTo skip_match
            nCharm(nCharm(0) + 1) = Trim(tMatches(iMatch).sSubMatches(1))
            nCharm(0) = nCharm(0) + 1
        Case "Willpower":
            If nWillpower(0) >= 6 Then GoTo skip_match
            nWillpower(nWillpower(0) + 1) = Trim(tMatches(iMatch).sSubMatches(1))
            nWillpower(0) = nWillpower(0) + 1
        Case "Spellcasting":
            If nSpellcasting(0) >= 6 Then GoTo skip_match
            nSpellcasting(nSpellcasting(0) + 1) = Trim(tMatches(iMatch).sSubMatches(1))
            nSpellcasting(0) = nSpellcasting(0) + 1
            
        Case "Hits":
            If nHitPoints(0) >= 6 Then GoTo skip_match
            nHitPoints(nHitPoints(0) + 1) = Trim(tMatches(iMatch).sSubMatches(2))
            nHitPoints(0) = nHitPoints(0) + 1
        Case "Mana":
            If nMaxMana(0) >= 6 Then GoTo skip_match
            nMaxMana(nMaxMana(0) + 1) = Trim(tMatches(iMatch).sSubMatches(2))
            nMaxMana(0) = nMaxMana(0) + 1
            
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
    
    sChar = mid(sPastedText, x + y - 1, 1)
    
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
                Select Case UCase(mid(sText, x2 + 1, Len(sText) - x2 - 1))
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
                If x = 16 Then
                    nWeaponNum(iChar) = tabItems.Fields("Number")
                    nWeaponSpeed(iChar) = tabItems.Fields("Speed")
                    nWeaponSTR(iChar) = tabItems.Fields("StrReq")
                End If
                
                nAccyWorn(iChar) = nAccyWorn(iChar) + tabItems.Fields("Accy")
                
                For y = 0 To 19
                    If tabItems.Fields("Abil-" & y) > 0 And tabItems.Fields("AbilVal-" & y) <> 0 Then
                        Select Case tabItems.Fields("Abil-" & y)
                            Case 7: tCharacter(iChar).nCrit = tCharacter(iChar).nCrit + tabItems.Fields("AbilVal-" & y)
                            Case 11: tCharacter(iChar).nPlusMaxDamage = tCharacter(iChar).nPlusMaxDamage + tabItems.Fields("AbilVal-" & y)
                            Case 13: tCharacter(iChar).nPlusBSaccy = tCharacter(iChar).nPlusBSaccy + tabItems.Fields("AbilVal-" & y)
                            Case 14: tCharacter(iChar).nPlusBSmindmg = tCharacter(iChar).nPlusBSmindmg + tabItems.Fields("AbilVal-" & y)
                            Case 15: tCharacter(iChar).nPlusBSmaxdmg = tCharacter(iChar).nPlusBSmaxdmg + tabItems.Fields("AbilVal-" & y)
                            Case 37: tCharacter(iChar).nMAPlusSkill(1) = tCharacter(iChar).nMAPlusSkill(1) + tabItems.Fields("AbilVal-" & y)
                            Case 40: tCharacter(iChar).nMAPlusAccy(1) = tCharacter(iChar).nMAPlusAccy(1) + tabItems.Fields("AbilVal-" & y)
                            Case 34: tCharacter(iChar).nMAPlusDmg(1) = tCharacter(iChar).nMAPlusDmg(1) + tabItems.Fields("AbilVal-" & y)
                            Case 38: tCharacter(iChar).nMAPlusSkill(2) = tCharacter(iChar).nMAPlusSkill(2) + tabItems.Fields("AbilVal-" & y)
                            Case 41: tCharacter(iChar).nMAPlusAccy(2) = tCharacter(iChar).nMAPlusAccy(2) + tabItems.Fields("AbilVal-" & y)
                            Case 35: tCharacter(iChar).nMAPlusDmg(2) = tCharacter(iChar).nMAPlusDmg(2) + tabItems.Fields("AbilVal-" & y)
                            Case 39: tCharacter(iChar).nMAPlusSkill(3) = tCharacter(iChar).nMAPlusSkill(3) + tabItems.Fields("AbilVal-" & y)
                            Case 42: tCharacter(iChar).nMAPlusAccy(3) = tCharacter(iChar).nMAPlusAccy(3) + tabItems.Fields("AbilVal-" & y)
                            Case 36: tCharacter(iChar).nMAPlusDmg(3) = tCharacter(iChar).nMAPlusDmg(3) + tabItems.Fields("AbilVal-" & y)
                            Case 19: tCharacter(iChar).nStealth = tCharacter(iChar).nStealth + tabItems.Fields("AbilVal-" & y)
                            Case 145: tCharacter(iChar).nManaRegen = tCharacter(iChar).nManaRegen + tabItems.Fields("AbilVal-" & y)
                            Case 34: 'dodge
                                nPlusDodge(iChar) = nPlusDodge(iChar) + tabItems.Fields("AbilVal-" & y)
                            Case 123: 'hpregen
                                nPlusRegen(iChar) = nPlusRegen(iChar) + tabItems.Fields("AbilVal-" & y)
                            Case 22: 'accy
                                If tabItems.Fields("AbilVal-" & y) > nAccyAbil(iChar) Then nAccyAbil(iChar) = tabItems.Fields("AbilVal-" & y)
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
    
    If val(sName(0)) >= iChar Then txtPastePartyName(iChar).Text = sName(iChar)
    If nAC(0) >= iChar Then txtPastePartyAC(iChar).Text = nAC(iChar)
    If nDR(0) >= iChar Then txtPastePartyDR(iChar).Text = nDR(iChar)
    If nMR(0) >= iChar Then txtPastePartyMR(iChar).Text = nMR(iChar)
    If nHitPoints(0) >= iChar Then txtPastePartyHitpoints(iChar).Text = nHitPoints(iChar): tCharacter(iChar).nHP = nHitPoints(iChar)
    
    If val(sClassName(0)) >= iChar Then
        If frmMain.cmbGlobalClass(0).ListCount > 0 Then
            For y = 0 To frmMain.cmbGlobalClass(0).ListCount - 1
                If frmMain.cmbGlobalClass(0).List(y) = sClassName(iChar) Then
                    nClass(iChar) = frmMain.cmbGlobalClass(0).ItemData(y)
                    tCharacter(iChar).nClass = nClass(iChar)
                End If
            Next
        End If
    End If
    
    If val(sRaceName(0)) >= iChar Then
        If frmMain.cmbGlobalRace(0).ListCount > 0 Then
            For y = 0 To frmMain.cmbGlobalRace(0).ListCount - 1
                If frmMain.cmbGlobalRace(0).List(y) = sRaceName(iChar) Then
                    nRace(iChar) = frmMain.cmbGlobalRace(0).ItemData(y)
                    tCharacter(iChar).nRace = nRace(iChar)
                End If
            Next
        End If
    End If
    
    If nClass(iChar) > 0 Then
        If ClassHasAbility(nClass(iChar), 51) >= 0 Then chkPastePartyAM(iChar).Value = 1 'anti-magic
        nTemp = ClassHasAbility(nClass(iChar), 34) 'dodge
        If nTemp <> -31337 Then nPlusDodge(iChar) = nPlusDodge(iChar) + nTemp
        nTemp = ClassHasAbility(nClass(iChar), 123) 'hpregen
        If nTemp <> -31337 Then nPlusRegen(iChar) = nPlusRegen(iChar) + nTemp
    End If
    
    If nRace(iChar) > 0 Then
        If RaceHasAbility(nRace(iChar), 51) >= 0 Then chkPastePartyAM(iChar).Value = 1 'anti-magic
        nTemp = RaceHasAbility(nRace(iChar), 34) 'dodge
        If nTemp <> -31337 Then nPlusDodge(iChar) = nPlusDodge(iChar) + nTemp
        nTemp = RaceHasAbility(nRace(iChar), 123) 'hpregen
        If nTemp <> -31337 Then nPlusRegen(iChar) = nPlusRegen(iChar) + nTemp
    End If
    
    If nLevel(iChar) > 0 And nAgility(iChar) > 0 And nCharm(iChar) > 0 Then
        txtPastePartyDodge(iChar).Text = CalcDodge(nLevel(iChar), nAgility(iChar), nCharm(iChar), nPlusDodge(iChar), nCurrentEnc(iChar), nMaxEnc(iChar))
        tCharacter(iChar).nDodge = val(txtPastePartyDodge(iChar).Text)
    End If
    
    If nLevel(iChar) > 0 And nHealth(iChar) > 0 Then
        txtPastePartyRegenHP(iChar).Text = CalcRestingRate(nLevel(iChar), nHealth(iChar), nPlusRegen(iChar), False)
        txtPastePartyRestHP(iChar).Text = CalcRestingRate(nLevel(iChar), nHealth(iChar), nPlusRegen(iChar), True)
        tCharacter(iChar).nHPRegen = val(txtPastePartyRegenHP(iChar).Text)
    End If
    
    If nMaxEnc(iChar) > 0 Then tCharacter(iChar).nEncumPCT = Round((nCurrentEnc(iChar) / nMaxEnc(iChar)) * 100)
    If tCharacter(iChar).nEncumPCT > 100 Then tCharacter(iChar).nEncumPCT = 100
    
    If nClass(iChar) > 0 And nLevel(iChar) > 0 Then
        'also below
        txtPastePartyACCY(iChar).Text = CalculateAccuracy(nClass(iChar), nLevel(iChar), _
            nStrength(iChar), nAgility(iChar), nIntellect(iChar), nCharm(iChar), _
            nAccyWorn(iChar), nAccyAbil(iChar), tCharacter(iChar).nEncumPCT)
        tCharacter(iChar).nAccuracy = val(txtPastePartyACCY(iChar).Text)
    End If
    
    tCharacter(iChar).nLevel = nLevel(iChar)
    tCharacter(iChar).nSTR = nStrength(iChar)
    tCharacter(iChar).nINT = nIntellect(iChar)
    tCharacter(iChar).nAGI = nAgility(iChar)
    tCharacter(iChar).nCHA = nCharm(iChar)
    tCharacter(iChar).nWis = nWillpower(iChar)
    tCharacter(iChar).nHEA = nHealth(iChar)
    tCharacter(iChar).nSpellcasting = nSpellcasting(iChar)
    tCharacter(iChar).nMaxMana = nMaxMana(iChar)
Next iChar

For iChar = 1 To 6
    If optPastyPartyAtkLast(iChar).Value = True And Len(optPastyPartyAtkLast(iChar).Tag) > 0 Then
        If optPastyPartyAtkLast(iChar).Tag <> Trim(txtPastePartyName(iChar).Text) Then sFindAtkLast = optPastyPartyAtkLast(iChar).Tag
        Exit For
    End If
Next iChar
If Len(sFindAtkLast) > 0 Then
    For iChar = 1 To 6
        If sFindAtkLast = Trim(txtPastePartyName(iChar).Text) Then
            optPastyPartyAtkLast(iChar).Value = True
            Exit For
        End If
    Next iChar
End If

bHoldPartyRefresh = False
Call CalculateAverageParty
DoEvents
bHoldPartyRefresh = True

fraPasteParty.Visible = True
txtText.Visible = False
DoEvents

For iChar = 1 To 6
    nAttackTypeMUD = a0_none
    sSpellAttackShort = ""
    nAttackSpellNum = 0
    If nWeaponNum(iChar) > 0 Or tCharacter(iChar).nSpellcasting > 0 Then
        sChar = Trim(txtPastePartyName(iChar).Text)
        If sChar = "" Then sChar = "Character " & iChar
        If tCharacter(iChar).nLevel > 0 Then sChar = sChar & " - " & tCharacter(iChar).nLevel
        If tCharacter(iChar).nClass > 0 Then sChar = sChar & " " & GetClassName(tCharacter(iChar).nClass)
        
        sText = "a"
        If val(txtPastePartySpellDMG(iChar).Text) > 0 Or val(txtPastePartyDMG(iChar).Text) > 0 Then sText = ""
        
        sText = InputBox("Enter attack for " & sChar & vbCrLf & vbCrLf & _
                    "phys attack: a, aa/bash, aaa/smash, bs, pu, kick, jk" & vbCrLf & _
                    "spell attack: enter short code of learnable spell (i.e. lbol)" & vbCrLf & vbCrLf & _
                    "cancel/anything else to skip", "Calculate Attack", sText)
        
        Select Case Trim(sText)
            Case "", "0": GoTo next_char
            Case "a", "at", "att", "attack": nAttackTypeMUD = a5_Normal
            Case "aa", "ba", "bash": nAttackTypeMUD = a6_Bash
            Case "aaa", "sm", "smash": nAttackTypeMUD = a7_Smash
            Case "p", "pu", "punch": nAttackTypeMUD = a1_Punch
            Case "k", "ki", "kick": nAttackTypeMUD = a2_Kick
            Case "j", "jk", "jumpk", "jumpkick": nAttackTypeMUD = a3_Jumpkick
            Case "bs", "backstab": nAttackTypeMUD = a4_Surprise
            Case Else:
                If Len(Trim(sText)) = 4 Then
                    sSpellAttackShort = Trim(sText)
                Else
                    GoTo next_char
                End If
        End Select
    End If
    
    If (nAttackTypeMUD = a6_Bash Or nAttackTypeMUD = a7_Smash) And nWeaponNum(iChar) = 0 Then
        MsgBox "No weapon detected.", vbExclamation + vbOKOnly
        GoTo next_char:
    ElseIf (nAttackTypeMUD = a1_Punch And tCharacter(iChar).nMAPlusSkill(1) = 0) _
            Or (nAttackTypeMUD = a2_Kick And tCharacter(iChar).nMAPlusSkill(2) = 0) _
            Or (nAttackTypeMUD = a3_Jumpkick And tCharacter(iChar).nMAPlusSkill(3) = 0) Then
        MsgBox "Proper MA skill not detected.", vbExclamation + vbOKOnly
        GoTo next_char:
    End If
    
    If Len(sSpellAttackShort) = 4 Then
        If tCharacter(iChar).nSpellcasting = 0 Then
            MsgBox "Spellcast rating not detected.", vbExclamation + vbOKOnly
        Else
            nAttackSpellNum = GetSpellByShort(sSpellAttackShort, nClass(iChar))
        End If
    End If
    
    If nAttackTypeMUD = a0_none And nAttackSpellNum > 0 Then
        
        If tCharacter(iChar).nClass > 0 Then
            tCharacter(iChar).nManaRegen = CalcManaRegen(tCharacter(iChar).nLevel, tCharacter(iChar).nINT, tCharacter(iChar).nWis, tCharacter(iChar).nCHA, _
                                        GetClassMageryLVL(tCharacter(iChar).nClass), GetClassMagery(tCharacter(iChar).nClass), tCharacter(iChar).nManaRegen)
        End If
        tSpellcast = CalculateSpellCast(tCharacter(iChar), nAttackSpellNum, tCharacter(iChar).nLevel)
        nSpellDamage(iChar) = tSpellcast.nAvgRoundDmg
        
    ElseIf nAttackTypeMUD <> a0_none Then
        
        If tCharacter(iChar).nClass > 0 Then
            nCombat = GetClassCombat(tCharacter(iChar).nClass)
            tCharacter(iChar).nCombat = nCombat
        End If
        
        If nAttackTypeMUD = a4_Surprise Or nAttackTypeMUD = a7_Smash Then 'backstab, smash
            nEnergy = 1000
        Else
            nEnergy = CalcEnergyUsed(nCombat, nLevel(iChar), nWeaponSpeed(iChar), nAgility(iChar), nStrength(iChar), _
                tCharacter(iChar).nEncumPCT, nWeaponSTR(iChar), , IIf(nAttackTypeMUD = a4_Surprise, True, False))
        End If
        tCharacter(iChar).nCrit = tCharacter(iChar).nCrit + CalcQuickAndDeadlyBonus(nAgility(iChar), nEnergy, tCharacter(iChar).nEncumPCT)
        
        If nAttackTypeMUD = a6_Bash Or nAttackTypeMUD = a7_Smash Then
            'also above
            txtPastePartyACCY(iChar).Text = CalculateAccuracy(tCharacter(iChar).nClass, tCharacter(iChar).nLevel, _
                tCharacter(iChar).nSTR, tCharacter(iChar).nAGI, tCharacter(iChar).nINT, tCharacter(iChar).nCHA, _
                nAccyWorn(iChar), nAccyAbil(iChar), tCharacter(iChar).nEncumPCT, , nAttackTypeMUD)
            tCharacter(iChar).nAccuracy = val(txtPastePartyACCY(iChar).Text)
        End If
        
        tAttack = CalculateAttack(tCharacter(iChar), nAttackTypeMUD, nWeaponNum(iChar))
        nPhysDamage(iChar) = tAttack.nRoundTotal
        nSwings(iChar) = tAttack.nSwings
        
    End If
    
    If nPhysDamage(iChar) > 0 Then
        txtPastePartyDMG(iChar).Text = nPhysDamage(iChar)
        txtPastePartySwings(iChar).Text = nSwings(iChar)
    ElseIf nSpellDamage(iChar) > 0 Then
        txtPastePartySpellDMG(iChar).Text = nSpellDamage(iChar)
    End If
    
next_char:
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

For x = 1 To 6
    If optPastyPartyAtkLast(x).Value Then bAtkLast = True
Next x

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
    Case 11: 'accy
        GoTo accy_only:
    Case 12: 'spdmg
        GoTo spdmg_only:
    Case 13: 'swings
        GoTo swings_only:
End Select

nTotal = 0
nCount = 0
For x = 1 To 6
    If Len(Trim(txtPastePartyAC(x).Text)) > 0 Then
        nTotal = nTotal + val(txtPastePartyAC(x).Text) + IIf(optPastyPartyAtkLast(x).Value, val(txtPastePartyAC(x).Text), 0)
        nCount = nCount + 1
    End If
Next x
If nCount > 0 Then txtPastePartyAC(0).Text = Round(nTotal / (nCount + IIf(bAtkLast, 1, 0)))
If nWhat > 0 Then GoTo out

dr_only:
nTotal = 0
nCount = 0
For x = 1 To 6
    If Len(Trim(txtPastePartyDR(x).Text)) > 0 Then
        nTotal = nTotal + val(txtPastePartyDR(x).Text) + IIf(optPastyPartyAtkLast(x).Value, val(txtPastePartyDR(x).Text), 0)
        nCount = nCount + 1
    End If
Next x
If nCount > 0 Then txtPastePartyDR(0).Text = Round(nTotal / (nCount + IIf(bAtkLast, 1, 0)))
If nWhat > 0 Then GoTo out

mr_only:
nTotal = 0
nCount = 0
For x = 1 To 6
    If Len(Trim(txtPastePartyMR(x).Text)) > 0 Then
        nTotal = nTotal + val(txtPastePartyMR(x).Text) + IIf(optPastyPartyAtkLast(x).Value, val(txtPastePartyMR(x).Text), 0)
        nCount = nCount + 1
    End If
Next x
If nCount > 0 Then txtPastePartyMR(0).Text = Round(nTotal / (nCount + IIf(bAtkLast, 1, 0)))
If nWhat > 0 Then GoTo out

dodge_only:
nTotal = 0
nCount = 0
For x = 1 To 6
    If Len(Trim(txtPastePartyDodge(x).Text)) > 0 Then
        nTotal = nTotal + val(txtPastePartyDodge(x).Text) + IIf(optPastyPartyAtkLast(x).Value, val(txtPastePartyDodge(x).Text), 0)
        nCount = nCount + 1
    End If
Next x
If nCount > 0 Then txtPastePartyDodge(0).Text = Round(nTotal / (nCount + IIf(bAtkLast, 1, 0)))
If nWhat > 0 Then GoTo out

hp_only:
nTotal = 0
nCount = 0
For x = 1 To 6
    If Len(Trim(txtPastePartyHitpoints(x).Text)) > 0 Then
        nTotal = nTotal + val(txtPastePartyHitpoints(x).Text)
        nCount = nCount + 1
    End If
Next x
If nCount > 0 Then txtPastePartyHitpoints(0).Text = Round(nTotal / nCount)
If nWhat > 0 Then GoTo out

regen_only:
nTotal = 0
nCount = 0
For x = 1 To 6
    If Len(Trim(txtPastePartyRegenHP(x).Text)) > 0 Then
        nTotal = nTotal + val(txtPastePartyRegenHP(x).Text) + IIf(optPastyPartyAtkLast(x).Value, val(txtPastePartyRegenHP(x).Text), 0)
        nCount = nCount + 1
    End If
Next x
If nCount > 0 Then txtPastePartyRegenHP(0).Text = Round(nTotal / (nCount + IIf(bAtkLast, 1, 0)))
If nWhat > 0 Then GoTo out

rest_only:
nTotal = 0
nCount = 0
For x = 1 To 6
    If Len(Trim(txtPastePartyRestHP(x).Text)) > 0 Then
        nTotal = nTotal + val(txtPastePartyRestHP(x).Text)
        nCount = nCount + 1
    End If
Next x
If nCount > 0 Then txtPastePartyRestHP(0).Text = Round(nTotal / nCount)
If nWhat > 0 Then GoTo out

heal_only:
nTotal = 0
nCount = 0
For x = 1 To 6
    If Len(Trim(txtPastePartyHeals(x).Text)) > 0 Then
        nTotal = nTotal + val(txtPastePartyHeals(x).Text)
        nCount = nCount + 1
    End If
Next x
If nTotal > 0 Then txtPastePartyHeals(0).Text = nTotal
If nWhat > 0 Then GoTo out

dmg_only:
nTotal = 0
nCount = 0
For x = 1 To 6
    If Len(Trim(txtPastePartyDMG(x).Text)) > 0 Then
        nTotal = nTotal + val(txtPastePartyDMG(x).Text)
        nCount = nCount + 1
    End If
Next x
If nCount > 0 Then txtPastePartyDMG(0).Text = Round(nTotal / nCount)
If nWhat > 0 Then GoTo swings_only

accy_only:
nTotal = 0
nCount = 0
For x = 1 To 6
    If Len(Trim(txtPastePartyACCY(x).Text)) > 0 Then
        nTotal = nTotal + val(txtPastePartyACCY(x).Text)
        nCount = nCount + 1
    End If
Next x
If nCount > 0 Then txtPastePartyACCY(0).Text = Round(nTotal / nCount)
If nWhat > 0 Then GoTo out

am_only:
nCount = 0
For x = 1 To 6
    If chkPastePartyAM(x).Value = 1 Then
        nCount = nCount + 1
    End If
Next x
txtPastePartyAMTotal.Text = nCount
If nWhat > 0 Then GoTo out

spdmg_only:
nTotal = 0
nCount = 0
For x = 1 To 6
    If Len(Trim(txtPastePartySpellDMG(x).Text)) > 0 Then
        nTotal = nTotal + val(txtPastePartySpellDMG(x).Text)
        nCount = nCount + 1
    End If
Next x
If nCount > 0 Then txtPastePartySpellDMG(0).Text = Round(nTotal / nCount)
If nWhat > 0 Then GoTo out

swings_only:
nTotal = 0
nCount = 0
For x = 1 To 6
    If Len(Trim(txtPastePartySwings(x).Text)) > 0 And Len(Trim(txtPastePartyDMG(x).Text)) > 0 Then
        nTotal = nTotal + val(txtPastePartySwings(x).Text)
        nCount = nCount + 1
    ElseIf Len(Trim(txtPastePartyDMG(x).Text)) > 0 Then 'e.g. and swings = "" or 0
        nTotal = nTotal + 1
        nCount = nCount + 1
    End If
Next x
If nCount > 0 Then txtPastePartySwings(0).Text = Round(nTotal / nCount, 1)
If nWhat > 0 Then GoTo out

out:
nPartySize = 0
For x = 1 To 6
    If chkPastePartyAM(x).Value = 1 _
        Or Trim(txtPastePartyAC(x).Text) <> "" _
        Or Trim(txtPastePartyDR(x).Text) <> "" _
        Or Trim(txtPastePartyMR(x).Text) <> "" _
        Or Trim(txtPastePartyDodge(x).Text) <> "" _
        Or Trim(txtPastePartyHitpoints(x).Text) <> "" _
        Or Trim(txtPastePartyRestHP(x).Text) <> "" _
        Or Trim(txtPastePartyRegenHP(x).Text) <> "" _
        Or Trim(txtPastePartyHeals(x).Text) <> "" _
        Or Trim(txtPastePartySwings(x).Text) <> "" _
        Or Trim(txtPastePartyACCY(x).Text) <> "" _
        Or Trim(txtPastePartySpellDMG(x).Text) <> "" _
        Or Trim(txtPastePartyDMG(x).Text) <> "" Then
        nPartySize = nPartySize + 1
    End If
Next x
txtPastePartyPartyTotal.Text = nPartySize

If bAtkLast Then
    For x = 1 To 6
        If optPastyPartyAtkLast(x).Value = True And Len(Trim(txtPastePartyName(x).Text)) > 0 Then
            optPastyPartyAtkLast(x).Tag = Trim(txtPastePartyName(x).Text)
        Else
            optPastyPartyAtkLast(x).Tag = ""
        End If
    Next x
End If
On Error Resume Next
Exit Sub
error:
Call HandleError("CalculateAverageParty")
Resume out:
End Sub

Private Sub cmdPaste_Click()
Dim nYesNo As Integer, sSkip1 As String, sSkip2 As String

'paste char
sSkip1 = "Paste the commands below into your game to get your stats." _
        & vbCrLf & "Copy and paste the output here." _
        & vbCrLf _
        & vbCrLf & "powers" & vbCrLf & "spells" & vbCrLf & "inventory" & vbCrLf & "stat" _
        & vbCrLf _
        & vbCrLf & "or, create a macro: sp^Mi^Mstat^M"

'paste party
sSkip2 = vbCrLf & "Paste each character's stat and inventory outputs, listed" _
                    & vbCrLf & "one after another, in this window and then click continue." _
            & vbCrLf & vbCrLf & "Note that any *spell buffs* for dodge and hp regen will not" _
                    & vbCrLf & "be accounted for as they are not reflected on the stat sheet." _
            & vbCrLf & vbCrLf & "You can click continue now to skip this and go to the party screen."
            
If fraPasteParty.Visible Then
    nYesNo = MsgBox("Are you sure?", vbYesNo + vbDefaultButton2 + vbQuestion, "Sure?")
    If Not nYesNo = vbYes Then Exit Sub
End If

If Not Clipboard.GetText = "" Then
    If Trim(txtText.Text) = "" Or txtText.Text = sSkip1 Or txtText.Text = sSkip2 Then
        nYesNo = vbYes
    ElseIf Not Trim(txtText.Text) = "" And fraPasteParty.Visible = False Then
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

Private Sub cmdPasteMegaDmg_Click(Index As Integer)
On Error GoTo error:
Dim tMatches() As RegexMatches, sRegexPattern As String ', sSubMatches() As String, sSubValues() As String
Dim sPastedText As String, nCell As Integer

sPastedText = Clipboard.GetText
If sPastedText = "" Then Exit Sub

If Index <= 6 Then
    nCell = Index
    sRegexPattern = "Round:[^\r\n]+Avg:(\d+)"
Else
    nCell = Index - 6
    sRegexPattern = "Pre[^\r\n]+(?:\r|\n)+[^\r\n]+Avg:(\d+)[^\r\n]*(?:\r|\n)+Round"
End If

tMatches() = RegExpFindv2(sPastedText, sRegexPattern, False, True, False)
If (UBound(tMatches()) = 0 And Len(tMatches(0).sFullMatch) = 0) Or Len(tMatches(0).sSubMatches(0)) = 0 Then Exit Sub

If Index <= 6 Then
    txtPastePartyDMG(nCell).Text = tMatches(0).sSubMatches(0)
Else
    txtPastePartySpellDMG(nCell).Text = tMatches(0).sSubMatches(0)
End If

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("cmdPasteMegaDmg_Click")
Resume out:
End Sub

Private Sub cmdPasteQ_Click(Index As Integer)
cmdContinue.SetFocus
Select Case Index
    Case 0: 'regen / rest
        MsgBox "The left value is the character's calculated non-resting hp regen rate. " _
            & "This occurs every 30 seconds. When you click continue, this value will be divided by 6 to get a per round [5 second] rate and then factor that into the sustainable damage/healing per round." _
            & vbCrLf & vbCrLf & "The right value is the hp regen rate while resting. This occurs every 20 seconds when resting and is factored into the calculation for time spent resting.", vbInformation
        
    Case 1: 'heals
        MsgBox "Enter each character's sustainable healing (to self or party) per round from spells. Note that sustainable here means without requiring to rest." _
            & vbCrLf & vbCrLf & "So, for instance, if you could cast MAHE once every 2 rounds without requiring to rest, take your average MAHE and divide it by 2." _
            & vbCrLf & vbCrLf & "[non-resting regen rate / 6] is added to the saved output to reflect natural hp regen.", vbInformation
        
    Case 2: 'attack last
        MsgBox "Selecting one character with attack last will cause that character's defenses to be counted an extra time in the average.", vbInformation
        
    Case 3: 'phys dmg
        MsgBox "Enter each character's average per-round PHYSICAL damage output." _
            & vbCrLf & vbCrLf & "Use the button to the right of each box to paste what you get from the copy button in MegaMUD. " _
            & "Note that MegaMUD's report will be relative to the defenses of the specific mobs you're fighting and will be " _
            & "over/undercutting each character's damage output when calculated against other lair/mob defenses." _
            & vbCrLf & vbCrLf & "Enter the damage as if it were VS 0/0/0 AC/DR/Dodge.", vbInformation
        
    Case 4: 'general help
        MsgBox "Only fields that have a value are considered in the averages. " _
            & "Likewise, only the fields in the saved area with actual values (a 0 is a value, a blank field is not) will be written back to MME when you click continue. " _
            & "This would allow you to only paste/populate some stats without altering others." _
            & vbCrLf & vbCrLf & "Be sure to click the header buttons for additional information.", vbInformation
    
    Case 5: 'magic dmg
        MsgBox "Enter each character's average per-round MAGIC/SPELL damage output." _
            & vbCrLf & vbCrLf & "Use the button to the right of each box to paste what you get from the copy button in MegaMUD. " _
            & "Note that MegaMUD's report will be relative to the defenses of the specific mobs you're fighting." _
            & vbCrLf & vbCrLf & "Enter the damage as if it were VS 50 MR so that the damage can be appropriately calculated against mob/lair defenses. All party magic damage is reduced by MR defense.", vbInformation
    
    Case 6: 'swings
        MsgBox "Enter the average of the number of swings from each melee character. This is utilized to help determine DR reduction. " _
            & "Note that the damage output entered should be the total average per round damage of each character from the party, without consideration of swings.", vbInformation + vbOKOnly
End Select
End Sub

Private Sub Form_Load()
On Error GoTo error:

Call SetWindowLong(Me.hWnd, GWL_HWNDPARENT, 0)

cmdPasteQ(2).Caption = "Attack" & vbCrLf & "Last"
lblLabelArray(4).Caption = "Anti" & vbCrLf & "Magic"

tWindowSize.twpMinWidth = 13260
tWindowSize.twpMinHeight = 4470
Call SubclassFormMinMaxSize(Me, tWindowSize)

If frmMain.WindowState = vbMinimized Then
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
Else
    Me.Left = frmMain.Left + ((frmMain.Width - Me.Width) / 2)
    Me.Top = frmMain.Top + ((frmMain.Height - Me.Height) / 2)
End If

timWindowMove.Enabled = True

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

txtText.Width = Me.ScaleWidth - txtText.Left - 50
txtText.Height = Me.ScaleHeight - txtText.Top - 50

End Sub

Private Sub optPastyPartyAtkLast_Click(Index As Integer)
Call CalculateAverageParty
On Error Resume Next
cmdContinue.SetFocus
End Sub

Private Sub timWindowMove_Timer()
Call MonitorFormTimer(Me)
End Sub

Private Sub txtPastePartyAC_Change(Index As Integer)
If Index = 0 Then Exit Sub
Call CalculateAverageParty(ac)
End Sub

Private Sub txtPastePartyAC_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyAC(Index))
End Sub

Private Sub txtPastePartyAC_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPastePartyACCY_Change(Index As Integer)
If Index = 0 Then Exit Sub
Call CalculateAverageParty(accy)
End Sub

Private Sub txtPastePartyACCY_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyACCY(Index))
End Sub

Private Sub txtPastePartyACCY_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPastePartyAMTotal_GotFocus()
Call SelectAll(txtPastePartyAMTotal)
End Sub

Private Sub txtPastePartyAMTotal_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPastePartyDMG_Change(Index As Integer)
If Index = 0 Then Exit Sub
Call CalculateAverageParty(dmg)
'If val(txtPastePartyDMG(Index).Text) > 0 And Trim(txtPastePartySpellDMG(Index).Text) = "" Then txtPastePartySpellDMG(Index).Text = "0"
End Sub

Private Sub txtPastePartyDMG_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyDMG(Index))
End Sub

Private Sub txtPastePartyDMG_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPastePartyDodge_Change(Index As Integer)
If Index = 0 Then Exit Sub
Call CalculateAverageParty(Dodge)
End Sub

Private Sub txtPastePartyDodge_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyDodge(Index))
End Sub

Private Sub txtPastePartyDodge_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPastePartyDR_Change(Index As Integer)
If Index = 0 Then Exit Sub
Call CalculateAverageParty(DR)
End Sub

Private Sub txtPastePartyDR_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyDR(Index))
End Sub

Private Sub txtPastePartyDR_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPastePartyHeals_Change(Index As Integer)
If Index = 0 Then Exit Sub
Call CalculateAverageParty(heal)
End Sub

Private Sub txtPastePartyHeals_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyHeals(Index))
End Sub

Private Sub txtPastePartyHeals_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPastePartyHitpoints_Change(Index As Integer)
If Index = 0 Then Exit Sub
Call CalculateAverageParty(HP)
End Sub

Private Sub txtPastePartyHitpoints_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyHitpoints(Index))
End Sub

Private Sub txtPastePartyHitpoints_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPastePartyMR_Change(Index As Integer)
If Index = 0 Then Exit Sub
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
If Index = 0 Then Exit Sub
Call CalculateAverageParty(Regen)
End Sub

Private Sub txtPastePartyRegenHP_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyRegenHP(Index))
End Sub

Private Sub txtPastePartyRegenHP_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPastePartyRestHP_Change(Index As Integer)
If Index = 0 Then Exit Sub
Call CalculateAverageParty(Rest)
End Sub

Private Sub txtPastePartyRestHP_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyRestHP(Index))
End Sub

Private Sub txtPastePartyRestHP_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPastePartySpellDMG_Change(Index As Integer)
If Index = 0 Then Exit Sub
Call CalculateAverageParty(SPdmg)
'If val(txtPastePartySpellDMG(Index).Text) > 0 And Trim(txtPastePartyDMG(Index).Text) = "" Then txtPastePartyDMG(Index).Text = "0"
End Sub

Private Sub txtPastePartySpellDMG_GotFocus(Index As Integer)
Call SelectAll(txtPastePartySpellDMG(Index))
End Sub

Private Sub txtPastePartySpellDMG_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPastePartySwings_Change(Index As Integer)
If Index = 0 Then Exit Sub
Call CalculateAverageParty(Swings)
End Sub

Private Sub txtPastePartySwings_GotFocus(Index As Integer)
Call SelectAll(txtPastePartySwings(Index))
End Sub

Private Sub txtPastePartySwings_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii, True)
End Sub

Private Sub txtText_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyA And (Shift And vbCtrlMask) <> 0 And Len(txtText.Text) > 0 Then
    txtText.SelStart = 0
    txtText.SelLength = Len(txtText.Text)
    KeyCode = 0
End If
End Sub

Private Sub txtText_KeyPress(KeyAscii As Integer)
If KeyAscii = 1 Then 'Ctrl+A ?
    KeyAscii = 0
End If
End Sub
