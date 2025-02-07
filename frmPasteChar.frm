VERSION 5.00
Object = "{20D5284F-7B23-4F0A-B8B1-6C9D18B64F1C}#1.0#0"; "exlimiter.ocx"
Begin VB.Form frmPasteChar 
   Caption         =   "Paste Characters/Equipment/Spells"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9345
   ControlBox      =   0   'False
   Icon            =   "frmPasteChar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   9345
   Begin VB.Frame fraPasteParty 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3855
      Left            =   60
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   7875
      Begin VB.Frame fraMISC 
         BorderStyle     =   0  'None
         Caption         =   "With the updates I have coming out in the  "
         Height          =   3495
         Left            =   240
         TabIndex        =   5
         Top             =   180
         Width           =   7395
         Begin VB.CommandButton cmdAttackLastQ 
            Caption         =   "?"
            Height          =   255
            Left            =   7080
            TabIndex        =   98
            Top             =   840
            Width           =   255
         End
         Begin VB.TextBox txtPastePartyPartyTotal 
            Height          =   285
            Left            =   720
            TabIndex        =   96
            Top             =   420
            Width           =   555
         End
         Begin VB.TextBox txtPastePartyRestHP 
            Height          =   285
            Index           =   6
            Left            =   4920
            TabIndex        =   56
            Top             =   3000
            Width           =   495
         End
         Begin VB.OptionButton optPastyPartyAtkLast 
            Caption         =   "Option1"
            Height          =   195
            Index           =   6
            Left            =   6660
            TabIndex        =   58
            Top             =   3060
            Width           =   195
         End
         Begin VB.CheckBox chkPastePartyAM 
            Caption         =   "Check1"
            Height          =   195
            Index           =   6
            Left            =   5880
            TabIndex        =   57
            ToolTipText     =   "Anti-Magic"
            Top             =   3060
            Width           =   195
         End
         Begin VB.TextBox txtPastePartyHitpoints 
            Height          =   285
            Index           =   6
            Left            =   4020
            TabIndex        =   55
            Top             =   3000
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyDodge 
            Height          =   285
            Index           =   6
            Left            =   3360
            TabIndex        =   54
            Top             =   3000
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyMR 
            Height          =   285
            Index           =   6
            Left            =   2700
            TabIndex        =   53
            Top             =   3000
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDR 
            Height          =   285
            Index           =   6
            Left            =   2040
            TabIndex        =   52
            Top             =   3000
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyAC 
            Height          =   285
            Index           =   6
            Left            =   1380
            TabIndex        =   51
            Top             =   3000
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyName 
            BackColor       =   &H80000016&
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   360
            TabIndex        =   82
            Top             =   3000
            Width           =   915
         End
         Begin VB.TextBox txtPastePartyRestHP 
            Height          =   285
            Index           =   5
            Left            =   4920
            TabIndex        =   48
            Top             =   2640
            Width           =   495
         End
         Begin VB.OptionButton optPastyPartyAtkLast 
            Caption         =   "Option1"
            Height          =   195
            Index           =   5
            Left            =   6660
            TabIndex        =   50
            Top             =   2700
            Width           =   195
         End
         Begin VB.CheckBox chkPastePartyAM 
            Caption         =   "Check1"
            Height          =   195
            Index           =   5
            Left            =   5880
            TabIndex        =   49
            ToolTipText     =   "Anti-Magic"
            Top             =   2700
            Width           =   195
         End
         Begin VB.TextBox txtPastePartyHitpoints 
            Height          =   285
            Index           =   5
            Left            =   4020
            TabIndex        =   47
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyDodge 
            Height          =   285
            Index           =   5
            Left            =   3360
            TabIndex        =   46
            Top             =   2640
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyMR 
            Height          =   285
            Index           =   5
            Left            =   2700
            TabIndex        =   45
            Top             =   2640
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDR 
            Height          =   285
            Index           =   5
            Left            =   2040
            TabIndex        =   44
            Top             =   2640
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyAC 
            Height          =   285
            Index           =   5
            Left            =   1380
            TabIndex        =   43
            Top             =   2640
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyName 
            BackColor       =   &H80000016&
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   360
            TabIndex        =   80
            Top             =   2640
            Width           =   915
         End
         Begin VB.TextBox txtPastePartyRestHP 
            Height          =   285
            Index           =   4
            Left            =   4920
            TabIndex        =   40
            Top             =   2280
            Width           =   495
         End
         Begin VB.OptionButton optPastyPartyAtkLast 
            Caption         =   "Option1"
            Height          =   195
            Index           =   4
            Left            =   6660
            TabIndex        =   42
            Top             =   2340
            Width           =   195
         End
         Begin VB.CheckBox chkPastePartyAM 
            Caption         =   "Check1"
            Height          =   195
            Index           =   4
            Left            =   5880
            TabIndex        =   41
            ToolTipText     =   "Anti-Magic"
            Top             =   2340
            Width           =   195
         End
         Begin VB.TextBox txtPastePartyHitpoints 
            Height          =   285
            Index           =   4
            Left            =   4020
            TabIndex        =   39
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyDodge 
            Height          =   285
            Index           =   4
            Left            =   3360
            TabIndex        =   38
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyMR 
            Height          =   285
            Index           =   4
            Left            =   2700
            TabIndex        =   37
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDR 
            Height          =   285
            Index           =   4
            Left            =   2040
            TabIndex        =   36
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyAC 
            Height          =   285
            Index           =   4
            Left            =   1380
            TabIndex        =   35
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyName 
            BackColor       =   &H80000016&
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   360
            TabIndex        =   78
            Top             =   2280
            Width           =   915
         End
         Begin VB.TextBox txtPastePartyRestHP 
            Height          =   285
            Index           =   3
            Left            =   4920
            TabIndex        =   32
            Top             =   1920
            Width           =   495
         End
         Begin VB.OptionButton optPastyPartyAtkLast 
            Caption         =   "Option1"
            Height          =   195
            Index           =   3
            Left            =   6660
            TabIndex        =   34
            Top             =   1980
            Width           =   195
         End
         Begin VB.CheckBox chkPastePartyAM 
            Caption         =   "Check1"
            Height          =   195
            Index           =   3
            Left            =   5880
            TabIndex        =   33
            ToolTipText     =   "Anti-Magic"
            Top             =   1980
            Width           =   195
         End
         Begin VB.TextBox txtPastePartyHitpoints 
            Height          =   285
            Index           =   3
            Left            =   4020
            TabIndex        =   31
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyDodge 
            Height          =   285
            Index           =   3
            Left            =   3360
            TabIndex        =   30
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyMR 
            Height          =   285
            Index           =   3
            Left            =   2700
            TabIndex        =   29
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDR 
            Height          =   285
            Index           =   3
            Left            =   2040
            TabIndex        =   28
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyAC 
            Height          =   285
            Index           =   3
            Left            =   1380
            TabIndex        =   27
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyName 
            BackColor       =   &H80000016&
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   360
            TabIndex        =   76
            Top             =   1920
            Width           =   915
         End
         Begin VB.TextBox txtPastePartyRestHP 
            Height          =   285
            Index           =   2
            Left            =   4920
            TabIndex        =   24
            Top             =   1560
            Width           =   495
         End
         Begin VB.OptionButton optPastyPartyAtkLast 
            Caption         =   "Option1"
            Height          =   195
            Index           =   2
            Left            =   6660
            TabIndex        =   26
            Top             =   1620
            Width           =   195
         End
         Begin VB.CheckBox chkPastePartyAM 
            Caption         =   "Check1"
            Height          =   195
            Index           =   2
            Left            =   5880
            TabIndex        =   25
            ToolTipText     =   "Anti-Magic"
            Top             =   1620
            Width           =   195
         End
         Begin VB.TextBox txtPastePartyHitpoints 
            Height          =   285
            Index           =   2
            Left            =   4020
            TabIndex        =   23
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyDodge 
            Height          =   285
            Index           =   2
            Left            =   3360
            TabIndex        =   22
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyMR 
            Height          =   285
            Index           =   2
            Left            =   2700
            TabIndex        =   21
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDR 
            Height          =   285
            Index           =   2
            Left            =   2040
            TabIndex        =   20
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyAC 
            Height          =   285
            Index           =   2
            Left            =   1380
            TabIndex        =   19
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyName 
            BackColor       =   &H80000016&
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   360
            TabIndex        =   74
            Top             =   1560
            Width           =   915
         End
         Begin VB.OptionButton optPastyPartyAtkLast 
            Caption         =   "Option1"
            Height          =   195
            Index           =   0
            Left            =   6660
            TabIndex        =   66
            Top             =   480
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.TextBox txtPastePartyAMTotal 
            Height          =   285
            Left            =   5700
            TabIndex        =   65
            Top             =   420
            Width           =   555
         End
         Begin VB.TextBox txtPastePartyRestHP 
            Height          =   285
            Index           =   0
            Left            =   4920
            TabIndex        =   64
            Top             =   420
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyHitpoints 
            Height          =   285
            Index           =   0
            Left            =   4020
            TabIndex        =   63
            Top             =   420
            Width           =   735
         End
         Begin VB.TextBox txtPastePartyDodge 
            Height          =   285
            Index           =   0
            Left            =   3360
            TabIndex        =   62
            Top             =   420
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyMR 
            Height          =   285
            Index           =   0
            Left            =   2700
            TabIndex        =   61
            Top             =   420
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDR 
            Height          =   285
            Index           =   0
            Left            =   2040
            TabIndex        =   60
            Top             =   420
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyAC 
            Height          =   285
            Index           =   0
            Left            =   1380
            TabIndex        =   59
            Top             =   420
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyRestHP 
            Height          =   285
            Index           =   1
            Left            =   4920
            TabIndex        =   15
            Top             =   1200
            Width           =   495
         End
         Begin VB.OptionButton optPastyPartyAtkLast 
            Caption         =   "Option1"
            Height          =   195
            Index           =   1
            Left            =   6660
            TabIndex        =   18
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
            Height          =   285
            Index           =   1
            Left            =   1380
            TabIndex        =   6
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDR 
            Height          =   285
            Index           =   1
            Left            =   2040
            TabIndex        =   8
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyMR 
            Height          =   285
            Index           =   1
            Left            =   2700
            TabIndex        =   10
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyDodge 
            Height          =   285
            Index           =   1
            Left            =   3360
            TabIndex        =   12
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtPastePartyHitpoints 
            Height          =   285
            Index           =   1
            Left            =   4020
            TabIndex        =   14
            Top             =   1200
            Width           =   735
         End
         Begin VB.CheckBox chkPastePartyAM 
            Caption         =   "Check1"
            Height          =   195
            Index           =   1
            Left            =   5880
            TabIndex        =   17
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
            TabIndex        =   97
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
            Left            =   6240
            TabIndex        =   95
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
            Left            =   6240
            TabIndex        =   94
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
            Left            =   6240
            TabIndex        =   93
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
            Left            =   6240
            TabIndex        =   92
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
            Left            =   6240
            TabIndex        =   91
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
            Left            =   6240
            TabIndex        =   90
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
            TabIndex        =   89
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
            TabIndex        =   88
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
            TabIndex        =   87
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
            TabIndex        =   86
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
            TabIndex        =   85
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
            TabIndex        =   84
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
            TabIndex        =   83
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
            TabIndex        =   81
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
            TabIndex        =   79
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
            TabIndex        =   77
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
            TabIndex        =   75
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
            Left            =   6420
            TabIndex        =   73
            Top             =   240
            Width           =   675
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   2  'Center
            Caption         =   "[===============  VALUES TO BE SAVED  ===============]"
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
            Left            =   660
            TabIndex        =   72
            Top             =   120
            Width           =   5655
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
            TabIndex        =   71
            Top             =   480
            Width           =   195
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   2  'Center
            Caption         =   "RestHP"
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
            Index           =   7
            Left            =   4860
            TabIndex        =   70
            Top             =   960
            Width           =   735
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
            TabIndex        =   69
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
            TabIndex        =   68
            Top             =   960
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
            Left            =   2760
            TabIndex        =   67
            Top             =   960
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
            Left            =   3300
            TabIndex        =   16
            Top             =   960
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
            Left            =   4080
            TabIndex        =   13
            Top             =   960
            Width           =   615
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
            Left            =   5700
            TabIndex        =   11
            Top             =   770
            Width           =   555
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   2  'Center
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
            Height          =   375
            Index           =   5
            Left            =   6420
            TabIndex        =   9
            Top             =   765
            Width           =   675
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
      Width           =   9240
   End
End
Attribute VB_Name = "frmPasteChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bPasteParty As Boolean
Dim bHoldPartyRefresh As Boolean

Private Sub cmdAttackLastQ_Click()
MsgBox "Selecting one character with attack last will cause that character's defenses to be counted twice in the average.", vbInformation
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

If Me.bPasteParty = True Then
    If fraPasteParty.Visible = False Then
        Call ParsePasteParty
        Exit Sub
    Else
        If Len(Trim(txtPastePartyAC(0).Text)) > 0 Then frmMain.txtMonsterLairFilter(1).Text = Trim(txtPastePartyAC(0).Text)
        If Len(Trim(txtPastePartyDR(0).Text)) > 0 Then frmMain.txtMonsterLairFilter(2).Text = Trim(txtPastePartyDR(0).Text)
        If Len(Trim(txtPastePartyMR(0).Text)) > 0 Then frmMain.txtMonsterLairFilter(3).Text = Trim(txtPastePartyMR(0).Text)
        If Len(Trim(txtPastePartyDodge(0).Text)) > 0 Then frmMain.txtMonsterLairFilter(4).Text = Trim(txtPastePartyDodge(0).Text)
        If Len(Trim(txtPastePartyHitpoints(0).Text)) > 0 Then frmMain.txtMonsterLairFilter(5).Text = Trim(txtPastePartyHitpoints(0).Text)
        If Len(Trim(txtPastePartyRestHP(0).Text)) > 0 Then frmMain.txtMonsterLairFilter(7).Text = Trim(txtPastePartyRestHP(0).Text)
        If Len(Trim(txtPastePartyAMTotal.Text)) > 0 Then frmMain.txtMonsterLairFilter(6).Text = Trim(txtPastePartyAMTotal.Text)
        If Len(Trim(txtPastePartyPartyTotal.Text)) > 0 Then frmMain.txtMonsterLairFilter(0).Text = Trim(txtPastePartyPartyTotal.Text)
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
'Dim sSearch As String, sText As String, sChar As String
', x2 As Integer
'Dim sEquipLoc(0 To 19) As String, bResult As Boolean, nTries As Integer
', bItemsFound As Boolean
'Dim nEncum As Long, nStat As String, sName As String, sWorn(0 To 1) As String
'Dim sCharFile As String, sSectionName As String, nResult As Integer, nYesNo As Integer
Dim tMatches() As RegexMatches, sRegexPattern As String, sSubMatches() As String, sSubValues() As String
Dim sName(6) As String, nMR(6) As Integer, nAC(6) As Integer, nDR(6) As Integer
Dim sRaceName(6) As String, sClassName(6) As String, nClass(6) As Integer, nRace(6) As Integer
Dim nCurrentEnc(6) As Long, nMaxEnc(6) As Long, nHitPoints(6) As Long
Dim nLevel(6) As Integer, nAgility(6) As Integer, nHealth(6) As Integer, nCharm(6) As Integer
Dim x As Integer, y As Integer, iMatch As Integer
Dim sClipBoardText As String


sClipBoardText = frmPasteChar.txtText.Text
If Len(sClipBoardText) < 10 Then GoTo canceled:

'Name: Kratos                           Lives/CP:      9/20
'Race: Half-Ogre   Exp: 5281866477      Perception:     75
'Class: Warrior    Level: 75            Stealth:         0
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
tMatches() = RegExpFindv2(sClipBoardText, sRegexPattern, False, True, False)
If UBound(tMatches()) = 0 And Len(tMatches(0).sFullMatch) = 0 Then
    MsgBox "No matching data.", vbOKOnly + vbExclamation, "Paste Party"
    GoTo canceled:
End If

bHoldPartyRefresh = True

optPastyPartyAtkLast(0).Value = True
For x = 0 To 6
    If x > 0 Then txtPastePartyName(x).Text = ""
    If x > 0 Then chkPastePartyAM(x).Value = 0
    txtPastePartyAC(x).Text = ""
    txtPastePartyDR(x).Text = ""
    txtPastePartyMR(x).Text = ""
    txtPastePartyDodge(x).Text = ""
    txtPastePartyHitpoints(x).Text = ""
    txtPastePartyRestHP(x).Text = ""
Next x
txtPastePartyAMTotal.Text = ""
txtPastePartyPartyTotal.Text = ""
fraPasteParty.Visible = True
txtText.Visible = False
cmdPaste.Enabled = False
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

For x = 1 To 6
    If Val(sName(0)) >= x Then txtPastePartyName(x).Text = sName(x)
    If nAC(0) >= x Then txtPastePartyAC(x).Text = nAC(x)
    If nDR(0) >= x Then txtPastePartyDR(x).Text = nDR(x)
    If nMR(0) >= x Then txtPastePartyMR(x).Text = nMR(x)
    If nHitPoints(0) >= x Then txtPastePartyHitpoints(x).Text = nHitPoints(x)
    
    If Val(sClassName(0)) >= x Then
        If frmMain.cmbGlobalClass(0).ListCount > 0 Then
            For y = 0 To frmMain.cmbGlobalClass(0).ListCount - 1
                If frmMain.cmbGlobalClass(0).List(y) = sClassName(x) Then
                    nClass(x) = frmMain.cmbGlobalClass(0).ItemData(y)
                End If
            Next
        End If
    End If
    
    If Val(sRaceName(0)) >= x Then
        If frmMain.cmbGlobalRace(0).ListCount > 0 Then
            For y = 0 To frmMain.cmbGlobalRace(0).ListCount - 1
                If frmMain.cmbGlobalRace(0).List(y) = sRaceName(x) Then
                    nRace(x) = frmMain.cmbGlobalRace(0).ItemData(y)
                End If
            Next
        End If
    End If
    
    'txtPastePartyDodge(x).Text = ""
    'txtPastePartyHitpoints(x).Text = ""
    'txtPastePartyRestHP(x).Text = ""
Next x

out:
On Error Resume Next
Exit Sub
canceled:
On Error Resume Next
Call cmdCancel_Click
Exit Sub
error:
Call HandleError("ParsePasteParty")
Resume out:
End Sub

Private Sub cmdPaste_Click()
Dim nYesNo As Integer

If Not Clipboard.GetText = "" Then
    If Not txtText.Text = "" Then
        nYesNo = MsgBox("Clear paste area first?", vbYesNo + vbDefaultButton1 + vbQuestion, "Clear?")
    Else
        nYesNo = vbYes
    End If
    
    If nYesNo = vbYes Then
        txtText.Text = Clipboard.GetText
    Else
        txtText.Text = Clipboard.GetText & vbCrLf & vbCrLf & txtText.Text
    End If
End If

End Sub

Private Sub Form_Load()

With EL1
    .CenterOnLoad = True
    .FormInQuestion = Me
    .MinWidth = 575
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
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = vbMinimized Then Exit Sub

txtText.Width = Me.Width - 400
txtText.Height = Me.Height - TITLEBAR_OFFSET - 1000

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

Private Sub txtPastePartyDodge_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyDodge(Index))
End Sub

Private Sub txtPastePartyDodge_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPastePartyDR_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyDR(Index))
End Sub

Private Sub txtPastePartyDR_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPastePartyHitpoints_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyHitpoints(Index))
End Sub

Private Sub txtPastePartyHitpoints_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
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

Private Sub txtPastePartyRestHP_GotFocus(Index As Integer)
Call SelectAll(txtPastePartyRestHP(Index))
End Sub

Private Sub txtPastePartyRestHP_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub
