VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{20D5284F-7B23-4F0A-B8B1-6C9D18B64F1C}#1.0#0"; "exlimiter.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "MMUD Explorer"
   ClientHeight    =   7950
   ClientLeft      =   5280
   ClientTop       =   3900
   ClientWidth     =   10575
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   10575
   Begin VB.Timer timMouseDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   10200
      Top             =   60
   End
   Begin VB.CheckBox chkGlobalFilter 
      Caption         =   "Level:"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   300
      Width           =   735
   End
   Begin VB.Frame fraDatVer 
      Caption         =   "Database Version"
      Height          =   675
      Left            =   6180
      TabIndex        =   19
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   435
         Left            =   3300
         TabIndex        =   729
         Top             =   180
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblDatVer 
         Alignment       =   2  'Center
         Caption         =   "v#.##"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   60
         TabIndex        =   20
         Top             =   240
         Width           =   4185
      End
   End
   Begin VB.Frame frmGlobalFilter 
      Caption         =   "Global Filter"
      Enabled         =   0   'False
      Height          =   675
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   6075
      Begin VB.CommandButton cmdFilterAll 
         Caption         =   "Filter All"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5100
         TabIndex        =   5
         Top             =   180
         Width           =   855
      End
      Begin VB.TextBox txtGlobalLevel 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   870
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "999"
         Top             =   255
         Width           =   495
      End
      Begin VB.ComboBox cmbGlobalClass 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1980
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   225
         Width           =   1275
      End
      Begin VB.ComboBox cmbGlobalAlignment 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   225
         Width           =   1155
      End
      Begin VB.Label lblClass 
         Caption         =   "Class:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1500
         TabIndex        =   17
         Top             =   300
         Width           =   495
      End
      Begin VB.Label lblAlign 
         Caption         =   "Align:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3360
         TabIndex        =   18
         Top             =   300
         Width           =   435
      End
   End
   Begin VB.Frame framButtons 
      Height          =   555
      Left            =   60
      TabIndex        =   21
      Top             =   660
      Width           =   10455
      Begin VB.CommandButton cmdNav 
         Caption         =   "Rooms"
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
         Index           =   10
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "F11"
         Top             =   120
         Width           =   795
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Shops"
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
         Index           =   9
         Left            =   8880
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "F10"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Monsters"
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
         Index           =   8
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "F9"
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Sundry"
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
         Index           =   7
         Left            =   7140
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "F8"
         Top             =   120
         Width           =   795
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Class/Race"
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
         Index           =   6
         Left            =   5940
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "F7"
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Character"
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
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "F6"
         Top             =   120
         Width           =   1035
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Equipment"
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
         Index           =   4
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "F5"
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Compare"
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
         Index           =   3
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "F4"
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Spells"
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
         Index           =   2
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "F3"
         Top             =   120
         Width           =   795
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Armour"
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
         Index           =   1
         Left            =   1140
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "F2"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdNav 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Weapons"
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
         Index           =   0
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "F1"
         Top             =   120
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog oComDag 
      Left            =   1200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin exlimiter.EL EL1 
      Left            =   360
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
   End
   Begin VB.Frame framNav 
      Caption         =   "Weapons"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6075
      Index           =   0
      Left            =   60
      TabIndex        =   26
      Top             =   1200
      Width           =   10455
      Begin VB.CommandButton cmdFilter 
         Caption         =   "&Remove Filter"
         Height          =   555
         Index           =   7
         Left            =   9540
         TabIndex        =   40
         Top             =   240
         Width           =   795
      End
      Begin VB.TextBox txtWeaponSpeed 
         Height          =   315
         Left            =   7140
         TabIndex        =   37
         Text            =   "9999"
         Top             =   465
         Width           =   555
      End
      Begin VB.ComboBox cmbWeaponSpeed 
         Height          =   315
         ItemData        =   "frmMain.frx":0CCA
         Left            =   6480
         List            =   "frmMain.frx":0CD4
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   465
         Width           =   615
      End
      Begin VB.CheckBox chkWeaponNonMagical 
         Caption         =   "Non-Magic"
         Height          =   195
         Left            =   5280
         TabIndex        =   35
         Top             =   570
         Width           =   1155
      End
      Begin VB.CheckBox chkWeaponStaffOnly 
         Caption         =   "Staff Only"
         Height          =   195
         Left            =   4260
         TabIndex        =   33
         Top             =   570
         Width           =   1035
      End
      Begin VB.ComboBox cmbWeaponMagicLevel 
         Height          =   315
         Left            =   7800
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   465
         Width           =   915
      End
      Begin VB.CheckBox chkWeaponNoLimit 
         Caption         =   "No Limit"
         Height          =   195
         Left            =   5280
         TabIndex        =   34
         Top             =   330
         Width           =   975
      End
      Begin VB.CheckBox chkBSAble 
         Caption         =   "BS'able"
         Height          =   195
         Left            =   4260
         TabIndex        =   32
         Top             =   330
         Width           =   855
      End
      Begin VB.CheckBox chkHanded 
         Caption         =   "2H Sharp"
         Height          =   195
         Index           =   3
         Left            =   3240
         TabIndex        =   31
         Top             =   570
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkHanded 
         Caption         =   "1H Sharp"
         Height          =   195
         Index           =   2
         Left            =   2220
         TabIndex        =   29
         Top             =   570
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkHanded 
         Caption         =   "2H Blunt"
         Height          =   195
         Index           =   1
         Left            =   3240
         TabIndex        =   30
         Top             =   330
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkHanded 
         Caption         =   "1H Blunt"
         Height          =   195
         Index           =   0
         Left            =   2220
         TabIndex        =   28
         Top             =   330
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "&Apply Filter"
         Height          =   555
         Index           =   6
         Left            =   8820
         TabIndex        =   39
         Top             =   240
         Width           =   675
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Next"
         Height          =   255
         Index           =   5
         Left            =   1140
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtWeaponFind 
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   510
         Width           =   1995
      End
      Begin MMUD_Explorer.cntSplitter splSplitterNS 
         Height          =   5115
         Index           =   0
         Left            =   120
         TabIndex        =   738
         Top             =   840
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   9022
         Begin MMUD_Explorer.cntSplitter splSplitterWE 
            Height          =   975
            Index           =   0
            Left            =   0
            TabIndex        =   739
            Top             =   4140
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   1720
            Begin VB.TextBox txtWeaponDetail 
               Height          =   975
               Left            =   0
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   42
               Top             =   0
               Width           =   4035
            End
            Begin MSComctlLib.ListView lvWeaponLoc 
               Height          =   975
               Left            =   6660
               TabIndex        =   43
               Top             =   0
               Width           =   3555
               _ExtentX        =   6271
               _ExtentY        =   1720
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
         End
         Begin MSComctlLib.ListView lvWeapons 
            Height          =   3735
            Left            =   0
            TabIndex        =   41
            Tag             =   "STRETCHALL"
            Top             =   0
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   6588
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Label lblLabelArray 
         Alignment       =   2  'Center
         Caption         =   "Speed"
         Height          =   195
         Index           =   17
         Left            =   6480
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblLabelArray 
         Alignment       =   2  'Center
         Caption         =   "Magic LVL"
         Height          =   195
         Index           =   18
         Left            =   7800
         TabIndex        =   27
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame framNav 
      Caption         =   "Rooms"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6075
      Index           =   10
      Left            =   60
      TabIndex        =   208
      Top             =   1200
      Visible         =   0   'False
      Width           =   10455
      Begin VB.CommandButton cmdMapAddMegaCodes 
         Caption         =   "Add &Megamud Codes and Copy"
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
         Left            =   8100
         TabIndex        =   258
         Top             =   5700
         Width           =   2235
      End
      Begin VB.CommandButton cmdMapSwitch 
         Caption         =   "&Switch to Manual Edit"
         Height          =   375
         Left            =   5040
         TabIndex        =   255
         Top             =   5520
         Width           =   1875
      End
      Begin VB.CommandButton cmdMapCopyToClip 
         Caption         =   "Cop&y to Clipboard"
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
         Left            =   8100
         TabIndex        =   257
         Top             =   5460
         Width           =   2235
      End
      Begin VB.CommandButton cmdMapMoveClear 
         Caption         =   "Cl&ear"
         Height          =   375
         Left            =   7080
         TabIndex        =   256
         Top             =   5520
         Width           =   855
      End
      Begin VB.TextBox txtMapMove 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   663
         Top             =   5220
         Width           =   4815
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         Caption         =   "Presets"
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   7860
         TabIndex        =   260
         Top             =   240
         Width           =   2475
         Begin VB.CommandButton cmdEditPreset 
            Caption         =   "!"
            Height          =   315
            Index           =   9
            Left            =   2220
            TabIndex        =   254
            Top             =   3120
            Width           =   195
         End
         Begin VB.CommandButton cmdEditPreset 
            Caption         =   "!"
            Height          =   315
            Index           =   8
            Left            =   2220
            TabIndex        =   252
            Top             =   2820
            Width           =   195
         End
         Begin VB.CommandButton cmdEditPreset 
            Caption         =   "!"
            Height          =   315
            Index           =   7
            Left            =   2220
            TabIndex        =   250
            Top             =   2520
            Width           =   195
         End
         Begin VB.CommandButton cmdEditPreset 
            Caption         =   "!"
            Height          =   315
            Index           =   6
            Left            =   2220
            TabIndex        =   248
            Top             =   2220
            Width           =   195
         End
         Begin VB.CommandButton cmdEditPreset 
            Caption         =   "!"
            Height          =   315
            Index           =   5
            Left            =   2220
            TabIndex        =   246
            Top             =   1920
            Width           =   195
         End
         Begin VB.CommandButton cmdEditPreset 
            Caption         =   "!"
            Height          =   315
            Index           =   4
            Left            =   2220
            TabIndex        =   244
            Top             =   1620
            Width           =   195
         End
         Begin VB.CommandButton cmdEditPreset 
            Caption         =   "!"
            Height          =   315
            Index           =   3
            Left            =   2220
            TabIndex        =   242
            Top             =   1320
            Width           =   195
         End
         Begin VB.CommandButton cmdEditPreset 
            Caption         =   "!"
            Height          =   315
            Index           =   2
            Left            =   2220
            TabIndex        =   240
            Top             =   1020
            Width           =   195
         End
         Begin VB.CommandButton cmdEditPreset 
            Caption         =   "!"
            Height          =   315
            Index           =   1
            Left            =   2220
            TabIndex        =   238
            Top             =   720
            Width           =   195
         End
         Begin VB.CommandButton cmdEditPreset 
            Caption         =   "!"
            Height          =   315
            Index           =   0
            Left            =   2220
            TabIndex        =   236
            Top             =   420
            Width           =   195
         End
         Begin VB.CommandButton cmdMapPreset 
            Caption         =   "Lava Fields"
            Height          =   315
            Index           =   9
            Left            =   120
            TabIndex        =   253
            Top             =   3120
            Width           =   2055
         End
         Begin VB.CommandButton cmdMapPreset 
            Caption         =   "Ancient Ruin"
            Height          =   315
            Index           =   8
            Left            =   120
            TabIndex        =   251
            Top             =   2820
            Width           =   2055
         End
         Begin VB.CommandButton cmdMapPreset 
            Caption         =   "Storm Fortress"
            Height          =   315
            Index           =   7
            Left            =   120
            TabIndex        =   249
            Top             =   2520
            Width           =   2055
         End
         Begin VB.CommandButton cmdMapPreset 
            Caption         =   "Black Fortress"
            Height          =   315
            Index           =   6
            Left            =   120
            TabIndex        =   247
            Top             =   2220
            Width           =   2055
         End
         Begin VB.CommandButton cmdMapPreset 
            Caption         =   "Commander Markus"
            Height          =   315
            Index           =   5
            Left            =   120
            TabIndex        =   245
            Top             =   1920
            Width           =   2055
         End
         Begin VB.CommandButton cmdMapPreset 
            Caption         =   "Rhudar"
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   243
            Top             =   1620
            Width           =   2055
         End
         Begin VB.CommandButton cmdMapPreset 
            Caption         =   "Lost City"
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   241
            Top             =   1320
            Width           =   2055
         End
         Begin VB.CommandButton cmdMapPreset 
            Caption         =   "Arlysia"
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   239
            Top             =   1020
            Width           =   2055
         End
         Begin VB.CommandButton cmdMapPreset 
            Caption         =   "Khazarad"
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   237
            Top             =   720
            Width           =   2055
         End
         Begin VB.CommandButton cmdMapPreset 
            Caption         =   "Town Square"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   235
            Top             =   420
            Width           =   2055
         End
         Begin VB.CommandButton cmdResetPresets 
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
            Height          =   195
            Left            =   1800
            TabIndex        =   234
            Top             =   180
            Width           =   615
         End
         Begin VB.CommandButton cmdMapPresetSelect 
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   1320
            TabIndex        =   233
            Top             =   180
            Width           =   315
         End
         Begin VB.CommandButton cmdMapPresetSelect 
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   1020
            TabIndex        =   232
            Top             =   180
            Width           =   315
         End
         Begin VB.CommandButton cmdMapPresetSelect 
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   720
            TabIndex        =   231
            Top             =   180
            Width           =   315
         End
         Begin VB.CommandButton cmdMapPresetSelect 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   420
            TabIndex        =   230
            Top             =   180
            Width           =   315
         End
         Begin VB.CommandButton cmdMapPresetSelect 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   229
            Top             =   180
            Width           =   315
         End
      End
      Begin VB.PictureBox picMap 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   4845
         Left            =   120
         ScaleHeight     =   321
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   321
         TabIndex        =   261
         Top             =   330
         Width           =   4845
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   1
            Left            =   60
            TabIndex        =   262
            Top             =   60
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   2
            Left            =   300
            TabIndex        =   263
            Top             =   60
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   3
            Left            =   540
            TabIndex        =   264
            Top             =   60
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   4
            Left            =   780
            TabIndex        =   265
            Top             =   60
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   5
            Left            =   1020
            TabIndex        =   266
            Top             =   60
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   6
            Left            =   1260
            TabIndex        =   267
            Top             =   60
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   7
            Left            =   1500
            TabIndex        =   268
            Top             =   60
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   8
            Left            =   1740
            TabIndex        =   269
            Top             =   60
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   9
            Left            =   1980
            TabIndex        =   270
            Top             =   60
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   10
            Left            =   2220
            TabIndex        =   271
            Top             =   60
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   11
            Left            =   2460
            TabIndex        =   272
            Top             =   60
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   12
            Left            =   2700
            TabIndex        =   273
            Top             =   60
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   13
            Left            =   2940
            TabIndex        =   274
            Top             =   60
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   14
            Left            =   3180
            TabIndex        =   275
            Top             =   60
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   15
            Left            =   3420
            TabIndex        =   276
            Top             =   60
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   16
            Left            =   3660
            TabIndex        =   277
            Top             =   60
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   17
            Left            =   3900
            TabIndex        =   278
            Top             =   60
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   18
            Left            =   4140
            TabIndex        =   279
            Top             =   60
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   19
            Left            =   4380
            TabIndex        =   280
            Top             =   60
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   20
            Left            =   4620
            TabIndex        =   281
            Top             =   60
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   21
            Left            =   60
            TabIndex        =   282
            Top             =   300
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   22
            Left            =   300
            TabIndex        =   283
            Top             =   300
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   23
            Left            =   540
            TabIndex        =   284
            Top             =   300
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   24
            Left            =   780
            TabIndex        =   285
            Top             =   300
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   25
            Left            =   1020
            TabIndex        =   286
            Top             =   300
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   26
            Left            =   1260
            TabIndex        =   287
            Top             =   300
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   27
            Left            =   1500
            TabIndex        =   288
            Top             =   300
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   28
            Left            =   1740
            TabIndex        =   289
            Top             =   300
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   29
            Left            =   1980
            TabIndex        =   290
            Top             =   300
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   30
            Left            =   2220
            TabIndex        =   291
            Top             =   300
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   31
            Left            =   2460
            TabIndex        =   292
            Top             =   300
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   32
            Left            =   2700
            TabIndex        =   293
            Top             =   300
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   33
            Left            =   2940
            TabIndex        =   294
            Top             =   300
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   34
            Left            =   3180
            TabIndex        =   295
            Top             =   300
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   35
            Left            =   3420
            TabIndex        =   296
            Top             =   300
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   36
            Left            =   3660
            TabIndex        =   297
            Top             =   300
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   37
            Left            =   3900
            TabIndex        =   298
            Top             =   300
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   38
            Left            =   4140
            TabIndex        =   299
            Top             =   300
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   39
            Left            =   4380
            TabIndex        =   300
            Top             =   300
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   40
            Left            =   4620
            TabIndex        =   301
            Top             =   300
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   41
            Left            =   60
            TabIndex        =   302
            Top             =   540
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   42
            Left            =   300
            TabIndex        =   303
            Top             =   540
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   43
            Left            =   540
            TabIndex        =   304
            Top             =   540
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   44
            Left            =   780
            TabIndex        =   305
            Top             =   540
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   45
            Left            =   1020
            TabIndex        =   306
            Top             =   540
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   46
            Left            =   1260
            TabIndex        =   307
            Top             =   540
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   47
            Left            =   1500
            TabIndex        =   308
            Top             =   540
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   48
            Left            =   1740
            TabIndex        =   309
            Top             =   540
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   49
            Left            =   1980
            TabIndex        =   310
            Top             =   540
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   50
            Left            =   2220
            TabIndex        =   311
            Top             =   540
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   51
            Left            =   2460
            TabIndex        =   312
            Top             =   540
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   52
            Left            =   2700
            TabIndex        =   313
            Top             =   540
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   53
            Left            =   2940
            TabIndex        =   314
            Top             =   540
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   54
            Left            =   3180
            TabIndex        =   315
            Top             =   540
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   55
            Left            =   3420
            TabIndex        =   316
            Top             =   540
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   56
            Left            =   3660
            TabIndex        =   317
            Top             =   540
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   57
            Left            =   3900
            TabIndex        =   318
            Top             =   540
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   58
            Left            =   4140
            TabIndex        =   319
            Top             =   540
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   59
            Left            =   4380
            TabIndex        =   320
            Top             =   540
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   60
            Left            =   4620
            TabIndex        =   321
            Top             =   540
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   61
            Left            =   60
            TabIndex        =   322
            Top             =   780
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   62
            Left            =   300
            TabIndex        =   323
            Top             =   780
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   63
            Left            =   540
            TabIndex        =   324
            Top             =   780
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   64
            Left            =   780
            TabIndex        =   325
            Top             =   780
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   65
            Left            =   1020
            TabIndex        =   326
            Top             =   780
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   66
            Left            =   1260
            TabIndex        =   327
            Top             =   780
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   67
            Left            =   1500
            TabIndex        =   328
            Top             =   780
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   68
            Left            =   1740
            TabIndex        =   329
            Top             =   780
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   69
            Left            =   1980
            TabIndex        =   330
            Top             =   780
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   70
            Left            =   2220
            TabIndex        =   331
            Top             =   780
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   71
            Left            =   2460
            TabIndex        =   332
            Top             =   780
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   72
            Left            =   2700
            TabIndex        =   333
            Top             =   780
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   73
            Left            =   2940
            TabIndex        =   334
            Top             =   780
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   74
            Left            =   3180
            TabIndex        =   335
            Top             =   780
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   75
            Left            =   3420
            TabIndex        =   336
            Top             =   780
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   76
            Left            =   3660
            TabIndex        =   337
            Top             =   780
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   77
            Left            =   3900
            TabIndex        =   338
            Top             =   780
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   78
            Left            =   4140
            TabIndex        =   339
            Top             =   780
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   79
            Left            =   4380
            TabIndex        =   340
            Top             =   780
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   80
            Left            =   4620
            TabIndex        =   341
            Top             =   780
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   81
            Left            =   60
            TabIndex        =   342
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   82
            Left            =   300
            TabIndex        =   343
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   83
            Left            =   540
            TabIndex        =   344
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   84
            Left            =   780
            TabIndex        =   345
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   85
            Left            =   1020
            TabIndex        =   346
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   86
            Left            =   1260
            TabIndex        =   347
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   87
            Left            =   1500
            TabIndex        =   348
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   88
            Left            =   1740
            TabIndex        =   349
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   89
            Left            =   1980
            TabIndex        =   350
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   90
            Left            =   2220
            TabIndex        =   351
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   91
            Left            =   2460
            TabIndex        =   352
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   92
            Left            =   2700
            TabIndex        =   353
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   93
            Left            =   2940
            TabIndex        =   354
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   94
            Left            =   3180
            TabIndex        =   355
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   95
            Left            =   3420
            TabIndex        =   356
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   96
            Left            =   3660
            TabIndex        =   357
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   97
            Left            =   3900
            TabIndex        =   358
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   98
            Left            =   4140
            TabIndex        =   359
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   99
            Left            =   4380
            TabIndex        =   360
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   100
            Left            =   4620
            TabIndex        =   361
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   101
            Left            =   60
            TabIndex        =   362
            Top             =   1260
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   102
            Left            =   300
            TabIndex        =   363
            Top             =   1260
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   103
            Left            =   540
            TabIndex        =   364
            Top             =   1260
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   104
            Left            =   780
            TabIndex        =   365
            Top             =   1260
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   105
            Left            =   1020
            TabIndex        =   366
            Top             =   1260
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   106
            Left            =   1260
            TabIndex        =   367
            Top             =   1260
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   107
            Left            =   1500
            TabIndex        =   368
            Top             =   1260
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   108
            Left            =   1740
            TabIndex        =   369
            Top             =   1260
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   109
            Left            =   1980
            TabIndex        =   370
            Top             =   1260
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   110
            Left            =   2220
            TabIndex        =   371
            Top             =   1260
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   111
            Left            =   2460
            TabIndex        =   372
            Top             =   1260
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   112
            Left            =   2700
            TabIndex        =   373
            Top             =   1260
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   113
            Left            =   2940
            TabIndex        =   374
            Top             =   1260
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   114
            Left            =   3180
            TabIndex        =   375
            Top             =   1260
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   115
            Left            =   3420
            TabIndex        =   376
            Top             =   1260
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   116
            Left            =   3660
            TabIndex        =   377
            Top             =   1260
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   117
            Left            =   3900
            TabIndex        =   378
            Top             =   1260
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   118
            Left            =   4140
            TabIndex        =   379
            Top             =   1260
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   119
            Left            =   4380
            TabIndex        =   380
            Top             =   1260
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   120
            Left            =   4620
            TabIndex        =   381
            Top             =   1260
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   121
            Left            =   60
            TabIndex        =   382
            Top             =   1500
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   122
            Left            =   300
            TabIndex        =   383
            Top             =   1500
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   123
            Left            =   540
            TabIndex        =   384
            Top             =   1500
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   124
            Left            =   780
            TabIndex        =   385
            Top             =   1500
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   125
            Left            =   1020
            TabIndex        =   386
            Top             =   1500
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   126
            Left            =   1260
            TabIndex        =   387
            Top             =   1500
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   127
            Left            =   1500
            TabIndex        =   388
            Top             =   1500
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   128
            Left            =   1740
            TabIndex        =   389
            Top             =   1500
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   129
            Left            =   1980
            TabIndex        =   390
            Top             =   1500
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   130
            Left            =   2220
            TabIndex        =   391
            Top             =   1500
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   131
            Left            =   2460
            TabIndex        =   392
            Top             =   1500
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   132
            Left            =   2700
            TabIndex        =   393
            Top             =   1500
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   133
            Left            =   2940
            TabIndex        =   394
            Top             =   1500
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   134
            Left            =   3180
            TabIndex        =   395
            Top             =   1500
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   135
            Left            =   3420
            TabIndex        =   396
            Top             =   1500
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   136
            Left            =   3660
            TabIndex        =   397
            Top             =   1500
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   137
            Left            =   3900
            TabIndex        =   398
            Top             =   1500
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   138
            Left            =   4140
            TabIndex        =   399
            Top             =   1500
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   139
            Left            =   4380
            TabIndex        =   400
            Top             =   1500
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   140
            Left            =   4620
            TabIndex        =   401
            Top             =   1500
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   141
            Left            =   60
            TabIndex        =   402
            Top             =   1740
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   142
            Left            =   300
            TabIndex        =   403
            Top             =   1740
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   143
            Left            =   540
            TabIndex        =   404
            Top             =   1740
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   144
            Left            =   780
            TabIndex        =   405
            Top             =   1740
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   145
            Left            =   1020
            TabIndex        =   406
            Top             =   1740
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   146
            Left            =   1260
            TabIndex        =   407
            Top             =   1740
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   147
            Left            =   1500
            TabIndex        =   408
            Top             =   1740
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   148
            Left            =   1740
            TabIndex        =   409
            Top             =   1740
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   149
            Left            =   1980
            TabIndex        =   410
            Top             =   1740
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   150
            Left            =   2220
            TabIndex        =   411
            Top             =   1740
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   151
            Left            =   2460
            TabIndex        =   412
            Top             =   1740
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   152
            Left            =   2700
            TabIndex        =   413
            Top             =   1740
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   153
            Left            =   2940
            TabIndex        =   414
            Top             =   1740
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   154
            Left            =   3180
            TabIndex        =   415
            Top             =   1740
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   155
            Left            =   3420
            TabIndex        =   416
            Top             =   1740
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   156
            Left            =   3660
            TabIndex        =   417
            Top             =   1740
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   157
            Left            =   3900
            TabIndex        =   418
            Top             =   1740
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   158
            Left            =   4140
            TabIndex        =   419
            Top             =   1740
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   159
            Left            =   4380
            TabIndex        =   420
            Top             =   1740
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   160
            Left            =   4620
            TabIndex        =   421
            Top             =   1740
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   161
            Left            =   60
            TabIndex        =   422
            Top             =   1980
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   162
            Left            =   300
            TabIndex        =   423
            Top             =   1980
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   163
            Left            =   540
            TabIndex        =   424
            Top             =   1980
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   164
            Left            =   780
            TabIndex        =   425
            Top             =   1980
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   165
            Left            =   1020
            TabIndex        =   426
            Top             =   1980
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   166
            Left            =   1260
            TabIndex        =   427
            Top             =   1980
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   167
            Left            =   1500
            TabIndex        =   428
            Top             =   1980
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   168
            Left            =   1740
            TabIndex        =   429
            Top             =   1980
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   169
            Left            =   1980
            TabIndex        =   430
            Top             =   1980
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   170
            Left            =   2220
            TabIndex        =   431
            Top             =   1980
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   171
            Left            =   2460
            TabIndex        =   432
            Top             =   1980
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   172
            Left            =   2700
            TabIndex        =   433
            Top             =   1980
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   173
            Left            =   2940
            TabIndex        =   434
            Top             =   1980
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   174
            Left            =   3180
            TabIndex        =   435
            Top             =   1980
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   175
            Left            =   3420
            TabIndex        =   436
            Top             =   1980
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   176
            Left            =   3660
            TabIndex        =   437
            Top             =   1980
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   177
            Left            =   3900
            TabIndex        =   438
            Top             =   1980
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   178
            Left            =   4140
            TabIndex        =   439
            Top             =   1980
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   179
            Left            =   4380
            TabIndex        =   440
            Top             =   1980
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   180
            Left            =   4620
            TabIndex        =   441
            Top             =   1980
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   181
            Left            =   60
            TabIndex        =   442
            Top             =   2220
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   182
            Left            =   300
            TabIndex        =   443
            Top             =   2220
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   183
            Left            =   540
            TabIndex        =   444
            Top             =   2220
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   184
            Left            =   780
            TabIndex        =   445
            Top             =   2220
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   185
            Left            =   1020
            TabIndex        =   446
            Top             =   2220
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   186
            Left            =   1260
            TabIndex        =   447
            Top             =   2220
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   187
            Left            =   1500
            TabIndex        =   448
            Top             =   2220
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   188
            Left            =   1740
            TabIndex        =   449
            Top             =   2220
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   189
            Left            =   1980
            TabIndex        =   450
            Top             =   2220
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   190
            Left            =   2220
            TabIndex        =   451
            Top             =   2220
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   191
            Left            =   2460
            TabIndex        =   452
            Top             =   2220
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   192
            Left            =   2700
            TabIndex        =   453
            Top             =   2220
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   193
            Left            =   2940
            TabIndex        =   454
            Top             =   2220
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   194
            Left            =   3180
            TabIndex        =   455
            Top             =   2220
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   195
            Left            =   3420
            TabIndex        =   456
            Top             =   2220
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   196
            Left            =   3660
            TabIndex        =   457
            Top             =   2220
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   197
            Left            =   3900
            TabIndex        =   458
            Top             =   2220
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   198
            Left            =   4140
            TabIndex        =   459
            Top             =   2220
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   199
            Left            =   4380
            TabIndex        =   460
            Top             =   2220
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   200
            Left            =   4620
            TabIndex        =   461
            Top             =   2220
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   201
            Left            =   60
            TabIndex        =   462
            Top             =   2460
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   202
            Left            =   300
            TabIndex        =   463
            Top             =   2460
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   203
            Left            =   540
            TabIndex        =   464
            Top             =   2460
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   204
            Left            =   780
            TabIndex        =   465
            Top             =   2460
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   205
            Left            =   1020
            TabIndex        =   466
            Top             =   2460
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   206
            Left            =   1260
            TabIndex        =   467
            Top             =   2460
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   207
            Left            =   1500
            TabIndex        =   468
            Top             =   2460
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   208
            Left            =   1740
            TabIndex        =   469
            Top             =   2460
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   209
            Left            =   1980
            TabIndex        =   470
            Top             =   2460
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   210
            Left            =   2220
            TabIndex        =   471
            Top             =   2460
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   211
            Left            =   2460
            TabIndex        =   472
            Top             =   2460
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   212
            Left            =   2700
            TabIndex        =   473
            Top             =   2460
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   213
            Left            =   2940
            TabIndex        =   474
            Top             =   2460
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   214
            Left            =   3180
            TabIndex        =   475
            Top             =   2460
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   215
            Left            =   3420
            TabIndex        =   476
            Top             =   2460
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   216
            Left            =   3660
            TabIndex        =   477
            Top             =   2460
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   217
            Left            =   3900
            TabIndex        =   478
            Top             =   2460
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   218
            Left            =   4140
            TabIndex        =   479
            Top             =   2460
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   219
            Left            =   4380
            TabIndex        =   480
            Top             =   2460
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   220
            Left            =   4620
            TabIndex        =   481
            Top             =   2460
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   221
            Left            =   60
            TabIndex        =   482
            Top             =   2700
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   222
            Left            =   300
            TabIndex        =   484
            Top             =   2700
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   223
            Left            =   540
            TabIndex        =   485
            Top             =   2700
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   224
            Left            =   780
            TabIndex        =   486
            Top             =   2700
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   225
            Left            =   1020
            TabIndex        =   487
            Top             =   2700
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   226
            Left            =   1260
            TabIndex        =   488
            Top             =   2700
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   227
            Left            =   1500
            TabIndex        =   489
            Top             =   2700
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   228
            Left            =   1740
            TabIndex        =   490
            Top             =   2700
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   229
            Left            =   1980
            TabIndex        =   491
            Top             =   2700
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   230
            Left            =   2220
            TabIndex        =   492
            Top             =   2700
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   231
            Left            =   2460
            TabIndex        =   493
            Top             =   2700
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   232
            Left            =   2700
            TabIndex        =   494
            Top             =   2700
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   233
            Left            =   2940
            TabIndex        =   495
            Top             =   2700
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   234
            Left            =   3180
            TabIndex        =   496
            Top             =   2700
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   235
            Left            =   3420
            TabIndex        =   497
            Top             =   2700
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   236
            Left            =   3660
            TabIndex        =   498
            Top             =   2700
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   237
            Left            =   3900
            TabIndex        =   499
            Top             =   2700
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   238
            Left            =   4140
            TabIndex        =   500
            Top             =   2700
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   239
            Left            =   4380
            TabIndex        =   501
            Top             =   2700
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   240
            Left            =   4620
            TabIndex        =   502
            Top             =   2700
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   241
            Left            =   60
            TabIndex        =   503
            Top             =   2940
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   242
            Left            =   300
            TabIndex        =   504
            Top             =   2940
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   243
            Left            =   540
            TabIndex        =   505
            Top             =   2940
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   244
            Left            =   780
            TabIndex        =   506
            Top             =   2940
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   245
            Left            =   1020
            TabIndex        =   507
            Top             =   2940
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   246
            Left            =   1260
            TabIndex        =   508
            Top             =   2940
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   247
            Left            =   1500
            TabIndex        =   509
            Top             =   2940
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   248
            Left            =   1740
            TabIndex        =   510
            Top             =   2940
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   249
            Left            =   1980
            TabIndex        =   511
            Top             =   2940
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   250
            Left            =   2220
            TabIndex        =   512
            Top             =   2940
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   251
            Left            =   2460
            TabIndex        =   513
            Top             =   2940
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   252
            Left            =   2700
            TabIndex        =   514
            Top             =   2940
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   253
            Left            =   2940
            TabIndex        =   515
            Top             =   2940
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   254
            Left            =   3180
            TabIndex        =   516
            Top             =   2940
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   255
            Left            =   3420
            TabIndex        =   517
            Top             =   2940
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   256
            Left            =   3660
            TabIndex        =   518
            Top             =   2940
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   257
            Left            =   3900
            TabIndex        =   519
            Top             =   2940
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   258
            Left            =   4140
            TabIndex        =   520
            Top             =   2940
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   259
            Left            =   4380
            TabIndex        =   521
            Top             =   2940
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   260
            Left            =   4620
            TabIndex        =   522
            Top             =   2940
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   261
            Left            =   60
            TabIndex        =   523
            Top             =   3180
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   262
            Left            =   300
            TabIndex        =   524
            Top             =   3180
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   263
            Left            =   540
            TabIndex        =   525
            Top             =   3180
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   264
            Left            =   780
            TabIndex        =   526
            Top             =   3180
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   265
            Left            =   1020
            TabIndex        =   527
            Top             =   3180
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   266
            Left            =   1260
            TabIndex        =   528
            Top             =   3180
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   267
            Left            =   1500
            TabIndex        =   529
            Top             =   3180
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   268
            Left            =   1740
            TabIndex        =   530
            Top             =   3180
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   269
            Left            =   1980
            TabIndex        =   483
            Top             =   3180
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   270
            Left            =   2220
            TabIndex        =   531
            Top             =   3180
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   271
            Left            =   2460
            TabIndex        =   532
            Top             =   3180
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   272
            Left            =   2700
            TabIndex        =   533
            Top             =   3180
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   273
            Left            =   2940
            TabIndex        =   534
            Top             =   3180
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   274
            Left            =   3180
            TabIndex        =   535
            Top             =   3180
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   275
            Left            =   3420
            TabIndex        =   536
            Top             =   3180
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   276
            Left            =   3660
            TabIndex        =   537
            Top             =   3180
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   277
            Left            =   3900
            TabIndex        =   538
            Top             =   3180
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   278
            Left            =   4140
            TabIndex        =   539
            Top             =   3180
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   279
            Left            =   4380
            TabIndex        =   540
            Top             =   3180
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   280
            Left            =   4620
            TabIndex        =   541
            Top             =   3180
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   281
            Left            =   60
            TabIndex        =   542
            Top             =   3420
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   282
            Left            =   300
            TabIndex        =   543
            Top             =   3420
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   283
            Left            =   540
            TabIndex        =   544
            Top             =   3420
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   284
            Left            =   780
            TabIndex        =   545
            Top             =   3420
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   285
            Left            =   1020
            TabIndex        =   546
            Top             =   3420
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   286
            Left            =   1260
            TabIndex        =   547
            Top             =   3420
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   287
            Left            =   1500
            TabIndex        =   548
            Top             =   3420
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   288
            Left            =   1740
            TabIndex        =   549
            Top             =   3420
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   289
            Left            =   1980
            TabIndex        =   550
            Top             =   3420
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   290
            Left            =   2220
            TabIndex        =   551
            Top             =   3420
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   291
            Left            =   2460
            TabIndex        =   552
            Top             =   3420
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   292
            Left            =   2700
            TabIndex        =   553
            Top             =   3420
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   293
            Left            =   2940
            TabIndex        =   554
            Top             =   3420
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   294
            Left            =   3180
            TabIndex        =   555
            Top             =   3420
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   295
            Left            =   3420
            TabIndex        =   556
            Top             =   3420
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   296
            Left            =   3660
            TabIndex        =   557
            Top             =   3420
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   297
            Left            =   3900
            TabIndex        =   558
            Top             =   3420
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   298
            Left            =   4140
            TabIndex        =   559
            Top             =   3420
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   299
            Left            =   4380
            TabIndex        =   560
            Top             =   3420
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   300
            Left            =   4620
            TabIndex        =   561
            Top             =   3420
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   301
            Left            =   60
            TabIndex        =   562
            Top             =   3660
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   302
            Left            =   300
            TabIndex        =   563
            Top             =   3660
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   303
            Left            =   540
            TabIndex        =   564
            Top             =   3660
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   304
            Left            =   780
            TabIndex        =   565
            Top             =   3660
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   305
            Left            =   1020
            TabIndex        =   566
            Top             =   3660
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   306
            Left            =   1260
            TabIndex        =   567
            Top             =   3660
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   307
            Left            =   1500
            TabIndex        =   568
            Top             =   3660
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   308
            Left            =   1740
            TabIndex        =   569
            Top             =   3660
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   309
            Left            =   1980
            TabIndex        =   570
            Top             =   3660
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   310
            Left            =   2220
            TabIndex        =   571
            Top             =   3660
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   311
            Left            =   2460
            TabIndex        =   572
            Top             =   3660
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   312
            Left            =   2700
            TabIndex        =   573
            Top             =   3660
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   313
            Left            =   2940
            TabIndex        =   574
            Top             =   3660
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   314
            Left            =   3180
            TabIndex        =   575
            Top             =   3660
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   315
            Left            =   3420
            TabIndex        =   576
            Top             =   3660
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   316
            Left            =   3660
            TabIndex        =   577
            Top             =   3660
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   317
            Left            =   3900
            TabIndex        =   578
            Top             =   3660
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   318
            Left            =   4140
            TabIndex        =   579
            Top             =   3660
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   319
            Left            =   4380
            TabIndex        =   580
            Top             =   3660
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   320
            Left            =   4620
            TabIndex        =   581
            Top             =   3660
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   321
            Left            =   60
            TabIndex        =   582
            Top             =   3900
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   322
            Left            =   300
            TabIndex        =   583
            Top             =   3900
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   323
            Left            =   540
            TabIndex        =   584
            Top             =   3900
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   324
            Left            =   780
            TabIndex        =   585
            Top             =   3900
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   325
            Left            =   1020
            TabIndex        =   586
            Top             =   3900
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   326
            Left            =   1260
            TabIndex        =   587
            Top             =   3900
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   327
            Left            =   1500
            TabIndex        =   588
            Top             =   3900
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   328
            Left            =   1740
            TabIndex        =   589
            Top             =   3900
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   329
            Left            =   1980
            TabIndex        =   590
            Top             =   3900
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   330
            Left            =   2220
            TabIndex        =   591
            Top             =   3900
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   331
            Left            =   2460
            TabIndex        =   592
            Top             =   3900
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   332
            Left            =   2700
            TabIndex        =   593
            Top             =   3900
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   333
            Left            =   2940
            TabIndex        =   594
            Top             =   3900
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   334
            Left            =   3180
            TabIndex        =   595
            Top             =   3900
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   335
            Left            =   3420
            TabIndex        =   596
            Top             =   3900
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   336
            Left            =   3660
            TabIndex        =   597
            Top             =   3900
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   337
            Left            =   3900
            TabIndex        =   598
            Top             =   3900
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   338
            Left            =   4140
            TabIndex        =   599
            Top             =   3900
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   339
            Left            =   4380
            TabIndex        =   600
            Top             =   3900
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   340
            Left            =   4620
            TabIndex        =   601
            Top             =   3900
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   341
            Left            =   60
            TabIndex        =   602
            Top             =   4140
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   342
            Left            =   300
            TabIndex        =   603
            Top             =   4140
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   343
            Left            =   540
            TabIndex        =   604
            Top             =   4140
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   344
            Left            =   780
            TabIndex        =   605
            Top             =   4140
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   345
            Left            =   1020
            TabIndex        =   606
            Top             =   4140
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   346
            Left            =   1260
            TabIndex        =   607
            Top             =   4140
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   347
            Left            =   1500
            TabIndex        =   608
            Top             =   4140
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   348
            Left            =   1740
            TabIndex        =   609
            Top             =   4140
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   349
            Left            =   1980
            TabIndex        =   610
            Top             =   4140
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   350
            Left            =   2220
            TabIndex        =   611
            Top             =   4140
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   351
            Left            =   2460
            TabIndex        =   612
            Top             =   4140
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   352
            Left            =   2700
            TabIndex        =   613
            Top             =   4140
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   353
            Left            =   2940
            TabIndex        =   614
            Top             =   4140
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   354
            Left            =   3180
            TabIndex        =   615
            Top             =   4140
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   355
            Left            =   3420
            TabIndex        =   616
            Top             =   4140
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   356
            Left            =   3660
            TabIndex        =   617
            Top             =   4140
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   357
            Left            =   3900
            TabIndex        =   618
            Top             =   4140
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   358
            Left            =   4140
            TabIndex        =   619
            Top             =   4140
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   359
            Left            =   4380
            TabIndex        =   620
            Top             =   4140
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   360
            Left            =   4620
            TabIndex        =   621
            Top             =   4140
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   361
            Left            =   60
            TabIndex        =   622
            Top             =   4380
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   362
            Left            =   300
            TabIndex        =   623
            Top             =   4380
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   363
            Left            =   540
            TabIndex        =   624
            Top             =   4380
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   364
            Left            =   780
            TabIndex        =   625
            Top             =   4380
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   365
            Left            =   1020
            TabIndex        =   626
            Top             =   4380
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   366
            Left            =   1260
            TabIndex        =   627
            Top             =   4380
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   367
            Left            =   1500
            TabIndex        =   628
            Top             =   4380
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   368
            Left            =   1740
            TabIndex        =   629
            Top             =   4380
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   369
            Left            =   1980
            TabIndex        =   630
            Top             =   4380
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   370
            Left            =   2220
            TabIndex        =   631
            Top             =   4380
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   371
            Left            =   2460
            TabIndex        =   632
            Top             =   4380
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   372
            Left            =   2700
            TabIndex        =   633
            Top             =   4380
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   373
            Left            =   2940
            TabIndex        =   634
            Top             =   4380
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   374
            Left            =   3180
            TabIndex        =   635
            Top             =   4380
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   375
            Left            =   3420
            TabIndex        =   636
            Top             =   4380
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   376
            Left            =   3660
            TabIndex        =   637
            Top             =   4380
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   377
            Left            =   3900
            TabIndex        =   638
            Top             =   4380
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   378
            Left            =   4140
            TabIndex        =   639
            Top             =   4380
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   379
            Left            =   4380
            TabIndex        =   640
            Top             =   4380
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   380
            Left            =   4620
            TabIndex        =   641
            Top             =   4380
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   381
            Left            =   60
            TabIndex        =   642
            Top             =   4620
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   382
            Left            =   300
            TabIndex        =   643
            Top             =   4620
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   383
            Left            =   540
            TabIndex        =   644
            Top             =   4620
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   384
            Left            =   780
            TabIndex        =   645
            Top             =   4620
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   385
            Left            =   1020
            TabIndex        =   646
            Top             =   4620
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   386
            Left            =   1260
            TabIndex        =   647
            Top             =   4620
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   387
            Left            =   1500
            TabIndex        =   648
            Top             =   4620
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   388
            Left            =   1740
            TabIndex        =   649
            Top             =   4620
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   389
            Left            =   1980
            TabIndex        =   650
            Top             =   4620
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   390
            Left            =   2220
            TabIndex        =   651
            Top             =   4620
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   391
            Left            =   2460
            TabIndex        =   652
            Top             =   4620
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   392
            Left            =   2700
            TabIndex        =   653
            Top             =   4620
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   393
            Left            =   2940
            TabIndex        =   654
            Top             =   4620
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   394
            Left            =   3180
            TabIndex        =   655
            Top             =   4620
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   395
            Left            =   3420
            TabIndex        =   656
            Top             =   4620
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   396
            Left            =   3660
            TabIndex        =   657
            Top             =   4620
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   397
            Left            =   3900
            TabIndex        =   658
            Top             =   4620
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   398
            Left            =   4140
            TabIndex        =   659
            Top             =   4620
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   399
            Left            =   4380
            TabIndex        =   660
            Top             =   4620
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblRoomCell 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   135
            Index           =   400
            Left            =   4620
            TabIndex        =   661
            Top             =   4620
            Visible         =   0   'False
            Width           =   135
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   4950
         TabIndex        =   209
         Top             =   240
         Width           =   2955
         Begin VB.CommandButton cmdMapShowUnused 
            Caption         =   "S&how Blocks"
            Height          =   315
            Left            =   1500
            TabIndex        =   227
            Top             =   3120
            Width           =   1275
         End
         Begin VB.CommandButton cmdMapOpenExternal 
            Caption         =   "E&xternal Map"
            Height          =   315
            Left            =   120
            TabIndex        =   226
            Top             =   3120
            Width           =   1395
         End
         Begin VB.CommandButton cmdMapLastRoom 
            Caption         =   "L&ast Room"
            Height          =   315
            Left            =   1500
            TabIndex        =   225
            ToolTipText     =   "Goes back one room"
            Top             =   2820
            Width           =   1275
         End
         Begin VB.CheckBox chkMapOptions 
            Caption         =   "Don't Show Tooltips"
            Height          =   195
            Index           =   5
            Left            =   480
            TabIndex        =   221
            Top             =   2250
            Width           =   1995
         End
         Begin VB.CommandButton cmdViewMapLegend 
            Caption         =   "Help/&Legend"
            Height          =   315
            Left            =   120
            TabIndex        =   224
            Top             =   2820
            Width           =   1395
         End
         Begin VB.CommandButton cmdMapFindText 
            Caption         =   "Find &Next"
            Height          =   315
            Index           =   1
            Left            =   1500
            TabIndex        =   223
            Top             =   2520
            Width           =   1275
         End
         Begin VB.CommandButton cmdMapFindText 
            Caption         =   "&Find Room"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   222
            Top             =   2520
            Width           =   1395
         End
         Begin VB.CommandButton cmdDrawMap 
            Caption         =   "&Redraw"
            Height          =   315
            Index           =   1
            Left            =   1440
            TabIndex        =   215
            Top             =   600
            Width           =   1035
         End
         Begin VB.CheckBox chkMapOptions 
            Caption         =   "Don't Mark Commands"
            Height          =   195
            Index           =   4
            Left            =   480
            TabIndex        =   220
            Top             =   2010
            Width           =   2115
         End
         Begin VB.CheckBox chkMapOptions 
            Caption         =   "Don't Mark NPCs"
            Height          =   195
            Index           =   3
            Left            =   480
            TabIndex        =   219
            Top             =   1755
            Width           =   1995
         End
         Begin VB.CheckBox chkMapOptions 
            Caption         =   "Don't Mark Lairs"
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   218
            Top             =   1500
            Width           =   1995
         End
         Begin VB.TextBox txtRoomMap 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   420
            MaxLength       =   5
            TabIndex        =   212
            Text            =   "1"
            Top             =   300
            Width           =   1035
         End
         Begin VB.TextBox txtRoomRoom 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1440
            MaxLength       =   5
            TabIndex        =   213
            Text            =   "1"
            Top             =   300
            Width           =   1035
         End
         Begin VB.CommandButton cmdDrawMap 
            Caption         =   "&Draw"
            Height          =   315
            Index           =   0
            Left            =   420
            TabIndex        =   214
            Top             =   600
            Width           =   1035
         End
         Begin VB.CheckBox chkMapOptions 
            Caption         =   "Follow Map Changes"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   216
            Top             =   990
            Width           =   1875
         End
         Begin VB.CheckBox chkMapOptions 
            Caption         =   "Don't Follow Hidden"
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   217
            Top             =   1260
            Width           =   1995
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   2  'Center
            Caption         =   "Map"
            Height          =   195
            Index           =   19
            Left            =   420
            TabIndex        =   210
            Top             =   120
            Width           =   1035
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   2  'Center
            Caption         =   "Room"
            Height          =   195
            Index           =   20
            Left            =   1440
            TabIndex        =   211
            Top             =   120
            Width           =   1035
         End
      End
      Begin MSComctlLib.ListView lvMapLoc 
         Height          =   1455
         Left            =   4950
         TabIndex        =   228
         Top             =   3720
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label lblLabelArray 
         Caption         =   "Place your cursor in this box to the left and move around the map with your keypad."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   21
         Left            =   5040
         TabIndex        =   664
         Top             =   5220
         Width           =   5235
      End
      Begin VB.Label lblMapBG 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   4845
         Left            =   120
         TabIndex        =   662
         Top             =   330
         Width           =   4845
      End
      Begin VB.Label Label41 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   120
         TabIndex        =   259
         Top             =   240
         Width           =   4875
      End
   End
   Begin VB.Frame framNav 
      Caption         =   "Shops"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6075
      Index           =   9
      Left            =   60
      TabIndex        =   84
      Top             =   1200
      Visible         =   0   'False
      Width           =   10455
      Begin VB.CommandButton cmdShopConvert 
         Caption         =   "$"
         Height          =   255
         Left            =   10080
         TabIndex        =   96
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkShopShowCharm 
         Caption         =   "Show Selling"
         Height          =   255
         Index           =   1
         Left            =   8640
         TabIndex        =   95
         Top             =   240
         Width           =   1275
      End
      Begin VB.CheckBox chkShopShowCharm 
         Caption         =   "Show Buying"
         Height          =   255
         Index           =   0
         Left            =   7140
         TabIndex        =   94
         Top             =   240
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.CommandButton cmdShopAlterCharm 
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
         Height          =   255
         Index           =   1
         Left            =   6660
         TabIndex        =   93
         Top             =   240
         Width           =   315
      End
      Begin VB.CommandButton cmdShopAlterCharm 
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
         Height          =   255
         Index           =   0
         Left            =   6360
         TabIndex        =   92
         Top             =   240
         Width           =   315
      End
      Begin VB.TextBox txtShopCharm 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5460
         MaxLength       =   4
         TabIndex        =   91
         Text            =   "0"
         Top             =   240
         Width           =   795
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Next"
         Height          =   255
         Index           =   7
         Left            =   3330
         TabIndex        =   87
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   255
         Index           =   6
         Left            =   2190
         TabIndex        =   86
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtShopFind 
         Height          =   285
         Left            =   120
         TabIndex        =   85
         Top             =   240
         Width           =   1995
      End
      Begin VB.TextBox txtShopDetail 
         Height          =   435
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   89
         Top             =   4440
         Width           =   4335
      End
      Begin MSComctlLib.ListView lvShops 
         Height          =   3855
         Left            =   120
         TabIndex        =   88
         Tag             =   "STRETCHALL"
         Top             =   540
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   6800
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvShopDetail 
         Height          =   5415
         Left            =   4500
         TabIndex        =   97
         Tag             =   "STRETCHALL"
         Top             =   540
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   9551
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvShopLoc 
         Height          =   1035
         Left            =   120
         TabIndex        =   90
         Top             =   4920
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1826
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblLabelArray 
         AutoSize        =   -1  'True
         Caption         =   "Charm:"
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
         Index           =   22
         Left            =   4575
         TabIndex        =   734
         Top             =   255
         Width           =   735
      End
   End
   Begin VB.Frame framNav 
      Caption         =   "Monsters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6075
      Index           =   8
      Left            =   60
      TabIndex        =   83
      Top             =   1200
      Visible         =   0   'False
      Width           =   10455
      Begin VB.CheckBox chkMonstersNoRegenLookUp 
         Caption         =   "Don't Lookup Monster Regen in Detail Window"
         Height          =   255
         Left            =   5460
         TabIndex        =   758
         Top             =   240
         Width           =   4095
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Next"
         Height          =   255
         Index           =   9
         Left            =   3570
         TabIndex        =   756
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   255
         Index           =   8
         Left            =   2250
         TabIndex        =   755
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtMonsterFind 
         Height          =   285
         Left            =   120
         TabIndex        =   754
         Top             =   240
         Width           =   1995
      End
      Begin VB.CommandButton cmdMonHelp 
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
         Height          =   255
         Left            =   5040
         TabIndex        =   757
         Top             =   240
         Width           =   255
      End
      Begin MMUD_Explorer.cntSplitter splMonsterSplit 
         Height          =   5415
         Index           =   0
         Left            =   120
         TabIndex        =   753
         Top             =   540
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   9551
         Begin MSComctlLib.ListView lvMonsterDetail 
            Height          =   4995
            Left            =   5580
            TabIndex        =   761
            Tag             =   "STRETCHALL"
            Top             =   0
            Width           =   4275
            _ExtentX        =   7541
            _ExtentY        =   8811
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ListView lvMonsters 
            Height          =   4995
            Left            =   0
            TabIndex        =   760
            Tag             =   "STRETCHALL"
            Top             =   0
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   8811
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
   End
   Begin VB.Frame framNav 
      Caption         =   "Sundry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6075
      Index           =   7
      Left            =   60
      TabIndex        =   99
      Top             =   1200
      Visible         =   0   'False
      Width           =   10455
      Begin VB.CommandButton cmdSundryChests 
         Caption         =   "View Chest Contents"
         Height          =   255
         Left            =   8220
         TabIndex        =   103
         Top             =   240
         Width           =   2115
      End
      Begin MMUD_Explorer.cntSplitter splSplitterWE 
         Height          =   5415
         Index           =   6
         Left            =   120
         TabIndex        =   750
         Top             =   540
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   9551
         Begin MSComctlLib.ListView lvOtherItems 
            Height          =   3735
            Left            =   0
            TabIndex        =   104
            Tag             =   "STRETCHALL"
            Top             =   0
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   6588
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MMUD_Explorer.cntSplitter splSplitterNS 
            Height          =   5415
            Index           =   6
            Left            =   6360
            TabIndex        =   751
            Top             =   0
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   9551
            Begin VB.TextBox txtOtherItemDetail 
               Height          =   1875
               Left            =   0
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   105
               Top             =   0
               Width           =   2775
            End
            Begin MSComctlLib.ListView lvOtherItemLoc 
               Height          =   1575
               Left            =   0
               TabIndex        =   106
               Top             =   2880
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   2778
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
         End
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Next"
         Height          =   255
         Index           =   11
         Left            =   3780
         TabIndex        =   102
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   255
         Index           =   10
         Left            =   2280
         TabIndex        =   101
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtOtherItemsFind 
         Height          =   285
         Left            =   120
         TabIndex        =   100
         Top             =   240
         Width           =   1995
      End
   End
   Begin VB.Frame framNav 
      Caption         =   "Classes / Races"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6075
      Index           =   6
      Left            =   60
      TabIndex        =   98
      Top             =   1200
      Visible         =   0   'False
      Width           =   10455
      Begin VB.CommandButton cmdClassRace 
         Caption         =   "&Races"
         Height          =   315
         Index           =   1
         Left            =   1500
         Style           =   1  'Graphical
         TabIndex        =   731
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdClassRace 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Classes"
         Height          =   315
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   730
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtRaceDetail 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   733
         Top             =   5580
         Visible         =   0   'False
         Width           =   10215
      End
      Begin MSComctlLib.ListView lvClasses 
         Height          =   4935
         Left            =   120
         TabIndex        =   732
         Tag             =   "STRETCHALL"
         Top             =   600
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   8705
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtClassDetail 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   736
         Top             =   5580
         Width           =   10215
      End
      Begin MSComctlLib.ListView lvRaces 
         Height          =   4935
         Left            =   120
         TabIndex        =   735
         Tag             =   "STRETCHALL"
         Top             =   600
         Visible         =   0   'False
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   8705
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame framNav 
      Caption         =   "Character"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6075
      Index           =   5
      Left            =   60
      TabIndex        =   762
      Top             =   1200
      Visible         =   0   'False
      Width           =   10455
      Begin VB.CommandButton cmdInvenPasteChar 
         Caption         =   "&Paste Character..."
         Height          =   435
         Index           =   1
         Left            =   180
         TabIndex        =   799
         Top             =   5520
         Width           =   1755
      End
      Begin VB.Frame fraChar 
         Caption         =   "Stats"
         Height          =   5175
         Index           =   0
         Left            =   120
         TabIndex        =   840
         Top             =   240
         Width           =   3915
         Begin VB.CommandButton cmdCharChangeStats 
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
            Height          =   255
            Index           =   11
            Left            =   3420
            TabIndex        =   789
            Top             =   4020
            Width           =   315
         End
         Begin VB.CommandButton cmdCharChangeStats 
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
            Height          =   255
            Index           =   9
            Left            =   3420
            TabIndex        =   786
            Top             =   3660
            Width           =   315
         End
         Begin VB.CommandButton cmdCharChangeStats 
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
            Height          =   255
            Index           =   7
            Left            =   3420
            TabIndex        =   783
            Top             =   3300
            Width           =   315
         End
         Begin VB.CommandButton cmdCharChangeStats 
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
            Height          =   255
            Index           =   5
            Left            =   3420
            TabIndex        =   780
            Top             =   2940
            Width           =   315
         End
         Begin VB.CommandButton cmdCharChangeStats 
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
            Height          =   255
            Index           =   3
            Left            =   3420
            TabIndex        =   777
            Top             =   2580
            Width           =   315
         End
         Begin VB.CommandButton cmdCharChangeStats 
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
            Height          =   255
            Index           =   1
            Left            =   3420
            TabIndex        =   774
            Top             =   2220
            Width           =   315
         End
         Begin VB.CommandButton cmdCharChangeStats 
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
            Height          =   255
            Index           =   13
            Left            =   3420
            TabIndex        =   769
            Top             =   1380
            Width           =   315
         End
         Begin VB.TextBox txtCharStats 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   1380
            MaxLength       =   4
            TabIndex        =   787
            Text            =   "0"
            Top             =   4020
            Width           =   735
         End
         Begin VB.TextBox txtCharStats 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   1380
            MaxLength       =   4
            TabIndex        =   781
            Text            =   "0"
            Top             =   3300
            Width           =   735
         End
         Begin VB.TextBox txtGlobalLevel 
            Alignment       =   2  'Center
            Height          =   300
            Index           =   0
            Left            =   1380
            MaxLength       =   3
            TabIndex        =   767
            Text            =   "999"
            Top             =   1380
            Width           =   1635
         End
         Begin VB.ComboBox cmbGlobalClass 
            Height          =   315
            Index           =   0
            Left            =   1380
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   765
            Top             =   660
            Width           =   1635
         End
         Begin VB.TextBox txtCharStats 
            Alignment       =   2  'Center
            Height          =   300
            Index           =   0
            Left            =   1380
            MaxLength       =   4
            TabIndex        =   772
            Text            =   "100"
            Top             =   2220
            Width           =   735
         End
         Begin VB.ComboBox cmbGlobalRace 
            Height          =   315
            Index           =   0
            Left            =   1380
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   766
            Top             =   1020
            Width           =   1635
         End
         Begin VB.TextBox txtCharStats 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   1380
            MaxLength       =   4
            TabIndex        =   778
            Text            =   "0"
            Top             =   2940
            Width           =   735
         End
         Begin VB.TextBox txtCharStats 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   1380
            MaxLength       =   4
            TabIndex        =   775
            Text            =   "0"
            Top             =   2580
            Width           =   735
         End
         Begin VB.TextBox txtCharStats 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   1380
            MaxLength       =   4
            TabIndex        =   784
            Text            =   "0"
            Top             =   3660
            Width           =   735
         End
         Begin VB.TextBox txtCharMaxStats 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            Height          =   285
            Index           =   0
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   846
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   2220
            Width           =   855
         End
         Begin VB.TextBox txtCharMaxStats 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            Height          =   285
            Index           =   1
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   845
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   2580
            Width           =   855
         End
         Begin VB.TextBox txtCharMaxStats 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            Height          =   285
            Index           =   2
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   844
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   2940
            Width           =   855
         End
         Begin VB.TextBox txtCharMaxStats 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            Height          =   285
            Index           =   3
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   843
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   3300
            Width           =   855
         End
         Begin VB.TextBox txtCharMaxStats 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            Height          =   285
            Index           =   4
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   842
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   3660
            Width           =   855
         End
         Begin VB.TextBox txtCharMaxStats 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            Height          =   285
            Index           =   5
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   841
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   4020
            Width           =   855
         End
         Begin VB.TextBox txtCharName 
            Height          =   285
            Left            =   1380
            MaxLength       =   32
            TabIndex        =   764
            Top             =   300
            Width           =   2355
         End
         Begin VB.CommandButton cmdCharChangeStats 
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
            Height          =   255
            Index           =   0
            Left            =   3120
            TabIndex        =   773
            Top             =   2220
            Width           =   315
         End
         Begin VB.CommandButton cmdCharChangeStats 
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
            Height          =   255
            Index           =   2
            Left            =   3120
            TabIndex        =   776
            Top             =   2580
            Width           =   315
         End
         Begin VB.CommandButton cmdCharChangeStats 
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
            Height          =   255
            Index           =   4
            Left            =   3120
            TabIndex        =   779
            Top             =   2940
            Width           =   315
         End
         Begin VB.CommandButton cmdCharChangeStats 
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
            Height          =   255
            Index           =   6
            Left            =   3120
            TabIndex        =   782
            Top             =   3300
            Width           =   315
         End
         Begin VB.CommandButton cmdCharChangeStats 
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
            Height          =   255
            Index           =   8
            Left            =   3120
            TabIndex        =   785
            Top             =   3660
            Width           =   315
         End
         Begin VB.CommandButton cmdCharChangeStats 
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
            Height          =   255
            Index           =   10
            Left            =   3120
            TabIndex        =   788
            Top             =   4020
            Width           =   315
         End
         Begin VB.CommandButton cmdCharChangeStats 
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
            Height          =   315
            Index           =   14
            Left            =   2160
            TabIndex        =   771
            Top             =   1800
            Width           =   855
         End
         Begin VB.CommandButton cmdCharChangeStats 
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
            Height          =   255
            Index           =   12
            Left            =   3120
            TabIndex        =   768
            Top             =   1380
            Width           =   315
         End
         Begin VB.CommandButton cmdCharChangeStats 
            Caption         =   "Reload"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   15
            Left            =   1380
            TabIndex        =   770
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label lblLabelArray 
            AutoSize        =   -1  'True
            Caption         =   "Charm:"
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
            Left            =   180
            TabIndex        =   857
            Top             =   4020
            Width           =   735
         End
         Begin VB.Label lblLabelArray 
            AutoSize        =   -1  'True
            Caption         =   "Agility:"
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
            Left            =   180
            TabIndex        =   856
            Top             =   3300
            Width           =   720
         End
         Begin VB.Label lblLabelArray 
            AutoSize        =   -1  'True
            Caption         =   "Level:"
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
            TabIndex        =   855
            Top             =   1380
            Width           =   645
         End
         Begin VB.Label lblLabelArray 
            AutoSize        =   -1  'True
            Caption         =   "Class:"
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
            Index           =   6
            Left            =   180
            TabIndex        =   854
            Top             =   660
            Width           =   660
         End
         Begin VB.Label lblLabelArray 
            AutoSize        =   -1  'True
            Caption         =   "Strength:"
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
            Left            =   180
            TabIndex        =   853
            Top             =   2220
            Width           =   930
         End
         Begin VB.Label lblLabelArray 
            AutoSize        =   -1  'True
            Caption         =   "Race:"
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
            Left            =   180
            TabIndex        =   852
            Top             =   1020
            Width           =   630
         End
         Begin VB.Label lblLabelArray 
            AutoSize        =   -1  'True
            Caption         =   "Intellect:"
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
            Left            =   180
            TabIndex        =   851
            Top             =   2580
            Width           =   885
         End
         Begin VB.Label lblLabelArray 
            AutoSize        =   -1  'True
            Caption         =   "Health:"
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
            Index           =   23
            Left            =   180
            TabIndex        =   850
            Top             =   3660
            Width           =   750
         End
         Begin VB.Label lblLabelArray 
            AutoSize        =   -1  'True
            Caption         =   "Willpower:"
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
            Index           =   24
            Left            =   180
            TabIndex        =   849
            Top             =   2940
            Width           =   1095
         End
         Begin VB.Label lblLabelArray 
            AutoSize        =   -1  'True
            Caption         =   "Name:"
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
            Index           =   25
            Left            =   180
            TabIndex        =   848
            Top             =   300
            Width           =   690
         End
         Begin VB.Label lblStatCalc 
            Appearance      =   0  'Flat
            Caption         =   "CP:"
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
            Height          =   615
            Left            =   180
            TabIndex        =   847
            Top             =   4440
            Width           =   3555
         End
      End
      Begin VB.Frame fraChar 
         Caption         =   "Hit Points / Hit Point Regen"
         Height          =   1335
         Index           =   1
         Left            =   4140
         TabIndex        =   836
         Top             =   240
         Width           =   4095
         Begin VB.TextBox txtCharHPRegen 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   180
            MaxLength       =   4
            TabIndex        =   790
            Top             =   840
            Width           =   1035
         End
         Begin VB.Label lblLabelArray 
            AutoSize        =   -1  'True
            Caption         =   "Regen Bonus"
            Height          =   195
            Index           =   26
            Left            =   180
            TabIndex        =   839
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lblCharRestRate 
            Appearance      =   0  'Flat
            Caption         =   "Resting:"
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
            Height          =   495
            Left            =   1380
            TabIndex        =   838
            Top             =   660
            Width           =   2535
         End
         Begin VB.Label lblCharMaxHP 
            Appearance      =   0  'Flat
            Caption         =   "HP Range:"
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
            Height          =   255
            Left            =   180
            TabIndex        =   837
            Top             =   300
            Width           =   3735
         End
      End
      Begin VB.Frame fraChar 
         Caption         =   "Spellcasting / Mana / Mana Regen"
         Height          =   1635
         Index           =   2
         Left            =   4140
         TabIndex        =   831
         Top             =   1620
         Width           =   4095
         Begin VB.TextBox txtCharManaRegen 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   180
            MaxLength       =   4
            TabIndex        =   791
            Top             =   1200
            Width           =   1035
         End
         Begin VB.Label lblCharManaRate 
            Appearance      =   0  'Flat
            Caption         =   "Mana Regen:"
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
            Height          =   495
            Left            =   1380
            TabIndex        =   835
            Top             =   1020
            Width           =   2535
         End
         Begin VB.Label lblLabelArray 
            AutoSize        =   -1  'True
            Caption         =   "Regen Bonus"
            Height          =   195
            Index           =   27
            Left            =   180
            TabIndex        =   834
            Top             =   960
            Width           =   975
         End
         Begin VB.Label lblCharMaxMana 
            Appearance      =   0  'Flat
            Caption         =   "Mana:"
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
            Height          =   255
            Left            =   180
            TabIndex        =   833
            Top             =   600
            Width           =   3735
         End
         Begin VB.Label lblCharSC 
            Appearance      =   0  'Flat
            Caption         =   "Spellcasting:"
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
            Height          =   255
            Left            =   180
            TabIndex        =   832
            Top             =   300
            Width           =   3735
         End
      End
      Begin VB.Frame fraChar 
         Caption         =   "Completed Quests"
         Height          =   1455
         Index           =   4
         Left            =   4140
         TabIndex        =   830
         Top             =   3960
         Width           =   4095
         Begin VB.ComboBox cmbChar2ndAlign 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2460
            Style           =   2  'Dropdown List
            TabIndex        =   798
            Top             =   1020
            Width           =   1455
         End
         Begin VB.CheckBox chkCharQuests 
            Height          =   255
            Index           =   5
            Left            =   2160
            TabIndex        =   797
            Top             =   1020
            Width           =   195
         End
         Begin VB.CheckBox chkCharQuests 
            Caption         =   "Ice Sorceress (+1 AC)"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   792
            Top             =   300
            Width           =   1935
         End
         Begin VB.CheckBox chkCharQuests 
            Caption         =   "High Druid (+1 SC)"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   793
            Top             =   660
            Width           =   1935
         End
         Begin VB.CheckBox chkCharQuests 
            Caption         =   "Bishop (+3 ACC)"
            Height          =   255
            Index           =   3
            Left            =   2160
            TabIndex        =   795
            Top             =   300
            Width           =   1635
         End
         Begin VB.CheckBox chkCharQuests 
            Caption         =   "ARD (+1 Crit, +2 SC)"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   794
            Top             =   1020
            Width           =   1815
         End
         Begin VB.CheckBox chkCharQuests 
            Caption         =   "Apparatus (+1 DG)"
            Height          =   255
            Index           =   4
            Left            =   2160
            TabIndex        =   796
            Top             =   660
            Width           =   1635
         End
      End
      Begin VB.Frame fraChar 
         Caption         =   "Picklocks"
         Height          =   615
         Index           =   3
         Left            =   4140
         TabIndex        =   828
         Top             =   3300
         Width           =   4095
         Begin VB.Label lblCharPicklocks 
            Appearance      =   0  'Flat
            Caption         =   "Picklocks: "
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
            Height          =   255
            Left            =   180
            TabIndex        =   829
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.CommandButton cmdCharCopytoClip 
         Caption         =   "&Copy Full Character"
         Height          =   435
         Index           =   0
         Left            =   6360
         TabIndex        =   802
         Top             =   5520
         Width           =   1755
      End
      Begin VB.CommandButton cmdCharCopytoClip 
         Caption         =   "Copy Only Stats"
         Height          =   435
         Index           =   1
         Left            =   4200
         TabIndex        =   801
         Top             =   5520
         Width           =   1755
      End
      Begin VB.CommandButton cmdCharCopytoClip 
         Caption         =   "Copy Only CP Info"
         Height          =   435
         Index           =   2
         Left            =   2160
         TabIndex        =   800
         Top             =   5520
         Width           =   1755
      End
      Begin VB.CommandButton cmdCharReset 
         Caption         =   "Reset Character Fields"
         Height          =   435
         Left            =   8340
         TabIndex        =   803
         Top             =   5520
         Width           =   1995
      End
      Begin VB.Frame fraChar 
         Caption         =   "Bless Calc (Only)"
         Height          =   4755
         Index           =   5
         Left            =   8340
         TabIndex        =   763
         Top             =   420
         Width           =   1995
         Begin VB.ComboBox cmbCharBless 
            Height          =   315
            Index           =   0
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   806
            Text            =   "cmbCharBless"
            Top             =   480
            Width           =   1635
         End
         Begin VB.ComboBox cmbCharBless 
            Height          =   315
            Index           =   1
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   808
            Text            =   "cmbCharBless"
            Top             =   840
            Width           =   1635
         End
         Begin VB.ComboBox cmbCharBless 
            Height          =   315
            Index           =   2
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   810
            Text            =   "cmbCharBless"
            Top             =   1200
            Width           =   1635
         End
         Begin VB.ComboBox cmbCharBless 
            Height          =   315
            Index           =   3
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   812
            Text            =   "cmbCharBless"
            Top             =   1560
            Width           =   1635
         End
         Begin VB.ComboBox cmbCharBless 
            Height          =   315
            Index           =   4
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   814
            Text            =   "cmbCharBless"
            Top             =   1920
            Width           =   1635
         End
         Begin VB.ComboBox cmbCharBless 
            Height          =   315
            Index           =   5
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   816
            Text            =   "cmbCharBless"
            Top             =   2280
            Width           =   1635
         End
         Begin VB.ComboBox cmbCharBless 
            Height          =   315
            Index           =   6
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   818
            Text            =   "cmbCharBless"
            Top             =   2640
            Width           =   1635
         End
         Begin VB.ComboBox cmbCharBless 
            Height          =   315
            Index           =   7
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   820
            Text            =   "cmbCharBless"
            Top             =   3000
            Width           =   1635
         End
         Begin VB.ComboBox cmbCharBless 
            Height          =   315
            Index           =   8
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   822
            Text            =   "cmbCharBless"
            Top             =   3360
            Width           =   1635
         End
         Begin VB.ComboBox cmbCharBless 
            Height          =   315
            Index           =   9
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   824
            Text            =   "cmbCharBless"
            Top             =   3720
            Width           =   1635
         End
         Begin VB.CommandButton cmdCharResetBless 
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
            Height          =   195
            Index           =   0
            Left            =   1140
            TabIndex        =   805
            Top             =   240
            Width           =   795
         End
         Begin VB.CommandButton cmdCharResetBless 
            Caption         =   "Reload"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   804
            Top             =   240
            Width           =   915
         End
         Begin VB.CommandButton cmdCharBlessJump 
            Caption         =   ">"
            Height          =   315
            Index           =   0
            Left            =   1800
            TabIndex        =   807
            Top             =   480
            Width           =   135
         End
         Begin VB.CommandButton cmdCharBlessJump 
            Caption         =   ">"
            Height          =   315
            Index           =   1
            Left            =   1800
            TabIndex        =   809
            Top             =   840
            Width           =   135
         End
         Begin VB.CommandButton cmdCharBlessJump 
            Caption         =   ">"
            Height          =   315
            Index           =   2
            Left            =   1800
            TabIndex        =   811
            Top             =   1200
            Width           =   135
         End
         Begin VB.CommandButton cmdCharBlessJump 
            Caption         =   ">"
            Height          =   315
            Index           =   3
            Left            =   1800
            TabIndex        =   813
            Top             =   1560
            Width           =   135
         End
         Begin VB.CommandButton cmdCharBlessJump 
            Caption         =   ">"
            Height          =   315
            Index           =   4
            Left            =   1800
            TabIndex        =   815
            Top             =   1920
            Width           =   135
         End
         Begin VB.CommandButton cmdCharBlessJump 
            Caption         =   ">"
            Height          =   315
            Index           =   5
            Left            =   1800
            TabIndex        =   817
            Top             =   2280
            Width           =   135
         End
         Begin VB.CommandButton cmdCharBlessJump 
            Caption         =   ">"
            Height          =   315
            Index           =   6
            Left            =   1800
            TabIndex        =   819
            Top             =   2640
            Width           =   135
         End
         Begin VB.CommandButton cmdCharBlessJump 
            Caption         =   ">"
            Height          =   315
            Index           =   7
            Left            =   1800
            TabIndex        =   821
            Top             =   3000
            Width           =   135
         End
         Begin VB.CommandButton cmdCharBlessJump 
            Caption         =   ">"
            Height          =   315
            Index           =   8
            Left            =   1800
            TabIndex        =   823
            Top             =   3360
            Width           =   135
         End
         Begin VB.CommandButton cmdCharBlessJump 
            Caption         =   ">"
            Height          =   315
            Index           =   9
            Left            =   1800
            TabIndex        =   825
            Top             =   3720
            Width           =   135
         End
         Begin VB.Label lblCharBless 
            Alignment       =   2  'Center
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   827
            Top             =   4080
            Width           =   1695
         End
         Begin VB.Label lblLabelArray 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Mana Regen Needed"
            ForeColor       =   &H80000011&
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   826
            Top             =   4500
            Width           =   1785
         End
      End
   End
   Begin VB.Frame framNav 
      Caption         =   "Compare"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6075
      Index           =   3
      Left            =   60
      TabIndex        =   107
      Top             =   1200
      Visible         =   0   'False
      Width           =   10455
      Begin VB.CommandButton cmdCompareNav 
         Caption         =   "&Monsters"
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
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCompareClear 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   6420
         TabIndex        =   115
         ToolTipText     =   "Clear Spell List"
         Top             =   300
         Width           =   255
      End
      Begin VB.CommandButton cmdCompareQ 
         Caption         =   "?"
         Height          =   315
         Left            =   10020
         TabIndex        =   117
         Top             =   240
         Width           =   315
      End
      Begin VB.CommandButton cmdClearAllCompares 
         Caption         =   "&Clear All"
         Height          =   315
         Left            =   8580
         TabIndex        =   116
         Top             =   240
         Width           =   1395
      End
      Begin VB.CommandButton cmdCompareClear 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   4620
         TabIndex        =   113
         ToolTipText     =   "Clear Spell List"
         Top             =   300
         Width           =   255
      End
      Begin VB.CommandButton cmdCompareClear 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2940
         TabIndex        =   111
         ToolTipText     =   "Clear Armour List"
         Top             =   300
         Width           =   255
      End
      Begin VB.CommandButton cmdCompareClear 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1260
         TabIndex        =   109
         ToolTipText     =   "Clear Weapon List"
         Top             =   300
         Width           =   255
      End
      Begin VB.CommandButton cmdCompareNav 
         Caption         =   "&Spells"
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
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdCompareNav 
         Caption         =   "&Armour"
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
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdCompareNav 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Weapons"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   240
         Width           =   1095
      End
      Begin VB.Frame framCompareNav 
         Caption         =   "Weapons"
         Height          =   5355
         Index           =   0
         Left            =   120
         TabIndex        =   129
         Top             =   600
         Width           =   10215
         Begin MMUD_Explorer.cntSplitter splSplitterNS 
            Height          =   4995
            Index           =   3
            Left            =   120
            TabIndex        =   744
            Top             =   240
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   8811
            Begin MSComctlLib.ListView lvWeaponCompare 
               Height          =   2595
               Left            =   0
               TabIndex        =   118
               Tag             =   "STRETCHALL"
               Top             =   0
               Width           =   9975
               _ExtentX        =   17595
               _ExtentY        =   4577
               View            =   3
               LabelEdit       =   1
               Sorted          =   -1  'True
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
            Begin MMUD_Explorer.cntSplitter splSplitterWE 
               Height          =   975
               Index           =   3
               Left            =   0
               TabIndex        =   745
               Top             =   4020
               Width           =   9975
               _ExtentX        =   17595
               _ExtentY        =   1720
               Begin VB.TextBox txtWeaponCompareDetail 
                  Height          =   975
                  Left            =   0
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   119
                  Top             =   0
                  Width           =   4275
               End
               Begin MSComctlLib.ListView lvWeaponCompareLoc 
                  Height          =   975
                  Left            =   6180
                  TabIndex        =   120
                  Top             =   0
                  Width           =   3795
                  _ExtentX        =   6694
                  _ExtentY        =   1720
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  FullRowSelect   =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   1
                  NumItems        =   0
               End
            End
         End
      End
      Begin VB.Frame framCompareNav 
         Caption         =   "Spells"
         Height          =   5355
         Index           =   2
         Left            =   120
         TabIndex        =   131
         Top             =   600
         Width           =   10215
         Begin MMUD_Explorer.cntSplitter splSplitterNS 
            Height          =   4995
            Index           =   5
            Left            =   120
            TabIndex        =   748
            Top             =   240
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   8811
            Begin MSComctlLib.ListView lvSpellCompare 
               Height          =   2355
               Left            =   0
               TabIndex        =   121
               Tag             =   "STRETCHALL"
               Top             =   0
               Width           =   9975
               _ExtentX        =   17595
               _ExtentY        =   4154
               View            =   3
               LabelEdit       =   1
               Sorted          =   -1  'True
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
            Begin MMUD_Explorer.cntSplitter splSplitterWE 
               Height          =   1155
               Index           =   5
               Left            =   0
               TabIndex        =   749
               Top             =   3840
               Width           =   9975
               _ExtentX        =   17595
               _ExtentY        =   2037
               Begin VB.TextBox txtSpellCompareDetail 
                  Height          =   1155
                  Left            =   0
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   122
                  Top             =   0
                  Width           =   3375
               End
               Begin MSComctlLib.ListView lvSpellCompareLoc 
                  Height          =   1155
                  Left            =   7020
                  TabIndex        =   123
                  Top             =   0
                  Width           =   2955
                  _ExtentX        =   5212
                  _ExtentY        =   2037
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  FullRowSelect   =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   1
                  NumItems        =   0
               End
            End
         End
      End
      Begin VB.Frame framCompareNav 
         Caption         =   "Armour"
         Height          =   5355
         Index           =   1
         Left            =   120
         TabIndex        =   130
         Top             =   600
         Width           =   10215
         Begin MMUD_Explorer.cntSplitter splSplitterNS 
            Height          =   4995
            Index           =   4
            Left            =   120
            TabIndex        =   746
            Top             =   240
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   8811
            Begin MSComctlLib.ListView lvArmourCompare 
               Height          =   2895
               Left            =   0
               TabIndex        =   124
               Tag             =   "STRETCHALL"
               Top             =   0
               Width           =   9975
               _ExtentX        =   17595
               _ExtentY        =   5106
               View            =   3
               LabelEdit       =   1
               Sorted          =   -1  'True
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
            Begin MMUD_Explorer.cntSplitter splSplitterWE 
               Height          =   975
               Index           =   4
               Left            =   0
               TabIndex        =   747
               Top             =   4020
               Width           =   9975
               _ExtentX        =   17595
               _ExtentY        =   1720
               Begin VB.TextBox txtArmourCompareDetail 
                  Height          =   975
                  Left            =   0
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   125
                  Top             =   0
                  Width           =   3315
               End
               Begin MSComctlLib.ListView lvArmourCompareLoc 
                  Height          =   975
                  Left            =   7080
                  TabIndex        =   126
                  Top             =   0
                  Width           =   2895
                  _ExtentX        =   5106
                  _ExtentY        =   1720
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  FullRowSelect   =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   1
                  NumItems        =   0
               End
            End
         End
      End
      Begin VB.Frame framCompareNav 
         Caption         =   "Monsters"
         Height          =   5355
         Index           =   3
         Left            =   120
         TabIndex        =   752
         Top             =   600
         Width           =   10215
         Begin MMUD_Explorer.cntSplitter splMonsterSplit 
            Height          =   4995
            Index           =   1
            Left            =   120
            TabIndex        =   759
            Top             =   240
            Width           =   9915
            _ExtentX        =   17489
            _ExtentY        =   8811
            Begin MSComctlLib.ListView lvMonsterCompare 
               Height          =   4515
               Left            =   0
               TabIndex        =   127
               Tag             =   "STRETCHALL"
               Top             =   0
               Width           =   5115
               _ExtentX        =   9022
               _ExtentY        =   7964
               View            =   3
               LabelEdit       =   1
               Sorted          =   -1  'True
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
            Begin MSComctlLib.ListView lvMonsterCompareLoc 
               Height          =   4575
               Left            =   5640
               TabIndex        =   128
               Top             =   0
               Width           =   3915
               _ExtentX        =   6906
               _ExtentY        =   8070
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
         End
      End
   End
   Begin VB.Frame framNav 
      Caption         =   "Spells"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6075
      Index           =   2
      Left            =   60
      TabIndex        =   64
      Top             =   1200
      Visible         =   0   'False
      Width           =   10455
      Begin MMUD_Explorer.cntSplitter splSplitterNS 
         Height          =   5115
         Index           =   2
         Left            =   120
         TabIndex        =   742
         Top             =   840
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   9022
         Begin MSComctlLib.ListView lvSpells 
            Height          =   2535
            Left            =   0
            TabIndex        =   76
            Tag             =   "STRETCHALL"
            Top             =   0
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   4471
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MMUD_Explorer.cntSplitter splSplitterWE 
            Height          =   1155
            Index           =   2
            Left            =   0
            TabIndex        =   743
            Top             =   3960
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   2037
            Begin VB.TextBox txtSpellDetail 
               Height          =   1155
               Left            =   0
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   77
               Top             =   0
               Width           =   4635
            End
            Begin MSComctlLib.ListView lvSpellLoc 
               Height          =   1155
               Left            =   7080
               TabIndex        =   78
               Top             =   0
               Width           =   3135
               _ExtentX        =   5530
               _ExtentY        =   2037
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
         End
      End
      Begin VB.ComboBox cmbContainsAbil 
         Height          =   315
         Index           =   0
         ItemData        =   "frmMain.frx":0CE0
         Left            =   5100
         List            =   "frmMain.frx":0CE2
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "&Remove Filter"
         Height          =   555
         Index           =   3
         Left            =   9540
         TabIndex        =   75
         Top             =   240
         Width           =   795
      End
      Begin VB.ComboBox cmbSpellAttackType 
         Height          =   315
         ItemData        =   "frmMain.frx":0CE4
         Left            =   6300
         List            =   "frmMain.frx":0CE6
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox chkSpellLearnable 
         Caption         =   "Learnable Only"
         Height          =   435
         Left            =   3960
         TabIndex        =   70
         ToolTipText     =   "Learnable Only"
         Top             =   360
         Width           =   1035
      End
      Begin VB.ComboBox cmbSpellTarget 
         Height          =   315
         ItemData        =   "frmMain.frx":0CE8
         Left            =   7500
         List            =   "frmMain.frx":0CEA
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox cmbSpellMageryLevel 
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox cmbSpellMagery 
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtSpellFind 
         Height          =   285
         Left            =   120
         TabIndex        =   65
         Top             =   510
         Width           =   1875
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Next"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   67
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "&Apply Filter"
         Height          =   555
         Index           =   2
         Left            =   8820
         TabIndex        =   74
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblLabelArray 
         Alignment       =   2  'Center
         Caption         =   "Contains"
         Height          =   255
         Index           =   28
         Left            =   5100
         TabIndex        =   737
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblLabelArray 
         Alignment       =   2  'Center
         Caption         =   "Attack Type"
         Height          =   255
         Index           =   30
         Left            =   6300
         TabIndex        =   81
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblLabelArray 
         Alignment       =   2  'Center
         Caption         =   "Target"
         Height          =   195
         Index           =   15
         Left            =   7560
         TabIndex        =   82
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblLabelArray 
         Caption         =   "Level:"
         Height          =   195
         Index           =   13
         Left            =   3120
         TabIndex        =   80
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblLabelArray 
         Caption         =   "Magery:"
         Height          =   195
         Index           =   14
         Left            =   2100
         TabIndex        =   79
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame framNav 
      Caption         =   "Armour"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6075
      Index           =   1
      Left            =   60
      TabIndex        =   44
      Top             =   1200
      Visible         =   0   'False
      Width           =   10455
      Begin VB.CommandButton cmdFilter 
         Caption         =   "&Remove Filter"
         Height          =   555
         Index           =   5
         Left            =   9540
         TabIndex        =   59
         Top             =   240
         Width           =   795
      End
      Begin VB.CheckBox chkArmourNonMagic 
         Caption         =   "Non-Magic"
         Height          =   195
         Left            =   5580
         TabIndex        =   56
         Top             =   540
         Width           =   1095
      End
      Begin VB.CheckBox chkArmourNoLimit 
         Caption         =   "No Limit"
         Height          =   195
         Left            =   5580
         TabIndex        =   55
         Top             =   300
         Width           =   975
      End
      Begin VB.CheckBox chkArmourType 
         Caption         =   "Plate"
         Height          =   195
         Index           =   6
         Left            =   4800
         TabIndex        =   54
         Top             =   570
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkArmourType 
         Caption         =   "Scale"
         Height          =   195
         Index           =   5
         Left            =   4020
         TabIndex        =   53
         Top             =   570
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkArmourType 
         Caption         =   "Chain"
         Height          =   195
         Index           =   4
         Left            =   4020
         TabIndex        =   52
         Top             =   300
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.CheckBox chkArmourType 
         Caption         =   "Leather"
         Height          =   195
         Index           =   3
         Left            =   3120
         TabIndex        =   51
         Top             =   570
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkArmourType 
         Caption         =   "Ninja"
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   50
         Top             =   300
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.CheckBox chkArmourType 
         Caption         =   "Silk"
         Height          =   195
         Index           =   1
         Left            =   2220
         TabIndex        =   49
         Top             =   570
         Value           =   1  'Checked
         Width           =   675
      End
      Begin VB.CheckBox chkArmourType 
         Caption         =   "Natural"
         Height          =   195
         Index           =   0
         Left            =   2220
         TabIndex        =   48
         Top             =   300
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.ComboBox cmbArmourWorn 
         Height          =   315
         Left            =   6900
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "&Apply Filter"
         Height          =   555
         Index           =   4
         Left            =   8820
         TabIndex        =   58
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox txtArmourFind 
         Height          =   285
         Left            =   120
         TabIndex        =   45
         Top             =   510
         Width           =   1995
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Next"
         Height          =   255
         Index           =   3
         Left            =   1140
         TabIndex        =   47
         Top             =   240
         Width           =   975
      End
      Begin MMUD_Explorer.cntSplitter splSplitterNS 
         Height          =   5115
         Index           =   1
         Left            =   120
         TabIndex        =   740
         Top             =   840
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   9022
         Begin MMUD_Explorer.cntSplitter splSplitterWE 
            Height          =   975
            Index           =   1
            Left            =   0
            TabIndex        =   741
            Top             =   4140
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   1720
            Begin VB.TextBox txtArmourDetail 
               Height          =   975
               Left            =   0
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   61
               Top             =   0
               Width           =   2895
            End
            Begin MSComctlLib.ListView lvArmourLoc 
               Height          =   975
               Left            =   6180
               TabIndex        =   62
               Top             =   0
               Width           =   3915
               _ExtentX        =   6906
               _ExtentY        =   1720
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
         End
         Begin MSComctlLib.ListView lvArmour 
            Height          =   1815
            Left            =   0
            TabIndex        =   60
            Tag             =   "STRETCHALL"
            Top             =   0
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   3201
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Label lblLabelArray 
         Alignment       =   2  'Center
         Caption         =   "Worn On"
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
         Left            =   6900
         TabIndex        =   63
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame framNav 
      Caption         =   "Equipment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Index           =   4
      Left            =   60
      TabIndex        =   132
      Top             =   1200
      Visible         =   0   'False
      Width           =   10455
      Begin VB.CommandButton cmdEquipGoto 
         Caption         =   ">"
         Height          =   315
         Index           =   16
         Left            =   3720
         TabIndex        =   192
         Top             =   6240
         Width           =   195
      End
      Begin VB.CommandButton cmdEquipGoto 
         Caption         =   ">"
         Height          =   315
         Index           =   15
         Left            =   3720
         TabIndex        =   189
         Top             =   5940
         Width           =   195
      End
      Begin VB.CommandButton cmdEquipGoto 
         Caption         =   ">"
         Height          =   315
         Index           =   14
         Left            =   3720
         TabIndex        =   186
         Top             =   5640
         Width           =   195
      End
      Begin VB.CommandButton cmdEquipGoto 
         Caption         =   ">"
         Height          =   315
         Index           =   13
         Left            =   3720
         TabIndex        =   183
         Top             =   5340
         Width           =   195
      End
      Begin VB.CommandButton cmdEquipGoto 
         Caption         =   ">"
         Height          =   315
         Index           =   12
         Left            =   3720
         TabIndex        =   180
         Top             =   5040
         Width           =   195
      End
      Begin VB.CommandButton cmdEquipGoto 
         Caption         =   ">"
         Height          =   315
         Index           =   11
         Left            =   3720
         TabIndex        =   177
         Top             =   4740
         Width           =   195
      End
      Begin VB.CommandButton cmdEquipGoto 
         Caption         =   ">"
         Height          =   315
         Index           =   10
         Left            =   3720
         TabIndex        =   174
         Top             =   4440
         Width           =   195
      End
      Begin VB.CommandButton cmdEquipGoto 
         Caption         =   ">"
         Height          =   315
         Index           =   9
         Left            =   3720
         TabIndex        =   171
         Top             =   4140
         Width           =   195
      End
      Begin VB.CommandButton cmdEquipGoto 
         Caption         =   ">"
         Height          =   315
         Index           =   8
         Left            =   3720
         TabIndex        =   168
         Top             =   3840
         Width           =   195
      End
      Begin VB.CommandButton cmdEquipGoto 
         Caption         =   ">"
         Height          =   315
         Index           =   7
         Left            =   3720
         TabIndex        =   165
         Top             =   3540
         Width           =   195
      End
      Begin VB.CommandButton cmdEquipGoto 
         Caption         =   ">"
         Height          =   315
         Index           =   6
         Left            =   3720
         TabIndex        =   162
         Top             =   3240
         Width           =   195
      End
      Begin VB.CommandButton cmdEquipGoto 
         Caption         =   ">"
         Height          =   315
         Index           =   5
         Left            =   3720
         TabIndex        =   159
         Top             =   2940
         Width           =   195
      End
      Begin VB.CommandButton cmdEquipGoto 
         Caption         =   ">"
         Height          =   315
         Index           =   4
         Left            =   3720
         TabIndex        =   156
         Top             =   2640
         Width           =   195
      End
      Begin VB.CommandButton cmdEquipGoto 
         Caption         =   ">"
         Height          =   315
         Index           =   3
         Left            =   3720
         TabIndex        =   153
         Top             =   2340
         Width           =   195
      End
      Begin VB.CommandButton cmdEquipGoto 
         Caption         =   ">"
         Height          =   315
         Index           =   2
         Left            =   3720
         TabIndex        =   150
         Top             =   2040
         Width           =   195
      End
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   0
         ItemData        =   "frmMain.frx":0CEC
         Left            =   1140
         List            =   "frmMain.frx":0CEE
         Sorted          =   -1  'True
         TabIndex        =   143
         Text            =   "cmbEquip"
         Top             =   840
         Width           =   2535
      End
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   1
         ItemData        =   "frmMain.frx":0CF0
         Left            =   1140
         List            =   "frmMain.frx":0CF2
         Sorted          =   -1  'True
         TabIndex        =   146
         Text            =   "cmbEquip"
         Top             =   1140
         Width           =   2535
      End
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   17
         ItemData        =   "frmMain.frx":0CF4
         Left            =   1140
         List            =   "frmMain.frx":0CF6
         Sorted          =   -1  'True
         TabIndex        =   858
         Text            =   "cmbEquip"
         Top             =   1440
         Width           =   2535
      End
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   18
         ItemData        =   "frmMain.frx":0CF8
         Left            =   1140
         List            =   "frmMain.frx":0CFA
         Sorted          =   -1  'True
         TabIndex        =   859
         Text            =   "cmbEquip"
         Top             =   1740
         Width           =   2535
      End
      Begin VB.CommandButton cmdEquipGoto 
         Caption         =   ">"
         Height          =   315
         Index           =   18
         Left            =   3720
         TabIndex        =   863
         Top             =   1740
         Width           =   195
      End
      Begin VB.CommandButton cmdEquipGoto 
         Caption         =   ">"
         Height          =   315
         Index           =   17
         Left            =   3720
         TabIndex        =   862
         Top             =   1440
         Width           =   195
      End
      Begin VB.CheckBox chkEquipHold 
         Caption         =   "Face"
         Height          =   195
         Index           =   18
         Left            =   120
         TabIndex        =   861
         Top             =   1800
         Width           =   1035
      End
      Begin VB.CheckBox chkEquipHold 
         Caption         =   "Eyes"
         Height          =   195
         Index           =   17
         Left            =   120
         TabIndex        =   860
         Top             =   1500
         Width           =   1035
      End
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   2
         ItemData        =   "frmMain.frx":0CFC
         Left            =   1140
         List            =   "frmMain.frx":0CFE
         Sorted          =   -1  'True
         TabIndex        =   149
         Text            =   "cmbEquip"
         Top             =   2040
         Width           =   2535
      End
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   3
         ItemData        =   "frmMain.frx":0D00
         Left            =   1140
         List            =   "frmMain.frx":0D02
         Sorted          =   -1  'True
         TabIndex        =   152
         Text            =   "cmbEquip"
         Top             =   2340
         Width           =   2535
      End
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   4
         ItemData        =   "frmMain.frx":0D04
         Left            =   1140
         List            =   "frmMain.frx":0D06
         Sorted          =   -1  'True
         TabIndex        =   155
         Text            =   "cmbEquip"
         Top             =   2640
         Width           =   2535
      End
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   5
         ItemData        =   "frmMain.frx":0D08
         Left            =   1140
         List            =   "frmMain.frx":0D0A
         Sorted          =   -1  'True
         TabIndex        =   158
         Text            =   "cmbEquip"
         Top             =   2940
         Width           =   2535
      End
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   6
         ItemData        =   "frmMain.frx":0D0C
         Left            =   1140
         List            =   "frmMain.frx":0D0E
         Sorted          =   -1  'True
         TabIndex        =   161
         Text            =   "cmbEquip"
         Top             =   3240
         Width           =   2535
      End
      Begin VB.CheckBox chkEquipHold 
         Caption         =   "Wrist"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   160
         Top             =   3300
         Width           =   1035
      End
      Begin VB.CheckBox chkInvenHideCharStats 
         Caption         =   "Hide Character Specific Stats"
         Height          =   195
         Left            =   5220
         TabIndex        =   203
         Top             =   900
         Width           =   2535
      End
      Begin VB.CommandButton cmdInvenAdditionalWeight 
         Caption         =   "Additional Item Weight"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   195
         Top             =   1500
         Width           =   915
      End
      Begin VB.CommandButton cmdResetFonts 
         Caption         =   "Reset Fonts"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   201
         Top             =   4680
         Width           =   675
      End
      Begin VB.CommandButton cmdInvenSetStatFont 
         Caption         =   "Stat Font"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   200
         Top             =   4200
         Width           =   675
      End
      Begin VB.CommandButton cmdInvenEmptyLists 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   4740
         TabIndex        =   198
         Top             =   3000
         Width           =   255
      End
      Begin VB.CommandButton cmdInvenEmptyLists 
         Caption         =   "Empty Lists"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   4080
         TabIndex        =   197
         Top             =   3000
         Width           =   675
      End
      Begin VB.CommandButton cmdInvenCompareAll 
         Caption         =   "Add All to &Compare"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4080
         TabIndex        =   196
         Top             =   2220
         Width           =   915
      End
      Begin VB.CheckBox chkInvenAddWeight 
         Caption         =   "(Use)"
         Height          =   195
         Left            =   4200
         TabIndex        =   193
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtInvenAddWeight 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4080
         TabIndex        =   194
         Top             =   1200
         Width           =   915
      End
      Begin VB.CommandButton cmdInvenModStrength 
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
         Height          =   255
         Index           =   1
         Left            =   10020
         TabIndex        =   206
         Top             =   840
         Width           =   315
      End
      Begin VB.CommandButton cmdInvenModStrength 
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
         Height          =   255
         Index           =   0
         Left            =   9720
         TabIndex        =   205
         Top             =   840
         Width           =   315
      End
      Begin VB.CommandButton cmdInvenNextBest 
         Caption         =   "Ne&xt Best"
         Height          =   375
         Left            =   6420
         TabIndex        =   137
         Top             =   180
         Width           =   1095
      End
      Begin VB.PictureBox picStats 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   4695
         Left            =   5160
         ScaleHeight     =   311
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   343
         TabIndex        =   667
         Top             =   1200
         Width           =   5175
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   3
            Left            =   2820
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   726
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   420
            Width           =   735
         End
         Begin VB.TextBox txtStat 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   2
            Left            =   1740
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   725
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   420
            Width           =   795
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   4
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   695
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   2595
            Width           =   735
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   5
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   694
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   795
            Width           =   735
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   6
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   693
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   1095
            Width           =   735
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   7
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   692
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   1395
            Width           =   735
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   8
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   691
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   1695
            Width           =   735
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   9
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   690
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   1995
            Width           =   735
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   10
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   689
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   2295
            Width           =   735
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   11
            Left            =   4320
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   688
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   1395
            Width           =   795
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   12
            Left            =   4320
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   687
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   2595
            Width           =   795
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   13
            Left            =   4320
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   686
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   2295
            Width           =   795
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   14
            Left            =   4320
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   685
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   1695
            Width           =   795
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   16
            Left            =   4320
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   684
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   795
            Width           =   795
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   17
            Left            =   4320
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   683
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   1095
            Width           =   795
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   18
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   682
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   2895
            Width           =   735
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   19
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   681
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   3195
            Width           =   735
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   20
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   680
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   3495
            Width           =   735
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   21
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   679
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   3795
            Width           =   735
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   22
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   678
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   4095
            Width           =   735
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   25
            Left            =   4320
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   677
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   3195
            Width           =   795
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   26
            Left            =   4320
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   676
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   3495
            Width           =   795
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   27
            Left            =   4320
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   675
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   3795
            Width           =   795
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   28
            Left            =   4320
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   674
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   4095
            Width           =   795
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   29
            Left            =   4320
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   673
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   4395
            Width           =   795
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   23
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   672
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   4395
            Width           =   735
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   24
            Left            =   4320
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   671
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   2895
            Width           =   795
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   15
            Left            =   4320
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   670
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   1995
            Width           =   795
         End
         Begin VB.TextBox txtStat 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   1
            Left            =   2820
            Locked          =   -1  'True
            MaxLength       =   7
            TabIndex        =   669
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   60
            Width           =   735
         End
         Begin VB.TextBox txtStat 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   290
            Index           =   0
            Left            =   1740
            Locked          =   -1  'True
            MaxLength       =   7
            TabIndex        =   668
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   60
            Width           =   795
         End
         Begin VB.Label lblInvenSlash 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   300
            Index           =   1
            Left            =   2610
            TabIndex        =   728
            Top             =   420
            Width           =   150
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Armour Class:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   727
            Top             =   435
            Width           =   1560
         End
         Begin VB.Label lblEncumLevel 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "None (0%)"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   600
            Index           =   1
            Left            =   3600
            TabIndex        =   724
            Top             =   135
            Width           =   1560
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Picklocks:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   22
            Left            =   120
            TabIndex        =   723
            Top             =   4095
            Width           =   1200
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Traps:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   21
            Left            =   120
            TabIndex        =   722
            Top             =   3795
            Width           =   720
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Prot Evil:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   20
            Left            =   120
            TabIndex        =   721
            Top             =   3495
            Width           =   1200
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Stealth:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   19
            Left            =   120
            TabIndex        =   720
            Top             =   3195
            Width           =   960
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Perception:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   18
            Left            =   120
            TabIndex        =   719
            Top             =   2895
            Width           =   1320
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "SpellCast:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   9
            Left            =   120
            TabIndex        =   718
            Top             =   1995
            Width           =   1200
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Mana:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   6
            Left            =   120
            TabIndex        =   717
            Top             =   1095
            Width           =   600
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "HitPoints:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   5
            Left            =   120
            TabIndex        =   716
            Top             =   795
            Width           =   1200
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Max Damage:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   11
            Left            =   2400
            TabIndex        =   715
            Top             =   1395
            Width           =   1320
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Magic Resist:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   24
            Left            =   2400
            TabIndex        =   714
            Top             =   2895
            Width           =   1560
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Lightning Res.:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   29
            Left            =   2400
            TabIndex        =   713
            Top             =   4395
            Width           =   1800
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Accuracy:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   10
            Left            =   120
            TabIndex        =   712
            Top             =   2295
            Width           =   1080
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Crits:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   7
            Left            =   120
            TabIndex        =   711
            Top             =   1410
            Width           =   720
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Dodge:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   8
            Left            =   120
            TabIndex        =   710
            Top             =   1710
            Width           =   720
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Illuminatn:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   23
            Left            =   120
            TabIndex        =   709
            Top             =   4395
            Width           =   1320
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Water Resist:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   26
            Left            =   2400
            TabIndex        =   708
            Top             =   3495
            Width           =   1560
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Fire Resist:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   27
            Left            =   2400
            TabIndex        =   707
            Top             =   3795
            Width           =   1440
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Stone Resist:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   25
            Left            =   2400
            TabIndex        =   706
            Top             =   3195
            Width           =   1560
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "HitPoint Regen:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   16
            Left            =   2400
            TabIndex        =   705
            Top             =   795
            Width           =   1800
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "BS Accuracy:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   13
            Left            =   2400
            TabIndex        =   704
            Top             =   2295
            Width           =   1440
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "BS Min Damage:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   14
            Left            =   2400
            TabIndex        =   703
            Top             =   1695
            Width           =   1680
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "BS Max Damage:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   15
            Left            =   2400
            TabIndex        =   702
            Top             =   1995
            Width           =   1680
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Cold Resist:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   28
            Left            =   2400
            TabIndex        =   701
            Top             =   4095
            Width           =   1440
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Mana Regen:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   17
            Left            =   2400
            TabIndex        =   700
            Top             =   1095
            Width           =   1320
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Shield Damage:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   12
            Left            =   2400
            TabIndex        =   699
            Top             =   2595
            Width           =   1680
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Encum.:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   698
            Top             =   2595
            Width           =   840
         End
         Begin VB.Label lblInvenSlash 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   300
            Index           =   0
            Left            =   2610
            TabIndex        =   697
            Top             =   60
            Width           =   150
         End
         Begin VB.Label lblInvenStats 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Encumbrance:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   696
            Top             =   75
            Width           =   1440
         End
      End
      Begin VB.CommandButton cmdEquipCheckAll 
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1740
         TabIndex        =   141
         Top             =   630
         Width           =   615
      End
      Begin VB.CommandButton cmdEquipCheckAll 
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1155
         TabIndex        =   140
         Top             =   630
         Width           =   555
      End
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   7
         ItemData        =   "frmMain.frx":0D10
         Left            =   1140
         List            =   "frmMain.frx":0D12
         Sorted          =   -1  'True
         TabIndex        =   164
         Text            =   "cmbEquip"
         Top             =   3540
         Width           =   2535
      End
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   8
         ItemData        =   "frmMain.frx":0D14
         Left            =   1140
         List            =   "frmMain.frx":0D16
         Sorted          =   -1  'True
         TabIndex        =   167
         Text            =   "cmbEquip"
         Top             =   3840
         Width           =   2535
      End
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   9
         ItemData        =   "frmMain.frx":0D18
         Left            =   1140
         List            =   "frmMain.frx":0D1A
         Sorted          =   -1  'True
         TabIndex        =   170
         Text            =   "cmbEquip"
         Top             =   4140
         Width           =   2535
      End
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   10
         ItemData        =   "frmMain.frx":0D1C
         Left            =   1140
         List            =   "frmMain.frx":0D1E
         Sorted          =   -1  'True
         TabIndex        =   173
         Text            =   "cmbEquip"
         Top             =   4440
         Width           =   2535
      End
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   11
         ItemData        =   "frmMain.frx":0D20
         Left            =   1140
         List            =   "frmMain.frx":0D22
         Sorted          =   -1  'True
         TabIndex        =   176
         Text            =   "cmbEquip"
         Top             =   4740
         Width           =   2535
      End
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   12
         ItemData        =   "frmMain.frx":0D24
         Left            =   1140
         List            =   "frmMain.frx":0D26
         Sorted          =   -1  'True
         TabIndex        =   179
         Text            =   "cmbEquip"
         Top             =   5040
         Width           =   2535
      End
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   13
         ItemData        =   "frmMain.frx":0D28
         Left            =   1140
         List            =   "frmMain.frx":0D2A
         Sorted          =   -1  'True
         TabIndex        =   182
         Text            =   "cmbEquip"
         Top             =   5340
         Width           =   2535
      End
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   14
         ItemData        =   "frmMain.frx":0D2C
         Left            =   1140
         List            =   "frmMain.frx":0D2E
         Sorted          =   -1  'True
         TabIndex        =   185
         Text            =   "cmbEquip"
         Top             =   5640
         Width           =   2535
      End
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   15
         ItemData        =   "frmMain.frx":0D30
         Left            =   1140
         List            =   "frmMain.frx":0D32
         Sorted          =   -1  'True
         TabIndex        =   188
         Text            =   "cmbEquip"
         Top             =   5940
         Width           =   2535
      End
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   16
         ItemData        =   "frmMain.frx":0D34
         Left            =   1140
         List            =   "frmMain.frx":0D36
         Sorted          =   -1  'True
         TabIndex        =   191
         Text            =   "cmbEquip"
         Top             =   6240
         Width           =   2535
      End
      Begin VB.CheckBox chkEquipHold 
         Caption         =   "Weapon"
         Height          =   195
         Index           =   16
         Left            =   120
         TabIndex        =   190
         Top             =   6300
         Width           =   1035
      End
      Begin VB.CheckBox chkEquipHold 
         Caption         =   "Off-Hand"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   187
         Top             =   6000
         Width           =   1035
      End
      Begin VB.CheckBox chkEquipHold 
         Caption         =   "Worn"
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   184
         Top             =   5700
         Width           =   1035
      End
      Begin VB.CheckBox chkEquipHold 
         Caption         =   "Feet"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   181
         Top             =   5400
         Width           =   1035
      End
      Begin VB.CheckBox chkEquipHold 
         Caption         =   "Legs"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   178
         Top             =   5100
         Width           =   1035
      End
      Begin VB.CheckBox chkEquipHold 
         Caption         =   "Finger"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   172
         Top             =   4500
         Width           =   1035
      End
      Begin VB.CheckBox chkEquipHold 
         Caption         =   "Finger"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   169
         Top             =   4200
         Width           =   1035
      End
      Begin VB.CheckBox chkEquipHold 
         Caption         =   "Hands"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   166
         Top             =   3900
         Width           =   1035
      End
      Begin VB.CheckBox chkEquipHold 
         Caption         =   "Waist"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   175
         Top             =   4800
         Width           =   1035
      End
      Begin VB.CheckBox chkEquipHold 
         Caption         =   "Wrist"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   163
         Top             =   3600
         Width           =   1035
      End
      Begin VB.CheckBox chkEquipHold 
         Caption         =   "Arms"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   157
         Top             =   3000
         Width           =   1035
      End
      Begin VB.CheckBox chkEquipHold 
         Caption         =   "Torso"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   154
         Top             =   2700
         Width           =   1035
      End
      Begin VB.CheckBox chkEquipHold 
         Caption         =   "Back"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   151
         Top             =   2400
         Width           =   1035
      End
      Begin VB.CheckBox chkEquipHold 
         Caption         =   "Neck"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   148
         Top             =   2100
         Width           =   1035
      End
      Begin VB.CheckBox chkEquipHold 
         Caption         =   "Ears"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   145
         Top             =   1200
         Width           =   1035
      End
      Begin VB.CheckBox chkEquipHold 
         Caption         =   "Head"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   142
         Top             =   900
         Width           =   1035
      End
      Begin VB.CommandButton cmdInvenClipboard 
         Caption         =   "Cop&y to Clipboard"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4080
         TabIndex        =   202
         Top             =   5340
         Width           =   915
      End
      Begin VB.CommandButton cmdInvenPasteChar 
         Caption         =   "&Paste Char.."
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
         Index           =   0
         Left            =   9000
         TabIndex        =   139
         Top             =   180
         Width           =   1335
      End
      Begin VB.CheckBox chkInvenNoLimited 
         Caption         =   "Non- &Limiteds"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7680
         TabIndex        =   138
         ToolTipText     =   "Non-Limiteds"
         Top             =   240
         Width           =   1395
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "Re&move Filter"
         Height          =   315
         Index           =   1
         Left            =   1980
         TabIndex        =   134
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "&Apply Global Filter"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   133
         Top             =   240
         Width           =   1755
      End
      Begin VB.CommandButton cmdEquipGoto 
         Caption         =   ">"
         Height          =   315
         Index           =   1
         Left            =   3720
         TabIndex        =   147
         Top             =   1140
         Width           =   195
      End
      Begin VB.CommandButton cmdEquipGoto 
         Caption         =   ">"
         Height          =   315
         Index           =   0
         Left            =   3720
         TabIndex        =   144
         Top             =   840
         Width           =   195
      End
      Begin VB.TextBox txtInvenStrength 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   8880
         MaxLength       =   4
         TabIndex        =   204
         Text            =   "100"
         Top             =   840
         Width           =   795
      End
      Begin VB.CommandButton cmdInvenFindBest 
         Caption         =   "&Find Best..."
         Height          =   375
         Left            =   5160
         TabIndex        =   136
         Top             =   180
         Width           =   1155
      End
      Begin VB.CommandButton cmdInvenReset 
         Caption         =   "&Reset"
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
         Left            =   3720
         TabIndex        =   135
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton cmdSetInvenLabelFont 
         Caption         =   "Label Font"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   199
         Top             =   3720
         Width           =   675
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   10320
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblLabelArray 
         AutoSize        =   -1  'True
         Caption         =   "Jump"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   7
         Left            =   3660
         TabIndex        =   666
         Top             =   660
         Width           =   345
      End
      Begin VB.Label lblLabelArray 
         AutoSize        =   -1  'True
         Caption         =   "Hold"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   8
         Left            =   60
         TabIndex        =   665
         Top             =   660
         Width           =   300
      End
      Begin VB.Label lblLabelArray 
         Caption         =   "Strength:"
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
         Index           =   9
         Left            =   7860
         TabIndex        =   207
         Top             =   870
         Width           =   1095
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Fi&le"
      Index           =   0
      Begin VB.Menu mnuOpenDataFile 
         Caption         =   "&Open Data File"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuB 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOptItem 
         Caption         =   "&Load Character"
         Index           =   0
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuFileOptItem 
         Caption         =   "&Revert to Saved"
         Index           =   1
         Shortcut        =   %{BKSP}
      End
      Begin VB.Menu mnuFileOptItem 
         Caption         =   "&Save Character"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileOptItem 
         Caption         =   "Save Character &As..."
         Index           =   3
      End
      Begin VB.Menu mnuFileOptItem 
         Caption         =   "&Close File"
         Index           =   4
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^{F12}
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Options"
      Index           =   1
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuReload 
         Caption         =   "&Reload Program"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuRemoveFilters 
         Caption         =   "Re&move All Filters"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuC 
         Caption         =   "-"
      End
      Begin VB.Menu mnuJumpToCompare 
         Caption         =   "Jump to Equip/Compare on Add"
         Checked         =   -1  'True
         Shortcut        =   ^J
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Tools"
      Index           =   2
      Begin VB.Menu mnuToolsItems 
         Caption         =   "&Backstab Calculator"
         Index           =   0
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuToolsItems 
         Caption         =   "E&xternal Map"
         Index           =   1
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuToolsItems 
         Caption         =   "E&xp Calculator"
         Index           =   2
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuToolsItems 
         Caption         =   "&Notepad"
         Index           =   3
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuToolsItems 
         Caption         =   "&Swing Calculator"
         Index           =   4
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Help"
      Index           =   3
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   0
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Release Notes"
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Check for Updates"
         Index           =   2
      End
   End
   Begin VB.Menu mnuItemsPopUp 
      Caption         =   "ItemsPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuItemsPopUpItem 
         Caption         =   "Add to Compare"
         Index           =   0
      End
      Begin VB.Menu mnuItemsPopUpItem 
         Caption         =   "Calculate Swings"
         Index           =   1
      End
      Begin VB.Menu mnuItemsPopUpItem 
         Caption         =   "Copy to Clipboard"
         Index           =   2
      End
      Begin VB.Menu mnuItemsPopUpItem 
         Caption         =   "Equip Item"
         Index           =   3
      End
      Begin VB.Menu mnuItemsPopUpItem 
         Caption         =   "Calculate Backstab Damage"
         Index           =   4
      End
   End
   Begin VB.Menu mnuSpellsPopUp 
      Caption         =   "SpellPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuSpellsPopUpItem 
         Caption         =   "Add to Compare"
         Index           =   0
      End
      Begin VB.Menu mnuSpellsPopUpItem 
         Caption         =   "Copy to Clipboard"
         Index           =   1
      End
      Begin VB.Menu mnuSpellsPopUpItem 
         Caption         =   "What casts this spell?"
         Index           =   2
      End
   End
   Begin VB.Menu mnuAuxPopUp 
      Caption         =   "AuxPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuAuxPopUpItem 
         Caption         =   "Copy to Clipboard"
         Index           =   0
      End
      Begin VB.Menu mnuAuxPopUpItem 
         Caption         =   "Where/How is this Monster Summoned?"
         Index           =   1
      End
      Begin VB.Menu mnuAuxPopUpItem 
         Caption         =   "Add Items to Compare"
         Index           =   2
      End
      Begin VB.Menu mnuAuxPopUpItem 
         Caption         =   "Add to Monster Compare"
         Index           =   3
      End
      Begin VB.Menu mnuAuxPopUpItem 
         Caption         =   "Copy Chest to Clipboard"
         Index           =   4
      End
   End
   Begin VB.Menu mnuFindBest 
      Caption         =   "FindBestMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuFindBestItem 
         Caption         =   "Armour"
         Index           =   0
         Begin VB.Menu mnuFindBestArmour 
            Caption         =   "AC/DR Combo"
            Index           =   0
         End
         Begin VB.Menu mnuFindBestArmour 
            Caption         =   "AC"
            Index           =   1
         End
         Begin VB.Menu mnuFindBestArmour 
            Caption         =   "DR"
            Index           =   2
         End
         Begin VB.Menu mnuFindBestArmour 
            Caption         =   "Dodge"
            Index           =   3
         End
         Begin VB.Menu mnuFindBestArmour 
            Caption         =   "Prot. from Evil"
            Index           =   4
         End
         Begin VB.Menu mnuFindBestArmour 
            Caption         =   "Prot. from Good"
            Index           =   5
         End
      End
      Begin VB.Menu mnuFindBestItem 
         Caption         =   "Attacking"
         Index           =   1
         Begin VB.Menu mnuFindBestAttack 
            Caption         =   "Accuracy"
            Index           =   0
         End
         Begin VB.Menu mnuFindBestAttack 
            Caption         =   "BS Accuracy"
            Index           =   1
         End
         Begin VB.Menu mnuFindBestAttack 
            Caption         =   "BS Min"
            Index           =   2
         End
         Begin VB.Menu mnuFindBestAttack 
            Caption         =   "BS Max"
            Index           =   3
         End
         Begin VB.Menu mnuFindBestAttack 
            Caption         =   "Crits"
            Index           =   4
         End
         Begin VB.Menu mnuFindBestAttack 
            Caption         =   "DamageShield"
            Index           =   5
         End
         Begin VB.Menu mnuFindBestAttack 
            Caption         =   "Max Damage"
            Index           =   6
         End
         Begin VB.Menu mnuFindBestAttackMystics 
            Caption         =   "Mystic Attacks"
            Begin VB.Menu mnuFindBestMystics 
               Caption         =   "JumpKick ACY"
               Index           =   0
            End
            Begin VB.Menu mnuFindBestMystics 
               Caption         =   "Jumpkcik DMG"
               Index           =   1
            End
            Begin VB.Menu mnuFindBestMystics 
               Caption         =   "Kick ACY"
               Index           =   2
            End
            Begin VB.Menu mnuFindBestMystics 
               Caption         =   "Kick DMG"
               Index           =   3
            End
            Begin VB.Menu mnuFindBestMystics 
               Caption         =   "Punch ACY"
               Index           =   4
            End
            Begin VB.Menu mnuFindBestMystics 
               Caption         =   "Punch DMG"
               Index           =   5
            End
         End
      End
      Begin VB.Menu mnuFindBestItem 
         Caption         =   "Resistances"
         Index           =   2
         Begin VB.Menu mnuFindBestResist 
            Caption         =   "All Magic (MR)"
            Index           =   0
         End
         Begin VB.Menu mnuFindBestResist 
            Caption         =   "Cold"
            Index           =   1
         End
         Begin VB.Menu mnuFindBestResist 
            Caption         =   "Fire"
            Index           =   2
         End
         Begin VB.Menu mnuFindBestResist 
            Caption         =   "Lightning"
            Index           =   3
         End
         Begin VB.Menu mnuFindBestResist 
            Caption         =   "Stone"
            Index           =   4
         End
         Begin VB.Menu mnuFindBestResist 
            Caption         =   "Water"
            Index           =   5
         End
      End
      Begin VB.Menu mnuFindBestItem 
         Caption         =   "Stats"
         Index           =   3
         Begin VB.Menu mnuFindBestStats 
            Caption         =   "Encumberance"
            Index           =   0
         End
         Begin VB.Menu mnuFindBestStats 
            Caption         =   "HP"
            Index           =   1
         End
         Begin VB.Menu mnuFindBestStats 
            Caption         =   "HP Regen"
            Index           =   2
         End
         Begin VB.Menu mnuFindBestStats 
            Caption         =   "Illumination"
            Index           =   3
         End
         Begin VB.Menu mnuFindBestStats 
            Caption         =   "Mana"
            Index           =   4
         End
         Begin VB.Menu mnuFindBestStats 
            Caption         =   "Mana Regen"
            Index           =   5
         End
         Begin VB.Menu mnuFindBestStats 
            Caption         =   "Picklocks"
            Index           =   6
         End
         Begin VB.Menu mnuFindBestStats 
            Caption         =   "Spellcasting"
            Index           =   7
         End
         Begin VB.Menu mnuFindBestStats 
            Caption         =   "Stealth"
            Index           =   8
         End
         Begin VB.Menu mnuFindBestStats 
            Caption         =   "Thievery"
            Index           =   9
         End
         Begin VB.Menu mnuFindBestStats 
            Caption         =   "Traps"
            Index           =   10
         End
      End
   End
   Begin VB.Menu mnuMapPopUp 
      Caption         =   "MapMenuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuMapPopUpItem 
         Caption         =   "Follow Up and Redraw"
         Index           =   0
      End
      Begin VB.Menu mnuMapPopUpItem 
         Caption         =   "Follow Down and Redraw"
         Index           =   1
      End
      Begin VB.Menu mnuMapPopUpItem 
         Caption         =   "Redraw From Here"
         Index           =   2
      End
   End
   Begin VB.Menu mnuEquipGoto 
      Caption         =   "Equip Goto"
      Visible         =   0   'False
      Begin VB.Menu mnuEquipGotoItem 
         Caption         =   "Goto Item"
         Index           =   0
      End
      Begin VB.Menu mnuEquipGotoItem 
         Caption         =   "Add to Compare"
         Index           =   1
      End
      Begin VB.Menu mnuEquipGotoItem 
         Caption         =   "Calculate Swings"
         Index           =   2
      End
      Begin VB.Menu mnuEquipGotoItem 
         Caption         =   "Calculate Backstab Damage"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Const WM_SETREDRAW = &HB
Private msOldString As String ' module level global
Private miStart As Integer    ' module level global
Private miLength As Integer   ' module level global

'// main menu stuff <--

'Private Type Declarations
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

'Private API Declarations
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
Private Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32.dll" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Private Declare Function GetMenuItemInfo Lib "user32.dll" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long

'Private Constants
Private Const MIIM_STATE = &H1
Private Const MIIM_ID = &H2
Private Const MIIM_TYPE = &H10
Private Const MFT_SEPARATOR = &H800
Private Const MFT_STRING = &H0
Private Const MFS_ENABLED = &H0
Private Const MFS_CHECKED = &H8

'// main menu stuff -->

Private Enum EnumDrawRoom
    drSquare = 0
    drStar = 1
    drOpenCircle = 2
    drUp = 3
    drDown = 4
    drCircle = 5
    drLineN = 6
    drLineS = 7
    drLineE = 8
    drLineW = 9
    drLineNE = 10
    drLineNW = 11
    drLineSE = 12
    drLineSW = 13
End Enum

Private Enum enmFindBest
    Armour = 0
    Attack = 1
    Resist = 2
    Stats = 3
    Mystics = 4
End Enum

Private Type TypeEquipWinner
    Value As Long
    Number As Long
End Type

Private Type TypeGetEquip
    nEquip As Integer
    sText As String
End Type

Dim nBSStealth As Integer
Dim bMouseDown As Boolean
Dim objWorkingListView As ListView
Dim sNormalCaption As String
Dim bKeepSortOrder As Boolean
Dim bSortOrderAsc As Boolean
Dim oLastColumnSorted As ColumnHeader
Dim bDontRefresh As Boolean
Dim bStartup As Boolean
Dim sAddWeight As String
Dim sCharSwings As String
Dim bPrevInstanceWarned As Boolean

Public bAutoSave As Boolean
Public bNO_RECENT_FILES As Boolean
Public bNameInTitle As Boolean
Public bDontSpanNav As Boolean
Public bReloadProgram As Boolean
Public bCharLoaded As Boolean
Public bDontCallTerminate As Boolean
Public bNoAlwaysOnTop As Boolean
Public bDontSetMainFocus As Boolean
Public bDontSaveSettings As Boolean
Public nWindowState As Integer

Dim bInvenNextBest As Boolean
Dim nInvenLastIndex(1) As Integer
Dim nInvenExcludedItems() As Long
Public bInvenUse2ndWrist  As Boolean

Dim nLastShopDetailIndex As Integer
Dim nLastShopSort As Integer
Dim nLastWeaponSort As Integer
Dim nLastSpellSort As Integer
Dim nLastArmourSort As Integer
Dim nLastEquipIndex As Integer

Dim nMapLastFind(0 To 2) As Long
Dim nMapLastCellIndex As Integer
Dim bMapStillMapping As Boolean
Dim sMapSECorner As Integer
Dim nMapRowLength As Integer
Public nMapStartRoom As Long
Public nMapStartMap As Long
Dim nMapCenterCell As Integer
Dim sMapSearch As String
Dim nMapLastRoom As Long
Dim nMapLastMap As Long
Dim nMapCurrentRecord As Variant
Public bMapSwapButtons As Boolean
Public bMapCancelFind As Boolean
Dim CellRoom(1 To 400, 1 To 2) As Long
Dim UnchartedCells(1 To 400) As Integer
Dim StopBuild As Boolean

Dim TTlbl As clsToolTip
Dim objToolTip As clsToolTip

Private Sub cmbChar2ndAlign_Click()

bPromptSave = True
Call RefreshAll
End Sub

Private Sub cmbGlobalAlignment_Change()
bPromptSave = True
End Sub

Private Sub Form_Load()
On Error GoTo Error:
Dim fso As FileSystemObject, sFile As String, x As Integer, bResult As Boolean

'bNO_RECENT_FILES = True  'TURN OFF (comment out) BEFORE RELEASE

bSortOrderAsc = True
bSuppressErrors = False
bStartup = True
bDontRefresh = True
bCharLoaded = False
sNormalCaption = App.Title & " v" & App.Major & "." & App.Minor '& " BETA"
If bNO_RECENT_FILES Then sNormalCaption = sNormalCaption & " (NO RECENT)"
Me.Caption = sNormalCaption

If bAppTerminating Then GoTo term:

If App.PrevInstance And Not bPrevInstanceWarned Then
    x = MsgBox(App.Title & " is already running, open another copy?", _
        vbExclamation + vbYesNo + vbDefaultButton2)
    If x = vbNo Then
        GoTo term:
    End If
End If
bPrevInstanceWarned = True

Load frmLoad
frmLoad.lblCaption.Caption = "Loading ..."
frmLoad.Show
DoEvents

Call GetTitleBarOffset
With EL1
    .CenterOnLoad = True
    .FormInQuestion = Me
    .MinWidth = 713
    .MinHeight = 533 + (TITLEBAR_OFFSET / 10)
    .EnableLimiter = True
End With

If Not bNO_RECENT_FILES Then
    nMenuItemID = 1000 'initialize to one thousand in order to avoid conflicts with existing menu item IDs
    oldWindowProc = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf WindowProc) 'set up a new window procedure for this form and save a pointer to the original one as 'oldWindowProc'
End If

'Set objToolTip = Nothing
Set objToolTip = Nothing
'Set objToolTip = New clsToolTip
Set objToolTip = New clsToolTip

'With objToolTip
'    .DelayTime = 20
'    .VisibleTime = 20000
'    .BkColor = &HC0FFFF
'    .TxtColor = &H0
'    .Style = ttStyleStandard
'    '.Style = ttStyleStandard
'End With

With objToolTip
    .DelayTime = 20
    .VisibleTime = 20000
    .BkColor = &HC0FFFF
    .TxtColor = &H0
    .Style = ttStyleStandard
    '.Style = ttStyleStandard
End With

Set fso = CreateObject("Scripting.FileSystemObject")

If Right(App.Path, 1) = "\" Then
    INIFileName = App.Path & "settings.ini"
Else
    INIFileName = App.Path & "\settings.ini"
End If

If Not fso.FileExists(INIFileName) Then Call CreateSettings
'If Not fso.FileExists(ReadINI("Settings", "DataFile")) Then
'    Call WriteINI("Settings", "DataFile", "data-v1.11n.mdb")
'End If

Call SetUpFormObjects
Call SetupSplitters

sFile = ReadINI("Settings", "DataFile")

If InStr(1, sFile, "\") = 0 Then
    If Right(App.Path, 1) = "\" Then
        sFile = App.Path & sFile
    Else
        sFile = App.Path & "\" & sFile
    End If
End If
    
If Not fso.FileExists(sFile) Then
    MsgBox "Datafile (" & sFile & ")" & vbCrLf & "was not found.  Select open data file from the File menu.", vbInformation
    GoTo skipload:
End If

'GoTo skipload:

If ReadINI("Settings", "OnlyInGame") = 1 Then
    bOnlyInGame = True
Else
    bOnlyInGame = False
End If

If ReadINI("Settings", "HideRecordNumbers") = "1" Then
    bHideRecordNumbers = True
Else
    bHideRecordNumbers = False
End If

bResult = OpenTables(sFile)
If bResult = False Then GoTo skipload:
bResult = LoadInfo
If bResult = False Then GoTo skipload:

frmLoad.lblCaption.Caption = "Loading Classes..."
DoEvents
Call LoadClasses

frmLoad.lblCaption.Caption = "Loading Races..."
DoEvents
Call LoadRaces

If ReadINI("Settings", "LoadItems") = 1 Then
    frmLoad.lblCaption.Caption = "Loading Items..."
    DoEvents
    Call LoadItems
End If

If ReadINI("Settings", "LoadMonsters") = 1 Then
    chkMonstersNoRegenLookUp.Value = ReadINI("Setting", "LookUpMonsterRegen")
    frmLoad.lblCaption.Caption = "Loading Monsters..."
    DoEvents
    Call LoadMonsters
End If

If ReadINI("Settings", "LoadShops") = 1 Then
    frmLoad.lblCaption.Caption = "Loading Shops..."
    DoEvents
    Call LoadShops
End If

If ReadINI("Settings", "LoadSpells") = 1 Then
    frmLoad.lblCaption.Caption = "Loading Spells..."
    DoEvents
    Call LoadSpells
End If

For x = 0 To 9
    Call AutoSizeDropDownWidth(cmbCharBless(x))
    Call ExpandCombo(cmbCharBless(x), HeightOnly, DoubleWidth, fraChar(5).hwnd)
    cmbCharBless(x).SelLength = 0
Next x

DoEvents
frmLoad.lblCaption.Caption = "Loading Settings..."
DoEvents
Call LoadSettings
Call InvenCalcEncum
DoEvents

If chkGlobalFilter.Value = 0 Then Call ResetFilterOptions

bDontRefresh = False
Call RefreshAll
GoTo out:

skipload:
bSuppressErrors = True
GoTo out:

out:
bPromptSave = False
bSortOrderAsc = True
bDontRefresh = False
bStartup = False
Set fso = Nothing
Me.Show
Unload frmLoad
DoEvents
Call SetupSplitterSizes

Exit Sub
Error:
Call HandleError("Main_Load")
Resume Next
term:
Set fso = Nothing
bDontCallTerminate = True
bDontSaveSettings = True
Unload Me
End Sub

Private Sub chkCharQuests_Click(Index As Integer)

bPromptSave = True

If Index = 5 Then
    If chkCharQuests(5).Value = 1 Then
        cmbChar2ndAlign.Enabled = True
    Else
        cmbChar2ndAlign.Enabled = False
    End If
End If

Call RefreshAll
End Sub

Private Sub chkGlobalFilter_Click()
Dim x As Integer
On Error GoTo Error:

If chkGlobalFilter.Value = 1 Then
    frmGlobalFilter.Enabled = True
    txtGlobalLevel(1).Enabled = True
    lblClass.Enabled = True
    cmbGlobalClass(1).Enabled = True
    'cmbGlobalClass(1).Enabled = True
    lblAlign.Enabled = True
    cmbGlobalAlignment.Enabled = True
    cmdFilterAll.Enabled = True
    'Call ResetFilterOptions(, , , True)
    'Call cmbGlobalClass_Click(0)
    Call SetupClass
    Call RefreshAll
    
'    For x = 0 To 6
'        chkArmourType(x).Enabled = False
'    Next x
'    For x = 0 To 3
'        chkHanded(x).Enabled = False
'    Next x
'    chkWeaponStaffOnly.Enabled = False
Else
    frmGlobalFilter.Enabled = False
    txtGlobalLevel(1).Enabled = False
    lblClass.Enabled = False
    cmbGlobalClass(1).Enabled = False
    'cmbGlobalClass(1).Enabled = flase
    lblAlign.Enabled = False
    cmbGlobalAlignment.Enabled = False
    cmdFilterAll.Enabled = False
    Call ResetFilterOptions(, , , True)
    
'    For x = 0 To 6
'        chkArmourType(x).Enabled = True
'    Next x
'    For x = 0 To 3
'        chkHanded(x).Enabled = True
'    Next x
'    chkWeaponStaffOnly.Enabled = True
End If

'If Not lvWeapons.SelectedItem Is Nothing Then
'    Call lvWeapons_ItemClick(lvWeapons.SelectedItem)
'End If
'If Not lvWeaponCompare.SelectedItem Is Nothing Then
'    Call lvWeaponCompare_ItemClick(lvWeaponCompare.SelectedItem)
'End If
'
'If Not lvArmour.SelectedItem Is Nothing Then
'    Call lvArmour_ItemClick(lvArmour.SelectedItem)
'End If
'If Not lvArmourCompare.SelectedItem Is Nothing Then
'    Call lvArmourCompare_ItemClick(lvArmourCompare.SelectedItem)
'End If

If Not lvSpells.SelectedItem Is Nothing Then
    Call lvSpells_ItemClick(lvSpells.SelectedItem)
End If
If Not lvSpellCompare.SelectedItem Is Nothing Then
    Call lvSpellCompare_ItemClick(lvSpellCompare.SelectedItem)
End If

Exit Sub

Error:
Call HandleError("chkGlobalFilter_Click")
    
End Sub

Private Sub chkInvenAddWeight_Click()
Call RefreshAll
End Sub

Private Sub chkInvenHideCharStats_Click()
Call RefreshAll
End Sub

Private Sub chkShopShowCharm_Click(Index As Integer)

If Index = 1 Then
    If chkShopShowCharm(0).Value = 1 And chkShopShowCharm(1).Value = 1 Then
        chkShopShowCharm(0).Value = 0
        Exit Sub
    End If
Else
    If chkShopShowCharm(0).Value = 1 And chkShopShowCharm(1).Value = 1 Then
        chkShopShowCharm(1).Value = 0
        Exit Sub
    End If
End If

Call RefreshShops
End Sub

Private Sub chkWeaponStaffOnly_Click()
Dim x As Integer

On Error GoTo Error:

'If chkGlobalFilter.Value = 1 Then Exit Sub

If chkWeaponStaffOnly.Value = 1 Then
    For x = 0 To 3
        chkHanded(x).Value = 1
        chkHanded(x).Enabled = False
    Next
Else
    For x = 0 To 3
        chkHanded(x).Enabled = True
    Next
End If

Exit Sub

Error:
Call HandleError

End Sub

Private Sub cmbCharBless_Click(Index As Integer)

bPromptSave = True
Call RefreshCharBless
End Sub

Private Sub cmbCharBless_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = AutoComplete(cmbCharBless(Index), KeyAscii)
End Sub

Private Sub cmbCharBless_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Dim x As Integer, sText As String
'
'If KeyCode = vbKeyTab Or KeyCode = vbKeyShift Or _
'     KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or _
'     KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or _
'     KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then Exit Sub
'
'If cmbCharBless(Index).ListCount = 0 Then Exit Sub
'
'sText = cmbCharBless(Index).Text
'If sText = "" Then Exit Sub
'
'DoEvents
'For x = nStart To cmbCharBless(Index).ListCount - 1
'    If LCase(Left(cmbCharBless(Index).List(x), Len(sText))) = LCase(sText) Then
'        cmbCharBless(Index).ListIndex = x
'        cmbCharBless(Index).SelStart = Len(sText)
'        cmbCharBless(Index).SelLength = Len(cmbCharBless(Index).Text) - Len(sText)
'        Exit Sub
'    End If
'Next x
End Sub

'Private Sub cmdCharCalcTrueDamage_Click()
'Dim nSwings As Double, nAvgRound As Double
'
'On Error GoTo error:
'
'nSwings = Val(InputBox("Enter your swings in the box below (raw swing number from swing calc can be used) " _
'    & "and have a copy of a megamud player's statistics in your clipboard and press 'OK'.", _
'    "Enter Swings", sCharSwings, Me.Left + cmdCharCalcTrueDamage.Left - 5000, Me.Top + cmdCharCalcTrueDamage.Top - 1000))
'If nSwings <= 0 Then Exit Sub
'
'nAvgRound = CalcTrueAverage(nSwings)
'If nAvgRound > 0 Then
'    MsgBox "True Average Round: " & nAvgRound, vbInformation
'End If
'
'Exit Sub
'error:
'Call HandleError("cmdCharCalcTrueDamage_Click")
'
'End Sub

Private Sub cmbEquip_Click(Index As Integer)

bPromptSave = True

nEquippedItem(Index) = cmbEquip(Index).ItemData(cmbEquip(Index).ListIndex)
Call RefreshAll

End Sub

Private Sub cmbEquip_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = AutoComplete(cmbEquip(Index), KeyAscii)
End Sub

Private Sub cmbGlobalClass_Click(Index As Integer)
Dim x As Integer
On Error GoTo Error:

bPromptSave = True

With cmbGlobalClass
    For x = 0 To .UBound
        If Not x = Index Then
            If Not .item(x).ListIndex = .item(Index).ListIndex Then
                .item(x).ListIndex = .item(Index).ListIndex
                Exit Sub
            End If
        End If
    Next x
End With

Call SetupClass
Call RefreshAll

Exit Sub
Error:
Call HandleError("cmbGlobalClass_Click")
End Sub

Private Sub cmbGlobalRace_Click(Index As Integer)
Dim x As Integer

On Error GoTo Error:

bPromptSave = True

With cmbGlobalRace
    For x = 0 To .UBound
        If Not x = Index Then
            If Not .item(x).ListIndex = .item(Index).ListIndex Then
                .item(x).ListIndex = .item(Index).ListIndex
                Exit Sub
            End If
        End If
    Next x
End With

If cmbGlobalRace(0).ItemData(cmbGlobalRace(0).ListIndex) = 0 Or cmbGlobalRace(0).Text = "(none)" Then
    Call RefreshAll
    Exit Sub
End If

If tabRaces.RecordCount = 0 Then Exit Sub

tabRaces.Index = "pkRaces"
tabRaces.Seek "=", cmbGlobalRace(0).ItemData(cmbGlobalRace(0).ListIndex)
If tabRaces.NoMatch = True Then Exit Sub

txtCharMaxStats(0).Text = tabRaces.Fields("mSTR") & "-" & tabRaces.Fields("xSTR")
txtCharMaxStats(1).Text = tabRaces.Fields("mINT") & "-" & tabRaces.Fields("xINT")
txtCharMaxStats(2).Text = tabRaces.Fields("mWIL") & "-" & tabRaces.Fields("xWIL")
txtCharMaxStats(3).Text = tabRaces.Fields("mAGL") & "-" & tabRaces.Fields("xAGL")
txtCharMaxStats(4).Text = tabRaces.Fields("mHEA") & "-" & tabRaces.Fields("xHEA")
txtCharMaxStats(5).Text = tabRaces.Fields("mCHM") & "-" & tabRaces.Fields("xCHM")

txtCharMaxStats(0).Tag = tabRaces.Fields("mSTR")
txtCharMaxStats(1).Tag = tabRaces.Fields("mINT")
txtCharMaxStats(2).Tag = tabRaces.Fields("mWIL")
txtCharMaxStats(3).Tag = tabRaces.Fields("mAGL")
txtCharMaxStats(4).Tag = tabRaces.Fields("mHEA")
txtCharMaxStats(5).Tag = tabRaces.Fields("mCHM")

For x = 0 To 5
    If Val(txtCharStats(x)) < Val(txtCharMaxStats(x).Tag) Then
        txtCharStats(x) = Val(txtCharMaxStats(x).Tag)
    End If
Next x

Call RefreshAll

Exit Sub

Error:
Call HandleError("cmbGlobalRace_Click")

End Sub

Private Sub cmdCharBlessJump_Click(Index As Integer)
Call GotoSpell(cmbCharBless(Index).ItemData(cmbCharBless(Index).ListIndex))
End Sub

Private Sub cmdCharChangeStats_Click(Index As Integer)

If Not bMouseDown Then Call ModifyCharStats(Index)

End Sub

Private Sub cmdCharChangeStats_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

bMouseDown = True

Do While bMouseDown
    timMouseDown.Enabled = True
    Call ModifyCharStats(Index)
    Do While timMouseDown.Enabled
        DoEvents
    Loop
Loop

End Sub

Private Sub cmdCharChangeStats_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
bMouseDown = False
End Sub

Private Sub cmdCharCopytoClip_Click(Index As Integer)
Dim sCP As String, x As Integer

If Index = 0 Then
    Call CopyChartoClip
ElseIf Index = 1 Then
    Call CopyChartoClip(True)
ElseIf Index = 2 Then
    For x = 0 To 5
        Select Case x
            Case 0:
                sCP = "s" & txtCharStats(x).Text
            Case 1:
                sCP = sCP & ", i" & txtCharStats(x).Text
            Case 2:
                sCP = sCP & ", w" & txtCharStats(x).Text
            Case 3:
                sCP = sCP & ", a" & txtCharStats(x).Text
            Case 4:
                sCP = sCP & ", h" & txtCharStats(x).Text
            Case 5:
                sCP = sCP & ", c" & txtCharStats(x).Text
        End Select
    Next x
    If Not Val(lblStatCalc.Tag) = 0 Then
        sCP = sCP & " (" & Val(lblStatCalc.Tag) & " CP remaining)"
    End If
    
    Clipboard.clear
    Clipboard.SetText sCP
End If
End Sub

Private Sub cmdCharReset_Click()
Dim x As Integer
On Error GoTo Error:

bDontRefresh = True

For x = 0 To 5
    chkCharQuests(x).Value = 0
Next x

cmbGlobalClass(0).ListIndex = 0
cmbGlobalRace(0).ListIndex = 0
txtGlobalLevel(0).Text = 999

txtCharName.Text = ""

For x = 0 To 5
    txtCharStats(x).Text = "0"
    txtCharMaxStats(x).Text = "0"
    txtCharMaxStats(x).Tag = "0"
Next x

For x = 0 To 9
    cmbCharBless(x).ListIndex = 0
Next x

bDontRefresh = False
Call RefreshAll
out:
bDontRefresh = False
Exit Sub
Error:
Call HandleError("cmdCharReset_Click")
Resume out:

End Sub

Private Sub cmdCharResetBless_Click(Index As Integer)
Dim x As Integer, sFile As String, sSectionName As String, nSpell As Long, y As Long

On Error GoTo Error:

bDontRefresh = True
If Index = 0 Then
    For x = 0 To 9
        cmbCharBless(x).ListIndex = 0
    Next x
Else
    sSectionName = RemoveCharacter(lblDatVer.Caption, " ")
    If bCharLoaded Then
        sFile = ReadINI(sSectionName, "LastCharFile")
        If Not FileExists(sFile) Then
            sFile = ""
        Else
            sSectionName = "Bless"
        End If
    End If
    
    For x = 0 To 9
        nSpell = ReadINI(sSectionName, "Bless" & x, sFile)
        If nSpell > 0 Then
            For y = 0 To cmbCharBless(x).ListCount - 1
                If cmbCharBless(x).ItemData(y) = nSpell Then
                    cmbCharBless(x).ListIndex = y
                    Exit For
                End If
            Next y
        Else
            cmbCharBless(x).ListIndex = 0
        End If
    Next x
End If

cmbCharBless(0).SetFocus
bDontRefresh = False
Call RefreshCharBless
    
Exit Sub
Error:
Call HandleError("cmdCharResetBless_Click")
bDontRefresh = False
End Sub

Private Sub cmdClassRace_Click(Index As Integer)

If Index = 0 Then
    lvClasses.Visible = True
    txtClassDetail.Visible = True
    lvRaces.Visible = False
    txtRaceDetail.Visible = False
    cmdClassRace(0).BackColor = &HC0C0FF
    cmdClassRace(1).BackColor = &H8000000F
Else
    cmdClassRace(0).BackColor = &H8000000F
    cmdClassRace(1).BackColor = &HC0C0FF
    lvClasses.Visible = False
    txtClassDetail.Visible = False
    lvRaces.Visible = True
    txtRaceDetail.Visible = True
End If

End Sub

Private Sub cmdClearAllCompares_Click()
Dim x As Integer

For x = 0 To cmdCompareClear.UBound
    Call cmdCompareClear_Click(x)
Next x

End Sub

Private Sub cmdCompareClear_Click(Index As Integer)

Select Case Index
    Case 0:
        lvWeaponCompare.ListItems.clear
        txtWeaponCompareDetail.Text = ""
        lvWeaponCompareLoc.ListItems.clear
    Case 1:
        lvArmourCompare.ListItems.clear
        txtArmourCompareDetail.Text = ""
        lvArmourCompareLoc.ListItems.clear
    Case 2:
        lvSpellCompare.ListItems.clear
        txtSpellCompareDetail.Text = ""
        lvSpellCompareLoc.ListItems.clear
    Case 3:
        lvMonsterCompare.ListItems.clear
        'txtMonsterCompareDetail.Text = ""
        lvMonsterCompareLoc.ListItems.clear
End Select

End Sub

Private Sub cmdCompareNav_Click(Index As Integer)
Dim x As Integer

On Error GoTo Error:

For x = 0 To 3
    If x = Index Then
        framCompareNav(Index).Visible = True
        cmdCompareNav(x).BackColor = &HC0C0FF
    Else
        cmdCompareNav(x).BackColor = &H8000000F
        framCompareNav(x).Visible = False
    End If
Next x

Exit Sub

Error:
Call HandleError

End Sub

Private Sub cmdCompareQ_Click()
MsgBox "Right click on a weapon, piece of armour, spell, or monster to add it to the compare lists.", vbInformation
End Sub

Private Sub cmdDrawMap_Click(Index As Integer)

If Index = 0 Then
    If Val(txtRoomMap.Text) > 32767 Then txtRoomMap.Text = 32767
    If Val(txtRoomRoom.Text) > 32767 Then txtRoomRoom.Text = 32767
    Call MapStartMapping(Val(txtRoomMap.Text), Val(txtRoomRoom.Text))
Else
    Call MapStartMapping(nMapStartMap, nMapStartRoom)
End If


End Sub

Private Sub cmdEditPreset_Click(Index As Integer)
Dim sSectionName As String, bRestoreMap As Boolean
Dim cReg As clsRegistryRoutines
On Error GoTo Error:
Set cReg = New clsRegistryRoutines

If bLegit Then
    sSectionName = "Custom_Presets"
Else
    sSectionName = RemoveCharacter(lblDatVer.Caption, " ") & "_Presets"
End If
'sSectionName = RemoveCharacter(lblDatVer.Caption, " ") & "_Presets"

cReg.hkey = HKEY_LOCAL_MACHINE
cReg.KeyRoot = "Software\MMUD Explorer\Presets"
cReg.Subkey = sSectionName

If FormIsLoaded("frmMap") Then
    If frmMap.WindowState = vbNormal Then
        bRestoreMap = True
        frmMap.WindowState = vbMinimized
    End If
End If

Unload frmEditPreset
Load frmEditPreset
frmEditPreset.nPreset = Val(cmdMapPreset(Index).Tag)
frmEditPreset.lblCaption.Caption = "Editing Preset #" & (cmdMapPreset(Index).Tag + 1)
frmEditPreset.txtMap.Text = cReg.GetRegistryValue("Map" & cmdMapPreset(Index).Tag, 0) 'ReadINI(sSectionName, "Map" & cmdMapPreset(index).Tag)
frmEditPreset.txtRoom.Text = cReg.GetRegistryValue("Room" & cmdMapPreset(Index).Tag, 0) 'ReadINI(sSectionName, "Room" & cmdMapPreset(index).Tag)
frmEditPreset.txtCaption.Text = cReg.GetRegistryValue("Name" & cmdMapPreset(Index).Tag, "unset") 'ReadINI(sSectionName, "Name" & cmdMapPreset(index).Tag)
Set frmEditPreset.objFormOwner = Me
DoEvents
frmEditPreset.Show vbModal, Me

If Not frmEditPreset.nPreset < 0 Then
    Call cReg.SetRegistryValue("Map" & cmdMapPreset(Index).Tag, frmEditPreset.txtMap.Text, REG_SZ)
    Call cReg.SetRegistryValue("Room" & cmdMapPreset(Index).Tag, frmEditPreset.txtRoom.Text, REG_SZ)
    Call cReg.SetRegistryValue("Name" & cmdMapPreset(Index).Tag, frmEditPreset.txtCaption.Text, REG_SZ)
'    Call WriteINI(sSectionName, "Map" & cmdMapPreset(index).Tag, frmEditPreset.txtMap.Text)
'    Call WriteINI(sSectionName, "Room" & cmdMapPreset(index).Tag, frmEditPreset.txtRoom.Text)
'    Call WriteINI(sSectionName, "Name" & cmdMapPreset(index).Tag, frmEditPreset.txtCaption.Text)
    
    cmdMapPreset(Index).Caption = frmEditPreset.txtCaption.Text
End If

Unload frmEditPreset

If FormIsLoaded("frmMap") Then Call frmMap.LoadPresets

If bRestoreMap Then frmMap.WindowState = vbNormal

Exit Sub
Error:
Call HandleError("cmdEditPreset_Click")
Unload frmEditPreset
End Sub

Private Sub cmdEquipCheckAll_Click(Index As Integer)
Dim x As Integer, nSet As Integer

If Index = 0 Then nSet = 1 Else nSet = 0

For x = 0 To 18
    chkEquipHold(x).Value = nSet
Next x

End Sub

Private Sub cmdEquipGoto_Click(Index As Integer)
nLastEquipIndex = Index

If Index = 16 Then
    mnuEquipGotoItem(2).Visible = True
    mnuEquipGotoItem(3).Visible = True
    PopupMenu mnuEquipGoto, , , cmdEquipGoto(Index).Top + 100
Else
    mnuEquipGotoItem(2).Visible = False
    mnuEquipGotoItem(3).Visible = False
    PopupMenu mnuEquipGoto
End If

End Sub

Private Sub cmdFilter_Click(Index As Integer)

Me.MousePointer = vbHourglass
DoEvents

Select Case Index
    Case 0: 'filter inven
        Call FilterInvenItems(False)
    Case 1: 'remove inven
        Call FilterInvenItems(True)
    Case 2: 'filter spells
        lvSpellLoc.ListItems.clear
        txtSpellDetail.Text = ""
        lvSpellCompareLoc.ListItems.clear
        txtSpellCompareDetail.Text = ""
        If chkGlobalFilter.Value = 1 Then
            Call FilterSpells(True)
        Else
            Call FilterSpells(False)
        End If
    Case 3: 'remove spells
        lvSpellLoc.ListItems.clear
        txtSpellDetail.Text = ""
        lvSpellCompareLoc.ListItems.clear
        txtSpellCompareDetail.Text = ""
        Call ResetFilterOptions(True, True, False, False)
        Call FilterSpells(False)
        If chkGlobalFilter.Value = 1 Then
            Call SetupClass
            Call RefreshAll
        End If
    Case 4: 'Filter armour
        lvArmourLoc.ListItems.clear
        txtArmourDetail.Text = ""
        lvArmourCompareLoc.ListItems.clear
        txtArmourCompareDetail.Text = ""
        If chkGlobalFilter.Value = 1 Then
            Call FilterArmour(True)
        Else
            Call FilterArmour(False)
        End If
    Case 5: 'remove armour
        lvArmourLoc.ListItems.clear
        txtArmourDetail.Text = ""
        lvArmourCompareLoc.ListItems.clear
        txtArmourCompareDetail.Text = ""
        Call ResetFilterOptions(False, True, True, False)
        Call FilterArmour(False)
        If chkGlobalFilter.Value = 1 Then
            Call SetupClass
            Call RefreshAll
        End If
    Case 6: 'filter weapons
        lvWeaponLoc.ListItems.clear
        txtWeaponDetail.Text = ""
        lvWeaponCompareLoc.ListItems.clear
        txtWeaponCompareDetail.Text = ""
        If chkGlobalFilter.Value = 1 Then
            Call FilterWeapons(True)
        Else
            Call FilterWeapons(False)
        End If
    Case 7: 'remove weapons
        lvWeaponLoc.ListItems.clear
        txtWeaponDetail.Text = ""
        lvWeaponCompareLoc.ListItems.clear
        txtWeaponCompareDetail.Text = ""
        Call ResetFilterOptions(True, False, True, False)
        Call FilterWeapons(False)
        If chkGlobalFilter.Value = 1 Then
            Call SetupClass
            Call RefreshAll
        End If
    Case 8:
    Case 9:
End Select

Me.MousePointer = vbDefault
DoEvents

End Sub

Private Sub cmdFilterAll_Click()
On Error GoTo Error:

Call FilterAll(True)

Exit Sub
Error:
Call HandleError
End Sub

Private Sub cmdFind_Click(Index As Integer)
Dim bFound As Boolean

Select Case Index
    Case 0, 1: 'spell
        bFound = SearchLV(IIf(Index = 0, 0, vbKeyRight), lvSpells, txtSpellFind)
        If bFound Then Call lvSpells_ItemClick(lvSpells.SelectedItem)
        
    Case 2, 3: 'armour find
        bFound = SearchLV(IIf(Index = 2, 0, vbKeyRight), lvArmour, txtArmourFind)
        If bFound Then Call lvArmour_ItemClick(lvArmour.SelectedItem)
    
    Case 4, 5: 'weapons find
        bFound = SearchLV(IIf(Index = 4, 0, vbKeyRight), lvWeapons, txtWeaponFind)
        If bFound Then Call lvWeapons_ItemClick(lvWeapons.SelectedItem)
    
    Case 6, 7: 'shops
        bFound = SearchLV(IIf(Index = 6, 0, vbKeyRight), lvShops, txtShopFind)
        If bFound Then Call lvShops_ItemClick(lvShops.SelectedItem)
    
    Case 8, 9: 'monsters
        bFound = SearchLV(IIf(Index = 8, 0, vbKeyRight), lvMonsters, txtMonsterFind)
        If bFound Then Call lvMonsters_ItemClick(lvMonsters.SelectedItem)
        
    Case 10, 11: 'other items
        bFound = SearchLV(IIf(Index = 10, 0, vbKeyRight), lvOtherItems, txtOtherItemsFind)
        If bFound Then Call lvOtherItems_ItemClick(lvOtherItems.SelectedItem)

End Select

End Sub

Private Sub cmdInvenAdditionalWeight_Click()
Dim sStr As String, nEncum As Integer, x As Integer

On Error GoTo Error:

sStr = InputBox("Enter your reported current encumbrance " _
    & "value and it will calculate the difference from the equipment " _
    & "items you've selected.  Alternatively, you can paste a text capture of your " _
    & "'inventory' output, including the ""Ecumbrance:"" line with the 'Paste Character' button.", _
    "Enter total encumbrance", sAddWeight)
If Val(sStr) = 0 Then Exit Sub

If Val(sStr) > 1000000 Then sStr = 1000000

sAddWeight = Val(sStr)

For x = 0 To UBound(nEquippedItem())
    If nEquippedItem(x) > 0 Then
        nEncum = nEncum + GetItemWeight(nEquippedItem(x))
    End If
Next x

If Val(sStr) > nEncum Then
    txtInvenAddWeight.Text = Val(sStr) - nEncum
End If

Exit Sub
Error:
Call HandleError("cmdInvenAdditionalWeight_Click")
End Sub

Private Sub cmdInvenClipboard_Click()
Call InvenCopytoClipboard
End Sub

Private Sub cmdInvenCompareAll_Click()
Dim x As Integer, nResult As Integer

For x = 0 To UBound(nEquippedItem())
    If nEquippedItem(x) > 0 Then
         nResult = CompareAddItem(nEquippedItem(x))
         If nResult = vbCancel Then Exit Sub
    End If
Next x

End Sub

Private Sub cmdInvenEmptyLists_Click(Index As Integer)

If Index = 1 Then
    MsgBox "When you do a Find Best, it searches through the the items that are currently listed in each of " _
        & "the equipment slots.  What you can do with this is empty the slots and then create your own list of items " _
        & "by right clicking on an item and choosing 'Equip Item'.  After you have all of the items you want in the lists, " _
        & "you can then do a Find Best on those items.  To restore the full lists, either apply or remove the global filter.  " _
        & "If you wanted to save these lists, you could add all of your items to the compare lists first and then equip them " _
        & "since the compare lists are saved on a character save.", vbInformation
    Exit Sub
End If

bDontRefresh = True
Call InvenSetupEquip
bDontRefresh = False
Call RefreshAll
End Sub

Private Sub cmdInvenFindBest_Click()
bInvenNextBest = False
PopupMenu mnuFindBest
End Sub

Private Sub cmdInvenModStrength_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
bMouseDown = True

Do While bMouseDown
    timMouseDown.Enabled = True
    If Index = 0 Then
        txtCharStats(0).Text = Val(txtCharStats(0).Text) - 1
    Else
        txtCharStats(0).Text = Val(txtCharStats(0).Text) + 1
    End If
    Do While timMouseDown.Enabled
        DoEvents
    Loop
Loop
End Sub

Private Sub cmdInvenModStrength_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
bMouseDown = False
End Sub

Private Sub cmdInvenNextBest_Click()
bInvenNextBest = True
Call InvenFindBest(nInvenLastIndex(0), nInvenLastIndex(1))
'PopupMenu mnuFindBest
End Sub

Private Sub cmdInvenPasteChar_Click(Index As Integer)
Call PasteCharacter
End Sub

Private Sub cmdInvenReset_Click()

Call InvenClear

End Sub

Private Sub cmdInvenSetStatFont_Click()
On Error GoTo Error:
Dim x As Integer

oComDag.FLAGS = cdlCFBoth Or cdlCFForceFontExist
oComDag.FontName = txtStat(0).Font.name
oComDag.FontSize = txtStat(0).Font.Size
oComDag.FontBold = txtStat(0).Font.Bold
oComDag.FontItalic = txtStat(0).Font.Italic

' Display the Font dialog box
On Error GoTo canceled:
oComDag.ShowFont
On Error GoTo Error:

Call SetStatFonts(oComDag.FontName, oComDag.FontSize, oComDag.FontBold, oComDag.FontItalic)

canceled:
Exit Sub
Error:
HandleError
End Sub

Private Sub cmdMapAddMegaCodes_Click()
On Error GoTo Error:
Dim sText As String, sNewText As String, x As Long, y As Long, nSteps As Long

sText = txtMapMove.Text
If Len(sText) < 1 Then
    MsgBox "No text to copy.", vbInformation
    Exit Sub
End If

sText = "FFFFFFFF:0080:" & sText
nSteps = 1
y = 1
Do While y < Len(sText)
    'x = InStr(x, sText, Chr(10))
    x = InStr(y, sText, Chr(13))
    If x > 0 Then
       ' Debug.Print "--begin---"
       ' Debug.Print Left(sText, x - 1)
       ' Debug.Print Right(sText, Len(sText) - x - 1)
       ' Debug.Print "---end----"
        sText = Left(sText, x + 1) & "FFFFFFFF:0080:" & Right(sText, Len(sText) - x - 1)
        nSteps = nSteps + 1
        y = x + 2
    Else
        Exit Do
    End If
    DoEvents
Loop

sText = "[LOOP NAME][AUTHOR]" & vbCrLf _
    & "[CODE:FROM ROOM GROUP:FROM ROOM NAME]" & vbCrLf _
    & "[CODE:TO ROOM GROUP:TO ROOM NAME]" & vbCrLf _
    & "START ROOM CHECKSUM:END ROOM CHECKSUM:" & nSteps & ":-1:0:Needed Items:Path fails:when finished" _
    & vbCrLf & sText

Clipboard.clear
Clipboard.SetText sText

MsgBox "Copied " & nSteps & " steps to the clipboard.", vbInformation

Exit Sub

Error:
Call HandleError("cmdMapAddMegaCodes_Click")
End Sub

Private Sub cmdMapCopyToClip_Click()
On Error GoTo Error:
Dim sText As String

sText = txtMapMove.Text
If Len(sText) < 1 Then
    MsgBox "No text to copy.", vbInformation
    Exit Sub
End If

Clipboard.clear
Clipboard.SetText sText

Exit Sub

Error:
Call HandleError("cmdMapCopyToClip_Click")
End Sub

Private Sub cmdMapFindText_Click(Index As Integer)
On Error GoTo Error:
Dim sTemp As String

If tabRooms.RecordCount = 0 Then Exit Sub

tabRooms.Index = "idxRooms"
If Index = 0 Or nMapLastFind(0) = 0 Or nMapLastFind(1) = 0 Then
    sTemp = InputBox("Enter text to search for.", "Search for room name", sMapSearch)
    If sTemp = "" Then Exit Sub
    
    sMapSearch = sTemp
    nMapLastFind(2) = 0
    tabRooms.MoveFirst
Else
    tabRooms.Seek "=", nMapLastFind(0), nMapLastFind(1)
    If tabRooms.NoMatch Then
        MsgBox "Room " & nMapLastFind(0) & "/" & nMapLastFind(1) & " not found.", vbInformation
        Exit Sub
    End If
    tabRooms.MoveNext
End If
DoEvents

'Call UnloadForms("frmMain")
Me.Enabled = False
If FormIsLoaded("frmMap") Then frmMap.Enabled = False

bMapCancelFind = False
Load frmProgressBar
Call frmProgressBar.SetRange(tabRooms.RecordCount)
frmProgressBar.ProgressBar.Value = nMapLastFind(2)
frmProgressBar.lblCaption.Caption = "Searching for Room Name ..."
Set frmProgressBar.objFormOwner = Me

DoEvents
frmProgressBar.Show vbModeless, Me
DoEvents

Do Until tabRooms.EOF Or bMapCancelFind
    If InStr(1, LCase(tabRooms.Fields("Name")), LCase(sMapSearch)) > 0 Then Exit Do
    Call frmProgressBar.IncreaseProgress
    tabRooms.MoveNext
    DoEvents
Loop
If tabRooms.EOF Then
    nMapLastFind(0) = 0
    nMapLastFind(1) = 0
    nMapLastFind(2) = 0
    MsgBox "Name not found.", vbInformation
    GoTo out:
End If

nMapLastFind(0) = tabRooms.Fields("Map Number")
nMapLastFind(1) = tabRooms.Fields("Room Number")
nMapLastFind(2) = frmProgressBar.ProgressBar.Value

If Not bMapCancelFind Then
    Call MapStartMapping(tabRooms.Fields("Map Number"), tabRooms.Fields("Room Number"))
End If

out:
Unload frmProgressBar
Me.Enabled = True
If FormIsLoaded("frmMap") Then frmMap.Enabled = True
frmMain.SetFocus
Exit Sub

Error:
Call HandleError("cmdMapFindText_Click")
Resume out:
End Sub

Private Sub cmdMapLastRoom_Click()
On Error GoTo Error:

Call MapStartMapping(nMapLastMap, nMapLastRoom)

out:
Exit Sub
Error:
Call HandleError("cmdMapLastRoom_Click")
Resume out:
End Sub

Private Sub cmdMapMoveClear_Click()
Dim nYesNo As Integer
nYesNo = MsgBox("Are you sure you want to clear the box?", vbYesNo + vbDefaultButton2 + vbQuestion)
If nYesNo = vbYes Then txtMapMove.Text = ""
End Sub

Private Sub cmdMapOpenExternal_Click()
On Error Resume Next
Load frmMap
If nMapStartMap > 0 And nMapStartRoom > 0 Then
    Call frmMap.MapStartMapping(nMapStartMap, nMapStartRoom)
End If
frmMap.Show
Me.WindowState = vbMinimized
frmMap.SetFocus
End Sub

Private Sub cmdMapPreset_Click(Index As Integer)
Dim nMap As Long, nRoom As Long, sSectionName As String
Dim cReg As clsRegistryRoutines
On Error GoTo Error:

Set cReg = New clsRegistryRoutines

If bLegit Then
    sSectionName = "Custom_Presets"
Else
    sSectionName = RemoveCharacter(lblDatVer.Caption, " ") & "_Presets"
End If

cReg.hkey = HKEY_LOCAL_MACHINE
cReg.KeyRoot = "Software\MMUD Explorer\Presets"
cReg.Subkey = sSectionName

nMap = cReg.GetRegistryValue("Map" & cmdMapPreset(Index).Tag, 0) 'Val(ReadINI(sSectionName, "Map" & cmdMapPreset(index).Tag))
nRoom = cReg.GetRegistryValue("Room" & cmdMapPreset(Index).Tag, 0) 'Val(ReadINI(sSectionName, "Room" & cmdMapPreset(index).Tag))

Call MapStartMapping(nMap, nRoom)

out:
Exit Sub
Error:
Call HandleError("cmdMapPreset_Click")
Resume out:

End Sub

Private Sub cmdMapPresetSelect_Click(Index As Integer)
Dim nStart As Integer, x As Integer, sSectionName As String
Dim cReg As clsRegistryRoutines

Set cReg = New clsRegistryRoutines

If bLegit Then
    sSectionName = "Custom_Presets"
Else
    sSectionName = RemoveCharacter(lblDatVer.Caption, " ") & "_Presets"
End If

cReg.hkey = HKEY_LOCAL_MACHINE
cReg.KeyRoot = "Software\MMUD Explorer\Presets"
cReg.Subkey = sSectionName

Select Case Index
    Case 0: nStart = 0
    Case 1: nStart = 10
    Case 2: nStart = 20
    Case 3: nStart = 30
    Case 4: nStart = 40
    Case Else: Exit Sub
End Select

For x = nStart To nStart + 9
    cmdMapPreset(x Mod 10).Caption = cReg.GetRegistryValue("Name" & x, "unset") 'ReadINI(sSectionName, "Name" & x)
    cmdMapPreset(x Mod 10).Tag = x
Next x

End Sub

Private Sub cmdMapShowUnused_Click()
Dim x As Integer

If cmdMapShowUnused.Caption = "S&how Blocks" Then
    For x = 1 To 400
        lblRoomCell(x).Visible = True
    Next
    cmdMapShowUnused.Caption = "&Hide Blocks"
Else
    For x = 1 To 400
        If CellRoom(x, 1) = 0 Then lblRoomCell(x).Visible = False
    Next
    cmdMapShowUnused.Caption = "S&how Blocks"
End If

End Sub

Private Sub cmdMapSwitch_Click()
If cmdMapSwitch.Tag = "1" Then
    cmdMapSwitch.Tag = ""
    cmdMapSwitch.Caption = "&Switch to Manual Edit"
Else
    cmdMapSwitch.Tag = "1"
    cmdMapSwitch.Caption = "&Switch to Map Move"
End If
txtMapMove.SetFocus
End Sub

Private Sub cmdMonHelp_Click()
MsgBox "You can right click on a monster to copy it to the" & vbCrLf _
    & "clipboard or look up how it is summoned.", vbInformation
End Sub

Public Sub cmdNav_Click(Index As Integer)
Dim x As Integer

On Error GoTo Error:

For x = 0 To cmdNav().UBound
    If x = Index Then
        framNav(Index).Visible = True
        cmdNav(x).BackColor = &HFFC0C0
    Else
        cmdNav(x).BackColor = &H8000000F
        framNav(x).Visible = False
    End If
Next x

If Me.WindowState = vbMinimized Then Me.WindowState = nWindowState

If bDontSetMainFocus = True Then Exit Sub
cmdNav(Index).SetFocus
Select Case Index
    Case 0: 'items
        txtWeaponFind.SetFocus
        'Call SelectAll(txtWeaponFind)
    Case 1: 'armour
        txtArmourFind.SetFocus
        'Call SelectAll(txtArmourFind)
    Case 2: 'spells
        txtSpellFind.SetFocus
        'Call SelectAll(txtSpellFind)
    Case 3: 'compare
        cmdCompareNav(0).SetFocus
    Case 4: 'inven
        cmdFilter(0).SetFocus
    Case 5: 'shops
        txtCharName.SetFocus
    Case 6: 'class/race
        cmdClassRace(0).SetFocus
    Case 7: 'sundry
        txtOtherItemsFind.SetFocus
    Case 8: 'monsters
        txtMonsterFind.SetFocus
        'Call SelectAll(txtMonsterFind)
    Case 9: 'shops
        txtShopFind.SetFocus
        'Call SelectAll(txtShopFind)
        'Call SelectAll(txtOtherItemsFind)
    Case 10: 'rooms
        txtRoomMap.SetFocus
        
End Select

Exit Sub

Error:
Call HandleError

End Sub

'Private Sub Form_Load()
' 'JUST SAMPLE FOR DEMO PURPOSES
' 'YOU CAN OBVIOUSLY ADD WHAT YOU NEED
' 'TO ADD HERE
'    With Combo1
'        .AddItem "Blue"
'        .AddItem "Green"
'        .AddItem "Yellow"
'
'        .ListIndex = 0
'        .Text = .List(0)
'        .SelStart = 0
'        .SelLength = Len(.Text)
'        msOldString = .Text
'        miStart = 0
'        miLength = .SelLength
'    End With
'End Sub
'End Sub

Private Sub cmdResetFonts_Click()
On Error GoTo Error:

Call SetStatFonts("Terminal", 9, False, False)
Call SetLabelFonts("Terminal", 9, False, False)

Exit Sub
Error:
Call HandleError("Reset Fonts")
End Sub

Private Sub cmdResetPresets_Click()
Dim nYesNo As Integer

nYesNo = MsgBox("Are you sure you want to reset the presets to the default set?", vbYesNo + vbDefaultButton2 + vbQuestion, "Reset Presets?")

If nYesNo = vbYes Then Call LoadPresets(True)

End Sub

Private Sub cmdSetInvenLabelFont_Click()
On Error GoTo Error:
Dim x As Integer

oComDag.FLAGS = cdlCFBoth Or cdlCFForceFontExist
oComDag.FontName = lblInvenStats(2).Font.name
oComDag.FontSize = lblInvenStats(2).Font.Size
oComDag.FontBold = lblInvenStats(2).Font.Bold
oComDag.FontItalic = lblInvenStats(2).Font.Italic

' Display the Font dialog box
On Error GoTo canceled:
oComDag.ShowFont
On Error GoTo Error:

Call SetLabelFonts(oComDag.FontName, oComDag.FontSize, oComDag.FontBold, oComDag.FontItalic)

canceled:
Exit Sub
Error:
HandleError
End Sub

Private Sub cmdShopAlterCharm_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
bMouseDown = True

Do While bMouseDown
    timMouseDown.Enabled = True
    If Index = 0 Then
        If Val(txtShopCharm.Text) - 1 < 0 Then Exit Sub
        txtShopCharm.Text = Val(txtShopCharm.Text) - 1
    Else
        txtShopCharm.Text = Val(txtShopCharm.Text) + 1
    End If
    Do While timMouseDown.Enabled
        DoEvents
    Loop
Loop
End Sub

Private Sub cmdShopAlterCharm_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
bMouseDown = False
End Sub

Private Sub cmdShopConvert_Click()

Call MsgBox("100 platinum pieces == 1 runic coin" _
            & vbCrLf & "100 gold crowns == 1 platinum piece" _
            & vbCrLf & "10 silver nobles == 1 gold crown" _
            & vbCrLf & "10 copper farthings == 1 silver noble" _
            & vbCrLf & "" _
            & vbCrLf & "e.g." _
            & vbCrLf & "" _
            & vbCrLf & "1 runic coin == 1,000,000 copper" _
            & vbCrLf & "1 platinum pieces == 10,000 copper" _
            & vbCrLf & "1 gold crowns == 100 copper" _
            & vbCrLf & "1 silver nobles == 10 copper" _
            , vbInformation Or vbDefaultButton1, "The currency conversion rates are:")

End Sub

Private Sub cmdSundryChests_Click()

On Error GoTo Error:
Dim x As Long, y As Long, nTBNumber As Long, sData As String, nNest As Long
Dim nChestItems() As Currency, nDataPos As Long, oLI As ListItem

If lvOtherItems.SelectedItem Is Nothing Then Exit Sub

tabItems.Index = "pkItems"
tabItems.Seek "=", Val(lvOtherItems.SelectedItem.Text)
If tabItems.NoMatch Then
    MsgBox "Item #" & Val(lvOtherItems.SelectedItem.Text) & " not found.", vbExclamation
    Exit Sub
End If

If Not tabItems.Fields("ItemType") = 8 Then
    MsgBox tabItems.Fields("Name") & " is not a container.", vbExclamation
    Exit Sub
End If

For x = 0 To 19
    If tabItems.Fields("Abil-" & x) = 43 And tabItems.Fields("AbilVal-" & x) > 0 Then
        tabSpells.Index = "pkSpells"
        tabSpells.Seek "=", tabItems.Fields("AbilVal-" & x)
        If Not tabSpells.NoMatch Then
            For y = 0 To 9
                If tabSpells.Fields("Abil-" & y) = 148 Then 'castsp
                    If tabSpells.Fields("AbilVal-" & y) = 0 Then
                        If tabSpells.Fields("MinBase") > 0 Then
                            nTBNumber = tabSpells.Fields("MinBase")
                        Else
                            nTBNumber = tabSpells.Fields("MaxBase")
                        End If
                    Else
                        nTBNumber = tabSpells.Fields("AbilVal-" & y)
                    End If
                    
                    tabTBInfo.Index = "pkTBInfo"
                    tabTBInfo.Seek "=", nTBNumber
                    If Not tabTBInfo.NoMatch Then
                        If Not tabTBInfo.Fields("Action") = Chr(0) Then
                            sData = LCase(tabTBInfo.Fields("Action"))
                            GoTo dig:
                        End If
                    End If
                End If
            Next y
        End If
    End If
Next x
MsgBox "Failed to find chest data.", vbExclamation
GoTo out:

dig:
ReDim nChestItems(1 To 3, 0) '1=number, 2=percent, 3=percent failure
nDataPos = 1

Do While InStr(nDataPos, sData, "random ") > 0
    nDataPos = InStr(nDataPos, sData, "random ") + Len("random ")
    
    For x = nDataPos To Len(sData)
        Select Case Mid(sData, x, 1)
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
            Case Else: Exit For
        End Select
    Next x
    
    If x > nDataPos Then
        nTBNumber = Val(Mid(sData, nDataPos, x - nDataPos))
        Call GetChestItems(nChestItems(), nTBNumber, nNest)
    End If
Loop

If UBound(nChestItems(), 2) > 0 Then
    lvOtherItemLoc.ListItems.clear
    For x = 0 To UBound(nChestItems(), 2)
        If nChestItems(1, x) > 0 Then
            Set oLI = lvOtherItemLoc.ListItems.Add
            oLI.Text = "Item: " & GetItemName(nChestItems(1, x), bHideRecordNumbers) _
                & " - " & Round(nChestItems(2, x) * 100, 1) & "%"
            oLI.Tag = nChestItems(1, x)
        End If
    Next x
    Call SortListView(lvOtherItemLoc, 1, ldtstring, True)
End If

out:
On Error Resume Next
Erase nChestItems()
Set oLI = Nothing
Exit Sub

Error:
Call HandleError("cmdSundryChests_Click")
Resume out:
End Sub

Private Sub cmdViewMapLegend_Click()

If cmdViewMapLegend.Tag = "1" Then
    Unload frmMapLegend
    cmdViewMapLegend.Tag = "0"
Else
    cmdViewMapLegend.Tag = "1"
    frmMapLegend.Show False, Me
    Set frmMapLegend.objFormOwner = Me
End If

End Sub

Private Sub Command1_Click()
''''''''' CONTROL COUNT
'''''
'''''Dim objTemp As Control, x As Integer
'''''On Error GoTo Error:
'''''
'''''Dim sFile As String, fso As FileSystemObject, ts As TextStream
'''''
'''''sFile = "c:\out.txt"
'''''
'''''Set fso = CreateObject("Scripting.FileSystemObject")
'''''Set ts = fso.CreateTextFile(sFile, True)
'''''
'''''For Each objTemp In frmMain.Controls
'''''    If Not objTemp.index > 0 Then
'''''        x = x + 1
'''''        ts.WriteLine objTemp.name
'''''    End If
'''''nextx:
'''''
'''''    Set objTemp = Nothing
'''''Next
'''''ts.WriteLine x
'''''
'''''ts.Close
'''''
'''''Set ts = Nothing
'''''Set fso = Nothing
'''''
''''''''' /CONTROL COUNT
'''''
''''''''''''Dim qdfQuery As QueryDef, strSQL As String, rstTemp As Recordset
''''''''''''
''''''''''''On Error GoTo error:
''''''''''''
''''''''''''strSQL = "SELECT * FROM [Items] WHERE " & SQL_NumArray("Ability", "OR", 19, 0, "=", 4)
''''''''''''
''''''''''''Debug.Print strSQL
''''''''''''
''''''''''''Set qdfQuery = DB.CreateQueryDef("", strSQL)
''''''''''''Set rstTemp = qdfQuery.OpenRecordset(dbOpenSnapshot)
''''''''''''
''''''''''''If rstTemp.RecordCount > 0 Then rstTemp.MoveLast
''''''''''''MsgBox rstTemp.RecordCount
''''''''''''
''''''''''''quit:
'''''''''''''On Error Resume Next
''''''''''''rstTemp.Close
''''''''''''qdfQuery.Close
''''''''''''
''''''''''''Exit Sub
''''''''''''Error:
''''''''''''Call HandleError("Command1_Click")
''''''''''''Resume quit:
'''''
''''''''' CONTROL COUNT
'''''Exit Sub
'''''
'''''Error:
'''''x = x + 1
'''''ts.WriteLine objTemp.name
'''''Resume nextx:
'''''Call HandleError("Command1_Click")
''''''''' /CONTROL COUNT
End Sub

Private Function CompareAddItem(ByVal nNum As Long) As Integer
Dim oLI As ListItem, nYesNo As Integer

On Error GoTo Error:

tabItems.Index = "pkItems"
tabItems.Seek "=", nNum
If tabItems.NoMatch = True Then
    'MsgBox "Record " & nNum & " not found."
    Exit Function
End If


Select Case tabItems.Fields("ItemType")
    Case 0: 'armour
        bPromptSave = True
        Set oLI = lvArmourCompare.FindItem(nNum, lvwText, , 0)
        If Not oLI Is Nothing Then
            nYesNo = MsgBox(tabItems.Fields("Name") _
                & " is already on the compare list, add it again?", _
                vbYesNoCancel + vbDefaultButton2 + vbQuestion)
            If nYesNo = vbNo Then
                GoTo quit:
            ElseIf nYesNo = vbCancel Then
                CompareAddItem = vbCancel
                GoTo quit:
            End If
        End If
        Call AddArmour2LV(lvArmourCompare)
        If txtArmourCompareDetail.Text = "" Then
            Call lvArmourCompare_ItemClick(lvArmourCompare.ListItems(1))
        End If
        
        If mnuJumpToCompare.Checked Then
            Call cmdNav_Click(3) 'compare
            Call cmdCompareNav_Click(1)
        End If
    Case 1: 'weapon
        bPromptSave = True
        Set oLI = lvWeaponCompare.FindItem(nNum, lvwText, , 0)
        If Not oLI Is Nothing Then
            nYesNo = MsgBox(tabItems.Fields("Name") _
                & " is already on the compare list, add it again?", _
                vbYesNoCancel + vbDefaultButton2 + vbQuestion)
            If nYesNo = vbNo Then
                GoTo quit:
            ElseIf nYesNo = vbCancel Then
                GoTo quit:
            End If
        End If

        Call AddWeapon2LV(lvWeaponCompare)
        If txtWeaponCompareDetail.Text = "" Then
            Call lvWeaponCompare_ItemClick(lvWeaponCompare.ListItems(1))
        End If
        
        If mnuJumpToCompare.Checked Then
            Call cmdNav_Click(3) 'compare
            Call cmdCompareNav_Click(0)
        End If
End Select

quit:
Set oLI = Nothing

Exit Function
Error:
Call HandleError("CompareAddItem")
Set oLI = Nothing
End Function

Private Function CompareAddMonster(ByVal nNum As Long, Optional ByVal bNoPrompt As Boolean) As Integer
Dim oLI As ListItem, nYesNo As Integer

On Error GoTo Error:

tabMonsters.Index = "pkMonsters"
tabMonsters.Seek "=", nNum
If tabMonsters.NoMatch = True Then
    'MsgBox "Record " & nNum & " not found."
    Exit Function
End If


If Not bNoPrompt Then
    Set oLI = lvMonsterCompare.FindItem(nNum, lvwText, , 0)
    If Not oLI Is Nothing Then
        nYesNo = MsgBox(tabMonsters.Fields("Name") _
            & " is already on the compare list, add it again?", _
            vbYesNoCancel + vbDefaultButton2 + vbQuestion)
        If nYesNo = vbNo Then
            GoTo quit:
        ElseIf nYesNo = vbCancel Then
            CompareAddMonster = vbCancel
            GoTo quit:
        End If
    End If
End If

Call AddMonster2LV(lvMonsterCompare)
If lvMonsterCompare.ListItems.Count = 1 Then
    Set lvMonsterCompare.SelectedItem = lvMonsterCompare.ListItems(1)
    lvMonsterCompare.ListItems(1).Selected = True
    Call lvMonsterCompare_ItemClick(lvMonsterCompare.ListItems(1))
End If

If mnuJumpToCompare.Checked Then
    Call cmdNav_Click(3) 'compare
    Call cmdCompareNav_Click(3) 'mon
End If

quit:
Set oLI = Nothing

Exit Function
Error:
Call HandleError("CompareAddMonster")
Set oLI = Nothing
End Function

Private Sub CompareAddShopItems()
Dim oLI As ListItem

On Error GoTo Error:

bPromptSave = True
For Each oLI In lvShopDetail.ListItems
    If oLI.Selected Then
        If Val(oLI.Text) = 0 Then GoTo skip:
        
        tabItems.Index = "pkItems"
        tabItems.Seek "=", Val(oLI.Text)
            
        If tabItems.NoMatch = True Then
            MsgBox "Record " & oLI.Text & " not found."
        Else
            Select Case tabItems.Fields("ItemType")
                Case 0:
                    If Not tabItems.Fields("Worn") = 0 Then
                        Call AddArmour2LV(lvArmourCompare)
                        If txtArmourCompareDetail.Text = "" Then Call lvArmourCompare_ItemClick(lvArmourCompare.ListItems(1))
                    End If
                Case 1:
                    Call AddWeapon2LV(lvWeaponCompare)
                    If txtWeaponCompareDetail.Text = "" Then Call lvWeaponCompare_ItemClick(lvWeaponCompare.ListItems(1))
            End Select
            
        End If
        
skip:
    End If
Next

Set oLI = Nothing

Exit Sub

Error:
Call HandleError
End Sub

Private Sub CopyChartoClip(Optional ByVal bStatsOnly As Boolean)
Dim str As String, nExp As Currency, sExp As String
On Error GoTo Error:

'Name: Syntax Blackvail                 Lives/CP:      9/0
'Race: Dwarf       Exp: 791580519       Perception:     73
'Class: Paladin    Level: 55            Stealth:         0
'Hits:   566/671   Armour Class:  82/27 Thievery:        0
'Mana: *  84/126   Spellcasting: 174    Traps:           0
'                                       Picklocks:       0
'Strength:  130    Agility: 116         Tracking:        0
'Intellect: 90     Health:  110         Martial Arts:   21
'Willpower: 82     Charm:   51          MagicRes:       94

'line 1
str = "Name: " & txtCharName.Text & String(33 - Len(txtCharName.Text), " ")
str = str & "Lives/CP:" & String(7 - Len(lblStatCalc.Tag), " ") & "9/" & lblStatCalc.Tag

'line 2
nExp = CalcExpNeededByRaceClass(Val(txtGlobalLevel(0).Text), _
    cmbGlobalClass(0).ItemData(cmbGlobalClass(0).ListIndex), _
    cmbGlobalRace(0).ItemData(cmbGlobalRace(0).ListIndex))
sExp = CStr(nExp)
str = str & vbCrLf & "Race: " & cmbGlobalRace(0).Text & String(12 - Len(cmbGlobalRace(0).Text), " ")
str = str & "Exp: " & sExp & String(16 - Len(sExp), " ") & "Perception:      0"

'line 3
str = str & vbCrLf & "Class: " & cmbGlobalClass(0).Text & String(11 - Len(cmbGlobalClass(0).Text), " ")
str = str & "Level: " & txtGlobalLevel(0).Text & String(14 - Len(txtGlobalLevel(0).Text), " ")
str = str & "Stealth:         0"

'line 4
str = str & vbCrLf & "Hits: " & lblCharMaxHP.Tag & String(12 - Len(lblCharMaxHP.Tag), " ")
str = str & "Armour Class: " & txtStat(2).Tag & "/" & txtStat(3).Tag & _
    String(7 - Len(txtStat(2).Tag & "/" & txtStat(3).Tag), " ")
str = str & "Thievery:        0"

'line 5
If fraChar(2).Enabled Then
    str = str & vbCrLf & "Mana: " & lblCharMaxMana.Tag & "-" & lblCharMaxMana.Tag & _
        String(12 - Len(lblCharMaxMana.Tag & "-" & lblCharMaxMana.Tag), " ")
    str = str & "Spellcasting: " & lblCharSC.Tag & String(7 - Len(lblCharSC.Tag), " ")
    str = str & "Traps:           0"
Else
    str = str & vbCrLf & "                                       Traps:           0"
End If

'line 6
If fraChar(3).Enabled Then
    str = str & vbCrLf & String(39, " ") & "Picklocks:" & _
        String(8 - Len(lblCharPicklocks.Tag), " ") & lblCharPicklocks.Tag
Else
    str = str & vbCrLf & String(39, " ") & "Picklocks:       0"
End If

'line 7
str = str & vbCrLf & "Strength:  " & txtCharStats(0).Text & String(7 - Len(txtCharStats(0).Text), " ")
str = str & "Agility: " & txtCharStats(3).Text & String(12 - Len(txtCharStats(3).Text), " ")
str = str & "Tracking:        0"

'line 8
str = str & vbCrLf & "Intellect: " & txtCharStats(1).Text & String(7 - Len(txtCharStats(1).Text), " ")
str = str & "Health:  " & txtCharStats(4).Text & String(12 - Len(txtCharStats(4).Text), " ")
str = str & "Martial Arts:    0"

'line 9
str = str & vbCrLf & "Willpower: " & txtCharStats(2).Text & String(7 - Len(txtCharStats(2).Text), " ")
str = str & "Charm:   " & txtCharStats(5).Text & String(12 - Len(txtCharStats(5).Text), " ")
str = str & "MagicRes:        0"

'str = str & vbCrLf & vbCrLf & lblCharRestRate.Caption
'If fraChar(2).Enabled Then str = str & vbCrLf & lblCharManaRate.Caption

Clipboard.clear
Clipboard.SetText str

If bStatsOnly Then Exit Sub

Call InvenCopytoClipboard(True, True)

Exit Sub
Error:
Call HandleError("CharCopytoClip")
End Sub

Private Sub CopyMonsterToClipboard(ByRef lvDetail) ', ByRef txtDetail)
On Error GoTo Error:
Dim oLI As ListItem, str As String, x As Integer
    
For Each oLI In lvDetail.ListItems
    str = str & oLI.Text
    If Len(oLI.Text) <= 16 Then str = str & String(17 - Len(oLI.Text), " ")
    str = str & oLI.SubItems(1) & vbCrLf
Next

'If Not txtDetail.Text = "" Then str = str & vbCrLf & "Abilities: " & txtDetail.Text

If Not str = "" Then
    Clipboard.clear
    Clipboard.SetText str
End If

Set oLI = Nothing

Exit Sub
Error:
HandleError
Set oLI = Nothing

End Sub

Private Sub CopyShopToClipboard()
On Error GoTo Error:
Dim oLI As ListItem, oLSI As ListSubItem, oCH As ColumnHeader, str As String, x As Integer

x = 0
For Each oCH In lvShops.ColumnHeaders
    
    str = str & oCH.Text & ": "
    
    Select Case x
        Case 0: 'number
            str = str & lvShops.SelectedItem.Text & ", "
        Case 1: 'name
            str = str & lvShops.SelectedItem.SubItems(x) & vbCrLf
        Case Else:
            str = str & lvShops.SelectedItem.SubItems(x)
    End Select
    
    x = x + 1
Next

str = str & " -- " & txtShopDetail.Text

str = str & vbCrLf & lvShopLoc.ColumnHeaders(1).Text & ": "
x = 1
For Each oLI In lvShopLoc.ListItems
    If x > 1 Then str = str & ", "
    str = str & oLI.Text
    x = x + 1
Next

str = str & vbCrLf & vbCrLf

x = 1
For Each oCH In lvShopDetail.ColumnHeaders
    If x = 2 Then 'name
        str = str & oCH.Text
        str = str & String(30 - Len(oCH.Text), " ")
    Else
        str = str & oCH.Text
        If Len(oCH.Text) < 7 Then str = str & String(7 - Len(oCH.Text), " ")
    End If
    x = x + 1
Next

str = str & vbCrLf

For Each oLI In lvShopDetail.ListItems
    str = str & oLI.Text
    If Len(oLI.Text) <= 6 Then str = str & String(7 - Len(oLI.Text), " ")
    
    x = 1
    For Each oLSI In oLI.ListSubItems
    
        If x = 1 Then
            str = str & Left(oLSI, 29)
            str = str & String(30 - Len(oLSI), " ")
        Else
            str = str & oLSI
            If Len(oLSI) < 7 Then str = str & String(7 - Len(oLSI), " ")
        End If
        
        x = x + 1
    Next
    str = str & vbCrLf
Next

If Not str = "" Then
    Clipboard.clear
    Clipboard.SetText str
End If

Set oLI = Nothing
Set oLSI = Nothing
Set oCH = Nothing

Exit Sub
Error:
HandleError
Set oLI = Nothing
Set oLSI = Nothing
Set oCH = Nothing
End Sub

Private Sub FilterAll(ByVal EnableGlobalFilter As Boolean)
On Error GoTo Error:
'filterall(false) = remove all filters

Me.MousePointer = vbHourglass
DoEvents

'Call LockWindowUpdate(Me.hwnd)
Me.Enabled = False

Call ResetFilterOptions

If EnableGlobalFilter Then
    chkGlobalFilter.Value = 1
    Call cmbGlobalClass_Click(0)
    Call FilterWeapons(True)
    Call FilterArmour(True)
    Call FilterSpells(True)
    Call FilterInvenItems(False)
Else
    Call FilterWeapons(False)
    Call FilterArmour(False)
    Call FilterSpells(False)
    Call FilterInvenItems(True)
End If

quit:
Me.Enabled = True
Me.MousePointer = vbDefault
'Call LockWindowUpdate(0&)

Exit Sub

Error:
Call HandleError("FilterAll")
Resume quit:
End Sub

Private Sub FilterArmour(ByVal UseGlobalFilter As Boolean)
On Error GoTo Error:
Dim oLI As ListItem, x As Integer, nClass As Integer, bMagical As Boolean
Dim bClassOK As Boolean, nAlign As Integer, nNotAlign As Integer, bFiltered As Boolean

If tabItems.RecordCount = 0 Then Exit Sub

'Call InvenSetupEquip(True, False)
lvArmour.ListItems.clear
DoEvents

LockWindowUpdate Me.hwnd

DoEvents
tabItems.MoveFirst
Do Until tabItems.EOF

    nClass = 0
    bMagical = False
    bClassOK = False
    
    If tabItems.Fields("ItemType") = 0 Then 'armour
        
        If bOnlyInGame And tabItems.Fields("In Game") = 0 Then GoTo MoveNext:
        
        If chkArmourNoLimit.Value = 1 And tabItems.Fields("Limit") <> 0 Then GoTo skip:
        
        If cmbArmourWorn.ListIndex > 0 Then
            Select Case cmbArmourWorn.ItemData(cmbArmourWorn.ListIndex)
                Case 1: If Not tabItems.Fields("Worn") = 0 Then GoTo skip:  'NoWhere
                Case 2: If Not tabItems.Fields("Worn") = 2 Then GoTo skip:  '"Head"
                Case 3: If Not tabItems.Fields("Worn") = 3 Then GoTo skip:  '"Hands"
                Case 4: If Not tabItems.Fields("Worn") = 4 And Not tabItems.Fields("Worn") = 13 Then GoTo skip:  '"Finger"
                Case 5: If Not tabItems.Fields("Worn") = 5 Then GoTo skip:  '"Feet"
                Case 6: If Not tabItems.Fields("Worn") = 6 Then GoTo skip:  '"Arms"
                Case 7: If Not tabItems.Fields("Worn") = 7 Then GoTo skip:  '"Back"
                Case 8: If Not tabItems.Fields("Worn") = 8 Then GoTo skip:  '"Neck"
                Case 9: If Not tabItems.Fields("Worn") = 9 Then GoTo skip:  '"Legs"
                Case 10: If Not tabItems.Fields("Worn") = 10 Then GoTo skip:  '"Waist"
                Case 11: If Not tabItems.Fields("Worn") = 11 Then GoTo skip: '"Torso"
                Case 12: If Not tabItems.Fields("Worn") = 12 Then GoTo skip: '"Off-Hand"
                Case 13: If Not tabItems.Fields("Worn") = 14 And Not tabItems.Fields("Worn") = 17 Then GoTo skip: '"Wrist"
                Case 14: If Not tabItems.Fields("Worn") = 15 Then GoTo skip: '"Ears"
                Case 15: If Not tabItems.Fields("Worn") = 16 Then GoTo skip: '"Worn"
                Case 16: If Not tabItems.Fields("Worn") = 18 Then GoTo skip: '"Eyes"
                Case 17: If Not tabItems.Fields("Worn") = 19 Then GoTo skip: '"Face"
            End Select
        End If
        
        'If UseGlobalFilter Then
        '    If TestGlobalFilter = False Then GoTo skip:
        'Else 'no global filter
            For x = 0 To 6
                If chkArmourType(x).Value = 0 Then
                    Select Case x
                        Case 0: If tabItems.Fields("ArmourType") = 0 Then GoTo skip:
                        Case 1: If tabItems.Fields("ArmourType") = 1 Then GoTo skip:
                        Case 2: If tabItems.Fields("ArmourType") = 2 Then GoTo skip:
                        Case 3: If tabItems.Fields("ArmourType") >= 3 And tabItems.Fields("ArmourType") <= 6 Then GoTo skip:
                        Case 4: If tabItems.Fields("ArmourType") = 7 Then GoTo skip:
                        Case 5: If tabItems.Fields("ArmourType") = 8 Then GoTo skip:
                        Case 6: If tabItems.Fields("ArmourType") = 9 Then GoTo skip:
                    End Select
                End If
            Next
            
            For x = 0 To 19
                Select Case tabItems.Fields("Abil-" & x)
                    Case 28: bMagical = True    'magical check
                End Select
            Next
            
            If UseGlobalFilter Then
                If TestGlobalFilter(tabItems.Fields("Number")) = False Then GoTo skip:
            End If
        'End If
        
        If chkArmourNonMagic.Value = 1 And bMagical = True Then GoTo skip:
        
        Call AddArmour2LV(lvArmour)
    End If
    
GoTo MoveNext:
skip:
bFiltered = True
MoveNext:
    tabItems.MoveNext
    'DoEvents
Loop

For Each oLI In lvArmour.ListItems
    oLI.Selected = False
Next

bKeepSortOrder = True
Call lvArmour_ColumnClick(lvArmour.ColumnHeaders(nLastArmourSort))

If lvArmour.ListItems.Count >= 1 Then Call lvArmour_ItemClick(lvArmour.ListItems(1))
lvArmour.Refresh

If bFiltered Then cmdNav(1).Caption = "*Armour*" Else cmdNav(1).Caption = "Armour"

DoEvents
out:
On Error Resume Next
LockWindowUpdate 0&
Set oLI = Nothing

Exit Sub
Error:
Call HandleError("FilterArmour")
Resume out:

End Sub
'*******************************************************************************
' menu commands
'-------------------------------------------------------------------------------

Private Sub FilterInvenItems(Optional ByVal bRemoveFilter As Boolean)
On Error GoTo Error:
Dim x As Integer, y As Integer
Dim bFiltered As Boolean, nStartItems() As Long

ReDim nStartItems(0 To UBound(nEquippedItem()))

bDontRefresh = True

For x = 0 To UBound(nEquippedItem())
    If nEquippedItem(x) > 0 Then
        nStartItems(x) = nEquippedItem(x)
    End If
Next x

If bRemoveFilter Then
    Call LoadInvenItems
    
    For x = 0 To UBound(nEquippedItem())
        If cmbEquip(x).ListCount > 0 Then
            If nStartItems(x) > 0 Then
                For y = 0 To cmbEquip(x).ListCount - 1
                    If cmbEquip(x).ItemData(y) = nStartItems(x) Then
                        cmbEquip(x).ListIndex = y
                        Exit For
                    End If
                Next y
            End If
        End If
    Next x
    bDontRefresh = False
    Call RefreshAll
    GoTo out:
End If

If tabClasses.RecordCount = 0 Then GoTo out:

Call InvenSetupEquip
chkGlobalFilter.Value = 1
DoEvents

dofilter:
If tabItems.RecordCount = 0 Then GoTo out:

tabItems.MoveFirst
DoEvents

Do Until tabItems.EOF
    If bOnlyInGame And tabItems.Fields("In Game") = 0 Then GoTo MoveNext:
    If tabItems.Fields("ItemType") > 1 Then GoTo skip: 'not armour or weapon
    If TestGlobalFilter(tabItems.Fields("Number")) = False Then GoTo skip:
    Call InvenAddEquip(tabItems.Fields("Number"), tabItems.Fields("Name"), tabItems.Fields("ItemType"), tabItems.Fields("Worn"))
GoTo MoveNext:
skip:
bFiltered = True
MoveNext:
    tabItems.MoveNext
Loop

For x = 0 To cmbEquip().UBound
    If cmbEquip(x).ListCount > 0 Then
        If nStartItems(x) > 0 Then
            For y = 0 To cmbEquip(x).ListCount - 1
                If cmbEquip(x).ItemData(y) = nStartItems(x) Then
                    cmbEquip(x).ListIndex = y
                    Exit For
                End If
            Next y
        End If
    End If
Next x

If bFiltered Then cmdNav(4).Caption = "*Equipment*" Else cmdNav(4).Caption = "Equipment"

bDontRefresh = False
Call RefreshAll
out:
On Error Resume Next
Erase nStartItems()
Exit Sub
Error:
Call HandleError
bDontRefresh = False
Resume out:
End Sub

Private Sub FilterSpells(ByVal UseGlobalFilter As Boolean)
On Error GoTo Error:
Dim oLI As ListItem, x As Integer, nAlign As Integer, nNotAlign As Integer, nTarget As Integer
Dim bFiltered As Boolean, bHasAbility As Boolean

If tabSpells.RecordCount = 0 Then Exit Sub

lvSpells.ListItems.clear
DoEvents

'LockWindowUpdate lvSpells.hwnd
LockWindowUpdate Me.hwnd

DoEvents
'97, good
'98, evil
'112, neutral
'110, not good
'111, not evil
'113, not neutral

'0 "User"
'1 "Self"
'2 "Self or User"
'3 "Divided Area (not self)"
'4 "Monster"
'5 "Divided Area (incl self)"
'6 "Any"
'7 "Item"
'8 "Monster or User"
'9 "Divided Attack Area"
'10 "Divided Party Area"
'11 "Full Area"
'12 "Full Attack Area"
'13 "Full Party Area"

tabSpells.MoveFirst
Do Until tabSpells.EOF
    bHasAbility = False
    nAlign = 0
    nNotAlign = 0

    If Not cmbSpellMagery.ListIndex = 0 Then
        If Not cmbSpellMagery.ListIndex = tabSpells.Fields("Magery") Then GoTo skip:
        If Not cmbSpellMageryLevel.ListIndex = 0 Then
            If cmbSpellMageryLevel.ListIndex < tabSpells.Fields("MageryLVL") Then GoTo skip:
        End If
    End If

    'magery 5 is kai
    If chkSpellLearnable.Value = 1 And Not cmbSpellMagery.ListIndex = 5 And tabSpells.Fields("Learnable") = 0 Then GoTo skip:
    
    If Not cmbSpellAttackType.ListIndex = 0 Then
        If Not tabSpells.Fields("AttType") = cmbSpellAttackType.ListIndex - 1 Then GoTo skip: '-1 cause of "any"
    End If
    
    If Not cmbSpellTarget.ListIndex = 0 Then
        nTarget = tabSpells.Fields("Targets")
        Select Case cmbSpellTarget.ListIndex
            Case 1: 'self
                Select Case nTarget
                    Case 1, 2:
                    Case Else: GoTo skip:
                End Select
            
            Case 2: 'user
                Select Case nTarget
                    Case 0, 2, 8:
                    Case Else: GoTo skip:
                End Select
            
            Case 3: 'monster
                Select Case nTarget
                    Case 4, 6, 8:
                    Case Else: GoTo skip:
                End Select
            
            Case 4: 'party
                Select Case nTarget
                    Case 5, 10, 13:
                    Case Else: GoTo skip:
                End Select
            
            Case 5: 'room
                Select Case nTarget
                    Case 11, 12, 9, 3:
                    Case Else: GoTo skip:
                End Select
            
        End Select
    End If
            
    If UseGlobalFilter Then
        If Val(txtGlobalLevel(0).Text) < tabSpells.Fields("ReqLevel") Then GoTo skip:
        
        For x = 0 To 9
            Select Case tabSpells.Fields("Abil-" & x)
                Case 0:
                    
                Case 97, 98, 112: 'good/evil/neutral abils
                    nAlign = tabSpells.Fields("Abil-" & x)
                    Select Case cmbGlobalAlignment.ListIndex
                        Case 0:
                        Case 1: 'good
                            If Not nAlign = 97 Then GoTo skip:
                        Case 2: 'netural
                            If Not nAlign = 112 Then GoTo skip:
                        Case 3: 'evil
                            If Not nAlign = 98 Then GoTo skip:
                    End Select
            
                Case 110, 111, 113: 'notgood/notevil/notneutral abils
                    nNotAlign = tabSpells.Fields("Abil-" & x)
                    Select Case cmbGlobalAlignment.ListIndex
                        Case 0:
                        Case 1: 'good
                            If nNotAlign = 110 Then GoTo skip:
                        Case 2: 'netural
                            If nNotAlign = 113 Then GoTo skip:
                        Case 3: 'evil
                            If nNotAlign = 111 Then GoTo skip:
                    End Select

            End Select
            If cmbContainsAbil(0).ItemData(cmbContainsAbil(0).ListIndex) > 0 Then
                If tabSpells.Fields("Abil-" & x) = cmbContainsAbil(0).ItemData(cmbContainsAbil(0).ListIndex) Then
                    bHasAbility = True
                End If
            End If
        Next x
    Else
        For x = 0 To 9
            If cmbContainsAbil(0).ItemData(cmbContainsAbil(0).ListIndex) > 0 Then
                If tabSpells.Fields("Abil-" & x) = cmbContainsAbil(0).ItemData(cmbContainsAbil(0).ListIndex) Then
                    bHasAbility = True
                End If
            End If
        Next x
    End If
    
    If cmbContainsAbil(0).ItemData(cmbContainsAbil(0).ListIndex) > 0 Then
        If Not bHasAbility Then GoTo skip:
    End If
    
    Call AddSpell2LV(lvSpells)

GoTo MoveNext:
skip:
bFiltered = True
MoveNext:
    tabSpells.MoveNext
    'DoEvents
Loop

For Each oLI In lvSpells.ListItems
    oLI.Selected = False
Next

bKeepSortOrder = True
Call lvSpells_ColumnClick(lvSpells.ColumnHeaders(nLastSpellSort))

If lvSpells.ListItems.Count >= 1 Then Call lvSpells_ItemClick(lvSpells.ListItems(1))

If bFiltered Then cmdNav(2).Caption = "*Spells*" Else cmdNav(2).Caption = "Spells"

lvSpells.Refresh
DoEvents
out:
On Error Resume Next
LockWindowUpdate 0&
Set oLI = Nothing

Exit Sub
Error:
Call HandleError("FilterSpells")
Resume out:

End Sub

Private Sub FilterWeapons(ByVal UseGlobalFilter As Boolean)
On Error GoTo Error:
Dim oLI As ListItem, x As Integer, nClass As Integer
Dim bBSAble As Boolean, nMagical As Integer, nHitMagic As Integer, bFiltered As Boolean
Dim bClassOK As Boolean, bStaff As Boolean, nAlign As Integer, nNotAlign As Integer

If tabItems.RecordCount = 0 Then Exit Sub

'Call InvenSetupEquip(False, True)
lvWeapons.ListItems.clear
DoEvents

'LockWindowUpdate lvWeapons.hwnd
LockWindowUpdate Me.hwnd

DoEvents
tabItems.MoveFirst
Do Until tabItems.EOF

    nClass = 0
    bBSAble = False
    nMagical = 0
    nHitMagic = 0
    bStaff = False
    bClassOK = False
    
    If tabItems.Fields("ItemType") = 1 Then 'weapon
        
        If bOnlyInGame And tabItems.Fields("In Game") = 0 Then GoTo MoveNext:
        
        If chkWeaponNoLimit.Value = 1 And tabItems.Fields("Limit") <> 0 Then GoTo skip:
        
        Select Case cmbWeaponSpeed.ListIndex
            Case 0: '<=
                If Val(txtWeaponSpeed.Text) < tabItems.Fields("Speed") Then GoTo skip:
            Case 1: '>=
                If Val(txtWeaponSpeed.Text) > tabItems.Fields("Speed") Then GoTo skip:
        End Select
        
        'If UseGlobalFilter Then
        '    If TestGlobalFilter(tabItems.Fields("Number")) = False Then GoTo skip:
        '    bStaff = True
        'Else
            For x = 0 To 3
                If chkHanded(x).Value = 0 And tabItems.Fields("WeaponType") = x Then GoTo skip:
            Next
        'End If
        
        For x = 0 To 19
            Select Case tabItems.Fields("Abil-" & x)
                Case 116: 'BSable check
                    bBSAble = True
                    
                Case 28: 'magical check
                    nMagical = tabItems.Fields("AbilVal-" & x)
                
                Case 142: 'hitmagic check
                    nHitMagic = tabItems.Fields("AbilVal-" & x)
                
                Case 59: 'class ok
                    bStaff = True
            End Select
        Next x
        
        If UseGlobalFilter Then
            If TestGlobalFilter(tabItems.Fields("Number")) = False Then GoTo skip:
            bStaff = True
        End If
        
        If chkWeaponStaffOnly.Value = 1 And bStaff = False Then GoTo skip:
        
        If cmbWeaponMagicLevel.ListIndex > nMagical And _
            cmbWeaponMagicLevel.ListIndex > nHitMagic Then GoTo skip:
        
        If chkWeaponNonMagical.Value = 1 And nMagical > 0 Then GoTo skip:
        
        If chkBSAble.Value = 1 And bBSAble = False Then GoTo skip:
        
        Call AddWeapon2LV(lvWeapons)
    End If
    
GoTo MoveNext:
skip:
bFiltered = True
MoveNext:
    tabItems.MoveNext
    'DoEvents
Loop

For Each oLI In lvWeapons.ListItems
    oLI.Selected = False
Next

bKeepSortOrder = True
Call lvWeapons_ColumnClick(lvWeapons.ColumnHeaders(nLastWeaponSort))

If lvWeapons.ListItems.Count >= 1 Then Call lvWeapons_ItemClick(lvWeapons.ListItems(1))
lvWeapons.Refresh

If bFiltered = True Then cmdNav(0).Caption = "*Weapons*" Else cmdNav(0).Caption = "Weapons"

DoEvents
out:
On Error Resume Next
LockWindowUpdate 0&
Set oLI = Nothing

Exit Sub
Error:
Call HandleError("FilterWeapons")
Resume out:

End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If Shift Then Exit Sub
If Shift And vbAltMask Then Exit Sub
If Shift And vbCtrlMask Then Exit Sub

Select Case KeyCode
    Case 112: Call cmdNav_Click(0) 'f1
    Case 113: Call cmdNav_Click(1) 'f2
    Case 114: Call cmdNav_Click(2) 'f3
    Case 115: Call cmdNav_Click(3) 'f4
    Case 116: Call cmdNav_Click(4) 'f5
    Case 117: Call cmdNav_Click(5) 'f6
    Case 118: Call cmdNav_Click(6) 'f7
    Case 119: Call cmdNav_Click(7) 'f8
    Case 120: Call cmdNav_Click(8) 'f9
    Case 121: Call cmdNav_Click(9) 'f10
    Case 122: Call cmdNav_Click(10) 'f11
    Case 123: 'f12
End Select

If KeyCode >= 112 And KeyCode <= 123 Then KeyCode = 0
End Sub


Private Sub Form_Resize()
On Error Resume Next
Dim x As Integer, nWidth As Long, nHeight As Long, nTwipsEnlarged As Long

If Me.WindowState = vbMinimized Then Exit Sub

nWindowState = Me.WindowState

nWidth = Me.Width - 210
nHeight = Me.Height - 1950 - TITLEBAR_OFFSET

If Not bDontSpanNav Then
    nTwipsEnlarged = Fix((Me.Width - 10695) / 11)
    'Debug.Print nTwipsEnlarged
    framButtons.Width = nWidth
    fraDatVer.Width = nWidth - 6115
    lblDatVer.Width = fraDatVer.Width - 140
End If

For x = 0 To 10 'WHEN CHANGING THE SPAN STUFF, CHANGE IN frmSettings.cmdSave TOO
    framNav(x).Height = nHeight
    framNav(x).Width = nWidth
    If Not bDontSpanNav Then
        Select Case x
            Case 0:
                cmdNav(x).Width = 1095 + nTwipsEnlarged
            Case 1:
                cmdNav(x).Left = cmdNav(x - 1).Left + cmdNav(x - 1).Width - 15
                cmdNav(x).Width = 855 + nTwipsEnlarged
            Case 2:
                cmdNav(x).Left = cmdNav(x - 1).Left + cmdNav(x - 1).Width - 15
                cmdNav(x).Width = 795 + nTwipsEnlarged
            Case 3:
                cmdNav(x).Left = cmdNav(x - 1).Left + cmdNav(x - 1).Width - 15
                cmdNav(x).Width = 975 + nTwipsEnlarged
            Case 4:
                cmdNav(x).Left = cmdNav(x - 1).Left + cmdNav(x - 1).Width - 15
                cmdNav(x).Width = 1215 + nTwipsEnlarged
            Case 5:
                cmdNav(x).Left = cmdNav(x - 1).Left + cmdNav(x - 1).Width - 15
                cmdNav(x).Width = 1035 + nTwipsEnlarged
            Case 6:
                cmdNav(x).Left = cmdNav(x - 1).Left + cmdNav(x - 1).Width - 15
                cmdNav(x).Width = 1215 + nTwipsEnlarged
            Case 7:
                cmdNav(x).Left = cmdNav(x - 1).Left + cmdNav(x - 1).Width - 15
                cmdNav(x).Width = 795 + nTwipsEnlarged
            Case 8:
                cmdNav(x).Left = cmdNav(x - 1).Left + cmdNav(x - 1).Width - 15
                cmdNav(x).Width = 975 + nTwipsEnlarged
            Case 9:
                cmdNav(x).Left = cmdNav(x - 1).Left + cmdNav(x - 1).Width - 15
                cmdNav(x).Width = 735 + nTwipsEnlarged
            Case 10:
                cmdNav(x).Left = cmdNav(x - 1).Left + cmdNav(x - 1).Width - 15
                cmdNav(x).Width = 795 + nTwipsEnlarged
        End Select
    End If
Next x

For x = 0 To 3
    framCompareNav(x).Height = nHeight - 720
    framCompareNav(x).Width = nWidth - 240
Next

'wep/armr/spl
For x = 0 To 2
    splSplitterNS(x).Width = nWidth - 240
    splSplitterNS(x).Height = nHeight - 1000
Next x

'compares
For x = 3 To 5
    splSplitterNS(x).Width = nWidth - 480
    splSplitterNS(x).Height = nHeight - 1110
Next x
    
'sundry
splSplitterWE(6).Width = nWidth - 240
splSplitterWE(6).Height = nHeight - 720

'monsters
splMonsterSplit(0).Width = nWidth - 240
splMonsterSplit(0).Height = nHeight - 650
splMonsterSplit(1).Width = nWidth - 480
splMonsterSplit(1).Height = nHeight - 1110


'txtWeaponDetail.Top = nHeight - 1095
'txtWeaponDetail.Width = nWidth - 4260
'lvWeaponLoc.Top = nHeight - 1095
'lvWeaponLoc.Left = nWidth - 4035
'lvWeapons.Height = nHeight - 1980
'lvWeapons.Width = nWidth - 240
'
'txtArmourDetail.Top = nHeight - 1095
'txtArmourDetail.Width = nWidth - 4260
'lvArmourLoc.Top = nHeight - 1095
'lvArmourLoc.Left = nWidth - 4035
'lvArmour.Height = nHeight - 1980
'lvArmour.Width = nWidth - 240
'
'txtSpellDetail.Top = nHeight - 1275
'txtSpellDetail.Width = nWidth - 3780
'lvSpellLoc.Top = nHeight - 1275
'lvSpellLoc.Left = nWidth - 3615
'lvSpells.Height = nHeight - 2160
'lvSpells.Width = nWidth - 240
'lvSpells.ColumnHeaders(8).Width = nWidth - 6455
'

'
'lvWeaponCompareLoc.Left = nWidth - 4155
'lvWeaponCompareLoc.Top = nHeight - 1815
'txtWeaponCompareDetail.Width = nWidth - 4380
'txtWeaponCompareDetail.Top = nHeight - 1815
'lvWeaponCompare.Height = nHeight - 2160
'lvWeaponCompare.Width = nWidth - 480
'
'lvArmourCompareLoc.Left = nWidth - 4155
'lvArmourCompareLoc.Top = nHeight - 1815
'txtArmourCompareDetail.Width = nWidth - 4380
'txtArmourCompareDetail.Top = nHeight - 1815
'lvArmourCompare.Height = nHeight - 2160
'lvArmourCompare.Width = nWidth - 480
'
'lvSpellCompareLoc.Left = nWidth - 3855
'lvSpellCompareLoc.Top = nHeight - 1995
'txtSpellCompareDetail.Width = nWidth - 4020
'txtSpellCompareDetail.Top = nHeight - 1995
'lvSpellCompare.Height = nHeight - 2280
'lvSpellCompare.Width = nWidth - 480
'lvSpellCompare.ColumnHeaders(8).Width = nWidth - 6555
'
'txtMonsterDetail.Top = nHeight - 1035
'txtMonsterDetail.Width = nWidth - 5880
'lvMonsterDetail.Height = nHeight - 1380
'lvMonsterDetail.Width = nWidth - 5880
'lvMonsters.Height = nHeight - 720

'txtMonsterCompareDetail.Top = nHeight - 1835
'txtMonsterCompareDetail.Width = nWidth - 6130
'lvMonsterCompareLoc.Height = nHeight - 2120
'lvMonsterCompareLoc.Width = nWidth - 6130
'lvMonsterCompare.Height = nHeight - 1060

txtShopDetail.Top = nHeight - 1635
lvShopLoc.Top = nHeight - 1175
lvShops.Height = nHeight - 2220
lvShopDetail.Height = nHeight - 360 - 300
lvShopDetail.Width = nWidth - 4605
If lvShopDetail.ColumnHeaders.Count = 5 Then
    lvShopDetail.ColumnHeaders(5).Width = nWidth - 6255
ElseIf lvShopDetail.ColumnHeaders.Count = 2 Then
    lvShopDetail.ColumnHeaders(2).Width = nWidth - 6255
End If

txtRaceDetail.Top = nHeight - 795 + 300
txtRaceDetail.Width = nWidth - 240
lvRaces.Height = nHeight - 1120
lvRaces.Width = nWidth - 240
lvRaces.ColumnHeaders(11).Width = nWidth - 8700

txtClassDetail.Top = nHeight - 795 + 300
txtClassDetail.Width = nWidth - 240
lvClasses.Height = nHeight - 1120
lvClasses.Width = nWidth - 240
lvClasses.ColumnHeaders(9).Width = nWidth - 7555

'txtOtherItemDetail.Height = nHeight - 3160
'txtOtherItemDetail.Width = nWidth - 6740
'lvOtherItemLoc.Top = nHeight - 2515
'lvOtherItemLoc.Width = nWidth - 6740
'lvOtherItems.Height = nHeight - 700

txtMapMove.Height = nHeight - 5340
End Sub

Public Sub Form_Unload(Cancel As Integer)
Dim retval As Long
On Error Resume Next

If bDontSaveSettings Then GoTo term:

Cancel = SaveSettings
If Not Cancel = 0 Then
    modMain.bCancelTerminate = True
    Exit Sub
End If

term:
timMouseDown.Enabled = False
Set objWorkingListView = Nothing

'Set objToolTip = Nothing
Set objToolTip = Nothing
Set oLastColumnSorted = Nothing

If Not bNO_RECENT_FILES Then
    retval = SetWindowLong(Me.hwnd, GWL_WNDPROC, oldWindowProc) 'restore this window's original procedure before it unloads
End If

If Not bDontCallTerminate Then Call AppTerminate

End Sub

Public Sub GotoItem(nNum As Long)
On Error GoTo Error:
Dim oLI As ListItem, oLV As ListView, nYesNo As Integer
If nNum <= 0 Then Exit Sub

tabItems.Index = "pkItems"
tabItems.Seek "=", nNum
If tabItems.NoMatch = True Then
    Exit Sub
End If

Select Case tabItems.Fields("ItemType")
    Case 0:
        If tabItems.Fields("Worn") = 0 Then
            Set oLV = lvOtherItems
        Else
            Set oLV = lvArmour
        End If
    Case 1:
        Set oLV = lvWeapons
    Case Else:
        Set oLV = lvOtherItems
End Select
GoTo find:

tryagain:
tabItems.Index = "pkItems"
tabItems.Seek "=", nNum

find:
Set oLI = oLV.FindItem(nNum, lvwText, , 0)

If Not oLI Is Nothing Then
    Select Case tabItems.Fields("ItemType")
        Case 0: 'armour/other
            If tabItems.Fields("Worn") = 0 Then
                Call ClearListViewSelections(lvOtherItems)
                'Set lvOtherItems.SelectedItem = oLI
                Call lvOtherItems_ItemClick(oLI)
                Call cmdNav_Click(7)
            Else
                Call ClearListViewSelections(lvArmour)
                Call lvArmour_ItemClick(oLI)
                Call cmdNav_Click(1)
            End If
        Case 1: 'weapon
            Call ClearListViewSelections(lvWeapons)
            Call lvWeapons_ItemClick(oLI)
            Call cmdNav_Click(0)
        Case Else: 'other
            Call ClearListViewSelections(lvOtherItems)
            'Set lvOtherItems.SelectedItem = oLI
            Call lvOtherItems_ItemClick(oLI)
            Call cmdNav_Click(7)
    End Select
Else
    Select Case tabItems.Fields("ItemType")
        Case 0:
            If tabItems.Fields("Worn") = 0 Then
                MsgBox "Item " & nNum & " was not found in the current sundry list."
            Else
                nYesNo = MsgBox("Item " & nNum & " was not found in the current armour list." & vbCrLf _
                    & "Remove filter and try again?", vbYesNo + vbDefaultButton1 + vbQuestion)
                If nYesNo = vbYes Then Call cmdFilter_Click(5): GoTo tryagain:
            End If
        Case 1:
            nYesNo = MsgBox("Item " & nNum & " was not found in the current weapon list." & vbCrLf _
                & "Remove filter and try again?", vbYesNo + vbDefaultButton1 + vbQuestion)
            If nYesNo = vbYes Then Call cmdFilter_Click(7): GoTo tryagain:
        Case Else:
            MsgBox "Item " & nNum & " was not found in the current sundry list."
    End Select
End If

Set oLI = Nothing
Set oLV = Nothing

Exit Sub
Error:
Call HandleError
Set oLI = Nothing
Set oLV = Nothing

End Sub

Public Sub GotoLocation(oliSource As ListItem, Optional ByVal AuxNumber As Long, _
    Optional ByRef objFormOwner As Form)
On Error GoTo Error:
Dim oLI As ListItem, oLV As ListView, x As Integer, sStr As String, nNum As Long
Dim RoomExits As RoomExitType

If oliSource.Text = "... plus more." Then
    MsgBox "There were more references for this record but because of file size issues it was truncated.", vbInformation
    GoTo out:
End If

nNum = 16 'number of sStr's listed below +1
For x = 1 To nNum
    Select Case x
        Case 1: sStr = "Room"
        Case 2: sStr = "Monster"
        Case 3: sStr = "Textblock"
        Case 4: sStr = "Item"
        Case 5: sStr = "Spell"
        Case 6: sStr = "Shop"
        Case 7: sStr = "Group"
        Case 8: sStr = "Commands"
        Case 9: sStr = "NPC"
        Case 10: sStr = "Lair"
        Case 11: sStr = "Teleport"
        Case 12: sStr = "Action"
        Case 13: sStr = "Execute"
        Case 14: sStr = "Casts"
        Case 15: sStr = "Summon"
    End Select
    
    If Left(oliSource.Text, Len(sStr)) = sStr Then Exit For
    If x = nNum Then Exit Sub
Next x

If objFormOwner Is Nothing Then Set objFormOwner = Me

Select Case x
    Case 9, 10, 15: 'change to monster
        x = 2
    Case 12, 7: 'change to room
        x = 1
    Case 14: 'spell
        x = 5
End Select

If x = 1 Or x = 11 Then 'room/group
    RoomExits = ExtractMapRoom(oliSource.Tag)
Else
    nNum = Val(oliSource.Tag)
    If nNum <= 0 Then Exit Sub
End If

Me.MousePointer = vbHourglass
DoEvents

Select Case x
    Case 1, 11: 'room/group
        Set oLV = Nothing
    Case 2: 'monster
        Set oLV = lvMonsters
    Case 3: 'textblock
        'Load frmResults
        Call frmResults.SetupResultsWindow(True, objFormOwner, _
            IIf(AuxNumber > 0, AuxNumber, nMapStartMap))
        Call frmResults.CreateExecutionTree(nNum)
        frmResults.Show vbModeless, IIf(bNoAlwaysOnTop And Not objFormOwner Is frmMap, Nothing, objFormOwner)
        'this has to be here 'cause for some damn reason the activate event keeps firing off on frmMap when the cmdNav_Click goes (i think)
        If objFormOwner Is frmMap Then
            If frmMap.chkMapOptions(6).Value = 0 Then
                Call SetTopMostWindow(frmMap.hwnd, True)
            End If
        End If
        GoTo out:
    Case 4: 'item
        'this has to be here 'cause for some damn reason the activate event keeps firing off on frmMap when the cmdNav_Click goes (i think)
        If objFormOwner Is frmMap Then
            If frmMap.chkMapOptions(6).Value = 0 Then
                If frmMap.chkMapOptions(8).Value = 1 Then
                    Call SetTopMostWindow(frmMap.hwnd, False)
                    Me.SetFocus
                End If
            End If
        End If
        Call GotoItem(nNum)
        GoTo out:

    Case 5: 'spell
        'this has to be here 'cause for some damn reason the activate event keeps firing off on frmMap when the cmdNav_Click goes (i think)
        If objFormOwner Is frmMap Then
            If frmMap.chkMapOptions(6).Value = 0 Then
                If frmMap.chkMapOptions(8).Value = 1 Then
                    Call SetTopMostWindow(frmMap.hwnd, False)
                    Me.SetFocus
                End If
            End If
        End If
        Call GotoSpell(nNum)
        GoTo out:

    Case 6: 'shop
        Set oLV = lvShops
    Case 8: 'command
        'Load frmResults
        Call frmResults.SetupResultsWindow(True, objFormOwner, _
            IIf(AuxNumber > 0, AuxNumber, nMapStartMap))
        Call frmResults.CreateCommandTree(nNum, True, False)
        frmResults.Show vbModeless, IIf(bNoAlwaysOnTop And Not objFormOwner Is frmMap, Nothing, objFormOwner)
        'this has to be here 'cause for some damn reason the activate event keeps firing off on frmMap when the cmdNav_Click goes (i think)
        If objFormOwner Is frmMap Then
            If frmMap.chkMapOptions(6).Value = 0 Then
                Call SetTopMostWindow(frmMap.hwnd, True)
            End If
        End If
        GoTo out:

    Case 13: 'execute
        'Load frmResults
        Call frmResults.SetupResultsWindow(True, objFormOwner, _
            IIf(AuxNumber > 0, AuxNumber, nMapStartMap))
        Call frmResults.CreateCommandTree(nNum, False, False)
        frmResults.Show vbModeless, IIf(bNoAlwaysOnTop And Not objFormOwner Is frmMap, Nothing, objFormOwner)
        'this has to be here 'cause for some damn reason the activate event keeps firing off on frmMap when the cmdNav_Click goes (i think)
        If objFormOwner Is frmMap Then
            If frmMap.chkMapOptions(6).Value = 0 Then
                Call SetTopMostWindow(frmMap.hwnd, True)
            End If
        End If
        GoTo out:
End Select

If x = 1 Or x = 11 Then 'rooms/group
    If objFormOwner Is Me Then
        If x = 11 Then 'teleport
            If Len(txtMapMove.Text) > 0 Then
                sStr = GetTextblockCMDText("teleport " & RoomExits.Room & " " & RoomExits.Map, , GetRoomCMDTB(, nMapStartMap, nMapStartRoom))
                If Not sStr = "" Then
                    txtMapMove.Text = IIf(txtMapMove.Text = "", "", txtMapMove.Text & vbCrLf) & "@party " & sStr '"00000000:0080:"
                    txtMapMove.SelStart = Len(txtMapMove.Text)
                    txtMapMove.SelLength = 0
                End If
            End If
        End If
        Call cmdNav_Click(10)
    End If
    Call objFormOwner.MapStartMapping(RoomExits.Map, RoomExits.Room)
    GoTo check_exmap:
Else
    Set oLI = oLV.FindItem(nNum, lvwText, , 0)

    If Not oLI Is Nothing Then
        Select Case x
            Case 1: 'room
            Case 2: 'monster
                'this has to be here 'cause for some damn reason the activate event keeps firing off on frmMap when the cmdNav_Click goes (i think)
                If objFormOwner Is frmMap Then
                    If frmMap.chkMapOptions(6).Value = 0 Then
                        If frmMap.chkMapOptions(8).Value = 1 Then
                            Call SetTopMostWindow(frmMap.hwnd, False)
                            Me.SetFocus
                        End If
                    End If
                End If
                Call ClearListViewSelections(lvMonsters)
                Call lvMonsters_ItemClick(oLI)
                Call cmdNav_Click(8) 'monster
                
            Case 3: 'textblock
                GoTo out:
            Case 4: 'item
                'this has to be here 'cause for some damn reason the activate event keeps firing off on frmMap when the cmdNav_Click goes (i think)
                If objFormOwner Is frmMap Then
                    If frmMap.chkMapOptions(6).Value = 0 Then
                        If frmMap.chkMapOptions(8).Value = 1 Then
                            Call SetTopMostWindow(frmMap.hwnd, False)
                            Me.SetFocus
                        End If
                    End If
                End If
                Select Case tabItems.Fields("ItemType")
                    Case 0:
                        If tabItems.Fields("Worn") = 0 Then
                            Call ClearListViewSelections(lvOtherItems)
                            Call lvOtherItems_ItemClick(oLI)
                            Call cmdNav_Click(7) 'sundry
                        Else
                            Call ClearListViewSelections(lvArmour)
                            Call lvArmour_ItemClick(oLI)
                            Call cmdNav_Click(1) ' armour
                        End If
                    Case 1:
                        Call ClearListViewSelections(lvWeapons)
                        Call lvWeapons_ItemClick(oLI)
                        Call cmdNav_Click(0) ' weapons
                    Case Else:
                        Call ClearListViewSelections(lvOtherItems)
                        Call lvOtherItems_ItemClick(oLI)
                        Call cmdNav_Click(7) 'sundry
                End Select
        
            Case 5: 'spell
                'this has to be here 'cause for some damn reason the activate event keeps firing off on frmMap when the cmdNav_Click goes (i think)
                If objFormOwner Is frmMap Then
                    If frmMap.chkMapOptions(6).Value = 0 Then
                        If frmMap.chkMapOptions(8).Value = 1 Then
                            Call SetTopMostWindow(frmMap.hwnd, False)
                            Me.SetFocus
                        End If
                    End If
                End If
                Call ClearListViewSelections(lvSpells)
                Call lvSpells_ItemClick(oLI)
                Call cmdNav_Click(2) 'spell
                
            Case 6: 'shop
                'this has to be here 'cause for some damn reason the activate event keeps firing off on frmMap when the cmdNav_Click goes (i think)
                If objFormOwner Is frmMap Then
                    If frmMap.chkMapOptions(6).Value = 0 Then
                        If frmMap.chkMapOptions(8).Value = 1 Then
                            Call SetTopMostWindow(frmMap.hwnd, False)
                            Me.SetFocus
                        End If
                    End If
                End If
                Call ClearListViewSelections(lvShops)
                Call lvShops_ItemClick(oLI)
                Call cmdNav_Click(9) 'shop
                Call ClearListViewSelections(lvShopDetail)
                
                If AuxNumber > 0 Then
                    Set oLI = Nothing
                    Set oLI = lvShopDetail.FindItem(AuxNumber, lvwText, , 0)
                    If Not oLI Is Nothing Then
                        Call lvShopDetail_ItemClick(oLI)
                    End If
                End If
        End Select
        GoTo check_exmap:
    Else
        MsgBox sStr & " " & nNum & " not found in current " & sStr & " list."
    End If
End If
GoTo out:

check_exmap:
'If objFormOwner Is frmMap Then
'    If frmMap.chkMapOptions(6).Value = 0 Then
'        If frmMap.chkMapOptions(8).Value = 1 Then
'            Call SetTopMostWindow(frmMap.hwnd, False)
'            Me.SetFocus
'        End If
'    End If
'End If

out:
On Error Resume Next
Me.MousePointer = vbDefault
Set oLI = Nothing
Set oLV = Nothing
Exit Sub
Error:
Call HandleError("GotoLocation")
Resume out:
End Sub

Public Sub GotoMonster(nNum As Long)
On Error GoTo Error:
Dim oLI As ListItem, nYesNo As Integer
If nNum <= 0 Then Exit Sub

tabMonsters.Index = "pkMonsters"
tabMonsters.Seek "=", nNum
If tabMonsters.NoMatch = True Then
    Exit Sub
End If
GoTo find:

tryagain:
tabMonsters.Index = "pkMonsters"
tabMonsters.Seek "=", nNum

find:
Set oLI = lvMonsters.FindItem(nNum, lvwText, , 0)

If Not oLI Is Nothing Then
    Call ClearListViewSelections(lvMonsters)
    Call lvMonsters_ItemClick(oLI)
    Call cmdNav_Click(8)
Else
    MsgBox "Monster " & nNum & " was not found in the current Monster list.", vbInformation
End If

Set oLI = Nothing
Exit Sub
Error:
Call HandleError
Set oLI = Nothing
End Sub

Public Sub GotoSpell(nNum As Long)
On Error GoTo Error:
Dim oLI As ListItem, nYesNo As Integer
If nNum <= 0 Then Exit Sub

tabSpells.Index = "pkSpells"
tabSpells.Seek "=", nNum
If tabSpells.NoMatch = True Then
    Exit Sub
End If
GoTo find:

tryagain:
tabSpells.Index = "pkSpells"
tabSpells.Seek "=", nNum

find:
Set oLI = lvSpells.FindItem(nNum, lvwText, , 0)

If Not oLI Is Nothing Then
    Call ClearListViewSelections(lvSpells)
    Call lvSpells_ItemClick(oLI)
    Call cmdNav_Click(2)
Else
    nYesNo = MsgBox("Spell " & nNum & " was not found in the current spell list." & vbCrLf _
        & "Remove filter and try again?", vbYesNo + vbDefaultButton1 + vbQuestion)
    If nYesNo = vbYes Then Call cmdFilter_Click(3): GoTo tryagain:
End If

Set oLI = Nothing
Exit Sub
Error:
Call HandleError
Set oLI = Nothing
End Sub

Public Sub InvenAddEquip(ByVal nNum As Long, ByVal sName As String, ByVal ItemType As Integer, ByVal WornOn As Integer)
On Error GoTo Error:
Dim bResult As Boolean, x As Integer

Select Case ItemType
    Case 0: 'armour
        Select Case WornOn
            Case 0: '"Nowhere"
                x = -1
            Case 1: '"Everywhere"
                x = -1
            Case 2: '"Head"
                x = 0
            Case 3: '"Hands"
                x = 8
            Case 4: '"Finger"
                x = 9
            Case 5: '"Feet"
                x = 13
            Case 6: '"Arms"
                x = 5
            Case 7: '"Back"
                x = 3
            Case 8: '"Neck"
                x = 2
            Case 9: '"Legs"
                x = 12
            Case 10: '"Waist"
                x = 11
            Case 11: '"Torso"
                x = 4
            Case 12: '"Off-Hand"
                x = 15
            Case 13: '"Finger"
                x = 10
            Case 14: '"Wrist"
                x = 6
            Case 15: '"Ears"
                x = 1
            Case 16: '"Worn"
                x = 14
            Case 18: '"Eyes"
                x = 17
            Case 19: '"Face"
                x = 18
            Case Else:
                x = -1
        End Select
    Case 1: 'weapon
        x = 16
    Case Else:
        x = -1
End Select

Select Case x
    Case 9, 10: 'fingers
        cmbEquip(9).AddItem sName & " (" & nNum & ")"
        cmbEquip(9).ItemData(cmbEquip(9).NewIndex) = nNum
        cmbEquip(10).AddItem sName & " (" & nNum & ")"
        cmbEquip(10).ItemData(cmbEquip(10).NewIndex) = nNum
    Case 6, 7: 'wrists
        cmbEquip(6).AddItem sName & " (" & nNum & ")"
        cmbEquip(6).ItemData(cmbEquip(6).NewIndex) = nNum
        cmbEquip(7).AddItem sName & " (" & nNum & ")"
        cmbEquip(7).ItemData(cmbEquip(7).NewIndex) = nNum
    Case Is >= 0:
        cmbEquip(x).AddItem sName & " (" & nNum & ")"
        cmbEquip(x).ItemData(cmbEquip(x).NewIndex) = nNum
End Select

Exit Sub

Error:
Call HandleError
End Sub

Private Sub InvenCalcEncum()
On Error GoTo Error:

txtStat(1).Text = CalcEncum(Val(txtCharStats(0).Text), Val(txtStat(4).Text))

Exit Sub
Error:
Call HandleError
End Sub

Private Sub InvenCalcStats()
Dim x As Integer, y As Integer, sToolTip As String, nAC As Single, nDR As Single
Dim sName As String, Equip As TypeGetEquip, nRaceBonus As Long, nStrengthBonus As Long
Dim StatTips(0 To 29) As String

On Error GoTo Error:

Call InvenResetStats

For y = 0 To UBound(nEquippedItem())
    If nEquippedItem(y) > 0 Then
        If cmbEquip(y).ListIndex < 0 Then
            If cmbEquip(y).ListCount > 0 Then
                For x = 0 To cmbEquip(y).ListCount - 1
                    If cmbEquip(y).ItemData(x) = nEquippedItem(y) Then
                        cmbEquip(y).ListIndex = x
                        Exit Sub
                    End If
                Next x
                nEquippedItem(y) = 0
            Else
                nEquippedItem(y) = 0
            End If
        Else
            If Not cmbEquip(y).ItemData(cmbEquip(y).ListIndex) = nEquippedItem(y) Then
                For x = 0 To cmbEquip(y).ListCount - 1
                    If cmbEquip(y).ItemData(x) = nEquippedItem(y) Then
                        cmbEquip(y).ListIndex = x
                        Exit Sub
                    End If
                Next x
                nEquippedItem(y) = 0
            End If
        End If
    End If
Next y

For y = 0 To UBound(nEquippedItem())
    sToolTip = ""
    
    objToolTip.DelToolTip cmbEquip(y).hwnd
    
    For x = 0 To 29
        objToolTip.DelToolTip txtStat(x).hwnd
    Next x
    
    'If cmbEquip(y).ListCount = 0 Then GoTo skip:
    If nEquippedItem(y) < 1 Then GoTo skip:
    
    'If cmbEquip(y).ItemData(cmbEquip(y).ListIndex) = 0 Then GoTo skip:
    
    tabItems.Index = "pkItems"
    tabItems.Seek "=", nEquippedItem(y) 'cmbEquip(y).ItemData(cmbEquip(y).ListIndex)
    If tabItems.NoMatch Then GoTo skip:
    
    If y = 16 Then 'weapon
        If tabItems.Fields("WeaponType") = 1 Or tabItems.Fields("WeaponType") = 3 Then '2handed
            If nEquippedItem(15) > 0 Then 'shield
                chkEquipHold(15).ForeColor = &HFF&
                chkEquipHold(16).ForeColor = &HFF&
            Else
                chkEquipHold(15).ForeColor = &H0&
                chkEquipHold(16).ForeColor = &H0&
            End If
        Else
            chkEquipHold(15).ForeColor = &H0&
            chkEquipHold(16).ForeColor = &H0&
        End If
    End If
    
    sName = tabItems.Fields("Name")
    
    If Not tabItems.Fields("Encum") = 0 Then
        txtStat(0).Text = Val(txtStat(0).Text) + tabItems.Fields("Encum")
        If Not sToolTip = "" Then sToolTip = sToolTip & ", "
        sToolTip = sToolTip & "Enc: " & tabItems.Fields("Encum")
        
        If Not StatTips(0) = "" Then StatTips(0) = StatTips(0) & vbCrLf
        StatTips(0) = StatTips(0) & sName & " (" & tabItems.Fields("Encum") & ")"
    End If
    
    nAC = tabItems.Fields("ArmourClass") / 10 'RoundUp(tabItems.Fields("ArmourClass") / 10)
    nDR = tabItems.Fields("DamageResist") / 10
    If Not nAC = 0 Or Not nDR = 0 Then
        txtStat(2).Text = Round(Val(txtStat(2).Text) + nAC, 1)
        txtStat(3).Text = Round(Val(txtStat(3).Text) + nDR, 1)
        If Not sToolTip = "" Then sToolTip = sToolTip & ", "
        sToolTip = sToolTip & "AC: " & nAC & "/" & nDR
        
        If nAC > 0 Then
            If Not StatTips(2) = "" Then StatTips(2) = StatTips(2) & vbCrLf
            StatTips(2) = StatTips(2) & sName & " (" & nAC & "/" & nDR & ")"
        End If
        If nDR > 0 Then
            If Not StatTips(3) = "" Then StatTips(3) = StatTips(3) & vbCrLf
            StatTips(3) = StatTips(3) & sName & " (" & nAC & "/" & nDR & ")"
        End If
    End If
    
    If Not tabItems.Fields("Accy") = 0 Then
        txtStat(10).Text = Val(txtStat(10).Text) + tabItems.Fields("Accy")
        If Not sToolTip = "" Then sToolTip = sToolTip & ", "
        sToolTip = sToolTip & "Accy: " & tabItems.Fields("Accy")
        
        If Not StatTips(10) = "" Then StatTips(10) = StatTips(10) & vbCrLf
        StatTips(10) = StatTips(10) & sName & " (" & tabItems.Fields("Accy") & ")"
    End If
    
    'quick stats
    For x = 0 To 19
        If tabItems.Fields("Abil-" & x) > 0 And tabItems.Fields("AbilVal-" & x) <> 0 Then
            Equip = InvenGetEquipInfo(tabItems.Fields("Abil-" & x), tabItems.Fields("AbilVal-" & x))
            
            If Not tabItems.Fields("Number") = nEquippedItem(y) Then tabItems.Seek "=", nEquippedItem(y)
            
            If Equip.nEquip > 0 Then
                If Not Equip.sText = "" Then
                    If Not sToolTip = "" Then sToolTip = sToolTip & ", "
                    sToolTip = sToolTip & Equip.sText
                End If
                If Equip.nEquip = 3 Then 'dr
                    txtStat(Equip.nEquip).Text = Round(Val(txtStat(Equip.nEquip).Text) + (tabItems.Fields("AbilVal-" & x) / 10))
                    If Not StatTips(Equip.nEquip) = "" Then StatTips(Equip.nEquip) = StatTips(Equip.nEquip) & vbCrLf
                    StatTips(Equip.nEquip) = StatTips(Equip.nEquip) & sName & " (" & (tabItems.Fields("AbilVal-" & x) / 10) & ")"
                Else
                    txtStat(Equip.nEquip).Text = Val(txtStat(Equip.nEquip).Text) + tabItems.Fields("AbilVal-" & x)
                    If Not StatTips(Equip.nEquip) = "" Then StatTips(Equip.nEquip) = StatTips(Equip.nEquip) & vbCrLf
                    StatTips(Equip.nEquip) = StatTips(Equip.nEquip) & sName & " (" & tabItems.Fields("AbilVal-" & x) & ")"
                End If
            ElseIf Not Equip.sText = "" Then
                If Not sToolTip = "" Then sToolTip = sToolTip & ", "
                sToolTip = sToolTip & Equip.sText
                'txtStat(Equip.nEquip).Text = Val(txtStat(Equip.nEquip).Text) + tabItems.Fields("AbilVal-" & x)
            End If
        End If
    Next x

    'sToolTip = sName & " - " & tabItems.Fields("Number") & vbCrLf & sToolTip
    objToolTip.SetToolTipObj cmbEquip(y).hwnd, sToolTip, False
skip:
Next y

If cmbGlobalClass(0).ListIndex > 0 And tabClasses.RecordCount > 0 And chkInvenHideCharStats.Value = 0 Then
    tabClasses.Index = "pkClasses"
    tabClasses.Seek "=", cmbGlobalClass(0).ItemData(cmbGlobalClass(0).ListIndex)
    If Not tabClasses.NoMatch Then
        For x = 0 To 9
            Equip = InvenGetEquipInfo(tabClasses.Fields("Abil-" & x), 0)
                    
            If Equip.nEquip > 0 Then
                If Equip.nEquip = 3 Then 'dr
                    txtStat(Equip.nEquip).Text = Round(Val(txtStat(Equip.nEquip).Text) + (tabClasses.Fields("AbilVal-" & x) / 10), 1)
                    If Not StatTips(Equip.nEquip) = "" Then StatTips(Equip.nEquip) = StatTips(Equip.nEquip) & vbCrLf
                    StatTips(Equip.nEquip) = StatTips(Equip.nEquip) & "Class: " & tabClasses.Fields("Name") & " (" & (tabClasses.Fields("AbilVal-" & x) / 10) & ")"
                Else
                    txtStat(Equip.nEquip).Text = Val(txtStat(Equip.nEquip).Text) + tabClasses.Fields("AbilVal-" & x)
                    If Not StatTips(Equip.nEquip) = "" Then StatTips(Equip.nEquip) = StatTips(Equip.nEquip) & vbCrLf
                    StatTips(Equip.nEquip) = StatTips(Equip.nEquip) & "Class: " & tabClasses.Fields("Name") & " (" & (tabClasses.Fields("AbilVal-" & x)) & ")"
                End If
            End If
        Next x
    End If
End If

If cmbGlobalRace(0).ListIndex > 0 And tabRaces.RecordCount > 0 And chkInvenHideCharStats.Value = 0 Then

    nRaceBonus = GetRaceHPBonus(cmbGlobalRace(0).ItemData(cmbGlobalRace(0).ListIndex))
    If Not nRaceBonus = 0 Then
        nRaceBonus = nRaceBonus * Val(txtGlobalLevel(0).Text)
        txtStat(5).Text = Val(txtStat(5).Text) + nRaceBonus
        StatTips(5) = IIf(StatTips(5) = "", "", StatTips(5) _
            & vbCrLf) & "Race: " & GetRaceName(cmbGlobalRace(0).ItemData(cmbGlobalRace(0).ListIndex)) _
            & " (" & nRaceBonus & ")"
    End If
    
    tabRaces.Index = "pkRaces"
    tabRaces.Seek "=", cmbGlobalRace(0).ItemData(cmbGlobalRace(0).ListIndex)
    If Not tabRaces.NoMatch Then
        For x = 0 To 9
            Equip = InvenGetEquipInfo(tabRaces.Fields("Abil-" & x), 0)
                    
            If Equip.nEquip > 0 Then
                If Equip.nEquip = 3 Then 'dr
                    txtStat(Equip.nEquip).Text = Round(Val(txtStat(Equip.nEquip).Text) + (tabRaces.Fields("AbilVal-" & x) / 10), 1)
                    If Not StatTips(Equip.nEquip) = "" Then StatTips(Equip.nEquip) = StatTips(Equip.nEquip) & vbCrLf
                    StatTips(Equip.nEquip) = StatTips(Equip.nEquip) & "Race: " & tabRaces.Fields("Name") & " (" & (tabRaces.Fields("AbilVal-" & x) / 10) & ")"
                Else
                    txtStat(Equip.nEquip).Text = Val(txtStat(Equip.nEquip).Text) + tabRaces.Fields("AbilVal-" & x)
                    If Not StatTips(Equip.nEquip) = "" Then StatTips(Equip.nEquip) = StatTips(Equip.nEquip) & vbCrLf
                    StatTips(Equip.nEquip) = StatTips(Equip.nEquip) & "Race: " & tabRaces.Fields("Name") & " (" & (tabRaces.Fields("AbilVal-" & x)) & ")"
                End If
            End If
        Next x
    End If
End If

If Val(txtCharStats(0).Text) > 0 And chkInvenHideCharStats.Value = 0 Then 'strength max damage bonus
'    STR   Damage     STR   Damage
'    20     -3        90     +4
'    30     -2       100     +5
'    40     -1       110     +6
'    50     +0       120     +7
'    60     +1       130     +8
'    70     +2       140     +9
'    80     +3       150    +10     etc...
    nStrengthBonus = Fix((Val(txtCharStats(0).Text) + 50) / 10) - 10
    If Not nStrengthBonus = 0 Then
        StatTips(11) = IIf(StatTips(11) = "", "", StatTips(11) & vbCrLf) & "Strength (" _
                & nStrengthBonus & ")"
            txtStat(11).Text = Val(txtStat(11).Text) + nStrengthBonus
    End If
End If

If chkInvenAddWeight.Value = 1 Then
    If Val(txtInvenAddWeight.Text) > 0 Then
        StatTips(0) = IIf(StatTips(0) = "", "", StatTips(0) & vbCrLf) & "Additional Items (" _
            & txtInvenAddWeight.Text & ")"
        txtStat(0).Text = Val(txtStat(0).Text) + Val(txtInvenAddWeight.Text)
    End If
End If

If chkInvenHideCharStats.Value = 0 Then
    For x = 0 To 5
        If chkCharQuests(x).Value = 1 Then
            Select Case x
                Case 0: 'ice ho
                    txtStat(2).Text = Val(txtStat(2).Text) + 1
                    StatTips(2) = IIf(StatTips(2) = "", "", StatTips(2) & vbCrLf) _
                        & "Quest: Ice Sorceress (1)"
                Case 1: 'high druid
                    txtStat(9).Text = Val(txtStat(9).Text) + 1
                    StatTips(9) = IIf(StatTips(9) = "", "", StatTips(9) & vbCrLf) _
                        & "Quest: High Druid (1)"
                Case 2: 'red dragon
                    txtStat(7).Text = Val(txtStat(7).Text) + 1
                    txtStat(9).Text = Val(txtStat(9).Text) + 2
                    StatTips(7) = IIf(StatTips(7) = "", "", StatTips(7) & vbCrLf) _
                        & "Quest: Adult Red Dragon (1)"
                    StatTips(9) = IIf(StatTips(9) = "", "", StatTips(9) & vbCrLf) _
                        & "Quest: Adult Red Dragon (2)"
                Case 3: 'bishop
                    txtStat(10).Text = Val(txtStat(10).Text) + 3
                    StatTips(10) = IIf(StatTips(10) = "", "", StatTips(10) & vbCrLf) _
                        & "Quest: Bishop (3)"
                Case 4: 'apparatus
                    txtStat(8).Text = Val(txtStat(8).Text) + 1
                    StatTips(8) = IIf(StatTips(8) = "", "", StatTips(8) & vbCrLf) _
                        & "Quest: Apparatus (1)"
                Case 5: '2nd quest
                    Select Case cmbChar2ndAlign.ListIndex
                        Case 0: 'nothing
                        Case 1: '"+1 max damage"
                            txtStat(11).Text = Val(txtStat(11).Text) + 1
                            StatTips(11) = IIf(StatTips(11) = "", "", StatTips(11) & vbCrLf) _
                                & "Quest: 2nd Align (1)"
                        Case 2: '"+1 AC, +6 mana"
                            txtStat(2).Text = Val(txtStat(2).Text) + 1
                            StatTips(2) = IIf(StatTips(2) = "", "", StatTips(2) & vbCrLf) _
                                & "Quest: 2nd Align (1)"
                            txtStat(6).Text = Val(txtStat(6).Text) + 6
                            StatTips(6) = IIf(StatTips(6) = "", "", StatTips(6) & vbCrLf) _
                                & "Quest: 2nd Align (6)"
                        Case 3: '"+1 SC, +10 mana"
                            txtStat(9).Text = Val(txtStat(9).Text) + 1
                            StatTips(9) = IIf(StatTips(9) = "", "", StatTips(9) & vbCrLf) _
                                & "Quest: 2nd Align (1)"
                            txtStat(6).Text = Val(txtStat(6).Text) + 10
                            StatTips(6) = IIf(StatTips(6) = "", "", StatTips(6) & vbCrLf) _
                                & "Quest: 2nd Align (10)"
                        Case 4: '"+4 mana, +6 bs min/max, +1 stealth"
                            txtStat(6).Text = Val(txtStat(6).Text) + 4
                            StatTips(6) = IIf(StatTips(6) = "", "", StatTips(6) & vbCrLf) _
                                & "Quest: 2nd Align (4)"
                            txtStat(14).Text = Val(txtStat(14).Text) + 6
                            StatTips(14) = IIf(StatTips(14) = "", "", StatTips(14) & vbCrLf) _
                                & "Quest: 2nd Align (6)"
                            txtStat(15).Text = Val(txtStat(15).Text) + 6
                            StatTips(15) = IIf(StatTips(15) = "", "", StatTips(15) & vbCrLf) _
                                & "Quest: 2nd Align (6)"
                            txtStat(19).Text = Val(txtStat(19).Text) + 1
                            StatTips(19) = IIf(StatTips(19) = "", "", StatTips(6) & vbCrLf) _
                                & "Quest: 2nd Align (1)"
                        Case 5: '"+10 bs min/max, +2 stealth"
                            txtStat(14).Text = Val(txtStat(14).Text) + 10
                            StatTips(14) = IIf(StatTips(14) = "", "", StatTips(14) & vbCrLf) _
                                & "Quest: 2nd Align (10)"
                            txtStat(15).Text = Val(txtStat(15).Text) + 10
                            StatTips(15) = IIf(StatTips(15) = "", "", StatTips(15) & vbCrLf) _
                                & "Quest: 2nd Align (10)"
                            txtStat(19).Text = Val(txtStat(19).Text) + 2
                            StatTips(19) = IIf(StatTips(19) = "", "", StatTips(6) & vbCrLf) _
                                & "Quest: 2nd Align (2)"
                    End Select
            End Select
        End If
    Next x
End If

'If Val(txtStat(16).Text) > 0 Then 'hp regen
    txtCharHPRegen.Text = Val(txtStat(16).Text)
'End If

'If Val(txtStat(17).Text) > 0 Then 'mana regen
    txtCharManaRegen.Text = Val(txtStat(17).Text)
'End If

Call InvenCalcEncum

For x = 0 To 29
    If Not StatTips(x) = "" Then objToolTip.SetToolTipObj txtStat(x).hwnd, StatTips(x), False
Next x

If nEquippedItem(15) < 1 And nEquippedItem(16) < 1 Then
    chkEquipHold(15).ForeColor = &H0&
    chkEquipHold(16).ForeColor = &H0&
End If

txtStat(2).Tag = Fix(Val(txtStat(2).Text))
txtStat(3).Tag = Fix(Val(txtStat(3).Text))

Call InvenColorCodeStats
'Call txtCharStats_Change(0)

Exit Sub

Error:
Call HandleError
End Sub

Private Sub InvenClear()
Dim x As Integer

bDontRefresh = True

For x = 0 To cmbEquip().UBound
    If chkEquipHold(x).Value = 0 Then cmbEquip(x).ListIndex = 0
Next x

If Not bStartup Then bDontRefresh = False
Call RefreshAll
End Sub

Private Sub InvenColorCodeStats()
Dim x As Integer

On Error GoTo Error:

For x = 4 To txtStat().Count - 1
    Select Case Val(txtStat(x).Text)
        Case Is < 0:
            txtStat(x).ForeColor = &HFF&
        Case Is > 0:
            txtStat(x).ForeColor = &HFFFFFF
            txtStat(x).Text = "+" & Val(txtStat(x).Text)
    End Select
Next x

Exit Sub

Error:
Call HandleError

End Sub

Private Sub InvenCopytoClipboard(Optional ByVal bAppend As Boolean, Optional ByVal bNoCharStats As Boolean)
Dim str As String, x As Integer, sItem As String, sStats As String ', sQuests As String

On Error GoTo Error:

If Not bNoCharStats Then
    If cmbGlobalClass(0).ListIndex > 0 Then str = str & "Class: " & cmbGlobalClass(0).Text & vbCrLf
    If cmbGlobalRace(0).ListIndex > 0 Then str = str & "Race: " & cmbGlobalRace(0).Text & vbCrLf
    If Val(txtGlobalLevel(0).Text) > 0 Then str = str & "Level: " & txtGlobalLevel(0).Text & vbCrLf
    If Val(txtCharStats(0).Text) > 0 Then str = str & "Strength: " & txtCharStats(0).Text & vbCrLf
End If

str = str & vbCrLf

If Not Val(txtStat(2).Text) = 0 Or Not Val(txtStat(3).Text) = 0 Then
    str = str & "Armour Class: " & txtStat(2).Text & "/" & txtStat(3).Text & vbCrLf
End If

str = str & "Encumberance: " & txtStat(0).Text & "/" & txtStat(1).Text _
    & " " & lblEncumLevel(1).Tag & vbCrLf & vbCrLf & "He is equipped with:" & vbCrLf & vbCrLf

For x = 0 To cmbEquip().UBound
    If cmbEquip(x).ListIndex > 0 Then
        sItem = Mid(cmbEquip(x).Text, 1, InStr(1, cmbEquip(x).Text, "(") - 2)
        str = str & sItem & String(31 - Len(sItem), " ") _
            & "(" & GetEquipCaption(x) & ")" & vbCrLf
    End If
Next x

For x = 5 To 29
    If Not Val(txtStat(x).Text) = 0 Then
        If Not sStats = "" Then sStats = sStats & ", "
        sStats = sStats & Left(lblInvenStats(x).Caption, InStr(1, lblInvenStats(x).Caption, ":") - 1) & " " & txtStat(x).Text
    End If
Next
If Not sStats = "" Then str = str & vbCrLf & "Stats: " & sStats

'For x = 0 To 4
'    If chkCharQuests(x).Value = 1 Then
'        If Not sQuests = "" Then sQuests = sQuests & ", "
'        sQuests = sQuests & chkCharQuests(x).Caption
'    End If
'Next x
'If Not sQuests = "" Then str = str & vbCrLf & vbCrLf & "Quests: " & sQuests

If Not str = "" Then
    If bAppend Then str = Clipboard.GetText & vbCrLf & vbCrLf & str
    Clipboard.clear
    Clipboard.SetText str
End If

Exit Sub

Error:
Call HandleError("InvenCopytoClipboard")

End Sub

Private Sub InvenEquipItem(ByVal nNum As Long, Optional ByVal bJumpToNav As Boolean)
On Error GoTo Error:
Dim oCMB As ComboBox, x As Integer, nTries As Integer

If nNum = 0 Then Exit Sub

tabItems.Index = "pkItems"
tabItems.Seek "=", nNum
If tabItems.NoMatch = True Then
    'MsgBox "Item " & nNum & " not found."
    Exit Sub
End If

Select Case tabItems.Fields("ItemType")
    Case 0: 'armour
        Select Case tabItems.Fields("Worn")
            Case 0: '"Nowhere"
            Case 1: '"Everywhere"
            Case 2: '"Head"
                Set oCMB = cmbEquip(0)
            Case 3: '"Hands"
                Set oCMB = cmbEquip(8)
            Case 4, 13: '"Finger"
                If nEquippedItem(9) > 0 Then
                    If nEquippedItem(9) = nNum Then GoTo done:
                    If nEquippedItem(10) = nNum Then GoTo done:
                    Set oCMB = cmbEquip(10)
                ElseIf nEquippedItem(10) > 0 Then
                    If nEquippedItem(10) = nNum Then GoTo done:
                    Set oCMB = cmbEquip(9)
                Else
                    Set oCMB = cmbEquip(9)
                End If
            Case 5: '"Feet"
                Set oCMB = cmbEquip(13)
            Case 6: '"Arms"
                Set oCMB = cmbEquip(5)
            Case 7: '"Back"
                Set oCMB = cmbEquip(3)
            Case 8: '"Neck"
                Set oCMB = cmbEquip(2)
            Case 9: '"Legs"
                Set oCMB = cmbEquip(12)
            Case 10: '"Waist"
                Set oCMB = cmbEquip(11)
            Case 11: '"Torso"
                Set oCMB = cmbEquip(4)
            Case 12: '"Off-Hand"
                Set oCMB = cmbEquip(15)
            Case 14: '"Wrist"
                If nEquippedItem(6) > 0 Then
                    If nEquippedItem(6) = nNum Then GoTo done:
                    If nEquippedItem(7) = nNum Then GoTo done:
                    If bInvenUse2ndWrist Then
                        Set oCMB = cmbEquip(7)
                    Else
                        Set oCMB = cmbEquip(6)
                    End If
                ElseIf nEquippedItem(7) > 0 Then
                    If nEquippedItem(7) = nNum Then GoTo done:
                    Set oCMB = cmbEquip(6)
                Else
                    Set oCMB = cmbEquip(6)
                End If
            Case 15: '"Ears"
                Set oCMB = cmbEquip(1)
            Case 16: '"Worn"
                Set oCMB = cmbEquip(14)
            Case 18: '"Eyes"
                Set oCMB = cmbEquip(17)
            Case 19: '"Face"
                Set oCMB = cmbEquip(18)
            Case Else:
        End Select
        
        
    Case 1: 'weapon
        Set oCMB = cmbEquip(16)
End Select

nTries = 0
If Not oCMB Is Nothing Then
again:
    For x = 0 To oCMB.ListCount - 1
        If oCMB.ItemData(x) = nNum Then
            oCMB.ListIndex = x
            GoTo done:
        End If
    Next
    
    If nTries > 0 Then GoTo done:
    
    Call InvenAddEquip(nNum, tabItems.Fields("Name"), tabItems.Fields("ItemType"), tabItems.Fields("Worn"))
    nTries = 1 'add the item then try to search once again
    GoTo again:
End If

done:
If bJumpToNav Then Call cmdNav_Click(4)

Set oCMB = Nothing
Exit Sub
Error:
Call HandleError("InvenEquipItem")
Set oCMB = Nothing
End Sub

Private Sub InvenFindBest(ByVal Index As Integer, ByVal nFindBestEnm As enmFindBest)
Dim nAbility As Integer, nAbility2 As Integer, nAbility3 As Integer, sField As String
Dim x As Integer, y As Integer, z As Integer, LastFindBest(18) As Long
Dim tWinner(18) As TypeEquipWinner, tPosWinner(18) As TypeEquipWinner
Dim bNo2Handed As Boolean, nStart As Integer, nResult As Boolean
On Error GoTo Error:

'.... very confusing.  tposwinner is the possible winner that it finds when searching
'.... twinner is the current winning item (best item)
'.... lastfindbest is the value of the last item selected

Me.Enabled = False
Me.MousePointer = vbHourglass
DoEvents

nInvenLastIndex(0) = Index
nInvenLastIndex(1) = nFindBestEnm

bDontRefresh = True

nAbility = 0
nAbility2 = 0
sField = ""

Select Case nFindBestEnm
    Case 0: 'Armour
        Select Case Index
            Case 0: 'ac/dr
                
            Case 1: 'ac
                sField = "ArmourClass"
            Case 2: 'dr
                sField = "DamageResist"
            Case 3: 'dodge
                nAbility = 34
            Case 4: 'prot evil
                nAbility = 24
            Case 5: 'prot good
                nAbility = 25
        End Select
    Case 1: 'Attack
        Select Case Index
            Case 0: 'acc
                nAbility = 22
                nAbility2 = 105
                nAbility3 = 106
                sField = "Accy"
                
            Case 1: 'bs acc
                nAbility = 116
            Case 2: 'bs min
                nAbility = 117
            Case 3: 'bs max
                nAbility = 118
            Case 4: 'crit
                nAbility = 58
            Case 5: 'dam sh
                nAbility = 72
            Case 6: 'max dam
                nAbility = 4
        End Select
    Case 2: 'Resist
        Select Case Index
            Case 0: 'mr
                nAbility = 36
            Case 1: 'cold
                nAbility = 3
            Case 2: 'fire
               nAbility = 5
            Case 3: 'light
                nAbility = 66
            Case 4: 'stone
                nAbility = 65
            Case 5: 'water
                nAbility = 147
        End Select
    Case 3: 'Stat
        Select Case Index
            Case 0: 'enc
                nAbility = 96
            Case 1: 'hp
                nAbility = 88
            Case 2: 'hp rgn
                nAbility = 123
            Case 3: 'ill
                nAbility = 13
                nAbility2 = 14
            Case 4: 'mana
                nAbility = 69
            Case 5: 'mana rgn
                nAbility = 145
            Case 6: 'pick
                nAbility = 37
                nAbility2 = 180
            Case 7: 'spell
                nAbility = 70
            Case 8: 'stealth
                nAbility = 27
            Case 9: 'thiev
                nAbility = 39
            Case 10: 'trap
                nAbility = 40
                nAbility2 = 41
                nAbility3 = 179
        End Select
    Case 4: 'Mystics
        Select Case Index
            Case 0: 'jump acc
                nAbility = 91
            Case 1: 'jump dmg
                nAbility = 94
            Case 2: 'kick acc
                nAbility = 90
            Case 3: 'kick dmg
                nAbility = 93
            Case 4: 'punch acc
                nAbility = 89
            Case 5: 'punch dmg
                nAbility = 92
        End Select
End Select

If bInvenNextBest Then
    For x = 0 To UBound(nEquippedItem())
        If nEquippedItem(x) < 1 Then GoTo next_nextbest:

        tabItems.Index = "pkItems"
        tabItems.Seek "=", nEquippedItem(x)
        If tabItems.NoMatch Then GoTo next_nextbest:
        
        For z = 0 To UBound(nInvenExcludedItems())
            If nInvenExcludedItems(z) = nEquippedItem(x) Then
                Exit For
            Else
                If z = UBound(nInvenExcludedItems()) Then
                    ReDim Preserve nInvenExcludedItems(0 To z + 1)
                    nInvenExcludedItems(z + 1) = nEquippedItem(x)
                End If
            End If
        Next z
        
        If nFindBestEnm = Armour And Index = 0 Then 'AC/DR
            LastFindBest(x) = tabItems.Fields("ArmourClass") + tabItems.Fields("DamageResist")
            'tWinner(x).Value = tabItems.Fields("ArmourClass") + tabItems.Fields("DamageResist")
            'tWinner(x).Number = tabItems.Fields("Number")
        Else
            For z = 0 To 19
                If nAbility > 0 Then 'first ability check
                    If tabItems.Fields("Abil-" & z) = nAbility Then
                        LastFindBest(x) = tabItems.Fields("AbilVal-" & z)
                        'tWinner(x).Value = tabItems.Fields("AbilVal-" & z)
                        'tWinner(x).Number = tabItems.Fields("Number")
                        GoTo next_nextbest:
                    End If
                End If

                If nAbility2 > 0 Then 'second ability check
                    If tabItems.Fields("Abil-" & z) = nAbility2 Then
                        LastFindBest(x) = tabItems.Fields("AbilVal-" & z)
                        'tWinner(x).Value = tabItems.Fields("AbilVal-" & z)
                        'tWinner(x).Number = cmbEquip(x).ItemData(y)
                        GoTo next_nextbest:
                    End If
                End If
                
                If nAbility3 > 0 Then 'third ability check
                    If tabItems.Fields("Abil-" & z) = nAbility3 Then
                        LastFindBest(x) = tabItems.Fields("AbilVal-" & z)
                        'tWinner(x).Value = tabItems.Fields("AbilVal-" & z)
                        'tWinner(x).Number = cmbEquip(x).ItemData(y)
                        GoTo next_nextbest:
                    End If
                End If
            Next z

            If Not sField = "" Then 'if we're checking on a non-ability
                LastFindBest(x) = tabItems.Fields(sField)
                'tWinner(x).Value = tabItems.Fields(sField)
                'tWinner(x).Number = cmbEquip(x).ItemData(y)
                GoTo next_nextbest:
            End If
        End If

next_nextbest:
    Next x
Else
    Erase nInvenExcludedItems()
    ReDim nInvenExcludedItems(0)
End If

nStart = 0
recheck:
'DoEvents
For x = nStart To cmbEquip().UBound 'x=cmbequip()
    If cmbEquip(x).ListCount = 0 Then GoTo skip:
    If chkEquipHold(x).Value = 1 Then GoTo skip:
    
    For y = 0 To cmbEquip(x).ListCount - 1 'y=listindex
        tabItems.Index = "pkItems"
        tabItems.Seek "=", cmbEquip(x).ItemData(y)
        If tabItems.NoMatch Then GoTo nextitem:
        
        If chkInvenNoLimited.Value = 1 And tabItems.Fields("Limit") > 0 Then GoTo nextitem:
        
        If x = 16 And bNo2Handed = True Then 'weapon
            If tabItems.Fields("WeaponType") = 1 Or tabItems.Fields("WeaponType") = 3 Then GoTo nextitem:
        End If
        
        If nFindBestEnm = Armour And Index = 0 Then 'AC/DR
            If tabItems.Fields("ArmourClass") + tabItems.Fields("DamageResist") > tPosWinner(x).Value Then
                nResult = InvenFindBestDupeFail(x, y, tPosWinner())
                If nResult = False Then GoTo nextitem:
                tPosWinner(x).Value = tabItems.Fields("ArmourClass") + tabItems.Fields("DamageResist")
                tPosWinner(x).Number = cmbEquip(x).ItemData(y)
            End If
        Else
            For z = 0 To 19
                If nAbility > 0 Then 'first ability check
                    If tabItems.Fields("Abil-" & z) = nAbility Then
                        If tabItems.Fields("AbilVal-" & z) > tPosWinner(x).Value Then
                            nResult = InvenFindBestDupeFail(x, y, tPosWinner())
                            If nResult = False Then GoTo nextitem:
                            tPosWinner(x).Value = tabItems.Fields("AbilVal-" & z)
                            tPosWinner(x).Number = cmbEquip(x).ItemData(y)
                        End If
                    End If
                End If
                
                If nAbility2 > 0 Then 'second ability check
                    If tabItems.Fields("Abil-" & z) = nAbility2 Then
                        If tabItems.Fields("AbilVal-" & z) > tPosWinner(x).Value Then
                            nResult = InvenFindBestDupeFail(x, y, tPosWinner())
                            If nResult = False Then GoTo nextitem:
                            tPosWinner(x).Value = tabItems.Fields("AbilVal-" & z)
                            tPosWinner(x).Number = cmbEquip(x).ItemData(y)
                        End If
                    End If
                End If
                
                If nAbility3 > 0 Then 'third ability check
                    If tabItems.Fields("Abil-" & z) = nAbility3 Then
                        If tabItems.Fields("AbilVal-" & z) > tPosWinner(x).Value Then
                            nResult = InvenFindBestDupeFail(x, y, tPosWinner())
                            If nResult = False Then GoTo nextitem:
                            tPosWinner(x).Value = tabItems.Fields("AbilVal-" & z)
                            tPosWinner(x).Number = cmbEquip(x).ItemData(y)
                        End If
                    End If
                End If
            Next z
            
            If Not sField = "" Then 'if we're checking on a non-ability
                If tabItems.Fields(sField) > tPosWinner(x).Value Then
                    nResult = InvenFindBestDupeFail(x, y, tPosWinner())
                    If nResult = False Then GoTo nextitem:
                    tPosWinner(x).Value = tabItems.Fields(sField)
                    tPosWinner(x).Number = cmbEquip(x).ItemData(y)
                End If
            End If
        End If
        
        If tPosWinner(x).Number > 0 Then 'if we found an item
            If bInvenNextBest Then
                If tPosWinner(x).Value <= LastFindBest(x) Then
                
                    For z = 0 To UBound(nInvenExcludedItems())
                        If nInvenExcludedItems(z) = cmbEquip(x).ItemData(y) Then
                            tPosWinner(x).Number = 0
                            tPosWinner(x).Value = 0
                        End If
                    Next z
                    
                    If tPosWinner(x).Value > tWinner(x).Value Then
                        tWinner(x).Number = tPosWinner(x).Number
                        tWinner(x).Value = tPosWinner(x).Value
                        'tPosWinner(x).Value = 0
                    Else
                        tPosWinner(x).Number = 0
                        tPosWinner(x).Value = 0
                    End If
                Else
                    tPosWinner(x).Number = 0
                    tPosWinner(x).Value = 0
                End If
            Else
                tWinner(x).Number = tPosWinner(x).Number
                tWinner(x).Value = tPosWinner(x).Value
                'tPosWinner(x).Number = 0
            End If
        End If
nextitem:
    Next y 'next item in list
    
    If tWinner(x).Number > 0 Then 'if we found a winner
           
        'tWinner(x).Number = tPosWinner(x).Number
        'tWinner(x).Value = tPosWinner(x).Value
        
        For y = 0 To cmbEquip(x).ListCount - 1
            If cmbEquip(x).ItemData(y) = tWinner(x).Number Then
                cmbEquip(x).ListIndex = y
                LastFindBest(x) = tWinner(x).Value
                Exit For
            End If
        Next y
    End If
skip:
Next x 'next piece of equipment

If nEquippedItem(15) > 0 And nEquippedItem(16) > 0 Then 'if weapon and off-hand got selected
    tabItems.Index = "pkItems"
    tabItems.Seek "=", tWinner(16).Number
    If Not tabItems.NoMatch Then
        If tabItems.Fields("WeaponType") = 1 Or tabItems.Fields("WeaponType") = 3 Then 'if the weapon is 2-handed

            If chkEquipHold(15).Value = 0 And chkEquipHold(16).Value = 0 Then 'neither held
                If tWinner(15).Value >= tWinner(16).Value Then
                    cmbEquip(16).ListIndex = 0
                    tWinner(16).Value = 0
                    tWinner(16).Number = 0
                    tPosWinner(16).Number = 0
                    tPosWinner(16).Value = 0
                    bNo2Handed = True
                    nStart = 16
                    GoTo recheck:
                Else
                    cmbEquip(15).ListIndex = 0
                End If

            ElseIf chkEquipHold(15).Value = 1 And chkEquipHold(16).Value = 0 Then 'off-hand held
                cmbEquip(16).ListIndex = 0
                tWinner(16).Value = 0
                tWinner(16).Number = 0
                tPosWinner(16).Number = 0
                tPosWinner(16).Value = 0
                bNo2Handed = True
                nStart = 16
                GoTo recheck:

            ElseIf chkEquipHold(15).Value = 0 And chkEquipHold(16).Value = 1 Then 'weapon held
                cmbEquip(15).ListIndex = 0

            End If
        End If
    End If
End If

For x = 0 To 18
    If Not tWinner(x).Number = 0 Then
        Exit For
    End If
    If x = 18 Then MsgBox "Nothing found.", vbInformation
Next x

out:
On Error Resume Next
Me.MousePointer = vbDefault
Me.Enabled = True
bDontRefresh = False
Call RefreshAll
Exit Sub
Error:
Call HandleError("InvenFindBest")
Resume out:
End Sub

Private Function InvenFindBestDupeFail(ByVal nCmbEquipNum As Integer, _
    ByVal nCmbEquipListIndex As Integer, ByRef tPosWinner() As TypeEquipWinner) As Boolean

On Error GoTo Error:

'before:
'Select Case nCmbEquipNum
'    Case 6, 7: 'wrists
'        If nCmbEquipNum = 7 And Not bInvenUse2ndWrist Then Exit Function
'        If tPosWinner(6).Number = cmbEquip(nCmbEquipNum).ItemData(nCmbEquipListIndex) Then Exit Function
'        If tPosWinner(7).Number = cmbEquip(nCmbEquipNum).ItemData(nCmbEquipListIndex) Then Exit Function
'        If cmbEquip(6).ItemData(cmbEquip(6).ListIndex) = cmbEquip(nCmbEquipNum).ItemData(nCmbEquipListIndex) Then Exit Function
'        If cmbEquip(7).ItemData(cmbEquip(7).ListIndex) = cmbEquip(nCmbEquipNum).ItemData(nCmbEquipListIndex) Then Exit Function
'    Case 9, 10: 'fingers
'        If tPosWinner(10).Number = cmbEquip(nCmbEquipNum).ItemData(nCmbEquipListIndex) Then Exit Function
'        If tPosWinner(9).Number = cmbEquip(nCmbEquipNum).ItemData(nCmbEquipListIndex) Then Exit Function
'        If cmbEquip(10).ItemData(cmbEquip(10).ListIndex) = cmbEquip(nCmbEquipNum).ItemData(nCmbEquipListIndex) Then Exit Function
'        If cmbEquip(9).ItemData(cmbEquip(9).ListIndex) = cmbEquip(nCmbEquipNum).ItemData(nCmbEquipListIndex) Then Exit Function
'End Select

'after:
Select Case nCmbEquipNum
    Case 6, 7: 'wrists
        If nCmbEquipNum = 7 And Not bInvenUse2ndWrist Then Exit Function
        If tPosWinner(6).Number = cmbEquip(nCmbEquipNum).ItemData(nCmbEquipListIndex) Then Exit Function
        If tPosWinner(7).Number = cmbEquip(nCmbEquipNum).ItemData(nCmbEquipListIndex) Then Exit Function
        If nEquippedItem(6) = cmbEquip(nCmbEquipNum).ItemData(nCmbEquipListIndex) Then Exit Function
        If nEquippedItem(7) = cmbEquip(nCmbEquipNum).ItemData(nCmbEquipListIndex) Then Exit Function
    Case 9, 10: 'fingers
        If tPosWinner(9).Number = cmbEquip(nCmbEquipNum).ItemData(nCmbEquipListIndex) Then Exit Function
        If tPosWinner(10).Number = cmbEquip(nCmbEquipNum).ItemData(nCmbEquipListIndex) Then Exit Function
        If nEquippedItem(9) = cmbEquip(nCmbEquipNum).ItemData(nCmbEquipListIndex) Then Exit Function
        If nEquippedItem(10) = cmbEquip(nCmbEquipNum).ItemData(nCmbEquipListIndex) Then Exit Function
End Select


proceed:
InvenFindBestDupeFail = True

Exit Function
Error:
Call HandleError("InvenFindBestDupeFail")

End Function

Private Function InvenGetEquipInfo(ByVal nAbility As Integer, ByVal nAbilityValue As Integer) As TypeGetEquip

If nAbilityValue > 0 Then InvenGetEquipInfo.sText = GetAbilityStats(nAbility, nAbilityValue)

InvenGetEquipInfo.nEquip = -1
Select Case nAbility
    Case 0: 'nothing
    Case 2: '2=AC
        InvenGetEquipInfo.nEquip = 2
        'InvenGetEquipInfo.sText = "AC: "
    Case 3: '3=res_cold
        InvenGetEquipInfo.nEquip = 28
        'InvenGetEquipInfo.sText = "Cold Res: "
    Case 4: '4=max dmg
        InvenGetEquipInfo.nEquip = 11
        'InvenGetEquipInfo.sText = "Max Dmg: "
    Case 5: '5=res_fire
        InvenGetEquipInfo.nEquip = 27
        'InvenGetEquipInfo.sText = "Fire Res: "
    Case 7: '7=DR
        InvenGetEquipInfo.nEquip = 3
        'InvenGetEquipInfo.sText = "DR: "
    Case 10: '10=AC
        InvenGetEquipInfo.nEquip = 2
        'InvenGetEquipInfo.sText = "AC: "
    Case 13: '13=illu
        InvenGetEquipInfo.nEquip = 23
        'InvenGetEquipInfo.sText = "Illu: "
    Case 14: '14=roomillu
        InvenGetEquipInfo.nEquip = 23
        'InvenGetEquipInfo.sText = "RoomIllu: "
    Case 22: '22=acc
        InvenGetEquipInfo.nEquip = 10
        'InvenGetEquipInfo.sText = "Accy: "
    Case 24: '42=prev
        InvenGetEquipInfo.nEquip = 20
    '25=prgd
    Case 27: '27=stealth
        InvenGetEquipInfo.nEquip = 19
        'InvenGetEquipInfo.sText = "Stealth: "
    Case 34: '34=dodge
        InvenGetEquipInfo.nEquip = 8
        'InvenGetEquipInfo.sText = "Dodge: "
    Case 36: '36=MR
        InvenGetEquipInfo.nEquip = 24
        'InvenGetEquipInfo.sText = "MR: "
    Case 37: '37=picklocks
        InvenGetEquipInfo.nEquip = 22
        'InvenGetEquipInfo.sText = "Picks: "
    Case 38: '38=tracking
        'InvenGetEquipInfo.nEquip = 23
        ''InvenGetEquipInfo.sText = "Tracking: "
    Case 39: '39=thievery
        'InvenGetEquipInfo.nEquip = 20
        'InvenGetEquipInfo.sText = "Thievery: "
    Case 40: '40=findtraps
        InvenGetEquipInfo.nEquip = 21
        'InvenGetEquipInfo.sText = "Traps: "
    '41=disarmtraps
    '44=int
    '45=wis
    '46=str
    '47=hea
    '48=agi
    '49=chm
    Case 58: '58=crits
        InvenGetEquipInfo.nEquip = 7
        'InvenGetEquipInfo.sText = "Crits: "
    Case 65: '65=res_stone
        InvenGetEquipInfo.nEquip = 25
        'InvenGetEquipInfo.sText = "Stone Res: "
    Case 66: '66=res_lit
        InvenGetEquipInfo.nEquip = 29
        'InvenGetEquipInfo.sText = "Light Res: "
    Case 69: '69=max mana
        InvenGetEquipInfo.nEquip = 6
        'InvenGetEquipInfo.sText = "Mana: "
    Case 70: '70=SC
        InvenGetEquipInfo.nEquip = 9
        'InvenGetEquipInfo.sText = "SC: "
    Case 72: '72=damageshield
        InvenGetEquipInfo.nEquip = 12
        'InvenGetEquipInfo.sText = "Shock: "
    Case 77: '77=percep
        InvenGetEquipInfo.nEquip = 18
        'InvenGetEquipInfo.sText = "Percep: "
    '87=speed
    Case 88: '88=alter hp
        InvenGetEquipInfo.nEquip = 5
        'InvenGetEquipInfo.sText = "HP: "
    '89=punchavg
    '89=kickavg
    '89=jumpavg
    '89=punchdmg
    '89=kickdmg
    '89=jumpdmg
    Case 96: '96=encum
        InvenGetEquipInfo.nEquip = 4
        'InvenGetEquipInfo.sText = "Enc%: "
    Case 105: '105=acc
        InvenGetEquipInfo.nEquip = 10
        'InvenGetEquipInfo.sText = "Accy: "
    Case 106: '106=acc
        InvenGetEquipInfo.nEquip = 10
        'InvenGetEquipInfo.sText = "Accy: "
    Case 116: '116=bsaccu
        InvenGetEquipInfo.nEquip = 13
        'InvenGetEquipInfo.sText = "BS Accy: "
    Case 117: '117=bsmin
        InvenGetEquipInfo.nEquip = 14
        'InvenGetEquipInfo.sText = "BS Min: "
    Case 118: '118=bsmax
        InvenGetEquipInfo.nEquip = 15
        'InvenGetEquipInfo.sText = "BS Max: "
    Case 123: '123=hpregen
        InvenGetEquipInfo.nEquip = 16
        'InvenGetEquipInfo.sText = "HP Rgn: "
    Case 142: '142=hitmagic
        'InvenGetEquipInfo.nEquip = 31
        ''InvenGetEquipInfo.sText = "Hit Magic: "
    Case 145: '145=manaregen
        InvenGetEquipInfo.nEquip = 17
        'InvenGetEquipInfo.sText = "Mana Rgn: "
    Case 147: '147=res_water
        InvenGetEquipInfo.nEquip = 26
        'InvenGetEquipInfo.sText = "Water Res: "
    Case 179: '179=find trap value
        InvenGetEquipInfo.nEquip = 21
        'InvenGetEquipInfo.sText = "Traps: "
    Case 180: '180=pick value
        InvenGetEquipInfo.nEquip = 22
        'InvenGetEquipInfo.sText = "Picks: "
    
End Select
End Function

Private Sub InvenResetStats()
Dim x As Integer

For x = 0 To txtStat().Count - 1
    txtStat(x).Text = "0"
    txtStat(x).Tag = ""
    txtStat(x).ForeColor = &HC0C000
Next x

End Sub

Private Sub InvenSetupEquip() 'Optional ArmourOnly As Boolean, Optional WeaponOnly As Boolean)
Dim x As Integer ', y As Integer, z As Integer

On Error GoTo Error:

'If ArmourOnly Then
'    y = 0
'    z = 15
'ElseIf WeaponOnly Then
'    y = 16
'    z = 16
'Else
'    y = 0
'    z = 16
'End If

For x = 0 To cmbEquip().UBound
    objToolTip.DelToolTip cmbEquip(x).hwnd
    cmbEquip(x).clear
    cmbEquip(x).AddItem "(none)", 0
    cmbEquip(x).ItemData(cmbEquip(x).NewIndex) = 0
    cmbEquip(x).ListIndex = 0
    nEquippedItem(x) = 0
    Call ExpandCombo(cmbEquip(x), HeightOnly, DoubleWidth, framNav(4).hwnd)
    cmbEquip(x).SelLength = 0
Next

Exit Sub

Error:
Call HandleError("InvenSetupEquip")
bDontRefresh = False
End Sub

Private Sub lblRoomCell_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Error:

nMapLastCellIndex = Index
lvMapLoc.ListItems.clear

If CellRoom(Index, 1) = 0 Then
    If Button = 2 And Shift = 1 Then
        Call MapStartMapping(nMapStartMap, nMapStartRoom, Index)
        Exit Sub
    Else
        Exit Sub
    End If
End If

If bMapSwapButtons Then
    If Button = 2 Then
        Button = 1
    ElseIf Button = 1 Then
        Button = 2
    End If
End If

If Button = 1 Then
    Call MapGetRoomLoc(CellRoom(Index, 1), CellRoom(Index, 2))
ElseIf Button = 2 Then
    If lblRoomCell(Index).BackColor = &HFF00& Then '-- up
        Call PopUpMapMenu(True, False)
    ElseIf lblRoomCell(Index).BackColor = &HFFFF& Then '-- down
        Call PopUpMapMenu(False, True)
    ElseIf lblRoomCell(Index).BackColor = &HFFFF00 Then '-- both
        Call PopUpMapMenu(True, True)
    Else
        If Shift = 1 Then
            Call MapStartMapping(nMapStartMap, nMapStartRoom, Index)
        Else
            Call MapStartMapping(CellRoom(Index, 1), CellRoom(Index, 2))
        End If
    End If
End If

Exit Sub
Error:
Call HandleError

End Sub

Private Sub LoadCharacter(ByVal bPromptForFile As Boolean, Optional ByVal strFileName As String, _
    Optional ByVal bClearSelections As Boolean)
On Error GoTo Error:
Dim sFile As String, nItem As Long, sCompares As String, x As Integer, y As Integer
Dim sSectionName As String, bJustLoad As Boolean, sFileTitle As String
Dim bLoadCompare As Boolean, bLoadInven As Boolean, sName As String

sSectionName = RemoveCharacter(lblDatVer.Caption, " ")

If bPromptForFile Then
    oComDag.Filter = "MMUD Explorer Character (*.mmec)|*.mmec"
    oComDag.DialogTitle = "Load Character"
    sFile = ReadINI(sSectionName, "LastLoadName")
    If Len(sFile) < 4 Then
        oComDag.FileName = "Character.mmec"
    Else
        oComDag.FileName = sFile
    End If
    
    sFile = ReadINI(sSectionName, "LastLoadDir")
    If FolderExists(sFile) Then
        oComDag.InitDir = sFile
    Else
        oComDag.InitDir = App.Path
    End If
    
    On Error GoTo canceled:
    oComDag.ShowOpen
    If oComDag.FileName = "" Then GoTo canceled:
    
    sFile = oComDag.FileName
    sFileTitle = oComDag.FileTitle
    
    Me.Enabled = False
    frmLoadChar.Tag = "-1"
    frmLoadChar.Show vbModal, Me
    Me.Enabled = True
    
    If frmLoadChar.Tag = "-1" Then GoTo canceled:
    
    If frmLoadChar.chkInvenLoad.Value = 1 Then bLoadInven = True
    If frmLoadChar.chkCompareLoad.Value = 1 Then bLoadCompare = True
Else
    bJustLoad = True
    sFile = strFileName
    If Not sFile = "" Then
        For x = 1 To Len(sFile)
            If InStr(x, sFile, "\") > 0 Then
                y = InStr(x, sFile, "\") + 1
                x = y
            End If
        Next x
        If y = 0 Then y = 1
        sFileTitle = Mid(sFile, y)
    End If
End If

On Error GoTo Error:

bCharLoaded = False
If Not sFile = "" Then
    If Not UCase(Right(sFile, 5)) = ".MMEC" Then sFile = sFile & ".mmec"
    
    If Not FileExists(sFile) Then
        MsgBox "Character File (" & sFile & ") was not found.", vbExclamation
        sFile = ""
        sSectionName = RemoveCharacter(lblDatVer.Caption, " ")
        GoTo loadfromINI:
    End If
    
    bCharLoaded = True
    
    Call RecentFileAdd(sFile)
    
    Call WriteINI(sSectionName, "LastLoadDir", Left(sFile, Len(sFile) - Len(sFileTitle) - 1))
    Call WriteINI(sSectionName, "LastLoadName", sFileTitle)
    Call WriteINI(sSectionName, "LastCharFile", sFile)
Else
    Call RecentFileAdd
End If

loadfromINI:
bDontRefresh = True

If Not sFile = "" Then sSectionName = "PlayerInfo"
'chkGlobalFilter.Value = ReadINI(sSectionName, "UseGlobalFilter", sFile)
txtGlobalLevel(0).Text = ReadINI(sSectionName, "Level", sFile)
cmbGlobalAlignment.ListIndex = ReadINI(sSectionName, "Alignment", sFile)

sName = ReadINI(sSectionName, "Name", sFile)
If sName = "0" Then sName = ""
If bNameInTitle And bCharLoaded Then
    If Len(sName) > 1 Then
        Me.Caption = sNormalCaption & " (" & sName & ")"
    Else
        Me.Caption = sNormalCaption & " (" & sFileTitle & ")"
    End If
Else
    If bCharLoaded Then
        Me.Caption = sNormalCaption & " (" & sFileTitle & ")"
    Else
        Me.Caption = sNormalCaption
    End If
End If
txtCharName.Text = sName

txtCharStats(0).Text = ReadINI(sSectionName, "Strength", sFile)
txtCharStats(1).Text = ReadINI(sSectionName, "Intellect", sFile)
txtCharStats(2).Text = ReadINI(sSectionName, "Widsom", sFile)
txtCharStats(3).Text = ReadINI(sSectionName, "Agility", sFile)
txtCharStats(4).Text = ReadINI(sSectionName, "Health", sFile)
txtCharStats(5).Text = ReadINI(sSectionName, "Charm", sFile)

txtInvenAddWeight.Text = ReadINI(sSectionName, "AddWeight", sFile)
chkInvenAddWeight.Value = ReadINI(sSectionName, "UseAddWeight", sFile)

For x = 0 To 4
    chkCharQuests(x).Value = ReadINI(sSectionName, "Quest" & x, sFile)
Next x
cmbChar2ndAlign.ListIndex = ReadINI(sSectionName, "Quest_2nd", sFile)

x = Val(ReadINI(sSectionName, "Class", sFile))
If Not cmbGlobalClass(0).ListCount = 0 Then
    For y = 0 To cmbGlobalClass(0).ListCount - 1
        If cmbGlobalClass(0).ItemData(y) = x Then
            cmbGlobalClass(0).ListIndex = y
            Exit For
        End If
    Next y
End If

x = Val(ReadINI(sSectionName, "Race", sFile))
If Not cmbGlobalRace(0).ListCount = 0 Then
    For y = 0 To cmbGlobalRace(0).ListCount - 1
        If cmbGlobalRace(0).ItemData(y) = x Then
            cmbGlobalRace(0).ListIndex = y
            Exit For
        End If
    Next y
End If

If Not bJustLoad Then
    If frmLoadChar.optFilter(0).Value = True Then 'filter all based on char
        Call FilterAll(True)
    ElseIf frmLoadChar.optFilter(1).Value = True Then 'leave as is
    ElseIf frmLoadChar.optFilter(2).Value = True Then 'reset filters
        Call FilterAll(False)
    End If
End If

If Not sFile = "" Then sSectionName = "Bless"
For x = 0 To 9
    nItem = ReadINI(sSectionName, "Bless" & x, sFile)
    If nItem > 0 Then
        For y = 0 To cmbCharBless(x).ListCount - 1
            If cmbCharBless(x).ItemData(y) = nItem Then
                cmbCharBless(x).ListIndex = y
                Exit For
            End If
        Next y
    Else
        cmbCharBless(x).ListIndex = 0
    End If
Next x

If bJustLoad Or bLoadInven Then
    If bJustLoad And bClearSelections = True Then
        Call InvenClear
    ElseIf bJustLoad = False And frmLoadChar.chkInvenClear.Value = 1 Then
        Call InvenClear
    End If
    
    If Not sFile = "" Then sSectionName = "Inventory"
    
    bDontRefresh = True
    Call InvenEquipItem(Val(ReadINI(sSectionName, "Head", sFile)), False)
    Call InvenEquipItem(Val(ReadINI(sSectionName, "Ears", sFile)), False)
    Call InvenEquipItem(Val(ReadINI(sSectionName, "Neck", sFile)), False)
    Call InvenEquipItem(Val(ReadINI(sSectionName, "Back", sFile)), False)
    Call InvenEquipItem(Val(ReadINI(sSectionName, "Torso", sFile)), False)
    Call InvenEquipItem(Val(ReadINI(sSectionName, "Arms", sFile)), False)
    Call InvenEquipItem(Val(ReadINI(sSectionName, "Wrist", sFile)), False)
    Call InvenEquipItem(Val(ReadINI(sSectionName, "Wrist2", sFile)), False)
    Call InvenEquipItem(Val(ReadINI(sSectionName, "Waist", sFile)), False)
    Call InvenEquipItem(Val(ReadINI(sSectionName, "Hands", sFile)), False)
    Call InvenEquipItem(Val(ReadINI(sSectionName, "Finger1", sFile)), False)
    Call InvenEquipItem(Val(ReadINI(sSectionName, "Finger2", sFile)), False)
    Call InvenEquipItem(Val(ReadINI(sSectionName, "Legs", sFile)), False)
    Call InvenEquipItem(Val(ReadINI(sSectionName, "Feet", sFile)), False)
    Call InvenEquipItem(Val(ReadINI(sSectionName, "Worn", sFile)), False)
    Call InvenEquipItem(Val(ReadINI(sSectionName, "Off-Hand", sFile)), False)
    Call InvenEquipItem(Val(ReadINI(sSectionName, "Weapon", sFile)), False)
    Call InvenEquipItem(Val(ReadINI(sSectionName, "Eyes", sFile)), False)
    Call InvenEquipItem(Val(ReadINI(sSectionName, "Face", sFile)), False)
    
End If

If bLoadCompare Or bJustLoad Then
    If bJustLoad And bClearSelections Then
        Call cmdClearAllCompares_Click
    ElseIf bJustLoad = False And frmLoadChar.chkCompareClear.Value = 1 Then
        Call cmdClearAllCompares_Click
    End If
    
    If Not sFile = "" Then sSectionName = "Compare"
    
    sCompares = ReadINI(sSectionName, "WeaponCompare", sFile)
    x = 0
    Do While Not InStr(x + 1, sCompares, ",") = 0
        y = InStr(x + 1, sCompares, ",")
        
        tabItems.Index = "pkItems"
        tabItems.Seek "=", Val(Mid(sCompares, x + 1, y - x - 1))
        If tabItems.NoMatch = False Then
            Call AddWeapon2LV(lvWeaponCompare)
        End If
        x = y
    Loop
    If lvWeaponCompare.ListItems.Count > 0 Then
        Call lvWeaponCompare_ItemClick(lvWeaponCompare.ListItems(1))
    End If
    
    sCompares = ReadINI(sSectionName, "ArmourCompare", sFile)
    x = 0
    Do While Not InStr(x + 1, sCompares, ",") = 0
        y = InStr(x + 1, sCompares, ",")
        
        tabItems.Index = "pkItems"
        tabItems.Seek "=", Val(Mid(sCompares, x + 1, y - x - 1))
        If tabItems.NoMatch = False Then
            Call AddArmour2LV(lvArmourCompare)
        End If
        x = y
    Loop
    If lvArmourCompare.ListItems.Count > 0 Then
        Call lvArmourCompare_ItemClick(lvArmourCompare.ListItems(1))
    End If
    
    sCompares = ReadINI(sSectionName, "SpellCompare", sFile)
    x = 0
    Do While Not InStr(x + 1, sCompares, ",") = 0
        y = InStr(x + 1, sCompares, ",")
        
        tabSpells.Index = "pkSpells"
        tabSpells.Seek "=", Val(Mid(sCompares, x + 1, y - x - 1))
        If tabSpells.NoMatch = False Then
            Call AddSpell2LV(lvSpellCompare)
        End If
        x = y
    Loop
    If lvSpellCompare.ListItems.Count > 0 Then
        Call lvSpellCompare_ItemClick(lvSpellCompare.ListItems(1))
    End If
End If


canceled:
If Not bStartup Then bDontRefresh = False
Call RefreshAll
Me.Enabled = True
bPromptSave = False

Exit Sub
Error:
Call HandleError("LoadCharacter")
Me.Enabled = True
If Not bStartup Then bDontRefresh = False
End Sub

Private Sub LoadClasses()
On Error GoTo Error:
Dim x As Integer, oLI As ListItem

lvClasses.ListItems.clear

cmbGlobalClass(x).clear
'cmbClass.Clear

If tabClasses.RecordCount = 0 Then
    For x = 0 To 1
        cmbGlobalClass(x).AddItem "Warrior"
        cmbGlobalClass(x).ItemData(cmbGlobalClass(x).NewIndex) = 1
        cmbGlobalClass(x).AddItem "Witchunter"
        cmbGlobalClass(x).ItemData(cmbGlobalClass(x).NewIndex) = 2
        cmbGlobalClass(x).AddItem "Paladin"
        cmbGlobalClass(x).ItemData(cmbGlobalClass(x).NewIndex) = 3
        cmbGlobalClass(x).AddItem "Cleric"
        cmbGlobalClass(x).ItemData(cmbGlobalClass(x).NewIndex) = 4
        cmbGlobalClass(x).AddItem "Priest"
        cmbGlobalClass(x).ItemData(cmbGlobalClass(x).NewIndex) = 5
        cmbGlobalClass(x).AddItem "Missionary"
        cmbGlobalClass(x).ItemData(cmbGlobalClass(x).NewIndex) = 6
        cmbGlobalClass(x).AddItem "Ninja"
        cmbGlobalClass(x).ItemData(cmbGlobalClass(x).NewIndex) = 7
        cmbGlobalClass(x).AddItem "Thief"
        cmbGlobalClass(x).ItemData(cmbGlobalClass(x).NewIndex) = 8
        cmbGlobalClass(x).AddItem "Bard"
        cmbGlobalClass(x).ItemData(cmbGlobalClass(x).NewIndex) = 9
        cmbGlobalClass(x).AddItem "Gypsy"
        cmbGlobalClass(x).ItemData(cmbGlobalClass(x).NewIndex) = 10
        cmbGlobalClass(x).AddItem "Warlock"
        cmbGlobalClass(x).ItemData(cmbGlobalClass(x).NewIndex) = 11
        cmbGlobalClass(x).AddItem "Mage"
        cmbGlobalClass(x).ItemData(cmbGlobalClass(x).NewIndex) = 12
        cmbGlobalClass(x).AddItem "Druid"
        cmbGlobalClass(x).ItemData(cmbGlobalClass(x).NewIndex) = 13
        cmbGlobalClass(x).AddItem "Ranger"
        cmbGlobalClass(x).ItemData(cmbGlobalClass(x).NewIndex) = 14
        cmbGlobalClass(x).AddItem "Mystic"
        cmbGlobalClass(x).ItemData(cmbGlobalClass(x).NewIndex) = 15
        
        cmbGlobalClass(x).AddItem "Any", 0
        cmbGlobalClass(x).ListIndex = 0
        
        Select Case x
            Case 0:
                Call ExpandCombo(cmbGlobalClass(x), HeightOnly, TripleWidth, fraChar(0).hwnd)
            Case 1:
                Call ExpandCombo(cmbGlobalClass(x), HeightOnly, TripleWidth, frmGlobalFilter.hwnd)
        End Select
        
        Exit Sub
    Next x
End If

tabClasses.MoveFirst

Do Until tabClasses.EOF

    For x = 0 To 1
        cmbGlobalClass(x).AddItem tabClasses.Fields("Name") ', tabClasses.Fields("Number")
        cmbGlobalClass(x).ItemData(cmbGlobalClass(x).NewIndex) = tabClasses.Fields("Number")
    Next x
    
    Call AddClass2LV(lvClasses)

    tabClasses.MoveNext
Loop

For x = 0 To 1
    cmbGlobalClass(x).AddItem "Any", 0
    cmbGlobalClass(x).ListIndex = 0
    
    Select Case x
        Case 0:
            Call ExpandCombo(cmbGlobalClass(x), HeightOnly, TripleWidth, fraChar(0).hwnd)
        Case 1:
            Call ExpandCombo(cmbGlobalClass(x), HeightOnly, TripleWidth, frmGlobalFilter.hwnd)
    End Select
Next x

bKeepSortOrder = True
Call lvClasses_ColumnClick(lvClasses.ColumnHeaders(2))

If lvClasses.ListItems.Count >= 1 Then Call lvClasses_ItemClick(lvClasses.ListItems(1))

lvClasses.Refresh

Set oLI = Nothing

Exit Sub
Error:
HandleError
Set oLI = Nothing

End Sub

Private Function LoadInfo() As Boolean
On Error GoTo Error:
Dim nNMRVer As Double

If tabInfo.RecordCount = 0 Then
    MsgBox "Error opening info table of database.", vbCritical + vbOKOnly
    Exit Function
End If

tabInfo.MoveFirst
lblDatVer.Caption = tabInfo.Fields("Dat File Version")
If LCase(lblDatVer.Caption) = "v1.11l" Then lblDatVer.Caption = "v1.11L"

fraDatVer.Caption = "Database Version (Created " & tabInfo.Fields("Date") & ")"

nNMRVer = Val(ExtractNumbersFromString(tabInfo.Fields("NMR Version")))
If nNMRVer < 1.55 Then
    MsgBox sNormalCaption & " introduces new features that this database does not support." _
        & vbCrLf & "Please download/open a newer database.", vbCritical + vbOKOnly
    Exit Function
End If

On Error GoTo skip:

If tabInfo.Fields("Legit") = 1 Then
    bLegit = True
Else
    If Not tabInfo.Fields("Custom") = Chr(0) Then
        lblDatVer.Caption = tabInfo.Fields("Custom") & " - " & lblDatVer.Caption
    End If
    bLegit = False
End If

skip:
LoadInfo = True
Exit Function
Error:
Call HandleError("LoadInfo")
End Function

Private Sub LoadInvenItems()
On Error GoTo Error:
If tabItems.RecordCount = 0 Then Exit Sub

Call InvenSetupEquip
tabItems.MoveFirst
DoEvents

Do Until tabItems.EOF
    If bOnlyInGame And tabItems.Fields("In Game") = 0 Then GoTo skip:
    Call InvenAddEquip(tabItems.Fields("Number"), tabItems.Fields("Name"), tabItems.Fields("ItemType"), tabItems.Fields("Worn"))
skip:
    tabItems.MoveNext
Loop

cmdNav(4).Caption = "Equipment"

Exit Sub
Error:
Call HandleError("LoadInvenItems")
End Sub

Private Sub LoadItems()
On Error GoTo Error:
Dim oLI As ListItem, x As Integer, sName As String

If tabItems.RecordCount = 0 Then Exit Sub
tabItems.MoveFirst

'--[weapon]
'
'damage | speed | str | ac | acc | bs acc | enc | level | limit
'special
'
'--[armour]
'
'type | level | enc | ac | location | acc | limit
'special

lvArmour.ListItems.clear
lvWeapons.ListItems.clear
lvOtherItems.ListItems.clear
Call InvenSetupEquip

DoEvents
Do Until tabItems.EOF
    
    If bOnlyInGame And tabItems.Fields("In Game") = 0 Then GoTo skip:
    
    sName = tabItems.Fields("Name")
    If sName = "" Or Left(sName, 3) = "sdf" Then GoTo skip:
    
    Select Case tabItems.Fields("ItemType")
        Case 0: 'armour
            If tabItems.Fields("Worn") = 0 Then
                Call AddOtherItem2LV(lvOtherItems)
            Else
                Call AddArmour2LV(lvArmour, True)
            End If
            
        Case 1: 'weapons
            Call AddWeapon2LV(lvWeapons, True)
            
        Case Else: 'other
            Call AddOtherItem2LV(lvOtherItems)
            
    End Select

skip:
    tabItems.MoveNext
    'DoEvents
Loop


bKeepSortOrder = True
Call lvWeapons_ColumnClick(lvWeapons.ColumnHeaders(2))
bKeepSortOrder = True
Call lvArmour_ColumnClick(lvArmour.ColumnHeaders(2))
bKeepSortOrder = True
Call lvOtherItems_ColumnClick(lvOtherItems.ColumnHeaders(2))

If lvWeapons.ListItems.Count >= 1 Then Call lvWeapons_ItemClick(lvWeapons.ListItems(1))
If lvArmour.ListItems.Count >= 1 Then Call lvArmour_ItemClick(lvArmour.ListItems(1))
If lvOtherItems.ListItems.Count >= 1 Then Call lvOtherItems_ItemClick(lvOtherItems.ListItems(1))

'lvWeapons.Refresh
'lvArmour.Refresh
'lvOtherItems.Refresh

Set oLI = Nothing

Exit Sub
Error:
HandleError
Set oLI = Nothing

End Sub

Private Sub LoadMonsters()
On Error GoTo Error:
Dim x As Integer, oLI As ListItem

lvMonsters.ListItems.clear

tabMonsters.MoveFirst
Do Until tabMonsters.EOF

    If bOnlyInGame And tabMonsters.Fields("In Game") = 0 Then GoTo skip:
    
    Call AddMonster2LV(lvMonsters)

skip:
    tabMonsters.MoveNext
Loop

bKeepSortOrder = True
Call lvMonsters_ColumnClick(lvMonsters.ColumnHeaders(2))

If lvMonsters.ListItems.Count >= 1 Then Call lvMonsters_ItemClick(lvMonsters.ListItems(1))

lvMonsters.Refresh

Set oLI = Nothing

Exit Sub
Error:
Call HandleError("LoadMonsters")
Set oLI = Nothing

End Sub

Public Sub LoadPresets(Optional ByVal bReset As Boolean)
Dim x As Integer, sSectionName As String, nMap As Long, nRoom As Long, sName As String
Dim cReg As clsRegistryRoutines, nError As Integer, bResult As Boolean

On Error GoTo Error:

Set cReg = New clsRegistryRoutines

If bLegit Then
    sSectionName = "Custom_Presets"
Else
    sSectionName = RemoveCharacter(lblDatVer.Caption, " ") & "_Presets"
End If

cReg.hkey = HKEY_LOCAL_MACHINE
cReg.KeyRoot = "Software\MMUD Explorer\Presets"
cReg.Subkey = sSectionName

'If cReg.KeyExists = False Then
'    nError = RegCreateKeyPath(HKEY_LOCAL_MACHINE, "Software\MMUD Explorer\Presets\" & sSectionName)
'    If nError > 0 Then GoTo Error:
'
'    For x = 0 To 49
'        nMap = Val(ReadINI(sSectionName, "Map" & x))
'        nRoom = Val(ReadINI(sSectionName, "Room" & x))
'        sName = ReadINI(sSectionName, "Name" & x)
'
'        If nMap = 0 Or nRoom = 0 Or sName = "" Then
'            Select Case x
'                Case 0: nMap = 10: nRoom = 271: sName = "Aged Titan"
'                Case 1: nMap = 3: nRoom = 560: sName = "Ancient Ruin"
'                Case 2: nMap = 17: nRoom = 2269: sName = "Arlysia"
'                Case 3: nMap = 7: nRoom = 1176: sName = "Black Fortress"
'                Case 4: nMap = 3: nRoom = 669: sName = "Black Wastelands"
'                Case 5: nMap = 17: nRoom = 241: sName = "Blackwood Graveyard"
'                Case 6: nMap = 8: nRoom = 461: sName = "Dark-Elf Castle"
'                Case 7: nMap = 6: nRoom = 552: sName = "Gnome Village"
'                Case 8: nMap = 12: nRoom = 1919: sName = "Great Pyramid"
'                Case 9: nMap = 6: nRoom = 1255: sName = "Khazarad"
'                Case 10: nMap = 7: nRoom = 884: sName = "Lava Fields"
'                Case 11: nMap = 16: nRoom = 454: sName = "Lost City"
'                Case 12: nMap = 12: nRoom = 5: sName = "Nekojin Village"
'                Case 13: nMap = 2: nRoom = 2523: sName = "Rhudar"
'                Case 14: nMap = 12: nRoom = 2099: sName = "Saracen Fort"
'                Case 15: nMap = 12: nRoom = 1173: sName = "Small Pyramid"
'                Case 16: nMap = 16: nRoom = 1179: sName = "Storm Fortress"
'                Case 17: nMap = 16: nRoom = 1: sName = "Tasloi Village"
'                Case 18: nMap = 1: nRoom = 224: sName = "Town Square"
'                Case 19: nMap = 16: nRoom = 1990: sName = "Volcano"
'                Case Else: nMap = 1: nRoom = 1: sName = "unset"
'            End Select
'
'            Call cReg.SetRegistryValue("Map" & x, nMap, REG_SZ)
'            Call cReg.SetRegistryValue("Room" & x, nRoom, REG_SZ)
'            Call cReg.SetRegistryValue("Name" & x, sName, REG_SZ)
'        End If
'
'    Next x
'End If

nError = RegCreateKeyPath(HKEY_LOCAL_MACHINE, "Software\MMUD Explorer\Presets\" & sSectionName)
If nError > 0 Then GoTo Error:

If bReset Then
    For x = 0 To 49
        bResult = cReg.SetRegistryValue("Map" & x, "0", REG_SZ)
        If bResult = False Then Err.Raise 0, "LoadPresets", "Error Setting Registry Values"
        'Call WriteINI(sSectionName, "Map" & x, "0")
    Next
End If

For x = 0 To 49
    nMap = Val(cReg.GetRegistryValue("Map" & x, 0))
    nRoom = Val(cReg.GetRegistryValue("Room" & x, 0))
    sName = cReg.GetRegistryValue("Name" & x, 0)
    
    If nMap = 0 Or nRoom = 0 Or sName = "" Then
        Select Case x
            Case 0: nMap = 10: nRoom = 271: sName = "Aged Titan"
            Case 1: nMap = 3: nRoom = 560: sName = "Ancient Ruin"
            Case 2: nMap = 17: nRoom = 2269: sName = "Arlysia"
            Case 3: nMap = 7: nRoom = 1176: sName = "Black Fortress"
            Case 4: nMap = 3: nRoom = 669: sName = "Black Wastelands"
            Case 5: nMap = 17: nRoom = 241: sName = "Blackwood Graveyard"
            Case 6: nMap = 8: nRoom = 461: sName = "Dark-Elf Castle"
            Case 7: nMap = 6: nRoom = 552: sName = "Gnome Village"
            Case 8: nMap = 12: nRoom = 1919: sName = "Great Pyramid"
            Case 9: nMap = 6: nRoom = 1255: sName = "Khazarad"
            Case 10: nMap = 7: nRoom = 884: sName = "Lava Fields"
            Case 11: nMap = 16: nRoom = 454: sName = "Lost City"
            Case 12: nMap = 12: nRoom = 5: sName = "Nekojin Village"
            Case 13: nMap = 2: nRoom = 2523: sName = "Rhudar"
            Case 14: nMap = 12: nRoom = 2099: sName = "Saracen Fort"
            Case 15: nMap = 12: nRoom = 1173: sName = "Small Pyramid"
            Case 16: nMap = 16: nRoom = 1179: sName = "Storm Fortress"
            Case 17: nMap = 16: nRoom = 1: sName = "Tasloi Village"
            Case 18: nMap = 1: nRoom = 224: sName = "Town Square"
            Case 19: nMap = 16: nRoom = 1990: sName = "Volcano"
            Case Else: nMap = 1: nRoom = 1: sName = "unset"
        End Select
        
        Call cReg.SetRegistryValue("Map" & x, nMap, REG_SZ)
        Call cReg.SetRegistryValue("Room" & x, nRoom, REG_SZ)
        Call cReg.SetRegistryValue("Name" & x, sName, REG_SZ)
        
'        Call WriteINI(sSectionName, "Map" & x, nMap)
'        Call WriteINI(sSectionName, "Room" & x, nRoom)
'        Call WriteINI(sSectionName, "Name" & x, sName)
    End If
    
Next x

For x = 0 To 9
    cmdMapPreset(x).Caption = cReg.GetRegistryValue("Name" & x, "unset")
    cmdMapPreset(x).Tag = x
Next x

If FormIsLoaded("frmMap") Then
    Call frmMap.LoadPresets
End If

out:
Exit Sub
Error:
Call HandleError("LoadPresets")
Resume out:

End Sub

Private Sub LoadRaces()
On Error GoTo Error:
Dim x As Integer, oLI As ListItem


lvRaces.ListItems.clear

cmbGlobalRace(x).clear
'cmbRaces.Clear
If tabRaces.RecordCount = 0 Then
    For x = 0 To 0
        cmbGlobalRace(x).AddItem "Human"
        cmbGlobalRace(x).ItemData(cmbGlobalRace(x).NewIndex) = 1
        cmbGlobalRace(x).AddItem "Dwarf"
        cmbGlobalRace(x).ItemData(cmbGlobalRace(x).NewIndex) = 2
        cmbGlobalRace(x).AddItem "Gnome"
        cmbGlobalRace(x).ItemData(cmbGlobalRace(x).NewIndex) = 3
        cmbGlobalRace(x).AddItem "Halfling"
        cmbGlobalRace(x).ItemData(cmbGlobalRace(x).NewIndex) = 4
        cmbGlobalRace(x).AddItem "Elf"
        cmbGlobalRace(x).ItemData(cmbGlobalRace(x).NewIndex) = 5
        cmbGlobalRace(x).AddItem "Half-Elf"
        cmbGlobalRace(x).ItemData(cmbGlobalRace(x).NewIndex) = 6
        cmbGlobalRace(x).AddItem "Dark-Elf"
        cmbGlobalRace(x).ItemData(cmbGlobalRace(x).NewIndex) = 7
        cmbGlobalRace(x).AddItem "Half-Orc"
        cmbGlobalRace(x).ItemData(cmbGlobalRace(x).NewIndex) = 8
        cmbGlobalRace(x).AddItem "Goblin"
        cmbGlobalRace(x).ItemData(cmbGlobalRace(x).NewIndex) = 9
        cmbGlobalRace(x).AddItem "Half-Ogre"
        cmbGlobalRace(x).ItemData(cmbGlobalRace(x).NewIndex) = 10
        cmbGlobalRace(x).AddItem "Kang"
        cmbGlobalRace(x).ItemData(cmbGlobalRace(x).NewIndex) = 11
        cmbGlobalRace(x).AddItem "Nekojin"
        cmbGlobalRace(x).ItemData(cmbGlobalRace(x).NewIndex) = 12
        cmbGlobalRace(x).AddItem "Gaunt One"
        cmbGlobalRace(x).ItemData(cmbGlobalRace(x).NewIndex) = 13
       
        cmbGlobalRace(x).AddItem "Any", 0
        cmbGlobalRace(x).ListIndex = 0
        
        Call ExpandCombo(cmbGlobalRace(x), HeightOnly, TripleWidth, fraChar(0).hwnd)
    Next x
    
    Exit Sub
End If


tabRaces.MoveFirst
Do Until tabRaces.EOF
        
    For x = 0 To 0
      cmbGlobalRace(x).AddItem tabRaces.Fields("Name") ', tabRaces.Fields("Number")
      cmbGlobalRace(x).ItemData(cmbGlobalRace(x).NewIndex) = tabRaces.Fields("Number")
    Next x
    
    Call AddRace2LV(lvRaces)
    tabRaces.MoveNext
Loop

For x = 0 To 0
    cmbGlobalRace(x).AddItem "Any", 0
    cmbGlobalRace(x).ListIndex = 0
    
    Call ExpandCombo(cmbGlobalRace(x), HeightOnly, TripleWidth, fraChar(0).hwnd)
Next x

bKeepSortOrder = True
Call lvRaces_ColumnClick(lvRaces.ColumnHeaders(2))

If lvRaces.ListItems.Count >= 1 Then Call lvRaces_ItemClick(lvRaces.ListItems(1))

lvRaces.Refresh

Call RaceColorCode(lvRaces)

Set oLI = Nothing

Exit Sub
Error:
HandleError
Set oLI = Nothing

End Sub

Public Sub LoadSettings()
Dim sSectionName As String, sName As String, nSize As Integer, bBold As Boolean, bItalic As Boolean
Dim x As Integer, sFileTitle() As String, nLng As Long

On Error GoTo Error:

sSectionName = RemoveCharacter(lblDatVer.Caption, " ")

Call CheckINIReadOnly
Call LoadPresets

If ReadINI("Settings", "AutoSaveLastChar") = "1" Then
    bAutoSave = True
Else
    bAutoSave = False
End If

If ReadINI("Settings", "NameInTitle") = "1" Then
    bNameInTitle = True
Else
    bNameInTitle = False
End If

If ReadINI("Settings", "Use2ndWrist") = "1" Then
    chkEquipHold(7).Enabled = True
    cmdEquipGoto(7).Enabled = True
    cmbEquip(7).Enabled = True
    bInvenUse2ndWrist = True
Else
    chkEquipHold(7).Enabled = False
    cmdEquipGoto(7).Enabled = False
    cmbEquip(7).Enabled = False
    bInvenUse2ndWrist = False
End If

If ReadINI("Settings", "SwapMapButtons") = "1" Then
    bMapSwapButtons = True
Else
    bMapSwapButtons = False
End If

If ReadINI("Settings", "NoAlwaysOnTop") = "1" Then
    bNoAlwaysOnTop = True
Else
    bNoAlwaysOnTop = False
End If

If Not Len(ReadINI("Settings", "LabelFontName")) < 3 Then
    sName = ReadINI("Settings", "LabelFontName")
    nSize = Int(ReadINI("Settings", "LabelFontSize"))
    If ReadINI("Settings", "LabelFontBold") = "True" Then
        bBold = True
    Else
        bBold = False
    End If
    If ReadINI("Settings", "LabelFontItal") = "True" Then
        bItalic = True
    Else
        bItalic = False
    End If
    Call SetLabelFonts(sName, nSize, bBold, bItalic)
End If
If Not Len(ReadINI("Settings", "StatFontName")) < 3 Then
    sName = ReadINI("Settings", "StatFontName")
    nSize = Int(ReadINI("Settings", "StatFontSize"))
    If ReadINI("Settings", "StatFontBold") = "True" Then
        bBold = True
    Else
        bBold = False
    End If
    If ReadINI("Settings", "StatFontItal") = "True" Then
        bItalic = True
    Else
        bItalic = False
    End If
    Call SetStatFonts(sName, nSize, bBold, bItalic)
End If

'character
If ReadINI("Settings", "AutoLoadLastChar") = "1" Then
    If FileExists(ReadINI(sSectionName, "LastCharFile")) Then
        Call LoadCharacter(False, ReadINI(sSectionName, "LastCharFile"))
    Else
        Call LoadCharacter(False)
    End If
Else
    Call LoadCharacter(False)
End If

'settings
For x = 1 To 5
    sRecentFiles(x, 2) = ReadINI("Settings", "Recent" & x)
Next x
Call RecentFileAdd

If Val(txtGlobalLevel(0).Text) = 0 Then txtGlobalLevel(0).Text = 99

If ReadINI("Settings", "DontSpanNavButtons") = "1" Then
    bDontSpanNav = True
Else
    bDontSpanNav = False
End If

nLng = Val(ReadINI("Settings", "Top"))
If nLng > 0 Then Me.Top = nLng
nLng = Val(ReadINI("Settings", "Left"))
If nLng > 0 Then Me.Left = nLng

Me.Width = Val(ReadINI("Settings", "Width"))
Me.Height = Val(ReadINI("Settings", "Height"))
If Val(ReadINI("Settings", "Maximized")) = 1 Then
    Me.WindowState = vbMaximized
End If
DoEvents

chkMapOptions(0).Value = ReadINI("Settings", "MapFollowMap")
chkMapOptions(1).Value = ReadINI("Settings", "MapNoHidden")
chkMapOptions(2).Value = ReadINI("Settings", "MapNoLairs")
chkMapOptions(3).Value = ReadINI("Settings", "MapNoNPC")
chkMapOptions(4).Value = ReadINI("Settings", "MapNoCMD")
chkMapOptions(5).Value = ReadINI("Settings", "MapNoTooltips")
chkMonstersNoRegenLookUp.Value = ReadINI("Setting", "LookUpMonsterRegen")
chkGlobalFilter.Value = ReadINI("Settings", "UseGlobalFilter")

If Val(ReadINI("Settings", "JumpToCompare")) = 1 Then
    mnuJumpToCompare.Checked = True
Else
    mnuJumpToCompare.Checked = False
End If

If Val(ReadINI("Settings", "FilterAll")) = 1 Then Call FilterAll(True)

Erase sFileTitle()
Exit Sub
Error:
Call HandleError("LoadSettings")
Resume Next
End Sub

Private Sub LoadShops()
On Error GoTo Error:
Dim x As Integer, oLI As ListItem


lvShops.ListItems.clear

tabShops.MoveFirst
Do Until tabShops.EOF
    If bOnlyInGame And tabShops.Fields("In Game") = 0 Then GoTo skip:
    
    Call AddShop2LV(lvShops)

skip:
    tabShops.MoveNext
Loop

bKeepSortOrder = True
Call lvShops_ColumnClick(lvShops.ColumnHeaders(2))

If lvShops.ListItems.Count >= 1 Then Call lvShops_ItemClick(lvShops.ListItems(1))

lvShops.Refresh

Set oLI = Nothing

Exit Sub
Error:
HandleError
Set oLI = Nothing

End Sub

Private Sub LoadSpells()
On Error GoTo Error:
Dim oLI As ListItem, x As Integer

If tabSpells.RecordCount = 0 Then Exit Sub
tabSpells.MoveFirst

'num
'name
'Short
'Magery
'Level
'mana
'difficulty
Call SetupCharBless
lvSpells.ListItems.clear

DoEvents
Do Until tabSpells.EOF
    
    'If bOnlyLearnable And tabSpells.Fields("Learnable") = 0 Then GoTo skip:
    
    Call AddSpell2LV(lvSpells, True)

skip:
    tabSpells.MoveNext
    'DoEvents
Loop

bKeepSortOrder = True
Call lvSpells_ColumnClick(lvSpells.ColumnHeaders(nLastSpellSort))

If lvSpells.ListItems.Count >= 1 Then Call lvSpells_ItemClick(lvSpells.ListItems(1))

lvSpells.Refresh

Set oLI = Nothing

Exit Sub
Error:
HandleError
Set oLI = Nothing

End Sub

Public Sub LookUpMonsterRegen(ByVal nNumber As Long, _
    bInResults As Boolean, Optional ByRef DestLV As ListView)
On Error GoTo Error:
'Dim sLoc As String, x As Long, y As Long, oLI As ListItem, tRoom As RoomExitType
Dim bHideTemp As Boolean

If nNumber < 1 Then Exit Sub

If bInResults Then
    Set DestLV = frmResults.lvResults
Else
    If DestLV Is Nothing Then Exit Sub
End If

tabMonsters.Index = "pkMonsters"
tabMonsters.Seek "=", nNumber
If tabMonsters.NoMatch Then Exit Sub
If Len(tabMonsters.Fields("Summoned By")) < 5 Then Exit Sub

If bInResults Then
    'Load frmResults
    Call frmResults.SetupResultsWindow(False, Me, nMapStartMap)
    DestLV.ListItems.clear
End If

bHideTemp = bHideRecordNumbers
bHideRecordNumbers = False

If bInResults Then
    Call GetLocations(tabMonsters.Fields("Summoned By"), DestLV, True, , , , True)
Else
    Call GetLocations(tabMonsters.Fields("Summoned By"), DestLV, True, , , True, True)
End If

bHideRecordNumbers = bHideTemp

'sLoc = tabMonsters.Fields("Summoned By")
'x = 0
'Do While Not InStr(x + 1, sLoc, ",") = 0
'    y = InStr(x + 1, sLoc, ",")
'
'    tRoom = ExtractMapRoom(Mid(sLoc, x + 1, y - x - 1))
'    tabRooms.Index = "idxRooms"
'    tabRooms.Seek "=", tRoom.Map, tRoom.Room
'    If tabRooms.NoMatch = False Then
'        Set oLI = DestLV.ListItems.Add()
'        oLI.Text = "Room: " & tabRooms.Fields("Name") & " (" & tRoom.Map & "/" & tRoom.Room & ")"
'        oLI.Tag = tRoom.Map & "/" & tRoom.Room
'    End If
'    x = y
'Loop

If bInResults Then
    If DestLV.ListItems.Count > 0 Then
        frmResults.lblCaption = "Monster " & GetMonsterName(nNumber, bHideRecordNumbers) & " is summoned by:"
        frmResults.Show vbModeless, IIf(bNoAlwaysOnTop, Nothing, Me)
    Else
        Unload frmResults
        DoEvents
        MsgBox "Unknown.", vbOKOnly + vbInformation
    End If
End If

Exit Sub

Error:
Call HandleError
bHideRecordNumbers = bHideTemp
End Sub

Private Sub LookUpSpellCast(objLV As ListView)

On Error GoTo Error:

If objLV.SelectedItem Is Nothing Then Exit Sub

tabSpells.Index = "pkSpells"
tabSpells.Seek "=", Val(objLV.SelectedItem.Text)
If tabSpells.NoMatch Then
    MsgBox "Cannot find Spell."
End If

If Len(tabSpells.Fields("Casted By")) < 5 Then
    MsgBox "Nothing.", vbOKOnly + vbInformation
    Exit Sub
End If

'Load frmResults
Call frmResults.SetupResultsWindow(False, Me, nMapStartMap)
Call GetLocations(tabSpells.Fields("Casted By"), frmResults.lvResults)
If frmResults.lvResults.ListItems.Count > 0 Then
    frmResults.lblCaption = "Spell " & GetSpellName(Val(objLV.SelectedItem.Text), bHideRecordNumbers) & " is casted by:"
    frmResults.Show vbModeless, IIf(bNoAlwaysOnTop, Nothing, Me)
Else
    Unload frmResults
    DoEvents
    MsgBox "Nothing.", vbOKOnly + vbInformation
End If

Exit Sub

Error:
Call HandleError("LookUpSpellCast")

End Sub

Private Sub lvArmour_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim bSort As Boolean, nSort As ListDataType

nLastArmourSort = ColumnHeader.Index
If bKeepSortOrder Then
    bSort = IIf(lvArmour.SortOrder = lvwDescending, False, True)
    bKeepSortOrder = False
Else
    If oLastColumnSorted Is ColumnHeader Then
        If bSortOrderAsc = True Then
            bSortOrderAsc = False
        Else
            bSortOrderAsc = True
        End If
    End If
    bSort = bSortOrderAsc
    Set oLastColumnSorted = ColumnHeader
End If

Select Case ColumnHeader.Index
    Case 1, 5, 6, 7, 8, 9, 10: nSort = ldtnumber
    Case Else: nSort = ldtstring
End Select
SortListView lvArmour, ColumnHeader.Index, nSort, bSort
End Sub

Public Sub lvArmour_ItemClick(ByVal item As MSComctlLib.ListItem)

Set lvArmour.SelectedItem = item
Call ProcessListViewClick(item, txtArmourDetail, lvArmourLoc)

End Sub

Private Sub lvArmour_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
     Call PopUpItemsMenu(lvArmour)
End If
End Sub

Private Sub lvArmourCompare_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim bSort As Boolean, nSort As ListDataType

nLastArmourSort = ColumnHeader.Index
If bKeepSortOrder Then
    bSort = IIf(lvArmourCompare.SortOrder = lvwDescending, False, True)
    bKeepSortOrder = False
Else
    If oLastColumnSorted Is ColumnHeader Then
        If bSortOrderAsc = True Then
            bSortOrderAsc = False
        Else
            bSortOrderAsc = True
        End If
    End If
    bSort = bSortOrderAsc
    Set oLastColumnSorted = ColumnHeader
End If

Select Case ColumnHeader.Index
    Case 1, 5, 6, 7, 8, 9, 10: nSort = ldtnumber
    Case Else: nSort = ldtstring
End Select
SortListView lvArmourCompare, ColumnHeader.Index, nSort, bSort
End Sub

Public Sub lvArmourCompare_ItemClick(ByVal item As MSComctlLib.ListItem)

Set lvArmourCompare.SelectedItem = item
Call ProcessListViewClick(item, txtArmourCompareDetail, lvArmourCompareLoc)

End Sub

'********************* /weapon compare


'********************* armour compare

Private Sub lvArmourCompare_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
     Call PopUpItemsMenu(lvArmourCompare)
End If
End Sub

Private Sub lvArmourCompareLoc_DblClick()
If lvArmourCompareLoc.ListItems.Count = 0 Then Exit Sub
Call GotoLocation(lvArmourCompareLoc.SelectedItem, Val(lvArmourCompare.SelectedItem.Text))
End Sub

Private Sub lvArmourLoc_DblClick()
If lvArmourLoc.ListItems.Count = 0 Then Exit Sub
Call GotoLocation(lvArmourLoc.SelectedItem, Val(lvArmour.SelectedItem.Text))
End Sub

Private Sub lvClasses_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim bSort As Boolean, nSort As ListDataType

If bKeepSortOrder Then
    bSort = IIf(lvClasses.SortOrder = lvwDescending, False, True)
    bKeepSortOrder = False
Else
    If oLastColumnSorted Is ColumnHeader Then
        If bSortOrderAsc = True Then
            bSortOrderAsc = False
        Else
            bSortOrderAsc = True
        End If
    End If
    bSort = bSortOrderAsc
    Set oLastColumnSorted = ColumnHeader
End If

Select Case ColumnHeader.Index
    Case 1, 3, 7, 8: nSort = ldtnumber
    Case Else: nSort = ldtstring
End Select
SortListView lvClasses, ColumnHeader.Index, nSort, bSort
End Sub

Private Sub lvClasses_ItemClick(ByVal item As MSComctlLib.ListItem)
On Error GoTo Error:

Call PullClassDetail(Val(item.Text), txtClassDetail)

item.Selected = True
item.EnsureVisible

Exit Sub

Error:
Call HandleError

End Sub

Private Sub lvClasses_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Call PopUpAuxMenu(lvClasses)
End If
End Sub

Private Sub lvMapLoc_DblClick()
If lvMapLoc.ListItems.Count = 0 Then Exit Sub
Call GotoLocation(lvMapLoc.SelectedItem)
End Sub

Private Sub lvMonsterCompare_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim bSort As Boolean, nSort As ListDataType

If bKeepSortOrder Then
    bSort = IIf(lvMonsterCompare.SortOrder = lvwDescending, False, True)
    bKeepSortOrder = False
Else
    If oLastColumnSorted Is ColumnHeader Then
        If bSortOrderAsc = True Then
            bSortOrderAsc = False
        Else
            bSortOrderAsc = True
        End If
    End If
    bSort = bSortOrderAsc
    Set oLastColumnSorted = ColumnHeader
End If

Select Case ColumnHeader.Index
    Case 2: nSort = ldtstring
    Case 3, 4:
        Call SortListViewByTag(lvMonsterCompare, ColumnHeader.Index, ldtnumber, bSort)
        Exit Sub
    Case Else: nSort = ldtnumber
End Select
SortListView lvMonsterCompare, ColumnHeader.Index, nSort, bSort
End Sub

Private Sub lvMonsterCompare_ItemClick(ByVal item As MSComctlLib.ListItem)
On Error GoTo Error:

Set lvMonsterCompare.SelectedItem = item
Call PullMonsterDetail(Val(item.Text), lvMonsterCompareLoc) ', txtMonsterCompareDetail)

item.Selected = True
item.EnsureVisible

Exit Sub

Error:
Call HandleError
End Sub

Private Sub lvMonsterCompare_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
     Call PopUpAuxMenu(lvMonsterCompare)
End If
End Sub

Private Sub lvMonsterCompareLoc_DblClick()
Dim tRoomExits As RoomExitType

Select Case LCase(lvMonsterCompareLoc.SelectedItem.Tag)
    Case "item":
        Call GotoItem(Val(lvMonsterCompareLoc.SelectedItem.ListSubItems(1).Tag))
    Case "spell":
        Call GotoSpell(Val(lvMonsterCompareLoc.SelectedItem.ListSubItems(1).Tag))
    Case "greet_text":
        'Load frmResults
        lvMonsters.SetFocus
        Call frmResults.SetupResultsWindow(True, Me, nMapStartMap)
        Call frmResults.CreateCommandTree(Val(lvMonsterDetail.SelectedItem.ListSubItems(1).Tag), False, True)
        frmResults.Show vbModeless, IIf(bNoAlwaysOnTop, Nothing, Me)
        
    Case "textblock":
        lvMonsters.SetFocus
        Call frmResults.SetupResultsWindow(True, Me, nMapStartMap)
        Call frmResults.CreateExecutionTree(Val(lvMonsterDetail.SelectedItem.ListSubItems(1).Tag))
        frmResults.Show vbModeless, IIf(bNoAlwaysOnTop, Nothing, Me)
        
    Case "monster":
        Call GotoMonster(Val(lvMonsterCompareLoc.SelectedItem.ListSubItems(1).Tag))
    Case "room":
        tRoomExits = ExtractMapRoom(lvMonsterCompareLoc.SelectedItem.ListSubItems(1).Tag)
        If tRoomExits.Map > 0 And tRoomExits.Room > 0 Then
            Call cmdNav_Click(10)
            Call MapStartMapping(tRoomExits.Map, tRoomExits.Room)
        End If
End Select

End Sub

Private Sub lvMonsterDetail_DblClick()
Dim tRoomExits As RoomExitType

Select Case LCase(lvMonsterDetail.SelectedItem.Tag)
    Case "item":
        Call GotoItem(Val(lvMonsterDetail.SelectedItem.ListSubItems(1).Tag))
    Case "spell":
        Call GotoSpell(Val(lvMonsterDetail.SelectedItem.ListSubItems(1).Tag))
    Case "greet_text":
        'Load frmResults
        lvMonsters.SetFocus
        Call frmResults.SetupResultsWindow(True, Me, nMapStartMap)
        Call frmResults.CreateCommandTree(Val(lvMonsterDetail.SelectedItem.ListSubItems(1).Tag), False, True)
        frmResults.Show vbModeless, IIf(bNoAlwaysOnTop, Nothing, Me)
        
    Case "textblock":
        lvMonsters.SetFocus
        Call frmResults.SetupResultsWindow(True, Me, nMapStartMap)
        Call frmResults.CreateExecutionTree(Val(lvMonsterDetail.SelectedItem.ListSubItems(1).Tag))
        frmResults.Show vbModeless, IIf(bNoAlwaysOnTop, Nothing, Me)
        
    Case "monster":
        Call GotoMonster(Val(lvMonsterDetail.SelectedItem.ListSubItems(1).Tag))
    Case "room":
        tRoomExits = ExtractMapRoom(lvMonsterDetail.SelectedItem.ListSubItems(1).Tag)
        If tRoomExits.Map > 0 And tRoomExits.Room > 0 Then
            Call cmdNav_Click(10)
            Call MapStartMapping(tRoomExits.Map, tRoomExits.Room)
        End If
End Select

End Sub

Private Sub lvMonsterDetail_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Call PopUpAuxMenu(lvMonsters)
End If
End Sub

Private Sub lvMonsters_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim bSort As Boolean, nSort As ListDataType

If bKeepSortOrder Then
    bSort = IIf(lvMonsters.SortOrder = lvwDescending, False, True)
    bKeepSortOrder = False
Else
    If oLastColumnSorted Is ColumnHeader Then
        If bSortOrderAsc = True Then
            bSortOrderAsc = False
        Else
            bSortOrderAsc = True
        End If
    End If
    bSort = bSortOrderAsc
    Set oLastColumnSorted = ColumnHeader
End If

Select Case ColumnHeader.Index
    Case 2: nSort = ldtstring
    Case 3, 4:
        Call SortListViewByTag(lvMonsters, ColumnHeader.Index, ldtnumber, bSort)
        Exit Sub
    Case Else: nSort = ldtnumber
End Select
SortListView lvMonsters, ColumnHeader.Index, nSort, bSort
End Sub

Public Sub lvMonsters_ItemClick(ByVal item As MSComctlLib.ListItem)
On Error GoTo Error:

Set lvMonsters.SelectedItem = item

Call PullMonsterDetail(Val(item.Text), lvMonsterDetail) ', txtMonsterDetail)

item.Selected = True
item.EnsureVisible

Exit Sub

Error:
Call HandleError

End Sub

Private Sub lvMonsters_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Call PopUpAuxMenu(lvMonsters)
End If
End Sub

Private Sub lvOtherItemLoc_DblClick()
If lvOtherItemLoc.ListItems.Count = 0 Then Exit Sub
Call GotoLocation(lvOtherItemLoc.SelectedItem, Val(lvOtherItems.SelectedItem.Text))
End Sub

Private Sub lvOtherItemLoc_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 2 Then
    Call PopUpAuxMenu(lvOtherItemLoc)
End If

End Sub

Private Sub lvOtherItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim bSort As Boolean, nSort As ListDataType

If bKeepSortOrder Then
    bSort = IIf(lvOtherItems.SortOrder = lvwDescending, False, True)
    bKeepSortOrder = False
Else
    If oLastColumnSorted Is ColumnHeader Then
        If bSortOrderAsc = True Then
            bSortOrderAsc = False
        Else
            bSortOrderAsc = True
        End If
    End If
    bSort = bSortOrderAsc
    Set oLastColumnSorted = ColumnHeader
End If

Select Case ColumnHeader.Index
    Case 1, 4, 5: nSort = ldtnumber
    Case Else: nSort = ldtstring
End Select
SortListView lvOtherItems, ColumnHeader.Index, nSort, bSort
End Sub

Public Sub lvOtherItems_ItemClick(ByVal item As MSComctlLib.ListItem)
On Error GoTo Error:

Set lvOtherItems.SelectedItem = item

tabItems.Index = "pkItems"
tabItems.Seek "=", Val(item.Text)
If tabItems.NoMatch = True Then
    MsgBox "Record not found."
Else
    Call PullItemDetail(txtOtherItemDetail, lvOtherItemLoc)
End If

If Not lvOtherItems.SelectedItem Is Nothing Then
    If LCase(lvOtherItems.SelectedItem.ListSubItems(2)) = "container" Then
        cmdSundryChests.Enabled = True
    Else
        cmdSundryChests.Enabled = False
    End If
Else
    cmdSundryChests.Enabled = True
End If

item.Selected = True
item.EnsureVisible

Exit Sub

Error:
Call HandleError

End Sub

Private Sub lvOtherItems_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Call PopUpAuxMenu(lvOtherItems)
End If
End Sub

Private Sub lvRaces_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim bSort As Boolean, nSort As ListDataType

If bKeepSortOrder Then
    bSort = IIf(lvRaces.SortOrder = lvwDescending, False, True)
    bKeepSortOrder = False
Else
    If oLastColumnSorted Is ColumnHeader Then
        If bSortOrderAsc = True Then
            bSortOrderAsc = False
        Else
            bSortOrderAsc = True
        End If
    End If
    bSort = bSortOrderAsc
    Set oLastColumnSorted = ColumnHeader
End If

Select Case ColumnHeader.Index
    Case 2, 11: nSort = ldtstring
    Case Else: nSort = ldtnumber
End Select
SortListView lvRaces, ColumnHeader.Index, nSort, bSort
End Sub

Private Sub lvRaces_ItemClick(ByVal item As MSComctlLib.ListItem)

On Error GoTo Error:

Call PullRaceDetail(Val(item.Text), txtRaceDetail)

item.Selected = True
item.EnsureVisible

Exit Sub

Error:
Call HandleError

End Sub

Private Sub lvRaces_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Call PopUpAuxMenu(lvRaces)
End If
End Sub

Private Sub lvShopDetail_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'Dim nSort As ListDataType
Dim bSort As Boolean

If bKeepSortOrder Then
    bSort = IIf(lvShopDetail.SortOrder = lvwDescending, False, True)
    bKeepSortOrder = False
Else
    If oLastColumnSorted Is ColumnHeader Then
        If bSortOrderAsc = True Then
            bSortOrderAsc = False
        Else
            bSortOrderAsc = True
        End If
    End If
    bSort = bSortOrderAsc
    Set oLastColumnSorted = ColumnHeader
End If

If lvShopDetail.ColumnHeaders.Count = 2 Then 'training
    SortListViewByTag lvShopDetail, ColumnHeader.Index, ldtnumber, bSort
Else
    nLastShopSort = ColumnHeader.Index
    Select Case ColumnHeader.Index
        Case 2: 'name
            SortListView lvShopDetail, ColumnHeader.Index, ldtstring, bSort
        Case 5: 'cost
            SortListViewByTag lvShopDetail, ColumnHeader.Index, ldtnumber, bSort
        Case Else:
            SortListView lvShopDetail, ColumnHeader.Index, ldtnumber, bSort
    End Select
End If

End Sub
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%   /COLUMN CLICKS

Private Sub lvShopDetail_DblClick()
If lvShopDetail.SelectedItem Is Nothing Then Exit Sub

Call GotoItem(lvShopDetail.SelectedItem.Text)

End Sub

Public Sub lvShopDetail_ItemClick(ByVal item As MSComctlLib.ListItem)
item.Selected = True
item.EnsureVisible
nLastShopDetailIndex = item.Index
End Sub

Private Sub lvShopDetail_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Call PopUpItemsMenu(lvShopDetail)
End If
End Sub

Private Sub lvShopLoc_DblClick()
If lvShopLoc.ListItems.Count = 0 Then Exit Sub
Call GotoLocation(lvShopLoc.SelectedItem, Val(lvShopLoc.SelectedItem.Text))
End Sub

Private Sub lvShops_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim bSort As Boolean, nSort As ListDataType

If bKeepSortOrder Then
    bSort = IIf(lvShops.SortOrder = lvwDescending, False, True)
    bKeepSortOrder = False
Else
    If oLastColumnSorted Is ColumnHeader Then
        If bSortOrderAsc = True Then
            bSortOrderAsc = False
        Else
            bSortOrderAsc = True
        End If
    End If
    bSort = bSortOrderAsc
    Set oLastColumnSorted = ColumnHeader
End If

Select Case ColumnHeader.Index
    Case 1: nSort = ldtnumber
    Case Else:
        nSort = ldtstring
End Select
SortListView lvShops, ColumnHeader.Index, nSort, bSort
End Sub

'######################################## Item Clicks

Public Sub lvShops_ItemClick(ByVal item As MSComctlLib.ListItem)

On Error GoTo Error:

Set lvShops.SelectedItem = item

Call PullShopDetail(Val(item.Text), lvShopDetail, txtShopDetail, lvShopLoc, _
    IIf(chkShopShowCharm(0).Value = 1 Or chkShopShowCharm(1).Value = 1, _
    Val(txtCharStats(5).Text), 0), IIf(chkShopShowCharm(1).Value = 1, True, False))

item.Selected = True
item.EnsureVisible

bKeepSortOrder = True
If lvShopDetail.ColumnHeaders.Count = 2 Then 'training
    Call lvShopDetail_ColumnClick(lvShopDetail.ColumnHeaders(2))
Else
    Call lvShopDetail_ColumnClick(lvShopDetail.ColumnHeaders(nLastShopSort))
End If

If nLastShopDetailIndex < lvShopDetail.ListItems.Count Then
    Call ClearListViewSelections(lvShopDetail)
    Call lvShopDetail_ItemClick(lvShopDetail.ListItems(nLastShopDetailIndex))
End If

Exit Sub

Error:
Call HandleError("lvShops_ItemClick")

End Sub

'######################################## /Item Clicks


Private Sub lvShops_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Call PopUpAuxMenu(lvShops)
End If
End Sub

Private Sub lvSpellCompare_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim bSort As Boolean, nSort As ListDataType

On Error GoTo Error:

nLastSpellSort = ColumnHeader.Index
If bKeepSortOrder Then
    bSort = IIf(lvSpellCompare.SortOrder = lvwDescending, False, True)
    bKeepSortOrder = False
Else
    If oLastColumnSorted Is ColumnHeader Then
        If bSortOrderAsc = True Then
            bSortOrderAsc = False
        Else
            bSortOrderAsc = True
        End If
    End If
    bSort = bSortOrderAsc
    Set oLastColumnSorted = ColumnHeader
End If

Select Case ColumnHeader.Index
    Case 1, 5, 6, 7: nSort = ldtnumber
    Case Else: nSort = ldtstring
End Select
SortListView lvSpellCompare, ColumnHeader.Index, nSort, bSort

Exit Sub

Error:
Call HandleError("lvSpellCompare_ColumnClick")
End Sub

Public Sub lvSpellCompare_ItemClick(ByVal item As MSComctlLib.ListItem)
Dim oLI As ListItem
On Error GoTo Error:

Set lvSpellCompare.SelectedItem = item

tabSpells.Index = "pkSpells"
tabSpells.Seek "=", Val(item.Text)
If tabSpells.NoMatch = True Then
    MsgBox "Record not found."
Else
    Call PullSpellDetail(Val(item.Text), txtSpellCompareDetail, lvSpellCompareLoc)
End If

Set oLI = Nothing

Exit Sub

Error:
Call HandleError("lvSpellCompare_ItemClick")

End Sub

'*************************** /armour compare

'*************************** spell compare

Private Sub lvSpellCompare_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Call PopUpSpellsMenu(lvSpellCompare)
End If
End Sub

Private Sub lvSpellCompareLoc_DblClick()
If lvSpellCompareLoc.ListItems.Count = 0 Then Exit Sub
Call GotoLocation(lvSpellCompareLoc.SelectedItem)
End Sub

Private Sub lvSpellLoc_DblClick()
On Error GoTo Error:

If lvSpellLoc.ListItems.Count = 0 Then Exit Sub
If FormIsLoaded("frmMap") Then
    If frmMap.lvMapLoc.ListItems.Count > 0 Then
        If Val(frmMap.lvMapLoc.SelectedItem.Tag) = Val(lvSpells.SelectedItem.Text) Then
            Call GotoLocation(lvSpellLoc.SelectedItem, frmMap.nMapStartMap, frmMap)
            If frmMap.chkMapOptions(6).Value = 0 Then Call SetTopMostWindow(frmMap.hwnd, True)
            Exit Sub
        End If
    End If
End If

Call GotoLocation(lvSpellLoc.SelectedItem)

Exit Sub

Error:
Call HandleError("lvSpellLoc_DblClick")
End Sub

Private Sub lvSpells_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim bSort As Boolean, nSort As ListDataType

On Error GoTo Error:

nLastSpellSort = ColumnHeader.Index
If bKeepSortOrder Then
    bSort = IIf(lvSpells.SortOrder = lvwDescending, False, True)
    bKeepSortOrder = False
Else
    If oLastColumnSorted Is ColumnHeader Then
        If bSortOrderAsc = True Then
            bSortOrderAsc = False
        Else
            bSortOrderAsc = True
        End If
    End If
    bSort = bSortOrderAsc
    Set oLastColumnSorted = ColumnHeader
End If

Select Case ColumnHeader.Index
    Case 1, 5, 6, 7: nSort = ldtnumber
    Case Else: nSort = ldtstring
End Select
SortListView lvSpells, ColumnHeader.Index, nSort, bSort

Exit Sub

Error:
Call HandleError("lvSpells_ColumnClick")
End Sub

Public Sub lvSpells_ItemClick(ByVal item As MSComctlLib.ListItem)

Set lvSpells.SelectedItem = item
Call PullSpellDetail(Val(item.Text), txtSpellDetail, lvSpellLoc)

item.Selected = True
item.EnsureVisible

End Sub

Private Sub lvSpells_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Call PopUpSpellsMenu(lvSpells)
End If
End Sub

Private Sub lvWeaponCompare_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim bSort As Boolean, nSort As ListDataType

On Error GoTo Error:

nLastWeaponSort = ColumnHeader.Index
If bKeepSortOrder Then
    bSort = IIf(lvWeaponCompare.SortOrder = lvwDescending, False, True)
    bKeepSortOrder = False
Else
    If oLastColumnSorted Is ColumnHeader Then
        If bSortOrderAsc = True Then
            bSortOrderAsc = False
        Else
            bSortOrderAsc = True
        End If
    End If
    bSort = bSortOrderAsc
    Set oLastColumnSorted = ColumnHeader
End If

Select Case ColumnHeader.Index
    Case 1, 4, 5, 6, 7, 8, 10, 11, 12, 13: nSort = ldtnumber
    Case Else: nSort = ldtstring
End Select
SortListView lvWeaponCompare, ColumnHeader.Index, nSort, bSort

Exit Sub

Error:
Call HandleError("lvWeaponCompare_ColumnClick")
End Sub

Public Sub lvWeaponCompare_ItemClick(ByVal item As MSComctlLib.ListItem)

Set lvWeaponCompare.SelectedItem = item
Call ProcessListViewClick(item, txtWeaponCompareDetail, lvWeaponCompareLoc)

End Sub

'********************* weapon compare

Private Sub lvWeaponCompare_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
     Call PopUpItemsMenu(lvWeaponCompare)
End If
End Sub

Private Sub lvWeaponCompareLoc_DblClick()
If lvWeaponCompareLoc.ListItems.Count = 0 Then Exit Sub
Call GotoLocation(lvWeaponCompareLoc.SelectedItem, Val(lvWeaponCompare.SelectedItem.Text))
End Sub

Private Sub lvWeaponLoc_DblClick()
If lvWeaponLoc.ListItems.Count = 0 Then Exit Sub
Call GotoLocation(lvWeaponLoc.SelectedItem, Val(lvWeapons.SelectedItem.Text))
End Sub

Private Sub lvWeapons_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim bSort As Boolean, nSort As ListDataType

On Error GoTo Error:

nLastWeaponSort = ColumnHeader.Index
If bKeepSortOrder Then
    bSort = IIf(lvWeapons.SortOrder = lvwDescending, False, True)
    bKeepSortOrder = False
Else
    If oLastColumnSorted Is ColumnHeader Then
        If bSortOrderAsc = True Then
            bSortOrderAsc = False
        Else
            bSortOrderAsc = True
        End If
    End If
    bSort = bSortOrderAsc
    Set oLastColumnSorted = ColumnHeader
End If

Select Case ColumnHeader.Index
    Case 1, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14: nSort = ldtnumber
    Case Else: nSort = ldtstring
End Select
SortListView lvWeapons, ColumnHeader.Index, nSort, bSort

Exit Sub

Error:
Call HandleError("lvWeapons_ColumnClick")
End Sub

Public Sub lvWeapons_ItemClick(ByVal item As MSComctlLib.ListItem)

Set lvWeapons.SelectedItem = item
Call ProcessListViewClick(item, txtWeaponDetail, lvWeaponLoc)

End Sub

Private Sub lvWeapons_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Call PopUpItemsMenu(lvWeapons)
End If
End Sub

Private Function MapActivateCell(ByVal FromCell As Integer, ByVal direction As Integer, ByVal ExitType As Integer) As Integer
Dim temp As Integer, LineColor As Long

'0 = N = -20
'1 = S = +20
'2 = E = +1
'3 = W = -1
'4 = NE = -19
'5 = NW = -21
'6 = SE = +21
'7 = SW = +19

'figure out which cell is to be activated
On Error GoTo Error:

Select Case direction
    Case 0: 'north
        MapActivateCell = (FromCell - 20)
        'checking to see if it's on the north edge
        If MapActivateCell < 1 Then
            Call MapDrawOnRoom(lblRoomCell(FromCell), drLineN, 4, Grey)
            GoTo DontActivate
        End If

    Case 1: 'south
        MapActivateCell = (FromCell + 20)
        'checking to see if it's on the south edge
        If MapActivateCell > sMapSECorner Then
            Call MapDrawOnRoom(lblRoomCell(FromCell), drLineS, 4, Grey)
            GoTo DontActivate
        End If

    Case 2: 'east
        MapActivateCell = (FromCell + 1)
        'checking to see if it's on the east edge
        For temp = nMapRowLength To sMapSECorner Step 20
            If FromCell = temp Then
                Call MapDrawOnRoom(lblRoomCell(FromCell), drLineE, 4, Grey)
                GoTo DontActivate
            End If
        Next
        
    Case 3: 'west
        MapActivateCell = (FromCell - 1)
        'checking to see if it's on the west edge
        For temp = 1 To sMapSECorner Step 20
            If FromCell = temp Then
                Call MapDrawOnRoom(lblRoomCell(FromCell), drLineW, 4, Grey)
                GoTo DontActivate
            End If
        Next

    Case 4: 'northeast
        MapActivateCell = (FromCell - 19)
        'checking to see if it's on the north edge
        If MapActivateCell < 1 Then
            Call MapDrawOnRoom(lblRoomCell(FromCell), drLineNE, 4, Grey)
            GoTo DontActivate
        End If
        'checking to see if it's on the east edge
        For temp = nMapRowLength To sMapSECorner Step 20
            If FromCell = temp Then
                Call MapDrawOnRoom(lblRoomCell(FromCell), drLineNE, 4, Grey)
                GoTo DontActivate
            End If
        Next

    Case 5: 'northwest
        MapActivateCell = (FromCell - 21)
        'checking to see if it's on the north edge
        If MapActivateCell < 1 Then
            Call MapDrawOnRoom(lblRoomCell(FromCell), drLineNW, 4, Grey)
            GoTo DontActivate:
        End If
        'checking to see if it's on the west edge
        For temp = 1 To sMapSECorner Step 20
            If FromCell = temp Then
                Call MapDrawOnRoom(lblRoomCell(FromCell), drLineNW, 4, Grey)
                GoTo DontActivate
            End If
        Next

    Case 6: 'southeast
        MapActivateCell = (FromCell + 21)
        'checking to see if it's on the south edge
        If MapActivateCell > sMapSECorner Then
            Call MapDrawOnRoom(lblRoomCell(FromCell), drLineSE, 4, Grey)
            GoTo DontActivate
        End If
        'checking to see if it's on the east edge
        For temp = nMapRowLength To sMapSECorner Step 20
            If FromCell = temp Then
                Call MapDrawOnRoom(lblRoomCell(FromCell), drLineSE, 4, Grey)
                GoTo DontActivate
            End If
        Next

    Case 7: 'southwest
        MapActivateCell = (FromCell + 19)
        'checking to see if it's on the south edge
        If MapActivateCell > sMapSECorner Then
            Call MapDrawOnRoom(lblRoomCell(FromCell), drLineSW, 4, Grey)
            GoTo DontActivate:
        End If
        'checking to see if it's on the west edge
        For temp = 1 To sMapSECorner Step 20
            If FromCell = temp Then
                Call MapDrawOnRoom(lblRoomCell(FromCell), drLineSW, 4, Grey)
                GoTo DontActivate
            End If
        Next

    Case 8:
        GoTo DontActivate:

    Case 9:
        GoTo DontActivate:
    
    Case Else:
        GoTo DontActivate:
        
End Select

If MapActivateCell < 1 Or MapActivateCell > sMapSECorner Then GoTo DontActivate:

'set line mode
'ScaleMode = vbPixels
DrawWidth = 4

'pick line color
Select Case ExitType
    Case 2: LineColor = 10    'l green - key
    Case 3: LineColor = 10    'l green - item
    Case 4: LineColor = 10    'l green - toll
    Case 5: LineColor = 11    'l cyan - action
    Case 6: LineColor = 5     'd magenta - hidden
    Case 7: LineColor = 9     'l blue - door/gate
    Case 8: LineColor = 13    'l magenta - map change
    Case 9: LineColor = 12    'l red - trap/spell trap
    Case 10: LineColor = 14   'l yellow - text
    Case 11: LineColor = 9    'l blue - door/gate
    Case 12: LineColor = 11   'l cyan - remote action
    Case 13: LineColor = 4    'd red - class
    Case 14: LineColor = 4    'd red - race
    Case 15: LineColor = 4    'd red - level
    Case 16: LineColor = 2    'gray - timed
    Case 20: LineColor = 4    'd red - alignment
    Case 23: LineColor = 4    'd red - ability
    Case 24: LineColor = 12   'l red - trap/spell trap
    Case Else: LineColor = 8 '0  'black - anything else
End Select
    
'If chkNoColors.value = 1 Then LineColor = 0
'If chkNoLineColors.value = 1 Then LineColor = 0

'draw the line
Select Case direction
    Case 0: Call MapDrawOnRoom(lblRoomCell(FromCell), drLineN, 4, LineColor)
    Case 1: Call MapDrawOnRoom(lblRoomCell(FromCell), drLineS, 4, LineColor)
    Case 2: Call MapDrawOnRoom(lblRoomCell(FromCell), drLineE, 4, LineColor)
    Case 3: Call MapDrawOnRoom(lblRoomCell(FromCell), drLineW, 4, LineColor)
    Case 4: Call MapDrawOnRoom(lblRoomCell(FromCell), drLineNE, 4, LineColor)
    Case 5: Call MapDrawOnRoom(lblRoomCell(FromCell), drLineNW, 4, LineColor)
    Case 6: Call MapDrawOnRoom(lblRoomCell(FromCell), drLineSE, 4, LineColor)
    Case 7: Call MapDrawOnRoom(lblRoomCell(FromCell), drLineSW, 4, LineColor)
End Select

'if the cell to be activated has already been mapped, dont map it again
If UnchartedCells(MapActivateCell) = 2 Then GoTo DontActivate:

Select Case ExitType
    Case 12: MapActivateCell = -1 'if it's a remote action, dont map it
    Case 8: 'if it's a map change, check to see if it should be mapped
        If chkMapOptions(0).Value = 1 Then
            lblRoomCell(MapActivateCell).BackColor = &H0
        Else
            MapActivateCell = -1
        End If
    Case Else: lblRoomCell(MapActivateCell).BackColor = &H0
End Select

Exit Function
DontActivate:
MapActivateCell = -1

Exit Function

Error:
Call HandleError("MapActivateCell")

End Function

Private Sub MapDrawOnRoom(ByRef oLabel As Label, ByVal drDrawType As EnumDrawRoom, ByVal nSize As Integer, ByVal nColor As QBColorCode)
Dim x1 As Integer, x2 As Integer, y1 As Integer, y2 As Integer
Dim nTemp As Integer

nTemp = picMap.DrawWidth

'If chkNoColors.value = 1 Then nColor = Black

Select Case drDrawType
    Case 0: 'square
        picMap.DrawWidth = nSize
        x1 = oLabel.Left
        y1 = oLabel.Top
        x2 = oLabel.Left + oLabel.Width
        y2 = oLabel.Top + oLabel.Height
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor), BF
        
    Case 1: 'star
        picMap.DrawWidth = nSize
        '/
        x1 = oLabel.Left - 4
        y1 = oLabel.Top + oLabel.Height + 4
        x2 = oLabel.Left + 4
        y2 = oLabel.Top - 4
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
        
        '\
        x1 = x2
        y1 = y2
        x2 = oLabel.Left + oLabel.Width + 4
        y2 = oLabel.Top + oLabel.Height + 4
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
        
        '\
        x1 = x2
        y1 = y2
        x2 = oLabel.Left - 4
        y2 = oLabel.Top
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
        
        '-
        x1 = x2
        y1 = y2
        x2 = oLabel.Left + oLabel.Width + 4
        y2 = y1
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
        
        '/
        x1 = x2
        y1 = y2
        x2 = oLabel.Left - 4
        y2 = oLabel.Top + oLabel.Height + 4
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
        
    Case 2: 'open circle
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        picMap.Circle (x1, y1), 8, QBColor(nColor)
      
     Case 3: 'up
        picMap.DrawWidth = nSize
        x1 = oLabel.Left
        y1 = oLabel.Top
        x2 = oLabel.Left + oLabel.Width
        y2 = oLabel.Top + 2
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor), B
        
     Case 4: 'down
        picMap.DrawWidth = nSize
        x1 = oLabel.Left - 1
        y1 = oLabel.Top + oLabel.Height - 1
        x2 = oLabel.Left + oLabel.Width
        y2 = y1 + 2
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor), B
    
    Case 5: 'circle
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        picMap.Circle (x1, y1), 5, QBColor(nColor)
    
    Case 6: 'LineN
        'If chkNoLineColors.value = 1 Then nColor = Black
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        x2 = x1
        y2 = y1 - 8
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor), BF
        
    Case 7: 'LineS
        'If chkNoLineColors.value = 1 Then nColor = Black
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        x2 = x1
        y2 = y1 + 9
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor), BF
        
    Case 8: 'LineE
        'If chkNoLineColors.value = 1 Then nColor = Black
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        x2 = x1 + 9
        y2 = y1
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor), BF
        
    Case 9: 'LineW
        'If chkNoLineColors.value = 1 Then nColor = Black
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        x2 = x1 - 8
        y2 = y1
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor), BF
        
    Case 10: 'LineNE
        'If chkNoLineColors.value = 1 Then nColor = Black
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        x2 = x1 + 8
        y2 = y1 - 8
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
        
    Case 11: 'LineNW
        'If chkNoLineColors.value = 1 Then nColor = Black
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 5
        y1 = oLabel.Top + 5
        x2 = x1 - 8
        y2 = y1 - 8
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
        
    Case 12: 'LineSE
        'If chkNoLineColors.value = 1 Then nColor = Black
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 5
        y1 = oLabel.Top + 5
        x2 = x1 + 8
        y2 = y1 + 8
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
    
    Case 13: 'LineSW
        'If chkNoLineColors.value = 1 Then nColor = Black
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        x2 = x1 - 8
        y2 = y1 + 8
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
        
End Select

picMap.DrawWidth = nTemp
End Sub

Private Sub MapGetRoomLoc(ByVal nMapNumber As Long, ByVal nRoomNumber As Long)
On Error GoTo Error:
Dim x As Long, sLook As String, nExitType As Integer, RoomExit As RoomExitType
Dim oLI As ListItem, RoomExit2 As RoomExitType, sArray() As String
Dim nRecNum As Long, y As Long, sNumbers As String, sData As String, sCommand As String
Dim nMap As Long, nRoom As Long, sChar As String, nDataPos As Long, sLine As String

'=============================================================================
'
'                 NOTE: THIS ROUTINE IS ON BOTH frmMain AND frmMap
'
'=============================================================================

tabRooms.Index = "idxRooms"
tabRooms.Seek "=", nMapNumber, nRoomNumber
If tabRooms.NoMatch Then
    MsgBox "Room (" & nMapNumber & "/" & nRoomNumber & ") was not found."
    Exit Sub
End If

lvMapLoc.ColumnHeaders(1).Text = "References [" & tabRooms.Fields("Name") & " (" & nMapNumber & "/" & nRoomNumber & ")]"
nDataPos = 1

If tabRooms.Fields("CMD") > 0 Then 'chkMapOptions(4).Value = 0 And
    tabTBInfo.Index = "pkTBInfo"
    tabTBInfo.Seek "=", tabRooms.Fields("CMD")
    If tabTBInfo.NoMatch = False Then
        sData = tabTBInfo.Fields("Action")
        
        Do While nDataPos < Len(sData)
            x = InStr(nDataPos, sData, Chr(10))
            If x = 0 Then x = Len(sData)
            sLine = Mid(sData, nDataPos, x - nDataPos)
            nDataPos = x + 1
            
            x = InStr(1, sLine, "teleport ")
            If x > 0 Then
                y = x + Len("teleport ")
                x = y
                
                Do While y <= Len(sLine)
                    sChar = Mid(sLine, y, 1)
                    Select Case sChar
                        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
                        Case " ":
                            If y > x And nRoom = 0 Then
                                nRoom = Val(Mid(sLine, x, y - x))
                                x = y + 1
                            Else
                                nMap = Val(Mid(sLine, x, y - x))
                                Exit Do
                            End If
                        Case Else:
                            If y > x And nRoom = 0 Then
                                nRoom = Val(Mid(sLine, x, y - x))
                                Exit Do
                            Else
                                nMap = Val(Mid(sLine, x, y - x))
                                Exit Do
                            End If
                            Exit Do
                    End Select
                    y = y + 1
                Loop
                
                If Not nRoom = 0 Then
                    sCommand = Left(sLine, InStr(1, sLine, ":") - 1)
                    If nMap = 0 Then nMap = nMapNumber
                    
                    For Each oLI In lvMapLoc.ListItems
                        If oLI.Tag = nMap & "/" & nRoom Then GoTo skiptele:
                    Next
                    
                    Set oLI = lvMapLoc.ListItems.Add()
                    oLI.Text = "Teleport: " & sCommand _
                        & " --> " & GetRoomName(, nMap, nRoom, False)
                    oLI.Tag = nMap & "/" & nRoom
                End If
skiptele:
                nRoom = 0
                nMap = 0
            End If
            tabRooms.Seek "=", nMapNumber, nRoomNumber
        Loop
        
        Set oLI = lvMapLoc.ListItems.Add()
        oLI.Text = "Commands: Textblock " & tabRooms.Fields("CMD")
        oLI.Tag = tabRooms.Fields("CMD")
    End If
End If

If chkMapOptions(3).Value = 0 And tabRooms.Fields("NPC") > 0 Then
    Set oLI = lvMapLoc.ListItems.Add()
    oLI.Text = "NPC: " & GetMonsterName(tabRooms.Fields("NPC"), bHideRecordNumbers)
    oLI.Tag = tabRooms.Fields("NPC")
End If

If tabRooms.Fields("Shop") > 0 Then
    Set oLI = lvMapLoc.ListItems.Add()
    oLI.Text = "Shop: " & GetShopName(tabRooms.Fields("Shop"), bHideRecordNumbers) '& "(" & tabRooms.Fields("Shop") & ")"
    oLI.Tag = tabRooms.Fields("Shop")
End If

If tabRooms.Fields("Spell") > 0 Then
    Set oLI = lvMapLoc.ListItems.Add()
    oLI.Text = "Spell: " & GetSpellName(tabRooms.Fields("Spell"), bHideRecordNumbers)
    oLI.Tag = tabRooms.Fields("Spell")
End If

For x = 0 To 9
    Select Case x
        Case 0: sLook = "N"
        Case 1: sLook = "S"
        Case 2: sLook = "E"
        Case 3: sLook = "W"
        Case 4: sLook = "NE"
        Case 5: sLook = "NW"
        Case 6: sLook = "SE"
        Case 7: sLook = "SW"
        Case 8: sLook = "U"
        Case 9: sLook = "D"
    End Select
    
    nExitType = 0
    If Not Val(tabRooms.Fields(sLook)) = 0 Then
        RoomExit = ExtractMapRoom(tabRooms.Fields(sLook))
        
        If Len(RoomExit.ExitType) > 2 Then
            Select Case Left(RoomExit.ExitType, 5)
                Case "(Key:": nExitType = 2
                Case "(Item": nExitType = 3
                Case "(Toll": nExitType = 4
                Case "(Hidd": nExitType = 6
                Case "(Door": nExitType = 7
                Case "(Trap": nExitType = 9
                Case "(Text": nExitType = 10
                Case "(Gate": nExitType = 11
                Case "Actio": nExitType = 12
                Case "(Clas": nExitType = 13
                Case "(Race": nExitType = 14
                Case "(Leve": nExitType = 15
                Case "(Time": nExitType = 16
                Case "(Tick": nExitType = 17
                Case "(Max ": nExitType = 18
                Case "(Bloc": nExitType = 19
                Case "(Alig": nExitType = 20
                Case "(Dela": nExitType = 21
                Case "(Cast": nExitType = 22
                Case "(Abil": nExitType = 23
                Case "(Spel": nExitType = 24
            End Select
        End If
        
        Select Case nExitType
            Case 0:
            Case 2, 3, 17:
                nRecNum = ExtractNumbersFromString(RoomExit.ExitType)
                If nRecNum > 0 Then
                    Set oLI = lvMapLoc.ListItems.Add()
                    oLI.Text = "Item: " & GetItemName(nRecNum, bHideRecordNumbers) '& " (" & nRecNum & ")"
                    oLI.Tag = nRecNum
                End If
            Case 22, 24:
'                nRecNum = ExtractNumbersFromString(RoomExit.ExitType)
'                If nRecNum > 0 Then
'                    Set oLI = lvMapLoc.ListItems.Add()
'                    oLI.Text = "Spell: " & GetSpellName(nRecNum, bHideRecordNumbers) '& " (" & nRecNum & ")"
'                    oLI.Tag = nRecNum
'                End If
                nRecNum = ExtractValueFromString(RoomExit.ExitType, "pre-") ' ExtractNumbersFromString(RoomExit.ExitType)
                If nRecNum > 0 Then
                    Set oLI = lvMapLoc.ListItems.Add()
                    oLI.Text = "Spell: " & GetSpellName(nRecNum, bHideRecordNumbers) '& " (" & nRecNum & ")"
                    oLI.Tag = nRecNum
                End If
                nRecNum = ExtractValueFromString(RoomExit.ExitType, "post-") ' ExtractNumbersFromString(RoomExit.ExitType)
                If nRecNum > 0 Then
                    Set oLI = lvMapLoc.ListItems.Add()
                    oLI.Text = "Spell: " & GetSpellName(nRecNum, bHideRecordNumbers) '& " (" & nRecNum & ")"
                    oLI.Tag = nRecNum
                End If
            Case 12:
                RoomExit2 = ExtractMapRoom(RoomExit.ExitType)
                If RoomExit2.Map > 0 Then
                    sChar = "Action On: " & GetRoomName(, RoomExit2.Map, RoomExit2.Room, False) '& " (" & RoomExit2.Map & "/" & RoomExit2.Room & ")"
                    For Each oLI In lvMapLoc.ListItems
                        If oLI.Text = sChar Then GoTo nextexit:
                    Next
                    Set oLI = lvMapLoc.ListItems.Add()
                    oLI.Text = sChar
                    oLI.Tag = RoomExit2.Map & "/" & RoomExit2.Room
                    
                    tabRooms.Seek "=", nMapNumber, nRoomNumber
                End If
                If InStr(1, tabRooms.Fields(sLook), "(Item: ") > 0 Then
                    nRecNum = ExtractValueFromString(tabRooms.Fields(sLook), "(Item: ")
                    If nRecNum > 0 Then
                        Set oLI = lvMapLoc.ListItems.Add()
                        oLI.Text = "Item: " & GetItemName(nRecNum, bHideRecordNumbers) '& " (" & nRecNum & ")"
                        oLI.Tag = nRecNum
                    End If
                End If
        End Select
    ElseIf Left(tabRooms.Fields(sLook), 6) = "Action" Then
        RoomExit2 = ExtractMapRoom(tabRooms.Fields(sLook))
        If RoomExit2.Map > 0 Then
            sChar = "Action On: " & GetRoomName(, RoomExit2.Map, RoomExit2.Room, False) '& " (" & RoomExit2.Map & "/" & RoomExit2.Room & ")"
            For Each oLI In lvMapLoc.ListItems
                If oLI.Text = sChar Then GoTo nextexit:
            Next
            Set oLI = lvMapLoc.ListItems.Add()
            oLI.Text = sChar
            oLI.Tag = RoomExit2.Map & "/" & RoomExit2.Room

            tabRooms.Seek "=", nMapNumber, nRoomNumber
        End If
        If InStr(1, tabRooms.Fields(sLook), "(Item: ") > 0 Then
            nRecNum = ExtractValueFromString(tabRooms.Fields(sLook), "(Item: ")
            If nRecNum > 0 Then
                Set oLI = lvMapLoc.ListItems.Add()
                oLI.Text = "Item: " & GetItemName(nRecNum, bHideRecordNumbers) '& " (" & nRecNum & ")"
                oLI.Tag = nRecNum
            End If
        End If
    End If
nextexit:
Next x

If chkMapOptions(2).Value = 0 And Len(tabRooms.Fields("Lair")) > 1 Then
    tabMonsters.Index = "pkMonsters"
    sNumbers = Mid(tabRooms.Fields("Lair"), InStr(1, tabRooms.Fields("Lair"), ":") + 2)
    x = 0
    Do While Not InStr(x + 1, sNumbers, ",") = 0
        y = InStr(x + 1, sNumbers, ",")
        
        tabMonsters.Seek "=", Val(Mid(sNumbers, x + 1, y - x - 1))
        If tabMonsters.NoMatch = False Then
            Set oLI = lvMapLoc.ListItems.Add()
            oLI.Text = "Lair: " & tabMonsters.Fields("Name") & IIf(bHideRecordNumbers, "", "(" & tabMonsters.Fields("Number") & ")")
            oLI.Tag = tabMonsters.Fields("Number")
        End If
        x = y
    Loop
End If

If Len(tabRooms.Fields("Placed")) > 1 Then
    sArray() = Split(tabRooms.Fields("Placed"), ",")
    If UBound(sArray()) >= 0 Then
        For x = 0 To UBound(sArray())
            If Val(sArray(x)) > 0 Then
                tabItems.Index = "pkItems"
                tabItems.Seek "=", Val(sArray(0))
                If tabItems.NoMatch = False Then
                    Set oLI = lvMapLoc.ListItems.Add()
                    oLI.Text = "Item: " & tabItems.Fields("Name") & IIf(bHideRecordNumbers, "", "(" & tabItems.Fields("Number") & ")")
                    oLI.Tag = tabItems.Fields("Number")
                End If
            End If
        Next x
    End If
    Erase sArray()
End If

'If lvMapLoc.ListItems.Count > 0 Then
'    Call SortListView(lvMapLoc, 1, ldtstring, True)
'End If

Set oLI = Nothing
Exit Sub
Error:
Call HandleError("MapGetRoomLoc")
Set oLI = Nothing
End Sub

Private Sub MapGoDirection(ByVal nSourceMapNumber As Long, ByVal nSourceRoomNumber As Long, ByVal sDirection As String)
On Error GoTo Error:
Dim RoomExits As RoomExitType

tabRooms.Index = "idxRooms"
tabRooms.Seek "=", nSourceMapNumber, nSourceRoomNumber
If tabRooms.NoMatch Then
    MsgBox "Source room (" & nSourceMapNumber & "/" & nSourceRoomNumber & ") not found."
    Exit Sub
End If

RoomExits = ExtractMapRoom(tabRooms.Fields(sDirection))
If Not RoomExits.Map = 0 And Not RoomExits.Room = 0 Then
    Call MapStartMapping(RoomExits.Map, RoomExits.Room)
End If
Exit Sub
Error:
Call HandleError("MapGoDirection")
End Sub

Private Sub MapMapExits(Cell As Integer, Room As Long, Map As Long)
Dim ActivatedCell As Integer, x As Integer, y As Long
Dim rc As RECT, ToolTipString As String, sText As String, sPlaced As String
Dim sRemote As String, sMonsters As String, sArray() As String
Dim RoomExit As RoomExitType, sLook As String, nExitType As Integer, sRoomCMDs As String

On Error GoTo Error:

'=============================================================================
'
'                 NOTE: THIS ROUTINE IS ON BOTH frmMain AND frmMap
'
'=============================================================================

CellRoom(Cell, 1) = Map
CellRoom(Cell, 2) = Room

tabRooms.Index = "idxRooms"
tabRooms.Seek "=", Map, Room
If tabRooms.NoMatch Then
    UnchartedCells(Cell) = 2
    Call MapDrawOnRoom(lblRoomCell(Cell), drSquare, 8, BrightRed)
    ToolTipString = "Map " & Map & " Room " & Room
    rc.Left = lblRoomCell(Cell).Left
    rc.Top = lblRoomCell(Cell).Top
    rc.Bottom = (lblRoomCell(Cell).Top + lblRoomCell(Cell).Height)
    rc.Right = (lblRoomCell(Cell).Left + lblRoomCell(Cell).Width)
    objToolTip.SetToolTipItem picMap.hwnd, 0, rc.Left, rc.Top, rc.Right, rc.Bottom, ToolTipString, False
    Exit Sub
End If

ToolTipString = Map & "/" & Room & " - " & tabRooms.Fields("Name")

If chkMapOptions(4).Value = 0 And tabRooms.Fields("CMD") > 0 Then
    sRoomCMDs = vbCrLf & vbCrLf & "Room commands: " & GetTextblockCMDS(tabRooms.Fields("CMD"))
    Call MapDrawOnRoom(lblRoomCell(Cell), drSquare, 6, BrightGreen)
Else
    sRoomCMDs = ""
End If

If chkMapOptions(3).Value = 0 And tabRooms.Fields("NPC") > 0 Then
    ToolTipString = ToolTipString & vbCrLf & "NPC: " & GetMonsterName(tabRooms.Fields("NPC"), bHideRecordNumbers)
    Call MapDrawOnRoom(lblRoomCell(Cell), drOpenCircle, 2, BrightRed)
End If

If Len(tabRooms.Fields("Placed")) > 1 Then
    sArray() = Split(tabRooms.Fields("Placed"), ",")
    If UBound(sArray()) >= 0 Then
        For x = 0 To UBound(sArray())
            If Val(sArray(x)) > 0 Then
                If Not sPlaced = "" Then sPlaced = sPlaced & ", "
                sPlaced = sPlaced & GetItemName(Val(sArray(0)), bHideRecordNumbers)
            End If
        Next x
        ToolTipString = ToolTipString & vbCrLf & "Placed Items: " & sPlaced
        'Call MapDrawOnRoom(lblRoomCell(Cell), drOpenCircle, 2, BrightRed)
    End If
    Erase sArray()
End If

If chkMapOptions(2).Value = 0 And Len(tabRooms.Fields("Lair")) > 1 Then
    sMonsters = GetMultiMonsterNames(Mid(tabRooms.Fields("Lair"), InStr(1, tabRooms.Fields("Lair"), ":") + 2), bHideRecordNumbers)
    sMonsters = "Also Here " & Left(tabRooms.Fields("Lair"), InStr(1, tabRooms.Fields("Lair"), ":") + 1) & sMonsters
    Call MapDrawOnRoom(lblRoomCell(Cell), drCircle, 5, BrightMagenta)
End If

If tabRooms.Fields("Shop") > 2 Then
    ToolTipString = ToolTipString & vbCrLf & "Shop: " & GetShopName(tabRooms.Fields("Shop"), bHideRecordNumbers) '& "(" & tabRooms.Fields("Shop") & ")"
End If

If tabRooms.Fields("Spell") > 0 Then
    ToolTipString = ToolTipString & vbCrLf & "Room Spell: " & GetSpellName(tabRooms.Fields("Spell"), bHideRecordNumbers)
End If

'map exits
For x = 0 To 9
    Select Case x
        Case 0: sLook = "N"
        Case 1: sLook = "S"
        Case 2: sLook = "E"
        Case 3: sLook = "W"
        Case 4: sLook = "NE"
        Case 5: sLook = "NW"
        Case 6: sLook = "SE"
        Case 7: sLook = "SW"
        Case 8: sLook = "U"
        Case 9: sLook = "D"
    End Select
    
    nExitType = 0
    If Left(tabRooms.Fields(sLook), 6) = "Action" Then
        sRemote = sRemote & vbCrLf & tabRooms.Fields(sLook)
        If chkMapOptions(4).Value = 0 Then Call MapDrawOnRoom(lblRoomCell(Cell), drSquare, 6, BrightGreen)
    
    ElseIf Not Val(tabRooms.Fields(sLook)) = 0 Then
        RoomExit = ExtractMapRoom(tabRooms.Fields(sLook))
        
        If Len(RoomExit.ExitType) > 2 Then
            Select Case Left(RoomExit.ExitType, 5)
                Case "(Key:": nExitType = 2
                Case "(Item": nExitType = 3
                Case "(Toll": nExitType = 4
                Case "(Hidd": nExitType = 6
                Case "(Door": nExitType = 7
                Case "(Trap": nExitType = 9
                Case "(Text": nExitType = 10
                Case "(Gate": nExitType = 11
                Case "Actio": nExitType = 12
                Case "(Clas": nExitType = 13
                Case "(Race": nExitType = 14
                Case "(Leve": nExitType = 15
                Case "(Time": nExitType = 16
                Case "(Tick": nExitType = 17
                Case "(Max ": nExitType = 18
                Case "(Bloc": nExitType = 19
                Case "(Alig": nExitType = 20
                Case "(Dela": nExitType = 21
                Case "(Cast": nExitType = 22
                Case "(Abil": nExitType = 23
                Case "(Spel": nExitType = 24
            End Select
        End If
        If Not RoomExit.Map = Map Then nExitType = 8 'map change
        
        'sText = sText & vbCrLf & sLook & ": " & RoomExit.Map & "/" & RoomExit.Room

        'note order of case'ings is important here
        Select Case nExitType
            Case 2: 'key
                y = ExtractValueFromString(RoomExit.ExitType, "Key: ")
                sText = sText & vbCrLf & sLook & " (Key: " _
                    & GetItemName(y, bHideRecordNumbers) _
                    & " " & Mid(RoomExit.ExitType, InStr(1, RoomExit.ExitType, y) + Len(CStr(y)) + 1)

                ActivatedCell = MapActivateCell(Cell, x, nExitType)
                If ActivatedCell = -1 Then GoTo skip:

                If chkMapOptions(1).Value = 1 And nExitType = 6 Then GoTo skip:

                CellRoom(ActivatedCell, 1) = Map
                CellRoom(ActivatedCell, 2) = RoomExit.Room
                If UnchartedCells(ActivatedCell) = 0 Then UnchartedCells(ActivatedCell) = 1

            Case 3: 'item
                y = ExtractValueFromString(RoomExit.ExitType, "Item: ")
                sText = sText & vbCrLf & sLook & " (Item): " _
                    & GetItemName(y, bHideRecordNumbers) _
                    & " " & Mid(RoomExit.ExitType, InStr(1, RoomExit.ExitType, y) + Len(CStr(y)) + 1)

                ActivatedCell = MapActivateCell(Cell, x, nExitType)
                If ActivatedCell = -1 Then GoTo skip:

                If chkMapOptions(1).Value = 1 And nExitType = 6 Then GoTo skip:

                CellRoom(ActivatedCell, 1) = Map
                CellRoom(ActivatedCell, 2) = RoomExit.Room
                If UnchartedCells(ActivatedCell) = 0 Then UnchartedCells(ActivatedCell) = 1
                
            Case 8: 'map change
                ActivatedCell = MapActivateCell(Cell, x, nExitType)
                If ActivatedCell = -1 Then GoTo skip:
                If chkMapOptions(0).Value = 1 Then
                    CellRoom(ActivatedCell, 1) = RoomExit.Map
                    CellRoom(ActivatedCell, 2) = RoomExit.Room
                    If UnchartedCells(ActivatedCell) = 0 Then UnchartedCells(ActivatedCell) = 1
                End If
            Case 12: 'action
                sRemote = sRemote & vbCrLf & tabRooms.Fields(sLook)
                If chkMapOptions(4).Value = 0 Then Call MapDrawOnRoom(lblRoomCell(Cell), drSquare, 6, BrightGreen)
            Case Is > 0:
                sText = sText & vbCrLf & sLook & ": " & RoomExit.ExitType
                ActivatedCell = MapActivateCell(Cell, x, nExitType)
                If ActivatedCell = -1 Then GoTo skip:
                
                If chkMapOptions(1).Value = 1 And nExitType = 6 Then GoTo skip:
                
                CellRoom(ActivatedCell, 1) = Map
                CellRoom(ActivatedCell, 2) = RoomExit.Room
                If UnchartedCells(ActivatedCell) = 0 Then UnchartedCells(ActivatedCell) = 1

            Case Else:
                ActivatedCell = MapActivateCell(Cell, x, nExitType) 'nExitType)
                If ActivatedCell = -1 Then GoTo skip:
                CellRoom(ActivatedCell, 1) = Map
                CellRoom(ActivatedCell, 2) = RoomExit.Room
                If UnchartedCells(ActivatedCell) = 0 Then UnchartedCells(ActivatedCell) = 1
        End Select
    End If
skip:
Next x

'set color of this room
If Val(tabRooms.Fields("U")) = 0 And Val(tabRooms.Fields("D")) = 0 Then
    lblRoomCell(Cell).BackColor = &HC0C0C0   '&H0& '-- nothing
ElseIf Val(tabRooms.Fields("U")) > 0 And Val(tabRooms.Fields("D")) = 0 Then
    lblRoomCell(Cell).BackColor = &HFF00& '-- up
ElseIf Val(tabRooms.Fields("U")) = 0 And Val(tabRooms.Fields("D")) > 0 Then
    lblRoomCell(Cell).BackColor = &HFFFF& '-- down
Else
    lblRoomCell(Cell).BackColor = &HFFFF00 '-- both
End If

If chkMapOptions(5).Value = 0 Then
    ToolTipString = ToolTipString & sText & IIf(sRemote = "", "", vbCrLf & sRemote) & sRoomCMDs _
        & IIf(sMonsters = "", "", vbCrLf & vbCrLf & sMonsters)
    
    rc.Left = lblRoomCell(Cell).Left
    rc.Top = lblRoomCell(Cell).Top
    rc.Bottom = (lblRoomCell(Cell).Top + lblRoomCell(Cell).Height)
    rc.Right = (lblRoomCell(Cell).Left + lblRoomCell(Cell).Width)
    objToolTip.SetToolTipItem picMap.hwnd, 0, rc.Left, rc.Top, rc.Right, rc.Bottom, ToolTipString, False
End If

UnchartedCells(Cell) = 2

Exit Sub

Error:
Call HandleError("MapMapExits")
End Sub

Public Sub MapStartMapping(ByVal nStartMap As Long, ByVal nStartRoom As Long, Optional nCenterCell As Integer)
On Error GoTo Error:
Dim x As Integer, nMapSize As Integer, bCheckAgain As Boolean, y As Integer

If bMapStillMapping Then Exit Sub

tabRooms.Index = "idxRooms"
tabRooms.Seek "=", nStartMap, nStartRoom
If tabRooms.NoMatch Then
    MsgBox "Room " & nStartMap & "/" & nStartRoom & " was not found.", vbInformation
    'framNav(10).Caption = "Rooms"
    Exit Sub
Else
    framNav(10).Caption = "Rooms -- " & tabRooms.Fields("Name") & " (" & nStartMap & "/" & nStartRoom & ")  "
End If

If Not nMapStartRoom = nStartRoom Then
    nMapLastRoom = nMapStartRoom
    nMapLastMap = nMapStartMap
End If

bMapStillMapping = True
Call LockWindowUpdate(Me.hwnd)

'picMap.Visible = False
picMap.Cls
Me.MousePointer = vbHourglass
DoEvents
'20x20
sMapSECorner = 400
nMapRowLength = 20
If Not nCenterCell = 0 Then nMapCenterCell = nCenterCell
If nMapCenterCell = 0 Then nMapCenterCell = 210
If nMapCenterCell > sMapSECorner Then nMapCenterCell = 210

For x = 1 To 400
    objToolTip.DelToolTip picMap.hwnd, 0
    lblRoomCell(x).BackColor = &HFFFFFF
    lblRoomCell(x).Visible = False
    lblRoomCell(x).Tag = 0
    UnchartedCells(x) = 0
    CellRoom(x, 1) = 0
    CellRoom(x, 2) = 0
Next x

StopBuild = False

nMapStartRoom = nStartRoom
nMapStartMap = nStartMap

CellRoom(nMapCenterCell, 1) = nMapStartMap
CellRoom(nMapCenterCell, 2) = nMapStartRoom

Call MapMapExits(nMapCenterCell, nMapStartRoom, nMapStartMap)

DoEvents
again:
bCheckAgain = False
For x = 1 To sMapSECorner
    If StopBuild = True Then GoTo Cancel:
    If UnchartedCells(x) = 1 Then
        For y = 1 To sMapSECorner
            If Not CellRoom(x, 1) = 0 Then
                If Not x = y Then
                    If CellRoom(y, 2) = CellRoom(x, 2) Then
                        If CellRoom(y, 1) = CellRoom(x, 1) Then
                            CellRoom(x, 2) = 0
                            CellRoom(x, 1) = 0
                            UnchartedCells(x) = 0
                            GoTo skiproom:
                        End If
                    End If
                End If
            End If
        Next y
        Call MapMapExits(x, CellRoom(x, 2), CellRoom(x, 1))
        bCheckAgain = True
    End If
skiproom:
    'DoEvents
Next x
If bCheckAgain Then GoTo again:

Call MapDrawOnRoom(lblRoomCell(nMapCenterCell), drSquare, 4, BrightBlue)

DoEvents
cmdMapShowUnused.Caption = "S&how Blocks"
For x = 1 To 400
    If Not CellRoom(x, 1) = 0 Then lblRoomCell(x).Visible = True
Next x
DoEvents

Call lblRoomCell_MouseDown(nMapCenterCell, IIf(bMapSwapButtons, 2, 1), 0, 0, 0)
'picMap.Visible = True

Cancel:
On Error Resume Next
Me.MousePointer = vbDefault
bMapStillMapping = False
Call LockWindowUpdate(0&)

Exit Sub
Error:
Call HandleError("MapStartMapping")
Resume Cancel:
End Sub

'*************************** /spell compare

Private Sub mnuAuxPopUpItem_Click(Index As Integer)
On Error GoTo Error:
Dim oLI As ListItem, nResult As Integer, sClip As String, x As Long

Select Case Index
    Case 0: 'Copy
        Select Case objWorkingListView.name
            Case "lvClasses", "lvRaces", "lvShopDetail":
                Call CopyLVLinetoClipboard(objWorkingListView)
            Case "lvShops":
                Call CopyShopToClipboard
            Case "lvOtherItems":
                Call CopyLVLinetoClipboard(objWorkingListView, txtOtherItemDetail, lvOtherItemLoc)
            Case "lvMonsters":
                Call CopyMonsterToClipboard(lvMonsterDetail) ', txtMonsterCompareDetail)
            Case "lvMonsterCompare":
                Call CopyMonsterToClipboard(lvMonsterCompareLoc)
            Case "lvCalcExp":
                If frmExpCalc.Visible = False Then
                    If bNoAlwaysOnTop Then
                        frmExpCalc.Show vbModeless, Me
                    Else
                        frmExpCalc.Show vbModeless
                    End If
                End If
                Call CopyWholeLVtoClipboard(frmExpCalc.lvCalcExp, True)
        End Select
    Case 1: 'look up monster
        If Not objWorkingListView.SelectedItem Is Nothing Then
            Call LookUpMonsterRegen(Val(objWorkingListView.SelectedItem.Text), True)
        End If
    Case 2: 'add shop compare
        Call CompareAddShopItems
    Case 3: 'add/remove monster compare
        If mnuAuxPopUpItem(3).Tag = "2" Then 'remove
            If lvMonsterCompare.ListItems.Count = 0 Then Exit Sub
            x = 1
            Do While x <= lvMonsterCompare.ListItems.Count
                If lvMonsterCompare.ListItems(x).Selected Then
                    Call lvMonsterCompare.ListItems.Remove(x)
                Else
                    x = x + 1
                End If
            Loop
        Else 'add
            For Each oLI In lvMonsters.ListItems
                If oLI.Selected Then
                    nResult = CompareAddMonster(Val(oLI.Text))
                    If nResult = vbCancel Then GoTo quit:
                End If
            Next oLI
        End If
        
    Case 4: 'copy chest
        If Not lvOtherItems.SelectedItem Is Nothing Then
            Call cmdSundryChests_Click
            If lvOtherItemLoc.ListItems.Count > 0 Then
                sClip = "Chest info for " & lvOtherItems.SelectedItem.ListSubItems(1) _
                    & " (" & lvOtherItems.SelectedItem.Text & "):"
                For Each oLI In lvOtherItemLoc.ListItems
                    sClip = sClip & vbCrLf & oLI.Text
                Next oLI
                Clipboard.clear
                Clipboard.SetText sClip
            End If
        End If
End Select

quit:
Set oLI = Nothing
Exit Sub

Error:
Call HandleError("mnuAuxPopUpItem_Click")
Set oLI = Nothing
End Sub

Private Sub mnuEquipGotoItem_Click(Index As Integer)
'Dim nMinDMG As Long, nMaxDMG As Long, sTemp As String
'Dim x As Integer, nMaxDMGBonus As Integer
On Error GoTo Error:

'If cmbEquip(nLastEquipIndex).ListIndex < 1 Then Exit Sub
If nEquippedItem(nLastEquipIndex) < 1 Then Exit Sub

Select Case Index
    Case 0: 'goto item
        Call GotoItem(nEquippedItem(nLastEquipIndex))
    Case 1: 'add to compare
        Call CompareAddItem(nEquippedItem(nLastEquipIndex))
    Case 2: 'calc swings
        Call frmSwingCalc.GotoWeapon(nEquippedItem(nLastEquipIndex))
        If frmSwingCalc.Visible = False Then
            If bNoAlwaysOnTop Then
                frmSwingCalc.Show vbModeless
            Else
                frmSwingCalc.Show vbModeless, Me
            End If
        Else
            frmSwingCalc.SetFocus
        End If
    Case 3: 'calc bs damage
        tabItems.Seek "=", nEquippedItem(nLastEquipIndex)
        If Not tabItems.NoMatch Then
            frmBSCalc.GotoWeapon (nEquippedItem(nLastEquipIndex))
            frmBSCalc.Show
            frmBSCalc.SetFocus
        Else
            MsgBox "Item not found.", vbOKCancel
        End If
        
End Select

out:
Exit Sub
Error:
Call HandleError("mnuEquipGotoItem_Click")
Resume out:
End Sub

Private Sub mnuExit_Click()

Unload Me

End Sub

Private Sub mnuFileOptItem_Click(Index As Integer)
Dim sSectionName As String, nYesNo As Integer, nResult As Integer
sSectionName = RemoveCharacter(lblDatVer.Caption, " ")

Select Case Index
    Case 0: 'load
        If bCharLoaded Then
            If bAutoSave Then
                If FileExists(ReadINI(sSectionName, "LastCharFile")) Then
                    nResult = SaveCharacter(False, ReadINI(sSectionName, "LastCharFile"))
                    If nResult = -1 Then Exit Sub
                Else
                    If bPromptSave = True Then Call SaveCharacter(True)
                End If
            Else
                If bPromptSave = True Then
                    nYesNo = MsgBox("Save character file first?", vbYesNoCancel + vbQuestion + vbDefaultButton3)
                    If nYesNo = vbYes Then
                        If FileExists(ReadINI(sSectionName, "LastCharFile")) Then
                            nResult = SaveCharacter(False, ReadINI(sSectionName, "LastCharFile"))
                        Else
                            nResult = SaveCharacter(True)
                        End If
                        If nResult = -1 Then Exit Sub
                    ElseIf nYesNo = vbCancel Then
                        Exit Sub
                    End If
                End If
            End If
        End If
        Call LoadCharacter(True)
    Case 1: 'revert
        If bCharLoaded Then
            If FileExists(ReadINI(sSectionName, "LastCharFile")) Then
                Call LoadCharacter(False, ReadINI(sSectionName, "LastCharFile"), True)
            Else
                Call LoadCharacter(True)
            End If
        Else
            Call LoadCharacter(False, , True)
        End If
    Case 2: 'save
        If bCharLoaded Then
            If FileExists(ReadINI(sSectionName, "LastCharFile")) Then
                Call SaveCharacter(False, ReadINI(sSectionName, "LastCharFile"))
            Else
                Call SaveCharacter(True)
            End If
        Else
            Call SaveCharacter(True)
        End If
    Case 3: 'save as
        Call SaveCharacter(True)
    Case 4: 'close
        If bCharLoaded Then
            If bAutoSave Then
                If FileExists(ReadINI(sSectionName, "LastCharFile")) Then
                    nResult = SaveCharacter(False, ReadINI(sSectionName, "LastCharFile"))
                    If nResult = -1 Then Exit Sub
                Else
                    If bPromptSave = True Then Call SaveCharacter(True)
                End If
            Else
                If bPromptSave = True Then
                    nYesNo = MsgBox("Save character file first?", vbYesNoCancel + vbQuestion + vbDefaultButton3)
                    If nYesNo = vbYes Then
                        If FileExists(ReadINI(sSectionName, "LastCharFile")) Then
                            nResult = SaveCharacter(False, ReadINI(sSectionName, "LastCharFile"))
                        Else
                            nResult = SaveCharacter(True)
                        End If
                        If nResult = -1 Then Exit Sub
                    ElseIf nYesNo = vbCancel Then
                        Exit Sub
                    End If
                End If
            End If
            bCharLoaded = False
            Me.Caption = sNormalCaption
            Call LoadCharacter(False, , True)
        End If
End Select

End Sub

Private Sub mnuFindBestArmour_Click(Index As Integer)
Call InvenFindBest(Index, Armour)
End Sub

Private Sub mnuFindBestAttack_Click(Index As Integer)
Call InvenFindBest(Index, Attack)
End Sub

Private Sub mnuFindBestItem_Click(Index As Integer)
'Call InvenFindBest(Index)
End Sub

Private Sub mnuFindBestMystics_Click(Index As Integer)
Call InvenFindBest(Index, Mystics)
End Sub

Private Sub mnuFindBestResist_Click(Index As Integer)
Call InvenFindBest(Index, Resist)
End Sub

Private Sub mnuFindBestStats_Click(Index As Integer)
Call InvenFindBest(Index, Stats)
End Sub

Private Sub mnuHelp_Click(Index As Integer)
Dim sTemp As String

Select Case Index
    Case 0: frmAbout.Show vbModal, Me
    Case 1: frmHelpChangeLog.Show vbModal, Me
    Case 2:
        If tabInfo.RecordCount > 0 Then
            tabInfo.MoveFirst
            sTemp = tabInfo.Fields("UpdateURL")
            If Len(sTemp) < 5 Then sTemp = "http://www.mudinfo.net/mmudexp.php"
        Else
            sTemp = "http://www.mudinfo.net/mmudexp.php"
        End If
        Call ShellExecute(0&, "open", sTemp, vbNullString, vbNullString, vbNormalFocus)
End Select

End Sub

Private Sub mnuItemsPopUpItem_Click(Index As Integer)
Dim oLI As ListItem
Dim nMinDmg As Long, nMaxDmg As Long, sTemp As String
Dim x As Long, bClassStealth As Boolean, nMaxDMGBonus As Integer

On Error GoTo Error:

Select Case Index
    Case 0: 'add/remove compare
        bPromptSave = True
        If mnuItemsPopUpItem(0).Tag = "1" Then 'add
            For Each oLI In objWorkingListView.ListItems
                If oLI.Selected Then
                     Call CompareAddItem(Val(oLI.Text))
                End If
            Next oLI
        ElseIf mnuItemsPopUpItem(0).Tag = "2" Then 'remove
            If objWorkingListView.ListItems.Count = 0 Then Exit Sub
            
            x = 1
            Do While x <= objWorkingListView.ListItems.Count
                If objWorkingListView.ListItems(x).Selected Then
                    Call objWorkingListView.ListItems.Remove(x)
                Else
                    x = x + 1
                End If
            Loop
            
            If objWorkingListView.ListItems.Count > 0 Then
                Select Case objWorkingListView.name
                    Case "lvWeaponCompare":
                        Call ProcessListViewClick(objWorkingListView.ListItems(1), txtWeaponCompareDetail, lvWeaponCompareLoc)
                    Case "lvArmourCompare":
                        Call ProcessListViewClick(objWorkingListView.ListItems(1), txtArmourCompareDetail, lvArmourCompareLoc)
'                    Case "lvMonsterCompare":
'                        Set lvMonsterCompare.SelectedItem = lvMonsterCompare.ListItems(1)
'                        lvMonsterCompare.ListItems(1).Selected = True
'                        lvMonsterCompare.ListItems(1).EnsureVisible
'                        Call lvMonsterCompare_ItemClick(lvMonsterCompare.SelectedItem)
                End Select
            Else
                Select Case objWorkingListView.name
                    Case "lvWeaponCompare":
                        Call cmdCompareClear_Click(0)
                    Case "lvArmourCompare":
                        Call cmdCompareClear_Click(1)
'                    Case "lvMonsterCompare":
'                        Call cmdCompareClear_Click(3)
                End Select
            End If
        End If

    Case 1: 'calc swings
        Call frmSwingCalc.GotoWeapon(Val(objWorkingListView.SelectedItem.Text))
        If frmSwingCalc.Visible = False Then
            If bNoAlwaysOnTop Then
                frmSwingCalc.Show vbModeless
            Else
                frmSwingCalc.Show vbModeless, Me
            End If
        Else
            frmSwingCalc.SetFocus
        End If
        
    Case 2: 'copy to clip
        Select Case objWorkingListView.name
            Case "lvWeapons":
                Call CopyLVLinetoClipboard(objWorkingListView, txtWeaponDetail, lvWeaponLoc)
            Case "lvWeaponCompare":
                Call CopyLVLinetoClipboard(objWorkingListView, txtWeaponCompareDetail, lvWeaponCompareLoc)
            Case "lvArmour":
                Call CopyLVLinetoClipboard(objWorkingListView, txtArmourDetail, lvArmourLoc)
            Case "lvArmourCompare":
                Call CopyLVLinetoClipboard(objWorkingListView, txtArmourCompareDetail, lvArmourCompareLoc)
            Case "lvShopDetail":
                Call CopyLVLinetoClipboard(objWorkingListView)
            Case "lvMonsterCompare":
                Call CopyMonsterToClipboard(lvMonsterCompareLoc) ', txtMonsterCompareDetail)
        End Select
        
    Case 3: 'equip
        bPromptSave = True
        For Each oLI In objWorkingListView.ListItems
            bDontRefresh = True
            If oLI.Selected = True Then Call InvenEquipItem(Val(oLI.Text), mnuJumpToCompare.Checked)
            Set oLI = Nothing
        Next
        bDontRefresh = False
        Call RefreshAll
    Case 4: 'calc bs
        tabItems.Seek "=", Val(objWorkingListView.SelectedItem.Text)
        If Not tabItems.NoMatch Then
            Call frmBSCalc.GotoWeapon(Val(objWorkingListView.SelectedItem.Text))
            
            If Val(txtGlobalLevel(0).Text) > 0 Then
                frmBSCalc.txtLevel.Text = Val(txtGlobalLevel(0).Text)
            End If
            
            frmBSCalc.txtMaxDMG.Text = Val(txtStat(11).Text)
            
            bClassStealth = GetClassStealth(cmbGlobalClass(0).ItemData(cmbGlobalClass(0).ListIndex))
            If bClassStealth Then
                frmBSCalc.chkClassStealth.Value = 1
            Else
                frmBSCalc.chkClassStealth.Value = 0
            End If
            
            frmBSCalc.txtBSMinDMG.Text = Val(txtStat(14).Text)
            frmBSCalc.txtBSMaxDMG.Text = Val(txtStat(15).Text)
            
            frmBSCalc.Show
            frmBSCalc.SetFocus
        Else
            MsgBox "Item not found.", vbOKCancel
        End If
        
End Select

Set oLI = Nothing
bDontRefresh = False
Exit Sub

Error:
Call HandleError("mnuItemsPopUpItem_Click")
Set oLI = Nothing
bDontRefresh = False
End Sub

Private Sub mnuJumpToCompare_Click()
If mnuJumpToCompare.Checked = True Then
    mnuJumpToCompare.Checked = False
Else
    mnuJumpToCompare.Checked = True
End If
End Sub

Private Sub mnuMapPopUpItem_Click(Index As Integer)
On Error GoTo Error:

Select Case Index
    Case 0: 'up
        Call MapGoDirection(CellRoom(nMapLastCellIndex, 1), CellRoom(nMapLastCellIndex, 2), "U")
    Case 1: 'down
        Call MapGoDirection(CellRoom(nMapLastCellIndex, 1), CellRoom(nMapLastCellIndex, 2), "D")
    Case 2: 'redraw
        Call MapStartMapping(CellRoom(nMapLastCellIndex, 1), CellRoom(nMapLastCellIndex, 2))
End Select

Exit Sub

Error:
Call HandleError("mnuMapPopUpItem_Click")
End Sub

Private Sub mnuOpenDataFile_Click()
On Error GoTo Error:
Dim sFile As String, fso As FileSystemObject

Set fso = CreateObject("Scripting.FileSystemObject")

oComDag.Filter = "MDB Files (*.mdb)|*.mdb"
oComDag.DialogTitle = "Select Data File"
oComDag.FileName = ReadINI("Settings", "DataFile")

If fso.FolderExists(ReadINI("Settings", "LastDataDir")) Then
    oComDag.InitDir = ReadINI("Settings", "LastDataDir")
Else
    oComDag.InitDir = App.Path
End If

On Error GoTo canceled:
oComDag.ShowOpen
If oComDag.FileName = "" Then GoTo canceled:

On Error GoTo Error:

sFile = oComDag.FileName
If UCase(Right(sFile, 4)) <> ".MDB" Then sFile = sFile & ".mdb"

If Not fso.FileExists(sFile) Then
    MsgBox "File not found."
    Set fso = Nothing
    Exit Sub
End If

Call WriteINI("Settings", "LastDataDir", Left(sFile, Len(sFile) - Len(oComDag.FileTitle) - 1))
    
If Left(sFile, Len(sFile) - Len(oComDag.FileTitle) - 1) = App.Path Then
    Call WriteINI("Settings", "DataFile", oComDag.FileTitle)
Else
    Call WriteINI("Settings", "DataFile", sFile)
End If

Call AppReload(False)
canceled:
Set fso = Nothing
Exit Sub
Error:
Call HandleError("mnuOpenDataFile_Click")
Set fso = Nothing
End Sub

'()()()()()()()())()(()()()()()() mudview lookups

'Private Sub mnuMudviewMonsters_Click()
'    If lvMonsters.ListItems.Count < 1 Then Exit Sub
'    Call MudviewLookup(Monster, lvMonsters.SelectedItem.Text)
'End Sub
'
'Private Sub mnuMudviewWeapons_Click()
'    If lvWeapons.ListItems.Count < 1 Then Exit Sub
'    Call MudviewLookup(Item, lvWeapons.SelectedItem.Text)
'End Sub
'Private Sub mnuMudviewWeaponCompare_Click()
'    If lvWeaponCompare.ListItems.Count < 1 Then Exit Sub
'    Call MudviewLookup(Item, lvWeaponCompare.SelectedItem.Text)
'End Sub
'Private Sub mnuMudviewArmour_Click()
'    If lvArmour.ListItems.Count < 1 Then Exit Sub
'    Call MudviewLookup(Item, lvArmour.SelectedItem.Text)
'End Sub
'Private Sub mnuMudviewArmourCompare_Click()
'    If lvArmourCompare.ListItems.Count < 1 Then Exit Sub
'    Call MudviewLookup(Item, lvArmourCompare.SelectedItem.Text)
'End Sub
'Private Sub mnuMudviewOtherItems_Click()
'    If lvOtherItems.ListItems.Count < 1 Then Exit Sub
'    Call MudviewLookup(Item, lvOtherItems.SelectedItem.Text)
'End Sub
'Private Sub mnuMudviewSpells_Click()
'    If lvSpells.ListItems.Count < 1 Then Exit Sub
'    Call MudviewLookup(Spell, lvSpells.SelectedItem.Text)
'End Sub
'Private Sub mnuMudviewSpellCompare_Click()
'    If lvSpellCompare.ListItems.Count < 1 Then Exit Sub
'    Call MudviewLookup(Spell, lvSpellCompare.SelectedItem.Text)
'End Sub
'Private Sub mnuMudviewShops_Click()
'    If lvShops.ListItems.Count < 1 Then Exit Sub
'    Call MudviewLookup(Shop, lvShops.SelectedItem.Text)
'End Sub
'Private Sub mnuMudviewRaces_Click()
'    If lvRaces.ListItems.Count < 1 Then Exit Sub
'    Call MudviewLookup(Race, lvRaces.SelectedItem.Text)
'End Sub
'Private Sub mnuMudviewClasses_Click()
'    If lvClasses.ListItems.Count < 1 Then Exit Sub
'    Call MudviewLookup(Class, lvClasses.SelectedItem.Text)
'End Sub


'*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*   clipboard copying

Private Sub mnuReload_Click()
Call AppReload(False)
End Sub

Private Sub mnuRemoveFilters_Click()
Call FilterAll(False)
End Sub

Private Sub mnuSettings_Click()

If FormIsLoaded("frmMap") Then frmMap.WindowState = vbMinimized

bReloadProgram = False
frmSettings.Show vbModal, Me
If bReloadProgram Then Call AppReload(True)

End Sub

Private Sub mnuSpellsPopUpItem_Click(Index As Integer)
Dim oLI As ListItem, x As Long, nYesNo As Integer

On Error GoTo Error:

Select Case Index
    Case 0: 'add/remove compare
        bPromptSave = True
        If mnuSpellsPopUpItem(0).Tag = "1" Then 'add
            For Each oLI In objWorkingListView.ListItems
                If oLI.Selected Then
                    tabSpells.Index = "pkSpells"
                    tabSpells.Seek "=", Val(oLI.Text)
                        
                    If Not tabSpells.NoMatch = True Then
                        Set oLI = lvSpellCompare.FindItem(Val(oLI.Text), lvwText, , 0)
                        If Not oLI Is Nothing Then
                            nYesNo = MsgBox(tabSpells.Fields("Name") & " is already on the compare list, add it again?", vbYesNoCancel + vbDefaultButton2 + vbQuestion)
                            If nYesNo = vbNo Then
                                Set oLI = Nothing
                                GoTo nextoli:
                            ElseIf nYesNo = vbCancel Then
                                Set oLI = Nothing
                                Exit Sub
                            End If
                        End If
                        Call AddSpell2LV(lvSpellCompare)
                        
                        If txtSpellCompareDetail.Text = "" Then Call lvSpellCompare_ItemClick(lvSpellCompare.ListItems(1))
                        
                        If mnuJumpToCompare.Checked Then
                            Call cmdNav_Click(3) 'compare
                            Call cmdCompareNav_Click(2)
                        End If
                    End If
                End If
nextoli:
            Next oLI
        ElseIf mnuSpellsPopUpItem(0).Tag = "2" Then 'remove
            bPromptSave = True
            If objWorkingListView.ListItems.Count = 0 Then Exit Sub
            
            x = 1
            Do While x <= objWorkingListView.ListItems.Count
                If objWorkingListView.ListItems(x).Selected Then
                    Call objWorkingListView.ListItems.Remove(x)
                Else
                    x = x + 1
                End If
            Loop
            
            If lvSpellCompare.ListItems.Count > 0 Then
                Call lvSpellCompare_ItemClick(objWorkingListView.ListItems(1))
            Else
                Call cmdCompareClear_Click(2)
            End If

        End If

    Case 1: 'copy to clip
        Select Case objWorkingListView.name
            Case "lvSpells":
                Call CopyLVLinetoClipboard(objWorkingListView, txtSpellDetail, lvSpellLoc, 7)
            Case "lvSpellCompare":
                Call CopyLVLinetoClipboard(objWorkingListView, txtSpellCompareDetail, lvSpellCompareLoc, 7)
        End Select
        
    Case 2: 'casts
        Call LookUpSpellCast(objWorkingListView)
        
End Select

Set oLI = Nothing

Exit Sub

Error:
Call HandleError("mnuSpellsPopUpItem_Click")
Set oLI = Nothing
End Sub

Private Sub mnuToolsItems_Click(Index As Integer)
Dim bClassStealth As Boolean

Select Case Index
    Case 0: 'bs calc
        If Val(txtGlobalLevel(0).Text) > 0 Then
            frmBSCalc.txtLevel.Text = Val(txtGlobalLevel(0).Text)
        End If
        
        frmBSCalc.txtMaxDMG.Text = Val(txtStat(11).Text)
        
        bClassStealth = GetClassStealth(cmbGlobalClass(0).ItemData(cmbGlobalClass(0).ListIndex))
        If bClassStealth Then
            frmBSCalc.chkClassStealth.Value = 1
        Else
            frmBSCalc.chkClassStealth.Value = 0
        End If
        
        frmBSCalc.txtBSMinDMG.Text = Val(txtStat(14).Text)
        frmBSCalc.txtBSMaxDMG.Text = Val(txtStat(15).Text)
        
        frmBSCalc.Show
        frmBSCalc.SetFocus
        
    Case 2: 'exp
        If bNoAlwaysOnTop Then
            frmExpCalc.Show vbModeless
        Else
            frmExpCalc.Show vbModeless, Me
        End If
        
    Case 4: 'swing
        If bNoAlwaysOnTop Then
            frmSwingCalc.Show vbModeless
        Else
            frmSwingCalc.Show vbModeless, Me
        End If
        
    Case 1: 'ext
        Call cmdMapOpenExternal_Click
    
    Case 3: 'note
        frmNotepad.Show
End Select

End Sub

Private Sub ModifyCharStats(Index As Integer)
Dim x As Integer, sFile As String, sSectionName As String
On Error GoTo Error:

If Index > 11 Then
    Select Case Index
        Case 12: 'lvl -
            If Val(txtGlobalLevel(0).Text) > 0 Then
                txtGlobalLevel(0).Text = Val(txtGlobalLevel(0).Text) - 1
            End If
        Case 13: 'lvl +
            txtGlobalLevel(0).Text = Val(txtGlobalLevel(0).Text) + 1
        Case 14: 'reset
            For x = 0 To 5
                txtCharStats(x).Text = Val(txtCharMaxStats(x).Tag)
            Next x
        Case 15: 'reload
            sSectionName = RemoveCharacter(lblDatVer.Caption, " ")
            If bCharLoaded Then
                sFile = ReadINI(sSectionName, "LastCharFile")
                If Not FileExists(sFile) Then
                    sFile = ""
                Else
                    sSectionName = "PlayerInfo"
                End If
            End If
            txtCharStats(0).Text = ReadINI(sSectionName, "Strength", sFile)
            txtCharStats(1).Text = ReadINI(sSectionName, "Intellect", sFile)
            txtCharStats(2).Text = ReadINI(sSectionName, "Widsom", sFile)
            txtCharStats(3).Text = ReadINI(sSectionName, "Agility", sFile)
            txtCharStats(4).Text = ReadINI(sSectionName, "Health", sFile)
            txtCharStats(5).Text = ReadINI(sSectionName, "Charm", sFile)
            
    End Select
    Exit Sub
End If

If ((Index + 2) Mod 2) > 0 Then '+
    If Val(txtCharStats((Index - 1) / 2).Text) > 9999 Then Exit Sub
    txtCharStats((Index - 1) / 2).Text = Val(txtCharStats((Index - 1) / 2).Text) + 1
Else '-
    If Val(txtCharStats(Index / 2).Text) < 1 Then Exit Sub
    txtCharStats(Index / 2).Text = Val(txtCharStats(Index / 2).Text) - 1
End If

Exit Sub

Error:
Call HandleError("ModifyCharStats")
End Sub

Private Sub PasteCharacter()
On Error GoTo Error:
Dim sSearch As String, sText As String, sChar As String, sField As String
Dim x As Integer, y As Integer, x2 As Integer
Dim sEquipLoc(0 To 18) As String, bResult As Boolean, nTries As Integer
Dim sLevel As String, sRaceName As String, sClassName As String
Dim nEncum As Long, nStat As String, sName As String

'x = current position in string
'y = length of next possible (current) string match
'x2 = starting point of field match

If tabItems.RecordCount = 0 Then
    MsgBox "No item records", vbExclamation
    Exit Sub
End If

Me.Enabled = False
Load frmPasteChar
frmPasteChar.txtText = Clipboard.GetText
frmPasteChar.Tag = "-1"
frmPasteChar.Show vbModal, Me
Me.Enabled = True
frmPasteChar.Hide

If frmPasteChar.Tag = "-1" Then GoTo canceled:

sSearch = frmPasteChar.txtText.Text

If Len(sSearch) < 10 Then GoTo canceled:

x = 1
y = 1
x2 = -1
Do Until x + y > Len(sSearch) + 1
    
    sChar = Mid(sSearch, x + y - 1, 1)
    
    bResult = TestPasteChar(sChar)
    If bResult = False Then GoTo next_y:
    
    sText = RemoveCharacter(sText & sChar, " ")
    'If Right(sText, 2) = "  " Then sText = Left(sText, Len(sText) - 1)
    
    If Not InStr(1, LCase(sText), "equippedwith:") = 0 Then
        GoTo clear:
    ElseIf Not InStr(1, LCase(sText), "arecarrying") = 0 Then
        GoTo clear:
    End If
    
    Select Case sChar
        Case ",":
            GoTo clear:
        Case "(":
            x2 = Len(sText)
        Case ")":
            If x2 = -1 Then GoTo clear:
            
            Select Case UCase(Mid(sText, x2 + 1, Len(sText) - x2 - 1))
                Case "HEAD": sEquipLoc(0) = Left(sText, x2 - 1)
                Case "EARS": sEquipLoc(1) = Left(sText, x2 - 1)
                Case "EYES": sEquipLoc(3) = Left(sText, x2 - 1)
                Case "FACE": sEquipLoc(3) = Left(sText, x2 - 1)
                Case "NECK": sEquipLoc(2) = Left(sText, x2 - 1)
                Case "BACK": sEquipLoc(3) = Left(sText, x2 - 1)
                Case "TORSO": sEquipLoc(4) = Left(sText, x2 - 1)
                Case "ARMS": sEquipLoc(5) = Left(sText, x2 - 1)
                Case "WRIST":
                    If Not sEquipLoc(6) = "" Then
                        If sEquipLoc(7) = "" Then
                            sEquipLoc(7) = Left(sText, x2 - 1)
                        End If
                    Else
                        sEquipLoc(6) = Left(sText, x2 - 1)
                    End If
                Case "WAIST": sEquipLoc(11) = Left(sText, x2 - 1)
                Case "FINGER":
                    If Not sEquipLoc(9) = "" Then
                        If sEquipLoc(10) = "" Then
                            sEquipLoc(10) = Left(sText, x2 - 1)
                        End If
                    Else
                        sEquipLoc(9) = Left(sText, x2 - 1)
                    End If
                Case "HANDS": sEquipLoc(8) = Left(sText, x2 - 1)
                Case "LEGS": sEquipLoc(12) = Left(sText, x2 - 1)
                Case "FEET": sEquipLoc(13) = Left(sText, x2 - 1)
                Case "WORN": sEquipLoc(14) = Left(sText, x2 - 1)
                Case "OFF-HAND": sEquipLoc(15) = Left(sText, x2 - 1)
                Case "WEAPONHAND": sEquipLoc(16) = Left(sText, x2 - 1)
                Case "TWOHANDED": sEquipLoc(16) = Left(sText, x2 - 1)
            End Select
            
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

x = InStr(1, sSearch, "Race: ")
If x > 0 Then
    x = x + 6 '6=len("race: ")
    y = InStr(x, sSearch, "Exp:") 'exp is the next thing in the string for stats
    If y > x + 20 Then y = 0 'just incase "exp:" is somewhere way down in the paste
    If y > 0 Then
        If InStr(1, LTrim(RTrim(Mid(sSearch, x, y - x))), Chr(10)) > 0 Then y = 0
        'if there is a carriage return inside the matched string
    End If
    If y = 0 Then y = InStr(x, sSearch, Chr(13))
    If y = 0 Then y = InStr(x, sSearch, Chr(10))
    If y > x Then sRaceName = LTrim(RTrim(Mid(sSearch, x, y - x)))
End If

x = InStr(1, sSearch, "Class: ")
If x > 0 Then
    x = x + 7 '7=len("class: ")
    y = InStr(x, sSearch, "Level:")
    If y > x + 15 Then y = 0
    If y > 0 Then
        If InStr(1, LTrim(RTrim(Mid(sSearch, x, y - x))), Chr(10)) > 0 Then y = 0
    End If
    If y = 0 Then y = InStr(x, sSearch, Chr(13))
    If y = 0 Then y = InStr(x, sSearch, Chr(10))
    If y > x Then sClassName = LTrim(RTrim(Mid(sSearch, x, y - x)))
End If

x = InStr(1, sSearch, "Name: ")
If x > 0 Then
    x = x + 6 '6=len("name: ")
    y = InStr(x, sSearch, "Lives/CP:")
    If y > x + 35 Then y = 0
    If y > 0 Then
        If InStr(1, LTrim(RTrim(Mid(sSearch, x, y - x))), Chr(10)) > 0 Then y = 0
    End If
    If y = 0 Then y = InStr(x, sSearch, Chr(13))
    If y = 0 Then y = InStr(x, sSearch, Chr(10))
    If y > x Then sName = LTrim(RTrim(Mid(sSearch, x, y - x)))
End If


bDontRefresh = True

nStat = ExtractValueFromString(sSearch, "Level:")
If nStat > 0 Then txtGlobalLevel(0).Text = nStat

nStat = ExtractValueFromString(sSearch, "Strength:")
If nStat > 0 Then txtCharStats(0).Text = nStat

nStat = ExtractValueFromString(sSearch, "Intellect:")
If nStat > 0 Then txtCharStats(1).Text = nStat

nStat = ExtractValueFromString(sSearch, "Willpower:")
If nStat > 0 Then txtCharStats(2).Text = nStat

nStat = ExtractValueFromString(sSearch, "Agility:")
If nStat > 0 Then txtCharStats(3).Text = nStat

nStat = ExtractValueFromString(sSearch, "Health:")
If nStat > 0 Then txtCharStats(4).Text = nStat

nStat = ExtractValueFromString(sSearch, "Charm:")
If nStat > 0 Then txtCharStats(5).Text = nStat

If Not sName = "" Then txtCharName = sName

nEncum = Val(ExtractValueFromString(sSearch, "Encumbrance:"))

If Not sRaceName = "" Then
    If cmbGlobalRace(0).ListCount > 0 Then
        For x = 0 To cmbGlobalRace(0).ListCount - 1
            If cmbGlobalRace(0).List(x) = sRaceName Then
                cmbGlobalRace(0).ListIndex = x
            End If
        Next
    End If
End If

If Not sClassName = "" Then
    If cmbGlobalClass(0).ListCount > 0 Then
        For x = 0 To cmbGlobalClass(0).ListCount - 1
            If cmbGlobalClass(0).List(x) = sClassName Then
                cmbGlobalClass(0).ListIndex = x
            End If
        Next
    End If
End If

'add to inven
tabItems.MoveFirst
DoEvents
Do Until tabItems.EOF
    
'    If tabItems.Fields("Number") = 56 Then
'        Debug.Print tabItems.Fields("Number")
'    End If
    
    If bOnlyInGame And tabItems.Fields("In Game") = 0 Then GoTo skip:
    
    sText = RemoveCharacter(tabItems.Fields("Name"), " ")
    For x = 0 To cmbEquip().UBound
        If sText = sEquipLoc(x) Then
'            If x = 3 Then
'                Debug.Print x
'            End If
            If x = 7 And Not bInvenUse2ndWrist Then GoTo skip:
            If cmbEquip(x).ListCount > 0 Then
                nTries = 0
tryagain:
                For y = 0 To cmbEquip(x).ListCount - 1
                    If cmbEquip(x).ItemData(y) = tabItems.Fields("Number") Then
                        
                        If nEncum > 0 Then
                            nEncum = nEncum - tabItems.Fields("Encum")
                        Else
                            nEncum = 0
                        End If
                        
                        sEquipLoc(x) = ""
                        If chkEquipHold(x).Value = 1 Then GoTo skip:
                        
                        cmbEquip(x).ListIndex = y
                        GoTo skip:
                    End If
                Next y
                
                If nTries > 0 Then GoTo skip:
                Call InvenAddEquip(tabItems.Fields("Number"), tabItems.Fields("Name"), tabItems.Fields("ItemType"), tabItems.Fields("Worn"))
                nTries = 1
                GoTo tryagain:
            End If
        End If
    Next x
skip:
    tabItems.MoveNext
Loop

If nEncum > 0 Then
    txtInvenAddWeight.Text = nEncum
    If chkInvenAddWeight.Value = 0 Then chkInvenAddWeight.Value = 1
End If

canceled:
bDontRefresh = False
Call RefreshAll
Exit Sub
Error:
Call HandleError("PasteCharacter")
Me.Enabled = True
bDontRefresh = False
End Sub

Public Sub PopUpAuxMenu(objWorkingLV As ListView)
On Error GoTo Error:

Set objWorkingListView = objWorkingLV

If objWorkingListView.name = "lvOtherItemLoc" Then 'chests
    If cmdSundryChests.Enabled = False Then
        Set objWorkingListView = lvOtherItems
    End If
End If

Select Case objWorkingListView.name
    Case "lvMonsters":
        mnuAuxPopUpItem(3).Tag = "1"
        mnuAuxPopUpItem(3).Caption = "Add to Compare"
        mnuAuxPopUpItem(0).Visible = True   'copy to clip
        mnuAuxPopUpItem(1).Visible = True   'search monster
        mnuAuxPopUpItem(2).Visible = False  'shop add to item compare
        mnuAuxPopUpItem(3).Visible = True   'add to monster compare
        mnuAuxPopUpItem(4).Visible = False  'copy chest
    Case "lvMonsterCompare":
        mnuAuxPopUpItem(3).Tag = "2"
        mnuAuxPopUpItem(3).Caption = "Remove from Compare"
        mnuAuxPopUpItem(0).Visible = True   'copy to clip
        mnuAuxPopUpItem(1).Visible = True   'search monster
        mnuAuxPopUpItem(2).Visible = False  'shop add to item compare
        mnuAuxPopUpItem(3).Visible = True   'add to monster compare
        mnuAuxPopUpItem(4).Visible = False  'copy chest
    Case "lvShopDetail":
        mnuAuxPopUpItem(0).Visible = True   'copy to clip
        mnuAuxPopUpItem(1).Visible = False  'search monster
        mnuAuxPopUpItem(2).Visible = True   'shop add to item compare
        mnuAuxPopUpItem(3).Visible = False  'add to monster compare
        mnuAuxPopUpItem(4).Visible = False  'copy chest
    Case "lvOtherItemLoc":
        mnuAuxPopUpItem(4).Visible = True   'copy chest
        mnuAuxPopUpItem(0).Visible = False  'copy to clip
        mnuAuxPopUpItem(1).Visible = False  'search monster
        mnuAuxPopUpItem(2).Visible = False  'shop add to item compare
        mnuAuxPopUpItem(3).Visible = False  'add to monster compare
    Case Else:
        mnuAuxPopUpItem(0).Visible = True   'copy to clip
        mnuAuxPopUpItem(1).Visible = False  'search monster
        mnuAuxPopUpItem(2).Visible = False  'shop add to item compare
        mnuAuxPopUpItem(3).Visible = False  'add to monster compare
        mnuAuxPopUpItem(4).Visible = False  'copy chest
End Select

DoEvents

PopupMenu mnuAuxPopUp

Exit Sub

Error:
Call HandleError("PopUpAuxMenu")

End Sub

Private Sub PopUpItemsMenu(objWorkingLV As ListView)
On Error GoTo Error:

'Case 0: 'add/remove compare
'Case 1: 'calc swings
'Case 2: 'copy to clip
'Case 3: 'equip

Set objWorkingListView = objWorkingLV

Select Case objWorkingListView.name
    Case "lvWeapons":
        mnuItemsPopUpItem(0).Caption = "Add To Compare"
        mnuItemsPopUpItem(0).Tag = "1"
        mnuItemsPopUpItem(1).Visible = True  'swings
        mnuItemsPopUpItem(3).Visible = True 'equip
    Case "lvWeaponCompare":
        mnuItemsPopUpItem(0).Caption = "Remove From Compare"
        mnuItemsPopUpItem(0).Tag = "2"
        mnuItemsPopUpItem(1).Visible = True  'swings
        mnuItemsPopUpItem(3).Visible = True 'equip
    Case "lvArmour":
        mnuItemsPopUpItem(0).Caption = "Add To Compare"
        mnuItemsPopUpItem(0).Tag = "1"
        mnuItemsPopUpItem(1).Visible = False 'swings
        mnuItemsPopUpItem(3).Visible = True 'equip
    Case "lvArmourCompare":
        mnuItemsPopUpItem(0).Caption = "Remove From Compare"
        mnuItemsPopUpItem(0).Tag = "2"
        mnuItemsPopUpItem(1).Visible = False 'swings
        mnuItemsPopUpItem(3).Visible = True 'equip
    Case "lvShopDetail":
        mnuItemsPopUpItem(0).Caption = "Add To Compare"
        mnuItemsPopUpItem(0).Tag = "1"
        mnuItemsPopUpItem(1).Visible = False 'swings
        mnuItemsPopUpItem(3).Visible = True 'equip
    Case "lvMonsterCompare":
        mnuItemsPopUpItem(0).Caption = "Remove From Compare"
        mnuItemsPopUpItem(0).Tag = "2"
        mnuItemsPopUpItem(1).Visible = False 'swings
        mnuItemsPopUpItem(3).Visible = False 'equip
    Case Else:
        Exit Sub
End Select

DoEvents

PopupMenu mnuItemsPopUp

Exit Sub

Error:
Call HandleError("PopUpItemMenu")

End Sub

Public Sub PopUpMapMenu(ByVal bUp As Boolean, bDown As Boolean)
On Error GoTo Error:


If bUp Then mnuMapPopUpItem(0).Visible = True Else mnuMapPopUpItem(0).Visible = False
If bDown Then mnuMapPopUpItem(1).Visible = True Else mnuMapPopUpItem(1).Visible = False

DoEvents
PopupMenu mnuMapPopUp

Exit Sub

Error:
Call HandleError("PopUpMapMenu")

End Sub

Private Sub PopUpSpellsMenu(objWorkingLV As ListView)
On Error GoTo Error:

'Case 0: 'add/remove compare
'Case 1: 'copy to clip
'Case 2: 'casts?

Set objWorkingListView = objWorkingLV

Select Case objWorkingListView.name
    Case "lvSpells":
        mnuSpellsPopUpItem(0).Caption = "Add To Compare"
        mnuSpellsPopUpItem(0).Tag = "1"
    Case "lvSpellCompare":
        mnuSpellsPopUpItem(0).Caption = "Remove From Compare"
        mnuSpellsPopUpItem(0).Tag = "2"
    Case Else:
        Exit Sub
End Select

DoEvents

PopupMenu mnuSpellsPopUp

Exit Sub

Error:
Call HandleError("PopUpSpellMenu")

End Sub

Private Sub ProcessListViewClick(ByRef objListItem As ListItem, ByRef objDetailText As TextBox, _
    ByRef objLocationLV As ListView)
On Error GoTo Error:

If objListItem Is Nothing Then Exit Sub

tabItems.Index = "pkItems"
tabItems.Seek "=", Val(objListItem.Text)
If tabItems.NoMatch = True Then
    MsgBox "Record not found."
    Exit Sub
Else
    Call PullItemDetail(objDetailText, objLocationLV)
End If

objListItem.Selected = True
objListItem.EnsureVisible

Exit Sub

Error:
Call HandleError("ProcessListViewClick")
End Sub

Public Function RecentFileAdd(Optional ByVal sFile As String)
Dim x As Integer, y As Integer, sName As String, sFileTitle() As String
On Error GoTo Error:

'first check if this is a recent file
If Not sFile = "" Then
    For y = 1 To 5
        If sRecentFiles(y, 2) = sFile Then
            'it is, move only what's above it down
            If y > 1 Then
                For x = y To 2 Step -1
                    sRecentFiles(x, 2) = sRecentFiles(x - 1, 2)
                Next x
            End If
            Exit For
        End If
        'it's not, move all files down a level ...
        If y = 5 Then
            For x = 5 To 2 Step -1
                sRecentFiles(x, 2) = sRecentFiles(x - 1, 2)
            Next x
        End If
    Next y
    sRecentFiles(1, 2) = sFile
End If

For x = 1 To 5
    If FileExists(sRecentFiles(x, 2)) = False Then
        sRecentFiles(x, 1) = "-none-"
        sRecentFiles(x, 2) = ""
    Else
        If InStr(1, sRecentFiles(x, 2), "\") > 0 Then
            sFileTitle() = Split(sRecentFiles(x, 2), "\", , vbTextCompare)
        Else
            ReDim sFileTitle(0)
            sFileTitle(0) = sRecentFiles(x, 2)
        End If
        
        If bNameInTitle Then
            sName = ReadINI("PlayerInfo", "Name", sRecentFiles(x, 2))
            If sName = "" Or sName = "0" Then sName = sFileTitle(UBound(sFileTitle()))
        Else
            sName = sFileTitle(UBound(sFileTitle()))
        End If
        If Len(sName) > 33 Then
            sName = Left(sName, 10) & "..." & Right(sName, 20)
        End If
        sRecentFiles(x, 1) = sName
    End If
Next x

Call RecentFilesUpdate

Erase sFileTitle()
Exit Function
Error:
Call HandleError("RecentFileAdd")
Erase sFileTitle()
End Function

Public Function RecentFilesLoad(ByVal nNumber As String)
Dim nYesNo As Integer, nResult As Integer, sSectionName As String
On Error GoTo Error:

If FileExists(sRecentFiles(nNumber, 2)) = False Then
    MsgBox "File not found.", vbInformation
    Exit Function
End If

If bCharLoaded Then
    sSectionName = RemoveCharacter(lblDatVer.Caption, " ")
    If bAutoSave Then
        If FileExists(ReadINI(sSectionName, "LastCharFile")) Then
            nResult = SaveCharacter(False, ReadINI(sSectionName, "LastCharFile"))
            If nResult = -1 Then Exit Function
        Else
            If bPromptSave = True Then Call SaveCharacter(True)
        End If
    Else
        If bPromptSave = True Then
            nYesNo = MsgBox("Save character file first?", vbYesNoCancel + vbQuestion + vbDefaultButton3)
            If nYesNo = vbYes Then
                If FileExists(ReadINI(sSectionName, "LastCharFile")) Then
                    nResult = SaveCharacter(False, ReadINI(sSectionName, "LastCharFile"))
                Else
                    nResult = SaveCharacter(True)
                End If
                If nResult = -1 Then Exit Function
            ElseIf nYesNo = vbCancel Then
                Exit Function
            End If
        End If
    End If
End If

Call LoadCharacter(False, sRecentFiles(nNumber, 2), True)

Exit Function
Error:
Call HandleError("LoadCharacterFile")
End Function

Private Sub RecentFilesUpdate()
Dim newItem As String 'holds the item that the user wants to add to the menu
Dim hMenu As Long 'handle to this form's menu
Dim hSubMenu 'handle to one of the sub menus
'Dim iNum As Long 'the number of items currently on the menu
Dim menuInfo As MENUITEMINFO 'holds info about the current menu
Dim nItemNum As Integer
Dim x As Integer

If bNO_RECENT_FILES Then Exit Sub

nItemNum = 7 'location in File menu to start inserting

hMenu = GetMenu(Me.hwnd) 'retreive a handle to this form's menu
hSubMenu = GetSubMenu(hMenu, 0) 'retreive a handle to the menu (0-based array)
'iNum = GetMenuItemCount(hSubMenu) 'determine how many items are currently on this menu

If nMenuItemID = 1000 Then 'if this is the first recent added then add a separator bar to the menu
    With menuInfo 'add a separator bar to this menu
        .cbSize = Len(menuInfo) 'set the length of the menu structure
        .fMask = MIIM_ID Or MIIM_TYPE 'declare which parts of the menu structure to use
        .fType = MFT_SEPARATOR 'define the type of menu item (separator)
        .wID = nMenuItemID 'set the structure ID
    End With
    Call InsertMenuItem(hSubMenu, nItemNum, 1, menuInfo) 'add the separator bar to the menu
    nMenuItemID = nMenuItemID + 1 'increment for the next new menu item
    nItemNum = nItemNum + 1
    
    For x = 1 To 5
        With menuInfo 'add the new menu item to the top of this menu
            .cbSize = Len(menuInfo) 'set the length of the menu structure
            .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE 'declare which parts of the menu structure to use
            .fType = MFT_STRING 'define the type of menu item (text)
            .fState = MFS_ENABLED 'this item should be enabled on the menu
            .dwTypeData = sRecentFiles(x, 1) 'the text of the new menu item
            .cch = Len(.dwTypeData)
            .wID = nMenuItemID 'nMenuItemID 'set the structure ID (this ID is used to add functionality to this menu item)
        End With
        Call InsertMenuItem(hSubMenu, nItemNum, 1, menuInfo)
        nItemNum = nItemNum + 1
        nMenuItemID = nMenuItemID + 1 'increment for the next new menu item
    Next x
Else
    nItemNum = nItemNum + 1
    For x = 1 To 5
        With menuInfo 'add the new menu item to the top of this menu
            .cbSize = Len(menuInfo) 'set the length of the menu structure
            .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE 'declare which parts of the menu structure to use
            .fType = MFT_STRING 'define the type of menu item (text)
            .fState = MFS_ENABLED 'this item should be enabled on the menu
            .dwTypeData = sRecentFiles(x, 1) 'the text of the new menu item
            .cch = Len(.dwTypeData)
            .wID = x + 1000 'nMenuItemID 'set the structure ID (this ID is used to add functionality to this menu item)
        End With
        Call SetMenuItemInfo(hSubMenu, nItemNum, 1, menuInfo)
        nItemNum = nItemNum + 1
        'Call InsertMenuItem(hSubMenu, nItemNum, 1, menuInfo)
        'nItemNum = nItemNum + 1
        'nMenuItemID = nMenuItemID + 1 'increment for the next new menu item
    Next x
End If


'With menuInfo 'add the new menu item to the top of this menu
'    .cbSize = Len(menuInfo) 'set the length of the menu structure
'    .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE 'declare which parts of the menu structure to use
'    .fType = MFT_STRING 'define the type of menu item (text)
'    .fState = MFS_ENABLED 'this item should be enabled on the menu
'    .dwTypeData = sFile 'the text of the new menu item
'    .cch = Len(.dwTypeData)
'    .wID = nMenuItemID 'set the structure ID (this ID is used to add functionality to this menu item)
'End With
'Call InsertMenuItem(hSubMenu, 8, 1, menuInfo)
'nMenuItemID = nMenuItemID + 1 'increment for the next new menu item
    
End Sub

Private Sub RefreshAll()

If bDontRefresh Then Exit Sub

Call InvenCalcStats

Call RefreshCPs
Call RefreshHitPoints
Call RefreshMagic
Call RefreshPicklocks
Call RefreshShops
Call RefreshCharBless

End Sub

Private Sub RefreshCharBless()
Dim x As Integer, nTotal As Double, nSetLevel As Long, sQuick As String
Dim nDur As Double, nDurIncr As Long, nDurLVLs As Long, nLevel As Long

On Error GoTo Error:

If bDontRefresh Then Exit Sub

If tabSpells.RecordCount = 0 Then Exit Sub

tabSpells.Index = "pkSpells"
nSetLevel = Val(txtGlobalLevel(0).Text)
If nSetLevel = 0 Then nSetLevel = 1

For x = 0 To 9
    objToolTip.DelToolTip cmbCharBless(x).hwnd
    
    If cmbCharBless(x).ListIndex >= 0 Then
        If cmbCharBless(x).ItemData(cmbCharBless(x).ListIndex) > 0 Then
            tabSpells.Seek "=", cmbCharBless(x).ItemData(cmbCharBless(x).ListIndex)
            If tabSpells.NoMatch = False Then
                nLevel = nSetLevel
                
                If nLevel > tabSpells.Fields("Cap") And tabSpells.Fields("Cap") > 0 Then nLevel = tabSpells.Fields("Cap")
                If nLevel < tabSpells.Fields("ReqLevel") Then nLevel = tabSpells.Fields("ReqLevel")
                
                nDur = tabSpells.Fields("Dur")
                nDurIncr = tabSpells.Fields("DurInc")
                nDurLVLs = tabSpells.Fields("DurIncLVLs")
                
                If Not nDurLVLs = 0 And Not nDurIncr = 0 Then
                    nDur = nDur + (Round(nDurIncr / nDurLVLs, 3) * nLevel)
                End If
                
                'nSpellNest = 0
                If chkGlobalFilter.Value = 1 Then
                    sQuick = PullSpellEQ(True, nSetLevel)
                Else
                    sQuick = PullSpellEQ(False)
                End If
                
                objToolTip.SetToolTipObj cmbCharBless(x).hwnd, tabSpells.Fields("Name") _
                    & " (" & tabSpells.Fields("Short") & ") -- Mana: " _
                    & tabSpells.Fields("ManaCost") & vbCrLf & "EQ: " & sQuick, False
                
                nTotal = nTotal + Round(tabSpells.Fields("ManaCost") / nDur, 3)
            End If
        End If
    Else
        If cmbCharBless(x).ListCount > 0 Then
            cmbCharBless(x).ListIndex = 0
            Exit Sub
        End If
    End If
Next x

nTotal = nTotal * 10
lblCharBless.Caption = nTotal

Exit Sub

Error:
Call HandleError("RefreshCharBless")

End Sub

Private Sub RefreshCPs()
Dim x As Integer, nCPUsed As Long, y As Integer, nCP As Long
Dim nLevelReq As Long, nBaseCP As Long

On Error GoTo Error:

For x = 0 To 5
    objToolTip.DelToolTip txtCharMaxStats(x).hwnd
    
    nBaseCP = 0
    nCP = Val(txtCharStats(x).Text) - Val(txtCharMaxStats(x).Tag)
    If nCP < 0 Then nCP = 0
    For y = 1 To Fix(nCP / 10)
        If y = 10 Then Exit For
        nBaseCP = nBaseCP + (10 * y)
    Next y
    If y = 10 Then
        nBaseCP = nBaseCP + ((nCP - 90) * y)
    Else
        nBaseCP = nBaseCP + ((nCP Mod 10) * y)
    End If
    
    If nBaseCP > 0 Then
        objToolTip.SetToolTipObj txtCharMaxStats(x).hwnd, "CP Used: " & nBaseCP, False
    End If
    
    nCPUsed = nCPUsed + nBaseCP
Next x

nLevelReq = 1
nBaseCP = GetRaceCP(cmbGlobalRace(0).ItemData(cmbGlobalRace(0).ListIndex))
nCP = nBaseCP
Do While nCP < nCPUsed
    nCP = nCP + (Fix((nLevelReq) / 10) * 5) + 10
    nLevelReq = nLevelReq + 1
Loop

If Val(txtGlobalLevel(0).Text) > 0 Then
    nCP = nBaseCP + CalcCPLevel(Val(txtGlobalLevel(0).Text))
End If

lblStatCalc.Caption = "Level Required:  " & nLevelReq _
    & vbCrLf & "CPs Used/Avail:  " & nCPUsed & "/" & nCP - nCPUsed
lblStatCalc.Tag = nCP - nCPUsed

If Not txtInvenStrength.Text = txtCharStats(0).Text Then
    txtInvenStrength.Text = txtCharStats(0).Text
End If

If Not txtShopCharm.Text = txtCharStats(5).Text Then
    txtShopCharm.Text = txtCharStats(5).Text
End If

Exit Sub
Error:
Call HandleError("RefreshCPs")

End Sub

Private Sub RefreshHitPoints()
On Error GoTo Error:
Dim sMin As String, nMin As Long, sMax As String, nMax As Long  ', nRaceBonus As Long
Dim sTag As String

lblCharRestRate.Caption = "Normal:  " & CalcRestingRate(Val(txtGlobalLevel(0).Text), _
    Val(txtCharStats(4).Text), Val(txtCharHPRegen.Text)) _
    & vbCrLf & "Resting:  " & CalcRestingRate(Val(txtGlobalLevel(0).Text), _
    Val(txtCharStats(4).Text), Val(txtCharHPRegen.Text), True)

If cmbGlobalClass(0).ListIndex > 0 Then
    nMin = GetClassMinHP(cmbGlobalClass(0).ItemData(cmbGlobalClass(0).ListIndex))
    nMax = GetClassMaxHP(cmbGlobalClass(0).ItemData(cmbGlobalClass(0).ListIndex))
    
'    If cmbGlobalRace(0).ListIndex > 0 Then
'        nRaceBonus = Val(txtGlobalLevel(0).Text) * _
'            GetRaceHPBonus(cmbGlobalRace(0).ItemData(cmbGlobalRace(0).ListIndex))
'    End If
    
    sMin = CalcMaxHP(nMax - nMin, Val(txtGlobalLevel(0).Text), _
        Val(txtCharStats(4).Text), nMin) '+ nRaceBonus
    sTag = Val(sMin) + Val(txtStat(5).Text)
    
    sMax = CalcMaxHP((nMax - nMin) * Val(txtGlobalLevel(0).Text), _
        Val(txtGlobalLevel(0).Text), Val(txtCharStats(4).Text), nMin)
    
    sTag = sTag & "-" & (Val(sMax) + Val(txtStat(5).Text))
    
    lblCharMaxHP.Caption = "HP Range: " & sMin & "-" & sMax
    lblCharMaxHP.Tag = sTag
    
    If Not Val(txtStat(5).Text) = 0 Then
        If Val(txtStat(5).Text) > 0 Then
            lblCharMaxHP.Caption = lblCharMaxHP.Caption & " (+" & Val(txtStat(5).Text) & ")"
        Else
            lblCharMaxHP.Caption = lblCharMaxHP.Caption & " (" & Val(txtStat(5).Text) & ")"
        End If
    End If
Else
    lblCharMaxHP.Caption = "HP Range: ? - ?"
End If

Exit Sub
Error:
Call HandleError("CalcRestingRatePrivate")
End Sub

Private Sub RefreshMagic()
On Error GoTo Error:
Dim nMageryLVL As Integer, nMagery As enmMagicEnum, nLevel As Long
Dim nSC As Long, nSCBonsus As Long, nMaxMana As Long, nMaxManaBonus As Long
Dim sTemp As String

If cmbGlobalClass(0).ListIndex > 0 Then
    nLevel = Val(txtGlobalLevel(0).Text)
    nMagery = GetClassMagery(cmbGlobalClass(0).ItemData(cmbGlobalClass(0).ListIndex))
    nMageryLVL = GetClassMageryLVL(cmbGlobalClass(0).ItemData(cmbGlobalClass(0).ListIndex))
    
    lblCharManaRate.Caption = "Mana Regen: " & CalcManaRegen(nLevel, _
        Val(txtCharStats(1).Text), Val(txtCharStats(2).Text), Val(txtCharStats(5).Text), _
        nMageryLVL, nMagery, Val(txtCharManaRegen.Text), False) & vbCrLf _
        & "Medi. Ticks: " & CalcManaRegen(nLevel, _
        Val(txtCharStats(1).Text), Val(txtCharStats(2).Text), Val(txtCharStats(5).Text), _
        nMageryLVL, nMagery, , True)
    
    nMaxMana = CalcMaxMana(nLevel, nMageryLVL)
    nMaxManaBonus = Val(txtStat(6).Text)
    If Not nMaxManaBonus = 0 Then
    
        If nMaxManaBonus > 0 Then
            sTemp = " +"
        Else
            sTemp = " "
        End If
        
        lblCharMaxMana.Caption = "Max Mana: " & nMaxMana + nMaxManaBonus _
            & " (" & nMaxMana & sTemp & nMaxManaBonus & ")"
        lblCharMaxMana.Tag = nMaxMana + nMaxManaBonus
    Else
        lblCharMaxMana.Caption = "Max Mana: " & nMaxMana
        lblCharMaxMana.Tag = nMaxMana
    End If
    
    nSC = CalcSpellCasting(nLevel, _
        Val(txtCharStats(1).Text), Val(txtCharStats(2).Text), Val(txtCharStats(5).Text), _
        nMageryLVL, nMagery)
    nSCBonsus = Val(txtStat(9).Text)
    If Not nSCBonsus = 0 Then
        
        If nMaxManaBonus > 0 Then
            sTemp = " +"
        Else
            sTemp = " "
        End If
        
        lblCharSC.Caption = "Spellcasting: " & nSC + nSCBonsus _
            & " (" & nSC & sTemp & nSCBonsus & ")"
        lblCharSC.Tag = nSC + nSCBonsus
    Else
        lblCharSC.Caption = "Spellcasting: " & nSC
        lblCharSC.Tag = nSC
    End If
Else
    lblCharManaRate.Caption = "Mana Regen: ?"
    lblCharSC.Caption = "Spellcasting: ?"
    lblCharMaxMana.Caption = "Max Mana: ?"
End If
Exit Sub

Error:
Call HandleError("RefreshMagic")
End Sub

Private Sub RefreshPicklocks()
Dim nBasePicks As Long, nBonusPicks As Long
Dim sTemp As String
On Error GoTo Error:

nBasePicks = CalcPicklocks(Val(txtGlobalLevel(0).Text), _
    Val(txtCharStats(3).Text), Val(txtCharStats(1).Text))
nBonusPicks = Val(txtStat(22).Text)

If Not nBonusPicks = 0 Then
    If nBonusPicks > 0 Then
        sTemp = " +"
    Else
        sTemp = " "
    End If
    
    lblCharPicklocks.Caption = "Picklocks: " & (nBasePicks + nBonusPicks) _
        & " (" & nBasePicks & sTemp & nBonusPicks & ")"
    lblCharPicklocks.Tag = (nBasePicks + nBonusPicks)
Else
    lblCharPicklocks.Caption = "Picklocks: " & nBasePicks
    lblCharPicklocks.Tag = nBasePicks
End If

Exit Sub

Error:
Call HandleError("RefreshPicklocks")
    
End Sub

Private Sub RefreshShops()

On Error GoTo Error:

If lvShops.ListItems.Count > 0 Then
    Call lvShops_ItemClick(lvShops.SelectedItem)
End If

Exit Sub
Error:
Call HandleError("RefreshShops")

End Sub

Private Sub ResetFilterOptions(Optional bNotArmour As Boolean, _
    Optional bNotWeapons As Boolean, Optional bNotSpells As Boolean, _
    Optional bClassFiltersOnly As Boolean)
Dim x As Integer
On Error GoTo Error:

If Not bNotArmour Then
    'armour opts
    For x = 0 To 6
        chkArmourType(x).Value = 1
    Next
    chkArmourNoLimit.Value = 0
    chkArmourNonMagic.Value = 0
    If Not bClassFiltersOnly Then
        cmbArmourWorn.ListIndex = 0
    End If
End If

If Not bNotWeapons Then
    'weapon opts
    chkWeaponStaffOnly.Value = 0
    For x = 0 To 3
        chkHanded(x).Value = 1
    Next
    chkWeaponNonMagical.Value = 0
    If Not bClassFiltersOnly Then
        chkWeaponNoLimit.Value = 0
        cmbWeaponSpeed.ListIndex = 0
        txtWeaponSpeed.Text = 9999
        cmbWeaponMagicLevel.ListIndex = 0
    End If
End If

If Not bNotSpells Then
    'spell opts
    cmbSpellMagery.ListIndex = 0
    cmbSpellMageryLevel.ListIndex = 0
    If Not bClassFiltersOnly Then
        chkSpellLearnable.Value = 0
        cmbSpellAttackType.ListIndex = 0
        cmbSpellTarget.ListIndex = 0
        cmbContainsAbil(0).ListIndex = 0
    End If
End If

Exit Sub

Error:
Call HandleError("ResetFilterOptions")
    
End Sub

Public Function SaveCharacter(ByVal bPromptForFile As Boolean, Optional ByVal strFileName As String) As Integer
On Error GoTo Error:
Dim sFile As String, str As String, oLI As ListItem, sName As String
Dim sSectionName As String, x As Integer, sFileTitle As String, y As Integer

sSectionName = RemoveCharacter(lblDatVer.Caption, " ")

If bPromptForFile Then
    oComDag.Filter = "MMUD Explorer Character (*.mmec)|*.mmec"
    oComDag.DialogTitle = "Save Character"
    If Len(ReadINI(sSectionName, "LastSaveName")) < 4 Then
        oComDag.FileName = "Character.mmec"
    Else
        oComDag.FileName = ReadINI(sSectionName, "LastSaveName")
    End If
    oComDag.InitDir = ReadINI(sSectionName, "LastSaveDir")
    
    On Error GoTo canceled:
    oComDag.ShowSave
    If oComDag.FileName = "" Then GoTo canceled:
    
    sFile = oComDag.FileName
    sFileTitle = oComDag.FileTitle
Else
    sFile = strFileName
    If Not sFile = "" Then
        For x = 1 To Len(sFile)
            If InStr(x, sFile, "\") > 0 Then
                y = InStr(x, sFile, "\") + 1
                x = y
            End If
        Next x
        If y = 0 Then y = 1
        sFileTitle = Mid(sFile, y)
    End If
End If

On Error GoTo Error:

If Not sFile = "" Then
    If Not LCase(Right(sFile, 5)) = ".mmec" Then sFile = sFile & ".mmec"
    
    bCharLoaded = True
    
    'need to create the file before adding to recent
    Call WriteINI("PlayerInfo", "Name", txtCharName.Text, sFile)
    Call RecentFileAdd(sFile)
    
    'these dont get ", sFile" but the others do
    Call WriteINI(sSectionName, "LastSaveDir", _
        Left(sFile, Len(sFile) - Len(sFileTitle) - 1))
    Call WriteINI(sSectionName, "LastSaveName", sFileTitle)
    Call WriteINI(sSectionName, "LastCharFile", sFile) '<-- (sFile is needed here)
    
    'Me.Caption = sNormalCaption & " (" & sFileTitle & ")"
End If

sName = txtCharName.Text
If bNameInTitle And bCharLoaded Then
    If Len(sName) > 1 Then
        Me.Caption = sNormalCaption & " (" & sName & ")"
    Else
        Me.Caption = sNormalCaption & " (" & sFileTitle & ")"
    End If
Else
    If bCharLoaded Then
        Me.Caption = sNormalCaption & " (" & sFileTitle & ")"
    Else
        Me.Caption = sNormalCaption
    End If
End If

If Not sFile = "" Then sSectionName = "PlayerInfo"
'Call WriteINI(sSectionName, "UseGlobalFilter", chkGlobalFilter.Value, sFile)
Call WriteINI(sSectionName, "Class", cmbGlobalClass(0).ItemData(cmbGlobalClass(0).ListIndex), sFile)
Call WriteINI(sSectionName, "Race", cmbGlobalRace(0).ItemData(cmbGlobalRace(0).ListIndex), sFile)
Call WriteINI(sSectionName, "Level", txtGlobalLevel(0).Text, sFile)
Call WriteINI(sSectionName, "Alignment", cmbGlobalAlignment.ListIndex, sFile)

Call WriteINI(sSectionName, "Name", txtCharName.Text, sFile)
Call WriteINI(sSectionName, "Strength", txtCharStats(0).Text, sFile)
Call WriteINI(sSectionName, "Intellect", txtCharStats(1).Text, sFile)
Call WriteINI(sSectionName, "Widsom", txtCharStats(2).Text, sFile)
Call WriteINI(sSectionName, "Agility", txtCharStats(3).Text, sFile)
Call WriteINI(sSectionName, "Health", txtCharStats(4).Text, sFile)
Call WriteINI(sSectionName, "Charm", txtCharStats(5).Text, sFile)

Call WriteINI(sSectionName, "AddWeight", txtInvenAddWeight.Text, sFile)
Call WriteINI(sSectionName, "UseAddWeight", chkInvenAddWeight.Value, sFile)

For x = 0 To 4
    Call WriteINI(sSectionName, "Quest" & x, chkCharQuests(x).Value, sFile)
Next x
Call WriteINI(sSectionName, "Quest_2nd", cmbChar2ndAlign.ListIndex, sFile)

On Error Resume Next
If Not sFile = "" Then sSectionName = "Inventory"
Call WriteINI(sSectionName, "Head", nEquippedItem(0), sFile)
Call WriteINI(sSectionName, "Ears", nEquippedItem(1), sFile)
Call WriteINI(sSectionName, "Neck", nEquippedItem(2), sFile)
Call WriteINI(sSectionName, "Back", nEquippedItem(3), sFile)
Call WriteINI(sSectionName, "Torso", nEquippedItem(4), sFile)
Call WriteINI(sSectionName, "Arms", nEquippedItem(5), sFile)
Call WriteINI(sSectionName, "Wrist", nEquippedItem(6), sFile)
Call WriteINI(sSectionName, "Wrist2", nEquippedItem(7), sFile)
Call WriteINI(sSectionName, "Hands", nEquippedItem(8), sFile)
Call WriteINI(sSectionName, "Finger1", nEquippedItem(9), sFile)
Call WriteINI(sSectionName, "Finger2", nEquippedItem(10), sFile)
Call WriteINI(sSectionName, "Waist", nEquippedItem(11), sFile)
Call WriteINI(sSectionName, "Legs", nEquippedItem(12), sFile)
Call WriteINI(sSectionName, "Feet", nEquippedItem(13), sFile)
Call WriteINI(sSectionName, "Worn", nEquippedItem(14), sFile)
Call WriteINI(sSectionName, "Off-Hand", nEquippedItem(15), sFile)
Call WriteINI(sSectionName, "Weapon", nEquippedItem(16), sFile)
Call WriteINI(sSectionName, "Eyes", nEquippedItem(17), sFile)
Call WriteINI(sSectionName, "Face", nEquippedItem(18), sFile)

If Not sFile = "" Then sSectionName = "Bless"
For x = 0 To 9
    Call WriteINI(sSectionName, "Bless" & x, cmbCharBless(x).ItemData(cmbCharBless(x).ListIndex), sFile)
Next x

If Not sFile = "" Then sSectionName = "Compare"

x = 1
str = ""
For Each oLI In lvWeaponCompare.ListItems
    str = str & oLI.Text & ","
    Set oLI = Nothing
    x = x + 1
    If x > 101 Then Exit For
Next
Call WriteINI(sSectionName, "WeaponCompare", str, sFile)

x = 1
str = ""
For Each oLI In lvArmourCompare.ListItems
    str = str & oLI.Text & ","
    Set oLI = Nothing
    x = x + 1
    If x > 101 Then Exit For
Next
Call WriteINI(sSectionName, "ArmourCompare", str, sFile)

x = 1
str = ""
For Each oLI In lvSpellCompare.ListItems
    str = str & oLI.Text & ","
    Set oLI = Nothing
    x = x + 1
    If x > 101 Then Exit For
Next
Call WriteINI(sSectionName, "SpellCompare", str, sFile)

bPromptSave = False

GoTo out:

canceled:
SaveCharacter = -1

out:
On Error Resume Next
Set oLI = Nothing
Exit Function
Error:
Call HandleError("SaveCharacter")
Resume out:
End Function

Public Function SaveSettings() As Integer
Dim sSectionName As String, nYesNo As Integer, nResult As Integer, x As Integer
On Error GoTo Error:

sSectionName = RemoveCharacter(lblDatVer.Caption, " ")

If cmbGlobalClass(0).ListIndex < 0 Then Exit Function

If bCharLoaded Then
    If bAutoSave Then
        If FileExists(ReadINI(sSectionName, "LastCharFile")) Then
            nResult = SaveCharacter(False, ReadINI(sSectionName, "LastCharFile"))
            If nResult = -1 Then
                SaveSettings = -1
                Exit Function
            End If
        Else
            If bPromptSave = True Then Call SaveCharacter(True)
        End If
    Else
        If bPromptSave = True Then
            nYesNo = MsgBox("Save Character File first?", vbYesNoCancel + vbQuestion + vbDefaultButton3)
            If nYesNo = vbYes Then
                If FileExists(ReadINI(sSectionName, "LastCharFile")) Then
                    nResult = SaveCharacter(False, ReadINI(sSectionName, "LastCharFile"))
                Else
                    nResult = SaveCharacter(True)
                End If
                If nResult = -1 Then
                    SaveSettings = -1
                    Exit Function
                End If
            ElseIf nYesNo = vbCancel Then
                SaveSettings = -1
                Exit Function
            End If
        End If
    End If
Else
    Call SaveCharacter(False)
End If

If Me.WindowState = vbMaximized Then
    Call WriteINI("Settings", "Maximized", 1)
Else
    Call WriteINI("Settings", "Top", Me.Top)
    Call WriteINI("Settings", "Left", Me.Left)
    Call WriteINI("Settings", "Width", Me.Width)
    Call WriteINI("Settings", "Height", Me.Height)
    Call WriteINI("Settings", "Maximized", 0)
End If

Call WriteINI("Settings", "WeaponSplitNS", splSplitterNS(0).Position)
Call WriteINI("Settings", "WeaponSplitWE", splSplitterWE(0).Position)
'Call WriteINI("Settings", "ArmourSplitNS", splSplitterNS(1).Position)
'Call WriteINI("Settings", "ArmourSplitWE", splSplitterWE(1).Position)
'Call WriteINI("Settings", "SpellSplitNS", splSplitterNS(2).Position)
'Call WriteINI("Settings", "SpellSplitWE", splSplitterWE(2).Position)
Call WriteINI("Settings", "WeaponCompSplitNS", splSplitterNS(3).Position)
Call WriteINI("Settings", "WeaponCompSplitWE", splSplitterWE(3).Position)
'Call WriteINI("Settings", "ArmourCompSplitNS", splSplitterNS(4).Position)
'Call WriteINI("Settings", "ArmourCompSplitWE", splSplitterWE(4).Position)
'Call WriteINI("Settings", "SpellCompSplitNS", splSplitterNS(5).Position)
'Call WriteINI("Settings", "SpellCompSplitWE", splSplitterWE(5).Position)
Call WriteINI("Settings", "OtherSplitNS", splSplitterNS(6).Position)
Call WriteINI("Settings", "OtherSplitWE", splSplitterWE(6).Position)
Call WriteINI("Settings", "MonsterSplitWE", splMonsterSplit(0).Position)
Call WriteINI("Settings", "MonsterCompSplitWE", splMonsterSplit(1).Position)

Call WriteINI("Settings", "MapFollowMap", chkMapOptions(0).Value)
Call WriteINI("Settings", "MapNoHidden", chkMapOptions(1).Value)
Call WriteINI("Settings", "MapNoLairs", chkMapOptions(2).Value)
Call WriteINI("Settings", "MapNoNPC", chkMapOptions(3).Value)
Call WriteINI("Settings", "MapNoCMD", chkMapOptions(4).Value)
Call WriteINI("Settings", "MapNoTooltips", chkMapOptions(5).Value)

Call WriteINI("Settings", "UseGlobalFilter", chkGlobalFilter.Value)
Call WriteINI("Setting", "LookUpMonsterRegen", chkMonstersNoRegenLookUp.Value)

For x = 1 To 5
    Call WriteINI("Settings", "Recent" & x, sRecentFiles(x, 2))
Next x

If mnuJumpToCompare.Checked = True Then
    Call WriteINI("Settings", "JumpToCompare", 1)
Else
    Call WriteINI("Settings", "JumpToCompare", 0)
End If

Exit Function
Error:
Call HandleError("SaveSettings")
Resume Next
End Function

Private Sub SetLabelFonts(sName As String, nSize As Integer, bBold As Boolean, bItalic As Boolean)
On Error GoTo Error:
Dim x As Integer

Call WriteINI("Settings", "LabelFontName", sName)
Call WriteINI("Settings", "LabelFontSize", nSize)
Call WriteINI("Settings", "LabelFontBold", bBold)
Call WriteINI("Settings", "LabelFontItal", bItalic)

For x = 2 To 29
    lblInvenStats(x).Font.name = sName
    lblInvenStats(x).Font.Size = nSize
    lblInvenStats(x).Font.Bold = bBold
    lblInvenStats(x).Font.Italic = bItalic
Next x

DoEvents
Exit Sub
Error:
HandleError
End Sub

Private Sub SetStatFonts(sName As String, nSize As Integer, bBold As Boolean, bItalic As Boolean)
On Error GoTo Error:
Dim x As Integer

Call WriteINI("Settings", "StatFontName", sName)
Call WriteINI("Settings", "StatFontSize", nSize)
Call WriteINI("Settings", "StatFontBold", bBold)
Call WriteINI("Settings", "StatFontItal", bItalic)

For x = 0 To 29
    txtStat(x).Font.name = sName
    txtStat(x).Font.Size = nSize
    txtStat(x).Font.Bold = bBold
    txtStat(x).Font.Italic = bItalic
Next x

For x = 0 To 1
    lblInvenSlash(x).Font.name = sName
    lblInvenSlash(x).Font.Size = nSize + 1
    lblInvenSlash(x).Font.Bold = bBold
    lblInvenSlash(x).Font.Italic = bItalic
Next x

lblEncumLevel(1).Font.name = sName
lblEncumLevel(1).Font.Size = nSize
lblEncumLevel(1).Font.Bold = bBold
lblEncumLevel(1).Font.Italic = bItalic

DoEvents
Exit Sub
Error:
HandleError
End Sub

Private Sub SetupCharBless()
Dim x As Integer ', y As Integer, z As Integer

On Error GoTo Error:

For x = 0 To cmbCharBless().UBound
    cmbCharBless(x).clear
    cmbCharBless(x).AddItem "(none)", 0
    cmbCharBless(x).ItemData(cmbCharBless(x).NewIndex) = 0
    cmbCharBless(x).ListIndex = 0
Next x

Exit Sub

Error:
Call HandleError("SetupCharBless")
End Sub

Private Sub SetupClass()
Dim x As Integer, y As Integer
Dim bAntiMagic As Boolean, bHasPicks As Boolean
Dim nMagicLVL As Integer, nMagery As Integer, bHasMagic As Boolean
On Error GoTo Error:

If cmbGlobalClass(0).ItemData(cmbGlobalClass(0).ListIndex) = 0 Or cmbGlobalClass(0).Text = "(none)" Then
    Call ResetFilterOptions(, , , True)
    GoTo noclass:
End If

If tabClasses.RecordCount = 0 Then GoTo noclass:

If chkCharQuests(5).Value = 0 Or cmbChar2ndAlign.ListIndex = 0 Then
    Select Case cmbGlobalClass(0).ItemData(cmbGlobalClass(0).ListIndex)
        Case 1, 2, 3, 15: cmbChar2ndAlign.ListIndex = 1
        Case 4, 11: cmbChar2ndAlign.ListIndex = 2
        Case 5, 12, 13: cmbChar2ndAlign.ListIndex = 3
        Case 6, 9, 10: cmbChar2ndAlign.ListIndex = 4
        Case 7, 8, 14: cmbChar2ndAlign.ListIndex = 5
    End Select
End If

tabClasses.Index = "pkClasses"
tabClasses.Seek "=", cmbGlobalClass(0).ItemData(cmbGlobalClass(0).ListIndex)
If tabClasses.NoMatch = False Then GoTo checkclass:

noclass:
If chkCharQuests(5).Value = 0 Then cmbChar2ndAlign.ListIndex = 0
bHasPicks = True
bAntiMagic = False
nMagicLVL = 3
nMagery = 0
bHasMagic = True
GoTo setup:

checkclass:
'Select Case tabClasses.Fields("ArmourType")
'    Case 0: y = 0 '"Natural"
'    Case 1: y = 1 '"Silk"      '"Robes"
'    Case 2: y = 2 '"Ninja"     '"Padded"
'    Case 3, 4, 5, 6: y = 3 '"Leather" '"Soft Leather","Soft Studded","Rigid Leather","Rigid Studded"
'    Case 7: y = 4 '"Chainmail"
'    Case 8: y = 5 '"Scalemail"
'    Case 9: y = 6 '"Platemail"
'    Case Else: y = 6
'End Select
'
'For x = 0 To 6
'    chkArmourType(x).Value = 0
'Next
'For x = 0 To y
'    chkArmourType(x).Value = 1
'Next
For x = 0 To 6
    chkArmourType(x).Value = 1
Next

'chkWeaponStaffOnly.Value = 0
For x = 0 To 3
    chkHanded(x).Value = 1
Next
If tabClasses.Fields("WeaponType") = 9 Then
    chkWeaponStaffOnly.Value = 1
Else
    chkWeaponStaffOnly.Value = 0
End If

'Select Case tabClasses.Fields("WeaponType")
'    Case 0: '"1H Blunt"
'        chkHanded(0).Value = 1
'    Case 1: '"2H Blunt"
'        chkHanded(1).Value = 1
'    Case 2: '"1H Sharp"
'        chkHanded(2).Value = 1
'    Case 3: '"2H Sharp"
'        chkHanded(3).Value = 1
'    Case 4: '"Any 1H"
'        chkHanded(0).Value = 1
'        chkHanded(2).Value = 1
'    Case 5: '"Any 2H"
'        chkHanded(1).Value = 1
'        chkHanded(3).Value = 1
'    Case 6: '"Any Sharp"
'        chkHanded(2).Value = 1
'        chkHanded(3).Value = 1
'    Case 7: '"Any Blunt"
'        chkHanded(0).Value = 1
'        chkHanded(1).Value = 1
'    Case 8: '"All Weapons"
'        For x = 0 To 3
'            chkHanded(x).Value = 1
'        Next
'    Case 9: '"Staff"
'        chkWeaponStaffOnly.Value = 1
'    Case Else:
'        For x = 0 To 3
'            chkHanded(x).Value = 1
'        Next
'End Select

For x = 0 To 9
    Select Case tabClasses.Fields("Abil-" & x)
        Case 0:
        Case 51: 'anti-magic check
            bAntiMagic = True
        Case 37: 'picklocks
            bHasPicks = True
    End Select
Next x

nMagicLVL = tabClasses.Fields("MageryLVL")
If nMagicLVL > 3 Then nMagicLVL = 3
nMagery = tabClasses.Fields("MageryType")
If tabClasses.Fields("MageryType") > 0 Then bHasMagic = True

setup:
cmbSpellMageryLevel.ListIndex = nMagicLVL
cmbSpellMagery.ListIndex = nMagery

If bHasPicks Then
    fraChar(3).Enabled = True
    lblCharPicklocks.Enabled = True
Else
    fraChar(3).Enabled = False
    lblCharPicklocks.Enabled = False
End If

If bAntiMagic Then
    chkArmourNonMagic.Value = 1 'anti-magic check
    chkWeaponNonMagical.Value = 1
Else
    chkArmourNonMagic.Value = 0
    chkWeaponNonMagical.Value = 0
End If

If bHasMagic Then
    chkSpellLearnable.Value = 1
    fraChar(2).Enabled = True
    lblCharSC.Enabled = True
    lblCharMaxMana.Enabled = True
    lblCharManaRate.Enabled = True
Else
    chkSpellLearnable.Value = 0
    fraChar(2).Enabled = False
    lblCharSC.Enabled = False
    lblCharMaxMana.Enabled = False
    lblCharManaRate.Enabled = False
End If

Exit Sub

Error:
Call HandleError("SetupClass")

End Sub

Private Sub SetUpFormObjects()
On Error Resume Next
Dim x As Integer, strTemp As String

nLastShopSort = 2
nLastWeaponSort = 2
nLastSpellSort = 2
nLastArmourSort = 2
nLastShopDetailIndex = 1
ReDim nInvenExcludedItems(0)

cmbChar2ndAlign.clear
cmbChar2ndAlign.AddItem "2nd Alignment Quest ...", 0
cmbChar2ndAlign.AddItem "+1 Max Damage", 1
cmbChar2ndAlign.AddItem "+1 AC, +6 Mana", 2
cmbChar2ndAlign.AddItem "+1 SC, +10 Mana", 3
cmbChar2ndAlign.AddItem "+4 Mana, +6 BS Min/Max, +1 Stealth", 4
cmbChar2ndAlign.AddItem "+10 BS Min/Max, +2 Stealth", 5
Call AutoSizeDropDownWidth(cmbChar2ndAlign)
cmbChar2ndAlign.ListIndex = 0

cmbArmourWorn.clear
cmbArmourWorn.AddItem "Nowhere"
cmbArmourWorn.ItemData(cmbArmourWorn.NewIndex) = 1
cmbArmourWorn.AddItem "Head"
cmbArmourWorn.ItemData(cmbArmourWorn.NewIndex) = 2
cmbArmourWorn.AddItem "Hands"
cmbArmourWorn.ItemData(cmbArmourWorn.NewIndex) = 3
cmbArmourWorn.AddItem "Finger"
cmbArmourWorn.ItemData(cmbArmourWorn.NewIndex) = 4
cmbArmourWorn.AddItem "Feet"
cmbArmourWorn.ItemData(cmbArmourWorn.NewIndex) = 5
cmbArmourWorn.AddItem "Arms"
cmbArmourWorn.ItemData(cmbArmourWorn.NewIndex) = 6
cmbArmourWorn.AddItem "Back"
cmbArmourWorn.ItemData(cmbArmourWorn.NewIndex) = 7
cmbArmourWorn.AddItem "Neck"
cmbArmourWorn.ItemData(cmbArmourWorn.NewIndex) = 8
cmbArmourWorn.AddItem "Legs"
cmbArmourWorn.ItemData(cmbArmourWorn.NewIndex) = 9
cmbArmourWorn.AddItem "Waist"
cmbArmourWorn.ItemData(cmbArmourWorn.NewIndex) = 10
cmbArmourWorn.AddItem "Torso"
cmbArmourWorn.ItemData(cmbArmourWorn.NewIndex) = 11
cmbArmourWorn.AddItem "Off-Hand"
cmbArmourWorn.ItemData(cmbArmourWorn.NewIndex) = 12
cmbArmourWorn.AddItem "Wrist"
cmbArmourWorn.ItemData(cmbArmourWorn.NewIndex) = 13
cmbArmourWorn.AddItem "Ears"
cmbArmourWorn.ItemData(cmbArmourWorn.NewIndex) = 14
cmbArmourWorn.AddItem "Worn"
cmbArmourWorn.ItemData(cmbArmourWorn.NewIndex) = 15
cmbArmourWorn.AddItem "Eyes"
cmbArmourWorn.ItemData(cmbArmourWorn.NewIndex) = 16
cmbArmourWorn.AddItem "Face"
cmbArmourWorn.ItemData(cmbArmourWorn.NewIndex) = 17
cmbArmourWorn.AddItem "Anywhere", 0
cmbArmourWorn.ListIndex = 0
Call ExpandCombo(cmbArmourWorn, HeightOnly, TripleWidth, framNav(1).hwnd)

cmbWeaponMagicLevel.clear
cmbWeaponMagicLevel.AddItem "Any", 0
cmbWeaponMagicLevel.AddItem "1", 1
cmbWeaponMagicLevel.AddItem "2", 2
cmbWeaponMagicLevel.AddItem "3", 3
cmbWeaponMagicLevel.AddItem "4", 4
cmbWeaponMagicLevel.AddItem "5", 5
cmbWeaponMagicLevel.AddItem "6+", 6
cmbWeaponMagicLevel.ListIndex = 0

cmbSpellMagery.clear
cmbSpellMagery.AddItem "Any", 0
cmbSpellMagery.AddItem "Mage", 1
cmbSpellMagery.AddItem "Priest", 2
cmbSpellMagery.AddItem "Druid", 3
cmbSpellMagery.AddItem "Bard", 4
cmbSpellMagery.AddItem "Kai", 5
cmbSpellMagery.ListIndex = 0

cmbSpellMageryLevel.clear
cmbSpellMageryLevel.AddItem "Any", 0
cmbSpellMageryLevel.AddItem "1", 1
cmbSpellMageryLevel.AddItem "2", 2
cmbSpellMageryLevel.AddItem "3", 3
cmbSpellMageryLevel.ListIndex = 0

cmbSpellTarget.clear
cmbSpellTarget.AddItem "Any", 0
cmbSpellTarget.AddItem "Self", 1
cmbSpellTarget.AddItem "User", 2
cmbSpellTarget.AddItem "Monster", 3
cmbSpellTarget.AddItem "Party", 4
cmbSpellTarget.AddItem "Room", 5
cmbSpellTarget.ListIndex = 0

cmbSpellAttackType.clear
cmbSpellAttackType.AddItem "Any", 0
cmbSpellAttackType.AddItem "Cold", 1
cmbSpellAttackType.AddItem "Hot", 2
cmbSpellAttackType.AddItem "Stone", 3
cmbSpellAttackType.AddItem "Lightning", 4
cmbSpellAttackType.AddItem "Normal", 5
cmbSpellAttackType.AddItem "Water", 6
cmbSpellAttackType.AddItem "Poison", 7
cmbSpellAttackType.ListIndex = 0

cmbContainsAbil(0).clear
For x = 1 To 187
    strTemp = GetAbilityName(x)
    If Not strTemp = "" Then
        cmbContainsAbil(0).AddItem strTemp
        cmbContainsAbil(0).ItemData(cmbContainsAbil(0).NewIndex) = x
    End If
Next x
cmbContainsAbil(0).AddItem "Any", 0
Call AutoSizeDropDownWidth(cmbContainsAbil(0))
Call ExpandCombo(cmbContainsAbil(0), HeightOnly, DoubleWidth, framNav(2).hwnd)
cmbContainsAbil(0).ListIndex = 0

cmbGlobalAlignment.clear
cmbGlobalAlignment.AddItem "Any"
cmbGlobalAlignment.AddItem "Good"
cmbGlobalAlignment.AddItem "Neutral"
cmbGlobalAlignment.AddItem "Evil"
cmbGlobalAlignment.ListIndex = 0

lvWeapons.ColumnHeaders.clear
lvWeapons.ColumnHeaders.Add 1, "Number", "#", 600, lvwColumnLeft
lvWeapons.ColumnHeaders.Add 2, "Name", "Name", 2000, lvwColumnCenter
lvWeapons.ColumnHeaders.Add 3, "Type", "Type", 1000, lvwColumnCenter
lvWeapons.ColumnHeaders.Add 4, "Min Dmg", "Min Dmg", 850, lvwColumnCenter
lvWeapons.ColumnHeaders.Add 5, "Max Dmg", "Max Dmg", 850, lvwColumnCenter
lvWeapons.ColumnHeaders.Add 6, "Speed", "Speed", 700, lvwColumnCenter
lvWeapons.ColumnHeaders.Add 7, "LVL", "LVL", 500, lvwColumnCenter
lvWeapons.ColumnHeaders.Add 8, "Str", "Str", 500, lvwColumnCenter
lvWeapons.ColumnHeaders.Add 9, "Enc", "Enc", 600, lvwColumnCenter
lvWeapons.ColumnHeaders.Add 10, "AC", "AC", 650, lvwColumnCenter
lvWeapons.ColumnHeaders.Add 11, "Acc", "Acc", 500, lvwColumnCenter
lvWeapons.ColumnHeaders.Add 12, "BS", "BS", 500, lvwColumnCenter
lvWeapons.ColumnHeaders.Add 13, "Limit", "Limit", 500, lvwColumnCenter
lvWeapons.ColumnHeaders.Add 14, "Dmg/Spd", "Dmg/Spd", 900, lvwColumnCenter

lvWeaponCompare.ColumnHeaders.clear
lvWeaponCompare.ColumnHeaders.Add 1, "Number", "#", 550, lvwColumnLeft
lvWeaponCompare.ColumnHeaders.Add 2, "Name", "Name", 1900, lvwColumnCenter
lvWeaponCompare.ColumnHeaders.Add 3, "Type", "Type", 1000, lvwColumnCenter
lvWeaponCompare.ColumnHeaders.Add 4, "Min Dmg", "Min Dmg", 1000, lvwColumnCenter
lvWeaponCompare.ColumnHeaders.Add 5, "Max Dmg", "Max Dmg", 1000, lvwColumnCenter
lvWeaponCompare.ColumnHeaders.Add 6, "Speed", "Speed", 700, lvwColumnCenter
lvWeaponCompare.ColumnHeaders.Add 7, "LVL", "LVL", 500, lvwColumnCenter
lvWeaponCompare.ColumnHeaders.Add 8, "Str", "Str", 500, lvwColumnCenter
lvWeaponCompare.ColumnHeaders.Add 9, "Enc", "Enc", 600, lvwColumnCenter
lvWeaponCompare.ColumnHeaders.Add 10, "AC", "AC", 650, lvwColumnCenter
lvWeaponCompare.ColumnHeaders.Add 11, "Acc", "Acc", 500, lvwColumnCenter
lvWeaponCompare.ColumnHeaders.Add 12, "BS", "BS", 500, lvwColumnCenter
lvWeaponCompare.ColumnHeaders.Add 13, "Limit", "Limit", 500, lvwColumnCenter
lvWeaponCompare.ColumnHeaders.Add 14, "Dmg/Spd", "Dmg/Spd", 900, lvwColumnCenter

lvArmour.ColumnHeaders.clear
lvArmour.ColumnHeaders.Add 1, "Number", "#", 600, lvwColumnLeft
lvArmour.ColumnHeaders.Add 2, "Name", "Name", 2500, lvwColumnCenter
lvArmour.ColumnHeaders.Add 3, "Worn", "Worn", 1000, lvwColumnCenter
lvArmour.ColumnHeaders.Add 4, "Armr Type", "Armr Type", 1500, lvwColumnCenter
lvArmour.ColumnHeaders.Add 5, "Level", "Level", 600, lvwColumnCenter
lvArmour.ColumnHeaders.Add 6, "Enc", "Enc", 800, lvwColumnCenter
lvArmour.ColumnHeaders.Add 7, "AC", "AC/DR", 900, lvwColumnCenter
lvArmour.ColumnHeaders.Add 8, "Acc", "Acc", 500, lvwColumnCenter
lvArmour.ColumnHeaders.Add 9, "Limit", "Limit", 500, lvwColumnCenter
lvArmour.ColumnHeaders.Add 10, "Enc/AC", "AC/Enc", 850, lvwColumnCenter

lvArmourCompare.ColumnHeaders.clear
lvArmourCompare.ColumnHeaders.Add 1, "Number", "#", 600, lvwColumnLeft
lvArmourCompare.ColumnHeaders.Add 2, "Name", "Name", 2300, lvwColumnCenter
lvArmourCompare.ColumnHeaders.Add 3, "Worn", "Worn", 1000, lvwColumnCenter
lvArmourCompare.ColumnHeaders.Add 4, "Armr Type", "Armr Type", 1500, lvwColumnCenter
lvArmourCompare.ColumnHeaders.Add 5, "Level", "Level", 600, lvwColumnCenter
lvArmourCompare.ColumnHeaders.Add 6, "Enc", "Enc", 800, lvwColumnCenter
lvArmourCompare.ColumnHeaders.Add 7, "AC/DR", "AC", 900, lvwColumnCenter
lvArmourCompare.ColumnHeaders.Add 8, "Acc", "Acc", 500, lvwColumnCenter
lvArmourCompare.ColumnHeaders.Add 9, "Limit", "Limit", 500, lvwColumnCenter
lvArmourCompare.ColumnHeaders.Add 10, "Enc/AC", "AC/Enc", 850, lvwColumnCenter

lvOtherItems.ColumnHeaders.clear
lvOtherItems.ColumnHeaders.Add 1, "Number", "Number", 750, lvwColumnLeft
lvOtherItems.ColumnHeaders.Add 2, "Name", "Name", 2500, lvwColumnCenter
lvOtherItems.ColumnHeaders.Add 3, "Type", "Type", 1000, lvwColumnCenter
lvOtherItems.ColumnHeaders.Add 4, "Enc", "Enc", 800, lvwColumnCenter
lvOtherItems.ColumnHeaders.Add 5, "Limit", "Limit", 500, lvwColumnCenter
'lvOtherItems.ColumnHeaders.Add 6, "Abilities", "Abilities", 3700, lvwColumnLeft

lvClasses.ColumnHeaders.clear
lvClasses.ColumnHeaders.Add 1, "Number", "#", 400, lvwColumnLeft
lvClasses.ColumnHeaders.Add 2, "Name", "Name", 1300, lvwColumnCenter
lvClasses.ColumnHeaders.Add 3, "Exp%", "Exp", 700, lvwColumnCenter
lvClasses.ColumnHeaders.Add 4, "Weapon", "Weapon", 1400, lvwColumnCenter
lvClasses.ColumnHeaders.Add 5, "Armour", "Armour", 1000, lvwColumnCenter
lvClasses.ColumnHeaders.Add 6, "Magic", "Magic", 900, lvwColumnCenter
lvClasses.ColumnHeaders.Add 7, "Cmbt", "Cmbt", 600, lvwColumnCenter
lvClasses.ColumnHeaders.Add 8, "HP", "HP", 600, lvwColumnCenter
lvClasses.ColumnHeaders.Add 9, "Abilities", "Abilities", 3000, lvwColumnLeft

lvRaces.ColumnHeaders.clear
lvRaces.ColumnHeaders.Add 1, "Number", "#", 400, lvwColumnLeft
lvRaces.ColumnHeaders.Add 2, "Name", "Name", 1300, lvwColumnCenter
lvRaces.ColumnHeaders.Add 3, "Exp%", "Exp", 700, lvwColumnCenter
lvRaces.ColumnHeaders.Add 4, "HP", "HP", 500, lvwColumnCenter
lvRaces.ColumnHeaders.Add 5, "Str", "Str", 850, lvwColumnCenter
lvRaces.ColumnHeaders.Add 6, "Int", "Int", 850, lvwColumnCenter
lvRaces.ColumnHeaders.Add 7, "Wis", "Wis", 850, lvwColumnCenter
lvRaces.ColumnHeaders.Add 8, "Agi", "Agi", 850, lvwColumnCenter
lvRaces.ColumnHeaders.Add 9, "Hea", "Hea", 850, lvwColumnCenter
lvRaces.ColumnHeaders.Add 10, "Cha", "Cha", 850, lvwColumnCenter
lvRaces.ColumnHeaders.Add 11, "Abilities", "Abilities", 2000, lvwColumnLeft

lvSpells.ColumnHeaders.clear
lvSpells.ColumnHeaders.Add 1, "Number", "#", 600, lvwColumnLeft
lvSpells.ColumnHeaders.Add 2, "Name", "Name", 2000, lvwColumnCenter
lvSpells.ColumnHeaders.Add 3, "Short", "Short", 650, lvwColumnCenter
lvSpells.ColumnHeaders.Add 4, "Magery", "Magery", 900, lvwColumnCenter
lvSpells.ColumnHeaders.Add 5, "LVL", "LVL", 500, lvwColumnCenter
lvSpells.ColumnHeaders.Add 6, "Mana", "Mana", 650, lvwColumnCenter
lvSpells.ColumnHeaders.Add 7, "Diff", "Diff", 500, lvwColumnCenter
lvSpells.ColumnHeaders.Add 8, "Detail", "Detail", 4000, lvwColumnLeft

lvSpellCompare.ColumnHeaders.clear
lvSpellCompare.ColumnHeaders.Add 1, "Number", "#", 600, lvwColumnLeft
lvSpellCompare.ColumnHeaders.Add 2, "Name", "Name", 2000, lvwColumnCenter
lvSpellCompare.ColumnHeaders.Add 3, "Short", "Short", 650, lvwColumnCenter
lvSpellCompare.ColumnHeaders.Add 4, "Magery", "Magery", 900, lvwColumnCenter
lvSpellCompare.ColumnHeaders.Add 5, "LVL", "LVL", 500, lvwColumnCenter
lvSpellCompare.ColumnHeaders.Add 6, "Mana", "Mana", 650, lvwColumnCenter
lvSpellCompare.ColumnHeaders.Add 7, "Diff", "Diff", 500, lvwColumnCenter
lvSpellCompare.ColumnHeaders.Add 8, "Detail", "Detail", 3900, lvwColumnLeft

lvMonsters.ColumnHeaders.clear
lvMonsters.ColumnHeaders.Add 1, "Number", "#", 500, lvwColumnLeft
lvMonsters.ColumnHeaders.Add 2, "Name", "Name", 1800, lvwColumnCenter
lvMonsters.ColumnHeaders.Add 3, "Exp", "Exp", 1100, lvwColumnCenter
lvMonsters.ColumnHeaders.Add 4, "HP", "HP", 750, lvwColumnCenter
lvMonsters.ColumnHeaders.Add 5, "Rgn", "Rgn", 550, lvwColumnCenter
lvMonsters.ColumnHeaders.Add 6, "Mgc", "Mgc", 550, lvwColumnCenter
lvMonsters.ColumnHeaders.Add 7, "Exp/HP", "Exp/HP", 1000, lvwColumnCenter

lvMonsterDetail.ColumnHeaders.clear
lvMonsterDetail.ColumnHeaders.Add 1, "Stat", "Stat", 1400, lvwColumnLeft
lvMonsterDetail.ColumnHeaders.Add 2, "Detail", "Detail  (double click items/spells/textblocks to jump)", 30000, lvwColumnLeft

lvMonsterCompare.ColumnHeaders.clear
lvMonsterCompare.ColumnHeaders.Add 1, "Number", "#", 500, lvwColumnLeft
lvMonsterCompare.ColumnHeaders.Add 2, "Name", "Name", 1800, lvwColumnCenter
lvMonsterCompare.ColumnHeaders.Add 3, "Exp", "Exp", 1100, lvwColumnCenter
lvMonsterCompare.ColumnHeaders.Add 4, "HP", "HP", 750, lvwColumnCenter
lvMonsterCompare.ColumnHeaders.Add 5, "Rgn", "Rgn", 550, lvwColumnCenter
lvMonsterCompare.ColumnHeaders.Add 6, "Mgc", "Mgc", 550, lvwColumnCenter
lvMonsterCompare.ColumnHeaders.Add 7, "Exp/HP", "Exp/HP", 1000, lvwColumnCenter

lvMonsterCompareLoc.ColumnHeaders.clear
lvMonsterCompareLoc.ColumnHeaders.Add 1, "Stat", "Stat", 1400, lvwColumnLeft
lvMonsterCompareLoc.ColumnHeaders.Add 2, "Detail", "Detail  (double click items/spells/textblocks to jump)", 30000, lvwColumnLeft

lvShops.ColumnHeaders.clear
lvShops.ColumnHeaders.Add 1, "Number", "#", 500, lvwColumnLeft
lvShops.ColumnHeaders.Add 2, "Name", "Name", 2300, lvwColumnCenter
lvShops.ColumnHeaders.Add 3, "Type", "Type", 1050, lvwColumnCenter

'lvShopDetail.ColumnHeaders.clear
'lvShopDetail.ColumnHeaders.Add 1, "Number", "#", 500, lvwColumnLeft
'lvShopDetail.ColumnHeaders.Add 2, "Name", "Name", 2000, lvwColumnCenter
'lvShopDetail.ColumnHeaders.Add 3, "Max", "Max", 550, lvwColumnCenter
'lvShopDetail.ColumnHeaders.Add 4, "Time", "Time", 800, lvwColumnCenter
'lvShopDetail.ColumnHeaders.Add 5, "Rgn%", "Rgn%", 600, lvwColumnCenter
'lvShopDetail.ColumnHeaders.Add 6, "Rgn#", "Rgn#", 600, lvwColumnCenter
'lvShopDetail.ColumnHeaders.Add 7, "Cost", "Cost", 1200, lvwColumnLeft

cmbWeaponSpeed.ListIndex = 0

lvMapLoc.ColumnHeaders.clear
lvMapLoc.ColumnHeaders.Add 1, "References", "References", 5000

lvShopLoc.ColumnHeaders.clear
lvShopLoc.ColumnHeaders.Add 1, "Assigned", "Assigned To", 3700

lvSpellLoc.ColumnHeaders.clear
lvSpellLoc.ColumnHeaders.Add 1, "References", "References", 3100

lvSpellCompareLoc.ColumnHeaders.clear
lvSpellCompareLoc.ColumnHeaders.Add 1, "References", "References", 3100

lvWeaponLoc.ColumnHeaders.clear
lvWeaponLoc.ColumnHeaders.Add 1, "References", "References", 3200 '"Location", "Obtained From", 3200

lvArmourLoc.ColumnHeaders.clear
lvArmourLoc.ColumnHeaders.Add 1, "References", "References", 3200 '"Location", "Obtained From", 3200

lvOtherItemLoc.ColumnHeaders.clear
lvOtherItemLoc.ColumnHeaders.Add 1, "References", "References", 3200 '"Location", "Obtained From", 3200

'compare
lvWeaponCompareLoc.ColumnHeaders.clear
lvWeaponCompareLoc.ColumnHeaders.Add 1, "References", "References", 3200

lvArmourCompareLoc.ColumnHeaders.clear
lvArmourCompareLoc.ColumnHeaders.Add 1, "References", "References", 3200

'lvSpellCompareCasted.ColumnHeaders.Clear
'lvSpellCompareCasted.ColumnHeaders.Add 1, "Casted", "Casted By", 3200

Call cmdCompareNav_Click(0)
Call InvenSetupEquip
Call SetupCharBless

End Sub

Private Sub SetupSplitters()
'Dim x As Long

On Error GoTo Error:

splSplitterWE(0).Orientation = cSPLTOrientationVertical
splSplitterWE(0).FullDrag = True
splSplitterWE(0).MinimumSize(cSPLTLeftOrTopPanel) = 30 'picSplitMain.ScaleX(picSplitLeft.Width, picSplitMain.ScaleMode, vbPixels)
splSplitterWE(0).MinimumSize(cSPLTRightOrBottomPanel) = 75
splSplitterWE(0).KeepProportion = True
splSplitterWE(0).SplitterSize = 7
splSplitterWE(0).Bind txtWeaponDetail, lvWeaponLoc

splSplitterNS(0).Orientation = cSPLTOrientationHorizontal
splSplitterNS(0).MinimumSize(cSPLTLeftOrTopPanel) = 30 'picSplitMain.ScaleX(picSplitLeft.Width, picSplitMain.ScaleMode, vbPixels)
splSplitterNS(0).MinimumSize(cSPLTRightOrBottomPanel) = 30
splSplitterNS(0).FullDrag = True
splSplitterNS(0).SplitterSize = 7
splSplitterNS(0).KeepProportion = True
splSplitterNS(0).Bind lvWeapons, splSplitterWE(0)

splSplitterWE(1).Orientation = cSPLTOrientationVertical
splSplitterWE(1).FullDrag = True
splSplitterWE(1).MinimumSize(cSPLTLeftOrTopPanel) = 30 'picSplitMain.ScaleX(picSplitLeft.Width, picSplitMain.ScaleMode, vbPixels)
splSplitterWE(1).MinimumSize(cSPLTRightOrBottomPanel) = 75
splSplitterWE(1).KeepProportion = True
splSplitterWE(1).SplitterSize = 7
splSplitterWE(1).Bind txtArmourDetail, lvArmourLoc

splSplitterNS(1).Orientation = cSPLTOrientationHorizontal
splSplitterNS(1).MinimumSize(cSPLTLeftOrTopPanel) = 30 'picSplitMain.ScaleX(picSplitLeft.Width, picSplitMain.ScaleMode, vbPixels)
splSplitterNS(1).MinimumSize(cSPLTRightOrBottomPanel) = 30
splSplitterNS(1).FullDrag = True
splSplitterNS(1).SplitterSize = 7
splSplitterNS(1).KeepProportion = True
splSplitterNS(1).Bind lvArmour, splSplitterWE(1)

splSplitterWE(2).Orientation = cSPLTOrientationVertical
splSplitterWE(2).FullDrag = True
splSplitterWE(2).MinimumSize(cSPLTLeftOrTopPanel) = 30 'picSplitMain.ScaleX(picSplitLeft.Width, picSplitMain.ScaleMode, vbPixels)
splSplitterWE(2).MinimumSize(cSPLTRightOrBottomPanel) = 75
splSplitterWE(2).KeepProportion = True
splSplitterWE(2).SplitterSize = 7
splSplitterWE(2).Bind txtSpellDetail, lvSpellLoc

splSplitterNS(2).Orientation = cSPLTOrientationHorizontal
splSplitterNS(2).MinimumSize(cSPLTLeftOrTopPanel) = 30 'picSplitMain.ScaleX(picSplitLeft.Width, picSplitMain.ScaleMode, vbPixels)
splSplitterNS(2).MinimumSize(cSPLTRightOrBottomPanel) = 30
splSplitterNS(2).FullDrag = True
splSplitterNS(2).SplitterSize = 7
splSplitterNS(2).KeepProportion = True
splSplitterNS(2).Bind lvSpells, splSplitterWE(2)

splSplitterWE(3).Orientation = cSPLTOrientationVertical
splSplitterWE(3).FullDrag = True
splSplitterWE(3).MinimumSize(cSPLTLeftOrTopPanel) = 30 'picSplitMain.ScaleX(picSplitLeft.Width, picSplitMain.ScaleMode, vbPixels)
splSplitterWE(3).MinimumSize(cSPLTRightOrBottomPanel) = 75
splSplitterWE(3).KeepProportion = True
splSplitterWE(3).SplitterSize = 7
splSplitterWE(3).Bind txtWeaponCompareDetail, lvWeaponCompareLoc

splSplitterNS(3).Orientation = cSPLTOrientationHorizontal
splSplitterNS(3).MinimumSize(cSPLTLeftOrTopPanel) = 30 'picSplitMain.ScaleX(picSplitLeft.Width, picSplitMain.ScaleMode, vbPixels)
splSplitterNS(3).MinimumSize(cSPLTRightOrBottomPanel) = 30
splSplitterNS(3).FullDrag = True
splSplitterNS(3).SplitterSize = 7
splSplitterNS(3).KeepProportion = True
splSplitterNS(3).Bind lvWeaponCompare, splSplitterWE(3)

splSplitterWE(4).Orientation = cSPLTOrientationVertical
splSplitterWE(4).FullDrag = True
splSplitterWE(4).MinimumSize(cSPLTLeftOrTopPanel) = 30 'picSplitMain.ScaleX(picSplitLeft.Width, picSplitMain.ScaleMode, vbPixels)
splSplitterWE(4).MinimumSize(cSPLTRightOrBottomPanel) = 75
splSplitterWE(4).KeepProportion = True
splSplitterWE(4).SplitterSize = 7
splSplitterWE(4).Bind txtArmourCompareDetail, lvArmourCompareLoc

splSplitterNS(4).Orientation = cSPLTOrientationHorizontal
splSplitterNS(4).MinimumSize(cSPLTLeftOrTopPanel) = 30 'picSplitMain.ScaleX(picSplitLeft.Width, picSplitMain.ScaleMode, vbPixels)
splSplitterNS(4).MinimumSize(cSPLTRightOrBottomPanel) = 30
splSplitterNS(4).FullDrag = True
splSplitterNS(4).SplitterSize = 7
splSplitterNS(4).KeepProportion = True
splSplitterNS(4).Bind lvArmourCompare, splSplitterWE(4)

splSplitterWE(5).Orientation = cSPLTOrientationVertical
splSplitterWE(5).FullDrag = True
splSplitterWE(5).MinimumSize(cSPLTLeftOrTopPanel) = 30 'picSplitMain.ScaleX(picSplitLeft.Width, picSplitMain.ScaleMode, vbPixels)
splSplitterWE(5).MinimumSize(cSPLTRightOrBottomPanel) = 75
splSplitterWE(5).KeepProportion = True
splSplitterWE(5).SplitterSize = 7
splSplitterWE(5).Bind txtSpellCompareDetail, lvSpellCompareLoc

splSplitterNS(5).Orientation = cSPLTOrientationHorizontal
splSplitterNS(5).MinimumSize(cSPLTLeftOrTopPanel) = 30 'picSplitMain.ScaleX(picSplitLeft.Width, picSplitMain.ScaleMode, vbPixels)
splSplitterNS(5).MinimumSize(cSPLTRightOrBottomPanel) = 30
splSplitterNS(5).FullDrag = True
splSplitterNS(5).SplitterSize = 7
splSplitterNS(5).KeepProportion = True
splSplitterNS(5).Bind lvSpellCompare, splSplitterWE(5)

splSplitterWE(6).Orientation = cSPLTOrientationVertical
splSplitterWE(6).FullDrag = True
splSplitterWE(6).MinimumSize(cSPLTLeftOrTopPanel) = 30 'picSplitMain.ScaleX(picSplitLeft.Width, picSplitMain.ScaleMode, vbPixels)
splSplitterWE(6).MinimumSize(cSPLTRightOrBottomPanel) = 75
splSplitterWE(6).KeepProportion = False
splSplitterWE(6).SplitterSize = 7
splSplitterWE(6).Bind lvOtherItems, splSplitterNS(6)

splSplitterNS(6).Orientation = cSPLTOrientationHorizontal
splSplitterNS(6).MinimumSize(cSPLTLeftOrTopPanel) = 30 'picSplitMain.ScaleX(picSplitLeft.Width, picSplitMain.ScaleMode, vbPixels)
splSplitterNS(6).MinimumSize(cSPLTRightOrBottomPanel) = 30
splSplitterNS(6).FullDrag = True
splSplitterNS(6).SplitterSize = 7
splSplitterNS(6).KeepProportion = True
splSplitterNS(6).Bind txtOtherItemDetail, lvOtherItemLoc

splMonsterSplit(0).Orientation = cSPLTOrientationVertical
splMonsterSplit(0).FullDrag = True
splMonsterSplit(0).MinimumSize(cSPLTLeftOrTopPanel) = 30 'picSplitMain.ScaleX(picSplitLeft.Width, picSplitMain.ScaleMode, vbPixels)
splMonsterSplit(0).MinimumSize(cSPLTRightOrBottomPanel) = 75
splMonsterSplit(0).KeepProportion = False
splMonsterSplit(0).SplitterSize = 7
splMonsterSplit(0).Bind lvMonsters, lvMonsterDetail

splMonsterSplit(1).Orientation = cSPLTOrientationVertical
splMonsterSplit(1).FullDrag = True
splMonsterSplit(1).MinimumSize(cSPLTLeftOrTopPanel) = 30 'picSplitMain.ScaleX(picSplitLeft.Width, picSplitMain.ScaleMode, vbPixels)
splMonsterSplit(1).MinimumSize(cSPLTRightOrBottomPanel) = 75
splMonsterSplit(1).KeepProportion = False
splMonsterSplit(1).SplitterSize = 7
splMonsterSplit(1).Bind lvMonsterCompare, lvMonsterCompareLoc

'x = Val(ReadINI("Settings", "WeaponSplitNS"))
'If x < 50 Then x = 250
'splSplitterNS(0).Position = x
'x = Val(ReadINI("Settings", "WeaponSplitWE"))
'If x < 50 Then x = 380
'splSplitterWE(0).Position = x
'
'x = Val(ReadINI("Settings", "ArmourSplitNS"))
'If x < 50 Then x = 250
'splSplitterNS(1).Position = x
'x = Val(ReadINI("Settings", "ArmourSplitWE"))
'If x < 50 Then x = 380
'splSplitterWE(1).Position = x
'
'x = Val(ReadINI("Settings", "SpellSplitNS"))
'If x < 50 Then x = 250
'splSplitterNS(2).Position = x
'x = Val(ReadINI("Settings", "SpellSplitWE"))
'If x < 50 Then x = 380
'splSplitterWE(2).Position = x
'
'x = Val(ReadINI("Settings", "WeaponCompSplitNS"))
'If x < 50 Then x = 240
'splSplitterNS(3).Position = x
'x = Val(ReadINI("Settings", "WeaponCompSplitWE"))
'If x < 50 Then x = 340
'splSplitterWE(3).Position = x
'
'x = Val(ReadINI("Settings", "ArmourCompSplitNS"))
'If x < 50 Then x = 240
'splSplitterNS(4).Position = x
'x = Val(ReadINI("Settings", "ArmourCompSplitWE"))
'If x < 50 Then x = 340
'splSplitterWE(4).Position = x
'
'x = Val(ReadINI("Settings", "SpellCompSplitNS"))
'If x < 50 Then x = 240
'splSplitterNS(5).Position = x
'x = Val(ReadINI("Settings", "SpellCompSplitWE"))
'If x < 50 Then x = 340
'splSplitterWE(5).Position = x
'
'x = Val(ReadINI("Settings", "OtherSplitNS"))
'If x < 50 Then x = 140
'splSplitterNS(6).Position = x
'x = Val(ReadINI("Settings", "OtherSplitWE"))
'If x < 50 Then x = 400
'splSplitterWE(6).Position = x


out:
Exit Sub
Error:
Call HandleError("SetupSplitters")
Resume out:

End Sub

Private Sub SetupSplitterSizes()
Dim x As Long

On Error GoTo Error:

x = Val(ReadINI("Settings", "WeaponSplitNS"))
If x < 50 Then x = 250
splSplitterNS(0).Position = x
x = Val(ReadINI("Settings", "WeaponSplitWE"))
If x < 50 Then x = 380
splSplitterWE(0).Position = x

'x = Val(ReadINI("Settings", "ArmourSplitNS"))
'If x < 50 Then x = 250
'splSplitterNS(1).Position = x
'x = Val(ReadINI("Settings", "ArmourSplitWE"))
'If x < 50 Then x = 380
'splSplitterWE(1).Position = x
'
'x = Val(ReadINI("Settings", "SpellSplitNS"))
'If x < 50 Then x = 250
'splSplitterNS(2).Position = x
'x = Val(ReadINI("Settings", "SpellSplitWE"))
'If x < 50 Then x = 380
'splSplitterWE(2).Position = x

x = Val(ReadINI("Settings", "WeaponCompSplitNS"))
If x < 50 Then x = 240
splSplitterNS(3).Position = x
x = Val(ReadINI("Settings", "WeaponCompSplitWE"))
If x < 50 Then x = 340
splSplitterWE(3).Position = x

'x = Val(ReadINI("Settings", "ArmourCompSplitNS"))
'If x < 50 Then x = 240
'splSplitterNS(4).Position = x
'x = Val(ReadINI("Settings", "ArmourCompSplitWE"))
'If x < 50 Then x = 340
'splSplitterWE(4).Position = x
'
'x = Val(ReadINI("Settings", "SpellCompSplitNS"))
'If x < 50 Then x = 240
'splSplitterNS(5).Position = x
'x = Val(ReadINI("Settings", "SpellCompSplitWE"))
'If x < 50 Then x = 340
'splSplitterWE(5).Position = x

x = Val(ReadINI("Settings", "OtherSplitNS"))
If x < 50 Then x = 140
splSplitterNS(6).Position = x
x = Val(ReadINI("Settings", "OtherSplitWE"))
If x < 50 Then x = 400
splSplitterWE(6).Position = x


x = Val(ReadINI("Settings", "MonsterSplitWE"))
If x < 50 Then x = 400
splMonsterSplit(0).Position = x
x = Val(ReadINI("Settings", "MonsterCompSplitWE"))
If x < 50 Then x = 400
splMonsterSplit(1).Position = x

'For x = 0 To splSplitterNS().UBound
'    Call splSplitterNS(x).Resize
'    Call splSplitterWE(x).Resize
'Next x

Call SyncSplitters(0)

Exit Sub
Error:
Call HandleError("SetupSplitterSizes")

End Sub

Private Sub splSplitterNS_MouseUp(Index As Integer)
Call SyncSplitters(Index)
End Sub

Private Sub splSplitterNS_Resize(Index As Integer)
'Dim x As Integer
'
'If splSplitterNS(index).m_bInDrag Then Exit Sub
'
'Select Case index
'    Case 0, 1, 2: 'wep,arm,spl
'        For x = 0 To 2
'            If Not splSplitterNS(x).Position = splSplitterNS(index).Position Then
'                splSplitterNS(x).Position = splSplitterNS(index).Position
'            End If
'        Next x
'    Case 3, 4, 5:
'        For x = 3 To 5
'            If Not splSplitterNS(x).Position = splSplitterNS(index).Position Then
'                splSplitterNS(x).Position = splSplitterNS(index).Position
'            End If
'        Next x
'End Select

End Sub

Private Sub splSplitterWE_MouseUp(Index As Integer)
Call SyncSplitters(Index)
End Sub

Private Sub splSplitterWE_Resize(Index As Integer)
Dim oLV As ListView, oLVExtended As ListView, oCH As ColumnHeader, z As Long ', x As Integer
On Error GoTo Error:

'If splSplitterWE(Index).m_bInDrag Then Exit Sub
'
'Select Case Index
'    Case 0, 1, 2: 'wep,arm,spl
'        For x = 0 To 2
'            If Not splSplitterWE(x).Position = splSplitterWE(Index).Position Then
'                splSplitterWE(x).Position = splSplitterWE(Index).Position
'            End If
'        Next x
'    Case 3, 4, 5:
'        For x = 3 To 5
'            If Not splSplitterWE(x).Position = splSplitterWE(Index).Position Then
'                splSplitterWE(x).Position = splSplitterWE(Index).Position
'            End If
'        Next x
'End Select

Select Case Index
    Case 0: 'weapon
        Set oLV = lvWeaponLoc
    Case 1: 'armour
        Set oLV = lvArmourLoc
    Case 2: 'spell
        Set oLV = lvSpellLoc
        Set oLVExtended = lvSpells
    Case 3: 'wep compare
        Set oLV = lvWeaponCompareLoc
    Case 4: 'armr compare
        Set oLV = lvArmourCompareLoc
    Case 5: 'spell compare
        Set oLV = lvSpellCompareLoc
        Set oLVExtended = lvSpellCompare
    Case 6:
        Set oLV = lvOtherItemLoc
End Select

If Not oLV Is Nothing Then
    If oLV.ColumnHeaders.Count > 0 Then
        oLV.ColumnHeaders(1).Width = oLV.Width - 500
    End If
End If

If Not oLVExtended Is Nothing Then
    If oLV.ColumnHeaders.Count > 0 Then
        For Each oCH In oLVExtended.ColumnHeaders
            z = z + oCH.Width
            Set oCH = Nothing
        Next oCH
        Set oCH = oLVExtended.ColumnHeaders(oLVExtended.ColumnHeaders.Count)
        oCH.Width = oLVExtended.Width - z + oCH.Width - 500
    End If
End If

out:
Set oLV = Nothing
Set oLVExtended = Nothing
Set oCH = Nothing
Exit Sub

Error:
Call HandleError("splSplitterWE_Resize")
Resume out:
End Sub

Private Sub SyncSplitters(Index As Integer)
On Error GoTo Error:
Dim x As Integer, y1 As Integer, y2 As Integer

Select Case Index
    Case 0:
        y1 = 0
        y2 = 3
    Case 1:
        y1 = 1
        y2 = 3
    Case 2:
        y1 = 2
        y2 = 3
    Case 3:
        y1 = 0
        y2 = 3
    Case 4:
        y1 = 0
        y2 = 4
    Case 5:
        y1 = 0
        y2 = 5
End Select

For x = 0 To 2
    If Not splSplitterNS(x).Position = splSplitterNS(y1).Position Then
        splSplitterNS(x).Position = splSplitterNS(y1).Position
    End If
    If Not splSplitterWE(x).Position = splSplitterWE(y1).Position Then
        splSplitterWE(x).Position = splSplitterWE(y1).Position
    End If
Next x
For x = 3 To 5
    If Not splSplitterNS(x).Position = splSplitterNS(y2).Position Then
        splSplitterNS(x).Position = splSplitterNS(y2).Position
    End If
    If Not splSplitterWE(x).Position = splSplitterWE(y2).Position Then
        splSplitterWE(x).Position = splSplitterWE(y2).Position
    End If
Next x


Exit Sub
Error:
Call HandleError("SyncSplitters")
End Sub

Private Function TestGlobalFilter(ByVal nItemNumber As Long) As Boolean
Dim x As Integer, bClassOK As Boolean, nClass As Integer, bNoMagic As Boolean
Dim nNotAlign As Integer, nAlign As Integer, nClassWeapon As Integer, nClassArmour As Integer
On Error GoTo Error:

If Not tabItems.Fields("Number") = nItemNumber Then
    tabItems.Index = "pkItems"
    tabItems.Seek "=", nItemNumber
    If tabItems.NoMatch Then Exit Function
End If

If tabClasses.RecordCount = 0 Or cmbGlobalClass(0).ListIndex < 1 Then
    TestGlobalFilter = True
    GoTo skip:
End If

tabClasses.Index = "pkClasses"
tabClasses.Seek "=", cmbGlobalClass(0).ItemData(cmbGlobalClass(0).ListIndex)
If tabClasses.NoMatch = True Then
    TestGlobalFilter = True
    GoTo skip:
End If

nClassArmour = tabClasses.Fields("ArmourType")
nClassWeapon = tabClasses.Fields("WeaponType")

For x = 0 To 9
    If tabClasses.Fields("Abil-" & x) = 51 Then 'anti-magic check
        bNoMagic = True
        Exit For
    End If
Next

For x = 0 To 19
    Select Case tabItems.Fields("Abil-" & x)
        Case 0:
        Case 135: 'min level check
            If tabItems.Fields("AbilVal-" & x) > Val(txtGlobalLevel(0).Text) Then GoTo skip:

        Case 136: 'max level check
            If tabItems.Fields("AbilVal-" & x) < Val(txtGlobalLevel(0).Text) Then GoTo skip:

        Case 59: 'classok
            If tabItems.Fields("AbilVal-" & x) = cmbGlobalClass(0).ItemData(cmbGlobalClass(0).ListIndex) Then
                bClassOK = True
            End If
            
        Case 28: 'magical check
            If bNoMagic = True Then GoTo skip:

        Case 97, 98, 112: 'good/evil/neutral abils
            nAlign = tabItems.Fields("Abil-" & x)
            Select Case cmbGlobalAlignment.ListIndex
                Case 0:
                Case 1: 'good
                    If Not nAlign = 97 Then GoTo skip:
                Case 2: 'netural
                    If Not nAlign = 112 Then GoTo skip:
                Case 3: 'evil
                    If Not nAlign = 98 Then GoTo skip:
            End Select
    
        Case 110, 111, 113: 'notgood/notevil/notneutral abils
            nNotAlign = tabItems.Fields("Abil-" & x)
            Select Case cmbGlobalAlignment.ListIndex
                Case 0:
                Case 1: 'good
                    If nNotAlign = 110 Then GoTo skip:
                Case 2: 'netural
                    If nNotAlign = 113 Then GoTo skip:
                Case 3: 'evil
                    If nNotAlign = 111 Then GoTo skip:
            End Select
            
    End Select
Next
          
If bClassOK = False Then
    If cmbGlobalClass(0).ItemData(cmbGlobalClass(0).ListIndex) = 0 Then
        nClass = 1
    Else
        For x = 0 To 9
            'nclass = 0 = undetermined
            'nclass = -1 = there are class restrictions and no match found yet
            'nclass = 1 = match found
            If Not tabItems.Fields("ClassRest-" & x) = 0 And nClass = 0 Then nClass = -1
            If tabItems.Fields("ClassRest-" & x) = cmbGlobalClass(0).ItemData(cmbGlobalClass(0).ListIndex) Then
                bClassOK = True
                nClass = 1
                Exit For
            End If
        Next x
    End If
    
    If nClass = -1 Then GoTo skip: 'if it fails class checks
End If

'class, level, and alignment ok ... now check armour/weapon restrictions
If bClassOK Then GoTo skip_type_check:

Select Case tabItems.Fields("ItemType")
    Case 0: 'armour
        If nClassArmour < tabItems.Fields("ArmourType") Then GoTo skip:
        
        If tabItems.Fields("Worn") = 12 Then
            Select Case nClassWeapon
                Case 0, 2, 4, 9: '1H
                    If Not bClassOK Then GoTo skip:
            End Select
        End If
        
    Case 1: 'weapon
        Select Case nClassWeapon
            '0=1h blunt
            '1=2h blunt
            '2=1h sharp
            '3=2h sharp
            Case 0: '"1H Blunt"
                If Not tabItems.Fields("WeaponType") = 0 Then GoTo skip:
            Case 1: '"2H Blunt"
                If Not tabItems.Fields("WeaponType") = 1 Then GoTo skip:
            Case 2: '"1H Sharp"
                If Not tabItems.Fields("WeaponType") = 2 Then GoTo skip:
            Case 3: '"2H Sharp"
                If Not tabItems.Fields("WeaponType") = 3 Then GoTo skip:
            Case 4: '"Any 1H"
                If Not tabItems.Fields("WeaponType") = 0 And Not tabItems.Fields("WeaponType") = 2 Then GoTo skip:
            Case 5: '"Any 2H"
                If Not tabItems.Fields("WeaponType") = 1 And Not tabItems.Fields("WeaponType") = 3 Then GoTo skip:
            Case 6: '"Any Sharp"
                If Not tabItems.Fields("WeaponType") >= 2 Then GoTo skip:
            Case 7: '"Any Blunt"
                If Not tabItems.Fields("WeaponType") <= 1 Then GoTo skip:
            Case 8: '"All Weapons"
            Case 9: '"Staff"
                If bClassOK = False Then GoTo skip:
        End Select
End Select

skip_type_check:

TestGlobalFilter = True
skip:
Exit Function

Error:
Call HandleError("TestGlobalFilter")
    
End Function

Private Sub timMouseDown_Timer()
timMouseDown.Enabled = False
End Sub

Private Sub txtArmourCompareDetail_GotFocus()
'Call SelectAll(txtArmourCompareDetail)

End Sub

Private Sub txtArmourDetail_GotFocus()
'Call SelectAll(txtArmourDetail)

End Sub

Private Sub txtArmourFind_GotFocus()
Call SelectAll(txtArmourFind)
End Sub

Private Sub txtArmourFind_KeyUp(KeyCode As Integer, Shift As Integer)
Dim bFound As Boolean

bFound = SearchLV(KeyCode, lvArmour, txtArmourFind)
If bFound Then Call lvArmour_ItemClick(lvArmour.SelectedItem)
End Sub

Private Sub txtCharHPRegen_Change()

'bPromptSave = True
Call RefreshHitPoints
End Sub

Private Sub txtCharHPRegen_GotFocus()
Call SelectAll(txtCharHPRegen)
End Sub

Private Sub txtCharHPRegen_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtCharManaRegen_Change()

'bPromptSave = True
Call RefreshMagic
End Sub

Private Sub txtCharManaRegen_GotFocus()
Call SelectAll(txtCharManaRegen)
End Sub

Private Sub txtCharName_Change()
bPromptSave = True
If bNameInTitle Then
    If txtCharName.Text = "" Then
        Me.Caption = sNormalCaption
    Else
        Me.Caption = sNormalCaption & " (" & txtCharName.Text & ")"
    End If
End If
End Sub

Private Sub txtCharName_GotFocus()
Call SelectAll(txtCharName)
End Sub

Private Sub txtCharStats_Change(Index As Integer)

bPromptSave = True

Call RefreshAll

End Sub

Private Sub txtCharStats_GotFocus(Index As Integer)
Call SelectAll(txtCharStats(Index))
End Sub

Private Sub txtCharStats_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtClassDetail_GotFocus()
Call SelectAll(txtClassDetail)

End Sub

Private Sub txtGlobalLevel_Change(Index As Integer)
Dim x As Integer

bPromptSave = True

With txtGlobalLevel()
    For x = 0 To .UBound
        If Not x = Index Then
            If Not .item(x).Text = .item(Index).Text Then
                .item(x).Text = .item(Index).Text
                Exit Sub
            End If
        End If
    Next x
End With

If Not lvSpells.SelectedItem Is Nothing Then
    Call lvSpells_ItemClick(lvSpells.SelectedItem)
End If
If Not lvSpellCompare.SelectedItem Is Nothing Then
    Call lvSpellCompare_ItemClick(lvSpellCompare.SelectedItem)
End If

Call RefreshAll
End Sub

Private Sub txtGlobalLevel_GotFocus(Index As Integer)
Call SelectAll(txtGlobalLevel(Index))
End Sub

Private Sub txtGlobalLevel_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtInvenAddWeight_Change()

bPromptSave = True
Call RefreshAll
End Sub

Private Sub txtInvenAddWeight_GotFocus()
Call SelectAll(txtInvenAddWeight)
End Sub

Private Sub txtInvenAddWeight_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtInvenStrength_Change()

bPromptSave = True
            
If Not txtInvenStrength.Text = txtCharStats(0).Text Then
    txtCharStats(0).Text = txtInvenStrength.Text
    Exit Sub
End If

Call RefreshAll
End Sub

Private Sub txtInvenStrength_GotFocus()
Call SelectAll(txtInvenStrength)
End Sub

Private Sub txtInvenStrength_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtMapMove_KeyPress(KeyAscii As Integer)
Dim sLook As String, RoomExit As RoomExitType
Dim nExitType As Integer, nRecNum As Long
On Error GoTo Error:

If bMapStillMapping Then
    KeyAscii = 0
    Exit Sub
End If

If cmdMapSwitch.Tag = "1" Then Exit Sub

Select Case KeyAscii
    Case 8, 26: Exit Sub
End Select

If Len(txtMapMove.Text) > 30000 Then
    nRecNum = MsgBox("Maximum Length Reached, Clear Direction Box?", vbQuestion + vbYesNo + vbDefaultButton2)
    If nRecNum = vbYes Then
        txtMapMove.Text = ""
    Else
        GoTo out:
    End If
End If

tabRooms.Index = "idxRooms"
tabRooms.Seek "=", nMapStartMap, nMapStartRoom
If tabRooms.NoMatch Then GoTo out:

Select Case KeyAscii
    Case 46: 'd
        sLook = "D"
    Case 48: 'u
        sLook = "U"
    Case 49: 'sw
        sLook = "SW"
    Case 50: 's
        sLook = "S"
    Case 51: 'se
        sLook = "SE"
    Case 52: 'w
        sLook = "W"
    'Case 53:
    Case 54: 'e
        sLook = "E"
    Case 55: 'nw
        sLook = "NW"
    Case 56: 'n
        sLook = "N"
    Case 57: 'ne
        sLook = "NE"
    Case Else: GoTo out:
End Select

If Left(tabRooms.Fields(sLook), 6) = "Action" Then
    GoTo out:
ElseIf Not Val(tabRooms.Fields(sLook)) = 0 Then
    RoomExit = ExtractMapRoom(tabRooms.Fields(sLook))
    
    tabRooms.Index = "idxRooms"
    tabRooms.Seek "=", RoomExit.Map, RoomExit.Room
    If tabRooms.NoMatch Then
        MsgBox "Error going in that direction."
        GoTo out:
    End If
Else
    GoTo out:
End If

If Len(RoomExit.ExitType) > 2 Then
    Select Case Left(RoomExit.ExitType, 5)
        Case "(Key:": nExitType = 2
        Case "(Item": nExitType = 3
        Case "(Toll": nExitType = 4
        Case "(Hidd": nExitType = 6
        Case "(Door": nExitType = 7
        Case "(Trap": nExitType = 9
        Case "(Text": nExitType = 10
        Case "(Gate": nExitType = 11
        Case "Actio": nExitType = 12
        Case "(Clas": nExitType = 13
        Case "(Race": nExitType = 14
        Case "(Leve": nExitType = 15
        Case "(Time": nExitType = 16
        Case "(Tick": nExitType = 17
        Case "(Max ": nExitType = 18
        Case "(Bloc": nExitType = 19
        Case "(Alig": nExitType = 20
        Case "(Dela": nExitType = 21
        Case "(Cast": nExitType = 22
        Case "(Abil": nExitType = 23
        Case "(Spel": nExitType = 24
    End Select
End If
If Not RoomExit.Map = nMapStartMap Then nExitType = 8 'map change

Select Case nExitType
    Case 8: 'map change
    Case 12: 'action
        GoTo out:
    Case 3: 'item
        nRecNum = ExtractNumbersFromString(RoomExit.ExitType)
        If nRecNum > 0 Then
            sLook = sLook & " -- (Requires " & GetItemName(nRecNum, bHideRecordNumbers) & ")"
        End If
    Case 2: 'key
        nRecNum = ExtractNumbersFromString(RoomExit.ExitType)
        If nRecNum > 0 Then
            sLook = sLook & "[" & "use " & GetItemName(nRecNum, bHideRecordNumbers) & " " & sLook & "]"
        Else
            sLook = sLook '& ": open door " & sLook
        End If
    Case 6:
        If InStr(1, LCase(RoomExit.ExitType), "action") > 0 Then
            sLook = sLook & " -- " & RoomExit.ExitType
        ElseIf InStr(1, LCase(RoomExit.ExitType), "passable") > 0 Then
            'nothing
        Else
            sLook = sLook & "[search " & sLook & "]"
        End If
    Case 10:
        sLook = ExtractTextCommand(RoomExit.ExitType)
    Case 4, 9, 13, 14, 15, 20: '
        sLook = sLook & " -- " & RoomExit.ExitType
End Select

txtMapMove.Text = IIf(txtMapMove.Text = "", "", txtMapMove.Text & vbCrLf) & sLook '"00000000:0080:"
txtMapMove.SelStart = Len(txtMapMove.Text)
txtMapMove.SelLength = 0
Call MapStartMapping(RoomExit.Map, RoomExit.Room)
out:
KeyAscii = 0

Exit Sub
Error:
Call HandleError

End Sub

Private Sub txtMonsterFind_GotFocus()
Call SelectAll(txtMonsterFind)
End Sub

Private Sub txtMonsterFind_KeyUp(KeyCode As Integer, Shift As Integer)
Dim bFound As Boolean

bFound = SearchLV(KeyCode, lvMonsters, txtMonsterFind)
If bFound Then Call lvMonsters_ItemClick(lvMonsters.SelectedItem)

End Sub

Private Sub txtOtherItemDetail_GotFocus()
'Call SelectAll(txtOtherItemDetail)

End Sub

Private Sub txtOtherItemsFind_GotFocus()
Call SelectAll(txtOtherItemsFind)
End Sub

Private Sub txtOtherItemsFind_KeyUp(KeyCode As Integer, Shift As Integer)
Dim bFound As Boolean

bFound = SearchLV(KeyCode, lvOtherItems, txtOtherItemsFind)
If bFound Then Call lvOtherItems_ItemClick(lvOtherItems.SelectedItem)

End Sub

Private Sub txtRaceDetail_GotFocus()
Call SelectAll(txtRaceDetail)

End Sub

Private Sub txtRoomMap_GotFocus()
Call SelectAll(txtRoomMap)
End Sub

Private Sub txtRoomMap_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call cmdDrawMap_Click(0)
    KeyAscii = 0
End If
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtRoomRoom_GotFocus()
Call SelectAll(txtRoomRoom)
End Sub

Private Sub txtRoomRoom_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call cmdDrawMap_Click(0)
    KeyAscii = 0
End If
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtShopCharm_Change()

If Not txtShopCharm.Text = txtCharStats(5).Text Then
    txtCharStats(5).Text = txtShopCharm.Text
    Exit Sub
End If

Call RefreshShops
End Sub

Private Sub txtShopCharm_GotFocus()
Call SelectAll(txtShopCharm)
End Sub

Private Sub txtShopCharm_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtShopDetail_GotFocus()
Call SelectAll(txtShopDetail)

End Sub

Private Sub txtShopFind_GotFocus()
Call SelectAll(txtShopFind)
End Sub

Private Sub txtShopFind_KeyUp(KeyCode As Integer, Shift As Integer)
Dim bFound As Boolean

bFound = SearchLV(KeyCode, lvShops, txtShopFind)
If bFound Then Call lvShops_ItemClick(lvShops.SelectedItem)

End Sub

Private Sub txtSpellCompareDetail_GotFocus()
'Call SelectAll(txtSpellCompareDetail)

End Sub

Private Sub txtSpellDetail_GotFocus()
'Call SelectAll(txtSpellDetail)

End Sub

Private Sub txtSpellFind_GotFocus()
Call SelectAll(txtSpellFind)
End Sub

Private Sub txtSpellFind_KeyUp(KeyCode As Integer, Shift As Integer)
Dim bFound As Boolean

bFound = SearchLV(KeyCode, lvSpells, txtSpellFind)
If bFound Then Call lvSpells_ItemClick(lvSpells.SelectedItem)

End Sub

Private Sub txtStat_Change(Index As Integer)
Dim nTest As Single, sToolTipString As String, rc As RECT
On Error GoTo Error:

Select Case Index
    Case 4: 'encum +%
        Call InvenCalcEncum
    Case 0, 1, 4: 'encum/max encum
        If Val(txtStat(1).Text) = 0 Then
            If Val(txtStat(0).Text) >= 1 Then
                lblEncumLevel(1).Caption = "- Heavy" & vbCrLf & "(" & Val(txtStat(0).Text) & "%)"
                lblEncumLevel(1).Tag = "- Heavy (" & Val(txtStat(0).Text) & "%)"
            Else
                lblEncumLevel(1).Caption = "- None" & vbCrLf & "(0%)"
                lblEncumLevel(1).Tag = "- None (0%)"
            End If
            GoTo tooltip:
        End If
        

'        nTest = Val(txtStat(0).Text) / Val(txtStat(1).Text)
'        If nTest < 0.17 Then
'            lblEncumLevel(1).Caption = "- None" & vbCrLf & "(" & (Round(Val(nTest), 2) * 100) & "%)"
'        ElseIf nTest < 0.34 Then
'            lblEncumLevel(1).Caption = "- Light" & vbCrLf & "(" & (Round(Val(nTest), 2) * 100) & "%)"
'        ElseIf nTest < 0.67 Then
'            lblEncumLevel(1).Caption = "- Medium" & vbCrLf & "(" & (Round(Val(nTest), 2) * 100) & "%)"
'        Else
'            lblEncumLevel(1).Caption = "- Heavy" & vbCrLf & "(" & (Round(Val(nTest), 2) * 100) & "%)"
'        End If
        nTest = Fix(Val(txtStat(0).Text) / Val(txtStat(1).Text) * 100)
        If nTest < 17 Then
            lblEncumLevel(1).Caption = "- None" & vbCrLf & "(" & nTest & "%)"
            lblEncumLevel(1).Tag = "- None (" & nTest & "%)"
        ElseIf nTest < 34 Then
            lblEncumLevel(1).Caption = "- Light" & vbCrLf & "(" & nTest & "%)"
            lblEncumLevel(1).Tag = "- Light (" & nTest & "%)"
        ElseIf nTest < 67 Then
            lblEncumLevel(1).Caption = "- Medium" & vbCrLf & "(" & nTest & "%)"
            lblEncumLevel(1).Tag = "- Medium (" & nTest & "%)"
        Else
            lblEncumLevel(1).Caption = "- Heavy" & vbCrLf & "(" & nTest & "%)"
            lblEncumLevel(1).Tag = "- Heavy (" & nTest & "%)"
        End If
        
tooltip:
        sToolTipString = GetEncumPercents(Val(txtStat(1).Text))
        
        objToolTip.DelToolTip txtStat(1).hwnd
        objToolTip.SetToolTipObj txtStat(1).hwnd, sToolTipString
        
        objToolTip.DelToolTip picStats.hwnd, 1
        rc.Left = lblEncumLevel(1).Left
        rc.Top = lblEncumLevel(1).Top
        rc.Bottom = (lblEncumLevel(1).Top + lblEncumLevel(1).Height)
        rc.Right = (lblEncumLevel(1).Left + lblEncumLevel(1).Width)
        objToolTip.SetToolTipItem picStats.hwnd, 1, rc.Left, rc.Top, rc.Right, rc.Bottom, sToolTipString, False
    Case Else:
End Select

Exit Sub

Error:
Call HandleError

End Sub

Private Sub txtWeaponCompareDetail_GotFocus()
'Call SelectAll(txtWeaponCompareDetail)

End Sub

Private Sub txtWeaponDetail_GotFocus()
'Call SelectAll(txtWeaponDetail)

End Sub

Private Sub txtWeaponFind_GotFocus()
Call SelectAll(txtWeaponFind)
End Sub

Private Sub txtWeaponFind_KeyUp(KeyCode As Integer, Shift As Integer)
Dim bFound As Boolean

bFound = SearchLV(KeyCode, lvWeapons, txtWeaponFind)
If bFound Then Call lvWeapons_ItemClick(lvWeapons.SelectedItem)

End Sub

Private Sub txtWeaponSpeed_GotFocus()
Call SelectAll(txtWeaponSpeed)
End Sub

Private Sub txtWeaponSpeed_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub
