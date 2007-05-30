VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMap 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7350
   Icon            =   "frmMap.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMapControls 
      BackColor       =   &H00000000&
      Caption         =   "Map Control"
      ForeColor       =   &H00E0E0E0&
      Height          =   1395
      Left            =   120
      TabIndex        =   94
      Top             =   420
      Visible         =   0   'False
      Width           =   1395
      Begin VB.CommandButton cmdMove 
         Caption         =   "D"
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
         Index           =   9
         Left            =   870
         MaskColor       =   &H80000016&
         TabIndex        =   32
         Top             =   990
         Width           =   375
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "U"
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
         Index           =   8
         Left            =   150
         MaskColor       =   &H80000016&
         TabIndex        =   31
         Top             =   990
         Width           =   375
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "SE"
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
         Index           =   6
         Left            =   870
         MaskColor       =   &H80000016&
         TabIndex        =   30
         Top             =   750
         Width           =   375
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "S"
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
         Index           =   1
         Left            =   510
         MaskColor       =   &H80000016&
         TabIndex        =   29
         Top             =   750
         Width           =   375
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "SW"
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
         Index           =   7
         Left            =   150
         MaskColor       =   &H80000016&
         TabIndex        =   28
         Top             =   750
         Width           =   375
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "E"
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
         Index           =   2
         Left            =   870
         MaskColor       =   &H80000016&
         TabIndex        =   27
         Top             =   510
         Width           =   375
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "W"
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
         Index           =   3
         Left            =   150
         MaskColor       =   &H80000016&
         TabIndex        =   26
         Top             =   510
         Width           =   375
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "NE"
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
         Index           =   4
         Left            =   870
         MaskColor       =   &H80000016&
         TabIndex        =   25
         Top             =   270
         Width           =   375
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "N"
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
         Index           =   0
         Left            =   510
         MaskColor       =   &H80000016&
         TabIndex        =   24
         Top             =   270
         Width           =   375
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "NW"
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
         Index           =   5
         Left            =   150
         MaskColor       =   &H80000016&
         TabIndex        =   23
         Top             =   270
         Width           =   375
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00000000&
      Caption         =   "Options"
      ForeColor       =   &H00E0E0E0&
      Height          =   4035
      Left            =   2280
      TabIndex        =   93
      Top             =   420
      Visible         =   0   'False
      Width           =   2595
      Begin VB.CommandButton cmdQ 
         Caption         =   "?"
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
         Left            =   2280
         TabIndex        =   17
         Top             =   2280
         Width           =   195
      End
      Begin VB.CheckBox chkMapOptions 
         BackColor       =   &H00000000&
         Caption         =   "Allow Main To Overlap"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   16
         Top             =   2340
         Width           =   2115
      End
      Begin VB.ComboBox cmbMapSize 
         Height          =   315
         ItemData        =   "frmMap.frx":0CCA
         Left            =   180
         List            =   "frmMap.frx":0CD4
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2580
         Width           =   2235
      End
      Begin VB.CommandButton cmdQ 
         Caption         =   "?"
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
         Left            =   2280
         TabIndex        =   14
         Top             =   1860
         Width           =   195
      End
      Begin VB.CheckBox chkMapOptions 
         BackColor       =   &H00000000&
         Caption         =   "Show Map Controls"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   13
         Top             =   1860
         Width           =   1875
      End
      Begin VB.CheckBox chkMapOptions 
         BackColor       =   &H00000000&
         Caption         =   "Not ""Always on Top"""
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   15
         Top             =   2100
         Width           =   1935
      End
      Begin VB.CommandButton cmdQ 
         Caption         =   "?"
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
         Left            =   2280
         TabIndex        =   6
         Top             =   240
         Width           =   195
      End
      Begin VB.CommandButton cmdViewMapLegend 
         Caption         =   "View Help/&Legend"
         Height          =   315
         Left            =   180
         TabIndex        =   22
         Top             =   3600
         Width           =   2235
      End
      Begin VB.CommandButton cmdMapShowUnused 
         Caption         =   "S&how Unused Blocks"
         Height          =   315
         Left            =   180
         TabIndex        =   21
         Top             =   3300
         Width           =   2235
      End
      Begin VB.CommandButton cmdMapFindText 
         Caption         =   "Find &Next"
         Height          =   315
         Index           =   1
         Left            =   1380
         TabIndex        =   20
         Top             =   3000
         Width           =   1035
      End
      Begin VB.CommandButton cmdMapFindText 
         Caption         =   "&Find Room"
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CheckBox chkMapOptions 
         BackColor       =   &H00000000&
         Caption         =   "Don't Follow Hidden Exits"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   510
         Width           =   2235
      End
      Begin VB.CheckBox chkMapOptions 
         BackColor       =   &H00000000&
         Caption         =   "Follow Map Changes"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   240
         Width           =   1875
      End
      Begin VB.CheckBox chkMapOptions 
         BackColor       =   &H00000000&
         Caption         =   "Don't Mark Lairs"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   750
         Width           =   2235
      End
      Begin VB.CheckBox chkMapOptions 
         BackColor       =   &H00000000&
         Caption         =   "Don't Mark NPCs"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   10
         Top             =   1005
         Width           =   2235
      End
      Begin VB.CheckBox chkMapOptions 
         BackColor       =   &H00000000&
         Caption         =   "Don't Mark Commands"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   11
         Top             =   1260
         Width           =   2235
      End
      Begin VB.CheckBox chkMapOptions 
         BackColor       =   &H00000000&
         Caption         =   "Don't Show Tooltips"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   12
         Top             =   1500
         Width           =   2235
      End
   End
   Begin VB.Frame fraPresets 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Presets"
      ForeColor       =   &H00E0E0E0&
      Height          =   3855
      Left            =   2280
      TabIndex        =   60
      Top             =   420
      Visible         =   0   'False
      Width           =   2595
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
         Height          =   255
         Index           =   4
         Left            =   1320
         TabIndex        =   37
         Top             =   300
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
         Height          =   255
         Index           =   3
         Left            =   1020
         TabIndex        =   36
         Top             =   300
         Width           =   315
      End
      Begin VB.CommandButton cmdEditPreset 
         Caption         =   "!"
         Height          =   315
         Index           =   9
         Left            =   2280
         TabIndex        =   58
         Top             =   3420
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPreset 
         Caption         =   "!"
         Height          =   315
         Index           =   8
         Left            =   2280
         TabIndex        =   56
         Top             =   3120
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPreset 
         Caption         =   "!"
         Height          =   315
         Index           =   7
         Left            =   2280
         TabIndex        =   54
         Top             =   2820
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPreset 
         Caption         =   "!"
         Height          =   315
         Index           =   6
         Left            =   2280
         TabIndex        =   52
         Top             =   2520
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPreset 
         Caption         =   "!"
         Height          =   315
         Index           =   5
         Left            =   2280
         TabIndex        =   50
         Top             =   2220
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPreset 
         Caption         =   "!"
         Height          =   315
         Index           =   4
         Left            =   2280
         TabIndex        =   48
         Top             =   1920
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPreset 
         Caption         =   "!"
         Height          =   315
         Index           =   3
         Left            =   2280
         TabIndex        =   46
         Top             =   1620
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPreset 
         Caption         =   "!"
         Height          =   315
         Index           =   2
         Left            =   2280
         TabIndex        =   44
         Top             =   1320
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPreset 
         Caption         =   "!"
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   42
         Top             =   1020
         Width           =   195
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
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   35
         Top             =   300
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
         Height          =   255
         Index           =   1
         Left            =   420
         TabIndex        =   34
         Top             =   300
         Width           =   315
      End
      Begin VB.CommandButton cmdMapPreset 
         Caption         =   "Lava Fields"
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   57
         Top             =   3420
         Width           =   2115
      End
      Begin VB.CommandButton cmdMapPreset 
         Caption         =   "Ancient Ruin"
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   55
         Top             =   3120
         Width           =   2115
      End
      Begin VB.CommandButton cmdMapPreset 
         Caption         =   "Storm Fortress"
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   53
         Top             =   2820
         Width           =   2115
      End
      Begin VB.CommandButton cmdMapPreset 
         Caption         =   "Black Fortress"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   51
         Top             =   2520
         Width           =   2115
      End
      Begin VB.CommandButton cmdMapPreset 
         Caption         =   "Commander Markus"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   49
         Top             =   2220
         Width           =   2115
      End
      Begin VB.CommandButton cmdMapPreset 
         Caption         =   "Rhudar"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   47
         Top             =   1920
         Width           =   2115
      End
      Begin VB.CommandButton cmdMapPreset 
         Caption         =   "Lost City"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   45
         Top             =   1620
         Width           =   2115
      End
      Begin VB.CommandButton cmdMapPreset 
         Caption         =   "Arlysia"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   43
         Top             =   1320
         Width           =   2115
      End
      Begin VB.CommandButton cmdMapPreset 
         Caption         =   "Khazarad"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   1020
         Width           =   2115
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   300
         Width           =   315
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
         Height          =   255
         Left            =   1800
         TabIndex        =   38
         Top             =   300
         Width           =   675
      End
      Begin VB.CommandButton cmdMapPreset 
         Caption         =   "Town Square"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Width           =   2115
      End
      Begin VB.CommandButton cmdEditPreset 
         Caption         =   "!"
         Height          =   315
         Index           =   0
         Left            =   2280
         TabIndex        =   40
         Top             =   720
         Width           =   195
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   60
      TabIndex        =   92
      Top             =   60
      Width           =   7245
      Begin VB.CommandButton cmdDrawMap 
         Caption         =   "&Draw"
         Default         =   -1  'True
         Height          =   315
         Index           =   0
         Left            =   1380
         TabIndex        =   2
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPresets 
         Caption         =   "&Presets"
         Height          =   315
         Left            =   3900
         TabIndex        =   5
         ToolTipText     =   "Goes back one room"
         Top             =   0
         Width           =   915
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "&Options"
         Height          =   315
         Left            =   2880
         TabIndex        =   4
         ToolTipText     =   "Goes back one room"
         Top             =   0
         Width           =   975
      End
      Begin VB.TextBox txtRoomRoom 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   600
         MaxLength       =   5
         TabIndex        =   1
         Text            =   "1"
         Top             =   0
         Width           =   735
      End
      Begin VB.TextBox txtRoomMap 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   0
         MaxLength       =   5
         TabIndex        =   0
         Text            =   "1"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdDrawMap 
         Caption         =   "&Last"
         Height          =   315
         Index           =   1
         Left            =   2220
         TabIndex        =   3
         ToolTipText     =   "Back to last room"
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   7245
      Left            =   60
      ScaleHeight     =   481
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   481
      TabIndex        =   59
      Top             =   390
      Width           =   7245
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   900
         Left            =   7020
         TabIndex        =   964
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   899
         Left            =   6780
         TabIndex        =   963
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   898
         Left            =   6540
         TabIndex        =   962
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   897
         Left            =   6300
         TabIndex        =   961
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   896
         Left            =   6060
         TabIndex        =   960
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   895
         Left            =   5820
         TabIndex        =   959
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   894
         Left            =   5580
         TabIndex        =   958
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   893
         Left            =   5340
         TabIndex        =   957
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   892
         Left            =   5100
         TabIndex        =   956
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   891
         Left            =   4860
         TabIndex        =   955
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   890
         Left            =   4620
         TabIndex        =   954
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   889
         Left            =   4380
         TabIndex        =   953
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   888
         Left            =   4140
         TabIndex        =   952
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   887
         Left            =   3900
         TabIndex        =   951
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   886
         Left            =   3660
         TabIndex        =   950
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   885
         Left            =   3420
         TabIndex        =   949
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   884
         Left            =   3180
         TabIndex        =   948
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   883
         Left            =   2940
         TabIndex        =   947
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   882
         Left            =   2700
         TabIndex        =   946
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   881
         Left            =   2460
         TabIndex        =   945
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   880
         Left            =   2220
         TabIndex        =   944
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   879
         Left            =   1980
         TabIndex        =   943
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   878
         Left            =   1740
         TabIndex        =   942
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   877
         Left            =   1500
         TabIndex        =   941
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   876
         Left            =   1260
         TabIndex        =   940
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   875
         Left            =   1020
         TabIndex        =   939
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   874
         Left            =   780
         TabIndex        =   938
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   873
         Left            =   540
         TabIndex        =   937
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   872
         Left            =   300
         TabIndex        =   936
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   871
         Left            =   60
         TabIndex        =   935
         Top             =   7020
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   870
         Left            =   7020
         TabIndex        =   934
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   869
         Left            =   6780
         TabIndex        =   933
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   868
         Left            =   6540
         TabIndex        =   932
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   867
         Left            =   6300
         TabIndex        =   931
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   866
         Left            =   6060
         TabIndex        =   930
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   865
         Left            =   5820
         TabIndex        =   929
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   864
         Left            =   5580
         TabIndex        =   928
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   863
         Left            =   5340
         TabIndex        =   927
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   862
         Left            =   5100
         TabIndex        =   926
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   861
         Left            =   4860
         TabIndex        =   925
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   860
         Left            =   4620
         TabIndex        =   924
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   859
         Left            =   4380
         TabIndex        =   923
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   858
         Left            =   4140
         TabIndex        =   922
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   857
         Left            =   3900
         TabIndex        =   921
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   856
         Left            =   3660
         TabIndex        =   920
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   855
         Left            =   3420
         TabIndex        =   919
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   854
         Left            =   3180
         TabIndex        =   918
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   853
         Left            =   2940
         TabIndex        =   917
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   852
         Left            =   2700
         TabIndex        =   916
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   851
         Left            =   2460
         TabIndex        =   915
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   850
         Left            =   2220
         TabIndex        =   914
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   849
         Left            =   1980
         TabIndex        =   913
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   848
         Left            =   1740
         TabIndex        =   912
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   847
         Left            =   1500
         TabIndex        =   911
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   846
         Left            =   1260
         TabIndex        =   910
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   845
         Left            =   1020
         TabIndex        =   909
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   844
         Left            =   780
         TabIndex        =   908
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   843
         Left            =   540
         TabIndex        =   907
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   842
         Left            =   300
         TabIndex        =   906
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   841
         Left            =   60
         TabIndex        =   905
         Top             =   6780
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   840
         Left            =   7020
         TabIndex        =   904
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   839
         Left            =   6780
         TabIndex        =   903
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   838
         Left            =   6540
         TabIndex        =   902
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   837
         Left            =   6300
         TabIndex        =   901
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   836
         Left            =   6060
         TabIndex        =   900
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   835
         Left            =   5820
         TabIndex        =   899
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   834
         Left            =   5580
         TabIndex        =   898
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   833
         Left            =   5340
         TabIndex        =   897
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   832
         Left            =   5100
         TabIndex        =   896
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   831
         Left            =   4860
         TabIndex        =   895
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   830
         Left            =   4620
         TabIndex        =   894
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   829
         Left            =   4380
         TabIndex        =   893
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   828
         Left            =   4140
         TabIndex        =   892
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   827
         Left            =   3900
         TabIndex        =   891
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   826
         Left            =   3660
         TabIndex        =   890
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   825
         Left            =   3420
         TabIndex        =   889
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   824
         Left            =   3180
         TabIndex        =   888
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   823
         Left            =   2940
         TabIndex        =   887
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   822
         Left            =   2700
         TabIndex        =   886
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   821
         Left            =   2460
         TabIndex        =   885
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   820
         Left            =   2220
         TabIndex        =   884
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   819
         Left            =   1980
         TabIndex        =   883
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   818
         Left            =   1740
         TabIndex        =   882
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   817
         Left            =   1500
         TabIndex        =   881
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   816
         Left            =   1260
         TabIndex        =   880
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   815
         Left            =   1020
         TabIndex        =   879
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   814
         Left            =   780
         TabIndex        =   878
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   813
         Left            =   540
         TabIndex        =   877
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   812
         Left            =   300
         TabIndex        =   876
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   811
         Left            =   60
         TabIndex        =   875
         Top             =   6540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   810
         Left            =   7020
         TabIndex        =   874
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   809
         Left            =   6780
         TabIndex        =   873
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   808
         Left            =   6540
         TabIndex        =   872
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   807
         Left            =   6300
         TabIndex        =   871
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   806
         Left            =   6060
         TabIndex        =   870
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   805
         Left            =   5820
         TabIndex        =   869
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   804
         Left            =   5580
         TabIndex        =   868
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   803
         Left            =   5340
         TabIndex        =   867
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   802
         Left            =   5100
         TabIndex        =   866
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   801
         Left            =   4860
         TabIndex        =   865
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   800
         Left            =   4620
         TabIndex        =   864
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   799
         Left            =   4380
         TabIndex        =   863
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   798
         Left            =   4140
         TabIndex        =   862
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   797
         Left            =   3900
         TabIndex        =   861
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   796
         Left            =   3660
         TabIndex        =   860
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   795
         Left            =   3420
         TabIndex        =   859
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   794
         Left            =   3180
         TabIndex        =   858
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   793
         Left            =   2940
         TabIndex        =   857
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   792
         Left            =   2700
         TabIndex        =   856
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   791
         Left            =   2460
         TabIndex        =   855
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   790
         Left            =   2220
         TabIndex        =   854
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   789
         Left            =   1980
         TabIndex        =   853
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   788
         Left            =   1740
         TabIndex        =   852
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   787
         Left            =   1500
         TabIndex        =   851
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   786
         Left            =   1260
         TabIndex        =   850
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   785
         Left            =   1020
         TabIndex        =   849
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   784
         Left            =   780
         TabIndex        =   848
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   783
         Left            =   540
         TabIndex        =   847
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   782
         Left            =   300
         TabIndex        =   846
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   781
         Left            =   60
         TabIndex        =   845
         Top             =   6300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   780
         Left            =   7020
         TabIndex        =   844
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   779
         Left            =   6780
         TabIndex        =   843
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   778
         Left            =   6540
         TabIndex        =   842
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   777
         Left            =   6300
         TabIndex        =   841
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   776
         Left            =   6060
         TabIndex        =   840
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   775
         Left            =   5820
         TabIndex        =   839
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   774
         Left            =   5580
         TabIndex        =   838
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   773
         Left            =   5340
         TabIndex        =   837
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   772
         Left            =   5100
         TabIndex        =   836
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   771
         Left            =   4860
         TabIndex        =   835
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   770
         Left            =   4620
         TabIndex        =   834
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   769
         Left            =   4380
         TabIndex        =   833
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   768
         Left            =   4140
         TabIndex        =   832
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   767
         Left            =   3900
         TabIndex        =   831
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   766
         Left            =   3660
         TabIndex        =   830
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   765
         Left            =   3420
         TabIndex        =   829
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   764
         Left            =   3180
         TabIndex        =   828
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   763
         Left            =   2940
         TabIndex        =   827
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   762
         Left            =   2700
         TabIndex        =   826
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   761
         Left            =   2460
         TabIndex        =   825
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   760
         Left            =   2220
         TabIndex        =   824
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   759
         Left            =   1980
         TabIndex        =   823
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   758
         Left            =   1740
         TabIndex        =   822
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   757
         Left            =   1500
         TabIndex        =   821
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   756
         Left            =   1260
         TabIndex        =   820
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   755
         Left            =   1020
         TabIndex        =   819
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   754
         Left            =   780
         TabIndex        =   818
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   753
         Left            =   540
         TabIndex        =   817
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   752
         Left            =   300
         TabIndex        =   816
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   751
         Left            =   60
         TabIndex        =   815
         Top             =   6060
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   750
         Left            =   7020
         TabIndex        =   814
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   749
         Left            =   6780
         TabIndex        =   813
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   748
         Left            =   6540
         TabIndex        =   812
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   747
         Left            =   6300
         TabIndex        =   811
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   746
         Left            =   6060
         TabIndex        =   810
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   745
         Left            =   5820
         TabIndex        =   809
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   744
         Left            =   5580
         TabIndex        =   808
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   743
         Left            =   5340
         TabIndex        =   807
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   742
         Left            =   5100
         TabIndex        =   806
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   741
         Left            =   4860
         TabIndex        =   805
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   740
         Left            =   4620
         TabIndex        =   804
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   739
         Left            =   4380
         TabIndex        =   803
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   738
         Left            =   4140
         TabIndex        =   802
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   737
         Left            =   3900
         TabIndex        =   801
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   736
         Left            =   3660
         TabIndex        =   800
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   735
         Left            =   3420
         TabIndex        =   799
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   734
         Left            =   3180
         TabIndex        =   798
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   733
         Left            =   2940
         TabIndex        =   797
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   732
         Left            =   2700
         TabIndex        =   796
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   731
         Left            =   2460
         TabIndex        =   795
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   730
         Left            =   2220
         TabIndex        =   794
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   729
         Left            =   1980
         TabIndex        =   793
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   728
         Left            =   1740
         TabIndex        =   792
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   727
         Left            =   1500
         TabIndex        =   791
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   726
         Left            =   1260
         TabIndex        =   790
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   725
         Left            =   1020
         TabIndex        =   789
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   724
         Left            =   780
         TabIndex        =   788
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   723
         Left            =   540
         TabIndex        =   787
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   722
         Left            =   300
         TabIndex        =   786
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   721
         Left            =   60
         TabIndex        =   785
         Top             =   5820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   720
         Left            =   7020
         TabIndex        =   784
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   719
         Left            =   6780
         TabIndex        =   783
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   718
         Left            =   6540
         TabIndex        =   782
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   717
         Left            =   6300
         TabIndex        =   781
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   716
         Left            =   6060
         TabIndex        =   780
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   715
         Left            =   5820
         TabIndex        =   779
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   714
         Left            =   5580
         TabIndex        =   778
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   713
         Left            =   5340
         TabIndex        =   777
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   712
         Left            =   5100
         TabIndex        =   776
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   711
         Left            =   4860
         TabIndex        =   775
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   710
         Left            =   4620
         TabIndex        =   774
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   709
         Left            =   4380
         TabIndex        =   773
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   708
         Left            =   4140
         TabIndex        =   772
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   707
         Left            =   3900
         TabIndex        =   771
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   706
         Left            =   3660
         TabIndex        =   770
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   705
         Left            =   3420
         TabIndex        =   769
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   704
         Left            =   3180
         TabIndex        =   768
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   703
         Left            =   2940
         TabIndex        =   767
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   702
         Left            =   2700
         TabIndex        =   766
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   701
         Left            =   2460
         TabIndex        =   765
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   700
         Left            =   2220
         TabIndex        =   764
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   699
         Left            =   1980
         TabIndex        =   763
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   698
         Left            =   1740
         TabIndex        =   762
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   697
         Left            =   1500
         TabIndex        =   761
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   696
         Left            =   1260
         TabIndex        =   760
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   695
         Left            =   1020
         TabIndex        =   759
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   694
         Left            =   780
         TabIndex        =   758
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   693
         Left            =   540
         TabIndex        =   757
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   692
         Left            =   300
         TabIndex        =   756
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   691
         Left            =   60
         TabIndex        =   755
         Top             =   5580
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   690
         Left            =   7020
         TabIndex        =   754
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   689
         Left            =   6780
         TabIndex        =   753
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   688
         Left            =   6540
         TabIndex        =   752
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   687
         Left            =   6300
         TabIndex        =   751
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   686
         Left            =   6060
         TabIndex        =   750
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   685
         Left            =   5820
         TabIndex        =   749
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   684
         Left            =   5580
         TabIndex        =   748
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   683
         Left            =   5340
         TabIndex        =   747
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   682
         Left            =   5100
         TabIndex        =   746
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   681
         Left            =   4860
         TabIndex        =   745
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   680
         Left            =   4620
         TabIndex        =   744
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   679
         Left            =   4380
         TabIndex        =   743
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   678
         Left            =   4140
         TabIndex        =   742
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   677
         Left            =   3900
         TabIndex        =   741
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   676
         Left            =   3660
         TabIndex        =   740
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   675
         Left            =   3420
         TabIndex        =   739
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   674
         Left            =   3180
         TabIndex        =   738
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   673
         Left            =   2940
         TabIndex        =   737
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   672
         Left            =   2700
         TabIndex        =   736
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   671
         Left            =   2460
         TabIndex        =   735
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   670
         Left            =   2220
         TabIndex        =   734
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   669
         Left            =   1980
         TabIndex        =   733
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   668
         Left            =   1740
         TabIndex        =   732
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   667
         Left            =   1500
         TabIndex        =   731
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   666
         Left            =   1260
         TabIndex        =   730
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   665
         Left            =   1020
         TabIndex        =   729
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   664
         Left            =   780
         TabIndex        =   728
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   663
         Left            =   540
         TabIndex        =   727
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   662
         Left            =   300
         TabIndex        =   726
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   661
         Left            =   60
         TabIndex        =   725
         Top             =   5340
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   660
         Left            =   7020
         TabIndex        =   724
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   659
         Left            =   6780
         TabIndex        =   723
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   658
         Left            =   6540
         TabIndex        =   722
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   657
         Left            =   6300
         TabIndex        =   721
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   656
         Left            =   6060
         TabIndex        =   720
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   655
         Left            =   5820
         TabIndex        =   719
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   654
         Left            =   5580
         TabIndex        =   718
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   653
         Left            =   5340
         TabIndex        =   717
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   652
         Left            =   5100
         TabIndex        =   716
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   651
         Left            =   4860
         TabIndex        =   715
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   650
         Left            =   4620
         TabIndex        =   714
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   649
         Left            =   4380
         TabIndex        =   713
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   648
         Left            =   4140
         TabIndex        =   712
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   647
         Left            =   3900
         TabIndex        =   711
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   646
         Left            =   3660
         TabIndex        =   710
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   645
         Left            =   3420
         TabIndex        =   709
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   644
         Left            =   3180
         TabIndex        =   708
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   643
         Left            =   2940
         TabIndex        =   707
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   642
         Left            =   2700
         TabIndex        =   706
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   641
         Left            =   2460
         TabIndex        =   705
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   640
         Left            =   2220
         TabIndex        =   704
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   639
         Left            =   1980
         TabIndex        =   703
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   638
         Left            =   1740
         TabIndex        =   702
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   637
         Left            =   1500
         TabIndex        =   701
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   636
         Left            =   1260
         TabIndex        =   700
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   635
         Left            =   1020
         TabIndex        =   699
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   634
         Left            =   780
         TabIndex        =   698
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   633
         Left            =   540
         TabIndex        =   697
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   632
         Left            =   300
         TabIndex        =   696
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   631
         Left            =   60
         TabIndex        =   695
         Top             =   5100
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   630
         Left            =   7020
         TabIndex        =   694
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   629
         Left            =   6780
         TabIndex        =   693
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   628
         Left            =   6540
         TabIndex        =   692
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   627
         Left            =   6300
         TabIndex        =   691
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   626
         Left            =   6060
         TabIndex        =   690
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   625
         Left            =   5820
         TabIndex        =   689
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   624
         Left            =   5580
         TabIndex        =   688
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   623
         Left            =   5340
         TabIndex        =   687
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   622
         Left            =   5100
         TabIndex        =   686
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   621
         Left            =   4860
         TabIndex        =   685
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   620
         Left            =   4620
         TabIndex        =   684
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   619
         Left            =   4380
         TabIndex        =   683
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   618
         Left            =   4140
         TabIndex        =   682
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   617
         Left            =   3900
         TabIndex        =   681
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   616
         Left            =   3660
         TabIndex        =   680
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   615
         Left            =   3420
         TabIndex        =   679
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   614
         Left            =   3180
         TabIndex        =   678
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   613
         Left            =   2940
         TabIndex        =   677
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   612
         Left            =   2700
         TabIndex        =   676
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   611
         Left            =   2460
         TabIndex        =   675
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   610
         Left            =   2220
         TabIndex        =   674
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   609
         Left            =   1980
         TabIndex        =   673
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   608
         Left            =   1740
         TabIndex        =   672
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   607
         Left            =   1500
         TabIndex        =   671
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   606
         Left            =   1260
         TabIndex        =   670
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   605
         Left            =   1020
         TabIndex        =   669
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   604
         Left            =   780
         TabIndex        =   668
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   603
         Left            =   540
         TabIndex        =   667
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   602
         Left            =   300
         TabIndex        =   666
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   601
         Left            =   60
         TabIndex        =   665
         Top             =   4860
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   600
         Left            =   7020
         TabIndex        =   664
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
         Index           =   599
         Left            =   6780
         TabIndex        =   663
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
         Index           =   598
         Left            =   6540
         TabIndex        =   662
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
         Index           =   597
         Left            =   6300
         TabIndex        =   661
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
         Index           =   596
         Left            =   6060
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
         Index           =   595
         Left            =   5820
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
         Index           =   594
         Left            =   5580
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
         Index           =   593
         Left            =   5340
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
         Index           =   592
         Left            =   5100
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
         Index           =   591
         Left            =   4860
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
         Index           =   590
         Left            =   4620
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
         Index           =   589
         Left            =   4380
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
         Index           =   588
         Left            =   4140
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
         Index           =   587
         Left            =   3900
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
         Index           =   586
         Left            =   3660
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
         Index           =   585
         Left            =   3420
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
         Index           =   584
         Left            =   3180
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
         Index           =   583
         Left            =   2940
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
         Index           =   582
         Left            =   2700
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
         Index           =   581
         Left            =   2460
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
         Index           =   580
         Left            =   2220
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
         Index           =   579
         Left            =   1980
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
         Index           =   578
         Left            =   1740
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
         Index           =   577
         Left            =   1500
         TabIndex        =   641
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
         Index           =   576
         Left            =   1260
         TabIndex        =   640
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
         Index           =   575
         Left            =   1020
         TabIndex        =   639
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
         Index           =   574
         Left            =   780
         TabIndex        =   638
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
         Index           =   573
         Left            =   540
         TabIndex        =   637
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
         Index           =   572
         Left            =   300
         TabIndex        =   636
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
         Index           =   571
         Left            =   60
         TabIndex        =   635
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
         Index           =   570
         Left            =   7020
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
         Index           =   569
         Left            =   6780
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
         Index           =   568
         Left            =   6540
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
         Index           =   567
         Left            =   6300
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
         Index           =   566
         Left            =   6060
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
         Index           =   565
         Left            =   5820
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
         Index           =   564
         Left            =   5580
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
         Index           =   563
         Left            =   5340
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
         Index           =   562
         Left            =   5100
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
         Index           =   561
         Left            =   4860
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
         Index           =   560
         Left            =   4620
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
         Index           =   559
         Left            =   4380
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
         Index           =   558
         Left            =   4140
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
         Index           =   557
         Left            =   3900
         TabIndex        =   621
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
         Index           =   556
         Left            =   3660
         TabIndex        =   620
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
         Index           =   555
         Left            =   3420
         TabIndex        =   619
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
         Index           =   554
         Left            =   3180
         TabIndex        =   618
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
         Index           =   553
         Left            =   2940
         TabIndex        =   617
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
         Index           =   552
         Left            =   2700
         TabIndex        =   616
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
         Index           =   551
         Left            =   2460
         TabIndex        =   615
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
         Index           =   550
         Left            =   2220
         TabIndex        =   614
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
         Index           =   549
         Left            =   1980
         TabIndex        =   613
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
         Index           =   548
         Left            =   1740
         TabIndex        =   612
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
         Index           =   547
         Left            =   1500
         TabIndex        =   611
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
         Index           =   546
         Left            =   1260
         TabIndex        =   610
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
         Index           =   545
         Left            =   1020
         TabIndex        =   609
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
         Index           =   544
         Left            =   780
         TabIndex        =   608
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
         Index           =   543
         Left            =   540
         TabIndex        =   607
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
         Index           =   542
         Left            =   300
         TabIndex        =   606
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
         Index           =   541
         Left            =   60
         TabIndex        =   605
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
         Index           =   540
         Left            =   7020
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
         Index           =   539
         Left            =   6780
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
         Index           =   538
         Left            =   6540
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
         Index           =   537
         Left            =   6300
         TabIndex        =   601
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
         Index           =   536
         Left            =   6060
         TabIndex        =   600
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
         Index           =   535
         Left            =   5820
         TabIndex        =   599
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
         Index           =   534
         Left            =   5580
         TabIndex        =   598
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
         Index           =   533
         Left            =   5340
         TabIndex        =   597
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
         Index           =   532
         Left            =   5100
         TabIndex        =   596
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
         Index           =   531
         Left            =   4860
         TabIndex        =   595
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
         Index           =   530
         Left            =   4620
         TabIndex        =   594
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
         Index           =   529
         Left            =   4380
         TabIndex        =   593
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
         Index           =   528
         Left            =   4140
         TabIndex        =   592
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
         Index           =   527
         Left            =   3900
         TabIndex        =   591
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
         Index           =   526
         Left            =   3660
         TabIndex        =   590
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
         Index           =   525
         Left            =   3420
         TabIndex        =   589
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
         Index           =   524
         Left            =   3180
         TabIndex        =   588
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
         Index           =   523
         Left            =   2940
         TabIndex        =   587
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
         Index           =   522
         Left            =   2700
         TabIndex        =   586
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
         Index           =   521
         Left            =   2460
         TabIndex        =   585
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
         Index           =   520
         Left            =   2220
         TabIndex        =   584
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
         Index           =   519
         Left            =   1980
         TabIndex        =   583
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
         Index           =   518
         Left            =   1740
         TabIndex        =   582
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
         Index           =   517
         Left            =   1500
         TabIndex        =   581
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
         Index           =   516
         Left            =   1260
         TabIndex        =   580
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
         Index           =   515
         Left            =   1020
         TabIndex        =   579
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
         Index           =   514
         Left            =   780
         TabIndex        =   578
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
         Index           =   513
         Left            =   540
         TabIndex        =   577
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
         Index           =   512
         Left            =   300
         TabIndex        =   576
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
         Index           =   511
         Left            =   60
         TabIndex        =   575
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
         Index           =   510
         Left            =   7020
         TabIndex        =   574
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
         Index           =   509
         Left            =   6780
         TabIndex        =   573
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
         Index           =   508
         Left            =   6540
         TabIndex        =   572
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
         Index           =   507
         Left            =   6300
         TabIndex        =   571
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
         Index           =   506
         Left            =   6060
         TabIndex        =   570
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
         Index           =   505
         Left            =   5820
         TabIndex        =   569
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
         Index           =   504
         Left            =   5580
         TabIndex        =   568
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
         Index           =   503
         Left            =   5340
         TabIndex        =   567
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
         Index           =   502
         Left            =   5100
         TabIndex        =   566
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
         Index           =   501
         Left            =   4860
         TabIndex        =   565
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
         Index           =   500
         Left            =   4620
         TabIndex        =   564
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
         Index           =   499
         Left            =   4380
         TabIndex        =   563
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
         Index           =   498
         Left            =   4140
         TabIndex        =   562
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
         Index           =   497
         Left            =   3900
         TabIndex        =   561
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
         Index           =   496
         Left            =   3660
         TabIndex        =   560
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
         Index           =   495
         Left            =   3420
         TabIndex        =   559
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
         Index           =   494
         Left            =   3180
         TabIndex        =   558
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
         Index           =   493
         Left            =   2940
         TabIndex        =   557
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
         Index           =   492
         Left            =   2700
         TabIndex        =   556
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
         Index           =   491
         Left            =   2460
         TabIndex        =   555
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
         Index           =   490
         Left            =   2220
         TabIndex        =   554
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
         Index           =   489
         Left            =   1980
         TabIndex        =   553
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
         Index           =   488
         Left            =   1740
         TabIndex        =   552
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
         Index           =   487
         Left            =   1500
         TabIndex        =   551
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
         Index           =   486
         Left            =   1260
         TabIndex        =   550
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
         Index           =   485
         Left            =   1020
         TabIndex        =   549
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
         Index           =   484
         Left            =   780
         TabIndex        =   548
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
         Index           =   483
         Left            =   540
         TabIndex        =   547
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
         Index           =   482
         Left            =   300
         TabIndex        =   546
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
         Index           =   481
         Left            =   60
         TabIndex        =   545
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
         Index           =   480
         Left            =   7020
         TabIndex        =   544
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
         Index           =   479
         Left            =   6780
         TabIndex        =   543
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
         Index           =   478
         Left            =   6540
         TabIndex        =   542
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
         Index           =   477
         Left            =   6300
         TabIndex        =   541
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
         Index           =   476
         Left            =   6060
         TabIndex        =   540
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
         Index           =   475
         Left            =   5820
         TabIndex        =   539
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
         Index           =   474
         Left            =   5580
         TabIndex        =   538
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
         Index           =   473
         Left            =   5340
         TabIndex        =   537
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
         Index           =   472
         Left            =   5100
         TabIndex        =   536
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
         Index           =   471
         Left            =   4860
         TabIndex        =   535
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
         Index           =   470
         Left            =   4620
         TabIndex        =   534
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
         Index           =   469
         Left            =   4380
         TabIndex        =   533
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
         Index           =   468
         Left            =   4140
         TabIndex        =   532
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
         Index           =   467
         Left            =   3900
         TabIndex        =   531
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
         Index           =   466
         Left            =   3660
         TabIndex        =   530
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
         Index           =   465
         Left            =   3420
         TabIndex        =   529
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
         Index           =   464
         Left            =   3180
         TabIndex        =   528
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
         Index           =   463
         Left            =   2940
         TabIndex        =   527
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
         Index           =   462
         Left            =   2700
         TabIndex        =   526
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
         Index           =   461
         Left            =   2460
         TabIndex        =   525
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
         Index           =   460
         Left            =   2220
         TabIndex        =   524
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
         Index           =   459
         Left            =   1980
         TabIndex        =   523
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
         Index           =   458
         Left            =   1740
         TabIndex        =   522
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
         Index           =   457
         Left            =   1500
         TabIndex        =   521
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
         Index           =   456
         Left            =   1260
         TabIndex        =   520
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
         Index           =   455
         Left            =   1020
         TabIndex        =   519
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
         Index           =   454
         Left            =   780
         TabIndex        =   518
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
         Index           =   453
         Left            =   540
         TabIndex        =   517
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
         Index           =   452
         Left            =   300
         TabIndex        =   516
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
         Index           =   451
         Left            =   60
         TabIndex        =   515
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
         Index           =   450
         Left            =   7020
         TabIndex        =   514
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
         Index           =   449
         Left            =   6780
         TabIndex        =   513
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
         Index           =   448
         Left            =   6540
         TabIndex        =   512
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
         Index           =   447
         Left            =   6300
         TabIndex        =   511
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
         Index           =   446
         Left            =   6060
         TabIndex        =   510
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
         Index           =   445
         Left            =   5820
         TabIndex        =   509
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
         Index           =   444
         Left            =   5580
         TabIndex        =   508
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
         Index           =   443
         Left            =   5340
         TabIndex        =   507
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
         Index           =   442
         Left            =   5100
         TabIndex        =   506
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
         Index           =   441
         Left            =   4860
         TabIndex        =   505
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
         Index           =   440
         Left            =   4620
         TabIndex        =   504
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
         Index           =   439
         Left            =   4380
         TabIndex        =   503
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
         Index           =   438
         Left            =   4140
         TabIndex        =   502
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
         Index           =   437
         Left            =   3900
         TabIndex        =   501
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
         Index           =   436
         Left            =   3660
         TabIndex        =   500
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
         Index           =   435
         Left            =   3420
         TabIndex        =   499
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
         Index           =   434
         Left            =   3180
         TabIndex        =   498
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
         Index           =   433
         Left            =   2940
         TabIndex        =   497
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
         Index           =   432
         Left            =   2700
         TabIndex        =   496
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
         Index           =   431
         Left            =   2460
         TabIndex        =   495
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
         Index           =   430
         Left            =   2220
         TabIndex        =   494
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
         Index           =   429
         Left            =   1980
         TabIndex        =   493
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
         Index           =   428
         Left            =   1740
         TabIndex        =   492
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
         Index           =   427
         Left            =   1500
         TabIndex        =   491
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
         Index           =   426
         Left            =   1260
         TabIndex        =   490
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
         Index           =   425
         Left            =   1020
         TabIndex        =   489
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
         Index           =   424
         Left            =   780
         TabIndex        =   488
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
         Index           =   423
         Left            =   540
         TabIndex        =   487
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
         Index           =   422
         Left            =   300
         TabIndex        =   486
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
         Index           =   421
         Left            =   60
         TabIndex        =   485
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
         Index           =   420
         Left            =   7020
         TabIndex        =   484
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
         Index           =   419
         Left            =   6780
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
         Index           =   418
         Left            =   6540
         TabIndex        =   482
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
         Index           =   417
         Left            =   6300
         TabIndex        =   481
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
         Index           =   416
         Left            =   6060
         TabIndex        =   480
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
         Index           =   415
         Left            =   5820
         TabIndex        =   479
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
         Index           =   414
         Left            =   5580
         TabIndex        =   478
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
         Index           =   413
         Left            =   5340
         TabIndex        =   477
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
         Index           =   412
         Left            =   5100
         TabIndex        =   476
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
         Index           =   411
         Left            =   4860
         TabIndex        =   475
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
         Index           =   410
         Left            =   4620
         TabIndex        =   474
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
         Index           =   409
         Left            =   4380
         TabIndex        =   473
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
         Index           =   408
         Left            =   4140
         TabIndex        =   472
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
         Index           =   407
         Left            =   3900
         TabIndex        =   471
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
         Index           =   406
         Left            =   3660
         TabIndex        =   470
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
         Index           =   405
         Left            =   3420
         TabIndex        =   469
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
         Index           =   404
         Left            =   3180
         TabIndex        =   468
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
         Index           =   403
         Left            =   2940
         TabIndex        =   467
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
         Index           =   402
         Left            =   2700
         TabIndex        =   466
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
         Index           =   401
         Left            =   2460
         TabIndex        =   465
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
         Index           =   400
         Left            =   2220
         TabIndex        =   464
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
         Index           =   399
         Left            =   1980
         TabIndex        =   463
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
         Index           =   398
         Left            =   1740
         TabIndex        =   462
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
         Index           =   397
         Left            =   1500
         TabIndex        =   461
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
         Index           =   396
         Left            =   1260
         TabIndex        =   460
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
         Index           =   395
         Left            =   1020
         TabIndex        =   459
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
         Index           =   394
         Left            =   780
         TabIndex        =   458
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
         Index           =   393
         Left            =   540
         TabIndex        =   457
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
         Index           =   392
         Left            =   300
         TabIndex        =   456
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
         Index           =   391
         Left            =   60
         TabIndex        =   455
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
         Index           =   390
         Left            =   7020
         TabIndex        =   454
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
         Index           =   389
         Left            =   6780
         TabIndex        =   453
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
         Index           =   388
         Left            =   6540
         TabIndex        =   452
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
         Index           =   387
         Left            =   6300
         TabIndex        =   451
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
         Index           =   386
         Left            =   6060
         TabIndex        =   450
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
         Index           =   385
         Left            =   5820
         TabIndex        =   449
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
         Index           =   384
         Left            =   5580
         TabIndex        =   448
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
         Index           =   383
         Left            =   5340
         TabIndex        =   447
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
         Index           =   382
         Left            =   5100
         TabIndex        =   446
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
         Index           =   381
         Left            =   4860
         TabIndex        =   445
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
         Index           =   380
         Left            =   4620
         TabIndex        =   444
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
         Index           =   379
         Left            =   4380
         TabIndex        =   443
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
         Index           =   378
         Left            =   4140
         TabIndex        =   442
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
         Index           =   377
         Left            =   3900
         TabIndex        =   441
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
         Index           =   376
         Left            =   3660
         TabIndex        =   440
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
         Index           =   375
         Left            =   3420
         TabIndex        =   439
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
         Index           =   374
         Left            =   3180
         TabIndex        =   438
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
         Index           =   373
         Left            =   2940
         TabIndex        =   437
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
         Index           =   372
         Left            =   2700
         TabIndex        =   436
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
         Index           =   371
         Left            =   2460
         TabIndex        =   435
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
         Index           =   370
         Left            =   2220
         TabIndex        =   434
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
         Index           =   369
         Left            =   1980
         TabIndex        =   433
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
         Index           =   368
         Left            =   1740
         TabIndex        =   432
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
         Index           =   367
         Left            =   1500
         TabIndex        =   431
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
         Index           =   366
         Left            =   1260
         TabIndex        =   430
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
         Index           =   365
         Left            =   1020
         TabIndex        =   429
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
         Index           =   364
         Left            =   780
         TabIndex        =   428
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
         Index           =   363
         Left            =   540
         TabIndex        =   427
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
         Index           =   362
         Left            =   300
         TabIndex        =   426
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
         Index           =   361
         Left            =   60
         TabIndex        =   425
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
         Index           =   360
         Left            =   7020
         TabIndex        =   424
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
         Index           =   359
         Left            =   6780
         TabIndex        =   423
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
         Index           =   358
         Left            =   6540
         TabIndex        =   422
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
         Index           =   357
         Left            =   6300
         TabIndex        =   421
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
         Index           =   356
         Left            =   6060
         TabIndex        =   420
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
         Index           =   355
         Left            =   5820
         TabIndex        =   419
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
         Index           =   354
         Left            =   5580
         TabIndex        =   418
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
         Index           =   353
         Left            =   5340
         TabIndex        =   417
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
         Index           =   352
         Left            =   5100
         TabIndex        =   416
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
         Index           =   351
         Left            =   4860
         TabIndex        =   415
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
         Index           =   350
         Left            =   4620
         TabIndex        =   414
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
         Index           =   349
         Left            =   4380
         TabIndex        =   413
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
         Index           =   348
         Left            =   4140
         TabIndex        =   412
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
         Index           =   347
         Left            =   3900
         TabIndex        =   411
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
         Index           =   346
         Left            =   3660
         TabIndex        =   410
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
         Index           =   345
         Left            =   3420
         TabIndex        =   409
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
         Index           =   344
         Left            =   3180
         TabIndex        =   408
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
         Index           =   343
         Left            =   2940
         TabIndex        =   407
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
         Index           =   342
         Left            =   2700
         TabIndex        =   406
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
         Index           =   341
         Left            =   2460
         TabIndex        =   405
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
         Index           =   340
         Left            =   2220
         TabIndex        =   404
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
         Index           =   339
         Left            =   1980
         TabIndex        =   403
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
         Index           =   338
         Left            =   1740
         TabIndex        =   402
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
         Index           =   337
         Left            =   1500
         TabIndex        =   401
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
         Index           =   336
         Left            =   1260
         TabIndex        =   400
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
         Index           =   335
         Left            =   1020
         TabIndex        =   399
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
         Index           =   334
         Left            =   780
         TabIndex        =   398
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
         Index           =   333
         Left            =   540
         TabIndex        =   397
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
         Index           =   332
         Left            =   300
         TabIndex        =   396
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
         Index           =   331
         Left            =   60
         TabIndex        =   395
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
         Index           =   330
         Left            =   7020
         TabIndex        =   394
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
         Index           =   329
         Left            =   6780
         TabIndex        =   393
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
         Index           =   328
         Left            =   6540
         TabIndex        =   392
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
         Index           =   327
         Left            =   6300
         TabIndex        =   391
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
         Index           =   326
         Left            =   6060
         TabIndex        =   390
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
         Index           =   325
         Left            =   5820
         TabIndex        =   389
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
         Index           =   324
         Left            =   5580
         TabIndex        =   388
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
         Index           =   323
         Left            =   5340
         TabIndex        =   387
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
         Index           =   322
         Left            =   5100
         TabIndex        =   386
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
         Index           =   321
         Left            =   4860
         TabIndex        =   385
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
         Index           =   320
         Left            =   4620
         TabIndex        =   384
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
         Index           =   319
         Left            =   4380
         TabIndex        =   383
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
         Index           =   318
         Left            =   4140
         TabIndex        =   382
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
         Index           =   317
         Left            =   3900
         TabIndex        =   381
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
         Index           =   316
         Left            =   3660
         TabIndex        =   380
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
         Index           =   315
         Left            =   3420
         TabIndex        =   379
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
         Index           =   314
         Left            =   3180
         TabIndex        =   378
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
         Index           =   313
         Left            =   2940
         TabIndex        =   377
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
         Index           =   312
         Left            =   2700
         TabIndex        =   376
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
         Index           =   311
         Left            =   2460
         TabIndex        =   375
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
         Index           =   310
         Left            =   2220
         TabIndex        =   374
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
         Index           =   309
         Left            =   1980
         TabIndex        =   373
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
         Index           =   308
         Left            =   1740
         TabIndex        =   372
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
         Index           =   307
         Left            =   1500
         TabIndex        =   371
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
         Index           =   306
         Left            =   1260
         TabIndex        =   370
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
         Index           =   305
         Left            =   1020
         TabIndex        =   369
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
         Index           =   304
         Left            =   780
         TabIndex        =   368
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
         Index           =   303
         Left            =   540
         TabIndex        =   367
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
         Index           =   302
         Left            =   300
         TabIndex        =   366
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
         Index           =   301
         Left            =   60
         TabIndex        =   365
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
         Index           =   300
         Left            =   7020
         TabIndex        =   364
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
         Index           =   299
         Left            =   6780
         TabIndex        =   363
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
         Index           =   298
         Left            =   6540
         TabIndex        =   362
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
         Index           =   297
         Left            =   6300
         TabIndex        =   361
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
         Index           =   296
         Left            =   6060
         TabIndex        =   360
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
         Index           =   295
         Left            =   5820
         TabIndex        =   359
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
         Index           =   294
         Left            =   5580
         TabIndex        =   358
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
         Index           =   293
         Left            =   5340
         TabIndex        =   357
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
         Index           =   292
         Left            =   5100
         TabIndex        =   356
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
         Index           =   291
         Left            =   4860
         TabIndex        =   355
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
         Index           =   290
         Left            =   4620
         TabIndex        =   354
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
         Index           =   289
         Left            =   4380
         TabIndex        =   353
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
         Index           =   288
         Left            =   4140
         TabIndex        =   352
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
         Index           =   287
         Left            =   3900
         TabIndex        =   351
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
         Index           =   286
         Left            =   3660
         TabIndex        =   350
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
         Index           =   285
         Left            =   3420
         TabIndex        =   349
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
         Index           =   284
         Left            =   3180
         TabIndex        =   348
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
         Index           =   283
         Left            =   2940
         TabIndex        =   347
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
         Index           =   282
         Left            =   2700
         TabIndex        =   346
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
         Index           =   281
         Left            =   2460
         TabIndex        =   345
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
         Index           =   280
         Left            =   2220
         TabIndex        =   344
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
         Index           =   279
         Left            =   1980
         TabIndex        =   343
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
         Index           =   278
         Left            =   1740
         TabIndex        =   342
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
         Index           =   277
         Left            =   1500
         TabIndex        =   341
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
         Index           =   276
         Left            =   1260
         TabIndex        =   340
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
         Index           =   275
         Left            =   1020
         TabIndex        =   339
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
         Index           =   274
         Left            =   780
         TabIndex        =   338
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
         Index           =   273
         Left            =   540
         TabIndex        =   337
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
         Index           =   272
         Left            =   300
         TabIndex        =   336
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
         Index           =   271
         Left            =   60
         TabIndex        =   335
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
         Index           =   270
         Left            =   7020
         TabIndex        =   334
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
         Index           =   269
         Left            =   6780
         TabIndex        =   333
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
         Index           =   268
         Left            =   6540
         TabIndex        =   332
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
         Index           =   267
         Left            =   6300
         TabIndex        =   331
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
         Index           =   266
         Left            =   6060
         TabIndex        =   330
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
         Index           =   265
         Left            =   5820
         TabIndex        =   329
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
         Index           =   264
         Left            =   5580
         TabIndex        =   328
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
         Index           =   263
         Left            =   5340
         TabIndex        =   327
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
         Index           =   262
         Left            =   5100
         TabIndex        =   326
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
         Index           =   261
         Left            =   4860
         TabIndex        =   325
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
         Index           =   260
         Left            =   4620
         TabIndex        =   324
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
         Index           =   259
         Left            =   4380
         TabIndex        =   323
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
         Index           =   258
         Left            =   4140
         TabIndex        =   322
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
         Index           =   257
         Left            =   3900
         TabIndex        =   321
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
         Index           =   256
         Left            =   3660
         TabIndex        =   320
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
         Index           =   255
         Left            =   3420
         TabIndex        =   319
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
         Index           =   254
         Left            =   3180
         TabIndex        =   318
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
         Index           =   253
         Left            =   2940
         TabIndex        =   317
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
         Index           =   252
         Left            =   2700
         TabIndex        =   316
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
         Index           =   251
         Left            =   2460
         TabIndex        =   315
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
         Index           =   250
         Left            =   2220
         TabIndex        =   314
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
         Index           =   249
         Left            =   1980
         TabIndex        =   313
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
         Index           =   248
         Left            =   1740
         TabIndex        =   312
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
         Index           =   247
         Left            =   1500
         TabIndex        =   311
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
         Index           =   246
         Left            =   1260
         TabIndex        =   310
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
         Index           =   245
         Left            =   1020
         TabIndex        =   309
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
         Index           =   244
         Left            =   780
         TabIndex        =   308
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
         Index           =   243
         Left            =   540
         TabIndex        =   307
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
         Index           =   242
         Left            =   300
         TabIndex        =   306
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
         Index           =   241
         Left            =   60
         TabIndex        =   305
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
         Index           =   240
         Left            =   7020
         TabIndex        =   304
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
         Index           =   239
         Left            =   6780
         TabIndex        =   303
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
         Index           =   238
         Left            =   6540
         TabIndex        =   302
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
         Index           =   237
         Left            =   6300
         TabIndex        =   301
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
         Index           =   236
         Left            =   6060
         TabIndex        =   300
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
         Index           =   235
         Left            =   5820
         TabIndex        =   299
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
         Index           =   234
         Left            =   5580
         TabIndex        =   298
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
         Index           =   233
         Left            =   5340
         TabIndex        =   297
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
         Index           =   232
         Left            =   5100
         TabIndex        =   296
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
         Index           =   231
         Left            =   4860
         TabIndex        =   295
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
         Index           =   230
         Left            =   4620
         TabIndex        =   294
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
         Index           =   229
         Left            =   4380
         TabIndex        =   293
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
         Index           =   228
         Left            =   4140
         TabIndex        =   292
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
         Index           =   227
         Left            =   3900
         TabIndex        =   291
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
         Index           =   226
         Left            =   3660
         TabIndex        =   290
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
         Index           =   225
         Left            =   3420
         TabIndex        =   289
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
         Index           =   224
         Left            =   3180
         TabIndex        =   288
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
         Index           =   223
         Left            =   2940
         TabIndex        =   287
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
         Index           =   222
         Left            =   2700
         TabIndex        =   286
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
         Index           =   221
         Left            =   2460
         TabIndex        =   285
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
         Index           =   220
         Left            =   2220
         TabIndex        =   284
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
         Index           =   219
         Left            =   1980
         TabIndex        =   283
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
         Index           =   218
         Left            =   1740
         TabIndex        =   282
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
         Index           =   217
         Left            =   1500
         TabIndex        =   281
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
         Index           =   216
         Left            =   1260
         TabIndex        =   280
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
         Index           =   215
         Left            =   1020
         TabIndex        =   279
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
         Index           =   214
         Left            =   780
         TabIndex        =   278
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
         Index           =   213
         Left            =   540
         TabIndex        =   277
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
         Index           =   212
         Left            =   300
         TabIndex        =   276
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
         Index           =   211
         Left            =   60
         TabIndex        =   275
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
         Index           =   210
         Left            =   7020
         TabIndex        =   274
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
         Index           =   209
         Left            =   6780
         TabIndex        =   273
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
         Index           =   208
         Left            =   6540
         TabIndex        =   272
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
         Index           =   207
         Left            =   6300
         TabIndex        =   271
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
         Index           =   206
         Left            =   6060
         TabIndex        =   270
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
         Index           =   205
         Left            =   5820
         TabIndex        =   269
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
         Index           =   204
         Left            =   5580
         TabIndex        =   268
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
         Index           =   203
         Left            =   5340
         TabIndex        =   267
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
         Index           =   202
         Left            =   5100
         TabIndex        =   266
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
         Index           =   201
         Left            =   4860
         TabIndex        =   265
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
         Index           =   200
         Left            =   4620
         TabIndex        =   264
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
         Index           =   199
         Left            =   4380
         TabIndex        =   263
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
         Index           =   198
         Left            =   4140
         TabIndex        =   262
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
         Index           =   197
         Left            =   3900
         TabIndex        =   261
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
         Index           =   196
         Left            =   3660
         TabIndex        =   260
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
         Index           =   195
         Left            =   3420
         TabIndex        =   259
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
         Index           =   194
         Left            =   3180
         TabIndex        =   258
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
         Index           =   193
         Left            =   2940
         TabIndex        =   257
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
         Index           =   192
         Left            =   2700
         TabIndex        =   256
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
         Index           =   191
         Left            =   2460
         TabIndex        =   255
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
         Index           =   190
         Left            =   2220
         TabIndex        =   254
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
         Index           =   189
         Left            =   1980
         TabIndex        =   253
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
         Index           =   188
         Left            =   1740
         TabIndex        =   252
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
         Index           =   187
         Left            =   1500
         TabIndex        =   251
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
         Index           =   186
         Left            =   1260
         TabIndex        =   250
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
         Index           =   185
         Left            =   1020
         TabIndex        =   249
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
         Index           =   184
         Left            =   780
         TabIndex        =   248
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
         Index           =   183
         Left            =   540
         TabIndex        =   247
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
         Index           =   182
         Left            =   300
         TabIndex        =   246
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
         Index           =   181
         Left            =   60
         TabIndex        =   245
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
         Index           =   180
         Left            =   7020
         TabIndex        =   244
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
         Index           =   179
         Left            =   6780
         TabIndex        =   243
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
         Index           =   178
         Left            =   6540
         TabIndex        =   242
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
         Index           =   177
         Left            =   6300
         TabIndex        =   241
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
         Index           =   176
         Left            =   6060
         TabIndex        =   240
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
         Index           =   175
         Left            =   5820
         TabIndex        =   239
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
         Index           =   174
         Left            =   5580
         TabIndex        =   238
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
         Index           =   173
         Left            =   5340
         TabIndex        =   237
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
         Index           =   172
         Left            =   5100
         TabIndex        =   236
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
         Index           =   171
         Left            =   4860
         TabIndex        =   235
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
         Index           =   170
         Left            =   4620
         TabIndex        =   234
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
         Index           =   169
         Left            =   4380
         TabIndex        =   233
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
         Index           =   168
         Left            =   4140
         TabIndex        =   232
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
         Index           =   167
         Left            =   3900
         TabIndex        =   231
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
         Index           =   166
         Left            =   3660
         TabIndex        =   230
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
         Index           =   165
         Left            =   3420
         TabIndex        =   229
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
         Index           =   164
         Left            =   3180
         TabIndex        =   228
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
         Index           =   163
         Left            =   2940
         TabIndex        =   227
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
         Index           =   162
         Left            =   2700
         TabIndex        =   226
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
         Index           =   161
         Left            =   2460
         TabIndex        =   225
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
         Index           =   160
         Left            =   2220
         TabIndex        =   224
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
         Index           =   159
         Left            =   1980
         TabIndex        =   223
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
         Index           =   158
         Left            =   1740
         TabIndex        =   222
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
         Index           =   157
         Left            =   1500
         TabIndex        =   221
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
         Index           =   156
         Left            =   1260
         TabIndex        =   220
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
         Index           =   155
         Left            =   1020
         TabIndex        =   219
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
         Index           =   154
         Left            =   780
         TabIndex        =   218
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
         Index           =   153
         Left            =   540
         TabIndex        =   217
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
         Index           =   152
         Left            =   300
         TabIndex        =   216
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
         Index           =   151
         Left            =   60
         TabIndex        =   215
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
         Index           =   150
         Left            =   7020
         TabIndex        =   214
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
         Index           =   149
         Left            =   6780
         TabIndex        =   213
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
         Index           =   148
         Left            =   6540
         TabIndex        =   212
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
         Index           =   147
         Left            =   6300
         TabIndex        =   211
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
         Index           =   146
         Left            =   6060
         TabIndex        =   210
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
         Index           =   145
         Left            =   5820
         TabIndex        =   209
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
         Index           =   144
         Left            =   5580
         TabIndex        =   208
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
         Index           =   143
         Left            =   5340
         TabIndex        =   207
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
         Index           =   142
         Left            =   5100
         TabIndex        =   206
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
         Index           =   141
         Left            =   4860
         TabIndex        =   205
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
         Index           =   140
         Left            =   4620
         TabIndex        =   204
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
         Index           =   139
         Left            =   4380
         TabIndex        =   203
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
         Index           =   138
         Left            =   4140
         TabIndex        =   202
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
         Index           =   137
         Left            =   3900
         TabIndex        =   201
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
         Index           =   136
         Left            =   3660
         TabIndex        =   200
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
         Index           =   135
         Left            =   3420
         TabIndex        =   199
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
         Index           =   134
         Left            =   3180
         TabIndex        =   198
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
         Index           =   133
         Left            =   2940
         TabIndex        =   197
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
         Index           =   132
         Left            =   2700
         TabIndex        =   196
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
         Index           =   131
         Left            =   2460
         TabIndex        =   195
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
         Index           =   130
         Left            =   2220
         TabIndex        =   194
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
         Index           =   129
         Left            =   1980
         TabIndex        =   193
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
         Index           =   128
         Left            =   1740
         TabIndex        =   192
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
         Index           =   127
         Left            =   1500
         TabIndex        =   191
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
         Index           =   126
         Left            =   1260
         TabIndex        =   190
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
         Index           =   125
         Left            =   1020
         TabIndex        =   189
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
         Index           =   124
         Left            =   780
         TabIndex        =   188
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
         Index           =   123
         Left            =   540
         TabIndex        =   187
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
         Index           =   122
         Left            =   300
         TabIndex        =   186
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
         Index           =   121
         Left            =   60
         TabIndex        =   185
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
         Index           =   120
         Left            =   7020
         TabIndex        =   184
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
         Index           =   119
         Left            =   6780
         TabIndex        =   183
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
         Index           =   118
         Left            =   6540
         TabIndex        =   182
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
         Index           =   117
         Left            =   6300
         TabIndex        =   181
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
         Index           =   116
         Left            =   6060
         TabIndex        =   180
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
         Index           =   115
         Left            =   5820
         TabIndex        =   179
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
         Index           =   114
         Left            =   5580
         TabIndex        =   178
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
         Index           =   113
         Left            =   5340
         TabIndex        =   177
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
         Index           =   112
         Left            =   5100
         TabIndex        =   176
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
         Index           =   111
         Left            =   4860
         TabIndex        =   175
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
         Index           =   110
         Left            =   4620
         TabIndex        =   174
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
         Index           =   109
         Left            =   4380
         TabIndex        =   173
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
         Index           =   108
         Left            =   4140
         TabIndex        =   172
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
         Index           =   107
         Left            =   3900
         TabIndex        =   171
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
         Index           =   106
         Left            =   3660
         TabIndex        =   170
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
         Index           =   105
         Left            =   3420
         TabIndex        =   169
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
         Index           =   104
         Left            =   3180
         TabIndex        =   168
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
         Index           =   103
         Left            =   2940
         TabIndex        =   167
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
         Index           =   102
         Left            =   2700
         TabIndex        =   166
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
         Index           =   101
         Left            =   2460
         TabIndex        =   165
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
         Index           =   100
         Left            =   2220
         TabIndex        =   164
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
         Index           =   99
         Left            =   1980
         TabIndex        =   163
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
         Index           =   98
         Left            =   1740
         TabIndex        =   162
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
         Index           =   97
         Left            =   1500
         TabIndex        =   161
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
         Index           =   96
         Left            =   1260
         TabIndex        =   160
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
         Index           =   95
         Left            =   1020
         TabIndex        =   159
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
         Index           =   94
         Left            =   780
         TabIndex        =   158
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
         Index           =   93
         Left            =   540
         TabIndex        =   157
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
         Index           =   92
         Left            =   300
         TabIndex        =   156
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
         Index           =   91
         Left            =   60
         TabIndex        =   155
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
         Index           =   90
         Left            =   7020
         TabIndex        =   154
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
         Index           =   89
         Left            =   6780
         TabIndex        =   153
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
         Index           =   88
         Left            =   6540
         TabIndex        =   152
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
         Index           =   87
         Left            =   6300
         TabIndex        =   151
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
         Index           =   86
         Left            =   6060
         TabIndex        =   150
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
         Index           =   85
         Left            =   5820
         TabIndex        =   149
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
         Index           =   84
         Left            =   5580
         TabIndex        =   148
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
         Index           =   83
         Left            =   5340
         TabIndex        =   147
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
         Index           =   82
         Left            =   5100
         TabIndex        =   146
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
         Index           =   81
         Left            =   4860
         TabIndex        =   145
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
         Index           =   80
         Left            =   4620
         TabIndex        =   144
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
         Index           =   79
         Left            =   4380
         TabIndex        =   143
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
         Index           =   78
         Left            =   4140
         TabIndex        =   142
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
         Index           =   77
         Left            =   3900
         TabIndex        =   141
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
         Index           =   76
         Left            =   3660
         TabIndex        =   140
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
         Index           =   75
         Left            =   3420
         TabIndex        =   139
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
         Index           =   74
         Left            =   3180
         TabIndex        =   138
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
         Index           =   73
         Left            =   2940
         TabIndex        =   137
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
         Index           =   72
         Left            =   2700
         TabIndex        =   136
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
         Index           =   71
         Left            =   2460
         TabIndex        =   135
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
         Index           =   70
         Left            =   2220
         TabIndex        =   134
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
         Index           =   69
         Left            =   1980
         TabIndex        =   133
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
         Index           =   68
         Left            =   1740
         TabIndex        =   132
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
         Index           =   67
         Left            =   1500
         TabIndex        =   131
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
         Index           =   66
         Left            =   1260
         TabIndex        =   130
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
         Index           =   65
         Left            =   1020
         TabIndex        =   129
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
         Index           =   64
         Left            =   780
         TabIndex        =   128
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
         Index           =   63
         Left            =   540
         TabIndex        =   127
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
         Index           =   62
         Left            =   300
         TabIndex        =   126
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
         TabIndex        =   125
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
         Left            =   7020
         TabIndex        =   124
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
         Index           =   59
         Left            =   6780
         TabIndex        =   123
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
         Index           =   58
         Left            =   6540
         TabIndex        =   122
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
         Index           =   57
         Left            =   6300
         TabIndex        =   121
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
         Index           =   56
         Left            =   6060
         TabIndex        =   120
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
         Index           =   55
         Left            =   5820
         TabIndex        =   119
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
         Index           =   54
         Left            =   5580
         TabIndex        =   118
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
         Index           =   53
         Left            =   5340
         TabIndex        =   117
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
         Index           =   52
         Left            =   5100
         TabIndex        =   116
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
         Index           =   51
         Left            =   4860
         TabIndex        =   115
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
         Index           =   50
         Left            =   4620
         TabIndex        =   114
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
         Index           =   49
         Left            =   4380
         TabIndex        =   113
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
         Index           =   48
         Left            =   4140
         TabIndex        =   112
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
         Index           =   47
         Left            =   3900
         TabIndex        =   111
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
         Index           =   46
         Left            =   3660
         TabIndex        =   110
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
         Index           =   45
         Left            =   3420
         TabIndex        =   109
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
         Index           =   44
         Left            =   3180
         TabIndex        =   108
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
         Index           =   43
         Left            =   2940
         TabIndex        =   107
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
         Index           =   42
         Left            =   2700
         TabIndex        =   106
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
         Left            =   2460
         TabIndex        =   105
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
         Left            =   2220
         TabIndex        =   104
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
         Left            =   1980
         TabIndex        =   103
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
         Left            =   1740
         TabIndex        =   102
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
         Left            =   1500
         TabIndex        =   101
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
         Left            =   1260
         TabIndex        =   100
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
         Left            =   1020
         TabIndex        =   99
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
         Left            =   780
         TabIndex        =   98
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
         Left            =   540
         TabIndex        =   97
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
         Left            =   300
         TabIndex        =   96
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
         Left            =   60
         TabIndex        =   95
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
         Left            =   7020
         TabIndex        =   91
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
         Index           =   29
         Left            =   6780
         TabIndex        =   90
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
         Index           =   28
         Left            =   6540
         TabIndex        =   89
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
         Index           =   27
         Left            =   6300
         TabIndex        =   88
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
         Index           =   26
         Left            =   6060
         TabIndex        =   87
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
         Index           =   25
         Left            =   5820
         TabIndex        =   86
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
         Index           =   24
         Left            =   5580
         TabIndex        =   85
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
         Index           =   23
         Left            =   5340
         TabIndex        =   84
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
         Index           =   22
         Left            =   5100
         TabIndex        =   83
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
         Left            =   4860
         TabIndex        =   82
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
         TabIndex        =   81
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
         TabIndex        =   80
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
         TabIndex        =   79
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
         TabIndex        =   78
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
         TabIndex        =   77
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
         TabIndex        =   76
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
         TabIndex        =   75
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
         TabIndex        =   74
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
         TabIndex        =   73
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
         TabIndex        =   72
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
         TabIndex        =   71
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
         TabIndex        =   70
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
         TabIndex        =   69
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
         TabIndex        =   68
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
         TabIndex        =   67
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
         TabIndex        =   66
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
         TabIndex        =   65
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
         TabIndex        =   64
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
         TabIndex        =   63
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
         Index           =   1
         Left            =   60
         TabIndex        =   62
         Top             =   60
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin MSComctlLib.ListView lvMapLoc 
      Height          =   1035
      Left            =   60
      TabIndex        =   61
      Top             =   7680
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   1826
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   14737632
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
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
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

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
Dim CellRoom(1 To 900, 1 To 2) As Long
Dim UnchartedCells(1 To 900) As Integer
Dim StopBuild As Boolean

Dim TTlbl As clsToolTip

Private Sub cmbMapSize_Click()
Select Case cmbMapSize.ListIndex
    Case 1: '30x30
        nMapCenterCell = 436
    Case Else: '20x20
        nMapCenterCell = 280
        
End Select
End Sub

Private Sub Form_Activate()
If chkMapOptions(6).Value = 0 Then Call SetTopMostWindow(Me.hWnd, True)
End Sub

Private Sub Form_Load()
On Error GoTo Error:
Dim lR As Long

Set TTlbl = New clsToolTip

With TTlbl
    .DelayTime = 20
    .VisibleTime = 20000
    .BkColor = &HC0FFFF
    .TxtColor = &H0
    .Style = ttStyleStandard
    '.Style = ttStyleStandard
End With

If Not ReadINI("Settings", "MapExternalOnTop") = "1" Then
    lR = SetTopMostWindow(Me.hWnd, True)
Else
    chkMapOptions(6).Value = 1
End If

lvMapLoc.ColumnHeaders.clear
lvMapLoc.ColumnHeaders.Add 1, "References", "References", 4500

bMapSwapButtons = frmMain.bMapSwapButtons

Me.Top = ReadINI("Settings", "ExMapTop")
Me.Left = ReadINI("Settings", "ExMapLeft")

chkMapOptions(0).Value = ReadINI("Settings", "ExMapFollowMap")
chkMapOptions(1).Value = ReadINI("Settings", "ExMapNoHidden")
chkMapOptions(2).Value = ReadINI("Settings", "ExMapNoLairs")
chkMapOptions(3).Value = ReadINI("Settings", "ExMapNoNPC")
chkMapOptions(4).Value = ReadINI("Settings", "ExMapNoCMD")
chkMapOptions(5).Value = ReadINI("Settings", "ExMapNoTooltips")
chkMapOptions(8).Value = ReadINI("Settings", "ExMapMainOverlap")

Call LoadPresets

cmbMapSize.ListIndex = Val(ReadINI("Settings", "ExMapSize"))

Call ResizeMap

Exit Sub
Error:
Call HandleError("Form_Load")
Resume Next
End Sub

Private Sub chkMapOptions_Click(Index As Integer)
Dim lR As Long

If Index = 6 Then
    If chkMapOptions(6).Value = 1 Then
        lR = SetTopMostWindow(Me.hWnd, False)
    Else
        lR = SetTopMostWindow(Me.hWnd, True)
    End If
    If FormIsLoaded("frmResults") Then
        If frmResults.objFormOwner Is Me Then
            If chkMapOptions(6).Value = 1 Then
                lR = SetTopMostWindow(frmResults.hWnd, False)
            Else
                lR = SetTopMostWindow(frmResults.hWnd, True)
            End If
        End If
    End If
ElseIf Index = 7 Then
    If chkMapOptions(7).Value = 1 Then
        fraMapControls.Visible = True
    Else
        fraMapControls.Visible = False
    End If
End If

End Sub

Private Sub cmdMove_Click(Index As Integer)
On Error GoTo Error:
Dim sLook As String, RoomExit As RoomExitType
Dim nExitType As Integer, nRecNum As Long

tabRooms.Index = "idxRooms"
tabRooms.Seek "=", nMapStartMap, nMapStartRoom
If tabRooms.NoMatch Then GoTo out:

Select Case Index
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

Call MapStartMapping(RoomExit.Map, RoomExit.Room)

out:
Exit Sub
Error:
Call HandleError("cmdMove_Click")
Resume out:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If Me.ActiveControl Is txtRoomMap Then
    Exit Sub
ElseIf Me.ActiveControl Is txtRoomRoom Then
    Exit Sub
End If

Select Case KeyAscii
    Case 46: 'd
        Call cmdMove_Click(9)
    Case 48: 'u
        Call cmdMove_Click(8)
    Case 49: 'sw
        Call cmdMove_Click(7)
    Case 50: 's
        Call cmdMove_Click(1)
    Case 51: 'se
        Call cmdMove_Click(6)
    Case 52: 'w
        Call cmdMove_Click(3)
    'Case 53:
    Case 54: 'e
        Call cmdMove_Click(2)
    Case 55: 'nw
        Call cmdMove_Click(5)
    Case 56: 'n
        Call cmdMove_Click(0)
    Case 57: 'ne
        Call cmdMove_Click(4)
End Select

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

Private Sub fraMapControls_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
fraMapControls.Top = y
fraMapControls.Left = x
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

Private Sub cmdMapPresetSelect_Click(Index As Integer)
Dim nStart As Integer, x As Integer, sSectionName As String
Dim cReg As clsRegistryRoutines

Set cReg = New clsRegistryRoutines

If InStr(1, frmMain.lblDatVer.Caption, "-") = 0 Then
    sSectionName = "Custom_Presets"
Else
    sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ") & "_Presets"
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
    cmdMapPreset(x Mod 10).Caption = cReg.GetRegistryValue("Name" & x, "unset")
    cmdMapPreset(x Mod 10).Tag = x
Next x
End Sub

Private Sub cmdQ_Click(Index As Integer)

Select Case Index
    Case 0:
        MsgBox "Clicking ""Options"" again will hide the options window and refresh the map.", vbInformation
    Case 1:
        MsgBox "You can also use your keypad to move around on the map.", vbInformation
    Case 2:
        MsgBox "This will allow the 'Main' MMUD Explorer window to overlap the" & vbCrLf _
            & "map window (when set to 'Always on Top') when double clicking one" & vbCrLf _
            & "of the references below.  (Click the Map window again to re-activate" & vbCrLf _
            & "the 'Always on Top' functionality.)", vbInformation
End Select

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

fraOptions.Visible = False
Me.Enabled = False
frmMain.Enabled = False

bMapCancelFind = False

If chkMapOptions(6).Value = 0 Then Call SetTopMostWindow(Me.hWnd, False)

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
On Error Resume Next
Unload frmProgressBar
Me.Enabled = True
If chkMapOptions(6).Value = 0 Then Call SetTopMostWindow(Me.hWnd, True)
frmMain.Enabled = True
Me.SetFocus
Exit Sub

Error:
Call HandleError("cmdMapFindText_Click")
Resume out:
End Sub

Private Sub cmdMapShowUnused_Click()
Dim x As Integer

If cmdMapShowUnused.Caption = "S&how Unused Blocks" Then
    For x = 1 To 900
        lblRoomCell(x).Visible = True
    Next
    cmdMapShowUnused.Caption = "&Hide Unused Blocks"
Else
    For x = 1 To 900
        If CellRoom(x, 1) = 0 Then lblRoomCell(x).Visible = False
    Next
    cmdMapShowUnused.Caption = "S&how Unused Blocks"
End If

fraOptions.Visible = False
End Sub

Private Sub cmdViewMapLegend_Click()
On Error GoTo Error:

If cmdViewMapLegend.Tag = "1" Then
    Unload frmMapLegend
    cmdViewMapLegend.Tag = "0"
Else
    cmdViewMapLegend.Tag = "1"
    frmMapLegend.Show vbModeless, Me
    Set frmMapLegend.objFormOwner = Me
    
    If chkMapOptions(6).Value = 0 Then Call SetTopMostWindow(Me.hWnd, True)
    
    'Call SetOwner(frmMapLegend.hwnd, Me.hwnd)
'    If chkMapOptions(6).Value = 1 Then
'        lR = SetTopMostWindow(frmMap.hwnd, False)
'    Else
'        lR = SetTopMostWindow(frmMap.hwnd, True)
'    End If
End If
'fraOptions.Visible = False

Exit Sub

Error:
Call HandleError("cmdViewMapLegend_Click")
End Sub


Private Sub cmdDrawMap_Click(Index As Integer)
fraOptions.Visible = False
If Index = 0 Then
    If Val(txtRoomMap.Text) > 32767 Then txtRoomMap.Text = 32767
    If Val(txtRoomRoom.Text) > 32767 Then txtRoomRoom.Text = 32767
    Call MapStartMapping(Val(txtRoomMap.Text), Val(txtRoomRoom.Text))
Else
    Call MapStartMapping(nMapLastMap, nMapLastRoom)
End If
End Sub

Private Sub ResizeMap()
On Error GoTo Error:

If Me.WindowState = vbMinimized Then Me.WindowState = vbNormal

Select Case cmbMapSize.ListIndex
    Case 1: '30x30
        sMapSECorner = 900
        nMapRowLength = 30
        If nMapCenterCell > sMapSECorner Then nMapCenterCell = 436
        If nMapCenterCell = 0 Then nMapCenterCell = 436
        Me.Height = 9135 + TITLEBAR_OFFSET
        Me.Width = 7440
        lvMapLoc.Top = 7680
        lvMapLoc.Width = 7245
        picMap.Width = 7245
        picMap.Height = 7245
        lvMapLoc.ColumnHeaders(1).Width = 6800
        
    Case Else: '20x20
        sMapSECorner = 590
        nMapRowLength = 20
        If nMapCenterCell > sMapSECorner Then nMapCenterCell = 280
        If nMapCenterCell = 0 Then nMapCenterCell = 280
        Me.Height = 6735 + TITLEBAR_OFFSET
        Me.Width = 5055
        lvMapLoc.Top = 5280
        lvMapLoc.Width = 4845
        picMap.Width = 4845
        picMap.Height = 4845
        lvMapLoc.ColumnHeaders(1).Width = 4400
        
End Select

out:
Exit Sub
Error:
Call HandleError("ResizeMap")
Resume out:

End Sub
Public Sub MapStartMapping(ByVal nStartMap As Long, ByVal nStartRoom As Long, Optional nCenterCell As Integer)
On Error GoTo Error:
Dim x As Integer, nMapSize As Integer, bCheckAgain As Boolean, y As Integer

If bMapStillMapping Then Exit Sub

tabRooms.Index = "idxRooms"
tabRooms.Seek "=", nStartMap, nStartRoom
If tabRooms.NoMatch Then
    MsgBox "Room " & nStartMap & "/" & nStartRoom & " was not found.", vbInformation
    'Me.Caption = "Rooms"
    Exit Sub
Else
    Me.Caption = "Map -- " & tabRooms.Fields("Name") & " (" & nStartMap & "/" & nStartRoom & ")  "
End If

'If chkMapOptions(6).Value = 0 Then Call SetTopMostWindow(Me.hwnd, True)

If Not nMapStartRoom = nStartRoom Then
    nMapLastRoom = nMapStartRoom
    nMapLastMap = nMapStartMap
End If

bMapStillMapping = True
Call LockWindowUpdate(Me.hWnd)

'picMap.Visible = False
picMap.Cls
Me.MousePointer = vbHourglass
DoEvents
'20x20
'sMapSECorner = 400
'nMapRowLength = 20
If Not nCenterCell = 0 Then nMapCenterCell = nCenterCell
'If nMapCenterCell > sMapSECorner Then nMapCenterCell = 210

For x = 1 To 900
    TTlbl.DelToolTip picMap.hWnd, 0
    lblRoomCell(x).BackColor = &HFFFFFF
    lblRoomCell(x).Visible = False
    lblRoomCell(x).Tag = 0
    UnchartedCells(x) = 0
    CellRoom(x, 1) = 0
    CellRoom(x, 2) = 0
Next x

Call ResizeMap

StopBuild = False

nMapStartRoom = nStartRoom
nMapStartMap = nStartMap

CellRoom(nMapCenterCell, 1) = nMapStartMap
CellRoom(nMapCenterCell, 2) = nMapStartRoom

fraPresets.Visible = False
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

'For x = nMapCenterCell To 1 Step -1 '1 To sMapSECorner
'    If StopBuild = True Then GoTo Cancel:
'    If UnchartedCells(x) = 1 Then
'        For y = 1 To sMapSECorner
'            If Not CellRoom(x, 1) = 0 Then
'                If Not x = y Then
'                    If CellRoom(y, 2) = CellRoom(x, 2) Then
'                        If CellRoom(y, 1) = CellRoom(x, 1) Then
'                            CellRoom(x, 2) = 0
'                            CellRoom(x, 1) = 0
'                            UnchartedCells(x) = 0
'                            GoTo skiproom:
'                        End If
'                    End If
'                End If
'            End If
'        Next y
'        Call MapMapExits(x, CellRoom(x, 2), CellRoom(x, 1))
'        bCheckAgain = True
'    End If
'skiproom:
'    'DoEvents
'Next x
'For x = nMapCenterCell To sMapSECorner
'    If StopBuild = True Then GoTo Cancel:
'    If UnchartedCells(x) = 1 Then
'        For y = 1 To sMapSECorner
'            If Not CellRoom(x, 1) = 0 Then
'                If Not x = y Then
'                    If CellRoom(y, 2) = CellRoom(x, 2) Then
'                        If CellRoom(y, 1) = CellRoom(x, 1) Then
'                            CellRoom(x, 2) = 0
'                            CellRoom(x, 1) = 0
'                            UnchartedCells(x) = 0
'                            GoTo skiproom2:
'                        End If
'                    End If
'                End If
'            End If
'        Next y
'        Call MapMapExits(x, CellRoom(x, 2), CellRoom(x, 1))
'        bCheckAgain = True
'    End If
'skiproom2:
'    'DoEvents
'Next x

If bCheckAgain Then GoTo again:

Call MapDrawOnRoom(lblRoomCell(nMapCenterCell), drSquare, 4, BrightBlue)

DoEvents
cmdMapShowUnused.Caption = "S&how Unused Blocks"
For x = 1 To 900
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

Private Sub MapMapExits(Cell As Integer, Room As Long, Map As Long)
Dim ActivatedCell As Integer, x As Integer
Dim rc As RECT, ToolTipString As String, sText As String, y As Long
Dim sRemote As String, sMonsters As String, sArray() As String, sPlaced As String
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
    TTlbl.SetToolTipItem picMap.hWnd, 0, rc.Left, rc.Top, rc.Right, rc.Bottom, ToolTipString, False
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

If chkMapOptions(2).Value = 0 And Not tabRooms.Fields("Lair") = Chr(0) Then
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
    TTlbl.SetToolTipItem picMap.hWnd, 0, rc.Left, rc.Top, rc.Right, rc.Bottom, ToolTipString, False
End If

UnchartedCells(Cell) = 2

Exit Sub

Error:
Call HandleError("MapMapExits")
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
        MapActivateCell = (FromCell - 30)
        'checking to see if it's on the north edge
        If MapActivateCell < 1 Then
            Call MapDrawOnRoom(lblRoomCell(FromCell), drLineN, 4, Grey)
            GoTo DontActivate
        End If

    Case 1: 'south
        MapActivateCell = (FromCell + 30)
        'checking to see if it's on the south edge
        If MapActivateCell > sMapSECorner Then
            Call MapDrawOnRoom(lblRoomCell(FromCell), drLineS, 4, Grey)
            GoTo DontActivate
        End If

    Case 2: 'east
        MapActivateCell = (FromCell + 1)
        'checking to see if it's on the east edge
        For temp = nMapRowLength To sMapSECorner Step 30
            If FromCell = temp Then
                Call MapDrawOnRoom(lblRoomCell(FromCell), drLineE, 4, Grey)
                GoTo DontActivate
            End If
        Next
        
    Case 3: 'west
        MapActivateCell = (FromCell - 1)
        'checking to see if it's on the west edge
        For temp = 1 To sMapSECorner Step 30
            If FromCell = temp Then
                Call MapDrawOnRoom(lblRoomCell(FromCell), drLineW, 4, Grey)
                GoTo DontActivate
            End If
        Next

    Case 4: 'northeast
        MapActivateCell = (FromCell - 29)
        'checking to see if it's on the north edge
        If MapActivateCell < 1 Then
            Call MapDrawOnRoom(lblRoomCell(FromCell), drLineNE, 4, Grey)
            GoTo DontActivate
        End If
        'checking to see if it's on the east edge
        For temp = nMapRowLength To sMapSECorner Step 30
            If FromCell = temp Then
                Call MapDrawOnRoom(lblRoomCell(FromCell), drLineNE, 4, Grey)
                GoTo DontActivate
            End If
        Next

    Case 5: 'northwest
        MapActivateCell = (FromCell - 31)
        'checking to see if it's on the north edge
        If MapActivateCell < 1 Then
            Call MapDrawOnRoom(lblRoomCell(FromCell), drLineNW, 4, Grey)
            GoTo DontActivate:
        End If
        'checking to see if it's on the west edge
        For temp = 1 To sMapSECorner Step 30
            If FromCell = temp Then
                Call MapDrawOnRoom(lblRoomCell(FromCell), drLineNW, 4, Grey)
                GoTo DontActivate
            End If
        Next

    Case 6: 'southeast
        MapActivateCell = (FromCell + 31)
        'checking to see if it's on the south edge
        If MapActivateCell > sMapSECorner Then
            Call MapDrawOnRoom(lblRoomCell(FromCell), drLineSE, 4, Grey)
            GoTo DontActivate
        End If
        'checking to see if it's on the east edge
        For temp = nMapRowLength To sMapSECorner Step 30
            If FromCell = temp Then
                Call MapDrawOnRoom(lblRoomCell(FromCell), drLineSE, 4, Grey)
                GoTo DontActivate
            End If
        Next

    Case 7: 'southwest
        MapActivateCell = (FromCell + 29)
        'checking to see if it's on the south edge
        If MapActivateCell > sMapSECorner Then
            Call MapDrawOnRoom(lblRoomCell(FromCell), drLineSW, 4, Grey)
            GoTo DontActivate:
        End If
        'checking to see if it's on the west edge
        For temp = 1 To sMapSECorner Step 30
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
Dim x As Long, sLook As String, nExitType As Integer, RoomExit As RoomExitType, oLI As ListItem, RoomExit2 As RoomExitType
Dim nRecNum As Long, y As Long, sNumbers As String, sCommand As String, nMap As Long, nRoom As Long, sChar As String
Dim sArray() As String

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

If tabRooms.Fields("CMD") > 0 Then 'chkMapOptions(4).Value = 0 And
    tabTBInfo.Index = "pkTBInfo"
    tabTBInfo.Seek "=", tabRooms.Fields("CMD")
    If tabTBInfo.NoMatch = False Then
        sCommand = tabTBInfo.Fields("Action")
        x = InStr(1, sCommand, "teleport ")
        If x > 0 Then
            Do While x < Len(sCommand)
                x = x + Len("teleport ") 'position x just after the search text
                y = x
                Do While y < Len(sCommand) + 2
                    sChar = Mid(sCommand, y, 1)
                    Select Case sChar
                        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
                        Case " ":
                            If y > x And nRoom = 0 Then
                                nRoom = Val(Mid(sCommand, x, y - x))
                                x = y + 1
                            Else
                                nMap = Val(Mid(sCommand, x, y - x))
                                Exit Do
                            End If
                        Case Else:
                            If y > x And nRoom = 0 Then
                                nRoom = Val(Mid(sCommand, x, y - x))
                                Exit Do
                            Else
                                nMap = Val(Mid(sCommand, x, y - x))
                                Exit Do
                            End If
                            Exit Do
                    End Select
                    y = y + 1
                Loop
                
                If Not nRoom = 0 Then
                    If nMap = 0 Then nMap = nMapNumber
                    For Each oLI In lvMapLoc.ListItems
                        If oLI.Tag = nMap & "/" & nRoom Then GoTo skiptele:
                    Next
                    
                    Set oLI = lvMapLoc.ListItems.Add()
                    oLI.Text = "Teleport: " & GetTextblockCMDText("teleport " & nRoom & " " & nMap, sCommand) _
                        & " --> " & GetRoomName(, nMap, nRoom, False)
                    oLI.Tag = nMap & "/" & nRoom
                End If
skiptele:
                nRoom = 0
                nMap = 0
                x = InStr(y, sCommand, "teleport ")
                If x = 0 Then x = Len(sCommand)
            Loop
            tabRooms.Seek "=", nMapNumber, nRoomNumber
        End If
        
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

Private Sub cmdOptions_Click()
If fraOptions.Visible = True Then
    fraOptions.Visible = False
    If nMapStartMap > 0 And nMapStartRoom > 0 Then
        Call MapStartMapping(nMapStartMap, nMapStartRoom)
    End If
    Exit Sub
End If

fraPresets.Visible = False
fraOptions.Visible = True
End Sub

Private Sub cmdPresets_Click()
If fraPresets.Visible = True Then
    fraPresets.Visible = False
    Exit Sub
End If

fraOptions.Visible = False
fraPresets.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Call WriteINI("Settings", "ExMapFollowMap", chkMapOptions(0).Value)
Call WriteINI("Settings", "ExMapNoHidden", chkMapOptions(1).Value)
Call WriteINI("Settings", "ExMapNoLairs", chkMapOptions(2).Value)
Call WriteINI("Settings", "ExMapNoNPC", chkMapOptions(3).Value)
Call WriteINI("Settings", "ExMapNoCMD", chkMapOptions(4).Value)
Call WriteINI("Settings", "ExMapNoTooltips", chkMapOptions(5).Value)
Call WriteINI("Settings", "MapExternalOnTop", chkMapOptions(6).Value)
Call WriteINI("Settings", "ExMapSize", cmbMapSize.ListIndex)
Call WriteINI("Settings", "ExMapMainOverlap", chkMapOptions(8).Value)

Set TTlbl = Nothing

If Not Me.WindowState = vbMinimized And Not Me.WindowState = vbMaximized Then
    Call WriteINI("Settings", "ExMapTop", Me.Top)
    Call WriteINI("Settings", "ExMapLeft", Me.Left)
End If

If Not bAppTerminating Then
    If frmMain.WindowState = vbMinimized Then frmMain.WindowState = frmMain.nWindowState
    frmMain.SetFocus
End If

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
    fraOptions.Visible = False
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

Private Sub lvMapLoc_DblClick()
Dim lR As Long

On Error GoTo Error:

If lvMapLoc.ListItems.Count = 0 Then Exit Sub
Call frmMain.GotoLocation(lvMapLoc.SelectedItem, nMapStartMap, Me)
'If frmMain.WindowState = vbMinimized Then frmMain.WindowState = vbNormal
'frmMain.SetFocus

'If chkMapOptions(6).Value = 0 Then
'    If FormIsLoaded("frmResults") Then
'        If frmResults.objFormOwner Is Me Then
'            lR = SetTopMostWindow(frmResults.hwnd, True)
'        End If
'    End If
'End If
DoEvents
out:
Exit Sub
Error:
Call HandleError("lvMapLoc_DblClick")
Resume out:
End Sub

Private Sub txtRoomMap_GotFocus()
Call SelectAll(txtRoomMap)
End Sub

Private Sub txtRoomRoom_GotFocus()
Call SelectAll(txtRoomRoom)
End Sub

Public Sub LoadPresets(Optional ByVal bReset As Boolean)
Dim x As Integer, sSectionName As String, nMap As Long, nRoom As Long, sName As String
Dim cReg As clsRegistryRoutines, nError As Integer, bResult As Boolean

On Error GoTo Error:

Set cReg = New clsRegistryRoutines

Me.MousePointer = vbHourglass

If InStr(1, frmMain.lblDatVer.Caption, "-") = 0 Then
    sSectionName = "Custom_Presets"
Else
    sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ") & "_Presets"
End If

nError = RegCreateKeyPath(HKEY_LOCAL_MACHINE, "Software\MMUD Explorer\Presets\" & sSectionName)
If nError > 0 Then GoTo Error:

cReg.hkey = HKEY_LOCAL_MACHINE
cReg.KeyRoot = "Software\MMUD Explorer\Presets"
cReg.Subkey = sSectionName

If bReset Then
    For x = 0 To 49
        bResult = cReg.SetRegistryValue("Map" & x, "0", REG_SZ)
        If bResult = False Then Err.Raise 0, "LoadPresets", "Error Setting Registry Values"
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
    End If
    
Next x

For x = 0 To 9
    cmdMapPreset(x).Caption = cReg.GetRegistryValue("Name" & x, "unset")
    cmdMapPreset(x).Tag = x
Next x

Me.MousePointer = vbDefault

Exit Sub

Error:
Call HandleError("LoadPresets")

End Sub

Private Sub cmdResetPresets_Click()
Dim nYesNo As Integer

nYesNo = MsgBox("Are you sure you want to reset the presets to the default set?", vbYesNo + vbDefaultButton2 + vbQuestion, "Reset Presets?")

If nYesNo = vbYes Then Call LoadPresets(True)

Call frmMain.LoadPresets

End Sub

Private Sub cmdMapPreset_Click(Index As Integer)
Dim nMap As Long, nRoom As Long, sSectionName As String
Dim cReg As clsRegistryRoutines
On Error GoTo Error:

Set cReg = New clsRegistryRoutines

If InStr(1, frmMain.lblDatVer.Caption, "-") = 0 Then
    sSectionName = "Custom_Presets"
Else
    sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ") & "_Presets"
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

Private Sub cmdEditPreset_Click(Index As Integer)
Dim sSectionName As String, lR As Long
Dim cReg As clsRegistryRoutines
On Error GoTo Error:
Set cReg = New clsRegistryRoutines

If InStr(1, frmMain.lblDatVer.Caption, "-") = 0 Then
    sSectionName = "Custom_Presets"
Else
    sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ") & "_Presets"
End If
'sSectionName = RemoveCharacter(lblDatVer.Caption, " ") & "_Presets"

cReg.hkey = HKEY_LOCAL_MACHINE
cReg.KeyRoot = "Software\MMUD Explorer\Presets"
cReg.Subkey = sSectionName

Unload frmEditPreset
Load frmEditPreset
frmEditPreset.nPreset = Val(cmdMapPreset(Index).Tag)
frmEditPreset.lblCaption.Caption = "Editing Preset #" & (cmdMapPreset(Index).Tag + 1)
frmEditPreset.txtMap.Text = cReg.GetRegistryValue("Map" & cmdMapPreset(Index).Tag, 0) 'ReadINI(sSectionName, "Map" & cmdMapPreset(index).Tag)
frmEditPreset.txtRoom.Text = cReg.GetRegistryValue("Room" & cmdMapPreset(Index).Tag, 0) 'ReadINI(sSectionName, "Room" & cmdMapPreset(index).Tag)
frmEditPreset.txtCaption.Text = cReg.GetRegistryValue("Name" & cmdMapPreset(Index).Tag, "unset") 'ReadINI(sSectionName, "Name" & cmdMapPreset(index).Tag)
Set frmEditPreset.objFormOwner = Me
If chkMapOptions(6).Value = 0 Then lR = SetTopMostWindow(Me.hWnd, False)
DoEvents
frmEditPreset.Show vbModal, Me
If chkMapOptions(6).Value = 0 Then lR = SetTopMostWindow(Me.hWnd, True)

If Not frmEditPreset.nPreset < 0 Then
    Call cReg.SetRegistryValue("Map" & cmdMapPreset(Index).Tag, frmEditPreset.txtMap.Text, REG_SZ)
    Call cReg.SetRegistryValue("Room" & cmdMapPreset(Index).Tag, frmEditPreset.txtRoom.Text, REG_SZ)
    Call cReg.SetRegistryValue("Name" & cmdMapPreset(Index).Tag, frmEditPreset.txtCaption.Text, REG_SZ)
    cmdMapPreset(Index).Caption = frmEditPreset.txtCaption.Text
End If

Unload frmEditPreset

Call frmMain.LoadPresets

Exit Sub
Error:
Call HandleError("cmdEditPreset_Click")
Unload frmEditPreset
End Sub
