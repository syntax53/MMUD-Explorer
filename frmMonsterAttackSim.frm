VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmMonsterAttackSim 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monster Attack Simulator"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14850
   Icon            =   "frmMonsterAttackSim.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   14850
   Begin VB.Timer timWindowMove 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.CheckBox chkHideEnergy 
      Caption         =   "Hide Energy Info."
      Height          =   195
      Left            =   7920
      TabIndex        =   63
      Top             =   5040
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.Frame fraStats 
      Caption         =   "Results"
      Height          =   4455
      Left            =   9000
      TabIndex        =   3
      Top             =   0
      Width           =   5775
      Begin VB.TextBox txtStatResistDodge 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   5025
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txtStatDmgResist 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   4365
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txtStatSuccess 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   3705
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txtStatAvgRound 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   2985
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   2280
         Width           =   675
      End
      Begin VB.TextBox txtStatTrueCast 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   2205
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   2280
         Width           =   675
      End
      Begin VB.TextBox txtStatResistDodge 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   5025
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   1860
         Width           =   615
      End
      Begin VB.TextBox txtStatDmgResist 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   4365
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1860
         Width           =   615
      End
      Begin VB.TextBox txtStatSuccess 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   3705
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1860
         Width           =   615
      End
      Begin VB.TextBox txtStatAvgRound 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   2985
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1860
         Width           =   675
      End
      Begin VB.TextBox txtStatTrueCast 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   2205
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1860
         Width           =   675
      End
      Begin VB.TextBox txtStatResistDodge 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   5025
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtStatDmgResist 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   4365
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtStatSuccess 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   3705
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtStatAvgRound 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   2985
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1440
         Width           =   675
      End
      Begin VB.TextBox txtStatTrueCast 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   2205
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1440
         Width           =   675
      End
      Begin VB.TextBox txtStatResistDodge 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   5025
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1020
         Width           =   615
      End
      Begin VB.TextBox txtStatDmgResist 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   4365
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1020
         Width           =   615
      End
      Begin VB.TextBox txtStatSuccess 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   3705
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1020
         Width           =   615
      End
      Begin VB.TextBox txtStatAvgRound 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   2985
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1020
         Width           =   675
      End
      Begin VB.TextBox txtStatTrueCast 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   2205
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1020
         Width           =   675
      End
      Begin VB.TextBox txtStatResistDodge 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   5025
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "100%"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtStatDmgResist 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   4365
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "100%"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtStatSuccess 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   3705
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "100%"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtStatAvgRound 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   2985
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "99999"
         Top             =   600
         Width           =   675
      End
      Begin VB.TextBox txtStatTrueCast 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   2205
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "45.5%"
         Top             =   600
         Width           =   675
      End
      Begin VB.Label lblResultsAttBreakdown 
         Alignment       =   2  'Center
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
         Left            =   120
         TabIndex        =   69
         Top             =   3420
         Width           =   5475
      End
      Begin VB.Line Line1 
         X1              =   5580
         X2              =   120
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label lblResultsMaxRound 
         Alignment       =   2  'Center
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
         TabIndex        =   41
         Top             =   3840
         Width           =   5475
      End
      Begin VB.Label lblResultsAvgDmg 
         Alignment       =   2  'Center
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
         TabIndex        =   40
         Top             =   2940
         Width           =   5475
      End
      Begin VB.Label lblHeadings 
         Alignment       =   2  'Center
         Caption         =   "Attack"
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
         TabIndex        =   4
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label lblAttack 
         Caption         =   "1"
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblAttack 
         Caption         =   "2"
         Height          =   360
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   1020
         Width           =   2055
      End
      Begin VB.Label lblAttack 
         Caption         =   "3"
         Height          =   360
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label lblAttack 
         Caption         =   "4"
         Height          =   360
         Index           =   3
         Left            =   120
         TabIndex        =   28
         Top             =   1860
         Width           =   2055
      End
      Begin VB.Label lblAttack 
         Caption         =   "5"
         Height          =   360
         Index           =   4
         Left            =   120
         TabIndex        =   34
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label lblHeadings 
         Alignment       =   2  'Center
         Caption         =   "%resist /dodge"
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
         Index           =   18
         Left            =   4965
         TabIndex        =   9
         Top             =   180
         Width           =   735
      End
      Begin VB.Label lblHeadings 
         Alignment       =   2  'Center
         Caption         =   "%dmg Resist"
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
         Index           =   17
         Left            =   4365
         TabIndex        =   8
         Top             =   180
         Width           =   615
      End
      Begin VB.Label lblHeadings 
         Alignment       =   2  'Center
         Caption         =   "% Hit"
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
         Index           =   16
         Left            =   3705
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblHeadings 
         Alignment       =   2  'Center
         Caption         =   "Avg Hit"
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
         Index           =   15
         Left            =   2925
         TabIndex        =   6
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblHeadings 
         Alignment       =   2  'Center
         Caption         =   "True Attk%"
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
         Index           =   13
         Left            =   2265
         TabIndex        =   5
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.TextBox txtNumRounds 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "M/dd/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      Left            =   8220
      MaxLength       =   6
      TabIndex        =   66
      Text            =   "2000"
      Top             =   5340
      Width           =   915
   End
   Begin VB.CheckBox chkDynamicRounds 
      Alignment       =   1  'Right Justify
      Caption         =   "or Dynamic:"
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
      Left            =   9360
      TabIndex        =   67
      ToolTipText     =   "This will run the sim in 1,000 round increments untl the change in result is < 0.001%"
      Top             =   5400
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkCombatMaxRoundOnly 
      Caption         =   "Show combat log only for max round seen."
      Height          =   195
      Left            =   7080
      TabIndex        =   62
      Top             =   4740
      Width           =   3435
   End
   Begin VB.CommandButton cmdRunSim 
      Caption         =   "Run Simulator"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   11520
      TabIndex        =   64
      Top             =   5100
      Width           =   3075
   End
   Begin VB.Frame fraChar 
      Caption         =   "Character Defenses"
      Height          =   1635
      Left            =   120
      TabIndex        =   42
      Top             =   4620
      Width           =   5895
      Begin VB.CommandButton cmdAlwaysDodgeQ 
         Caption         =   "?"
         Height          =   315
         Left            =   3960
         TabIndex        =   53
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox txtElementalResist 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   3540
         MaxLength       =   4
         TabIndex        =   58
         Text            =   "0"
         Top             =   1140
         Width           =   735
      End
      Begin VB.TextBox txtElementalResist 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   1020
         MaxLength       =   4
         TabIndex        =   55
         Text            =   "0"
         Top             =   1140
         Width           =   735
      End
      Begin VB.TextBox txtElementalResist 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   2700
         MaxLength       =   4
         TabIndex        =   57
         Text            =   "0"
         Top             =   1140
         Width           =   735
      End
      Begin VB.TextBox txtElementalResist 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   1860
         MaxLength       =   4
         TabIndex        =   56
         Text            =   "0"
         Top             =   1140
         Width           =   735
      End
      Begin VB.TextBox txtElementalResist 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   180
         MaxLength       =   4
         TabIndex        =   54
         Text            =   "0"
         Top             =   1140
         Width           =   735
      End
      Begin VB.CheckBox chkAlwaysDodge 
         Caption         =   "MegaMUD Dodge"
         Height          =   375
         Left            =   2760
         TabIndex        =   52
         Top             =   420
         Width           =   1215
      End
      Begin VB.CommandButton cmdResetUserDefs 
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
         Left            =   3900
         TabIndex        =   43
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdResetUserDefs 
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
         Left            =   4860
         TabIndex        =   44
         Top             =   0
         Width           =   855
      End
      Begin VB.TextBox txtUserAC 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   180
         MaxLength       =   4
         TabIndex        =   49
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtUserDodge 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1860
         MaxLength       =   4
         TabIndex        =   51
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtUserMR 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   4500
         MaxLength       =   4
         TabIndex        =   59
         Top             =   1140
         Width           =   795
      End
      Begin VB.CheckBox chkUserAntiMagic 
         Caption         =   "Anti-Magic"
         Height          =   255
         Left            =   4560
         TabIndex        =   61
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtUserDR 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   1020
         MaxLength       =   4
         TabIndex        =   50
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdMRNote 
         Caption         =   "!"
         Height          =   315
         Left            =   5340
         TabIndex        =   60
         Top             =   1140
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "R-Stone"
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
         Left            =   3525
         TabIndex        =   74
         Top             =   900
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "R-Cold"
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
         Left            =   1095
         TabIndex        =   73
         Top             =   900
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "R-Water"
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
         Left            =   2670
         TabIndex        =   72
         Top             =   900
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "R-Fire"
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
         Left            =   285
         TabIndex        =   71
         Top             =   900
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "R-Litng"
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
         Left            =   1890
         TabIndex        =   70
         Top             =   900
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Dodge%"
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
         Left            =   1860
         TabIndex        =   47
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "AC"
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
         Left            =   240
         TabIndex        =   45
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         Index           =   3
         Left            =   4530
         TabIndex        =   48
         Top             =   900
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "DR"
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
         Index           =   4
         Left            =   1110
         TabIndex        =   46
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.TextBox txtCombatLog 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   480
      Width           =   8775
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
      Left            =   3720
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   60
      Width           =   5175
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   315
      Left            =   6240
      TabIndex        =   68
      Top             =   5880
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "# Rounds to Sim:"
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
      Left            =   6540
      TabIndex        =   65
      Top             =   5400
      Width           =   1605
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLabelArray 
      Alignment       =   1  'Right Justify
      Caption         =   "Choose Monster:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3315
   End
End
Attribute VB_Name = "frmMonsterAttackSim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Dim tWindowSize As WindowSizeProperties

Public nLastPosTop As Long
Public nLastPosLeft As Long
Public nLastPosMoved As Long
Public nLastPosMonitor As Long

Public nLastTimerTop As Long
Public nLastTimerLeft As Long

Private Sub chkUserAntiMagic_Click()
If chkUserAntiMagic.Value = 0 Then
    chkUserAntiMagic.FontBold = False
Else
    chkUserAntiMagic.FontBold = True
End If
End Sub

Private Sub cmbMonsterList_Click()
Call ResetFields
End Sub

Private Sub cmbMonsterList_KeyPress(KeyAscii As Integer)
KeyAscii = AutoComplete(cmbMonsterList, KeyAscii, False)
End Sub

Private Sub cmdResetUserDefs_Click(Index As Integer)
On Error GoTo error:
Dim nParty As Integer

chkAlwaysDodge.Value = 0

If Index = 0 Then
    fraChar.Caption = "Character Defenses"
    txtUserAC.Text = 0
    txtUserDR.Text = 0
    txtUserDodge.Text = 0
    txtUserMR.Text = 50
    chkUserAntiMagic.Value = 0
    txtElementalResist(0).Text = 0
    txtElementalResist(1).Text = 0
    txtElementalResist(2).Text = 0
    txtElementalResist(3).Text = 0
    txtElementalResist(5).Text = 0
Else
    If frmMain.optMonsterFilter(1).Value = True And val(frmMain.txtMonsterLairFilter(0).Text) > 1 Then
        nParty = val(frmMain.txtMonsterLairFilter(0).Text)
    End If
    If nParty < 1 Then nParty = 1
    If nParty > 6 Then nParty = 6
    
    If nParty = 1 Then
        fraChar.Caption = "Character Defenses"
        txtUserAC.Text = Round(val(frmMain.lblInvenCharStat(2).Caption))
        txtUserDR.Text = Round(val(frmMain.lblInvenCharStat(3).Caption))
        txtUserMR.Text = Round(val(frmMain.txtCharMR.Text))
        txtUserDodge.Text = Round(val(frmMain.lblCharDodge.Tag))
        chkUserAntiMagic.Value = frmMain.chkCharAntiMagic.Value
        txtElementalResist(0).Text = frmMain.lblInvenCharStat(28).Tag 'col
        txtElementalResist(1).Text = frmMain.lblInvenCharStat(27).Tag 'fir
        txtElementalResist(2).Text = frmMain.lblInvenCharStat(25).Tag 'sto
        txtElementalResist(3).Text = frmMain.lblInvenCharStat(29).Tag 'lit
        txtElementalResist(5).Text = frmMain.lblInvenCharStat(26).Tag 'wat
    Else
        fraChar.Caption = "PARTY Defenses"
        'txtMonsterLairFilter... 0-#, 1-ac, 2-dr, 3-mr, 4-dodge, 5-HP, 6-#antimag, 7-hpregen, 8-accy
        txtUserAC.Text = Round(val(frmMain.txtMonsterLairFilter(1).Text))
        txtUserDR.Text = Round(val(frmMain.txtMonsterLairFilter(2).Text))
        txtUserMR.Text = Round(val(frmMain.txtMonsterLairFilter(3).Text))
        txtUserDodge.Text = Round(val(frmMain.txtMonsterLairFilter(4).Text))
        If val(frmMain.txtMonsterLairFilter(6).Text) > 1 Then chkUserAntiMagic.Value = 1
        txtElementalResist(0).Text = 0
        txtElementalResist(1).Text = 0
        txtElementalResist(2).Text = 0
        txtElementalResist(3).Text = 0
        txtElementalResist(5).Text = 0
    End If
End If

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("cmdResetUserDefs_Click")
Resume out:
End Sub

Private Sub cmdRunSim_Click()
On Error GoTo error:
Dim clsMonAtkSimThisForm As New clsMonsterAttackSim, x As Integer

Me.Enabled = False

Call ResetFields

If cmbMonsterList.ItemData(cmbMonsterList.ListIndex) <= 0 Then Exit Sub

If val(txtNumRounds.Text) > 500000 Then txtNumRounds.Text = 500000

Call clsMonAtkSimThisForm.ResetValues
Set clsMonAtkSimThisForm.cProgressBar = ProgressBar
clsMonAtkSimThisForm.bUseCPU = False
clsMonAtkSimThisForm.nCombatLogMaxRounds = 100
If chkCombatMaxRoundOnly.Value = 1 Then clsMonAtkSimThisForm.bCombatLogMaxRoundOnly = True
clsMonAtkSimThisForm.nNumberOfRounds = val(txtNumRounds.Text)
clsMonAtkSimThisForm.nUserMR = 50
clsMonAtkSimThisForm.bGreaterMUD = bGreaterMUD
clsMonAtkSimThisForm.bDynamicCalc = IIf(chkDynamicRounds.Value = 1, True, False)
clsMonAtkSimThisForm.nDynamicCalcDifference = 0.0001
If chkHideEnergy.Value = 1 Then clsMonAtkSimThisForm.bHideEnergyInfo = True
If chkAlwaysDodge.Value = 1 Then clsMonAtkSimThisForm.bDodgeBeforeAC = True

If val(txtUserAC.Text) > 0 Then clsMonAtkSimThisForm.nUserAC = val(txtUserAC.Text)
If val(txtUserDR.Text) > 0 Then clsMonAtkSimThisForm.nUserDR = val(txtUserDR.Text)
If val(txtUserDodge.Text) > 0 Then clsMonAtkSimThisForm.nUserDodge = val(txtUserDodge.Text)
If val(txtUserMR.Text) > 0 Then clsMonAtkSimThisForm.nUserMR = val(txtUserMR.Text)

If val(txtElementalResist(0).Text) > 0 Then clsMonAtkSimThisForm.nUserRCOL = val(txtElementalResist(0).Text)
If val(txtElementalResist(1).Text) > 0 Then clsMonAtkSimThisForm.nUserRFIR = val(txtElementalResist(1).Text)
If val(txtElementalResist(2).Text) > 0 Then clsMonAtkSimThisForm.nUserRSTO = val(txtElementalResist(2).Text)
If val(txtElementalResist(3).Text) > 0 Then clsMonAtkSimThisForm.nUserRLIT = val(txtElementalResist(3).Text)
If val(txtElementalResist(5).Text) > 0 Then clsMonAtkSimThisForm.nUserRWAT = val(txtElementalResist(5).Text)

If chkUserAntiMagic.Value = 1 Then clsMonAtkSimThisForm.nUserAntiMagic = 1

Call PopulateMonsterDataToAttackSim(cmbMonsterList.ItemData(cmbMonsterList.ListIndex), clsMonAtkSimThisForm)

For x = 0 To 4
    If Len(clsMonAtkSimThisForm.sAtkName(x)) > 0 Then
        lblAttack(x).Caption = clsMonAtkSimThisForm.sAtkName(x)
    Else
        lblAttack(x).Caption = ""
    End If
Next x

If clsMonAtkSimThisForm.nNumberOfRounds > 0 Then clsMonAtkSimThisForm.RunSim

txtCombatLog.Text = Trim(clsMonAtkSimThisForm.sCombatLog)
If clsMonAtkSimThisForm.nTotalAttacks > 0 And clsMonAtkSimThisForm.nNumberOfRounds > 0 Then
    lblResultsAvgDmg.Caption = "AVG Dmg/Rnd: " & Round(clsMonAtkSimThisForm.nTotalDamage / clsMonAtkSimThisForm.nNumberOfRounds, 1)
    lblResultsMaxRound.Caption = "Max/Seen: " & clsMonAtkSimThisForm.GetMaxDamage & "/" & clsMonAtkSimThisForm.nMaxRoundDamage
    lblResultsAttBreakdown.Caption = "(Physical/Spell: " & Round(clsMonAtkSimThisForm.nAverageDamagePhys) & "/" & Round(clsMonAtkSimThisForm.nAverageDamageSpell) & ")"
    
    For x = 0 To 4
        If clsMonAtkSimThisForm.nAtkType(x) > 0 Then
            txtStatTrueCast(x).Text = Round(clsMonAtkSimThisForm.nStatAtkAttempted(x) / clsMonAtkSimThisForm.nTotalAttacks, 3) * 100
            'txtStatAttRound(x).Text = Round(clsMonAtkSimThisForm.nStatAtkAttempted(x) / clsMonAtkSimThisForm.nNumberOfRounds, 2)
            
            If clsMonAtkSimThisForm.nStatAtkTotalDamage(x) > 0 And clsMonAtkSimThisForm.nStatAtkHits(x) Then
                txtStatAvgRound(x).Text = Round(clsMonAtkSimThisForm.nStatAtkTotalDamage(x) / clsMonAtkSimThisForm.nStatAtkHits(x))
            Else
                txtStatAvgRound(x).Text = 0
            End If
            
            If clsMonAtkSimThisForm.nStatAtkAttempted(x) > 0 Then
                txtStatSuccess(x).Text = Round(clsMonAtkSimThisForm.nStatAtkHits(x) / clsMonAtkSimThisForm.nStatAtkAttempted(x), 3) * 100
            Else
                txtStatSuccess(x).Text = 0
            End If
            
            If clsMonAtkSimThisForm.nStatAtkDmgResisted(x) <> 0 Then
                txtStatDmgResist(x).Text = IIf(clsMonAtkSimThisForm.nStatAtkTotalDamage(x) = 0, 100, _
                    Round(clsMonAtkSimThisForm.nStatAtkDmgResisted(x) / (clsMonAtkSimThisForm.nStatAtkDmgResisted(x) + clsMonAtkSimThisForm.nStatAtkTotalDamage(x)), 3) * 100)
            Else
                txtStatDmgResist(x).Text = 0
            End If
            
            If clsMonAtkSimThisForm.nStatAtkAttempted(x) > 0 And clsMonAtkSimThisForm.nAtkType(x) = 2 Then 'spell
                txtStatResistDodge(x).Text = Round(clsMonAtkSimThisForm.nStatAtkAttemptDodgedOrResisted(x) / clsMonAtkSimThisForm.nStatAtkAttempted(x), 3) * 100
            
            'update 2024.01.13
            'ElseIf clsMonAtkSimThisForm.nStatAtkHits(x) > 0 Or clsMonAtkSimThisForm.nStatAtkAttemptDodgedOrResisted(x) > 0 Then
            ElseIf clsMonAtkSimThisForm.nStatAtkAttemptDodgedOrResisted(x) > 0 And clsMonAtkSimThisForm.nStatAtkAttempted(x) > 0 Then
                'update 2024.01.13
                'txtStatResistDodge(x).Text = Round(clsMonAtkSimThisForm.nStatAtkAttemptDodgedOrResisted(x) / (clsMonAtkSimThisForm.nStatAtkHits(x) + clsMonAtkSimThisForm.nStatAtkAttemptDodgedOrResisted(x)), 3) * 100
                txtStatResistDodge(x).Text = Round(clsMonAtkSimThisForm.nStatAtkAttemptDodgedOrResisted(x) / clsMonAtkSimThisForm.nStatAtkAttempted(x), 3) * 100
            Else
                txtStatResistDodge(x).Text = 0
            End If
        End If
    Next x
End If

out:
On Error Resume Next
ProgressBar.Value = 0
Me.Enabled = True
Exit Sub
error:
Call HandleError("cmdRunSim_Click")
Resume out:
End Sub


Private Sub cmdAlwaysDodgeQ_Click()
MsgBox "MME can now calculate your dodge value and should have populated it for you.  The value you see in MegaMUD is likely less than your actual dodge value.  " _
    & "This is because dodge is checked after the AC check and determined to be a hit.  This is compounded by different mobs with different accuracy causing the reported dodge value to fluctuate." _
    & vbCrLf & vbCrLf _
    & "This option will cause dodge to be checked before AC and match what MegaMUD sees.  You can use this if the MegaMUD dodge value is all you know or what you want to go by.", vbInformation
End Sub

Private Sub Form_Load()
On Error GoTo error:

Call SetWindowLong(Me.hWnd, GWL_HWNDPARENT, 0)

'stop windows from resizing fixed-size windows when changing dpi
If bDPIAwareMode Then Call SubclassFormMinMaxSize(Me, tWindowSize, True)

Call ResetFields
Call LoadMonsters

If frmMain.WindowState = vbMinimized Then
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
Else
    Me.Left = frmMain.Left + ((frmMain.Width - Me.Width) / 2)
    Me.Top = frmMain.Top + ((frmMain.Height - Me.Height) / 2)
End If

timWindowMove.Enabled = True

Call cmdResetUserDefs_Click(1)

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("Form_Load")
Resume out:
End Sub

Public Sub ResetFields()
Dim x As Integer
On Error GoTo error:

lblResultsAvgDmg.Caption = ""
lblResultsMaxRound.Caption = ""
lblResultsAttBreakdown.Caption = ""

For x = 0 To 4
    lblAttack(x).Caption = (x + 1) & "."
    txtStatTrueCast(x).Text = ""
    'txtStatAttRound(x).Text = ""
    txtStatAvgRound(x).Text = ""
    txtStatSuccess(x).Text = ""
    txtStatDmgResist(x).Text = ""
    txtStatResistDodge(x).Text = ""
Next x

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("ResetMonsterFields")
Resume out:
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
Call ExpandCombo(cmbMonsterList, HeightOnly, DoubleWidth, frmMonsterAttackSim.hWnd)
cmbMonsterList.SelLength = 0

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("LoadMonsters")
Resume out:
End Sub

Private Sub Form_Resize()
'CheckPosition Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
frmMain.Show
frmMain.SetFocus
End Sub

Private Sub timWindowMove_Timer()
Call MonitorFormTimer(Me)
End Sub

Private Sub txtElementalResist_GotFocus(Index As Integer)
Call SelectAll(txtElementalResist(Index))
End Sub

Private Sub txtElementalResist_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtUserAC_GotFocus()
Call SelectAll(txtUserAC)
End Sub
Private Sub txtUserAC_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub
Private Sub txtUserDodge_GotFocus()
Call SelectAll(txtUserDodge)
End Sub
Private Sub txtUserDodge_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub
Private Sub txtUserMR_GotFocus()
Call SelectAll(txtUserMR)
End Sub
Private Sub txtUserMR_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub
Private Sub txtUserDR_GotFocus()
Call SelectAll(txtUserDR)
End Sub
Private Sub txtUserDR_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub
Private Sub cmdMRNote_Click()
MsgBox "Note: MR < 50 gives negative resistance when MR is taken into account.", vbInformation
End Sub
Private Sub chkDynamicRounds_Click()
If chkDynamicRounds.Value = 1 Then
    txtNumRounds.Enabled = False
    txtNumRounds.BackColor = &H8000000F
Else
    txtNumRounds.BackColor = &H80000005
    txtNumRounds.Enabled = True
End If
End Sub
Private Sub txtNumRounds_GotFocus()
Call SelectAll(txtNumRounds)
End Sub
Private Sub txtNumRounds_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Public Sub GotoMonster(ByVal nMonster As Long)
Dim x As Integer

For x = 0 To cmbMonsterList.ListCount - 1
    If cmbMonsterList.ItemData(x) = nMonster Then
        cmbMonsterList.ListIndex = x
        Exit For
    End If
Next x

End Sub




