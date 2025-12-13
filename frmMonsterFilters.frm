VERSION 5.00
Begin VB.Form frmMonsterFilters 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extra Monster Filters"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5295
   Icon            =   "frmMonsterFilters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExec 
      Cancel          =   -1  'True
      Caption         =   "Cancel/Close"
      Height          =   495
      Index           =   2
      Left            =   3840
      TabIndex        =   39
      Top             =   5460
      Width           =   1335
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "Save/Apply"
      Height          =   495
      Index           =   1
      Left            =   1980
      TabIndex        =   38
      Top             =   5460
      Width           =   1275
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "Save/Close"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   37
      Top             =   5460
      Width           =   1275
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   35
      Text            =   "0"
      ToolTipText     =   "Filter by monster damage output"
      Top             =   1860
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3060
      MaxLength       =   6
      TabIndex        =   33
      Text            =   "0"
      ToolTipText     =   "Filter by monster damage output"
      Top             =   1860
      Width           =   855
   End
   Begin VB.TextBox txtLairEXP 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   4200
      MaxLength       =   6
      TabIndex        =   31
      Text            =   "0"
      ToolTipText     =   "Filter by monster damage output"
      Top             =   540
      Width           =   795
   End
   Begin VB.TextBox txtGameLimit 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3180
      MaxLength       =   6
      TabIndex        =   29
      Text            =   "0"
      ToolTipText     =   "Filter by monster damage output"
      Top             =   540
      Width           =   795
   End
   Begin VB.TextBox txtMR 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   27
      Text            =   "0"
      ToolTipText     =   "Filter by monster damage output"
      Top             =   540
      Width           =   795
   End
   Begin VB.TextBox txtDR 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1140
      MaxLength       =   6
      TabIndex        =   25
      Text            =   "0"
      ToolTipText     =   "Filter by monster damage output"
      Top             =   540
      Width           =   795
   End
   Begin VB.TextBox txtAC 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   180
      MaxLength       =   6
      TabIndex        =   23
      Text            =   "0"
      ToolTipText     =   "Filter by monster damage output"
      Top             =   540
      Width           =   795
   End
   Begin VB.ComboBox cmbAbilityOp 
      Height          =   315
      Index           =   2
      ItemData        =   "frmMonsterFilters.frx":09AA
      Left            =   3180
      List            =   "frmMonsterFilters.frx":09B4
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox txtAbilityVal 
      Height          =   315
      Index           =   2
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   21
      Text            =   "0"
      Top             =   4920
      Width           =   495
   End
   Begin VB.ComboBox cmbAbilities 
      Height          =   315
      Index           =   2
      Left            =   660
      Sorted          =   -1  'True
      TabIndex        =   20
      Text            =   "cmbAbilities"
      Top             =   4920
      Width           =   2475
   End
   Begin VB.ComboBox cmbAbilityOp 
      Height          =   315
      Index           =   1
      ItemData        =   "frmMonsterFilters.frx":09C0
      Left            =   3180
      List            =   "frmMonsterFilters.frx":09CA
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   4560
      Width           =   615
   End
   Begin VB.TextBox txtAbilityVal 
      Height          =   315
      Index           =   1
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   18
      Text            =   "0"
      Top             =   4560
      Width           =   495
   End
   Begin VB.ComboBox cmbAbilities 
      Height          =   315
      Index           =   1
      Left            =   660
      Sorted          =   -1  'True
      TabIndex        =   17
      Text            =   "cmbAbilities"
      Top             =   4560
      Width           =   2475
   End
   Begin VB.ComboBox cmbAbilityOp 
      Height          =   315
      Index           =   0
      ItemData        =   "frmMonsterFilters.frx":09D6
      Left            =   3180
      List            =   "frmMonsterFilters.frx":09E0
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox txtAbilityVal 
      Height          =   315
      Index           =   0
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   14
      Text            =   "0"
      Top             =   4200
      Width           =   495
   End
   Begin VB.ComboBox cmbAbilities 
      Height          =   315
      Index           =   0
      ItemData        =   "frmMonsterFilters.frx":09EC
      Left            =   660
      List            =   "frmMonsterFilters.frx":09EE
      Sorted          =   -1  'True
      TabIndex        =   13
      Text            =   "cmbAbilities"
      Top             =   4200
      Width           =   2475
   End
   Begin VB.CheckBox chkAtkNoFear 
      Caption         =   "No Fear"
      Height          =   255
      Left            =   3660
      TabIndex        =   12
      ToolTipText     =   "Only Undead"
      Top             =   3360
      Width           =   1155
   End
   Begin VB.CheckBox chkAtkNoConfusion 
      Caption         =   "No Confusion"
      Height          =   255
      Left            =   3660
      TabIndex        =   11
      ToolTipText     =   "Only Undead"
      Top             =   2940
      Width           =   1515
   End
   Begin VB.CheckBox chkAtkNoPoison 
      Caption         =   "No Poison"
      Height          =   255
      Left            =   3660
      TabIndex        =   10
      ToolTipText     =   "Only Undead"
      Top             =   2520
      Width           =   1275
   End
   Begin VB.Frame fraCash 
      Caption         =   "Drops Coin"
      Height          =   1395
      Left            =   180
      TabIndex        =   3
      Top             =   2340
      Width           =   3315
      Begin VB.OptionButton optCash 
         Caption         =   "Runic"
         Height          =   315
         Index           =   5
         Left            =   1500
         TabIndex        =   9
         Top             =   960
         Width           =   915
      End
      Begin VB.OptionButton optCash 
         Caption         =   "Platinum+"
         Height          =   315
         Index           =   4
         Left            =   1500
         TabIndex        =   8
         Top             =   600
         Width           =   1275
      End
      Begin VB.OptionButton optCash 
         Caption         =   "Gold+"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1035
      End
      Begin VB.OptionButton optCash 
         Caption         =   "Silver+"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optCash 
         Caption         =   "Drops Any Coin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1500
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optCash 
         Caption         =   "No Filter"
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
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Non-Hostile VS Neutral/Good"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      ToolTipText     =   "Only Undead"
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CheckBox chkNonHostile_vEvil 
      Caption         =   "Non-Hostile VS Evil"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      ToolTipText     =   "Only Undead"
      Top             =   1500
      Width           =   1815
   End
   Begin VB.CheckBox chkIsUndead 
      Caption         =   "Is Undead"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      ToolTipText     =   "Only Undead"
      Top             =   1080
      Width           =   1155
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
      Left            =   4080
      TabIndex        =   36
      Top             =   1260
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
      Left            =   3060
      TabIndex        =   34
      Top             =   1260
      Width           =   855
   End
   Begin VB.Label lblLabelArray 
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
      Index           =   5
      Left            =   4260
      TabIndex        =   32
      Top             =   120
      Width           =   675
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
      Left            =   3180
      TabIndex        =   30
      Top             =   120
      Width           =   795
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
      Left            =   2160
      TabIndex        =   28
      Top             =   120
      Width           =   795
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
      Height          =   195
      Index           =   2
      Left            =   1140
      TabIndex        =   26
      Top             =   300
      Width           =   795
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
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   24
      Top             =   300
      Width           =   795
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
      Left            =   660
      TabIndex        =   16
      Top             =   3960
      Width           =   3675
   End
End
Attribute VB_Name = "frmMonsterFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbAbilities_Change(Index As Integer)

End Sub

Private Sub cmbAbilities_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = AutoComplete(cmbAbilities(Index), KeyAscii, False)
End Sub

Private Sub Form_Load()
On Error GoTo error:
Dim y As Integer, x As Integer, sAbilityList() As String

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
    Call ExpandCombo(cmbAbilities(y), HeightOnly, DoubleWidth, frmMonsterFilters.hWnd)
    cmbAbilities(y).ListIndex = 0
    cmbAbilityOp(y).ListIndex = 0
Next y


out:
On Error Resume Next
Exit Sub
error:
Call HandleError("Form_Load")
Resume out:
End Sub

Private Sub lblLabelArray_Click(Index As Integer)

End Sub

Private Sub txtAbilityVal_GotFocus(Index As Integer)
Call SelectAll(txtAbilityVal(Index))
End Sub

Private Sub txtAbilityVal_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub
