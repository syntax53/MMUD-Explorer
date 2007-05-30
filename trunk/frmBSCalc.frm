VERSION 5.00
Begin VB.Form frmBSCalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backstab Calculator"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5115
   Icon            =   "frmBSCalc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   5115
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNote 
      Caption         =   "Readme"
      Height          =   375
      Left            =   2640
      TabIndex        =   27
      Top             =   2400
      Width           =   915
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   1860
      TabIndex        =   26
      Top             =   2400
      Width           =   795
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   3720
      TabIndex        =   28
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdCopytoClip 
      Caption         =   "Cop&y to Clipboard"
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   25
      Top             =   2400
      Width           =   1635
   End
   Begin VB.Frame Frame2 
      Caption         =   "Stats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4995
      Begin VB.CommandButton cmdAlterLevel 
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
         Left            =   1860
         TabIndex        =   11
         Top             =   960
         Width           =   315
      End
      Begin VB.TextBox txtStrength 
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   9
         Text            =   "100"
         Top             =   960
         Width           =   675
      End
      Begin VB.CommandButton cmdAlterLevel 
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
         Left            =   1560
         TabIndex        =   10
         Top             =   960
         Width           =   315
      End
      Begin VB.CheckBox chkClassStealth 
         Caption         =   "Class Stealth"
         Height          =   315
         Left            =   3480
         TabIndex        =   24
         Top             =   1380
         Width           =   1335
      End
      Begin VB.CommandButton cmdAlterLevel 
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
         Left            =   1860
         TabIndex        =   8
         Top             =   600
         Width           =   315
      End
      Begin VB.CommandButton cmdAlterLevel 
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
         Left            =   4500
         TabIndex        =   18
         Top             =   600
         Width           =   315
      End
      Begin VB.CommandButton cmdAlterLevel 
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
         Left            =   4200
         TabIndex        =   17
         Top             =   600
         Width           =   315
      End
      Begin VB.CommandButton cmdAlterLevel 
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
         Left            =   4500
         TabIndex        =   21
         Top             =   960
         Width           =   315
      End
      Begin VB.CommandButton cmdAlterLevel 
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
         Left            =   4500
         TabIndex        =   15
         Top             =   240
         Width           =   315
      End
      Begin VB.CommandButton cmdAlterLevel 
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
         Left            =   4200
         TabIndex        =   20
         Top             =   960
         Width           =   315
      End
      Begin VB.CommandButton cmdAlterLevel 
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
         Left            =   4200
         TabIndex        =   14
         Top             =   240
         Width           =   315
      End
      Begin VB.CommandButton cmdAlterLevel 
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
         Left            =   1560
         TabIndex        =   7
         Top             =   600
         Width           =   315
      End
      Begin VB.CommandButton cmdAlterLevel 
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
         Left            =   1860
         TabIndex        =   4
         Top             =   240
         Width           =   315
      End
      Begin VB.CommandButton cmdAlterLevel 
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
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   315
      End
      Begin VB.CommandButton cmdJump 
         Caption         =   ">"
         Height          =   315
         Left            =   3120
         TabIndex        =   23
         Top             =   1380
         Width           =   195
      End
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   2
         Text            =   "50"
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox txtStealth 
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   6
         Text            =   "100"
         Top             =   600
         Width           =   675
      End
      Begin VB.TextBox txtBSMinDMG 
         Height          =   285
         Left            =   3480
         MaxLength       =   5
         TabIndex        =   16
         Text            =   "0"
         Top             =   600
         Width           =   675
      End
      Begin VB.TextBox txtBSMaxDMG 
         Height          =   285
         Left            =   3480
         MaxLength       =   6
         TabIndex        =   19
         Text            =   "0"
         Top             =   960
         Width           =   675
      End
      Begin VB.ComboBox cmbWeapon 
         Height          =   315
         ItemData        =   "frmBSCalc.frx":0CCA
         Left            =   840
         List            =   "frmBSCalc.frx":0CCC
         Sorted          =   -1  'True
         TabIndex        =   22
         Text            =   "cmbWeapon"
         Top             =   1380
         Width           =   2235
      End
      Begin VB.TextBox txtMaxDMG 
         Height          =   285
         Left            =   3480
         MaxLength       =   4
         TabIndex        =   13
         Text            =   "0"
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Strength"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   600
      End
      Begin VB.Label lblDMG 
         Alignment       =   2  'Center
         Caption         =   "00 - 00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   1800
         Width           =   4755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Level"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Stealth"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "BS Min DMG"
         Height          =   195
         Index           =   3
         Left            =   2400
         TabIndex        =   29
         Top             =   600
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "BS Max DMG"
         Height          =   195
         Index           =   4
         Left            =   2400
         TabIndex        =   30
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Weapon"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Max Damage"
         Height          =   195
         Index           =   6
         Left            =   2400
         TabIndex        =   12
         Top             =   240
         Width           =   945
      End
   End
   Begin VB.Timer timMouseDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   60
   End
End
Attribute VB_Name = "frmBSCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Dim bMouseDown As Boolean
Dim bDontRefresh As Boolean

Private Const WM_SETREDRAW = &HB
Private msOldString As String ' module level global
Private miStart As Integer    ' module level global
Private miLength As Integer   ' module level global

'Dim objToolTip As clsToolTip


Private Sub chkClassStealth_Click()
Call CalcBS
End Sub

Private Sub cmbWeapon_Click()

'objToolTip.DelToolTip cmdJump.hWnd

Call CalcBS
End Sub

Private Sub cmbWeapon_KeyPress(KeyAscii As Integer)
KeyAscii = AutoComplete(cmbWeapon, KeyAscii)
End Sub

Private Sub cmdNote_Click()
MsgBox "The max damage and strength fields are separate.  That means if you have 60 " _
    & "strength (which would incur a +1 max damage) you would still put 1 in the max damage " _
    & "field.  The strength field is here because every 10 strength over 100 adds +2 min " _
    & "damage to the weapon.", vbInformation
End Sub

Private Sub cmdReset_Click()
Call Form_Load
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
On Error GoTo Error:
Dim x As Integer, nCombat As Integer, bClassStealth As Boolean

'Set objToolTip = New clsToolTip

bDontRefresh = True

Call LoadWeapons
Call GetStealth

If Val(frmMain.txtCharStats(0).Text) > 0 Then
    txtStrength.Text = Val(frmMain.txtCharStats(0).Text)
End If

If Val(frmMain.txtGlobalLevel(0).Text) > 0 Then
    txtLevel.Text = Val(frmMain.txtGlobalLevel(0).Text)
End If

If Not Val(frmMain.txtStat(11).Text) = 0 Then
    txtMaxDMG.Text = Val(frmMain.txtStat(11).Text)
End If

If frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex) > 0 Then
    bClassStealth = GetClassStealth(frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex))
    If bClassStealth Then
        chkClassStealth.Value = 1
    Else
        chkClassStealth.Value = 0
    End If
End If

If Not Val(frmMain.txtStat(14).Text) = 0 Then txtBSMinDMG.Text = Val(frmMain.txtStat(14).Text)
If Not Val(frmMain.txtStat(15).Text) = 0 Then txtBSMaxDMG.Text = Val(frmMain.txtStat(15).Text)

If nEquippedItem(16) > 0 Then
    Call GotoWeapon(nEquippedItem(16))
End If

bDontRefresh = False
Call CalcBS

Exit Sub
Error:
Call HandleError("BSCalc_Load")
Resume Next
End Sub

Private Sub GetStealth()
Dim x As Integer, sFile As String, sSectionName As String, nSpell As Long, y As Long

On Error GoTo Error:

sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ")
If frmMain.bCharLoaded Then
    sFile = ReadINI(sSectionName, "LastCharFile")
    If Not FileExists(sFile) Then
        sFile = ""
    Else
        sSectionName = "PlayerInfo"
    End If
End If

txtStealth.Text = Val(ReadINI(sSectionName, "BSStealth", sFile))
    
Exit Sub
Error:
Call HandleError("GetStealth")
End Sub

Private Sub WriteStealth()
Dim x As Integer, sFile As String, sSectionName As String, nSpell As Long, y As Long

On Error GoTo Error:

sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ")
If frmMain.bCharLoaded Then
    sFile = ReadINI(sSectionName, "LastCharFile")
    If Not FileExists(sFile) Then
        sFile = ""
    Else
        sSectionName = "PlayerInfo"
    End If
End If

Call WriteINI(sSectionName, "BSStealth", Val(txtStealth.Text), sFile)
    
Exit Sub
Error:
Call HandleError("WriteStealth")
End Sub

Private Sub cmdAlterLevel_Click(Index As Integer)
    
If Not bMouseDown Then Call AlterLevel(Index)

End Sub

Private Sub AlterLevel(ByVal Index As Integer)

On Error GoTo Error:

If Index = 0 Then 'minus LEVEL
    If Val(txtLevel.Text) <= 0 Then
        txtLevel.Text = 0
    Else
        txtLevel.Text = Val(txtLevel.Text) - 1
    End If
ElseIf Index = 1 Then 'plus
    If Val(txtLevel.Text) >= 1000 Then
        txtLevel.Text = 1000
    Else
        txtLevel.Text = Val(txtLevel.Text) + 1
    End If
ElseIf Index = 2 Then 'minus stea
    If Val(txtStealth.Text) <= 0 Then
        txtStealth.Text = 0
    Else
        txtStealth.Text = Val(txtStealth.Text) - 1
    End If
ElseIf Index = 3 Then 'plus
    If Val(txtStealth.Text) >= 1000 Then
        txtStealth.Text = 1000
    Else
        txtStealth.Text = Val(txtStealth.Text) + 1
    End If
ElseIf Index = 4 Then 'minus max dmg
    If Val(txtMaxDMG.Text) < -1000 Then
        txtMaxDMG.Text = -1000
    Else
        txtMaxDMG.Text = Val(txtMaxDMG.Text) - 1
    End If
ElseIf Index = 5 Then 'plus
    If Val(txtMaxDMG.Text) >= 1000 Then
        txtMaxDMG.Text = 1000
    Else
        txtMaxDMG.Text = Val(txtMaxDMG.Text) + 1
    End If
ElseIf Index = 6 Then 'minus bs min
    If Val(txtBSMinDMG.Text) < -1000 Then
        txtBSMinDMG.Text = -1000
    Else
        txtBSMinDMG.Text = Val(txtBSMinDMG.Text) - 1
    End If
ElseIf Index = 7 Then 'plus
    If Val(txtBSMinDMG.Text) >= 1000 Then
        txtBSMinDMG.Text = 1000
    Else
        txtBSMinDMG.Text = Val(txtBSMinDMG.Text) + 1
    End If
ElseIf Index = 8 Then 'minus bs max
    If Val(txtBSMaxDMG.Text) < -1000 Then
        txtBSMaxDMG.Text = -1000
    Else
        txtBSMaxDMG.Text = Val(txtBSMaxDMG.Text) - 1
    End If
ElseIf Index = 9 Then 'plus
    If Val(txtBSMaxDMG.Text) >= 1000 Then
        txtBSMaxDMG.Text = 1000
    Else
        txtBSMaxDMG.Text = Val(txtBSMaxDMG.Text) + 1
    End If
ElseIf Index = 10 Then 'minus str max
    If Val(txtStrength.Text) < -1000 Then
        txtStrength.Text = -1000
    Else
        txtStrength.Text = Val(txtStrength.Text) - 1
    End If
ElseIf Index = 11 Then 'plus
    If Val(txtStrength.Text) >= 1000 Then
        txtStrength.Text = 1000
    Else
        txtStrength.Text = Val(txtStrength.Text) + 1
    End If
End If
'Call CalcBS

Exit Sub

Error:
Call HandleError("AlterLevel")
    
End Sub
Private Sub cmdAlterLevel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

bMouseDown = True

Do While bMouseDown
    timMouseDown.Enabled = True
    Call AlterLevel(Index)
    Do While timMouseDown.Enabled
        DoEvents
    Loop
Loop

'bMouseDown = True
'
'Do While bMouseDown
'    timMouseDown.Enabled = True
'    If Index = 0 Then 'minus LEVEL
'        If Val(txtLevel.Text) <= 0 Then
'            txtLevel.Text = 0
'        Else
'            txtLevel.Text = Val(txtLevel.Text) - 1
'        End If
'    ElseIf Index = 1 Then 'plus
'        If Val(txtLevel.Text) >= 9999 Then
'            txtLevel.Text = 9999
'        Else
'            txtLevel.Text = Val(txtLevel.Text) + 1
'        End If
'    ElseIf Index = 2 Then 'minus AGL
'        If Val(txtStealth.Text) <= 0 Then
'            txtStealth.Text = 0
'        Else
'            txtStealth.Text = Val(txtStealth.Text) - 1
'        End If
'    ElseIf Index = 3 Then 'plus
'        If Val(txtStealth.Text) >= 9999 Then
'            txtStealth.Text = 9999
'        Else
'            txtStealth.Text = Val(txtStealth.Text) + 1
'        End If
'    ElseIf Index = 4 Then 'minus STR
'        If Val(txtMaxDMG.Text) <= 0 Then
'            txtMaxDMG.Text = 0
'        Else
'            txtMaxDMG.Text = Val(txtMaxDMG.Text) - 1
'        End If
'    ElseIf Index = 5 Then 'plus
'        If Val(txtMaxDMG.Text) >= 9999 Then
'            txtMaxDMG.Text = 9999
'        Else
'            txtMaxDMG.Text = Val(txtMaxDMG.Text) + 1
'        End If
'    ElseIf Index = 6 Then 'minus ENC
'        If Val(txtBSMinDMG.Text) <= 0 Then
'            txtBSMinDMG.Text = 0
'        Else
'            txtBSMinDMG.Text = Val(txtBSMinDMG.Text) - 25
'        End If
'    ElseIf Index = 7 Then 'plus
'        If Val(txtBSMinDMG.Text) >= 99999 Then
'            txtBSMinDMG.Text = 99999
'        Else
'            txtBSMinDMG.Text = Val(txtBSMinDMG.Text) + 25
'        End If
'    ElseIf Index = 8 Then 'minus MAX ENC
'        If Val(txtBSMaxDMG.Text) <= 0 Then
'            txtBSMaxDMG.Text = 0
'        Else
'            txtBSMaxDMG.Text = Val(txtBSMaxDMG.Text) - 1
'        End If
'    ElseIf Index = 9 Then 'plus
'        If Val(txtBSMaxDMG.Text) >= 99999 Then
'            txtBSMaxDMG.Text = 99999
'        Else
'            txtBSMaxDMG.Text = Val(txtBSMaxDMG.Text) + 1
'        End If
'    End If
'    Call CalcBS
'    Do While timMouseDown.Enabled
'        DoEvents
'    Loop
'Loop

End Sub

Private Sub cmdAlterLevel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
bMouseDown = False
DoEvents
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdCopytoClip_Click(Index As Integer)
Dim str As String
On Error GoTo Error:

If cmbWeapon.ListIndex < 0 Then Exit Sub

tabItems.Seek "=", cmbWeapon.ItemData(cmbWeapon.ListIndex)
If tabItems.NoMatch Then Exit Sub

str = "BS Damage: " & lblDMG.Caption & vbCrLf

str = str & tabItems.Fields("Name") & ": " _
    & tabItems.Fields("Min") & " - " & tabItems.Fields("Max")

str = str & vbCrLf & "Strength: " & Val(txtStrength.Text)

If Not Val(txtMaxDMG.Text) = 0 Then
    If Val(txtMaxDMG.Text) > 0 Then
        str = str & vbCrLf & "Max Damage: +" & txtMaxDMG.Text
    Else
        str = str & vbCrLf & "Max Damage: " & txtMaxDMG.Text
    End If
End If
 
If Not Val(txtBSMinDMG.Text) = 0 Then
    If Val(txtBSMinDMG.Text) > 0 Then
        str = str & vbCrLf & "MinBS: +" & txtBSMinDMG.Text
    Else
        str = str & vbCrLf & "MinBS: " & txtBSMinDMG.Text
    End If
End If

If Not Val(txtBSMaxDMG.Text) = 0 Then
    If Not Val(txtBSMinDMG.Text) = 0 Then
        str = str & ", "
    Else
        str = str & vbCrLf
    End If
    
    If Val(txtBSMaxDMG.Text) > 0 Then
        str = str & "MaxBS: +" & txtBSMaxDMG.Text
    Else
        str = str & "MaxBS: " & txtBSMaxDMG.Text
    End If
End If

 str = str & vbCrLf & "Level: " & txtLevel.Text & ", Stealth: " & txtStealth.Text _
    & vbCrLf & "Class Stealth: " _
    & IIf(chkClassStealth.Value = 1, "Yes", "No")

If Not str = "" Then
    Clipboard.clear
    Clipboard.SetText str
End If

Exit Sub

Error:
Call HandleError("cmdCopytoClip_Click")
End Sub

Private Sub cmdJump_Click()
If cmbWeapon.ListIndex < 0 Then Exit Sub
Call frmMain.GotoItem(cmbWeapon.ItemData(cmbWeapon.ListIndex))
End Sub

Public Sub GotoWeapon(ByVal nItem As Long)
Dim x As Integer

For x = 0 To cmbWeapon.ListCount - 1
    If cmbWeapon.ItemData(x) = nItem Then
        cmbWeapon.ListIndex = x
        Exit For
    End If
Next x

End Sub

Private Sub LoadWeapons()
On Error GoTo Error:
If tabItems.RecordCount = 0 Then Exit Sub

Me.MousePointer = vbHourglass
tabItems.MoveFirst
DoEvents

cmbWeapon.clear

Do Until tabItems.EOF
    If bOnlyInGame And tabItems.Fields("In Game") = 0 Then GoTo skip:
    If tabItems.Fields("ItemType") = 1 Then
        cmbWeapon.AddItem (tabItems.Fields("Name") & " (" & tabItems.Fields("Number") & ")")
        cmbWeapon.ItemData(cmbWeapon.NewIndex) = tabItems.Fields("Number")
    End If
skip:
    tabItems.MoveNext
Loop

If cmbWeapon.ListCount > 0 Then
    cmbWeapon.ListIndex = 0
    Call AutoSizeDropDownWidth(cmbWeapon)
    Call ExpandCombo(cmbWeapon, HeightOnly, DoubleWidth, Frame2.hWnd)
    cmbWeapon.SelLength = 0
End If

Me.MousePointer = vbDefault
Exit Sub
Error:
Call HandleError("SwingCalc_LoadItems")
Me.MousePointer = vbDefault

End Sub

Private Sub CalcBS()
Dim nMinDmg As Long, nMaxDmg As Long, nBSStealth As Integer, nDMG_Mod As Integer
Dim x As Integer, bClassStealth As Boolean, nMaxDMGBonus As Integer
On Error GoTo Error:

If bDontRefresh Then Exit Sub

If cmbWeapon.ListIndex < 0 Then Exit Sub

tabItems.Index = "pkItems"
tabItems.Seek "=", cmbWeapon.ItemData(cmbWeapon.ListIndex)
If Not tabItems.NoMatch Then
    For x = 0 To 19
        If tabItems.Fields("Abil-" & x) = 116 Then Exit For 'bs accu
        If x = 19 Then
            lblDMG.Caption = "No BS"
            Exit Sub
        End If
    Next x

    If Val(txtStealth.Text) > 1000 Then txtStealth.Text = 1000
    nBSStealth = Val(txtStealth.Text)
    nMaxDMGBonus = Val(txtMaxDMG.Text)
    If chkClassStealth.Value = 1 Then bClassStealth = True
    
    nMinDmg = tabItems.Fields("Min")
    nMaxDmg = tabItems.Fields("Max")
     
    If Val(txtStrength.Text) > 109 Then
        nMinDmg = nMinDmg + ((Fix(Val(txtStrength.Text) / 10) - 10) * 2)
    End If
    
    nMaxDmg = nMaxDmg + nMaxDMGBonus
    
    If nMaxDmg < nMinDmg Then nMaxDmg = nMinDmg
    
    nDMG_Mod = Val(txtBSMinDMG.Text)
    nMinDmg = CalcBSDamage(Val(txtLevel.Text), nBSStealth, _
        nMinDmg, nDMG_Mod, bClassStealth) '+ 12
    
    nDMG_Mod = Val(txtBSMaxDMG.Text)
    nMaxDmg = CalcBSDamage(Val(txtLevel.Text), nBSStealth, _
        nMaxDmg, nDMG_Mod, bClassStealth)
    
    If nMaxDmg < nMinDmg Then nMaxDmg = nMinDmg
    
    lblDMG.Caption = nMinDmg & " - " & nMaxDmg & " (AVG: " & Fix((nMaxDmg + nMinDmg) / 2) & ")"
End If

Exit Sub

Error:
Call HandleError("CalcBS")

End Sub
Private Sub Form_Unload(Cancel As Integer)
If Not bAppTerminating Then frmMain.SetFocus
Call WriteStealth
'Set objToolTip = Nothing
End Sub


Private Sub timMouseDown_Timer()
timMouseDown.Enabled = False
End Sub

Private Sub txtAgility_Change()
Call CalcBS
End Sub

Private Sub txtBSMaxDMG_Change()
Call CalcBS
End Sub

Private Sub txtBSMinDMG_Change()
Call CalcBS
End Sub

Private Sub txtLevel_Change()
Call CalcBS
End Sub

Private Sub txtMaxDMG_Change()
Call CalcBS
End Sub

Private Sub txtStealth_Change()
Call CalcBS
End Sub

Private Sub txtStealth_GotFocus()
Call SelectAll(txtStealth)
End Sub

Private Sub txtStealth_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtStealth_KeyUp(KeyCode As Integer, Shift As Integer)
'Call CalcBS
End Sub

Private Sub txtBSMinDMG_GotFocus()
Call SelectAll(txtBSMinDMG)
End Sub

Private Sub txtBSMinDMG_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtBSMinDMG_KeyUp(KeyCode As Integer, Shift As Integer)
'Call CalcBS
End Sub

Private Sub txtLevel_GotFocus()
Call SelectAll(txtLevel)
End Sub

Private Sub txtLevel_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtLevel_KeyUp(KeyCode As Integer, Shift As Integer)
'Call CalcBS
End Sub

Private Sub txtBSMaxDMG_GotFocus()
Call SelectAll(txtBSMaxDMG)
End Sub

Private Sub txtBSMaxDMG_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtBSMaxDMG_KeyUp(KeyCode As Integer, Shift As Integer)
'Call CalcBS
End Sub

Private Sub txtMaxDMG_GotFocus()
Call SelectAll(txtMaxDMG)
End Sub

Private Sub txtMaxDMG_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtMaxDMG_KeyUp(KeyCode As Integer, Shift As Integer)
'Call CalcBS
End Sub

Private Sub txtStrength_Change()
Call CalcBS
End Sub

Private Sub txtStrength_GotFocus()
Call SelectAll(txtStrength)

End Sub
