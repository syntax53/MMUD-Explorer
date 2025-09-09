VERSION 5.00
Begin VB.Form frmBSCalc 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backstab Calculator"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   Icon            =   "frmBSCalc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6000
   Begin VB.Timer timCalc 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   840
   End
   Begin VB.Timer timButtonPress 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   420
   End
   Begin VB.Timer timWindowMove 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdNote 
      Caption         =   "Help"
      Height          =   375
      Left            =   3180
      TabIndex        =   27
      Top             =   2400
      Width           =   1035
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   2040
      TabIndex        =   26
      Top             =   2400
      Width           =   1035
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   4560
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
      Width           =   5835
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
         Width           =   1275
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
      Begin VB.Label lblStealthAdj 
         AutoSize        =   -1  'True
         Caption         =   "xxx"
         Height          =   195
         Left            =   4920
         TabIndex        =   37
         Top             =   1440
         Width           =   225
      End
      Begin VB.Label lblBSMaxAdj 
         AutoSize        =   -1  'True
         Caption         =   "xxx"
         Height          =   195
         Left            =   4920
         TabIndex        =   36
         Top             =   960
         Width           =   225
      End
      Begin VB.Label lblBSMinAdj 
         AutoSize        =   -1  'True
         Caption         =   "xxx"
         Height          =   195
         Left            =   4920
         TabIndex        =   35
         Top             =   600
         Width           =   225
      End
      Begin VB.Label lblMaxAdj 
         AutoSize        =   -1  'True
         Caption         =   "xxx"
         Height          =   195
         Left            =   4920
         TabIndex        =   34
         Top             =   240
         Width           =   225
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
         Width           =   5595
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
End
Attribute VB_Name = "frmBSCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Dim tWindowSize As WindowSizeProperties

Dim bMouseDown As Boolean
Dim bDontRefresh As Boolean

Public nLastPosTop As Long
Public nLastPosLeft As Long
Public nLastPosMoved As Long
Public nLastPosMonitor As Long

Public nLastTimerTop As Long
Public nLastTimerLeft As Long

Private Sub chkClassStealth_Click()
Call CalcBS
End Sub

Private Sub cmbWeapon_Click()

'objToolTip.DelToolTip cmdJump.hWnd

timCalc.Enabled = False: timCalc.Enabled = True: 'Call CalcBS
End Sub

Private Sub cmbWeapon_KeyPress(KeyAscii As Integer)
KeyAscii = AutoComplete(cmbWeapon, KeyAscii)
End Sub

Private Sub cmdNote_Click()
MsgBox "Max damage bonus from strength is not automatically added. Specify all max damage bonus (including anything from strength) in that field. " _
    & "+MIN damage from strength IS automatically added." & vbCrLf & vbCrLf _
    & "The tool will automatically add related stats from the selected item to the stats for the calculation. " _
    & "If the global filter is enabled and the chosen weapon is different from the equipped weapon, those stats will be subtracted. " _
    & "To prevent any subtraction of stats, disable the global filter on the main window." & vbCrLf & vbCrLf _
    & "Note: Stat/input fields are not updated once the tool is loaded. Click reset to refresh stats.", vbInformation
End Sub

Private Sub cmdReset_Click()
Call Form_Load
End Sub


Private Sub Form_Load()
On Error GoTo error:
Dim bClassStealth As Boolean

Call SetWindowLong(Me.hWnd, GWL_HWNDPARENT, 0)

'stop windows from resizing fixed-size windows when changing dpi
If bDPIAwareMode Then Call SubclassFormMinMaxSize(Me, tWindowSize, True)

bDontRefresh = True
Me.MousePointer = vbHourglass
DoEvents

lblMaxAdj.Caption = ""
lblBSMinAdj.Caption = ""
lblBSMaxAdj.Caption = ""
lblStealthAdj.Caption = ""

Call LoadWeapons

If val(frmMain.txtCharStats(0).Text) > 0 Then
    txtStrength.Text = val(frmMain.txtCharStats(0).Text)
End If

If val(frmMain.txtGlobalLevel(0).Text) > 0 Then
    txtLevel.Text = val(frmMain.txtGlobalLevel(0).Text)
End If

If Not val(frmMain.lblInvenCharStat(11).Caption) = 0 Then
    txtMaxDMG.Text = val(frmMain.lblInvenCharStat(11).Caption)
End If

If Not val(frmMain.lblInvenCharStat(19).Caption) = 0 Then
    txtStealth.Text = val(frmMain.lblInvenCharStat(19).Caption)
End If

If frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex) > 0 Then
    bClassStealth = GetClassStealth(frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex))
    If bClassStealth Then
        chkClassStealth.Value = 1
    Else
        chkClassStealth.Value = 0
    End If
End If

If Not val(frmMain.lblInvenCharStat(14).Caption) = 0 Then txtBSMinDMG.Text = val(frmMain.lblInvenCharStat(14).Caption)
If Not val(frmMain.lblInvenCharStat(15).Caption) = 0 Then txtBSMaxDMG.Text = val(frmMain.lblInvenCharStat(15).Caption)

If nEquippedItem(16) > 0 Then
    Call GotoWeapon(nEquippedItem(16))
End If

If Not Me.Visible Then
    If frmMain.WindowState = vbMinimized Then
        Me.Top = (Screen.Height - Me.Height) / 2
        Me.Left = (Screen.Width - Me.Width) / 2
    Else
        Me.Left = frmMain.Left + ((frmMain.Width - Me.Width) / 2)
        Me.Top = frmMain.Top + ((frmMain.Height - Me.Height) / 2)
    End If
End If
timWindowMove.Enabled = True
bDontRefresh = False
Call CalcBS

Me.MousePointer = vbDefault
Exit Sub
error:
Call HandleError("BSCalc_Load")
Resume Next
End Sub

'Private Sub GetStealth()
'Dim sFile As String, sSectionName As String, sCharFile As String
'
'On Error GoTo error:
'
'sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ")
'
'sCharFile = ReadINI(sSectionName, "LastCharFile")
'If Len(sSessionLastCharFile) > 0 Then sCharFile = sSessionLastCharFile
'If Not FileExists(sCharFile) Then
'    sCharFile = ""
'    sSessionLastCharFile = ""
'End If
'
'If frmMain.bCharLoaded Then
'    sFile = sCharFile
'    If Not FileExists(sFile) Then
'        sFile = ""
'        sSessionLastCharFile = ""
'    Else
'        sSectionName = "PlayerInfo"
'    End If
'End If
'
'txtStealth.Text = val(ReadINI(sSectionName, "BSStealth", sFile))
'
'Exit Sub
'error:
'Call HandleError("GetStealth")
'End Sub

Private Sub WriteStealth()
Dim sFile As String, sSectionName As String, sCharFile As String

On Error GoTo error:

sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ")

sCharFile = ReadINI(sSectionName, "LastCharFile")
If Len(sSessionLastCharFile) > 0 Then sCharFile = sSessionLastCharFile
If Not FileExists(sCharFile) Then
    sCharFile = ""
    sSessionLastCharFile = ""
End If

If frmMain.bCharLoaded Then
    sFile = sCharFile
    If Not FileExists(sFile) Then
        sFile = ""
        sSessionLastCharFile = ""
    Else
        sSectionName = "PlayerInfo"
    End If
End If

Call WriteINI(sSectionName, "BSStealth", val(txtStealth.Text), sFile)
    
Exit Sub
error:
Call HandleError("WriteStealth")
End Sub

Private Sub cmdAlterLevel_Click(Index As Integer)
    
If Not bMouseDown Then Call AlterLevel(Index)

End Sub

Private Sub AlterLevel(ByVal Index As Integer)

On Error GoTo error:

If Index = 0 Then 'minus LEVEL
    If val(txtLevel.Text) <= 0 Then
        txtLevel.Text = 0
    Else
        txtLevel.Text = val(txtLevel.Text) - 1
    End If
ElseIf Index = 1 Then 'plus
    If val(txtLevel.Text) >= 1000 Then
        txtLevel.Text = 1000
    Else
        txtLevel.Text = val(txtLevel.Text) + 1
    End If
ElseIf Index = 2 Then 'minus stea
    If val(txtStealth.Text) <= 0 Then
        txtStealth.Text = 0
    Else
        txtStealth.Text = val(txtStealth.Text) - 1
    End If
ElseIf Index = 3 Then 'plus
    If val(txtStealth.Text) >= 1000 Then
        txtStealth.Text = 1000
    Else
        txtStealth.Text = val(txtStealth.Text) + 1
    End If
ElseIf Index = 4 Then 'minus max dmg
    If val(txtMaxDMG.Text) < -1000 Then
        txtMaxDMG.Text = -1000
    Else
        txtMaxDMG.Text = val(txtMaxDMG.Text) - 1
    End If
ElseIf Index = 5 Then 'plus
    If val(txtMaxDMG.Text) >= 1000 Then
        txtMaxDMG.Text = 1000
    Else
        txtMaxDMG.Text = val(txtMaxDMG.Text) + 1
    End If
ElseIf Index = 6 Then 'minus bs min
    If val(txtBSMinDMG.Text) < -1000 Then
        txtBSMinDMG.Text = -1000
    Else
        txtBSMinDMG.Text = val(txtBSMinDMG.Text) - 1
    End If
ElseIf Index = 7 Then 'plus
    If val(txtBSMinDMG.Text) >= 1000 Then
        txtBSMinDMG.Text = 1000
    Else
        txtBSMinDMG.Text = val(txtBSMinDMG.Text) + 1
    End If
ElseIf Index = 8 Then 'minus bs max
    If val(txtBSMaxDMG.Text) < -1000 Then
        txtBSMaxDMG.Text = -1000
    Else
        txtBSMaxDMG.Text = val(txtBSMaxDMG.Text) - 1
    End If
ElseIf Index = 9 Then 'plus
    If val(txtBSMaxDMG.Text) >= 1000 Then
        txtBSMaxDMG.Text = 1000
    Else
        txtBSMaxDMG.Text = val(txtBSMaxDMG.Text) + 1
    End If
ElseIf Index = 10 Then 'minus str max
    If val(txtStrength.Text) < -1000 Then
        txtStrength.Text = -1000
    Else
        txtStrength.Text = val(txtStrength.Text) - 1
    End If
ElseIf Index = 11 Then 'plus
    If val(txtStrength.Text) >= 1000 Then
        txtStrength.Text = 1000
    Else
        txtStrength.Text = val(txtStrength.Text) + 1
    End If
End If
'Call CalcBS

Exit Sub

error:
Call HandleError("AlterLevel")
    
End Sub
Private Sub cmdAlterLevel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

bMouseDown = True

Do While bMouseDown
    timButtonPress.Enabled = True
    Call AlterLevel(Index)
    Do While timButtonPress.Enabled
        DoEvents
    Loop
Loop

'bMouseDown = True
'
'Do While bMouseDown
'    timButtonPress.Enabled = True
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
'    Do While timButtonPress.Enabled
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
On Error GoTo error:

If cmbWeapon.ListIndex < 0 Then Exit Sub

tabItems.Seek "=", cmbWeapon.ItemData(cmbWeapon.ListIndex)
If tabItems.NoMatch Then
    tabItems.MoveFirst
    Exit Sub
End If

str = "BS Damage: " & lblDMG.Caption & vbCrLf

str = str & tabItems.Fields("Name") & ": " _
    & tabItems.Fields("Min") & " - " & tabItems.Fields("Max")

str = str & vbCrLf & "Strength: " & val(txtStrength.Text)

If Not val(txtMaxDMG.Text) = 0 Then
    If val(txtMaxDMG.Text) > 0 Then
        str = str & vbCrLf & "Max Damage: +" & txtMaxDMG.Text
    Else
        str = str & vbCrLf & "Max Damage: " & txtMaxDMG.Text
    End If
End If
 
If Not val(txtBSMinDMG.Text) = 0 Then
    If val(txtBSMinDMG.Text) > 0 Then
        str = str & vbCrLf & "MinBS: +" & txtBSMinDMG.Text
    Else
        str = str & vbCrLf & "MinBS: " & txtBSMinDMG.Text
    End If
End If

If Not val(txtBSMaxDMG.Text) = 0 Then
    If Not val(txtBSMinDMG.Text) = 0 Then
        str = str & ", "
    Else
        str = str & vbCrLf
    End If
    
    If val(txtBSMaxDMG.Text) > 0 Then
        str = str & "MaxBS: +" & txtBSMaxDMG.Text
    Else
        str = str & "MaxBS: " & txtBSMaxDMG.Text
    End If
End If

str = str & vbCrLf & "Level: " & txtLevel.Text & ", Stealth: " & txtStealth.Text _
    & vbCrLf & "Class Stealth: " _
    & IIf(chkClassStealth.Value = 1, "Yes", "No")

If Not str = "" Then
    'Clipboard.clear
    'Clipboard.SetText str
    Call SetClipboardText(str)
End If

Exit Sub

error:
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
On Error GoTo error:
Dim bHasBS As Boolean, x As Integer

If tabItems.RecordCount = 0 Then Exit Sub

tabItems.MoveFirst
DoEvents

cmbWeapon.clear

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
            cmbWeapon.AddItem (tabItems.Fields("Name") & " (" & tabItems.Fields("Number") & ")")
            cmbWeapon.ItemData(cmbWeapon.NewIndex) = tabItems.Fields("Number")
        End If
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

Exit Sub
error:
Call HandleError("SwingCalc_LoadItems")
End Sub

Private Sub CalcBS()
Dim nMinDmg As Long, nMaxDmg As Long, nStealth As Integer, nPlusBSmindmg As Integer, nPlusBSmaxdmg As Integer
Dim x As Integer, bClassStealth As Boolean, nPlusMaxDamage As Integer, nMinStrBonus As Integer
Dim nWeaponNumber As Long, tStatSlot As tAbilityToStatSlot, nVal As Long
On Error GoTo error:

If bDontRefresh Then Exit Sub
If cmbWeapon.ListIndex < 0 Then Exit Sub
nWeaponNumber = cmbWeapon.ItemData(cmbWeapon.ListIndex)

lblMaxAdj.Caption = ""
lblBSMinAdj.Caption = ""
lblBSMaxAdj.Caption = ""
lblStealthAdj.Caption = ""

tabItems.Index = "pkItems"
tabItems.Seek "=", nWeaponNumber
If Not tabItems.NoMatch Then
    For x = 0 To 19
        If tabItems.Fields("Abil-" & x) = 116 Then Exit For 'bs accu
        If x = 19 Then
            lblDMG.Caption = "No BS"
            Exit Sub
        End If
    Next x
    
    If val(txtBSMinDMG.Text) > 1000 Then txtBSMinDMG.Text = 1000
    If val(txtBSMinDMG.Text) < 0 Then txtBSMinDMG.Text = 0
    nPlusBSmindmg = val(txtBSMinDMG.Text)
    
    If val(txtBSMaxDMG.Text) > 1000 Then txtBSMaxDMG.Text = 1000
    If val(txtBSMaxDMG.Text) < 0 Then txtBSMaxDMG.Text = 0
    nPlusBSmaxdmg = val(txtBSMaxDMG.Text)
    
    If val(txtLevel.Text) > 1000 Then txtLevel.Text = 1000
    If val(txtLevel.Text) < 0 Then txtLevel.Text = 0
    
    If val(txtStealth.Text) > 1000 Then txtStealth.Text = 1000
    If val(txtStealth.Text) < 0 Then txtStealth.Text = 0
    nStealth = val(txtStealth.Text)
    
    If val(txtMaxDMG.Text) > 1000 Then txtMaxDMG.Text = 1000
    If val(txtMaxDMG.Text) < 0 Then txtMaxDMG.Text = 0
    nPlusMaxDamage = val(txtMaxDMG.Text)
    
    nMinStrBonus = Fix((val(txtStrength.Text) - 100) / 10)
    If Not bGreaterMUD Then nMinStrBonus = nMinStrBonus * 2
    If nMinStrBonus < 0 Then nMinStrBonus = 0
    
    If chkClassStealth.Value = 1 Then bClassStealth = True
    
    If nGlobalCharWeaponNumber(0) <> nWeaponNumber Or frmMain.chkGlobalFilter.Value = 0 Then
        For x = 0 To 19
            If tabItems.Fields("Abil-" & x) > 0 And tabItems.Fields("AbilVal-" & x) <> 0 Then
                nVal = tabItems.Fields("AbilVal-" & x)
                tStatSlot = GetAbilityStatSlot(tabItems.Fields("Abil-" & x), nVal)
                If Not tabItems.Fields("Number") = nWeaponNumber Then tabItems.Seek "=", nWeaponNumber
                If tStatSlot.nEquip > 0 Then
                    Select Case tStatSlot.nEquip
                        Case 11: nPlusMaxDamage = nPlusMaxDamage + nVal: lblMaxAdj.Caption = val(lblMaxAdj.Caption) + nVal
                        Case 14: nPlusBSmindmg = nPlusBSmindmg + nVal: lblBSMinAdj.Caption = val(lblBSMinAdj.Caption) + nVal
                        Case 15: nPlusBSmaxdmg = nPlusBSmaxdmg + nVal: lblBSMaxAdj.Caption = val(lblBSMaxAdj.Caption) + nVal
                        Case 19: nStealth = nStealth + nVal: lblStealthAdj.Caption = val(lblStealthAdj.Caption) + nVal
                    End Select
                End If
            End If
        Next x
        If val(lblMaxAdj.Caption) > 0 Then lblMaxAdj.Caption = "+" & lblMaxAdj.Caption
        If val(lblBSMinAdj.Caption) > 0 Then lblBSMinAdj.Caption = "+" & lblBSMinAdj.Caption
        If val(lblBSMaxAdj.Caption) > 0 Then lblBSMaxAdj.Caption = "+" & lblBSMaxAdj.Caption
        If val(lblStealthAdj.Caption) > 0 Then lblStealthAdj.Caption = "+" & lblStealthAdj.Caption
    End If
    
    If nGlobalCharWeaponNumber(0) <> nWeaponNumber And frmMain.chkGlobalFilter.Value = 1 Then
        If nGlobalCharWeaponMaxDmg(0) <> 0 Then
            nPlusMaxDamage = nPlusMaxDamage - nGlobalCharWeaponMaxDmg(0)
            lblMaxAdj.Caption = lblMaxAdj.Caption & IIf(nGlobalCharWeaponMaxDmg(0) < 0, "+", "-") & Abs(nGlobalCharWeaponMaxDmg(0))
        End If
        If nGlobalCharWeaponBSmindmg(0) <> 0 Then
            nPlusBSmindmg = nPlusBSmindmg - nGlobalCharWeaponBSmindmg(0)
            lblBSMinAdj.Caption = lblBSMinAdj.Caption & IIf(nGlobalCharWeaponBSmindmg(0) < 0, "+", "-") & Abs(nGlobalCharWeaponBSmindmg(0))
        End If
        If nGlobalCharWeaponBSmaxdmg(0) <> 0 Then
            nPlusBSmaxdmg = nPlusBSmaxdmg - nGlobalCharWeaponBSmaxdmg(0)
            lblBSMaxAdj.Caption = lblBSMaxAdj.Caption & IIf(nGlobalCharWeaponBSmaxdmg(0) < 0, "+", "-") & Abs(nGlobalCharWeaponBSmaxdmg(0))
        End If
        If nGlobalCharWeaponStealth(0) <> 0 Then
            nStealth = nStealth - nGlobalCharWeaponStealth(0)
            lblStealthAdj.Caption = lblStealthAdj.Caption & IIf(nGlobalCharWeaponStealth(0) < 0, "+", "-") & Abs(nGlobalCharWeaponStealth(0))
        End If
        
        If tabItems.Fields("WeaponType") = 1 Or tabItems.Fields("WeaponType") = 3 Then
            '+this weapon is two-handed...
            If nGlobalCharWeaponNumber(1) > 0 Then
                '+off-hand currently equipped. subtract those stats too...
                If nGlobalCharWeaponMaxDmg(1) <> 0 Then
                    nPlusMaxDamage = nPlusMaxDamage - nGlobalCharWeaponMaxDmg(1)
                    lblMaxAdj.Caption = lblMaxAdj.Caption & IIf(nGlobalCharWeaponMaxDmg(1) < 0, "+", "-") & Abs(nGlobalCharWeaponMaxDmg(1))
                End If
                If nGlobalCharWeaponBSmindmg(1) <> 0 Then
                    nPlusBSmindmg = nPlusBSmindmg - nGlobalCharWeaponBSmindmg(1)
                    lblBSMinAdj.Caption = lblBSMinAdj.Caption & IIf(nGlobalCharWeaponBSmindmg(1) < 0, "+", "-") & Abs(nGlobalCharWeaponBSmindmg(1))
                End If
                If nGlobalCharWeaponBSmaxdmg(1) <> 0 Then
                    nPlusBSmaxdmg = nPlusBSmaxdmg - nGlobalCharWeaponBSmaxdmg(1)
                    lblBSMaxAdj.Caption = lblBSMaxAdj.Caption & IIf(nGlobalCharWeaponBSmaxdmg(1) < 0, "+", "-") & Abs(nGlobalCharWeaponBSmaxdmg(1))
                End If
                If nGlobalCharWeaponStealth(1) <> 0 Then
                    nStealth = nStealth - nGlobalCharWeaponStealth(1)
                    lblStealthAdj.Caption = lblStealthAdj.Caption & IIf(nGlobalCharWeaponStealth(1) < 0, "+", "-") & Abs(nGlobalCharWeaponStealth(1))
                End If
            End If
        End If
    End If
    
    nMinDmg = tabItems.Fields("Min") + nMinStrBonus
    nMaxDmg = tabItems.Fields("Max") + nPlusMaxDamage
    If nMaxDmg < nMinDmg Then nMaxDmg = nMinDmg
    
    nMinDmg = CalcBSDamage(val(txtLevel.Text), nStealth, nMinDmg, nPlusBSmindmg, bClassStealth)
    nMaxDmg = CalcBSDamage(val(txtLevel.Text), nStealth, nMaxDmg, nPlusBSmaxdmg, bClassStealth)
    
    lblDMG.Caption = nMinDmg & " - " & nMaxDmg & " (AVG: " & Round((nMaxDmg + nMinDmg) / 2) & ")"
Else
    tabItems.MoveFirst
End If

Exit Sub

error:
Call HandleError("CalcBS")

End Sub

Private Sub Form_Resize()
'CheckPosition Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not bAppTerminating Then frmMain.SetFocus
Call WriteStealth
'Set objToolTip = Nothing
End Sub


Private Sub timButtonPress_Timer()
timButtonPress.Enabled = False
End Sub

Private Sub timCalc_Timer()
timCalc.Enabled = False
Call CalcBS
End Sub

Private Sub timWindowMove_Timer()
Call MonitorFormTimer(Me)
End Sub

Private Sub txtBSMaxDMG_Change()
timCalc.Enabled = False: timCalc.Enabled = True: 'Call CalcBS
End Sub

Private Sub txtBSMinDMG_Change()
timCalc.Enabled = False: timCalc.Enabled = True: 'Call CalcBS
End Sub

Private Sub txtLevel_Change()
timCalc.Enabled = False: timCalc.Enabled = True: 'Call CalcBS
End Sub

Private Sub txtMaxDMG_Change()
timCalc.Enabled = False: timCalc.Enabled = True: 'Call CalcBS
End Sub

Private Sub txtStealth_Change()
timCalc.Enabled = False: timCalc.Enabled = True: 'Call CalcBS
End Sub

Private Sub txtStealth_GotFocus()
Call SelectAll(txtStealth)
End Sub

Private Sub txtStealth_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtBSMinDMG_GotFocus()
Call SelectAll(txtBSMinDMG)
End Sub

Private Sub txtBSMinDMG_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtLevel_GotFocus()
Call SelectAll(txtLevel)
End Sub

Private Sub txtLevel_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub


Private Sub txtBSMaxDMG_GotFocus()
Call SelectAll(txtBSMaxDMG)
End Sub

Private Sub txtBSMaxDMG_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub


Private Sub txtMaxDMG_GotFocus()
Call SelectAll(txtMaxDMG)
End Sub

Private Sub txtMaxDMG_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub


Private Sub txtStrength_Change()
timCalc.Enabled = False: timCalc.Enabled = True: 'Call CalcBS
End Sub

Private Sub txtStrength_GotFocus()
Call SelectAll(txtStrength)

End Sub
