VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10110
   Begin VB.CommandButton cmdRecreateINI 
      Caption         =   "Recreate settings.ini"
      Height          =   375
      Left            =   3660
      TabIndex        =   3
      Top             =   5820
      Width           =   2115
   End
   Begin VB.Frame Frame4 
      Caption         =   "Settings"
      Height          =   5535
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   9915
      Begin VB.TextBox txtMonsterSimRounds 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   5640
         MaxLength       =   5
         TabIndex        =   31
         ToolTipText     =   "Min = 100, Max = 10000, Default = 500 (more rounds = more accurate but slow calc time)"
         Top             =   1140
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Caption         =   "Monster Exp/Hour Calculation Models"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3675
         Left            =   5640
         TabIndex        =   25
         Top             =   1680
         Width           =   3855
         Begin VB.OptionButton optEPH_Model 
            Caption         =   "No Recovery"
            Height          =   255
            Index           =   1
            Left            =   2160
            TabIndex        =   52
            ToolTipText     =   "Will average both models for movement and kill time, but not account for recovery time."
            Top             =   300
            Width           =   1575
         End
         Begin VB.CommandButton cmdModelQ 
            Caption         =   "Reset Values"
            Height          =   315
            Index           =   1
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   2700
            Width           =   1395
         End
         Begin VB.CommandButton cmdModelQ 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Models?"
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
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   3180
            Width           =   1395
         End
         Begin VB.CommandButton cmdCEPHB_Q 
            Caption         =   "XP"
            Height          =   315
            Index           =   3
            Left            =   2820
            TabIndex        =   49
            Top             =   2280
            Width           =   735
         End
         Begin VB.CommandButton cmdCEPHB_Q 
            Caption         =   "Move"
            Height          =   315
            Index           =   2
            Left            =   2820
            TabIndex        =   48
            Top             =   1860
            Width           =   735
         End
         Begin VB.CommandButton cmdCEPHB_Q 
            Caption         =   "Mana"
            Height          =   315
            Index           =   1
            Left            =   2820
            TabIndex        =   47
            Top             =   1440
            Width           =   735
         End
         Begin VB.OptionButton optEPH_Model 
            Caption         =   "Model [B]:"
            Height          =   255
            Index           =   3
            Left            =   2160
            TabIndex        =   46
            Top             =   660
            Width           =   1335
         End
         Begin VB.CommandButton cmdCEPHB_Q 
            Caption         =   "DMG"
            Height          =   315
            Index           =   0
            Left            =   2820
            TabIndex        =   45
            Top             =   1020
            Width           =   735
         End
         Begin VB.TextBox txtCEPHB_DMG 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   2160
            MaxLength       =   4
            TabIndex        =   44
            ToolTipText     =   "Min = 0.01, Max = 2.99, Default = 0.9"
            Top             =   1020
            Width           =   615
         End
         Begin VB.TextBox txtCEPHB_XP 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   2160
            MaxLength       =   4
            TabIndex        =   43
            ToolTipText     =   "Min = 0.01, Max = 2.99, Default = 1"
            Top             =   2280
            Width           =   615
         End
         Begin VB.TextBox txtCEPHB_Move 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   2160
            MaxLength       =   4
            TabIndex        =   42
            ToolTipText     =   "Min = 0.01, Max = 2.99, Default = 0.9"
            Top             =   1860
            Width           =   615
         End
         Begin VB.TextBox txtCEPHB_Mana 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   2160
            MaxLength       =   4
            TabIndex        =   41
            ToolTipText     =   "Min = 0.01, Max = 2.99, Default = 0.95"
            Top             =   1440
            Width           =   615
         End
         Begin VB.OptionButton optEPH_Model 
            Caption         =   "Model [A]:"
            Height          =   255
            Index           =   2
            Left            =   300
            TabIndex        =   40
            Top             =   660
            Width           =   1335
         End
         Begin VB.CommandButton cmdCEPHA_Q 
            Caption         =   "Cluster"
            Height          =   315
            Index           =   4
            Left            =   960
            TabIndex        =   38
            Top             =   2700
            Width           =   735
         End
         Begin VB.CommandButton cmdCEPHA_Q 
            Caption         =   "Reco"
            Height          =   315
            Index           =   3
            Left            =   960
            TabIndex        =   37
            Top             =   2280
            Width           =   735
         End
         Begin VB.CommandButton cmdCEPHA_Q 
            Caption         =   "Move"
            Height          =   315
            Index           =   2
            Left            =   960
            TabIndex        =   36
            Top             =   1860
            Width           =   735
         End
         Begin VB.CommandButton cmdCEPHA_Q 
            Caption         =   "Mana"
            Height          =   315
            Index           =   1
            Left            =   960
            TabIndex        =   35
            Top             =   1440
            Width           =   735
         End
         Begin VB.CommandButton cmdCEPHA_Q 
            Caption         =   "DMG"
            Height          =   315
            Index           =   0
            Left            =   960
            TabIndex        =   34
            Top             =   1020
            Width           =   735
         End
         Begin VB.TextBox txtCEPHA_ClusterMx 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   300
            MaxLength       =   3
            TabIndex        =   33
            ToolTipText     =   "Min = 1, Max = 255, Default = 10 (lower = more areas considered clusters)"
            Top             =   2700
            Width           =   615
         End
         Begin VB.CheckBox chkExpHrCalcByCharacter 
            Caption         =   "Save these settings per character"
            Height          =   375
            Left            =   300
            TabIndex        =   30
            Top             =   3180
            Width           =   1875
         End
         Begin VB.TextBox txtCEPHA_Mana 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   300
            MaxLength       =   4
            TabIndex        =   29
            ToolTipText     =   "Min = 0.01, Max = 2.99, Default = 1"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtCEPHA_Move 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   300
            MaxLength       =   4
            TabIndex        =   28
            ToolTipText     =   "Min = 0.01, Max = 2.99, Default = 1"
            Top             =   1860
            Width           =   615
         End
         Begin VB.TextBox txtCEPHA_MoveReco 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   300
            MaxLength       =   4
            TabIndex        =   27
            ToolTipText     =   "Min = 0.01, Max = 2.99, Default = 0.85"
            Top             =   2280
            Width           =   615
         End
         Begin VB.TextBox txtCEPHA_DMG 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   300
            MaxLength       =   4
            TabIndex        =   26
            ToolTipText     =   "Min = 0.01, Max = 2.99, Default = 1"
            Top             =   1020
            Width           =   615
         End
         Begin VB.OptionButton optEPH_Model 
            Caption         =   "Average Both"
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
            Left            =   300
            TabIndex        =   39
            ToolTipText     =   "Will average both models."
            Top             =   300
            Value           =   -1  'True
            Width           =   1635
         End
         Begin VB.Line Line2 
            X1              =   1920
            X2              =   1920
            Y1              =   1080
            Y2              =   2580
         End
      End
      Begin VB.CheckBox chkDontLookupMonsterRegen 
         Caption         =   "Don't lookup monster in detail (minor performance boost)"
         Height          =   255
         Left            =   180
         TabIndex        =   20
         Top             =   3780
         Width           =   4635
      End
      Begin VB.CheckBox chkSwapWindowTitle 
         Caption         =   "Put the character / filename at the start of MME's window title"
         Height          =   255
         Left            =   180
         TabIndex        =   17
         Top             =   2880
         Width           =   4815
      End
      Begin VB.CheckBox chkShowCharacterName 
         Caption         =   "Show character name in window title instead of filename"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   2580
         Width           =   4995
      End
      Begin VB.CheckBox chkAutoCalcMonDamage 
         Caption         =   "Auto-Calculate monster damage vs character stats in real-time"
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   1380
         Width           =   5175
      End
      Begin VB.CheckBox chkRemoveListEquip 
         Caption         =   "Remove item/spell from saved lists when equipping or learning"
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   1080
         Width           =   4995
      End
      Begin VB.CheckBox chkWindowSnap 
         Caption         =   "Disable Window/Display Snap (could cause window to get lost on disconnected or reconfigured monitors)"
         Height          =   435
         Left            =   180
         TabIndex        =   24
         Top             =   4980
         Width           =   4875
      End
      Begin VB.CheckBox chkNavSpan 
         Caption         =   "Don't span navigation buttons on resize"
         Height          =   255
         Left            =   180
         TabIndex        =   21
         Top             =   4080
         Width           =   3435
      End
      Begin VB.CheckBox chkUseWrist 
         Caption         =   "Use Second Wrist Slot"
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   3180
         Width           =   2175
      End
      Begin VB.CheckBox chkHideRecordNumbers 
         Caption         =   "Hide record numbers when referencing names"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   3480
         Width           =   3735
      End
      Begin VB.CheckBox chkWindowsOnTop 
         Caption         =   "Don't make Results window stay on top of main window"
         Height          =   255
         Left            =   180
         TabIndex        =   23
         Top             =   4680
         Width           =   4695
      End
      Begin VB.CheckBox chkSwapMapButtons 
         Caption         =   "Swap left/right mouse buttons for maps"
         Height          =   255
         Left            =   180
         TabIndex        =   22
         Top             =   4380
         Width           =   3435
      End
      Begin VB.CheckBox chkAutoSaveChar 
         Caption         =   "Always Auto-Save Loaded Character"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   1680
         Width           =   3135
      End
      Begin VB.CheckBox chkAutoLoadChar 
         Caption         =   "Auto-Load last character associated with database"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   1980
         Width           =   4155
      End
      Begin VB.CheckBox chkFilterAll 
         Caption         =   """ Filter All "" on program load"
         Enabled         =   0   'False
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   2280
         Width           =   2415
      End
      Begin VB.CheckBox chkInGame 
         Caption         =   "Only load items, monsters, and shops that are in the game (requires reload after changing)"
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   600
         Width           =   7035
      End
      Begin VB.CommandButton cmdNone 
         Caption         =   "None"
         Height          =   435
         Left            =   8580
         TabIndex        =   9
         Top             =   300
         Width           =   1095
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "All"
         Height          =   435
         Left            =   7380
         TabIndex        =   8
         Top             =   300
         Width           =   1035
      End
      Begin VB.CheckBox chkLoadShops 
         Caption         =   "Load Shops"
         Height          =   255
         Left            =   4140
         TabIndex        =   7
         Top             =   300
         Width           =   1215
      End
      Begin VB.CheckBox chkLoadMonsters 
         Caption         =   "Load Monsters"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   300
         Width           =   1395
      End
      Begin VB.CheckBox chkLoadSpells 
         Caption         =   "Load Spells"
         Height          =   255
         Left            =   1380
         TabIndex        =   5
         Top             =   300
         Width           =   1215
      End
      Begin VB.CheckBox chkLoadItems 
         Caption         =   "Load Items"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Rounds to sim when calculating mon dmg (more = more accurate but slower)"
         Height          =   435
         Left            =   6360
         TabIndex        =   32
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   180
         X2              =   9660
         Y1              =   960
         Y2              =   960
      End
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
      Height          =   375
      Left            =   8700
      TabIndex        =   1
      Top             =   5820
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
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
      Left            =   120
      TabIndex        =   0
      Top             =   5820
      Width           =   1575
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit
Dim tWindowSize As WindowSizeProperties
Dim bPerCharOnLoad As Boolean

Private Sub chkAutoLoadChar_Click()
DoEvents
If chkAutoLoadChar.Value = 1 Then
    chkFilterAll.Enabled = True
Else
    chkFilterAll.Enabled = False
End If
End Sub

Private Sub chkExpHrCalcByCharacter_Click()
On Error GoTo error:
Dim x As Integer

If frmMain.bCharLoaded And sSessionLastCharFile <> "" Then
    If bPerCharOnLoad And chkExpHrCalcByCharacter.Value = 0 Then
        x = MsgBox("Reload values from from the global settings?", vbYesNo + vbQuestion + vbDefaultButton1)
        If x = vbYes Then
            Call frmMain.LoadExpPerHourKnobs(False, True)
            Call PopulateExpFields
            Call frmMain.LoadExpPerHourKnobs(True)
            bPerCharOnLoad = False
        End If
    ElseIf Not bPerCharOnLoad And chkExpHrCalcByCharacter.Value = 1 Then
        x = MsgBox("Reload values from current character settings?", vbYesNo + vbQuestion + vbDefaultButton1)
        If x = vbYes Then
            Call frmMain.LoadExpPerHourKnobs(True)
            Call PopulateExpFields
            Call frmMain.LoadExpPerHourKnobs(False, True)
            bPerCharOnLoad = True
        End If
    End If
End If

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("chkExpHrCalcByCharacter_Click")
Resume out:
End Sub

Private Sub cmdAll_Click()

chkLoadItems.Value = 1
chkLoadSpells.Value = 1
'chkLoadClasses.Value = 1
'chkLoadRaces.Value = 1
chkLoadMonsters.Value = 1
chkLoadShops.Value = 1

End Sub


Private Sub cmdCEPHA_Q_Click(Index As Integer)
Select Case Index
    Case 0: MsgBox "Scales incoming damage. Larger multiplier makes thing hurt more, smaller numbers hurt less.", vbInformation
    Case 1: MsgBox "Scales mana recovery. Directly effects the time spent recovering. Smaller multiplier = less time spent.", vbInformation
    Case 2: MsgBox "Scales movement/lairs. Larger multiplier scales up movement and vice versa.", vbInformation
    Case 3: MsgBox "Scales recovery while moving. A larger multiplier will allow more recovery to be crediting while moving.", vbInformation
    Case 4: MsgBox "This model has trouble detecting areas like gnoll encampment where the lairs are in a small cluster of a thousand rooms.  " _
        & "If the calculated AvgWalk movement is <= 2 AND the total lairs in the area multiplied by this number is still less than the " _
        & "total rooms in the area, then consider it a cluser and greatly reduce movement penalties. (gnolls: 1.3 avg walk, 13lairs*10mult = 130 which is < the 1200+ rooms in the area)", vbInformation
End Select
End Sub

Private Sub cmdCEPHB_Q_Click(Index As Integer)
Select Case Index
    Case 0: MsgBox "Scales incoming damage. Larger multiplier makes thing hurt more (and then recover more). Smaller numbers hurt less.", vbInformation
    Case 1: MsgBox "Scales mana recovery. Smaller multiplier = less time spent.", vbInformation
    Case 2: MsgBox "Scales movement/lairs. Larger multiplier scales up movement and vice versa.", vbInformation
    Case 3: MsgBox "Directly multiplies the exp/hr result before being returned.", vbInformation
End Select
End Sub

Private Sub cmdModelQ_Click(Index As Integer)
If Index = 0 Then
    MsgBox "I used ChatGPT to help build both of these models.  Model A came first.  Call it model Alpha.  " _
        & "Model B, call it Beta, was built second using gained knowledge and additional information and should be better, but there are instances where model A is more accurate.  " _
        & "Model B also provides more straightforward tuning knobs.  Choosing one model will save CPU cycles." _
        & vbCrLf & vbCrLf _
        & "Choosing 'No Recovery' will average both models for movement and kill time, based on chosen attack (damage out), but provide just the standard filter for incoming damage.  e.g. No recovery time will be accounted for.", vbInformation + vbOKOnly
ElseIf Index = 1 Then
    'reset
    txtCEPHA_DMG.Text = 1
    txtCEPHA_Mana.Text = 1
    txtCEPHA_Move.Text = 1
    txtCEPHA_MoveReco.Text = 0.85
    txtCEPHA_ClusterMx.Text = 10
    txtCEPHB_DMG.Text = 0.9
    txtCEPHB_Mana.Text = 0.95
    txtCEPHB_Move.Text = 0.9
    txtCEPHB_XP.Text = 1
End If
End Sub

Private Sub cmdNone_Click()

chkLoadItems.Value = 0
chkLoadSpells.Value = 0
'chkLoadClasses.Value = 0
'chkLoadRaces.Value = 0
chkLoadMonsters.Value = 0
chkLoadShops.Value = 0

End Sub

Private Sub Form_Load()
On Error GoTo error:
Dim sSectionName As String

'stop windows from resizing fixed-size windows when changing dpi
If bDPIAwareMode Then Call SubclassFormMinMaxSize(Me, tWindowSize, True)

sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ")

If val(ReadINI("Settings", "LoadItems", , 1)) > 0 Then chkLoadItems.Value = 1
If val(ReadINI("Settings", "LoadSpells", , 1)) > 0 Then chkLoadSpells.Value = 1
If val(ReadINI("Settings", "LoadMonsters", , 1)) > 0 Then chkLoadMonsters.Value = 1
If val(ReadINI("Settings", "LoadShops", , 1)) > 0 Then chkLoadShops.Value = 1
If val(ReadINI("Settings", "OnlyInGame", , 1)) > 0 Then chkInGame.Value = 1
If val(ReadINI("Settings", "FilterAll")) > 0 Then chkFilterAll.Value = 1
If val(ReadINI("Settings", "AutoLoadLastChar")) > 0 Then chkAutoLoadChar.Value = 1
If val(ReadINI("Settings", "AutoSaveLastChar")) > 0 Then chkAutoSaveChar.Value = 1
If val(ReadINI("Settings", "SwapMapButtons")) > 0 Then chkSwapMapButtons.Value = 1
If val(ReadINI("Settings", "NoAlwaysOnTop")) > 0 Then chkWindowsOnTop.Value = 1
If val(ReadINI("Settings", "HideRecordNumbers")) > 0 Then chkHideRecordNumbers.Value = 1
If val(ReadINI("Settings", "Use2ndWrist", , 1)) > 0 Then chkUseWrist.Value = 1
If val(ReadINI("Settings", "NameInTitle")) > 0 Then chkShowCharacterName.Value = 1
If val(ReadINI("Settings", "DontSpanNavButtons")) > 0 Then chkNavSpan.Value = 1
If val(ReadINI("Settings", "DisableWindowSnap")) > 0 Then chkWindowSnap.Value = 1
If val(ReadINI("Settings", "RemoveListEquip")) > 0 Then chkRemoveListEquip.Value = 1
If val(ReadINI("Settings", "AutoCalcMonDamage", , "1")) > 0 Then chkAutoCalcMonDamage.Value = 1
If val(ReadINI("Settings", "SwapWindowTitle")) > 0 Then chkSwapWindowTitle.Value = 1
If val(ReadINI("Settings", "DontLookupMonsterRegen")) > 0 Then chkDontLookupMonsterRegen.Value = 1
If val(ReadINI("Settings", "ExpPerHourKnobsByCharacter")) > 0 Then
    bPerCharOnLoad = True
    chkExpHrCalcByCharacter.Value = 1
End If

txtMonsterSimRounds.Text = nGlobalMonsterSimRounds

Call PopulateExpFields
Call chkAutoLoadChar_Click

If frmMain.WindowState = vbMinimized Then
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
Else
    Me.Left = frmMain.Left + ((frmMain.Width - Me.Width) / 2)
    Me.Top = frmMain.Top + ((frmMain.Height - Me.Height) / 2)
End If

out:
Exit Sub
error:
Call HandleError("Form_Load")
Resume out:
End Sub

Private Sub PopulateExpFields()
On Error GoTo error:

txtCEPHA_DMG.Text = nGlobal_cephA_DMG
txtCEPHA_Mana.Text = nGlobal_cephA_Mana
txtCEPHA_MoveReco.Text = nGlobal_cephA_MoveRecover
txtCEPHA_ClusterMx.Text = nGlobal_cephA_ClusterMx
txtCEPHA_Move.Text = nGlobal_cephA_Move

txtCEPHB_DMG.Text = nGlobal_cephB_DMG
txtCEPHB_Mana.Text = nGlobal_cephB_Mana
txtCEPHB_Move.Text = nGlobal_cephB_Move
txtCEPHB_XP.Text = nGlobal_cephB_XP

Select Case eGlobalExpHrModel
    Case 0, 1: 'default, average
        optEPH_Model(0).Value = True
    Case 2: 'modelA
        optEPH_Model(2).Value = True
    Case 3: 'modelB
        optEPH_Model(3).Value = True
    Case 99:
        optEPH_Model(1).Value = True
End Select
Call optEPH_Model_Click(0)

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("PopulateExpFields")
Resume out:
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdRecreateINI_Click()
Dim nYesNo As Integer

nYesNo = MsgBox("Are you sure you want to delete and recreate your setting.ini (resetting all values to defaults)?", vbYesNo + vbDefaultButton2 + vbQuestion, "Reset Settings?")
If nYesNo <> vbYes Then Exit Sub

frmMain.bReloadWithNewSettings = True
Unload Me

End Sub

Private Sub cmdSave_Click()
Dim sSectionName As String, x As Integer, sFile As String

On Error GoTo error:

sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ")

nGlobalMonsterSimRounds = val(txtMonsterSimRounds.Text)
If nGlobalMonsterSimRounds < 100 Then nGlobalMonsterSimRounds = 100
If nGlobalMonsterSimRounds > 10000 Then nGlobalMonsterSimRounds = 10000

nGlobal_cephA_DMG = Round(val(txtCEPHA_DMG.Text), 2)
If nGlobal_cephA_DMG < 0.01 Then nGlobal_cephA_DMG = 0.01
If nGlobal_cephA_DMG > 2.99 Then nGlobal_cephA_DMG = 2.99

nGlobal_cephA_Mana = Round(val(txtCEPHA_Mana.Text), 2)
If nGlobal_cephA_Mana < 0.01 Then nGlobal_cephA_Mana = 0.01
If nGlobal_cephA_Mana > 2.99 Then nGlobal_cephA_Mana = 2.99

nGlobal_cephA_MoveRecover = Round(val(txtCEPHA_MoveReco.Text), 2)
If nGlobal_cephA_MoveRecover < 0.01 Then nGlobal_cephA_MoveRecover = 0.01
If nGlobal_cephA_MoveRecover > 2.99 Then nGlobal_cephA_MoveRecover = 2.99

nGlobal_cephA_Move = Round(val(txtCEPHA_Move.Text), 2)
If nGlobal_cephA_Move < 0.01 Then nGlobal_cephA_Move = 0.01
If nGlobal_cephA_Move > 2.99 Then nGlobal_cephA_Move = 2.99

nGlobal_cephA_ClusterMx = Round(val(txtCEPHA_ClusterMx.Text))
If nGlobal_cephA_ClusterMx < 1 Then nGlobal_cephA_ClusterMx = 1
If nGlobal_cephA_ClusterMx > 255 Then nGlobal_cephA_ClusterMx = 255

'---

nGlobal_cephB_DMG = Round(val(txtCEPHB_DMG.Text), 2)
If nGlobal_cephB_DMG < 0.01 Then nGlobal_cephB_DMG = 0.01
If nGlobal_cephB_DMG > 2.99 Then nGlobal_cephB_DMG = 2.99

nGlobal_cephB_Mana = Round(val(txtCEPHB_Mana.Text), 2)
If nGlobal_cephB_Mana < 0.01 Then nGlobal_cephB_Mana = 0.01
If nGlobal_cephB_Mana > 2.99 Then nGlobal_cephB_Mana = 2.99

nGlobal_cephB_Move = Round(val(txtCEPHB_Move.Text), 2)
If nGlobal_cephB_Move < 0.01 Then nGlobal_cephB_Move = 0.01
If nGlobal_cephB_Move > 2.99 Then nGlobal_cephB_Move = 2.99

nGlobal_cephB_XP = Round(val(txtCEPHB_XP.Text), 2)
If nGlobal_cephB_XP < 0.01 Then nGlobal_cephB_XP = 0.01
If nGlobal_cephB_XP > 2.99 Then nGlobal_cephB_XP = 2.99

'---

If optEPH_Model(0).Value = True Then
    eGlobalExpHrModel = average
ElseIf optEPH_Model(1).Value = True Then
    eGlobalExpHrModel = basic_dmg
ElseIf optEPH_Model(2).Value = True Then
    eGlobalExpHrModel = modelA
ElseIf optEPH_Model(3).Value = True Then
    eGlobalExpHrModel = modelB
End If

Call WriteINI("Settings", "LoadItems", chkLoadItems.Value)
Call WriteINI("Settings", "LoadSpells", chkLoadSpells.Value)
Call WriteINI("Settings", "LoadMonsters", chkLoadMonsters.Value)
Call WriteINI("Settings", "LoadShops", chkLoadShops.Value)
Call WriteINI("Settings", "OnlyInGame", chkInGame.Value)
Call WriteINI("Settings", "FilterAll", chkFilterAll.Value)
Call WriteINI("Settings", "AutoLoadLastChar", chkAutoLoadChar.Value)
Call WriteINI("Settings", "AutoSaveLastChar", chkAutoSaveChar.Value)
Call WriteINI("Settings", "SwapMapButtons", chkSwapMapButtons.Value)
Call WriteINI("Settings", "NoAlwaysOnTop", chkWindowsOnTop.Value)
Call WriteINI("Settings", "HideRecordNumbers", chkHideRecordNumbers.Value)
Call WriteINI("Settings", "Use2ndWrist", chkUseWrist.Value)
Call WriteINI("Settings", "NameInTitle", chkShowCharacterName.Value)
Call WriteINI("Settings", "DontSpanNavButtons", chkNavSpan.Value)
Call WriteINI("Settings", "DisableWindowSnap", chkWindowSnap.Value)
Call WriteINI("Settings", "RemoveListEquip", chkRemoveListEquip.Value)
Call WriteINI("Settings", "AutoCalcMonDamage", chkAutoCalcMonDamage.Value)
Call WriteINI("Settings", "SwapWindowTitle", chkSwapWindowTitle.Value)
Call WriteINI("Settings", "DontLookupMonsterRegen", chkDontLookupMonsterRegen.Value)
Call WriteINI("Settings", "MonsterSimRounds", nGlobalMonsterSimRounds)

Call WriteINI("Settings", "ExpPerHourKnobsByCharacter", chkExpHrCalcByCharacter.Value)

If chkExpHrCalcByCharacter.Value = 1 And frmMain.bCharLoaded = True And sSessionLastCharFile <> "" Then sFile = sSessionLastCharFile
Call WriteINI("Settings", "ExpHrModel", eGlobalExpHrModel, sFile)

Call WriteINI("Settings", "cephA_DMG", nGlobal_cephA_DMG, sFile)
Call WriteINI("Settings", "cephA_Mana", nGlobal_cephA_Mana, sFile)
Call WriteINI("Settings", "cephA_Move", nGlobal_cephA_Move, sFile)
Call WriteINI("Settings", "cephA_MoveRecover", nGlobal_cephA_MoveRecover, sFile)
Call WriteINI("Settings", "cephA_ClusterMx", nGlobal_cephA_ClusterMx, sFile)

Call WriteINI("Settings", "cephB_DMG", nGlobal_cephB_DMG, sFile)
Call WriteINI("Settings", "cephB_Mana", nGlobal_cephB_Mana, sFile)
Call WriteINI("Settings", "cephB_Move", nGlobal_cephB_Move, sFile)
Call WriteINI("Settings", "cephB_XP", nGlobal_cephB_XP, sFile)

If chkDontLookupMonsterRegen.Value = 1 Then
    frmMain.bDontLookupMonRegen = True
Else
    frmMain.bDontLookupMonRegen = False
End If

'Call WriteINI("Settings", "FilterAllChar", chkFilterAllChar.Value)

If chkAutoCalcMonDamage.Value = 1 Then
    frmMain.bAutoCalcMonDamage = True
Else
    frmMain.bAutoCalcMonDamage = False
End If

If chkRemoveListEquip.Value = 1 Then
    frmMain.bRemoveListEquip = True
Else
    frmMain.bRemoveListEquip = False
End If

If chkAutoSaveChar.Value = 1 Then
    frmMain.bAutoSave = True
Else
    frmMain.bAutoSave = False
End If

If chkNavSpan.Value = 1 Then
    frmMain.bDontSpanNav = True
Else
    frmMain.bDontSpanNav = False
End If

If chkNavSpan.Value = 1 Then
    frmMain.framButtons.Width = 13335
    frmMain.fraDatVer.Width = 6915
    frmMain.lblDatVer.Width = 6705

    For x = 0 To 10
        Select Case x
            Case 0:
                frmMain.cmdNav(x).Width = 1335
            Case 1:
                frmMain.cmdNav(x).Left = frmMain.cmdNav(x - 1).Left + frmMain.cmdNav(x - 1).Width - 15
                frmMain.cmdNav(x).Width = 1095
            Case 2:
                frmMain.cmdNav(x).Left = frmMain.cmdNav(x - 1).Left + frmMain.cmdNav(x - 1).Width - 15
                frmMain.cmdNav(x).Width = 1035
            Case 3:
                frmMain.cmdNav(x).Left = frmMain.cmdNav(x - 1).Left + frmMain.cmdNav(x - 1).Width - 15
                frmMain.cmdNav(x).Width = 1215
            Case 4:
                frmMain.cmdNav(x).Left = frmMain.cmdNav(x - 1).Left + frmMain.cmdNav(x - 1).Width - 15
                frmMain.cmdNav(x).Width = 1455
            Case 5:
                frmMain.cmdNav(x).Left = frmMain.cmdNav(x - 1).Left + frmMain.cmdNav(x - 1).Width - 15
                frmMain.cmdNav(x).Width = 1335
            Case 6:
                frmMain.cmdNav(x).Left = frmMain.cmdNav(x - 1).Left + frmMain.cmdNav(x - 1).Width - 15
                frmMain.cmdNav(x).Width = 1455
            Case 7:
                frmMain.cmdNav(x).Left = frmMain.cmdNav(x - 1).Left + frmMain.cmdNav(x - 1).Width - 15
                frmMain.cmdNav(x).Width = 1095
            Case 8:
                frmMain.cmdNav(x).Left = frmMain.cmdNav(x - 1).Left + frmMain.cmdNav(x - 1).Width - 15
                frmMain.cmdNav(x).Width = 1215
            Case 9:
                frmMain.cmdNav(x).Left = frmMain.cmdNav(x - 1).Left + frmMain.cmdNav(x - 1).Width - 15
                frmMain.cmdNav(x).Width = 1035
            Case 10:
                frmMain.cmdNav(x).Left = frmMain.cmdNav(x - 1).Left + frmMain.cmdNav(x - 1).Width - 15
                frmMain.cmdNav(x).Width = 1095
        End Select
    Next x
End If

If chkShowCharacterName.Value = 1 Then
    frmMain.bNameInTitle = True
Else
    frmMain.bNameInTitle = False
End If

If chkUseWrist.Value = 1 Then
    frmMain.chkEquipHold(7).Enabled = True
    frmMain.cmdEquipGoto(7).Enabled = True
    frmMain.cmbEquip(7).Enabled = True
    frmMain.bInvenUse2ndWrist = True
Else
    frmMain.chkEquipHold(7).Enabled = False
    frmMain.cmdEquipGoto(7).Enabled = False
    frmMain.cmbEquip(7).Enabled = False
    frmMain.bInvenUse2ndWrist = False
    If frmMain.cmbEquip(7).ListIndex > 0 Then frmMain.cmbEquip(7).ListIndex = 0
End If

If chkHideRecordNumbers.Value = 1 Then
    bHideRecordNumbers = True
Else
    bHideRecordNumbers = False
End If

If chkWindowsOnTop.Value = 1 Then
    frmMain.bNoAlwaysOnTop = True
Else
    frmMain.bNoAlwaysOnTop = False
End If

If chkSwapMapButtons.Value = 1 Then
    frmMain.bMapSwapButtons = True
    If FormIsLoaded("frmMap") Then frmMap.bMapSwapButtons = True
Else
    frmMain.bMapSwapButtons = False
    If FormIsLoaded("frmMap") Then frmMap.bMapSwapButtons = False
End If

If chkWindowSnap.Value = 1 Then
    frmMain.bDisableWindowSnap = True
Else
    frmMain.bDisableWindowSnap = False
End If

Call frmMain.Form_Resize_Event
Call frmMain.RefreshMonsterCombatGUI

out:
On Error Resume Next
Unload Me
Exit Sub
error:
Call HandleError("cmdSave_Click")
Resume out:
End Sub

Private Sub optEPH_Model_Click(Index As Integer)
Dim x As Integer
For x = 0 To optEPH_Model.Count - 1
    If optEPH_Model(x).Value = True Then
        optEPH_Model(x).FontBold = True
    Else
        optEPH_Model(x).FontBold = False
    End If
Next x
End Sub

Private Sub txtCEPHA_ClusterMx_GotFocus()
Call SelectAll(txtCEPHA_ClusterMx)
End Sub

Private Sub txtCEPHA_ClusterMx_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtCEPHA_DMG_GotFocus()
Call SelectAll(txtCEPHA_DMG)
End Sub

Private Sub txtCEPHA_DMG_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii, True)
End Sub

Private Sub txtCEPHA_Mana_GotFocus()
Call SelectAll(txtCEPHA_Mana)
End Sub

Private Sub txtCEPHA_Mana_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii, True)
End Sub

Private Sub txtCEPHA_MoveReco_GotFocus()
Call SelectAll(txtCEPHA_MoveReco)
End Sub

Private Sub txtCEPHA_MoveReco_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii, True)
End Sub

Private Sub txtCEPHB_DMG_GotFocus()
Call SelectAll(txtCEPHB_DMG)
End Sub

Private Sub txtCEPHB_DMG_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii, True)
End Sub

Private Sub txtCEPHB_Mana_GotFocus()
Call SelectAll(txtCEPHB_Mana)
End Sub

Private Sub txtCEPHB_Mana_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii, True)
End Sub

Private Sub txtCEPHB_Move_GotFocus()
Call SelectAll(txtCEPHB_Move)
End Sub

Private Sub txtCEPHB_Move_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii, True)
End Sub

Private Sub txtCEPHB_XP_GotFocus()
Call SelectAll(txtCEPHB_XP)
End Sub

Private Sub txtCEPHB_XP_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii, True)
End Sub

Private Sub txtMonsterSimRounds_GotFocus()
Call SelectAll(txtMonsterSimRounds)
End Sub

Private Sub txtMonsterSimRounds_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtCEPHA_Move_GotFocus()
Call SelectAll(txtCEPHA_Move)
End Sub

Private Sub txtCEPHA_Move_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii, True)
End Sub
