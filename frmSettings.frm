VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   10110
   Begin VB.CommandButton cmdRecreateINI 
      Caption         =   "Recreate settings.ini"
      Height          =   375
      Left            =   4020
      TabIndex        =   3
      Top             =   6900
      Width           =   2115
   End
   Begin VB.Frame Frame4 
      Caption         =   "Settings"
      Height          =   6675
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   9915
      Begin VB.CheckBox chkSwapMonDetailOutput 
         Caption         =   "In lair mode, print damage out and lair/scripting information first"
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   3840
         Width           =   5115
      End
      Begin VB.TextBox txtMonsterSimRounds 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   5700
         MaxLength       =   5
         TabIndex        =   26
         ToolTipText     =   "Min = 100, Max = 10000, Default = 500 (more rounds = more accurate but slow calc time)"
         Top             =   1440
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Caption         =   "Monster Exp/Hour Calculation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   5760
         TabIndex        =   38
         Top             =   2040
         Width           =   3375
         Begin VB.CheckBox chkEPH_Model 
            Caption         =   "Model D (New)"
            Height          =   255
            Index           =   3
            Left            =   1680
            TabIndex        =   49
            Top             =   660
            Width           =   1455
         End
         Begin VB.CheckBox chkEPH_Model 
            Caption         =   "Show in Detail"
            Height          =   195
            Index           =   99
            Left            =   1680
            TabIndex        =   48
            Top             =   1020
            Width           =   1575
         End
         Begin VB.CheckBox chkEPH_Model 
            Caption         =   "No Recovery"
            Height          =   255
            Index           =   98
            Left            =   300
            TabIndex        =   47
            Top             =   1020
            Width           =   1335
         End
         Begin VB.CheckBox chkEPH_Model 
            Caption         =   "Model C"
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   46
            Top             =   300
            Width           =   1035
         End
         Begin VB.CheckBox chkEPH_Model 
            Caption         =   "Model B"
            Height          =   255
            Index           =   1
            Left            =   300
            TabIndex        =   45
            Top             =   660
            Width           =   1035
         End
         Begin VB.CheckBox chkEPH_Model 
            Caption         =   "Model A"
            Height          =   255
            Index           =   0
            Left            =   300
            TabIndex        =   44
            Top             =   300
            Width           =   1035
         End
         Begin VB.CommandButton cmdModelQ 
            Caption         =   "Reset Values"
            Height          =   315
            Index           =   1
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   3120
            Width           =   1395
         End
         Begin VB.CommandButton cmdModelQ 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Info"
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
            Left            =   2340
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   1440
            Width           =   855
         End
         Begin VB.CommandButton cmdCEPHB_Q 
            Caption         =   "XP"
            Height          =   315
            Index           =   3
            Left            =   1380
            TabIndex        =   34
            Top             =   2700
            Width           =   735
         End
         Begin VB.CommandButton cmdCEPHB_Q 
            Caption         =   "Move"
            Height          =   315
            Index           =   2
            Left            =   1380
            TabIndex        =   32
            Top             =   2280
            Width           =   735
         End
         Begin VB.CommandButton cmdCEPHB_Q 
            Caption         =   "Mana"
            Height          =   315
            Index           =   1
            Left            =   1380
            TabIndex        =   30
            Top             =   1860
            Width           =   735
         End
         Begin VB.CommandButton cmdCEPHB_Q 
            Caption         =   "DMG"
            Height          =   315
            Index           =   0
            Left            =   1380
            TabIndex        =   28
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox txtCEPH_DMG 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   720
            MaxLength       =   4
            TabIndex        =   27
            ToolTipText     =   "Min = 0.01, Max = 10.0, Default = 1.0"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtCEPH_XP 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   720
            MaxLength       =   4
            TabIndex        =   33
            ToolTipText     =   "Min = 0.01, Max = 10.0, Default = 1.0"
            Top             =   2700
            Width           =   615
         End
         Begin VB.TextBox txtCEPH_Move 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   720
            MaxLength       =   4
            TabIndex        =   31
            ToolTipText     =   "Min = 0.01, Max = 10.0, Default = 1.0"
            Top             =   2280
            Width           =   615
         End
         Begin VB.TextBox txtCEPH_Mana 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   720
            MaxLength       =   4
            TabIndex        =   29
            ToolTipText     =   "Min = 0.01, Max = 10.0, Default = 1.0"
            Top             =   1860
            Width           =   615
         End
         Begin VB.CheckBox chkExpHrCalcByCharacter 
            Caption         =   "Save these settings per character"
            Height          =   375
            Left            =   600
            TabIndex        =   36
            Top             =   3540
            Width           =   1875
         End
      End
      Begin VB.CheckBox chkDontLookupMonsterRegen 
         Caption         =   "Don't lookup monster spawn locations in detail pane"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   4140
         Width           =   4695
      End
      Begin VB.CheckBox chkSwapWindowTitle 
         Caption         =   "Put the character / filename at the start of MME's window title"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   2880
         Width           =   4815
      End
      Begin VB.CheckBox chkShowCharacterName 
         Caption         =   "Show ""Character Name (File Name)"" in window title"
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   2580
         Width           =   4995
      End
      Begin VB.CheckBox chkAutoCalcMonDamage 
         Caption         =   "Auto-Calculate monster damage vs character stats when clicked"
         Height          =   255
         Left            =   180
         TabIndex        =   17
         Top             =   3540
         Width           =   5175
      End
      Begin VB.CheckBox chkRemoveListEquip 
         Caption         =   "Remove item/spell from saved lists when equipping or learning"
         Height          =   255
         Left            =   180
         TabIndex        =   20
         Top             =   4800
         Width           =   4995
      End
      Begin VB.CheckBox chkWindowSnap 
         Caption         =   "Disable Window/Display Snap (could cause window to get lost on disconnected or reconfigured monitors)"
         Height          =   255
         Left            =   180
         TabIndex        =   25
         Top             =   6300
         Width           =   8955
      End
      Begin VB.CheckBox chkNavSpan 
         Caption         =   "Don't span navigation buttons on resize"
         Height          =   255
         Left            =   180
         TabIndex        =   22
         Top             =   5400
         Width           =   3435
      End
      Begin VB.CheckBox chkUseWrist 
         Caption         =   "Use Second Wrist Slot"
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   1380
         Width           =   2175
      End
      Begin VB.CheckBox chkHideRecordNumbers 
         Caption         =   "Hide record numbers when referencing names"
         Height          =   255
         Left            =   180
         TabIndex        =   21
         Top             =   5100
         Width           =   3735
      End
      Begin VB.CheckBox chkWindowsOnTop 
         Caption         =   "Don't make Results window stay on top of main window"
         Height          =   255
         Left            =   180
         TabIndex        =   24
         Top             =   6000
         Width           =   4695
      End
      Begin VB.CheckBox chkSwapMapButtons 
         Caption         =   "Swap left/right mouse buttons for maps"
         Height          =   255
         Left            =   180
         TabIndex        =   23
         Top             =   5700
         Width           =   3435
      End
      Begin VB.CheckBox chkAutoSaveChar 
         Caption         =   "Always Auto-Save Loaded Character"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   2280
         Width           =   3135
      End
      Begin VB.CheckBox chkAutoLoadChar 
         Caption         =   "Auto-Load last character associated with database"
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   1680
         Width           =   4155
      End
      Begin VB.CheckBox chkFilterAll 
         Caption         =   """ Filter All "" on program load"
         Enabled         =   0   'False
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   1980
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "More Monster-Related Stuff:"
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
         Left            =   5640
         TabIndex        =   43
         Top             =   1080
         Width           =   2400
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Interface-Related Stuff:"
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
         Left            =   180
         TabIndex        =   42
         Top             =   4500
         Width           =   2010
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Monster-Related Stuff:"
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
         Left            =   180
         TabIndex        =   41
         Top             =   3240
         Width           =   1920
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Character-Related Stuff:"
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
         Left            =   180
         TabIndex        =   40
         Top             =   1080
         Width           =   2070
      End
      Begin VB.Label Label1 
         Caption         =   "Rounds to sim when calculating mon dmg (more = more accurate but slower)"
         Height          =   435
         Left            =   6420
         TabIndex        =   39
         Top             =   1440
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
      Top             =   6900
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
      Top             =   6900
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

If bCharLoaded And sSessionLastCharFile <> "" Then
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
    Case 1: MsgBox "Scales mana recovery. Smaller multiplier = less time spent recovering mana.", vbInformation
    Case 2: MsgBox "Scales movement. Larger multiplier = more time spent moving.", vbInformation
    Case 3: MsgBox "Directly multiplies the exp/hr result before being returned.", vbInformation
End Select
End Sub

Private Sub cmdModelQ_Click(Index As Integer)
If Index = 0 Then
    MsgBox "I used ChatGPT to help build models A-C. Model A came first using a limited set of inputs." & vbCrLf & vbCrLf _
        & "Model B was built second using gained knowledge and additional information, but there are instances where model A provides more accurate predictions." & vbCrLf & vbCrLf _
        & "Model C adds more input variables, such as first round damage and min round damage, to provide more accuracy over just using an average." & vbCrLf & vbCrLf _
		& "Model D was with Claude where I asked it to evaluate the other models and it concluded that C was already good, but tweaked some of its deficiencies." & vbCrLf & vbCrLf _
        & "No Recovery: average all chosen models for movement and kill time, based on chosen attack (damage out), but provide just the standard filter for incoming damage.  e.g. No recovery time will be accounted for." & vbCrLf & vbCrLf _
        & "Show in Detail: will show each model's estimates in the detail pane when viewing a monster." & vbCrLf & vbCrLf _
        & "Chosen models will be averaged together.", vbInformation + vbOKOnly
ElseIf Index = 1 Then
    'reset
    txtCEPH_DMG.Text = 1
    txtCEPH_Mana.Text = 1
    txtCEPH_Move.Text = 1
    txtCEPH_XP.Text = 1
    
    chkEPH_Model(0).Value = 1
    chkEPH_Model(1).Value = 1
    chkEPH_Model(2).Value = 1
    chkEPH_Model(3).Value = 1
    chkEPH_Model(98).Value = 0
    chkEPH_Model(99).Value = 0
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
If val(ReadINI("Settings", "AutoLoadLastChar", , 1)) > 0 Then chkAutoLoadChar.Value = 1
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
If val(ReadINI("Settings", "MobPrintCharDamageOutFirst")) > 0 Then chkSwapMonDetailOutput.Value = 1
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

txtCEPH_DMG.Text = nGlobal_cephDMG_Knob
txtCEPH_Mana.Text = nGlobal_cephMana_Knob
txtCEPH_Move.Text = nGlobal_cephMove_Knob
txtCEPH_XP.Text = nGlobal_cephXP_Knob

chkEPH_Model(0).Value = IIf(bGlobal_cephModelA, 1, 0)
chkEPH_Model(1).Value = IIf(bGlobal_cephModelB, 1, 0)
chkEPH_Model(2).Value = IIf(bGlobal_cephModelC, 1, 0)
chkEPH_Model(3).Value = IIf(bGlobal_cephModelD, 1, 0)
chkEPH_Model(98).Value = IIf(bGlobal_cephRecoveryOnly, 1, 0)
chkEPH_Model(99).Value = IIf(bGlobal_cephShowAll, 1, 0)

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


nGlobal_cephDMG_Knob = Round(val(txtCEPH_DMG.Text), 2)
If nGlobal_cephDMG_Knob < 0.01 Then nGlobal_cephDMG_Knob = 0.01
If nGlobal_cephDMG_Knob > 10 Then nGlobal_cephDMG_Knob = 10

nGlobal_cephMana_Knob = Round(val(txtCEPH_Mana.Text), 2)
If nGlobal_cephMana_Knob < 0.01 Then nGlobal_cephMana_Knob = 0.01
If nGlobal_cephMana_Knob > 10 Then nGlobal_cephMana_Knob = 10

nGlobal_cephMove_Knob = Round(val(txtCEPH_Move.Text), 2)
If nGlobal_cephMove_Knob < 0.01 Then nGlobal_cephMove_Knob = 0.01
If nGlobal_cephMove_Knob > 10 Then nGlobal_cephMove_Knob = 10

nGlobal_cephXP_Knob = Round(val(txtCEPH_XP.Text), 2)
If nGlobal_cephXP_Knob < 0.01 Then nGlobal_cephXP_Knob = 0.01
If nGlobal_cephXP_Knob > 10 Then nGlobal_cephXP_Knob = 10

If chkEPH_Model(0).Value = 1 Then
    bGlobal_cephModelA = True
Else
    bGlobal_cephModelA = False
End If
If chkEPH_Model(1).Value = 1 Then
    bGlobal_cephModelB = True
Else
    bGlobal_cephModelB = False
End If
If chkEPH_Model(2).Value = 1 Then
    bGlobal_cephModelC = True
Else
    bGlobal_cephModelC = False
End If
If chkEPH_Model(3).Value = 1 Then
    bGlobal_cephModelD = True
Else
    bGlobal_cephModelD = False
End If

If bGlobal_cephModelA = False And bGlobal_cephModelB = False And bGlobal_cephModelC = False And bGlobal_cephModelD = False Then
    bGlobal_cephModelA = True
    bGlobal_cephModelB = True
    bGlobal_cephModelC = True
End If

If chkEPH_Model(98).Value = 1 Then
    bGlobal_cephRecoveryOnly = True
Else
    bGlobal_cephRecoveryOnly = False
End If

If chkEPH_Model(99).Value = 1 Then
    bGlobal_cephShowAll = True
Else
    bGlobal_cephShowAll = False
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
Call WriteINI("Settings", "MobPrintCharDamageOutFirst", chkSwapMonDetailOutput.Value)
Call WriteINI("Settings", "MonsterSimRounds", nGlobalMonsterSimRounds)

Call WriteINI("Settings", "ExpPerHourKnobsByCharacter", chkExpHrCalcByCharacter.Value)
If chkExpHrCalcByCharacter.Value = 1 And bCharLoaded = True And sSessionLastCharFile <> "" Then sFile = sSessionLastCharFile

Call WriteINI("Settings", "cephDMG", nGlobal_cephDMG_Knob, sFile)
Call WriteINI("Settings", "cephMana", nGlobal_cephMana_Knob, sFile)
Call WriteINI("Settings", "cephMove", nGlobal_cephMove_Knob, sFile)
Call WriteINI("Settings", "cephXP", nGlobal_cephXP_Knob, sFile)
Call WriteINI("Settings", "cephModelA", IIf(bGlobal_cephModelA, 1, 0), sFile)
Call WriteINI("Settings", "cephModelB", IIf(bGlobal_cephModelB, 1, 0), sFile)
Call WriteINI("Settings", "cephModelC", IIf(bGlobal_cephModelC, 1, 0), sFile)
Call WriteINI("Settings", "cephModelD", IIf(bGlobal_cephModelD, 1, 0), sFile)
Call WriteINI("Settings", "cephRecovery", IIf(bGlobal_cephRecoveryOnly, 1, 0), sFile)
Call WriteINI("Settings", "cephShowAll", IIf(bGlobal_cephShowAll, 1, 0), sFile)

If chkDontLookupMonsterRegen.Value = 1 Then
    frmMain.bDontLookupMonRegen = True
Else
    frmMain.bDontLookupMonRegen = False
End If

If chkSwapMonDetailOutput.Value = 1 Then
    bMobPrintCharDamageOutFirst = True
Else
    bMobPrintCharDamageOutFirst = False
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

If chkShowCharacterName.Value = 1 Then
    frmMain.bNameInTitle = True
Else
    frmMain.bNameInTitle = False
End If

If chkUseWrist.Value = 1 Then
    frmMain.chkEquipHold(7).Enabled = True
    frmMain.cmdEquipGoto(7).Enabled = True
    frmMain.cmbEquip(7).Enabled = True
    bInvenUse2ndWrist = True
Else
    frmMain.chkEquipHold(7).Enabled = False
    frmMain.cmdEquipGoto(7).Enabled = False
    frmMain.cmbEquip(7).Enabled = False
    bInvenUse2ndWrist = False
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
Call frmMain.Form_Resize_Event
Unload Me
Exit Sub
error:
Call HandleError("cmdSave_Click")
Resume out:
End Sub


Private Sub txtCEPH_DMG_GotFocus()
Call SelectAll(txtCEPH_DMG)
End Sub

Private Sub txtCEPH_DMG_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii, True)
End Sub

Private Sub txtCEPH_Mana_GotFocus()
Call SelectAll(txtCEPH_Mana)
End Sub

Private Sub txtCEPH_Mana_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii, True)
End Sub

Private Sub txtCEPH_Move_GotFocus()
Call SelectAll(txtCEPH_Move)
End Sub

Private Sub txtCEPH_Move_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii, True)
End Sub

Private Sub txtCEPH_XP_GotFocus()
Call SelectAll(txtCEPH_XP)
End Sub

Private Sub txtCEPH_XP_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii, True)
End Sub

Private Sub txtMonsterSimRounds_GotFocus()
Call SelectAll(txtMonsterSimRounds)
End Sub

Private Sub txtMonsterSimRounds_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub



