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
      Begin VB.Frame Frame1 
         Caption         =   "Monster Exp/Dmg Calculations"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4275
         Left            =   5460
         TabIndex        =   25
         Top             =   1080
         Width           =   4275
         Begin VB.CheckBox chkExpHrCalcByCharacter 
            Caption         =   "Save these settings per character"
            Height          =   195
            Left            =   180
            TabIndex        =   39
            Top             =   3840
            Width           =   3195
         End
         Begin VB.TextBox txtManaScaleFactor 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   180
            MaxLength       =   4
            TabIndex        =   37
            ToolTipText     =   "Min = 0, Max =2.0, Default = 1 (lower = less time spent recovering mana)"
            Top             =   1380
            Width           =   615
         End
         Begin VB.CommandButton cmdMoveNote 
            Caption         =   "?"
            Height          =   315
            Left            =   3780
            TabIndex        =   36
            Top             =   2520
            Width           =   315
         End
         Begin VB.TextBox txtGlobalRouteBias 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   180
            MaxLength       =   4
            TabIndex        =   34
            ToolTipText     =   "Min = 0, Max = 2, Default = 1 (lower = less movement time)"
            Top             =   3240
            Width           =   615
         End
         Begin VB.TextBox txtGlobalRoomDensityRef 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   180
            MaxLength       =   4
            TabIndex        =   32
            ToolTipText     =   "Min = 0, Max = 0.99, Default = 0.25 (lower = less movement time)"
            Top             =   2520
            Width           =   615
         End
         Begin VB.TextBox txtMovementRecoveryRatio 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   180
            MaxLength       =   4
            TabIndex        =   30
            ToolTipText     =   "Min = 0, Max =2.0, Default = 0.85 (higher = more healing while moving)"
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtDmgScaleFactor 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   180
            MaxLength       =   4
            TabIndex        =   28
            ToolTipText     =   "Min = 0, Max = 2.0, Default = 1 (lower = less damage/resting)"
            Top             =   900
            Width           =   615
         End
         Begin VB.TextBox txtMonsterSimRounds 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   180
            MaxLength       =   5
            TabIndex        =   26
            ToolTipText     =   "Min = 100, Max = 10000, Default = 500 (more rounds = more accurate but slow calc time)"
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Scale factor for time spent recovering mana."
            Height          =   315
            Index           =   3
            Left            =   900
            TabIndex        =   38
            Top             =   1440
            Width           =   3255
         End
         Begin VB.Label Label3 
            Caption         =   "Scale factor for the lair/room ratio effecting movement [mostly applies to smaller areas / fewer lairs)"
            Height          =   615
            Index           =   2
            Left            =   900
            TabIndex        =   35
            Top             =   3120
            Width           =   2835
         End
         Begin VB.Label Label3 
            Caption         =   "Scale factor for the lair/room density effecting movement [mostly applies to larger areas / plenty of lairs)"
            Height          =   615
            Index           =   1
            Left            =   900
            TabIndex        =   33
            Top             =   2400
            Width           =   2895
         End
         Begin VB.Label Label3 
            Caption         =   "Fraction of movement time that can be credited to resting time"
            Height          =   495
            Index           =   0
            Left            =   900
            TabIndex        =   31
            Top             =   1860
            Width           =   2775
         End
         Begin VB.Label Label2 
            Caption         =   "Scale factor for excess damage (equating to resting time). Tweak in small increments."
            Height          =   435
            Left            =   900
            TabIndex        =   29
            Top             =   840
            Width           =   3195
         End
         Begin VB.Label Label1 
            Caption         =   "Rounds to sim when calculating mon dmg (more = more accurate but slower)"
            Height          =   435
            Left            =   900
            TabIndex        =   27
            Top             =   300
            Width           =   3015
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
         Height          =   315
         Left            =   6840
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "All"
         Height          =   315
         Left            =   5640
         TabIndex        =   8
         Top             =   240
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
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   180
         X2              =   7920
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

Private Sub chkAutoLoadChar_Click()
DoEvents
If chkAutoLoadChar.Value = 1 Then
    chkFilterAll.Enabled = True
Else
    chkFilterAll.Enabled = False
End If
End Sub

Private Sub cmdAll_Click()

chkLoadItems.Value = 1
chkLoadSpells.Value = 1
'chkLoadClasses.Value = 1
'chkLoadRaces.Value = 1
chkLoadMonsters.Value = 1
chkLoadShops.Value = 1

End Sub

Private Sub cmdMoveNote_Click()
MsgBox "Note: Movement time can be limited on the fly via menu option.", vbInformation
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

chkLoadItems.Value = Val(ReadINI("Settings", "LoadItems", , 1))
chkLoadSpells.Value = Val(ReadINI("Settings", "LoadSpells", , 1))
chkLoadMonsters.Value = Val(ReadINI("Settings", "LoadMonsters", , 1))
chkLoadShops.Value = Val(ReadINI("Settings", "LoadShops", , 1))
chkInGame.Value = Val(ReadINI("Settings", "OnlyInGame", , 1))
chkFilterAll.Value = Val(ReadINI("Settings", "FilterAll"))
chkAutoLoadChar.Value = Val(ReadINI("Settings", "AutoLoadLastChar"))
chkAutoSaveChar.Value = Val(ReadINI("Settings", "AutoSaveLastChar"))
chkSwapMapButtons.Value = Val(ReadINI("Settings", "SwapMapButtons"))
chkWindowsOnTop.Value = Val(ReadINI("Settings", "NoAlwaysOnTop"))
chkHideRecordNumbers.Value = Val(ReadINI("Settings", "HideRecordNumbers"))
chkUseWrist.Value = Val(ReadINI("Settings", "Use2ndWrist", , 1))
chkShowCharacterName.Value = Val(ReadINI("Settings", "NameInTitle"))
chkNavSpan.Value = Val(ReadINI("Settings", "DontSpanNavButtons"))
chkWindowSnap.Value = Val(ReadINI("Settings", "DisableWindowSnap"))
chkRemoveListEquip.Value = Val(ReadINI("Settings", "RemoveListEquip"))
chkAutoCalcMonDamage.Value = Val(ReadINI("Settings", "AutoCalcMonDamage", , "1"))
chkSwapWindowTitle.Value = Val(ReadINI("Settings", "SwapWindowTitle"))
chkDontLookupMonsterRegen.Value = Val(ReadINI("Settings", "DontLookupMonsterRegen"))

chkExpHrCalcByCharacter.Value = Val(ReadINI("Settings", "ExpPerHourKnobsByCharacter"))
txtMonsterSimRounds.Text = nGlobalMonsterSimRounds
txtDmgScaleFactor.Text = nGlobalDmgScaleFactor
txtManaScaleFactor.Text = nGlobalManaScaleFactor
txtMovementRecoveryRatio.Text = nGlobalMovementRecoveryRatio
txtGlobalRoomDensityRef.Text = nGlobalRoomDensityRef
txtGlobalRouteBias.Text = nGlobalRoomRouteBias

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

nGlobalDmgScaleFactor = Val(txtDmgScaleFactor.Text)
If nGlobalDmgScaleFactor < 0 Then nGlobalDmgScaleFactor = 0
If nGlobalDmgScaleFactor > 2# Then nGlobalDmgScaleFactor = 2#

nGlobalManaScaleFactor = Val(txtManaScaleFactor.Text)
If nGlobalManaScaleFactor < 0 Then nGlobalManaScaleFactor = 0
If nGlobalManaScaleFactor > 2# Then nGlobalManaScaleFactor = 2#

nGlobalMonsterSimRounds = Val(txtMonsterSimRounds.Text)
If nGlobalMonsterSimRounds < 100 Then nGlobalMonsterSimRounds = 100
If nGlobalMonsterSimRounds > 10000 Then nGlobalMonsterSimRounds = 10000

nGlobalMovementRecoveryRatio = Val(txtMovementRecoveryRatio.Text)
If nGlobalMovementRecoveryRatio < 0 Then nGlobalMovementRecoveryRatio = 0
If nGlobalMovementRecoveryRatio > 2# Then nGlobalMovementRecoveryRatio = 2#

nGlobalRoomDensityRef = Round(Val(txtGlobalRoomDensityRef.Text), 2)
If nGlobalRoomDensityRef < 0 Then nGlobalRoomDensityRef = 0
If nGlobalRoomDensityRef > 0.99 Then nGlobalRoomDensityRef = 0.99

nGlobalRoomRouteBias = Round(Val(txtGlobalRouteBias.Text), 2)
If nGlobalRoomRouteBias < 0 Then nGlobalRoomDensityRef = 0
If nGlobalRoomRouteBias > 2# Then nGlobalRoomDensityRef = 2#

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

Call WriteINI("Settings", "ExpPerHourKnobsByCharacter", chkExpHrCalcByCharacter.Value)

If chkExpHrCalcByCharacter.Value = 1 And frmMain.bCharLoaded = True And sSessionLastCharFile <> "" Then sFile = sSessionLastCharFile
Call WriteINI("Settings", "MonsterSimRounds", nGlobalMonsterSimRounds, sFile)
Call WriteINI("Settings", "DmgScaleFactor", nGlobalDmgScaleFactor, sFile)
Call WriteINI("Settings", "ManaScaleFactor", nGlobalManaScaleFactor, sFile)
Call WriteINI("Settings", "MovementRecoveryRatio", nGlobalMovementRecoveryRatio, sFile)
Call WriteINI("Settings", "RoomDensityRef", nGlobalRoomDensityRef, sFile)
Call WriteINI("Settings", "RoomRouteBias", nGlobalRoomRouteBias, sFile)

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

out:
On Error Resume Next
Unload Me
Exit Sub
error:
Call HandleError("cmdSave_Click")
Resume out:
End Sub

Private Sub txtDmgScaleFactor_GotFocus()
Call SelectAll(txtDmgScaleFactor)
End Sub

Private Sub txtDmgScaleFactor_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii, True)
End Sub

Private Sub txtManaScaleFactor_GotFocus()
Call SelectAll(txtManaScaleFactor)
End Sub

Private Sub txtManaScaleFactor_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii, True)
End Sub

Private Sub txtMovementRecoveryRatio_GotFocus()
Call SelectAll(txtMovementRecoveryRatio)
End Sub

Private Sub txtMovementRecoveryRatio_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii, True)
End Sub

Private Sub txtMonsterSimRounds_GotFocus()
Call SelectAll(txtMonsterSimRounds)
End Sub

Private Sub txtMonsterSimRounds_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtGlobalRoomDensityRef_GotFocus()
Call SelectAll(txtGlobalRoomDensityRef)
End Sub

Private Sub txtGlobalRoomDensityRef_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii, True)
End Sub

Private Sub txtGlobalRouteBias_GotFocus()
Call SelectAll(txtGlobalRouteBias)
End Sub

Private Sub txtGlobalRouteBias_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii, True)
End Sub
