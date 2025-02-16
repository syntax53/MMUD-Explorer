VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   8325
   Begin VB.CommandButton cmdRecreateINI 
      Caption         =   "Recreate settings.ini"
      Height          =   375
      Left            =   3060
      TabIndex        =   3
      Top             =   5580
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      Caption         =   "Settings"
      Height          =   5355
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   8175
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
         Height          =   3375
         Left            =   5040
         TabIndex        =   25
         Top             =   1080
         Width           =   2955
         Begin VB.TextBox txtTheoreticalAvgMaxLairsPerRegenPeriod 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   180
            MaxLength       =   3
            TabIndex        =   32
            ToolTipText     =   "Min = 1, Max = 100, Default = 36 (larger number = larger penalty for fewer lairs)"
            Top             =   2760
            Width           =   615
         End
         Begin VB.TextBox txtMonsterLairRatioMultiplier 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   180
            MaxLength       =   2
            TabIndex        =   30
            ToolTipText     =   "Min = 1, Max = 10, Default = 3 (larger = require more non-lairs to have an effect)"
            Top             =   2040
            Width           =   615
         End
         Begin VB.TextBox txtDmgScaleFactor 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   180
            MaxLength       =   5
            TabIndex        =   28
            ToolTipText     =   "Min = 0.1, Max = 2.0, Default = 0.9 (lower = less resting)"
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtMonsterSimRounds 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   180
            MaxLength       =   5
            TabIndex        =   26
            ToolTipText     =   "Min = 100, Max = 10000, Default = 500 (more rounds = more accurate but slow calc time)"
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "maximum number of single mob lairs that can be cleared before regen"
            Height          =   675
            Left            =   900
            TabIndex        =   33
            ToolTipText     =   "Min = 100, Max = 10000"
            Top             =   2580
            Width           =   1935
         End
         Begin VB.Label Label3 
            Caption         =   "multiplier for non-lairs vs lairs to start reducing exp due to travel time"
            Height          =   615
            Left            =   900
            TabIndex        =   31
            ToolTipText     =   "Min = 100, Max = 10000"
            Top             =   1860
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Scale factor for excess damage = resting time.  Tweak in small increments."
            Height          =   675
            Left            =   900
            TabIndex        =   29
            Top             =   1020
            Width           =   1995
         End
         Begin VB.Label Label1 
            Caption         =   "rounds to sim mon dmg (accuracy vs speed)"
            Height          =   435
            Left            =   900
            TabIndex        =   27
            ToolTipText     =   "Min = 100, Max = 10000"
            Top             =   300
            Width           =   1815
         End
      End
      Begin VB.CheckBox chkDontLookupMonsterRegen 
         Caption         =   "Don't lookup monster in detail (performance)"
         Height          =   255
         Left            =   180
         TabIndex        =   20
         Top             =   3780
         Width           =   7695
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
         Height          =   255
         Left            =   180
         TabIndex        =   24
         Top             =   4980
         Width           =   7875
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
         Caption         =   "Don't make Result, Exp, Swing Calc windows stay on top of main window"
         Height          =   255
         Left            =   180
         TabIndex        =   23
         Top             =   4680
         Width           =   5895
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
      Left            =   6960
      TabIndex        =   1
      Top             =   5580
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
      Left            =   60
      TabIndex        =   0
      Top             =   5580
      Width           =   1155
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

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

Private Sub cmdNone_Click()

chkLoadItems.Value = 0
chkLoadSpells.Value = 0
'chkLoadClasses.Value = 0
'chkLoadRaces.Value = 0
chkLoadMonsters.Value = 0
chkLoadShops.Value = 0

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim sSectionName As String

sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ")

chkLoadItems.Value = ReadINI("Settings", "LoadItems", , 1)
chkLoadSpells.Value = ReadINI("Settings", "LoadSpells", , 1)
chkLoadMonsters.Value = ReadINI("Settings", "LoadMonsters", , 1)
chkLoadShops.Value = ReadINI("Settings", "LoadShops", , 1)
chkInGame.Value = ReadINI("Settings", "OnlyInGame", , 1)
chkFilterAll.Value = ReadINI("Settings", "FilterAll")
chkAutoLoadChar.Value = ReadINI("Settings", "AutoLoadLastChar")
chkAutoSaveChar.Value = ReadINI("Settings", "AutoSaveLastChar")
chkSwapMapButtons.Value = ReadINI("Settings", "SwapMapButtons")
chkWindowsOnTop.Value = ReadINI("Settings", "NoAlwaysOnTop")
chkHideRecordNumbers.Value = ReadINI("Settings", "HideRecordNumbers")
chkUseWrist.Value = ReadINI("Settings", "Use2ndWrist", , 1)
chkShowCharacterName.Value = ReadINI("Settings", "NameInTitle")
chkNavSpan.Value = ReadINI("Settings", "DontSpanNavButtons")
chkWindowSnap.Value = ReadINI("Settings", "DisableWindowSnap")
chkRemoveListEquip.Value = ReadINI("Settings", "RemoveListEquip")
chkAutoCalcMonDamage.Value = ReadINI("Settings", "AutoCalcMonDamage", , "1")
chkSwapWindowTitle.Value = ReadINI("Settings", "SwapWindowTitle")
chkDontLookupMonsterRegen.Value = ReadINI("Settings", "DontLookupMonsterRegen")
txtMonsterSimRounds.Text = nMonsterSimRounds
txtDmgScaleFactor.Text = nDmgScaleFactor
txtMonsterLairRatioMultiplier.Text = nMonsterLairRatioMultiplier
txtTheoreticalAvgMaxLairsPerRegenPeriod.Text = nTheoreticalAvgMaxLairsPerRegenPeriod

Call chkAutoLoadChar_Click

If frmMain.WindowState = vbMinimized Then
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
Else
    Me.Left = frmMain.Left + ((frmMain.Width - Me.Width) / 2)
    Me.Top = frmMain.Top + ((frmMain.Height - Me.Height) / 2)
End If

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
Dim sSectionName As String, X As Integer, nWidth As Long, nTwipsEnlarged As Long

On Error GoTo error:

sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ")

nDmgScaleFactor = Val(txtDmgScaleFactor.Text)
If nDmgScaleFactor < 0.1 Then nDmgScaleFactor = 0.1
If nDmgScaleFactor > 2 Then nDmgScaleFactor = 2#

nMonsterSimRounds = Val(txtMonsterSimRounds.Text)
If nMonsterSimRounds < 100 Then nMonsterSimRounds = 100
If nMonsterSimRounds > 10000 Then nMonsterSimRounds = 10000

nMonsterLairRatioMultiplier = Val(txtMonsterLairRatioMultiplier.Text)
If nMonsterLairRatioMultiplier < 1 Then nMonsterLairRatioMultiplier = 1
If nMonsterLairRatioMultiplier > 10 Then nMonsterLairRatioMultiplier = 10

nTheoreticalAvgMaxLairsPerRegenPeriod = Val(txtTheoreticalAvgMaxLairsPerRegenPeriod.Text)
If nTheoreticalAvgMaxLairsPerRegenPeriod < 1 Then nTheoreticalAvgMaxLairsPerRegenPeriod = 1
If nTheoreticalAvgMaxLairsPerRegenPeriod > 100 Then nTheoreticalAvgMaxLairsPerRegenPeriod = 100

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
Call WriteINI("Settings", "MonsterSimRounds", nMonsterSimRounds)
Call WriteINI("Settings", "DmgScaleFactor", nDmgScaleFactor)
Call WriteINI("Settings", "MonsterLairRatioMultiplier", nMonsterLairRatioMultiplier)
Call WriteINI("Settings", "TheoreticalAvgMaxLairsPerRegenPeriod", nTheoreticalAvgMaxLairsPerRegenPeriod)

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

nWidth = frmMain.Width - 210

If chkNavSpan.Value = 0 Then
    nTwipsEnlarged = Fix((frmMain.Width - 10695) / 11)
    'Debug.Print nTwipsEnlarged
    frmMain.framButtons.Width = nWidth
    frmMain.fraDatVer.Width = nWidth - 6115
    frmMain.lblDatVer.Width = frmMain.fraDatVer.Width - 140
Else
    nWidth = 10485
    nTwipsEnlarged = 0
    'Debug.Print nTwipsEnlarged
    frmMain.framButtons.Width = nWidth
    frmMain.fraDatVer.Width = nWidth - 6115
    frmMain.lblDatVer.Width = frmMain.fraDatVer.Width - 140
End If

For X = 0 To 10
    Select Case X
        Case 0:
            frmMain.cmdNav(X).Width = 1095 + nTwipsEnlarged
        Case 1:
            frmMain.cmdNav(X).Left = frmMain.cmdNav(X - 1).Left + frmMain.cmdNav(X - 1).Width - 15
            frmMain.cmdNav(X).Width = 855 + nTwipsEnlarged
        Case 2:
            frmMain.cmdNav(X).Left = frmMain.cmdNav(X - 1).Left + frmMain.cmdNav(X - 1).Width - 15
            frmMain.cmdNav(X).Width = 795 + nTwipsEnlarged
        Case 3:
            frmMain.cmdNav(X).Left = frmMain.cmdNav(X - 1).Left + frmMain.cmdNav(X - 1).Width - 15
            frmMain.cmdNav(X).Width = 975 + nTwipsEnlarged
        Case 4:
            frmMain.cmdNav(X).Left = frmMain.cmdNav(X - 1).Left + frmMain.cmdNav(X - 1).Width - 15
            frmMain.cmdNav(X).Width = 1215 + nTwipsEnlarged
        Case 5:
            frmMain.cmdNav(X).Left = frmMain.cmdNav(X - 1).Left + frmMain.cmdNav(X - 1).Width - 15
            frmMain.cmdNav(X).Width = 1035 + nTwipsEnlarged
        Case 6:
            frmMain.cmdNav(X).Left = frmMain.cmdNav(X - 1).Left + frmMain.cmdNav(X - 1).Width - 15
            frmMain.cmdNav(X).Width = 1215 + nTwipsEnlarged
        Case 7:
            frmMain.cmdNav(X).Left = frmMain.cmdNav(X - 1).Left + frmMain.cmdNav(X - 1).Width - 15
            frmMain.cmdNav(X).Width = 795 + nTwipsEnlarged
        Case 8:
            frmMain.cmdNav(X).Left = frmMain.cmdNav(X - 1).Left + frmMain.cmdNav(X - 1).Width - 15
            frmMain.cmdNav(X).Width = 975 + nTwipsEnlarged
        Case 9:
            frmMain.cmdNav(X).Left = frmMain.cmdNav(X - 1).Left + frmMain.cmdNav(X - 1).Width - 15
            frmMain.cmdNav(X).Width = 735 + nTwipsEnlarged
        Case 10:
            frmMain.cmdNav(X).Left = frmMain.cmdNav(X - 1).Left + frmMain.cmdNav(X - 1).Width - 15
            frmMain.cmdNav(X).Width = 795 + nTwipsEnlarged
    End Select
Next X

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

Unload Me

Exit Sub
error:
Call HandleError("cmdSave_Click")
End Sub

Private Sub txtDmgScaleFactor_GotFocus()
Call SelectAll(txtDmgScaleFactor)
End Sub

Private Sub txtDmgScaleFactor_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii, True)
End Sub

Private Sub txtMonsterSimRounds_GotFocus()
Call SelectAll(txtMonsterSimRounds)
End Sub

Private Sub txtMonsterSimRounds_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub
