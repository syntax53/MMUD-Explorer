VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   5625
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRecreateINI 
      Caption         =   "Recreate settings.ini"
      Height          =   375
      Left            =   1860
      TabIndex        =   3
      Top             =   4140
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      Caption         =   "Settings"
      Height          =   3975
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   5475
      Begin VB.CheckBox chkNavSpan 
         Caption         =   "Don't span navigation buttons on resize"
         Height          =   435
         Left            =   180
         TabIndex        =   19
         Top             =   3360
         Width           =   2235
      End
      Begin VB.CheckBox chkShowCharacterName 
         Caption         =   "Show character names instead of filenames"
         Height          =   435
         Left            =   2640
         TabIndex        =   18
         Top             =   3360
         Width           =   2475
      End
      Begin VB.CheckBox chkUseWrist 
         Caption         =   "Use Second Wrist Slot"
         Height          =   255
         Left            =   180
         TabIndex        =   17
         Top             =   1140
         Width           =   2175
      End
      Begin VB.CheckBox chkHideRecordNumbers 
         Caption         =   "Hide record number when referencing names"
         Height          =   435
         Left            =   180
         TabIndex        =   16
         Top             =   2760
         Width           =   2355
      End
      Begin VB.CheckBox chkWindowsOnTop 
         Caption         =   "Don't make Result, Exp, Swing Calc windows stay on top of main window"
         Height          =   675
         Left            =   180
         TabIndex        =   15
         Top             =   2040
         Width           =   2235
      End
      Begin VB.CheckBox chkSwapMapButtons 
         Caption         =   "Swap left/right mouse buttons for map."
         Height          =   375
         Left            =   180
         TabIndex        =   14
         Top             =   1560
         Width           =   2295
      End
      Begin VB.CheckBox chkAutoSaveChar 
         Caption         =   "Auto-Save Loaded Character (Per Database Version)"
         Height          =   375
         Left            =   2640
         TabIndex        =   13
         Top             =   2160
         Width           =   2415
      End
      Begin VB.CheckBox chkAutoLoadChar 
         Caption         =   "Auto-Load Last Character (Per Database Version)"
         Height          =   495
         Left            =   2640
         TabIndex        =   12
         Top             =   1500
         Width           =   2355
      End
      Begin VB.CheckBox chkFilterAll 
         Caption         =   """ Filter All "" on load"
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   1140
         Width           =   1755
      End
      Begin VB.CheckBox chkInGame 
         Caption         =   "Only load items, monsters, and shops that are in the game"
         Height          =   435
         Left            =   2640
         TabIndex        =   10
         Top             =   2760
         Width           =   2475
      End
      Begin VB.CommandButton cmdNone 
         Caption         =   "None"
         Height          =   315
         Left            =   2760
         TabIndex        =   9
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "All"
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   180
         Width           =   1035
      End
      Begin VB.CheckBox chkLoadShops 
         Caption         =   "Load Shops"
         Height          =   255
         Left            =   4140
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkLoadMonsters 
         Caption         =   "Load Monsters"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   600
         Width           =   1395
      End
      Begin VB.CheckBox chkLoadSpells 
         Caption         =   "Load Spells"
         Height          =   255
         Left            =   1380
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkLoadItems 
         Caption         =   "Load Items"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   5340
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
      Left            =   4320
      TabIndex        =   1
      Top             =   4140
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
      Top             =   4140
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

chkLoadItems.Value = ReadINI("Settings", "LoadItems")
chkLoadSpells.Value = ReadINI("Settings", "LoadSpells")
chkLoadMonsters.Value = ReadINI("Settings", "LoadMonsters")
chkLoadShops.Value = ReadINI("Settings", "LoadShops")
chkInGame.Value = ReadINI("Settings", "OnlyInGame")
chkFilterAll.Value = ReadINI("Settings", "FilterAll")
chkAutoLoadChar.Value = ReadINI("Settings", "AutoLoadLastChar")
chkAutoSaveChar.Value = ReadINI("Settings", "AutoSaveLastChar")
chkSwapMapButtons.Value = ReadINI("Settings", "SwapMapButtons")
chkWindowsOnTop.Value = ReadINI("Settings", "NoAlwaysOnTop")
chkHideRecordNumbers.Value = ReadINI("Settings", "HideRecordNumbers")
chkUseWrist.Value = ReadINI("Settings", "Use2ndWrist")
chkShowCharacterName.Value = ReadINI("Settings", "NameInTitle")
chkNavSpan.Value = ReadINI("Settings", "DontSpanNavButtons")

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdRecreateINI_Click()

frmMain.bReloadProgram = True
Unload Me

End Sub

Private Sub cmdSave_Click()
Dim sSectionName As String, x As Integer, nWidth As Long, nTwipsEnlarged As Long

On Error GoTo Error:

sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ")

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

For x = 0 To 10
    Select Case x
        Case 0:
            frmMain.cmdNav(x).Width = 1095 + nTwipsEnlarged
        Case 1:
            frmMain.cmdNav(x).Left = frmMain.cmdNav(x - 1).Left + frmMain.cmdNav(x - 1).Width - 15
            frmMain.cmdNav(x).Width = 855 + nTwipsEnlarged
        Case 2:
            frmMain.cmdNav(x).Left = frmMain.cmdNav(x - 1).Left + frmMain.cmdNav(x - 1).Width - 15
            frmMain.cmdNav(x).Width = 795 + nTwipsEnlarged
        Case 3:
            frmMain.cmdNav(x).Left = frmMain.cmdNav(x - 1).Left + frmMain.cmdNav(x - 1).Width - 15
            frmMain.cmdNav(x).Width = 975 + nTwipsEnlarged
        Case 4:
            frmMain.cmdNav(x).Left = frmMain.cmdNav(x - 1).Left + frmMain.cmdNav(x - 1).Width - 15
            frmMain.cmdNav(x).Width = 1215 + nTwipsEnlarged
        Case 5:
            frmMain.cmdNav(x).Left = frmMain.cmdNav(x - 1).Left + frmMain.cmdNav(x - 1).Width - 15
            frmMain.cmdNav(x).Width = 1035 + nTwipsEnlarged
        Case 6:
            frmMain.cmdNav(x).Left = frmMain.cmdNav(x - 1).Left + frmMain.cmdNav(x - 1).Width - 15
            frmMain.cmdNav(x).Width = 1215 + nTwipsEnlarged
        Case 7:
            frmMain.cmdNav(x).Left = frmMain.cmdNav(x - 1).Left + frmMain.cmdNav(x - 1).Width - 15
            frmMain.cmdNav(x).Width = 795 + nTwipsEnlarged
        Case 8:
            frmMain.cmdNav(x).Left = frmMain.cmdNav(x - 1).Left + frmMain.cmdNav(x - 1).Width - 15
            frmMain.cmdNav(x).Width = 975 + nTwipsEnlarged
        Case 9:
            frmMain.cmdNav(x).Left = frmMain.cmdNav(x - 1).Left + frmMain.cmdNav(x - 1).Width - 15
            frmMain.cmdNav(x).Width = 735 + nTwipsEnlarged
        Case 10:
            frmMain.cmdNav(x).Left = frmMain.cmdNav(x - 1).Left + frmMain.cmdNav(x - 1).Width - 15
            frmMain.cmdNav(x).Width = 795 + nTwipsEnlarged
    End Select
Next x

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

Unload Me

Exit Sub
Error:
Call HandleError("cmdSave_Click")
End Sub

