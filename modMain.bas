Attribute VB_Name = "modMain"
#Const DEVELOPMENT_MODE = 0 'TURN OFF BEFORE RELEASE - LOC 1/3

#If DEVELOPMENT_MODE Then
    Public Const DEVELOPMENT_MODE_RT As Boolean = True
#Else
    Public Const DEVELOPMENT_MODE_RT As Boolean = False
#End If
Option Explicit
Option Base 0

Global Const TOTAL_STAT_LBLS = 46

Public DebugLogFileHandle As Integer
Public DebugLogFilePath As String
Global bCancelLaunch As Boolean
Global bDPIAwareMode As Boolean
Global bHideRecordNumbers As Boolean
Global bMobPrintCharDamageOutFirst As Boolean
Global bOnlyInGame As Boolean
Global nGlobalMonsterSimRounds As Long
Global bCharLoaded As Boolean
Global bStartup As Boolean
Global bDontSyncSplitters As Boolean
Global nNMRVer As Double
Global nOSversion As cnWin32Ver
Global sCurrentDatabaseFile As String
Global sForceCharacterFile As String
'Global bOnlyLearnable As Boolean

'max 1-mob lairs you can clear in 3 minutes (average lair regen of 3 minutes divided by average round of 5 seconds = 36 lairs
'Global nTheoreticalMaxLairsPerRegenPeriod As Integer

Global sMonsterDamageVsCharDefenseConfig As String
Global sGlobalCharDefenseDescription As String
Global bDontPromptCalcCharMonsterDamage As Boolean
Global bMonsterDamageVsPartyCalculated As Boolean
Global bDontPromptCalcPartyMonsterDamage As Boolean
Global nLastItemSortCol As Integer
Public tLastAvgLairInfo As LairInfoType

'Global nGlobalCharAccyStats As Long
Global nGlobalCharAccyItems As Long
Global nGlobalCharAccyAbils As Long
Global nGlobalCharAccyOther As Long
Global nGlobalCharPlusDodge As Long
Global nGlobalCharPlusMR As Long
Global nGlobalCharQnDbonus As Long
Global nGlobalCharWeaponNumber(1) As Long '0=weapon, 1=offhand
Global nGlobalCharWeaponAccy(1) As Long
Global nGlobalCharWeaponCrit(1) As Long
Global nGlobalCharWeaponSTR(1) As Long
Global nGlobalCharWeaponAGI(1) As Long
Global nGlobalCharWeaponMaxDmg(1) As Long
Global nGlobalCharWeaponBSaccy(1) As Long
Global nGlobalCharWeaponBSmindmg(1) As Long
Global nGlobalCharWeaponBSmaxdmg(1) As Long
Global nGlobalCharWeaponPunchSkill(1) As Long
Global nGlobalCharWeaponPunchAccy(1) As Long
Global nGlobalCharWeaponPunchDmg(1) As Long
Global nGlobalCharWeaponKickSkill(1) As Long
Global nGlobalCharWeaponKickAccy(1) As Long
Global nGlobalCharWeaponKickDmg(1) As Long
Global nGlobalCharWeaponJkSkill(1) As Long
Global nGlobalCharWeaponJkAccy(1) As Long
Global nGlobalCharWeaponJkDmg(1) As Long
Global nGlobalCharWeaponStealth(1) As Long

Public Enum eAttackTypeMME
    a0_oneshot = 0
    a1_PhysAttack = 1
    a2_Spell = 2
    a3_SpellAny = 3
    a4_MartialArts = 4
    a5_Manual = 5
    a6_PhysBash = 6
    a7_PhysSmash = 7
End Enum

Global nGlobalAttackTypeMME As eAttackTypeMME '0-none, 1-weapon, 2/3-spell, 4-MA, 5-manual
Global bGlobalAttackBackstab As Boolean
Global nGlobalAttackBackstabWeapon As Long
Global nGlobalAttackMA As Integer '1-punch, 2-kick, 3-jumpkick
Global nGlobalAttackSpellNum As Long
Global nGlobalAttackSpellLVL As Integer
Global nGlobalAttackManualP As Long
Global nGlobalAttackManualM As Long
Global sGlobalAttackConfig As String
Global bGlobalAttackUseMeditate As Boolean
Global nGlobalAttackHealType As Integer '0-none, 2/3-spell, 4-manual
Global nGlobalAttackHealSpellNum As Long
Global nGlobalAttackHealSpellLVL As Integer
Global nGlobalAttackHealRounds As Integer
Global nGlobalAttackHealManual As Long
Global nGlobalAttackHealValue As Long
Global nGlobalAttackHealCost As Double

Public Type tAbilityToStatSlot
    nEquip As Integer
    sText As String
End Type

Public Enum QBColorCode
    Black = 0
    Blue = 1
    green = 2
    Cyan = 3
    Red = 4
    Magenta = 5
    Yellow = 6
    white = 7
    Grey = 8
    BrightBlue = 9
    BrightGreen = 10
    BrightCyan = 11
    BrightRed = 12
    BrightMagenta = 13
    BrightYellow = 14
    BrightWhite = 15
End Enum

Public Enum eExpandBy
    Percent50 = 0
    Percent75 = 1
    DoubleWidth = 2
    TripleWidth = 3
    QuadWidth = 4
    NoExpand = 5
End Enum
Public Enum eExpandType
    WidthOnly = 0
    HeightOnly = 1
    HeightAndWidth = 2
End Enum

Public Enum eAttackRestrictions
    AR000_Unknown = 0
    AR001_None = &H1
    AR023_Undead = &H2 'abil 23 AffectsUndead / undead flag
    AR080_Animal = &H4 'abil 80 AffectsAnimals / 78 animal
    AR108_Living = &H8 'abil 108 AffectsLiving / 109 NonLiving
End Enum

Public Enum eDefenseFlags
    DF023_IsUndead = &H1 'abil 23 AffectsUndead / undead flag
    DF078_IsAnimal = &H2 'abil 80 AffectsAnimals / 78 animal
    DF109_IsLiving = &H4 'abil 108 AffectsLiving / 109 NonLiving
    DFIAM_IsAntiMag = &H8
End Enum

Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_GETITEMHEIGHT = &H154
Private Const CB_SETDROPPEDWIDTH = &H160
Private Const CB_GETDROPPEDWIDTH = &H15F
Private Const CB_SETDROPPEDCONTROLRECT = &H160
Private Const DT_CALCRECT = &H400
Public Const WM_SETREDRAW As Long = 11

Public bUseDwmAPI As Boolean
Public bPromptSave As Boolean
Public bCancelTerminate As Boolean
Public bAppTerminating As Boolean
Public bAppReallyTerminating As Boolean

Public sRecentFiles(1 To 5, 1 To 2) As String '1=shown, 2=filename
Public sRecentDBs(1 To 5, 1 To 2) As String '1=shown, 2=filename
Public nEquippedItem(0 To 19) As Long
Public nLearnedSpells(0 To 99) As Long
Public nLearnedSpellClass As Integer
Public bLegit As Boolean
Public bDisableKaiAutolearn As Boolean
Public sSessionLastCharFile As String
Public sSessionLastLoadDir As String
Public sSessionLastLoadName As String
Public sSessionLastSaveDir As String
Public sSessionLastSaveName As String

Public clsMonAtkSim As clsMonsterAttackSim

Public LoadChar_CheckFilterOnReload As Boolean
Public LoadChar_chkInvenLoad As Boolean
Public LoadChar_chkInvenClear As Boolean
Public LoadChar_chkCompareLoad As Boolean
Public LoadChar_chkCompareClear As Boolean
Public LoadChar_optFilter As Integer

Public Type POINTAPI
   x As Long
   y As Long
End Type

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private gLogFSO                 As Object  ' Scripting.FileSystemObject
Private gLogTS                  As Object  ' Scripting.TextStream
Private gLogPath                As String
Private gLogImmediateFlush      As Boolean ' True => close & reopen after every line
Private gLogIsAppend            As Boolean

Private Const CSIDL_APPDATA As Long = &H1A    ' \AppData\Roaming
Private Const S_OK As Long = 0&

Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long


Public Declare Function GetWindowRect Lib "user32" _
  (ByVal hWnd As Long, _
   lpRect As RECT) As Long

Private Declare Function ScreenToClient Lib "user32" _
  (ByVal hWnd As Long, _
   lpPoint As POINTAPI) As Long

Public Declare Function SendMessageLong Lib "user32" Alias _
        "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function DrawText Lib "user32" Alias _
    "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, _
    ByVal nCount As Long, lpRect As RECT, ByVal wFormat _
    As Long) As Long

Private Declare Function SHGetFolderPathA Lib "shfolder" ( _
    ByVal hwndOwner As Long, ByVal nFolder As Long, _
    ByVal hToken As Long, ByVal dwFlags As Long, _
    ByVal pszPath As String) As Long

'Public Declare Function CalcExpNeeded Lib "lltmmudxp" (ByVal Level As Long, ByVal Chart As Long) As Currency
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Const VK_LBUTTON = &H1
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long

' ===== Debug controls =====
Public Function Pct(ByVal x As Double) As String: Pct = Format$(x, "0.00%"): End Function
Public Function F6(ByVal x As Double) As String: F6 = Format$(x, "0.000000"): End Function
Public Function F3(ByVal x As Double) As String: F3 = Format$(x, "0.000"): End Function
Public Function F2(ByVal x As Double) As String: F2 = Format$(x, "0.00"): End Function
Public Function F1(ByVal x As Double) As String: F1 = Format$(x, "0.0"): End Function

' Open the log. If sPath="", defaults to App.Path & "\_DebugLog.txt"
' bAppend: False = overwrite; True = append
' bImmediateFlush: True = close/reopen after each WriteLine (forces disk write)
Public Sub InitDebugLog(Optional ByVal sPath As String, _
                        Optional ByVal bAppend As Boolean = False, _
                        Optional ByVal bImmediateFlush As Boolean = True)
    On Error GoTo EH

    If sPath = "" Then sPath = sGlobalWorkingDirectory & "\_DebugLog.txt"
    gLogPath = sPath
    gLogImmediateFlush = bImmediateFlush
    gLogIsAppend = bAppend

    Set gLogFSO = CreateObject("Scripting.FileSystemObject")

    If bAppend And gLogFSO.FileExists(sPath) Then
        ' ForAppending = 8
        Set gLogTS = gLogFSO.OpenTextFile(sPath, 8, True)
    Else
        ' Create (overwrite=True, unicode=False for classic ANSI; set True for Unicode if desired)
        Set gLogTS = gLogFSO.CreateTextFile(sPath, True, False)
    End If
    Exit Sub

EH:
    ' If init fails, leave objects as Nothing (no logging)
End Sub

' Print a single message line to the log (adds CRLF)
Public Sub DebugLogPrint(ByVal Msg As String)
    On Error Resume Next
    If gLogTS Is Nothing Then Exit Sub

    gLogTS.WriteLine Msg

    If gLogImmediateFlush Then
        ' Force the OS to commit by closing and reopening in append mode
        gLogTS.Close
        ' ForAppending = 8
        Set gLogTS = gLogFSO.OpenTextFile(gLogPath, 8, True)
    End If
End Sub

' Close the log
Public Sub DebugCloseLog()
    On Error Resume Next
    If Not gLogTS Is Nothing Then gLogTS.Close
    Set gLogTS = Nothing
    Set gLogFSO = Nothing
    gLogPath = vbNullString
End Sub
'===============================================================

Private Sub Main()
On Error GoTo fail

nOSversion = GetWin32Ver
bCancelLaunch = False

Load frmMain

If bCancelLaunch Or bAppTerminating Then
    If FormIsLoaded("frmMain") Then
        Unload frmMain
    End If
End If

Exit Sub

fail:
ExitApp 1
End Sub

Public Function GetAppDataDir() As String
    Dim buf As String * 260
    If SHGetFolderPathA(0&, CSIDL_APPDATA, 0&, 0&, buf) = S_OK Then
        GetAppDataDir = Left$(buf, InStr(buf, vbNullChar) - 1)
    End If
End Function

Public Sub RefreshCombatHealingValues()
Dim tHealSpell As tSpellCastValues, bUseCharacter As Boolean, tChar As tCharacterProfile
On Error GoTo error:

nGlobalAttackHealCost = 0
nGlobalAttackHealValue = 0
If frmMain.optMonsterFilter(1).Value = True And val(frmMain.txtMonsterLairFilter(0).Text) > 1 Then
    nGlobalAttackHealValue = val(frmMain.txtMonsterDamage.Text)
    Exit Sub
End If

If frmMain.chkGlobalFilter.Value = 1 Then bUseCharacter = True

Select Case nGlobalAttackHealType
    Case 0: 'infinite
        nGlobalAttackHealValue = 999999
    Case 1: 'base
        nGlobalAttackHealValue = 0
    Case 2, 3: 'spell
        If nGlobalAttackHealSpellNum > 0 Then
            If nGlobalAttackHealSpellLVL < 0 Then nGlobalAttackHealSpellLVL = 0
            If nGlobalAttackHealSpellLVL > 9999 Then nGlobalAttackHealSpellLVL = 9999
            If nGlobalAttackHealRounds < 1 Then nGlobalAttackHealRounds = 1
            If nGlobalAttackHealRounds > 50 Then nGlobalAttackHealRounds = 50
            Call PopulateCharacterProfile(tChar, False, True)
            If bUseCharacter Then
                tHealSpell = CalculateSpellCast(tChar, nGlobalAttackHealSpellNum, tChar.nLevel)
            Else
                tHealSpell = CalculateSpellCast(tChar, nGlobalAttackHealSpellNum)
            End If
            nGlobalAttackHealCost = Round(tHealSpell.nManaCost / nGlobalAttackHealRounds, 2)
            nGlobalAttackHealValue = Round(tHealSpell.nAvgCast / nGlobalAttackHealRounds, 2)
        End If
    Case 4: 'manual
        nGlobalAttackHealValue = nGlobalAttackHealManual
End Select

If nGlobalAttackHealCost < 0.25 Then nGlobalAttackHealCost = 0
If nGlobalAttackHealCost > 9999 Then nGlobalAttackHealCost = 9999
If nGlobalAttackHealValue < 0 Then nGlobalAttackHealValue = 0
If nGlobalAttackHealValue > 999999 Then nGlobalAttackHealValue = 999999

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("RefreshCombatHealingValues")
Resume out:
End Sub

Public Sub SetCharDefenseDescription()
On Error GoTo error:
Dim sConfig As String

sConfig = CalcEncumbrancePercent(val(frmMain.lblInvenCharStat(0).Caption), val(frmMain.lblInvenCharStat(1).Caption)) 'encum

If frmMain.cmbGlobalClass(0).ListIndex < 0 Then Exit Sub

If frmMain.chkGlobalFilter.Value = 1 Then
    sConfig = sConfig & "_" & val(frmMain.txtGlobalLevel(0).Text) 'lvl
    sConfig = sConfig & "_" & frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex) 'class
End If

sConfig = sConfig & "_" & frmMain.lblInvenCharStat(2).Tag 'ac
sConfig = sConfig & "_" & frmMain.lblInvenCharStat(3).Tag 'dr
sConfig = sConfig & "_" & frmMain.txtCharMR.Text 'mr
sConfig = sConfig & "_" & frmMain.lblInvenCharStat(8).Tag 'dodge
sConfig = sConfig & "_" & frmMain.chkCharAntiMagic.Value 'anti-magic
sConfig = sConfig & "_" & frmMain.lblInvenCharStat(28).Tag 'rcol
sConfig = sConfig & "_" & frmMain.lblInvenCharStat(27).Tag 'rfir
sConfig = sConfig & "_" & frmMain.lblInvenCharStat(25).Tag 'rsto
sConfig = sConfig & "_" & frmMain.lblInvenCharStat(29).Tag 'rlit
sConfig = sConfig & "_" & frmMain.lblInvenCharStat(26).Tag 'rwat

If sConfig <> sGlobalCharDefenseDescription Then bDontPromptCalcCharMonsterDamage = False
sGlobalCharDefenseDescription = sConfig

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("SetCharDefenseDescription")
Resume out:
End Sub

Public Sub SetCurrentAttackTypeConfig()
On Error GoTo error:
Dim sConfig As String

'sGlobalAttackConfig is a just string that creates a way to test to see if the config changed.
'used during calculcations to determine if it should re-run attack simulations.

sConfig = CStr(nGlobalAttackTypeMME)
If frmMain.chkGlobalFilter.Value = 1 Then
    sConfig = sConfig & "_" & val(frmMain.txtGlobalLevel(0).Text) 'lvl
    sConfig = sConfig & "_" & frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex) 'class
    sConfig = sConfig & "_" & CalcEncumbrancePercent(val(frmMain.lblInvenCharStat(0).Caption), val(frmMain.lblInvenCharStat(1).Caption)) 'encum
End If

Select Case nGlobalAttackTypeMME
    Case 1, 6, 7, 4: 'weap, bash, smash, MA, backstab
        sConfig = sConfig & "_" & nGlobalCharWeaponNumber(0)
        sConfig = sConfig & "_" & nGlobalCharWeaponNumber(1)
        If frmMain.chkGlobalFilter.Value = 1 Then
            sConfig = sConfig & "_" & val(frmMain.txtCharStats(0).Tag) 'str
            sConfig = sConfig & "_" & val(frmMain.txtCharStats(3).Tag) 'agi
            sConfig = sConfig & "_" & val(frmMain.lblInvenCharStat(7).Tag) - nGlobalCharQnDbonus 'nCritChance
            sConfig = sConfig & "_" & val(frmMain.lblInvenCharStat(11).Tag) 'nPlusMaxDamage
            sConfig = sConfig & "_" & val(frmMain.lblInvenCharStat(30).Tag) 'nPlusMinDamage
            sConfig = sConfig & "_" & val(frmMain.lblInvenCharStat(10).Tag) 'nAttackAccuracy
            sConfig = sConfig & "_" & val(frmMain.lblInvenCharStat(13).Tag) 'nPlusBSaccy
            sConfig = sConfig & "_" & val(frmMain.lblInvenCharStat(14).Tag) 'nPlusBSmindmg
            sConfig = sConfig & "_" & val(frmMain.lblInvenCharStat(15).Tag) 'nPlusBSmaxdmg
            sConfig = sConfig & "_" & val(frmMain.lblInvenCharStat(19).Tag) 'nStealth
            If val(frmMain.lblInvenStats(19).Tag) >= 2 Then sConfig = sConfig & "1" 'bClassStealth
        Else
            sConfig = sConfig & "_default"
        End If
        
    Case 2, 3: 'spell
        sConfig = sConfig & "_" & nGlobalAttackSpellNum
        If frmMain.chkGlobalFilter.Value = 1 Then
            sConfig = sConfig & "_" & val(frmMain.txtGlobalLevel(0).Text)
        Else
            sConfig = sConfig & "_" & nGlobalAttackSpellLVL
        End If
        If bGlobalAttackUseMeditate Then sConfig = sConfig & "_med"
        
    Case 5: 'manual
        sConfig = sConfig & "_" & CStr(nGlobalAttackManualP) & "_" & CStr(nGlobalAttackManualM)
End Select

If nGlobalAttackTypeMME = a4_MartialArts Then 'MA
    sConfig = sConfig & "_" & CStr(nGlobalAttackMA)
    Select Case nGlobalAttackMA
        Case 1: 'punch
            sConfig = sConfig & "_" & CStr(nGlobalCharWeaponPunchSkill(0) + nGlobalCharWeaponPunchSkill(1))
            sConfig = sConfig & "_" & CStr(nGlobalCharWeaponPunchAccy(0) + nGlobalCharWeaponPunchAccy(1))
            sConfig = sConfig & "_" & CStr(nGlobalCharWeaponPunchDmg(0) + nGlobalCharWeaponPunchDmg(1))
        Case 2: 'kick
            sConfig = sConfig & "_" & CStr(nGlobalCharWeaponKickSkill(0) + nGlobalCharWeaponKickSkill(1))
            sConfig = sConfig & "_" & CStr(nGlobalCharWeaponKickAccy(0) + nGlobalCharWeaponKickAccy(1))
            sConfig = sConfig & "_" & CStr(nGlobalCharWeaponKickDmg(0) + nGlobalCharWeaponKickDmg(1))
        Case 3: 'jk
            sConfig = sConfig & "_" & CStr(nGlobalCharWeaponJkSkill(0) + nGlobalCharWeaponJkSkill(1))
            sConfig = sConfig & "_" & CStr(nGlobalCharWeaponJkAccy(0) + nGlobalCharWeaponJkAccy(1))
            sConfig = sConfig & "_" & CStr(nGlobalCharWeaponJkDmg(0) + nGlobalCharWeaponJkDmg(1))
    End Select
End If

If nGlobalAttackTypeMME > a0_oneshot And bGlobalAttackBackstab Then sConfig = sConfig & "_BS" & CStr(nGlobalAttackBackstabWeapon)

If sGlobalAttackConfig <> sConfig Then Call ClearSavedDamageVsMonster
sGlobalAttackConfig = sConfig

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("SetCurrentAttackTypeKEY")
Resume out:
End Sub

Public Function IsDllAvailable(ByVal DllName As String) As Boolean
On Error GoTo error:
Dim hLib As Long

hLib = LoadLibrary(DllName)
If hLib <> 0 Then
    IsDllAvailable = True
    FreeLibrary hLib
Else
    IsDllAvailable = False
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("IsDllAvailable")
Resume out:
End Function

Public Sub LearnOrUnlearnSpell(nSpell As Long)
On Error GoTo error:

If in_long_arr(ByVal nSpell, nLearnedSpells()) Then
    Call UnLearnSpell(nSpell)
Else
    Call LearnSpell(nSpell)
End If

If FormIsLoaded("frmPopUpOptions") Then Unload frmPopUpOptions

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("LearnOrUnlearnSpell")
Resume out:
End Sub

Public Sub LearnSpell(nSpell As Long)
Dim x As Integer
On Error GoTo error:

For x = 0 To 99
    If nLearnedSpells(x) = 0 Then
        nLearnedSpells(x) = nSpell
        If nLearnedSpells(x) > 0 And frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex) > 0 And nLearnedSpellClass <> frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex) Then
            nLearnedSpellClass = frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex)
        End If
        Exit For
    End If
Next x

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("LearnSpell")
Resume out:
End Sub

Public Sub UnLearnSpell(nSpell As Long)
Dim x As Integer
On Error GoTo error:

For x = 0 To 99
    If nLearnedSpells(x) = nSpell Then
        nLearnedSpells(x) = 0
    End If
Next x

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("UnLearnSpell")
Resume out:
End Sub

Public Sub ExpandCombo(ByRef Combo As ComboBox, ByVal ExpandType As eExpandType, _
    ByVal ExpandBy As eExpandBy, Optional ByVal hFrame As Long)

    Dim lRet As Long
    Dim pt As POINTAPI
    Dim rc As RECT
    Dim lComboWidth As Long
    Dim lNewHeight As Long
    Dim lItemHeight As Long
    
    If ExpandType <> HeightOnly Then
        lComboWidth = (Combo.Width / Screen.TwipsPerPixelX)
        Select Case ExpandBy
            Case 0:
                lComboWidth = lComboWidth + (lComboWidth * 0.5)
            Case 1:
                lComboWidth = lComboWidth + (lComboWidth * 0.75)
            Case 2:
                lComboWidth = lComboWidth * 2
            Case 3:
                lComboWidth = lComboWidth * 3
            Case 4:
                lComboWidth = lComboWidth * 4
        End Select
        lRet = SendMessage(Combo.hWnd, CB_SETDROPPEDCONTROLRECT, lComboWidth, 0)
        
    End If
    
    If ExpandType <> WidthOnly Then
        lComboWidth = Combo.Width / Screen.TwipsPerPixelX
        lItemHeight = SendMessage(Combo.hWnd, CB_GETITEMHEIGHT, 0, 0)
        Select Case ExpandBy
            Case 1:
                'lComboWidth = lComboWidth + (lComboWidth * 0.75)
                lNewHeight = lItemHeight * 16
            Case 2:
                'lComboWidth = lComboWidth * 2
                lNewHeight = lItemHeight * 18
            Case 3:
                'lComboWidth = lComboWidth * 3
                lNewHeight = lItemHeight * 26
            Case 4:
                'lComboWidth = lComboWidth * 4
                lNewHeight = lItemHeight * 32
            Case Else:
                lNewHeight = lItemHeight * 14
                'lComboWidth = lComboWidth + (lComboWidth * 0.5)
        End Select
        Call GetWindowRect(Combo.hWnd, rc)
        pt.x = rc.Left
        pt.y = rc.Top
        Call ScreenToClient(hFrame, pt)
        Call MoveWindow(Combo.hWnd, pt.x, pt.y, lComboWidth, lNewHeight, True)
    End If
    
End Sub

Public Function AutoSizeDropDownWidth(Combo As Object) As Boolean
'**************************************************************
'PURPOSE: Automatically size the combo box drop down width
'         based on the width of the longest item in the combo box

'PARAMETERS: Combo - ComboBox to size

'RETURNS: True if successful, false otherwise

'ASSUMPTIONS: 1. Form's Scale Mode is vbTwips, which is why
'                conversion from twips to pixels are made.
'                API functions require units in pixels
'
'             2. Combo Box's parent is a form or other
'                container that support the hDC property

'EXAMPLE: AutoSizeDropDownWidth Combo1
'****************************************************************
Dim lRet As Long
Dim lCurrentWidth As Single
Dim rectCboText As RECT
Dim lParentHDC As Long
Dim lListCount As Long
Dim lCtr As Long
Dim lTempWidth As Long
Dim lWidth As Long
Dim sSavedFont As String
Dim sngSavedSize As Single
Dim bSavedBold As Boolean
Dim bSavedItalic As Boolean
Dim bSavedUnderline As Boolean
Dim bFontSaved As Boolean

On Error GoTo errorHandler

If Not TypeOf Combo Is ComboBox Then Exit Function
lParentHDC = Combo.Parent.hdc
If lParentHDC = 0 Then Exit Function
lListCount = Combo.ListCount
If lListCount = 0 Then Exit Function


'Change font of parent to combo box's font
'Save first so it can be reverted when finished
'this is necessary for drawtext API Function
'which is used to determine longest string in combo box
With Combo.Parent

    sSavedFont = .FontName
    sngSavedSize = .FontSize
    bSavedBold = .FontBold
    bSavedItalic = .FontItalic
    bSavedUnderline = .FontUnderline
    
    .FontName = Combo.FontName
    .FontSize = Combo.FontSize
    .FontBold = Combo.FontBold
    .FontItalic = Combo.FontItalic
    .FontUnderline = Combo.FontItalic

End With

bFontSaved = True

'Get the width of the largest item
For lCtr = 0 To lListCount
   DrawText lParentHDC, Combo.List(lCtr), -1, rectCboText, _
        DT_CALCRECT
   'adjust the number added (20 in this case to
   'achieve desired right margin
   lTempWidth = rectCboText.Right - rectCboText.Left + 20

   If (lTempWidth > lWidth) Then
      lWidth = lTempWidth
   End If
Next
 
lCurrentWidth = SendMessageLong(Combo.hWnd, CB_GETDROPPEDWIDTH, _
    0, 0)

If lCurrentWidth > lWidth Then 'current drop-down width is
'                               sufficient

    AutoSizeDropDownWidth = True
    GoTo errorHandler
    Exit Function
End If
 
'don't allow drop-down width to
'exceed screen.width
 
   If lWidth > Screen.Width \ Screen.TwipsPerPixelX - 20 Then _
    lWidth = Screen.Width \ Screen.TwipsPerPixelX - 20

lRet = SendMessageLong(Combo.hWnd, CB_SETDROPPEDWIDTH, lWidth, 0)

AutoSizeDropDownWidth = lRet > 0
errorHandler:
On Error Resume Next
If bFontSaved Then
'restore parent's font settings
  With Combo.Parent
    .FontName = sSavedFont
    .FontSize = sngSavedSize
    .FontUnderline = bSavedUnderline
    .FontBold = bSavedBold
    .FontItalic = bSavedItalic
 End With
End If
End Function

Public Sub PullItemDetail(DetailTB As TextBox, LocationLV As ListView, Optional ByVal nAttackTypeMUD As eAttackTypeMUD, Optional ByVal bFullDetails As Boolean)
Dim sStr As String, sAbil As String, x As Integer, sCasts As String, nPercent As Integer
Dim sNegate As String, sClasses As String, sRaces As String, sClassOk As String
Dim sUses As String, sGetDrop As String, oLI As ListItem, nNumber As Long
Dim y As Integer, bCompareWeapon As Boolean, bCompareArmor As Boolean, nInvenSlot1 As Integer, nInvenSlot2 As Integer
Dim sCompareText1 As String, sCompareText2 As String, tabItems1 As Recordset, tabItems2 As Recordset
Dim sTemp1 As String, sTemp2 As String, sTemp3 As String, bFlag1 As Boolean, bFlag2 As Boolean
Dim nClassRestrictions(0 To 2, 0 To 9) As Long, nRaceRestrictions(0 To 2, 0 To 9) As Long
Dim nNegateSpells(0 To 2, 0 To 9) As Long, nAbils(0 To 2, 0 To 19, 0 To 2) As Long, sAbilText(0 To 2, 0 To 19) As String
Dim nReturnValue As Long, nMatchReturnValue As Long, sClassOk1 As String, sClassOk2 As String
Dim sCastSp1 As String, sCastSp2 As String, bCastSpFlag(0 To 2) As Boolean, nPct(0 To 2) As Integer, bForceCalc As Boolean
Dim tWeaponDmg As tAttackDamage, sWeaponDmg As String, nSpeedAdj As Integer, bCalcCombat As Boolean, bUseCharacter As Boolean
Dim tCharacter As tCharacterProfile, nBSacc As Integer, nLVLreq As Integer, nAcc As Integer ', bGetsSpellBonus As Boolean

On Error GoTo error:

DetailTB.Text = ""
If bStartup Then Exit Sub

nNumber = tabItems.Fields("Number")

If frmMain.chkGlobalFilter.Value = 1 Then bUseCharacter = True
If frmMain.chkWeaponOptions(3).Value = 1 Then bCalcCombat = True

nSpeedAdj = 100
If bCalcCombat Then
    If nAttackTypeMUD = 0 Then nAttackTypeMUD = frmMain.cmbWeaponCombos(1).ItemData(frmMain.cmbWeaponCombos(1).ListIndex)
    If frmMain.chkWeaponOptions(4).Value = 1 Then nSpeedAdj = 85
ElseIf nAttackTypeMUD = 0 Then
    nAttackTypeMUD = 5
End If

If tabItems.Fields("ItemType") = 1 Then Call PopulateCharacterProfile(tCharacter, bUseCharacter, True, nAttackTypeMUD, nNumber)
If Not tabItems.Fields("Number") = nNumber Then tabItems.Seek "=", nNumber

nInvenSlot1 = -1
nInvenSlot2 = -1

Select Case tabItems.Fields("ItemType")
    Case 0: 'armour
        If tabItems.Fields("Worn") <= 0 Then
            'nada
        Else
            If tabItems.Fields("Worn") <= UBound(nEquippedItem) Then
                Select Case tabItems.Fields("Worn")
                    Case 0: '"Nowhere"
                    Case 1: '"Everywhere"
                        nInvenSlot1 = 19
                    Case 2: '"Head"
                        nInvenSlot1 = 0
                    Case 3: '"Hands"
                        nInvenSlot1 = 8
                    Case 4, 13: '"Finger"
                        If nEquippedItem(9) > 0 Then
                            nInvenSlot1 = 9
                            If nEquippedItem(10) > 0 Then nInvenSlot2 = 10
                        ElseIf nEquippedItem(10) > 0 Then
                            nInvenSlot1 = 10
                        End If
                    Case 5: '"Feet"
                        nInvenSlot1 = 13
                    Case 6: '"Arms"
                        nInvenSlot1 = 5
                    Case 7: '"Back"
                        nInvenSlot1 = 3
                    Case 8: '"Neck"
                        nInvenSlot1 = 2
                    Case 9: '"Legs"
                        nInvenSlot1 = 12
                    Case 10: '"Waist"
                        nInvenSlot1 = 11
                    Case 11: '"Torso"
                        nInvenSlot1 = 4
                    Case 12: '"Off-Hand"
                        nInvenSlot1 = 15
                    Case 14: '"Wrist"
                        If nEquippedItem(6) > 0 Then
                            nInvenSlot1 = 6
                            If nEquippedItem(7) > 0 Then nInvenSlot2 = 7
                        ElseIf nEquippedItem(7) > 0 Then
                            nInvenSlot1 = 7
                        End If
                    Case 15: '"Ears"
                        nInvenSlot1 = 1
                    Case 16: '"Worn"
                        nInvenSlot1 = 14
                    Case 18: '"Eyes"
                        nInvenSlot1 = 17
                    Case 19: '"Face"
                        nInvenSlot1 = 18
                    Case Else:
                End Select
                
                If nInvenSlot1 >= 0 Then
                    If nEquippedItem(nInvenSlot1) > 0 Then
                        bCompareArmor = True
                    Else
                        nInvenSlot1 = -1
                    End If
                End If
                
                If nInvenSlot2 >= 0 Then
                    If nEquippedItem(nInvenSlot2) > 0 Then
                        bCompareArmor = True
                    Else
                        nInvenSlot2 = -1
                    End If
                End If
                
            End If
        End If
        
    Case 1: 'weapons
        nInvenSlot1 = 16
        If nEquippedItem(nInvenSlot1) > 0 Then
            bCompareWeapon = True
        End If
        
    Case Else: 'other
        'nada
        
End Select

If Not bCompareWeapon And Not bCompareArmor Then
    nInvenSlot1 = -1
    nInvenSlot2 = -1
End If

If nInvenSlot1 >= 0 Then
    If nEquippedItem(nInvenSlot1) = nNumber Then nInvenSlot1 = -1
End If
If nInvenSlot2 >= 0 Then
    If nEquippedItem(nInvenSlot2) = nNumber Then nInvenSlot2 = -1
End If

If nInvenSlot2 >= 0 And nInvenSlot1 < 0 Then
    nInvenSlot1 = nInvenSlot2
    nInvenSlot2 = -1
End If

If nInvenSlot1 < 0 And nInvenSlot2 < 0 Then
    bCompareWeapon = False
    bCompareArmor = False
End If

If bCompareWeapon Or bCompareArmor And nInvenSlot1 >= 0 Then
    Set tabItems1 = DB.OpenRecordset("Items")
    tabItems1.Index = "pkItems"
    tabItems1.Seek "=", nEquippedItem(nInvenSlot1)
    If tabItems1.NoMatch = True Then
        If nInvenSlot2 < 0 Then
            bCompareWeapon = False
            bCompareArmor = False
            nInvenSlot1 = -1
            nInvenSlot2 = -1
            tabItems1.Close
            Set tabItems1 = Nothing
        End If
    End If
    
    If nInvenSlot2 >= 0 Then
        Set tabItems2 = DB.OpenRecordset("Items")
        tabItems2.Index = "pkItems"
        tabItems2.Seek "=", nEquippedItem(nInvenSlot2)
        If tabItems2.NoMatch = True Then
            If nInvenSlot1 < 0 Then
                bCompareWeapon = False
                bCompareArmor = False
                nInvenSlot1 = -1
                nInvenSlot2 = -1
                tabItems2.Close
                Set tabItems2 = Nothing
            End If
        End If
    End If
End If

'#################
If tabItems.Fields("UseCount") > 0 Then
    sUses = tabItems.Fields("UseCount")
    If tabItems.Fields("Retain After Uses") = 1 Then
        sUses = sUses & " start/max"
    Else
        sUses = sUses & " (destroys after uses)"
    End If
End If

If nInvenSlot1 >= 0 Then
    If tabItems1.Fields("UseCount") > 0 Then
        sTemp1 = tabItems1.Fields("UseCount")
        If tabItems1.Fields("Retain After Uses") = 1 Then
            sTemp1 = "Uses: " & sTemp1 & " start/max"
        Else
            sTemp1 = "Uses: " & sTemp1 & " (destroys after uses)"
        End If
    End If
    If sTemp1 <> sUses Then sCompareText1 = AutoAppend(sCompareText1, sTemp1)
End If

If nInvenSlot2 >= 0 Then
    If tabItems2.Fields("UseCount") > 0 Then
        sTemp2 = tabItems2.Fields("UseCount")
        If tabItems2.Fields("Retain After Uses") = 1 Then
            sTemp2 = "Uses: " & sTemp2 & " start/max"
        Else
            sTemp2 = "Uses: " & sTemp2 & " (destroys after uses)"
        End If
    End If
    If sTemp2 <> sUses Then sCompareText2 = AutoAppend(sCompareText2, sTemp2)
End If


'#################
If tabItems.Fields("Gettable") = 0 Then sGetDrop = AutoAppend(sGetDrop, "Not Getable")
If nInvenSlot1 >= 0 Then
    If tabItems.Fields("Gettable") <> tabItems1.Fields("Gettable") Then
        If tabItems1.Fields("Gettable") = 0 Then
            sCompareText1 = AutoAppend(sCompareText1, "+Getable")
        Else
            sCompareText1 = AutoAppend(sCompareText1, "+Not Getable")
        End If
    End If
End If
If nInvenSlot2 >= 0 Then
    If tabItems.Fields("Gettable") <> tabItems2.Fields("Gettable") Then
        If tabItems2.Fields("Gettable") = 0 Then
            sCompareText2 = AutoAppend(sCompareText2, "+Getable")
        Else
            sCompareText2 = AutoAppend(sCompareText2, "+Not Getable")
        End If
    End If
End If

'#################
If tabItems.Fields("Not Droppable") = 1 Then sGetDrop = AutoAppend(sGetDrop, "Not Droppable")
If nInvenSlot1 >= 0 Then
    If tabItems.Fields("Not Droppable") <> tabItems1.Fields("Not Droppable") Then
        If tabItems1.Fields("Not Droppable") = 1 Then
            sCompareText1 = AutoAppend(sCompareText1, "+Droppable")
        Else
            sCompareText1 = AutoAppend(sCompareText1, "+Not Droppable")
        End If
    End If
End If
If nInvenSlot2 >= 0 Then
    If tabItems.Fields("Not Droppable") <> tabItems2.Fields("Not Droppable") Then
        If tabItems2.Fields("Not Droppable") = 1 Then
            sCompareText2 = AutoAppend(sCompareText2, "+Droppable")
        Else
            sCompareText2 = AutoAppend(sCompareText2, "+Not Droppable")
        End If
    End If
End If

'#################
If tabItems.Fields("Destroy On Death") = 1 Then sGetDrop = AutoAppend(sGetDrop, "Destroys On Death")
If nInvenSlot1 >= 0 Then
    If tabItems.Fields("Destroy On Death") <> tabItems1.Fields("Destroy On Death") Then
        If tabItems1.Fields("Destroy On Death") = 1 Then
            sCompareText1 = AutoAppend(sCompareText1, "-Destroy On Death")
        Else
            sCompareText1 = AutoAppend(sCompareText1, "+Destroy On Death")
        End If
    End If
End If
If nInvenSlot2 >= 0 Then
    If tabItems.Fields("Destroy On Death") <> tabItems2.Fields("Destroy On Death") Then
        If tabItems2.Fields("Destroy On Death") = 1 Then
            sCompareText2 = AutoAppend(sCompareText2, "-Destroy On Death")
        Else
            sCompareText2 = AutoAppend(sCompareText2, "+Destroy On Death")
        End If
    End If
End If

'#################

For x = 0 To 19
    
    If tabItems.Fields("Abil-" & x) > 0 Then
        nAbils(0, x, 0) = tabItems.Fields("Abil-" & x)
        'If getindex_array_long_3d(nAbils, nAbils(0, x, 0), nReturnValue, 2, 0, , , x, 0) Then
        '    nAbils(0, nReturnValue, 1) = nAbils(0, nReturnValue, 1) + tabItems.Fields("AbilVal-" & x)
        '    nAbils(0, x, 0) = 0
        'Else
            nAbils(0, x, 1) = tabItems.Fields("AbilVal-" & x)
        'End If
    End If
    
    If nInvenSlot1 >= 0 Then
        If tabItems1.Fields("Abil-" & x) > 0 Then
            nAbils(1, x, 0) = tabItems1.Fields("Abil-" & x)
            'If getindex_array_long_3d(nAbils, nAbils(1, x, 0), nReturnValue, 2, 1, , , x, 0) Then
            '    nAbils(1, nReturnValue, 1) = nAbils(1, nReturnValue, 1) + tabItems1.Fields("AbilVal-" & x)
            '    nAbils(1, x, 0) = 0
            'Else
                nAbils(1, x, 1) = tabItems1.Fields("AbilVal-" & x)
            'End If
        End If
    End If
    
    If nInvenSlot2 >= 0 Then
        If tabItems2.Fields("Abil-" & x) > 0 Then
            nAbils(2, x, 0) = tabItems2.Fields("Abil-" & x)
            'If getindex_array_long_3d(nAbils, nAbils(2, x, 0), nReturnValue, 2, 2, , , x, 0) Then
            '    nAbils(2, nReturnValue, 1) = nAbils(2, nReturnValue, 1) + tabItems2.Fields("AbilVal-" & x)
            '    nAbils(2, x, 0) = 0
            'Else
                nAbils(2, x, 1) = tabItems2.Fields("AbilVal-" & x)
            'End If
        End If
    End If
Next


For x = 0 To 19
    If nAbils(0, x, 0) > 0 Then
        Select Case nAbils(0, x, 0)
            Case 116: '116-bsacc
                'If Not DetailTB.name = "txtWeaponCompareDetail" And _
                    Not DetailTB.name = "txtWeaponDetail" Then
                    
                    sTemp1 = GetAbilityStats(nAbils(0, x, 0), nAbils(0, x, 1), LocationLV, , True)
                    sAbilText(0, x) = sTemp1
                    sAbil = AutoAppend(sAbil, sTemp1)
                'End If
                nBSacc = nAbils(0, x, 1)
            Case 22, 105, 106, 135:  '22-acc, 105-acc, 106-acc, 135-minlvl
                'If Not DetailTB.name = "txtWeaponCompareDetail" And _
                '    Not DetailTB.name = "txtWeaponDetail" And _
                '    Not DetailTB.name = "txtArmourCompareDetail" And _
                '    Not DetailTB.name = "txtArmourDetail" Then
    
                    sTemp1 = GetAbilityStats(nAbils(0, x, 0), nAbils(0, x, 1), LocationLV, , True)
                    sAbilText(0, x) = sTemp1
                    sAbil = AutoAppend(sAbil, sTemp1)
                'End If
                If nAbils(0, x, 0) = 135 Then
                    nLVLreq = nAbils(0, x, 1)
                Else
                    nAcc = nAcc + nAbils(0, x, 1)
                End If
            Case 59: 'class ok
                sTemp1 = GetClassName(nAbils(0, x, 1))
                sAbilText(0, x) = sTemp1
                sClassOk = AutoAppend(sClassOk, sTemp1)
                
            Case 43: 'casts spell
                'nSpellNest = 0 'make sure this doesn't nest too deep
                sCasts = AutoAppend(sCasts, "[" & GetSpellName(nAbils(0, x, 1), bHideRecordNumbers) _
                    & ", " & PullSpellEQ(True, 0, nAbils(0, x, 1), , , , True, , , , , tCharacter.nSpellDmgBonus))
                If Not nPercent = 0 Then
                    sCasts = sCasts & ", " & nPercent & "%]"
                Else
                    sCasts = sCasts & "]"
                End If
                sAbilText(0, x) = sCasts
                
                'Set oLI = LocationLV.ListItems.Add
                'oLI.Text = ""
                'oLI.ListSubItems.Add 1, , "Casts: " & GetSpellName(nAbils(0, x, 1), bHideRecordNumbers)
                'oLI.ListSubItems(1).Tag = nAbils(0, x, 1)
            
            Case 114: '%spell
                nPercent = nAbils(0, x, 1)
                
            Case Else:
                sTemp1 = GetAbilityStats(nAbils(0, x, 0), nAbils(0, x, 1), LocationLV, , True)
                sAbilText(0, x) = sTemp1
                sAbil = AutoAppend(sAbil, sTemp1)
                
        End Select
    End If
Next x
nAcc = nAcc + tabItems.Fields("Accy")


If nInvenSlot1 >= 0 Then
    sTemp1 = ""
    sTemp2 = ""
    bCastSpFlag(0) = False
    bCastSpFlag(1) = False
    bCastSpFlag(2) = False
    bFlag1 = False
    bFlag2 = False
    For y = 0 To 2
        
        nPct(0) = 0
        nPct(1) = 0
        nPct(2) = 0
        For x = 0 To 19
            
            nMatchReturnValue = -32000
            
            If nAbils(y, x, 0) > 0 Then
            
                If nAbils(y, x, 0) = 59 Then nMatchReturnValue = nAbils(y, x, 1) 'classok
                If nAbils(y, x, 0) = 43 Then
                    nMatchReturnValue = nAbils(y, x, 1) 'casts spell
                    bCastSpFlag(y) = True
                End If
                If nAbils(y, x, 0) = 114 Then nPct(y) = nAbils(y, x, 1) '%spell
                
'                If nAbils(y, x, 0) = 117 Then
'                    Debug.Print 1
'                End If
                
                If y = 0 Then
                    
                    If Not getval_array_long_3d(nAbils, nAbils(y, x, 0), nReturnValue, 1, nMatchReturnValue, 1, , , , 0) Then
                        
                        sTemp3 = GetAbilDiffText(nAbils(y, x, 0), nAbils(y, x, 1), 0, sAbilText(y, x), nPct(y))
                        If Len(sTemp3) > 0 Then
                            If nAbils(y, x, 0) = 59 Then 'classok
                                sClassOk1 = AutoAppend(sClassOk1, "+" & sTemp3)
                            ElseIf nAbils(y, x, 0) = 43 Then 'casts spell
                                sCastSp1 = AutoAppend(sCastSp1, "+" & sTemp3)
                            Else
                                bFlag1 = True
                                sTemp1 = AutoAppend(sTemp1, IIf(nAbils(y, x, 1) = 0, "+", "") & sTemp3)
                            End If
                        End If
                        
                    ElseIf nReturnValue <> nAbils(y, x, 1) Then
                        
                        sTemp3 = GetAbilDiffText(nAbils(y, x, 0), nAbils(y, x, 1), nReturnValue, sAbilText(y, x), nPct(y))
                        If Len(sTemp3) > 0 Then
                            bFlag1 = True
                            sTemp1 = AutoAppend(sTemp1, sTemp3)
                        End If
                        
                    End If
                    
                    If nInvenSlot2 >= 0 Then
                    
                        If Not getval_array_long_3d(nAbils, nAbils(y, x, 0), nReturnValue, 1, nMatchReturnValue, 2, , , , 0) Then
                        
                            sTemp3 = GetAbilDiffText(nAbils(y, x, 0), nAbils(y, x, 1), 0, sAbilText(y, x), nPct(y))
                            If Len(sTemp3) > 0 Then
                                If nAbils(y, x, 0) = 59 Then 'classok
                                    sClassOk2 = AutoAppend(sClassOk2, "+" & sTemp3)
                                ElseIf nAbils(y, x, 0) = 43 Then 'casts spell
                                    sCastSp2 = AutoAppend(sCastSp2, "+" & sTemp3)
                                Else
                                    bFlag2 = True
                                    sTemp2 = AutoAppend(sTemp2, IIf(nAbils(y, x, 1) = 0, "+", "") & sTemp3)
                                End If
                            End If
                            
                        ElseIf nReturnValue <> nAbils(y, x, 1) Then
                        
                            sTemp3 = GetAbilDiffText(nAbils(y, x, 0), nAbils(y, x, 1), nReturnValue, sAbilText(y, x), nPct(y))
                            If Len(sTemp3) > 0 Then
                                bFlag2 = True
                                sTemp2 = AutoAppend(sTemp2, sTemp3)
                            End If
                            
                        End If
                        
                    End If
                Else
                    If Not getval_array_long_3d(nAbils, nAbils(y, x, 0), nReturnValue, 1, nMatchReturnValue, 0, , , , 0) Then
                        
                        sTemp3 = GetAbilDiffText(nAbils(y, x, 0), nAbils(y, x, 1), 0, , nPct(y), True)
                        If Len(sTemp3) > 0 Then
                            If nAbils(y, x, 0) = 59 Then 'classok
                                If y = 1 Then
                                    sClassOk1 = AutoAppend(sClassOk1, "-" & sTemp3)
                                Else
                                    sClassOk2 = AutoAppend(sClassOk2, "-" & sTemp3)
                                End If
                            ElseIf nAbils(y, x, 0) = 43 Then 'casts spell
                                If y = 1 Then
                                    sCastSp1 = AutoAppend(sCastSp1, "-" & sTemp3)
                                Else
                                    sCastSp2 = AutoAppend(sCastSp2, "-" & sTemp3)
                                End If
                            Else
                                If y = 1 Then
                                    bFlag1 = True
                                    sTemp1 = AutoAppend(sTemp1, IIf(nAbils(y, x, 1) = 0, "-", "") & sTemp3)
                                Else
                                    bFlag2 = True
                                    sTemp2 = AutoAppend(sTemp2, IIf(nAbils(y, x, 1) = 0, "-", "") & sTemp3)
                                End If
                            End If
                        End If
                        
                    End If
                End If
            End If
        Next x
    Next y
    
    If Len(sTemp1) > 0 Then
        sTemp3 = "Abilities"
        If Len(sAbil) = 0 And bFlag1 Then
            sTemp3 = sTemp3 & " [none]: " & sTemp1
        Else
            sTemp3 = sTemp3 & ": " & sTemp1
        End If
        sCompareText1 = AutoAppend(sCompareText1, sTemp3, " -- ")
    End If

    If Len(sTemp2) > 0 Then
        sTemp3 = "Abilities"
        If Len(sAbil) = 0 And bFlag2 Then
            sTemp3 = sTemp3 & " [none]: " & sTemp2
        Else
            sTemp3 = sTemp3 & ": " & sTemp2
        End If
        sCompareText2 = AutoAppend(sCompareText2, sTemp3, " -- ")
    End If
    
    If Len(sCastSp1) > 0 Then
        sTemp3 = "Casts"
        If Len(sCasts) = 0 And bCastSpFlag(1) Then
            sTemp3 = sTemp3 & " [none]: " & sCastSp1
        Else
            sTemp3 = sTemp3 & ": " & sCastSp1
        End If
        sCompareText1 = AutoAppend(sCompareText1, sTemp3, " -- ")
    End If

    If Len(sCastSp2) > 0 Then
        sTemp3 = "Casts"
        If Len(sCasts) = 0 And bCastSpFlag(2) Then
            sTemp3 = sTemp3 & " [none]: " & sCastSp2
        Else
            sTemp3 = sTemp3 & ": " & sCastSp2
        End If
        sCompareText2 = AutoAppend(sCompareText2, sTemp3, " -- ")
    End If
    
    If Len(sClassOk1) > 0 Then
        sTemp3 = "ClassOk"
        sTemp3 = sTemp3 & ": " & sClassOk1
        sCompareText1 = AutoAppend(sCompareText1, sTemp3, " -- ")
    End If

    If Len(sClassOk2) > 0 Then
        sTemp3 = "ClassOk"
        sTemp3 = sTemp3 & ": " & sClassOk2
        sCompareText2 = AutoAppend(sCompareText2, sTemp3, " -- ")
    End If
End If

'#################

For x = 0 To 9
    If tabItems.Fields("ClassRest-" & x) <> 0 Then
        sClasses = AutoAppend(sClasses, GetClassName(tabItems.Fields("ClassRest-" & x)))
        nClassRestrictions(0, x) = tabItems.Fields("ClassRest-" & x)
    End If
    If nInvenSlot1 >= 0 Then
        If tabItems1.Fields("ClassRest-" & x) <> 0 Then nClassRestrictions(1, x) = tabItems1.Fields("ClassRest-" & x)
    End If
    If nInvenSlot2 >= 0 Then
        If tabItems2.Fields("ClassRest-" & x) <> 0 Then nClassRestrictions(2, x) = tabItems2.Fields("ClassRest-" & x)
    End If
Next

If nInvenSlot1 >= 0 Then
    sTemp1 = ""
    sTemp2 = ""
    bFlag1 = False
    bFlag2 = False
    For y = 0 To 2
        For x = 0 To 9
            If nClassRestrictions(y, x) > 0 Then
                If y = 0 Then
                    If Not in_array_long_md(nClassRestrictions, nClassRestrictions(y, x), 1) Then
                        sTemp1 = AutoAppend(sTemp1, "+" & GetClassName(nClassRestrictions(y, x)))
                    End If
                    If nInvenSlot2 >= 0 Then
                        If Not in_array_long_md(nClassRestrictions, nClassRestrictions(y, x), 2) Then
                            sTemp2 = AutoAppend(sTemp2, "+" & GetClassName(nClassRestrictions(y, x)))
                        End If
                    End If
                Else
                    If y = 1 Then 'mark that there are values found
                        bFlag1 = True
                    Else
                        bFlag2 = True
                    End If
                    If Not in_array_long_md(nClassRestrictions, nClassRestrictions(y, x), 0) Then
                        If y = 1 Then
                            sTemp1 = AutoAppend(sTemp1, "-" & GetClassName(nClassRestrictions(y, x)))
                        Else
                            sTemp2 = AutoAppend(sTemp2, "-" & GetClassName(nClassRestrictions(y, x)))
                        End If
                    End If
                End If
            End If
        Next x
    Next y
    
    If Len(sTemp1) > 0 Then
        sTemp3 = "Classes"
        If Len(sClasses) = 0 Then
            sTemp3 = sTemp3 & " [not-restricted]"
        ElseIf Not bFlag1 Then
            sTemp3 = sTemp3 & " [+restricted]" & ": " & sTemp1
        Else
            sTemp3 = sTemp3 & ": " & sTemp1
        End If
        sCompareText1 = AutoAppend(sCompareText1, sTemp3, " -- ")
    End If

    If Len(sTemp2) > 0 Then
        sTemp3 = "Classes"
        If Len(sClasses) = 0 Then
            sTemp3 = sTemp3 & " [not-restricted]"
        ElseIf Not bFlag2 Then
            sTemp3 = sTemp3 & " [+restricted]" & ": " & sTemp2
        Else
            sTemp3 = sTemp3 & ": " & sTemp2
        End If
        sCompareText2 = AutoAppend(sCompareText2, sTemp3, " -- ")
    End If
End If


'#################

For x = 0 To 9
    If tabItems.Fields("RaceRest-" & x) <> 0 Then
        sRaces = AutoAppend(sRaces, GetRaceName(tabItems.Fields("RaceRest-" & x)))
        nRaceRestrictions(0, x) = tabItems.Fields("RaceRest-" & x)
    End If
    If nInvenSlot1 >= 0 Then
        If tabItems1.Fields("RaceRest-" & x) <> 0 Then nRaceRestrictions(1, x) = tabItems1.Fields("RaceRest-" & x)
    End If
    If nInvenSlot2 >= 0 Then
        If tabItems2.Fields("RaceRest-" & x) <> 0 Then nRaceRestrictions(2, x) = tabItems2.Fields("RaceRest-" & x)
    End If
Next

If nInvenSlot1 >= 0 Then
    sTemp1 = ""
    sTemp2 = ""
    bFlag1 = False
    bFlag2 = False
    For y = 0 To 2
        For x = 0 To 9
            If nRaceRestrictions(y, x) > 0 Then
                If y = 0 Then
                    If Not in_array_long_md(nRaceRestrictions, nRaceRestrictions(y, x), 1) Then
                        sTemp1 = AutoAppend(sTemp1, "+" & GetRaceName(nRaceRestrictions(y, x)))
                    End If
                    If nInvenSlot2 >= 0 Then
                        If Not in_array_long_md(nRaceRestrictions, nRaceRestrictions(y, x), 2) Then
                            sTemp2 = AutoAppend(sTemp2, "+" & GetRaceName(nRaceRestrictions(y, x)))
                        End If
                    End If
                Else
                    If y = 1 Then 'mark that there are values found
                        bFlag1 = True
                    Else
                        bFlag2 = True
                    End If
                    If Not in_array_long_md(nRaceRestrictions, nRaceRestrictions(y, x), 0) Then
                        If y = 1 Then
                            sTemp1 = AutoAppend(sTemp1, "-" & GetRaceName(nRaceRestrictions(y, x)))
                        Else
                            sTemp2 = AutoAppend(sTemp2, "-" & GetRaceName(nRaceRestrictions(y, x)))
                        End If
                    End If
                End If
            End If
        Next x
    Next y
    
    If Len(sTemp1) > 0 Then
        sTemp3 = "Races"
        If Len(sRaces) = 0 Then
            sTemp3 = sTemp3 & " [not-restricted]"
        ElseIf Not bFlag1 Then
            sTemp3 = sTemp3 & " [+restricted]" & ": " & sTemp1
        Else
            sTemp3 = sTemp3 & ": " & sTemp1
        End If
        sCompareText1 = AutoAppend(sCompareText1, sTemp3, " -- ")
    End If

    If Len(sTemp2) > 0 Then
        sTemp3 = "Races"
        If Len(sRaces) = 0 Then
            sTemp3 = sTemp3 & " [not-restricted]"
        ElseIf Not bFlag2 Then
            sTemp3 = sTemp3 & " [+restricted]" & ": " & sTemp2
        Else
            sTemp3 = sTemp3 & ": " & sTemp2
        End If
        sCompareText2 = AutoAppend(sCompareText2, sTemp3, " -- ")
    End If
End If

'#################

For x = 0 To 9
    If tabItems.Fields("NegateSpell-" & x) <> 0 Then
        sNegate = AutoAppend(sNegate, GetSpellName(tabItems.Fields("NegateSpell-" & x)))
        nNegateSpells(0, x) = tabItems.Fields("NegateSpell-" & x)
    End If
    If nInvenSlot1 >= 0 Then
        If tabItems1.Fields("NegateSpell-" & x) <> 0 Then nNegateSpells(1, x) = tabItems1.Fields("NegateSpell-" & x)
    End If
    If nInvenSlot2 >= 0 Then
        If tabItems2.Fields("NegateSpell-" & x) <> 0 Then nNegateSpells(2, x) = tabItems2.Fields("NegateSpell-" & x)
    End If
Next

If nInvenSlot1 >= 0 Then
    sTemp1 = ""
    sTemp2 = ""
    bFlag1 = False
    bFlag2 = False
    For y = 0 To 2
        For x = 0 To 9
            If nNegateSpells(y, x) > 0 Then
                If y = 0 Then
                    If Not in_array_long_md(nNegateSpells, nNegateSpells(y, x), 1) Then
                        sTemp1 = AutoAppend(sTemp1, "+" & GetSpellName(nNegateSpells(y, x)))
                    End If
                    If nInvenSlot2 >= 0 Then
                        If Not in_array_long_md(nNegateSpells, nNegateSpells(y, x), 2) Then
                            sTemp2 = AutoAppend(sTemp2, "+" & GetSpellName(nNegateSpells(y, x)))
                        End If
                    End If
                Else
                    If y = 1 Then 'mark that there are values found
                        bFlag1 = True
                    Else
                        bFlag2 = True
                    End If
                    If Not in_array_long_md(nNegateSpells, nNegateSpells(y, x), 0) Then
                        If y = 1 Then
                            sTemp1 = AutoAppend(sTemp1, "-" & GetSpellName(nNegateSpells(y, x)))
                        Else
                            sTemp2 = AutoAppend(sTemp2, "-" & GetSpellName(nNegateSpells(y, x)))
                        End If
                    End If
                End If
            End If
        Next x
    Next y
    
    If Len(sTemp1) > 0 Then
        sTemp3 = "Negate"
        If Len(sNegate) = 0 And bFlag1 Then
            sTemp3 = sTemp3 & " [none]: " & sTemp1
        Else
            sTemp3 = sTemp3 & ": " & sTemp1
        End If
        sCompareText1 = AutoAppend(sCompareText1, sTemp3, " -- ")
    End If

    If Len(sTemp2) > 0 Then
        sTemp3 = "Negate"
        If Len(sNegate) = 0 And bFlag2 Then
            sTemp3 = sTemp3 & " [none]: " & sTemp2
        Else
            sTemp3 = sTemp3 & ": " & sTemp2
        End If
        sCompareText2 = AutoAppend(sCompareText2, sTemp3, " -- ")
    End If
    
End If

'#################

Call GetLocations(tabItems.Fields("Obtained From"), LocationLV, , , nNumber, , , True)
If nNMRVer >= 1.7 Then Call GetLocations(tabItems.Fields("References"), LocationLV, True, , , , , True)

If Not tabItems.Fields("Number") = nNumber Then
    tabItems.Index = "pkItems"
    tabItems.Seek "=", nNumber
End If

For x = 0 To 19
    If nAbils(0, x, 0) > 0 Then
        Select Case nAbils(0, x, 0)
            Case 43: 'casts spell
                'this is just so it adds any references to the locationlv after being cleared from the above GetLocations
                sTemp1 = PullSpellEQ(True, 0, nAbils(0, x, 1), LocationLV, , , True)
                
                Set oLI = LocationLV.ListItems.Add
                oLI.Text = ""
                oLI.ListSubItems.Add 1, , "Casts: " & GetSpellName(nAbils(0, x, 1), bHideRecordNumbers)
                oLI.ListSubItems(1).Tag = nAbils(0, x, 1)
                
            Case 114: '%spell
                nPercent = nAbils(0, x, 1)
                
        End Select
    End If
Next x

If Not tabItems.Fields("Number") = nNumber Then
    tabItems.Index = "pkItems"
    tabItems.Seek "=", nNumber
End If

'####################

If Not sGetDrop = "" Then
    sStr = AutoAppend(sStr, sGetDrop, " -- ")
End If
If Not sUses = "" Then
    sStr = AutoAppend(sStr, "Uses: " & sUses, " -- ")
End If
If Not sAbil = "" Then
    sStr = AutoAppend(sStr, "Abilities: " & sAbil, " -- ")
End If
If Not sCasts = "" Then
    sStr = AutoAppend(sStr, "Casts: " & sCasts, " -- ")
End If
If Not sClassOk = "" Then
    sStr = AutoAppend(sStr, "ClassOK: " & sClassOk, " -- ")
End If
If Not sClasses = "" Then
    sStr = AutoAppend(sStr, "Classes: " & sClasses, " -- ")
End If
If Not sRaces = "" Then
    sStr = AutoAppend(sStr, "Races: " & sRaces, " -- ")
End If
If Not sNegate = "" Then
    sStr = AutoAppend(sStr, "Negates: " & sNegate, " -- ")
End If

If nInvenSlot1 >= 0 Then
    If bCompareWeapon Then sCompareText1 = AutoPrepend(sCompareText1, CompareWeapons(tabItems, tabItems1))
    If bCompareArmor Then sCompareText1 = AutoPrepend(sCompareText1, CompareArmor(tabItems, tabItems1))
    sStr = AutoAppend(sStr, "Compared to " & tabItems1.Fields("Name") & ": " & sCompareText1, vbCrLf & vbCrLf)
End If
If nInvenSlot2 >= 0 Then
    If bCompareWeapon Then sCompareText2 = AutoPrepend(sCompareText2, CompareWeapons(tabItems, tabItems2))
    If bCompareArmor Then sCompareText2 = AutoPrepend(sCompareText2, CompareArmor(tabItems, tabItems2))
    sStr = AutoAppend(sStr, "Compared to " & tabItems2.Fields("Name") & ": " & sCompareText2, vbCrLf & vbCrLf)
End If

'weapon damage
If tabItems.Fields("ItemType") = 1 Then
    
    If nAttackTypeMUD = a4_Surprise And bUseCharacter Then
        If GetClassStealth = False And GetRaceStealth = False Then bForceCalc = True
    End If
    
    'Call PopulateCharacterProfile(tCharacter, bUseCharacter, True, nAttackTypeMUD, nNumber)
    If Not tabItems.Fields("Number") = nNumber Then tabItems.Seek "=", nNumber
    
    tWeaponDmg = CalculateAttack( _
                    tCharacter, _
                    nAttackTypeMUD, _
                    tabItems.Fields("Number"), _
                    False, _
                    nSpeedAdj, _
                    IIf(bCalcCombat, val(frmMain.txtWeaponExtras(2).Text), 0), _
                    IIf(bCalcCombat, val(frmMain.txtWeaponExtras(3).Text), 0), _
                    IIf(bCalcCombat, val(frmMain.txtWeaponExtras(4).Text), 0), _
                    sCasts, bForceCalc)
    
    If Not tabItems.Fields("Number") = nNumber Then tabItems.Seek "=", nNumber
    
    If tWeaponDmg.nSwings > 0 Then
        Select Case nAttackTypeMUD
            Case 1: sWeaponDmg = "Punch Damage"
            Case 2: sWeaponDmg = "Kick Damage"
            Case 3: sWeaponDmg = "Jumpkick Damage"
            Case 4: sWeaponDmg = "Backstab Damage"
            Case 6: sWeaponDmg = "Bash Damage"
            Case 7: sWeaponDmg = "Smash Damage"
            Case Else:
                sWeaponDmg = "Damage"
        End Select
        If bUseCharacter = False Then sWeaponDmg = sWeaponDmg & " (@lvl 255)"
        sWeaponDmg = sWeaponDmg & ": "
        sWeaponDmg = sWeaponDmg & tWeaponDmg.nRoundTotal & "/round @ " & Round(tWeaponDmg.nSwings, 1) & " swings w/" & tWeaponDmg.nHitChance & "% hit chance"
        If nAttackTypeMUD = 4 And bForceCalc Then sWeaponDmg = sWeaponDmg & " (forced race stealth)"
        
        sWeaponDmg = sWeaponDmg & " - Avg Hit: " & tWeaponDmg.nAvgHit
        
        If tWeaponDmg.nMaxCrit > 0 And tWeaponDmg.nCritChance > 0 Then
            sWeaponDmg = AutoAppend(sWeaponDmg, "Avg/Max Crit: " & tWeaponDmg.nAvgCrit & "/" & tWeaponDmg.nMaxCrit)
            sWeaponDmg = sWeaponDmg & " (" & tWeaponDmg.nCritChance & "%"
            If tWeaponDmg.nQnDBonus > 0 Then sWeaponDmg = sWeaponDmg & " w/" & tWeaponDmg.nQnDBonus & "qnd"
            sWeaponDmg = sWeaponDmg & ")"
        End If
        
        If tWeaponDmg.nAvgExtraHit > 0 Then
            sWeaponDmg = AutoAppend(sWeaponDmg, "Avg Extra: " & tWeaponDmg.nAvgExtraHit)
            If tWeaponDmg.nAvgExtraHit <> tWeaponDmg.nAvgExtraSwing Then
                sWeaponDmg = sWeaponDmg & " (avg " & tWeaponDmg.nAvgExtraSwing & "/swing)"
            End If
        End If
        
        If tWeaponDmg.nRoundTotal > 0 And tWeaponDmg.nRoundTotal <> tWeaponDmg.nFirstRoundDamage And tWeaponDmg.nSwings <> Fix(tWeaponDmg.nSwings) Then
            sWeaponDmg = sWeaponDmg & " - 1st Round: " & tWeaponDmg.nFirstRoundDamage & " dmg @ " & Fix(tWeaponDmg.nSwings) & " swings"
        End If
        
        If bFullDetails Then
            sTemp1 = "Min/Max Hit: " & tWeaponDmg.nMinDmg & "/" & tWeaponDmg.nMaxDmg
            sTemp1 = sTemp1 & ", Speed: " & tabItems.Fields("Speed")
            If nLVLreq > 0 Then sTemp1 = sTemp1 & ", LVL Req: " & nLVLreq
            If tabItems.Fields("StrReq") > 0 Then sTemp1 = sTemp1 & ", STR Req: " & tabItems.Fields("StrReq")
            If (tabItems.Fields("ArmourClass") + tabItems.Fields("DamageResist")) > 0 Then
                sTemp1 = sTemp1 & ", AC/DR: " & RoundUp(tabItems.Fields("ArmourClass") / 10) & "/" & (tabItems.Fields("DamageResist") / 10)
            End If
            If nAcc > 0 Then sTemp1 = sTemp1 & ", Accy: " & nAcc
            sTemp1 = sTemp1 & ", BS: " & IIf(nBSacc > 0, nBSacc, "No")
            If tabItems.Fields("Limit") > 0 Then sTemp1 = sTemp1 & ", Limit: " & tabItems.Fields("Limit")
            sWeaponDmg = sTemp1 & vbCrLf & sWeaponDmg
        End If
        
        sWeaponDmg = sWeaponDmg & vbCrLf
        
        If bUseCharacter And tabItems.Fields("StrReq") > val(frmMain.txtCharStats(0).Tag) Then
            sWeaponDmg = sWeaponDmg & "Notice: Character Strength (" & val(frmMain.txtCharStats(0).Tag) & ") < Strength Requirement (" & tabItems.Fields("StrReq") & ")" & vbCrLf
        End If
        
        sWeaponDmg = sWeaponDmg & vbCrLf
    End If
ElseIf tabItems.Fields("ItemType") = 0 And bFullDetails Then
    sTemp1 = "Armour Type: " & GetArmourType(tabItems.Fields("ArmourType"))
    If nLVLreq > 0 Then sTemp1 = sTemp1 & ", LVL Req: " & nLVLreq
    If (tabItems.Fields("ArmourClass") + tabItems.Fields("DamageResist")) > 0 Then
        sTemp1 = sTemp1 & ", AC/DR: " & RoundUp(tabItems.Fields("ArmourClass") / 10) & "/" & (tabItems.Fields("DamageResist") / 10)
    End If
    If nAcc > 0 Then sTemp1 = sTemp1 & ", Accy: " & nAcc
    If tabItems.Fields("Limit") > 0 Then sTemp1 = sTemp1 & ", Limit: " & tabItems.Fields("Limit")
    sWeaponDmg = sTemp1 & vbCrLf & vbCrLf
End If

If Not tabItems.Fields("Number") = nNumber Then
    tabItems.Index = "pkItems"
    tabItems.Seek "=", nNumber
End If

DetailTB.Text = sWeaponDmg & sStr

If LocationLV.ListItems.Count > 0 Then
'    If nLastItemSortCol > LocationLV.ColumnHeaders.Count Then nLastItemSortCol = 1
'    If nLastItemSortCol = 1 Then
'        Call SortListViewByTag(LocationLV, 1, ldtnumber, False)
'    Else
'        Call SortListView(LocationLV, nLastItemSortCol, ldtstring, False)
'    End If
    
    Call LV_RefreshSort(LocationLV, 1, ldtnumber, True, False)
End If

out:
On Error Resume Next
Set oLI = Nothing
If Not tabItems2 Is Nothing Then tabItems2.Close
Set tabItems2 = Nothing
If Not tabItems1 Is Nothing Then tabItems1.Close
Set tabItems1 = Nothing
Exit Sub

error:
Call HandleError("PullItemDetail")
Resume out:
End Sub

Private Function CompareArmor(tabItem1 As Recordset, tabItem2 As Recordset) As String
Dim nTemp As Long, nTemp2 As Double
On Error GoTo error:

If tabItem1.Fields("Worn") <> tabItem2.Fields("Worn") Then
    CompareArmor = AutoAppend(CompareArmor, GetWornType(tabItem2.Fields("Worn")) & " -> " & GetWornType(tabItem1.Fields("Worn")))
End If

If tabItem1.Fields("ArmourType") <> tabItem2.Fields("ArmourType") Then
    CompareArmor = AutoAppend(CompareArmor, GetArmourType(tabItem2.Fields("ArmourType")) & " -> " & GetArmourType(tabItem1.Fields("ArmourType")))
End If

If tabItem1.Fields("Encum") <> tabItem2.Fields("Encum") Then
    nTemp = tabItem1.Fields("Encum") - tabItem2.Fields("Encum")
    If nTemp <> 0 Then CompareArmor = AutoAppend(CompareArmor, "Encum: " & IIf(nTemp > 0, "+", "") & nTemp)
End If

If tabItem1.Fields("Accy") <> tabItem2.Fields("Accy") Then
    nTemp = tabItem1.Fields("Accy") - tabItem2.Fields("Accy")
    If nTemp <> 0 Then CompareArmor = AutoAppend(CompareArmor, "WepAccy: " & IIf(nTemp > 0, "+", "") & nTemp)
End If

If tabItem1.Fields("ArmourClass") <> tabItem2.Fields("ArmourClass") Or tabItem1.Fields("DamageResist") <> tabItem2.Fields("DamageResist") Then
    nTemp = RoundUp(tabItem1.Fields("ArmourClass") / 10) - RoundUp(tabItem2.Fields("ArmourClass") / 10)
    nTemp2 = (tabItem1.Fields("DamageResist") / 10) - (tabItem2.Fields("DamageResist") / 10)
    If nTemp <> 0 Or nTemp2 <> 0 Then CompareArmor = AutoAppend(CompareArmor, "AC: " & IIf(nTemp > 0, "+", "") & nTemp & "/" & IIf(nTemp2 > 0, "+", "") & nTemp2)
End If

'If tabItem1.Fields("Encum") <> tabItem2.Fields("Encum") Or tabItem1.Fields("ArmourClass") <> tabItem2.Fields("ArmourClass") Or tabItem1.Fields("DamageResist") <> tabItem2.Fields("DamageResist") Then
'    nTemp = Get_Enc_Ratio(tabItem1.Fields("Encum"), tabItem1.Fields("ArmourClass"), tabItem1.Fields("DamageResist")) - Get_Enc_Ratio(tabItem2.Fields("Encum"), tabItem2.Fields("ArmourClass"), tabItem2.Fields("DamageResist"))
'    If nTemp <> 0 Then CompareArmor = AutoAppend(CompareArmor, "AC/Enc: " & IIf(nTemp > 0, "+", "") & nTemp)
'End If

If tabItem1.Fields("Limit") <> tabItem2.Fields("Limit") Then
    nTemp = tabItem1.Fields("Limit") - tabItem2.Fields("Limit")
    If nTemp <> 0 Then CompareArmor = AutoAppend(CompareArmor, "Limit: " & IIf(nTemp > 0, "+", "") & nTemp)
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CompareArmor")
Resume out:
End Function


Private Function CompareWeapons(tabItem1 As Recordset, tabItem2 As Recordset) As String
Dim nTemp As Long, nTemp2 As Double
On Error GoTo error:

If tabItem1.Fields("WeaponType") <> tabItem2.Fields("WeaponType") Then
    CompareWeapons = AutoAppend(CompareWeapons, GetWeaponType(tabItem2.Fields("WeaponType")) & " -> " & GetWeaponType(tabItem1.Fields("WeaponType")))
End If

If tabItem1.Fields("Min") <> tabItem2.Fields("Min") Or tabItem1.Fields("Max") <> tabItem2.Fields("Max") Then
    nTemp = Round(tabItem1.Fields("Min") + tabItem1.Fields("Max") / 2, 0) - Round(tabItem2.Fields("Min") + tabItem2.Fields("Max") / 2, 0)
    If nTemp <> 0 Then CompareWeapons = AutoAppend(CompareWeapons, "Damage: " & IIf(nTemp > 0, "+", "") & nTemp)
End If

If tabItem1.Fields("Speed") <> tabItem2.Fields("Speed") Then
    nTemp = tabItem1.Fields("Speed") - tabItem2.Fields("Speed")
    If nTemp <> 0 Then CompareWeapons = AutoAppend(CompareWeapons, "Speed: " & IIf(nTemp > 0, "+", "") & nTemp)
End If

If tabItem1.Fields("Encum") <> tabItem2.Fields("Encum") Then
    nTemp = tabItem1.Fields("Encum") - tabItem2.Fields("Encum")
    If nTemp <> 0 Then CompareWeapons = AutoAppend(CompareWeapons, "Encum: " & IIf(nTemp > 0, "+", "") & nTemp)
End If

If tabItem1.Fields("Accy") <> tabItem2.Fields("Accy") Then
    nTemp = tabItem1.Fields("Accy") - tabItem2.Fields("Accy")
    If nTemp <> 0 Then CompareWeapons = AutoAppend(CompareWeapons, "WepAccy: " & IIf(nTemp > 0, "+", "") & nTemp)
End If

If tabItem1.Fields("ArmourClass") <> tabItem2.Fields("ArmourClass") Or tabItem1.Fields("DamageResist") <> tabItem2.Fields("DamageResist") Then
    nTemp = RoundUp(tabItem1.Fields("ArmourClass") / 10) - RoundUp(tabItem2.Fields("ArmourClass") / 10)
    nTemp2 = (tabItem1.Fields("DamageResist") / 10) - (tabItem2.Fields("DamageResist") / 10)
    If nTemp <> 0 Or nTemp2 <> 0 Then CompareWeapons = AutoAppend(CompareWeapons, "AC: " & IIf(nTemp > 0, "+", "") & nTemp & "/" & IIf(nTemp2 > 0, "+", "") & nTemp2)
End If

If tabItem1.Fields("StrReq") <> tabItem2.Fields("StrReq") Then
    nTemp = tabItem1.Fields("StrReq") - tabItem2.Fields("StrReq")
    If nTemp <> 0 Then CompareWeapons = AutoAppend(CompareWeapons, "StrReq: " & IIf(nTemp > 0, "+", "") & nTemp & " (" & tabItem1.Fields("StrReq") & ")")
End If

If tabItem1.Fields("Limit") <> tabItem2.Fields("Limit") Then
    nTemp = tabItem1.Fields("Limit") - tabItem2.Fields("Limit")
    If nTemp <> 0 Then CompareWeapons = AutoAppend(CompareWeapons, "Limit: " & IIf(nTemp > 0, "+", "") & nTemp & " (" & tabItem1.Fields("Limit") & ")")
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CompareWeapons")
Resume out:
End Function

Public Sub PullClassDetail(nClassNum As Long, DetailTB As TextBox)

On Error GoTo error:

Dim sAbil As String, x As Integer

DetailTB.Text = ""
If bStartup Then Exit Sub

tabClasses.Index = "pkClasses"
tabClasses.Seek "=", nClassNum
If tabClasses.NoMatch = True Then
    DetailTB.Text = "Class not found"
    tabClasses.MoveFirst
    Exit Sub
End If

'sStr = ClipNull(tabClasses.Fields("Name")) & " (" & tabClasses.Fields("Number") & ")"

'sStr = "Exp: " & (Val(tabClasses.Fields("ExpTable")) + 100) & "%"
'sStr = sStr & " -- Weapon: " & GetClassWeaponType(tabClasses.Fields("WeaponType"))
'sStr = sStr & " -- Armour: " & GetArmourType(tabClasses.Fields("ArmourType"))
'sStr = sStr & " -- Magic: " & GetMagery(tabClasses.Fields("MageryType"), tabClasses.Fields("MageryLVL"))
'sStr = sStr & " -- Combat: " & (tabClasses.Fields("CombatLVL") - 2)
'sStr = sStr & " -- HP: " & tabClasses.Fields("MinHits") & "-" & (tabClasses.Fields("MinHits") + tabClasses.Fields("MaxHits"))

For x = 0 To 9
    If Not tabClasses.Fields("Abil-" & x) = 0 Then
        Select Case tabClasses.Fields("Abil-" & x)
            Case 0:
            Case 59: 'classok
            Case Else:
                If sAbil <> "" Then sAbil = sAbil & ", "
                sAbil = sAbil & GetAbilityStats(tabClasses.Fields("Abil-" & x), tabClasses.Fields("AbilVal-" & x))
                If Right(sAbil, 2) = ", " Then sAbil = Left(sAbil, Len(sAbil) - 2)
                If tabClasses.Fields("Number") <> nClassNum Then tabClasses.Seek "=", nClassNum
        End Select
    End If
Next

DetailTB.Text = sAbil

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("PullClassDetail")
Resume out:
End Sub
Public Sub PullRaceDetail(nRaceNum As Long, DetailTB As TextBox)

On Error GoTo error:

Dim sAbil As String, x As Integer

DetailTB.Text = ""
If bStartup Then Exit Sub

tabRaces.Index = "pkRaces"
tabRaces.Seek "=", nRaceNum
If tabRaces.NoMatch = True Then
    DetailTB.Text = "Race not found"
    tabRaces.MoveFirst
    Exit Sub
End If

'sStr = ClipNull(tabRaces.Fields("Name")) & " (" & tabRaces.Fields("Number") & ")"
'sStr = "Exp: " & tabRaces.Fields("ExpTable") & "%, HP Bonus: " & tabRaces.Fields("HPPerLVL")
'
'sStr = sStr & vbCrLf _
'    & "Str: " & tabRaces.Fields("mSTR") & "-" & tabRaces.Fields("xSTR") & " ... " & "Agi: " & tabRaces.Fields("mAGL") & "-" & tabRaces.Fields("xAGL") & vbCrLf _
'    & "Int: " & tabRaces.Fields("mINT") & "-" & tabRaces.Fields("xINT") & " ... " & "Hea: " & tabRaces.Fields("mHEA") & "-" & tabRaces.Fields("xHEA") & vbCrLf _
'    & "Wis: " & tabRaces.Fields("mWIL") & "-" & tabRaces.Fields("xWIL") & " ... " & "Cha: " & tabRaces.Fields("mCHM") & "-" & tabRaces.Fields("xCHM")

For x = 0 To 9
    If Not tabRaces.Fields("Abil-" & x) = 0 Then
        Select Case tabRaces.Fields("Abil-" & x)
            Case 0:
            Case Else:
                If sAbil <> "" Then sAbil = sAbil & ", "
                sAbil = sAbil & GetAbilityStats(tabRaces.Fields("Abil-" & x), tabRaces.Fields("AbilVal-" & x))
                If Right(sAbil, 2) = ", " Then sAbil = Left(sAbil, Len(sAbil) - 2)
                If tabRaces.Fields("Number") <> nRaceNum Then tabRaces.Seek "=", nRaceNum
        End Select
    End If
Next

DetailTB.Text = sAbil



out:
On Error Resume Next
Exit Sub
error:
Call HandleError("PullRaceDetail")
Resume out:
End Sub
Public Sub PullMonsterDetail(nMonsterNum As Long, DetailLV As ListView, Optional ByVal nLookupLimit = 100)
Dim sAbil As String, x As Integer, y As Integer, sTemp As String, sTemp2 As String, sExpEa As String
Dim sCash As String, nCash As Currency, nPercent As Integer, nTemp As Long, tExpInfo As tExpPerHourInfo
Dim oLI As ListItem, nExp As Currency, nLocalMonsterDamage As MonAttackSimReturn, nMonsterEnergy As Long
Dim sReducedCoin As String, nReducedCoin As Currency, nDamage As Currency, nTemp2 As Long
Dim nAvgDmg As Long, nExpDmgHP As Currency, nExpPerHour As Currency, nExpPerHourEA As Currency, nPossyPCT As Currency, nMobDodge As Integer
Dim nScriptValue As Currency, nLairPCT As Currency, nPossSpawns As Long, sPossSpawns As String, sScriptValue As String
Dim tAvgLairInfo As LairInfoType, sArr() As String, bHasAttacks As Boolean, bNeedBlankRow As Boolean, nMobDmg As Long
Dim nDamageOut As Long, sDefenseDesc As String, nDamageVMob As Currency, tBackStab As tAttackDamage
Dim nMaxLairsBeforeRegen As Currency, bHasAntiMagic As Boolean, tAttack As tAttackDamage, sHeader As String
Dim tSpellcast As tSpellCastValues, nCalcDamageHP As Long, nSurpriseDamageOut As Long, nMinDamageOut As Long
Dim nCalcDamageAC As Long, nCalcDamageDR As Long, nCalcDamageDodge As Long, nCalcDamageMR As Long, nCalcDamageHPRegen As Long
Dim nCalcDamageNumMobs As Currency, bUseCharacter As Boolean, iAttack As Integer
Dim tCharProfile As tCharacterProfile, tForcedCharProfile As tCharacterProfile, tBackStabProfile As tCharacterProfile
Dim nDmgOut() As Currency, nSpeedAdj As Integer, sBackstabText As String, nOverrideRTK As Double, sImmuTXT As String
Dim nSpellImmuLVL As Integer, nWeaponMagic As Integer, nBackstabWeaponMagic As Integer, nMagicLVL As Integer
Dim nCalcSpellImmuLVL As Integer, nCalcMagicLVL As Integer, nBSDefense As Integer
Dim nMobElementalResist(5) As Integer, nCalcElementalResist(5) As Integer
'Dim bIsLiving As Boolean, bIsAnimal As Boolean, bIsUndead As Boolean, bIsAntiMagic As Boolean
'Dim bMobIsLiving As Boolean, bMobIsAnimal As Boolean, bMobIsUndead As Boolean, bMobIsAntiMagic As Boolean
Dim DF_Flags As eDefenseFlags, eMobDefenseFlags As eDefenseFlags, eAttackFlags As eAttackRestrictions, bValidTarget As Boolean

On Error GoTo error:

DetailLV.ListItems.clear
If bStartup Then Exit Sub

tabMonsters.Index = "pkMonsters"
tabMonsters.Seek "=", nMonsterNum
If tabMonsters.NoMatch = True Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Monster not found"
    'DetailTB.Text = "Monster not found"
    Set oLI = Nothing
    tabMonsters.MoveFirst
    Exit Sub
End If

If frmMain.chkGlobalFilter.Value = 1 Then bUseCharacter = True
nSpeedAdj = 100
eAttackFlags = AR000_Unknown

Call RefreshCombatHealingValues
Call PopulateCharacterProfile(tCharProfile)

'tForcedCharProfile == a forced character profile, e.g. either current character or a generic character, but NOT a party character
If nGlobalAttackTypeMME = a4_MartialArts And bUseCharacter Then 'MA
    'this is to get proper +skill/accy/dmg stats
    Call PopulateCharacterProfile(tForcedCharProfile, bUseCharacter, True, IIf(nGlobalAttackMA > 1, nGlobalAttackMA, 1))
ElseIf tCharProfile.nParty > 1 And bUseCharacter Then
    'in a party, tCharProfile set earlier will have party stats
    'some of the the attack calculations in pull monster detail are [this char] vs [mob]
    Call PopulateCharacterProfile(tForcedCharProfile, bUseCharacter, True)
Else
    tForcedCharProfile = tCharProfile 'this means that the char profile is the normal character profile
End If

If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum

Set oLI = DetailLV.ListItems.Add()
oLI.Text = "Name"
oLI.Bold = True
oLI.ListSubItems.Add (1), "Detail", tabMonsters.Fields("Name") & " (" & nMonsterNum & ")"
oLI.ListSubItems(1).Bold = True

Set oLI = DetailLV.ListItems.Add()
oLI.Text = "Experience"
If UseExpMulti Then
    nExp = tabMonsters.Fields("EXP") * tabMonsters.Fields("ExpMulti")
Else
    nExp = tabMonsters.Fields("EXP")
End If
oLI.ListSubItems.Add (1), "Detail", IIf(nExp > 0, Format(nExp, "#,#"), 0)

If Not tabMonsters.Fields("RegenTime") = 0 Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Regen Time"
    oLI.ListSubItems.Add (1), "Detail", tabMonsters.Fields("RegenTime") & " hour(s)"
End If

If nNMRVer >= 1.83 Then
    If Not tabMonsters.Fields("GameLimit") = 0 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "Game Limit"
        If tabMonsters.Fields("RegenTime") = 0 Then oLI.ForeColor = RGB(204, 0, 0)
        oLI.ListSubItems.Add (1), "Detail", tabMonsters.Fields("GameLimit")
        If tabMonsters.Fields("RegenTime") = 0 Then oLI.ListSubItems(1).ForeColor = RGB(204, 0, 0)
    End If
End If

Set oLI = DetailLV.ListItems.Add()
oLI.Text = "Type"
oLI.ListSubItems.Add (1), "Detail", GetMonType(tabMonsters.Fields("Type"))


sTemp = GetMonAlignment(tabMonsters.Fields("Align"))
'Case 0: GetMonAlignment = "Good"
'Case 1: GetMonAlignment = "Evil"
'Case 2: GetMonAlignment = "Chaotic Evil"
'Case 3: GetMonAlignment = "Neutral"
'Case 4: GetMonAlignment = "Lawful Good"
'Case 5: GetMonAlignment = "Neutral Evil"
'Case 6: GetMonAlignment = "Lawful Evil"
'does NOT attack good or neutral:   good, lawful good, neutral
'does NOT attack evil: lawful evil, good, lawful good, neutral
Select Case tabMonsters.Fields("Align")
    Case 0, 3, 4:
        sTemp = sTemp & " [Not-Hostile]"
    Case 6:
        If bUseCharacter And frmMain.cmbGlobalAlignment.ListIndex = 3 Then 'evil aligned
            sTemp = sTemp & " [Not-Hostile to EVIL]"
        Else
            sTemp = sTemp & " [Hostile]"
        End If
    Case Else:
        sTemp = sTemp & " [Hostile]"
End Select
Set oLI = DetailLV.ListItems.Add()
oLI.Text = "Alignment"
oLI.ListSubItems.Add (1), "Detail", sTemp

If tabMonsters.Fields("Undead") = 1 Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Undead"
    oLI.ListSubItems.Add (1), "Detail", "Yes"
    eMobDefenseFlags = eMobDefenseFlags Or DF023_IsUndead
End If

Set oLI = DetailLV.ListItems.Add()
oLI.Text = "HPs"
'oLI.ForeColor = RGB(204, 0, 0)
oLI.ListSubItems.Add (1), "Detail", tabMonsters.Fields("HP") & " (Regens: " & tabMonsters.Fields("HPRegen") & " HPs every 90 seconds [18 rounds])"
'oLI.ListSubItems(1).ForeColor = RGB(204, 0, 0)

Set oLI = DetailLV.ListItems.Add()
oLI.Text = "AC/DR"
oLI.ListSubItems.Add (1), "Detail", tabMonsters.Fields("ArmourClass") & "/" & tabMonsters.Fields("DamageResist")

Set oLI = DetailLV.ListItems.Add()
oLI.Text = "MR"
oLI.ListSubItems.Add (1), "Detail", tabMonsters.Fields("MagicRes")

Set oLI = DetailLV.ListItems.Add()
oLI.Text = "Follow %"
oLI.ListSubItems.Add (1), "Detail", tabMonsters.Fields("Follow%")

Set oLI = DetailLV.ListItems.Add()
oLI.Text = "Charm LVL"
oLI.ListSubItems.Add (1), "Detail", tabMonsters.Fields("CharmLVL")

'cash
nCash = nCash + (tabMonsters.Fields("R") * 1000000)
nCash = nCash + (tabMonsters.Fields("P") * 10000)
nCash = nCash + (tabMonsters.Fields("G") * 100)
nCash = nCash + (tabMonsters.Fields("S") * 10)
nCash = nCash + tabMonsters.Fields("C")

sReducedCoin = "Copper"
nReducedCoin = nCash
If nReducedCoin > 0 Then
    If nCash >= 10000000 Then
        nReducedCoin = nCash / 1000000
        sReducedCoin = "Runic"
    ElseIf nCash >= 100000 Then
        nReducedCoin = nCash / 10000
        sReducedCoin = "Platinum"
    ElseIf nCash >= 1000 Then
        nReducedCoin = nCash / 100
        sReducedCoin = "Gold"
    ElseIf nCash >= 100 Then
        nReducedCoin = nCash / 10
        sReducedCoin = "Silver"
    End If
    
    If Not sReducedCoin = "Copper" Then
        nReducedCoin = Round(nReducedCoin, 2)
    End If
    
    sCash = Format(nReducedCoin, "##,##0.00")
    If Right(sCash, 3) = ".00" Then sCash = Left(sCash, Len(sCash) - 3)

    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Cash (up to)"
    oLI.ListSubItems.Add (1), "Detail", sCash & " " & sReducedCoin
End If

If Not tabMonsters.Fields("Weapon") = 0 Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Weapon"
    oLI.Tag = "Item"
    oLI.ListSubItems.Add (1), "Detail", GetItemName(tabMonsters.Fields("Weapon"), bHideRecordNumbers)
    oLI.ListSubItems(1).Tag = tabMonsters.Fields("Weapon")
End If

If tabMonsters.Fields("CreateSpell") > 0 Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Create Spell:"
    oLI.Tag = "Spell"
    'oLI.ForeColor = &HC00000
    'nSpellNest = 0 'make sure this doesn't nest too deep
    oLI.ListSubItems.Add (1), "Detail", "[" & GetSpellName(tabMonsters.Fields("CreateSpell"), bHideRecordNumbers) _
        & ", " & PullSpellEQ(False, , tabMonsters.Fields("CreateSpell")) & "]"
    If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
    oLI.ListSubItems(1).Tag = tabMonsters.Fields("CreateSpell")
    'oLI.ListSubItems(1).ForeColor = &HC00000
End If

If tabMonsters.Fields("DeathSpell") > 0 Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Death Spell"
    oLI.Tag = "Spell"
    'oLI.ForeColor = &HC00000
    'nSpellNest = 0 'make sure this doesn't nest too deep
    oLI.ListSubItems.Add (1), "Detail", "[" & GetSpellName(tabMonsters.Fields("DeathSpell"), bHideRecordNumbers) _
        & ", " & PullSpellEQ(False, , tabMonsters.Fields("DeathSpell")) & "]"
    If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
    oLI.ListSubItems(1).Tag = tabMonsters.Fields("DeathSpell")
    If SpellHasAbility(tabMonsters.Fields("DeathSpell"), 60) >= 0 Then 'fear
        oLI.ListSubItems(1).ForeColor = &HC0&
        oLI.ListSubItems(1).Bold = True
    ElseIf SpellHasAbility(tabMonsters.Fields("DeathSpell"), 19) >= 0 Then 'poison
        oLI.ListSubItems(1).ForeColor = &H8000&
        oLI.ListSubItems(1).Bold = True
    ElseIf SpellHasAbility(tabMonsters.Fields("DeathSpell"), 71) >= 0 Then 'confusion
        oLI.ListSubItems(1).ForeColor = &H80FF&
        oLI.ListSubItems(1).Bold = True
    End If
End If

If tabMonsters.Fields("GreetTXT") > 0 Then
    tabTBInfo.Index = "pkTBInfo"
    tabTBInfo.Seek "=", tabMonsters.Fields("GreetTXT")
    If Not tabTBInfo.NoMatch Then
        If Not LCase(Left(tabTBInfo.Fields("Action"), 4)) = "heh:" _
            And Not LCase(Left(tabTBInfo.Fields("Action"), 8)) = "nothing:" _
            And Not LCase(Left(tabTBInfo.Fields("Action"), 5)) = "yada:" _
            And Not LCase(Left(tabTBInfo.Fields("Action"), 5)) = "hehe:" _
            And Not LCase(Left(tabTBInfo.Fields("Action"), 5)) = "shit:" Then
            
            Set oLI = DetailLV.ListItems.Add()
            oLI.Tag = "greet_text"
            oLI.Text = "Greet Commands:"
            oLI.ForeColor = &H8000&
            oLI.ListSubItems.Add (1), "Detail", "Textblock " & tabMonsters.Fields("GreetTXT")
            oLI.ListSubItems(1).Tag = tabMonsters.Fields("GreetTXT")
            oLI.ListSubItems(1).ForeColor = &H8000&
        End If
    End If
End If

If nNMRVer >= 1.83 Then
    If tabMonsters.Fields("BSDefense") > 0 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "BS Defense:"
        oLI.ForeColor = &HC00000
        oLI.ListSubItems.Add (1), "Detail", tabMonsters.Fields("BSDefense")
        oLI.ListSubItems(1).ForeColor = &HC00000
    End If
End If

nTemp = 1 'TEMP FLAG FOR LIVING (set to 0 if nonliving/109 encountered)
For x = 0 To 9 'abilities
    If tabMonsters.Fields("Abil-" & x) > 0 And Not tabMonsters.Fields("Abil-" & x) = 146 Then '146=guarded by (handled below)
        If sAbil <> "" Then sAbil = sAbil & ", "
        sAbil = sAbil & GetAbilityStats(tabMonsters.Fields("Abil-" & x), tabMonsters.Fields("AbilVal-" & x))
        If Right(sAbil, 2) = ", " Then sAbil = Left(sAbil, Len(sAbil) - 2)
        
        If tabMonsters.Fields("Abil-" & x) = 34 And tabMonsters.Fields("AbilVal-" & x) > 0 Then 'dodge
            nMobDodge = tabMonsters.Fields("AbilVal-" & x)
            If bUseCharacter And val(frmMain.lblInvenCharStat(10).Tag) > 0 Then
                sAbil = sAbil & " (" & Fix((tabMonsters.Fields("AbilVal-" & x) * 10) / Fix(val(frmMain.lblInvenCharStat(10).Tag) / 8)) & "% @ " _
                    & val(frmMain.lblInvenCharStat(10).Tag) & " accy)"
            End If
            
        ElseIf tabMonsters.Fields("Abil-" & x) = 51 Then 'anti-magic
            bHasAntiMagic = True
            eMobDefenseFlags = eMobDefenseFlags Or DFIAM_IsAntiMag
            
        ElseIf tabMonsters.Fields("Abil-" & x) = 139 Then 'spellimmu
            nSpellImmuLVL = tabMonsters.Fields("AbilVal-" & x)
            
        ElseIf tabMonsters.Fields("Abil-" & x) = 28 Then 'magical
            nMagicLVL = tabMonsters.Fields("AbilVal-" & x)
        
        ElseIf tabMonsters.Fields("Abil-" & x) = 109 Then 'nonliving
            nTemp = 0
        
        ElseIf tabMonsters.Fields("Abil-" & x) = 78 Then 'animal
            eMobDefenseFlags = eMobDefenseFlags Or DF078_IsAnimal
           
        ElseIf tabMonsters.Fields("Abil-" & x) = 3 Then 'rcol
            nMobElementalResist(0) = tabMonsters.Fields("AbilVal-" & x)
            
        ElseIf tabMonsters.Fields("Abil-" & x) = 5 Then 'rfir
            nMobElementalResist(1) = tabMonsters.Fields("AbilVal-" & x)
               
        ElseIf tabMonsters.Fields("Abil-" & x) = 65 Then 'rsto
            nMobElementalResist(2) = tabMonsters.Fields("AbilVal-" & x)
                
        ElseIf tabMonsters.Fields("Abil-" & x) = 66 Then 'rlit
            nMobElementalResist(3) = tabMonsters.Fields("AbilVal-" & x)
                
        ElseIf tabMonsters.Fields("Abil-" & x) = 147 Then 'rwat
            nMobElementalResist(5) = tabMonsters.Fields("AbilVal-" & x)
                
        End If
    End If
Next x
If nTemp = 1 Then eMobDefenseFlags = eMobDefenseFlags Or DF109_IsLiving

If Not sAbil = "" Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Abilities: "
    oLI.ForeColor = &HC00000
    'oLI.ForeColor = &HFF00FF
    'oLI.Bold = True
    oLI.ListSubItems.Add (1), "Detail", sAbil
    oLI.ListSubItems(1).ForeColor = &HC00000
    'oLI.ListSubItems(1).ForeColor = &HFF00FF
    'oLI.Height = oLI.Height * 2
End If

Set oLI = Nothing
For x = 0 To 9 'mon guards
    If tabMonsters.Fields("Abil-" & x) = 146 And tabMonsters.Fields("AbilVal-" & x) > 0 Then
        If oLI Is Nothing Then
            Set oLI = DetailLV.ListItems.Add()
            oLI.Text = "Guarded by: "
            oLI.ForeColor = &H800080
            oLI.Tag = "Monster"
        Else
            Set oLI = DetailLV.ListItems.Add()
            oLI.Text = ""
            oLI.Tag = "Monster"
        End If
        
        oLI.ListSubItems.Add (1), "Detail", GetMonsterName(tabMonsters.Fields("AbilVal-" & x), bHideRecordNumbers)
        oLI.ListSubItems(1).ForeColor = &H800080
        
        tabMonsters.Seek "=", nMonsterNum
        oLI.ListSubItems(1).Tag = tabMonsters.Fields("AbilVal-" & x)
    End If
Next

Set oLI = DetailLV.ListItems.Add()
oLI.Text = ""
bNeedBlankRow = False

If bMobPrintCharDamageOutFirst And nNMRVer >= 1.83 And frmMain.optMonsterFilter(1).Value = True Then GoTo char_damage_out:
mob_attacks:

If bNeedBlankRow Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
    bNeedBlankRow = False
End If

y = 0
For x = 0 To 9 'item drops
    If Not tabMonsters.Fields("DropItem-" & x) = 0 Then
        y = y + 1
        Set oLI = DetailLV.ListItems.Add()
        If y = 1 Then
            oLI.Text = "Item Drops"
            oLI.Bold = True
        Else
            oLI.Text = ""
        End If
        oLI.Tag = "Item"
        
        oLI.ListSubItems.Add (1), "Detail", y & ". " & GetItemName(tabMonsters.Fields("DropItem-" & x), bHideRecordNumbers) _
            & " (" & tabMonsters.Fields("DropItem%-" & x) & "%)"
        oLI.ListSubItems(1).Tag = tabMonsters.Fields("DropItem-" & x)
        bNeedBlankRow = True
    End If
Next

'[START OF MOB'S ATTACKS]:
If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
bHasAttacks = False
For x = 0 To 4 'between round spells
    If Not tabMonsters.Fields("MidSpell-" & x) = 0 Then bHasAttacks = True
    If tabMonsters.Fields("AttType-" & x) > 0 And tabMonsters.Fields("AttType-" & x) <= 3 And tabMonsters.Fields("Att%-" & x) > 0 Then bHasAttacks = True
    If bHasAttacks Then Exit For
Next x

nMonsterEnergy = 1000
If nNMRVer >= 1.71 Then
    If Not tabMonsters.Fields("Energy") = 0 Then
        nMonsterEnergy = tabMonsters.Fields("Energy")
    End If
End If

If bNeedBlankRow Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
    bNeedBlankRow = False
End If

If bHasAttacks Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Mob's Attacks"
    oLI.Bold = True
    If nNMRVer >= 1.71 Then
        If Not tabMonsters.Fields("Energy") = 0 Then
            oLI.ListSubItems.Add (1), "Detail", nMonsterEnergy & " energy/round"
        End If
    End If
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
    bNeedBlankRow = False
End If

bNeedBlankRow = False
If bHasAttacks Then
    If nNMRVer >= 1.8 Then
        nLocalMonsterDamage = CalculateMonsterAvgDmg(nMonsterNum, 0) 'this is to get max damage
        nLocalMonsterDamage.nAverageDamage = tabMonsters.Fields("AvgDmg")
    Else
        nLocalMonsterDamage = CalculateMonsterAvgDmg(nMonsterNum, nGlobalMonsterSimRounds)
        If nMonsterDamageVsChar(nMonsterNum) = -1 Then
            nMonsterDamageVsChar(nMonsterNum) = nLocalMonsterDamage.nAverageDamage 'this is to get damage for older MME exports
        End If
    End If
    If nLocalMonsterDamage.nAverageDamage > 0 Or nLocalMonsterDamage.nMaxDamage > 0 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "Dmg/Round *"
        oLI.ForeColor = RGB(204, 0, 0)
    
        If nLocalMonsterDamage.nAverageDamage < nLocalMonsterDamage.nMaxDamage Then
            oLI.ListSubItems.Add (1), "Detail", "AVG: " & nLocalMonsterDamage.nAverageDamage & ", Max: " & nLocalMonsterDamage.nMaxDamage
        Else
            oLI.ListSubItems.Add (1), "Detail", "AVG: " & nLocalMonsterDamage.nAverageDamage
        End If
        
        If nNMRVer >= 1.8 Then
            oLI.ListSubItems(1).Text = oLI.ListSubItems(1).Text & "   * before character defenses, calculated when DB created"
        Else
            oLI.ListSubItems(1).Text = oLI.ListSubItems(1).Text & "   * before character defenses, " & nGlobalMonsterSimRounds & " round sim"
        End If
        oLI.ListSubItems(1).ForeColor = RGB(204, 0, 0)
        bNeedBlankRow = True
    End If
    
    nDamage = -1
    If bUseCharacter Then
        If frmMain.bAutoCalcMonDamage Then
            nDamage = CalculateMonsterDamageVsChar(nMonsterNum)
        ElseIf nMonsterDamageVsChar(nMonsterNum) >= 0 Then
            nDamage = nMonsterDamageVsChar(nMonsterNum)
        End If
    
        If nDamage >= 0 Then
            'nMonsterDamageVsChar(nMonsterNum) = nDamage
        
            Set oLI = DetailLV.ListItems.Add()
            oLI.Text = "Dmg/Round *"
            oLI.ForeColor = RGB(144, 4, 214)
        
            If frmMain.bAutoCalcMonDamage Then
                oLI.ListSubItems.Add (1), "Detail", "AVG: " & nDamage & ", Max Seen: " & clsMonAtkSim.nMaxRoundDamage
            Else
                oLI.ListSubItems.Add (1), "Detail", "AVG: " & nDamage
            End If
            
            oLI.ListSubItems(1).Text = oLI.ListSubItems(1).Text & "   * versus current character defenses, " & nGlobalMonsterSimRounds & " round sim"
            oLI.ListSubItems(1).ForeColor = RGB(144, 4, 214)
            bNeedBlankRow = True
        End If
    End If
    
    nDamage = -1
    If tCharProfile.nParty > 1 Then 'vs party
        If frmMain.bAutoCalcMonDamage Then
            nDamage = CalculateMonsterDamageVsChar(nMonsterNum, True)
        ElseIf nMonsterDamageVsParty(nMonsterNum) >= 0 Then
            nDamage = nMonsterDamageVsParty(nMonsterNum)
        End If
        
        If nDamage >= 0 Then
            Set oLI = DetailLV.ListItems.Add()
            oLI.Text = "Dmg/Round *"
            oLI.ForeColor = &H40C0&
            
            If frmMain.bAutoCalcMonDamage Then
                oLI.ListSubItems.Add (1), "Detail", "AVG: " & nDamage & ", Max Seen: " & clsMonAtkSim.nMaxRoundDamage
            Else
                oLI.ListSubItems.Add (1), "Detail", "AVG: " & nDamage
            End If
            
            oLI.ListSubItems(1).Text = oLI.ListSubItems(1).Text & "   * versus current PARTY defenses, " & nGlobalMonsterSimRounds & " round sim"
            oLI.ListSubItems(1).ForeColor = &H40C0&
            bNeedBlankRow = True
        End If
    End If
    If bNeedBlankRow Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = ""
        bNeedBlankRow = False
    End If
    
    If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
    
    nPercent = 0
    y = 0
    For x = 0 To 4 'between round spells
        If Not tabMonsters.Fields("MidSpell-" & x) = 0 Then
            
            y = y + 1
            Set oLI = DetailLV.ListItems.Add()
            If y = 1 Then
                oLI.Text = "Between Rounds"
            Else
                oLI.Text = ""
            End If
            oLI.Tag = "Spell"
            
            nPercent = tabMonsters.Fields("MidSpell%-" & x) - nPercent
            'nSpellNest = 0 'make sure this doesn't nest too deep
            oLI.ListSubItems.Add (1), "Detail", "(" & nPercent & "%) [" & _
                GetSpellName(tabMonsters.Fields("MidSpell-" & x), bHideRecordNumbers) _
                & ", " & PullSpellEQ(True, tabMonsters.Fields("MidSpellLVL-" & x), _
                tabMonsters.Fields("MidSpell-" & x), , , True) & "]"
            If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
            oLI.ListSubItems(1).Tag = tabMonsters.Fields("MidSpell-" & x) & "@" & tabMonsters.Fields("MidSpellLVL-" & x)
            
            If SpellHasAbility(tabMonsters.Fields("MidSpell-" & x), 60) >= 0 Then 'fear
                oLI.ListSubItems(1).ForeColor = &HC0&
                oLI.ListSubItems(1).Bold = True
            ElseIf SpellHasAbility(tabMonsters.Fields("MidSpell-" & x), 19) >= 0 Then 'poison
                oLI.ListSubItems(1).ForeColor = &H8000&
                oLI.ListSubItems(1).Bold = True
            ElseIf SpellHasAbility(tabMonsters.Fields("MidSpell-" & x), 71) >= 0 Then 'confusion
                oLI.ListSubItems(1).ForeColor = &H80FF&
                oLI.ListSubItems(1).Bold = True
            ElseIf SpellHasAbility(tabMonsters.Fields("MidSpell-" & x), 95) >= 0 Then 'slay
                oLI.ListSubItems(1).ForeColor = &HC0&
                oLI.ListSubItems(1).Bold = True
            ElseIf SpellHasAbility(tabMonsters.Fields("MidSpell-" & x), 13) <= -999 Then 'illu
                'oLI.ListSubItems(1).ForeColor = &HC0&
                oLI.ListSubItems(1).Bold = True
            End If
            
            nPercent = tabMonsters.Fields("MidSpell%-" & x)
            
            bNeedBlankRow = True
        End If
    Next
    If bNeedBlankRow Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = ""
        bNeedBlankRow = False
    End If
    
    
    nPercent = 0
    y = 0
    For x = 0 To 4 'attacks
        If tabMonsters.Fields("AttType-" & x) > 0 And tabMonsters.Fields("AttType-" & x) <= 3 And tabMonsters.Fields("Att%-" & x) > 0 Then
            y = y + 1
            Set oLI = DetailLV.ListItems.Add()
            bNeedBlankRow = True
            
            nPercent = tabMonsters.Fields("Att%-" & x) - nPercent
            If nPercent < 0 Then nPercent = 0
            
            If nNMRVer >= 1.8 Then
                If Round(tabMonsters.Fields("AttTrue%-" & x)) = 0 Then
                    oLI.Text = "(" & nPercent & "%) "
                Else
                    oLI.Text = "(" & Round(tabMonsters.Fields("AttTrue%-" & x)) & "%) "
                End If
                
                If Len(Trim(tabMonsters.Fields("AttName-" & x))) = 0 Then
                    oLI.Text = oLI.Text & "Attack " & y
                Else
                    oLI.Text = oLI.Text & Trim(tabMonsters.Fields("AttName-" & x))
                End If
            Else
                oLI.Text = "Attack " & y & " (" & nPercent & "%)"
            End If
            'oLI.ListSubItems.Add (1), "Detail", GetMonAttackType(tabMonsters.Fields("AttType-" & x))
            
            nPercent = tabMonsters.Fields("Att%-" & x)
            
            Select Case tabMonsters.Fields("AttType-" & x)
                Case 1, 3: 'normal, rob
                    
                    'Set oLI = DetailLV.ListItems.Add()
                    'oLI.Text = ""
                    oLI.ListSubItems.Add (1), "Detail", "Min-Max: " & tabMonsters.Fields("AttMin-" & x) & "-" & tabMonsters.Fields("AttMax-" & x)
                    
                    bNeedBlankRow = True
                    Set oLI = DetailLV.ListItems.Add()
                    oLI.Text = ""
                    oLI.ListSubItems.Add (1), "Detail", "Accuracy: " & tabMonsters.Fields("AttAcc-" & x)
                    
                    If tabMonsters.Fields("AttEnergy-" & x) > 0 Then
                        Set oLI = DetailLV.ListItems.Add()
                        oLI.Text = ""
                        oLI.ListSubItems.Add (1), "Detail", "Energy: " & tabMonsters.Fields("AttEnergy-" & x) _
                            & " (Max " & Fix(nMonsterEnergy / tabMonsters.Fields("AttEnergy-" & x)) & "x/round)"
                    Else
                        Set oLI = DetailLV.ListItems.Add()
                        oLI.Text = ""
                        oLI.ListSubItems.Add (1), "Detail", "Energy: " & tabMonsters.Fields("AttEnergy-" & x)
                    End If
                    
                    If Not tabMonsters.Fields("AttHitSpell-" & x) = 0 Then
                        bNeedBlankRow = True
                        Set oLI = DetailLV.ListItems.Add()
                        oLI.Text = ""
                        oLI.Tag = "Spell"
                        'nSpellNest = 0 'make sure this doesn't nest too deep
                        oLI.ListSubItems.Add (1), "Detail", "Hit Spell: [" & _
                            GetSpellName(tabMonsters.Fields("AttHitSpell-" & x), bHideRecordNumbers) _
                            & ", " & PullSpellEQ(False, , tabMonsters.Fields("AttHitSpell-" & x)) & "]"
                        If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
                        oLI.ListSubItems(1).Tag = tabMonsters.Fields("AttHitSpell-" & x)
                        
                        If SpellHasAbility(tabMonsters.Fields("AttHitSpell-" & x), 60) >= 0 Then 'fear
                            oLI.ListSubItems(1).ForeColor = &HC0&
                            oLI.ListSubItems(1).Bold = True
                        ElseIf SpellHasAbility(tabMonsters.Fields("AttHitSpell-" & x), 19) >= 0 Then 'poison
                            oLI.ListSubItems(1).ForeColor = &H8000&
                            oLI.ListSubItems(1).Bold = True
                        ElseIf SpellHasAbility(tabMonsters.Fields("AttHitSpell-" & x), 71) >= 0 Then 'confusion
                            oLI.ListSubItems(1).ForeColor = &H80FF&
                            oLI.ListSubItems(1).Bold = True
                        ElseIf SpellHasAbility(tabMonsters.Fields("AttHitSpell-" & x), 13) <= -999 Then 'illu
                            'oLI.ListSubItems(1).ForeColor = &HC0&
                            oLI.ListSubItems(1).Bold = True
                        End If
                    End If
                    
                Case 2: 'spell
                    bNeedBlankRow = True
                    'Set oLI = DetailLV.ListItems.Add()
                    'oLI.Text = ""
                    oLI.Tag = "Spell"
                    'nSpellNest = 0 'make sure this doesn't nest too deep
                    oLI.ListSubItems.Add (1), "Detail", "Spell: [" & _
                        GetSpellName(tabMonsters.Fields("AttAcc-" & x), bHideRecordNumbers) _
                        & ", " & PullSpellEQ(True, tabMonsters.Fields("AttMax-" & x), tabMonsters.Fields("AttAcc-" & x), , , True) & "]"
                    If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
                    oLI.ListSubItems(1).Tag = tabMonsters.Fields("AttAcc-" & x) & "@" & tabMonsters.Fields("AttMax-" & x)
                    
                    If SpellHasAbility(tabMonsters.Fields("AttAcc-" & x), 60) >= 0 Then 'fear
                        oLI.ListSubItems(1).ForeColor = &HC0&
                        oLI.ListSubItems(1).Bold = True
                    ElseIf SpellHasAbility(tabMonsters.Fields("AttAcc-" & x), 19) >= 0 Then 'poison
                        oLI.ListSubItems(1).ForeColor = &H8000&
                        oLI.ListSubItems(1).Bold = True
                    ElseIf SpellHasAbility(tabMonsters.Fields("AttAcc-" & x), 71) >= 0 Then 'confusion
                        oLI.ListSubItems(1).ForeColor = &H80FF&
                        oLI.ListSubItems(1).Bold = True
                    ElseIf SpellHasAbility(tabMonsters.Fields("AttAcc-" & x), 95) >= 0 Then 'slay
                        oLI.ListSubItems(1).ForeColor = &HC0&
                        oLI.ListSubItems(1).Bold = True
                    ElseIf SpellHasAbility(tabMonsters.Fields("AttAcc-" & x), 13) <= -999 Then 'illu
                        'oLI.ListSubItems(1).ForeColor = &HC0&
                        oLI.ListSubItems(1).Bold = True
                    End If
                    
                    If SpellIsAreaAttack(tabMonsters.Fields("AttAcc-" & x)) Then
                        Set oLI = DetailLV.ListItems.Add()
                        oLI.Text = ""
                        oLI.ListSubItems.Add (1), "Detail", "Target: " & GetSpellTargets(tabSpells.Fields("Targets"))
                        oLI.ListSubItems(1).ForeColor = &HFF&
                        'oLI.ListSubItems(1).Bold = True
                    End If
                    
                    Set oLI = DetailLV.ListItems.Add()
                    oLI.Text = ""
                    oLI.ListSubItems.Add (1), "Detail", "Success %: " & tabMonsters.Fields("AttMin-" & x)
                    
                    If tabMonsters.Fields("AttEnergy-" & x) > 0 Then
                        Set oLI = DetailLV.ListItems.Add()
                        oLI.Text = ""
                        oLI.ListSubItems.Add (1), "Detail", "Energy: " & tabMonsters.Fields("AttEnergy-" & x) _
                            & " (Max " & Fix(nMonsterEnergy / tabMonsters.Fields("AttEnergy-" & x)) & "x/round)"
                    Else
                        Set oLI = DetailLV.ListItems.Add()
                        oLI.Text = ""
                        oLI.ListSubItems.Add (1), "Detail", "Energy: " & tabMonsters.Fields("AttEnergy-" & x)
                    End If
                    
                    If Not tabMonsters.Fields("AttHitSpell-" & x) = 0 Then
                        Set oLI = DetailLV.ListItems.Add()
                        oLI.Text = ""
                        oLI.Tag = "Spell"
                        'nSpellNest = 0 'make sure this doesn't nest too deep
                        oLI.ListSubItems.Add (1), "Detail", "Hit Spell: [" & _
                            GetSpellName(tabMonsters.Fields("AttHitSpell-" & x), bHideRecordNumbers) _
                            & ", " & PullSpellEQ(False, , tabMonsters.Fields("AttHitSpell-" & x)) & "]"
                            
                        If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
                        
                        oLI.ListSubItems(1).Tag = tabMonsters.Fields("AttHitSpell-" & x)
                        
                        If SpellHasAbility(tabMonsters.Fields("AttHitSpell-" & x), 60) >= 0 Then 'fear
                            oLI.ListSubItems(1).ForeColor = &HC0&
                            oLI.ListSubItems(1).Bold = True
                        ElseIf SpellHasAbility(tabMonsters.Fields("AttHitSpell-" & x), 19) >= 0 Then 'poison
                            oLI.ListSubItems(1).ForeColor = &H8000&
                            oLI.ListSubItems(1).Bold = True
                        ElseIf SpellHasAbility(tabMonsters.Fields("AttHitSpell-" & x), 71) >= 0 Then 'confusion
                            oLI.ListSubItems(1).ForeColor = &H80FF&
                            oLI.ListSubItems(1).Bold = True
                        ElseIf SpellHasAbility(tabMonsters.Fields("AttHitSpell-" & x), 95) >= 0 Then 'slay
                            oLI.ListSubItems(1).ForeColor = &HC0&
                            oLI.ListSubItems(1).Bold = True
                        ElseIf SpellHasAbility(tabMonsters.Fields("AttHitSpell-" & x), 13) <= -999 Then 'illu
                            'oLI.ListSubItems(1).ForeColor = &HC0&
                            oLI.ListSubItems(1).Bold = True
                        End If
                    End If
                    
                    If SpellIsAreaAttack(tabMonsters.Fields("AttAcc-" & x)) Then
                        If GetSpellDuration(tabMonsters.Fields("AttAcc-" & x), tabMonsters.Fields("AttMax-" & x), True) = 0 Then
                            nTemp = SpellHasAbility(tabMonsters.Fields("AttAcc-" & x), 1) '1=damage
                            If nTemp > -1 Then
                                bNeedBlankRow = True
                                Set oLI = DetailLV.ListItems.Add()
                                oLI.Text = ""
                                oLI.ListSubItems.Add (1), "Detail", "This is an invalid spell and will not be cast (area attack spells from regular monster casts must use ability 17, Damage-MR)."
                                oLI.ListSubItems(1).Bold = True
                            End If
                        End If
                    End If
            End Select
            
            If bNeedBlankRow Then
                Set oLI = DetailLV.ListItems.Add()
                oLI.Text = ""
                bNeedBlankRow = False
            End If
        End If
    Next
Else
    If bNeedBlankRow Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = ""
        bNeedBlankRow = False
    End If
End If

':[END OF MOB'S ATTACKS]
If bMobPrintCharDamageOutFirst And nNMRVer >= 1.83 And frmMain.optMonsterFilter(1).Value = True Then GoTo done_attacks:
char_damage_out:

If bNeedBlankRow Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
    bNeedBlankRow = False
End If

'[START OF DAMAGE OUT + SCRIPTING / LAIR STUFF]:
If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum

nAvgDmg = 0
nMobDmg = 0
nExpDmgHP = 0
sDefenseDesc = ""

If tCharProfile.nParty > 1 And nMonsterDamageVsParty(nMonsterNum) >= 0 Then
    nMobDmg = nMonsterDamageVsParty(nMonsterNum)
    sDefenseDesc = " (vs party defenses)"
ElseIf bUseCharacter And nMonsterDamageVsChar(nMonsterNum) >= 0 Then
    nMobDmg = nMonsterDamageVsChar(nMonsterNum)
    sDefenseDesc = " (vs char defenses)"
ElseIf nNMRVer >= 1.8 Then
    nMobDmg = tabMonsters.Fields("AvgDmg")
    sDefenseDesc = " (vs default defenses)"
ElseIf nMonsterDamageVsDefault(nMonsterNum) >= 0 Then
    nMobDmg = nMonsterDamageVsDefault(nMonsterNum)
    sDefenseDesc = " (vs default defenses)"
Else
    nMobDmg = 0
End If
nAvgDmg = nMobDmg

If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum

If nNMRVer >= 1.83 And InStr(1, tabMonsters.Fields("Summoned By"), "lair", vbTextCompare) > 0 Then
    'If tLastAvgLairInfo.sGroupIndex <> tabMonsters.Fields("Summoned By") Or tLastAvgLairInfo.sGlobalAttackConfig <> sGlobalAttackConfig Then
        tLastAvgLairInfo = GetLairAveragesFromLocs(tabMonsters.Fields("Summoned By"))
    'End If
ElseIf tLastAvgLairInfo.sGroupIndex <> "" Then
    tLastAvgLairInfo = GetLairInfo("") 'reset
End If
tAvgLairInfo = tLastAvgLairInfo

If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum

If tCharProfile.nParty = 1 Then
    
    If nGlobalCharWeaponNumber(0) > 0 And (nGlobalAttackTypeMME = a1_PhysAttack Or nGlobalAttackTypeMME = a6_PhysBash Or nGlobalAttackTypeMME = a7_PhysSmash) Then
        
        nWeaponMagic = ItemHasAbility(nGlobalCharWeaponNumber(0), 28) 'magical
        nTemp = ItemHasAbility(nGlobalCharWeaponNumber(0), 142) 'hitmagic
        If nTemp > nWeaponMagic Then nWeaponMagic = nTemp
        
'    ElseIf nGlobalAttackSpellNum > 0 And (nGlobalAttackTypeMME = a2_Spell Or nGlobalAttackTypeMME = a3_SpellAny) Then
'
'        If SpellHasAbility(nGlobalAttackSpellNum, 23) >= 0 Then eAttackFlags = eAttackFlags Or AR023_Undead
    
    ElseIf (tCharProfile.nClass > 0 Or tCharProfile.nRace > 0) And (nGlobalAttackTypeMME = a1_PhysAttack Or nGlobalAttackTypeMME = a4_MartialArts) Then  'punch/MA
        
        If tCharProfile.nClass > 0 Then
            nTemp2 = ClassHasAbility(tCharProfile.nClass, 142)
            If nTemp2 < 1 Then nTemp2 = 0
            If nTemp2 > nWeaponMagic Then nWeaponMagic = nTemp2
        End If
        If tCharProfile.nRace > 0 Then
            nTemp2 = RaceHasAbility(tCharProfile.nRace, 142)
            If nTemp2 < 1 Then nTemp2 = 0
            If nTemp2 > nWeaponMagic Then nWeaponMagic = nTemp2
        End If
    End If
    
    If nMagicLVL > 0 And nGlobalAttackTypeMME > a0_oneshot And bGlobalAttackBackstab = True Then
        nTemp = nGlobalAttackBackstabWeapon
        If nGlobalAttackBackstabWeapon = 0 Then nTemp = nGlobalCharWeaponNumber(0)
        If nTemp > 0 Then
            nBackstabWeaponMagic = ItemHasAbility(nTemp, 28) 'magical
            nTemp2 = ItemHasAbility(nTemp, 142) 'hitmagic
            If nTemp2 > nBackstabWeaponMagic Then nBackstabWeaponMagic = nTemp2
        
        ElseIf nGlobalAttackTypeMME = a4_MartialArts Or (nGlobalAttackTypeMME = a1_PhysAttack And nGlobalCharWeaponNumber(0) = 0) Then
            'this would mean we already checked for class/race hitmagic above
            nBackstabWeaponMagic = nWeaponMagic
            
        ElseIf tCharProfile.nClass > 0 Or tCharProfile.nRace > 0 Then 'surprise punch
            
            If tCharProfile.nClass > 0 Then
                nTemp2 = ClassHasAbility(tCharProfile.nClass, 142)
                If nTemp2 < 1 Then nTemp2 = 0
                If nTemp2 > nBackstabWeaponMagic Then nBackstabWeaponMagic = nTemp2
            End If
            If tCharProfile.nRace > 0 Then
                nTemp2 = RaceHasAbility(tCharProfile.nRace, 142)
                If nTemp2 < 1 Then nTemp2 = 0
                If nTemp2 > nBackstabWeaponMagic Then nBackstabWeaponMagic = nTemp2
            End If
        End If
    End If
End If
If nWeaponMagic < 0 Then nWeaponMagic = 0
If nBackstabWeaponMagic < 0 Then nBackstabWeaponMagic = 0

If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum

For iAttack = 1 To IIf(tAvgLairInfo.nTotalLairs > 0, 2, 1) 'And frmMain.optMonsterFilter(1).Value = True
    
    If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
    
    nDamageOut = 0
    nMinDamageOut = -9999
    nOverrideRTK = 0
    sBackstabText = ""
    tBackStab.nRoundTotal = 0
    tBackStab.sAttackDetail = ""
    sImmuTXT = ""
    DF_Flags = 0
    bValidTarget = False
    Select Case iAttack
        Case 1:
            sHeader = "Damage vs Mob"
            nAvgDmg = nMobDmg
            nCalcDamageHP = tabMonsters.Fields("HP")
            nCalcDamageHPRegen = tabMonsters.Fields("HPRegen")
            nCalcDamageAC = tabMonsters.Fields("ArmourClass")
            nCalcDamageDR = tabMonsters.Fields("DamageResist")
            nCalcDamageDodge = nMobDodge
            nCalcDamageMR = tabMonsters.Fields("MagicRes")
            For x = 0 To 5
                nCalcElementalResist(x) = nMobElementalResist(x)
            Next x
            'bIsAntiMagic = bHasAntiMagic
            nCalcDamageNumMobs = 1
            nCalcSpellImmuLVL = nSpellImmuLVL
            nCalcMagicLVL = nMagicLVL
            'bIsUndead = IIf(tabMonsters.Fields("Undead") = 1, True, False)
            DF_Flags = eMobDefenseFlags
        Case 2:
            'note that exp/hour calculation below is dependant on tSpellCast getting calculated here
            'and the lair version of it getting calculated last when calculating by lair
            sHeader = "Damage vs Lair"
            nAvgDmg = tAvgLairInfo.nAvgDmgLair
            nCalcDamageHP = tAvgLairInfo.nAvgHP
            nCalcDamageHPRegen = 0
            nCalcDamageAC = tAvgLairInfo.nAvgAC
            nCalcDamageDR = tAvgLairInfo.nAvgDR
            nCalcDamageDodge = tAvgLairInfo.nAvgDodge
            nCalcDamageMR = tAvgLairInfo.nAvgMR
            nCalcElementalResist(0) = tAvgLairInfo.nAvgRCOL
            nCalcElementalResist(1) = tAvgLairInfo.nAvgRFIR
            nCalcElementalResist(2) = tAvgLairInfo.nAvgRSTO
            nCalcElementalResist(3) = tAvgLairInfo.nAvgRLIT
            nCalcElementalResist(5) = tAvgLairInfo.nAvgRWAT
            nCalcDamageNumMobs = tAvgLairInfo.nMaxRegen
            nOverrideRTK = tAvgLairInfo.nRTK
            nCalcSpellImmuLVL = tAvgLairInfo.nSpellImmuLVL
            nCalcMagicLVL = tAvgLairInfo.nMagicLVL
            'bIsAntiMagic = IIf(tAvgLairInfo.nNumAntiMagic >= (tAvgLairInfo.nMobs / 2), True, False)
            'bIsUndead = IIf(tAvgLairInfo.nNumUndeads >= (tAvgLairInfo.nMobs * LAIR_FLAG_RATIO), True, False)
            If tAvgLairInfo.nNumAntiMagic > 0 And tAvgLairInfo.nNumAntiMagic >= (tAvgLairInfo.nMobs / 2) Then DF_Flags = DF_Flags Or DFIAM_IsAntiMag
            If tAvgLairInfo.nNumUndeads > 0 And tAvgLairInfo.nNumUndeads >= (tAvgLairInfo.nMobs * LAIR_FLAG_RATIO) Then DF_Flags = DF_Flags Or DF023_IsUndead
            If tAvgLairInfo.nNumLiving > 0 And tAvgLairInfo.nNumLiving >= (tAvgLairInfo.nMobs * LAIR_FLAG_RATIO) Then DF_Flags = DF_Flags Or DF109_IsLiving
            If tAvgLairInfo.nNumAnimals > 0 And tAvgLairInfo.nNumAnimals >= (tAvgLairInfo.nMobs * LAIR_FLAG_RATIO) Then DF_Flags = DF_Flags Or DF078_IsAnimal
            
        Case Else: GoTo no_attack
    End Select
    
    If tCharProfile.nParty > 1 Or nGlobalAttackTypeMME = a5_Manual Then '5=manual
        
        If tCharProfile.nParty > 1 Then
            sTemp = "/round (party)"
        Else
            sTemp = "/round (manual)"
        End If
        
        If iAttack = 2 Then  'lair
            nDamageOut = tAvgLairInfo.nDamageOut
            nMinDamageOut = tAvgLairInfo.nMinDamageOut
            nSurpriseDamageOut = tAvgLairInfo.nSurpriseDamageOut
            If nSurpriseDamageOut > 0 Then sBackstabText = " + " & CStr(nSurpriseDamageOut) & " surprise round"
            Call AddMonsterDamageOutText(DetailLV, sHeader, nDamageOut & sTemp & sImmuTXT & sBackstabText, , _
                nDamageOut, nCalcDamageHP, nCalcDamageHPRegen, nAvgDmg, tCharProfile.nHP, sDefenseDesc, nCalcDamageNumMobs, , nOverrideRTK, _
                nSurpriseDamageOut, , nMinDamageOut)
        Else
            'not lair + (party or manual attack)
            nDmgOut = GetDamageOutput(tabMonsters.Fields("Number"), nCalcDamageAC, nCalcDamageDR, nCalcDamageMR, nCalcDamageDodge, _
                DF_Flags, , nCalcSpellImmuLVL, nCalcMagicLVL)
                
            nDamageOut = nDmgOut(0)
            nMinDamageOut = nDmgOut(1)
            nSurpriseDamageOut = nDmgOut(2)
            
'            If tCharProfile.nParty = 1 And nCalcMagicLVL > 0 And bGlobalAttackBackstab And nGlobalAttackBackstabWeapon > 0 _
'                And nGlobalAttackTypeMME > a0_oneshot And nBackstabWeaponMagic < nCalcMagicLVL Then
'                nSurpriseDamageOut = 0
'            End If
            If nSurpriseDamageOut > -9999 Then
                If nSurpriseDamageOut = -9998 Then
                    nSurpriseDamageOut = -9999
                    sBackstabText = " + 0 surprise round [immune:MagicLVL]"
                Else
                    sBackstabText = " + " & CStr(nSurpriseDamageOut) & " surprise round"
                End If
            End If
            
            If nDamageOut = -9998 Then
                sImmuTXT = " [immune:MagicLVL]"
                nDamageOut = 0
            End If
            Call AddMonsterDamageOutText(DetailLV, sHeader, nDamageOut & sTemp & sImmuTXT & sBackstabText, , _
                nDamageOut, nCalcDamageHP, nCalcDamageHPRegen, nAvgDmg, tCharProfile.nHP, sDefenseDesc, nCalcDamageNumMobs, , nOverrideRTK, _
                nSurpriseDamageOut, , nMinDamageOut)
        End If
        
    Else
        'nGlobalAttackTypeMME (from frmPopUpOptions): 0-none/one-shot, 1-weapon, 2-spell user, 3-spell any, 4-MA, 5-manual, 6-bash, 7-smash
        'CalculateAttack > nAttackTypeMUD: 1-punch, 2-kick, 3-jumpkick, 4-surprise, 5-normal, 6-bash, 7-smash
        
        If tCharProfile.nParty = 1 And nGlobalAttackTypeMME > a0_oneshot And bGlobalAttackBackstab Then
            If nCalcMagicLVL > 0 And nBackstabWeaponMagic < nCalcMagicLVL Then
                nSurpriseDamageOut = 0
                sBackstabText = " + 0 surprise round [immune:MagicLVL]"
            Else
                nTemp = 0
                If nGlobalAttackBackstabWeapon > 0 Then
                    nTemp = nGlobalAttackBackstabWeapon
                ElseIf nGlobalAttackBackstabWeapon = 0 And nGlobalCharWeaponNumber(0) > 0 Then
                    nTemp = nGlobalCharWeaponNumber(0)
                End If
                Call PopulateCharacterProfile(tBackStabProfile, False, True, a4_Surprise, nTemp)
                If nNMRVer >= 1.83 Then nBSDefense = tabMonsters.Fields("BSDefense")
                tBackStab = CalculateAttack(tBackStabProfile, a4_Surprise, nTemp, False, nSpeedAdj, nCalcDamageAC, nCalcDamageDR, nCalcDamageDodge, , , , , nBSDefense)
                nSurpriseDamageOut = tBackStab.nRoundTotal
                sBackstabText = " + " & CStr(nSurpriseDamageOut) & " surprise round"
            End If
        End If
        
        Select Case nGlobalAttackTypeMME
            Case 1, 6, 7: 'eq'd weapon, bash, smash
                
                'allowing this to still happen even if we know it will be immune in order to populate tAttack
                If nGlobalAttackTypeMME = a6_PhysBash Then 'bash w/wep
                    tAttack = CalculateAttack(tForcedCharProfile, a6_Bash, nGlobalCharWeaponNumber(0), False, nSpeedAdj, nCalcDamageAC, nCalcDamageDR, nCalcDamageDodge)
                ElseIf nGlobalAttackTypeMME = a7_PhysSmash Then 'smash w/wep
                    tAttack = CalculateAttack(tForcedCharProfile, a7_Smash, nGlobalCharWeaponNumber(0), False, nSpeedAdj, nCalcDamageAC, nCalcDamageDR, nCalcDamageDodge)
                Else 'EQ'd Weapon reg attack
                    tAttack = CalculateAttack(tForcedCharProfile, a5_Normal, nGlobalCharWeaponNumber(0), False, nSpeedAdj, nCalcDamageAC, nCalcDamageDR, nCalcDamageDodge)
                End If
                
                If nCalcMagicLVL > 0 And nWeaponMagic < nCalcMagicLVL And nGlobalCharWeaponNumber(0) > 0 Then
                    nDamageOut = 0
                    nMinDamageOut = 0
                    nOverrideRTK = 0
                    sImmuTXT = " [immune:MagicLVL]"
                Else
                    nDamageOut = tAttack.nRoundTotal
                    nMinDamageOut = tAttack.nMinDmg
                    If tAttack.nAvgExtraHit > 0 And tAttack.nAvgExtraHit = tAttack.nAvgExtraSwing Then nMinDamageOut = nMinDamageOut + tAttack.nAvgExtraHit
                    nMinDamageOut = nMinDamageOut * tAttack.nSwings
                End If
                
                Call AddMonsterDamageOutText(DetailLV, sHeader, nDamageOut & "/round (" & tAttack.sAttackDesc & ")" & sImmuTXT & sBackstabText, tAttack.sAttackDetail, _
                    nDamageOut, nCalcDamageHP, nCalcDamageHPRegen, nAvgDmg, tCharProfile.nHP, sDefenseDesc, nCalcDamageNumMobs, , nOverrideRTK, _
                    nSurpriseDamageOut, tBackStab.sAttackDetail, nMinDamageOut)
                
            Case 2, 3:
                '2-spell learned: GetSpellShort(nGlobalAttackSpellNum) & " @ " & Val(txtGlobalLevel(0).Text)
                '3-spell any: GetSpellShort(nGlobalAttackSpellNum) & " @ " & nGlobalAttackSpellLVL
                If nGlobalAttackSpellNum <= 0 Then GoTo no_attack:
                
                tSpellcast = CalculateSpellCast(tForcedCharProfile, nGlobalAttackSpellNum, _
                                IIf(nGlobalAttackTypeMME = a3_SpellAny, nGlobalAttackSpellLVL, tForcedCharProfile.nLevel), _
                                nCalcDamageMR, (DF_Flags And DFIAM_IsAntiMag) <> 0, nCalcElementalResist(0), nCalcElementalResist(1), _
                                nCalcElementalResist(2), nCalcElementalResist(3), nCalcElementalResist(5))
                
                If nCalcSpellImmuLVL = 0 Or tSpellcast.nCastLevel > nCalcSpellImmuLVL Then
                    If eAttackFlags = AR000_Unknown Then
                        If SpellSeek(nGlobalAttackSpellNum) Then
                            For x = 0 To 9
                                Select Case tabSpells.Fields("Abil-" & x) 'tabSpells.Fields("AbilVal-" & x)
                                    Case 0: 'nada
                                    Case 23: eAttackFlags = (eAttackFlags Or AR023_Undead)
                                    Case 80: eAttackFlags = (eAttackFlags Or AR080_Animal)
                                    Case 108: eAttackFlags = (eAttackFlags Or AR108_Living)
                                End Select
                            Next x
                            If eAttackFlags <= AR001_None Then bValidTarget = True
                        End If
                    ElseIf eAttackFlags = AR001_None Then
                        bValidTarget = True
                    End If
                    
                    If bValidTarget = False Then
                        If eAttackFlags > AR001_None Then
                            If (eAttackFlags And AR023_Undead) <> 0 Then
                                If (DF_Flags And DF023_IsUndead) <> 0 Then bValidTarget = True
                            ElseIf (eAttackFlags And AR080_Animal) <> 0 Then
                                If (DF_Flags And DF078_IsAnimal) <> 0 Then bValidTarget = True
                            ElseIf (eAttackFlags And AR108_Living) <> 0 Then
                                If (DF_Flags And DF109_IsLiving) <> 0 Then bValidTarget = True
                            End If
                        Else
                            bValidTarget = True
                        End If
                    End If
                End If
                
                
                If bValidTarget Then
                    nDamageOut = tSpellcast.nAvgRoundDmg
                    nMinDamageOut = tSpellcast.nMinCast * tSpellcast.nNumCasts
                Else
                    nDamageOut = 0
                    nMinDamageOut = 0
                    If (nCalcSpellImmuLVL > 0 And tSpellcast.nCastLevel <= nCalcSpellImmuLVL) Then sImmuTXT = AutoAppend(sImmuTXT, "SpellImmuLVL", "+")
                    If ((eAttackFlags And AR023_Undead) <> 0 And (DF_Flags And DF023_IsUndead) = 0) Then sImmuTXT = AutoAppend(sImmuTXT, "NotUndead", "+")
                    If ((eAttackFlags And AR080_Animal) <> 0 And (DF_Flags And DF078_IsAnimal) = 0) Then sImmuTXT = AutoAppend(sImmuTXT, "NotAnimal", "+")
                    If ((eAttackFlags And AR108_Living) <> 0 And (DF_Flags And DF109_IsLiving) = 0) Then sImmuTXT = AutoAppend(sImmuTXT, "NotLiving", "+")
                    sImmuTXT = " [immune:" & sImmuTXT & "]"
                End If
                
                Call AddMonsterDamageOutText(DetailLV, sHeader, _
                    IIf(sImmuTXT = "", tSpellcast.sAvgRound, 0) & "/round (" & tSpellcast.sSpellName & ")" & sImmuTXT & sBackstabText, tSpellcast.sMMA, _
                    nDamageOut, nCalcDamageHP, nCalcDamageHPRegen, nAvgDmg, tCharProfile.nHP, sDefenseDesc, nCalcDamageNumMobs, tSpellcast.nOOM, nOverrideRTK, _
                    nSurpriseDamageOut, tBackStab.sAttackDetail, nMinDamageOut)
                
            Case 4: 'martial arts attack
                '1-Punch, 2-Kick, 3-JumpKick
                Select Case nGlobalAttackMA
                    Case 2: 'kick
                        tAttack = CalculateAttack(tForcedCharProfile, a2_Kick, , False, nSpeedAdj, nCalcDamageAC, nCalcDamageDR, nCalcDamageDodge)
                        nDamageOut = tAttack.nRoundTotal
                        nMinDamageOut = tAttack.nMinDmg * tAttack.nSwings
                    Case 3: 'jumpkick
                        tAttack = CalculateAttack(tForcedCharProfile, a3_Jumpkick, , False, nSpeedAdj, nCalcDamageAC, nCalcDamageDR, nCalcDamageDodge)
                        nDamageOut = tAttack.nRoundTotal
                        nMinDamageOut = tAttack.nMinDmg * tAttack.nSwings
                    Case Else: 'punch
                        tAttack = CalculateAttack(tForcedCharProfile, a1_Punch, , False, nSpeedAdj, nCalcDamageAC, nCalcDamageDR, nCalcDamageDodge)
                        nDamageOut = tAttack.nRoundTotal
                        nMinDamageOut = tAttack.nMinDmg * tAttack.nSwings
                End Select
                Call AddMonsterDamageOutText(DetailLV, sHeader, tAttack.nRoundTotal & "/round (" & tAttack.sAttackDesc & ")" & sBackstabText, tAttack.sAttackDetail, _
                    nDamageOut, nCalcDamageHP, nCalcDamageHPRegen, nAvgDmg, tCharProfile.nHP, sDefenseDesc, nCalcDamageNumMobs, , nOverrideRTK, _
                    nSurpriseDamageOut, tBackStab.sAttackDetail, nMinDamageOut)
                
            'Case 5: 'manual
                'nDamageOut = nGlobalAttackManualP
                'nDamageOutSpell = nGlobalAttackManualM
                'Call AddMonsterDamageOutText(DetailLV, sHeader, nGlobalAttackManualP & "/round (manual)", , nDamageOut, nCalcDamageHP, nCalcDamageHPRegen, nAvgDmg, tCharProfile.nHP, sDefenseDesc, nCalcDamageNumMobs)
                
            Case Else: '1-Shot All
                nDamageOut = 9999999
                Call AddMonsterDamageOutText(DetailLV, sHeader, "(assuming one-shot)")
                
        End Select
    End If
no_attack:
    If iAttack = 1 Then nDamageVMob = nDamageOut
    If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
Next iAttack

If tAvgLairInfo.nTotalLairs > 0 And frmMain.optMonsterFilter(1).Value = True Then
    nAvgDmg = tAvgLairInfo.nAvgDmg
    If tCharProfile.nParty > 1 Then
        sDefenseDesc = " (LAIR avg vs party defenses)"
    Else
        sDefenseDesc = " (LAIR avg vs defenses)"
    End If
'Else
'    nAvgDmg = nMobDmg
End If

If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum

If nExp <= 1 Or tabMonsters.Fields("HP") < 1 Then GoTo done_scripting:

Set oLI = DetailLV.ListItems.Add()

If tAvgLairInfo.nTotalLairs > 0 And (tabMonsters.Fields("RegenTime") > 0 Or InStr(1, tabMonsters.Fields("Summoned By"), "Room", vbTextCompare) > 0) Then
    oLI.Text = "Scripting vs MOB"
ElseIf tAvgLairInfo.nTotalLairs > 0 And frmMain.optMonsterFilter(1).Value = True Then
    oLI.Text = "Scripting vs Lair"
Else
    oLI.Text = "Scripting"
End If
oLI.Bold = True

If nNMRVer < 1.83 Or frmMain.optMonsterFilter(1).Value = False Then GoTo script_value:

If Not bUseCharacter Then oLI.ListSubItems.Add (1), "Detail", "(global filter inactive. default stats used.)"

Set oLI = DetailLV.ListItems.Add()
oLI.Text = "Scripting Estimate"

nExpPerHour = 0
If tabMonsters.Fields("RegenTime") = 0 And tAvgLairInfo.nTotalLairs > 0 And frmMain.optMonsterFilter(1).Value = True Then
    
    tExpInfo = CalcExpPerHour(tAvgLairInfo.nAvgExp, tAvgLairInfo.nAvgDelay, tAvgLairInfo.nMaxRegen, tAvgLairInfo.nTotalLairs, _
                    tAvgLairInfo.nPossSpawns, tAvgLairInfo.nRTK, tAvgLairInfo.nDamageOut, tCharProfile.nHP, tCharProfile.nHPRegen, _
                    tAvgLairInfo.nAvgDmgLair, tAvgLairInfo.nAvgHP, , tCharProfile.nDamageThreshold, _
                    tSpellcast.nManaCost, tCharProfile.nSpellOverhead, tCharProfile.nMaxMana, tCharProfile.nManaRegen, tCharProfile.nMeditateRate, _
                    tAvgLairInfo.nAvgWalk, tCharProfile.nEncumPCT, , tAvgLairInfo.nSurpriseDamageOut)
    
    nExpPerHour = tExpInfo.nExpPerHour
    
ElseIf tabMonsters.Fields("RegenTime") > 0 Or InStr(1, tabMonsters.Fields("Summoned By"), "Room", vbTextCompare) > 0 Then
    
    tExpInfo = CalcExpPerHour(nExp, tabMonsters.Fields("RegenTime"), 1, -1, _
                    , , nDamageVMob, tCharProfile.nHP, tCharProfile.nHPRegen, _
                    nMobDmg, tabMonsters.Fields("HP"), tabMonsters.Fields("HPRegen"), tCharProfile.nDamageThreshold, _
                    tSpellcast.nManaCost, tCharProfile.nSpellOverhead, tCharProfile.nMaxMana, tCharProfile.nManaRegen, tCharProfile.nMeditateRate, _
                    0, tCharProfile.nEncumPCT, , nSurpriseDamageOut)
    
    nExpPerHour = tExpInfo.nExpPerHour

ElseIf nNMRVer >= 1.83 Then
    tExpInfo.sRTCText = "(No lairs and not assigned as an NPC)"
End If

If nExpPerHour > 0 And tCharProfile.nParty > 1 Then
    nExpPerHourEA = Round(nExpPerHour / tCharProfile.nParty, 1)
Else
    nExpPerHourEA = nExpPerHour
End If

If nExpPerHour > 0 Then
    If nExpPerHour > 1000000 Then
        sTemp = Format((nExpPerHour / 1000000), "#,#.00") & " M"
    ElseIf nExpPerHour > 1000 Then
        sTemp = Format((nExpPerHour / 1000), "#,#.0") & " K"
    Else
        sTemp = IIf(nExpPerHour > 0, Format(RoundUp(nExpPerHour), "#,#"), "0")
    End If
    
    sTemp = sTemp & "/hr"
    If nExpPerHour <> nExpPerHourEA And nExpPerHourEA > 0 Then
        
        If nExpPerHourEA > 1000000 Then
            sExpEa = Format((nExpPerHourEA / 1000000), "#,#.0") & " M"
        ElseIf nExpPerHourEA > 1000 Then
            sExpEa = Format((nExpPerHourEA / 1000), "#,#.0") & " K"
        Else
            sExpEa = IIf(nExpPerHourEA > 0, Format(RoundUp(nExpPerHourEA), "#,#"), "0")
        End If
        
        sTemp = sTemp & " (" & sExpEa & "/hr ea.)"
        
    End If
    
ElseIf nExpPerHour = -1 Then
    If tabMonsters.Fields("RegenTime") = 0 And tAvgLairInfo.nTotalLairs > 0 And frmMain.optMonsterFilter(1).Value = True Then
        sTemp = "The lairs of this mob"
    Else
        sTemp = "This mob"
    End If
    sTemp = sTemp & " deemed undefeatable against current stats " & sTemp2 & "."
Else
    sTemp = "0"
End If

oLI.ListSubItems.Add (1), "Detail", sTemp

If Len(tExpInfo.sRTCText) > 0 Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
    oLI.ListSubItems.Add (1), "Detail", tExpInfo.sRTCText
End If

If Len(tExpInfo.sMoveText) > 0 Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
    oLI.ListSubItems.Add (1), "Detail", tExpInfo.sMoveText
End If

'If Len(tExpInfo.sLairText) > 0 Then
'    Set oLI = DetailLV.ListItems.Add()
'    oLI.Text = ""
'    oLI.ListSubItems.Add (1), "Detail", tExpInfo.sLairText
'End If

If tExpInfo.nTimeRecovering > 0 And nExpPerHour >= 0 Then
    sTemp = ""
    If Len(tExpInfo.sManaRecovery) > 0 And Len(tExpInfo.sHitpointRecovery) > 0 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = ""
        oLI.ListSubItems.Add (1), "Detail", tExpInfo.sTimeRecovering ' & " (derived from combination of below)"
        sTemp = " > "
    End If
    
    If Len(tExpInfo.sHitpointRecovery) > 0 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = ""
        oLI.ListSubItems.Add (1), "Detail", sTemp & tExpInfo.sHitpointRecovery
    End If
    
    If Len(tExpInfo.sManaRecovery) > 0 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = ""
        oLI.ListSubItems.Add (1), "Detail", sTemp & tExpInfo.sManaRecovery & IIf(tSpellcast.nOOM > 0, " (OOM every " & tSpellcast.nOOM & " rounds)", "")
    End If
End If

If tabMonsters.Fields("RegenTime") = 0 And tAvgLairInfo.nTotalLairs > 0 And frmMain.optMonsterFilter(1).Value = True Then
    If bUseCharacter And tCharProfile.nParty < 2 Then 'no party, vs char
        If nAvgDmg > tCharProfile.nDamageThreshold Or tExpInfo.nHitpointRecovery > 0 Then
            If Len(sMonsterDamageVsCharDefenseConfig) > 0 And sMonsterDamageVsCharDefenseConfig <> sGlobalCharDefenseDescription Then
                Set oLI = DetailLV.ListItems.Add()
                oLI.Text = ""
                oLI.ListSubItems.Add (1), "Detail", "Monster damage vs Character defense sims may be stale."
                oLI.ListSubItems(1).Bold = True
                Set oLI = DetailLV.ListItems.Add()
                oLI.Text = ""
                oLI.ListSubItems.Add (1), "Detail", "Old or base damage utilized where missing. Recalculate all from options menu" & IIf(bDontPromptCalcCharMonsterDamage, ".", " or apply filter.")
                oLI.ListSubItems(1).Bold = False
            ElseIf Len(sMonsterDamageVsCharDefenseConfig) = 0 And Len(sGlobalCharDefenseDescription) > 0 Then
                Set oLI = DetailLV.ListItems.Add()
                oLI.Text = ""
                oLI.ListSubItems.Add (1), "Detail", "All monster damage vs Character defense sims not calculated."
                oLI.ListSubItems(1).Bold = True
                Set oLI = DetailLV.ListItems.Add()
                oLI.Text = ""
                oLI.ListSubItems.Add (1), "Detail", "Old or base damage utilized where missing. Calculate from options menu" & IIf(bDontPromptCalcCharMonsterDamage, ".", " or apply filter.")
                oLI.ListSubItems(1).Bold = False
            End If
        End If
    ElseIf tCharProfile.nParty > 1 Then 'vs party
        If bMonsterDamageVsPartyCalculated = False Or bDontPromptCalcPartyMonsterDamage = False Then
            Set oLI = DetailLV.ListItems.Add()
            oLI.Text = ""
            oLI.ListSubItems.Add (1), "Detail", "Monster damage vs Party defense sims may be incomplete. Calculate all from options menu or re-apply filter."
            oLI.ListSubItems(1).Bold = True
        End If
    End If
End If

'Set oLI = DetailLV.ListItems.Add()
'oLI.Text = ""

'================================================================
'================================================================
'==============   2025.06.15 ===============
'================================================================

'a lot of this repeated in addmonsterlv
script_value:
If (tAvgLairInfo.nTotalLairs = 0 Or frmMain.optMonsterFilter(1).Value = False) And (frmMain.optMonsterFilter(1).Value = False Or nNMRVer < 1.83) Then
    nExpDmgHP = 0
    If nAvgDmg > 0 Or tabMonsters.Fields("HP") > 0 Then
        nExpDmgHP = Round(nExp / ((nAvgDmg * 2) + tabMonsters.Fields("HP")), 2) * 100
    Else
        nExpDmgHP = nExp
    End If
    
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Exp/((Dmg*2)+HP)"
    oLI.ListSubItems.Add (1), "Detail", IIf(nExpDmgHP > 0, Format(nExpDmgHP, "#,#"), 0) & " (" & nExp & " / ((" & nAvgDmg & " x 2) + " & tabMonsters.Fields("HP") & ")) * 100" & sDefenseDesc
    
    If frmMain.chkGlobalFilter.Value = 0 And nMonsterDamageVsChar(nMonsterNum) >= 0 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = " "
        oLI.ListSubItems.Add (1), "Detail", "Calculated damage vs character defenses not utilized because global filter is disabled"
    End If
End If

'skip script value stuff if calculating by lair
If tAvgLairInfo.nTotalLairs > 0 And frmMain.optMonsterFilter(1).Value = True Then GoTo done_scripting:
    
'this is for scripting value
nPossSpawns = 0
nLairPCT = 0
nMaxLairsBeforeRegen = 36 'nTheoreticalMaxLairsPerRegenPeriod
If InStr(1, tabMonsters.Fields("Summoned By"), "(lair)", vbTextCompare) > 0 Then
    nPossSpawns = InstrCount(tabMonsters.Fields("Summoned By"), "(lair)")
    sPossSpawns = nPossSpawns
    
    If nMonsterPossy(nMonsterNum) > 0 Then nMaxLairsBeforeRegen = Round(nMaxLairsBeforeRegen / nMonsterPossy(nMonsterNum), 2)
    If nPossSpawns < nMaxLairsBeforeRegen Then
        nLairPCT = Round(nPossSpawns / nMaxLairsBeforeRegen, 2)
    Else
        nLairPCT = 1
    End If
End If

nPossyPCT = 1
nScriptValue = 0
If nMonsterPossy(nMonsterNum) > 0 Then
    nPossyPCT = 1 + ((nMonsterPossy(nMonsterNum) - 1) / 5)
    If nPossyPCT > 3 Then nPossyPCT = 3
    If nPossyPCT < 1 Then nPossyPCT = 1
End If

If nNMRVer >= 1.83 And frmMain.chkGlobalFilter.Value = 0 Then
    nScriptValue = tabMonsters.Fields("ScriptValue")
ElseIf tabMonsters.Fields("RegenTime") = 0 And nLairPCT > 0 Then
    If nPossyPCT > 1 And (tAvgLairInfo.nTotalLairs = 0 Or frmMain.optMonsterFilter(1).Value = False) And (nAvgDmg > 0 Or tabMonsters.Fields("HP")) Then
        nExpDmgHP = Round(nExp / (((nAvgDmg * 2) + tabMonsters.Fields("HP")) * nPossyPCT), 2) * 100
    End If
    nScriptValue = nExpDmgHP * nLairPCT
End If

If nScriptValue > 0 Then
    If nScriptValue > 1000000000 Then
        sScriptValue = Format((nScriptValue / 1000000), "#,#M")
    ElseIf nScriptValue > 1000000 Then
        sScriptValue = Format((nScriptValue / 1000), "#,#K")
    Else
        sScriptValue = IIf(nScriptValue > 0, Format(RoundUp(nScriptValue), "#,#"), "0")
    End If
    
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Script Value"
    
    If nLairPCT < 1 Or nPossyPCT > 1 Then
        If nPossyPCT > 1 Then
            sTemp = sScriptValue & " calculated by (Exp/(((Dmg*2)+HP)*Possy%))"
            If nLairPCT < 1 Then sTemp = sTemp & "*Lair%"
            oLI.ListSubItems.Add (1), "Detail", sTemp
            
            Set oLI = DetailLV.ListItems.Add()
            oLI.Text = " "
            sTemp = "Calculation: ((" & nExp & " / ( ((" & nAvgDmg & " x 2) + " & tabMonsters.Fields("HP") & ") * " & nPossyPCT & ")) * 100)"
            If nLairPCT < 1 Then sTemp = sTemp & " * " & nLairPCT
            oLI.ListSubItems.Add (1), "Detail", sTemp
            
            Set oLI = DetailLV.ListItems.Add()
            oLI.Text = " "
            oLI.ListSubItems.Add (1), "Detail", "DMG+HP inflated by " & ((nPossyPCT * 100) - 100) & "% to account for mob's possy"
        Else
            sTemp = sScriptValue & " calculated from [Exp/((Dmg*2)+HP)] above"
            oLI.ListSubItems.Add (1), "Detail", sTemp
        End If
    Else
        If nNMRVer >= 1.83 Then
            sTemp = sScriptValue & " calculated by [Exp/((Dmg*2)+HP)]"
            oLI.ListSubItems.Add (1), "Detail", sTemp
        Else
            sTemp = "Same as [Exp/((Dmg*2)+HP)] above as there are no effects from monster possy or lair difficiency"
            oLI.ListSubItems.Add (1), "Detail", sTemp
        End If
    End If
    
    If nLairPCT < 1 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = " "
        oLI.ListSubItems.Add (1), "Detail", "Value reduced by " & ((1 - nLairPCT) * 100) & "% due insufficient number of mobs/lairs to outpace max regen"
    End If
    
    If nNMRVer >= 1.83 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = " "
        oLI.ListSubItems.Add (1), "Detail", "The final value is averaged amongst all of the monsters within the lairs that this monster spawns."
    End If
End If

If nNMRVer < 1.82 Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
    oLI.ListSubItems.Add (1), "Detail", "Note: This database is old and does not provide max regen per lair.  Calculations limited."
    oLI.ListSubItems(1).Bold = True
End If

done_scripting:
If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum

If tAvgLairInfo.nTotalLairs > 0 Or (nNMRVer >= 1.82 And nMonsterPossy(nMonsterNum) > 0) Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Lair Stats"
    oLI.Bold = True
    
    If tabMonsters.Fields("RegenTime") > 0 And tAvgLairInfo.nMobs > 0 Then
        oLI.ListSubItems.Add (1), "Detail", "Note: Mobs with regen time >0 are not included in lair stats"
    End If
End If

If tAvgLairInfo.nTotalLairs > 0 Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Total Lairs"
    oLI.ListSubItems.Add (1), "Detail", tAvgLairInfo.nTotalLairs
    
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "AVG # Mobs/Lair"
    If nMonsterSpawnChance(nMonsterNum) > 0 Then
        oLI.ListSubItems.Add (1), "Detail", nMonsterPossy(nMonsterNum) _
            & "  (" & (nMonsterSpawnChance(nMonsterNum) * 100) & "% avg chance for this monster to spawn/lair)"
    Else
        oLI.ListSubItems.Add (1), "Detail", tAvgLairInfo.nMaxRegen
    End If
    
    If tAvgLairInfo.nAvgDelay > 0 And tAvgLairInfo.nMaxRegen > 0 Then
        nTemp = Round((tAvgLairInfo.nAvgDelay * 60) / (tAvgLairInfo.nMaxRegen * 5))
        sTemp = " (lairs/regen period: " & nTemp & " @ " & tAvgLairInfo.nMaxRegen & " RTC"
        If tAvgLairInfo.nRTC > tAvgLairInfo.nMaxRegen Then
            nTemp2 = Round((tAvgLairInfo.nAvgDelay * 60) / (tAvgLairInfo.nRTC * 5))
            If nTemp2 <> nTemp Then sTemp = sTemp & " [" & nTemp2 & " @ " & tAvgLairInfo.nRTC & " RTC]"
        End If
        sTemp = sTemp & ")"
        
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "AVG Regen"
        If bGreaterMUD Then
            oLI.ListSubItems.Add (1), "Detail", (tAvgLairInfo.nAvgDelay - 1) & "m 30s" & sTemp
        Else
            oLI.ListSubItems.Add (1), "Detail", tAvgLairInfo.nAvgDelay & " minutes" & sTemp
        End If
    End If
    
    If tAvgLairInfo.nAvgWalk > 0 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "AVG Walk"
        oLI.ListSubItems.Add (1), "Detail", tAvgLairInfo.nAvgWalk & " rooms lair to lair"
    End If
    
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "AVG Exp"
    oLI.ListSubItems.Add (1), "Detail", PutCommas(tabMonsters.Fields("AvgLairExp")) & "  (" & PutCommas(Round(tabMonsters.Fields("AvgLairExp") / tAvgLairInfo.nMaxRegen)) & "/mob)"

    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "AVG HP"
    oLI.ListSubItems.Add (1), "Detail", PutCommas(tAvgLairInfo.nAvgHP) & "  (" & Round(tAvgLairInfo.nAvgHP / tAvgLairInfo.nMaxRegen) & "/mob)"
    
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "AVG AC/DR"
    oLI.ListSubItems.Add (1), "Detail", tAvgLairInfo.nAvgAC & "/" & tAvgLairInfo.nAvgDR
    
    If tAvgLairInfo.nAvgDodge > 0 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "AVG Dodge"
        If bUseCharacter And val(frmMain.lblInvenCharStat(10).Tag) > 0 Then
            oLI.ListSubItems.Add (1), "Detail", tAvgLairInfo.nAvgDodge _
                & " (" & Fix((tAvgLairInfo.nAvgDodge * 10) / Fix(val(frmMain.lblInvenCharStat(10).Tag) / 8)) & "% @ " & val(frmMain.lblInvenCharStat(10).Tag) & " accy)"
        Else
            oLI.ListSubItems.Add (1), "Detail", tAvgLairInfo.nAvgDodge
        End If
    End If
    
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "AVG BS Defense"
    oLI.ListSubItems.Add (1), "Detail", tAvgLairInfo.nAvgBSDefense
    If tAvgLairInfo.nAvgBSDefense <> 0 And bGlobalAttackBackstab Then
        oLI.ForeColor = &HC00000
        oLI.ListSubItems(1).ForeColor = &HC00000
    End If
    
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "AVG MR"
    oLI.ListSubItems.Add (1), "Detail", tAvgLairInfo.nAvgMR
    
    If tAvgLairInfo.nNumAntiMagic > 0 And tAvgLairInfo.nMobs > 0 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "AVG Anti-Magic"
        oLI.ListSubItems.Add (1), "Detail", Round((tAvgLairInfo.nNumAntiMagic / tAvgLairInfo.nMobs) * 100) & "% of mobs/lair"
    End If
    
    If (tAvgLairInfo.nAvgRCOL <> 0 Or tAvgLairInfo.nAvgRFIR <> 0 Or tAvgLairInfo.nAvgRSTO <> 0 Or tAvgLairInfo.nAvgRLIT <> 0 Or tAvgLairInfo.nAvgRWAT <> 0) Then
        sTemp = ""
        If tAvgLairInfo.nAvgRCOL <> 0 Then sTemp = AutoAppend(sTemp, "Cold: " & IIf(tAvgLairInfo.nAvgRCOL > 0, "+", "") & tAvgLairInfo.nAvgRCOL, ", ")
        If tAvgLairInfo.nAvgRFIR <> 0 Then sTemp = AutoAppend(sTemp, "Fire: " & IIf(tAvgLairInfo.nAvgRFIR > 0, "+", "") & tAvgLairInfo.nAvgRFIR, ", ")
        If tAvgLairInfo.nAvgRSTO <> 0 Then sTemp = AutoAppend(sTemp, "Stone: " & IIf(tAvgLairInfo.nAvgRSTO > 0, "+", "") & tAvgLairInfo.nAvgRSTO, ", ")
        If tAvgLairInfo.nAvgRLIT <> 0 Then sTemp = AutoAppend(sTemp, "Litng: " & IIf(tAvgLairInfo.nAvgRLIT > 0, "+", "") & tAvgLairInfo.nAvgRLIT, ", ")
        If tAvgLairInfo.nAvgRWAT <> 0 Then sTemp = AutoAppend(sTemp, "Water: " & IIf(tAvgLairInfo.nAvgRWAT > 0, "+", "") & tAvgLairInfo.nAvgRWAT, ", ")
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "AVG EL. Resist"
        oLI.ListSubItems.Add (1), "Detail", sTemp
    End If

    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "AVG DMG/mob"
    oLI.ListSubItems.Add (1), "Detail", PutCommas(tAvgLairInfo.nAvgDmg) & "/mob/round"
    If tAvgLairInfo.nDamageMitigated > 0 Then
        oLI.ListSubItems(1).Text = oLI.ListSubItems(1).Text & " (" & tAvgLairInfo.nDamageMitigated & " dmg/round mitigated)"
    End If
    
    If tAvgLairInfo.nRTK > 1 Or tAvgLairInfo.nRTC > 1 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "AVG Rounds"
        oLI.ListSubItems.Add (1), "Detail", tAvgLairInfo.nRTK & " RTK/mob, " & tAvgLairInfo.nRTC & " RTC/lair"
    End If
    
    If tAvgLairInfo.nRTC > 1 Or tAvgLairInfo.nAvgDmg <> tAvgLairInfo.nAvgDmgLair Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "AVG DMG/clear"
        oLI.ListSubItems.Add (1), "Detail", PutCommas(tAvgLairInfo.nAvgDmgLair) & "/round, " & PutCommas(Round(tAvgLairInfo.nAvgDmgLair * tAvgLairInfo.nRTC)) & "/clear (average damage taken, before any healing)"
    End If
    
    If tAvgLairInfo.nMagicLVL + tAvgLairInfo.nMaxMagicLVL > 0 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "Effective MagicLVL"
        sTemp = tAvgLairInfo.nMagicLVL & " (immune to attacks with < " & tAvgLairInfo.nMagicLVL & " magic/hitmagic)"
        If tAvgLairInfo.nMaxMagicLVL > tAvgLairInfo.nMagicLVL Then
            sTemp = sTemp & " - Max MagicLVL in lairs: " & tAvgLairInfo.nMaxMagicLVL
            oLI.ListSubItems.Add (1), "Detail", sTemp
            If nGlobalCharWeaponNumber(0) > 0 And (nGlobalAttackTypeMME = a1_PhysAttack Or nGlobalAttackTypeMME = a6_PhysBash Or nGlobalAttackTypeMME = a7_PhysSmash) Then
                If nWeaponMagic < tAvgLairInfo.nMaxMagicLVL Then
                    oLI.ForeColor = RGB(204, 0, 0)
                    oLI.ListSubItems(1).ForeColor = RGB(204, 0, 0)
                    oLI.Bold = True
                End If
            End If
            If nGlobalAttackTypeMME > a0_oneshot And bGlobalAttackBackstab = True Then
                If nBackstabWeaponMagic < tAvgLairInfo.nMaxMagicLVL Then
                    oLI.ForeColor = RGB(204, 0, 0)
                    oLI.ListSubItems(1).ForeColor = RGB(204, 0, 0)
                    oLI.Bold = True
                End If
            End If
        Else
            oLI.ListSubItems.Add (1), "Detail", sTemp
        End If
    End If
    
    If tAvgLairInfo.nSpellImmuLVL + tAvgLairInfo.nMaxSpellImmuLVL > 0 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "Effective SpellImmu"
        sTemp = tAvgLairInfo.nSpellImmuLVL & " (immune to spells <= level " & tAvgLairInfo.nSpellImmuLVL & ")"
        If tAvgLairInfo.nMaxSpellImmuLVL > tAvgLairInfo.nSpellImmuLVL Then
            sTemp = sTemp & " - Max SpellImmu in lairs: " & tAvgLairInfo.nMaxSpellImmuLVL
            oLI.ListSubItems.Add (1), "Detail", sTemp
            If nGlobalAttackSpellNum > 0 And (nGlobalAttackTypeMME = a2_Spell Or nGlobalAttackTypeMME = a3_SpellAny) Then
                If tSpellcast.nCastLevel > 0 And tSpellcast.nCastLevel <= tAvgLairInfo.nMaxSpellImmuLVL Then
                    oLI.ForeColor = RGB(204, 0, 0)
                    oLI.ListSubItems(1).ForeColor = RGB(204, 0, 0)
                    oLI.Bold = True
                End If
            End If
        Else
            oLI.ListSubItems.Add (1), "Detail", sTemp
        End If
    End If
    
    If (tAvgLairInfo.nNumUndeads > 0 And tAvgLairInfo.nMobs > 0) Or (tAvgLairInfo.nMobs > 0 And eAttackFlags And AR023_Undead) Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "Effective Undead"
        oLI.ListSubItems.Add (1), "Detail", Round((tAvgLairInfo.nNumUndeads / RoundUp(tAvgLairInfo.nMobs)) * 100) & "% of mobs/lair"
        If ((eAttackFlags And AR023_Undead) <> 0) And nGlobalAttackSpellNum > 0 And (nGlobalAttackTypeMME = a2_Spell Or nGlobalAttackTypeMME = a3_SpellAny) Then
            If Round(tAvgLairInfo.nNumUndeads / RoundUp(tAvgLairInfo.nMobs) * 100) < 100 Then
                oLI.ForeColor = RGB(204, 0, 0)
                oLI.ListSubItems(1).ForeColor = RGB(204, 0, 0)
                oLI.Bold = True
            End If
        End If
    End If
    
    If (tAvgLairInfo.nNumLiving > 0 And tAvgLairInfo.nNumLiving < tAvgLairInfo.nMobs) Or (tAvgLairInfo.nMobs > 0 And eAttackFlags And AR108_Living) Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "Effective Living"
        oLI.ListSubItems.Add (1), "Detail", Round((tAvgLairInfo.nNumLiving / RoundUp(tAvgLairInfo.nMobs)) * 100) & "% of mobs/lair"
        If ((eAttackFlags And AR108_Living) <> 0) And nGlobalAttackSpellNum > 0 And (nGlobalAttackTypeMME = a2_Spell Or nGlobalAttackTypeMME = a3_SpellAny) Then
            If Round(tAvgLairInfo.nNumLiving / RoundUp(tAvgLairInfo.nMobs) * 100) < 100 Then
                oLI.ForeColor = RGB(204, 0, 0)
                oLI.ListSubItems(1).ForeColor = RGB(204, 0, 0)
                oLI.Bold = True
            End If
        End If
    End If
    
    If (tAvgLairInfo.nNumAnimals > 0 And tAvgLairInfo.nMobs > 0) Or (tAvgLairInfo.nMobs > 0 And eAttackFlags And AR080_Animal) Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "Effective Animals"
        oLI.ListSubItems.Add (1), "Detail", Round((tAvgLairInfo.nNumAnimals / RoundUp(tAvgLairInfo.nMobs)) * 100) & "% of mobs/lair"
        If ((eAttackFlags And AR080_Animal) <> 0) And nGlobalAttackSpellNum > 0 And (nGlobalAttackTypeMME = a2_Spell Or nGlobalAttackTypeMME = a3_SpellAny) Then
            If Round(tAvgLairInfo.nNumAnimals / RoundUp(tAvgLairInfo.nMobs) * 100) < 100 Then
                oLI.ForeColor = RGB(204, 0, 0)
                oLI.ListSubItems(1).ForeColor = RGB(204, 0, 0)
                oLI.Bold = True
            End If
        End If
    End If
    
    If InStr(1, tAvgLairInfo.sMobList, ",", vbTextCompare) > 0 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "Other Lair Mobs"
        sArr() = Split(tAvgLairInfo.sMobList, ",")
        y = 0
        For x = 0 To UBound(sArr())
            If val(sArr(x)) <> nMonsterNum Then
                If y > 0 Then
                    Set oLI = DetailLV.ListItems.Add()
                    oLI.Text = ""
                End If
                oLI.ListSubItems.Add (1), "Detail", GetMonsterName(sArr(x), bHideRecordNumbers)
                tabMonsters.Seek "=", nMonsterNum
                oLI.Tag = "monster"
                oLI.ListSubItems(1).Tag = sArr(x)
                y = y + 1
                If y > 14 And UBound(sArr()) > 20 Then
                    Set oLI = DetailLV.ListItems.Add()
                    oLI.Text = ""
                    oLI.ListSubItems.Add 1, , "... plus " & (UBound(sArr()) - y) & " more."
                    x = UBound(sArr()) + 1
                End If
            End If
        Next x
    End If
    
ElseIf nNMRVer >= 1.82 And nMonsterPossy(nMonsterNum) > 0 Then
'    Set oLI = DetailLV.ListItems.Add()
'    oLI.Text = ""
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Avg # Mobs / Lair"
    oLI.ListSubItems.Add (1), "Detail", nMonsterPossy(nMonsterNum)
    If nMonsterSpawnChance(nMonsterNum) > 0 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "Avg Spawn Chance"
        oLI.ListSubItems.Add (1), "Detail", (nMonsterSpawnChance(nMonsterNum) * 100) & "%  (the chance for this monster to spawn per lair)"
    End If
End If

If nNMRVer < 1.83 Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = " "
    oLI.ListSubItems.Add (1), "Detail", "This database is out of date and unable to use all features."
    oLI.ListSubItems(1).ForeColor = &H80000015
End If

Set oLI = DetailLV.ListItems.Add()
oLI.Text = ""

':[END OF DAMAGE OUT + SCRIPTING / LAIR STUFF]
If bMobPrintCharDamageOutFirst And nNMRVer >= 1.83 And frmMain.optMonsterFilter(1).Value = True Then GoTo mob_attacks:
done_attacks:

If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum

If Not frmMain.bDontLookupMonRegen Then
    If Len(tabMonsters.Fields("Summoned By")) > 4 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "Spawns via ..."
        oLI.Bold = True
        Call frmMain.LookUpMonsterRegen(nMonsterNum, False, DetailLV, nLookupLimit)
    End If
End If

out:
Set oLI = Nothing
On Error Resume Next
Exit Sub
error:
Call HandleError("PullMonsterDetail")
Resume out:
End Sub

'----------------------------------------------------------------------
' Compute the average of non-zero elements in an array
'----------------------------------------------------------------------
Public Function CalcAverageNonZero(ByRef arrData() As Double) As Double
On Error GoTo error:

Dim sum As Double, cnt As Long
Dim i As Long
For i = LBound(arrData) To UBound(arrData)
    If arrData(i) <> 0 Then
        sum = sum + arrData(i)
        cnt = cnt + 1
    End If
Next i
If cnt > 0 Then
    CalcAverageNonZero = sum / cnt
Else
    CalcAverageNonZero = 0
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalcAverageNonZero")
Resume out:
End Function

Private Sub AddMonsterDamageOutText(ByRef DetailLV As ListView, ByVal sHeader As String, ByVal sDetail As String, Optional ByVal sDetail2 As String, _
    Optional ByVal nDamageOut As Long = -9999, Optional ByVal nMobHealth As Long, Optional ByVal nMobHPRegen As Long, _
    Optional ByVal nMobDamage As Long = -1, Optional ByVal nCharHealth As Long, Optional ByVal sDefenseText As String, _
    Optional ByVal nNumMobs As Double = 1, Optional ByVal nOOM As Integer, Optional ByVal nOverrideRTK As Double, _
    Optional ByVal nSurpriseDamageOut As Double = -9999, Optional ByVal sSurpriseDamageOut As String, _
    Optional ByVal nMinDamageOut As Long = -9999)
On Error GoTo error:
Dim oLI As ListItem, tCombatRounds As tCombatRoundInfo, bUseCharacter As Boolean, sExtText As String

If frmMain.chkGlobalFilter.Value = 1 Then bUseCharacter = True

tCombatRounds = CalcCombatRounds(nDamageOut, nMobHealth, nMobDamage, nCharHealth, nMobHPRegen, nNumMobs, nOverrideRTK, nSurpriseDamageOut, nMinDamageOut)

Set oLI = DetailLV.ListItems.Add()
oLI.Bold = True
oLI.Text = sHeader
oLI.ListSubItems.Add (1), "Detail", sDetail

If InStr(1, sDetail, "immune:", vbTextCompare) > 0 Then
    oLI.ForeColor = RGB(204, 0, 0)
    oLI.ListSubItems(1).ForeColor = RGB(204, 0, 0)
End If

If nSurpriseDamageOut > 0 And Len(sSurpriseDamageOut) > 0 Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
    oLI.ListSubItems.Add (1), "Detail", sSurpriseDamageOut
End If

If Not sDetail2 = "" Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
    oLI.ListSubItems.Add (1), "Detail", sDetail2
End If

If Len(tCombatRounds.sRTK & tCombatRounds.sRTD) > 0 Then
    
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
    If Len(tCombatRounds.sRTK) > 0 And Len(tCombatRounds.sRTD) > 0 Then
        oLI.ListSubItems.Add (1), "Detail", AutoAppend(tCombatRounds.sRTK, tCombatRounds.sRTD, " ") & tCombatRounds.sSuccess & sDefenseText
    Else
        oLI.ListSubItems.Add (1), "Detail", tCombatRounds.sRTK & tCombatRounds.sRTD & tCombatRounds.sSuccess
    End If
    
    If (tCombatRounds.nRTD > 0 And tCombatRounds.nRTK > 1) And tCombatRounds.nSuccess < 70 Then
        'nRTD > 0 And (nRTK = 0 Or (nRTK * 1.1) > (nRTD * 0.9))
        oLI.ListSubItems(1).ForeColor = &HC0&
        'oLI.ListSubItems(1).Bold = True
    ElseIf tCombatRounds.nRTD >= 1 And tCombatRounds.nSuccess >= 95 Then
        oLI.ListSubItems(1).ForeColor = &H8000&
    End If
End If

If nOOM > 0 And nOOM < 100 Then
    If bUseCharacter And (val(frmMain.lblCharBless.Caption) > 0 Or nGlobalAttackHealCost > 0) Then sExtText = " (with current heals/bless set)"
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
    oLI.ListSubItems.Add (1), "Detail", "OOM in " & nOOM & " rounds" & sExtText
    If (nOOM * 0.9) < (tCombatRounds.nRTK * 1.1) Then oLI.ListSubItems(1).ForeColor = &HC0&
End If

Set oLI = DetailLV.ListItems.Add()
oLI.Text = ""

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("AddMonsterDamageOutText")
Resume out:
End Sub

Public Sub PullShopDetail(nShopNum As Long, DetailLV As ListView, _
    DetailTB As TextBox, lvAssigned As ListView, ByVal nCharm As Integer, ByVal bShowSell As Boolean)

On Error GoTo error:

Dim sStr As String, x As Integer, nRegenTime As Integer, sRegenTime As String
Dim oLI As ListItem, tCostType As tItemValue, nCopper As Currency, sCopper As String
Dim nCharmMod As Double, sCharmMod As String
Dim nReducedCoin As Currency, sReducedCoin As String

DetailLV.ListItems.clear
If bStartup Then Exit Sub

tabShops.Index = "pkShops"
tabShops.Seek "=", nShopNum
If tabShops.NoMatch = True Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Shop not found"
    DetailTB.Text = "Shop not found"
    Set oLI = Nothing
    tabShops.MoveFirst
    Exit Sub
End If


If nCharm > 0 Or bShowSell Then
    If bShowSell And Not tabShops.Fields("ShopType") = 8 Then
        nCharmMod = Fix(nCharm / 2) + 25
        sCharmMod = nCharmMod & "% cost pre-markup"
    Else
        nCharmMod = 1 - ((Fix(nCharm / 5) - 10) / 100)
        If nCharmMod > 1 Then
            sCharmMod = Abs(1 - CCur(nCharmMod)) * 100 & "% Markup"
        ElseIf nCharmMod < 1 Then
            sCharmMod = val(1 - CCur(nCharmMod)) * 100 & "% Discount"
        Else
            sCharmMod = "Retail Value"
        End If
    End If
End If

    
If tabShops.Fields("ShopType") = 8 Then 'Case 8: GetShopType = "Training"
    If Not DetailLV.ColumnHeaders.Count = 2 Then
        DetailLV.ColumnHeaders.clear
        DetailLV.ColumnHeaders.Add 1, "Level", "LVL", 1000, lvwColumnLeft
        DetailLV.ColumnHeaders.Add 2, "Cost", "Cost", 4000, lvwColumnLeft
    End If
    
    frmMain.chkShopShowCharm(0).Enabled = False
    frmMain.chkShopShowCharm(1).Enabled = False
    
    frmMain.lblCharmMod.Caption = ""
    
    For x = tabShops.Fields("MinLVL") To tabShops.Fields("MaxLVL")
        sReducedCoin = ""
        nReducedCoin = 0
        
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = x
        nCopper = CalcMoneyRequiredToTrain(x - 1, tabShops.Fields("Markup%")) '* nCharmMod
        If nCopper < 0 Then nCopper = 0
        
        nCopper = Round(nCopper)
        
        sReducedCoin = "Copper"
        If nCopper >= 10000000 Then
            nReducedCoin = nCopper / 1000000
            sReducedCoin = "Runic"
        ElseIf nCopper >= 100000 Then
            nReducedCoin = nCopper / 10000
            sReducedCoin = "Platinum"
        ElseIf nCopper >= 1000 Then
            nReducedCoin = nCopper / 100
            sReducedCoin = "Gold"
        ElseIf nCopper >= 100 Then
            nReducedCoin = nCopper / 10
            sReducedCoin = "Silver"
        End If
        If nReducedCoin > 0 Then nReducedCoin = Round(nReducedCoin, 2)
        
        sCopper = Format(nCopper, "##,##0.00")
        If Right(sCopper, 3) = ".00" Then sCopper = Left(sCopper, Len(sCopper) - 3)
        
        If nReducedCoin = 0 Then
            oLI.ListSubItems.Add (1), "Cost", IIf(nCopper <= 0, "Free", sCopper & " Copper")
        Else
            sStr = Format(nReducedCoin, "##,##0.00")
            If Right(sStr, 3) = ".00" Then sStr = Left(sStr, Len(sStr) - 3)
            oLI.ListSubItems.Add (1), "Cost", Format(nCopper, "#,#") & " copper (" & sStr & " " & sReducedCoin & ")"
        End If
        
        oLI.ListSubItems(1).Tag = nCopper
    Next x
    
Else
    frmMain.chkShopShowCharm(0).Enabled = True
    frmMain.chkShopShowCharm(1).Enabled = True
    
    frmMain.lblCharmMod.Caption = sCharmMod
    
    If Not DetailLV.ColumnHeaders.Count = 5 Then
        DetailLV.ColumnHeaders.clear
        DetailLV.ColumnHeaders.Add 1, "Number", "#", 700, lvwColumnLeft
        DetailLV.ColumnHeaders.Add 2, "Name", "Name", 2000, lvwColumnCenter
        DetailLV.ColumnHeaders.Add 3, "Max", "Max", 550, lvwColumnCenter
        DetailLV.ColumnHeaders.Add 4, "Regen", "Regen", 1650, lvwColumnCenter
        'DetailLV.ColumnHeaders.Add 5, "Rgn%", "Rgn%", 600, lvwColumnCenter
        'DetailLV.ColumnHeaders.Add 6, "Rgn#", "Rgn#", 600, lvwColumnCenter
        DetailLV.ColumnHeaders.Add 5, "Cost", "Cost", 5000, lvwColumnLeft
    End If
    
    For x = 0 To 19
        sReducedCoin = ""
        nReducedCoin = 0
        
        If tabShops.Fields("Item-" & x) = 0 Then GoTo skip:
        
        Set oLI = DetailLV.ListItems.Add()
        
        oLI.Text = tabShops.Fields("Item-" & x)
        oLI.ListSubItems.Add (1), "Name", GetItemName(tabShops.Fields("Item-" & x), True)
        oLI.ListSubItems.Add (2), "Max", tabShops.Fields("Max-" & x)
        
        sRegenTime = ""
        If tabShops.Fields("Time-" & x) > 0 And tabShops.Fields("%-" & x) > 0 _
            And tabShops.Fields("Amount-" & x) > 0 Then
            
            nRegenTime = tabShops.Fields("Time-" & x)
            If nRegenTime >= (60 * 24) Then 'days
                sRegenTime = sRegenTime & Int(nRegenTime / (60 * 24)) & "d"
                nRegenTime = nRegenTime - (Int(nRegenTime / (60 * 24)) * (60 * 24))
            End If
            If nRegenTime >= 60 Then 'hours
                sRegenTime = sRegenTime & Int(nRegenTime / 60) & "h"
                nRegenTime = nRegenTime - (Int(nRegenTime / 60) * 60)
            End If
            If nRegenTime > 0 Then 'minutes
                sRegenTime = sRegenTime & nRegenTime & "m"
            End If
            
            sRegenTime = tabShops.Fields("%-" & x) & "% for " & tabShops.Fields("Amount-" & x) _
                & " per " & sRegenTime
        Else
            sRegenTime = "no regen"
        End If
        
        oLI.ListSubItems.Add (3), "Time", sRegenTime
        
        tCostType = GetItemValue(tabShops.Fields("Item-" & x), nCharm, tabShops.Fields("Markup%"))
        
        If bShowSell Then
            oLI.ListSubItems.Add (4), "Cost", tCostType.sFriendlySell
            oLI.ListSubItems(4).Tag = tCostType.nCopperSell
        Else
            oLI.ListSubItems.Add (4), "Cost", tCostType.sFriendlyBuy
            oLI.ListSubItems(4).Tag = tCostType.nCopperBuy
        End If
skip:
    Next
End If


sStr = "Levels: " & tabShops.Fields("MinLVL") & " to " & tabShops.Fields("MaxLVL")
sStr = sStr & " -- Markup: " & tabShops.Fields("Markup%") & "%"

If Not tabShops.Fields("ClassRest") = 0 Then
    sStr = sStr & " -- Class: " & GetClassName(tabShops.Fields("ClassRest"))
End If
    
DetailTB.Text = sStr

Call GetLocations(tabShops.Fields("Assigned To"), lvAssigned)


out:
On Error Resume Next
Set oLI = Nothing
Exit Sub
error:
Call HandleError("PullShopDetail")
Resume out:
End Sub


Public Sub PullSpellDetail(nSpellNum As Long, DetailTB As TextBox, LocationLV As ListView)
On Error GoTo error:
Dim sSpellDetail As String, sRemoves As String, sArr() As String, x As Integer ', y As Integer
Dim bCalcCombat As Boolean, bUseCharacter As Boolean
Dim tSpellcast As tSpellCastValues, bBR As Boolean
Dim nCastLVL As Long, sSpellEQ As String
Dim tChar As tCharacterProfile, sBonusDamage As String

DetailTB.Text = ""
If bStartup Then Exit Sub

tabSpells.Index = "pkSpells"
tabSpells.Seek "=", nSpellNum
If tabSpells.NoMatch = True Then
    DetailTB.Text = "spell not found"
    tabSpells.MoveFirst
    Exit Sub
End If

If frmMain.chkSpellOptions(0).Value = 1 And val(frmMain.txtSpellOptions(0).Text) > 0 Then bCalcCombat = True
If frmMain.chkGlobalFilter.Value = 1 And val(frmMain.txtGlobalLevel(1).Text) > 0 Then bUseCharacter = True

LocationLV.ListItems.clear

If bUseCharacter Then nCastLVL = val(frmMain.txtGlobalLevel(1).Text)
Call PopulateCharacterProfile(tChar, False, True)

If bCalcCombat Then
    tSpellcast = CalculateSpellCast(tChar, nSpellNum, nCastLVL, _
            val(frmMain.txtSpellOptions(0).Text), IIf(frmMain.chkSpellOptions(2).Value = 1, True, False))
Else
    tSpellcast = CalculateSpellCast(tChar, nSpellNum, nCastLVL)
End If
nCastLVL = tSpellcast.nCastLevel

'If tChar.nSpellDmgBonus > 0 Then sBonusDamage = " (+" & tChar.nSpellDmgBonus & "% spell damage)"

If bCalcCombat And (tSpellcast.bDoesDamage Or tSpellcast.bDoesHeal) And Len(tSpellcast.sAvgRound) > 0 Then
    sSpellDetail = AutoAppend(sSpellDetail, tSpellcast.sAvgRound & sBonusDamage, vbCrLf)
    bBR = True
ElseIf tSpellcast.nDuration > 1 And Len(tSpellcast.sLVLincreases) > 0 And nCastLVL > 1 And bUseCharacter = False Then
    bQuickSpell = True
    sSpellDetail = AutoAppend(sSpellDetail, PullSpellEQ(True, nCastLVL, , , , , , , , tSpellcast.nMinCast, tSpellcast.nMaxCast)) & sBonusDamage
    bQuickSpell = False
    bBR = True
End If

If (bCalcCombat Or Not bUseCharacter) And Len(tSpellcast.sMMA) > 0 And tSpellcast.nMinCast > 0 _
    And (tSpellcast.nMinCast <> tSpellcast.nMaxCast Or tSpellcast.nMaxCast <> tSpellcast.nAvgCast) Then
    sSpellDetail = AutoAppend(sSpellDetail, tSpellcast.sMMA, vbCrLf)
    bBR = True
End If

'(tabSpells.Fields("ManaCost") * tSpellCast.nNumCasts) > (Val(frmMain.lblCharManaRate.Tag) - Val(frmMain.lblCharBless.Caption)) _
    And Val(frmMain.lblCharMaxMana.Tag) > 0 And (tSpellCast.bDoesDamage Or tSpellCast.bDoesHeal) And tSpellCast.nDuration = 1
If bCalcCombat And bUseCharacter And tSpellcast.nOOM > 0 Then
    'reads better here when calculating combat (also below)
    sSpellDetail = AutoAppend(sSpellDetail, "OOM in " & tSpellcast.nOOM & " rounds", vbCrLf)
    If tSpellcast.nDuration > 1 And (val(frmMain.lblCharBless.Caption) > 0 Or nGlobalAttackHealCost > 0) Then
        sSpellDetail = sSpellDetail & " (after " & (tSpellcast.nOOM \ tSpellcast.nDuration) & " casts, with current heals/bless set)"
    ElseIf (val(frmMain.lblCharBless.Caption) > 0 Or nGlobalAttackHealCost > 0) Then
        sSpellDetail = sSpellDetail & " (with current heals/bless set)"
    ElseIf tSpellcast.nDuration > 1 Then
        sSpellDetail = sSpellDetail & " (after " & (tSpellcast.nOOM \ tSpellcast.nDuration) & " casts)"
    End If
    If tSpellcast.nDuration > 1 And tSpellcast.nCastChance > 0 And tSpellcast.nCastChance < 100 Then
        sSpellDetail = sSpellDetail & " @ " & tSpellcast.nCastChance & "% chance to cast"
    End If
    bBR = True
End If

If bBR Then sSpellDetail = sSpellDetail & vbCrLf: bBR = False

If bUseCharacter Then
    sSpellEQ = PullSpellEQ(True, val(frmMain.txtGlobalLevel(0).Text), , LocationLV, , , , , , tSpellcast.nMinCast, tSpellcast.nMaxCast)
Else
    sSpellEQ = PullSpellEQ(False, , , LocationLV, , , , , , tSpellcast.nMinCast, tSpellcast.nMaxCast)
End If
If Not tabSpells.Fields("Number") = nSpellNum Then tabSpells.Seek "=", nSpellNum

If InStr(1, sSpellEQ, " -- RemovesSpells", vbTextCompare) > 0 Then
    sRemoves = Trim(Mid(sSpellEQ, InStr(1, sSpellEQ, " -- RemovesSpells(", vbTextCompare) + 4, Len(sSpellEQ)))
    sSpellEQ = Left(sSpellEQ, Len(sSpellEQ) - Len(sRemoves) - 4)
    sRemoves = Trim(Mid(sRemoves, Len(" -- RemovesSpells(") - 3, Len(sRemoves) - Len(" -- RemovesSpells(") + 3))
End If

If Not tabSpells.Fields("Cap") = 0 Then
    If bUseCharacter Then
        sSpellEQ = "LVL Cap: " & tabSpells.Fields("Cap") & " " & sSpellEQ
    Else
        sSpellEQ = "LVL Cap: " & tabSpells.Fields("Cap") & ", " & sSpellEQ
    End If
End If

If Len(sSpellEQ) > 0 Then sSpellDetail = AutoAppend(sSpellDetail, sSpellEQ, vbCrLf)
If tSpellcast.nManaCost > 0 And tSpellcast.nNumCasts > 1 Then
    sSpellDetail = AutoAppend(sSpellDetail, " (" & Fix(tSpellcast.nManaCost / tSpellcast.nNumCasts) & " mana/ea)", "")
End If

If bUseCharacter And Len(tSpellcast.sLVLincreases) > 0 Then
    'If Len(sSpellEQ) > 0 Then sSpellDetail = sSpellDetail & vbCrLf
    sSpellDetail = AutoAppend(sSpellDetail, tSpellcast.sLVLincreases, vbCrLf)
End If

If bCalcCombat = False And bUseCharacter And tSpellcast.nOOM > 0 Then
    'reads better here when NOT calculating combat (also above)
    sSpellDetail = AutoAppend(sSpellDetail, "OOM in " & tSpellcast.nOOM & " rounds", vbCrLf)
    If (val(frmMain.lblCharBless.Caption) > 0) Then sSpellDetail = sSpellDetail & " (with current bless set)"
End If

If Not tabSpells.Fields("Number") = nSpellNum Then tabSpells.Seek "=", nSpellNum

sSpellDetail = sSpellDetail & vbCrLf & vbCrLf & "Target: " & GetSpellTargets(tabSpells.Fields("Targets"))
If tabSpells.Fields("Diff") <> 0 And tabSpells.Fields("Diff") < 200 Then sSpellDetail = AutoAppend(sSpellDetail, "Difficulty: " & tabSpells.Fields("Diff"))
sSpellDetail = AutoAppend(sSpellDetail, "Attack Type: " & SpellAttackTypeEnum(tabSpells.Fields("AttType")))

If nNMRVer >= 1.8 Then
    If tabSpells.Fields("TypeOfResists") = 1 Then
        sSpellDetail = sSpellDetail & ", Fully-Resistable by Anti-Magic Only"
        If tSpellcast.nFullResistChance > 0 Then sSpellDetail = sSpellDetail & " (" & tSpellcast.nFullResistChance & "%)"
    ElseIf tabSpells.Fields("TypeOfResists") = 2 Then
        sSpellDetail = sSpellDetail & ", Fully-Resistable by All"
        If tSpellcast.nFullResistChance > 0 Then sSpellDetail = sSpellDetail & " (" & tSpellcast.nFullResistChance & "%)"
    Else
        sSpellDetail = sSpellDetail & ", Can Not be Fully-Resisted"
    End If
End If

'If tChar.nSpellDmgBonus > 0 And (tSpellcast.bDoesDamage Or (tSpellcast.bDoesHeal And bGreaterMUD)) Then
'    sSpellDetail = AutoAppend(sSpellDetail, "+" & tChar.nSpellDmgBonus & "% dmg")
'    If bGreaterMUD Then sSpellDetail = sSpellDetail & "/heal"
'End If

If nNMRVer >= 1.7 Then
    If Len(tabSpells.Fields("Classes")) > 2 And Not tabSpells.Fields("Classes") = "(*)" Then
        
        sSpellDetail = sSpellDetail & vbCrLf & "Class Restricted (via learning method): "
        
        sArr() = StringOfNumbersToArray(tabSpells.Fields("Classes"))
        For x = 0 To UBound(sArr())
            If x > 0 Then sSpellDetail = sSpellDetail & ", "
            sSpellDetail = sSpellDetail & GetClassName(sArr(x))
        Next x
    End If
End If

If Not sRemoves = "" Then sSpellDetail = sSpellDetail & vbCrLf & "Removes: " & sRemoves

DetailTB.Text = sSpellDetail

Call GetLocations(tabSpells.Fields("Learned From"), LocationLV, True, "(learn) ")
If Not tabSpells.Fields("Number") = nSpellNum Then tabSpells.Seek "=", nSpellNum
Call GetLocations(tabSpells.Fields("Casted By"), LocationLV, True)

out:
On Error Resume Next
bQuickSpell = False
Exit Sub
error:
Call HandleError("PullSpellDetail")
Resume out:
End Sub

Public Function CalcRoundsToOOM(ByVal ManaCost As Double, ByVal MaxMana As Long, ByVal RegenRate As Double, _
    Optional ByVal nCastChance As Integer, Optional ByVal nDuration As Long = 1) As Integer
On Error GoTo error:
Dim rounds As Long, CurrentMana As Double
Dim RoundsPerRegen As Long, regenBetween As Long, nRoundsDuration As Integer ', RoundsPerFail As Long
Dim ReturnOnFail As Long, FailAccumulation As Long, bCastAttempt As Boolean

If ManaCost > MaxMana Then Exit Function 'never cast

Const RoundSecs As Long = 5
Const RegenSecs As Long = 30

RoundsPerRegen = RegenSecs \ RoundSecs

If nDuration < 1 Then nDuration = 1
If nCastChance <= 0 Then nCastChance = 100
If nCastChance < 100 And ManaCost > 0 Then ReturnOnFail = Fix(ManaCost / 2)

If nDuration > 1 Then
    regenBetween = (nDuration \ RoundsPerRegen) * RegenRate
    If regenBetween >= (ManaCost + ((ManaCost / 2) * (1 - (nCastChance / 100)))) Then Exit Function 'never oom
Else
    If RegenRate >= (ManaCost * RoundsPerRegen) Then Exit Function  'never oom
End If

CurrentMana = MaxMana
rounds = 0
nRoundsDuration = 0

Do While CurrentMana >= ManaCost And rounds < 999
    rounds = rounds + 1
    If nDuration > 1 Then nRoundsDuration = nRoundsDuration + 1
    
    bCastAttempt = False
    If nDuration = 1 Or rounds = 1 Or nRoundsDuration = 1 Then
        bCastAttempt = True
    ElseIf (nRoundsDuration Mod nDuration) = 0 Then
        bCastAttempt = True
    End If
    
    If bCastAttempt Then CurrentMana = CurrentMana - ManaCost
    
    If nCastChance < 100 And bCastAttempt Then
        FailAccumulation = FailAccumulation + (100 - nCastChance)
        If FailAccumulation >= 100 - ((100 - nCastChance) / 2) Then
            CurrentMana = CurrentMana + ReturnOnFail
            FailAccumulation = FailAccumulation - 100
            If nDuration > 1 Then nRoundsDuration = 0
        End If
    End If
    
    If (rounds Mod RoundsPerRegen) = 0 Then
        CurrentMana = CurrentMana + RegenRate
        If CurrentMana > MaxMana Then
            CurrentMana = MaxMana
            If rounds > 200 Then GoTo out: 'assume won't run out
        End If
    End If
Loop

If rounds = 999 Then rounds = 0
CalcRoundsToOOM = rounds

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalcRoundsToOOM")
Resume out:
End Function

Public Function CalcManaRecoveryRounds(ByVal nMaxMana As Long, ByVal nRegenRate As Long, _
    Optional ByVal nMeditateRate As Long, Optional ByVal nPercentage As Integer = 95) As Long
On Error GoTo error:

'--- MP DEBUG HEADER ----------------------------------------------------
If bDebugExpPerHour Then
    Debug.Print "MPDBG --- CalcManaRecoveryRounds ---"
    Debug.Print "  nMaxMana=" & nMaxMana & _
                "; nRegenRate=" & nRegenRate & _
                "; nMeditateRate=" & nMeditateRate & _
                "; nPercentage=" & nPercentage
End If
'----------------------------------------------------------------------'

' Returns 5s rounds needed to go from 0 MP to target MP **while meditating**.
' Cadence:
'   - Regen tick: every 30s, amount = nRegenRate  (base + manaRegenBonus)
'   - Meditate tick: every 10s, amount = nMeditateRate (base only)
'
' Caller later converts rounds to a time fraction and also applies HP-rest overlap.

Const RoundSecs As Long = 5
Const RegenTick As Long = 30
Const MediTick  As Long = 10

Dim targetMana As Long, Mana As Long
Dim rounds As Long, t As Long
Dim nextRegen As Long, nextMedi As Long

If nPercentage > 0 Then
    targetMana = Fix(nMaxMana * (nPercentage / 100#))
Else
    targetMana = nMaxMana
End If
If targetMana <= 0 Then CalcManaRecoveryRounds = 0: Exit Function

Mana = 0
rounds = 0
t = 0
nextRegen = RegenTick
nextMedi = MediTick

Do While Mana < targetMana And rounds < 999
    rounds = rounds + 1
    t = t + RoundSecs

    ' 30s regen ticks (can fire multiple times if RoundSecs > 30, guarded by loop)
    Do While t >= nextRegen
        Mana = Mana + nRegenRate
        nextRegen = nextRegen + RegenTick
        If Mana >= targetMana Then Exit Do
    Loop

    ' 10s meditate ticks
    If nMeditateRate > 0 Then
        Do While t >= nextMedi
            Mana = Mana + nMeditateRate
            nextMedi = nextMedi + MediTick
            If Mana >= targetMana Then Exit Do
        Loop
    End If
Loop

If Mana < targetMana Then
    ' Could not reach target with given rates within safety cap.
    CalcManaRecoveryRounds = 999
Else
    CalcManaRecoveryRounds = rounds
End If

If bDebugExpPerHour Then
    Debug.Print "MPDBG --- CalcManaRecoveryRounds ---"
    Debug.Print "  targetMana=" & targetMana _
        & "; nRegenRate=" & nRegenRate _
        & "; nMeditateRate=" & nMeditateRate
    Debug.Print "  roundsNeeded=" & rounds & " (=" & rounds * RoundSecs & "s)"
End If

out:
Exit Function
error:
Call HandleError("CalcManaRecoveryRounds")
Resume out:
End Function


Public Function Get_Enc_Ratio(nENC As Long, nVal1 As Long, Optional nVal2 As Long) As Currency
Dim nTotal As Long

nTotal = nVal1 + nVal2

If nTotal > 0 Then
    If nENC < 1 Then
        Get_Enc_Ratio = nTotal '* 10
    Else
        Get_Enc_Ratio = Round(nTotal / nENC, 4) * 100 'Round((nTotal / 10) / nENC, 5) * 1000
    End If
Else
    Get_Enc_Ratio = 0
End If

End Function
Public Sub AddArmour2LV(lv As ListView, Optional AddToInven As Boolean, Optional nAbility As Integer)
On Error GoTo error:
Dim oLI As ListItem, x As Integer, sName As String, nAbilityVal As Integer
Dim sAbil As String

sName = tabItems.Fields("Name")
If sName = "" Then GoTo skip:

Set oLI = lv.ListItems.Add()
oLI.Text = tabItems.Fields("Number")

oLI.ListSubItems.Add (1), "Name", tabItems.Fields("Name")
oLI.ListSubItems.Add (2), "Worn", GetWornType(tabItems.Fields("Worn"))
oLI.ListSubItems(2).Tag = tabItems.Fields("Worn")
oLI.ListSubItems.Add (3), "Armr Type", GetArmourType(tabItems.Fields("ArmourType"))
oLI.ListSubItems.Add (4), "Level", 0
oLI.ListSubItems.Add (5), "Enc", tabItems.Fields("Encum")
oLI.ListSubItems.Add (6), "AC", (tabItems.Fields("ArmourClass") / 10) & "/" & (tabItems.Fields("DamageResist") / 10)
oLI.ListSubItems(6).Tag = tabItems.Fields("ArmourClass") + tabItems.Fields("DamageResist")

oLI.ListSubItems.Add (7), "Acc", tabItems.Fields("Accy")
oLI.ListSubItems.Add (8), "Crits", 0
oLI.ListSubItems.Add (9), "Limit", tabItems.Fields("Limit")

oLI.ListSubItems.Add (10), "AC/Enc", _
    Get_Enc_Ratio(tabItems.Fields("Encum"), tabItems.Fields("ArmourClass"), tabItems.Fields("DamageResist"))
    
For x = 0 To 19
    Select Case tabItems.Fields("Abil-" & x)
        Case 0:
        Case 58: ' crits
            oLI.ListSubItems(8).Text = tabItems.Fields("AbilVal-" & x)
            
        Case 135: 'min level
            oLI.ListSubItems(4).Text = tabItems.Fields("AbilVal-" & x)
        
        Case 22, 105, 106: 'acc
            oLI.ListSubItems(7).Text = val(oLI.ListSubItems(7).Text) + tabItems.Fields("AbilVal-" & x)
    End Select
    
    If nAbility > 0 And tabItems.Fields("Abil-" & x) = nAbility Then
        nAbilityVal = tabItems.Fields("AbilVal-" & x)
    ElseIf nAbility = 0 And tabItems.Fields("Abil-" & x) > 0 Then
        'sAbil = AutoAppend(sAbil, GetAbilityStats(tabItems.Fields("Abil-" & x), tabItems.Fields("AbilVal-" & x)))
    End If
Next x

If nAbility > 0 Then
    oLI.ListSubItems.Add (11), "Ability", nAbilityVal
ElseIf Len(sAbil) > 0 Then
    'oLI.ListSubItems.Add (11), "Ability", sAbil
Else
    oLI.ListSubItems.Add (11), "Ability", ""
End If

If AddToInven Then Call frmMain.InvenAddEquip(tabItems.Fields("Number"), sName, tabItems.Fields("ItemType"), tabItems.Fields("Worn"))
    
skip:
Set oLI = Nothing


out:
On Error Resume Next
Exit Sub
error:
Call HandleError("AddArmour2LV")
Resume out:
End Sub
Public Sub AddOtherItem2LV(lv As ListView)

On Error GoTo error:

Dim oLI As ListItem

If tabItems.Fields("Name") = "" Then GoTo skip:

Set oLI = lv.ListItems.Add()
oLI.Text = tabItems.Fields("Number")

oLI.ListSubItems.Add (1), "Name", tabItems.Fields("Name")
oLI.ListSubItems.Add (2), "Type", GetItemType(tabItems.Fields("ItemType"))
oLI.ListSubItems.Add (3), "Enc", tabItems.Fields("Encum")
oLI.ListSubItems.Add (4), "Limit", tabItems.Fields("Limit")

'For x = 0 To 9
'    If Not tabItems.Fields("Abil-" & x).Value = 0 Then
'        If sAbil <> "" Then sAbil = sAbil & ", "
'        sAbil = sAbil & GetAbilityStats(tabItems.Fields("Abil-" & x), tabItems.Fields("AbilVal-" & x))
'        If Right(sAbil, 2) = ", " Then sAbil = Left(sAbil, Len(sAbil) - 2)
'    End If
'Next

'oLI.ListSubItems.Add (5), "Abils", sAbil

skip:
Set oLI = Nothing


out:
On Error Resume Next
Exit Sub
error:
Call HandleError("AddOtherItem2LV")
Resume out:
End Sub

Public Sub AddWeapon2LV(lv As ListView, tChar As tCharacterProfile, Optional AddToInven As Boolean, Optional nAbility As Integer, _
    Optional ByVal nAttackTypeMUD As eAttackTypeMUD, Optional ByRef sCasts As String = "", Optional ByVal bForceCalc As Boolean)
On Error GoTo error:
Dim oLI As ListItem, x As Integer, sName As String, nSpeed As Integer, nAbilityVal As Integer, sTemp1 As String, sTemp2 As String
Dim tWeaponDmg As tAttackDamage, nSpeedAdj As Integer, bUseCharacter As Boolean, bCalcCombat As Boolean, nNumber As Long

If frmMain.chkGlobalFilter.Value = 1 Then bUseCharacter = True
If frmMain.chkWeaponOptions(3).Value = 1 Then bCalcCombat = True

nSpeedAdj = 100
If bCalcCombat Then
    If nAttackTypeMUD = 0 Then nAttackTypeMUD = frmMain.cmbWeaponCombos(1).ItemData(frmMain.cmbWeaponCombos(1).ListIndex)
    If frmMain.chkWeaponOptions(4).Value = 1 Then nSpeedAdj = 85
ElseIf nAttackTypeMUD = 0 Then
    nAttackTypeMUD = a5_Normal
End If

If nAttackTypeMUD = -4 Then 'surprise punch
    nAttackTypeMUD = a4_Surprise
    sName = "Punch (Surprise)"
    nNumber = 0
    bCalcCombat = True
Else
    sName = tabItems.Fields("Name")
    If sName = "" Then GoTo skip:
    nNumber = tabItems.Fields("Number")
End If

Set oLI = lv.ListItems.Add()
oLI.Text = nNumber
oLI.Tag = nAttackTypeMUD

'If bUseCharacter Then Call PopulateCharacterProfile(tChar, bUseCharacter, True, nAttackTypeMUD, nNumber)

tWeaponDmg = CalculateAttack( _
    tChar, _
    nAttackTypeMUD, _
    nNumber, _
    False, _
    nSpeedAdj, _
    IIf(bCalcCombat, val(frmMain.txtWeaponExtras(2).Text), 0), _
    IIf(bCalcCombat, val(frmMain.txtWeaponExtras(3).Text), 0), _
    IIf(bCalcCombat, val(frmMain.txtWeaponExtras(4).Text), 0), _
    sCasts, _
    bForceCalc)

oLI.ListSubItems.Add (1), "Name", sName
oLI.ListSubItems.Add (2), "Wepn Type", IIf(nNumber > 0, GetWeaponType(tabItems.Fields("WeaponType")), "Fists")
oLI.ListSubItems.Add (3), "Min Dmg", IIf(bCalcCombat, tWeaponDmg.nMinDmg, tabItems.Fields("Min"))
oLI.ListSubItems.Add (4), "Max Dmg", IIf(bCalcCombat, tWeaponDmg.nMaxDmg, tabItems.Fields("Max"))
If tWeaponDmg.nAttackSpeed > 0 Then
    oLI.ListSubItems.Add (5), "Speed", tWeaponDmg.nAttackSpeed
Else
    oLI.ListSubItems.Add (5), "Speed", IIf(nNumber > 0, tabItems.Fields("Speed"), 1150)
End If
oLI.ListSubItems.Add (6), "Level", 0
oLI.ListSubItems.Add (7), "Str", IIf(nNumber > 0, tabItems.Fields("StrReq"), 0)
oLI.ListSubItems.Add (8), "Enc", IIf(nNumber > 0, tabItems.Fields("Encum"), 0)
oLI.ListSubItems.Add (9), "AC", IIf(nNumber > 0, RoundUp(tabItems.Fields("ArmourClass") / 10) & "/" & (tabItems.Fields("DamageResist") / 10), "0/0")
oLI.ListSubItems.Add (10), "Acc", 0 'tabItems.Fields("Accy")
oLI.ListSubItems.Add (11), "BS Acc", IIf(nNumber > 0, "No", 0)
oLI.ListSubItems.Add (12), "Crits", 0
oLI.ListSubItems.Add (13), "Limit", IIf(nNumber > 0, tabItems.Fields("Limit"), 0)

If nNumber = 0 Then GoTo no_number1:

For x = 0 To 19
    Select Case tabItems.Fields("Abil-" & x)
        Case 0:
        Case 58: 'crits
            oLI.ListSubItems(12).Text = tabItems.Fields("AbilVal-" & x)
            
        Case 22, 105, 106: 'acc
            oLI.ListSubItems(10).Text = tabItems.Fields("AbilVal-" & x)
        
        Case 135: 'min level
            oLI.ListSubItems(6).Text = tabItems.Fields("AbilVal-" & x)
            
        Case 116: 'bs accu
            oLI.ListSubItems(11).Text = tabItems.Fields("AbilVal-" & x)
      
    End Select
    
    If nAbility > 0 And tabItems.Fields("Abil-" & x) = nAbility Then
        nAbilityVal = tabItems.Fields("AbilVal-" & x)
    End If
Next x
oLI.ListSubItems(10).Text = val(oLI.ListSubItems(10).Text) + tabItems.Fields("Accy")

nSpeed = tabItems.Fields("Speed")
no_number1:
If nAttackTypeMUD <> a4_Surprise And nSpeed > 0 And tWeaponDmg.nRoundTotal > 0 And tWeaponDmg.nSwings > 0 Then
    oLI.ListSubItems.Add (14), "Dmg/Spd", Round(tWeaponDmg.nRoundTotal / tWeaponDmg.nSwings / nSpeed, 4) * 1000
Else
    oLI.ListSubItems.Add (14), "Dmg/Spd", 0
End If

oLI.ListSubItems.Add (15), "#Swings", Round(tWeaponDmg.nSwings, 2)

If nAttackTypeMUD = 4 Then 'backstab
    oLI.ListSubItems.Add (16), "xSwings", tWeaponDmg.nAvgHit
Else
    oLI.ListSubItems.Add (16), "xSwings", tWeaponDmg.nRoundPhysical
End If
oLI.ListSubItems.Add (17), "Extra", Round(tWeaponDmg.nAvgExtraSwing * tWeaponDmg.nSwings)
oLI.ListSubItems.Add (18), "Dmg/Rnd", tWeaponDmg.nRoundTotal
oLI.ListSubItems.Add (19), "Dmg/1st", tWeaponDmg.nFirstRoundDamage
oLI.ListSubItems(19).Tag = tWeaponDmg.nFirstRoundDamage + Round(tWeaponDmg.nRoundTotal / 100, 2)

'NOTE THAT THERE IS SOME MANUAL ADDING TO LV IN FILTER WEAPONS FOR MA ATTACKS

If nAbility > 0 Then
    Select Case nAbility
        Case 43: 'castssp
            sTemp1 = GetSpellName(nAbilityVal, True)
            sTemp2 = PullSpellEQ(bUseCharacter, tChar.nLevel, nAbilityVal, , , , , , True, , , tChar.nSpellDmgBonus)
            oLI.ListSubItems.Add (20), "Ability", sTemp1 & ": " & sTemp2
        Case Else:
            oLI.ListSubItems.Add (20), "Ability", nAbilityVal
    End Select
ElseIf nAbility = -1 Then
    Select Case nAttackTypeMUD
        Case 1: oLI.ListSubItems.Add (20), "Ability", "Punch"
        Case 2: oLI.ListSubItems.Add (20), "Ability", "Kick"
        Case 3: oLI.ListSubItems.Add (20), "Ability", "Jumpkick"
        Case 4: oLI.ListSubItems.Add (20), "Ability", "Backstab"
        Case 5: oLI.ListSubItems.Add (20), "Ability", "Normal"
        Case 6: oLI.ListSubItems.Add (20), "Ability", "Bash"
        Case 7: oLI.ListSubItems.Add (20), "Ability", "Smash"
        Case Else: oLI.ListSubItems.Add (20), "Ability", ""
    End Select
ElseIf nAttackTypeMUD = a4_Surprise And nNumber = 0 Then
    oLI.ListSubItems.Add (20), "Ability", tWeaponDmg.nHitChance & "% hit, Avg Hit: " _
        & ((tWeaponDmg.nMinDmg + tWeaponDmg.nMaxDmg + (tWeaponDmg.nAvgExtraSwing * 2)) \ 2)
Else
    oLI.ListSubItems.Add (20), "Ability", ""
End If

If AddToInven Then Call frmMain.InvenAddEquip(nNumber, sName, tabItems.Fields("ItemType"), tabItems.Fields("Worn"))

skip:
Set oLI = Nothing

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("AddWeapon2LV")
Resume out:
End Sub

Public Function GetDamageOutput(Optional ByVal nSingleMonster As Long, _
    Optional ByVal nVSAC As Long, Optional ByVal nVSDR As Long, Optional ByVal nVSMR As Long, _
    Optional ByVal nVSDodge As Long = -1, Optional ByVal ePassedDefenseFlags As eDefenseFlags, _
    Optional ByVal nSpeedAdj As Integer = 100, Optional ByVal nSpellImmuLVL As Integer, _
    Optional ByVal nVSMagicLVL As Integer, Optional ByVal eAttackFlags As eAttackRestrictions, _
    Optional ByVal nVSBSDefense As Integer, Optional ByVal nVSrcol As Integer, _
    Optional ByVal nVSrfir As Integer, Optional ByVal nVSrsto As Integer, _
    Optional ByVal nVSrlit As Integer, Optional ByVal nVSrwat As Integer, _
    Optional ByVal bForceCharacter As Boolean) As Currency()
On Error GoTo error:
Dim x As Integer, tAttack As tAttackDamage, tSpellcast As tSpellCastValues, nParty As Integer
Dim nReturnDamage As Currency, nReturnMinDamage As Currency, nReturn(3) As Currency
Dim nDMG_Physical As Double, nDMG_Spell As Double, nAccy As Long, nSwings As Double, nTemp As Long
Dim tCharacter As tCharacterProfile, nAttackTypeMUD As eAttackTypeMUD, nReturnSurpriseDamage As Long
Dim tBackStab As tAttackDamage, nTemp2 As Long, nWeaponMagic As Long, nBackstabWeaponMagic As Long
Dim DF_Flags As eDefenseFlags, bValidTarget As Boolean, nReturnSwings As Double

'results of this look as a value -9990 or greater as a return having value.
'e.g. < -9990 == no damage done and certain values meaning different things.
'-9998 = immune

nReturnDamage = -9999
nReturnMinDamage = -9999
nReturnSurpriseDamage = -9999
nReturnSwings = 0
nReturn(0) = nReturnDamage 'nReturnSurpriseDamage not included in nReturnDamage
nReturn(1) = nReturnMinDamage
nReturn(2) = nReturnSurpriseDamage
nReturn(3) = nReturnSwings

DF_Flags = ePassedDefenseFlags

nAccy = -1
nSpeedAdj = 100
nParty = 1
If frmMain.optMonsterFilter(1).Value = True Then nParty = val(frmMain.txtMonsterLairFilter(0).Text)
If nParty < 1 Then nParty = 1
If nParty > 6 Then nParty = 6

If nParty > 1 Then
    nDMG_Physical = val(frmMain.txtMonsterDamageOUT(0).Text)
    nDMG_Spell = val(frmMain.txtMonsterDamageOUT(1).Text)
    nAccy = val(frmMain.txtMonsterLairFilter(8).Text)
    nSwings = val(frmMain.txtMonsterLairFilter(9).Text)
    If nSwings < 1 Then nSwings = 1
    If nSwings > 6 Then nSwings = 6
    
ElseIf nGlobalAttackTypeMME = a0_oneshot Then 'oneshot
    nReturnDamage = 9999999
    nReturnMinDamage = nReturnDamage
    nReturnSwings = 1
    GoTo done:
    
ElseIf nGlobalAttackTypeMME = a5_Manual Then 'manual
    nDMG_Physical = nGlobalAttackManualP
    nDMG_Spell = nGlobalAttackManualM
    If frmMain.chkGlobalFilter.Value = 1 Then
        nAccy = val(frmMain.lblInvenCharStat(10).Tag)
    Else
        nAccy = 9999
    End If
End If

If nSingleMonster < 1 Then GoTo getdamage:

If nParty = 1 Then 'not party
    If sCharDamageVsMonsterConfig = sGlobalAttackConfig Then
        If nCharDamageVsMonster(nSingleMonster) >= 0 And nCharMinDamageVsMonster(nSingleMonster) >= 0 Then
            nReturnDamage = nCharDamageVsMonster(nSingleMonster)
            nReturnMinDamage = nCharMinDamageVsMonster(nSingleMonster)
            nReturnSurpriseDamage = nCharSurpriseDamageVsMonster(nSingleMonster)
            GoTo done:
        End If
    Else
        ClearSavedDamageVsMonster 'this also sets sCharDamageVsMonsterConfig = sGlobalAttackConfig
    End If
End If

On Error GoTo seek2:
If tabMonsters.Fields("Number") = nSingleMonster Then GoTo monready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabMonsters.Index = "pkMonsters"
tabMonsters.Seek "=", nSingleMonster
If tabMonsters.NoMatch = True Then
    tabMonsters.MoveFirst
    GoTo out:
End If

monready:
nVSAC = tabMonsters.Fields("ArmourClass")
nVSDR = tabMonsters.Fields("DamageResist")
nVSMR = tabMonsters.Fields("MagicRes")
If nNMRVer >= 1.83 Then nVSBSDefense = tabMonsters.Fields("BSDefense")
nTemp = 1 'TEMP FLAG FOR LIVING (set to 0 if nonliving/109 encountered)
For x = 0 To 9
    Select Case tabMonsters.Fields("Abil-" & x)
        Case 0: 'nada
        Case 3: nVSrcol = tabMonsters.Fields("AbilVal-" & x)
        Case 5: nVSrfir = tabMonsters.Fields("AbilVal-" & x)
        Case 65: nVSrsto = tabMonsters.Fields("AbilVal-" & x)
        Case 66: nVSrlit = tabMonsters.Fields("AbilVal-" & x)
        Case 147: nVSrwat = tabMonsters.Fields("AbilVal-" & x)
        Case 28: nVSMagicLVL = tabMonsters.Fields("AbilVal-" & x)
        Case 34: nVSDodge = tabMonsters.Fields("AbilVal-" & x)
        Case 51: DF_Flags = DF_Flags Or DFIAM_IsAntiMag
        Case 78: DF_Flags = DF_Flags Or DF078_IsAnimal
        Case 109: nTemp = 0 'nonliving
        Case 139: nSpellImmuLVL = tabMonsters.Fields("AbilVal-" & x)
    End Select
Next x
If nTemp = 1 Then DF_Flags = DF_Flags Or DF109_IsLiving
If tabMonsters.Fields("Undead") = 1 Then DF_Flags = DF_Flags Or DF023_IsUndead

getdamage:
If nVSDodge < 0 Then nVSDodge = 0

'SUPRISE DAMAGE
If nParty = 1 And nGlobalAttackTypeMME > a0_oneshot And bGlobalAttackBackstab = True Then
    nTemp = 0
    If nGlobalAttackBackstabWeapon > 0 Then
        nTemp = nGlobalAttackBackstabWeapon
    ElseIf nGlobalAttackBackstabWeapon = 0 And nGlobalCharWeaponNumber(0) > 0 Then
        nTemp = nGlobalCharWeaponNumber(0)
    End If
    
    Call PopulateCharacterProfile(tCharacter, bForceCharacter, False, a4_Surprise, nTemp)
    
    If nVSMagicLVL > 0 Then
        If nTemp > 0 Then
            nBackstabWeaponMagic = ItemHasAbility(nTemp, 28) 'magical
            nTemp2 = ItemHasAbility(nTemp, 142) 'hitmagic
            If nTemp2 > nBackstabWeaponMagic Then nBackstabWeaponMagic = nTemp2
        ElseIf tCharacter.nClass > 0 Or tCharacter.nRace > 0 Then 'surprise punch
            If tCharacter.nClass > 0 Then
                nTemp2 = ClassHasAbility(tCharacter.nClass, 142)
                If nTemp2 < 1 Then nTemp2 = 0
                If nTemp2 > nBackstabWeaponMagic Then nBackstabWeaponMagic = nTemp2
            End If
            If tCharacter.nRace > 0 Then
                nTemp2 = RaceHasAbility(tCharacter.nRace, 142)
                If nTemp2 < 1 Then nTemp2 = 0
                If nTemp2 > nBackstabWeaponMagic Then nBackstabWeaponMagic = nTemp2
            End If
        End If
    End If
    If nBackstabWeaponMagic < 0 Then nBackstabWeaponMagic = 0
    
    If nVSMagicLVL <= nBackstabWeaponMagic Then
        tBackStab = CalculateAttack(tCharacter, a4_Surprise, nTemp, False, nSpeedAdj, nVSAC, nVSDR, nVSDodge, , , , , nVSBSDefense)
        nReturnSurpriseDamage = tBackStab.nRoundTotal
        If nReturnSwings < 1 Then nReturnSwings = 1
    Else
        nReturnSurpriseDamage = -9998
        nReturnSwings = 0
    End If
End If

If nParty > 1 Or nGlobalAttackTypeMME = a5_Manual Then 'party or manual
    nReturnDamage = 0
    If nDMG_Physical > 0 Then
        Call PopulateCharacterProfile(tCharacter, bForceCharacter, False, a5_Normal)
        If nParty = 1 Then
            tAttack = CalculateAttack(tCharacter, a5_Normal, 0, False, nSpeedAdj, _
                nVSAC, nVSDR, nVSDodge, , True, nDMG_Physical, nAccy)
        Else
            tAttack = CalculateAttack(tCharacter, a5_Normal, 0, False, nSpeedAdj, _
                nVSAC, 0, nVSDodge, , True, (nDMG_Physical - (nVSDR * nSwings)), nAccy)
        End If
        nReturnDamage = nReturnDamage + tAttack.nRoundTotal
    End If
    If nDMG_Spell > 0 Then
        nReturnDamage = nReturnDamage + CalculateResistDamage(nDMG_Spell, nVSMR, , True, False, (DF_Flags And DFIAM_IsAntiMag) <> 0)
    End If
    nReturnMinDamage = nReturnDamage
    If nReturnSwings < 1 And (nReturnMinDamage + nReturnDamage) > 0 Then nReturnSwings = 1
    GoTo done:
End If

Select Case nGlobalAttackTypeMME
    Case 1, 6, 7: 'eq'd weapon, bash, smash
        If nGlobalAttackTypeMME = a6_PhysBash Or nGlobalAttackTypeMME = a7_PhysSmash Then
            nAttackTypeMUD = nGlobalAttackTypeMME
        Else
            nAttackTypeMUD = a5_Normal
        End If
        
        If nVSMagicLVL > 0 And nGlobalCharWeaponNumber(0) > 0 Then
            nWeaponMagic = ItemHasAbility(nGlobalCharWeaponNumber(0), 28) 'magical
            nTemp2 = ItemHasAbility(nGlobalCharWeaponNumber(0), 142) 'hitmagic
            If nTemp2 > nWeaponMagic Then nWeaponMagic = nTemp2
        End If
        If nWeaponMagic < 0 Then nWeaponMagic = 0
        
        If nVSMagicLVL <= nWeaponMagic Then
            If nGlobalCharWeaponNumber(0) > 0 Then
                Call PopulateCharacterProfile(tCharacter, bForceCharacter, False, nAttackTypeMUD)
                tAttack = CalculateAttack(tCharacter, nAttackTypeMUD, nGlobalCharWeaponNumber(0), False, nSpeedAdj, nVSAC, nVSDR, nVSDodge)
                nReturnDamage = tAttack.nRoundTotal
                nReturnSwings = Round(tAttack.nSwings, 2)
            End If
        Else
            nReturnDamage = -9998
            nReturnSwings = 0
        End If

    Case 2, 3:
        '2-spell learned: GetSpellShort(nGlobalAttackSpellNum) & " @ " & Val(txtGlobalLevel(0).Text)
        '3-spell any: GetSpellShort(nGlobalAttackSpellNum) & " @ " & nGlobalAttackSpellLVL
        If nGlobalAttackSpellNum > 0 Then
            Call PopulateCharacterProfile(tCharacter, bForceCharacter, True)

            tSpellcast = CalculateSpellCast(tCharacter, nGlobalAttackSpellNum, _
                            IIf(nGlobalAttackTypeMME = a3_SpellAny, nGlobalAttackSpellLVL, tCharacter.nLevel), nVSMR, (DF_Flags And DFIAM_IsAntiMag) <> 0, _
                            nVSrcol, nVSrfir, nVSrsto, nVSrlit, nVSrwat)
            
            If nSpellImmuLVL = 0 Or tSpellcast.nCastLevel > nSpellImmuLVL Then
                If eAttackFlags = AR000_Unknown Then
                    If SpellSeek(nGlobalAttackSpellNum) Then
                        For x = 0 To 9
                            Select Case tabSpells.Fields("Abil-" & x) 'tabSpells.Fields("AbilVal-" & x)
                                Case 0: 'nada
                                Case 23: eAttackFlags = (eAttackFlags Or AR023_Undead)
                                Case 80: eAttackFlags = (eAttackFlags Or AR080_Animal)
                                Case 108: eAttackFlags = (eAttackFlags Or AR108_Living)
                            End Select
                        Next x
                        If eAttackFlags <= AR001_None Then bValidTarget = True
                    End If
                ElseIf eAttackFlags = AR001_None Then
                    bValidTarget = True
                End If
                
                If bValidTarget = False Then
                    If eAttackFlags > AR001_None Then
                        If (eAttackFlags And AR023_Undead) <> 0 Then
                            If (DF_Flags And DF023_IsUndead) <> 0 Then bValidTarget = True
                        ElseIf (eAttackFlags And AR080_Animal) <> 0 Then
                            If (DF_Flags And DF078_IsAnimal) <> 0 Then bValidTarget = True
                        ElseIf (eAttackFlags And AR108_Living) <> 0 Then
                            If (DF_Flags And DF109_IsLiving) <> 0 Then bValidTarget = True
                        End If
                    Else
                        bValidTarget = True
                    End If
                End If
            End If
                
            If bValidTarget Then
                nReturnDamage = tSpellcast.nAvgRoundDmg
                nReturnSwings = tSpellcast.nNumCasts
            Else
                nReturnDamage = -9998
                nReturnSwings = 0
            End If
        End If

    Case 4: 'martial arts attack
        '1-Punch, 2-Kick, 3-JumpKick
        Call PopulateCharacterProfile(tCharacter, bForceCharacter, False, IIf(nGlobalAttackMA > 1, nGlobalAttackMA, 1))
        Select Case nGlobalAttackMA
            Case 2: 'kick
                tAttack = CalculateAttack(tCharacter, a2_Kick, , False, nSpeedAdj, nVSAC, nVSDR, nVSDodge)
                nReturnDamage = tAttack.nRoundTotal
                nReturnSwings = tAttack.nSwings
            Case 3: 'jumpkick
                tAttack = CalculateAttack(tCharacter, a3_Jumpkick, , False, nSpeedAdj, nVSAC, nVSDR, nVSDodge)
                nReturnDamage = tAttack.nRoundTotal
                nReturnSwings = tAttack.nSwings
            Case Else: 'punch
                tAttack = CalculateAttack(tCharacter, a1_Punch, , False, nSpeedAdj, nVSAC, nVSDR, nVSDodge)
                nReturnDamage = tAttack.nRoundTotal
                nReturnSwings = tAttack.nSwings
        End Select

End Select

If nReturnDamage > -9990 Then
    If tAttack.nSwings > 0 Then
        nReturnMinDamage = tAttack.nMinDmg
        If tAttack.nAvgExtraHit > 0 And tAttack.nAvgExtraHit = tAttack.nAvgExtraSwing Then nReturnMinDamage = nReturnMinDamage + tAttack.nAvgExtraHit
        nReturnMinDamage = tAttack.nMinDmg * tAttack.nSwings
    ElseIf tSpellcast.nMinCast > 0 Then
        nReturnMinDamage = tSpellcast.nMinCast * tSpellcast.nNumCasts
    End If
End If

If nSingleMonster > 0 And nParty = 1 Then
    nCharDamageVsMonster(nSingleMonster) = nReturnDamage
    nCharMinDamageVsMonster(nSingleMonster) = nReturnMinDamage
    nCharSurpriseDamageVsMonster(nSingleMonster) = nReturnSurpriseDamage
End If

done:
nReturn(0) = nReturnDamage 'nReturnSurpriseDamage not included in nReturnDamage
nReturn(1) = nReturnMinDamage
nReturn(2) = nReturnSurpriseDamage
nReturn(3) = nReturnSwings
GetDamageOutput = nReturn

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetDamageOutput")
Resume out:
End Function

Public Sub PopulateCharacterProfile(ByRef tChar As tCharacterProfile, Optional ByVal bForceUseChar As Boolean, Optional ByVal bForceNoParty As Boolean, _
    Optional ByVal nAttackTypeMUD As eAttackTypeMUD, Optional ByVal nWeaponNumber As Long)
On Error GoTo error:
Dim bUseCharacter As Boolean, bCalcAccy As Boolean, nWeapon As Long
Dim nNormAccyAdj As Integer, nBSAccyAdj As Integer

If frmMain.chkGlobalFilter.Value = 1 Or bForceUseChar Then bUseCharacter = True

If frmMain.optMonsterFilter(1).Value = True And val(frmMain.txtMonsterLairFilter(0).Text) > 1 Then
    tChar.nParty = val(frmMain.txtMonsterLairFilter(0).Text)
End If
If tChar.nParty < 1 Then tChar.nParty = 1
If tChar.nParty > 6 Then tChar.nParty = 6

If (bUseCharacter And tChar.nParty < 2) Or bForceUseChar Then
    tChar.bIsLoadedCharacter = True
    tChar.nLevel = val(frmMain.txtGlobalLevel(0).Text)
    tChar.nClass = frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex)
    tChar.nRace = frmMain.cmbGlobalRace(0).ItemData(frmMain.cmbGlobalRace(0).ListIndex)
    tChar.nCombat = GetClassCombat(frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex))
    tChar.nEncumPCT = CalcEncumbrancePercent(val(frmMain.lblInvenCharStat(0).Caption), val(frmMain.lblInvenCharStat(1).Caption))
    tChar.nDodge = val(frmMain.lblInvenCharStat(8).Tag)
    tChar.nDodgeCap = GetDodgeCap(tChar.nClass)
    tChar.nSTR = val(frmMain.txtCharStats(0).Tag)
    tChar.nAGI = val(frmMain.txtCharStats(3).Tag)
    tChar.nINT = val(frmMain.txtCharStats(1).Tag)
    tChar.nCHA = val(frmMain.txtCharStats(5).Tag)
    tChar.nCrit = val(frmMain.lblInvenCharStat(7).Tag)
    tChar.nPlusMaxDamage = val(frmMain.lblInvenCharStat(11).Tag)
    tChar.nPlusMinDamage = val(frmMain.lblInvenCharStat(30).Tag)
    tChar.nPlusBSaccy = val(frmMain.lblInvenCharStat(13).Tag)
    tChar.nPlusBSmindmg = val(frmMain.lblInvenCharStat(14).Tag)
    tChar.nPlusBSmaxdmg = val(frmMain.lblInvenCharStat(15).Tag)
    tChar.nStealth = val(frmMain.lblInvenCharStat(19).Tag)
    tChar.nHP = val(frmMain.lblCharMaxHP.Tag)
    tChar.nHPRegen = val(frmMain.lblCharRestRate.Tag)
    If bGlobalAttackUseMeditate Then tChar.nMeditateRate = val(frmMain.txtCharManaRegen.Tag)
    tChar.nMaxMana = val(frmMain.lblCharMaxMana.Tag)
    tChar.nManaRegen = val(frmMain.lblCharManaRate.Tag)
    tChar.nDamageThreshold = nGlobalAttackHealValue
    tChar.nSpellcasting = val(frmMain.lblCharSC.Tag)
    tChar.nSpellDmgBonus = val(frmMain.lblInvenCharStat(33).Tag)
'    If bGreaterMUD And tChar.nSpellcasting > 150 Then
'        tChar.nSpellDmgBonus = tChar.nSpellDmgBonus + GMUD_GetSpDmgMultiplierFromSC(tChar.nSpellcasting)
'    End If
    tChar.nSpellOverhead = nGlobalAttackHealCost + (val(frmMain.lblCharBless.Caption) / 6)
    
    If (nGlobalAttackTypeMME = a2_Spell Or nGlobalAttackTypeMME = a3_SpellAny) And nGlobalAttackSpellNum > 0 Then   'spell attack
        tChar.nSpellAttackCost = GetSpellManaCost(nGlobalAttackSpellNum)
    End If
    
    If nAttackTypeMUD = a4_Surprise Then
        bCalcAccy = True
    
    ElseIf bGreaterMUD And nAttackTypeMUD > a0_none Then
        Select Case nAttackTypeMUD
            Case 1, 2, 3, 4, 5:  'pu, ki, jk, bs, a
                If nGlobalAttackTypeMME = a6_PhysBash Or nGlobalAttackTypeMME = a7_PhysSmash Then bCalcAccy = True
            Case 6, 7: 'bash, smash
                If nGlobalAttackTypeMME <> nAttackTypeMUD Then bCalcAccy = True
        End Select
    End If
    
    If bCalcAccy Then
        If nAttackTypeMUD = a4_Surprise Then
            nWeapon = 0
            If nWeaponNumber > 0 Then
                nWeapon = nWeaponNumber
            ElseIf nWeaponNumber < 0 Then 'punch
                nWeapon = 0
            ElseIf bGlobalAttackBackstab And nGlobalAttackBackstabWeapon > 0 Then
                nWeapon = nGlobalAttackBackstabWeapon
            ElseIf Not bGlobalAttackBackstab Or nGlobalAttackBackstabWeapon = 0 Then
                nWeapon = nGlobalCharWeaponNumber(0)
            End If
            
            If nWeapon <> nGlobalCharWeaponNumber(0) Then
                If nWeapon > 0 Then
                    nNormAccyAdj = ItemHasAbility(nWeapon, 22)
                    If nNormAccyAdj < 0 Then nNormAccyAdj = 0
                    nBSAccyAdj = ItemHasAbility(nWeapon, 116)
                    If nBSAccyAdj < 0 Then nBSAccyAdj = 0
                    
                    If nGlobalCharWeaponNumber(1) > 0 Then
                        If IsTwoHandedWeapon(nWeapon) Then
                            tChar.nPlusBSaccy = tChar.nPlusBSaccy - nGlobalCharWeaponBSaccy(1)
                            nNormAccyAdj = nNormAccyAdj - nGlobalCharWeaponAccy(1)
                        End If
                    End If
                End If
                tChar.nPlusBSaccy = tChar.nPlusBSaccy + nBSAccyAdj - nGlobalCharWeaponBSaccy(0)
                nNormAccyAdj = nNormAccyAdj - nGlobalCharWeaponAccy(0)
            End If
            
            tChar.nAccuracy = CalculateBackstabAccuracy(tChar.nStealth, tChar.nAGI, tChar.nPlusBSaccy, _
                GetClassStealth(tChar.nClass), nGlobalCharAccyAbils + nGlobalCharAccyOther + nNormAccyAdj, _
                tChar.nLevel, tChar.nSTR, GetItemStrReq(nWeapon))
        Else
            tChar.nAccuracy = CalculateAccuracy(tChar.nClass, tChar.nLevel, tChar.nSTR, tChar.nAGI, tChar.nINT, tChar.nCHA, _
                nGlobalCharAccyItems, nGlobalCharAccyOther + nGlobalCharAccyAbils, tChar.nEncumPCT, , nAttackTypeMUD)
        End If
    Else
        tChar.nAccuracy = val(frmMain.lblInvenCharStat(10).Tag)
    End If
        
    'punch
    tChar.nMAPlusSkill(1) = val(frmMain.lblInvenCharStat(37).Tag)
    tChar.nMAPlusAccy(1) = val(frmMain.lblInvenCharStat(40).Tag)
    tChar.nMAPlusDmg(1) = val(frmMain.lblInvenCharStat(34).Tag)
    'Kick
    tChar.nMAPlusSkill(2) = val(frmMain.lblInvenCharStat(38).Tag)
    tChar.nMAPlusAccy(2) = val(frmMain.lblInvenCharStat(41).Tag)
    tChar.nMAPlusDmg(2) = val(frmMain.lblInvenCharStat(35).Tag)
    'Jumpkick
    tChar.nMAPlusSkill(3) = val(frmMain.lblInvenCharStat(39).Tag)
    tChar.nMAPlusAccy(3) = val(frmMain.lblInvenCharStat(42).Tag)
    tChar.nMAPlusDmg(3) = val(frmMain.lblInvenCharStat(36).Tag)
    
ElseIf tChar.nParty > 1 And Not bForceNoParty Then 'vs party
    'txtMonsterLairFilter... 0-#, 1-ac, 2-dr, 3-mr, 4-dodge, 5-HP, 6-#antimag, 7-hpregen, 8-accy
    tChar.nHP = val(frmMain.txtMonsterLairFilter(5).Text)
    If tChar.nHP < 1 Then
        frmMain.txtMonsterLairFilter(5).Text = 1
        tChar.nHP = 1
    End If
    tChar.nHP = tChar.nHP * tChar.nParty
    tChar.nHPRegen = val(frmMain.txtMonsterLairFilter(7).Text) * tChar.nParty
    tChar.nDamageThreshold = val(frmMain.txtMonsterDamage.Text)
    tChar.nAccuracy = val(frmMain.txtMonsterLairFilter(8).Text)
Else 'no party / not char
    tChar.nLevel = 255
    tChar.nCombat = 5
    tChar.nSTR = 255
    tChar.nAGI = 255
    tChar.nStealth = 255
    tChar.nAccuracy = 999
    tChar.nPlusBSaccy = 999
    tChar.bClassStealth = True
    tChar.bRaceStealth = True
    tChar.nHP = 1000
    tChar.nHPRegen = tChar.nHP * 0.05
    If nAttackTypeMUD >= a1_Punch And nAttackTypeMUD <= a3_Jumpkick Then
        tChar.nMAPlusSkill(1) = 1
        tChar.nMAPlusSkill(2) = 1
        tChar.nMAPlusSkill(3) = 1
    End If
    If nNMRVer < 1.83 Then
        tChar.nDamageThreshold = val(frmMain.txtMonsterDamage.Text)
    Else
        tChar.nDamageThreshold = nGlobalAttackHealValue
        If (nGlobalAttackTypeMME = a2_Spell Or nGlobalAttackTypeMME = a3_SpellAny) And nGlobalAttackSpellNum > 0 Then   'spell attack
            tChar.nSpellAttackCost = GetSpellManaCost(nGlobalAttackSpellNum)
        End If
    End If
End If

If tChar.nDamageThreshold < 0 Then tChar.nDamageThreshold = 0
If tChar.nDamageThreshold > 9999999 Then tChar.nDamageThreshold = 9999999
If tChar.nHP < 1 Then tChar.nHP = 1
If tChar.nHPRegen < 1 Then tChar.nHPRegen = 1

If tChar.nMaxMana < 0 Then tChar.nMaxMana = 0
If tChar.nManaRegen < 0 Then tChar.nManaRegen = 0
If tChar.nSpellOverhead < 0 Then tChar.nSpellOverhead = 0
If tChar.nSpellAttackCost < 0 Then tChar.nSpellAttackCost = 0
If tChar.nEncumPCT < 0 Then tChar.nEncumPCT = 0
If tChar.nAccuracy < 0 Then tChar.nAccuracy = 0

If tChar.nMaxMana > 9999999 Then tChar.nMaxMana = 9999999
If tChar.nManaRegen > 9999999 Then tChar.nManaRegen = 9999999
If tChar.nSpellOverhead > 9999999 Then tChar.nSpellOverhead = 9999999
If tChar.nSpellAttackCost > 9999999 Then tChar.nSpellAttackCost = 9999999
If tChar.nAccuracy > 9999999 Then tChar.nAccuracy = 9999999
If tChar.nEncumPCT > 100 Then tChar.nEncumPCT = 100

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("PopulateCharacterProfile")
Resume out:
End Sub


Public Sub AddSpell2LV(lv As ListView, tChar As tCharacterProfile, Optional ByVal AddBless As Boolean)
On Error GoTo error:
Dim oLI As ListItem, sName As String, x As Integer, nSpell As Long, sTimesCast As String
Dim nSpellDamage As Currency, nSpellDuration As Long, bUseCharacter As Boolean, nManaCost As Long
Dim bCalcCombat As Boolean, nCastPCT As Double, tSpellcast As tSpellCastValues

If frmMain.chkSpellOptions(0).Value = 1 And val(frmMain.txtSpellOptions(0).Text) > 0 Then bCalcCombat = True
If frmMain.chkGlobalFilter.Value = 1 And val(frmMain.txtGlobalLevel(1).Text) > 0 Then bUseCharacter = True

nSpell = tabSpells.Fields("Number")
sName = tabSpells.Fields("Name")
If sName = "" Then GoTo skip:
If Left(sName, 1) = "1" Then GoTo skip:
If Left(LCase(sName), 3) = "sdf" Then GoTo skip:

Set oLI = lv.ListItems.Add()
oLI.Text = nSpell

oLI.ListSubItems.Add (1), "Name", sName
oLI.ListSubItems.Add (2), "Short", tabSpells.Fields("Short")
oLI.ListSubItems.Add (3), "Magery", GetMagery(tabSpells.Fields("Magery"), tabSpells.Fields("MageryLVL"))
oLI.ListSubItems.Add (4), "Level", tabSpells.Fields("ReqLevel")

If bUseCharacter Then
    If bCalcCombat Then
        tSpellcast = CalculateSpellCast(tChar, nSpell, tChar.nLevel, val(frmMain.txtSpellOptions(0).Text), IIf(frmMain.chkSpellOptions(2).Value = 1, True, False))
    Else
        tSpellcast = CalculateSpellCast(tChar, nSpell, tChar.nLevel)
    End If
Else
    tSpellcast = CalculateSpellCast(tChar, nSpell, tabSpells.Fields("ReqLevel"))
End If

If tabSpells.Fields("ManaCost") > 0 Then
    If tSpellcast.nNumCasts > 1 Then
        nManaCost = (tabSpells.Fields("ManaCost") * tSpellcast.nNumCasts)
    Else
        nManaCost = tabSpells.Fields("ManaCost")
    End If
End If
oLI.ListSubItems.Add (5), "Mana", nManaCost

If Not tabSpells.Fields("Number") = nSpell Then
    tabSpells.Index = "pkSpells"
    tabSpells.Seek "=", nSpell
    If tabSpells.NoMatch = True Then
        tabSpells.MoveFirst
        Exit Sub
    End If
End If

nSpellDuration = tSpellcast.nDuration
nCastPCT = tSpellcast.nCastChance / 100

If bUseCharacter Then
    oLI.ListSubItems.Add (6), "Diff", tSpellcast.nCastChance & "%"
Else
    oLI.ListSubItems.Add (6), "Diff", tabSpells.Fields("Diff")
End If


If tabSpells.Fields("Learnable") = 1 Or tabSpells.Fields("ManaCost") > 0 Then

    oLI.ListSubItems.Add (7), "Dmg", (tSpellcast.nAvgRoundDmg * tSpellcast.nDuration) 'Round(nSpellDamage)
    
    nSpellDamage = 0
    If tSpellcast.nAvgRoundDmg > 0 Then
        If tabSpells.Fields("ManaCost") > 0 Then
            If tSpellcast.nNumCasts > 1 Then
                nSpellDamage = Round(tSpellcast.nAvgRoundDmg / nManaCost, 1)
            Else
                nSpellDamage = Round((tSpellcast.nAvgRoundDmg * tSpellcast.nDuration) / tabSpells.Fields("ManaCost"), 1)
            End If
        End If
    End If
    
    oLI.ListSubItems.Add (8), "Dmg/M", nSpellDamage
    oLI.ListSubItems.Add (9), "Heal", (tSpellcast.nAvgRoundHeals * tSpellcast.nDuration) 'Round(nSpellDamage)
    
    nSpellDamage = 0
    If tSpellcast.nAvgRoundHeals <> 0 Then
        If tabSpells.Fields("ManaCost") > 0 Then
            If tSpellcast.nNumCasts > 1 Then
                nSpellDamage = Round(tSpellcast.nAvgRoundHeals / nManaCost, 1)
            Else
                nSpellDamage = Round((tSpellcast.nAvgRoundHeals * tSpellcast.nDuration) / tabSpells.Fields("ManaCost"), 1)
            End If
        End If
    End If
    
    oLI.ListSubItems.Add (10), "Heal/M", nSpellDamage
Else
    oLI.ListSubItems.Add (7), "Dmg", 0
    oLI.ListSubItems.Add (8), "Dmg/M", 0
    oLI.ListSubItems.Add (9), "Heal", 0
    oLI.ListSubItems.Add (10), "Heal/M", 0
End If

bQuickSpell = True
If lv.name = "lvSpellBook" And FormIsLoaded("frmSpellBook") And bUseCharacter Then
    If val(frmSpellBook.txtLevel) > 0 Then
        oLI.ListSubItems.Add (11), "Detail", PullSpellEQ(True, val(frmSpellBook.txtLevel), nSpell, Nothing, , , , , True, _
                                                tSpellcast.nMinCast, tSpellcast.nMaxCast) & sTimesCast
    Else
        oLI.ListSubItems.Add (11), "Detail", PullSpellEQ(False, , nSpell, Nothing, , , , , True, _
                                                tSpellcast.nMinCast, tSpellcast.nMaxCast) & sTimesCast
    End If
Else
    If bUseCharacter Then
        oLI.ListSubItems.Add (11), "Detail", PullSpellEQ(True, val(frmMain.txtGlobalLevel(1).Text), nSpell, Nothing, , , , , True, _
                                                tSpellcast.nMinCast, tSpellcast.nMaxCast) & sTimesCast
    Else
        oLI.ListSubItems.Add (11), "Detail", PullSpellEQ(False, , nSpell, Nothing, , , , , True, _
                                                tSpellcast.nMinCast, tSpellcast.nMaxCast) & sTimesCast
    End If
End If
bQuickSpell = False

If Not tabSpells.Fields("Number") = nSpell Then tabSpells.Seek "=", nSpell

If AddBless Then
    If tabSpells.Fields("Learnable") = 1 Or Len(tabSpells.Fields("Learned From")) > 0 _
        Or (tabSpells.Fields("Magery") = 5 And tabSpells.Fields("ReqLevel") > 0) Then
                    
        Select Case tabSpells.Fields("Targets")
            Case 0: GoTo skip: 'GetSpellTargets = "User"
            Case 1: 'GetSpellTargets = "Self"
            Case 2: 'GetSpellTargets = "Self or User"
            Case 3: GoTo skip: 'GetSpellTargets = "Divided Area (not self)"
            Case 4: GoTo skip: 'GetSpellTargets = "Monster"
            Case 5: 'GetSpellTargets = "Divided Area (incl self)"
            Case 6: GoTo skip: 'GetSpellTargets = "Any"
            Case 7: GoTo skip: 'GetSpellTargets = "Item"
            Case 8: GoTo skip: 'GetSpellTargets = "Monster or User"
            Case 9: GoTo skip: 'GetSpellTargets = "Divided Attack Area"
            Case 10: GoTo skip: ' GetSpellTargets = "Divided Party Area"
            Case 11: GoTo skip: ' GetSpellTargets = "Full Area"
            Case 12: GoTo skip: ' GetSpellTargets = "Full Attack Area"
            Case 13: 'GetSpellTargets = "Full Party Area"
            Case Else: GoTo skip: 'GetSpellTargets = "Unknown (" & nNum & ")"
        End Select

        If tabSpells.Fields("Dur") > 0 Then
            'nothing
        ElseIf tabSpells.Fields("DurInc") > 0 And _
            tabSpells.Fields("DurIncLVLs") > 0 Then
            'nothing
        Else
            GoTo skip:
        End If
        
        For x = 0 To 9
            frmMain.cmbCharBless(x).AddItem sName & " (" & nSpell & ")"
            frmMain.cmbCharBless(x).ItemData(frmMain.cmbCharBless(x).NewIndex) = nSpell
        Next x
    End If
End If
    
skip:
Set oLI = Nothing


out:
On Error Resume Next
bQuickSpell = False
Exit Sub
error:
Call HandleError("AddSpell2LV")
Resume out:
End Sub

Public Sub AddRace2LV(lv As ListView)

On Error GoTo error:

Dim oLI As ListItem, x As Integer, sAbil As String
    
    If tabRaces.Fields("Name") = "" Then GoTo skip:
    
    Set oLI = lv.ListItems.Add()
    oLI.Text = tabRaces.Fields("Number")
    
    oLI.ListSubItems.Add (1), "Name", tabRaces.Fields("Name")
    oLI.ListSubItems.Add (2), "Exp%", tabRaces.Fields("ExpTable") & "%"
    oLI.ListSubItems.Add (3), "HP", tabRaces.Fields("HPPerLVL")
    oLI.ListSubItems.Add (4), "Str", tabRaces.Fields("mSTR") & "-" & tabRaces.Fields("xSTR")
    oLI.ListSubItems.Add (5), "Int", tabRaces.Fields("mINT") & "-" & tabRaces.Fields("xINT")
    oLI.ListSubItems.Add (6), "Wis", tabRaces.Fields("mWIL") & "-" & tabRaces.Fields("xWIL")
    oLI.ListSubItems.Add (7), "Agi", tabRaces.Fields("mAGL") & "-" & tabRaces.Fields("xAGL")
    oLI.ListSubItems.Add (8), "Hea", tabRaces.Fields("mHEA") & "-" & tabRaces.Fields("xHEA")
    oLI.ListSubItems.Add (9), "Cha", tabRaces.Fields("mCHM") & "-" & tabRaces.Fields("xCHM")
    
    For x = 0 To 9
        Select Case tabRaces.Fields("Abil-" & x)
            Case 0:
            Case Else:
                If sAbil <> "" Then sAbil = sAbil & ", "
                sAbil = sAbil & GetAbilityStats(tabRaces.Fields("Abil-" & x), tabRaces.Fields("AbilVal-" & x))
                If Right(sAbil, 2) = ", " Then sAbil = Left(sAbil, Len(sAbil) - 2)
        End Select
    Next

    oLI.ListSubItems.Add (10), "Abilities", sAbil
    
skip:
Set oLI = Nothing


out:
On Error Resume Next
Exit Sub
error:
Call HandleError("AddRace2LV")
Resume out:
End Sub

Public Function IsMobKillable(ByVal nCharDMG As Double, ByVal nCharHP As Long, ByVal nMobDmg As Double, ByVal nMobHP As Long, _
    Optional ByVal nCharHPRegen As Integer = 0, Optional ByVal nMobHPRegen As Long = 0) As Boolean
On Error GoTo error:
Dim nFactor As Double, nRoundsToKill As Double, nRoundsToDeath As Double
Dim nMobTotalHP As Long, nCharTotalHP As Long, nEffDmg As Double, nRegenPerRound As Double

If nCharDMG <= 0 And nMobHP > 0 Then Exit Function
If nMobHP < 1 Then
    IsMobKillable = True
    Exit Function
End If

nFactor = 0.25
nCharDMG = nCharDMG * (nFactor + 1)
nRoundsToKill = nMobHP / nCharDMG
If nRoundsToKill < 1 Then nRoundsToKill = 1

If nRoundsToKill > 1 Then
    nRegenPerRound = nMobHPRegen / 18
    If nRoundsToKill < 18 Then
        'reducing by precentage change to reach the hp tick at 90 seconds
        nRegenPerRound = nRegenPerRound * (nRoundsToKill / 18)
    End If
End If

If nRegenPerRound > 0 Then
    nEffDmg = nCharDMG - nRegenPerRound
    If nEffDmg <= 0 Then Exit Function
    nRoundsToKill = nMobHP / nEffDmg
    If nRoundsToKill < 1 Then nRoundsToKill = 1
    nMobTotalHP = nMobHP + (nRegenPerRound * nRoundsToKill)
Else
    nMobTotalHP = nMobHP
End If

If nRoundsToKill > 720 Then Exit Function 'would take over an hour to kill... prehaps nRegenTime should be worked into here to allow >1hr?

nMobDmg = nMobDmg * (1 - nFactor)
If nMobDmg <= 0 Then
    IsMobKillable = True
    Exit Function
End If

If nCharHPRegen > 0 Then
    nCharTotalHP = nCharHP + ((nCharHP / nMobDmg) * (nCharHPRegen / 3 / 6))
Else
    nCharTotalHP = nCharHP
End If
nRoundsToDeath = nCharTotalHP / nMobDmg

If nRoundsToDeath >= nRoundsToKill Then IsMobKillable = True

out:
On Error Resume Next
Exit Function
error:
Call HandleError("IsMobKillable")
Resume out:
End Function

Public Sub AddMonster2LV(lv As ListView, tChar As tCharacterProfile, Optional ByVal nDamageOut As Long = -9999, _
    Optional ByVal nPassEXP As Currency = -1, Optional ByVal nPassRecovery As Double = -1, _
    Optional ByVal nSurpriseDamageOut As Long = -9999)
On Error GoTo error:
Dim oLI As ListItem, sName As String, nExp As Currency, nHP As Currency, x As Integer
Dim nAvgDmg As Long, nExpDmgHP As Currency, nIndex As Integer, nMagicLVL As Integer
Dim nScriptValue As Currency, nLairPCT As Currency, nPossSpawns As Long
Dim nMaxLairsBeforeRegen As Currency, nPossyPCT As Currency, bAsterisks As Boolean, sTemp As String
Dim tAvgLairInfo As LairInfoType, nTimeRecovering As Double, sTemp2 As String
Dim nMonsterNum As Long, nDmgOut() As Currency
Dim tExpInfo As tExpPerHourInfo, nMobDodge As Integer, bUseCharacter As Boolean
Dim bHasAntiMagic As Boolean, nParty As Integer 'tSpellcast As tSpellCastValues, nTemp As Long, nMaxAcc As Long, nAccAvg As Long
Dim sSpellExtraTypes As String, sSpellAttackTypes As String
Dim nExpPerHour As Currency, nPercent As Integer ', nTotalMeleeAttackPercentage As Integer

nMonsterNum = tabMonsters.Fields("Number")
If frmMain.chkGlobalFilter.Value = 1 Then bUseCharacter = True

If frmMain.optMonsterFilter(1).Value = True And val(frmMain.txtMonsterLairFilter(0).Text) > 1 Then
    nParty = val(frmMain.txtMonsterLairFilter(0).Text)
End If
If nParty < 1 Then nParty = 1
If nParty > 6 Then nParty = 6

If nNMRVer >= 1.83 And lv.hWnd = frmMain.lvMonsters.hWnd And frmMain.optMonsterFilter(1).Value = True _
    And (tLastAvgLairInfo.sGroupIndex <> tabMonsters.Fields("Summoned By") Or tLastAvgLairInfo.sGlobalAttackConfig <> sGlobalAttackConfig) Then
    tLastAvgLairInfo = GetLairAveragesFromLocs(tabMonsters.Fields("Summoned By"))
ElseIf (nNMRVer < 1.83 Or lv.hWnd <> frmMain.lvMonsters.hWnd Or frmMain.optMonsterFilter(1).Value = False) And Not tLastAvgLairInfo.sGroupIndex = "" Then
    tLastAvgLairInfo = GetLairInfo("") 'reset
End If

tAvgLairInfo = tLastAvgLairInfo
If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum

sName = tabMonsters.Fields("Name")
If sName = "" Or Left(sName, 3) = "sdf" Then GoTo skip:

For x = 0 To 9 'abilities
    If Not tabMonsters.Fields("Abil-" & x) = 0 Then
        Select Case tabMonsters.Fields("Abil-" & x)
            Case 0:
            Case 28: 'magical
                nMagicLVL = tabMonsters.Fields("AbilVal-" & x)
            Case 34: 'dodge
                nMobDodge = tabMonsters.Fields("AbilVal-" & x)
            Case 51: 'anti-magic
                bHasAntiMagic = True
            'Case 139: 'spellimmu
                'nSpellImmuLVL = tabMonsters.Fields("AbilVal-" & x)
        End Select
    End If
Next

Set oLI = lv.ListItems.Add()
oLI.Text = tabMonsters.Fields("Number")

nIndex = 1 '2
oLI.ListSubItems.Add (nIndex), "Name", sName

nIndex = nIndex + 1 '3
oLI.ListSubItems.Add (nIndex), "RGN", tabMonsters.Fields("RegenTime")
oLI.ListSubItems(nIndex).Tag = tabMonsters.Fields("RegenTime")

nIndex = nIndex + 1 '4
If UseExpMulti Then
    nExp = tabMonsters.Fields("EXP") * tabMonsters.Fields("ExpMulti")
Else
    nExp = tabMonsters.Fields("EXP")
End If
oLI.ListSubItems.Add (nIndex), "Exp", IIf(nExp > 0, Format(nExp, "#,#"), 0)
oLI.ListSubItems(nIndex).Tag = nExp

nIndex = nIndex + 1 '5
sTemp = ""
If tAvgLairInfo.nTotalLairs > 0 And tabMonsters.Fields("RegenTime") = 0 Then
    nHP = tAvgLairInfo.nAvgHP
    sTemp = "*"
Else
    nHP = tabMonsters.Fields("HP")
End If
oLI.ListSubItems.Add (nIndex), "HP", IIf(nHP > 0, Format(nHP, "#,#"), 0) & sTemp
oLI.ListSubItems(nIndex).Tag = nHP

nIndex = nIndex + 1 '6
oLI.ListSubItems.Add (nIndex), "AC/DR", tabMonsters.Fields("ArmourClass") & "/" & tabMonsters.Fields("DamageResist")
oLI.ListSubItems(nIndex).Tag = tabMonsters.Fields("ArmourClass") + tabMonsters.Fields("DamageResist")

nIndex = nIndex + 1 '7
oLI.ListSubItems.Add (nIndex), "Dodge", nMobDodge
oLI.ListSubItems(nIndex).Tag = nMobDodge

nIndex = nIndex + 1 '8
oLI.ListSubItems.Add (nIndex), "MR", tabMonsters.Fields("MagicRes")
oLI.ListSubItems(nIndex).Tag = tabMonsters.Fields("MagicRes")

nPercent = 0
For x = 0 To 4 'between round spells
    If Not tabMonsters.Fields("MidSpell-" & x) = 0 Then
        nPercent = tabMonsters.Fields("MidSpell%-" & x) - nPercent
        If nPercent > 0 Then
            If SpellDoesDamage(tabMonsters.Fields("MidSpell-" & x), True) Then
                sSpellExtraTypes = sSpellExtraTypes & SpellAttackTypeEnum(GetSpellAttackType(tabMonsters.Fields("MidSpell-" & x)), True)
            End If
        End If
        If nPercent < 0 Then nPercent = 0
    End If
Next

Const MAJ_THRESH_PCT As Long = 51
Dim meleeTotalPct As Long, prevCumPct As Long, currCum As Long
Dim nAcc As Long, maxAcc As Long
Dim uniqAcc(0 To 4) As Long, uniqPct(0 To 4) As Long, uniqCount As Long
Dim i As Long, found As Long
Dim domIdx As Long, domPct As Long, domAcc As Long

For x = 0 To 4
    If tabMonsters.Fields("AttType-" & x) > 0 And tabMonsters.Fields("AttType-" & x) <= 3 And tabMonsters.Fields("Att%-" & x) > 0 Then
        If nNMRVer >= 1.8 Then
            nPercent = Round(tabMonsters.Fields("AttTrue%-" & x))
        Else
            currCum = CLng(tabMonsters.Fields("Att%-" & x))
            nPercent = currCum - prevCumPct
            prevCumPct = currCum
        End If
        If nPercent < 0 Then nPercent = 0

        Select Case tabMonsters.Fields("AttType-" & x)
            Case 1, 3  ' melee (normal/rob)
                nAcc = tabMonsters.Fields("AttAcc-" & x)
                found = -1
                For i = 0 To uniqCount - 1
                    If uniqAcc(i) = nAcc Then
                        found = i
                        Exit For
                    End If
                Next i

                If found >= 0 Then
                    uniqPct(found) = uniqPct(found) + nPercent
                Else
                    uniqAcc(uniqCount) = nAcc
                    uniqPct(uniqCount) = nPercent
                    uniqCount = uniqCount + 1
                End If

                meleeTotalPct = meleeTotalPct + nPercent
                If nAcc > maxAcc Then maxAcc = nAcc

            Case 2  ' spell
                sSpellAttackTypes = sSpellAttackTypes & SpellAttackTypeEnum(GetSpellAttackType(tabMonsters.Fields("AttAcc-" & x)), True)
        End Select

        If tabMonsters.Fields("AttHitSpell-" & x) > 0 Then
            If SpellDoesDamage(tabMonsters.Fields("AttHitSpell-" & x), True) Then
                sSpellExtraTypes = sSpellExtraTypes & SpellAttackTypeEnum(GetSpellAttackType(tabMonsters.Fields("AttHitSpell-" & x)), True)
            End If
        End If
    End If
Next x

' Edge case: if nothing summed, treat total as 100 so threshold is meaningful
If meleeTotalPct < 1 Then meleeTotalPct = 100

' Find dominant (majority) accuracy group
domIdx = -1: domPct = 0: domAcc = 0
If uniqCount > 0 Then
    For i = 0 To uniqCount - 1
        ' choose the largest share; tie-breaker = higher accuracy
        If (uniqPct(i) > domPct) Or (uniqPct(i) = domPct And uniqAcc(i) > domAcc) Then
            domPct = uniqPct(i)
            domAcc = uniqAcc(i)
            domIdx = i
        End If
    Next i
End If

Dim hasMajority As Boolean
hasMajority = False
If domIdx >= 0 Then
    ' majority threshold: MAJ_THRESH_PCT% of meleeTotalPct (e.g., 70%)
    If domPct * 100& >= MAJ_THRESH_PCT * meleeTotalPct Then
        hasMajority = True
    End If
End If

' ----- Output to subitem #9 (Acc (Maj/Mx)) -----
nIndex = nIndex + 1 ' 9
If hasMajority Then
    ' If majority and max differ meaningfully, show both, else show one
    If Abs(maxAcc - domAcc) > 2 Then
        oLI.ListSubItems.Add (nIndex), "Acc (Maj/Mx)", CStr(domAcc) & "/" & CStr(maxAcc)
    Else
        oLI.ListSubItems.Add (nIndex), "Acc (Maj/Mx)", CStr(domAcc)
    End If
    oLI.ListSubItems(nIndex).Tag = domAcc
Else
    ' No majority: show Max (safest for gearing)
    oLI.ListSubItems.Add (nIndex), "Acc (Maj/Mx)", CStr(maxAcc)
    oLI.ListSubItems(nIndex).Tag = maxAcc
End If

nIndex = nIndex + 1 '10
nAvgDmg = -1
sTemp = ""
If tAvgLairInfo.nTotalLairs > 0 And tabMonsters.Fields("RegenTime") = 0 Then
    nAvgDmg = tAvgLairInfo.nAvgDmgLair
    sTemp = "*"
    bAsterisks = True
Else
    nAvgDmg = GetPreCalculatedMonsterDamage(tabMonsters.Fields("Number"), sTemp2, nParty)
End If
oLI.ListSubItems.Add (nIndex), "Damage", IIf(nAvgDmg > 0, Format(nAvgDmg, "#,##"), IIf(nAvgDmg = 0, 0, "?")) & sTemp
oLI.ListSubItems(nIndex).Tag = nAvgDmg
If nParty > 1 Then 'vs party
    If nMonsterDamageVsParty(tabMonsters.Fields("Number")) >= 0 Then oLI.ListSubItems(nIndex).ForeColor = RGB(193, 0, 232)
ElseIf bUseCharacter And nMonsterDamageVsChar(tabMonsters.Fields("Number")) >= 0 Then
    oLI.ListSubItems(nIndex).ForeColor = RGB(193, 0, 232)
End If

If InStr(1, tabMonsters.Fields("Summoned By"), "(lair)", vbTextCompare) > 0 Then
    nPossSpawns = InstrCount(tabMonsters.Fields("Summoned By"), "(lair)")
End If

sTemp = ""
nIndex = nIndex + 1 '11
If nNMRVer >= 1.83 And frmMain.optMonsterFilter(1).Value = True And lv.hWnd = frmMain.lvMonsters.hWnd Then
    'by lair (and exp by hour)
    bAsterisks = False
    If nPassEXP < 0 Or nPassRecovery < 0 Then
        'Call PopulateCharacterProfile(tChar)
        
        If tabMonsters.Fields("RegenTime") = 0 And tAvgLairInfo.nTotalLairs > 0 Then

            tExpInfo = CalcExpPerHour(tLastAvgLairInfo.nAvgExp, tLastAvgLairInfo.nAvgDelay, tLastAvgLairInfo.nMaxRegen, tLastAvgLairInfo.nTotalLairs, _
                            tLastAvgLairInfo.nPossSpawns, tLastAvgLairInfo.nRTK, tLastAvgLairInfo.nDamageOut, tChar.nHP, tChar.nHPRegen, _
                            tLastAvgLairInfo.nAvgDmgLair, tLastAvgLairInfo.nAvgHP, , tChar.nDamageThreshold, _
                            tChar.nSpellAttackCost, tChar.nSpellOverhead, tChar.nMaxMana, tChar.nManaRegen, tChar.nMeditateRate, _
                            tLastAvgLairInfo.nAvgWalk, tChar.nEncumPCT)

        ElseIf tabMonsters.Fields("RegenTime") > 0 Or InStr(1, tabMonsters.Fields("Summoned By"), "Room", vbTextCompare) > 0 Then
            
            If nDamageOut = -9999 Or (tChar.nParty = 1 And nSurpriseDamageOut = -9999 And nGlobalAttackTypeMME > a0_oneshot And bGlobalAttackBackstab = True) Then
                
                nDmgOut = GetDamageOutput(nMonsterNum)
                nDamageOut = IIf(nDmgOut(0) > -9990, nDmgOut(0), 0)
                nSurpriseDamageOut = IIf(nDmgOut(2) > -9990, nDmgOut(2), 0)
                
            End If
            
            tExpInfo = CalcExpPerHour(nExp, tabMonsters.Fields("RegenTime"), 1, -1, _
                            , , nDamageOut, tChar.nHP, tChar.nHPRegen, _
                            nAvgDmg, tabMonsters.Fields("HP"), tabMonsters.Fields("HPRegen"), tChar.nDamageThreshold, _
                            tChar.nSpellAttackCost, tChar.nSpellOverhead, tChar.nMaxMana, tChar.nManaRegen, tChar.nMeditateRate, _
                            0, tChar.nEncumPCT, , nSurpriseDamageOut)
        End If
    End If
    
    If nPassEXP >= 0 Then
        nExpPerHour = nPassEXP
    Else
        nExpPerHour = tExpInfo.nExpPerHour
        If nExpPerHour > 0 And nParty > 1 Then
            nExpPerHour = Round(nExpPerHour / nParty)
        End If
    End If
    
    If nPassRecovery >= 0 Then
        nTimeRecovering = nPassRecovery
    Else
        nTimeRecovering = tExpInfo.nTimeRecovering
    End If
    
    If nExpPerHour > 1000000 Then
        sTemp = Format((nExpPerHour / 1000000), "#,#.0") & " M"
    ElseIf nExpPerHour > 1000 Then
        sTemp = Format((nExpPerHour / 1000), "#,#.0") & " K"
    Else
        sTemp = IIf(nExpPerHour > 0, Format(RoundUp(nExpPerHour), "#,#"), "0")
    End If
    
    If nExpPerHour > 0 And nParty > 1 Then
        sTemp = sTemp & "/hr ea."
    Else
        sTemp = sTemp & "/hr"
    End If
    
    oLI.ListSubItems.Add (nIndex), "Exp/(Dmg+HP)", sTemp & IIf(bAsterisks, " *", "")
    oLI.ListSubItems(nIndex).Tag = nExpPerHour
    
    If nParty > 1 Then
        If nMonsterDamageVsParty(tabMonsters.Fields("Number")) >= 0 Then oLI.ListSubItems(nIndex).ForeColor = RGB(193, 0, 232)
    ElseIf bUseCharacter And nMonsterDamageVsChar(tabMonsters.Fields("Number")) >= 0 Then
        oLI.ListSubItems(nIndex).ForeColor = RGB(193, 0, 232)
    End If
    
ElseIf nExp > 0 Then

    If nAvgDmg > 0 Or nHP > 0 Then
        If nAvgDmg < 0 Then nAvgDmg = 0
        nExpDmgHP = Round(nExp / ((nAvgDmg * 2) + nHP), 2) * 100
    Else
        nExpDmgHP = nExp * 100
    End If
    
    oLI.ListSubItems.Add (nIndex), "Exp/(Dmg+HP)", IIf(nExpDmgHP > 0, Format(nExpDmgHP, "#,#"), 0) & IIf(bAsterisks, "*", "")
    oLI.ListSubItems(nIndex).Tag = nExpDmgHP
    
    If nParty > 1 Then
        If nMonsterDamageVsParty(tabMonsters.Fields("Number")) >= 0 Then oLI.ListSubItems(nIndex).ForeColor = RGB(193, 0, 232)
    ElseIf bUseCharacter And nMonsterDamageVsChar(tabMonsters.Fields("Number")) >= 0 Then
        oLI.ListSubItems(nIndex).ForeColor = RGB(193, 0, 232)
    End If
Else
    oLI.ListSubItems.Add (nIndex), "Exp/(Dmg+HP)", 0
    oLI.ListSubItems(nIndex).Tag = nExp
End If

nIndex = nIndex + 1 '12
If nNMRVer >= 1.83 Then
    If frmMain.optMonsterFilter(1).Value = True And lv.hWnd = frmMain.lvMonsters.hWnd Then
        'by lair - resting rate substituted here
        oLI.ListSubItems.Add (nIndex), "Lair Exp", Round(nTimeRecovering * 100) & "%"
        oLI.ListSubItems(nIndex).Tag = Round(nTimeRecovering * 100)
    Else
        oLI.ListSubItems.Add (nIndex), "Lair Exp", PutCommas(tabMonsters.Fields("AvgLairExp"))
        oLI.ListSubItems(nIndex).Tag = tabMonsters.Fields("AvgLairExp")
    End If
Else
    'script value (phased out in 1.83+)...
    nMaxLairsBeforeRegen = 36 'nTheoreticalMaxLairsPerRegenPeriod
    If nMonsterPossy(tabMonsters.Fields("Number")) > 0 Then nMaxLairsBeforeRegen = Round(nMaxLairsBeforeRegen / nMonsterPossy(tabMonsters.Fields("Number")), 2)
    
    If nPossSpawns < nMaxLairsBeforeRegen Then
        nLairPCT = Round(nPossSpawns / nMaxLairsBeforeRegen, 2)
    Else
        nLairPCT = 1
    End If
    
    nPossyPCT = 1
    If tabMonsters.Fields("RegenTime") = 0 And nLairPCT > 0 Then
        If nMonsterPossy(tabMonsters.Fields("Number")) > 0 Then
            nPossyPCT = 1 + ((nMonsterPossy(tabMonsters.Fields("Number")) - 1) / 5)
            If nPossyPCT > 3 Then nPossyPCT = 3
            If nPossyPCT < 1 Then nPossyPCT = 1
        End If
        
        If nAvgDmg > 0 Or nHP > 0 Then
            nExpDmgHP = Round(nExp / (((nAvgDmg * 2) + nHP) * nPossyPCT), 2) * 100
        Else
            nExpDmgHP = nExp
        End If
        
        nScriptValue = nExpDmgHP * nLairPCT
    End If
    
    If nScriptValue > 1000000000 Then
        oLI.ListSubItems.Add (nIndex), "Script Value", Format((nScriptValue / 1000000), "#,#M") & sTemp
    ElseIf nScriptValue > 1000000 Then
        oLI.ListSubItems.Add (nIndex), "Script Value", Format((nScriptValue / 1000), "#,#K") & sTemp
    Else
        oLI.ListSubItems.Add (nIndex), "Script Value", IIf(nScriptValue > 0, Format(RoundUp(nScriptValue), "#,#"), "0") & sTemp
    End If
    oLI.ListSubItems(nIndex).Tag = nScriptValue
    
    If nParty > 1 Then
        If nMonsterDamageVsParty(tabMonsters.Fields("Number")) >= 0 Then oLI.ListSubItems(nIndex).ForeColor = RGB(193, 0, 232)
    ElseIf nMonsterDamageVsChar(tabMonsters.Fields("Number")) >= 0 And bUseCharacter Then
        oLI.ListSubItems(nIndex).ForeColor = RGB(193, 0, 232)
    End If
End If

nIndex = nIndex + 1 '13
oLI.ListSubItems.Add (nIndex), "Lairs", IIf(nPossSpawns > 0, nPossSpawns, "")
oLI.ListSubItems(nIndex).Tag = nPossSpawns

If nNMRVer >= 1.82 Then
    nIndex = nIndex + 1 '14
    If nMonsterPossy(tabMonsters.Fields("Number")) > 0 Then
        If nMonsterSpawnChance(tabMonsters.Fields("Number")) > 0 Then 'only populated in nNMRVer >= 1.83
            oLI.ListSubItems.Add (nIndex), "Mobs/Spwn", nMonsterPossy(tabMonsters.Fields("Number")) & " / " & (nMonsterSpawnChance(tabMonsters.Fields("Number")) * 100) & "%"
            oLI.ListSubItems(nIndex).Tag = (nMonsterSpawnChance(tabMonsters.Fields("Number")) * 100) * nMonsterPossy(tabMonsters.Fields("Number"))
        Else
            oLI.ListSubItems.Add (nIndex), "#Mobs", nMonsterPossy(tabMonsters.Fields("Number"))
            oLI.ListSubItems(nIndex).Tag = nMonsterPossy(tabMonsters.Fields("Number"))
        End If
    Else
        If nNMRVer >= 1.83 Then
            oLI.ListSubItems.Add (nIndex), "Mobs/Spwn", ""
        Else
            oLI.ListSubItems.Add (nIndex), "#Mobs", ""
        End If
        oLI.ListSubItems(nIndex).Tag = 0
    End If
End If

nIndex = nIndex + 1 '15 (14 < 1.82)
oLI.ListSubItems.Add (nIndex), "Mag.", IIf(nMagicLVL > 0, nMagicLVL, "")
oLI.ListSubItems(nIndex).Tag = nMagicLVL

nIndex = nIndex + 1 '16 (15 < 1.82)
oLI.ListSubItems.Add (nIndex), "Undead", IIf(tabMonsters.Fields("Undead") > 0, "X", "")
oLI.ListSubItems(nIndex).Tag = tabMonsters.Fields("Undead")

nIndex = nIndex + 1 '17 (16 < 1.82)
If Len(sSpellExtraTypes) > 0 Then sSpellAttackTypes = sSpellAttackTypes & "+" & sSpellExtraTypes
If Len(sSpellAttackTypes) > 0 Then sSpellAttackTypes = SortLettersWithSeparator(sSpellAttackTypes, "+")
oLI.ListSubItems.Add (nIndex), "Spell Atk.", sSpellAttackTypes

skip:
Set oLI = Nothing

out:
On Error Resume Next
If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum
Exit Sub
error:
Call HandleError("AddMonster2LV")
Resume out:
End Sub

Public Sub AddShop2LV(lv As ListView)

On Error GoTo error:

Dim oLI As ListItem, sName As String
    
    sName = tabShops.Fields("Name")
    If sName = "" Or Left(LCase(sName), 3) = "sdf" Then GoTo skip:
    If sName = "Leave this blank" Then GoTo skip:
    
    Set oLI = lv.ListItems.Add()
    oLI.Text = tabShops.Fields("Number")
    
    oLI.ListSubItems.Add (1), "Name", sName
    oLI.ListSubItems.Add (2), "Type", GetShopType(tabShops.Fields("ShopType"))
    
skip:
Set oLI = Nothing


out:
On Error Resume Next
Exit Sub
error:
Call HandleError("AddShop2LV")
Resume out:
End Sub

Public Sub AddClass2LV(lv As ListView)

On Error GoTo error:

Dim oLI As ListItem, x As Integer, sAbil As String
    
    If tabClasses.Fields("Name") = "" Then GoTo skip:
    
    Set oLI = lv.ListItems.Add()
    oLI.Text = tabClasses.Fields("Number")
    
    oLI.ListSubItems.Add (1), "Name", tabClasses.Fields("Name")
    oLI.ListSubItems.Add (2), "Exp%", (val(tabClasses.Fields("ExpTable")) + 100) & "%"
    oLI.ListSubItems.Add (3), "Weapon", GetClassWeaponType(tabClasses.Fields("WeaponType"))
    oLI.ListSubItems.Add (4), "Armour", GetArmourType(tabClasses.Fields("ArmourType"))
    oLI.ListSubItems.Add (5), "Magic", GetMagery(tabClasses.Fields("MageryType"), tabClasses.Fields("MageryLVL"))
    oLI.ListSubItems.Add (6), "Cmbt", tabClasses.Fields("CombatLVL") - 2
    oLI.ListSubItems.Add (7), "HP", tabClasses.Fields("MinHits") & "-" & (tabClasses.Fields("MinHits") + tabClasses.Fields("MaxHits"))
    
    For x = 0 To 9
        Select Case tabClasses.Fields("Abil-" & x)
            Case 0:
            Case 59: 'class ok
            Case Else:
                If sAbil <> "" Then sAbil = sAbil & ", "
                sAbil = sAbil & GetAbilityStats(tabClasses.Fields("Abil-" & x), tabClasses.Fields("AbilVal-" & x))
                If Right(sAbil, 2) = ", " Then sAbil = Left(sAbil, Len(sAbil) - 2)
        End Select
    Next

    oLI.ListSubItems.Add (8), "Abilities", sAbil
    
skip:
Set oLI = Nothing


out:
On Error Resume Next
Exit Sub
error:
Call HandleError("AddClass2LV")
Resume out:
End Sub



Public Sub RaceColorCode(lv As ListView)
On Error GoTo error:
Dim oLI As ListItem, x As Integer
Dim Stat(1 To 6, 1 To 2) As Integer, Min(1 To 6) As Integer, Max(1 To 6) As Integer, nRaces As Integer
'1-6 = str, int, wis, agi, hea, cha

'stat, 1 = min
'stat, 2 = max
'min, total min
'max, total max
'then get avg

For Each oLI In lv.ListItems
    
    tabRaces.Index = "pkRaces"
    tabRaces.Seek "=", val(oLI.Text)
    If tabRaces.NoMatch = False Then
        
        nRaces = nRaces + 1
        Stat(1, 1) = val(tabRaces.Fields("mSTR"))
        Stat(2, 1) = val(tabRaces.Fields("mINT"))
        Stat(3, 1) = val(tabRaces.Fields("mWIL"))
        Stat(4, 1) = val(tabRaces.Fields("mAGL"))
        Stat(5, 1) = val(tabRaces.Fields("mHEA"))
        Stat(6, 1) = val(tabRaces.Fields("mCHM"))
        Stat(1, 2) = val(tabRaces.Fields("xSTR"))
        Stat(2, 2) = val(tabRaces.Fields("xINT"))
        Stat(3, 2) = val(tabRaces.Fields("xWIL"))
        Stat(4, 2) = val(tabRaces.Fields("xAGL"))
        Stat(5, 2) = val(tabRaces.Fields("xHEA"))
        Stat(6, 2) = val(tabRaces.Fields("xCHM"))
        
        For x = 1 To 6
            Min(x) = Min(x) + Stat(x, 1)
            Max(x) = Max(x) + Stat(x, 2)
        Next x
        
    End If
Next oLI

If nRaces = 0 Then Exit Sub

For x = 1 To 6
    Stat(x, 1) = Min(x) / nRaces 'avg min
    Stat(x, 2) = Max(x) / nRaces 'avg max
    Stat(x, 1) = Stat(x, 1) + Stat(x, 2) 'avg total
Next

For Each oLI In lv.ListItems
    
    tabRaces.Index = "pkRaces"
    tabRaces.Seek "=", val(oLI.Text)
    If tabRaces.NoMatch = False Then
        
        
        If val(tabRaces.Fields("mSTR")) + val(tabRaces.Fields("xSTR")) < Stat(1, 1) - 20 Then oLI.ListSubItems(4).ForeColor = &H80&
        If val(tabRaces.Fields("mINT")) + val(tabRaces.Fields("xINT")) < Stat(2, 1) - 20 Then oLI.ListSubItems(5).ForeColor = &H80&
        If val(tabRaces.Fields("mWIL")) + val(tabRaces.Fields("xWIL")) < Stat(3, 1) - 20 Then oLI.ListSubItems(6).ForeColor = &H80&
        If val(tabRaces.Fields("mAGL")) + val(tabRaces.Fields("xAGL")) < Stat(4, 1) - 20 Then oLI.ListSubItems(7).ForeColor = &H80&
        If val(tabRaces.Fields("mHEA")) + val(tabRaces.Fields("xHEA")) < Stat(5, 1) - 20 Then oLI.ListSubItems(8).ForeColor = &H80&
        If val(tabRaces.Fields("mCHM")) + val(tabRaces.Fields("xCHM")) < Stat(6, 1) - 20 Then oLI.ListSubItems(9).ForeColor = &H80&
        
        If val(tabRaces.Fields("mSTR")) + val(tabRaces.Fields("xSTR")) > Stat(1, 1) + 20 Then oLI.ListSubItems(4).ForeColor = &H8000&
        If val(tabRaces.Fields("mINT")) + val(tabRaces.Fields("xINT")) > Stat(2, 1) + 20 Then oLI.ListSubItems(5).ForeColor = &H8000&
        If val(tabRaces.Fields("mWIL")) + val(tabRaces.Fields("xWIL")) > Stat(3, 1) + 20 Then oLI.ListSubItems(6).ForeColor = &H8000&
        If val(tabRaces.Fields("mAGL")) + val(tabRaces.Fields("xAGL")) > Stat(4, 1) + 20 Then oLI.ListSubItems(7).ForeColor = &H8000&
        If val(tabRaces.Fields("mHEA")) + val(tabRaces.Fields("xHEA")) > Stat(5, 1) + 20 Then oLI.ListSubItems(8).ForeColor = &H8000&
        If val(tabRaces.Fields("mCHM")) + val(tabRaces.Fields("xCHM")) > Stat(6, 1) + 20 Then oLI.ListSubItems(9).ForeColor = &H8000&
        
    End If
Next

out:
On Error Resume Next
tabRaces.MoveFirst
Set oLI = Nothing
Exit Sub
error:
Call HandleError("RaceColorCode")
Resume out:
End Sub

Public Function SearchLV(ByVal KeyCode As Integer, oLVW As ListView, oTXT As TextBox) As Boolean
Dim i As Long, SearchStart As Long, SelectText As String, bSearchAgain As Boolean
Dim nCIndex As Integer

On Error GoTo error:

If oLVW.ListItems.Count < 1 Then Exit Function

If KeyCode = vbKeyUp Then Exit Function
If KeyCode = vbKeyLeft Then Exit Function
If KeyCode = vbKeyBack Then Exit Function
If KeyCode = vbKeyDown Then oLVW.SetFocus
If KeyCode = vbKeyRight Then bSearchAgain = True
If KeyCode = vbKeyShift Then Exit Function 'shift
If KeyCode = vbKeyControl Then Exit Function 'control
If KeyCode = 18 Then Exit Function 'alt
If KeyCode = vbKeyTab Then Exit Function 'alt

If oTXT.Text = "" Then Exit Function
SelectText = oTXT.Text

If bSearchAgain = True Then 'searching for next?
    SearchStart = oLVW.SelectedItem.Index + 1
Else
    SearchStart = 1
End If

If Not SearchStart + 1 <= oLVW.ListItems.Count Then Exit Function 'if it's the last item in the list

nCIndex = oLVW.SelectedItem.Index

For i = SearchStart To oLVW.ListItems.Count
    If Not InStr(1, LCase(oLVW.ListItems(i).ListSubItems(1).Text), LCase(SelectText)) = 0 Then
        If Not i = nCIndex Then
            SearchLV = True
            nCIndex = i
        End If
        Exit For
    End If
Next
        
If SearchLV Then
    For i = 1 To oLVW.ListItems.Count
        oLVW.ListItems(i).Selected = False
    Next
    Set oLVW.SelectedItem = oLVW.ListItems(nCIndex)
    oLVW.SelectedItem.EnsureVisible
End If

Exit Function
error:
Call HandleError
End Function

Public Sub CopyWholeLVtoClipboard(lv As ListView, Optional ByVal UsePeriods As Boolean)
On Error GoTo error:
Dim oLI As ListItem, oLSI As ListSubItem, oCH As ColumnHeader
Dim str As String, x As Integer, sSpacer As String, nLongText() As Integer
    
str = ""
sSpacer = IIf(UsePeriods, ".", " ")

ReDim nLongText(0 To lv.ColumnHeaders.Count - 1)

'find longest text(s)
For Each oLI In lv.ListItems
    If Len(oLI.Text) > nLongText(0) Then nLongText(0) = Len(oLI.Text)
    x = 1
    For Each oLSI In oLI.ListSubItems
        If Len(oLSI.Text) > nLongText(x) Then nLongText(x) = Len(oLSI.Text)
        x = x + 1
    Next
Next
'put on 3 spaces
For x = 0 To UBound(nLongText())
    nLongText(x) = nLongText(x) + 3
Next

x = 0
For Each oCH In lv.ColumnHeaders
    str = str & oCH.Text
    str = str & " " & String(nLongText(x) - Len(oCH.Text), " ") & " "
    x = x + 1
Next

str = str & vbCrLf

For Each oLI In lv.ListItems
    str = str & oLI.Text
    str = str & " " & String(nLongText(0) - Len(oLI.Text), sSpacer) & " "
    
    x = 1
    For Each oLSI In oLI.ListSubItems
        str = str & oLSI.Text
        If Not x = oLI.ListSubItems.Count Then str = str & " " & String(nLongText(x) - Len(oLSI.Text), sSpacer) & " "
        x = x + 1
    Next
    str = str & vbCrLf
Next

If Not str = "" Then
    'Clipboard.clear
    'Clipboard.SetText str
    Call SetClipboardText(str)
End If

Set oLI = Nothing
Set oLSI = Nothing
Set oCH = Nothing

Exit Sub
error:
Call HandleError("CopyWholeLVtoClip")
Set oLI = Nothing
Set oLSI = Nothing
Set oCH = Nothing
End Sub
Public Sub CopyLVLinetoClipboard(lv As ListView, Optional DetailTB As TextBox, _
    Optional LocationLV As ListView, Optional ByVal nExcludeColumn As Integer = -1, Optional bNameOnly As Boolean = False)
On Error GoTo error:
Dim oLI As ListItem, oLI2 As ListItem, oCH As ColumnHeader
Dim str As String, x As Integer, nCount As Integer

If lv.ListItems.Count < 1 Then Exit Sub

nCount = 1
For Each oLI In lv.ListItems
    If oLI.Selected Then
        If nCount > 100 Then GoTo done:
        If nCount > 1 Then
            If bNameOnly Then
                str = str & ", "
            Else
                str = str & vbCrLf & vbCrLf
            End If
        End If
        
        x = 0
        For Each oCH In lv.ColumnHeaders
            If Not x = nExcludeColumn Then
                If bNameOnly Then
                    If (lv.name = "lvMapLoc" Or lv.name = "lvSpellLoc" Or lv.name = "lvShopLoc" Or lv.name = "lvSpellCompareLoc") And x = 0 Then
                        If InStr(1, oLI.Text, ":", vbTextCompare) > 0 Then
                            str = str & Trim(Mid(oLI.Text, InStr(1, oLI.Text, ":", vbTextCompare) + 1, 999))
                        Else
                            str = str & oLI.Text
                        End If
                    ElseIf oCH.Text = "Name" Then
                        If x = 0 Then
                            str = str & oLI.Text
                        Else
                            str = str & oLI.SubItems(x)
                        End If
                    ElseIf Right(lv.name, 3) = "Loc" And oLI.ListSubItems.Count > 0 Then
                        If InStr(1, oLI.SubItems(x), ":", vbTextCompare) > 0 Then
                            str = str & Trim(Mid(oLI.SubItems(x), InStr(1, oLI.SubItems(x), ":", vbTextCompare) + 1, 999))
                        Else
                            str = str & oLI.SubItems(x)
                        End If
                    End If
                Else
                    If Not x = 0 Then str = str & ", "
                    
                    str = str & oCH.Text & ": "
                    'If Len(oCH.Text) <= 9 Then str = str & String(10 - Len(oCH.Text), " ")
                    
                    If x = 0 Then
                        str = str & oLI.Text
                    Else
                        str = str & oLI.SubItems(x)
                    End If
                End If
            End If
            x = x + 1
        Next oCH
        
        Select Case lv.name
            Case "lvWeapons":
                Call frmMain.lvWeapons_ItemClick(oLI)
            Case "lvArmour":
                Call frmMain.lvArmour_ItemClick(oLI)
            Case "lvSpells":
                Call frmMain.lvSpells_ItemClick(oLI)
            Case "lvWeaponCompare":
                Call frmMain.lvWeaponCompare_ItemClick(oLI)
            Case "lvArmourCompare":
                Call frmMain.lvArmourCompare_ItemClick(oLI)
            Case "lvSpellCompare":
                Call frmMain.lvSpellCompare_ItemClick(oLI)
        End Select
        
        If Not bNameOnly And Not DetailTB Is Nothing Then
            If Not DetailTB.Text = "" Then
                str = str & vbCrLf & ">> " & Replace(DetailTB.Text, vbCrLf & "Compared to", ">> Compared to")
            End If
        End If
        
        If Not bNameOnly And Not LocationLV Is Nothing Then
            If LocationLV.ListItems.Count > 0 Then
                If LocationLV.ListItems.Count > 5 Then
                    str = str & vbCrLf & "References--" & vbCrLf
                Else
                    str = str & vbCrLf & ">> Refs: "
                End If
                
                x = 1
                For Each oLI2 In LocationLV.ListItems
                    If LocationLV.ListItems.Count > 5 Then
                        If x > 1 Then str = str & vbCrLf
                    Else
                        If x > 1 Then str = str & ", "
                    End If
                    
                    str = str & oLI2.Text
                    
                    If oLI2.ListSubItems.Count >= 1 Then
                        If Len(Trim(oLI2.Text)) > 0 Then str = str & ": "
                        str = str & oLI2.ListSubItems(1).Text
                    End If
                    x = x + 1
                Next oLI2
                Set oLI2 = Nothing
            End If
        End If
        nCount = nCount + 1
    End If
    Set oLI = Nothing
    DoEvents
Next oLI

done:
If Not str = "" Then
    'Clipboard.clear
    'Clipboard.SetText str
    Call SetClipboardText(str)
End If

out:
Set oLI = Nothing
Set oLI2 = Nothing
Set oCH = Nothing

Exit Sub
error:
Call HandleError("CopyLVLinetoClip")
Resume out:
End Sub

Public Sub GetLocations(ByVal sLoc As String, lv As ListView, _
    Optional bDontClear As Boolean, Optional ByVal sHeader As String, _
    Optional ByVal nAuxValue As Long, Optional ByVal bTwoColumns As Boolean, _
    Optional ByVal bDontSort As Boolean, Optional ByVal bPercentColumn As Boolean, _
    Optional ByVal sFooter As String, Optional ByVal nLimit As Integer)
On Error GoTo error:
Dim sLook As String, sChar As String, sTest As String, oLI As ListItem, sPercent As String
Dim x As Integer, y1 As Integer, y2 As Integer, z As Integer, nValue As Long, x2 As Integer
Dim sLocation As String, nPercent As Currency, nPercent2 As Currency, sTemp As String, nSpawnChance As Currency
Dim sDisplayFooter As String, sLairRegex As String, sRoomKey As String, nMarkup As Integer
Dim tMatches() As RegexMatches, nMaxRegen As Integer, sGroupIndex As String, tLairInfo As LairInfoType
Dim tValue As tItemValue, sShopValue As String, nTemp As Double

Dim nCount As Integer

sDisplayFooter = sFooter

If Not bDontClear Then lv.ListItems.clear
If bDontSort Then lv.Sorted = False

If Len(sLoc) < 5 Then Exit Sub

sLairRegex = "Group\(lair\): (\d+)\/(\d+)"
If nNMRVer >= 1.82 Then sLairRegex = "\[(\d+)\]" & sLairRegex
If nNMRVer >= 1.83 Then sLairRegex = "\[([\d\-]+)\]" & sLairRegex

sTest = LCase(sLoc)

For z = 1 To 12
    
    'now that a regex function has been built, this would be much better handled that way.........
    x = 1
    Select Case z
        Case 1: sLook = "room "
        Case 2: sLook = "monster #"
        Case 3: sLook = "textblock #"
        Case 4: sLook = "textblock(rndm) #"
        Case 5: sLook = "item #"
        Case 6: sLook = "spell #"
        Case 7: sLook = "shop #"
        Case 8: sLook = "shop(sell) #"
        Case 9: sLook = "shop(nogen) #"
        Case 10: sLook = "group(lair): "
        Case 11: sLook = "group: "
        Case 12: sLook = "npc #"
    End Select

checknext:
    If nLimit > 0 And nCount > nLimit Then GoTo finish:
    
    sPercent = ""
    If Not InStr(x, sTest, sLook) = 0 Then
        
        x = InStr(x, sTest, sLook) 'sets x to the position of the string we're looking for
        
'        If z = 10 Then
'            y1 = x + 1
'            GoTo nonumber:
'        End If
        
        y1 = x + Len(sLook) 'len of string searching (to position y1 at first number)
        y2 = 0
nextnumber:
        sChar = Mid(sTest, y1 + y2, 1)
        Select Case sChar
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "/", ".":
                If Not y1 + y2 - 1 = Len(sTest) Then
                    y2 = y2 + 1
                    GoTo nextnumber:
                End If
            Case "+": 'end of string
                Exit Sub
            Case "(": 'percent
                x2 = InStr(y1 + y2, sTest, ")")
                If Not x2 = 0 Then
                    sPercent = " " & Mid(sTest, y1 + y2, x2 - y1 - y2 + 1)
                End If
            Case Else:
        End Select
        
        If y2 = 0 Then
            'if there were no numbers after the string
            x = y1
            GoTo checknext:
        End If
        
        If Not z = 1 Or z = 10 Or z = 11 Then 'not room or group
            nValue = val(Mid(sTest, y1, y2))
        End If

nonumber:
        If bPercentColumn Then
            nPercent = ExtractNumbersFromString(sPercent)
            
            If InStr(1, sFooter, "%)", vbTextCompare) > 0 Then
                nPercent2 = ExtractNumbersFromString(Mid(sFooter, InStr(1, sFooter, "(", vbTextCompare)))
                If nPercent2 > 0 Then
                    sDisplayFooter = Replace(sFooter, "(" & nPercent2 & "%)", "")
                    If nPercent > 0 Then
                        nPercent = (nPercent * nPercent2) / 100
                    Else
                        nPercent = nPercent2
                    End If
                End If
            End If
            
            If nPercent > 1 Then
                nPercent = Round(nPercent)
            Else
                nPercent = Round(nPercent, 2)
            End If
            'sPercent = " (" & nPercent & "%)"
        End If
        
        Select Case z
            Case 1: '"room "
                If nLimit > 0 Then nCount = nCount + 1
                If nLimit > 0 And nCount > nLimit Then GoTo skip:
                sLocation = "Room: "
                If Not sHeader = "" Then
                    If InStr(1, sHeader, ":", vbTextCompare) = 0 Then
                        sLocation = sLocation & sHeader
                    Else
                        sLocation = sHeader
                    End If
                End If
                
                Set oLI = lv.ListItems.Add()
                If bPercentColumn Then
                    oLI.Text = ""
                    If nPercent > 0 Then oLI.Text = nPercent & "%"
                    oLI.Tag = nPercent
                    oLI.ListSubItems.Add 1, , sLocation & GetRoomName(Mid(sTest, y1, y2), , , bHideRecordNumbers) & sDisplayFooter
                    
                ElseIf bTwoColumns Then
                    oLI.Text = sLocation
                    oLI.ListSubItems.Add 1, , GetRoomName(Mid(sTest, y1, y2), , , bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = "Room"
                Else
                    oLI.Text = sLocation & GetRoomName(Mid(sTest, y1, y2), , , bHideRecordNumbers) & sPercent & sDisplayFooter
                End If
                
                If nAuxValue > 0 And (bDontClear = True Or Len(sHeader) > 0) Then
                    If bTwoColumns Or bPercentColumn Then
                        oLI.ListSubItems(1).Tag = nAuxValue
                    Else
                        oLI.Tag = nAuxValue
                    End If
                Else
                    If bTwoColumns Or bPercentColumn Then
                        oLI.ListSubItems(1).Tag = Mid(sTest, y1, y2)
                    Else
                        oLI.Tag = Mid(sTest, y1, y2)
                    End If
                End If
                
            Case 2: '"monster #"
                If nLimit > 0 Then nCount = nCount + 1
                If nLimit > 0 And nCount > nLimit Then GoTo skip:
                sLocation = "Monster: "
                Set oLI = lv.ListItems.Add()
                If bPercentColumn Then
                    oLI.Text = ""
                    If nPercent > 0 Then oLI.Text = nPercent & "%"
                    oLI.Tag = nPercent
                    oLI.ListSubItems.Add 1, , sLocation & sHeader & GetMonsterName(nValue, bHideRecordNumbers) & sDisplayFooter
                    oLI.ListSubItems(1).Tag = nValue
                ElseIf bTwoColumns Then
                    oLI.Text = sLocation & sHeader
                    oLI.ListSubItems.Add 1, , GetMonsterName(nValue, bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = "monster"
                    oLI.ListSubItems(1).Tag = nValue
                Else
                    oLI.Text = sLocation & sHeader & GetMonsterName(nValue, bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = nValue
                End If:
                
            Case 3: '"textblock #"
                If nLimit > 0 Then nCount = nCount + 1
                If nLimit > 0 And nCount > nLimit Then GoTo skip:
                sLocation = "Textblock "
                Set oLI = lv.ListItems.Add()
                If bPercentColumn Then
                    oLI.Text = ""
                    If nPercent > 0 Then oLI.Text = nPercent & "%"
                    oLI.Tag = nPercent
                    oLI.ListSubItems.Add 1, , sLocation & sHeader & nValue & sDisplayFooter
                    oLI.ListSubItems(1).Tag = nValue
                    
                ElseIf bTwoColumns Then
                    oLI.Text = sLocation & sHeader
                    oLI.ListSubItems.Add 1, , nValue & sPercent & sDisplayFooter
                    oLI.Tag = "textblock"
                    oLI.ListSubItems(1).Tag = nValue
                Else
                    oLI.Text = sLocation & sHeader & nValue & sPercent & sDisplayFooter
                    oLI.Tag = nValue
                End If
                
            Case 4: '"textblock(rndm) #"
                If nLimit > 0 Then nCount = nCount + 1
                If nLimit > 0 And nCount > nLimit Then GoTo skip:
                sLocation = "Textblock "
                Set oLI = lv.ListItems.Add()
                If bPercentColumn Then
                    oLI.Text = ""
                    If nPercent > 0 Then oLI.Text = nPercent & "%"
                    oLI.Tag = nPercent
                    oLI.ListSubItems.Add 1, , sLocation & sHeader & nValue & sDisplayFooter
                    oLI.ListSubItems(1).Tag = nValue
                ElseIf bTwoColumns Then
                    oLI.Text = sLocation & sHeader
                    oLI.ListSubItems.Add 1, , nValue & sPercent & sDisplayFooter
                    oLI.Tag = "textblock"
                    oLI.ListSubItems(1).Tag = nValue
                Else
                    oLI.Text = sLocation & sHeader & nValue & sPercent & sDisplayFooter
                    oLI.Tag = nValue
                End If
                
            Case 5: '"item #"
                If nLimit > 0 Then nCount = nCount + 1
                If nLimit > 0 And nCount > nLimit Then GoTo skip:
                sLocation = "Item: "
                Set oLI = lv.ListItems.Add()
                If bPercentColumn Then
                    oLI.Text = ""
                    If nPercent > 0 Then oLI.Text = nPercent & "%"
                    oLI.Tag = nPercent
                    oLI.ListSubItems.Add 1, , sLocation & sHeader & GetItemName(nValue, bHideRecordNumbers) & sDisplayFooter
                    oLI.ListSubItems(1).Tag = nValue
                ElseIf bTwoColumns Then
                    oLI.Text = sLocation & sHeader
                    oLI.ListSubItems.Add 1, , GetItemName(nValue, bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = "item"
                    oLI.ListSubItems(1).Tag = nValue
                Else
                    oLI.Text = sLocation & sHeader & GetItemName(nValue, bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = nValue
                End If
                
                If ItemIsChest(nValue) And sHeader = "" And sFooter = "" Then
                    Call GetLocations(tabItems.Fields("Obtained From"), lv, True, , , , True, bPercentColumn, " -> " & tabItems.Fields("Name") & sPercent, nLimit - nCount)
                End If
                
            Case 6: '"spell #"
                If nLimit > 0 Then nCount = nCount + 1
                If nLimit > 0 And nCount > nLimit Then GoTo skip:
                sLocation = "Spell: "
                Set oLI = lv.ListItems.Add()
                If bPercentColumn Then
                    oLI.Text = ""
                    If nPercent > 0 Then oLI.Text = nPercent & "%"
                    oLI.Tag = nPercent
                    oLI.ListSubItems.Add 1, , sLocation & sHeader & GetSpellName(nValue, bHideRecordNumbers) & sDisplayFooter
                    oLI.ListSubItems(1).Tag = nValue
                ElseIf bTwoColumns Then
                    oLI.Text = sLocation & sHeader
                    oLI.ListSubItems.Add 1, , GetSpellName(nValue, bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = "spell"
                    oLI.ListSubItems(1).Tag = nValue
                Else
                    oLI.Text = sLocation & sHeader & GetSpellName(nValue, bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = nValue
                End If
                
            Case 7: '"shop #"
                If nLimit > 0 Then nCount = nCount + 1
                If nLimit > 0 And nCount > nLimit Then GoTo skip:
                sShopValue = ""
                If nValue > 0 And nAuxValue > 0 Then
                    nMarkup = GetShopMarkup(nValue)
                    tValue = GetItemValue(nAuxValue, IIf(frmMain.chkGlobalFilter.Value = 1, val(frmMain.txtCharStats(5).Tag), 0), nMarkup)
                    If tValue.nBaseCost > 0 Then sShopValue = " - Value: " & tValue.sFriendlyBuyShort & "/" & tValue.sFriendlySellShort
                End If
                
                If bPercentColumn And nAuxValue > 0 Then
                    nPercent = GetItemShopRegenPCT(nValue, nAuxValue)
                    If nPercent > 0 Then
                        sTemp = GetShopLocation(nValue)
                        sTemp = Join(Split(sTemp, ","), "(" & nPercent & "%),")
                        If Not Right(sTemp, 2) = "%)" Then sTemp = sTemp & "(" & nPercent & "%)"
                        Call GetLocations(sTemp, lv, True, "Shop: ", nValue, , , bPercentColumn, sShopValue, nLimit - nCount)
                    Else
                        Call GetLocations(GetShopLocation(nValue), lv, True, "Shop: ", nValue, , , bPercentColumn, sShopValue, nLimit - nCount)
                    End If
                Else
                    Call GetLocations(GetShopLocation(nValue), lv, True, "Shop: ", nValue, , , , sShopValue, nLimit - nCount)
                End If
                
            Case 8: '"shop(sell) #"
                If nLimit > 0 Then nCount = nCount + 1
                If nLimit > 0 And nCount > nLimit Then GoTo skip:
                sShopValue = ""
                If nValue > 0 And nAuxValue > 0 Then
                    nMarkup = GetShopMarkup(nValue)
                    tValue = GetItemValue(nAuxValue, IIf(frmMain.chkGlobalFilter.Value = 1, val(frmMain.txtCharStats(5).Tag), 0), nMarkup, , True)
                    If tValue.nCopperSell > 0 Then sShopValue = " - Value: " & tValue.sFriendlySellShort
                End If
                Call GetLocations(GetShopLocation(nValue), lv, True, "Shop (sell): ", nValue, , , bPercentColumn, sShopValue, nLimit - nCount)
                
            Case 9: '"shop(nogen) #"
                If nLimit > 0 Then nCount = nCount + 1
                If nLimit > 0 And nCount > nLimit Then GoTo skip:
                sShopValue = ""
                If nValue > 0 And nAuxValue > 0 Then
                    nMarkup = GetShopMarkup(nValue)
                    tValue = GetItemValue(nAuxValue, IIf(frmMain.chkGlobalFilter.Value = 1, val(frmMain.txtCharStats(5).Tag), 0), nMarkup)
                    If tValue.nBaseCost > 0 Then sShopValue = " - Value: " & tValue.sFriendlyBuyShort & "/" & tValue.sFriendlySellShort
                End If
                Call GetLocations(GetShopLocation(nValue), lv, True, "Shop (nogen): ", nValue, , , bPercentColumn, sShopValue, nLimit - nCount)
                
            Case 10: 'group (lair)
                sLocation = "Group(Lair): "
                nMaxRegen = 0
                sGroupIndex = "0-0-0"
                nSpawnChance = 0
                If nNMRVer >= 1.83 Then tLairInfo = GetLairInfo("") 'reset
                
                If (y1 - Len(sLook) - 2) > 0 Then
                    For x2 = (y1 - Len(sLook) - 1) To 1 Step -1
                        sChar = Mid(sTest, x2, 1)
                        If sChar = "," Then Exit For
                    Next
                    sTemp = Mid(sTest, x2 + 1, y1 + y2 - x2 - 1)
                Else
                    sTemp = Mid(sTest, y1 - Len(sLook), y1 + y2 - 1)
                End If
                
                tMatches() = RegExpFindv2(sTemp, sLairRegex, False)
                If UBound(tMatches()) > 0 Or Len(tMatches(0).sFullMatch) > 0 Then
                    If nNMRVer >= 1.83 Then
                        '[7-8-9][6]Group(lair): 1/2345
                        sGroupIndex = tMatches(0).sSubMatches(0)
                        nMaxRegen = val(tMatches(0).sSubMatches(1))
                        sRoomKey = tMatches(0).sSubMatches(2) & "/" & tMatches(0).sSubMatches(3)
                        tLairInfo = GetLairInfo(sGroupIndex, nMaxRegen)
                        If tLairInfo.nMobs > 0 Then
                            nSpawnChance = Round(1 - (1 - (1 / tLairInfo.nMobs)) ^ nMaxRegen, 2) * 100
                            '1 - (1 - (x / y)) ^ z
                            '(x / y) == (1) of (y) totalmobs
                            'z == maxregen (chance to spawn)
                        End If
                    ElseIf nNMRVer >= 1.82 Then
                        '[6]Group(lair): 1/2345
                        nMaxRegen = val(tMatches(0).sSubMatches(0))
                        sRoomKey = tMatches(0).sSubMatches(1) & "/" & tMatches(0).sSubMatches(2)
                    Else
                        'Group(lair): 1/2345
                        sRoomKey = tMatches(0).sSubMatches(0) & "/" & tMatches(0).sSubMatches(1)
                    End If
                Else
                    sRoomKey = Mid(sTest, y1, y2)
                End If
                
                If nSpawnChance > 0 Then
                    sLocation = "Lair " & nMaxRegen & " (" & nSpawnChance & "%)"
                ElseIf nMaxRegen > 0 Then
                    sLocation = "Lair " & nMaxRegen
                Else
                    sLocation = "Lair"
                End If
                
                nTemp = GetRoomRegen(sRoomKey)
                If nTemp > 0 Then sLocation = sLocation & " - " & nTemp & "m"
                
                Set oLI = lv.ListItems.Add()
                If bPercentColumn Then
                    oLI.Text = ""
                    If nPercent > 0 Then oLI.Text = nPercent & "%"
                    oLI.Tag = nPercent
                    oLI.ListSubItems.Add 1, , sLocation & GetRoomName(sRoomKey, , , bHideRecordNumbers) & sDisplayFooter
                    oLI.ListSubItems(1).Tag = sRoomKey
                ElseIf bTwoColumns Then
                    oLI.Text = sLocation
                    sTemp = GetRoomName(sRoomKey, , , bHideRecordNumbers) & sPercent & sDisplayFooter
                    If tLairInfo.nAvgExp > 0 Then sTemp = sTemp & ", Exp: " & FormatNumber(tLairInfo.nAvgExp * tLairInfo.nMaxRegen, 0, , , vbTrue)
                    'If tLairInfo.nScriptValue > 0 Then sTemp = sTemp & ", SV: " & FormatNumber(tLairInfo.nScriptValue, 0, , , vbTrue)
                    If tLairInfo.nAvgDmgLair > 0 And tLairInfo.nDamageOut > 0 And (nGlobalAttackTypeMME > a0_oneshot And nGlobalAttackTypeMME <> a5_Manual) Then
                        sTemp = sTemp & ", Dmg In/Out: " & Round(tLairInfo.nAvgDmgLair) & "/" & tLairInfo.nDamageOut
                    ElseIf tLairInfo.nAvgDmgLair > 0 Then
                        sTemp = sTemp & ", Dmg In: " & Round(tLairInfo.nAvgDmgLair)
                    ElseIf tLairInfo.nDamageOut > 0 And (nGlobalAttackTypeMME > a0_oneshot And nGlobalAttackTypeMME <> a5_Manual) Then
                        sTemp = sTemp & ", Dmg Out: " & tLairInfo.nDamageOut
                    End If
                    oLI.ListSubItems.Add 1, , sTemp
                    oLI.Tag = "Room"
                    oLI.ListSubItems(1).Tag = sRoomKey
                Else
                    oLI.Text = sLocation & GetRoomName(sRoomKey, , , bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = sRoomKey
                End If
                
            Case 11: 'group
                If nLimit > 0 Then nCount = nCount + 1
                If nLimit > 0 And nCount > nLimit Then GoTo skip:
                sLocation = "Group: "
                Set oLI = lv.ListItems.Add()
                If bPercentColumn Then
                    oLI.Text = ""
                    If nPercent > 0 Then oLI.Text = nPercent & "%"
                    oLI.Tag = nPercent
                    oLI.ListSubItems.Add 1, , sLocation & GetRoomName(Mid(sTest, y1, y2), , , bHideRecordNumbers) & sDisplayFooter
                    oLI.ListSubItems(1).Tag = Mid(sTest, y1, y2)
                ElseIf bTwoColumns Then
                    oLI.Text = sLocation
                    oLI.ListSubItems.Add 1, , GetRoomName(Mid(sTest, y1, y2), , , bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = "Room"
                    oLI.ListSubItems(1).Tag = Mid(sTest, y1, y2)
                Else
                    oLI.Text = sLocation & GetRoomName(Mid(sTest, y1, y2), , , bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = Mid(sTest, y1, y2)
                End If
            
            Case 12: '"NPC #"
                If nLimit > 0 Then nCount = nCount + 1
                If nLimit > 0 And nCount > nLimit Then GoTo skip:
                sLocation = "NPC: "
                Set oLI = lv.ListItems.Add()
                If bPercentColumn Then
                    oLI.Text = ""
                    If nPercent > 0 Then oLI.Text = nPercent & "%"
                    oLI.Tag = nPercent
                    oLI.ListSubItems.Add 1, , sLocation & sHeader & GetMonsterName(nValue, bHideRecordNumbers) & sDisplayFooter
                    oLI.ListSubItems(1).Tag = nValue
                ElseIf bTwoColumns Then
                    oLI.Text = sLocation & sHeader
                    oLI.ListSubItems.Add 1, , GetMonsterName(nValue, bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = "monster"
                    oLI.ListSubItems(1).Tag = nValue
                Else
                    oLI.Text = sLocation & sHeader & GetMonsterName(nValue, bHideRecordNumbers) & sPercent & sDisplayFooter
                    oLI.Tag = nValue
                End If
                
        End Select
        
skip:
        x = y1
        GoTo checknext:
    End If
Next z

finish:

If lv.ListItems.Count > 1 And Not bDontSort And Not bPercentColumn Then
    Call SortListView(lv, 1, ldtstring, True)
    lv.Sorted = False
End If

If lv.ListItems.Count > 1 Then
    If Right(sLoc, 2) = "+" & Chr(0) Then
        Set oLI = lv.ListItems.Add(lv.ListItems.Count + 1)
        If bTwoColumns Or bPercentColumn Then
            oLI.ListSubItems.Add 1, , "... plus more."
        Else
            oLI.Text = "... plus more."
        End If
        oLI.Tag = 0
    ElseIf nLimit > 0 And nCount >= nLimit And sHeader = "" And sFooter = "" Then
        Set oLI = lv.ListItems.Add(lv.ListItems.Count + 1)
        If bTwoColumns Or bPercentColumn Then
            oLI.ListSubItems.Add 1, , "... plus " & (nCount - nLimit) & " more. Double-click to see all."
        Else
            oLI.Text = "... plus " & (nCount - nLimit) & " more. Double-click to see all."
        End If
        oLI.Tag = "nolimit"
    End If
End If

Set oLI = Nothing
Exit Sub

error:
HandleError
Set oLI = Nothing

End Sub

'Public Function GetLocations_STR(ByVal sLoc As String) As String
'Dim sLook As String, sChar As String, sTest As String, sSuffix As String
'Dim x As Integer, y1 As Integer, y2 As Integer, z As Integer, nValue As Long, x2 As Integer
'
'If Len(sLoc) < 5 Then
'    GetLocations_STR = "None."
'    Exit Function
'End If
'
'sTest = LCase(sLoc)
'
'For z = 1 To 10
'
'    x = 1
'    Select Case z
'        Case 1: sLook = "room "
'        Case 2: sLook = "monster #"
'        Case 3: sLook = "textblock #"
'        Case 4: sLook = "textblock(rndm) #"
'        Case 5: sLook = "item #"
'        Case 6: sLook = "spell #"
'        Case 7: sLook = "shop #"
'        Case 8: sLook = "shop(sell) #"
'        Case 9: sLook = "shop(nogen) #"
'        Case 10: sLook = "group "
'    End Select
'
'checknext:
'    sSuffix = ""
'    If Not InStr(x, sTest, sLook) = 0 Then
'
'
'        x = InStr(x, sTest, sLook) 'sets x to the position of the string we're looking for
'
'        If z = 10 Then
'            y1 = x + 1
'            GoTo nonumber:
'        End If
'
'        y1 = x + Len(sLook) 'len of string searching (to position y1 at first number)
'        y2 = 0
'nextnumber:
'        sChar = Mid(sTest, y1 + y2, 1)
'        Select Case sChar
'            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "/":
'                If Not y1 + y2 - 1 = Len(sTest) Then
'                    y2 = y2 + 1
'                    GoTo nextnumber:
'                End If
'            Case "+": 'end of string
'                Exit Function
'            Case "(": 'precent
'                x2 = InStr(y1 + y2, sTest, ")")
'                If Not x2 = 0 Then
'                    sSuffix = " " & Mid(sTest, y1 + y2, x2 - y1 - y2 + 1)
'                End If
'            Case Else:
'        End Select
'
'        If y2 = 0 Then
'            'if there were no numbers after the string
'            x = y1
'            GoTo checknext:
'        End If
'
'        If Not z = 1 Then 'not room or group
'            nValue = Val(Mid(sTest, y1, y2))
'        End If
'
'nonumber:
'        If Not GetLocations_STR = "" Then GetLocations_STR = GetLocations_STR & ", "
'
'        Select Case z
'            Case 1: '"room "
'                GetLocations_STR = GetLocations_STR & "Room: " & GetRoomName(Mid(sTest, y1, y2)) & sSuffix
'
'            Case 2: '"monster #"
'                GetLocations_STR = GetLocations_STR & "Monster: " & GetMonsterName(nValue, True) & sSuffix
'
'            Case 3: '"textblock #"
'                GetLocations_STR = GetLocations_STR & "Textblock " & nValue & sSuffix
'
'            Case 4: '"textblock(rndm) #"
'                GetLocations_STR = GetLocations_STR & "Textblock " & nValue & " (random)" & sSuffix
'
'            Case 5: '"item #"
'                GetLocations_STR = GetLocations_STR & "Item: " & GetItemName(nValue) & sSuffix
'
'            Case 6: '"spell #"
'                GetLocations_STR = GetLocations_STR & "Spell: " & GetSpellName(nValue) & sSuffix
'
'            Case 7: '"shop #"
'                GetLocations_STR = GetLocations_STR & "Shop: " & GetShopName(nValue) & sSuffix
'
'            Case 8: '"shop(sell) #"
'                GetLocations_STR = GetLocations_STR & "Shop: " & GetShopName(nValue) & " (sell only)" & sSuffix
'
'            Case 9: '"shop(nogen) #"
'                GetLocations_STR = GetLocations_STR & "Shop: " & GetShopName(nValue) & " (wont regen)" & sSuffix
'
'            Case 10: 'group
'                GetLocations_STR = GetLocations_STR & "Regular room regeneration via group."
'
'        End Select
'
'        x = y1
'        GoTo checknext:
'    End If
'Next z
'
'If GetLocations_STR = "" Then GetLocations_STR = "None."
'
'End Function

Public Sub SelectAll(ByRef TB As TextBox)
Dim tX As Integer
DoEvents
tX = GetAsyncKeyState(VK_LBUTTON)
If tX And 32768 Then 'mousebutton is down
    frmMain.timSelectAll.Enabled = True
    Do While (tX And 32768) And frmMain.timSelectAll.Enabled
        DoEvents
        tX = GetAsyncKeyState(VK_LBUTTON)
    Loop
End If
TB.SelStart = 0
TB.SelLength = Len(TB.Text)
End Sub


Public Function GetEquipCaption(ByVal nIndex As Integer) As String

Select Case nIndex
    Case 0: GetEquipCaption = "Head"
    Case 1: GetEquipCaption = "Ears"
    Case 2: GetEquipCaption = "Neck"
    Case 3: GetEquipCaption = "Back"
    Case 4: GetEquipCaption = "Torso"
    Case 5: GetEquipCaption = "Arms"
    Case 6: GetEquipCaption = "Wrist"
    Case 7: GetEquipCaption = "Wrist"
    Case 8: GetEquipCaption = "Hands"
    Case 9: GetEquipCaption = "Finger"
    Case 10: GetEquipCaption = "Finger"
    Case 11: GetEquipCaption = "Waist"
    Case 12: GetEquipCaption = "Legs"
    Case 13: GetEquipCaption = "Feet"
    Case 14: GetEquipCaption = "Worn"
    Case 15: GetEquipCaption = "Off-Hand"
    Case 16: GetEquipCaption = "Weapon Hand"
    Case 17: GetEquipCaption = "Eyes"
    Case 18: GetEquipCaption = "Face"
    Case 19: GetEquipCaption = "Worn"
End Select

End Function

Public Sub AppReload(ByVal bNewSettings As Boolean)

frmMain.bDontCallTerminate = True
Call AppTerminate
DoEvents

bAppTerminating = False
bAppReallyTerminating = False
If bNewSettings Then Call CreateSettings
DoEvents

sSessionLastCharFile = ""
sSessionLastLoadDir = ""
sSessionLastLoadName = ""
sSessionLastSaveDir = ""
sSessionLastSaveName = ""

Load frmMain
frmMain.bDontCallTerminate = False
End Sub

Public Sub AppTerminate()
On Error Resume Next

bAppTerminating = True
bCancelTerminate = False

If FormIsLoaded("frmMain") Then
    frmMain.timButtonPress.Enabled = False
    frmMain.timRefreshDelay.Enabled = False
    frmMain.timSelectAll.Enabled = False
    frmMain.timWindowMove(0).Enabled = False
    frmMain.timWindowMove(1).Enabled = False
    frmMain.timWindowMove(2).Enabled = False
End If

Call UnloadForms("none")
On Error Resume Next

If bCancelTerminate Then
    bAppTerminating = False
    bAppReallyTerminating = False
    Exit Sub
End If

Call CloseDatabases

ExitApp
End Sub

Public Sub ExitApp(Optional ByVal exitCode As Long = 0)

On Error Resume Next

bAppReallyTerminating = True

' 1) Tell each form to stop timers/background work in its own code (guards should check bAppReallyTerminating).
' 2) Unload all forms, back to front (handles hidden & modeless too).
Dim i As Long

For i = Forms.Count - 1 To 0 Step -1
    Unload Forms(i)
Next i

' Give pending Unload/Terminate handlers a chance to run.
DoEvents

' If anything reloaded itself, try again once.
If Forms.Count > 0 Then
    For i = Forms.Count - 1 To 0 Step -1
        Unload Forms(i)
    Next i
    DoEvents
End If

End Sub

Public Function RegCreateKeyPath(ByVal enmHKEY As hkey, ByVal strKeyPath As String) As Integer
'****************************************************************************
' By Syntax53
' Create a path of keys
' Inputs: HKEY, KeyPath
' Return: Error code, 0=no error
'****************************************************************************
On Error GoTo error:
Dim cReg As clsRegistryRoutines
Dim x As Long, y As Long, KeyArray() As String

Set cReg = New clsRegistryRoutines

cReg.hkey = enmHKEY

x = InStr(1, strKeyPath, "\")
If x = 0 Then
    cReg.KeyRoot = ""
    cReg.Subkey = strKeyPath
    If Not cReg.KeyExists Then Call cReg.CreateKey(strKeyPath)
    GoTo quit:
End If

KeyArray() = Split(strKeyPath, "\", , vbTextCompare)
'ReDim KeyArray(0)
'KeyArray(0) = Mid(strKeyPath, 1, x - 1)
'y = y + 1
'
'Do While InStr(x + 1, strKeyPath, "\") > 0
'    ReDim Preserve KeyArray(0 To y)
'    KeyArray(y) = Mid(strKeyPath, x + 1, InStr(x + 1, strKeyPath, "\") - x - 1)
'    x = InStr(x + 1, strKeyPath, "\")
'    y = y + 1
'Loop
'
'ReDim Preserve KeyArray(0 To y)
'KeyArray(y) = Mid(strKeyPath, x + 1)

For x = 0 To UBound(KeyArray())
    cReg.KeyRoot = ""
    For y = 0 To (x - 1)
        cReg.KeyRoot = cReg.KeyRoot & KeyArray(y)
        If Not y = (x - 1) Then cReg.KeyRoot = cReg.KeyRoot & "\"
    Next
    cReg.Subkey = KeyArray(x)
    If Not cReg.KeyExists Then Call cReg.CreateKey(cReg.Subkey)
Next x

quit:
Set cReg = Nothing
Erase KeyArray()
Exit Function

error:
RegCreateKeyPath = Err.Number
Resume quit:
End Function

Public Function StringOfNumbersToArray(sNumberString As String) As String()
Dim x As Long, sRet() As String
On Error GoTo error:

If InStr(1, sNumberString, ",", vbTextCompare) = 0 Then
    ReDim sRet(0)
    sRet(0) = val(Replace(Replace(sNumberString, "(", "", 1, -1, vbTextCompare), ")", "", 1, -1, vbTextCompare))
Else
    sRet = Split(sNumberString, ",")
    For x = 0 To UBound(sRet())
        sRet(x) = val(Replace(Replace(sRet(x), "(", "", 1, -1, vbTextCompare), ")", "", 1, -1, vbTextCompare))
    Next x
End If

StringOfNumbersToArray = sRet

out:
On Error Resume Next
Exit Function
error:
Call HandleError("NumberStringToArray")
Resume out:
End Function

Public Sub ColorListviewRow(lv As ListView, RowNbr As Long, RowColor As OLE_COLOR, Optional bAndBold As Boolean)

On Error GoTo error:

'***************************************************************************
'Purpose: Color a ListView Row
'Inputs : lv - The ListView
'         RowNbr - The index of the row to be colored
'         RowColor - The color to color it
'Outputs: None
'***************************************************************************
    
Dim itmX As ListItem
Dim lvSI As ListSubItem
Dim intIndex As Integer


Set itmX = lv.ListItems(RowNbr)
itmX.ForeColor = RowColor
If bAndBold Then
    itmX.Bold = True
Else
    itmX.Bold = False
End If
For intIndex = 1 To itmX.ListSubItems.Count
    Set lvSI = itmX.ListSubItems(intIndex)
    lvSI.ForeColor = RowColor
    If bAndBold Then
        lvSI.Bold = True
    Else
        lvSI.Bold = False
    End If
    DoEvents
Next
'lv.ListItems(2).Selected = True
'lv.ListItems(1).Selected = True

Set itmX = Nothing
Set lvSI = Nothing
   

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("ColorListviewRow")
Resume out:
End Sub

Public Function InstrCount(StringToSearch As String, StringToFind As String) As Long
    If Len(StringToFind) > 0 Then
        InstrCount = UBound(Split(StringToSearch, StringToFind))
    End If
End Function

Public Function Get_MegaMUD_ExitsCode(ByVal nMapNum As Long, ByVal nRoomNum As Long) As String
Dim RoomExit As RoomExitType, x As Integer, sRoomName As String
Dim sExits(9) As String, nExitVal(9) As Integer, nExitPosition(9) As Integer
Dim nExitsCalculated(1 To 5) As Integer
Dim bDoor As Boolean
On Error GoTo error:

sExits(0) = "N"
sExits(1) = "S"
sExits(2) = "E"
sExits(3) = "W"
sExits(4) = "NE"
sExits(5) = "NW"
sExits(6) = "SE"
sExits(7) = "SW"
sExits(8) = "U"
sExits(9) = "D"

nExitVal(0) = 1 '"N"
nExitVal(1) = 4 '"S"
nExitVal(2) = 1 '"E"
nExitVal(3) = 4 '"W"
nExitVal(4) = 1 '"NE"
nExitVal(5) = 4 '"NW"
nExitVal(6) = 1 '"SE"
nExitVal(7) = 4 '"SW"
nExitVal(8) = 1 '"U"
nExitVal(9) = 4 '"D"

nExitPosition(0) = 5 '"N"
nExitPosition(1) = 5 '"S"
nExitPosition(2) = 4 '"E"
nExitPosition(3) = 4 '"W"
nExitPosition(4) = 3 '"NE"
nExitPosition(5) = 3 '"NW"
nExitPosition(6) = 2 '"SE"
nExitPosition(7) = 2 '"SW"
nExitPosition(8) = 1 '"U"
nExitPosition(9) = 1 '"D"

tabRooms.Index = "idxRooms"
tabRooms.Seek "=", nMapNum, nRoomNum
If tabRooms.NoMatch Then GoTo out:

sRoomName = tabRooms.Fields("Name")

For x = 0 To 9
    bDoor = False
    If Not val(tabRooms.Fields(sExits(x))) = 0 Then
        RoomExit = ExtractMapRoom(tabRooms.Fields(sExits(x)))
        If RoomExit.Map > 0 And RoomExit.Room > 0 Then
            If Len(RoomExit.ExitType) > 2 Then
                Select Case Left(RoomExit.ExitType, 5)
                    Case "(Key:": bDoor = True
                    Case "(Item":
                    Case "(Toll":
                    Case "(Hidd": GoTo exit_not_seen:
                    Case "(Door": bDoor = True
                    Case "(Trap":
                    Case "(Text": GoTo exit_not_seen:
                    Case "(Gate": bDoor = True
                    Case "Actio": GoTo exit_not_seen:
                    Case "(Clas":
                    Case "(Race":
                    Case "(Leve":
                    Case "(Time":
                    Case "(Tick":
                    Case "(Max ":
                    Case "(Bloc":
                    Case "(Alig":
                    Case "(Dela":
                    Case "(Cast":
                    Case "(Abil":
                    Case "(Spel":
                End Select
            End If
            nExitsCalculated(nExitPosition(x)) = nExitsCalculated(nExitPosition(x)) + (nExitVal(x) * IIf(bDoor, 2, 1))
        End If
    End If
exit_not_seen:
Next x

For x = 1 To 5
    Get_MegaMUD_ExitsCode = Get_MegaMUD_ExitsCode & Hex(nExitsCalculated(x))
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("Get_MegaMUD_ExitsCode")
End Function

    
Public Function Get_MegaMUD_RoomHash(ByVal sRoomName As String, Optional ByVal nMapNum As Long, Optional ByVal nRoomNum As Long) As String
Dim x As Integer, nValue As Long
On Error GoTo error:

Get_MegaMUD_RoomHash = "FFF"

If nMapNum > 0 And nRoomNum > 0 Then
    tabRooms.Index = "idxRooms"
    tabRooms.Seek "=", nMapNum, nRoomNum
    If tabRooms.NoMatch Then GoTo out:
    sRoomName = tabRooms.Fields("Name")
End If

'// Calculate the checksum of the room name
'dwCheckSum = 0;
'for (i = 0; pRec->szName[i]; i++)
'    dwCheckSum += (DWORD) ((int)(pRec->szName[i] * (i+1)));
'dwCheckSum = dwCheckSum << 20;
    
x = 1
Do While x <= Len(sRoomName)
    nValue = nValue + (x * asc(Mid(sRoomName, x, 1)))
    x = x + 1
Loop

Get_MegaMUD_RoomHash = Hex(nValue)
Get_MegaMUD_RoomHash = Right(Get_MegaMUD_RoomHash, 3)
If Len(Get_MegaMUD_RoomHash) < 3 Then Get_MegaMUD_RoomHash = String(3 - Len(Get_MegaMUD_RoomHash), "0") & Get_MegaMUD_RoomHash

out:
On Error Resume Next
Exit Function
error:
Call HandleError("Get_MegaMUD_RoomHash")
Resume out:
End Function

Function in_array_long_md(ByRef SearchArray() As Long, ByVal nFindValue As Long, _
    Optional ByVal nIndexDimension1 As Integer = -32000, Optional ByVal nNotIndexDimension1 As Integer = -32000, _
    Optional ByVal nIndexDimension2 As Integer = -32000, Optional ByVal nNotIndexDimension2 As Integer = -32000) As Boolean
Dim x As Long, y As Long
On Error GoTo error:

For x = LBound(SearchArray(), 1) To UBound(SearchArray(), 1)
    If nIndexDimension1 <> -32000 And x <> nIndexDimension1 Then GoTo skipx:
    If nNotIndexDimension1 <> -32000 And x = nNotIndexDimension1 Then GoTo skipx:
    For y = LBound(SearchArray(), 2) To UBound(SearchArray(), 2)
        If nIndexDimension2 <> -32000 And y <> nIndexDimension2 Then GoTo skipy:
        If nNotIndexDimension2 <> -32000 And y = nNotIndexDimension2 Then GoTo skipy:
        If SearchArray(x, y) = nFindValue Then
            in_array_long_md = True
            Exit Function
        End If
skipy:
    Next y
skipx:
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("in_array_long_md")
Resume out:
End Function

Function in_array_long_3d(ByRef SearchArray() As Long, ByVal nFindValue As Long, _
    Optional ByVal nIndexDimension1 As Integer = -32000, Optional ByVal nNotIndexDimension1 As Integer = -32000, _
    Optional ByVal nIndexDimension2 As Integer = -32000, Optional ByVal nNotIndexDimension2 As Integer = -32000, _
    Optional ByVal nIndexDimension3 As Integer = -32000, Optional ByVal nNotIndexDimension3 As Integer = -32000) As Boolean
Dim x As Long, y As Long, z As Long
On Error GoTo error:

For x = LBound(SearchArray(), 1) To UBound(SearchArray(), 1)
    If nIndexDimension1 <> -32000 And x <> nIndexDimension1 Then GoTo skipx:
    If nNotIndexDimension1 <> -32000 And x = nNotIndexDimension1 Then GoTo skipx:
    For y = LBound(SearchArray(), 2) To UBound(SearchArray(), 2)
        If nIndexDimension2 <> -32000 And y <> nIndexDimension2 Then GoTo skipy:
        If nNotIndexDimension2 <> -32000 And y = nNotIndexDimension2 Then GoTo skipy:
        For z = LBound(SearchArray(), 3) To UBound(SearchArray(), 3)
            If nIndexDimension3 <> -32000 And y <> nIndexDimension3 Then GoTo skipz:
            If nNotIndexDimension3 <> -32000 And y = nNotIndexDimension3 Then GoTo skipz:
            If SearchArray(x, y, z) = nFindValue Then
                in_array_long_3d = True
                Exit Function
            End If
skipz:
        Next z
skipy:
    Next y
skipx:
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("in_array_long_3d")
Resume out:
End Function

Function getval_array_long_3d(ByRef SearchArray() As Long, ByVal nFindValue As Long, _
    ByRef nReturnValueByRef As Long, Optional ByVal nReturnIndex = 1, Optional ByVal nMatchReturnValue As Integer = -32000, _
    Optional ByVal nIndexDimension1 As Integer = -32000, Optional ByVal nNotIndexDimension1 As Integer = -32000, _
    Optional ByVal nIndexDimension2 As Integer = -32000, Optional ByVal nNotIndexDimension2 As Integer = -32000, _
    Optional ByVal nIndexDimension3 As Integer = -32000, Optional ByVal nNotIndexDimension3 As Integer = -32000) As Boolean
Dim x As Long, y As Long, z As Long
On Error GoTo error:
nReturnValueByRef = 0

For x = UBound(SearchArray(), 1) To LBound(SearchArray(), 1) Step -1
    If nIndexDimension1 <> -32000 And x <> nIndexDimension1 Then GoTo skipx:
    If nNotIndexDimension1 <> -32000 And x = nNotIndexDimension1 Then GoTo skipx:
    
    For y = UBound(SearchArray(), 2) To LBound(SearchArray(), 2) Step -1
        If nIndexDimension2 <> -32000 And y <> nIndexDimension2 Then GoTo skipy:
        If nNotIndexDimension2 <> -32000 And y = nNotIndexDimension2 Then GoTo skipy:
        
        For z = UBound(SearchArray(), 3) To LBound(SearchArray(), 3) Step -1
            If nIndexDimension3 <> -32000 And z <> nIndexDimension3 Then GoTo skipz:
            If nNotIndexDimension3 <> -32000 And z = nNotIndexDimension3 Then GoTo skipz:
            
            If SearchArray(x, y, z) = nFindValue Then
                
                If nMatchReturnValue <> -32000 And SearchArray(x, y, nReturnIndex) <> nMatchReturnValue Then GoTo skipz:
                
                nReturnValueByRef = SearchArray(x, y, nReturnIndex)
                getval_array_long_3d = True
                Exit Function
            End If
skipz:
        Next z
skipy:
    Next y
skipx:
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("getval_array_long_3d")
Resume out:
End Function

Function getindex_array_long_3d(ByRef SearchArray() As Long, ByVal nFindValue As Long, _
    ByRef nReturnIndexByRef As Long, ByVal nReturnDimension As Long, _
    Optional ByVal nIndexDimension1 As Integer = -32000, Optional ByVal nNotIndexDimension1 As Integer = -32000, _
    Optional ByVal nIndexDimension2 As Integer = -32000, Optional ByVal nNotIndexDimension2 As Integer = -32000, _
    Optional ByVal nIndexDimension3 As Integer = -32000, Optional ByVal nNotIndexDimension3 As Integer = -32000) As Boolean
Dim x As Long, y As Long, z As Long
On Error GoTo error:
nReturnIndexByRef = 0

For x = LBound(SearchArray(), 1) To UBound(SearchArray(), 1)
    If nIndexDimension1 <> -32000 And x <> nIndexDimension1 Then GoTo skipx:
    If nNotIndexDimension1 <> -32000 And x = nNotIndexDimension1 Then GoTo skipx:
    
    For y = LBound(SearchArray(), 2) To UBound(SearchArray(), 2)
        If nIndexDimension2 <> -32000 And y <> nIndexDimension2 Then GoTo skipy:
        If nNotIndexDimension2 <> -32000 And y = nNotIndexDimension2 Then GoTo skipy:
        
        For z = LBound(SearchArray(), 3) To UBound(SearchArray(), 3)
            If nIndexDimension3 <> -32000 And z <> nIndexDimension3 Then GoTo skipz:
            If nNotIndexDimension3 <> -32000 And z = nNotIndexDimension3 Then GoTo skipz:
            
            If SearchArray(x, y, z) = nFindValue Then
                
                'If nMatchReturnValue <> -32000 And SearchArray(x, y, nReturnIndex) <> nMatchReturnValue Then GoTo skipz:
                
                Select Case nReturnDimension
                    Case 1:
                        nReturnIndexByRef = x
                        getindex_array_long_3d = True
                    Case 2:
                        nReturnIndexByRef = y
                        getindex_array_long_3d = True
                    Case 3:
                        nReturnIndexByRef = z
                        getindex_array_long_3d = True
                End Select
                
                Exit Function
            End If
skipz:
        Next z
skipy:
    Next y
skipx:
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("getindex_array_long_3d")
Resume out:
End Function

Function GetAbilDiffText(nAbilNumber, ByVal nValue1 As Long, ByVal nValue2 As Long, Optional ByVal sCurrentText As String, _
    Optional ByVal nPercent As Integer, Optional ByVal bOppositeMath As Boolean) As String
On Error GoTo error:

Dim nValue As Long
If bOppositeMath Then
    nValue = nValue2 - nValue1
Else
    nValue = nValue1 - nValue2
End If

'59 and 43 currently coded so that nValue1 and nValue2 would always be equal when reaching this function

Select Case nAbilNumber
    'Case 135: 'minlvl
    '    GetAbilDiffText = "Min LVL: " & nValue1
    '    If nValue2 > 0 Then GetAbilDiffText = GetAbilDiffText & " vs " & nValue2
    Case 59: 'class ok
        GetAbilDiffText = GetClassName(nValue1)
        
    Case 43: 'casts spell
        If nValue1 > 0 Then
            GetAbilDiffText = "[" & GetSpellName(nValue1, bHideRecordNumbers) & ", " & PullSpellEQ(True, 0, nValue1)
            If Not nPercent = 0 Then
                GetAbilDiffText = GetAbilDiffText & ", " & nPercent & "%]"
            Else
                GetAbilDiffText = GetAbilDiffText & "]"
            End If
        End If
        'GetAbilDiffText = GetSpellName(nValue1, bHideRecordNumbers)
        
    Case 114: '%spell
                
    Case Else:
        GetAbilDiffText = GetAbilityStats(nAbilNumber, nValue)
        If nAbilNumber = 135 Then GetAbilDiffText = GetAbilDiffText & " (" & IIf(bOppositeMath, nValue2, nValue1) & ")"
End Select

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetAbilDiffText")
Resume out:
End Function

Function ExtractRoomActions(ByVal sExit As String) As String
On Error GoTo error:

ExtractRoomActions = sExit
If InStr(1, ExtractRoomActions, ":", vbTextCompare) > 0 Then
    ExtractRoomActions = Trim(Mid(ExtractRoomActions, InStr(1, ExtractRoomActions, ":", vbTextCompare) + 1, 999))
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("ExtractRoomActions")
Resume out:
End Function

Public Function ClearSavedDamageVsMonster()
On Error GoTo error:
Dim x As Long

For x = 0 To UBound(nCharDamageVsMonster)
    nCharDamageVsMonster(x) = -1
    nCharMinDamageVsMonster(x) = -1
    nCharSurpriseDamageVsMonster(x) = -1
Next x
sCharDamageVsMonsterConfig = sGlobalAttackConfig

out:
On Error Resume Next
Exit Function
error:
Call HandleError("ClearSavedDamageVsMonster")
Resume out:
End Function

Public Function ClearMonsterDamageVsCharALL(Optional bPartyInstead As Boolean = False)
On Error GoTo error:
Dim x As Long

If bPartyInstead Then
    For x = 0 To UBound(nMonsterDamageVsParty)
        nMonsterDamageVsParty(x) = -1
    Next x
    bMonsterDamageVsPartyCalculated = False
    bDontPromptCalcPartyMonsterDamage = False
Else
    For x = 0 To UBound(nMonsterDamageVsChar)
        nMonsterDamageVsChar(x) = -1
    Next x
    sMonsterDamageVsCharDefenseConfig = ""
    bDontPromptCalcCharMonsterDamage = False
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("ClearMonsterDamageVsCharALL")
Resume out:
End Function

Public Function CalculateMonsterDamageVsCharALL(Optional bPartyInstead As Boolean = False)
On Error GoTo error:
Dim nInterval As Integer, nDamage As Currency, bHasAttack As Boolean, x As Integer

frmMain.bMapCancelFind = False
frmMain.Enabled = False
frmMain.timWindowMove(0).Enabled = False

Load frmProgressBar
Call frmProgressBar.SetRange(tabMonsters.RecordCount / 5)
frmProgressBar.ProgressBar.Value = 1
If bPartyInstead Then
    frmProgressBar.lblCaption.Caption = "Calculate mob dmg vs party..."
Else
    frmProgressBar.lblCaption.Caption = "Calculate mob dmg vs char..."
End If
Set frmProgressBar.objFormOwner = frmMain

DoEvents
frmProgressBar.Show vbModeless, frmMain
DoEvents
'Call LockWindowUpdate(frmMain.hWnd)
nInterval = 1

tabMonsters.MoveFirst
Do While tabMonsters.EOF = False
    
    bHasAttack = False
    For x = 0 To 4
        If bHasAttack = True Then Exit For
        If tabMonsters.Fields("AttType-" & x) > 0 And tabMonsters.Fields("AttType-" & x) < 4 Then bHasAttack = True
    Next x
    For x = 0 To 4
        If bHasAttack = True Then Exit For
        If tabMonsters.Fields("MidSpell-" & x) > 0 Then bHasAttack = True
    Next x
    
    If bHasAttack Then
        nDamage = CalculateMonsterDamageVsChar(tabMonsters.Fields("Number"), bPartyInstead)
    Else
        If bPartyInstead Then
            nMonsterDamageVsParty(tabMonsters.Fields("Number")) = 0
        Else
            nMonsterDamageVsChar(tabMonsters.Fields("Number")) = 0
        End If
    End If
    
skip:
    If nInterval > 5 Then
        Call frmProgressBar.IncreaseProgress
        nInterval = 1
    ElseIf nInterval > 0 Then
        nInterval = nInterval + 1
    End If
    
    DoEvents
    If frmMain.bMapCancelFind Then Exit Do
    
    tabMonsters.MoveNext
Loop
tabMonsters.MoveFirst

If bPartyInstead Then
    bMonsterDamageVsPartyCalculated = True
    bDontPromptCalcPartyMonsterDamage = True
Else
    Call SetCharDefenseDescription
    sMonsterDamageVsCharDefenseConfig = sGlobalCharDefenseDescription
    bDontPromptCalcCharMonsterDamage = True
End If

out:
On Error Resume Next
frmMain.Enabled = True
Call EnsureAppForeground(frmMain)
frmProgressBar.Hide
DoEvents
If FormIsLoaded("frmProgressBar") Then Unload frmProgressBar
frmMain.timWindowMove(0).Enabled = True
Exit Function
error:
Call HandleError("CalculateMonsterDamageVsCharALL")
Resume out:
End Function

Public Function CalculateMonsterDamageVsChar(ByVal nMonsterNumber As Long, Optional bPartyInstead As Boolean = False) As Currency
On Error GoTo error:
Dim nNon As Integer, nAnti As Integer

If nMonsterNumber <= 0 Then Exit Function
If nGlobalMonsterSimRounds < 100 Then nGlobalMonsterSimRounds = 100
If nGlobalMonsterSimRounds > 10000 Then nGlobalMonsterSimRounds = 10000

If val(frmMain.txtMonsterLairFilter(0).Text) < 2 Or val(frmMain.txtMonsterLairFilter(6).Text) < 1 _
    Or val(frmMain.txtMonsterLairFilter(0).Text) = val(frmMain.txtMonsterLairFilter(6).Text) _
    Or bPartyInstead = False Then
    
    nAnti = 0: If val(frmMain.txtMonsterLairFilter(0).Text) = val(frmMain.txtMonsterLairFilter(6).Text) Then nAnti = 1
    
    Call SetupMonsterAttackSimWithCharStats(nGlobalMonsterSimRounds, False, bPartyInstead, nAnti)
    Call PopulateMonsterDataToAttackSim(nMonsterNumber, clsMonAtkSim)
    If clsMonAtkSim.nNumberOfRounds > 0 Then clsMonAtkSim.RunSim
    CalculateMonsterDamageVsChar = clsMonAtkSim.nAverageDamage
Else 'party
    nAnti = val(frmMain.txtMonsterLairFilter(6).Text)
    nNon = val(frmMain.txtMonsterLairFilter(0).Text) - nAnti
    If nAnti > 0 Then
        Call SetupMonsterAttackSimWithCharStats(nGlobalMonsterSimRounds, False, bPartyInstead, 1)
        Call PopulateMonsterDataToAttackSim(nMonsterNumber, clsMonAtkSim)
        If clsMonAtkSim.nNumberOfRounds > 0 Then clsMonAtkSim.RunSim
        CalculateMonsterDamageVsChar = clsMonAtkSim.nAverageDamage * nAnti
    End If
    If nNon > 0 Then
        Call SetupMonsterAttackSimWithCharStats(nGlobalMonsterSimRounds, False, bPartyInstead, 0)
        Call PopulateMonsterDataToAttackSim(nMonsterNumber, clsMonAtkSim)
        If clsMonAtkSim.nNumberOfRounds > 0 Then clsMonAtkSim.RunSim
        CalculateMonsterDamageVsChar = CalculateMonsterDamageVsChar + (clsMonAtkSim.nAverageDamage * nNon)
    End If
    CalculateMonsterDamageVsChar = CalculateMonsterDamageVsChar / (nAnti + nNon)
End If


If bPartyInstead Then
    nMonsterDamageVsParty(nMonsterNumber) = Round(CalculateMonsterDamageVsChar, 1)
Else
    nMonsterDamageVsChar(nMonsterNumber) = Round(CalculateMonsterDamageVsChar, 1)
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalculateMonsterDamageVsChar")
Resume out:
End Function

Public Function GetPreCalculatedMonsterDamage(ByVal nMonsterNumber As Long, ByRef sReturn As String, Optional ByVal nParty As Integer) As Double
On Error GoTo error:
Dim bUseCharacter As Boolean

If frmMain.chkGlobalFilter.Value = 1 Then bUseCharacter = True
If nParty < 1 Then
    If frmMain.optMonsterFilter(1).Value = True Then nParty = val(frmMain.txtMonsterLairFilter(0).Text)
End If
If nParty < 1 Then nParty = 1

If nMonsterNumber < 1 Then
    If nParty > 1 Then
        sReturn = "vs Party"
    ElseIf nParty = 1 And bUseCharacter Then
        sReturn = "vs Char"
    Else
        sReturn = "(default)"
    End If
    Exit Function
End If

If nParty > 1 And nMonsterDamageVsParty(nMonsterNumber) >= 0 Then
    GetPreCalculatedMonsterDamage = nMonsterDamageVsParty(nMonsterNumber)
    sReturn = "vs Party"
ElseIf nParty = 1 And bUseCharacter And nMonsterDamageVsChar(nMonsterNumber) >= 0 Then
    GetPreCalculatedMonsterDamage = nMonsterDamageVsChar(nMonsterNumber)
    sReturn = "vs Char"
ElseIf nMonsterDamageVsDefault(nMonsterNumber) >= 0 Then
    GetPreCalculatedMonsterDamage = nMonsterDamageVsDefault(nMonsterNumber)
    sReturn = "(default)"
Else
    GetPreCalculatedMonsterDamage = GetMonsterAvgDmgFromDB(nMonsterNumber)
    sReturn = "(default)"
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetPreCalculatedMonsterDamage")
Resume out:
End Function

Public Sub SetupMonsterAttackSimWithCharStats(Optional ByVal nRounds As Integer = 50, Optional ByVal bDynamic As Boolean = True, Optional ByVal bPartyInstead As Boolean = False, Optional ByVal nPartyAntiMagic As Integer = 0)
On Error GoTo error:

Call clsMonAtkSim.ResetValues
clsMonAtkSim.bUseCPU = False
clsMonAtkSim.nCombatLogMaxRounds = 100

clsMonAtkSim.bCombatLogMaxRoundOnly = True

clsMonAtkSim.nNumberOfRounds = nRounds

If bDynamic Then
    clsMonAtkSim.bDynamicCalc = 1
Else
    clsMonAtkSim.bDynamicCalc = 0
End If
clsMonAtkSim.bGreaterMUD = bGreaterMUD
clsMonAtkSim.nDynamicCalcDifference = 0.001
clsMonAtkSim.nUserMR = 50

If bPartyInstead Then
    If val(frmMain.txtMonsterLairFilter(1).Text) > 0 Then clsMonAtkSim.nUserAC = val(frmMain.txtMonsterLairFilter(1).Text)
    If val(frmMain.txtMonsterLairFilter(2).Text) > 0 Then clsMonAtkSim.nUserDR = val(frmMain.txtMonsterLairFilter(2).Text)
    If val(frmMain.txtMonsterLairFilter(3).Text) > 0 Then clsMonAtkSim.nUserMR = val(frmMain.txtMonsterLairFilter(3).Text)
    If val(frmMain.txtMonsterLairFilter(4).Text) > 0 Then clsMonAtkSim.nUserDodge = val(frmMain.txtMonsterLairFilter(4).Text)
    If nPartyAntiMagic = 1 Then clsMonAtkSim.nUserAntiMagic = 1
Else
    If val(frmMain.lblInvenCharStat(2).Caption) > 0 Then clsMonAtkSim.nUserAC = val(frmMain.lblInvenCharStat(2).Caption)
    If val(frmMain.lblInvenCharStat(3).Caption) > 0 Then clsMonAtkSim.nUserDR = val(frmMain.lblInvenCharStat(3).Caption)
    If val(frmMain.txtCharMR.Text) > 0 Then clsMonAtkSim.nUserMR = val(frmMain.txtCharMR.Text)
    If val(frmMain.lblCharDodge.Tag) > 0 Then clsMonAtkSim.nUserDodge = val(frmMain.lblCharDodge.Tag)
    If frmMain.chkCharAntiMagic.Value = 1 Then clsMonAtkSim.nUserAntiMagic = 1
    clsMonAtkSim.nUserRCOL = frmMain.lblInvenCharStat(28).Tag 'col
    clsMonAtkSim.nUserRFIR = frmMain.lblInvenCharStat(27).Tag 'fir
    clsMonAtkSim.nUserRSTO = frmMain.lblInvenCharStat(25).Tag 'sto
    clsMonAtkSim.nUserRLIT = frmMain.lblInvenCharStat(29).Tag 'lit
    clsMonAtkSim.nUserRWAT = frmMain.lblInvenCharStat(26).Tag 'wat
End If

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("PopulateAttackSimCharStats")
Resume out:
End Sub

Public Function CountListviewSelections(ByRef oLV As ListView) As Long
Dim x As Long
On Error GoTo error:

CountListviewSelections = 0

If oLV.ListItems.Count < 1 Then Exit Function
For x = 1 To oLV.ListItems.Count - 1
    If oLV.ListItems(x).Selected = True Then CountListviewSelections = CountListviewSelections + 1
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CountListviewSelections")
Resume out:
End Function

Public Function ControlExists(ByRef oForm As Form, ByVal sName As String, Optional ByVal nIndex As Integer = -1) As Boolean
On Error GoTo error:
Dim ctl As Control
Dim blnExist As Boolean

blnExist = False

For Each ctl In oForm.Controls
    If ctl.name = sName And TypeName(oForm.Controls(sName)) = "Object" Then
        If nIndex >= 0 Then
            If ctl.Index = nIndex Then
                blnExist = True
                Exit For
            End If
        Else
            blnExist = True
            Exit For
        End If
    End If
Next ctl

ControlExists = blnExist

out:
On Error Resume Next
Exit Function
error:
Call HandleError("ControlExists")
Resume out:
End Function


Public Function GetAbilityStatSlot(ByVal nAbility As Integer, ByVal nAbilityValue As Integer) As tAbilityToStatSlot

If nAbilityValue > 0 Then GetAbilityStatSlot.sText = GetAbilityStats(nAbility, nAbilityValue)

GetAbilityStatSlot.nEquip = -1
Select Case nAbility
    Case 0: 'nothing
    Case 2: '2=AC
        GetAbilityStatSlot.nEquip = 2
        'GetAbilityStatSlot.sText = "AC: "
    Case 3: '3=res_cold
        GetAbilityStatSlot.nEquip = 28
        'GetAbilityStatSlot.sText = "Cold Res: "
    Case 4: '4=max dmg
        GetAbilityStatSlot.nEquip = 11
        'GetAbilityStatSlot.sText = "Max Dmg: "
    Case 5: '5=res_fire
        GetAbilityStatSlot.nEquip = 27
        'GetAbilityStatSlot.sText = "Fire Res: "
    Case 7: '7=DR
        GetAbilityStatSlot.nEquip = 3
        'GetAbilityStatSlot.sText = "DR: "
'    Case 9: 'shadow
'        If bGreaterMUD Then GetAbilityStatSlot.nEquip = 2
    Case 10: '10=AC(BLUR)
        GetAbilityStatSlot.nEquip = 2
        'GetAbilityStatSlot.sText = "AC: "
    Case 13: '13=illu
        GetAbilityStatSlot.nEquip = 23
        'GetAbilityStatSlot.sText = "Illu: "
    Case 14: '14=roomillu
        GetAbilityStatSlot.nEquip = 23
        'GetAbilityStatSlot.sText = "RoomIllu: "
    '21=immu poison
    Case 22: '22=acc
        GetAbilityStatSlot.nEquip = 10
        'GetAbilityStatSlot.sText = "Accy: "
    Case 24: '24=prev prot evil
        GetAbilityStatSlot.nEquip = 20
    Case 25: '25=prgd prot good
        GetAbilityStatSlot.nEquip = 32
    Case 27: '27=stealth
        GetAbilityStatSlot.nEquip = 19
        'GetAbilityStatSlot.sText = "Stealth: "
    Case 29: GetAbilityStatSlot.nEquip = 37 'punch skill
    Case 30: GetAbilityStatSlot.nEquip = 38 'kick skill
    Case 34: '34=dodge
        GetAbilityStatSlot.nEquip = 8
        'GetAbilityStatSlot.sText = "Dodge: "
    Case 35: GetAbilityStatSlot.nEquip = 39 'jk skill
    Case 36: '36=MR
        GetAbilityStatSlot.nEquip = 24
        'GetAbilityStatSlot.sText = "MR: "
    Case 37: '37=picklocks
        GetAbilityStatSlot.nEquip = 22
        'GetAbilityStatSlot.sText = "Picks: "
    Case 38: '38=tracking
        'GetAbilityStatSlot.nEquip = 23
        ''GetAbilityStatSlot.sText = "Tracking: "
    Case 39: '39=thievery
        'GetAbilityStatSlot.nEquip = 20
        'GetAbilityStatSlot.sText = "Thievery: "
    Case 40: '40=findtraps
        GetAbilityStatSlot.nEquip = 21
        'GetAbilityStatSlot.sText = "Traps: "
    '41=disarmtraps
    Case 44: '44=int
        GetAbilityStatSlot.nEquip = 104
    Case 45: '45=wis
        GetAbilityStatSlot.nEquip = 124
    Case 46: '46=str
        GetAbilityStatSlot.nEquip = 101
    Case 47: '47=hea
        GetAbilityStatSlot.nEquip = 123
    Case 48: '48=agi
        GetAbilityStatSlot.nEquip = 102
    Case 49: '49=chm
        GetAbilityStatSlot.nEquip = 103
    Case 58: '58=crits
        GetAbilityStatSlot.nEquip = 7
        'GetAbilityStatSlot.sText = "Crits: "
    Case 65: '65=res_stone
        GetAbilityStatSlot.nEquip = 25
        'GetAbilityStatSlot.sText = "Stone Res: "
    Case 66: '66=res_lit
        GetAbilityStatSlot.nEquip = 29
        'GetAbilityStatSlot.sText = "Light Res: "
    Case 67: GetAbilityStatSlot.nEquip = 31 'quickness
    Case 69: '69=max mana
        GetAbilityStatSlot.nEquip = 6
        'GetAbilityStatSlot.sText = "Mana: "
    Case 70: '70=SC
        GetAbilityStatSlot.nEquip = 9
        'GetAbilityStatSlot.sText = "SC: "
    Case 72: '72=damageshield
        GetAbilityStatSlot.nEquip = 12
        'GetAbilityStatSlot.sText = "Shock: "
    Case 77: '77=percep
        GetAbilityStatSlot.nEquip = 18
        'GetAbilityStatSlot.sText = "Percep: "
    '87=speed
    Case 88: '88=alter hp
        GetAbilityStatSlot.nEquip = 5
        'GetAbilityStatSlot.sText = "HP: "
    
    Case 89: GetAbilityStatSlot.nEquip = 40 'punch accy
    Case 90: GetAbilityStatSlot.nEquip = 41 'kick accy
    Case 91: GetAbilityStatSlot.nEquip = 42 'jumpkick accy
    
    Case 92: GetAbilityStatSlot.nEquip = 34 'punch dmg
    Case 93: GetAbilityStatSlot.nEquip = 35 'kick dmg
    Case 94: GetAbilityStatSlot.nEquip = 36 'jumpkick dmg
    
    Case 96: '96=encum
        GetAbilityStatSlot.nEquip = 4
        'GetAbilityStatSlot.sText = "Enc%: "
    Case 105: '105=acc
        GetAbilityStatSlot.nEquip = 10
        'GetAbilityStatSlot.sText = "Accy: "
    Case 106: '106=acc
        GetAbilityStatSlot.nEquip = 10
        'GetAbilityStatSlot.sText = "Accy: "
    Case 116: '116=bsaccu
        GetAbilityStatSlot.nEquip = 13
        'GetAbilityStatSlot.sText = "BS Accy: "
    Case 117: '117=bsmin
        GetAbilityStatSlot.nEquip = 14
        'GetAbilityStatSlot.sText = "BS Min: "
    Case 118: '118=bsmax
        GetAbilityStatSlot.nEquip = 15
        'GetAbilityStatSlot.sText = "BS Max: "
    Case 123: '123=hpregen
        GetAbilityStatSlot.nEquip = 16
        'GetAbilityStatSlot.sText = "HP Rgn: "
    Case 142: '142=hitmagic
        'GetAbilityStatSlot.nEquip = 31
        ''GetAbilityStatSlot.sText = "Hit Magic: "
    Case 145: '145=manaregen
        GetAbilityStatSlot.nEquip = 17
        'GetAbilityStatSlot.sText = "Mana Rgn: "
    Case 147: '147=res_water
        GetAbilityStatSlot.nEquip = 26
        'GetAbilityStatSlot.sText = "Water Res: "
    Case 165: GetAbilityStatSlot.nEquip = 33 'alter spell dmg
    Case 179: '179=find trap value
        GetAbilityStatSlot.nEquip = 21
        'GetAbilityStatSlot.sText = "Traps: "
    Case 180: '180=pick value
        GetAbilityStatSlot.nEquip = 22
        'GetAbilityStatSlot.sText = "Picks: "
    
End Select
End Function

Public Function GetCurrentAttackName(Optional ByVal bForceAttackDesc As Boolean) As String
On Error GoTo error:

Select Case nGlobalAttackTypeMME
    Case 1, 6, 7: 'eq'd weapon, bash, smash
        If nEquippedItem(16) > 0 Then
            If nGlobalAttackTypeMME = a6_PhysBash Then
                GetCurrentAttackName = "bash"
            ElseIf nGlobalAttackTypeMME = a7_PhysSmash Then
                GetCurrentAttackName = "smash"
            Else
                If bForceAttackDesc Then
                    GetCurrentAttackName = "normal attack"
                Else
                    GetCurrentAttackName = "weapon"
                End If
            End If
            If bForceAttackDesc Then GetCurrentAttackName = GetCurrentAttackName & " w/" & GetItemName(nEquippedItem(16), bHideRecordNumbers)
        Else
            GetCurrentAttackName = "no wepn!"
        End If
    Case 2: 'spell learned
        If bForceAttackDesc Then GetCurrentAttackName = GetSpellName(nGlobalAttackSpellNum, bHideRecordNumbers) & " / "
        GetCurrentAttackName = GetCurrentAttackName & GetSpellShort(nGlobalAttackSpellNum)
        If bForceAttackDesc Then GetCurrentAttackName = GetCurrentAttackName & " @ LVL"
    
    Case 3: 'spell any
        If bForceAttackDesc Then GetCurrentAttackName = GetSpellName(nGlobalAttackSpellNum, bHideRecordNumbers) & " / "
        GetCurrentAttackName = GetCurrentAttackName & GetSpellShort(nGlobalAttackSpellNum) & "@" & nGlobalAttackSpellLVL
        
    Case 4: 'martial arts attack
        '1-Punch, 2-Kick, 3-JumpKick
        Select Case nGlobalAttackMA
            Case 2: 'kick
                GetCurrentAttackName = "kick"
            Case 3: 'jumpkick
                GetCurrentAttackName = "jumpkick"
            Case Else: 'punch
                GetCurrentAttackName = "punch"
        End Select
        
    Case 5: 'manual
        If nGlobalAttackManualP + nGlobalAttackManualM = 0 Then
            GetCurrentAttackName = "zero"
        ElseIf nGlobalAttackManualP > 0 And nGlobalAttackManualM <= 0 Then
            GetCurrentAttackName = nGlobalAttackManualP & " phys"
        ElseIf nGlobalAttackManualP <= 0 And nGlobalAttackManualM > 0 Then
            GetCurrentAttackName = nGlobalAttackManualM & " mag"
        Else
            GetCurrentAttackName = nGlobalAttackManualP & "/" & nGlobalAttackManualM & " dmg"
        End If
        
    Case Else:
        GetCurrentAttackName = "one-shot"
        
End Select

If nGlobalAttackTypeMME > a0_oneshot And bGlobalAttackBackstab Then
    If bForceAttackDesc And nGlobalAttackBackstabWeapon > 0 Then
        GetCurrentAttackName = AutoAppend(GetCurrentAttackName, "+bs", vbCrLf)
    Else
        GetCurrentAttackName = GetCurrentAttackName & "+bs"
    End If
    
    If nGlobalAttackBackstabWeapon > 0 Then
        If bForceAttackDesc Then
            GetCurrentAttackName = GetCurrentAttackName & " w/" & GetItemName(nGlobalAttackBackstabWeapon, bHideRecordNumbers)
        Else
            GetCurrentAttackName = GetCurrentAttackName & "*"
        End If
    End If
End If

If Not bForceAttackDesc And frmMain.optMonsterFilter(1).Value = True And nNMRVer >= 1.83 And eGlobalExpHrModel <> basic_dmg Then
    Call RefreshCombatHealingValues
    Select Case nGlobalAttackHealType
        Case 0: 'infinite
            GetCurrentAttackName = AutoAppend(GetCurrentAttackName, "invincible", " / ")
        Case 1: 'base
            GetCurrentAttackName = AutoAppend(GetCurrentAttackName, "passive", " / ")
        Case 2, 3: 'spell
            If nGlobalAttackHealSpellNum > 0 Then
                GetCurrentAttackName = AutoAppend(GetCurrentAttackName, GetSpellShort(nGlobalAttackHealSpellNum), " / ")
                If nGlobalAttackHealType = 3 Then
                    GetCurrentAttackName = GetCurrentAttackName & "@" & nGlobalAttackHealSpellLVL
                    GetCurrentAttackName = RemoveCharacter(GetCurrentAttackName, " ")
                End If
                'If nGlobalAttackHealRounds > 1 And nGlobalAttackTypeMME <> a3_SpellAny And nGlobalAttackHealType <> 3 Then GetCurrentAttackName = GetCurrentAttackName & "/" & nGlobalAttackHealRounds & "r"
                If nGlobalAttackHealValue > 0 Then GetCurrentAttackName = GetCurrentAttackName & "(" & nGlobalAttackHealValue & ")"
            End If
        Case 4: 'manual
            GetCurrentAttackName = AutoAppend(GetCurrentAttackName, nGlobalAttackHealValue & " heals", " / ")
    End Select
    
    If bGlobalAttackUseMeditate And (nGlobalAttackTypeMME = a2_Spell Or nGlobalAttackTypeMME = a3_SpellAny Or nGlobalAttackHealType = 2 Or nGlobalAttackHealType = 3) Then
        GetCurrentAttackName = AutoAppend(GetCurrentAttackName, "+m", "")
        If nGlobalAttackTypeMME = a3_SpellAny Or nGlobalAttackHealType = 3 Then
            GetCurrentAttackName = RemoveCharacter(GetCurrentAttackName, " ")
        End If
    End If
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetCurrentAttackName")
Resume out:
End Function
