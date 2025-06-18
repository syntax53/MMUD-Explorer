Attribute VB_Name = "modMain"
Option Explicit
Option Base 0

Global bDPIAwareMode As Boolean
Global bDEVELOPMENT_MODE As Boolean
Global bHideRecordNumbers As Boolean
Global bOnlyInGame As Boolean
Global nMonsterSimRounds As Long
Global nDmgScaleFactor As Double
Global nMonsterLairRatioMultiplier As Double
Global bStartup As Boolean
Global bDontSyncSplitters As Boolean
Global nNMRVer As Double
Global nOSversion As cnWin32Ver
Global sCurrentDatabaseFile As String
Global sForceCharacterFile As String
'Global bOnlyLearnable As Boolean
Global nTheoreticalAvgMaxLairsPerRegenPeriod As Integer 'max 1-mob lairs you can clear in 3 minutes (average lair regen of 3 minutes divided by average round of 5 seconds = 36 lairs)
Global bMonsterDamageVsCharCalculated As Boolean
Global bMonsterDamageVsPartyCalculated As Boolean
Global bDontPromptCalcCharMonsterDamage As Boolean
Global bDontPromptCalcPartyMonsterDamage As Boolean
Global nLastItemSortCol As Integer
Public tLastAvgLairInfo As LairInfoType
    
Global nCurrentCharAccyWornItems As Long
Global nCurrentCharAccyAbil22 As Long
Global nCurrentCharQnDbonus As Long
Global nCurrentCharWeaponNumber(1) As Long '0=weapon, 1=offhand
Global nCurrentCharWeaponAccy(1) As Long
Global nCurrentCharWeaponCrit(1) As Long
Global nCurrentCharWeaponMaxDmg(1) As Long
Global nCurrentCharWeaponBSaccy(1) As Long
Global nCurrentCharWeaponBSmindmg(1) As Long
Global nCurrentCharWeaponBSmaxdmg(1) As Long
Global nCurrentCharWeaponPunchSkill(1) As Long
Global nCurrentCharWeaponPunchAccy(1) As Long
Global nCurrentCharWeaponPunchDmg(1) As Long
Global nCurrentCharWeaponKickSkill(1) As Long
Global nCurrentCharWeaponKickAccy(1) As Long
Global nCurrentCharWeaponKickDmg(1) As Long
Global nCurrentCharWeaponJkSkill(1) As Long
Global nCurrentCharWeaponJkAccy(1) As Long
Global nCurrentCharWeaponJkDmg(1) As Long
Global nCurrentCharWeaponStealth(1) As Long

Global nCurrentAttackType As Integer '0-none, 1-weapon, 2/3-spell, 4-MA, 5-manual
Global nCurrentAttackMA As Integer
Global nCurrentAttackSpellNum As Long
Global nCurrentAttackSpellLVL As Integer
Global nCurrentAttackManual As Long
Global sCurrentAttackConfig As String

Public Type TypeGetEquip
    nEquip As Integer
    sText As String
End Type

Public Type tSpellCastValues
    nMinCast As Long
    nMaxCast As Long
    nAvgCast As Long
    nNumCasts As Double
    nCastChance As Integer
    nAvgRoundDmg As Long
    nAvgRoundHeals As Long
    nDuration As Integer
    nDamageResisted As Long
    nFullResistChance As Integer
    bDoesHeal As Boolean
    bDoesDamage As Boolean
    sAvgRound As String
    sLVLincreases As String
    sMMA As String
    sSpellName As String
End Type

Public Type tAttackDamage
    nMinDmg As Long
    nMaxDmg As Long
    nHitChance As Integer
    nDodgeChance As Integer
    nCritChance As Integer
    nQnDBonus As Integer
    nAccy As Integer
    nAvgHit As Long
    nAvgCrit As Long
    nAvgExtraHit As Long
    nAvgExtraSwing As Long
    nSwings As Double
    nRoundPhysical As Long
    nRoundTotal As Long
    sAttackDesc As String
    sAttackDetail As String
End Type

Public Enum QBColorCode
    Black = 0
    Blue = 1
    Green = 2
    Cyan = 3
    Red = 4
    Magenta = 5
    Yellow = 6
    White = 7
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

Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_GETITEMHEIGHT = &H154
Private Const CB_SETDROPPEDWIDTH = &H160
Private Const CB_GETDROPPEDWIDTH = &H15F
Private Const CB_SETDROPPEDCONTROLRECT = &H160
Private Const DT_CALCRECT = &H400
Public Const WM_SETREDRAW As Long = 11

Public bUse_dwmapi As Boolean
Public bPromptSave As Boolean
Public bCancelTerminate As Boolean
Public bAppTerminating As Boolean
Public sRecentFiles(1 To 5, 1 To 2) As String '1=shown, 2=filename
Public sRecentDBs(1 To 5, 1 To 2) As String '1=shown, 2=filename
Public nEquippedItem(0 To 19) As Long
Public nLearnedSpells(0 To 99) As Long
Public nLearnedSpellClass As Integer
Public bLegit As Boolean
Public bGreaterMUD As Boolean
Public bDisableKaiAutolearn As Boolean
Public sSessionLastCharFile As String
Public sSessionLastLoadDir As String
Public sSessionLastLoadName As String
Public sSessionLastSaveDir As String
Public sSessionLastSaveName As String
Public sPartyPasteHeals As String

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
'
'Public Const MONITOR_DEFAULTTONEAREST = &H2
'
'Public Type MONITORINFO
'    cbSize As Long
'    rcMonitor As RECT
'    rcWork As RECT
'    dwFlags As Long
'End Type

Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

'Public Declare Function MoveWindow Lib "user32" _
'  (ByVal hwnd As Long, _
'   ByVal x As Long, ByVal y As Long, _
'   ByVal nWidth As Long, _
'   ByVal nHeight As Long, _
'   ByVal bRepaint As Long) As Long

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
    
'Public Declare Function CalcExpNeeded Lib "lltmmudxp" (ByVal Level As Long, ByVal Chart As Long) As Currency
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Const VK_LBUTTON = &H1
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
'===
'Private Type InitCommonControlsExStruct
'    lngSize As Long
'    lngICC As Long
'End Type
'Private Declare Function InitCommonControls Lib "comctl32" () As Long
'Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As InitCommonControlsExStruct) As Boolean
'
Private Sub Main()
''''
''''    Dim iccex As InitCommonControlsExStruct, hMod As Long
''''    Const ICC_ALL_CLASSES As Long = &HFDFF& ' combination of all known values
''''    ' constant descriptions: http://msdn.microsoft.com/en-us/library/bb775507%28VS.85%29.aspx
''''
''''    With iccex
''''       .lngSize = LenB(iccex)
''''       .lngICC = ICC_ALL_CLASSES    ' you really should customize this value from the available constants
''''    End With
''''    On Error Resume Next ' error? Requires IEv3 or above
''''    hMod = LoadLibrary("shell32.dll")
''''    InitCommonControlsEx iccex
''''    If Err Then
''''        InitCommonControls ' try Win9x version
''''        Err.clear
''''    End If
''''    On Error GoTo 0
''''    '... show your main form next (i.e., Form1.Show)
nOSversion = GetWin32Ver
Load frmMain
''''    If hMod Then FreeLibrary hMod
''''
''''
'''''** Tip 1: Avoid using VB Frames when applying XP/Vista themes
'''''          In place of VB Frames, use pictureboxes instead.
'''''** Tip 2: Avoid using Graphical Style property of buttons, checkboxes and option buttons
'''''          Doing so will prevent them from being themed.
''''
End Sub

Public Sub SetCurrentAttackTypeConfig()
On Error GoTo error:
Dim sConfig As String

sConfig = CStr(nCurrentAttackType)

Select Case nCurrentAttackType
    Case 1, 6, 7, 4: 'weap, bash, smash, MA
        sConfig = sConfig & "_" & nCurrentCharWeaponNumber(0)
        sConfig = sConfig & "_" & nCurrentCharWeaponNumber(1)
        If frmMain.chkGlobalFilter.Value = 1 Then
            sConfig = sConfig & "_" & Val(frmMain.txtGlobalLevel(0).Text) 'lvl
            sConfig = sConfig & frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex) 'class
            sConfig = sConfig & CalcEncumbrancePercent(Val(frmMain.lblInvenCharStat(0).Caption), Val(frmMain.lblInvenCharStat(1).Caption)) 'encum
            sConfig = sConfig & Val(frmMain.txtCharStats(0).Text) 'str
            sConfig = sConfig & Val(frmMain.txtCharStats(3).Text) 'agi
            sConfig = sConfig & Val(frmMain.lblInvenCharStat(7).Tag) - nCurrentCharQnDbonus 'nCritChance
            sConfig = sConfig & Val(frmMain.lblInvenCharStat(11).Tag) 'nPlusMaxDamage
            sConfig = sConfig & Val(frmMain.lblInvenCharStat(30).Tag) 'nPlusMinDamage
            sConfig = sConfig & Val(frmMain.lblInvenCharStat(10).Tag) 'nAttackAccuracy
            sConfig = sConfig & Val(frmMain.lblInvenCharStat(13).Tag) 'nPlusBSaccy
            sConfig = sConfig & Val(frmMain.lblInvenCharStat(14).Tag) 'nPlusBSmindmg
            sConfig = sConfig & Val(frmMain.lblInvenCharStat(15).Tag) 'nPlusBSmaxdmg
            sConfig = sConfig & Val(frmMain.lblInvenCharStat(19).Tag) 'nStealth
            If Val(frmMain.lblInvenStats(19).Tag) >= 2 Then sConfig = sConfig & "1" 'bClassStealth
        Else
            sConfig = sConfig & "_default"
        End If
        
    Case 2, 3: 'spell
        sConfig = sConfig & "_" & nCurrentAttackSpellNum
        If frmMain.chkGlobalFilter.Value = 1 Then
            sConfig = sConfig & "_" & Val(frmMain.txtGlobalLevel(0).Text)
        Else
            sConfig = sConfig & "_" & nCurrentAttackSpellLVL
        End If
        
    Case 5: 'manual
        sConfig = sConfig & "_" & CStr(nCurrentAttackManual)
End Select

If nCurrentAttackType = 4 Then 'MA
    sConfig = sConfig & "_" & CStr(nCurrentAttackMA)
    Select Case nCurrentAttackMA
        Case 1: 'punch
            sConfig = sConfig & "_" & CStr(nCurrentCharWeaponPunchSkill(0) + nCurrentCharWeaponPunchSkill(1))
            sConfig = sConfig & "_" & CStr(nCurrentCharWeaponPunchAccy(0) + nCurrentCharWeaponPunchAccy(1))
            sConfig = sConfig & "_" & CStr(nCurrentCharWeaponPunchDmg(0) + nCurrentCharWeaponPunchDmg(1))
        Case 2: 'kick
            sConfig = sConfig & "_" & CStr(nCurrentCharWeaponKickSkill(0) + nCurrentCharWeaponKickSkill(1))
            sConfig = sConfig & "_" & CStr(nCurrentCharWeaponKickAccy(0) + nCurrentCharWeaponKickAccy(1))
            sConfig = sConfig & "_" & CStr(nCurrentCharWeaponKickDmg(0) + nCurrentCharWeaponKickDmg(1))
        Case 3: 'jk
            sConfig = sConfig & "_" & CStr(nCurrentCharWeaponJkSkill(0) + nCurrentCharWeaponJkSkill(1))
            sConfig = sConfig & "_" & CStr(nCurrentCharWeaponJkAccy(0) + nCurrentCharWeaponJkAccy(1))
            sConfig = sConfig & "_" & CStr(nCurrentCharWeaponJkDmg(0) + nCurrentCharWeaponJkDmg(1))
    End Select
End If

If sCurrentAttackConfig <> sConfig Then Call ClearSavedDamageVsMonster
sCurrentAttackConfig = sConfig

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

Public Function CalcExpNeeded(ByVal startlevel As Long, ByVal exptable As Long) As Currency
'FROM: https://www.mudinfo.net/viewtopic.php?p=7703
On Error GoTo error:
Dim nModifiers() As Integer, i As Long, j As Currency, K As Currency, exp_multiplier As Long, exp_divisor As Long, Ret() As Currency
Dim lastexp As Currency, startexp As Currency, running_exp_tabulation As Currency, billions_tabulator As Currency
Dim potential_new_exp As Currency, ALTERNATE_NEW_EXP As Currency
Dim MAX_UINT As Double, numlevels As Integer, num_divides As Integer

MAX_UINT = 4294967295#
numlevels = 1

ReDim Ret(startlevel To (startlevel + numlevels - 1))

For i = 1 To (startlevel + numlevels - 1)
    startexp = lastexp
    If i = 1 Then
        running_exp_tabulation = 0
    ElseIf i = 2 Then
        running_exp_tabulation = exptable * 10
    Else
        If i <= 26 Then 'levels 1-26
            nModifiers() = GetExpModifiers(i)
            exp_multiplier = nModifiers(0)
            exp_divisor = nModifiers(1)
        ElseIf i <= 55 Then 'levels 27-55
            exp_multiplier = 115
            exp_divisor = 100
        ElseIf i <= 58 Then 'levels 56-58
            exp_multiplier = 109
            exp_divisor = 100
        Else 'levels 59+
            exp_multiplier = 108
            exp_divisor = 100
        End If
        
        If i = 97 Then
            'Debug.Print i
        End If
        
        If exp_multiplier = 0 Or exp_divisor = 0 Then
            potential_new_exp = 0
        Else
            potential_new_exp = running_exp_tabulation * exp_multiplier
        End If
        
        If potential_new_exp > MAX_UINT Then 'UINT ROLLOVER #1
            num_divides = 0
            Do While potential_new_exp > MAX_UINT
                running_exp_tabulation = Fix(running_exp_tabulation / 100)
                potential_new_exp = running_exp_tabulation * exp_multiplier
                num_divides = num_divides + 1
            Loop
            If num_divides > 1 Then
                ALTERNATE_NEW_EXP = Fix((running_exp_tabulation * exp_multiplier * 100) / exp_divisor)
            Else
                ALTERNATE_NEW_EXP = Fix(potential_new_exp / exp_divisor)
            End If
            Do While num_divides > 0
                ALTERNATE_NEW_EXP = ALTERNATE_NEW_EXP * 100
                num_divides = num_divides - 1
            Loop
        Else
            ALTERNATE_NEW_EXP = Fix(potential_new_exp / exp_divisor)
        End If
        
        j = (1000000 * exp_multiplier * billions_tabulator)
        Do While j > MAX_UINT
            j = j - MAX_UINT - 1 'UINT ROLLOVER #2
        Loop
        Do While j >= 1000000000
            j = j - 1000000000
            billions_tabulator = billions_tabulator + 1
        Loop
        
        K = (j + ALTERNATE_NEW_EXP)
        Do While K >= 1000000000
            K = K - 1000000000
            billions_tabulator = billions_tabulator + 1
        Loop
        
        running_exp_tabulation = K
    End If
    
    lastexp = running_exp_tabulation + (billions_tabulator * 1000000000)
    
    If i >= startlevel Then
        Ret(i) = lastexp
    End If
Next i

CalcExpNeeded = Ret(startlevel)

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalcExpNeeded")
Resume out:
End Function

Private Function GetExpModifiers(ByVal nLevel As Integer) As Integer()
Dim Ret(1) As Integer
Ret(0) = 0
Ret(1) = 0

Select Case nLevel
    Case 3:
        Ret(0) = 40
        Ret(1) = 20
        'return [40, 20];
    Case 4, 5:
        Ret(0) = 44
        Ret(1) = 24
        'return [44, 24];
    Case 6, 7:
        Ret(0) = 48
        Ret(1) = 28
        'return [48, 28];
    Case 8, 9:
        Ret(0) = 52
        Ret(1) = 32
        'return [52, 32];
    Case 10, 11:
        Ret(0) = 56
        Ret(1) = 36
        'return [56, 36];
    Case 12, 13:
        Ret(0) = 60
        Ret(1) = 40
        'return [60, 40];
    Case 14, 15:
        Ret(0) = 65
        Ret(1) = 45
        'return [65, 45];
    Case 16, 17:
        Ret(0) = 70
        Ret(1) = 50
        'return [70, 50];
    Case 18:
        Ret(0) = 75
        Ret(1) = 55
        'return [75, 55];
    Case Else:
        If nLevel <= 26 Then
            Ret(0) = 50
            Ret(1) = 40
            'return [50, 40];
        Else
            Ret(0) = 23
            Ret(1) = 20
            'return [23, 20];
        End If

End Select

GetExpModifiers = Ret

'function GetExpModifiers($iLevel) {
'    switch ($iLevel) {
'        Case 3:
'            return [40, 20];
'        Case 4:
'        Case 5:
'            return [44, 24];
'        Case 6:
'        Case 7:
'            return [48, 28];
'        Case 8:
'        Case 9:
'            return [52, 32];
'        Case 10:
'        Case 11:
'            return [56, 36];
'        Case 12:
'        Case 13:
'            return [60, 40];
'        Case 14:
'        Case 15:
'            return [65, 45];
'        Case 16:
'        Case 17:
'            return [70, 50];
'        Case 18:
'            return [75, 55];
'default:
'            if ($iLevel <= 26) {
'                return [50, 40];
'            } else {
'                return [23, 20];
'            }
'    }
'}
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

On Error GoTo ErrorHandler

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
    GoTo ErrorHandler
    Exit Function
End If
 
'don't allow drop-down width to
'exceed screen.width
 
   If lWidth > Screen.Width \ Screen.TwipsPerPixelX - 20 Then _
    lWidth = Screen.Width \ Screen.TwipsPerPixelX - 20

lRet = SendMessageLong(Combo.hWnd, CB_SETDROPPEDWIDTH, lWidth, 0)

AutoSizeDropDownWidth = lRet > 0
ErrorHandler:
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

Public Sub PullItemDetail(DetailTB As TextBox, LocationLV As ListView, Optional ByVal nAttackType As Integer)
Dim sStr As String, sAbil As String, x As Integer, sCasts As String, nPercent As Integer
Dim sNegate As String, sClasses As String, sRaces As String, sClassOk As String
Dim sUses As String, sGetDrop As String, oLI As ListItem, nNumber As Long
Dim y As Integer, bCompareWeapon As Boolean, bCompareArmor As Boolean, nInvenSlot1 As Integer, nInvenSlot2 As Integer
Dim sCompareText1 As String, sCompareText2 As String, tabItems1 As Recordset, tabItems2 As Recordset
Dim sTemp1 As String, sTemp2 As String, sTemp3 As String, bFlag1 As Boolean, bFlag2 As Boolean
Dim nClassRestrictions(0 To 2, 0 To 9) As Long, nRaceRestrictions(0 To 2, 0 To 9) As Long
Dim nNegateSpells(0 To 2, 0 To 9) As Long, nAbils(0 To 2, 0 To 19, 0 To 2) As Long, sAbilText(0 To 2, 0 To 19) As String
Dim nReturnValue As Long, nMatchReturnValue As Long, sClassOk1 As String, sClassOk2 As String
Dim sCastSp1 As String, sCastSp2 As String, bCastSpFlag(0 To 2) As Boolean, nPct(0 To 2) As Integer
Dim tWeaponDmg As tAttackDamage, sWeaponDmg As String, nSpeedAdj As Integer

DetailTB.Text = ""
If bStartup Then Exit Sub

nSpeedAdj = 100
If frmMain.chkWeaponOptions(3).Value = 1 Then
    If nAttackType = 0 Then nAttackType = frmMain.cmbWeaponCombos(1).ItemData(frmMain.cmbWeaponCombos(1).ListIndex)
    If frmMain.chkWeaponOptions(4).Value = 1 Then nSpeedAdj = 85
ElseIf nAttackType = 0 Then
    nAttackType = 5
End If

nInvenSlot1 = -1
nInvenSlot2 = -1

'sStr = ClipNull(tabItems.Fields("Name")) & " (" & tabItems.Fields("Number") & ")"

On Error GoTo error:

nNumber = tabItems.Fields("Number")

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
                If Not DetailTB.name = "txtWeaponCompareDetail" And _
                    Not DetailTB.name = "txtWeaponDetail" Then
                    
                    sTemp1 = GetAbilityStats(nAbils(0, x, 0), nAbils(0, x, 1), LocationLV, , True)
                    sAbilText(0, x) = sTemp1
                    sAbil = AutoAppend(sAbil, sTemp1)
                End If
                
            Case 22, 105, 106, 135:  '22-acc, 105-acc, 106-acc, 135-minlvl
                If Not DetailTB.name = "txtWeaponCompareDetail" And _
                    Not DetailTB.name = "txtWeaponDetail" And _
                    Not DetailTB.name = "txtArmourCompareDetail" And _
                    Not DetailTB.name = "txtArmourDetail" Then
    
                    sTemp1 = GetAbilityStats(nAbils(0, x, 0), nAbils(0, x, 1), LocationLV, , True)
                    sAbilText(0, x) = sTemp1
                    sAbil = AutoAppend(sAbil, sTemp1)
                End If
                
            Case 59: 'class ok
                sTemp1 = GetClassName(nAbils(0, x, 1))
                sAbilText(0, x) = sTemp1
                sClassOk = AutoAppend(sClassOk, sTemp1)
                
            Case 43: 'casts spell
                'nSpellNest = 0 'make sure this doesn't nest too deep
                sCasts = AutoAppend(sCasts, "[" & GetSpellName(nAbils(0, x, 1), bHideRecordNumbers) _
                    & ", " & PullSpellEQ(True, 0, nAbils(0, x, 1), , , , True))
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
    
    tWeaponDmg = CalculateAttack(nAttackType, tabItems.Fields("Number"), _
                    IIf(frmMain.chkWeaponOptions(3).Value = 1, True, False), _
                    False, _
                    nSpeedAdj, _
                    IIf(frmMain.chkWeaponOptions(3).Value = 1, Val(frmMain.txtWeaponExtras(2).Text), 0), _
                    IIf(frmMain.chkWeaponOptions(3).Value = 1, Val(frmMain.txtWeaponExtras(3).Text), 0), _
                    IIf(frmMain.chkWeaponOptions(3).Value = 1, Val(frmMain.txtWeaponExtras(4).Text), 0), _
                    sCasts)
        
    If tWeaponDmg.nSwings > 0 Then
        Select Case nAttackType
            Case 1: sWeaponDmg = "Punch Damage: "
            Case 2: sWeaponDmg = "Kick Damage: "
            Case 3: sWeaponDmg = "Jumpkick Damage: "
            Case 4: sWeaponDmg = "Backstab Damage: "
            Case 6: sWeaponDmg = "Bash Damage: "
            Case 7: sWeaponDmg = "Smash Damage: "
            Case Else:
                sWeaponDmg = "Damage: "
        End Select
        
        sWeaponDmg = sWeaponDmg & tWeaponDmg.nRoundTotal & "/round @ " & Round(tWeaponDmg.nSwings, 1) & " swings w/" & tWeaponDmg.nHitChance & "% hit chance"
        sWeaponDmg = sWeaponDmg & " - Avg Hit: " & tWeaponDmg.nAvgHit
        
        If tWeaponDmg.nAvgCrit > 0 And tWeaponDmg.nCritChance > 0 Then
            sWeaponDmg = AutoAppend(sWeaponDmg, "Avg Crit: " & tWeaponDmg.nAvgCrit)
            sWeaponDmg = sWeaponDmg & " (" & tWeaponDmg.nCritChance & "%"
            If tWeaponDmg.nQnDBonus > 0 Then sWeaponDmg = sWeaponDmg & " w/" & tWeaponDmg.nQnDBonus & "qnd"
            sWeaponDmg = sWeaponDmg & ")"
        End If
        
        If tWeaponDmg.nAvgExtraHit > 0 Then
            sWeaponDmg = AutoAppend(sWeaponDmg, "Avg Extra: " & tWeaponDmg.nAvgExtraHit)
            If tWeaponDmg.nAvgExtraHit <> tWeaponDmg.nAvgExtraSwing Then
                sWeaponDmg = sWeaponDmg & " (avg " & tWeaponDmg.nAvgExtraSwing & "/swing)"
            Else
                
            End If
        End If
        
        sWeaponDmg = sWeaponDmg & vbCrLf
        
        If (frmMain.chkGlobalFilter.Value = 1 Or frmMain.chkWeaponOptions(3).Value = 1) And tabItems.Fields("StrReq") > Val(frmMain.txtCharStats(0).Text) Then
            sWeaponDmg = sWeaponDmg & "Notice: Character Strength (" & Val(frmMain.txtCharStats(0).Text) & ") < Strength Requirement (" & tabItems.Fields("StrReq") & ")" & vbCrLf
        End If
        
        sWeaponDmg = sWeaponDmg & vbCrLf
    End If
End If

DetailTB.Text = sWeaponDmg & sStr

If LocationLV.ListItems.Count > 0 Then
    If nLastItemSortCol > LocationLV.ColumnHeaders.Count Then nLastItemSortCol = 1
    If nLastItemSortCol = 1 Then
        Call SortListViewByTag(LocationLV, 1, ldtnumber, False)
    Else
        Call SortListView(LocationLV, nLastItemSortCol, ldtstring, False)
    End If
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
Dim sCash As String, nCash As Currency, nPercent As Integer, nTest As Long, nMobExpPerHour() As Currency
Dim oLI As ListItem, nExp As Currency, nLocalMonsterDamage As MonAttackSimReturn, nMonsterEnergy As Long
Dim sReducedCoin As String, nReducedCoin As Currency, nDamage As Currency, nRestingRate As Double
Dim nAvgDmg As Long, nExpDmgHP As Currency, nExpPerHour As Currency, nExpPerHourEA As Currency, nPossyPCT As Currency, nMobDodge As Integer
Dim nScriptValue As Currency, nLairPCT As Currency, nPossSpawns As Long, sPossSpawns As String, sScriptValue As String
Dim tAvgLairInfo As LairInfoType, sArr() As String, bHasAttacks As Boolean, bSpacer As Boolean
Dim nDamageOut As Long, nHPRegen As Long, nSpellCastLVL As Long, nSpellDuration As Long, sDefenseDesc As String
Dim nParty As Integer, nCharHealth As Long, nMaxLairsBeforeRegen As Currency, bHasAntiMagic As Boolean
Dim nExpReductionLairRatio As Double, sExpReductionLairRatio As String, tSpellCast As tSpellCastValues
Dim nExpReductionMaxLairs As Double, sExpReductionMaxLairs As String, tAttack As tAttackDamage
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
        If frmMain.chkGlobalFilter.Value = 1 And frmMain.cmbGlobalAlignment.ListIndex = 3 Then 'evil aligned
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
End If

Set oLI = DetailLV.ListItems.Add()
oLI.Text = "HPs"
oLI.ForeColor = RGB(204, 0, 0)
oLI.ListSubItems.Add (1), "Detail", tabMonsters.Fields("HP") & " (Regens: " & tabMonsters.Fields("HPRegen") & " HPs/click)"
oLI.ListSubItems(1).ForeColor = RGB(204, 0, 0)

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

For x = 0 To 9 'abilities
    If tabMonsters.Fields("Abil-" & x) > 0 And Not tabMonsters.Fields("Abil-" & x) = 146 Then '146=guarded by (handled below)
        If sAbil <> "" Then sAbil = sAbil & ", "
        sAbil = sAbil & GetAbilityStats(tabMonsters.Fields("Abil-" & x), tabMonsters.Fields("AbilVal-" & x))
        If Right(sAbil, 2) = ", " Then sAbil = Left(sAbil, Len(sAbil) - 2)
        
        If tabMonsters.Fields("Abil-" & x) = 34 And tabMonsters.Fields("AbilVal-" & x) > 0 Then 'dodge
            nMobDodge = tabMonsters.Fields("AbilVal-" & x)
            If frmMain.chkGlobalFilter.Value = 1 And Val(frmMain.lblInvenCharStat(10).Tag) > 0 Then
                sAbil = sAbil & " (" & Fix((tabMonsters.Fields("AbilVal-" & x) * 10) / Fix(Val(frmMain.lblInvenCharStat(10).Tag) / 8)) & "% @ " _
                    & Val(frmMain.lblInvenCharStat(10).Tag) & " accy)"
            End If
        ElseIf tabMonsters.Fields("Abil-" & x) = 51 Then
            bHasAntiMagic = True
        End If
    End If
Next x
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
            oLI.Text = ""
            
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

bSpacer = False
For x = 0 To 9 'item drops
    If Not tabMonsters.Fields("DropItem-" & x) = 0 Then bSpacer = True
    If bSpacer Then Exit For
Next x
If bSpacer Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
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
    End If
Next

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
If bHasAttacks Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Attacks"
    oLI.Bold = True
    If nNMRVer >= 1.71 Then
        If Not tabMonsters.Fields("Energy") = 0 Then
            oLI.ListSubItems.Add (1), "Detail", nMonsterEnergy & " energy/round"
        End If
    End If
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
End If

bSpacer = False
If bHasAttacks Then
    If nNMRVer >= 1.8 Then
        nLocalMonsterDamage = CalculateMonsterAvgDmg(nMonsterNum, 0) 'this is to get max damage
        nLocalMonsterDamage.nAverageDamage = tabMonsters.Fields("AvgDmg")
    Else
        nLocalMonsterDamage = CalculateMonsterAvgDmg(nMonsterNum, nMonsterSimRounds)
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
            oLI.ListSubItems(1).Text = oLI.ListSubItems(1).Text & "   * before character defenses, " & nMonsterSimRounds & " round sim"
        End If
        oLI.ListSubItems(1).ForeColor = RGB(204, 0, 0)
        bSpacer = True
    End If
    
    nDamage = -1
    If frmMain.chkGlobalFilter.Value = 1 Then
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
            
            oLI.ListSubItems(1).Text = oLI.ListSubItems(1).Text & "   * versus current character defenses, " & nMonsterSimRounds & " round sim"
            oLI.ListSubItems(1).ForeColor = RGB(144, 4, 214)
            bSpacer = True
        End If
    End If
    
    nDamage = -1
    If frmMain.optMonsterFilter(1).Value = True And Val(frmMain.txtMonsterLairFilter(0).Text) > 1 Then 'vs party
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
            
            oLI.ListSubItems(1).Text = oLI.ListSubItems(1).Text & "   * versus current PARTY defenses, " & nMonsterSimRounds & " round sim"
            oLI.ListSubItems(1).ForeColor = &H40C0&
            bSpacer = True
        End If
    End If
    If bSpacer Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = ""
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
        End If
    Next
    If y > 0 Then 'add blank line if there was entried added
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = ""
    End If
    
    
    nPercent = 0
    y = 0
    For x = 0 To 4 'attacks
        If tabMonsters.Fields("AttType-" & x) > 0 And tabMonsters.Fields("AttType-" & x) <= 3 And tabMonsters.Fields("Att%-" & x) > 0 Then
            y = y + 1
            Set oLI = DetailLV.ListItems.Add()
            
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
                            nTest = SpellHasAbility(tabMonsters.Fields("AttAcc-" & x), 1) '1=damage
                            If nTest > -1 Then
                                Set oLI = DetailLV.ListItems.Add()
                                oLI.Text = ""
                                oLI.ListSubItems.Add (1), "Detail", "This is an invalid spell and will not be cast (area attack spells from regular monster casts must use ability 17, Damage-MR)."
                                oLI.ListSubItems(1).Bold = True
                            End If
                        End If
                    End If
            End Select
            
            Set oLI = DetailLV.ListItems.Add()
            oLI.Text = ""
        End If
    Next
Else
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
End If

'For x = 0 To 9 'abilities
'    If Not tabMonsters.Fields("Abil-" & x) = 0 Then
'        Select Case tabMonsters.Fields("Abil-" & x)
'            Case 0:
'            Case 146: 'mon guards
'                If sMonGuards <> "" Then sMonGuards = sMonGuards & ", "
'                sMonGuards = sMonGuards & GetMonsterName(tabMonsters.Fields("AbilVal-" & x), bHideRecordNumbers)
'                tabMonsters.Seek "=", nMonsterNum
'
'            Case Else:
'                If sAbil <> "" Then sAbil = sAbil & ", "
'                sAbil = sAbil & GetAbilityStats(tabMonsters.Fields("Abil-" & x), tabMonsters.Fields("AbilVal-" & x))
'                If Right(sAbil, 2) = ", " Then sAbil = Left(sAbil, Len(sAbil) - 2)
'        End Select
'    End If
'Next
'
'DetailTB.Text = sAbil
'If Not sMonGuards = "" Then
'    sMonGuards = "Guarded By: " & sMonGuards
'    If sAbil = "" Then
'        DetailTB.Text = sMonGuards
'    Else
'        DetailTB.Text = sAbil & ", " & sMonGuards
'    End If
'End If

If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum

nCharHealth = 1
nHPRegen = 0
nParty = 1
nRestingRate = 0
If tAvgLairInfo.nMobs > 0 Then nAvgDmg = tAvgLairInfo.nAvgDmg

If frmMain.chkGlobalFilter.Value = 1 And (frmMain.optMonsterFilter(1).Value = False Or Val(frmMain.txtMonsterLairFilter(0).Text) < 2) Then 'no party, vs char
    nCharHealth = Val(frmMain.lblCharMaxHP.Tag)
    nHPRegen = Val(frmMain.lblCharRestRate.Tag)
    
ElseIf frmMain.optMonsterFilter(1).Value = True And Val(frmMain.txtMonsterLairFilter(0).Text) > 1 Then 'vs party
    nParty = Val(frmMain.txtMonsterLairFilter(0).Text)
    nCharHealth = Val(frmMain.txtMonsterLairFilter(5).Text)
    If nCharHealth < 1 Then
        frmMain.txtMonsterLairFilter(7).Text = 1
        nCharHealth = 1
    End If
    nCharHealth = nCharHealth * Val(frmMain.txtMonsterLairFilter(0).Text) 'note: nCharHealth is avg * party to match tAvgLairInfo values
    nHPRegen = Val(frmMain.txtMonsterLairFilter(7).Text)
    
Else
    nCharHealth = nAvgDmg * 2
    nHPRegen = nCharHealth * 0.05
End If

If nCharHealth < 1 Then nCharHealth = 1
If nHPRegen < 1 Then nHPRegen = 1
If nParty > 6 Then nParty = 6
If nParty < 1 Then nParty = 1

'nDamageOut = Val(frmMain.txtMonsterDamageOUT(0).Text) * nParty
'nDamageOut = 0
'nDamageOutSpell = 0
nDamageOut = 0
If nParty > 1 Then
    nDamageOut = Val(frmMain.txtMonsterDamageOUT.Text) * nParty
    If nNMRVer >= 1.83 Then Call AddMonsterDamageOutText(DetailLV, nDamageOut & "/round (party)")
    'nDamageOutSpell = Val(frmMain.txtMonsterDamageOUT(1).Text) * nParty
Else
    'nCurrentAttackType (from frmPopUpOptions): 0-none/one-shot, 1-weapon, 2-spell user, 3-spell any, 4-MA, 5-manual, 6-bash, 7-smash
    'CalculateAttack > nAttackType: 1-punch, 2-kick, 3-jumpkick, 4-surprise, 5-normal, 6-bash, 7-smash
    Select Case nCurrentAttackType
        Case 1, 6, 7: 'eq'd weapon, bash, smash
            If nCurrentCharWeaponNumber(0) > 0 Then
                If nCurrentAttackType = 6 Then 'bash w/wep
                    tAttack = CalculateAttack(6, nCurrentCharWeaponNumber(0), True, False, 100, tabMonsters.Fields("ArmourClass"), tabMonsters.Fields("DamageResist"), nMobDodge)
                    nDamageOut = tAttack.nRoundTotal
                ElseIf nCurrentAttackType = 7 Then 'smash w/wep
                    tAttack = CalculateAttack(7, nCurrentCharWeaponNumber(0), True, False, 100, tabMonsters.Fields("ArmourClass"), tabMonsters.Fields("DamageResist"), nMobDodge)
                    nDamageOut = tAttack.nRoundTotal
                Else 'EQ'd Weapon reg attack
                    tAttack = CalculateAttack(5, nCurrentCharWeaponNumber(0), True, False, 100, tabMonsters.Fields("ArmourClass"), tabMonsters.Fields("DamageResist"), nMobDodge)
                    nDamageOut = tAttack.nRoundTotal
                End If
                Call AddMonsterDamageOutText(DetailLV, tAttack.nRoundTotal & "/round (" & tAttack.sAttackDesc & ")", tAttack.sAttackDetail)
            Else
                GoTo no_attack:
            End If
            
        Case 2, 3:
            '2-spell learned: GetSpellShort(nCurrentAttackSpellNum) & " @ " & Val(txtGlobalLevel(0).Text)
            '3-spell any: GetSpellShort(nCurrentAttackSpellNum) & " @ " & nCurrentAttackSpellLVL
            If nCurrentAttackSpellNum <= 0 Then GoTo no_attack:
            If frmMain.chkGlobalFilter.Value = 1 Then
                tSpellCast = CalculateSpellCast(nCurrentAttackSpellNum, Val(frmMain.txtGlobalLevel(0).Text), Val(frmMain.lblCharSC.Tag), _
                    tabMonsters.Fields("MagicRes"), bHasAntiMagic)
            Else
                tSpellCast = CalculateSpellCast(nCurrentAttackSpellNum, 0, 0, tabMonsters.Fields("MagicRes"), bHasAntiMagic)
            End If
            nDamageOut = tSpellCast.nAvgRoundDmg
            Call AddMonsterDamageOutText(DetailLV, tSpellCast.sAvgRound & "/round (" & tSpellCast.sSpellName & ")", tSpellCast.sMMA)
            
        Case 4: 'martial arts attack
            '1-Punch, 2-Kick, 3-JumpKick
            Select Case nCurrentAttackMA
                Case 2: 'kick
                    tAttack = CalculateAttack(2, , True, False, 100, tabMonsters.Fields("ArmourClass"), tabMonsters.Fields("DamageResist"), nMobDodge)
                    nDamageOut = tAttack.nRoundTotal
                Case 3: 'jumpkick
                    tAttack = CalculateAttack(3, , True, False, 100, tabMonsters.Fields("ArmourClass"), tabMonsters.Fields("DamageResist"), nMobDodge)
                    nDamageOut = tAttack.nRoundTotal
                Case Else: 'punch
                    tAttack = CalculateAttack(1, , True, False, 100, tabMonsters.Fields("ArmourClass"), tabMonsters.Fields("DamageResist"), nMobDodge)
                    nDamageOut = tAttack.nRoundTotal
            End Select
            Call AddMonsterDamageOutText(DetailLV, tAttack.nRoundTotal & "/round (" & tAttack.sAttackDesc & ")", tAttack.sAttackDetail)
            
        Case 5: 'manual
            nDamageOut = nCurrentAttackManual
            'nDamageOutSpell = nCurrentAttackManualMag
            If nNMRVer >= 1.83 Then Call AddMonsterDamageOutText(DetailLV, nCurrentAttackManual & "/round (manual)")
            
        Case Else: '1-Shot All
            nDamageOut = 9999999
            If nNMRVer >= 1.83 Then Call AddMonsterDamageOutText(DetailLV, "(assuming one-shot)")
            'nDamageOutSpell = 9999999
            
    End Select
End If
no_attack:
'If nDamageOut < 0 Then nDamageOut = 0
'If nDamageOutSpell < 0 Then nDamageOutSpell = 0

If nExp <= 1 Or tabMonsters.Fields("HP") < 1 Then GoTo done_scripting:

Set oLI = DetailLV.ListItems.Add()
oLI.Text = "Scripting"
oLI.Bold = True
'Set oLI = DetailLV.ListItems.Add()
'oLI.Text = ""

If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum

'================================================================
'================================================================
'==============   2025.06.15 ===============
'================================================================

nAvgDmg = 0
nExpDmgHP = 0
sDefenseDesc = ""
If frmMain.chkGlobalFilter.Value = 1 And nMonsterDamageVsChar(nMonsterNum) >= 0 Then
    nAvgDmg = nMonsterDamageVsChar(nMonsterNum)
    sDefenseDesc = " (vs char defenses)"
ElseIf nNMRVer >= 1.8 Then
    nAvgDmg = tabMonsters.Fields("AvgDmg")
    sDefenseDesc = " (vs default defenses)"
ElseIf nMonsterDamageVsDefault(nMonsterNum) >= 0 Then
    nAvgDmg = nMonsterDamageVsDefault(nMonsterNum)
    sDefenseDesc = " (vs default defenses)"
Else
    nAvgDmg = 0
End If

If nNMRVer < 1.83 Then GoTo script_value:

If nNMRVer >= 1.83 And InStr(1, tabMonsters.Fields("Summoned By"), "lair", vbTextCompare) > 0 Then
    tLastAvgLairInfo = GetAverageLairValuesFromLocs(tabMonsters.Fields("Summoned By"))
ElseIf tLastAvgLairInfo.sGroupIndex <> "" Then
    tLastAvgLairInfo = GetLairInfo("") 'reset
End If
tAvgLairInfo = tLastAvgLairInfo

If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum

Set oLI = DetailLV.ListItems.Add()
oLI.Text = "Scripting Estimate"

If tabMonsters.Fields("RegenTime") = 0 And tAvgLairInfo.nMobs > 0 Then
    
    If IsMobKillable(nDamageOut, nCharHealth, nAvgDmg, tAvgLairInfo.nAvgHP, nHPRegen) = False Then
        nExpPerHour = -1
        nRestingRate = 1
        sTemp2 = "(" & nDamageOut & "dmg/" & nCharHealth & "hp vs " & nAvgDmg & "dmg/" & tAvgLairInfo.nAvgHP & "hp)"
    Else
        nExpPerHour = tAvgLairInfo.nAvgExp
        nRestingRate = tAvgLairInfo.nRestRate
    End If
    
    nMaxLairsBeforeRegen = nTheoreticalAvgMaxLairsPerRegenPeriod
    If nExpPerHour > 0 And tAvgLairInfo.nMaxRegen > 0 Then
        nMaxLairsBeforeRegen = Round(nMaxLairsBeforeRegen / tAvgLairInfo.nMaxRegen, 2)
        If nDamageOut > 0 And (nDamageOut * tAvgLairInfo.nMaxRegen) < tAvgLairInfo.nAvgHP Then
            nExpReductionMaxLairs = ((nDamageOut * tAvgLairInfo.nMaxRegen) / tAvgLairInfo.nAvgHP)
            nMaxLairsBeforeRegen = Round(nMaxLairsBeforeRegen * nExpReductionMaxLairs, 2)
            If tAvgLairInfo.nMobs > nMaxLairsBeforeRegen Then
                sExpReductionMaxLairs = Round((1 - nExpReductionMaxLairs) * 100) & "% kill time reduction due to time spent attacking (driven by [dmg out] vs [mob HP])"
            Else
                nExpReductionMaxLairs = 0
            End If
        End If
    End If
    
    If nExpPerHour > 0 Then
        nPossSpawns = InstrCount(tabMonsters.Fields("Summoned By"), "Group:") + tAvgLairInfo.nMobs
        If nPossSpawns > (tAvgLairInfo.nMobs * nMonsterLairRatioMultiplier) Then '(nmobs = # lairs) ... indication of a lot of walking distance between lairs
            nExpReductionLairRatio = ((tAvgLairInfo.nMobs * nMonsterLairRatioMultiplier) / nPossSpawns)
            sExpReductionLairRatio = Round((1 - nExpReductionLairRatio) * 100) & "% exp reduction due to the ratio of lairs to non-lairs (meaning increased travel time, presumably)."
            nExpPerHour = Round(nExpReductionLairRatio * nExpPerHour)
        End If
    End If
    
ElseIf tabMonsters.Fields("RegenTime") > 0 Or InStr(1, tabMonsters.Fields("Summoned By"), "Room", vbTextCompare) > 0 Then
        
    nMobExpPerHour() = CalcMobExpPerHour(nMonsterNum, nDamageOut, nCharHealth, nAvgDmg, tabMonsters.Fields("HP"), _
        nHPRegen, tabMonsters.Fields("HPRegen"), Val(frmMain.txtMonsterDamage.Text), nParty)
    
    nExpPerHour = nMobExpPerHour(0)
    
    'If tAvgLairInfo.nMobs > 0 And tabMonsters.Fields("RegenTime") > 0 And tAvgLairInfo.nRestRate > 0 Then
    '    nRestingRate = tAvgLairInfo.nRestRate
    'Else
        nRestingRate = nMobExpPerHour(1)
    'End If
        
Else
    nExpPerHour = 0
End If

If nExpPerHour > 0 And frmMain.optMonsterFilter(1).Value = True And Val(frmMain.txtMonsterLairFilter(0).Text) > 1 Then
    nExpPerHourEA = Round(nExpPerHour / Val(frmMain.txtMonsterLairFilter(0).Text))
Else
    nExpPerHourEA = nExpPerHour
End If

If nExpPerHour > 0 Then
    If nExpPerHour > 1000000 Then
        sTemp = Format((nExpPerHour / 1000000), "#,#.0") & " M"
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
    If tabMonsters.Fields("RegenTime") = 0 And tAvgLairInfo.nMobs > 0 Then
        sTemp = "The lairs of this mob"
    Else
        sTemp = "This mob"
    End If
    sTemp = sTemp & " deemed undefeatable against current stats " & sTemp2 & "."
Else
    sTemp = "0"
End If

oLI.ListSubItems.Add (1), "Detail", sTemp

If Len(sExpReductionMaxLairs) > 0 Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
    oLI.ListSubItems.Add (1), "Detail", sExpReductionMaxLairs
End If

If Len(sExpReductionLairRatio) > 0 Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
    oLI.ListSubItems.Add (1), "Detail", sExpReductionLairRatio
End If

If nRestingRate > 0 And nExpPerHour <> -1 Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
    oLI.ListSubItems.Add (1), "Detail", Round(nRestingRate * 100) & "% combat time reduction for time spent resting (driven by [dmg out], excess incoming damage over [HEALS], and [rest HP] rate)"
End If

If tabMonsters.Fields("RegenTime") = 0 And tAvgLairInfo.nMobs > 0 Then
    If frmMain.chkGlobalFilter.Value = 1 And nParty < 2 Then 'no party, vs char
        If nAvgDmg > Val(frmMain.txtMonsterDamage.Text) Or nRestingRate > 0 Then
            If bMonsterDamageVsCharCalculated = False Then
                Set oLI = DetailLV.ListItems.Add()
                oLI.Text = ""
                oLI.ListSubItems.Add (1), "Detail", "Character damage vs all monsters not calculated. Base damage utilized where missing. Calculate from options menu."
                oLI.ListSubItems(1).Bold = True
            ElseIf bDontPromptCalcCharMonsterDamage = False Then
                Set oLI = DetailLV.ListItems.Add()
                oLI.Text = ""
                oLI.ListSubItems.Add (1), "Detail", "Character damage vs all monsters may be stale. Calculate from options menu."
                oLI.ListSubItems(1).Bold = True
            End If
        End If
    ElseIf nParty > 1 Then 'vs party
        If bMonsterDamageVsPartyCalculated = False Then
            Set oLI = DetailLV.ListItems.Add()
            oLI.Text = ""
            oLI.ListSubItems.Add (1), "Detail", "Party damage vs all monsters not calculated. Base damage utilized where missing. Click the [Calc. Party DMG] button."
            oLI.ListSubItems(1).Bold = True
        ElseIf bDontPromptCalcPartyMonsterDamage = False Then
            Set oLI = DetailLV.ListItems.Add()
            oLI.Text = ""
            oLI.ListSubItems.Add (1), "Detail", "Party damage vs all monsters may be stale. Click the [Calc. Party DMG] button."
            oLI.ListSubItems(1).Bold = True
        End If
    End If
End If

Set oLI = DetailLV.ListItems.Add()
oLI.Text = ""

'================================================================
'================================================================
'==============   2025.06.15 ===============
'================================================================

'a lot of this repeated in addmonsterlv
script_value:

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

If frmMain.optMonsterFilter(1).Value = True And Val(frmMain.txtMonsterLairFilter(0).Text) > 1 And nMonsterDamageVsParty(nMonsterNum) >= 0 Then 'vs party
    nExpDmgHP = 0
    nAvgDmg = nMonsterDamageVsParty(nMonsterNum)
    If nAvgDmg > 0 Or tabMonsters.Fields("HP") > 0 Then
        nExpDmgHP = Round(nExp / ((nAvgDmg * 2) + tabMonsters.Fields("HP")), 2) * 100
    Else
        nExpDmgHP = nExp
    End If
    
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Exp/((Dmg*2)+HP)"
    oLI.ListSubItems.Add (1), "Detail", IIf(nExpDmgHP > 0, Format(nExpDmgHP, "#,#"), 0) & " (" & nExp & " / ((" & nAvgDmg & " x 2) + " & tabMonsters.Fields("HP") & ")) * 100" & " (vs party defenses)"
End If

'a lot of this repeated in addmonsterlv
nPossSpawns = 0
nLairPCT = 0
nMaxLairsBeforeRegen = nTheoreticalAvgMaxLairsPerRegenPeriod
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
    If nAvgDmg > 0 Or tabMonsters.Fields("HP") > 0 Then
        nExpDmgHP = Round(nExp / (((nAvgDmg * 2) + tabMonsters.Fields("HP")) * nPossyPCT), 2) * 100
    Else
        nExpDmgHP = nExp
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

If tAvgLairInfo.nMobs > 0 Or (nNMRVer >= 1.82 And nMonsterPossy(nMonsterNum) > 0) Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Lair Stats"
    oLI.Bold = True
    
    If tabMonsters.Fields("RegenTime") > 0 And tAvgLairInfo.nMobs > 0 Then
        oLI.ListSubItems.Add (1), "Detail", "Note: Mobs with regen time >0 are not included in lair stats"
    End If
End If

If tAvgLairInfo.nMobs > 0 Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "Total Lairs"
    oLI.ListSubItems.Add (1), "Detail", tAvgLairInfo.nMobs
    
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "AVG # Mobs/Lair"
    If nMonsterSpawnChance(nMonsterNum) > 0 Then
        oLI.ListSubItems.Add (1), "Detail", nMonsterPossy(nMonsterNum) _
            & "  (" & (nMonsterSpawnChance(nMonsterNum) * 100) & "% chance for this monster to spawn/lair)"
    Else
        oLI.ListSubItems.Add (1), "Detail", tAvgLairInfo.nMaxRegen
    End If
    
    If nNMRVer >= 1.83 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "AVG Exp"
        oLI.ListSubItems.Add (1), "Detail", PutCommas(tabMonsters.Fields("AvgLairExp")) & "  (" & PutCommas(Round(tabMonsters.Fields("AvgLairExp") / tAvgLairInfo.nMaxRegen)) & "/mob)"
    End If

    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "AVG DMG/Round"
    oLI.ListSubItems.Add (1), "Detail", PutCommas(tAvgLairInfo.nAvgDmg) & "  (" & Round(tAvgLairInfo.nAvgDmg / tAvgLairInfo.nMaxRegen) & "/mob)"
    If tAvgLairInfo.nDamageAdjustment > 0 Then
        oLI.ListSubItems(1).Text = oLI.ListSubItems(1).Text & " - " & tAvgLairInfo.nDamageAdjustment & " dmg/mob mitigated"
    End If
    
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "AVG HP"
    oLI.ListSubItems.Add (1), "Detail", PutCommas(tAvgLairInfo.nAvgHP) & "  (" & Round(tAvgLairInfo.nAvgHP / tAvgLairInfo.nMaxRegen) & "/mob)"
    
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "AVG AC/DR"
    oLI.ListSubItems.Add (1), "Detail", tAvgLairInfo.nAvgAC & "/" & tAvgLairInfo.nAvgDR
    
    If tAvgLairInfo.nAvgDodge > 0 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "AVG Dodge"
        If frmMain.chkGlobalFilter.Value = 1 And Val(frmMain.lblInvenCharStat(10).Tag) > 0 Then
            oLI.ListSubItems.Add (1), "Detail", tAvgLairInfo.nAvgDodge _
                & " (" & Fix((tAvgLairInfo.nAvgDodge * 10) / Fix(Val(frmMain.lblInvenCharStat(10).Tag) / 8)) & "% @ " & Val(frmMain.lblInvenCharStat(10).Tag) & " accy)"
        Else
            oLI.ListSubItems.Add (1), "Detail", tAvgLairInfo.nAvgDodge
        End If
    End If
    
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = "AVG MR"
    oLI.ListSubItems.Add (1), "Detail", tAvgLairInfo.nAvgMR
    
    If InStr(1, tAvgLairInfo.sMobList, ",", vbTextCompare) > 0 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = "Other Lair Mobs"
        sArr() = Split(tAvgLairInfo.sMobList, ",")
        y = 0
        For x = 0 To UBound(sArr())
            If Val(sArr(x)) <> nMonsterNum Then
                If y > 0 Then
                    Set oLI = DetailLV.ListItems.Add()
                    oLI.Text = ""
                End If
                oLI.ListSubItems.Add (1), "Detail", GetMonsterName(sArr(x), bHideRecordNumbers)
                tabMonsters.Seek "=", nMonsterNum
                oLI.Tag = "monster"
                oLI.ListSubItems(1).Tag = sArr(x)
                y = y + 1
                If y > 9 And UBound(sArr()) > 14 Then
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

If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum

If Not frmMain.bDontLookupMonRegen Then
    If Len(tabMonsters.Fields("Summoned By")) > 4 Then
        Set oLI = DetailLV.ListItems.Add()
        oLI.Text = ""
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

Private Sub AddMonsterDamageOutText(ByRef DetailLV As ListView, ByVal sDetail As String, Optional ByVal sDetail2 As String)
On Error GoTo error:
Dim oLI As ListItem

'Set oLI = DetailLV.ListItems.Add()
'oLI.Text = ""

Set oLI = DetailLV.ListItems.Add()
oLI.Bold = True
oLI.Text = "Damage vs Mob"
oLI.ListSubItems.Add (1), "Detail", sDetail

If Not sDetail2 = "" Then
    Set oLI = DetailLV.ListItems.Add()
    oLI.Text = ""
    oLI.ListSubItems.Add (1), "Detail", sDetail2
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
Dim oLI As ListItem, tCostType As typItemCostDetail, nCopper As Currency, sCopper As String
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
        
        sCharmMod = nCharmMod & "% cost (pre-markup)"
    Else
        nCharmMod = 1 - ((Fix(nCharm / 5) - 10) / 100)
        If nCharmMod > 1 Then
            sCharmMod = Abs(1 - CCur(nCharmMod)) * 100 & "% Markup"
        ElseIf nCharmMod < 1 Then
            sCharmMod = Val(1 - CCur(nCharmMod)) * 100 & "% Discount"
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
        DetailLV.ColumnHeaders.Add 1, "Number", "#", 0, lvwColumnLeft
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
'        oLI.ListSubItems.Add (4), "Rgn %", tabShops.Fields("%-" & x)
'        oLI.ListSubItems.Add (5), "Rgn #", tabShops.Fields("Amount-" & x)
        
        If bShowSell Then
            tCostType = GetItemCost(tabShops.Fields("Item-" & x), 0)
        Else
            tCostType = GetItemCost(tabShops.Fields("Item-" & x), tabShops.Fields("Markup%"))
        End If
        
        Select Case tCostType.Coin
            Case 0: 'GetCostType = "Copper"
                nCopper = Val(tCostType.Cost)
            Case 1: 'GetCostType = "Silver"
                nCopper = Val(tCostType.Cost) * 10
            Case 2: 'GetCostType = "Gold"
                nCopper = Val(tCostType.Cost) * 100
            Case 3: 'GetCostType = "Platinum"
                nCopper = Val(tCostType.Cost) * 10000
            Case 4: 'GetCostType = "Runic"
                nCopper = Val(tCostType.Cost) * 1000000
            Case Else:
                nCopper = tCostType.Cost
        End Select
        
        If nCharm > 0 Or bShowSell Then
            If bShowSell Then
                nCopper = nCharmMod * nCopper
                Do While nCopper > 4294967295# 'for the overflow bug
                    nCopper = nCopper - 4294967295#
                Loop
                nCopper = Fix(nCopper / 100)
                
                'nCopper = nCopper * nCharmMod
            Else
                nCopper = (nCharmMod * nCopper)
                Do While nCopper > 4294967295# 'for the overflow bug
                    nCopper = nCopper - 4294967295#
                Loop
                'nCopper = nCopper * nCharmMod
            End If
            If nCopper <= 0 Then nCopper = 0
        End If
        
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
            oLI.ListSubItems.Add (4), "Cost", IIf(nCopper <= 0, "Free", sCopper & " Copper")
        Else
            sStr = Format(nReducedCoin, "##,##0.00")
            If Right(sStr, 3) = ".00" Then sStr = Left(sStr, Len(sStr) - 3)
            oLI.ListSubItems.Add (4), "Cost", Format(nCopper, "#,#") & " copper (" & sStr & " " & sReducedCoin & ")"
        End If
        oLI.ListSubItems(4).Tag = nCopper
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
Dim sSpellDetail As String, sRemoves As String, sArr() As String, x As Integer, y As Integer
'Dim nSpellDamage As Currency, nSpellDuration As Long, nTotalResist As Double
Dim bCalcCombat As Boolean, bUseCharacter As Boolean, nTemp As Currency, sTemp As String, sTemp2 As String
Dim bDamageMinusMR As Boolean, nCastPCT As Double, tSpellCast As tSpellCastValues, bBR As Boolean
Dim nCastLVL As Long, sSpellEQ As String ', sCastCalc As String, sMMA As String, sCastLVL As String
'Dim tSpellMinMaxDur As SpellMinMaxDur, nCastChance As Double

DetailTB.Text = ""
If bStartup Then Exit Sub

tabSpells.Index = "pkSpells"
tabSpells.Seek "=", nSpellNum
If tabSpells.NoMatch = True Then
    DetailTB.Text = "spell not found"
    tabSpells.MoveFirst
    Exit Sub
End If

If frmMain.chkSpellOptions(0).Value = 1 And Val(frmMain.txtSpellOptions(0).Text) > 0 Then bCalcCombat = True
If frmMain.chkGlobalFilter.Value = 1 And Val(frmMain.txtGlobalLevel(1).Text) > 0 Then bUseCharacter = True

LocationLV.ListItems.clear

If bUseCharacter Then
    sSpellEQ = PullSpellEQ(True, Val(frmMain.txtGlobalLevel(0).Text), , LocationLV)
Else
    sSpellEQ = PullSpellEQ(False, , , LocationLV)
End If
If Not tabSpells.Fields("Number") = nSpellNum Then tabSpells.Seek "=", nSpellNum

'If tabSpells.Fields("EnergyCost") > 0 And tabSpells.Fields("EnergyCost") <= 500 Then
'    sSpellEQ = sSpellEQ & ", x" & Fix(1000 / tabSpells.Fields("EnergyCost")) & " times/round"
'End If

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

If bUseCharacter Then nCastLVL = Val(frmMain.txtGlobalLevel(1).Text)

If bCalcCombat Then
    If bUseCharacter Then
        tSpellCast = CalculateSpellCast(nSpellNum, nCastLVL, Val(frmMain.lblCharSC.Tag), _
            Val(frmMain.txtSpellOptions(0).Text), IIf(frmMain.chkSpellOptions(2).Value = 1, True, False))
    Else
        tSpellCast = CalculateSpellCast(nSpellNum, nCastLVL, 0, _
            Val(frmMain.txtSpellOptions(0).Text), IIf(frmMain.chkSpellOptions(2).Value = 1, True, False))
    End If
Else
    If bUseCharacter Then
        tSpellCast = CalculateSpellCast(nSpellNum, nCastLVL, Val(frmMain.lblCharSC.Tag))
    Else
        tSpellCast = CalculateSpellCast(nSpellNum, nCastLVL)
    End If
End If

If bCalcCombat And (tSpellCast.bDoesDamage Or tSpellCast.bDoesHeal) And Len(tSpellCast.sAvgRound) > 0 Then
    sSpellDetail = AutoAppend(sSpellDetail, tSpellCast.sAvgRound, vbCrLf)
    bBR = True
End If

If (bCalcCombat Or Not bUseCharacter) And Len(tSpellCast.sMMA) > 0 And tSpellCast.nMinCast > 0 _
    And (tSpellCast.nMinCast <> tSpellCast.nMaxCast Or tSpellCast.nMaxCast <> tSpellCast.nAvgCast) Then
    sSpellDetail = AutoAppend(sSpellDetail, tSpellCast.sMMA, vbCrLf)
    bBR = True
End If

If bBR Then sSpellDetail = sSpellDetail & vbCrLf: bBR = False

If Len(sSpellEQ) > 0 Then sSpellDetail = AutoAppend(sSpellDetail, sSpellEQ, vbCrLf)

If bUseCharacter And Len(tSpellCast.sLVLincreases) > 0 Then
    sSpellDetail = AutoAppend(sSpellDetail, tSpellCast.sLVLincreases, vbCrLf)
End If

'=============================

'If Not bUseCharacter And nCastLVL > 0 Then sCastLVL = " (@lvl " & nCastLVL & ")"

'If bCalcCombat And (tSpellCast.bDoesDamage Or tSpellCast.bDoesHeal) Then
'
'    If tSpellCast.bDoesDamage And tSpellCast.bDoesHeal Then
'        sCastCalc = "Avg Damage+Heals/Round" & sCastLVL & ": " & IIf(tSpellCast.nDuration > 1, tSpellCast.nAvgCast, tSpellCast.nAvgRoundDmg) & " dmg + " & tSpellCast.nAvgRoundHeals & " heals"
'    ElseIf tSpellCast.bDoesDamage Then
'        sCastCalc = "Avg Damage/Round" & sCastLVL & ": " & IIf(tSpellCast.nDuration > 1, tSpellCast.nAvgCast, tSpellCast.nAvgRoundDmg)
'    ElseIf tSpellCast.bDoesHeal Then
'        sCastCalc = "Avg Healing/Round" & sCastLVL & ": " & tSpellCast.nAvgRoundHeals
'    End If
'
'    If tSpellCast.nDuration > 1 Then
'        sCastCalc = sCastCalc & " for " & tSpellCast.nDuration & " rounds (" & ((tSpellCast.nAvgRoundDmg + tSpellCast.nAvgRoundHeals) * tSpellCast.nDuration) & " total)"
'        If tSpellCast.nDamageResisted > 0 Then sCastCalc = sCastCalc & " after " & tSpellCast.nDamageResisted & "% damage resisted"
'        sTemp = ""
'        If bUseCharacter And tSpellCast.nCastChance < 100 Then sTemp = AutoAppend(sTemp, (100 - tSpellCast.nCastChance) & "% chance to fail cast", " and ")
'        If tSpellCast.nFullResistChance > 0 Then sTemp = AutoAppend(sTemp, tSpellCast.nFullResistChance & "% chance to fully-resist", " and ")
'        If Not sTemp = "" Then sCastCalc = sCastCalc & ", not including " & sTemp
'    Else
'        If bUseCharacter And tSpellCast.nCastChance < 100 Then sCastCalc = sCastCalc & " @ " & tSpellCast.nCastChance & "% chance to cast"
'        If tSpellCast.nDamageResisted > 0 Then sCastCalc = sCastCalc & ", " & tSpellCast.nDamageResisted & "% damage resisted"
'        If tSpellCast.nFullResistChance > 0 Then sCastCalc = sCastCalc & ", " & tSpellCast.nFullResistChance & "% chance to fully-resist"
'    End If
'End If

'If (bCalcCombat Or Not bUseCharacter) And tSpellCast.nMinCast > 0 And (tSpellCast.nMinCast <> tSpellCast.nMaxCast Or tSpellCast.nMaxCast <> tSpellCast.nAvgCast) Then
'    sMMA = "Min/Max/Avg Cast" & sCastLVL & ": " & tSpellCast.nMinCast & "/" & tSpellCast.nMaxCast & "/" & tSpellCast.nAvgCast
'    If tSpellCast.nNumCasts > 1 Then sMMA = sMMA & " x" & tSpellCast.nNumCasts & "/round"
'    If bCalcCombat And tSpellCast.nDuration = 1 Then
'        If tSpellCast.nFullResistChance > 0 And tSpellCast.nCastChance < 100 Then
'            sMMA = sMMA & " (before full resist & cast % reductions)"
'        ElseIf tSpellCast.nFullResistChance > 0 Then
'            sMMA = sMMA & " (before full resist reduction)"
'        ElseIf tSpellCast.nCastChance < 100 Then
'            sMMA = sMMA & " (before cast % reduction)"
'        End If
'    End If
'End If

'If Len(sCastCalc) > 0 Then sSpellDetail = AutoAppend(sSpellDetail, sCastCalc, vbCrLf)
'If Len(sMMA) > 0 Then sSpellDetail = AutoAppend(sSpellDetail, sMMA, vbCrLf)
'If Len(sCastCalc & sMMA) > 0 Then sSpellDetail = sSpellDetail & vbCrLf

'If bUseCharacter And (tabSpells.Fields("Cap") = 0 Or tabSpells.Fields("Cap") > tabSpells.Fields("ReqLevel")) _
'    And ((tabSpells.Fields("MinInc") > 0 And tabSpells.Fields("MinIncLVLs") > 0) _
'        Or (tabSpells.Fields("MaxInc") > 0 And tabSpells.Fields("MaxIncLVLs") > 0) _
'        Or (tabSpells.Fields("DurInc") > 0 And tabSpells.Fields("DurIncLVLs") > 0)) Then
'
'    sTemp = ""
'    tSpellMinMaxDur = GetCurrentSpellMinMax(False)
'    y = 0
'    For x = 0 To 9
'        If tabSpells.Fields("Abil-" & x) > 0 And tabSpells.Fields("AbilVal-" & x) = 0 Then
'            Select Case tabSpells.Fields("Abil-" & x)
'                Case 23, 51, 52, 80, 97, 98, 100, 108 To 113, 119, 138, 144:
'                        'ignore:
'                        '23 - effectsundead
'                        '51: 'anti magic
'                        '52: 'evil in combat
'                        '80: 'effects animal
'                        '97-98 - good/evil only
'                        '100: 'loyal
'                        '108: 'effects living
'                        '109 To 113: 'nonliving, notgood, notevil, neutral, not neutral
'                        '112 - neut only
'                        '119: 'del@main
'                        '138: 'roomvis
'                        '144: 'non magic spell
'                Case Else:
'                    y = y + 1
'                    sTemp = AutoAppend(sTemp, GetAbilityStats(tabSpells.Fields("Abil-" & x), 0, , False))
'            End Select
'        End If
'    Next x
'    If Not sTemp = "" Then
'        If CStr(tSpellMinMaxDur.nMin) <> tSpellMinMaxDur.sMin Then sTemp2 = AutoAppend(sTemp2, "Min: " & tSpellMinMaxDur.sMin)
'        If CStr(tSpellMinMaxDur.nMax) <> tSpellMinMaxDur.sMax Then sTemp2 = AutoAppend(sTemp2, "Max: " & tSpellMinMaxDur.sMax)
'        If CStr(tSpellMinMaxDur.nDur) <> tSpellMinMaxDur.sDur Then sTemp2 = AutoAppend(sTemp2, "Duration: " & tSpellMinMaxDur.sDur)
'        sSpellDetail = sSpellDetail & vbCrLf & "LVL Increases: " & sTemp2
'
'        If y > 1 And (CStr(tSpellMinMaxDur.nMin) <> tSpellMinMaxDur.sMin Or CStr(tSpellMinMaxDur.nMax) <> tSpellMinMaxDur.sMax) Then
'            sSpellDetail = sSpellDetail & " for: " & sTemp
'        End If
'    End If
'End If

If Not tabSpells.Fields("Number") = nSpellNum Then tabSpells.Seek "=", nSpellNum

sSpellDetail = sSpellDetail & vbCrLf & vbCrLf & "Target: " & GetSpellTargets(tabSpells.Fields("Targets"))
sSpellDetail = sSpellDetail & ", Attack Type: " & GetSpellAttackType(tabSpells.Fields("AttType"))

If nNMRVer >= 1.8 Then
    If tabSpells.Fields("TypeOfResists") = 1 Then
        sSpellDetail = sSpellDetail & ", Fully-Resistable by Anti-Magic Only"
        If tSpellCast.nFullResistChance > 0 Then sSpellDetail = sSpellDetail & " (" & tSpellCast.nFullResistChance & "%)"
    ElseIf tabSpells.Fields("TypeOfResists") = 2 Then
        sSpellDetail = sSpellDetail & ", Fully-Resistable by All"
        If tSpellCast.nFullResistChance > 0 Then sSpellDetail = sSpellDetail & " (" & tSpellCast.nFullResistChance & "%)"
    Else
        sSpellDetail = sSpellDetail & ", Can Not be Fully-Resisted"
    End If
End If

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
Exit Sub
error:
Call HandleError("PullSpellDetail")
Resume out:
End Sub

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
Public Sub AddArmour2LV(LV As ListView, Optional AddToInven As Boolean, Optional nAbility As Integer)
On Error GoTo error:
Dim oLI As ListItem, x As Integer, sName As String, nAbilityVal As Integer
Dim sAbil As String

sName = tabItems.Fields("Name")
If sName = "" Then GoTo skip:

Set oLI = LV.ListItems.Add()
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
        Case 58: ' crits
            oLI.ListSubItems(8).Text = tabItems.Fields("AbilVal-" & x)
            
        Case 135: 'min level
            oLI.ListSubItems(4).Text = tabItems.Fields("AbilVal-" & x)
        
        Case 22, 105, 106: 'acc
            oLI.ListSubItems(7).Text = Val(oLI.ListSubItems(7).Text) + tabItems.Fields("AbilVal-" & x)
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
Public Sub AddOtherItem2LV(LV As ListView)

On Error GoTo error:

Dim oLI As ListItem

If tabItems.Fields("Name") = "" Then GoTo skip:

Set oLI = LV.ListItems.Add()
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
Public Sub AddWeapon2LV(LV As ListView, Optional AddToInven As Boolean, Optional nAbility As Integer, _
    Optional ByVal nAttackType As Integer, Optional ByRef sCasts As String = "")
On Error GoTo error:
Dim oLI As ListItem, x As Integer, sName As String, nSpeed As Integer, nAbilityVal As Integer
Dim tWeaponDmg As tAttackDamage, nSpeedAdj As Integer, bUseCharacter As Boolean

sName = tabItems.Fields("Name")
If sName = "" Then GoTo skip:

If frmMain.chkWeaponOptions(3).Value = 1 Then bUseCharacter = True

nSpeedAdj = 100
If bUseCharacter Then
    If nAttackType = 0 Then nAttackType = frmMain.cmbWeaponCombos(1).ItemData(frmMain.cmbWeaponCombos(1).ListIndex)
    If frmMain.chkWeaponOptions(4).Value = 1 Then nSpeedAdj = 85
ElseIf nAttackType = 0 Then
    nAttackType = 5
End If

Set oLI = LV.ListItems.Add()
oLI.Text = tabItems.Fields("Number")

tWeaponDmg = CalculateAttack( _
    nAttackType, _
    tabItems.Fields("Number"), _
    bUseCharacter, _
    False, _
    nSpeedAdj, _
    IIf(bUseCharacter, Val(frmMain.txtWeaponExtras(2).Text), 0), _
    IIf(bUseCharacter, Val(frmMain.txtWeaponExtras(3).Text), 0), _
    IIf(bUseCharacter, Val(frmMain.txtWeaponExtras(4).Text), 0), _
    sCasts)
                
oLI.ListSubItems.Add (1), "Name", tabItems.Fields("Name")
oLI.ListSubItems.Add (2), "Wepn Type", GetWeaponType(tabItems.Fields("WeaponType"))
oLI.ListSubItems.Add (3), "Min Dmg", IIf(bUseCharacter, tWeaponDmg.nMinDmg, tabItems.Fields("Min"))
oLI.ListSubItems.Add (4), "Max Dmg", IIf(bUseCharacter, tWeaponDmg.nMaxDmg, tabItems.Fields("Max"))
oLI.ListSubItems.Add (5), "Speed", tabItems.Fields("Speed")
oLI.ListSubItems.Add (6), "Level", 0
oLI.ListSubItems.Add (7), "Str", tabItems.Fields("StrReq")
oLI.ListSubItems.Add (8), "Enc", tabItems.Fields("Encum")
oLI.ListSubItems.Add (9), "AC", RoundUp(tabItems.Fields("ArmourClass") / 10) & "/" & (tabItems.Fields("DamageResist") / 10)
oLI.ListSubItems.Add (10), "Acc", 0 'tabItems.Fields("Accy")
oLI.ListSubItems.Add (11), "BS Acc", "No"
oLI.ListSubItems.Add (12), "Crits", 0

For x = 0 To 19
    Select Case tabItems.Fields("Abil-" & x)
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

oLI.ListSubItems(10).Text = Val(oLI.ListSubItems(10).Text) + tabItems.Fields("Accy")
oLI.ListSubItems.Add (13), "Limit", tabItems.Fields("Limit")
                    
nSpeed = tabItems.Fields("Speed")
If nAttackType <> 4 And nSpeed > 0 And tWeaponDmg.nRoundTotal > 0 And tWeaponDmg.nSwings > 0 Then
    oLI.ListSubItems.Add (14), "Dmg/Spd", Round(tWeaponDmg.nRoundTotal / tWeaponDmg.nSwings / nSpeed, 4) * 1000
Else
    oLI.ListSubItems.Add (14), "Dmg/Spd", 0
End If

If nAttackType = 4 Then 'backstab
    oLI.ListSubItems.Add (15), "xSwings", tWeaponDmg.nAvgHit
Else
    oLI.ListSubItems.Add (15), "xSwings", tWeaponDmg.nRoundPhysical
End If
oLI.ListSubItems.Add (16), "Extra", Round(tWeaponDmg.nAvgExtraSwing * tWeaponDmg.nSwings)
oLI.ListSubItems.Add (17), "Dmg/Rnd", tWeaponDmg.nRoundTotal

If nAbility > 0 Then
    oLI.ListSubItems.Add (18), "Ability", nAbilityVal
ElseIf nAbility = -1 Then
    Select Case nAttackType
        Case 1: oLI.ListSubItems.Add (18), "Ability", "Punch"
        Case 2: oLI.ListSubItems.Add (18), "Ability", "Kick"
        Case 3: oLI.ListSubItems.Add (18), "Ability", "Jumpkick"
        Case 4: oLI.ListSubItems.Add (18), "Ability", "Backstab"
        Case 5: oLI.ListSubItems.Add (18), "Ability", "Normal"
        Case 6: oLI.ListSubItems.Add (18), "Ability", "Bash"
        Case 7: oLI.ListSubItems.Add (18), "Ability", "Smash"
        Case Else: oLI.ListSubItems.Add (18), "Ability", ""
    End Select
Else
    oLI.ListSubItems.Add (18), "Ability", ""
End If

If AddToInven Then Call frmMain.InvenAddEquip(tabItems.Fields("Number"), sName, tabItems.Fields("ItemType"), tabItems.Fields("Worn"))

skip:
Set oLI = Nothing

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("AddWeapon2LV")
Resume out:
End Sub

Public Function CalculateSpellCast(ByVal nSpellNum As Long, Optional ByRef nCastLVL As Long, Optional ByVal nSpellcasting As Long, _
    Optional ByVal nVSMR As Long, Optional ByVal bVSAntiMagic As Boolean) As tSpellCastValues
On Error GoTo error:
Dim x As Integer, y As Integer, tSpellMinMaxDur As SpellMinMaxDur, nDamage As Long, nHeals As Long
Dim nMinCast As Long, nMaxCast As Long, nSpellAvgCast As Long, nSpellDuration As Long, nFullResistChance As Integer
Dim nCastChance As Integer, bDamageMinusMR As Boolean, nCasts As Double, nRoundTotal As Long
Dim sAvgRound As String, bLVLspecified As Boolean, sLVLincreases As String, sMMA As String
Dim nTemp As Long, nTemp2 As Long, sTemp As String, sTemp2 As String, sCastLVL As String, sAbil As String

On Error GoTo seekit:
If tabSpells.Fields("Number") = nSpellNum Then GoTo ready:

seekit:
Resume seekit2:
seekit2:
On Error GoTo error:
tabSpells.Index = "pkSpells"
tabSpells.Seek "=", nSpellNum
If tabSpells.NoMatch = True Then
    tabSpells.MoveFirst
    Exit Function
End If

ready:
On Error GoTo error:
CalculateSpellCast.sSpellName = tabSpells.Fields("Name")

If nCastLVL <= 0 Then
    If nCastLVL < tabSpells.Fields("Cap") Then nCastLVL = tabSpells.Fields("Cap")
Else
    bLVLspecified = True
End If
If nCastLVL < tabSpells.Fields("ReqLevel") Then nCastLVL = tabSpells.Fields("ReqLevel")
If nCastLVL > tabSpells.Fields("Cap") And tabSpells.Fields("Cap") > 0 Then nCastLVL = tabSpells.Fields("Cap")

tSpellMinMaxDur = GetCurrentSpellMinMax(IIf(nCastLVL > 0, True, False), nCastLVL)

nMinCast = tSpellMinMaxDur.nMin 'GetSpellMinDamage(nSpellNum, nCastLVL)
nMaxCast = tSpellMinMaxDur.nMax 'GetSpellMaxDamage(nSpellNum, nCastLVL)
nSpellDuration = tSpellMinMaxDur.nDur 'GetSpellDuration(nSpellNum, nCastLVL)
If nSpellDuration < 1 Then nSpellDuration = 1
nSpellAvgCast = Round((nMinCast + nMaxCast) / 2)
'
'If nEnergyRem = 0 Then nEnergyRem = 1000
'nEnergyRem = nEnergyRem - tabSpells.Fields("EnergyCost")
'If nEnergyRem < 1 Then nEnergyRem = 1
'
'If nEnergyRem >= 143 And tabSpells.Fields("EnergyCost") >= 143 Then
'    If nEndCast = 0 Then
'        If tabSpells.Fields("EnergyCost") <= 500 Then
'            GetSpellMinDamage = GetSpellMinDamage + (GetSpellMinDamage * Fix(nEnergyRem / tabSpells.Fields("EnergyCost")))
'        End If
'    Else
'        GetSpellMinDamage = GetSpellMinDamage + GetSpellMinDamage(nEndCast, nCastLevel, nEnergyRem, bForMonster)
'    End If
'End If

If Not tabSpells.Fields("Number") = nSpellNum Then tabSpells.Seek "=", nSpellNum

If nSpellcasting > 0 And tabSpells.Fields("Diff") < 200 Then
    nCastChance = nSpellcasting + tabSpells.Fields("Diff")
    If nCastChance < 0 Then nCastChance = 0
    If tabSpells.Fields("Magery") = 5 Then 'kai
        If nCastChance > 100 Then nCastChance = 100
    Else
        If nCastChance > 98 Then nCastChance = 98
    End If
Else
    nCastChance = 100
End If


For x = 0 To 9
    Select Case tabSpells.Fields("Abil-" & x)
        Case 1: 'dmg
            CalculateSpellCast.bDoesDamage = True
            If tabSpells.Fields("AbilVal-" & x) <> 0 Then
                nDamage = nDamage + tabSpells.Fields("AbilVal-" & x)
            Else
                nDamage = nDamage + nSpellAvgCast
            End If
            
        Case 8: 'drain
            CalculateSpellCast.bDoesDamage = True
            CalculateSpellCast.bDoesHeal = True
            If tabSpells.Fields("AbilVal-" & x) <> 0 Then
                nDamage = tabSpells.Fields("AbilVal-" & x)
                nHeals = tabSpells.Fields("AbilVal-" & x)
            Else
                nDamage = nDamage + nSpellAvgCast
                nHeals = nHeals + nSpellAvgCast
            End If
            
        Case 17: 'dmg-mr
            CalculateSpellCast.bDoesDamage = True
            If tabSpells.Fields("AbilVal-" & x) = 0 Then
                bDamageMinusMR = True
                nTemp = nSpellAvgCast
            Else
                nTemp = tabSpells.Fields("AbilVal-" & x)
            End If
            
            If nVSMR Then
                nTemp2 = CalculateResistDamage(nTemp, nVSMR, tabSpells.Fields("TypeOfResists"), True, False, bVSAntiMagic, 0)
                CalculateSpellCast.nDamageResisted = CalculateSpellCast.nDamageResisted + (nTemp - nTemp2)
                nDamage = nDamage + nTemp2
            Else
                nDamage = nDamage + nTemp
            End If
            
        Case 18: 'healing
            CalculateSpellCast.bDoesHeal = True
            If tabSpells.Fields("AbilVal-" & x) <> 0 Then
                nHeals = tabSpells.Fields("AbilVal-" & x)
            Else
                nHeals = nHeals + nSpellAvgCast
            End If
            
    End Select
Next x

If nVSMR > 0 Then
    If bDamageMinusMR Then
        nMinCast = CalculateResistDamage(nMinCast, nVSMR, tabSpells.Fields("TypeOfResists"), bDamageMinusMR, False, bVSAntiMagic, 0)
        nMaxCast = CalculateResistDamage(nMaxCast, nVSMR, tabSpells.Fields("TypeOfResists"), bDamageMinusMR, False, bVSAntiMagic, 0)
        nSpellAvgCast = Round((nMinCast + nMaxCast) / 2)
    End If
    
    If tabSpells.Fields("TypeOfResists") = 2 Or (tabSpells.Fields("TypeOfResists") = 1 And bVSAntiMagic) Then
        nFullResistChance = Fix(nVSMR / 2)
        If nFullResistChance > 98 Then nFullResistChance = 98
    End If
End If

If tabSpells.Fields("EnergyCost") > 0 And tabSpells.Fields("EnergyCost") <= 500 Then
    nCasts = Fix(1000 / tabSpells.Fields("EnergyCost"))
Else
    nCasts = 1
End If

CalculateSpellCast.nMinCast = nMinCast 'Fix(nMinCast / nCasts)
CalculateSpellCast.nMaxCast = nMaxCast 'Fix(nMaxCast / nCasts)
CalculateSpellCast.nAvgCast = nSpellAvgCast 'Fix(nSpellAvgCast / nCasts)
CalculateSpellCast.nNumCasts = nCasts
CalculateSpellCast.nCastChance = nCastChance
CalculateSpellCast.nAvgRoundDmg = Round(((nDamage * nCasts) * (nCastChance / 100#)) * (1# - (nFullResistChance / 100#)))
CalculateSpellCast.nAvgRoundHeals = Round(((nHeals * nCasts) * (nCastChance / 100#)) * (1# - (nFullResistChance / 100#)))
CalculateSpellCast.nDuration = nSpellDuration
CalculateSpellCast.nFullResistChance = nFullResistChance

If CalculateSpellCast.nDamageResisted > 0 Then
    If nDamage = 0 Then
        CalculateSpellCast.nDamageResisted = 100
    Else
        nTemp = CalculateSpellCast.nDamageResisted
        CalculateSpellCast.nDamageResisted = Round((nTemp / (nDamage + nTemp)) * 100)
    End If
End If

'===========================

If Not bLVLspecified And nCastLVL > 0 Then sCastLVL = " (@lvl " & nCastLVL & ")"

If CalculateSpellCast.bDoesDamage Or CalculateSpellCast.bDoesHeal Then
    
    If CalculateSpellCast.bDoesDamage And CalculateSpellCast.bDoesHeal Then
        sAvgRound = sCastLVL & IIf(nSpellDuration > 1, nSpellAvgCast, CalculateSpellCast.nAvgRoundDmg) & " damage + " & CalculateSpellCast.nAvgRoundHeals & " heals/round"
    ElseIf CalculateSpellCast.bDoesDamage Then
        sAvgRound = sCastLVL & IIf(nSpellDuration > 1, nSpellAvgCast, CalculateSpellCast.nAvgRoundDmg) & " damage/round"
    ElseIf CalculateSpellCast.bDoesHeal Then
        sAvgRound = sCastLVL & CalculateSpellCast.nAvgRoundHeals & " healing/round"
    End If
    
    If nSpellDuration > 1 Then
        sAvgRound = sAvgRound & " for " & nSpellDuration & " rounds (" & ((CalculateSpellCast.nAvgRoundDmg + CalculateSpellCast.nAvgRoundHeals) * nSpellDuration) & " total)"
        If CalculateSpellCast.nDamageResisted > 0 Then sAvgRound = sAvgRound & " after " & CalculateSpellCast.nDamageResisted & "% damage resisted"
        sTemp = ""
        If bLVLspecified And nCastChance < 100 Then sTemp = AutoAppend(sTemp, (100 - nCastChance) & "% chance to fail cast", " and ")
        If CalculateSpellCast.nFullResistChance > 0 Then sTemp = AutoAppend(sTemp, CalculateSpellCast.nFullResistChance & "% chance to fully-resist", " and ")
        If Not sTemp = "" Then sAvgRound = sAvgRound & ", not including " & sTemp
    Else
        If bLVLspecified And nCastChance < 100 Then sAvgRound = sAvgRound & " @ " & nCastChance & "% chance to cast"
        If CalculateSpellCast.nDamageResisted > 0 Then sAvgRound = sAvgRound & ", " & CalculateSpellCast.nDamageResisted & "% damage resisted"
        If CalculateSpellCast.nFullResistChance > 0 Then sAvgRound = sAvgRound & ", " & CalculateSpellCast.nFullResistChance & "% chance to fully-resist"
    End If
End If

If CalculateSpellCast.nMinCast > 0 And (CalculateSpellCast.nMinCast <> CalculateSpellCast.nMaxCast Or CalculateSpellCast.nMaxCast <> nSpellAvgCast) Then
    sMMA = "Min/Max/Avg Cast" & sCastLVL & ": " & CalculateSpellCast.nMinCast & "/" & CalculateSpellCast.nMaxCast & "/" & nSpellAvgCast
    If CalculateSpellCast.nNumCasts > 1 Then sMMA = sMMA & " x" & CalculateSpellCast.nNumCasts & "/round"
    If bLVLspecified And nSpellDuration = 1 Then
        If CalculateSpellCast.nFullResistChance > 0 And nCastChance < 100 Then
            sMMA = sMMA & " (before full resist & cast % reductions)"
        ElseIf CalculateSpellCast.nFullResistChance > 0 Then
            sMMA = sMMA & " (before full resist reduction)"
        ElseIf nCastChance < 100 Then
            sMMA = sMMA & " (before cast % reduction)"
        End If
    End If
End If

If (tabSpells.Fields("Cap") = 0 Or tabSpells.Fields("Cap") > tabSpells.Fields("ReqLevel")) _
    And ((tabSpells.Fields("MinInc") > 0 And tabSpells.Fields("MinIncLVLs") > 0) _
        Or (tabSpells.Fields("MaxInc") > 0 And tabSpells.Fields("MaxIncLVLs") > 0) _
        Or (tabSpells.Fields("DurInc") > 0 And tabSpells.Fields("DurIncLVLs") > 0)) Then

    sTemp = ""
    sTemp2 = ""
    y = 0
    For x = 0 To 9
        If tabSpells.Fields("Abil-" & x) > 0 Then
            Select Case tabSpells.Fields("Abil-" & x)
                Case 23, 51, 52, 80, 97, 98, 100, 108 To 113, 115, 119, 122, 138, 144, 151, 164, 178:
                    'ignore:
                    '23 - effectsundead
                    '51: 'anti magic
                    '52: 'evil in combat
                    '80: 'effects animal
                    '97-98 - good/evil only
                    '100: 'loyal
                    '108: 'effects living
                    '109 To 113: 'nonliving, notgood, notevil, neutral, not neutral
                    '112 - neut only
                    '115 - descmsg
                    '119: 'del@main
                    '122: removespell
                    '138: 'roomvis
                    '144: 'non magic spell
                    '151,164: endcast, endcast%
                    '178: shadowform
                Case 137:
                    '137-shock... really ignore
                Case Else:
                    sAbil = GetAbilityStats(tabSpells.Fields("Abil-" & x), 0, , False)
                    If Len(sAbil) > 0 Then
                        y = y + 1
                        If tabSpells.Fields("AbilVal-" & x) = 0 Then
                            sTemp = AutoAppend(sTemp, sAbil)
                        Else
                            sTemp2 = AutoAppend(sTemp2, sAbil)
                        End If
                    End If
            End Select
        End If
    Next x
    If Not sTemp = "" Then
        If Not sTemp2 = "" Then sTemp = sTemp & " (not " & sTemp2 & ")"
        sTemp2 = ""
        tSpellMinMaxDur = GetCurrentSpellMinMax(False)
        If CStr(tSpellMinMaxDur.nDur) <> tSpellMinMaxDur.sDur Then sTemp2 = AutoAppend(sTemp2, "Duration: " & tSpellMinMaxDur.sDur)
        If CStr(tSpellMinMaxDur.nMin) <> tSpellMinMaxDur.sMin Then sTemp2 = AutoAppend(sTemp2, "Min: " & tSpellMinMaxDur.sMin)
        If CStr(tSpellMinMaxDur.nMax) <> tSpellMinMaxDur.sMax Then sTemp2 = AutoAppend(sTemp2, "Max: " & tSpellMinMaxDur.sMax)
        If Not sTemp2 = "" Then
            sLVLincreases = "LVL Increases: " & sTemp2
            If y > 1 Then sLVLincreases = sLVLincreases & " for: " & sTemp
        End If
    End If
End If

If Len(sAvgRound) > 0 Then CalculateSpellCast.sAvgRound = sAvgRound
If Len(sMMA) > 0 Then CalculateSpellCast.sMMA = sMMA
If Len(sLVLincreases) > 0 Then CalculateSpellCast.sLVLincreases = sLVLincreases

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalculateSpellCast")
Resume out:
End Function

Public Function GetDamageOutput(Optional ByVal nSingleMonster As Long, Optional ByVal sLairsSummonedBy As String, _
    Optional ByVal bAntiMagicSpecifed As Boolean, Optional ByRef bAntiMagic As Boolean, Optional ByRef nMobDodge = -1) As Currency
On Error GoTo error:
Dim x As Integer, nSpeedAdj As Integer
Dim nVSAC As Long, nVSDR As Long, nVSDodge As Long, nVSMR As Long
Dim tAttack As tAttackDamage, tSpellCast As tSpellCastValues, tAvgLairInfo As LairInfoType

If nCurrentAttackType = 0 Then 'onshot
    GetDamageOutput = 9999999
    'nDamageOutSpell = 9999999
    Exit Function
ElseIf nCurrentAttackType = 5 Then 'manual
    GetDamageOutput = nCurrentAttackManual
    'nDamageOutSpell = nCurrentAttackManualMag
    Exit Function
End If

If nSingleMonster = 0 And Len(sLairsSummonedBy) = 0 Then Exit Function
nSpeedAdj = 100

If Len(sLairsSummonedBy) > 0 Then GoTo lair:

If sCharDamageVsMonsterConfig = sCurrentAttackConfig Then
    If nCharDamageVsMonster(nSingleMonster) >= 0 Then
        GetDamageOutput = nCharDamageVsMonster(nSingleMonster)
        Exit Function
    End If
Else
    ClearSavedDamageVsMonster 'this also sets sCharDamageVsMonsterConfig = sCurrentAttackConfig
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
    Exit Function
End If

monready:
nVSAC = tabMonsters.Fields("ArmourClass")
nVSDR = tabMonsters.Fields("DamageResist")
nVSMR = tabMonsters.Fields("MagicRes")
If nMobDodge < 0 Or bAntiMagicSpecifed = False Then
    For x = 0 To 9
        If tabMonsters.Fields("Abil-" & x) = 34 Then 'dodge
            nMobDodge = tabMonsters.Fields("AbilVal-" & x)
        ElseIf tabMonsters.Fields("Abil-" & x) = 51 Then 'anti-magic
            bAntiMagic = True
        End If
    Next
    If nMobDodge < 0 Then nMobDodge = 0
End If
nVSDodge = nMobDodge
GoTo getdamage:

lair:
If nNMRVer < 1.83 Then Exit Function
If tLastAvgLairInfo.sGroupIndex <> sLairsSummonedBy Then tLastAvgLairInfo = GetAverageLairValuesFromLocs(sLairsSummonedBy)
tAvgLairInfo = tLastAvgLairInfo

nVSAC = tAvgLairInfo.nAvgAC
nVSDR = tAvgLairInfo.nAvgDR
nVSMR = tAvgLairInfo.nAvgMR
nVSDodge = tAvgLairInfo.nAvgDodge

getdamage:
Select Case nCurrentAttackType
    Case 1, 6, 7: 'eq'd weapon, bash, smash
        If nCurrentCharWeaponNumber(0) > 0 Then
        
            If nCurrentAttackType = 6 Then 'bash w/wep
                tAttack = CalculateAttack(6, nCurrentCharWeaponNumber(0), True, False, nSpeedAdj, nVSAC, nVSDR, nVSDodge)
                GetDamageOutput = tAttack.nRoundTotal
                
            ElseIf nCurrentAttackType = 7 Then 'smash w/wep
                tAttack = CalculateAttack(7, nCurrentCharWeaponNumber(0), True, False, nSpeedAdj, nVSAC, nVSDR, nVSDodge)
                GetDamageOutput = tAttack.nRoundTotal
                
            Else 'EQ'd Weapon reg attack
                tAttack = CalculateAttack(5, nCurrentCharWeaponNumber(0), True, False, nSpeedAdj, nVSAC, nVSDR, nVSDodge)
                GetDamageOutput = tAttack.nRoundTotal
            End If
            
        End If

    Case 2, 3:
        '2-spell learned: GetSpellShort(nCurrentAttackSpellNum) & " @ " & Val(txtGlobalLevel(0).Text)
        '3-spell any: GetSpellShort(nCurrentAttackSpellNum) & " @ " & nCurrentAttackSpellLVL
        If nCurrentAttackSpellNum > 0 Then
        
            If frmMain.chkGlobalFilter.Value = 1 Then
                tSpellCast = CalculateSpellCast(nCurrentAttackSpellNum, Val(frmMain.txtGlobalLevel(0).Text), Val(frmMain.lblCharSC.Tag), nVSMR, bAntiMagic)
            Else
                tSpellCast = CalculateSpellCast(nCurrentAttackSpellNum, 0, 0, nVSMR, bAntiMagic)
            End If
            GetDamageOutput = tSpellCast.nAvgRoundDmg
            
        End If

    Case 4: 'martial arts attack
        '1-Punch, 2-Kick, 3-JumpKick
        Select Case nCurrentAttackMA
            Case 2: 'kick
                tAttack = CalculateAttack(2, , True, False, nSpeedAdj, nVSAC, nVSDR, nVSDodge)
                GetDamageOutput = tAttack.nRoundTotal
                
            Case 3: 'jumpkick
                tAttack = CalculateAttack(3, , True, False, nSpeedAdj, nVSAC, nVSDR, nVSDodge)
                GetDamageOutput = tAttack.nRoundTotal
                
            Case Else: 'punch
                tAttack = CalculateAttack(1, , True, False, nSpeedAdj, nVSAC, nVSDR, nVSDodge)
                GetDamageOutput = tAttack.nRoundTotal
        End Select

End Select

If Len(sLairsSummonedBy) = 0 Then '(therefore vs monster)
    nCharDamageVsMonster(nSingleMonster) = GetDamageOutput
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetDamageOutput")
Resume out:
End Function

Public Function CalculateAttack(ByVal nAttackType As Integer, Optional ByVal nWeaponNumber As Long, Optional ByVal bUseCharacter As Boolean, _
    Optional ByVal bAbil68Slow As Boolean, Optional ByVal nSpeedAdj As Integer = 100, Optional ByVal nVSAC As Long, Optional ByVal nVSDR As Long, _
    Optional ByVal nVSDodge As Long, Optional ByRef sCasts As String = "") As tAttackDamage
On Error GoTo error:
Dim x As Integer, nAvgHit As Currency, nPlusMaxDamage As Integer, nCritChance As Integer, nAvgCrit As Long
Dim nPercent As Double, nDurDamage As Currency, nDurCount As Integer, nTemp As Integer, nPlusMinDamage As Integer
Dim tMatches() As RegexMatches, sRegexPattern As String, sAttackDetail As String
Dim sArr() As String, iMatch As Integer, nExtraTMP As Currency, nExtraAvgSwing As Currency, nCount As Integer, nExtraPCT As Double
Dim nEncum As Currency, nEnergy As Long, nCombat As Currency, nQnDBonus As Currency, nSwings As Double, nExtraAvgHit As Currency
Dim nMinCrit As Long, nMaxCrit As Long, nStrReq As Integer, nAttackAccuracy As Currency, nPercent2 As Double
Dim nDmgMin As Long, nDmgMax As Long, nAttackSpeed As Integer, nMAPlusAccy As Long, nMAPlusDmg As Long, nMAPlusSkill As Integer
Dim nLevel As Integer, nStrength As Integer, nAgility As Integer, nPlusBSaccy As Integer, nPlusBSmindmg As Integer, nPlusBSmaxdmg As Integer
Dim nStealth As Integer, bClassStealth As Boolean, nDamageBonus As Integer, nHitChance As Currency
Dim tStatIndex As TypeGetEquip

'nAttackType:
'1-punch, 2-kick, 3-jumpkick
'4-surprise, 5-normal, 6-bash, 7-smash
If nAttackType <= 0 Then nAttackType = 5
If nWeaponNumber = 0 And nAttackType > 5 Then Exit Function 'bash/smash

If bUseCharacter Then
    nLevel = Val(frmMain.txtGlobalLevel(0).Text)
    nCombat = GetClassCombat(frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex))
    nEncum = CalcEncumbrancePercent(Val(frmMain.lblInvenCharStat(0).Caption), Val(frmMain.lblInvenCharStat(1).Caption))
    nStrength = Val(frmMain.txtCharStats(0).Text)
    nAgility = Val(frmMain.txtCharStats(3).Text)
    nCritChance = Val(frmMain.lblInvenCharStat(7).Tag) - nCurrentCharQnDbonus
    nPlusMaxDamage = Val(frmMain.lblInvenCharStat(11).Tag)
    nPlusMinDamage = Val(frmMain.lblInvenCharStat(30).Tag)
    nAttackAccuracy = Val(frmMain.lblInvenCharStat(10).Tag)
    nPlusBSaccy = Val(frmMain.lblInvenCharStat(13).Tag)
    nPlusBSmindmg = Val(frmMain.lblInvenCharStat(14).Tag)
    nPlusBSmaxdmg = Val(frmMain.lblInvenCharStat(15).Tag)
    nStealth = Val(frmMain.lblInvenCharStat(19).Tag)
    If Val(frmMain.lblInvenStats(19).Tag) >= 2 Then bClassStealth = True
    
    Select Case nAttackType
        Case 1: 'Punch
            nMAPlusSkill = Val(frmMain.lblInvenCharStat(37).Tag)
            nMAPlusAccy = Val(frmMain.lblInvenCharStat(40).Tag)
            nMAPlusDmg = Val(frmMain.lblInvenCharStat(34).Tag)
        Case 2: 'Kick
            nMAPlusSkill = Val(frmMain.lblInvenCharStat(38).Tag)
            nMAPlusAccy = Val(frmMain.lblInvenCharStat(41).Tag)
            nMAPlusDmg = Val(frmMain.lblInvenCharStat(35).Tag)
        Case 3: 'Jumpkick
            nMAPlusSkill = Val(frmMain.lblInvenCharStat(39).Tag)
            nMAPlusAccy = Val(frmMain.lblInvenCharStat(42).Tag)
            nMAPlusDmg = Val(frmMain.lblInvenCharStat(36).Tag)
    End Select
    
    If (nCurrentCharWeaponNumber(0) > 0 And nWeaponNumber > 0 And nWeaponNumber <> nCurrentCharWeaponNumber(0)) Then
        'subtract current item's stats from overall stats
        nAttackAccuracy = nAttackAccuracy - nCurrentCharWeaponAccy(0)
        nCritChance = nCritChance - nCurrentCharWeaponCrit(0)
        nPlusMaxDamage = nPlusMaxDamage - nCurrentCharWeaponMaxDmg(0)
        nPlusBSaccy = nPlusBSaccy - nCurrentCharWeaponBSaccy(0)
        nPlusBSmindmg = nPlusBSmindmg - nCurrentCharWeaponBSmindmg(0)
        nPlusBSmaxdmg = nPlusBSmaxdmg - nCurrentCharWeaponBSmaxdmg(0)
        nStealth = nStealth - nCurrentCharWeaponStealth(0)
        
        Select Case nAttackType
            Case 1: 'Punch
                nMAPlusSkill = nMAPlusSkill - nCurrentCharWeaponPunchSkill(0)
                nMAPlusAccy = nMAPlusAccy - nCurrentCharWeaponPunchAccy(0)
                nMAPlusDmg = nMAPlusDmg - nCurrentCharWeaponPunchDmg(0)
            Case 2: 'Kick
                nMAPlusSkill = nMAPlusSkill - nCurrentCharWeaponKickSkill(0)
                nMAPlusAccy = nMAPlusAccy - nCurrentCharWeaponKickAccy(0)
                nMAPlusDmg = nMAPlusDmg - nCurrentCharWeaponKickDmg(0)
            Case 3: 'Jumpkick
                nMAPlusSkill = nMAPlusSkill - nCurrentCharWeaponJkSkill(0)
                nMAPlusAccy = nMAPlusAccy - nCurrentCharWeaponJkAccy(0)
                nMAPlusDmg = nMAPlusDmg - nCurrentCharWeaponJkDmg(0)
        End Select
    ElseIf nAttackType <= 3 Then
        'weapon accuracy does not count towards mystic attacks
        nAttackAccuracy = nAttackAccuracy - nCurrentCharWeaponAccy(0)
    End If
Else
    nLevel = 255
    nCombat = 3
    nStrength = 255
    nAgility = 255
    nMAPlusSkill = 1
    nAttackAccuracy = 999
    nPlusBSaccy = 999
    bClassStealth = True
End If

If nWeaponNumber = 0 Then GoTo non_weapon_attack:

On Error GoTo seek2:
If tabItems.Fields("Number") = nWeaponNumber Then GoTo item_ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabItems.Index = "pkItems"
tabItems.Seek "=", nWeaponNumber
If tabItems.NoMatch = True Then
    tabItems.MoveFirst
    Exit Function
End If

item_ready:
On Error GoTo error:

If bUseCharacter And nWeaponNumber > 0 And nWeaponNumber <> nCurrentCharWeaponNumber(0) Then
    'current weapon is different than this weapon...
    If tabItems.Fields("WeaponType") = 1 Or tabItems.Fields("WeaponType") = 3 Then
        '+this weapon is two-handed...
        If nCurrentCharWeaponNumber(1) > 0 Then
            '+off-hand currently equipped. subtract those stats too...
            nAttackAccuracy = nAttackAccuracy - nCurrentCharWeaponAccy(1)
            nCritChance = nCritChance - nCurrentCharWeaponCrit(1)
            nPlusMaxDamage = nPlusMaxDamage - nCurrentCharWeaponMaxDmg(1)
            nPlusBSaccy = nPlusBSaccy - nCurrentCharWeaponBSaccy(1)
            nPlusBSmindmg = nPlusBSmindmg - nCurrentCharWeaponBSmindmg(1)
            nPlusBSmaxdmg = nPlusBSmaxdmg - nCurrentCharWeaponBSmaxdmg(1)
            nStealth = nStealth - nCurrentCharWeaponStealth(1)
            Select Case nAttackType
                Case 1: 'Punch
                    nMAPlusSkill = nMAPlusSkill - nCurrentCharWeaponPunchSkill(1)
                    nMAPlusAccy = nMAPlusAccy - nCurrentCharWeaponPunchAccy(1)
                    nMAPlusDmg = nMAPlusDmg - nCurrentCharWeaponPunchDmg(1)
                Case 2: 'Kick
                    nMAPlusSkill = nMAPlusSkill - nCurrentCharWeaponKickSkill(1)
                    nMAPlusAccy = nMAPlusAccy - nCurrentCharWeaponKickAccy(1)
                    nMAPlusDmg = nMAPlusDmg - nCurrentCharWeaponKickDmg(1)
                Case 3: 'Jumpkick
                    nMAPlusSkill = nMAPlusSkill - nCurrentCharWeaponJkSkill(1)
                    nMAPlusAccy = nMAPlusAccy - nCurrentCharWeaponJkAccy(1)
                    nMAPlusDmg = nMAPlusDmg - nCurrentCharWeaponJkDmg(1)
            End Select
        End If
    End If
    
    'now add in current item's stats...
    If nAttackType > 3 Then
        'weapon accuracy does not count towards mystic attacks
        nAttackAccuracy = nAttackAccuracy + tabItems.Fields("Accy")
    End If
    
    For x = 0 To 19
        If tabItems.Fields("Abil-" & x) > 0 And tabItems.Fields("AbilVal-" & x) <> 0 Then
            tStatIndex = InvenGetEquipInfo(tabItems.Fields("Abil-" & x), tabItems.Fields("AbilVal-" & x))
            If tStatIndex.nEquip > 0 Then
                Select Case tStatIndex.nEquip
                    Case 7: nCritChance = nCritChance + tabItems.Fields("AbilVal-" & x)
                    Case 11: nPlusMaxDamage = nPlusMaxDamage + tabItems.Fields("AbilVal-" & x)
                    Case 13: nPlusBSaccy = nPlusBSaccy + tabItems.Fields("AbilVal-" & x)
                    Case 14: nPlusBSmindmg = nPlusBSmindmg + tabItems.Fields("AbilVal-" & x)
                    Case 15: nPlusBSmaxdmg = nPlusBSmaxdmg + tabItems.Fields("AbilVal-" & x)
                    Case 37: If nAttackType = 1 Then nMAPlusSkill = nMAPlusSkill + tabItems.Fields("AbilVal-" & x) 'pu
                    Case 40: If nAttackType = 1 Then nMAPlusAccy = nMAPlusAccy + tabItems.Fields("AbilVal-" & x) 'pu
                    Case 34: If nAttackType = 1 Then nMAPlusDmg = nMAPlusDmg + tabItems.Fields("AbilVal-" & x) 'pu
                    Case 38: If nAttackType = 2 Then nMAPlusSkill = nMAPlusSkill + tabItems.Fields("AbilVal-" & x) 'kick
                    Case 41: If nAttackType = 2 Then nMAPlusAccy = nMAPlusAccy + tabItems.Fields("AbilVal-" & x) 'kick
                    Case 35: If nAttackType = 2 Then nMAPlusDmg = nMAPlusDmg + tabItems.Fields("AbilVal-" & x) 'kick
                    Case 39: If nAttackType = 3 Then nMAPlusSkill = nMAPlusSkill + tabItems.Fields("AbilVal-" & x) 'jk
                    Case 42: If nAttackType = 3 Then nMAPlusAccy = nMAPlusAccy + tabItems.Fields("AbilVal-" & x) 'jk
                    Case 36: If nAttackType = 3 Then nMAPlusDmg = nMAPlusDmg + tabItems.Fields("AbilVal-" & x) 'jk
                    Case 19: nStealth = nStealth + tabItems.Fields("AbilVal-" & x)
                End Select
            End If
        End If
    Next x
End If

If nAttackType <= 3 Then GoTo non_weapon_attack:

CalculateAttack.sAttackDesc = tabItems.Fields("Name")
nStrReq = tabItems.Fields("StrReq")
nDmgMin = tabItems.Fields("Min")
nDmgMax = tabItems.Fields("Max")
nAttackSpeed = tabItems.Fields("Speed")
If bAbil68Slow Then nAttackSpeed = Fix((nAttackSpeed * 3) / 2)

GoTo calc_energy:

non_weapon_attack:
If nAttackType <= 3 And nMAPlusSkill <= 0 Then Exit Function
CalculateAttack.sAttackDesc = "Punch"

Select Case nAttackType
    Case 1: 'Punch
        nAttackSpeed = 1150
        If bAbil68Slow Then nAttackSpeed = 1750
    Case 2: 'Kick
        nAttackSpeed = 1400
        If bAbil68Slow Then nAttackSpeed = 2000
    Case 3: 'Jumpkick
        nAttackSpeed = 1900
        If bAbil68Slow Then nAttackSpeed = 2650
    Case 4, 5: 'surprise/normal Punch
        nAttackSpeed = 1200
        If bAbil68Slow Then nAttackSpeed = 1800
    Case Else:
        Exit Function
End Select

If nAttackType < 4 Then
    nTemp = nLevel
    If nTemp > 20 Then nTemp = 20
    
    nDmgMin = nMAPlusSkill * nTemp
    If nDmgMin < 0 Then nDmgMin = nDmgMin + 7 'it's in the dll... not sure why as this would only happen is skill was < 0, but just in case.
    nDmgMin = Fix(nDmgMin / 8) + 2

    Select Case nAttackType
        Case 1: 'Punch
            nDmgMax = nMAPlusSkill * (nTemp + 3)
            If nDmgMax < 0 Then nDmgMax = nDmgMax + 3 'it's in the dll... not sure why as this would only happen is skill was < 0, but just in case.
            nDmgMax = Fix(nDmgMax / 4) + 6
        Case 2: 'Kick
            nDmgMax = nMAPlusSkill * nTemp
            nDmgMax = Fix(nDmgMax / 6) + 7
        Case 3: 'Jumpkick
            nDmgMax = nMAPlusSkill * nTemp
            nDmgMax = Fix(nDmgMax / 6) + 8
    End Select
Else 'attacking without +punch or without a weapon
    nDmgMin = 1
    nDmgMax = 4
End If

calc_energy:
If nAttackType = 4 Or nAttackType = 7 Then 'backstab, smash
    nEnergy = 1000
    nSwings = 1
Else
    nEnergy = CalcEnergyUsed(nCombat, nLevel, nAttackSpeed, nAgility, nStrength, nEncum, nStrReq, nSpeedAdj, IIf(nAttackType = 4, True, False))
End If

If bUseCharacter Then
    If nStrength >= nStrReq Then
        nQnDBonus = CalcQuickAndDeadlyBonus(nAgility, nEnergy, nEncum)
        nCritChance = nCritChance + nQnDBonus
    End If
    If nCritChance > 40 Then nCritChance = (40 + Fix((nCritChance - 40) / 3)) 'diminishing returns
End If

If nAttackType = 6 Then nEnergy = nEnergy * 2 'bash
If nEnergy < 200 Then nEnergy = 200
If nEnergy > 1000 Then nEnergy = 1000
nSwings = Round((1000 / nEnergy), 4)
If nSwings > 5 Then nSwings = 5

nDmgMin = nDmgMin + nPlusMinDamage
nDmgMax = nDmgMax + nPlusMaxDamage
If nDmgMin > nDmgMax Then nDmgMin = nDmgMax
If nDmgMin < 0 Then nDmgMin = 0
If nDmgMax < 0 Then nDmgMax = 0


If nAttackType < 4 Then
    nAttackAccuracy = nAttackAccuracy + nMAPlusAccy
    nDmgMin = nDmgMin + nMAPlusDmg
    nDmgMax = nDmgMax + nMAPlusDmg
    If nAttackType = 2 Then 'kick
        nDamageBonus = 33
        CalculateAttack.sAttackDesc = "Kick"
    ElseIf nAttackType = 3 Then 'jk
        nDamageBonus = 66
        CalculateAttack.sAttackDesc = "JumpKick"
    End If
    
ElseIf nAttackType = 4 Then 'surprise
    If CalculateAttack.sAttackDesc = "Punch" Then
        CalculateAttack.sAttackDesc = "Surprise Punch"
    Else
        CalculateAttack.sAttackDesc = "backstab with " & CalculateAttack.sAttackDesc
    End If
    
    nCritChance = 0
    nQnDBonus = 0
    
    nTemp = (nLevel * 2) + Fix(nStealth / 10)
    nDmgMin = (nDmgMin * 2) + nTemp + nPlusBSmindmg
    nDmgMax = (nDmgMax * 2) + nTemp + nPlusBSmaxdmg
    
    If Not bClassStealth Then
        nDmgMin = Fix((nDmgMin * 75) / 100)
        nDmgMax = Fix((nDmgMax * 75) / 100)
    End If
    
    nDmgMin = Fix(((nLevel + 100) * nDmgMin) / 100)
    nDmgMax = Fix(((nLevel + 100) * nDmgMax) / 100)
    
    nAttackAccuracy = Fix((nStealth + nAgility) / 2)
    If bClassStealth Then 'class or class+race
        nAttackAccuracy = nAttackAccuracy + 5
    Else 'race only
        nAttackAccuracy = nAttackAccuracy - 15
    End If
    nAttackAccuracy = nAttackAccuracy + Fix(nPlusBSaccy / 2) + IIf(bUseCharacter, nCurrentCharAccyAbil22, 0)
    
ElseIf nAttackType = 6 Then 'bash
    nCritChance = 0
    nQnDBonus = 0
    nDamageBonus = 10
    nAttackAccuracy = nAttackAccuracy - 15
    CalculateAttack.sAttackDesc = "bash with " & CalculateAttack.sAttackDesc
ElseIf nAttackType = 7 Then 'smash
    nCritChance = 0
    nQnDBonus = 0
    nDamageBonus = 20
    nAttackAccuracy = nAttackAccuracy - 20
    CalculateAttack.sAttackDesc = "smash with " & CalculateAttack.sAttackDesc
End If

'If PARTY_FRONTRANK Then
'    if NOT ATTACK_TYPE = 3 then CHAR_ACCY += 5 'NOT jumpkick
'ElseIf PARTY_BACKRANK Then
'    if NOT ATTACK_TYPE = 3 then CHAR_ACCY -= 10 'NOT jumpkick
'End If

'ACCY_PENALTY += MONSTER_ABILITY104 'Abil 104 = DefenseModifier

nHitChance = 100
If nVSAC > 0 Then
    If nAttackType = 4 Then 'surprise
        nHitChance = nAttackAccuracy - nVSAC
    Else
        'SuccessChance = Round(1 - (((m_nUserAC * m_nUserAC) / 100) / ((nAttack_AdjSuccessChance * nAttack_AdjSuccessChance) / 140)), 2) * 100
        nHitChance = Round(1 - (((nVSAC * nVSAC) / 100) / ((nAttackAccuracy * nAttackAccuracy) / 140)), 2) * 100
    End If
    If nHitChance < 9 Then nHitChance = 9
    If nHitChance > 99 Then nHitChance = 99
End If
If nVSDodge < 0 And nVSAC > 0 Then
    'the dll provides for a x% chance for AC to be ignored if dodge is negative
    'i.e.: (-dodge+100) = chance to ignore AC check and have a 99% hit chance
    'so, if dodge was -10, there would be a 10% chance to ignore AC
    'i'm simulating this by just reducing the hitchance at scale.
    nPercent = ((nVSDodge + 100) / 100)  'chance for 99% hit
    nPercent2 = 1 - nPercent 'chance for regular hit chance
    nHitChance = (99 * nPercent) + (nHitChance * nPercent2)
    If nHitChance < 9 Then nHitChance = 9
ElseIf nVSDodge > 0 And nAttackAccuracy > 0 Then
    nPercent = Fix((nVSDodge * 10) / Fix(nAttackAccuracy / 8))
    If nPercent > 95 Then nPercent = 95
    If nAttackType = 4 Then nPercent = Fix(nPercent / 5) 'backstab
    CalculateAttack.nDodgeChance = nPercent
    nPercent = (nPercent / 100) 'chance to dodge
    nHitChance = (nHitChance * (1 - nPercent))
    If nHitChance < 9 Then nHitChance = 9
End If
nHitChance = nHitChance / 100

If nDamageBonus > 0 Then
    nDmgMin = Fix((nDmgMin * (100 + nDamageBonus)) / 100)
    nDmgMax = Fix((nDmgMax * (100 + nDamageBonus)) / 100)
End If

nDmgMin = nDmgMin - nVSDR
nDmgMax = nDmgMax - nVSDR
If nDmgMin < 0 Then nDmgMin = 0
If nDmgMax < 0 Then nDmgMax = 0

If nCritChance > 0 Then
    nMinCrit = nDmgMax * 2
    nMaxCrit = nDmgMax * 4
    If nMinCrit > nMaxCrit Then nMaxCrit = nMinCrit
    nAvgCrit = Round((nMinCrit + nMaxCrit + 1) / 2) - nVSDR
End If

If nAttackType = 6 Then 'bash
    'nAvgHit = nAvgHit * 3
    nDmgMin = nDmgMin * 3
    nDmgMax = nDmgMax * 3
ElseIf nAttackType = 7 Then 'smash
    'nAvgHit = nAvgHit * 5
    nDmgMin = nDmgMin * 5
    nDmgMax = nDmgMax * 5
End If
nAvgHit = Round((nDmgMin + nDmgMax + 1) / 2) ' - nVSDR

If Len(sCasts) = 0 And nWeaponNumber > 0 And nAttackType > 3 Then
    For x = 0 To 19
        Select Case tabItems.Fields("Abil-" & x)
            Case 43: 'casts spell
                sCasts = AutoAppend(sCasts, "[" & GetSpellName(tabItems.Fields("AbilVal-" & x), bHideRecordNumbers) _
                    & ", " & PullSpellEQ(True, 0, tabItems.Fields("AbilVal-" & x), , , , True), "|")
                If Not nPercent = 0 Then
                    sCasts = sCasts & ", " & nPercent & "%]"
                Else
                    sCasts = sCasts & "]"
                End If
                
            Case 114: '%spell
                nPercent = tabItems.Fields("AbilVal-" & x)
              
        End Select
    Next x
End If

If Len(sCasts) > 0 And nWeaponNumber > 0 And nAttackType > 3 Then
    'this is matching against:
    '[fire burns(979), Damage 5 to 15, 100%] -- would produce 1 full match and 3 submatches for the 5, 15, and 100
    'or: [lacerate(985), Damage 3 to 12, AffectsLivingOnly, for 10 rounds, 100%] -- this will produce the same matching as above with the 10 rounds being ignored
    'or: [fire burns(979), Damage 5 to 15, 25%], [ice freezes(978), Damage 5 to 15, 25%] -- would produce 2 full matches, each with 3 submatches
    'or: [{rocks shred(977), Damage 5 to 15} OR {ice freezes(978), Damage 5 to 15} OR {fire burns(979), Damage 5 to 15} OR {acid sears(980), Damage 5 to 15} OR {lightning shocks(981), Damage 5 to 15}], 100%]
    '     ...which would produce only 1 full match with all of the damage numbers and final percentage as submatches
    sRegexPattern = "(?:(?:Damage(?:\(-MR\))?|DrainLife) (-?\d+) to (-?\d+)[^\]]*, (\d+)%|\[(?:{[^\[\{\}\]]+, (?:Damage(?:\(-MR\))?|DrainLife) (-?\d+) to (-?\d+)[^\]\}]*(?:} OR ))(?:{[^\[\{\}\]]+, (?:Damage(?:\(-MR\))?|DrainLife) (-?\d+) to (-?\d+)[^\]\}]*(?:} OR )?)?(?:{[^\[\{\}\]]+, (?:Damage(?:\(-MR\))?|DrainLife) (-?\d+) to (-?\d+)[^\]\}]*(?:} OR )?)?(?:{[^\[\{\}\]]+, (?:Damage(?:\(-MR\))?|DrainLife) (-?\d+) to (-?\d+)[^\]\}]*(?:} OR )?)?(?:{[^\[\{\}\]]+, (?:Damage(?:\(-MR\))?|DrainLife) (-?\d+) to (-?\d+)[^\]\}]*(?:} OR )?)?(?:{[^\[\{\}\]]+, (?:Damage(?:\(-MR\))?|DrainLife) (-?\d+) to (-?\d+)[^\]\}]*(?:} OR )?)?}], (\d+)%)"
    tMatches() = RegExpFindv2(sCasts, sRegexPattern, False, False, False)
    If UBound(tMatches()) = 0 And Len(tMatches(0).sFullMatch) = 0 Then GoTo done_extra:
       
'    If tabItems.Fields("Number") = 2577 Then
'        Debug.Print tabItems.Fields("Number")
'    End If
    
    For iMatch = 0 To UBound(tMatches())
        If UBound(tMatches(iMatch).sSubMatches()) < 2 Then GoTo skip_match:
        
        If InStr(1, tMatches(iMatch).sFullMatch, "} or {", vbTextCompare) > 0 Then
            'multiple spells with equal chance
            sArr() = Split(tMatches(iMatch).sFullMatch, "} or {", , vbTextCompare)
        Else
            ReDim sArr(0)
            sArr(0) = tMatches(iMatch).sFullMatch
        End If
        
        nExtraTMP = 0
        nCount = 0
        nDurDamage = 0
        nDurCount = 0
        For x = 0 To UBound(tMatches(iMatch).sSubMatches()) - 1
            nTemp = x - Fix((x + 1) / 2) 'index that refers to the full text string of the match for these two damage values
            If UBound(sArr()) >= nTemp Then
                If InStr(1, sArr(nTemp), ", for", vbTextCompare) > 0 And InStr(1, sArr(nTemp), "rounds", vbTextCompare) > 0 Then
                    nDurDamage = nDurDamage + Abs(Val(tMatches(iMatch).sSubMatches(x)))
                    nDurCount = nDurCount + 1
                    nCount = nCount + 1
                    x = x + 1 'get the next number
                    If UBound(tMatches(iMatch).sSubMatches()) >= (x + 1) Then 'plus another because there should also be the percentage at the end
                        nDurDamage = nDurDamage + Abs(Val(tMatches(iMatch).sSubMatches(x)))
                        nDurCount = nDurCount + 1
                        nCount = nCount + 1 'still counting here because its presence would reduce the chance of casting the other spells in the group, thereby reducing their overall effect on the average damage
                    End If
                    GoTo skip_submatch:
                End If
            End If
            nExtraTMP = nExtraTMP + Abs(Val(tMatches(iMatch).sSubMatches(x)))
            nCount = nCount + 1
skip_submatch:
        Next x
        
        If nCount > 0 Then nExtraTMP = Round(nExtraTMP / nCount, 2)
        nExtraAvgHit = nExtraAvgHit + nExtraTMP
        nExtraPCT = Round(Val(tMatches(iMatch).sSubMatches(UBound(tMatches(iMatch).sSubMatches()))) / 100, 2)
        nExtraTMP = Round(nExtraTMP * nExtraPCT, 2)
        
        'dividing durection by SWINGS so it actually counts only once when it multiplies by SWINGS later (e.g. we're adding one tick of the duration damage to the total per-round damage)
        If nDurCount > 0 Then nExtraTMP = nExtraTMP + Round(((nDurDamage / nDurCount) * nExtraPCT) / nSwings, 2)
        
        nExtraAvgSwing = nExtraAvgSwing + nExtraTMP
skip_match:
    Next iMatch
    
    If UBound(tMatches()) > 0 Then nExtraAvgHit = Round(nExtraAvgHit / (UBound(tMatches()) + 1))
    nExtraAvgSwing = Round(nExtraAvgSwing)
End If
done_extra:

CalculateAttack.nMinDmg = nDmgMin
CalculateAttack.nMaxDmg = nDmgMax
CalculateAttack.nAvgHit = nAvgHit
CalculateAttack.nAvgCrit = nAvgCrit
CalculateAttack.nAvgExtraHit = nExtraAvgHit
CalculateAttack.nAvgExtraSwing = nExtraAvgSwing
CalculateAttack.nCritChance = nCritChance
CalculateAttack.nQnDBonus = nQnDBonus
CalculateAttack.nSwings = nSwings
CalculateAttack.nAccy = nAttackAccuracy

nPercent = (nCritChance / 100) 'chance to crit
CalculateAttack.nRoundPhysical = (((1 - nPercent) * nAvgHit) + (nPercent * nAvgCrit)) * nSwings * nHitChance
CalculateAttack.nRoundTotal = CalculateAttack.nRoundPhysical + (nExtraAvgSwing * nSwings * nHitChance)
CalculateAttack.nHitChance = Round(nHitChance * 100)

If nSwings > 0 And (nAvgHit + nAvgCrit) > 0 Then
    sAttackDetail = "Swings: " & Round(nSwings, 1) & ", Avg Hit: " & nAvgHit
    If nAvgCrit > 0 Then
        sAttackDetail = AutoAppend(sAttackDetail, "Avg Crit: " & nAvgCrit)
        If nCritChance > 0 Then sAttackDetail = sAttackDetail & " (" & nCritChance & "%)"
    End If
    If CalculateAttack.nHitChance > 0 Then sAttackDetail = AutoAppend(sAttackDetail, "Hit: " & CalculateAttack.nHitChance & "%")
End If
CalculateAttack.sAttackDetail = sAttackDetail

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalculateAttack")
Resume out:
End Function

Public Sub AddSpell2LV(LV As ListView, Optional ByVal AddBless As Boolean)
On Error GoTo error:
Dim oLI As ListItem, sName As String, x As Integer, nSpell As Long, sTimesCast As String
Dim nSpellDamage As Currency, nSpellDuration As Long, nTemp As Long, bUseCharacter As Boolean
Dim bCalcCombat As Boolean, bDamageMinusMR As Boolean, nCastPCT As Double, tSpellCast As tSpellCastValues

If frmMain.chkSpellOptions(0).Value = 1 And Val(frmMain.txtSpellOptions(0).Text) > 0 Then bCalcCombat = True
If frmMain.chkGlobalFilter.Value = 1 And Val(frmMain.txtGlobalLevel(1).Text) > 0 Then bUseCharacter = True

nSpell = tabSpells.Fields("Number")
sName = tabSpells.Fields("Name")
If sName = "" Then GoTo skip:
If Left(sName, 1) = "1" Then GoTo skip:
If Left(LCase(sName), 3) = "sdf" Then GoTo skip:

Set oLI = LV.ListItems.Add()
oLI.Text = nSpell

oLI.ListSubItems.Add (1), "Name", sName
oLI.ListSubItems.Add (2), "Short", tabSpells.Fields("Short")
oLI.ListSubItems.Add (3), "Magery", GetMagery(tabSpells.Fields("Magery"), tabSpells.Fields("MageryLVL"))
oLI.ListSubItems.Add (4), "Level", tabSpells.Fields("ReqLevel")
oLI.ListSubItems.Add (5), "Mana", tabSpells.Fields("ManaCost")

If bUseCharacter Then
    If bCalcCombat Then
        tSpellCast = CalculateSpellCast(nSpell, Val(frmMain.txtGlobalLevel(1).Text), Val(frmMain.lblCharSC.Tag), _
            Val(frmMain.txtSpellOptions(0).Text), IIf(frmMain.chkSpellOptions(2).Value = 1, True, False))
    Else
        tSpellCast = CalculateSpellCast(nSpell, Val(frmMain.txtGlobalLevel(1).Text), Val(frmMain.lblCharSC.Tag))
    End If
Else
    tSpellCast = CalculateSpellCast(nSpell, tabSpells.Fields("ReqLevel"))
End If

If Not tabSpells.Fields("Number") = nSpell Then
    tabSpells.Index = "pkSpells"
    tabSpells.Seek "=", nSpell
    If tabSpells.NoMatch = True Then
        tabSpells.MoveFirst
        Exit Sub
    End If
End If

nSpellDuration = tSpellCast.nDuration
nCastPCT = tSpellCast.nCastChance / 100

If bUseCharacter Then
    oLI.ListSubItems.Add (6), "Diff", tSpellCast.nCastChance & "%"
Else
    oLI.ListSubItems.Add (6), "Diff", tabSpells.Fields("Diff")
End If

'nCastPCT = 1
'If bUseCharacter Then
'    If tabSpells.Fields("Diff") >= 200 Then
'        nTemp = 100
'    Else
'        nTemp = Val(frmMain.lblCharSC.Tag) + tabSpells.Fields("Diff")
'        If nTemp < 0 Then nTemp = 0
'        If nTemp > 98 Then nTemp = 98
'        nCastPCT = nTemp / 100
'    End If
'    oLI.ListSubItems.Add (6), "Diff", nTemp & "%"
'Else
'    oLI.ListSubItems.Add (6), "Diff", tabSpells.Fields("Diff")
'End If

If tabSpells.Fields("Learnable") = 1 Or tabSpells.Fields("ManaCost") > 0 Then
    
    'DMG
    
'    If bUseCharacter Then
'        nSpellDamage = GetSpellMinDamage(nSpell, Val(frmMain.txtGlobalLevel(1).Text))
'        nSpellDamage = nSpellDamage + GetSpellMaxDamage(nSpell, Val(frmMain.txtGlobalLevel(1).Text))
'        nSpellDuration = GetSpellDuration(nSpell, Val(frmMain.txtGlobalLevel(1).Text))
'    Else
'        nSpellDamage = GetSpellMinDamage(nSpell)
'        nSpellDamage = nSpellDamage + GetSpellMaxDamage(nSpell)
'        nSpellDuration = GetSpellDuration(nSpell)
'    End If
    
'    If Not tabSpells.Fields("Number") = nSpell Then
'        tabSpells.Index = "pkSpells"
'        tabSpells.Seek "=", nSpell
'        If tabSpells.NoMatch = True Then Exit Sub
'    End If
    
'    If nSpellDuration < 1 Then nSpellDuration = 1
'    nSpellDamage = Round((nSpellDamage / 2) * nSpellDuration)
    
'    If nSpellDamage > 0 And bCalcCombat Then
'        For x = 0 To 9
'            If tabSpells.Fields("Abil-" & x) = 17 Then 'Damage-MR
'                bDamageMinusMR = True
'                Exit For
'            End If
'        Next x
'
'        nSpellDamage = CalculateResistDamage(nSpellDamage, Val(frmMain.txtSpellOptions(0).Text), _
'            tabSpells.Fields("TypeOfResists"), bDamageMinusMR, True, False, 0)
'        nSpellDamage = Round(nSpellDamage * nCastPCT)
'    End If
    oLI.ListSubItems.Add (7), "Dmg", (tSpellCast.nAvgRoundDmg * tSpellCast.nDuration) 'Round(nSpellDamage)
    
    nSpellDamage = 0
    If tSpellCast.nAvgRoundDmg > 0 Then
        If tabSpells.Fields("ManaCost") > 0 Then
            If tSpellCast.nNumCasts > 1 Then
                nSpellDamage = Round(tSpellCast.nAvgRoundDmg / (tabSpells.Fields("ManaCost") * tSpellCast.nNumCasts), 1)
            Else
                nSpellDamage = Round((tSpellCast.nAvgRoundDmg * tSpellCast.nDuration) / tabSpells.Fields("ManaCost"), 1)
            End If
        End If
    End If
    
    oLI.ListSubItems.Add (8), "Dmg/M", nSpellDamage
    
    
    'HEALING
'    If bUseCharacter Then
'        nSpellDamage = GetSpellMinDamage(nSpell, Val(frmMain.txtGlobalLevel(1).Text), , , True)
'        nSpellDamage = nSpellDamage + GetSpellMaxDamage(nSpell, Val(frmMain.txtGlobalLevel(1).Text), , , True)
'        nSpellDuration = GetSpellDuration(nSpell, Val(frmMain.txtGlobalLevel(1).Text))
'    Else
'        nSpellDamage = GetSpellMinDamage(nSpell, , , , True)
'        nSpellDamage = nSpellDamage + GetSpellMaxDamage(nSpell, , , , True)
'        nSpellDuration = GetSpellDuration(nSpell)
'    End If
'
'    If Not tabSpells.Fields("Number") = nSpell Then
'        tabSpells.Index = "pkSpells"
'        tabSpells.Seek "=", nSpell
'        If tabSpells.NoMatch = True Then Exit Sub
'    End If

'    If nSpellDuration < 1 Then nSpellDuration = 1
'    nSpellDamage = (nSpellDamage / 2) * nSpellDuration

    oLI.ListSubItems.Add (9), "Heal", (tSpellCast.nAvgRoundHeals * tSpellCast.nDuration) 'Round(nSpellDamage)
    
    nSpellDamage = 0
    If tSpellCast.nAvgRoundHeals <> 0 Then
        If tabSpells.Fields("ManaCost") > 0 Then
            If tSpellCast.nNumCasts > 1 Then
                nSpellDamage = Round(tSpellCast.nAvgRoundHeals / (tabSpells.Fields("ManaCost") * tSpellCast.nNumCasts), 1)
            Else
                nSpellDamage = Round((tSpellCast.nAvgRoundHeals * tSpellCast.nDuration) / tabSpells.Fields("ManaCost"), 1)
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

'If tabSpells.Fields("EnergyCost") > 0 And tabSpells.Fields("EnergyCost") <= 500 Then
'    sTimesCast = ", x" & Fix(1000 / tabSpells.Fields("EnergyCost")) & " times/round"
'End If

bQuickSpell = True
If LV.name = "lvSpellBook" And FormIsLoaded("frmSpellBook") And bUseCharacter Then
    If Val(frmSpellBook.txtLevel) > 0 Then
        oLI.ListSubItems.Add (11), "Detail", PullSpellEQ(True, Val(frmSpellBook.txtLevel), nSpell, Nothing, , , , , True) & sTimesCast
    Else
        oLI.ListSubItems.Add (11), "Detail", PullSpellEQ(False, , nSpell, Nothing, , , , , True) & sTimesCast
    End If
Else
    If bUseCharacter Then
        oLI.ListSubItems.Add (11), "Detail", PullSpellEQ(True, Val(frmMain.txtGlobalLevel(1).Text), nSpell, Nothing, , , , , True) & sTimesCast
    Else
        oLI.ListSubItems.Add (11), "Detail", PullSpellEQ(False, , nSpell, Nothing, , , , , True) & sTimesCast
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
Exit Sub
error:
Call HandleError("AddSpell2LV")
Resume out:
End Sub

Public Sub AddRace2LV(LV As ListView)

On Error GoTo error:

Dim oLI As ListItem, x As Integer, sAbil As String
    
    If tabRaces.Fields("Name") = "" Then GoTo skip:
    
    Set oLI = LV.ListItems.Add()
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

Public Function IsMobKillable(ByVal nCharDMG As Double, ByVal nCharHP As Long, ByVal nMobDMG As Double, ByVal nMobHP As Long, Optional ByVal nCharHPRegen As Integer = 0, Optional ByVal nMobHPRegen As Long = 0) As Boolean
On Error GoTo error:
Dim nFactor As Double, nRoundsToKill As Double, nRoundsToDeath As Double
Dim nMobTotalHP As Long, nCharTotalHP As Long

If nCharDMG <= 0 Or nMobHP <= 0 Then Exit Function

nFactor = 0.25

nCharDMG = nCharDMG * (nFactor + 1)
If nCharDMG <= 0 Then Exit Function

nMobDMG = nMobDMG * (1 - nFactor)
If nMobDMG <= 0 Then
    IsMobKillable = True
    Exit Function
End If

If nMobHPRegen > 0 Then
    nMobTotalHP = nMobHP + ((nMobHP / nCharDMG) * (nMobHPRegen / 6))
Else
    nMobTotalHP = nMobHP
End If
nRoundsToKill = nMobTotalHP / nCharDMG
If nRoundsToKill < 1 Then nRoundsToKill = 1

If nCharHPRegen > 0 Then
    nCharTotalHP = nCharHP + ((nCharHP / nMobDMG) * (nCharHPRegen / 3 / 6))
Else
    nCharTotalHP = nCharHP
End If
nRoundsToDeath = nCharTotalHP / nMobDMG

If nRoundsToDeath >= nRoundsToKill Then IsMobKillable = True

out:
On Error Resume Next
Exit Function
error:
Call HandleError("IsMobKillable")
Resume out:
End Function

Public Sub AddMonster2LV(LV As ListView)
On Error GoTo error:
Dim oLI As ListItem, sName As String, nExp As Currency, nHP As Currency, x As Integer
Dim nAvgDmg As Long, nExpDmgHP As Currency, nIndex As Integer, nMagicLVL As Integer
Dim nScriptValue As Currency, nLairPCT As Currency, nPossSpawns As Long, sPossSpawns As String
Dim nMaxLairsBeforeRegen As Currency, nPossyPCT As Currency, bAsterisks As Boolean, sTemp As String
Dim tAvgLairInfo As LairInfoType, nParty As Integer, nRestingRate As Double
Dim nCharHealth As Long, nDamageOut As Long, nHPRegen As Long, nMonsterNum As Long
Dim nMobExpPerHour() As Currency, nMobDodge As Integer
Dim tAttack As tAttackDamage, tSpellCast As tSpellCastValues, bHasAntiMagic As Boolean

nMonsterNum = tabMonsters.Fields("Number")

If nNMRVer >= 1.83 And LV.hWnd = frmMain.lvMonsters.hWnd And frmMain.optMonsterFilter(1).Value = True And tLastAvgLairInfo.sGroupIndex <> tabMonsters.Fields("Summoned By") Then
    tLastAvgLairInfo = GetAverageLairValuesFromLocs(tabMonsters.Fields("Summoned By"))
ElseIf (nNMRVer < 1.83 Or LV.hWnd <> frmMain.lvMonsters.hWnd Or frmMain.optMonsterFilter(1).Value = False) And Not tLastAvgLairInfo.sGroupIndex = "" Then
    tLastAvgLairInfo = GetLairInfo("") 'reset
End If

tAvgLairInfo = tLastAvgLairInfo
If Not tabMonsters.Fields("Number") = nMonsterNum Then tabMonsters.Seek "=", nMonsterNum

sName = tabMonsters.Fields("Name")
If sName = "" Or Left(sName, 3) = "sdf" Then GoTo skip:

For x = 0 To 9 'abilities
    If Not tabMonsters.Fields("Abil-" & x) = 0 Then
        Select Case tabMonsters.Fields("Abil-" & x)
            Case 28: 'magical
                nMagicLVL = tabMonsters.Fields("AbilVal-" & x)
            Case 34: 'dodge
                nMobDodge = tabMonsters.Fields("AbilVal-" & x)
            Case 51: 'anti-magic
                bHasAntiMagic = True
        End Select
    End If
Next

Set oLI = LV.ListItems.Add()
oLI.Text = tabMonsters.Fields("Number")

nIndex = 1
oLI.ListSubItems.Add (nIndex), "Name", sName

nIndex = nIndex + 1
oLI.ListSubItems.Add (nIndex), "RGN", tabMonsters.Fields("RegenTime")
oLI.ListSubItems(nIndex).Tag = tabMonsters.Fields("RegenTime")

nIndex = nIndex + 1
If UseExpMulti Then
    nExp = tabMonsters.Fields("EXP") * tabMonsters.Fields("ExpMulti")
Else
    nExp = tabMonsters.Fields("EXP")
End If
oLI.ListSubItems.Add (nIndex), "Exp", IIf(nExp > 0, Format(nExp, "#,#"), 0)
oLI.ListSubItems(nIndex).Tag = nExp

sTemp = ""
If tAvgLairInfo.nMobs > 0 Then
    nHP = tAvgLairInfo.nAvgHP
    sTemp = "*"
Else
    nHP = tabMonsters.Fields("HP")
End If
nIndex = nIndex + 1
oLI.ListSubItems.Add (nIndex), "HP", IIf(nHP > 0, Format(nHP, "#,#"), 0) & sTemp
oLI.ListSubItems(nIndex).Tag = nHP

nIndex = nIndex + 1
oLI.ListSubItems.Add (nIndex), "AC/DR", tabMonsters.Fields("ArmourClass") & "/" & tabMonsters.Fields("DamageResist")
oLI.ListSubItems(nIndex).Tag = tabMonsters.Fields("ArmourClass") + tabMonsters.Fields("DamageResist")

nIndex = nIndex + 1
oLI.ListSubItems.Add (nIndex), "Dodge", nMobDodge
oLI.ListSubItems(nIndex).Tag = nMobDodge

nIndex = nIndex + 1
oLI.ListSubItems.Add (nIndex), "MR", tabMonsters.Fields("MagicRes")
oLI.ListSubItems(nIndex).Tag = tabMonsters.Fields("MagicRes")

'a lot of this repeated in pullmonsterdetail
nIndex = nIndex + 1
nAvgDmg = -1
sTemp = ""
If tAvgLairInfo.nMobs > 0 And tabMonsters.Fields("RegenTime") = 0 Then
    nAvgDmg = tAvgLairInfo.nAvgDmg
    sTemp = "*"
    bAsterisks = True
Else
    If frmMain.optMonsterFilter(1).Value = True And Val(frmMain.txtMonsterLairFilter(0).Text) > 1 And nMonsterDamageVsParty(tabMonsters.Fields("Number")) >= 0 Then 'vs party
        nAvgDmg = nMonsterDamageVsParty(tabMonsters.Fields("Number"))
    ElseIf frmMain.chkGlobalFilter.Value = 1 And nMonsterDamageVsChar(tabMonsters.Fields("Number")) >= 0 Then
        nAvgDmg = nMonsterDamageVsChar(tabMonsters.Fields("Number"))
    ElseIf nNMRVer >= 1.8 Then
        nAvgDmg = tabMonsters.Fields("AvgDmg")
    ElseIf nMonsterDamageVsDefault(tabMonsters.Fields("Number")) >= 0 Then
        nAvgDmg = nMonsterDamageVsDefault(tabMonsters.Fields("Number"))
    Else
        bAsterisks = True
    End If
End If
oLI.ListSubItems.Add (nIndex), "Damage", IIf(nAvgDmg > 0, Format(nAvgDmg, "#,##"), IIf(nAvgDmg = 0, 0, "?")) & sTemp
oLI.ListSubItems(nIndex).Tag = nAvgDmg
If frmMain.optMonsterFilter(1).Value = True And Val(frmMain.txtMonsterLairFilter(0).Text) > 1 Then 'vs party
    If nMonsterDamageVsParty(tabMonsters.Fields("Number")) >= 0 Then oLI.ListSubItems(nIndex).ForeColor = RGB(193, 0, 232)
ElseIf frmMain.chkGlobalFilter.Value = 1 And nMonsterDamageVsChar(tabMonsters.Fields("Number")) >= 0 Then
    oLI.ListSubItems(nIndex).ForeColor = RGB(193, 0, 232)
    'oLI.ListSubItems(nIndex).Bold = True
End If

'If tabMonsters.Fields("Number") = 511 Then
'    Debug.Print tabMonsters.Fields("Number")
'End If

'a lot of this repeated in filtermonsters
nIndex = nIndex + 1
nExpDmgHP = 0
If nNMRVer >= 1.83 And frmMain.optMonsterFilter(1).Value = True And LV.hWnd = frmMain.lvMonsters.hWnd Then 'by lair (and exp by hour)
    
    bAsterisks = False
    nCharHealth = 1
    nHPRegen = 0
    nParty = 1
    nRestingRate = 0
    
    If frmMain.chkGlobalFilter.Value = 1 And Val(frmMain.txtMonsterLairFilter(0).Text) < 2 Then 'no party, vs char
        nCharHealth = Val(frmMain.lblCharMaxHP.Tag)
        nHPRegen = Val(frmMain.lblCharRestRate.Tag)
        
    ElseIf Val(frmMain.txtMonsterLairFilter(0).Text) > 1 Then 'vs party
        nParty = Val(frmMain.txtMonsterLairFilter(0).Text)
        nCharHealth = Val(frmMain.txtMonsterLairFilter(5).Text)
        If nCharHealth < 1 Then
            frmMain.txtMonsterLairFilter(7).Text = 1
            nCharHealth = 1
        End If
        nCharHealth = nCharHealth * Val(frmMain.txtMonsterLairFilter(0).Text) 'note: nCharHealth is avg * party to match tLastAvgLairInfo values
        nHPRegen = Val(frmMain.txtMonsterLairFilter(7).Text)
        
    Else
        nCharHealth = nAvgDmg * 2
        nHPRegen = nCharHealth * 0.05
    End If
    
    If nCharHealth < 1 Then nCharHealth = 1
    If nHPRegen < 1 Then nHPRegen = 1
    If nParty > 6 Then nParty = 6
    If nParty < 1 Then nParty = 1
    
    If nParty > 1 Then
        nDamageOut = Val(frmMain.txtMonsterDamageOUT.Text) * nParty
        If nDamageOut < 0 Then nDamageOut = 0
    End If
    
    If tabMonsters.Fields("RegenTime") = 0 And tLastAvgLairInfo.nMobs > 0 Then
        nExpDmgHP = tLastAvgLairInfo.nAvgExp
        nRestingRate = tLastAvgLairInfo.nRestRate
        If nParty < 2 Then nDamageOut = tLastAvgLairInfo.nDamageOut
        'can we kill the mobs?
        
        If IsMobKillable(nDamageOut, nCharHealth, nAvgDmg, tLastAvgLairInfo.nAvgHP, nHPRegen) = False Then
            nExpDmgHP = 0
            nRestingRate = 1
        End If
        
    ElseIf tabMonsters.Fields("RegenTime") > 0 Or InStr(1, tabMonsters.Fields("Summoned By"), "Room", vbTextCompare) > 0 Then
        
        If nParty < 2 Then
            Select Case nCurrentAttackType
                Case 1, 6, 7: 'eq'd weapon, bash, smash
                    If nCurrentCharWeaponNumber(0) > 0 Then
                        If nCurrentAttackType = 6 Then 'bash w/wep
                            tAttack = CalculateAttack(6, nCurrentCharWeaponNumber(0), True, False, 100, tabMonsters.Fields("ArmourClass"), tabMonsters.Fields("DamageResist"), nMobDodge)
                            nDamageOut = tAttack.nRoundTotal
                        ElseIf nCurrentAttackType = 7 Then 'smash w/wep
                            tAttack = CalculateAttack(7, nCurrentCharWeaponNumber(0), True, False, 100, tabMonsters.Fields("ArmourClass"), tabMonsters.Fields("DamageResist"), nMobDodge)
                            nDamageOut = tAttack.nRoundTotal
                        Else 'EQ'd Weapon reg attack
                            tAttack = CalculateAttack(5, nCurrentCharWeaponNumber(0), True, False, 100, tabMonsters.Fields("ArmourClass"), tabMonsters.Fields("DamageResist"), nMobDodge)
                            nDamageOut = tAttack.nRoundTotal
                        End If
                    Else
                        GoTo no_attack:
                    End If
                    
                Case 2, 3:
                    '2-spell learned: GetSpellShort(nCurrentAttackSpellNum) & " @ " & Val(txtGlobalLevel(0).Text)
                    '3-spell any: GetSpellShort(nCurrentAttackSpellNum) & " @ " & nCurrentAttackSpellLVL
                    If nCurrentAttackSpellNum <= 0 Then GoTo no_attack:
                    If frmMain.chkGlobalFilter.Value = 1 Then
                        tSpellCast = CalculateSpellCast(nCurrentAttackSpellNum, Val(frmMain.txtGlobalLevel(0).Text), Val(frmMain.lblCharSC.Tag), _
                            tabMonsters.Fields("MagicRes"), bHasAntiMagic)
                    Else
                        tSpellCast = CalculateSpellCast(nCurrentAttackSpellNum, 0, 0, tabMonsters.Fields("MagicRes"), bHasAntiMagic)
                    End If
                    nDamageOut = tSpellCast.nAvgRoundDmg
                    
                Case 4: 'martial arts attack
                    '1-Punch, 2-Kick, 3-JumpKick
                    Select Case nCurrentAttackMA
                        Case 2: 'kick
                            tAttack = CalculateAttack(2, , True, False, 100, tabMonsters.Fields("ArmourClass"), tabMonsters.Fields("DamageResist"), nMobDodge)
                            nDamageOut = tAttack.nRoundTotal
                        Case 3: 'jumpkick
                            tAttack = CalculateAttack(3, , True, False, 100, tabMonsters.Fields("ArmourClass"), tabMonsters.Fields("DamageResist"), nMobDodge)
                            nDamageOut = tAttack.nRoundTotal
                        Case Else: 'punch
                            tAttack = CalculateAttack(1, , True, False, 100, tabMonsters.Fields("ArmourClass"), tabMonsters.Fields("DamageResist"), nMobDodge)
                            nDamageOut = tAttack.nRoundTotal
                    End Select
                    
                Case 5: 'manual
                    nDamageOut = nCurrentAttackManual
                    'nDamageOutSpell = nCurrentAttackManualMag
                    
                Case Else: '1-Shot All
                    nDamageOut = 9999999
                    'nDamageOutSpell = 9999999
                    
            End Select
        End If
no_attack:

        nMobExpPerHour() = CalcMobExpPerHour(tabMonsters.Fields("Number"), nDamageOut, nCharHealth, nAvgDmg, tabMonsters.Fields("HP"), _
            nHPRegen, tabMonsters.Fields("HPRegen"), Val(frmMain.txtMonsterDamage.Text), nParty)
        
        nExpDmgHP = nMobExpPerHour(0)
        nRestingRate = nMobExpPerHour(1)
        
    Else
        nExpDmgHP = 0
    End If
    
    If nExpDmgHP > 0 And tLastAvgLairInfo.nMobs > 0 Then
        nPossSpawns = InstrCount(tabMonsters.Fields("Summoned By"), "Group:") + tLastAvgLairInfo.nMobs
        If nPossSpawns > (tLastAvgLairInfo.nMobs * nMonsterLairRatioMultiplier) Then '(nmobs = # lairs) ... indication of a lot of walking distance between lairs
            nExpDmgHP = Round(((tLastAvgLairInfo.nMobs * nMonsterLairRatioMultiplier) / nPossSpawns) * nExpDmgHP)
        End If
    End If
    
    If nExpDmgHP > 0 And Val(frmMain.txtMonsterLairFilter(0).Text) > 1 Then
        nExpDmgHP = Round(nExpDmgHP / Val(frmMain.txtMonsterLairFilter(0).Text))
    End If
    
    If nExpDmgHP > 1000000 Then
        sTemp = Format((nExpDmgHP / 1000000), "#,#.0") & " M"
    ElseIf nExpDmgHP > 1000 Then
        sTemp = Format((nExpDmgHP / 1000), "#,#.0") & " K"
    Else
        sTemp = IIf(nExpDmgHP > 0, Format(RoundUp(nExpDmgHP), "#,#"), "0")
    End If
    
    If nExpDmgHP > 0 And Val(frmMain.txtMonsterLairFilter(0).Text) > 1 Then
        sTemp = sTemp & "/hr ea."
    Else
        sTemp = sTemp & "/hr"
    End If
    
    oLI.ListSubItems.Add (nIndex), "Exp/(Dmg+HP)", sTemp & IIf(bAsterisks, " *", "")
    oLI.ListSubItems(nIndex).Tag = nExpDmgHP
    
    If Val(frmMain.txtMonsterLairFilter(0).Text) > 1 Then
        If nMonsterDamageVsParty(tabMonsters.Fields("Number")) >= 0 Then oLI.ListSubItems(nIndex).ForeColor = RGB(193, 0, 232)
    ElseIf frmMain.chkGlobalFilter.Value = 1 And nMonsterDamageVsChar(tabMonsters.Fields("Number")) >= 0 Then
        oLI.ListSubItems(nIndex).ForeColor = RGB(193, 0, 232)
    End If
    
ElseIf nExp > 0 Then
    
'    If nMonsterNum = 281 Then
'        Debug.Print nMonsterNum
'    End If
    
    If nAvgDmg > 0 Or nHP > 0 Then
        If nAvgDmg < 0 Then nAvgDmg = 0
        nExpDmgHP = Round(nExp / ((nAvgDmg * 2) + nHP), 2) * 100
    Else
        nExpDmgHP = nExp * 100
    End If
    
    oLI.ListSubItems.Add (nIndex), "Exp/(Dmg+HP)", IIf(nExpDmgHP > 0, Format(nExpDmgHP, "#,#"), 0) & IIf(bAsterisks, "*", "")
    oLI.ListSubItems(nIndex).Tag = nExpDmgHP
    
    If frmMain.optMonsterFilter(1).Value = True And Val(frmMain.txtMonsterLairFilter(0).Text) > 1 Then
        If nMonsterDamageVsParty(tabMonsters.Fields("Number")) >= 0 Then oLI.ListSubItems(nIndex).ForeColor = RGB(193, 0, 232)
    ElseIf frmMain.chkGlobalFilter.Value = 1 And nMonsterDamageVsChar(tabMonsters.Fields("Number")) >= 0 Then
        oLI.ListSubItems(nIndex).ForeColor = RGB(193, 0, 232)
    End If
Else
    oLI.ListSubItems.Add (nIndex), "Exp/(Dmg+HP)", 0
    oLI.ListSubItems(nIndex).Tag = nExp
End If

'a lot of this repeated in pullmonsterdetail AND apply monster filter
nPossSpawns = 0
nLairPCT = 0
nMaxLairsBeforeRegen = nTheoreticalAvgMaxLairsPerRegenPeriod
If InStr(1, tabMonsters.Fields("Summoned By"), "(lair)", vbTextCompare) > 0 Then
    nPossSpawns = InstrCount(tabMonsters.Fields("Summoned By"), "(lair)")
    sPossSpawns = nPossSpawns
    
    If nMonsterPossy(tabMonsters.Fields("Number")) > 0 Then nMaxLairsBeforeRegen = Round(nMaxLairsBeforeRegen / nMonsterPossy(tabMonsters.Fields("Number")), 2)
    If nPossSpawns < nMaxLairsBeforeRegen Then
        nLairPCT = Round(nPossSpawns / nMaxLairsBeforeRegen, 2)
    Else
        nLairPCT = 1
    End If
End If

sTemp = ""
nPossyPCT = 1
nScriptValue = 0
nIndex = nIndex + 1
'If tAvgLairInfo.nMobs > 0 Then
    'nScriptValue = tAvgLairInfo.nScriptValue
    'sTemp = "*"
    
'Else if...
If nNMRVer >= 1.83 And (nMonsterDamageVsChar(tabMonsters.Fields("Number")) < 0 Or frmMain.chkGlobalFilter.Value = 0) Then
    nScriptValue = tabMonsters.Fields("ScriptValue")
    
ElseIf tabMonsters.Fields("RegenTime") = 0 And nLairPCT > 0 Then
    
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

If nNMRVer >= 1.83 And frmMain.optMonsterFilter(1).Value = True And LV.hWnd = frmMain.lvMonsters.hWnd Then 'by lair
    oLI.ListSubItems.Add (nIndex), "Script Value", Round(nRestingRate * 100) & "%" '(resting rate substituted here for by lair)
    oLI.ListSubItems(nIndex).Tag = Round(nRestingRate * 100)
Else
    If nScriptValue > 1000000000 Then
        oLI.ListSubItems.Add (nIndex), "Script Value", Format((nScriptValue / 1000000), "#,#M") & sTemp
    ElseIf nScriptValue > 1000000 Then
        oLI.ListSubItems.Add (nIndex), "Script Value", Format((nScriptValue / 1000), "#,#K") & sTemp
    Else
        oLI.ListSubItems.Add (nIndex), "Script Value", IIf(nScriptValue > 0, Format(RoundUp(nScriptValue), "#,#"), "0") & sTemp
    End If
    oLI.ListSubItems(nIndex).Tag = nScriptValue
End If

If frmMain.optMonsterFilter(1).Value = True And Val(frmMain.txtMonsterLairFilter(0).Text) > 1 Then
    If nMonsterDamageVsParty(tabMonsters.Fields("Number")) >= 0 Then oLI.ListSubItems(nIndex).ForeColor = RGB(193, 0, 232)
ElseIf nMonsterDamageVsChar(tabMonsters.Fields("Number")) >= 0 And frmMain.chkGlobalFilter.Value = 1 Then
    oLI.ListSubItems(nIndex).ForeColor = RGB(193, 0, 232)
End If

If nNMRVer >= 1.83 Then
    nIndex = nIndex + 1
    oLI.ListSubItems.Add (nIndex), "Lair Exp", PutCommas(tabMonsters.Fields("AvgLairExp"))
    oLI.ListSubItems(nIndex).Tag = tabMonsters.Fields("AvgLairExp")
End If

nIndex = nIndex + 1
oLI.ListSubItems.Add (nIndex), "Lairs", sPossSpawns
oLI.ListSubItems(nIndex).Tag = nPossSpawns

If nNMRVer >= 1.82 Then
    nIndex = nIndex + 1
    If nMonsterPossy(tabMonsters.Fields("Number")) > 0 Then
        If nMonsterSpawnChance(tabMonsters.Fields("Number")) > 0 Then
            oLI.ListSubItems.Add (nIndex), "Mobs/Spwn", nMonsterPossy(tabMonsters.Fields("Number")) & " / " & (nMonsterSpawnChance(tabMonsters.Fields("Number")) * 100) & "%"
            oLI.ListSubItems(nIndex).Tag = (nMonsterSpawnChance(tabMonsters.Fields("Number")) * 100) * nMonsterPossy(tabMonsters.Fields("Number"))
        Else
            oLI.ListSubItems.Add (nIndex), "#Mobs", nMonsterPossy(tabMonsters.Fields("Number"))
            oLI.ListSubItems(nIndex).Tag = nMonsterPossy(tabMonsters.Fields("Number"))
        End If
    Else
        oLI.ListSubItems.Add (nIndex), "Mobs/Spwn", ""
        oLI.ListSubItems(nIndex).Tag = 0
    End If
End If

nIndex = nIndex + 1
oLI.ListSubItems.Add (nIndex), "Mag.", IIf(nMagicLVL > 0, nMagicLVL, "")
oLI.ListSubItems(nIndex).Tag = nMagicLVL

nIndex = nIndex + 1
oLI.ListSubItems.Add (nIndex), "Undead", IIf(tabMonsters.Fields("Undead") > 0, "X", "")
oLI.ListSubItems(nIndex).Tag = tabMonsters.Fields("Undead")

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

Public Function CalcMobExpPerHour(ByVal nMonster As Long, ByVal nCharDMG As Double, ByVal nCharHP As Long, ByVal nMobDMG As Double, ByVal nMobHP As Long, _
    Optional ByVal nCharHPRegen As Integer = 0, Optional ByVal nMobHPRegen As Long = 0, Optional ByVal nDamageThreshold As Long = 0, Optional ByVal nParty As Integer = 0) As Currency()
On Error GoTo error:
Dim nReturn(1) As Currency, nExp As Currency, nTimeFactor As Double, nRestingRate As Double
'0 = exp
'1 = resting rate
CalcMobExpPerHour = nReturn

If nMonster = 0 Then Exit Function
If tabMonsters.RecordCount = 0 Then Exit Function

On Error GoTo lookup:
If tabMonsters.Fields("Number") = nMonster Then GoTo continue:
GoTo lookup2:

lookup:
Resume lookup2:
lookup2:
On Error GoTo error:
tabMonsters.Index = "pkMonsters"
tabMonsters.Seek "=", nMonster
If tabMonsters.NoMatch = True Then
    tabMonsters.MoveFirst
    Exit Function
End If
continue:

'can we kill the mob?
If IsMobKillable(nCharDMG, nCharHP, nMobDMG, nMobHP, nCharHPRegen, nMobHPRegen) = False Then
    nReturn(0) = -1
    nReturn(1) = 1
    CalcMobExpPerHour = nReturn
    Exit Function
End If

If UseExpMulti Then
    nExp = tabMonsters.Fields("EXP") * tabMonsters.Fields("ExpMulti")
Else
    nExp = tabMonsters.Fields("EXP")
End If

If tabMonsters.Fields("RegenTime") > 0 Then
    nExp = Round(nExp / tabMonsters.Fields("RegenTime"))
Else
    nExp = nExp * nTheoreticalAvgMaxLairsPerRegenPeriod
    nTimeFactor = 20
    If nDamageThreshold > 0 And nDamageThreshold < nMobDMG Then
        nRestingRate = CalcPercentTimeSpentResting(nMobDMG - nDamageThreshold, nCharDMG, nMobHP, nCharHPRegen)
    ElseIf nDamageThreshold = 0 Then
        nRestingRate = CalcPercentTimeSpentResting(nMobDMG - (nCharHPRegen / 3 / 6), nCharDMG, nMobHP, nCharHPRegen)
    End If
    nTimeFactor = 20 * (1 - nRestingRate)
    nExp = nExp * nTimeFactor
End If

nReturn(0) = nExp
nReturn(1) = nRestingRate
CalcMobExpPerHour = nReturn

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalcMobExpPerHour")
Resume out:
End Function

Public Sub AddShop2LV(LV As ListView)

On Error GoTo error:

Dim oLI As ListItem, sName As String
    
    sName = tabShops.Fields("Name")
    If sName = "" Or Left(LCase(sName), 3) = "sdf" Then GoTo skip:
    If sName = "Leave this blank" Then GoTo skip:
    
    Set oLI = LV.ListItems.Add()
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

Public Sub AddClass2LV(LV As ListView)

On Error GoTo error:

Dim oLI As ListItem, x As Integer, sAbil As String
    
    If tabClasses.Fields("Name") = "" Then GoTo skip:
    
    Set oLI = LV.ListItems.Add()
    oLI.Text = tabClasses.Fields("Number")
    
    oLI.ListSubItems.Add (1), "Name", tabClasses.Fields("Name")
    oLI.ListSubItems.Add (2), "Exp%", (Val(tabClasses.Fields("ExpTable")) + 100) & "%"
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



Public Sub RaceColorCode(LV As ListView)
On Error GoTo error:
Dim oLI As ListItem, x As Integer
Dim Stat(1 To 6, 1 To 2) As Integer, Min(1 To 6) As Integer, Max(1 To 6) As Integer, nRaces As Integer
'1-6 = str, int, wis, agi, hea, cha

'stat, 1 = min
'stat, 2 = max
'min, total min
'max, total max
'then get avg

For Each oLI In LV.ListItems
    
    tabRaces.Index = "pkRaces"
    tabRaces.Seek "=", Val(oLI.Text)
    If tabRaces.NoMatch = False Then
        
        nRaces = nRaces + 1
        Stat(1, 1) = Val(tabRaces.Fields("mSTR"))
        Stat(2, 1) = Val(tabRaces.Fields("mINT"))
        Stat(3, 1) = Val(tabRaces.Fields("mWIL"))
        Stat(4, 1) = Val(tabRaces.Fields("mAGL"))
        Stat(5, 1) = Val(tabRaces.Fields("mHEA"))
        Stat(6, 1) = Val(tabRaces.Fields("mCHM"))
        Stat(1, 2) = Val(tabRaces.Fields("xSTR"))
        Stat(2, 2) = Val(tabRaces.Fields("xINT"))
        Stat(3, 2) = Val(tabRaces.Fields("xWIL"))
        Stat(4, 2) = Val(tabRaces.Fields("xAGL"))
        Stat(5, 2) = Val(tabRaces.Fields("xHEA"))
        Stat(6, 2) = Val(tabRaces.Fields("xCHM"))
        
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

For Each oLI In LV.ListItems
    
    tabRaces.Index = "pkRaces"
    tabRaces.Seek "=", Val(oLI.Text)
    If tabRaces.NoMatch = False Then
        
        
        If Val(tabRaces.Fields("mSTR")) + Val(tabRaces.Fields("xSTR")) < Stat(1, 1) - 20 Then oLI.ListSubItems(4).ForeColor = &H80&
        If Val(tabRaces.Fields("mINT")) + Val(tabRaces.Fields("xINT")) < Stat(2, 1) - 20 Then oLI.ListSubItems(5).ForeColor = &H80&
        If Val(tabRaces.Fields("mWIL")) + Val(tabRaces.Fields("xWIL")) < Stat(3, 1) - 20 Then oLI.ListSubItems(6).ForeColor = &H80&
        If Val(tabRaces.Fields("mAGL")) + Val(tabRaces.Fields("xAGL")) < Stat(4, 1) - 20 Then oLI.ListSubItems(7).ForeColor = &H80&
        If Val(tabRaces.Fields("mHEA")) + Val(tabRaces.Fields("xHEA")) < Stat(5, 1) - 20 Then oLI.ListSubItems(8).ForeColor = &H80&
        If Val(tabRaces.Fields("mCHM")) + Val(tabRaces.Fields("xCHM")) < Stat(6, 1) - 20 Then oLI.ListSubItems(9).ForeColor = &H80&
        
        If Val(tabRaces.Fields("mSTR")) + Val(tabRaces.Fields("xSTR")) > Stat(1, 1) + 20 Then oLI.ListSubItems(4).ForeColor = &H8000&
        If Val(tabRaces.Fields("mINT")) + Val(tabRaces.Fields("xINT")) > Stat(2, 1) + 20 Then oLI.ListSubItems(5).ForeColor = &H8000&
        If Val(tabRaces.Fields("mWIL")) + Val(tabRaces.Fields("xWIL")) > Stat(3, 1) + 20 Then oLI.ListSubItems(6).ForeColor = &H8000&
        If Val(tabRaces.Fields("mAGL")) + Val(tabRaces.Fields("xAGL")) > Stat(4, 1) + 20 Then oLI.ListSubItems(7).ForeColor = &H8000&
        If Val(tabRaces.Fields("mHEA")) + Val(tabRaces.Fields("xHEA")) > Stat(5, 1) + 20 Then oLI.ListSubItems(8).ForeColor = &H8000&
        If Val(tabRaces.Fields("mCHM")) + Val(tabRaces.Fields("xCHM")) > Stat(6, 1) + 20 Then oLI.ListSubItems(9).ForeColor = &H8000&
        
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

Public Sub CopyWholeLVtoClipboard(LV As ListView, Optional ByVal UsePeriods As Boolean)
On Error GoTo error:
Dim oLI As ListItem, oLSI As ListSubItem, oCH As ColumnHeader
Dim str As String, x As Integer, sSpacer As String, nLongText() As Integer
    
str = ""
sSpacer = IIf(UsePeriods, ".", " ")

ReDim nLongText(0 To LV.ColumnHeaders.Count - 1)

'find longest text(s)
For Each oLI In LV.ListItems
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
For Each oCH In LV.ColumnHeaders
    str = str & oCH.Text
    str = str & " " & String(nLongText(x) - Len(oCH.Text), " ") & " "
    x = x + 1
Next

str = str & vbCrLf

For Each oLI In LV.ListItems
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
Public Sub CopyLVLinetoClipboard(LV As ListView, Optional DetailTB As TextBox, _
    Optional LocationLV As ListView, Optional ByVal nExcludeColumn As Integer = -1, Optional bNameOnly As Boolean = False)
On Error GoTo error:
Dim oLI As ListItem, oLI2 As ListItem, oCH As ColumnHeader
Dim str As String, x As Integer, nCount As Integer

If LV.ListItems.Count < 1 Then Exit Sub

nCount = 1
For Each oLI In LV.ListItems
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
        For Each oCH In LV.ColumnHeaders
            If Not x = nExcludeColumn Then
                If bNameOnly Then
                    If (LV.name = "lvMapLoc" Or LV.name = "lvSpellLoc" Or LV.name = "lvShopLoc") And x = 0 Then
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
                    ElseIf LV.name = "lvWeaponLoc" Or LV.name = "lvArmourLoc" Then
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
        
        Select Case LV.name
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

Public Sub GetLocations(ByVal sLoc As String, LV As ListView, _
    Optional bDontClear As Boolean, Optional ByVal sHeader As String, _
    Optional ByVal nAuxValue As Long, Optional ByVal bTwoColumns As Boolean, _
    Optional ByVal bDontSort As Boolean, Optional ByVal bPercentColumn As Boolean, _
    Optional ByVal sFooter As String, Optional ByVal nLimit As Integer)
On Error GoTo error:
Dim sLook As String, sChar As String, sTest As String, oLI As ListItem, sPercent As String
Dim x As Integer, y1 As Integer, y2 As Integer, z As Integer, nValue As Long, x2 As Integer
Dim sLocation As String, nPercent As Currency, nPercent2 As Currency, sTemp As String, nSpawnChance As Currency
Dim sDisplayFooter As String, sLairRegex As String, sRoomKey As String
Dim tMatches() As RegexMatches, nMaxRegen As Integer, sGroupIndex As String, tLairInfo As LairInfoType
Dim nCount As Integer

sDisplayFooter = sFooter

If Not bDontClear Then LV.ListItems.clear
If bDontSort Then LV.Sorted = False

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
            nValue = Val(Mid(sTest, y1, y2))
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
                
                Set oLI = LV.ListItems.Add()
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
                Set oLI = LV.ListItems.Add()
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
                Set oLI = LV.ListItems.Add()
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
                Set oLI = LV.ListItems.Add()
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
                Set oLI = LV.ListItems.Add()
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
                    Call GetLocations(tabItems.Fields("Obtained From"), LV, True, , , , True, bPercentColumn, " -> " & tabItems.Fields("Name") & sPercent, nLimit - nCount)
                End If
                
            Case 6: '"spell #"
                If nLimit > 0 Then nCount = nCount + 1
                If nLimit > 0 And nCount > nLimit Then GoTo skip:
                sLocation = "Spell: "
                Set oLI = LV.ListItems.Add()
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
                If bPercentColumn And nAuxValue > 0 Then
                    nPercent = GetItemShopRegenPCT(nValue, nAuxValue)
                    If nPercent > 0 Then
                        sTemp = GetShopLocation(nValue)
                        sTemp = Join(Split(sTemp, ","), "(" & nPercent & "%),")
                        If Not Right(sTemp, 2) = "%)" Then sTemp = sTemp & "(" & nPercent & "%)"
                        Call GetLocations(sTemp, LV, True, "Shop: ", nValue, , , bPercentColumn, , nLimit - nCount)
                    Else
                        Call GetLocations(GetShopLocation(nValue), LV, True, "Shop: ", nValue, , , bPercentColumn, , nLimit - nCount)
                    End If
                Else
                    Call GetLocations(GetShopLocation(nValue), LV, True, "Shop: ", nValue, , , , , nLimit - nCount)
                End If
                
            Case 8: '"shop(sell) #"
                Call GetLocations(GetShopLocation(nValue), LV, True, "Shop (sell): ", nValue, , , bPercentColumn, , nLimit - nCount)
'
            Case 9: '"shop(nogen) #"
                Call GetLocations(GetShopLocation(nValue), LV, True, "Shop (nogen): ", nValue, , , bPercentColumn, , nLimit - nCount)
'
            Case 10: 'group (lair)
                If nLimit > 0 Then nCount = nCount + 1
                If nLimit > 0 And nCount > nLimit Then GoTo skip:
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
                        nMaxRegen = Val(tMatches(0).sSubMatches(1))
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
                        nMaxRegen = Val(tMatches(0).sSubMatches(0))
                        sRoomKey = tMatches(0).sSubMatches(1) & "/" & tMatches(0).sSubMatches(2)
                    Else
                        'Group(lair): 1/2345
                        sRoomKey = tMatches(0).sSubMatches(0) & "/" & tMatches(0).sSubMatches(1)
                    End If
                Else
                    sRoomKey = Mid(sTest, y1, y2)
                End If
                
                If nSpawnChance > 0 Then
                    sLocation = "Group(Lair " & nMaxRegen & " / " & nSpawnChance & "%)"
                ElseIf nMaxRegen > 0 Then
                    sLocation = "Group(Lair " & nMaxRegen & ")"
                Else
                    sLocation = "Group(Lair)"
                End If
                
                Set oLI = LV.ListItems.Add()
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
                    If tLairInfo.nAvgDmg > 0 And tLairInfo.nDamageOut > 0 And (nCurrentAttackType > 0 And nCurrentAttackType <> 5) Then
                        sTemp = sTemp & ", Dmg In/Out: " & (tLairInfo.nAvgDmg * tLairInfo.nMaxRegen) & "/" & tLairInfo.nDamageOut
                    ElseIf tLairInfo.nAvgDmg > 0 Then
                        sTemp = sTemp & ", Dmg In: " & (tLairInfo.nAvgDmg * tLairInfo.nMaxRegen)
                    ElseIf tLairInfo.nDamageOut > 0 And (nCurrentAttackType > 0 And nCurrentAttackType <> 5) Then
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
                Set oLI = LV.ListItems.Add()
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
                Set oLI = LV.ListItems.Add()
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

If LV.ListItems.Count > 1 And Not bDontSort And Not bPercentColumn Then
    Call SortListView(LV, 1, ldtstring, True)
    LV.Sorted = False
End If

If LV.ListItems.Count > 1 Then
    If Right(sLoc, 2) = "+" & Chr(0) Then
        Set oLI = LV.ListItems.Add(LV.ListItems.Count + 1)
        If bTwoColumns Then
            oLI.ListSubItems.Add 1, , "... plus more."
        Else
            oLI.Text = "... plus more."
        End If
        oLI.Tag = 0
    ElseIf nLimit > 0 And nCount >= nLimit And sHeader = "" And sFooter = "" Then
        Set oLI = LV.ListItems.Add(LV.ListItems.Count + 1)
        If bTwoColumns Then
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

Private Function GetShopLocation(ByVal nNum As Long) As String
On Error GoTo error:

tabShops.Index = "pkShops"
tabShops.Seek "=", nNum
If tabShops.NoMatch Then
    GetShopLocation = ""
    tabShops.MoveFirst
    Exit Function
End If

GetShopLocation = tabShops.Fields("Assigned To")

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetShopLocation")
Resume out:
End Function

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
bAppTerminating = True
bCancelTerminate = False

Call UnloadForms("none")
DoEvents

If bCancelTerminate Then
    bAppTerminating = False
    Exit Sub
End If

Call CloseDatabases
DoEvents

'End
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
    sRet(0) = Val(Replace(Replace(sNumberString, "(", "", 1, -1, vbTextCompare), ")", "", 1, -1, vbTextCompare))
Else
    sRet = Split(sNumberString, ",")
    For x = 0 To UBound(sRet())
        sRet(x) = Val(Replace(Replace(sRet(x), "(", "", 1, -1, vbTextCompare), ")", "", 1, -1, vbTextCompare))
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

' Omit pvarMirror, plngLeft & plngRight; they are used internally during recursion
Public Sub MergeSort1(ByRef pvarArray As Variant, Optional pvarMirror As Variant, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngMid As Long
    Dim l As Long
    Dim r As Long
    Dim O As Long
    Dim varSwap As Variant
 
    If plngRight = 0 Then
        plngLeft = LBound(pvarArray)
        plngRight = UBound(pvarArray)
        ReDim pvarMirror(plngLeft To plngRight)
    End If
    lngMid = plngRight - plngLeft
    Select Case lngMid
        Case 0
        Case 1
            If pvarArray(plngLeft) > pvarArray(plngRight) Then
                varSwap = pvarArray(plngLeft)
                pvarArray(plngLeft) = pvarArray(plngRight)
                pvarArray(plngRight) = varSwap
            End If
        Case Else
            lngMid = lngMid \ 2 + plngLeft
            MergeSort1 pvarArray, pvarMirror, plngLeft, lngMid
            MergeSort1 pvarArray, pvarMirror, lngMid + 1, plngRight
            ' Merge the resulting halves
            l = plngLeft ' start of first (left) half
            r = lngMid + 1 ' start of second (right) half
            O = plngLeft ' start of output (mirror array)
            Do
                If pvarArray(r) < pvarArray(l) Then
                    pvarMirror(O) = pvarArray(r)
                    r = r + 1
                    If r > plngRight Then
                        For l = l To lngMid
                            O = O + 1
                            pvarMirror(O) = pvarArray(l)
                        Next
                        Exit Do
                    End If
                Else
                    pvarMirror(O) = pvarArray(l)
                    l = l + 1
                    If l > lngMid Then
                        For r = r To plngRight
                            O = O + 1
                            pvarMirror(O) = pvarArray(r)
                        Next
                        Exit Do
                    End If
                End If
                O = O + 1
            Loop
            For O = plngLeft To plngRight
                pvarArray(O) = pvarMirror(O)
            Next
    End Select
End Sub

Public Sub ColorListviewRow(LV As ListView, RowNbr As Long, RowColor As OLE_COLOR, Optional bAndBold As Boolean)

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


Set itmX = LV.ListItems(RowNbr)
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
    If Not Val(tabRooms.Fields(sExits(x))) = 0 Then
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
    nValue = nValue + (x * Asc(Mid(sRoomName, x, 1)))
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

Public Function ClearSavedDamageVsMonster(Optional bPartyInstead As Boolean = False)
On Error GoTo error:
Dim x As Long

If bPartyInstead Then
    For x = 0 To UBound(nPartyDamageVsMonster)
        nPartyDamageVsMonster(x) = -1
    Next x
    'sPartyDamageVsMonsterConfig = ...
Else
    For x = 0 To UBound(nCharDamageVsMonster)
        nCharDamageVsMonster(x) = -1
    Next x
    sCharDamageVsMonsterConfig = sCurrentAttackConfig
End If

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
    bMonsterDamageVsCharCalculated = False
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
Call LockWindowUpdate(frmMain.hWnd)
nInterval = 1

tabMonsters.MoveFirst
Do While tabMonsters.EOF = False
    
    bHasAttack = False
    For x = 0 To (4 Or bHasAttack)
        If tabMonsters.Fields("AttType-" & x) > 0 And tabMonsters.Fields("AttType-" & x) < 4 Then bHasAttack = True
    Next x
    For x = 0 To (4 Or bHasAttack)
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

If bPartyInstead Then
    bMonsterDamageVsPartyCalculated = True
    bDontPromptCalcPartyMonsterDamage = True
Else
    bMonsterDamageVsCharCalculated = True
    bDontPromptCalcCharMonsterDamage = True
End If

out:
On Error Resume Next

If FormIsLoaded("frmProgressBar") Then
    Call LockWindowUpdate(0&)
    Unload frmProgressBar
End If
frmMain.Enabled = True
Exit Function
error:
Call HandleError("CalculateMonsterDamageVsCharALL")
Resume out:
End Function

Public Function CalculateMonsterDamageVsChar(ByVal nMonsterNumber As Long, Optional bPartyInstead As Boolean = False) As Currency
On Error GoTo error:
Dim nNon As Integer, nAnti As Integer

If nMonsterNumber <= 0 Then Exit Function
If nMonsterSimRounds < 100 Then nMonsterSimRounds = 100
If nMonsterSimRounds > 10000 Then nMonsterSimRounds = 10000

If Val(frmMain.txtMonsterLairFilter(0).Text) < 2 Or Val(frmMain.txtMonsterLairFilter(6).Text) < 1 _
    Or Val(frmMain.txtMonsterLairFilter(0).Text) = Val(frmMain.txtMonsterLairFilter(6).Text) _
    Or bPartyInstead = False Then
    
    nAnti = 0: If Val(frmMain.txtMonsterLairFilter(0).Text) = Val(frmMain.txtMonsterLairFilter(6).Text) Then nAnti = 1
    
    Call SetupMonsterAttackSimWithCharStats(nMonsterSimRounds, False, bPartyInstead, nAnti)
    Call PopulateMonsterDataToAttackSim(nMonsterNumber, clsMonAtkSim)
    If clsMonAtkSim.nNumberOfRounds > 0 Then clsMonAtkSim.RunSim
    CalculateMonsterDamageVsChar = clsMonAtkSim.nAverageDamage
Else 'party
    nAnti = Val(frmMain.txtMonsterLairFilter(6).Text)
    nNon = Val(frmMain.txtMonsterLairFilter(0).Text) - nAnti
    If nAnti > 0 Then
        Call SetupMonsterAttackSimWithCharStats(nMonsterSimRounds, False, bPartyInstead, 1)
        Call PopulateMonsterDataToAttackSim(nMonsterNumber, clsMonAtkSim)
        If clsMonAtkSim.nNumberOfRounds > 0 Then clsMonAtkSim.RunSim
        CalculateMonsterDamageVsChar = clsMonAtkSim.nAverageDamage * nAnti
    End If
    If nNon > 0 Then
        Call SetupMonsterAttackSimWithCharStats(nMonsterSimRounds, False, bPartyInstead, 0)
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
    If Val(frmMain.txtMonsterLairFilter(1).Text) > 0 Then clsMonAtkSim.nUserAC = Val(frmMain.txtMonsterLairFilter(1).Text)
    If Val(frmMain.txtMonsterLairFilter(2).Text) > 0 Then clsMonAtkSim.nUserDR = Val(frmMain.txtMonsterLairFilter(2).Text)
    If Val(frmMain.txtMonsterLairFilter(3).Text) > 0 Then clsMonAtkSim.nUserMR = Val(frmMain.txtMonsterLairFilter(3).Text)
    If Val(frmMain.txtMonsterLairFilter(4).Text) > 0 Then clsMonAtkSim.nUserDodge = Val(frmMain.txtMonsterLairFilter(4).Text)
    If nPartyAntiMagic = 1 Then clsMonAtkSim.nUserAntiMagic = 1
Else
    If Val(frmMain.txtCharAC.Text) > 0 Then clsMonAtkSim.nUserAC = Val(frmMain.txtCharAC.Text)
    If Val(frmMain.lblInvenCharStat(3).Caption) > 0 Then clsMonAtkSim.nUserDR = Val(frmMain.lblInvenCharStat(3).Caption)
    If Val(frmMain.txtCharMR.Text) > 0 Then clsMonAtkSim.nUserMR = Val(frmMain.txtCharMR.Text)
    If Val(frmMain.lblCharDodge.Tag) > 0 Then clsMonAtkSim.nUserDodge = Val(frmMain.lblCharDodge.Tag)
    If frmMain.chkCharAntiMagic.Value = 1 Then clsMonAtkSim.nUserAntiMagic = 1
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

Public Function CalcPercentTimeSpentResting(ByVal nDmgIN As Double, ByVal nDmgOUT As Double, ByVal nMobHP As Double, ByVal nRestHP As Double, _
    Optional ByVal nMobs As Integer = 0) As Double
On Error GoTo error:
Dim nKillTime As Double
Dim nDmgInTotal As Double
Dim nNetDmg As Double
Dim nRestTime As Double
Dim nTotalTime As Double
Dim nRestPCT As Double

If nMobs < 1 Then nMobs = 1
If nRestHP < 1 Then nRestHP = 1

If nDmgOUT >= nMobHP Then
    nKillTime = 0.5
    If nMobs > 1 Then nKillTime = nKillTime * nMobs
ElseIf nDmgOUT > 0 Then
    nKillTime = (nMobHP / nDmgOUT) * (1 - (nDmgOUT / (nDmgOUT + nMobHP)))
Else
    CalcPercentTimeSpentResting = 1
    Exit Function
End If
If nKillTime < 0.5 Then nKillTime = 0.5

nDmgInTotal = nKillTime * nDmgIN * (nDmgScaleFactor - Exp(-1 * nKillTime))

nNetDmg = nDmgInTotal - (nKillTime * (nRestHP / 18))
If nNetDmg < 0 Then nNetDmg = 0

nRestTime = (nNetDmg ^ nDmgScaleFactor) / (nRestHP / 3)

nTotalTime = nKillTime + nRestTime

If nTotalTime > 0 Then
    nRestPCT = (nRestTime / nTotalTime)
Else
    nRestPCT = 0
End If

If nRestPCT < 0 Then nRestPCT = 0
If nRestPCT > 1 Then nRestPCT = 1

CalcPercentTimeSpentResting = nRestPCT

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalcPercentTimeSpentResting")
Resume out:
End Function

Public Function InvenGetEquipInfo(ByVal nAbility As Integer, ByVal nAbilityValue As Integer) As TypeGetEquip

If nAbilityValue > 0 Then InvenGetEquipInfo.sText = GetAbilityStats(nAbility, nAbilityValue)

InvenGetEquipInfo.nEquip = -1
Select Case nAbility
    Case 0: 'nothing
    Case 2: '2=AC
        InvenGetEquipInfo.nEquip = 2
        'InvenGetEquipInfo.sText = "AC: "
    Case 3: '3=res_cold
        InvenGetEquipInfo.nEquip = 28
        'InvenGetEquipInfo.sText = "Cold Res: "
    Case 4: '4=max dmg
        InvenGetEquipInfo.nEquip = 11
        'InvenGetEquipInfo.sText = "Max Dmg: "
    Case 5: '5=res_fire
        InvenGetEquipInfo.nEquip = 27
        'InvenGetEquipInfo.sText = "Fire Res: "
    Case 7: '7=DR
        InvenGetEquipInfo.nEquip = 3
        'InvenGetEquipInfo.sText = "DR: "
    Case 10: '10=AC
        InvenGetEquipInfo.nEquip = 2
        'InvenGetEquipInfo.sText = "AC: "
    Case 13: '13=illu
        InvenGetEquipInfo.nEquip = 23
        'InvenGetEquipInfo.sText = "Illu: "
    Case 14: '14=roomillu
        InvenGetEquipInfo.nEquip = 23
        'InvenGetEquipInfo.sText = "RoomIllu: "
    Case 21: InvenGetEquipInfo.nEquip = 32 'immu poison
    Case 22: '22=acc
        InvenGetEquipInfo.nEquip = 10
        'InvenGetEquipInfo.sText = "Accy: "
    Case 24: '42=prev
        InvenGetEquipInfo.nEquip = 20
    '25=prgd
    Case 27: '27=stealth
        InvenGetEquipInfo.nEquip = 19
        'InvenGetEquipInfo.sText = "Stealth: "
    Case 29: InvenGetEquipInfo.nEquip = 37 'punch skill
    Case 30: InvenGetEquipInfo.nEquip = 38 'kick skill
    Case 34: '34=dodge
        InvenGetEquipInfo.nEquip = 8
        'InvenGetEquipInfo.sText = "Dodge: "
    Case 35: InvenGetEquipInfo.nEquip = 39 'jk skill
    Case 36: '36=MR
        InvenGetEquipInfo.nEquip = 24
        'InvenGetEquipInfo.sText = "MR: "
    Case 37: '37=picklocks
        InvenGetEquipInfo.nEquip = 22
        'InvenGetEquipInfo.sText = "Picks: "
    Case 38: '38=tracking
        'InvenGetEquipInfo.nEquip = 23
        ''InvenGetEquipInfo.sText = "Tracking: "
    Case 39: '39=thievery
        'InvenGetEquipInfo.nEquip = 20
        'InvenGetEquipInfo.sText = "Thievery: "
    Case 40: '40=findtraps
        InvenGetEquipInfo.nEquip = 21
        'InvenGetEquipInfo.sText = "Traps: "
    '41=disarmtraps
    '44=int
    '45=wis
    '46=str
    '47=hea
    '48=agi
    '49=chm
    Case 58: '58=crits
        InvenGetEquipInfo.nEquip = 7
        'InvenGetEquipInfo.sText = "Crits: "
    Case 65: '65=res_stone
        InvenGetEquipInfo.nEquip = 25
        'InvenGetEquipInfo.sText = "Stone Res: "
    Case 66: '66=res_lit
        InvenGetEquipInfo.nEquip = 29
        'InvenGetEquipInfo.sText = "Light Res: "
    Case 67: InvenGetEquipInfo.nEquip = 31 'quickness
    Case 69: '69=max mana
        InvenGetEquipInfo.nEquip = 6
        'InvenGetEquipInfo.sText = "Mana: "
    Case 70: '70=SC
        InvenGetEquipInfo.nEquip = 9
        'InvenGetEquipInfo.sText = "SC: "
    Case 72: '72=damageshield
        InvenGetEquipInfo.nEquip = 12
        'InvenGetEquipInfo.sText = "Shock: "
    Case 77: '77=percep
        InvenGetEquipInfo.nEquip = 18
        'InvenGetEquipInfo.sText = "Percep: "
    '87=speed
    Case 88: '88=alter hp
        InvenGetEquipInfo.nEquip = 5
        'InvenGetEquipInfo.sText = "HP: "
    
    Case 89: InvenGetEquipInfo.nEquip = 40 'punch accy
    Case 90: InvenGetEquipInfo.nEquip = 41 'kick accy
    Case 91: InvenGetEquipInfo.nEquip = 42 'jumpkick accy
    
    Case 92: InvenGetEquipInfo.nEquip = 34 'punch dmg
    Case 93: InvenGetEquipInfo.nEquip = 35 'kick dmg
    Case 94: InvenGetEquipInfo.nEquip = 36 'jumpkick dmg
    
    Case 96: '96=encum
        InvenGetEquipInfo.nEquip = 4
        'InvenGetEquipInfo.sText = "Enc%: "
    Case 105: '105=acc
        InvenGetEquipInfo.nEquip = 10
        'InvenGetEquipInfo.sText = "Accy: "
    Case 106: '106=acc
        InvenGetEquipInfo.nEquip = 10
        'InvenGetEquipInfo.sText = "Accy: "
    Case 116: '116=bsaccu
        InvenGetEquipInfo.nEquip = 13
        'InvenGetEquipInfo.sText = "BS Accy: "
    Case 117: '117=bsmin
        InvenGetEquipInfo.nEquip = 14
        'InvenGetEquipInfo.sText = "BS Min: "
    Case 118: '118=bsmax
        InvenGetEquipInfo.nEquip = 15
        'InvenGetEquipInfo.sText = "BS Max: "
    Case 123: '123=hpregen
        InvenGetEquipInfo.nEquip = 16
        'InvenGetEquipInfo.sText = "HP Rgn: "
    Case 142: '142=hitmagic
        'InvenGetEquipInfo.nEquip = 31
        ''InvenGetEquipInfo.sText = "Hit Magic: "
    Case 145: '145=manaregen
        InvenGetEquipInfo.nEquip = 17
        'InvenGetEquipInfo.sText = "Mana Rgn: "
    Case 147: '147=res_water
        InvenGetEquipInfo.nEquip = 26
        'InvenGetEquipInfo.sText = "Water Res: "
    Case 165: InvenGetEquipInfo.nEquip = 33 'alter spell dmg
    Case 179: '179=find trap value
        InvenGetEquipInfo.nEquip = 21
        'InvenGetEquipInfo.sText = "Traps: "
    Case 180: '180=pick value
        InvenGetEquipInfo.nEquip = 22
        'InvenGetEquipInfo.sText = "Picks: "
    
End Select
End Function

