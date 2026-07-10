Attribute VB_Name = "modMenuSubClass"
Option Explicit

'Public API Declarations
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Public Constants
'Public Const GWL_WNDPROC = -4
Public Const WM_COMMAND = &H111
Private Const WM_XBUTTONUP As Long = &H20C&
Private Const WM_APPCOMMAND As Long = &H319&
Private Const XBUTTON1 As Long = &H1&
Private Const APPCOMMAND_BROWSER_BACKWARD As Long = 1&

'Public Variables
Public nMenuItemID As Integer 'holds the item identification number of the newly added menu items
Public oldWindowProc As Long 'a pointer to this form's old window procedure

'=== dark menu bar support (bDarkMode) ===
'Windows 8+ sends WM_UAH* messages that allow owner-drawing the menu bar,
'which is otherwise always drawn in the light system colors.
Private Const WM_UAHDRAWMENU As Long = &H91
Private Const WM_UAHDRAWMENUITEM As Long = &H92
Private Const WM_NCPAINT As Long = &H85
Private Const WM_NCACTIVATE As Long = &H86
Private Const OBJID_MENU As Long = &HFFFFFFFD
Private Const ODS_SELECTED As Long = &H1
Private Const ODS_GRAYED As Long = &H2
Private Const ODS_HOTLIGHT As Long = &H40
Private Const ODS_INACTIVE As Long = &H80
Private Const ODS_NOACCEL As Long = &H100
Private Const DT_CENTER As Long = &H1
Private Const DT_VCENTER As Long = &H4
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_HIDEPREFIX As Long = &H100000
Private Const BKMODE_TRANSPARENT As Long = 1
Private Const MF_BYPOSITION As Long = &H400
Private Const DKMENU_HILITE As Long = &H403E3E

Private Type RECT_DKM
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type MENUBARINFO
    cbSize As Long
    rcBar As RECT_DKM
    hMenu As Long
    hwndMenu As Long
    fFlags As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT_DKM, ByVal hBrush As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT_DKM, ByVal wFormat As Long) As Long
Private Declare Function GetMenuStringA Lib "user32" (ByVal hMenu As Long, ByVal uIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal uFlag As Long) As Long
Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hWnd As Long, ByVal idObject As Long, ByVal idItem As Long, pmbi As MENUBARINFO) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT_DKM) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT_DKM) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Private Function HiWordUnsigned(ByVal dwValue As Long) As Long
    If dwValue < 0 Then
        HiWordUnsigned = ((dwValue And &H7FFF0000) \ &H10000) Or &H8000&
    Else
        HiWordUnsigned = (dwValue \ &H10000) And &HFFFF&
    End If
End Function

Private Function GetAppCommand(ByVal lParam As Long) As Long
    GetAppCommand = HiWordUnsigned(lParam) And &HFFF&
End Function

Public Function MenuWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long 'Processes window messages
    'There is no way for Visual Basic to create an event handler
    'to process whatever functions that need to be performed by the
    'newly created menu items. To work around this problem, it is necessary
    'to create this 'WindowProc' function to manually process the WM_COMMAND
    'messages that the new menu items send to the form's window...
    
    Dim retval As Long  'holds the return value

    Select Case uMsg
        Case WM_UAHDRAWMENU
            If bDarkMode Then
                If DarkMenuBarBackground(hWnd, lParam) = 1 Then
                    MenuWindowProc = 1
                    Exit Function
                End If
            End If

        Case WM_UAHDRAWMENUITEM
            If bDarkMode Then
                If DarkMenuBarItem(lParam) = 1 Then
                    MenuWindowProc = 1
                    Exit Function
                End If
            End If

        Case WM_NCPAINT, WM_NCACTIVATE
            If bDarkMode Then
                'let the default paint happen, then paint over the light
                '1px line that remains under the menu bar
                retval = CallWindowProc(oldWindowProc, hWnd, uMsg, wParam, lParam)
                Call DarkMenuBarBottomLine(hWnd)
                MenuWindowProc = retval
                Exit Function
            End If

        Case WM_XBUTTONUP
            If HiWordUnsigned(wParam) = XBUTTON1 Then
                If frmMain.NavHistoryBack Then
                    MenuWindowProc = 1
                    Exit Function
                End If
            End If

        Case WM_APPCOMMAND
            If GetAppCommand(lParam) = APPCOMMAND_BROWSER_BACKWARD Then
                If frmMain.NavHistoryBack Then
                    MenuWindowProc = 1
                    Exit Function
                End If
            End If
    End Select

    If uMsg = WM_COMMAND Then
        If wParam >= 1000 Then 'if the window command was received from one of our new menu items
            'This is where you set up event handling for our new menu items.
            'EXAMPLE:
            Select Case wParam
'                Case 1000: 'First New Menu Item (be careful, the item may be a separator bar!)
'                    Do Something
                Case 1001: 'Second New Menu Item
                    Call frmMain.RecentFilesLoad(1)
                Case 1002: 'Third New Menu Item
                    Call frmMain.RecentFilesLoad(2)
                Case 1003:
                    Call frmMain.RecentFilesLoad(3)
                Case 1004:
                    Call frmMain.RecentFilesLoad(4)
                Case 1005:
                    Call frmMain.RecentFilesLoad(5)
            End Select
            
            'Sample event handling (changes the form's background color):
            'Randomize
            'frmMain.BackColor = QBColor(CInt(Rnd * 15))
        End If
    End If
    retval = CallWindowProc(oldWindowProc, hWnd, uMsg, wParam, lParam) 'use this form's original window procedure to finish processing this message
    MenuWindowProc = retval 'set the WindowProc function equal to whatever this form's original window procedure would have returned
End Function

'=========================================================================
' WM_UAHDRAWMENU: fill the whole menu bar strip with the dark background.
' lParam points at a UAHMENU struct {hmenu, hdc, dwFlags}.
'=========================================================================
Private Function DarkMenuBarBackground(ByVal hWnd As Long, ByVal lParam As Long) As Long
Dim nMenuDC As Long
Dim tMBI As MENUBARINFO
Dim tRW As RECT_DKM
Dim hBr As Long

Call CopyMemory(nMenuDC, ByVal (lParam + 4), 4&)
If nMenuDC = 0 Then Exit Function

tMBI.cbSize = Len(tMBI)
If GetMenuBarInfo(hWnd, OBJID_MENU, 0, tMBI) = 0 Then Exit Function

'rcBar is in screen coordinates; the DC is a window DC
Call GetWindowRect(hWnd, tRW)
tMBI.rcBar.Left = tMBI.rcBar.Left - tRW.Left
tMBI.rcBar.Right = tMBI.rcBar.Right - tRW.Left
tMBI.rcBar.Top = tMBI.rcBar.Top - tRW.Top
tMBI.rcBar.Bottom = tMBI.rcBar.Bottom - tRW.Top

hBr = CreateSolidBrush(DK_FORM_BACK)
Call FillRect(nMenuDC, tMBI.rcBar, hBr)
Call DeleteObject(hBr)

DarkMenuBarBackground = 1
End Function

'=========================================================================
' WM_UAHDRAWMENUITEM: draw one top-level menu bar item. lParam points at a
' UAHDRAWMENUITEM struct: DRAWITEMSTRUCT (48 bytes), then UAHMENU (12
' bytes), then UAHMENUITEM (iPosition is its first member).
'=========================================================================
Private Function DarkMenuBarItem(ByVal lParam As Long) As Long
Dim nDC As Long, hMenu As Long
Dim nState As Long, nPos As Long, nLen As Long, nFlags As Long
Dim tRC As RECT_DKM
Dim hBr As Long
Dim sText As String

Call CopyMemory(nState, ByVal (lParam + 16), 4&)   'DRAWITEMSTRUCT.itemState
Call CopyMemory(nDC, ByVal (lParam + 24), 4&)      'DRAWITEMSTRUCT.hDC
Call CopyMemory(tRC, ByVal (lParam + 28), 16&)     'DRAWITEMSTRUCT.rcItem
Call CopyMemory(hMenu, ByVal (lParam + 48), 4&)    'UAHMENU.hmenu
Call CopyMemory(nPos, ByVal (lParam + 60), 4&)     'UAHMENUITEM.iPosition
If nDC = 0 Then Exit Function

sText = String$(260, vbNullChar)
nLen = GetMenuStringA(hMenu, nPos, sText, 259, MF_BYPOSITION)
If nLen > 0 Then
    sText = Left$(sText, nLen)
Else
    sText = ""
End If

If (nState And (ODS_SELECTED Or ODS_HOTLIGHT)) Then
    hBr = CreateSolidBrush(DKMENU_HILITE)
Else
    hBr = CreateSolidBrush(DK_FORM_BACK)
End If
Call FillRect(nDC, tRC, hBr)
Call DeleteObject(hBr)

Call SetBkMode(nDC, BKMODE_TRANSPARENT)
If (nState And (ODS_GRAYED Or ODS_INACTIVE)) Then
    Call SetTextColor(nDC, DK_TEXT_DIM)
Else
    Call SetTextColor(nDC, DK_TEXT)
End If

nFlags = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
If (nState And ODS_NOACCEL) Then nFlags = nFlags Or DT_HIDEPREFIX
Call DrawTextA(nDC, sText, -1, tRC, nFlags)

DarkMenuBarItem = 1
End Function

'=========================================================================
' Paints over the light 1px line Windows leaves between the menu bar and
' the client area (called after default WM_NCPAINT/WM_NCACTIVATE).
'=========================================================================
Private Sub DarkMenuBarBottomLine(ByVal hWnd As Long)
Dim tMBI As MENUBARINFO
Dim tRC As RECT_DKM, tRW As RECT_DKM
Dim nDC As Long, hBr As Long

tMBI.cbSize = Len(tMBI)
If GetMenuBarInfo(hWnd, OBJID_MENU, 0, tMBI) = 0 Then Exit Sub

Call GetClientRect(hWnd, tRC)
Call MapWindowPoints(hWnd, 0, tRC, 2)
Call GetWindowRect(hWnd, tRW)
tRC.Left = tRC.Left - tRW.Left
tRC.Right = tRC.Right - tRW.Left
tRC.Top = tRC.Top - tRW.Top
tRC.Bottom = tRC.Top
tRC.Top = tRC.Top - 1

nDC = GetWindowDC(hWnd)
If nDC = 0 Then Exit Sub
hBr = CreateSolidBrush(DK_FORM_BACK)
Call FillRect(nDC, tRC, hBr)
Call DeleteObject(hBr)
Call ReleaseDC(hWnd, nDC)
End Sub
