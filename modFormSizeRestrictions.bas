Attribute VB_Name = "modFormSizeRestrictions"
Option Explicit
'Global nDebugHWND As Long
'
' Notes on subclassing with Comctl32.DLL:
'
'   1.  A subclassed function will get executed even AFTER the IDE "Stop" button is pressed.
'       This gives us an opportunity to un-subclass everything if things are done correctly.
'       Things that will still crash the IDE:
'
'       *   Executing the "END" statement in code.
'       *   Clicking IDE "Stop" on modal form loaded after something else is subclassed.
'       *   Clicking the "End" button after a runtime error on the "End", "Debug", "Help" form.
'
'   2.  "Each subclass is uniquely identified by the address of the pfnSubclass and its uIdSubclass"
'       (quote from Microsoft.com).
'
'   3.  For a particular hWnd, the last procedure subclassed will be the first to execute.
'
'   4.  If we call SetWindowSubclass repeatedly with the same hWnd, same pfnSubclass,
'       same uIdSubclass, and same dwRefData, it does nothing at all.
'       Not even the order of the subclassed functions will change,
'       even if other functions were subclassed later, and then SetWindowSubclass was
'       called again with the same hWnd, pfnSubclass, uIdSubclass, and dwRefData.
'
'   5.  Similar to the above, if we call SetWindowSubclass repeatedly,
'       and nothing changes but the dwRefData, the dwRefData is changed like we want,
'       but the order of execution of the functions still stays the same as it was.
'        "To change reference data you can make subsequent calls to SetWindowSubclass"
'       (quote from Microsoft.com).
'
'   6.  When un-subclassing, we can call RemoveWindowSubclass in any order we like, with no harm.
'
'   7.  We don't have to call DefSubclassProc in a particular subclassed function, but if we don't,
'       all other "downstream" subclassed functions won't execute.
'
'   8.  In the subclassed function, if uMsg = WM_DESTROY we should absolutely call
'       DefSubclassProc so that other possible "downstream" procedures can also un-subclassed.
'
'   9.  Things that are cleared BEFORE the subclass proc is executed again when the
'       IDE "Stop" button is clicked (i.e., before "uMsg = WM_DESTROY"):
'       *   All COM objects are uninstantiated (including Collections).
'       *   All dynamic arrays are erased.
'       *   All static arrays are reset (i.e., set to zero, vbNullString, etc.)
'       *   ALL variables are reset, including local Static variables.
'
'   10. Continuing on the above, even after all that is done, we can still make use of
'       variables, just recognizing that they'll be "fresh" variables.
'
'   11. The dwRefData can be used for whatever we want.  It's stored by Comctl32.DLL and is
'       returned everytime the subclassed procedure is called, or when explicitly requested by
'       a call to GetWindowSubclass.

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
Private Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoW" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Declare Function GetMenuItemRect Lib "user32" (ByVal hwnd As Long, _
    ByVal hMenu As Long, ByVal nPos As Long, lpRect As RECT) As Long
    
Private Declare Function SetWindowSubclass Lib "comctl32.dll" Alias "#410" ( _
    ByVal hwnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, _
    Optional ByVal dwRefData As Long) As Long
    
Private Declare Function RemoveWindowSubclass Lib "comctl32.dll" Alias "#412" ( _
    ByVal hwnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
    
Private Declare Function NextSubclassProcOnChain Lib "comctl32.dll" Alias "#413" ( _
    ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
Private Declare Function AdjustWindowRectEx Lib "user32.dll" ( _
    ByRef lpRect As RECT, _
    ByVal dwStyle As Long, _
    ByVal bMenu As Long, _
    ByVal dwExStyle As Long) As Long
    
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" (ByVal hwnd As Long, _
   ByVal nIndex As Long, ByVal dwNewLong As Long) _
   As Long

Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

'**************************************************************************************
' The following MODULE level stuff is specific to individual subclassing needs.
'**************************************************************************************
Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type WindowSizeRestrictions
    twpMinWidth As Long
    twpMaxWidth As Long
    twpMinHeight As Long
    twpMaxHeight As Long
    pxlMinWidth As Long
    pxlMaxWidth As Long
    pxlMinHeight As Long
    pxlMaxHeight As Long
    ScaleFactor As Single
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Private Enum MONITORFROMWINDOW_FLAGS
    MONITOR_DEFAULTTONULL = &H0& 'If the monitor is not found, return WIN32_NULL.
    MONITOR_DEFAULTTOPRIMARY = &H1& 'If the monitor is not found, return the primary monitor.
    MONITOR_DEFAULTTONEAREST = &H2& 'If the monitor is not found, return the nearest monitor.
End Enum

'Private Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type

Dim bSetWhenSubclassing_UsedByIdeStop As Boolean ' Never goes false once set by first subclassing, unless IDE Stop button is clicked.
Public gbAllowSubclassing As Boolean    ' Be sure to turn this on if you're going to use subclassing.

Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const WIN32_FALSE As Long = 0
Private Const WIN32_TRUE As Long = Not WIN32_FALSE
Private Const SPI_GETWORKAREA As Long = &H30
Private Const SM_CYSCREEN As Long = &H1
Private Const SM_CYCAPTION As Long = 4    ' Height of the window caption (title bar)
Private Const SM_CYMENU As Long = 15      ' Height of a single-line menu bar
Private Const SM_CXFRAME As Long = 32     ' Width of the sizing border for a resizable window
Private Const SM_CYFRAME As Long = 33     ' Height of the sizing border for a resizable window
Private Const GWL_STYLE As Long = (-16&)
Private Const GWL_EXSTYLE As Long = (-20&)
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const WM_DESTROY As Long = &H2&
Private Const WM_UAHDESTROYWINDOW As Long = &H90&
Private Const WM_GETMINMAXINFO As Long = &H24&
Private Const WM_DPICHANGED As Long = &H2E0

'**************************************************************************************
'**************************************************************************************
'**************************************************************************************

Private Sub SubclassSomeWindow(hwnd As Long, AddressOf_ProcToSubclass As Long, dwRefData As Long)
    ' This can be called AFTER the initial subclassing to update dwRefData.
    If Not gbAllowSubclassing Then Exit Sub
    bSetWhenSubclassing_UsedByIdeStop = True
    Call SetWindowSubclass(hwnd, AddressOf_ProcToSubclass, hwnd, dwRefData)
End Sub

Private Sub UnSubclassSomeWindow(hwnd As Long, AddressOf_ProcToSubclass As Long)
    Call RemoveWindowSubclass(hwnd, AddressOf_ProcToSubclass, hwnd)
End Sub

Private Function IdeStopButtonClicked() As Boolean
    ' The following works because all variables are cleared when the STOP button is clicked,
    ' even though other code may still execute such as Windows calling some of the subclassing procedures below.
    IdeStopButtonClicked = Not bSetWhenSubclassing_UsedByIdeStop
End Function

'**************************************************************************************
'**************************************************************************************
'**************************************************************************************

Public Sub SubclassFormMinMaxSize(frm As VB.Form, tMinMaxSize As WindowSizeRestrictions, Optional ByVal bFixToCurrentSize As Boolean)
Dim nPixelWidth As Long, nPixelHeight As Long, hdc As Long
' It's PIXELS.
'
' MUST be done in Form_Load event so Windows doesn't resize form on small monitors.
' Also, move (such as center) the form after calling so that WM_GETMINMAXINFO is fired.
' Can be called repeatedly to change MinWidth, MinHeight, MaxWidth, and MaxHeight with no harm done.
' Although, all must be supplied that you wish to maintain.
'
' Not supplying an argument (i.e., leaving it zero) will cause it to be ignored.

If tMinMaxSize.ScaleFactor = 0 Then
    hdc = GetDC(0)
    tMinMaxSize.ScaleFactor = GetDeviceCaps(frm.hdc, LOGPIXELSX) / 96
    hdc = ReleaseDC(0, hdc)
End If

If bFixToCurrentSize Then
    tMinMaxSize.twpMinWidth = frm.ScaleWidth
    tMinMaxSize.twpMinHeight = frm.ScaleHeight
    tMinMaxSize.twpMaxWidth = frm.ScaleWidth
    tMinMaxSize.twpMaxHeight = frm.ScaleHeight
    
    nPixelWidth = ConvertScale(frm.ScaleWidth, vbTwips, vbPixels, tMinMaxSize.ScaleFactor)
    nPixelHeight = ConvertScale(frm.ScaleHeight, vbTwips, vbPixels, tMinMaxSize.ScaleFactor)
    tMinMaxSize.pxlMinWidth = nPixelWidth
    tMinMaxSize.pxlMinHeight = nPixelHeight
    tMinMaxSize.pxlMaxWidth = nPixelWidth
    tMinMaxSize.pxlMaxHeight = nPixelHeight
Else
    With tMinMaxSize
        If .twpMinWidth And Not .pxlMinWidth Then .pxlMinWidth = ConvertScale(.twpMinWidth, vbTwips, vbPixels, tMinMaxSize.ScaleFactor)
        If .twpMinHeight And Not .pxlMinHeight Then .pxlMinHeight = ConvertScale(.twpMinHeight, vbTwips, vbPixels, tMinMaxSize.ScaleFactor)
        If .twpMaxWidth And Not .pxlMaxWidth Then .pxlMaxWidth = ConvertScale(.twpMaxWidth, vbTwips, vbPixels, tMinMaxSize.ScaleFactor)
        If .twpMaxHeight And Not .pxlMaxHeight Then .pxlMaxHeight = ConvertScale(.twpMaxHeight, vbTwips, vbPixels, tMinMaxSize.ScaleFactor)
    End With
End If

With tMinMaxSize
    If .pxlMinWidth > .pxlMaxWidth And .pxlMaxWidth <> 0 Then .pxlMaxWidth = .pxlMinWidth
    If .pxlMinHeight > .pxlMaxHeight And .pxlMaxHeight <> 0 Then .pxlMaxHeight = .pxlMinHeight
End With

SubclassSomeWindow frm.hwnd, AddressOf MinMaxSize_Proc, VarPtr(tMinMaxSize)

End Sub

Public Sub UN_SubclassFormSizeRestriction(frm As VB.Form)
    UnSubclassSomeWindow frm.hwnd, AddressOf MinMaxSize_Proc
End Sub

Private Function MinMaxSize_Proc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, dwRefData As WindowSizeRestrictions) As Long
'lParam As MINMAXINFO
Dim bProcessed As Boolean, mmi As MINMAXINFO, minRECT As RECT, maxRECT As RECT, hMenu As Long
Dim NewMinWidth As Long, NewMinHeight As Long, NewMaxWidth As Long, NewMaxHeight As Long, rMenu As RECT
Dim lNewDPI As Long, captionHeight As Long, menuHeight As Long, borderWidth As Long, borderHeight As Long

If IdeStopButtonClicked Then
    MinMaxSize_Proc = NextSubclassProcOnChain(hwnd, uMsg, wParam, lParam)
    Exit Function
End If

Select Case uMsg
    Case WM_GETMINMAXINFO
    
        hMenu = GetMenu(hwnd)
        If hMenu > 0 Then
            If GetMenuItemRect(hwnd, hMenu, 0, rMenu) Then
                hMenu = 1
                menuHeight = (rMenu.Bottom - rMenu.Top)
            Else
                hMenu = 0
            End If
        End If
        
        If dwRefData.ScaleFactor > 0 And dwRefData.ScaleFactor <> 1 Then ' And 1 = 2
            captionHeight = GetSystemMetrics(SM_CYCAPTION) '* dwRefData.ScaleFactor
            'If hMenu Then menuHeight = GetSystemMetrics(SM_CYMENU) * dwRefData.ScaleFactor
            'menuHeight = menuHeight * dwRefData.ScaleFactor
            borderWidth = GetSystemMetrics(SM_CXFRAME) '* dwRefData.ScaleFactor
            borderHeight = GetSystemMetrics(SM_CYFRAME) '* dwRefData.ScaleFactor
            
            If dwRefData.pxlMinWidth Or dwRefData.pxlMinHeight Then
                If dwRefData.pxlMinWidth Then minRECT.Right = dwRefData.pxlMinWidth
                If dwRefData.pxlMinHeight Then minRECT.Bottom = dwRefData.pxlMinHeight
                minRECT.Bottom = minRECT.Bottom + captionHeight + menuHeight + borderHeight * 2
                minRECT.Right = minRECT.Right + borderWidth * 2
            End If
            
            If dwRefData.pxlMaxWidth Or dwRefData.pxlMaxHeight Then
                If dwRefData.pxlMaxWidth Then maxRECT.Right = dwRefData.pxlMaxWidth
                If dwRefData.pxlMaxHeight Then maxRECT.Bottom = dwRefData.pxlMaxHeight
                maxRECT.Bottom = maxRECT.Bottom + captionHeight + menuHeight + borderHeight * 2
                maxRECT.Right = maxRECT.Right + borderWidth * 2
            End If
        Else
            If dwRefData.pxlMinWidth Or dwRefData.pxlMinHeight Then
                If dwRefData.pxlMinWidth Then minRECT.Right = dwRefData.pxlMinWidth
                If dwRefData.pxlMinHeight Then minRECT.Bottom = dwRefData.pxlMinHeight
                AdjustWindowRectEx minRECT, GetWindowLong(hwnd, GWL_STYLE), hMenu, GetWindowLong(hwnd, GWL_EXSTYLE)
            End If
    
            If dwRefData.pxlMaxWidth Or dwRefData.pxlMaxHeight Then
                If dwRefData.pxlMaxWidth Then maxRECT.Right = dwRefData.pxlMaxWidth
                If dwRefData.pxlMaxHeight Then maxRECT.Bottom = dwRefData.pxlMaxHeight
                AdjustWindowRectEx maxRECT, GetWindowLong(hwnd, GWL_STYLE), hMenu, GetWindowLong(hwnd, GWL_EXSTYLE)
            End If
        End If
        
        If dwRefData.pxlMinWidth Then NewMinWidth = minRECT.Right - minRECT.Left
        If dwRefData.pxlMinHeight Then NewMinHeight = minRECT.Bottom - minRECT.Top
        If dwRefData.pxlMaxWidth Then NewMaxWidth = maxRECT.Right - maxRECT.Left
        If dwRefData.pxlMaxHeight Then NewMaxHeight = maxRECT.Bottom - maxRECT.Top
        
        RtlMoveMemory mmi, ByVal lParam, LenB(mmi)
        
        If NewMinWidth And NewMinWidth <> mmi.ptMinTrackSize.x Then
            mmi.ptMinTrackSize.x = NewMinWidth
            bProcessed = True
        End If

        If NewMinHeight And NewMinHeight <> mmi.ptMinTrackSize.y Then
            mmi.ptMinTrackSize.y = NewMinHeight
            bProcessed = True
        End If

        If NewMaxWidth And (NewMaxWidth <> mmi.ptMaxTrackSize.x Or NewMaxWidth <> mmi.ptMaxSize.x) Then
            mmi.ptMaxTrackSize.x = NewMaxWidth
            mmi.ptMaxSize.x = NewMaxWidth
            bProcessed = True
        End If
        
        If NewMaxHeight And (NewMaxHeight <> mmi.ptMaxTrackSize.y Or NewMaxHeight <> mmi.ptMaxSize.y) Then
            mmi.ptMaxTrackSize.y = NewMaxHeight
            mmi.ptMaxSize.y = NewMaxHeight
            bProcessed = True
        End If
        
        If bProcessed Then
            RtlMoveMemory ByVal lParam, mmi, LenB(mmi)
            Exit Function
        End If
        
    Case WM_DESTROY, WM_UAHDESTROYWINDOW
        UnSubclassSomeWindow hwnd, AddressOf modFormSizeRestrictions.MinMaxSize_Proc
     
    Case WM_DPICHANGED
        bDPIAwareMode = True
        lNewDPI = wParam And &HFFFF&
        dwRefData.ScaleFactor = lNewDPI / 96
        If dwRefData.twpMinWidth Then dwRefData.pxlMinWidth = ConvertScale(dwRefData.twpMinWidth, vbTwips, vbPixels, dwRefData.ScaleFactor)
        If dwRefData.twpMinHeight Then dwRefData.pxlMinHeight = ConvertScale(dwRefData.twpMinHeight, vbTwips, vbPixels, dwRefData.ScaleFactor)
        If dwRefData.twpMaxWidth Then dwRefData.pxlMaxWidth = ConvertScale(dwRefData.twpMaxWidth, vbTwips, vbPixels, dwRefData.ScaleFactor)
        If dwRefData.twpMaxHeight Then dwRefData.pxlMinHeight = ConvertScale(dwRefData.twpMaxHeight, vbTwips, vbPixels, dwRefData.ScaleFactor)
        
End Select

MinMaxSize_Proc = NextSubclassProcOnChain(hwnd, uMsg, wParam, lParam)

End Function

'**************************************************************************************
'**************************************************************************************
'**************************************************************************************

Public Sub ResizeForm(frmHwnd As Long, nSetWidth As Long, nSetHeight As Long, Optional ByVal nScaleFactor As Single, Optional bAsPixels As Boolean)
On Error GoTo error:
Dim captionHeight As Long, menuHeight As Long, borderWidth As Long, borderHeight As Long
Dim currentStyle As Long, currentExStyle As Long, hMenu As Long
Dim rCurWindow As RECT, rNewWindow As RECT, rMenu As RECT

rNewWindow.Left = 0
rNewWindow.Top = 0

If bAsPixels Then
    rNewWindow.Right = nSetWidth
    rNewWindow.Bottom = nSetHeight
Else
    rNewWindow.Right = ConvertScale(nSetWidth, vbTwips, vbPixels, nScaleFactor)
    rNewWindow.Bottom = ConvertScale(nSetHeight, vbTwips, vbPixels, nScaleFactor)
End If

hMenu = GetMenu(frmHwnd)
If hMenu > 0 Then
    If GetMenuItemRect(frmHwnd, hMenu, 0, rMenu) Then
        hMenu = 1
        menuHeight = (rMenu.Bottom - rMenu.Top)
    Else
        hMenu = 0
    End If
End If

If nScaleFactor > 0 And nScaleFactor <> 1 Then ' And 1 = 2
    captionHeight = GetSystemMetrics(SM_CYCAPTION) '* nScaleFactor
    'If hMenu Then menuHeight = GetSystemMetrics(SM_CYMENU) * nScaleFactor
    'menuHeight = menuHeight * nScaleFactor
    borderWidth = GetSystemMetrics(SM_CXFRAME) '* nScaleFactor
    borderHeight = GetSystemMetrics(SM_CYFRAME) '* nScaleFactor
    rNewWindow.Bottom = rNewWindow.Bottom + captionHeight + menuHeight + borderHeight * 2
    rNewWindow.Right = rNewWindow.Right + borderWidth * 2
Else
    If hMenu > 0 Then hMenu = 1
    AdjustWindowRectEx rNewWindow, GetWindowLong(frmHwnd, GWL_STYLE), hMenu, GetWindowLong(frmHwnd, GWL_EXSTYLE)
End If

Call GetWindowRect(frmHwnd, rCurWindow)
'currentStyle = GetWindowLong(frmHwnd, GWL_STYLE)
'currentExStyle = GetWindowLong(frmHwnd, GWL_EXSTYLE)
'
'AdjustWindowRectEx rNewWindow, currentStyle, hMenu, currentExStyle

'frm.Width = ConvertScale(rNewWindow.Right - rNewWindow.Left, vbPixels, vbTwips)
'frm.Height = ConvertScale(rNewWindow.Bottom - rNewWindow.Top, vbPixels, vbTwips)

Call MoveWindow(frmHwnd, rCurWindow.Left, rCurWindow.Top, (rNewWindow.Right - rNewWindow.Left), (rNewWindow.Bottom - rNewWindow.Top), True)

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("ResizeForm")
Resume out:
End Sub

Public Sub ConvertFixedSizeForm(frm As VB.Form, Optional bAllowMaximize As Boolean)
On Error GoTo error:
Dim nWidth As Long, nHeight As Long ', ScaleFactor As Single

nWidth = frm.ScaleWidth
nHeight = frm.ScaleHeight
'ScaleFactor = GetDeviceCaps(frm.hDC, LOGPIXELSX)

If bAllowMaximize Then
    Call SetWindowLong(frm.hwnd, GWL_STYLE, GetWindowLong(frm.hwnd, GWL_STYLE) Xor _
        (WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX))
Else
    Call SetWindowLong(frm.hwnd, GWL_STYLE, GetWindowLong(frm.hwnd, GWL_STYLE) Xor _
        (WS_THICKFRAME Or WS_MINIMIZEBOX))
End If

Call SetWindowPos(frm.hwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOMOVE Or _
    SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)

'frm.BorderStyle = 2 'sizable
'frm.Caption = frm.Caption 'force redraw

Call ResizeForm(frm.hwnd, nWidth, nHeight)

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("ConvertFixedSizeForm")
Resume out:
End Sub


'**************************************************************************************
'**************************************************************************************
'**************************************************************************************

Public Function GetTwipsPerPixel(Optional ByVal nScaleFactor As Single) As String
On Error GoTo error:
Dim hdc As Long, lDPI_X As Long, lDPI_Y As Long, sngTPP_X As Single, sngTPP_Y As Single
Const HimetricPerPixel As Single = 26.45834

If nScaleFactor > 0 Then
    sngTPP_X = 1440 / (96 * nScaleFactor)
    sngTPP_Y = 1440 / (96 * nScaleFactor)
Else
    hdc = GetDC(0)
    lDPI_X = GetDeviceCaps(hdc, LOGPIXELSX): lDPI_Y = GetDeviceCaps(hdc, LOGPIXELSY)
    sngTPP_X = 1440 / lDPI_X: sngTPP_Y = 1440 / lDPI_Y
    hdc = ReleaseDC(0, hdc)
End If

GetTwipsPerPixel = "x-" & sngTPP_X & ", y-" & sngTPP_Y
out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetTwipsPerPixel")
Resume out:
End Function

Public Function ConvertScale(ByVal sngValue As Single, ByVal ScaleFrom As ScaleModeConstants, ByVal ScaleTo As ScaleModeConstants, Optional ByVal nScaleFactor As Single) As Single
On Error GoTo error:
Dim hdc As Long, lDPI_X As Long, lDPI_Y As Long, sngTPP_X As Single, sngTPP_Y As Single
Const HimetricPerPixel As Single = 26.45834

If nScaleFactor > 0 Then
    sngTPP_X = 1440 / (96 * nScaleFactor)
    sngTPP_Y = 1440 / (96 * nScaleFactor)
Else
    hdc = GetDC(0)
    lDPI_X = GetDeviceCaps(hdc, LOGPIXELSX): lDPI_Y = GetDeviceCaps(hdc, LOGPIXELSY)
    sngTPP_X = 1440 / lDPI_X: sngTPP_Y = 1440 / lDPI_Y
    hdc = ReleaseDC(0, hdc)
End If

Select Case True
    Case ScaleFrom = ScaleTo
        ConvertScale = sngValue
    Case (ScaleFrom = vbTwips) And (ScaleTo = vbPixels)
        ConvertScale = sngValue / sngTPP_X
    Case (ScaleFrom = vbPixels) And (ScaleTo = vbTwips)
        ConvertScale = sngValue * sngTPP_X
    Case (ScaleFrom = vbTwips) And (ScaleTo = vbPoints)
        ConvertScale = sngValue / 20
    Case (ScaleFrom = vbPoints) And (ScaleTo = vbTwips)
        ConvertScale = sngValue * 20
    Case (ScaleFrom = vbPixels) And (ScaleTo = vbPoints)
        ConvertScale = sngValue * sngTPP_X / 20
    Case (ScaleFrom = vbPoints) And (ScaleTo = vbPixels)
        ConvertScale = sngValue * 20 / sngTPP_X
    Case (ScaleFrom = vbPixels) And (ScaleTo = vbHimetric)
        ConvertScale = sngValue * HimetricPerPixel
    Case (ScaleFrom = vbHimetric) And (ScaleTo = vbPixels)
        ConvertScale = sngValue / HimetricPerPixel
End Select

out:
On Error Resume Next
Exit Function
error:
Call HandleError("ConvertScale")
Resume out:
End Function
