Attribute VB_Name = "modFormSizeRestrictions"
Option Explicit

Public gbAllowSubclassing As Boolean

Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
Private Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoW" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

'Private Declare Function GetDpiForWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
Private Declare Function MonitorFromWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal dwFlags As MONITORFROMWINDOW_FLAGS) As Long
Private Declare Function GetDpiForMonitor Lib "shcore.dll" (ByVal hMonitor As Long, ByVal dpiType As MonitorDpiTypeEnum, ByRef dpiX As Long, ByRef dpiY As Long) As Long
Private Declare Function AdjustWindowRectExForDpi Lib "user32.dll" (ByRef lpRect As RECT, ByVal dwStyle As Long, ByVal bMenu As Long, ByVal dwExStyle As Long, ByVal dpi As Long) As Long
Private Declare Function GetSystemMetricsForDpi Lib "user32.dll" (ByVal nIndex As Long, ByVal dpi As Long) As Long

Private Declare Function GetMenuItemRect Lib "user32" (ByVal hWnd As Long, _
    ByVal hMenu As Long, ByVal nPos As Long, lpRect As RECT) As Long
    
Private Declare Function SetWindowSubclass Lib "comctl32.dll" Alias "#410" ( _
    ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, _
    Optional ByVal dwRefData As Long) As Long
    
Private Declare Function RemoveWindowSubclass Lib "comctl32.dll" Alias "#412" ( _
    ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
    
Private Declare Function NextSubclassProcOnChain Lib "comctl32.dll" Alias "#413" ( _
    ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
Private Declare Function AdjustWindowRectEx Lib "user32.dll" ( _
    ByRef lpRect As RECT, _
    ByVal dwStyle As Long, _
    ByVal bMenu As Long, _
    ByVal dwExStyle As Long) As Long
    
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" (ByVal hWnd As Long, _
   ByVal nIndex As Long, ByVal dwNewLong As Long) _
   As Long

Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
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
    lastScaleFactor As Single
    primaryTPP As Single
End Type

Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Private Type WINDOWPOS
    hWnd As Long
    hWndInsertAfter As Long
    x As Long
    y As Long
    cx As Long
    cy As Long
    flags As Long
End Type

Private Enum MONITORFROMWINDOW_FLAGS
    MONITOR_DEFAULTTONULL = &H0& 'If the monitor is not found, return WIN32_NULL.
    MONITOR_DEFAULTTOPRIMARY = &H1& 'If the monitor is not found, return the primary monitor.
    MONITOR_DEFAULTTONEAREST = &H2& 'If the monitor is not found, return the nearest monitor.
End Enum

Private Enum MonitorDpiTypeEnum
    MDT_EFFECTIVE_DPI = 0 ' (default) The effective DPI (almost always 96). This value should be used when determining the correct scale factor for scaling UI elements. This incorporates the scale factor set by the user for this specific display.
    MDT_ANGULAR_DPI = 1   ' The angular DPI. This DPI ensures rendering at a compliant angular resolution on the screen for us. This does not include the scale factor set by the user for this specific display.
    MDT_RAW_DPI = 2       ' The raw DPI (PHYSICAL for monitor's dimensions, with Win10 scaling built-in). This value is the linear DPI of the screen as measured on the screen itself. Use this value when you want to read the pixel density and not the recommended scaling setting. This does not include the scale factor set by the user for this specific display and is not guaranteed to be a supported DPI value.
End Enum

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
Private Const SM_CXBORDER As Long = 5     ' The width of a window border, in pixels.
Private Const SM_CYBORDER As Long = 6     ' The height of a window border, in pixels.
Private Const SM_CXPADDEDBORDER = 92
Private Const GWL_STYLE As Long = (-16&)
Private Const GWL_EXSTYLE As Long = (-20&)
Private Const GWL_WNDPROC As Long = -4
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE As Long = &H1
Private Const WM_DESTROY As Long = &H2&
Private Const WM_NCDESTROY As Long = &H82
Private Const WM_UAHDESTROYWINDOW As Long = &H90&
Private Const WM_GETMINMAXINFO As Long = &H24&
Private Const WM_WINDOWPOSCHANGING As Long = &H46
Private Const WM_DPICHANGED As Long = &H2E0
Private Const HimetricPerPixel As Single = 26.45834

'**************************************************************************************
'**************************************************************************************
'**************************************************************************************

Private Sub UnSubclassSomeWindow(hWnd As Long, uIdSubclass As Long)
    Call RemoveWindowSubclass(hWnd, AddressOf MinMaxSize_Proc, uIdSubclass)
End Sub

Public Sub UN_SubclassFormSizeRestriction(frm As VB.Form)
    UnSubclassSomeWindow frm.hWnd, ObjPtr(frm)
End Sub

Private Sub SubclassSomeWindow(hWnd As Long, uIdSubclass As Long, dwRefData As Long)
    If Not gbAllowSubclassing Then Exit Sub
    Call SetWindowSubclass(hWnd, AddressOf MinMaxSize_Proc, uIdSubclass, dwRefData)
End Sub

Public Sub SubclassFormMinMaxSize(frm As VB.Form, tMinMaxSize As WindowSizeRestrictions, Optional ByVal bFixToCurrentSize As Boolean, Optional ByVal bUpdateOnly As Boolean)
Dim nPixelWidth As Long, nPixelHeight As Long, borderSize As Long, hMenu As Long, rMenu As RECT
Dim captionHeight As Long, menuHeight As Long, borderWidth As Long, borderHeight As Long, borderPad As Long
Dim nScreenTPPfactor As Single, rMinWindow As RECT, rMaxWindow As RECT, nWindowStyle As Long, nWindowStyleEx As Long
Dim rNonClientSize As RECT

nScreenTPPfactor = 15 / Screen.TwipsPerPixelX
If tMinMaxSize.ScaleFactor = 0 Then tMinMaxSize.ScaleFactor = GetDpiForWindow_Proxy(frm.hWnd) / 96

If nScreenTPPfactor <> 1 Or tMinMaxSize.ScaleFactor <> 1 Then bDPIAwareMode = True

If bFixToCurrentSize Then

    tMinMaxSize.twpMinWidth = frm.ScaleWidth
    tMinMaxSize.twpMinHeight = frm.ScaleHeight
    tMinMaxSize.twpMaxWidth = frm.ScaleWidth
    tMinMaxSize.twpMaxHeight = frm.ScaleHeight
    
    nPixelWidth = ConvertScale(frm.Width, vbTwips, vbPixels, nScreenTPPfactor)
    nPixelHeight = ConvertScale(frm.Height, vbTwips, vbPixels, nScreenTPPfactor)
    tMinMaxSize.pxlMinWidth = nPixelWidth
    tMinMaxSize.pxlMinHeight = nPixelHeight
    tMinMaxSize.pxlMaxWidth = nPixelWidth
    tMinMaxSize.pxlMaxHeight = nPixelHeight
    
Else
    
    hMenu = GetMenu(frm.hWnd)
    If hMenu > 0 Then
        If GetMenuItemRect(frm.hWnd, hMenu, 0, rMenu) Then 'does the menu have size to it (e.g. not hidden)
            hMenu = 1
            menuHeight = (rMenu.Bottom - rMenu.Top)
        Else
            hMenu = 0
        End If
    End If
    
    nWindowStyle = GetWindowLong(frm.hWnd, GWL_STYLE)
    nWindowStyleEx = GetWindowLong(frm.hWnd, GWL_EXSTYLE)
    AdjustWindowRectExForDpi_Proxy rNonClientSize, nWindowStyle, hMenu, nWindowStyleEx, frm.hWnd, (96 * tMinMaxSize.ScaleFactor)
    
    captionHeight = Abs(rNonClientSize.Top) - rNonClientSize.Bottom
    borderHeight = rNonClientSize.Bottom * 2
    borderWidth = rNonClientSize.Right * 2
    
    If tMinMaxSize.twpMinWidth > 0 Or tMinMaxSize.twpMinHeight > 0 Then
        nPixelWidth = 0: nPixelHeight = 0
        If tMinMaxSize.twpMinWidth > 0 Then nPixelWidth = ConvertScale(tMinMaxSize.twpMinWidth, vbTwips, vbPixels, nScreenTPPfactor) + borderWidth
        If tMinMaxSize.twpMinHeight > 0 Then nPixelHeight = ConvertScale(tMinMaxSize.twpMinHeight, vbTwips, vbPixels, nScreenTPPfactor) + borderHeight + captionHeight
        rMinWindow.Right = nPixelWidth
        rMinWindow.Bottom = nPixelHeight
    End If
    
    If tMinMaxSize.twpMinWidth = tMinMaxSize.twpMaxWidth And tMinMaxSize.twpMinHeight = tMinMaxSize.twpMaxHeight Then
        rMaxWindow = rMinWindow
    ElseIf tMinMaxSize.twpMaxWidth > 0 Or tMinMaxSize.twpMaxHeight > 0 Then
        nPixelWidth = 0: nPixelHeight = 0
        If tMinMaxSize.twpMaxWidth > 0 Then nPixelWidth = ConvertScale(tMinMaxSize.twpMaxWidth, vbTwips, vbPixels, nScreenTPPfactor) + borderWidth
        If tMinMaxSize.twpMaxHeight > 0 Then nPixelHeight = ConvertScale(tMinMaxSize.twpMaxHeight, vbTwips, vbPixels, nScreenTPPfactor) + borderHeight + captionHeight
        rMaxWindow.Right = nPixelWidth
        rMaxWindow.Bottom = nPixelHeight
    End If
    
    With tMinMaxSize
        .pxlMinWidth = 0
        .pxlMinHeight = 0
        .pxlMaxWidth = 0
        .pxlMaxHeight = 0
        If .twpMinWidth Then .pxlMinWidth = rMinWindow.Right - rMinWindow.Left
        If .twpMinHeight Then .pxlMinHeight = rMinWindow.Bottom - rMinWindow.Top
        If .twpMaxWidth Then .pxlMaxWidth = rMaxWindow.Right - rMaxWindow.Left
        If .twpMaxHeight Then .pxlMaxHeight = rMaxWindow.Bottom - rMaxWindow.Top
    End With
    
End If

With tMinMaxSize
    If .pxlMinWidth > .pxlMaxWidth And .pxlMaxWidth <> 0 Then .pxlMaxWidth = .pxlMinWidth
    If .pxlMinHeight > .pxlMaxHeight And .pxlMaxHeight <> 0 Then .pxlMaxHeight = .pxlMinHeight
End With

tMinMaxSize.lastScaleFactor = tMinMaxSize.ScaleFactor
tMinMaxSize.primaryTPP = Screen.TwipsPerPixelX
If bUpdateOnly Then Exit Sub

SubclassSomeWindow frm.hWnd, ObjPtr(frm), VarPtr(tMinMaxSize)

End Sub

'**************************************************************************************
'**************************************************************************************
'**************************************************************************************

Private Function MinMaxSize_Proc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal objForm As Form, dwRefData As WindowSizeRestrictions) As Long
Dim bProcessed As Boolean, mmi As MINMAXINFO, hMenu As Long
Dim NewMinWidth As Long, NewMinHeight As Long, NewMaxWidth As Long, NewMaxHeight As Long
Dim lNewDPI As Long, captionHeight As Long, menuHeight As Long, borderWidth As Long, borderHeight As Long
Dim r As RECT, rMenu As RECT, minRECT As RECT, maxRECT As RECT, wp As WINDOWPOS
Dim borderSize As Long, nTwipWidth As Long, nTwipHeight As Long

Select Case uMsg
    Case WM_GETMINMAXINFO
        
        If dwRefData.ScaleFactor <> dwRefData.lastScaleFactor Or dwRefData.primaryTPP <> Screen.TwipsPerPixelX Then
            SubclassFormMinMaxSize objForm, dwRefData, False, True
        End If
        
        If dwRefData.pxlMinWidth Then NewMinWidth = dwRefData.pxlMinWidth
        If dwRefData.pxlMinHeight Then NewMinHeight = dwRefData.pxlMinHeight
        If dwRefData.pxlMaxWidth Then NewMaxWidth = dwRefData.pxlMaxWidth
        If dwRefData.pxlMaxHeight Then NewMaxHeight = dwRefData.pxlMaxHeight
        
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
        
    Case WM_DESTROY, WM_UAHDESTROYWINDOW, WM_NCDESTROY
        UnSubclassSomeWindow hWnd, ObjPtr(objForm)
    
    Case WM_WINDOWPOSCHANGING
        'this would prevent resizing when dragging across screens
'        If dwRefData.ScaleFactor <> dwRefData.lastScaleFactor And bDPIAwareMode Then
'            RtlMoveMemory wp, ByVal lParam, LenB(wp)
'            wp.flags = wp.flags Or SWP_NOSIZE
'            RtlMoveMemory ByVal lParam, wp, LenB(wp)
'        End If
        
    Case WM_DPICHANGED
        bDPIAwareMode = True
        lNewDPI = wParam And &HFFFF&
        dwRefData.ScaleFactor = lNewDPI / 96
        
End Select

MinMaxSize_Proc = NextSubclassProcOnChain(hWnd, uMsg, wParam, lParam)

End Function

'**************************************************************************************
'**************************************************************************************
'**************************************************************************************

Public Sub ResizeForm(frm As VB.Form, nSetClientWidthTwips As Long, nSetClientHeightTwips As Long, Optional ByVal nScaleFactor As Single, Optional ByVal bByTwips As Boolean)
On Error GoTo error:
Dim captionHeight As Long, menuHeight As Long, borderWidth As Long, borderHeight As Long, borderPad As Long
Dim TWIPcaptionHeight As Long, TWIPmenuHeight As Long, TWIPborderWidth As Long, TWIPwidth As Long, TWIPheight As Long
Dim AdjDPIcaptionHeight As Long, AdjDPIborderWidth As Long, AdjDPIborderHeight As Long, AdjDPIwidth As Long, AdjDPIheight As Long
Dim gsmDPIcaptionHeight As Long, gsmDPIborderWidth As Long, gsmDPIborderHeight As Long, gsmDPIwidth As Long, gsmDPIheight As Long, gsmDPIborderPad As Long
Dim nDPI As Integer, hMenu As Long, nPxlWidth As Single, nPxlHeight As Single
Dim rCurWindow As RECT, rNewWindow As RECT, rMenu As RECT
Dim AdjDPIrNewWindow As RECT, gsmDPIrNewWindow As RECT, rMon As RECT
Dim nScreenTPPfactor As Single

nScreenTPPfactor = 15 / Screen.TwipsPerPixelX

rNewWindow.Left = 0
rNewWindow.Top = 0

If bByTwips Then
    rNewWindow.Right = nSetClientWidthTwips
    rNewWindow.Bottom = nSetClientHeightTwips
Else
    rNewWindow.Right = ConvertScale(nSetClientWidthTwips, vbTwips, vbPixels, nScreenTPPfactor)
    rNewWindow.Bottom = ConvertScale(nSetClientHeightTwips, vbTwips, vbPixels, nScreenTPPfactor)
    
    hMenu = GetMenu(frm.hWnd)
    If hMenu > 0 Then
        If GetMenuItemRect(frm.hWnd, hMenu, 0, rMenu) Then
            hMenu = 1
            menuHeight = (rMenu.Bottom - rMenu.Top)
        Else
            hMenu = 0
        End If
    End If
End If

If bByTwips Then
    borderWidth = (frm.Width - frm.ScaleWidth) / 2
    borderHeight = borderWidth
    captionHeight = (frm.Height - frm.ScaleHeight) - (borderHeight * 2)
    rNewWindow.Top = 0 - (captionHeight + borderHeight)
    rNewWindow.Left = 0 - borderWidth
    rNewWindow.Bottom = rNewWindow.Bottom + borderHeight
    rNewWindow.Right = rNewWindow.Right + borderWidth
    
ElseIf nScaleFactor > 0 And nScaleFactor <> 1 Then
    AdjustWindowRectExForDpi_Proxy rNewWindow, GetWindowLong(frm.hWnd, GWL_STYLE), hMenu, GetWindowLong(frm.hWnd, GWL_EXSTYLE), frm.hWnd, (96 * nScaleFactor)
    
Else
    AdjustWindowRectEx rNewWindow, GetWindowLong(frm.hWnd, GWL_STYLE), hMenu, GetWindowLong(frm.hWnd, GWL_EXSTYLE)
End If

If bByTwips Then
    frm.Width = (rNewWindow.Right - rNewWindow.Left)
    frm.Height = (rNewWindow.Bottom - rNewWindow.Top)
Else
    Call GetWindowRect(frm.hWnd, rCurWindow)
    Call MoveWindow(frm.hWnd, rCurWindow.Left, rCurWindow.Top, rNewWindow.Right - rNewWindow.Left, rNewWindow.Bottom - rNewWindow.Top, True)
End If

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("ResizeForm")
Resume out:
End Sub

Public Sub ConvertFixedSizeForm(frm As VB.Form, Optional bAllowMaximize As Boolean)
'this will convert a fixed single (1) form to sizable (2) at runtime.
On Error GoTo error:
Dim nWidth As Long, nHeight As Long

nWidth = frm.ScaleWidth
nHeight = frm.ScaleHeight

If bAllowMaximize Then
    Call SetWindowLong(frm.hWnd, GWL_STYLE, GetWindowLong(frm.hWnd, GWL_STYLE) Xor _
        (WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX))
Else
    Call SetWindowLong(frm.hWnd, GWL_STYLE, GetWindowLong(frm.hWnd, GWL_STYLE) Xor _
        (WS_THICKFRAME Or WS_MINIMIZEBOX))
End If

Call SetWindowPos(frm.hWnd, 0&, 0&, 0&, 0&, 0&, SWP_NOMOVE Or _
    SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)

Call ResizeForm(frm, nWidth, nHeight)

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

Public Function GetTwipsPerPixel(Optional ByVal nScaleFactor As Single) As Long
On Error GoTo error:
Dim hdc As Long, lDPI_X As Long, sngTPP_X As Single
Const HimetricPerPixel As Single = 26.45834

If nScaleFactor > 0 Then
    sngTPP_X = 1440 / (96 * nScaleFactor)
Else
    hdc = GetDC(0)
    lDPI_X = GetDeviceCaps(hdc, LOGPIXELSX)
    sngTPP_X = 1440 / lDPI_X
    hdc = ReleaseDC(0, hdc)
End If

GetTwipsPerPixel = sngTPP_X

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetTwipsPerPixel")
Resume out:
End Function

Public Function ConvertScale(ByVal sngValue As Single, ByVal ScaleFrom As ScaleModeConstants, ByVal ScaleTo As ScaleModeConstants, Optional ByVal nScaleFactor As Single) As Single
On Error GoTo error:
Dim hdc As Long, lDPI_X As Long, sngTPP_X As Single, nScreenTPPfactor As Single

If nScaleFactor > 0 Then
    sngTPP_X = 1440 / (96 * nScaleFactor)
Else
    hdc = GetDC(0)
    lDPI_X = GetDeviceCaps(hdc, LOGPIXELSX)
    sngTPP_X = 1440 / lDPI_X
    hdc = ReleaseDC(0, hdc)
End If

If ScaleFrom = vbPixels And ScaleTo = vbPixels Then 'convert to primary monitor scale factor
    nScreenTPPfactor = 15 / Screen.TwipsPerPixelX
    If nScaleFactor <> nScreenTPPfactor Then
        sngValue = sngValue / nScaleFactor 'convert pixels from current scalefactor back to 1.0
        sngValue = sngValue * nScreenTPPfactor 'convert to primary monitor scalefactor
        ConvertScale = sngValue
        Exit Function
    End If
End If

Select Case True
    Case ScaleFrom = ScaleTo
        ConvertScale = sngValue
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

Public Function AdjustWindowRectExForDpi_Proxy(ByRef lpRect As RECT, ByVal dwStyle As Long, ByVal bMenu As Long, ByVal dwExStyle As Long, Optional frmHwnd As Long, Optional ByVal nDPI As Integer) As Long
On Error GoTo error:
Dim Ret As Long, rNonClientSize As RECT, captionHeight As Long, borderWidth As Long, borderHeight As Long
Dim nScaleFactor As Single

If nDPI = 0 And frmHwnd > 0 Then nDPI = GetDpiForWindow_Proxy(frmHwnd)
If nDPI = 0 Then nDPI = 96
nScaleFactor = nDPI / 96

If nOSversion < Win10 Or nDPI = 96 Then
    If nDPI = 96 Then
        AdjustWindowRectExForDpi_Proxy = AdjustWindowRectEx(lpRect, dwStyle, bMenu, dwExStyle)
    Else
        Ret = AdjustWindowRectEx(rNonClientSize, dwStyle, bMenu, dwExStyle)
        If Ret = 0 Then Exit Function
        
        captionHeight = Abs(rNonClientSize.Top) - rNonClientSize.Bottom
        borderHeight = rNonClientSize.Bottom
        borderWidth = rNonClientSize.Right
        
        lpRect.Left = lpRect.Left - (borderWidth * nScaleFactor)
        lpRect.Right = lpRect.Right + (borderWidth * nScaleFactor)
        lpRect.Top = lpRect.Top - ((captionHeight + borderHeight) * nScaleFactor)
        lpRect.Bottom = lpRect.Bottom + (borderHeight * nScaleFactor)
        
        AdjustWindowRectExForDpi_Proxy = Ret
    End If
Else
    AdjustWindowRectExForDpi_Proxy = AdjustWindowRectExForDpi(lpRect, dwStyle, bMenu, dwExStyle, nDPI)
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("AdjustWindowRectExForDpi_Proxy")
Resume out:
End Function

Public Function GetDpiForWindow_Proxy(ByVal hWnd As Long) As Long
On Error GoTo error:
Dim hMonitor As Long, dpiX As Long, dpiY As Long, OSVer As cnWin32Ver, hdc As Long

If nOSversion < Win8_1 Then GoTo default:

hMonitor = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST)
GetDpiForMonitor hMonitor, MDT_EFFECTIVE_DPI, dpiX, dpiY

If dpiX < 24 Then GoTo default:
GetDpiForWindow_Proxy = dpiX
GoTo out:

default:
hdc = GetDC(0)
GetDpiForWindow_Proxy = GetDeviceCaps(hdc, LOGPIXELSX)
hdc = ReleaseDC(0, hdc)

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetDpiForWindow_Proxy")
Resume out:
End Function

Public Function GetMonitorDimensions(ByVal hWnd As Long) As RECT
On Error GoTo error:
Dim hMon As Long, mi As MONITORINFO

mi.cbSize = Len(mi)

hMon = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST)
If hMon <> 0 Then
    If GetMonitorInfo(hMon, mi) <> 0 Then
        GetMonitorDimensions.Right = mi.rcMonitor.Right
        GetMonitorDimensions.Left = mi.rcMonitor.Left
        GetMonitorDimensions.Bottom = mi.rcMonitor.Bottom
        GetMonitorDimensions.Top = mi.rcMonitor.Top
    End If
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetMonitorDimensions")
Resume out:
End Function
