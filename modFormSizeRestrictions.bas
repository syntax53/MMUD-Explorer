Attribute VB_Name = "modFormSizeRestrictions"
Option Explicit

Public gbAllowSubclassing As Boolean    ' Be sure to turn this on if you're going to use subclassing.

Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
Private Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoW" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

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
    lastScaleFactor As Single
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

'Private Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type

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

'**************************************************************************************
'**************************************************************************************
'**************************************************************************************

Private Sub SubclassSomeWindow(hWnd As Long, uIdSubclass As Long, dwRefData As Long)
    ' This can be called AFTER the initial subclassing to update dwRefData.
    If Not gbAllowSubclassing Then Exit Sub
    Call SetWindowSubclass(hWnd, AddressOf MinMaxSize_Proc, uIdSubclass, dwRefData)
End Sub

Private Sub UnSubclassSomeWindow(hWnd As Long, uIdSubclass As Long)
    Call RemoveWindowSubclass(hWnd, AddressOf MinMaxSize_Proc, uIdSubclass)
End Sub

'**************************************************************************************
'**************************************************************************************
'**************************************************************************************

Public Sub SubclassFormMinMaxSize(frm As VB.Form, tMinMaxSize As WindowSizeRestrictions, Optional ByVal bFixToCurrentSize As Boolean, Optional ByVal bUpdateOnly As Boolean)
Dim nPixelWidth As Long, nPixelHeight As Long, borderSize As Long, hMenu As Long, rMenu As RECT
Dim captionHeight As Long, menuHeight As Long, borderWidth As Long, borderHeight As Long, borderPad As Long
Dim PrimaryTPP As Single

PrimaryTPP = 15 / Screen.TwipsPerPixelX

If tMinMaxSize.ScaleFactor = 0 Then tMinMaxSize.ScaleFactor = GetDpiForWindowProxy(frm.hWnd) / 96

If bFixToCurrentSize Then
'    tMinMaxSize.twpMinWidth = frm.ScaleWidth
'    tMinMaxSize.twpMinHeight = frm.ScaleHeight
'    tMinMaxSize.twpMaxWidth = frm.ScaleWidth
'    tMinMaxSize.twpMaxHeight = frm.ScaleHeight
'
'    nPixelWidth = ConvertScale(frm.ScaleWidth, vbTwips, vbPixels, tMinMaxSize.ScaleFactor)
'    nPixelHeight = ConvertScale(frm.ScaleHeight, vbTwips, vbPixels, tMinMaxSize.ScaleFactor)
'    tMinMaxSize.pxlMinWidth = nPixelWidth
'    tMinMaxSize.pxlMinHeight = nPixelHeight
'    tMinMaxSize.pxlMaxWidth = nPixelWidth
'    tMinMaxSize.pxlMaxHeight = nPixelHeight
    tMinMaxSize.twpMinWidth = frm.ScaleWidth
    tMinMaxSize.twpMinHeight = frm.ScaleHeight
    tMinMaxSize.twpMaxWidth = frm.ScaleWidth
    tMinMaxSize.twpMaxHeight = frm.ScaleHeight
    
    nPixelWidth = ConvertScale(frm.Width, vbTwips, vbPixels, PrimaryTPP)
    nPixelHeight = ConvertScale(frm.Height, vbTwips, vbPixels, PrimaryTPP)
    tMinMaxSize.pxlMinWidth = nPixelWidth
    tMinMaxSize.pxlMinHeight = nPixelHeight
    tMinMaxSize.pxlMaxWidth = nPixelWidth
    tMinMaxSize.pxlMaxHeight = nPixelHeight
Else
    'borderSize = (frm.Width - frm.ScaleWidth)
    'captionHeight = (frm.Height - frm.ScaleHeight) - borderSize
    
    captionHeight = GetSystemMetrics(SM_CYCAPTION) * tMinMaxSize.ScaleFactor
    borderPad = GetSystemMetrics(SM_CXPADDEDBORDER)
    borderWidth = (borderPad + GetSystemMetrics(SM_CXFRAME)) * tMinMaxSize.ScaleFactor
    borderHeight = (borderPad + GetSystemMetrics(SM_CYFRAME)) * tMinMaxSize.ScaleFactor
    
    hMenu = GetMenu(frm.hWnd)
    If hMenu > 0 Then
        If GetMenuItemRect(frm.hWnd, hMenu, 0, rMenu) Then
            hMenu = 1
            menuHeight = (rMenu.Bottom - rMenu.Top)
        Else
            hMenu = 0
        End If
    End If
    nPixelWidth = (borderWidth * 2)
    nPixelHeight = captionHeight + menuHeight + (borderHeight * 2)
    
    With tMinMaxSize
        .pxlMinWidth = 0
        .pxlMinHeight = 0
        .pxlMaxWidth = 0
        .pxlMaxHeight = 0
'        If .twpMinWidth Then .pxlMinWidth = ConvertScale(.twpMinWidth + borderSize, vbTwips, vbPixels, 1)
'        If .twpMinHeight Then .pxlMinHeight = ConvertScale(.twpMinHeight + borderSize + captionHeight, vbTwips, vbPixels, 1)
'        If .twpMaxWidth Then .pxlMaxWidth = ConvertScale(.twpMaxWidth + borderSize, vbTwips, vbPixels, 1)
'        If .twpMaxHeight Then .pxlMaxHeight = ConvertScale(.twpMaxHeight + borderSize + captionHeight, vbTwips, vbPixels, 1)
        If .twpMinWidth Then .pxlMinWidth = ConvertScale(.twpMinWidth, vbTwips, vbPixels, PrimaryTPP) + nPixelWidth
        If .twpMinHeight Then .pxlMinHeight = ConvertScale(.twpMinHeight, vbTwips, vbPixels, PrimaryTPP) + nPixelHeight
        If .twpMaxWidth Then .pxlMaxWidth = ConvertScale(.twpMaxWidth, vbTwips, vbPixels, PrimaryTPP) + nPixelWidth
        If .twpMaxHeight Then .pxlMaxHeight = ConvertScale(.twpMaxHeight, vbTwips, vbPixels, PrimaryTPP) + nPixelHeight
    End With
End If

With tMinMaxSize
    If .pxlMinWidth > .pxlMaxWidth And .pxlMaxWidth <> 0 Then .pxlMaxWidth = .pxlMinWidth
    If .pxlMinHeight > .pxlMaxHeight And .pxlMaxHeight <> 0 Then .pxlMaxHeight = .pxlMinHeight
End With

tMinMaxSize.lastScaleFactor = tMinMaxSize.ScaleFactor
If bUpdateOnly Then Exit Sub

SubclassSomeWindow frm.hWnd, ObjPtr(frm), VarPtr(tMinMaxSize)

End Sub

Public Sub UN_SubclassFormSizeRestriction(frm As VB.Form)
    UnSubclassSomeWindow frm.hWnd, ObjPtr(frm)
End Sub

Private Function MinMaxSize_Proc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal objForm As Form, dwRefData As WindowSizeRestrictions) As Long
'lParam As MINMAXINFO
Dim bProcessed As Boolean, mmi As MINMAXINFO, hMenu As Long
Dim NewMinWidth As Long, NewMinHeight As Long, NewMaxWidth As Long, NewMaxHeight As Long
Dim lNewDPI As Long, captionHeight As Long, menuHeight As Long, borderWidth As Long, borderHeight As Long
Dim R As RECT, rMenu As RECT, minRECT As RECT, maxRECT As RECT, wp As WINDOWPOS
Dim borderSize As Long, nTwipWidth As Long, nTwipHeight As Long

Select Case uMsg
    Case WM_GETMINMAXINFO
        
'                                '        borderSize = (objForm.Width - objForm.ScaleWidth) / 2
'                                '        captionHeight = (objForm.Height - objForm.ScaleHeight) - (borderSize * 2)
'                                '
'                                '        If dwRefData.twpMinWidth Or dwRefData.twpMinHeight Then
'                                '            If dwRefData.twpMinWidth Then minRECT.Right = dwRefData.twpMinWidth
'                                '            If dwRefData.twpMinHeight Then minRECT.Bottom = dwRefData.twpMinHeight
'                                '            AdjustWindowRectEx minRECT, GetWindowLong(hWnd, GWL_STYLE), hMenu, GetWindowLong(hWnd, GWL_EXSTYLE)
'                                '        End If
'                                '
'                                '        If dwRefData.pxlMinWidth = dwRefData.pxlMaxWidth And dwRefData.pxlMinHeight = dwRefData.pxlMaxHeight Then
'                                '            maxRECT = minRECT
'                                '        ElseIf dwRefData.pxlMaxWidth Or dwRefData.pxlMaxHeight Then
'                                '            If dwRefData.pxlMaxWidth Then maxRECT.Right = dwRefData.pxlMaxWidth
'                                '            If dwRefData.pxlMaxHeight Then maxRECT.Bottom = dwRefData.pxlMaxHeight
'                                '            AdjustWindowRectEx maxRECT, GetWindowLong(hWnd, GWL_STYLE), hMenu, GetWindowLong(hWnd, GWL_EXSTYLE)
'                                '        End If
        
'        hMenu = GetMenu(hWnd)
'        If hMenu > 0 Then
'            If GetMenuItemRect(hWnd, hMenu, 0, rMenu) Then
'                hMenu = 1
'                menuHeight = (rMenu.Bottom - rMenu.Top)
'            Else
'                hMenu = 0
'            End If
'        End If
'
'        If dwRefData.ScaleFactor > 0 And dwRefData.ScaleFactor <> 1 Then ' And 1 = 2
'            captionHeight = GetSystemMetrics(SM_CYCAPTION) * dwRefData.ScaleFactor
'            'If hMenu Then menuHeight = GetSystemMetrics(SM_CYMENU) * dwRefData.ScaleFactor
'            'menuHeight = menuHeight * dwRefData.ScaleFactor
'            borderWidth = GetSystemMetrics(SM_CXFRAME) * dwRefData.ScaleFactor
'            borderHeight = GetSystemMetrics(SM_CYFRAME) * dwRefData.ScaleFactor
'
'            If dwRefData.pxlMinWidth Or dwRefData.pxlMinHeight Then
'                If dwRefData.pxlMinWidth Then minRECT.Right = dwRefData.pxlMinWidth
'                If dwRefData.pxlMinHeight Then minRECT.Bottom = dwRefData.pxlMinHeight
'                minRECT.Bottom = minRECT.Bottom + captionHeight + menuHeight + borderHeight * 2
'                minRECT.Right = minRECT.Right + borderWidth * 2
'            End If
'
'            If dwRefData.pxlMaxWidth Or dwRefData.pxlMaxHeight Then
'                If dwRefData.pxlMaxWidth Then maxRECT.Right = dwRefData.pxlMaxWidth
'                If dwRefData.pxlMaxHeight Then maxRECT.Bottom = dwRefData.pxlMaxHeight
'                maxRECT.Bottom = maxRECT.Bottom + captionHeight + menuHeight + borderHeight * 2
'                maxRECT.Right = maxRECT.Right + borderWidth * 2
'            End If
'        Else
'            If dwRefData.pxlMinWidth Or dwRefData.pxlMinHeight Then
'                If dwRefData.pxlMinWidth Then minRECT.Right = dwRefData.pxlMinWidth
'                If dwRefData.pxlMinHeight Then minRECT.Bottom = dwRefData.pxlMinHeight
'                AdjustWindowRectEx minRECT, GetWindowLong(hWnd, GWL_STYLE), hMenu, GetWindowLong(hWnd, GWL_EXSTYLE)
'            End If
'
'            If dwRefData.pxlMinWidth = dwRefData.pxlMaxWidth And dwRefData.pxlMinHeight = dwRefData.pxlMaxHeight Then
'                maxRECT = minRECT
'            ElseIf dwRefData.pxlMaxWidth Or dwRefData.pxlMaxHeight Then
'                If dwRefData.pxlMaxWidth Then maxRECT.Right = dwRefData.pxlMaxWidth
'                If dwRefData.pxlMaxHeight Then maxRECT.Bottom = dwRefData.pxlMaxHeight
'                AdjustWindowRectEx maxRECT, GetWindowLong(hWnd, GWL_STYLE), hMenu, GetWindowLong(hWnd, GWL_EXSTYLE)
'            End If
'        End If
'
'        If dwRefData.pxlMinWidth Then NewMinWidth = minRECT.Right - minRECT.Left
'        If dwRefData.pxlMinHeight Then NewMinHeight = minRECT.Bottom - minRECT.Top
'        If dwRefData.pxlMaxWidth Then NewMaxWidth = maxRECT.Right - maxRECT.Left
'        If dwRefData.pxlMaxHeight Then NewMaxHeight = maxRECT.Bottom - maxRECT.Top
        
        If dwRefData.ScaleFactor <> dwRefData.lastScaleFactor Then
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
        'prevent resizing when dragging across screens
'        If dwRefData.ScaleFactor <> dwRefData.lastScaleFactor And bDPIAwareMode Then
'            RtlMoveMemory wp, ByVal lParam, LenB(wp)
'            wp.flags = wp.flags Or SWP_NOSIZE
'            RtlMoveMemory ByVal lParam, wp, LenB(wp)
'        End If
        
    Case WM_DPICHANGED
        bDPIAwareMode = True
        lNewDPI = wParam And &HFFFF&
        dwRefData.ScaleFactor = lNewDPI / 96
        'UnSubclassSomeWindow hWnd, ObjPtr(objForm)
        'SubclassFormMinMaxSize objForm, dwRefData
'        RtlMoveMemory R, ByVal lParam, LenB(R)
'        With R
'            SetWindowPos hWnd, 0, .Left, .Top, .Right - .Left, .Bottom - .Top, SWP_NOACTIVATE Or SWP_NOOWNERZORDER Or SWP_NOZORDER ' Resize the form to reflect the new DPI changes
'        End With
        'Exit Function
        'If dwRefData.twpMinWidth Then dwRefData.pxlMinWidth = ConvertScale(dwRefData.twpMinWidth, vbTwips, vbPixels, dwRefData.ScaleFactor)
        'If dwRefData.twpMinHeight Then dwRefData.pxlMinHeight = ConvertScale(dwRefData.twpMinHeight, vbTwips, vbPixels, dwRefData.ScaleFactor)
        'If dwRefData.twpMaxWidth Then dwRefData.pxlMaxWidth = ConvertScale(dwRefData.twpMaxWidth, vbTwips, vbPixels, dwRefData.ScaleFactor)
        'If dwRefData.twpMaxHeight Then dwRefData.pxlMaxHeight = ConvertScale(dwRefData.twpMaxHeight, vbTwips, vbPixels, dwRefData.ScaleFactor)
        
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

'...Dim currentStyle As Long, currentExStyle As Long

rNewWindow.Left = 0
rNewWindow.Top = 0

If bByTwips Then
    rNewWindow.Right = nSetClientWidthTwips
    rNewWindow.Bottom = nSetClientHeightTwips
Else
    rNewWindow.Right = ConvertScale(nSetClientWidthTwips, vbTwips, vbPixels, nScaleFactor)
    rNewWindow.Bottom = ConvertScale(nSetClientHeightTwips, vbTwips, vbPixels, nScaleFactor)
    
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

rMon = GetMonitorDimensions(frm.hWnd)
MsgBox "ResizeForm CLIENT before: Right:" & rNewWindow.Right & ", Bottom:" & rNewWindow.Bottom _
    & vbCrLf & vbCrLf & "TPPx:" & Screen.TwipsPerPixelX & ", GetTPP:" & GetTwipsPerPixel & ", GetTPP(" & Round(nScaleFactor, 2) & "):" & GetTwipsPerPixel(nScaleFactor) _
    & vbCrLf & "MonDPI:" & GetDpiForWindowProxy(frm.hWnd) & ", MonSize:" & (rMon.Right - rMon.Left) & "x" & (rMon.Bottom - rMon.Top)

If bByTwips Then
    borderWidth = (frm.Width - frm.ScaleWidth) / 2
    borderHeight = borderWidth
    captionHeight = (frm.Height - frm.ScaleHeight) - (borderHeight * 2)
    
    rNewWindow.Top = 0 - (captionHeight + borderHeight)
    rNewWindow.Left = 0 - borderWidth
    rNewWindow.Bottom = rNewWindow.Bottom + borderHeight
    rNewWindow.Right = rNewWindow.Right + borderWidth

ElseIf nScaleFactor > 0 And nScaleFactor <> 1 Or 1 = 1 Then
    
    nDPI = (96 * nScaleFactor)
    
    'AdjustWindowRectExForDpi method...
    AdjDPIrNewWindow = rNewWindow
    AdjustWindowRectExForDpi AdjDPIrNewWindow, GetWindowLong(frm.hWnd, GWL_STYLE), hMenu, GetWindowLong(frm.hWnd, GWL_EXSTYLE), nDPI
    
        AdjDPIborderWidth = Abs(AdjDPIrNewWindow.Left) 'starting from 0, the left value goes negative, accouting for just the left border
        AdjDPIborderHeight = AdjDPIrNewWindow.Bottom - rNewWindow.Bottom 'rNewWindow.Bottom is the client height prior to adjustment, so the adjusted bottom minues that leaves only the bottom border
        AdjDPIcaptionHeight = Abs(AdjDPIrNewWindow.Top) - AdjDPIborderHeight 'likewise, the title bar + top border go negative into the .top
        AdjDPIwidth = AdjDPIrNewWindow.Right - AdjDPIrNewWindow.Left
        AdjDPIheight = AdjDPIrNewWindow.Bottom - AdjDPIrNewWindow.Top
        
    'GetSystemMetricsForDpi method...
    gsmDPIcaptionHeight = GetSystemMetricsForDpi(SM_CYCAPTION, nDPI)
    gsmDPIborderPad = GetSystemMetricsForDpi(SM_CXPADDEDBORDER, nDPI)
    gsmDPIborderWidth = GetSystemMetricsForDpi(SM_CXFRAME, nDPI) + gsmDPIborderPad
    gsmDPIborderHeight = GetSystemMetricsForDpi(SM_CYFRAME, nDPI) + gsmDPIborderPad
    
        'doing it this way to simulate how AdjustWindowRectEx functions
        gsmDPIrNewWindow = rNewWindow
        gsmDPIrNewWindow.Top = 0 - (gsmDPIcaptionHeight + gsmDPIborderHeight)
        gsmDPIrNewWindow.Left = 0 - gsmDPIborderWidth
        gsmDPIrNewWindow.Bottom = gsmDPIrNewWindow.Bottom + gsmDPIborderHeight
        gsmDPIrNewWindow.Right = gsmDPIrNewWindow.Right + gsmDPIborderWidth
        gsmDPIwidth = (gsmDPIrNewWindow.Right - gsmDPIrNewWindow.Left)
        gsmDPIheight = (gsmDPIrNewWindow.Bottom - gsmDPIrNewWindow.Top)
    
    
    'GetSystemMetrics * nScaleFactor method...
    captionHeight = GetSystemMetrics(SM_CYCAPTION) * nScaleFactor
    borderPad = GetSystemMetrics(SM_CXPADDEDBORDER)
    borderWidth = (borderPad + GetSystemMetrics(SM_CXFRAME)) * nScaleFactor
    borderHeight = (borderPad + GetSystemMetrics(SM_CYFRAME)) * nScaleFactor
    
        'doing it this way to simulate how AdjustWindowRectEx functions
        rNewWindow.Top = 0 - (captionHeight + borderHeight + menuHeight)
        rNewWindow.Left = 0 - borderWidth
        rNewWindow.Bottom = rNewWindow.Bottom + borderHeight
        rNewWindow.Right = rNewWindow.Right + borderWidth
    
    
    'form exterior / twips method...
    TWIPborderWidth = (frm.Width - frm.ScaleWidth) / 2
    TWIPcaptionHeight = frm.Height - frm.ScaleHeight - (TWIPborderWidth * 2)
    TWIPwidth = nSetClientWidthTwips + (TWIPborderWidth * 2)
    TWIPheight = nSetClientHeightTwips + TWIPcaptionHeight + (TWIPborderWidth * 2)
    
        rNewWindow.Top = 0
        rNewWindow.Left = 0
        rNewWindow.Bottom = ConvertScale(TWIPheight, vbTwips, vbPixels, nScaleFactor)
        rNewWindow.Right = ConvertScale(TWIPwidth, vbTwips, vbPixels, nScaleFactor)
        
    MsgBox "         AdjDPI | GSMdpi | GSM*" & Round(nScaleFactor, 2) & " | TWIP > PIXEL" _
        & vbCrLf & "brdr width  : " & AdjDPIborderWidth & " | " & gsmDPIborderWidth & " | " & borderWidth & " | " & TWIPborderWidth & " > " & ConvertScale(TWIPborderWidth, vbTwips, vbPixels, nScaleFactor) _
        & vbCrLf & "brdr height : " & AdjDPIborderHeight & " | " & gsmDPIborderHeight & " | " & borderHeight & " | " & TWIPborderWidth & " > " & ConvertScale(TWIPborderWidth, vbTwips, vbPixels, nScaleFactor) _
        & vbCrLf & "caption     : " & AdjDPIcaptionHeight & " | " & gsmDPIcaptionHeight & " | " & captionHeight & " | " & TWIPcaptionHeight & " > " & ConvertScale(TWIPcaptionHeight, vbTwips, vbPixels, nScaleFactor) _
        & vbCrLf & vbCrLf & "width WINDOW pixels > twips (AdjDPI |  GSMdpi | GSM*" & Round(nScaleFactor, 2) & " | TWIP)--" _
        & vbCrLf & AdjDPIwidth & " > " & ConvertScale(AdjDPIwidth, vbPixels, vbTwips, nScaleFactor) _
            & " | " & gsmDPIwidth & " > " & ConvertScale(gsmDPIwidth, vbPixels, vbTwips, nScaleFactor) _
            & " | " & (rNewWindow.Right - rNewWindow.Left) & " > " & ConvertScale((rNewWindow.Right - rNewWindow.Left), vbPixels, vbTwips, nScaleFactor) _
            & " | " & ConvertScale(TWIPwidth, vbTwips, vbPixels, nScaleFactor) & " < " & TWIPwidth _
        & vbCrLf & vbCrLf & "height WINDOW pixels > twips (AdjDPI |  GSMdpi | GSM*" & Round(nScaleFactor, 2) & " | TWIP)--" _
        & vbCrLf & AdjDPIheight & " > " & ConvertScale(AdjDPIheight, vbPixels, vbTwips, nScaleFactor) _
            & " | " & gsmDPIheight & " > " & ConvertScale(gsmDPIheight, vbPixels, vbTwips, nScaleFactor) _
            & " | " & (rNewWindow.Bottom - rNewWindow.Top) & " > " & ConvertScale((rNewWindow.Bottom - rNewWindow.Top), vbPixels, vbTwips, nScaleFactor) _
            & " | " & ConvertScale(TWIPheight, vbTwips, vbPixels, nScaleFactor) & " < " & TWIPheight _
        & vbCrLf & vbCrLf & "TPPx:" & Screen.TwipsPerPixelX & ", GetTPP:" & GetTwipsPerPixel & ", GetTPP(" & Round(nScaleFactor, 2) & "):" & GetTwipsPerPixel(nScaleFactor) _
    & vbCrLf & "MonDPI:" & GetDpiForWindowProxy(frm.hWnd) & ", MonSize:" & (rMon.Right - rMon.Left) & "x" & (rMon.Bottom - rMon.Top)
    
Else
    AdjustWindowRectEx rNewWindow, GetWindowLong(frm.hWnd, GWL_STYLE), hMenu, GetWindowLong(frm.hWnd, GWL_EXSTYLE)
End If

MsgBox "ResizeForm WINDOW adjusted: t" & rNewWindow.Top & " L" & rNewWindow.Left & " r" & rNewWindow.Right & " b" & rNewWindow.Bottom _
    & vbCrLf & vbCrLf & "TPPx:" & Screen.TwipsPerPixelX & ", GetTPP:" & GetTwipsPerPixel & ", GetTPP(" & Round(nScaleFactor, 2) & "):" & GetTwipsPerPixel(nScaleFactor) _
    & vbCrLf & "MonDPI:" & GetDpiForWindowProxy(frm.hWnd) & ", MonSize:" & (rMon.Right - rMon.Left) & "x" & (rMon.Bottom - rMon.Top)

If bByTwips Then
    
    'nScaleFactor 1 because movewindow auto scales
    nPxlWidth = ConvertScale(rNewWindow.Right - rNewWindow.Left, vbTwips, vbPixels, 1)
    nPxlHeight = ConvertScale(rNewWindow.Bottom - rNewWindow.Top, vbTwips, vbPixels, 1)
    MsgBox "set PXL: w" & nPxlWidth & "*1.25 h" & nPxlHeight & "*1.25" _
        & vbCrLf & "set TWP: w" & (rNewWindow.Right - rNewWindow.Left) _
        & " h" & (rNewWindow.Bottom - rNewWindow.Top) _
        & vbCrLf & "MonDPI:" & GetDpiForWindowProxy(frm.hWnd) & ", MonSize:" & (rMon.Right - rMon.Left) & "x" & (rMon.Bottom - rMon.Top)
    
    Call GetWindowRect(frm.hWnd, rCurWindow)
    Call MoveWindow(frm.hWnd, rCurWindow.Left, rCurWindow.Top, nPxlWidth, nPxlHeight, True)
    
    'MsgBox "set TP: w" & (rNewWindow.Right - rNewWindow.Left) & " h" & (rNewWindow.Bottom - rNewWindow.Top)
    'frm.Width = (rNewWindow.Right - rNewWindow.Left)
    'frm.Height = (rNewWindow.Bottom - rNewWindow.Top)
Else
    'Call GetWindowRect(frm.hWnd, rCurWindow)
    
    MsgBox "set WINDOW PXL: w" & (rNewWindow.Right - rNewWindow.Left) & " h" & (rNewWindow.Bottom - rNewWindow.Top) _
        & vbCrLf & "set WINDOW TWP: w" & ConvertScale(rNewWindow.Right - rNewWindow.Left, vbPixels, vbTwips, nScaleFactor) _
        & " h" & ConvertScale(rNewWindow.Bottom - rNewWindow.Top, vbPixels, vbTwips, nScaleFactor) _
    & vbCrLf & vbCrLf & "TPPx:" & Screen.TwipsPerPixelX & ", GetTPP:" & GetTwipsPerPixel & ", GetTPP(" & Round(nScaleFactor, 2) & "):" & GetTwipsPerPixel(nScaleFactor) _
    & vbCrLf & "MonDPI:" & GetDpiForWindowProxy(frm.hWnd) & ", MonSize:" & (rMon.Right - rMon.Left) & "x" & (rMon.Bottom - rMon.Top)
        
    'Call MoveWindow(frm.hWnd, rCurWindow.Left, rCurWindow.Top, (rNewWindow.Right - rNewWindow.Left), (rNewWindow.Bottom - rNewWindow.Top), True)
    
    frm.Width = ConvertScale(rNewWindow.Right - rNewWindow.Left, vbPixels, vbTwips, nScaleFactor)
    frm.Height = ConvertScale(rNewWindow.Bottom - rNewWindow.Top, vbPixels, vbTwips, nScaleFactor)
End If

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
    Call SetWindowLong(frm.hWnd, GWL_STYLE, GetWindowLong(frm.hWnd, GWL_STYLE) Xor _
        (WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX))
Else
    Call SetWindowLong(frm.hWnd, GWL_STYLE, GetWindowLong(frm.hWnd, GWL_STYLE) Xor _
        (WS_THICKFRAME Or WS_MINIMIZEBOX))
End If

Call SetWindowPos(frm.hWnd, 0&, 0&, 0&, 0&, 0&, SWP_NOMOVE Or _
    SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)

'frm.BorderStyle = 2 'sizable
'frm.Caption = frm.Caption 'force redraw

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
Dim hDC As Long, lDPI_X As Long, sngTPP_X As Single
Const HimetricPerPixel As Single = 26.45834

If nScaleFactor > 0 Then
    sngTPP_X = 1440 / (96 * nScaleFactor)
Else
    hDC = GetDC(0)
    lDPI_X = GetDeviceCaps(hDC, LOGPIXELSX)
    sngTPP_X = 1440 / lDPI_X
    hDC = ReleaseDC(0, hDC)
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
Dim hDC As Long, lDPI_X As Long, sngTPP_X As Single
Const HimetricPerPixel As Single = 26.45834

If nScaleFactor > 0 Then
    sngTPP_X = 1440 / (96 * nScaleFactor)
Else
    hDC = GetDC(0)
    lDPI_X = GetDeviceCaps(hDC, LOGPIXELSX)
    sngTPP_X = 1440 / lDPI_X
    hDC = ReleaseDC(0, hDC)
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

Public Function GetDpiForWindowProxy(ByVal hWnd As Long) As Long
On Error GoTo error:
Dim hMonitor As Long, dpiX As Long, dpiY As Long, OSVer As cnWin32Ver, hDC As Long

OSVer = Win32Ver
If OSVer < Win8_1 Then GoTo default:

hMonitor = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST)
GetDpiForMonitor hMonitor, MDT_EFFECTIVE_DPI, dpiX, dpiY

If dpiX < 1 Then GoTo default:
GetDpiForWindowProxy = dpiX
GoTo out:

default:
hDC = GetDC(0)
GetDpiForWindowProxy = GetDeviceCaps(hDC, LOGPIXELSX) / 96
hDC = ReleaseDC(0, hDC)

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetDpiForWindowProxy")
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
