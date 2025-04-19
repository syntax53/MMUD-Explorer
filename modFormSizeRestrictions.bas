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
Private Declare Function MonitorFromWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal dwFlags As MONITORFROMWINDOW_FLAGS) As Long
Private Declare Function GetDpiForMonitor Lib "shcore.dll" (ByVal hMonitor As Long, ByVal dpiType As MonitorDpiTypeEnum, ByRef dpiX As Long, ByRef dpiY As Long) As Long
Private Declare Function AdjustWindowRectExForDpi Lib "user32.dll" (ByRef lpRect As RECT, ByVal dwStyle As Long, ByVal bMenu As Long, ByVal dwExStyle As Long, ByVal Dpi As Long) As Long

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
Private Const GWL_STYLE As Long = (-16&)
Private Const GWL_EXSTYLE As Long = (-20&)
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const WM_DESTROY As Long = &H2&
Private Const WM_NCDESTROY As Long = &H82
Private Const WM_UAHDESTROYWINDOW As Long = &H90&
Private Const WM_GETMINMAXINFO As Long = &H24&
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

Public Sub SubclassFormMinMaxSize(frm As VB.Form, tMinMaxSize As WindowSizeRestrictions, Optional ByVal bFixToCurrentSize As Boolean)
Dim nPixelWidth As Long, nPixelHeight As Long, hDC As Long

If tMinMaxSize.ScaleFactor = 0 Then
    hDC = GetDC(0)
    tMinMaxSize.ScaleFactor = GetDeviceCaps(frm.hDC, LOGPIXELSX) / 96
    hDC = ReleaseDC(0, hDC)
End If

'MsgBox "subclass scale: " & tMinMaxSize.ScaleFactor

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
        .pxlMinWidth = 0
        .pxlMinHeight = 0
        .pxlMaxWidth = 0
        .pxlMaxHeight = 0
        If .twpMinWidth Then .pxlMinWidth = ConvertScale(.twpMinWidth, vbTwips, vbPixels, tMinMaxSize.ScaleFactor)
        If .twpMinHeight Then .pxlMinHeight = ConvertScale(.twpMinHeight, vbTwips, vbPixels, tMinMaxSize.ScaleFactor)
        If .twpMaxWidth Then .pxlMaxWidth = ConvertScale(.twpMaxWidth, vbTwips, vbPixels, tMinMaxSize.ScaleFactor)
        If .twpMaxHeight Then .pxlMaxHeight = ConvertScale(.twpMaxHeight, vbTwips, vbPixels, tMinMaxSize.ScaleFactor)
    End With
End If

With tMinMaxSize
    If .pxlMinWidth > .pxlMaxWidth And .pxlMaxWidth <> 0 Then .pxlMaxWidth = .pxlMinWidth
    If .pxlMinHeight > .pxlMaxHeight And .pxlMaxHeight <> 0 Then .pxlMaxHeight = .pxlMinHeight
End With

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
Dim R As RECT, rMenu As RECT, minRECT As RECT, maxRECT As RECT

Select Case uMsg
    Case WM_GETMINMAXINFO
    
        hMenu = GetMenu(hWnd)
        If hMenu > 0 Then
            If GetMenuItemRect(hWnd, hMenu, 0, rMenu) Then
                hMenu = 1
                menuHeight = (rMenu.Bottom - rMenu.Top)
            Else
                hMenu = 0
            End If
        End If
        
        If dwRefData.ScaleFactor > 0 And dwRefData.ScaleFactor <> 1 Then ' And 1 = 2
            captionHeight = GetSystemMetrics(SM_CYCAPTION) * dwRefData.ScaleFactor
            'If hMenu Then menuHeight = GetSystemMetrics(SM_CYMENU) * dwRefData.ScaleFactor
            'menuHeight = menuHeight * dwRefData.ScaleFactor
            borderWidth = GetSystemMetrics(SM_CXFRAME) * dwRefData.ScaleFactor
            borderHeight = GetSystemMetrics(SM_CYFRAME) * dwRefData.ScaleFactor
            
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
                AdjustWindowRectEx minRECT, GetWindowLong(hWnd, GWL_STYLE), hMenu, GetWindowLong(hWnd, GWL_EXSTYLE)
            End If
            
            If dwRefData.pxlMinWidth = dwRefData.pxlMaxWidth And dwRefData.pxlMinHeight = dwRefData.pxlMaxHeight Then
                maxRECT = minRECT
            ElseIf dwRefData.pxlMaxWidth Or dwRefData.pxlMaxHeight Then
                If dwRefData.pxlMaxWidth Then maxRECT.Right = dwRefData.pxlMaxWidth
                If dwRefData.pxlMaxHeight Then maxRECT.Bottom = dwRefData.pxlMaxHeight
                AdjustWindowRectEx maxRECT, GetWindowLong(hWnd, GWL_STYLE), hMenu, GetWindowLong(hWnd, GWL_EXSTYLE)
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
            'RtlMoveMemory ByVal lParam, mmi, LenB(mmi)
            Exit Function
        End If
        
    Case WM_DESTROY, WM_UAHDESTROYWINDOW, WM_NCDESTROY
        UnSubclassSomeWindow hWnd, ObjPtr(objForm)
     
    Case WM_DPICHANGED
        bDPIAwareMode = True
        'RtlMoveMemory r, ByVal lParam, LenB(r)
        'With r
        '    SetWindowPos hWnd, 0, .Left, .Top, .Right - .Left, .Bottom - .Top, SWP_NOACTIVATE Or SWP_NOOWNERZORDER Or SWP_NOZORDER ' Resize the form to reflect the new DPI changes
        'End With
        lNewDPI = wParam And &HFFFF&
        MsgBox "new dpi: " & lNewDPI
        dwRefData.ScaleFactor = lNewDPI / 96
        If dwRefData.twpMinWidth Then dwRefData.pxlMinWidth = ConvertScale(dwRefData.twpMinWidth, vbTwips, vbPixels, dwRefData.ScaleFactor)
        If dwRefData.twpMinHeight Then dwRefData.pxlMinHeight = ConvertScale(dwRefData.twpMinHeight, vbTwips, vbPixels, dwRefData.ScaleFactor)
        If dwRefData.twpMaxWidth Then dwRefData.pxlMaxWidth = ConvertScale(dwRefData.twpMaxWidth, vbTwips, vbPixels, dwRefData.ScaleFactor)
        If dwRefData.twpMaxHeight Then dwRefData.pxlMaxHeight = ConvertScale(dwRefData.twpMaxHeight, vbTwips, vbPixels, dwRefData.ScaleFactor)
        
End Select

MinMaxSize_Proc = NextSubclassProcOnChain(hWnd, uMsg, wParam, lParam)

End Function

'**************************************************************************************
'**************************************************************************************
'**************************************************************************************

Public Sub ResizeForm(frm As VB.Form, nSetClientWidthTwips As Long, nSetClientHeightTwips As Long, Optional ByVal nScaleFactor As Single, Optional ByVal bByTwips As Boolean)
On Error GoTo error:
Dim captionHeight As Long, menuHeight As Long, borderWidth As Long, borderHeight As Long
Dim TWIPcaptionHeight As Long, TWIPmenuHeight As Long, TWIPborderWidth As Long, TWIPwidth As Long, TWIPheight As Long
Dim AdjDPIcaptionHeight As Long, AdjDPIborderWidth As Long, AdjDPIborderHeight As Long, AdjDPIwidth As Long, AdjDPIheight As Long
Dim currentStyle As Long, currentExStyle As Long, hMenu As Long
Dim rCurWindow As RECT, rNewWindow As RECT, rMenu As RECT
Dim AdjDPIrNewWindow As RECT

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

MsgBox "before: t" & rNewWindow.Top & " L" & rNewWindow.Left & " r" & rNewWindow.Right & " b" & rNewWindow.Bottom

If bByTwips Then
    borderWidth = (frm.Width - frm.ScaleWidth) / 2
    borderHeight = borderWidth
    captionHeight = (frm.Height - frm.ScaleHeight) - (borderHeight * 2)
    
    rNewWindow.Top = 0 - (captionHeight + borderHeight + menuHeight)
    rNewWindow.Left = 0 - borderWidth
    rNewWindow.Bottom = rNewWindow.Bottom + borderHeight
    rNewWindow.Right = rNewWindow.Right + borderWidth

ElseIf nScaleFactor > 0 And nScaleFactor <> 1 Then
    
    'AdjustWindowRectExForDpi method...
    AdjDPIrNewWindow = rNewWindow
    AdjustWindowRectExForDpi AdjDPIrNewWindow, GetWindowLong(frm.hWnd, GWL_STYLE), hMenu, GetWindowLong(frm.hWnd, GWL_EXSTYLE), (96 * nScaleFactor)
    
    AdjDPIborderWidth = Abs(AdjDPIrNewWindow.Left) 'starting from 0, the left value goes negative, accouting for just the left border
    AdjDPIborderHeight = AdjDPIrNewWindow.Bottom - rNewWindow.Bottom 'rNewWindow.Bottom is the client height prior to adjustment, so the adjusted bottom minues that leaves only the bottom border
    AdjDPIcaptionHeight = Abs(AdjDPIrNewWindow.Top) - AdjDPIborderHeight 'likewise, the title bar + top border go negative into the .top
    AdjDPIwidth = AdjDPIrNewWindow.Right - AdjDPIrNewWindow.Left
    AdjDPIheight = AdjDPIrNewWindow.Bottom - AdjDPIrNewWindow.Top
    
    'form exterior / twips method...
    TWIPborderWidth = (frm.Width - frm.ScaleWidth) / 2
    TWIPcaptionHeight = frm.Height - frm.ScaleHeight - (TWIPborderWidth * 2)
    TWIPwidth = nSetClientWidthTwips + (TWIPborderWidth * 2)
    TWIPheight = nSetClientHeightTwips + TWIPcaptionHeight + (TWIPborderWidth * 2)
    
    'GetSystemMetrics method...
    captionHeight = GetSystemMetrics(SM_CYCAPTION) * nScaleFactor
    borderWidth = GetSystemMetrics(SM_CXFRAME) * nScaleFactor
    borderHeight = GetSystemMetrics(SM_CYFRAME) * nScaleFactor
    
    'doing it this way to simulate how AdjustWindowRectEx functions
    rNewWindow.Top = 0 - (captionHeight + borderHeight + menuHeight)
    rNewWindow.Left = 0 - borderWidth
    rNewWindow.Bottom = rNewWindow.Bottom + borderHeight
    rNewWindow.Right = rNewWindow.Right + borderWidth
    
    MsgBox "" _
                 & "        : AdjDPI | GSM | TWIP > PIXEL" _
        & vbCrLf & "brdr width  : " & AdjDPIborderWidth & " | " & borderWidth & " | " & TWIPborderWidth & " > " & ConvertScale(TWIPborderWidth, vbTwips, vbPixels, nScaleFactor) _
        & vbCrLf & "brdr height : " & AdjDPIborderHeight & " | " & borderHeight & " | " & TWIPborderWidth & " > " & ConvertScale(TWIPborderWidth, vbTwips, vbPixels, nScaleFactor) _
        & vbCrLf & "caption     : " & AdjDPIcaptionHeight & " | " & captionHeight & " | " & TWIPcaptionHeight & " > " & ConvertScale(TWIPcaptionHeight, vbTwips, vbPixels, nScaleFactor) _
        & vbCrLf & vbCrLf & "width pixels > twips (AdjDPI | GSM | TWIP)--" _
        & vbCrLf & AdjDPIwidth & " > " & ConvertScale(AdjDPIwidth, vbPixels, vbTwips, nScaleFactor) _
            & " | " & (rNewWindow.Right - rNewWindow.Left) & " > " & ConvertScale((rNewWindow.Right - rNewWindow.Left), vbPixels, vbTwips, nScaleFactor) _
            & " | " & ConvertScale(TWIPwidth, vbTwips, vbPixels, nScaleFactor) & " < " & TWIPwidth _
        & vbCrLf & vbCrLf & "height pixels > twips (AdjDPI | GSM | TWIP)--" _
        & vbCrLf & AdjDPIheight & " > " & ConvertScale(AdjDPIheight, vbPixels, vbTwips, nScaleFactor) _
            & " | " & (rNewWindow.Bottom - rNewWindow.Top) & " > " & ConvertScale((rNewWindow.Bottom - rNewWindow.Top), vbPixels, vbTwips, nScaleFactor) _
            & " | " & ConvertScale(TWIPheight, vbTwips, vbPixels, nScaleFactor) & " < " & TWIPheight
    
Else
    AdjustWindowRectEx rNewWindow, GetWindowLong(frm.hWnd, GWL_STYLE), hMenu, GetWindowLong(frm.hWnd, GWL_EXSTYLE)
End If
MsgBox "adjusted: t" & rNewWindow.Top & " L" & rNewWindow.Left & " r" & rNewWindow.Right & " b" & rNewWindow.Bottom

If bByTwips Then
    MsgBox "set TP: w" & (rNewWindow.Right - rNewWindow.Left) & " h" & (rNewWindow.Bottom - rNewWindow.Top)
    
    frm.Width = (rNewWindow.Right - rNewWindow.Left)
    frm.Height = (rNewWindow.Bottom - rNewWindow.Top)
Else
    Call GetWindowRect(frm.hWnd, rCurWindow)
    
    MsgBox "set PXL: w" & (rNewWindow.Right - rNewWindow.Left) & " h" & (rNewWindow.Bottom - rNewWindow.Top) _
        & vbCrLf & "set TWP: w" & ConvertScale(rNewWindow.Right - rNewWindow.Left, vbPixels, vbTwips, nScaleFactor) _
        & " h" & ConvertScale(rNewWindow.Bottom - rNewWindow.Top, vbPixels, vbTwips, nScaleFactor)
        
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
Dim hMonitor As Long, dpiX As Long, dpiY As Long, OSVer As cnWin32Ver

OSVer = Win32Ver
If OSVer < Win8_1 Then
    GetDpiForWindowProxy = 96
    Exit Function
End If

hMonitor = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST)

GetDpiForMonitor hMonitor, MDT_EFFECTIVE_DPI, dpiX, dpiY
GetDpiForWindowProxy = dpiX

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetDpiForWindowProxy")
Resume out:
End Function

