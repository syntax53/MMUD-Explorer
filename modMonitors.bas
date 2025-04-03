Attribute VB_Name = "modMonitors"
Option Explicit
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function EnumDisplayMonitors Lib "user32" (ByVal hdc As Long, lprcClip As Any, ByVal lpfnEnum As Long, dwData As Any) As Long
Private Declare Function MonitorFromRect Lib "user32" (ByRef lprc As RECT, ByVal dwFlags As Long) As Long
Public Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function UnionRect Lib "user32" (lprcDst As RECT, lprcSrc1 As RECT, lprcSrc2 As RECT) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function DwmGetWindowAttribute Lib "dwmapi.dll" (ByVal hwnd As Long, ByVal dwAttribute As Long, ByRef pvAttribute As Any, ByVal cbAttribute As Long) As Long
Public Declare Function MonitorFromWindow Lib "user32" (ByVal hwnd As Long, ByVal dwFlags As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
                                                    ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
                                                    ByVal cy As Long, ByVal wFlags As Long) As Long

Const DWMWA_EXTENDED_FRAME_BOUNDS = 9&

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

Public Const MONITOR_DEFAULTTONEAREST = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOACTIVATE = &H10
Public Const LOGPIXELSX As Long = 88
Public Const LOGPIXELSY As Long = 90

Dim rcMonitors() As RECT 'coordinate array for all monitors
Dim rcVS         As RECT 'coordinates for Virtual Screen

Public Function EnumMonitors(F As Form) As Long
    Dim N As Long
    EnumDisplayMonitors 0, ByVal 0&, AddressOf MonitorEnumProc, N
'    With F
'        .Move .Left, .Top, (rcVS.Right - rcVS.Left) * 2 + .Width - .ScaleWidth, (rcVS.Bottom - rcVS.Top) * 2 + .Height - .ScaleHeight
'    End With
'    F.Scale (rcVS.Left, rcVS.Top)-(rcVS.Right, rcVS.Bottom)
'    F.Caption = N & " Monitor" & IIf(N > 1, "s", vbNullString)
'    F.lblMonitors(0).Appearance = 0 'Flat
'    F.lblMonitors(0).BorderStyle = 1 'FixedSingle
'    For N = 0 To N - 1
'        If N Then
'            Load F.lblMonitors(N)
'            F.lblMonitors(N).Visible = True
'        End If
'        With rcMonitors(N)
'            F.lblMonitors(N).Move .Left, .Top, .Right - .Left, .Bottom - .Top
'            F.lblMonitors(N).Caption = "Monitor " & N + 1 & vbLf & _
'                .Right - .Left & " x " & .Bottom - .Top & vbLf & _
'                "(" & .Left & ", " & .Top & ")-(" & .Right & ", " & .Bottom & ")"
'        End With
'    Next
End Function
Private Function MonitorEnumProc(ByVal hMonitor As Long, ByVal hdcMonitor As Long, lprcMonitor As RECT, dwData As Long) As Long
    ReDim Preserve rcMonitors(dwData)
    rcMonitors(dwData) = lprcMonitor
    UnionRect rcVS, rcVS, lprcMonitor 'merge all monitors together to get the virtual screen coordinates
    dwData = dwData + 1 'increase monitor count
    MonitorEnumProc = 1 'continue
End Function

'Public Sub SavePosition(hWnd As Long)
'    Dim rc As RECT
'    GetWindowRect hWnd, rc 'save position in pixel units
'    SaveSetting "Multi Monitor Demo", "Position", "Left", rc.Left
'    SaveSetting "Multi Monitor Demo", "Position", "Top", rc.Top
'End Sub

'Public Function GetPosition(F As Form) As RECT
'On Error GoTo error:
'Dim rc As RECT, hWnd As Long
'
'hWnd = GetCurrentMonitor(F)
'If hWnd <> 0 Then
'    GetWindowRect hWnd, rc
'    GetPosition.Top = rc.Top
'    GetPosition.Bottom = rc.Bottom
'    GetPosition.Left = rc.Left
'    GetPosition.Right = rc.Right
'End If
'
'out:
'On Error Resume Next
'Exit Function
'error:
'Call HandleError("GetPosition")
'Resume out:
'End Function
'
'Public Sub SetPosition(F As Form, nTop As Long, nLeft As Long)
'On Error GoTo error:
'Dim hWnd As Long, mi As MONITORINFO, rc As RECT
'
'hWnd = GetCurrentMonitor(F)
'If hWnd <> 0 Then
'    OffsetRect rc, nLeft - rc.Left, nTop - rc.Top
'    mi.cbSize = Len(mi)
'    GetMonitorInfo hWnd, mi
'    If rc.Left < mi.rcWork.Left Then OffsetRect rc, mi.rcWork.Left - rc.Left, 0
'    If rc.Right > mi.rcWork.Right Then OffsetRect rc, mi.rcWork.Right - rc.Right, 0
'    If rc.Top < mi.rcWork.Top Then OffsetRect rc, 0, mi.rcWork.Top - rc.Top
'    If rc.Bottom > mi.rcWork.Bottom Then OffsetRect rc, 0, mi.rcWork.Bottom - rc.Bottom
'    MoveWindow hWnd, rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top, 0
'End If
'
'out:
'On Error Resume Next
'Exit Sub
'error:
'Call HandleError("SetPosition")
'Resume out:
'End Sub

Public Sub LoadPosition(hwnd As Long)
    Dim rc As RECT, Left As Long, Top As Long, hMonitor As Long, mi As MONITORINFO
    GetWindowRect hwnd, rc 'obtain the window rectangle
    'move the window rectangle to position saved previously
    Left = GetSetting("Multi Monitor Demo", "Position", "Left", rc.Left)
    Top = GetSetting("Multi Monitor Demo", "Position", "Top", rc.Left)
    OffsetRect rc, Left - rc.Left, Top - rc.Top
    'find the monitor closest to window rectangle
    hMonitor = MonitorFromRect(rc, MONITOR_DEFAULTTONEAREST)
    'get info about monitor coordinates and working area
    mi.cbSize = Len(mi)
    GetMonitorInfo hMonitor, mi
    'adjust the window rectangle so it fits inside the work area of the monitor
    If rc.Left < mi.rcWork.Left Then OffsetRect rc, mi.rcWork.Left - rc.Left, 0
    If rc.Right > mi.rcWork.Right Then OffsetRect rc, mi.rcWork.Right - rc.Right, 0
    If rc.Top < mi.rcWork.Top Then OffsetRect rc, 0, mi.rcWork.Top - rc.Top
    If rc.Bottom > mi.rcWork.Bottom Then OffsetRect rc, 0, mi.rcWork.Bottom - rc.Bottom
    'move the window to new calculated position
    MoveWindow hwnd, rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top, 0
End Sub

Public Function GetCurrentMonitor(F As Form) As Long
Dim rc As RECT
GetWindowRect F.hwnd, rc
GetCurrentMonitor = MonitorFromRect(rc, MONITOR_DEFAULTTONEAREST)
End Function
Public Sub CheckPosition(F As Form)
    Dim rc As RECT, Left As Long, Top As Long, hMonitor As Long, mi As MONITORINFO
    Dim bMove As Boolean
        
    If bUse_dwmapi = False Then Exit Sub
    If frmMain.bDisableWindowSnap Then Exit Sub
    If F.WindowState = vbMinimized Then Exit Sub
    If F.WindowState = vbMaximized Then Exit Sub
    
    GetWindowRect F.hwnd, rc 'obtain the window rectangle
    'move the window rectangle to position saved previously
    Left = F.Left
    Top = F.Top
    
    '###############################
    
'    RECT rect, frame;
'    GetWindowRect(hwnd, &rect);
'    DwmGetWindowAttribute(hwnd, DWMWA_EXTENDED_FRAME_BOUNDS, &frame, sizeof(RECT));
'
'    //rect should be `0, 0, 1280, 1024`
'    //frame should be `7, 0, 1273, 1017`
'
'    RECT border;
'    border.left = frame.left - rect.left;
'    border.top = frame.top - rect.top;
'    border.right = rect.right - frame.right;
'    border.bottom = rect.bottom - frame.bottom;
'
'    //border should be `7, 0, 7, 7`
    Dim rcWindow As RECT, rcFrame As RECT, nRet As Long
    nRet = GetWindowRect(F.hwnd, rcWindow)
    nRet = DwmGetWindowAttribute(F.hwnd, DWMWA_EXTENDED_FRAME_BOUNDS, rcFrame, Len(rcWindow))
    
    Dim rcBorder As RECT
    
    'COMMENTED THESE OUT ON 2021.08.06 AS IT WAS READING WEIRD ON A DELL LAPTOP WITH ZOOM
    'rcBorder.Left = rcFrame.Left - rcWindow.Left
    'rcBorder.Top = rcFrame.Top - rcWindow.Top
    'rcBorder.Right = rcWindow.Right - rcFrame.Right
    'rcBorder.Bottom = rcWindow.Bottom - rcFrame.Bottom
        
    '###############################
    
    
    'OffsetRect rc, Left - rc.Left, Top - rc.Top
    'find the monitor closest to window rectangle
    hMonitor = MonitorFromRect(rc, MONITOR_DEFAULTTONEAREST)
    'get info about monitor coordinates and working area
    mi.cbSize = Len(mi)
    GetMonitorInfo hMonitor, mi
    'adjust the window rectangle so it fits inside the work area of the monitor
    If rc.Left + rcBorder.Left + rcBorder.Right < mi.rcWork.Left Then
        OffsetRect rc, mi.rcWork.Left - rc.Left, 0
        bMove = True
    End If
    If rc.Right - rcBorder.Left - rcBorder.Right > mi.rcWork.Right Then
        OffsetRect rc, mi.rcWork.Right - rc.Right, 0
        bMove = True
    End If
    If rc.Top + rcBorder.Top + rcBorder.Bottom < mi.rcWork.Top Then
        OffsetRect rc, 0, mi.rcWork.Top - rc.Top
        bMove = True
    End If
    If rc.Bottom - rcBorder.Top - rcBorder.Bottom > mi.rcWork.Bottom Then
        OffsetRect rc, 0, mi.rcWork.Bottom - rc.Bottom
        bMove = True
    End If
    'move the window to new calculated position
    If bMove Then
        MoveWindow F.hwnd, rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top, 0
        If Not F.WindowState = vbMinimized Then
            F.Hide
            F.Show
        End If
    End If
End Sub

Public Sub MonitorFormTimer(oForm As Form)
Dim nCurrentMonitor As Long, nSecOfDay As Long
On Error GoTo error:

If Not oForm.WindowState = vbMinimized Then
    If Not Not oForm.Left = oForm.nLastPosLeft Or Not oForm.Top = oForm.nLastPosTop Then
        nSecOfDay = Hour(Now) * 60 * 60
        nSecOfDay = nSecOfDay + (Minute(Now) * 60)
        nSecOfDay = nSecOfDay + Second(Now)
        
        If oForm.nLastPosMoved = 0 Or (oForm.nLastTimerTop <> oForm.Top Or oForm.nLastTimerLeft <> oForm.Left) Then
            oForm.nLastPosMoved = nSecOfDay
            oForm.nLastTimerTop = oForm.Top
            oForm.nLastTimerLeft = oForm.Left
        End If
        
        If oForm.nLastPosMoved > nSecOfDay - 1 Then Exit Sub
        
        oForm.nLastPosMoved = 0
    End If
    
    nCurrentMonitor = GetCurrentMonitor(oForm)
    If Not Not oForm.Left = oForm.nLastPosLeft Or Not oForm.Top = oForm.nLastPosTop Then
        CheckPosition oForm
        oForm.nLastPosLeft = oForm.Left
        oForm.nLastPosTop = oForm.Top
        nCurrentMonitor = GetCurrentMonitor(oForm)
    End If
    
    DoEvents
    
    If Not oForm.nLastPosMonitor = nCurrentMonitor Then
        If oForm.nLastPosMonitor <> 0 Then
            'oForm.Hide
            'oForm.Show
        End If
        oForm.nLastPosMonitor = GetCurrentMonitor(oForm)
    End If
End If

out:
Exit Sub
error:
Resume out:
End Sub
