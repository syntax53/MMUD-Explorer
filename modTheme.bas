Attribute VB_Name = "modTheme"
Option Explicit

'=========================================================================
' Dark mode theming.
' bDarkMode is read from settings.ini ([Settings] DarkMode=1) early in
' frmMain.Form_Load, before any other form loads. Switching modes requires
' an app restart; every form calls ApplyDarkTheme(Me) in its Form_Load,
' which is a no-op when bDarkMode is False (the design-time colors ARE the
' light theme).
'
' Notes on scope/intentional exclusions:
'  - CommandButtons have no ForeColor in VB6 (captions are always drawn in
'    the system button-text color, i.e. black), so their face is remapped
'    to a readable mid-gray (DK_BTN_FACE) instead of a true dark color.
'    All buttons were switched to Style=1 (Graphical) so BackColor takes
'    effect; in an unthemed app this renders identically in light mode.
'  - Map room cells (lblRoomCell) and any label with a custom opaque
'    BackColor (map legend swatches, the black character-stat panel) keep
'    their explicit colors -- those colors are data, not chrome.
'  - Runtime color assignments elsewhere in the code route through TColor()
'    (or TBtnColor() for button faces) so semantic colors stay readable on
'    the dark background via lightness inversion.
'  - The 3D sunken client-edge borders (white highlight edge) are stripped
'    from field controls in dark mode and replaced with a thin system
'    border, which Windows draws in a muted gray.
'=========================================================================

Global bDarkMode As Boolean

'palette (OLE_COLOR values, &H00BBGGRR)
Public Const DK_FORM_BACK As Long = &H202020    'window/dialog background
Public Const DK_FIELD_BACK As Long = &H262525   'text/list/combo field background
Public Const DK_TEXT As Long = &HE0E0E0         'standard text
Public Const DK_TEXT_DIM As Long = &H909090     'disabled/gray text
Public Const DK_BTN_FACE As Long = &H989898     'button face (captions stay black)
Public Const DK_LINE As Long = &H505050         'separator lines

Private Declare Function DwmSetWindowAttribute Lib "dwmapi.dll" (ByVal hWnd As Long, ByVal dwAttribute As Long, ByRef pvAttribute As Any, ByVal cbAttribute As Long) As Long

'undocumented uxtheme exports (Win10 1809+); only called when build >= 17763
Private Declare Function SetPreferredAppMode Lib "uxtheme.dll" Alias "#135" (ByVal nAppMode As Long) As Long
Private Declare Sub FlushMenuThemes Lib "uxtheme.dll" Alias "#136" ()

Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const DWMWA_USE_IMMERSIVE_DARK_MODE As Long = 20
Private Const DWMWA_USE_IMMERSIVE_DARK_MODE_PRE_20H1 As Long = 19
Private Const APPMODE_FORCEDARK As Long = 2

Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const WS_BORDER As Long = &H800000
Private Const WS_EX_CLIENTEDGE As Long = &H200
'SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED
Private Const SWP_BORDERFLAGS As Long = &H37

Private Type TH_OSVERSIONINFO
    OSVSize As Long
    dwVerMajor As Long
    dwVerMinor As Long
    dwBuildNumber As Long
    PlatformID As Long
    szCSDVersion As String * 128
End Type
Private Declare Function TH_GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As TH_OSVERSIONINFO) As Long

'=========================================================================
' Call once at startup, right after bDarkMode has been read from the INI.
' Forces dark popup/context menus app-wide on supported Windows 10+ builds.
'=========================================================================
Public Sub InitDarkMode()
On Error Resume Next
If Not bDarkMode Then Exit Sub

If GetOSBuildNumber() >= 17763 Then
    Call SetPreferredAppMode(APPMODE_FORCEDARK)
    Call FlushMenuThemes
End If

End Sub

Private Function GetOSBuildNumber() As Long
Dim tOSV As TH_OSVERSIONINFO
On Error Resume Next
tOSV.OSVSize = Len(tOSV)
If TH_GetVersionEx(tOSV) = 1 Then GetOSBuildNumber = tOSV.dwBuildNumber
End Function

'=========================================================================
' Dark title bar for a single window (no-op when not in dark mode).
'=========================================================================
Public Sub ApplyDarkTitleBar(ByVal hWnd As Long)
Dim nVal As Long
On Error Resume Next
If Not bDarkMode Then Exit Sub
If Not bUseDwmAPI Then Exit Sub
If nOSversion < Win10 Then Exit Sub

nVal = 1
If DwmSetWindowAttribute(hWnd, DWMWA_USE_IMMERSIVE_DARK_MODE, nVal, 4) <> 0 Then
    Call DwmSetWindowAttribute(hWnd, DWMWA_USE_IMMERSIVE_DARK_MODE_PRE_20H1, nVal, 4)
End If

End Sub

'=========================================================================
' Strips the 3D sunken client edge (its highlight edge reads as a bright
' white line on a dark background) and optionally replaces it with a thin
' system border, which Windows draws in a muted gray.
'=========================================================================
Private Sub MuteSunkenBorder(ByVal hWnd As Long, ByVal bAddThinBorder As Boolean)
Dim nStyle As Long
On Error Resume Next
If hWnd = 0 Then Exit Sub

nStyle = GetWindowLongA(hWnd, GWL_EXSTYLE)
If (nStyle And WS_EX_CLIENTEDGE) Then
    Call SetWindowLongA(hWnd, GWL_EXSTYLE, nStyle And Not WS_EX_CLIENTEDGE)
    If bAddThinBorder Then
        nStyle = GetWindowLongA(hWnd, GWL_STYLE)
        Call SetWindowLongA(hWnd, GWL_STYLE, nStyle Or WS_BORDER)
    End If
    Call SetWindowPos(hWnd, 0, 0, 0, 0, 0, SWP_BORDERFLAGS)
End If

End Sub

'=========================================================================
' Re-colors a form and its controls for dark mode. Call from Form_Load.
' Safe to call unconditionally -- exits immediately in light mode.
'=========================================================================
Public Sub ApplyDarkTheme(frm As Object)
Dim ctl As Object
Dim sName As String

If Not bDarkMode Then Exit Sub
On Error Resume Next

frm.BackColor = DK_FORM_BACK
Call ApplyDarkTitleBar(frm.hWnd)

For Each ctl In frm.Controls

    sName = ""
    sName = ctl.Name

    If InStr(1, ctl.Tag & "", "notheme", vbTextCompare) = 0 Then

        Select Case TypeName(ctl)

            Case "Label"
                'map room cells keep their colors (code compares BackColor to
                'specific values); custom opaque backs are intentional styling
                If Not sName = "lblRoomCell" Then
                    If ctl.BackStyle = 0 Then
                        ctl.ForeColor = TColor(ctl.ForeColor)
                    ElseIf (ctl.BackColor And &H80000000) Then
                        ctl.BackColor = TColor(ctl.BackColor)
                        ctl.ForeColor = TColor(ctl.ForeColor)
                    End If
                End If

            Case "TextBox", "ComboBox", "ListBox"
                If ctl.BackColor = vbWhite Or ctl.BackColor = &H80000005 Then
                    ctl.BackColor = DK_FIELD_BACK
                Else
                    ctl.BackColor = TColor(ctl.BackColor)
                End If
                ctl.ForeColor = TColor(ctl.ForeColor)
                If Not TypeName(ctl) = "ComboBox" Then Call MuteSunkenBorder(ctl.hWnd, True)

            Case "CheckBox", "OptionButton"
                If (ctl.BackColor And &H80000000) Then ctl.BackColor = TColor(ctl.BackColor)
                ctl.ForeColor = TColor(ctl.ForeColor)

            Case "Frame"
                If (ctl.BackColor And &H80000000) Then ctl.BackColor = TColor(ctl.BackColor)
                ctl.ForeColor = TColor(ctl.ForeColor)
                'the etched 3D frame border is drawn with the system highlight
                'color (white); captionless frames drop it entirely (a frame
                'with BorderStyle=0 does not draw its caption, so captioned
                'frames keep the border)
                If Len(Trim$(ctl.Caption & "")) = 0 Then ctl.BorderStyle = 0

            Case "CommandButton"
                'graphical-style buttons honor BackColor; captions are always
                'black, so remap only the default face to a readable mid-gray
                'and leave custom highlight colors alone
                If ctl.BackColor = &H8000000F Or ctl.BackColor = &H80000005 Then ctl.BackColor = DK_BTN_FACE

            Case "ListView", "TreeView"
                ctl.BackColor = DK_FIELD_BACK
                ctl.ForeColor = DK_TEXT
                Call MuteSunkenBorder(ctl.hWnd, True)

            Case "Line"
                ctl.BorderColor = DK_LINE

            Case "PictureBox"
                If ctl.BackColor = &H8000000F Then ctl.BackColor = DK_FORM_BACK
                Call MuteSunkenBorder(ctl.hWnd, False)

            Case "cntSplitter"
                ctl.BackColor = DK_FORM_BACK

            'Shape, Image, Timer, Menu, etc: intentionally untouched

        End Select

    End If

Next ctl

End Sub

'=========================================================================
' Translates a color for the active theme. In light mode returns the color
' unchanged. In dark mode, maps the common system colors onto the dark
' palette and inverts the lightness of explicit RGB colors (hue/saturation
' preserved), so e.g. dark red on white becomes light red on dark. The
' inverted lightness is clamped away from pure white/black so results stay
' muted against the dark background.
' NOTE: never wrap a value in TColor twice -- the second pass re-inverts.
'=========================================================================
Public Function TColor(ByVal nColor As Long) As Long

If Not bDarkMode Then
    TColor = nColor
    Exit Function
End If

If (nColor And &H80000000) Then
    Select Case nColor
        Case &H80000005                             'window background
            TColor = DK_FIELD_BACK
        Case &H80000008, &H80000012                 'window text, button text
            TColor = DK_TEXT
        Case &H8000000F, &H80000016                 'button face
            TColor = DK_FORM_BACK
        Case &H80000010, &H80000011, &H80000015     'shadow, gray text
            TColor = DK_TEXT_DIM
        Case Else                                   'other system colors as-is
            TColor = nColor
    End Select
Else
    TColor = InvertLightness(nColor)
End If

End Function

'=========================================================================
' Same idea as TColor but for CommandButton faces: only the default system
' face is remapped (captions are always black, so the face must stay a
' readable mid-gray); explicit highlight colors pass through unchanged.
'=========================================================================
Public Function TBtnColor(ByVal nColor As Long) As Long

If bDarkMode And (nColor = &H8000000F Or nColor = &H80000005) Then
    TBtnColor = DK_BTN_FACE
Else
    TBtnColor = nColor
End If

End Function

'=========================================================================
' RGB -> HSL, L = 1 - L (clamped), -> RGB. Preserves hue and saturation.
'=========================================================================
Private Function InvertLightness(ByVal nColor As Long) As Long
Dim r As Double, g As Double, b As Double
Dim mx As Double, mn As Double
Dim h As Double, s As Double, l As Double
Dim d As Double

r = (nColor And &HFF&) / 255#
g = ((nColor \ &H100&) And &HFF&) / 255#
b = ((nColor \ &H10000) And &HFF&) / 255#

mx = r: If g > mx Then mx = g
If b > mx Then mx = b
mn = r: If g < mn Then mn = g
If b < mn Then mn = b

l = (mx + mn) / 2#

If mx = mn Then
    h = 0#
    s = 0#
Else
    d = mx - mn
    If l > 0.5 Then
        s = d / (2# - mx - mn)
    Else
        s = d / (mx + mn)
    End If
    If mx = r Then
        h = (g - b) / d
        If g < b Then h = h + 6#
    ElseIf mx = g Then
        h = (b - r) / d + 2#
    Else
        h = (r - g) / d + 4#
    End If
    h = h / 6#
End If

l = 1# - l
If l > 0.85 Then l = 0.85   'keep light results off pure white
If l < 0.12 Then l = 0.12   'keep dark results off pure black

If s = 0# Then
    r = l: g = l: b = l
Else
    Dim q As Double, p As Double
    If l < 0.5 Then
        q = l * (1# + s)
    Else
        q = l + s - (l * s)
    End If
    p = (2# * l) - q
    r = Hue2RGB(p, q, h + (1# / 3#))
    g = Hue2RGB(p, q, h)
    b = Hue2RGB(p, q, h - (1# / 3#))
End If

InvertLightness = RGB(CLng(r * 255#), CLng(g * 255#), CLng(b * 255#))

End Function

Private Function Hue2RGB(ByVal p As Double, ByVal q As Double, ByVal t As Double) As Double

If t < 0# Then t = t + 1#
If t > 1# Then t = t - 1#

If t < (1# / 6#) Then
    Hue2RGB = p + ((q - p) * 6# * t)
ElseIf t < 0.5 Then
    Hue2RGB = q
ElseIf t < (2# / 3#) Then
    Hue2RGB = p + ((q - p) * ((2# / 3#) - t) * 6#)
Else
    Hue2RGB = p
End If

End Function
