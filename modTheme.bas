Attribute VB_Name = "modTheme"
Option Explicit

'=========================================================================
' Dark mode theming.
' bDarkMode is read from settings.ini ([Settings] DarkMode=1) in Sub Main,
' before any form loads. Switching modes requires an app restart; every
' form calls ApplyDarkTheme(Me) in its Form_Load, which is a no-op when
' bDarkMode is False (the design-time colors ARE the light theme).
'
' Notes on scope/intentional exclusions:
'  - CommandButtons have no ForeColor in VB6 (captions are always drawn in
'    the system button-text color, i.e. black), so their face is remapped
'    to a readable mid-gray (DK_BTN_FACE) instead of a true dark color.
'    All buttons were switched to Style=1 (Graphical) so BackColor takes
'    effect; in an unthemed app this renders identically in light mode.
'  - Frames draw their border with system 3D colors that cannot be
'    recolored, and BorderStyle=0 also hides the caption -- so in dark
'    mode the system border is dropped, a muted group-box border (top line
'    at caption mid-height, with a gap around the caption) is drawn from a
'    subclass, and the caption is recreated as an overlay label. Code that
'    changes a frame caption at runtime must go through SetFrameCaption so
'    the overlay and the border gap stay in sync.
'  - Map room cells (lblRoomCell) and any label with a custom opaque
'    BackColor (map legend swatches, the black character-stat panel) keep
'    their explicit colors -- those colors are data, not chrome.
'  - Runtime color assignments elsewhere in the code route through TColor()
'    (or TBtnColor() for button faces) so semantic colors stay readable on
'    the dark background via lightness inversion.
'  - The 3D sunken client-edge borders (white highlight edge) are stripped
'    from field controls in dark mode and replaced with a thin system
'    border. ComboBoxes paint their sunken border internally instead, so
'    they get a subclassed post-paint that overdraws it in muted colors.
'  - Scrollbars, listview headers, and popup menus keep the system light
'    look: the uxtheme dark-mode calls for them proved ineffective here and
'    destabilized the IDE (they act on the whole process, i.e. VB6.EXE
'    itself when running from the IDE), so they were removed.
'  - Subclassing is gated by gbAllowSubclassing, matching the app's
'    existing conventions.
'=========================================================================

Global bDarkMode As Boolean

'palette (OLE_COLOR values, &H00BBGGRR)
Public Const DK_FORM_BACK As Long = &H202020    'window/dialog background
Public Const DK_FIELD_BACK As Long = &H262525   'text/list/combo field background
Public Const DK_TEXT As Long = &HE0E0E0         'standard text
Public Const DK_TEXT_DIM As Long = &H909090     'disabled/gray text
Public Const DK_BTN_FACE As Long = &H989898     'button face (captions stay black)
Public Const DK_LINE As Long = &H505050         'separator lines / muted borders

Private Declare Function DwmSetWindowAttribute Lib "dwmapi.dll" (ByVal hWnd As Long, ByVal dwAttribute As Long, ByRef pvAttribute As Any, ByVal cbAttribute As Long) As Long

Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Type RECT_DKT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT_DKT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT_DKT, ByVal hBrush As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT_DKT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function SetWindowSubclass Lib "comctl32.dll" Alias "#410" ( _
    ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, _
    Optional ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32.dll" Alias "#412" ( _
    ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function NextSubclassProcOnChain Lib "comctl32.dll" Alias "#413" ( _
    ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const DWMWA_USE_IMMERSIVE_DARK_MODE As Long = 20
Private Const DWMWA_USE_IMMERSIVE_DARK_MODE_PRE_20H1 As Long = 19

Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const WS_BORDER As Long = &H800000
Private Const WS_EX_CLIENTEDGE As Long = &H200
'SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED
Private Const SWP_BORDERFLAGS As Long = &H37

Private Const WM_PAINT_DKT As Long = &HF
Private Const WM_DESTROY_DKT As Long = &H2
Private Const WM_NCDESTROY_DKT As Long = &H82
Private Const DKT_COMBO_SUBCLASS_ID As Long = &H444B31   'arbitrary unique ids
Private Const DKT_FRAME_SUBCLASS_ID As Long = &H444B32

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

Private Function DKT_CanSubclass() As Boolean
DKT_CanSubclass = gbAllowSubclassing
End Function

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
' Packs the geometry the frame-border subclass needs into one Long:
' caption half-height (border line y), and the left/right pixel edges of
' the caption gap. Computed from the overlay caption label.
'=========================================================================
Private Function PackFrameCaptionInfo(oLbl As Object) As Long
Dim nHalfH As Long, nGapL As Long, nGapR As Long
On Error Resume Next

nHalfH = CLng(oLbl.Height / Screen.TwipsPerPixelY) \ 2
nGapL = CLng(oLbl.Left / Screen.TwipsPerPixelX) - 2
nGapR = CLng((oLbl.Left + oLbl.Width) / Screen.TwipsPerPixelX) + 2

If nHalfH < 0 Then nHalfH = 0
If nHalfH > 255 Then nHalfH = 255
If nGapL < 0 Then nGapL = 0
If nGapL > 255 Then nGapL = 255
If nGapR < 0 Then nGapR = 0
If nGapR > 65535 Then nGapR = 65535

PackFrameCaptionInfo = nHalfH + (nGapL * &H100&) + (nGapR * &H10000)

End Function

'=========================================================================
' Re-colors a form and its controls for dark mode. Call from Form_Load.
' Safe to call unconditionally -- exits immediately in light mode.
'=========================================================================
Public Sub ApplyDarkTheme(frm As Object)
Dim ctl As Object
Dim oCap As Object
Dim sName As String
Dim sCap As String
Dim nCapN As Long
Dim nFramePack As Long
Dim bFrameBorder As Boolean

If Not bDarkMode Then Exit Sub
On Error Resume Next

frm.BackColor = DK_FORM_BACK
Call ApplyDarkTitleBar(frm.hWnd)

For Each ctl In frm.Controls

    sName = ""
    sName = ctl.name

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

            Case "TextBox", "ListBox"
                'a custom opaque back (e.g. the black txtMapMove/room boxes)
                'is intentional styling -- leave those untouched
                If (ctl.BackColor And &H80000000) Or ctl.BackColor = vbWhite Then
                    If ctl.BackColor = vbWhite Or ctl.BackColor = &H80000005 Then
                        ctl.BackColor = DK_FIELD_BACK
                    Else
                        ctl.BackColor = TColor(ctl.BackColor)
                    End If
                    ctl.ForeColor = TColor(ctl.ForeColor)
                    Call MuteSunkenBorder(ctl.hWnd, True)
                End If

            Case "ComboBox"
                If (ctl.BackColor And &H80000000) Or ctl.BackColor = vbWhite Then
                    If ctl.BackColor = vbWhite Or ctl.BackColor = &H80000005 Then
                        ctl.BackColor = DK_FIELD_BACK
                    Else
                        ctl.BackColor = TColor(ctl.BackColor)
                    End If
                    ctl.ForeColor = TColor(ctl.ForeColor)
                    'the combo paints its sunken border internally (not via
                    'the client-edge style), so overdraw it after every paint
                    If DKT_CanSubclass() Then Call SetWindowSubclass(ctl.hWnd, AddressOf DarkComboBorderProc, DKT_COMBO_SUBCLASS_ID, 0)
                End If

            Case "CheckBox", "OptionButton"
                'a custom opaque back (e.g. the map options panel, which is
                'already styled dark) is intentional -- leave those alone
                If (ctl.BackColor And &H80000000) Then
                    ctl.BackColor = TColor(ctl.BackColor)
                    ctl.ForeColor = TColor(ctl.ForeColor)
                End If

            Case "Frame"
                'frames with a custom opaque back (e.g. the map options
                'panel) are intentionally styled -- keep their original
                'border, caption, and colors
                If (ctl.BackColor And &H80000000) = 0 Then GoTo nextframe:
                ctl.BackColor = TColor(ctl.BackColor)
                ctl.ForeColor = TColor(ctl.ForeColor)
                'the frame border is drawn with system 3D colors that cannot
                'be recolored, and BorderStyle=0 also hides the caption -- so
                'drop the system border, redraw a muted group-box border from
                'a subclass, and recreate the caption as an overlay label.
                'ctl.Caption stays intact for any code that reads it; runtime
                'caption changes must go through SetFrameCaption.
                bFrameBorder = False
                If Not ctl.BorderStyle = 0 Then bFrameBorder = True
                nFramePack = 0
                sCap = ""
                sCap = ctl.Caption
                If bFrameBorder And Len(Trim$(sCap)) > 0 Then
                    nCapN = nCapN + 1
                    Set oCap = frm.Controls.Add("VB.Label", "lblDkFrameCap" & CStr(nCapN))
                    Set oCap.Container = ctl
                    Set oCap.Font = ctl.Font
                    oCap.BackStyle = 0
                    oCap.ForeColor = ctl.ForeColor
                    oCap.Caption = sCap
                    oCap.AutoSize = True
                    oCap.Move 90, 0
                    oCap.ZOrder 0
                    oCap.Visible = True
                    nFramePack = PackFrameCaptionInfo(oCap)
                    Set oCap = Nothing
                End If
                ctl.BorderStyle = 0
                If bFrameBorder And DKT_CanSubclass() Then Call SetWindowSubclass(ctl.hWnd, AddressOf DarkFrameBorderProc, DKT_FRAME_SUBCLASS_ID, nFramePack)
nextframe:

            Case "CommandButton"
                'graphical-style buttons honor BackColor; captions are always
                'black, so remap only the default face to a readable mid-gray
                'and leave custom highlight colors alone
                If ctl.BackColor = &H8000000F Or ctl.BackColor = &H80000005 Then ctl.BackColor = DK_BTN_FACE

            Case "ListView", "TreeView"
                ctl.BackColor = DK_FIELD_BACK
                ctl.ForeColor = DK_TEXT
                'the gridline color is baked into the control (drawn light);
                'no way to recolor it on classic comctl, so drop the grid
                ctl.GridLines = False
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
' Sets a frame caption AND keeps the dark-mode overlay caption label and
' the border-gap geometry (see ApplyDarkTheme) in sync. All runtime frame
' caption changes must use this.
'=========================================================================
Public Sub SetFrameCaption(fra As Object, ByVal sCaption As String)
Dim ctl As Object
On Error Resume Next

fra.Caption = sCaption
If Not bDarkMode Then Exit Sub

For Each ctl In fra.Parent.Controls
    If TypeName(ctl) = "Label" Then
        If Left$(ctl.name, 13) = "lblDkFrameCap" Then
            If ctl.Container Is fra Then
                ctl.Caption = sCaption
                'the caption width changed: update the border-gap geometry
                '(SetWindowSubclass with the same proc/id updates dwRefData)
                If DKT_CanSubclass() Then Call SetWindowSubclass(fra.hWnd, AddressOf DarkFrameBorderProc, DKT_FRAME_SUBCLASS_ID, PackFrameCaptionInfo(ctl))
            End If
        End If
    End If
Next ctl

End Sub

'=========================================================================
' Sets a frame ForeColor AND keeps the dark-mode overlay caption label in
' sync (the real frame caption is no longer drawn in dark mode). Runtime
' frame caption-color changes must use this.
'=========================================================================
Public Sub SetFrameForeColor(fra As Object, ByVal nColor As Long)
Dim ctl As Object
On Error Resume Next

fra.ForeColor = nColor
If Not bDarkMode Then Exit Sub

For Each ctl In fra.Parent.Controls
    If TypeName(ctl) = "Label" Then
        If Left$(ctl.name, 13) = "lblDkFrameCap" Then
            If ctl.Container Is fra Then ctl.ForeColor = nColor
        End If
    End If
Next ctl

End Sub

'=========================================================================
' Subclass proc for ComboBoxes in dark mode: after the control paints its
' classic sunken border (bright system 3D colors), overdraw the two border
' rings in muted dark colors. The drop-down arrow button keeps the classic
' face, matching the app's gray buttons.
'=========================================================================
Public Function DarkComboBorderProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Dim tRC As RECT_DKT
Dim nDC As Long
Dim hBr As Long

DarkComboBorderProc = NextSubclassProcOnChain(hWnd, uMsg, wParam, lParam)

Select Case uMsg

    Case WM_PAINT_DKT
        If bDarkMode Then
            Call GetClientRect(hWnd, tRC)
            nDC = GetDC(hWnd)
            If Not nDC = 0 Then
                hBr = CreateSolidBrush(DK_LINE)
                Call FrameRect(nDC, tRC, hBr)
                Call DeleteObject(hBr)
                tRC.Left = tRC.Left + 1
                tRC.Top = tRC.Top + 1
                tRC.Right = tRC.Right - 1
                tRC.Bottom = tRC.Bottom - 1
                hBr = CreateSolidBrush(DK_FIELD_BACK)
                Call FrameRect(nDC, tRC, hBr)
                Call DeleteObject(hBr)
                Call ReleaseDC(hWnd, nDC)
            End If
        End If

    Case WM_DESTROY_DKT, WM_NCDESTROY_DKT
        Call UnsubclassDarkComboBorder(hWnd)

End Select

End Function

'AddressOf cannot reference a function from inside its own body (the name
'resolves to the return-value variable there), hence these helpers
Private Sub UnsubclassDarkComboBorder(ByVal hWnd As Long)
On Error Resume Next
Call RemoveWindowSubclass(hWnd, AddressOf DarkComboBorderProc, DKT_COMBO_SUBCLASS_ID)
End Sub

Private Sub UnsubclassDarkFrameBorder(ByVal hWnd As Long)
On Error Resume Next
Call RemoveWindowSubclass(hWnd, AddressOf DarkFrameBorderProc, DKT_FRAME_SUBCLASS_ID)
End Sub

'=========================================================================
' Subclass proc for Frames in dark mode: the system border was removed
' (BorderStyle=0), so draw a muted group-box border after every paint --
' the top line runs at caption mid-height with a gap around the caption,
' matching where the original etched border was drawn. dwRefData carries
' the packed caption geometry (see PackFrameCaptionInfo); 0 = no caption,
' draw a plain rectangle at the edges.
'=========================================================================
Public Function DarkFrameBorderProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Dim tRC As RECT_DKT
Dim tSeg As RECT_DKT
Dim nDC As Long
Dim hBr As Long
Dim nTop As Long, nGapL As Long, nGapR As Long

DarkFrameBorderProc = NextSubclassProcOnChain(hWnd, uMsg, wParam, lParam)

Select Case uMsg

    Case WM_PAINT_DKT
        If bDarkMode Then
            Call GetClientRect(hWnd, tRC)
            nDC = GetDC(hWnd)
            If Not nDC = 0 Then
                hBr = CreateSolidBrush(DK_LINE)
                If dwRefData = 0 Then
                    Call FrameRect(nDC, tRC, hBr)
                Else
                    nTop = dwRefData And &HFF&
                    nGapL = (dwRefData \ &H100&) And &HFF&
                    nGapR = (dwRefData \ &H10000) And &HFFFF&
                    If nGapR > tRC.Right Then nGapR = tRC.Right
                    'left edge
                    tSeg.Left = 0: tSeg.Right = 1
                    tSeg.Top = nTop: tSeg.Bottom = tRC.Bottom
                    Call FillRect(nDC, tSeg, hBr)
                    'right edge
                    tSeg.Left = tRC.Right - 1: tSeg.Right = tRC.Right
                    Call FillRect(nDC, tSeg, hBr)
                    'bottom edge
                    tSeg.Left = 0: tSeg.Right = tRC.Right
                    tSeg.Top = tRC.Bottom - 1: tSeg.Bottom = tRC.Bottom
                    Call FillRect(nDC, tSeg, hBr)
                    'top edge, left of the caption
                    tSeg.Top = nTop: tSeg.Bottom = nTop + 1
                    tSeg.Left = 0: tSeg.Right = nGapL
                    Call FillRect(nDC, tSeg, hBr)
                    'top edge, right of the caption
                    tSeg.Left = nGapR: tSeg.Right = tRC.Right
                    Call FillRect(nDC, tSeg, hBr)
                End If
                Call DeleteObject(hBr)
                Call ReleaseDC(hWnd, nDC)
            End If
        End If

    Case WM_DESTROY_DKT, WM_NCDESTROY_DKT
        Call UnsubclassDarkFrameBorder(hWnd)

End Select

End Function

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
        Case &H80000008, &H80000012, &H80000007     'window text, button text, menu text
            TColor = DK_TEXT
        Case &H8000000F, &H80000016, &H80000004     'button face, menu background
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
Dim h As Double, s As Double, L As Double
Dim d As Double

r = (nColor And &HFF&) / 255#
g = ((nColor \ &H100&) And &HFF&) / 255#
b = ((nColor \ &H10000) And &HFF&) / 255#

mx = r: If g > mx Then mx = g
If b > mx Then mx = b
mn = r: If g < mn Then mn = g
If b < mn Then mn = b

L = (mx + mn) / 2#

If mx = mn Then
    h = 0#
    s = 0#
Else
    d = mx - mn
    If L > 0.5 Then
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

L = 1# - L
If L > 0.85 Then L = 0.85   'keep light results off pure white
If L < 0.12 Then L = 0.12   'keep dark results off pure black

If s = 0# Then
    r = L: g = L: b = L
Else
    Dim q As Double, p As Double
    If L < 0.5 Then
        q = L * (1# + s)
    Else
        q = L + s - (L * s)
    End If
    p = (2# * L) - q
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
