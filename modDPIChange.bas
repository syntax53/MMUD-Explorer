Attribute VB_Name = "modDPIChange"
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Private Type POINTAPI
'    x As Long
'    y As Long
'End Type
'
'Private Type MINMAXINFO
'    ptReserved As POINTAPI
'    ptMaxSize As POINTAPI
'    ptMaxPosition As POINTAPI
'    ptMinTrackSize As POINTAPI
'    ptMaxTrackSize As POINTAPI
'End Type
'
'Private lpPrevWndProc As Long

Private Const WM_NCDESTROY As Long = &H82, WM_DPICHANGED As Long = &H2E0, LOGPIXELSX As Long = &H58
Private Const SWP_NOACTIVATE As Long = &H10, SWP_NOOWNERZORDER As Long = &H200, SWP_NOZORDER As Long = &H4
'Private Const WM_GETMINMAXINFO As Long = &H24
'Private Const GWL_WNDPROC As Long = -4
'Private Const MIN_WIDTH As Long = 13500 ' Minimum width in twips (15 twips = 1 pixel)

Private Declare Function SetWindowSubclass Lib "comctl32" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" Alias "#412" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
'    ByVal lpPrevWndProc As Long, ByVal hWnd As Long, _
'    ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private lNewDPI As Long, sngScaleFactor As Single, xControl As Control

Public Function SubclassForm(objForm As Form, Optional dwRefData As Long) As Boolean
    If dwRefData = 0 Then dwRefData = GetDeviceCaps(objForm.hDC, LOGPIXELSX)
    SubclassForm = SetWindowSubclass(objForm.hWnd, AddressOf WndProc, ObjPtr(objForm), dwRefData)
End Function

Public Function UnSubclassForm(hWnd As Long, uIdSubclass As Long) As Boolean
    UnSubclassForm = RemoveWindowSubclass(hWnd, AddressOf WndProc, uIdSubclass)
End Function

Private Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As RECT, ByVal objForm As Form, ByVal dwRefData As Long) As Long
    Select Case uMsg
        Case WM_NCDESTROY ' Remove subclassing as the window is about to be destroyed
            UnSubclassForm hWnd, ObjPtr(objForm)
        Case WM_DPICHANGED ' This message signals a change in the DPI of the current monitor or the window was dragged to a monitor with a different DPI
            lNewDPI = wParam And &HFFFF&: sngScaleFactor = lNewDPI / dwRefData: SubclassForm objForm, lNewDPI ' Calculate the new DPI and ScaleFactor and update dwRefData
            With lParam
                SetWindowPos hWnd, 0, .Left, .Top, .Right - .Left, .Bottom - .Top, SWP_NOACTIVATE Or SWP_NOOWNERZORDER Or SWP_NOZORDER ' Resize the form to reflect the new DPI changes
            End With
            objForm.Font.Size = objForm.Font.Size * sngScaleFactor
            For Each xControl In objForm.Controls ' After the form is resized do the same for all its controls
                If Not (TypeOf xControl Is Timer Or TypeOf xControl Is Menu) Then ' Do not process controls without dimensions
                    With xControl
                        .Left = .Left * sngScaleFactor: .Top = .Top * sngScaleFactor: .Width = .Width * sngScaleFactor ' Common properties for most controls
                        Select Case True
                            Case TypeOf xControl Is Label, TypeOf xControl Is TextBox, TypeOf xControl Is CommandButton, _
                                 TypeOf xControl Is Frame, TypeOf xControl Is PictureBox, _
                                 TypeOf xControl Is ListBox, TypeOf xControl Is OptionButton, TypeOf xControl Is CheckBox
                                .Font.Size = .Font.Size * sngScaleFactor: .Height = .Height * sngScaleFactor
                            Case TypeOf xControl Is ComboBox ' Height is ReadOnly for a ComboBox
                                .Font.Size = .Font.Size * sngScaleFactor
                            Case TypeOf xControl Is HScrollBar, TypeOf xControl Is Image ' These controls don't have a Font property
                                .Height = .Height * sngScaleFactor
                        End Select
                    End With
                End If
            Next xControl
    End Select
    WndProc = DefSubclassProc(hWnd, uMsg, wParam, VarPtr(lParam))
End Function

