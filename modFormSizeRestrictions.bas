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
'
'
Public gbAllowSubclassing As Boolean    ' Be sure to turn this on if you're going to use subclassing.
'
Private Const WM_DESTROY As Long = &H2&, WM_UAHDESTROYWINDOW As Long = &H90&, WM_GETMINMAXINFO As Long = &H24&

Private Declare Function SetWindowSubclass Lib "comctl32.dll" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, Optional ByVal dwRefData As Long) As Long
Private Declare Function GetWindowSubclass Lib "comctl32.dll" Alias "#411" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, pdwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32.dll" Alias "#412" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function NextSubclassProcOnChain Lib "comctl32.dll" Alias "#413" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Declare Function DefSubclassProc Lib "comctl32.dll" Alias "#413" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'
Dim bSetWhenSubclassing_UsedByIdeStop As Boolean ' Never goes false once set by first subclassing, unless IDE Stop button is clicked.
'
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef Source As Any, ByVal Bytes As Long)
'
'**************************************************************************************
' The following MODULE level stuff is specific to individual subclassing needs.
'**************************************************************************************
'
Private Enum ExtraDataIDs
    ' These must be unique for each piece of extra data.
    ' They just give us 4 bytes each managed by ComCtl32.
    ID_ForMaxSize = 1
End Enum
#If False Then  ' Intellisense fix.
    Dim ID_ForMaxSize
#End If
'
Public Type POINTAPI
    x As Long
    y As Long
End Type
Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
Public Type WindowSizeRestrictions
    MinWidth As Integer
    MaxWidth As Integer
    MinHeight As Integer
    MaxHeight As Integer
End Type

'**************************************************************************************
'**************************************************************************************
'**************************************************************************************

Private Sub SubclassSomeWindow(hWnd As Long, AddressOf_ProcToSubclass As Long, dwRefData As Long)
    ' This can be called AFTER the initial subclassing to update dwRefData.
    If Not gbAllowSubclassing Then Exit Sub
    bSetWhenSubclassing_UsedByIdeStop = True
    Call SetWindowSubclass(hWnd, AddressOf_ProcToSubclass, hWnd, dwRefData)
End Sub

Private Sub UnSubclassSomeWindow(hWnd As Long, AddressOf_ProcToSubclass As Long)
    Call RemoveWindowSubclass(hWnd, AddressOf_ProcToSubclass, hWnd)
End Sub

Private Function IdeStopButtonClicked() As Boolean
    ' The following works because all variables are cleared when the STOP button is clicked,
    ' even though other code may still execute such as Windows calling some of the subclassing procedures below.
    IdeStopButtonClicked = Not bSetWhenSubclassing_UsedByIdeStop
End Function

'**************************************************************************************
'**************************************************************************************
'**************************************************************************************

Public Sub SubclassFormMinMaxSize(frm As VB.Form, tMinMaxSize As WindowSizeRestrictions)
    ' It's PIXELS.
    '
    ' MUST be done in Form_Load event so Windows doesn't resize form on small monitors.
    ' Also, move (such as center) the form after calling so that WM_GETMINMAXINFO is fired.
    ' Can be called repeatedly to change MinWidth, MinHeight, MaxWidth, and MaxHeight with no harm done.
    ' Although, all must be supplied that you wish to maintain.
    '
    ' Not supplying an argument (i.e., leaving it zero) will cause it to be ignored.
    With tMinMaxSize
        If .MinWidth > .MaxWidth And .MaxWidth <> 0 Then .MaxWidth = .MinWidth
        If .MinHeight > .MaxHeight And .MaxHeight <> 0 Then .MaxHeight = .MinHeight
    End With
    SubclassSomeWindow frm.hWnd, AddressOf MinMaxSize_Proc, VarPtr(tMinMaxSize)
End Sub

Public Sub SubclassFormFixedSize(frm As VB.Form, tMinMaxSize As WindowSizeRestrictions)
    ' This fixes the size of a window at its size when this is called
    ' NOTICE:  Be sure the window is moved (possibly centered) AFTER this is call, or we may not see WM_GETMINMAXINFO until a bit later.
    Dim PelWidth As Long
    Dim PelHeight As Long
    PelWidth = ConvertScale(frm.Width, vbTwips, vbPixels)
    PelHeight = ConvertScale(frm.Height, vbTwips, vbPixels)
    tMinMaxSize.MinWidth = PelWidth
    tMinMaxSize.MaxWidth = PelWidth
    tMinMaxSize.MinHeight = PelHeight
    tMinMaxSize.MaxHeight = PelHeight
    
    SubclassSomeWindow frm.hWnd, AddressOf MinMaxSize_Proc, VarPtr(tMinMaxSize)
End Sub

Public Sub UN_SubclassFormSizeRestriction(frm As VB.Form)
    UnSubclassSomeWindow frm.hWnd, AddressOf MinMaxSize_Proc
End Sub

Private Function MinMaxSize_Proc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As MINMAXINFO, ByVal uIdSubclass As Long, dwRefData As WindowSizeRestrictions) As Long
Dim bProcessed As Boolean

If IdeStopButtonClicked Then
    MinMaxSize_Proc = NextSubclassProcOnChain(hWnd, uMsg, wParam, VarPtr(lParam))
    Exit Function
End If

Select Case uMsg
    Case WM_GETMINMAXINFO
        With dwRefData
            If .MinWidth And .MinWidth <> lParam.ptMinTrackSize.x Then lParam.ptMinTrackSize.x = .MinWidth: bProcessed = True
            If .MinHeight And .MinHeight <> lParam.ptMinTrackSize.y Then lParam.ptMinTrackSize.y = .MinHeight: bProcessed = True
            If .MaxWidth And .MaxWidth <> lParam.ptMaxTrackSize.x Then lParam.ptMaxTrackSize.x = .MaxWidth: bProcessed = True
            If .MaxHeight And .MaxHeight <> lParam.ptMaxTrackSize.y Then lParam.ptMaxTrackSize.y = .MaxHeight: bProcessed = True
        End With
        If bProcessed Then Exit Function
        
    Case WM_DESTROY, WM_UAHDESTROYWINDOW
        UnSubclassSomeWindow hWnd, AddressOf modFormSizeRestrictions.MinMaxSize_Proc
        
End Select

MinMaxSize_Proc = NextSubclassProcOnChain(hWnd, uMsg, wParam, VarPtr(lParam))
End Function
