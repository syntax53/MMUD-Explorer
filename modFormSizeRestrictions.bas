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
'

'**************************************************************************************
'**************************************************************************************
'**************************************************************************************
'
' Generic subclassing procedures (used in many of the specific subclassing).
'
'**************************************************************************************
'**************************************************************************************
'**************************************************************************************

Public Function RTrimNull(s As String) As String
    Dim i As Integer
    i = InStr(s, vbNullChar)
    If i Then
        RTrimNull = Left$(s, i - 1)
    Else
        RTrimNull = s
    End If
End Function

'Private Sub SubclassExtraData(hWnd As Long, dwRefData As Long, ID As ExtraDataIDs)
'    ' This is used solely to store extra data.
'    '
'    If Not gbAllowSubclassing Then Exit Sub
'    '
'    bSetWhenSubclassing_UsedByIdeStop = True
'    Call SetWindowSubclass(hWnd, AddressOf DummyProcForExtraData, ID, dwRefData)
'End Sub

'Private Function GetSubclassRefData(hWnd As Long, AddressOf_ProcToSubclass As Long) As Long
'    ' This one is used only to fetch the optional dwRefData you may have specified when calling SubclassSomeWindow.
'    ' Typically this would only be used by the subclassed procedure, but it is available to anyone.
'    Call GetWindowSubclass(hWnd, AddressOf_ProcToSubclass, hWnd, GetSubclassRefData)
'End Function

'Private Function GetExtraData(hWnd As Long, ID As ExtraDataIDs) As Long
'    Call GetWindowSubclass(hWnd, AddressOf DummyProcForExtraData, ID, GetExtraData)
'End Function

'Private Function IsSubclassed(hWnd As Long, AddressOf_ProcToSubclass As Long) As Boolean
'    ' This just tells us we're already subclassed.
'    Dim dwRefData As Long
'    IsSubclassed = GetWindowSubclass(hWnd, AddressOf_ProcToSubclass, hWnd, dwRefData) = 1&
'End Function


'Private Sub UnSubclassExtraData(hWnd As Long, ID As ExtraDataIDs)
'    Call RemoveWindowSubclass(hWnd, AddressOf DummyProcForExtraData, ID)
'End Sub

Private Function ProcedureAddress(AddressOf_TheProc As Long)
    ' A private "helper" function for writing the AddressOf_... functions (see above notes).
    ProcedureAddress = AddressOf_TheProc
End Function

'Private Function DummyProcForExtraData(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
'    ' Just used for SubclassExtraData (and GetExtraData and UnSubclassExtraData).
'    If uMsg = WM_DESTROY Then Call RemoveWindowSubclass(hWnd, AddressOf_DummyProc, uIdSubclass)
'    DummyProcForExtraData = NextSubclassProcOnChain(hWnd, uMsg, wParam, lParam)
'End Function

'Private Function AddressOf_DummyProc() As Long
'    AddressOf_DummyProc = ProcedureAddress(AddressOf DummyProcForExtraData)
'End Function

Private Function IdeStopButtonClicked() As Boolean
    ' The following works because all variables are cleared when the STOP button is clicked,
    ' even though other code may still execute such as Windows calling some of the subclassing procedures below.
    IdeStopButtonClicked = Not bSetWhenSubclassing_UsedByIdeStop
End Function

'**************************************************************************************
'**************************************************************************************
'**************************************************************************************
'
' The following are our functions to be subclassed, along with their AddressOf_... function.
' All of the following should be Private to make sure we don't accidentally call it,
' except for the first procedure that's actually used to initiate the subclassing.
'
'**************************************************************************************
'**************************************************************************************
'**************************************************************************************

Public Sub SubclassFormFixedSize(frm As VB.Form)
    '
    ' This fixes the size of a window, even if it won't fit on a monitor.
    '
    ' On this one, we use dwRefData on the first time through so we can do some setup (see FixedSize_RefData).
    ' We can't use GetWindowRect.  It reports an already resized value.
    '
    ' NOTE: If done in the form LOAD event, the form will NOT have been resized from a smaller monitor.
    '       If done in form ACTIVATE or anywhere else, we're too late, and the form will have been resized.
    '
    ' ALSO: If you're in the IDE, and the monitors aren't big enough, do NOT open the form in design mode.
    '       So long as you don't open it, everything is fine, although you can NOT compile in the IDE.
    '       If you're compiling without large enough monitors, you MUST do a command line compile.
    '
    ' This can simultaneously be used by as many forms as will need it.
    '
    ' NOTICE:  Be sure the window is moved (possibly centered) AFTER this is call, or we may not see WM_GETMINMAXINFO until a bit later.
    '
    Dim tMinMax As WindowSizeRestrictions
    Dim PelWidth As Long
    Dim PelHeight As Long
    PelWidth = frm.Width \ Screen.TwipsPerPixelX
    PelHeight = frm.Height \ Screen.TwipsPerPixelY
    tMinMax.MinWidth = PelWidth
    tMinMax.MaxWidth = PelWidth
    tMinMax.MinHeight = PelHeight
    tMinMax.MaxHeight = PelHeight
    
    SubclassSomeWindow frm.hWnd, AddressOf FixedSize_Proc, FixedSize_RefData(frm)
End Sub

Private Function FixedSize_Proc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
    'If hWnd = nDebugHWND Then Debug.Print uMsg
    If uMsg = WM_DESTROY Then
        UnSubclassSomeWindow hWnd, AddressOf_FixedSize_Proc
        FixedSize_Proc = NextSubclassProcOnChain(hWnd, uMsg, wParam, lParam)
        Exit Function
    End If
    If IdeStopButtonClicked Then ' Protect the IDE.  Don't execute any specific stuff if we're stopping.  We may run into COM objects or other variables that no longer exist.
        FixedSize_Proc = NextSubclassProcOnChain(hWnd, uMsg, wParam, lParam)
        Exit Function
    End If
    '
    Dim PelWidth As Long
    Dim PelHeight As Long
    Dim MMI As MINMAXINFO
    Const WM_GETMINMAXINFO As Long = &H24&
    '
    ' And now we force our size to not change.
    If uMsg = WM_GETMINMAXINFO Then
        ' Force the form to stay at initial size.
        PelWidth = dwRefData And &HFFFF&
        PelHeight = (dwRefData And &H7FFF0000) \ &H10000
        '
        CopyMemory MMI, ByVal lParam, LenB(MMI)
        '
        MMI.ptMinTrackSize.x = PelWidth
        MMI.ptMinTrackSize.y = PelHeight
        MMI.ptMaxTrackSize.x = PelWidth
        MMI.ptMaxTrackSize.y = PelHeight
        '
        CopyMemory ByVal lParam, MMI, LenB(MMI)
        Exit Function ' If we process the message, we must return 0 and not let more subclassed procedures execute.
    End If
    '
    ' Give control to other procs, if they exist.
    FixedSize_Proc = NextSubclassProcOnChain(hWnd, uMsg, wParam, lParam)
End Function

Private Function FixedSize_RefData(frm As VB.Form) As Long
    ' We must use this to pass the form's initial width and height.
    ' Note that using GetWindowRect absolutely doesn't work.  It reports an already resized value.
    '
    Dim PelWidth As Long
    Dim PelHeight As Long
    '
    PelWidth = frm.Width \ Screen.TwipsPerPixelX
    PelHeight = frm.Height \ Screen.TwipsPerPixelY
    '
    ' Push PelHeight to high two-bytes, and add PelWidth.
    ' This will easily accomodate any monitor in the foreseeable future.
    FixedSize_RefData = (PelHeight * &H10000 + PelWidth)
End Function

Private Function AddressOf_FixedSize_Proc() As Long
    AddressOf_FixedSize_Proc = ProcedureAddress(AddressOf FixedSize_Proc)
End Function
'**************************************************************************************
'**************************************************************************************
'**************************************************************************************

Public Sub UN_SubclassFormFixedSize(frm As VB.Form)
    UnSubclassSomeWindow frm.hWnd, AddressOf FixedSize_Proc
End Sub

Public Sub UN_SubclassFormMinMaxSize(frm As VB.Form)
    UnSubclassSomeWindow frm.hWnd, AddressOf MinMaxSize_Proc
End Sub

Private Sub SubclassSomeWindow(hWnd As Long, AddressOf_ProcToSubclass As Long, dwRefData As Long)
    ' This just always uses hWnd for uIdSubclass, as we never have a need to subclass the same window to the same proc.
    ' The uniqueness is pfnSubclass and uIdSubclass (2nd and 3rd argument below).
    '
    ' This can be called AFTER the initial subclassing to update dwRefData.
    '
    If Not gbAllowSubclassing Then Exit Sub
    '
    bSetWhenSubclassing_UsedByIdeStop = True
    Call SetWindowSubclass(hWnd, AddressOf_ProcToSubclass, hWnd, dwRefData)
End Sub

Private Sub UnSubclassSomeWindow(hWnd As Long, AddressOf_ProcToSubclass As Long)
    ' Only needed if we specifically want to un-subclass before we're closing the form (or control),
    ' otherwise, it's automatically taken care of when the window closes.
    '
    ' Be careful, some subclassing may require additional cleanup that's not done here.
    Call RemoveWindowSubclass(hWnd, AddressOf_ProcToSubclass, hWnd)
End Sub

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
    '
    ' Some validation before subclassing.
    'If MinWidth > MaxWidth And MaxWidth <> 0 Then MaxWidth = MinWidth
    'If MinHeight > MaxHeight And MaxHeight <> 0 Then MaxHeight = MinHeight
    '
    'SubclassSomeWindow frm.hWnd, AddressOf MinMaxSize_Proc, CLng(MinHeight * &H10000 + MinWidth)
    'SubclassExtraData frm.hWnd, CLng(MaxHeight * &H10000 + MaxWidth), ID_ForMaxSize
    
    With tMinMaxSize
        If .MinWidth > .MaxWidth And .MaxWidth <> 0 Then .MaxWidth = .MinWidth
        If .MinHeight > .MaxHeight And .MaxHeight <> 0 Then .MaxHeight = .MinHeight
    End With
    SubclassSomeWindow frm.hWnd, AddressOf MinMaxSize_Proc, VarPtr(tMinMaxSize)
End Sub

Private Function MinMaxSize_Proc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As MINMAXINFO, ByVal uIdSubclass As Long, dwRefData As WindowSizeRestrictions) As Long
'Dim MinWidth As Long
'Dim MinHeight As Long
'Dim MaxWidth As Long
'Dim MaxHeight As Long
'Dim MMI As MINMAXINFO
Dim bProcessed As Boolean

If IdeStopButtonClicked Then ' Protect the IDE.  Don't execute any specific stuff if we're stopping.  We may run into COM objects or other variables that no longer exist.
    MinMaxSize_Proc = NextSubclassProcOnChain(hWnd, uMsg, wParam, VarPtr(lParam))
    Exit Function
End If

Select Case uMsg
'        Case WM_GETMINMAXINFO
'            MinWidth = dwRefData And &HFFFF&
'            MinHeight = (dwRefData And &H7FFF0000) \ &H10000
'            dwRefData = GetExtraData(hWnd, ID_ForMaxSize)
'            MaxWidth = dwRefData And &HFFFF&
'            MaxHeight = (dwRefData And &H7FFF0000) \ &H10000
'            '
'            CopyMemory MMI, ByVal lParam, LenB(MMI)
'            If MinWidth <> 0 Then MMI.ptMinTrackSize.x = MinWidth
'            If MinHeight <> 0 Then MMI.ptMinTrackSize.y = MinHeight
'            If MaxWidth <> 0 Then MMI.ptMaxTrackSize.x = MaxWidth
'            If MaxHeight <> 0 Then MMI.ptMaxTrackSize.y = MaxHeight
'            CopyMemory ByVal lParam, MMI, LenB(MMI)
'            Exit Function ' If we process the message, we must return 0 and not let more subclass procedures execute.
    Case WM_GETMINMAXINFO
        With dwRefData
            If .MinWidth And .MinWidth <> lParam.ptMinTrackSize.x Then lParam.ptMinTrackSize.x = .MinWidth: bProcessed = True
            If .MinHeight And .MinHeight <> lParam.ptMinTrackSize.y Then lParam.ptMinTrackSize.y = .MinHeight: bProcessed = True
            If .MaxWidth And .MaxWidth <> lParam.ptMaxTrackSize.x Then lParam.ptMaxTrackSize.x = .MaxWidth: bProcessed = True
            If .MaxHeight And .MaxHeight <> lParam.ptMaxTrackSize.y Then lParam.ptMaxTrackSize.y = .MaxHeight: bProcessed = True
        End With
        If bProcessed Then Exit Function
        
    Case WM_DESTROY, WM_UAHDESTROYWINDOW
        UnSubclassSomeWindow hWnd, AddressOf modFormSizeRestrictions.MinMaxSize_Proc 'AddressOf_MinMaxSize_Proc
        
End Select

' Give control to other procs, if they exist.
MinMaxSize_Proc = NextSubclassProcOnChain(hWnd, uMsg, wParam, VarPtr(lParam))
End Function

'Private Function AddressOf_MinMaxSize_Proc() As Long
'    AddressOf_MinMaxSize_Proc = ProcedureAddress(AddressOf MinMaxSize_Proc)
'End Function
