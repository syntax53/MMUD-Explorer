Attribute VB_Name = "modSyntaxsFunc"
Option Explicit
Option Base 0

Global Const STATE_SYSTEM_FOCUSABLE = &H100000
Global Const STATE_SYSTEM_INVISIBLE = &H8000
Global Const STATE_SYSTEM_OFFSCREEN = &H10000
Global Const STATE_SYSTEM_UNAVAILABLE = &H1
Global Const STATE_SYSTEM_PRESSED = &H8
Global Const CCHILDREN_TITLEBAR = 5
Global Const LB_GETITEMRECT = &H198
Global Const CB_GETDROPPEDCONTROLRECT = &H15F
'Global Const CB_GETITEMHEIGHT = &H154
Global Const MF_BYPOSITION = &H400&
Global Const MF_DISABLED = &H2&
Global TITLEBAR_OFFSET As Integer

Global Const LongOffset = 4294967296#
Global Const MaxLong = 2147483647
Global Const IntOffset = 65536
Global Const MaxInt = 32767
Public Const CB_FINDSTRING = &H14C

Public Type ResizeCons
    LeftGap As Long
    TopGap As Long
    RightGap As Long
    BottomGap As Long
End Type

Public Enum ListDataType
    ldtstring = 0
    ldtnumber = 1
    ldtDateTime = 2
End Enum

Type TITLEBARINFO
    cbSize As Long
    rcTitleBar As RECT
    rgstate(CCHILDREN_TITLEBAR) As Long
End Type

Public bSuppressErrors As Boolean

'Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
'Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
'Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetTitleBarInfo Lib "user32" (ByVal hwnd As Long, ByRef pti As TITLEBARINFO) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Declare Function SetWindowPos Lib "user32" _
      (ByVal hwnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long


Const GWL_EXSTYLE = -20
Const GWL_HINSTANCE = -6
Const GWL_HWNDPARENT = -8
Const GWL_ID = -12
Const GWL_STYLE = -16
Const GWL_USERDATA = -21
Const GWL_WNDPROC = -4
Const DWL_DLGPROC = 4
Const DWL_MSGRESULT = 0
Const DWL_USER = 8

'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
'    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetParent Lib "user32" _
    (ByVal FormHwnd As Long, Optional ByVal NewHwnd As Long) As Long
    

Public Function GetOwner(ByVal HwndofForm) As Long
    GetOwner = GetWindowLong(HwndofForm, GWL_HWNDPARENT)
End Function

Public Function SetOwner(ByVal HwndtoUse, ByVal HwndofOwner) As Long
    SetOwner = SetWindowLong(HwndtoUse, GWL_HWNDPARENT, HwndofOwner)
'    If HwndofOwner = 0 Then
'        SetOwner = SetParent(HwndtoUse, Null)
'    Else
'        SetOwner = SetParent(HwndtoUse, HwndofOwner)
'    End If
End Function

Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) _
   As Long

   If Topmost = True Then 'Make the window topmost
      SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, _
         0, FLAGS)
   Else
      SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, _
         0, 0, FLAGS)
      SetTopMostWindow = False
   End If
End Function

Public Function in_long_arr(ByVal nSearch As Long, arrLong() As Long) As Boolean
Dim nIndex As Long
On Error GoTo error:

For nIndex = LBound(arrLong) To UBound(arrLong)
    If arrLong(nIndex) = nSearch Then
        in_long_arr = True
        Exit For
    End If
Next

out:
On Error Resume Next
Exit Function
error:
Call HandleError("in_long_arr")
Resume out:

End Function

Public Function in_str_arr(ByVal sSearch As String, arrString() As String) As Boolean
Dim nIndex As Integer
On Error GoTo error:

For nIndex = LBound(arrString) To UBound(arrString)
    If arrString(nIndex) = sSearch Then
        in_str_arr = True
        Exit For
    End If
Next

out:
On Error Resume Next
Exit Function
error:
Call HandleError("in_str_arr")
Resume out:

End Function

Public Sub ClearListViewSelections(ByRef LV As ListView)
Dim oLI As ListItem

On Error GoTo error:

For Each oLI In LV.ListItems
    oLI.Selected = False
    Set oLI = Nothing
Next

out:
Set oLI = Nothing
Exit Sub
error:
Call HandleError("ClearListViewSelections")
Resume out:

End Sub

'Public Function ClipNull(ByVal DataToClip As String, Optional ByVal length As Integer) As String
'On Error GoTo Error:
'Dim i As Long
'
'If length = 0 Then length = Len(DataToClip)
'
'For i = 1 To length
'    If Mid(DataToClip, i, 1) = Chr(0) Then
'        ClipNull = Mid(DataToClip, 1, i - 1)
'        Exit Function
'    End If
'Next i
'
'ClipNull = DataToClip
'
'Exit Function
'Error:
'Call HandleError("ClipNull")
'ClipNull = "error"
'End Function


Public Function ExtractNumbersFromString(ByVal sString As String) As Variant
Dim x As Integer, sNewString As String, bIgnoreDecimal As Boolean

On Error GoTo error:

ExtractNumbersFromString = 0
sNewString = ""
bIgnoreDecimal = False

For x = 1 To Len(sString)
    Select Case Mid(sString, x, 1)
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0":
            sNewString = sNewString & Mid(sString, x, 1)
        Case ".":
            If Not sNewString = "" And Not bIgnoreDecimal Then
                sNewString = sNewString & Mid(sString, x, 1)
                bIgnoreDecimal = True
            End If
        Case "-":
            If sNewString = "" Then
                sNewString = sNewString & Mid(sString, x, 1)
            End If
        Case Else:
            If sNewString = "-" Then
                sNewString = ""
            ElseIf Not sNewString = "" Then
                GoTo out:
            End If
    End Select
Next

out:
ExtractNumbersFromString = Val(sNewString)

Exit Function
error:
Call HandleError("ExtractNumbersFromString")

End Function

Public Function ExtractValueFromString(ByVal sWholeString As String, ByVal sSearchText As String) As Long
Dim x As Long, y As Long, sChar As String * 1

On Error GoTo error:

x = InStr(1, sWholeString, sSearchText, vbTextCompare)
If x > 0 Then
    x = x + Len(sSearchText) 'position x just after the search text
    y = x
    Do Until y > Len(sWholeString)
        sChar = Mid(sWholeString, y, 1)
        Select Case sChar
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
            Case " ":
                If y > x Then
                    Exit Do
                Else
                    x = x + 1
                End If
            Case Else: Exit Do
        End Select
        y = y + 1
    Loop
    If y > x Then ExtractValueFromString = Val(Mid(sWholeString, x, y - x))
    'If ExtractValueFromString = "0" Then ExtractValueFromString = ""
End If

out:
Exit Function
error:
Call HandleError("ExtractValueFromString")
Resume out:

End Function

Public Function FileExists(ByVal FileName As String) As Boolean
Dim fso As FileSystemObject

On Error GoTo error:

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(FileName) Then FileExists = True

out:
Set fso = Nothing
Exit Function
error:
Call HandleError("FileExists")
Resume out:
End Function

Public Function FolderExists(ByVal FolderPath As String) As Boolean
Dim fso As FileSystemObject

On Error GoTo error:

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FolderExists(FolderPath) Then FolderExists = True

out:
Set fso = Nothing
Exit Function
error:
Call HandleError("FileExists")
Resume out:
End Function

Public Sub GetTitleBarOffset()
Dim TitleInfo As TITLEBARINFO, OSVer As cnWin32Ver

On Error GoTo error:

OSVer = Win32Ver
If OSVer <= win95 Then GoTo win95:

TitleInfo.cbSize = Len(TitleInfo)
GetTitleBarInfo frmMain.hwnd, TitleInfo

TITLEBAR_OFFSET = (TitleInfo.rcTitleBar.Bottom * Screen.TwipsPerPixelY) - (TitleInfo.rcTitleBar.Top * Screen.TwipsPerPixelY)

If TITLEBAR_OFFSET > 285 Then '285 is the standard height
    TITLEBAR_OFFSET = TITLEBAR_OFFSET - 285
Else
    TITLEBAR_OFFSET = 0
End If

Exit Sub

win95:

TITLEBAR_OFFSET = 0

Exit Sub
error:
Call HandleError("GetTitleBarOffset")

End Sub

Public Sub HandleError(Optional ByVal ErrorSource As String)
Dim nYesNo As Integer

If bSuppressErrors Then
    Err.clear
    Exit Sub
End If

Select Case Err.Number
    Case 70:
        nYesNo = MsgBox("Error 70: File is locked by another process!" _
            & vbCrLf & "Terminate Application?", vbExclamation + vbYesNo + vbDefaultButton2)
    Case Else:
        If Len(ErrorSource) > 1 Then
            nYesNo = MsgBox("Error " & Err.Number & " in [" & ErrorSource & "]" & vbCrLf _
                & Err.Description & vbCrLf _
                & "Terminate Application?", vbCritical + vbYesNo + vbDefaultButton2)
        Else
            nYesNo = MsgBox("Error " & Err.Number & ": " & Err.Description _
                & vbCrLf & "Terminate Application?", vbYesNo + vbCritical + vbDefaultButton2)
        End If
End Select

If nYesNo = vbYes Then
    frmMain.bDontCallTerminate = True
    frmMain.bDontSaveSettings = True
    Call AppTerminate
    End
End If

Err.clear
End Sub

Private Function InvNumber(ByVal Number As String) As String
'*******************************************************************************
' Modifies a numeric string to allow it to be sorted alphabetically
'-------------------------------------------------------------------------------

    Static i As Integer
    For i = 1 To Len(Number)
        Select Case Mid$(Number, i, 1)
        Case "-": Mid$(Number, i, 1) = " "
        Case "0": Mid$(Number, i, 1) = "9"
        Case "1": Mid$(Number, i, 1) = "8"
        Case "2": Mid$(Number, i, 1) = "7"
        Case "3": Mid$(Number, i, 1) = "6"
        Case "4": Mid$(Number, i, 1) = "5"
        Case "5": Mid$(Number, i, 1) = "4"
        Case "6": Mid$(Number, i, 1) = "3"
        Case "7": Mid$(Number, i, 1) = "2"
        Case "8": Mid$(Number, i, 1) = "1"
        Case "9": Mid$(Number, i, 1) = "0"
        End Select
    Next
    InvNumber = Number
    
'*******************************************************************************
'
'-------------------------------------------------------------------------------
End Function

Public Function NumberKeysOnly(ByVal KeyAscii As Integer) As Integer
NumberKeysOnly = KeyAscii
If KeyAscii = 3 Or KeyAscii = 22 Then Exit Function 'control+v, control+c
If KeyAscii < 48 Or KeyAscii > 57 Then NumberKeysOnly = 0
If KeyAscii = 8 Then NumberKeysOnly = KeyAscii
If KeyAscii = 45 Then NumberKeysOnly = KeyAscii
End Function

Public Function PutCommas(ByVal sNumber As String) As String
On Error GoTo error:
Dim x As Integer, y As Integer, z As Integer

If Len(sNumber) < 4 Then
    PutCommas = sNumber
    Exit Function
End If

z = 1
y = Len(sNumber)
For x = 1 To y
    PutCommas = Mid(sNumber, y - x + 1, 1) & PutCommas
    
    If z > 2 And Not z = y Then
        If z Mod 3 = 0 Then PutCommas = "," & PutCommas
    End If
    
    z = z + 1
Next

Exit Function
error:
Call HandleError("PutCommas")
End Function

Public Function AutoPrepend(ByVal sStringToPrepend As String, ByVal sPrepend As String, Optional ByVal sGlue As String = ", ") As String
On Error GoTo error:

If Len(sStringToPrepend) > 0 Then
    If Len(Trim(sPrepend)) > 0 Then
        AutoPrepend = sPrepend & sGlue & sStringToPrepend
    Else
        AutoPrepend = sStringToPrepend
    End If
Else
    AutoPrepend = sPrepend
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("AutoAppend")
Resume out:
End Function

Public Function AutoAppend(ByVal sStringToAppend As String, ByVal sAppend As String, Optional ByVal sGlue As String = ", ") As String
On Error GoTo error:

If Len(sStringToAppend) > 0 Then
    If Len(Trim(sAppend)) > 0 Then
        AutoAppend = sStringToAppend & sGlue & sAppend
    Else
        AutoAppend = sStringToAppend
    End If
Else
    AutoAppend = sAppend
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("AutoAppend")
Resume out:
End Function

Public Function RemoveCharacter(ByVal DataToTest As String, ByVal sChar As String) As String
On Error GoTo error:
Dim i As Long

For i = 1 To Len(DataToTest)
    If Not Mid(DataToTest, i, 1) = sChar Then
        RemoveCharacter = RemoveCharacter & Mid(DataToTest, i, 1)
    End If
Next i

Exit Function
error:
Call HandleError("RemoveCharacter")
RemoveCharacter = "error"
End Function

Public Function RemoveVowles(ByVal sStr As String)
On Error GoTo error:
Dim x As Long, sChar As String

If Len(sStr) = 0 Then Exit Function

RemoveVowles = Mid(sStr, 1, 1)

'2 because commonly you want the first vowel
For x = 2 To Len(sStr)
    sChar = Mid(sStr, x, 1)
    Select Case sChar
        Case "a", "e", "i", "o", "u":
        Case Else:
            RemoveVowles = RemoveVowles & sChar
    End Select
Next

Exit Function
error:
Call HandleError("HandleError")
End Function

Public Function RoundUp(ByVal nNumber As Double) As Double

On Error GoTo error:

If 0 < nNumber - Int(nNumber) Then
    RoundUp = Int(nNumber) + 1
Else
    RoundUp = Int(nNumber)
End If

out:
Exit Function
error:
Call HandleError("RoundUp")
Resume out:

End Function


Public Sub SortListView(ListView As ListView, ByVal Index As Integer, ByVal DataType As ListDataType, ByVal Ascending As Boolean)

'*******************************************************************************
' Sort a ListView by String, Number, or DateTime
'
' Parameters:
'
'   ListView    Reference to the ListView control to be sorted.
'   Index       Index of the column in the ListView to be sorted. The first
'               column in a ListView has an index value of 1.
'   DataType    Sets whether the data in the column is to be sorted
'               alphabetically, numerically, or by date.
'   Ascending   Sets the direction of the sort. True sorts A-Z (Ascending),
'               and False sorts Z-A (descending)
'-------------------------------------------------------------------------------

    On Error Resume Next
    Dim i As Integer
    Dim l As Long
    Dim strFormat As String
    
    ' Display the hourglass cursor whilst sorting
    
    Dim lngCursor As Long
    lngCursor = ListView.MousePointer
    ListView.MousePointer = vbHourglass
    
    ' Prevent the ListView control from updating on screen - this is to hide
    ' the changes being made to the listitems, and also to speed up the sort
    
    If ListView.ListItems.Count > 75 Then LockWindowUpdate frmMain.hwnd 'ListView.hWnd
    
    Dim blnRestoreFromTag As Boolean
    
    Select Case DataType
    Case ldtstring
        
        ' Sort alphabetically. This is the only sort provided by the
        ' MS ListView control (at this time), and as such we don't really
        ' need to do much here
    
        blnRestoreFromTag = False
        
    Case ldtnumber
    
        ' Sort Numerically
    
        strFormat = String$(20, "0") & "." & String$(10, "0")
        
        ' Loop through the values in this column. Re-format the values so
        ' as they can be sorted alphabetically, having already stored their
        ' text values in the tag, along with the tag's original value
    
        With ListView.ListItems
            If (Index = 1) Then
                For l = 1 To .Count
                    With .item(l)
                        .Tag = .Text & Chr$(0) & .Tag
'                        If IsNumeric(.Text) Then
                            If CDbl(Val(.Text)) >= 0 Then
                                .Text = Format(CDbl(Val(.Text)), strFormat)
                            Else
                                .Text = "&" & InvNumber(Format(0 - CDbl(Val(.Text)), strFormat))
                            End If
'                        Else
'                            .Text = ""
'                        End If
                    End With
                Next l
            Else
                For l = 1 To .Count
                    With .item(l).ListSubItems(Index - 1)
                        .Tag = .Text & Chr$(0) & .Tag
'                        If IsNumeric(.Text) Then
                            If CDbl(Val(.Text)) >= 0 Then
                                .Text = Format(CDbl(Val(.Text)), strFormat)
                            Else
                                .Text = "&" & InvNumber(Format(0 - CDbl(Val(.Text)), strFormat))
                            End If
'                        Else
'                            .Text = ""
'                        End If
                    End With
                Next l
            End If
        End With
        
        blnRestoreFromTag = True
    
    Case ldtDateTime
    
        ' Sort by date.
        
        strFormat = "YYYYMMDDHhNnSs"
        
        Dim dte As Date
    
        ' Loop through the values in this column. Re-format the dates so as they
        ' can be sorted alphabetically, having already stored their visible
        ' values in the tag, along with the tag's original value
    
        With ListView.ListItems
            If (Index = 1) Then
                For l = 1 To .Count
                    With .item(l)
                        .Tag = .Text & Chr$(0) & .Tag
                        dte = CDate(.Text)
                        .Text = Format$(dte, strFormat)
                    End With
                Next l
            Else
                For l = 1 To .Count
                    With .item(l).ListSubItems(Index - 1)
                        .Tag = .Text & Chr$(0) & .Tag
                        dte = CDate(.Text)
                        .Text = Format$(dte, strFormat)
                    End With
                Next l
            End If
        End With
        
        blnRestoreFromTag = True
        
    End Select
    
    ' Sort the ListView Alphabetically
    
    ListView.SortOrder = IIf(Ascending, lvwAscending, lvwDescending)
    ListView.SortKey = Index - 1
    ListView.Sorted = True
    
    ' Restore the Text Values if required
    
    If blnRestoreFromTag Then
        
        ' Restore the previous values to the 'cells' in this column of the list
        ' from the tags, and also restore the tags to their original values
        
        With ListView.ListItems
            If (Index = 1) Then
                For l = 1 To .Count
                    With .item(l)
                        i = InStr(.Tag, Chr$(0))
                        .Text = Left$(.Tag, i - 1)
                        .Tag = Mid$(.Tag, i + 1)
                    End With
                Next l
            Else
                For l = 1 To .Count
                    With .item(l).ListSubItems(Index - 1)
                        i = InStr(.Tag, Chr$(0))
                        .Text = Left$(.Tag, i - 1)
                        .Tag = Mid$(.Tag, i + 1)
                    End With
                Next l
            End If
        End With
    End If
    
    ' Unlock the list window so that the OCX can update it
    
    LockWindowUpdate 0&
    
    ' Restore the previous cursor
    
    ListView.MousePointer = lngCursor
    

End Sub

Public Sub SortListViewByTag(ListView As ListView, ByVal Index As Integer, ByVal DataType As ListDataType, ByVal Ascending As Boolean)

'*******************************************************************************
' Sort a ListView by String, Number, or DateTime
'
' Parameters:
'
'   ListView    Reference to the ListView control to be sorted.
'   Index       Index of the column in the ListView to be sorted. The first
'               column in a ListView has an index value of 1.
'   DataType    Sets whether the data in the column is to be sorted
'               alphabetically, numerically, or by date.
'   Ascending   Sets the direction of the sort. True sorts A-Z (Ascending),
'               and False sorts Z-A (descending)
'-------------------------------------------------------------------------------

    On Error Resume Next
    Dim i As Integer
    Dim l As Long
    Dim strFormat As String
    
    ' Display the hourglass cursor whilst sorting
    
    Dim lngCursor As Long
    lngCursor = ListView.MousePointer
    ListView.MousePointer = vbHourglass
    
    ' Prevent the ListView control from updating on screen - this is to hide
    ' the changes being made to the listitems, and also to speed up the sort
    
    If ListView.ListItems.Count > 75 Then LockWindowUpdate frmMain.hwnd 'ListView.hWnd
    
    Dim blnRestoreFromTag As Boolean
    
    Select Case DataType
    Case ldtstring
        
        ' Sort alphabetically. This is the only sort provided by the
        ' MS ListView control (at this time), and as such we don't really
        ' need to do much here
    
        blnRestoreFromTag = False
        
    Case ldtnumber
    
        ' Sort Numerically
    
        strFormat = String$(20, "0") & "." & String$(10, "0")
        
        ' Loop through the values in this column. Re-format the values so
        ' as they can be sorted alphabetically, having already stored their
        ' text values in the tag, along with the tag's original value
    
        With ListView.ListItems
            If (Index = 1) Then
                For l = 1 To .Count
                    With .item(l)
                        '.Tag = .Text & Chr$(0) & .Tag
                        .Tag = .Tag & Chr$(0) & .Text
'                        If IsNumeric(.Text) Then
                            If CDbl(Val(Replace(.Text, "%", ""))) >= 0 Then
                                .Text = Format(CDbl(Val(.Tag)), strFormat)
                            Else
                                .Text = "&" & InvNumber(Format(0 - CDbl(Val(.Tag)), strFormat))
                            End If
'                        Else
'                            .Text = ""
'                        End If
                    End With
                Next l
            Else
                For l = 1 To .Count
                    With .item(l).ListSubItems(Index - 1)
                        '.Tag = .Text & Chr$(0) & .Tag
                        .Tag = .Tag & Chr$(0) & .Text
'                        If IsNumeric(.Text) Then
                            If CDbl(Val(Replace(.Text, "%", ""))) >= 0 Then
                                .Text = Format(CDbl(Val(.Tag)), strFormat)
                            Else
                                .Text = "&" & InvNumber(Format(0 - CDbl(Val(.Tag)), strFormat))
                            End If
'                        Else
'                            .Text = ""
'                        End If
                    End With
                Next l
            End If
        End With
        
        blnRestoreFromTag = True
    
    Case ldtDateTime
    
        ' Sort by date.
        
        strFormat = "YYYYMMDDHhNnSs"
        
        Dim dte As Date
    
        ' Loop through the values in this column. Re-format the dates so as they
        ' can be sorted alphabetically, having already stored their visible
        ' values in the tag, along with the tag's original value
    
        With ListView.ListItems
            If (Index = 1) Then
                For l = 1 To .Count
                    With .item(l)
                        .Tag = .Text & Chr$(0) & .Tag
                        dte = CDate(.Text)
                        .Text = Format$(dte, strFormat)
                    End With
                Next l
            Else
                For l = 1 To .Count
                    With .item(l).ListSubItems(Index - 1)
                        .Tag = .Text & Chr$(0) & .Tag
                        dte = CDate(.Text)
                        .Text = Format$(dte, strFormat)
                    End With
                Next l
            End If
        End With
        
        blnRestoreFromTag = True
        
    End Select
    
    ' Sort the ListView Alphabetically
    
    ListView.SortOrder = IIf(Ascending, lvwAscending, lvwDescending)
    ListView.SortKey = Index - 1
    ListView.Sorted = True
    
    ' Restore the Text Values if required
    
    If blnRestoreFromTag Then
        
        ' Restore the previous values to the 'cells' in this column of the list
        ' from the tags, and also restore the tags to their original values
        
        With ListView.ListItems
            If (Index = 1) Then
                For l = 1 To .Count
                    With .item(l)
                        i = InStr(.Tag, Chr$(0))
                        .Text = Mid$(.Tag, i + 1)
                        .Tag = Left$(.Tag, i - 1)
                    End With
                Next l
            Else
                For l = 1 To .Count
                    With .item(l).ListSubItems(Index - 1)
                        i = InStr(.Tag, Chr$(0))
                        .Text = Mid$(.Tag, i + 1)
                        .Tag = Left$(.Tag, i - 1)
                    End With
                Next l
            End If
        End With
    End If
    
    ' Unlock the list window so that the OCX can update it
    
    LockWindowUpdate 0&
    
    ' Restore the previous cursor
    
    ListView.MousePointer = lngCursor
    

End Sub

Public Sub UnloadForms(ByVal sDontUnload As String)
On Error GoTo error:
Dim objFrm As Form

For Each objFrm In Forms
    If Not objFrm.name = sDontUnload And Not objFrm.name = "frmMain" Then Unload objFrm
Next

If Not sDontUnload = "frmMain" Then
    Unload frmMain
End If

Set objFrm = Nothing
Exit Sub
error:
Call HandleError("HandleError")
Set objFrm = Nothing
End Sub

Public Function FormIsLoaded(ByVal sFormName As String) As Boolean
On Error GoTo error:
Dim objFrm As Form

For Each objFrm In Forms
    If objFrm.name = sFormName Then
        FormIsLoaded = True
        Exit For
    End If
Next objFrm

Set objFrm = Nothing
Exit Function
error:
Call HandleError("FormIsLoaded")
Set objFrm = Nothing
End Function

Public Function PutCrLF(ByVal sString As String) As String
Dim x As Integer, y As Integer

On Error GoTo error:

y = InStr(1, sString, Chr(10))
If y = 0 Then
    PutCrLF = sString
    Exit Function
End If

x = 1
Do While x < Len(sString)
    y = InStr(x, sString, Chr(10))
    If y = 0 Then
        PutCrLF = PutCrLF & Mid(sString, x)
        Exit Do
    End If
    PutCrLF = PutCrLF & Mid(sString, x, y - x) & vbCrLf
    x = y + 1
Loop

Exit Function

error:
Call HandleError("PutCrLF")

End Function

'**************************************
' Name: a AutoComplete Very Simple!
' Description:VERY SIMPLE cut and paste
'     funtion for the Keypress event of a comb
'     obox. Just paste this code into a module
'     or form and call the function from the K
'     eyPress event. KeyAscii = AutoComplete(c
'     boCombobox, KeyAscii,Optional UpperCase)
'
' By: James Berard
'
'
' Inputs:None
'
' Returns:None
'
'Assumes:None
'
'Side Effects:None
'This code is copyrighted and has limite
'     d warranties.
'Please see http://www.Planet-Source-Cod
'     e.com/xq/ASP/txtCodeId.43911/lngWId.1/qx
'     /vb/scripts/ShowCode.htm
'for details.
'**************************************

Public Function AutoComplete(cbCombo As ComboBox, sKeyAscii As Integer, Optional bMatchCase As Boolean) As Integer
    Dim lngFind As Long, intPos As Integer, intLength As Integer
    Dim tStr As String, intCurrent As Integer


    With cbCombo
        intCurrent = IIf(.ListIndex >= 0, .ListIndex, 0)
        
        If sKeyAscii = 8 Then
            If .SelStart <= 1 Then
                .ListIndex = 0
                AutoComplete = 0
                Exit Function
            End If
            .SelStart = .SelStart - 1
            .SelLength = 32000
            .SelText = ""
        Else
            intPos = .SelStart '// save intial cursor position
            tStr = .Text '// save string

            If bMatchCase = True Then
                .SelText = Chr(sKeyAscii) '// change string. (leave case alone)
            Else
                .SelText = LCase(Chr(sKeyAscii)) '// change string. (lowercase only)
            End If
        End If
        
        lngFind = SendMessage(.hwnd, CB_FINDSTRING, 0, ByVal .Text) '// Find string in combobox

        If lngFind = -1 Then '// if string not found
            .ListIndex = intCurrent
            .Text = tStr '// set old string (used for boxes that require charachter monitoring
            .SelStart = intPos '// set cursor position
            .SelLength = (Len(.Text) - intPos) '// set selected length
            AutoComplete = 0 '// return 0 value to KeyAscii
            Exit Function
            
        Else '// If string found
            intPos = .SelStart '// save cursor position
            intLength = Len(.List(lngFind)) - Len(.Text) '// save remaining highlighted text length
            .ListIndex = lngFind
            '.SelText = .SelText & Right(.List(lngFind), intLength) '// change new text in string
            '.Text = .List(lngFind)'// Use this inst
            '     ead of the above .Seltext line to set th
            '     e text typed to the exact case of the it
            '     em selected in the combo box.
            .SelStart = intPos '// set cursor position
            .SelLength = intLength '// set selected length
        End If
    End With
    
End Function


Public Function IsAlphaNumeric(TestString As String) As Boolean

Dim sTemp As String
Dim iLen As Integer
Dim iCtr As Integer
Dim sChar As String

'returns true if all characters in a string are alphabetical
' or numeric
'returns false otherwise or for empty string

sTemp = TestString
iLen = Len(sTemp)
If iLen > 0 Then
    For iCtr = 1 To iLen
        sChar = Mid(sTemp, iCtr, 1)
        If Not sChar Like "[0-9A-Za-z]" Then Exit Function
    Next
    
    IsAlphaNumeric = True
End If

End Function

Public Function IsAlphaBetical(TestString As String, Optional bAllowSpace As Boolean) As Boolean

Dim sTemp As String
Dim iLen As Integer
Dim iCtr As Integer
Dim sChar As String

'returns true if all characters in a string are alphabetical
'returns false otherwise or for empty string

sTemp = TestString
iLen = Len(sTemp)
If iLen > 0 Then
    For iCtr = 1 To iLen
        sChar = Mid(sTemp, iCtr, 1)
        If bAllowSpace Then
            If Not sChar Like "[A-Za-z ]" Then Exit Function
        Else
            If Not sChar Like "[A-Za-z]" Then Exit Function
        End If
    Next
    
    IsAlphaBetical = True
End If
End Function
Function RegExpFind(LookIn As String, PatternStr As String, Optional pos, _
    Optional MatchCase As Boolean = True, Optional ReturnType As Long = 0, _
    Optional MultiLine As Boolean = False) As String()
    
    ' Function written by Patrick G. Matthews.  You may use and distribute this code freely,
    ' as long as you properly credit and attribute authorship and the URL of where you
    ' found the code
    
    ' This function relies on the VBScript version of Regular Expressions, and thus some of
    ' the functionality available in Perl and/or .Net may not be available.  The full extent
    ' of what functionality will be available on any given computer is based on which version
    ' of the VBScript runtime is installed on that computer
    
    ' This function uses Regular Expressions to parse a string (LookIn), and return matches to a
    ' pattern (PatternStr).  Use Pos to indicate which match you want:
    ' Pos omitted               : function returns a zero-based array of all matches
    ' Pos = 1                   : the first match
    ' Pos = 2                   : the second match
    ' Pos = <positive integer>  : the Nth match
    ' Pos = 0                   : the last match
    ' Pos = -1                  : the last match
    ' Pos = -2                  : the 2nd to last match
    ' Pos = <negative integer>  : the Nth to last match
    ' If Pos is non-numeric, or if the absolute value of Pos is greater than the number of
    ' matches, the function returns an empty string.  If no match is found, the function returns
    ' an empty string.  (Earlier versions of this code used zero for the last match; this is
    ' retained for backward compatibility)
    
    ' If MatchCase is omitted or True (default for RegExp) then the Pattern must match case (and
    ' thus you may have to use [a-zA-Z] instead of just [a-z] or [A-Z]).
    
    ' ReturnType indicates what information you want to return:
    ' ReturnType = 0            : the matched values
    ' ReturnType = 1            : the starting character positions for the matched values
    ' ReturnType = 2            : the lengths of the matched values
    
    ' If you use this function in Excel, you can use range references for any of the arguments.
    ' If you use this in Excel and return the full array, make sure to set up the formula as an
    ' array formula.  If you need the array formula to go down a column, use TRANSPOSE()
    
    ' Note: RegExp counts the character positions for the Match.FirstIndex property as starting
    ' at zero.  Since VB6 and VBA has strings starting at position 1, I have added one to make
    ' the character positions conform to VBA/VB6 expectations
    
    ' Normally as an object variable I would set the RegX variable to Nothing; however, in cases
    ' where a large number of calls to this function are made, making RegX a static variable that
    ' preserves its state in between calls significantly improves performance
    
    Static RegX As Object
    Dim TheMatches As Object
    Dim Answer() As String
    Dim Counter As Long
    
    ' Evaluate Pos.  If it is there, it must be numeric and converted to Long
    ReDim RegExpFind(0)
    
    If Not IsMissing(pos) Then
        If Not IsNumeric(pos) Then
            Exit Function
        Else
            pos = CLng(pos)
        End If
    End If
    
    ' Evaluate ReturnType
    
    If ReturnType < 0 Or ReturnType > 2 Then
        Exit Function
    End If
    
    ' Create instance of RegExp object if needed, and set properties
    
    If RegX Is Nothing Then Set RegX = CreateObject("VBScript.RegExp")
    With RegX
        .Pattern = PatternStr
        .Global = True
        .IgnoreCase = Not MatchCase
        .MultiLine = MultiLine
    End With
        
    ' Test to see if there are any matches
    
    If RegX.Test(LookIn) Then
        
        ' Run RegExp to get the matches, which are returned as a zero-based collection
        
        Set TheMatches = RegX.Execute(LookIn)
        
        ' Test to see if Pos is negative, which indicates the user wants the Nth to last
        ' match.  If it is, then based on the number of matches convert Pos to a positive
        ' number, or zero for the last match
        
        If Not IsMissing(pos) Then
            If pos < 0 Then
                If pos = -1 Then
                    pos = 0
                Else
                    
                    ' If Abs(Pos) > number of matches, then the Nth to last match does not
                    ' exist.  Return a zero-length string
                    
                    If Abs(pos) <= TheMatches.Count Then
                        pos = TheMatches.Count + pos + 1
                    Else
                        GoTo Cleanup
                    End If
                End If
            End If
        End If
        
        ' If Pos is missing, user wants array of all matches.  Build it and assign it as the
        ' function's return value
        
        If IsMissing(pos) Then
            ReDim Answer(0 To TheMatches.Count - 1)
            For Counter = 0 To UBound(Answer)
                Select Case ReturnType
                    Case 0: Answer(Counter) = TheMatches(Counter)
                    Case 1: Answer(Counter) = TheMatches(Counter).FirstIndex + 1
                    Case 2: Answer(Counter) = TheMatches(Counter).length
                End Select
            Next
            RegExpFind = Answer
        
        ' User wanted the Nth match (or last match, if Pos = 0).  Get the Nth value, if possible
        
        Else
            Select Case pos
                Case 0                          ' Last match
                    Select Case ReturnType
                        Case 0: RegExpFind = TheMatches(TheMatches.Count - 1)
                        Case 1: RegExpFind = TheMatches(TheMatches.Count - 1).FirstIndex + 1
                        Case 2: RegExpFind = TheMatches(TheMatches.Count - 1).length
                    End Select
                Case 1 To TheMatches.Count      ' Nth match
                    Select Case ReturnType
                        Case 0: RegExpFind = TheMatches(pos - 1)
                        Case 1: RegExpFind = TheMatches(pos - 1).FirstIndex + 1
                        Case 2: RegExpFind = TheMatches(pos - 1).length
                    End Select
                Case Else                       ' Invalid item number
                    'nada
            End Select
        End If
    
    ' If there are no matches, return empty string
    
    Else
        'nada
    End If
    
Cleanup:
    ' Release object variables
    
    Set TheMatches = Nothing
    
End Function

Function EscapeRegex(sText As String) As String
On Error GoTo error:

EscapeRegex = sText
EscapeRegex = Replace(EscapeRegex, "(", "\(")
EscapeRegex = Replace(EscapeRegex, ")", "\)")
EscapeRegex = Replace(EscapeRegex, "[", "\[")
EscapeRegex = Replace(EscapeRegex, "]", "\]")
EscapeRegex = Replace(EscapeRegex, ".", "\.")
EscapeRegex = Replace(EscapeRegex, "$", "\$")
EscapeRegex = Replace(EscapeRegex, "^", "\^")

out:
On Error Resume Next
Exit Function
error:
Call HandleError("EscapeRegexPattern")
Resume out:
End Function

