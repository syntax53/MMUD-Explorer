Attribute VB_Name = "modMenuSubClass"
Option Explicit

'Public API Declarations
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Public Constants
Public Const GWL_WNDPROC = -4
Public Const WM_COMMAND = &H111

'Public Variables
Public nMenuItemID As Integer 'holds the item identification number of the newly added menu items
Public oldWindowProc As Long 'a pointer to this form's old window procedure

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long 'Processes window messages
    'There is no way for Visual Basic to create an event handler
    'to process whatever functions that need to be performed by the
    'newly created menu items. To work around this problem, it is necessary
    'to create this 'WindowProc' function to manually process the WM_COMMAND
    'messages that the new menu items send to the form's window...
    
    Dim retval As Long  'holds the return value

    If uMsg = WM_COMMAND Then
        If wParam >= 1000 Then 'if the window command was received from one of our new menu items
            'This is where you set up event handling for our new menu items.
            'EXAMPLE:
            Select Case wParam
'                Case 1000: 'First New Menu Item (be careful, the item may be a separator bar!)
'                    Do Something
                Case 1001: 'Second New Menu Item
                    Call frmMain.RecentFilesLoad(1)
                Case 1002: 'Third New Menu Item
                    Call frmMain.RecentFilesLoad(2)
                Case 1003:
                    Call frmMain.RecentFilesLoad(3)
                Case 1004:
                    Call frmMain.RecentFilesLoad(4)
                Case 1005:
                    Call frmMain.RecentFilesLoad(5)
            End Select
            
            'Sample event handling (changes the form's background color):
            'Randomize
            'frmMain.BackColor = QBColor(CInt(Rnd * 15))
        End If
    End If
    retval = CallWindowProc(oldWindowProc, hWnd, uMsg, wParam, lParam) 'use this form's original window procedure to finish processing this message
    WindowProc = retval 'set the WindowProc function equal to whatever this form's original window procedure would have returned
End Function
