Attribute VB_Name = "modWin32Ver"
Option Explicit
'#####################################################################################
'#  Determine the Win32 Operating System Version Via API (modWin32Ver.bas)
'#      By: Nick Campbeln
'#
'#      Revision History:
'#          1.0.3 (April 18, 2025):
'#              ChatGPT and syntax53 updated to include newer operating systems
'#          1.0.2 (Aug 11, 2002):
'#              Switched GetVersionEx() form Public to Private
'#          1.0.1 (Aug 6, 2002):
'#              Fixed a (very) stupid coding error in isWin2k() - Renamed function from isWin2000() to isWin2k() and forgot to change the return values in the function to the same name - D'oh!
'#          1.0 (Aug 4, 2002):
'#              Initial Release
'#
'#      Copyright � 2002 Nick Campbeln (opensource@nick.campbeln.com)
'#          This source code is provided 'as-is', without any express or implied warranty. In no event will the author(s) be held liable for any damages arising from the use of this source code. Permission is granted to anyone to use this source code for any purpose, including commercial applications, and to alter it and redistribute it freely, subject to the following restrictions:
'#          1. The origin of this source code must not be misrepresented; you must not claim that you wrote the original source code. If you use this source code in a product, an acknowledgment in the product documentation would be appreciated but is not required.
'#          2. Altered source versions must be plainly marked as such, and must not be misrepresented as being the original source code.
'#          3. This notice may not be removed or altered from any source distribution.
'#              (NOTE: This license is borrowed from zLib.)
'#
'#  Please remember to vote on PSC.com if you like this code!
'#  Code URL: http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=37628&lngWId=1
'#####################################################################################

    '#### Functions/Consts/Types used for GetWin32Ver()
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Type OSVERSIONINFO
    OSVSize As Long
    dwVerMajor As Long
    dwVerMinor As Long
    dwBuildNumber As Long           '#### NT: Build Number, 9x: High-Order has Major/Minor ver, Low-Order has build
    PlatformID As Long
    szCSDVersion As String * 128    '#### NT: ie- "Service Pack 3", 9x: 'arbitrary additional information'
End Type

Public Enum cnWin32Ver
    UnknownOS = 0
    win95 = 1
    Win98 = 2
    WinME = 3
    WinNT4 = 4
    Win2k = 5
    WinXP = 6
    WinVista = 7
    Win7 = 8
    Win8 = 9
    Win8_1 = 10
    Win10 = 11
    Win11 = 12
End Enum


'#####################################################################################
'# Public subs/functions
'#####################################################################################
'#########################################################
'# Returns the asso. cnWin32Ver eNum value of the current Win32 OS
'#########################################################
Public Function GetWin32Ver() As cnWin32Ver
    Dim oOSV As OSVERSIONINFO
    oOSV.OSVSize = Len(oOSV)
   
        '#### If the API returned a valid value
    If GetVersionEx(oOSV) = 1 Then
        Select Case oOSV.PlatformID
            Case VER_PLATFORM_WIN32_NT
                Select Case True
                    Case oOSV.dwVerMajor = 10 And oOSV.dwVerMinor = 0
                        ' Windows 10 vs Windows 11 differentiation by build number
                        If oOSV.dwBuildNumber >= 22000 Then
                            GetWin32Ver = Win11
                        Else
                            GetWin32Ver = Win10
                        End If
                    Case oOSV.dwVerMajor = 6 And oOSV.dwVerMinor = 3
                        GetWin32Ver = Win8_1
                    Case oOSV.dwVerMajor = 6 And oOSV.dwVerMinor = 2
                        GetWin32Ver = Win8
                    Case oOSV.dwVerMajor = 6 And oOSV.dwVerMinor = 1
                        GetWin32Ver = Win7
                    Case oOSV.dwVerMajor = 6 And oOSV.dwVerMinor = 0
                        GetWin32Ver = WinVista
                    Case oOSV.dwVerMajor = 5 And oOSV.dwVerMinor = 1
                        GetWin32Ver = WinXP
                    Case oOSV.dwVerMajor = 5 And oOSV.dwVerMinor = 0
                        GetWin32Ver = Win2k
                    Case oOSV.dwVerMajor = 4
                        GetWin32Ver = WinNT4
                    Case Else
                        GetWin32Ver = UnknownOS
                End Select
            Case VER_PLATFORM_WIN32_WINDOWS
                Select Case oOSV.dwVerMinor
                    Case 90
                        GetWin32Ver = WinME
                    Case Is > 0
                        GetWin32Ver = Win98
                    Case 0
                        GetWin32Ver = win95
                    Case Else
                        GetWin32Ver = UnknownOS
                End Select
            Case Else
                GetWin32Ver = UnknownOS
        End Select

'replaced with above 2025.04.18
'            '#### If we're running WinXP
'            '####    If VER_PLATFORM_WIN32_NT, dwVerMajor is 5 and dwVerMinor is 1, it's WinXP
'        If (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 5 And oOSV.dwVerMinor = 1) Then
'           GetWin32Ver = WinXP
'
'            '#### If we're running WinNT2000 (NT5)
'            '####    If VER_PLATFORM_WIN32_NT, dwVerMajor is 5 and dwVerMinor is 0, it's Win2k
'        ElseIf (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 5 And oOSV.dwVerMinor = 0) Then
'           GetWin32Ver = Win2k
'
'            '#### If we're running WinNT4
'            '####    If VER_PLATFORM_WIN32_NT and dwVerMajor is 4
'        ElseIf (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 4) Then
'           GetWin32Ver = WinNT4
'
'            '#### If we're running Windows ME
'            '####    If VER_PLATFORM_WIN32_WINDOWS and
'            '####    dwVerMajor = 4,  and dwVerMinor > 0, return true
'        ElseIf (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS And oOSV.dwVerMajor = 4 And oOSV.dwVerMinor = 90) Then
'           GetWin32Ver = WinME
'
'            '#### If we're running Win98
'            '####    If VER_PLATFORM_WIN32_WINDOWS and
'            '####    dwVerMajor => 4, or dwVerMajor = 4 and
'            '####    dwVerMinor > 0, return true
'        ElseIf (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And (oOSV.dwVerMajor > 4) Or (oOSV.dwVerMajor = 4 And oOSV.dwVerMinor > 0) Then
'           GetWin32Ver = Win98
'
'            '#### If we're running Win95
'            '####    If VER_PLATFORM_WIN32_WINDOWS and
'            '####    dwVerMajor = 4, and dwVerMinor = 0,
'        ElseIf (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS And oOSV.dwVerMajor = 4 And oOSV.dwVerMinor = 0) Then
'           GetWin32Ver = win95
'
'            '#### Else the OS is not reconized by this function
'        Else
'            GetWin32Ver = UnknownOS
'        End If
'
'        '#### Else the OS is not reconized by this function
    Else
        GetWin32Ver = UnknownOS
    End If
End Function


'#########################################################
'# Returns true if the OS is WinNT4, Win2k or WinXP
'#########################################################
Public Function isNT() As Boolean
        '#### Determine the return value of GetWin32Ver() and set the return value accordingly
    Select Case GetWin32Ver()
        Case WinNT4, Win2k, WinXP
            isNT = True
        Case Else
            isNT = False
    End Select
End Function


'#########################################################
'# Returns true if the OS is Win95, Win98 or WinME
'#########################################################
Public Function is9x() As Boolean
        '#### Determine the return value of GetWin32Ver() and set the return value accordingly
    Select Case GetWin32Ver()
        Case win95, Win98, WinME
            is9x = True
        Case Else
            is9x = False
    End Select
End Function


'#########################################################
'# Returns true if the OS is WinXP
'#########################################################
Public Function isWinXP() As Boolean
    Dim oOSV As OSVERSIONINFO
    oOSV.OSVSize = Len(oOSV)

        '#### If the API returned a valid value
    If (GetVersionEx(oOSV) = 1) Then
        isWinXP = (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 5 And oOSV.dwVerMinor = 1)
    End If
End Function


'#########################################################
'# Returns true if the OS is Win2k
'#########################################################
Public Function isWin2k() As Boolean
    Dim oOSV As OSVERSIONINFO
    oOSV.OSVSize = Len(oOSV)

        '#### If the API returned a valid value
    If (GetVersionEx(oOSV) = 1) Then
        isWin2k = (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 5 And oOSV.dwVerMinor = 0)
    End If
End Function


'#########################################################
'# Returns true if the OS is WinNT4
'#########################################################
Public Function isWinNT4() As Boolean
    Dim oOSV As OSVERSIONINFO
    oOSV.OSVSize = Len(oOSV)

        '#### If the API returned a valid value
    If (GetVersionEx(oOSV) = 1) Then
        isWinNT4 = (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 4)
    End If
End Function


'#########################################################
'# Returns true if the OS is WinME
'#########################################################
Public Function isWinME() As Boolean
    Dim oOSV As OSVERSIONINFO
    oOSV.OSVSize = Len(oOSV)

        '#### If the API returned a valid value
    If (GetVersionEx(oOSV) = 1) Then
        isWinME = (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS And oOSV.dwVerMajor = 4 And oOSV.dwVerMinor = 90)
    End If
End Function


'#########################################################
'# Returns true if the OS is Win98
'#########################################################
Public Function isWin98() As Boolean
    Dim oOSV As OSVERSIONINFO
    oOSV.OSVSize = Len(oOSV)

        '#### If the API returned a valid value
    If (GetVersionEx(oOSV) = 1) Then
         isWin98 = (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And (oOSV.dwVerMajor > 4) Or (oOSV.dwVerMajor = 4 And oOSV.dwVerMinor > 0)
    End If
End Function


'#########################################################
'# Returns true if the OS is Win95
'#########################################################
Public Function isWin95() As Boolean
    Dim oOSV As OSVERSIONINFO
    oOSV.OSVSize = Len(oOSV)

        '#### If the API returned a valid value
    If (GetVersionEx(oOSV) = 1) Then
         isWin95 = (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS And oOSV.dwVerMajor = 4 And oOSV.dwVerMinor = 0)
    End If
End Function
