Attribute VB_Name = "modSettings"

Global INIReadOnly As Boolean
Global INIFileName As String

Private Ret As String

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function ReadINI(ByVal Section As String, ByVal Key As String, Optional ByVal AlternateINIFile As String) As Variant
On Error GoTo Error:

If AlternateINIFile = "" Then AlternateINIFile = INIFileName

reread:
Ret = Space$(255)
retlen = GetPrivateProfileString(Section, Key, "", Ret, Len(Ret), ByVal AlternateINIFile)
If retlen = 0 Then
    
    If INIReadOnly = True Then Exit Function
    Call WriteINI(Section, Key, 0, AlternateINIFile)
    
    Select Case UCase(Section)
        Case "SETTINGS":
            If UCase(Key) = "DATAFILE" Then Call WriteINI(Section, Key, "data-v1.11n.mdb", ByVal AlternateINIFile)
            If Left(UCase(Key), 4) = "LOAD" Then Call WriteINI(Section, Key, 1, ByVal AlternateINIFile)
            If UCase(Key) = "ONLYINGAME" Then Call WriteINI(Section, Key, 1, ByVal AlternateINIFile)
            If Left(UCase(Key), 9) = "INVENSTAT" Then Call WriteINI(Section, Key, 1, ByVal AlternateINIFile)
    End Select
    
    GoTo reread:
End If

Ret = Left$(Ret, retlen)
ReadINI = Ret

If Left(UCase(Key), 6) = Left(UCase("Global"), 6) And Val(ReadINI) < 0 Then
    If INIReadOnly = True Then Exit Function
    Call WriteINI(Section, Key, 0, AlternateINIFile): GoTo reread:
End If

Exit Function
Error:
HandleError
End Function

Public Sub WriteINI(ByVal Section As String, ByVal Key As String, ByVal Text As String, Optional ByVal AlternateINIFile As String)
On Error GoTo Error:

If AlternateINIFile = "" Then AlternateINIFile = INIFileName

If INIReadOnly = True Then Exit Sub
Call WritePrivateProfileString(Section, Key, Text, ByVal AlternateINIFile)

Exit Sub
Error:
HandleError
End Sub

Public Sub CreateSettings()
On Error GoTo Error:
Dim fso As FileSystemObject
Dim sAppPath As String

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(INIFileName) Then fso.DeleteFile INIFileName, True

fso.CreateTextFile INIFileName, True

If Right(App.Path, 1) = "\" Then
    sAppPath = App.Path
Else
    sAppPath = App.Path & "\"
End If

Call WriteINI("Settings", "DataFile", "data-v1.11o.mdb")

Set fso = Nothing
        
Exit Sub
Error:
HandleError
Set fso = Nothing
End Sub

Public Sub CheckINIReadOnly()
Dim fso As FileSystemObject, nYesNo As Integer, oFile As File
On Error GoTo Error:

INIReadOnly = False

Set fso = CreateObject("Scripting.FileSystemObject")
Set oFile = fso.GetFile(INIFileName)

If oFile.Attributes And ReadOnly Then
    INIReadOnly = True
    nYesNo = MsgBox("settings.ini is marked 'read only,' attempt to fix?" & vbCrLf & "(settings cannot be saved otherwise)", vbYesNo, "settings.ini is read-only...")
    If Not nYesNo = vbNo Then
        oFile.Attributes = oFile.Attributes - 1
        INIReadOnly = False
    End If
End If

Set oFile = Nothing
Set fso = Nothing
Exit Sub

Error:
HandleError
Set oFile = Nothing
Set fso = Nothing
End Sub
