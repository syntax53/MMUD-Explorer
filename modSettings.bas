Attribute VB_Name = "modSettings"

Global INIReadOnly As Boolean
Global INIFileName As String

Private Ret As String

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Sub CleanSettings(ByVal inputFile As String, ByVal outputFile As String)
'thank you ChatGPT-4o for 90% of this. Prompt:
'provide vb6 code that will read a text file, loop through each line, reading it,
'selectively storing lines in a temporary variable, and then writing the file back from that variable.
'As reading the file, we will look for lines that start and end with brackets ( "[" and "]" ).  This indicates a new section in the file.
'If the section is "[Settings]", that section will always be saved.
'For all other sections, we will read the line of data in between, which will be in the format of: Value = Data
'As we read lines within a section, we are looking for the following VALUE's: DataFile OR LastCharFile.  The DATA of those values should be a file path.  We want to check if that file exists.
'If neither VALUES are found OR *all* file paths point to non-existent files, then we will exclude that section and all VALUE/DATA pairs within it in the final output.
On Error GoTo error:
Dim fso As Object
Dim fileIn As Object
Dim fileOut As Object
Dim sLine As String
Dim sNewSettingsContent As String
Dim sCurrentSectionContent As String
Dim sCurrentSectionTag As String
Dim bKeepSection As Boolean
Dim bHasValidFile As Boolean
Dim sKey As String, sValue As String

' FileSystemObject for file operations
Set fso = CreateObject("Scripting.FileSystemObject")

' Check if input file exists
If Not fso.FileExists(inputFile) Then Exit Sub

' Open input file for reading
Set fileIn = fso.OpenTextFile(inputFile, 1) ' ForReading

' Initialize variables
sNewSettingsContent = ""
sCurrentSectionContent = ""
bKeepSection = False
bHasValidFile = False
sCurrentSectionTag = ""
sOriginalContent = ""

' Read file sLine by sLine
Do Until fileIn.AtEndOfStream
    sLine = Trim(fileIn.ReadLine)
    sOriginalContent = sOriginalContent & sLine & vbCrLf
    
    ' Check if sLine indicates a new section
    If Left(sLine, 1) = "[" And Right(sLine, 1) = "]" Then
        ' Process previous section before moving to the new one
        If bKeepSection Or bHasValidFile Then
            sNewSettingsContent = sNewSettingsContent & sCurrentSectionContent
        End If
        
        ' Reset section-specific variables
        sCurrentSectionContent = sLine & vbCrLf
        sCurrentSectionTag = sLine
        bKeepSection = (sCurrentSectionTag = "[Settings]") ' Always keep [Settings]
        bHasValidFile = False
        
    ElseIf InStr(sLine, "=") > 0 Then
        
        If bHasValidFile = False Then
            ' Process sKey-sValue pairs
            sKey = Trim(Left(sLine, InStr(sLine, "=") - 1))
            sValue = Trim(Mid(sLine, InStr(sLine, "=") + 1))
            
            ' If sKey matches the specified values, check if file exists
            If sKey = "DataFile" Or sKey = "LastCharFile" Then
                If fso.FileExists(sValue) Then
                    bHasValidFile = True
                End If
            End If
        End If
        
        ' Always store section content in case it's needed
        sCurrentSectionContent = sCurrentSectionContent & sLine & vbCrLf
    Else
        ' Store non-sKey-sValue content (blank lines, comments, etc.)
        ' sCurrentSectionContent = sCurrentSectionContent & sLine & vbCrLf
    End If
Loop

' Final check for the last section
If bKeepSection Or bHasValidFile Then
    sNewSettingsContent = sNewSettingsContent & sCurrentSectionContent
End If

' Close input file
fileIn.Close

If Not Trim(sNewSettingsContent) = Trim(sOriginalContent) Then
    ' Write back to output file
    Set fileOut = fso.OpenTextFile(outputFile, 2, True) ' ForWriting, Create if missing
    fileOut.Write sNewSettingsContent
    fileOut.Close
End If

' Cleanup
Set fileIn = Nothing
Set fileOut = Nothing
Set fso = Nothing

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("CleanSettings")
Resume out:
End Sub

Public Function ReadINI(ByVal Section As String, ByVal Key As String, Optional ByVal AlternateINIFile As String, Optional ByVal sDefaultValue As String = "0") As Variant
Dim nTries As Integer
On Error GoTo error:

If AlternateINIFile = "" Then AlternateINIFile = INIFileName

reread:
Ret = Space$(255)
retlen = GetPrivateProfileString(Section, Key, "", Ret, Len(Ret), ByVal AlternateINIFile)
If retlen = 0 Then
    
    If INIReadOnly = True Then Exit Function
    Call WriteINI(Section, Key, sDefaultValue, AlternateINIFile)
    
    Select Case UCase(Section)
        Case "SETTINGS":
            If UCase(Key) = "DATAFILE" Then Call WriteINI(Section, Key, "data-v1.11p.mdb", ByVal AlternateINIFile)
            If Left(UCase(Key), 4) = "LOAD" Then Call WriteINI(Section, Key, 1, ByVal AlternateINIFile)
            If UCase(Key) = "ONLYINGAME" Then Call WriteINI(Section, Key, 1, ByVal AlternateINIFile)
            If Left(UCase(Key), 9) = "INVENSTAT" Then Call WriteINI(Section, Key, 1, ByVal AlternateINIFile)
    End Select
    
    nTries = nTries + 1
    If nTries <= 1 Then GoTo reread:
End If

Ret = Left$(Ret, retlen)
ReadINI = Ret

If Left(UCase(Key), 6) = Left(UCase("Global"), 6) And val(ReadINI) < 0 Then
    If INIReadOnly = True Then Exit Function
    Call WriteINI(Section, Key, 0, AlternateINIFile)
    nTries = nTries + 1
    If nTries <= 2 Then GoTo reread:
End If

Exit Function
error:
HandleError
End Function

Public Sub WriteINI(ByVal Section As String, ByVal Key As String, ByVal Text As String, Optional ByVal AlternateINIFile As String)
On Error GoTo error:

If AlternateINIFile = "" Then AlternateINIFile = INIFileName

If INIReadOnly = True Then Exit Sub
Call WritePrivateProfileString(Section, Key, Text, ByVal AlternateINIFile)

Exit Sub
error:
HandleError
End Sub

Public Function GetSettingsFilePath() As String
    Dim fso As Object
    Dim sLocalPath As String
    Dim sAppData As String
    Dim Wsh As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' --- Try program folder first ---
    If Right$(App.Path, 1) = "\" Then
        sLocalPath = App.Path & "settings.ini"
    Else
        sLocalPath = App.Path & "\settings.ini"
    End If
    
    On Error Resume Next
    ' Try to create a dummy test file
    Dim ts As Object
    Set ts = fso.CreateTextFile(sLocalPath, True)
    If Err.Number = 0 Then
        ' Success, use this path
        ts.Close
        fso.DeleteFile sLocalPath, True  ' cleanup test
        GetSettingsFilePath = sLocalPath
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    
    ' --- Fall back to AppData ---
    Set Wsh = CreateObject("WScript.Shell")
    sAppData = Wsh.SpecialFolders("AppData") & "\MyApp"
    If Not fso.FolderExists(sAppData) Then
        fso.CreateFolder sAppData
    End If
    
    GetSettingsFilePath = sAppData & "\settings.ini"
End Function

Public Function ResolveSettingsPath(ByRef bNewCreated As Boolean) As String
    On Error GoTo fail
    Const ForReading As Long = 1
    Const ForWriting As Long = 2
    Const ForAppending As Long = 8
    
    Dim fso As Object
    Dim Wsh As Object
    Dim localPath As String
    Dim appDataDir As String
    Dim appDataPath As String
    Dim ts As Object
    
    bNewCreated = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Build local path: App.Path\settings.ini
    If Right$(App.Path, 1) = "\" Then
        localPath = App.Path & "settings.ini"
    Else
        localPath = App.Path & "\settings.ini"
    End If
    
    ' --- Step 1: If it exists AND is writable, use it ---
    If fso.FileExists(localPath) Then
        On Error Resume Next
        Set ts = fso.OpenTextFile(localPath, ForAppending, False)
        If Err.Number = 0 Then
            ts.Close
            ResolveSettingsPath = localPath
            On Error GoTo 0
            GoTo done
        End If
        ' not writable ? fall through to step 3
        On Error GoTo 0
    Else
        ' --- Step 2: It does not exist; try to create it locally ---
        On Error Resume Next
        Set ts = fso.OpenTextFile(localPath, ForWriting, True) ' create if missing
        If Err.Number = 0 Then
            ts.Close
            bNewCreated = True
            ResolveSettingsPath = localPath
            On Error GoTo 0
            GoTo done
        End If
        ' creation failed ? fall through to step 3
        On Error GoTo 0
    End If
    
    ' --- Step 3: Fallback to AppData\MMUD Explorer\settings.ini ---
    Set Wsh = CreateObject("WScript.Shell")
    appDataDir = Wsh.SpecialFolders("AppData") & "\MMUD Explorer"
    If Not fso.FolderExists(appDataDir) Then
        fso.CreateFolder appDataDir
    End If
    appDataPath = appDataDir & "\settings.ini"
    
    If Not fso.FileExists(appDataPath) Then
        ' create a new file in AppData
        Set ts = fso.OpenTextFile(appDataPath, ForWriting, True)
        ts.Close
        bNewCreated = True
    End If
    
    ResolveSettingsPath = appDataPath
    GoTo done

fail:
    ' If something unexpected happens, last-resort: try AppData anyway
    On Error Resume Next
    Set Wsh = CreateObject("WScript.Shell")
    appDataDir = Wsh.SpecialFolders("AppData") & "\MMUD Explorer"
    If Not fso Is Nothing Then
        If Not fso.FolderExists(appDataDir) Then fso.CreateFolder appDataDir
    End If
    ResolveSettingsPath = appDataDir & "\settings.ini"
    If Not fso Is Nothing Then
        If Not fso.FileExists(ResolveSettingsPath) Then
            Set ts = fso.OpenTextFile(ResolveSettingsPath, ForWriting, True)
            ts.Close
            bNewCreated = True
        End If
    End If
done:
    Set ts = Nothing
    Set Wsh = Nothing
    Set fso = Nothing
End Function

Public Sub CreateSettings()
On Error GoTo error:
    Dim fso As Object
    Dim ts As Object
    
    ' INIFileName is already set by ResolveSettingsPath
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create only if missing; do NOT delete existing file
    If Not fso.FileExists(INIFileName) Then
        Set ts = fso.OpenTextFile(INIFileName, 2, True) ' ForWriting, create if missing
        ts.Close
    End If
    
    ' Write initial/default values (this should append/replace keys, not wipe the file)
    Call WriteINI("Settings", "DataFile", "data-v1.11p.mdb")
    
    Set ts = Nothing
    Set fso = Nothing
    Exit Sub
error:
    Call HandleError("CreateSettings")
    Set ts = Nothing
    Set fso = Nothing
End Sub

'Public Sub CreateSettings()
'On Error GoTo error:
'Dim fso As FileSystemObject
'Dim sAppPath As String
'
'Set fso = CreateObject("Scripting.FileSystemObject")
'
'If fso.FileExists(INIFileName) Then fso.DeleteFile INIFileName, True
'
'fso.CreateTextFile INIFileName, True
'
'If Right(App.Path, 1) = "\" Then
'    sAppPath = App.Path
'Else
'    sAppPath = App.Path & "\"
'End If
'
'Call WriteINI("Settings", "DataFile", "data-v1.11p.mdb")
'
'Set fso = Nothing
'
'Exit Sub
'error:
'HandleError
'Set fso = Nothing
'End Sub

Public Sub CheckINIReadOnly()
Dim fso As FileSystemObject, nYesNo As Integer, oFile As File
On Error GoTo error:

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

error:
HandleError
Set oFile = Nothing
Set fso = Nothing
End Sub
