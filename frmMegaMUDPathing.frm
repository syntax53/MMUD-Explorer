VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmMegaMUDPath 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MegaMUD Pathing"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9165
   Icon            =   "frmMegaMUDPathing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9165
   Begin VB.Timer timWindowMove 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5520
      Top             =   60
   End
   Begin VB.TextBox txtPickLock 
      Height          =   315
      Left            =   6540
      MaxLength       =   4
      TabIndex        =   32
      Text            =   "300"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "+"
      Height          =   255
      Index           =   11
      Left            =   4000
      MaskColor       =   &H80000016&
      TabIndex        =   31
      Top             =   420
      Width           =   375
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   3960
      MaskColor       =   &H80000016&
      TabIndex        =   20
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   3240
      MaskColor       =   &H80000016&
      TabIndex        =   21
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "SE"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3960
      MaskColor       =   &H80000016&
      TabIndex        =   22
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3600
      MaskColor       =   &H80000016&
      TabIndex        =   23
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "SW"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   3240
      MaskColor       =   &H80000016&
      TabIndex        =   24
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3960
      MaskColor       =   &H80000016&
      TabIndex        =   25
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   3600
      MaskColor       =   &H80000016&
      TabIndex        =   30
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3240
      MaskColor       =   &H80000016&
      TabIndex        =   26
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "NE"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3960
      MaskColor       =   &H80000016&
      TabIndex        =   27
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3600
      MaskColor       =   &H80000016&
      TabIndex        =   28
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "NW"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3240
      MaskColor       =   &H80000016&
      TabIndex        =   29
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdUndoNote 
      Caption         =   "!"
      Height          =   315
      Left            =   8760
      TabIndex        =   19
      Top             =   360
      Width           =   315
   End
   Begin VB.CommandButton cmdUndoStep 
      Caption         =   "Undo Step"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7380
      TabIndex        =   18
      Top             =   360
      Width           =   1275
   End
   Begin VB.CommandButton cmdResetCurrent 
      Caption         =   "Set to Current Room"
      Height          =   435
      Left            =   7200
      TabIndex        =   13
      Top             =   6240
      Width           =   1875
   End
   Begin VB.TextBox txtCurrentRoom 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   5880
      Width           =   4155
   End
   Begin VB.CommandButton cmdGotoLast 
      Caption         =   "Go to Current Posiiton"
      Height          =   435
      Left            =   4920
      TabIndex        =   12
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CommandButton cmdGotoStart 
      Caption         =   "Go to Starting Room"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton cmdResetStart 
      Caption         =   "Set to Current Room"
      Height          =   375
      Left            =   7200
      TabIndex        =   9
      Top             =   5160
      Width           =   1875
   End
   Begin VB.TextBox txtStartingRoom 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4800
      Width           =   4155
   End
   Begin VB.TextBox txtMapMove 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   4215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   360
      Width           =   4575
   End
   Begin VB.CommandButton cmdMapMoveClear 
      Caption         =   "Cl&ear All"
      Height          =   375
      Left            =   2700
      TabIndex        =   15
      Top             =   5700
      Width           =   1995
   End
   Begin VB.CommandButton cmdMapCopyToClip 
      Caption         =   "Cop&y to Clipboard"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   5700
      Width           =   2355
   End
   Begin VB.CommandButton cmdMapSwitch 
      Caption         =   "&Switch to Manual Edit"
      Height          =   495
      Left            =   2700
      TabIndex        =   4
      Top             =   5100
      Width           =   1995
   End
   Begin VB.CommandButton cmdMapAddMegaCodes 
      Caption         =   "Add Headers and Save Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   6180
      Width           =   4575
   End
   Begin VB.CommandButton cmdMapCommand 
      Caption         =   "Enter a Manual Command in the Current Room"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5100
      Width           =   2355
   End
   Begin MSComctlLib.ListView lvHistory 
      Height          =   3795
      Left            =   4800
      TabIndex        =   16
      Tag             =   "STRETCHALL"
      Top             =   720
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   6694
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog oComDag 
      Left            =   240
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label lblLabelArray 
      AutoSize        =   -1  'True
      Caption         =   "Place your cursor in the black box below and move around the map with your keypad."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   21
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   8790
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Picklock flag if < than:"
      Height          =   195
      Left            =   4860
      TabIndex        =   1
      Top             =   420
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Note: You can double-click teleports from the main window references.  The command will be recorded."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   4500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Current Pathing Position:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4920
      TabIndex        =   10
      Top             =   5640
      Width           =   2130
   End
   Begin VB.Label lblStartingRoom 
      AutoSize        =   -1  'True
      Caption         =   "Starting Room:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4920
      TabIndex        =   6
      Top             =   4560
      Width           =   1275
   End
End
Attribute VB_Name = "frmMegaMUDPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Enum MegaRoomFlags
    STEPF_NONE = &H0
    STEPF_DARK = &H1
    STEPF_REST = &H2
    STEPF_DONTREST = &H4
    STEPF_NOLIGHT = &H8
    STEPF_STASH = &H10
    STEPF_NOATTACK = &H40
    STEPF_RELEARN = &H80
    STEPF_SUSPECT = &H100
    STEPF_DISARM = &H200
    STEPF_CANPICK = &H400
    STEPF_CYCLED = &H8000
    STEPF_MASK = &HFFFF
End Enum

Dim nLocalMapStartMap As Long
Dim nLocalMapStartRoom As Long
Public nLastMapRecorded As Long
Public nLastRoomRecorded As Long

Public nLastPosTop As Long
Public nLastPosLeft As Long
Public nLastPosMoved As Long
Public nLastPosMonitor As Long

Public nLastTimerTop As Long
Public nLastTimerLeft As Long

Private Sub cmdGotoLast_Click()
On Error GoTo error:

Call frmMain.MapStartMapping(nLastMapRecorded, nLastRoomRecorded)

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("cmdGotoLast_Click")
Resume out:
End Sub

Private Sub cmdGotoStart_Click()
On Error GoTo error:

Call frmMain.MapStartMapping(nLocalMapStartMap, nLocalMapStartRoom)

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("cmdGotoStart_Click")
Resume out:
End Sub

Private Sub cmdMapCommand_Click()
On Error GoTo error:
Dim sTemp As String

sTemp = InputBox("Enter Command")
If Len(sTemp) > 0 Then
    If Not Trim(txtMapMove.Text) = "" Then txtMapMove.Text = Trim(txtMapMove.Text) & vbCrLf
    txtMapMove.Text = txtMapMove.Text & Get_MegaMUD_RoomHash("", frmMain.nMapStartMap, frmMain.nMapStartRoom) & Get_MegaMUD_ExitsCode(frmMain.nMapStartMap, frmMain.nMapStartRoom)
    txtMapMove.Text = Replace(txtMapMove.Text, vbCrLf & vbCrLf, vbCrLf)
    txtMapMove.Text = txtMapMove.Text & ":0000:" & sTemp
    txtMapMove.SetFocus
    txtMapMove.SelStart = Len(txtMapMove.Text)
    txtMapMove.SelLength = 0
    Call AddHistory(sTemp, "", frmMain.nMapStartMap, frmMain.nMapStartRoom)
End If

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("cmdMapCommand_Click")
Resume out:
End Sub
Private Sub cmdMapAddMegaCodes_Click()
On Error GoTo error:
Dim sText As String, sNewText As String, x As Long, y As Long, nSteps As Long
Dim sFile As String, fso As FileSystemObject, oTS As TextStream, oFile As File, oFolder As Folder
Dim sLine As String, sArr() As String, oSubFolder As Folder, oSubFolder2 As Folder
Dim sFileHeader(3) As String, sNeededItem As String, sPathSteps() As String
Dim sStartingRoomChecksum As String, sEndingRoomChecksum As String
Dim sStartRoomName As String, sEndRoomName As String
Dim sStartRoomCode As String, sEndRoomCode As String

txtMapMove.Text = Replace(txtMapMove.Text, vbCrLf & vbCrLf, vbCrLf)

sText = txtMapMove.Text
If Len(sText) < 1 Then
    MsgBox "No text to copy.", vbInformation
    Exit Sub
End If

'sText = "[LOOP NAME][AUTHOR]" & vbCrLf _
    & "[CODE:FROM ROOM GROUP:FROM ROOM NAME]" & vbCrLf _
    & "[CODE:TO ROOM GROUP:TO ROOM NAME]" & vbCrLf _
    & "START ROOM CHECKSUM:END ROOM CHECKSUM:" & nSteps & ":-1:0:Needed Items:Path fails:when finished"

sPathSteps() = Split(txtMapMove.Text, vbCrLf)
If Not IsDimmed(sPathSteps) Then GoTo out:

nSteps = UBound(sPathSteps()) + 1

If Len(sPathSteps(0)) > 2 And Left(sPathSteps(0), 1) = ":" And Right(sPathSteps(0), 1) = ":" Then
    sNeededItem = Replace(sPathSteps(0), ":", "")
    sPathSteps(0) = ""
    nSteps = nSteps - 1
End If

sStartRoomName = GetRoomName("", nLocalMapStartMap, nLocalMapStartRoom, True)
sStartingRoomChecksum = Get_MegaMUD_RoomHash(sStartRoomName) & Get_MegaMUD_ExitsCode(nLocalMapStartMap, nLocalMapStartRoom)
sStartRoomCode = Get_MegaMUD_RoomHash(sStartRoomName) & "9"

sEndRoomName = GetRoomName("", nLastMapRecorded, nLastRoomRecorded, True)
sEndingRoomChecksum = Get_MegaMUD_RoomHash(sEndRoomName) & Get_MegaMUD_ExitsCode(nLastMapRecorded, nLastRoomRecorded)
sEndRoomCode = Get_MegaMUD_RoomHash(sEndRoomName) & "9"

sFileHeader(0) = "[][]"
'sFileHeader(1) = "[FFFF:Custom Paths:" & sStartRoomName
'sFileHeader(2) = "[FFFF:Custom Paths:" & sEndRoomName
sFileHeader(3) = sStartingRoomChecksum & ":" & sEndingRoomChecksum & ":" & nSteps & ":-1:0:" & sNeededItem & "::"

MsgBox "On the next screen you will be asked to select the MegaMUD rooms.md database file. " _
    & "Choose the MAIN rooms.md file (the one under the ""Default"" folder).  Other rooms.md files around it will also be examined." _
    & vbCrLf & "This will allow us to lookup matching starting and ending rooms and populate those fields.", vbInformation
    
Set fso = CreateObject("Scripting.FileSystemObject")

oComDag.Filter = "Rooms.md file (Rooms.md)|Rooms.md"
oComDag.DialogTitle = "Select MegaMUD Rooms.md File"
oComDag.FileName = "Rooms.md"

If fso.FolderExists(ReadINI("Settings", "Last_MegaMUD_DBFolder", , "C:\Program Files (x86)\Megamud\Default")) Then
    oComDag.InitDir = ReadINI("Settings", "Last_MegaMUD_DBFolder")
ElseIf fso.FileExists("C:\Program Files (x86)\Megamud\Default\Rooms.md") Then
    oComDag.InitDir = "C:\Program Files (x86)\Megamud\Default"
ElseIf fso.FileExists("C:\Megamud\Default\Rooms.md") Then
    oComDag.InitDir = "C:\Megamud\Default"
ElseIf fso.FileExists(Environ("USERPROFILE") & "\AppData\Local\VirtualStore\Program Files (x86)\Megamud\Default\Rooms.md") Then
    oComDag.InitDir = Environ("USERPROFILE") & "\AppData\Local\VirtualStore\Program Files (x86)\Megamud\Default"
Else
    oComDag.InitDir = App.Path
End If

On Error GoTo canceled:
oComDag.ShowOpen
If oComDag.FileName = "" Then GoTo canceled:

continue1:
On Error GoTo error:

sFile = oComDag.FileName
If Not UCase(Right(sFile, 3)) = ".MD" Then sFile = sFile & ".MD"

If Not fso.FileExists(sFile) Then
    x = MsgBox("File not found or file open canceled, continue anyway?", vbYesNo + vbQuestion)
    If Not x = vbYes Then GoTo out:
    GoTo skip_room_lookup:
End If

Set oFile = fso.GetFile(sFile)
Call WriteINI("Settings", "Last_MegaMUD_DBFolder", oFile.ParentFolder)
Set oFile = Nothing

Set oTS = fso.OpenTextFile(sFile, ForReading)
Call ScanRoomsDatabase(oTS, sFileHeader(), sStartRoomCode, sEndRoomCode, sStartingRoomChecksum, sEndingRoomChecksum)
oTS.Close
Set oTS = Nothing

If sFileHeader(1) = "" Or sFileHeader(2) = "" Then
    Set oFile = fso.GetFile(sFile)
    Set oFolder = oFile.ParentFolder
    Set oFolder = oFolder.ParentFolder
    If oFolder.SubFolders.Count > 0 Then
        For Each oSubFolder In oFolder.SubFolders
            If fso.FileExists(oSubFolder.Path & "\rooms.md") Then
                Set oTS = fso.OpenTextFile(oSubFolder.Path & "\rooms.md", ForReading)
                Call ScanRoomsDatabase(oTS, sFileHeader(), sStartRoomCode, sEndRoomCode, sStartingRoomChecksum, sEndingRoomChecksum)
                oTS.Close
                Set oTS = Nothing
                If Not sFileHeader(1) = "" And Not sFileHeader(2) = "" Then GoTo skip_room_lookup
            End If
            
            If oSubFolder.SubFolders.Count > 0 Then
                
                If oSubFolder.SubFolders.Count > 0 Then
                    For Each oSubFolder2 In oSubFolder.SubFolders
                        If fso.FileExists(oSubFolder2.Path & "\rooms.md") Then
                            Set oTS = fso.OpenTextFile(oSubFolder2.Path & "\rooms.md", ForReading)
                            Call ScanRoomsDatabase(oTS, sFileHeader(), sStartRoomCode, sEndRoomCode, sStartingRoomChecksum, sEndingRoomChecksum)
                            oTS.Close
                            Set oTS = Nothing
                            If Not sFileHeader(1) = "" And Not sFileHeader(2) = "" Then GoTo skip_room_lookup
                        End If
                    Next oSubFolder2
                End If
                
            End If
        Next oSubFolder
    End If
End If

skip_room_lookup:
If Not sFileHeader(1) = "" And Not sFileHeader(2) = "" Then
    MsgBox "Starting and ending rooms matched to existing MegaMUD rooms!" & vbCrLf _
        & "Start: " & sFileHeader(1) & vbCrLf _
        & "End: " & sFileHeader(2), vbInformation
End If

If sFileHeader(1) = "" Then
   '[FFFF:Custom Paths:" & sStartRoomName
   sText = InputBox("START room code / group not found for " & vbCrLf & sStartRoomName & vbCrLf & vbCrLf & "Enter 4 character megamud room code to use (must be unique if new to avoid conflict).", , "FFFF")
   If sText = "" Then GoTo out:
   sFileHeader(1) = "[" & sText & ":"
   
   sText = InputBox("START room code / group not found for " & vbCrLf & sStartRoomName & vbCrLf & vbCrLf & "Enter group name to associate with the room.", , "Custom Rooms")
   If sText = "" Then GoTo out:
   sFileHeader(1) = sFileHeader(1) & sText & ":" & sStartRoomName & "]"
End If

If sFileHeader(2) = "" And Not sStartingRoomChecksum = sEndingRoomChecksum Then
   '[FFFF:Custom Paths:" & sEndRoomName
   sText = InputBox("END room code / group not found for " & vbCrLf & sEndRoomName & vbCrLf & vbCrLf & "Enter 4 character megamud room code to use (must be unique if new to avoid conflict).", , "FFFF")
   If sText = "" Then GoTo out:
   sFileHeader(2) = "[" & sText & ":"
   
   sText = InputBox("END room code / group not found for " & vbCrLf & sEndRoomName & vbCrLf & vbCrLf & "Enter group name to associate with the room.", , "Custom Rooms")
   If sText = "" Then GoTo out:
   sFileHeader(2) = sFileHeader(2) & sText & ":" & sEndRoomName & "]"
End If

If sStartingRoomChecksum = sEndingRoomChecksum Then
    sText = InputBox("This appears to be a loop.  Enter the name of the loop:")
    If sText = "" Then GoTo out:
    sFileHeader(0) = "[" & sText & "]"
    sFileHeader(2) = ""
Else
    sFileHeader(0) = "[]"
End If

sText = InputBox("Enter Author", , ReadINI("Settings", "MegaMUD_Path_Author", , "Custom"))
If sText = "" Then GoTo out:
Call WriteINI("Settings", "MegaMUD_Path_Author", sText)
sFileHeader(0) = sFileHeader(0) & "[" & sText & "]"

sText = ""
For x = 0 To 3
    If Not Trim(sFileHeader(x)) = "" Then
        sText = sText & sFileHeader(x) & vbCrLf
    End If
Next x
For x = 0 To UBound(sPathSteps())
    If Not Trim(sPathSteps(x)) = "" Then
        sText = sText & sPathSteps(x) & vbCrLf
    End If
Next x

MsgBox "We are ready to save the path. It is recommended to save the file *outside* of your megamud install " _
    & "(such as your desktop) and then utilize the Add path feature from the" & vbCrLf _
    & "Options -> Game Data -> Paths -> ""Add..."" button. " & vbCrLf _
    & "Any new rooms included in the path will automatically be added by megamud.", vbInformation

oComDag.Filter = "MegaMUD .MP (*.mp)|*.mp"
oComDag.DialogTitle = "Save MegaMUD Path"
If sStartingRoomChecksum = sEndingRoomChecksum Then
    oComDag.FileName = sStartRoomCode & "LOOP.mp"
Else
    oComDag.FileName = sStartRoomCode & sEndRoomCode & ".mp"
End If
oComDag.InitDir = ReadINI("Settings", "LastMegaPathDir", , Environ("USERPROFILE") & "\Desktop")

saveagain:
On Error GoTo out:
oComDag.ShowSave
If oComDag.FileName = "" Then GoTo out:

If fso.FileExists(oComDag.FileName) Then
    x = MsgBox("File Exists, Overwrite?", vbQuestion + vbYesNoCancel + vbDefaultButton2)
    If x = vbCancel Then
        GoTo out:
    ElseIf x = vbYes Then
        Call fso.DeleteFile(oComDag.FileName, True)
    Else
        GoTo saveagain:
    End If
End If

Set oTS = fso.OpenTextFile(oComDag.FileName, ForWriting, True)
oTS.Write sText
oTS.Close
Set oTS = Nothing

Set oFile = fso.GetFile(oComDag.FileName)
Call WriteINI("Settings", "LastMegaPathDir", oFile.ParentFolder)

out:
On Error Resume Next
Set fso = Nothing
Set oTS = Nothing
Set oFile = Nothing
Exit Sub

canceled:
Resume continue1:

Exit Sub
error:
Call HandleError("cmdMapAddMegaCodes_Click")
Resume out:
End Sub

Private Sub ScanRoomsDatabase(ByRef oTS, ByRef sFileHeader() As String, ByRef sStartRoomCode As String, ByRef sEndRoomCode As String, _
    ByRef sStartingRoomChecksum As String, ByRef sEndingRoomChecksum As String)
Dim sLine As String, sArr() As String

Do While oTS.AtEndOfStream = False
    '0        1        2 3 4 5    6            7
    'CAB00180:00004040:0:0:0:AALY:Ancient Ruin:Ancient Ruin Dark Alley
    sLine = oTS.ReadLine
    sArr() = Split(sLine, ":")
    If UBound(sArr()) >= 7 Then
        If sArr(0) = sStartingRoomChecksum And sFileHeader(1) = "" Then
            '[CODE:FROM ROOM GROUP:FROM ROOM NAME]
            sFileHeader(1) = "[" & sArr(5) & ":" & sArr(6) & ":" & sArr(7) & "]"
            sStartRoomCode = sArr(5)
        End If
        If sArr(0) = sEndingRoomChecksum And sFileHeader(2) = "" Then
            '[CODE:TO ROOM GROUP:TO ROOM NAME]
            sFileHeader(2) = "[" & sArr(5) & ":" & sArr(6) & ":" & sArr(7) & "]"
            sEndRoomCode = sArr(5)
        End If
    End If
Loop

End Sub

Private Sub cmdMapSwitch_Click()
If cmdMapSwitch.Tag = "1" Then
    cmdMapSwitch.Tag = ""
    cmdMapSwitch.Caption = "&Switch to Manual Edit"
Else
    cmdMapSwitch.Tag = "1"
    cmdMapSwitch.Caption = "&Switch to Map Move"
End If
txtMapMove.SetFocus
End Sub
Private Sub cmdMapCopyToClip_Click()
On Error GoTo error:
Dim sText As String

sText = txtMapMove.Text
If Len(sText) < 1 Then
    MsgBox "No text to copy.", vbInformation
    Exit Sub
End If

'Clipboard.clear
'Clipboard.SetText sText
Call SetClipboardText(sText)

MsgBox "Copied.", vbInformation

Exit Sub

error:
Call HandleError("cmdMapCopyToClip_Click")
End Sub
Private Sub cmdMapMoveClear_Click()
Dim nYesNo As Integer
nYesNo = MsgBox("Are you sure you want to clear all?", vbYesNo + vbDefaultButton2 + vbQuestion)
If nYesNo = vbYes Then
    txtMapMove.Text = ""
    nLocalMapStartMap = 0
    nLocalMapStartRoom = 0
    nLastMapRecorded = 0
    nLastRoomRecorded = 0
    txtCurrentRoom.Text = ""
    txtStartingRoom.Text = ""
    lvHistory.ListItems.clear
End If
End Sub

Private Sub cmdMove_Click(Index As Integer)
Dim x As Integer
On Error GoTo error:

Select Case Index
    Case 0: Call txtMapMove_KeyPress(56) 'n
    Case 1: Call txtMapMove_KeyPress(50) 's
    Case 2: Call txtMapMove_KeyPress(54) 'e
    Case 3: Call txtMapMove_KeyPress(52) 'w
    Case 4: Call txtMapMove_KeyPress(57) 'ne
    Case 5: Call txtMapMove_KeyPress(55) 'nw
    Case 6: Call txtMapMove_KeyPress(51) 'se
    Case 7: Call txtMapMove_KeyPress(49) 'sw
    Case 8: Call txtMapMove_KeyPress(48) 'u
    Case 9: Call txtMapMove_KeyPress(46) 'd
End Select

If Index = 10 Then
    For x = 0 To 10
        cmdMove(x).Visible = False
    Next x
    cmdMove(11).Visible = True
ElseIf Index = 11 Then
    cmdMove(11).Visible = False
    For x = 0 To 10
        cmdMove(x).Visible = True
    Next x
End If

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("cmdMove_Click")
Resume out:
End Sub

Private Sub cmdResetCurrent_Click()
On Error GoTo error:
Dim x As Integer

If Len(Trim(txtMapMove.Text)) > 0 Then
    x = MsgBox("Are you sure?", vbQuestion + vbYesNo)
    If Not x = vbYes Then Exit Sub
End If

Call SetCurrentPosition(frmMain.nMapStartMap, frmMain.nMapStartRoom)

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("cmdResetCurrent_Click")
Resume out:
End Sub

Private Sub cmdResetStart_Click()
On Error GoTo error:
Dim x As Integer

If Len(Trim(txtMapMove.Text)) > 0 Then
    x = MsgBox("Are you sure?  If you have steps in the path the preceed this room then when you add the megamud codes later the wrong starting room codes will be entered.", vbQuestion + vbYesNo)
    If Not x = vbYes Then Exit Sub
End If
Call ResetStartingRoom

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("cmdResetStart_Click")
Resume out:
End Sub

Private Sub cmdUndoNote_Click()
MsgBox "Note: This simply removes the last line in the box and sets the current position to the previous room in the history.  If you've made manual edits and start undo'ing past those edits then things will start to get messed up.", vbInformation
End Sub

Private Sub cmdUndoStep_Click()
Dim sArr() As String, sText As String, x As Integer, nRoom As RoomExitType
On Error GoTo error:

If Right(txtMapMove.Text, 2) = vbCrLf Then txtMapMove.Text = Left(txtMapMove.Text, Len(txtMapMove.Text) - 2)

If Not Trim(txtMapMove.Text) = "" Then
    txtMapMove.Text = Replace(Trim(txtMapMove.Text), vbCrLf & vbCrLf, vbCrLf)
    sArr() = Split(txtMapMove.Text, vbCrLf)
    If IsDimmed(sArr()) Then
        For x = 0 To UBound(sArr()) - 1
            sText = sText & sArr(x) & vbCrLf
        Next x
        If Len(sText) > 1 Then sText = Left(sText, Len(sText) - 2)
        txtMapMove.Text = sText
        
        txtMapMove.SetFocus
        txtMapMove.SelStart = Len(txtMapMove.Text)
        txtMapMove.SelLength = 0
    End If
End If

If lvHistory.ListItems.Count = 0 Then Exit Sub
nRoom = ExtractMapRoom(lvHistory.ListItems(lvHistory.ListItems.Count).Tag)
If nRoom.Map > 0 And nRoom.Room > 0 Then
    Call frmMain.MapStartMapping(nRoom.Map, nRoom.Room)
    Call SetCurrentPosition(nRoom.Map, nRoom.Room)
End If

lvHistory.ListItems.Remove lvHistory.ListItems.Count
If lvHistory.ListItems.Count = 0 Then Exit Sub
lvHistory.ListItems(lvHistory.ListItems.Count).Selected = True
lvHistory.ListItems(lvHistory.ListItems.Count).EnsureVisible

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("cmdUndoStep_Click")
Resume out:
End Sub

Private Sub Form_Load()
On Error GoTo error:
Dim nTemp As Long
'SubclassForm Me
Call cmdMove_Click(10)
lvHistory.ColumnHeaders.Add , , "Room (dbl-click goto)", 3700

txtPickLock.Text = ReadINI("Settings", "MegaPathPicklocks", , 300)
If Val(txtPickLock.Text) < 1 Then txtPickLock.Text = 1

If frmMain.nMapStartMap > 0 And frmMain.nMapStartRoom > 0 Then
    Call ResetStartingRoom
    Call SetCurrentPosition(frmMain.nMapStartMap, frmMain.nMapStartRoom)
End If


nTemp = Val(ReadINI("Settings", "MegaPathTop"))
If nTemp = 0 Then
    If frmMain.WindowState = vbMinimized Then
        nTemp = (Screen.Height - Me.Height) / 2
    Else
        nTemp = frmMain.Top + ((frmMain.Height - Me.Height) / 2)
    End If
End If
Me.Top = nTemp

nTemp = Val(ReadINI("Settings", "MegaPathLeft"))
If nTemp = 0 Then
    If frmMain.WindowState = vbMinimized Then
        nTemp = (Screen.Width - Me.Width) / 2
    Else
        nTemp = frmMain.Left + ((frmMain.Width - Me.Width) / 2)
    End If
End If
Me.Left = nTemp


timWindowMove.Enabled = True

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("Form_Load")
Resume out:
End Sub

Private Sub ResetStartingRoom()
On Error GoTo error:

nLocalMapStartMap = frmMain.nMapStartMap
nLocalMapStartRoom = frmMain.nMapStartRoom

txtStartingRoom.Text = nLocalMapStartMap & "/" & nLocalMapStartRoom & ": " & GetRoomName("", nLocalMapStartMap, nLocalMapStartRoom, True)

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("ResetStartingRoom")
Resume out:
End Sub

Private Sub Form_Resize()
'CheckPosition Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Call WriteINI("Settings", "MegaPathPicklocks", txtPickLock.Text)
Call WriteINI("Settings", "MegaPathTop", Me.Top)
Call WriteINI("Settings", "MegaPathLeft", Me.Left)

frmMain.txtMapMove.Enabled = True
frmMain.txtMapMove.Text = ""
End Sub

Private Sub lvHistory_DblClick()
On Error GoTo error:
Dim nRoom As RoomExitType

If lvHistory.ListItems.Count = 0 Then Exit Sub
If lvHistory.SelectedItem Is Nothing Then Exit Sub
If lvHistory.SelectedItem.Tag = "" Then Exit Sub
nRoom = ExtractMapRoom(lvHistory.SelectedItem.Tag)
If nRoom.Map > 0 And nRoom.Room > 0 Then
    Call frmMain.MapStartMapping(nRoom.Map, nRoom.Room)
End If


out:
On Error Resume Next
Exit Sub
error:
Call HandleError("lvHistory_DblClick")
Resume out:
End Sub

Private Sub timWindowMove_Timer()
Call MonitorFormTimer(Me)
End Sub

Private Sub txtMapMove_KeyPress(KeyAscii As Integer)
Dim sLook As String, RoomExit As RoomExitType, x As Integer
Dim nExitType As Integer, nRecNum As Long, sRoomName As String
Dim nTest As Integer, sActions(9) As String, sTemp As String
Dim sCurrentRoomMegaMudCode As String, nRoomFlags As Long, sRoomFlags As String

On Error GoTo error:

If (nLocalMapStartMap = 0 Or nLocalMapStartRoom = 0 Or Trim(txtMapMove.Text) = "") And frmMain.nMapStartMap > 0 And frmMain.nMapStartRoom > 0 Then
    Call ResetStartingRoom
End If

If cmdMapSwitch.Tag = "1" Then Exit Sub

If frmMain.bMapStillMapping Then GoTo out:

Select Case KeyAscii
    Case 8, 26: Exit Sub
End Select

If Len(txtMapMove.Text) > 30000 Then
    nRecNum = MsgBox("Maximum Length Reached, Clear Direction Box?", vbQuestion + vbYesNo + vbDefaultButton2)
    If nRecNum = vbYes Then
        txtMapMove.Text = ""
    Else
        GoTo out:
    End If
End If

tabRooms.Index = "idxRooms"
tabRooms.Seek "=", frmMain.nMapStartMap, frmMain.nMapStartRoom
If tabRooms.NoMatch Then GoTo out:

sRoomName = tabRooms.Fields("Name")
sCurrentRoomMegaMudCode = Get_MegaMUD_RoomHash(sRoomName) & Get_MegaMUD_ExitsCode(frmMain.nMapStartMap, frmMain.nMapStartRoom)

If nLastMapRecorded > 0 And nLastRoomRecorded > 0 Then
    If nLocalMapStartMap = frmMain.nMapStartMap And nLocalMapStartRoom = frmMain.nMapStartRoom And txtMapMove.Text = "" Then
        Call SetCurrentPosition(frmMain.nMapStartMap, frmMain.nMapStartRoom)
    Else
        If nLastMapRecorded <> frmMain.nMapStartMap Or nLastRoomRecorded <> frmMain.nMapStartRoom Then
            nTest = MsgBox("Warning, current room (" & frmMain.nMapStartMap & "/" & frmMain.nMapStartRoom & ") does not match the current " _
                & "pathing position (" & nLastMapRecorded & "/" & nLastRoomRecorded & ").  Continue Anyway?" & vbCrLf & vbCrLf _
                & "(Use the ""Go to Current Position"" button.)", vbExclamation + vbYesNo + vbDefaultButton2)
            If Not nTest = vbYes Then GoTo out:
        End If
    End If
End If

Select Case KeyAscii
    Case 46: 'd
        sLook = "D"
    Case 48: 'u
        sLook = "U"
    Case 49: 'sw
        sLook = "SW"
    Case 50: 's
        sLook = "S"
    Case 51: 'se
        sLook = "SE"
    Case 52: 'w
        sLook = "W"
    'Case 53:
    Case 54: 'e
        sLook = "E"
    Case 55: 'nw
        sLook = "NW"
    Case 56: 'n
        sLook = "N"
    Case 57: 'ne
        sLook = "NE"
    Case Else: GoTo out:
End Select

If Left(tabRooms.Fields(sLook), 6) = "Action" Then
    GoTo out:
ElseIf Not Val(tabRooms.Fields(sLook)) = 0 Then
    RoomExit = ExtractMapRoom(tabRooms.Fields(sLook))
    
    tabRooms.Index = "idxRooms"
    tabRooms.Seek "=", RoomExit.Map, RoomExit.Room
    If tabRooms.NoMatch Then
        MsgBox "Error going in that direction."
        GoTo out:
    End If
Else
    GoTo out:
End If

If Len(RoomExit.ExitType) > 2 Then
    Select Case Left(RoomExit.ExitType, 5)
        Case "(Key:": nExitType = 2
        Case "(Item": nExitType = 3
        Case "(Toll": nExitType = 4
        Case "(Hidd": nExitType = 6
        Case "(Door": nExitType = 7
        Case "(Trap": nExitType = 9
        Case "(Text": nExitType = 10
        Case "(Gate": nExitType = 11
        Case "Actio": nExitType = 12
        Case "(Clas": nExitType = 13
        Case "(Race": nExitType = 14
        Case "(Leve": nExitType = 15
        Case "(Time": nExitType = 16
        Case "(Tick": nExitType = 17
        Case "(Max ": nExitType = 18
        Case "(Bloc": nExitType = 19
        Case "(Alig": nExitType = 20
        Case "(Dela": nExitType = 21
        Case "(Cast": nExitType = 22
        Case "(Abil": nExitType = 23
        Case "(Spel": nExitType = 24
    End Select
End If
If Not RoomExit.Map = frmMain.nMapStartMap Then nExitType = 8 'map change

Select Case nExitType
    Case 8: 'map change
    Case 12: 'action
        GoTo out:
    Case 3: 'item
        nRecNum = ExtractNumbersFromString(RoomExit.ExitType)
        If nRecNum > 0 Then
            'sLook = sLook & " -- (Requires " & GetItemName(nRecNum, bHideRecordNumbers) & ")"
            sTemp = GetItemName(nRecNum, bHideRecordNumbers)
            If InStr(1, txtMapMove.Text, ":" & sTemp & ":", vbTextCompare) = 0 Then
                If Left(txtMapMove.Text, 1) = ":" Then
                    MsgBox "Warning: We have already listed a required item for this path (see first line in box) and this step requires another item (" & sTemp & ")." _
                        & " MegaMUD does not support more than 1 required item per path so this item will not be added to the path requirements.", vbExclamation
                    txtMapMove.SetFocus
                Else
                    txtMapMove.Text = ":" & sTemp & ":" & vbCrLf & txtMapMove.Text
                End If
            End If
        End If
    Case 2: 'key
        nRecNum = ExtractNumbersFromString(RoomExit.ExitType)
        If nRecNum > 0 Then
            sLook = sLook & "[use " & GetItemName(nRecNum, bHideRecordNumbers) & " " & sLook & "]"
        Else
            sLook = sLook '& ": open door " & sLook
        End If
        If RoomExit.ExitType Like "*[[]or #* picklocks*" Then
            nTest = InStr(1, RoomExit.ExitType, "[or ", vbTextCompare)
            If nTest > 0 Then
                nTest = ExtractNumbersFromString(Mid(RoomExit.ExitType, nTest + 3))
                If nTest > 0 And nTest < Val(txtPickLock.Text) Then
                    nRoomFlags = nRoomFlags + MegaRoomFlags.STEPF_CANPICK
                End If
            End If
        End If
    Case 6:
        If InStr(1, LCase(RoomExit.ExitType), "action") > 0 Then
            nTest = ExtractValueFromString(RoomExit.ExitType, "needs ")
            If nTest > 0 Then
                For x = 1 To nTest
                    sActions(x) = InputBox("Enter action # " & x & vbCrLf & vbCrLf _
                        & "If the action is not entered in this room, type a zero (0)." & vbCrLf _
                        & "To cancel the move and look up the command press cancel.")
                    If sActions(x) = "" Then GoTo out:
                    If sActions(x) = "0" Then
                        sActions(x) = ""
                        GoTo cont
                    End If
                Next x
cont:
            Else
                sLook = sLook & " -- " & RoomExit.ExitType
            End If
            
        ElseIf InStr(1, LCase(RoomExit.ExitType), "passable") > 0 Then
            'nothing
        Else
            sLook = sLook & "[search " & sLook & "]"
        End If
    Case 9:
        nRoomFlags = nRoomFlags + MegaRoomFlags.STEPF_DISARM
    Case 10:
        sLook = ExtractTextCommand(RoomExit.ExitType)
    Case 4, 13, 14, 15, 20: '
       ' sLook = sLook & " -- " & RoomExit.ExitType
End Select

If Not Trim(txtMapMove.Text) = "" Then txtMapMove.Text = Trim(txtMapMove.Text) & vbCrLf
    
txtMapMove.Text = txtMapMove.Text & sCurrentRoomMegaMudCode

txtMapMove.Text = Replace(txtMapMove.Text, vbCrLf & vbCrLf, vbCrLf)

sRoomFlags = Hex(nRoomFlags)
If Len(sRoomFlags) < 4 Then sRoomFlags = String(4 - Len(sRoomFlags), "0") & sRoomFlags

txtMapMove.Text = txtMapMove.Text & ":" & sRoomFlags & ":" & sLook
If Len(sActions(1)) > 0 Then
    txtMapMove.Text = txtMapMove.Text & "["
    For x = 1 To 9
        If Len(sActions(x)) > 0 Then
            If x > 1 Then txtMapMove.Text = txtMapMove.Text & ","
            txtMapMove.Text = txtMapMove.Text & sActions(x)
        End If
    Next x
    txtMapMove.Text = txtMapMove.Text & "]"
End If

If frmMain.WindowState = vbMinimized Or frmMain.framNav(10).Visible = False Then
    If frmMain.WindowState = vbMinimized Then frmMain.WindowState = vbNormal
    frmMain.bDontSetMainFocus = True
    Call frmMain.cmdNav_Click(10)
    frmMain.bDontSetMainFocus = False
    txtMapMove.SetFocus
End If

txtMapMove.SelStart = Len(txtMapMove.Text)
txtMapMove.SelLength = 0
Call AddHistory(sLook, sRoomName, frmMain.nMapStartMap, frmMain.nMapStartRoom)
Call SetCurrentPosition(RoomExit.Map, RoomExit.Room)
Call frmMain.MapStartMapping(RoomExit.Map, RoomExit.Room)

out:
KeyAscii = 0

Exit Sub
error:
Call HandleError("txtMapMove_KeyPress")

End Sub

Public Sub AddHistory(ByVal sCommand As String, ByVal sRoomName As String, ByVal nMap As Long, ByVal nRoom As Long)
On Error GoTo error:

If sRoomName = "" Then sRoomName = GetRoomName("", nMap, nRoom, True)

lvHistory.ListItems.Add , , sCommand & " - " & sRoomName & " (" & nMap & "/" & nRoom & ")"
lvHistory.ListItems(lvHistory.ListItems.Count).Tag = nMap & "/" & nRoom
lvHistory.ListItems(lvHistory.ListItems.Count).Selected = True
lvHistory.ListItems(lvHistory.ListItems.Count).EnsureVisible

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("AddHistory")
Resume out:
End Sub

Public Sub SetCurrentPosition(ByVal nMap As Long, ByVal nRoom As Long)
On Error GoTo error:

nLastMapRecorded = nMap
nLastRoomRecorded = nRoom
txtCurrentRoom.Text = nMap & "/" & nRoom & ": " & GetRoomName("", nMap, nRoom, True)

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("SetCurrentPosition")
Resume out:
End Sub

Private Sub txtPickLock_GotFocus()
Call SelectAll(txtPickLock)
End Sub
