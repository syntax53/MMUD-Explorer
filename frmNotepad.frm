VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmNotepad 
   Caption         =   " Notepad"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5355
   Icon            =   "frmNotepad.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   5355
   Begin VB.Timer timWindowMove 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin MSComDlg.CommonDialog oComDag 
      Left            =   3960
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.TextBox txtNotepad 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   5235
   End
   Begin VB.CommandButton cmdDO 
      Cancel          =   -1  'True
      Caption         =   "Cl&ose"
      Height          =   315
      Index           =   3
      Left            =   4020
      TabIndex        =   4
      Top             =   0
      Width           =   1275
   End
   Begin VB.CommandButton cmdDO 
      Caption         =   "&Paste"
      Height          =   315
      Index           =   2
      Left            =   2700
      TabIndex        =   3
      Top             =   0
      Width           =   1275
   End
   Begin VB.CommandButton cmdDO 
      Caption         =   "&Copy"
      Height          =   315
      Index           =   1
      Left            =   1380
      TabIndex        =   2
      Top             =   0
      Width           =   1275
   End
   Begin VB.CommandButton cmdDO 
      Caption         =   "&Save As..."
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   1275
   End
End
Attribute VB_Name = "frmNotepad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Dim tWindowSize As WindowSizeProperties

Public nLastPosTop As Long
Public nLastPosLeft As Long
Public nLastPosMoved As Long
Public nLastPosMonitor As Long

Public nLastTimerTop As Long
Public nLastTimerLeft As Long

Private Sub cmdDO_Click(Index As Integer)
Dim sTemp As String, sFile As String, nPos As Long, nLen As Long
Dim sSectionName As String, x As Integer, sCharFile As String
Dim fso As FileSystemObject, ts As TextStream
On Error GoTo error:

txtNotepad.SetFocus
nPos = txtNotepad.SelStart
nLen = txtNotepad.SelLength

sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ")

sCharFile = ReadINI(sSectionName, "LastCharFile")
If Len(sSessionLastCharFile) > 0 Then sCharFile = sSessionLastCharFile
If Not FileExists(sCharFile) Then
    sCharFile = ""
    sSessionLastCharFile = ""
End If

Select Case Index
    Case 0: 'save
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        If frmMain.bCharLoaded Then
            sFile = sCharFile
            If Not FileExists(sFile) Then
                sFile = ""
                sSessionLastCharFile = ""
            Else
                sSectionName = "PlayerInfo"
            End If
        End If
        
        oComDag.Filter = "Text Files (*.txt)|*.txt"
        oComDag.DialogTitle = "Save Notepad"
        sTemp = ReadINI(sSectionName, "LastNotepadFile", sFile)
        If Len(sTemp) < 4 Then
            oComDag.FileName = "MME-Notepad.txt"
        Else
            oComDag.FileName = sTemp
        End If
        oComDag.InitDir = ReadINI(sSectionName, "LastNotepadDir", sFile)
        
saveagain:
        On Error GoTo out:
        oComDag.ShowSave
        If oComDag.FileName = "" Then GoTo out:
        
        If fso.FileExists(oComDag.FileName) Then
            x = MsgBox("File Exists, Overwrite?", vbQuestion + vbYesNoCancel + vbDefaultButton2)
            If x = vbCancel Then
                Exit Sub
            ElseIf x = vbYes Then
                Call fso.DeleteFile(oComDag.FileName, True)
            Else
                GoTo saveagain:
            End If
        End If
        
        Set ts = fso.OpenTextFile(oComDag.FileName, ForWriting, True)
        ts.Write txtNotepad.Text
        
        Call WriteINI(sSectionName, "LastNotepadFile", oComDag.FileTitle, sFile)
        Call WriteINI(sSectionName, "LastNotepadDir", oComDag.FileName, sFile)
        
    Case 1: 'copy
        If txtNotepad.SelLength = 0 Then
            'Clipboard.SetText txtNotepad.Text
            Call SetClipboardText(txtNotepad.Text)
        Else
            'Clipboard.SetText txtNotepad.SelText
            Call SetClipboardText(txtNotepad.SelText)
        End If
        
        txtNotepad.SetFocus
        txtNotepad.SelStart = nPos
        txtNotepad.SelLength = nLen
        
    Case 2: 'paste
        If Clipboard.GetText = "" Then Exit Sub
        
        If txtNotepad.SelLength = 0 Then
            txtNotepad.Text = Left(txtNotepad.Text, txtNotepad.SelStart) & Clipboard.GetText _
                & Right(txtNotepad.Text, Len(txtNotepad.Text) - txtNotepad.SelStart)
        Else
            txtNotepad.SelText = Clipboard.GetText
        End If
        
        nLen = Len(Clipboard.GetText)
        txtNotepad.SetFocus
        txtNotepad.SelStart = nPos + nLen
        'txtNotepad.SelLength = nLen
        
    Case 3: 'close
        Unload Me
End Select

out:
Set fso = Nothing
Set ts = Nothing
Exit Sub
error:
Call HandleError("cmdDO_Click")
Resume out:

End Sub

Private Sub Form_Load()
On Error GoTo error:
Dim nTemp As Long

Call SetWindowLong(Me.hWnd, GWL_HWNDPARENT, 0)

tWindowSize.twpMinWidth = 5355
tWindowSize.twpMinHeight = 5235
Call SubclassFormMinMaxSize(Me, tWindowSize)

'Me.Height = ReadINI("Settings", "NotepadHeight", , 5000)
'Me.Width = ReadINI("Settings", "NotepadWidth", , 9000)
Call ResizeForm(Me, ReadINI("Settings", "NotepadWidth", , 9000), ReadINI("Settings", "NotepadHeight", , 5500))

nTemp = Val(ReadINI("Settings", "NotepadTOP"))
If nTemp = 0 Then
    If frmMain.WindowState = vbMinimized Then
        nTemp = (Screen.Height - Me.Height) / 2
    Else
        nTemp = frmMain.Top + ((frmMain.Height - Me.Height) / 2)
    End If
End If
Me.Top = nTemp

nTemp = Val(ReadINI("Settings", "NotepadLeft"))
If nTemp = 0 Then
    If frmMain.WindowState = vbMinimized Then
        nTemp = (Screen.Width - Me.Width) / 2
    Else
        nTemp = frmMain.Left + ((frmMain.Width - Me.Width) / 2)
    End If
End If
Me.Left = nTemp

timWindowMove.Enabled = True

If ReadINI("Settings", "NotepadMaxed") = "1" Then Me.WindowState = vbMaximized

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("Form_Load")
Resume out:
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = vbMinimized Then Exit Sub

'txtNotepad.Width = Me.Width - 320
'txtNotepad.Height = Me.Height - TITLEBAR_OFFSET - 950

txtNotepad.Width = Me.ScaleWidth - txtNotepad.Left - 50
txtNotepad.Height = Me.ScaleHeight - txtNotepad.Top - 50

'CheckPosition Me
End Sub

Private Sub Form_Unload(Cancel As Integer)

If Me.WindowState = vbNormal Then
    Call WriteINI("Settings", "NotepadTOP", Me.Top)
    Call WriteINI("Settings", "NotepadLeft", Me.Left)
    Call WriteINI("Settings", "NotepadHeight", Me.Height)
    Call WriteINI("Settings", "NotepadWidth", Me.Width)
End If

If Me.WindowState = vbMaximized Then
    Call WriteINI("Settings", "NotepadMaxed", 1)
Else
    Call WriteINI("Settings", "NotepadMaxed", 0)
End If

End Sub

Private Sub timWindowMove_Timer()
Call MonitorFormTimer(Me)
End Sub
