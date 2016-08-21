VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{20D5284F-7B23-4F0A-B8B1-6C9D18B64F1C}#1.0#0"; "exlimiter.ocx"
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
   Begin exlimiter.EL EL1 
      Left            =   4560
      Top             =   4440
      _ExtentX        =   1270
      _ExtentY        =   1270
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

Private Sub cmdDO_Click(Index As Integer)
Dim sTemp As String, sFile As String, nPos As Long, nLen As Long
Dim sSectionName As String, x As Integer
Dim fso As FileSystemObject, ts As TextStream
On Error GoTo Error:

txtNotepad.SetFocus
nPos = txtNotepad.SelStart
nLen = txtNotepad.SelLength

Select Case Index
    Case 0: 'save
        Set fso = CreateObject("Scripting.FileSystemObject")
        sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ")

        If frmMain.bCharLoaded Then
            sFile = ReadINI(sSectionName, "LastCharFile")
            If Not FileExists(sFile) Then
                sFile = ""
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
            Clipboard.SetText txtNotepad.Text
        Else
            Clipboard.SetText txtNotepad.SelText
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
Error:
Call HandleError("cmdDO_Click")
Resume out:

End Sub

Private Sub Form_Load()

With EL1
    .CenterOnLoad = False
    .FormInQuestion = Me
    .MinWidth = 365
    .MinHeight = 100
    .EnableLimiter = True
End With

On Error Resume Next

Me.Top = ReadINI("Settings", "NotepadTOP")
Me.Left = ReadINI("Settings", "NotepadLeft")
If Me.Top < 2 Then Me.Top = frmMain.Top
If Me.Left < 2 Then Me.Left = frmMain.Left

Me.Height = ReadINI("Settings", "NotepadHeight")
Me.Width = ReadINI("Settings", "NotepadWidth")

If ReadINI("Settings", "NotepadMaxed") = "1" Then Me.WindowState = vbMaximized

End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = vbMinimized Then Exit Sub

txtNotepad.Width = Me.Width - 240
txtNotepad.Height = Me.Height - TITLEBAR_OFFSET - 825

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
