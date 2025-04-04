VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmProgressBar 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timWindowMove 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   3495
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   315
         Left            =   1020
         TabIndex        =   0
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   60
         TabIndex        =   2
         Top             =   300
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Dim nScale As Integer
Dim nScaleCount As Long

Public nLastPosTop As Long
Public nLastPosLeft As Long
Public nLastPosMoved As Long
Public nLastPosMonitor As Long

Public nLastTimerTop As Long
Public nLastTimerLeft As Long

Public objFormOwner As Form

Private Sub cmdCancel_Click()

Select Case lblCaption.Caption
    Case "Searching for Room Name ...", "Searching path steps...", "Calculate mob dmg vs char...", "Calculate mob dmg vs party...", "Applying Filters...", "Removing Filters...":
        objFormOwner.bMapCancelFind = True
        DoEvents
End Select

Me.Hide
End Sub

Private Sub Form_Load()
Dim nLng As Long
On Error Resume Next
SubclassForm Me
If frmMain.WindowState = vbMinimized Then
    nLng = Val(ReadINI("Settings", "Top", , 0))
    If nLng <> 0 Then
        Me.Top = nLng
    Else
        Me.Top = (Screen.Height - Me.Height) / 2
    End If
    
    nLng = Val(ReadINI("Settings", "Left", , 0))
    If nLng <> 0 Then
        Me.Left = nLng
    Else
        Me.Left = (Screen.Width - Me.Width) / 2
    End If
Else
    Me.Left = frmMain.Left + ((frmMain.Width - Me.Width) / 2)
    Me.Top = frmMain.Top + ((frmMain.Height - Me.Height) / 2)
End If

nScale = 0
nScaleCount = 1
ProgressBar.Value = 0
ProgressBar.Min = 0
ProgressBar.Max = 32767

timWindowMove.Enabled = True

End Sub

Public Sub SetRange(ByVal MaxValue As Double, Optional ByVal bReset As Boolean = True)
Dim nNewMax As Integer

nScale = 0

If MaxValue > MaxInt Then
    If MaxValue / 2 < MaxInt Then
        nScale = 2
        nNewMax = MaxValue / 2
    ElseIf MaxValue / 4 < MaxInt Then
        nScale = 4
        nNewMax = MaxValue / 4
    ElseIf MaxValue / 8 < MaxInt Then
        nScale = 8
        nNewMax = MaxValue / 8
    ElseIf MaxValue / 10 < MaxInt Then
        nScale = 10
        nNewMax = MaxValue / 10
    Else
        MaxValue = MaxInt
    End If
Else
    nNewMax = MaxValue
End If

nNewMax = Fix(nNewMax)

nScaleCount = 1
If bReset Then ProgressBar.Value = 0
If ProgressBar.Max < nNewMax Then
    If ProgressBar.Value > nNewMax Then ProgressBar.Value = ProgressBar.Max
End If
ProgressBar.Min = 0
ProgressBar.Max = nNewMax
End Sub

Public Sub IncreaseProgress()
    If nScale > 0 Then
        If nScaleCount = nScale Then
            If ProgressBar.Value + 1 < ProgressBar.Max Then ProgressBar.Value = ProgressBar.Value + 1
            nScaleCount = 1
        Else
            nScaleCount = nScaleCount + 1
        End If
    Else
        If ProgressBar.Value + 1 < ProgressBar.Max Then ProgressBar.Value = ProgressBar.Value + 1
    End If
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub

CheckPosition Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set objFormOwner = Nothing
End Sub

Private Sub timWindowMove_Timer()
Call MonitorFormTimer(Me)
End Sub
