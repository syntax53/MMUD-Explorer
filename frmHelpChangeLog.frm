VERSION 5.00
Begin VB.Form frmHelpChangeLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ChangeLog"
   ClientHeight    =   6252
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   10032
   Icon            =   "frmHelpChangeLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6252
   ScaleWidth      =   10032
   Begin VB.Timer timWindowMove 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Height          =   6135
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmHelpChangeLog.frx":0CCA
      Top             =   60
      Width           =   9915
   End
End
Attribute VB_Name = "frmHelpChangeLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Public nLastPosTop As Long
Public nLastPosLeft As Long
Public nLastPosMoved As Long
Public nLastPosMonitor As Long

Public nLastTimerTop As Long
Public nLastTimerLeft As Long

Private Sub Form_Load()
On Error Resume Next
Dim x As Integer, y As Integer
x = ConvertScale(10104, vbTwips, vbPixels) 'width
y = ConvertScale(6672, vbTwips, vbPixels) 'height
SubclassFormMinMaxSize Me, x, y, x, y
'SubclassForm Me
If frmMain.WindowState = vbMinimized Then
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
Else
    Me.Left = frmMain.Left + ((frmMain.Width - Me.Width) / 2)
    Me.Top = frmMain.Top + ((frmMain.Height - Me.Height) / 2)
End If
timWindowMove.Enabled = True

End Sub

Private Sub Form_Resize()
'
'    Dim lUseWidth As Long
'    Dim lUseHeight As Long
'
'    Const MINWIDTH As Long = 3000
'    Const MINHEIGHT As Long = 3000
'
'    'Copy the current width and height to our variables
'    lUseWidth = Me.Width
'    lUseHeight = Me.Height
'
'    'Set a minimum limit on the lUseWidth and lUseHeight variables
'    If lUseWidth < MINWIDTH Then lUseWidth = MINWIDTH
'    If lUseHeight < MINHEIGHT Then lUseHeight = MINHEIGHT
'
'    'Set the size of the textbox using the values in lUseWidth and lUseHeight
'    With Text1
'        .Move .Left, .Top, lUseWidth - 125, lUseHeight - TITLEBAR_OFFSET - 125
'    End With
 CheckPosition Me
End Sub

Private Sub timWindowMove_Timer()
Call MonitorFormTimer(Me)
End Sub
