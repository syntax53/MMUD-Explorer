VERSION 5.00
Begin VB.Form frmLoad 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Loading ..."
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   Moveable        =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4410
      Left            =   -45
      ScaleHeight     =   4410
      ScaleWidth      =   3960
      TabIndex        =   0
      Top             =   -60
      Width           =   3960
      Begin VB.Label lblCaption 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   180
         MousePointer    =   11  'Hourglass
         TabIndex        =   1
         Top             =   4020
         Width           =   2355
      End
      Begin VB.Image Image1 
         Height          =   4350
         Left            =   40
         Picture         =   "frmLoad.frx":0000
         Stretch         =   -1  'True
         Top             =   60
         Width           =   3900
      End
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Sub Form_Load()
Dim X As Long, Y As Long
Dim mi As MONITORINFO, hMonitor As Long, wr As RECT
Dim workWidth As Long, workHeight As Long
Dim hDC As Long, dpiX As Long, dpiY As Long
Dim localTwipsPerPixelX As Double, localTwipsPerPixelY As Double
Dim winWidth As Long, winHeight As Long
Dim newLeft As Long, newTop As Long
'On Error Resume Next
On Error GoTo error:
SubclassForm Me
X = Val(ReadINI("Settings", "Top", , 0))
Y = Val(ReadINI("Settings", "Left", , 0))

If X <> 0 And Y <> 0 Then
    Me.Top = X + (frmMain.Height / 2)
    Me.Left = Y + (frmMain.Width / 2)
    DoEvents
'    mi.cbSize = Len(mi)
'    hMonitor = MonitorFromWindow(Me.hwnd, MONITOR_DEFAULTTONEAREST)
'    GetMonitorInfo hMonitor, mi
'
'    workWidth = mi.rcWork.Right - mi.rcWork.Left
'    workHeight = mi.rcWork.Bottom - mi.rcWork.Top
'
'    hdc = GetDC(Me.hwnd)
'    dpiX = GetDeviceCaps(hdc, LOGPIXELSX)
'    dpiY = GetDeviceCaps(hdc, LOGPIXELSY)
'    ReleaseDC Me.hwnd, hdc
'
'    localTwipsPerPixelX = 1440 / dpiX
'    localTwipsPerPixelY = 1440 / dpiY
'
'    winWidth = CLng(Me.Width / localTwipsPerPixelX)
'    winHeight = CLng(Me.Height / localTwipsPerPixelY)
'
'    newLeft = mi.rcWork.Left + ((workWidth - winWidth) \ 2)
'    newTop = mi.rcWork.Top + ((workHeight - winHeight) \ 2)
'
'    SetWindowPos Me.hwnd, 0, newLeft, newTop, winWidth, winHeight, SWP_NOZORDER Or SWP_NOACTIVATE
Else
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
End If

out:
On Error Resume Next
Me.Visible = True
Exit Sub
error:
Call HandleError("Form_Load")
Resume out:
End Sub

