VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3855
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmAbout.frx":0000
      Top             =   480
      Width           =   4155
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   315
      Left            =   3360
      TabIndex        =   0
      Top             =   3420
      Width           =   975
   End
   Begin VB.Label lblMMESource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "View the project page on GitHub for support"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   120
      MousePointer    =   2  'Cross
      TabIndex        =   6
      Top             =   3060
      Width           =   4155
   End
   Begin VB.Label lblMudinfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "mudinfo.net - your source for mmud information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   120
      MousePointer    =   2  'Cross
      TabIndex        =   5
      Top             =   2760
      Width           =   4155
   End
   Begin VB.Label lblSynEmail 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "syntax53"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   2205
      MousePointer    =   2  'Cross
      TabIndex        =   2
      Top             =   3420
      Width           =   1155
   End
   Begin VB.Label lblb2yb 
      BackColor       =   &H00000000&
      Caption         =   "Brought to you by: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3420
      Width           =   1995
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   4275
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Dim TTlbl2 As clsToolTip

Private Sub cmdClose_Click()
Unload Me
End Sub


Private Sub Form_Load()
On Error Resume Next
Dim rc As RECT

Set TTlbl2 = New clsToolTip

With TTlbl2
    .DelayTime = 25
    .VisibleTime = 10000
    .BkColor = &HC0FFFF
    .TxtColor = &H0
    .Style = ttStyleBalloon
End With

TTlbl2.Style = 1

lblCaption.Caption = "MajorMUD Explorer"

rc.Left = lblMudinfo.Left \ Screen.TwipsPerPixelX
rc.Top = lblMudinfo.Top \ Screen.TwipsPerPixelY
rc.Bottom = (lblMudinfo.Top + lblMudinfo.Height) \ Screen.TwipsPerPixelX
rc.Right = (lblMudinfo.Left + lblMudinfo.Width) \ Screen.TwipsPerPixelY
TTlbl2.SetToolTipItem Me.hwnd, 0, rc.Left, rc.Top, rc.Right, rc.Bottom, "http://www.mudinfo.net/", False

rc.Left = lblSynEmail.Left \ Screen.TwipsPerPixelX
rc.Top = lblSynEmail.Top \ Screen.TwipsPerPixelY
rc.Bottom = (lblSynEmail.Top + lblSynEmail.Height) \ Screen.TwipsPerPixelX
rc.Right = (lblSynEmail.Left + lblSynEmail.Width) \ Screen.TwipsPerPixelY
TTlbl2.SetToolTipItem Me.hwnd, 0, rc.Left, rc.Top, rc.Right, rc.Bottom, "mailto: syntax53@mudinfo.net", False

If Not frmMain.WindowState = vbMinimized Then
    Me.Left = frmMain.Left + (frmMain.Width / 4)
    Me.Top = frmMain.Top + (frmMain.Height / 4)
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set TTlbl2 = Nothing
End Sub


Private Sub lblMMESource_Click()
Call ShellExecute(0&, "open", "https://github.com/syntax53/MMUD-Explorer/issues", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub lblMudinfo_Click()
Call ShellExecute(0&, "open", "http://www.mudinfo.net/", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub lblSynEmail_Click()
    Call ShellExecute(0&, "open", "mailto:syntax53@mudinfo.net &subject=MMUD Explorer", vbNullString, vbNullString, vbNormalFocus)
End Sub

