VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4725
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00E0E0E0&
      Height          =   2895
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmAbout.frx":0000
      Top             =   480
      Width           =   7515
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   315
      Left            =   6660
      TabIndex        =   0
      Top             =   4260
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
      Left            =   1680
      MousePointer    =   2  'Cross
      TabIndex        =   6
      Top             =   3900
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
      Left            =   1680
      MousePointer    =   2  'Cross
      TabIndex        =   5
      Top             =   3600
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
      Left            =   4185
      MousePointer    =   2  'Cross
      TabIndex        =   2
      Top             =   4320
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
      Left            =   2100
      TabIndex        =   3
      Top             =   4320
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
      Width           =   7695
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
'SubclassForm Me
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

TTlbl2.DelToolTip Me.hWnd, 0
rc.Left = lblMudinfo.Left ' \ Screen.TwipsPerPixelX
rc.Top = lblMudinfo.Top ' \ Screen.TwipsPerPixelY
rc.Bottom = (lblMudinfo.Top + lblMudinfo.Height) ' \ Screen.TwipsPerPixelX
rc.Right = (lblMudinfo.Left + lblMudinfo.Width) ' \ Screen.TwipsPerPixelY
TTlbl2.SetToolTipItem Me.hWnd, 0, _
    ConvertScale(rc.Left, vbTwips, vbPixels), _
    ConvertScale(rc.Top, vbTwips, vbPixels), _
    ConvertScale(rc.Right, vbTwips, vbPixels), _
    ConvertScale(rc.Bottom, vbTwips, vbPixels), _
    "http://www.mudinfo.net/", False

TTlbl2.DelToolTip Me.hWnd, 1
rc.Left = lblSynEmail.Left ' \ Screen.TwipsPerPixelX
rc.Top = lblSynEmail.Top ' \ Screen.TwipsPerPixelY
rc.Bottom = (lblSynEmail.Top + lblSynEmail.Height) ' \ Screen.TwipsPerPixelX
rc.Right = (lblSynEmail.Left + lblSynEmail.Width) ' \ Screen.TwipsPerPixelY
TTlbl2.SetToolTipItem Me.hWnd, 1, _
    ConvertScale(rc.Left, vbTwips, vbPixels), _
    ConvertScale(rc.Top, vbTwips, vbPixels), _
    ConvertScale(rc.Right, vbTwips, vbPixels), _
    ConvertScale(rc.Bottom, vbTwips, vbPixels), _
    "mailto: syntax53@mudinfo.net", False

If frmMain.WindowState = vbMinimized Then
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
Else
    Me.Left = frmMain.Left + ((frmMain.Width - Me.Width) / 2)
    Me.Top = frmMain.Top + ((frmMain.Height - Me.Height) / 2)
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

