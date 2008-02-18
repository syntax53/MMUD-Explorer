VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3795
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmAbout.frx":0000
      Top             =   540
      Width           =   4155
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   315
      Left            =   3360
      TabIndex        =   0
      Top             =   3060
      Width           =   975
   End
   Begin VB.Label lblGhaleonLink 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "ghaleon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2205
      MousePointer    =   2  'Cross
      TabIndex        =   7
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblb2yb2 
      BackColor       =   &H00000000&
      Caption         =   "Post 1.68 Versions by:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   2160
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
      ForeColor       =   &H00FF0000&
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2205
      MousePointer    =   2  'Cross
      TabIndex        =   2
      Top             =   3060
      Width           =   1155
   End
   Begin VB.Label lblb2yb 
      BackColor       =   &H00000000&
      Caption         =   "Brought to you by: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3060
      Width           =   1635
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
      ForeColor       =   &H00C00000&
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

lblCaption.Caption = frmMain.Caption

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

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set TTlbl2 = Nothing
End Sub

Private Sub lblGhaleonLink_Click()
    Call ShellExecute(0&, "open", "telnet:quicksilverbbs.com", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub lblMudinfo_Click()
Call ShellExecute(0&, "open", "http://www.mudinfo.net/", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub lblSynEmail_Click()
    Call ShellExecute(0&, "open", "mailto:syntax53@mudinfo.net &subject=MMUD Explorer", vbNullString, vbNullString, vbNormalFocus)
End Sub

