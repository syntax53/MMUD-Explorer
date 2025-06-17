VERSION 5.00
Begin VB.Form frmCoinConvert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Coin Converter"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6210
   Icon            =   "frmCoinConvert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timWindowMove 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   660
      Top             =   3480
   End
   Begin VB.CommandButton cmdCharm 
      Caption         =   "Apply Charm"
      Height          =   510
      Left            =   4575
      TabIndex        =   16
      Top             =   3435
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   120
      Top             =   3480
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   435
      Left            =   120
      TabIndex        =   9
      Top             =   2940
      Width           =   5895
      Begin VB.OptionButton optCoinBottom 
         Caption         =   "Copper"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   60
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton optCoinBottom 
         Caption         =   "Silver"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1320
         TabIndex        =   11
         Top             =   60
         Width           =   1050
      End
      Begin VB.OptionButton optCoinBottom 
         Caption         =   "Gold"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   2400
         TabIndex        =   12
         Top             =   60
         Width           =   930
      End
      Begin VB.OptionButton optCoinBottom 
         Caption         =   "Platinum"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   3360
         TabIndex        =   13
         Top             =   60
         Width           =   1290
      End
      Begin VB.OptionButton optCoinBottom 
         Caption         =   "Runic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   4680
         TabIndex        =   14
         Top             =   60
         Width           =   1050
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   5895
      Begin VB.OptionButton optCoinTop 
         Caption         =   "Copper"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   1170
      End
      Begin VB.OptionButton optCoinTop 
         Caption         =   "Silver"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1320
         TabIndex        =   4
         Top             =   0
         Width           =   1050
      End
      Begin VB.OptionButton optCoinTop 
         Caption         =   "Gold"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   2400
         TabIndex        =   5
         Top             =   0
         Value           =   -1  'True
         Width           =   930
      End
      Begin VB.OptionButton optCoinTop 
         Caption         =   "Platinum"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   3360
         TabIndex        =   6
         Top             =   0
         Width           =   1290
      End
      Begin VB.OptionButton optCoinTop 
         Caption         =   "Runic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   4680
         TabIndex        =   7
         Top             =   0
         Width           =   1050
      End
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   315
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "frmCoinConvert.frx":0CCA
      Top             =   120
      Width           =   5520
   End
   Begin VB.TextBox txtCoin 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   2100
      MaxLength       =   10
      TabIndex        =   15
      Text            =   "100"
      Top             =   3480
      Width           =   1860
   End
   Begin VB.TextBox txtCoin 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   2100
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "1"
      Top             =   1740
      Width           =   1860
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "\/ Convert /\"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2400
      TabIndex        =   8
      Top             =   2670
      Width           =   1230
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   3900
      X2              =   6000
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   60
      X2              =   2100
      Y1              =   2820
      Y2              =   2820
   End
End
Attribute VB_Name = "frmCoinConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Enum eCoins
    Copper = 1#
    Silver = 10#
    Gold = 100#
    Platinum = 10000#
    Runic = 1000000#
End Enum

Public nLastPosTop As Long
Public nLastPosLeft As Long
Public nLastPosMoved As Long
Public nLastPosMonitor As Long

Public nLastTimerTop As Long
Public nLastTimerLeft As Long

Dim tWindowSize As WindowSizeProperties
Dim eCurrentCoin(1) As eCoins

Private Sub EnableCharmButton()
On Error GoTo error:
Dim nCharmMod As Double, nCharm As Integer, sCharmMod As String
nCharm = Val(frmMain.txtCharStats(5).Text)

nCharmMod = 1 - ((Fix(nCharm / 5) - 10) / 100)
If nCharmMod > 1 Then
    cmdCharm.Tag = CCur(nCharmMod) * 100
    sCharmMod = Abs(1 - CCur(nCharmMod)) * 100 & "% Markup"
ElseIf nCharmMod < 1 Then
    cmdCharm.Tag = CCur(nCharmMod) * 100
    sCharmMod = Val(1 - CCur(nCharmMod)) * 100 & "% Discount"
End If

If nCharmMod = 0 Then
    cmdCharm.Visible = False
Else
    cmdCharm.Caption = "Apply" & vbCrLf & sCharmMod
    cmdCharm.Visible = True
End If

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("EnableCharmButton")
Resume out:
End Sub

Private Sub cmdCharm_Click()
Dim nCopper As Double
On Error GoTo error:

If Val(txtCoin(1).Text) < 1 Then Exit Sub

nCopper = ConvertCoin(Val(txtCoin(1).Text), eCurrentCoin(1), Copper)
nCopper = Round(nCopper * (Val(cmdCharm.Tag) / 100), 8)

txtCoin(1).Text = ConvertCoin(nCopper, Copper, eCurrentCoin(1))

If Val(txtCoin(1).Text) < 1 Then txtCoin(1).Text = 1
If Val(txtCoin(1).Text) > 999999999# Then txtCoin(1).Text = 999999999#

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("cmdCharm_Click")
Resume out:
End Sub

Private Sub Form_Load()
On Error GoTo error:
Dim x As Integer, y As Integer

Call SetWindowLong(Me.hWnd, GWL_HWNDPARENT, 0)

'stop windows from resizing fixed-size windows when changing dpi
If bDPIAwareMode Then Call SubclassFormMinMaxSize(Me, tWindowSize, True)

If frmMain.WindowState = vbMinimized Then
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
Else
    Me.Left = frmMain.Left + ((frmMain.Width - Me.Width) / 2)
    Me.Top = frmMain.Top + ((frmMain.Height - Me.Height) / 2)
End If


eCurrentCoin(0) = Gold
eCurrentCoin(1) = Copper
txtCoin(0).Text = 1
txtCoin(1).Text = 100

Call CalcCoin(0, 1)

If frmMain.chkGlobalFilter.Value = 1 And Val(frmMain.txtCharStats(5).Text) > 0 Then Call EnableCharmButton

timWindowMove.Enabled = True

Exit Sub
error:
Call HandleError("CoinConvert_Load")
Resume Next
End Sub

Private Sub CalcCoin(nSourceIndex As Integer, nDestinationIndex As Integer, Optional ByVal eForceSource As eCoins)
On Error GoTo error:
Dim eCoinFrom As eCoins, eCoinTo As eCoins

If nSourceIndex = 0 Or nDestinationIndex = 0 Then
    If optCoinTop(4).Value = True Then
        If nSourceIndex = 0 Then eCoinFrom = Runic
        If nDestinationIndex = 0 Then eCoinTo = Runic
    ElseIf optCoinTop(3).Value = True Then
        If nSourceIndex = 0 Then eCoinFrom = Platinum
        If nDestinationIndex = 0 Then eCoinTo = Platinum
    ElseIf optCoinTop(2).Value = True Then
        If nSourceIndex = 0 Then eCoinFrom = Gold
        If nDestinationIndex = 0 Then eCoinTo = Gold
    ElseIf optCoinTop(1).Value = True Then
        If nSourceIndex = 0 Then eCoinFrom = Silver
        If nDestinationIndex = 0 Then eCoinTo = Silver
    Else
        If nSourceIndex = 0 Then eCoinFrom = Copper
        If nDestinationIndex = 0 Then eCoinTo = Copper
    End If
End If

If nSourceIndex = 1 Or nDestinationIndex = 1 Then
    If optCoinBottom(4).Value = True Then
        If nSourceIndex = 1 Then eCoinFrom = Runic
        If nDestinationIndex = 1 Then eCoinTo = Runic
    ElseIf optCoinBottom(3).Value = True Then
        If nSourceIndex = 1 Then eCoinFrom = Platinum
        If nDestinationIndex = 1 Then eCoinTo = Platinum
    ElseIf optCoinBottom(2).Value = True Then
        If nSourceIndex = 1 Then eCoinFrom = Gold
        If nDestinationIndex = 1 Then eCoinTo = Gold
    ElseIf optCoinBottom(1).Value = True Then
        If nSourceIndex = 1 Then eCoinFrom = Silver
        If nDestinationIndex = 1 Then eCoinTo = Silver
    Else
        If nSourceIndex = 1 Then eCoinFrom = Copper
        If nDestinationIndex = 1 Then eCoinTo = Copper
    End If
End If

If eForceSource > 0 Then eCoinFrom = eForceSource

eCurrentCoin(nDestinationIndex) = eCoinTo
txtCoin(nDestinationIndex).Text = ConvertCoin(Val(txtCoin(nSourceIndex).Text), eCoinFrom, eCoinTo)

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("CalcCoin")
Resume out:
End Sub

Private Function ConvertCoin(ByVal nCoins As Double, ByVal ConvertFrom As eCoins, ByVal ConvertTo As eCoins) As Double
On Error GoTo error:
Dim nCopper As Double

nCopper = nCoins * ConvertFrom
If nCopper > 9999999999# Then nCopper = 9999999999#
ConvertCoin = Round(nCopper / ConvertTo, 8)

out:
On Error Resume Next
Exit Function
error:
Call HandleError("ConvertCoin")
Resume out:
End Function

Private Sub optCoinBottom_Click(Index As Integer)
Call CalcCoin(1, 1, eCurrentCoin(1))
End Sub

Private Sub optCoinTop_Click(Index As Integer)
Call CalcCoin(0, 0, eCurrentCoin(0))
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
If Timer1.Tag = 0 Then
    Call CalcCoin(0, 1)
Else
    Call CalcCoin(1, 0)
End If
End Sub

Private Sub timWindowMove_Timer()
Call MonitorFormTimer(Me)
End Sub

Private Sub txtCoin_GotFocus(Index As Integer)
Call SelectAll(txtCoin(Index))
End Sub

Private Sub txtCoin_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii, True)
End Sub

Private Sub txtCoin_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Call CalcDelay(Index)
End Sub

Private Sub CalcDelay(nIndex As Integer)

Timer1.Enabled = False
Timer1.Tag = nIndex
Timer1.Enabled = True

End Sub

Private Sub txtCoin_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Call CalcDelay(Index)
End Sub
