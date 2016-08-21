VERSION 5.00
Begin VB.Form frmMapLegend 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Legend"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "frmMapLegend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6240
   Begin VB.Label lblRoomCell 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   135
      Index           =   1
      Left            =   2340
      TabIndex        =   27
      Top             =   1800
      Width           =   135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   4080
      X2              =   4080
      Y1              =   60
      Y2              =   2580
   End
   Begin VB.Label lblRoomCell 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   135
      Index           =   0
      Left            =   2340
      TabIndex        =   26
      Top             =   1110
      Width           =   135
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "NPC Assigned"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   15
      Left            =   2640
      TabIndex        =   25
      Top             =   1110
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "- Right click on an up/down exit to have the option of following and redrawing."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   4260
      TabIndex        =   24
      Top             =   2040
      Width           =   1875
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Lair Room"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   17
      Left            =   2640
      TabIndex        =   23
      Top             =   1440
      Width           =   1275
   End
   Begin VB.Label lblRoomCell 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   135
      Index           =   2
      Left            =   2340
      TabIndex        =   22
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Command In Room (Remote or Local)"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   16
      Left            =   2640
      TabIndex        =   21
      Top             =   1740
      Width           =   1335
   End
   Begin VB.Label lblRoomCell 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   135
      Index           =   8
      Left            =   2340
      TabIndex        =   20
      Top             =   2220
      Width           =   135
   End
   Begin VB.Label lblRoomCell 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   135
      Index           =   7
      Left            =   2340
      TabIndex        =   19
      Top             =   480
      Width           =   135
   End
   Begin VB.Label lblRoomCell 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   135
      Index           =   6
      Left            =   2340
      TabIndex        =   18
      Top             =   180
      Width           =   135
   End
   Begin VB.Label lblRoomCell 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   5
      Left            =   2340
      TabIndex        =   17
      Top             =   780
      Width           =   135
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Starting Point"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   13
      Left            =   2640
      TabIndex        =   16
      Top             =   2220
      Width           =   1275
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Class / Race / Level / Alignment / Ability"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   9
      Left            =   840
      TabIndex        =   15
      Top             =   1620
      Width           =   1275
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Timed"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   8
      Left            =   840
      TabIndex        =   14
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Line linLegend 
      BorderColor     =   &H00000040&
      BorderWidth     =   5
      Index           =   9
      X1              =   180
      X2              =   660
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Line linLegend 
      BorderColor     =   &H00004000&
      BorderWidth     =   5
      Index           =   8
      X1              =   180
      X2              =   660
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Door / Gate"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   7
      Left            =   840
      TabIndex        =   13
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Hidden"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   5
      Left            =   840
      TabIndex        =   12
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Key / Item / Toll"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   4
      Left            =   840
      TabIndex        =   11
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   3
      Left            =   840
      TabIndex        =   10
      Top             =   900
      Width           =   1095
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Trap / Spell Trap"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   2
      Left            =   840
      TabIndex        =   9
      Top             =   660
      Width           =   1095
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Map Change"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   1
      Left            =   840
      TabIndex        =   8
      Top             =   420
      Width           =   1095
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   7
      Top             =   180
      Width           =   1095
   End
   Begin VB.Line linLegend 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      Index           =   7
      X1              =   180
      X2              =   660
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line linLegend 
      BorderColor     =   &H00800080&
      BorderWidth     =   5
      Index           =   5
      X1              =   180
      X2              =   660
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Line linLegend 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   5
      Index           =   4
      X1              =   180
      X2              =   660
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line linLegend 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   5
      Index           =   3
      X1              =   180
      X2              =   660
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line linLegend 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Index           =   2
      X1              =   180
      X2              =   660
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line linLegend 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   5
      Index           =   1
      X1              =   180
      X2              =   660
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line linLegend 
      BorderColor     =   &H00808080&
      BorderWidth     =   5
      Index           =   0
      X1              =   180
      X2              =   660
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "-Right Click: Redraw map from selected room"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   4260
      TabIndex        =   6
      Top             =   480
      Width           =   1875
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "-Hover mouse over room to see Information"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   4260
      TabIndex        =   5
      Top             =   1620
      Width           =   1875
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "-Left Click: View record references of selected room."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   4260
      TabIndex        =   4
      Top             =   60
      Width           =   1875
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "-Shift+Right Click: Redraw map using selected cell as the ""center"".  This may also be done on unused cells."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   675
      Left            =   4260
      TabIndex        =   3
      Top             =   900
      Width           =   1875
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Exit Up"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   10
      Left            =   2640
      TabIndex        =   2
      Top             =   180
      Width           =   1275
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Exit Down"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   11
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   1275
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Exit Up and Down"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   12
      Left            =   2640
      TabIndex        =   0
      Top             =   780
      Width           =   1275
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   225
      Left            =   2295
      Top             =   2175
      Width           =   225
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      Height          =   255
      Left            =   2280
      Shape           =   1  'Square
      Top             =   1740
      Width           =   255
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FF00FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   4
      Height          =   195
      Left            =   2310
      Shape           =   3  'Circle
      Top             =   1410
      Width           =   195
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   255
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   1050
      Width           =   255
   End
End
Attribute VB_Name = "frmMapLegend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit
Public objFormOwner As Form

Private Sub Form_Load()
Dim nTmp As Long

On Error GoTo Error:

nTmp = ReadINI("Settings", "LegendTop")
Me.Top = IIf(nTmp > 1, nTmp, Me.ScaleHeight / 2)

nTmp = ReadINI("Settings", "LegendLeft")
Me.Left = IIf(nTmp > 1, nTmp, Me.ScaleWidth / 2)

Exit Sub

Error:
Call HandleError("Form_Load")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If Not Me.WindowState = vbMinimized And Not Me.WindowState = vbMaximized Then
    Call WriteINI("Settings", "LegendTop", Me.Top)
    Call WriteINI("Settings", "LegendLeft", Me.Left)
End If
If Not objFormOwner Is Nothing Then
    If Not bAppTerminating Then objFormOwner.cmdViewMapLegend.Tag = "0"
    Set objFormOwner = Nothing
End If
End Sub
