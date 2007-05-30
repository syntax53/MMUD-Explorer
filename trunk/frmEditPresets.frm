VERSION 5.00
Begin VB.Form frmEditPreset 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton cmdGetStats 
         Caption         =   "&Get Current Rooms Stats"
         Height          =   255
         Left            =   120
         MaskColor       =   &H8000000F&
         TabIndex        =   3
         Top             =   1320
         Width           =   3255
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   2460
         TabIndex        =   5
         Top             =   1740
         Width           =   975
      End
      Begin VB.TextBox txtCaption 
         Height          =   285
         Left            =   960
         MaxLength       =   25
         TabIndex        =   2
         Top             =   900
         Width           =   2415
      End
      Begin VB.TextBox txtRoom 
         Height          =   285
         Left            =   2520
         TabIndex        =   1
         Top             =   540
         Width           =   855
      End
      Begin VB.TextBox txtMap 
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   540
         Width           =   795
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Caption"
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
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Room"
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
         Height          =   195
         Left            =   1920
         TabIndex        =   9
         Top             =   540
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Map"
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
         Left            =   480
         TabIndex        =   8
         Top             =   540
         Width           =   375
      End
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Editing Preset #"
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
         TabIndex        =   7
         Top             =   180
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmEditPreset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Public objFormOwner As Form
Public nPreset As Integer


Private Sub cmdCancel_Click()
nPreset = -1
Me.Hide
End Sub

Private Sub cmdGetStats_Click()
txtMap.Text = objFormOwner.nMapStartMap
txtRoom.Text = objFormOwner.nMapStartRoom
txtCaption.Text = GetRoomName(, objFormOwner.nMapStartMap, objFormOwner.nMapStartRoom, True)
cmdOK.SetFocus
End Sub

Private Sub cmdOK_Click()
Me.Hide
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set objFormOwner = Nothing
End Sub

Private Sub txtCaption_GotFocus()
Call SelectAll(txtCaption)
End Sub

Private Sub txtMap_GotFocus()
Call SelectAll(txtMap)
End Sub

Private Sub txtRoom_GotFocus()
Call SelectAll(txtRoom)
End Sub
