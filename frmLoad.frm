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
   Picture         =   "frmLoad.frx":0000
   ScaleHeight     =   4350
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblCaption 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
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
      Left            =   1320
      MousePointer    =   11  'Hourglass
      TabIndex        =   0
      Top             =   3900
      Width           =   2355
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
Dim x As Long, y As Long
On Error Resume Next

x = Val(ReadINI("Settings", "Top", , 0))
y = Val(ReadINI("Settings", "Height", , 0))
If x <> 0 Then
    If y > 0 Then x = x + ((y - Me.Height) / 2)
    Me.Top = x
Else
    Me.Top = (Screen.Height - Me.Height) / 2
End If

x = Val(ReadINI("Settings", "Left", , 0))
y = Val(ReadINI("Settings", "Width", , 0))
If x <> 0 Then
    If y > 0 Then x = x + ((y - Me.Width) / 2)
    Me.Left = x
Else
    Me.Left = (Screen.Width - Me.Width) / 2
End If

End Sub
