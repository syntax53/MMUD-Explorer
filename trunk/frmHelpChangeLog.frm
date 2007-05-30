VERSION 5.00
Begin VB.Form frmHelpChangeLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ChangeLog"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   Icon            =   "frmHelpChangeLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6810
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   4935
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmHelpChangeLog.frx":0CCA
      Top             =   0
      Width           =   6795
   End
End
Attribute VB_Name = "frmHelpChangeLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Private Sub Form_Load()
On Error Resume Next

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
 
End Sub

