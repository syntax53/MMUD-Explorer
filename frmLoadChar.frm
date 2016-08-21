VERSION 5.00
Begin VB.Form frmLoadChar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Load Character ..."
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2595
   ControlBox      =   0   'False
   Icon            =   "frmLoadChar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   2595
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   2295
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   2475
      Begin VB.CheckBox chkInvenLoad 
         Caption         =   "Load Inventory Equipment"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   1  'Checked
         Width           =   2235
      End
      Begin VB.CheckBox chkCompareLoad 
         Caption         =   "Load Compare Lists"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox chkCompareClear 
         Caption         =   "Clear First"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1155
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "Filter all based on Char."
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   1740
         Width           =   2235
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "Leave current filtering"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   1500
         Value           =   -1  'True
         Width           =   2235
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "Reset Filters"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1980
         Width           =   2235
      End
      Begin VB.CheckBox chkInvenClear 
         Caption         =   "Clear First"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Value           =   1  'Checked
         Width           =   1155
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2340
         Y1              =   1380
         Y2              =   1380
      End
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "&Continue"
      Default         =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   2460
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Ca&ncel"
      Height          =   375
      Left            =   1500
      TabIndex        =   8
      Top             =   2460
      Width           =   1035
   End
End
Attribute VB_Name = "frmLoadChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkCompareLoad_Click()

If chkCompareLoad.Value = 1 Then
    chkCompareClear.Enabled = True
Else
    chkCompareClear.Enabled = False
End If

End Sub

Private Sub chkInvenLoad_Click()

If chkInvenLoad.Value = 1 Then
    chkInvenClear.Enabled = True
Else
    chkInvenClear.Enabled = False
End If

End Sub

Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdContinue_Click()
Me.Tag = "1"
Me.Hide
End Sub

'Private Sub optLoadFrom_Click(Index As Integer)
'
'If optLoadFrom(1).Value = True Then
'    chkInvenLoad.Value = 0
'    chkInvenLoad.Enabled = False
'    chkCompareLoad.Value = 0
'    chkCompareLoad.Enabled = False
'Else
'    chkInvenLoad.Value = 1
'    chkInvenLoad.Enabled = True
'    chkInvenClear.Value = 1
'    chkCompareLoad.Value = 1
'    chkCompareLoad.Enabled = True
'    chkCompareClear.Value = 1
'End If
'
'End Sub
Private Sub Form_Load()

End Sub
