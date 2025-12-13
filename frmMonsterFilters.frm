VERSION 5.00
Begin VB.Form frmMonsterFilters 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monster Filter"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkMonsterUndead 
      Height          =   255
      Left            =   1140
      TabIndex        =   3
      ToolTipText     =   "Only Undead"
      Top             =   1200
      Width           =   195
   End
   Begin VB.CheckBox chkMonsterDropCash 
      Height          =   255
      Left            =   1140
      TabIndex        =   2
      ToolTipText     =   "Drops Coin"
      Top             =   840
      Width           =   195
   End
   Begin VB.PictureBox picMonsterFilterPics 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1380
      Picture         =   "frmMonsterFilters.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Drops Coin"
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox picMonsterFilterPics 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1380
      Picture         =   "frmMonsterFilters.frx":027D
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Only Undead"
      Top             =   1185
      Width           =   255
   End
End
Attribute VB_Name = "frmMonsterFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
