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
   Picture         =   "frmLoad.frx":0000
   ScaleHeight     =   4350
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
