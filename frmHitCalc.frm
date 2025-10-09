VERSION 5.00
Begin VB.Form frmHitCalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hit Calculator"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6480
   Icon            =   "frmHitCalc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timButtonPress 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5880
      Top             =   120
   End
   Begin VB.Timer timWindowMove 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.Frame Frame4 
      Height          =   6075
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   435
         Left            =   180
         TabIndex        =   22
         Top             =   240
         Width           =   5895
         Begin VB.OptionButton optHitCalcType 
            Caption         =   "Backstab"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   4020
            TabIndex        =   24
            Top             =   90
            Width           =   1530
         End
         Begin VB.OptionButton optHitCalcType 
            Caption         =   "Normal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   2400
            TabIndex        =   23
            Top             =   90
            Value           =   -1  'True
            Width           =   1350
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Calculate For:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   25
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   435
         Left            =   240
         TabIndex        =   18
         Top             =   1800
         Width           =   6015
         Begin VB.OptionButton optDefender 
            Caption         =   "Manual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   4680
            TabIndex        =   36
            Top             =   120
            Width           =   1230
         End
         Begin VB.OptionButton optDefender 
            Caption         =   "vs Player"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   3060
            TabIndex        =   21
            Top             =   120
            Width           =   1470
         End
         Begin VB.OptionButton optDefender 
            Caption         =   "vs Mob"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   1620
            TabIndex        =   20
            Top             =   120
            Width           =   1230
         End
         Begin VB.OptionButton optDefender 
            Caption         =   "vs Char."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   19
            Top             =   120
            Value           =   -1  'True
            Width           =   1410
         End
      End
      Begin VB.CommandButton cmdCharHitCalc 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2340
         TabIndex        =   17
         Top             =   3720
         Width           =   315
      End
      Begin VB.CommandButton cmdCharHitCalc 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   3660
         TabIndex        =   16
         Top             =   3720
         Width           =   315
      End
      Begin VB.CommandButton cmdCharHitCalc 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   4260
         TabIndex        =   15
         Top             =   3720
         Width           =   315
      End
      Begin VB.CommandButton cmdCharHitCalc 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   5760
         TabIndex        =   14
         Top             =   3720
         Width           =   315
      End
      Begin VB.CommandButton cmdCharHitCalc 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   3720
         Width           =   315
      End
      Begin VB.CommandButton cmdCharHitCalc 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1740
         TabIndex        =   12
         Top             =   3720
         Width           =   315
      End
      Begin VB.TextBox txtHitCalc 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   600
         MaxLength       =   5
         TabIndex        =   11
         Text            =   "0"
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox txtHitCalc 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   4620
         MaxLength       =   5
         TabIndex        =   10
         Text            =   "10"
         ToolTipText     =   "Dodge value, not dodge %"
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox txtHitCalc 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   600
         MaxLength       =   4
         TabIndex        =   9
         Text            =   "100"
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox txtHitCalc 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   2700
         MaxLength       =   4
         TabIndex        =   8
         Text            =   "0"
         Top             =   3720
         Width           =   915
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   990
         Width           =   6015
         Begin VB.OptionButton optAttacker 
            Caption         =   "Current Char."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   7
            Top             =   120
            Value           =   -1  'True
            Width           =   1890
         End
         Begin VB.OptionButton optAttacker 
            Caption         =   "Monster"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   2340
            TabIndex        =   6
            Top             =   120
            Width           =   1290
         End
         Begin VB.OptionButton optAttacker 
            Caption         =   "Manual "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   3960
            TabIndex        =   5
            Top             =   120
            Width           =   1710
         End
      End
      Begin VB.ComboBox cmbMonsterList 
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
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   2640
         Width           =   4815
      End
      Begin VB.CommandButton cmdCharHitCalc 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   1740
         TabIndex        =   2
         Top             =   4800
         Width           =   315
      End
      Begin VB.CommandButton cmdCharHitCalc 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   240
         TabIndex        =   1
         Top             =   4800
         Width           =   315
      End
      Begin VB.Label lblSubLeft 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         TabIndex        =   35
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   180
         X2              =   2220
         Y1              =   1740
         Y2              =   1740
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   4020
         X2              =   6120
         Y1              =   1740
         Y2              =   1740
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Defender"
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
         Left            =   2520
         TabIndex        =   34
         Top             =   1590
         Width           =   1230
      End
      Begin VB.Label lbl2ndD 
         Alignment       =   2  'Center
         Caption         =   "Secondary Defense:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   300
         TabIndex        =   33
         Top             =   4260
         Width           =   1710
      End
      Begin VB.Label lblDodge 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "vs Dodge:"
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
         Left            =   4635
         TabIndex        =   32
         Top             =   3420
         Width           =   1080
      End
      Begin VB.Label lblACC 
         Alignment       =   2  'Center
         Caption         =   "Attacker Accuracy:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   630
         TabIndex        =   31
         Top             =   3180
         Width           =   1050
      End
      Begin VB.Label lblAC 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "vs AC:"
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
         Left            =   2790
         TabIndex        =   30
         Top             =   3420
         Width           =   810
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Attacker"
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
         Left            =   2520
         TabIndex        =   29
         Top             =   780
         Width           =   1230
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   4020
         X2              =   6120
         Y1              =   930
         Y2              =   930
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   180
         X2              =   2220
         Y1              =   930
         Y2              =   930
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         Index           =   4
         X1              =   180
         X2              =   6180
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label lblLabelArray 
         AutoSize        =   -1  'True
         Caption         =   "Monster:"
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
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   2700
         Width           =   900
      End
      Begin VB.Label lblMain 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   2340
         TabIndex        =   27
         Top             =   4320
         Width           =   3735
      End
      Begin VB.Label lblSubRight 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2340
         TabIndex        =   26
         Top             =   5280
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmHitCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Public nLastPosTop As Long
Public nLastPosLeft As Long
Public nLastPosMoved As Long
Public nLastPosMonitor As Long

Public nLastTimerTop As Long
Public nLastTimerLeft As Long

Dim nMonsterAccy As Long
Dim nMonsterAC As Long
Dim nMonsterDodge As Long
Dim nMonsterBSdef As Long

Dim nDefenderAC As Long
Dim nDefenderDodge As Long
Dim bDefenderSeeHidden As Boolean
Dim nDefenderPerception As Long
Dim nDefenderProtEvil As Long
Dim nDefenderVileWard As Long
Dim nDefenderEvilness As Long
Dim bDefenderShadow As Boolean
Dim nDefenderClass As Long

Dim nCharVileWard As Long
Dim nCharPerception As Long
Dim nCharEvilness As Long
Dim bCharSeeHidden As Boolean

Dim ntimButtonPressCount As Long
Dim nPlayer
Dim bDontRefresh As Boolean
Dim bMouseDown As Boolean
Dim tWindowSize As WindowSizeProperties

Private Sub PromptForPlayerSecondaryDefense(Optional ByVal nPREV As Long, Optional ByVal nVileWard As Long, Optional ByVal nEvilNess As Integer, _
    Optional ByVal bShadow As Boolean, Optional ByVal nPercep As Long)
On Error GoTo error:
Dim str As String

If nPercep >= 0 Then
    str = InputBox("Enter defender's preception value (for BS only)", "Player Secondary Defense Calculation", nPercep)
    If str = "" Then Exit Sub
    If val(str) < 0 Then
        nDefenderPerception = 0
    ElseIf val(str) > 9999 Then
        nDefenderPerception = 9999
    Else
        nDefenderPerception = val(str)
    End If
End If

str = InputBox("Enter defender's protection from evil value", "Player Secondary Defense Calculation", nPREV)
If str = "" Then Exit Sub
If val(str) < 0 Then
    nDefenderProtEvil = 0
ElseIf val(str) > 9999 Then
    nDefenderProtEvil = 9999
Else
    nDefenderProtEvil = val(str)
End If

str = InputBox("Enter defender's vile ward value (before being divided by 10)", "Player Secondary Defense Calculation", nVileWard)
If str = "" Then Exit Sub
If val(str) < 0 Then
    nDefenderVileWard = 0
ElseIf val(str) > 9999 Then
    nDefenderVileWard = 9999
Else
    nDefenderVileWard = val(str)
End If

If nDefenderVileWard > 0 Then
    str = InputBox("Enter defender's ""evilness"" (for vile ward)" & vbCrLf & vbCrLf _
                & "Answer 0 for seedy or less (no value), 1 for outlaw/criminal (50% value), or 2 for villian+", _
                        "Player Secondary Defense Calculation", nDefenderEvilness)
    If val(str) < 0 Then
        nDefenderEvilness = 0
    ElseIf val(str) > 2 Then
        nDefenderEvilness = 2
    Else
        nDefenderEvilness = val(str)
    End If
End If

str = InputBox("Does defender have shadow/shadowstealth?" & vbCrLf & vbCrLf & "Answer 0 for no, 1 for yes", "Player Secondary Defense Calculation", IIf(bShadow, 1, 0))
If str = "" Then Exit Sub
If val(str) > 0 Then
    bDefenderShadow = True
Else
    bDefenderShadow = False
End If

out:
Exit Sub
error:
Call HandleError("PromptForPlayerSecondaryDefense")
Resume out:
End Sub

Private Sub PromptForPlayerDodgeBSDefense(Optional ByVal nDodge As Long, Optional ByVal nPercep As Long, Optional ByVal bSeeH As Boolean)
On Error GoTo error:
Dim str As String

str = InputBox("Enter defender's dodge value", "Player BS Defense Calculation", nDodge)
If str = "" Then Exit Sub
If val(str) < 0 Then
    nDefenderDodge = 0
ElseIf val(str) > 9999 Then
    nDefenderDodge = 9999
Else
    nDefenderDodge = val(str)
End If

str = InputBox("Enter defender's perception value", "Player BS Defense Calculation", nPercep)
If str = "" Then Exit Sub
If val(str) < 0 Then
    nDefenderPerception = 0
ElseIf val(str) > 9999 Then
    nDefenderPerception = 9999
Else
    nDefenderPerception = val(str)
End If

str = InputBox("Does defender have see hidden?" & vbCrLf & vbCrLf & "Answer 0 for no, 1 for yes", "Player BS Defense Calculation", IIf(bSeeH, 1, 0))
If str = "" Then Exit Sub
If val(str) > 0 Then
    bDefenderSeeHidden = True
Else
    bDefenderSeeHidden = False
End If

out:
Exit Sub
error:
Call HandleError("PromptForPlayerDodgeBSDefense")
Resume out:
End Sub

Public Sub DoHitCalc()
On Error GoTo error:
Dim nAC As Long, nAccy As Long, nHitChance As Currency, nTotalHitPercent As Currency
Dim nDodge As Long, nDodgeChance As Currency, sArr() As String
Dim sPrint As String, nTemp As Long, nAux As Long, nShadow As Long
Dim nSecondaryDef As Long, nProtEv As Long, nPerception As Long, nVileWard As Long, eEvil As eEvilPoints
Dim bShadow As Boolean, bSeeHidden As Boolean, bBackstab As Boolean, bVSplayer As Boolean
Dim nClass As Integer, nDefense() As Long

nAccy = Fix(val(txtHitCalc(0).Text))
nAC = Fix(val(txtHitCalc(1).Text))
nDodge = Fix(val(txtHitCalc(2).Text))
nAux = Fix(val(txtHitCalc(3).Text))

If nAccy > 9999 Then nAccy = 9999: If nAccy < 1 Then nAccy = 1
If nAC > 9999 Then nAC = 9999: If nAC < 0 Then nAC = 0
If nDodge > 9999 Then nDodge = 9999: If nDodge < -999 Then nDodge = 0
If nAux > 9999 Then nAux = 9999: If nAux < 0 Then nAux = 0

If optHitCalcType(1).Value = True Then bBackstab = True
If optDefender(2).Value = True Then bVSplayer = True

If optDefender(0).Value = 1 Then
    nClass = frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex)
ElseIf optDefender(2).Value = 1 Then
    nClass = nDefenderClass
Else
    nClass = -1
End If

'GET HIT CHANCE
If nAC + nAux > 0 Or (bBackstab And bVSplayer And bGreaterMUD And Len(txtHitCalc(3).Tag) > 0) Then
    If bBackstab Then '[BACKSTAB]
        If bVSplayer Then '[BACKSTAB+PLAYER]
            If bGreaterMUD Then '[BACKSTAB+PLAYER+GREATERMUD]
                sArr = Split(txtHitCalc(3).Tag, ",", , vbTextCompare)
                If UBound(sArr) < 4 Then ReDim sArr(4)
                nPerception = val(sArr(0))
                nProtEv = val(sArr(1))
                nVileWard = val(sArr(2))
                nTemp = val(sArr(3))
                If nTemp >= 2 Then
                    eEvil = e7_FIEND
                ElseIf nTemp >= 1 Then
                    eEvil = e5_Criminal
                End If
                nShadow = val(sArr(4))
                If nShadow > 0 Then bShadow = True
            Else '[BACKSTAB+PLAYER+STOCK]
                nPerception = nAux
            End If
        Else '[BACKSTAB+MOB] (same for stock and gmud)
            nSecondaryDef = nAux 'bs defense in this case
        End If
    Else 'NORMAL ATTACK
        nSecondaryDef = nAux 'actual secondary defenses
    End If
End If

If bGreaterMUD Then
    If (nDodge > 0 Or (bBackstab And bVSplayer And Len(txtHitCalc(2).Tag) > 0)) Then
        If bBackstab And bVSplayer And Len(txtHitCalc(2).Tag) > 0 Then
            sArr = Split(txtHitCalc(2).Tag, ",", , vbTextCompare)
            If UBound(sArr) < 2 Then ReDim sArr(2)
            nDodge = val(sArr(0))
            nPerception = val(sArr(1))
            If val(sArr(2)) > 0 Then bSeeHidden = True
        End If
    End If
End If

'need implement protection from good...
nDefense = CalculateAttackDefense(nAccy, nAC, nDodge, nSecondaryDef, nProtEv, 0, nPerception, _
    nVileWard, eEvil, bShadow, bSeeHidden, bBackstab, bVSplayer, nClass)

nHitChance = nDefense(0)
nDodgeChance = nDefense(1)

If bBackstab And bVSplayer Then
    If nSecondaryDef > 0 And Len(txtHitCalc(3).Tag) > 0 Then txtHitCalc(3).Text = nSecondaryDef
    If nDodge > 0 And Len(txtHitCalc(2).Tag) > 0 Then txtHitCalc(2).Text = nDodge
End If

prin:
sPrint = "Hit: " & nHitChance & "%"
If nDodgeChance > 0 Then
    sPrint = sPrint & vbCrLf & "Dodge: " & Round(nDodgeChance) & "%"
    nTotalHitPercent = Round(nHitChance - ((nHitChance * (nDodgeChance / 100))))
    sPrint = sPrint & vbCrLf & "Overall Hit: " & nTotalHitPercent & "%"
Else
    sPrint = sPrint & vbCrLf & "Dodge: 0%"
    sPrint = sPrint & vbCrLf & "Overall Hit: " & nHitChance & "%"
End If

lblMain.Caption = sPrint

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("DoHitCalc")
Resume out:
End Sub

Private Sub cmbMonsterList_Click()
On Error GoTo error:

If cmbMonsterList.ListCount < 1 Then Exit Sub
If cmbMonsterList.ListIndex < 0 Then Exit Sub
If cmbMonsterList.ItemData(cmbMonsterList.ListIndex) < 1 Then Exit Sub
If GetMonsterData(cmbMonsterList.ItemData(cmbMonsterList.ListIndex)) Then
    If optAttacker(1).Value = True Then 'attacking - accy
        txtHitCalc(0).Text = nMonsterAccy
    End If
    If optDefender(1).Value = True Then 'defending - ac + dodge
        txtHitCalc(1).Text = nMonsterAC
        txtHitCalc(2).Text = nMonsterDodge
        If optHitCalcType(1).Value = True Then 'backstab - bs defense
            txtHitCalc(3).Text = nMonsterBSdef
        End If
    End If
End If

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("cmbMonsterList_Click")
Resume out:
End Sub

Private Sub cmbMonsterList_KeyPress(KeyAscii As Integer)
KeyAscii = AutoComplete(cmbMonsterList, KeyAscii, False)
End Sub

Public Sub SetHitCalcVals()
On Error GoTo error:
Dim nBSWep As Long, nBSAccyAdj As Integer, nNormAccyAdj As Integer
Dim bCharAttack As Boolean, bMobAttack As Boolean, bManualAttack As Boolean
Dim bVSchar As Boolean, bVSmob As Boolean, bVSplayer As Boolean, bVSmanual As Boolean
Dim bNormal As Boolean, bBackstab As Boolean

If optHitCalcType(0).Value = True Then
    bNormal = True
ElseIf optHitCalcType(1).Value = True Then
    bBackstab = True
Else
    Exit Sub
End If

If optAttacker(0).Value = True Then
    bCharAttack = True
ElseIf optAttacker(1).Value = True Then
    bMobAttack = True
ElseIf optAttacker(2).Value = True Then
    bManualAttack = True
Else
    Exit Sub
End If

If optDefender(0).Value = True Then
    bVSchar = True
ElseIf optDefender(1).Value = True Then
    bVSmob = True
ElseIf optDefender(2).Value = True Then
    bVSplayer = True
ElseIf optDefender(3).Value = True Then
    bVSmanual = True
Else
    Exit Sub
End If

lblSubRight.Caption = "Hit Min-Cap: " & GetHitMin() & "%-" & GetHitCap() & "%"
If bGreaterMUD Then
    lblSubRight.Caption = lblSubRight.Caption & ", Dodge DR-Cap: " & GetDodgeCap(, True) & "-" & GetDodgeCap() & "%"
Else
    lblSubRight.Caption = lblSubRight.Caption & ", Dodge Cap: " & GetDodgeCap() & "%"
End If

bDontRefresh = True
If bBackstab And (bVSplayer Or bVSchar) Then
    txtHitCalc(2).Locked = True 'dodge
    txtHitCalc(3).Locked = True '2nd
    txtHitCalc(2).BackColor = &H8000000F
    txtHitCalc(3).BackColor = &H8000000F
    txtHitCalc(2).Text = "click"
    txtHitCalc(3).Text = "click"
ElseIf bNormal And bVSplayer Then
    txtHitCalc(2).Locked = False 'dodge
    txtHitCalc(3).Locked = True '2nd
    txtHitCalc(2).BackColor = &H80000005
    txtHitCalc(3).BackColor = &H8000000F
    If txtHitCalc(2).Text = "click" Then txtHitCalc(2).Text = 0
    txtHitCalc(3).Text = "click"
Else
    txtHitCalc(2).Locked = False 'dodge
    txtHitCalc(3).Locked = False '2nd
    txtHitCalc(2).BackColor = &H80000005
    txtHitCalc(3).BackColor = &H80000005
    If txtHitCalc(2).Text = "click" Then txtHitCalc(2).Text = 0
    If txtHitCalc(3).Text = "click" Then txtHitCalc(3).Text = 0
End If


'ATTACKER / ACCY
If bBackstab And bCharAttack Then 'bs + current character
    
    If bGlobalAttackBackstab And nGlobalAttackBackstabWeapon > 0 Then
        nBSWep = nGlobalAttackBackstabWeapon
    ElseIf Not bGlobalAttackBackstab Or nGlobalAttackBackstabWeapon = 0 Then
        nBSWep = nGlobalCharWeaponNumber(0)
    End If
    
    'not currently accounting for removal of shield if new bs weapon is two hander...
    If nBSWep <> nGlobalCharWeaponNumber(0) Then
        If nBSWep > 0 Then
            nNormAccyAdj = ItemHasAbility(nBSWep, 22)
            If nNormAccyAdj < 0 Then nNormAccyAdj = 0
            nBSAccyAdj = ItemHasAbility(nBSWep, 116)
            If nBSAccyAdj < 0 Then nBSAccyAdj = 0
        End If
        nNormAccyAdj = nNormAccyAdj - nGlobalCharWeaponAccy(0)
        nBSAccyAdj = nBSAccyAdj - nGlobalCharWeaponBSaccy(0)
    End If
    
    txtHitCalc(0).Text = CalculateBackstabAccuracy(val(frmMain.lblInvenCharStat(19).Tag), val(frmMain.txtCharStats(3).Tag), _
        val(frmMain.lblInvenCharStat(13).Tag) + nBSAccyAdj, _
        GetClassStealth(frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex)), _
        nGlobalCharAccyAbils + nGlobalCharAccyOther + nNormAccyAdj, val(frmMain.txtGlobalLevel(0).Text), val(frmMain.txtCharStats(0).Tag), GetItemStrReq(nBSWep))

ElseIf bNormal And bCharAttack Then 'normal + current character
    
    txtHitCalc(0).Text = val(frmMain.lblInvenCharStat(10).Tag)  'acc
    
    If nGlobalAttackTypeMME = a6_PhysBash Then
        txtHitCalc(0).Text = val(txtHitCalc(0).Text) + nGlobalAttackAccyAdj
        lblSubRight.Caption = AutoAppend(lblSubRight.Caption, "Bash " & nGlobalAttackAccyAdj & " acc", ", ")
    ElseIf nGlobalAttackTypeMME = a7_PhysSmash Then
        txtHitCalc(0).Text = val(txtHitCalc(0).Text) + nGlobalAttackAccyAdj
        lblSubRight.Caption = AutoAppend(lblSubRight.Caption, "Smash " & nGlobalAttackAccyAdj & " acc", ", ")
    ElseIf nGlobalAttackTypeMME = a4_MartialArts Then
        If nGlobalAttackMA = 2 Then 'kick
            txtHitCalc(0).Text = val(txtHitCalc(0).Text) + nGlobalAttackAccyAdj
            lblSubRight.Caption = AutoAppend(lblSubRight.Caption, "Kick " & nGlobalAttackAccyAdj & " acc", ", ")
        ElseIf nGlobalAttackMA = 3 Then 'jk
            txtHitCalc(0).Text = val(txtHitCalc(0).Text) + nGlobalAttackAccyAdj
            lblSubRight.Caption = AutoAppend(lblSubRight.Caption, "Jumpkick " & nGlobalAttackAccyAdj & " acc", ", ")
        End If
    End If

ElseIf bNormal And bMobAttack Then 'normal + monster

    txtHitCalc(0).Text = nMonsterAccy
    
End If

'DEFENDER - AC, DODGE, SECONDARY DEFENSE
If bVSmob Then 'monster
    
    txtHitCalc(1).Text = nMonsterAC
    txtHitCalc(2).Text = nMonsterDodge
    lblDodge.Caption = "vs Dodge:"
    If bBackstab Then 'backstab
        txtHitCalc(3).Text = nMonsterBSdef
        lblSubLeft.Caption = "BS Defense"
    Else
        txtHitCalc(3).Text = 0
        lblSubLeft.Caption = ""
    End If
    
ElseIf bVSchar Or bVSplayer Then 'char/player
    
    If bVSchar Then
        txtHitCalc(1).Text = Fix(val(frmMain.lblInvenCharStat(2).Tag)) 'ac
        
        If bBackstab Then
            If bGreaterMUD Then
            Else
            End If
        Else
            txtHitCalc(2).Text = val(frmMain.lblInvenCharStat(8).Tag) 'dodge
            txtHitCalc(3).Tag = val(frmMain.lblInvenCharStat(20).Tag) + val(frmMain.lblInvenCharStat(32).Tag) 'prot.evil/good
            lblDodge.Caption = "vs Dodge:"
            lblSubLeft.Caption = "Prot. Evil/Good (incl.), Shadow (add +10)"
            If bGreaterMUD Then lblSubLeft.Caption = lblSubLeft.Caption & ", Vile Ward (add [VW val/10] for villian+ -OR- [VW val/2/10] if criminal)"
        End If
                
    ElseIf bVSplayer Then
        
        txtHitCalc(1).Text = nDefenderAC
        If bBackstab Then
            If bGreaterMUD Then
                If (nDefenderDodge + nDefenderPerception + nDefenderProtEvil + nDefenderVileWard) = 0 Then
                    Call PromptForPlayerDodgeBSDefense(nDefenderDodge, nDefenderPerception, bDefenderSeeHidden)
                    Call PromptForPlayerSecondaryDefense(nDefenderProtEvil, nDefenderVileWard, nDefenderEvilness, bDefenderShadow, -1)
                End If
                lblDodge.Caption = "DG+Percep+SeeH:"
                lblSubLeft.Caption = "Percep:" & nDefenderPerception
                lblSubLeft.Caption = "Prot. Evil/Good: " & nDefenderProtEvil
            Else
                txtHitCalc(2).Text = nDefenderDodge
                txtHitCalc(3).Text = nDefenderPerception
                lblDodge.Caption = "vs Dodge:"
                lblSubLeft.Caption = "Perception"
            End If
        Else
            lblDodge.Caption = "vs Dodge:"
            txtHitCalc(2).Text = nDefenderDodge
            
            If (nDefenderPerception + nDefenderProtEvil + nDefenderVileWard) = 0 Then
                Call PromptForPlayerSecondaryDefense(nDefenderProtEvil, nDefenderVileWard, nDefenderEvilness, bDefenderShadow)
            End If
            
            lblSubLeft.Caption = "Prot. Evil/Good: " & nDefenderProtEvil
            If bDefenderShadow Then
                lblSubLeft.Caption = ", Shadow 10"
            Else
                lblSubLeft.Caption = ", Shadow 0"
            End If
            If bGreaterMUD Then
                lblSubLeft.Caption = lblSubLeft.Caption & ", Vile Ward: " & GetVileWardValue(nDefenderVileWard, nDefenderEvilness)
            End If
        End If
    End If
    
    
    
    If bNormal Then 'normal
    
        
        If bVSchar Then
        Else
        End If
        
        lblSubLeft.Caption = "Prot. Evil/Good, Shadow"
        If bGreaterMUD Then lblSubLeft.Caption = lblSubLeft.Caption & ", Vile Ward"
        
        
    ElseIf bBackstab Then 'bs
        
        If bGreaterMUD Then
            lblDodge.Caption = "DG+Percep+SeeH:"
            lblSubLeft.Caption = "Percep:" & nDefenderPerception
        Else
            lblSubLeft.Caption = "Perception"
        End If
    End If
    
    If bNormal Then
        
        
        
        lblSubLeft.Caption = lblSubLeft.Caption & "+PrEV:" & nDefenderProtEvil
        If bDefenderShadow Then lblSubLeft.Caption = lblSubLeft.Caption & ", +Shadow"
        
    End If
        
    If bGreaterMUD Then
        lblSubLeft.Caption = lblSubLeft.Caption & ", VileWard:" & nDefenderVileWard
        If nDefenderVileWard > 0 Then
            If nDefenderEvilness >= 2 Then
                lblSubLeft.Caption = lblSubLeft.Caption & ", VileWard:" & nDefenderVileWard & "(villian)"
            ElseIf nDefenderEvilness >= 1 Then
                lblSubLeft.Caption = lblSubLeft.Caption & ", VileWard:" & nDefenderVileWard & "(criminal)"
            Else
                lblSubLeft.Caption = lblSubLeft.Caption & ", VileWard: 0 (not evil)"
            End If
        End If
        If bDefenderShadow Then lblSubLeft.Caption = lblSubLeft.Caption & ", +Shadow"
    End If
    
ElseIf bVSmanual Then 'manual
    lblDodge.Caption = "vs Dodge:"
    If bBackstab Then
        lblSubLeft.Caption = "Perception"
    Else
        lblSubLeft.Caption = "Prot. Evil/Good, Shadow"
        If bGreaterMUD Then lblSubLeft.Caption = lblSubLeft.Caption & ", Vile Ward"
    End If
End If

out:
On Error Resume Next
bDontRefresh = False
Call DoHitCalc
Exit Sub
error:
Call HandleError("SetHitCalcVals")
Resume out:
End Sub

Private Function GetVileWardValue(ByVal nDefenderVileWard As Long, ByVal nDefenderEvilness As Long) As Long
If nDefenderVileWard > 0 Then
    If nDefenderEvilness >= 2 Then
        GetVileWardValue = nDefenderVileWard
    ElseIf nDefenderEvilness >= 1 Then
        GetVileWardValue = nDefenderVileWard \ 2
    Else
        GetVileWardValue = 0
    End If
End If
End Function

Private Sub cmdCharHitCalc_Click(Index As Integer)
If Not bMouseDown Then Call ModifyHitCalcValues(Index)
timButtonPress.Enabled = False
bMouseDown = False
End Sub

Private Sub cmdCharHitCalc_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

bMouseDown = True
ntimButtonPressCount = 1

Do While bMouseDown And ntimButtonPressCount < 101
    ntimButtonPressCount = ntimButtonPressCount + 1
    timButtonPress.Enabled = True
    Call ModifyHitCalcValues(Index)
    Do While timButtonPress.Enabled
        DoEvents
    Loop
Loop
bMouseDown = False


End Sub

Private Sub ModifyHitCalcValues(ByVal Index As Integer)
On Error GoTo error:

If txtHitCalc(Index).Text = "click" Then Exit Sub

Select Case Index
    Case 0: 'acc -
        txtHitCalc(0).Text = Fix(val(txtHitCalc(0).Text)) - 1
    Case 1: 'acc +
        txtHitCalc(0).Text = Fix(val(txtHitCalc(0).Text)) + 1
    Case 2: 'ac -
        txtHitCalc(1).Text = Fix(val(txtHitCalc(1).Text)) - 1
    Case 3: 'ac +
        txtHitCalc(1).Text = Fix(val(txtHitCalc(1).Text)) + 1
    Case 4: 'dg -
        txtHitCalc(2).Text = Fix(val(txtHitCalc(2).Text)) - 1
    Case 5: 'dg +
        txtHitCalc(2).Text = Fix(val(txtHitCalc(2).Text)) + 1
    Case 6: '2nd -
        txtHitCalc(3).Text = Fix(val(txtHitCalc(3).Text)) - 1
    Case 7: '2nd +
        txtHitCalc(3).Text = Fix(val(txtHitCalc(3).Text)) + 1
End Select

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("ModifyHitCalcValues")
Resume out:
End Sub

Private Sub cmdCharHitCalc_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
bMouseDown = False
End Sub

Private Sub Form_Load()
On Error GoTo error:

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

timWindowMove.Enabled = True

Call LoadMonsters
Call SetHitCalcVals

Exit Sub
error:
Call HandleError("frmHitCalc_Load")
Resume Next
End Sub

Private Sub LoadMonsters()
Dim sName As String
On Error GoTo error:

cmbMonsterList.clear

tabMonsters.MoveFirst
Do While Not tabMonsters.EOF
    If bOnlyInGame And tabMonsters.Fields("In Game") = 0 Then GoTo MoveNext:
    
    sName = tabMonsters.Fields("Name")
    If sName = "" Or Left(sName, 3) = "sdf" Then GoTo MoveNext:
    sName = sName & " (" & tabMonsters.Fields("Number") & ")"
    
    cmbMonsterList.AddItem sName
    cmbMonsterList.ItemData(cmbMonsterList.NewIndex) = tabMonsters.Fields("Number")
    
MoveNext:
    tabMonsters.MoveNext
Loop
tabMonsters.MoveFirst

If cmbMonsterList.ListCount = 0 Then Exit Sub

cmbMonsterList.ListIndex = 0
Call AutoSizeDropDownWidth(cmbMonsterList)
Call ExpandCombo(cmbMonsterList, HeightOnly, DoubleWidth, Me.hWnd)
cmbMonsterList.SelLength = 0
Call cmbMonsterList_Click

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("LoadMonsters")
Resume out:
End Sub

Private Function GetMonsterData(ByVal nMonster As Long) As Boolean
On Error GoTo error:
Dim x As Integer

On Error GoTo seek2:
If tabMonsters.Fields("Number") = nMonster Then GoTo ready:
GoTo seekit:

seek2:
Resume seekit:
seekit:
On Error GoTo error:
tabMonsters.Index = "pkMonsters"
tabMonsters.Seek "=", nMonster
If tabMonsters.NoMatch = True Then
    tabMonsters.MoveFirst
    Exit Function
End If

ready:
nMonsterAccy = 0
nMonsterDodge = 0
nMonsterAC = tabMonsters.Fields("ArmourClass")
If nNMRVer >= 1.83 Then
    nMonsterBSdef = tabMonsters.Fields("BSDefense")
Else
    nMonsterBSdef = 0
End If

For x = 0 To 9 'abilities
    If tabMonsters.Fields("Abil-" & x) = 34 And tabMonsters.Fields("AbilVal-" & x) > 0 Then 'dodge
        nMonsterDodge = tabMonsters.Fields("AbilVal-" & x)
        Exit For
    End If
Next x

For x = 0 To 4 'attacks
    If tabMonsters.Fields("AttType-" & x) > 0 And tabMonsters.Fields("AttType-" & x) <= 3 And tabMonsters.Fields("Att%-" & x) > 0 Then
        Select Case tabMonsters.Fields("AttType-" & x)
            Case 1, 3  ' melee (normal/rob)
                If tabMonsters.Fields("AttAcc-" & x) > nMonsterAccy Then nMonsterAccy = tabMonsters.Fields("AttAcc-" & x)
                
        End Select
    End If
Next x

GetMonsterData = True

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetMonsterData")
Resume out:
End Function


Private Sub optAttacker_Click(Index As Integer)
Dim x As Integer
For x = 0 To 2
    If optAttacker(x).Value = True Then
        optAttacker(x).FontBold = True
    Else
        optAttacker(x).FontBold = False
    End If
Next x
Call SetHitCalcVals
Call DoHitCalc
End Sub

Private Sub optDefender_Click(Index As Integer)
Dim x As Integer
For x = 0 To 3
    If optDefender(x).Value = True Then
        optDefender(x).FontBold = True
    Else
        optDefender(x).FontBold = False
    End If
Next x
Call SetHitCalcVals
Call DoHitCalc
End Sub

Private Sub optHitCalcType_Click(Index As Integer)
Dim x As Integer

For x = 0 To 1
    If optHitCalcType(x).Value = True Then
        optHitCalcType(x).FontBold = True
    Else
        optHitCalcType(x).FontBold = False
    End If
Next x

If optHitCalcType(1).Value = True Then 'if backstab
    If optAttacker(1).Value = True Then optAttacker(2).Value = True 'no monster attacker
    optAttacker(1).Enabled = False
Else
    optAttacker(1).Enabled = True
End If

Call SetHitCalcVals
Call DoHitCalc

End Sub

Private Sub timWindowMove_Timer()
Call MonitorFormTimer(Me)
End Sub


Private Sub timButtonPress_Timer()
If bAppReallyTerminating Or bAppTerminating Then Exit Sub
timButtonPress.Enabled = False
End Sub

Private Sub txtHitCalc_Change(Index As Integer)
Call DoHitCalc
End Sub

Private Sub txtHitCalc_GotFocus(Index As Integer)
Call SelectAll(txtHitCalc(Index))
End Sub

Private Sub txtHitCalc_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii, False)
End Sub
