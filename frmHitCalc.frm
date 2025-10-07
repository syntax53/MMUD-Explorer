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
         Left            =   2400
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
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   8
         Text            =   "0"
         Top             =   3720
         Width           =   795
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
            Caption         =   "Manual Player"
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
            Width           =   2010
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
         Height          =   555
         Left            =   240
         TabIndex        =   35
         Top             =   5340
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
      Begin VB.Label lblLabelArray 
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
         Index           =   63
         Left            =   300
         TabIndex        =   33
         Top             =   4260
         Width           =   1710
      End
      Begin VB.Label lblLabelArray 
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
         Index           =   62
         Left            =   4635
         TabIndex        =   32
         Top             =   3420
         Width           =   1080
      End
      Begin VB.Label lblLabelArray 
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
         Index           =   34
         Left            =   630
         TabIndex        =   31
         Top             =   3180
         Width           =   1050
      End
      Begin VB.Label lblLabelArray 
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
         Index           =   32
         Left            =   2850
         TabIndex        =   30
         Top             =   3420
         Width           =   690
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
         Left            =   2400
         TabIndex        =   27
         Top             =   4320
         Width           =   3675
      End
      Begin VB.Label lblSubRight 
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
         Height          =   555
         Left            =   2400
         TabIndex        =   26
         Top             =   5340
         Width           =   3675
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

Dim nPlayer
Dim bDontRefresh As Boolean
Dim bMouseDown As Boolean
Dim tWindowSize As WindowSizeProperties

Public Sub DoHitCalc()
On Error GoTo error:
Dim nAC As Long, nAccy As Long, nHitChance As Currency, nTotalHitPercent As Currency
Dim nDodge As Long, nDodgeChance As Currency, sArr() As String
Dim sPrint As String, nTemp As Long, nAux As Long, nShadow As Long
Dim nSecondaryDef As Long, nProtEv As Long, nPerception As Long, nVileWard As Long, eEvil As eEvilPoints
Dim bShadow As Boolean, bSeeHidden As Boolean, bBackstab As Boolean, bVsPlayer As Boolean
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
If optDefender(2).Value = True Then bVsPlayer = True

If optDefender(0).Value = 1 Then
    nClass = frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex)
ElseIf optDefender(2).Value = 1 Then
    nClass = nDefenderClass
Else
    nClass = -1
End If

'GET HIT CHANCE
If nAC + nAux > 0 Or (bBackstab And bVsPlayer And bGreaterMUD And Len(txtHitCalc(3).Tag) > 0) Then
    If bBackstab Then '[BACKSTAB]
        If bVsPlayer Then '[BACKSTAB+PLAYER]
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
    If (nDodge > 0 Or (bBackstab And bVsPlayer And Len(txtHitCalc(2).Tag) > 0)) Then
        If bBackstab And bVsPlayer And Len(txtHitCalc(2).Tag) > 0 Then
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
    nVileWard, eEvil, bShadow, bSeeHidden, bBackstab, bVsPlayer, nClass)

nHitChance = nDefense(0)
nDodgeChance = nDefense(1)

If bBackstab And bVsPlayer Then
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

bDontRefresh = True
lblMain.Caption = ""
lblSubLeft.Caption = ""
lblSubRight.Caption = ""

'ATTACKER - ACC
If optHitCalcType(1).Value = True And optAttacker(0).Value = True Then 'bs + current character
    
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

ElseIf optHitCalcType(0).Value = True And optAttacker(0).Value = True Then 'normal + current character
    
    txtHitCalc(0).Text = val(frmMain.lblInvenCharStat(10).Tag) 'acc
    
    If nGlobalAttackTypeMME = a6_PhysBash Then
        txtHitCalc(0).Text = val(txtHitCalc(0).Text) - 15
        lblSubRight.Caption = AutoAppend(lblSubRight.Caption, "Bash -15 acc", ", ")
    ElseIf nGlobalAttackTypeMME = a7_PhysSmash Then
        txtHitCalc(0).Text = val(txtHitCalc(0).Text) - 25
        lblSubRight.Caption = AutoAppend(lblSubRight.Caption, "Smash -25 acc", ", ")
    ElseIf nGlobalAttackTypeMME = a4_MartialArts Then
        If nGlobalAttackMA = 1 Then 'kick
            txtHitCalc(0).Text = val(txtHitCalc(0).Text) - 10
            lblSubRight.Caption = AutoAppend(lblSubRight.Caption, "Kick -10 acc", ", ")
        ElseIf nGlobalAttackMA = 2 Then 'jk
            txtHitCalc(0).Text = val(txtHitCalc(0).Text) - 15
            lblSubRight.Caption = AutoAppend(lblSubRight.Caption, "Jumpkick -15 acc", ", ")
        End If
    End If

ElseIf optHitCalcType(0).Value = True And optAttacker(1).Value = True Then 'normal + monster

    txtHitCalc(0).Text = nMonsterAccy
    
End If

'DEFENDER - AC, DODGE, SECONDARY DEFENSE
If optDefender(0).Value = True Then 'current char
    txtHitCalc(1).Text = Fix(val(frmMain.lblInvenCharStat(2).Tag)) 'ac
    txtHitCalc(2).Text = val(frmMain.lblInvenCharStat(8).Tag) 'dodge
    txtHitCalc(3).Text = val(frmMain.lblInvenCharStat(20).Tag) + val(frmMain.lblInvenCharStat(32).Tag) 'prot.evil/good
    
ElseIf optDefender(1).Value = True Then 'monster
    txtHitCalc(1).Text = nMonsterAC
    txtHitCalc(2).Text = nMonsterDodge
    If optAttacker(0).Value = True Then 'backstab
        txtHitCalc(3).Text = nMonsterBSdef
    Else
        txtHitCalc(3).Text = 0
    End If
    
ElseIf optDefender(2).Value = True Then 'player
    
    txtHitCalc(1).Text = nDefenderAC
    txtHitCalc(2).Text = nDefenderDodge
    txtHitCalc(3).Text = val(frmMain.lblInvenCharStat(20).Tag) + val(frmMain.lblInvenCharStat(32).Tag) 'prot.evil/good
    
    
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

Private Sub cmdCharHitCalc_Click(Index As Integer)
If Not bMouseDown Then Call ModifyHitCalcValues(Index)
timButtonPress.Enabled = False
bMouseDown = False
End Sub

Private Sub cmdCharHitCalc_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

bMouseDown = True
Do While bMouseDown
    timButtonPress.Enabled = True
    Call ModifyHitCalcValues(Index)
    Do While timButtonPress.Enabled
        DoEvents
    Loop
Loop

End Sub

Private Sub ModifyHitCalcValues(ByVal Index As Integer)
On Error GoTo error:

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
    If optAttacker(x).Value = True Then
        optAttacker(x).FontBold = True
    Else
        optAttacker(x).FontBold = False
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
