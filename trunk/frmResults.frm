VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20D5284F-7B23-4F0A-B8B1-6C9D18B64F1C}#1.0#0"; "exlimiter.ocx"
Begin VB.Form frmResults 
   Caption         =   "Results (click to jump)"
   ClientHeight    =   2595
   ClientLeft      =   660
   ClientTop       =   945
   ClientWidth     =   5145
   Icon            =   "frmResults.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   5145
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4620
      Top             =   1920
   End
   Begin VB.Frame fraTree 
      BorderStyle     =   0  'None
      Height          =   2595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4395
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Next"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3420
         TabIndex        =   10
         Top             =   0
         Width           =   555
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   9
         Top             =   0
         Width           =   555
      End
      Begin VB.CommandButton cmdQ 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   8
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdCollapse 
         Caption         =   "&Expand"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   900
         TabIndex        =   7
         Top             =   0
         Width           =   855
      End
      Begin VB.CheckBox chkHideTextblocks 
         Caption         =   "Simple"
         Height          =   195
         Left            =   1920
         TabIndex        =   6
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdCollapse 
         Caption         =   "&Collapse"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   855
      End
      Begin MSComctlLib.TreeView tvwResults 
         Height          =   2295
         Left            =   0
         TabIndex        =   1
         Top             =   300
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   4048
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin exlimiter.EL EL1 
      Left            =   4440
      Top             =   900
      _ExtentX        =   1270
      _ExtentY        =   1270
   End
   Begin MSComctlLib.ImageList ilImages 
      Left            =   4500
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame fraLV 
      BorderStyle     =   0  'None
      Height          =   2595
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4395
      Begin MSComctlLib.ListView lvResults 
         Height          =   2355
         Left            =   0
         TabIndex        =   3
         Top             =   270
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   4154
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblCaption 
         Height          =   255
         Left            =   45
         TabIndex        =   4
         Top             =   0
         Width           =   4275
      End
   End
   Begin VB.Menu mnuExpand 
      Caption         =   "ExpandCollapse"
      Visible         =   0   'False
      Begin VB.Menu mnuExpandItem 
         Caption         =   "Collapse Branch"
         Index           =   0
      End
      Begin VB.Menu mnuExpandItem 
         Caption         =   "Expand Branch"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' NOTE: this is the ugliest code i think i've ever created
' -syntax

Option Explicit
Option Base 0

'ntreemode = 1 == execution tree
'ntreemode = 2 == regualar textblock line
'ntreemode = 3 == room command
'ntreemode = 4 == greet text

Public objFormOwner As Form
Public nDefaultMap As Long
Dim nLastNode As Long
Dim sFind As String
Dim nTreeMode As Integer
Dim nNest As Integer
Dim nNestMax As Integer
Dim ScannedTB() As Boolean
Dim nWindowState As Integer

Private Sub Form_Load()
Dim nTmp As Long
On Error GoTo Error:

With EL1
    .CenterOnLoad = True
    .FormInQuestion = Me
    .MinWidth = 300
    .MinHeight = 200 + (TITLEBAR_OFFSET / 10)
    .EnableLimiter = True
End With

lvResults.ColumnHeaders.clear
lvResults.ColumnHeaders.Add 1, "Location", "Location/Execution Matches", 3500

chkHideTextblocks.Value = ReadINI("Settings", "HideTextblockResults")
'
'nTmp = ReadINI("Settings", "ResultsTop")
'Me.Top = IIf(nTmp > 1, nTmp, frmMain.Top)
'
'nTmp = ReadINI("Settings", "ResultsLeft")
'Me.Left = IIf(nTmp > 1, nTmp, frmMain.Left)
'
'nTmp = ReadINI("Settings", "ResultsWidth")
'Me.Width = IIf(nTmp > 4335, nTmp, 4335)
'
'nTmp = ReadINI("Settings", "ResultsHeight")
'Me.Height = IIf(nTmp > 3465, nTmp, 3465)

Exit Sub

Error:
Call HandleError("Form_Load")

End Sub

Private Sub chkHideTextblocks_Click()

On Error GoTo Error:

If tvwResults.Nodes.Count < 1 Then Exit Sub

If nTreeMode = 1 Then 'ntreemode = 1 == execution tree
    Call CreateExecutionTree(Val(tvwResults.Nodes(1).Tag))
ElseIf nTreeMode = 2 Then 'ntreemode = 2 == regualar textblock line
    Call CreateCommandTree(Val(tvwResults.Nodes(1).Tag), False, False)
ElseIf nTreeMode = 3 Then 'ntreemode = 3 == room command
    Call CreateCommandTree(Val(tvwResults.Nodes(1).Tag), True, False)
ElseIf nTreeMode = 4 Then 'ntreemode = 4 == greet text
    Call CreateCommandTree(Val(tvwResults.Nodes(1).Tag), False, True)
End If

Exit Sub

Error:
Call HandleError("chkHideTextblocks_Click")

End Sub

Private Sub cmdCollapse_Click(Index As Integer)
Dim x As Integer, bExpanded As Boolean
On Error GoTo Error:

If tvwResults.Nodes.Count < 1 Then Exit Sub

Me.MousePointer = vbHourglass
DoEvents
Call LockWindowUpdate(Me.hWnd)
If Index = 1 Then bExpanded = True

For x = 1 To tvwResults.Nodes.Count
    tvwResults.Nodes(x).Expanded = bExpanded
Next x

tvwResults.Nodes(1).Expanded = True

out:
Me.MousePointer = vbDefault
Call LockWindowUpdate(0&)
Exit Sub

Error:
Call HandleError("cmdCollapse_Click")
Resume out:
End Sub

Private Sub cmdFind_Click(Index As Integer)
Dim sTemp As String, nStartNode As Long, x As Long

If tvwResults.Nodes.Count < 1 Then Exit Sub

If Index = 0 Or sFind = "" Then
    sTemp = InputBox("Enter text to search for.", "Search for text", sFind)
    If sTemp = "" Then Exit Sub
    sFind = sTemp
    nStartNode = 1
Else
    nStartNode = tvwResults.SelectedItem.Index + 1
End If

For x = nStartNode To tvwResults.Nodes.Count
    If InStr(1, LCase(tvwResults.Nodes(x)), LCase(sFind)) > 0 Then
        tvwResults.SelectedItem = tvwResults.Nodes(x)
        Exit For
    End If
Next x

If x = tvwResults.Nodes.Count + 1 Then
    MsgBox "Not Found.", vbInformation
End If

End Sub

Private Sub cmdQ_Click()
    MsgBox "Clicking on ""[Link To] Textblock X"" lines will copy that textblock's raw data to your clipboard.", vbInformation
End Sub

Public Sub SetupResultsWindow(ByVal bTreeMode As Boolean, ByRef objSetFormOwner As Form, _
    Optional ByVal nSetDefaultMap As Long)
Dim lR As Long

On Error GoTo Error:

If FormIsLoaded("frmResults") Then Unload Me

Load Me
DoEvents
Set objFormOwner = objSetFormOwner

'If Not bNoOnTopofOwner Or objFormOwner Is frmMap Then
'    Call SetOwner(Me.hwnd, objFormOwner.hwnd)
''    If objFormOwner Is frmMap Then
''        If frmMap.chkMapOptions(6).Value = 0 Then Call SetTopMostWindow(frmMap.hwnd, True)
''    End If
'End If

If nSetDefaultMap > 0 Then nDefaultMap = nSetDefaultMap

If Me.WindowState = vbMinimized Then
    Me.WindowState = vbNormal
End If

If Not Me.WindowState = vbMaximized Then
    If Me.Visible Then
        If bTreeMode And fraTree.Visible = False Then
            Call WriteINI("Settings", "ResultsTop", Me.Top)
            Call WriteINI("Settings", "ResultsLeft", Me.Left)
            Call WriteINI("Settings", "ResultsWidth", Me.Width)
            Call WriteINI("Settings", "ResultsHeight", Me.Height)
            Me.Top = ReadINI("Settings", "ResultsTreeTop")
            Me.Left = ReadINI("Settings", "ResultsTreeLeft")
            Me.Width = ReadINI("Settings", "ResultsTreeWidth")
            Me.Height = ReadINI("Settings", "ResultsTreeHeight")
        ElseIf Not bTreeMode And fraLV.Visible = False Then
            Call WriteINI("Settings", "ResultsTreeTop", Me.Top)
            Call WriteINI("Settings", "ResultsTreeLeft", Me.Left)
            Call WriteINI("Settings", "ResultsTreeWidth", Me.Width)
            Call WriteINI("Settings", "ResultsTreeHeight", Me.Height)
            Me.Top = ReadINI("Settings", "ResultsTop")
            Me.Left = ReadINI("Settings", "ResultsLeft")
            Me.Width = ReadINI("Settings", "ResultsWidth")
            Me.Height = ReadINI("Settings", "ResultsHeight")
        End If
    Else
        If bTreeMode Then
            Me.Top = ReadINI("Settings", "ResultsTreeTop")
            Me.Left = ReadINI("Settings", "ResultsTreeLeft")
            Me.Width = ReadINI("Settings", "ResultsTreeWidth")
            Me.Height = ReadINI("Settings", "ResultsTreeHeight")
        Else
            Me.Top = ReadINI("Settings", "ResultsTop")
            Me.Left = ReadINI("Settings", "ResultsLeft")
            Me.Width = ReadINI("Settings", "ResultsWidth")
            Me.Height = ReadINI("Settings", "ResultsHeight")
        End If
    End If
End If

If bTreeMode Then
    fraTree.Visible = True
    fraLV.Visible = False
Else
    fraTree.Visible = False
    fraLV.Visible = True
End If

out:
Exit Sub
Error:
Call HandleError("SetupResultsWindow")
Resume out:

End Sub

Private Sub Form_Resize()
On Error Resume Next

If Me.WindowState = vbMinimized Then Exit Sub

nWindowState = Me.WindowState

lvResults.Width = Me.Width - 100
lvResults.Height = Me.Height - TITLEBAR_OFFSET - 650
lvResults.ColumnHeaders(1).Width = lvResults.Width - 500
fraLV.Width = Me.Width - 130
fraLV.Height = Me.Height + TITLEBAR_OFFSET

tvwResults.Width = Me.Width - 130
tvwResults.Height = Me.Height - TITLEBAR_OFFSET - 700
fraTree.Width = Me.Width - 130
fraTree.Height = Me.Height + TITLEBAR_OFFSET
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Call WriteINI("Settings", "HideTextblockResults", chkHideTextblocks.Value)
If Not Me.WindowState = vbMinimized And Not Me.WindowState = vbMaximized Then
    If fraTree.Visible Then
        Call WriteINI("Settings", "ResultsTreeTop", Me.Top)
        Call WriteINI("Settings", "ResultsTreeLeft", Me.Left)
        Call WriteINI("Settings", "ResultsTreeWidth", Me.Width)
        Call WriteINI("Settings", "ResultsTreeHeight", Me.Height)
    Else
        Call WriteINI("Settings", "ResultsTop", Me.Top)
        Call WriteINI("Settings", "ResultsLeft", Me.Left)
        Call WriteINI("Settings", "ResultsWidth", Me.Width)
        Call WriteINI("Settings", "ResultsHeight", Me.Height)
    End If
End If
If Not objFormOwner Is Nothing Then
    If Not bAppTerminating Then objFormOwner.SetFocus
    Set objFormOwner = Nothing
End If
End Sub

Private Sub lvResults_Click()
On Error GoTo Error:

If objFormOwner Is Nothing Then Set objFormOwner = frmMain

If lvResults.SelectedItem Is Nothing Then Exit Sub
frmMain.bDontSetMainFocus = True
Call frmMain.GotoLocation(lvResults.SelectedItem, , objFormOwner)
'objFormOwner.SetFocus
frmMain.bDontSetMainFocus = False

Exit Sub

Error:
Call HandleError("lvResults_Click")
End Sub

Private Sub lvResults_KeyUp(KeyCode As Integer, Shift As Integer)
Call lvResults_Click
End Sub

Public Sub CreateExecutionTree(ByVal nTextblockNumber As Long)
On Error GoTo Error:
Dim nStatus As Integer, Line As String
Dim NodX As Node, i As Integer, CurrentSubTree As Integer, imgX As ListImage
Dim WorkingTree As Integer, CurrentTree As Integer, sLoc As String

If tabTBInfo.RecordCount = 0 Then Exit Sub

tabTBInfo.Index = "pkTBInfo"
tabTBInfo.Seek "=", nTextblockNumber
If tabTBInfo.NoMatch Then
    MsgBox "Textblock " & nTextblockNumber & " not found."
    GoTo out:
End If

chkHideTextblocks.Enabled = True
nTreeMode = 1

If Me.Visible Then
    Me.MousePointer = vbHourglass
    DoEvents
    Call LockWindowUpdate(Me.hWnd)
End If
tvwResults.Nodes.clear
nLastNode = 0
ilImages.ListImages.clear
Set imgX = ilImages.ListImages.Add(1, "STAR", LoadResPicture("STAR", vbResBitmap))
Set imgX = ilImages.ListImages.Add(2, "ARROW", LoadResPicture("ARROW", vbResBitmap))
Set imgX = ilImages.ListImages.Add(3, "PAPER", LoadResPicture("PAPER", vbResBitmap))
Set imgX = ilImages.ListImages.Add(4, "MONSTER", LoadResPicture("MONSTER", vbResBitmap))
Set imgX = ilImages.ListImages.Add(5, "ROOM", LoadResPicture("ROOM", vbResBitmap))
Set imgX = ilImages.ListImages.Add(6, "BOLT", LoadResPicture("BOLT", vbResBitmap))
Set imgX = ilImages.ListImages.Add(7, "MONEY", LoadResPicture("MONEY", vbResBitmap))
Set imgX = ilImages.ListImages.Add(8, "ITEM", LoadResPicture("ITEM", vbResBitmap))
ilImages.MaskColor = vbMagenta
tvwResults.ImageList = ilImages

Erase ScannedTB()
ReDim ScannedTB(1)
Set NodX = tvwResults.Nodes.Add(, , "NODE1", "Textblock " & nTextblockNumber & " is executed by ...", 3)
NodX.Expanded = True
NodX.Tag = nTextblockNumber

nNest = 0
nNestMax = 50
Call AddExecutionNode(nTextblockNumber, 1)

out:
On Error Resume Next
Me.MousePointer = vbDefault
Call LockWindowUpdate(0&)
DoEvents
timWait.Enabled = True
Set NodX = Nothing
Set imgX = Nothing

Exit Sub
Error:
Call HandleError("CreateExecutionTree")
Resume out:
End Sub

Private Sub AddExecutionNode(ByVal nTextblockNumber As Long, ByVal nCurrentNode As Integer)
Dim sLook As String, sChar As String, sTest As String, sSuffix As String
Dim x As Integer, y1 As Integer, y2 As Integer, z As Integer, nValue As Long, x2 As Integer
Dim NodX As Node, nodY As Node
Dim sLoc As String, sTemp As String

On Error GoTo Error:

If chkHideTextblocks.Value = 1 Then
    If nTextblockNumber > UBound(ScannedTB()) Then
        ReDim Preserve ScannedTB(nTextblockNumber)
    End If
    If ScannedTB(nTextblockNumber) = True Then GoTo out:
    ScannedTB(nTextblockNumber) = True
Else
    If nCurrentNode > 1 Then
'        If tvwResults.Nodes(nCurrentNode).Parent.Key = "NODE1" Then
'            Erase ScannedTB()
'            ReDim ScannedTB(1)
'        End If
'        If nTextblockNumber > UBound(ScannedTB()) Then
'            ReDim Preserve ScannedTB(nTextblockNumber)
'        End If
'        If ScannedTB(nTextblockNumber) = True Then
'            Set nodX = tvwResults.Nodes.Add("NODE" & nCurrentNode, tvwChild, _
'                "NODE" & tvwResults.Nodes.Count + 1, "Loop detected (Reference to Textblock " & nTextblockNumber & ", but already scanned in tree.)", 1)
'            nodX.Tag = 0
'            nodX.Expanded = True
'            GoTo out:
'        End If
'        ScannedTB(nTextblockNumber) = True
        Set nodY = tvwResults.Nodes(nCurrentNode)
        If nCurrentNode > 1 Then
            Do
                If Not nodY.Parent.Key = "NODE1" Then
                    Set nodY = nodY.Parent
                    If InStr(1, nodY.Text, "Textblock " & nTextblockNumber) > 0 Then
                        Set NodX = tvwResults.Nodes.Add("NODE" & nCurrentNode, tvwChild, _
                            "NODE" & tvwResults.Nodes.Count + 1, "Loop detected (" & "Textblock " & nTextblockNumber & ") ... quitting.", 1)
                        NodX.Tag = 0
                        NodX.Expanded = True
                        GoTo out:
                    End If
                Else
                    Exit Do
                End If
            Loop
        End If
    End If
End If

nNest = nNest + 1
If nNest > nNestMax Then
    If nNestMax + 1 = nNest Then
        z = MsgBox("MMUD Explorer has nested through " & nNestMax & " textblocks so far, continue for another 50 blocks?", vbYesNo + vbDefaultButton1)
        If z = vbYes Then
            nNestMax = nNestMax + 50
        Else
            Set NodX = tvwResults.Nodes.Add("NODE" & 1, tvwChild, _
                "NODE" & tvwResults.Nodes.Count + 1, "Too many references ... quitting.", 1)
            NodX.Tag = 0
            NodX.Expanded = True
            Exit Sub
        End If
    Else
        Exit Sub
    End If
End If

tabTBInfo.Index = "pkTBInfo"
tabTBInfo.Seek "=", nTextblockNumber
If tabTBInfo.NoMatch Then
    Exit Sub
End If

sLoc = tabTBInfo.Fields("Called From")
If Len(sLoc) < 5 Then Exit Sub
sTest = LCase(sLoc)

For z = 1 To 6
    
    x = 1
    Select Case z
        Case 1: sLook = "room "
        Case 2: sLook = "monster #"
        Case 3: sLook = "textblock #"
        Case 4: sLook = "textblock(rndm) #"
        Case 5: sLook = "item #"
        Case 6: sLook = "spell #"
    End Select

checknext:
    sSuffix = ""
    If Not InStr(x, sTest, sLook) = 0 Then
        
        x = InStr(x, sTest, sLook) 'sets x to the position of the string we're looking for
        
'        If z = 10 Then
'            y1 = x + 1
'            GoTo nonumber:
'        End If
        
        y1 = x + Len(sLook) 'len of string searching (to position y1 at first number)
        y2 = 0
nextnumber:
        sChar = Mid(sTest, y1 + y2, 1)
        Select Case sChar
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "/":
                If Not y1 + y2 - 1 = Len(sTest) Then
                    y2 = y2 + 1
                    GoTo nextnumber:
                End If
            Case "+": 'end of string
                Exit Sub
            Case "(": 'precent
                x2 = InStr(y1 + y2, sTest, ")")
                If Not x2 = 0 Then
                    sSuffix = " " & Mid(sTest, y1 + y2, x2 - y1 - y2 + 1)
                End If
            Case Else:
        End Select
        
        If y2 = 0 Then
            'if there were no numbers after the string
            x = y1
            GoTo checknext:
        End If
        
        If Not z = 1 Or z = 10 Then 'not room or group
            nValue = Val(Mid(sTest, y1, y2))
        End If

nonumber:
        Select Case z
            Case 1: '"room "
                sTemp = GetRoomName(Mid(sTest, y1, y2), , , False)
                If chkHideTextblocks.Value = 1 Then
                    For x = 1 To tvwResults.Nodes.Count
                        If InStr(1, tvwResults.Nodes(x).Text, "Room: " & sTemp) > 0 Then
                            GoTo out:
                        End If
                    Next x
                End If
                
                Set NodX = tvwResults.Nodes.Add("NODE" & nCurrentNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, "Room: " & sTemp & sSuffix, 5)
                NodX.Tag = Mid(sTest, y1, y2)
                NodX.Expanded = True
                NodX.Bold = True
            Case 2: '"monster #"
                Set NodX = tvwResults.Nodes.Add("NODE" & nCurrentNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, "Monster: " & GetMonsterName(nValue, False) & sSuffix, 4)
                NodX.Tag = nValue
                NodX.Expanded = True
                NodX.Bold = True
            Case 3, 4: '"textblock #"
                If chkHideTextblocks.Value = 0 Then
'                    Set nodY = tvwResults.Nodes(nCurrentNode)
'                    If nCurrentNode > 1 Then
'                        Do
'                            If Not nodY.Parent.Key = "NODE1" Then
'                                Set nodY = nodY.Parent
'                                If InStr(1, nodY.Text, "Textblock " & nValue) > 0 Then
'                                    Set nodX = tvwResults.Nodes.Add("NODE" & nCurrentNode, tvwChild, _
'                                        "NODE" & tvwResults.Nodes.Count + 1, "Loop detected (" & "Textblock " & nValue & ") ... quitting.", 1)
'                                    nodX.Tag = 0
'                                    nodX.Expanded = True
'                                    GoTo out:
'                                End If
'                            Else
'                                Exit Do
'                            End If
'                        Loop
'                    End If
                    
                    sTemp = GetTextblockTrigger(nValue, nTextblockNumber)
                    Set NodX = tvwResults.Nodes.Add("NODE" & nCurrentNode, tvwChild, _
                        "NODE" & tvwResults.Nodes.Count + 1, "Textblock " & nValue & sSuffix _
                        & IIf(sTemp = "", "", " " & sTemp), 3)
                    NodX.Tag = nValue
                    NodX.Expanded = True
                    Call AddExecutionNode(nValue, tvwResults.Nodes.Count)
                Else
                    Call AddExecutionNode(nValue, 1)
                End If
                
'            Case 4: '"textblock(rndm) #"
'                Set nodX = tvwResults.Nodes.Add("NODE" & nCurrentNode, tvwChild, _
'                    "NODE" & tvwResults.Nodes.Count + 1, "Textblock " & nValue & sSuffix & " (random)", 3)
'                nodX.Tag = nValue
'                nodX.Expanded = True
'                Call AddExecutionNode(nValue, tvwResults.Nodes.Count)
            Case 5: '"item #"
                Set NodX = tvwResults.Nodes.Add("NODE" & nCurrentNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, "Item: " & GetItemName(nValue, False) & sSuffix, "ARROW")
                NodX.Tag = nValue
                NodX.Expanded = True
                NodX.Bold = True
            Case 6: '"spell #"
                Set NodX = tvwResults.Nodes.Add("NODE" & nCurrentNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, "Spell: " & GetSpellName(nValue, False) & sSuffix, "ARROW")
                NodX.Tag = nValue
                NodX.Expanded = True
                NodX.Bold = True
        End Select
        
        x = y1
        GoTo checknext:
    End If
Next z

out:
On Error Resume Next
Set NodX = Nothing
Set nodY = Nothing
Exit Sub
Error:
Call HandleError("AddNode")
Resume out:
End Sub

Public Sub CreateCommandTree(ByVal nTextblockNumber As Long, _
    ByVal bRoomCommands As Boolean, ByVal bGreetText As Boolean)
On Error GoTo Error:
Dim nStatus As Integer, Line As String
Dim NodX As Node, i As Integer, CurrentSubTree As Integer, imgX As ListImage
Dim WorkingTree As Integer, CurrentTree As Integer, sLoc As String

If tabTBInfo.RecordCount = 0 Then Exit Sub

tabTBInfo.Index = "pkTBInfo"
tabTBInfo.Seek "=", nTextblockNumber
If tabTBInfo.NoMatch Then
    MsgBox "Textblock " & nTextblockNumber & " not found."
    GoTo out:
End If

chkHideTextblocks.Enabled = True
If bRoomCommands Then
    nTreeMode = 3
ElseIf bGreetText Then
    nTreeMode = 4
Else
    nTreeMode = 2
End If

If Me.Visible Then
    Me.MousePointer = vbHourglass
    DoEvents
    Call LockWindowUpdate(Me.hWnd)
End If

tvwResults.Nodes.clear
nLastNode = 0

ilImages.ListImages.clear
Set imgX = ilImages.ListImages.Add(1, "STAR", LoadResPicture("STAR", vbResBitmap))
Set imgX = ilImages.ListImages.Add(2, "ARROW", LoadResPicture("ARROW", vbResBitmap))
Set imgX = ilImages.ListImages.Add(3, "PAPER", LoadResPicture("PAPER", vbResBitmap))
Set imgX = ilImages.ListImages.Add(4, "MONSTER", LoadResPicture("MONSTER", vbResBitmap))
Set imgX = ilImages.ListImages.Add(5, "ROOM", LoadResPicture("ROOM", vbResBitmap))
Set imgX = ilImages.ListImages.Add(6, "BOLT", LoadResPicture("BOLT", vbResBitmap))
Set imgX = ilImages.ListImages.Add(7, "MONEY", LoadResPicture("MONEY", vbResBitmap))
Set imgX = ilImages.ListImages.Add(8, "ITEM", LoadResPicture("ITEM", vbResBitmap))
Set imgX = ilImages.ListImages.Add(9, "REDARROW", LoadResPicture("REDARROW", vbResBitmap))
ilImages.MaskColor = vbMagenta
tvwResults.ImageList = ilImages

Erase ScannedTB()
ReDim ScannedTB(1)
Set NodX = tvwResults.Nodes.Add(, , "NODE1", "Textblock " & nTextblockNumber & " has the following commands ...", 3)
NodX.Expanded = True
NodX.Tag = nTextblockNumber

nNest = 0
nNestMax = 50
Call AddCommandNode(nTextblockNumber, 1, bRoomCommands, False, bGreetText)

'MsgBox tvwResults.Nodes.Count
out:
On Error Resume Next
Me.MousePointer = vbDefault
Call LockWindowUpdate(0&)
DoEvents
timWait.Enabled = True
Set NodX = Nothing
Set imgX = Nothing

Exit Sub
Error:
Call HandleError("CreateTree")
Resume out:
End Sub

Private Sub AddCommandNode(ByVal nTextblockNumber As Long, ByVal nCurrentNode As Integer, _
    ByVal bRoomCommands As Boolean, ByVal bRandom As Boolean, ByVal bGreetText As Boolean)
Dim sChar As String
Dim x As Integer, y As Integer, y2 As Integer, z As Integer, nValue As Long, x2 As Integer
Dim NodX As Node, nodY As Node, sTextblockData As String, sLine As String, nNode As Long
Dim sCommand As String, nDataPos As Integer, nLinePos As Integer, nTotalLines As Integer, nCurrLine As Integer
Dim sLineCommand As String, nMap As Long, nRoom As Long, nPercent1 As Long, nPercent2 As Long
Dim nRepeats As Long, sLastCommand As String, nRepeatNode As Long
On Error GoTo Error:

If nCurrentNode > 1 Then
    Set nodY = tvwResults.Nodes(nCurrentNode)
    If nCurrentNode > 1 Then
        Do
            If Not nodY.Parent.Key = "NODE1" Then
                Set nodY = nodY.Parent
                If InStr(1, nodY.Text, "Textblock " & nTextblockNumber) > 0 Then
                    Set NodX = tvwResults.Nodes.Add("NODE" & nCurrentNode, tvwChild, _
                        "NODE" & tvwResults.Nodes.Count + 1, "Loop detected (" & "Textblock " & nTextblockNumber & ") ... quitting.", 1)
                    NodX.Tag = 0
                    NodX.Expanded = True
                    GoTo out:
                ElseIf InStr(1, nodY.Text, "random: " & nTextblockNumber) > 0 Then
                    Set NodX = tvwResults.Nodes.Add("NODE" & nCurrentNode, tvwChild, _
                        "NODE" & tvwResults.Nodes.Count + 1, "Loop detected (" & "random: " & nTextblockNumber & ") ... quitting.", 1)
                    NodX.Tag = 0
                    NodX.Expanded = True
                    GoTo out:
                End If
            Else
                Exit Do
            End If
        Loop
    End If
End If

nNest = nNest + 1
If nNest > nNestMax Then
    If nNestMax + 1 = nNest Then
        z = MsgBox("MMUD Explorer has nested through " & nNestMax & " textblocks so far, continue for another 50 blocks?", vbYesNo + vbDefaultButton1)
        If z = vbYes Then
            nNestMax = nNestMax + 50
        Else
            Set NodX = tvwResults.Nodes.Add("NODE" & 1, tvwChild, _
                "NODE" & tvwResults.Nodes.Count + 1, "Too many references ... quitting.", 1)
            NodX.Tag = 0
            NodX.Expanded = True
            Exit Sub
        End If
    Else
        Exit Sub
    End If
End If

If nTextblockNumber = 9379 Then
    Debug.Print ""
End If

tabTBInfo.Index = "pkTBInfo"
tabTBInfo.Seek "=", nTextblockNumber
If tabTBInfo.NoMatch Then
    Set NodX = tvwResults.Nodes.Add("NODE" & nCurrentNode, tvwChild, _
        "NODE" & tvwResults.Nodes.Count + 1, "Textblock not found, could be because it required items not in the game.", 1)
    NodX.Tag = 0
    NodX.Expanded = True
    Exit Sub
End If

If tabTBInfo.Fields("Action") = Chr(0) Then
    GoTo nada:
End If
sTextblockData = tabTBInfo.Fields("Action")

'get total number of lines, only really concerned if it's > 1 so we exit before continuing
nTotalLines = 0
nDataPos = 1
x2 = InStr(nDataPos, sTextblockData, Chr(10))
If x2 > 0 Then
    Do While (nDataPos < Len(sTextblockData) + 1) And nTotalLines < 2
        x2 = InStr(nDataPos, sTextblockData, Chr(10))
        If x2 = 0 Then
            If Len(sTextblockData) - nDataPos > 1 Then nTotalLines = nTotalLines + 1
            Exit Do
        End If
        If Not x2 = nDataPos Then
            nTotalLines = nTotalLines + 1
        End If
        nDataPos = x2 + 1
    Loop
Else
    nTotalLines = 1
End If

nDataPos = 0
If Not bGreetText And Not bRoomCommands And Not bRandom Then GoTo no_commands:

'get first command
nDataPos = 1
nDataPos = InStr(nDataPos, sTextblockData, ":")
If nDataPos = 0 Then GoTo nada:

sCommand = Mid(sTextblockData, 1, nDataPos - 1)

no_commands:
nPercent1 = 0
nPercent2 = 0

nDataPos = nDataPos + 1
Do While nDataPos < Len(sTextblockData) 'loops through lines
    If bRandom Then
        nPercent2 = nPercent1
        nPercent1 = Val(sCommand)

        sCommand = (nPercent1 - nPercent2) & "%"
    End If
    
    If bRoomCommands Or bGreetText Then
        Set NodX = tvwResults.Nodes.Add("NODE" & nCurrentNode, tvwChild, _
            "NODE" & tvwResults.Nodes.Count + 1, IIf(bRandom = True, "", "Command: ") & sCommand, 1)
        NodX.Expanded = True
        nNode = tvwResults.Nodes.Count
    ElseIf bRandom Then
        Set NodX = tvwResults.Nodes.Add("NODE" & nCurrentNode, tvwChild, _
            "NODE" & tvwResults.Nodes.Count + 1, sCommand, 1)
        NodX.Expanded = True
        nNode = tvwResults.Nodes.Count
    Else
        nNode = nCurrentNode
    End If
    
    x2 = InStr(nDataPos, sTextblockData, Chr(10))
    If x2 = 0 Then x2 = Len(sTextblockData)
    sLine = Mid(sTextblockData, nDataPos, x2 - nDataPos)
    
    If sLine = "" Then GoTo next_line:
    
    If Not bRoomCommands And Not bGreetText And Not bRandom And nTotalLines > 1 Then
        nCurrLine = nCurrLine + 1
        Set NodX = tvwResults.Nodes.Add("NODE" & nCurrentNode, tvwChild, _
            "NODE" & tvwResults.Nodes.Count + 1, "Line " & nCurrLine, 3)
        NodX.Expanded = True
        nNode = tvwResults.Nodes.Count
    End If
    
    If bGreetText Then
        tvwResults.Nodes(nNode).Text = tvwResults.Nodes(nNode).Text & " --> Textblock " & Val(sLine)
        Call AddCommandNode(Val(sLine), tvwResults.Nodes.Count, False, False, False)
        GoTo next_line:
    End If
    
'    If chkHideTextblocks.Value = 0 Then
'        Set NodX = tvwResults.Nodes.Add("NODE" & nNode, tvwChild, _
'            "NODE" & tvwResults.Nodes.Count + 1, "--> Raw Line: " & sLine, 3)
'        NodX.Expanded = True
'        NodX.Tag = nTextblockNumber
'    End If
    
    nLinePos = 1
    nLinePos = InStr(nLinePos, sLine, ":")
    If nLinePos = 0 Then nLinePos = Len(sLine) + 1
    
    'nRepeats = 1
    'nRepeatNode = 0
    sLastCommand = ""
    sLineCommand = Mid(sLine, 1, nLinePos - 1)
    
    Do While nLinePos < Len(sLine) + 2
        
'        If Right(sLineCommand, 3) = "898" Then
'            Debug.Print ""
'        End If
        If sLineCommand = sLastCommand Then
            nRepeats = nRepeats + 1
            GoTo next_cmd:
        Else
            'this code is also below
            If nRepeats > 1 Then
                tvwResults.Nodes(nRepeatNode).Text = tvwResults.Nodes(nRepeatNode).Text _
                    & " (x" & nRepeats & ")"
            End If
            nRepeatNode = tvwResults.Nodes.Count + 1
            nRepeats = 1
        End If
        
        sLastCommand = sLineCommand
        
        nRoom = 0
        nMap = 0
        If InStr(1, sLineCommand, "cast ") > 0 Then
            nValue = ExtractValueFromString(sLineCommand, "cast ")
            If nValue > 0 Then
                Set NodX = tvwResults.Nodes.Add("NODE" & nNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, "Cast: " & GetSpellName(nValue, bHideRecordNumbers), "BOLT")
                NodX.Tag = nValue
                NodX.Expanded = True
                NodX.Bold = True
            End If
        ElseIf InStr(1, sLineCommand, "item ") > 0 Then
            nValue = ExtractValueFromString(sLineCommand, "item ")
'                If nValue = 223 Then
'                    Debug.Print 223
'                End If
            If nValue > 0 Then
                y = InStr(1, sLineCommand, "item ")
                Set NodX = tvwResults.Nodes.Add("NODE" & nNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, "Item, " & Left(sLineCommand, y - 1) & ": " _
                    & GetItemName(nValue, bHideRecordNumbers), "ITEM")
                NodX.Tag = nValue
                NodX.Expanded = True
                NodX.Bold = True
            End If
        ElseIf InStr(1, sLineCommand, "ability ") > 0 And Not InStr(1, sLineCommand, "testability") > 0 Then
            nValue = ExtractValueFromString(sLineCommand, "ability ")
            If nValue > 0 Then
                y = InStr(1, sLineCommand, "ability ")
                Set NodX = tvwResults.Nodes.Add("NODE" & nNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, "Ability, " & Left(sLineCommand, y - 1) _
                    & ": " & GetAbilityName(nValue) & " (" & sLineCommand & ")", "ARROW")
                NodX.Tag = nValue
                NodX.Expanded = True
                NodX.Bold = False
            End If
        ElseIf InStr(1, sLineCommand, "class ") > 0 Then
            nValue = ExtractValueFromString(sLineCommand, "class ")
            If nValue > 0 Then
                Set NodX = tvwResults.Nodes.Add("NODE" & nNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, "Class: " & GetClassName(nValue), "REDARROW")
                NodX.Tag = nValue
                NodX.Expanded = True
                NodX.Bold = True
            End If
        ElseIf InStr(1, sLineCommand, "race ") > 0 Then
            nValue = ExtractValueFromString(sLineCommand, "race ")
            If nValue > 0 Then
                Set NodX = tvwResults.Nodes.Add("NODE" & nNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, "Race: " & GetRaceName(nValue), "REDARROW")
                NodX.Tag = nValue
                NodX.Expanded = True
                NodX.Bold = True
            End If
        ElseIf InStr(1, sLineCommand, "addexp ") > 0 Then
            nValue = ExtractValueFromString(sLineCommand, "addexp ")
            If nValue > 0 Then
                Set NodX = tvwResults.Nodes.Add("NODE" & nNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, "AddExp: " & PutCommas(nValue), "ARROW")
                NodX.Tag = nValue
                NodX.Expanded = True
                NodX.Bold = True
            End If
        ElseIf InStr(1, sLineCommand, "learnspell ") > 0 Then
            nValue = ExtractValueFromString(sLineCommand, "learnspell ")
            If nValue > 0 Then
                Set NodX = tvwResults.Nodes.Add("NODE" & nNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, "Learnspell: " & GetSpellName(nValue, bHideRecordNumbers), "BOLT")
                NodX.Tag = nValue
                NodX.Expanded = True
                NodX.Bold = True
            End If
        ElseIf InStr(1, sLineCommand, "checkspell ") > 0 Then
            nValue = ExtractValueFromString(sLineCommand, "checkspell ")
            If nValue > 0 Then
                Set NodX = tvwResults.Nodes.Add("NODE" & nNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, "Checkspell: " & GetSpellName(nValue, bHideRecordNumbers), "BOLT")
                NodX.Tag = nValue
                NodX.Expanded = True
                NodX.Bold = True
            End If
        ElseIf InStr(1, sLineCommand, "summon ") > 0 Then
            nValue = ExtractValueFromString(sLineCommand, "summon ")
            If nValue > 0 Then
                Set NodX = tvwResults.Nodes.Add("NODE" & nNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, "Summon: " & GetMonsterName(nValue, bHideRecordNumbers), "MONSTER")
                NodX.Tag = nValue
                NodX.Expanded = True
                NodX.Bold = True
            End If
        ElseIf InStr(1, sLineCommand, "random ") > 0 Then
            nValue = ExtractValueFromString(sLineCommand, "random ")
            If nValue > 0 Then
                Set NodX = tvwResults.Nodes.Add("NODE" & nNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, "random: " & nValue, "PAPER")
                NodX.Tag = nValue
                NodX.Expanded = True
            End If
            Call AddCommandNode(nValue, tvwResults.Nodes.Count, False, True, False)
        ElseIf InStr(1, sLineCommand, "price ") > 0 Then
            nValue = ExtractValueFromString(sLineCommand, "price ")
            Select Case UCase(Right(sLineCommand, 1))
                Case "R": sChar = " runic"
                Case "P": sChar = " platinum"
                Case "G": sChar = " gold"
                Case "S": sChar = " silver"
                Case Else: sChar = " copper"
            End Select
            If nValue > 0 Then
                Set NodX = tvwResults.Nodes.Add("NODE" & nNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, "Cost: " & PutCommas(nValue) & sChar, "MONEY")
                NodX.Tag = nValue
                NodX.Expanded = True
                NodX.Bold = True
            End If
        ElseIf InStr(1, sLineCommand, "givecoins ") > 0 Then
            nValue = ExtractValueFromString(sLineCommand, "givecoins ")
            Select Case UCase(Right(sLineCommand, 1))
                Case "R": sChar = " runic"
                Case "P": sChar = " platinum"
                Case "G": sChar = " gold"
                Case "S": sChar = " silver"
                Case Else: sChar = " copper"
            End Select
            If nValue > 0 Then
                Set NodX = tvwResults.Nodes.Add("NODE" & nNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, "Givecoins: " & PutCommas(nValue) & sChar, "MONEY")
                NodX.Tag = nValue
                NodX.Expanded = True
                NodX.Bold = True
            End If
        ElseIf InStr(1, sLineCommand, "teleport ") > 0 Then
            x = InStr(1, sLineCommand, "teleport ") + Len("teleport ")
            y = x
            Do While y <= Len(sLineCommand) + 1
                sChar = Mid(sLineCommand, y, 1)
                Select Case sChar
                    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
                    Case " ":
                        If y > x And nRoom = 0 Then
                            nRoom = Val(Mid(sLineCommand, x, y - x))
                            x = y + 1
                        Else
                            nMap = Val(Mid(sLineCommand, x, y - x))
                            Exit Do
                        End If
                    Case Else:
                        If y > x And nRoom = 0 Then
                            nRoom = Val(Mid(sLineCommand, x, y - x))
                        Else
                            nMap = Val(Mid(sLineCommand, x, y - x))
                        End If
                        Exit Do
                End Select
                y = y + 1
            Loop
            
            If nMap = 0 Then nMap = nDefaultMap
            
            If Not nRoom = 0 And Not nMap = 0 Then
                'If nMap = 0 Then nMap = nMapNumber
                
                Set NodX = tvwResults.Nodes.Add("NODE" & nNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, "Teleport: " & GetRoomName(, nMap, nRoom, False), "ROOM")
                NodX.Tag = nMap & "/" & nRoom
                NodX.Expanded = True
                NodX.Bold = True
            Else
                Set NodX = tvwResults.Nodes.Add("NODE" & nNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, sLineCommand, 3)
                NodX.Expanded = True
            End If
        ElseIf InStr(1, sLineCommand, "remoteaction ") > 0 Then
            x = InStr(1, sLineCommand, "remoteaction ") + Len("remoteaction ")
            y = x
            Do While y <= Len(sLineCommand) + 1
                sChar = Mid(sLineCommand, y, 1)
                Select Case sChar
                    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
                    Case " ":
                        If y > x And nRoom = 0 Then
                            nRoom = Val(Mid(sLineCommand, x, y - x))
                            'position to after message
                            x = InStr(y + 1, sLineCommand, " ")
                            If x = 0 Then Exit Do
                            'position to after first variable
                            x = InStr(x + 1, sLineCommand, " ")
                            If x = 0 Then Exit Do
                            y = x + 1
                            x = x + 1
                        End If
                    Case Else:
                        If y > x And nRoom = 0 Then
                            nRoom = Val(Mid(sLineCommand, x, y - x))
                        ElseIf y > x Then
                            Select Case Val(Mid(sLineCommand, x, y - x))
                                Case 0: sChar = " (on the N exit)"
                                Case 1: sChar = " (on the S exit)"
                                Case 2: sChar = " (on the E exit)"
                                Case 3: sChar = " (on the W exit)"
                                Case 4: sChar = " (on the NE exit)"
                                Case 5: sChar = " (on the NW exit)"
                                Case 6: sChar = " (on the SE exit)"
                                Case 7: sChar = " (on the SW exit)"
                                Case 8: sChar = " (on the U exit)"
                                Case 9: sChar = " (on the D exit)"
                            End Select
                        End If
                        Exit Do
                End Select
                y = y + 1
            Loop
            
            nMap = nDefaultMap
            
            If Not nRoom = 0 And Not nMap = 0 Then
                'If nMap = 0 Then nMap = nMapNumber
                
                Set NodX = tvwResults.Nodes.Add("NODE" & nNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, "Remote Action" & sChar & ": " & GetRoomName(, nMap, nRoom, False), "ROOM")
                NodX.Tag = nMap & "/" & nRoom
                NodX.Expanded = True
                NodX.Bold = True
            Else
                Set NodX = tvwResults.Nodes.Add("NODE" & nNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, sLineCommand, "PAPER")
                NodX.Expanded = True
            End If
        ElseIf InStr(1, sLineCommand, "testskill ") > 0 Then
            nValue = 0
            x = InStr(1, sLineCommand, "testskill ") + Len("testskill ")
            x = InStr(x, sLineCommand, " ") + 1 'position after agility,health,etc
            If x = 1 Then GoTo no_testskill_tb:
            x = InStr(x, sLineCommand, " ") + 1 'position after test amount
            If x = 1 Then GoTo no_testskill_tb:
            y = x
            Do While y <= Len(sLineCommand) + 1
                sChar = Mid(sLineCommand, y, 1)
                Select Case sChar
                    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
                    Case Else:
                        If y > x And nRoom = 0 Then
                            nValue = Val(Mid(sLineCommand, x, y - x))
                        End If
                        Exit Do
                End Select
                y = y + 1
            Loop
no_testskill_tb:
            If Not nValue = 0 Then
                Set NodX = tvwResults.Nodes.Add("NODE" & nNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, sLineCommand, "PAPER")
                NodX.Tag = 0
                NodX.Expanded = True
                NodX.Bold = False
                Call AddCommandNode(nValue, tvwResults.Nodes.Count, False, False, False)
            Else
                Set NodX = tvwResults.Nodes.Add("NODE" & nNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, sLineCommand, "PAPER")
                NodX.Expanded = True
            End If
            '''''''''''''''''''''''''''''''''''''''''
            
        Else
            If chkHideTextblocks.Value = 0 Then
                Set NodX = tvwResults.Nodes.Add("NODE" & nNode, tvwChild, _
                    "NODE" & tvwResults.Nodes.Count + 1, sLineCommand, "PAPER")
                NodX.Expanded = True
            End If
        End If
        
next_cmd:
        'Debug.Print sLine
        nLinePos = InStr(nLinePos, sLine, ":") + 1 'position nLinePos to start of next command
        If nLinePos = 1 Then
            'this code is also above
            If nRepeats > 1 Then
                tvwResults.Nodes(nRepeatNode).Text = tvwResults.Nodes(nRepeatNode).Text _
                    & " (x" & nRepeats & ")"
            End If
            nRepeatNode = tvwResults.Nodes.Count + 1
            nRepeats = 1
            Exit Do
        End If
        
        y2 = InStr(nLinePos, sLine, ":")
        If y2 = 0 Then y2 = Len(sLine) + 1
        If y2 = nLinePos Then GoTo next_cmd:
        sLineCommand = Mid(sLine, nLinePos, y2 - nLinePos)
        
        nLinePos = nLinePos + 1
    Loop
    
next_line:
    nDataPos = InStr(nDataPos, sTextblockData, Chr(10)) + 1
    If nDataPos = 1 Then Exit Do
    
    If bGreetText Or bRoomCommands Or bRandom Then
        x2 = InStr(nDataPos, sTextblockData, ":")
        If x2 = 0 Then x2 = Len(sTextblockData) + 1
        If x2 = nDataPos Then GoTo next_line:
        sCommand = Mid(sTextblockData, nDataPos, x2 - nDataPos)
        
        nDataPos = x2 + 1
    End If
Loop

GoTo out:

nada:
If chkHideTextblocks.Value = 0 Then
    Set NodX = tvwResults.Nodes.Add("NODE" & nCurrentNode, tvwChild, _
        "NODE" & tvwResults.Nodes.Count + 1, "Dialog", 3)
    NodX.Tag = 0
    NodX.Expanded = True
End If

If tabTBInfo.Fields("LinkTo") > 0 Then
    If chkHideTextblocks.Value = 0 Then
        Set NodX = tvwResults.Nodes.Add("NODE" & nCurrentNode, tvwChild, _
            "NODE" & tvwResults.Nodes.Count + 1, "Link to Textblock " & tabTBInfo.Fields("LinkTo"), 3)
        NodX.Tag = tabTBInfo.Fields("LinkTo")
        NodX.Expanded = True
    End If
    Call AddCommandNode(tabTBInfo.Fields("LinkTo"), tvwResults.Nodes.Count, False, False, False)
End If

out:

On Error Resume Next
Set NodX = Nothing
Set nodY = Nothing
Exit Sub
Error:
Call HandleError("AddCommandNode")
Resume out:
End Sub


Private Sub mnuExpandItem_Click(Index As Integer)
Dim bExpanded As Boolean

On Error GoTo Error:

If nLastNode < 1 Then Exit Sub

Me.MousePointer = vbHourglass
DoEvents
Call LockWindowUpdate(Me.hWnd)

If Index = 0 Then
    bExpanded = False
Else
    bExpanded = True
End If

tvwResults.Nodes(nLastNode).Expanded = bExpanded
Call ExpandBranch(nLastNode, bExpanded)
tvwResults.Nodes(nLastNode).EnsureVisible

out:
Me.MousePointer = vbDefault
Call LockWindowUpdate(0&)
Exit Sub
Error:
Call HandleError("mnuExpandItem_Click")
Resume out:
End Sub

Private Sub ExpandBranch(ByVal nParentNode As Long, ByVal bExpand As Boolean)
Dim x As Long

If tvwResults.Nodes(nParentNode).Children = 0 Then Exit Sub

x = tvwResults.Nodes(nParentNode).Child.Index
Do Until x = tvwResults.Nodes(x).LastSibling.Index
    tvwResults.Nodes(x).Expanded = bExpand
    If tvwResults.Nodes(x).Children > 0 Then Call ExpandBranch(x, bExpand)
    x = tvwResults.Nodes(x).Next.Index
    'x = x + 1
Loop
tvwResults.Nodes(x).Expanded = bExpand
If tvwResults.Nodes(x).Children > 0 Then Call ExpandBranch(x, bExpand)

End Sub
Private Sub timWait_Timer()
timWait.Enabled = False
End Sub

Private Sub tvwResults_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 2 Then
    PopupMenu mnuExpand
End If

End Sub

Private Sub tvwResults_NodeClick(ByVal Node As MSComctlLib.Node)
Dim oLI As ListItem, oLV As ListView, x As Integer, sStr As String, nNum As Long
Dim RoomExits As RoomExitType
On Error GoTo Error:

If timWait.Enabled = True Then Exit Sub
nLastNode = Node.Index
'Debug.Print Node.Index
If Val(Node.Tag) < 1 Then Exit Sub
frmMain.bDontSetMainFocus = True

If objFormOwner Is Nothing Then Set objFormOwner = frmMain

nNum = 15
For x = 1 To nNum
    Select Case x
        Case 1: sStr = "Room"
        Case 2: sStr = "Monster"
        Case 3: sStr = "Textblock"
        Case 4: sStr = "Item"
        Case 5: sStr = "Spell"
        Case 6: sStr = "Shop"
        Case 7: sStr = "Group"
        Case 8: sStr = "Teleport"
        Case 9: sStr = "Summon"
        Case 10: sStr = "Cast"
        Case 11: sStr = "--> Raw Line"
        Case 12: sStr = "Learnspell"
        Case 13: sStr = "Remote Action"
        Case 14: sStr = "Link to"
        Case 15: sStr = "Checkspell"
    End Select
    
    If Left(Node.Text, Len(sStr)) = sStr Then Exit For
    If x = nNum Then GoTo out:
Next x

Select Case x
    Case 8, 13: x = 1
    Case 9: x = 2
    Case 10, 12, 15: x = 5
    Case 11, 14: x = 3
End Select

If x = 1 Or x = 7 Then 'room/group
    RoomExits = ExtractMapRoom(Node.Tag)
Else
    nNum = Val(Node.Tag)
    If nNum <= 0 Then GoTo out:
End If

Select Case x
    Case 1, 7: 'room/group
        Set oLV = Nothing
    Case 2: 'monster
        Set oLV = frmMain.lvMonsters
    Case 3: 'textblock
        tabTBInfo.Index = "pkTBInfo"
        tabTBInfo.Seek "=", Val(Node.Tag)
        If tabTBInfo.NoMatch Then GoTo out:
        
'        If Not Len(tabTBInfo.Fields("Action")) < 2 Then
'            If Len(tabTBInfo.Fields("Action")) > 900 Then
'                Clipboard.clear
'                Clipboard.SetText PutCrLF(tabTBInfo.Fields("Action"))
'                MsgBox "Raw textblock too long to display; copied to clipboard.", vbInformation
'            Else
                If Not tabTBInfo.Fields("Action") = Chr(0) Then
                    Clipboard.clear
                    Clipboard.SetText "Raw textblock (" & Val(Node.Tag) & "): " & vbCrLf & PutCrLF(tabTBInfo.Fields("Action"))
                End If
'                MsgBox "Raw textblock (" & Val(Node.Tag) & "): " & vbCrLf & tabTBInfo.Fields("Action"), vbInformation
'            End If
'        End If
        GoTo out:
    Case 4: 'item
        Call frmMain.GotoItem(nNum)
        'this has to be here 'cause for some damn reason the activate event keeps firing off on frmMap when the cmdNav_Click goes (i think)
        If objFormOwner Is frmMap Then
            If frmMap.chkMapOptions(6).Value = 0 Then
                If frmMap.chkMapOptions(8).Value = 1 Then
                    Call SetTopMostWindow(frmMap.hWnd, False)
                    frmMain.SetFocus
                End If
            End If
        End If
        GoTo out:

    Case 5: 'spell
        Call frmMain.GotoSpell(nNum)
        'this has to be here 'cause for some damn reason the activate event keeps firing off on frmMap when the cmdNav_Click goes (i think)
        If objFormOwner Is frmMap Then
            If frmMap.chkMapOptions(6).Value = 0 Then
                If frmMap.chkMapOptions(8).Value = 1 Then
                    Call SetTopMostWindow(frmMap.hWnd, False)
                    frmMain.SetFocus
                End If
            End If
        End If
        GoTo out:
    Case 6: 'shop
        Set oLV = frmMain.lvShops
End Select

If x = 1 Or x = 7 Then 'rooms/group
    If objFormOwner Is frmMain Then Call frmMain.cmdNav_Click(10)
    Call objFormOwner.MapStartMapping(RoomExits.Map, RoomExits.Room)
Else
    Set oLI = oLV.FindItem(nNum, lvwText, , 0)

    If Not oLI Is Nothing Then
        Select Case x
            Case 1: 'room
            Case 2: 'monster
                Call ClearListViewSelections(oLV)
                Call frmMain.lvMonsters_ItemClick(oLI)
                Call frmMain.cmdNav_Click(8) 'monster
                
            Case 3: 'textblock
                GoTo out:
            Case 4: 'item
                Select Case tabItems.Fields("ItemType")
                    Case 0:
                        If tabItems.Fields("Worn") = 0 Then
                            Call ClearListViewSelections(oLV)
                            Call frmMain.lvOtherItems_ItemClick(oLI)
                            Call frmMain.cmdNav_Click(7) 'sundry
                        Else
                            Call ClearListViewSelections(oLV)
                            Call frmMain.lvArmour_ItemClick(oLI)
                            Call frmMain.cmdNav_Click(1) ' armour
                        End If
                    Case 1:
                        Call ClearListViewSelections(oLV)
                        Call frmMain.lvWeapons_ItemClick(oLI)
                        Call frmMain.cmdNav_Click(0) ' weapons
                    Case Else:
                        Call ClearListViewSelections(oLV)
                        Call frmMain.lvOtherItems_ItemClick(oLI)
                        Call frmMain.cmdNav_Click(7) 'sundry
                End Select
        
            Case 5: 'spell
                Call ClearListViewSelections(oLV)
                Call frmMain.lvSpells_ItemClick(oLI)
                Call frmMain.cmdNav_Click(2) 'spell
                
            Case 6: 'shop
                Call ClearListViewSelections(oLV)
                Call frmMain.lvShops_ItemClick(oLI)
                Call frmMain.cmdNav_Click(9) 'shop
                Call ClearListViewSelections(frmMain.lvShopDetail)

        End Select
        'this has to be here 'cause for some damn reason the activate event keeps firing off on frmMap when the cmdNav_Click goes (i think)
        If objFormOwner Is frmMap Then
            If frmMap.chkMapOptions(6).Value = 0 Then
                If frmMap.chkMapOptions(8).Value = 1 Then
                    Call SetTopMostWindow(frmMap.hWnd, False)
                    frmMain.SetFocus
                End If
            End If
        End If
    Else
        MsgBox sStr & " " & nNum & " not found in current " & sStr & " list."
    End If
End If


If Not objFormOwner Is frmMap Then
    If Not frmMap.chkMapOptions(6).Value = 0 Then
        Me.SetFocus
    End If
End If

out:
frmMain.bDontSetMainFocus = False
Set oLI = Nothing
Set oLV = Nothing
Exit Sub

Error:
Call HandleError("tvwResults_NodeClick")
Resume out:
End Sub
