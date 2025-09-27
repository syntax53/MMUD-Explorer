Attribute VB_Name = "modListViewExt"
Option Explicit
Option Base 0

Private Const SEP As String = vbNullChar
Private Const SORT_TOKEN As String = "SORTSTATE:"
Private Const DEBUG_STICKY As Boolean = False   ' flip to False to silence

Public Enum ListDataType
    ldtstring = 0
    ldtnumber = 1
    ldtDateTime = 2
End Enum

' Extend the stored state so we can reapply exactly:
'  - LastCol:   1-based column index
'  - Asc:       0 = Desc, 1 = Asc
'  - DType:     ListDataType (your enum, e.g., ldtstring=0, ldtnumber=1, ldtdate=2)
'  - ByTag:     0 = SortListView, 1 = SortListViewByTag
Private Type tSortState
    LastCol As Integer
    asc     As Integer
    DType   As Integer
    ByTag   As Integer
End Type

Private m_LastSortWasTagCol As Boolean


'--- return the ListView column index for the visible Flag column (fallback to 2) ---
Private Function LV_GetFlagColIndex(ByVal lv As ListView) As Long
    Dim i As Long
    On Error Resume Next
    For i = 1 To lv.ColumnHeaders.Count
        If StrComp(lv.ColumnHeaders(i).Key, "Flag", vbTextCompare) = 0 Then
            LV_GetFlagColIndex = i
            Exit Function
        End If
    Next
    On Error GoTo 0

    For i = 1 To lv.ColumnHeaders.Count
        If StrComp(lv.ColumnHeaders(i).Text, "Flag", vbTextCompare) = 0 Then
            LV_GetFlagColIndex = i
            Exit Function
        End If
    Next

    LV_GetFlagColIndex = 2
End Function

'--- map the visible flag into a sort "bucket" (BUY/SELL -> SHOP; others unchanged) ---
Private Function LV_MapSortBucket(ByVal flagBase As String) As String
    Select Case UCase$(Trim$(flagBase))
        Case "BUY", "SELL": LV_MapSortBucket = "SHOP"
        Case "CARRIED":     LV_MapSortBucket = "CARRIED"
        Case "STASH":       LV_MapSortBucket = "STASH"
        Case "MANUAL":      LV_MapSortBucket = "MANUAL"
        Case "DROP":        LV_MapSortBucket = "DROP"
        Case "HIDE":        LV_MapSortBucket = "HIDE"
        Case "PICKUP":      LV_MapSortBucket = "PICKUP"
        Case Else:          LV_MapSortBucket = "ZZZ"
    End Select
End Function

'--- convert bucket to a 2-digit rank for primary sort (BUY+SELL share rank) ---
Private Function LV_SortBucketRank(ByVal bucket As String) As String
    Select Case bucket
        Case "CARRIED": LV_SortBucketRank = "01"
        Case "STASH":   LV_SortBucketRank = "02"
        Case "MANUAL":  LV_SortBucketRank = "03"
        Case "SHOP":    LV_SortBucketRank = "04" ' BUY and SELL both land here
        Case "DROP":    LV_SortBucketRank = "05"
        Case "HIDE":    LV_SortBucketRank = "06"
        Case "PICKUP":  LV_SortBucketRank = "07"
        Case Else:      LV_SortBucketRank = "99"
    End Select
End Function

'--- build the primary token from a ListItem by reading its visible Flag text ---
Private Function LV_GetPriFromLI(ByVal lv As ListView, ByRef li As ListItem) As String
    Dim flagCol As Long, sFlagText As String, baseFlag As String
    flagCol = LV_GetFlagColIndex(lv)

    If flagCol <= 1 Then
        sFlagText = CStr(li.Text)
    ElseIf li.ListSubItems.Count >= (flagCol - 1) Then
        sFlagText = li.ListSubItems(flagCol - 1).Text
    Else
        sFlagText = ""
    End If

    ' strip any qty suffix like "BUY x3" -> "BUY"
    baseFlag = GetFlagBase(sFlagText)
    LV_GetPriFromLI = LV_SortBucketRank(LV_MapSortBucket(baseFlag))
End Function

' Returns the first token before a space.
' Handles Null, empty, and multi-space strings without raising errors.
Private Function GetFlagBase(ByVal anyText As Variant) As String
    Dim s As String, p As Long
    If IsNull(anyText) Then
        s = ""
    Else
        s = CStr(anyText)
    End If
    s = Trim$(s)
    If LenB(s) = 0 Then
        GetFlagBase = ""
        Exit Function
    End If
    p = InStr(1, s, " ")
    If p > 0 Then
        GetFlagBase = Left$(s, p - 1)
    Else
        GetFlagBase = s
    End If
End Function


' Safe getters
Private Function LV_GetCell(ByRef li As ListItem, ByVal colIndex As Integer) As String
    If colIndex <= 1 Then
        LV_GetCell = li.Text
    Else
        If li.ListSubItems.Count >= colIndex - 1 Then
            LV_GetCell = li.ListSubItems(colIndex - 1).Text
        Else
            LV_GetCell = ""
        End If
    End If
End Function

Private Function LV_FlagBase(ByRef li As ListItem) As String
    Dim s As String
    s = ""
    If li.ListSubItems.Count >= 2 Then s = Trim$(li.ListSubItems(2).Text)
    LV_FlagBase = UCase$(ParseActionBase(s))
End Function

' Dump a row’s key parts
'Private Sub DBG_PrintRow(ByVal label As String, ByVal i As Long, ByRef li As ListItem, _
'                         ByVal pri As Long, ByVal secKey As String, ByVal tie As String, _
'                         ByVal composite As String, ByVal clickedCol As Long, ByVal hidCol As Long, _
'                         ByVal tagVal As String)
'    If Not DEBUG_STICKY Then Exit Sub
'    Debug.Print label & " #" & Right$("0000" & CStr(i), 4) & _
'                "  PRI=" & Right$("000" & CStr(pri), 3) & _
'                "  FLAG=" & LV_FlagBase(li) & _
'                "  SECKEY=[" & secKey & "]  TIE=" & tie & _
'                "  HIDDEN=[" & LV_GetCell(li, hidCol) & "]" & _
'                "  COMPOSITE=[" & composite & "]" & _
'                "  CLICKCOL=[" & LV_GetCell(li, clickedCol) & "]" & _
'                "  TAG=[" & tagVal & "]"
'End Sub

' After-sort mini dump
Private Sub DBG_DumpAfterSort(ByVal title As String, ByVal lv As ListView, ByVal clickedCol As Long, ByVal hidCol As Long, ByVal maxRows As Long)
    If Not DEBUG_STICKY Then Exit Sub
    Dim n As Long, li As ListItem
    Debug.Print "---- AFTER SORT (" & title & ")  hidCol=" & hidCol & "  clickCol=" & clickedCol & "  Count=" & lv.ListItems.Count
    For n = 1 To lv.ListItems.Count
        If n > maxRows Then Exit For
        Set li = lv.ListItems(n)
        Debug.Print "  [" & Right$("0000" & CStr(n), 4) & "]  HIDDEN=[" & LV_GetCell(li, hidCol) & "]  FLAG=" & LV_FlagBase(li) & _
                    "  CLICKCOL=[" & LV_GetCell(li, clickedCol) & "]  TAG=[" & CStr(li.Tag) & "]  NAME=[" & LV_GetCell(li, 2) & "]"
    Next
End Sub

' Before-sort mini dump (composite computed in caller and written to hidden)
Private Sub DBG_DumpBeforeSort(ByVal title As String, ByVal lv As ListView, ByVal clickedCol As Long, ByVal hidCol As Long, ByVal maxRows As Long)
    If Not DEBUG_STICKY Then Exit Sub
    Dim i As Long, li As ListItem
    Debug.Print "---- BEFORE SORT (" & title & ") hidCol=" & hidCol & "  clickCol=" & clickedCol & "  Count=" & lv.ListItems.Count
    For i = 1 To lv.ListItems.Count
        If i > maxRows Then Exit For
        Set li = lv.ListItems(i)
        Debug.Print "  [" & Right$("0000" & CStr(i), 4) & "]  HIDDEN=[" & LV_GetCell(li, hidCol) & "]  FLAG=" & LV_FlagBase(li) & _
                    "  CLICKCOL=[" & LV_GetCell(li, clickedCol) & "]  TAG=[" & CStr(li.Tag) & "]  NAME=[" & LV_GetCell(li, 2) & "]"
    Next
End Sub


Public Sub LV_DeleteSelected(ByRef lvListView As ListView)
On Error GoTo error:
    Dim i As Long, selCount As Long
    Dim highestSelIdx As Long
    Dim originalCount As Long
    Dim targetNewIndex As Long
    Dim resp As VbMsgBoxResult
    Dim bCarriedAddedOrRemoved As Boolean
    
    If lvListView Is Nothing Then Exit Sub
    If lvListView.ListItems.Count = 0 Then Exit Sub
    
    originalCount = lvListView.ListItems.Count
    
    ' Count selected items and find the highest selected index
    highestSelIdx = 0
    For i = 1 To originalCount
        If lvListView.ListItems(i).Selected Then
            selCount = selCount + 1
            If i > highestSelIdx Then highestSelIdx = i
        End If
    Next i
    
    If selCount = 0 Then Exit Sub  ' nothing to delete
    
    ' If more than one selected, confirm first
    If selCount > 1 Then
        resp = MsgBox("Are you sure you want to delete the selected rows?", _
                      vbYesNo Or vbQuestion Or vbDefaultButton2, "Confirm Delete")
        If resp <> vbYes Then Exit Sub
    End If
    
    ' Compute the target row to select after deletion:
    ' We want the row that was originally just AFTER the highest selected row.
    ' After deleting 'selCount' rows (all <= highestSelIdx), that original "next" row
    ' moves to: highestSelIdx - selCount + 1.
    targetNewIndex = highestSelIdx - selCount + 1
    ' We will clamp this to the valid range after deletion is done.
    
    ' Delete selected items (reverse to avoid index shifting issues)
    For i = originalCount To 1 Step -1
        If lvListView.ListItems(i).Selected Then
            If lvListView.ListItems(i).ListSubItems.Count >= 2 Then
                If InStr(1, lvListView.ListItems(i).ListSubItems(2).Text, "CARRIED", vbTextCompare) > 0 Then bCarriedAddedOrRemoved = True
            End If
            lvListView.ListItems.Remove i
        End If
    Next i
    
    ' If nothing remains, we're done
    If lvListView.ListItems.Count = 0 Then Exit Sub
    
    ' Clamp targetNewIndex into [1 .. newCount]
    If targetNewIndex < 1 Then targetNewIndex = 1
    If targetNewIndex > lvListView.ListItems.Count Then targetNewIndex = lvListView.ListItems.Count
    
    ' Select and reveal the target row so repeated deletes are quick
    With lvListView.ListItems(targetNewIndex)
        .Selected = True
        On Error Resume Next
        .EnsureVisible
        On Error GoTo 0
    End With
    
    If bCarriedAddedOrRemoved Then
        If bCharLoaded And Not bStartup Then bPromptSave = True
        Call frmMain.RefreshAll
    End If
    lvListView.SetFocus
    Exit Sub
    
out:
On Error Resume Next
Exit Sub
error:
Call HandleError("LV_DeleteSelected")
Resume out:
End Sub

Public Sub LV_InvertSelection(ByRef lvListView As ListView)
On Error GoTo error:

    Dim i As Long
    
    If lvListView Is Nothing Then Exit Sub
    If lvListView.ListItems.Count = 0 Then Exit Sub
    
    For i = 1 To lvListView.ListItems.Count
        lvListView.ListItems(i).Selected = Not lvListView.ListItems(i).Selected
    Next i

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("LV_InvertSelection")
Resume out:
End Sub


'--- UPDATED: apply action to SubItem(2), with +/- quantity handling for indices 15/16
'--- UPDATED: apply action to SubItem(2) (main action) and SubItem(10) (CARRIED/STASH),
'             with +/- quantity handling for indices 15/16 (incl. PICKUP now)
'--- UPDATED: all actions in SubItem(2); +/- qty for DROP/HIDE/SELL/PICKUP/STASH; CARRIED no qty
Public Sub LV_SetActionCell(ByRef lvListView As ListView, ByVal actionIndex As Long)
On Error GoTo done
    Dim i As Long
    Dim li As ListItem
    Dim curText As String, baseAction As String
    Dim newText As String
    Dim qty As Long
    Dim anySelected As Boolean
    Dim bCarriedAddedOrRemoved As Boolean
    
    If lvListView Is Nothing Then Exit Sub
    If lvListView.ListItems.Count = 0 Then Exit Sub
    If lvListView.ColumnHeaders.Count < 3 Then Exit Sub ' need SubItem(2)

    For i = 1 To lvListView.ListItems.Count
        If lvListView.ListItems(i).Selected Then
            anySelected = True
            Set li = lvListView.ListItems(i)
            
            ' Ensure subitem(2) exists
            Do While li.ListSubItems.Count < 2
                li.ListSubItems.Add , , ""
            Loop
            
            curText = Trim$(li.ListSubItems(2).Text)
            Call ParseActionAndQty(curText, baseAction, qty) ' qty>=1 on return
            
            If actionIndex > 0 Then
                If actionIndex = 12 Then
                    bCarriedAddedOrRemoved = True
                ElseIf actionIndex <> 15 And actionIndex <> 16 Then  '12 = carried, 15/16 = -/+
                    If baseAction = "CARRIED" Then bCarriedAddedOrRemoved = True
                End If
            End If
            
            Select Case actionIndex
                Case 9      ' DROP/HIDE/<clear> cycle
                    Select Case baseAction
                        Case "DROP":   newText = "HIDE"
                        Case "HIDE":   newText = ""                  ' clear (qty discarded)
                        Case Else:     newText = "DROP"
                    End Select
                    
                    
                Case 10     ' PICKUP toggle
                    If baseAction = "PICKUP" Then
                        newText = ""                                  ' clear
                    Else
                        newText = "PICKUP"
                    End If
                    
                Case 11     ' SELL toggle
'                    If baseAction = "SELL" Then
'                        newText = ""                                  ' clear
'                    Else
'                        newText = "SELL"
'                    End If
                    Select Case baseAction
                        Case "SELL":   newText = "BUY"
                        Case "BUY":    newText = ""                  ' clear (qty discarded)
                        Case Else:     newText = "SELL"
                    End Select
                    
                Case 12     ' CARRIED toggle (no qty)
                    If baseAction = "CARRIED" Then
                        newText = ""                                  ' clear
                        ' (shop repopulation intentionally removed)
                        ' li.ListSubItems(2).Text = GetShopRoomNames(Val(li.Text), , bHideRecordNumbers)
                    Else
                        newText = "CARRIED"
                    End If

                Case 17     ' STASH toggle (qty allowed via +/- later)
                    If baseAction = "STASH" Then
                        newText = ""                                  ' clear
                        ' (shop repopulation intentionally removed)
                        ' li.ListSubItems(2).Text = GetShopRoomNames(Val(li.Text), , bHideRecordNumbers)
                    Else
                        newText = "STASH"
                    End If

                Case 15     ' minus: adjust qty for DROP/HIDE/SELL/PICKUP/STASH
                    Select Case baseAction
                        Case "DROP", "HIDE", "BUY", "SELL", "PICKUP", "STASH"
                            If qty > 1 Then qty = qty - 1
                            If qty <= 1 Then
                                newText = baseAction
                            Else
                                newText = baseAction & " x" & CStr(qty)
                            End If
                        Case Else
                            GoTo nextItem
                    End Select

                Case 16     ' plus: adjust qty for DROP/HIDE/SELL/PICKUP/STASH
                    Select Case baseAction
                        Case "DROP", "HIDE", "BUY", "SELL", "PICKUP", "STASH"
                            qty = qty + 1
                            If qty <= 1 Then
                                newText = baseAction
                            Else
                                newText = baseAction & " x" & CStr(qty)
                            End If
                        Case Else
                            GoTo nextItem
                    End Select
                    
                Case 18     'clear
                    newText = ""
                    
                Case Else
                    GoTo nextItem
            End Select

            li.ListSubItems(2).Text = newText
        End If
nextItem:
    Next i

    If anySelected Then lvListView.SetFocus
done:
    If bCarriedAddedOrRemoved Then
        If bCharLoaded And Not bStartup Then bPromptSave = True
        Call frmMain.RefreshAll
    End If
    Exit Sub
ErrHandler:
    ' Optional: Handle/log error
End Sub


'--- UPDATED: copy selected actions to clipboard; repeats by quantity; ignores CARRIED/blank
'--- UPDATED: copy selected actions to clipboard; repeats by quantity;
'             ignores blank and any CARRIED/STASH (these live in SubItem(10))
'--- UPDATED: copy selected actions to clipboard; repeats by quantity; ignores blank/CARRIED
Public Sub LV_CopySelectedActionsToClipboard(ByRef lvListView As ListView)
    Dim i As Long, k As Long
    Dim li As ListItem
    Dim itemName As String
    Dim cellTxt As String
    Dim act As String
    Dim qty As Long
    Dim cmd As String
    Dim outBuf As String

    If lvListView Is Nothing Then Exit Sub
    If lvListView.ListItems.Count = 0 Then Exit Sub

    For i = 1 To lvListView.ListItems.Count
        If lvListView.ListItems(i).Selected Then
            Set li = lvListView.ListItems(i)
            If li.ListSubItems.Count >= 2 Then
                itemName = Trim$(li.ListSubItems(1).Text)   ' item name
                cellTxt = Trim$(li.ListSubItems(2).Text)     ' action + optional x#
                
                Call ParseActionAndQty(cellTxt, act, qty)    ' qty>=1; act upper-cased
                
                Select Case act
                    Case "DROP":    cmd = "drop " & itemName
                    Case "HIDE":    cmd = "hide " & itemName
                    Case "SELL":    cmd = "sell " & itemName
                    Case "BUY":     cmd = "buy " & itemName
                    Case "PICKUP":  cmd = "get " & itemName
                    Case "STASH":   cmd = "stash " & itemName
                    Case "", "CARRIED", "MANUAL"
                        cmd = ""     ' ignore
                    Case Else
                        cmd = ""     ' unknown tag -> ignore
                End Select

                If LenB(cmd) <> 0 Then
                    For k = 1 To qty
                        If LenB(outBuf) <> 0 Then outBuf = outBuf & vbCrLf
                        outBuf = outBuf & cmd
                    Next k
                End If
            End If
        End If
    Next i

    If LenB(outBuf) <> 0 Then
        Clipboard.clear
        Clipboard.SetText outBuf
    End If
End Sub


'--- helper: parse action + optional trailing quantity "x#"
'--- helper: parse action + optional trailing quantity "x#"
'--- helper: parse action + optional trailing quantity "x#"
Public Sub ParseActionAndQty(ByVal sIn As String, ByRef actionOut As String, ByRef qtyOut As Long)
    Dim s As String, pos As Long, tail As String, maybeNum As String
    s = Trim$(sIn)
    actionOut = UCase$(s)
    qtyOut = 1
    
    If LenB(s) = 0 Then Exit Sub
    
    ' 1) Try the " ... x#" form (with a space)
    pos = InStrRev(s, " ")
    If pos > 0 Then
        tail = Mid$(s, pos + 1)
        If LCase$(Left$(tail, 1)) = "x" And Len(tail) > 1 Then
            maybeNum = Mid$(tail, 2)
            If IsNumeric(maybeNum) Then
                qtyOut = val(maybeNum)
                If qtyOut < 1 Then qtyOut = 1
                actionOut = UCase$(Trim$(Left$(s, pos - 1)))
                Exit Sub
            End If
        End If
    End If
    
    ' 2) Try the "...x#" form (no space)
    Dim j As Long, startDigits As Long
    startDigits = 0
    For j = Len(s) To 1 Step -1
        Dim ch As String
        ch = Mid$(s, j, 1)
        If ch Like "[0-9]" Then
            startDigits = j
        ElseIf (ch = "x" Or ch = "X") And startDigits > 0 And j = startDigits - 1 Then
            maybeNum = Mid$(s, j + 1)
            If IsNumeric(maybeNum) Then
                qtyOut = val(maybeNum)
                If qtyOut < 1 Then qtyOut = 1
                actionOut = UCase$(Trim$(Left$(s, j - 1)))
                Exit Sub
            End If
            Exit For
        Else
            Exit For
        End If
    Next j
End Sub



Public Sub SortListView(ListView As ListView, ByVal Index As Integer, ByVal DataType As ListDataType, ByVal Ascending As Boolean)

'*******************************************************************************
' Sort a ListView by String, Number, or DateTime
'
' Parameters:
'
'   ListView    Reference to the ListView control to be sorted.
'   Index       Index of the column in the ListView to be sorted. The first
'               column in a ListView has an index value of 1.
'   DataType    Sets whether the data in the column is to be sorted
'               alphabetically, numerically, or by date.
'   Ascending   Sets the direction of the sort. True sorts A-Z (Ascending),
'               and False sorts Z-A (descending)
'-------------------------------------------------------------------------------

    On Error Resume Next
    Dim i As Integer
    Dim L As Long
    Dim strFormat As String
    
    ' Display the hourglass cursor whilst sorting
    
    Dim lngCursor As Long
    lngCursor = ListView.MousePointer
    ListView.MousePointer = vbHourglass
    
    ' Prevent the ListView control from updating on screen - this is to hide
    ' the changes being made to the listitems, and also to speed up the sort
    
    If ListView.ListItems.Count > 75 Then LockWindowUpdate frmMain.hWnd 'ListView.hWnd
    
    Dim blnRestoreFromTag As Boolean
    
    Select Case DataType
    Case ldtstring
        
        ' Sort alphabetically. This is the only sort provided by the
        ' MS ListView control (at this time), and as such we don't really
        ' need to do much here
    
        blnRestoreFromTag = False
        
    Case ldtnumber
    
        ' Sort Numerically
    
        strFormat = String$(20, "0") & "." & String$(10, "0")
        
        ' Loop through the values in this column. Re-format the values so
        ' as they can be sorted alphabetically, having already stored their
        ' text values in the tag, along with the tag's original value
    
        With ListView.ListItems
            If (Index = 1) Then
                For L = 1 To .Count
                    With .item(L)
                        .Tag = .Text & Chr$(0) & .Tag
'                        If IsNumeric(.Text) Then
                            If CDbl(val(.Text)) >= 0 Then
                                .Text = Format(CDbl(val(.Text)), strFormat)
                            Else
                                .Text = "&" & InvNumber(Format(0 - CDbl(val(.Text)), strFormat))
                            End If
'                        Else
'                            .Text = ""
'                        End If
                    End With
                Next L
            Else
                For L = 1 To .Count
                    With .item(L).ListSubItems(Index - 1)
                        .Tag = .Text & Chr$(0) & .Tag
'                        If IsNumeric(.Text) Then
                            If CDbl(val(.Text)) >= 0 Then
                                .Text = Format(CDbl(val(.Text)), strFormat)
                            Else
                                .Text = "&" & InvNumber(Format(0 - CDbl(val(.Text)), strFormat))
                            End If
'                        Else
'                            .Text = ""
'                        End If
                    End With
                Next L
            End If
        End With
        
        blnRestoreFromTag = True
    
    Case ldtDateTime
    
        ' Sort by date.
        
        strFormat = "YYYYMMDDHhNnSs"
        
        Dim dte As Date
    
        ' Loop through the values in this column. Re-format the dates so as they
        ' can be sorted alphabetically, having already stored their visible
        ' values in the tag, along with the tag's original value
    
        With ListView.ListItems
            If (Index = 1) Then
                For L = 1 To .Count
                    With .item(L)
                        .Tag = .Text & Chr$(0) & .Tag
                        dte = CDate(.Text)
                        .Text = Format$(dte, strFormat)
                    End With
                Next L
            Else
                For L = 1 To .Count
                    With .item(L).ListSubItems(Index - 1)
                        .Tag = .Text & Chr$(0) & .Tag
                        dte = CDate(.Text)
                        .Text = Format$(dte, strFormat)
                    End With
                Next L
            End If
        End With
        
        blnRestoreFromTag = True
        
    End Select
    
    ' Sort the ListView Alphabetically
    
    ListView.SortOrder = IIf(Ascending, lvwAscending, lvwDescending)
    ListView.SortKey = Index - 1
    ListView.Sorted = True
    
    ' Restore the Text Values if required
    
    If blnRestoreFromTag Then
        
        ' Restore the previous values to the 'cells' in this column of the list
        ' from the tags, and also restore the tags to their original values
        
        With ListView.ListItems
            If (Index = 1) Then
                For L = 1 To .Count
                    With .item(L)
                        i = InStr(.Tag, Chr$(0))
                        .Text = Left$(.Tag, i - 1)
                        .Tag = Mid$(.Tag, i + 1)
                    End With
                Next L
            Else
                For L = 1 To .Count
                    With .item(L).ListSubItems(Index - 1)
                        i = InStr(.Tag, Chr$(0))
                        .Text = Left$(.Tag, i - 1)
                        .Tag = Mid$(.Tag, i + 1)
                    End With
                Next L
            End If
        End With
    End If
    
    ' Unlock the list window so that the OCX can update it
    
    LockWindowUpdate 0&
    
    ' Restore the previous cursor
    
    ListView.MousePointer = lngCursor
    

End Sub

Public Sub SortListViewByTag(ListView As ListView, ByVal Index As Integer, ByVal DataType As ListDataType, ByVal Ascending As Boolean)

'*******************************************************************************
' Sort a ListView by String, Number, or DateTime
'
' Parameters:
'
'   ListView    Reference to the ListView control to be sorted.
'   Index       Index of the column in the ListView to be sorted. The first
'               column in a ListView has an index value of 1.
'   DataType    Sets whether the data in the column is to be sorted
'               alphabetically, numerically, or by date.
'   Ascending   Sets the direction of the sort. True sorts A-Z (Ascending),
'               and False sorts Z-A (descending)
'-------------------------------------------------------------------------------

    On Error Resume Next
    Dim i As Integer
    Dim L As Long
    Dim strFormat As String
    
    ' Display the hourglass cursor whilst sorting
    
    Dim lngCursor As Long
    lngCursor = ListView.MousePointer
    ListView.MousePointer = vbHourglass
    
    ' Prevent the ListView control from updating on screen - this is to hide
    ' the changes being made to the listitems, and also to speed up the sort
    
    If ListView.ListItems.Count > 75 Then LockWindowUpdate frmMain.hWnd 'ListView.hWnd
    
    Dim blnRestoreFromTag As Boolean
    
    Select Case DataType
    Case ldtstring
        
        ' Sort alphabetically. This is the only sort provided by the
        ' MS ListView control (at this time), and as such we don't really
        ' need to do much here
    
        blnRestoreFromTag = False
        
    Case ldtnumber
    
        ' Sort Numerically
    
        strFormat = String$(20, "0") & "." & String$(10, "0")
        
        ' Loop through the values in this column. Re-format the values so
        ' as they can be sorted alphabetically, having already stored their
        ' text values in the tag, along with the tag's original value
    
        With ListView.ListItems
            If (Index = 1) Then
                For L = 1 To .Count
                    With .item(L)
                        '.Tag = .Text & Chr$(0) & .Tag
                        .Tag = .Tag & Chr$(0) & .Text
'                        If IsNumeric(.Text) Then
                            If CDbl(val(Replace(.Text, "%", ""))) >= 0 Then
                                .Text = Format(CDbl(val(.Tag)), strFormat)
                            Else
                                .Text = "&" & InvNumber(Format(0 - CDbl(val(.Tag)), strFormat))
                            End If
'                        Else
'                            .Text = ""
'                        End If
                    End With
                Next L
            Else
                For L = 1 To .Count
                    With .item(L).ListSubItems(Index - 1)
                        '.Tag = .Text & Chr$(0) & .Tag
                        .Tag = .Tag & Chr$(0) & .Text
'                        If IsNumeric(.Text) Then
                            If CDbl(val(Replace(.Text, "%", ""))) >= 0 Then
                                .Text = Format(CDbl(val(.Tag)), strFormat)
                            Else
                                .Text = "&" & InvNumber(Format(0 - CDbl(val(.Tag)), strFormat))
                            End If
'                        Else
'                            .Text = ""
'                        End If
                    End With
                Next L
            End If
        End With
        
        blnRestoreFromTag = True
    
    Case ldtDateTime
    
        ' Sort by date.
        
        strFormat = "YYYYMMDDHhNnSs"
        
        Dim dte As Date
    
        ' Loop through the values in this column. Re-format the dates so as they
        ' can be sorted alphabetically, having already stored their visible
        ' values in the tag, along with the tag's original value
    
        With ListView.ListItems
            If (Index = 1) Then
                For L = 1 To .Count
                    With .item(L)
                        .Tag = .Text & Chr$(0) & .Tag
                        dte = CDate(.Text)
                        .Text = Format$(dte, strFormat)
                    End With
                Next L
            Else
                For L = 1 To .Count
                    With .item(L).ListSubItems(Index - 1)
                        .Tag = .Text & Chr$(0) & .Tag
                        dte = CDate(.Text)
                        .Text = Format$(dte, strFormat)
                    End With
                Next L
            End If
        End With
        
        blnRestoreFromTag = True
        
    End Select
    
    ' Sort the ListView Alphabetically
    
    ListView.SortOrder = IIf(Ascending, lvwAscending, lvwDescending)
    ListView.SortKey = Index - 1
    ListView.Sorted = True
    
    ' Restore the Text Values if required
    
    If blnRestoreFromTag Then
        
        ' Restore the previous values to the 'cells' in this column of the list
        ' from the tags, and also restore the tags to their original values
        
        With ListView.ListItems
            If (Index = 1) Then
                For L = 1 To .Count
                    With .item(L)
                        i = InStr(.Tag, Chr$(0))
                        .Text = Mid$(.Tag, i + 1)
                        .Tag = Left$(.Tag, i - 1)
                    End With
                Next L
            Else
                For L = 1 To .Count
                    With .item(L).ListSubItems(Index - 1)
                        i = InStr(.Tag, Chr$(0))
                        .Text = Mid$(.Tag, i + 1)
                        .Tag = Left$(.Tag, i - 1)
                    End With
                Next L
            End If
        End With
    End If
    
    ' Unlock the list window so that the OCX can update it
    
    LockWindowUpdate 0&
    
    ' Restore the previous cursor
    
    ListView.MousePointer = lngCursor
    

End Sub

Public Function InvNumber(ByVal Number As String) As String
'*******************************************************************************
' Modifies a numeric string to allow it to be sorted alphabetically
'-------------------------------------------------------------------------------

    Static i As Integer
    For i = 1 To Len(Number)
        Select Case Mid$(Number, i, 1)
        Case "-": Mid$(Number, i, 1) = " "
        Case "0": Mid$(Number, i, 1) = "9"
        Case "1": Mid$(Number, i, 1) = "8"
        Case "2": Mid$(Number, i, 1) = "7"
        Case "3": Mid$(Number, i, 1) = "6"
        Case "4": Mid$(Number, i, 1) = "5"
        Case "5": Mid$(Number, i, 1) = "4"
        Case "6": Mid$(Number, i, 1) = "3"
        Case "7": Mid$(Number, i, 1) = "2"
        Case "8": Mid$(Number, i, 1) = "1"
        Case "9": Mid$(Number, i, 1) = "0"
        End Select
    Next
    InvNumber = Number
    
'*******************************************************************************
'
'-------------------------------------------------------------------------------
End Function

'--- read/write state (non-destructive on existing Tag) ---
Private Function LV_ReadSortState(ByVal lv As ListView, ByRef st As tSortState) As Boolean
    Dim parts() As String, i As Long, s As String
    s = CStr(lv.Tag)
    If LenB(s) = 0 Then Exit Function

    parts = Split(s, SEP)
    For i = LBound(parts) To UBound(parts)
        If Left$(parts(i), Len(SORT_TOKEN)) = SORT_TOKEN Then
            Dim body As String, kv() As String, j As Long, piece As String
            body = Mid$(parts(i), Len(SORT_TOKEN) + 1)
            kv = Split(body, ";")
            For j = LBound(kv) To UBound(kv)
                piece = Trim$(kv(j))
                If Left$(piece, 4) = "col=" Then st.LastCol = val(Mid$(piece, 5))
                If Left$(piece, 4) = "asc=" Then st.asc = IIf(val(Mid$(piece, 5)) <> 0, 1, 0)
                If Left$(piece, 6) = "dtype=" Then st.DType = val(Mid$(piece, 7))
                If Left$(piece, 6) = "bytag=" Then st.ByTag = IIf(val(Mid$(piece, 7)) <> 0, 1, 0)
            Next j
            LV_ReadSortState = True
            Exit Function
        End If
    Next i
End Function

Private Sub LV_WriteSortState(ByVal lv As ListView, ByRef st As tSortState)
    Dim s As String, parts() As String, i As Long, found As Boolean
    Dim rebuilt As String, piece As String, newSeg As String

    newSeg = SORT_TOKEN & "col=" & CStr(st.LastCol) _
           & ";asc=" & CStr(st.asc) _
           & ";dtype=" & CStr(st.DType) _
           & ";bytag=" & CStr(st.ByTag)

    s = CStr(lv.Tag)
    If InStr(1, s, SORT_TOKEN, vbBinaryCompare) = 0 Then
        If LenB(s) = 0 Then
            lv.Tag = newSeg
        Else
            lv.Tag = s & SEP & newSeg
        End If
        Exit Sub
    End If

    parts = Split(s, SEP)
    For i = LBound(parts) To UBound(parts)
        piece = parts(i)
        If Left$(piece, Len(SORT_TOKEN)) = SORT_TOKEN Then
            piece = newSeg
            found = True
        End If
        If LenB(rebuilt) = 0 Then
            rebuilt = piece
        Else
            rebuilt = rebuilt & SEP & piece
        End If
    Next i

    If Not found Then
        If LenB(rebuilt) = 0 Then
            rebuilt = newSeg
        Else
            rebuilt = rebuilt & SEP & newSeg
        End If
    End If

    lv.Tag = rebuilt
End Sub

' Decide next Asc/Desc by rules and persist; unchanged from before except we also write dtype/bytag.
Public Function LV_GetNextAscending( _
    ByVal lv As ListView, _
    ByVal col As MSComctlLib.ColumnHeader, _
    ByVal DataType As ListDataType, _
    Optional ByVal textAscFirst As Boolean = True, _
    Optional ByVal numberDescFirst As Boolean = True _
) As Boolean

    Dim st As tSortState, has As Boolean
    has = LV_ReadSortState(lv, st)

    If Not has Or st.LastCol <> col.Index Then
        Select Case DataType
            Case ldtstring: LV_GetNextAscending = textAscFirst
            Case ldtnumber, ldtDateTime: LV_GetNextAscending = Not numberDescFirst ' Desc first => False
            Case Else: LV_GetNextAscending = textAscFirst
        End Select
    Else
        LV_GetNextAscending = (st.asc = 0) ' toggle
    End If

    lv.SortOrder = IIf(LV_GetNextAscending, lvwAscending, lvwDescending)

    st.LastCol = col.Index
    st.asc = IIf(LV_GetNextAscending, 1, 0)
    st.DType = DataType ' store now; caller also sets ByTag in LV_Sort_ColumnClick
    LV_WriteSortState lv, st
End Function

' Wrapper to be called FROM your *_ColumnClick handlers.
' Computes next direction, calls the proper sort routine, and saves full state (incl. ByTag).
Public Sub LV_Sort_ColumnClick( _
    ByVal lv As ListView, _
    ByVal col As MSComctlLib.ColumnHeader, _
    Optional ByVal DataType As ListDataType = ldtstring, _
    Optional ByVal useByTag As Boolean, _
    Optional ByVal forceAsc As Boolean, _
    Optional ByVal forceDesc As Boolean _
)
    Dim bAsc As Boolean
    Dim st As tSortState

    ' Decide direction
    If forceAsc Then
        bAsc = True
    ElseIf forceDesc Then
        bAsc = False
    Else
        bAsc = LV_GetNextAscending(lv, col, DataType)  ' this also sets SortOrder + writes partial state
    End If

    ' Keep the arrow synced even on forced paths
    lv.SortOrder = IIf(bAsc, lvwAscending, lvwDescending)

    ' Perform the actual sort
    If useByTag Then
        SortListViewByTag lv, col.Index, DataType, bAsc
    Else
        SortListView lv, col.Index, DataType, bAsc
    End If

    ' FINALIZE/PERSIST full state for Refresh to use (always set LastCol!)
    Call LV_ReadSortState(lv, st)   ' ok if empty; we'll fill it
    st.LastCol = col.Index          ' <-- critical to avoid drift on refresh
    st.DType = DataType
    st.ByTag = IIf(useByTag, 1, 0)
    st.asc = IIf(bAsc, 1, 0)
    LV_WriteSortState lv, st
End Sub


'==== Replace LV_RefreshSort with this version ====
' Optional defaults are ONLY used when no prior sort has been recorded for this ListView.
' After applying defaults the first time, they are persisted as the ListView's sort state.
Public Sub LV_RefreshSort( _
    ByVal lv As ListView, _
    Optional ByVal defaultCol As Integer = 1, _
    Optional ByVal defaultDType As ListDataType = ldtstring, _
    Optional ByVal defaultByTag As Boolean = False, _
    Optional ByVal defaultAsc As Boolean = True _
)
    Dim st As tSortState
    Dim haveState As Boolean
    Dim asc As Boolean

    haveState = LV_ReadSortState(lv, st) And (st.LastCol > 0)

    If haveState Then
        ' Reapply EXACT previous sort (no toggle)
        lv.SortOrder = IIf(st.asc <> 0, lvwAscending, lvwDescending)
        If st.ByTag <> 0 Then
            SortListViewByTag lv, st.LastCol, st.DType, (st.asc <> 0)
        Else
            SortListView lv, st.LastCol, st.DType, (st.asc <> 0)
        End If
        Exit Sub
    End If

    ' No prior state recorded — only proceed if caller provided a default column
    If defaultCol <= 0 Then Exit Sub
    
    asc = defaultAsc
    
    ' Apply the default sort
    lv.SortOrder = IIf(asc, lvwAscending, lvwDescending)
    If defaultByTag Then
        SortListViewByTag lv, defaultCol, defaultDType, asc
    Else
        SortListView lv, defaultCol, defaultDType, asc
    End If

    ' Persist as the ListView's sort state so future refreshes need no args
    st.LastCol = defaultCol
    st.asc = IIf(asc, 1, 0)
    st.DType = defaultDType
    st.ByTag = IIf(defaultByTag, 1, 0)
    LV_WriteSortState lv, st
End Sub

' Call this instead of LV_Sort_ColumnClick from your ColumnClick handler.
' - If sticky is enabled for this LV, we route to LV_Sort_WithStickyByAction / LV_Sort_WithStickyByTagCol
' - Else (not sticky) we call SortListView directly (so we can honor forceAsc/forceDesc)
Public Sub LV_Sort_ColumnClickOrSticky( _
    ByVal lv As ListView, _
    ByVal ColumnHeader As MSComctlLib.ColumnHeader, _
    ByVal nSortType As ListDataType, _
    ByVal bSortTag As Boolean, _
    Optional ByVal forceAsc As Boolean, _
    Optional ByVal forceDesc As Boolean _
)
    On Error GoTo fail

    ' --- small helper to resolve ASC with overrides ---
    Dim functionDefaultAsc As Boolean
    Dim ResolveAsc As Boolean
    ' if both overrides are set, prefer Asc (and log)
'    If forceAsc And forceDesc Then
'        If DEBUG_STICKY Then Debug.Print "LV_Sort_ColumnClickOrSticky: both forceAsc & forceDesc set; using forceAsc"
'    End If

    Dim orderCSV As String
    Dim stickyOn As Boolean: stickyOn = LV_IsStickyEnabled(lv, orderCSV)

    ' =========================
    ' STICKY + TAG-COLUMN PATH
    ' =========================
    If bSortTag And stickyOn Then
        Dim st As tSortState, have As Boolean
        have = LV_ReadSortState(lv, st)

        ' default (when no overrides): DESC on first TagCol click, else toggle
        If m_LastSortWasTagCol = False Or Not have Or st.LastCol <> ColumnHeader.Index Or st.ByTag = 0 Then
            functionDefaultAsc = False   ' first click in TagCol mode => DESC
        Else
            functionDefaultAsc = (st.asc = 0) ' toggle
        End If

        ' apply overrides
        If forceAsc Xor forceDesc Then
            ResolveAsc = forceAsc
        Else
            ResolveAsc = functionDefaultAsc
        End If

'        If DEBUG_STICKY Then Debug.Print "CLICK-TAGCOL nextAsc=", ResolveAsc

        ' do the sticky sort w/ TagCol secondary
        LV_Sort_WithStickyByTagCol lv, ColumnHeader.Index, nSortType, ResolveAsc, orderCSV

        ' persist state
        st.LastCol = ColumnHeader.Index
        st.asc = IIf(ResolveAsc, 1, 0)
        st.DType = nSortType
        st.ByTag = 1
        LV_WriteSortState lv, st

        m_LastSortWasTagCol = True
        Exit Sub
    End If

    ' =========================================
    ' STICKY + NORMAL COLUMN (secondary = column)
    ' =========================================
    If stickyOn And Not bSortTag Then
        Dim nextAsc2 As Boolean, st2 As tSortState, have2 As Boolean
        have2 = LV_ReadSortState(lv, st2)

        ' default (when no overrides): your existing "next asc" logic
        nextAsc2 = LV_GetNextAscending(lv, ColumnHeader, nSortType)

        ' apply overrides
        If forceAsc Xor forceDesc Then
            nextAsc2 = forceAsc
        End If

        LV_Sort_WithStickyByAction lv, ColumnHeader, nSortType, nextAsc2, orderCSV

        ' persist state
        st2.LastCol = ColumnHeader.Index
        st2.asc = IIf(nextAsc2, 1, 0)
        st2.DType = nSortType
        st2.ByTag = 0
        LV_WriteSortState lv, st2

        m_LastSortWasTagCol = False
        Exit Sub
    End If

    ' =================
    ' NON-STICKY PATHS
    ' =================
    ' Here we can finally honor forceAsc/forceDesc even without sticky by calling SortListView directly.
    Dim plainAsc As Boolean
    If forceAsc Xor forceDesc Then
        plainAsc = forceAsc
        SortListView lv, ColumnHeader.Index, nSortType, plainAsc

        Dim st3 As tSortState
        st3.LastCol = ColumnHeader.Index
        st3.asc = IIf(plainAsc, 1, 0)
        st3.DType = nSortType
        st3.ByTag = IIf(bSortTag, 1, 0)
        LV_WriteSortState lv, st3

        m_LastSortWasTagCol = (bSortTag <> 0)
        Exit Sub
    Else
        ' No overrides -> fall back to your legacy handler
        LV_Sort_ColumnClick lv, ColumnHeader, nSortType, bSortTag
        ' You may want to read back the state here if your legacy writes it;
        ' minimally, clear the tag-mode flag since we don't know.
        m_LastSortWasTagCol = (bSortTag <> 0)
        Exit Sub
    End If

fail:
    ' Last-resort fallback to legacy
    LV_Sort_ColumnClick lv, ColumnHeader, nSortType, bSortTag
End Sub



Public Sub LV_RefreshSort_RespectingSticky(ByVal lv As ListView)
    On Error GoTo fail
    Dim st As tSortState
    Dim orderCSV As String
    If Not LV_ReadSortState(lv, st) Then Exit Sub

    ' If last sort was an explicit Tag sort, keep your legacy path:
    If st.ByTag <> 0 Then
        Call LV_RefreshSort(lv)
        Exit Sub
    End If

    ' If sticky is enabled, re-apply sticky with the stored column & direction
    If LV_IsStickyEnabled(lv, orderCSV) Then
        Dim ch As ColumnHeader
        Set ch = lv.ColumnHeaders(st.LastCol)
        Dim asc As Boolean: asc = (st.asc <> 0)
        LV_Sort_WithStickyByAction lv, ch, st.DType, asc, orderCSV
    Else
        ' otherwise your normal refresh
        Call LV_RefreshSort(lv)
    End If
    Exit Sub
fail:
    Call LV_RefreshSort(lv)
End Sub


' Return index of a second hidden column used to persist a row-sequence tie-break
Private Function LV_EnsureHiddenRowSeqColumn(ByVal lv As ListView) As Long
    Dim idx As Long, i As Long, parts() As String, tagStr As String
    tagStr = CStr(lv.Tag)
    parts = Split(tagStr, vbNullChar)

    ' Reuse token if present
    For i = LBound(parts) To UBound(parts)
        If Left$(parts(i), 13) = "HIDROWSEQCOL:" Then
            idx = val(Mid$(parts(i), 14))
            Exit For
        End If
    Next

    If idx < 1 Or idx > lv.ColumnHeaders.Count Then
        Dim ch As ColumnHeader
        Set ch = lv.ColumnHeaders.Add(, "col_hidden_rowseq_" & (lv.ColumnHeaders.Count + 1), "", 0)
        idx = ch.Index

        ' rewrite Tag tokens (preserve others)
        Dim kept() As String, n As Long: n = -1
        For i = LBound(parts) To UBound(parts)
            If LenB(parts(i)) > 0 And Left$(parts(i), 13) <> "HIDROWSEQCOL:" Then
                n = n + 1: ReDim Preserve kept(0 To n): kept(n) = parts(i)
            End If
        Next
        If n >= 0 Then
            tagStr = Join(kept, vbNullChar) & vbNullChar & "HIDROWSEQCOL:" & CStr(idx)
        Else
            tagStr = "HIDROWSEQCOL:" & CStr(idx)
        End If
        lv.Tag = tagStr
    End If

    On Error Resume Next
    With lv.ColumnHeaders(idx)
        If .SubItemIndex <> idx - 1 Then .SubItemIndex = idx - 1
        If .Width <> 0 Then .Width = 0
    End With
    On Error GoTo 0

    LV_EnsureHiddenRowSeqColumn = idx
End Function

' Read the row-seq value from the hidden column; 0 if not set
Private Function LV_GetRowSeq(ByRef li As ListItem, ByVal rowSeqCol As Long) As Long
    If rowSeqCol <= 1 Then
        LV_GetRowSeq = val(CStr(li.Tag)) ' (never used; we don't store here)
    Else
        EnsureSubItemExists li, rowSeqCol
        LV_GetRowSeq = val(CStr(li.ListSubItems(rowSeqCol - 1).Text))
    End If
End Function

' Assign a new sequence if missing (monotonic per ListView)
Public Sub LV_AssignRowSeqIfMissing(ByVal lv As ListView, ByRef li As ListItem)
    Dim rowCol As Long: rowCol = LV_EnsureHiddenRowSeqColumn(lv)
    EnsureSubItemExists li, rowCol
    If LenB(li.ListSubItems(rowCol - 1).Text) = 0 Or val(li.ListSubItems(rowCol - 1).Text) = 0 Then
        ' find current max seq
        Dim i As Long, cur As Long, mx As Long, li2 As ListItem
        For i = 1 To lv.ListItems.Count
            Set li2 = lv.ListItems(i)
            cur = LV_GetRowSeq(li2, rowCol)
            If cur > mx Then mx = cur
        Next
        li.ListSubItems(rowCol - 1).Text = CStr(mx + 1)
    End If
End Sub


' Determine the *next* asc/desc based on your stored state:
' - If clicking the same column, flip direction
' - If new column, default to Ascending
'Private Function LV_ComputeNextAsc(ByVal lv As ListView, ByVal clickedCol As Long) As Boolean
'    Dim st As tSortState
'    If LV_ReadSortState(lv, st) Then
'        If st.LastCol = clickedCol Then
'            LV_ComputeNextAsc = (st.asc = 0) ' flip
'        Else
'            LV_ComputeNextAsc = True         ' new column -> Asc
'        End If
'    Else
'        LV_ComputeNextAsc = True             ' no prior state -> Asc
'    End If
'End Function

' Enable/disable sticky grouping for this LV.
' Stores a compact token into lv.Tag without disturbing your existing SORTSTATE.
Public Sub LV_EnableSticky(ByVal lv As ListView, ByVal enable As Boolean, _
                           Optional ByVal csvOrder As String = "CARRIED,STASH,MANUAL,BUY,SELL,DROP,HIDE,PICKUP")
    Dim s As String: s = CStr(lv.Tag)
    s = LV_RemoveToken(s, "STICKY2:")
    If enable Then
        If LenB(csvOrder) = 0 Then csvOrder = "CARRIED,STASH,MANUAL,BUY,SELL,DROP,HIDE,PICKUP"
        If LenB(s) > 0 Then s = s & vbNullChar
        s = s & "STICKY2:" & csvOrder
    End If
    lv.Tag = s
End Sub

Public Function LV_IsStickyEnabled(ByVal lv As ListView, Optional ByRef csvOrder As String = "") As Boolean
    Dim s As String: s = CStr(lv.Tag)
    Dim parts() As String, i As Long
    parts = Split(s, vbNullChar)
    For i = LBound(parts) To UBound(parts)
        If Left$(parts(i), 8) = "STICKY2:" Then
            csvOrder = Mid$(parts(i), 9)
            If LenB(csvOrder) = 0 Then csvOrder = "CARRIED,STASH,MANUAL,BUY,SELL,DROP,HIDE,PICKUP"
            LV_IsStickyEnabled = True
            Exit Function
        End If
    Next
    csvOrder = "CARRIED,STASH,MANUAL,BUY,SELL,DROP,HIDE,PICKUP"
End Function

Private Function LV_RemoveToken(ByVal tagStr As String, ByVal prefix As String) As String
    If LenB(tagStr) = 0 Then Exit Function
    Dim parts() As String, i As Long, out() As String, n As Long
    parts = Split(tagStr, vbNullChar)
    ReDim out(0): n = -1
    For i = LBound(parts) To UBound(parts)
        If LenB(parts(i)) > 0 Then
            If Left$(parts(i), Len(prefix)) <> prefix Then
                n = n + 1: ReDim Preserve out(0 To n): out(n) = parts(i)
            End If
        End If
    Next
    If n >= 0 Then LV_RemoveToken = Join(out, vbNullChar)
End Function

' Sort the list by each item's .Tag (string compare), ascending only.
' Technique: swap Text<->Tag, do a normal column-1 string sort, then swap back.
' This avoids relying on project-specific SortListViewByTag behavior.
Public Sub LV_SortByCompositeTagAscending(ByVal lv As ListView)
    On Error GoTo done

    Dim i As Long, li As ListItem
    If lv.ListItems.Count <= 1 Then Exit Sub

    ' 1) Swap Text <-> Tag
    For i = 1 To lv.ListItems.Count
        Set li = lv.ListItems(i)
        Dim tmp As String
        tmp = li.Text
        li.Text = CStr(li.Tag)
        li.Tag = tmp
    Next i

    ' 2) Sort column 1 ascending (plain string)
    SortListView lv, 1, ldtstring, True

    ' 3) Swap back
    For i = 1 To lv.ListItems.Count
        Set li = lv.ListItems(i)
        Dim tmp2 As String
        tmp2 = li.Tag
        li.Tag = li.Text
        li.Text = tmp2
    Next i

done:
End Sub

'=====================================================================
' Sticky-group sort by ListSubItems(2) (action tags like CARRIED/STASH/…)
'   • Rows are grouped in a fixed order by SubItem(2) "base action"
'   • Within each group, rows are sub-sorted by the user-selected column
'   • Works with your existing SortListViewByTag (string compare)
'
' Usage example (in your ColumnClick handler):
'   LV_Sort_WithStickyByAction lv, ColumnHeader, chosenDataType, bAsc, _
'       "CARRIED,STASH,MANUAL,BUY,SELL,DROP,HIDE,PICKUP"
'
' Notes:
'   - We *always* call SortListViewByTag with Ascending:=True.
'     The asc/desc choice for the user column is encoded into the composite Tag key.
'   - Unknown/blank actions fall to the end.
'=====================================================================

Public Sub LV_Sort_WithStickyByAction( _
    ByVal lv As ListView, _
    ByRef col As MSComctlLib.ColumnHeader, _
    ByVal DataType As ListDataType, _
    ByVal bAsc As Boolean, _
    Optional ByVal csvStickyOrder As String = "CARRIED,STASH,MANUAL,BUY,SELL,DROP,HIDE,PICKUP" _
)
    On Error GoTo fail

    ' Priority map for flags in SubItem(2)
    Dim priMap As Object: Set priMap = CreateObject("Scripting.Dictionary")
    priMap.CompareMode = 1 ' vbTextCompare

    Dim parts() As String, i As Long, name As String
    parts = Split(csvStickyOrder, ",")
    For i = LBound(parts) To UBound(parts)
        name = UCase$(Trim$(parts(i)))
        If LenB(name) > 0 Then If Not priMap.Exists(name) Then priMap.Add name, i + 1
    Next

    Const PRI_UNKNOWN As Long = 9998
    Const PRI_BLANK   As Long = 9999
    Dim SEP As String: SEP = Chr$(31)

    ' Pre-scan range for numeric/datetime in the clicked column
    Dim minV As Double, maxV As Double
    If DataType = ldtnumber Or DataType = ldtDateTime Then
        Call GetListViewColMinMax(lv, col.Index, DataType, minV, maxV)
    Else
        minV = 0: maxV = 0
    End If

    ' Use hidden sort column; DO NOT touch li.Tag
    Dim wasCreated As Boolean
    Dim hidCol As Long: hidCol = LV_EnsureHiddenSortColumn(lv, wasCreated)
'    If DEBUG_STICKY Then Debug.Print "STICKY-ACTION: hidCol=" & hidCol & "  clickCol=" & col.Index & "  asc=" & bAsc
    
    Dim rowSeqCol As Long: rowSeqCol = LV_EnsureHiddenRowSeqColumn(lv)
    
    If wasCreated Then
        Dim iPrime As Long, liPrime As ListItem
        lv.ColumnHeaders(hidCol).Width = 1
        For iPrime = 1 To lv.ListItems.Count
            Set liPrime = lv.ListItems(iPrime)
            LV_SetCellText liPrime, hidCol, ""
        Next
        lv.Refresh
    End If


    Dim L As Long, li As ListItem
    Dim baseAction As String, pri As Long
    Dim secKey As String, tie As String
    Dim rawVal As String, composite As String

    For L = 1 To lv.ListItems.Count
        Set li = lv.ListItems(L)

        ' Priority from SubItem(2)
        EnsureSubItemExists li, 2
        baseAction = UCase$(ParseActionBase(Trim$(li.ListSubItems(2).Text)))
'        If LenB(baseAction) = 0 Then
'            pri = PRI_BLANK
'        ElseIf priMap.Exists(baseAction) Then
'            pri = priMap(baseAction)
'        Else
'            pri = PRI_UNKNOWN
'        End If
        pri = LV_GetPriFromLI(lv, li)
        
        ' Secondary from the clicked column
        rawVal = LV_GetCell(li, col.Index)
        secKey = BuildSecondaryKey(rawVal, DataType, bAsc, minV, maxV)  ' uses SanitizeKey/Invert for desc

        ' Stable tie-break
        'tie = Right$("00000000" & CStr(L), 8)
        tie = Right$("00000000" & CStr(LV_GetRowSeq(li, rowSeqCol)), 8)

        ' Composite -> write to hidden column (not Tag)
        composite = Right$("000" & CStr(pri), 3) & SEP & secKey & SEP & tie
        
'        If L <= 10 And DEBUG_STICKY Then Call DBG_PrintRow("PRE-ACT", L, li, pri, secKey, tie, composite, col.Index, hidCol, CStr(li.Tag))

        LV_SetCellText li, hidCol, composite
    Next

    ' Sort by hidden column ascending; desc already encoded in secKey
    Dim oldW As Single: oldW = lv.ColumnHeaders(hidCol).Width
    If oldW = 0 Then lv.ColumnHeaders(hidCol).Width = 1
    
'    If DEBUG_STICKY Then
'        Dim newW As Single: newW = lv.ColumnHeaders(hidCol).Width
'        Debug.Print "SORT-HIDDEN", " idx=" & hidCol, _
'                    " oldW=" & oldW, " newW=" & newW, _
'                    " rows=" & lv.ListItems.Count
'    End If

    DBG_DumpBeforeSort "ACTION", lv, 1, hidCol, 1
    
    SortListView lv, hidCol, ldtstring, True
    
    DBG_DumpAfterSort "ACTION", lv, 1, hidCol, 1
    
    If oldW = 0 Then lv.ColumnHeaders(hidCol).Width = 0
    Exit Sub

fail:
'    If DEBUG_STICKY Then Debug.Print "!! FAIL in", "LV_Sort_WithStickyByAction", _
'        " Err=" & Err.Number & " (" & Err.Description & ")", _
'        " hidCol=" & hidCol & " colCount=" & lv.ColumnHeaders.Count
    Err.clear
    SortListView lv, col.Index, DataType, bAsc

End Sub


' Sticky sort where the *secondary* key is the item's Tag (string/number/datetime).
' Writes composite keys straight into the hidden sort column (no Text/Tag swapping).
Public Sub LV_Sort_WithStickyByTag( _
    ByVal lv As ListView, _
    ByVal DataType As ListDataType, _
    ByVal bAsc As Boolean, _
    Optional ByVal csvStickyOrder As String = "CARRIED,STASH,MANUAL,BUY,SELL,DROP,HIDE,PICKUP" _
)
    On Error GoTo fail

    ' Build action priority map (CARRIED, STASH, ...)
    Dim priMap As Object: Set priMap = CreateObject("Scripting.Dictionary")
    priMap.CompareMode = 1 ' vbTextCompare

    Dim parts() As String, i As Long, name As String
    parts = Split(csvStickyOrder, ",")
    For i = LBound(parts) To UBound(parts)
        name = UCase$(Trim$(parts(i)))
        If LenB(name) > 0 Then If Not priMap.Exists(name) Then priMap.Add name, i + 1
    Next

    Const PRI_UNKNOWN As Long = 9998
    Const PRI_BLANK   As Long = 9999
    Dim SEP As String: SEP = Chr$(31)

    ' Pre-scan min/max of Tag if numeric/datetime (for desc encoding)
    Dim minV As Double, maxV As Double, gotAny As Boolean
    If DataType = ldtnumber Or DataType = ldtDateTime Then
        Dim v As Double, li0 As ListItem
        Dim t As String
        For i = 1 To lv.ListItems.Count
            Set li0 = lv.ListItems(i)
            t = CStr(li0.Tag)
            Select Case DataType
                Case ldtnumber
                    v = val(StripNumFmt(t))
                Case ldtDateTime
                    On Error Resume Next
                    v = CDbl(CDate(t))
                    If Err.Number <> 0 Then v = 0#
                    On Error GoTo 0
            End Select

            If Not gotAny Then
                minV = v
                maxV = v
                gotAny = True
            Else
                If v < minV Then minV = v
                If v > maxV Then maxV = v
            End If
        Next
    End If

    ' Ensure hidden column exists and fill composite keys *into that column*
    Dim hidCol As Long: hidCol = LV_EnsureHiddenSortColumn(lv)
'    If DEBUG_STICKY Then Debug.Print "STICKY-TAG: hidCol=" & hidCol & "  asc=" & bAsc & "  dtype=" & DataType

    Dim L As Long, li As ListItem
    Dim baseAction As String, pri As Long
    Dim secKey As String, tie As String
    Dim rawTag As String

    For L = 1 To lv.ListItems.Count
        Set li = lv.ListItems(L)

        ' Priority from SubItem(2)
        EnsureSubItemExists li, 2
        baseAction = UCase$(ParseActionBase(Trim$(li.ListSubItems(2).Text)))
        If LenB(baseAction) = 0 Then
            pri = PRI_BLANK
        ElseIf priMap.Exists(baseAction) Then
            pri = priMap(baseAction)
        Else
            pri = PRI_UNKNOWN
        End If

        ' Secondary from the item's Tag (string/number/datetime)
        rawTag = CStr(li.ListSubItems(9).Tag)
        secKey = BuildSecondaryKey(rawTag, DataType, bAsc, minV, maxV)
                    
        ' Stable tiebreak
        tie = Right$("00000000" & CStr(L), 8)

'        If L <= 10 And DEBUG_STICKY Then Call DBG_PrintRow("PRE-TAG", L, li, pri, secKey, tie, _
'            Right$("000" & CStr(pri), 3) & Chr$(31) & secKey & Chr$(31) & tie, _
'            1, hidCol, CStr(li.Tag))   ' clickCol=1 (not used here)
        
        ' Write composite directly to hidden column (do NOT touch li.Tag)
        LV_SetCellText li, hidCol, Right$("000" & CStr(pri), 3) & SEP & secKey & SEP & tie
    Next

    ' Sort hidden column ascending (desc already encoded in the key)
    Dim oldW As Single: oldW = lv.ColumnHeaders(hidCol).Width
    If oldW = 0 Then lv.ColumnHeaders(hidCol).Width = 1
    DBG_DumpBeforeSort "TAG", lv, 1, hidCol, 1
    SortListView lv, hidCol, ldtstring, True
    DBG_DumpAfterSort "TAG", lv, 1, hidCol, 10
    If oldW = 0 Then lv.ColumnHeaders(hidCol).Width = 0
    Exit Sub

fail:
    ' If anything happens, fall back to your legacy tag sort
    
    SortListViewByTag lv, 1, DataType, bAsc
End Sub


' Return (and create if needed) a hidden column used solely for internal sorting.
' The column is width 0, stored in lv.Tag as HIDDENSORTCOL:<n> so we reuse it.
Private Function LV_EnsureHiddenSortColumn(ByVal lv As ListView, Optional ByRef wasCreated As Boolean = False) As Long
    Dim tagStr As String, parts() As String, i As Long, idx As Long
    tagStr = CStr(lv.Tag)
    parts = Split(tagStr, vbNullChar)

    ' Try to reuse token
    For i = LBound(parts) To UBound(parts)
        If Left$(parts(i), 14) = "HIDDENSORTCOL:" Then
            idx = val(Mid$(parts(i), 15))
            Exit For
        End If
    Next

    ' Create a new hidden column if missing/out-of-range
    If idx < 1 Or idx > lv.ColumnHeaders.Count Then
        Dim ch As ColumnHeader
        Set ch = lv.ColumnHeaders.Add(, "col_hidden_sort_key_" & (lv.ColumnHeaders.Count + 1), "", 0)
        idx = ch.Index
        wasCreated = True

        ' Rewrite token
        Dim kept() As String, n As Long: n = -1
        For i = LBound(parts) To UBound(parts)
            If LenB(parts(i)) > 0 And Left$(parts(i), 14) <> "HIDDENSORTCOL:" Then
                n = n + 1: ReDim Preserve kept(0 To n): kept(n) = parts(i)
            End If
        Next
        If n >= 0 Then
            tagStr = Join(kept, vbNullChar) & vbNullChar & "HIDDENSORTCOL:" & CStr(idx)
        Else
            tagStr = "HIDDENSORTCOL:" & CStr(idx)
        End If
        lv.Tag = tagStr
    End If

    ' Ensure mapping/hidden
    On Error Resume Next
    With lv.ColumnHeaders(idx)
        If .SubItemIndex <> idx - 1 Then .SubItemIndex = idx - 1
        If .Width <> 0 Then .Width = 0
    End With
    On Error GoTo 0

    LV_EnsureHiddenSortColumn = idx
End Function



' Safe, stable sort on our composite key without touching Text/Tag.
' Writes the composite key (already in li.Tag) into the hidden column's subitem
' and uses your existing SortListView to sort that column ascending (string).
Public Sub LV_SortByCompositeKey_HiddenColumn(ByVal lv As ListView)
    On Error GoTo done
    Dim i As Long, li As ListItem, colIdx As Long
    If lv.ListItems.Count <= 1 Then Exit Sub
    colIdx = LV_EnsureHiddenSortColumn(lv)
    For i = 1 To lv.ListItems.Count
        Set li = lv.ListItems(i)
        LV_SetCellText li, colIdx, CStr(li.Tag) ' write composite key to hidden column
    Next i
    SortListView lv, colIdx, ldtstring, True   ' always ascending; desc is encoded
done:
End Sub

' Set text into any column (1..N). Column 1 = ListItem.Text; >1 = SubItems.
Private Sub LV_SetCellText(ByRef li As ListItem, ByVal colIndex As Integer, ByVal sText As String)
    If colIndex <= 1 Then
        li.Text = sText
    Else
        Do While li.ListSubItems.Count < colIndex - 1
            li.ListSubItems.Add , , ""
        Loop
        li.ListSubItems(colIndex - 1).Text = sText
    End If
End Sub
Private Function ParseActionBase(ByVal s As String) As String
    ' Delegates to ParseActionAndQty to avoid duplicate parsing logic
    Dim act As String, qtyTmp As Long
    Call ParseActionAndQty(s, act, qtyTmp)
    ParseActionBase = act
End Function


' Return text for column N (1-based). Column 1 = ListItem.Text; others = SubItems(N-1)
' Ensure at least N subitems exist
Private Sub EnsureSubItemExists(ByRef li As ListItem, ByVal needIdx As Integer)
    Do While li.ListSubItems.Count < needIdx - 1
        li.ListSubItems.Add , , ""
    Loop
End Sub

Private Function StripNumFmt(ByVal s As String) As String
    Dim t As String
    t = Replace$(s, ",", "")
    t = Replace$(t, "$", "")
    t = Trim$(t)
    StripNumFmt = t
End Function



' Build a string key which sorts in the desired direction *as a string*.
' For numbers/dates, we encode into fixed-width strings; for desc we invert via (max - v).
' For text, we upper-case; for desc we invert lexicographically using a simple wide-char transform.
' Builds a normalized secondary sort key from raw text, given a data type and order.
' - Numeric & Date/Time: converted to a monotonic numeric key and padded via Currency (4dp exact).
' - String: uppercased + trimmed; for descending, invert to preserve ListView ASC sort mechanics.
Private Function BuildSecondaryKey( _
    ByVal s As String, _
    ByVal DataType As ListDataType, _
    ByVal asc As Boolean, _
    ByVal minV As Double, _
    ByVal maxV As Double _
) As String
    Select Case DataType
        Case ldtnumber
            Dim vN As Double
            vN = val(StripNumFmt(s))
            If maxV = minV Then
                BuildSecondaryKey = PadNumberKey(0#)
            ElseIf asc Then
                BuildSecondaryKey = PadNumberKey(vN - minV)
            Else
                BuildSecondaryKey = PadNumberKey(maxV - vN)
            End If

        Case ldtDateTime
            Dim vD As Double
            On Error Resume Next
            vD = CDbl(CDate(s))
            If Err.Number <> 0 Then vD = 0#
            On Error GoTo 0
            If maxV = minV Then
                BuildSecondaryKey = PadNumberKey(0#)
            ElseIf asc Then
                BuildSecondaryKey = PadNumberKey(vD - minV)
            Else
                BuildSecondaryKey = PadNumberKey(maxV - vD)
            End If

        Case Else ' strings
            Dim u As String
            u = UCase$(Trim$(s))
            If asc Then
                BuildSecondaryKey = SanitizeKey(u)
            Else
                BuildSecondaryKey = SanitizeKey(InvertStringForDesc(u))
            End If
    End Select
End Function



Private Function SanitizeKey(ByVal s As String) As String
    ' strip our internal separator and cap length for safety
    s = Replace$(s, Chr$(31), " ")
    If Len(s) > 64 Then s = Left$(s, 64)
    SanitizeKey = s
End Function


' Fixed-width numeric key (left-padded with zeros) for safe string compare
Private Function PadNumberKey(ByVal v As Double) As String
    Dim scaled As Currency
    scaled = v * 10000@     ' 4 dec places
    PadNumberKey = Right$("000000000000000000" & CStr(CDbl(scaled)), 18)
End Function

' Fixed-width numeric key using Variant Decimal (28 digits) to avoid overflow
' Sticky sort where the SECONDARY key is read from the specified column's SubItem.Tag
' tagCol: 1-based column index (e.g., 10 for your "Value by Tag" column)
Public Sub LV_Sort_WithStickyByTagCol( _
    ByVal lv As ListView, _
    ByVal tagCol As Long, _
    ByVal DataType As ListDataType, _
    ByVal bAsc As Boolean, _
    Optional ByVal csvStickyOrder As String = "CARRIED,STASH,MANUAL,BUY,SELL,DROP,HIDE,PICKUP" _
)
    On Error GoTo fail

    ' --- Build priority map for SubItem(2) flags ---
    Dim priMap As Object: Set priMap = CreateObject("Scripting.Dictionary")
    priMap.CompareMode = vbTextCompare
    Dim parts() As String, i As Long, name As String
    parts = Split(csvStickyOrder, ",")
    For i = LBound(parts) To UBound(parts)
        name = UCase$(Trim$(parts(i)))
        If LenB(name) > 0 Then If Not priMap.Exists(name) Then priMap.Add name, i + 1
    Next
    Const PRI_UNKNOWN As Long = 9998
    Const PRI_BLANK   As Long = 9999
    Dim SEP As String: SEP = Chr$(31)

    ' --- Pre-scan min/max for numeric/datetime on the SOURCE: SubItem(tagCol-1).Tag ---
    Dim minV As Double, maxV As Double, gotAny As Boolean, v As Double
    If DataType = ldtnumber Or DataType = ldtDateTime Then
        Dim liScan As ListItem, src As String
        For i = 1 To lv.ListItems.Count
            Set liScan = lv.ListItems(i)
            src = GetCellTagFromColumn(liScan, tagCol)
            Select Case DataType
                Case ldtnumber
                    v = val(StripNumFmt(src))
                Case ldtDateTime
                    On Error Resume Next
                    v = CDbl(CDate(src))
                    If Err.Number <> 0 Then v = 0#
                    On Error GoTo 0
            End Select
            If Not gotAny Then
                minV = v
                maxV = v
                gotAny = True
            Else
                If v < minV Then minV = v
                If v > maxV Then maxV = v
            End If
        Next
    End If

    ' --- Ensure hidden sort column ---
    Dim wasCreated As Boolean
    Dim hidCol As Long: hidCol = LV_EnsureHiddenSortColumn(lv, wasCreated)
'    If DEBUG_STICKY Then Debug.Print "STICKY-TAGCOL: hidCol=" & hidCol & "  tagCol=" & tagCol & "  asc=" & bAsc
    
    Dim rowSeqCol As Long: rowSeqCol = LV_EnsureHiddenRowSeqColumn(lv)

    ' PRIME on first creation so subitems bind immediately
    If wasCreated Then
        Dim iPrime As Long, liPrime As ListItem
        ' make it visible (width=1) while we initialize; some builds ignore width=0 writes
        lv.ColumnHeaders(hidCol).Width = 1
        For iPrime = 1 To lv.ListItems.Count
            Set liPrime = lv.ListItems(iPrime)
            ' ensure subitem exists and set empty (or any stub)
            LV_SetCellText liPrime, hidCol, ""
        Next
        lv.Refresh
    End If

    
    ' --- Build composite into hidden column (do NOT touch li.Tag) ---
    Dim L As Long, li As ListItem
    Dim baseAction As String, pri As Long
    Dim secKey As String, tie As String
    Dim rawTag As String, composite As String

    For L = 1 To lv.ListItems.Count
        Set li = lv.ListItems(L)

        ' Priority by SubItem(2) flag
        EnsureSubItemExists li, 2
        baseAction = UCase$(ParseActionBase(Trim$(li.ListSubItems(2).Text)))
'        If LenB(baseAction) = 0 Then
'            pri = PRI_BLANK
'        ElseIf priMap.Exists(baseAction) Then
'            pri = priMap(baseAction)
'        Else
'            pri = PRI_UNKNOWN
'        End If
        pri = LV_GetPriFromLI(lv, li)
        
        ' Secondary from SubItem(tagCol-1).Tag
        rawTag = GetCellTagFromColumn(li, tagCol)
        secKey = BuildSecondaryKey(rawTag, DataType, bAsc, minV, maxV)

        ' Stable tie
        'tie = Right$("00000000" & CStr(L), 8)
        tie = Right$("00000000" & CStr(LV_GetRowSeq(li, rowSeqCol)), 8)
        
        ' Composite ? hidden column
        composite = Right$("000" & CStr(pri), 3) & SEP & secKey & SEP & tie
        LV_SetCellText li, hidCol, composite

        ' Optional debug:
'        If L <= 10 Then Debug.Print "PRE-TAGCOL #" & L & " PRI=" & pri & " FLAG=" & baseAction & " SRC=[" & rawTag & "] SECKEY=[" & secKey & "]"
    Next

    ' --- Sort hidden column ascending (desc encoded in key) ---
    Dim oldW As Single: oldW = lv.ColumnHeaders(hidCol).Width
    If oldW = 0 Then lv.ColumnHeaders(hidCol).Width = 1
    
'    If DEBUG_STICKY Then
'        Dim newW As Single: newW = lv.ColumnHeaders(hidCol).Width
'        Debug.Print "SORT-HIDDEN", " idx=" & hidCol, _
'                    " oldW=" & oldW, " newW=" & newW, _
'                    " rows=" & lv.ListItems.Count
'    End If

    DBG_DumpBeforeSort "TAGCOL", lv, 1, hidCol, 1
    
    SortListView lv, hidCol, ldtstring, True
    
    DBG_DumpAfterSort "TAGCOL", lv, 1, hidCol, 1
    
    If oldW = 0 Then lv.ColumnHeaders(hidCol).Width = 0
    Exit Sub

fail:
'    If DEBUG_STICKY Then Debug.Print "!! FAIL in", "LV_Sort_WithStickyByTagCol", _
'        " Err=" & Err.Number & " (" & Err.Description & ")", _
'        " hidCol=" & hidCol & " colCount=" & lv.ColumnHeaders.Count
    Err.clear
    
    SortListView lv, tagCol, DataType, bAsc

End Sub

' Read the .Tag from a given column (1-based). For col=1 it returns ListItem.Tag.
Private Function GetCellTagFromColumn(ByRef li As ListItem, ByVal colIndex As Long) As String
    If colIndex <= 1 Then
        GetCellTagFromColumn = CStr(li.Tag)
    Else
        EnsureSubItemExists li, colIndex
        GetCellTagFromColumn = CStr(li.ListSubItems(colIndex - 1).Tag)
    End If
End Function

' Reverse lexicographic for ANSI safely:
' - Upper-case the input BEFORE calling this (you already do).
' - Mirrors A..Z -> Z..A; 0..9 -> 9..0
' - Space -> "~" (tilde, highest printable) so spaces invert correctly
' - Other printable ASCII 33..126 mirrored across 32..126
' - Non-printables normalized to space
Private Function InvertStringForDesc(ByVal s As String) As String
    Dim i As Long, ch As Integer, out As String, c As String * 1
    out = Space$(Len(s))

    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        ch = asc(c)  ' ANSI 0..255

        Select Case ch
            Case 65 To 90  ' A..Z
                Mid$(out, i, 1) = Chr$(asc("Z") - (ch - asc("A")))
            Case 48 To 57   ' 0..9
                Mid$(out, i, 1) = Chr$(asc("9") - (ch - asc("0")))
            Case 32         ' space
                Mid$(out, i, 1) = "~"  ' 126
            Case 33 To 126  ' other printable ASCII, mirror across 32..126
                Mid$(out, i, 1) = Chr$(158 - ch) ' 32+126=158
            Case Else
                Mid$(out, i, 1) = " "            ' normalize non-printables
        End Select
    Next

    InvertStringForDesc = out
End Function


' Scan the chosen column to get min/max for numeric/datetime keys
Private Function GetListViewColMinMax( _
    ByVal lv As ListView, _
    ByVal colIndex As Integer, _
    ByVal DataType As ListDataType, _
    ByRef minV As Double, _
    ByRef maxV As Double _
) As Boolean
    Dim i As Long, li As ListItem, v As Double, gotAny As Boolean
    minV = 0#: maxV = 0#: gotAny = False
    For i = 1 To lv.ListItems.Count
        Set li = lv.ListItems(i)
        Select Case DataType
            Case ldtnumber
                If IsNumeric(LV_GetCell(li, colIndex)) Then
                    v = CDbl(LV_GetCell(li, colIndex))
                Else
                    v = 0#
                End If
            Case ldtDateTime
                On Error Resume Next
                v = CDbl(CDate(LV_GetCell(li, colIndex)))
                If Err.Number <> 0 Then v = 0#
                On Error GoTo 0
            Case Else
                v = 0#
        End Select
        If Not gotAny Then
            minV = v
            maxV = v
            gotAny = True
        Else
            If v < minV Then minV = v
            If v > maxV Then maxV = v
        End If
    Next
    GetListViewColMinMax = gotAny
End Function


