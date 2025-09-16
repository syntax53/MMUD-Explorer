Attribute VB_Name = "modItemParse"
Option Explicit

' ==============================
' Public data shapes
' ==============================
Public Type ItemParseResult
    sEquipped()  As String  ' "Item (Slot)" (deduped later)
    sInventory() As String  ' "name" or "name (N)" consolidated later
    sKeys()      As String  ' keys list items (normalized to singular base name; consolidated)
    sGround()    As String  ' ground items (consolidated across room groups)
End Type

Public Type ItemMatch
    Number       As Long
    name         As String
    ItemType     As Long
    Worn         As Long
    WeaponType   As Long
    ObtainedFrom As String
End Type

Public Type ShopToken
    ShopNumber As Long
    bNObuy     As Boolean  ' True means shop does not sell to player (sell-only)
    bNoSell    As Boolean  ' (unused here, future-safe)
End Type


' ======================================================================
' Room-based ground aggregation (no double-count on repeated searches)
' ======================================================================
Private Type RoomAgg
    key    As String ' room name || exits (both lower)
    names() As String
    counts() As Long
    used   As Long
End Type

' ==============================
' Constants (headers / markers)
' ==============================
Private Const HDR_INV  As String = "You are carrying"
Private Const HDR_KEYS As String = "You have the following keys"
Private Const HDR_NOKEY As String = "You have no keys"
Private Const HDR_NOTC As String = "You notice"

' ==============================
' Public entry
' ==============================
Public Function ParseGameTextInventory(ByVal sInput As String) As ItemParseResult
On Error GoTo fail
    Dim tOut As ItemParseResult
    Dim sNorm As String
    Dim vLines() As String
    Dim i As Long, sLine As String

    sNorm = Replace$(sInput, vbCr, vbNullString)   ' normalize to LF only
    vLines = Split(sNorm, vbLf)

    ' init zero-length arrays
    ReDim tOut.sEquipped(0)
    ReDim tOut.sInventory(0)
    ReDim tOut.sKeys(0)
    ReDim tOut.sGround(0)

    ' ---------- Pass 1: scan lines & collect blobs ----------
    i = 0
    Do While i <= UBound(vLines)
        sLine = Trim$(vLines(i))

        If StartsWithCI(sLine, HDR_INV) Then
            Dim blobI As String
            blobI = CollectInventoryBlob(vLines, i)
            ParseInventoryBlob blobI, tOut

        ElseIf StartsWithCI(sLine, HDR_KEYS) Or StartsWithCI(sLine, HDR_NOKEY) Then
            Dim blobK As String
            blobK = CollectKeysBlob(vLines, i)
            ParseKeysBlob blobK, tOut

        ElseIf StartsWithCI(sLine, HDR_NOTC) Then
            Dim blobN As String
            blobN = CollectNoticeBlob(vLines, i)
            ParseNoticeBlob blobN, tOut
        End If

        i = i + 1
    Loop

    ' ---------- Pass 1b: rebuild GROUND by room (dedupe repeated searches) ----------
    ConsolidateGroundByRoom sNorm, tOut

    ' ---------- Pass 2: consolidate lists ----------
    ConsolidateList tOut.sInventory, False
    ConsolidateList tOut.sGround, False
    ConsolidateList tOut.sKeys, True        ' normalize key names ("... keys" -> "... key"; "golden idol (2)")

    ' ---------- Pass 3: dedupe equipped ----------
    DedupeEquipped tOut.sEquipped

    ParseGameTextInventory = tOut
    Exit Function

fail:
    HandleError "ParseGameTextInventory"
    ParseGameTextInventory = tOut
End Function


' ======================================================================
' Inventory blob collection/parsing
' ======================================================================
Private Function CollectInventoryBlob(ByRef vLines() As String, ByRef i As Long) As String
On Error GoTo fail
    Dim s As String, j As Long, sLine As String
    Dim pos As Long, leftPart As String

    s = vbNullString
    j = i
    Do While j <= UBound(vLines)
        sLine = Trim$(vLines(j))

        If ShouldStopSection(sLine, (j <> i)) Then Exit Do

        ' mid-line header split (e.g., "... stormmetal greataxe You have the following keys: ...")
        pos = FindInlineBoundaryPos(sLine)
        If pos > 1 Then
            leftPart = Trim$(Left$(sLine, pos - 1))
            If LenB(leftPart) > 0 Then
                If LenB(s) > 0 Then s = s & " "
                s = s & leftPart
            End If
            vLines(j) = Mid$(sLine, pos)   ' let main loop see the header on this same line next time
            Exit Do
        End If

        If LenB(s) > 0 Then s = s & " "
        s = s & sLine
        j = j + 1
    Loop

    i = j - 1
    CollectInventoryBlob = s
    Exit Function
fail:
    HandleError "CollectInventoryBlob"
    CollectInventoryBlob = s
End Function

Private Sub ParseInventoryBlob(ByVal sBlob As String, ByRef tOut As ItemParseResult)
On Error GoTo fail
    Dim s As String
    Dim toks() As String, k As Long, it As String

    s = sBlob
    If StartsWithCI(s, HDR_INV) Then s = Trim$(Mid$(s, Len(HDR_INV) + 1))

    s = SqueezeSpaces(Replace$(s, vbLf, " "))
    SplitToArray s, toks

    If (Not Not toks) <> 0 Then
        For k = LBound(toks) To UBound(toks)
            it = Trim$(toks(k))
            If LenB(it) = 0 Then GoTo nxt
            If IsBracketLine(it) Then GoTo nxt
            If IsCashItem(it) Then GoTo nxt

            If IsEquippedItem(it) Then
                AddString tOut.sEquipped, it
            Else
                AddString tOut.sInventory, it
            End If
nxt:
        Next k
    End If
    Exit Sub
fail:
    HandleError "ParseInventoryBlob"
End Sub


' ======================================================================
' Keys blob collection/parsing
' ======================================================================
Private Function CollectKeysBlob(ByRef vLines() As String, ByRef i As Long) As String
On Error GoTo fail
    Dim s As String, j As Long, sLine As String
    Dim pos As Long, leftPart As String

    s = vbNullString
    j = i
    Do While j <= UBound(vLines)
        sLine = Trim$(vLines(j))

        If ShouldStopSection(sLine, (j <> i)) Then Exit Do

        pos = FindInlineBoundaryPos(sLine)
        If pos > 1 Then
            leftPart = Trim$(Left$(sLine, pos - 1))
            If LenB(leftPart) > 0 Then
                If LenB(s) > 0 Then s = s & " "
                s = s & leftPart
            End If
            vLines(j) = Mid$(sLine, pos)
            Exit Do
        End If

        If LenB(s) > 0 Then s = s & " "
        s = s & sLine

        If Right$(sLine, 1) = "." Then
            j = j + 1
            Exit Do
        End If

        j = j + 1
    Loop

    i = j - 1
    CollectKeysBlob = s
    Exit Function
fail:
    HandleError "CollectKeysBlob"
    CollectKeysBlob = s
End Function

Private Sub ParseKeysBlob(ByVal sBlob As String, ByRef tOut As ItemParseResult)
On Error GoTo fail
    Dim s As String
    Dim toks() As String, k As Long, it As String

    s = sBlob

    If StartsWithCI(s, HDR_NOKEY) Then
        Exit Sub
    ElseIf StartsWithCI(s, HDR_KEYS) Then
        s = Trim$(Mid$(s, Len(HDR_KEYS) + 1))
        If Left$(s, 1) = ":" Then s = Trim$(Mid$(s, 2))
    End If

    If InStr(1, LCase$(s), "no keys", vbTextCompare) > 0 Then Exit Sub

    s = RTrim$(s)
    If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)

    s = SqueezeSpaces(Replace$(s, vbLf, " "))

    SplitToArray s, toks
    If (Not Not toks) <> 0 Then
        For k = LBound(toks) To UBound(toks)
            it = Trim$(toks(k))
            If LenB(it) > 0 Then AddString tOut.sKeys, it
        Next k
    End If
    Exit Sub
fail:
    HandleError "ParseKeysBlob"
End Sub


' ======================================================================
' Ground blob collection/parsing (line-wise)
' ======================================================================
Private Function CollectNoticeBlob(ByRef vLines() As String, ByRef i As Long) As String
On Error GoTo fail
    Dim s As String, j As Long, sLine As String
    Dim pos As Long, leftPart As String

    s = vbNullString
    j = i
    Do While j <= UBound(vLines)
        sLine = Trim$(vLines(j))

        If IsBracketLine(sLine) Then j = j + 1: GoTo cont

        If ShouldStopSection(sLine, (j <> i)) Then Exit Do

        pos = FindInlineBoundaryPos(sLine)
        If pos > 1 Then
            leftPart = Trim$(Left$(sLine, pos - 1))
            If LenB(leftPart) > 0 Then
                If LenB(s) > 0 Then s = s & " "
                s = s & leftPart
            End If
            vLines(j) = Mid$(sLine, pos)
            Exit Do
        End If

        If LenB(s) > 0 Then s = s & " "
        s = s & sLine

        If LCase$(Right$(sLine, 5)) = "here." Then
            j = j + 1
            Exit Do
        End If

        j = j + 1
cont:
    Loop

    i = j - 1
    CollectNoticeBlob = s
    Exit Function
fail:
    HandleError "CollectNoticeBlob"
    CollectNoticeBlob = s
End Function

Private Sub ParseNoticeBlob(ByVal sBlob As String, ByRef tOut As ItemParseResult)
On Error GoTo fail
    Dim s As String
    Dim toks() As String, k As Long, it As String

    s = sBlob
    If StartsWithCI(s, HDR_NOTC) Then s = Trim$(Mid$(s, Len(HDR_NOTC) + 1))
    If LCase$(Right$(s, 5)) = "here." Then s = Left$(s, Len(s) - 5)

    ' remove bracket inserts and newlines before splitting
    s = StripBracketedChunks(s)
    s = SqueezeSpaces(Replace$(s, vbLf, " "))

    SplitToArray s, toks
    If (Not Not toks) <> 0 Then
        For k = LBound(toks) To UBound(toks)
            it = Trim$(toks(k))
            If LenB(it) = 0 Then GoTo nxt
            If IsBracketLine(it) Then GoTo nxt
            If IsCashItem(it) Then GoTo nxt
            AddString tOut.sGround, it
nxt:
        Next k
    End If
    Exit Sub
fail:
    HandleError "ParseNoticeBlob"
End Sub

Private Sub ConsolidateGroundByRoom(ByVal sAll As String, ByRef tOut As ItemParseResult)
On Error GoTo fail
    Dim v() As String, i As Long, line As String
    Dim curRoomName As String, curExits As String, curKey As String
    Dim pending As String
    Dim rooms() As RoomAgg, roomUsed As Long

    sAll = Replace$(sAll, vbCr, vbNullString)
    v = Split(sAll, vbLf)

    curRoomName = "": curExits = "": curKey = "": pending = ""

    For i = 0 To UBound(v)
        line = Trim$(v(i))
        If LenB(line) = 0 Then GoTo cont

        ' movement commands => start of a new room likely; flush pending to unknown
        If IsMovementCommand(line) Then
            If LenB(pending) > 0 Then
                AddNoticeToRoom rooms, roomUsed, LCase$("__unknown||__unknown"), pending
                pending = ""
            End If
            curRoomName = "": curExits = "": curKey = ""
            GoTo cont
        End If

        ' candidate room name (not a header, not a marker line)
        If Not StartsWithCI(line, HDR_NOTC) _
        And Not StartsWithCI(line, "also here:") _
        And InStr(1, LCase$(line), "obvious exits:", vbTextCompare) = 0 _
        And Left$(line, 1) <> "[" _
        And Not StartsWithCI(line, "name:") _
        And Not StartsWithCI(line, "race:") _
        And Not StartsWithCI(line, "class:") _
        And Not StartsWithCI(line, HDR_INV) _
        And Not StartsWithCI(line, HDR_KEYS) _
        And Not StartsWithCI(line, HDR_NOKEY) Then
            curRoomName = line
        End If

        If StartsWithCI(line, HDR_NOTC) Then
            ' collect full "You notice ... here." span (may wrap)
            Dim span As String
            span = line
            Do While i < UBound(v)
                If LCase$(Right$(Trim$(v(i)), 5)) = "here." Then Exit Do
                i = i + 1
                span = span & " " & Trim$(v(i))
            Loop
            pending = SqueezeSpaces(StripBracketedChunks(span))
            GoTo cont
        End If

        If InStr(1, LCase$(line), "obvious exits:", vbTextCompare) > 0 Then
            curExits = line
            curKey = LCase$(Trim$(curRoomName) & "||" & Trim$(curExits))
            If LenB(pending) > 0 Then
                AddNoticeToRoom rooms, roomUsed, curKey, pending
                pending = ""
            End If
        End If

cont:
    Next i

    ' if something pending without exits, attach to unknown bucket
    If LenB(pending) > 0 Then
        AddNoticeToRoom rooms, roomUsed, LCase$("__unknown||__unknown"), pending
        pending = ""
    End If

    ' merge per-room (MAX within room), then SUM across rooms
    Dim totN() As String, totC() As Long, totU As Long
    Dim r As Long, n As Long

    For r = 0 To roomUsed - 1
        For n = 0 To rooms(r).used - 1
            AccumulateSum totN, totC, totU, rooms(r).names(n), rooms(r).counts(n)
        Next n
    Next r

    ' write back to tOut.sGround as "name (N)"
    If totU = 0 Then
        ReDim tOut.sGround(0)
        tOut.sGround(0) = vbNullString
    Else
        ReDim tOut.sGround(totU - 1)
        For n = 0 To totU - 1
            If totC(n) > 1 Then
                tOut.sGround(n) = totN(n) & " (" & CStr(totC(n)) & ")"
            Else
                tOut.sGround(n) = totN(n)
            End If
        Next n
    End If
    Exit Sub
fail:
    HandleError "ConsolidateGroundByRoom"
End Sub

Private Sub AddNoticeToRoom(ByRef rooms() As RoomAgg, ByRef used As Long, ByVal key As String, ByVal span As String)
    Dim idx As Long
    idx = GetOrCreateRoomIndex(rooms, used, key)

    ' strip header/tail & split
    If StartsWithCI(span, HDR_NOTC) Then span = Trim$(Mid$(span, Len(HDR_NOTC) + 1))
    If LCase$(Right$(span, 5)) = "here." Then span = Left$(span, Len(span) - 5)
    span = SqueezeSpaces(Replace$(StripBracketedChunks(span), vbLf, " "))

    Dim toks() As String, i As Long, nm As String, qty As Long, t As String
    SplitToArray span, toks
    If (Not Not toks) = 0 Then Exit Sub

    For i = LBound(toks) To UBound(toks)
        t = Trim$(toks(i))
        If LenB(t) = 0 Then GoTo nxt
        If IsCashItem(t) Then GoTo nxt
        ParseNameAndQty t, nm, qty
        nm = LCase$(nm)
        If qty < 1 Then qty = 1
        AccumulateMax rooms(idx).names, rooms(idx).counts, rooms(idx).used, nm, qty
nxt:
    Next i
End Sub

Private Function GetOrCreateRoomIndex(ByRef rooms() As RoomAgg, ByRef used As Long, ByVal key As String) As Long
    Dim i As Long
    For i = 0 To used - 1
        If rooms(i).key = key Then GetOrCreateRoomIndex = i: Exit Function
    Next
    If used = 0 Then
        ReDim rooms(0)
    Else
        ReDim Preserve rooms(used)
    End If
    rooms(used).key = key
    rooms(used).used = 0
    GetOrCreateRoomIndex = used
    used = used + 1
End Function

Private Sub AccumulateMax(ByRef names() As String, ByRef counts() As Long, ByRef used As Long, ByVal nm As String, ByVal qty As Long)
    Dim i As Long
    For i = 0 To used - 1
        If names(i) = nm Then
            If qty > counts(i) Then counts(i) = qty
            Exit Sub
        End If
    Next
    If used = 0 Then
        ReDim names(0): ReDim counts(0)
    Else
        ReDim Preserve names(used): ReDim Preserve counts(used)
    End If
    names(used) = nm
    counts(used) = qty
    used = used + 1
End Sub

Private Sub AccumulateSum(ByRef names() As String, ByRef counts() As Long, ByRef used As Long, ByVal nm As String, ByVal qty As Long)
    Dim i As Long
    For i = 0 To used - 1
        If names(i) = nm Then counts(i) = counts(i) + qty: Exit Sub
    Next
    If used = 0 Then
        ReDim names(0): ReDim counts(0)
    Else
        ReDim Preserve names(used): ReDim Preserve counts(used)
    End If
    names(used) = nm
    counts(used) = qty
    used = used + 1
End Sub

Private Function IsMovementCommand(ByVal line As String) As Boolean
    Dim s As String
    s = LCase$(Trim$(line))
    ' allow prompts like "[HP=...]:n"
    Dim p As Long
    p = InStr(s, "]:")
    If p > 0 Then s = Mid$(s, p + 2)
    s = Trim$(s)

    Select Case s
        Case "n", "s", "e", "w", "nw", "ne", "sw", "se", _
             "north", "south", "east", "west", _
             "northwest", "northeast", "southwest", "southeast"
            IsMovementCommand = True
        Case Else
            IsMovementCommand = False
    End Select
End Function


' ======================================================================
' Section boundary detection
' ======================================================================
Private Function IsSectionBoundary(ByVal sLine As String) As Boolean
    Dim sL As String
    sL = LCase$(Trim$(sLine))

    If LenB(sL) = 0 Then IsSectionBoundary = True: Exit Function

    If StartsWithCI(sL, LCase$(HDR_INV)) Then IsSectionBoundary = True: Exit Function
    If StartsWithCI(sL, LCase$(HDR_KEYS)) Then IsSectionBoundary = True: Exit Function
    If StartsWithCI(sL, LCase$(HDR_NOKEY)) Then IsSectionBoundary = True: Exit Function
    If StartsWithCI(sL, LCase$(HDR_NOTC)) Then IsSectionBoundary = True: Exit Function

    If StartsWithCI(sL, "wealth:") Then IsSectionBoundary = True: Exit Function
    If StartsWithCI(sL, "encumbrance:") Then IsSectionBoundary = True: Exit Function

    If StartsWithCI(sL, "name:") Then IsSectionBoundary = True: Exit Function
    If StartsWithCI(sL, "race:") Then IsSectionBoundary = True: Exit Function
    If StartsWithCI(sL, "class:") Then IsSectionBoundary = True: Exit Function

    If Left$(sL, 1) = "[" Then IsSectionBoundary = True: Exit Function
    If InStr(1, sL, "obvious exits:", vbTextCompare) > 0 Then IsSectionBoundary = True: Exit Function
    If StartsWithCI(sL, "also here:") Then IsSectionBoundary = True: Exit Function

    If InStr(1, sL, "*combat", vbTextCompare) > 0 Then IsSectionBoundary = True: Exit Function
End Function

Private Function ShouldStopSection(ByVal sLine As String, ByVal bPastFirst As Boolean) As Boolean
    ShouldStopSection = (bPastFirst And IsSectionBoundary(sLine))
End Function

' return earliest position (>1) where any known header appears mid-line
Private Function FindInlineBoundaryPos(ByVal sLine As String) As Long
    Dim sL As String, p As Long, best As Long
    sL = LCase$(sLine): best = 0

    p = InStr(1, sL, LCase$(HDR_KEYS)):  If p > 1 Then If best = 0 Or p < best Then best = p
    p = InStr(1, sL, LCase$(HDR_NOKEY)): If p > 1 Then If best = 0 Or p < best Then best = p
    p = InStr(1, sL, LCase$(HDR_INV)):   If p > 1 Then If best = 0 Or p < best Then best = p
    p = InStr(1, sL, LCase$(HDR_NOTC)):  If p > 1 Then If best = 0 Or p < best Then best = p
    p = InStr(1, sL, "wealth:"):         If p > 1 Then If best = 0 Or p < best Then best = p
    p = InStr(1, sL, "encumbrance:"):    If p > 1 Then If best = 0 Or p < best Then best = p

    FindInlineBoundaryPos = best
End Function

Private Function StartsWithCI(ByVal s As String, ByVal prefix As String) As Boolean
    StartsWithCI = (LCase$(Left$(s, Len(prefix))) = LCase$(prefix))
End Function


' ======================================================================
' Consolidation helpers (counts, plural/singular, formatting)
' ======================================================================
Private Sub ConsolidateList(ByRef a() As String, ByVal bIsKeysList As Boolean)
On Error GoTo fail
    Dim keys() As String, counts() As Long, nKeys As Long
    Dim i As Long, nm As String, cnt As Long, idx As Long

    If Not SafeArrayHasData(a) Then Exit Sub

    nKeys = 0
    For i = LBound(a) To UBound(a)
        If LenB(a(i)) > 0 Then
            ParseCountAndName a(i), nm, cnt, bIsKeysList
            idx = FindKeyIndex(keys, nKeys, nm)
            If idx < 0 Then
                If nKeys = 0 Then
                    ReDim keys(0): ReDim counts(0)
                Else
                    ReDim Preserve keys(nKeys): ReDim Preserve counts(nKeys)
                End If
                keys(nKeys) = nm
                counts(nKeys) = cnt
                nKeys = nKeys + 1
            Else
                counts(idx) = counts(idx) + cnt
            End If
        End If
    Next i

    If nKeys = 0 Then
        ReDim a(0): a(0) = vbNullString
        Exit Sub
    End If

    ReDim a(nKeys - 1)
    For i = 0 To nKeys - 1
        If counts(i) > 1 Then
            a(i) = keys(i) & " (" & CStr(counts(i)) & ")"
        Else
            a(i) = keys(i)
        End If
    Next i
    Exit Sub
fail:
    HandleError "ConsolidateList"
End Sub

Private Sub ParseCountAndName(ByVal token As String, ByRef nameOut As String, ByRef countOut As Long, ByVal bIsKeysList As Boolean)
    Dim s As String, p As Long, firstWord As String
    s = Trim$(token)

    Do While Len(s) > 0 And (Right$(s, 1) = "." Or Right$(s, 1) = ";")
        s = Left$(s, Len(s) - 1)
        s = RTrim$(s)
    Loop

    p = InStr(1, s, " ")
    If p > 0 Then
        firstWord = Left$(s, p - 1)
        If IsNumeric(firstWord) Then
            countOut = CLng(val(firstWord))
            nameOut = Trim$(Mid$(s, p + 1))
        Else
            countOut = 1
            nameOut = s
        End If
    Else
        countOut = 1
        nameOut = s
    End If

    nameOut = SqueezeSpaces(nameOut)
    nameOut = Trim$(nameOut)

    If bIsKeysList Then
        ' merge '... keys' with '... key'
        If LCase$(Right$(nameOut, 5)) = " keys" Then nameOut = Left$(nameOut, Len(nameOut) - 1)
        ' ensure singular base name when count > 1 (DB uses singular item names)
        If countOut > 1 Then nameOut = SingularizeSimple(nameOut)
    End If

    nameOut = LCase$(nameOut)
End Sub

Private Function SingularizeSimple(ByVal s As String) As String
    Dim L As Long
    s = Trim$(s): L = Len(s)
    If L = 0 Then SingularizeSimple = s: Exit Function

    If L >= 3 And Right$(s, 3) = "ies" Then
        SingularizeSimple = Left$(s, L - 3) & "y"
        Exit Function
    End If

    If L >= 2 And Right$(s, 2) = "es" Then
        If Right$(s, 4) = "sses" Or Right$(s, 4) = "ches" Or Right$(s, 4) = "shes" Or Right$(s, 4) = "xes" Then
            SingularizeSimple = Left$(s, L - 2)
            Exit Function
        End If
    End If

    If L >= 2 And Right$(s, 1) = "s" And Right$(s, 2) <> "ss" Then
        SingularizeSimple = Left$(s, L - 1)
    Else
        SingularizeSimple = s
    End If
End Function

Private Function FindKeyIndex(ByRef keys() As String, ByVal used As Long, ByVal nm As String) As Long
    Dim i As Long
    For i = 0 To used - 1
        If keys(i) = nm Then FindKeyIndex = i: Exit Function
    Next
    FindKeyIndex = -1
End Function

Private Sub DedupeEquipped(ByRef a() As String)
On Error GoTo fail
    Dim seen() As String, out() As String
    Dim used As Long, oUsed As Long
    Dim i As Long, keyL As String

    If Not SafeArrayHasData(a) Then Exit Sub
    used = 0: oUsed = 0
    For i = LBound(a) To UBound(a)
        keyL = LCase$(Trim$(a(i)))
        If LenB(keyL) = 0 Then GoTo nxt
        If IndexOf(seen, used, keyL) < 0 Then
            If used = 0 Then ReDim seen(0) Else ReDim Preserve seen(used)
            seen(used) = keyL: used = used + 1

            If oUsed = 0 Then ReDim out(0) Else ReDim Preserve out(oUsed)
            out(oUsed) = Trim$(a(i)): oUsed = oUsed + 1
        End If
nxt:
    Next i
    If oUsed = 0 Then
        ReDim a(0): a(0) = vbNullString
    Else
        a = out
    End If
    Exit Sub
fail:
    HandleError "DedupeEquipped"
End Sub

Private Function IndexOf(ByRef arr() As String, ByVal used As Long, ByVal keyL As String) As Long
    Dim i As Long
    For i = 0 To used - 1
        If arr(i) = keyL Then IndexOf = i: Exit Function
    Next
    IndexOf = -1
End Function


' ======================================================================
' Populate ListView (QTY column + best shop via GetItemValue)
' ======================================================================
Public Sub PopulateItemManagerFromParsed(ByRef tItems As ItemParseResult, ByRef lvItemManager As ListView)
On Error GoTo fail
    Dim doEquipped As Boolean, doKeys As Boolean
    Dim haveEquipped As Boolean, haveKeys As Boolean, haveInv As Boolean, haveGround As Boolean
    Dim totalToAdd As Long

    haveEquipped = SafeArrayHasData(tItems.sEquipped)
    haveKeys = SafeArrayHasData(tItems.sKeys)
    haveInv = SafeArrayHasData(tItems.sInventory)
    haveGround = SafeArrayHasData(tItems.sGround)

    doEquipped = True
    doKeys = True

    If haveEquipped Then
        If MsgBox("Import EQUIPPED items into Item Manager?", vbQuestion Or vbYesNo, "Import Equipped?") = vbNo Then doEquipped = False
    End If
    If haveKeys Then
        If MsgBox("Import KEYS (from inventory) into Item Manager?", vbQuestion Or vbYesNo, "Import Keys?") = vbNo Then doKeys = False
    End If

    totalToAdd = 0
    If doEquipped Then totalToAdd = totalToAdd + ArrayCount(tItems.sEquipped)
    If doKeys Then totalToAdd = totalToAdd + ArrayCount(tItems.sKeys)
    If haveGround Then totalToAdd = totalToAdd + ArrayCount(tItems.sGround)
    If haveInv Then totalToAdd = totalToAdd + ArrayCount(tItems.sInventory)

    If totalToAdd > 0 And lvItemManager.ListItems.Count > 0 Then
        If MsgBox("Clear the existing Item Manager list first?", vbQuestion Or vbYesNo, "Clear Existing?") = vbYes Then
            lvItemManager.ListItems.clear
        End If
    End If

    If doEquipped And haveEquipped Then AddSectionItems tItems.sEquipped, "Equipped", lvItemManager, True
    If doKeys And haveKeys Then AddSectionItems tItems.sKeys, "Inventory", lvItemManager, False
    If haveGround Then AddSectionItems tItems.sGround, "Ground", lvItemManager, False
    If haveInv Then AddSectionItems tItems.sInventory, "Inventory", lvItemManager, False
    Exit Sub
fail:
    MsgBox "PopulateItemManagerFromParsed error: " & Err.Description, vbExclamation
End Sub

Private Function ArrayCount(ByRef a() As String) As Long
    If SafeArrayHasData(a) Then ArrayCount = UBound(a) - LBound(a) + 1 Else ArrayCount = 0
End Function

' sectionName: "Equipped", "Inventory", "Ground"
Private Sub AddSectionItems(ByRef arrItems() As String, ByVal sectionName As String, _
                            ByRef LV As ListView, ByVal isEquippedSection As Boolean)
On Error GoTo fail
    Dim i As Long, hits() As ItemMatch, hitCount As Long
    Dim baseName As String, qty As Long

    If (Not Not arrItems) = 0 Then Exit Sub

    For i = LBound(arrItems) To UBound(arrItems)
        If isEquippedSection Then
            baseName = NormalizeEquippedName(arrItems(i))
            qty = 1
        Else
            ParseNameAndQty arrItems(i), baseName, qty
            If qty < 1 Then qty = 1
        End If
        If Len(baseName) = 0 Then GoTo nxt

        Erase hits
        hitCount = GetItemsByExactNameArr(baseName, hits)
        If hitCount <= 0 Then GoTo nxt

        ' One row per item (use first exact DB match)
        AddListViewRowForItem hits(0), sectionName, LV, qty
nxt:
    Next i
    Exit Sub
fail:
    MsgBox "AddSectionItems error: " & Err.Description, vbExclamation
End Sub

' === DB lookup ===
Private Function GetItemsByExactNameArr(ByVal exactName As String, ByRef hits() As ItemMatch) As Long
On Error GoTo fail
    Dim rs As DAO.Recordset
    Dim cnt As Long
    Dim itm As ItemMatch
    Dim nm As String, inGame As Long

    If tabItems Is Nothing Then Exit Function

    Set rs = tabItems.Clone
    If (rs.BOF And rs.EOF) Then rs.Close: Exit Function

    rs.MoveFirst
    Do While Not rs.EOF
        nm = NzStr(rs.Fields("Name").Value)
        inGame = NzLong(rs.Fields("In Game").Value)
        If inGame <> 0 Then
            If StrComp(nm, exactName, vbBinaryCompare) = 0 Then
                itm.Number = NzLong(rs.Fields("Number").Value)
                itm.name = nm
                itm.ItemType = NzLong(rs.Fields("ItemType").Value)
                itm.Worn = NzLong(rs.Fields("Worn").Value)
                itm.WeaponType = NzLong(rs.Fields("WeaponType").Value)
                itm.ObtainedFrom = NzStr(rs.Fields("Obtained From").Value)

                If cnt = 0 Then
                    ReDim hits(0)
                Else
                    ReDim Preserve hits(cnt)
                End If
                hits(cnt) = itm
                cnt = cnt + 1
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    GetItemsByExactNameArr = cnt
    Exit Function
fail:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    GetItemsByExactNameArr = 0
End Function

' === Row creation (choose best shop by lowest nCopperBuy via GetItemValue) ===
Private Sub AddListViewRowForItem(ByRef hit As ItemMatch, ByRef sectionName As String, ByRef LV As ListView, ByVal qty As Long)
On Error GoTo fail
    Dim shops() As ShopToken, shopCnt As Long
    Dim s As Long, nCharm As Integer
    Dim itemVal As tItemValue, chosen As tItemValue
    Dim bestIdx As Long, bestBuy As Double, anyBuy As Boolean
    Dim shopName As String, shopCell As String, valueCell As String
    Dim extraCount As Long

    nCharm = val(frmMain.txtCharStats(5).Text)

    shopCnt = ExtractShopsFromObtainedFrom(hit.ObtainedFrom, shops)
    If shopCnt = 0 Then
        shopCell = "none"
        valueCell = ""
        AddOneRow LV, hit, sectionName, qty, shopCell, valueCell, 0#
        Exit Sub
    End If

    bestIdx = -1: bestBuy = 0#: anyBuy = False
    For s = 0 To shopCnt - 1
        itemVal = GetItemValue(hit.Number, nCharm, 0, shops(s).ShopNumber, shops(s).bNObuy)
        If Not shops(s).bNObuy And itemVal.nBaseCost > 0 Then
            If Not anyBuy Or itemVal.nCopperBuy < bestBuy Then
                anyBuy = True
                bestBuy = itemVal.nCopperBuy
                bestIdx = s
                chosen = itemVal
            End If
        End If
    Next s

    If bestIdx = -1 Then
        bestIdx = 0
        chosen = GetItemValue(hit.Number, nCharm, 0, shops(0).ShopNumber, shops(0).bNObuy)
    End If

    shopName = GetShopRoomNames(shops(bestIdx).ShopNumber, , bHideRecordNumbers)
    extraCount = shopCnt - 1
    If extraCount > 0 Then
        shopCell = shopName & " +" & CStr(extraCount) & " more"
    Else
        shopCell = shopName
    End If

    If chosen.nBaseCost > 0 Then
        If shops(bestIdx).bNObuy Then
            valueCell = "(sell) " & chosen.sFriendlySellShort
        Else
            valueCell = chosen.sFriendlyBuyShort & " / " & chosen.sFriendlySellShort
        End If
    Else
        valueCell = ""
    End If

    AddOneRow LV, hit, sectionName, qty, shopCell, valueCell, chosen.nCopperBuy
    Exit Sub
fail:
    MsgBox "AddListViewRowForItem error: " & Err.Description, vbExclamation
End Sub

Private Sub AddOneRow(ByRef LV As ListView, ByRef hit As ItemMatch, ByVal sectionName As String, _
                      ByVal qty As Long, ByVal shopCell As String, ByVal valueCell As String, _
                      ByVal copperBuy As Double)
    Dim oLI As ListItem
    Dim wornText As String, usableText As String

    ' Worn text by ItemType:
    Select Case hit.ItemType
        Case 0: wornText = GetWornType(hit.Worn)       ' Armour
        Case 1: wornText = GetWeaponType(hit.WeaponType) ' Weapon
        Case Else: wornText = "Nowhere"
    End Select

    usableText = IIf(frmMain.TestGlobalFilter(hit.Number), "Yes", "No")

    Set oLI = LV.ListItems.Add()
    oLI.Text = CStr(hit.Number)                                  ' Col 1: Number
    oLI.ListSubItems.Add 1, "Name", hit.name                     ' Col 2
    oLI.ListSubItems.Add 2, "Source", sectionName                ' Col 3
    oLI.ListSubItems.Add 3, "QTY", CStr(qty)                     ' Col 4
    oLI.ListSubItems.Add 4, "Type", GetItemType(hit.ItemType)    ' Col 5
    oLI.ListSubItems.Add 5, "Worn", wornText                     ' Col 6
    oLI.ListSubItems.Add 6, "Usable", usableText                 ' Col 7
    oLI.ListSubItems.Add 7, "Shop", shopCell                     ' Col 8
    oLI.ListSubItems.Add 8, "Value", valueCell                   ' Col 9

    ' For sorting: keep BUY price (0 if NA)
    If copperBuy > 0# Then
        oLI.ListSubItems(8).Tag = CStr(copperBuy)
    Else
        oLI.ListSubItems(8).Tag = "0"
    End If
End Sub


' ======================================================================
' Tokenization & small helpers
' ======================================================================
' "Name (N)" -> base name + qty; if none, qty=1
Private Sub ParseNameAndQty(ByVal raw As String, ByRef baseName As String, ByRef qty As Long)
    Dim s As String, pOpen As Long, pClose As Long, inside As String
    s = Trim$(raw)
    baseName = s
    qty = 1

    pClose = InStrRev(s, ")")
    If pClose > 0 Then
        pOpen = InStrRev(s, "(", pClose)
        If pOpen > 0 And pOpen < pClose And pClose = Len(s) Then
            inside = Mid$(s, pOpen + 1, pClose - pOpen - 1)
            If IsNumeric(Trim$(inside)) Then
                qty = CLng(val(inside))
                baseName = Trim$(Left$(s, pOpen - 1))
            End If
        End If
    End If
End Sub

' Equipped lines: "Item Name (Slot)" -> just the name
Private Function NormalizeEquippedName(ByVal raw As String) As String
    Dim s As String, pOpen As Long, pClose As Long
    s = Trim$(raw)
    pClose = InStrRev(s, ")")
    If pClose > 0 Then
        pOpen = InStrRev(s, "(", pClose)
        If pOpen > 0 And pClose = Len(s) Then
            NormalizeEquippedName = Trim$(Left$(s, pOpen - 1))
            Exit Function
        End If
    End If
    NormalizeEquippedName = s
End Function

' Splits by comma into String() without Variants
Private Sub SplitToArray(ByVal s As String, ByRef aOut() As String)
On Error GoTo fail
    Dim startPos As Long, commaPos As Long, tok As String
    Dim n As Long
    Erase aOut: n = -1: startPos = 1
    Do
        commaPos = InStr(startPos, s, ",")
        If commaPos = 0 Then
            tok = Trim$(Mid$(s, startPos))
            If LenB(tok) > 0 Then
                n = n + 1
                If n = 0 Then ReDim aOut(0) Else ReDim Preserve aOut(n)
                aOut(n) = tok
            End If
            Exit Do
        Else
            tok = Trim$(Mid$(s, startPos, commaPos - startPos))
            If LenB(tok) > 0 Then
                n = n + 1
                If n = 0 Then ReDim aOut(0) Else ReDim Preserve aOut(n)
                aOut(n) = tok
            End If
            startPos = commaPos + 1
        End If
    Loop
    Exit Sub
fail:
    HandleError "SplitToArray"
End Sub

Private Function SqueezeSpaces(ByVal s As String) As String
    Do While InStr(1, s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop
    SqueezeSpaces = s
End Function

Private Function SafeArrayHasData(ByRef a() As String) As Boolean
    On Error GoTo nope
    If (UBound(a) - LBound(a)) >= 0 Then SafeArrayHasData = True
    Exit Function
nope:
End Function

Private Sub AddString(ByRef a() As String, ByVal s As String)
    Dim n As Long
    On Error Resume Next
    n = UBound(a) + 1
    If Err.Number <> 0 Then
        Err.clear
        ReDim a(0): a(0) = s
        Exit Sub
    End If
    On Error GoTo 0
    ReDim Preserve a(n)
    a(n) = s
End Sub

' Remove bracketed client inserts inside a span
Private Function StripBracketedChunks(ByVal s As String) As String
    Dim res As String, ch As String * 1, i As Long, depth As Long
    res = vbNullString: depth = 0
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = "[" Then
            depth = depth + 1
        ElseIf ch = "]" Then
            If depth > 0 Then depth = depth - 1
        Else
            If depth = 0 Then res = res & ch
        End If
    Next i
    StripBracketedChunks = res
End Function

' Filters
Private Function IsCashItem(ByVal sItem As String) As Boolean
    Dim sL As String
    sL = " " & LCase$(Trim$(sItem)) & " "
    If InStr(sL, " gold crown ") > 0 Or InStr(sL, " gold crowns ") > 0 Then IsCashItem = True: Exit Function
    If InStr(sL, " silver noble ") > 0 Or InStr(sL, " silver nobles ") > 0 Then IsCashItem = True: Exit Function
    If InStr(sL, " copper farthing ") > 0 Or InStr(sL, " copper farthings ") > 0 Then IsCashItem = True: Exit Function
    If InStr(sL, " platinum piece ") > 0 Or InStr(sL, " platinum pieces ") > 0 Then IsCashItem = True: Exit Function
    If InStr(sL, " runic coin ") > 0 Or InStr(sL, " runic coins ") > 0 Then IsCashItem = True: Exit Function
End Function

Private Function IsEquippedItem(ByVal sItem As String) As Boolean
    Dim p As Long
    If Right$(sItem, 1) <> ")" Then Exit Function
    p = InStrRev(sItem, " (")
    If p > 0 And p < Len(sItem) Then IsEquippedItem = True
End Function

Private Function IsBracketLine(ByVal sLine As String) As Boolean
    Dim t As String
    t = Trim$(sLine)
    If Len(t) >= 2 Then
        IsBracketLine = (Left$(t, 1) = "[" And Right$(t, 1) = "]")
    Else
        IsBracketLine = False
    End If
End Function

' Shop parsing (from "Obtained From")
Private Function ExtractShopsFromObtainedFrom(ByVal obtained As String, ByRef shops() As ShopToken) As Long
    Dim parts() As String, i As Long, t As String, n As Long
    Dim tok As ShopToken, cnt As Long

    If Len(Trim$(obtained)) = 0 Then ExtractShopsFromObtainedFrom = 0: Exit Function

    parts = Split(obtained, ",")
    For i = LBound(parts) To UBound(parts)
        t = LCase$(Trim$(parts(i)))
        If Left$(t, 4) = "shop" Then
            tok.bNObuy = False: tok.bNoSell = False
            If InStr(t, "(sell)") > 0 Then tok.bNObuy = True
            n = ExtractFirstNumber(t)
            If n > 0 Then
                tok.ShopNumber = n
                If cnt = 0 Then
                    ReDim shops(0)
                Else
                    ReDim Preserve shops(cnt)
                End If
                shops(cnt) = tok
                cnt = cnt + 1
            End If
        End If
    Next i

    ExtractShopsFromObtainedFrom = cnt
End Function

Private Function ExtractFirstNumber(ByVal s As String) As Long
    Dim i As Long, digits As String, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        ElseIf Len(digits) > 0 Then
            Exit For
        End If
    Next i
    If Len(digits) > 0 Then ExtractFirstNumber = CLng(digits)
End Function

' Null helpers
Private Function NzStr(ByVal v As Variant) As String
    If IsNull(v) Then NzStr = "" Else NzStr = CStr(v)
End Function

Private Function NzLong(ByVal v As Variant) As Long
    If IsNull(v) Or v = "" Then NzLong = 0 Else NzLong = CLng(v)
End Function

' Minimal error handler (keeps module self-contained)
Private Sub HandleError(ByVal where As String)
    Debug.Print where & " error: "; Err.Number & " - " & Err.Description
    Err.clear
End Sub


