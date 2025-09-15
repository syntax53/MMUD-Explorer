Attribute VB_Name = "modItemParse"
Option Explicit

' ----------------------------------------------------------------
' Represent a DB hit without keeping a live cursor into tabItems
' ----------------------------------------------------------------
Public Type ItemMatch
    Number As Long
    name As String
    ItemType As Long
    Worn As Long
    WeaponType As Long
    ObtainedFrom As String
End Type


Private Type ShopToken
    ShopNumber As Long
    bNoSell As Boolean
    bNObuy As Boolean
End Type


'================================================================================
' Data shape for results
'================================================================================
Public Type ItemParseResult
    sEquipped() As String          ' items with slots, e.g., "white gold ring (Finger)" (deduped)
    EquippedCount As Integer
    sInventory() As String         ' non-key, non-cash, NOT equipped (consolidated -> "name (N)")
    InventoryCount As Integer
    sKeys() As String              ' keys from "You have the following keys:" (consolidated -> "name (N)")
    KeysCount As Integer
    sGround() As String            ' items seen via "You notice ..." (includes keys; cash removed; consolidated)
    GroundCount As Integer
End Type

Private Const HDR_INV   As String = "You are carrying"
Private Const HDR_KEYS  As String = "You have the following keys"
Private Const HDR_NOKEY As String = "You have no keys"
Private Const HDR_NOTC  As String = "You notice"

'================================================================================
' Public entry point
'================================================================================
Public Function ParseGameTextInventory(ByVal sInput As String) As ItemParseResult
On Error GoTo error:
    Dim tOut As ItemParseResult
    Dim sNorm As String
    Dim vLines() As String
    Dim i As Long
    Dim sLine As String

    sNorm = sInput
    sNorm = Replace$(sNorm, vbCr, vbNullString) ' normalize to LF only (keep LFs until tokenizing)
    vLines = Split(sNorm, vbLf)

    ' init zero-length arrays
    ReDim tOut.sEquipped(0)
    ReDim tOut.sInventory(0)
    ReDim tOut.sKeys(0)
    ReDim tOut.sGround(0)

    ' ---- Pass 1: collect inventory blocks + keys blocks + ground blocks ----
    i = 0
    Do While i <= UBound(vLines)
        sLine = Trim$(vLines(i))

        If StartsWithCI(sLine, HDR_INV) Then
            Dim sInvBlob As String
            sInvBlob = CollectInventoryBlob(vLines, i)   ' advances i appropriately
            ParseInventoryBlob sInvBlob, tOut            ' fills equipped + inventory (raw)
        ElseIf StartsWithCI(sLine, HDR_KEYS) Then
            Dim sKeysBlob As String
            sKeysBlob = CollectKeysBlob(vLines, i)
            ParseKeysBlob sKeysBlob, tOut
        ElseIf StartsWithCI(sLine, HDR_NOKEY) Then
            ' explicit "no keys" section: just consume line(s) and stop at boundary
            Dim sNoKey As String
            sNoKey = CollectNoKeysBlob(vLines, i)
            ' nothing to add
        ElseIf StartsWithCI(sLine, HDR_NOTC) Then
            Dim sNoticeBlob As String
            sNoticeBlob = CollectNoticeBlob(vLines, i)
            ParseNoticeBlob sNoticeBlob, tOut
        End If

        i = i + 1
    Loop

    ' ---- Pass 1b (rescue): regex-scan the whole text for any missed ground blobs ----
    RescueGroundWithRegex sNorm, tOut

    ' ---- Pass 2: consolidate counts & clean punctuation ----
    ConsolidateList tOut.sInventory, False        ' -> "name (N)"
    ConsolidateList tOut.sGround, False           ' -> "name (N)"
    ConsolidateList tOut.sKeys, True              ' -> "name (N)" with key-name singularization

    ' ---- Pass 3: dedupe equipped (no counts)
    DedupeEquipped tOut.sEquipped
    
    If Len(Trim(tOut.sEquipped(0))) > 0 Then tOut.EquippedCount = UBound(tOut.sEquipped) + 1
    If Len(Trim(tOut.sInventory(0))) > 0 Then tOut.InventoryCount = UBound(tOut.sInventory) + 1
    If Len(Trim(tOut.sKeys(0))) > 0 Then tOut.KeysCount = UBound(tOut.sKeys) + 1
    If Len(Trim(tOut.sGround(0))) > 0 Then tOut.GroundCount = UBound(tOut.sGround) + 1
    
out:
    ParseGameTextInventory = tOut
    Exit Function
error:
    Call HandleError("ParseGameTextInventory")
    Resume out:
End Function

'================================================================================
' Inventory blob collection/parsing
'================================================================================

Private Function CollectInventoryBlob(ByRef vLines() As String, ByRef i As Long) As String
On Error GoTo error:
    Dim s As String, j As Long, sLine As String
    Dim pos As Long, sLeft As String, sRight As String
    s = vbNullString
    j = i

    Do While j <= UBound(vLines)
        sLine = Trim$(vLines(j))

        If ShouldStopSection(sLine, (j <> i)) Then Exit Do

        ' stop if another section header appears mid-line; split the line
        pos = FindInlineBoundaryPos(sLine)
        If pos > 1 Then
            sLeft = Trim$(Left$(sLine, pos - 1))
            sRight = Mid$(sLine, pos)
            If LenB(sLeft) > 0 Then
                If LenB(s) > 0 Then s = s & " "
                s = s & sLeft
            End If
            vLines(j) = sRight ' show next section to main loop
            Exit Do
        End If

        If LenB(s) > 0 Then s = s & " "
        s = s & sLine
        j = j + 1
    Loop

    i = j - 1
    CollectInventoryBlob = s
    Exit Function
error:
    Call HandleError("CollectInventoryBlob")
    CollectInventoryBlob = s
End Function

'==============================================================================
' Keys collection/parsing (including "You have no keys.")
'==============================================================================

Private Function CollectKeysBlob(ByRef vLines() As String, ByRef i As Long) As String
On Error GoTo error:
    Dim s As String, j As Long, sLine As String
    Dim pos As Long, sLeft As String, sRight As String
    s = vbNullString
    j = i

    Do While j <= UBound(vLines)
        sLine = Trim$(vLines(j))

        If ShouldStopSection(sLine, (j <> i)) Then Exit Do

        pos = FindInlineBoundaryPos(sLine)
        If pos > 1 Then
            sLeft = Trim$(Left$(sLine, pos - 1))
            sRight = Mid$(sLine, pos)
            If LenB(sLeft) > 0 Then
                If LenB(s) > 0 Then s = s & " "
                s = s & sLeft
            End If
            vLines(j) = sRight
            Exit Do
        End If

        If LenB(s) > 0 Then s = s & " "
        s = s & sLine

        If Right$(sLine, 1) = "." Then
            j = j + 1 ' consume the period line
            Exit Do
        End If

        j = j + 1
    Loop

    i = j - 1
    CollectKeysBlob = s
    Exit Function
error:
    Call HandleError("CollectKeysBlob")
    CollectKeysBlob = s
End Function

' Explicit "You have no keys." section consumer (so it doesn't bleed into inventory)
Private Function CollectNoKeysBlob(ByRef vLines() As String, ByRef i As Long) As String
On Error GoTo error:
    Dim s As String, j As Long, sLine As String
    Dim pos As Long
    s = vbNullString
    j = i

    Do While j <= UBound(vLines)
        sLine = Trim$(vLines(j))

        If ShouldStopSection(sLine, (j <> i)) Then Exit Do

        pos = FindInlineBoundaryPos(sLine)
        If pos > 1 Then
            ' consume up to the boundary and return
            If LenB(s) > 0 Then s = s & " "
            s = s & Trim$(Left$(sLine, pos - 1))
            vLines(j) = Mid$(sLine, pos)
            Exit Do
        End If

        If LenB(s) > 0 Then s = s & " "
        s = s & sLine

        ' usually ends on the same line
        If Right$(sLine, 1) = "." Then
            j = j + 1
            Exit Do
        End If

        j = j + 1
    Loop

    i = j - 1
    CollectNoKeysBlob = s
    Exit Function
error:
    Call HandleError("CollectNoKeysBlob")
    CollectNoKeysBlob = s
End Function

'==============================================================================
' Ground "You notice ..." collection/parsing (line-wise)
'==============================================================================

Private Function CollectNoticeBlob(ByRef vLines() As String, ByRef i As Long) As String
On Error GoTo error:
    Dim s As String, j As Long, sLine As String
    Dim pos As Long, sLeft As String, sRight As String
    s = vbNullString
    j = i

    Do While j <= UBound(vLines)
        sLine = Trim$(vLines(j))

        ' Skip bracketed helper lines between wraps (e.g., “[...added to known items]”)
        If IsBracketLine(sLine) Then
            j = j + 1
            GoTo continue_loop
        End If

        If ShouldStopSection(sLine, (j <> i)) Then Exit Do

        ' handle mid-line boundary (rare but seen)
        pos = FindInlineBoundaryPos(sLine)
        If pos > 1 Then
            sLeft = Trim$(Left$(sLine, pos - 1))
            sRight = Mid$(sLine, pos)
            If LenB(sLeft) > 0 Then
                If LenB(s) > 0 Then s = s & " "
                s = s & sLeft
            End If
            vLines(j) = sRight
            Exit Do
        End If

        If LenB(s) > 0 Then s = s & " "
        s = s & sLine

        ' Normal terminator: “here.”
        If LCase$(Right$(sLine, 5)) = "here." Then
            j = j + 1 ' consume the “here.” line
            Exit Do
        End If

        j = j + 1
continue_loop:
    Loop

    i = j - 1
    CollectNoticeBlob = s
    Exit Function
error:
    Call HandleError("CollectNoticeBlob")
    CollectNoticeBlob = s
End Function

'================================================================================
' Unified section boundary detection
'================================================================================

Private Function IsSectionBoundary(ByVal sLine As String) As Boolean
    Dim sL As String
    sL = LCase$(Trim$(sLine))

    If LenB(sL) = 0 Then IsSectionBoundary = True: Exit Function

    ' canonical headers / blocks
    If StartsWithCI(sL, LCase$(HDR_INV)) Then IsSectionBoundary = True: Exit Function
    If StartsWithCI(sL, LCase$(HDR_KEYS)) Then IsSectionBoundary = True: Exit Function
    If StartsWithCI(sL, LCase$(HDR_NOKEY)) Then IsSectionBoundary = True: Exit Function
    If StartsWithCI(sL, LCase$(HDR_NOTC)) Then IsSectionBoundary = True: Exit Function

    ' common tail sections after inventory/keys
    If StartsWithCI(sL, "wealth:") Then IsSectionBoundary = True: Exit Function
    If StartsWithCI(sL, "encumbrance:") Then IsSectionBoundary = True: Exit Function

    ' stat card headers (start of a new character dump)
    If StartsWithCI(sL, "name:") Then IsSectionBoundary = True: Exit Function
    If StartsWithCI(sL, "race:") Then IsSectionBoundary = True: Exit Function
    If StartsWithCI(sL, "class:") Then IsSectionBoundary = True: Exit Function

    ' prompts & room text markers
    If Left$(sL, 1) = "[" Then IsSectionBoundary = True: Exit Function       ' e.g., [HP=...]:
    If InStr(1, sL, "obvious exits:", vbTextCompare) > 0 Then IsSectionBoundary = True: Exit Function
    If StartsWithCI(sL, "also here:") Then IsSectionBoundary = True: Exit Function

    ' common combat flags that appear as standalone lines
    If InStr(1, sL, "*combat", vbTextCompare) > 0 Then IsSectionBoundary = True: Exit Function
End Function

Private Function ShouldStopSection(ByVal sLine As String, ByVal bPastFirst As Boolean) As Boolean
    ShouldStopSection = (bPastFirst And IsSectionBoundary(sLine))
End Function

'================================================================================
' Parse blobs
'================================================================================

Private Sub ParseInventoryBlob(ByVal sBlob As String, ByRef tOut As ItemParseResult)
On Error GoTo error:
    Dim aTok() As String
    Dim k As Long
    Dim sItem As String
    Dim sAfterPrefix As String

    sAfterPrefix = sBlob
    If StartsWithCI(sAfterPrefix, HDR_INV) Then
        sAfterPrefix = Trim$(Mid$(sAfterPrefix, Len(HDR_INV) + 1))
    End If

    sAfterPrefix = RemoveTrailingAfterInventory(sAfterPrefix)
    Call SplitToArray(NormalizeWhitespace(sAfterPrefix), aTok)

    If SafeArrayHasData(aTok) Then
        For k = LBound(aTok) To UBound(aTok)
            sItem = Trim$(aTok(k))
            If LenB(sItem) = 0 Then GoTo continue_k
            If IsBracketLine(sItem) Then GoTo continue_k
            If IsCashItem(sItem) Then GoTo continue_k

            If IsEquippedItem(sItem) Then
                AddString tOut.sEquipped, sItem
            Else
                AddString tOut.sInventory, sItem   ' raw; consolidated later
            End If
continue_k:
        Next k
    End If

out:
    Exit Sub
error:
    Call HandleError("ParseInventoryBlob")
    Resume out:
End Sub

Private Function RemoveTrailingAfterInventory(ByVal s As String) As String
    RemoveTrailingAfterInventory = s
End Function

Private Sub ParseKeysBlob(ByVal sBlob As String, ByRef tOut As ItemParseResult)
On Error GoTo error:
    Dim s As String
    Dim aTok() As String
    Dim k As Long
    Dim sItem As String

    If StartsWithCI(sBlob, HDR_KEYS) Then
        s = Mid$(sBlob, Len(HDR_KEYS) + 1)
        s = Trim$(s)
        If Left$(s, 1) = ":" Then s = Trim$(Mid$(s, 2))
    Else
        s = sBlob
    End If

    If InStr(1, LCase$(s), "no keys", vbTextCompare) > 0 Then Exit Sub

    s = RTrim$(s)
    If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)

    Call SplitToArray(NormalizeWhitespace(s), aTok)
    If SafeArrayHasData(aTok) Then
        For k = LBound(aTok) To UBound(aTok)
            sItem = Trim$(aTok(k))
            If LenB(sItem) > 0 Then AddString tOut.sKeys, sItem   ' raw; consolidated later
        Next k
    End If

out:
    Exit Sub
error:
    Call HandleError("ParseKeysBlob")
    Resume out:
End Sub

Private Sub ParseNoticeBlob(ByVal sBlob As String, ByRef tOut As ItemParseResult)
On Error GoTo error:
    Dim s As String
    Dim aTok() As String
    Dim k As Long
    Dim sItem As String

    s = sBlob
    If StartsWithCI(s, HDR_NOTC) Then s = Trim$(Mid$(s, Len(HDR_NOTC) + 1))
    ' remove trailing "here." if present
    If LCase$(Right$(s, 5)) = "here." Then s = Left$(s, Len(s) - 5)

    Call SplitToArray(NormalizeWhitespace(s), aTok)
    If SafeArrayHasData(aTok) Then
        For k = LBound(aTok) To UBound(aTok)
            sItem = Trim$(aTok(k))
            If LenB(sItem) = 0 Then GoTo continue_k
            If IsBracketLine(sItem) Then GoTo continue_k
            If IsCashItem(sItem) Then GoTo continue_k
            AddString tOut.sGround, sItem   ' raw; consolidated later
continue_k:
        Next k
    End If

out:
    Exit Sub
error:
    Call HandleError("ParseNoticeBlob")
    Resume out:
End Sub

'================================================================================
' Regex rescue for ground blobs
'================================================================================

' Finds all "You notice ... here." spans across newlines/brackets and parses them.
Private Sub RescueGroundWithRegex(ByVal sAll As String, ByRef tOut As ItemParseResult)
On Error GoTo error:
    Dim tMatches() As RegexMatches
    Dim i As Long
    Dim sSpan As String

    ' Pattern: You notice + anything (incl. newlines) + here.
    ' VBScript.RegExp has no singleline-dot, so use [\s\S]*?
    tMatches = RegExpFindv2(sAll, HDR_NOTC & "\s+([\s\S]*?)\s+here\.", False, True, False)
    If UBound(tMatches) = 0 And LenB(tMatches(0).sFullMatch) = 0 Then GoTo out

    For i = LBound(tMatches) To UBound(tMatches)
        sSpan = tMatches(i).sSubMatches(0)
        ' remove bracketed annotations, normalize whitespace, then split
        sSpan = StripBracketedChunks(sSpan)
        sSpan = NormalizeWhitespace(sSpan)
        If LenB(sSpan) > 0 Then
            Dim aTok() As String, k As Long, sItem As String
            Call SplitToArray(sSpan, aTok)
            If SafeArrayHasData(aTok) Then
                For k = LBound(aTok) To UBound(aTok)
                    sItem = Trim$(aTok(k))
                    If LenB(sItem) = 0 Then GoTo next_k
                    If IsCashItem(sItem) Then GoTo next_k
                    AddString tOut.sGround, sItem
next_k:
                Next k
            End If
        End If
    Next i

out:
    Exit Sub
error:
    Call HandleError("RescueGroundWithRegex")
    Resume out:
End Sub

' Remove any [...] chunks inside a text span (works across inline inserts)
Private Function StripBracketedChunks(ByVal s As String) As String
    Dim res As String, ch As String * 1
    Dim i As Long, depth As Long
    res = vbNullString
    depth = 0
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

'================================================================================
' Consolidation (counts, punctuation, normalization)  -- no Variants, no For Each
'================================================================================

' Consolidate items: merges duplicates, applies counts, optional keys-name normalization.
Private Sub ConsolidateList(ByRef a() As String, ByVal bIsKeysList As Boolean)
On Error GoTo error:
    Dim keys() As String, counts() As Long
    Dim nKeys As Long, i As Long
    Dim nm As String, cnt As Long, idx As Long

    If Not SafeArrayHasData(a) Then Exit Sub

    nKeys = 0
    Erase keys
    Erase counts

    For i = LBound(a) To UBound(a)
        If LenB(a(i)) > 0 Then
            ParseCountAndName a(i), nm, cnt, bIsKeysList   ' -> nm (lowercased), cnt>=1
            If bIsKeysList Then nm = NormalizeKeyName(nm, cnt)
            idx = FindKeyIndex(keys, nKeys, nm)
            If idx < 0 Then
                ReDim Preserve keys(nKeys)
                ReDim Preserve counts(nKeys)
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
        GoTo out
    End If

    ReDim a(nKeys - 1)
    For i = 0 To nKeys - 1
        If counts(i) > 1 Then
            a(i) = keys(i) & " (" & CStr(counts(i)) & ")"
        Else
            a(i) = keys(i)
        End If
    Next i

out:
    Exit Sub
error:
    Call HandleError("ConsolidateList")
    Resume out:
End Sub

' Normalize key-name for merging/printing (keys list only).
' - "... keys" -> "... key"
' - if count>1 and last word ends with single 's', drop it (e.g., "golden idols" -> "golden idol")
Private Function NormalizeKeyName(ByVal nameLower As String, ByVal countVal As Long) As String
    Dim s As String, p As Long, lastWord As String
    s = Trim$(nameLower)

    If Right$(s, 5) = " keys" Then
        NormalizeKeyName = Left$(s, Len(s) - 1) ' -> " key"
        Exit Function
    End If

    If countVal > 1 Then
        p = InStrRev(s, " ")
        If p > 0 Then
            lastWord = Mid$(s, p + 1)
            If Right$(lastWord, 1) = "s" And Right$(lastWord, 2) <> "ss" Then
                s = Left$(s, p) & Left$(lastWord, Len(lastWord) - 1)
            End If
        Else
            If Right$(s, 1) = "s" And Right$(s, 2) <> "ss" Then
                s = Left$(s, Len(s) - 1)
            End If
        End If
    End If

    NormalizeKeyName = s
End Function

' Linear search for existing key (nm already lowercased)
Private Function FindKeyIndex(ByRef keys() As String, ByVal nKeys As Long, ByVal nm As String) As Long
    Dim i As Long
    For i = 0 To nKeys - 1
        If keys(i) = nm Then FindKeyIndex = i: Exit Function
    Next i
    FindKeyIndex = -1
End Function

Private Sub ParseCountAndName(ByVal token As String, ByRef nameOut As String, ByRef countOut As Long, ByVal bIsKeysList As Boolean)
    Dim s As String, p As Long, firstWord As String
    s = Trim$(token)

    ' strip trailing punctuation (periods/semicolons)
    Do While Len(s) > 0 And (Right$(s, 1) = "." Or Right$(s, 1) = ";")
        s = Left$(s, Len(s) - 1)
        s = RTrim$(s)
    Loop

    ' leading count?
    p = InStr(1, s, " ")
    If p > 0 Then
        firstWord = Left$(s, p - 1)
        If IsNumeric(firstWord) Then
            countOut = CLng(firstWord)
            nameOut = Trim$(Mid$(s, p + 1))
        Else
            countOut = 1
            nameOut = s
        End If
    Else
        countOut = 1
        nameOut = s
    End If

    ' normalize inner spaces and lowercase (collapse any LF/TABs beforehand)
    nameOut = NormalizeWhitespace(nameOut)
    nameOut = LCase$(Trim$(nameOut))   ' canonicalize for merging
End Sub

'================================================================================
' Equipped dedupe (case-insensitive, by full "name (Slot)" text)
'================================================================================
Private Sub DedupeEquipped(ByRef a() As String)
On Error GoTo error:
    Dim out() As String, nOut As Long
    Dim seen() As String, nSeen As Long
    Dim i As Long, nm As String

    If Not SafeArrayHasData(a) Then Exit Sub

    nOut = 0
    nSeen = 0

    For i = LBound(a) To UBound(a)
        nm = LCase$(NormalizeWhitespace(a(i)))
        If LenB(nm) = 0 Then GoTo continue_i
        If IndexOf(seen, nSeen, nm) < 0 Then
            ReDim Preserve seen(nSeen): seen(nSeen) = nm: nSeen = nSeen + 1
            ReDim Preserve out(nOut):  out(nOut) = Trim$(NormalizeWhitespace(a(i))): nOut = nOut + 1
        End If
continue_i:
    Next i

    If nOut = 0 Then
        ReDim a(0): a(0) = vbNullString
    Else
        a = out
    End If
    Exit Sub
error:
    Call HandleError("DedupeEquipped")
End Sub

Private Function IndexOf(ByRef arr() As String, ByVal n As Long, ByVal keyL As String) As Long
    Dim i As Long
    For i = 0 To n - 1
        If arr(i) = keyL Then IndexOf = i: Exit Function
    Next i
    IndexOf = -1
End Function

'================================================================================
' Splitting & small helpers
'================================================================================

' Splits on commas into a String() (trims each token). No Variant usage.
Private Sub SplitToArray(ByVal s As String, ByRef aOut() As String)
On Error GoTo error:
    Dim startPos As Long, commaPos As Long, tok As String
    Dim n As Long

    Erase aOut
    n = -1
    startPos = 1

    Do
        commaPos = InStr(startPos, s, ",")
        If commaPos = 0 Then
            tok = Trim$(Mid$(s, startPos))
            If LenB(tok) > 0 Then
                n = n + 1
                ReDim Preserve aOut(n)
                aOut(n) = tok
            End If
            Exit Do
        Else
            tok = Trim$(Mid$(s, startPos, commaPos - startPos))
            If LenB(tok) > 0 Then
                n = n + 1
                ReDim Preserve aOut(n)
                aOut(n) = tok
            End If
            startPos = commaPos + 1
        End If
    Loop

out:
    Exit Sub
error:
    Call HandleError("SplitToArray")
    Resume out:
End Sub

' Collapse ALL whitespace (LF/CR/TAB/multiple spaces) to single spaces.
Private Function NormalizeWhitespace(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace$(t, vbCr, " ")
    t = Replace$(t, vbLf, " ")
    t = Replace$(t, vbTab, " ")
    Do While InStr(1, t, "  ") > 0
        t = Replace$(t, "  ", " ")
    Loop
    NormalizeWhitespace = Trim$(t)
End Function

Private Function SqueezeSpaces(ByVal s As String) As String
    ' Kept for compatibility; now just calls NormalizeWhitespace.
    SqueezeSpaces = NormalizeWhitespace(s)
End Function

Private Function SafeArrayHasData(ByRef a() As String) As Boolean
    On Error GoTo nope
    If (UBound(a) - LBound(a)) >= 0 Then SafeArrayHasData = True
    Exit Function
nope:
End Function

' Returns earliest position (>1) of any known boundary phrase inside sLine.
Private Function FindInlineBoundaryPos(ByVal sLine As String) As Long
    Dim sL As String
    Dim p As Long, best As Long

    sL = LCase$(sLine)
    best = 0

    p = InStr(1, sL, LCase$(HDR_KEYS))
    If p > 1 Then If best = 0 Or p < best Then best = p

    p = InStr(1, sL, LCase$(HDR_NOKEY))
    If p > 1 Then If best = 0 Or p < best Then best = p

    p = InStr(1, sL, LCase$(HDR_INV))
    If p > 1 Then If best = 0 Or p < best Then best = p

    p = InStr(1, sL, LCase$(HDR_NOTC))
    If p > 1 Then If best = 0 Or p < best Then best = p

    p = InStr(1, sL, "wealth:")
    If p > 1 Then If best = 0 Or p < best Then best = p

    p = InStr(1, sL, "encumbrance:")
    If p > 1 Then If best = 0 Or p < best Then best = p

    FindInlineBoundaryPos = best
End Function

Private Function StartsWithCI(ByVal s As String, ByVal prefix As String) As Boolean
    StartsWithCI = (LCase$(Left$(s, Len(prefix))) = LCase$(prefix))
End Function

'================================================================================
' Filters
'================================================================================

Private Function IsCashItem(ByVal sItem As String) As Boolean
    Dim sL As String
    sL = " " & LCase$(Trim$(NormalizeWhitespace(sItem))) & " "
    If InStr(sL, " gold crown ") > 0 Or InStr(sL, " gold crowns ") > 0 Then IsCashItem = True: Exit Function
    If InStr(sL, " silver noble ") > 0 Or InStr(sL, " silver nobles ") > 0 Then IsCashItem = True: Exit Function
    If InStr(sL, " copper farthing ") > 0 Or InStr(sL, " copper farthings ") > 0 Then IsCashItem = True: Exit Function
    If InStr(sL, " platinum piece ") > 0 Or InStr(sL, " platinum pieces ") > 0 Then IsCashItem = True: Exit Function
    If InStr(sL, " runic coin ") > 0 Or InStr(sL, " runic coins ") > 0 Then IsCashItem = True: Exit Function
End Function

' Equipped item detection: ends with ")"; has " (" before it.
Private Function IsEquippedItem(ByVal sItem As String) As Boolean
    Dim p As Long
    If Right$(sItem, 1) <> ")" Then Exit Function
    p = InStrRev(sItem, " (")
    If p > 0 And p < Len(sItem) Then IsEquippedItem = True
End Function

' Lines like: [runed silk pants added to known items] or any bracketed helper
Private Function IsBracketLine(ByVal sLine As String) As Boolean
    Dim t As String
    t = Trim$(sLine)
    If Len(t) >= 2 Then
        IsBracketLine = (Left$(t, 1) = "[" And Right$(t, 1) = "]")
    Else
        IsBracketLine = False
    End If
End Function

'================================================================================
' Append helper
'================================================================================
Private Sub AddString(ByRef a() As String, ByVal s As String)
    Dim n As Long
    On Error Resume Next
    n = UBound(a) + 1
    If Err.Number <> 0 Then
        Err.clear
        ReDim a(0)
        a(0) = s
        Exit Sub
    End If
    On Error GoTo 0
    ReDim Preserve a(n)
    a(n) = s
End Sub


' ==============================================================
' Item Manager population from parsed text (tItems)
' Requires:
'   - Public tabItems As DAO.Recordset  (opened on Items table)
'   - External functions:
'       GetItemType(ByVal nType As Long) As String
'       GetWornType(ByVal nWorn As Long) As String
'       TestGlobalFilter(ByVal nItemNumber As Long) As String
'       GetShopName(ByVal nShopNum As Long, Optional ByVal bShort As Boolean = True) As String
'       GetItemValueQuickString(ByVal nItemNumber As Long, Optional ByVal nCharm As Integer, _
'           Optional ByVal nMarkup As Integer, Optional ByVal nShopNumber As Long, _
'           Optional ByVal bNoSell As Boolean, Optional ByVal bNoBuy As Boolean) As String
'       GetItemCopperCost(ByVal nItemNumber As Long) As Long
'
' Notes:
'   - Prompts to import Equipped / Keys.
'   - If items remain and ListView already has rows, prompt to clear.
'   - Exact [Name] matches only; skip items with [In Game] = 0 or no matches.
'   - If multiple exact-name matches exist, include all.
'   - Quantity: add N rows (Ground capped at 5 per unique item).
'   - Shop replication: one row per matching shop token; else Shop="none", Value="(no value)".
' ==============================================================

Public Sub PopulateItemManagerFromParsed(ByRef tItems As ItemParseResult, ByRef lvItemManager As ListView)
    On Error GoTo fail

    Dim haveEquipped As Boolean, haveKeys As Boolean, haveInv As Boolean, haveGround As Boolean
    Dim doEquipped As Boolean, doKeys As Boolean
    Dim totalToAdd As Long

    haveEquipped = (tItems.EquippedCount > 0)
    haveKeys = (tItems.KeysCount > 0)            ' keys list from inventory only
    haveInv = (tItems.InventoryCount > 0)
    haveGround = (tItems.GroundCount > 0)

    doEquipped = True
    doKeys = True

    If haveEquipped Then
        If MsgBox("Import EQUIPPED items into Item Manager?", vbQuestion Or vbYesNo, "Import Equipped?") = vbNo Then
            doEquipped = False
        End If
    End If

    If haveKeys Then
        If MsgBox("Import KEYS (from inventory) into Item Manager?", vbQuestion Or vbYesNo, "Import Keys?") = vbNo Then
            doKeys = False
        End If
    End If

    totalToAdd = 0
    If doEquipped Then totalToAdd = totalToAdd + tItems.EquippedCount
    If doKeys Then totalToAdd = totalToAdd + tItems.KeysCount
    If haveGround Then totalToAdd = totalToAdd + tItems.GroundCount
    If haveInv Then totalToAdd = totalToAdd + tItems.InventoryCount

    If totalToAdd > 0 And lvItemManager.ListItems.Count > 0 Then
        If MsgBox("Clear the existing Item Manager list first?", vbQuestion Or vbYesNo, "Clear Existing?") = vbYes Then
            lvItemManager.ListItems.clear
        End If
    End If

    If doEquipped And haveEquipped Then
        AddSectionItems tItems.sEquipped, "Equipped", lvItemManager, True
    End If
    If doKeys And haveKeys Then
        AddSectionItems tItems.sKeys, "Inventory", lvItemManager, False
    End If
    If haveGround Then
        AddSectionItems tItems.sGround, "Ground", lvItemManager, False
    End If
    If haveInv Then
        AddSectionItems tItems.sInventory, "Inventory", lvItemManager, False
    End If

    Exit Sub
fail:
    MsgBox "PopulateItemManagerFromParsed error: " & Err.Description, vbExclamation
End Sub

' Return index for name, creating new slot if needed.
Private Function GetOrCreateNameIndex(ByRef names() As String, ByRef counts() As Integer, _
                                      ByRef used As Long, ByVal name As String) As Long
    Dim i As Long
    For i = 0 To used - 1
        If StrComp(names(i), name, vbBinaryCompare) = 0 Then
            GetOrCreateNameIndex = i
            Exit Function
        End If
    Next

    If used = 0 Then
        ReDim names(0)
        ReDim counts(0)
    Else
        ReDim Preserve names(used)
        ReDim Preserve counts(used)
    End If
    names(used) = name
    counts(used) = 0
    GetOrCreateNameIndex = used
    used = used + 1
End Function

'====================[ Section Driver ]====================
' sectionName: "Equipped", "Inventory", "Key", "Ground"
Private Sub AddSectionItems(ByRef arrItems() As String, ByVal sectionName As String, _
                            ByRef LV As ListView, ByVal isEquippedSection As Boolean)
    On Error GoTo fail

    Dim i As Long, h As Long
    Dim baseName As String
    Dim qty As Long, addQty As Long
    Dim hits() As ItemMatch
    Dim hitCount As Long
    Dim toAdd As Long
    Dim idx As Long, remaining As Long

    ' Per-section cap trackers (name -> how many rows already added for that item)
    Dim capNames() As String, capCounts() As Integer, capUsed As Long

    If (Not Not arrItems) = 0 Then Exit Sub ' uninitialized

    For i = LBound(arrItems) To UBound(arrItems)
        If isEquippedSection Then
            baseName = NormalizeEquippedName(arrItems(i))   ' strip "(Slot)"
            qty = 1
        Else
            ParseNameAndQty arrItems(i), baseName, qty      ' strip "(N)" -> qty
            If Len(baseName) = 0 Then GoTo nextI
        End If

        ' Determine remaining capacity for this item in THIS section (max 5)
        idx = GetOrCreateNameIndex(capNames, capCounts, capUsed, baseName)
        remaining = 5 - capCounts(idx)
        If remaining <= 0 Then GoTo nextI

        ' Desired adds per DB match (will be further limited by remaining)
        addQty = qty

        Erase hits
        hitCount = GetItemsByExactNameArr(baseName, hits)   ' exact match(es); in-game only
        If hitCount = 0 Then GoTo nextI

        For h = 0 To hitCount - 1
            If remaining <= 0 Then Exit For

            toAdd = addQty
            If toAdd > remaining Then toAdd = remaining

            Do While toAdd > 0
                AddListViewRowsForItem hits(h), sectionName, LV
                capCounts(idx) = capCounts(idx) + 1
                remaining = remaining - 1
                toAdd = toAdd - 1
                If remaining <= 0 Then Exit Do
            Loop
        Next h

nextI:
    Next i
    Exit Sub
fail:
    MsgBox "AddSectionItems error: " & Err.Description, vbExclamation
End Sub


'====================[ DB Lookup (no Variants out) ]====================
' Fills hits() with exact-name matches where [In Game] <> 0.
' Returns count (0 if none).
' Fills hits() with exact-name matches where [In Game] <> 0.
' Returns count (0 if none). Works with forward-only / snapshot / dynaset.
Private Function GetItemsByExactNameArr(ByVal exactName As String, ByRef hits() As ItemMatch) As Long
    On Error GoTo fail

    Dim rs As DAO.Recordset
    Dim cnt As Long
    Dim itm As ItemMatch
    Dim nm As String
    Dim inGame As Long

    If tabItems Is Nothing Then Exit Function

    Set rs = tabItems.Clone
    If (rs.BOF And rs.EOF) Then
        rs.Close
        Exit Function
    End If

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
                itm.WeaponType = NzLong(rs.Fields("WeaponType").Value)   ' <— NEW
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


'====================[ Row Creation ]====================
Private Sub AddListViewRowsForItem(ByRef hit As ItemMatch, ByRef sectionName As String, ByRef LV As ListView)
    On Error GoTo fail

    Dim shops() As ShopToken
    Dim shopCnt As Long
    Dim s As Long
    Dim shopName As String
    Dim valStr As String
    Dim nCharm As Integer
    Dim itemVal As tItemValue
    
    nCharm = val(frmMain.txtCharStats(5).Text)
    
    shopCnt = ExtractShopsFromObtainedFrom(hit.ObtainedFrom, shops)

    If shopCnt = 0 Then
        AddOneRow LV, hit, sectionName, "none", "(no value)"
    Else
        For s = 0 To shopCnt - 1
            shopName = GetShopName(shops(s).ShopNumber, True)
            itemVal = GetItemValue(hit.Number, nCharm, 0, shops(s).ShopNumber, shops(s).bNObuy)
            If itemVal.nBaseCost > 0 Then
                If shops(s).bNObuy Then
                    valStr = itemVal.sFriendlySell & " (sell)"
                Else
                    valStr = itemVal.sFriendlyBuy & " / " & itemVal.sFriendlySell
                End If
            Else
                valStr = ""
            End If
            AddOneRow LV, hit, sectionName, shopName, valStr
        Next s
    End If
    Exit Sub
fail:
    MsgBox "AddListViewRowsForItem error: " & Err.Description, vbExclamation
End Sub

Private Sub AddOneRow(ByRef LV As ListView, ByRef hit As ItemMatch, ByVal sectionName As String, _
                      ByVal shopCell As String, ByVal valueCell As String)
    Dim oLI As ListItem
    Dim wornText As String
    Dim usableText As String

    ' Worn text by ItemType:
    Select Case hit.ItemType
        Case 0          ' Armour
            wornText = GetWornType(hit.Worn)
        Case 1          ' Weapon
            wornText = GetWeaponType(hit.WeaponType)
        Case Else
            wornText = "Nowhere"
    End Select

    ' Usable: Yes/No
    If frmMain.TestGlobalFilter(hit.Number) Then
        usableText = "Yes"
    Else
        usableText = "No"
    End If

    Set oLI = LV.ListItems.Add()
    oLI.Text = CStr(hit.Number)                                ' Col 1: Number

    oLI.ListSubItems.Add 1, "Name", hit.name                   ' Col 2
    oLI.ListSubItems.Add 2, "Source", sectionName              ' Col 3
    oLI.ListSubItems.Add 3, "Type", GetItemType(hit.ItemType)  ' Col 4
    oLI.ListSubItems.Add 4, "Worn", wornText                   ' Col 5
    oLI.ListSubItems.Add 5, "Usable", usableText               ' Col 6
    oLI.ListSubItems.Add 6, "Shop", shopCell                   ' Col 7
    oLI.ListSubItems.Add 7, "Value", valueCell                 ' Col 8

    ' store copper cost into Value subitem's Tag for later sorting
    oLI.ListSubItems(7).Tag = CStr(GetItemCopperCost(hit.Number))
End Sub


'====================[ Parsing Helpers ]====================
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

'====================[ Shop Parsing (array out) ]====================
' Returns count and fills shops() with matches from "Obtained From"
' Recognizes: "shop #", "shop(sell) #", "shop(nogen) #"
Private Function ExtractShopsFromObtainedFrom(ByVal obtained As String, ByRef shops() As ShopToken) As Long
    Dim parts() As String, i As Long, t As String, n As Long
    Dim tok As ShopToken
    Dim cnt As Long

    If Len(Trim$(obtained)) = 0 Then
        ExtractShopsFromObtainedFrom = 0
        Exit Function
    End If

    parts = Split(obtained, ",")
    For i = LBound(parts) To UBound(parts)
        t = LCase$(Trim$(parts(i)))
        If Left$(t, 4) = "shop" Then
            tok.bNObuy = False: tok.bNoSell = False
            If InStr(t, "(sell)") > 0 Then
                tok.bNObuy = True
            End If
            ' (nogen) still allows buy/sell; flags remain False

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

'====================[ Small Utils (typed) ]====================
Private Function SQLQuote(ByVal s As String) As String
    SQLQuote = "'" & Replace$(s, "'", "''") & "'"
End Function

Private Function NzStr(ByVal v As Variant) As String
    If IsNull(v) Then NzStr = "" Else NzStr = CStr(v)
End Function

Private Function NzLong(ByVal v As Variant) As Long
    If IsNull(v) Or v = "" Then
        NzLong = 0
    Else
        NzLong = CLng(v)
    End If
End Function


