Attribute VB_Name = "modQuestConfig"
'======================================================================
' modQuestConfig (format v2) - data-driven "Completed Quests" rewards
'----------------------------------------------------------------------
' Phase 2: quests are a dynamic, ordered list identified by a stable
' Key (not a fixed 0..11 slot). Built-ins load from MME-Quests.txt
' (recreated from the embedded default if missing). MME-QuestsCustom.txt
' overrides a built-in (same Key) or appends a new quest (new Key).
'
' Save/paste persistence uses the Key, so adding/reordering quests never
' breaks a saved character. The stat applier lives in frmMain
' (ApplyQuestRewards / ApplyOneQuestTerm); the list UI lives in
' modQuestUI.
'======================================================================
Option Explicit

Public Const QPASS_PREENCUM As Integer = 0
Public Const QPASS_MAIN As Integer = 1
Public Const QOPT_MAX As Integer = 9

Public Type QuestTerm
    Target As String          ' "sN" | "accy" | "accymax" | "dodge" | "str"
    Value As Long
End Type

Public Type QuestDef
    Key As String             ' stable id (persistence)
    Builtin As Boolean        ' from the shipped default file
    Dirty As Boolean          ' edited at runtime -> save to custom file
    SlotIndex As Integer      ' legacy built-in slot 0..11 (-1 if custom)
    ChoiceField As Integer    ' legacy combo field 0..3 (-1 if none / custom)
    HasChoice As Boolean
    Name As String
    Engine As String          ' both | stock | gmud
    Reward As String          ' simple-quest term string
    OptionCount As Integer
    OptionLabel(0 To QOPT_MAX) As String
    OptionStr(0 To QOPT_MAX) As String
    OptionExport(0 To QOPT_MAX) As String
    ImportRule As String      ' "abil:minval"
    ExportRule As String      ' "abil|val"
End Type

Public g_Quests() As QuestDef
Public g_QuestCount As Integer
Public g_QuestsLoaded As Boolean

'----------------------------------------------------------------------
Public Sub LoadQuestDefs()
    Dim sBase As String, sMain As String, sCustom As String
    On Error GoTo fail

    If Right$(App.Path, 1) = "\" Then sBase = App.Path Else sBase = App.Path & "\"
    sMain = sBase & "MME-Quests.txt"
    sCustom = sBase & "MME-QuestsCustom.txt"

    g_QuestCount = 0
    ReDim g_Quests(0 To 31)

    If Len(Dir$(sMain)) = 0 Then WriteDefaultQuestFile sMain
    ParseQuestFileInto sMain, True
    If Len(Dir$(sCustom)) > 0 Then ParseQuestFileInto sCustom, False

    g_QuestsLoaded = True
    Exit Sub
fail:
    g_QuestsLoaded = True       ' fail-safe: whatever parsed stays usable
End Sub

'----------------------------------------------------------------------
Public Function QuestIndexByKey(ByVal sKey As String) As Integer
    Dim i As Integer
    QuestIndexByKey = -1
    For i = 0 To g_QuestCount - 1
        If StrComp(g_Quests(i).Key, sKey, vbTextCompare) = 0 Then QuestIndexByKey = i: Exit Function
    Next i
End Function

Public Function QuestIndexBySlot(ByVal nSlot As Integer) As Integer
    Dim i As Integer
    QuestIndexBySlot = -1
    For i = 0 To g_QuestCount - 1
        If g_Quests(i).Builtin And g_Quests(i).SlotIndex = nSlot Then QuestIndexBySlot = i: Exit Function
    Next i
End Function

'----------------------------------------------------------------------
Private Function AddOrGetQuest(ByVal sKey As String, ByVal bBuiltinFile As Boolean) As Integer
    Dim idx As Integer, j As Integer
    idx = QuestIndexByKey(sKey)
    If idx >= 0 Then AddOrGetQuest = idx: Exit Function
    If g_QuestCount > UBound(g_Quests) Then ReDim Preserve g_Quests(0 To g_QuestCount + 15)
    idx = g_QuestCount
    g_Quests(idx).Key = sKey
    g_Quests(idx).Builtin = bBuiltinFile
    g_Quests(idx).SlotIndex = -1
    g_Quests(idx).ChoiceField = -1
    g_Quests(idx).HasChoice = False
    g_Quests(idx).Engine = "both"
    g_Quests(idx).OptionCount = 0
    For j = 0 To QOPT_MAX
        g_Quests(idx).OptionLabel(j) = ""
        g_Quests(idx).OptionStr(j) = ""
        g_Quests(idx).OptionExport(j) = ""
    Next j
    g_QuestCount = g_QuestCount + 1
    AddOrGetQuest = idx
End Function

'----------------------------------------------------------------------
Private Sub ParseQuestFileInto(ByVal sPath As String, ByVal bBuiltinFile As Boolean)
    Dim f As Integer, ln As String, p As Integer
    Dim sectID As String, sectKey As String, haveSect As Boolean
    Dim k As String, v As String, idx As Integer, optIdx As Integer

    ' Two-pass per block: buffer fields until we know the Key, then commit.
    Dim bufName As String, bufEngine As String, bufReward As String
    Dim bufImport As String, bufExport As String, bufChoice As Integer
    Dim bufHasChoice As Boolean
    Dim bufOptL(0 To QOPT_MAX) As String, bufOptT(0 To QOPT_MAX) As String
    Dim bufOptExp(0 To QOPT_MAX) As String
    Dim bufOptCount As Integer, i As Integer

    f = FreeFile
    Open sPath For Input As #f
    Do
        If EOF(f) Then
            If haveSect Then CommitBlock bBuiltinFile, sectID, sectKey, bufName, bufEngine, bufReward, _
                bufImport, bufExport, bufChoice, bufHasChoice, bufOptL, bufOptT, bufOptExp, bufOptCount
            Exit Do
        End If
        Line Input #f, ln
        p = InStr(ln, "#"): If p > 0 Then ln = Left$(ln, p - 1)
        ln = Trim$(ln)
        If Len(ln) > 0 Then
            If Left$(ln, 1) = "[" Then
                If haveSect Then CommitBlock bBuiltinFile, sectID, sectKey, bufName, bufEngine, bufReward, _
                    bufImport, bufExport, bufChoice, bufHasChoice, bufOptL, bufOptT, bufOptExp, bufOptCount
                ' reset buffers
                bufName = "": bufEngine = "both": bufReward = "": bufImport = "": bufExport = ""
                bufChoice = -1: bufHasChoice = False: bufOptCount = 0
                For i = 0 To QOPT_MAX: bufOptL(i) = "": bufOptT(i) = "": bufOptExp(i) = "": Next i
                sectKey = ""
                p = InStr(ln, "]")
                sectID = Trim$(Mid$(ln, 2, p - 2))
                If LCase$(Left$(sectID, 6)) = "quest " Then sectID = Trim$(Mid$(sectID, 7))
                haveSect = True
            ElseIf haveSect Then
                p = InStr(ln, "=")
                If p > 0 Then
                    k = LCase$(Trim$(Left$(ln, p - 1)))
                    v = Trim$(Mid$(ln, p + 1))
                    Select Case k
                        Case "key": sectKey = v
                        Case "name": bufName = v
                        Case "engine": bufEngine = LCase$(v)
                        Case "reward": bufReward = v
                        Case "import": bufImport = v
                        Case "export": bufExport = v
                        Case "choice": bufChoice = CInt(Val(v)): bufHasChoice = True
                        Case Else
                            If Left$(k, 6) = "option" Then
                                optIdx = CInt(Val(Mid$(k, 7)))
                                If optIdx >= 0 And optIdx <= QOPT_MAX Then
                                    If InStr(k, "_export") > 0 Then
                                        bufOptExp(optIdx) = v
                                    Else
                                        SplitLabelTerms v, bufOptL(optIdx), bufOptT(optIdx)
                                    End If
                                    bufHasChoice = True
                                    If optIdx + 1 > bufOptCount Then bufOptCount = optIdx + 1
                                End If
                            End If
                    End Select
                End If
            End If
        End If
    Loop
    Close #f
End Sub

Private Sub CommitBlock(ByVal bBuiltinFile As Boolean, ByVal sectID As String, ByVal sectKey As String, _
    ByVal nm As String, ByVal eng As String, ByVal rw As String, ByVal imp As String, ByVal exp As String, _
    ByVal ch As Integer, ByVal hasCh As Boolean, ByRef optL() As String, ByRef optT() As String, ByRef optExp() As String, ByVal optC As Integer)

    Dim key As String, idx As Integer, i As Integer, isNum As Boolean
    isNum = (Len(sectID) > 0 And IsNumeric(sectID))
    key = sectKey
    If Len(key) = 0 Then If isNum Then key = "slot" & sectID Else key = sectID
    If Len(key) = 0 Then Exit Sub

    idx = AddOrGetQuest(key, bBuiltinFile)
    g_Quests(idx).Name = nm
    g_Quests(idx).Engine = eng
    g_Quests(idx).Reward = rw
    g_Quests(idx).ImportRule = imp
    g_Quests(idx).ExportRule = exp
    g_Quests(idx).HasChoice = hasCh
    g_Quests(idx).ChoiceField = ch
    g_Quests(idx).OptionCount = optC
    For i = 0 To QOPT_MAX
        g_Quests(idx).OptionLabel(i) = optL(i)
        g_Quests(idx).OptionStr(i) = optT(i)
        g_Quests(idx).OptionExport(i) = optExp(i)
    Next i
    If isNum Then g_Quests(idx).SlotIndex = CInt(sectID)
    If Not bBuiltinFile Then g_Quests(idx).Dirty = True   ' came from custom file
End Sub

'----------------------------------------------------------------------
Private Sub SplitLabelTerms(ByVal v As String, ByRef outLabel As String, ByRef outTerms As String)
    Dim p As Integer
    p = InStr(v, "|")
    If p > 0 Then
        outLabel = Trim$(Left$(v, p - 1))
        outTerms = Trim$(Mid$(v, p + 1))
    Else
        outLabel = ""           ' phase-1 style: whole RHS is terms
        outTerms = Trim$(v)
    End If
End Sub

'----------------------------------------------------------------------
' Return the terms string for a quest (simple) or a chosen option.
Public Function QuestExportCodes(ByVal qi As Integer, ByVal optIdx As Integer, ByVal bGmud As Boolean) As String
    Dim raw As String, parts() As String, i As Integer, code As String, s As String
    If qi < 0 Or qi > g_QuestCount - 1 Then Exit Function
    If g_Quests(qi).Engine = "gmud" And Not bGmud Then Exit Function
    If g_Quests(qi).Engine = "stock" And bGmud Then Exit Function
    If g_Quests(qi).HasChoice Then
        If optIdx < 0 Or optIdx > QOPT_MAX Then Exit Function
        raw = g_Quests(qi).OptionExport(optIdx)
    Else
        raw = g_Quests(qi).ExportRule
    End If
    If Len(Trim$(raw)) = 0 Then Exit Function
    parts = Split(raw, ",")
    For i = 0 To UBound(parts)
        code = Trim$(parts(i))
        If Left$(code, 6) = "[gmud]" Then
            If Not bGmud Then code = "" Else code = Trim$(Mid$(code, 7))
        ElseIf Left$(code, 7) = "[stock]" Then
            If bGmud Then code = "" Else code = Trim$(Mid$(code, 8))
        End If
        If Len(code) > 0 Then
            If Len(s) > 0 Then s = s & ","
            s = s & code
        End If
    Next i
    QuestExportCodes = s
End Function

Public Function QuestImportMatches(ByVal qi As Integer, ByVal code As Long, ByVal v As Long, ByVal bGmud As Boolean) As Boolean
    Dim r As String, p As Integer
    If qi < 0 Or qi > g_QuestCount - 1 Then Exit Function
    r = g_Quests(qi).ImportRule
    If Len(Trim$(r)) = 0 Then Exit Function
    If g_Quests(qi).Engine = "gmud" And Not bGmud Then Exit Function
    If g_Quests(qi).Engine = "stock" And bGmud Then Exit Function
    p = InStr(r, ":")
    If p <= 0 Then Exit Function
    QuestImportMatches = (code = CLng(Val(Left$(r, p - 1))) And v >= CLng(Val(Mid$(r, p + 1))))
End Function

Public Function OptionTerms(ByVal qi As Integer, ByVal optIdx As Integer) As String
    If qi < 0 Or qi > g_QuestCount - 1 Then Exit Function
    If g_Quests(qi).HasChoice Then
        If optIdx >= 0 And optIdx <= QOPT_MAX Then OptionTerms = g_Quests(qi).OptionStr(optIdx)
    Else
        OptionTerms = g_Quests(qi).Reward
    End If
End Function

'----------------------------------------------------------------------
' Parse a term string, engine-filtered. Fills terms() (0-based), returns count.
Public Function ParseQuestTerms(ByVal sTerms As String, ByVal sEngine As String, _
                                ByRef terms() As QuestTerm) As Integer
    Dim parts() As String, i As Integer, t As String, q As String, c As Integer, n As Integer
    n = 0
    ReDim terms(0 To 31)
    If Len(Trim$(sTerms)) = 0 Then ParseQuestTerms = 0: Exit Function
    parts = Split(sTerms, ",")
    For i = 0 To UBound(parts)
        t = Trim$(parts(i))
        If Len(t) > 0 Then
            If Left$(t, 1) = "[" Then
                c = InStr(t, "]")
                If c > 0 Then
                    q = LCase$(Trim$(Mid$(t, 2, c - 2)))
                    t = Trim$(Mid$(t, c + 1))
                    If q <> LCase$(sEngine) Then GoTo nextp
                End If
            End If
            c = InStr(t, ":")
            If c > 0 Then
                terms(n).Target = LCase$(Trim$(Left$(t, c - 1)))
                terms(n).Value = CLng(Val(Mid$(t, c + 1)))
                n = n + 1
            End If
        End If
nextp:
    Next i
    ParseQuestTerms = n
End Function

'----------------------------------------------------------------------
' Short human summary of a term set, e.g. "+1 AC, +2 SC" (for list captions).
Public Function SummaryForTerms(ByVal sTerms As String, ByVal sEngine As String) As String
    Dim terms() As QuestTerm, n As Integer, i As Integer, s As String, lbl As String, vl As Long
    n = ParseQuestTerms(sTerms, sEngine, terms())
    For i = 0 To n - 1
        vl = terms(i).Value
        Select Case terms(i).Target
            Case "s2": lbl = "AC"
            Case "s4": lbl = "% Enc"
            Case "s5": lbl = "HP"
            Case "s6": lbl = "Mana"
            Case "s7": lbl = "Crit"
            Case "s9": lbl = "SC"
            Case "s11": lbl = "Max Dmg"
            Case "s14": lbl = "BS Min"
            Case "s15": lbl = "BS Max"
            Case "s17": lbl = "ManaRgn"
            Case "s19": lbl = "Stealth"
            Case "accy", "accymax": lbl = "Accy"
            Case "dodge": lbl = "DG"
            Case "str": lbl = "Str"
            Case Else: lbl = terms(i).Target
        End Select
        s = AutoAppendLocal(s, "+" & vl & " " & lbl, ", ")
    Next i
    SummaryForTerms = s
End Function

Private Function AutoAppendLocal(ByVal sBase As String, ByVal sAdd As String, ByVal sSep As String) As String
    If Len(sBase) = 0 Then AutoAppendLocal = sAdd Else AutoAppendLocal = sBase & sSep & sAdd
End Function

'----------------------------------------------------------------------
' Persist custom + edited quests to MME-QuestsCustom.txt.
Public Sub SaveCustomQuests()
    Dim sBase As String, sPath As String, f As Integer, i As Integer, j As Integer
    On Error Resume Next
    If Right$(App.Path, 1) = "\" Then sBase = App.Path Else sBase = App.Path & "\"
    sPath = sBase & "MME-QuestsCustom.txt"
    f = FreeFile
    Open sPath For Output As #f
    Print #f, "# MME-QuestsCustom.txt - per-realm overrides/additions (auto-saved)."
    Print #f, "# Same format as MME-Quests.txt. Blocks here override matching Keys."
    Print #f, ""
    For i = 0 To g_QuestCount - 1
        If (Not g_Quests(i).Builtin) Or g_Quests(i).Dirty Then
            If g_Quests(i).Builtin Then
                Print #f, "[Quest " & g_Quests(i).SlotIndex & "]"
            Else
                Print #f, "[Quest " & g_Quests(i).Key & "]"
            End If
            Print #f, "Key = " & g_Quests(i).Key
            Print #f, "Name = " & g_Quests(i).Name
            Print #f, "Engine = " & g_Quests(i).Engine
            If Len(g_Quests(i).ImportRule) > 0 Then Print #f, "Import = " & g_Quests(i).ImportRule
            If Len(g_Quests(i).ExportRule) > 0 Then Print #f, "Export = " & g_Quests(i).ExportRule
            If g_Quests(i).HasChoice Then
                If g_Quests(i).ChoiceField >= 0 Then Print #f, "Choice = " & g_Quests(i).ChoiceField
                For j = 0 To g_Quests(i).OptionCount - 1
                    Print #f, "Option" & j & " = " & g_Quests(i).OptionLabel(j) & " | " & g_Quests(i).OptionStr(j)
                    If Len(g_Quests(i).OptionExport(j)) > 0 Then Print #f, "Option" & j & "_Export = " & g_Quests(i).OptionExport(j)
                Next j
            Else
                Print #f, "Reward = " & g_Quests(i).Reward
            End If
            Print #f, ""
        End If
    Next i
    Close #f
End Sub

'----------------------------------------------------------------------
Private Sub WriteDefaultQuestFile(ByVal sPath As String)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open sPath For Output As #f
    Print #f, DefaultQuestText()
    Close #f
End Sub

Private Function DefaultQuestText() As String
Dim s As String, nl As String
nl = vbCrLf
s = "# =====================================================================" & nl
s = s & "#  MME-Quests.txt  --  Completed-Quest reward definitions (format v2)" & nl
s = s & "#  for the MMUD Explorer ""Completed Quests"" character panel." & nl
s = s & "# ---------------------------------------------------------------------" & nl
s = s & "#  This is the SHIPPED DEFAULT set. It is recreated automatically if" & nl
s = s & "#  missing, so behaviour is identical out of the box. Edit freely." & nl
s = s & "#" & nl
s = s & "#  Per-realm additions/overrides go in MME-QuestsCustom.txt (same" & nl
s = s & "#  format). A block there with the same Key overrides the built-in;" & nl
s = s & "#  a block with a new Key is appended as a new quest in the list." & nl
s = s & "#" & nl
s = s & "#  QUEST BLOCK" & nl
s = s & "#    [Quest <id>]            id = built-in slot number (0..11) or any slug" & nl
s = s & "#    Key    = <slug>         stable id used for save/paste (never reorder-broken)" & nl
s = s & "#    Name   = <text>         shown in the list and in stat tooltips" & nl
s = s & "#    Engine = both|stock|gmud" & nl
s = s & "#    Choice = <comboField>   0..3 selects which option set; omit = simple quest" & nl
s = s & "#    Import = <abil>:<min>   (optional) paste ability id that ticks this quest" & nl
s = s & "#    Export = <abil>|<val>   (optional) ability emitted when exporting char" & nl
s = s & "#    Reward = <terms>        simple quest" & nl
s = s & "#    Option<k> = <label> | <terms>   choice quest, k = dropdown index" & nl
s = s & "#" & nl
s = s & "#  TERMS (comma separated)  <target>:<value>" & nl
s = s & "#    sN  accy  accymax  dodge  str        (slot N or special target)" & nl
s = s & "#    engine qualifier prefix: [stock] or [gmud]   (no prefix = both)" & nl
s = s & "#  SLOTS: 2=AC 4=Encum 5=MaxHP 6=Mana 7=Crits 9=SC 11=MaxDmg" & nl
s = s & "#         14=BSMin 15=BSMax 17=ManaRegen 19=Stealth" & nl
s = s & "# =====================================================================" & nl
s = s & nl
s = s & "[Quest 0]" & nl
s = s & "Key = ice_sorceress" & nl
s = s & "Name = Ice Sorceress" & nl
s = s & "Engine = both" & nl
s = s & "Import = 125:1" & nl
s = s & "Export = 125|2" & nl
s = s & "Reward = s2:1" & nl
s = s & nl
s = s & "[Quest 1]" & nl
s = s & "Key = high_druid" & nl
s = s & "Name = High Druid" & nl
s = s & "Engine = both" & nl
s = s & "Import = 129:1" & nl
s = s & "Export = 129|2" & nl
s = s & "Reward = s9:1" & nl
s = s & nl
s = s & "[Quest 2]" & nl
s = s & "Key = adult_red_dragon" & nl
s = s & "Name = Adult Red Dragon" & nl
s = s & "Engine = both" & nl
s = s & "Import = 131:2" & nl
s = s & "Export = 131|3" & nl
s = s & "Reward = s7:1, s9:2" & nl
s = s & nl
s = s & "[Quest 3]" & nl
s = s & "Key = bishop" & nl
s = s & "Name = Bishop" & nl
s = s & "Engine = both" & nl
s = s & "Import = 130:1" & nl
s = s & "Export = 130|2" & nl
s = s & "Reward = [stock]accymax:3, [gmud]accy:3" & nl
s = s & nl
s = s & "[Quest 4]" & nl
s = s & "Key = apparatus" & nl
s = s & "Name = Apparatus" & nl
s = s & "Engine = both" & nl
s = s & "Import = 132:1" & nl
s = s & "Export = 132|2" & nl
s = s & "Reward = dodge:1" & nl
s = s & nl
s = s & "[Quest 5]" & nl
s = s & "Key = 2nd_align" & nl
s = s & "Name = 2nd Align" & nl
s = s & "Engine = both" & nl
s = s & "Choice = 0" & nl
s = s & "Option0 = (none) |" & nl
s = s & "Option1 = +1 Max Damage (+5 Accy in GMUD) | s11:1, [gmud]accy:5" & nl
s = s & "Option2 = +1 AC, +6 Mana | s2:1, s6:6" & nl
s = s & "Option3 = +1 SC / +5 ManaRgn, +10 Mana | s6:10, [stock]s9:1, [gmud]s17:5" & nl
s = s & "Option4 = +4 Mana, +6 BS Min/Max, +1 Stealth | s6:4, s14:6, s15:6, s19:1" & nl
s = s & "Option5 = +10 BS Min/Max, +2 Stealth | s14:10, s15:10, s19:2" & nl
s = s & "Option1_Export = 4|1, [gmud]22|5" & nl
s = s & "Option2_Export = 69|6, 2|1" & nl
s = s & "Option3_Export = 69|10, 70|1" & nl
s = s & "Option4_Export = 69|4, 27|1, 117|6, 118|6" & nl
s = s & "Option5_Export = 27|2, 117|10, 118|10" & nl
s = s & nl
s = s & "[Quest 6]" & nl
s = s & "Key = opaline" & nl
s = s & "Name = Opaline" & nl
s = s & "Engine = gmud" & nl
s = s & "Import = 200:6" & nl
s = s & "Export = 88|100, 200|6" & nl
s = s & "Reward = s5:100" & nl
s = s & nl
s = s & "[Quest 7]" & nl
s = s & "Key = cartographer" & nl
s = s & "Name = Cartographer" & nl
s = s & "Engine = gmud" & nl
s = s & "Import = 203:1" & nl
s = s & "Export = 96|3, 203|1" & nl
s = s & "Reward = s4:3" & nl
s = s & nl
s = s & "[Quest 8]" & nl
s = s & "Key = loremaster" & nl
s = s & "Name = Loremaster" & nl
s = s & "Engine = gmud" & nl
s = s & "Import = 202:1" & nl
s = s & "Export = 2|1, 202|1" & nl
s = s & "Reward = s2:1" & nl
s = s & nl
s = s & "[Quest 9]" & nl
s = s & "Key = 6th_align" & nl
s = s & "Name = 6th Align" & nl
s = s & "Engine = gmud" & nl
s = s & "Choice = 1" & nl
s = s & "Option0 = (none) |" & nl
s = s & "Option1 = War/WHunter/Paladin (+5 Accy,+1 MaxDmg,+5% Enc,+50 HP) | accy:5, s11:1, s4:5, s5:50" & nl
s = s & "Option2 = Cleric/Warlock (+2 AC,+5 ManaRgn,+3% Enc,+50 HP) | s2:2, s17:5, s4:3, s5:50" & nl
s = s & "Option3 = Priest/Mage/Druid (+50 SC,+10 ManaRgn,+50 HP) | s9:50, s17:10, s5:50" & nl
s = s & "Option4 = Missy/Bard/Gypsy (+10 BS,+5 ManaRgn,+10 Stl,+50 HP) | s14:10, s15:10, s17:5, s19:10, s5:50" & nl
s = s & "Option5 = Thief/Ninja/Ranger (+15 BS,+10 Stl,+50 HP) | s14:15, s15:15, s19:10, s5:50" & nl
s = s & "Option6 = Mystic (+10 Accy,+1 MaxDmg,+1 Crit,+50 HP) | accy:10, s11:1, s7:1, s5:50" & nl
s = s & "Option1_Export = 22|5, 4|1, 96|5, 88|50" & nl
s = s & "Option2_Export = 2|2, 145|5, 96|3, 88|50" & nl
s = s & "Option3_Export = 70|25, 145|10, 88|50" & nl
s = s & "Option4_Export = 117|10, 118|10, 145|5, 27|10, 88|50" & nl
s = s & "Option5_Export = 117|15, 118|15, 27|10, 88|50" & nl
s = s & "Option6_Export = 22|10, 4|1, 58|1, 88|50" & nl
s = s & nl
s = s & "[Quest 10]" & nl
s = s & "Key = dread_wraith" & nl
s = s & "Name = Dread Wraith" & nl
s = s & "Engine = gmud" & nl
s = s & "Choice = 2" & nl
s = s & "Option0 = (none) |" & nl
s = s & "Option1 = War/Pal/Cler/Missy/Ninja/Thief/Bard/Gypsy/WL/Rngr/Mystic (+1 AC,+1 Crit) | s2:1, s7:1" & nl
s = s & "Option2 = Witchhunter (+1 AC,+2 Crit) | s2:1, s7:2" & nl
s = s & "Option3 = Priest/Mage/Druid (+1 AC) | s2:1" & nl
s = s & "Option1_Export = 2|1, 58|1" & nl
s = s & "Option2_Export = 2|1, 58|2" & nl
s = s & "Option3_Export = 2|1" & nl
s = s & nl
s = s & "[Quest 11]" & nl
s = s & "Key = renfry" & nl
s = s & "Name = Renfry" & nl
s = s & "Engine = gmud" & nl
s = s & "Choice = 3" & nl
s = s & "Option0 = (none) |" & nl
s = s & "Option1 = 1st: +10% Encum, +1 Max Dmg | s4:10, s11:1" & nl
s = s & "Option2 = 2nd: +10 Strength (plus 1st) | s4:10, str:10, s11:1" & nl
s = s & "Option1_Export = 96|10, 4|1" & nl
s = s & "Option2_Export = 96|10, 4|1, 46|10" & nl
s = s & nl
s = s & "# ---------------------------------------------------------------------" & nl
s = s & "# CUSTOM QUEST TEMPLATE (copy into MME-QuestsCustom.txt and edit):" & nl
s = s & "#" & nl
s = s & "# [Quest my_realm_quest]" & nl
s = s & "# Key    = my_realm_quest" & nl
s = s & "# Name   = My Realm Quest" & nl
s = s & "# Engine = both" & nl
s = s & "# Import = 250:1" & nl
s = s & "# Reward = s2:2, s5:25" & nl
s = s & "#" & nl
s = s & "# ...or a choice quest:" & nl
s = s & "# [Quest my_choice_quest]" & nl
s = s & "# Key    = my_choice_quest" & nl
s = s & "# Name   = My Choice Quest" & nl
s = s & "# Engine = gmud" & nl
s = s & "# Choice = 0" & nl
s = s & "# Option0 = (none) |" & nl
s = s & "# Option1 = Warriors (+2 AC) | s2:2" & nl
s = s & "# Option2 = Casters (+25 SC) | s9:25" & nl
s = s & "# ---------------------------------------------------------------------" & nl
DefaultQuestText = s
End Function