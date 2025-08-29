Attribute VB_Name = "CheshireCAT"
' ========= INVIO TESTO E INSERIMENTO RISPOSTA =========
' ========= INVIO TESTO E INSERIMENTO RISPOSTA =========
Public Sub InviaTestoAChat()
    Dim selectedText As String
    Dim response As String
    Dim startRange As Range

    If Selection.Type = wdSelectionIP Then
        MsgBox "Nessun testo selezionato", vbExclamation
        Exit Sub
    End If

    'Costruisce un payload pulito: testo normale + tabelle convertite in Markdown
    selectedText = BuildMessageFromSelection(Selection.Range)
    If Len(Trim$(selectedText)) = 0 Then
        MsgBox "La selezione è vuota dopo la normalizzazione.", vbExclamation
        Exit Sub
    End If

    response = CheshireCat_Chat(selectedText) ' <-- modulo API

    Set startRange = Selection.Range
    startRange.Collapse Direction:=wdCollapseEnd
    startRange.Select

    Selection.InsertAfter vbNewLine
    Selection.Collapse Direction:=wdCollapseEnd

    ' Inserisce subito testo + tabelle formattate
    InsertAIResponseWithMarkdownTables response
End Sub


Public Sub CancellaCronologiaChat()
    Dim jwtToken As String
    Dim success As Boolean
    
    jwtToken = GetJWToken() ' <-- dal modulo API
    If Left$(jwtToken, 6) = "Errore" Then
        MsgBox "Errore durante il recupero del token: " & jwtToken
        Exit Sub
    End If
    
    success = ClearChatHistory(jwtToken) ' <-- dal modulo API
    If success Then
        MsgBox "Cronologia cancellata con successo!"
    Else
        MsgBox "Errore durante la cancellazione della cronologia."
    End If
End Sub

' ========= CONVERSIONE / INSERIMENTO TABELLE MARKDOWN =========

' Inserisce testo e converte TUTTE le tabelle Markdown trovate (2a, 3a, ...).
Private Sub InsertAIResponseWithMarkdownTables(ByVal response As String)
    Dim rest As String
    Dim pre As String, tbl As String, post_ As String
    Dim hasTable As Boolean
    Dim safety As Long
    
    ' Normalizza i fine riga
    rest = NormalizeToLf(Replace(response, "\n", vbLf))
    
    safety = 0 ' guardia anti-loop
    Do
        hasTable = ExtractFirstMarkdownTableBlock(rest, pre, tbl, post_)
        
        If hasTable = False Then
            ' Nessuna (ulteriore) tabella: inserisce tutto il residuo come testo
            If Len(pre) > 0 Then InsertMarkdownInlineToSelection pre, False, False
            Exit Do
        End If
        
        ' 1) Testo prima della tabella corrente
        If Len(pre) > 0 Then
            InsertMarkdownInlineToSelection pre, False, False
        End If
        
        ' 2) Tabella corrente
        EnsureParagraphBeforeInsertion
        CreateAndInsertWordTableFromMarkdown tbl, Selection.Range
        
        ' 3) Se c'è ancora del testo dopo, aggiungi una riga di separazione
        rest = post_
        If Len(rest) > 0 Then Selection.TypeParagraph
        
        safety = safety + 1
        If safety > 20 Then Exit Do ' estrema sicurezza
    Loop
End Sub

' Inserisce un a-capo se il punto di inserimento non è già all'inizio di una nuova riga
Private Sub EnsureParagraphBeforeInsertion()
    Dim lastChar As String
    lastChar = ""
    On Error Resume Next
    lastChar = Right$(Selection.Range.text, 1)
    On Error GoTo 0
    If lastChar <> vbCr Then Selection.TypeParagraph
End Sub


' === Strumento manuale: seleziona tabella Markdown e convertila in tabella Word
Public Sub ConvertiTabellaMarkdown()
    Dim selRng As Range
    Dim md As String
    
    If Selection Is Nothing Then
        MsgBox "Nessuna selezione.", vbExclamation
        Exit Sub
    End If
    
    Set selRng = Selection.Range
    md = Trim(GetSelectedMarkdownTableText(selRng))
    
    If Len(md) = 0 Then
        MsgBox "Seleziona il testo della tabella Markdown (righe con '|' ).", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo EH
    ConvertMarkdownToWord md, selRng
    MsgBox "Tabella convertita con successo!", vbInformation
    Exit Sub
EH:
    MsgBox "Errore durante la conversione: " & Err.Description, vbCritical
End Sub

Private Function GetSelectedMarkdownTableText(ByVal rng As Range) As String
    Dim tx As String, lines() As String, i As Long, out As String
    
    tx = rng.text
    tx = Replace(tx, vbCrLf, vbLf)
    tx = Replace(tx, vbCr, vbLf)
    
    lines = Split(tx, vbLf)
    For i = LBound(lines) To UBound(lines)
        Dim ln As String
        ln = Trim$(lines(i))
        If ln <> "```" And LCase$(ln) <> "```markdown" And ln <> "" Then
            out = out & ln & vbLf
        End If
    Next i
    
    If Right$(out, 1) = vbLf Then out = Left$(out, Len(out) - 1)
    GetSelectedMarkdownTableText = out
End Function

Private Sub ConvertMarkdownToWord(ByVal markdown As String, ByVal targetRange As Range)
    Dim rawLines() As String, lines() As String
    Dim i As Long, j As Long, rowCount As Long, colCount As Long
    Dim wordTable As Table
    Dim sepLineIndex As Long
    Dim actualRow As Long
    Dim rowVals() As String

    rawLines = Split(markdown, vbLf)
    lines = FilterNonEmpty(rawLines)

    rowCount = UBound(lines) - LBound(lines) + 1
    If rowCount < 2 Then Err.Raise vbObjectError + 1, , "Tabella Markdown non valida (righe insufficienti)."

    colCount = CountMarkdownColumns(lines(0))
    If colCount < 1 Then Err.Raise vbObjectError + 2, , "Impossibile determinare il numero di colonne."

    sepLineIndex = FindSeparatorIndex(lines, colCount)
    If sepLineIndex = -1 Then Err.Raise vbObjectError + 3, , "Riga separatrice Markdown non trovata."

    targetRange.text = ""
    Set wordTable = ActiveDocument.tables.Add(Range:=targetRange, numRows:=(rowCount - 1), NumColumns:=colCount)

    actualRow = 1
    For i = 0 To UBound(lines)
        If i <> sepLineIndex Then
            rowVals = SplitMarkdownRow(lines(i), colCount)
            For j = 1 To colCount
                Dim r As Range, cellText As String
                Set r = wordTable.cell(actualRow, j).Range
                ' r include il carattere di fine cella: lo escludo, poi scrivo con il writer inline
                r.End = r.End - 1
                r.text = ""
                r.Select
                cellText = rowVals(j - 1)
                ' Header (riga logica 1) in bold di default, ma B/I rispettano i marker
                InsertMarkdownInlineToSelection cellText, (actualRow = 1), False
            Next j
            actualRow = actualRow + 1
        End If
    Next i

    ApplyNiceTableFormatting wordTable

    Dim afterTbl As Range
    Set afterTbl = wordTable.Range
    afterTbl.Collapse wdCollapseEnd
    afterTbl.Select
End Sub


' === Rilevamento primo blocco tabella nel testo
Private Function ExtractFirstMarkdownTableBlock(ByVal src As String, _
                                               ByRef preText As String, _
                                               ByRef tableBlock As String, _
                                               ByRef postText As String) As Boolean
    Dim lines() As String, i As Long, j As Long, k As Long
    Dim hdr As String, sep As String, colCount As Long
    Dim startIdx As Long, sepIdx As Long, endIdx As Long
    
    lines = Split(src, vbLf)
    startIdx = -1: sepIdx = -1: endIdx = -1
    
    For i = LBound(lines) To UBound(lines) - 1
        hdr = Trim$(lines(i))
        If IsFenceLine(hdr) Or hdr = "" Then GoTo NextI
        If CountPipes(hdr) >= 2 Then
            For j = i + 1 To UBound(lines)
                sep = Trim$(lines(j))
                If sep <> "" And Not IsFenceLine(sep) Then
                    colCount = CountMarkdownColumns(hdr)
                    If colCount > 0 And (IsSeparatorLine(sep, colCount) Or IsSeparatorLike(sep)) Then
                        startIdx = i: sepIdx = j
                    End If
                    Exit For
                End If
            Next j
            If startIdx <> -1 Then Exit For
        End If
NextI:
    Next i
    
    If startIdx = -1 Then
        preText = src: tableBlock = "": postText = ""
        ExtractFirstMarkdownTableBlock = False
        Exit Function
    End If
    
    endIdx = sepIdx
    For k = sepIdx + 1 To UBound(lines)
        If Trim$(lines(k)) = "" Or IsFenceLine(Trim$(lines(k))) Then Exit For
        If InStr(1, lines(k), "|") = 0 Then Exit For
        endIdx = k
    Next k
    
    preText = JoinSubArray(lines, LBound(lines), startIdx - 1, vbLf)
    tableBlock = JoinSubArray(lines, startIdx, endIdx, vbLf)
    postText = ""
    If endIdx + 1 <= UBound(lines) Then
        postText = JoinSubArray(lines, endIdx + 1, UBound(lines), vbLf)
    End If
    
    ExtractFirstMarkdownTableBlock = True
End Function

' === Creazione tabella Word dalla stringa Markdown (con B/I inline)
Private Sub CreateAndInsertWordTableFromMarkdown(ByVal markdown As String, ByVal targetRange As Range)
    Dim rawLines() As String, lines() As String
    Dim i As Long, rowCount As Long, colCount As Long
    Dim sepLineIndex As Long, actualRow As Long
    Dim rowVals() As String
    Dim wordTable As Table
    
    rawLines = Split(NormalizeToLf(RemoveFences(markdown)), vbLf)
    lines = FilterNonEmpty(rawLines)
    
    rowCount = UBound(lines) - LBound(lines) + 1
    If rowCount < 2 Then Err.Raise vbObjectError + 7001, , "Tabella Markdown non valida."
    
    colCount = CountMarkdownColumns(lines(0))
    If colCount < 1 Then Err.Raise vbObjectError + 7002, , "Impossibile determinare il numero di colonne."
    
    sepLineIndex = FindSeparatorIndex(lines, colCount)
    If sepLineIndex = -1 Then Err.Raise vbObjectError + 7003, , "Riga separatrice Markdown non trovata."
    
    targetRange.text = ""
    Set wordTable = ActiveDocument.tables.Add(Range:=targetRange, numRows:=(rowCount - 1), NumColumns:=colCount)
    
    actualRow = 1
    For i = 0 To UBound(lines)
        If i <> sepLineIndex Then
            rowVals = SplitMarkdownRow(lines(i), colCount)
            Dim c As Long
            For c = 1 To colCount
                Dim r As Range, cellText As String
                Set r = wordTable.cell(actualRow, c).Range
                r.End = r.End - 1
                r.text = ""
                r.Select
                cellText = rowVals(c - 1)
                InsertMarkdownInlineToSelection cellText, (actualRow = 1), False
            Next c
            actualRow = actualRow + 1
        End If
    Next i
    
    ApplyNiceTableFormatting wordTable
    
    Dim afterTbl As Range
    Set afterTbl = wordTable.Range
    afterTbl.Collapse wdCollapseEnd
    afterTbl.Select
End Sub

' ========= Helpers Markdown / Word =========
Private Function NormalizeToLf(ByVal tx As String) As String
    tx = Replace(tx, vbCrLf, vbLf)
    tx = Replace(tx, vbCr, vbLf)
    NormalizeToLf = tx
End Function

Private Function RemoveFences(ByVal tx As String) As String
    Dim out As String, lines() As String, i As Long, ln As String
    lines = Split(NormalizeToLf(tx), vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(lines(i))
        If Not IsFenceLine(ln) Then out = out & lines(i) & vbLf
    Next i
    If Right$(out, 1) = vbLf Then out = Left$(out, Len(out) - 1)
    RemoveFences = out
End Function

Private Function FilterNonEmpty(arr() As String) As String()
    Dim tmp() As String, i As Long, v As String, k As Long
    ReDim tmp(0 To 0)
    For i = LBound(arr) To UBound(arr)
        v = Trim$(arr(i))
        If v <> "" Then
            If tmp(0) = "" Then
                tmp(0) = v
            Else
                ReDim Preserve tmp(0 To UBound(tmp) + 1)
                tmp(UBound(tmp)) = v
            End If
        End If
    Next i
    FilterNonEmpty = tmp
End Function

Private Function FindSeparatorIndex(lines() As String, ByVal colCount As Long) As Long
    Dim i As Long
    For i = 1 To UBound(lines)
        If IsSeparatorLine(lines(i), colCount) Or IsSeparatorLike(lines(i)) Then
            FindSeparatorIndex = i
            Exit Function
        End If
    Next i
    FindSeparatorIndex = -1
End Function

Private Function CountMarkdownColumns(ByVal rowText As String) As Long
    Dim parts() As String, i As Long, cnt As Long
    parts = Split(rowText, "|")
    For i = LBound(parts) To UBound(parts)
        If Trim$(parts(i)) <> "" Then cnt = cnt + 1
    Next i
    CountMarkdownColumns = cnt
End Function

Private Function IsSeparatorLine(ByVal rowText As String, ByVal colCount As Long) As Boolean
    Dim vals() As String, i As Long, tok As String, seen As Long
    vals = Split(rowText, "|")
    For i = LBound(vals) To UBound(vals)
        tok = Trim$(vals(i))
        If tok <> "" Then
            seen = seen + 1
            If Not IsSeparatorToken(tok) Then
                IsSeparatorLine = False
                Exit Function
            End If
        End If
    Next i
    IsSeparatorLine = (seen = colCount)
End Function

Private Function IsSeparatorLike(ByVal rowText As String) As Boolean
    IsSeparatorLike = (InStr(1, rowText, "---") > 0)
End Function

Private Function IsSeparatorToken(ByVal tok As String) As Boolean
    Dim s As String: s = tok
    s = Replace$(s, ":", "")
    s = Replace$(s, "-", "")
    If Len(Replace$(tok, ":", "")) - Len(Replace$(Replace$(tok, ":", ""), "-", "")) >= 3 Then
        Dim c As Long, ch As String
        For c = 1 To Len(tok)
            ch = Mid$(tok, c, 1)
            If ch <> "-" And ch <> ":" Then
                IsSeparatorToken = False
                Exit Function
            End If
        Next c
        IsSeparatorToken = True
    Else
        IsSeparatorToken = False
    End If
End Function

Private Function SplitMarkdownRow(ByVal rowText As String, ByVal colCount As Long) As String()
    Dim parts() As String, tmp() As String
    Dim i As Long, v As String, k As Long
    
    parts = Split(rowText, "|")
    ReDim tmp(0 To colCount - 1)
    k = 0
    For i = LBound(parts) To UBound(parts)
        v = Trim$(parts(i))
        If v <> "" Then
            If k <= UBound(tmp) Then
                tmp(k) = v
                k = k + 1
            End If
        End If
    Next i
    For i = k To UBound(tmp)
        tmp(i) = ""
    Next i
    SplitMarkdownRow = tmp
End Function

Private Sub ApplyNiceTableFormatting(ByVal tbl As Table)
    Dim i As Long, j As Long
    On Error Resume Next
    tbl.Style = "Tabella griglia"
    If Err.Number <> 0 Then
        Err.Clear: tbl.Style = "Table Grid"
    End If
    On Error GoTo 0
    
    With tbl
        .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Range.Font.Name = "Calibri"
        .Range.Font.Size = 11
        .Borders.OutsideLineStyle = wdLineStyleSingle
        .Borders.InsideLineStyle = wdLineStyleSingle
        
        For j = 1 To .Columns.Count
            With .cell(1, j).Range
                .Font.Bold = True
                .Shading.BackgroundPatternColor = RGB(240, 240, 240)
            End With
        Next j
        
        For i = 2 To .rows.Count
            If i Mod 2 = 0 Then
                For j = 1 To .Columns.Count
                    .cell(i, j).Range.Shading.BackgroundPatternColor = RGB(249, 249, 249)
                Next j
            End If
        Next i
        
        .AutoFitBehavior wdAutoFitContent
    End With
End Sub

' Inserisce testo con marker **bold** / *italic* e __bold__ / _italic_
' Supporta *** / ___ (bold+italic) e il backslash \ per escape dei marker.
' Scrive "a run": applica Bold/Italic solo al range appena inserito (niente toggle persistente).
Private Sub InsertMarkdownInlineToSelection(ByVal s As String, _
                                            Optional ByVal defaultBold As Boolean = False, _
                                            Optional ByVal defaultItalic As Boolean = False)
    Dim i As Long
    Dim buf As String
    Dim boldState As Boolean, italicState As Boolean

    s = NormalizeToLf(s)
    boldState = False: italicState = False
    i = 1

    Do While i <= Len(s)
        ' Escape: \* o \_ -> carattere letterale
        If Mid$(s, i, 1) = "\" Then
            If i < Len(s) Then
                buf = buf & Mid$(s, i + 1, 1)
                i = i + 2
            Else
                i = i + 1
            End If

        ' *** o ___ => toggle bold+italic
        ElseIf i <= Len(s) - 2 And (Mid$(s, i, 3) = "***" Or Mid$(s, i, 3) = "___") Then
            Call FlushRun(buf, defaultBold Or boldState, defaultItalic Or italicState)
            boldState = Not boldState
            italicState = Not italicState
            i = i + 3

        ' ** o __ => toggle bold
        ElseIf i <= Len(s) - 1 And (Mid$(s, i, 2) = "**" Or Mid$(s, i, 2) = "__") Then
            Call FlushRun(buf, defaultBold Or boldState, defaultItalic Or italicState)
            boldState = Not boldState
            i = i + 2

        ' * o _ => toggle italic
        ElseIf Mid$(s, i, 1) = "*" Or Mid$(s, i, 1) = "_" Then
            Call FlushRun(buf, defaultBold Or boldState, defaultItalic Or italicState)
            italicState = Not italicState
            i = i + 1

        ' A capo
        ElseIf Mid$(s, i, 1) = vbLf Then
            Call FlushRun(buf, defaultBold Or boldState, defaultItalic Or italicState)
            Selection.TypeParagraph
            i = i + 1

        ' Carattere normale
        Else
            buf = buf & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop

    ' Flush residuo
    Call FlushRun(buf, defaultBold Or boldState, defaultItalic Or italicState)
End Sub

' Scrive il buffer e applica Bold/Italic SOLO al range appena scritto
Private Sub FlushRun(ByRef buf As String, ByVal isBold As Boolean, ByVal isItalic As Boolean)
    Dim startPos As Long
    Dim r As Range
    If Len(buf) = 0 Then Exit Sub

    startPos = Selection.Range.Start
    Selection.TypeText buf
    Set r = ActiveDocument.Range(startPos, Selection.Range.Start)
    r.Font.Bold = isBold
    r.Font.Italic = isItalic

    buf = vbNullString
End Sub



' Utility per blocchi
Private Function JoinSubArray(arr() As String, ByVal a As Long, ByVal b As Long, ByVal sep As String) As String
    Dim i As Long, s As String
    If b < a Or a < LBound(arr) Or b > UBound(arr) Then Exit Function
    For i = a To b
        s = s & arr(i)
        If i < b Then s = s & sep
    Next i
    JoinSubArray = s
End Function

Private Function IsFenceLine(ByVal ln As String) As Boolean
    ln = LCase$(Trim$(ln))
    IsFenceLine = (ln = "```" Or ln = "```markdown")
End Function

Private Function CountPipes(ByVal s As String) As Long
    CountPipes = (Len(s) - Len(Replace$(s, "|", "")))
End Function

' Costruisce il testo da inviare all'API:
' - Selezione senza tabelle: normalizza + sanifica
' - Selezione con tabelle: testo (sanificato) + blocchi Markdown per ogni tabella
Private Function BuildMessageFromSelection(ByVal rng As Range) As String
    Dim out As String
    Dim curStart As Long
    Dim nextT As Table
    Dim txtRng As Range

    If rng.tables.Count = 0 Then
        out = SanitizeForApi(NormalizeParagraphs(rng.text))
        BuildMessageFromSelection = out
        Exit Function
    End If

    curStart = rng.Start
    Do
        Set nextT = FindNextTableInRange(rng, curStart)
        If nextT Is Nothing Then
            If curStart < rng.End Then
                Set txtRng = ActiveDocument.Range(curStart, rng.End)
                out = out & SanitizeForApi(NormalizeParagraphs(txtRng.text))
            End If
            Exit Do
        End If

        ' testo prima della tabella
        If nextT.Range.Start > curStart Then
            Set txtRng = ActiveDocument.Range(curStart, nextT.Range.Start)
            out = out & SanitizeForApi(NormalizeParagraphs(txtRng.text))
            If Len(out) > 0 And Right$(out, 1) <> vbLf Then out = out & vbLf
        End If

        ' tabella in Markdown (separata da righe vuote)
        If Len(out) > 0 And Right$(out, 1) <> vbLf Then out = out & vbLf
        out = out & TableToMarkdown(nextT) & vbLf
        curStart = nextT.Range.End
        If curStart < rng.End Then out = out & vbLf
    Loop

    BuildMessageFromSelection = Trim$(out)
End Function


' Trova la prossima tabella (in ordine di posizione) all'interno di rng, a partire da fromPos
Private Function FindNextTableInRange(ByVal rng As Range, ByVal fromPos As Long) As Table
    Dim t As Table
    Dim best As Table
    Dim bestStart As Long

    bestStart = 0
    For Each t In rng.tables
        If (fromPos >= t.Range.Start And fromPos < t.Range.End) Then
            Set FindNextTableInRange = t
            Exit Function
        ElseIf t.Range.Start >= fromPos Then
            If best Is Nothing Or t.Range.Start < bestStart Then
                Set best = t
                bestStart = t.Range.Start
            End If
        End If
    Next t

    If Not best Is Nothing Then Set FindNextTableInRange = best
End Function

' Converte una tabella Word in Markdown (GFM-like)
Private Function TableToMarkdown(ByVal tbl As Table) As String
    Dim r As Long, c As Long, cols As Long, rows As Long
    Dim line As String, sep As String, md As String

    rows = tbl.rows.Count
    cols = tbl.Columns.Count
    If cols = 0 Or rows = 0 Then Exit Function

    ' Header = prima riga
    line = ""
    For c = 1 To cols
        line = line & "| " & EscapePipes(Trim$(GetCellPlainText(tbl.cell(1, c)))) & " "
    Next c
    line = line & "|"
    md = line & vbLf

    ' Separatore con allineamento derivato dall'header
    sep = ""
    For c = 1 To cols
        Select Case tbl.cell(1, c).Range.ParagraphFormat.Alignment
            Case wdAlignParagraphCenter: sep = sep & "|:" & String$(3, "-") & ":"
            Case wdAlignParagraphRight:  sep = sep & "| " & String$(3, "-") & ":"
            Case Else:                   sep = sep & "| " & String$(3, "-") & " "
        End Select
        sep = sep & " "
    Next c
    sep = sep & "|"
    md = md & sep & vbLf

    ' Dati
    For r = 2 To rows
        line = ""
        For c = 1 To cols
            line = line & "| " & EscapePipes(Trim$(GetCellPlainText(tbl.cell(r, c)))) & " "
        Next c
        line = line & "|"
        md = md & line & vbLf
    Next r

    ' Sanifica (rimuove controlli residui)
    TableToMarkdown = SanitizeForApi(md)
End Function

' Estrae il testo "pulito" di una cella (senza i marcatori di fine cella/riga)
Private Function GetCellPlainText(ByVal cel As cell) As String
    Dim tx As String
    tx = cel.Range.text
    ' Rimuove i marcatori finali (CR + CellEnd Chr(7))
    If Len(tx) >= 2 Then
        tx = Replace(tx, Chr(13) & Chr(7), "")
    End If
    ' Gestione line break interni alla cella: CR/VT -> " / "
    tx = Replace(tx, vbCr, " / ")
    tx = Replace(tx, Chr(11), " / ")
    tx = Replace(tx, vbTab, " ")
    GetCellPlainText = tx
End Function

' Escapa i pipe che romperebbero la tabella Markdown
Private Function EscapePipes(ByVal s As String) As String
    EscapePipes = Replace(s, "|", "\|")
End Function

' Normalizza i paragrafi a \n
Private Function NormalizeParagraphs(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    s = Replace(s, Chr(11), vbLf) ' manual line break
    NormalizeParagraphs = s
End Function

' Rimuove i caratteri di controllo non stampabili (0..31 esclusi 9,10,13).
' Se replaceWith è non-vuoto, i caratteri scartati vengono sostituiti con tale stringa.
Public Function SanitizeControlChars(ByVal s As String, Optional ByVal replaceWith As String = "") As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim out As String

    If LenB(s) = 0 Then
        SanitizeControlChars = s
        Exit Function
    End If

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)

        If (code = 9 Or code = 10 Or code = 13 Or code >= 32) Then
            ' mantieni il carattere
            out = out & ch
        Else
            ' scarta (code 0..31 esclusi 9,10,13) oppure sostituisci
            If Len(replaceWith) > 0 Then
                out = out & replaceWith
            End If
        End If
    Next i

    SanitizeControlChars = out
End Function

' Pipeline di sanificazione "safe-for-API":
' - normalizza CR/LF in LF
' - rimuove/sostituisce controlli (0..31 tranne 9,10,13)
' - converte NBSP in spazio
' - rimuove zero-width e bidi invisibili
' - trim di spazi a fine riga
' - collassa righe vuote consecutive
Private Function SanitizeForApi(ByVal s As String) As String
    Dim t As String
    t = s

    ' Normalizza le nuove linee (idempotente)
    t = NormalizeParagraphs(t)

    ' Tab opzionale -> 4 spazi
    t = Replace(t, vbTab, "    ")

    ' NBSP -> spazio normale
    t = Replace(t, ChrW(160), " ")

    ' Rimuovi zero-width / bidi invisibili più comuni
    ' (ZWSP, ZWNJ, ZWJ, LRE, RLE, PDF, LRO, RLO, isolating format chars)
    t = RemoveCharsByCodepoints(t, Array( _
        8203, 8204, 8205, _
        8234, 8235, 8236, 8237, 8238, _
        8298, 8299, 8300, 8301, 8302, 8303 _
    ))

    ' Controlli non stampabili (eccetto 9,10,13)
    t = SanitizeControlChars(t, "")

    ' Trim spazi a fine riga
    t = TrimLineEnds(t)

    ' Collassa righe vuote multiple (max 2 consecutive)
    t = CollapseBlankLines(t, 2)

    SanitizeForApi = t
End Function


' Rimuove caratteri corrispondenti ai codepoint indicati
Private Function RemoveCharsByCodepoints(ByVal s As String, ByVal codes As Variant) As String
    Dim i As Long, ch As String, cp As Long, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        cp = AscW(ch)
        If Not InArrayLong(cp, codes) Then out = out & ch
    Next
    RemoveCharsByCodepoints = out
End Function

Private Function InArrayLong(ByVal v As Long, ByVal arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If v = CLng(arr(i)) Then InArrayLong = True: Exit Function
    Next
End Function

' Elimina spazi e tab a fine riga
Private Function TrimLineEnds(ByVal s As String) As String
    Dim lines() As String, i As Long
    Dim out As String
    lines = Split(s, vbLf)
    For i = LBound(lines) To UBound(lines)
        lines(i) = RTrim$(Replace(lines(i), vbTab, " "))
    Next
    out = Join(lines, vbLf)
    TrimLineEnds = out
End Function

' Mantiene al più "maxConsecutive" righe vuote di fila
Private Function CollapseBlankLines(ByVal s As String, Optional ByVal maxConsecutive As Long = 2) As String
    Dim lines() As String, i As Long, emptyRun As Long
    Dim out As String

    lines = Split(s, vbLf)
    emptyRun = 0
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(lines(i))) = 0 Then
            emptyRun = emptyRun + 1
            If emptyRun <= maxConsecutive Then
                out = out & vbLf
            End If
        Else
            emptyRun = 0
            If Len(out) > 0 Then out = out & lines(i) Else out = lines(i)
            If i < UBound(lines) Then out = out & vbLf
        End If
    Next

    ' Normalizza eventuale doppio vbLf iniziale/terminale
    Do While Left$(out, 1) = vbLf: out = Mid$(out, 2): Loop
    Do While Right$(out, 1) = vbLf And Right$(out, 2) = vbLf & vbLf: out = Left$(out, Len(out) - 1): Loop

    CollapseBlankLines = out
End Function

