Attribute VB_Name = "modMarkdownTables"
'==========================
' Modulo: modMarkdownTables
'==========================
Option Explicit

' Entrata principale richiamata dal menu:
' mySubMenu.OnAction = "ConvertiTabellaMarkdown"
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

' -- Estrae il testo selezionato e lo normalizza (fine riga, rimozione ```).
Private Function GetSelectedMarkdownTableText(ByVal rng As Range) As String
    Dim tx As String, lines() As String, i As Long, out As String
    
    tx = rng.text
    ' Normalizza i fine riga a vbLf
    tx = Replace(tx, vbCrLf, vbLf)
    tx = Replace(tx, vbCr, vbLf)
    
    ' Rimuovi eventuali fence ``` (``` o ```markdown)
    lines = Split(tx, vbLf)
    For i = LBound(lines) To UBound(lines)
        Dim ln As String
        ln = Trim$(lines(i))
        If ln <> "```" And LCase$(ln) <> "```markdown" And ln <> "" Then
            out = out & ln & vbLf
        End If
    Next i
    
    ' Rimuovi ultimo vbLf
    If Right$(out, 1) = vbLf Then out = Left$(out, Len(out) - 1)
    GetSelectedMarkdownTableText = out
End Function

' -- Converte markdown in tabella Word, sostituendo la selezione.
Private Sub ConvertMarkdownToWord(ByVal markdown As String, ByVal targetRange As Range)
    Dim rawLines() As String, lines() As String
    Dim i As Long, j As Long, rowCount As Long, colCount As Long
    Dim wordTable As Table
    Dim sepLineIndex As Long
    Dim actualRow As Long
    Dim rowVals() As String
    
    ' Split base e rimozione righe vuote
    rawLines = Split(markdown, vbLf)
    ReDim lines(0 To 0)
    For i = LBound(rawLines) To UBound(rawLines)
        Dim t As String
        t = Trim$(rawLines(i))
        If t <> "" Then
            If lines(0) = "" Then
                lines(0) = t
            Else
                ReDim Preserve lines(0 To UBound(lines) + 1)
                lines(UBound(lines)) = t
            End If
        End If
    Next i
    
    rowCount = UBound(lines) - LBound(lines) + 1
    If rowCount < 2 Then Err.Raise vbObjectError + 1, , "Tabella Markdown non valida (righe insufficienti)."
    
    ' Determina # colonne dalla prima riga (header)
    colCount = CountMarkdownColumns(lines(0))
    If colCount < 1 Then Err.Raise vbObjectError + 2, , "Impossibile determinare il numero di colonne."
    
    ' Individua la riga separatrice (di solito è la seconda).
    ' Accettiamo la classica sintassi: |---|:---:|---:| ecc.
    sepLineIndex = -1
    For i = 1 To UBound(lines) ' da seconda riga in poi
        If IsSeparatorLine(lines(i), colCount) Then
            sepLineIndex = i
            Exit For
        End If
    Next i
    If sepLineIndex = -1 Then
        ' fallback: consideriamo la seconda riga come separatrice se plausibile
        If rowCount >= 2 And IsSeparatorLike(lines(1)) Then
            sepLineIndex = 1
        Else
            Err.Raise vbObjectError + 3, , "Riga separatrice Markdown non trovata."
        End If
    End If
    
    ' Crea tabella su targetRange sostituendo la selezione
    targetRange.text = ""      ' svuota e usa quel range per l'inserimento
    Set wordTable = ActiveDocument.tables.Add(Range:=targetRange, numRows:=(rowCount - 1), NumColumns:=colCount)
    
    ' Popola: tutte le righe tranne la separatrice
    actualRow = 1
    For i = 0 To UBound(lines)
        If i <> sepLineIndex Then
            rowVals = SplitMarkdownRow(lines(i), colCount)
            For j = 1 To colCount
                wordTable.cell(actualRow, j).Range.text = rowVals(j - 1)
            Next j
            actualRow = actualRow + 1
        End If
    Next i
    
    ' Formattazione (stile IT/EN, zebra, header bold, bordi, autofit)
    ApplyNiceTableFormatting wordTable
End Sub

' -- Conta colonne dalla riga (ignora pipe iniziale/finale vuote)
Private Function CountMarkdownColumns(ByVal rowText As String) As Long
    Dim parts() As String, i As Long, cnt As Long
    parts = Split(rowText, "|")
    For i = LBound(parts) To UBound(parts)
        If Trim$(parts(i)) <> "" Then cnt = cnt + 1
    Next i
    CountMarkdownColumns = cnt
End Function

' -- Verifica che una riga sia una riga separatrice valida per colCount colonne
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
    ' Controllo blando: contiene almeno tre '-' consecutivi
    IsSeparatorLike = (InStr(1, rowText, "---") > 0)
End Function

' -- Token separatore valido: ---  :---  ---:  :---:
Private Function IsSeparatorToken(ByVal tok As String) As Boolean
    Dim s As String: s = tok
    s = Replace$(s, ":", "")
    s = Replace$(s, "-", "")
    ' valido se composto da soli '-' e/o ':', con almeno 3 '-'
    If Len(Replace$(tok, ":", "")) - Len(Replace$(Replace$(tok, ":", ""), "-", "")) >= 3 Then
        Dim c As Long
        For c = 1 To Len(tok)
            Dim ch As String
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

' -- Split di una riga markdown in colCount celle (trim, rimuove pipe estreme)
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
    
    ' Se meno campi del previsto, riempi vuoti
    For i = k To UBound(tmp)
        tmp(i) = ""
    Next i
    
    SplitMarkdownRow = tmp
End Function

' -- Applica stile, header bold con sfondo, zebra, bordi, autofit
Private Sub ApplyNiceTableFormatting(ByVal tbl As Table)
    Dim i As Long, j As Long
    On Error Resume Next
    tbl.Style = "Tabella griglia"     ' IT
    If Err.Number <> 0 Then
        Err.Clear
        tbl.Style = "Table Grid"      ' EN
    End If
    On Error GoTo 0
    
    With tbl
        .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Range.Font.Name = "Calibri"
        .Range.Font.Size = 11
        
        .Borders.OutsideLineStyle = wdLineStyleSingle
        .Borders.InsideLineStyle = wdLineStyleSingle
        
        ' Intestazione (prima riga logica)
        For j = 1 To .Columns.Count
            With .cell(1, j).Range
                .Font.Bold = True
                .Shading.BackgroundPatternColor = RGB(240, 240, 240)
            End With
        Next j
        
        ' Zebra dalle righe dati
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

