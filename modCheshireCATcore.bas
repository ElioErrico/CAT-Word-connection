Attribute VB_Name = "modCheshireCATcore"
' ==========================
' Modulo: modCheshireCATcore
' ==========================
Option Explicit

' ========= INVIO TESTO E INSERIMENTO RISPOSTA =========
Public Sub InviaTestoAChat()
    Dim selectedText As String
    Dim response As String
    Dim rGray As Range
    Dim insertAt As Range

    If Selection Is Nothing Or Selection.Type = wdSelectionIP Then
        MsgBox "Nessun testo selezionato", vbExclamation
        Exit Sub
    End If

    ' 1) Costruisci payload (testo + tabelle -> markdown)
    selectedText = BuildMessageFromSelection(Selection.Range)
    If Len(Trim$(selectedText)) = 0 Then
        MsgBox "La selezione è vuota dopo la normalizzazione.", vbExclamation
        Exit Sub
    End If

    ' 2) Grigia SOLO i caratteri selezionati (escludi il ¶ finale)
    Set rGray = Selection.Range.Duplicate
    modMarkdownHelper.GrayOutRange rGray

    ' 3) Punto di inserimento DOPO il range grigiato, con formati puliti
    Set insertAt = rGray.Duplicate
    insertAt.Collapse wdCollapseEnd
    insertAt.InsertParagraphAfter
    insertAt.Collapse wdCollapseEnd
    With insertAt
        .ParagraphFormat.Reset
        .Font.Reset
        .HighlightColorIndex = wdNoHighlight
        .Shading.Texture = wdTextureNone
        .Shading.ForegroundPatternColor = wdColorAutomatic
        .Shading.BackgroundPatternColor = wdColorAutomatic
    End With

    ' Sposta la Selection sul punto target
    Selection.SetRange insertAt.Start, insertAt.End

    ' 4) Chiama API
    response = modCheshireCatApi.CheshireCat_Chat(selectedText)
    If Len(response) = 0 Or Left$(response, 6) = "Errore" Then
        If Len(response) = 0 Then
            MsgBox "Risposta vuota dall'API.", vbExclamation
        Else
            MsgBox response, vbExclamation
        End If
        Exit Sub
    End If

    ' 5) Inserisci testo + tabelle markdown
    modMarkdownHelper.InsertAIResponseWithMarkdownTables response
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

