Attribute VB_Name = "modCheshireCatApi"
' ==========================
' Modulo: modCheshireCatApi
' ==========================
Option Explicit

' =========== CONFIG ===========
Public Const DEFAULT_URL As String = "http://192.168.71.63:1865"
Public Const DEFAULT_USERNAME As String = "admin"
Public Const DEFAULT_PASSWORD As String = "admin_psw"

' ===== TIMEOUTS (ms) =====
Private Const TO_RESOLVE  As Long = 5000      ' DNS
Private Const TO_CONNECT  As Long = 15000     ' TCP/TLS handshake
Private Const TO_SEND     As Long = 60000     ' invio request
Private Const TO_RECEIVE  As Long = 300000    ' attesa risposta (5 min)
Private Const MAX_RETRY   As Long = 2         ' n. tentativi extra su timeout

' ====== Dichiarazioni e utility di attesa ======
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Sub WaitMs(ByVal ms As Long)
    ' Attesa responsive per l'UI di Word (slice da 50 ms)
    Dim slice As Long
    slice = IIf(ms > 50, 50, ms)
    Do While ms > 0
        DoEvents
        Sleep slice
        ms = ms - slice
    Loop
End Sub

' ========== HTTP / AUTH ==========
' Ritorna il token JWT (oppure stringa che inizia con "Errore")
Public Function GetJWToken() As String
    Dim url As String, requestBody As String, responseText As String
    Dim tokenValue As String
    
    url = DEFAULT_URL & "/auth/token"
    requestBody = "{""username"": """ & EscapeJsonString(DEFAULT_USERNAME) & """, ""password"": """ & EscapeJsonString(DEFAULT_PASSWORD) & """}"
    
    responseText = HttpPostJson(url, requestBody)
    If Left$(responseText, 6) = "Errore" Then
        GetJWToken = responseText
        Exit Function
    End If
    
    tokenValue = ExtractJsonValue(responseText, "access_token")
    If tokenValue <> "" Then
        GetJWToken = tokenValue
    Else
        GetJWToken = "Token non trovato nella risposta"
    End If
End Function


' Invia testo al CheshireCat e ritorna la stringa di contenuto
Public Function CheshireCat_Chat(messageText As String) As String
    Dim jwtToken As String, url As String, requestBody As String
    Dim responseText As String, contentValue As String
    
    jwtToken = GetJWToken()
    If Left$(jwtToken, 5) = "Error" Or Left$(jwtToken, 6) = "Errore" Then
        CheshireCat_Chat = jwtToken
        Exit Function
    End If
    
    url = DEFAULT_URL & "/message"
    requestBody = "{""text"": """ & EscapeJsonString(messageText) & """}"
    
    responseText = HttpPostJson(url, requestBody, jwtToken)
    If Left$(responseText, 6) = "Errore" Then
        CheshireCat_Chat = responseText
        Exit Function
    End If
    
    contentValue = ExtractJsonValue(responseText, "content")
    If contentValue <> "" Then
        CheshireCat_Chat = contentValue
    Else
        CheshireCat_Chat = "Campo 'content' non trovato nella risposta"
    End If
End Function


' Classificazione semplice: ritorna solo la classe
Public Function CheshireCat_Classify(sentence As String, labels As Variant) As String
    Dim labels_list As String
    Dim i As Long
    Dim prompt As String
    
    If TypeName(labels) = "String" Then
        labels_list = "- " & Replace(labels, ",", vbNewLine & "- ")
    ElseIf IsArray(labels) Then
        For i = LBound(labels) To UBound(labels)
            labels_list = labels_list & "- " & labels(i) & vbNewLine
        Next i
    End If
    
    prompt = "Classify this sentence:" & vbNewLine & _
             """" & sentence & """" & vbNewLine & vbNewLine & _
             "Allowed classes are:" & vbNewLine & _
             labels_list & vbNewLine & vbNewLine & _
             "Just output the class, nothing else."
    
    CheshireCat_Classify = CheshireCat_Chat(prompt)
End Function

' Cancella la cronologia della conversazione (ritorna True se OK)
Public Function ClearChatHistory(ByVal jwtToken As String) As Boolean
    Dim httpRequest As Object
    Dim url As String
    
    url = DEFAULT_URL & "/memory/conversation_history"
    Set httpRequest = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    On Error GoTo ErrorHandler
    With httpRequest
        .Open "DELETE", url, False
        .setRequestHeader "Authorization", "Bearer " & jwtToken
        .send
    End With
    
    ClearChatHistory = (httpRequest.status = 200)
    Exit Function
ErrorHandler:
    ClearChatHistory = False
End Function

' ====== Helpers JSON ======
Private Function ExtractJsonValue(jsonText As String, key As String) As String
    Dim startPos As Long, valueStart As Long, endPos As Long
    startPos = InStr(1, jsonText, """" & key & """:""")
    If startPos = 0 Then Exit Function
    valueStart = startPos + Len("""" & key & """:""")
    endPos = InStr(valueStart, jsonText, """")
    If endPos > valueStart Then ExtractJsonValue = Mid$(jsonText, valueStart, endPos - valueStart)
End Function

Private Function EscapeJsonString(text As String) As String
    Dim s As String
    s = Replace(text, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    EscapeJsonString = s
End Function

Private Function HttpPostJson(ByVal url As String, ByVal body As String, _
                              Optional ByVal bearer As String = "") As String
    Dim http As Object, attempt As Long
    Dim status As Long, resp As String
    Dim errNum As Long, errDesc As String
    
    For attempt = 0 To MAX_RETRY
        Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        On Error GoTo SendErr
        
        ' NOTA: setTimeouts va chiamato prima di .send
        http.Open "POST", url, False
        http.setTimeouts TO_RESOLVE, TO_CONNECT, TO_SEND, TO_RECEIVE
        http.setRequestHeader "Content-Type", "application/json"
        If Len(bearer) > 0 Then http.setRequestHeader "Authorization", "Bearer " & bearer
        http.setRequestHeader "Connection", "Keep-Alive"
        
        http.send body
        
        status = http.status
        resp = http.responseText
        
        ' Successo
        If status >= 200 And status < 300 Then
            HttpPostJson = resp
            Exit Function
        End If
        
        ' Alcuni ambienti possono restituire 0 su errori di rete senza eccezione
        If status = 0 Then
            DoEvents
            WaitMs 1000 * (1 + attempt)  ' backoff lineare: 1s, 2s, 3s...
            GoTo ResumeNextAttempt
        End If
        
        ' Retry su errori transitori server/network
        If (status = 408 Or status = 429 Or status = 500 Or status = 502 Or status = 503 Or status = 504) Then
            DoEvents
            WaitMs 1000 * (1 + attempt)  ' backoff lineare
        Else
            HttpPostJson = "Errore HTTP: " & status & " - " & http.StatusText
            Exit Function
        End If
        
ResumeNextAttempt:
        On Error GoTo 0
        Set http = Nothing
    Next attempt
    
    HttpPostJson = "Errore: timeout/scadenza dopo " & (MAX_RETRY + 1) & " tentativi"
    Exit Function

SendErr:
    errNum = Err.Number: errDesc = Err.Description
    On Error GoTo 0
    If attempt < MAX_RETRY Then
        DoEvents
        WaitMs 1000 * (1 + attempt)      ' backoff anche su eccezioni COM (es. timeout)
        Resume ResumeNextAttempt
    Else
        HttpPostJson = "Errore: " & errNum & " - " & errDesc
    End If
End Function


Private Function HttpDelete(ByVal url As String, ByVal bearer As String) As Long
    Dim http As Object, attempt As Long
    For attempt = 0 To MAX_RETRY
        Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        On Error Resume Next
        http.Open "DELETE", url, False
        http.setTimeouts TO_RESOLVE, TO_CONNECT, TO_SEND, TO_RECEIVE
        http.setRequestHeader "Authorization", "Bearer " & bearer
        http.send
        If Err.Number = 0 And http.status > 0 Then
            HttpDelete = http.status
            Exit Function
        End If
        On Error GoTo 0
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 1 + attempt)
    Next attempt
    HttpDelete = 0
End Function




