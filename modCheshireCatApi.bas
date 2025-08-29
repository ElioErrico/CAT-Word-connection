Attribute VB_Name = "modCheshireCatApi"
' ==========================
' Modulo: modCheshireCatApi
' ==========================
Option Explicit

' =========== CONFIG ===========
Public Const DEFAULT_URL As String = "http://localhost:1865"
Public Const DEFAULT_USERNAME As String = "admin"
Public Const DEFAULT_PASSWORD As String = "admin"

' ========== HTTP / AUTH ==========
' Ritorna il token JWT (oppure stringa che inizia con "Errore")
Public Function GetJWToken() As String
    Dim httpRequest As Object
    Dim url As String
    Dim requestBody As String
    Dim responseText As String
    Dim tokenValue As String
    
    url = DEFAULT_URL & "/auth/token"
    requestBody = "{""username"": """ & EscapeJsonString(DEFAULT_USERNAME) & """, ""password"": """ & EscapeJsonString(DEFAULT_PASSWORD) & """}"
    Set httpRequest = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    On Error GoTo ErrorHandler
    With httpRequest
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/json"
        .send requestBody
    End With
    
    If httpRequest.Status <> 200 Then
        GetJWToken = "Errore HTTP: " & httpRequest.Status & " - " & httpRequest.StatusText
        Exit Function
    End If
    
    responseText = httpRequest.responseText
    tokenValue = ExtractJsonValue(responseText, "access_token")
    If tokenValue <> "" Then
        GetJWToken = tokenValue
    Else
        GetJWToken = "Token non trovato nella risposta"
    End If
    Exit Function
ErrorHandler:
    GetJWToken = "Errore: " & Err.Description
End Function

' Invia testo al CheshireCat e ritorna la stringa di contenuto
Public Function CheshireCat_Chat(messageText As String) As String
    Dim httpRequest As Object
    Dim url As String
    Dim requestBody As String
    Dim responseText As String
    Dim contentValue As String
    Dim jwtToken As String
    
    jwtToken = GetJWToken()
    If Left$(jwtToken, 5) = "Error" Or Left$(jwtToken, 6) = "Errore" Then
        CheshireCat_Chat = jwtToken
        Exit Function
    End If
    
    url = DEFAULT_URL & "/message"
    requestBody = "{""text"": """ & EscapeJsonString(messageText) & """}"
    Set httpRequest = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    On Error GoTo ErrorHandler
    With httpRequest
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & jwtToken
        .send requestBody
    End With
    
    If httpRequest.Status <> 200 Then
        CheshireCat_Chat = "Errore HTTP: " & httpRequest.Status & " - " & httpRequest.StatusText
        Exit Function
    End If
    
    responseText = httpRequest.responseText
    contentValue = ExtractJsonValue(responseText, "content")
    
    If contentValue <> "" Then
        CheshireCat_Chat = contentValue
    Else
        CheshireCat_Chat = "Campo 'content' non trovato nella risposta"
    End If
    
    Exit Function
ErrorHandler:
    CheshireCat_Chat = "Errore: " & Err.Description
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
    
    ClearChatHistory = (httpRequest.Status = 200)
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


