Attribute VB_Name = "MenuContestuale"
Option Explicit

Sub AddCheshireCatContextMenu()
    Dim myMenu As CommandBarPopup
    Dim mySubMenu As CommandBarButton
    
    ' Rimuovi il menu esistente se presente
    On Error Resume Next
    Application.CommandBars("Text").Controls("CheshireCat").Delete
    On Error GoTo 0
    
    ' Aggiungi il menu principale
    Set myMenu = Application.CommandBars("Text").Controls.Add( _
        Type:=msoControlPopup, _
        Temporary:=True)
    myMenu.Caption = "CheshireCat"
    
    ' Aggiungi la voce per inviare il testo
    Set mySubMenu = myMenu.Controls.Add(Type:=msoControlButton)
    mySubMenu.Caption = "Invia testo a CheshireCat"
    mySubMenu.OnAction = "InviaTestoAChat"
    mySubMenu.FaceId = 59
    
    ' Aggiungi la voce per cancellare la cronologia
    Set mySubMenu = myMenu.Controls.Add(Type:=msoControlButton)
    mySubMenu.Caption = "Cancella cronologia chat"
    mySubMenu.OnAction = "CancellaCronologiaChat"
    mySubMenu.FaceId = 100
    
    ' Aggiungi la voce per convertire tabella markdown
    Set mySubMenu = myMenu.Controls.Add(Type:=msoControlButton)
    mySubMenu.Caption = "Converti tabella markdown"
    mySubMenu.OnAction = "ConvertiTabellaMarkdown"
    mySubMenu.FaceId = 16
End Sub

Sub RemoveCheshireCatContextMenu()
    On Error Resume Next
    Application.CommandBars("Text").Controls("CheshireCat").Delete
    On Error GoTo 0
End Sub

' Esegui all'avvio di Word
Sub AutoExec()
    AddCheshireCatContextMenu
End Sub

' Esegui alla chiusura (opzionale)
Sub AutoExit()
    RemoveCheshireCatContextMenu
End Sub

