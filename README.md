# Integrazione CheshireCat per Microsoft Word
## Descrizione

Questo progetto integra le funzionalità di CheshireCat (un assistente AI) all'interno di Microsoft Word attraverso macro VBA, aggiungendo un menu contestuale e una scheda personalizzata nella barra multifunzione.
Installazione

## Importazione manuale dei moduli (consigliato)

    Scarica i file .bas dalla repository

    Apri Microsoft Word

    Premi ALT+F11 per aprire l'editor VBA

    Nel pannello "Project - VBAProject", clicca destro su "Normal"

    Seleziona "Importa file..."

    Importa tutti i file .bas dalla repository

    Salva il progetto e chiudi l'editor VBA


## Configurazione

## OPZIONALE Configurazione della barra multifunzione

    Vai su File → Opzioni → Personalizza barra multifunzione

    Crea una nuova scheda "CheshireCat"

    Aggiungi un nuovo gruppo alla scheda

    Dal menu "Scegli comandi da:" seleziona "Macro"

    Aggiungi le macro "InviaTestoAChat" e "CancellaCronologiaChat" al gruppo

    Personalizza le icone e i nomi dei pulsanti se desiderato

    Clicca "OK" per salvare

## Configurazione del menu contestuale

Il menu contestuale viene aggiunto automaticamente all'avvio di Word. Se non compare:

    Apri l'editor VBA (ALT+F11)

    Esegui manualmente la macro "AddCheshireCatContextMenu"


## Utilizzo

    Seleziona del testo in un documento Word

    Usa una delle seguenti opzioni:

        Clic col tasto destro → CheshireCat → Invia testo a CheshireCat

        Scheda CheshireCat nella barra multifunzione → Invia testo

    La risposta verrà inserita nel documento

## Personalizzazione

Modifica queste costanti nel codice per personalizzare la connessione:
vba

Private Const DEFAULT_URL As String = "http://192.168.71.63:1865"
Private Const DEFAULT_USERNAME As String = "admin"
Private Const DEFAULT_PASSWORD As String = "admin_psw"

## Requisiti di sistema

    Microsoft Word 2016 o versioni successive

    Abilitazione delle macro (File → Opzioni → Centro protezione → Impostazioni macro → "Abilita tutte le macro")


