Attribute VB_Name = "Modulo14"
Sub m14_SaveAsVersione()
    
    Dim nomeFile As String
    Dim cartella As String
    Dim estensione As String
    Dim descrizione As String
    Dim nomeBase As String
    Dim nuovoNumero As String
    Dim nuovoNome As String
    Dim risposta As String
    
    ' --- Percorso e nome del file corrente ---
    nomeFile = ThisWorkbook.Name
    cartella = ThisWorkbook.path & "\"
    
    ' Estrai l'estensione
    Dim posPunto As Long
    posPunto = InStrRev(nomeFile, ".")
    estensione = Mid(nomeFile, posPunto)               ' es: .xlsm
    nomeFile = Left(nomeFile, posPunto - 1)            ' es: budget_02_inserito costi
    
    ' --- Cerca se il nome contiene già il pattern _NN o _NN_descrizione ---
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "^(.+?)_(\d+)(_.*)?$"
    regEx.IgnoreCase = True
    
    Dim nuovoNum As Integer
    
    If regEx.Test(nomeFile) Then
        Dim matches As Object
        Set matches = regEx.Execute(nomeFile)
        nomeBase = matches(0).SubMatches(0)            ' parte prima del numero
        nuovoNum = CInt(matches(0).SubMatches(1)) + 1  ' incrementa il numero
    Else
        ' Il file non ha ancora il pattern: lo usiamo come base, partiamo da _00
        nomeBase = nomeFile
        nuovoNum = 0
    End If
    
    ' Formatta il numero con zero iniziale
    nuovoNumero = Format(nuovoNum, "00")
    
    ' --- Chiedi la descrizione (lascia vuoto per omettere) ---
    risposta = InputBox( _
        "Salvataggio versione progressiva" & vbCrLf & vbCrLf & _
        "Base: " & nomeBase & "_" & nuovoNumero & estensione & vbCrLf & vbCrLf & _
        "Inserisci una breve descrizione (opzionale):", _
        "Save As Versione", "")
    
    ' Se l'utente preme Annulla, esci
    If risposta = "" And StrPtr(risposta) = 0 Then
        MsgBox "Operazione annullata.", vbInformation
        Exit Sub
    End If
    
    ' Costruisci il nuovo nome
    If Trim(risposta) <> "" Then
        nuovoNome = nomeBase & "_" & nuovoNumero & "_" & Trim(risposta) & estensione
    Else
        nuovoNome = nomeBase & "_" & nuovoNumero & estensione
    End If
    
    ' --- Controlla se il file esiste già ---
    If Dir(cartella & nuovoNome) <> "" Then
        Dim conf As VbMsgBoxResult
        conf = MsgBox("Il file """ & nuovoNome & """ esiste già." & vbCrLf & "Vuoi sovrascriverlo?", _
                      vbYesNo + vbExclamation, "File esistente")
        If conf <> vbYes Then
            MsgBox "Operazione annullata.", vbInformation
            Exit Sub
        End If
    End If
    
    ' --- Salva ---
    Application.DisplayAlerts = False
    ThisWorkbook.SaveAs Filename:=cartella & nuovoNome, _
                        FileFormat:=ThisWorkbook.FileFormat
    Application.DisplayAlerts = True
    
    MsgBox "File salvato come:" & vbCrLf & nuovoNome, vbInformation, "Salvataggio completato"
    
End Sub

