Attribute VB_Name = "Modulo7"
' ===============================================================
' MACRO: EsportaFogliSelezionati_multipli_plus
' ---------------------------------------------------------------
' SCOPO:
'   Esporta in un unico file .xlsx tutti i fogli selezionati
'   con Ctrl+Click nel file di origine. Il nome del file viene
'   chiesto all'utente tramite InputBox; se lasciato vuoto
'   viene usato il nome di default "report_esport.xlsx".
'
' COMPORTAMENTO FORMULE:
'   - Le formule che collegano fogli TRA QUELLI ESPORTATI
'     vengono mantenute come formule (collegamento interno OK)
'   - Le formule che referenziano fogli NON esportati
'     vengono convertite in valori (file autonomo e indipendente)
'   - I link a file Excel esterni vengono anch'essi convertiti
'     in valori tramite BreakLink
'
' OUTPUT:
'   File .xlsx salvato nella stessa cartella del file di origine.
'   Se esiste gia', viene sovrascritto.
'
' UTILIZZO:
'   1) Selezionare i fogli da esportare con Ctrl+Click sui tab
'   2) Eseguire la macro
'   3) Inserire il nome del file nell'InputBox (o lasciare vuoto)
'   4) Al termine appare un messaggio di conferma con il riepilogo
'
' NOTE TECNICHE:
'   Quando Excel copia un foglio in un nuovo workbook, riscrive
'   automaticamente le formule aggiungendo il riferimento al file
'   di origine, es: =FoglioA!F44 diventa ='[File.xlsm]FoglioA'!F44
'   La funzione RipulisciFormula gestisce questo comportamento,
'   distinguendo i fogli esportati da quelli non esportati.
' ===============================================================
Sub m7_EsportaFogliSelezionati_multipli_plus()
    Dim wb As Workbook
    Dim wbExport As Workbook
    Dim ws As Worksheet
    Dim filePath As String
    Dim linkSource As Variant
    Dim selectedSheets() As String
    Dim i As Integer
    Dim sheetCount As Integer
    Dim cell As Range
    Dim formula As String
    Dim cleanFormula As String
    Dim wbName As String
    Dim sheetList As String
    Dim customName As String

    ' Nasconde aggiornamenti schermo e alert durante l'esecuzione
    ' per evitare che l'utente veda file temporanei aprirsi/chiudersi
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wb = ThisWorkbook

    ' Conta i fogli selezionati con Ctrl+Click sui tab
    sheetCount = wb.Windows(1).selectedSheets.count

    ' Controllo di sicurezza: almeno un foglio deve essere selezionato
    If sheetCount < 1 Then
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        MsgBox "Nessun foglio selezionato.", vbExclamation
        Exit Sub
    End If

    ' Memorizza i nomi dei fogli selezionati in un array
    ' (necessario perche' dopo la copia il contesto cambia)
    ReDim selectedSheets(1 To sheetCount)
    For i = 1 To sheetCount
        selectedSheets(i) = wb.Windows(1).selectedSheets(i).Name
    Next i

    ' Riabilita temporaneamente lo schermo per mostrare l'InputBox
    Application.ScreenUpdating = True

    ' Chiede all'utente il nome del file di esportazione
    customName = InputBox("Inserisci il nome del file di esportazione:" & vbCrLf & _
                          "(lascia vuoto per usare 'report_esport')", _
                          "Nome file export", "report_esport")

    ' Se l'utente clicca Annulla ? esce senza fare nulla
    If StrPtr(customName) = 0 Then
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Exit Sub
    End If

    ' Se lascia vuoto ? nome di default
    If Trim(customName) = "" Then customName = "report_esport"

    ' Rimuove eventuale .xlsx digitato dall'utente per evitare doppio suffisso
    If LCase(Right(Trim(customName), 5)) = ".xlsx" Then
        customName = Left(Trim(customName), Len(Trim(customName)) - 5)
    End If

    ' Riapplica la soppressione dello schermo dopo l'InputBox
    Application.ScreenUpdating = False

    ' Percorso di destinazione: stessa cartella del file di origine
    filePath = wb.path & "\" & Trim(customName) & ".xlsx"

    ' Nome del file di origine usato dalla funzione RipulisciFormula
    ' per riconoscere i riferimenti inseriti automaticamente da Excel
    wbName = wb.Name

    ' Crea un nuovo workbook vuoto e copia dentro tutti i fogli selezionati
    Set wbExport = Workbooks.Add
    For i = 1 To sheetCount
        Set ws = wb.Sheets(selectedSheets(i))
        ws.Copy after:=wbExport.Sheets(wbExport.Sheets.count)
    Next i

    ' Elimina il foglio vuoto (es. "Foglio1") creato automaticamente
    ' da Workbooks.Add, finche' non rimangono solo i fogli copiati
    Do While wbExport.Sheets.count > sheetCount
        wbExport.Sheets(1).Delete
    Loop

    ' Scansiona tutte le celle con formule in tutti i fogli esportati
    ' per gestire i riferimenti ad altri fogli
    For i = 1 To wbExport.Sheets.count
        For Each cell In wbExport.Sheets(i).UsedRange
            If cell.HasFormula Then
                formula = cell.formula
                ' Controlla se la formula contiene riferimenti al file di origine
                ' (Excel li aggiunge automaticamente durante la copia)
                ' Esempio: =FoglioA!F44 ? ='[MioFile.xlsm]FoglioA'!F44
                If InStr(formula, "[" & wbName & "]") > 0 Then
                    ' Passa la formula alla funzione RipulisciFormula:
                    '   - Se il foglio referenziato E' tra gli esportati
                    '     ? restituisce la formula pulita (senza nome file)
                    '   - Se il foglio referenziato NON e' tra gli esportati
                    '     ? restituisce "" come segnale di conversione in valore
                    cleanFormula = RipulisciFormula(formula, wbName, selectedSheets, sheetCount)
                    If cleanFormula = "" Then
                        ' Foglio non esportato: sostituisce la formula con il valore
                        ' cosi' il file esportato e' completamente indipendente
                        cell.Value = cell.Value
                    Else
                        ' Foglio esportato: ripristina la formula pulita
                        ' che collega correttamente i fogli nel nuovo file
                        cell.formula = cleanFormula
                    End If
                End If
            End If
        Next cell
    Next i

    ' Rompe eventuali link residui verso file Excel esterni
    ' (es. formule che puntano ad altri .xlsx non coinvolti nell'export)
    On Error Resume Next
    For Each linkSource In wbExport.LinkSources(xlExcelLinks)
        wbExport.BreakLink Name:=linkSource, Type:=xlExcelLinks
    Next linkSource
    On Error GoTo 0

    ' Pulisce eventuali dati residui negli appunti
    Application.CutCopyMode = False

    ' Salva il nuovo file in formato .xlsx e lo chiude
    wbExport.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbook
    wbExport.Close

    ' Riabilita aggiornamenti schermo e alert prima del messaggio finale
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' Costruisce la lista dei fogli esportati per il messaggio di conferma
    sheetList = ""
    For i = 1 To sheetCount
        sheetList = sheetList & "- " & selectedSheets(i) & vbCrLf
    Next i

    MsgBox "Esportati " & sheetCount & " fogli in:" & vbCrLf & filePath & vbCrLf & vbCrLf & _
           "Fogli esportati:" & vbCrLf & sheetList & vbCrLf & _
           "- Formule interne tra fogli esportati: OK" & vbCrLf & _
           "- Riferimenti a fogli non esportati ? valori" & vbCrLf & _
           "- Link a file esterni ? valori", vbInformation
End Sub

' ===============================================================
' FUNZIONE: RipulisciFormula
' ---------------------------------------------------------------
' SCOPO:
'   Analizza una formula che contiene riferimenti al file di
'   origine (aggiunti automaticamente da Excel durante la copia)
'   e decide cosa farne in base ai fogli esportati.
'
' PARAMETRI:
'   formula        - La formula della cella da analizzare
'   wbName         - Il nome del file di origine (es. "File.xlsm")
'   selectedSheets - Array con i nomi dei fogli esportati
'   sheetCount     - Numero di fogli esportati
'
' RESTITUISCE:
'   - La formula pulita (senza "[NomeFile]") se TUTTI i fogli
'     referenziati sono tra quelli esportati
'   - "" (stringa vuota) se almeno un foglio referenziato
'     NON e' tra quelli esportati ? segnale per convertire in valore
' ===============================================================
Function RipulisciFormula(formula As String, wbName As String, _
                           selectedSheets() As String, sheetCount As Integer) As String
    Dim result As String
    Dim i As Integer
    Dim sheetName As String

    result = formula

    ' Per ogni foglio esportato, rimuove il prefisso '[NomeFile]'
    ' dalla formula, coprendo i tre formati possibili che Excel genera:
    For i = 1 To sheetCount
        sheetName = selectedSheets(i)
        ' Formato 1 - apice di chiusura dopo il nome foglio:
        '   '[File.xlsm]NomeFoglio'!  ?  NomeFoglio'!
        result = Replace(result, "[" & wbName & "]" & sheetName & "'", sheetName & "'")
        ' Formato 2 - apici attorno a tutto il blocco file+foglio:
        '   '[File.xlsm]NomeFoglio'!  ?  NomeFoglio!
        result = Replace(result, "'[" & wbName & "]" & sheetName & "'!", sheetName & "!")
        ' Formato 3 - nessun apice (nomi senza spazi):
        '   [File.xlsm]NomeFoglio!  ?  NomeFoglio!
        result = Replace(result, "[" & wbName & "]" & sheetName & "!", sheetName & "!")
    Next i

    ' Dopo aver rimosso tutti i riferimenti ai fogli esportati,
    ' se nella formula compare ancora '[NomeFile]' significa che
    ' c'e' almeno un riferimento a un foglio NON esportato.
    ' In questo caso restituisce "" per segnalare la conversione in valore.
    If InStr(result, "[" & wbName & "]") > 0 Then
        RipulisciFormula = ""
    Else
        RipulisciFormula = result
    End If
End Function
