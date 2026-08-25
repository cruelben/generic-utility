Attribute VB_Name = "Modulo3"
Sub m3_VerificaSeReferenziate()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cellaSelezionata As Range
    Dim formulaTest As String
    Dim formulaTestApici As String
    Dim trovato As Boolean
    Dim riferimenti As String
    Dim reportFinale As String
    
    ' Inizializza la variabile per il report finale
    reportFinale = ""
    
    ' Cicla attraverso ogni cella selezionata
    For Each cellaSelezionata In Selection
        ' Formula senza apici
        formulaTest = cellaSelezionata.Worksheet.Name & "!" & cellaSelezionata.Address(False, False)
        ' Formula con apici per gestire i nomi dei fogli complessi
        formulaTestApici = "'" & cellaSelezionata.Worksheet.Name & "'!" & cellaSelezionata.Address(False, False)
        
        trovato = False
        riferimenti = "La cella " & cellaSelezionata.Address(False, False) & " è richiamata da: " & vbCrLf
        
        ' Cicla attraverso ogni foglio
        For Each ws In ThisWorkbook.Worksheets
            ' Cicla attraverso ogni cella con formule nel foglio
            For Each rng In ws.UsedRange
                If rng.HasFormula Then
                    ' Verifica se la formula contiene il riferimento alla cella selezionata
                    ' Verifica sia la formula senza apici sia con apici
                    If InStr(rng.formula, formulaTest) > 0 Or InStr(rng.formula, formulaTestApici) > 0 Then
                        riferimenti = riferimenti & ws.Name & "!" & rng.Address(False, False) & vbCrLf
                        trovato = True
                    End If
                End If
            Next rng
        Next ws
        
        ' Aggiungi il risultato della cella corrente al report finale solo se trovato
        If trovato Then
            reportFinale = reportFinale & riferimenti & vbCrLf
        End If
    Next cellaSelezionata
    
    ' Mostra il report finale solo se sono stati trovati riferimenti
    If Len(reportFinale) > 0 Then
        MsgBox reportFinale
    Else
        MsgBox "Nessuna delle celle selezionate è richiamata da altri fogli."
    End If
End Sub


