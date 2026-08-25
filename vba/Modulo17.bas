Attribute VB_Name = "Modulo17"
' ============================================================
' ELIMINA il foglio selezionato in ListBox1Fogli
' ============================================================
Public Sub m17_EliminaFoglioSelezionato()
    Dim frm As UserForm1UTILITY
    Set frm = UserForm1UTILITY
    
    If frm.ListBox1Fogli.ListIndex = -1 Then
        MsgBox "Seleziona uno sheet da eliminare.", vbExclamation, "Nessun foglio selezionato"
        Exit Sub
    End If
    
    Dim nomeFoglio As String
    nomeFoglio = frm.ListBox1Fogli.Value
    
    If ThisWorkbook.Worksheets.count = 1 Then
        MsgBox "Impossibile eliminare l'unico foglio del file.", vbCritical, "Operazione non consentita"
        Exit Sub
    End If
    
    Dim risposta As VbMsgBoxResult
    risposta = MsgBox("Sei sicuro di voler eliminare il foglio:" & vbCrLf & _
                      """" & nomeFoglio & """" & vbCrLf & vbCrLf & _
                      "L'operazione è IRREVERSIBILE.", _
                      vbYesNo + vbCritical, "Conferma eliminazione foglio")
    
    If risposta = vbNo Then Exit Sub
    
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(nomeFoglio).Delete
    Application.DisplayAlerts = True
    
    Call m17_AggiornalistBox1Fogli
    
    MsgBox "Foglio """ & nomeFoglio & """ eliminato.", vbInformation, "Eliminazione completata"
End Sub


' ============================================================
' DUPLICA il foglio selezionato in ListBox1Fogli
' ============================================================
Public Sub m17_DuplicaFoglioSelezionato()
    Dim frm As UserForm1UTILITY
    Set frm = UserForm1UTILITY
    
    If frm.ListBox1Fogli.ListIndex = -1 Then
        MsgBox "Seleziona uno sheet da duplicare.", vbExclamation, "Nessun foglio selezionato"
        Exit Sub
    End If
    
    Dim nomeFoglio As String
    nomeFoglio = frm.ListBox1Fogli.Value
    
    Dim nuovoNome As String
    nuovoNome = InputBox("Inserisci il nome per il foglio duplicato:", _
                         "Duplica foglio: """ & nomeFoglio & """", nomeFoglio & " (2)")
    
    If nuovoNome = "" Then Exit Sub
    
    ' Verifica che il nome non esista già
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        If LCase(sh.Name) = LCase(nuovoNome) Then
            MsgBox "Esiste già un foglio con il nome """ & nuovoNome & """." & vbCrLf & _
                   "Scegli un nome diverso.", vbExclamation, "Nome già esistente"
            Exit Sub
        End If
    Next sh
    
    ThisWorkbook.Worksheets(nomeFoglio).Copy after:=ThisWorkbook.Worksheets(nomeFoglio)
    ActiveSheet.Name = nuovoNome
    
    Call m17_AggiornalistBox1Fogli
    
    ' Riseleziona il nuovo foglio nella ListBox
    Dim i As Integer
    For i = 0 To frm.ListBox1Fogli.ListCount - 1
        If frm.ListBox1Fogli.List(i) = nuovoNome Then
            frm.ListBox1Fogli.ListIndex = i
            Exit For
        End If
    Next i
    
    MsgBox "Foglio """ & nomeFoglio & """ duplicato come """ & nuovoNome & """.", _
           vbInformation, "Duplicazione completata"
End Sub


' ============================================================
' RINOMINA il foglio selezionato in ListBox1Fogli
' ============================================================
Public Sub m17_RinominaFoglioSelezionato()
    Dim frm As UserForm1UTILITY
    Set frm = UserForm1UTILITY
    
    If frm.ListBox1Fogli.ListIndex = -1 Then
        MsgBox "Seleziona uno sheet da rinominare.", vbExclamation, "Nessun foglio selezionato"
        Exit Sub
    End If
    
    Dim vecchioNome As String
    vecchioNome = frm.ListBox1Fogli.Value
    
    Dim nuovoNome As String
    nuovoNome = InputBox("Inserisci il nuovo nome per il foglio:", _
                         "Rinomina foglio: """ & vecchioNome & """", vecchioNome)
    
    If nuovoNome = "" Or nuovoNome = vecchioNome Then Exit Sub
    
    ' Verifica che il nome non esista già
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        If LCase(sh.Name) = LCase(nuovoNome) Then
            MsgBox "Esiste già un foglio con il nome """ & nuovoNome & """." & vbCrLf & _
                   "Scegli un nome diverso.", vbExclamation, "Nome già esistente"
            Exit Sub
        End If
    Next sh
    
    ThisWorkbook.Worksheets(vecchioNome).Name = nuovoNome
    
    Call m17_AggiornalistBox1Fogli
    
    ' Riseleziona il foglio rinominato nella ListBox
    Dim i As Integer
    For i = 0 To frm.ListBox1Fogli.ListCount - 1
        If frm.ListBox1Fogli.List(i) = nuovoNome Then
            frm.ListBox1Fogli.ListIndex = i
            Exit For
        End If
    Next i
    
    MsgBox "Foglio rinominato: """ & vecchioNome & """ -> """ & nuovoNome & """.", _
           vbInformation, "Rinomina completata"
End Sub


' ============================================================
' AGGIORNA ListBox1Fogli con tutti i fogli del workbook
' ============================================================
Public Sub m17_AggiornalistBox1Fogli()
    Dim frm As UserForm1UTILITY
    Set frm = UserForm1UTILITY
    
    Dim sh As Worksheet
    frm.ListBox1Fogli.Clear
    
    For Each sh In ThisWorkbook.Worksheets
        frm.ListBox1Fogli.AddItem sh.Name
    Next sh
End Sub

