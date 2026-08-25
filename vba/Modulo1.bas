Attribute VB_Name = "Modulo1"
' ============================================================
'  MACRO: confronta_fogli
'  Confronta il foglio attivo (deve avere suffisso " (2)")
'  con il corrispondente foglio originale (senza suffisso)
'  e colora di ROSSO nel foglio " (2)" le celle cambiate.
' ============================================================
Sub m1_confronta_fogli()
    Dim wsNew   As Worksheet
    Dim wsRif   As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r       As Long
    Dim c       As Long
    Dim valNew  As Variant
    Dim valRif  As Variant
    Dim changed As Long
    Dim nomeRif As String
    Dim nomeNew As String
    Dim suffisso As String
    suffisso = " (2)"
    ' --- Il foglio attivo deve essere il " (2)" ---
    Set wsRif = ActiveSheet
    nomeRif = wsRif.Name
    If Right(nomeRif, Len(suffisso)) <> suffisso Then
        MsgBox "Il foglio attivo non termina con '" & suffisso & "'." & vbCrLf & _
               "Seleziona prima il foglio 'nome (2)' che vuoi confrontare.", _
               vbExclamation, "Foglio non valido"
        Exit Sub
    End If
    ' --- Ricava il nome del foglio corrente (senza suffisso " (2)") ---
    nomeNew = Left(nomeRif, Len(nomeRif) - Len(suffisso))
    On Error Resume Next
    Set wsNew = ThisWorkbook.Sheets(nomeNew)
    On Error GoTo 0
    If wsNew Is Nothing Then
        MsgBox "Foglio corrispondente '" & nomeNew & "' non trovato nella cartella.", _
               vbCritical, "Errore"
        Exit Sub
    End If
    ' --- Conferma ---
    Dim risposta As Integer
    risposta = MsgBox("Confronto:" & vbCrLf & _
                      "  Foglio attuale:     " & nomeNew & vbCrLf & _
                      "  Foglio riferimento: " & nomeRif & vbCrLf & vbCrLf & _
                      "Gli sfondi rossi esistenti in '" & nomeRif & "' verranno rimossi" & vbCrLf & _
                      "e ricalcolate le differenze." & vbCrLf & vbCrLf & _
                      "Continuare?", vbYesNo + vbQuestion, "Confronto fogli")
    If risposta = vbNo Then Exit Sub
    ' --- Area di confronto ---
    lastRow = Application.WorksheetFunction.Max( _
                wsNew.UsedRange.Row + wsNew.UsedRange.Rows.count - 1, _
                wsRif.UsedRange.Row + wsRif.UsedRange.Rows.count - 1)
    lastCol = Application.WorksheetFunction.Max( _
                wsNew.UsedRange.Column + wsNew.UsedRange.Columns.count - 1, _
                wsRif.UsedRange.Column + wsRif.UsedRange.Columns.count - 1)
    ' --- Reset sfondi rossi in " (2)" ---
    wsRif.UsedRange.Interior.ColorIndex = xlNone
    wsRif.UsedRange.Font.ColorIndex = xlAutomatic
    Application.ScreenUpdating = False
    changed = 0
    For r = 1 To lastRow
        For c = 1 To lastCol
            valNew = wsNew.Cells(r, c).Value
            valRif = wsRif.Cells(r, c).Value
            If IsEmpty(valNew) Then valNew = ""
            If IsEmpty(valRif) Then valRif = ""
            If CStr(valNew) <> CStr(valRif) Then
                wsRif.Cells(r, c).Interior.Color = RGB(255, 0, 0)
                changed = changed + 1
            End If
        Next c
    Next r
    Application.ScreenUpdating = True
    If changed = 0 Then
        MsgBox "Nessuna differenza trovata tra '" & nomeNew & "' e '" & nomeRif & "'.", _
               vbInformation, "Confronto completato"
    Else
        MsgBox "Confronto completato." & vbCrLf & vbCrLf & _
               "Celle modificate evidenziate in '" & nomeRif & "': " & changed, _
               vbInformation, "Confronto completato"
    End If
End Sub

