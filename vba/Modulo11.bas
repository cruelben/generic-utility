Attribute VB_Name = "Modulo11"
Sub m11_MostraErrori()

    Dim ws As Worksheet
    Dim cell As Range
    Dim contatore As Long
    Dim wsReport As Worksheet
    contatore = 0

    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("_Report_Errori")
    On Error GoTo 0

    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add(after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsReport.Name = "_Report_Errori"
    Else
        wsReport.Cells.Clear
    End If

    wsReport.Cells(1, 1) = "Foglio!Cella"
    wsReport.Cells(1, 2) = "Tipo errore"
    wsReport.Cells(1, 3) = "Descrizione"
    wsReport.Cells(1, 4) = "Formula"
    wsReport.Rows(1).Font.Bold = True
    wsReport.Rows(1).Interior.Color = RGB(68, 114, 196)
    wsReport.Rows(1).Font.Color = RGB(255, 255, 255)

    Dim riga As Long
    riga = 2

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "_Report_Errori" Or ws.Name = "_Report_Modifiche" Then GoTo ProssimoFoglio

        For Each cell In ws.UsedRange
            If IsError(cell.Value) Then

                Dim tipoErrore As String
                Dim descrizione As String
                Dim colore As Long
                tipoErrore = ""
                descrizione = ""

                Select Case CStr(cell.Value)
                    Case "Errore 2042"
                        tipoErrore = "#N/D"
                        descrizione = "Valore non disponibile — formula non trova corrispondenza (es. CERCA.VERT)"
                        colore = RGB(255, 199, 206)
                    Case "Errore 2015"
                        tipoErrore = "#VALORE!"
                        descrizione = "Tipo di dato errato — testo dove serve un numero o viceversa"
                        colore = RGB(255, 199, 206)
                    Case "Errore 2023"
                        tipoErrore = "#RIF!"
                        descrizione = "Riferimento non valido — cella o intervallo eliminato o spostato"
                        colore = RGB(255, 199, 206)
                    Case "Errore 2036"
                        tipoErrore = "#NOME?"
                        descrizione = "Nome non riconosciuto — funzione o nome definito inesistente"
                        colore = RGB(255, 235, 156)
                    Case "Errore 2007"
                        tipoErrore = "#DIV/0!"
                        descrizione = "Divisione per zero"
                        colore = RGB(255, 235, 156)
                    Case "Errore 2000"
                        tipoErrore = "#NULL!"
                        descrizione = "Intersezione vuota — due intervalli non si intersecano"
                        colore = RGB(198, 239, 206)
                    Case "Errore 2048"
                        tipoErrore = "#NUM!"
                        descrizione = "Valore numerico non valido — numero troppo grande/piccolo o radice negativa"
                        colore = RGB(198, 239, 206)
                    Case Else
                        tipoErrore = CStr(cell.Value)
                        descrizione = "Errore non classificato"
                        colore = RGB(220, 220, 220)
                End Select

                wsReport.Cells(riga, 1) = ws.Name & "!" & cell.Address(False, False)
                wsReport.Cells(riga, 2) = tipoErrore
                wsReport.Cells(riga, 3) = descrizione
                ' Colonna D: link dinamico alla cella in errore
                If InStr(ws.Name, " ") > 0 Then
                    wsReport.Cells(riga, 4).formula = "='" & ws.Name & "'!" & cell.Address(False, False)
                Else
                    wsReport.Cells(riga, 4).formula = "=" & ws.Name & "!" & cell.Address(False, False)
                End If
                wsReport.Cells(riga, 1).Interior.Color = colore
                wsReport.Cells(riga, 2).Interior.Color = colore
                wsReport.Cells(riga, 3).Interior.Color = colore
                wsReport.Cells(riga, 4).Interior.Color = colore

                riga = riga + 1
                contatore = contatore + 1
            End If
        Next cell
ProssimoFoglio:
    Next ws

    wsReport.Columns("A:D").AutoFit
    wsReport.Activate

    If contatore = 0 Then
        MsgBox "Nessun errore trovato in nessun foglio.", vbInformation, "Analisi errori"
    Else
        MsgBox contatore & " errori trovati." & vbCrLf & _
               "Vedi il foglio '_Report_Errori' per i dettagli.", vbInformation, "Analisi errori completata"
    End If

End Sub
