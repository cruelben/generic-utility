Attribute VB_Name = "Modulo12"
Public SnapshotValori As Object
Public SnapshotFormule As Object

Sub m12_ScattaSnapshot()

    Set SnapshotValori = CreateObject("Scripting.Dictionary")
    Set SnapshotFormule = CreateObject("Scripting.Dictionary")
    
    Dim ws   As Worksheet
    Dim cell As Range
    Dim chiave As String
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "_Report_Modifiche" Then
            For Each cell In ws.UsedRange
                chiave = ws.Name & "!" & cell.Address(False, False)
                
                If IsError(cell.Value) Then
                    SnapshotValori(chiave) = "ERR:" & CStr(cell.Value)
                Else
                    SnapshotValori(chiave) = cell.Value
                End If
                
                SnapshotFormule(chiave) = cell.formula
                
            Next cell
        End If
    Next ws
    
    Application.StatusBar = "Snapshot salvato alle " & Format(Now, "hh:mm:ss")

End Sub

Sub m12_MostraDifferenze()

    If SnapshotValori Is Nothing Then
        MsgBox "Nessuno snapshot disponibile." & vbCrLf & _
               "Esegui prima 'm12_ScattaSnapshot'.", vbExclamation
        Exit Sub
    End If
    
    Dim rChiave()    As String
    Dim rValPrima()  As String
    Dim rValOra()    As String
    Dim rFormPrima() As String
    Dim rFormOra()   As String
    Dim rTipo()      As String
    Dim contatore    As Long
    contatore = 0
    
    ReDim rChiave(0):    ReDim rValPrima(0):  ReDim rValOra(0)
    ReDim rFormPrima(0): ReDim rFormOra(0):   ReDim rTipo(0)
    
    Dim ws     As Worksheet
    Dim cell   As Range
    Dim chiave As String
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "_Report_Modifiche" Then
            For Each cell In ws.UsedRange
                chiave = ws.Name & "!" & cell.Address(False, False)
                
                Dim valOra  As String
                Dim formOra As String
                
                If IsError(cell.Value) Then
                    valOra = "ERR:" & CStr(cell.Value)
                Else
                    valOra = CStr(cell.Value)
                End If
                formOra = cell.formula
                
                If SnapshotValori.Exists(chiave) Then
                    
                    Dim valPrima  As String
                    Dim formPrima As String
                    valPrima = CStr(SnapshotValori(chiave))
                    formPrima = CStr(SnapshotFormule(chiave))
                    
                    Dim valCambiato As Boolean
                    valCambiato = False
                    If Left(valOra, 4) = "ERR:" Or Left(valPrima, 4) = "ERR:" Then
                        If valOra <> valPrima Then valCambiato = True
                    ElseIf IsNumeric(valOra) And IsNumeric(valPrima) Then
                        If Abs(CDbl(valOra) - CDbl(valPrima)) > 0.000001 Then valCambiato = True
                    ElseIf valOra <> valPrima Then
                        valCambiato = True
                    End If
                    
                    Dim formCambiata As Boolean
                    formCambiata = (formOra <> formPrima)
                    
                    Dim tipo As String
                    tipo = ""
                    If valCambiato And formCambiata Then
                        tipo = "Valore+Formula"
                    ElseIf formCambiata Then
                        tipo = "Formula"
                    ElseIf valCambiato Then
                        tipo = "Valore"
                    End If
                    
                    If tipo <> "" Then
                        ReDim Preserve rChiave(contatore)
                        ReDim Preserve rValPrima(contatore)
                        ReDim Preserve rValOra(contatore)
                        ReDim Preserve rFormPrima(contatore)
                        ReDim Preserve rFormOra(contatore)
                        ReDim Preserve rTipo(contatore)
                        rChiave(contatore) = chiave
                        rValPrima(contatore) = valPrima
                        rValOra(contatore) = valOra
                        rFormPrima(contatore) = formPrima
                        rFormOra(contatore) = formOra
                        rTipo(contatore) = tipo
                        contatore = contatore + 1
                    End If
                    
                Else
                    If valOra <> "" Or formOra <> "" Then
                        ReDim Preserve rChiave(contatore)
                        ReDim Preserve rValPrima(contatore)
                        ReDim Preserve rValOra(contatore)
                        ReDim Preserve rFormPrima(contatore)
                        ReDim Preserve rFormOra(contatore)
                        ReDim Preserve rTipo(contatore)
                        rChiave(contatore) = chiave
                        rValPrima(contatore) = ""
                        rValOra(contatore) = valOra
                        rFormPrima(contatore) = ""
                        rFormOra(contatore) = formOra
                        rTipo(contatore) = "Nuova"
                        contatore = contatore + 1
                    End If
                End If
                
            Next cell
        End If
    Next ws
    
    Dim k As Variant
    For Each k In SnapshotValori.Keys
        
        Dim parts()   As String
        parts = Split(k, "!")
        Dim nomeSheet As String: nomeSheet = parts(0)
        Dim indCell   As String: indCell = parts(1)
        
        If nomeSheet <> "_Report_Modifiche" Then
            
            On Error Resume Next
            Dim wsCheck As Worksheet
            Set wsCheck = Nothing
            Set wsCheck = ThisWorkbook.Sheets(nomeSheet)
            On Error GoTo 0
            
            If Not wsCheck Is Nothing Then
                Dim cellCheck As Range
                Set cellCheck = wsCheck.Range(indCell)
                
                Dim snapVal  As String: snapVal = CStr(SnapshotValori(k))
                Dim snapForm As String: snapForm = CStr(SnapshotFormule(k))
                
                Dim cellaVuota As Boolean
                If IsError(cellCheck.Value) Then
                    cellaVuota = False
                Else
                    cellaVuota = (cellCheck.Value = "" And cellCheck.formula = "")
                End If
                
                If cellaVuota Then
                    If snapVal <> "" Or snapForm <> "" Then
                        
                        Dim giaTrovata As Boolean: giaTrovata = False
                        Dim i As Long
                        For i = 0 To contatore - 1
                            If rChiave(i) = CStr(k) Then giaTrovata = True: Exit For
                        Next i
                        
                        If Not giaTrovata Then
                            ReDim Preserve rChiave(contatore)
                            ReDim Preserve rValPrima(contatore)
                            ReDim Preserve rValOra(contatore)
                            ReDim Preserve rFormPrima(contatore)
                            ReDim Preserve rFormOra(contatore)
                            ReDim Preserve rTipo(contatore)
                            rChiave(contatore) = CStr(k)
                            rValPrima(contatore) = snapVal
                            rValOra(contatore) = ""
                            rFormPrima(contatore) = snapForm
                            rFormOra(contatore) = ""
                            rTipo(contatore) = "Svuotata"
                            contatore = contatore + 1
                        End If
                        
                    End If
                End If
            End If
        End If
    Next k
    
    Dim wsReport As Worksheet
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("_Report_Modifiche")
    On Error GoTo 0
    
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add(after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsReport.Name = "_Report_Modifiche"
    Else
        wsReport.Cells.Clear
    End If
    
    Dim intestazioni As Variant
    intestazioni = Array("Foglio!Cella", "Valore prima", "Valore attuale", _
                         "Formula prima", "Formula attuale", "Tipo modifica")
    Dim c As Long
    For c = 0 To 5
        wsReport.Cells(1, c + 1) = intestazioni(c)
    Next c
    With wsReport.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    If contatore = 0 Then
        wsReport.Activate
        MsgBox "Nessuna modifica rilevata rispetto allo snapshot.", vbInformation, "Confronto"
        Exit Sub
    End If
    
    Dim colori(4) As Long
    colori(0) = RGB(198, 239, 206)
    colori(1) = RGB(255, 235, 156)
    colori(2) = RGB(255, 199, 206)
    colori(3) = RGB(189, 215, 238)
    colori(4) = RGB(255, 153, 204)
    
    For i = 0 To contatore - 1
        
        Dim riga As Long: riga = i + 2
        
        Dim parti() As String: parti = Split(rChiave(i), "!")
        Dim nFoglio As String: nFoglio = parti(0)
        Dim nCella  As String: nCella = parti(1)
        
        wsReport.Cells(riga, 1) = rChiave(i)
        wsReport.Cells(riga, 2) = rValPrima(i)
        wsReport.Cells(riga, 4) = rFormPrima(i)
        wsReport.Cells(riga, 6) = rTipo(i)
        
        If rTipo(i) = "Svuotata" Then
            wsReport.Cells(riga, 3) = "(vuota)"
            wsReport.Cells(riga, 5) = "(vuota)"
        Else
            wsReport.Cells(riga, 3).formula = "='" & nFoglio & "'!" & nCella
            wsReport.Cells(riga, 5) = rFormOra(i)
        End If
        
        Dim idxColore As Long
        Select Case rTipo(i)
            Case "Nuova":          idxColore = 0
            Case "Svuotata":       idxColore = 1
            Case "Valore":         idxColore = 2
            Case "Formula":        idxColore = 3
            Case "Valore+Formula": idxColore = 4
        End Select
        wsReport.Cells(riga, 1).Resize(1, 6).Interior.Color = colori(idxColore)
        
    Next i
    
    wsReport.Columns("A:F").AutoFit
    wsReport.Activate
    
    MsgBox contatore & " modifiche trovate rispetto allo snapshot." & vbCrLf & _
           "Vedi il foglio '_Report_Modifiche' per i dettagli.", _
           vbInformation, "Confronto completato"

End Sub

Sub m12_Ripristina()

    ' --- verifica snapshot disponibile ---
    If SnapshotValori Is Nothing Then
        MsgBox "Nessuno snapshot disponibile." & vbCrLf & _
               "Esegui prima 'm12_ScattaSnapshot'.", vbExclamation
        Exit Sub
    End If
    
    ' --- verifica che tutti i fogli dello snapshot esistano ancora ---
    Dim foglioMancante As String
    foglioMancante = ""
    
    Dim k As Variant
    For Each k In SnapshotValori.Keys
        Dim parts() As String
        parts = Split(k, "!")
        Dim nomeSheet As String: nomeSheet = parts(0)
        
        If nomeSheet <> "_Report_Modifiche" Then
            On Error Resume Next
            Dim wsTest As Worksheet
            Set wsTest = Nothing
            Set wsTest = ThisWorkbook.Sheets(nomeSheet)
            On Error GoTo 0
            
            If wsTest Is Nothing Then
                ' accumula nomi fogli mancanti senza duplicati
                If InStr(foglioMancante, "[" & nomeSheet & "]") = 0 Then
                    foglioMancante = foglioMancante & "[" & nomeSheet & "] "
                End If
            End If
        End If
    Next k
    
    If foglioMancante <> "" Then
        MsgBox "Ripristino non eseguito." & vbCrLf & vbCrLf & _
               "I seguenti fogli presenti nello snapshot non esistono piu' nel workbook:" & vbCrLf & _
               foglioMancante & vbCrLf & vbCrLf & _
               "Recupera i fogli eliminati e riprova.", vbCritical, "Ripristino impossibile"
        Exit Sub
    End If
    
    ' --- conferma utente ---
    Dim risposta As Integer
    risposta = MsgBox("Stai per ripristinare tutti i valori e le formule allo stato dello snapshot." & vbCrLf & _
                      "L'operazione sovrascrivera' le modifiche attuali." & vbCrLf & vbCrLf & _
                      "Continuare?", vbQuestion + vbYesNo, "Conferma ripristino")
    If risposta = vbNo Then Exit Sub
    
    ' --- esegui il ripristino ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim errori      As Long:  errori = 0
    Dim ripristinate As Long: ripristinate = 0
    Dim logErrori   As String: logErrori = ""
    
    For Each k In SnapshotValori.Keys
        
        parts = Split(k, "!")
        nomeSheet = parts(0)
        Dim indCell As String: indCell = parts(1)
        
        If nomeSheet <> "_Report_Modifiche" Then
            
            Dim ws   As Worksheet
            Dim cell As Range
            Set ws = ThisWorkbook.Sheets(nomeSheet)
            Set cell = ws.Range(indCell)
            
            Dim snapForm As String: snapForm = CStr(SnapshotFormule(k))
            Dim snapVal  As Variant: snapVal = SnapshotValori(k)
            
            On Error Resume Next
            
            If snapForm <> "" Then
                ' era una formula: la ripristina
                cell.formula = snapForm
            Else
                ' era un valore fisso: ripristina il valore
                ' se era un errore salvato come "ERR:..." svuota la cella
                If Left(CStr(snapVal), 4) = "ERR:" Then
                    cell.ClearContents
                Else
                    cell.Value = snapVal
                End If
            End If
            
            If Err.Number <> 0 Then
                errori = errori + 1
                If Len(logErrori) < 500 Then
                    logErrori = logErrori & CStr(k) & " (" & Err.Description & ")" & vbCrLf
                End If
                Err.Clear
            Else
                ripristinate = ripristinate + 1
            End If
            
            On Error GoTo 0
            
        End If
    Next k
    
    ' --- celle che nello snapshot erano vuote ma ora hanno contenuto: le svuota ---
    Dim ws2  As Worksheet
    Dim cell2 As Range
    For Each ws2 In ThisWorkbook.Worksheets
        If ws2.Name <> "_Report_Modifiche" Then
            For Each cell2 In ws2.UsedRange
                Dim chiave2 As String
                chiave2 = ws2.Name & "!" & cell2.Address(False, False)
                If Not SnapshotValori.Exists(chiave2) Then
                    On Error Resume Next
                    cell2.ClearContents
                    On Error GoTo 0
                End If
            Next cell2
        End If
    Next ws2
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' --- messaggio finale ---
    If errori = 0 Then
        MsgBox ripristinate & " celle ripristinate con successo.", _
               vbInformation, "Ripristino completato"
    Else
        MsgBox ripristinate & " celle ripristinate." & vbCrLf & _
               errori & " celle non ripristinate per errore:" & vbCrLf & vbCrLf & _
               logErrori, vbExclamation, "Ripristino completato con avvisi"
    End If

End Sub

