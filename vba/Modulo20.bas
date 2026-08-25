Attribute VB_Name = "Modulo20"
Sub m20_Clean()
'
' Lavora SOLO sul foglio ATTIVO. Esegue:
'   1) Rimozione colore di sfondo, colore carattere, grassetto,
'      corsivo, sottolineato, barrato, TUTTI i bordi,
'      allineamenti (orizzontale/verticale/testo a capo),
'      formattazione condizionale, unione celle,
'      note, commenti, collegamenti ipertestuali
'      (NON tocca NumberFormat, nome font, dimensione font)
'   2) Rimozione di TUTTE le forme/shape (frecce, cerchi, caselle di testo,
'      immagini, smartart, ecc. presenti sul foglio)
'   3) Rimozione Blocca riquadri (Visualizza > Blocca riquadri)
'   4) Sblocco protezione foglio, se presente e SENZA password
'      (se protetta con password, la macro avvisa e NON procede a sbloccarla)
'   5) Rimozione nomi di cella (Names) definiti che puntano a questo foglio
'   6) Rimozione filtro automatico (AutoFilter)
'   7) Rilevamento (SENZA modificare) di eventuali righe/colonne nascoste,
'      con avviso nel messaggio finale se presenti
'   8) Rimozione raggruppamenti (Outline / Struttura)
'   9) Conversione di eventuali Tabelle Excel (ListObject) in range normale
'      (mantiene i dati, rimuove solo la "struttura tabella")
'  10) Rimozione validazione dati (menu a tendina, messaggi di errore/input)
'
    Dim ws As Worksheet
    Dim rng As Range
    Dim shp As Shape
    Dim risposta As VbMsgBoxResult
    Dim eraProtetto As Boolean
    Dim erroreSblocco As Boolean
    Dim numForme As Long
    Dim commentCells As Range
    Dim cel As Range
    Dim nm As Name
    Dim numCommenti As Long
    Dim numNomi As Long
    Dim numCollegamenti As Long
    Dim i As Long
    Dim lo As ListObject
    Dim numTabelle As Long
    Dim numRigheNascoste As Long, numColNascoste As Long
    
    Set ws = ActiveSheet
    
    ' --- Messaggio di conferma con Annulla come opzione di default ---
    risposta = MsgBox("Stai per pulire il foglio '" & ws.Name & "':" & vbCrLf & vbCrLf & _
                       "Verranno rimossi: formattazione colori/bordi/allineamenti, " & _
                       "formattazione condizionale, unione celle, note/commenti, " & _
                       "collegamenti, nomi cella, filtro automatico, " & _
                       "raggruppamenti, struttura tabelle, validazione dati, forme, " & _
                       "blocca riquadri e protezione (se senza password)." & vbCrLf & vbCrLf & _
                       "Formati numerici (%, €, decimali, migliaia) e font (nome/dimensione) " & _
                       "NON verranno modificati." & vbCrLf & vbCrLf & _
                       "Operazione non annullabile con CTRL+Z. Procedere?", _
                       vbOKCancel + vbExclamation + vbDefaultButton2, _
                       "Conferma pulizia foglio - min_Clean")
    
    If risposta = vbCancel Then
        MsgBox "Operazione annullata. Nessuna modifica effettuata.", vbInformation, "Annullato"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' ============================================================
    ' 1) SBLOCCO PROTEZIONE FOGLIO (se presente, senza password)
    ' ============================================================
    eraProtetto = ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios
    erroreSblocco = False
    
    If eraProtetto Then
        On Error Resume Next
        ws.Unprotect ' senza password
        If Err.Number <> 0 Then
            erroreSblocco = True
            Err.Clear
        End If
        On Error GoTo 0
    End If
    
    ' Se è protetta con password, non possiamo procedere con le modifiche
    If erroreSblocco Then
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        MsgBox "Il foglio '" & ws.Name & "' è protetto da PASSWORD." & vbCrLf & _
               "Impossibile sbloccarlo automaticamente: inserisci la password manualmente" & vbCrLf & _
               "(Revisione > Rimuovi protezione foglio) e poi rilancia la macro.", _
               vbCritical, "Protezione con password rilevata"
        Exit Sub
    End If
    
    ' ============================================================
    ' 2) RIMOZIONE BLOCCA RIQUADRI
    ' ============================================================
    If ws.Parent.Windows(1).FreezePanes Then
        ws.Activate
        ActiveWindow.FreezePanes = False
    End If
    
    ' ============================================================
    ' 3) RIMOZIONE FORMATTAZIONE: sfondo, font color, grassetto,
    '    corsivo, sottolineato, barrato, bordi
    ' ============================================================
    On Error Resume Next
    Set rng = ws.UsedRange
    On Error GoTo 0
    
    If Not rng Is Nothing Then
        ' Sfondo
        rng.Interior.ColorIndex = xlNone
        rng.Interior.Pattern = xlNone
        
        ' Font: colore, grassetto, corsivo, sottolineato, barrato
        rng.Font.ColorIndex = xlAutomatic
        rng.Font.Bold = False
        rng.Font.Italic = False
        rng.Font.Underline = xlUnderlineStyleNone
        rng.Font.Strikethrough = False
        
        ' Bordi: rimuove contorno, interni, diagonali
        rng.Borders.LineStyle = xlNone
        
        ' Allineamenti: riporta ai valori predefiniti di Excel
        rng.HorizontalAlignment = xlGeneral
        rng.VerticalAlignment = xlBottom
        rng.WrapText = False
        rng.Orientation = 0
        rng.IndentLevel = 0
        
        ' Formattazione condizionale: rimuove tutte le regole presenti
        rng.FormatConditions.Delete
        
        ' Unione celle: separa eventuali celle unite nel range usato
        rng.UnMerge
        
        ' Note e commenti (sia note classiche che commenti moderni)
        On Error Resume Next
        Set commentCells = rng.SpecialCells(xlCellTypeComments)
        On Error GoTo 0
        
        numCommenti = 0
        If Not commentCells Is Nothing Then
            numCommenti = commentCells.count
            For Each cel In commentCells
                cel.ClearComments
                On Error Resume Next
                If Not cel.CommentThreaded Is Nothing Then cel.CommentThreaded.Delete
                On Error GoTo 0
            Next cel
        End If
        
        ' Collegamenti ipertestuali sul foglio
        numCollegamenti = ws.Hyperlinks.count
        ws.Hyperlinks.Delete
        
        ' Validazione dati (menu a tendina, messaggi di errore/input)
        On Error Resume Next
        rng.Validation.Delete
        On Error GoTo 0
    End If
    
    ' ============================================================
    ' 4) RIMOZIONE DI TUTTE LE FORME/SHAPE
    ' ============================================================
    numForme = ws.Shapes.count
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    
    ' ============================================================
    ' 5) RIMOZIONE NOMI DI CELLA (Names) RIFERITI A QUESTO FOGLIO
    ' ============================================================
    numNomi = 0
    
    ' Nomi con scope limitato al foglio (worksheet-level names)
    For Each nm In ws.Names
        On Error Resume Next
        nm.Delete
        If Err.Number = 0 Then numNomi = numNomi + 1
        Err.Clear
        On Error GoTo 0
    Next nm
    
    ' Nomi con scope a livello di intera cartella di lavoro (workbook-level names)
    ' ma che puntano a celle di QUESTO foglio
    For i = ThisWorkbook.Names.count To 1 Step -1
        On Error Resume Next
        If Not ThisWorkbook.Names(i).RefersToRange Is Nothing Then
            If ThisWorkbook.Names(i).RefersToRange.Worksheet.Name = ws.Name Then
                ThisWorkbook.Names(i).Delete
                numNomi = numNomi + 1
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next i
    
    ' ============================================================
    ' 6) RIMOZIONE FILTRO AUTOMATICO (AutoFilter)
    ' ============================================================
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    
    ' ============================================================
    ' 7) RILEVAMENTO RIGHE/COLONNE NASCOSTE (NON vengono modificate)
    ' ============================================================
    numRigheNascoste = 0
    numColNascoste = 0
    
    If Not rng Is Nothing Then
        Dim rw As Range, cl As Range
        For Each rw In rng.Rows
            If rw.Hidden Then numRigheNascoste = numRigheNascoste + 1
        Next rw
        For Each cl In rng.Columns
            If cl.Hidden Then numColNascoste = numColNascoste + 1
        Next cl
    End If
    
    ' ============================================================
    ' 8) RIMOZIONE RAGGRUPPAMENTI (Struttura/Outline)
    ' ============================================================
    On Error Resume Next
    ws.Cells.ClearOutline
    On Error GoTo 0
    
    ' ============================================================
    ' 9) CONVERSIONE TABELLE EXCEL (ListObject) IN RANGE NORMALE
    '    (mantiene i dati, rimuove solo la struttura tabella)
    ' ============================================================
    numTabelle = ws.ListObjects.count
    For Each lo In ws.ListObjects
        On Error Resume Next
        lo.Unlist
        On Error GoTo 0
    Next lo
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' ============================================================
    ' Messaggio finale riepilogativo
    ' ============================================================
    Dim msgFinale As String
    msgFinale = "Operazioni completate sul foglio '" & ws.Name & "':" & vbCrLf & vbCrLf
    msgFinale = msgFinale & "  • Colori sfondo/font, grassetto, corsivo, sottolineato, barrato rimossi" & vbCrLf
    msgFinale = msgFinale & "  • Bordi rimossi" & vbCrLf
    msgFinale = msgFinale & "  • Allineamenti ripristinati ai valori predefiniti" & vbCrLf
    msgFinale = msgFinale & "  • Formattazione condizionale rimossa" & vbCrLf
    msgFinale = msgFinale & "  • Unione celle rimossa (se presente)" & vbCrLf
    msgFinale = msgFinale & "  • Note/commenti rimossi: " & numCommenti & vbCrLf
    msgFinale = msgFinale & "  • Collegamenti ipertestuali rimossi: " & numCollegamenti & vbCrLf
    msgFinale = msgFinale & "  • Nomi di cella rimossi: " & numNomi & vbCrLf
    msgFinale = msgFinale & "  • Forme rimosse: " & numForme & vbCrLf
    msgFinale = msgFinale & "  • Filtro automatico: rimosso (se presente)" & vbCrLf
    If numRigheNascoste > 0 Or numColNascoste > 0 Then
        msgFinale = msgFinale & "  • ATTENZIONE: rilevate " & numRigheNascoste & _
                    " righe e " & numColNascoste & " colonne nascoste (NON modificate)" & vbCrLf
    Else
        msgFinale = msgFinale & "  • Righe/colonne nascoste: nessuna rilevata" & vbCrLf
    End If
    msgFinale = msgFinale & "  • Raggruppamenti (struttura): rimossi" & vbCrLf
    msgFinale = msgFinale & "  • Tabelle Excel convertite in range: " & numTabelle & vbCrLf
    msgFinale = msgFinale & "  • Blocca riquadri: rimosso (se presente)" & vbCrLf
    If eraProtetto Then
        msgFinale = msgFinale & "  • Protezione foglio: rimossa" & vbCrLf
    Else
        msgFinale = msgFinale & "  • Protezione foglio: non presente" & vbCrLf
    End If
    msgFinale = msgFinale & vbCrLf & "Formati numerici e font (nome/dimensione) mantenuti."
    
    MsgBox msgFinale, vbInformation, "min_Clean - Completato"
    
End Sub


