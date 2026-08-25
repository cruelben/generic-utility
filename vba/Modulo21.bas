Attribute VB_Name = "Modulo21"
Option Explicit

' ======================================================================================
' MACRO: m21_SemplificaCollegamentiCatena
'
' DESCRIZIONE:
' Semplifica e riordina le catene di collegamenti indiretti tra formule all'interno dell'intero
' cartella di lavoro Excel.
'
' LOGICA DI FUNZIONAMENTO:
' 1. Scansiona tutte le celle contenenti formule in tutti i fogli del file.
' 2. Individua i collegamenti semplici a singola cella (es. =Foglio1!A1) ed esclude formule
'    complesse, calcoli matematici o funzioni (es. SOMMA, +, *).
' 3. Risale la catena di dipendenze fino alla cella sorgente originale (valore statico).
' 4. Gestisce la priorità del foglio attivo in base alla scelta dell'utente.
' 5. Richiede una CONFERMA sintetica con titolo dinamico personalizzato.
' 6. Previene blocchi da riferimenti circolari tramite un registro interno delle visite.
' ======================================================================================

Public Sub m21_SemplificaCollegamentiCatena()
    Dim ws As Worksheet
    Dim cellaFormula As Range
    Dim celleConFormule As Range
    Dim cellaSorgente As Range
    Dim nuovaFormula As String
    Dim contatoreModifiche As Long
    Dim rispMsg As VbMsgBoxResult
    Dim confermataEsecuzione As VbMsgBoxResult
    Dim foglioPrioritario As Worksheet
    Dim testoMESSAGGIO As String
    Dim titoloMsg2 As String
    Dim nomeFoglioMaiuscolo As String
    
    ' Ottiene il nome del foglio attivo convertito in maiuscolo
    nomeFoglioMaiuscolo = UCase(ActiveSheet.Name)
    
    ' --- FASE 1: Interazione Utente e Scelta Modalità ---
    rispMsg = MsgBox("Vuoi dare la priorità al foglio attivo (""" & nomeFoglioMaiuscolo & """) come punto intermedio?" & vbCrLf & vbCrLf & _
                     "• SÌ: Le formule che passano da questo foglio punteranno ad esso." & vbCrLf & _
                     "• NO: Tutte le formule punteranno direttamente al valore sorgente primario." & vbCrLf & _
                     "• ANNULLA: Interrompe la macro.", _
                     vbYesNoCancel + vbQuestion, "Priorità foglio " & nomeFoglioMaiuscolo & " ?")
                     
    If rispMsg = vbCancel Then Exit Sub
    
    ' Imposta il foglio di priorità, il testo e il titolo per il secondo messaggio di conferma
    If rispMsg = vbYes Then
        Set foglioPrioritario = ActiveSheet
        testoMESSAGGIO = "puntamento prioritario al foglio: " & nomeFoglioMaiuscolo & vbCrLf & _
                         "e per i restanti ai valori origine." & vbCrLf & _
                         "confermi la riscrittura?"
        titoloMsg2 = "priorità " & nomeFoglioMaiuscolo
    Else
        Set foglioPrioritario = Nothing
        testoMESSAGGIO = "puntamento diretto alle celle con i VALORI ORIGINE" & vbCrLf & _
                         "confermi la riscrittura?"
        titoloMsg2 = "VALORI ORIGINE"
    End If
    
    ' --- FASE 1B: SECONDO MSGBOX DI CONFERMA SINTETICO ---
    confermataEsecuzione = MsgBox(testoMESSAGGIO, _
                                  vbYesNo + vbExclamation + vbDefaultButton2, _
                                  titoloMsg2)
                                  
    ' Se l'utente clicca NO oppure chiude la finestra, interrompe l'esecuzione
    If confermataEsecuzione = vbNo Then Exit Sub
    
    ' --- FASE 2: Ottimizzazione Prestazioni Excel ---
    On Error GoTo GestioneErrori
    Application.ScreenUpdating = False         ' Disabilita il refresh dello schermo
    Application.Calculation = xlCalculationManual ' Mette il calcolo in manuale per evitare rallentamenti
    Application.EnableEvents = False           ' Disabilita gli eventi VBA della cartella
    
    contatoreModifiche = 0
    
    ' --- FASE 3: Scansione di tutti i Fogli e di tutte le Formule ---
    For Each ws In ActiveWorkbook.Worksheets
        ' Isola soltanto le celle del foglio che contengono effettivamente una formula
        On Error Resume Next
        Set celleConFormule = ws.Cells.SpecialCells(xlCellTypeFormulas)
        On Error GoTo GestioneErrori
        
        If Not celleConFormule Is Nothing Then
            For Each cellaFormula In celleConFormule
                ' Risale la catena dei collegamenti fino a trovare la sorgente corretta
                Set cellaSorgente = TrovaCellaOrigine(cellaFormula, foglioPrioritario)
                
                ' Verifica che la sorgente sia valida e non sia la cella stessa
                If Not cellaSorgente Is Nothing Then
                    If Not (cellaSorgente.Worksheet.Name = cellaFormula.Worksheet.Name _
                       And cellaSorgente.Address = cellaFormula.Address) Then
                        
                        ' Costruisci il testo della nuova formula (=NomeFoglio!Cella)
                        nuovaFormula = "=" & CostruisciRiferimento(cellaFormula, cellaSorgente)
                        
                        ' Se la nuova formula è diversa da quella esistente, sovrascrivi
                        If UCase(Trim(cellaFormula.formula)) <> UCase(Trim(nuovaFormula)) Then
                            cellaFormula.formula = nuovaFormula
                            contatoreModifiche = contatoreModifiche + 1
                        End If
                    End If
                End If
            Next cellaFormula
        End If
        Set celleConFormule = Nothing
    Next ws

    ' --- FASE 4: Ripristino Ambito e Notifica Finale ---
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "Operazione m21 completata con successo!" & vbCrLf & _
           "Formule semplificate e riordinate: " & contatoreModifiche, vbInformation, "Esito m21"
    Exit Sub

' --- GESTIONE ERRORI ---
GestioneErrori:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Si è verificato un errore durante l'esecuzione di m21: " & Err.Description, vbCritical, "Errore m21"
End Sub

' ======================================================================================
' FUNZIONE: TrovaCellaOrigine
' SCOPO: Risale la catena dei collegamenti cella per cella.
' NOTE: Se rileva il foglio prioritario lungo la catena (e la partenza non è su quel foglio),
'       interrompe la risalita lì; altrimenti prosegue fino al valore statico finale.
' ======================================================================================
Private Function TrovaCellaOrigine(ByVal cellaIniziale As Range, ByVal foglioPrioritario As Worksheet) As Range
    Dim cellaCorrente As Range
    Dim prossimaCella As Range
    Dim visite As Object
    Dim chiave As String
    Dim cellaIntermediaPrioritaria As Range
    
    ' Dictionary per tracciare gli indirizzi visitati e prevenire loop infiniti (riferimenti circolari)
    Set visite = CreateObject("Scripting.Dictionary")
    Set cellaCorrente = cellaIniziale
    Set cellaIntermediaPrioritaria = Nothing
    
    Do While True
        ' Verifica se la cella corrente punta semplicemente a un'altra cella
        Set prossimaCella = EstraiRiferimentoSingolo(cellaCorrente)
        
        If prossimaCella Is Nothing Then
            ' Raggiunta la radice (cella con valore statico o formula complessa non riducibile)
            Exit Do
        Else
            ' Controllo anti-loop circolare
            chiave = prossimaCella.Worksheet.Name & "!" & prossimaCella.Address
            If visite.Exists(chiave) Then
                Exit Do
            End If
            visite.Add chiave, True
            
            ' Verifica della priorità del foglio attivo
            If Not foglioPrioritario Is Nothing Then
                If prossimaCella.Worksheet.Name = foglioPrioritario.Name Then
                    ' Salva questo passaggio se la cella di partenza NON apparteneva già al foglio prioritario
                    If cellaIniziale.Worksheet.Name <> foglioPrioritario.Name Then
                        Set cellaIntermediaPrioritaria = prossimaCella
                    End If
                End If
            End If
            
            Set cellaCorrente = prossimaCella
        End If
    Loop
    
    ' Restituisce la cella intermedia del foglio prioritario oppure la sorgente finale
    If Not cellaIntermediaPrioritaria Is Nothing Then
        Set TrovaCellaOrigine = cellaIntermediaPrioritaria
    Else
        Set TrovaCellaOrigine = cellaCorrente
    End If
End Function

' ======================================================================================
' FUNZIONE: EstraiRiferimentoSingolo
' SCOPO: Analizza il testo della formula. Se è un puntamento diretto e pulito a una
'       singola cella (es. =Foglio1!A1 o =A1), restituisce l'oggetto Range puntato.
'       Se contiene operazioni (+, -, *, /), funzioni o intervalli, restituisce Nothing.
' ======================================================================================
Private Function EstraiRiferimentoSingolo(ByVal c As Range) As Range
    If Not c.HasFormula Then Exit Function
    
    Dim fml As String
    fml = Trim(c.formula)
    If Left(fml, 1) = "=" Then fml = Trim(Mid(fml, 2))
    
    ' Scarta qualsiasi formula che contenga operatori matematici, intervalli o funzioni
    If InStr(fml, "+") > 0 Or InStr(fml, "-") > 0 Or InStr(fml, "*") > 0 Or _
       InStr(fml, "/") > 0 Or InStr(fml, "(") > 0 Or InStr(fml, ":") > 0 Or _
       InStr(fml, ",") > 0 Or InStr(fml, ";") > 0 Or InStr(fml, "&") > 0 Or _
       InStr(fml, "^") > 0 Then
        Exit Function
    End If
    
    On Error Resume Next
    Dim target As Range
    Dim posPuntoEscl As Long
    Dim nomeFoglio As String
    Dim indirizzoCella As String
    
    posPuntoEscl = InStrRev(fml, "!")
    
    If posPuntoEscl > 0 Then
        ' Riferimento tra fogli diversi (es. 'Foglio Dati'!A1)
        nomeFoglio = Left(fml, posPuntoEscl - 1)
        indirizzoCella = Mid(fml, posPuntoEscl + 1)
        
        ' Rimuovi apici dal nome del foglio se presenti
        If Left(nomeFoglio, 1) = "'" And Right(nomeFoglio, 1) = "'" Then
            nomeFoglio = Mid(nomeFoglio, 2, Len(nomeFoglio) - 2)
        End If
        
        Set target = c.Worksheet.Parent.Worksheets(nomeFoglio).Range(indirizzoCella)
    Else
        ' Riferimento all'interno dello stesso foglio
        Set target = c.Worksheet.Range(fml)
    End If
    On Error GoTo 0
    
    ' Conferma che il target sia una cella singola
    If Not target Is Nothing Then
        If target.Cells.count = 1 Then
            Set EstraiRiferimentoSingolo = target
        End If
    End If
End Function

' ======================================================================================
' FUNZIONE: CostruisciRiferimento
' SCOPO: Formatta correttamente la stringa di testo della nuova formula.
' NOTE: Gestisce automaticamente gli apici singoli per i fogli contenenti spazi
'       ed elimina il nome del foglio se la formula punta a una cella dello stesso foglio.
' ======================================================================================
Private Function CostruisciRiferimento(ByVal cellaOrigineFormula As Range, ByVal cellaDestinazione As Range) As String
    Dim nomeFoglio As String
    nomeFoglio = cellaDestinazione.Worksheet.Name
    
    ' Racchiude il nome del foglio tra apici se sono presenti spazi
    If InStr(nomeFoglio, " ") > 0 Then
        nomeFoglio = "'" & nomeFoglio & "'"
    End If
    
    ' Se la cella sorgente è nello stesso foglio della formula attiva
    If cellaDestinazione.Worksheet.Name = cellaOrigineFormula.Worksheet.Name Then
        CostruisciRiferimento = cellaDestinazione.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Else
        CostruisciRiferimento = nomeFoglio & "!" & cellaDestinazione.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    End If
End Function

