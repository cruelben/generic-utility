Attribute VB_Name = "Modulo6"
Option Explicit

' ======================================================================================
' MACRO: m6_estrazione (Versione Doppio Hub con Mappatura Destinazioni)
'
' SCOPO GENERALE:
'   Isolare e separare completamente le dipendenze da/verso un gruppo di fogli selezionati.
'   La macro crea due fogli Hub:
'     1. Hub Estrazione (default: "Estrazione"): Raccoglie i riferimenti contenuti nei
'        Fogli Selezionati che puntano ai Fogli Non Selezionati.
'     2. Hub Restante (default: "Restante"): Raccoglie i riferimenti contenuti nei
'        Fogli Non Selezionati che puntano ai Fogli Selezionati.
'
' STRUTTURA FOGLI HUB (Colonne A, B, C):
'   - Colonna A: Origine (Foglio e Cella sorgente).
'   - Colonna B: Valore / Collegamento Hub.
'   - Colonna C: "Nuova Destinazione" (Elenca il foglio e la cella dove il valore viene usato.
'                In caso di utilizzi multipli, le destinazioni vengono concatenate con '; ').
'
' COMPORTAMENTO CONFERMA UTENTE:
'   - OK: Compila i due fogli Hub E aggiorna le formule originali in tutti i fogli.
'   - ANNULLA: Compila i due fogli Hub (con elenchi e destinazioni) SENZA modificare le formule originali.
'
' NOTE PER MANUTENZIONE FUTURA:
'   1. PREVENZIONE CRASH EXCEL: Sempre rompere il gruppo di selezione prima di aggiungere fogli (Punto 4).
'   2. MAPPATURA DESTINAZIONI: Gestita dentro ProcessScanAndHub salvando la riga dell'Hub nel dizionario
'      refMap (memorizzato come "Riga|TargetRef"). Se la sorgente esiste già, aggiunge la nuova destinazione in Colonna C.
' ======================================================================================

Sub m6_estrazione()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    ' ----------------------------------------------------------------------------------
    ' 1. RICHIESTA NOMI PER I DUE FOGLI HUB ALL'UTENTE
    ' ----------------------------------------------------------------------------------
    Dim rifName1 As String, rifName2 As String
    
    rifName1 = InputBox("Inserisci il nome per il PRIMO foglio Hub (Selezionati -> Non Selezionati):", "Nome Hub Estrazione", "Estrazione")
    rifName1 = Trim(rifName1)
    If rifName1 = "" Then
        MsgBox "Operazione annullata dall'utente.", vbInformation, "Annullato"
        Exit Sub
    End If
    
    rifName2 = InputBox("Inserisci il nome per il SECONDO foglio Hub (Non Selezionati -> Selezionati):", "Nome Hub Restante", "Restante")
    rifName2 = Trim(rifName2)
    If rifName2 = "" Then
        MsgBox "Operazione annullata dall'utente.", vbInformation, "Annullato"
        Exit Sub
    End If
    
    If UCase(rifName1) = UCase(rifName2) Then
        MsgBox "I due fogli Hub devono avere nomi diversi tra loro. Operazione annullata.", vbExclamation, "Nomi Duplicati"
        Exit Sub
    End If
    
    ' ----------------------------------------------------------------------------------
    ' 2. CONTROLLO ESISTENZA FOGLI HUB NEL WORKBOOK
    ' ----------------------------------------------------------------------------------
    Dim wsCheck As Worksheet
    For Each wsCheck In wb.Worksheets
        If UCase(wsCheck.Name) = UCase(rifName1) Or UCase(wsCheck.Name) = UCase(rifName2) Then
            MsgBox "Uno o entrambi i fogli Hub ('" & rifName1 & "', '" & rifName2 & "') esistono già nella cartella di lavoro. Operazione annullata.", vbExclamation, "Foglio Esistente"
            Exit Sub
        End If
    Next wsCheck
    
    ' ----------------------------------------------------------------------------------
    ' 3. MEMORIZZAZIONE FOGLI SELEZIONATI DALL'UTENTE
    ' ----------------------------------------------------------------------------------
    Dim selNames As Object
    Set selNames = CreateObject("Scripting.Dictionary")
    Dim sh As Object
    
    For Each sh In ActiveWindow.selectedSheets
        If TypeOf sh Is Worksheet Then
            ' Esclude i nomi dei fogli Hub da creare
            If UCase(sh.Name) <> UCase(rifName1) And UCase(sh.Name) <> UCase(rifName2) Then
                selNames(sh.Name) = True
            End If
        End If
    Next sh
    
    If selNames.count = 0 Then
        MsgBox "Selezionare almeno un foglio di lavoro valido.", vbExclamation, "Attenzione"
        Exit Sub
    End If
    
    ' ----------------------------------------------------------------------------------
    ' 4. PREVENZIONE CRASH EXCEL: DESELEZIONE GRUPPO FOGLI
    ' ----------------------------------------------------------------------------------
    On Error Resume Next
    wb.Worksheets(CStr(selNames.Keys()(0))).Select
    On Error GoTo 0
    
    ' ----------------------------------------------------------------------------------
    ' 5. IDENTIFICAZIONE FOGLI NON SELEZIONATI
    ' ----------------------------------------------------------------------------------
    Dim nonSelSheets As Object
    Set nonSelSheets = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    
    For Each ws In wb.Worksheets
        If Not selNames.Exists(ws.Name) And UCase(ws.Name) <> UCase(rifName1) And UCase(ws.Name) <> UCase(rifName2) Then
            nonSelSheets(ws.Name) = True
        End If
    Next ws
    
    If nonSelSheets.count = 0 Then
        MsgBox "Tutti i fogli della cartella risultano selezionati. Selezionare solo un sottoinsieme di fogli.", vbInformation, "Attenzione"
        Exit Sub
    End If
    
    ' ----------------------------------------------------------------------------------
    ' 6. DOMANDA ALL'UTENTE: SOSTITUZIONE AUTOMATICA O SOLO COMPILAZIONE CONSULTAZIONE?
    ' ----------------------------------------------------------------------------------
    Dim risposta As VbMsgBoxResult
    Dim aggiornaFormule As Boolean
    
    risposta = MsgBox("Si desidera procedere con l'aggiornamento automatico dei collegamenti nelle formule verso i due fogli Hub?" & vbCrLf & vbCrLf & _
                      "• OK: Compila i due fogli Hub E aggiorna le formule originali in tutti i fogli interessati." & vbCrLf & _
                      "• ANNULLA: Compila comunque i due fogli Hub (con l'elenco delle destinazioni) SENZA modificare alcuna formula nei fogli.", _
                      vbOKCancel + vbQuestion, "Opzioni di Estrazione Bidirezionale")
                      
    aggiornaFormule = (risposta = vbOK)

    ' ----------------------------------------------------------------------------------
    ' 7. OPTIMIZATION: DISATTIVAZIONE RICALCOLO ED EVENTI PER MASSIMA VELOCITÀ
    ' ----------------------------------------------------------------------------------
    Dim oldCalc As XlCalculation
    oldCalc = Application.Calculation
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo CleanUp
    
    ' ----------------------------------------------------------------------------------
    ' 8. CREAZIONE DEI DUE FOGLI HUB E INTESTAZIONI (COLONNE A, B, C)
    ' ----------------------------------------------------------------------------------
    Dim wsHub1 As Worksheet, wsHub2 As Worksheet
    
    ' Hub 1: Estrazione (Selezionati -> Non Selezionati)
    Set wsHub1 = wb.Worksheets.Add(after:=wb.Worksheets(wb.Worksheets.count))
    wsHub1.Name = rifName1
    wsHub1.Cells(1, 1).Value = "Origine (Foglio Non Selezionato)"
    wsHub1.Cells(1, 2).Value = "Valore (Collegamento Hub Cella Singola)"
    wsHub1.Cells(1, 3).Value = "Nuova Destinazione"
    wsHub1.Rows(1).Font.Bold = True
    
    ' Hub 2: Restante (Non Selezionati -> Selezionati)
    Set wsHub2 = wb.Worksheets.Add(after:=wb.Worksheets(wb.Worksheets.count))
    wsHub2.Name = rifName2
    wsHub2.Cells(1, 1).Value = "Origine (Foglio Selezionato)"
    wsHub2.Cells(1, 2).Value = "Valore (Collegamento Hub Cella Singola)"
    wsHub2.Cells(1, 3).Value = "Nuova Destinazione"
    wsHub2.Rows(1).Font.Bold = True
    
    ' ----------------------------------------------------------------------------------
    ' 9. COSTRUZIONE PATTERN REGEX PER LE DUE DIREZIONI
    ' ----------------------------------------------------------------------------------
    Dim regEx1 As Object, regEx2 As Object
    Set regEx1 = CreateObject("VBScript.RegExp")
    Set regEx2 = CreateObject("VBScript.RegExp")
    
    ' Pattern 1: Cerca riferimenti ai fogli NON SELEZIONATI
    Dim sheetList1() As String
    ReDim sheetList1(0 To nonSelSheets.count - 1)
    Dim idx As Long: idx = 0
    Dim k As Variant
    For Each k In nonSelSheets.Keys
        sheetList1(idx) = EscapeRegExPattern(CStr(k))
        idx = idx + 1
    Next k
    regEx1.Global = True
    regEx1.IgnoreCase = True
    regEx1.Pattern = "('?(" & Join(sheetList1, "|") & ")'?![$]?[A-Za-z]+[$]?[0-9]+(:[$]?[A-Za-z]+[$]?[0-9]+)?)"
    
    ' Pattern 2: Cerca riferimenti ai fogli SELEZIONATI
    Dim sheetList2() As String
    ReDim sheetList2(0 To selNames.count - 1)
    idx = 0
    For Each k In selNames.Keys
        sheetList2(idx) = EscapeRegExPattern(CStr(k))
        idx = idx + 1
    Next k
    regEx2.Global = True
    regEx2.IgnoreCase = True
    regEx2.Pattern = "('?(" & Join(sheetList2, "|") & ")'?![$]?[A-Za-z]+[$]?[0-9]+(:[$]?[A-Za-z]+[$]?[0-9]+)?)"
    
    ' ----------------------------------------------------------------------------------
    ' 10. PASSAGGIO 1: SCANSIONE FOGLI SELEZIONATI -> COMPILAZIONE HUB 1 (ESTRAZIONE)
    ' ----------------------------------------------------------------------------------
    Dim countMod1 As Long: countMod1 = 0
    Dim currRow1 As Long: currRow1 = 2
    Dim refMap1 As Object: Set refMap1 = CreateObject("Scripting.Dictionary")
    
    Call ProcessScanAndHub(wb, selNames, regEx1, wsHub1, rifName1, refMap1, currRow1, aggiornaFormule, countMod1)
    
    ' ----------------------------------------------------------------------------------
    ' 11. PASSAGGIO 2: SCANSIONE FOGLI NON SELEZIONATI -> COMPILAZIONE HUB 2 (RESTANTE)
    ' ----------------------------------------------------------------------------------
    Dim countMod2 As Long: countMod2 = 0
    Dim currRow2 As Long: currRow2 = 2
    Dim refMap2 As Object: Set refMap2 = CreateObject("Scripting.Dictionary")
    
    Call ProcessScanAndHub(wb, nonSelSheets, regEx2, wsHub2, rifName2, refMap2, currRow2, aggiornaFormule, countMod2)

    ' Formattazione colonne nei fogli Hub
    wsHub1.Columns("A:C").AutoFit
    wsHub2.Columns("A:C").AutoFit

    ' ----------------------------------------------------------------------------------
    ' 12. MESSAGGIO FINALE DI RIEPILOGO
    ' ----------------------------------------------------------------------------------
    If aggiornaFormule Then
        MsgBox "Estrazione bidirezionale completata con successo!" & vbCrLf & vbCrLf & _
               "• Foglio Hub '" & rifName1 & "': " & (currRow1 - 2) & " voci estratte, " & countMod1 & " formule modificate nei fogli selezionati." & vbCrLf & _
               "• Foglio Hub '" & rifName2 & "': " & (currRow2 - 2) & " voci estratte, " & countMod2 & " formule modificate nei fogli non selezionati.", _
               vbInformation, "Completato"
    Else
        MsgBox "Estrazione completata in modalità consultazione!" & vbCrLf & vbCrLf & _
               "• Foglio Hub '" & rifName1 & "': " & (currRow1 - 2) & " voci individuate." & vbCrLf & _
               "• Foglio Hub '" & rifName2 & "': " & (currRow2 - 2) & " voci individuate." & vbCrLf & _
               "• Le formule nei fogli della cartella NON sono state modificate.", _
               vbInformation, "Consultazione Completata"
    End If

' --------------------------------------------------------------------------------------
' 13. CLEANUP & RIPRISTINO AMBIENTE EXCEL
' --------------------------------------------------------------------------------------
CleanUp:
    Application.StatusBar = False
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

' ======================================================================================
' ROUTINE DI SCANSIONE E COMPILAZIONE HUB CON MAPPATURA DESTINAZIONI (COLONNA C)
' ======================================================================================
Private Sub ProcessScanAndHub(ByRef wb As Workbook, _
                             ByRef targetSheetsDict As Object, _
                             ByRef regExObj As Object, _
                             ByRef wsHub As Worksheet, _
                             ByVal hubName As String, _
                             ByRef refMap As Object, _
                             ByRef currRow As Long, _
                             ByVal doUpdateFormulas As Boolean, _
                             ByRef countModified As Long)
                             
    Dim sheetNameKey As Variant
    Dim ws As Worksheet
    Dim formulaCells As Range, cell As Range
    Dim cellCount As Long
    Dim matches As Object, match As Object
    Dim fml As String, origRef As String, cleanKey As String, targetRef As String, rangePart As String
    Dim rawSheetName As String, destCellAddr As String
    Dim existingRow As Long, mapData As Variant
    
    For Each sheetNameKey In targetSheetsDict.Keys
        Set ws = wb.Worksheets(CStr(sheetNameKey))
        
        ' Ignora i fogli protetti da struttura/password
        If Not ws.ProtectContents Then
            Set formulaCells = Nothing
            On Error Resume Next
            Set formulaCells = Intersect(ws.UsedRange, ws.Cells.SpecialCells(xlCellTypeFormulas))
            On Error GoTo 0
            
            If Not formulaCells Is Nothing Then
                cellCount = 0
                
                For Each cell In formulaCells.Cells
                    cellCount = cellCount + 1
                    If cellCount Mod 50 = 0 Then DoEvents
                    
                    ' Salta formule matriciali e celle unite
                    If cell.HasArray Then GoTo NextCellInLoop
                    If cell.MergeCells Then GoTo NextCellInLoop
                    
                    fml = ""
                    On Error Resume Next
                    fml = cell.formula
                    On Error GoTo 0
                    
                    If fml <> "" Then
                        If regExObj.Test(fml) Then
                            Set matches = regExObj.Execute(fml)
                            
                            ' Costruisce la stringa con nome foglio e cella di destinazione attuale (es. 'Foglio10'!K21)
                            destCellAddr = "'" & ws.Name & "'!" & cell.Address(False, False)
                            
                            For Each match In matches
                                origRef = match.Value
                                rangePart = Mid(origRef, InStr(origRef, "!") + 1)
                                
                                rawSheetName = Left(origRef, InStr(origRef, "!") - 1)
                                If Left(rawSheetName, 1) = "'" And Right(rawSheetName, 1) = "'" Then
                                    rawSheetName = Mid(rawSheetName, 2, Len(rawSheetName) - 2)
                                End If
                                
                                cleanKey = "'" & rawSheetName & "'!" & rangePart
                                
                                ' --------------------------------------------------------------
                                ' COMPILAZIONE FOGLIO HUB (COLONNE A, B, C)
                                ' --------------------------------------------------------------
                                If InStr(rangePart, ":") > 0 Then
                                    ' CASO INTERVALLO/MATRICE: Replicato alle stesse coordinate
                                    On Error Resume Next
                                    wsHub.Range(rangePart).formula = "='" & rawSheetName & "'!" & rangePart
                                    On Error GoTo 0
                                    targetRef = "'" & hubName & "'!" & rangePart
                                Else
                                    ' CASO CELLA SINGOLA
                                    If Not refMap.Exists(cleanKey) Then
                                        ' Nuova voce: Compila Colonna A, B e C
                                        wsHub.Cells(currRow, 1).Value = cleanKey
                                        wsHub.Cells(currRow, 2).formula = "=" & cleanKey
                                        wsHub.Cells(currRow, 3).Value = destCellAddr
                                        
                                        targetRef = "'" & hubName & "'!" & wsHub.Cells(currRow, 2).Address(True, True)
                                        ' Salva nel dizionario sia il numero di riga che il riferimento Target
                                        refMap(cleanKey) = currRow & "|" & targetRef
                                        currRow = currRow + 1
                                    Else
                                        ' Voce già presente: Recupera riga e TargetRef
                                        mapData = Split(refMap(cleanKey), "|")
                                        existingRow = CLng(mapData(0))
                                        targetRef = CStr(mapData(1))
                                        
                                        ' Aggiunge la nuova destinazione in Colonna C se non già contenuta
                                        If InStr(wsHub.Cells(existingRow, 3).Value, destCellAddr) = 0 Then
                                            wsHub.Cells(existingRow, 3).Value = wsHub.Cells(existingRow, 3).Value & "; " & destCellAddr
                                        End If
                                    End If
                                End If
                                
                                ' Prepara la stringa di formula aggiornata
                                fml = Replace(fml, origRef, targetRef)
                            Next match
                            
                            ' --------------------------------------------------------------
                            ' EVENTUALE SCRITTURA NELLA CELLA SORGENTE
                            ' --------------------------------------------------------------
                            If doUpdateFormulas Then
                                On Error Resume Next
                                cell.formula = fml
                                If Err.Number = 0 Then
                                    countModified = countModified + 1
                                Else
                                    Err.Clear
                                End If
                                On Error GoTo 0
                            End If
                        End If
                    End If
NextCellInLoop:
                Next cell
            End If
        End If
    Next sheetNameKey
End Sub

' ======================================================================================
' FUNZIONE AUSILIARIA: EscapeRegExPattern
' SCOPO: Effettua l'escape dei caratteri riservati RegEx nei nomi dei fogli
' ======================================================================================
Private Function EscapeRegExPattern(ByVal str As String) As String
    Dim charsToEscape As Variant
    charsToEscape = Array("\", "^", "$", ".", "|", "?", "*", "+", "(", ")", "[", "]", "{", "}")
    Dim i As Long
    For i = LBound(charsToEscape) To UBound(charsToEscape)
        str = Replace(str, charsToEscape(i), "\" & charsToEscape(i))
    Next i
    EscapeRegExPattern = str
End Function

