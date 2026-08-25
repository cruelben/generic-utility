Attribute VB_Name = "Modulo5"
Option Explicit

Sub m5_CollegamentiEsterni()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim link As Variant
    Dim msg As String
    Dim risultati As String
    Dim linkFound As Boolean
    Dim i As Integer
    Dim fileNum As Integer
    Dim percorsoFile As String
    Dim celleUniche As Object
    Dim fileEsterni As Object
    Dim chiaveCella As String
    Dim nomeFileEsterno As String
    Dim posInizio As Long
    Dim posFine As Long
    Dim risposta As VbMsgBoxResult
    Dim k As Variant
    
    ' Dictionary per tenere traccia delle celle uniche e dei file esterni distinti
    Set celleUniche = CreateObject("Scripting.Dictionary")
    Set fileEsterni = CreateObject("Scripting.Dictionary")
    
    ' Inizializza il messaggio
    msg = "Collegamenti esterni trovati:" & vbCrLf
    risultati = "Collegamenti esterni trovati:" & vbCrLf
    i = 1 ' Inizializza il contatore
    
    ' Scansiona tutti i fogli di lavoro
    For Each ws In ThisWorkbook.Worksheets
        ' Scansiona tutte le celle utilizzate nel foglio di lavoro
        For Each rng In ws.UsedRange
            ' Verifica se la cella contiene una formula
            If rng.HasFormula Then
                ' Scansiona tutte le celle non vuote nella selezione
                For Each cell In rng.Cells
                    ' Verifica se la cella contiene un collegamento esterno
                    For Each link In Split(cell.formula, ",")
                        If InStr(link, "[") > 0 And InStr(link, "]") > 0 Then
                            ' Aggiungi il collegamento al messaggio con il numero sequenziale
                            msg = msg & i & ". Collegamento trovato nel foglio: " & ws.Name & ", cella: " & cell.Address & ", collegamento: " & link & vbCrLf
                            risultati = risultati & i & ". Collegamento trovato nel foglio: " & ws.Name & ", cella: " & cell.Address & ", collegamento: " & link & vbCrLf
                            linkFound = True
                            i = i + 1 ' Incrementa il contatore
                            
                            ' Traccia la cella come unica (foglio + indirizzo)
                            chiaveCella = ws.Name & "|" & cell.Address
                            If Not celleUniche.Exists(chiaveCella) Then
                                celleUniche.Add chiaveCella, cell
                            End If
                            
                            ' Estrae il nome del file esterno tra [ e ] per contare i file distinti
                            posInizio = InStr(link, "[")
                            posFine = InStr(link, "]")
                            If posInizio > 0 And posFine > posInizio Then
                                nomeFileEsterno = Mid(link, posInizio + 1, posFine - posInizio - 1)
                                If Not fileEsterni.Exists(nomeFileEsterno) Then
                                    fileEsterni.Add nomeFileEsterno, 1
                                End If
                            End If
                        End If
                    Next link
                Next cell
            End If
        Next rng
    Next ws
    
    ' Gestione risultato
    If Not linkFound Then
        MsgBox "Nessun collegamento a fogli esterni trovato.", vbInformation, "Collegamenti esterni"
    Else
        ' Messaggio riepilogativo con richiesta Si/No
        risposta = MsgBox("Trovati " & celleUniche.count & " collegamenti a " & fileEsterni.count & " file esterni." & vbCrLf & vbCrLf & _
                           "Vuoi che trasformi queste celle in valori?", vbYesNo + vbQuestion, "Collegamenti esterni")
        
        ' Se l'utente sceglie Si, converte le celle trovate in valori
        If risposta = vbYes Then
            For Each k In celleUniche.Keys
                celleUniche(k).Value = celleUniche(k).Value
            Next k
            
            ' Aggiorna il semaforo sullo UserForm, se aperto
            Call m5_AggiornaSemaforoCollegamenti
        End If
        
        ' Crea il file TXT (sempre, se ci sono risultati)
        percorsoFile = ThisWorkbook.path & "\CollegamentiEsterni.txt"
        fileNum = FreeFile
        Open percorsoFile For Output As #fileNum
        Print #fileNum, risultati
        Close #fileNum
        
        ' Apre il file con il Blocco Note
        Shell "notepad.exe " & Chr(34) & percorsoFile & Chr(34), vbNormalFocus
    End If
End Sub

' Restituisce True se nel workbook sono presenti collegamenti esterni
Public Function m5_EsistonoCollegamentiEsterni() As Boolean
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim link As Variant
    
    For Each ws In ThisWorkbook.Worksheets
        For Each rng In ws.UsedRange
            If rng.HasFormula Then
                For Each cell In rng.Cells
                    For Each link In Split(cell.formula, ",")
                        If InStr(link, "[") > 0 And InStr(link, "]") > 0 Then
                            m5_EsistonoCollegamentiEsterni = True
                            Exit Function
                        End If
                    Next link
                Next cell
            End If
        Next rng
    Next ws
    
    m5_EsistonoCollegamentiEsterni = False
End Function

' Aggiorna il semaforo (Label31) sullo UserForm1UTILITY, se aperto
Public Sub m5_AggiornaSemaforoCollegamenti()
    Dim frm As Object
    Dim trovato As Boolean
    
    trovato = m5_EsistonoCollegamentiEsterni()
    
    For Each frm In VBA.UserForms
        If frm.Name = "UserForm1UTILITY" Then
            frm.Label31.BackColor = IIf(trovato, RGB(255, 0, 0), RGB(198, 239, 206))
            Exit For
        End If
    Next frm
End Sub

