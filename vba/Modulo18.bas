Attribute VB_Name = "Modulo18"
Option Explicit

' --- Variabili di modulo: mantengono lo stato tra i lanci delle due macro ---
Private gPrimariaMemorizzata As Boolean
Private gPrimariaRange As Range
Private gPrimariaValori() As Variant


Sub m18_SELEZIONE1()

    Dim selCorrente As Range
    Dim risposta As VbMsgBoxResult

    If TypeName(Selection) <> "Range" Then
        MsgBox "Seleziona prima un intervallo di celle.", vbExclamation, "Nessuna selezione valida"
        Exit Sub
    End If

    Set selCorrente = Selection

    risposta = MsgBox( _
        "Selezione rilevata:" & vbCrLf & _
        selCorrente.Parent.Name & "!" & selCorrente.Address(False, False) & _
        " (" & selCorrente.Cells.count & " celle)" & vbCrLf & vbCrLf & _
        "SI = Procedi (memorizza come selezione primaria)" & vbCrLf & _
        "NO = Dimentica", _
        vbYesNo + vbQuestion, "SELEZIONE 1")

    If risposta = vbYes Then
        m18_MemorizzaPrimaria selCorrente
        MsgBox "Analizzata. Ora vai sul secondo foglio, seleziona le nuove celle e lancia m18_SELEZIONE2.", _
               vbInformation, "Selezione primaria memorizzata"
    Else
        gPrimariaMemorizzata = False
        Set gPrimariaRange = Nothing
        Erase gPrimariaValori
        MsgBox "Selezione dimenticata. Nessun dato memorizzato.", vbInformation, "Operazione annullata"
    End If

End Sub


Sub m18_SELEZIONE2()

    Dim selCorrente As Range
    Dim risposta As VbMsgBoxResult

    If Not gPrimariaMemorizzata Then
        MsgBox "Non c'è nessuna selezione primaria memorizzata." & vbCrLf & _
               "Lancia prima m18_SELEZIONE1.", vbExclamation, "Selezione primaria mancante"
        Exit Sub
    End If

    If TypeName(Selection) <> "Range" Then
        MsgBox "Seleziona prima un intervallo di celle.", vbExclamation, "Nessuna selezione valida"
        Exit Sub
    End If

    Set selCorrente = Selection

    risposta = MsgBox( _
        "Selezione primaria memorizzata:" & vbCrLf & _
        gPrimariaRange.Parent.Name & "!" & gPrimariaRange.Address(False, False) & _
        " (" & gPrimariaRange.Cells.count & " celle)" & vbCrLf & vbCrLf & _
        "Nuova selezione:" & vbCrLf & _
        selCorrente.Parent.Name & "!" & selCorrente.Address(False, False) & _
        " (" & selCorrente.Cells.count & " celle)" & vbCrLf & vbCrLf & _
        "SI = Confronta" & vbCrLf & _
        "NO = Dimentica (annulla entrambe le selezioni)", _
        vbYesNo + vbQuestion, "SELEZIONE 2")

    If risposta = vbNo Then
        gPrimariaMemorizzata = False
        Set gPrimariaRange = Nothing
        Erase gPrimariaValori
        MsgBox "Entrambe le selezioni sono state dimenticate.", vbInformation, "Operazione annullata"
        Exit Sub
    End If

    ' --- CONFRONTA ---
    m18_Confronta selCorrente

    ' reset dopo il confronto: pronto per un nuovo ciclo
    gPrimariaMemorizzata = False
    Set gPrimariaRange = Nothing
    Erase gPrimariaValori

End Sub


Private Sub m18_MemorizzaPrimaria(rng As Range)
    Dim cel As Range
    Dim n As Long

    Set gPrimariaRange = rng
    ReDim gPrimariaValori(1 To rng.Cells.count)

    n = 0
    For Each cel In rng.Cells
        n = n + 1
        gPrimariaValori(n) = cel.Value2   ' Value2: valore "grezzo", da formula o valore diretto
    Next cel

    gPrimariaMemorizzata = True
End Sub


Private Sub m18_Confronta(rngSecondaria As Range)

    Dim n1 As Long, n2 As Long, i As Long
    Dim cel As Range
    Dim valoriSecondaria() As Variant
    Dim diff As String
    Dim numDiff As Long
    Dim indiciDiff() As Long
    Dim risposta As VbMsgBoxResult
    Dim msg As String

    n1 = gPrimariaRange.Cells.count
    n2 = rngSecondaria.Cells.count

    If n1 <> n2 Then
        MsgBox "Le selezioni hanno dimensioni diverse, impossibile confrontare:" & vbCrLf & vbCrLf & _
               "Primaria (" & gPrimariaRange.Parent.Name & "!" & gPrimariaRange.Address(False, False) & "): " & n1 & " celle" & vbCrLf & _
               "Secondaria (" & rngSecondaria.Parent.Name & "!" & rngSecondaria.Address(False, False) & "): " & n2 & " celle", _
               vbExclamation, "Dimensioni diverse"
        Exit Sub
    End If

    ReDim valoriSecondaria(1 To n2)
    i = 0
    For Each cel In rngSecondaria.Cells
        i = i + 1
        valoriSecondaria(i) = cel.Value2
    Next cel

    ReDim indiciDiff(1 To n1)
    diff = ""
    numDiff = 0
    For i = 1 To n1
        If Not m18_ValoriEquivalenti(gPrimariaValori(i), valoriSecondaria(i)) Then
            numDiff = numDiff + 1
            indiciDiff(numDiff) = i
            If diff <> "" Then diff = diff & ", "
            diff = diff & gPrimariaRange.Cells(i).Address(False, False)
        End If
    Next i

    ' --- Nessuna differenza: VERIFICATO ---
    If numDiff = 0 Then
        MsgBox "VERIFICATO" & vbCrLf & vbCrLf & _
               "Le due selezioni contengono gli stessi valori equivalenti." & vbCrLf & vbCrLf & _
               gPrimariaRange.Parent.Name & "!" & gPrimariaRange.Address(False, False) & "   =   " & _
               rngSecondaria.Parent.Name & "!" & rngSecondaria.Address(False, False), _
               vbInformation, "Selezioni corrispondenti"
        Exit Sub
    End If

    ' --- Costruisco il messaggio: elenco se poche, solo conteggio se tante ---
    If numDiff = n1 Then
        msg = "Nessuna corrispondenza: TUTTE le celle della selezione primaria differiscono dalla secondaria."
    ElseIf numDiff <= 5 Then
        msg = "Trovate differenze in " & numDiff & " cella/e su " & n1 & "." & vbCrLf & vbCrLf & _
              "Celle della selezione primaria (" & gPrimariaRange.Parent.Name & "!" & gPrimariaRange.Address(False, False) & _
              ") che NON coincidono con la secondaria:" & vbCrLf & diff
    Else
        msg = "Trovate differenze in " & numDiff & " cella/e su " & n1 & "."
    End If

    ' --- DIVERSE: icona rossa ---
    MsgBox "DIVERSE" & vbCrLf & vbCrLf & msg, vbCritical, "Differenze trovate"

    ' --- Chiedo se evidenziare in giallo ---
    risposta = MsgBox( _
        "Vuoi evidenziare in giallo le celle diverse nella selezione 2 (" & _
        rngSecondaria.Parent.Name & "!" & rngSecondaria.Address(False, False) & ")?", _
        vbYesNo + vbQuestion, "Evidenziazione")

    If risposta = vbYes Then
        rngSecondaria.Interior.ColorIndex = xlColorIndexNone   ' pulisce eventuale colore precedente
        For i = 1 To numDiff
            rngSecondaria.Cells(indiciDiff(i)).Interior.Color = RGB(255, 255, 0)
        Next i
        MsgBox numDiff & " cella/e evidenziate in giallo sulla selezione 2.", vbInformation, "Evidenziazione completata"
    End If

End Sub









Private Function m18_ValoriEquivalenti(v1 As Variant, v2 As Variant) As Boolean
    Const TOLL As Double = 0.000001

    If IsEmpty(v1) And IsEmpty(v2) Then
        m18_ValoriEquivalenti = True
        Exit Function
    End If

    If IsEmpty(v1) Or IsEmpty(v2) Then
        m18_ValoriEquivalenti = False
        Exit Function
    End If

    If IsNumeric(v1) And IsNumeric(v2) Then
        m18_ValoriEquivalenti = (Abs(CDbl(v1) - CDbl(v2)) < TOLL)
    Else
        m18_ValoriEquivalenti = (CStr(v1) = CStr(v2))
    End If
End Function

