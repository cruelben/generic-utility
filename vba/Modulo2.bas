Attribute VB_Name = "Modulo2"
Option Explicit

Private m2_mFlashRange As Range
Private m2_mFlashColorIndexOriginale As Variant
Private m2_mFlashColorOriginale As Long
Private m2_mFlashPatternOriginale As XlPattern
Private m2_mFlashAvevaColore As Boolean
Private m2_mFlashOraFine As Date

Sub m2_VaiAllaCellaOrigine()
Attribute m2_VaiAllaCellaOrigine.VB_ProcData.VB_Invoke_Func = "y\n14"
    Dim f As String
    Dim nomeFoglio As String
    Dim indirizzo As String
    Dim p1 As Long, p2 As Long, pEscl As Long
    Dim foglioOrigine As Worksheet

    If Not ActiveCell.HasFormula Then
        MsgBox "La cella selezionata non contiene una formula.", vbExclamation
        Exit Sub
    End If

    Set foglioOrigine = ActiveSheet
    f = ActiveCell.formula

    If Left(f, 1) = "=" Then f = Mid(f, 2)

    Do While Len(f) > 0 And (Left(f, 1) = "+" Or Left(f, 1) = "-")
        f = Mid(f, 2)
    Loop

    If UCase(Left(f, 13)) = "CERCA.ORIZZ(" Or UCase(Left(f, 8)) = "HLOOKUP(" Then
        Call m2_GestiCercaOrizz(f, foglioOrigine)
        Exit Sub
    End If

    If UCase(Left(f, 11)) = "CERCA.VERT(" Or UCase(Left(f, 8)) = "VLOOKUP(" Then
        Call m2_GestiCercaVert(f, foglioOrigine)
        Exit Sub
    End If

    pEscl = InStr(f, "!")

    If pEscl > 0 Then
        If Left(f, 1) = "'" Then
            p1 = 2
            p2 = InStr(p1, f, "'")
            nomeFoglio = Mid(f, p1, p2 - p1)
        Else
            nomeFoglio = Left(f, pEscl - 1)
        End If

        indirizzo = Mid(f, pEscl + 1)
        indirizzo = Trim(m2_PrimoRiferimento(indirizzo))
        indirizzo = Replace(indirizzo, "$", "")

        If indirizzo = "" Then
            MsgBox "Non riesco a individuare un indirizzo di cella valido nella formula.", vbExclamation
            Exit Sub
        End If

        On Error GoTo Errore
        Worksheets(nomeFoglio).Activate
        Range(indirizzo).Select
        Call m2_EvidenziaTemporanea(Range(indirizzo))
        Exit Sub
    End If

    indirizzo = Trim(m2_PrimoRiferimento(f))
    indirizzo = Replace(indirizzo, "$", "")

    If indirizzo = "" Then
        MsgBox "Non trovo un riferimento di cella interpretabile nella formula.", vbExclamation
        Exit Sub
    End If

    On Error GoTo Errore
    foglioOrigine.Activate
    Range(indirizzo).Select
    Call m2_EvidenziaTemporanea(Range(indirizzo))
    Exit Sub

Errore:
    MsgBox "Non riesco a trovare il foglio o la cella indicata." & vbCrLf & _
           "Dettaglio: " & Err.Description, vbCritical
End Sub

Private Sub m2_GestiCercaOrizz(f As String, foglioOrigine As Worksheet)
    Dim args() As String
    Dim valoreRicerca As Variant
    Dim matriceRange As Range
    Dim indiceRiga As Long
    Dim j As Long
    Dim corrispondenzaApprox As Boolean

    args = m2_EstraiArgomenti(f)

    If UBound(args) < 2 Then
        MsgBox "Impossibile interpretare gli argomenti di CERCA.ORIZZ.", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErroreRange
    Set matriceRange = m2_RisolviRange(Trim(args(1)), foglioOrigine)
    If matriceRange Is Nothing Then GoTo ErroreRange

    indiceRiga = CLng(Trim(args(2)))
    valoreRicerca = m2_RisolviValore(Trim(args(0)), foglioOrigine)

    corrispondenzaApprox = m2_LeggiFlagCorrispondenza(args, 3)

    Dim colonnaFound As Long
    colonnaFound = 0

    If corrispondenzaApprox Then
        For j = 1 To matriceRange.Columns.count
            If m2_ConfrontaValori(matriceRange.Cells(1, j).Value, valoreRicerca) <= 0 Then
                colonnaFound = j
            Else
                Exit For
            End If
        Next j
    Else
        For j = 1 To matriceRange.Columns.count
            If m2_ConfrontaValori(matriceRange.Cells(1, j).Value, valoreRicerca) = 0 Then
                colonnaFound = j
                Exit For
            End If
        Next j
    End If

    If colonnaFound = 0 Then
        MsgBox "Valore '" & valoreRicerca & "' non trovato nella prima riga della matrice.", vbExclamation
        Exit Sub
    End If

    Dim cellaDestino As Range
    Set cellaDestino = matriceRange.Cells(indiceRiga, colonnaFound)
    cellaDestino.Worksheet.Activate
    cellaDestino.Select
    Call m2_EvidenziaTemporanea(cellaDestino)
    Exit Sub

ErroreRange:
    MsgBox "Non riesco a risolvere la matrice di CERCA.ORIZZ." & vbCrLf & _
           "Dettaglio: " & Err.Description, vbCritical
End Sub

Private Sub m2_GestiCercaVert(f As String, foglioOrigine As Worksheet)
    Dim args() As String
    Dim valoreRicerca As Variant
    Dim matriceRange As Range
    Dim indiceColonna As Long
    Dim i As Long
    Dim corrispondenzaApprox As Boolean

    args = m2_EstraiArgomenti(f)

    If UBound(args) < 2 Then
        MsgBox "Impossibile interpretare gli argomenti di CERCA.VERT.", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErroreRange
    Set matriceRange = m2_RisolviRange(Trim(args(1)), foglioOrigine)
    If matriceRange Is Nothing Then GoTo ErroreRange

    indiceColonna = CLng(Trim(args(2)))
    valoreRicerca = m2_RisolviValore(Trim(args(0)), foglioOrigine)

    corrispondenzaApprox = m2_LeggiFlagCorrispondenza(args, 3)

    Dim rigaFound As Long
    rigaFound = 0

    If corrispondenzaApprox Then
        For i = 1 To matriceRange.Rows.count
            If m2_ConfrontaValori(matriceRange.Cells(i, 1).Value, valoreRicerca) <= 0 Then
                rigaFound = i
            Else
                Exit For
            End If
        Next i
    Else
        For i = 1 To matriceRange.Rows.count
            If m2_ConfrontaValori(matriceRange.Cells(i, 1).Value, valoreRicerca) = 0 Then
                rigaFound = i
                Exit For
            End If
        Next i
    End If

    If rigaFound = 0 Then
        MsgBox "Valore '" & valoreRicerca & "' non trovato nella prima colonna della matrice.", vbExclamation
        Exit Sub
    End If

    Dim cellaDestino As Range
    Set cellaDestino = matriceRange.Cells(rigaFound, indiceColonna)
    cellaDestino.Worksheet.Activate
    cellaDestino.Select
    Call m2_EvidenziaTemporanea(cellaDestino)
    Exit Sub

ErroreRange:
    MsgBox "Non riesco a risolvere la matrice di CERCA.VERT." & vbCrLf & _
           "Dettaglio: " & Err.Description, vbCritical
End Sub

Private Function m2_LeggiFlagCorrispondenza(args() As String, idx As Long) As Boolean
    If UBound(args) < idx Then
        m2_LeggiFlagCorrispondenza = True
        Exit Function
    End If

    Dim sArg As String
    sArg = UCase(Trim(args(idx)))

    Select Case sArg
        Case "FALSO", "FALSE", "0"
            m2_LeggiFlagCorrispondenza = False
        Case Else
            m2_LeggiFlagCorrispondenza = True
    End Select
End Function

Private Function m2_ConfrontaValori(a As Variant, b As Variant) As Long
    If IsNumeric(a) And IsNumeric(b) Then
        If CDbl(a) < CDbl(b) Then
            m2_ConfrontaValori = -1
        ElseIf CDbl(a) > CDbl(b) Then
            m2_ConfrontaValori = 1
        Else
            m2_ConfrontaValori = 0
        End If
    Else
        m2_ConfrontaValori = StrComp(CStr(a), CStr(b), vbTextCompare)
    End If
End Function

Private Function m2_EstraiArgomenti(f As String) As String()
    Dim pOpen As Long, pClose As Long
    pOpen = InStr(f, "(")
    pClose = InStrRev(f, ")")

    If pOpen = 0 Or pClose = 0 Then
        m2_EstraiArgomenti = Split("", ";")
        Exit Function
    End If

    Dim inner As String
    inner = Mid(f, pOpen + 1, pClose - pOpen - 1)

    Dim risultato() As String
    ReDim risultato(0)
    Dim livello As Long
    livello = 0
    Dim dentroTesto As Boolean
    dentroTesto = False
    Dim corrente As String
    corrente = ""
    Dim k As Long
    Dim ch As String

    For k = 1 To Len(inner)
        ch = Mid(inner, k, 1)

        If ch = """" Then
            dentroTesto = Not dentroTesto
            corrente = corrente & ch
        ElseIf ch = "(" And Not dentroTesto Then
            livello = livello + 1
            corrente = corrente & ch
        ElseIf ch = ")" And Not dentroTesto Then
            livello = livello - 1
            corrente = corrente & ch
        ElseIf (ch = ";" Or ch = ",") And livello = 0 And Not dentroTesto Then
            ReDim Preserve risultato(UBound(risultato) + 1)
            risultato(UBound(risultato) - 1) = corrente
            corrente = ""
        Else
            corrente = corrente & ch
        End If
    Next k

    risultato(UBound(risultato)) = corrente
    m2_EstraiArgomenti = risultato
End Function

Private Function m2_RisolviRange(ref As String, wsDefault As Worksheet) As Range
    Dim pEscl As Long
    pEscl = InStr(ref, "!")

    On Error Resume Next
    If pEscl > 0 Then
        Dim nFoglio As String, nRange As String
        If Left(ref, 1) = "'" Then
            nFoglio = Mid(ref, 2, InStr(2, ref, "'") - 2)
            nRange = Mid(ref, InStr(2, ref, "'") + 2)
        Else
            nFoglio = Left(ref, pEscl - 1)
            nRange = Mid(ref, pEscl + 1)
        End If
        nRange = Replace(nRange, "$", "")
        Set m2_RisolviRange = Worksheets(nFoglio).Range(nRange)
    Else
        Dim refPulito As String
        refPulito = Replace(ref, "$", "")
        Set m2_RisolviRange = wsDefault.Range(refPulito)
    End If
    On Error GoTo 0

    If Not m2_RisolviRange Is Nothing Then Exit Function

    On Error Resume Next
    Err.Clear
    Set m2_RisolviRange = wsDefault.Names(ref).RefersToRange
    If m2_RisolviRange Is Nothing Then
        Err.Clear
        Set m2_RisolviRange = wsDefault.Parent.Names(ref).RefersToRange
    End If
    On Error GoTo 0
End Function

Private Function m2_RisolviValore(arg As String, wsDefault As Worksheet) As Variant
    Dim r As Range
    Set r = m2_RisolviRange(arg, wsDefault)
    If Not r Is Nothing Then
        m2_RisolviValore = r.Cells(1, 1).Value
    Else
        m2_RisolviValore = Replace(arg, """", "")
    End If
End Function

Private Sub m2_EvidenziaTemporanea(rng As Range, Optional durataSecondi As Double = 1.5)
    On Error Resume Next

    If Not m2_mFlashRange Is Nothing Then
        If m2_mFlashAvevaColore Then
            m2_mFlashRange.Interior.Pattern = m2_mFlashPatternOriginale
            m2_mFlashRange.Interior.Color = m2_mFlashColorOriginale
        Else
            m2_mFlashRange.Interior.ColorIndex = xlColorIndexNone
        End If

        If m2_mFlashOraFine <> 0 Then
            Application.OnTime EarliestTime:=m2_mFlashOraFine, _
                               Procedure:="m2_SpegniFlashEvidenziaTemporanea", _
                               Schedule:=False
        End If

        Set m2_mFlashRange = Nothing
    End If

    Set m2_mFlashRange = rng
    m2_mFlashColorIndexOriginale = rng.Interior.ColorIndex
    m2_mFlashColorOriginale = rng.Interior.Color
    m2_mFlashPatternOriginale = rng.Interior.Pattern
    m2_mFlashAvevaColore = (m2_mFlashColorIndexOriginale <> xlColorIndexNone)

    rng.Interior.Pattern = xlSolid
    rng.Interior.Color = RGB(255, 165, 0)

    Application.ScreenUpdating = True
    DoEvents

    m2_mFlashOraFine = Now + (durataSecondi / 86400#)
    Application.OnTime EarliestTime:=m2_mFlashOraFine, _
                       Procedure:="m2_SpegniFlashEvidenziaTemporanea", _
                       Schedule:=True

    On Error GoTo 0
End Sub

Public Sub m2_SpegniFlashEvidenziaTemporanea()
    On Error Resume Next

    If Not m2_mFlashRange Is Nothing Then
        If m2_mFlashAvevaColore Then
            m2_mFlashRange.Interior.Pattern = m2_mFlashPatternOriginale
            m2_mFlashRange.Interior.Color = m2_mFlashColorOriginale
        Else
            m2_mFlashRange.Interior.ColorIndex = xlColorIndexNone
        End If
    End If

    Set m2_mFlashRange = Nothing
    m2_mFlashOraFine = 0
    On Error GoTo 0
End Sub

Private Function m2_PrimoRiferimento(s As String) As String
    Dim i As Long
    Dim c As String

    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        Select Case c
            Case "A" To "Z", "a" To "z", "0" To "9", "$", ".", ":"
            Case Else
                m2_PrimoRiferimento = Left(s, i - 1)
                Exit Function
        End Select
    Next i

    m2_PrimoRiferimento = s
End Function

