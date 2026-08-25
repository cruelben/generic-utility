Attribute VB_Name = "Modulo13"
Sub m13_ImportaFoglioRicollegaFormule()
    Dim ws      As Worksheet
    Dim cell    As Range
    Dim usedRng As Range

    Set ws = ActiveSheet

    On Error Resume Next
    Set usedRng = ws.UsedRange
    On Error GoTo 0

    If usedRng Is Nothing Then
        MsgBox "Il foglio attivo e' vuoto.", vbInformation
        Exit Sub
    End If

    Dim mapping As Object
    Set mapping = CreateObject("Scripting.Dictionary")

    Dim totalConverted As Long
    Dim totalToValue   As Long
    Dim totalSkipped   As Long

    ' ===================================================
    ' PASSATA 1: raccoglie tutti i nomi foglio esterni
    ' ===================================================
    Dim extNames As New Collection
    Dim nm       As Variant
    Dim found    As Boolean

    For Each cell In usedRng.Cells
        If cell.HasFormula Then
            Dim extList() As String
            extList = GetExternalSheetNames(cell.formula)
            Dim i As Integer
            For i = 0 To UBound(extList)
                If extList(i) <> "" Then
                    found = False
                    For Each nm In extNames
                        If nm = extList(i) Then found = True: Exit For
                    Next nm
                    If Not found Then extNames.Add extList(i)
                End If
            Next i
        End If
    Next cell

    If extNames.count = 0 Then
        MsgBox "Nessun collegamento esterno trovato nel foglio attivo.", vbInformation
        Exit Sub
    End If

    ' ===================================================
    ' Per ogni foglio esterno: decidi l'azione
    ' ===================================================
    Dim shName      As String
    Dim localWs     As Worksheet
    Dim existsLocal As Boolean
    Dim matchedName As String

    For Each nm In extNames
        shName = CStr(nm)
        If Not mapping.Exists(shName) Then

            existsLocal = False
            matchedName = ""
            For Each localWs In ThisWorkbook.Worksheets
                If NormalizeSheetName(localWs.Name) = NormalizeSheetName(shName) Then
                    existsLocal = True
                    matchedName = localWs.Name
                    Exit For
                End If
            Next localWs

            If existsLocal Then
                mapping.Add shName, matchedName
            Else
                Dim msgTxt As String
                msgTxt = "Il foglio esterno '" & shName & "' NON esiste nel file corrente." & vbCrLf & vbCrLf & _
                         "Fogli disponibili:" & vbCrLf
                Dim wsT As Worksheet
                For Each wsT In ThisWorkbook.Worksheets
                    msgTxt = msgTxt & "  - " & wsT.Name & vbCrLf
                Next wsT
                msgTxt = msgTxt & vbCrLf & _
                         "Digita il nome del foglio locale corretto," & vbCrLf & _
                         "oppure lascia VUOTO per convertire in valore statico."

                Dim answer As String
                answer = InputBox(msgTxt, "Foglio non trovato: " & shName, "")

                If answer = "" Then
                    mapping.Add shName, ""
                Else
                    Dim validName  As Boolean
                    Dim resolvedName As String
                    validName = False
                    resolvedName = ""
                    For Each wsT In ThisWorkbook.Worksheets
                        If NormalizeSheetName(wsT.Name) = NormalizeSheetName(answer) Then
                            validName = True
                            resolvedName = wsT.Name
                            Exit For
                        End If
                    Next wsT

                    If validName Then
                        mapping.Add shName, resolvedName
                    Else
                        Dim resp As Integer
                        resp = MsgBox("Il foglio '" & answer & "' non esiste." & vbCrLf & _
                                      "Convertire le celle in valore statico?", _
                                      vbYesNo + vbExclamation, "Foglio non valido")
                        mapping.Add shName, IIf(resp = vbYes, "", "SKIP")
                    End If
                End If
            End If
        End If
    Next nm

    ' ===================================================
    ' PASSATA 2: sostituisce le formule
    ' ===================================================
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For Each cell In usedRng.Cells
        If cell.HasFormula Then
            Dim origFormula As String
            origFormula = cell.formula

            If HasExternalRef(origFormula) Then
                Dim newFormula As String
                Dim needsValue As Boolean
                Dim skipCell   As Boolean
                needsValue = False
                skipCell = False

                newFormula = ReplaceExternalRefs(origFormula, mapping, needsValue, skipCell)

                If skipCell Then
                    totalSkipped = totalSkipped + 1
                ElseIf needsValue Then
                    Dim cellVal As Variant
                    cellVal = cell.Value
                    cell.ClearContents
                    cell.Value = cellVal
                    totalToValue = totalToValue + 1
                ElseIf newFormula <> origFormula Then
                    On Error Resume Next
                    cell.formula = newFormula
                    If Err.Number <> 0 Then
                        Err.Clear
                        cell.FormulaLocal = newFormula
                        If Err.Number <> 0 Then
                            Err.Clear
                            totalSkipped = totalSkipped + 1
                        Else
                            totalConverted = totalConverted + 1
                        End If
                    Else
                        totalConverted = totalConverted + 1
                    End If
                    On Error GoTo 0
                End If
            End If
        End If
    Next cell

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ws.Calculate


    ' Aggiorna il semaforo sullo UserForm, se aperto
    Call AggiornaSemaforoCollegamenti


    MsgBox "Operazione completata su: " & ws.Name & vbCrLf & vbCrLf & _
           "  Formule ricollegare localmente : " & totalConverted & vbCrLf & _
           "  Celle convertite in valore      : " & totalToValue & vbCrLf & _
           "  Celle saltate / errori          : " & totalSkipped, _
           vbInformation, "m13_ImportaFoglioRicollegaFormule"
End Sub

' =============================================================
' Normalizza il nome foglio per confronto robusto
' (gestisce varianti codifica del simbolo euro e spazi speciali)
' =============================================================
Private Function NormalizeSheetName(s As String) As String
    Dim result As String
    result = s
    result = Replace(result, Chr(160), " ")          ' no-break space (ANSI) -> spazio
    result = Replace(result, ChrW(8201), " ")        ' thin space (Unicode) -> spazio
    result = Replace(result, Chr(128), ChrW(8364))   ' euro Windows-1252 -> euro Unicode
    result = LCase(Trim(result))
    NormalizeSheetName = result
End Function

' =============================================================
' Estrae tutti i nomi foglio esterni da una formula
' =============================================================
Private Function GetExternalSheetNames(formula As String) As String()
    Dim results() As String
    Dim count     As Integer
    count = 0
    ReDim results(0)
    results(0) = ""

    Dim pos As Integer
    pos = 1

    Do
        Dim bracket As Integer
        bracket = InStr(pos, formula, "]")
        If bracket = 0 Then Exit Do

        Dim excl As Integer
        excl = InStr(bracket + 1, formula, "!")
        If excl = 0 Then Exit Do

        Dim apostPost As Integer
        apostPost = InStr(bracket + 1, formula, "'")

        Dim endPos As Integer
        If apostPost > 0 And apostPost < excl Then
            endPos = apostPost - 1
        Else
            endPos = excl - 1
        End If

        Dim shName As String
        shName = Mid(formula, bracket + 1, endPos - bracket)
        shName = Replace(shName, "'", "")
        shName = Trim(shName)

        If shName <> "" Then
            ReDim Preserve results(count)
            results(count) = shName
            count = count + 1
        End If

        pos = excl + 1
    Loop

    GetExternalSheetNames = results
End Function

' =============================================================
Private Function HasExternalRef(formula As String) As Boolean
    HasExternalRef = (InStr(formula, "[") > 0 And InStr(formula, "]") > 0)
End Function

Private Function ReplaceExternalRefs(formula As String, _
                                     mapping As Object, _
                                     ByRef needsValue As Boolean, _
                                     ByRef skipCell As Boolean) As String
    Dim result As String
    result = formula
    needsValue = False
    skipCell = False

    Dim safetyCounter As Integer
    safetyCounter = 0

    Do While InStr(1, result, "[") > 0
        safetyCounter = safetyCounter + 1
        If safetyCounter > 500 Then Exit Do

        Dim bracketOpen  As Integer
        Dim bracketClose As Integer
        bracketOpen = InStr(1, result, "[")
        If bracketOpen = 0 Then Exit Do
        bracketClose = InStr(bracketOpen, result, "]")
        If bracketClose = 0 Then Exit Do

        Dim excl As Integer
        excl = InStr(bracketClose + 1, result, "!")
        If excl = 0 Then
            result = Left(result, bracketOpen - 1) & Chr(1) & Mid(result, bracketOpen + 1)
            GoTo NextIter
        End If

        Dim apostPost As Integer
        apostPost = InStr(bracketClose + 1, result, "'")

        Dim blockEnd As Integer
        If apostPost > 0 And apostPost < excl Then
            blockEnd = apostPost
        Else
            blockEnd = excl - 1
        End If

        ' ===================================================
        ' Determina blockStart: cerca l'apice di apertura reale
        ' con InStrRev invece di uno scan carattere-per-carattere
        ' che si confonde con trattini/operatori dentro il percorso
        ' (es. "07 - ENERGIA..."). Fallback a scan a operatori
        ' solo se il riferimento NON e' tra apici.
        ' ===================================================
        Dim blockStart  As Integer
        Dim aposBefore  As Integer
        Dim isQuoted    As Boolean

        isQuoted = (apostPost > 0 And apostPost < excl)

        If isQuoted Then
            aposBefore = InStrRev(result, "'", bracketOpen - 1)
            If aposBefore > 0 Then
                blockStart = aposBefore
            Else
                blockStart = bracketOpen ' fallback di sicurezza, non dovrebbe accadere
            End If
        Else
            blockStart = bracketOpen
            Dim k As Integer
            Dim ch As String
            For k = bracketOpen - 1 To 1 Step -1
                ch = Mid(result, k, 1)
                If ch = "=" Or ch = "," Or ch = "(" Or ch = "+" Or _
                   ch = "-" Or ch = "*" Or ch = "/" Or ch = "&" Then
                    Exit For
                End If
                blockStart = k
            Next k
        End If

        Dim shName As String
        shName = Mid(result, bracketClose + 1, blockEnd - bracketClose)
        shName = Replace(shName, "'", "")
        shName = Trim(shName)

        If mapping.Exists(shName) Then
            Dim localTarget As String
            localTarget = CStr(mapping(shName))

            If localTarget = "SKIP" Then
                skipCell = True
                result = Replace(result, Chr(1), "[")
                ReplaceExternalRefs = result
                Exit Function
            ElseIf localTarget = "" Then
                needsValue = True
                result = Replace(result, Chr(1), "[")
                ReplaceExternalRefs = result
                Exit Function
            Else
                Dim newRef As String
                If NeedsQuotes(localTarget) Then
                    newRef = "'" & localTarget & "'!"
                Else
                    newRef = localTarget & "!"
                End If

                Dim before As String
                Dim after  As String
                before = Left(result, blockStart - 1)
                after = Mid(result, excl + 1)
                result = before & newRef & after
            End If
        Else
            result = Left(result, bracketOpen - 1) & Chr(1) & Mid(result, bracketOpen + 1)
        End If

NextIter:
    Loop

    result = Replace(result, Chr(1), "[")
    ReplaceExternalRefs = result
End Function

' =============================================================
Private Function NeedsQuotes(shName As String) As Boolean
    Dim specialChars As String
    specialChars = " !@#$%^&*()-+=[]{}|;:,.<>?/\`~"
    Dim i As Integer
    For i = 1 To Len(specialChars)
        If InStr(shName, Mid(specialChars, i, 1)) > 0 Then
            NeedsQuotes = True
            Exit Function
        End If
    Next i
    ' Qualsiasi carattere non-ASCII (euro, accenti, ecc.)
    For i = 1 To Len(shName)
        If AscW(Mid(shName, i, 1)) > 127 Then
            NeedsQuotes = True
            Exit Function
        End If
    Next i
    If shName Like "[0-9]*" Then NeedsQuotes = True
End Function





