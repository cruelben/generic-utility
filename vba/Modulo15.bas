Attribute VB_Name = "Modulo15"
Dim formuleMemoriate() As String
Dim numRighe As Long
Dim numColonne As Long

Sub m15a_Copia_Formule()
    Dim rng As Range
    Dim i As Long, j As Long
    
    Set rng = Selection
    numRighe = rng.Rows.count
    numColonne = rng.Columns.count
    
    ReDim formuleMemoriate(1 To numRighe, 1 To numColonne)
    
    For i = 1 To numRighe
        For j = 1 To numColonne
            Dim cell As Range
            Set cell = rng.Cells(i, j)
            If cell.HasFormula Then
                formuleMemoriate(i, j) = cell.formula
            Else
                formuleMemoriate(i, j) = ""
            End If
        Next j
    Next i
    
    MsgBox "Copiate " & rng.count & " formule! Ora seleziona la cella di partenza e lancia Incolla_Formule.", vbInformation
End Sub

Sub m15b_Incolla_Formule()
    Dim rngStart As Range
    Dim i As Long, j As Long
    
    If numRighe = 0 Or numColonne = 0 Then
        MsgBox "Prima esegui Copia_Formule!", vbExclamation
        Exit Sub
    End If
    
    ' Prende solo la prima cella della selezione come punto di partenza
    Set rngStart = Selection.Cells(1, 1)
    
    For i = 1 To numRighe
        For j = 1 To numColonne
            Dim cellDest As Range
            Set cellDest = rngStart.Offset(i - 1, j - 1)
            If formuleMemoriate(i, j) <> "" Then
                cellDest.formula = formuleMemoriate(i, j)
            End If
        Next j
    Next i
    
    MsgBox "Incollate le formule a partire da " & rngStart.Address & "!", vbInformation
End Sub





