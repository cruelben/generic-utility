Attribute VB_Name = "Modulo9"
Sub m9_RICALCOLA_TUTTO()
    Dim ws As Worksheet
    Dim BeforeVal As Double, AfterVal As Double
    
    ' Calcolo un "totale" prima
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        BeforeVal = BeforeVal + Application.WorksheetFunction.Sum(ws.UsedRange)
        On Error GoTo 0
    Next ws
    
    ' Ricalcolo totale
    Application.CalculateFullRebuild
    
    ' Calcolo di nuovo il "totale" dopo
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        AfterVal = AfterVal + Application.WorksheetFunction.Sum(ws.UsedRange)
        On Error GoTo 0
    Next ws
    
    ' Confronto
    If BeforeVal = AfterVal Then
        MsgBox "Nessuna variazione rilevata dopo il ricalcolo.", vbInformation
    Else
        MsgBox "Attenzione: alcuni valori sono cambiati dopo il ricalcolo!", vbExclamation
    End If
End Sub
