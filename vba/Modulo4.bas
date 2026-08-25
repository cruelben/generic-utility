Attribute VB_Name = "Modulo4"
Sub m4_ControllaCollegamentiFoglioAttivo()
    Dim sh As Worksheet
    Dim cel As Range
    Dim foglioDaCancellare As String
    Dim msg As String
    Dim risultati As String
    Dim fileNum As Integer
    Dim percorsoFile As String
    
    ' Nome del foglio attivo
    foglioDaCancellare = ActiveSheet.Name
    msg = "Celle collegate al foglio """ & foglioDaCancellare & """:" & vbCrLf & vbCrLf
    risultati = "Celle collegate al foglio """ & foglioDaCancellare & """:" & vbCrLf & vbCrLf
    
    ' Cicla su tutti gli altri fogli
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name <> foglioDaCancellare Then
            For Each cel In sh.UsedRange
                If cel.HasFormula Then
                    If InStr(1, cel.formula, "'" & foglioDaCancellare & "'!", vbTextCompare) > 0 _
                       Or InStr(1, cel.formula, foglioDaCancellare & "!", vbTextCompare) > 0 Then
                        msg = msg & sh.Name & ", cella " & cel.Address(False, False) & vbCrLf
                        risultati = risultati & sh.Name & ", cella " & cel.Address(False, False) & vbCrLf
                    End If
                End If
            Next cel
        End If
    Next sh
    
    ' Mostra risultato
    If msg = "Celle collegate al foglio """ & foglioDaCancellare & """:" & vbCrLf & vbCrLf Then
        MsgBox "Nessuna formula fa riferimento al foglio " & foglioDaCancellare, vbInformation
    Else
        MsgBox msg, vbExclamation, "Collegamenti trovati"
        
        ' Crea il file TXT solo se ci sono risultati
        percorsoFile = ThisWorkbook.path & "\" & foglioDaCancellare & ".txt"
        fileNum = FreeFile
        Open percorsoFile For Output As #fileNum
        Print #fileNum, risultati
        Close #fileNum
        
        ' Apre il file con il Blocco Note
        Shell "notepad.exe " & Chr(34) & percorsoFile & Chr(34), vbNormalFocus
    End If
    
End Sub


