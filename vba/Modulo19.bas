Attribute VB_Name = "Modulo19"
Sub m19_elimina_spazio_nascosto()

    Dim ws As Worksheet, ur As Range
    Dim r As Long, c As Long
    Dim firstRow As Long, lastRow As Long
    Dim firstCol As Long, lastCol As Long
    Dim risposta As VbMsgBoxResult

    Set ws = ActiveSheet

    risposta = MsgBox("Nel foglio """ & ws.Name & """ verranno eliminate definitivamente le righe e le colonne nascoste." & _
        vbNewLine & vbNewLine & "Vuoi procedere?", vbYesNo + vbExclamation, "Conferma eliminazione")

    If risposta = vbNo Then Exit Sub

    On Error GoTo GestioneErrore

    Set ur = ws.UsedRange

    firstRow = ur.Row
    lastRow = ur.Row + ur.Rows.count - 1
    firstCol = ur.Column
    lastCol = ur.Column + ur.Columns.count - 1

    'Estende a destra se subito dopo ci sono colonne nascoste/larghezza 0
    Do While lastCol < ws.Columns.count
        If ws.Columns(lastCol + 1).Hidden Or ws.Columns(lastCol + 1).ColumnWidth = 0 Then
            lastCol = lastCol + 1
        Else
            Exit Do
        End If
    Loop

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For r = lastRow To firstRow Step -1
        If ws.Rows(r).Hidden Or ws.Rows(r).RowHeight = 0 Then ws.Rows(r).Delete
    Next r

    For c = lastCol To firstCol Step -1
        If ws.Columns(c).Hidden Or ws.Columns(c).ColumnWidth = 0 Then ws.Columns(c).Delete
    Next c

Fine:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ws.Range("A1").Select
    MsgBox "Operazione completata.", vbInformation
    Exit Sub

GestioneErrore:
    Resume Fine

End Sub



