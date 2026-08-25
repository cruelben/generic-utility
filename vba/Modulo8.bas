Attribute VB_Name = "Modulo8"
Sub m8_DOPPIA_FINESTRA_STESSO_FOGLIO()
Attribute m8_DOPPIA_FINESTRA_STESSO_FOGLIO.VB_ProcData.VB_Invoke_Func = "u\n14"

    Dim ws As Worksheet
    Set ws = ActiveSheet  ' memorizza il foglio corrente

    ' Apre una nuova finestra dello stesso file
    ActiveWorkbook.NewWindow

    ' Assicura che entrambe le finestre mostrino lo stesso foglio
    ws.Activate

    ' Dispone SOLO le finestre di questo workbook verticalmente
    ActiveWorkbook.Windows.Arrange ArrangeStyle:=xlVertical

End Sub
