Attribute VB_Name = "Modulo16"
Sub m16_MostraFormulePrecedenti_slim()
Attribute m16_MostraFormulePrecedenti_slim.VB_ProcData.VB_Invoke_Func = "p\n14"

    On Error Resume Next
    Dim cell As Range
    For Each cell In Selection
        cell.ShowPrecedents
    Next cell
    On Error GoTo 0


End Sub

Sub m16_RimuoviFrecce()
Attribute m16_RimuoviFrecce.VB_ProcData.VB_Invoke_Func = "o\n14"
    ActiveSheet.ClearArrows
End Sub



