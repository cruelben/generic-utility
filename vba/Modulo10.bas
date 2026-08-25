Attribute VB_Name = "Modulo10"
Option Explicit

Sub m10_Utility()
Attribute m10_Utility.VB_ProcData.VB_Invoke_Func = "m\n14"
    ' UserForm1UTILITY.Show
    UserForm1UTILITY.Show vbModeless
End Sub

Sub m10_EsportaTuttiIModuli_Completo()
    Dim VBProj As Object
    Dim VBComp As Object
    Dim percorso As String
    Dim f As Integer
    Set VBProj = ThisWorkbook.VBProject
    percorso = ThisWorkbook.path & "\ExportVBA\"
    If Dir(percorso, vbDirectory) = "" Then MkDir percorso
    
    ' 1. Esporta moduli, form e ThisWorkbook
    For Each VBComp In VBProj.VBComponents
        Select Case VBComp.Type
            Case 1 ' Moduli standard
                VBComp.Export percorso & VBComp.Name & ".bas"
            Case 3 ' UserForm
                VBComp.Export percorso & VBComp.Name & ".frm"
            Case 100 ' ThisWorkbook
                If VBComp.Name = "ThisWorkbook" Then
                    If VBComp.CodeModule.CountOfLines > 0 Then
                        VBComp.Export percorso & VBComp.Name & ".cls"
                    End If
                End If
        End Select
    Next VBComp
    
    ' 2. Crea ImportMod.txt con la macro di importazione
    f = FreeFile
    Open percorso & "ImportMod.txt" For Output As #f
    Print #f, "' ============================================================"
    Print #f, "' ISTRUZIONI:"
    Print #f, "' 1. Nel file Excel di destinazione apri il VBA Editor (ALT+F11)"
    Print #f, "' 2. Inserisci un nuovo modulo standard"
    Print #f, "' 3. Nelle Proprieta (F4) rinominalo da Modulo1 a ImportMod"
    Print #f, "' 4. Incolla tutto il codice qui sotto"
    Print #f, "' 5. Avvia la macro ImportaTuttiIModuli"
    Print #f, "' ============================================================"
    Print #f, ""
    Print #f, "Sub ImportaTuttiIModuli()"
    Print #f, ""
    Print #f, "    Dim VBProj As Object"
    Print #f, "    Dim VBComp As Object"
    Print #f, "    Dim percorso As String"
    Print #f, "    Dim file As String"
    Print #f, "    Dim nome As String"
    Print #f, "    Dim daRimuovere As Collection"
    Print #f, ""
    Print #f, "    Set VBProj = ThisWorkbook.VBProject"
    Print #f, "    percorso = ThisWorkbook.Path & ""\ExportVBA\"""
    Print #f, ""
    Print #f, "    If Dir(percorso, vbDirectory) = """" Then"
    Print #f, "        MsgBox ""Cartella ExportVBA non trovata!"", vbCritical"
    Print #f, "        Exit Sub"
    Print #f, "    End If"
    Print #f, ""
    Print #f, "    ' 1. Raccoglie i moduli standard da rimuovere (tutti tranne ImportMod e ModuloMainframe)"
    Print #f, "    Set daRimuovere = New Collection"
    Print #f, "    For Each VBComp In VBProj.VBComponents"
    Print #f, "        If VBComp.Type = 1 And VBComp.Name <> ""ImportMod"" And VBComp.Name <> ""ModuloMainframe"" Then"
    Print #f, "            daRimuovere.Add VBComp"
    Print #f, "        End If"
    Print #f, "    Next VBComp"
    Print #f, ""
    Print #f, "    ' 2. Li rimuove"
    Print #f, "    Dim item As Object"
    Print #f, "    For Each item In daRimuovere"
    Print #f, "        VBProj.VBComponents.Remove item"
    Print #f, "    Next item"
    Print #f, ""
    Print #f, "    ' 3. Importa tutto"
    Print #f, "    file = Dir(percorso & ""*.*"")"
    Print #f, "    Do While file <> """""
    Print #f, ""
    Print #f, "        Select Case LCase(Right(file, 4))"
    Print #f, ""
    Print #f, "            Case "".bas"""
    Print #f, "                VBProj.VBComponents.Import percorso & file"
    Print #f, ""
    Print #f, "            Case "".frm"""
    Print #f, "                nome = Left(file, InStrRev(file, ""."") - 1)"
    Print #f, "                If EsisteComponente(VBProj, nome) Then"
    Print #f, "                    VBProj.VBComponents.Remove VBProj.VBComponents(nome)"
    Print #f, "                End If"
    Print #f, "                VBProj.VBComponents.Import percorso & file"
    Print #f, ""
    Print #f, "            Case "".cls"""
    Print #f, "                nome = Left(file, InStrRev(file, ""."") - 1)"
    Print #f, "                If EsisteComponente(VBProj, nome) Then"
    Print #f, "                    With VBProj.VBComponents(nome).CodeModule"
    Print #f, "                        .DeleteLines 1, .CountOfLines"
    Print #f, "                        .AddFromFile percorso & file"
    Print #f, "                    With"
    Print #f, "                End If"
    Print #f, ""
    Print #f, "        End Select"
    Print #f, ""
    Print #f, "        file = Dir"
    Print #f, "    Loop"
    Print #f, ""
    Print #f, "    ' 4. Si autodistrugge"
    Print #f, "    VBProj.VBComponents.Remove VBProj.VBComponents(""ImportMod"")"
    Print #f, ""
    Print #f, "    MsgBox ""Importazione completata."", vbInformation"
    Print #f, ""
    Print #f, "End Sub"
    Print #f, ""
    Print #f, "Function EsisteComponente(VBProj As Object, nome As String) As Boolean"
    Print #f, "    On Error Resume Next"
    Print #f, "    EsisteComponente = Not VBProj.VBComponents(nome) Is Nothing"
    Print #f, "    On Error GoTo 0"
    Print #f, "End Function"
    Close #f
    
    ' 3. MsgBox finale con istruzioni
    MsgBox "Esportazione COMPLETA in:" & vbCrLf & percorso & vbCrLf & vbCrLf & _
           "ISTRUZIONI PER IMPORTARE:" & vbCrLf & vbCrLf & _
           "1. Apri il file Excel di destinazione" & vbCrLf & _
           "2. Apri il VBA Editor (ALT+F11)" & vbCrLf & _
           "3. Inserisci un nuovo modulo standard" & vbCrLf & _
           "4. Nelle Proprieta (F4) rinominalo in ImportMod" & vbCrLf & _
           "5. Apri il file ImportMod.txt dalla cartella ExportVBA" & vbCrLf & _
           "6. Copia tutto il contenuto e incollalo nel modulo" & vbCrLf & _
           "7. Avvia la macro ImportaTuttiIModuli", vbInformation
End Sub

