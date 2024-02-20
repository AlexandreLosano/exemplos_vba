Sub simulador()

    Plan7.Visible = xlSheetVisible
    Plan12.Visible = xlSheetHidden

End Sub

Sub gerador()

    Plan13.Visible = xlSheetVisible
    Plan12.Visible = xlSheetHidden

End Sub

Sub simulador_gerador()
    
    Plan13.Visible = xlSheetVisible
    Plan7.Visible = xlSheetHidden

End Sub

Sub gerador_simulador()

    Plan7.Visible = xlSheetVisible
    Plan13.Visible = xlSheetHidden

End Sub

Sub simulador_capa()

    Plan12.Visible = xlSheetVisible
    Plan7.Visible = xlSheetHidden

End Sub

Sub gerador_capa()
    
    Plan12.Visible = xlSheetVisible
    Plan13.Visible = xlSheetHidden

End Sub

Sub limpar_simulador()

    Plan7.Range("C5:K5").ClearContents
    Plan7.Range("C7:K7").ClearContents
    Plan7.Range("C9:K9").ClearContents
    Plan7.Range("C11:K11").ClearContents
    Plan7.Range("C13:K13").ClearContents
    Plan7.Range("C15:K15").ClearContents
    Plan7.Range("C17:K17").ClearContents
    'Plan7.Range("N9").ClearContents
    'Plan7.Range("N11").ClearContents
    Plan7.Range("E19").ClearContents
    Plan7.Range("H19").ClearContents
    Plan7.Range("K19").ClearContents
    Plan7.Range("P7") = "Sem Taxa de Combustível"
    Plan7.Range("P8").ClearContents
    Plan7.Range("P10") = "Com Taxa de Combustível"
    Plan7.Range("P11").ClearContents
    Plan7.Range("P13") = "Tempo de Trânsito"
    Plan7.Range("P14").ClearContents
    Plan7.Range("P16") = "Peso Considerado"
    Plan7.Range("P17").ClearContents



End Sub

Sub calcular()

    Plan7.Range("S7:S17").Select
    Selection.Copy
    Plan7.Range("P7:P17").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    'Plan7.Range("N9") = Plan8.Range("S14")
    'Plan7.Range("N11") = Plan8.Range("S16")
    Plan7.Range("A1").Select

End Sub

Sub dados_simulador_gerador()

    Plan13.Range("D5").Select
    Selection.ClearContents
    Plan13.Range("D7").Select
    Selection.ClearContents
    Plan13.Range("D9").Select
    Selection.ClearContents
    
    Plan13.Range("D5") = Plan13.Range("D11")
    Plan13.Range("D7") = Plan13.Range("D13")
    Plan13.Range("D9") = Plan13.Range("D15")

End Sub

Sub limpar_gerador()

    Plan13.Range("D5").Select
    Selection.ClearContents
    Plan13.Range("D7").Select
    Selection.ClearContents
    Plan13.Range("D9").Select
    Selection.ClearContents

End Sub

Sub gera_tabela()

Form_Load

If Plan13.Range("D7:G7").Text = "" Or Plan13.Range("D9:G9").Text = "" Then
    simulador1.Show
Else
   ' Plan13.Range("d5").Select
   ' Selection.ClearContents

    Plan7.Visible = True

    Sheets("Simulador").Select
    Plan10.Visible = True
    Plan10.Select
    Plan11.Visible = True
    Plan11.Select
    Plan14.Visible = True
    Plan14.Select
    Plan9.Visible = True
    Plan9.Select
    Sheet8.Visible = True
    Plan9.Select
        Sheets(Array("Express Exportação", "Economy Exportação", "Express Importação", _
        "Economy Importação", "Origens e Produtos")).Select
        Sheets("Express Exportação").Activate
    Plan9.Activate

    Plan7.Visible = False
    'Plan13.Range("C17").Select

     'Plan13.Visible = xlSheetHidden

     Salvar = Plan13.Range("D7").Text

     ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
     "C:\TNT-Express\" + Salvar + ".pdf", Quality:=xlQualityStandard, _
     IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
     True

     Application.ScreenUpdating = False
        
     Sheets("Gerador de Tabelas").Visible = True
     Sheets("Gerador de Tabelas").Select
        
     Plan10.Visible = False
     Plan11.Visible = False
     Plan14.Visible = False
     Plan9.Visible = False
     Sheet8.Visible = False


End If

End Sub

Private Sub Form_Load()

  If ArquivoExiste("C:\TNT-Express", True) Then  '<<< Nome da Pasta de Busca
    MsgBox "O Arquivo será salvo em C:\TNT-Express"
  Else
    MkDir "C:\TNT-Express" '<<< Nome da Pasta que será criada
    MsgBox ("Pasta TNT-Express criada em seu C:, é o local onde será salvo suas tabelas")
  End If

End Sub

Public Function ArquivoExiste(ByVal Caminho As String, Optional ByVal SomenteDiretorio As Boolean = False) As Boolean

    On Error Resume Next
    If SomenteDiretorio Then
        ArquivoExiste = GetAttr(Mid(Caminho, 1, InStrRev(Caminho, ""))) And vbDirectory
    Else
        ArquivoExiste = GetAttr(Caminho)
    End If
    On Error GoTo 0
End Function
