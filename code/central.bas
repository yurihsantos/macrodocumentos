Public Pri_Lin As Integer
Public Ult_Lin As Integer
Public Ult_Col As Integer
Public Fil_Lin As Integer

Public Local_Planilha As String
Public ColunaIterada As Integer
Public Critério_Mes  As Integer
Public Critério_Ano  As Integer

Public Dados_Lei As String
Public Dados_Mat As String
Public Dados_Pro As String
Public Dados_Esc As String
Public Dados_Pre As String
Public Rel_Comun As Long
Public Rel_Posit As Long
Public Rel_Negat As Long
Public Rel_Susta As Long
Public Rel_Anula As Long
Public Rel_NRepa As Long
Public Rel_Judic As Long
Public Rel_Extra As Long
Public Rel_Parti As Long

Public Qtd_Judic As Integer
Public Qtd_Extra As Integer
Public Qtd_Parti As Integer
Public Val_Judic As Single
Public Val_Extra As Single
Public Val_Parti As Single
Public Val_Geral As Single

Public Extens As String
Public NomeDo As String
Public PlaRel As Object

Public Registro As Integer
Public PlaTemp As String
Public WordApp As Object
Public Document As Object
'----------------------------------------------------------------------
Private Sub ReexibirPlanilhas()
  Sheets("Filtro").Visible = -1: Sheets("Comunicado").Visible = -1
  Sheets("DREI").Visible = -1:   Sheets("Relatório").Visible = -1
End Sub
'----------------------------------------------------------------------
Private Sub OcultarPlanilhas()
  Sheets("Filtro").Visible = 2: Sheets("Comunicado").Visible = 2
  Sheets("DREI").Visible = 2: Sheets("Relatório").Visible = 2
End Sub
'----------------------------------------------------------------------
Private Sub AtualizarTela0()
  Application.ScreenUpdating = False 'Impede que o Excel atualize a tela
  Application.DisplayAlerts = False 'Impede que o Excel exiba alertas
End Sub
'----------------------------------------------------------------------
Private Sub AtualizarTela1()
    Application.ScreenUpdating = True 'Permite que o Excel atualize a tela
    Application.DisplayAlerts = True 'Permite que o Excel exiba alertas
End Sub
'----------------------------------------------------------------------
Private Sub IntraPlanilha()
    AtualizarTela0
    ReexibirPlanilhas
    LimparFiltro
    QtdLinha
    QtdColuna
    FiltrarMes
    IntervaloFiltrado
End Sub
'----------------------------------------------------------------------
Private Sub QtdLinha()
    P1.Activate
    Ult_Lin = Range("TabelaDados").Rows.Count + 1 'Captura a Última Linha
End Sub
'----------------------------------------------------------------------
Private Sub QtdColuna()
    P1.Activate
    Ult_Col = Range("TabelaDados").Columns.Count 'Captura a Última Coluna
End Sub
'----------------------------------------------------------------------
Private Sub LimparFiltro()
    Range("TabelaDados").AutoFilter
End Sub
'----------------------------------------------------------------------
Private Sub FiltrarMes()
    Range("TabelaDados").AutoFilter _
    Field:=6, Operator:=xlFilterValues, _
    Criteria2:=Array(1, Critério_Mes & "/" & Critério_Ano)

    If NomeDo = "Comunicado" Then
        Range("TabelaDados").AutoFilter _
        Field:=19, Criteria1:=Array(1) 'Filtra a coluna 19 somente com o primeiro lote
    End If
End Sub
'----------------------------------------------------------------------
Private Sub IntervaloFiltrado()
    Fil_Lin = WorksheetFunction.Subtotal(3, Range("TabelaDados[Data]"))
    'Verifica quantos registros estão filtrados

    If Fil_Lin = 0 Then
        Exit Sub
        'Se não houver filtrados, então saia da função
    Else
        Range("TabelaDados").SpecialCells(xlCellTypeVisible).Copy
        Range("TabelaFiltro").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
        Fil_Lin = Fil_Lin + 1
        'Se houver filtrados, então os copie para a
        'tabela de filtrados e corrija o número da linha
    End If
End Sub
'----------------------------------------------------------------------
Sub Comunicado()
    IntraPlanilha
    ComunicadoInterno
    AbrirModeloExcel
    CopiarDados
    AjustarBorda
    ExportarPlanilha

    LimparFiltrados
    LimparFiltro
    OcultarPlanilhas
    AtualizarTela1
End Sub
'----------------------------------------------------------------------
Sub DREI()
    IntraPlanilha
    DREIInterno
    AbrirModeloExcel
    CopiarDados
    AjustarBorda
    ExportarPlanilha

    LimparFiltrados
    LimparFiltro
    OcultarPlanilhas
    AtualizarTela1
End Sub
'----------------------------------------------------------------------
Sub Relatorio()
    IntraPlanilha
    DadosRelatorio
    RelatorioInterno
    SalvarPlanilhaTemp
    DocumentoRelatorio
    ApagarPlanilhaTemporária

    LimparFiltrados
    LimparFiltro
    OcultarPlanilhas
    AtualizarTela1
End Sub
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------
Private Sub ComunicadoInterno()
    P3.Activate
    
    If Fil_Lin = 0 Then
       Exit Sub
    Else:
        Range("A2").Formula2R1C1 = "=TabelaFiltro[DAT]"
        Range("B2").Formula2R1C1 = "=TabelaFiltro[HOR]"
        Range("C2").Formula2R1C1 = "=TabelaFiltro[LOC]"
        Range("D2").Formula2R1C1 = "=TabelaFiltro[TIT]"
        Range("E2").Formula2R1C1 = "=TabelaFiltro[PRO]"
        Range("F2").Formula2R1C1 = "=""Publicado em ""&TabelaFiltro[PUB]&"" no ""&TabelaFiltro[LOP]"
    End If
End Sub
'----------------------------------------------------------------------
Private Sub DREIInterno()
    P4.Activate

    If Fil_Lin = 0 Then
       Exit Sub
    Else:
        Range("A2").Formula2R1C1 = "=TabelaFiltro[TIP]"
        Range("B2").Formula2R1C1 = "=TabelaFiltro[VEN]"
        Range("C2").Formula2R1C1 = "=TabelaFiltro[LOT]"
        Range("D2").Formula2R1C1 = "=TabelaFiltro[DES]"
        Range("E2").Formula2R1C1 = "=TabelaFiltro[MIN]"
        Range("F2").Formula2R1C1 = "=TabelaFiltro[VAL]"
    End If
End Sub
'----------------------------------------------------------------------
Private Sub DadosRelatorio()
    P2.Activate
    
    CO = "Comunicado": PO = "Positivo": NE = "Negativo": SU = "Sustado": AN = "Anulado"
    NA = "Não Repassado": JU = "Judicial": EX = "Extrajudicial": PA = "Particular"
    'TI = Tipo de Leilão, SI = Situação do Lote, VE = Valor de Venda
    
    Set TI = Range("TabelaFiltro[TIP]")
    Set SI = Range("TabelaFiltro[SIT]")
    Set VE = Range("TabelaFiltro[VAL]")
    
    Rel_Posit = Application.CountIf(SI, PO)
    Rel_Negat = Application.CountIf(SI, NE)
    Rel_Susta = Application.CountIf(SI, SU)
    Rel_Anula = Application.CountIf(SI, AN)
    Rel_NRepa = Application.CountIf(SI, NA)
    Rel_Comun = Rel_Posit + Rel_Negat + Rel_Susta + Rel_Anula

    Qtd_Judic = Application.CountIfs(TI, JU, SI, PO)
    Qtd_Extra = Application.CountIfs(TI, EX, SI, PO)
    Qtd_Parti = Application.CountIfs(TI, PA, SI, PO)

    Val_Judic = Application.SumIfs(VE, TI, JU, SI, PO)
    Val_Extra = Application.SumIfs(VE, TI, EX, SI, PO)
    Val_Parti = Application.SumIfs(VE, TI, PA, SI, PO)
    Val_Geral = Val_Judic + Val_Extra + Val_Parti
End Sub
'----------------------------------------------------------------------
Private Sub RelatorioInterno()
    With Sheets("Relatório")
        .Range("A2") = Critério_Mes & "/" & Critério_Ano      'Mês / Ano
        .Range("B2") = Dados_Lei                              'Leiloeiro
        .Range("C2") = Dados_Mat                              'Matrícula
        .Range("D2") = Dados_Pro                              'Pronome
        .Range("E2") = Dados_Esc                              'Escritório
        .Range("F2") = Dados_Pre                              'Preposto
        .Range("G2") = " "                                    '
        .Range("H2") = Rel_Comun                              'Comunicados
        .Range("I2") = Rel_Posit                              'Positivos
        .Range("J2") = Rel_Negat                              'Negativos
        .Range("K2") = Rel_Susta                              'Sustados
        .Range("L2") = Rel_Anula                              'Anulados
        .Range("M2") = Rel_NRepa                              'Não Repassados
        .Range("N2") = " "                                    '
        .Range("O2") = Qtd_Judic                              'Qtd. Judiciais
        .Range("P2") = Qtd_Extra                              'Qtd. Adm. Pub.
        .Range("Q2") = Qtd_Parti                              'Qtd. Particular
        .Range("R2") = " "                                    '
        .Range("S2") = Val_Judic                              'Soma Judicial
        .Range("T2") = Val_Extra                              'Soma Adm. Pub.
        .Range("U2") = Val_Parti                              'Soma Particular
        .Range("V2") = Val_Geral                              'Soma Geral
    End With
End Sub
'----------------------------------------------------------------------
Private Sub AbrirModeloExcel()
    Set PlaRel = Workbooks.Open(Environ("OneDrive") & _
    "\Escritorio Brame Leiloes\Modelos\JUCERJA\Relatórios Mensais\" _
    & NomeDo & ".xlsx")
End Sub
'----------------------------------------------------------------------
Private Sub CopiarDados()
    PlaRel.Activate
    PlaRel.Sheets(NomeDo).Range("F3").Value = Critério_Mes & "/" & Critério_Ano
    PlaRel.Sheets(NomeDo).Range("D3").Value = UCase(Dados_Lei)

    If Fil_Lin = 0 Then
      With PlaRel
        .Sheets(NomeDo).Range(Cells(6, 1), Cells(6, 6)).Merge
        .Sheets(NomeDo).Range(Cells(6, 1), Cells(6, 6)).Value = "NÃO HOUVE LEILÕES NESTE MÊS"
      End With
    Else:
      With ThisWorkbook.Sheets(NomeDo)
        .Activate
        .Range(Cells(2, 1), Cells(Fil_Lin, 6)).Copy
      End With
      With PlaRel
        .Activate
        .Sheets(NomeDo).Range(Cells(6, 1), Cells((Fil_Lin + 4), 6)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
      End With
   End If
End Sub
'----------------------------------------------------------------------
Private Sub AjustarBorda()
   If Fil_Lin = 0 Then
      With PlaRel.Sheets(NomeDo).Range(Cells(6, 1), Cells(6, 6))
          .Borders.LineStyle = xlContinuous
          .Font.Name = "Tenorite"
          .Font.Size = 11
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
       End With
   
   Else:
       Rows("6:" & Fil_Lin + 4).RowHeight = 40
       With PlaRel.Sheets(NomeDo).Range(Cells(6, 1), Cells((Fil_Lin + 4), 6))
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            .Font.Name = "Tenorite"
            .Font.Size = 8
            .WrapText = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    End If
End Sub
'----------------------------------------------------------------------
Private Sub ExportarPlanilha()
    PlaRel.Activate
    If Extens = "xlsx" Then
         ActiveWorkbook.SaveAs Filename:=Local_Planilha & _
         "\" & Critério_Mes & "." & Critério_Ano & " - " & NomeDo & "." & Extens
    Else: ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF _
         , Filename:=Local_Planilha & "\" & Critério_Mes & "." & Critério_Ano & " - " & NomeDo & "." & Extens
    End If
    PlaRel.Close
End Sub
'----------------------------------------------------------------------
Private Sub AbrirModeloWord()
    Set PlaRel = Workbooks.Open(Environ("OneDrive") & _
    "\Escritorio Brame Leiloes\Modelos\Informações de Leilão\" _
    & NomeDo & ".docx")
End Sub
'----------------------------------------------------------------------
Private Sub SalvarPlanilhaTemp()
    ThisWorkbook.Save
    Sheets("Relatório").Select
    Dim NovoWB As Workbook 'Seta uma variável para se referir a nova pasta de trabalho
    Set NovoWB = Workbooks.Add(xlWBATWorksheet) 'Cria esta nova aba
    
    With NovoWB 'Função With
    ThisWorkbook.ActiveSheet.Copy After:=.Worksheets(.Worksheets.Count) 'Copia a aba para nova planilha como Plan2
    .Worksheets(1).Delete 'Deleta Plan1
    .SaveAs ThisWorkbook.Path & "\$PlaTemp$.xlsx" 'Salva na mesma pasta
    .Close False
    End With
End Sub
'----------------------------------------------------------------------
Private Sub ApagarPlanilhaTemporária()
    Kill (ThisWorkbook.Path & "\$PlaTemp$.xlsx")
End Sub
'----------------------------------------------------------------------
Private Sub DocumentoRelatorio()
    
    Set WordApp = CreateObject("Word.Application")                          'Define Document como um documento do Word
    WordApp.Visible = True                                                  'Torna o Word visível
    Set Document = WordApp.Documents.Open(Environ("OneDrive") & _
    "\Escritorio Brame Leiloes\Modelos\JUCERJA\Relatórios Mensais\" _
    & NomeDo & ".docx")                                             'Abre um documento definido na pasta dos modelos
 
    PlaTemp = ThisWorkbook.Path & "\$PlaTemp$.xlsx"
        
    Document.MailMerge.OpenDataSource Name:=PlaTemp, _
    SQLStatement:="SELECT * FROM `Relatório$`" 'Seleciona o destinatário e ativa a mala direta

    'Desativa os resultados para limpar qualquer resultado anterior e retornar aos dados
    Document.MailMerge.ViewMailMergeFieldCodes = True 'Desativa a visualização dos resultados
    Document.MailMerge.ViewMailMergeFieldCodes = False 'Ativa a visualização dos resultados

    Document.MailMerge.DataSource.ActiveRecord = wdFirstRecord 'Define o primeiro registro da mala direta
    Registro = Document.MailMerge.DataSource.RecordCount 'Contador de registros

    If Extens = "docx" Then
        MalaDiretaWord
    ElseIf Extens = "pdf" Then
        MalaDiretaPDF
    End If

    Document.Save
    Document.Close
    WordApp.Quit
End Sub
'----------------------------------------------------------------------
Private Sub MalaDiretaWord()
 For Registro = 1 To Registro 'Salva todos os registros
   With Document.MailMerge
    .Destination = wdSendToNewDocument 'Mandar para um novo documento
    .DataSource.FirstRecord = ActiveDocument.MailMerge.DataSource.ActiveRecord 'Primeiro registro: atual
    .DataSource.LastRecord = ActiveDocument.MailMerge.DataSource.ActiveRecord 'Último registro: atual
    .Execute Pause:=False 'Finalizar ação
   End With
   ActiveDocument.SaveAs2 Filename:=Local_Planilha & "\" & Critério_Mes & "." & Critério_Ano & " - Relatório.docx", FileFormat:=wdFormatXMLDocument, CompatibilityMode:=wdCurrent 'Salvar o documento como Word
   ActiveDocument.Close 'Fecha o documento
   Document.MailMerge.DataSource.ActiveRecord = wdNextRecord 'Avança para o próximo registro
 Next Registro 'Próximo registro
End Sub
'----------------------------------------------------------------------
Private Sub MalaDiretaPDF()
    For Registro = 1 To Registro 'Salva todos os registros
        Document.SaveAs2 Filename:=Local_Planilha & "\" & Critério_Mes & "." & Critério_Ano & " - Relatório.pdf", FileFormat:=wdFormatPDF, CompatibilityMode:=wdCurrent 'Salva como PDF
        Document.MailMerge.DataSource.ActiveRecord = wdNextRecord 'Próximo registro da mala direta
    Next Registro 'Próximo registro
End Sub
'----------------------------------------------------------------------
Private Sub LimparFiltrados()
    ThisWorkbook.Activate
    
    If Fil_Lin = 0 Then
      P2.Activate
      Range("TabelaFiltro").Clear
    
      P3.Activate
      Rows("2:2").Delete Shift:=xlUp
    
      P4.Activate
      Rows("2:2").Delete Shift:=xlUp
    
      P5.Activate
      Rows("2:2").Delete Shift:=xlUp
    
    Else:
      P2.Activate
      Range("TabelaFiltro").Clear
      
      P3.Activate
      Rows("2:" & Fil_Lin).Delete Shift:=xlUp
      
      P4.Activate
      Rows("2:" & Fil_Lin).Delete Shift:=xlUp
      
      P5.Activate
      Rows("2:2").Clear
    End If
End Sub
'----------------------------------------------------------------------
Sub ReproduçãoAnual()
  For CriterioMensal = 1 To 12
    Critério_Mes = CriterioMensal
    Application.Run (NomeDo)
  Next
End Sub

