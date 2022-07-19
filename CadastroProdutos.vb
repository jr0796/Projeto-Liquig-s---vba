Private Sub UserForm_Activate()
    Dim ref As Worksheet
    Dim linha, coluna As Integer
    
    Set ref = ThisWorkbook.Worksheets(6)
    linha = 2
    coluna = 1

    'cabeçalho listviwer
    With ltw_resultado
            .AllowColumnReorder = False
            .Gridlines = True
            .View = lvwReport
            .FullRowSelect = True
            ltw_resultado.ColumnHeaders.Add Text:="Código", Width:=70
            'ltw_result.ColumnHeaders.Add Text:="CPF", Width:=60
            ltw_resultado.ColumnHeaders.Add Text:="Nome do Produto", Width:=90
            ltw_resultado.ColumnHeaders.Add Text:="Fornecedor", Width:=70
            'ltw_result.ColumnHeaders.Add Text:="RG", Width:=70
            ltw_resultado.ColumnHeaders.Add Text:="Preço de custo", Width:=70
            ltw_resultado.ColumnHeaders.Add Text:="Preço para venda", Width:=80
            ltw_resultado.ColumnHeaders.Add Text:="Data de cadastro", Width:=80
            'ltw_result.ColumnHeaders.Add Text:="CNH", Width:=60
            'ltw_result.ColumnHeaders.Add Text:="Função", Width:=60
            
    End With
    
    btn_atualiazar.Visible = False
    btn_excluir.Enabled = False
    btn_editar.Enabled = False
    
    
    'Carregando fornecedores na combo
    With ref
        Do Until Plan6.Cells(linha, coluna) = ""
            cbx_fornecedor.AddItem Plan6.Cells(linha, coluna)
            linha = linha + 1
            'coluna = coluna + 1
        Loop
            
        End With
    
    'Carregando data atual no DTimepicker
    dtp_data_cadastro.Value = DateValue(Now)
    
    
End Sub

Private Sub buscadeprodutos()
    Dim guia As Worksheet
    Dim linha As Integer
    Dim coluna As Integer
    Dim linhaslistviwer As Integer
    Dim valor_celula As String
    Dim contador_registros As Integer
    
    'setando a variável guia
    Set guia = ThisWorkbook.Worksheets(7)
    
    linha = 2
    coluna = 1
    contador_registros = 0
    
    'limpando Listview
    ltw_resultado.ListItems.Clear
    
    If txt_cod.Text = "" Then
    'limpando campos
       
        
        
    End If
    
    With guia
        
         Do Until Plan7.Cells(linha, 1) = ""
            valor_celula = Plan7.Cells(linha, coluna).Value
            
            If UCase(Left(valor_celula, Len(produtos_pesquisados))) = UCase(produtos_pesquisados) And txt_cod.Text <> "" Then
             
                With frm_cadastro_produtos.ltw_resultado.ListItems.Add
                     .Text = Plan7.Cells(linha, 1) 'Cod
                    .SubItems(1) = Plan7.Cells(linha, 2) 'Nome do produto
                    .SubItems(2) = Plan7.Cells(linha, 4) 'Fornecedor
                    .SubItems(3) = Plan7.Cells(linha, 6) 'Preço de custo
                    .SubItems(4) = Plan7.Cells(linha, 7) 'Preço de venda
                    .SubItems(5) = Plan7.Cells(linha, 3) 'Data de cadastro
                    '.SubItems(6) = Plan5.Cells(linha, 9)
                    '.SubItems(7) = Plan5.Cells(linha, 10)
                    '.SubItems(8) = Plan5.Cells(linha, 11)
                    '.SubItems(9) = Plan5.Cells(linha, 12)
                    linhaslistviwer = linhaslistviwer + 1
    End With
            End If
            linha = linha + 1
        
        Loop
    End With
   'zip
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Salvando a planilha e as informações
                ThisWorkbook.Save
End Sub
