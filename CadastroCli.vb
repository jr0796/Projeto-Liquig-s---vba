Private Sub UserForm_Activate()
     'cabeçalho listviwer
    With ltw_resultado
            .AllowColumnReorder = False
            .Gridlines = True
            .View = lvwReport
            .FullRowSelect = True
            ltw_resultado.ColumnHeaders.Add Text:="Nome", Width:=90
            'ltw_result.ColumnHeaders.Add Text:="CPF", Width:=60
            ltw_resultado.ColumnHeaders.Add Text:="Endereço", Width:=90
            ltw_resultado.ColumnHeaders.Add Text:="CEP", Width:=70
            'ltw_result.ColumnHeaders.Add Text:="RG", Width:=70
            ltw_resultado.ColumnHeaders.Add Text:="Telefone", Width:=70
            ltw_resultado.ColumnHeaders.Add Text:="Data", Width:=60
            ltw_resultado.ColumnHeaders.Add Text:="Bairro", Width:=70
            'ltw_result.ColumnHeaders.Add Text:="CNH", Width:=60
            'ltw_result.ColumnHeaders.Add Text:="Função", Width:=60
            
    End With
    
    'alimentando comboBox tipo de telefone
    cbx_tipo_tel.AddItem "Fixo"
    cbx_tipo_tel.AddItem "Celular"
    
    cbx_tipo_tel.Value = "Fixo"
    
    'desabilitando btn editar
    btn_editar.Enabled = False
    
    'escondendo botão atualizar
    btn_atualizar.Visible = False
    
    'desabilitar botão excluir
    btn_exluir.Enabled = False
    
    'Carregadno daata correta
    dtp_data_cadast.Value = DateValue(Now)
    
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Salvando planilha para não perder informação
    ThisWorkbook.Save
End Sub