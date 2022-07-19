Private Sub UserForm_Activate()
        Dim ref, ref2 As Worksheet
        Dim linha, linha2 As Integer
        
        linha = 2
        linha2 = 2
        
        'Escondendo os campos funcionários
        cbx_nome_funcionario.Visible = False
        Frame2.Visible = False
        Label2.Visible = False
        
        'Escondendo campos PGTO.
        cbx_tipo_pgto.Visible = False
        frame_tipo_pgto.Visible = False
        Label3.Visible = False
        
         With ref
            Do Until Plan4.Cells(linha, 1) = ""
            cbx_nome_funcionario.AddItem Plan4.Cells(linha, 1)
            linha = linha + 1
            Loop
        End With
        
         With ref2
            Do Until Plan1.Cells(linha2, 1) = ""
            cbx_tipo_pgto.AddItem Plan1.Cells(linha2, 2)
            linha2 = linha2 + 1
            Loop
        End With
        
        opt_opcao_todas.Value = True
        
        'data de hoje nos dtps
        DTPicker1.Value = DateValue(Now)
        DTPicker2.Value = DateValue(Now)
        
        Me.Width = 398.25
        Me.Left = 300
        
End Sub
