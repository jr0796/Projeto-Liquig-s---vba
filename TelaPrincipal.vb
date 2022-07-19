Private Sub UserForm_Activate()
     
     'Botões MINIMIZAR e MAXIMIZAR extraidos da internet
            
            Dim lngMyHandle As Long, lngCurrentStyle As Long, lngNewStyle As Long
                If Application.Version < 9 Then
                    lngMyHandle = FindWindow("THUNDERXFRAME", Me.Caption)
                Else
                    lngMyHandle = FindWindow("THUNDERDFRAME", Me.Caption)
                End If
                
                lngCurrentStyle = GetWindowLong(lngMyHandle, GWL_STYLE)
                lngNewStyle = lngCurrentStyle Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
                SetWindowLong lngMyHandle, GWL_STYLE, lngNewStyle
        
            
            '#########################################################################
            
     
    If aberto = 1 Then
    
            Application.Visible = False
            
            
            
             'statusbar
            StatusBar1.Panels.Add ("1")
            StatusBar1.Panels.Add ("2")
            StatusBar1.Panels.Add ("3")
            StatusBar1.Panels.Add ("4")
            StatusBar1.Panels.Add ("5")
            
            StatusBar1.Panels(1).Width = 200
            StatusBar1.Panels(2).Width = 100
            StatusBar1.Panels(3).Width = 100
            StatusBar1.Panels(4).Width = 100
            StatusBar1.Panels(5).Width = 550
            
            StatusBar1.Panels(1).Alignment = sbrLeft
            StatusBar1.Panels(2).Alignment = sbrCenter
            StatusBar1.Panels(3).Alignment = sbrCenter
            StatusBar1.Panels(4).Alignment = sbrCenter
            StatusBar1.Panels(5).Alignment = sbrCenter
        
        
        
            StatusBar1.Panels(1).Text = FormatDateTime(DateTime.Now, vbLongDate)
            StatusBar1.Panels(3).Text = "##########################"
            
            
            
            If TimeValue(Now) <= ("12:00:00") Then
                   StatusBar1.Panels(5).Text = "Bom dia"
            End If
                
            If TimeValue(Now) >= ("13:00:00") And TimeValue(Now) <= ("18:00:00") Then
                   StatusBar1.Panels(5).Text = "Boa tarde"
            End If
            
            If TimeValue(Now) >= ("19:00:00") And TimeValue(Now) <= ("24:00:00") Then
                   StatusBar1.Panels(5).Text = "Boa Noite"
            End If
            
            'Escondendo botão configurações;
            'btn_exibir.Visible = False
            
            'carregando relógio na statusbar
            Application.Run "inicio_contagem2"
            aberto = 2
        
        'Caregando um loop para verificação de estoque
        
       ' Application.Run "virificacao_estoque"
        
    End If
            
         Application.OnKey "%{END}", "call mostrar"
   
    
    
End Sub



Private Sub btn_cadastro_de_produtos_Click()
    frm_cadastro_produtos.Show
End Sub

Private Sub btn_cli_Click()
    frm_cadastro_cli.Show
End Sub

Private Sub btn_estok_Click()
    frm_estoque.Show
End Sub
Private Sub btn_exibir_Click()
    Application.Visible = True
End Sub

Private Sub btn_forn_Click()
    frm_cadastro_fornecedores.Show
End Sub
Private Sub btn_func_Click()
    frm_cadastro_func.Show
End Sub
Private Sub btn_produtos_Click()
    frm_cadastro_produtos.Show
End Sub

Private Sub btn_relatorios_Click()
    frm_relatorios.Show
End Sub

Private Sub btn_vendas_Click()
    frm_vendas.Show
End Sub

Private Sub btnCadastro_Click()

End Sub

Private Sub CommandButton51_Click()
    frm_estoque.Show
End Sub


Sub mostrar()
MsgBox = "Olá Mundo Opss!!"
btn_exibir.Visible = True





End Sub

Private Sub UserForm_Click()
    Application.Visible = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Application.Run "parar_relogio2"
    
    ThisWorkbook.Save
End Sub

Private Sub UserForm_Terminate()
    'Application.Visible = True
End Sub
