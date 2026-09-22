<%
'=======================================================================================
' PÁGINA: lca_busca_detalhamento_pedido_processar_edicao.asp
' DESCRIÇÃO: Processa e salva as edições de itens do pedido no banco de dados
' AUTOR: Sistema
' DATA: 2026-01-23
' SEGURANÇA: Validações, transações, log de alterações, proteção de dados
'=======================================================================================

Dim cod_pedido, acao
cod_pedido = Request.Form("cod_pedido")
acao = Request.Form("acao")

' Variáveis para controle de erros
Dim msg_sucesso, msg_erro, tem_erro
msg_sucesso = ""
msg_erro = ""
tem_erro = False

AbrirConexao()

' ======================= VALIDAÇÕES =======================

' 1. Validar permissão do usuário
Dim VALIDA_SUB22
VALIDA_SUB22 = formatNumber(0&inStr(sub_validacoes_paginas,"#22"),0)

If VALIDA_SUB22 = 0 Then
    msg_erro = "Você não possui permissão para editar pedidos"
    tem_erro = True
End If

' 2. Validar código do pedido
If cod_pedido = "" Or Not IsNumeric(cod_pedido) Then
    msg_erro = "Pedido inválido"
    tem_erro = True
End If

' 3. Se não houver erros, processar as edições
If Not tem_erro Then

    ' ======================= INICIAR TRANSAÇÃO =======================
    On Error Resume Next
    objConexao.BeginTrans()
    
    Dim sql, rs, qtd_alteracoes
    qtd_alteracoes = 0
    
    ' --------- EDITAR FICHAS DE IDENTIFICAÇÃO (Álbuns) ---------
    
    ' Receber lista de formandos alterados
    Dim cod_formandos_array, form_keys, i
    form_keys = Split(Request.Form("cod_formando_list"), ",")
    
    For i = LBound(form_keys) To UBound(form_keys)
        
        Dim cod_formando, inicial_novo, faturado_novo, valor_un_novo
        Dim inicial_original, faturado_original, valor_un_original
        Dim descricao_log
        
        cod_formando = Trim(form_keys(i))
        
        If cod_formando <> "" And IsNumeric(cod_formando) Then
            
            ' Receber valores novos do formulário
            inicial_novo = formatNumber(0&Request.Form("inicial_" & cod_formando), 0)
            faturado_novo = formatNumber(0&Request.Form("faturado_" & cod_formando), 0)
            valor_un_novo = CDbl(Request.Form("valor_un_" & cod_formando))
            
            ' Receber valores originais (para comparação)
            inicial_original = formatNumber(0&Request.Form("inicial_original_" & cod_formando), 0)
            faturado_original = formatNumber(0&Request.Form("faturado_original_" & cod_formando), 0)
            valor_un_original = CDbl(Request.Form("valor_un_original_" & cod_formando))
            
            ' Validar dados
            If inicial_novo < 0 Or faturado_novo < 0 Or valor_un_novo < 0 Then
                msg_erro = "Valores não podem ser negativos"
                tem_erro = True
                Exit For
            End If
            
            ' Validar se faturado <= inicial
            If faturado_novo > inicial_novo Then
                msg_erro = "Quantidade faturada não pode ser maior que quantidade inicial"
                tem_erro = True
                Exit For
            End If
            
            ' Se houve mudança, atualizar
            If inicial_novo <> inicial_original Or faturado_novo <> faturado_original Or valor_un_novo <> valor_un_original Then
                
                ' Executar UPDATE
                sql = "UPDATE TB_FICHAS_IDENTIFICACAO SET " &_
                      "INICIAL = " & inicial_novo & ", " &_
                      "FATURADO = " & faturado_novo & ", " &_
                      "VALOR_UN = " & valor_un_novo & " " &_
                      "WHERE COD_FORMANDO = " & cod_formando
                
                objConexao.Execute(sql)
                qtd_alteracoes = qtd_alteracoes + 1
                
                ' Registrar log
                descricao_log = "Editou álbum: Inicial " & inicial_original & " → " & inicial_novo & _
                                ", Faturado " & faturado_original & " → " & faturado_novo & _
                                ", Valor R$ " & formatNumber(valor_un_original) & " → R$ " & formatNumber(valor_un_novo)
                
                Call RegistrarLogAlteracao(cod_pedido, descricao_log, objConexao, Session("cod_usuario"))
                
            End If
            
        End If
        
    Next

    ' --------- EDITAR PRODUTOS À PARTE ---------
    
    Dim cod_produtos_array
    cod_produtos_array = Request.Form("cod_pedido_produto_list")
    
    If cod_produtos_array <> "" Then
        
        Dim prod_keys, j, cod_pedido_produto, qtd_nova, valor_novo
        Dim qtd_original, valor_original
        
        prod_keys = Split(cod_produtos_array, ",")
        
        For j = LBound(prod_keys) To UBound(prod_keys)
            
            cod_pedido_produto = Trim(prod_keys(j))
            
            If cod_pedido_produto <> "" And IsNumeric(cod_pedido_produto) Then
                
                qtd_nova = formatNumber(0&Request.Form("qtd_prod_" & cod_pedido_produto), 0)
                valor_novo = CDbl(Request.Form("valor_prod_" & cod_pedido_produto))
                
                qtd_original = formatNumber(0&Request.Form("qtd_prod_original_" & cod_pedido_produto), 0)
                valor_original = CDbl(Request.Form("valor_prod_original_" & cod_pedido_produto))
                
                ' Validar dados
                If qtd_nova < 0 Or valor_novo < 0 Then
                    msg_erro = "Valores não podem ser negativos"
                    tem_erro = True
                    Exit For
                End If
                
                ' Se houve mudança, atualizar
                If qtd_nova <> qtd_original Or valor_novo <> valor_original Then
                    
                    sql = "UPDATE TB_PEDIDOS_PRODUTOS SET " &_
                          "QUANTIDADE = " & qtd_nova & ", " &_
                          "VALOR_UN = " & valor_novo & " " &_
                          "WHERE COD_PEDIDO_PRODUTO = " & cod_pedido_produto
                    
                    objConexao.Execute(sql)
                    qtd_alteracoes = qtd_alteracoes + 1
                    
                    ' Registrar log
                    descricao_log = "Editou produto à parte: Qtd " & qtd_original & " → " & qtd_nova & _
                                    ", Valor R$ " & formatNumber(valor_original) & " → R$ " & formatNumber(valor_novo)
                    
                    Call RegistrarLogAlteracao(cod_pedido, descricao_log, objConexao, Session("cod_usuario"))
                    
                End If
                
            End If
            
        Next
        
    End If

    ' ======================= FINALIZAR TRANSAÇÃO =======================
    
    If tem_erro Then
        objConexao.RollbackTrans()
        msg_erro = "Erro ao salvar: " & msg_erro
    Else
        If qtd_alteracoes > 0 Then
            objConexao.CommitTrans()
            msg_sucesso = qtd_alteracoes & " item(ns) alterado(s) com sucesso!"
        Else
            objConexao.RollbackTrans()
            msg_sucesso = "Nenhuma alteração foi detectada"
        End If
    End If

    On Error GoTo 0

End If

FecharConexao()

' ======================= REDIRECIONAR =======================

Dim url_retorno

If tem_erro Then
    url_retorno = "lca_busca_detalhamento_pedido_editar.asp?cod_pedido=" & cod_pedido & "&msg_erro=" & Server.URLEncode(msg_erro)
Else
    url_retorno = "lca_busca_detalhamento_pedido.asp?cod_pedido=" & cod_pedido & "&msg=" & Server.URLEncode(msg_sucesso)
End If

Response.Redirect(url_retorno)

'=======================================================================================
' FUNÇÃO: RegistrarLogAlteracao
' DESCRIÇÃO: Registra alterações no log de auditoria
' PARÂMETROS: cod_pedido, descricao, objConexao, cod_usuario
'=======================================================================================

Function RegistrarLogAlteracao(cod_pedido, descricao, conexao, cod_usuario)
    
    Dim sql_log
    
    ' Esta função assume que existe uma tabela de log e uma classe/função disponível
    ' Adapte conforme a estrutura do seu sistema
    
    ' Exemplo genérico:
    sql_log = "INSERT INTO TB_ADM_LOG_ALTERACOES " &_
              "(COD_PEDIDO, DESCRICAO, COD_USUARIO, DATA_CADASTRO) VALUES " &_
              "(" & cod_pedido & ", '" & Replace(descricao, "'", "''") & "', " &_
              cod_usuario & ", NOW())"
    
    On Error Resume Next
    conexao.Execute(sql_log)
    On Error GoTo 0
    
End Function

%>
