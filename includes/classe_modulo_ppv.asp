<!-- #include file="ppv_responsaveis_padrao.asp" -->
<%
'//////////////////////////////////////////////////////////////////////////
'ROTINA E FUNÇÕES PARA O MÓDULO PPV - PRODUÇÃO PÓS-VENDA


'// Function que verifica se o pedido de compra é manual ou não 

Function moduloPPVpedidoManual(cod_pedido_compra, numero_pedido_compra, conexao_aberta)

	Dim aux, sqlWhere, rs

	if not conexao_aberta Then
		AbrirConexao()
	end if 

	if cod_pedido_compra <> "" then
		sqlWhere = " cod_pedido_compra = "&sqlNumero(cod_pedido_compra)
	else
		sqlWhere = " numero_pedido_compra = "&sqlNumero(numero_pedido_compra)
	end if

	set rs = objConexao.execute("SELECT pedido_manual FROM TB_PEDIDOS_COMPRAS WHERE "&sqlWhere)
		if not rs.Eof Then
			aux = rs(0)
		end if 
	rs.close
	set rs = nothing

	if not conexao_aberta then
		FecharConexao()
	end if 

moduloPPVpedidoManual = (aux = "S")
End Function


'// Function que lê o pedido e gera as OS necessárias, retornando 1 se deu tudo certo e > 1 para erros.
Function moduloPPVgerarOrdensServico(cod_pedido, conexao_aberta)

	If not conexao_aberta Then 
		AbrirConexao()
	End If
		
	Dim rs, rs2, rs3, rs4, rs5, rs6, rsV, sql, i, ii, retorna
	Dim cod_ordem_servico, data_inicio, cod_ordem_servico_etapa_atual
	Dim CONFERIDO_POR_COD_USUARIO, ppvPadraoVenda, ppvResponsavelEtapa
	i = 1
	retorna = 1 'Por padrão é 1 sem erros
						
				'1/)--------------------------------------
				'Informações do pedido e verifica se o mesmo está corrigido para gerar OS's
				Set rs2=objConexao.Execute("SELECT DATA_VENDA, CONFERIDO_POR_COD_USUARIO, PRE_VENDA FROM TB_PEDIDOS WHERE FINALIZA_CORRECAO = 1 AND COD_PEDIDO = "&COD_PEDIDO)
				
					If not rs2.Eof Then

						'2/)--------------------------------------
						'Se já estiver corrigido lista todos os produtos vendidos à parte, cuja FLAG gera_ordem_servico_prazo_cliente for > 0 (significa que é para gerar OS)
						CONFERIDO_POR_COD_USUARIO = rs2("CONFERIDO_POR_COD_USUARIO")
						ppvPadraoVenda = PPVResponsavelPadrao(rs2("PRE_VENDA"))
						
												
							'Seleciona todos os produtos vendidos à parte e os brindes (exceto as fotos 2, 20 e 21), exceto agregados também						
							sql = "SELECT TB_PEDIDOS_PRODUTOS.COD_FORMANDO, TB_PEDIDOS_PRODUTOS.COD_PRODUTO, TB_PEDIDOS_PRODUTOS.QUANTIDADE, TB_PRODUTOS.gera_ordem_servico_prazo_cliente, TB_PRODUTOS.NOME_PRODUTO FROM TB_PRODUTOS INNER JOIN TB_PEDIDOS_PRODUTOS ON (TB_PRODUTOS.COD_PRODUTO = TB_PEDIDOS_PRODUTOS.COD_PRODUTO) WHERE TB_PRODUTOS.gera_ordem_servico_prazo_cliente > 0 AND TB_PEDIDOS_PRODUTOS.COD_PEDIDO = "&cod_pedido&" UNION ALL SELECT TB_FICHAS_IDENTIFICACAO_BRINDES.COD_FORMANDO, TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO, Sum(TB_FICHAS_IDENTIFICACAO_BRINDES.QUANTIDADE) AS QUANTIDADE, TB_PRODUTOS.gera_ordem_servico_prazo_cliente, TB_PRODUTOS.NOME_PRODUTO FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_BRINDES ON (TB_PRODUTOS.COD_PRODUTO = TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO) WHERE TB_PRODUTOS.gera_ordem_servico_prazo_cliente > 0 AND TB_FICHAS_IDENTIFICACAO_BRINDES.COD_FORMANDO IN(SELECT COD_FORMANDO FROM  TB_PEDIDOS_FORMANDOS WHERE COD_PEDIDO = "&cod_pedido&") AND TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO NOT IN(2,20,21) GROUP BY TB_FICHAS_IDENTIFICACAO_BRINDES.cod_formando, TB_FICHAS_IDENTIFICACAO_BRINDES.cod_produto"
							
							'Response.Write(sql)
							'Response.End()
							Set rs3=objConexao.Execute(sql)
										
								If not rs3.Eof Then
									
									While not rs3.Eof
											
										'1/)--------------------------------------
										'Verifica se o item já possui OS gerada, se já houver então não gera para o item específico
										Set rsV = objConexao.Execute("SELECT cod_ordem_servico FROM TB_ORDENS_SERVICO WHERE cod_pedido = "&cod_pedido&" AND cod_formando = "&rs3("COD_FORMANDO")&" AND cod_produto = "&rs3("cod_produto"))
											
											If rsV.Eof Then
											'Significa que a OS NÃO foi gerada para o produto e pedido
												'///////////////////////////////////////////////////////////////////////////
													
													
													If int(0&rs3("COD_FORMANDO")) > 0 Then
														
																	
														'4/)--------------------------------------
														'Verifica qual a etapa inicial do produto, se não houver informa erro
														Set rs4=objConexao.Execute("SELECT cod_etapa FROM TB_PRODUTOS_ETAPAS WHERE cod_produto = "&sqlTexto(rs3("cod_produto"))&" order by numero_etapa asc limit 1")
														
															If rs4.Eof Then
																retorna = 3 'Algum produto que não possui etapas cadastradas
															Else
																'5/)--------------------------------------
																'Gera uma OS para cada item									
																sql = "INSERT INTO TB_ORDENS_SERVICO (cod_pedido, cod_produto, cod_formando, numero_ordem_servico, numero_ordem_servico_indice, quantidade, cod_ordem_servico_etapa_atual, cod_usuario_processou, data_base_prazo_cliente_inicio, status_ordem_servico, prazo_cliente_total) VALUES ("&sqlTexto(cod_pedido)&", "&sqlTexto(rs3("cod_produto"))&", "&sqlTexto(rs3("cod_formando"))&", "&sqlTexto(cod_pedido)&", "&sqlTexto(i)&", "&sqlNumero(rs3("quantidade"))&", NULL, "&sqlTexto(CONFERIDO_POR_COD_USUARIO)&", "&sqlData(rs2("DATA_VENDA"))&",1,"&sqlNumero(rs3("gera_ordem_servico_prazo_cliente"))&")"
																
																objConexao.Execute(sql)
																
																																
																'Response.Write(sql&"<br/>")
																
																i=i+1
																
																'6/)--------------------------------------
																'Busca o ID gerado
																Set rs5=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_ORDENS_SERVICO LIMIT 1")
																	If not rs5.Eof Then
																		cod_ordem_servico = rs5("ULTIMO_ID")	
																	End if
																	
																rs5.Close
																Set rs5=Nothing
																
																
													
																
																
																
																'7/)--------------------------------------
																'Lista as etapas do produto
																ii = 0
																Set rs5=objConexao.Execute("SELECT * FROM TB_PRODUTOS_ETAPAS WHERE cod_produto = "&sqlTexto(rs3("cod_produto"))&" order by numero_etapa asc")
																
																	While not rs5.Eof
																																		
																		If int(ii)=0 Then
																		'1º Etapa é necessário colocar a data de início com a data de cadastro da OS, para iniciar o processo.
																			data_inicio = sqlData(now())
																		Else
																			data_inicio = "NULL"
																		End If
																	
																		ppvResponsavelEtapa = rs5("cod_usuario_responsavel")
                                                                        If ii = 0 And LCase(Trim(rs5("descricao") & "")) = "triagem" And Len(ppvPadraoVenda) > 0 Then
                                                                            ppvResponsavelEtapa = ppvPadraoVenda
                                                                        End If
                                                                        sql = "INSERT INTO TB_ORDENS_SERVICO_ETAPAS (cod_ordem_servico, cod_etapa, numero_etapa, descricao, prazo_em_dias, instrucoes, necessita_gerar_pedido_compra, cod_usuario_responsavel, data_inicio) VALUES ("&sqlTexto(cod_ordem_servico)&", "&sqlTexto(rs5("cod_etapa"))&", "&sqlTexto(rs5("numero_etapa"))&", "&sqlTexto(rs5("descricao"))&", "&sqlTexto(rs5("prazo_em_dias"))&", "&sqlTexto(rs5("instrucoes"))&", "&sqlTexto(rs5("necessita_gerar_pedido_compra"))&", "&sqlTexto(ppvResponsavelEtapa)&","&data_inicio&")"
																		
																		'Response.Write(sql&"<br />")
																		objConexao.Execute(sql)
																														
																		
																		If int(ii)=0 Then
																		'Busca o cod_ordem_servico_etapa_atual (1ª no caso).
																				sql = "SELECT cod_ordem_servico_etapa FROM TB_ORDENS_SERVICO_ETAPAS WHERE cod_ordem_servico = "&cod_ordem_servico&" ORDER BY numero_etapa ASC LIMIT 1"
																				
																				Set rs6=objConexao.Execute(sql)
																					If not rs6.Eof Then
																						
																						sql = "UPDATE TB_ORDENS_SERVICO SET cod_ordem_servico_etapa_atual = "&rs6("cod_ordem_servico_etapa")&" WHERE cod_ordem_servico = "&cod_ordem_servico
																						
																						
																						
																						'Response.Write(sql&"<br />")
																						
																						objConexao.Execute(sql)	
																																												
																					End if 
																				rs6.Close
																				Set rs6=Nothing	
																		
																		End If
																		
																		ii=ii+1
																		
																		
																	rs5.MoveNext()
																	Wend
																	
																rs5.Close
																Set rs5=Nothing
																

																call logEventos("[cod_ordem_servico:"&cod_ordem_servico&"][cod_pedido:"&cod_pedido&"][cod_usuario:"&session("cod_usuario")&"][INFO:CRIACAO DE OS]", true)

															End If
														
														rs4.Close
				
														Set rs4=Nothing								
													
													End If
													
												'///////////////////////////////////////////////////////////////////////////
											End If
											
										rsV.Close
										Set rsV = Nothing
									
									rs3.MoveNext()
									Wend
									
								End If
							
							rs3.Close
							Set rs3=Nothing
								
					End If
					
	rs2.Close
	Set rs2=Nothing									

	If retorna > 1 Then
	'Limpa OS's, por segurança se foi gerado para o pedido somente se estiver com status 1
		
		'NÃO EXCLUI AGORA
	
		'Excluir Etapas
		'objConexao.Execute("SELECT cod_ordem_servico FROM TB_ORDENS_SERVICO_ETAPAS WHERE cod_ordem_servico IN(SELECT cod_ordem_servico FROM TB_ORDENS_SERVICO WHERE cod_pedido = "&cod_pedido&" and status_ordem_servico = 1)")
		'Excluir Histórico
		'objConexao.Execute("SELECT cod_ordem_servico FROM TB_ORDENS_SERVICO_HISTORICO WHERE cod_ordem_servico IN(SELECT cod_ordem_servico FROM TB_ORDENS_SERVICO WHERE cod_pedido = "&cod_pedido&" and status_ordem_servico = 1)")
		'Excluir OS
		'objConexao.Execute("DELETE FROM TB_ORDENS_SERVICO WHERE status_ordem_servico = 1 and cod_pedido = "&cod_pedido)
				
	End If
	
	
	If not conexao_aberta Then 
		FecharConexao()
	End If		
		
	moduloPPVgerarOrdensServico = retorna

End Function

'Função que lê o código de retorno da função de inserção de OS e retorna com o erro correspondente.
Function moduloPPVgerarOrdensServicoMsgRetorno(cod_mensagem_retorno,cod_pedido)

	Dim retorno
	
	If int(cod_mensagem_retorno) = 1 Then
		retorno = "OSs geradas com sucesso!"
	ElseIf int(cod_mensagem_retorno) = 2 Then
		retorno = "OSs NÃO geradas. Motivo: OSs já geradas anteriormente!"
	ElseIf int(cod_mensagem_retorno) = 3 Then
		retorno = "OSs NÃO geradas. Motivo: Algum produto que não possui etapas cadastradas!"
	End If
	
	moduloPPVgerarOrdensServicoMsgRetorno = retorno
	
End Function

'Dim msg_retorno
'msg_retorno = moduloPPVgerarOrdensServico(22300,false)
'Response.Write(moduloPPVgerarOrdensServicoMsgRetorno(msg_retorno,22300))
'msg_retorno = moduloPPVgerarOrdensServico(24100,false)
'Response.Write(moduloPPVgerarOrdensServicoMsgRetorno(msg_retorno,24100))

'Se informar: cod_ordem_servico, então retorna com mascara
'Se informar: ordem_servico_com_mascara, então retorna sem mascara

'Ex com máscara: 22300.016.2 [NUMERO_ORDEM_SERVICO].[COD_PRODUTO].[INDICE]
Function moduloPPVOSMascara(cod_ordem_servico_,ordem_servico_com_mascara,conexao_aberta)

	If not conexao_aberta Then 
		AbrirConexao()
	End If
	
	Dim rs, rs2, sql, temp, i, retorna
	Dim numero_ordem_servico, cod_produto, numero_ordem_servico_indice
	
	'Trata valor inserido
	Dim cod_ordem_servico
	cod_ordem_servico = Cstr(""&cod_ordem_servico_)
	
	If cod_ordem_servico <> "" Then
	'Colocar a mascara
	
		Set rs=objConexao.Execute("SELECT numero_ordem_servico, cod_produto, numero_ordem_servico_indice FROM TB_ORDENS_SERVICO WHERE cod_ordem_servico = "&cod_ordem_servico)
			
			If not rs.Eof Then
				
				numero_ordem_servico = preencheCaracter(rs("numero_ordem_servico"),"0",5)
				cod_produto = preencheCaracter(rs("cod_produto"),"0",3)
				numero_ordem_servico_indice = rs("numero_ordem_servico_indice")
				
				retorna = Cstr(""&numero_ordem_servico)&"."&cod_produto&"."&numero_ordem_servico_indice
				
			End If
			
		rs.Close
		Set rs=Nothing
	
	Elseif ordem_servico_com_mascara <> "" Then
	'Retirar a mascara
		temp = split(ordem_servico_com_mascara,".")
		'[NUMERO_ORDEM_SERVICO].[COD_PRODUTO].[INDICE]
			
			Set rs=objConexao.Execute("SELECT cod_ordem_servico FROM TB_ORDENS_SERVICO WHERE numero_ordem_servico = "&sqlTexto(temp(0))&" AND cod_produto = "&sqlTexto(temp(1))&" AND numero_ordem_servico_indice = "&sqlTexto(temp(2)))
				If not rs.Eof Then
					retorna = rs(0)
				End If
				
			rs.Close
			Set rs=Nothing
	
	End If
	

	If not conexao_aberta Then 
		FecharConexao()
	End If
	
	moduloPPVOSMascara = retorna

End Function

'Retorna o tipo de status
Function moduloPPVStatus(status_ordem_servico)
		
	Dim retorna
	
	If status_ordem_servico = 1 Then
		retorna = "Em produção"
	Elseif status_ordem_servico = 2 Then
		retorna = "Finalizados e não despachados"
	Elseif status_ordem_servico = 3 Then
		retorna = "Despachado para cliente"
	Elseif status_ordem_servico = 99 Then
		retorna = "Cancelados"
	End If

	moduloPPVStatus = retorna
	
End Function
%>