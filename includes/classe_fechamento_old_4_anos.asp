<%
'//////////////////////////////////////////////////////////////////////////
'ROTINA E FUNÇÕES PARA O MÓDULO DE FECHAMENTOS

Function fechamentoRetornaCodPeriodoVerificaErro(cod_fechamento_periodo)

	Dim retorna
	
	If cod_fechamento_periodo > 0 Then
	'Não há erro foi encontrado o cód de fechamento
		retorna = true
	Elseif cod_fechamento_periodo = -1 or cod_fechamento_periodo = 0 Then
		retorna = "Erro período não encontrado de acordo com a data da venda."
	Elseif cod_fechamento_periodo = -2 Then
		retorna = "Erro este pedido está programado para um pagamento que já ocorreu e o sistema não conseguiu encontrar um período recente para o incluir o pedido. Peça para o responsável pelo cadastramento de Períodos de Vendas incluir novos registros"
	Elseif cod_fechamento_periodo = -3 Then
		retorna = "Não foi possível encontrar período de fechamento devido ..."
	Elseif cod_fechamento_periodo = -4 Then
		retorna = "Não foi possível encontrar período de fechamento devido ..."
	Elseif cod_fechamento_periodo = -5 Then
		retorna = "Não foi possível encontrar período de fechamento devido ..."
	End If
	
	fechamentoRetornaCodPeriodoVerificaErro = retorna
	
End Function


'Rotina que calcula o cód perído e retorna o mais adequado para o período de acordo com a data atual
Function fechamentoRetornaCodPeriodo(cod_pedido,conexao_aberta)

	If not conexao_aberta Then
		AbrirConexao()
	End If
	
	'Variáveis primarias
	Dim rs, sql, temp, temp2
	Dim data_venda, data_cadastro, forma_pagamento, qtd_parcelas, cod_fechamento_periodo, verifica_comissao
	
	'Zera para iniciar
	cod_fechamento_periodo = 0
	
	'Variáveis do período
	Dim data_inicio, data_termino, data_termino_lancamento_vendas, data_pagamento, data_processado_pagamento
	
	'Busca as informações do pedido para efetuar os cálculos
	'//////////////
	sql = "SELECT data_venda, data_cadastro, forma_pagamento, qtd_parcelas FROM TB_PEDIDOS WHERE COD_PEDIDO = "&cod_pedido
			
	Set rs=objConexao.Execute(sql)
				
		If not rs.Eof Then
			data_venda		= formatDateTime(rs("data_venda"))
			data_cadastro	= formatDateTime(rs("data_cadastro"))
			forma_pagamento = rs("forma_pagamento")
			qtd_parcelas	= formatNumber(0&rs("qtd_parcelas"),0)
		Else
			Response.Write("ERRO PEDIDO NÃO ENCONTRADO!!!")
		End If
		
	rs.Close
	Set rs=Nothing
	
	
	'#########################################################################################################################
	'### 1/4 - COD_FECHAMENTO_PERIODO PADRÃO
	'### ETAPA 1 - Possui período de Vendas de acordo com a data da venda
	'#########################################################################################################################
			
			
			' 1) Verifica qual seria o período comum se não houvesse problemas com a venda. Data referência é a data da venda
			'Neste caso já está aberto a conexão

			cod_fechamento_periodo = int(fechamentoRetornaCodPeriodoEmRelacaoData(1,data_venda,true))
			'///////////////////////////////////////////////////////////////////////////////////////
			
			'2) Agora continua somente se foi encontrado pelo menos o período padrão	
			If cod_fechamento_periodo < 1 Then
							
						'Não encontrou nem período padrão!!!
						'Busca a próxima data válida para pagamento
							
						Set rs=objConexao.Execute("SELECT cod_fechamento_periodo FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE data_processado_pagamento IS NULL ORDER BY data_pagamento ASC ")
						
							If not rs.Eof Then
								cod_fechamento_periodo = rs(0)
							Else
								Response.Write("ERRO PERÍODO MAIS PRÓXIMO NÃO ENCONTRADO, FAVOR CADASTRAR CALENDÁRIO DE VENDAS URGENTE")
							End If
							
						rs.Close
						Set rs=Nothing
						
						msg_erro_readequacao_vendedor_recebe_em = "Não foi encontrado período válido para incluir venda em pagamento, por tanto o pedido foi realocado para o pagamento mais atual e mais próximo."
				
			End If	
			
				'Obs.: Neste ponto sempre estará aberta a conexão
				'////////////
				data_inicio 					= fechamentoRetornaData(1,cod_fechamento_periodo,true)
				data_termino 					= fechamentoRetornaData(2,cod_fechamento_periodo,true)
				data_termino_lancamento_vendas	= fechamentoRetornaData(3,cod_fechamento_periodo,true)
				data_pagamento					= fechamentoRetornaData(4,cod_fechamento_periodo,true)
				data_processado_pagamento		= fechamentoRetornaData(5,cod_fechamento_periodo,true)
			
				'Response.Write("Período selecionado incialmente :"&cod_fechamento_periodo&"<br />")
				temp = DateDiff("d",data_termino_lancamento_vendas,data_cadastro)				
	
	'###########################################################################################
	'### 2/4 - COD_FECHAMENTO_PERIODO PARA QUANDO EXCEDER O LIMITE DA DATA DE CADASTRO
	'### ETAPA 2 - Lançado até termino do prazo estabelecido para o período encontrado
	'###########################################################################################
				

				If temp > 0 Then 'Período excedeu limite de data de cadastro
						
						'/////////////////////////////
						'Busca novo período em relação a data de cadastro
						cod_fechamento_periodo = fechamentoRetornaCodPeriodoEmRelacaoData(1,data_cadastro,true)
						
										
						'////////////
						'Atualiza as informações novamente em relação a este novo período
						data_inicio 					= fechamentoRetornaData(1,cod_fechamento_periodo,true)
						data_termino 					= fechamentoRetornaData(2,cod_fechamento_periodo,true)
						data_termino_lancamento_vendas	= fechamentoRetornaData(3,cod_fechamento_periodo,true)
						data_pagamento					= fechamentoRetornaData(4,cod_fechamento_periodo,true)
						data_processado_pagamento		= fechamentoRetornaData(5,cod_fechamento_periodo,true)
						
						msg_erro_readequacao_vendedor_recebe_em = "Pedido realocado para pagamento para "&ArrumaData(data_pagamento)&" devido ter excedido o limite estabelecido do período em "&temp&" dia(s)."
						
				
				End If
				
				
				
				If len(data_processado_pagamento) > 2 Then				
	'###########################################################################################
	'### 3/4 - COD_FECHAMENTO_PERIODO PARA QUANDO O PAGAMENTO PADRÃO JÁ FOI PROCESSADO
	'###########################################################################################
				
				'4) Se a data do processamento do pagamento já tiver ocorrido então não poderá mais utilizar este período deverá ser o próximo
				'Obs.: Agora o sistema deverá verificar qual o próximo pagamento em relação a este selecionado para 
							
				'Pagamento já processado tenta o pagamento mais próxima superior ao pagamento encontrado
				'Para isso busca um período que não tenha sido processado o pagamento, pois nem sempre o próximo quer dizer que não foi processado. No caso verifica o último processamento de pagamento
				
					sql = "SELECT data_termino FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS order by data_processado_pagamento desc limit 1"
					
					Set rs=objConexao.Execute(sql)
					
						If not rs.Eof Then
							temp = formatDateTime(rs(0))
						End If
						
					rs.Close
					Set rs=Nothing
				
					'Incrementa em 1 dia em relação a data de referência encontrada do último pagamento ocorrido
					temp = DateAdd("d",1,temp)
				
					cod_fechamento_periodo = fechamentoRetornaCodPeriodoEmRelacaoData(1,temp,true)
				
					If cod_fechamento_periodo = -1 Then
						'Muda para o erro -2, pois o sistema está tentanto colocar o pedido em pagamento que não foi processado, mas não é possível, pois não foi encontrado
						cod_fechamento_periodo = -2
					End If
				
				
					'////////////
					'Atualiza as informações novamente em relação a este novo período
					data_inicio 					= fechamentoRetornaData(1,cod_fechamento_periodo,true)
					data_termino 					= fechamentoRetornaData(2,cod_fechamento_periodo,true)
					data_termino_lancamento_vendas	= fechamentoRetornaData(3,cod_fechamento_periodo,true)
					data_pagamento					= fechamentoRetornaData(4,cod_fechamento_periodo,true)
					data_processado_pagamento		= fechamentoRetornaData(5,cod_fechamento_periodo,true)
							
					msg_erro_readequacao_vendedor_recebe_em = "Pedido realocado para pagamento para "&ArrumaData(data_pagamento)&" devido ter o pagamento correto para este pedido já ter sido processado"
					
					 
				End If
			
	
	'###########################################################################################
	'### 4/4 - COD_FECHAMENTO_PERIODO ESTÁ DE ACORDO COM A ENTRADA
	'###########################################################################################
			If cod_fechamento_periodo > 0 Then
				
			'Nesta etapa pode ter alterado o perído no máximo.
			'1) cod_fechamento_periodo da etapa 1
			'2) cod_fechamento_periodo da etapa 2 ou seja um superior ao padrão
			'3) cod_fechamento_periodo da etapa 3 que poderá ser superior somente a etapa 1 ou superior a etapa 1 e 2 também
			'De qualquer forma em todas as situações 1, 2 e 3 deverá verificar se neste período novo encontrado não há problemas de entrada, pois caso contrário deverá sem alterado novamente o pagamento
					
					temp = data_pagamento
					'Response.Write(retornaPedidoPossuiEntradaEmRelacaoDataInformada(temp,cod_pedido,true))
					
					
					If not retornaPedidoPossuiEntradaEmRelacaoDataInformada(temp,cod_pedido,true) Then
					
					
					'Significa que não foi encontrado entrada para a data de pagamento atual, por tanto neste caso
					'somente há possibilidade de não ter dinheiro+cartão superior a <valor_considerado_entrada>
					'e há algum cheque com o menor BOM_PARA superior a data do pagamento
							
							'Por tanto busca este menor cheque:
							
							'Possibilidades 2
							'1) Se tiver cheque significa se está pré-datado para uma superior basta procurar um pagamento em relação esta data
							'2) Se não houver cheque significa que é um pedido com valor inverior ao <valor_considerado_entrada> como é o caso de pedidos de comissões e por tanto NÃO deverá ser alterado para o pagamento do mês seguinte
							
							Set rs2 = objConexao.Execute("SELECT BOM_PARA FROM TB_PEDIDOS_CHEQUE WHERE COD_PEDIDO ="&cod_pedido&" ORDER BY BOM_PARA ASC LIMIT 1")
										
									If not rs2.Eof Then
										temp = FormatDateTime(rs2("BOM_PARA"))
										
										'Encontrou por tanto possibilidade "1" (descrição acima)
										cod_fechamento_periodo = fechamentoRetornaCodPeriodoEmRelacaoData(2,temp,true)
										
										'Response.Write(">>>"&cod_fechamento_periodo&"<<<")
										
										If cod_fechamento_periodo > 0 Then
											'////////////
											'Atualiza as informações novamente em relação a este novo período
											data_inicio 					= fechamentoRetornaData(1,cod_fechamento_periodo,true)
											data_termino 					= fechamentoRetornaData(2,cod_fechamento_periodo,true)
											data_termino_lancamento_vendas	= fechamentoRetornaData(3,cod_fechamento_periodo,true)
											data_pagamento					= fechamentoRetornaData(4,cod_fechamento_periodo,true)
											data_processado_pagamento		= fechamentoRetornaData(5,cod_fechamento_periodo,true)										
											
											
											msg_erro_readequacao_vendedor_recebe_em = "Pedido realocado para pagamento para "&ArrumaData(data_pagamento)&" devido não possuir entrada no valor mínimo."
											
										End If
										
										
									Else
									'Não encontrou por tanto possibilidade "2" (descrição acima)
											
										'Neste caso não tenta encontrar outra data de pagamento, pois é um pedido com valor baixo.
										
									End If
							
							rs2.Close
							Set rs2 = Nothing
							
					End If
					
		
				
			End If
			
			
			
			
					'Response.Write("<br /><br />Informações do período que será utilizado<br />---------------------------------------------------------<br />")
			
					'Response.Write("Período selecionado: "&cod_fechamento_periodo&"<br />")
					'Response.Write("Data data_inicio: "&data_inicio&"<br />")
					'Response.Write("Data data_termino: "&data_termino&"<br />")
					'Response.Write("Data data_termino_lancamento_vendas: "&data_termino_lancamento_vendas&"<br />")
					'Response.Write("Data data_pagamento: "&data_pagamento&"<br />")
					'Response.Write("Data data_processado_pagamento: "&data_processado_pagamento&"<br />")			
					'Response.Write("--------------------------------------------------------- <br />")
			

	If not conexao_aberta Then
		FecharConexao()
	End If
	
	fechamentoRetornaCodPeriodo = cod_fechamento_periodo
	
End Function

'Função que retorna a forma de pagamento automática selecionada pelo sistema, atenção só deverá ser utilizada
'antes da finalização do pedido, pois efetua cálculos para finalização da correção do pedido e depois perdem o sentido.
'Retorna Formas de Pagamento
'1 = A vista
'2 = Parcelado
Function fechamentoRetornaFormaPagamentoPedido(data_venda,cod_pedido,conexao_aberta)
	
	If not conexao_aberta Then
		AbrirConexao()
	End If
	
	Dim rs, sql, temp, temp2, temp3
	Dim cod_fechamento_periodo, data_pagamento
	Dim retorna
	retorna = 0
					
						
						If Len(data_venda) < 2 Then
						
							'Busca a data da venda, pois no caso do vendedor lançar um pedido acima do limite da data de pagamento, poderá ser alterada erroneamente para à vista uma vez que passou o pagamento e o cheque por sua vez não ultrapassa a data do próximo pagamento.	
							'Para não corre o risco procura o próximo pagamento no qual o vendedor se baseou no momento da venda.
							Set rs = objConexao.Execute("SELECT data_venda FROM TB_PEDIDOS WHERE COD_PEDIDO="&cod_pedido)
									
								If not rs.Eof Then
									data_venda = formatDateTime(rs(0))
								End If
									
									
							rs.Close
							Set rs = Nothing
						
						End If
							
							
						'Busca um cod_fechamento_periodo de acordo com a data da venda, mesmo que este fechamento não possa ser utilizado pois já foi processado ou algo do genero, por servir apenas como base no momento da verificação dos cheques
						cod_fechamento_periodo = fechamentoRetornaCodPeriodoEmRelacaoData(1,data_venda,true)
											
						'Verifica se houve erro na busca e no caso de erro já responde através da função qual foi o erro
						temp = Cstr(""&fechamentoRetornaCodPeriodoVerificaErro(cod_fechamento_periodo))
						If temp <> "True"  Then
							
							'Não encontrou nem período padrão!!!
							'Busca a próxima data válida para pagamento
							
							Set rs=objConexao.Execute("SELECT cod_fechamento_periodo FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE data_processado_pagamento IS NULL ORDER BY data_pagamento ASC ")
				
								If not rs.Eof Then
									cod_fechamento_periodo = rs(0)
								End If
							
							rs.Close
							Set rs=Nothing
						
						End If
						
						
						data_pagamento = fechamentoRetornaData(4,cod_fechamento_periodo,true)
						
						'Response.Write(">>>"&data_pagamento)
						
						'//////////////////////////////////////////////////////////////////////
							'à vista (1) até que prove o contrário
							forma_pagamento = 1
							
								'Verifica se essas parcelas são na verdade titulares diferentes, com vencimento para até o dia 15
								'1) Cartão
								Set rs = objConexao.Execute("SELECT PARCELAS FROM TB_PEDIDOS_CARTAO WHERE COD_PEDIDO = "&cod_pedido)
										
										While not rs.Eof
											If int(rs("PARCELAS")) > 1 Then
											'Se tiver pelo menos 1 cartão do mesmo pedido com mais que 1 parcela automaticamente
											'irá entender como parcelado independente se é crédito ou débito									
												forma_pagamento = 2
											End If
											
										rs.MoveNext()
										Wend
										
								rs.Close
								Set rs = Nothing
								
								
								'2) Cheque
								Set rs = objConexao.Execute("SELECT BOM_PARA FROM TB_PEDIDOS_CHEQUE WHERE COD_PEDIDO = "&COD_PEDIDO)
										
										While not rs.Eof
											
												temp2 = formatDateTime(rs("BOM_PARA"))
												
												'Data limite do próprio pagamento selecionado acima com a diferença em relação
												'ao menor 
												temp = DateDiff("d",temp2,data_pagamento)										
												'Response.Write("Pedido: "&COD_PEDIDO&" : "&VENDEDOR_RECEBE_EM_CONFERIDO)
												
												If int(temp) < 0 Then
												'Siginifica que há pelo menos 1 parcela a prazo para depois do pagamento, não é considerado à vista.
														forma_pagamento = 2											
												End If
												
											
										rs.MoveNext()
										Wend
										
								rs.Close
								Set rs = Nothing

								'3) Boleto'
								Set rs = objConexao.Execute("SELECT count(*) as parcelas FROM TB_PEDIDOS_BOLETOS WHERE COD_PEDIDO = "&cod_pedido)
								
								While not rs.Eof
									If int(""&rs("parcelas")) > 1 Then
									'Se tiver mais de um boleto gerado irá setar como parcelado
										forma_pagamento = 2
									End If
									
								rs.MoveNext()
								Wend
										
								rs.Close
								Set rs = Nothing


							'//////////////////////////////////////////////////////////////////////
							retorna = forma_pagamento
						

	If not conexao_aberta Then
		FecharConexao()
	End If
	
		
	fechamentoRetornaFormaPagamentoPedido = retorna
		
End Function


'Função que verifica se o pedido possui entrada de pelo menos R$ 50,00 podendo ser alterado no futuro
'Se tiver retorna TRUE se não tiver retorna FALSE
'Tudo em relação a uma data de pagamento
Function retornaPedidoPossuiEntradaEmRelacaoDataInformada(data_pagamento,cod_pedido,conexao_aberta)


	If not conexao_aberta Then
		AbrirConexao()
	End If

	Dim rs, rs2, sql, retorna, temp, temp2
	Dim valor_considerado_avista
	Dim total_em_dinheiro, total_cartao, total_cheque
	
	valor_considerado_avista = 50 'Basta alterar este valor para alterar o padrão aceito no valor de entrada
	
	'Por padrão todas as vendas possuem entrada até que seja provado ao contrário
	retorna = true
	
	'1/4 --------------------------------------------------------------------------
	'Busca as informações dos valores totais do pedido
	
			Set rs=objConexao.Execute("SELECT PAGAMENTO_EM_DINHEIRO, XTOTAL_CHEQUE, XTOTAL_CARTAO FROM VW_PEDIDOS_TOTAL_RECEBIDO WHERE COD_PEDIDO = "&cod_pedido&" LIMIT 1")
				
				If not rs.Eof Then
					total_em_dinheiro	= (formatNumber(0&rs("PAGAMENTO_EM_DINHEIRO")))/1
					total_cartao 		= (formatNumber(0&rs("XTOTAL_CARTAO")))/1
					total_cheque 		= (formatNumber(0&rs("XTOTAL_CHEQUE")))/1
				End If
			
			rs.Close
			Set rs=Nothing
	
	'2/4 --------------------------------------------------------------------------
	'Verifica se há entrada e há pagamento em cheque

								temp = 0
								
								'Cartão e dinheiro são considerados à vista de imediato
								temp = (total_cartao+total_em_dinheiro)/1
								
								If temp < valor_considerado_avista Then
								'Não há entrada por tanto deverá verificar se há cheque à vista
								
									'Neste ponto não há mais entrada encontrada
									'Agora deverá verificar se possui entrada
									retorna = false
								
									If total_cheque > 0 Then
									'Significa que há pagamento em cheque, por tanto verifica-se se há pagamento à vista
															
										'3/4 --------------------------------------------------------------------------				
										'Visualiza o cheque com a data menor
										Set rs2 = objConexao.Execute("SELECT BOM_PARA FROM TB_PEDIDOS_CHEQUE WHERE COD_PEDIDO ="&cod_pedido&" ORDER BY BOM_PARA ASC LIMIT 1")
										
												If not rs2.Eof Then
													temp = FormatDateTime(rs2("BOM_PARA"))
													
													temp2 = DateDiff("d",data_pagamento,temp)											
														
													If temp2 <= 0 Then
														
														'4/4 -------------------------------------------------------------------
														'O menor cheque está inferior ou igual a data do pagamento, é considerado entrada.
														retorna = true										
													End If
												End If
											
										
										rs2.Close()
										Set rs2 = Nothing
										
									End If
									
								End If
						

	If not conexao_aberta Then
		FecharConexao()
	End If
	
	retornaPedidoPossuiEntradaEmRelacaoDataInformada = retorna

End Function

'Função que retorna a data de acordo com cod_fechamento_periodo e tipo
'tipo 1 = data_inicio
'tipo 2 = data_termino
'tipo 3 = data_termino_lancamento_vendas
'tipo 4 = data_pagamento
'tipo 5 = data_processado_pagamento
Function fechamentoRetornaData(tipo,cod_fechamento_periodo,conexao_aberta)

	Dim rs, sql, retorna
	
	
	retorna = ""
	
	If not conexao_aberta Then
		AbrirConexao()
	End If

	sql = "SELECT * FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE cod_fechamento_periodo = "&sqlTexto(cod_fechamento_periodo)
	
	Set rs=objConexao.Execute(sql)
				
			If not rs.Eof Then
				
				If tipo = 1 Then
					If isDate(rs("data_inicio")) Then
						retorna = formatDateTime(rs("data_inicio"))
					End If
				ElseIf tipo = 2 Then
					If isDate(rs("data_termino")) Then
						retorna = formatDateTime(rs("data_termino"))
					End If				
				ElseIf tipo = 3 Then
					If isDate(rs("data_termino_lancamento_vendas")) Then
						retorna = formatDateTime(rs("data_termino_lancamento_vendas"))
					End If				
				ElseIf tipo = 4 Then
					If isDate(rs("data_pagamento")) Then
						retorna = formatDateTime(rs("data_pagamento"))
					End If				
				ElseIf tipo = 5 Then
					If isDate(rs("data_processado_pagamento")) Then
						retorna = formatDateTime(rs("data_processado_pagamento"))
					End If
				End If
			End If
			
	rs.Close
	Set rs=Nothing
	


	If not conexao_aberta Then
		FecharConexao()
	End If
	
	fechamentoRetornaData = retorna

End Function


'Função que retorna qual o período de acordo com a data de referencia, obs não faz validações somente verifica o período de acordo com a data referencia que poderá ser diversas datas como data venda etc...
'tipo = 1 referencia pela data da venda
'tipo = 2 referencia pela data do pagamento (utilizado em consultas de cheques)
Function fechamentoRetornaCodPeriodoEmRelacaoData(tipo,data_referencia,conexao_aberta)

		Dim rs, sql, temp, temp2
		Dim data_inicio, data_termino, data_pagamento, cod_fechamento_periodo

		
		If not conexao_aberta Then
			AbrirConexao()
		End If
		
		If tipo = 1 Then
		
				'1.1) Verifica os períodos do mês de acordo com a data base
				'Normalmente há 1 ou 2 períodos por mês e não é permitido ter meses diferentes no período.
				'Em cada período verifica de acordo com o intervalo de tempo qual 
				sql = "SELECT * FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE MONTH(data_inicio) = "&MONTH(data_referencia)&" AND YEAR(data_inicio) = "&YEAR(data_referencia)
					
				Set rs=objConexao.Execute(sql)
				
				'Response.Write("Data Referencia: "&data_referencia&"<br />")
				
				If not rs.Eof Then
					
					While not rs.Eof
						
						data_inicio		= formatDateTime(rs("data_inicio"))	
						data_termino	= formatDateTime(rs("data_termino"))
						temp			= DateDiff("d",data_inicio,data_termino)
						
						'Response.Write("Início ---------------------------------------<br />")
						
						For i=0 to temp
						
							'Incrementa a data de início
							temp2 = DateAdd("d",i,data_inicio)
									
							'Compara se a data da venda é igual a data do laço ou seja está dentro do intervalo deste período			
							If (DateDiff("d",data_referencia,temp2) = 0) Then
								cod_fechamento_periodo = rs("cod_fechamento_periodo")
								'Response.Write("Encontrado >>> ")
							End If
							
							'Response.Write(temp2&"<br />")
												
						Next
						'Response.Write("<br />Fim ---------------------------------------<br />")
						
					rs.MoveNext()
					Wend
					
					
				Else
					cod_fechamento_periodo = -1 'Erro período não encontrado de acordo com a data da venda
				End If
				
			rs.Close
			Set rs=Nothing
		
		Elseif tipo = 2 Then
			
			'Tipo 2 busca em relação a data do pagamento verifica o mais próximo que não tenha sido processado
			sql = "SELECT cod_fechamento_periodo,data_pagamento,DATEDIFF(data_pagamento,"&sqlData(data_referencia)&") as diferenca FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE DATEDIFF(data_pagamento,"&sqlData(data_referencia)&") >= 0 AND data_processado_pagamento is null ORDER BY DATEDIFF(data_pagamento,"&sqlData(data_referencia)&") ASC LIMIT 1"
			
				Set rs=objConexao.Execute(sql)
		
					If not rs.Eof Then
						cod_fechamento_periodo = rs("cod_fechamento_periodo")
					Else
						cod_fechamento_periodo = -1
					End If
					
				rs.Close
				Set rs=Nothing
			
			
		End If
		
		
	If not conexao_aberta Then
		FecharConexao()
	End If

	fechamentoRetornaCodPeriodoemRelacaoData = cod_fechamento_periodo

End Function 

Function fechamentoRetornaTipoVenda(cod_pedido,conexao_aberta, cod_formando)

	If not conexao_aberta Then
		AbrirConexao()
	End If

	Dim rs, sql, tipo_venda, retorna, semestre_ano_conclusao_completo
	Dim TRANS_DATA_CONCLUSAO, TRANS_DATA_VIRA_REPASSE, TRANS_DATA_VENDA, DIFERENCA_DATA, DATA_VENDA
	'Tipos de Vendas:
	'1 Normal
	'2 Repasse
	'3 Repasse Especial 'INUTILIZADO DEVIDO AS TABELAS DE COMISSÕES FOREM SEPARADAS POR VENDEDORES
	
	'Verifica a data da venda
			sql = "SELECT DATA_VENDA FROM TB_PEDIDOS WHERE COD_PEDIDO = "&COD_PEDIDO
			
			Set rs=objConexao.Execute(sql)
				
				If not rs.Eof Then
					DATA_VENDA	= formatDateTime(rs("DATA_VENDA"))
				End If
			
			rs.Close
			Set rs=Nothing
				
		
	'Verifica a MENOR data de conclusão do pedido
	'Regra de que se houver mais que um formando sendo 1 repasse e o outro não, deverá contar como os 2 repasse
			sql = "SELECT DATE(MIN(SEMESTRE_ANO_CONCLUSAO_COMPLETO)) as SEMESTRE_ANO_CONCLUSAO_COMPLETO_MENOR FROM VW_PEDIDOS_FORMANDOS  WHERE COD_PEDIDO = "&cod_pedido&" and cod_formando = "&cod_formando
			
			'response.write(sql)
			Set rs = objConexao.Execute(sql)
				
				If not rs.Eof Then
									
					TRANS_DATA_CONCLUSAO	= formatDateTime(rs("SEMESTRE_ANO_CONCLUSAO_COMPLETO_MENOR"))
					TRANS_DATA_VIRA_REPASSE	= DateAdd("yyyy",4,TRANS_DATA_CONCLUSAO)
					TRANS_DATA_VENDA 		= formatDateTime(Year(DATA_VENDA)&"-"&Month(DATA_VENDA)&"-1")
									
					
					DIFERENCA_DATA = DateDiff("d",TRANS_DATA_VIRA_REPASSE,TRANS_DATA_VENDA)
					
					'Response.Write(TRANS_DATA_CONCLUSAO&"<br />")
					'Response.Write(TRANS_DATA_VIRA_REPASSE&"<br />")
					'Response.Write(TRANS_DATA_VENDA&"<br />")
							
					If DIFERENCA_DATA > -1 Then
					'Repasse
						TIPO_VENDA = 2
					Else
					'Normal
						TIPO_VENDA = 1
							
					End If
				
				Else
					TIPO_VENDA = 0
					Response.Write("ERRO PEDIDO NÃO ENCONTRADO!!!!")
				End If
				
			rs.Close
			Set rs = Nothing
	
	retorna = tipo_venda
	
	If not conexao_aberta Then
		FecharConexao()
	End If
	
	fechamentoRetornaTipoVenda = retorna
	
End Function

'Função que cálcula o valor total de comissão de um produto em um pedido
Function fechamentoRetornaValorComissaoPedidoProduto(cod_pedido,cod_produto,quantidade,valor_total,aliquota_desconto, conexao_aberta, cod_formando)

	If not conexao_aberta Then
		AbrirConexao()
	End If	
	
	Dim rs, rs2, rs3, sql, temp, temp2, retorna
	Dim cod_tabela_comissao_vendedores, forma_pagamento, tipo_venda, data_venda
	Dim aux_desc
	aux_desc = formatNumber(aliquota_desconto)
	retorna = 0
	
	sql = "SELECT TB_PEDIDOS.FORMA_PAGAMENTO, TB_PEDIDOS.TIPO_VENDA, TB_PEDIDOS.DATA_VENDA, TB_VENDEDORES.cod_tabela_comissao_vendedores FROM TB_VENDEDORES INNER JOIN TB_PEDIDOS ON TB_VENDEDORES.COD_VENDEDOR = TB_PEDIDOS.COD_VENDEDOR WHERE TB_PEDIDOS.COD_PEDIDO = "&cod_pedido
	Set rs=objConexao.Execute(sql) 
		
		If not rs.Eof Then
			cod_tabela_comissao_vendedores = ""&rs("cod_tabela_comissao_vendedores")
			forma_pagamento = formatNumber(0&rs("forma_pagamento"),0)
			tipo_venda = formatNumber(0&rs("tipo_venda"),0)
			data_venda = rs("data_venda")

			If int(forma_pagamento) < 1 and isDate(data_venda) Then
				'Cálcula a forma de pagamento
				forma_pagamento = fechamentoRetornaFormaPagamentoPedido(data_venda,cod_pedido,true)
			End if
			
			'If int(tipo_venda) < 1 Then
				'Cálcula o tipo da venda
				'tipo_venda = fechamentoRetornaTipoVenda(cod_pedido,true, cod_formando)
			'End If
			
			'Sempre buscamos o tipo da venda pois o tipo da venda passa a ser por formando do pedido e não mais por pedido
			tipo_venda = fechamentoRetornaTipoVenda(cod_pedido,true, cod_formando)



			'//////////////////////////////////////////////////////
			
			'Verifica novamente se conseguiu processar valores novos
			If int(forma_pagamento) > 0 and int(tipo_venda) > 0 Then
				
				'Corrige tipo de venda em desuso
				If tipo_venda = 3 Then
					tipo_venda = 2
				End If
				
				
				If tipo_venda = 1 and forma_pagamento = 1 Then
				'Venda normal e à vista
					temp = "normal"
					temp2 = "avista"
				Elseif  tipo_venda = 1 and forma_pagamento = 2 Then
				'Venda normal e à prazo
					temp = "normal"
					temp2 = "aprazo"			
				Elseif  tipo_venda = 2 and forma_pagamento = 1 Then
				'Venda repasse e à vista
					temp = "repasse"
					temp2 = "avista"				
				Elseif  tipo_venda = 2 and forma_pagamento = 2 Then
				'Venda repasse e à prazo
					temp = "repasse"
					temp2 = "aprazo"
				End If
				

				if temp = "repasse" and cod_tabela_comissao_vendedores = "" then
					cod_tabela_comissao_vendedores = fechamentoRetornaTabelaRepassePorAno(cod_formando,"", true)
				''	response.write ("FORMANDO: "&cod_formando)
				''	response.write ("<br>cod_tabela_comissao_vendedores: "&cod_tabela_comissao_vendedores)
				''	response.end 
				end if

				sql = "SELECT tipo_pagamento_"&temp&"_"&temp2&", valor_pagamento_"&temp&"_"&temp2&" FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES_PRODUTOS WHERE cod_produto = "&cod_produto&" and cod_tabela_comissao_vendedores = "&cod_tabela_comissao_vendedores				
				
				'Response.Write(sql)
				Set rs2=objConexao.Execute(sql)
					
					If not rs2.Eof Then
						
						If Cstr(""&rs2(0)) = "F" Then
							retorna = formatNumber(0&rs2(1))
							retorna = (retorna)/1
							retorna = (retorna*quantidade)/1
							if aux_desc > 0 Then
								retorna = retorna * (1-(aux_desc/100))
							end if 
						Else
							'Verifica o valor 
							retorna = formatNumber(0&rs2(1))
							temp = (retorna)/100
							'Response.Write(">"&temp&"<")
							retorna = temp*valor_total
							if aux_desc > 0 Then
								retorna = retorna * (1-(aux_desc/100))
							end if 
						End If
						
					End If
				
				rs2.Close
				Set rs2=Nothing
				
				
			End If
			
		Else
			Response.Write("ERRO: PEDIDO NÃO ENCONTRADO!!!!!!("&cod_pedido&")")
		
		End If
		
	rs.Close
	Set rs=Nothing

	If not conexao_aberta Then
		FecharConexao()
	End If
	
	fechamentoRetornaValorComissaoPedidoProduto = retorna
End Function


'tipo = 1 retorna os valores
'tipo = 2 retorna o detalhamento de como foi realizado o cálculo
Function fechamentoRetornaValorComissaoPedido(cod_pedido,tipo,conexao_aberta)

	If not conexao_aberta Then
		AbrirConexao()
	End If

	Dim rs, rs2, sql, retorna, temp
	Dim detalhamento_calculo
	Dim quantidade, valor_total, aliquota_desconto_foto, aliquota_desconto
	Dim cod_foto_padrao, pedido_extra 
	retorna = 0
	pedido_extra = "2"
	
	
		Set rs=objConexao.Execute("SELECT pedido_extra FROM TB_PEDIDOS WHERE cod_pedido = "&cod_pedido)
		
			If not rs.Eof Then
				pedido_extra = Cstr(""&rs("pedido_extra"))
			End If
	
		rs.Close
		Set rs=Nothing
		
	
		'1/3 - Soma os valores de todos os produtos vendidos
		'//////////////////////////////////////////////////////////////////////////		
			'Set rs2=objConexao.Execute("SELECT COD_PRODUTO, NOME_PRODUTO, SUM(QUANTIDADE) AS QUANTIDADE, SUM(QUANTIDADE*VALOR_UN) AS VALOR_TOTAL, aliquota_desconto FROM VW_TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO = "&cod_pedido&" GROUP BY COD_PRODUTO")
			Set rs2=objConexao.Execute("SELECT COD_PRODUTO, NOME_PRODUTO,  QUANTIDADE, QUANTIDADE*VALOR_UN AS VALOR_TOTAL, aliquota_desconto, cod_formando FROM VW_TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO = "&cod_pedido&" order by cod_formando")
						
				If not rs2.Eof Then
					
					detalhamento_calculo = "COMISSÃO POR PRODUTOS VENDIDOS À PARTE<br />"
					
					While not rs2.Eof 
						
						quantidade = formatNumber(0&rs2("quantidade"),0)
						valor_total = formatNumber(0&rs2("valor_total"))
						aliquota_desconto = formatNumber(0&rs2("aliquota_desconto"))
						quantidade = quantidade/1
						valor_total = valor_total/1
						temp = fechamentoRetornaValorComissaoPedidoProduto(cod_pedido,rs2("cod_produto"),quantidade,valor_total,aliquota_desconto,true, rs2("cod_formando"))
						retorna = (retorna+temp)/1
						
						detalhamento_calculo = detalhamento_calculo&" - "&rs2("NOME_PRODUTO")&"["&rs2("cod_formando")&"] -  "&formatCurrency(temp)

						if aliquota_desconto > 0 Then
							detalhamento_calculo = detalhamento_calculo & " (desc. "&aliquota_desconto&"%)"
						end if 
						detalhamento_calculo = detalhamento_calculo & "<br />"
						'Response.Write(rs2("cod_produto")&" | Comissão: "&retorna&" | Qtd.: "&quantidade&" | Valor Total: "&valor_total&" <br />")

					rs2.MoveNext()
					Wend
					
				End If
						
			rs2.Close
			Set rs2=Nothing
			
		If PEDIDO_EXTRA = "2" Then
			
			'2/3 - VALOR TOTAL DO ÁLBUM DE CADA FORMANDO DE ACORDO COM A FOTO PADRÃO DE CADA PRODUÇÃO
			'///////////////////////////////////////////////////////////////////////////
					
					'Seleciona todos os formandos
					Set rs=objConexao.Execute("SELECT TB_PEDIDOS_FORMANDOS.COD_FORMANDO, TB_FICHAS_IDENTIFICACAO.VALOR_UN, TB_FICHAS_IDENTIFICACAO.aliquota_desconto_foto, TB_FICHAS_IDENTIFICACAO.FATURADO, TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO,  IFNULL(TB_PRODUCAO.COD_FOTO_PADRAO,0)  AS COD_FOTO_PADRAO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO LEFT JOIN TB_PRODUCAO ON TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO = TB_PRODUCAO.COD_PRODUCAO WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&cod_pedido)
					
						If not rs.Eof Then
						
							detalhamento_calculo = detalhamento_calculo&"<br />COMISSÃO DE ÁLBUNS VENDIDOS<br />"
	
							While not rs.Eof
								
								cod_foto_padrao = Cstr(""&rs("cod_foto_padrao"))
								
								If cod_foto_padrao = "0" Then
								'Por definição escolhe a foto 
									cod_foto_padrao = "20"
								End If
								
								quantidade = formatNumber(0&rs("FATURADO"),0)
								valor_total = formatNumber(0&rs("VALOR_UN"))
								aliquota_desconto_foto = formatNumber(0&rs("aliquota_desconto_foto"))

								
								If quantidade > 0 Then
										
										valor_total = (quantidade*valor_total)
										'Response.Write("F padrão: "&cod_foto_padrao&" | Qtd.: "&quantidade&" | Valor Total: "&valor_total&" <br />")
										
										temp = fechamentoRetornaValorComissaoPedidoProduto(cod_pedido,cod_foto_padrao,quantidade,valor_total,aliquota_desconto_foto,true, rs("COD_FORMANDO"))
										retorna = (retorna+temp)/1
								
										
										detalhamento_calculo = detalhamento_calculo&" - "&rs("NOME_FORMANDO")&" - "&formatCurrency(temp)

										if aliquota_desconto_foto > 0 Then
											detalhamento_calculo = detalhamento_calculo& " (desc. "&aliquota_desconto_foto&"%)"
										end if 

										detalhamento_calculo = detalhamento_calculo&"<br />"
								
								End If								
	
							rs.MoveNext()
							Wend
							
						End If
				
					rs.Close
					Set rs=Nothing
					
			'3/3 - VALOR TOTAL DE ITENS VENDIDOS COMO PRODUTO AGREGADO
			'///////////////////////////////////////////////////////////////////////////
			
					Set rs2=objConexao.Execute("SELECT TB_PEDIDOS_FORMANDOS.COD_FORMANDO, TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO, SUM(TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.QUANTIDADE) AS QUANTIDADE, SUM(TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.QUANTIDADE * TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.VALOR_UN) AS VALOR_TOTAL FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&cod_pedido&" AND TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.VENDIDO = 1 and TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.FOI_BRINDE = 2 GROUP BY TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO")
								
						If not rs2.Eof Then
							
							detalhamento_calculo = detalhamento_calculo&"<br />COMISSÃO DE PRODUTOS AGREGADOS<br />"
							
							While not rs2.Eof 
								
								quantidade = formatNumber(0&rs2("quantidade"),0)
								valor_total = formatNumber(0&rs2("valor_total"))
								quantidade = quantidade/1
								valor_total = valor_total/1
								temp = fechamentoRetornaValorComissaoPedidoProduto(cod_pedido,rs2("cod_produto"),quantidade,valor_total,true, rs2("COD_FORMANDO"))
								
								
								retorna = (retorna+temp)/1
								detalhamento_calculo = detalhamento_calculo&" - Cod. Formando: "&rs2("COD_FORMANDO")&" - "&formatCurrency(temp)&"<br />"
								
							rs2.MoveNext()
							Wend
							
						End If
					rs2.Close
					Set rs2=Nothing
					
		End If
		
		

	If not conexao_aberta Then
		FecharConexao()
	End If
	
	If tipo = 2 Then
		retorna = detalhamento_calculo
	End If
	
	fechamentoRetornaValorComissaoPedido = retorna
	
End Function


Function fechamentoRetornaValorComissaoPedidoFormaPagamento(cod_pedido, tipo, debug, conexao_aberta)

	dim rs, qtd_parcelas, valor_comissao, retorno
  dim valor_pedido, valor_total_boletos, qtd_parcelas_boletos, porcentagem_boleto

  if not conexao_aberta then
  	AbrirConexao()
  end if

  'tipos retorno'
  ' boleto.representatividade  : representatividade de boleto perante o pedido
  ' boleto.comissao_total      : retorna valor total a ser pago para boletos
  ' boleto.comissao_parcelas   : valor individual da parcela (ignorando a 1º parcela)
  ' boleto.comissao_parcela1   : valor da 1 parcela corrigida
  ' boleto.descritivo          : informação formatada da forma de pagamento
  ' outros.representatividade  : representatividade de outros pagamentos perante o pedido
  ' outros.comissao_total      : informação formatada da forma de pagamento
  '/total do pedido
    Set rs = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_AGREGADOS WHERE COD_PEDIDO="&COD_PEDIDO)
  if not rs.Eof Then
    XTOTAL_AGREGADOS = cdbl(0&rs("XTOTAL_AGREGADOS"))
  End If
  rs.Close()
  Set rs = Nothing

  Set rs = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_ALBUNS WHERE COD_PEDIDO="&COD_PEDIDO)
  if not rs.Eof Then
    XTOTAL_ALBUNS = cdbl(0&rs("XTOTAL_ALBUNS"))
  End If
  rs.Close()
  Set rs = Nothing

  Set rs = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_PRODUTOS WHERE COD_PEDIDO="&COD_PEDIDO)
  if not rs.Eof Then
    XTOTAL_PRODUTOS_A_PARTE  = cdbl(0&rs("XTOTAL_PRODUTOS_A_PARTE"))
  End If
  rs.Close()
  Set rs = Nothing

  'BUSCAMOS DESCONTOS'
  Set rsProdutosTotal = objConexao.Execute("select SUM(valor_total_com_desconto) as valor_total_com_desconto from(SELECT SUM(faturado * (valor_un-(valor_un * (1-(ifnull(aliquota_desconto_foto,0)/100))))) as valor_total_com_desconto FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO in (SELECT cod_formando FROM TB_PEDIDOS_FORMANDOS WHERE cod_pedido =  "&COD_PEDIDO&") union all SELECT SUM(quantidade * (valor_un-(valor_un * (1-(ifnull(aliquota,0)/100))))) as valor_total_com_desconto  FROM TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO = "&cod_pedido&") as total_desconto")
  if not rsProdutosTotal.Eof Then
    XTOTAL_DESCONTO = formatNumber(0&rsProdutosTotal(0))/1
  End If
  rsProdutosTotal.Close()
  Set rsProdutosTotal = Nothing


  valor_pedido = XTOTAL_PRODUTOS_A_PARTE + XTOTAL_ALBUNS + XTOTAL_AGREGADOS - XTOTAL_DESCONTO

  '/total  em boletos
  set rs = objConexao.Execute("SELECT sum(valor) as valor_total, count(parcela) as parcelas_boletos FROM TB_PEDIDOS_BOLETOS WHERE COD_PEDIDO = "&sqlNumero(cod_pedido))
  if not rs.Eof then
    valor_total_boletos = cdbl(0&rs("valor_total"))
    qtd_parcelas_boletos = int(0&rs("parcelas_boletos"))
  end if 
  rs.Close
  set rs = Nothing

  valor_comissao = cdbl(fechamentoRetornaValorComissaoPedido(cod_pedido,1,true))

  'CALCULAMOS A REPRESENTATIVIDADE DO VALOR LANÇADO EM BOLETOS SOBRE O VALOR TOTAL DO PEDIDO 
	if valor_pedido > 0 then
	  porcentagem_boleto = (cdbl(""&valor_total_boletos)/valor_pedido)* 100
	  porcentagem_boleto = round(porcentagem_boleto,2)
	  valor_comissao_boleto =  round(valor_comissao * (porcentagem_boleto/100),2)
	else
		porcentagem_boleto = 0
		valor_comissao_boleto = 0
	end if
  'INCLUIMOS A COMISSAO NOS BOLETOS DE ACORDO COM A REPRESENTATIVIDADE DO VALOR LANÇADO EM BOLETOS SOBRE O VALOR TOTAL DO PEDIDO
  dim valor_primeira_comissao, valor_parcelas_comissao

  if valor_comissao_boleto > 0 then
	  valor_parcelas_comissao = cdbl(valor_comissao_boleto) / qtd_parcelas_boletos
	  valor_parcelas_comissao = cdbl(round(valor_parcelas_comissao,2))
	  valor_primeira_comissao = cdbl(valor_parcelas_comissao+(valor_comissao_boleto - (qtd_parcelas_boletos*valor_parcelas_comissao)))
  end if
  'CALCULAMOS A REPRESENTATIVIDADE/VALOR DAS OUTRAS FORMAS DE PAGAMENTO 
  Dim porcentagem_outros, valor_comissao_outros
  Dim PAGAMENTO_EM_DINHEIRO, XTOTAL_CHEQUE, XTOTAL_CARTAO


	'Set rs = objConexao.Execute("SELECT XTOTAL_BOLETO,PAGAMENTO_EM_DINHEIRO, XTOTAL_CHEQUE, XTOTAL_CARTAO FROM VW_PEDIDOS_TOTAL_RECEBIDO WHERE COD_PEDIDO = "&COD_PEDIDO&" LIMIT 1")
	'If not rs.Eof Then
	''	PAGAMENTO_EM_DINHEIRO = cdbl(0&rs("PAGAMENTO_EM_DINHEIRO"))
	''	XTOTAL_CHEQUE = cdbl(0&rs("XTOTAL_CHEQUE"))
	''	XTOTAL_CARTAO = cdbl(0&rs("XTOTAL_CARTAO"))
	
	''	valor_comissao_outros = (PAGAMENTO_EM_DINHEIRO + XTOTAL_CHEQUE +  XTOTAL_CARTAO)
	'End IF
	'rs.Close()
	'Set rs = Nothing

	porcentagem_outros = 100 - porcentagem_boleto

	valor_comissao_outros =  round(valor_comissao * (porcentagem_outros/100),2)

  if tipo = "boleto.representatividade" then
    retorno = porcentagem_boleto
  elseif tipo = "boleto.comissao_total" then
    retorno = valor_comissao_boleto
  elseif tipo = "boleto.comissao_parcelas" then
    retorno = valor_parcelas_comissao
  elseif tipo = "boleto.comissao_parcela1" then
    retorno = valor_primeira_comissao
  elseif tipo = "boleto.descritivo" then
    
  	if valor_comissao_boleto > 0 then
  		retorno = "R$ "&arrumaNumero(round(valor_comissao_boleto,2),"0,00",2)&" em "&qtd_parcelas_boletos&" x R$ "&arrumaNumero(round(valor_parcelas_comissao,2),"0,00",2)
  	else
  		retorno = "-"
  	end if

  elseif tipo = "outros.representatividade" then
  	retorno = porcentagem_outros
  elseif tipo = "outros.comissao_total" then
  	retorno = valor_comissao_outros
  end if

  '/////DEBUG - DUMP 
  if debug then
    response.write ("<br>valor_pedido: "&valor_pedido)
    response.write ("<br>valor_total_boletos: "&valor_total_boletos)
    response.write ("<br>qtd_parcelas_boletos: "&qtd_parcelas_boletos)
    response.write ("<br>porcentagem_boleto: "&porcentagem_boleto)
    response.write ("<br>valor_comissao: "&valor_comissao)
    response.write ("<br>valor_comissao_boleto: "&valor_comissao_boleto)
    response.write ("<br>valor_primeira_comissao: "&valor_primeira_comissao)
    response.write ("<br>valor_parcelas_comissao: "&valor_parcelas_comissao)
    response.write ("<br>porcentagem_outros: "&porcentagem_outros)
    response.write ("<br>valor_comissao_outros: "&valor_comissao_outros)
    response.write ("<br>tipo_retorno: "&tipo)
    response.write ("<br>retorno: "&retorno)
  end if
  '/////DEBUG - DUMP

   if not conexao_aberta then
  	fecharConexao()
  end if

fechamentoRetornaValorComissaoPedidoFormaPagamento = retorno
end function

'Função que cálcula o pagamento do vendedor após ter sido processado e retorna os valores por tipo
'tipo 1 = Total Comissão de Produtos
'tipo 2 = Total Ajuda de Custo
'tipo 3 = Total Comissão de Metas
'tipo 4 = Total Comissão de Coordenador
'tipo 5 = Total Pendencias (pode ser + ou -)
'tipo 6 = Total Pagamento em Boletos'
'tipo 100 = Total Arrecadado de todos os pedidos e formas de pagamento (utilizada para efetuar alguns cálculos)

Function fechamentoRetornaValorRecebidoVendedorTipo(tipo,cod_vendedor,cod_fechamento_periodo,conexao_aberta)

	Dim rs, rs2, sql, temp, temp2, temp3, retorna
	Dim cod_grupo_vendedores, qtd_vendedores_grupo_coordenadores, valor_arrecadado
	retorna = 0
	
	If not conexao_aberta Then
		AbrirConexao()
	End If
	
		'//////////////////
		'Informações do verdedor que poderão ser utilizadas
		Set rs=objConexao.Execute("SELECT cod_grupo_vendedores FROM TB_VENDEDORES WHERE cod_vendedor = "&cod_vendedor)			
			If not rs.Eof Then
				cod_grupo_vendedores = rs("cod_grupo_vendedores")
			End If
		rs.Close
		Set rs=Nothing
		'//////////////////
	
		If tipo = 1 Then

			Set rs=objConexao.Execute("SELECT SUM(VALOR_TOTAL_COMISSAO) as TOTAL FROM TB_FECHAMENTO_PAGAMENTO WHERE cod_vendedor = "&cod_vendedor&" and cod_fechamento_periodo = "&cod_fechamento_periodo&" group by cod_vendedor")			
				If not rs.Eof Then
					retorna = formatNumber(0&rs(0))
				End If
			rs.Close
			Set rs=Nothing

			Set rs=objConexao.Execute("SELECT SUM(valor_comissao_vendedor) as TOTAL FROM TB_FECHAMENTO_PAGAMENTO_BOLETOS WHERE cod_vendedor = "&cod_vendedor&" and cod_fechamento_periodo = "&cod_fechamento_periodo&" group by cod_vendedor")			
				If not rs.Eof Then
					retorna = cdbl(retorna) + cdbl(0&rs(0))
				End If
			rs.Close
			Set rs=Nothing

			
		Elseif tipo = 2 Then

			Set rs=objConexao.Execute("SELECT SUM(VALOR_AJUDA_CUSTO) AS XTOTAL_AJUDA_CUSTO FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_VENDEDOR = "&cod_vendedor&" AND cod_fechamento_periodo ="&cod_fechamento_periodo)
					If not rs.Eof Then
							retorna = formatNumber(0&rs(0))
					End If
			rs.Close
			Set rs=Nothing
	
		Elseif tipo = 3 Then
			
			If cod_grupo_vendedores > 0 Then
			
				'Obs.: Somente para vendedor do grupo repasse!
				Set rs = objConexao.Execute("SELECT SUM(QTD_ALBUNS) AS QTD_ALBUNS FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_VENDEDOR = "&cod_vendedor&" AND COD_GRUPO_VENDEDORES > 0 AND cod_fechamento_periodo ="&cod_fechamento_periodo)
										If not rs.Eof Then
												
												temp = formatNumber(0&rs("QTD_ALBUNS"))
												
												If int(temp) > 29 Then
													'1% do valor arrecadado total
													
													'Resgata o valor total arrecadado
													valor_arrecadado = formatNumber(0&fechamentoRetornaValorRecebidoVendedorTipo(100,cod_vendedor,cod_fechamento_periodo,true))												
													'1%
													retorna = valor_arrecadado/100
												End If
																				
										End If
										
				rs.Close
				Set rs = Nothing
								
							'2) META GRUPO //////////////////////////////////////////////////////////
								  
								  'Zera variáveis
								  temp = 0
								  temp2= 0
								  temp3= 0
								  qtd_vendedores_grupo_coordenadores = 0
								  
								  'Qtd de vendedores do grupo no pagamento
								  Set rs=objConexao.Execute("SELECT COUNT(*) FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_GRUPO_VENDEDORES = "&cod_grupo_vendedores&" GROUP BY COD_VENDEDOR",qtd_vendedores_grupo_coordenadores)
								  rs.Close
								  Set rs=Nothing
									
									
								  'Contabiliza qtd de álbum do grupo
								  Set rs = objConexao.Execute("SELECT SUM(QTD_ALBUNS) AS QTD_ALBUNS_VENDIDOS FROM TB_FECHAMENTO_PAGAMENTO WHERE cod_fechamento_periodo = "&cod_fechamento_periodo&" AND COD_GRUPO_VENDEDORES = "&cod_grupo_vendedores&" GROUP BY COD_GRUPO_VENDEDORES")
			  
									  If not rs.Eof Then
										temp = formatNumber(0&rs("QTD_ALBUNS_VENDIDOS"),0)
									  End If
			  
								  rs.Close
								  Set rs = Nothing
								
								  'Cálcula a média (qtd de álbuns/qtd de vendedores no grupo)
								  
								  If qtd_vendedores_grupo_coordenadores > 0 and temp > 0 Then
										temp2 = (temp/qtd_vendedores_grupo_coordenadores)
										
										If int(temp2) => 30 Then
										'1% do valor arrecadado total
											temp = (valor_arrecadado)/100
											
											'Soma no valor da meta se houver anterior
											retorno = (retorno+temp)/1
											
										End If
										
								  End If
								  
			  
			 
			End If
			 
		Elseif tipo = 4 Then
			
			'Somente para Mário e Renato
			If int(cod_vendedor) = 14 or int(cod_vendedor) = 7 or int(cod_vendedor) = 48 or int(cod_vendedor) = 157 Then
					If int(COD_VENDEDOR) = 7 Then
						cod_grupo_vendedores = 1 'Renato
					Elseif  int(COD_VENDEDOR) = 14 Then
						cod_grupo_vendedores = 2 'Mário
					Elseif  int(COD_VENDEDOR) = 48 Then
						cod_grupo_vendedores = 3 'Mantuam
          Elseif  int(COD_VENDEDOR) = 157 Then
						cod_grupo_vendedores = 4 'Vicente
					End If

					'Desativado pois é obtido de acordo com o formando da venda	
					'1) Percentual pago aos coordenadores
					'Set rs=objConexao.Execute("SELECT * FROM TB_FECHAMENTO_COMISSIONAMENTO_COORDENADORES WHERE DATA_INICIO <= "&sqlData(data_pagamento)&" ORDER BY DATA_INICIO DESC LIMIT 1")
					''  If not rs.Eof Then
					''	  temp = formatNumber(0&rs("PERCENTUAL"),0)
					''  End If
					'rs.Close
					'Set rs=Nothing

					'2) Total arrecadado do grupo


					Set rs=objConexao.Execute("SELECT cod_pedido FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND COD_GRUPO_VENDEDORES = "&COD_GRUPO_VENDEDORES&" ")
					If not rs.Eof Then
						While not rs.Eof 	

							'COMISSAO DE FOTOS
							'
							'Seleciona todos os formandos
							Set rs2=objConexao.Execute("SELECT TB_PEDIDOS_FORMANDOS.COD_FORMANDO, TB_FICHAS_IDENTIFICACAO.VALOR_UN, TB_FICHAS_IDENTIFICACAO.aliquota_desconto_foto, TB_FICHAS_IDENTIFICACAO.FATURADO, TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO,  IFNULL(TB_PRODUCAO.COD_FOTO_PADRAO,0)  AS COD_FOTO_PADRAO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO LEFT JOIN TB_PRODUCAO ON TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO = TB_PRODUCAO.COD_PRODUCAO WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&rs("cod_pedido"))
							
								If not rs2.Eof Then			
									While not rs2.Eof
										
										cod_foto_padrao = Cstr(""&rs2("cod_foto_padrao"))
										
										If cod_foto_padrao = "0" Then
										'Por definição escolhe a foto 
											cod_foto_padrao = "20"
										End If
										
										quantidade = formatNumber(0&rs2("FATURADO"),0)
										valor_total = formatNumber(0&rs2("VALOR_UN"))
										aliquota_desconto_foto = formatNumber(0&rs2("aliquota_desconto_foto"))

										If quantidade > 0 Then
											valor_total = (quantidade*valor_total)											
											temp = fechamentoRetornaValorComissaoCoordenadorPedidoProduto(rs("cod_pedido"),cod_foto_padrao,quantidade,valor_total,aliquota_desconto_foto,true, rs2("COD_FORMANDO"))
											retorna = (retorna+temp)/1
										End If								
			
									rs2.MoveNext()
									Wend
									
								End If
						
							rs2.Close
							Set rs2=Nothing

							'COMISSAO DE PRODUTOS
							Set rs2=objConexao.Execute("SELECT COD_PRODUTO, NOME_PRODUTO,  QUANTIDADE, QUANTIDADE*VALOR_UN AS VALOR_TOTAL, aliquota_desconto, cod_formando FROM VW_TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO = "&rs("cod_pedido")&" order by cod_formando")
								If not rs2.Eof Then
									While not rs2.Eof 
										percentual = cdbl(fechamentoRetornaPercentualRepasseCoodenadores(rs2("cod_formando"), "", true))

										valor_total = cdbl(0&rs2("VALOR_TOTAL"))
										aliquota_desconto = cdbl(0&rs2("aliquota_desconto"))
										valor_total = valor_total/1

										temp = (valor_total-(valor_total*(aliquota_desconto/100))) * (percentual/100)

										retorna = (retorna+temp)/1
									rs2.MoveNext()
									Wend
								End If
							rs2.Close
							Set rs2=Nothing
						rs.MoveNext()
						Wend
					End If
					rs.Close
					Set rs=Nothing
'response.end

					'2) Total arrecadado do grupo
					'Set rs=objConexao.Execute("SELECT SUM(TOTAL_RECEBIMENTO) AS TOTAL_RECEBIDO FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND COD_GRUPO_VENDEDORES = "&COD_GRUPO_VENDEDORES&" GROUP BY COD_FECHAMENTO_PERIODO")
					'If not rs.Eof Then
					''	retorna = formatNumber(0&rs("TOTAL_RECEBIDO"))/100 * temp	
					'End If
					'rs.Close
					'Set rs=Nothing

					'Set rs=objConexao.Execute("SELECT SUM(valor) AS TOTAL_RECEBIDO_BOLETOS FROM TB_FECHAMENTO_PAGAMENTO_BOLETOS WHERE COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND COD_GRUPO_VENDEDORES = "&COD_GRUPO_VENDEDORES&" GROUP BY COD_FECHAMENTO_PERIODO")
					'If not rs.Eof Then
					''	retorna = retorna + (formatNumber(0&rs("TOTAL_RECEBIDO_BOLETOS"))/100 * temp	)
					'End If
					'rs.Close
					'Set rs=Nothing
			
			End If
			
		Elseif tipo = 5 Then
			
				Set rs = objConexao.Execute("SELECT * FROM TB_FECHAMENTO_PENDENCIAS_PAGAMENTO WHERE COD_VENDEDOR = "&COD_VENDEDOR&" AND cod_fechamento_periodo ="&cod_fechamento_periodo)
  
					If not rs.Eof Then
				
						While not rs.Eof
							
							temp = 0
							temp = formatNumber(0&rs("VALOR"))
		
							If Cstr(rs("TIPO"))  = "A"  Then
							'Acréscimo
								retorna = (retorna+temp)/1
							Else
							'Descréscimo
								retorna = (retorna-temp)/1
							End If
	
						rs.MoveNext()
						Wend
	
				End If
  
			rs.Close
			Set rs = Nothing

		elseif tipo = 6 then

			Set rs = objConexao.Execute("SELECT SUM(valor_comissao_vendedor) FROM TB_FECHAMENTO_PAGAMENTO_BOLETOS WHERE COD_VENDEDOR = "&COD_VENDEDOR&" AND cod_fechamento_periodo ="&cod_fechamento_periodo)
			if not rs.eof then
				retorna = cdbl(0&rs(0))
			end if 
			rs.Close
			set rs = nothing 
			
		Elseif tipo = 100 Then
			
			Set rs=objConexao.Execute("SELECT SUM(TOTAL_RECEBIMENTO) AS TOTAL_RECEBIDO FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND COD_VENDEDOR = "&cod_vendedor&" GROUP BY COD_FECHAMENTO_PERIODO")
		
				If not rs.Eof Then
					retorna = formatNumber(0&rs(0))
				End If
		
			rs.Close
			Set rs=Nothing				
			
		End If
	
	'Formata com 2 casas decimais
	retorna = formatNumber(retorna,2)
	
	If not conexao_aberta Then
		FecharConexao()
	End If
	
	fechamentoRetornaValorRecebidoVendedorTipo = retorna

End Function

'Soma todos os tipos de todais válidos
Function fechamentoRetornaValorRecebidoVendedor(cod_vendedor,cod_fechamento_periodo,conexao_aberta)

	
	Dim rs, sql, retorna, i
	
	If not conexao_aberta Then
		AbrirConexao()
	End If	
	
	
	'Tipos de totais válidos
	For i=1 to 5
		
		temp = fechamentoRetornaValorRecebidoVendedorTipo(i,cod_vendedor,cod_fechamento_periodo,true)
		'Response.Write("Total do tipo"&i&" : "&temp&" <br />")
		
		retorna = (retorna+temp)/1
		
	Next

	If not conexao_aberta Then
		FecharConexao()
	End If
	
	
	fechamentoRetornaValorRecebidoVendedor = retorna
	
End Function





function fechamentoRetornaTabelaRepassePorAno(cod_formando, ano_conclusao, conexao_aberta)

	dim retorno, rs, ano_conclusao_

	if not conexao_aberta then
		AbrirConexao()
	end if

	if ano_conclusao = "" and cod_formando <> "" then
		set rs = objConexao.Execute("SELECT ano_conclusao FROM TB_FICHAS_IDENTIFICACAO WHERE cod_formando = "&sqlNumero(cod_formando))
		if not rs.Eof then
			ano_conclusao_ = rs(0)
		end if
		rs.Close
		set rs = Nothing
	else
		ano_conclusao_ = ano_conclusao
	end if

	'Response.Write(ano_conclusao_) 

	Set rs = objConexao.Execute("SELECT cod_tabela_comissao_vendedores FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES where tipo = 'R' AND ano_inicio <= "&sqlTexto(ano_conclusao_)&"and ano_fim >= "&sqlTexto(ano_conclusao_))
	if not rs.Eof then
		retorno = rs(0)
	end if 
	rs.Close
	set rs = nothing

	if not conexao_aberta then
		fecharConexao()
	end if
fechamentoRetornaTabelaRepassePorAno = retorno
end function

function fechamentoRetornaPercentualRepasseCoodenadores(cod_formando, ano_conclusao, conexao_aberta)

	dim retorno, rs, ano_conclusao_

	if not conexao_aberta then
		AbrirConexao()
	end if

	if ano_conclusao = "" and cod_formando <> "" then
		set rs = objConexao.Execute("SELECT ano_conclusao FROM TB_FICHAS_IDENTIFICACAO WHERE cod_formando = "&sqlNumero(cod_formando))
		if not rs.Eof then
			ano_conclusao_ = rs(0)
		end if
		rs.Close
		set rs = Nothing
	else
		ano_conclusao_ = ano_conclusao
	end if

	Set rs = objConexao.Execute("SELECT Percentual FROM TB_FECHAMENTO_COMISSIONAMENTO_COORDENADORES where ano_inicio <= "&sqlTexto(ano_conclusao_)&"and ano_fim >= "&sqlTexto(ano_conclusao_))
	if not rs.Eof then
		retorno = rs(0)
	end if 
	rs.Close
	set rs = nothing

	if not conexao_aberta then
		fecharConexao()
	end if
fechamentoRetornaPercentualRepasseCoodenadores = retorno
end function

'123aqui'
'1 - retorna valor'
'2 - retorna detalhamento'
Function fechamentoRetornaValorComissaoCoordenadorPedido(cod_pedido,tipo,conexao_aberta)

	If not conexao_aberta Then
		AbrirConexao()
	End If

	Dim rs, rs2, sql, retorna, temp
	Dim detalhamento_calculo
	Dim quantidade, valor_total, aliquota_desconto_foto, aliquota_desconto
	Dim cod_foto_padrao, pedido_extra 
	retorna = 0
	pedido_extra = "2"
	
	
		Set rs=objConexao.Execute("SELECT pedido_extra FROM TB_PEDIDOS WHERE cod_pedido = "&cod_pedido)
		
			If not rs.Eof Then
				pedido_extra = Cstr(""&rs("pedido_extra"))
			End If
	
		rs.Close
		Set rs=Nothing
		
	
		'1/3 - Soma os valores de todos os produtos vendidos
		'//////////////////////////////////////////////////////////////////////////		
			'Set rs2=objConexao.Execute("SELECT COD_PRODUTO, NOME_PRODUTO, SUM(QUANTIDADE) AS QUANTIDADE, SUM(QUANTIDADE*VALOR_UN) AS VALOR_TOTAL, aliquota_desconto FROM VW_TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO = "&cod_pedido&" GROUP BY COD_PRODUTO")
			Set rs2=objConexao.Execute("SELECT COD_PRODUTO, NOME_PRODUTO,  QUANTIDADE, QUANTIDADE*VALOR_UN AS VALOR_TOTAL, aliquota_desconto, cod_formando FROM VW_TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO = "&cod_pedido&" order by cod_formando")
						
				If not rs2.Eof Then
					
					While not rs2.Eof 
						
						quantidade = formatNumber(0&rs2("quantidade"),0)
						valor_total = formatNumber(0&rs2("valor_total"))
						aliquota_desconto = formatNumber(0&rs2("aliquota_desconto"))
						quantidade = quantidade/1
						valor_total = valor_total/1
						temp = fechamentoRetornaValorComissaoCoordenadorPedidoProduto(cod_pedido,rs2("cod_produto"),quantidade,valor_total,aliquota_desconto,true, rs2("cod_formando"))
						retorna = (retorna+temp)/1

					rs2.MoveNext()
					Wend
					
				End If
						
			rs2.Close
			Set rs2=Nothing
			
		If PEDIDO_EXTRA = "2" Then
			
			'2/3 - VALOR TOTAL DO ÁLBUM DE CADA FORMANDO DE ACORDO COM A FOTO PADRÃO DE CADA PRODUÇÃO
			'///////////////////////////////////////////////////////////////////////////
					
					'Seleciona todos os formandos
					Set rs=objConexao.Execute("SELECT TB_PEDIDOS_FORMANDOS.COD_FORMANDO, TB_FICHAS_IDENTIFICACAO.VALOR_UN, TB_FICHAS_IDENTIFICACAO.aliquota_desconto_foto, TB_FICHAS_IDENTIFICACAO.FATURADO, TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO,  IFNULL(TB_PRODUCAO.COD_FOTO_PADRAO,0)  AS COD_FOTO_PADRAO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO LEFT JOIN TB_PRODUCAO ON TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO = TB_PRODUCAO.COD_PRODUCAO WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&cod_pedido)
					
						If not rs.Eof Then
							While not rs.Eof
								
								cod_foto_padrao = Cstr(""&rs("cod_foto_padrao"))
								
								If cod_foto_padrao = "0" Then
								'Por definição escolhe a foto 
									cod_foto_padrao = "20"
								End If
								'response.write("cod_foto_padrao : "&cod_foto_padrao)
								'response.End
								quantidade = formatNumber(0&rs("FATURADO"),0)
								valor_total = formatNumber(0&rs("VALOR_UN"))
								aliquota_desconto_foto = formatNumber(0&rs("aliquota_desconto_foto"))
								
								If quantidade > 0 Then
										
									valor_total = (quantidade*valor_total)


									temp = fechamentoRetornaValorComissaoCoordenadorPedidoProduto(cod_pedido,cod_foto_padrao,quantidade,valor_total,aliquota_desconto_foto,true, rs("COD_FORMANDO"))

									retorna = (retorna+temp)/1								
								
								End If								
	
							rs.MoveNext()
							Wend
							
						End If
				
					rs.Close
					Set rs=Nothing
					
			'3/3 - VALOR TOTAL DE ITENS VENDIDOS COMO PRODUTO AGREGADO
			'///////////////////////////////////////////////////////////////////////////
			
					Set rs2=objConexao.Execute("SELECT TB_PEDIDOS_FORMANDOS.COD_FORMANDO, TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO, SUM(TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.QUANTIDADE) AS QUANTIDADE, SUM(TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.QUANTIDADE * TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.VALOR_UN) AS VALOR_TOTAL FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&cod_pedido&" AND TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.VENDIDO = 1 and TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.FOI_BRINDE = 2 GROUP BY TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO")
								
						If not rs2.Eof Then
							
							detalhamento_calculo = detalhamento_calculo&"<br />COMISSÃO DE PRODUTOS AGREGADOS<br />"

							While not rs2.Eof 
								quantidade = formatNumber(0&rs2("quantidade"),0)
								valor_total = formatNumber(0&rs2("valor_total"))
								quantidade = quantidade/1
								valor_total = valor_total/1
								temp = fechamentoRetornaValorComissaoCoordenadorPedidoProduto(cod_pedido,rs2("cod_produto"),quantidade,valor_total,true, rs2("COD_FORMANDO"))
								
								retorna = (retorna+temp)/1

							rs2.MoveNext()
							Wend
							
						End If
					rs2.Close
					Set rs2=Nothing
					
		End If
		
		

	If not conexao_aberta Then
		FecharConexao()
	End If
	
	If tipo = 2 Then
		retorna = detalhamento_calculo
	End If
	
	fechamentoRetornaValorComissaoCoordenadorPedido = retorna
	
End Function



Function fechamentoRetornaValorComissaoCoordenadorPedidoProduto(cod_pedido,cod_produto,quantidade,valor_total,aliquota_desconto, conexao_aberta, cod_formando)

	If not conexao_aberta Then
		AbrirConexao()
	End If	
	
	Dim rs, rs2, rs3, sql, temp, temp2, retorna, percentual
	Dim cod_tabela_comissao_vendedores, forma_pagamento, tipo_venda, data_venda
	Dim aux_desc
	aux_desc = formatNumber(aliquota_desconto)
	retorna = 0
	
	sql = "SELECT TB_PEDIDOS.FORMA_PAGAMENTO, TB_PEDIDOS.TIPO_VENDA, TB_PEDIDOS.DATA_VENDA, TB_VENDEDORES.cod_tabela_comissao_vendedores FROM TB_VENDEDORES INNER JOIN TB_PEDIDOS ON TB_VENDEDORES.COD_VENDEDOR = TB_PEDIDOS.COD_VENDEDOR WHERE TB_PEDIDOS.COD_PEDIDO = "&cod_pedido
	Set rs=objConexao.Execute(sql) 
		
		If not rs.Eof Then
			cod_tabela_comissao_vendedores = ""&rs("cod_tabela_comissao_vendedores")
			forma_pagamento = formatNumber(0&rs("forma_pagamento"),0)
			tipo_venda = formatNumber(0&rs("tipo_venda"),0)
			data_venda = rs("data_venda")

			percentual = fechamentoRetornaPercentualRepasseCoodenadores(cod_formando, "", true)

			'response.write valor_total
			'response.end

			retorna = cdbl(valor_total) * (cdbl(percentual)/100)			
		
			
		Else
			Response.Write("ERRO: PEDIDO NÃO ENCONTRADO!!!!!!("&cod_pedido&")")
		
		End If
		
	rs.Close
	Set rs=Nothing

	If not conexao_aberta Then
		FecharConexao()
	End If
	
	fechamentoRetornaValorComissaoCoordenadorPedidoProduto = retorna
End Function

%>
