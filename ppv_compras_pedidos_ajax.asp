<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<!-- #include file="files/includes/functions_local.asp" -->
<%
Dim tipo, i, retorno 
Dim valor(10)

tipo = Request.QueryString("ts")
i=0
While i <= ubound(valor)
	if i = 0 then
		valor(i) = Request.QueryString("valor")
	Else
		valor(i) = Request.QueryString("valor"&i)
	End if
i=i+1
Wend
AbrirConexao()
If tipo = "baixa_entrega" then
	'valor = cod_pedido_compra_produto
	'valor2 = cod_pedido_compra
	
	Set rs = Objconexao.Execute("SELECT * FROM VW_TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra ="&sqlNumero(valor(1)))
	If not rs.Eof Then
		ObjConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS_PRODUTOS SET status_produto = 'E', quantidade_entregue = "&sqlNumero(rs("quantidade_solicitada"))&" WHERE cod_pedido_compra_produto = "&sqlTexto(valor(0))) 
	End if
	rs.close()
	Set rs = nothing
	retorno = "<!--#true#window.location = window.location.href#-->"

Elseif tipo = "aprovar_pedido" then
	'valor = cod_pedido_compra
	if not pedidoCompraMudancaStatus(valor(0),2,"",true) then
		a = 1/0
	end if

Elseif tipo = "finalizar_pedido" then
	'valor = cod_pedido_compra
	dim tmep : tmep=  pedidoCompraMudancaStatus(valor(0),3,"",true)
	if not tmep then
		'response.write tmep
		a = 1/0
	Else
		'Entrada no estoque
		set rs2 = objConexao.Execute("SELECT * FROM  TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra = "&sqlNumero(valor(0)))
		if Not rs2.Eof then

			While not rs2.Eof 
				
				call estoqueMovimentacao("E", "pedido_compra", valor(0), rs2("cod_produto"), rs2("quantidade_entregue"), "Entrada automática, Pedido de Compras:" & valor(0), true)
				
			rs2.movenext
			Wend

		end if
		rs2.movenext
		set rs2 = nothing


	end if

Elseif tipo = "gerar_recibo" then
	'valor = cod_pedido_compra

	'Deletamos da tabela temporaria
	objConexao.Execute("DELETE FROM TB_PEDIDOS_COMPRAS_PRODUTOS_TEMPORARIA WHERE cod_usuario = "&sqlNumero(Session("cod_usuario")))
	
	'Inseriomos novos registros
	objConexao.Execute("INSERT INTO TB_PEDIDOS_COMPRAS_PRODUTOS_TEMPORARIA (cod_pedido_compra_produto, cod_usuario) SELECT cod_pedido_compra_produto, "&session("cod_usuario")&" FROM TB_PEDIDOS_COMPRAS_PRODUTOS WHERE status_produto = 'E' AND cod_pedido_compra = "&sqlNumero(valor(0)))


Elseif tipo = "cancelar_produto" then
	'valor = cod_pedido_compra_produto
	'valor2 = cod_pedido_compra	
	
	Set rs = Objconexao.Execute("SELECT * FROM VW_TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra ="&sqlNumero(valor(1)))
	If not rs.Eof Then
		ObjConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS_PRODUTOS SET status_produto = 'C' WHERE cod_pedido_compra_produto = "&sqlTexto(valor(0))) 
	End if
	rs.close()
	Set rs = nothing
	retorno = "<!--#true#window.location = window.location.href#-->"
	
Elseif tipo = "cee" and validaUsuario(2,,69,true) Then
'Cancelar entrada errada

	Dim cod_pedido_compra_produto, cod_ordem_servico_etapa, cod_ordem_servico, cod_pedido_compra
	cod_pedido_compra_produto = Base64Decode(valor(0))
	
	'valor(0) = cod_pedido_compra_produto
	Set rs=objConexao.Execute("SELECT VW_TB_PEDIDOS_COMPRAS_PRODUTOS.* FROM VW_TB_ORDENS_SERVICO INNER JOIN VW_TB_PEDIDOS_COMPRAS_PRODUTOS on VW_TB_ORDENS_SERVICO.COD_ORDEM_SERVICO = VW_TB_PEDIDOS_COMPRAS_PRODUTOS.COD_ORDEM_SERVICO WHERE VW_TB_PEDIDOS_COMPRAS_PRODUTOS.status_pedido_compra <> 0 AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.cod_pedido_compra_produto = "&cod_pedido_compra_produto)
	
		If not rs.Eof Then
			
			cod_ordem_servico_etapa = Cstr(""&rs("cod_ordem_servico_etapa"))
			cod_ordem_servico = Cstr(""&rs("cod_ordem_servico"))
			cod_pedido_compra = rs("cod_pedido_compra")
			
			'1/ Volta o pedido de compra como pendente.
			sql = "UPDATE TB_PEDIDOS_COMPRAS_PRODUTOS SET status_produto = 'P', quantidade_entregue = NULL WHERE cod_pedido_compra_produto = "&sqlTexto(cod_pedido_compra_produto)
			'Response.Write(sql&"<br />")
			objConexao.Execute(sql) 
		
			'2/ Volta pedido de compra como pendente
			sql = "UPDATE TB_PEDIDOS_COMPRAS SET STATUS_PEDIDO_COMPRA = 2 WHERE cod_pedido_compra = "&sqlTexto(cod_pedido_compra_produto)
			'Response.Write(sql&"<br />")
			objConexao.Execute(sql)
		
			'3/ Retorna etapa da OS
			If cod_ordem_servico_etapa <> "" Then
			
				'Verifica a qtd_etapas
				Set rs3=objConexao.Execute("Select count(*) as qtd_etapas From TB_ORDENS_SERVICO_ETAPAS where cod_ordem_servico = "&cod_ordem_servico&" group by cod_ordem_servico")
												
					If not rs3.Eof Then
						qtd_etapas = formatNumber(0&rs3("qtd_etapas"),0)
					End If
													
				rs3.Close
				Set rs3=Nothing
						
				'Verifica qual o número da etapa atual
				Set rs3=objConexao.Execute("Select numero_etapa From VW_TB_ORDENS_SERVICO where cod_ordem_servico = "&cod_ordem_servico)
					
					If not rs3.Eof Then
						numero_etapa = formatNumber(0&rs3("numero_etapa"),0)
					End If
													
				rs3.Close
				Set rs3=Nothing

				If int(qtd_etapas) = int(numero_etapa) Then
				'Deverá retorna a OS, pois é a última etapa
					sql = "UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_termino = NULL WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&sqlTexto(cod_ordem_servico_etapa)
					objConexao.Execute(sql)
					sql = "UPDATE TB_ORDENS_SERVICO SET status_ordem_servico = 1, data_finalizado = NOW(), cod_usuario_finalizou = NULL WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)
					objConexao.Execute(sql)
				Else
					'Finaliza a etapa atual				
					sql = "UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_termino = NULL WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&sqlTexto(cod_ordem_servico_etapa)	
					objConexao.Execute(sql)
					'Volta para etapa anterior
					Set rs3=objConexao.Execute("SELECT cod_ordem_servico_etapa FROM TB_ORDENS_SERVICO_ETAPAS WHERE cod_ordem_servico = "&cod_ordem_servico&" and numero_etapa IN(SELECT (numero_etapa+1) FROM VW_TB_ORDENS_SERVICO WHERE cod_ordem_servico = "&cod_ordem_servico&")")
										
							If not rs3.Eof Then
								sql = "UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_inicio = NULL WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&rs3("cod_ordem_servico_etapa")
								objConexao.Execute(sql)
								sql = "UPDATE TB_ORDENS_SERVICO SET cod_ordem_servico_etapa_atual = "&cod_ordem_servico_etapa&" WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)
								objConexao.Execute(sql)
							End If
										
					rs3.Close
					Set rs3=Nothing
								
				End If				
				
			End If
		
		End If
	
	rs.Close
	Set rs=Nothing
	
	retorno = "<!--#true#window.location = window.location.href#-->"
	
End if
FecharConexao()
response.Write(retorno)
%>