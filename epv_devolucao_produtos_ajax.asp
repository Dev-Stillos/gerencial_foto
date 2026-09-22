<% dim script_valida_form_master : script_valida_form_master = "NAO" %>
<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<!-- #include file="files/includes/functions_local.asp" -->
<%

response.charset = "ISO 8859-1"

dim retorno
dim t : t = request.form("t")
dim p(10)

for i = 0 to 10
	p(i) = request.form("p"&i)&""
next
'TESTE| LOG
'(OK) | (OK)  trocarRespOS(cod_ordem_servico, cod_usuario_resp, trocar_prox_etapas, conexao_aberta)

if t = "lista_produtos" then
	'Lista produtos em posse do vendedor
	retorno = listarProdutosVendedor (p(0), p(1), false)
end if

response.write retorno

'tipo_retorno = D:Listagem de devolução | R:Listagem Renovação
function listarProdutosVendedor (cod_vendedor, tipo_retorno, conexao_aberta)

	dim retorno,  rs
	dim quantidade_faltante, quantidade, quantidade_saida
	if not conexao_aberta then
		abrirConexao()
	end if

	if tipo_retorno = "D" or tipo_retorno = "" then

		Set rs = objConexao.Execute("SELECT sum(quantidade) as quantidade, sum(quantidade_saida) as quantidade_saida, cod_produto, nome_produto FROM VW_TB_ESTOQUE_MOVIMENTACOES_VENDEDORES WHERE tipo = 'E' and quantidade != quantidade_saida and cod_vendedor = "&sqlNumero(cod_vendedor)&" group by cod_produto")
		if not rs.eof then
			retorno = "<table class='lista_produtos' width='700px' style='border-collapse:collapse;' border='1' align='center'>"
			retorno = retorno & "<tr><th>Produtos</th><th>qtd em posse/ qtd devolução</th><th>Devolver</th></tr>"
			while not rs.eof
				
				quantidade = cstr("0"&rs("quantidade"))
				quantidade_saida = cstr("0"&rs("quantidade_saida"))
				quantidade_faltante = int(quantidade) - int(quantidade_saida)

				retorno = retorno & "<tr><td>"&rs("nome_produto")&"</td>"&_
									"<td align='center'>"&quantidade_faltante&" / <input type='text' class='inputTamanho1' value='"&quantidade_faltante&"' style='width:30px;' name='produto_"&rs("cod_produto")&"' /></td>"&_
								    "<td align='center'><input name='produtos' value='"&rs("cod_produto")&"' type='checkbox' /></td></tr>"
				rs.movenext()
			wend
			retorno = retorno & "</table>"
		end if
		rs.close
		set rs = nothing

	elseif tipo_retorno = "R" then

		Set rs = objConexao.Execute("SELECT cod_movimentacao_vendedor, quantidade, quantidade_saida, cod_produto, data_devolucao, nome_produto FROM VW_TB_ESTOQUE_MOVIMENTACOES_VENDEDORES WHERE tipo = 'E' and quantidade != quantidade_saida and cod_vendedor =  "&sqlNumero(cod_vendedor))'&" AND data_devolucao < NOW()")
		if not rs.eof then
			
			retorno = "<table class='lista_produtos' width='700px' style='border-collapse:collapse;' border='1' align='center'>"
			retorno = retorno & "<tr><th>#</th><th>Produtos</th><th>qtd em posse.</th><th>Data Devolução</th><th>Renovar por</th><th>Renovar</th></tr>"
			
			while not rs.eof
				quantidade = cstr("0"&rs("quantidade"))
				quantidade_saida = cstr("0"&rs("quantidade_saida"))
				quantidade_faltante = int(quantidade) - int(quantidade_saida)
				retorno = retorno & "<tr><td align='center'><strong>"&rs("cod_movimentacao_vendedor")&"</strong></td>"&_
									"<td>"&rs("nome_produto")&"</td>"&_
									"<td align='center'>"&quantidade_faltante&"</td>"&_
									"<td align='center'>"&arrumaData(mid(rs("data_devolucao"),1,10))&"</td>"&_
									"<td align='center'><input class='inputTamanho1' style='width:30px' type='text' name='renovacao_dias_"&rs("cod_movimentacao_vendedor")&"' value='20'> Dias</td>"&_
									"<td align='center'><input name='movimentacoes_renovacao' value='"&rs("cod_movimentacao_vendedor")&"' type='checkbox' /></td></tr>"
				rs.movenext()
			wend
			retorno = retorno & "</table>"
		end if
		rs.close
		set rs = nothing

	end if

	if not conexao_aberta then
		fecharConexao()
	end if

	listarProdutosVendedor = retorno
end function


%>