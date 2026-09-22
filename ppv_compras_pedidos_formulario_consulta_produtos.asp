<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<!-- #include virtual="gerencial_foto/files/includes/scripts.asp" -->
<!-- #include file="files/includes/functions_local.asp" -->
<%
Dim cod_pedido_compra, cod_pedido_compra_produto, imprimir_cabecalho, cod_fornecedor, numero_pedido_compra
Dim campo1, comando, comando_externo, status_produto
Dim cod_produto, obs, quantidade_solicitada
Dim pedido_manual, status_pedido_compra

campo1 = request.form("campo1")
comando = Request.QueryString("comando")
comando_externo = request.QueryString("comando_externo")
cod_pedido_compra = Request.QueryString("cod_pedido_compra")
cod_pedido_compra_produto = Request.QueryString("cod_pedido_compra_produto")
imprimir_cabecalho = Cstr(""&Request.QueryString("ic"))

redirecionar_incluir = "ppv_compras_pedidos_formulario_consulta_produtos.asp"
redirecionar_alterar = "ppv_compras_pedidos_formulario_consulta_produtos.asp"
redirecionar_excluir = "ppv_compras_pedidos_formulario_consulta_produtos.asp"


AbrirConexao()

Set rs = objConexao.execute("SELECT PEDIDO_MANUAL, status_pedido_compra, cod_fornecedor FROM TB_PEDIDOS_COMPRAS WHERE cod_pedido_compra = "&sqlNumero(cod_pedido_compra))
if not rs.Eof Then
	pedido_manual = rs(0)
	status_pedido_compra = rs(1)
	cod_fornecedor = rs(2)
end if
rs.Close
set rs = Nothing

FecharConexao()

if comando = "" and comando_externo <> "visualizar" and comando_externo <> "excluir" Then
	comando = "incluir"
elseif comando = "" and comando_externo = "visualizar" or comando_externo = "excluir" Then
	comando = "visualizar"
end if

'LISTAGEM DE COMANDO
If comando = "incluir" and campo1 = Empty Then
'.....
ElseIf comando = "incluir" and campo1 <> Empty Then
'Incluir novo registro
incluirRegistro(redirecionar_incluir)
titulo_pagina = titulo_pagina&" / Inclusão"
Elseif comando = "alterar" and campo1 = Empty Then
'Leitura inicial para possível Alteração
lerRegistro()
cod_pedido_compra = Request.QueryString("cod_pedido_compra")
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "alterar" and campo1 <> Empty Then
'Gravar Alteração
cod_pedido_compra = Request.Form("cod_pedido_compra")
alterarRegistro(redirecionar_alterar)
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 
'Leitura inicial para possível Exclusão
cod_pedido_compra = Request.QueryString("cod_pedido_compra")
lerRegistro()
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
'Exclui registro
cod_pedido_compra = Request.Form("cod_pedido_compra")
excluirRegistro(redirecionar_excluir)
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
'Visualiza registro

titulo_pagina = titulo_pagina&" / Visualização"
End If


sub lerFormulario()

	cod_pedido_compra = request.form("cod_pedido_compra")
	cod_pedido_compra_produto = request.form("cod_pedido_compra_produto")
	cod_produto = request.form("cod_produto")
	quantidade_solicitada = request.form("quantidade_solicitada")
	obs = request.form("obs")

end sub


sub lerRegistro()
	
	AbrirConexao()

	response.Write campo1

	set rs = objConexao.execute("SELECT * FROM TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra_produto = "&sqlNumero(cod_pedido_compra_produto) )
		if not rs.Eof Then
			
			cod_produto = rs("cod_produto")
			quantidade_solicitada = rs("quantidade_solicitada")
			obs = rs("obs")
			status_produto = rs("status_produto")

		end if 
	rs.close
	set rs = nothing

	FecharConexao()

end sub


sub incluirRegistro (redirecionar)

	lerFormulario()

	AbrirConexao()

		if (cod_produto <> "" and int(quantidade_solicitada) > 0 and obs <> "") Then

			objConexao.execute("INSERT INTO TB_PEDIDOS_COMPRAS_PRODUTOS (cod_produto, quantidade_solicitada, quantidade_entregue, obs, cod_pedido_compra, valor_custo_solicitado) SELECT "&sqlNumero(cod_produto)&","&sqlNumero(quantidade_solicitada)&",0 ,"&sqlTexto(obs)&","&sqlNumero(cod_pedido_compra)&", valor_custo FROM TB_FORNECEDOR_PRODUTOS WHERE cod_produto = "&sqlNumero(cod_produto)&" AND cod_fornecedor = "&sqlNumero(cod_fornecedor))

		end if 

	FecharConexao()

	Response.redirect(redirecionar&"?comando_externo="&comando_externo&"&cod_pedido_compra="&cod_pedido_compra)
	
end sub

sub alterarRegistro(redirecionar)

	lerFormulario()
	
	AbrirConexao()

	if (cod_produto <> "" and int(quantidade_solicitada) > 0 and obs <> "") Then

		objConexao.execute("UPDATE VW_TB_PEDIDOS_COMPRAS_PRODUTOS SET "&_
				"cod_produto = "&sqlNumero(cod_produto)&_
				",quantidade_solicitada = "&sqlNumero(quantidade_solicitada)&_
				",obs = "&sqlTexto(obs)&_
				" WHERE cod_pedido_compra_produto = "&sqlNumero(cod_pedido_compra_produto))

	end if

	FecharConexao()

	Response.redirect(redirecionar&"?comando_externo="&comando_externo&"&cod_pedido_compra="&cod_pedido_compra)

end sub
 
sub excluirRegistro(redirecionar)


	lerFormulario()
	
	AbrirConexao()

		objConexao.execute("DELETE FROM TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra_produto = "&sqlNumero(cod_pedido_compra_produto))
	
	FecharConexao()

	Response.redirect(redirecionar&"?comando_externo="&comando_externo&"&cod_pedido_compra="&cod_pedido_compra)

end sub
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="http://stillos.com.br/gerencial_foto/css/base.css" media="screen"></link>
<link rel="stylesheet" media="all" type="text/css" href="http://stillos.com.br/gerencial_foto/css/modal-message.css">
</head>
<body style="padding:10px; background-color:#FFF">
	<%If comando = "excluir" Then%>
		<div class='msgExcluir'>Atenção ao excluir o item a Ordem de Serviço voltará para a etapa de compra necessário, mesma que esteja finalizada.<br />Não voltará para etapa de compra apenas no caso da OS estiver como Cancelada.</div>
	<%End If%>
	<%
	AbrirConexao()

	if pedido_manual = "S" Then

		if  comando <> "visualizar" and status_pedido_compra < 2 then
		%>
		<form method="POST">
		<table align="center" style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC">
			<tr>
				<td>
					<div align='center'><strong>Produto</strong></div>
				</td>
				<td>
					<div align='center'><strong>Quantidade</strong></div>
				</td>
				<td>
					<div align='center'><strong>Observação</strong></div>
				</td>
				<td>
					<div align='center'>&nbsp;</div>
				</td>
			</tr>
			<tr>
				<td>
					<div align='center'><%=padraoFormularioSelect("cod_produto", "", "inputTamanho2", "","","Selecione",cod_produto, "SELECT TB_FORNECEDOR_PRODUTOS.cod_produto, TB_PRODUTOS.nome_produto FROM TB_PRODUTOS JOIN TB_FORNECEDOR_PRODUTOS ON TB_PRODUTOS.cod_produto = TB_FORNECEDOR_PRODUTOS.cod_produto WHERE TB_PRODUTOS.ativo = 1 AND TB_FORNECEDOR_PRODUTOS.cod_fornecedor = "&sqlNumero(cod_fornecedor), comando, false, false, false )%></div>
					</td>
					<td>
						<div align='center'><input type="text" class="inputTamanho1 padrao_numero" name='quantidade_solicitada' value="<%=quantidade_solicitada%>" style="text-align:right; width:60px;" /></div>
					</td>
					<td>
						<div align='center'><input type="text" class="inputTamanho2" name='obs' value='<%=obs%>' /></div>
					</td>
					<td>
						<input type="submit" value='  <%=Ucase(comando)%>  '>
					</td>

				</tr>
			</table>
			<input type="hidden" name="campo1" value="1">
			<input type="hidden" name="comando" value="<%=comando%>">
			<input type="hidden" name="comando_externo" value="<%=comando_externo%>">
			<input type="hidden" name="cod_pedido_compra" value="<%=cod_pedido_compra%>">
			<input type="hidden" name="cod_pedido_compra_produto" value="<%=cod_pedido_compra_produto%>">
		</form>
		<%
		end if

		if comando = "excluir" then %>
			<button onclick="window.location.href='<%=redirecionar_excluir%>?comando_externo=<%=comando_externo%>&cod_pedido_compra=<%=cod_pedido_compra%>'"> Cancelar </button>
		<%end if%>
		<%
		AbrirConexao()

		set rs = objConexao.execute("SELECT *, TB_PEDIDOS_COMPRAS_PRODUTOS.quantidade_solicitada * TB_PEDIDOS_COMPRAS_PRODUTOS.valor_custo_solicitado as valor_custo_total FROM TB_PEDIDOS_COMPRAS_PRODUTOS JOIN TB_PRODUTOS ON TB_PEDIDOS_COMPRAS_PRODUTOS.cod_produto = TB_PRODUTOS.cod_produto WHERE cod_pedido_compra = "&sqlNumero(cod_pedido_compra)) 

		if not rs.Eof Then
			%>
			<table width="99%" border="1" align="center" cellpadding="5" style="border-collapse:collapse; text-align: right;">
	<tr bgcolor="whitesmoke">
				<tr>
					<th><center>Produto</center></th>
					<th><center>Quantidade</center></th>
					<th><center>Valor Un.</center></th>
					<th><center>Valor Total</center></th>
					<th><center>Observação</center></th>
					<th><center>Stauts Produto</center></th>
					<th>&nbsp;</th>
				</tr>
			<%
			While not rs.Eof
				%>
					<tr>
						<td align="left"><%=rs("nome_produto")%></td>
						<td><%=rs("quantidade_solicitada")%></td>
						<td>R$ <%=arrumaNumero(rs("valor_custo_solicitado"),0,2)%></td>
						<td>R$ <%=arrumaNumero(rs("valor_custo_total"),0,2)%></td>
						<td align="left"><%=rs("obs")%></td>
						<td align="center"  style="font-size:11px">
							<div id="containerProdutoPedidoCompra_<%=base64Encode(rs("cod_pedido_compra_produto"))%>">
								<%If rs("status_produto") = "E" Then%>
									<div class="contasPagarReceberStatusQ" style="text-decoration:none; cursor:auto;">Entregue</div>
						
								<%Elseif rs("status_produto")  = "P" Then %>    
									<div class="contasPagarReceberStatusA_nao_vencido" style="text-decoration:none; cursor:auto;">Pendente</div>
					
								<%Elseif rs("status_produto")  = "C" Then%>
									<div class="contasPagarReceberStatusA_vencido" style="text-decoration:none; background-color:#999999; cursor:auto;">Cancelado</div>
								<%End If%>
							</div>    
						</td>
						<td align="center" width="110px">
							<% if status_pedido_compra < 2 then%>
							<table border="0" align="center" cellpadding="3" cellspacing="0">
      						<tr>
                    			<td>
										<a href="ppv_compras_pedidos_formulario_consulta_produtos.asp?comando_externo=<%=comando_externo%>&comando=alterar&cod_pedido_compra=<%=cod_pedido_compra%>&cod_pedido_compra_produto=<%=rs("cod_pedido_compra_produto")%>">
											<img src="images/icons/icon_alterar.png" alt="Alterar Registro" title="Alterar Registro" width="16" height="16" border="0">
										</a>
									</td>
									<td>
            						<a href="ppv_compras_pedidos_formulario_consulta_produtos.asp?comando_externo=<%=comando_externo%>&comando=excluir&cod_pedido_compra=<%=cod_pedido_compra%>&cod_pedido_compra_produto=<%=rs("cod_pedido_compra_produto")%>">
              							<img src="images/icons/icon_excluir.png" alt="Excluir Registro" title="Excluir Registro" border="0">
              						</a>
              					</td>
        						</tr>
							</table>
							<%
							else 
								response.Write ("Não é possível alterar produtos após a aprovação") 
							end if
							%>
						</td>
					</tr>
				<%
			rs.MoveNext
			Wend
			%>
			<table>
			<%
		end if
		rs.Close
		set rs = nothing


		FecharConexao()
	Else
	AbrirConexao()

		If imprimir_cabecalho = "S" Then
			Set rs=objConexao.Execute("SELECT * FROM VW_TB_PEDIDOS_COMPRAS WHERE cod_pedido_compra = "&cod_pedido_compra)
				If not rs.Eof Then
					cod_fornecedor = rs("cod_fornecedor")
					numero_pedido_compra = rs("NUMERO_PEDIDO_COMPRA")
				End If
			rs.Close
			Set rs=Nothing
		
			Set rs=objConexao.Execute("SELECT cod_fornecedor, razao_social FROM TB_FORNECEDORES WHERE cod_fornecedor = "&cod_fornecedor&" ORDER BY razao_social")
				If not rs.Eof Then
					razao_social = Ucase(""&rs("razao_social"))
				End If
			rs.Close
			Set rs=Nothing
			Response.Write("<h1>"&numero_pedido_compra&" - "&razao_social&"</h1>")
		End If
	Set rs=objConexao.Execute("SELECT VW_TB_ORDENS_SERVICO.NOME_FORMANDO, VW_TB_ORDENS_SERVICO.ENTIDADE_ENSINO, VW_TB_PEDIDOS_COMPRAS_PRODUTOS.* FROM VW_TB_PEDIDOS_COMPRAS_PRODUTOS  LEFT JOIN VW_TB_ORDENS_SERVICO on VW_TB_ORDENS_SERVICO.COD_ORDEM_SERVICO = VW_TB_PEDIDOS_COMPRAS_PRODUTOS.COD_ORDEM_SERVICO WHERE VW_TB_PEDIDOS_COMPRAS_PRODUTOS.cod_pedido_compra = "&sqlTexto(cod_pedido_compra)&" ORDER BY VW_TB_ORDENS_SERVICO.NOME_FORMANDO ASC")

	If not rs.Eof Then
	%>
	

	<table width="99%" border="1" align="center" cellpadding="5" style="border-collapse:collapse; text-align: right;">
	<tr bgcolor="whitesmoke">
		<td width="179"><div align="center"><strong>OS</strong></div></td>
		<td width="289"><div align="center"><strong>Formando</strong></div></td>
		<td width="178"><div align="center"><strong>Institui&ccedil;&atilde;o de Ensino</strong></div></td>
		<td width="168"><div align="center" style="font-weight: bold" align="center">Produto</div></td>
		<td width="64"><div align="center" style="font-weight: bold" align="center">Qtd.</div></td>
		<td width="156"><div align="center"><strong>Obs.</strong></div></td>

		<td align="center"><div align="center"><span style="font-weight: bold">Data Prevista Entrega</span></div></td>
		<td width="80">&nbsp;</td>
	</tr>
	<%
	While not rs.Eof
		status_produto = Cstr(""&rs("status_produto"))
		dias_atrasados = formatNumber(rs("dias_atrasados"))
	%>
	<tr <%If int(cod_pedido_compra_produto) = int(rs("cod_pedido_compra_produto")) Then%> bgcolor="#FFFFD2" <%End If%>>
		<td>
			<div align="center">
				<%=moduloPPVOSMascara(rs("cod_ordem_servico"),"",true)%>
			</div>
		</td>
		<td><div align="left"><%=rs("nome_formando")%></div></td>
		<td><div align="left"><span style="font-size:12px"><%=Ucase(rs("ENTIDADE_ENSINO"))%></span></div></td>
		<td><div align="left"><%=rs("nome_produto")%></div></td>
		<td><div align="center"><%=ArrumaNumero(rs("quantidade_solicitada"),"-",0)%></div></td>
		<td><div align="left"><%=rs("obs")%></div></td>
		<td><div align="center"><%=ArrumaData(rs("data_previsao_entrega"))%></div></td>
		<td align="center"  style="font-size:11px">
			<div id="containerProdutoPedidoCompra_<%=base64Encode(rs("cod_pedido_compra_produto"))%>">
				<%If status_produto = "E" Then%>
					<div class="contasPagarReceberStatusQ" style="text-decoration:none; cursor:auto;">Entregue</div>
				<%Elseif status_produto = "P" and dias_atrasados = 0 Then%>    
					<div class="contasPagarReceberStatusA_nodia" style="text-decoration:none; cursor:auto;">Pendente</div>
				<%Elseif status_produto = "P" and dias_atrasados < 1 Then%>    
					<div class="contasPagarReceberStatusA_nao_vencido" style="text-decoration:none; cursor:auto;">Pendente</div>
				<%Elseif status_produto = "P" and dias_atrasados > 0 Then%>
					<div class="contasPagarReceberStatusA_vencido" style="text-decoration:none; cursor:auto;">Pendente</div>
				<%Elseif status_produto = "C" Then%>
					<div class="contasPagarReceberStatusA_vencido" style="text-decoration:none; background-color:#999999; cursor:auto;">Cancelado</div>
				<%End If%>
			</div>    
		</td>
	</tr>

	<%
rs.MoveNext()
Wend
%>
</table>
<%Else%>
	<div class="msgGrande">Não há produtos inseridos neste pedido de compra!</div>
<%End If

rs.Close
Set rs=Nothing
FecharConexao()
end if ' end if tipo pedido


%>
<!-- #include virtual="gerencial_foto/files/includes/scripts_rodape.asp" -->
</body>
</html>