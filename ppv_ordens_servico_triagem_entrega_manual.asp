<%
Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
BTN_VOLTAR = 1
BTN_IMPRIMIR = 1
BTN_FONTE = 2
BTN_IMPRIMIR_LINK = "#"

Dim titulo_pagina
titulo_pagina = "Triagem Entrega | Transpostadora"

Dim calendario_necessario
calendario_necessario = "NÃO"
%>
<!-- #include virtual="gerencial_foto/includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include virtual="gerencial_foto/includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="https://stillos.com.br/gerencial_foto/css/base.css" media="screen"></link>
<link type="text/css" rel="stylesheet" href="https://stillos.com.br/gerencial_foto/css/chosen.min.css" media="screen"></link>
<link rel="stylesheet" media="all" type="text/css" href="https://stillos.com.br/gerencial_foto/css/modal-message.css">
<title>StillosWEB - Sistema corporativo - Vers&atilde;o 2.0</title>
<!-- #include virtual="gerencial_foto/files/includes/scripts.asp" -->
<!-- #include virtual="gerencial_foto/files/includes/functions_local.asp" -->

</head>
<body>
<%
campo1 = request.form("campo1")
cod_ordem_servico = request.form("cod_ordem_servico")
data_entregue = request.form("data_entregue")
descricao_produto = request.form("descricao_produto")
cod_pedido = request.form("cod_pedido")

tia = request.form("tia")


if campo1 = "1" then 

	if tia = "" and cod_ordem_servico <> "" then

		if cod_ordem_servico <> ""  and data_entregue <> "" and descricao_produto <> "" then
		abrirConexao

			objconexao.execute("INSERT INTO TB_PEDIDOS_ITENS_ENTREGUES_MANUAL (cod_ordem_servico,  data_entregue, descricao) VALUES ('"&sqlNumero(cod_ordem_servico)&"', "&sqlData(data_entregue)&", "&sqlTexto(descricao_produto)&")")

			retorno = "produtos Inseridos com sucesso"

		fecharConexao
		else

			retorno = "preencha todos os campos"
		end if

	elseif tia <> "" and cod_ordem_servico = "" then

		if tia <> ""  and data_entregue <> "" and descricao_produto <> "" then
		abrirConexao

			objconexao.execute("INSERT INTO TB_PEDIDOS_ITENS_ENTREGUES_MANUAL_MACK (tia,  data_entregue, descricao) VALUES ('"&sqlNumero(tia)&"', "&sqlData(data_entregue)&", "&sqlTexto(descricao_produto)&")")

			retorno = "produtos Inseridos com sucesso"

		fecharConexao
		else

			retorno = "preencha todos os campos"
		end if
	else
		retorno = "preencha todos os campos"
	end if

end if 
%>
<%if retorno <> "" then%>
	<div class="msgGrande">
		<%=retorno%>
	</div>
<%end if%>

<form style="margin:0px" method="post">
	<table width="99%" border="0" align="center" cellpadding="5" cellspacing="0">
		<tr>
			<td>
				<strong>C&oacute;d. OS / Cód. Pedido</strong>
			</td>
			<td>
				<input type="text" value="" name="cod_ordem_servico" class="inputTamanho2 padrao_numero" />
			</td>
		</tr>
		<tr>
			<td colspan="2" align="center">
				<strong> OU </strong>
			</td>
		</tr>
		<tr>
			<td>
				<strong>C&oacute;d. TIA</strong>
			</td>
			<td>
				<input type="text" value="" name="tia" class="inputTamanho2 padrao_numero" />
			</td>
		</tr>
		<tr>
			<td colspan="2" align="center">
				<strong> ---------------------------------------------------------------- </strong>
			</td>
		</tr>
		<tr>
			<td>
				<strong>Data</strong>
			</td>
			<td>
				<input type="text" value="" name="data_entregue" class="inputTamanho2 padrao_data" />
			</td>
		</tr>
		<tr>
			<td>
				<strong>Produto</strong>
			</td>
			<td>
				<input type="text" value="" name="descricao_produto" class="inputTamanho2" />
			</td>
		</tr>
		<tr>
			<td colspan="2"><input type="submit" value="GRAVAR"></td>
		</tr>
	</table>
	<input type="hidden" name="campo1" value="1">	
</form>

</body>
</html>
<!-- #include virtual="gerencial_foto/files/includes/scripts_rodape.asp" -->