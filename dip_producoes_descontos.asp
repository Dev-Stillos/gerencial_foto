<% dim script_valida_form_master : script_valida_form_master = "NAO" %>
<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<!-- #include file="files/includes/functions_local.asp" -->
<%
if not validaUsuario(2,,82, false) then
	response.write("Você não tem acesso")
	response.end
end if

Dim rs, cod_producao
cod_producao = request.queryString("cod_producao")

'---------------------------------------------------------------'
%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
		<link type="text/css" rel="stylesheet" media="all" href="css/bootstrap.min.css" />
		<link type="text/css" rel="stylesheet" media="screen" href="css/base.css"></link>
		<link type="text/css" rel="stylesheet" media="screen" href="css/chosen.min.css"></link>
		<link type="text/css" rel="stylesheet" media="all" type="text/css" href="css/modal-message.css">
		<title>StillosWEB - Sistema corporativo - Vers&atilde;o 2.0</title>
		<style>
			body{background-color: #fff;}
			.container{margin: 3px;}
		</style>
		<!-- #include virtual="gerencial_foto/files/includes/scripts.asp" -->
	</head>
	<body>
	<%
	'INICIO VALIDA��ES
	if not validaUsuario(2,,82, false) then
		%>
		<div class="alert alert-danger">
			<strong>Atençao</strong>
			<br />
			Você não tem acesso a esta página.
		</div>
		<%
		response.end()
	end if

	'Validamos se foi passado o codigo da produção'
	if cod_producao = "" then
		%>
		<div class="alert alert-danger">
			<strong>Atenção</strong>
			<br />
			Houve um erro inesperado, informe o departamento de TI
		</div>
		<%
		response.end()
	end if
	'FIM VALIDA��ES
	%>
	<!--
		SUB menu de a��es
	 -->
	<div class="submenu">
		<strong>Submenu:</strong>
		<a href='dip_producoes_descontos_formulario.asp?comando=incluir&cod_producao=<%=cod_producao%>'>Incluir novo desconto</a>
	</div>
	<div class="container">
	<%
	abrirConexao()


	Set rs = ObjConexao.Execute("SELECT * FROM VW_TB_PRODUCAO_DESCONTOS where cod_producao = "&sqlNumero(cod_producao))
	if not rs.Eof then
	%>
	<table border="1" width="100%">
		<tr>
			<th width="30px"> # </th>
			<th>Descrição</th>
			<th width="60px">Aliquota Padrão</th>
			<th width="80px">Qtd. de Prod.</th>
			<th width="140px">Contratos</th>
			<th width="60px">&nbsp;</th>
		</tr>
		<% While not rs.eof %>
			<tr>
				<td class="text-center"><%=rs("cod_producao_desconto")%></td>
				<td class="text-left"><%=rs("descricao")%></td>
				<td class="text-center"><%=rs("aliquota_padrao")%>%</td>
				<td class="text-center"><%=rs("qtd_produtos")%></td>
				<td class="text-center"><%=rs("contratos")%></td>
				<td class="text-center" width="60px">
					<table border='0'  align="center">
						<tr>
							<td width="20px">
								<a href='dip_producoes_descontos_formulario.asp?comando=visualizar&cod_producao=<%=rs("cod_producao")%>&cod_producao_desconto=<%=rs("cod_producao_desconto")%>'>
									<img src="images/icons/icon_visualizar.png" border="0">
								</a>
							</td>
							<td width="20px">
								<a href='dip_producoes_descontos_formulario.asp?comando=alterar&cod_producao=<%=rs("cod_producao")%>&cod_producao_desconto=<%=rs("cod_producao_desconto")%>'>
									<img src="images/icons/icon_alterar.png" border="0">
								</a>
							</td>
							<td width="20px">
								<a href='dip_producoes_descontos_formulario.asp?comando=excluir&cod_producao=<%=rs("cod_producao")%>&cod_producao_desconto=<%=rs("cod_producao_desconto")%>'>
									<img src="images/icons/icon_excluir.png" border="0">
								</a>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		<%
			rs.movenext
			Wend
		%>
		</table>
	<%
	else
		%>
		<div class='alert alert-warning'>
			Nenhum Desconto encontrado para esta produção.
		</div>
		<%
	end if
	rs.close
	set rs = nothing

	fecharConexao()
	%>
	</div>
	</body>
</html>
<!-- #include virtual="gerencial_foto/files/includes/scripts_rodape.asp" -->