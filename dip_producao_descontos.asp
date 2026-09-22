<% dim script_valida_form_master : script_valida_form_master = "NAO" %>
<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<!-- #include file="files/includes/functions_local.asp" -->
<%
if not validaUsuario(2,,82, false) then
	response.write("Voc� n�o tem acesso")
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
		<!-- #include virtual="gerencial_foto/files/includes/scripts.asp" -->
	</head>
	<body>
	<!--
		SUB menu de a��es
	 -->
	<div>
	</div>
	<%
	abrirConexao()

	Set rs = ObjConexao.Execute("SELECT * FROM VW_TB_PRODUCOES_DESCONTOS")
	if not rs.Eof then
	%>
	<table border="1" width="99%">
		<tr>
			<th> # </th>
			<th>Descrição</th>
			<th>Aliquota Padrão</th>
			<th>Quantidade de Produtos</th>
			<th>Contratos</th>
		</tr>
	<% While no rs.eof %>
		<tr>
			<td><%=rs("cod_producao_desconto")%></td>
			<td><%=rs("desconto")%></td>
			<td><%=rs("aliquota_padrao")%></td>
			<td><%=rs("qtd_produtos")%></td>
			<td><%=rs("contratos")%></td>
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
			Nenhum desconto encontrado para esta produção.
		</div>
		<%
	end if
	rs.close
	set rs = nothing

	fecharConexao()
	%>
	</body>
</html>