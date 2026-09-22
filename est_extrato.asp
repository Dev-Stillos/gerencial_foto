<!-- #include virtual="gerencial_foto/includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include virtual="gerencial_foto/includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include virtual="gerencial_foto/files/includes/functions_local.asp" -->

<%
abrirConexao()

dim cod_produto, periodo_inicio, periodo_fim
dim saldo_anterior, incluir_saldo_anterior

cod_produto = request.queryString("cod_produto")&""
periodo_parametro = request.queryString("periodo_parametro")&""
periodo_inicio = request.queryString("periodo_inicio")&""
periodo_fim = request.queryString("periodo_fim")&""

campo_data = " data_cadastro "

dim sqlQuery, sqlWhere, sqlOrderBy 

sqlQuery = "SELECT * FROM VW_TB_ESTOQUE_MOVIMENTACOES WHERE "
sqlOrderBy = " order by data_cadastro asc"

if cod_produto <> "" then
	sqlWhere = sqlWhere & " AND cod_produto = "&sqlNumero(cod_produto)
end if

if periodo_inicio <> "" and periodo_fim <> "" then

	sqlWhere = sqlWhere & " AND "&campo_data&" BETWEEN "&sqldata(periodo_inicio&" 00:00:00")&" AND "&sqldata(periodo_fim&" 23:59:59")
end if

if periodo_inicio <> "" and periodo_fim = "" then
	sqlWhere = sqlWhere & " AND "&campo_data&" > " & sqldata(periodo_inicio&" 00:00:00")
end if 

if periodo_inicio = "" and periodo_fim <> "" then
	sqlWhere = sqlWhere & " AND "&campo_data&" < " & sqldata(periodo_fim&" 23:59:59")
end if

if sqlWhere <> "" then
	sqlWhere = mid(sqlWhere,5, len(sqlWhere)-4)
end if

%>
<html>
<head>
	<title>PADRAO RELATORIO</title>
	<style type="text/css">
	body {font-family:arial;}
	table{font-size: 12px; width: 99%; margin-bottom: 20px;}
	table th{font-size: 14px;}
	table tr{text-align: center;}
	table tr td, table tr th{padding: 3px;}
	table tr td h3{border-bottom: 1px solid #000; text-align: left; margin: 10px 0px 3px 0px; padding: 0px;}
	.tableCollapse{border-collapse: collapse;}
	</style>
</head>
<body>
	<table>
		<tr>
			<td width="100px"><img src="images/layout/logomarca_foto_reportagens.jpg" /></td>
			<td align="left"><h1>Extrato de movimentação produto</h1></td>
		</tr>
	</table>
	
	<table class="tableCollapse" border='1'>
		<tr>
			<td align="left" colspan="2" ><strong>Gerado por:</strong> 
			<%
			set rs = objconexao.execute("SELECT NOME FROM TB_ADM_USUARIOS WHERE cod_usuario = "&sqlNumero(session("cod_usuario")))
			if not rs.Eof then
				response.write (rs(0))
			end if
			rs.close()
			set rs = nothing
			%>
			<strong>em</strong> <%=arrumaData(now())%> </td>
		</tr>
		<tr>
			<td align="left" width="50%"><strong>Produto:</strong>
			<%
				set rs = objconexao.execute("SELECT Nome_produto FROM TB_PRODUTOS WHERE cod_produto = "&sqlNumero(cod_produto))
				if not rs.Eof then
					response.write(rs("Nome_produto"))
				end if
				rs.close
				set rs = nothing
			%>
			</td>
			<td align="right" width="50%"><strong>Período:</strong>
			<%
				if periodo_inicio <> "" and periodo_fim <> "" then
					response.write("Data "&periodo_parametro&" "& periodo_inicio&" a "&periodo_fim)
				elseif periodo_inicio <> "" and periodo_fim = "" then
					response.write("Data "&periodo_parametro&" desde " & periodo_inicio)
				elseif periodo_inicio = "" and periodo_fim <> "" then
					response.write("Data "&periodo_parametro&" até " & periodo_inicio)
				else
					response.write("Não especificado")
				end if
			%>
			</td>
		</tr>
	</table>
	<%
	'response.write (sqlQuery & sqlWhere & sqlOrderBy)
	Dim agrupador : agrupador = ""
	set rs = objConexao.Execute(sqlQuery & sqlWhere & sqlOrderBy)
	if not rs.Eof then
	%>
	<table class="tableCollapse">
		<tr>
			<th>Data movimentação</th>
			<th>Descrição</th>
			<th>Tipo</th>
			<th>Quantidade</th>
		</tr>	
		<%
		while not rs.Eof
			if agrupador <> rs("cod_produto") then
				agrupador = rs("cod_produto")
				%>
				<tr>
					<td colspan="8"><h3><%=rs("cod_produto") & " - " &rs("Nome_produto")%></h3></td>
				</tr>
				<%
			end if %>
		<tr>
			<td><%=arrumadata(mid(rs("data_cadastro"),1,10))%></td>
			<td align="left"><%=rs("descricao")%></td>
			<td><%=rs("tipo")%></td>
			<td align="right"><%=rs("quantidade")%></td>
		</tr>
		<%
		rs.movenext()
		Wend
		%>
	</table>
	<%
		dim saldo_final

		if periodo_fim <> ""  then
	    	sqlWhere = " AND data_cadastro < "&sqldata(periodo_fim&" 23:59:59")
		end if 
		
		set rs2 = objconexao.execute("SELECT sum(quantidade) FROM TB_ESTOQUE_MOVIMENTACOES WHERE cod_produto = "&sqlNumero(cod_produto)&sqlWhere)
		if not rs2.Eof then
			saldo_final = rs2(0)
		end if
		rs2.close()
		set rs2 = nothing
	

		%>
		<table border="0" style="font-size:22px">
			<tr>
				<td width="50%" align="left">Total:</td>
				<td width="50%" align="right"><%=saldo_final%></td>
			</tr>
		</table>
		<%
	end if
	rs.close
	set rs = nothing
	%>
</body>
</html>