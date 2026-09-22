<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
abrirConexao()

dim cod_vendedor, periodo_inicio, periodo_fim

cod_vendedor = request.queryString("cod_vendedor")&""
periodo_parametro = request.queryString("periodo_parametro")&""
periodo_inicio = request.queryString("periodo_inicio")&""
periodo_fim = request.queryString("periodo_fim")&""

dim sqlQuery, sqlWhere, sqlOrderBy 

sqlQuery = "SELECT * FROM VW_TB_ESTOQUE_MOVIMENTACOES_VENDEDORES WHERE tipo = 'E' AND quantidade > quantidade_saida "
sqlOrderBy = " order by nome_vendedor, data_movimentacao"

if cod_vendedor <> "" then
	sqlWhere = sqlWhere & " AND cod_vendedor = "&sqlNumero(cod_vendedor)
end if

if cod_produto <> "" then
	sqlWhere = sqlWhere & " AND cod_produto = "&sqlNumero(cod_produto)
end if

if periodo_parametro = "devolucao" then
	campo_data = "data_devolucao"
elseif periodo_parametro = "retirada" then
	campo_data = "data_movimentacao"
else
	campo_data = ""
end if

if periodo_inicio <> "" and periodo_fim <> "" then
	'periodo_inicio = sqldata(periodo_inicio&" 00:00:00")
	'periodo_fim = sqldata(periodo_fim&" 23:59:59")
	sqlWhere = sqlWhere & " AND "&campo_data&" BETWEEN "&sqldata(periodo_inicio&" 00:00:00")&" AND "&sqldata(periodo_fim&" 23:59:59")
end if

if periodo_inicio <> "" and periodo_fim = "" then
	'periodo_inicio = sqldata(periodo_inicio&" 00:00:00")

	sqlWhere = sqlWhere & " AND "&campo_data&" > " & sqldata(periodo_inicio&" 00:00:00")
end if 

if periodo_inicio = "" and periodo_fim <> "" then
	'periodo_fim = sqldata(periodo_fim&" 23:59:59")
	sqlWhere = sqlWhere & " AND "&campo_data&" < " & sqldata(periodo_fim&" 23:59:59")
end if

%>
<html>
<head>
	<title>Itens por vendedores</title>
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
			<td align="left"><h1>Itens com vendedores</h1></td>
		</tr>
	</table>
	<table class="tableCollapse" border='1'>
		<tr>
			<td align="left" colspan="3" ><strong>Gerado por:</strong>
			<%
			set rs = objconexao.execute("SELECT NOME FROM TB_ADM_USUARIOS WHERE cod_usuario = "&sqlNumero(session("cod_usuario")))
			if not rs.Eof then
				response.write (rs(0))
			end if
			rs.close()
			set rs = nothing
			%> <strong>em</strong> <%=arrumaData(now())%>	</td>
		</tr>
		<tr>
			<td align="left" width="33%"><strong>Vendedor:</strong>
			<%
				if cod_vendedor <> "" then
					set rs = objConexao.Execute("SELECT nome_vendedor FROM TB_VENDEDORES WHERE cod_vendedor = "&sqlNumero(cod_vendedor))
					if not rs.Eof then
						response.write(rs(0))
					end if
					rs.close
					set rs = nothing
				else
					response.write("Não especificado")
				end if
			%>
			</td>
			<td align="left" width="33%"><strong>Produto:</strong>
			<%
				if cod_produto <> "" then
					set rs = objConexao.Execute("SELECT nome_produto FROM TB_PRODUTOS WHERE cod_produto = "&sqlNumero(cod_produto))
					if not rs.Eof then
						response.write(rs(0))
					end if
					rs.close
					set rs = nothing
				else
					response.write("Não especificado")
				end if
			%>
			</td>
			<td align="right" width="33%">
				<strong>Período:</strong> 
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
	Dim agrupador : agrupador = ""

	'response.write(sqlQuery & sqlWhere & sqlOrderBy)
	
	set rs = objConexao.Execute(sqlQuery & sqlWhere & sqlOrderBy)
	if not rs.Eof then
	%>
	<table class="tableCollapse">
		<tr>
			<th>Produto</th>
			<th>Quantidade</th>
			<th>Data Reitrada</th>
			<th>Data Devolução</th>
		</tr>
		<% 
		while not rs.Eof
			if agrupador <> rs("cod_vendedor") then
				agrupador = rs("cod_vendedor")
				%>
				<tr>
					<td colspan="8"><h3><%=rs("cod_vendedor") & " - " &rs("nome_vendedor")%></h3></td>
				</tr>
				<%
			end if

			quantidade_posse = int(rs("quantidade")) - int(rs("quantidade_saida"))
		%>
		<tr>
			<td align="left"><%=rs("nome_produto")%></td>
			<td><%=quantidade_posse%></td>
			<td><%=arrumadata(mid(rs("data_movimentacao"),1,10))%></td>
			<td><%=arrumadata(mid(rs("data_devolucao"),1,10))%></td>
		</tr>
		<% 
		rs.movenext
		Wend
		%>
	</table>
	<% 
	else
	%>
		<div class="msgGrandeErro">Nenhum resultado encontrado para os parametros informados</div>
	<%
	end if
	fecharConexao()
	%>
</body>
</html>