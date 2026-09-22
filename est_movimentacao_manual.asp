<%
Dim calendario_necessario, BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
calendario_necessario = "SIM"
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 1
BTN_IMPRIMIR_LINK = "#"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<h2>Compras</h2>
<!-- #include file="includes/estoque_sub_menu.asp" -->

<div class="subMenu">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td width="4%">A&ccedil&otildees:</td>
			<td width="96%">
				<a href="est_movimentacao_manual_formulario.asp">Incluir movimenta&ccedil&atildeo manual</a>
			</td>
		</tr>
	</table>
</div>

<%
	AbrirConexao()
	
	'/////////////////////////////////////////////////////////////	
	Dim sql, sqlWhere, sqlOrderBy
	sql = "SELECT * FROM VW_TB_ESTOQUE_MOVIMENTACOES "
	sqlOrderBy = " order by data_cadastro desc"
	
	'/////////////
	'Filtro
	Dim filtro_cod_movimentacao
	Dim filtro_cod_produto
	Dim periodo_inicio, periodo_fim
	Dim filtro_origem
	Dim filtro_tipo

	filtro_cod_movimentacao = Cstr(""&Request.QueryString("filtro_cod_movimentacao"))
	filtro_cod_produto = Cstr(""&Request.QueryString("filtro_cod_produto"))
	filtro_tipo = Cstr(""&Request.QueryString("filtro_tipo"))
	filtro_origem = Cstr(""&Request.QueryString("filtro_origem"))

	periodo_inicio = Cstr(""&Request.QueryString("periodo_inicio"))
	periodo_fim = Cstr(""&Request.QueryString("periodo_fim"))

	If Request.QueryString("filtro") <> "" or filtro_cod_movimentacao <> "" Then
		if filtro_cod_movimentacao <> "" then
			sqlWhere = " WHERE cod_movimentacao = "&sqlNumero(filtro_cod_movimentacao) 
		else
			If filtro_cod_produto <> "" Then
				sqlWhere = sqlWhere & " AND cod_produto = "&sqlNumero(filtro_cod_produto)&" "
			End if

			If filtro_tipo <> "" Then
				sqlWhere = sqlWhere & " AND tipo = "&sqlTexto(filtro_tipo)&" "
			End if

			if filtro_origem <> "" then
				sqlWhere = sqlWhere & " AND "&filtro_origem&" is not null "
			end if

			if periodo_fim <> "" or periodo_inicio <> "" then
				periodo_inicio = periodo_inicio & " 0:0:0"
				periodo_fim	= periodo_fim & " 23:59:59"

				if periodo_fim <> "" and periodo_inicio <> "" then
					sqlWhere = sqlWhere & " AND data_cadastro BETWEEN "&sqlData(periodo_inicio)&" and  "&sqlData(periodo_fim)

				elseif periodo_fim <> "" and periodo_inicio = "" then
					sqlWhere = sqlWhere & " AND data_cadastro <= "&sqlData(periodo_fim)

				elseif periodo_fim = "" and periodo_inicio <> "" then
					sqlWhere = sqlWhere & " AND data_cadastro >= "&sqlData(periodo_fim)

				end if
			end if

			if len(sqlWhere) > 0 Then
				sqlWhere = " WHERE " & mid(sqlWhere, 5, len(sqlWhere)-4)
			end if
		end if
	End If
	sql = sql&sqlWhere&sqlOrderBy
	%>
<h2><%=titulo_pagina%></h2><br />

<div class="subMenu" style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC">
	<form action="<%=caminho_pagina%>" name="FORMULARIO_FILTRO" method="get">
		<table cellpadding="4">
			<tr>
				<th colspan="2">
					<h6>Filtrar</h6>
				</th>
			</tr>
			<tr>
				<td align="left"><strong>Produto</strong></td>
				<td>
					<%=padraoFormularioSelect("filtro_cod_produto","","inputTamanho2","","","Todos", filtro_cod_produto,"SELECT cod_produto, nome_produto FROM TB_PRODUTOS order by nome_produto", "incluir", false, false, true)%>
				</td>
			</tr>
			<tr>
				<td><strong>Periodo movimenta&ccedil&atildeo</strong></td>
				<td>
					<table border="0" cellpadding="3">
						<tr>
							<td width="110"><input type="text" class="inputTamanho1 padrao_data" name="periodo_inicio"></td>
							<td width="10" align="center" style="text-weigth: bold;">&agrave</td>
							<td width="110"><input type="text" class="inputTamanho1 padrao_data" name="periodo_fim"></td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td><strong>Origem</strong></td>
				<td>
					<%=padraoFormularioSelect("filtro_origem","","inputTamanho2","","","Todos", filtro_origem,"PADRAO:cod_pedido#Pedido|cod_pedido_compra#Pedido Compra|cod_vendedor#Vendedor|cod_ordem_servico#Ordem servi&ccedilo", "incluir", false, false, true)%>
				</td>
			</tr>
			<tr>
				<td><strong>Tipo</strong></td>
				<td>
					<%=padraoFormularioSelect("filtro_tipo","","inputTamanho2","","","Todos", filtro_tipo,"PADRAO:E#Entrada|S#Saída", "incluir", false, false, true)%>
				</td>
			</tr>
			<tr>
				<td colspan="2" align="center">
					<input type="submit" class="btnFiltrarAmarelo" name="filtro" value=" ">
				</td>
			</tr>
		</table>
	</form>
</div>
<%
	Set rs=server.CreateObject("ADODB.recordset")
	rs.CursorLocation=3  'Cursor no cliente
	rs.PageSize=100
	rs.Open sql, objConexao



	Dim paginacaocao : paginacaocao = formatNumber(0&Request.QueryString("paginacao"),0)

	If not rs.Eof Then
		If paginacaocao > 0 Then
			rs.AbsolutePage=paginacaocao
		Else
			paginacaocao = 1
			rs.AbsolutePage=1
		End If	
	End If

If not rs.Eof Then
%>
	<table width="99%" border="0" align="center">
		<tr>
	    	<td width="47%">P&aacutegina <strong><%=rs.AbsolutePage%></strong> de <%=rs.PageCount%></td>
	    	<td width="53%"><div align="right">Total <strong><%=rs.RecordCount%></strong> Movimenta&ccedil&otildees</div></td>
	  	</tr>
	</table>
	<table width="99%" border="1" align="center" cellpadding="5" style="border-collapse:collapse; text-align: right;">
		<tr bgcolor="whitesmoke">
			<td width="7%"><div align="center" style="font-weight: bold">C&oacuted. Movimenta&ccedil&atildeo</div></td>
			<td width="7%"><div align="center" style="font-weight: bold">Data Movimenta&ccedil&atildeo</div></td>
			<td width="20%"><div align="center" style="font-weight: bold">Produto</div></td>
			<td width="4%"><div align="center" style="font-weight: bold">tipo</div></td>
			<td width="20%"><div align="center" style="font-weight: bold">Origem</div></td>
			<td width="20%"><div align="center" style="font-weight: bold">Descri&ccedil&atildeo</div></td>
			<td width="4%"><div align="center" style="font-weight: bold">Quantidade</div></td>
			<td width="2%"><div align="center" style="font-weight: bold">&nbsp;</div></td>
		</tr>
		<%
		Dim x
		For x=1 To rs.PageSize
			If Not rs.Eof Then
			%>
			<tr>
				<td><div align="center">
					<%=rs("cod_movimentacao")%>
				</div></td>
				<td><div align="center">
					<%=arrumaData(rs("data_cadastro"))%>
				</div></td>
			  	<td><div align="left">
			  		#<%=rs("cod_produto")%> - <%=rs("nome_produto")%>
			  	</div></td>
			  	<td><div align="center">
			  		<%
					if rs("tipo") = "E" then
						response.write "Entrada"
					elseif rs("tipo") = "S" then
						response.write "Sa&iacuteda"
					end if
					%>
			  	</div></td>
			  	<td><div align="left">
					<%=estoqueLink(rs)%>
				</div></td>
				<td><div align="left">
			  		<%=rs("descricao")%>
			  	</div></td>
			  	<td align="right">
					<%
						if rs("quantidade") > 0 then
							color = "00C"
						else
							color = "C00"
						end if
					%>
					<div style="font-weight:bold; color:#<%=color%>"><%=rs("quantidade")%></div>
				</td>
				<td align="center">
				    <table border="0" align="center" cellpadding="3" cellspacing="0">
				    <tr>
				        <td>
				        	<a href="est_movimentacao_manual_formulario.asp?comando=visualizar&cod_movimentacao=<%=rs("cod_movimentacao")%>">
				          		<img src="images/icons/icon_visualizar.png" alt="Visualizar Registro" title="Visualizar Registro" border="0" />
							</a>
						</td> 
						<td>
							<% if  estoqueMovimentacaoManual(rs) then%>
				        	<a href="est_movimentacao_manual_formulario.asp?comando=excluir&cod_movimentacao=<%=rs("cod_movimentacao")%>">
				          		<img src="images/icons/icon_excluir.png" alt="excluir Registro" title="excluir Registro" border="0" />
							</a>
							<% end if %>
						</td>
					</tr>
					</table>
				</td>
			</tr>
			<%
			rs.MoveNext()
			End If
		Next
		If rs.PageCount > 1 Then %>
			<tr bgcolor="whitesmoke">
				<td colspan="11"><%call paginacao(paginacaocao,rs.PageCount)%></td>
			</tr>
		<%End if%>
	</table>
	
	<% Else %>
	<div align="center" class="msgGrande">
		Nenhuma movimenta&ccedil&atildeo encontrada!</div>
	<%End if

	rs.Close
	Set rs=Nothing

	FecharConexao()
%>
<br />
<br />
<br />
<!-- #include file="files/includes/miolo-inferior.asp" -->
</body>
</html>