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
<script type="text/javascript">
	$(function(){
		$('.showTB').click(function(event){
			var d = $(this).attr('cod-produto');
			if ($(".produto_"+d).css("display") == 'none')
				$(".produto_"+d).css("display","table-row");
			else
				$(".produto_"+d).css("display","none");
		});
	});
</script>
<%
	AbrirConexao()

	'/////////////////////////////////////////////////////////////	
	Dim sql, sqlWhere, sqlOrderBy
	sql = "SELECT cod_produto, nome_produto, sum(quantidade) as pos_estoque, qtd_minima FROM VW_TB_ESTOQUE_MOVIMENTACOES "
	sqlOrderBy = " group by cod_produto, nome_produto order by nome_produto asc"
	
	'/////////////
	'Filtro
	Dim filtro_cod_produto
	Dim filtro_somente_ativos
	
	filtro_cod_produto = Cstr(""&Request.QueryString("filtro_cod_produto"))
	filtro_somente_ativos = Cstr(""&Request.QueryString("filtro_somente_ativos"))
	
	If Request.QueryString("filtro") <> "" Then
	
		If filtro_cod_produto <> "" Then
			sqlWhere = sqlWhere & " AND cod_produto = "&sqlNumero(filtro_cod_produto)&" "
		End if
		
		If filtro_somente_ativos <> "" Then
			'sqlWhere = sqlWhere & " AND situacao_produto = "&sqlTexto(filtro_somente_ativos)&" "
		End if

		if len(sqlWhere) > 0 Then
			sqlWhere = " WHERE " & mid(sqlWhere, 5, len(sqlWhere)-5)
		end if
	End If
	sql = sql&sqlWhere&sqlOrderBy
	%>
<h2><%=titulo_pagina%></h2><br />

<div class="subMenu" style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC">
	<form action="<%=caminho_pagina%>" name="FORMULARIO_FILTRO" method="get">
		<table>
			<tr>
				<th rowspan="3">
					<strong>Filtro</strong>
				</th>
			</tr>

			<tr>
				<td><strong>Produto</strong></td>
				<td>
					<%=padraoFormularioSelect("filtro_cod_produto","","inputTamanho2","","","Selecione", filtro_cod_produto,"SELECT cod_produto, nome_produto FROM TB_PRODUTOS order by nome_produto", "incluir", false, false, true)%>
				</td>
				<td rowspan="2"><input type="submit" class="btnFiltrarAmarelo" name="filtro" value=" "></td>
			</tr>
			<tr>
				<td><strong>Somente produtos ativos</strong></td>
				<td>
					<select name="filtro_somente_ativos" id="" class="inputtamanho1">
						<option value="S">Sim</option>
						<option value="N" selected>Não</option>
					</select>
				</td>
			</tr>
			<tr>
				<td colspan="2" align="center"></td>
			</tr>
		</table>
	</form>
</div>
<%
	Set rs=server.CreateObject("ADODB.recordset")
	rs.CursorLocation=3  'Cursor no cliente
	rs.PageSize=50
	rs.Open sql, objConexao

	Dim paginacaocao
	paginacaocao = formatNumber(0&Request.QueryString("paginacao"),0)

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
	    	<td width="47%">Página <strong><%=rs.AbsolutePage%></strong> de <%=rs.PageCount%></td>
	    	<td width="53%"><div align="right">Total <strong><%=rs.RecordCount%></strong> Movimentações</div></td>
	  	</tr>
	</table>

	<table width="99%" border="1" align="center" cellpadding="5" style="border-collapse:collapse; text-align: right;">
		<tr bgcolor="whitesmoke">
			<td width="7%"><div align="center" style="font-weight: bold">Cód. Produto</div></td>
			<td width="20%"><div align="center" style="font-weight: bold">Produto</div></td>
			<td width="10%"><div align="center" style="font-weight: bold">Qtd. Minima</div></td>
			<td width="8%"><div align="center" style="font-weight: bold">Qtd. em estoque</div></td>
		</tr>
		<%
		Dim x
		For x=1 To rs.PageSize
			If Not rs.Eof Then
			%>
			<tr>
				<td><div align="center">
					<a href="#" class="showTB" cod-produto="<%=rs("cod_produto")%>"><%=rs("cod_produto")%></a>
				</div></td>
			  	<td><div align="left">
			  		<a href="#" class="showTB" cod-produto="<%=rs("cod_produto")%>"><%=rs("nome_produto")%></a>
			  	</div></td>
			  	<td><div align="right"><%=arrumaNumero(rs("qtd_minima"),0,0)%></div></td>
			  	<td><div align="center">
					<span1>
						<%=rs("pos_estoque")%>
					</span>
				</div></td>
			</tr>
			<tr style="display:none" class='produto_<%=rs("cod_produto")%>'>
				<td colspan="9" align="left">
					<% 
					set rs2 = objConexao.execute("SELECT * FROM VW_TB_ESTOQUE_MOVIMENTACOES WHERE cod_produto = "&sqlNumero(rs("cod_produto"))&" order by data_cadastro desc limit 10")
					if not rs2.eof then
					%>
					<h5>Últimas 10 Movimentações</h5>
					<table style="border-collapse:collapse;" border="1" cellpadding="5" width="99%" align="center">
						<tr bgcolor="whitesmoke">
							<th>
								<div align="center" >Data movimentação</div>
							</th>
							<th>
								<div align="center" >Tipo movimentação</div>
							</th>
							<th>
								<div align="center" >Descrição</div>
							</th>
							<th>
								<div align="center" >Origem</div>
							</th>
							<th>
								<div align="center" >Quantidade</div>
							</th>
						</tr>
						<%
						while not rs2.eof
						%>
						<tr>
							<td>
								<div align="center"><%=arrumaData(rs2("data_cadastro"))%></div>
							</td>
							<td>
								<div align="center"><%

								if rs2("tipo") = "E" then
									response.write "Entrada"
								elseif rs2("tipo") = "S" then
									response.write "Sa�da"
								end if

								%></div>
							</td>
							<td>
								<div><%=rs2("descricao")%></div>
							</td>
							<td>
								<div>
									<%=estoqueLink(rs2)%>
								</div>
							</td>
							<td align="right">
								<%
									if rs2("quantidade") > 0 then
										color = "00C"
									else
										color = "C00"
									end if
								%>
								<div style="font-weight:bold; color:#<%=color%>"><%=rs2("quantidade")%></div>
							</td>
						</tr>
						<%
						rs2.MoveNext
						Wend
						%>
					</table>
					<%
					end if

					%>
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
		Nenhuma movimentação encontrada!</div>
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