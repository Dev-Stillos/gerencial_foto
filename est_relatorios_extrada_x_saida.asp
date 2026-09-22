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
<!-- #include file="includes/compras_sub_menu.asp" -->
<%
	AbrirConexao()
	
	'/////////////////////////////////////////////////////////////	
	Dim sql, sqlWhere, sqlOrderBy
	sql = "SELECT * FROM VW_TB_ESTOQUE_MOVIMENTACOES "
	sqlOrderBy = " order by nome_produto asc"
	
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
			sqlWhere = sqlWhere & " AND situacao_produto = "&sqlTexto(filtro_somente_ativos)&" "
		End if

		if len(sqlWhere) > 0 Then
			sqlWhere = mid(sqlWhere, 4, len(sqlWhere)-4)
		end if
	End If
	sql = sql&sqlWhere&sqlOrderBy
	%>
<h2><%=titulo_pagina%></h2><br />

<div class="subMenu" style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC">
	<form action="<%=caminho_pagina%>" name="FORMULARIO_FILTRO" method="get">
		<table>
			<thead>
				<tr>
					<th rowspan="3">Filtro</th>
				</tr>
			</thead>
			<tbody>
				<tr>
					<td>Produto</td>
					<td>
						<%=padraoFormularioSelect("filtro_cod_produto","","","","","Selecione", filtro_cod_produto,"SELECT cod_produto, nome_produto FROM TB_PRODUTOS order by nome_produto", "incluir", false, false, false)%>
					</td>
				</tr>
				<tr>
					<td>Somente produtos ativos</td>
					<td>
						<select name="filtro_somente_ativos" id="" class="inputtamanho1">
							<option value="S">Sim</option>
							<option value="N" selected>Não</option>
						</select>
					</td>
				</tr>
				<tr>
					<td colspan="2"><buttom>filtrar</buttom></td>
				</tr>
			</tbody>
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
%>

<table width="99%" border="0" align="center">
	<tr>
    	<td width="47%">Página <strong><%=rs.AbsolutePage%></strong> de <%=rs.PageCount%></td>
    	<td width="53%"><div align="right">Total <strong><%=rs.RecordCount%></strong> Movimentações</div></td>
  	</tr>
</table>

<table width="99%" border="1" align="center" cellpadding="5" style="border-collapse:collapse; text-align: right;">
<tr bgcolor="whitesmoke">
<td width="7%"><div align="center" style="font-weight: bold">C&oacute;d. Pedido Compra</div></td>
<td width="20%"><div align="center" style="font-weight: bold">Fornecedor</div></td>
<td width="14%"><div align="center"><strong>Produto</strong></div></td>
<td width="8%"><div align="center"><strong>Qtd. Pendente</strong></div></td>
<td width="10%" align="center"><div align="center"><strong>Qtd. Entregue</strong></div></td>
<td width="12%" align="center"><div align="center" style="font-weight: bold">Respons&aacute;vel pedido</div></td>
<td width="5%" align="center"><div align="center"><strong>Data Emiss&atilde;o</strong></div></td>
<td width="8%" align="center"><strong> Dias Atrasados</strong></td>
<td width="4%" align="center"><strong>Status</strong></td>
<td width="4%">&nbsp;</td>
<td width="8%">&nbsp;</td>
</tr>
<%
Dim x
For x=1 To rs.PageSize
If Not rs.Eof Then
%>
<tr>
	<td><div align="center"><%=rs("numero_pedido_compra")%></div></td>
  	<td><div align="left"><%=rs("nome_fantasia")%></div></td>
  	<td><div align="left"></div></td>
  	<td><div align="center"><%=qtd_pendente%></div></td>
  	<td><div align="center"><%=qtd_entregue%></div></td>
  	<td><div align="left"> <%=nomeUsuario(rs("cadastro_cod_usuario"),true)%> </div></td>
  	<td><div align="center"><%=ArrumaData(rs("data_cadastro"))%></div></td>
	<td><div align="center"></div></td>
	<td align="center">
    <table border="0" align="center" cellpadding="3" cellspacing="0">
    <tr>
        <td>
        	<a href="ppv_compras_pedidos_imprimir.asp?c=<%=rs("cod_pedido_compra")%>" target="_blank">
        		<img src="images/icons/icon_imprimir.png" alt="" title="" border="0" />
        	</a>
        </td>
        <td>
        	<a href="ppv_compras_pedidos_formulario.asp?comando=visualizar&cod_pedido_compra=<%=rs("cod_pedido_compra")%>">
          		<img src="images/icons/icon_visualizar.png" alt="Visualizar Registro" title="Visualizar Registro" border="0" />
			</a>
		</td> 
		<td>
			<a href="ppv_compras_pedidos_formulario.asp?comando=alterar&cod_pedido_compra=<%=rs("cod_pedido_compra")%>">
				<img src="images/icons/icon_alterar.png" alt="Alterar Registro" title="Alterar Registro" width="16" height="16" border="0" />
			</a>
		</td>
        <td>
        	<a href="ppv_compras_pedidos_formulario.asp?comando=excluir&cod_pedido_compra=<%=rs("cod_pedido_compra")%>">
            	<img src="images/icons/icon_excluir.png" alt="Excluir Registro" title="Excluir Registro" border="0" />
            </a>
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
<br />
<strong>&nbsp;Legenda de Status:</strong><br />


<table width="612" border="0" cellpadding="2" cellspacing="0">
<%
'Criador dinamico de legenda de status de pedido de compra, por enquanto há 7 status válidos (sem contas o status em cadastro)
For i=1 to 4%>
	<tr>
		<%If i=4 Then%>
			<td width="12" align="center">99</td>
			<td width="284"><%=pedidoCompraStatus(99)%></td>
		<%Else%>
			<td width="20" align="center"><%=i%></td>
			<td width="280"><%=pedidoCompraStatus(i)%></td>
		<%End if%>
	</tr>
<% Next %>
</table>

<%
	Else
	%>
    <div align="center" class="msgGrande">
    	Nenhum Pedido de Compra encontrado!</div>
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