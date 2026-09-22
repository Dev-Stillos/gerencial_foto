<%
Dim calendario_necessario, BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
calendario_necessario = "NAO"
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 1
BTN_IMPRIMIR_LINK = "#"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<h2>COB - Pedidos</h2>
<br>
<!-- #include file="includes/classe_modulo_cob.asp" -->
<!-- #include file="includes/sub_menu_cob.asp" -->
<%
	AbrirConexao()

	'/////////////////////////////////////////////////////////////	
	Dim sql, sqlWhere, sqlOrderBy
	sql = "SELECT pedbol.cod_pedido, pedbol.cod_sacado, sum(pedbol.valor) as valor_total_boletos, ifnull(sum(pedbol.valor_pago_efetivo),0) as valor_total_pago, count(pedbol.parcela) as qtd_parcelas, sac.nome FROM TB_PEDIDOS_BOLETOS pedbol inner join TB_PEDIDOS ped on ped.cod_pedido = pedbol.cod_pedido inner join TB_SACADOS sac on sac.cod_sacado = pedbol.cod_sacado "
	sqlOrderBy = " GROUP BY pedbol.cod_pedido, pedbol.cod_sacado, sac.nome  ORDER BY pedbol.cod_pedido desc"
	
	'/////////////
	'Filtro
	Dim filtro 
	dim filtro_nome, filtro_pedido

	filtro =  Request.QueryString("filtro")
	filtro_nome = Cstr(""&Request.QueryString("filtro_nome"))
	filtro_pedido = Cstr(""&Request.QueryString("filtro_pedido"))
	filtro_cod_vendedor = Cstr(""&Request.QueryString("filtro_cod_vendedor"))

	If filtro <> "" Then
		If filtro_nome <> "" Then
			sqlWhere = sqlWhere & " AND nome like "&sqlTexto("%"&filtro_nome&"%")&" "
		End if

		If filtro_pedido <> "" Then
			sqlWhere = sqlWhere & " AND pedbol.cod_pedido = "&sqlTexto(filtro_pedido)&" "
		End if

		If filtro_cod_vendedor  <> "" Then
			sqlWhere = sqlWhere & " AND pedbol.cod_vendedor = "&sqlTexto(filtro_cod_vendedor)&" "
		End if

		if len(sqlWhere) > 0 Then
			'sqlWhere = " WHERE   finaliza_vendedor = 1    " & mid(sqlWhere, 5, len(sqlWhere)-5)
			sqlWhere = " WHERE   finaliza_vendedor = 1    " & sqlWhere
		end if
	Else
		sqlWhere = " WHERE finaliza_vendedor = 1  "
	End If
	sql = sql&sqlWhere&sqlOrderBy


	'response.write sql
	%>
<h2><%=titulo_pagina%></h2><br />
<style type="text/css">
	.menu_paginacaocao{	word-break: break-all; text-align: left; }
	.menu_paginacaocao a{ line-height: 20px; }
</style>

<div class="subMenu" style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC">
	<form action="<%=caminho_pagina%>" name="FORMULARIO_FILTRO" method="get">
		<table cellpadding="4">
			<tr>
				<th rowspan="6">
					<strong>Filtro</strong>
				</th>
			</tr>
			<tr>
				<td><strong>Nome: </strong></td>
				<td>
					<input type="text" class="inputTamanho2" name="filtro_nome" value="<%=filtro_nome%>" maxlength='255'> 
				</td>
				<td rowspan="5"><input type="submit" class="btnFiltrarAmarelo" name="filtro" value=" "></td>
			</tr>

			<tr>
				<td><strong>Vendedor : </strong></td>
				<td>
					<%=padraoFormularioSelect("filtro_cod_vendedor", "", "inputTamanho2", "", "","Selecione" ,filtro_cod_vendedor , "SELECT cod_vendedor, nome_vendedor FROM TB_VENDEDORES WHERE ativado = 1 ORDER BY nome_vendedor ", "alterar", false, false, true )%>
				</td>
				<td colspan="2" align="center">
					
				&nbsp;

				</td>
			</tr>

			<tr>
				<td><strong>Pedido: </strong></td>
				<td>
					<input type="text" class="inputTamanho2" name="filtro_pedido" value="<%=filtro_pedido%>" maxlength='20'>
				</td>
			</tr>
			<tr>
				<td colspan="2" align="center">	&nbsp; </td>
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
	    	<td width="53%"><div align="right">Total <strong><%=rs.RecordCount%></strong> Formandos</div></td>
	  	</tr>
	</table>
	<table width="99%" border="1" align="center" cellpadding="5" style="border-collapse:collapse; text-align: right;">
		<tr bgcolor="whitesmoke">
			<td width="5%"><div align="center" style="font-weight: bold">Pedido</div></td>
			<td width="20%"><div align="center" style="font-weight: bold">Sacado</div></td>
			<td width="15%"><div align="center" style="font-weight: bold">Valor total boletos</div></td>
			<td width="7%"><div align="center" style="font-weight: bold">Valor total Pago</div></td>
			<td width="15%"><div align="center" style="font-weight: bold">Qtd. Parcelas</div></td>
			<td width="4%"><div align="center" style="font-weight: bold"></div></td>
		</tr>
		<%
		Dim x
		For x=1 To rs.PageSize
			If Not rs.Eof Then
			%>
			<tr>
				<td><div align="center"><%=rs("cod_pedido")%></div></td>
			  	<td><div align="left"><%=rs("nome")%></div></td>
			  	<td><div align="right">R$ <%=arrumaNUmero(rs("valor_total_boletos"),"0,00",2)%></div></td>
			  	<td><div align="right">R$ <%=arrumaNUmero(rs("valor_total_pago"),"0,00",2)%></div></td>
			  	<td><div align="center"><%=rs("qtd_parcelas")%></div></td>		  	
			  	<td>
			  		<div align="center">
			  			<table border="0" cellpadding="3">
							<tr>
								<td><a href="cob_pedidos_boletos_formulario.asp?comando=visualizar&cod_pedido=<%=rs("cod_pedido")%>"><img src="images/icons/icon_visualizar.png" alt="Visualizar registro"></a></td>
								<td><a href="cob_pedidos_boletos_formulario.asp?comando=alterar&cod_pedido=<%=rs("cod_pedido")%>"><img src="images/icons/icon_alterar.png" alt="Alterar registro"></a></td>
							</tr>
						</table>
			  		</div>
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
		Nenhum formando encontrado!</div>
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