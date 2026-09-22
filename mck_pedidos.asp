<%
Dim calendario_necessario, BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
calendario_necessario = "SIM"
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 1
BTN_IMPRIMIR_LINK = "#"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<h2>MCK - Pedidos</h2>
<br>
<!-- #include file="includes/sub_menu_mck.asp" -->

<br>

<div class="container" style="margin: 5px 0 20px 0;">
	<a class="btn btn-sm btn-success" href="mck_pedidos_formulario.asp" style="color: #fff;">Novo Pedido</a>
</div>

<%
	AbrirConexao()

	'/////////////////////////////////////////////////////////////	
	Dim sql, sqlWhere, sqlOrderBy, filtro
	sql = "SELECT pd.* ,  tmp.semestre_conclusao, tmp.ano_conclusao  FROM VW_TB_MACK_PEDIDOS pd  left join TB_MACK_TEMPORADAS tmp on pd.cod_mack_temporada = tmp.cod_mack_temporada "
	sqlOrderBy = " order by nome asc"
	
	'/////////////
	'Filtro
	Dim filtro_nome
	Dim filtro_tia
	
	filtro_nome = Cstr(""&Request.QueryString("filtro_nome"))
	filtro_tia = Cstr(""&Request.QueryString("filtro_tia"))
	filtro_cod_mack_temporada = Cstr(""&Request.QueryString("filtro_cod_mack_temporada"))


	if Request.QueryString("filtro") = "" and int(mackenzieTemporadaAtiva(true)) > 0 Then
		filtro_cod_mack_temporada	= int(mackenzieTemporadaAtiva(true))
		filtro = "1"
	else
		filtro = Request.QueryString("filtro")
	end if 

	If filtro <> "" Then
		
		if filtro_cod_mack_temporada <> "" Then
			sqlWhere = sqlWhere & " AND pd.cod_mack_temporada = "& filtro_cod_mack_temporada &" "
		end if


		If filtro_nome <> "" Then
			sqlWhere = sqlWhere & " AND nome like "&sqlTexto("%"&filtro_nome&"%")&" "
		End if

		If filtro_tia <> "" Then
			sqlWhere = sqlWhere & " AND tia like "&sqlTexto("%"&filtro_tia&"%")&" "
		End if

		if len(sqlWhere) > 0 Then
			sqlWhere = " WHERE " & mid(sqlWhere, 5, len(sqlWhere)-5)
		end if
	End If
	sql = sql&sqlWhere&sqlOrderBy

	'response.write sql & "aaaaaaaaaaaaaas"
	%>


<div class="subMenu" style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC">
	<form action="<%=caminho_pagina%>" name="FORMULARIO_FILTRO" method="get">
		<table cellpadding="4">
			<tr>
				<th rowspan="4">
					<strong>Filtro</strong>
				</th>
			</tr>

		<tr>
				<td><strong>Temporada</strong></td>
				<td>
					<%=padraoFormularioSelect("filtro_cod_mack_temporada","id_filtro_cod_mack_temporada","inputTamanho2","","","Selecione", filtro_cod_mack_temporada,"SELECT cod_mack_temporada, concat(semestre_conclusao,'/',ano_conclusao) FROM TB_MACK_TEMPORADAS order by ano_conclusao, semestre_conclusao desc", "incluir", false, false, true)%>
				</td>
				
			</tr>



			<tr>
				<td><strong>Nome: </strong></td>
				<td>
					<input type="text" class="inputTamanho2" name="filtro_nome">
				</td>
				<td rowspan="3"><input type="submit" class="btnFiltrarAmarelo" name="filtro" value=" "></td>
			</tr>

			<tr>
				<td><strong>TIA</strong></td>
				<td>
					<input type="text" class="inputTamanho2" name="filtro_tia">
				</td>
			</tr>

			<tr>
				<td colspan="2" align="center">
					
				&nbsp;

				</td>
			</tr>

			<tr>
				<th rowspan="3">
					<strong>Ordernar Por?:</strong>
				</th>
			</tr>

			<tr>
				<th rowspan="3" colspan="2">
					<%=padraoFormularioSelect("ordenacao", "", "inputTamanho2", "", "", "Selecione", cod_mack_forma_pagamento, "PADRAO:N#Nome|D#data Cadastro", "incluir", false, false, false)%>
				</th>
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
			<td width="7%"><div align="center" style="font-weight: bold"> # </div></td>
			<td width="20%"><div align="center" style="font-weight: bold">Nome</div></td>
			<td width="20%"><div align="center" style="font-weight: bold">TIA</div></td>
			<td width="10%"><div align="center" style="font-weight: bold">Temporada</div></td>
			<td width="7%"><div align="center" style="font-weight: bold">Qtd. Produtos</div></td>
			<td width="10%"><div align="center" style="font-weight: bold">Valor total</div></td>
			<td width="10%"><div align="center" style="font-weight: bold">Data Cadastro</div></td>
			<td width="10%"><div align="center" style="font-weight: bold">Usuário</div></td>
			<td width="8%"><div align="center" style="font-weight: bold"></div></td>
		</tr>
		<%
		Dim x
		For x=1 To rs.PageSize
			If Not rs.Eof Then
			%>
			<tr>
				<td><div align="center"><%=rs("cod_mack_pedido")%></div></td>
			  	<td><div align="left"><%=rs("nome")%></div></td>
			  	<td><div align="left"><%=rs("tia")%></div></td>
			  	<td><div align="center"><%=rs("semestre_conclusao")&"/"&rs("ano_conclusao")%></div></td>
			  	<td><div align="right">
			  		<%
			  		Set rs1 = objConexao.Execute("SELECT count(*) FROM TB_MACK_PEDIDOS_PRODUTOS WHERE cod_mack_pedido = "&sqlNumero(rs("cod_mack_pedido")))
			  		if not rs1.Eof Then
			  			response.write rs1(0)
			  		end if
			  		rs1.Close
			  		set rs1 = nothing
			  		%>
			  	</div></td>
			  	<td><div align="center"><%
			  		Set rs1 = objConexao.Execute("SELECT SUM(valor_unitario *quantidade) FROM TB_MACK_PEDIDOS_PRODUTOS WHERE cod_mack_pedido = "&sqlNumero(rs("cod_mack_pedido")))
			  		if not rs1.Eof Then
			  			response.write arrumaNumero(rs1(0),"",2)
			  		end if
			  		rs1.Close
			  		set rs1 = nothing
			  		%>
			  	</div></td>

				<td style="text-align:center;">
					<%=rs("data_cadastro")%>
				</td>
				<td>
					<div  style="text-align:center;"><%=nomeUsuario(rs("cod_usuario"),true)%></div>
					
				</td>
			  	<td >
			  		<div align="center">
			  			<table border="0" cellpadding="3">
							<tr>
								<td><a href="mck_pedidos_formulario.asp?comando=visualizar&cod_mack_pedido=<%=rs("cod_mack_pedido")%>"><img src="images/icons/icon_visualizar.png" alt="Visualizar registro"></a></td>
								<!--<td><a href="mck_pedidos_formulario.asp?comando=alterar&cod_mack_pedido=<%=rs("cod_mack_pedido")%>"><img src="images/icons/icon_alterar.png" alt="Alterar registro"></a></td>-->
								<td><a href="mck_pedidos_impressao.asp?c=<%=rs("cod_mack_pedido")%>"><img src="images/icons/icon_imprimir.png" alt="Excluir registro"></a></td>
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