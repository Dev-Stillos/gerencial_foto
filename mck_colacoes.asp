<%
Dim calendario_necessario, BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
calendario_necessario = "SIM"
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 1
BTN_IMPRIMIR_LINK = "#"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<h2>MCK - Colações</h2>
<br>
<!-- #include file="includes/sub_menu_mck.asp" -->
<%
	AbrirConexao()

	'/////////////////////////////////////////////////////////////	
	Dim sql, sqlWhere, sqlOrderBy, filtro
	sql = "SELECT col.* , temp.semestre_conclusao,  temp.ano_conclusao FROM TB_MACK_COLACOES as col inner join TB_MACK_TEMPORADAS as temp on temp.cod_mack_temporada = col.cod_mack_temporada"
	sqlOrderBy = " "
	'/////////////
	'Filtro
	Dim filtro_cod_mack_temporada

	

	if Request.QueryString("filtro") = "" and int(mackenzieTemporadaAtiva(true)) > 0 and  Request.QueryString("codmsg") = "" then
		filtro_cod_mack_temporada = int(mackenzieTemporadaAtiva(true))
		filtro = "1"
	else
		filtro_cod_mack_temporada = Cstr(""&Request.QueryString("filtro_cod_mack_temporada"))
		filtro = Request.QueryString("filtro")
	end if


	If filtro <> "" Then
		If filtro_cod_mack_temporada <> "" Then
			sqlWhere = sqlWhere & " AND col.cod_mack_temporada = "&sqlNumero(filtro_cod_mack_temporada)&" "
		End if

		if len(sqlWhere) > 0 Then
			sqlWhere = " WHERE " & mid(sqlWhere, 5, len(sqlWhere)-5)
		end if
	End If
	sql = sql&sqlWhere&sqlOrderBy
	%>
<h2><%=titulo_pagina%></h2><br />
<style type="text/css">
	.menu_paginacaocao{
		word-break: break-all;
		text-align: left;
	}
	.menu_paginacaocao a{
		line-height: 20px;
	}
</style>

<div class="container" style="margin: 5px 0 20px 0;">
	<a class="btn btn-sm btn-success" href="mck_colacoes_formulario.asp?comando=incluir" style="color: #fff;">Incluir colação</a>
</div>


<div class="subMenu" style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC">
	<form action="<%=caminho_pagina%>" name="FORMULARIO_FILTRO" method="get">
		<table cellpadding="4">
			<tr>
				<th rowspan="5">
					<strong>Filtro</strong>
				</th>
			</tr>
			<tr>
				<td><strong>Temporada: </strong></td>
				<td>
					<%=padraoFormularioSelect("filtro_cod_mack_temporada","","inputTamanho2","","","Selecione", filtro_cod_mack_temporada,"SELECT cod_mack_temporada, CONCAT(semestre_conclusao,'/',ano_conclusao)  FROM TB_MACK_TEMPORADAS order by sequencia", "incluir", false, false, true)%>
				</td>
				<td rowspan="5"><input type="submit" class="btnFiltrarAmarelo" name="filtro" value=" "></td>
			</tr>
		</table>
	</form>
</div>



<%
codmsg = Request.QueryString("codmsg")
if codmsg = "3" Then
	%>
	<div class="msgGrandeSucesso">
		Registro excluido com sucesso
	</div>
	<%
elseif codmsg = "4" Then
	%>
	<div class="msgGrandeErro">
		Registro não excluido pois existem outros registros relacionados
	</div>
	<%
end if 

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
	    	<td width="53%"><div align="right">Total <strong><%=rs.RecordCount%></strong> Colações</div></td>
	  	</tr>
	</table>
	<table width="99%" border="1" align="center" cellpadding="5" style="border-collapse:collapse; text-align: right;">
		<tr bgcolor="whitesmoke">
			<td width="7%"><div align="center" style="font-weight: bold"> # </div></td>
			<td width="7%"><div align="center" style="font-weight: bold"> seq. </div></td>
			<td width="20%"><div align="center" style="font-weight: bold">Descrição</div></td>
			<td width="20%"><div align="center" style="font-weight: bold">Temporada</div></td>
			<td width="10%"><div align="center" style="font-weight: bold">Contrato Interno</div></td>
			<td width="10%"><div align="center" style="font-weight: bold">Departamento</div></td>
			<td width="10%"><div align="center" style="font-weight: bold">Data</div></td>
			<td width="10%"><div align="center" style="font-weight: bold">Local</div></td>
			<td width="8%"><div align="center" style="font-weight: bold"></div></td>
		</tr>
		<%
		Dim x
		For x=1 To rs.PageSize
			If Not rs.Eof Then
			%>
			<tr>
				<td><div align="center"><%=rs("cod_mack_colacao")%></div></td>
				<td><div align="center"><%=rs("sequencia")%></div></td>
				<td><div align="center"><%=rs("descricao")%></div></td>
				<td><div align="left"><%=rs("semestre_conclusao")&"/"&rs("ano_conclusao")%></div></td>
				<td><div align="center"><%=rs("cod_contrato")%></div></td>
				<td><div align="center"><%=rs("unidade")%></div></td>
				<td><div align="center"><%=arrumadata(rs("data_colacao"))%></div></td>
				<td><div align="center">
				<%
					if rs("Local") = "H" Then
						response.write "Higienópolis"
					elseif rs("Local") = "C" Then
						response.write "Campinas"
					elseif rs("Local") = "A" Then
						response.write "Alphaville"
					end if
				%>
				</div></td>

			  	<td >
			  		<div align="center">
			  			<table border="0" cellpadding="3">
							<tr>
								<td><a href="mck_colacoes_formulario.asp?comando=visualizar&cod_mack_colacao=<%=rs("cod_mack_colacao")%>"><img src="images/icons/icon_visualizar.png" alt="Visualizar registro"></a></td>
								
								<td><a href="mck_colacoes_formulario.asp?comando=alterar&cod_mack_colacao=<%=rs("cod_mack_colacao")%>"><img src="images/icons/icon_alterar.png" alt="Alterar registro"></a></td>

								<td><a href="mck_colacoes_formulario.asp?comando=excluir&cod_mack_colacao=<%=rs("cod_mack_colacao")%>"><img src="images/icons/icon_excluir.png" alt="Alterar registro"></a></td>
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
		Nenhuma colação encontrada!
	</div>
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