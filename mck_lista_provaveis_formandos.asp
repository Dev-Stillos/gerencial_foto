<%
Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
BTN_VOLTAR = 1
BTN_IMPRIMIR = 1
BTN_FONTE = 2
BTN_IMPRIMIR_LINK = "#"

Dim titulo_pagina
titulo_pagina = "MCK - Listas de Prováveis Formandos"

Dim calendario_necessario
calendario_necessario = "NÃO"

%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<h2><%=titulo_pagina%></h2>
<br>
<!-- #include file="includes/sub_menu_mck.asp" -->
<br>

<%
	AbrirConexao()
	'/////////////////////////////////////////////////////////////	
	Dim sql, sqlWhere, sqlOrderBy
	sql = "SELECT tmlpf.*, tau.nome, concat(tnt.semestre_conclusao ,'/' ,tnt.ano_conclusao) as descricao FROM TB_MACK_LISTAS_PROVAVEIS_FORMANDOS tmlpf inner join TB_MACK_TEMPORADAS tnt on tmlpf.cod_mack_temporada = tnt.cod_mack_temporada inner join TB_ADM_USUARIOS tau on tau.cod_usuario = tnt.cod_usuario"
	sqlOrderBy = " order by data_cadastro desc"
	'/////////////////////////////////////////////////////////////
	'Filtro
	Dim filtro_cod_mack_temporada
	Dim filtro_somente_ativos, filtro_tia

	filtro_cod_mack_temporada = Cstr(""&Request.QueryString("filtro_cod_mack_temporada"))
	
	If Request.QueryString("filtro") <> "" Then
		If filtro_cod_mack_temporada <> "" Then
			sqlWhere = sqlWhere & " AND tnt.cod_mack_temporada = "&sqlNumero(filtro_cod_mack_temporada)&" "
		End if

		if len(sqlWhere) > 0 Then
			sqlWhere = " WHERE " & mid(sqlWhere, 5, len(sqlWhere)-5)
		end if
	End If
	sql = sql&sqlWhere&sqlOrderBy
	%>

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
	<a class="btn btn-sm btn-success" href="mck_lista_provaveis_formandos_formulario.asp?comando=incluir" style="color: #fff;">Incluir Nova lista</a>
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
					<%=padraoFormularioSelect("filtro_cod_mack_temporada","","inputTamanho2","","","Selecione", filtro_cod_mack_temporada,"SELECT cod_mack_temporada, concat(semestre_conclusao,'/' , ano_conclusao) as descricao FROM TB_MACK_TEMPORADAS order by cod_mack_temporada", "incluir", false, false, true)%>
				</td>
				<td rowspan="5"><input type="submit" class="btnFiltrarAmarelo" name="filtro" value=" "></td>
			</tr>
		</table>
	</form>
</div>
<%
	if Request.QueryString("msg") = "1" Then
	%>
		<div class='msgGrandeSucesso'>
			Arquivo importado com sucesso!
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
	    	<td width="53%"><div align="right">Total <strong><%=rs.RecordCount%></strong> Listas</div></td>
	  	</tr>
	</table>
	<table width="99%" border="1" align="center" cellpadding="5" style="border-collapse:collapse; text-align: right;">
		<tr bgcolor="whitesmoke">
			<td width="5%"><div align="center" style="font-weight: bold"> # </div></td>
			<td width="30%"><div align="center" style="font-weight: bold">Arquivo</div></td>
			<td width="10%"><div align="center" style="font-weight: bold">Temporada</div></td>
			<td width="10%"><div align="center" style="font-weight: bold">Vigente</div></td>
			<td width="10%"><div align="center" style="font-weight: bold">Data Cadastro</div></td>
			<td width="10%"><div align="center" style="font-weight: bold">Usuário</div></td>

		</tr>
		<%
		Dim x
		For x=1 To rs.PageSize
			If Not rs.Eof Then
			%>
			<tr>
				<td><div align="center"><%=rs("cod_mack_lista_provaveis_formandos")%></div></td>
			  	<td><div align="left"><a href='mck_lista_provaveis_formandos_download.asp?comando=download&cod_mack_lista_provaveis_formandos=<%=rs(0)%>' target="_blank"><%=rs("nome_arquivo_original")&"."&rs("extencao")%></a></div></td>
			  	<td><div align="center"><%=rs("descricao")%></div></td>
			  	<td><div align="center">
			  	<%
					if rs("vigente") = "S" Then
						response.write "Sim"
					else
						response.write "NÃO"
					end if
			  	%></div></td>
			  	<td><div align="center"><%=arrumaData(rs("data_cadastro"))%></div></td>
			  	<td><div align="right"><%=rs("nome")%></div></td>
			  
			  	
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
		Nenhuma Lista encontrada, para os filtros aplicados!
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