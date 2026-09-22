<%
Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
BTN_VOLTAR = 1
BTN_IMPRIMIR = 1
BTN_FONTE = 2
BTN_IMPRIMIR_LINK = "#"

Dim titulo_pagina
titulo_pagina = "MCK - Identificação de Formandos"

Dim calendario_necessario
calendario_necessario = "NÃO"

%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<h2><%=titulo_pagina%></h2>
<br>
<!-- #include file="includes/sub_menu_mck.asp" -->
<br>
<br>
<%
AbrirConexao()

	Dim nome_arquivo_original, data_cadastro, nome_usuario_lista, cod_mack_lista_provaveis_formandos

	'buscamos a lista vigente de acordo com a temporada
	set rs = objConexao.execute("SELECT tmlpf.*, tau.nome, concat(semestre_conclusao,'/',ano_conclusao) as descricao FROM TB_MACK_LISTAS_PROVAVEIS_FORMANDOS tmlpf inner join TB_MACK_TEMPORADAS tnt on tmlpf.cod_mack_temporada = tnt.cod_mack_temporada inner join TB_ADM_USUARIOS tau on tau.cod_usuario = tnt.cod_usuario where tmlpf.vigente = 'S' ORDER BY data_cadastro LIMIT 1  ")
	if not rs.Eof Then
		nome_arquivo_original = rs("nome_arquivo_original")
		data_cadastro = rs("data_cadastro")
		nome_usuario_lista = rs("nome")
		cod_mack_lista_provaveis_formandos  = rs("cod_mack_lista_provaveis_formandos")
	end if 
	rs.Close
	set rs = Nothing

	'/////////////////////////////////////////////////////////////	
	Dim sql, sqlWhere, sqlOrderBy
	sql = "SELECT * FROM TB_MACK_LISTAS_PROVAVEIS_FORMANDOS_NOMES "
	sqlOrderBy = " "
	sqlWhere = " WHERE  cod_mack_lista_provaveis_formandos = "&sqlNumero(cod_mack_lista_provaveis_formandos)&" "

	'/////////////////////////////////////////////////////////////
	'Filtro
	Dim filtro_nome, filtro_tia, filtro_departamento, filtro_curso

	filtro_departamento = Cstr(""&Request.QueryString("filtro_departamento"))
	filtro_curso = Cstr(""&Request.QueryString("filtro_curso"))
	filtro_nome = Cstr(""&Request.QueryString("filtro_nome"))
	filtro_tia = Cstr(""&Request.QueryString("filtro_tia"))

	If Request.QueryString("filtro") <> "" Then
		If filtro_departamento <> "" Then
			sqlWhere = sqlWhere & " AND departamento_sigla = "&sqltexto(filtro_departamento)&" "
		End if

		If filtro_curso <> "" Then
			sqlWhere = sqlWhere & " AND curso = "&sqltexto(filtro_curso)&" "
		End if

		If filtro_nome <> "" Then
			sqlWhere = sqlWhere & " AND nome like "&sqltexto("%"&filtro_nome&"%")&" "
		End if

		If filtro_tia <> "" Then
			sqlWhere = sqlWhere & " AND matricula = "&sqltexto(filtro_tia)&" "
		End if

		if len(sqlWhere) > 0 Then
			'sqlWhere = " WHERE " & mid(sqlWhere, 5, len(sqlWhere)-5)
		end if
	End If
	sql = sql&sqlWhere&sqlOrderBy
''	response.write sql
	%>
<br />
<style type="text/css">
	.menu_paginacaocao{
		word-break: break-all;
		text-align: left;
	}
	.menu_paginacaocao a{
		line-height: 20px;
	}
</style>

<div class="subMenu" style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC;  width:700px; margin:0 auto 0 auto;">
	<form action="<%=caminho_pagina%>" name="FORMULARIO_FILTRO" method="get">
		<table cellpadding="4">
			<tr>
				<th rowspan="5">
					<strong>Filtro</strong>
				</th>
			</tr>
			<tr>
				<td align="right"><strong>Departamento: </strong></td>
				<td>
					<%=padraoFormularioSelect("filtro_departamento","","inputTamanho3","","","Selecione", filtro_departamento,"SELECT departamento_sigla, concat(departamento_sigla,' - ',departamento) as nome_departamento FROM TB_MACK_LISTAS_PROVAVEIS_FORMANDOS_NOMES GROUP BY departamento_sigla,departamento order by departamento", "incluir", false, false, true)%>
				</td>
			</tr>
			<tr>
				<td align="right"><strong>Curso: </strong></td>
				<td>
					<%=padraoFormularioSelect("filtro_curso","","inputTamanho3","","","Selecione", filtro_curso,"SELECT curso, curso as nome_curso FROM TB_MACK_LISTAS_PROVAVEIS_FORMANDOS_NOMES GROUP BY curso order by curso", "incluir", false, false, true)%>
				</td>
			</tr>
			<tr align="right">
				<td><strong>Nome: </strong></td>
				<td>
					<input type="text" class="inputTamanho3" name='filtro_nome' value="<%=filtro_nome%>" maxlength="150">
				</td>
			</tr>
			<tr>
				<td align="right"><strong>Tia: </strong></td>
				<td>
					<input type="text" class="inputTamanho2" name='filtro_tia' value="<%=filtro_tia%>" maxlength="20">
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" class="btnFiltrarAmarelo" name="filtro" value=" ">
				</td>
			</tr>
				
		</table>
	</form>
	<div class="container" style="margin: 20px 0px 5px 250px">
		<a class="btn btn-sm btn-success" href="mck_formandos_formulario.asp?comando=incluir" style="color: #fff;">Incluir Formando sem lista</a>
	</div>	
</div>	
<% if Request.QueryString("msg") = "1" Then %>
	<div class='msgGrandeSucesso'>
		Formando identificado com sucesso!
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
	<table width="700px" border="0" align="center">
		<tr>
	    	<td width="47%">Página <strong><%=rs.AbsolutePage%></strong> de <%=rs.PageCount%></td>
	    	<td width="53%"><div align="right">Total <strong><%=rs.RecordCount%></strong> Formandos</div></td>
	  	</tr>
	</table>
	<table width="700px" border="1" align="center" cellpadding="5" style="border-collapse:collapse; text-align: right;">
		<tr bgcolor="whitesmoke">
			<td width="5%"><div align="center" style="font-weight: bold"> # </div></td>
			<td width="30%"><div align="center" style="font-weight: bold">Nome</div></td>
			<td width="10%"><div align="center" style="font-weight: bold">Matricula</div></td>
			<td width="10%"><div align="center" style="font-weight: bold"></div></td>
		</tr>
		<%
		Dim x
		For x=1 To rs.PageSize
			If Not rs.Eof Then
			%>
			<tr>
				<td><div align="center"><%=rs("departamento_sigla")%></div></td>
			  	<td><div align="left"><%=rs("nome")%></div></td>
			  	<td><div align="center"><%=rs("matricula")%></div></td>
			  	<td>
			  		<div align="center">
			  			<table border="0" cellpadding="3">
							<tr>
								<td><a href="mck_formandos_formulario.asp?comando=incluir&cod_mack_lista_provaveis_formandos_nome=<%=rs("cod_mack_lista_provaveis_formandos_nome")%>"><img src="images/icons/icon_visualizar.png" alt="Visualizar registro"></a></td>
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
	<br />
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