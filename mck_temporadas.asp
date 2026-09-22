<%
Dim calendario_necessario, BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
calendario_necessario = "SIM"
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 1
BTN_IMPRIMIR_LINK = "#"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<h2>MCK - Temporadas</h2>
<br>
<!-- #include file="includes/sub_menu_mck.asp" -->
<%
	AbrirConexao()

	'/////////////////////////////////////////////////////////////	
	Dim sql, sqlWhere, sqlOrderBy
	sql = "SELECT * FROM TB_MACK_TEMPORADAS "
	sqlOrderBy = " "
	'/////////////
	'Filtro
	Dim filtro_cod_mack_colacao
	Dim filtro_somente_ativos, filtro_tia
	
	filtro_cod_mack_colacao = Cstr(""&Request.QueryString("filtro_cod_mack_colacao"))
	filtro_nome = Cstr(""&Request.QueryString("filtro_nome"))
	filtro_tia = Cstr(""&Request.QueryString("filtro_tia"))
	
	If Request.QueryString("filtro") <> "" Then
	
		If filtro_nome <> "" Then
			sqlWhere = sqlWhere & " AND nome like "&sqlTexto("%"&filtro_nome&"%")&" "
		End if

		If filtro_tia <> "" Then
			sqlWhere = sqlWhere & " AND tia = "&sqlTexto(filtro_tia)&" "
		End if


		If filtro_cod_mack_colacao <> "" Then
			sqlWhere = sqlWhere & " AND cod_mack_colacao = "&sqlNumero(filtro_cod_mack_colacao)&" "
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
	<a class="btn btn-sm btn-success" href="mck_temporadas_formulario.asp?comando=incluir" style="color: #fff;">Incluir temporada</a>
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
	    	<td width="53%"><div align="right">Total <strong><%=rs.RecordCount%></strong> Temporadas</div></td>
	  	</tr>
	</table>
	<table width="99%" border="1" align="center" cellpadding="5" style="border-collapse:collapse; text-align: right;">
		<tr bgcolor="whitesmoke">
			<td width="7%"><div align="center" style="font-weight: bold"> # </div></td>
			<td width="20%"><div align="center" style="font-weight: bold">Sequência</div></td>
			<td width="20%"><div align="center" style="font-weight: bold">Temporada</div></td>
			<td width="10%"><div align="center" style="font-weight: bold">Data início</div></td>
			<td width="10%"><div align="center" style="font-weight: bold">Data Fim</div></td>
			<td width="10%"><div align="center" style="font-weight: bold">Modificado</div></td>
			<td width="10%"><div align="center" style="font-weight: bold">Usuário</div></td>
			<td width="8%"><div align="center" style="font-weight: bold"></div></td>
		</tr>
		<%
		Dim x
		For x=1 To rs.PageSize
			If Not rs.Eof Then
			%>
			<tr>
				<td><div align="center"><%=rs("cod_mack_temporada")%></div></td>
				<td><div align="center"><%=rs("sequencia")%></div></td>

				<td><div align="left"><%=rs("semestre_conclusao")&"/"&rs("ano_conclusao")%></div></td>
				<td><div align="center"><%=rs("inicio")%></div></td>
				<td><div align="center"><%=rs("fim")%></div></td>

			  	<td><div align="center"><%=rs("modificado")%></div></td>
				<td><div align="center"><%=nomeUsuario(rs("cod_usuario"), true)%></div></td>
			  	
			  	<td >
			  		<div align="center">
			  			<table border="0" cellpadding="3">
							<tr>
								<td><a href="mck_temporadas_formulario.asp?comando=visualizar&cod_mack_temporada=<%=rs("cod_mack_temporada")%>"><img src="images/icons/icon_visualizar.png" alt="Visualizar registro"></a></td>
								
								<td><a href="mck_temporadas_formulario.asp?comando=alterar&cod_mack_temporada=<%=rs("cod_mack_temporada")%>"><img src="images/icons/icon_alterar.png" alt="Alterar registro"></a></td>

								<td><a href="mck_temporadas_formulario.asp?comando=excluir&cod_mack_temporada=<%=rs("cod_mack_temporada")%>"><img src="images/icons/icon_excluir.png" alt="Alterar registro"></a></td>
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
		Nenhuma temporada encontrada!</div>
	<%End if

	rs.Close
	Set rs=Nothing

	FecharConexao()
%>
<br />
<br />
<br />
<!-- #include file="files/includes/miolo-inferior.asp" -->
<script type="text/javascript">
	$(function(){
		$('.showTB').click(function(){
			var d = $(this).attr('cod-produto');
			if ($(".produto_"+d).css("display") == 'none')
				$(".produto_"+d).css("display","table-row");
			else
				$(".produto_"+d).css("display","none");
		});
	});
</script>

</body>
</html>