<%
Dim calendario_necessario, BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
calendario_necessario = "SIM"
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 1
BTN_IMPRIMIR_LINK = "#"
%>
<!-- #include file="files/includes/miolo-superior.asp" <--></-->
<%
	AbrirConexao()

	'/////////////////////////////////////////////////////////////	
	Dim sql, sqlWhere, sqlOrderBy
	sql = "SELECT  TB_ENTIDADES.* , COUNT(TB_FICHAS_IDENTIFICACAO.cod_formando) as qtd FROM stillos3.TB_ENTIDADES join TB_FICHAS_IDENTIFICACAO ON TB_ENTIDADES.COD_CONTRATO = TB_FICHAS_IDENTIFICACAO.COD_CONTRATO where TB_ENTIDADES.ano <= 2007 "

	sqlOrderBy = " group by TB_ENTIDADES.cod_contrato, TB_ENTIDADES.descricao, TB_ENTIDADES.SEMESTRE, TB_ENTIDADES.ano, TB_ENTIDADES.SOMENTE_FOTO order by TB_ENTIDADES.ano desc, TB_ENTIDADES.COD_CONTRATO desc"

	
	'/////////////
	'Filtro
	dim filtro_contrato

	filtro_contrato = Cstr(""&Request.QueryString("filtro_contrato"))
	
	
	If Request.QueryString("filtro") <> "" Then
	
		If filtro_contrato <> "" Then
			sqlWhere = sqlWhere & " AND TB_FICHAS_IDENTIFICACAO.COD_CONTRATO like "&sqlTexto("%"&filtro_contrato&"%")&" "
		End if

		
	End If
sql = sql & sqlWhere & sqlOrderBy
	%>
<h2>Produções anteriores a 2007</h2><br />
<style type="text/css">
	
	.menu_paginacaocao{
		word-break: break-all;
		text-align: left;
	}
	.menu_paginacaocao a{
		line-height: 20px;
	}
</style>
<div class="subMenu" style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC">
	<form action="<%=caminho_pagina%>" name="FORMULARIO_FILTRO" method="get">
		<table cellpadding="4">
			<tr>
				<th rowspan="5">
					<strong>Filtro</strong>
				</th>
			</tr>

			<tr>
				<td><strong>Contrato: </strong></td>
				<td>
					<input type="text" class="inputTamanho2" name="filtro_contrato" value="<%=filtro_nome%>"> 
				</td>
				<td rowspan="5"><input type="submit" class="btnFiltrarAmarelo" name="filtro" value=" "></td>
			</tr>
			

			<tr>
				<td colspan="2" align="center">
					
				&nbsp;

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
			<td width="7%"><div align="center" style="font-weight: bold"> # </div></td>
			<td width="20%"><div align="center" style="font-weight: bold">Entidade</div></td>
			<td width="20%"><div align="center" style="font-weight: bold">Semestre / Ano</div></td>
			<td width="10%"><div align="center" style="font-weight: bold">Quantidade Formandos</div></td>
		</tr>
		<%
		Dim x
		For x=1 To rs.PageSize
			If Not rs.Eof Then
			%>
			<tr>
				<td><div align="center"><%=rs("cod_contrato")%></div></td>
			  	<td><div align="left"><a href='adm_busca_completa_depto_foto.asp?TIPO_BUSCA=5&BUSCAR_TERMO=<%=rs("cod_contrato")%>&Submit=+'><%=rs("descricao")%></a></div></td>
			  	<td><div align="right"><%=rs("SEMESTRE")&"/"&cstr(rs("ano"))%></div></td>
			  	<td><div align="right"><%=rs("qtd")%></div></td>
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