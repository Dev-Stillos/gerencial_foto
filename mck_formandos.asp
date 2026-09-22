<%
Dim calendario_necessario, BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
calendario_necessario = "SIM"
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 1
BTN_IMPRIMIR_LINK = "#"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<h2>MCK - Formandos</h2>
<br>
<!-- #include file="includes/sub_menu_mck.asp" -->
<%
	AbrirConexao()

	'/////////////////////////////////////////////////////////////	
	Dim sql, sqlWhere, sqlOrderBy
	sql = "SELECT * FROM VW_TB_MACK_FORMANDOS "
	sqlOrderBy = " ORDER BY cod_mack_formando "
	
	'/////////////
	'Filtro
	Dim filtro 
	Dim filtro_cod_mack_colacao, filtro_cod_mack_temporada
	Dim filtro_somente_ativos, filtro_tia, filtro_foto_cadastrada
	
	filtro_cod_mack_colacao = Cstr(""&Request.QueryString("filtro_cod_mack_colacao"))
	filtro_cod_mack_temporada= Cstr(""&Request.QueryString("filtro_cod_mack_temporada"))
	filtro_foto_cadastrada = Cstr(""&Request.QueryString("filtro_foto_cadastrada"))
	filtro_nome = Cstr(""&Request.QueryString("filtro_nome"))
	filtro_tia = Cstr(""&Request.QueryString("filtro_tia"))
	
	if Request.QueryString("filtro") = "" and int(mackenzieTemporadaAtiva(true)) > 0 Then
		filtro_cod_mack_temporada	= int(mackenzieTemporadaAtiva(true))
		filtro = "1"
	else
		filtro = Request.QueryString("filtro")
	end if 


	If filtro <> "" Then		
		If filtro_nome <> "" Then
			sqlWhere = sqlWhere & " AND nome like "&sqlTexto("%"&filtro_nome&"%")&" "
		End if

		if filtro_cod_mack_temporada <> "" Then
			sqlWhere = sqlWhere & " AND cod_mack_temporada = "& filtro_cod_mack_temporada &" "
		end if

		If filtro_tia <> "" Then
			sqlWhere = sqlWhere & " AND tia = "&sqlTexto(filtro_tia)&" "
		End if

		If filtro_cod_mack_colacao <> "" Then
			sqlWhere = sqlWhere & " AND cod_mack_colacao = "&sqlNumero(filtro_cod_mack_colacao)&" "
		End if

		if filtro_foto_cadastrada <> "" then

			if filtro_foto_cadastrada = "S" then
				sqlWhere = sqlWhere & " AND nome_arquivo <> '' and nome_arquivo is not null "	
			elseif	filtro_foto_cadastrada = "N" then
				sqlWhere = sqlWhere & " AND (nome_arquivo = '' or nome_arquivo is null) "	
			end if

		end if

		if len(sqlWhere) > 0 Then
			sqlWhere = " WHERE " & mid(sqlWhere, 5, len(sqlWhere)-5)
		end if
	End If
	sql = sql&sqlWhere&sqlOrderBy

	'response.write (sql)
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
	<a class="btn btn-sm btn-success" href="mck_identificacao.asp?comando=incluir" style="color: #fff;">Incluir Formando</a>
	<a class="btn btn-sm btn-success" href="mck_formandos_formulario.asp?comando=incluir" style="color: #fff;">Incluir Formando sem lista</a>
	<div class="pull-right">
		<!--<a id="btn-baixar-fotos-colacao" class="btn btn-sm btn-info" href="/gerencial_foto/fotos-mck.php?colacao=<%= FILTRO_COD_MACK_COLACAO %>&temporada=<%= filtro_cod_mack_temporada %>" style="color: #fff;">Baixar fotos por colação</a>-->
		<button id="btn-baixar-fotos-colacao" class="btn btn-sm btn-info" style="color: #fff;">Baixar fotos por colação</button>
	</div>
</div>
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
				<td><strong>TIA: </strong></td>
				<td>
					<input type="text" class="inputTamanho2" name="filtro_tia" value="<%=filtro_tia%>" maxlength='20'>
				</td>
			
			</tr>

			<tr>
				<td><strong>Temporada</strong></td>
				<td>
					<%=padraoFormularioSelect("filtro_cod_mack_temporada","id_filtro_cod_mack_temporada","inputTamanho2","","","Selecione", filtro_cod_mack_temporada,"SELECT cod_mack_temporada, concat(semestre_conclusao,'/',ano_conclusao) FROM TB_MACK_TEMPORADAS order by ano_conclusao, semestre_conclusao desc", "incluir", false, false, true)%>
				</td>
				
			</tr>

			<tr>
				<td><strong>colação</strong></td>
				<td id='selectCol'>
					<% if filtro_cod_mack_temporada <> "" then %>
					<%=padraoFormularioSelect("filtro_cod_mack_colacao","","inputTamanho2","","","Selecione", filtro_cod_mack_colacao,"SELECT cod_mack_colacao, descricao FROM TB_MACK_COLACOES where cod_mack_temporada = "&sqlNumero(filtro_cod_mack_temporada)&" order by descricao", "incluir", false, false, true)%>
					<% else %>
						Selecione uma temporada
					<% end if%>
				</td>
				
			</tr>

			<tr>
				<td><strong>Foto Cadastrada: </strong></td>
				<td>
					<%=padraoFormularioSelect("filtro_foto_cadastrada","id_filtro_foto_cadastrada","inputTamanho2","","","Selecione", filtro_foto_cadastrada,"PADRAO:S#Somente com foto cadastrada|N#Somente sem foto cadastrada", "incluir", false, false, true)%>
				</td>
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
			<td width="5%"><div align="center" style="font-weight: bold"> # </div></td>
			<td width="20%"><div align="center" style="font-weight: bold">Nome</div></td>
			<td width="15%"><div align="center" style="font-weight: bold">Curso</div></td>
			<td width="7%"><div align="center" style="font-weight: bold">TIA</div></td>
			<td width="15%"><div align="center" style="font-weight: bold">Colação</div></td>
			<td width="10%"><div align="center" style="font-weight: bold">Retirou Convites</div></td>
			<td width="10%"><div align="center" style="font-weight: bold">Possui foto cadastrada</div></td>
			<td width="4%"><div align="center" style="font-weight: bold"></div></td>
		</tr>
		<%
		Dim x
		For x=1 To rs.PageSize
			If Not rs.Eof Then
			%>
			<tr>
				<td><div align="center"><%=rs("cod_mack_formando")%></div></td>
			  	<td><div align="left"><%=rs("nome")%></div></td>
			  	<td><div align="right"><%=rs("curso")%></div></td>
			  	<td><div align="right"><%=rs("tia")%></div></td>
			  	<td><div align="center"><%=rs("descricao_colacao")%></div></td>
			  	<td><div align="center"><%
			  		if rs("retirou_convites") <> "" and rs("retirou_convites") = "S" then
			  				response.write "Sim"
			  			else
			  				response.write "Não"
			  			end if
				%></div></td>
			  	<td><div align="center"><%
			  			if rs("nome_arquivo") <> "" then
			  				response.write "Sim"
			  			else
			  				response.write "Não"
			  			end if
			  			%>
			  	</div></td>
			  	
			  	<td>
			  		<div align="center">
			  			<table border="0" cellpadding="3">
							<tr>
								<td><a href="mck_formandos_formulario.asp?comando=visualizar&cod_mack_formando=<%=rs("cod_mack_formando")%>"><img src="images/icons/icon_visualizar.png" alt="Visualizar registro"></a></td>
								<td><a href="mck_formandos_formulario.asp?comando=alterar&cod_mack_formando=<%=rs("cod_mack_formando")%>"><img src="images/icons/icon_alterar.png" alt="Alterar registro"></a></td>
								<td><a href="mck_formandos_formulario.asp?comando=alterar&cod_mack_formando=<%=rs("cod_mack_formando")%>&guia=2"><img width="22px" src="images/icons/icon_camera.png" alt="Alterar registro"></a></td>
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
<script src='js/FileSaver.min.js'></script>
<script type="text/javascript">
	$(function(){
		$('#id_filtro_cod_mack_temporada').change(function(){
			var request = $.ajax({
			  url: "mck_formandos_formulario_ajax.asp",
			  method: "POST",
			  data: { t : "selectColacoes" , "p0" : $('#id_filtro_cod_mack_temporada').val(), p1 : 0  },
			  dataType: "html",
			  beforeSend: function(){
			  	$('#id_filtro_cod_mack_temporada').attr('disabled', 'disabled');
			  	$( "#selectCol" ).html( "Aguarde..." );
			  }
			});
			 
			request.done(function( msg ) {
			  $('#id_filtro_cod_mack_temporada').removeAttr('disabled');
			  $( "#selectCol" ).html( msg );

			});
			 
			request.fail(function( jqXHR, textStatus ) {
			  console.log( "Request failed: " + textStatus );
			  console.log(  jqXHR)
			});
		});

		$("#btn-baixar-fotos-colacao").click(function(event){
			event.preventDefault();
			event.stopPropagation();
			
			var target = window.location.origin + "/gerencial_foto/fotos-mck.php?colacao=<%= FILTRO_COD_MACK_COLACAO %>&temporada=<%= filtro_cod_mack_temporada %>";
			var btn = $(this);
			
			btn.css("opacity",0.5);
			btn.css("cursor","default");
			btn.removeClass("btn-info").addClass("btn-warning").addClass("disabled");
			btn.html("Aguarde");
			
			window.open(target);
			
			btn.css("opacity",1);
			btn.css("cursor","pointer");
			btn.removeClass("btn-warning").removeClass("disabled").addClass("btn-info");
			btn.html("Baixar fotos por colação");
			console.log('Fim da requisição');
			/*$.ajax({
				url:target, 
				type: 'GET',
				contentType: 'application/zip',
				success: function(data){
					var blob = new Blob([str2bytes(data)], {type: "octet/stream"});
					saveAs(blob, "data1.zip"); 
					alert('sucesso 2');
				},
				error: function(data){
					alert('erro');
				},
				beforeSend: function(){
					btn.css("opacity",0.5);
					btn.css("cursor","default");
					btn.removeClass("btn-info").addClass("btn-warning").addClass("disabled");
					btn.html("Aguarde");
				},
				complete: function(){
					btn.css("opacity",1);
					btn.css("cursor","pointer");
					btn.removeClass("btn-warning").removeClass("disabled").addClass("btn-info");
					btn.html("Baixar fotos por colação");
					console.log('Fim da requisição');
				}
			});*/


			return false;
		});

	});
</script>

</body>
</html>