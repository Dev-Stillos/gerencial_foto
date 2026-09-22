<% dim script_valida_form_master : script_valida_form_master = "NAO" %>
<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<!-- #include file="files/includes/functions_local.asp" -->
<%

sub dislayErro(cod_erro)

	dim erro(4)
	erro(0) = "Você não possui este acesso."
	erro(1) = "Um erro inesperado ocorreu tente novamente mais tarde, se o erro persistir informe o departamento de TI"
	erro(2) = "Esta produção não possui contratos, ou todos os contratos j� est�o relacionados a uma produção"
	erro(3) = ""

	response.write("<div class='alert alert-danger'>")
	response.write(erro(cod_erro))
	response.write("</div>")
	response.end

end sub

if not validaUsuario(2,,82, false) then
	response.write("Você não tem acesso")
	response.end
end if

Dim rs, cod_producao, cod_producao_desconto, campo1, comando
Dim aliquota_padrao, descricao, aplicar_todos_contratos, contratos, contratos_desconto

campo1 = request.form("campo1")
cod_producao = request.queryString("cod_producao")
cod_producao_desconto = request.queryString("cod_producao_desconto")

if cod_producao = "" then'
	cod_producao = request.form("cod_producao")
end if

if cod_producao_desconto = "" then
	cod_producao_desconto = request.form("cod_producao_desconto")
end if


comando = request.queryString("comando")
if comando = "" then
	comando = request.form("comando")
	if comando = "" then
		comando = "incluir"
	end if
end if

abrirConexao()
contratos = ""

if comando = "alterar" and cod_producao_desconto <> "" then
	sqlWhere = " AND cod_producao_desconto <> "&sqlNumero(cod_producao_desconto)
end if 

'buscamos os contratos disponiveis para desocnto pois se não houver contratos aqui não prosseguimos

Set rs = objConexao.Execute("SELECT DISTINCT(COD_CONTRATO) AS COD_CONTRATO FROM TB_FICHAS_IDENTIFICACAO WHERE cod_contrato not in (SELECT cod_contrato FROM VW_TB_PRODUCAO_DESCONTOS_CONTRATOS where cod_producao = "&sqlNumero(cod_producao)&" "&sqlWhere&") AND COD_PRODUCAO="&sqlNumero(cod_producao)&" ORDER BY COD_CONTRATO ASC")
if not rs.Eof then
	while not rs.eof

		contratos = contratos&rs(0)&","

	rs.movenext
	wend
else

end if
rs.close
set rs = nothing

fecharConexao()

if contratos <> ""  then
	contratos = left(contratos,len(contratos)-1)
	contratos = split(contratos,",")
else
	if comando = "incluir" then
		cod_erro = 2
	end if
end if

sub lerRegistros()
	dim rs

	abrirConexao()

	set rs = objConexao.Execute("SELECT aliquota_padrao, descricao, aplicar_todos_contratos FROM TB_PRODUCAO_DESCONTOS WHERE cod_producao_desconto = "&sqlNumero(cod_producao_desconto))
	if not rs.Eof then
		aliquota_padrao = rs("aliquota_padrao")
		descricao = rs("descricao")
		aplicar_todos_contratos = rs("aplicar_todos_contratos")
	end if 
	rs.close
	set rs = nothing

	set rs = objConexao.Execute("SELECT cod_contrato FROM TB_PRODUCAO_DESCONTOS_CONTRATOS WHERE cod_producao_desconto = "&sqlNumero(cod_producao_desconto) )
	while not rs.eof 
		contratos_desconto = contratos_desconto & rs(0) & ","
	rs.movenext
	Wend
	rs.close
	set rs = nothing

	if contratos_desconto <> "" then
		contratos_desconto = left(contratos_desconto, len(contratos_desconto)-1)
	end if

	fecharConexao()
end sub

sub lerFormulario()

	cod_producao  = request.form("cod_producao")
	cod_producao_desconto  = request.form("cod_producao_desconto")
	aliquota_padrao = request.form("aliquota_padrao")
	descricao = request.form("descricao")
	aplicar_todos_contratos = request.form("aplicar_todos_contratos")
	contratos_desconto = request.form("contratos_desconto")

end sub

sub incluirRegistro(redirecionar)

	call lerFormulario()

	dim i, rs, aux, contratos_incluir

	abrirConexao()

		objConexao.execute("INSERT INTO TB_PRODUCAO_DESCONTOS (cod_producao, aliquota_padrao, descricao, aplicar_todos_contratos) VALUES ("&sqlNumero(cod_producao)&","&sqlNumero(aliquota_padrao)&", "&sqlTexto(descricao)&", "&sqlTexto(aplicar_todos_contratos)&")")

		'BUSCAMOS O CODIGO GERADO
		Set rs = objConexao.Execute("SELECT max(cod_producao_desconto) FROM TB_PRODUCAO_DESCONTOS")
		if not rs.Eof then
			aux = rs(0)
		end if
		rs.close
		set rs = nothing

		'INSERIMOS OS CONTRATOS NA TABELA DE RELACIONAMENTO
		if aplicar_todos_contratos = "S" then
			contratos_incluir = contratos
		else
			contratos_incluir = split(contratos_desconto,",")
		end if

		for i = 0 to ubound(contratos_incluir)

			objConexao.Execute("INSERT INTO TB_PRODUCAO_DESCONTOS_CONTRATOS (cod_contrato, cod_producao_desconto) VALUES ("&sqlTexto(contratos_incluir(i))&", "&sqlNumero(aux)&")")
		next

	fecharConexao()

	response.redirect(redirecionar&"?comando=alterar&cod_producao="&cod_producao&"&cod_producao_desconto="&aux)
end sub

sub alterarRegistro(redirecionar)

	call lerFormulario()

	abrirConexao()

		objConexao.Execute("UPDATE TB_PRODUCAO_DESCONTOS SET "&_
					" aliquota_padrao = "&sqlNumero(aliquota_padrao)&","&_
					" descricao = "&sqlTexto(descricao)&","&_
					" aplicar_todos_contratos = "&sqlTexto(aplicar_todos_contratos)&_
					" WHERE cod_producao_desconto = "&sqlNumero(cod_producao_desconto))

		'INSERIMOS OS CONTRATOS NA TABELA DE RELACIONAMENTO
		if aplicar_todos_contratos = "S" then
			contratos_incluir = contratos
		else
			contratos_incluir = split(contratos_desconto,",")
		end if

		'DELETAMOS PARA INCLUIR NOVAMENTE
		objConexao.Execute("DELETE FROM TB_PRODUCAO_DESCONTOS_CONTRATOS WHERE cod_producao_desconto = "&sqlNumero(cod_producao_desconto ))

		for i = 0 to ubound(contratos_incluir)

			objConexao.Execute("INSERT INTO TB_PRODUCAO_DESCONTOS_CONTRATOS (cod_contrato, cod_producao_desconto) VALUES ("&sqlNumero(contratos_incluir(i))&", "&sqlNumero(cod_producao_desconto)&")")
		next

	fecharConexao()
	response.redirect(redirecionar&"?cod_producao="&cod_producao&"&cod_producao_desconto="&cod_producao_desconto)
end sub

sub excluirRegistro(redirecionar)
	
	call lerFormulario()
	
	abrirConexao()

		if cod_producao_desconto <> "" and cod_producao_desconto <> "0" then

			objConexao.Execute("DELETE FROM TB_PRODUCAO_DESCONTOS WHERE cod_producao_desconto = "&sqlNumero(cod_producao_desconto))

			objConexao.Execute("DELETE FROM TB_PRODUCAO_DESCONTOS_CONTRATOS WHERE cod_producao_desconto = "&sqlNumero(cod_producao_desconto))

			objConexao.Execute("DELETE FROM TB_PRODUCAO_DESCONTOS_PRODUTOS WHERE cod_producao_desconto = "&sqlNumero(cod_producao_desconto))

		end if
	
	fecharConexao()
	response.redirect(redirecionar&"?cod_producao="&COD_PRODUCAO&"&cod_producao_desconto="&cod_producao_desconto)
end sub

dim redirecionar_incluir, redirecionar_alterar, redirecionar_excluir
redirecionar_incluir = "dip_producoes_descontos_formulario.asp"
redirecionar_alterar = "dip_producoes_descontos.asp"
redirecionar_excluir = "dip_producoes_descontos.asp"
'Sele��o de a��o
if comando = "visualizar" then
	call lerRegistros()

elseif comando = "incluir" AND campo1 <> "" then
	call incluirRegistro(redirecionar_incluir)

elseif comando = "alterar" AND campo1 = "" then
	call lerRegistros()

elseif comando = "alterar" AND campo1 <> "" then
	call alterarRegistro(redirecionar_alterar)

elseif comando = "excluir" AND campo1 = "" then
	call lerRegistros()

elseif comando = "excluir" AND campo1 <> "" then
	call excluirRegistro(redirecionar_excluir)

end if

'---------------------------------------------------------------'
%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
		<link type="text/css" rel="stylesheet" media="all" href="css/bootstrap.min.css" />
		<link type="text/css" rel="stylesheet" media="screen" href="css/base.css"></link>
		<link type="text/css" rel="stylesheet" media="screen" href="css/chosen.min.css"></link>
		<link type="text/css" rel="stylesheet" media="all" type="text/css" href="css/modal-message.css">
		<title>StillosWEB - Sistema corporativo - Vers&atilde;o 2.0</title>
		<style>
			body{background-color: #fff;}
			.container{margin: 3px;}
		</style>
		<!-- #include virtual="gerencial_foto/files/includes/scripts.asp" -->
		<script type="text/javascript">
		$(function(){
			$(".selecionaContrato").click(function(){
				var temp = $(".selecionaContrato:checked").val();

				if(temp == "N")
					$("#contratos_desconto").removeClass('hide');
				else
					$("#contratos_desconto").addClass('hide');
			})
		});
		</script>
	</head>
	<body>
	<%
	'INICIO VALIDA��ES
	if not validaUsuario(2,,82, false) then
		cod_erro = 0
	'Validamos se foi passado o codigo da produção'
	elseif cod_producao = "" then
		cod_erro = 1
	end if

	if cod_erro <> "" then
		call dislayErro(cod_erro)
	end if
	'FIM VALIDA��ES
	%>
	<!-- SUB menu de a��es -->
	<div class="submenu">
		<strong>Submenu:</strong>
		<a href='dip_producoes_descontos.asp?cod_producao=<%=cod_producao%>'>Voltar</a>
	</div>
<%
Dim qtd_tabs : qtd_tabs = 2
%>
<% if comando = "excluir" then %>
	<div class="alert alert-danger">
		<strong>
			Você deseja excluir este desconto, esta ação não poderá ser desfeita.
		</strong>
	</div>
<% end if%>
<div id="tabsJ">
	<ul>
		<li id="liTab1" class="current">
			<a href="javascript:tabsMostra('tab1',<%=qtd_tabs%>,1);">
				<span>Dados do Pedido de Compra</span>
			</a>
		</li>
		<li id="liTab2">
			<a href="javascript:tabsMostra('tab2',<%=qtd_tabs%>,2);">
				<span>Produtos</span>
			</a>
		</li>
	</ul>
</div>

<div id="tabsCaixa" style="height:350px;">
	<div id="tab1"><br />&nbsp;
		<form method="post" name="FORMULARIO_PADRAO">
			<table border="1" width="98%" cellpadding="3" align="center">
				<% if comando <> "incluir" then%>
				<tr>
					<td align="left"><strong>Cód. desconto</strong></td>
					<td><%=cod_producao_desconto%></td>
				</tr>
				<%end if%>
				<tr>
					<td align="left"><strong>Descrição</strong></td>
					<td>
						<% if comando = "incluir" or comando = "alterar" then %>
							<input type="text" name="descricao" value="<%=descricao%>" class='inputtamnho3' maxlength='255' />
						<% 
						else 
							response.write(descricao)
						end if
						%>
					</td>
				</tr>

				<tr>
					<td align="left" width="20%"><strong>Aliquota Padrão</strong></td>
					<td width="80%">
						<% if comando = "incluir" or comando = "alterar" then %>
							<input type="text" name="aliquota_padrao" value="<%=aliquota_padrao%>" class="padrao_numero inputtamnho2" maxlength='10' />
						<% 
						else 
							response.write(aliquota_padrao)
						end if
						%>
					</td>
				</tr>
				<tr>
					<td align="left"><strong>Aplicar a todos os contratos:</strong></td>
					<td>
						<% if comando = "incluir" or comando = "alterar" then %>
							<table border="0">
							<tr>
								<td>
									<label>
										<input type="radio" class="inputradio selecionaContrato" name='aplicar_todos_contratos' value="S" <% if aplicar_todos_contratos = "S" then%> checked='checked'<%end if%>>
										<strong>Sim</strong>
									</label>
								</td>
							</tr>
							<tr>
								<td>
									<label>
										<input type="radio" class="inputradio selecionaContrato" name='aplicar_todos_contratos' value="N" <% if aplicar_todos_contratos = "N" then%> checked='checked'<%end if%>>
										<strong>Não</strong>
									</label>
								</td>
							</tr>
							<tr>
								<td>
									<div id='contratos_desconto' <% if aplicar_todos_contratos <> "N" then%> class='hide' <%end if%> style="margin-left:20px; margin-top:5px;">
									<%
									for i = 0 to ubound(contratos)
										%>
										<label>
											<input type="checkbox" <% if instr(1, contratos_desconto, contratos(i) , 1) then %> checked="checked"<% end if %> class="inputradio" value="<%=contratos(i)%>" name='contratos_desconto'>
											<%=contratos(i)%>
										</label>
										<%
									next
									%>
									</div>
								</td>
							</tr>
							</table>
							<%
						else 
							response.write(contratos_desconto)
						end if
						%>
					</td>
				</tr>
				<tr>
					<td colspan="2" align="center">
						<a href='#' class='enviar'>
							<img src="images/layout/btnGravar.jpg">
						</a>
					</td>
				</tr>
			</table>
			<input type="hidden" name="campo1" value="1">
			<input type="hidden" name="cod_producao" value="<%=cod_producao%>">
			<input type="hidden" name="cod_producao_desconto" value="<%=cod_producao_desconto%>">
			</form>
			<br />
		</div>
		<div id="tab2" class="tabs">
			<br>&nbsp;
			<%if comando = "incluir" then%>
				<div class='alert alert-info'>
					<strong>Grave o Desconto antes de incluir Produtos</strong>
				</div>
			<%else%>
				<iframe src="dip_producoes_descontos_formulario_produtos.asp?comando_externo=<%=comando%>&cod_producao_desconto=<%=cod_producao_desconto%>" frameborder='0' width="100%" height="290px" id='iframeArquivo'></iframe>
			<%end if%>
		</div>
		</div>
	<!-- #include virtual="gerencial_foto/files/includes/scripts_rodape.asp" -->
	<script type="text/javascript">
	$(function(){
		$(".enviar").click(function(event){
			event.preventDefault();
			$("form[name=FORMULARIO_PADRAO]").submit();
		});
	})
	</script>
	</body>
</html>