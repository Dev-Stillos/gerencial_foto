<% dim script_valida_form_master : script_valida_form_master = "NAO" %>
<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<!-- #include file="files/includes/functions_local.asp" -->
<%

function pordeAlterarDesconto(cod_producao_desconto, conexao_aberta)
	

	if not conexao_aberta then
		abrirConexao()
	end if 

	dim retorno 

	set rs = objConexao.Execute("SELECT PRODUCAO_FINALIZADA FROM VW_TB_PRODUCAO_DESCONTOS WHERE cod_producao_desconto = "&sqlNumero(cod_producao_desconto))
	if  not rs.Eof then
		if rs("PRODUCAO_FINALIZADA")  = 1 then
			retorno = false
		else
			retorno = true
		end if
	else
		retorno = false
	end if
	rs.close
	set rs = nothing


	if not conexao_aberta then
		fecharConexao()
	end if 

pordeAlterarDesconto = retorno
end function


sub dislayErro(cod_erro)

	dim erro(4)
	erro(0) = "Você Não possui este acesso."
	erro(1) = "Um erro inesperado ocorreu tente novamente mais tarde, se o erro persistir informe o departamento de TI"
	erro(2) = "Esta produção Não possui contratos, ou todos os contratos já estão relacionados a uma produção"
	erro(3) = ""

	response.write("<div class='alert alert-danger'>")
	response.write(erro(cod_erro))
	response.write("</div>")
	response.end

end sub

if not validaUsuario(2,,82, false) then
	response.write("Você Não tem acesso")
	response.end
end if

Dim rs, cod_producao_desconto, campo1, comando
Dim cod_producao_desconto_produto, cod_produto, aliquota
Dim comando_externo

campo1 = request.form("campo1")
cod_producao_desconto = request.queryString("cod_producao_desconto")
cod_producao_desconto_produto = request.queryString("cod_producao_desconto_produto")

if cod_producao_desconto = "" then
	cod_producao_desconto = request.form("cod_producao_desconto")
end if

if cod_producao_desconto_produto = "" then
	cod_producao_desconto_produto = request.form("cod_producao_desconto_produto")
end if

comando_externo = request.queryString("comando_externo")
comando = request.queryString("comando")

if comando = "" then
	comando = request.form("comando")
	if comando = "" then
		comando = "incluir"
	end if
end if

sub lerRegistros()
	dim rs

	abrirConexao()

	set rs = objConexao.Execute("SELECT * FROM TB_PRODUCAO_DESCONTOS_PRODUTOS WHERE cod_producao_desconto_produto = "&sqlNumero(cod_producao_desconto_produto))
	if not rs.Eof then
		aliquota = rs("aliquota")
		cod_produto = rs("cod_produto")
	end if 
	rs.close
	set rs = nothing

	fecharConexao()
end sub

sub lerFormulario()

	cod_producao_desconto_produto  = request.form("cod_producao_desconto_produto")
	cod_producao_desconto  = request.form("cod_producao_desconto")
	aliquota = request.form("aliquota")
	cod_produto = request.form("cod_produto")


end sub

sub incluirRegistro(redirecionar)

	call lerFormulario()

	dim i, rs, aux, contratos_incluir

	abrirConexao()


	if  pordeAlterarDesconto(cod_producao_desconto, true) then

		if cod_produto <> "" then
			objConexao.execute("INSERT INTO TB_PRODUCAO_DESCONTOS_PRODUTOS (cod_producao_desconto, cod_produto, aliquota, cod_usuario, data_cadastro, modificado) VALUES ("&sqlNumero(cod_producao_desconto)&","&sqlNumero(cod_produto)&", "&sqlNumero(aliquota)&", "&sqlNumero(session("cod_usuario"))&", NOW(), NOW() )")
		else
			'objConexao.Execute("INSERT INTO TB_PRODUCAO_DESCONTOS_PRODUTOS (cod_producao_desconto, cod_produto, aliquota) SELECT "&sqlNumero(cod_producao_desconto)&", cod_produto,  "&sqlNumero(aliquota)&" FROM TB_PRODUTOS WHERE ativo = 1 ")
		end if
	else
		response.end
	end if

	fecharConexao()

	response.redirect(redirecionar&"?comando=incluir&cod_producao_desconto="&cod_producao_desconto)


end sub



sub alterarRegistro(redirecionar)

	call lerFormulario()

	abrirConexao()

	if  pordeAlterarDesconto(cod_producao_desconto, true) then

		if cod_produto <> "" then

			objConexao.Execute("UPDATE TB_PRODUCAO_DESCONTOS_PRODUTOS SET "&_
					" cod_producao_desconto = "&sqlNumero(cod_producao_desconto)&","&_
					" cod_produto = "&sqlNumero(cod_produto)&","&_
					" aliquota = "&sqlNumero(aliquota)&_
					" modificado = NOW()"&_
					" WHERE cod_producao_desconto_produto = "&sqlNumero(cod_producao_desconto_produto))
		else
			'objConexao.Execute("UPDATE TB_PRODUCAO_DESCONTOS_PRODUTOS SET "&_
			''		" aliquota = "&sqlNumero(aliquota)&_
			''		" WHERE cod_producao_desconto = "&sqlNumero(cod_producao_desconto))
		end if

	else
		response.end
	end if

	fecharConexao()
	response.redirect(redirecionar&"?cod_producao_desconto="&cod_producao_desconto)
end sub

sub excluirRegistro(redirecionar)
	
	call lerFormulario()
	
	abrirConexao()
		
		if  pordeAlterarDesconto(cod_producao_desconto, true) then
		
			if cod_produto <> "" then
				objConexao.Execute("DELETE FROM TB_PRODUCAO_DESCONTOS_PRODUTOS WHERE cod_producao_desconto_produto = "&sqlNumero(cod_producao_desconto_produto))
			else
				'objConexao.Execute("DELETE FROM TB_PRODUCAO_DESCONTOS_PRODUTOS WHERE cod_producao_desconto = "&sqlNumero(cod_producao_desconto))
			end if
		else
			response.end
		end if

	fecharConexao()
	response.redirect(redirecionar&"?cod_producao_desconto="&cod_producao_desconto)
end sub

dim redirecionar_incluir, redirecionar_alterar, redirecionar_excluir
redirecionar_incluir = "dip_producoes_descontos_formulario_produtos.asp"
redirecionar_alterar = "dip_producoes_descontos_formulario_produtos.asp"
redirecionar_excluir = "dip_producoes_descontos_formulario_produtos.asp"
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
	elseif cod_producao_desconto = "" then
		cod_erro = 1
	end if

	if cod_erro <> "" then
		call dislayErro(cod_erro)
	end if
	'FIM VALIDA��ES
	
if comando = "excluir" then %>
	<div class="alert alert-danger">
		<strong>
			Você deseja excluir este produto, esta ação Não poderá ser desfeita.
		</strong>
	</div>
<% end if

'response.write comando

if comando = "incluir" then
	
	abrirConexao()

	Set rs =objConexao.execute("SELECT aliquota_padrao FROM VW_TB_PRODUCAO_DESCONTOS WHERE cod_producao_desconto = "&sqlNumero(cod_producao_desconto))
	if not rs.Eof then
		aliquota = rs(0)
	end if 
	rs.close
	set rs = nothing

	fecharConexao()

end if


if comando_externo <> "visualizar" and comando_externo <> "excluir" then
%>
	<div class='submenu'>
		<form method="post" name="FORMULARIO_PADRAO">
			<table align="center">
				<tr>
					<th>Produto</th>
					<th>Desconto(%)</th>
					<th rowspan="2">
						<a href="javascript:void()" class='enviar'>
							<% if comando = "incluir" then %>
								<img src="images/layout/btnIncluir.jpg">
							<% elseif comando = "alterar" then %>
								<img src="images/layout/btnGravar.jpg">
							<% elseif comando = "excluir" then %>
								<img src="images/layout/btnExcluir.jpg">
							<% end if%>	
						</a>
						<%
							if comando = "alterar" OR comando = "excluir" then
						%>
							<br />
							<a href='?comando=incluir&cod_producao_desconto=<%=cod_producao_desconto%>' >Cancelar</a>
						<% end if %>

					</th>
				</tr>
				<tr>
					<td>
						<%=padraoFormularioSelect("cod_produto","","inputTamanho2","","","Selecione",cod_produto,"SELECT cod_produto, nome_produto FROM TB_PRODUTOS WHERE (cod_produto NOT IN (SELECT cod_produto FROM TB_PRODUCAO_DESCONTOS_PRODUTOS WHERE cod_producao_desconto = "&cod_producao_desconto&") OR cod_produto = '"&sqlNumero(cod_produto)&"') ORDER BY nome_produto ASC","incluir",false,false,false)%></td>

					</td>
					<td>
						<input name="aliquota" value='<%=arrumaNumero(aliquota,"",2)%>' class='inputTamanho1 padrao_valor' maxlength='6'>
					</td>
				</tr>
			</table>
			<input type="hidden" name="campo1" value="1">
			<input type="hidden" name="comando" value="<%=comando%>">
			<input type="hidden" name="cod_producao_desconto" value="<%=cod_producao_desconto%>">
			<input type="hidden" name="cod_producao_desconto_produto" value="<%=cod_producao_desconto_produto%>">
		</form>
	</div>
<% end if %>
	<script type="text/javascript">
	$(function(){
		$(".enviar").click(function(event){
			event.preventDefault();
			var cod_produto = $("select[name=cod_produto]").val()
			var aliquota = $("input[name=aliquota]").val()

		if (aliquota != '' && cod_produto != '')
				$("form[name=FORMULARIO_PADRAO]").submit();
			else
			{
				alert('Preencha os campos corretamente');
				return 0;
			}
		});
	})
	</script>
	<div>
		<%
		abrirConexao()
		Set rs = objConexao.execute("SELECT * FROM VW_TB_PRODUCAO_DESCONTOS_PRODUTOS WHERE cod_producao_desconto = "&sqlNumero(cod_producao_desconto)&" ORDER BY nome_produto")
		if not rs.Eof  then
		%>
		<table width="99%" border="1" align="center">
			<tr>
				<th> # </th>
				<th>Produto</th>
				<th>Desconto (%)</th>
				<th>&nbsp;</th>
			</tr>	
		<%
		while not rs.Eof
		%>
			<tr>
				<td align="center"><%=rs("cod_producao_desconto_produto")%></td>
				<td><%=rs("nome_produto")%></td>
				<td align="right"><%=arrumaNumero(rs("aliquota"),0,2) %></td>
				<td width="40px">
					<table border="0" width="40px">
						<tr>
							<td>
								<a href='?comando=alterar&cod_producao_desconto=<%=cod_producao_desconto%>&cod_producao_desconto_produto=<%=rs("cod_producao_desconto_produto")%>'>
									<img src="images/icons/icon_alterar.png" border="0">
								</a>
							</td>
							<td>
								<a href='?comando=excluir&cod_producao_desconto=<%=cod_producao_desconto%>&cod_producao_desconto_produto=<%=rs("cod_producao_desconto_produto")%>'>
									<img src="images/icons/icon_excluir.png" border="0">
								</a>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		<%
		rs.movenext
		Wend 
		%>
		</table>
		<% else %>
			<div class='alert alert-warning'>
				<strong>
					Nenhum produto encontrado.
				</strong>
			</div>
		<%
		end if
		rs.close
		set rs = nothing

		fecharConexao()
		%>
	</div>
	<!-- #include virtual="gerencial_foto/files/includes/scripts_rodape.asp" -->
	</body>
</html>