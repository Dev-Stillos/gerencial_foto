<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 1
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
  Dim titulo_pagina
  titulo_pagina = "MCK - Temporadas"

  Dim calendario_necessario
  calendario_necessario = "NÃO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->

<%
'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 10/10/2011 16:04:06

'Declara vari�veis para utiliza��o do sistema
Dim cod_mack_temporada, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario, cod_mack_lista_provaveis_formandos_nome

redirecionar_incluir	= "mck_temporadas.asp"
redirecionar_alterar	= "mck_temporadas.asp"
redirecionar_excluir	= "mck_temporadas.asp"
redirecionar_validacao	= "mck_temporadas.asp"
msgAoExcluir			= "Tem certeza que deseja excluir esta temporada?"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
cod_mack_lista_provaveis_formandos_nome = Request.QueryString("cod_mack_lista_provaveis_formandos_nome")

'Declara vari�veis
Dim sequencia, ano_conclusao, semestre_conclusao, inicio, fim

If comando = Empty Then
comando = "incluir"
titulo_pagina = titulo_pagina&" / Incluir"
End if

If comando = "incluir" and campo1 <> Empty Then
	'Incluir novo registro
	incluirRegistro(redirecionar_incluir)
	titulo_pagina = titulo_pagina&" / Inclusão"
Elseif comando = "incluir" and campo1 = Empty then
	getSequencia()
Elseif comando = "alterar" and campo1 = Empty Then
	'Leitura inicial para possível alteração
	cod_mack_temporada = Request.QueryString("cod_mack_temporada")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "alterar" and campo1 <> Empty Then
	'Gravar alteração
	cod_mack_temporada = Request.Form("cod_mack_temporada")
	alterarRegistro(redirecionar_alterar)
	titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 
	'Leitura inicial para possível exclusão
	cod_mack_temporada = Request.QueryString("cod_mack_temporada")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
	'Exclui registro
	cod_mack_temporada = Request.Form("cod_mack_temporada")
	excluirRegistro(redirecionar_excluir)
	titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
	'Visualiza registro
	cod_mack_temporada = Request.QueryString("cod_mack_temporada")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / Visualização"
End If

sub getSequencia()
	
AbrirConexao()

	dim rs 
	set rs = objConexao.execute("SELECT max(sequencia) from TB_MACK_TEMPORADAS")
	if not rs.Eof then
		sequencia = rs(0) + 1
	Else
		sequencia = 1
	end if
	rs.Close
	set rs = nothing
FecharConexao()
end sub

Sub lerFormulario()

	cod_mack_temporada = Request.Form("cod_mack_temporada")
	sequencia = Request.Form("sequencia")
	ano_conclusao = Request.Form("ano_conclusao")
	semestre_conclusao = Request.Form("semestre_conclusao")
	inicio = Request.Form("inicio")
	fim = Request.Form("fim")

End sub

Sub lerRegistro()
	
	dim rs

	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM TB_MACK_TEMPORADAS WHERE cod_mack_temporada = "&cod_mack_temporada)
		if not rs.eof then
			cod_mack_temporada = Rs("cod_mack_temporada")
			sequencia = Rs("sequencia")
			ano_conclusao = rs("ano_conclusao")
			semestre_conclusao = rs("semestre_conclusao")
			inicio = arrumadata(Rs("inicio"))
			fim = arrumadata(Rs("fim"))
		end if
		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)

	lerFormulario()

	AbrirConexao()

	dim rs
	dim sql : sql = "INSERT INTO TB_MACK_TEMPORADAS ("&_
						"sequencia, "&_
						"ano_conclusao, "&_
						"semestre_conclusao, "&_
						"inicio, "&_
						"fim, "&_
						"data_cadastro, "&_
						"cod_usuario, "&_
						"modificado "&_
						") VALUES  ( "&_
						sqlNumero(sequencia)&", "&_
						sqlTexto(ano_conclusao)&", "&_
						sqlTexto(semestre_conclusao)&", "&_
						sqlData(inicio)&", "&_
						sqlData(fim)&", "&_
						" now(), "&_
						sqlNumero(session("cod_usuario"))&", "&_
						" now() "&_
						")"

	response.write sql

	objConexao.Execute(sql)

	'Encontrar chave criada
	Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_MACK_TEMPORADAS LIMIT 1")
		If not rs.Eof Then
			cod_mack_temporada = rs("ULTIMO_ID")
		End if 
	rs.Close
	Set rs=Nothing

	FecharConexao()
	Response.Redirect(redirecionar&"?cod_mack_temporada="&cod_mack_temporada&"&codmsg=1")
End sub

Sub alterarRegistro(redirecionar)

	lerFormulario()

	AbrirConexao()

	dim sql : sql = "UPDATE TB_MACK_TEMPORADAS SET"&_
						 " sequencia = "&sqlNumero(sequencia)&_
						 ", ano_conclusao = "&sqltexto(ano_conclusao)&_
						 ", semestre_conclusao = "&sqltexto(semestre_conclusao)&_
						 ", inicio = "&sqlData(inicio)&_
						 ", fim = "&sqlData(fim)&_
						 ", modificado = now()"&_
						 ", cod_usuario = "&sqlNumero(session("cod_usuario"))&_
						 " WHERE cod_mack_temporada = "&sqlNumero(cod_mack_temporada)

	objConexao.execute(sql)

	FecharConexao()

	Response.Redirect(redirecionar&"?cod_mack_temporada="&cod_mack_temporada&"&codmsg=2")
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()

		dim rs, pode_excluir

		'verificamos se existe cola��es para esta temporada
		Set rs = objConexao.execute("SELECT * FROM TB_MACK_COLACOES WHERE cod_mack_temporada = "&sqlNumero(cod_mack_temporada))
		if not rs.eof then
			pode_excluir = false
		Else
			pode_excluir = true
		end if
		rs.Close
		set rs = nothing

		if pode_excluir then
			objConexao.Execute("DELETE FROM TB_MACK_TEMPORADAS WHERE cod_mack_temporada = "&cod_mack_temporada)
			Response.Redirect(redirecionar&"?cod_mack_temporada="&cod_mack_temporada&"&codmsg=3")
		Else
			Response.Redirect(redirecionar&"?cod_mack_temporada="&cod_mack_temporada&"&codmsg=4")
		end if
	FecharConexao()
End sub
'----------------------------------------------------------------------------------------
%>
<h2><%=titulo_pagina%></h2><br />
<!-- #include file="includes/sub_menu_mck.asp" -->
<br />

<%If comando = "excluir" Then%>
	<div class='msgExcluir'><%=msgAoExcluir%></div>
<%End If%>

<style type="text/css">
	fieldset{
		width: 820px;
		margin:10px auto 5px auto ;
	}
</style>

<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post'>
	<!--Dados Gerais-->
	<fieldset>
		<legend>Dados Gerais</legend>
		<table width='800px' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
			<%If comando <> "incluir" Then %><tr>
			<td width='27%'><strong>Cód. temporada</strong></td>
				<td width='73%'>
					<div align='left'>
					<% If comando <> "visualizar" and comando <> "excluir"  Then%>
						<span class='textoMaior'><strong><%=cod_mack_temporada%></strong></span><%Else%>
						<%
						If not isNull(cod_mack_temporada) Then
							response.write(cod_mack_temporada)
						End If
					End If %>
					</div>
				</td>
			</tr>
			<%end if%>
			<tr>
				<td width='27%'><strong>Conclusão *</strong><em>Semestre e ano de conclusão dos formandos</em></td>
				<td width='73%'>
					<div align='left'>
						<%=padraoFormularioSelect("semestre_conclusao", "", "inputTamanho1","","","Selecione", semestre_conclusao,"PADRAO:J#JUL|D#DEZ",comando, false, false, false) %>
						/
						<%=padraoFormularioSelect("ano_conclusao", "", "inputTamanho1","","","Selecione", ano_conclusao,"PADRAO:2026#2026|2027#2027|2028#2028|2029#2029|2030#2030|2031#2031|2032#2032|2033#2033|2034#2034|2035#2035",comando, false, false, false) %>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>Sequência *</strong></td>
				<td width='73%'>
					<div align='left'>
					<%If comando <> "visualizar" and comando <> "excluir"  Then%>
						<input name='sequencia' value="<%=sequencia%>" class='inputTamanho1 padrao_numero' type='text' maxlength='50' />
					<% Else 
						If not isNull(sequencia) Then
							response.write(sequencia)
						End If
					End If %>

					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>Data inicio *</strong></td>
				<td width='73%'>
					<div align='left'>
					<%If comando <> "visualizar" and comando <> "excluir"  Then%>
						<input name='inicio' class='inputTamanho2 padrao_data' value="<%=inicio%>" type='text' />
					<%
					Else
						If not isNull(inicio) Then
							response.write(inicio)
						End If
					End If %>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>Data Fim *</strong></td>
				<td width='73%'>
					<div align='left'>
					<%If comando <> "visualizar" and comando <> "excluir"  Then%>
						<input name='fim' class='inputTamanho2 padrao_data' value="<%=fim%>" type='text' />
					<%
					Else
						If not isNull(fim) Then
							response.write(fim)
						End If
					End If %>
					<%'=padraoFormularioSelect("cod_mack_colacao", "", "inputTamanho3",""," required='required' ","Selecione", cod_mack_colacao,"SELECT cod_mack_colacao, concat(descricao ,' - ',cast(DATE_FORMAT(data_colacao,'%d/%m/%Y') as char)) FROM TB_MACK_COLACOES ORDER BY descricao",comando, false, false, false) %>
					</div>
				</td>
			</tr>
			
		</table>	
		<br>

		<%If comando <> "visualizar" Then%>
		<table width='800px' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
		<tr>
			<td colspan='2' align='center'><div align='left'><strong>* - Campos obrigatórios</strong><br /></div></td>
		</tr>
		<tr>
			<% if comando <> "excluir" then%>
				<td colspan='2' align='center'><a href="javascript:salvar();"><img src="images/layout/btnGravar.jpg" /></a></td>
			<%else%>
				<td colspan='2' align='center'><a href="javascript:salvar();"><img src="images/layout/btnExcluir.jpg" /></a></td>
			<%end if %>
		</tr>
		</table>
	<%End If%>	
	</fieldset>
	<input type='hidden' value='1' name='campo1'  />
	<input type='hidden' name='cod_mack_temporada' value="<%=cod_mack_temporada%>" />
		
</form>

<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('SEMESTRE_CONCLUSAO#Semestre de Conclusão#2|ANO_CONCLUSAO#Ano de Conclusão#2|sequencia#Sequência#1|inicio#Data inicio#1|fim#Data fim#1',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>
<!-- #include file="files/includes/miolo-inferior.asp" -->