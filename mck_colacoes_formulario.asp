<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 1
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
  Dim titulo_pagina
  titulo_pagina = "MCK - Colações"

  Dim calendario_necessario
  calendario_necessario = "NÃO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->

<%
'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 10/10/2011 16:04:06

'Declara vari�veis para utiliza��o do sistema
Dim cod_mack_colacao, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario, cod_mack_lista_provaveis_formandos_nome

redirecionar_incluir	= "mck_colacoes.asp"
redirecionar_alterar	= "mck_colacoes.asp"
redirecionar_excluir	= "mck_colacoes.asp"
redirecionar_validacao	= "mck_colacoes.asp"
msgAoExcluir			= "Tem certeza que deseja excluir esta Colação?"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
cod_mack_lista_provaveis_formandos_nome = Request.QueryString("cod_mack_lista_provaveis_formandos_nome")

'Declara vari�veis
Dim cod_mack_temporada, cod_contrato, descricao, unidade, data_colacao, local, sequencia

If comando = Empty or comando = "incluir" Then
comando = "incluir"
titulo_pagina = titulo_pagina&" / Incluir"
End if

If comando = "incluir" and campo1 <> Empty Then
	'Incluir novo registro
	incluirRegistro(redirecionar_incluir)
	titulo_pagina = titulo_pagina&" / Inclusão"
Elseif comando = "incluir" and campo1 = Empty then
''	getSequencia()
Elseif comando = "alterar" and campo1 = Empty Then
	'Leitura inicial para possível Alteração
	cod_mack_colacao = Request.QueryString("cod_mack_colacao")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "alterar" and campo1 <> Empty Then
	'Gravar Alteração
	cod_mack_colacao = Request.Form("cod_mack_colacao")
	alterarRegistro(redirecionar_alterar)
	titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 
	'Leitura inicial para possível Exclusão
	cod_mack_colacao = Request.QueryString("cod_mack_colacao")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
	'Exclui registro
	cod_mack_colacao = Request.Form("cod_mack_colacao")
	excluirRegistro(redirecionar_excluir)
	titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
	'Visualiza registro
	cod_mack_colacao = Request.QueryString("cod_mack_colacao")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / Visualização"
End If

sub getSequencia()
	
AbrirConexao()

	dim rs 
	set rs = objConexao.execute("SELECT max(sequencia) from TB_MACK_COLACOES")
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

	cod_mack_colacao = Request.Form("cod_mack_colacao")
	cod_mack_temporada = Request.Form("cod_mack_temporada")
	descricao = Request.Form("descricao")
	unidade = UCASE(Request.Form("unidade"))
	data_colacao = Request.Form("data_colacao")
	local = Request.Form("local")
	sequencia = Request.form("sequencia")
End sub

Sub lerRegistro()
	
	dim rs

	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM TB_MACK_COLACOES WHERE cod_mack_colacao = "&cod_mack_colacao)
		if not rs.eof then
			'cod_mack_colacao = rs("cod_mack_colacao")
			cod_mack_temporada = rs("cod_mack_temporada")
			cod_contrato = rs("cod_contrato")
			descricao = rs("descricao")
			unidade = rs("unidade")
			data_colacao = arrumaData(rs("data_colacao"))
			local = rs("local")
			sequencia = rs("sequencia")
		end if
		rs.Close
		Set rs=Nothing
	FecharConexao()

End sub

Sub incluirRegistro(redirecionar)

	lerFormulario()

	AbrirConexao()

	dim rs
	'gereamos o contrato somente foto interno
	
	'buscamos o ultimo contrato gerado
	Set rs2=objConexao.Execute("SELECT Count(COD_CONTRATO) AS COD_CONTRATO FROM TB_ENTIDADES WHERE SOMENTE_FOTO = 1")
	If not rs2.Eof Then
		COD_CONTRATO = formatNumber(0&rs2(0),0)
		COD_CONTRATO = (COD_CONTRATO+1)/1
		COD_CONTRATO = Replace(COD_CONTRATO,",","")
		COD_CONTRATO = Replace(COD_CONTRATO,".","")
	Else
		COD_CONTRATO = 1
	End If

	COD_CONTRATO = COD_CONTRATO&"F"
		
	rs2.Close
	Set rs2=Nothing

	'Buscamos na temporada o ano e semestre de comclus�o 
	Set rs2=objConexao.Execute("SELECT ano_conclusao, semestre_conclusao FROM TB_MACK_TEMPORADAS WHERE cod_mack_temporada = "&sqlNumero(cod_mack_temporada))
	If not rs2.Eof Then
		ano_conclusao = rs2("ano_conclusao")
		semestre_conclusao = rs2("semestre_conclusao")
	End If
	rs2.Close
	Set rs2=Nothing

	'Monstamos a string de descricao
	descricao_contrato = descricao&" - "&arrumadata(data_colacao)

	'inserimos
	objConexao.Execute("INSERT INTO TB_ENTIDADES (COD_CONTRATO, DESCRICAO, SEMESTRE, ANO, SOMENTE_FOTO) VALUES ("&sqlTexto(COD_CONTRATO)&", "&sqlTexto(descricao_contrato)&", "&sqlTexto(semestre_conclusao)&", "&sqlTexto(ano_conclusao)&",1)")

	dim sql : sql = "INSERT INTO TB_MACK_COLACOES ("&_
						"cod_mack_temporada, "&_
						"COD_CONTRATO, "&_					
						"descricao, "&_
						"unidade, "&_
						"data_colacao, "&_
						"local, "&_
						"sequencia "&_
						") VALUES  ( "&_
						sqlNumero(cod_mack_temporada)&", "&_
						sqlTexto(COD_CONTRATO)&", "&_
						sqlTexto(descricao)&", "&_
						sqlTexto(unidade)&", "&_
						sqlData(data_colacao)&", "&_
						sqlTexto(local)&", "&_
						sqlNumero(0&sequencia)&" "&_
						")"

	objConexao.Execute(sql)

	'Encontrar chave criada
	Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_MACK_COLACOES LIMIT 1")
		If not rs.Eof Then
			cod_mack_colacao = rs("ULTIMO_ID")
		End if 
	rs.Close
	Set rs=Nothing

	FecharConexao()
	Response.Redirect(redirecionar&"?cod_mack_colacao="&cod_mack_colacao&"&codmsg=1&filtro=2&filtro_cod_mack_temporada="&cod_mack_temporada)
End sub

Sub alterarRegistro(redirecionar)

	lerFormulario()

	AbrirConexao()

	dim sql : sql = "UPDATE TB_MACK_COLACOES SET"&_
					 " cod_mack_temporada = "&sqlNumero(cod_mack_temporada)&_
					 ", descricao = "&sqltexto(descricao)&_
					 ", unidade = "&sqltexto(unidade)&_
					 ", data_colacao = "&sqlData(data_colacao)&_
					 ", local = "&sqltexto(local)&_
					 ", sequencia = "&sqltexto(0&sequencia)&_
					 " WHERE cod_mack_colacao = "&sqlNumero(cod_mack_colacao)

	objConexao.execute(sql)

	FecharConexao()

	Response.Redirect(redirecionar&"?cod_mack_colacao="&cod_mack_colacao&"&codmsg=2&filtro=2&filtro_cod_mack_temporada="&cod_mack_temporada)
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()

		dim rs, pode_excluir

		'verificamos se existe Colações para esta temporada
		Set rs = objConexao.execute("SELECT * FROM TB_MACK_FORMANDOS WHERE cod_mack_colacao = "&sqlNumero(cod_mack_colacao))
		if not rs.eof then
			pode_excluir = false
		Else
			pode_excluir = true
		end if
		set rs = nothing

		if pode_excluir then
			objConexao.Execute("DELETE FROM TB_MACK_COLACOES WHERE cod_mack_colacao = "&cod_mack_colacao)
			Response.Redirect(redirecionar&"?cod_mack_colacao="&cod_mack_colacao&"&codmsg=3")
		Else
			Response.Redirect(redirecionar&"?cod_mack_colacao="&cod_mack_colacao&"&codmsg=4")
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
						<span class='textoMaior'><strong><%=cod_mack_colacao%></strong></span><%Else%>
						<%
						If not isNull(cod_mack_colacao) Then
							response.write(cod_mack_colacao)
						End If
					End If %>
					</div>
				</td>
			</tr>
			<%end if%>
			<tr>
				<td width='27%'><strong>Descrição *</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='descricao' value="<%=descricao%>" class='inputTamanho3' type='text' maxlength='255' />
						<% Else 
							If not isNull(descricao) Then
								response.write(descricao)
							End If
						End If %>
					</div>
				</td>
			</tr>

			<tr>
				<td width='27%'><strong>Temporada *</strong></td>
				<td width='73%'>
					<div align='left'>
					
					<%=padraoFormularioSelect("cod_mack_temporada", "", "inputTamanho2",""," ","Selecione", cod_mack_temporada,"SELECT cod_mack_temporada, concat(semestre_conclusao,'/',ano_conclusao) as descricao FROM TB_MACK_TEMPORADAS ",comando, false, false, false) %>
					</div>
				</td>
			</tr>

			<% if comando <> "incluir" then %>
			<tr>
				<td width='27%'><strong>Contrato somente foto *</strong></td>
				<td width='73%'>
					<div align='left'>
						<%
						If not isNull(cod_contrato) and cod_contrato = Empty Then
							response.write(cod_contrato)
						End If
						%>
					</div>
				</td>
			</tr>
			<% end if %>
	

			<tr>
				<td width='27%'><strong>Departamento *</strong></td>
				<td width='73%'>
					<div align='left'>
					<%If comando <> "visualizar" and comando <> "excluir"  Then%>
						<input name='unidade' value="<%=unidade%>" class='inputTamanho1' type='text' maxlength='50' />
					<% Else 
						If not isNull(unidade) Then
							response.write(unidade)
						End If
					End If %>

					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>Data Colação *</strong></td>
				<td width='73%'>
					<div align='left'>
					<%If comando <> "visualizar" and comando <> "excluir"  Then%>
						<input name='data_colacao' class='inputTamanho1 padrao_data' value="<%=data_colacao%>" type='text' />
					<%
					Else
						If not isNull(data_colacao) Then
							response.write(data_colacao)
						End If
					End If %>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>Sequência </strong></td>
				<td width='73%'>
					<div align='left'>
					<%If comando <> "visualizar" and comando <> "excluir"  Then%>
						<input name='sequencia' class='inputTamanho1 ' value="<%=sequencia%>" type='text' />
					<%
					Else
						If not isNull(sequencia) Then
							response.write(sequencia)
						End If
					End If %>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>Local *</strong></td>
				<td width='73%'>
					<div align='left'>
					
					<%=padraoFormularioSelect("local", "", "inputTamanho1",""," required='required' ","Selecione", local,"PADRAO:H#Higienópolis|C#Campinas|A#Alphaville|S#Sumar�",comando, false, false, false) %>
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
	<input type='hidden' name='cod_mack_colacao' value="<%=cod_mack_colacao%>" />�
		
</form>

<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('descricao#Descrição#1|unidade#Departamento#1|data_colacao#Data Colação#1|LOCAL#Local#2',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>
<!-- #include file="files/includes/miolo-inferior.asp" -->