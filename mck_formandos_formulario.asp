<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 1
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
  Dim titulo_pagina
  titulo_pagina = "Formandos Identificados"

  Dim calendario_necessario
  calendario_necessario = "NÃO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%
'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 10/10/2011 16:04:06

'Declara vari�veis para utiliza��o do sistema
Dim cod_mack_formando, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario, cod_mack_lista_provaveis_formandos_nome

redirecionar_incluir	= "mck_formandos_formulario.asp"
redirecionar_alterar	= "mck_formandos.asp"
redirecionar_excluir	= "mck_formandos.asp"
redirecionar_validacao	= "mck_formandos.asp"
msgAoExcluir			= "Tem certeza que deseja excluir este formando?"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
cod_mack_lista_provaveis_formandos_nome = Request.QueryString("cod_mack_lista_provaveis_formandos_nome")

'Declara vari�veis
Dim cod_mack_colacao, nome, tia, curso, senha, numero_identificacao, nome_pai, nome_mae, telefone, celular, email_1, email_2, residencia_endereco, residencia_bairro, residencia_cep, residencia_cidade, residencia_estado, comercial_endereco, comercial_bairro, comercial_cep, comercial_cidade, comercial_estado, comercial_telefone, pais_endereco, pais_bairro, pais_cep, pais_cidade, pais_estado, pais_telefone, pais_celular, retirou_convites, data_retirou_convites, cod_mack_temporada

If comando = Empty Then
comando = "incluir"
titulo_pagina = titulo_pagina&" / Incluir"
End if

If comando = "incluir" and campo1 <> Empty Then
	'Incluir novo registro
	incluirRegistro(redirecionar_incluir)
	titulo_pagina = titulo_pagina&" / Inclusão"

ElseIf comando = "incluir" and campo1 = Empty Then
	'lemos o registro da lista de provaveis formandos	
	lerIdentificacao()
	titulo_pagina = titulo_pagina&" / Inclusão"

Elseif comando = "alterar" and campo1 = Empty Then
	'Leitura inicial para possível Alteração
	cod_mack_formando = Request.QueryString("cod_mack_formando")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "alterar" and campo1 <> Empty Then
	'Gravar Alteração
	cod_mack_formando = Request.Form("cod_mack_formando")
	alterarRegistro(redirecionar_alterar)
	titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 
	'Leitura inicial para possível Exclusão
	cod_mack_formando = Request.QueryString("cod_mack_formando")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
	'Exclui registro
	cod_mack_formando = Request.Form("cod_mack_formando")
	excluirRegistro(redirecionar_excluir)
	titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
	'Visualiza registro
	cod_mack_formando = Request.QueryString("cod_mack_formando")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / Visualização"
End If


'Fun��es
Sub lerIdentificacao()

	if cod_mack_lista_provaveis_formandos_nome <> empty then
	AbrirConexao()
		Set rs = objConexao.Execute("SELECT tmlpf.vigente, tmlpf.cod_mack_temporada, tmlpfn.*, concat(tnt.semestre_conclusao,'/',ano_conclusao) as descricao FROM TB_MACK_LISTAS_PROVAVEIS_FORMANDOS tmlpf  INNER JOIN TB_MACK_LISTAS_PROVAVEIS_FORMANDOS_NOMES tmlpfn on tmlpfn.cod_mack_lista_provaveis_formandos = tmlpf.cod_mack_lista_provaveis_formandos INNER JOIN TB_MACK_TEMPORADAS tnt on tmlpf.cod_mack_temporada = tnt.cod_mack_temporada WHERE tmlpfn.cod_mack_lista_provaveis_formandos_nome = "&sqlTexto(cod_mack_lista_provaveis_formandos_nome)&" AND vigente = 'S'")
		if not rs.Eof then
			nome = rs("nome")
			tia = rs("matricula")
			curso = rs("curso")
			departamento = rs("departamento")
			departamento_sigla = rs("departamento_sigla")

			cod_mack_temporada = mackenzieTemporadaAtiva(true)

			'removemos a data da Colação do curso
			curso = trim(mid(curso,1, instrrev(curso,"-") - 1 ))
		end if 
		rs.Close
		set rs = Nothing
		
		identificacao_sem_lista = false
	FecharConexao()
	end if

end sub


Sub lerFormulario()

	cod_mack_formando = Request.Form("cod_mack_formando")
	cod_mack_colacao = Request.Form("cod_mack_colacao")
	nome = Request.Form("nome")
	tia = Request.Form("tia")
	curso = Request.Form("curso")
	senha = Request.Form("senha")
	numero_identificacao = Request.Form("numero_identificacao")
	nome_pai = Request.Form("nome_pai")
	nome_mae = Request.Form("nome_mae")
	telefone = Request.Form("telefone")
	celular = Request.Form("celular")
	email_1 = Request.Form("email_1")
	email_2 = Request.Form("email_2")
	residencia_endereco = Request.Form("residencia_endereco")
	residencia_bairro = Request.Form("residencia_bairro")
	residencia_cep = Request.Form("residencia_cep")
	residencia_cidade = Request.Form("residencia_cidade")
	residencia_estado = Request.Form("residencia_estado")
	comercial_endereco = Request.Form("comercial_endereco")
	comercial_bairro = Request.Form("comercial_bairro")
	comercial_cep = Request.Form("comercial_cep")
	comercial_cidade = Request.Form("comercial_cidade")
	comercial_estado = Request.Form("comercial_estado")
	comercial_telefone = Request.Form("comercial_telefone")
	pais_endereco = Request.Form("pais_endereco")
	pais_bairro = Request.Form("pais_bairro")
	pais_cep = Request.Form("pais_cep")
	pais_cidade = Request.Form("pais_cidade")
	pais_estado = Request.Form("pais_estado")
	pais_telefone = Request.Form("pais_telefone")
	pais_celular = Request.Form("pais_celular")
	retirou_convites = Request.form("retirou_convites")

	cod_mack_lista_provaveis_formandos_nome = Request.form("cod_mack_lista_provaveis_formandos_nome")

End sub


Sub lerRegistro()
	
	dim rs

	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM TB_MACK_FORMANDOS WHERE cod_mack_formando = "&cod_mack_formando)
		if not rs.eof then
			cod_mack_formando = rs("cod_mack_formando")
			cod_mack_colacao = rs("cod_mack_colacao")
			nome = rs("nome")
			tia = rs("tia")
			curso = rs("curso")
			senha = rs("senha")
			numero_identificacao = rs("numero_identificacao")
			nome_pai = rs("nome_pai")
			nome_mae = rs("nome_mae")
			telefone = rs("telefone")
			celular = rs("celular")
			email_1 = rs("email_1")
			email_2 = rs("email_2")
			residencia_endereco = rs("residencia_endereco")
			residencia_bairro = rs("residencia_bairro")
			residencia_cep = rs("residencia_cep")
			residencia_cidade = rs("residencia_cidade")
			residencia_estado = rs("residencia_estado")
			comercial_endereco = rs("comercial_endereco")
			comercial_bairro = rs("comercial_bairro")
			comercial_cep = rs("comercial_cep")
			comercial_cidade = rs("comercial_cidade")
			comercial_estado = rs("comercial_estado")
			comercial_telefone = rs("comercial_telefone")
			pais_endereco = rs("pais_endereco")
			pais_bairro = rs("pais_bairro")
			pais_cep = rs("pais_cep")
			pais_cidade = rs("pais_cidade")
			pais_estado = rs("pais_estado")
			pais_telefone = rs("pais_telefone")
			pais_celular = rs("pais_celular")
			retirou_convites = rs("retirou_convites")
			data_retirou_convites = rs("data_retirou_convites")

			cod_mack_lista_provaveis_formandos_nome = rs("cod_mack_lista_provaveis_formandos_nome")
		Else
			response.write "formando NÃO encontrado"
			response.end
		end if
		rs.Close
		Set rs=Nothing


	Set rs=objConexao.Execute("SELECT cod_mack_temporada FROM TB_MACK_COLACOES WHERE cod_mack_colacao = "&cod_mack_colacao)
	if not rs.eof then
		cod_mack_temporada = rs(0)
	end if
	rs.Close
	Set rs=Nothing
	
	FecharConexao()

End sub

Sub incluirRegistro(redirecionar)

	lerFormulario()

	AbrirConexao()

	dim sql : sql = "INSERT INTO TB_MACK_FORMANDOS ("&_
						"cod_mack_colacao, "&_
						"nome, "&_
						"tia, "&_
						"senha, "&_
						"numero_identificacao, "&_
						"nome_pai, "&_
						"nome_mae, "&_
						"curso, "&_
						"telefone, "&_
						"celular, "&_
						"email_1, "&_
						"email_2, "&_
						"residencia_endereco, "&_
						"residencia_bairro, "&_
						"residencia_cep, "&_
						"residencia_cidade, "&_
						"residencia_estado, "&_
						"comercial_endereco, "&_
						"comercial_bairro, "&_
						"comercial_cep, "&_
						"comercial_cidade, "&_
						"comercial_estado, "&_
						"comercial_telefone, "&_
						"pais_endereco, "&_
						"pais_bairro, "&_
						"pais_cep, "&_
						"pais_cidade, "&_
						"pais_estado, "&_
						"pais_telefone, "&_
						"pais_celular, "&_
						"retirou_convites,"&_ 
						"data_retirou_convites,"&_
						"cod_mack_lista_provaveis_formandos_nome"&_
						") VALUES  ( "&_
						sqlNumero(cod_mack_colacao)&", "&_
						sqlTexto(nome)&", "&_
						sqlTexto(tia)&", "&_
						sqlTexto(senha)&", "&_
						sqlTexto(numero_identificacao)&", "&_
						sqlTexto(nome_pai)&", "&_
						sqlTexto(nome_mae)&", "&_
						sqlTexto(curso)&", "&_
						sqlTexto(telefone)&", "&_
						sqlTexto(celular)&", "&_
						sqlTexto(email_1)&", "&_
						sqlTexto(email_2)&", "&_
						sqlTexto(residencia_endereco)&", "&_
						sqlTexto(residencia_bairro)&", "&_
						sqlTexto(residencia_cep)&", "&_
						sqlTexto(residencia_cidade)&", "&_
						sqlTexto(residencia_estado)&", "&_
						sqlTexto(comercial_endereco)&", "&_
						sqlTexto(comercial_bairro)&", "&_
						sqlTexto(comercial_cep)&", "&_
						sqlTexto(comercial_cidade)&", "&_
						sqlTexto(comercial_estado)&", "&_
						sqlTexto(comercial_telefone)&", "&_
						sqlTexto(pais_endereco)&", "&_
						sqlTexto(pais_bairro)&", "&_
						sqlTexto(pais_cep)&", "&_
						sqlTexto(pais_cidade)&", "&_
						sqlTexto(pais_estado)&", "&_
						sqlTexto(pais_telefone)&", "&_
						sqlTexto(pais_celular)& ","&_
						sqlTexto(retirou_convites)&", "&_
						sqlData(data_retirou_convites)&", "&_
						sqlNumero(0&cod_mack_lista_provaveis_formandos_nome)&" "&_
						")"

	if retirou_convites = "S" then
		data_retirou_convites = NOW()
	else
		data_retirou_convites = null
	end if

	response.write(sql)

	objConexao.Execute(sql)

	'Encontrar chave criada
	Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_MACK_FORMANDOS LIMIT 1")
		If not rs.Eof Then
			cod_mack_formando = rs("ULTIMO_ID")
		End if 
	rs.Close
	Set rs=Nothing

	FecharConexao()
	Response.Redirect(redirecionar&"?comando=alterar&cod_mack_formando="&cod_mack_formando&"&codmsg=1&guia=2")
End sub

Sub alterarRegistro(redirecionar)
	
	lerFormulario()

	AbrirConexao()

	'verificamos se ja havia sido retirado	
	Set rsR = objConexao.execute("SELECT retirou_convites from TB_MACK_FORMANDOS WHERE retirou_convites = 'S' AND cod_mack_formando = "&sqlNumero(cod_mack_formando))
	if rsR.Eof and retirou_convites = "S" then
		data_retirou_convites = NOW()
	else
		data_retirou_convites = null
	end if
	rsR.Close
	set rsR = nothing

	dim sql : sql = "UPDATE TB_MACK_FORMANDOS SET"&_
						 " cod_mack_colacao = "&sqlNumero(cod_mack_colacao)&_
						 ", nome = "&sqltexto(nome)&_
						 ", tia="&sqltexto(tia)&_
						 ", curso = "&sqltexto(curso)&_
						 ", senha = "&sqltexto(senha)&_
						 ", numero_identificacao = "&sqltexto(numero_identificacao)&_
						 ", nome_pai = "&sqltexto(nome_pai)&_
						 ", nome_mae  = "&sqltexto(nome_mae)&_
						 ", telefone = "&sqltexto(telefone)&_
						 ", celular = "&sqltexto(celular)&_
						 ", email_1 = "&sqltexto(email_1)&_
						 ", email_2 = "&sqltexto(email_2)&_
						 ", residencia_endereco = "&sqltexto(residencia_endereco)&_
						 ", residencia_bairro = "&sqltexto(residencia_bairro)&_
						 ", residencia_cep = "&sqltexto(residencia_cep)&_
						 ", residencia_cidade = "&sqltexto(residencia_cidade)&_
						 ", residencia_estado = "&sqltexto(residencia_estado)&_
						 ", comercial_endereco = "&sqltexto(comercial_endereco)&_
						 ", comercial_bairro = "&sqltexto(comercial_bairro)&_
						 ", comercial_cep = "&sqltexto(comercial_cep)&_
						 ", comercial_cidade = "&sqltexto(comercial_cidade)&_
						 ", comercial_estado = "&sqltexto(comercial_estado)&_
						 ", comercial_telefone = "&sqltexto(comercial_telefone)&_
						 ", pais_endereco = "&sqltexto(pais_endereco)&_
						 ", pais_bairro = "&sqltexto(pais_bairro)&_
						 ", pais_cep = "&sqltexto(pais_cep)&_
						 ", pais_cidade = "&sqltexto(pais_cidade)&_
						 ", pais_estado = "&sqltexto(pais_estado)&_
						 ", pais_telefone = "&sqltexto(pais_telefone)&_
						 ", pais_celular = "&sqltexto(pais_celular)&_
						 ", retirou_convites = "&sqlTexto(retirou_convites)&_
						 ", data_retirou_convites = "&sqlData(data_retirou_convites)&_
						 " WHERE cod_mack_formando = "&sqlNumero(cod_mack_formando)

	response.write(sql)

	objConexao.execute(sql)

	FecharConexao()

	Response.Redirect(redirecionar&"?cod_mack_formando="&cod_mack_formando&"&codmsg=2")
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		if  estoqueVerificaManual(cod_mack_formando, true) then
			objConexao.Execute("DELETE FROM TB_MACK_FORMANDOS WHERE cod_mack_formando = "&cod_mack_formando)
		end if
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_mack_formando="&cod_mack_formando&"&codmsg=3")
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
<% Dim qtd_tabs : qtd_tabs = 3 %>
<div id="tabsJ">
	<ul>
		<li id="liTab1" class="current">
			<a href="javascript:tabsMostra('tab1',<%=qtd_tabs%>,1);">
				<span>Cadastro</span>
			</a>
		</li>
		<li id="liTab2">
			<a href="javascript:tabsMostra('tab2',<%=qtd_tabs%>,2);">
				<span>Foto Identificação</span>
			</a>
		</li>
    <li id="liTab3">
			<a href="javascript:tabsMostra('tab3',<%=qtd_tabs%>,3);">
				<span>Informação para Venda</span>
			</a>
		</li>
	</ul>
</div>

<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post'>
<div id="tabsCaixa">
	<div id="tab1" >
		<br><br>
		<!--Dados Gerais-->
		<fieldset>
			<legend>Dados Gerais</legend>
			<table width='800px' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
				<%If comando <> "incluir" Then %><tr>
				<td width='27%'><strong>Cód. Formando</strong></td>
					<td width='73%'>
						<div align='left'>
						<% If comando <> "visualizar" and comando <> "excluir"  Then%>
							<span class='textoMaior'><strong><%=cod_mack_formando%></strong></span><%Else%>
							<%
							If not isNull(cod_mack_formando) Then
								response.write(cod_mack_formando)
							End If
						End If %>
						</div>
					</td>
				</tr>
				<%end if%>
				<tr>
					<td width='27%'><strong>Nome *</strong></td>
					<td width='73%'>
						<div align='left'>
							<%If comando <> "visualizar" and comando <> "excluir"  Then%>
								<input name='nome' value="<%=nome%>" class='inputTamanho2' type='text' maxlength='255' />
							<% Else 
								If not isNull(nome) Then
									response.write(nome)
								End If
							End If %>
						</div>
					</td>
				</tr>
				<tr>
					<td width='27%'><strong>Tia *</strong></td>
					<td width='73%'>
						<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='tia' id='id_tia' value="<%=tia%>" class='inputTamanho2 padrao_numero' type='text' maxlength='50' /> <span id='msg_aguarde'></span>
						<% Else 
							If not isNull(tia) Then
								response.write(tia)
							End If
						End If %>

						</div>
					</td>
				</tr>
				<tr>
					<td width='27%'><strong>Curso *</strong></td>
					<td width='73%'>
						<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='curso' class='inputTamanho2' value="<%=curso%>" type='text' />
						<%
						Else
							If not isNull(curso) Then
								response.write(curso)
							End If
						End If %>
						</div>
					</td>
				</tr>
				<tr>
					<td width='27%'><strong>Temporada *</strong></td>
					<td width='73%'>
						<div align='left'>
							<%=padraoFormularioSelect("cod_mack_temporada","id_cod_mack_temporada","inputTamanho2","","","Selecione", cod_mack_temporada,"SELECT cod_mack_temporada, concat(semestre_conclusao,'/',ano_conclusao) FROM TB_MACK_TEMPORADAS order by ano_conclusao, semestre_conclusao desc", "incluir", false, false, false)%>
						</div>
					</td>
				</tr>
				<tr>
					<td width='27%'><strong>Colação *</strong></td>
					<td width='73%'>
						<div align='left' id='selectCol'>
							<% if comando = "visualizar" and comando = "excluir" then 
								if trim(""&cod_mack_temporada) = "" then
									cod_mack_temporada = mackenzieTemporadaAtiva(false)
								end if

							

								response.write(padraoFormularioSelect("cod_mack_colacao", "", "inputTamanho3",""," required='required' ","Selecione", cod_mack_colacao,"SELECT cod_mack_colacao, concat(descricao ,' - ',cast(DATE_FORMAT(data_colacao,'%d/%m/%Y') as char)) FROM TB_MACK_COLACOES WHERE cod_mack_temporada = "&cod_mack_temporada&" ORDER BY descricao",comando, false, false, false) )
							else
								response.write(padraoFormularioSelect("cod_mack_colacao", "", "inputTamanho3",""," required='required' ","Selecione", cod_mack_colacao,"SELECT cod_mack_colacao, concat(descricao ,' - ',cast(DATE_FORMAT(data_colacao,'%d/%m/%Y') as char)) FROM TB_MACK_COLACOES ORDER BY descricao",comando, false, false, false))
							end if %>
						</div>
					</td>
				</tr>
				<tr>
					<td width='27%'><strong>Número de identificação da ficha *</strong></td>
					<td width='73%'>
						<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='numero_identificacao' class='inputTamanho2' value="<%=numero_identificacao%>" type='text' />
						<%
						Else
							If not isNull(numero_identificacao) Then
								response.write(numero_identificacao)
							End If
						End If %>
						</div>
					</td>
				</tr>

				<tr>
					<td width='27%'><strong>Formando já retirou os convites? *</strong></td>
					<td width='73%'>
						<div align='left'>
						<%=padraoFormularioSelect("retirou_convites", "", "", "", "", "SEM_SELECIONE", retirou_convites, "PADRAO:N#NÃO|S#Sim", comando, false, false, false)%>
						<em>
						<% 
						if retirou_convites = "S" then
							response.write(" Retirou em:"&arrumaData(data_retirou_convites))
						end if
						%>
						</em>

						</div>
					</td>
				</tr>
			</table>	
		</fieldset>
		<!--Informa��es para Contato-->
		<fieldset>
			<legend>Informações para Contato</legend>
			<table width='800px' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
			<tr>
				<td width='27%'><strong>Telefone</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='telefone' value="<%=telefone%>" class='inputTamanho2 padrao_telefone' type='text' maxlength='255' />
						<% Else 
							If not isNull(telefone) Then
								response.write(telefone)
							End If
						End If %>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>Celular</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='celular' value="<%=celular%>" class='inputTamanho2 padrao_celular' type='text' maxlength='255' />
						<% Else 
							If not isNull(celular) Then
								response.write(celular)
							End If
						End If %>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>E-mail</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='email_1' value="<%=email_1%>" class='inputTamanho2' type='text' maxlength='255' />
						<% Else 
							If not isNull(email_1) Then
								response.write(email_1)
							End If
						End If %>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>E-mail alternativo</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='email_2' value="<%=email_2%>" class='inputTamanho2' type='text' maxlength='255' />
						<% Else 
							If not isNull(email_2) Then
								response.write(email_2)
							End If
						End If %>
					</div>
				</td>
			</tr>
			</table>
		</fieldset>
		<!--Endere�o Residencial-->
		<fieldset>
			<legend>Endereço Residencial</legend>
			<table width='800px' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
			<tr>
				<td width='27%'><strong>Endereço</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='residencia_endereco' value="<%=residencia_endereco%>" class='inputTamanho2' type='text' maxlength='255' />
						<% Else 
							If not isNull(residencia_endereco) Then
								response.write(residencia_endereco)
							End If
						End If %>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>Bairro</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='residencia_bairro' value="<%=residencia_bairro%>" class='inputTamanho2' type='text' maxlength='255' />
						<% Else 
							If not isNull(residencia_bairro) Then
								response.write(residencia_bairro)
							End If
						End If %>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>CEP</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='residencia_cep' value="<%=residencia_cep%>" class='inputTamanho2 padrao_cep' type='text' maxlength='255' />
						<% Else 
							If not isNull(residencia_cep) Then
								response.write(residencia_cep)
							End If
						End If %>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>Cidade</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='residencia_cidade' value="<%=residencia_cidade%>" class='inputTamanho2' type='text' maxlength='255' />
						<% Else 
							If not isNull(residencia_cidade) Then
								response.write(residencia_cidade)
							End If
						End If %>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>Estado</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='residencia_estado' value="<%=residencia_estado%>" class='inputTamanho2' type='text' maxlength='2' />
						<% Else 
							If not isNull(residencia_estado) Then
								response.write(residencia_estado)
							End If
						End If %>
					</div>
				</td>
			</tr>
			</table>
		</fieldset>
		<!--Endere�o Comercial-->
		<fieldset>
			<legend>Endereço Comercial</legend>
			<table width='800px' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
			<tr>
				<td width='27%'><strong>Endereço</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='comercial_endereco' value="<%=comercial_endereco%>" class='inputTamanho2' type='text' maxlength='255' />
						<% Else 
							If not isNull(comercial_endereco) Then
								response.write(comercial_endereco)
							End If
						End If %>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>Bairro</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='comercial_bairro' value="<%=comercial_bairro%>" class='inputTamanho2' type='text' maxlength='255' />
						<% Else 
							If not isNull(comercial_bairro) Then
								response.write(comercial_bairro)
							End If
						End If %>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>CEP</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='comercial_cep' value="<%=comercial_cep%>" class='inputTamanho2 padrao_cep' type='text' maxlength='255' />
						<% Else 
							If not isNull(comercial_cep) Then
								response.write(comercial_cep)
							End If
						End If %>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>Cidade</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='comercial_cidade' value="<%=comercial_cidade%>" class='inputTamanho2' type='text' maxlength='255' />
						<% Else 
							If not isNull(comercial_cidade) Then
								response.write(comercial_cidade)
							End If
						End If %>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>Estado</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='comercial_estado' value="<%=comercial_estado%>" class='inputTamanho2' type='text' maxlength='2' />
						<% Else 
							If not isNull(comercial_estado) Then
								response.write(comercial_estado)
							End If
						End If %>
					</div>
				</td>
			</tr>
			</table>
		</fieldset>
		<!--Informa��es dos Pais-->
		<fieldset>
			<legend>Informações dos Pais</legend>
			<table width='800px' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
			<tr>
				<td width='27%'><strong>Nome Pai</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='nome_pai' value="<%=nome_pai%>" class='inputTamanho2' type='text' maxlength='255' />
						<% Else 
							If not isNull(nome_pai) Then
								response.write(nome_pai)
							End If
						End If %>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>Nome Mãe</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='nome_mae' value="<%=nome_mae%>" class='inputTamanho2' type='text' maxlength='255' />
						<% Else 
							If not isNull(nome_mae) Then
								response.write(nome_mae)
							End If
						End If %>
					</div>
				</td>
			</tr>

			<tr>
				<td width='27%'><strong>Telefone</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='pais_telefone' value="<%=pais_telefone%>" class='inputTamanho2 padrao_celular' type='text' maxlength='255' />
						<% Else 
							If not isNull(pais_telefone) Then
								response.write(pais_telefone)
							End If
						End If %>
					</div>
				</td>
			</tr>
				<tr>
				<td width='27%'><strong>Celular</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='pais_celular' value="<%=pais_celular%>" class='inputTamanho2 padrao_celular' type='text' maxlength='255' />
						<% Else 
							If not isNull(pais_celular) Then
								response.write(pais_celular)
							End If
						End If %>
					</div>
				</td>
			</tr>
				<tr>
				<td width='27%'><strong>Endereço</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='pais_endereco' value="<%=pais_endereco%>" class='inputTamanho2' type='text' maxlength='255' />
						<% Else 
							If not isNull(pais_endereco) Then
								response.write(pais_endereco)
							End If
						End If %>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>Bairro</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='pais_bairro' value="<%=pais_bairro%>" class='inputTamanho2' type='text' maxlength='255' />
						<% Else 
							If not isNull(pais_bairro) Then
								response.write(pais_bairro)
							End If
						End If %>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>CEP</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='pais_cep' value="<%=pais_cep%>" class='inputTamanho2 padrao_cep' type='text' maxlength='255' />
						<% Else 
							If not isNull(pais_cep) Then
								response.write(pais_cep)
							End If
						End If %>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>Cidade</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='pais_cidade' value="<%=pais_cidade%>" class='inputTamanho2' type='text' maxlength='255' />
						<% Else 
							If not isNull(pais_cidade) Then
								response.write(pais_cidade)
							End If
						End If %>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>Estado</strong></td>
				<td width='73%'>
					<div align='left'>
						<%If comando <> "visualizar" and comando <> "excluir"  Then%>
							<input name='pais_estado' value="<%=pais_estado%>" class='inputTamanho2' type='text' maxlength='2' />
						<% Else 
							If not isNull(pais_estado) Then
								response.write(pais_estado)
							End If
						End If %>
					</div>
				</td>
			</tr>
			</table>
		</fieldset>

		<%If comando <> "visualizar" Then%>
			<table width='800px' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
			<tr>
				<td colspan='2' align='center'><div align='left'><strong>* - Campos obrigatórios</strong><br /></div></td>
			</tr>
			<tr>
				<td colspan='2' align='center'><a href="javascript:salvar();"><img src="images/layout/btnGravar.jpg" /></a></td>
			</tr>
			</table>
		<%End If%>		
		
		<input type='hidden' value='1' name='campo1'  />
		<input type='hidden' name='cod_mack_formando' value="<%=cod_mack_formando%>" />
		<input type='hidden' name='cod_mack_lista_provaveis_formandos_nome' value="<%=cod_mack_lista_provaveis_formandos_nome%>" />
	</div>

	<div id="tab2" class="tabs" ><br><br><br>
		<% if comando = "incluir" then %>
			<div class="msgGrande">Grave o regitro antes de incluir uma imagem.</div>
			<br><br>
		<% else %>
			<iframe  src="mck_formandos_formulario_upload.asp?cod_mack_formando=<%=cod_mack_formando%>&comando_externo=<%=comando%>" width="100%" height="400px;" border='0' style="border: 0px;"></iframe>
		<%end if %>
	</div>

  <div id="tab3" class="tabs" ><br><br><br>
		<% if comando = "incluir" then %>
			<div class="msgGrande">Grave o regitro antes de incluir uma venda.</div>
			<br><br>
		<% else %>
			<iframe  src="mck_formandos_formulario_converter.asp?cod_mack_formando=<%=cod_mack_formando%>&comando_externo=<%=comando%>" width="100%" height="400px;" border='0' style="border: 0px;"></iframe>
		<%end if %>
	</div>
</div>
</form>
<% if Request.QueryString("guia") then %>
<script type="text/javascript">
	$(function(){
		tabsMostra('tab2',<%=qtd_tabs%>,2);
	});
</script>

<% end if %>
<script type="text/javascript">
	$(function(){
		$('#id_tia').focusout(function(){
			if ($('#id_tia').val() != "")
			{
				var request = $.ajax({
				  url: "mck_formandos_formulario_ajax.asp",
				  method: "POST",
				  data: { t : "validaTia" , "p0" : $('#id_tia').val()  },
				  dataType: "html",
				  beforeSend: function(){
				  	$('input').attr('disabled', 'disabled');
				  	$('select').attr('disabled', 'disabled');
				  	$("#msg_aguarde").text('Aguarde...');
				  }
				});
				 
				request.done(function( msg ) {
				  $('input').removeAttr('disabled');
				  $('select').removeAttr('disabled');
				  if(msg == 'false')
				  {
				  	$("#msg_aguarde").text('');
				  	// $('#id_tia').focus();
				  	alert('O  TIA ('+$('#id_tia').val()+') já está cadastrado para outro formando');
				  }
				  else {
				  	$("#msg_aguarde").text('');	
				  }
				});
				 
				request.fail(function( jqXHR, textStatus ) {
				  console.log( "Request failed: " + textStatus );
				  console.log(  jqXHR)
				});
			}
		});

		$('#id_cod_mack_temporada').change(function(){
			var request = $.ajax({
			  url: "mck_formandos_formulario_ajax.asp",
			  method: "POST",
			  data: { t : "selectColacoes" , "p0" : $('#id_cod_mack_temporada').val(), p1 : 1  },
			  dataType: "html",
			  beforeSend: function(){
			  	$('#id_cod_mack_temporada').attr('disabled', 'disabled');
			  	$( "#selectCol" ).html( "Aguarde..." );
			  }
			});
			 
			request.done(function( msg ) {
			  $('#id_cod_mack_temporada').removeAttr('disabled');
			  $( "#selectCol" ).html( msg );

			});
			 
			request.fail(function( jqXHR, textStatus ) {
			  console.log( "Request failed: " + textStatus );
			  console.log(  jqXHR)
			});
		});
	});
</script>

<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('nome#Nome#1|tia#Tia#1|curso#Curso#1|COD_MACK_COLACAO#Colação#2|numero_identificacao#Número de identificação da ficha#1|RETIROU_CONVITES#Formando já retirou os convites?#2',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>
<!-- #include file="files/includes/miolo-inferior.asp" -->