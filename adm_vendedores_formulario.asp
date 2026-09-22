<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 1
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
  Dim titulo_pagina
  titulo_pagina = "Vendedores"


Dim calendario_necessario
calendario_necessario = "N&atilde;O"
Set Upload = New FreeASPUpload
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<!-- #include file="includes/freeASPUpload.asp" -->
<%

'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 26/7/2011 17:51:35

'Declara vari�veis para utiliza��o do sistema
Dim cod_vendedor, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario, endereco, obs

redirecionar_incluir	= "adm_usuarios.asp"
redirecionar_alterar	= "adm_usuarios.asp"
redirecionar_excluir	= "adm_usuarios.asp"
redirecionar_validacao	= "adm_usuarios.asp"
msgAoExcluir			= "Tem certeza que deseja excluir este vendedor?<br /><br />Voc� poder� inativa-lo ao inv�s de exclui-lo. Se houver Pedidos com este vendedor, o mesmo n�o ser� exclu�do"
actionformulario		= caminho_pagina


comando = Request.QueryString("comando")
campo1 =  Upload.Form("campo1")

'Declara vari�veis
Dim NOME_VENDEDOR, LOGIN, SENHA, E_MAIL, ATIVADO, COD_GRUPO_VENDEDORES, cod_tabela_comissao_vendedores, CPF,RG,NOME_PAI,NOME_MAE,DATA_NASCIMENTO, CPF_ARQUIVO, RG_ARQUIVO

If comando = Empty Then
comando = "incluir"
titulo_pagina = titulo_pagina&" / Inclus&atilde;o"
End if




If comando = "incluir" and campo1 <> Empty Then
'Incluir novo registro
incluirRegistro(redirecionar_incluir)
titulo_pagina = titulo_pagina&" / Inclus&atilde;o"
Elseif comando = "alterar" and campo1 = Empty Then
'Leitura inicial para poss�vel altera��o
cod_vendedor = Request.QueryString("cod_vendedor")
lerRegistro()
titulo_pagina = titulo_pagina&" / Altera&ccedil;&atilde;o"
Elseif 	comando = "alterar" and campo1 <> Empty Then
'Gravar altera��o
cod_vendedor = Upload.Form("cod_vendedor")

alterarRegistro(redirecionar_alterar)
titulo_pagina = titulo_pagina&" / Altera&ccedil;&atilde;o"
Elseif 	comando = "excluir" and campo1 = Empty Then 
'Leitura inicial para poss�vel exclus�o
cod_vendedor = Request.QueryString("cod_vendedor")
lerRegistro()
titulo_pagina = titulo_pagina&" / Exclus&atilde;o"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
'Exclui registro
cod_vendedor = Upload.Form("cod_vendedor")
excluirRegistro(redirecionar_excluir)
titulo_pagina = titulo_pagina&" / Exclus&atilde;o"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
'Visualiza registro
cod_vendedor = Request.QueryString("cod_vendedor")
lerRegistro()
titulo_pagina = titulo_pagina&" / Visualiza&ccedil;&atilde;o"
End If

if Upload.UploadedFiles.Exists("RG_ARQUIVO") then
		response.end
	end if

'Fun��es

Sub lerFormulario()



	COD_VENDEDOR = Upload.Form("COD_VENDEDOR")
	NOME_VENDEDOR = Ucase(Upload.Form("NOME_VENDEDOR"))
	LOGIN = Upload.Form("LOGIN")
	SENHA = Upload.Form("SENHA")
	E_MAIL = Upload.Form("E_MAIL")
	ATIVADO = Upload.Form("ATIVADO")
	COD_GRUPO_VENDEDORES = Upload.Form("COD_GRUPO_VENDEDORES")
	cod_tabela_comissao_vendedores = Upload.Form("cod_tabela_comissao_vendedores")
	endereco = Upload.Form("endereco")
	obs = Upload.form("obs")
	CPF = Upload.form("CPF")
	RG = Upload.form("RG")
	NOME_PAI = Upload.form("NOME_PAI")
	NOME_MAE = Upload.form("NOME_MAE")
	DATA_NASCIMENTO = Upload.form("DATA_NASCIMENTO")

End sub

Sub lerRegistro()
	AbrirConexao()

		Set rs=objConexao.Execute("SELECT * FROM TB_VENDEDORES WHERE cod_vendedor = "&cod_vendedor)

			NOME_VENDEDOR=trim(rs("NOME_VENDEDOR"))
			LOGIN=trim(rs("LOGIN"))
			SENHA=trim(rs("SENHA"))
			E_MAIL=trim(rs("E_MAIL"))
			ATIVADO=trim(rs("ATIVADO"))
			COD_GRUPO_VENDEDORES=trim(rs("COD_GRUPO_VENDEDORES"))
			cod_tabela_comissao_vendedores=trim(rs("cod_tabela_comissao_vendedores"))
			endereco = trim(rs("endereco"))
			obs = trim(rs("obs"))
			CPF = trim(rs("CPF"))
			RG = trim(rs("RG"))
			NOME_PAI = trim(rs("NOME_PAI"))
			NOME_MAE = trim(rs("NOME_MAE"))
			DATA_NASCIMENTO = trim(rs("DATA_NASCIMENTO"))

			RG_ARQUIVO = rs("ARQUIVO_RG")
			CPF_ARQUIVO = rs("ARQUIVO_CPF")

		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)
	lerFormulario()

		Dim chavesUnicasDuplicadas,camposUnicos
		call chaveUnica ("TB_VENDEDORES","LOGIN",LOGIN,redirecionar&"?cod_vendedor="&cod_vendedor&"&codmsg=4&sucesso=2&chaveUnica=",comando,"cod_vendedor",cod_vendedor)

	AbrirConexao()
		if cod_tabela_comissao_vendedores = 0 or cod_tabela_comissao_vendedores = "" then
			cod_tabela_comissao_vendedores = null
		end if 

		sql = "INSERT INTO TB_VENDEDORES (NOME_VENDEDOR, LOGIN, SENHA, E_MAIL, ATIVADO, COD_GRUPO_VENDEDORES, cod_tabela_comissao_vendedores, endereco, obs, CPF,RG,NOME_PAI,NOME_MAE,DATA_NASCIMENTO) VALUES ("&sqlTexto(NOME_VENDEDOR)&", "&sqlTexto(LOGIN)&", "&sqlTexto(SENHA)&", "&sqlTexto(E_MAIL)&", "&sqlTexto(ATIVADO)&", "&sqlNumero(COD_GRUPO_VENDEDORES)&","&cod_tabela_comissao_vendedores&", "&sqlTexto(endereco)&", "&sqlTexto(obs)&", "&sqlTexto(CPF)&","&sqltexto(RG)&","&sqlTexto(NOME_PAI)&","&sqltexto(NOME_MAE)&","&sqldata(DATA_NASCIMENTO)&")"

		objConexao.Execute(sql)

		'Encontrar chave criada
		Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_VENDEDORES LIMIT 1")
			If not rs.Eof Then
				cod_vendedor = rs("ULTIMO_ID")
			End if 
		rs.Close
		Set rs=Nothing
Set Upload = New FreeASPUpload

		'Buscamos todos os arquivos e enviamos
		if Upload.UploadedFiles.Exists("RG_ARQUIVO")  then
			
			'tratamos o nome do arquivo
			nome_arquivo_original = Upload.UploadedFiles.item("arquivo").fileName
			inicio_ext = instrrev(nome_arquivo_original,".")
			len_arquivo =  len(nome_arquivo_original)
			extencao = mid(nome_arquivo_original,inicio_ext+1,(len_arquivo-inicio_ext))
			nome_arquivo_original = replace(nome_arquivo_original, "."&extencao,"")

			'Tratamos o nome do arquivo
			nome_arquivo_servidor = cod_vendedor&"_rg"
			

			' extencao = lcase(replace(right(Upload.UploadedFiles.item("arquivo")).fileName,4),".","")
			if  extencao = "png" or extencao = "jpg" or extencao = "jpeg" then
				'SALVAMOS O ARQUIVO 
				call Upload.SaveAs("arquivo", nome_arquivo_servidor&"."&extencao, caminho_dados&"vendedores\")

				objConexao.Execute("UPDATE TB_VENDEDORES SET ARQUIVO_RG = "&sqltexto(nome_arquivo_servidor&"."&extencao)&" WHERE cod_vendedor = "&sqlNumero(cod_vendedor))
			end if
		end if

		if Upload.UploadedFiles.Exists("CPF_ARQUIVO")  then
		
			'tratamos o nome do arquivo
			nome_arquivo_original = Upload.UploadedFiles.item("arquivo").fileName
			inicio_ext = instrrev(nome_arquivo_original,".")
			len_arquivo =  len(nome_arquivo_original)
			extencao = mid(nome_arquivo_original,inicio_ext+1,(len_arquivo-inicio_ext))
			nome_arquivo_original = replace(nome_arquivo_original, "."&extencao,"")

			'Tratamos o nome do arquivo
			nome_arquivo_servidor = cod_vendedor&"_rg"
			
			' extencao = lcase(replace(right(Upload.UploadedFiles.item("arquivo")).fileName,4),".","")
			if  extencao = "png" or extencao = "jpg" or extencao = "jpeg" then
				'SALVAMOS O ARQUIVO 
				call Upload.SaveAs("arquivo", nome_arquivo_servidor&"."&extencao, caminho_dados&"vendedores\")

				objConexao.Execute("UPDATE TB_VENDEDORES SET ARQUIVO_CPF = "&sqltexto(nome_arquivo_servidor&"."&extencao)&" WHERE cod_vendedor = "&sqlNumero(cod_vendedor))
			end if
		end if

	FecharConexao()
	Response.Redirect(redirecionar&"?cod_vendedor="&cod_vendedor&"&codmsg=1")
End sub

Sub alterarRegistro(redirecionar)
	lerRegistro()
	lerFormulario()
		Dim chavesUnicasDuplicadas,camposUnicos
		call chaveUnica("TB_VENDEDORES","LOGIN",LOGIN,redirecionar&"?cod_vendedor="&cod_vendedor&"&codmsg=4&sucesso=2&chaveUnica=",comando,"cod_vendedor",cod_vendedor)

		if cod_tabela_comissao_vendedores = "0" or cod_tabela_comissao_vendedores = "" then
			cod_tabela_comissao_vendedores = "null"
		end if 

		AbrirConexao()
			objConexao.Execute("UPDATE TB_VENDEDORES SET COD_VENDEDOR = "&sqlNumero(COD_VENDEDOR)&", NOME_VENDEDOR = "&sqlTexto(NOME_VENDEDOR)&", LOGIN = "&sqlTexto(LOGIN)&", SENHA = "&sqlTexto(SENHA)&", E_MAIL = "&sqlTexto(E_MAIL)&", ATIVADO = "&sqlTexto(ATIVADO)&", COD_GRUPO_VENDEDORES = "&sqlNumero(COD_GRUPO_VENDEDORES)&", cod_tabela_comissao_vendedores = "&cod_tabela_comissao_vendedores&", endereco = "&sqlTexto(endereco)&", obs = "&sqlTexto(obs)&", CPF = "&sqlTexto(CPF)&" , RG = "&sqlTexto(RG)&", NOME_PAI = "&sqlTexto(NOME_PAI)&" , NOME_MAE = "&sqlTexto(NOME_MAE)&", DATA_NASCIMENTO = "&sqldata(DATA_NASCIMENTO)&" WHERE cod_vendedor = "&cod_vendedor)

		'Buscamos todos os arquivos e enviamos
		if Upload.UploadedFiles.Exists("rgarquivo")  then
			
			'tratamos o nome do arquivo
			nome_arquivo_original = Upload.UploadedFiles.item("rgarquivo").fileName
			inicio_ext = instrrev(nome_arquivo_original,".")
			len_arquivo =  len(nome_arquivo_original)
			extencao = mid(nome_arquivo_original,inicio_ext+1,(len_arquivo-inicio_ext))
			nome_arquivo_original = replace(nome_arquivo_original, "."&extencao,"")

			'Tratamos o nome do arquivo
			nome_arquivo_servidor = cod_vendedor&"_rg"


			' extencao = lcase(replace(right(Upload.UploadedFiles.item("arquivo")).fileName,4),".","")
			if  extencao = "png" or extencao = "jpg" or extencao = "jpeg" then
				'SALVAMOS O ARQUIVO 
				call deletarArquivo(  caminho_dados&"vendedores\"&RG_ARQUIVO)
				call Upload.SaveAs("rgarquivo", nome_arquivo_servidor&"."&extencao, caminho_dados&"vendedores\")

				objConexao.Execute("UPDATE TB_VENDEDORES SET ARQUIVO_RG = "&sqltexto(nome_arquivo_servidor&"."&extencao)&" WHERE cod_vendedor = "&sqlNumero(cod_vendedor))
			end if
		end if

		if Upload.UploadedFiles.Exists("cpfarquivo")  then
		
			'tratamos o nome do arquivo
			nome_arquivo_original = Upload.UploadedFiles.item("cpfarquivo").fileName
			inicio_ext = instrrev(nome_arquivo_original,".")
			len_arquivo =  len(nome_arquivo_original)
			extencao = mid(nome_arquivo_original,inicio_ext+1,(len_arquivo-inicio_ext))
			nome_arquivo_original = replace(nome_arquivo_original, "."&extencao,"")

			'Tratamos o nome do arquivo
			nome_arquivo_servidor = cod_vendedor&"_cpf"
			
			' extencao = lcase(replace(right(Upload.UploadedFiles.item("arquivo")).fileName,4),".","")
			if  extencao = "png" or extencao = "jpg" or extencao = "jpeg" then
				'SALVAMOS O ARQUIVO 
				
				call deletarArquivo(  caminho_dados&"vendedores\"&CPF_ARQUIVO)
				call Upload.SaveAs("cpfarquivo", nome_arquivo_servidor&"."&extencao, caminho_dados&"vendedores\")

				objConexao.Execute("UPDATE TB_VENDEDORES SET ARQUIVO_CPF = "&sqltexto(nome_arquivo_servidor&"."&extencao)&" WHERE cod_vendedor = "&sqlNumero(cod_vendedor))
			end if
		end if

		FecharConexao()
	Response.Redirect(redirecionar&"?cod_vendedor="&cod_vendedor&"&codmsg=2")
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()
		
		call verificaAntesExcluir(cod_vendedor,"TB_PEDIDOS","cod_vendedor",redirecionar_excluir&"?sucesso=2&codmsg=5")

	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_VENDEDORES WHERE cod_vendedor = "&cod_vendedor)
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_vendedor="&cod_vendedor&"&codmsg=3")
End sub

'----------------------------------------------------------------------------------------

%>
<h2><%=titulo_pagina%></h2><br />
<%If comando = "excluir" Then%>
	<div class='msgExcluir'><%=msgAoExcluir%></div>
<%End If%>

<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' m enctype="multipart/form-data" method='post'>
	<table width='90%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
	<%If comando <> "incluir" Then %>
	<tr>
		<td width='27%'><strong>C&oacute;d. Vendedor </strong></td>
		<td width='73%'><div align='left'>
		<%If comando <> "visualizar" and comando <> "excluir"  Then%>
			<span class='textoMaior'><strong><%=COD_VENDEDOR%></strong></span>
		<%Else%>
			<%If not isNull(COD_VENDEDOR) Then
				response.write(COD_VENDEDOR)
			End If
		End If%>
		</div>
		</td>
	</tr>
	<%End If%>
	<tr>
		<td width='27%'><strong>Nome completo *</strong></td>
		<td width='73%'>
			<div align='left'>
			<%If comando <> "visualizar" and comando <> "excluir"  Then%>
				<input name='NOME_VENDEDOR' value="<%=NOME_VENDEDOR%>" class='inputTamanho3' type='text' maxlength='250' />
			<%Else
				If not isNull(NOME_VENDEDOR) Then
					response.write(NOME_VENDEDOR)
				End If
			End If%>
			</div>
		</td>
	</tr>
	<tr>
		<td width='27%'><strong>Login (sem espa&ccedil;o ou acentos)*</strong></td>
		<td width='73%'>
			<div align='left'>
			<%If comando <> "visualizar" and comando <> "excluir"  Then%>
				<input name='LOGIN' value="<%=LOGIN%>" class='inputTamanho2' type='text' maxlength='60' />
			<%Else
				If not isNull(LOGIN) Then
					response.write(LOGIN)
				End If
			End If%>
			</div>
		</td>
	</tr>
	<tr>
		<td width='27%'><strong>Senha (escolha uma de no m&iacute;nimo 6 caracteres)*</strong></td>
		<td width='73%'>
			<div align='left'>
			<% If comando <> "visualizar" and comando <> "excluir"  Then %>
				<input name='SENHA' value="<%=SENHA%>" class='inputTamanho3' type='text' maxlength='60' />
			<% Else
				If not isNull(SENHA) Then
					response.write(SENHA)
				End If
			End If %>
			</div>
		</td>
	</tr>
	<tr>
		<td width='27%'><strong>E-mail </strong></td>
		<td width='73%'>
			<div align='left'>
			<%If comando <> "visualizar" and comando <> "excluir"  Then%>
				<input name='E_MAIL' value="<%=E_MAIL%>" class='inputTamanho3' type='text' maxlength='250' />
			<%Else
				If not isNull(E_MAIL) Then
					response.write(E_MAIL)
				End If
			End If%>
			</div>
		</td>
	</tr>
	<tr>
		<td><strong>CPF</strong></td>
		<td>
			<%If comando <> "visualizar" and comando <> "excluir"  Then%>
				<input type="text" class='padrao_cpf inputTamanho2' value="<%=CPF%>" name='CPF' />
			<%Else
				If not isNull(CPF) Then
					response.write(CPF)
				End If
			End If%>
		</td>
	</tr>
	<tr>
		<td><strong>CPF Arquivo</strong></td>
		<td>
			<%If comando <> "visualizar" and comando <> "excluir"  Then%>
				
				<input type="file" class='inputTamanho3' name='cpfarquivo' />
			<% End If %>

			<%
				If CPF_ARQUIVO <> Empty Then
					%><a href='adm_vendedores_formulario_download.asp?t=CPF&c=<%=cod_vendedor%>'>Visualizar</a><%
				End If
			%>
		</td>
	</tr>
	<tr>
		<td><strong>RG</strong></td>
		<td>
			<%If comando <> "visualizar" and comando <> "excluir"  Then%>
				<input type="text" class='inputTamanho2' name='RG' value="<%=RG%>" /> 
			<%Else
				If not isNull(RG) Then
					response.write(RG)
				End If
			End If%>
		</td>
	</tr>

	<tr>
		<td><strong>RG Arquivo</strong></td>
		<td>
			<%If comando <> "visualizar" and comando <> "excluir"  Then%>
				<input type="file" class='inputTamanho3' name='rgarquivo' /> 
			<%End If%>

			<% 
				If RG_ARQUIVO <> Empty Then
					%><a href='adm_vendedores_formulario_download.asp?t=RG&c=<%=cod_vendedor%>'>Visualizar</a><%
				End If
			%>
		</td>
	</tr>

	<tr>
		<td><strong>Nome do Pai</strong></td>
		<td>
			<%If comando <> "visualizar" and comando <> "excluir"  Then%>
				<input type="text" class='inputTamanho3' value="<%=NOME_PAI%>" name='NOME_PAI' /> 
			<%Else
				If not isNull(NOME_PAI) Then
					response.write(NOME_PAI)
				End If
			End If%>
		</td>
	</tr>
	<tr>
		<td><strong>Nome do M&atilde;e</strong></td>
		<td>
			<%If comando <> "visualizar" and comando <> "excluir"  Then%>
				<input type="text" class='inputTamanho3' value="<%=NOME_MAE%>" name='NOME_MAE' /> 
			<%Else
				If not isNull(NOME_MAE) Then
					response.write(NOME_MAE)
				End If
			End If%>
		</td>

	</tr>
	<tr>
		<td><strong>Data de nascimento</strong></td>
		<td>
			<%If comando <> "visualizar" and comando <> "excluir"  Then%>
				<input type="text" class='padrao_data inputTamanho1' value="<%=arrumadata(DATA_NASCIMENTO)%>" name='DATA_NASCIMENTO' />
			<%Else
				If not isNull(DATA_NASCIMENTO) Then
					response.write(DATA_NASCIMENTO)
				End If
			End If%>
		</td>
	</tr>
	<tr>
		<td width='27%'><strong>Endere&ccedil;o</strong></td>
		<td width='73%'>
			<div align='left'>
			<%If comando <> "visualizar" and comando <> "excluir"  Then%>
				<input name='ENDERECO' value="<%=ENDERECO%>" class='inputTamanho3' type='text' maxlength='250' />
			<%Else
				If not isNull(ENDERECO) Then
					response.write(ENDERECO)
				End If
			End If%>
			</div>
		</td>
	</tr>
	<tr>
		<td width='27%'><strong>Possui grupo de vendas especial *</strong></td>
		<td width='73%'>
			<div align='left'>
			<%=PadraoFormularioSelect("COD_GRUPO_VENDEDORES","","inputTamanho2","","","SEM_SELECIONE",COD_GRUPO_VENDEDORES,"PADRAO:0#Sem grupo especial|1#Renato|2#Mario|3#Paulo Mantuam|4#Vicente",comando,false,false,false)%>
			</div>
		</td>
	</tr>
	<tr>
		<td width='27%'><strong>Tabela de Comissionamento *</strong></td>
		<td width='73%'>
			<div align='left'>
			<%=PadraoFormularioSelect("COD_TABELA_COMISSAO_VENDEDORES","","inputTamanho3","","","Repasse",COD_TABELA_COMISSAO_VENDEDORES,"SELECT cod_tabela_comissao_vendedores, descricao FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES where COD_TABELA_COMISSAO_VENDEDORES = 1 ORDER BY descricao ASC",comando,false,false,false)%>
			</div>
		</td>
	</tr>
	<tr>
		<td width='27%'><strong>Ativo </strong></td>
		<td width='73%'>
			<div align='left'>
			<%=PadraoFormularioSelect("ATIVADO","","inputTamanho1","","","SEM_SELECIONE",ATIVADO,"PADRAO:1#Sim|2#N�o",comando,false,false,false)%>
			</div>
		</td>
	</tr>
	<tr>
		<td width='27%'><strong>Observa&ccedil;&otilde;es</strong></td>
		<td width='73%'>
			<div align='left'>
			<%If comando <> "visualizar" and comando <> "excluir"  Then%>
			<table style='width:420px;' border='0' cellspacing='0' cellpadding='0'>
				<tr>
					<td width='57'>
						<textarea rows='2' name='OBS' class='inputTamanho3' onkeydown='textCounter(this.form.OBS,this.form.charRestantes_OBS,500);' onkeyup='textCounter(this.form.OBS,this.form.charRestantes_OBS,500);'><%=OBS%></textarea>
					</td>
				</tr>
				<tr>
					<td>
						<table width='50' border='0' align='right' cellpadding='0' cellspacing='0'>
							<tr>
								<td>
									<div align='center'>
										<input class='inputContaCaracteres' tabindex='10000' type='text' name='charRestantes_OBS' value="<%=500- Len(OBS)%>" />
									</div>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<%Else
				If not isNull(OBS) Then
					response.write(OBS)
				End If
			End If%>
			</div>
		</td>
	</tr>

	<%If comando <> "visualizar" Then%>
	<tr>
		<td colspan='2' align='center'><div align='left'><strong>* - Campos obrigat&oacute;rios</strong><br /></div></td>
	</tr>
	<tr>
		<td colspan='2' align='center'><a href="javascript:salvar();"><img src="images/layout/btnGravar.jpg" /></a></td>
	</tr>
	<%End If%>
</table>

<input type="hidden" value="12" name="campo1" />
<input type='hidden' name='COD_VENDEDOR' value="<%=COD_VENDEDOR%>" /></form>
<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
	if(validaForm('NOME_VENDEDOR#Nome#1|LOGIN#Login#1|SENHA#Senha#1|CPF#CPF#1|RG#RG#1|NOME_PAI#Nome do Pai#1|NOME_MAE#Nome da sM�e#1|DATA_NASCIMENTO#Data de nascimento#1|ENDERECO#Endere�o#1|COD_GRUPO_VENDEDORES#Possui grupo de vendas especial#2',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->



