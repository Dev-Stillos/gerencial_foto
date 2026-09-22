<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 2
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
%>
<%
Dim calendario_necessario
calendario_necessario = "NÃO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%

'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 26/6/2012 10:27:21

'Declara vari�veis para utiliza��o do sistema
Dim cod_fornecedor, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

redirecionar_incluir	= "ppv_fornecedores.asp"
redirecionar_alterar	= "ppv_fornecedores.asp"
redirecionar_excluir	= "ppv_fornecedores.asp"
redirecionar_validacao	= "ppv_fornecedores.asp"
msgAoExcluir			= "Tem certeza que deseja excluir este fornecedor?<br /><br />Você poderá inativa-lo ao invés de exclui-lo. Se houver registros com este fornecedor, o mesmo NÃO ser� exclu�do"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
'Declara vari�veis
Dim  cod_fornecedor_stillos, RAZAO_SOCIAL, NOME_FANTASIA, CNPJ, ENDERECO, BAIRRO, CEP, CIDADE, ESTADO, TELEFONE, CONTATO, CARGO, ATIVADO, CADASTRADO_POR

If comando = Empty Then
comando = "incluir"
titulo_pagina = titulo_pagina&" / Inclusão"
End if


'#########################################################################

'#########################################################################

If comando = "incluir" and campo1 <> Empty Then
'Incluir novo registro
incluirRegistro(redirecionar_incluir)
titulo_pagina = titulo_pagina&" / Inclusão"
Elseif comando = "alterar" and campo1 = Empty Then
'Leitura inicial para poss�vel Alteração
cod_fornecedor = Request.QueryString("cod_fornecedor")
lerRegistro()
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "alterar" and campo1 <> Empty Then
'Gravar Alteração
cod_fornecedor = Request.Form("cod_fornecedor")
alterarRegistro(redirecionar_alterar)
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 
'Leitura inicial para possível Exclusão
cod_fornecedor = Request.QueryString("cod_fornecedor")
lerRegistro()
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
'Exclui registro
cod_fornecedor = Request.Form("cod_fornecedor")
excluirRegistro(redirecionar_excluir)
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
'Visualiza registro
cod_fornecedor = Request.QueryString("cod_fornecedor")
lerRegistro()
titulo_pagina = titulo_pagina&" / Visualização"
End If


'Fun��es

Sub lerFormulario()

COD_FORNECEDOR = Request.Form("COD_FORNECEDOR")
cod_fornecedor_stillos = Request.Form("cod_fornecedor_stillos")
RAZAO_SOCIAL = Request.Form("RAZAO_SOCIAL")
NOME_FANTASIA = Request.Form("NOME_FANTASIA")
CNPJ = Request.Form("CNPJ")
ENDERECO = Request.Form("ENDERECO")
BAIRRO = Request.Form("BAIRRO")
CEP = Request.Form("CEP")
CIDADE = Request.Form("CIDADE")
ESTADO = Request.Form("ESTADO")
TELEFONE = padraoTELEFONE(""&Request.Form("TELEFONE"))
CONTATO = Request.Form("CONTATO")
CARGO = Request.Form("CARGO")
ATIVADO = Request.Form("ATIVADO")

End sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM TB_FORNECEDORES WHERE cod_fornecedor = "&cod_fornecedor)

				cod_fornecedor_stillos=trim(rs("cod_fornecedor_stillos"))
				RAZAO_SOCIAL=trim(rs("RAZAO_SOCIAL"))
				NOME_FANTASIA=trim(rs("NOME_FANTASIA"))
				CNPJ= rs("CNPJ")
				ENDERECO=trim(rs("ENDERECO"))
				BAIRRO=trim(rs("BAIRRO"))
				CEP= rs("CEP")
				CIDADE=trim(rs("CIDADE"))
				ESTADO=trim(rs("ESTADO"))
				TELEFONE= padraoTELEFONE(rs("TELEFONE"))
				CONTATO=trim(rs("CONTATO"))
				CARGO=trim(rs("CARGO"))
				ATIVADO=trim(rs("ATIVADO"))
				CADASTRADO_POR=trim(rs("CADASTRADO_POR"))

		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("INSERT INTO TB_FORNECEDORES (cod_fornecedor_stillos, RAZAO_SOCIAL, NOME_FANTASIA, CNPJ, ENDERECO, BAIRRO, CEP, CIDADE, ESTADO, TELEFONE, CONTATO, CARGO, ATIVADO, CADASTRADO_POR) VALUES ("&sqlNumero(cod_fornecedor_stillos)&", "&sqlTexto(RAZAO_SOCIAL)&", "&sqlTexto(NOME_FANTASIA)&", "&sqlTexto(CNPJ)&", "&sqlTexto(ENDERECO)&", "&sqlTexto(BAIRRO)&", "&sqlTexto(CEP)&", "&sqlTexto(CIDADE)&", "&sqlTexto(ESTADO)&", "&sqlTexto(TELEFONE)&", "&sqlTexto(CONTATO)&", "&sqlTexto(CARGO)&", "&sqlTexto(ATIVADO)&", "&sqlTexto(Session("cod_usuario"))&")")
	
		'Encontrar chave criada
		Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_FORNECEDORES LIMIT 1")
			If not rs.Eof Then
				cod_fornecedor = rs("ULTIMO_ID")
			End if 
		rs.Close
		Set rs=Nothing

	FecharConexao()
	Response.Redirect(redirecionar&"?cod_fornecedor="&cod_fornecedor&"&codmsg=1")
End sub

Sub alterarRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("UPDATE TB_FORNECEDORES SET cod_fornecedor_stillos = "&sqlNumero(cod_fornecedor_stillos)&", RAZAO_SOCIAL = "&sqlTexto(RAZAO_SOCIAL)&", NOME_FANTASIA = "&sqlTexto(NOME_FANTASIA)&", CNPJ = "&sqlTexto(CNPJ)&", ENDERECO = "&sqlTexto(ENDERECO)&", BAIRRO = "&sqlTexto(BAIRRO)&", CEP = "&sqlTexto(CEP)&", CIDADE = "&sqlTexto(CIDADE)&", ESTADO = "&sqlTexto(ESTADO)&", TELEFONE = "&sqlTexto(TELEFONE)&", CONTATO = "&sqlTexto(CONTATO)&", CARGO = "&sqlTexto(CARGO)&", ATIVADO = "&sqlTexto(ATIVADO)&", CADASTRADO_POR = "&sqlNumero(Session("cod_usuario"))&" WHERE cod_fornecedor = "&cod_fornecedor)
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_fornecedor="&cod_fornecedor&"&codmsg=2")
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()
		'call verificaAntesExcluir(,cod_fornecedor,"TABELASRELACIONADAS;TABELASRELACIONADAS","cod_fornecedor",redirecionar_excluir&"?sucesso=2&codmsg=5")

	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_FORNECEDORES WHERE cod_fornecedor = "&cod_fornecedor)
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_fornecedor="&cod_fornecedor&"&codmsg=3")
End sub


'----------------------------------------------------------------------------------------

%>

<h2>Formul&aacute;rio de Fornecedor</h2><br />


<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post'>
<table width='90%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
<%If comando <> "incluir" Then %><tr>
<td width='27%'><strong>Cód. Fornecedor </strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<span class='textoMaior'><strong><%=COD_FORNECEDOR%></strong></span><%Else%>
<%If not isNull(COD_FORNECEDOR) Then%>
<%=COD_FORNECEDOR%>
<%End If%>
<%End If%>
</div></td>
</tr>
<%End If%>
<tr>
<td width='27%'><strong>Código Fornecedor Stillos*</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='COD_FORNECEDOR_STILLOS' maxlength='6' onkeypress='mascara(this,soNumeros)' class='inputTamanho3' value="<%=COD_FORNECEDOR_STILLOS%>" type='text' />
<%Else%>
<%If not isNull(COD_FORNECEDOR_STILLOS) Then%>
<%=COD_FORNECEDOR_STILLOS%>
<%End If%>
<%End If%>
</div></td>
</tr>


<tr>
	<td width='27%'><strong>Tipo de pessoa *</strong></td>
	<td width='73%'>
		<div align='left'>
			<%If comando <> "visualizar" and comando <> "excluir"  Then%>

				<select name='TIPOPESSOA' class='inputTamanho1'>				
						<%If CStr(TIPOPESSOA) =  "J" Then%>
								<option selected='selected' value="J">PJ</option>
						<%Else%>
								<option value="J">PJ</option>
						<%End if%>

						<%If CStr(TIPOPESSOA) =  "F" Then%>
							<option selected='selected' value="F">PF</option>
						<%Else%>
							<option value="F">PF</option>
						<%End if%>
				</select>

			<%End If%>
		</div>
	</td>
</tr>

<%response.Write(TIPOPESSOA)%>

<tr>
<td width='27%'><strong>Razão Social/Nome *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='RAZAO_SOCIAL' value="<%=RAZAO_SOCIAL%>" class='inputTamanho3' type='text' maxlength='255' />
<%Else%>
<%If not isNull(RAZAO_SOCIAL) Then%>
<%=RAZAO_SOCIAL%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Nome Fantasia/Apelido *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='NOME_FANTASIA' value="<%=NOME_FANTASIA%>" class='inputTamanho3' type='text' maxlength='255' />
<%Else%>
<%If not isNull(NOME_FANTASIA) Then%>
<%=NOME_FANTASIA%>
<%End If%>
<%End If%>
</div></td>
</tr>

<tr>
	<td width='27%'><strong>CNPJ/CPF *</strong></td>
	<td width='73%'>	
		<div align='left'>
			<%If comando <> "visualizar" and comando <> "excluir"  Then%>

				<%IF CStr(TIPOPESSOA) =  "F" Then%>
					<input name='CPF' onblur='ValidarCPF(this)' class='inputTamanho2 padrao_cpf' value="<%=CPF%>" type='text' size='40' maxlength='11' />
				<%Else%>	
					<input name='CNPJ' onblur='ValidarCNPJ(this)' class='inputTamanho2 padrao_cnpj' value="<%=CNPJ%>" type='text' size='40' maxlength='18' />		
				<%End if%>

			<%Else%>

				<%If not isNull(CNPJ) Then%>
					<%=CNPJ%>
				<%End If%>		

			<%End If%>
		</div>
	</td>
</tr>

<tr>
<td width='27%'><strong>Endereço </strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='ENDERECO' value="<%=ENDERECO%>" class='inputTamanho3' type='text' maxlength='255' />
<%Else%>
<%If not isNull(ENDERECO) Then%>
<%=ENDERECO%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Bairro </strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='BAIRRO' value="<%=BAIRRO%>" class='inputTamanho3' type='text' maxlength='255' />
<%Else%>
<%If not isNull(BAIRRO) Then%>
<%=BAIRRO%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>CEP </strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='CEP' value="<%=CEP%>" class='inputTamanho2 padrao_cep' type='text' maxlength='9' />
<%Else%>
<%If not isNull(CEP) Then%>
<%=CEP%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Cidade</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='CIDADE' value="<%=CIDADE%>" class='inputTamanho3' type='text' maxlength='255' />
<%Else%>
<%If not isNull(CIDADE) Then%>
<%=CIDADE%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Estado </strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='ESTADO' value="<%=ESTADO%>" class='inputTamanho3' type='text' maxlength='2' />
<%Else%>
<%If not isNull(ESTADO) Then%>
<%=ESTADO%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Telefone *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='TELEFONE' class='inputTamanho2 padrao_telefone' value="<%=TELEFONE%>" type='text' size='40' maxlength='14' /><%Else%>
<%If not isNull(TELEFONE) Then%>
<%=TELEFONE%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Contato *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='CONTATO' value="<%=CONTATO%>" class='inputTamanho3' type='text' maxlength='255' />
<%Else%>
<%If not isNull(CONTATO) Then%>
<%=CONTATO%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Cargo </strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='CARGO' value="<%=CARGO%>" class='inputTamanho3' type='text' maxlength='255' />
<%Else%>
<%If not isNull(CARGO) Then%>
<%=CARGO%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Ativo *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<select name='ATIVADO' class='inputTamanho1'>
<option value=''>Selecione</option>
<%'Padrao
Dim selRelPadrao14
selRelPadrao14 = "2" 

If ATIVADO = Empty or isNull(ATIVADO) Then
ATIVADO = selRelPadrao14
End if%>

<%If CStr(ATIVADO) =  "1" Then
%><option selected='selected' value="1">Sim</option>
<%Else%>
<option value="1">Sim</option>
<%End if%>
<%If CStr(ATIVADO) =  "2" Then
%><option selected='selected' value="2">NÃO</option>
<%Else%>
<option value="2">NÃO</option>
<%End if%>
</select>
<%Else%>
<%If not isNull(ATIVADO) Then%>
<%If CStr(ATIVADO) =  "1" Then%>
Sim<%End if%>
<%If CStr(ATIVADO) =  "2" Then%>
NÃO<%End if%>
<%End If%>
<%End If%>
</div></td>
</tr>
<%If comando <> "visualizar" Then%>
<tr>
<td colspan='2' align='center'><div align='left'><strong>* - Campos obrigatórios</strong><br /></div></td>
</tr>
<tr>
  <td colspan='2' align='center'><a href="javascript:salvar();"><img src="images/layout/btnGravar.jpg" /></a></td>
</tr>
<%End If%>
</table>
<input type='hidden' value='1' name='campo1'  />
<input type='hidden' name='COD_FORNECEDOR' value="<%=COD_FORNECEDOR%>" /> <input type='hidden' name='CADASTRADO_POR' value="<%=CADASTRADO_POR%>" /> </form>
<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('COD_FORNECEDOR_STILLOS#Código Forncedor Stillos#1|RAZAO_SOCIAL#Razão Social#1|NOME_FANTASIA#Nome Fantasia#1|CNPJ#CNPJ#1|TELEFONE#Telefone#1|ATIVADO#Ativo#2',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->
