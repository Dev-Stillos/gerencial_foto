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
'Criado OXWebDeveloper em: 26/6/2012 17:06:11

'Declara variáveis para utilização do sistema
Dim cod_forma_pagamento, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

redirecionar_incluir	= "ppv_formas_pagamento.asp"
redirecionar_alterar	= "ppv_formas_pagamento.asp"
redirecionar_excluir	= "ppv_formas_pagamento.asp"
redirecionar_validacao	= "ppv_formas_pagamento.asp"
msgAoExcluir			= "Tem certeza que deseja excluir esta forma de pagamento?<br /><br />Você poderá inativa-la ao invés de exclui-la. Se houver registros com esta forma de pagamento, a mesma não será excluída"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
'Declara variáveis
Dim descricao, condicao_pagamento, qtd_parcelas, parcelas, obs, cod_usuario, modificado, situacao

If comando = Empty Then
comando = "incluir"
titulo_pagina = titulo_pagina&" / Inclusão"
End if


'#########################################################################
'Validação de sub-lidações padrões com formulário
'#########################################################################

If comando = "incluir" and campo1 <> Empty Then
'Incluir novo registro
incluirRegistro(redirecionar_incluir)
titulo_pagina = titulo_pagina&" / Inclusão"
Elseif comando = "alterar" and campo1 = Empty Then
'Leitura inicial para possível alteração
cod_forma_pagamento = Request.QueryString("cod_forma_pagamento")
lerRegistro()
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "alterar" and campo1 <> Empty Then
'Gravar alteração
cod_forma_pagamento = Request.Form("cod_forma_pagamento")
alterarRegistro(redirecionar_alterar)
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 
'Leitura inicial para possível exclusão
cod_forma_pagamento = Request.QueryString("cod_forma_pagamento")
lerRegistro()
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
'Exclui registro
cod_forma_pagamento = Request.Form("cod_forma_pagamento")
excluirRegistro(redirecionar_excluir)
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
'Visualiza registro
cod_forma_pagamento = Request.QueryString("cod_forma_pagamento")
lerRegistro()
titulo_pagina = titulo_pagina&" / Visualização"
End If


'Funções

Sub lerFormulario()

cod_forma_pagamento = Request.Form("cod_forma_pagamento")
descricao = Request.Form("descricao")
condicao_pagamento = Request.Form("condicao_pagamento")
qtd_parcelas = Request.Form("qtd_parcelas")
parcelas = Request.Form("parcelas")
obs = Request.Form("obs")
cod_usuario = Session("cod_usuario")
modificado = now()
situacao = Request.Form("situacao")

End sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objconexao.Execute("SELECT * FROM TB_FORMAS_PAGAMENTO WHERE cod_forma_pagamento = "&cod_forma_pagamento)

				descricao=trim(rs("descricao"))
				condicao_pagamento=trim(rs("condicao_pagamento"))
				qtd_parcelas=trim(rs("qtd_parcelas"))
				parcelas=trim(rs("parcelas"))
				obs=trim(rs("obs"))
				cod_usuario=trim(rs("cod_usuario"))
				modificado=trim(rs("modificado"))
				situacao=trim(rs("situacao"))

		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objconexao.Execute("INSERT INTO TB_FORMAS_PAGAMENTO (descricao, condicao_pagamento, qtd_parcelas, parcelas, obs, cod_usuario, modificado, situacao) VALUES ("&sqlTexto(descricao)&", "&sqlTexto(condicao_pagamento)&", "&sqlTexto(qtd_parcelas)&", "&sqlTexto(parcelas)&", "&sqlTexto(obs)&", "&sqlTexto(cod_usuario)&", "&sqlData(modificado)&", "&sqlTexto(situacao)&")")

		'Encontrar chave criada
		Set rs=objconexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_FORMAS_PAGAMENTO LIMIT 1")
			If not rs.Eof Then
				cod_forma_pagamento = rs("ULTIMO_ID")
			End if 
		rs.Close
		Set rs=Nothing


	FecharConexao()
	Response.Redirect(redirecionar&"?cod_forma_pagamento="&cod_forma_pagamento&"&codmsg=1")
End sub

Sub alterarRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objconexao.Execute("UPDATE TB_FORMAS_PAGAMENTO SET descricao = "&sqlTexto(descricao)&", condicao_pagamento = "&sqlTexto(condicao_pagamento)&", qtd_parcelas = "&sqlTexto(qtd_parcelas)&", parcelas = "&sqlTexto(parcelas)&", obs = "&sqlTexto(obs)&", cod_usuario = "&sqlTexto(cod_usuario)&", modificado = "&sqlData(modificado)&", situacao = "&sqlTexto(situacao)&" WHERE cod_forma_pagamento = "&cod_forma_pagamento)
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_forma_pagamento="&cod_forma_pagamento&"&codmsg=2")
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()
		'call verificaAntesExcluir(,cod_forma_pagamento,"TABELASRELACIONADAS;TABELASRELACIONADAS","cod_forma_pagamento",redirecionar_excluir&"?sucesso=2&codmsg=5")

	AbrirConexao()
		objconexao.Execute("DELETE FROM TB_FORMAS_PAGAMENTO WHERE cod_forma_pagamento = "&cod_forma_pagamento)
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_forma_pagamento="&cod_forma_pagamento&"&codmsg=3")
End sub


'----------------------------------------------------------------------------------------

%>

<h2><%=titulo_pagina%></h2><br />

<%If comando = "excluir" Then%>
<script type="text/javascript">
setTimeout("parent.submenu.trocarPorGravarporExclusao()",1)
</script>
<div class='msgExcluir'><%=msgAoExcluir%></div>
<%Elseif comando = "visualizar" Then%>
<script type="text/javascript">
setTimeout("parent.submenu.comandoVisualizar()",1)
</script>
<%End If%>

<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post'>
<table width='90%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
<%If comando <> "incluir" Then %><tr>
<td width='27%'><strong>Cód. Forma Pagamento </strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<span class='textoMaior'><strong><%=COD_FORMA_PAGAMENTO%></strong></span><%Else%>
<%If not isNull(COD_FORMA_PAGAMENTO) Then%>
<%=COD_FORMA_PAGAMENTO%>
<%End If%>
<%End If%>
</div></td>
</tr>
<%End If%><tr>
<td width='27%'><strong>Descrição *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='DESCRICAO' value="<%=DESCRICAO%>" class='inputTamanho3' type='text' maxlength='255' />
<%Else%>
<%If not isNull(DESCRICAO) Then%>
<%=DESCRICAO%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Condição Pagamento *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<select name='CONDICAO_PAGAMENTO' class='inputTamanho1'>
<option value=''>Selecione</option>
<%'Padrao
Dim selRelPadrao3
selRelPadrao3 = "V" 

If CONDICAO_PAGAMENTO = Empty or isNull(CONDICAO_PAGAMENTO) Then
CONDICAO_PAGAMENTO = selRelPadrao3
End if%>

<%If CStr(CONDICAO_PAGAMENTO) =  "V" Then
%><option selected='selected' value="V"> À Vista</option>
<%Else%>
<option value="V"> À Vista</option>
<%End if%>
<%If CStr(CONDICAO_PAGAMENTO) =  "P" Then
%><option selected='selected' value="P">À Prazo</option>
<%Else%>
<option value="P">À Prazo</option>
<%End if%>
</select>
<%Else%>
<%If not isNull(CONDICAO_PAGAMENTO) Then%>
<%If CStr(CONDICAO_PAGAMENTO) =  "V" Then%>
 À Vista<%End if%>
<%If CStr(CONDICAO_PAGAMENTO) =  "P" Then%>
À Prazo<%End if%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Quantidade de Parcelas *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='QTD_PARCELAS' maxlength='4' onkeypress='mascara(this,soNumeros)' class='inputTamanho3' value="<%=QTD_PARCELAS%>" type='text' />
<%Else%>
<%If not isNull(QTD_PARCELAS) Then%>
<%=QTD_PARCELAS%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Parcelas *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='PARCELAS' value="<%=PARCELAS%>" class='inputTamanho3' type='text' maxlength='255' />
<%Else%>
<%If not isNull(PARCELAS) Then%>
<%=PARCELAS%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Observação </strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<table style='width:420px;' border='0' cellspacing='0' cellpadding='0'>
<tr>
<td width='57'><textarea rows='5' name='OBS' class='inputTamanho3' onkeydown='textCounter(this.form.OBS,this.form.charRestantes_OBS,300);' onkeyup='textCounter(this.form.OBS,this.form.charRestantes_OBS,300);'><%=OBS%></textarea></td>
</tr>
<tr>
<td><table width='50' border='0' align='right' cellpadding='0' cellspacing='0'>
<tr>
<td><div align='center'>
<input class='inputContaCaracteres' tabindex='10000' type='text' name='charRestantes_OBS' value="<%=300- Len(OBS)%>" />
</div></td>
</tr>
</table></td>
</tr>
</table>
<%Else%>
<%If not isNull(OBS) Then%>
<%=Replace(OBS,Chr(13),"<br />")%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Situação *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<select name='SITUACAO' class='inputTamanho2'>
<option value=''>Selecione</option>
<%'Padrao
Dim selRelPadrao9
selRelPadrao9 = "A" 

If SITUACAO = Empty or isNull(SITUACAO) Then
SITUACAO = selRelPadrao9
End if%>

<%If CStr(SITUACAO) =  "A" Then
%><option selected='selected' value="A">Ativo</option>
<%Else%>
<option value="A">Ativo</option>
<%End if%>
<%If CStr(SITUACAO) =  "I" Then
%><option selected='selected' value="I">Inativo</option>
<%Else%>
<option value="I">Inativo</option>
<%End if%>
</select>
<%Else%>
<%If not isNull(SITUACAO) Then%>
<%If CStr(SITUACAO) =  "A" Then%>
Ativo<%End if%>
<%If CStr(SITUACAO) =  "I" Then%>
Inativo<%End if%>
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
<input type='hidden' name='COD_FORMA_PAGAMENTO' value="<%=COD_FORMA_PAGAMENTO%>" /> </form>
<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('DESCRICAO#Descrição#1|CONDICAO_PAGAMENTO#Condição Pagamento#2|QTD_PARCELAS#Quantidade de Parcelas#1|PARCELAS#Parcelas#1|SITUACAO#Situação#2',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->
