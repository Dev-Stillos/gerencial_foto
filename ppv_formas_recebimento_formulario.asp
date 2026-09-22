				
<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 2
  BTN_FONTE = 1
  BTN_IMPRIMIR_LINK = "#"
  %>
<%
Dim calendario_necessario
calendario_necessario = "N&AtildeO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%

'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 26/6/2012 15:24:34

'Declara vari&aacuteveis para utiliza&ccedil&atildeo do sistema
Dim cod_forma_recebimento, pode_excluir, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

redirecionar_incluir	= "ppv_formas_recebimento.asp"
redirecionar_alterar	= "ppv_formas_recebimento.asp"
redirecionar_excluir	= "ppv_formas_recebimento.asp"
redirecionar_validacao	= "ppv_formas_recebimento.asp"
msgAoExcluir			= "Tem certeza que deseja excluir esta forma de recebimento?<br /><br />Voc&ecirc poder&aacute inativa-la ao inv&eacutes de exclui-la. Se houver registors com esta forma, a mesma n&atildeo ser&aacute exclu&iacuteda"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
'Declara vari&aacuteveis
Dim descricao, obs, cod_usuario, modificado, situacao

If comando = Empty Then
comando = "incluir"
titulo_pagina = titulo_pagina&" / Incluso"
End if


'#########################################################################
'Valida&ccedil&atildeo de sub-lida&ccedil&atildees padr&otildees com formul&aacuterios
'#########################################################################

If comando = "incluir" and campo1 <> Empty Then
'Incluir novo registro
incluirRegistro(redirecionar_incluir)
titulo_pagina = titulo_pagina&" / Inclus&atildeo"
Elseif comando = "alterar" and campo1 = Empty Then
'Leitura inicial para poss&iacutevel altera&ccedil&atildeo
cod_forma_recebimento = Request.QueryString("cod_forma_recebimento")
lerRegistro()
titulo_pagina = titulo_pagina&" / Altera&ccedil&atildeo"
Elseif 	comando = "alterar" and campo1 <> Empty Then
'Gravar altera&ccedil&atildeo
cod_forma_recebimento = Request.Form("cod_forma_recebimento")
alterarRegistro(redirecionar_alterar)
titulo_pagina = titulo_pagina&" / Altera&ccedil&atildeo"
Elseif 	comando = "excluir" and campo1 = Empty Then 
'Leitura inicial para poss&iacutevel exclus&atildeo
cod_forma_recebimento = Request.QueryString("cod_forma_recebimento")
lerRegistro()
titulo_pagina = titulo_pagina&" / Exclus&atildeo"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
'Exclui registro
cod_forma_recebimento = Request.Form("cod_forma_recebimento")
excluirRegistro(redirecionar_excluir)
titulo_pagina = titulo_pagina&" / Exclus&atildeo"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
'Visualiza registro
cod_forma_recebimento = Request.QueryString("cod_forma_recebimento")
lerRegistro()
titulo_pagina = titulo_pagina&" / Visualiza&ccedil&atildeo"
End If


'Fun&ccedil&otildees

Sub lerFormulario()

cod_forma_recebimento = Request.Form("cod_forma_recebimento")
descricao = Request.Form("descricao")
obs = Request.Form("obs")
pode_excluir = Request.Form("pode_excluir")
cod_usuario = Session("cod_usuario")
modificado = now()
situacao = Request.Form("situacao")

End sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM TB_FORMAS_RECEBIMENTO WHERE cod_forma_recebimento = "&cod_forma_recebimento)

				descricao=trim(rs("descricao"))
				obs=trim(rs("obs"))
				pode_excluir=trim(rs("pode_excluir"))
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
		objConexao.Execute("INSERT INTO TB_FORMAS_RECEBIMENTO (descricao, obs, cod_usuario, modificado, situacao) VALUES ("&sqlTexto(descricao)&", "&sqlTexto(obs)&", "&sqlTexto(cod_usuario)&", "&sqlData(modificado)&", "&sqlTexto(situacao)&")")

		'Encontrar chave criada
		Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_FORMAS_RECEBIMENTO LIMIT 1")
			If not rs.Eof Then
				cod_forma_recebimento = rs("ULTIMO_ID")
			End if 
		rs.Close
		Set rs=Nothing


	FecharConexao()
	Response.Redirect(redirecionar&"?cod_forma_recebimento="&cod_forma_recebimento&"&codmsg=1")
End sub

Sub alterarRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("UPDATE TB_FORMAS_RECEBIMENTO SET descricao = "&sqlTexto(descricao)&", obs = "&sqlTexto(obs)&", cod_usuario = "&sqlTexto(cod_usuario)&", modificado = "&sqlData(modificado)&", situacao = "&sqlTexto(situacao)&" WHERE cod_forma_recebimento = "&cod_forma_recebimento)
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_forma_recebimento="&cod_forma_recebimento&"&codmsg=2")
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()
		'call verificaAntesExcluir(,cod_forma_recebimento,"TABELASRELACIONADAS;TABELASRELACIONADAS","cod_forma_recebimento",redirecionar_excluir&"?sucesso=2&codmsg=5")

	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_FORMAS_RECEBIMENTO WHERE cod_forma_recebimento = "&cod_forma_recebimento)
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_forma_recebimento="&cod_forma_recebimento&"&codmsg=3")
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
<td width='27%'><strong>C&oacuted. Forma de Recebimento </strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<span class='textoMaior'><strong><%=COD_FORMA_RECEBIMENTO%></strong></span><%Else%>
<%If not isNull(COD_FORMA_RECEBIMENTO) Then%>
<%=COD_FORMA_RECEBIMENTO%>
<%End If%>
<%End If%>
</div></td>
</tr>
<%End If%><tr>
<td width='27%'><strong>Descri&ccedil&atildeo *</strong></td>
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
<td width='27%'><strong>Observa&ccedil&atildeo </strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<table style='width:420px;' border='0' cellspacing='0' cellpadding='0'>
<tr>
<td width='57'>
<textarea rows='5' name='OBS' class='inputTamanho3' onkeydown='textCounter(this.form.OBS,this.form.charRestantes_OBS,300);' onkeyup='textCounter(this.form.OBS,this.form.charRestantes_OBS,300);'><%=OBS%></textarea>
</td>
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
  <td width='27%'><strong>Situa&ccedil&atildeo *</strong></td>
  <td width='73%'>
  <div align='left'>
  <%If comando <> "visualizar" and comando <> "excluir"  Then%>
  <select name='SITUACAO' class='inputTamanho2'>
  <option value=''>Selecione</option>
  <%'Padrao
Dim selRelPadrao7
selRelPadrao7 = "A" 

If SITUACAO = Empty or isNull(SITUACAO) Then
SITUACAO = selRelPadrao7
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
<td colspan='2' align='center'><div align='left'><strong>* - Campos obrigat&oacuterios</strong><br /></div></td>
</tr>
<tr>
  <td colspan='2' align='center'><a href="javascript:salvar();"><img src="images/layout/btnGravar.jpg" /></a></td>
</tr>
<%End If%>
</table>
<input type='hidden' value='1' name='campo1'  />
<input type='hidden' value='<%=PODE_EXCLUIR%>' name='PODE_EXCLUIR'  />
<input type='hidden' name='COD_FORMA_RECEBIMENTO' value="<%=COD_FORMA_RECEBIMENTO%>" /> </form>
<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('DESCRICAO#Descri&ccedil&atildeo#1|SITUACAO#Situa&ccedil&atildeo#2',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->
