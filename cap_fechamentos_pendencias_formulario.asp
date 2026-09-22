<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 1
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
  Dim titulo_pagina
  titulo_pagina = "CAP - Fechamento Pend&ecirc;ncia de Pagamento Formul&aacute;rio"
%>				
<%
Dim calendario_necessario
calendario_necessario = "N&Atilde;O"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%

'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 11/8/2011 15:50:09


'Declara vari�veis para utiliza��o do sistema
Dim cod_pendencia, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

redirecionar_incluir	= "cap_fechamentos_pendencias.asp"
redirecionar_alterar	= "cap_fechamentos_pendencias.asp"
redirecionar_excluir	= "cap_fechamentos_pendencias.asp"
redirecionar_validacao	= "cap_fechamentos_pendencias.asp"
msgAoExcluir			= "Tem certeza que deseja excluir esta pend&ecirc;ncia de Pagamento?"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
'Declara vari�veis
Dim COD_VENDEDOR, TIPO, VALOR, MOTIVO, CADASTRADO_POR, cod_fechamento_periodo

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
cod_pendencia = Request.QueryString("cod_pendencia")
lerRegistro()
titulo_pagina = titulo_pagina&" / Altera&ccedil;&atilde;o"
Elseif 	comando = "alterar" and campo1 <> Empty Then
'Gravar altera��o
cod_pendencia = Request.Form("cod_pendencia")
alterarRegistro(redirecionar_alterar)
titulo_pagina = titulo_pagina&" / Altera&ccedil;&atilde;o"
Elseif 	comando = "excluir" and campo1 = Empty Then 
'Leitura inicial para poss�vel exclus�o
cod_pendencia = Request.QueryString("cod_pendencia")
lerRegistro()
titulo_pagina = titulo_pagina&" / Exclus&atilde;o"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
'Exclui registro
cod_pendencia = Request.Form("cod_pendencia")
excluirRegistro(redirecionar_excluir)
titulo_pagina = titulo_pagina&" / Exclus&atilde;o"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
'Visualiza registro
cod_pendencia = Request.QueryString("cod_pendencia")
lerRegistro()
titulo_pagina = titulo_pagina&" / Visualiza&ccedil;&atilde;o"
End If


'Fun��es

Sub lerFormulario()

COD_VENDEDOR = Request.Form("COD_VENDEDOR")
TIPO = Request.Form("TIPO")
VALOR = Request.Form("VALOR")
MOTIVO = Request.Form("MOTIVO")
CADASTRADO_POR = Request.Form("CADASTRADO_POR")
cod_fechamento_periodo = Request.Form("cod_fechamento_periodo")

End sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM TB_FECHAMENTO_PENDENCIAS_PAGAMENTO WHERE cod_pendencia = "&cod_pendencia)

				COD_VENDEDOR=trim(rs("COD_VENDEDOR"))
				TIPO=trim(rs("TIPO"))
				VALOR= ArrumaNumero(rs("VALOR"),"",2)
				MOTIVO=trim(rs("MOTIVO"))
				CADASTRADO_POR=trim(rs("CADASTRADO_POR"))
				cod_fechamento_periodo=trim(rs("cod_fechamento_periodo"))

		rs.Close
		Set rs=Nothing
	FecharConexao()
	
	'Revalida��o para verificar se n�o est�o tentando burlar o sistema e alterar uma pendencia que o pagamento j� foi realizado
	If int(dateDiff("d",ULTIMO_PAGAMENTO_PROCESSADO,cod_fechamento_periodo)) <= 0 Then
		Response.Redirect("cap_fechamentos_pendencias.asp?cod_pendencia="&cod_pendencia&"&codmsg=41&sucesso=2")	
	End If	
	
End sub

Sub incluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("INSERT INTO TB_FECHAMENTO_PENDENCIAS_PAGAMENTO (COD_VENDEDOR, TIPO, VALOR, MOTIVO, CADASTRADO_POR, cod_fechamento_periodo) VALUES ("&sqlNumero(COD_VENDEDOR)&", "&sqlTexto(TIPO)&", "&sqlNumero(VALOR)&", "&sqlTexto(MOTIVO)&", "&sqlTexto(CADASTRADO_POR)&", "&sqlTexto(cod_fechamento_periodo)&")")

		'Encontrar chave criada
		Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_FECHAMENTO_PENDENCIAS_PAGAMENTO LIMIT 1")
			If not rs.Eof Then
				cod_pendencia = rs("ULTIMO_ID")
			End if 
		rs.Close
		Set rs=Nothing


	FecharConexao()
	Response.Redirect(redirecionar&"?cod_pendencia="&cod_pendencia&"&codmsg=1")
End sub

Sub alterarRegistro(redirecionar)
	lerFormulario()
	
	AbrirConexao()
		objConexao.Execute("UPDATE TB_FECHAMENTO_PENDENCIAS_PAGAMENTO SET COD_VENDEDOR = "&sqlNumero(COD_VENDEDOR)&", TIPO = "&sqlTexto(TIPO)&", VALOR = "&sqlNumero(VALOR)&", MOTIVO = "&sqlTexto(MOTIVO)&", cod_fechamento_periodo = "&sqlTexto(cod_fechamento_periodo)&" WHERE cod_pendencia = "&cod_pendencia)
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_pendencia="&cod_pendencia&"&codmsg=2")
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_FECHAMENTO_PENDENCIAS_PAGAMENTO WHERE cod_pendencia = "&cod_pendencia)
	FecharConexao()
	
	Response.Redirect(redirecionar&"?cod_pendencia="&cod_pendencia&"&codmsg=3")
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
<td width='27%'><strong>C&oacute;d Pend&ecirc;ncia </strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<span class='textoMaior'><strong><%=COD_PENDENCIA%></strong></span><%Else%>
<%If not isNull(COD_PENDENCIA) Then%>
<%=COD_PENDENCIA%>
<%End If%>
<%End If%>
</div></td>
</tr>
<%End If%><tr>
<td width='27%'><strong>Vendedor *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<select name='COD_VENDEDOR' class='inputTamanho2'>
<option value=''>Selecione</option>
<%AbrirConexao()
Set rsE2=objConexao.Execute("SELECT cod_vendedor, nome_vendedor FROM TB_VENDEDORES where ATIVADO = 1 and lote = 2 ORDER BY nome_vendedor ASC")
While not rsE2.Eof 
'Padrao
Dim selRelPadrao2, verificador2
verificador2 = ""&rsE2(0)

selRelPadrao2 = "" 

If COD_VENDEDOR = Empty or isNull(COD_VENDEDOR) Then
COD_VENDEDOR = selRelPadrao2
End if

If CStr(verificador2) = CStr(COD_VENDEDOR)  Then%>
<option selected='selected' value="<%=rsE2(0)%>"><%=rsE2(1)%></option>
<%Else%>
<option value="<%=rsE2(0)%>"><%=rsE2(1)%></option>
<%End if%>
<%rsE2.MoveNext()
Wend
rsE2.Close
Set rsE2=Nothing
FecharConexao()%>

</select>
<%Else%>
<%If not isNull(COD_VENDEDOR) Then%>
<%AbrirConexao()
Set rsE2=objConexao.Execute("SELECT cod_vendedor, nome_vendedor FROM TB_VENDEDORES where ATIVADO = 1 and lote = 2 ORDER BY nome_vendedor ASC")
While not rsE2.Eof 
If CStr(rsE2(0)) = CStr(COD_VENDEDOR)  Then%>
<%=rsE2(1)%><%End if%>
<%rsE2.MoveNext()
Wend
rsE2.Close
Set rsE2=Nothing
FecharConexao()%>

<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Tipo *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<select name='TIPO' class='inputTamanho2'>
<option value=''>Selecione</option>
<%'Padrao
Dim selRelPadrao3
selRelPadrao3 = "" 

If TIPO = Empty or isNull(TIPO) Then
TIPO = selRelPadrao3
End if%>

<%If CStr(TIPO) =  "A" Then
%><option selected='selected' value="A">Acr&eacute;scimo</option>
<%Else%>
<option value="A">Acr&eacute;scimo</option>
<%End if%>
<%If CStr(TIPO) =  "D" Then
%><option selected='selected' value="D">Decr&eacute;scimo</option>
<%Else%>
<option value="D">Decr&eacute;scimo</option>
<%End if%>
</select>
<%Else%>
<%If not isNull(TIPO) Then%>
<%If CStr(TIPO) =  "A" Then%>
Acr&eacute;scimo<%End if%>
<%If CStr(TIPO) =  "D" Then%>
Decr&eacute;scimo<%End if%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Valor </strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='VALOR' value="<%=VALOR%>" class='inputTamanho1 padrao_valor' type='text' maxlength='200' />
<%Else%>
<%If not isNull(VALOR) Then%>
<%=VALOR%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Motivo *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<table style='width:420px;' border='0' cellspacing='0' cellpadding='0'>
<tr>
<td width='57'><textarea rows='5' name='MOTIVO' class='inputTamanho3' onkeydown='textCounter(this.form.MOTIVO,this.form.charRestantes_MOTIVO,255);' onkeyup='textCounter(this.form.MOTIVO,this.form.charRestantes_MOTIVO,255);'><%=MOTIVO%></textarea></td>
</tr>
<tr>
<td><table width='50' border='0' align='right' cellpadding='0' cellspacing='0'>
<tr>
  <td>
    <div align="right">
      <input class='inputContaCaracteres' tabindex='10000' style="width:25px" type='text' name='charRestantes_MOTIVO' value="<%=255- Len(MOTIVO)%>" />
  </div></td>
</tr>
</table></td>
</tr>
</table>
<%Else%>
<%If not isNull(MOTIVO) Then%>
<%=Replace(MOTIVO,Chr(13),"<br />")%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Agendar para pagamento *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<select name="cod_fechamento_periodo" class="inputTamanho2">
		
        <%
		AbrirConexao()
		Set rs2=objConexao.Execute("SELECT cod_fechamento_periodo, data_pagamento FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE data_processado_pagamento IS NULL ORDER BY data_pagamento ASC ")
								
		If not rs2.Eof Then
				While not rs2.Eof
					%>
					<option <%if int(cod_fechamento_periodo) = int(rs2(0)) Then%> selected="selected" <%End If%> value="<%=rs2(0)%>"><%=ArrumaData(rs2(1))%></option>
					<%
				rs2.MoveNext()
				Wend	
		End If
								
	rs2.Close
	Set rs2=Nothing
	FecharConexao()
	%>
     </select>   
        
		
<br />
(&Eacute; poss&iacute;vel selecionar apenas para pagamentos que n&atilde;o foram processados)
<%Else%>
<%End If%>
</div></td>
</tr>
<%If comando <> "visualizar" Then%>
<tr>
<td colspan='2' align='center'><div align='left'>
  <p><strong>Campos obrigat&oacute;rios, cont&eacute;m: *</strong></p>
  <p align="center"><a href="javascript:salvar();"><img border="0" src="images/layout/btnGravar.jpg" /></a>
  </p>
</div></td>
</tr>
<%End If%>
</table>
<input type='hidden' value='1' name='campo1'  />
<input type='hidden' name='COD_PENDENCIA' value="<%=COD_PENDENCIA%>" /> <input type='hidden' name='CADASTRADO_POR' value="<%=Session("cod_usuario")%>" /> </form>
<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('COD_VENDEDOR#Vendedor#2|TIPO#Tipo#2|MOTIVO#Motivo#5|cod_fechamento_periodo#Agendar para pagamento#1',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>
<!-- #include file="files/includes/miolo-inferior.asp" -->