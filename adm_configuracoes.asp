<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 1
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
  Dim titulo_pagina : titulo_pagina = "Aprovações Pendentes"

  Dim calendario_necessario : calendario_necessario = "NÃO"

%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%

'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 17/10/2011 14:29:33

'Declara variáveis para utilização do sistema
Dim cod_aprovacao, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

redirecionar_alterar	= ""
actionformulario		= "adm_aprovacoes_pendentes_formulario.asp"

campo1 = Request.Form("campo1")
comando = "alterar"
'Declara variáveis
Dim cod_categoria, solicitante_cod_usuario, responsavel_cod_usuario, motivo_solicitacao, motivo_responsavel, parametro, status_aprovacao

Dim descricao_geral, link_consulta_para_responsavel

If comando = Empty Then
comando = "incluir"
titulo_pagina = titulo_pagina&" / Inclusão"
End if

If comando = "incluir" and campo1 <> Empty Then

Elseif comando = "alterar" and campo1 = Empty Then
  'Leitura inicial para possível alteração
  cod_aprovacao = Request.QueryString("cod_aprovacao")
  lerRegistro()
  titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "alterar" and campo1 <> Empty Then
  'Gravar alteração
  cod_aprovacao = Request.Form("cod_aprovacao")
  alterarRegistro(redirecionar_alterar)
  titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 

Elseif 	comando = "excluir" and campo1 <> Empty Then 

Elseif 	comando = "visualizar" and campo1 = Empty Then 

End If


'Funções

Sub lerFormulario()

  cod_aprovacao = Request.Form("cod_aprovacao")
  cod_categoria = Request.Form("cod_categoria")
  solicitante_cod_usuario = Request.Form("solicitante_cod_usuario")
  responsavel_cod_usuario = Request.Form("responsavel_cod_usuario")
  data_cadastro = Request.Form("data_cadastro")
  motivo_solicitacao = Request.Form("motivo_solicitacao")
  motivo_responsavel = Request.Form("motivo_responsavel")
  parametro = Request.Form("parametro")
  status_aprovacao = Request.Form("STATUS_APROVACAO")

End sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT TB_APROVACOES.*, TB_APROVACOES_CATEGORIA.descricao_geral, TB_APROVACOES_CATEGORIA.descricao_interna_sobre_parametro, TB_APROVACOES_CATEGORIA.link_consulta_para_responsavel FROM TB_APROVACOES_CATEGORIA INNER JOIN TB_APROVACOES ON TB_APROVACOES_CATEGORIA.COD_CATEGORIA = TB_APROVACOES.COD_CATEGORIA where TB_APROVACOES.status = 0 and responsavel_cod_usuario = '"&Session("cod_usuario")&"' and cod_aprovacao = "&cod_aprovacao)
				
				if not rs.Eof Then
				
					cod_categoria=trim(rs("cod_categoria"))
					solicitante_cod_usuario=trim(rs("solicitante_cod_usuario"))
					responsavel_cod_usuario=trim(rs("responsavel_cod_usuario"))
					data_cadastro=trim(rs("data_cadastro"))
					descricao_geral=trim(rs("descricao_geral"))
					link_consulta_para_responsavel=trim(rs("link_consulta_para_responsavel"))
					
					
					motivo_solicitacao=trim(rs("motivo_solicitacao"))
					motivo_responsavel=trim(rs("motivo_responsavel"))
					parametro=trim(rs("parametro"))
					status_aprovacao=trim(rs("status"))
				
				Else
					Response.Redirect("adm_aprovacoes_pendentes.asp")
				End If

		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub alterarRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("UPDATE TB_APROVACOES SET motivo_responsavel = "&sqlTexto(motivo_responsavel)&", status = "&sqlTexto(status_aprovacao)&", data_finalizado = NOW() WHERE cod_aprovacao = "&cod_aprovacao)
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_aprovacao="&cod_aprovacao&"&codmsg=2")
End sub

'----------------------------------------------------------------------------------------

%>
<h2><%=titulo_pagina%></h2><br />

<div class='msgExcluir'>ATENÇÃO, É PERMITIDA APENAS UMA VEZ APROVAÇÃO OU REPROVAÇÃO</div>


<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post'>
<table width='90%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
<%If comando <> "incluir" Then %><tr>
<td width='27%' height="34"><strong>Cód. Aprovação </strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<span class='textoMaior'><strong><%=COD_APROVACAO%></strong></span><%Else%>
<%If not isNull(COD_APROVACAO) Then%>
<%=COD_APROVACAO%>
<%End If%>
<%End If%>
</div></td>
</tr>
<%End If%>
<tr>
<td width='27%'><strong>Solicitado por</strong></td>
<td width='73%'><div align="left">
<%
AbrirConexao()
  Set rs2=objConexao.Execute("Select nome FROM TB_ADM_USUARIOS WHERE COD_USUARIO = "&solicitante_cod_usuario)
    
    if not rs2.Eof Then
      Response.Write(rs2(0))
    End if
  
  rs2.Close
  Set rs2=Nothing
FecharConexao()
  %>
</div></td>
</tr>
<tr>
<td width='27%'><span style="font-weight: bold">Referente</span></td>
<td width='73%'><%=descricao_geral%></td>
</tr>
<tr>
<td width='27%'><span style="font-weight: bold">Mais informa&ccedil;&otilde;es</span></td>
<td width='73%'>
<a target="_blank" href="<%=link_consulta_para_responsavel%><%=parametro%>"><%=parametro%></a></td>
</tr>
<tr>
<td width='27%'><strong>Motivo da solicitação </strong></td>
<td width='73%'>
<div align='left'>
<%If not isNull(MOTIVO_SOLICITACAO) Then%>
<%=Replace(MOTIVO_SOLICITACAO,Chr(13),"<br />")%>
<%End If%>
</div></td>
</tr>

<tr>
<td width='27%'><strong>Mensagem do Responsável </strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<table style='width:420px;' border='0' cellspacing='0' cellpadding='0'>
<tr>
<td width='57'><textarea rows='4' name='MOTIVO_RESPONSAVEL' class='inputTamanho3' onkeydown='textCounter(this.form.MOTIVO_RESPONSAVEL,this.form.charRestantes_MOTIVO_RESPONSAVEL,255);' onkeyup='textCounter(this.form.MOTIVO_RESPONSAVEL,this.form.charRestantes_MOTIVO_RESPONSAVEL,255);'><%=MOTIVO_RESPONSAVEL%></textarea></td>
</tr>
<tr>
<td><table width='50' border='0' align='right' cellpadding='0' cellspacing='0'>
<tr>
<td><div align='center'>
<input class='inputContaCaracteres' style="border:solid 1px #FFFFFF; width:40px;" tabindex='10000' type='text' name='charRestantes_MOTIVO_RESPONSAVEL' value="<%=255- Len(MOTIVO_RESPONSAVEL)%>" />
</div></td>
</tr>
</table></td>
</tr>
</table>
<%Else%>
<%If not isNull(MOTIVO_RESPONSAVEL) Then%>
<%=Replace(MOTIVO_RESPONSAVEL,Chr(13),"<br />")%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Aprovar ou Reprovar *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<select name='STATUS_APROVACAO' class='inputTamanho1'>
<option value=''>Selecione</option>
<%'Padrao
Dim selRelPadrao10
selRelPadrao10 = "A" 

If STATUS_APROVACAO = Empty or isNull(STATUS_APROVACAO) Then
STATUS_APROVACAO = selRelPadrao10
End if%>

<%If CStr(STATUS_APROVACAO) =  "1" Then
%><option selected='selected' value="1">Aprovado</option>
<%Else%>
<option value="1">Aprovado</option>
<%End if%>
<%If CStr(STATUS_APROVACAO) =  "2" Then
%><option selected='selected' value="2">Reprovado</option>
<%Else%>
<option value="2">Reprovado</option>
<%End if%>
</select>
<%Else%>
<%If not isNull(STATUS_APROVACAO) Then%>
<%If CStr(STATUS_APROVACAO) =  "1" Then%>
Aprovado<%End if%>
<%If CStr(STATUS_APROVACAO) =  "2" Then%>
Reprovado<%End if%>
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
<input type='hidden' name='COD_APROVACAO' value="<%=COD_APROVACAO%>" /> <input type='hidden' name='COD_CATEGORIA' value="<%=COD_CATEGORIA%>" /> <input type='hidden' name='SOLICITANTE_COD_USUARIO' value="<%=SOLICITANTE_COD_USUARIO%>" /> <input type='hidden' name='RESPONSAVEL_COD_USUARIO' value="<%=RESPONSAVEL_COD_USUARIO%>" /> <input type='hidden' name='DATA_CADASTRO' value="<%=DATA_CADASTRO%>" /> <input type='hidden' name='PARAMETRO' value="<%=PARAMETRO%>" /> </form>
<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('STATUS_APROVACAO#Aprovar ou Reprovar#2',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->
