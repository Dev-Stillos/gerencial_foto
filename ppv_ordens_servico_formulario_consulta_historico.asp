			
<%
Dim calendario_necessario
calendario_necessario = "NÃO"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="http://stillos.com.br/gerencial_foto/css/base.css" media="screen"></link>
<link rel="stylesheet" media="all" type="text/css" href="http://stillos.com.br/gerencial_foto/css/modal-message.css">
<title>StillosWEB - Sistema corporativo - Vers&atilde;o 2.0</title>
<!-- #include virtual="gerencial_foto/includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include virtual="gerencial_foto/includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include virtual="gerencial_foto/files/includes/scripts.asp" -->
<!-- #include virtual="gerencial_foto/files/includes/functions_local.asp" -->

</head>
<body style="padding:10px; background-color:#FFF">
<%

'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 4/5/2011 21:59:58

'Declara vari�veis para utiliza��o do sistema
Dim id, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario, status_ordem_servico

redirecionar_incluir	= caminho_pagina
redirecionar_alterar	= caminho_pagina
redirecionar_excluir	= caminho_pagina
redirecionar_validacao	= caminho_pagina
msgAoExcluir			= "Tem certeza que deseja excluir este Histórico da Ordem de Serviço?"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
cod_ordem_servico = Request.QueryString("cod_ordem_servico")
'Declara vari�veis
Dim cod_ordem_servico, descricao, cod_usuario, modificado

If comando = Empty Then
comando = "incluir"
titulo_pagina = titulo_pagina&" / Inclusão"
End if

If comando = "incluir" and campo1 <> Empty Then
'Incluir novo registro
incluirRegistro(redirecionar_incluir)
titulo_pagina = titulo_pagina&" / Inclusão"
Elseif comando = "alterar" and campo1 = Empty Then
'Leitura inicial para poss�vel Alteração
id = Request.QueryString("id")
lerRegistro()
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "alterar" and campo1 <> Empty Then
'Gravar Alteração
id = Request.Form("id")
alterarRegistro(redirecionar_alterar)
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 
'Leitura inicial para poss�vel exclus�o
id = Request.QueryString("id")
lerRegistro()
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
'Exclui registro
id = Request.Form("id")
excluirRegistro(redirecionar_excluir)
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
'Visualiza registro
id = Request.QueryString("id")
lerRegistro()
titulo_pagina = titulo_pagina&" / Visualização"
End If


Sub lerFormulario()

id = Request.Form("id")
cod_ordem_servico = Request.Form("cod_ordem_servico")
descricao = Request.Form("descricao")
cod_usuario = Session("cod_usuario")
data_cadastro = Request.Form("data_cadastro")

End sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM TB_ORDENS_SERVICO_HISTORICO WHERE id = "&id)

				cod_ordem_servico=trim(rs("cod_ordem_servico"))
				descricao=trim(rs("descricao"))
				cod_usuario=trim(rs("cod_usuario"))
				data_cadastro=trim(rs("data_cadastro"))

		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
	
		
		'////
		'Verifica qual � a etapa atual de acordo com a base Temp_OrdemServico
		Dim etapa_atual
		etapa_atual = 1 'Por seguran�a etapa 1 caso NÃO tiver etapa atual ainda!
		
		Set rs=objConexao.Execute("SELECT cod_ordem_servico_etapa_atual FROM TB_ORDENS_SERVICO WHERE cod_ordem_servico = '"&cod_ordem_servico&"' ")
			
			If not rs.Eof Then
				etapa_atual = rs("cod_ordem_servico_etapa_atual")
               
			End if

		rs.Close
		Set rs=Nothing
        
		objConexao.Execute("INSERT INTO TB_ORDENS_SERVICO_HISTORICO (cod_ordem_servico, descricao, cod_ordem_servico_etapa,cod_usuario) VALUES ("&sqlNumero(cod_ordem_servico)&", "&sqlTexto(descricao)&", "&sqlTexto(etapa_atual)&", "&sqlNumero(cod_usuario)&")")

		'Encontrar chave criada
		Set rs=objConexao.Execute("SELECT @@IDENTITY AS 'ULTIMO_ID'")
			If not rs.Eof Then
				id = rs("ULTIMO_ID")
			End if 
		rs.Close
		Set rs=Nothing


	FecharConexao()
	Response.Redirect(redirecionar&"?cod_ordem_servico="&cod_ordem_servico&"&id="&id&"&codmsg=1")
End sub

Sub alterarRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("UPDATE TB_ORDENS_SERVICO_HISTORICO SET cod_ordem_servico = "&sqlNumero(cod_ordem_servico)&", descricao = "&sqlTexto(descricao)&", cod_usuario = "&sqlNumero(cod_usuario)&" WHERE id = "&id)
	FecharConexao()
	Response.Redirect(redirecionar&"?id="&id&"&codmsg=2")
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()
		'call verificaAntesExcluir(,id,"TABELASRELACIONADAS;TABELASRELACIONADAS","id",redirecionar_excluir&"?sucesso=2&codmsg=5")

	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_ORDENS_SERVICO_HISTORICO WHERE id = "&id)
	FecharConexao()
	Response.Redirect(redirecionar&"?id="&id&"&codmsg=3&cod_ordem_servico="&cod_ordem_servico)
End sub



'----------------------------------------------------------------------------------------

%>
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

<form action='<%=actionformulario%>?comando=<%=comando%>&cod_ordem_servico=<%=cod_ordem_servico%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post'>
<%
AbrirConexao()

	'///////////////////
	Set rs=objConexao.Execute("Select status_ordem_servico FROM TB_ORDENS_SERVICO Where cod_ordem_servico = '"&cod_ordem_servico&"'")
			
			If not rs.Eof Then
				status_ordem_servico = rs("status_ordem_servico")
			End If
	
	rs.Close
	Set rs=Nothing
	'///////////////////


	'Histórico
	Dim qtd_historicos
	qtd_historicos = 0
	
    
	Set rs=objConexao.Execute("Select TB_ORDENS_SERVICO_HISTORICO.*, TB_ORDENS_SERVICO_ETAPAS.numero_etapa From TB_ORDENS_SERVICO_ETAPAS inner join TB_ORDENS_SERVICO_HISTORICO on (TB_ORDENS_SERVICO_ETAPAS.cod_ordem_servico_etapa = TB_ORDENS_SERVICO_HISTORICO.cod_ordem_servico_etapa) Where TB_ORDENS_SERVICO_HISTORICO.cod_ordem_servico = '"&cod_ordem_servico&"' order by TB_ORDENS_SERVICO_HISTORICO.data_cadastro desc")
	%>
	
    <table width="96%" border="1" align="center" cellpadding="3" style="border-collapse:collapse">
<%
			If int(status_ordem_servico) = 1 or int(status_ordem_servico) = 2 Then
			'Pode alterar ou excluir somente se o status da OS for em andamento ou iniciada, caso contr�rio NÃO ser� permitido.
			%>
            <tr> 
              <td width="82%"><div align='center'>
<table border='0' align="center" cellpadding='0' cellspacing='0' style='width:480px;'>
<tr>
<td width='57'><textarea rows='3' name='DESCRICAO' style='width:480px;' class='inputTamanho3' onkeydown='textCounter(this.form.DESCRICAO,this.form.charRestantes_DESCRICAO,3000);' onkeyup='textCounter(this.form.DESCRICAO,this.form.charRestantes_DESCRICAO,3000);'><%=DESCRICAO%></textarea></td>
</tr>
<tr>
<td><table width='50' border='0' align='right' cellpadding='0' cellspacing='0'>
<tr>
<td><div align='center'>
<input class='inputContaCaracteres' tabindex='10000' type='text' name='charRestantes_DESCRICAO' value="<%=3000- Len(DESCRICAO)%>" />
</div></td>
</tr>
</table></td>
</tr>
</table>
</div></td>
              <td width="18%"><div align="center">
              
                        
              	<%
				If comando = "excluir" Then
				%>
                <a href="javascript:parent.history.go(-1)"><img src="images/layout/btnVoltar.jpg" /></a>
              	&nbsp;
                <a href="javascript:salvar();"><img src="images/layout/btnExcluir.jpg" /></a>
                <%Else%>
                <a href="javascript:salvar();"><img src="images/layout/btnGravar.jpg" /></a>
                <%End If%>
             </div></td>
			</tr>
<%
			End If
			
			While not rs.Eof
			
			DESCRICAO = rs("DESCRICAO")
			DESCRICAO = Replace(DESCRICAO,Chr(13),"<br />")
			qtd_historicos = qtd_historicos + 1
			
			%>
				
			<tr>
              <td colspan="2"><div align="left">
              
              <div style="padding:5px; background-color:#E5E5E5">
             <%=rs("numero_etapa")%>&ordf;) Etapa - <%=nomeUsuario(rs("cod_usuario"),true)%> - <%=rs("data_cadastro")%>:<br />
              </div>
<table width="100%" border="0" cellspacing="0" cellpadding="4">
  <tr>
    <td width="95%"><%=DESCRICAO%></td>
    <td width="5%"><div align="center">
    <%
	If int(rs("cod_usuario")) = Session("cod_usuario") and (int(status_ordem_servico) = 1 or int(status_ordem_servico) = 2) Then
	'S� permite exclus�o pelo pr�prio usu�rio que cadastrou o Histórico e se a OS estiver com Status Iniciada ou Em Andamento
	%>
    <a href="javascript:void(0);" onclick="abre('ppv_ordens_servico_formulario_consulta_historico.asp?comando=excluir&id=<%=rs(0)%>&cod_ordem_servico=<%=cod_ordem_servico%>')">
    <img src="images/icons/icon_excluir.png" width="16" height="16" border="0" />    </a>
    <%
	End If
	%>
    
    
    </div></td>
  </tr>
</table>


</div></td>
      </tr>	
				
			<%
			rs.MoveNext()
			Wend
			%>
            </table>
            
     
     <%If qtd_historicos = 0 and int(status_ordem_servico) = 3 Then%>
     <div class="msgGrande">NÃO Histórico nesta OS cadastrado!</div>
     <%End If%>
     
        
<br />
            <span style="font-style: italic">Observações:</span><br />
            <br />
            <span style="font-weight: bold">Inserir hist&oacute;rico:</span> Se a OS n&atilde;o estiver finalizada ou cancelada poder&aacute; ser inserido um hist&oacute;rico.<br />
            <br />
            <span style="font-weight: bold">Excluir hist&oacute;rico:</span> Somente por quem incluiu o hist&oacute;rico e se a OS n&atilde;o estiver finalizada ou cancelada.<br />
            <br />
<br />

            
<%	rs.Close
	Set rs=Nothing
	
FecharConexao()
%>
<input type='hidden' value="<%=id%>" name='id'  />
<input type='hidden' value='1' name='campo1'  />
<input type='hidden' name='cod_ordem_servico' value="<%=cod_ordem_servico%>" />
</form>
<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('DESCRICAO#Histórico#1',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>
 
<script type="text/javascript">
function salvar(){
	if(validaFormMaster()){
		document.FORMULARIO_PADRAO.submit();
	}
}
</script>
<!-- #include virtual="gerencial_foto/files/includes/scripts_rodape.asp" -->
</body>
</html>
