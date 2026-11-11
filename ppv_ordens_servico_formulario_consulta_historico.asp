			
<%
Dim calendario_necessario
calendario_necessario = "NÃO"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<link rel="stylesheet" media="all" type="text/css" href="css/modal-message.css">
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Hist&oacute;rico da ordem de servi&ccedil;o | Portal Stillos</title>
<style>
body.os-history { margin:0; padding:clamp(12px,2vw,20px); background:#fff; color:#292738; font:14px/1.5 system-ui,sans-serif; }
.os-history *, .os-history *::before, .os-history *::after { box-sizing:border-box; }
.os-history .history-compose, .os-history .history-entry { padding:14px; border:1px solid #e9e5ee; border-radius:10px; background:#faf8fc; }
.os-history .history-compose label { display:block; margin-bottom:7px; font-weight:700; }
.os-history .history-compose textarea { display:block; width:100%; min-height:84px; resize:vertical; padding:10px; border:1px solid #d9d2e1; border-radius:8px; background:#fff; color:inherit; font:inherit; }
.os-history .history-compose textarea:focus { outline:2px solid #9e125c; outline-offset:1px; }
.os-history .history-actions { display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px; margin-top:10px; }
.os-history .history-counter { color:#615b6e; font-size:12px; }
.os-history .history-counter input { width:4ch; border:0; padding:0; background:transparent; text-align:right; color:inherit; font:inherit; }
.os-history .history-button { display:inline-block; padding:7px 12px; border:1px solid #9e125c; border-radius:8px; background:#9e125c; color:#fff; text-decoration:none; font:650 13px/1.4 system-ui,sans-serif; cursor:pointer; }
.os-history .history-button:hover { background:#80104b; color:#fff; }
.os-history .history-secondary { border-color:#ddd5e5; background:#fff; color:#51445d; }
.os-history .history-danger { border-color:#edccd6; background:#fff2f5; color:#8d2543; }
.os-history .history-entry { margin-top:10px; }
.os-history .history-meta { display:flex; flex-wrap:wrap; gap:4px 12px; margin-bottom:8px; color:#615b6e; font-size:12px; }
.os-history .history-meta strong { color:#3e344e; }
.os-history .history-text { overflow-wrap:anywhere; }
.os-history .history-entry .history-button { margin-top:8px; }
.os-history .history-empty { margin:12px 0; color:#615b6e; }
.os-history .history-help { margin-top:12px; color:#615b6e; font-size:12px; }
.os-history .history-help summary { cursor:pointer; font-weight:650; }
</style>
<!-- #include virtual="gerencial_foto/includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include virtual="gerencial_foto/includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include virtual="gerencial_foto/files/includes/scripts.asp" -->
<!-- #include virtual="gerencial_foto/files/includes/functions_local.asp" -->

</head>
<body class="os-history">
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
	
<% If int(status_ordem_servico) = 1 or int(status_ordem_servico) = 2 Then %>
    <div class="history-compose">
        <label for="os-history-text">Anota&ccedil;&atilde;o da OS</label>
        <textarea id="os-history-text" rows="3" name="DESCRICAO" onkeydown="textCounter(this.form.DESCRICAO,this.form.charRestantes_DESCRICAO,3000);" onkeyup="textCounter(this.form.DESCRICAO,this.form.charRestantes_DESCRICAO,3000);"><%=Server.HTMLEncode(DESCRICAO & "")%></textarea>
        <div class="history-actions">
            <label class="history-counter">
                <input type="text" name="charRestantes_DESCRICAO" readonly="readonly" value="<%=3000 - Len(DESCRICAO)%>" /> caracteres restantes
            </label>
            <div>
            <% If comando = "excluir" Then %>
                <a class="history-button history-secondary" href="javascript:parent.history.go(-1)">Voltar</a>
                <button type="submit" class="history-button history-danger">Confirmar exclus&atilde;o</button>
            <% Else %>
                <button type="submit" class="history-button">Gravar</button>
            <% End If %>
            </div>
        </div>
    </div>
<% End If

            While not rs.Eof
			
			DESCRICAO = rs("DESCRICAO")
			DESCRICAO = Replace(DESCRICAO,Chr(13),"<br />")
			qtd_historicos = qtd_historicos + 1
			
			%>
				
    <article class="history-entry">
        <div class="history-meta">
            <strong><%=rs("numero_etapa")%>&ordf; etapa</strong>
            <span><%=nomeUsuario(rs("cod_usuario"),true)%></span>
            <span><%=rs("data_cadastro")%></span>
        </div>
        <div class="history-text"><%=DESCRICAO%></div>
        <% If int(rs("cod_usuario")) = Session("cod_usuario") and (int(status_ordem_servico) = 1 or int(status_ordem_servico) = 2) Then %>
            <a class="history-button history-danger" href="ppv_ordens_servico_formulario_consulta_historico.asp?comando=excluir&amp;id=<%=rs(0)%>&amp;cod_ordem_servico=<%=cod_ordem_servico%>">Excluir anota&ccedil;&atilde;o</a>
        <% End If %>
    </article>
    <%
            rs.MoveNext()
			Wend
			%>
    <% If qtd_historicos = 0 Then %>
        <p class="history-empty">Nenhuma anota&ccedil;&atilde;o registrada nesta OS.</p>
    <% End If %>

    <details class="history-help">
        <summary>Regras do hist&oacute;rico</summary>
        <p>O registro de anota&ccedil;&otilde;es est&aacute; dispon&iacute;vel enquanto a OS est&aacute; em andamento ou aguardando despacho. A exclus&atilde;o &eacute; permitida ao autor da anota&ccedil;&atilde;o nesses mesmos status.</p>
    </details>

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
 

<!-- #include virtual="gerencial_foto/files/includes/scripts_rodape.asp" -->
</body>
</html>
