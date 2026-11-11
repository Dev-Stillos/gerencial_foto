<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="files/includes/functions_local.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<%
Dim caminho_pagina
caminho_pagina = Replace(Request.ServerVariables("URL"), "/gerencial_foto/", "")
Dim cod_pedido, sucesso, codmsg
cod_pedido = Cstr(""&Request.Form("cod_pedido"))
cod_pedido = Replace(cod_pedido,".","")
cod_pedido = Replace(cod_pedido,"/","")
cod_pedido = Replace(cod_pedido," ","")
cod_pedido = Replace(cod_pedido,"-","")

	If cod_pedido <> "" Then
		
		AbrirConexao()
		
		Set rs=objConexao.Execute("SELECT cod_pedido FROM TB_PEDIDOS WHERE cod_pedido = "&sqlTexto(cod_pedido)&" and FINALIZA_CORRECAO = 1")
			if not rs.Eof then
				
				retorna = moduloPPVgerarOrdensServico(cod_pedido,true)
				'Response.Write(retorna)
				If int(retorna) = 1 Then
					sucesso = ""
					codmsg = 1	
				Else
					sucesso = 2
					codmsg = retorna
				End If				

			eLSE
				sucesso = 2
				codmsg = 100
			end if
		rs.close
		Set rs=nothing	
		
		FecharConexao()
		
		Response.Redirect("ppv_ordens_servico_os_manual_origem_vendas.asp?cod_pedido="&Server.URLEncode(cod_pedido)&"&codmsg="&codmsg&"&sucesso="&sucesso)
		
	End if

		
'----------------------------------------------------------------------------------------

%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Gera&ccedil;&atilde;o manual de OS | Portal Stillos</title>
<link rel="stylesheet" href="css/vendor/bootstrap.min.css">
<link rel="stylesheet" href="css/dashboard.css">
<style>
.manual-os-page .dashboard-nav a.active:not([aria-current]) { background:transparent; color:#d4d5e1; }
.manual-os-page .dashboard-nav a[aria-current="page"] { background:#b20b61; color:#fff; font-weight:600; }
.manual-os-page .ppv-module { margin-bottom:20px; }
.manual-os-page .subMenu { padding:12px 16px; background:#fff; border:1px solid var(--line); border-radius:10px; line-height:2; }
.manual-os-page .subMenu [aria-current="page"] { color:#88134f; font-weight:700; }
.manual-os-grid { display:grid; grid-template-columns:minmax(0,1.15fr) minmax(0,1fr); gap:18px; }
.manual-os-grid .dashboard-panel { height:auto; padding:clamp(16px,3vw,28px); }
.manual-os-form { margin-top:18px; }
.manual-os-form label { font-weight:650; }
.manual-os-actions { display:flex; justify-content:flex-end; flex-wrap:wrap; gap:8px; margin-top:16px; }
.manual-os-rules ol { padding-left:22px; margin:14px 0 0; }
.manual-os-rules li + li { margin-top:10px; }
.manual-os-dialog { box-sizing:border-box; width:min(480px,calc(100% - 24px)); padding:24px; border:1px solid var(--line); border-radius:14px; color:var(--ink); }
.manual-os-dialog::backdrop { background:#24273899; }
.manual-os-dialog p { margin:10px 0 0; overflow-wrap:anywhere; }
@media(max-width:980px) { .manual-os-grid { grid-template-columns:minmax(0,1fr); } }
@media(max-width:420px) { .manual-os-actions .btn { width:100%; } }
</style>
</head>
<body class="manual-os-page">
<a class="skip-link" href="#conteudo">Ir para o conte&uacute;do</a>
<div class="dashboard-layout">
  <aside class="dashboard-sidebar" id="menuPrincipal">
    <a class="dashboard-brand" href="adm_index.asp"><img src="images/logoparafundopreto.png" alt="Stillos" width="180" /><span>GERENCIAL FOTOS</span></a>
    <div class="nav-caption">&Aacute;REA DE TRABALHO</div>
    <!-- #include file="includes/dashboard_menu.asp" -->
  </aside>
  <div class="dashboard-workspace">
    <header class="dashboard-topbar">
      <button class="btn menu-toggle" id="ppvMenuToggle" type="button" aria-controls="menuPrincipal" aria-expanded="false">Menu</button>
      <div class="topbar-title">Portal Stillos <span>/ Produ&ccedil;&atilde;o p&oacute;s-venda</span></div>
      <div class="topbar-actions"><a class="user-name" href="adm_log_usuarios.asp"><%=Server.HTMLEncode(Session("NOME") & "")%></a><a href="manual_sistema.asp" target="_blank" rel="noopener">Manual</a><a class="btn btn-outline-secondary btn-sm" href="adm_sair.asp">Sair</a></div>
    </header>
    <main class="dashboard-main ppv-main" id="conteudo">
      <div class="dashboard-heading">
        <div><p class="eyebrow">PRODU&Ccedil;&Atilde;O P&Oacute;S-VENDA</p><h1>Gerar OS a partir de uma venda</h1><p>Informe o c&oacute;digo da venda para gerar as ordens de servi&ccedil;o eleg&iacute;veis.</p></div>
        <a href="ppv_ordens_servico.asp" class="btn btn-outline-secondary">Voltar &agrave;s ordens</a>
      </div>
      <div class="search-panel"><div><strong>Busca no portal</strong></div><!-- #include file="includes/dashboard_busca.asp" --></div>
      <div class="session-notices"><!-- #include file="includes/dashboard_avisos.asp" --></div>
      <nav class="ppv-module" aria-label="Produ&ccedil;&atilde;o p&oacute;s-venda">
        <!-- #include file="includes/classe_modulo_ppv_menu.asp" -->
      </nav>
      <%
      Dim vendaExibida, mensagemGeracaoOS
      codmsg = CStr("" & Request.QueryString("codmsg"))
      sucesso = CStr("" & Request.QueryString("sucesso"))
      vendaExibida = Server.HTMLEncode(Request.QueryString("cod_pedido") & "")
      If sucesso = "" And codmsg = "1" Then
      %>
      <p class="alert alert-success" role="status">Ordens de servi&ccedil;o geradas para a venda <strong><%=vendaExibida%></strong>.</p>
      <%ElseIf sucesso <> "" And codmsg <> "" Then%>
      <div class="alert alert-warning" role="status">
        <%Select Case codmsg
          Case "100"%>
          Pedido <strong><%=vendaExibida%></strong> indispon&iacute;vel. Confira o c&oacute;digo e se a corre&ccedil;&atilde;o da venda foi finalizada.
        <%Case "101"%>
          J&aacute; existem ordens de servi&ccedil;o para o pedido <strong><%=vendaExibida%></strong>. Para gerar novamente, cancele as OS da venda ou exclua-as quando permitido.
        <%Case "1", "2", "3"
            mensagemGeracaoOS = moduloPPVgerarOrdensServicoMsgRetorno(codmsg, "")
            Response.Write Server.HTMLEncode(mensagemGeracaoOS & "")
          Case Else%>
          A gera&ccedil;&atilde;o das ordens n&atilde;o foi conclu&iacute;da.
        <%End Select%>
      </div>
      <%End If%>
      <div class="manual-os-grid">
        <section class="dashboard-panel" aria-labelledby="manual-os-title">
          <h2 id="manual-os-title">Identificar a venda</h2>
          <form class="manual-os-form" id="manual-os-form" name="FORMULARIO_PADRAO" action="ppv_ordens_servico_os_manual_origem_vendas.asp" method="post">
            <label class="form-label" for="manual-os-sale">C&oacute;digo da venda</label>
            <input class="form-control" id="manual-os-sale" type="text" name="cod_pedido" inputmode="numeric" autocomplete="off" required>
            <div class="manual-os-actions"><button class="btn btn-primary" type="submit">Gerar ordens de servi&ccedil;o</button></div>
          </form>
        </section>
        <section class="dashboard-panel manual-os-rules" aria-labelledby="manual-os-rules">
          <h2 id="manual-os-rules">Regras de gera&ccedil;&atilde;o</h2>
          <ol>
            <li>Ordens j&aacute; existentes impedem uma nova gera&ccedil;&atilde;o. A exce&ccedil;&atilde;o s&atilde;o as ordens canceladas.</li>
            <li>A venda deve estar corrigida pelos respons&aacute;veis.</li>
            <li>A gera&ccedil;&atilde;o se aplica a produtos vendidos &agrave; parte com prazo do cliente cadastrado.</li>
          </ol>
        </section>
      </div>
      <dialog class="manual-os-dialog" id="manual-os-dialog" aria-labelledby="manual-os-confirm-title" aria-describedby="manual-os-confirm-sale">
        <h2 id="manual-os-confirm-title">Confirmar gera&ccedil;&atilde;o de OS?</h2>
        <p id="manual-os-confirm-sale"></p>
        <form class="manual-os-actions" method="dialog">
          <button class="btn btn-outline-secondary" value="cancelar" autofocus>Cancelar</button>
          <button class="btn btn-primary" value="gerar">Confirmar gera&ccedil;&atilde;o</button>
        </form>
      </dialog>
    </main>
    <footer class="dashboard-footer">&copy; <%=Year(Date)%> Grupo Stillos <span>Sistema administrativo interno</span></footer>
  </div>
</div>
<script>
(function () {
    var toggle = document.getElementById('ppvMenuToggle');
    var menu = document.getElementById('menuPrincipal');
    var previous = menu.querySelector('[aria-current="page"]');
    var current = menu.querySelector('a[href="ppv_ordens_servico.asp"]');
    if (previous) previous.removeAttribute('aria-current');
    if (current) current.setAttribute('aria-current', 'page');
    toggle.addEventListener('click', function () {
        var expanded = menu.classList.toggle('is-open');
        toggle.setAttribute('aria-expanded', String(expanded));
    });
    document.addEventListener('keydown', function (event) {
        if (event.key === 'Escape' && menu.classList.contains('is-open')) {
            menu.classList.remove('is-open');
            toggle.setAttribute('aria-expanded', 'false');
            toggle.focus();
        }
    });
})();

(function () {
    var submenu = document.querySelector('.ppv-module a[href="ppv_ordens_servico_os_manual_origem_vendas.asp"]');
    if (submenu) submenu.setAttribute('aria-current', 'page');
    var form = document.getElementById('manual-os-form');
    var sale = document.getElementById('manual-os-sale');
    var dialog = document.getElementById('manual-os-dialog');
    if (typeof dialog.showModal !== 'function') return;
    form.addEventListener('submit', function (event) {
        var code = sale.value.replace(/[. /-]/g, '');
        if (!code) {
            event.preventDefault();
            sale.value = '';
            sale.reportValidity();
            return;
        }
        event.preventDefault();
        document.getElementById('manual-os-confirm-sale').textContent = 'Venda ' + code;
        dialog.returnValue = '';
        dialog.showModal();
    });
    dialog.addEventListener('close', function () {
        if (dialog.returnValue === 'gerar') form.submit();
    });
})();
</script>
</body>
</html>
