<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
If Request.QueryString("limpar_agendadas") = "sim" Then

AbrirConexao()
 	objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_ETIQUETAS_TEMPORARIA WHERE COD_USUARIO="&Session("COD_USUARIO")&" AND AGENDADA=1")
FecharConexao()
	Response.Redirect("adm_etiquetas_agendadas.asp?msg=Etiquetas Agendadas deletadas com sucesso!")
End If


%>
<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="iso-8859-1">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Etiquetas agendadas | Portal Stillos</title>
  <link rel="stylesheet" href="css/vendor/bootstrap.min.css">
  <link rel="stylesheet" href="css/dashboard.css">
  <style>
    .queue-page .dashboard-nav a.active:not([aria-current]) { background:transparent; color:#d4d5e1; }
    .queue-page .dashboard-nav a.active:hover { background:#37394f; color:#fff; }
    .queue-page .queue-panel { height:auto; max-width:960px; margin-inline:auto; padding:clamp(18px, 3vw, 32px); }
    .queue-panel .panel-heading { margin-bottom:20px; }
    .queue-actions { display:flex; flex-wrap:wrap; gap:8px; }
    .queue-actions .btn-outline-danger { --bs-btn-color:#982542; --bs-btn-border-color:#e4b9c5; }
    .queue-total { display:inline-flex; align-items:center; margin-top:7px; color:#61425e; background:#f5edf4; border-radius:6px; padding:4px 9px; font-size:12px; font-weight:650; }
    .queue-table { margin:0; --bs-table-color:var(--ink); --bs-table-border-color:var(--line); }
    .queue-table th, .queue-table td { padding:12px; overflow-wrap:anywhere; }
    .queue-table thead th { background:#f5f3f8; color:#42465a; font-size:12px; }
    .queue-table td:last-child { font-variant-numeric:tabular-nums; }
    .queue-table a { font-weight:650; text-underline-offset:3px; }
    .queue-panel .empty-state { margin:0; }
    @media(max-width:480px) {
      .queue-actions, .queue-actions .btn { width:100%; }
      .queue-table th, .queue-table td { padding:10px 7px; }
    }
    .queue-dialog { width:min(440px, calc(100% - 32px)); padding:24px; border:1px solid var(--line); border-radius:14px; color:var(--ink); background:#fff; box-shadow:0 18px 60px #24273833; }
    .queue-dialog::backdrop { background:#24273899; }
    .queue-dialog p { margin:10px 0 20px; }
    .queue-dialog-actions { display:flex; justify-content:flex-end; flex-wrap:wrap; gap:8px; }
  </style>
</head>
<body class="queue-page">
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
      <div class="topbar-title">Portal Stillos <span>/ Etiquetas agendadas</span></div>
      <div class="topbar-actions"><a class="user-name" href="adm_log_usuarios.asp"><%=Server.HTMLEncode(Session("NOME") & "")%></a><a href="manual_sistema.asp" target="_blank" rel="noopener">Manual</a><a class="btn btn-outline-secondary btn-sm" href="adm_sair.asp">Sair</a></div>
    </header>
    <main class="dashboard-main ppv-main" id="conteudo">
      <div class="dashboard-heading">
        <div>
          <h1>Etiquetas agendadas</h1>
          <p>Consulte e imprima as etiquetas da sua fila.</p>
        </div>
        <a class="btn btn-outline-secondary" id="queueBack" href="adm_index.asp">Voltar</a>
      </div>
      <div class="search-panel"><strong>Busca no portal</strong><!-- #include file="includes/dashboard_busca.asp" --></div>
      <%If Len(Request.QueryString("msg") & "") > 0 Then%>
      <p class="alert alert-success" role="status"><%=Server.HTMLEncode(Request.QueryString("msg") & "")%></p>
      <%End If%>
      <%
      Dim rsFilaEtiquetas, totalFilaEtiquetas
      AbrirConexao()
      Set rsFilaEtiquetas = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO_ETIQUETAS_TEMPORARIA WHERE COD_USUARIO="&Session("COD_USUARIO")&" AND AGENDADA=1",totalFilaEtiquetas)
      %>
      <section class="dashboard-panel queue-panel" aria-labelledby="queueTitle">
        <div class="panel-heading">
          <div>
            <h2 id="queueTitle">Fila de impress&atilde;o</h2>
            <%If Not rsFilaEtiquetas.Eof Then%>
            <span class="queue-total">Total de etiquetas: <%=totalFilaEtiquetas%></span>
            <%End If%>
          </div>
          <%If Not rsFilaEtiquetas.Eof Then%>
          <div class="queue-actions">
            <a class="btn btn-primary" id="printLabels" href="dip_formandos_etiquetas.asp" target="Etiquetas">Imprimir etiquetas</a>
            <button class="btn btn-outline-danger" id="clearLabels" type="button" aria-haspopup="dialog">Limpar fila</button>
            <dialog class="queue-dialog" id="clearLabelsDialog" aria-labelledby="clearLabelsTitle" aria-describedby="clearLabelsDescription">
              <h2 id="clearLabelsTitle">Limpar etiquetas agendadas?</h2>
              <p id="clearLabelsDescription">Todas as etiquetas da sua fila ser&atilde;o removidas. Confirme se a impress&atilde;o j&aacute; foi conclu&iacute;da.</p>
              <form class="queue-dialog-actions" method="dialog">
                <button class="btn btn-outline-secondary" value="cancelar" autofocus>Cancelar</button>
                <button class="btn btn-danger" value="limpar">Limpar fila</button>
              </form>
            </dialog>
          </div>
          <%End If%>
        </div>
        <%If rsFilaEtiquetas.Eof Then%>
        <p class="empty-state">Nenhuma etiqueta agendada.</p>
        <%Else%>
        <table class="table queue-table" aria-labelledby="queueTitle">
          <thead>
            <tr><th scope="col">C&oacute;d. do formando</th><th scope="col">Data do agendamento</th></tr>
          </thead>
          <tbody>
            <%While Not rsFilaEtiquetas.Eof%>
            <tr>
              <td><a href="dip_busca_detalhamento_formando.asp?cod_formando=<%=Server.URLEncode(rsFilaEtiquetas("COD_FORMANDO") & "")%>"><%=Server.HTMLEncode(rsFilaEtiquetas("COD_FORMANDO") & "")%></a></td>
              <td><%=Server.HTMLEncode(rsFilaEtiquetas("DATA_CADASTRO") & "")%></td>
            </tr>
            <%
              rsFilaEtiquetas.MoveNext()
            Wend
            %>
          </tbody>
        </table>
        <%End If%>
      </section>
      <%
      rsFilaEtiquetas.Close
      Set rsFilaEtiquetas = Nothing
      FecharConexao()
      %>
    </main>
    <footer class="dashboard-footer">&copy; <%=Year(Date)%> Grupo Stillos <span>Sistema administrativo interno</span></footer>
  </div>
</div>
<script>
(function () {
    var toggle = document.getElementById('ppvMenuToggle');
    var menu = document.getElementById('menuPrincipal');
    var previous = menu.querySelector('[aria-current="page"]');
    if (previous) previous.removeAttribute('aria-current');
    var printLink = document.getElementById('printLabels');
    var clearButton = document.getElementById('clearLabels');
    var backLink = document.getElementById('queueBack');
    if (printLink) {
        printLink.addEventListener('click', function (event) {
            if (event.ctrlKey || event.metaKey || event.shiftKey || event.altKey) return;
            var popup = window.open(printLink.href, printLink.target, 'width=200,height=200,scrollbars=no');
            if (popup) event.preventDefault();
        });
    }
    if (clearButton) {
        var clearDialog = document.getElementById('clearLabelsDialog');
        clearButton.addEventListener('click', function () {
            clearDialog.returnValue = '';
            clearDialog.showModal();
        });
        clearDialog.addEventListener('close', function () {
            if (clearDialog.returnValue === 'limpar') {
                window.location.assign('?limpar_agendadas=sim');
            }
        });
    }
    backLink.addEventListener('click', function (event) {
        if (event.ctrlKey || event.metaKey || event.shiftKey || event.altKey || !document.referrer) return;
        var previousPage = new URL(document.referrer);
        if (previousPage.origin === window.location.origin && previousPage.pathname !== window.location.pathname) {
            event.preventDefault();
            window.history.back();
        }
    });
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
</script>
</body>
</html>
