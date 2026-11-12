<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Baixa para a diretoria | Portal Stillos</title>
<link rel="stylesheet" href="css/vendor/bootstrap.min.css">
<link rel="stylesheet" href="css/dashboard.css">
<style>
.cap-metrics { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; margin:0 0 16px; }
.cap-metric { min-width:0; padding:16px; border:1px solid #e4dfeb; border-radius:12px; background:#fff; }
.cap-metric dt { color:#57566a; font-size:.87em; font-weight:650; }
.cap-metric dd { margin:6px 0 0; font-size:clamp(19px,2vw,27px); line-height:1.15; font-weight:700; font-variant-numeric:tabular-nums; }
.cap-metric small { display:block; margin-top:6px; color:#57566a; }
.cap-overview-charts { display:grid; grid-template-columns:minmax(0,2fr) minmax(0,1fr); gap:12px; margin:0 0 18px; }
.cap-chart { min-width:0; padding:16px; border:1px solid #e4dfeb; border-radius:12px; background:#fff; }
.cap-chart h3 { margin:0 0 4px; color:#242738; font:700 1.05em/1.3 system-ui,sans-serif; }
.cap-chart p { margin:0 0 12px; color:#57566a; font-size:.9em; }
.cap-chart-bars { list-style:none; margin:0; padding:0; }
.cap-chart-bars li + li { margin-top:8px; }
.cap-chart-bars li > div:first-child { display:flex; justify-content:space-between; gap:8px; font-variant-numeric:tabular-nums; }
.cap-chart-track { height:7px; overflow:hidden; margin-top:3px; border-radius:7px; background:#eee9f4; }
.cap-chart-track span { display:block; height:100%; background:#a31962; }
.cap-chart-donut { display:grid; place-items:center; width:124px; aspect-ratio:1; margin:14px auto 10px; border-radius:50%; background:conic-gradient(#a31962 var(--cap-share),#eee9f4 0); }
.cap-chart-donut strong { display:grid; place-items:center; width:82px; aspect-ratio:1; border-radius:50%; background:#fff; }
.cap-chart .cap-chart-share { margin:0; text-align:center; }
.period-overview { margin-bottom:20px; }
.period-overview > h2 { margin-bottom:10px; }
.period-actions,.period-paging,.period-filter { display:flex; flex-wrap:wrap; align-items:center; gap:8px; }
.period-panel { height:auto; }
.period-panel header { display:flex; flex-wrap:wrap; justify-content:space-between; gap:12px; align-items:end; margin-bottom:12px; }
.period-scroll { overflow-x:auto; border:1px solid #e4dfeb; border-radius:9px; }
.period-table { width:100%; min-width:760px; margin:0; border-collapse:collapse; font-variant-numeric:tabular-nums; }
.period-table th,.period-table td { padding:9px; border-bottom:1px solid #eeeaf2; text-align:left; vertical-align:middle; }
.period-table th { background:#f1edf6; }
.period-table tbody tr:nth-child(even) { background:#faf9fc; }
.period-date { white-space:nowrap; }
.period-meta { display:block; margin-top:3px; font-size:.9em; color:#53566b; }
.period-badge { display:inline-block; padding:3px 7px; border-radius:6px; font-weight:650; background:#faf0dd; color:#754600; }
.period-done { background:#e5f2ea; color:#275e42; }
.period-paging { justify-content:center; margin-top:14px; }
.cap-chart-donut { margin-left:0; }
@media(max-width:850px) { .cap-metrics { grid-template-columns:repeat(2,minmax(0,1fr)); } }
@media(max-width:460px) { .cap-metrics { grid-template-columns:minmax(0,1fr); } }

.dispatch-card { height:auto; margin:0 0 16px; min-width:0; }
.dispatch-head,.dispatch-toolbar { display:flex; flex-wrap:wrap; justify-content:space-between; align-items:center; gap:10px; margin-bottom:12px; }
.dispatch-table { min-width:830px; }
.dispatch-table th:first-child,.dispatch-table td:first-child { text-align:center; }
.dispatch-table td:nth-child(2),.dispatch-table td:nth-child(n+4) { white-space:nowrap; }
.dispatch-table input[type="checkbox"] { accent-color:#b20b61; width:16px; height:16px; }
.dispatch-person { display:flex; flex-wrap:wrap; gap:4px 8px; }
.dispatch-person + .dispatch-person { margin-top:4px; }
.dispatch-person-code { color:#53566b; white-space:nowrap; }
.dispatch-bar { position:sticky; top:0; z-index:2; }
.dispatch-bar .dispatch-toolbar { margin:0; }
.dispatch-dialog { width:min(460px,calc(100% - 24px)); padding:20px; border:1px solid #e4dfeb; border-radius:12px; color:#242738; }
.dispatch-dialog::backdrop { background:#24273899; }
.dispatch-dialog h2 { margin-bottom:10px; }
[hidden] { display:none !important; }
</style>
</head>
<body>
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
      <div class="topbar-title">Portal Stillos <span>/ Baixa para a diretoria</span></div>
      <div class="topbar-actions"><a class="user-name" href="adm_log_usuarios.asp"><%=Server.HTMLEncode(Session("NOME") & "")%></a><a href="manual_sistema.asp" target="_blank" rel="noopener">Manual</a><a class="btn btn-outline-secondary btn-sm" href="adm_sair.asp">Sair</a></div>
    </header>
    <main class="dashboard-main ppv-main" id="conteudo">


      <header class="dashboard-heading">
        <div><p class="eyebrow">PEDIDOS PENDENTES</p><h1>Baixa para a diretoria</h1><p>Selecione os pedidos corrigidos que ser&atilde;o encaminhados.</p></div>
        <a class="btn btn-outline-secondary" href="lca_pedidos_pendentes.asp">Voltar aos pedidos pendentes</a>
      </header>
      <div class="search-panel"><strong>Busca no portal</strong>
        <!-- #include file="includes/dashboard_busca.asp" -->
      </div>
      <div class="session-notices">
        <!-- #include file="includes/dashboard_avisos.asp" -->
      </div>
      <%If Len(Request.QueryString("msg")) > 0 Then%><p class="alert alert-info" role="status"><%=Server.HTMLEncode(Request.QueryString("msg"))%></p><%End If%>
      <form id="dispatch-form" name="baixar_fichas" method="post" action="lca_pedidos_pendentes_baixar_para_diretoria_processar.asp">
        <dl class="cap-metrics">
          <div class="cap-metric"><dt>Pedidos dispon&iacute;veis</dt><dd id="dispatch-total">&mdash;</dd></div>
          <div class="cap-metric"><dt>Vendedores</dt><dd id="dispatch-vendors">&mdash;</dd></div>
          <div class="cap-metric"><dt>Pedidos selecionados</dt><dd id="dispatch-selected" aria-live="polite">0</dd></div>
        </dl>
        <section class="dashboard-panel dispatch-card dispatch-bar" id="dispatch-bar" aria-label="A&ccedil;&otilde;es do lote">
          <div class="dispatch-toolbar">
            <div class="period-actions">
              <button type="button" class="btn btn-outline-secondary" data-dispatch-select="all">Selecionar todos</button>
              <button type="button" class="btn btn-outline-secondary" data-dispatch-select="none">Limpar sele&ccedil;&atilde;o</button>
            </div>
            <button class="btn btn-primary" type="submit" name="btnTopo2" value="1" id="dispatch-submit">Baixar selecionados</button>
          </div>
        </section>
<%
Dim rs, rs2, rs3, rs4, sql, dispatchHasOrders
dispatchHasOrders = False
AbrirConexao()
sql = "SELECT COD_VENDEDOR, NOME_VENDEDOR FROM TB_VENDEDORES ORDER BY NOME_VENDEDOR"
Set rs = objConexao.Execute(sql)
While Not rs.Eof
    sql = "SELECT * FROM VW_PEDIDOS_COMPLETO WHERE COD_VENDEDOR="&rs("COD_VENDEDOR")&" AND DATA_DESCEU_DIRETORIA IS NULL AND FINALIZA_CORRECAO = 1 AND FINALIZA_VENDEDOR = 1 GROUP BY COD_PEDIDO ORDER BY NOME_FORMANDO"
    Set rs2 = objConexao.Execute(sql)
    If Not rs2.Eof Then
        dispatchHasOrders = True
%>
        <section class="dashboard-panel dispatch-card" data-dispatch-group aria-labelledby="dispatch-vendor-<%=rs("COD_VENDEDOR")%>">
          <header class="dispatch-head">
            <h2 id="dispatch-vendor-<%=rs("COD_VENDEDOR")%>"><%=Server.HTMLEncode(rs("NOME_VENDEDOR") & "")%></h2>
            <label><input type="checkbox" data-dispatch-vendor> Selecionar pedidos deste vendedor</label>
          </header>
          <div class="period-scroll" role="region" aria-labelledby="dispatch-vendor-<%=rs("COD_VENDEDOR")%>" tabindex="0">
          <table class="period-table dispatch-table">
            <thead><tr><th scope="col">Baixar</th><th scope="col">Pedido</th><th scope="col">Formandos</th><th scope="col">Total recebido</th><th scope="col">Lan&ccedil;amento</th><th scope="col">Chegada</th><th scope="col">Pagamento do vendedor</th></tr></thead>
            <tbody>
            <%While Not rs2.Eof%>
              <tr>
                <td><input type="checkbox" name="PEDIDO_<%=rs2("COD_PEDIDO")%>" value="1" data-dispatch-order aria-label="Selecionar pedido <%=rs2("COD_PEDIDO")%>"></td>
                <td><strong><%=rs2("COD_PEDIDO")%></strong><%If int(rs2("PEDIDO_EXTRA")) = 1 Then%> <small>(Extra)</small><%End If%></td>
                <td>
<%
sql = 	"SELECT TB_PEDIDOS_FORMANDOS.COD_FORMANDO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO "&_
					"FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON "&_
					"TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO "&_
					"WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&rs2("COD_PEDIDO")&" ORDER BY NOME_FORMANDO"
			Set rs3 = objConexao.Execute(sql)
				
		
				While Not rs3.Eof%>
                  <div class="dispatch-person"><span class="dispatch-person-code"><%=digito(rs3("COD_FORMANDO"))%></span><span><%=Server.HTMLEncode(rs3("NOME_FORMANDO") & "")%></span></div>
                <%rs3.MoveNext
                Wend
                rs3.Close
                Set rs3 = Nothing%>
                </td>
                <td><%
Set rs4 = objConexao.Execute("SELECT * FROM VW_PEDIDOS_TOTAL_RECEBIDO WHERE COD_PEDIDO="&rs2("COD_PEDIDO"))
			Dim PAGAMENTO_EM_DINHEIRO, XTOTAL_CHEQUE, XTOTAL_CARTAO, XTOTAL_BOLETO, TOTAL_PAGO
			
			TOTAL_PAGO = 0
				
				If not rs4.Eof Then
				
					PAGAMENTO_EM_DINHEIRO	= formatNumber(0&rs4("PAGAMENTO_EM_DINHEIRO"))/1
					XTOTAL_CHEQUE			= formatNumber(0&rs4("XTOTAL_CHEQUE"))/1
					XTOTAL_CARTAO			= formatNumber(0&rs4("XTOTAL_CARTAO"))/1
					XTOTAL_BOLETO			= formatNumber(0&rs4("XTOTAL_BOLETO"))/1
					TOTAL_PAGO				= (PAGAMENTO_EM_DINHEIRO + XTOTAL_CHEQUE + XTOTAL_CARTAO + XTOTAL_BOLETO)/1
					
					Response.Write(formatNumber(TOTAL_PAGO))
					
				End If
			
			rs4.Close()
			Set rs4 = Nothing
			%>
            </td>
                <td><%=ArrumaData(rs2("DATA_CADASTRO"))%></td>
                <td><%=ArrumaData(rs2("DATA_CHEGADA"))%></td>
                <td><%If IsDate(rs2("VENDEDOR_RECEBE_EM")) Then%><%=MonthName(Month(rs2("VENDEDOR_RECEBE_EM"))) & " / " & Year(rs2("VENDEDOR_RECEBE_EM"))%><%Else%>&mdash;<%End If%></td>
              </tr>
            <%rs2.MoveNext
            Wend%>
            </tbody>
          </table>
          </div>
        </section>
<%
    End If
    rs2.Close
    Set rs2 = Nothing
    rs.MoveNext
Wend
rs.Close
Set rs = Nothing
FecharConexao()
%>
        <%If Not dispatchHasOrders Then%><p class="empty-state">Nenhum pedido corrigido aguardando baixa para a diretoria.</p><%End If%>
      </form>
      <dialog class="dispatch-dialog" id="dispatch-confirm" aria-labelledby="dispatch-confirm-title">
        <h2 id="dispatch-confirm-title">Confirmar baixa para a diretoria</h2>
        <p id="dispatch-confirm-count"></p>
        <form method="dialog" class="period-actions">
          <button class="btn btn-outline-secondary" value="cancelar" autofocus>Voltar</button>
          <button class="btn btn-primary" value="confirmar">Confirmar baixa</button>
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
    var current = menu.querySelector('a[href="lca_pedidos_pendentes.asp"]');
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
    var form = document.getElementById('dispatch-form');
    var groups = Array.from(form.querySelectorAll('[data-dispatch-group]'));
    var orders = Array.from(form.querySelectorAll('[data-dispatch-order]'));
    var submit = document.getElementById('dispatch-submit');
    var dialog = document.getElementById('dispatch-confirm');
    var selected = 0;
    document.getElementById('dispatch-total').textContent = orders.length;
    document.getElementById('dispatch-vendors').textContent = groups.length;
    document.getElementById('dispatch-bar').hidden = orders.length === 0;

    function updateSelection() {
        selected = orders.filter(function (order) { return order.checked; }).length;
        document.getElementById('dispatch-selected').textContent = selected;
        submit.disabled = selected === 0;
        groups.forEach(function (group) {
            var toggle = group.querySelector('[data-dispatch-vendor]');
            var total = group.querySelectorAll('[data-dispatch-order]').length;
            var checked = group.querySelectorAll('[data-dispatch-order]:checked').length;
            toggle.checked = checked === total;
            toggle.indeterminate = checked > 0 && checked < total;
        });
    }

    form.addEventListener('change', function (event) {
        if (event.target.matches('[data-dispatch-vendor]')) {
            var group = event.target.closest('[data-dispatch-group]');
            group.querySelectorAll('[data-dispatch-order]').forEach(function (order) {
                order.checked = event.target.checked;
            });
        }
        updateSelection();
    });
    form.addEventListener('click', function (event) {
        var button = event.target.closest('[data-dispatch-select]');
        if (!button) return;
        orders.forEach(function (order) { order.checked = button.dataset.dispatchSelect === 'all'; });
        updateSelection();
    });
    form.addEventListener('submit', function (event) {
        event.preventDefault();
        updateSelection();
        if (!selected) return;
        dialog.returnValue = '';
        document.getElementById('dispatch-confirm-count').textContent =
            selected === 1 ? '1 pedido selecionado.' : selected + ' pedidos selecionados.';
        dialog.showModal();
    });
    dialog.addEventListener('close', function () {
        if (dialog.returnValue !== 'confirmar') return;
        submit.disabled = true;
        HTMLFormElement.prototype.submit.call(form);
    });
    updateSelection();
})();
</script>
</body>
</html>
