<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Comissionamento de coordenadores | Portal Stillos</title>
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

.coordinator-overview { margin-bottom:18px; }
.coordinator-overview h2 { margin-bottom:10px; }
.coordinator-percent { font-size:1.12em; font-weight:700; }
.coordinator-incomplete { display:block; margin-top:4px; color:#754600; font-size:.85em; }
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
      <div class="topbar-title">Portal Stillos <span>/ Comissionamento de coordenadores</span></div>
      <div class="topbar-actions"><a class="user-name" href="adm_log_usuarios.asp"><%=Server.HTMLEncode(Session("NOME") & "")%></a><a href="manual_sistema.asp" target="_blank" rel="noopener">Manual</a><a class="btn btn-outline-secondary btn-sm" href="adm_sair.asp">Sair</a></div>
    </header>
    <main class="dashboard-main ppv-main" id="conteudo">


      <header class="dashboard-heading">
        <div><p class="eyebrow">FECHAMENTOS</p><h1>Comissionamento de coordenadores</h1><p>Consulte os percentuais por faixa de ano de conclus&atilde;o.</p></div>
        <div class="period-actions">
          <a class="btn btn-outline-secondary" href="cap_fechamentos.asp">Voltar a Fechamentos</a>
          <a class="btn btn-primary" href="cap_fechamentos_comissionamento_coordenadores_formulario.asp">Nova regra de comiss&atilde;o</a>
        </div>
      </header>
      <div class="search-panel"><strong>Busca no portal</strong>
        <!-- #include file="includes/dashboard_busca.asp" -->
      </div>
      <div class="session-notices">
        <!-- #include file="includes/dashboard_avisos.asp" -->
      </div>
      <%
      Select Case Request.QueryString("codmsg")
          Case "1"
      %><p class="alert alert-success" role="status">Regra de comiss&atilde;o cadastrada.</p><%
          Case "3"
      %><p class="alert alert-success" role="status">Regra de comiss&atilde;o exclu&iacute;da.</p><%
      End Select
      If Len(Request.QueryString("msg")) > 0 Then
      %><p class="alert alert-info" role="status"><%=Server.HTMLEncode(Request.QueryString("msg"))%></p><%
      End If

      Dim coordinatorRows, coordinatorCount, coordinatorConfigured, coordinatorIndex
      Dim coordinatorStart, coordinatorEnd, coordinatorRate, coordinatorId
      AbrirConexao()
      Set rs = objConexao.Execute("SELECT ano_inicio, ano_fim, percentual, COD_COMISSIONAMENTO_COORDENADORES FROM stillos3.TB_FECHAMENTO_COMISSIONAMENTO_COORDENADORES ORDER BY ano_inicio DESC, COD_COMISSIONAMENTO_COORDENADORES DESC")
      coordinatorCount = 0
      coordinatorConfigured = 0
      If Not rs.Eof Then
          coordinatorRows = rs.GetRows()
          coordinatorCount = UBound(coordinatorRows, 2) + 1
          For coordinatorIndex = 0 To coordinatorCount - 1
              If HasCoordinatorRange(coordinatorRows(0, coordinatorIndex), coordinatorRows(1, coordinatorIndex)) Then
                  coordinatorConfigured = coordinatorConfigured + 1
              End If
          Next
      End If
      rs.Close
      Set rs = Nothing
      FecharConexao()

      Function HasCoordinatorRange(startYear, endYear)
          HasCoordinatorRange = False
          If Not IsNumeric(startYear) Or Not IsNumeric(endYear) Then Exit Function
          HasCoordinatorRange = (CDbl(startYear) > 0 And CDbl(endYear) >= CDbl(startYear))
      End Function
      %>
      <section class="coordinator-overview" aria-labelledby="coordinator-overview-title">
        <h2 id="coordinator-overview-title">Vis&atilde;o geral</h2>
        <dl class="cap-metrics">
          <div class="cap-metric"><dt>Regras cadastradas</dt><dd><%=coordinatorCount%></dd></div>
          <div class="cap-metric"><dt>Com faixa de anos v&aacute;lida</dt><dd><%=coordinatorConfigured%></dd></div>
          <div class="cap-metric"><dt>Com faixa incompleta ou inv&aacute;lida</dt><dd><%=coordinatorCount - coordinatorConfigured%></dd></div>
        </dl>
      </section>
      <section class="dashboard-panel period-panel" aria-labelledby="coordinator-list">
        <header><h2 id="coordinator-list">Regras de comiss&atilde;o</h2></header>
        <%If coordinatorCount = 0 Then%>
          <p class="empty-state">Nenhuma regra de comiss&atilde;o cadastrada para coordenadores.</p>
        <%Else%>
        <div class="period-scroll" role="region" aria-labelledby="coordinator-list" tabindex="0">
          <table class="period-table" style="min-width:540px">
            <thead><tr><th scope="col">Ano inicial</th><th scope="col">Ano final</th><th scope="col">Comiss&atilde;o sobre as vendas</th><th scope="col">A&ccedil;&otilde;es</th></tr></thead>
            <tbody>
            <%For coordinatorIndex = 0 To coordinatorCount - 1
                coordinatorStart = coordinatorRows(0, coordinatorIndex)
                coordinatorEnd = coordinatorRows(1, coordinatorIndex)
                coordinatorRate = coordinatorRows(2, coordinatorIndex)
                coordinatorId = coordinatorRows(3, coordinatorIndex)
            %>
              <tr>
                <td><%If Len(Trim(coordinatorStart & "")) > 0 Then%><%=Server.HTMLEncode(coordinatorStart & "")%><%Else%>&mdash;<%End If%>
                  <%If Not HasCoordinatorRange(coordinatorStart, coordinatorEnd) Then%><span class="coordinator-incomplete">Verifique a faixa de anos</span><%End If%>
                </td>
                <td><%If Len(Trim(coordinatorEnd & "")) > 0 Then%><%=Server.HTMLEncode(coordinatorEnd & "")%><%Else%>&mdash;<%End If%></td>
                <td class="coordinator-percent"><%If IsNumeric(coordinatorRate) Then%><%=FormatNumber(coordinatorRate, 2)%>%<%Else%>&mdash;<%End If%></td>
                <td><a class="btn btn-outline-danger btn-sm" href="cap_fechamentos_comissionamento_coordenadores_formulario.asp?comando=excluir&amp;COD_COMISSIONAMENTO_COORDENADORES=<%=coordinatorId%>">Excluir</a></td>
              </tr>
            <%Next%>
            </tbody>
          </table>
        </div>
        <%End If%>
      </section>
    </main>
    <footer class="dashboard-footer">&copy; <%=Year(Date)%> Grupo Stillos <span>Sistema administrativo interno</span></footer>
  </div>
</div>
<script>
(function () {
    var toggle = document.getElementById('ppvMenuToggle');
    var menu = document.getElementById('menuPrincipal');
    var previous = menu.querySelector('[aria-current="page"]');
    var current = menu.querySelector('a[href="cap_fechamentos.asp"]');
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

</script>
</body>
</html>
