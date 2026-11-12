<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Pend&ecirc;ncias de pagamento | Portal Stillos</title>
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

.pending-overview { margin-bottom:18px; }
.pending-overview h2 { margin-bottom:10px; }
.pending-chart { display:grid; grid-template-columns:auto minmax(0,1fr); align-items:center; gap:16px; margin-bottom:16px; }
.pending-chart .cap-chart-donut { width:96px; margin:0; }
.pending-chart .cap-chart-donut strong { width:64px; }
.pending-motive { min-width:200px; max-width:420px; overflow-wrap:anywhere; white-space:pre-line; }
.pending-type,.pending-value { white-space:nowrap; }
.pending-review { background:#f1edf6; color:#493d60; }
@media(max-width:470px) { .pending-chart { grid-template-columns:1fr; } }
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
      <div class="topbar-title">Portal Stillos <span>/ Pend&ecirc;ncias de pagamento</span></div>
      <div class="topbar-actions"><a class="user-name" href="adm_log_usuarios.asp"><%=Server.HTMLEncode(Session("NOME") & "")%></a><a href="manual_sistema.asp" target="_blank" rel="noopener">Manual</a><a class="btn btn-outline-secondary btn-sm" href="adm_sair.asp">Sair</a></div>
    </header>
    <main class="dashboard-main ppv-main" id="conteudo">


      <header class="dashboard-heading">
        <div><p class="eyebrow">FECHAMENTOS</p><h1>Pend&ecirc;ncias de pagamento</h1><p>Acompanhe os acr&eacute;scimos e decr&eacute;scimos vinculados aos pagamentos dos vendedores.</p></div>
        <div class="period-actions">
          <a class="btn btn-outline-secondary" href="cap_fechamentos.asp">Voltar a Fechamentos</a>
          <a class="btn btn-primary" href="cap_fechamentos_pendencias_formulario.asp">Cadastrar pend&ecirc;ncia</a>
        </div>
      </header>
      <div class="search-panel"><strong>Busca no portal</strong>
        <!-- #include file="includes/dashboard_busca.asp" -->
      </div>
      <div class="session-notices">
        <!-- #include file="includes/dashboard_avisos.asp" -->
      </div>
<%
Dim pendingMessage, pendingAlert
pendingAlert = "success"
If Request.QueryString("sucesso") = "" Then
    Select Case Request.QueryString("codmsg")
        Case "1": pendingMessage = "Pend&ecirc;ncia cadastrada."
        Case "2": pendingMessage = "Pend&ecirc;ncia alterada."
        Case "3": pendingMessage = "Pend&ecirc;ncia exclu&iacute;da."
    End Select
ElseIf Request.QueryString("codmsg") = "41" Then
    pendingAlert = "danger"
    pendingMessage = "O pagamento de refer&ecirc;ncia j&aacute; foi conclu&iacute;do. A pend&ecirc;ncia n&atilde;o pode ser alterada ou exclu&iacute;da."
End If
If Len(pendingMessage & "") > 0 Then%>
      <p class="alert alert-<%=pendingAlert%>" role="status"><%=pendingMessage%></p>
<%End If

Dim pendingSql, pendingFilter, pendingPage, pendingPattern, pendingIndex
Dim pendingTotal, pendingOpen, pendingProcessed, pendingReview, pendingShare
Dim pendingState, pendingPayment, pendingIsCredit, pendingId
pendingFilter = Request.QueryString("situacao")
pendingPage = 1
Set pendingPattern = New RegExp
pendingPattern.Pattern = "^[1-9][0-9]{0,8}$"
If pendingPattern.Test(Request.QueryString("pagina")) Then pendingPage = CLng(Request.QueryString("pagina"))
Set pendingPattern = Nothing

AbrirConexao()
pendingSql = "SELECT d.*, v.nome_vendedor, u.nome AS nome_cadastro, p.data_pagamento AS pagamento_periodo, " & _
    "CASE WHEN p.cod_fechamento_periodo IS NOT NULL THEN CASE WHEN p.data_processado_pagamento IS NULL THEN 0 ELSE 1 END " & _
    "WHEN d.vendedor_recebe_em IS NOT NULL AND corte.primeiro_pendente IS NOT NULL THEN " & _
    "CASE WHEN DATE(d.vendedor_recebe_em) <= DATE(corte.primeiro_pendente) THEN 1 ELSE 0 END ELSE 2 END AS situacao_pendencia " & _
    "FROM TB_FECHAMENTO_PENDENCIAS_PAGAMENTO d " & _
    "INNER JOIN TB_VENDEDORES v ON v.cod_vendedor = d.cod_vendedor " & _
    "INNER JOIN TB_ADM_USUARIOS u ON u.cod_usuario = d.cadastrado_por " & _
    "LEFT JOIN TB_FECHAMENTO_PAGAMENTO_PERIODOS p ON p.cod_fechamento_periodo = d.cod_fechamento_periodo " & _
    "CROSS JOIN (SELECT MIN(data_pagamento) AS primeiro_pendente FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE data_processado_pagamento IS NULL) corte " & _
    "ORDER BY d.data_cadastro DESC, d.cod_pendencia DESC"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3
rs.PageSize = 20
rs.Open pendingSql, objConexao
pendingTotal = rs.RecordCount
pendingOpen = 0
pendingProcessed = 0
pendingReview = 0
While Not rs.Eof
    Select Case CLng(rs("situacao_pendencia"))
        Case 0: pendingOpen = pendingOpen + 1
        Case 1: pendingProcessed = pendingProcessed + 1
        Case 2: pendingReview = pendingReview + 1
    End Select
    rs.MoveNext
Wend
pendingShare = 0
If pendingTotal > 0 Then pendingShare = 100 * pendingProcessed / pendingTotal
If pendingTotal > 0 Then rs.MoveFirst
Select Case pendingFilter
    Case "abertas": rs.Filter = "situacao_pendencia = 0"
    Case "processadas": rs.Filter = "situacao_pendencia = 1"
    Case "revisao": rs.Filter = "situacao_pendencia = 2"
    Case Else: pendingFilter = ""
End Select
If rs.RecordCount > 0 Then
    If pendingPage > rs.PageCount Then pendingPage = rs.PageCount
    rs.AbsolutePage = pendingPage
End If
%>
      <section class="pending-overview" aria-labelledby="pending-overview">
        <h2 id="pending-overview">Vis&atilde;o geral</h2>
        <dl class="cap-metrics">
          <div class="cap-metric"><dt>Registros no hist&oacute;rico</dt><dd><%=FormatNumber(pendingTotal,0)%></dd></div>
          <div class="cap-metric"><dt>Em aberto</dt><dd><%=FormatNumber(pendingOpen,0)%></dd></div>
          <div class="cap-metric"><dt>Processadas</dt><dd><%=FormatNumber(pendingProcessed,0)%></dd></div>
          <div class="cap-metric"><dt>A conferir</dt><dd><%=FormatNumber(pendingReview,0)%></dd></div>
        </dl>
        <%If pendingTotal > 0 Then%>
        <div class="cap-chart pending-chart">
          <div class="cap-chart-donut" style="--cap-share:<%=CInt(pendingShare)%>%" aria-hidden="true"><strong><%=FormatNumber(pendingShare,1)%>%</strong></div>
          <div><h3>Processamento no hist&oacute;rico</h3><p><strong><%=FormatNumber(pendingShare,1)%>% processadas</strong> &middot; <%=FormatNumber(pendingProcessed,0)%> de <%=FormatNumber(pendingTotal,0)%> registros.</p></div>
        </div>
        <%End If%>
      </section>
      <section class="dashboard-panel period-panel" aria-labelledby="pending-list">
        <header>
          <div><h2 id="pending-list">Pend&ecirc;ncias cadastradas</h2><p class="panel-description"><%=FormatNumber(rs.RecordCount,0)%> no filtro atual</p></div>
          <form method="get" class="period-filter">
            <label for="pending-filter">Situa&ccedil;&atilde;o</label>
            <select class="form-select w-auto" name="situacao" id="pending-filter">
              <option value="">Todas</option>
              <option value="abertas"<%If pendingFilter = "abertas" Then%> selected<%End If%>>Em aberto</option>
              <option value="processadas"<%If pendingFilter = "processadas" Then%> selected<%End If%>>Processadas</option>
              <option value="revisao"<%If pendingFilter = "revisao" Then%> selected<%End If%>>A conferir</option>
            </select>
            <button class="btn btn-outline-secondary" type="submit">Filtrar</button>
          </form>
        </header>
        <%If rs.Eof Then%>
        <p class="empty-state">Nenhuma pend&ecirc;ncia encontrada.</p>
        <%Else%>
        <div class="period-scroll" role="region" aria-labelledby="pending-list" tabindex="0">
          <table class="period-table">
            <thead><tr><th scope="col">Pagamento</th><th scope="col">Vendedor</th><th scope="col">Ajuste</th><th scope="col">Motivo</th><th scope="col">Cadastro</th><th scope="col">Situa&ccedil;&atilde;o</th><th scope="col">A&ccedil;&otilde;es</th></tr></thead>
            <tbody>
            <%For pendingIndex = 1 To rs.PageSize
                If rs.Eof Then Exit For
                pendingState = CLng(rs("situacao_pendencia"))
                pendingPayment = rs("pagamento_periodo")
                If Not IsDate(pendingPayment) Then pendingPayment = rs("vendedor_recebe_em")
                pendingIsCredit = (UCase(rs("tipo") & "") = "A")
                pendingId = rs("cod_pendencia")
            %>
            <tr>
              <td class="period-date"><%If IsDate(pendingPayment) Then%><%=ArrumaData(pendingPayment)%><%Else%>&mdash;<%End If%></td>
              <td><strong><%=Server.HTMLEncode(rs("nome_vendedor") & "")%></strong></td>
              <td><span class="pending-type"><%If pendingIsCredit Then%>Acr&eacute;scimo (+)<%Else%>Decr&eacute;scimo (-)<%End If%></span><strong class="period-meta pending-value"><%If IsNumeric(rs("valor")) Then%><%=FormatCurrency(rs("valor"))%><%Else%>&mdash;<%End If%></strong></td>
              <td class="pending-motive"><%=Server.HTMLEncode(rs("motivo") & "")%></td>
              <td><%=Server.HTMLEncode(rs("nome_cadastro") & "")%><span class="period-meta period-date"><%=ArrumaData(rs("data_cadastro"))%></span></td>
              <td><%Select Case pendingState
                  Case 0%><span class="period-badge">Em aberto</span><%
                  Case 1%><span class="period-badge period-done">Processada</span><%
                  Case 2%><span class="period-badge pending-review" title="Sem informa&ccedil;&atilde;o suficiente para confirmar o processamento">A conferir</span><%
                  End Select%></td>
              <td><%If pendingState = 0 Then%><div class="period-actions">
                <a class="btn btn-outline-secondary btn-sm" href="cap_fechamentos_pendencias_formulario.asp?comando=alterar&amp;cod_pendencia=<%=pendingId%>">Editar</a>
                <a class="btn btn-outline-danger btn-sm" href="cap_fechamentos_pendencias_formulario.asp?comando=excluir&amp;cod_pendencia=<%=pendingId%>">Excluir</a>
              </div><%Else%>&mdash;<%End If%></td>
            </tr>
            <%rs.MoveNext
            Next%>
            </tbody>
          </table>
        </div>
        <%If rs.PageCount > 1 Then%>
        <nav class="period-paging" aria-label="P&aacute;ginas das pend&ecirc;ncias">
          <%If pendingPage > 1 Then%><a class="btn btn-outline-secondary btn-sm" href="?situacao=<%=pendingFilter%>&amp;pagina=<%=pendingPage-1%>">Anterior</a><%End If%>
          <span>P&aacute;gina <%=pendingPage%> de <%=rs.PageCount%></span>
          <%If pendingPage < rs.PageCount Then%><a class="btn btn-outline-secondary btn-sm" href="?situacao=<%=pendingFilter%>&amp;pagina=<%=pendingPage+1%>">Pr&oacute;xima</a><%End If%>
        </nav>
        <%End If%>
        <%End If
        rs.Close
        Set rs = Nothing
        FecharConexao()%>
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
