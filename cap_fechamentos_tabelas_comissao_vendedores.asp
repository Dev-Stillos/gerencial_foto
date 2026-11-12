<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Tabelas de comissionamento | Portal Stillos</title>
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

.commission-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(min(100%,365px),1fr)); gap:14px; align-items:start; }
.commission-card { height:auto; min-width:0; }
.commission-card h3 { margin:0 0 4px; font-size:1.1em; overflow-wrap:anywhere; }
.commission-meta { color:#53566b; font-size:.9em; margin:0 0 12px; }
.commission-card dl { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:8px; margin:0 0 12px; }
.commission-card dl > div { border:1px solid #e4dfeb; border-radius:8px; padding:8px; }
.commission-card dt { font-size:.85em; font-weight:500; }
.commission-card dd { font-size:1.15em; font-weight:700; margin:4px 0 0; }
.commission-card .period-actions { margin-top:12px; }
.commission-chart { margin-bottom:16px; }
@media(max-width:460px) { .commission-card dl { grid-template-columns:1fr; } }

.commission-preview { width:min(1240px,calc(100% - 24px)); max-height:calc(100dvh - 24px); padding:0; border:1px solid #e4dfeb; border-radius:12px; color:#242738; }
.commission-preview[open] { display:flex; flex-direction:column; }
.commission-preview::backdrop { background:#24273899; }
.commission-preview-header { display:flex; flex-wrap:wrap; align-items:center; justify-content:space-between; gap:10px; padding:12px 16px; border-bottom:1px solid #e4dfeb; }
.commission-preview-frame { width:100%; height:80dvh; min-height:0; border:0; }
.commission-preview-status { margin:0; padding:12px 16px; }
.commission-preview [hidden] { display:none !important; }
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
      <div class="topbar-title">Portal Stillos <span>/ Tabelas de comissionamento</span></div>
      <div class="topbar-actions"><a class="user-name" href="adm_log_usuarios.asp"><%=Server.HTMLEncode(Session("NOME") & "")%></a><a href="manual_sistema.asp" target="_blank" rel="noopener">Manual</a><a class="btn btn-outline-secondary btn-sm" href="adm_sair.asp">Sair</a></div>
    </header>
    <main class="dashboard-main ppv-main" id="conteudo">


      <header class="dashboard-heading">
        <div><p class="eyebrow">FECHAMENTOS</p><h1>Tabelas de comissionamento</h1><p>Consulte os produtos e vendedores vinculados. Use Editar comiss&otilde;es para ajustar percentuais ou valores fixos por produto.</p></div>
        <div class="period-actions">
          <a class="btn btn-outline-secondary" href="cap_fechamentos.asp">Voltar a Fechamentos</a>
          <a class="btn btn-primary" href="cap_fechamentos_tabelas_comissao_vendedores_formulario.asp">Cadastrar tabela</a>
        </div>
      </header>
      <div class="search-panel"><strong>Busca no portal</strong>
        <!-- #include file="includes/dashboard_busca.asp" -->
      </div>
      <div class="session-notices">
        <!-- #include file="includes/dashboard_avisos.asp" -->
      </div>
      <%Dim sucesso, codmsg
	codmsg	= CStr(""&Request.QueryString("codmsg"))
	sucesso = CStr(""&Request.QueryString("sucesso"))
		
	If sucesso = "" and codmsg <> "" Then%>
    	<div class="alert alert-success" role="status">
        	<%
			If codmsg = "1" Then
				Response.Write("Tabela de Comiss&atilde;o de Vendedores inclu&iacute;da com sucesso!")
			Elseif codmsg = "2" Then
				Response.Write("Tabela de Comiss&atilde;o de Vendedores alterada com sucesso!")
			Elseif codmsg = "3" Then	
				Response.Write("Tabela de Comiss&atilde;o de Vendedores exclu&iacute;da com sucesso!")
			Elseif codmsg = "101" Then	
				Response.Write("Alterado a comiss&atilde;o de Produtos com sucesso!")
			End If	
			%>
            
        
        </div>
    <%ElseIf sucesso <> "" and codmsg <> "" Then%>
        <div class="alert alert-danger" role="alert">
			<%
			If codmsg = "5" Then
				Response.Write("Tabela de Comiss&atilde;o de Vendedores n&atilde;o exclu&iacute;da, pois h&aacute; vendedores utilizando esta tabela!")
			ElseIf codmsg = "4" Then
				Response.Write("Tabela de Comiss&atilde;o de Vendedores n&atilde;o inclu&iacute;da, pois j&aacute; existe outra tabela com o mesmo nome!")
			
			'/////////
			'ElseIf codmsg = "401" Then
			'ElseIf codmsg = "402" Then
			'ElseIf codmsg = "403" Then
			'ElseIf codmsg = "404" Then
			End If	
			%>
</div>
	<%End If%>
<%
Dim commissionPage, commissionInput, commissionPattern, commissionSql
Dim commissionActive, commissionPending, commissionCoverage, commissionRow, commissionId
commissionPage = 1
commissionInput = Request.QueryString("pagina")
Set commissionPattern = New RegExp
commissionPattern.Pattern = "^[1-9][0-9]{0,8}$"
If commissionPattern.Test(commissionInput) Then commissionPage = CLng(commissionInput)
Set commissionPattern = Nothing

AbrirConexao()
Set rs2 = objConexao.Execute("SELECT COUNT(*) AS qtd FROM TB_PRODUTOS WHERE ativo = '1' AND PAGAMENTO_COMISSAO = 'S'")
commissionActive = CLng(rs2("qtd"))
rs2.Close
Set rs2 = Nothing

commissionSql = "SELECT t.*, u.nome AS nome_cadastro, " & _
    "(SELECT COUNT(*) FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES_PRODUTOS cp WHERE cp.cod_tabela_comissao_vendedores = t.cod_tabela_comissao_vendedores) AS qtd_produtos, " & _
    "(SELECT COUNT(*) FROM TB_VENDEDORES v WHERE v.cod_tabela_comissao_vendedores = t.cod_tabela_comissao_vendedores) AS qtd_vendedores, " & _
    "(SELECT COUNT(*) FROM TB_PRODUTOS p WHERE p.ativo = '1' AND p.PAGAMENTO_COMISSAO = 'S' " & _
    "AND NOT EXISTS (SELECT 1 FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES_PRODUTOS cp WHERE cp.cod_tabela_comissao_vendedores = t.cod_tabela_comissao_vendedores AND cp.cod_produto = p.cod_produto)) AS qtd_pendentes " & _
    "FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES t LEFT JOIN TB_ADM_USUARIOS u ON u.cod_usuario = t.cod_usuario " & _
    "ORDER BY t.descricao DESC, t.cod_tabela_comissao_vendedores DESC"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3
rs.PageSize = 50
rs.Open commissionSql, objConexao
If Not rs.Eof Then
    If commissionPage > rs.PageCount Then commissionPage = rs.PageCount
    rs.AbsolutePage = commissionPage
End If
%>
      <section class="period-overview" aria-labelledby="commission-overview">
        <h2 id="commission-overview">Vis&atilde;o geral</h2>
        <dl class="cap-metrics">
          <div class="cap-metric"><dt>Tabelas cadastradas</dt><dd><%=FormatNumber(rs.RecordCount,0)%></dd></div>
          <div class="cap-metric"><dt>Produtos ativos com comiss&atilde;o</dt><dd><%=FormatNumber(commissionActive,0)%></dd></div>
        </dl>
      </section>
      <section aria-labelledby="commission-list">
        <h2 class="mb-3" id="commission-list">Tabelas cadastradas</h2>
        <%If rs.Eof Then%>
        <p class="empty-state">Nenhuma tabela de comissionamento cadastrada.</p>
        <%Else%>
        <div class="commission-grid">
          <%For commissionRow = 1 To rs.PageSize
              If rs.Eof Then Exit For
              commissionId = rs("cod_tabela_comissao_vendedores")
              commissionPending = CLng(rs("qtd_pendentes"))
              commissionCoverage = 0
              If commissionActive > 0 Then commissionCoverage = 100 * (commissionActive - commissionPending) / commissionActive
          %>
          <article class="dashboard-panel commission-card" aria-labelledby="commission-<%=commissionId%>">
            <h3 id="commission-<%=commissionId%>"><%=Server.HTMLEncode(rs("descricao") & "")%></h3>
            <p class="commission-meta">Tabela <%=commissionId%> &middot; Cadastro: <%=Server.HTMLEncode(rs("nome_cadastro") & "")%><br>&Uacute;ltima intera&ccedil;&atilde;o: <%=ArrumaData(rs("modificado"))%></p>
            <dl>
              <div><dt>Produtos cadastrados</dt><dd><%=FormatNumber(rs("qtd_produtos"),0)%></dd></div>
              <div><dt>Produtos ativos pendentes</dt><dd><span class="period-badge<%If commissionPending = 0 Then%> period-done<%End If%>"><%=FormatNumber(commissionPending,0)%></span></dd></div>
              <div><dt>Vendedores vinculados</dt><dd><%=FormatNumber(rs("qtd_vendedores"),0)%></dd></div>
            </dl>
            <%If commissionActive > 0 Then%>
            <div class="commission-chart">
              <strong>Cobertura de produtos ativos: <%=FormatNumber(commissionCoverage,1)%>%</strong>
              <div class="cap-chart-track" aria-hidden="true"><span style="width:<%=CInt(commissionCoverage)%>%"></span></div>
            </div>
            <%End If%>
            <div class="period-actions">
              <a class="btn btn-primary btn-sm" href="cap_fechamentos_tabelas_comissao_vendedores_formulario_consulta_produtos_formulario.asp?cod_tabela_comissao_vendedores=<%=commissionId%>">Editar comiss&otilde;es</a>
              <a class="btn btn-outline-secondary btn-sm" data-commission-print href="cap_fechamentos_tabelas_comissao_vendedores_formulario_consulta_produtos_formulario_imprimir.asp?cod_tabela_comissao_vendedores=<%=commissionId%>">Imprimir tabela</a>
              <a class="btn btn-outline-secondary btn-sm" href="cap_fechamentos_tabelas_comissao_vendedores_formulario.asp?cod_tabela_comissao_vendedores=<%=commissionId%>&amp;comando=alterar">Editar</a>
              <a class="btn btn-outline-danger btn-sm" href="cap_fechamentos_tabelas_comissao_vendedores_formulario.asp?cod_tabela_comissao_vendedores=<%=commissionId%>&amp;comando=excluir">Excluir</a>
            </div>
          </article>
          <%rs.MoveNext
          Next%>
        </div>
        <%If rs.PageCount > 1 Then%>
        <nav class="period-paging" aria-label="P&aacute;ginas das tabelas">
          <%If commissionPage > 1 Then%><a class="btn btn-outline-secondary btn-sm" href="?pagina=<%=commissionPage-1%>">Anterior</a><%End If%>
          <span>P&aacute;gina <%=commissionPage%> de <%=rs.PageCount%></span>
          <%If commissionPage < rs.PageCount Then%><a class="btn btn-outline-secondary btn-sm" href="?pagina=<%=commissionPage+1%>">Pr&oacute;xima</a><%End If%>
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
<dialog class="commission-preview" id="commission-preview" aria-labelledby="commission-preview-title">
  <header class="commission-preview-header">
    <h2 id="commission-preview-title">Pr&eacute;via da tabela de comiss&otilde;es</h2>
    <div class="period-actions">
      <button class="btn btn-primary btn-sm" type="button" id="commission-print-button" disabled>Imprimir / Salvar PDF</button>
      <form method="dialog"><button class="btn btn-outline-secondary btn-sm" autofocus>Fechar</button></form>
    </div>
  </header>
  <p class="commission-preview-status" id="commission-preview-status" role="status" hidden></p>
  <iframe class="commission-preview-frame" id="commission-preview-frame" title="Pr&eacute;via da tabela de comiss&otilde;es" hidden></iframe>
</dialog>
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


(function () {
    var dialog = document.getElementById('commission-preview');
    var frame = document.getElementById('commission-preview-frame');
    var status = document.getElementById('commission-preview-status');
    var printButton = document.getElementById('commission-print-button');
    var trigger;

    document.addEventListener('click', function (event) {
        var link = event.target.closest('[data-commission-print]');
        if (!link || typeof dialog.showModal !== 'function' || event.ctrlKey || event.metaKey || event.shiftKey || event.altKey) return;
        event.preventDefault();
        trigger = link;
        var url = new URL(link.href);
        url.searchParams.set('modal', '1');
        status.textContent = 'Carregando tabela...';
        status.hidden = false;
        frame.hidden = false;
        printButton.disabled = true;
        frame.src = url.href;
        dialog.showModal();
    });
    frame.addEventListener('load', function () {
        if (!dialog.open || !frame.hasAttribute('src')) return;
        var preview = frame.contentDocument;
        if (preview && preview.URL === 'about:blank') return;
        var ready = preview && preview.querySelector('[data-commission-report]');
        printButton.disabled = !ready;
        status.hidden = Boolean(ready);
        if (!ready) status.textContent = 'A pr\u00e9via da tabela n\u00e3o est\u00e1 dispon\u00edvel.';
    });
    printButton.addEventListener('click', function () {
        if (!printButton.disabled) frame.contentWindow.print();
    });
    dialog.addEventListener('close', function () {
        frame.hidden = true;
        frame.removeAttribute('src');
        printButton.disabled = true;
        if (trigger) trigger.focus();
    });
})();
</script>
</body>
</html>
