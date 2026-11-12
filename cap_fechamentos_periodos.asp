<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Per&iacute;odos de vendas | Portal Stillos</title>
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
      <div class="topbar-title">Portal Stillos <span>/ Per&iacute;odos de vendas</span></div>
      <div class="topbar-actions"><a class="user-name" href="adm_log_usuarios.asp"><%=Server.HTMLEncode(Session("NOME") & "")%></a><a href="manual_sistema.asp" target="_blank" rel="noopener">Manual</a><a class="btn btn-outline-secondary btn-sm" href="adm_sair.asp">Sair</a></div>
    </header>
    <main class="dashboard-main ppv-main" id="conteudo">

      <header class="dashboard-heading">
        <div><p class="eyebrow">FECHAMENTOS</p><h1>Per&iacute;odos de vendas</h1><p>Acompanhe as datas de lan&ccedil;amento, pagamento e processamento.</p></div>
        <div class="period-actions">
          <a class="btn btn-outline-secondary" href="cap_fechamentos.asp">Voltar a Fechamentos</a>
          <a class="btn btn-primary" href="cap_fechamentos_periodos_formulario.asp">Cadastrar per&iacute;odo</a>
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
				Response.Write("Per&iacute;odo de Venda inclu&iacute;do com sucesso!")
			Elseif codmsg = "2" Then
				Response.Write("Per&iacute;odo de Venda alterado com sucesso!")
			Elseif codmsg = "3" Then	
				Response.Write("Per&iacute;odo de Venda exclu&iacute;do com sucesso!")
			End If	
			%>
            
        
        </div>
    <%ElseIf sucesso <> "" and codmsg <> "" Then%>
        <div class="alert alert-danger" role="alert">
			<%
			If codmsg = "3" Then
				Response.Write("Per&iacute;odo de Vendas n&atilde;o exclu&iacute;do, pois h&aacute; pedidos com este per&iacute;odo!")
			ElseIf codmsg = "4" Then
				Response.Write("Per&iacute;odo de Vendas n&atilde;o inclu&iacute;do, pois &eacute; permitido somente 1 pagamento por dia, devido os recibos de venda serem gerados com a data do pagamento!")
			
			'/////////
			ElseIf codmsg = "401" Then
				Response.Write("Data in&iacute;cio maior que Data T&eacute;rmino do Per&iacute;odo de Venda N&Atilde;O inclu&iacute;do!")
			ElseIf codmsg = "402" Then
				Response.Write("Data t&eacute;rmino de venda maior que data limite de lan&ccedil;amento de Venda N&Atilde;O inclu&iacute;do!")
			ElseIf codmsg = "403" Then
				Response.Write("Data do pagamento n&atilde;o pode ser inferior a data m&aacute;xima de lan&ccedil;amento de vendas N&Atilde;O inclu&iacute;do!")
			ElseIf codmsg = "404" Then
				Response.Write("Voc&ecirc; est&aacute; tentando inserir um per&iacute;odo, por&eacute;m j&aacute; existe um outro Per&iacute;odo de Venda dentro do Per&iacute;odo informado, n&atilde;o &eacute; permitido misturar as datas dos per&iacute;odos de venda. Por tanto per&iacute;odo N&Atilde;O inclu&iacutedo!")
			ElseIf codmsg = "405" Then
				Response.Write("Voc&ecirc; est&aacute; tentando inserir um per&iacute;odo com meses diferentes n&atilde;o &eacute; permitido data in&iacute;cio e data t&eacute;rmino terem meses diferentes. Por tanto per&iacute;odo N&Atilde;O inclu&iacute;do!")
			End If	
			%>
</div>
	<%End If%>
<%
Dim periodStatus, periodWhere, periodPage, periodInput, periodPattern
Dim periodTotal, periodProcessed, periodPending, periodNext, periodShare, periodRow
periodStatus = Request.QueryString("situacao")
Select Case periodStatus
    Case "pendentes"
        periodWhere = " WHERE p.data_processado_pagamento IS NULL"
    Case "processados"
        periodWhere = " WHERE p.data_processado_pagamento IS NOT NULL"
    Case Else
        periodStatus = ""
        periodWhere = ""
End Select
periodPage = 1
periodInput = Request.QueryString("pagina")
Set periodPattern = New RegExp
periodPattern.Pattern = "^[1-9][0-9]{0,8}$"
If periodPattern.Test(periodInput) Then periodPage = CLng(periodInput)
Set periodPattern = Nothing

AbrirConexao()
Set rs2 = objConexao.Execute("SELECT COUNT(*) AS total, COUNT(data_processado_pagamento) AS processados, MIN(CASE WHEN data_processado_pagamento IS NULL THEN data_pagamento END) AS proximo FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS")
periodTotal = CLng(rs2("total"))
periodProcessed = CLng(rs2("processados"))
periodPending = periodTotal - periodProcessed
periodNext = rs2("proximo")
periodShare = 0
If periodTotal > 0 Then periodShare = 100 * periodProcessed / periodTotal
rs2.Close
Set rs2 = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3
rs.PageSize = 50
rs.Open "SELECT p.*, u.nome AS nome_cadastro FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS p LEFT JOIN TB_ADM_USUARIOS u ON u.cod_usuario = p.cod_usuario" & periodWhere & " ORDER BY p.data_termino DESC, p.cod_fechamento_periodo DESC", objConexao
If Not rs.Eof Then
    If periodPage > rs.PageCount Then periodPage = rs.PageCount
    rs.AbsolutePage = periodPage
End If
%>
      <section class="period-overview" aria-labelledby="period-overview-title">
        <h2 id="period-overview-title">Vis&atilde;o geral</h2>
        <dl class="cap-metrics">
          <div class="cap-metric"><dt>Per&iacute;odos cadastrados</dt><dd><%=FormatNumber(periodTotal,0)%></dd></div>
          <div class="cap-metric"><dt>Processados</dt><dd><%=FormatNumber(periodProcessed,0)%></dd></div>
          <div class="cap-metric"><dt>Pendentes de processamento</dt><dd><%=FormatNumber(periodPending,0)%></dd></div>
          <div class="cap-metric"><dt>Primeiro pagamento pendente</dt><dd><%If IsDate(periodNext) Then%><%=ArrumaData(periodNext)%><%Else%>&mdash;<%End If%></dd></div>
        </dl>
        <%If periodTotal > 0 Then%>
        <section class="cap-chart" aria-labelledby="period-chart-title">
          <h3 id="period-chart-title">Processamento dos per&iacute;odos cadastrados</h3>
          <div class="cap-chart-donut" style="--cap-share:<%=CInt(periodShare)%>%" aria-hidden="true"><strong><%=FormatNumber(periodShare,1)%>%</strong></div>
          <p><strong><%=FormatNumber(periodShare,1)%>% processados</strong> &middot; <%=FormatNumber(periodProcessed,0)%> de <%=FormatNumber(periodTotal,0)%> per&iacute;odos. <%=FormatNumber(periodPending,0)%> pendentes.</p>
        </section>
        <%End If%>
      </section>
      <section class="dashboard-panel period-panel" aria-labelledby="period-list-title">
        <header>
          <div><h2 id="period-list-title">Per&iacute;odos cadastrados</h2><p class="panel-description"><%=FormatNumber(rs.RecordCount,0)%> no filtro atual</p></div>
          <div class="period-actions">
            <form method="get" class="period-filter" aria-label="Filtrar per&iacute;odos">
              <label for="period-status">Situa&ccedil;&atilde;o</label>
              <select name="situacao" id="period-status" class="form-select w-auto">
                <option value="">Todos</option>
                <option value="pendentes"<%If periodStatus = "pendentes" Then%> selected<%End If%>>Pendentes</option>
                <option value="processados"<%If periodStatus = "processados" Then%> selected<%End If%>>Processados</option>
              </select>
              <button class="btn btn-outline-secondary" type="submit">Filtrar</button>
            </form>
            <a class="btn btn-outline-secondary" href="cap_fechamentos_periodos_imprimir.asp" target="_blank" rel="noopener">Imprimir listagem completa</a>
          </div>
        </header>
        <%If rs.Eof Then%>
        <p class="empty-state">Nenhum per&iacute;odo encontrado.</p>
        <%Else%>
        <div class="period-scroll" role="region" aria-labelledby="period-list-title" tabindex="0">
          <table class="period-table">
            <thead><tr><th scope="col">C&oacute;digo</th><th scope="col">Per&iacute;odo de vendas</th><th scope="col">Limite de lan&ccedil;amento</th><th scope="col">Pagamento</th><th scope="col">Processamento</th><th scope="col">Cadastro</th><th scope="col">A&ccedil;&otilde;es</th></tr></thead>
            <tbody>
            <%For periodRow = 1 To rs.PageSize
                If rs.Eof Then Exit For%>
              <tr>
                <td><strong><%=rs("cod_fechamento_periodo")%></strong></td>
                <td class="period-date"><%=ArrumaData(rs("data_inicio"))%> a <%=ArrumaData(rs("data_termino"))%></td>
                <td class="period-date"><%=ArrumaData(rs("data_termino_lancamento_vendas"))%></td>
                <td class="period-date"><strong><%=ArrumaData(rs("data_pagamento"))%></strong></td>
                <td><%If IsNull(rs("data_processado_pagamento")) Then%><span class="period-badge">Pendente</span><%Else%><span class="period-badge period-done">Processado</span><span class="period-meta period-date"><%=ArrumaData(rs("data_processado_pagamento"))%></span><%End If%></td>
                <td><%=Server.HTMLEncode(rs("nome_cadastro") & "")%><span class="period-meta period-date"><%=ArrumaData(rs("modificado"))%></span></td>
                <td><a class="btn btn-outline-danger btn-sm" href="cap_fechamentos_periodos_formulario.asp?cod_fechamento_periodo=<%=rs("cod_fechamento_periodo")%>&amp;comando=excluir" aria-label="Excluir per&iacute;odo <%=rs("cod_fechamento_periodo")%>">Excluir</a></td>
              </tr>
            <%rs.MoveNext
            Next%>
            </tbody>
          </table>
        </div>
        <%If rs.PageCount > 1 Then%>
        <nav class="period-paging" aria-label="P&aacute;ginas dos per&iacute;odos">
          <%If periodPage > 1 Then%><a class="btn btn-outline-secondary btn-sm" href="?situacao=<%=periodStatus%>&amp;pagina=<%=periodPage-1%>">Anterior</a><%End If%>
          <span>P&aacute;gina <%=periodPage%> de <%=rs.PageCount%></span>
          <%If periodPage < rs.PageCount Then%><a class="btn btn-outline-secondary btn-sm" href="?situacao=<%=periodStatus%>&amp;pagina=<%=periodPage+1%>">Pr&oacute;xima</a><%End If%>
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
