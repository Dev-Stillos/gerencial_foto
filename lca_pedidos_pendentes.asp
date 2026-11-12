<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!--
'*** Instru��es
'1) Relaciona todos os vendedores
'2) Mostra somentes os vendedores que possuem pedidos pendentes
-->
<%
Dim msg, cod_vendedor
msg = Request.QueryString("msg")

AbrirConexao()
Dim valida_permissao_para_alteracao_pedidos, valida_entrada_venda_vendedores, valida_pagina_baixa_diretoria
valida_permissao_para_alteracao_pedidos = validaUsuario(1,37,,true)
valida_entrada_venda_vendedores = validaUsuario(2,,70,true)
valida_pagina_baixa_diretoria = validaUsuario(1,66,,true)

FecharConexao()

Dim queueIdPattern, queueVendor, queueVendorQuery, queueReceipt
Set queueIdPattern = New RegExp
queueIdPattern.Pattern = "^[1-9][0-9]{0,9}$"
queueVendor = Request.QueryString("vid")
queueVendorQuery = Request.QueryString("cod_vendedor")
If queueVendor <> "all" And Not queueIdPattern.Test(queueVendor) Then queueVendor = ""
If Not queueIdPattern.Test(queueVendorQuery) Then queueVendorQuery = ""
If Request.Form("btnV") <> Empty And valida_entrada_venda_vendedores And queueVendorQuery <> "" Then

Dim PEDIDO_CHEGADA_VALIDA, PEDIDOS_RECIBO, COD_VENDEDOR_RECIBO
	AbrirConexao()
	Set rs=objConexao.Execute("SELECT COD_PEDIDO FROM TB_PEDIDOS WHERE DATA_DESCEU_DIRETORIA IS NULL AND DATA_CHEGADA IS NULL AND FINALIZA_CORRECAO <> 3 AND FINALIZA_VENDEDOR = 1 AND COD_VENDEDOR = " & queueVendorQuery)
		
		PEDIDOS_RECIBO = ""
		COD_VENDEDOR_RECIBO = Request.QueryString("cod_vendedor")
		
		While not rs.Eof
			
			PEDIDO_CHEGADA_VALIDA = formatNumber(0&Request.Form("PEDIDO_CHEGADA_"&rs("COD_PEDIDO")),0)
			
			'Response.Write(Request.Form("PEDIDO_CHEGADA_"&rs("COD_PEDIDO"))&"<br>")
			
			If int(PEDIDO_CHEGADA_VALIDA) = 1 Then
				
				sql = "UPDATE TB_PEDIDOS SET DATA_CHEGADA = NOW( ) WHERE COD_PEDIDO ="&rs("COD_PEDIDO")&" LIMIT 1"
				objConexao.Execute(sql)
				
				PEDIDOS_RECIBO = PEDIDOS_RECIBO&", "&rs("COD_PEDIDO")
				
			End If
			
		rs.MoveNext()
		Wend
		
	rs.Close
	Set rs=Nothing
	FecharConexao()
		
	If Len(PEDIDOS_RECIBO) > 2 Then
        msg = "Entrada de pedidos efetuada com sucesso!"
        PEDIDOS_RECIBO = Mid(PEDIDOS_RECIBO,3)
    End If
	'Response.Redirect("lca_pedidos_pendentes.asp?msg=")
End If

If Request.QueryString("comando") = "limparChegada" And valida_entrada_venda_vendedores And queueIdPattern.Test(Request.QueryString("COD_PEDIDO")) Then
	
	AbrirConexao()
		objConexao.Execute("UPDATE TB_PEDIDOS SET DATA_CHEGADA = NULL WHERE COD_PEDIDO ='"&Request.QueryString("COD_PEDIDO")&"' LIMIT 1")
	FecharConexao()	
	
	Response.Redirect("lca_pedidos_pendentes.asp?msg=Chegada limpa com sucesso!")
End If

%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Pedidos pendentes | Portal Stillos</title>
<link rel="stylesheet" href="css/base.css" media="screen">
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

.order-queue-metrics { margin-bottom:16px; }
.order-queue-panel { height:auto; margin-bottom:16px; }
.order-queue-panel header { display:flex; align-items:end; justify-content:space-between; flex-wrap:wrap; gap:12px; margin-bottom:12px; }
.order-queue-panel input[type=search] { min-width:min(240px,100%); }
.order-queue-panel .period-table { min-width:0; }
.order-queue-panel .period-table th,.order-queue-panel .period-table td { padding:8px 12px; }
.order-queue-panel .period-table td:first-child { overflow-wrap:anywhere; }
.order-queue-panel .period-table td:last-child { text-align:right; }
.order-queue-chart { margin-bottom:16px; }
.order-queue-detail { height:auto; margin-bottom:16px; min-width:0; }
.order-queue-detail-scroll { overflow-x:auto; }
.order-queue-detail table[border="1"] { border-collapse:collapse; border-color:#e4dfeb; }
.order-queue-detail td { overflow-wrap:anywhere; }
.style1 { color:#275e42; font-weight:700; }
.style2 { color:#8a2149; font-weight:700; }
[hidden] { display:none !important; }


.queue-mobile-selection { display:flex; margin:0 0 12px; }
.queue-vendor-name { font-size:1.2em; margin:0 0 8px; color:#242738; }
.order-queue-detail .queue-orders-table { width:100%; min-width:900px; border-collapse:collapse; border:1px solid #e4dfeb; }
.queue-orders-table thead th { padding:10px; text-align:left; background:#f1edf6; color:#3e344e; }
.queue-orders-table > tbody > tr > td { padding:10px; vertical-align:top; border-bottom:1px solid #e4dfeb; }
.queue-orders-table > tbody > tr:nth-child(even) { background:#faf9fc; }
.queue-col-order,.queue-col-amount,.queue-col-dates,.queue-col-people a { white-space:nowrap; }
.queue-order-field + .queue-order-field { margin-top:4px; }
.queue-sub-label { display:block; color:#53566b; font-size:.82em; font-weight:650; }
.queue-col-dates .queue-sub-label { margin-right:4px; }
.queue-status-negative { background:#f9e7ed; color:#86234a; }
.queue-action-note { color:#53566b; font-size:.85em; }
.queue-detail-footer { display:flex; flex-wrap:wrap; justify-content:space-between; align-items:center; gap:10px; margin-top:12px; }
.queue-preview { width:min(920px,calc(100% - 24px)); max-height:calc(100dvh - 24px); padding:16px; border:1px solid #e4dfeb; border-radius:12px; }
.queue-preview::backdrop { background:#24273899; }
.queue-preview-actions { display:flex; flex-wrap:wrap; gap:8px; margin-bottom:12px; }
.queue-report-head { display:flex; justify-content:space-between; align-items:center; gap:12px; padding-bottom:8px; margin-bottom:12px; border-bottom:2px solid #b20b61; }
.queue-report-head img { width:130px; max-width:30%; height:auto; border-radius:6px; background:#292b40; }
.queue-report-head h2,.queue-report-head p { margin:0; }
.queue-print-order { padding:8px 0; border-bottom:1px solid #e4dfeb; break-inside:avoid; }
.queue-print-order h3 { font-size:1.1em; margin:0 0 4px; }
.queue-print-fields { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:5px 12px; margin:0; }
.queue-print-fields dd { margin:0; overflow-wrap:anywhere; }
#queue-print-sheet { display:none; }

@media screen and (max-width:1190px) {
  .order-queue-detail .queue-orders-table,.queue-orders-table > tbody { display:block; min-width:0; }
  .queue-orders-table > thead { display:none; }
  .queue-order-row { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:6px; margin:8px; padding:8px; border:1px solid #e4dfeb; border-radius:8px; }
  .queue-orders-table > tbody > tr > td { min-width:0; white-space:normal; border:0; padding:4px; }
  .queue-order-row > td::before { content:attr(data-queue-heading); display:block; font-weight:700; color:#242738; margin-bottom:4px; }
  .queue-col-people { grid-column:span 2; }
}
@media screen and (max-width:550px) {
  .queue-order-row { grid-template-columns:repeat(2,minmax(0,1fr)); }
  .queue-col-people { grid-column:1/-1; }
}
@page queue-orders { size:A4 portrait; margin:12mm; }
@media print {
  body.queue-printing > :not(#queue-print-sheet) { display:none !important; }
  body.queue-printing #queue-print-sheet { display:block; page:queue-orders; color:#111; }
  #queue-print-sheet .queue-report-head img { width:30mm; background:transparent; filter:brightness(0); }
  #queue-print-sheet .queue-report-head { border-color:#444; break-after:avoid; }
}
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
      <div class="topbar-title">Portal Stillos <span>/ Pedidos pendentes</span></div>
      <div class="topbar-actions"><a class="user-name" href="adm_log_usuarios.asp"><%=Server.HTMLEncode(Session("NOME") & "")%></a><a href="manual_sistema.asp" target="_blank" rel="noopener">Manual</a><a class="btn btn-outline-secondary btn-sm" href="adm_sair.asp">Sair</a></div>
    </header>
    <main class="dashboard-main ppv-main" id="conteudo">


      <header class="dashboard-heading">
        <div><h1>Pedidos pendentes</h1><p>Acompanhe os pedidos por vendedor que ainda n&atilde;o foram baixados para a diretoria.</p></div>
        <div class="period-actions">
          <%If queueVendor <> "" Or queueVendorQuery <> "" Then%>
            <a class="btn btn-outline-secondary" href="lca_pedidos_pendentes.asp">Voltar aos vendedores</a>
          <%Else%>
            <a class="btn btn-outline-secondary" href="adm_index.asp">Voltar ao painel</a>
          <%End If%>
          <%If valida_pagina_baixa_diretoria Then%><a class="btn btn-primary" href="lca_pedidos_pendentes_baixar_para_diretoria.asp">Baixar fichas</a><%End If%>
        </div>
      </header>
      <div class="search-panel"><strong>Busca no portal</strong>
        <!-- #include file="includes/dashboard_busca.asp" -->
      </div>
      <div class="session-notices">
        <!-- #include file="includes/dashboard_avisos.asp" -->
      </div>
      <%If Len(msg & "") > 0 Then%><p class="alert alert-info" role="status"><%=Server.HTMLEncode(msg & "")%></p><%End If%>
      <%If Len(PEDIDOS_RECIBO & "") > 0 Then%><p><a class="btn btn-outline-secondary" target="_blank" rel="noopener" href="lca_pedidos_pendentes_recibo_venda.asp?pedidos=<%=Server.URLEncode(PEDIDOS_RECIBO)%>&amp;cod_vendedor=<%=queueVendorQuery%>">Imprimir recibo de entrada</a></p><%End If%>
<%
Dim vendedorId, COUNT_PEDIDOS
vendedorId = queueVendor
codVendedor = queueVendorQuery
COUNT_PEDIDOS = 0
AbrirConexao()
If vendedorId = "" And codVendedor = "" Then
    Dim rs5, sql1, queueRows, queueCount, queueIndex, queueMaximum, queueLeader
    sql1 = "SELECT v.COD_VENDEDOR, v.NOME_VENDEDOR, COUNT(DISTINCT p.COD_PEDIDO) AS PENDENTES FROM VW_PEDIDOS_COMPLETO p JOIN TB_VENDEDORES v ON v.COD_VENDEDOR = p.COD_VENDEDOR WHERE p.DATA_DESCEU_DIRETORIA IS NULL AND p.FINALIZA_CORRECAO <> 3 AND p.FINALIZA_VENDEDOR = 1 GROUP BY v.COD_VENDEDOR, v.NOME_VENDEDOR ORDER BY v.NOME_VENDEDOR"
    Set rs5 = objConexao.Execute(sql1)
    queueCount = 0
    queueMaximum = 0
    If Not rs5.Eof Then
        queueRows = rs5.GetRows()
        queueCount = UBound(queueRows, 2) + 1
        For queueIndex = 0 To queueCount - 1
            COUNT_PEDIDOS = COUNT_PEDIDOS + CLng(queueRows(2, queueIndex))
            If CLng(queueRows(2, queueIndex)) > queueMaximum Then
                queueMaximum = CLng(queueRows(2, queueIndex))
                queueLeader = queueRows(1, queueIndex) & ""
            End If
        Next
    End If
    rs5.Close
    Set rs5 = Nothing
%>
<section class="order-queue-metrics" aria-labelledby="queue-overview">
  <h2 id="queue-overview" class="mb-2">Vis&atilde;o geral</h2>
  <dl class="cap-metrics">
    <div class="cap-metric"><dt>Pedidos na fila</dt><dd><%=FormatNumber(COUNT_PEDIDOS,0)%></dd></div>
    <div class="cap-metric"><dt>Vendedores com pedidos pendentes</dt><dd><%=FormatNumber(queueCount,0)%></dd></div>
    <div class="cap-metric"><dt>Maior fila individual</dt><dd><%=FormatNumber(queueMaximum,0)%></dd><small><%=Server.HTMLEncode(queueLeader & "")%></small></div>
  </dl>
</section>
<section class="cap-chart order-queue-chart" id="queue-chart" aria-labelledby="queue-chart-title" hidden>
  <h3 id="queue-chart-title">Vendedores com mais pedidos pendentes</h3>
  <ul class="cap-chart-bars" id="queue-bars"></ul>
</section>
<section class="dashboard-panel order-queue-panel" aria-labelledby="queue-list-title">
  <header>
    <div><h2 id="queue-list-title">Fila por vendedor</h2><p class="panel-description" id="queue-count" role="status"><%=queueCount%> vendedores</p></div>
    <%If queueCount > 0 Then%>
    <div class="period-actions">
      <div><label for="queue-search">Buscar vendedor</label><input class="form-control" id="queue-search" type="search"></div>
      <a class="btn btn-outline-secondary" href="lca_pedidos_pendentes.asp?vid=all">Visualizar todos os pedidos</a>
    </div>
    <%End If%>
  </header>
  <%If queueCount = 0 Then%>
  <p class="empty-state">Nenhum pedido pendente na fila.</p>
  <%Else%>
  <p id="queue-empty" class="empty-state" hidden>Nenhum vendedor encontrado.</p>
  <div class="period-scroll" role="region" aria-labelledby="queue-list-title" tabindex="0">
    <table class="period-table">
      <thead><tr><th scope="col">Vendedor</th><th scope="col">Pedidos pendentes</th><th scope="col">A&ccedil;&otilde;es</th></tr></thead>
      <tbody>
      <%For queueIndex = 0 To queueCount - 1%>
      <tr data-queue-count="<%=queueRows(2, queueIndex)%>">
        <td><strong><%=Server.HTMLEncode(queueRows(1, queueIndex) & "")%></strong></td>
        <td><%=FormatNumber(queueRows(2, queueIndex),0)%></td>
        <td><a class="btn btn-outline-secondary btn-sm" href="lca_pedidos_pendentes.asp?vid=<%=queueRows(0, queueIndex)%>">Ver pedidos</a></td>
      </tr>
      <%Next%>
      </tbody>
    </table>
  </div>
  <%End If%>
</section>
<%

		Else
		  
		  Dim rs, rs2, rs3, rs4, sql
		  '1) Consulta todos os vendedores, por�m s� mostra na tela os que possuem pedidos pendentes
		  If vendedorId = "all" then
		 	sql = "SELECT COD_VENDEDOR, NOME_VENDEDOR FROM TB_VENDEDORES ORDER BY NOME_VENDEDOR"
		  Elseif codVendedor <> "" then
			sql = "SELECT COD_VENDEDOR, NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR = "&codVendedor&" ORDER BY NOME_VENDEDOR"
		  Else
			sql = "SELECT COD_VENDEDOR, NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR = "&vendedorId&" ORDER BY NOME_VENDEDOR"
		  End If

		  Set rs = objConexao.Execute(sql)
			Dim COUNT_PENDENTES, COUNT_FINALIZADOS 
			COUNT_PENDENTES		= 0
			COUNT_FINALIZADOS	= 0
		  	
			While not rs.Eof 
		  	
			  
		 		 	'2) Verifica se h&aacute; registros pendentes do vendedor acima
			'		Filtros: Vendedor, N&atilde;o desceu para diretoria (Pendentes) e Pedidos Ativos
			sql = "SELECT * FROM VW_PEDIDOS_COMPLETO WHERE COD_VENDEDOR="&rs("COD_VENDEDOR")&" AND DATA_DESCEU_DIRETORIA IS NULL AND FINALIZA_CORRECAO <> 3 AND FINALIZA_VENDEDOR = 1 GROUP BY COD_PEDIDO ORDER BY NOME_FORMANDO"	  	
			Set rs2 = objConexao.Execute(sql)
			
			If not rs2.Eof Then
					
			DIM COUNT_PENDENTES_POR_VENDEDOR, COUNT_FINALIZADOS_POR_VENDEDOR
			COUNT_PENDENTES_POR_VENDEDOR 	= 0
			COUNT_FINALIZADOS_POR_VENDEDOR 	= 0
			%>
			<section class="dashboard-panel order-queue-detail"><div class="order-queue-detail-scroll"><form id="queue-seller-<%=rs("COD_VENDEDOR")%>" name="poliimpressao" method="post" action="lca_pedidos_pendentes_poliimpressao.asp" target="_blank" style="margin:0px">
				<table style="margin-top:10px" cellpadding="5" cellspacing="0" align="center" width="100%">
			<tr>
			<td>
		      <h2 class="queue-vendor-name"><%=Server.HTMLEncode(Ucase(rs("NOME_VENDEDOR") & ""))%></h2></td>
			</tr>
			</table>
			    <div class="queue-mobile-selection period-actions">
  <button type="button" class="btn btn-outline-secondary btn-sm" onclick="selecaoVendedor_<%=rs("COD_VENDEDOR")%>()">Selecionar ou limpar impress&atilde;o</button>
  <%If valida_entrada_venda_vendedores Then%><button type="button" class="btn btn-outline-secondary btn-sm" onclick="selecaoVendedorChegada_<%=rs("COD_VENDEDOR")%>()">Selecionar ou limpar entrada</button><%End If%>
</div>
<input type="hidden" name="VERIFICA_SELECIONADO_<%=rs("COD_VENDEDOR")%>" value="0">
<input type="hidden" name="VERIFICA_SELECIONADO_CHEGADA_<%=rs("COD_VENDEDOR")%>" value="0">
<table class="queue-orders-table">
  <thead><tr>
    <th scope="col">Pedido</th><th scope="col">Formandos</th><th scope="col">Venda</th>
    <th scope="col">Total recebido</th><th scope="col">Datas</th><th scope="col">Situa&ccedil;&atilde;o e a&ccedil;&otilde;es</th>
  </tr></thead>
  <tbody>
                  <%
				Dim JAVA_CHECKED_TRUE, JAVA_CHECKED_FALSE, JAVA_CHEGADA_CHECKED_TRUE, JAVA_CHEGADA_CHECKED_FALSE
				JAVA_CHECKED_TRUE	= ""
				JAVA_CHECKED_FALSE	= ""
				JAVA_CHEGADA_CHECKED_TRUE 	= ""
				JAVA_CHEGADA_CHECKED_FALSE 	= ""
				  
				  While not rs2.Eof
				  
				  DATA_CHEGADA = ""&rs2("DATA_CHEGADA")
				  %>
                  <tr class="queue-order-row">
  <td class="queue-col-order" data-queue-heading="Pedido">
    <div class="queue-order-field" data-queue-label="Pedido"><div >
                        <%If int(rs2("FINALIZA_CORRECAO")) = 1 Then%>
                      
						     <table width="100%" border="0" cellspacing="0" cellpadding="0">
							  <tr>
							    <td width="72%"><div ><a href="lca_busca_detalhamento_pedido.asp?cod_pedido=<%=rs2("COD_PEDIDO")%>"><%=rs2("COD_PEDIDO")%></a><%If int(rs2("PEDIDO_EXTRA")) = 1 Then%>&nbsp;<em class="textoMenor">(Extra)</em><%End If%></div></td>
							    <td width="28%"><div >
							      <input class="inputRadio" type="checkbox" name="PEDIDO_<%=rs2("COD_PEDIDO")%>" value="1" />		
							    </div></td>
							  </tr>
							</table>
						
						
						
                        <%Else%>
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
							  <tr>
							    <td width="72%"><div ><%=rs2("COD_PEDIDO")%><%If int(rs2("PEDIDO_EXTRA")) = 1 Then%>&nbsp;<em class="textoMenor">(Extra)</em><%End If%></div></td>
							    <td width="28%">
									<div >
							&nbsp;		</div>	</td>
							  </tr>
							</table>
	                        <%End If%>
	                    </div></div>
  </td>
  <td class="queue-col-people" data-queue-heading="Formandos">
    <div class="queue-order-field" data-queue-label="Formandos"><div >
                        <%
						sql = 	"SELECT TB_PEDIDOS_FORMANDOS.COD_FORMANDO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO "&_
								"FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON "&_
								"TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO "&_
								"WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&rs2("COD_PEDIDO")&" ORDER BY NOME_FORMANDO"
						Set rs3 = objConexao.Execute(sql)
							
					
						While not rs3.Eof
											
						%>
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="23%"><div > <a href="dip_busca_detalhamento_formando.asp?cod_formando=<%=rs3("COD_FORMANDO")%>"> <%=digito(rs3("COD_FORMANDO"))%> </a> </div></td>
                            <td width="57%"><div ><%=rs3("NOME_FORMANDO")%></div></td>
                          </tr>
                        </table>
                      <%
				rs3.MoveNext()
				Wend
				
			rs3.Close
			Set rs3 = Nothing
			%>
                    </div></div>
  </td>
  <td class="queue-col-sale" data-queue-heading="Venda">
    <div class="queue-order-field" data-queue-label="Pagamento"><small class="queue-sub-label">Pagamento</small>
                    <%
					If int(rs2("FINALIZA_CORRECAO")) = 1 Then
					
						If int(rs2("FORMA_PAGAMENTO")) = 1 Then
							Response.Write("&Agrave; vista")
						Elseif int(rs2("FORMA_PAGAMENTO")) = 2 Then
							Response.Write("&Agrave; prazo")					
						End If
					Else
					%>
                    <em class="textoMenor">N&atilde;o dispon&iacute;vel</em>
                    <%
					End If
					%>                    </div>
    <div class="queue-order-field" data-queue-label="Tipo de venda"><small class="queue-sub-label">Tipo de venda</small> <%
					If int(rs2("FINALIZA_CORRECAO")) = 1 Then
					
						If int(rs2("TIPO_VENDA")) = 1 Then
							Response.Write("Normal")
						Elseif int(rs2("TIPO_VENDA")) = 2 Then
							Response.Write("Repasse")
						Elseif int(rs2("TIPO_VENDA")) = 3 Then
							Response.Write("Repasse Especial")
						End If
					Else
					%>
                    <em class="textoMenor">N&atilde;o dispon&iacute;vel</em>
                    <%
					End If
					%></div>
    <div class="queue-order-field" data-queue-label="Ajuda de custo"><small class="queue-sub-label">Ajuda de custo</small><%
					If int(rs2("FINALIZA_CORRECAO")) = 1 Then
						VALOR_AJUDA_CUSTO = formatNumber(0&rs2("VALOR_AJUDA_CUSTO"))
						'VALOR_AJUDA_CUSTO = formatCurrency(VALOR_AJUDA_CUSTO)
						Response.Write(VALOR_AJUDA_CUSTO)
					Else
					%>
                    <em class="textoMenor">N&atilde;o dispon&iacute;vel</em>
                    <%
					End If
					%></div>
    <div class="queue-order-field" data-queue-label="Grupo especial"><small class="queue-sub-label">Grupo especial</small>  <%
					If int(rs2("FINALIZA_CORRECAO")) = 1 Then
					
						If int(rs2("COD_GRUPO_VENDEDORES")) = 1 Then
							Response.Write("Renato")
						Elseif int(rs2("COD_GRUPO_VENDEDORES")) = 2 Then
							Response.Write("M�rio")	
						Elseif int(rs2("COD_GRUPO_VENDEDORES")) = 3 Then
							Response.Write("Paulo M.")	
						Elseif int(rs2("COD_GRUPO_VENDEDORES")) = 4 Then
              Response.Write("Vicente")	
						Elseif int(rs2("COD_GRUPO_VENDEDORES")) = 0 Then
							Response.Write("N&atilde;o h&aacute;")	
						End If
					
					Else
					%>
                    <em class="textoMenor">N&atilde;o dispon&iacute;vel</em>
                    <%
					End If
					%></div>
  </td>
  <td class="queue-col-amount" data-queue-heading="Total recebido">
    <div class="queue-order-field" data-queue-label="Total recebido"><div >
                      <%
         objConexao.Execute("SET SQL_BIG_SELECTS=1")
			Set rs4 = objConexao.Execute("SELECT * FROM VW_PEDIDOS_TOTAL_RECEBIDO WHERE COD_PEDIDO="&rs2("COD_PEDIDO"))
			Dim PAGAMENTO_EM_DINHEIRO, XTOTAL_CHEQUE, XTOTAL_CARTAO, XTOTAL_BOLETO, TOTAL_PAGO
			
			TOTAL_PAGO = 0
				
				If not rs4.Eof Then
				
					PAGAMENTO_EM_DINHEIRO	= formatNumber(0&rs4("PAGAMENTO_EM_DINHEIRO"))/1
					XTOTAL_CHEQUE			= formatNumber(0&rs4("XTOTAL_CHEQUE"))/1
					XTOTAL_CARTAO			= formatNumber(0&rs4("XTOTAL_CARTAO"))/1
					XTOTAL_BOLETO			= cdbl(0&rs4("XTOTAL_BOLETO"))/1
					TOTAL_PAGO				= (PAGAMENTO_EM_DINHEIRO + XTOTAL_CHEQUE + XTOTAL_CARTAO + XTOTAL_BOLETO)/1
					
					Response.Write(formatNumber(TOTAL_PAGO))

				End If
			
			rs4.Close()
			Set rs4 = Nothing
			%>
                    </div></div>
  </td>
  <td class="queue-col-dates" data-queue-heading="Datas">
    <div class="queue-order-field" data-queue-label="Lan&ccedil;amento"><small class="queue-sub-label">Lan&ccedil;amento</small><div ><%=ArrumaData(rs2("DATA_CADASTRO"))%></div></div>
    <div class="queue-order-field" data-queue-label="Chegada"><small class="queue-sub-label">Chegada</small><div >
                    <%If Len(DATA_CHEGADA) > 3 Then%>
					<%=ArrumaData(DATA_CHEGADA)%>
                    	<%If int(rs2("FINALIZA_CORRECAO")) = 2 and valida_entrada_venda_vendedores Then%>
                        <a href="javascript:limparChegada(<%=rs2("COD_PEDIDO")%>)"><img src="images/icons/icon_red.gif" border="0" /></a>
                        <%End If%>
                    <%Else%>
                    <input class="inputRadio" type="checkbox" name="PEDIDO_CHEGADA_<%=rs2("COD_PEDIDO")%>" value="1" />
                    <%End If%>
                    </div></div>
    <div class="queue-order-field" data-queue-label="Vendedor recebe em"><small class="queue-sub-label">Vendedor recebe em</small><div >
                        <%If rs2("VENDEDOR_RECEBE_EM") <> "" Then%>
                        <%=arrumaData(rs2("VENDEDOR_RECEBE_EM"))%>
                        <%Else%>
                      -
                      <%End iF%>
                    </div></div>
  </td>
  <td class="queue-col-actions" data-queue-heading="Situa&ccedil;&atilde;o e a&ccedil;&otilde;es">
    <div class="queue-order-field" data-queue-label="Conferido por"><small class="queue-sub-label">Conferido por</small><%
					If int(rs2("FINALIZA_CORRECAO")) = 1 Then
					
						Set rs4=objConexao.Execute("SELECT NOME FROM TB_ADM_USUARIOS WHERE COD_USUARIO = "&sqlTexto(rs2("CONFERIDO_POR_COD_USUARIO")))	
							
							If not rs4.Eof Then
							
									Response.Write(Left(rs4("nome"),3))
							Else
							
									Response.Write("-")
							
							End If
							
						
						rs4.Close
						Set rs4=Nothing
					
					Else
					%>
                    <em class="textoMenor">N&atilde;o dispon&iacute;vel</em>
                    <%
					End If
					%></div>
    <div class="queue-order-field" data-queue-label="Situa&ccedil;&atilde;o"><small class="queue-sub-label">Situa&ccedil;&atilde;o</small><div >
                        <%If rs2("FINALIZA_CORRECAO") = 1 Then
						COUNT_FINALIZADOS_POR_VENDEDOR = COUNT_FINALIZADOS_POR_VENDEDOR + 1
						
						%>
                        <span class="period-badge period-done" title="<%=ArrumaData(rs2("DATA_CORRIGIDO_DEP_FOTO"))%>">Finalizado</span>
                        <%Else
						
						COUNT_PENDENTES_POR_VENDEDOR = COUNT_PENDENTES_POR_VENDEDOR + 1	
									
								'Verifica se est&aacute; aguardando aprova&ccedil;&atilde;o			
								Set rs3=objConexao.Execute("SELECT * FROM TB_APROVACOES WHERE parametro = "&sqlTexto(rs2("COD_PEDIDO")))
								
									If not rs3.Eof Then
									
											If int(rs3("status")) = 1 Then%>
                               	  <span class="period-badge period-done" title="Aprovado para finalizar a venda">Aprovado</span>
                                            
                                  <%Elseif int(rs3("status")) = 2 Then%>
                               	  <span class="period-badge queue-status-negative">Reprovado</span>
                                            <%Else%>
                               	  <span class="period-badge ">Aguardando aprova&ccedil;&atilde;o</span>
                                  <%
											End if
									Else%>
                               	  <span class="period-badge queue-status-negative">Pendente de corre&ccedil;&atilde;o</span>
                                  <%
									End If
								
								rs3.Close
								Set rs3=Nothing
								%>
                        <%End If%>
                    </div></div>
    <div class="queue-order-field" data-queue-label="Corre&ccedil;&atilde;o"><div >
                    <%If Len(DATA_CHEGADA) > 3 Then%>
                    	<%If valida_permissao_para_alteracao_pedidos Then%>
                    		<a class="btn btn-outline-secondary btn-sm" href="lca_pedidos_pendentes_alteracoes_part_1.asp?cod_pedido=<%=rs2("COD_PEDIDO")%>">Corrigir pedido</a>
                        <%Else%>
                    	 -
                        <%End If%>
                    <%Else%>
                    <span class="queue-action-note">Registre a entrada para corrigir.</span>
                    <%End If%>
                    </div></div>
  </td>
</tr>
                <%
				
				If int(rs2("FINALIZA_CORRECAO")) = 1 Then
				JAVA_CHECKED_TRUE	= JAVA_CHECKED_TRUE&" d.PEDIDO_"&rs2("COD_PEDIDO")&".checked=""checked""; "
				JAVA_CHECKED_FALSE	= JAVA_CHECKED_FALSE&" d.PEDIDO_"&rs2("COD_PEDIDO")&".checked=""""; "
				End If
				
				If Len(DATA_CHEGADA) < 3 Then
				JAVA_CHEGADA_CHECKED_TRUE	= JAVA_CHEGADA_CHECKED_TRUE&" d.PEDIDO_CHEGADA_"&rs2("COD_PEDIDO")&".checked=""checked""; "
				JAVA_CHEGADA_CHECKED_FALSE	= JAVA_CHEGADA_CHECKED_FALSE&" d.PEDIDO_CHEGADA_"&rs2("COD_PEDIDO")&".checked=""""; "
				End If
				'
				rs2.MoveNext()
				Wend
				%>
				<script language="javascript">
						function selecaoVendedor_<%=rs("COD_VENDEDOR")%>(){
							var d = document.getElementById('queue-seller-<%=rs("COD_VENDEDOR")%>')
							if(d.VERIFICA_SELECIONADO_<%=rs("COD_VENDEDOR")%>.value == 0){
								d.VERIFICA_SELECIONADO_<%=rs("COD_VENDEDOR")%>.value = 1;
								 selecionartodasVendedor_<%=rs("COD_VENDEDOR")%>();
							}else{
								d.VERIFICA_SELECIONADO_<%=rs("COD_VENDEDOR")%>.value = 0;
								tirarselecaodeTodasVendedor_<%=rs("COD_VENDEDOR")%>()
							}
						}
				
						function selecionartodasVendedor_<%=rs("COD_VENDEDOR")%>(){
							var d = document.getElementById('queue-seller-<%=rs("COD_VENDEDOR")%>')		
							<%=JAVA_CHECKED_TRUE%>
						}
						
						function tirarselecaodeTodasVendedor_<%=rs("COD_VENDEDOR")%>(){
							var d = document.getElementById('queue-seller-<%=rs("COD_VENDEDOR")%>')
							<%=JAVA_CHECKED_FALSE%>
						}
						
						//DATA DE CHEGADA
						
						function selecaoVendedorChegada_<%=rs("COD_VENDEDOR")%>(){
							var d = document.getElementById('queue-seller-<%=rs("COD_VENDEDOR")%>')
							if(d.VERIFICA_SELECIONADO_CHEGADA_<%=rs("COD_VENDEDOR")%>.value == 0){
								d.VERIFICA_SELECIONADO_CHEGADA_<%=rs("COD_VENDEDOR")%>.value = 1;
								 selecionartodasVendedorChegada_<%=rs("COD_VENDEDOR")%>();
							}else{
								d.VERIFICA_SELECIONADO_CHEGADA_<%=rs("COD_VENDEDOR")%>.value = 0;
								tirarselecaodeTodasVendedorChegada_<%=rs("COD_VENDEDOR")%>()
							}
						}
						
						
						function selecionartodasVendedorChegada_<%=rs("COD_VENDEDOR")%>(){
							var d = document.getElementById('queue-seller-<%=rs("COD_VENDEDOR")%>')		
							<%=JAVA_CHEGADA_CHECKED_TRUE%>
						}
						
						function tirarselecaodeTodasVendedorChegada_<%=rs("COD_VENDEDOR")%>(){
							var d = document.getElementById('queue-seller-<%=rs("COD_VENDEDOR")%>')
							<%=JAVA_CHEGADA_CHECKED_FALSE%>
						}
						
					</script>
                </tbody></table>
                <%
                COUNT_PENDENTES = COUNT_PENDENTES + COUNT_PENDENTES_POR_VENDEDOR
                COUNT_FINALIZADOS = COUNT_FINALIZADOS + COUNT_FINALIZADOS_POR_VENDEDOR
                COUNT_PEDIDOS = COUNT_FINALIZADOS + COUNT_PENDENTES
                %>
                <div class="queue-detail-footer">
                  <div class="period-actions">
                    <button type="button" class="btn btn-outline-secondary btn-sm" data-queue-preview>Imprimir selecionados</button>
                    <%If valida_entrada_venda_vendedores Then%>
                    <button type="button" class="btn btn-primary btn-sm" onclick="confirmaChegada(<%=rs("COD_VENDEDOR")%>)">Registrar entrada e gerar recibo</button>
                    <%End If%>
                  </div>
                  <span><strong><%=COUNT_PENDENTES_POR_VENDEDOR%></strong> pendentes de corre&ccedil;&atilde;o &middot; <strong><%=COUNT_FINALIZADOS_POR_VENDEDOR%></strong> finalizados</span>
                </div>
<input type="hidden" name="btnV" value="1">
</form></div></section>
			<%
			End If
			
			rs2.Close
			Set rs2 = Nothing
			%>
			<%
		 	rs.MoveNext()
		  	Wend
	      	
		  rs.Close()
		  Set rs = Nothing
		 End If 
		  FecharConexao()
		  %>

<%If vendedorId <> "" Or codVendedor <> "" Then%>
<section class="dashboard-panel order-queue-panel" aria-label="Totais dos pedidos exibidos">
  <strong><%=COUNT_PEDIDOS%> pedidos exibidos</strong> &middot; <%=COUNT_PENDENTES%> pendentes de corre&ccedil;&atilde;o &middot; <%=COUNT_FINALIZADOS%> finalizados
</section>
<%End If%>

    </main>
    <footer class="dashboard-footer">&copy; <%=Year(Date)%> Grupo Stillos <span>Sistema administrativo interno</span></footer>
  </div>
</div>
<dialog class="queue-preview" id="queue-preview" aria-label="Pr&eacute;via dos pedidos selecionados">
  <div class="queue-preview-actions">
    <button type="button" class="btn btn-primary" id="queue-summary-print">Imprimir resumo / Salvar PDF</button>
    <button type="button" class="btn btn-outline-secondary" id="queue-official-print">Gerar PDF dos pedidos</button>
    <form method="dialog"><button class="btn btn-outline-secondary" autofocus>Fechar</button></form>
  </div>
  <p id="queue-preview-status" role="status" hidden></p>
  <article id="queue-preview-document">
    <header class="queue-report-head"><div><h2>Resumo dos pedidos selecionados</h2><p id="queue-report-vendor"></p></div><img src="images/logoparafundopreto.png" alt="Stillos"></header>
    <div id="queue-report-orders"></div>
  </article>
</dialog>
<div id="queue-print-sheet"></div>
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


function confirmaChegada(vendor) {
    if (!window.confirm('Dar entrada nas vendas selecionadas e gerar o recibo?')) return;
    var form = document.getElementById('queue-seller-' + vendor);
    form.target = '_self';
    form.action = 'lca_pedidos_pendentes.asp?cod_vendedor=' + encodeURIComponent(vendor);
    form.submit();
}
function limparChegada(order) {
    if (window.confirm('Limpar a chegada do pedido ' + order + '?')) {
        location.assign('lca_pedidos_pendentes.asp?comando=limparChegada&cod_pedido=' + encodeURIComponent(order));
    }
}
(function () {
    var search = document.getElementById('queue-search');
    if (!search) return;
    var rows = Array.from(document.querySelectorAll('[data-queue-count]'));
    var chart = document.getElementById('queue-chart');
    var bars = document.getElementById('queue-bars');
    var count = document.getElementById('queue-count');
    var empty = document.getElementById('queue-empty');
    var top = rows.slice().sort(function (a,b) { return Number(b.dataset.queueCount) - Number(a.dataset.queueCount); }).slice(0,5);
    var maximum = Number(top[0].dataset.queueCount);
    top.forEach(function (row) {
        var item = document.createElement('li');
        var label = document.createElement('div');
        var name = document.createElement('span');
        var value = document.createElement('strong');
        var track = document.createElement('div');
        var fill = document.createElement('span');
        name.textContent = row.cells[0].textContent.trim();
        value.textContent = row.cells[1].textContent.trim();
        label.append(name,value);
        track.className = 'cap-chart-track';
        track.setAttribute('aria-hidden','true');
        fill.style.width = (100 * Number(row.dataset.queueCount) / maximum) + '%';
        track.append(fill);
        item.append(label,track);
        bars.append(item);
    });
    chart.hidden = false;
    function normalize(value) {
        return value.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim();
    }
    search.addEventListener('input', function () {
        var term = normalize(search.value);
        var visible = 0;
        rows.forEach(function (row) {
            row.hidden = !normalize(row.cells[0].textContent).includes(term);
            if (!row.hidden) visible += 1;
        });
        count.textContent = visible + ' de ' + rows.length + ' vendedores';
        empty.hidden = visible !== 0;
    });
})();

(function () {
    var dialog = document.getElementById('queue-preview');
    var preview = document.getElementById('queue-preview-document');
    var status = document.getElementById('queue-preview-status');
    var list = document.getElementById('queue-report-orders');
    var print = document.getElementById('queue-summary-print');
    var official = document.getElementById('queue-official-print');
    var sheet = document.getElementById('queue-print-sheet');
    var form;
    var trigger;
    document.addEventListener('click', function (event) {
        var link = event.target.closest('[data-queue-preview]');
        if (!link) return;
        event.preventDefault();
        trigger = link;
        form = link.closest('form');
        var orders = Array.from(form.querySelectorAll('input[name^="PEDIDO_"]:not([name^="PEDIDO_CHEGADA_"]):checked'));
        status.hidden = orders.length > 0;
        status.textContent = 'Selecione pelo menos um pedido para imprimir.';
        preview.hidden = print.disabled = official.disabled = orders.length === 0;
        list.replaceChildren();
        document.getElementById('queue-report-vendor').textContent =
            form.querySelector('.queue-vendor-name').textContent + ' | ' + orders.length + ' pedido(s)';
        orders.forEach(function (order) {
            var row = order.closest('.queue-order-row');
            var section = document.createElement('section');
            var title = document.createElement('h3');
            var fields = document.createElement('dl');
            section.className = 'queue-print-order';
            fields.className = 'queue-print-fields';
            title.textContent = 'Pedido ' + order.name.slice(7);
            row.querySelectorAll('[data-queue-label]').forEach(function (cell, index) {
                if (index === 0 || index >= 11) return;
                var field = document.createElement('div');
                var label = document.createElement('dt');
                var value = document.createElement('dd');
                label.textContent = cell.dataset.queueLabel;
                var display = cell.cloneNode(true);
                var subLabel = display.querySelector('.queue-sub-label');
                if (subLabel) subLabel.remove();
                value.textContent = display.textContent.trim().replace(/\s+/g, ' ') || '—';
                field.append(label,value);
                fields.append(field);
            });
            section.append(title,fields);
            list.append(section);
        });
        dialog.showModal();
    });
    print.addEventListener('click', function () {
        var copy = preview.cloneNode(true);
        copy.removeAttribute('id');
        copy.querySelectorAll('[id]').forEach(function (element) { element.removeAttribute('id'); });
        sheet.replaceChildren(copy);
        document.body.classList.add('queue-printing');
        window.print();
    });
    official.addEventListener('click', function () {
        form.action = 'lca_pedidos_pendentes_poliimpressao.asp';
        form.target = '_blank';
        form.submit();
    });
    window.addEventListener('afterprint', function () {
        document.body.classList.remove('queue-printing');
        sheet.replaceChildren();
    });
    dialog.addEventListener('close', function () {
        if (trigger) trigger.focus();
    });
})();
</script>
</body>
</html>
