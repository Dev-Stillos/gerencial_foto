<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
 
Response.Buffer = False
Dim cod_fechamento_periodo, vendedor_recebe_em, autorizacao_diretoria_para_finalizar
cod_fechamento_periodo = Request.QueryString("cod_fechamento_periodo")
	
	AbrirConexao()

	Set rs2=objConexao.Execute("SELECT data_pagamento, autorizacao_diretoria_para_finalizar FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE cod_fechamento_periodo = "&cod_fechamento_periodo)
								
		If not rs2.Eof Then
			vendedor_recebe_em = formatDateTime(rs2(0))
			autorizacao_diretoria_para_finalizar = Cstr(""&rs2("autorizacao_diretoria_para_finalizar"))
		End If
								
	rs2.Close
	Set rs2=Nothing
	FecharConexao()	
	
Dim VERIFICA_ERRO
VERIFICA_ERRO = 2
%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Verifica&ccedil;&otilde;es do fechamento</title>
<style>
:root { --ink:#242738; --muted:#575a70; --line:#e4dfeb; --accent:#b20b61; }
* { box-sizing:border-box; }
body { margin:0; background:#f6f7fb; color:var(--ink); font:15px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif; }
main { width:min(1360px,100%); margin:auto; padding:clamp(12px,3vw,30px); }
.check-header,.check-card { background:#fff; border:1px solid var(--line); border-radius:12px; padding:clamp(14px,2vw,22px); margin-bottom:14px; }
.check-header { border-left:4px solid var(--accent); }
.check-header h1 { font-size:clamp(22px,4vw,30px); line-height:1.2; margin:0 0 5px; }
.check-header p { margin:0; color:var(--muted); }
.check-card h2 { display:flex; align-items:center; gap:10px; margin:0 0 10px; font-size:clamp(16px,2vw,19px); }
.check-number,.check-badge { flex:none; display:inline-grid; place-items:center; width:30px; height:30px; border-radius:9px; background:#f9eaf2; color:var(--accent); font-weight:750; }
.check-badge.ok { background:#e9f7ee; color:#21633c; }
.check-badge.attention { background:#fff0dd; color:#864909; }
.check-card table { width:100%; margin:0; border-collapse:collapse; border:0; float:none; }
.check-card table td { padding:8px 10px; border:0; }
.check-card table[border="0"] td:last-child { width:44px; }
.check-card > div[style*="border-bottom"] { border:0 !important; }
.check-card > br { display:none; }
.check-scroll { max-width:100%; overflow:auto; border:1px solid var(--line); border-radius:8px; margin:10px 0; }
.check-scroll table { min-width:510px; }
.check-scroll table tr:first-child td { background:#f0edf5; color:#433b51; font-size:13px; }
.check-scroll table tr:not(:first-child) td { border-top:1px solid #edeaf3; }
.check-scroll table tr:nth-child(even) td { background:#faf9fc; }
.check-final .check-scroll table { min-width:0; }
a { color:var(--accent); }
:focus-visible { outline:2px solid var(--accent); outline-offset:3px; }
.msgGrandeErro,.msgGrandeSucesso { padding:10px 12px; border-radius:8px; }
.msgGrandeErro { background:#fff1f3; color:#8d2441; }
.msgGrandeSucesso { background:#e9f7ee; color:#21633c; }
.process-button { border:0; border-radius:8px; background:var(--accent); color:#fff; font:inherit; font-weight:700; padding:10px 18px; cursor:pointer; }
.process-button:disabled { opacity:.65; cursor:wait; }
#statusProcessamento:not(:empty) { margin-top:10px; }
#statusProcessamento[aria-busy="true"]::after { content:""; display:inline-block; width:1em; height:1em; margin-left:.55em; vertical-align:-.12em; border:2px solid #e4dfeb; border-top-color:var(--accent); border-radius:50%; animation:check-spin .75s linear infinite; }
@keyframes check-spin { to { transform:rotate(360deg); } }
@media(prefers-reduced-motion:reduce) { #statusProcessamento::after { animation:none !important; } }
[hidden] { display:none !important; }
.check-metrics { display:grid; grid-template-columns:repeat(auto-fit,minmax(min(100%,180px),1fr)); gap:10px; margin:12px 0 18px; }
.check-metrics div { display:grid; align-content:start; gap:4px; border:1px solid var(--line); background:#faf8fc; border-radius:8px; padding:12px; }
.check-metrics span,.check-hint { color:var(--muted); font-size:13px; }
.check-metrics strong { font-size:24px; line-height:1.2; }
.check-metrics meter { width:100%; height:12px; accent-color:var(--accent); }
.check-overview h3 { font-size:15px; margin:0 0 3px; }
.check-hint { margin:0 0 10px; }
.check-chart { display:grid; gap:7px; margin:12px 0 18px; }
.check-bar-row { display:grid; grid-template-columns:minmax(0,260px) minmax(0,1fr) minmax(42px,auto); gap:8px; align-items:center; font-size:13px; }
.check-bar-row strong { text-align:right; }
.check-track { height:8px; border-radius:8px; background:#f0edf5; overflow:hidden; }
.check-track span { display:block; height:100%; background:var(--accent); border-radius:8px; }
.check-filters { display:flex; gap:12px; flex-wrap:wrap; margin-bottom:9px; }
.check-filters label { display:grid; gap:4px; flex:1 1 200px; font-weight:650; }
.check-filters input,.check-filters select,.check-pager button { font:inherit; padding:7px 9px; border:1px solid var(--line); border-radius:7px; background:#fff; color:var(--ink); min-width:0; }
.check-pager { display:flex; justify-content:flex-end; align-items:center; gap:9px; flex-wrap:wrap; font-size:13px; }
.check-pager button:disabled { opacity:.55; }
@media(max-width:550px) { .check-bar-row { grid-template-columns:minmax(0,1.4fr) minmax(0,1fr) 38px; } }
.check-decision { border:1px solid #c6e6d1; background:#edf8f1; border-left:4px solid #27804d; border-radius:9px; padding:12px 14px; margin:8px 0 15px; }
.check-decision.is-blocked { border-color:#f0c9d1; border-left-color:#ab294a; background:#fff1f4; }
.check-decision h3,.check-decision p { margin:0; }
.check-decision p { margin-top:4px; }
.check-overview-grid { display:grid; grid-template-columns:minmax(220px,.7fr) minmax(0,1.3fr); gap:12px; margin:14px 0; }
.check-distribution,.check-next { margin:0; padding:15px; border:1px solid var(--line); border-radius:10px; }
.check-distribution figcaption { font-weight:750; text-align:center; }
.check-donut { width:120px; height:120px; border-radius:50%; margin:12px auto; display:grid; place-items:center; background:conic-gradient(#ab294a 0 var(--blocked-end),#a96c1a var(--blocked-end) var(--review-end),#27804d var(--review-end) 100%); }
.check-donut > span { background:#fff; width:91px; height:91px; border-radius:50%; display:grid; align-content:center; justify-items:center; }
.check-donut strong { font-size:28px; line-height:1; }
.check-donut small { color:var(--muted); }
.check-legend { padding:0; margin:0; display:flex; flex-wrap:wrap; justify-content:center; gap:4px 11px; list-style:none; font-size:13px; }
.check-legend li { display:flex; gap:5px; align-items:center; }
.check-dot { width:9px; height:9px; border-radius:50%; }
.check-dot.state-blocked { background:#ab294a; }
.check-dot.state-review { background:#a96c1a; }
.check-dot.state-clear { background:#27804d; }
.check-state { display:inline-block; flex:none; border-radius:16px; padding:3px 8px; font-size:11px; font-weight:750; }
.check-state.state-blocked { color:#8d2441; background:#fff1f4; }
.check-state.state-review { color:#794b11; background:#fff0dd; }
.check-state.state-clear { color:#21633c; background:#e9f7ee; }
.check-card h2 { flex-wrap:wrap; }
.check-next ol { padding-left:20px; margin:8px 0 12px; }
.check-next li + li { margin-top:6px; }
@media(max-width:650px) { .check-overview-grid { grid-template-columns:minmax(0,1fr); } }
</style>
</head>
<body>
<main>
<header class="check-header">
<h1>Verifica&ccedil;&otilde;es do fechamento</h1>
<p>Per&iacute;odo #<%=Server.HTMLEncode(cod_fechamento_periodo & "")%> &middot; Pagamento em <%=Server.HTMLEncode(vendedor_recebe_em & "")%></p>
</header>
<section class="check-card check-overview" id="check-overview" aria-labelledby="check-overview-title" hidden>
  <h2 id="check-overview-title">Resumo das verifica&ccedil;&otilde;es</h2>
  <div id="check-decision" class="check-decision" role="status">
    <h3 id="check-decision-title"></h3><p id="check-decision-text"></p>
  </div>
  <div class="check-overview-grid">
    <figure class="check-distribution">
      <figcaption>Situa&ccedil;&atilde;o das 7 verifica&ccedil;&otilde;es</figcaption>
      <div class="check-donut" id="check-donut" aria-hidden="true"><span><strong id="check-donut-blocked"></strong><small>bloqueios</small></span></div>
      <ul class="check-legend" id="check-legend"></ul>
    </figure>
    <div class="check-next">
      <h3>Como finalizar o fechamento</h3>
      <ol id="check-guidance"></ol>
      <a href="#check-final">Ir para o processamento &darr;</a>
    </div>
  </div>
  <div class="check-metrics">
    <div><span>Se&ccedil;&otilde;es com registros ou avisos</span><strong id="check-metric">0 de 7</strong><meter id="check-meter" min="0" max="7" value="0" aria-label="Seções com registros ou avisos"></meter></div>
    <div><span>Pedidos ainda n&atilde;o processados</span><strong id="check-pending-count"></strong></div>
    <div><span>Total de pedidos corrigidos</span><strong id="check-total-count"></strong></div>
  </div>
  <h3>Registros por verifica&ccedil;&atilde;o</h3>
  <p class="check-hint">As barras comparam a quantidade de registros de cada lista. Um pedido pode aparecer em mais de uma verifica&ccedil;&atilde;o.</p>
  <div class="check-chart" id="check-chart"></div>
  <div class="check-filters">
    <label>Buscar nas tabelas<input id="check-query" type="search" placeholder="Pedido, produto ou vendedor"></label>
    <label>Verifica&ccedil;&atilde;o<select id="check-section"><option value="">Todas</option><option value="attention">Com registros ou avisos</option></select></label>
  </div>
  <p id="check-filter-status" class="check-hint" role="status"></p>
</section>

<%
AbrirConexao()
%>
<section class="check-card" aria-labelledby="check-1"><h2 id="check-1"><span class="check-number">1</span>Pedidos anteriores ainda n&atilde;o processados</h2>

<%
sql = "SELECT COD_PEDIDO, DATA_VENDA, DATA_CADASTRO, VENDEDOR_RECEBE_EM FROM TB_PEDIDOS WHERE VENDEDOR_RECEBE_EM < "&sqlData(VENDEDOR_RECEBE_EM)&" AND FINALIZA_CORRECAO = 1 AND PAGAMENTO_PROCESSADO = 2"
'Response.Write(sql)
Set rs=objConexao.Execute(sql)

If not rs.Eof Then

VERIFICA_ERRO = 1
%>
<div class="check-scroll" role="region" aria-label="Tabela de verifica&ccedil;&atilde;o" tabindex="0"><table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
        <td><div align="center"><strong>C&oacute;d. Pedido</strong></div></td>
        <td><div align="center"><strong>Data Venda</strong></div></td>
        <td><div align="center"><strong>Data Cadastro</strong></div></td>
        <td><div align="center"><strong>Vendedor Recebe em:</strong></div></td>
  </tr>
    <%
    While not rs.Eof
    %> 
      <tr>
        <td><div align="center"><a href="lca_busca_detalhamento_pedido.asp?cod_pedido=<%=rs("COD_PEDIDO")%>" target="_blank" rel="noopener"><%=rs("COD_PEDIDO")%></a></div></td>
        <td><div align="center"><%=rs("DATA_VENDA")%></div></td>
        <td><div align="center"><%=rs("DATA_CADASTRO")%></div></td>
        <td><div align="center"><%=rs("VENDEDOR_RECEBE_EM")%></div></td>
      </tr>
    <%
    rs.MoveNext()
    Wend
    %>
</table></div>
<br />
<div style="border-bottom:solid 1px #CCCCCC">
<table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>H&aacute; pedidos n&atilde;o processados programados para um recebimento anterior a este pagamento. Corrija antes de continuar</strong></td>
    <td width="8%"><span class="check-badge attention" aria-label="Requer aten&ccedil;&atilde;o">!</span></td>
  </tr>
</table>
</div>
<%
Else
%>
<div style="border-bottom:solid 1px #CCCCCC">
<table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>Todos os pedidos anteriores a este fechamento j&aacute; foram processados.</strong></td>
    <td width="8%"><span class="check-badge ok" aria-label="Sem pend&ecirc;ncia">&#10003;</span></td>
  </tr>
</table>
</div>
<%
End If

rs.Close
Set rs=Nothing
%>
<br />
<br />

</section>
<section class="check-card" aria-labelledby="check-2"><h2 id="check-2"><span class="check-number">2</span>Dados obrigat&oacute;rios dos pedidos</h2>

<%
sql = "SELECT COD_PEDIDO, DATA_VENDA, DATA_CADASTRO, VENDEDOR_RECEBE_EM FROM TB_PEDIDOS WHERE VENDEDOR_RECEBE_EM = "&sqlData(VENDEDOR_RECEBE_EM)&" AND FINALIZA_CORRECAO = 1 AND (TIPO_VENDA = 0 or cod_fechamento_periodo = 0)"
'Response.Write(sql)
Set rs=objConexao.Execute(sql)

If not rs.Eof Then

VERIFICA_ERRO = 1
%>
<div class="check-scroll" role="region" aria-label="Tabela de verifica&ccedil;&atilde;o" tabindex="0"><table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
        <td><div align="center"><strong>C&oacute;d. Pedido</strong></div></td>
        <td><div align="center"><strong>Data Venda</strong></div></td>
        <td><div align="center"><strong>Data Cadastro</strong></div></td>
        <td><div align="center"><strong>Vendedor Recebe em:</strong></div></td>
  </tr>
    <%
    While not rs.Eof
    %> 
      <tr>
        <td><div align="center"><%=rs("COD_PEDIDO")%></div></td>
        <td><div align="center"><%=rs("DATA_VENDA")%></div></td>
        <td><div align="center"><%=rs("DATA_CADASTRO")%></div></td>
        <td><div align="center"><%=rs("VENDEDOR_RECEBE_EM")%></div></td>
      </tr>
    <%
    rs.MoveNext()
    Wend
    %>
</table></div>
    <br />
    <div style="border-bottom:solid 1px #CCCCCC">
    <table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>Informe o Administrador do Sistema a Falha</strong></td>
    <td width="8%"><span class="check-badge attention" aria-label="Requer aten&ccedil;&atilde;o">!</span></td>
  </tr>
</table>
</div>
<%
Else
%>
<div style="border-bottom:solid 1px #CCCCCC">
<table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>N&atilde;o h&aacute; pedidos com estes campos em branco</strong></td>
    <td width="8%"><span class="check-badge ok" aria-label="Sem pend&ecirc;ncia">&#10003;</span></td>
  </tr>
</table>
</div>
<%
End If

rs.Close
Set rs=Nothing
%>
<br />
<br />
</section>
<section class="check-card" aria-labelledby="check-3"><h2 id="check-3"><span class="check-number">3</span>Pedidos futuros j&aacute; processados</h2>
<%
sql = "SELECT COD_PEDIDO, DATA_VENDA, DATA_CADASTRO, VENDEDOR_RECEBE_EM FROM TB_PEDIDOS WHERE VENDEDOR_RECEBE_EM >= "&sqlData(VENDEDOR_RECEBE_EM)&" AND PAGAMENTO_PROCESSADO = 1"
'Response.Write(sql)
Set rs=objConexao.Execute(sql)

If not rs.Eof Then

VERIFICA_ERRO = 1
%>
<div class="check-scroll" role="region" aria-label="Tabela de verifica&ccedil;&atilde;o" tabindex="0"><table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
        <td width="15%"><div align="center"><strong>C&oacute;d. Pedido</strong></div></td>
        <td width="18%"><div align="center"><strong>Data Venda</strong></div></td>
        <td width="21%"><div align="center"><strong>Data Cadastro</strong></div></td>
        <td width="21%"><div align="center"><strong>Data Processamento</strong></div></td>
        <td width="25%"><div align="center"><strong>Vendedor Recebe em:</strong></div></td>
      </tr>
    <%
    While not rs.Eof
    %> 
      <tr>
        <td><div align="center"><%=rs("COD_PEDIDO")%></div></td>
        <td><div align="center"><%=rs("DATA_VENDA")%></div></td>
        <td><div align="center"><%=rs("DATA_CADASTRO")%></div></td>
        <td>
			<div align="center">
			  <%
			Set rs2=objCOnexao.Execute("SELECT DATA_PROCESSAMENTO FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_PEDIDO = "&rs("COD_PEDIDO"))
					
					If not rs2.Eof Then
							Response.Write(rs2("DATA_PROCESSAMENTO"))
					Else
							Response.Write("N&Atilde;O PROCESSADO DE FATO")
							
					End If
			
			rs2.Close
			Set rs2=Nothing
			%>
			  
		  </div></td>
        <td><div align="center"><%=rs("VENDEDOR_RECEBE_EM")%></div></td>
      </tr>
    <%
    rs.MoveNext()
    Wend
    %>
</table></div>
    <br />
    <div style="border-bottom:solid 1px #CCCCCC">
    <table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>H&aacute; pedidos j&aacute; processados para um pagamento futuro a este, n&atilde;o &eacute; poss&iacute;vel continuar</strong></td>
    <td width="8%"><span class="check-badge attention" aria-label="Requer aten&ccedil;&atilde;o">!</span></td>
  </tr>
</table>
</div>
<%
Else
%>
<div style="border-bottom:solid 1px #CCCCCC">
    <table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>N&atilde;o h&aacute; pedidos futuros j&aacute; processados</strong></td>
    <td width="8%"><span class="check-badge ok" aria-label="Sem pend&ecirc;ncia">&#10003;</span></td>
  </tr>
</table>
</div>
<%
End If

rs.Close
Set rs=Nothing
%><br />

</section>
<section class="check-card" aria-labelledby="check-4"><h2 id="check-4"><span class="check-number">4</span>Comiss&otilde;es cadastradas por produto</h2>
<%
Dim qtd_tabelas, VERIFICA_ERRO_tipo4, temp
VERIFICA_ERRO_tipo4 = 2
qtd_tabelas = 0
Set rs=objConexao.Execute("SELECT COUNT(*) as qtd FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES")

If not rs.Eof Then
	qtd_tabelas = formatNumber(0&rs(0),0)/1
End If

rs.Close
Set rs=Nothing

qtd_tabelas = qtd_tabelas/1

sql = "SELECT TB_PRODUTOS.COD_PRODUTO, TB_PRODUTOS.NOME_PRODUTO, COUNT(TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES_PRODUTOS.cod_tabela_comissao_vendedores) AS QTD_TABELAS FROM TB_PRODUTOS LEFT JOIN TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES_PRODUTOS ON TB_PRODUTOS.COD_PRODUTO = TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES_PRODUTOS.COD_PRODUTO WHERE TB_PRODUTOS.PAGAMENTO_COMISSAO = 'S' AND  TB_PRODUTOS.COD_PRODUTO IN (SELECT  cod_produto FROM TB_PEDIDOS_PRODUTOS INNER JOIN TB_PEDIDOS ON TB_PEDIDOS_PRODUTOS.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO WHERE   TB_PEDIDOS.vendedor_recebe_em = "&sqlData(VENDEDOR_RECEBE_EM)&") GROUP BY TB_PRODUTOS.COD_PRODUTO"

Set rs=objConexao.Execute(sql)

If not rs.Eof Then
%>

<div class="check-scroll" role="region" aria-label="Tabela de verifica&ccedil;&atilde;o" tabindex="0"><table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
        <td width="27%"><div align="center"><strong>C&oacute;d. Produto</strong></div></td>
        <td width="30%"><div align="center"><strong>Produto</strong></div></td>
        <td width="43%"><div align="center"><strong>N&atilde;o Cadastrado na(s) tabela(s)</strong></div></td>
  </tr>
    <%
    While not rs.Eof
		
		temp = formatNumber(0&rs("QTD_TABELAS"),0)
		temp = temp/1
		
		If temp < qtd_tabelas Then
		VERIFICA_ERRO = 1
		VERIFICA_ERRO_tipo4 = 1
    %> 
      <tr>
        <td><div align="center"><%=rs("COD_PRODUTO")%></div></td>
        <td><div align="center"><%=rs("NOME_PRODUTO")%></div></td>
        <td><div align="center">PRODUTO N&Atilde;O CADASTRADO EM TODAS AS TABELAS</div></td>
      </tr>
    <%
		End If
	
    rs.MoveNext()
    Wend
    %>
</table></div>
    <br />
    
<%
End If

rs.Close
Set rs=Nothing
%>
<%If VERIFICA_ERRO_tipo4 = 1 Then%>
<div style="border-bottom:solid 1px #CCCCCC">
    <table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>H&aacute; produtos que n&atilde;o est&atilde;o cadastrados em todas as tabelas de comissionamento de vendedores.</strong></td>
    <td width="8%"><span class="check-badge attention" aria-label="Requer aten&ccedil;&atilde;o">!</span></td>
  </tr>
</table>
</div>
<%Else%>
<div style="border-bottom:solid 1px #CCCCCC">
    <table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>Todos os produtos est&atilde;o cadastrados em todas as tabelas de pre&ccedil;os para este pagamento.</strong></td>
    <td width="8%"><span class="check-badge ok" aria-label="Sem pend&ecirc;ncia">&#10003;</span></td>
  </tr>
</table>
</div>
<%End If%>
<br />
<br />
</section>
<section class="check-card" aria-labelledby="check-5"><h2 id="check-5"><span class="check-number">5</span>Confer&ecirc;ncia da Diretoria</h2>
<%
sql = "SELECT * FROM TB_PEDIDOS WHERE FINALIZA_CORRECAO = 1 AND FINALIZA_DIRETORIA = 2 AND VENDEDOR_RECEBE_EM = "&sqlData(VENDEDOR_RECEBE_EM)
'Response.Write(sql)
Set rs=objConexao.Execute(sql)

If not rs.Eof Then
	
	If autorizacao_diretoria_para_finalizar = "N" Then
		VERIFICA_ERRO = 1
	End If
%>
<div class="check-scroll" role="region" aria-label="Tabela de verifica&ccedil;&atilde;o" tabindex="0"><table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
        <td width="36%"><div align="center"><strong>C&oacute;d. Pedido</strong></div></td>
        <td width="28%"><div align="center"><strong>Data Venda</strong></div></td>
        <td width="36%"><div align="center"><strong>Data desceu p/ Diretoria</strong></div></td>
  </tr>
    <%
    While not rs.Eof
    %> 
      <tr>
        <td><div align="center"><a href="lca_busca_detalhamento_pedido.asp?cod_pedido=<%=rs("COD_PEDIDO")%>" target="_blank" rel="noopener"><%=rs("COD_PEDIDO")%></a></div></td>
        <td><div align="center"><%=rs("DATA_VENDA")%></div></td>
        <td><div align="center"><%=rs("DATA_DESCEU_DIRETORIA")%></div></td>
      </tr>
    <%
    rs.MoveNext()
    Wend
    %>
</table></div>
<br />
    <div style="border-bottom:solid 1px #CCCCCC">
    <table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>H&aacute; pedidos que n&atilde;o foram conferidos pela diretoria para este pagamento.</strong>
    	
    </td>
    <td width="8%"><span class="check-badge attention" aria-label="Requer aten&ccedil;&atilde;o">!</span></td>
  </tr>
</table>

	<%
		If autorizacao_diretoria_para_finalizar = "N" Then
		%>
    		<div class="msgGrandeErro">A diretoria N&Atilde;O aprovou continuar este pagamento sem finalizar a corre&ccedil;&atilde;o de todos os pedidos</div>
    	<%
		Else
		%>
        	<div class="msgGrandeSucesso">A diretoria aprovou continuar este pagamento sem finalizar a corre&ccedil;&atilde;o de todos os pedidos</div>
        <%End If%>

</div>
<%
Else
%>
<div style="border-bottom:solid 1px #CCCCCC">
    <table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>Todos os pedidos para este fechamento foram conferidos pela Diretoria</strong></td>
    <td width="8%"><span class="check-badge ok" aria-label="Sem pend&ecirc;ncia">&#10003;</span></td>
  </tr>
</table>
</div>
<%
End If

rs.Close
Set rs=Nothing
%><br />

</section>
<section class="check-card" aria-labelledby="check-6"><h2 id="check-6"><span class="check-number">6</span>Pedidos programados para pagamentos futuros</h2>

<%
sql = "SELECT COD_PEDIDO, DATA_VENDA, DATA_CADASTRO, COD_VENDEDOR, VENDEDOR_RECEBE_EM FROM TB_PEDIDOS WHERE VENDEDOR_RECEBE_EM > "&sqlData(VENDEDOR_RECEBE_EM)&" AND FINALIZA_CORRECAO = 1 AND DATA_VENDA < "&sqlData(VENDEDOR_RECEBE_EM)&" ORDER BY VENDEDOR_RECEBE_EM DESC"
'Response.Write(sql)
Set rs=objConexao.Execute(sql)


If not rs.Eof Then
%>
<div class="check-scroll" role="region" aria-label="Tabela de verifica&ccedil;&atilde;o" tabindex="0"><table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
        <td width="17%"><div align="center"><strong>C&oacute;d. Pedido</strong></div></td>
        <td width="17%"><div align="center"><strong>Nome Vendedor</strong></div></td>
        <td width="14%"><div align="center"><strong>Data Venda</strong></div></td>
        <td width="21%"><div align="center"><strong>Data Cadastro</strong></div></td>
        <td width="31%"><div align="center"><strong>Vendedor Recebe em:</strong></div></td>
      </tr>
    <%
    While not rs.Eof
    %> 
      <tr>
        <td><div align="center"><a href="lca_busca_detalhamento_pedido.asp?cod_pedido=<%=rs("COD_PEDIDO")%>" target="_blank" rel="noopener"><%=rs("COD_PEDIDO")%></a></div></td>
        <td>
			<div align="left">
			  <%
			Set rs2=objConexao.Execute("SELECT NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR = "&rs("COD_VENDEDOR"))
				
				If not rs2.Eof Then	
					Response.Write(rs2("NOME_VENDEDOR"))		
				End If
				
			rs2.Close
			Set rs2=Nothing
			%>
			  
	        </div></td>
        <td><div align="center"><%=rs("DATA_VENDA")%></div></td>
        <td><div align="center"><%=rs("DATA_CADASTRO")%></div></td>
        <td><div align="center"><%=rs("VENDEDOR_RECEBE_EM")%></div></td>
      </tr>
    <%
    rs.MoveNext()
    Wend
    %>
</table></div>
<%
Else
%>
<div style="border-bottom:solid 1px #CCCCCC">
<table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>N&atilde;o h&aacute; pedidos programados para pagamentos futuros</strong></td>
    <td width="8%"><span class="check-badge ok" aria-label="Sem pend&ecirc;ncia">&#10003;</span></td>
  </tr>
</table>
</div>
<%
End If

rs.Close
Set rs=Nothing
%>
<br />
</section>
<section class="check-card" aria-labelledby="check-7"><h2 id="check-7"><span class="check-number">7</span>Pend&ecirc;ncias deste fechamento</h2>
<%
sql = "SELECT TB_FECHAMENTO_PENDENCIAS_PAGAMENTO.*, TB_VENDEDORES.NOME_VENDEDOR, TB_ADM_USUARIOS.NOME FROM TB_VENDEDORES INNER JOIN TB_FECHAMENTO_PENDENCIAS_PAGAMENTO ON TB_VENDEDORES.COD_VENDEDOR = TB_FECHAMENTO_PENDENCIAS_PAGAMENTO.COD_VENDEDOR INNER JOIN TB_ADM_USUARIOS ON TB_FECHAMENTO_PENDENCIAS_PAGAMENTO.CADASTRADO_POR = TB_ADM_USUARIOS.COD_USUARIO WHERE TB_FECHAMENTO_PENDENCIAS_PAGAMENTO.cod_fechamento_periodo = "&cod_fechamento_periodo&" ORDER BY VENDEDOR_RECEBE_EM DESC"
'Response.Write(sql)
Set rs=objConexao.Execute(sql)


If not rs.Eof Then
%>
<div class="check-scroll" role="region" aria-label="Tabela de verifica&ccedil;&atilde;o" tabindex="0"><table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="27%"><div align="center"><strong>Vendedor</strong></div></td>
            <td width="10%"><div align="center"><strong>Tipo</strong></div></td>
            <td width="12%"><div align="center"><strong>Valor</strong></div></td>
            <td width="33%"><div align="center"><strong>Motivo</strong></div></td>
            <td width="18%"><div align="center"><strong>Data do Cadastro</strong></div>
     </td>
  </tr>

		  		 <%
	 
		While not rs.Eof
			
				%>
				<tr>
							<td><div align="left"><%=rs("nome_vendedor")%></div></td>
							<td align="center">
<%
							If Ucase(Cstr(rs("TIPO"))) = "A" Then
								Response.Write("Acr�scimo (+)")
							Else
								Response.Write("Decr�scimo (-)")
							End If
							%>                            </td>
							<td><div align="right"><%=FormatCurrency(rs("VALOR"))%></div></td>
							<td><div align="left"><%=rs("MOTIVO")%></div></td>
							<td><div align="center"><%=ArrumaData(rs("DATA_CADASTRO"))%> / <%=rs("NOME")%> </div></td>
				</tr>
		        <%
			rs.MoveNext()
			Wend	
      
		  %>
</table></div>
<%Else%>
<table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>N&atilde;o h&aacute; pend&ecirc;ncias para este fechamento</strong></td>
    <td width="8%"><span class="check-badge ok" aria-label="Sem pend&ecirc;ncia">&#10003;</span></td>
  </tr>
</table>
<%
End If

rs.Close
Set rs=Nothing
%>
<br /><br />
<!--
<div style="border-bottom:solid 1px #CCCCCC; background-color:#FFFFCC">

<form name="formulario_padrao" action="adm_rotina_fechamento_extrato_pagamento_part2_tipo2.asp" style="margin:0px">
<input type="hidden" name="VENDEDOR_RECEBE_EM" value="<%=VENDEDOR_RECEBE_EM%>" />

'FORMA DE PAGAMENTO NOVA N�O � MAIS PERMITIDO ALTERA��O DO DIA DE PAGAMENTO
<table width="600" border="0" cellpadding="5" cellspacing="0">
  <tr>
    <td width="52%" class="textoMedio"><strong>Escolha o dia do pagamento dos recibos</strong></td>
    <td width="48%">
  		<select name="DIA_VENCIMENTO" class="inputTamanho1">
			<option value="10">10</option>
        	<option value="11">11</option>
            <option value="12">12</option>
        	<option value="13">13</option>
            <option value="14">14</option>
        	<option value="15" selected="selected">15</option>
            <option value="16">16</option>
            <option value="17">17</option>
            <option value="18">18</option>
            <option value="19">19</option>
            <option value="20">20</option>
        </select>
    </td>
  </tr>
</table>

</form>
</div>
-->
<br />
<br />
</section>
<section class="check-card check-final" aria-labelledby="check-final"><h2 id="check-final">Processamento</h2>

<div class="check-scroll" role="region" aria-label="Tabela de verifica&ccedil;&atilde;o" tabindex="0"><table width="100%"  border="1" align="center" cellpadding="5" style="border-collapse:collapse">	
                    
					
                    <tr>
               			<td colspan="15"><div align="center" style="border-bottom:solid 2PX #CCCCCC; margin-bottom:15PX;"><strong>QTD. TOTAL DE PEDIDOS N&Atilde;O PROCESSADOS:</strong> <span class="style2">
               			 <%
							Set rs2=objConexao.Execute("SELECT COUNT(*) AS QTD_PEDIDOS FROM TB_PEDIDOS WHERE FINALIZA_CORRECAO = 1 AND PAGAMENTO_PROCESSADO = 2 AND VENDEDOR_RECEBE_EM = "&sqlData(VENDEDOR_RECEBE_EM))
				
								If not rs2.Eof Then
									Response.Write(rs2("QTD_PEDIDOS"))
									temp = formatNumber(0&rs2(0),0)
								End If
				
							rs2.Close
							Set rs2=Nothing
						%>
                	</span></div>
                  <div align="center"><strong>QTD. TOTAL DE PEDIDOS TOTAL:</strong> <span class="style2">
                 	 <%
						Set rs2=objConexao.Execute("SELECT COUNT(*) AS QTD_PEDIDOS FROM TB_PEDIDOS WHERE FINALIZA_CORRECAO = 1 AND VENDEDOR_RECEBE_EM = "&sqlData(VENDEDOR_RECEBE_EM))
						
							If not rs2.Eof Then
								Response.Write(rs2("QTD_PEDIDOS"))
							
							End If
						
						rs2.Close
						Set rs2=Nothing
						%>
                </span></div></td>
  </tr>
                 
                 	<tr>
    <td height="31" colspan="15" align="center">
        <% If int(VERIFICA_ERRO) = 2 and int(temp) > 0 Then %>

            <button type="button" id="btnProcessar" class="process-button" data-periodo="<%=Server.HTMLEncode(cod_fechamento_periodo & "")%>" data-pagamento="<%=Server.HTMLEncode(vendedor_recebe_em & "")%>">
                Processar Pagamento
            </button>

            <div id="statusProcessamento" role="status"></div>

        <% Else %>
            <div class="check-decision is-blocked">
          <%If int(VERIFICA_ERRO) <> 2 Then%>
          <strong>Pagamento bloqueado.</strong> Resolva as verifica&ccedil;&otilde;es sinalizadas como bloqueio e reabra este fechamento.
          <%Else%>
          <strong>Sem pedidos para processar.</strong> Confira os totais e o per&iacute;odo selecionado.
          <%End If%>
        </div>
        <% End If %>
    </td>
</tr>
			</table></div>
<br />
<br />     
          
</section>
<%
FecharConexao()
%>
<script>
document.addEventListener('DOMContentLoaded', function () {
    const button = document.getElementById('btnProcessar');
    if (!button) return;
    const status = document.getElementById('statusProcessamento');
    const pedidos = new Set();
    const url = new URL('adm_rotina_fechamento_extrato_pagamento_part2_tipo2.asp', document.baseURI);
    url.searchParams.set('COD_FECHAMENTO_PERIODO', button.dataset.periodo);
    url.searchParams.set('VENDEDOR_RECEBE_EM', button.dataset.pagamento);
    url.searchParams.set('modo', 'modal');

    button.addEventListener('click', async function () {
        button.disabled = true;
        status.setAttribute('aria-busy', 'true');
        status.textContent = 'Processando… Mantenha este modal aberto.';
        try {
            while (true) {
                const response = await fetch(url, {cache: 'no-store', redirect: 'error'});
                if (!response.ok) throw new Error('HTTP ' + response.status);
                const result = new DOMParser().parseFromString(await response.text(), 'text/html');
                const outcome = result.querySelector('[data-fechamento-resultado]');
                if (!outcome) throw new Error('O servidor não confirmou o resultado da etapa.');
                const state = outcome.dataset.fechamentoResultado;
                if (state === 'concluido') {
                    status.textContent = 'Fechamento concluído. Pedidos processados nesta execução: ' + pedidos.size + '.';
                    break;
                }
                const pedido = outcome.dataset.pedido;
                if (state !== 'pedido' || !/^\d+$/.test(pedido) || pedidos.has(pedido)) {
                    throw new Error('O servidor não confirmou avanço no processamento.');
                }
                pedidos.add(pedido);
                status.textContent = pedidos.size + ' pedido(s) processado(s) nesta execução. Mantenha o modal aberto.';
            }
        } catch (error) {
            status.textContent = 'Processamento interrompido. Pedidos confirmados nesta execução: ' + pedidos.size +
                '. Feche e reabra as verificações para conferir o estado antes de retomar. ' + error.message;
        } finally {
            status.removeAttribute('aria-busy');
        }
    });
});
</script>
<script>
(function () {
    const overview = document.getElementById('check-overview');
    const query = document.getElementById('check-query');
    const select = document.getElementById('check-section');
    const chart = document.getElementById('check-chart');
    const format = new Intl.NumberFormat('pt-BR');
    const normalize = value => value.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
    const pageSize = 15;
    const sections = Array.from(document.querySelectorAll('section[aria-labelledby^="check-"]'))
        .filter(section => /^check-[1-7]$/.test(section.getAttribute('aria-labelledby')))
        .map(section => {
            const id = section.getAttribute('aria-labelledby');
            const title = document.getElementById(id).lastChild.textContent.trim();
            const tables = Array.from(section.querySelectorAll('.check-scroll table')).map(table => {
                const rows = Array.from(table.rows).slice(1).map(row => ({row, text: normalize(row.textContent)}));
                const controls = document.createElement('div');
                controls.className = 'check-pager';
                const prev = document.createElement('button');
                const next = document.createElement('button');
                const label = document.createElement('span');
                prev.type = next.type = 'button';
                prev.textContent = 'Anterior';
                next.textContent = 'Próxima';
                label.setAttribute('role', 'status');
                controls.append(prev, label, next);
                table.parentElement.after(controls);
                const model = {rows, controls, prev, next, label, page: 0};
                prev.onclick = () => { model.page--; render(); };
                next.onclick = () => { model.page++; render(); };
                return model;
            });
            const count = tables.reduce((sum, table) => sum + table.rows.length, 0);
            return {id, section, title, tables, count, attention: count > 0 || Boolean(section.querySelector('.attention,.msgGrandeErro'))};
        });

    function classify(item) {
        if (!item.attention) return 'clear';
        if (item.id === 'check-5') return item.section.querySelector('.msgGrandeErro') ? 'blocked' : 'review';
        return /^check-[1-4]$/.test(item.id) ? 'blocked' : 'review';
    }

    function reveal(item) {
        query.value = '';
        select.value = item.id;
        reset();
    }

    const labels = {
        blocked: 'Bloqueio', review: 'Requer revisão', clear: 'Sem pendência'
    };
    const counts = {blocked: 0, review: 0, clear: 0};
    sections.forEach(item => {
        item.state = classify(item);
        counts[item.state]++;
        const badge = document.createElement('span');
        badge.className = 'check-state state-' + item.state;
        badge.textContent = labels[item.state];
        document.getElementById(item.id).append(badge);
    });

    const totalChecks = sections.length || 1;
    const donut = document.getElementById('check-donut');
    donut.style.setProperty('--blocked-end', (counts.blocked / totalChecks * 100) + '%');
    donut.style.setProperty('--review-end', ((counts.blocked + counts.review) / totalChecks * 100) + '%');
    document.getElementById('check-donut-blocked').textContent = counts.blocked;
    const legend = document.getElementById('check-legend');
    Object.entries(counts).forEach(([key, count]) => {
        const line = document.createElement('li');
        const dot = document.createElement('span');
        dot.className = 'check-dot state-' + key;
        dot.setAttribute('aria-hidden', 'true');
        line.append(dot, labels[key] + ': ' + count);
        legend.append(line);
    });

    const decision = document.getElementById('check-decision');
    const heading = document.getElementById('check-decision-title');
    const note = document.getElementById('check-decision-text');
    const processAvailable = Boolean(document.getElementById('btnProcessar'));
    if (counts.blocked) {
        decision.classList.add('is-blocked');
        heading.textContent = 'Fechamento bloqueado';
        note.textContent = 'Resolva os bloqueios indicados abaixo e reabra o fechamento para atualizar as verificações.';
    } else if (processAvailable) {
        heading.textContent = counts.review ? 'Processamento liberado com avisos' : 'Pronto para processar';
        note.textContent = counts.review ? 'Revise os avisos antes de clicar em Processar Pagamento.' : 'Confira os totais e clique em Processar Pagamento.';
    } else {
        heading.textContent = 'Sem pedidos para processar';
        note.textContent = 'Confira se este é o período pretendido e se os pedidos foram vinculados a ele.';
    }

    const instructions = {
        'check-1': 'Confira os pedidos de períodos anteriores e regularize o processamento no fechamento correspondente.',
        'check-2': 'Abra os pedidos indicados e revise o tipo de venda e o vínculo com o período de fechamento.',
        'check-3': 'Confira os pedidos já processados com pagamento futuro e regularize o fechamento correspondente.',
        'check-4': 'Cadastre a comissão dos produtos listados nas tabelas de comissionamento aplicáveis.',
        'check-5': 'Encaminhe os pedidos pendentes para conferência da Diretoria. Se for necessário seguir por exceção, solicite a autorização da Diretoria para o período.',
        'check-6': 'Confira se as datas futuras de recebimento dos pedidos estão corretas.',
        'check-7': 'Revise os lançamentos de pendências de pagamento vinculados ao período.'
    };
    const guidance = document.getElementById('check-guidance');
    const byPriority = ['blocked', 'review'].flatMap(state => sections.filter(item => item.state === state));
    byPriority.forEach(item => {
        const step = document.createElement('li');
        const link = document.createElement('a');
        link.href = '#' + item.id;
        link.textContent = item.title;
        link.onclick = () => reveal(item);
        const description = item.id === 'check-5' && item.state === 'review'
            ? 'A conferência segue pendente. A regra atual permite prosseguir neste período. Revise a lista antes de processar.'
            : instructions[item.id];
        step.append(link, ' — ' + description);
        guidance.append(step);
    });
    const last = document.createElement('li');
    last.textContent = counts.blocked
        ? 'Após corrigir, feche o modal e abra Gerar fechamento novamente para conferir se o processamento foi liberado.'
        : processAvailable ? 'Confirme o período e os totais. Em seguida, clique em Processar Pagamento.' : 'Confira o período e os totais na seção Processamento.';
    guidance.append(last);

    const max = Math.max(1, ...sections.map(item => item.count));
    const attention = sections.filter(item => item.attention).length;
    document.getElementById('check-metric').textContent = attention + ' de ' + sections.length;
    document.getElementById('check-meter').value = attention;
    const totals = document.querySelectorAll('.check-final .style2');
    ['check-pending-count','check-total-count'].forEach((id, index) => {
        document.getElementById(id).textContent = totals[index].textContent.trim();
    });
    sections.forEach(item => {
        select.add(new Option(item.title, item.id));
        const line = document.createElement('div');
        line.className = 'check-bar-row';
        const link = document.createElement('a');
        link.href = '#' + item.id;
        link.textContent = item.title;
        link.onclick = () => reveal(item);
        const track = document.createElement('span');
        track.className = 'check-track';
        track.setAttribute('aria-hidden', 'true');
        const bar = document.createElement('span');
        bar.style.width = (item.count / max * 100) + '%';
        track.append(bar);
        const count = document.createElement('strong');
        count.textContent = format.format(item.count);
        line.append(link, track, count);
        chart.append(line);
    });

    function render() {
        const term = normalize(query.value.trim());
        let visible = 0;
        sections.forEach(item => {
            const selected = !select.value || select.value === item.id || (select.value === 'attention' && item.attention);
            let matches = 0;
            item.tables.forEach(table => {
                const filtered = table.rows.filter(entry => !term || entry.text.includes(term));
                matches += filtered.length;
                const pages = Math.max(1, Math.ceil(filtered.length / pageSize));
                table.page = Math.min(Math.max(table.page, 0), pages - 1);
                const begin = table.page * pageSize;
                const shown = new Set(filtered.slice(begin, begin + pageSize));
                table.rows.forEach(entry => { entry.row.hidden = !shown.has(entry); });
                table.controls.hidden = filtered.length <= pageSize;
                table.prev.disabled = table.page === 0;
                table.next.disabled = table.page >= pages - 1;
                table.label.textContent = (begin + 1) + '–' + Math.min(begin + pageSize, filtered.length) + ' de ' + filtered.length;
            });
            item.section.hidden = !selected || (term && !matches);
            if (!item.section.hidden) visible++;
        });
        document.getElementById('check-filter-status').textContent = visible ? visible + ' de ' + sections.length + ' verificações exibidas.' : 'Nenhuma verificação corresponde aos filtros.';
    }

    function reset() {
        sections.forEach(item => item.tables.forEach(table => { table.page = 0; }));
        render();
    }
    query.addEventListener('input', reset);
    select.addEventListener('change', reset);
    overview.hidden = false;
    render();
}());
</script>
<span data-cap-process-ready hidden></span>
</main>
</body>
</html>



