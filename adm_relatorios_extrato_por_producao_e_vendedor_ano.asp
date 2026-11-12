<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Function DataFiltroRetirada(valor)
    Dim partes, padrao, ano, mes, dia, dataValidada
    Set padrao = New RegExp
    padrao.Pattern = "^[0-9]{4}-[0-9]{1,2}-[0-9]{1,2}$"
    If padrao.Test(valor) Then
        partes = Split(valor, "-")
        ano = CInt(partes(0))
        mes = CInt(partes(1))
        dia = CInt(partes(2))
    Else
        padrao.Pattern = "^[0-9]{1,2}/[0-9]{1,2}/[0-9]{4}$"
        If Not padrao.Test(valor) Then Exit Function
        partes = Split(valor, "/")
        ano = CInt(partes(2))
        mes = CInt(partes(1))
        dia = CInt(partes(0))
    End If
    If ano < 1000 Or ano > 9998 Or mes < 1 Or mes > 12 Or dia < 1 Or dia > 31 Then Exit Function
    dataValidada = DateSerial(ano, mes, dia)
    If Month(dataValidada) = mes And Day(dataValidada) = dia Then DataFiltroRetirada = dataValidada
End Function


Function ValorRetirada(consulta, vendedor, campo)
    Dim resultado, valor
    ValorRetirada = 0
    Set resultado = objConexao.Execute(Replace(consulta, "{}", vendedor))
    If Not resultado.EOF Then
        valor = resultado(campo).Value
        If Not IsNull(valor) Then ValorRetirada = CDbl(valor)
    End If
    resultado.Close
End Function

Dim inicio, termino, dataInicio, dataFim, fimExclusivo, relErro, relPeriodo
dataInicio = DataFiltroRetirada(Trim(Request.QueryString("DATA_INICIO")))
dataFim = DataFiltroRetirada(Trim(Request.QueryString("DATA_TERMINO")))
If IsEmpty(dataInicio) Or IsEmpty(dataFim) Then
    relErro = "Informe as datas inicial e final."
ElseIf dataInicio > dataFim Then
    relErro = "A data inicial deve ser anterior ou igual &agrave; data final."
Else
    fimExclusivo = DateAdd("d", 1, dataFim)
    inicio = Year(dataInicio) & "-" & Right("0" & Month(dataInicio),2) & "-" & Right("0" & Day(dataInicio),2)
    termino = Year(fimExclusivo) & "-" & Right("0" & Month(fimExclusivo),2) & "-" & Right("0" & Day(fimExclusivo),2)
    relPeriodo = FormatDateTime(dataInicio,2) & " a " & FormatDateTime(dataFim,2)
End If
%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Retirada de &aacute;lbuns por per&iacute;odo | Portal Stillos</title>
<link rel="stylesheet" href="css/vendor/bootstrap.min.css">
<link rel="stylesheet" href="css/dashboard.css">
<style>
.month-report { max-width:1150px; margin:24px auto; padding:clamp(16px,4vw,36px); background:#fff; border:1px solid #e4dfeb; border-radius:14px; }
.month-tools,.month-header { display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:12px; }
.month-header { margin:20px 0; padding-bottom:14px; border-bottom:2px solid #b20b61; }
.month-header p { margin:0; }
.month-logo { width:150px; padding:10px; border-radius:8px; background:#292b40; }
.month-kpis { display:grid; grid-template-columns:repeat(auto-fit,minmax(min(100%,175px),1fr)); gap:12px; margin:12px 0 16px; }
.month-kpis > div { padding:12px; border:1px solid #e4dfeb; border-radius:10px; }
.month-kpis dd { margin:4px 0 0; font-weight:700; font-size:1.15em; }
.month-scroll { overflow:auto; }
.month-table { width:100%; min-width:520px; border-collapse:collapse; margin-top:10px; font-variant-numeric:tabular-nums; }
.month-table td,.month-table th { padding:8px; text-align:right; border-bottom:1px solid #e4dfeb; }
.month-table th:first-child { text-align:left; }
.month-table thead,.month-table tfoot { background:#f1edf6; }
.month-table tbody tr:nth-child(even) { background:#faf9fc; }
.month-note { margin:14px 0 0; color:#53566b; }
@media screen { .report-embedded .month-report { margin:12px auto; } }
.report-signatures { display:flex; flex-wrap:wrap; gap:24px; margin-top:28px; break-inside:avoid; }
.report-signatures span { flex:1 1 180px; border-top:1px solid #777; padding-top:4px; }
.month-table .report-text { text-align:left; }
.month-table th,.month-table td { overflow-wrap:anywhere; }
[hidden] { display:none !important; }
.month-report { max-width:1480px; }
.ret-filter { display:grid; gap:6px; width:min(100%,420px); margin:0 0 14px; }
.ret-filter select { min-width:0; width:100%; }
.ret-table { min-width:0; }
.ret-label { display:none; }
.ret-products { margin:0; padding:0; list-style:none; }
.ret-products li { display:flex; justify-content:space-between; gap:8px; }
.ret-products li + li { margin-top:3px; }
.ret-meter { display:block; width:100%; min-width:44px; }
.ret-scope:empty { display:none; }
.ret-dashboard { display:grid; grid-template-columns:minmax(0,2fr) minmax(0,1fr); gap:14px; margin:0 0 24px; }
.ret-panel { min-width:0; border:1px solid #e4dfeb; border-radius:12px; padding:16px; }
.ret-panel-heading { display:flex; flex-wrap:wrap; align-items:baseline; justify-content:space-between; gap:4px 12px; margin-bottom:10px; }
.ret-panel .ret-panel-heading h3 { margin:0; }
.ret-help-button { display:inline-flex; gap:6px; align-items:center; padding:4px 8px; border:1px solid #e4dfeb; border-radius:7px; color:#861451; background:#fff; font:inherit; }
.ret-help-button:hover { background:#f6f2f8; }
.ret-help-button span { display:grid; place-items:center; width:18px; height:18px; border:1px solid currentColor; border-radius:50%; font-weight:700; }
.ret-help { width:min(560px,calc(100% - 24px)); max-height:calc(100dvh - 24px); overflow:auto; padding:clamp(16px,4vw,24px); border:1px solid #e4dfeb; border-radius:14px; color:#242738; }
.ret-help::backdrop { background:#24273899; }
.ret-help h2 { margin:0 0 14px; }
.ret-help h3 { margin:16px 0 5px; font-size:1em; }
.ret-help p { margin:0 0 8px; }
.ret-help-example { padding:10px 12px; background:#f6f2f8; border-radius:8px; overflow-wrap:anywhere; }
.ret-help form { display:flex; justify-content:flex-end; margin-top:16px; }
.ret-panel h3 { margin:0 0 10px; font-size:1.05em; }
.ret-legend { display:flex; flex-wrap:wrap; gap:6px 16px; margin:0 0 12px; color:#53566b; }
.ret-legend span::before { content:""; display:inline-block; width:10px; height:10px; margin-right:6px; border-radius:2px; background:#74668f; }
.ret-legend span + span::before { background:#a31962; }
.ret-bars { margin:0; padding:0; list-style:none; }
.ret-bars li + li { margin-top:12px; }
.ret-bars strong { display:block; margin-bottom:4px; overflow-wrap:anywhere; }
.ret-bar-row { display:grid; grid-template-columns:minmax(0,1fr) minmax(4ch,max-content); align-items:center; gap:8px; font-variant-numeric:tabular-nums; }
.ret-bar-row + .ret-bar-row { margin-top:3px; }
.ret-bar-track { height:8px; border-radius:5px; background:#f1edf6; overflow:hidden; }
.ret-bar-track span { display:block; height:100%; background:#74668f; }
.ret-bar-row + .ret-bar-row .ret-bar-track span { background:#a31962; }
.ret-dashboard-rate { display:block; font-size:1.8em; font-weight:700; }
.ret-dashboard meter { display:block; width:100%; height:24px; margin:4px 0; }
.ret-dashboard meter::-webkit-meter-optimum-value { background:#a31962; }
.ret-dashboard meter::-moz-meter-bar { background:#a31962; }
.ret-dashboard-caption { margin:4px 0 0; color:#53566b; }
@media screen and (max-width:760px) {
  .ret-dashboard { grid-template-columns:minmax(0,1fr); }
}
@media screen and (max-width:1100px) {
  .ret-table { display:block; }
  .ret-table thead { position:absolute; width:1px; height:1px; overflow:hidden; clip-path:inset(50%); }
  .ret-table tbody { display:grid; gap:12px; }
  .ret-table tbody tr { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); border:1px solid #e4dfeb; border-radius:10px; padding:4px 10px; }
  .ret-table tbody :is(th,td) { min-width:0; text-align:left; }
  .ret-table tbody th,.ret-wide { grid-column:1 / -1; }
  .ret-label { display:block; color:#53566b; font-size:.86em; font-weight:650; }
  .ret-meter { max-width:140px; }
}
@page { size:A4 landscape; margin:10mm; }
@media print {
  body { margin:0; background:#fff; color:#111; font-size:10pt; line-height:1.3; }
  .month-report { margin:0; padding:2mm 3mm; max-width:none; border:0; border-radius:0; }
  .month-tools,.ret-filter,.ret-meter,.ret-table summary,.ret-dashboard,.ret-help { display:none; }
  .month-header { margin:0 0 5mm; padding-bottom:3mm; border-bottom:1px solid #555; break-after:avoid; }
  .month-header h1 { font-size:16pt; margin:0 0 1mm; }
  .month-logo { width:30mm; padding:0; background:transparent; filter:brightness(0); }
  .month-scroll { overflow:visible; }
  .month-table { min-width:0; width:100%; table-layout:fixed; margin-top:0; }
  .month-table th,.month-table td { padding:1.2mm 1mm; border-bottom:1px solid #ccc; }
  .month-table thead { display:table-header-group; }
  .month-table thead th { border-top:1px solid #777; border-bottom:1px solid #777; }
  .month-table tfoot { display:table-row-group; }
  .month-table tr,.month-kpis,.report-signatures { break-inside:avoid; }
  .month-kpis { display:flex; flex-wrap:wrap; gap:2mm 6mm; margin:0 0 4mm; }
  .month-kpis > div { padding:0; border:0; flex:1 1 auto; }
  .month-kpis dt { font-size:9pt; }
  .month-kpis dd { font-size:10pt; margin:1mm 0 0; }
  .report-signatures { margin:12mm 0 0; gap:10mm; }
  .report-signatures span { flex:1 1 0; font-size:9pt; }

  .ret-table th,.ret-table td { box-sizing:border-box; overflow-wrap:normal; word-break:normal; }
  .ret-table thead th { font-size:9pt; line-height:1.2; }
  .ret-table td.ret-wide { text-align:left; }
  .ret-table details .ret-products { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:.45mm 2mm; font-size:8.5pt; line-height:1.2; }
  .ret-table .ret-products li { display:grid; grid-template-columns:minmax(0,1fr) auto; gap:1mm; margin:0; align-items:start; }
  .ret-table .ret-products strong { white-space:nowrap; }
  .ret-table thead th:nth-child(1) { width:19%; }
  .ret-table thead th:nth-child(2) { width:7%; }
  .ret-table thead th:nth-child(3) { width:7%; }
  .ret-table thead th:nth-child(4) { width:11%; }
  .ret-table thead th:nth-child(5) { width:11%; }
  .ret-table thead th:nth-child(6) { width:23%; }
  .ret-table thead th:nth-child(7) { width:9%; }
  .ret-table thead th:nth-child(8) { width:13%; }
}
</style>
</head>
<body<%If Request.QueryString("modal") = "1" Then%> class="report-embedded"<%End If%>>
<main class="month-report" data-report-preview-content>
  <%If Request.QueryString("modal") <> "1" Then%>
  <nav class="month-tools" aria-label="A&ccedil;&otilde;es do relat&oacute;rio">
    <a class="btn btn-outline-secondary" href="adm_relatorios.asp">Voltar aos relat&oacute;rios</a>
    <%If Len(relErro) = 0 Then%><button type="button" class="btn btn-primary" onclick="window.print()">Imprimir / Salvar PDF</button><%End If%>
  </nav>
<%End If%>
  <header class="month-header">
    <div><h1>Retirada de &aacute;lbuns por per&iacute;odo</h1><p><%=Server.HTMLEncode(relPeriodo & "")%></p></div>
    <img class="month-logo" src="images/logoparafundopreto.png" alt="Stillos">
  </header>
<%If Len(relErro) > 0 Then%>
<p class="alert alert-warning" data-report-error><%=relErro%></p>
<%Else
    Dim sql, sql_1, sql_2, sql_3, sql_4, sql_5, vendedores, produtoRs, produtos, nomesProdutos, i
    Dim codigo, nome, retirados, vendidos, valor, albunsPeriodo, recebido, aproveitamento, quantidadeProduto
    sql = "SELECT NOME_VENDEDOR, TB_VENDEDORES.COD_VENDEDOR, SUM(QTD_FORMANDOS_INICIAL) AS QTD_RETIRADO "&_
    "FROM TB_PRODUCAO_LOTES INNER JOIN TB_VENDEDORES ON TB_VENDEDORES.COD_VENDEDOR = TB_PRODUCAO_LOTES.COD_VENDEDOR_INICIAL "&_
    "WHERE  DATA_SAIDA_INICIAL >= '"&INICIO&"' and  DATA_SAIDA_INICIAL < '"&TERMINO&"'  AND TB_VENDEDORES.COD_VENDEDOR <> 27 GROUP BY TB_VENDEDORES.COD_VENDEDOR, NOME_VENDEDOR ORDER BY NOME_VENDEDOR, TB_VENDEDORES.COD_VENDEDOR"
    sql_1 = "SELECT (SUM(IFNULL(PAGAMENTO_EM_DINHEIRO,0)) + SUM(IFNULL(XTOTAL_CHEQUE,0))+ SUM(IFNULL(XTOTAL_BOLETO,0)) + SUM(IFNULL(XTOTAL_CARTAO,0))) AS VALOR_VENDIDO FROM VW_EXTRATO_TOTAL_RECEBIDO_COMPLETO WHERE DATA_VENDA >= '"&INICIO&"' AND DATA_VENDA < '"&TERMINO&"' AND COD_PEDIDO IN ( SELECT COD_PEDIDO FROM TB_PEDIDOS_FORMANDOS WHERE COD_FORMANDO IN (SELECT COD_FORMANDO FROM TB_PRODUCAO_LOTES_FORMANDOS WHERE DATA_SAIDA >= '"&INICIO&"' AND DATA_SAIDA < '"&TERMINO&"' AND TB_PRODUCAO_LOTES_FORMANDOS.COD_VENDEDOR = COD_VENDEDOR)) AND COD_VENDEDOR = {}"
    sql_2 = "SELECT sum(VW_QTD_FORMANDOS_POR_PEDIDO.QTD_FORMANDOS) AS QTD_VENDIDO FROM VW_QTD_FORMANDOS_POR_PEDIDO INNER JOIN TB_PEDIDOS ON VW_QTD_FORMANDOS_POR_PEDIDO.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO  WHERE DATA_VENDA >= '"&INICIO&"' AND DATA_VENDA < '"&TERMINO&"' AND TB_PEDIDOS.COD_PEDIDO IN ( SELECT COD_PEDIDO FROM TB_PEDIDOS_FORMANDOS WHERE COD_FORMANDO IN (SELECT COD_FORMANDO FROM TB_PRODUCAO_LOTES_FORMANDOS WHERE DATA_SAIDA >= '"&INICIO&"' AND DATA_SAIDA < '"&TERMINO&"'))  AND FINALIZA_CORRECAO = 1 AND  TB_PEDIDOS.COD_VENDEDOR = {} "
    sql_3 = "SELECT TB_PRODUTOS.COD_PRODUTO, NOME_PRODUTO, SUM(QUANTIDADE) as QUANTIDADE "&_
    "FROM TB_PEDIDOS_FORMANDOS "&_
    "INNER JOIN TB_PEDIDOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_FORMANDOS.COD_PEDIDO  "&_
    "INNER JOIN TB_PRODUCAO_LOTES_FORMANDOS ON TB_PRODUCAO_LOTES_FORMANDOS.COD_FORMANDO = TB_PEDIDOS_FORMANDOS.COD_FORMANDO  "&_
    "INNER JOIN TB_PEDIDOS_PRODUTOS ON TB_PEDIDOS_PRODUTOS.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO "&_
    "INNER JOIN TB_PRODUTOS ON TB_PRODUTOS.COD_PRODUTO = TB_PEDIDOS_PRODUTOS.COD_PRODUTO "&_
    "WHERE TB_PEDIDOS.DATA_VENDA IS NOT NULL AND TB_PRODUCAO_LOTES_FORMANDOS.DATA_SAIDA >= '"&INICIO&"' "&_
    "AND TB_PRODUCAO_LOTES_FORMANDOS.DATA_SAIDA < '"&TERMINO&"' AND TB_PRODUTOS.COD_PRODUTO IN (41,35,34,1,16,31,7,9) "&_
    "AND TB_PRODUCAO_LOTES_FORMANDOS.COD_VENDEDOR = {} "&_
    " GROUP BY TB_PRODUTOS.COD_PRODUTO, NOME_PRODUTO ORDER BY TB_PRODUTOS.COD_PRODUTO"
    sql_4 = "SELECT (SUM(IFNULL(PAGAMENTO_EM_DINHEIRO,0)) + SUM(IFNULL(XTOTAL_CHEQUE,0)) + SUM(IFNULL(XTOTAL_BOLETO,0)) + SUM(IFNULL(XTOTAL_CARTAO,0))) AS TOTAL_RECEBIDO_GERAL FROM VW_EXTRATO_TOTAL_RECEBIDO_COMPLETO WHERE DATA_VENDA >= '"&INICIO&"' AND DATA_VENDA < '"&TERMINO&"' AND COD_VENDEDOR = {}"
    sql_5 = "SELECT SUM( VW_QTD_FORMANDOS_POR_PEDIDO.QTD_FORMANDOS ) AS TOTAL_ALBUNS_GERAL FROM VW_QTD_FORMANDOS_POR_PEDIDO INNER JOIN TB_PEDIDOS ON VW_QTD_FORMANDOS_POR_PEDIDO.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO WHERE TB_PEDIDOS.DATA_VENDA >= '"&INICIO&"' AND TB_PEDIDOS.DATA_VENDA < '"&TERMINO&"' AND FINALIZA_CORRECAO =1 AND COD_VENDEDOR = {};"
    AbrirConexao()
    Set vendedores = objConexao.Execute(sql)
    nomesProdutos = Array("DVD", "Trio", "Slide show", "Foto-tela", "Revista", "Mini &aacute;lbum", "Porta-retrato digital", "Bolsa")
    If vendedores.EOF Then
%>
<p class="alert alert-info" role="status">Nenhuma retirada encontrada para o per&iacute;odo.</p>
<%Else%>
<section id="ret-summary" hidden aria-labelledby="ret-title">
  <h2 id="ret-title">Resumo das retiradas</h2>
  <dl class="month-kpis">
    <div><dt>&Aacute;lbuns retirados</dt><dd id="ret-withdrawn"></dd></div>
    <div><dt>&Aacute;lbuns vendidos</dt><dd id="ret-sold"></dd></div>
    <div><dt>Aproveitamento</dt><dd id="ret-rate"></dd></div>
    <div><dt>Valor vendido</dt><dd id="ret-amount"></dd></div>
  </dl>
  <label class="ret-filter">Vendedor
    <select class="form-select" id="ret-seller"><option value="">Todos os vendedores</option></select>
  </label>
  <p class="ret-scope" id="ret-scope" role="status"></p>
  <div class="ret-dashboard">
    <section class="ret-panel" aria-labelledby="ret-chart-title">
      <div class="ret-panel-heading">
        <h3 id="ret-chart-title">Retiradas e vendas por vendedor</h3>
        <button type="button" class="ret-help-button" id="ret-help-button" aria-haspopup="dialog" aria-controls="ret-help" hidden><span aria-hidden="true">?</span> Como interpretar</button>
      </div>
      <p class="ret-legend" aria-hidden="true"><span>Retirados</span><span>Vendidos</span></p>
      <ol class="ret-bars" id="ret-chart"></ol>
    </section>
    <section class="ret-panel" aria-labelledby="ret-gauge-title">
      <h3 id="ret-gauge-title">Aproveitamento das retiradas</h3>
      <strong class="ret-dashboard-rate" id="ret-gauge-value"></strong>
      <meter id="ret-gauge" min="0" max="100" value="0" aria-labelledby="ret-gauge-title" aria-describedby="ret-gauge-caption"></meter>
      <p class="ret-dashboard-caption" id="ret-gauge-caption"></p>
    </section>
  </div>
</section>
<div class="month-scroll" tabindex="0" role="group" aria-label="Retiradas por vendedor">
<table class="month-table ret-table">
<thead><tr><th scope="col">Vendedor</th><th scope="col">Retirados</th><th scope="col">Vendidos</th><th scope="col">Valor vendido</th><th scope="col">Aproveitamento</th><th scope="col">Produtos vendidos</th><th scope="col">&Aacute;lbuns no per&iacute;odo</th><th scope="col">Arrecada&ccedil;&atilde;o no per&iacute;odo</th></tr></thead>
<tbody>
<%Do Until vendedores.EOF
    codigo = vendedores("COD_VENDEDOR").Value
    nome = Server.HTMLEncode(vendedores("NOME_VENDEDOR") & "")
    retirados = CDbl("0" & vendedores("QTD_RETIRADO"))
    valor = ValorRetirada(sql_1, codigo, "VALOR_VENDIDO")
    vendidos = ValorRetirada(sql_2, codigo, "QTD_VENDIDO")
    recebido = ValorRetirada(sql_4, codigo, "TOTAL_RECEBIDO_GERAL")
    albunsPeriodo = ValorRetirada(sql_5, codigo, "TOTAL_ALBUNS_GERAL")
    aproveitamento = 0
    If retirados > 0 Then aproveitamento = 100 * vendidos / retirados
    produtos = Array(0,0,0,0,0,0,0,0)
    Set produtoRs = objConexao.Execute(Replace(sql_3, "{}", codigo))
    Do Until produtoRs.EOF
        quantidadeProduto = produtoRs("QUANTIDADE").Value
        If IsNull(quantidadeProduto) Then quantidadeProduto = 0
        Select Case CLng(produtoRs("COD_PRODUTO"))
            Case 1: produtos(0) = quantidadeProduto
            Case 7: produtos(1) = quantidadeProduto
            Case 9: produtos(2) = quantidadeProduto
            Case 16: produtos(3) = quantidadeProduto
            Case 31: produtos(4) = quantidadeProduto
            Case 34: produtos(5) = quantidadeProduto
            Case 35: produtos(6) = quantidadeProduto
            Case 41: produtos(7) = quantidadeProduto
        End Select
        produtoRs.MoveNext
    Loop
    produtoRs.Close
%>
<tr data-seller="<%=codigo%>" data-withdrawn="<%=retirados%>" data-sold="<%=vendidos%>" data-amount="<%=Replace(CStr(valor), ",", ".")%>">
  <th scope="row"><%=nome%></th>
  <td><span class="ret-label" aria-hidden="true">Retirados</span><%=FormatNumber(retirados,0)%></td>
  <td><span class="ret-label" aria-hidden="true">Vendidos</span><%=FormatNumber(vendidos,0)%></td>
  <td><span class="ret-label" aria-hidden="true">Valor vendido</span><%=FormatCurrency(valor,2)%></td>
  <td><span class="ret-label" aria-hidden="true">Aproveitamento</span><%=FormatNumber(aproveitamento,2)%>%<meter class="ret-meter" min="0" max="100" value="<%=Replace(CStr(aproveitamento), ",", ".")%>" aria-label="Aproveitamento"></meter></td>
  <td class="ret-wide"><details><summary>Ver produtos</summary><ul class="ret-products">
    <%For i = 0 To 7%><li><span><%=nomesProdutos(i)%></span><strong><%=FormatNumber(produtos(i),0)%></strong></li><%Next%>
  </ul></details></td>
  <td><span class="ret-label" aria-hidden="true">&Aacute;lbuns no per&iacute;odo</span><%=FormatNumber(albunsPeriodo,0)%></td>
  <td><span class="ret-label" aria-hidden="true">Arrecada&ccedil;&atilde;o no per&iacute;odo</span><%=FormatCurrency(recebido,2)%></td>
</tr>
<%  vendedores.MoveNext
Loop%>
</tbody></table></div>
<%End If
vendedores.Close
FecharConexao()
End If%>
<dialog class="ret-help" id="ret-help" aria-labelledby="ret-help-title">
  <h2 id="ret-help-title">Como interpretar o gr&aacute;fico</h2>
  <h3>O que as barras mostram</h3>
  <p>A barra roxa mostra os &aacute;lbuns retirados e a rosa, os vendidos. Na vis&atilde;o geral, o gr&aacute;fico mostra os cinco vendedores com mais &aacute;lbuns vendidos. Ao selecionar um vendedor, as barras mostram os n&uacute;meros dele.</p>
  <h3>Como o aproveitamento &eacute; calculado</h3>
  <p>Divida os &aacute;lbuns vendidos pelos retirados e multiplique por 100. O medidor usa os totais do filtro atual.</p>
  <p class="ret-help-example"><strong id="ret-help-calculation"></strong></p>
  <p>A consulta de vendas considera pedidos com corre&ccedil;&atilde;o finalizada, vinculados a retiradas dentro do per&iacute;odo. Vendas feitas depois da data final ficam fora do c&aacute;lculo.</p>
  <form method="dialog"><button class="btn btn-primary" autofocus>Entendi</button></form>
</dialog>
</main>
<script>
(function () {
    var summary = document.getElementById('ret-summary');
    if (!summary) return;
    var select = document.getElementById('ret-seller');
    var rows = Array.from(document.querySelectorAll('.ret-table [data-seller]'));
    var number = new Intl.NumberFormat('pt-BR');
    var money = new Intl.NumberFormat('pt-BR', {style:'currency',currency:'BRL'});
    var percent = new Intl.NumberFormat('pt-BR', {style:'percent',maximumFractionDigits:2});
    var chart = document.getElementById('ret-chart');
    var chartTitle = document.getElementById('ret-chart-title');
    var gauge = document.getElementById('ret-gauge');
    var help = document.getElementById('ret-help');
    var helpButton = document.getElementById('ret-help-button');
    var calculation = document.getElementById('ret-help-calculation');
    if (typeof help.showModal === 'function') {
        helpButton.addEventListener('click', function () { help.showModal(); });
        help.addEventListener('close', function () { helpButton.focus(); });
        helpButton.hidden = false;
    }
    var sellers = rows.map(function (row) {
        var name = row.cells[0].textContent.trim();
        select.add(new Option(name, row.dataset.seller));
        return {
            row:row,
            name:name,
            withdrawn:Number(row.dataset.withdrawn),
            sold:Number(row.dataset.sold),
            amount:Number(row.dataset.amount)
        };
    });

    function chartBar(value, maximum, label) {
        var line = document.createElement('div');
        var track = document.createElement('span');
        var bar = document.createElement('span');
        var count = document.createElement('span');
        line.className = 'ret-bar-row';
        line.setAttribute('aria-label', label + ': ' + number.format(value));
        line.setAttribute('role', 'img');
        track.className = 'ret-bar-track';
        bar.style.width = (maximum ? 100 * Math.max(0, value) / maximum : 0) + '%';
        count.textContent = number.format(value);
        track.appendChild(bar);
        line.append(track, count);
        return line;
    }

    function update() {
        var selected = select.value;
        var withdrawn = 0, sold = 0, amount = 0;
        var visible = sellers.filter(function (seller) {
            seller.row.hidden = Boolean(selected && seller.row.dataset.seller !== selected);
            return !seller.row.hidden;
        });
        visible.forEach(function (seller) {
            withdrawn += seller.withdrawn;
            sold += seller.sold;
            amount += seller.amount;
        });
        var rate = withdrawn ? sold / withdrawn : 0;
        document.getElementById('ret-withdrawn').textContent = number.format(withdrawn);
        document.getElementById('ret-sold').textContent = number.format(sold);
        document.getElementById('ret-rate').textContent = percent.format(rate);
        document.getElementById('ret-amount').textContent = money.format(amount);
        document.getElementById('ret-scope').textContent = selected ? 'Vendedor: ' + select.options[select.selectedIndex].text : '';

        var top = visible.slice().sort(function (a, b) {
            return b.sold - a.sold || b.withdrawn - a.withdrawn || a.name.localeCompare(b.name, 'pt-BR');
        }).slice(0, 5);
        var maximum = top.reduce(function (max, seller) { return Math.max(max, seller.withdrawn, seller.sold); }, 0);
        chartTitle.textContent = selected ? 'Retiradas e vendas do vendedor' : 'Cinco vendedores com mais \u00e1lbuns vendidos';
        chart.replaceChildren();
        top.forEach(function (seller) {
            var item = document.createElement('li');
            var name = document.createElement('strong');
            name.textContent = seller.name;
            item.append(name, chartBar(seller.withdrawn, maximum, 'Retirados'), chartBar(seller.sold, maximum, 'Vendidos'));
            chart.appendChild(item);
        });
        calculation.textContent = withdrawn
            ? number.format(sold) + ' \xf7 ' + number.format(withdrawn) + ' \xd7 100 = ' + percent.format(rate)
            : 'Sem retiradas no filtro atual, o painel exibe 0%.';
        gauge.value = Math.min(100, Math.max(0, 100 * rate));
        document.getElementById('ret-gauge-value').textContent = percent.format(rate);
        document.getElementById('ret-gauge-caption').textContent = number.format(sold) + ' vendidos / ' + number.format(withdrawn) + ' retirados';
    }
    var productLists = Array.from(document.querySelectorAll('.ret-table details'));
    var closedBeforePrint = [];
    window.addEventListener('beforeprint', function () {
        closedBeforePrint = productLists.filter(function (details) { return !details.open; });
        closedBeforePrint.forEach(function (details) { details.open = true; });
    });
    window.addEventListener('afterprint', function () {
        closedBeforePrint.forEach(function (details) { details.open = false; });
    });
    select.addEventListener('change', update);
    window.addEventListener('pageshow', update);
    update();
    summary.hidden = false;
})();
</script>
</body>
</html>
