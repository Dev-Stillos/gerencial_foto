<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim rs, rsFotos, rsProdutos, mes, ano, periodo, erroPeriodo, formatoPeriodo, embutido
Dim totalFotos, valorFotos, totalVendido, valorProduto

AbrirConexao()
Set rs = objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '43' and cod_perfil = '" & Session("cod_perfil") & "'")
If rs.EOF Then
    rs.Close
    FecharConexao()
    Session.Abandon()
    Response.Redirect("index.asp")
End If
rs.Close
Set rs = Nothing

mes = Trim(Request.QueryString("MES_INICIO"))
ano = Trim(Request.QueryString("ANO_INICIO"))
embutido = Request.QueryString("modal") = "1"

Set formatoPeriodo = New RegExp
formatoPeriodo.Pattern = "^(0?[1-9]|1[0-2])/[1-9][0-9]{3}$"
erroPeriodo = Not formatoPeriodo.Test(mes & "/" & ano)
Set formatoPeriodo = Nothing

If erroPeriodo Then
    FecharConexao()
Else
    periodo = DateSerial(CInt(ano), CInt(mes), 1)
    totalFotos = 0
    valorFotos = 0
    Set rsFotos = objConexao.Execute("SELECT SUM(FATURADO) AS FATURADO, SUM(FATURADO*VALOR_UN) AS TOTAL_FATURADO_ARRECADADO FROM VW_PEDIDOS_COMPLETO WHERE month(DATA_VENDA) = "&Month(periodo)&" AND year(DATA_VENDA) = "&Year(periodo)&" and FINALIZA_CORRECAO = 1")
    If Not rsFotos.EOF Then
        totalFotos = FormatNumber(0 & rsFotos("FATURADO"), 0)
        valorFotos = FormatNumber(0 & rsFotos("TOTAL_FATURADO_ARRECADADO"), 0)
    End If
    rsFotos.Close
    Set rsFotos = Nothing

    totalVendido = 0 + valorFotos
    Set rsProdutos = objConexao.Execute("SELECT TB_PEDIDOS_PRODUTOS.COD_PRODUTO, TB_PRODUTOS.NOME_PRODUTO, SUM(TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS QUANTIDADE, SUM(TB_PEDIDOS_PRODUTOS.VALOR_UN*TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS VALOR_TOTAL FROM TB_PEDIDOS  INNER JOIN TB_PEDIDOS_PRODUTOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO INNER JOIN TB_PRODUTOS ON TB_PEDIDOS_PRODUTOS.COD_PRODUTO = TB_PRODUTOS.COD_PRODUTO WHERE month(TB_PEDIDOS.DATA_VENDA) = "&Month(periodo)&" AND  year(TB_PEDIDOS.DATA_VENDA) = "&Year(periodo)&" and TB_PEDIDOS.FINALIZA_CORRECAO = 1 GROUP BY TB_PEDIDOS_PRODUTOS.COD_PRODUTO ")
End If
%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Relat&oacute;rio Mensal de Produtos Vendidos | Portal Stillos</title>
<link rel="stylesheet" href="css/vendor/bootstrap.min.css">
<link rel="stylesheet" href="css/dashboard.css">
<style>
body { background:#f4f4f8; color:#242738; }
.sales-report { max-width:1150px; margin:24px auto; padding:clamp(16px,4vw,36px); background:#fff; border:1px solid #e4dfeb; border-radius:14px; }
.sales-tools,.sales-header { display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; }
.sales-header { margin:16px 0 22px; padding-bottom:14px; border-bottom:2px solid #b20b61; }
.sales-header h1 { margin:0 0 6px; font-size:clamp(1.45rem,4vw,2rem); }
.sales-header p { margin:0; color:#53566b; text-transform:capitalize; }
.sales-logo { width:150px; padding:10px; border-radius:8px; background:#292b40; }
.sales-kpis { display:grid; grid-template-columns:repeat(auto-fit,minmax(min(100%,200px),1fr)); gap:12px; margin:0 0 24px; }
.sales-kpis > div { padding:14px; border:1px solid #e4dfeb; border-radius:10px; }
.sales-kpis dd { margin:4px 0 0; font-weight:700; font-size:1.25rem; }
.sales-scroll { overflow:auto; }
.sales-table { width:100%; border-collapse:collapse; font-variant-numeric:tabular-nums; }
.sales-table :is(th,td) { padding:10px; border-bottom:1px solid #e4dfeb; overflow-wrap:anywhere; }
.sales-table .sales-number { text-align:right; white-space:nowrap; }
.sales-table thead,.sales-table tfoot { background:#f1edf6; }
.sales-table tbody tr:nth-child(even) { background:#faf9fc; }
.sales-notes { margin-top:20px; color:#53566b; }
.sales-notes h2,.sales-report section h2 { font-size:1.1rem; margin:0 0 8px; }
.sales-notes ul { margin:0; padding-left:20px; }
.sales-notes li + li { margin-top:4px; }
.report-embedded .sales-report { margin:12px auto; }
@page { size:A4 portrait; margin:12mm; }
@media print {
    body { background:#fff; color:#000; font-size:10pt; }
    .sales-report,.report-embedded .sales-report { margin:0; padding:0; max-width:none; border:0; }
    .sales-tools { display:none; }
    .sales-header { margin-top:0; }
    .sales-scroll { overflow:visible; }
    .sales-table thead { display:table-header-group; }
    .sales-table tfoot { display:table-row-group; }
    .sales-table :is(th,td) { padding:5px; }
    .sales-table tr,.sales-kpis,.sales-notes { break-inside:avoid; }
    .sales-logo { print-color-adjust:exact; -webkit-print-color-adjust:exact; }
}
</style>
</head>
<body<%If embutido Then%> class="report-embedded"<%End If%>>
<main class="sales-report" data-report-preview-content>
<%If Not embutido Then%>
    <nav class="sales-tools" aria-label="A&ccedil;&otilde;es do relat&oacute;rio">
        <a href="adm_relatorios.asp" class="btn btn-outline-secondary">Voltar aos relat&oacute;rios</a>
        <%If Not erroPeriodo Then%>
        <button type="button" class="btn btn-primary" onclick="window.print()">Imprimir / Salvar PDF</button>
        <%End If%>
    </nav>
<%End If%>
    <header class="sales-header">
        <div>
            <h1>Relat&oacute;rio Mensal de Produtos Vendidos</h1>
            <%If Not erroPeriodo Then%><p><%=MonthName(Month(periodo))%> de <%=Year(periodo)%></p><%End If%>
        </div>
        <img class="sales-logo" src="images/logoparafundopreto.png" alt="Stillos">
    </header>
<%If erroPeriodo Then%>
    <p class="alert alert-warning" data-report-error>Selecione um m&ecirc;s e um ano v&aacute;lidos para abrir o relat&oacute;rio.</p>
<%Else%>
    <dl class="sales-kpis">
        <div><dt>Fotos vendidas</dt><dd><%=totalFotos%></dd></div>
        <div><dt>Valor das fotos vendidas</dt><dd><%=FormatCurrency(valorFotos)%></dd></div>
    </dl>
    <section aria-labelledby="sales-products-title">
        <h2 id="sales-products-title">Venda de produtos</h2>
        <div class="sales-scroll">
            <table class="sales-table" aria-labelledby="sales-products-title">
                <thead>
                    <tr><th scope="col">Produto</th><th scope="col" class="sales-number">Quantidade</th><th scope="col" class="sales-number">Subtotal</th></tr>
                </thead>
                <tbody>
<%If rsProdutos.EOF Then%>
                    <tr><td colspan="3">Nenhum produto vendido no per&iacute;odo.</td></tr>
<%Else
    Do While Not rsProdutos.EOF
        valorProduto = FormatNumber(0 & rsProdutos("VALOR_TOTAL"))
        totalVendido = totalVendido + valorProduto
%>
                    <tr>
                        <th scope="row"><%=Server.HTMLEncode(rsProdutos("NOME_PRODUTO") & "")%></th>
                        <td class="sales-number"><%=rsProdutos("QUANTIDADE")%></td>
                        <td class="sales-number"><%=FormatCurrency(valorProduto)%></td>
                    </tr>
<%
        rsProdutos.MoveNext
    Loop
End If
rsProdutos.Close
Set rsProdutos = Nothing
FecharConexao()
%>
                </tbody>
                <tfoot>
                    <tr><th scope="row" colspan="2">Total vendido (fotos e produtos)</th><td class="sales-number"><strong><%=FormatCurrency(totalVendido)%></strong></td></tr>
                </tfoot>
            </table>
        </div>
    </section>
    <aside class="sales-notes" aria-labelledby="sales-notes-title">
        <h2 id="sales-notes-title">Crit&eacute;rios do relat&oacute;rio</h2>
        <ul>
            <li>Entram somente vendas corrigidas pelo departamento fotogr&aacute;fico.</li>
            <li>Fotos produzidas e vendidas aparecem no resumo. Fotos vendidas &agrave; parte aparecem na tabela de produtos.</li>
            <li>Os valores declarados na venda podem diferir do total arrecadado devido ao arredondamento no faturamento.</li>
        </ul>
    </aside>
<%End If%>
</main>
</body>
</html>
