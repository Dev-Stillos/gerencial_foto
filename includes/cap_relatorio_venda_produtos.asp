<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%Dim modoVendaFornecedor : modoVendaFornecedor = "lista"%>
<%
Function FornecedorData(valor)
    Dim formato, partes, dia, mes, ano, dataValida
    Set formato = New RegExp
    formato.Pattern = "^([0-9]{4}-[0-9]{2}-[0-9]{2}|[0-9]{1,2}/[0-9]{1,2}/[0-9]{4})$"
    If Not formato.Test(valor) Then Exit Function
    If InStr(valor, "-") > 0 Then
        partes = Split(valor, "-")
        ano = CInt(partes(0))
        mes = CInt(partes(1))
        dia = CInt(partes(2))
    Else
        partes = Split(valor, "/")
        dia = CInt(partes(0))
        mes = CInt(partes(1))
        ano = CInt(partes(2))
    End If
    If ano < 1000 Or mes < 1 Or mes > 12 Or dia < 1 Or dia > 31 Then Exit Function
    dataValida = DateSerial(ano, mes, dia)
    If Month(dataValida) = mes And Day(dataValida) = dia Then FornecedorData = dataValida
End Function

Function FornecedorHtml(valor)
    FornecedorHtml = Server.HTMLEncode(valor & "")
End Function

Dim DATA_INICIO, DATA_FIM, COD_PRODUTO_DETALHADO, vendaErro, vendaEmbutida, vendaRecibo, vendaLista
Dim vendaRs, vendaContratos, vendaFormandos, vendaSql, vendaCodigo, vendaFormato
Dim vendaNome, vendaValor, vendaMostrarNome, vendaTitulo, vendaPeriodo, vendaParametros, vendaVoltar, vendaQuantidade
Dim vendaContrato, vendaCodigoContrato, vendaTituloContrato, vendaSeparadorContrato

vendaLista = modoVendaFornecedor = "lista"
vendaRecibo = modoVendaFornecedor = "recibo"
vendaEmbutida = Request.QueryString("modal") = "1"
vendaTitulo = "Venda de produtos para fornecedor"
vendaErro = ""
vendaVoltar = "adm_relatorios.asp"

AbrirConexao()
Set vendaRs = objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '44' and cod_perfil = '"&Session("cod_perfil")&"'")
If vendaRs.EOF Then
    vendaRs.Close
    FecharConexao()
    Session.Abandon()
    Response.Redirect("index.asp")
End If
vendaRs.Close

If vendaRecibo Then
    vendaTitulo = "Recibo de fornecedor"
    DATA_INICIO = FornecedorData(Trim(Request.Form("DATA_INICIO")))
    DATA_FIM = FornecedorData(Trim(Request.Form("DATA_FIM")))
    vendaCodigo = Trim(Request.Form("detalhar_produto"))
    vendaNome = UCase(Trim(Request.Form("nome_fornecedor")))
    vendaValor = Trim(Request.Form("valor_produto"))
    vendaMostrarNome = Request.Form("imprimir_nome_formando") = "S"
Else
    DATA_INICIO = FornecedorData(Trim(Request.QueryString("DATA_INI")))
    DATA_FIM = FornecedorData(Trim(Request.QueryString("DATA_FIM")))
    vendaCodigo = Trim(Request.QueryString("detalhar_produto"))
    If Not vendaLista Then vendaTitulo = "Venda por contrato e formando"
End If

If IsEmpty(DATA_INICIO) Or IsEmpty(DATA_FIM) Then
    vendaErro = "Informe uma data inicial e uma data final v&aacute;lidas."
ElseIf DATA_INICIO > DATA_FIM Then
    vendaErro = "A data inicial deve ser anterior ou igual &agrave; data final."
Else
    vendaPeriodo = ArrumaData(DATA_INICIO) & " a " & ArrumaData(DATA_FIM)
    vendaParametros = "DATA_INI=" & Server.URLEncode(ArrumaData(DATA_INICIO)) & "&amp;DATA_FIM=" & Server.URLEncode(ArrumaData(DATA_FIM))
    If vendaEmbutida Then vendaParametros = vendaParametros & "&amp;modal=1"
    If Not vendaLista Then vendaVoltar = "cap_relatorio_venda_produtos.asp?" & vendaParametros
End If

Set vendaFormato = New RegExp
If Not vendaLista And Len(vendaErro) = 0 Then
    vendaFormato.Pattern = "^[0-9]{1,10}$"
    If vendaFormato.Test(vendaCodigo) Then
        COD_PRODUTO_DETALHADO = vendaCodigo
        If vendaRecibo Then vendaVoltar = "cap_relatorio_venda_produtos_detalhado.asp?" & vendaParametros & "&amp;detalhar_produto=" & COD_PRODUTO_DETALHADO
    Else
        vendaErro = "Selecione um produto v&aacute;lido."
    End If
End If

If vendaRecibo And Len(vendaErro) = 0 Then
    If Len(vendaValor) = 0 Then vendaValor = "0"
    vendaFormato.Pattern = "^([0-9]{1,12}|[0-9]{1,3}(\.[0-9]{3}){1,3})(,[0-9]{1,2})?$"
    If vendaFormato.Test(vendaValor) Then
        Dim vendaValorPartes
        vendaValorPartes = Split(Replace(vendaValor, ".", ""), ",")
        vendaValor = CCur(vendaValorPartes(0))
        If UBound(vendaValorPartes) = 1 Then
            vendaValor = vendaValor + CCur(CInt(vendaValorPartes(1)) / (10 ^ Len(vendaValorPartes(1))))
        End If
    Else
        vendaErro = "Informe o valor unit&aacute;rio em reais, no formato 12,50."
    End If
End If
Set vendaFormato = Nothing

If Len(vendaErro) = 0 Then
    If vendaLista Then
        vendaSql = "SELECT TB_PEDIDOS_PRODUTOS.COD_PRODUTO, SUM(TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS QTD_POR_PRODUTO, TB_PRODUTOS.NOME_PRODUTO FROM TB_PEDIDOS INNER JOIN TB_PEDIDOS_PRODUTOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO INNER JOIN TB_PRODUTOS on TB_PEDIDOS_PRODUTOS.COD_PRODUTO = TB_PRODUTOS.COD_PRODUTO WHERE TB_PEDIDOS.FINALIZA_CORRECAO = 1 AND  TB_PEDIDOS.DATA_VENDA BETWEEN "&sqlData(DATA_INICIO)&" AND "&sqlData(DATA_FIM)&" GROUP BY TB_PEDIDOS_PRODUTOS.COD_PRODUTO ORDER BY TB_PRODUTOS.NOME_PRODUTO"
    Else
        vendaSql = "SELECT TB_PEDIDOS_PRODUTOS.COD_PRODUTO, SUM(TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS QTD_POR_PRODUTO, TB_PRODUTOS.NOME_PRODUTO FROM TB_PEDIDOS INNER JOIN TB_PEDIDOS_PRODUTOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO INNER JOIN TB_PRODUTOS on TB_PEDIDOS_PRODUTOS.COD_PRODUTO = TB_PRODUTOS.COD_PRODUTO WHERE TB_PEDIDOS.FINALIZA_CORRECAO = 1 AND  TB_PEDIDOS.DATA_VENDA BETWEEN "&sqlData(DATA_INICIO)&" AND "&sqlData(DATA_FIM)&" AND TB_PEDIDOS_PRODUTOS.COD_PRODUTO = '"&COD_PRODUTO_DETALHADO&"' GROUP BY TB_PEDIDOS_PRODUTOS.COD_PRODUTO ORDER BY TB_PRODUTOS.NOME_PRODUTO"
    End If
    Set vendaRs = objConexao.Execute(vendaSql)
    If vendaRs.EOF And Not vendaLista Then vendaErro = "Nenhuma venda encontrada para este produto no per&iacute;odo."
End If
%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title><%=vendaTitulo%> | Portal Stillos</title>
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

.sales-tools a { margin-right:auto; }
.sales-report section + section { margin-top:22px; }
.sales-report .sales-contract { margin:24px 0 8px; }
.sales-form { display:flex; flex-wrap:wrap; align-items:end; gap:12px; margin:0 0 24px; }
.sales-form label span { display:block; margin-bottom:4px; }
.sales-form fieldset { margin:0; padding:0; border:0; min-width:0; }
.sales-form legend { float:none; font-size:inherit; margin:0 0 4px; }
.sales-form fieldset label { margin-right:12px; }
.sales-signature { margin:34px auto 0; padding-top:8px; border-top:1px solid; max-width:320px; text-align:center; }
@media print {
    /* A previa em iframe pode abrir a impressao sem as margens da regra @page. */
    html,body { margin:0 !important; }
    body.report-embedded { padding:8mm 9mm !important; box-sizing:border-box; }
    .sales-action,.sales-form { display:none !important; }
    .sales-header { flex-wrap:nowrap; gap:5mm; margin:0 0 5mm; padding-bottom:3mm; border-bottom:1pt solid #444; }
    .sales-header h1 { font-size:17pt; line-height:1.15; margin-bottom:2mm; }
    .sales-header p { font-size:10pt; text-transform:none; }
    .sales-logo { width:31mm; padding:2mm; flex-shrink:0; }
    .sales-kpis { gap:3mm; margin:0 0 5mm; }
    .sales-kpis > div { padding:2.5mm 3mm; border:0.5pt solid #aaa; border-radius:2mm; }
    .sales-kpis dd { font-size:11pt; }
    .sales-report section + section { margin-top:5mm; }
    .sales-report .sales-contract { margin:0 0 2mm; font-size:10.5pt; line-height:1.25; break-after:avoid; }
    .sales-contract-quantity { white-space:nowrap; }
    .sales-table { font-size:10pt; }
    .sales-table :is(th,td) { padding:1.8mm 2mm; }
    .sales-table tbody th { font-weight:normal; }
    .sales-table thead th { border-bottom:0.75pt solid #777; }
    .sales-table a { color:inherit; text-decoration:none; }
    .sales-table .sales-order { width:22mm; text-align:right; white-space:nowrap; }
    .sales-table .sales-number { width:27mm; }
    .sales-signature { break-inside:avoid; }
}
</style>
</head>
<body<%If vendaEmbutida Then%> class="report-embedded"<%End If%>>
<main class="sales-report" data-report-preview-content>
    <nav class="sales-tools" aria-label="A&ccedil;&otilde;es do relat&oacute;rio">
        <%If Not vendaEmbutida Or Not vendaLista Then%>
        <a class="btn btn-outline-secondary" href="<%=vendaVoltar%>">Voltar</a>
        <%End If%>
        <%If Not vendaEmbutida And Len(vendaErro) = 0 Then%>
        <button class="btn btn-primary" type="button" onclick="window.print()">Imprimir / Salvar PDF</button>
        <%End If%>
    </nav>
    <header class="sales-header">
        <div><h1><%=vendaTitulo%></h1><p><%=FornecedorHtml(vendaPeriodo)%></p></div>
        <img class="sales-logo" src="images/logoparafundopreto.png" alt="Stillos">
    </header>
<%If Len(vendaErro) > 0 Then%>
    <p class="alert alert-warning" data-report-error><%=vendaErro%></p>
<%ElseIf vendaLista Then%>
    <div class="sales-scroll">
        <table class="sales-table">
            <thead><tr><th scope="col">Produto</th><th scope="col" class="sales-number">Quantidade vendida</th><th scope="col" class="sales-action">Detalhamento</th></tr></thead>
            <tbody>
            <%If vendaRs.EOF Then%>
            <tr><td colspan="3">Nenhum produto encontrado no per&iacute;odo.</td></tr>
            <%End If
            Do While Not vendaRs.EOF%>
            <tr>
                <th scope="row"><%=FornecedorHtml(vendaRs("NOME_PRODUTO"))%></th>
                <td class="sales-number"><%=FornecedorHtml(vendaRs("QTD_POR_PRODUTO"))%></td>
                <td class="sales-action"><a href="cap_relatorio_venda_produtos_detalhado.asp?<%=vendaParametros%>&amp;detalhar_produto=<%=Server.URLEncode(vendaRs("COD_PRODUTO") & "")%>">Ver contratos e formandos</a></td>
            </tr>
            <%  vendaRs.MoveNext
            Loop%>
            </tbody>
        </table>
    </div>
<%Else
    vendaQuantidade = vendaRs("QTD_POR_PRODUTO")
%>
    <dl class="sales-kpis">
        <div><dt>Produto</dt><dd><%=FornecedorHtml(vendaRs("NOME_PRODUTO"))%></dd></div>
        <div><dt>Quantidade total</dt><dd><%=FornecedorHtml(vendaQuantidade)%></dd></div>
        <%If vendaRecibo Then%>
        <div><dt>Valor unit&aacute;rio</dt><dd><%=FormatCurrency(vendaValor)%></dd></div>
        <div><dt>Valor total</dt><dd><%=FormatCurrency(CDbl(vendaQuantidade) * CDbl(vendaValor))%></dd></div>
        <%End If%>
    </dl>
    <%If vendaRecibo Then%>
    <p><strong>Fornecedor</strong> <%=FornecedorHtml(vendaNome)%><br><strong>Data de emiss&atilde;o</strong> <%=ArrumaData(Date())%></p>
    <%Else%>
    <form class="sales-form" method="post" target="_blank" action="cap_relatorio_venda_produtos_detalhado_imprimir.asp">
        <input type="hidden" name="DATA_INICIO" value="<%=FornecedorHtml(ArrumaData(DATA_INICIO))%>">
        <input type="hidden" name="DATA_FIM" value="<%=FornecedorHtml(ArrumaData(DATA_FIM))%>">
        <input type="hidden" name="detalhar_produto" value="<%=COD_PRODUTO_DETALHADO%>">
        <label><span>Fornecedor</span><input class="form-control" name="nome_fornecedor"></label>
        <label><span>Valor unit&aacute;rio (R$)</span><input class="form-control" name="valor_produto" inputmode="decimal" placeholder="0,00"></label>
        <fieldset><legend>Incluir nomes no recibo</legend>
            <label><input type="radio" name="imprimir_nome_formando" value="S" checked> Sim</label>
            <label><input type="radio" name="imprimir_nome_formando" value="N"> N&atilde;o</label>
        </fieldset>
        <button class="btn btn-primary" type="submit">Gerar recibo</button>
    </form>
    <%End If
    Set vendaContratos = objConexao.Execute("SELECT TB_PEDIDOS_PRODUTOS.COD_PRODUTO, SUM(TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS QTD_POR_CONTRATO, TB_FICHAS_IDENTIFICACAO.COD_CONTRATO, TB_ENTIDADES.DESCRICAO FROM TB_PEDIDOS INNER JOIN TB_PEDIDOS_PRODUTOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO INNER JOIN TB_FICHAS_IDENTIFICACAO ON TB_PEDIDOS_PRODUTOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO INNER JOIN TB_ENTIDADES ON TB_FICHAS_IDENTIFICACAO.COD_CONTRATO = TB_ENTIDADES.COD_CONTRATO  WHERE TB_PEDIDOS.FINALIZA_CORRECAO = 1 AND TB_PEDIDOS.DATA_VENDA BETWEEN "&sqlData(DATA_INICIO)&" AND "&sqlData(DATA_FIM)&" AND  TB_PEDIDOS_PRODUTOS.COD_PRODUTO = '"&COD_PRODUTO_DETALHADO&"' GROUP BY TB_FICHAS_IDENTIFICACAO.COD_CONTRATO,TB_PEDIDOS_PRODUTOS.COD_PRODUTO")
    If vendaContratos.EOF Then%>
    <p>Nenhum contrato encontrado para detalhar este produto.</p>
    <%End If
    Do While Not vendaContratos.EOF
        vendaCodigoContrato = Trim(vendaContratos("COD_CONTRATO") & "")
        vendaTituloContrato = Trim(vendaContratos("DESCRICAO") & "")
        If Len(vendaTituloContrato) = 0 Then
            vendaTituloContrato = vendaCodigoContrato
        ElseIf Len(vendaCodigoContrato) > 0 Then
            vendaSeparadorContrato = Mid(vendaTituloContrato, Len(vendaCodigoContrato) + 1, 1)
            If StrComp(Left(vendaTituloContrato, Len(vendaCodigoContrato)), vendaCodigoContrato, vbTextCompare) <> 0 _
                Or InStr(" -:/", vendaSeparadorContrato) = 0 Then
                vendaTituloContrato = vendaCodigoContrato & " - " & vendaTituloContrato
            End If
        End If
        vendaContrato = Replace(vendaContratos("COD_CONTRATO") & "", "'", "''")
        Set vendaFormandos = objConexao.Execute("SELECT TB_PEDIDOS_PRODUTOS.*, SUM(TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS QTD_POR_FORMANDO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, TB_PEDIDOS.COD_VENDEDOR, TB_PEDIDOS.DATA_VENDA FROM TB_PEDIDOS INNER JOIN TB_PEDIDOS_PRODUTOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO INNER JOIN TB_FICHAS_IDENTIFICACAO ON TB_PEDIDOS_PRODUTOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO WHERE TB_PEDIDOS.DATA_VENDA BETWEEN "&sqlData(DATA_INICIO)&" AND "&sqlData(DATA_FIM)&" AND COD_PRODUTO = '"&COD_PRODUTO_DETALHADO&"' AND TB_FICHAS_IDENTIFICACAO.COD_CONTRATO = '"&vendaContrato&"' AND TB_PEDIDOS.FINALIZA_CORRECAO = 1 GROUP BY COD_FORMANDO")
    %>
    <section>
        <h2 class="sales-contract"><%=FornecedorHtml(vendaTituloContrato)%> <span class="sales-contract-quantity">&middot; <%=FornecedorHtml(vendaContratos("QTD_POR_CONTRATO"))%> un.</span></h2>
        <div class="sales-scroll">
            <table class="sales-table">
                <thead><tr><th scope="col">Formando</th><%If Not vendaRecibo Then%><th scope="col" class="sales-order">Pedido</th><%End If%><th scope="col" class="sales-number">Quantidade</th></tr></thead>
                <tbody>
                <%Do While Not vendaFormandos.EOF%>
                <tr>
                    <th scope="row">
                        <%If vendaRecibo Then%>
                            <%=FornecedorHtml(digito(vendaFormandos("COD_FORMANDO")))%>
                            <%If vendaMostrarNome Then%> <%=FornecedorHtml(vendaFormandos("NOME_FORMANDO"))%><%End If%>
                        <%Else%>
                            <a target="_blank" rel="noopener" href="dip_busca_detalhamento_formando.asp?cod_formando=<%=Server.URLEncode(vendaFormandos("COD_FORMANDO") & "")%>"><%=FornecedorHtml(digito(vendaFormandos("COD_FORMANDO")))%></a>
                            <%=FornecedorHtml(vendaFormandos("NOME_FORMANDO"))%>
                        <%End If%>
                    </th>
                    <%If Not vendaRecibo Then%>
                    <td class="sales-order"><a target="_blank" rel="noopener" href="lca_busca_detalhamento_pedido.asp?cod_pedido=<%=Server.URLEncode(vendaFormandos("COD_PEDIDO") & "")%>"><%=FornecedorHtml(vendaFormandos("COD_PEDIDO"))%></a></td>
                    <%End If%>
                    <td class="sales-number"><%=FornecedorHtml(vendaFormandos("QTD_POR_FORMANDO"))%></td>
                </tr>
                <%  vendaFormandos.MoveNext
                Loop%>
                </tbody>
            </table>
        </div>
    </section>
    <%
        vendaFormandos.Close
        vendaContratos.MoveNext
    Loop
    vendaContratos.Close
    If vendaRecibo Then%><p class="sales-signature">Assinatura do fornecedor</p><%End If
End If
If vendaRs.State <> 0 Then vendaRs.Close
Set vendaRs = Nothing
FecharConexao()
%>
</main>
</body>
</html>

