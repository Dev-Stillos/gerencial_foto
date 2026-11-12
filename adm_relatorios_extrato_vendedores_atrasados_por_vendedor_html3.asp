<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<%
If not validaUsuario(2,,15,false) Then
  Response.Redirect("inicio.asp")
End If


Dim relErro, relPeriodo, posseVendedor, posseAtrasados, possePadrao, posseWhere
Dim posseRs, posseResumo, posseTotal, posseMedia, posseNome
posseAtrasados = (Request.QueryString("SOMENTE_VENDEDORES_EM_ATRASO") = "S")
posseVendedor = Trim(Request.QueryString("COD_VENDEDOR_BLOQUEADO"))
If Len(Request.QueryString("SOMENTE_VENDEDORES_EM_ATRASO")) = 0 Then
    If Len(Request.QueryString("COD_VENDEDOR")) > 0 Then
        posseVendedor = Trim(Request.QueryString("COD_VENDEDOR"))
    Else
        posseAtrasados = True
    End If
End If

Set possePadrao = New RegExp
possePadrao.Pattern = "^[0-9]{1,10}$"
If Not possePadrao.Test(posseVendedor) Then
    relErro = "Selecione um vendedor para consultar o relat&oacute;rio."
ElseIf CDbl(posseVendedor) = 0 Then
    relErro = "Selecione um vendedor para consultar o relat&oacute;rio."
End If

If Len(relErro) = 0 Then
    If posseAtrasados Then
        relPeriodo = "Somente álbuns em atraso"
    Else
        relPeriodo = "Todos os álbuns em posse"
    End If
    AbrirConexao()
    posseWhere = " FROM VW_MOVIMENTACOES_COMPLETA WHERE COD_VENDEDOR = " & posseVendedor & _
        " AND DATA_SAIDA IS NOT NULL AND DATA_RETORNO IS NULL AND TIPO = 0"
    If posseAtrasados Then posseWhere = posseWhere & " AND DIAS_ATRASADOS >= 0"
    Set posseResumo = objConexao.Execute("SELECT COUNT(*) AS TOTAL, AVG(DIAS_ATRASADOS) AS MEDIA" & posseWhere)
    posseTotal = posseResumo("TOTAL").Value
    posseMedia = posseResumo("MEDIA").Value
    posseResumo.Close
    Set posseResumo = Nothing
    Set posseRs = objConexao.Execute("SELECT COD_FORMANDO, NOME_FORMANDO, NOME_VENDEDOR, DATA_SAIDA, " & _
        "DIAS_ATRASADOS, ALBUM, DESCRICAO" & posseWhere & " ORDER BY DESCRICAO, NOME_FORMANDO, COD_FORMANDO")
    If Not posseRs.EOF Then posseNome = posseRs("NOME_VENDEDOR").Value
End If
%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>&Aacute;lbuns em posse do vendedor | Portal Stillos</title>
<link rel="stylesheet" href="css/vendor/bootstrap.min.css">
<link rel="stylesheet" href="css/dashboard.css">
<style>
.month-report { max-width:1150px; margin:24px auto; padding:clamp(16px,4vw,36px); background:#fff; border:1px solid #e4dfeb; border-radius:14px; }
.month-tools,.month-header { display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:12px; }
.month-header { margin:20px 0; padding-bottom:14px; border-bottom:2px solid #b20b61; }
.month-header p { margin:0; }
.month-logo { width:150px; padding:10px; border-radius:8px; background:#292b40; }
.month-kpis { display:grid; grid-template-columns:repeat(auto-fit,minmax(min(100%,175px),1fr)); gap:12px; margin:12px 0 16px; }
.month-kpis > div { min-width:0; padding:0; }
.month-kpis { padding:12px 0 16px; border-bottom:1px solid #e4dfeb; }
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
@page { size:A4 landscape; margin:10mm; }
@media print {
  body { margin:0; background:#fff; color:#111; font-size:10pt; line-height:1.3; }
  .month-report { margin:0; padding:6mm 7mm; max-width:none; border:0; border-radius:0; }
  .month-tools { display:none; }
  .month-header { margin:0 0 5mm; padding-bottom:3mm; border-bottom:1px solid #555; break-after:avoid; }
  .month-header h1 { font-size:16pt; margin:0 0 1mm; }
  .month-logo { width:30mm; padding:0; background:transparent; filter:brightness(0); }
  .month-scroll { overflow:visible; }
  .month-table { min-width:0; width:100%; table-layout:fixed; margin-top:0; }
  .month-table th,.month-table td { padding:1.8mm 1.5mm; border-bottom:1px solid #ccc; }
  .month-table thead { display:table-header-group; }
  .month-table thead th { border-top:1px solid #777; border-bottom:1px solid #777; }
  .month-table tfoot { display:table-row-group; }
  .month-table tr,.month-kpis,.report-signatures { break-inside:avoid; }
  .month-kpis { display:flex; flex-wrap:wrap; gap:2mm 6mm; margin:0 0 4mm; padding:0; border:0; }
  .month-kpis > div { padding:0; border:0; flex:1 1 auto; }
  .month-kpis dt { font-size:9pt; }
  .month-kpis dd { font-size:10pt; margin:1mm 0 0; }
  .report-signatures { margin:12mm 0 0; gap:10mm; }
  .report-signatures span { flex:1 1 0; font-size:9pt; }
  .month-table thead th:nth-child(1) { width:12%; }
  .month-table thead th:nth-child(2) { width:7%; }
  .month-table thead th:nth-child(3) { width:23%; }
  .month-table thead th:nth-child(4) { width:30%; }
  .month-table thead th:nth-child(5) { width:18%; }
  .month-table thead th:nth-child(6) { width:10%; }
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
    <div><h1>&Aacute;lbuns em posse do vendedor</h1><p><%=Server.HTMLEncode(relPeriodo & "")%></p></div>
    <img class="month-logo" src="images/logoparafundopreto.png" alt="Stillos">
  </header>
<%If Len(relErro) > 0 Then%>
  <p class="alert alert-warning" role="alert" data-report-error><%=relErro%></p>
<%Else
    If posseRs.EOF Then
%>
  <p class="alert alert-info" role="status">Nenhum &aacute;lbum encontrado para este vendedor e filtro.</p>
<%Else%>
  <dl class="month-kpis">
    <div><dt>Representante comercial</dt><dd><%=Server.HTMLEncode(posseNome & "")%></dd></div>
    <div><dt>&Aacute;lbuns listados</dt><dd><%=posseTotal%></dd></div>
    <div><dt>M&eacute;dia de dias excedentes</dt><dd><%If IsNull(posseMedia) Then%>&mdash;<%Else%><%=FormatNumber(posseMedia, 1)%><%End If%></dd></div>
    <div><dt>Emiss&atilde;o</dt><dd><%=Server.HTMLEncode(Now() & "")%></dd></div>
  </dl>
  <div class="month-scroll" tabindex="0" role="group" aria-label="&Aacute;lbuns em posse">
    <table class="month-table">
      <thead><tr><th scope="col">C&oacute;d. formando</th><th scope="col">&Aacute;lbum</th><th scope="col" class="report-text">Formando</th><th scope="col" class="report-text">Institui&ccedil;&atilde;o de ensino</th><th scope="col">Data de sa&iacute;da</th><th scope="col">Dias excedentes</th></tr></thead>
      <tbody>
      <%Do Until posseRs.EOF%>
        <tr><th scope="row"><%=Server.HTMLEncode(digito(posseRs("COD_FORMANDO").Value) & "")%></th><td><%=Server.HTMLEncode(posseRs("ALBUM") & "")%></td><td class="report-text"><%=Server.HTMLEncode(posseRs("NOME_FORMANDO") & "")%></td><td class="report-text"><%=Server.HTMLEncode(posseRs("DESCRICAO") & "")%></td><td><%=Server.HTMLEncode(ArrumaData(posseRs("DATA_SAIDA").Value) & "")%></td><td><%=Server.HTMLEncode(posseRs("DIAS_ATRASADOS") & "")%></td></tr>
      <%  posseRs.MoveNext
        Loop%>
      </tbody>
    </table>
  </div>
  <div class="report-signatures"><span>Assinatura do emissor</span><span>Assinatura do representante comercial</span></div>
<%  End If
    posseRs.Close
    Set posseRs = Nothing
    FecharConexao()
End If%>
</main>
</body>
</html>
