<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Validação especial para verificar se possui sub-validação específica
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '13' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se não encontrar então redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
		
	rs.Close
	Set rs=Nothing
	
FecharConexao()

Function DataFiltroDiretoria(valor)
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
    If Month(dataValidada) = mes And Day(dataValidada) = dia Then DataFiltroDiretoria = dataValidada
End Function

Dim relErro, relPeriodo, fichasInicio, fichasFim, fichasFiltro, fichasLimite, fichasSql, fichasRs, fichasTotal
If Len(Request.QueryString("filtro1")) > 0 Then
    fichasLimite = " LIMIT 100"
ElseIf Len(Request.QueryString("filtro2")) > 0 Then
    fichasInicio = DataFiltroDiretoria(Trim(Request.QueryString("DATA_INI")))
    fichasFim = DataFiltroDiretoria(Trim(Request.QueryString("DATA_FIM")))
    If IsEmpty(fichasInicio) Or IsEmpty(fichasFim) Then
        relErro = "Informe datas v&aacute;lidas para consultar o relat&oacute;rio."
    ElseIf fichasInicio > fichasFim Then
        relErro = "A data inicial deve ser anterior ou igual &agrave; data final."
    Else
        relPeriodo = FormatDateTime(fichasInicio, 2) & " a " & FormatDateTime(fichasFim, 2)
        fichasFiltro = " WHERE DATA_CORRIGIDO_DIRETORIA >= " & sqlData(fichasInicio) & _
            " AND DATA_CORRIGIDO_DIRETORIA < " & sqlData(DateAdd("d", 1, fichasFim))
    End If
Else
    relErro = "Escolha os &uacute;ltimos 100 registros ou informe um per&iacute;odo."
End If

If Len(relErro) = 0 Then
    AbrirConexao()
    fichasSql = "SELECT DATA_CORRIGIDO_DIRETORIA, COD_PEDIDO, NOME_VENDEDOR " & _
        "FROM VW_PEDIDOS_AUTENTICADOS_DIRETORIA" & fichasFiltro & _
        " ORDER BY DATA_CORRIGIDO_DIRETORIA DESC, COD_PEDIDO DESC" & fichasLimite
    Set fichasRs = objConexao.Execute(fichasSql)
End If
%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Fichas autenticadas pela Diretoria | Portal Stillos</title>
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
.fichas-filter { display:grid; gap:5px; max-width:420px; margin-bottom:14px; }
.fichas-filter select { min-width:0; width:100%; text-overflow:ellipsis; }
.fichas-selection:empty { display:none; }
.fichas-selection { margin:0 0 14px; overflow-wrap:anywhere; }
.fichas-panels { display:grid; grid-template-columns:minmax(0,3fr) minmax(0,2fr); gap:12px; margin:0 0 18px; }
.fichas-panel { min-width:0; border:1px solid #e4dfeb; border-radius:10px; padding:14px; }
.fichas-bars { padding:0; margin:10px 0 0; list-style:none; }
.fichas-bars li + li { margin-top:8px; }
.fichas-bar-label { display:flex; justify-content:space-between; gap:8px; overflow-wrap:anywhere; }
.fichas-track { display:block; height:7px; margin-top:3px; border-radius:5px; background:#eee9f4; overflow:hidden; }
.fichas-track span { display:block; height:100%; background:#a31962; }
.fichas-panel meter { display:block; width:100%; height:20px; accent-color:#a31962; }
.fichas-panel meter::-webkit-meter-optimum-value { background:#a31962; }
.fichas-print-scope { display:none; }
@media(max-width:700px) { .fichas-panels { grid-template-columns:minmax(0,1fr); } }
@page { size:A4 portrait; margin:10mm; }
@media print {
  body { margin:0; background:#fff; color:#111; font-size:10pt; line-height:1.3; }
  .month-report { margin:0; padding:6mm 7mm; max-width:none; border:0; border-radius:0; }
  .month-tools,.fichas-filter,.fichas-selection,.fichas-panels { display:none; }
  .fichas-print-scope:not(:empty) { display:block; margin:0 0 3mm; }
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
  .month-kpis { display:flex; flex-wrap:wrap; gap:2mm 6mm; margin:0 0 4mm; }
  .month-kpis > div { padding:0; border:0; flex:1 1 auto; }
  .month-kpis dt { font-size:9pt; }
  .month-kpis dd { font-size:10pt; margin:1mm 0 0; }
  .report-signatures { margin:12mm 0 0; gap:10mm; }
  .report-signatures span { flex:1 1 0; font-size:9pt; }
  .month-table thead th:nth-child(1) { width:34%; }
  .month-table thead th:nth-child(2) { width:16%; }
  .month-table thead th:nth-child(3) { width:50%; }
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
    <div><h1>Fichas autenticadas pela Diretoria</h1><p><%If Len(fichasLimite) > 0 Then%>&Uacute;ltimos 100 registros<%Else%><%=Server.HTMLEncode(relPeriodo & "")%><%End If%></p></div>
    <img class="month-logo" src="images/logoparafundopreto.png" alt="Stillos">
  </header>
<%If Len(relErro) > 0 Then%>
  <p class="alert alert-warning" role="alert" data-report-error><%=relErro%></p>
<%Else
    fichasTotal = 0
    If fichasRs.EOF Then
%>
  <p class="alert alert-info" role="status">Nenhuma ficha autenticada no per&iacute;odo consultado.</p>
<%Else%>
  <section id="fichas-dashboard" hidden aria-labelledby="fichas-overview-title">
    <h2 id="fichas-overview-title">Vis&atilde;o geral</h2>
    <label class="fichas-filter" for="fichas-seller">Vendedor
      <select class="form-select" id="fichas-seller" aria-describedby="fichas-selection"><option value="">Todos os vendedores</option></select>
    </label>
    <p class="fichas-selection" id="fichas-selection" role="status"></p>
    <p class="fichas-print-scope" id="fichas-print-scope"></p>
    <dl class="month-kpis">
      <div><dt>Fichas exibidas</dt><dd id="fichas-count"></dd></div>
      <div><dt>Vendedores na sele&ccedil;&atilde;o</dt><dd id="fichas-sellers"></dd></div>
    </dl>
    <div class="fichas-panels">
      <section class="fichas-panel"><h3 id="fichas-bars-title">Fichas por vendedor</h3><ul class="fichas-bars" id="fichas-bars"></ul></section>
      <section class="fichas-panel">
        <h3>Participa&ccedil;&atilde;o na consulta</h3>
        <p id="fichas-share"></p>
        <meter id="fichas-meter" min="0" max="100" value="0" aria-labelledby="fichas-share"></meter>
      </section>
    </div>
  </section>
  <div class="month-scroll" tabindex="0" role="group" aria-label="Fichas autenticadas">
    <table class="month-table">
      <thead><tr><th scope="col">Data e hora da autentica&ccedil;&atilde;o</th><th scope="col">Pedido</th><th scope="col" class="report-text">Vendedor</th></tr></thead>
      <tbody>
      <%Do Until fichasRs.EOF
          fichasTotal = fichasTotal + 1
      %>
        <tr><th scope="row"><%=Server.HTMLEncode(fichasRs("DATA_CORRIGIDO_DIRETORIA") & "")%></th><td><%=Server.HTMLEncode(fichasRs("COD_PEDIDO") & "")%></td><td class="report-text"><%=Server.HTMLEncode(fichasRs("NOME_VENDEDOR") & "")%></td></tr>
      <%  fichasRs.MoveNext
        Loop%>
      </tbody>
      <tfoot><tr><th scope="row">Total de registros exibidos</th><td colspan="2"><strong id="fichas-table-count"><%=fichasTotal%></strong></td></tr></tfoot>
    </table>
  </div>
<%  End If
    fichasRs.Close
    Set fichasRs = Nothing
    FecharConexao()
End If%>
</main>
<script>
(function () {
    var dashboard = document.getElementById('fichas-dashboard');
    if (!dashboard) return;

    var select = document.getElementById('fichas-seller');
    var rows = Array.from(document.querySelectorAll('.month-table tbody tr'));
    var counts = new Map();
    var unknown = 'Vendedor n\u00e3o informado';
    rows.forEach(function (row) {
        var name = row.cells[2].textContent.trim() || unknown;
        counts.set(name, (counts.get(name) || 0) + 1);
    });
    var names = Array.from(counts.keys()).sort(function (a, b) { return a.localeCompare(b, 'pt-BR'); });
    names.forEach(function (name) { select.add(new Option(name, name)); });
    var ranked = names.slice().sort(function (a, b) { return counts.get(b) - counts.get(a); });
    var format = new Intl.NumberFormat('pt-BR');
    var percent = new Intl.NumberFormat('pt-BR', { style:'percent', maximumFractionDigits:1 });

    function render() {
        var seller = select.value;
        var total = seller ? counts.get(seller) : rows.length;
        rows.forEach(function (row) {
            row.hidden = Boolean(seller && (row.cells[2].textContent.trim() || unknown) !== seller);
        });
        document.getElementById('fichas-count').textContent = format.format(total);
        document.getElementById('fichas-sellers').textContent = format.format(seller ? 1 : names.length);
        document.getElementById('fichas-table-count').textContent = format.format(total);
        document.getElementById('fichas-selection').textContent = seller;
        document.getElementById('fichas-print-scope').textContent = seller ? 'Vendedor: ' + seller : '';

        var top = seller ? [seller] : ranked.slice(0, 5);
        var maximum = counts.get(ranked[0]);
        var bars = document.getElementById('fichas-bars');
        document.getElementById('fichas-bars-title').textContent = seller ? 'Fichas do vendedor' : 'Vendedores com mais fichas';
        bars.replaceChildren();
        top.forEach(function (name) {
            var item = document.createElement('li');
            var label = document.createElement('div');
            var count = document.createElement('strong');
            var track = document.createElement('span');
            var bar = document.createElement('span');
            label.className = 'fichas-bar-label';
            count.textContent = format.format(counts.get(name));
            label.append(name, count);
            track.className = 'fichas-track';
            track.setAttribute('aria-hidden', 'true');
            bar.style.width = (100 * counts.get(name) / maximum) + '%';
            track.append(bar);
            item.append(label, track);
            bars.append(item);
        });
        var share = counts.get(top[0]) / rows.length;
        document.getElementById('fichas-share').textContent = percent.format(share) + ' das fichas: ' + top[0];
        document.getElementById('fichas-meter').value = share * 100;
    }

    select.addEventListener('change', render);
    dashboard.hidden = false;
    render();
})();
</script>
</body>
</html>
