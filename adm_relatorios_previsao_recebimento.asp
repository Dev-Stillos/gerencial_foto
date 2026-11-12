<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Validação especial para verificar se possui sub-validação específica
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '72' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se não encontrar então redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
		
	rs.Close
	Set rs=Nothing
	
FecharConexao()


Dim inicioMes, inicioAno, fimMes, fimAno, mesAtual, mesFinal, limite, filtroMes
Dim mensalSql, mensalRs, chaveMes, totalGeral, totaisMensais(3), valoresMensais(3), indiceMensal
Dim linhasMensais, percentualMensal, mensalRotulo, mensalErro, formaMensal

Set filtroMes = New RegExp
filtroMes.Pattern = "^[0-9]{1,2}$"
inicioMes = Trim(Request.QueryString("MES_INICIO"))
fimMes = Trim(Request.QueryString("MES_FIM"))
If Not filtroMes.Test(inicioMes) Or Not filtroMes.Test(fimMes) Then mensalErro = "Informe os meses inicial e final."
filtroMes.Pattern = "^[0-9]{4}$"
inicioAno = Trim(Request.QueryString("ANO_INICIO"))
fimAno = Trim(Request.QueryString("ANO_FIM"))
If Not filtroMes.Test(inicioAno) Or Not filtroMes.Test(fimAno) Then mensalErro = "Informe os anos inicial e final."
If Len(mensalErro) = 0 Then
    If CInt(inicioMes) < 1 Or CInt(inicioMes) > 12 Or CInt(fimMes) < 1 Or CInt(fimMes) > 12 _
       Or CInt(inicioAno) < 1000 Or CInt(fimAno) < 1000 Or CInt(fimAno) > 9998 Then
        mensalErro = "Informe meses e anos v&aacute;lidos."
    Else
        mesAtual = DateSerial(CInt(inicioAno), CInt(inicioMes), 1)
        mesFinal = DateSerial(CInt(fimAno), CInt(fimMes), 1)
        If mesAtual > mesFinal Then mensalErro = "O m&ecirc;s inicial deve ser anterior ou igual ao m&ecirc;s final."
    End If
End If

If Len(mensalErro) = 0 Then
    mensalRotulo = MonthName(Month(mesAtual)) & " de " & Year(mesAtual) & " a " & MonthName(Month(mesFinal)) & " de " & Year(mesFinal)
    limite = DateAdd("m", 1, mesFinal)
    mensalSql = "SELECT YEAR(P.DATA_VENDA)*100+MONTH(P.DATA_VENDA) AS MES_CHAVE, " & _
        "SUM(IFNULL(V.XTOTAL_CARTAO,0)) AS CARTAO, SUM(IFNULL(V.XTOTAL_CHEQUE,0)) AS CHEQUE, " & _
        "SUM(IFNULL(V.PAGAMENTO_EM_DINHEIRO,0)) AS DINHEIRO, SUM(IFNULL(V.XTOTAL,0)) AS TOTAL " & _
        "FROM VW_PEDIDOS_TOTAL_RECEBIDO V INNER JOIN TB_PEDIDOS P ON P.COD_PEDIDO=V.COD_PEDIDO " & _
        "WHERE P.FINALIZA_CORRECAO=1 AND P.DATA_VENDA >= " & sqlData(mesAtual) & _
        " AND P.DATA_VENDA < " & sqlData(limite) & " GROUP BY MES_CHAVE ORDER BY MES_CHAVE"
    AbrirConexao()
    Set mensalRs = objConexao.Execute(mensalSql)
    For indiceMensal = 0 To 3
        totaisMensais(indiceMensal) = CCur(0)
    Next
    Do While mesAtual <= mesFinal
        chaveMes = Year(mesAtual)*100+Month(mesAtual)
        For indiceMensal = 0 To 3
            valoresMensais(indiceMensal) = CCur(0)
        Next
        If Not mensalRs.EOF Then
            If CLng(mensalRs("MES_CHAVE")) = chaveMes Then
                For indiceMensal = 0 To 3
                    valoresMensais(indiceMensal) = CCur(mensalRs.Fields(indiceMensal+1).Value)
                Next
                mensalRs.MoveNext
            End If
        End If
        linhasMensais = linhasMensais & "<tr><th scope=""row"">" & Server.HTMLEncode(MonthName(Month(mesAtual))) & " / " & Year(mesAtual) & "</th>"
        For indiceMensal = 0 To 3
            totaisMensais(indiceMensal) = totaisMensais(indiceMensal) + valoresMensais(indiceMensal)
            linhasMensais = linhasMensais & "<td>" & FormatNumber(valoresMensais(indiceMensal), 2) & "</td>"
        Next
        linhasMensais = linhasMensais & "</tr>"
        mesAtual = DateAdd("m", 1, mesAtual)
    Loop
    mensalRs.Close
    Set mensalRs = Nothing
    FecharConexao()
    totalGeral = totaisMensais(3)
End If
formaMensal = Array("Cart&atilde;o de cr&eacute;dito", "Cheque", "Dinheiro", "Total de vendas")
%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Vendas mensais por forma de pagamento | Portal Stillos</title>
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
@page { size:A4 portrait; margin:12mm; }
@media print {
  body { background:#fff; font-size:10pt; }
  .month-report { margin:0; padding:0; border:0; }
  .month-tools { display:none; }
  .month-logo { width:32mm; padding:0; background:transparent; filter:brightness(0); }
  .month-scroll { overflow:visible; }
  .month-table { min-width:0; }
  .month-table thead { display:table-header-group; }
  .month-table tfoot { display:table-row-group; }
  .month-table tr,.month-kpis { break-inside:avoid; }
  .month-header,h2 { break-after:avoid; }
}
</style>
</head>
<body<%If Request.QueryString("modal") = "1" Then%> class="report-embedded"<%End If%>>
<main class="month-report" data-report-preview-content>
  <%If Request.QueryString("modal") <> "1" Then%>
  <nav class="month-tools" aria-label="A&ccedil;&otilde;es do relat&oacute;rio">
    <a class="btn btn-outline-secondary" href="adm_relatorios.asp">Voltar aos relat&oacute;rios</a>
    <%If Len(mensalErro) = 0 Then%><button type="button" class="btn btn-primary" onclick="window.print()">Imprimir / Salvar PDF</button><%End If%>
  </nav>
<%End If%>
  <header class="month-header">
    <div><h1>Vendas mensais por forma de pagamento</h1><p><%=Server.HTMLEncode(mensalRotulo & "")%></p></div>
    <img class="month-logo" src="images/logoparafundopreto.png" alt="Stillos">
  </header>
  <%If Len(mensalErro) > 0 Then%>
  <p class="alert alert-warning" role="alert" data-report-error><%=mensalErro%></p>
  <%Else%>
  <h2>Vis&atilde;o geral</h2>
  <dl class="month-kpis">
    <%For indiceMensal = 0 To 3%>
    <div><dt><%=formaMensal(indiceMensal)%></dt><dd><%=FormatCurrency(totaisMensais(indiceMensal), 2)%></dd>
      <%If totalGeral > 0 And indiceMensal < 3 Then
          percentualMensal = 100 * totaisMensais(indiceMensal) / totalGeral
      %><small><%=FormatNumber(percentualMensal, 1)%>% do total de vendas</small><%End If%>
    </div>
    <%Next%>
  </dl>
  <section aria-labelledby="month-results">
    <h2 id="month-results">Totais por m&ecirc;s</h2>
    <div class="month-scroll" tabindex="0" role="group" aria-labelledby="month-results">
      <table class="month-table">
        <thead><tr><th scope="col">M&ecirc;s / ano</th><%For indiceMensal = 0 To 3%><th scope="col"><%=formaMensal(indiceMensal)%></th><%Next%></tr></thead>
        <tbody><%=linhasMensais%></tbody>
        <tfoot><tr><th scope="row">Total do per&iacute;odo</th><%For indiceMensal = 0 To 3%><td><strong><%=FormatNumber(totaisMensais(indiceMensal), 2)%></strong></td><%Next%></tr></tfoot>
      </table>
    </div>
    <p class="month-note">Os valores s&atilde;o agrupados pela data da venda, independentemente do vencimento das parcelas.</p>
  </section>
  <%End If%>
</main>
</body>
</html>
