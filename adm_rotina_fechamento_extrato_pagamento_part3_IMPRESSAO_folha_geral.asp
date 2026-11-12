<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/classe_fechamento.asp" -->
<%
Dim COD_FECHAMENTO_PERIODO
COD_FECHAMENTO_PERIODO	= Request.QueryString("COD_FECHAMENTO_PERIODO")


'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
Dim data_pagamento, data_inicio, data_termino
Dim idsExVendedoresFechamento
idsExVendedoresFechamento = ""

AbrirConexao()

	Set rs=objConexao.Execute("Select *, CASE WHEN DATA_TERMINO >= '2026-11-15' THEN 'S' ELSE 'N' END AS ocultar_ex_vendedores FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE cod_fechamento_periodo = "&cod_fechamento_periodo)
		If not rs.Eof Then
			data_pagamento	= ArrumaData(rs("data_pagamento"))
			If CStr(rs.Fields("ocultar_ex_vendedores").Value & "") = "S" Then idsExVendedoresFechamento = "7,10,14,48,228,255,276,312,335,341,353,363,369,371,372"
			data_inicio 	= ArrumaData(rs("data_inicio"))
			data_termino	= ArrumaData(rs("data_termino"))
		End If
	rs.Close
	Set rs=Nothing

FecharConexao()

'////////////////////////////////////////////////////////////////////////////////////


%>
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Folha Geral</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css" media="all">
<!--
.style1 {
	font-size: 18px;
	font-weight: bold;
}
.folha {
    page-break-after: always;

}
body,td,th {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 18px;
}
.style2 {
	font-size: 27px;
	font-weight: bold;
}
-->
</style>
<style type="text/css">
<!--
.style2 {font-weight: bold}
.style8 {font-size: 13px; font-weight: bold; }
.style9 {font-size: 13px}
-->
</style>
<style>
* { box-sizing:border-box; }
body { margin:0; background:#f6f7fb; color:#242738; font:14px/1.45 system-ui,Arial,sans-serif; }
.cap-document { max-width:1400px; margin:14px auto; padding:clamp(12px,3vw,26px); border:1px solid #e4dfeb; border-radius:12px; background:#fff; }
.cap-doc-header { display:flex; align-items:center; justify-content:space-between; gap:12px; padding-bottom:12px; margin-bottom:16px; border-bottom:2px solid #b20b61; }
.cap-doc-header h1 { margin:0; font:700 clamp(18px,3vw,27px)/1.2 system-ui,sans-serif; }
.cap-doc-logo { flex:none; width:clamp(95px,20vw,150px); height:auto; background:#292b40; border-radius:7px; }
.cap-doc-toolbar { display:flex; justify-content:flex-end; margin-bottom:10px; }
.cap-doc-button { display:inline-block; padding:8px 12px; border:1px solid #d8d0e1; border-radius:7px; background:#fff; color:#88134f; font:600 13px/1.3 system-ui,sans-serif; cursor:pointer; }
.cap-doc-primary { background:#b20b61; color:#fff; border-color:#b20b61; }
.cap-doc-scroll { overflow-x:auto; }
.cap-doc-scroll table { border-collapse:collapse; border:0; }
.cap-doc-scroll th, .cap-doc-scroll td { padding:6px 5px; border-color:#e4dfeb; font-size:inherit; }
.cap-doc-scroll th { background:#f3f0f7; color:#242738; }
.cap-doc-scroll .folha { width:100%; margin:12px 0; border:1px solid #e4dfeb; border-radius:8px; }
.cap-doc-scroll .folha td { padding:14px; }
.cap-doc-scroll .folha p { margin:8px 0; }
.cap-boleto-filter { display:flex; flex-wrap:wrap; gap:10px; align-items:flex-end; margin:0 0 14px; padding:12px; background:#f7f5fa; border-radius:8px; }
.cap-boleto-filter label { display:block; font-weight:600; }
.cap-boleto-filter select, .cap-boleto-edit input[type=text] { display:block; max-width:100%; margin-top:4px; padding:7px; border:1px solid #d8d0e1; border-radius:6px; background:#fff; color:inherit; }
.cap-boleto-table { width:100%; font-size:12px; }
.cap-boleto-table tbody tr:nth-child(even) { background:#faf9fc; }
.cap-boleto-status { display:inline-block; padding:3px 6px; border-radius:5px; white-space:nowrap; background:#f9e7ed; color:#86234a; }
.cap-boleto-status.is-paid { background:#e5f2ea; color:#275e42; }
.cap-boleto-edit { width:min(430px,calc(100% - 24px)); padding:20px; border:1px solid #e4dfeb; border-radius:12px; color:#242738; }
.cap-boleto-edit::backdrop { background:#24273899; }
.cap-boleto-edit h2 { margin-top:0; }
@media screen and (max-width:720px) {
  .cap-boleto-table, .cap-boleto-table tbody { display:block; }
  .cap-boleto-table thead { position:absolute; width:1px; height:1px; overflow:hidden; clip-path:inset(50%); }
  .cap-boleto-table tbody tr { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); margin-bottom:10px; border:1px solid #e4dfeb; border-radius:8px; }
  .cap-boleto-table td { display:block; text-align:left; overflow-wrap:anywhere; }
  .cap-boleto-table td::before { content:attr(data-label); display:block; font-weight:650; }
  .cap-boleto-table td[colspan] { grid-column:1/-1; }
}
@page { size:A4 portrait; margin:10mm; }
@media print {
  body { background:white; color:#111; font-size:9.5pt; }
  .cap-document { max-width:none; margin:0; padding:5mm 8mm; border:0; }
  .cap-doc-logo { width:32mm; background:transparent; filter:brightness(0); }
  .cap-doc-header h1 { font-size:15pt; }
  .cap-doc-toolbar, .cap-boleto-filter, .cap-boleto-edit, .cap-boleto-action { display:none !important; }
  .cap-doc-scroll { overflow:visible; }
  .cap-doc-scroll table { max-width:100%; }
  .cap-doc-scroll th, .cap-doc-scroll td { padding:4px 3px; }
  .cap-doc-scroll tr { break-inside:avoid; }
  .cap-doc-scroll thead { display:table-header-group; }
  .cap-doc-scroll .folha { break-after:page; }
  .cap-doc-scroll .folha:last-of-type { break-after:auto; }
}
</style>
</head>
<script language="javascript">


	function mClk2(src, url)
			{
				//if(event.srcElement.tagName=='TD')
				//	{
				//		src.children.tags('A')[0].click();
				//	}
			location.href = url;

			}

</script>
<body class="cap-report-page">
<main class="cap-document" data-cap-report>
<%If Request.QueryString("modal") <> "1" Then%><div class="cap-doc-toolbar"><button class="cap-doc-button cap-doc-primary" type="button" onclick="window.print()">Imprimir / Salvar PDF</button></div><%End If%>
<header class="cap-doc-header"><h1>Folha geral de pagamento</h1><img class="cap-doc-logo" src="images/logoparafundopreto.png" alt="Stillos"></header>
<div class="cap-doc-scroll">

<div align="center"><strong>Reportagens Fotogr&aacute;fica - Resumo Folha de Pagamento</strong><br>
<strong>Pagamento de <%=data_pagamento%> - Vendas de <%=data_inicio%> &agrave; <%=data_termino%></strong></div>

<table width="100%" border="0" cellspacing="0" cellpadding="3">
  <tr>
    <td width="25%"><div align="center" class="style8">Consultor</div></td>
    <td width="8%"><div align="center" class="style8">Comiss&atilde;o</div></td>
    <td width="12%"><div align="center" class="style8">Ajuda de Custo</div></td>
    <td width="10%"><div align="center" class="style8">Pend&ecirc;ncias</div></td>
    <td width="10%"><div align="center" class="style8">Coordenador</div></td>
    <td width="12%"><div align="center" class="style8">Metas</div></td>
    <td width="23%"><div align="center" class="style8">Total a Receber</div></td>
  </tr>

<%
AbrirConexao()


Dim rsTotos
Dim NOME_VENDEDOR, COD_VENDEDOR
Dim TOTAL_GERAL_FOLHA_PAGAMENTO, TOTAL_A_RECEBER_POR_CONSULTOR
TOTAL_A_RECEBER_POR_CONSULTOR = 0
TOTAL_GERAL_FOLHA_PAGAMENTO = 0

			'Obs.: Nos casos dos vendedores coordenadores 7 e 14 os mesmos dever�o aparecer mesmo que n�o tenham venda

			sql = "SELECT "&_
		"TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR ,"&_
		"TB_FECHAMENTO_PAGAMENTO.COD_GRUPO_VENDEDORES,"&_
		"TB_VENDEDORES.NOME_VENDEDOR "&_
		"FROM TB_VENDEDORES INNER JOIN TB_FECHAMENTO_PAGAMENTO ON TB_VENDEDORES.COD_VENDEDOR = TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR "&_
		"WHERE TB_FECHAMENTO_PAGAMENTO.COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR <> 27 OR TB_VENDEDORES.COD_VENDEDOR IN (SELECT cod_vendedor FROM TB_FECHAMENTO_PENDENCIAS_PAGAMENTO WHERE cod_fechamento_periodo = '"&COD_FECHAMENTO_PERIODO&"') or TB_VENDEDORES.COD_VENDEDOR IN(7,14,48) "&_
		"GROUP BY TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR ORDER BY NOME_VENDEDOR"



Dim filtroExVendedores
filtroExVendedores = ""
If Len(idsExVendedoresFechamento) > 0 Then
    filtroExVendedores = " WHERE teste.COD_VENDEDOR NOT IN (" & idsExVendedoresFechamento & ")"
End If
	sql ="SELECT * FROM (SELECT "&_
	"TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR ,"&_
	"TB_VENDEDORES.NOME_VENDEDOR "&_
	"FROM TB_VENDEDORES INNER JOIN TB_FECHAMENTO_PAGAMENTO ON TB_VENDEDORES.COD_VENDEDOR = TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR "&_
	"WHERE (TB_FECHAMENTO_PAGAMENTO.COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR <> 27) OR TB_VENDEDORES.COD_VENDEDOR IN (SELECT cod_vendedor FROM TB_FECHAMENTO_PENDENCIAS_PAGAMENTO WHERE cod_fechamento_periodo = '"&COD_FECHAMENTO_PERIODO&"') or TB_VENDEDORES.COD_VENDEDOR IN(7,14,48) "&_
	"GROUP BY TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR UNION ALL "&_
	"SELECT TB_FECHAMENTO_PAGAMENTO_BOLETOS.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR FROM TB_VENDEDORES INNER JOIN TB_FECHAMENTO_PAGAMENTO_BOLETOS ON TB_VENDEDORES.COD_VENDEDOR = TB_FECHAMENTO_PAGAMENTO_BOLETOS.COD_VENDEDOR WHERE (TB_FECHAMENTO_PAGAMENTO_BOLETOS.COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&") GROUP BY TB_FECHAMENTO_PAGAMENTO_BOLETOS.COD_VENDEDOR ORDER BY NOME_VENDEDOR) AS teste" & filtroExVendedores & " GROUP BY teste.NOME_VENDEDOR"




			'Response.Write(sql)

			Set rsTodos = objConexao.Execute(sql)




			While not rsTodos.Eof

				NOME_VENDEDOR				= Ucase(rsTodos("NOME_VENDEDOR"))
				COD_VENDEDOR 				= rsTodos("COD_VENDEDOR")


TOTAL_A_RECEBER_POR_CONSULTOR = formatCurrency(fechamentoRetornaValorRecebidoVendedor(COD_VENDEDOR,cod_fechamento_periodo,true))
TOTAL_GERAL_FOLHA_PAGAMENTO = TOTAL_GERAL_FOLHA_PAGAMENTO + TOTAL_A_RECEBER_POR_CONSULTOR
%>
  <tr>
    <td class="style9" style="border-bottom:solid 1px #999999"><%=NOME_VENDEDOR%></td>
<td class="style9" style="border-bottom:solid 1px #999999"><div align="right"><%=formatNumber("" & fechamentoRetornaValorRecebidoVendedorTipo(1,cod_vendedor,cod_fechamento_periodo,true))%></div></td>    
<td class="style9" style="border-bottom:solid 1px #999999"><div align="right"><%=formatNumber(0&fechamentoRetornaValorRecebidoVendedorTipo(2,cod_vendedor,cod_fechamento_periodo,true))%></div></td>
	<td class="style9" style="border-bottom:solid 1px #999999"><div align="right"><%=formatNumber(0&fechamentoRetornaValorRecebidoVendedorTipo(3,cod_vendedor,cod_fechamento_periodo,true))%></div></td>
    <td class="style9" style="border-bottom:solid 1px #999999"><div align="right"><%=formatNumber(0&fechamentoRetornaValorRecebidoVendedorTipo(4,cod_vendedor,cod_fechamento_periodo,true))%></div></td>
    <td class="style9" style="border-bottom:solid 1px #999999"><div align="right"><%=formatNumber(0&fechamentoRetornaValorRecebidoVendedorTipo(5,cod_vendedor,cod_fechamento_periodo,true))%></div></td>
    <td style="border-bottom:solid 1px #999999"><div align="right" class="style9"><%=formatNumber(TOTAL_A_RECEBER_POR_CONSULTOR)%></div></td>
  </tr>



<%

			rsTodos.MoveNext()
			Wend

			rsTodos.Close
			Set rsTodos = Nothing
			FecharConexao()
%>
  <tr>
    <td colspan="6" style="border-top:solid 1px #333333"></td>
    <td style="border-top:solid 1px #333333"><div align="right"><strong><%=formatCurrency(TOTAL_GERAL_FOLHA_PAGAMENTO)%></strong></div></td>
  </tr>
</table>
</div></main>
<%If Request.QueryString("modal") <> "1" Then%><script>window.addEventListener('load', function () { window.print(); });</script><%End If%>
</body>
</html>
