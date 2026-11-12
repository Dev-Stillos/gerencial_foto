<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/classe_fechamento.asp" -->
<%
Dim COD_FECHAMENTO_PERIODO, capPeriodPattern
COD_FECHAMENTO_PERIODO	= Request.QueryString("COD_FECHAMENTO_PERIODO")


Set capPeriodPattern = New RegExp
capPeriodPattern.Pattern = "^[1-9][0-9]{0,9}$"
If Not capPeriodPattern.Test(COD_FECHAMENTO_PERIODO) Then
    Response.Status = "400 Bad Request"
    Response.Write "Selecione um per&iacute;odo v&aacute;lido."
    Response.End
End If
Set capPeriodPattern = Nothing

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

If Len(data_pagamento & "") = 0 Then
    Response.Status = "404 Not Found"
    Response.Write "Per&iacute;odo n&atilde;o encontrado."
    Response.End
End If

'////////////////////////////////////////////////////////////////////////////////////


%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Previs&atilde;o da folha de pagamento</title>
<style>
* { box-sizing:border-box; }
body { margin:0; background:#f6f7fb; color:#242738; font:14px/1.45 system-ui,Arial,sans-serif; }
.cap-document { max-width:1150px; margin:16px auto; padding:clamp(12px,3vw,28px); border:1px solid #e4dfeb; border-radius:12px; background:#fff; }
.cap-doc-toolbar { text-align:right; margin:0 0 12px; }
.cap-doc-toolbar button { padding:8px 12px; border:0; border-radius:7px; background:#b20b61; color:white; font:inherit; cursor:pointer; }
.forecast-header { display:flex; flex-wrap:wrap; align-items:center; justify-content:space-between; gap:12px; padding-bottom:12px; border-bottom:2px solid #b20b61; }
.forecast-header h1 { margin:0; font-size:clamp(18px,3vw,26px); }
.forecast-logo { width:clamp(100px,20vw,150px); height:auto; border-radius:7px; background:#292b40; }
.forecast-status { display:inline-block; margin:7px 0 0; padding:3px 8px; background:#f9e7ed; color:#86234a; border-radius:5px; font-weight:650; }
.forecast-period { display:flex; flex-wrap:wrap; gap:10px 26px; margin:14px 0; padding:10px 14px; border-radius:8px; background:#f5f2f8; }
.forecast-period dt { color:#514960; font-size:.88em; }
.forecast-period dd { margin:2px 0 0; font-weight:700; }
.forecast-scroll { overflow-x:auto; border:1px solid #e4dfeb; border-radius:8px; }
.forecast-scroll:focus-visible { outline:2px solid #b20b61; }
.forecast-table { width:100%; min-width:570px; border-collapse:collapse; font-variant-numeric:tabular-nums; }
.forecast-table th, .forecast-table td { padding:7px 9px; border-bottom:1px solid #e4dfeb; text-align:right; white-space:nowrap; }
.forecast-table tr > :first-child { text-align:left; white-space:normal; }
.forecast-table thead, .forecast-table tfoot { background:#f0ebf5; }
.forecast-table tbody tr:nth-child(even) { background:#faf9fc; }
.forecast-table tr > :last-child { font-weight:700; }
.forecast-table tfoot td { font-size:1.08em; }
.forecast-note { margin:12px 0 0; color:#514960; }
@page { size:A4 landscape; margin:10mm; }
@media print {
  body { color:#111; background:#fff; font-size:9pt; line-height:1.25; }
  .cap-document { margin:0; max-width:none; padding:3mm 4mm; border:0; }
  .cap-doc-toolbar { display:none; }
  .forecast-header { padding-bottom:2mm; border-color:#444; break-inside:avoid; }
  .forecast-header h1 { font-size:14pt; }
  .forecast-logo { width:30mm; background:transparent; filter:brightness(0); }
  .forecast-period { margin:2mm 0; padding:1.5mm 2mm; break-inside:avoid; }
  .forecast-scroll { overflow:visible; }
  .forecast-table { min-width:0; }
  .forecast-table th, .forecast-table td { padding:1.2mm 2mm; }
  .forecast-table thead { display:table-header-group; }
  .forecast-table tfoot { display:table-row-group; }
  .forecast-table tr, .forecast-note { break-inside:avoid; }
}
</style>
</head>
<body>
<main class="cap-document" data-cap-report>
<% If Request.QueryString("modal") <> "1" Then %>
<div class="cap-doc-toolbar"><button type="button" onclick="window.print()">Imprimir / Salvar PDF</button></div>
<% End If %>
<header class="forecast-header">
  <div><h1>Previs&atilde;o da folha de pagamento</h1><p class="forecast-status">Provis&oacute;rio &middot; Fechamento n&atilde;o finalizado</p></div>
  <img class="forecast-logo" src="images/logoparafundopreto.png" alt="Stillos">
</header>
<dl class="forecast-period">
  <div><dt>Data de pagamento</dt><dd><%=data_pagamento%></dd></div>
  <div><dt>Per&iacute;odo de vendas</dt><dd><%=data_inicio%> a <%=data_termino%></dd></div>
</dl>
<div class="forecast-scroll" tabindex="0" role="region" aria-label="Valores por consultor">
<table class="forecast-table">
<thead><tr>
  <th scope="col">Consultor</th>
  <th scope="col">Comiss&atilde;o</th>
  <th scope="col">Ajuda de custo</th>
  <th scope="col">Pend&ecirc;ncias</th>
  <th scope="col">Total a receber</th>
</tr></thead>
<tbody>
<%
AbrirConexao()


Dim rsTotos
Dim NOME_VENDEDOR, COD_VENDEDOR
Dim TOTAL_GERAL_FOLHA_PAGAMENTO, TOTAL_A_RECEBER_POR_CONSULTOR
Dim TOTAL_COMISSAO, TOTAL_AJUDA_CUSTO, TOTAL_PENDENCIA_PAGAMENTO
TOTAL_A_RECEBER_POR_CONSULTOR = 0
TOTAL_GERAL_FOLHA_PAGAMENTO = 0

			'Obs.: Nos casos dos vendedores coordenadores 7 e 14 os mesmos dever�o aparecer mesmo que não tenham venda
				
Dim filtroExVendedores
filtroExVendedores = ""
If Len(idsExVendedoresFechamento) > 0 Then
    filtroExVendedores = " AND TB_PEDIDOS.COD_VENDEDOR NOT IN (" & idsExVendedoresFechamento & ")"
End If

			sql = "SELECT TB_PEDIDOS.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR FROM TB_VENDEDORES LEFT JOIN TB_PEDIDOS ON (TB_VENDEDORES.COD_VENDEDOR = TB_PEDIDOS.COD_VENDEDOR) WHERE TB_PEDIDOS.finaliza_correcao = 1 AND  TB_PEDIDOS.cod_fechamento_periodo = "&cod_fechamento_periodo&" AND TB_VENDEDORES.COD_VENDEDOR <> 27" & filtroExVendedores & " GROUP BY TB_PEDIDOS.COD_VENDEDOR ORDER BY TB_VENDEDORES.NOME_VENDEDOR "	
																
			'Response.Write(sql)

			Set rsTodos = objConexao.Execute(sql)

			While not rsTodos.Eof
			
				
			
				NOME_VENDEDOR				= Ucase(rsTodos("NOME_VENDEDOR"))
				COD_VENDEDOR 				= rsTodos("COD_VENDEDOR")
				
				'Response.Write(NOME_VENDEDOR)
				
				TOTAL_COMISSAO = 0
				TOTAL_AJUDA_CUSTO = 0
				TOTAL_PENDENCIA_PAGAMENTO = 0
				
				'1) Total de comiss�o por pedido
				Set rs2=objConexao.Execute("SELECT COD_PEDIDO FROM TB_PEDIDOS WHERE cod_fechamento_periodo = "&cod_fechamento_periodo&" AND COD_VENDEDOR = "&COD_VENDEDOR&" AND FINALIZA_CORRECAO = 1")
				
					While not rs2.Eof
					
							TOTAL_COMISSAO = (TOTAL_COMISSAO+fechamentoRetornaValorComissaoPedido(rs2("cod_pedido"),1,true))/1
							
					rs2.MoveNext()
					Wend
				
				
				rs2.Close
				Set rs2=Nothing
				
				
				'2) Total de Ajuda de Custo
				sql = "SELECT SUM(VALOR_AJUDA_CUSTO) AS TOTAL_AJUDA_CUSTO FROM  TB_PEDIDOS WHERE cod_fechamento_periodo = "&cod_fechamento_periodo&" AND COD_VENDEDOR = "&COD_VENDEDOR&" AND FINALIZA_CORRECAO = 1"
				
				Set rs2=objConexao.Execute(sql)
				
				
					If not rs2.Eof Then
						TOTAL_AJUDA_CUSTO = formatNumber(0&rs2("TOTAL_AJUDA_CUSTO"))
					End If
								
				
				rs2.Close
				Set rs2=Nothing
				
				'3) Pend�ncia pagamento
				TOTAL_PENDENCIA_PAGAMENTO = (fechamentoRetornaValorRecebidoVendedorTipo(5,cod_vendedor,cod_fechamento_periodo,true))/1
				
				TOTAL_A_RECEBER_POR_CONSULTOR = (TOTAL_COMISSAO + TOTAL_AJUDA_CUSTO + TOTAL_PENDENCIA_PAGAMENTO)/1
				
				TOTAL_GERAL_FOLHA_PAGAMENTO = (TOTAL_GERAL_FOLHA_PAGAMENTO+TOTAL_A_RECEBER_POR_CONSULTOR)/1
				
				
				
%>
  <tr>
    <th scope="row"><%=Server.HTMLEncode(NOME_VENDEDOR & "")%></th>
    <td><%=formatNumber(TOTAL_COMISSAO)%></td>
    <td><%=formatNumber(TOTAL_AJUDA_CUSTO)%></td>
    <td><%=formatNumber(TOTAL_PENDENCIA_PAGAMENTO)%></td>
    <td><%=formatNumber(TOTAL_A_RECEBER_POR_CONSULTOR)%></td>
  </tr>
<%

			rsTodos.MoveNext()
			Wend
			
			rsTodos.Close
			Set rsTodos = Nothing
			FecharConexao()
%>
</tbody>
<tfoot><tr>
  <th colspan="4" scope="row">Total previsto</th>
  <td><%=formatCurrency(TOTAL_GERAL_FOLHA_PAGAMENTO)%></td>
</tr></tfoot>
</table>
</div>
<p class="forecast-note">A previs&atilde;o n&atilde;o inclui os totais de coordena&ccedil;&atilde;o e de metas. Vendas conferidas pelo departamento administrativo fotogr&aacute;fico.</p>
</main>
</body>
</html>
