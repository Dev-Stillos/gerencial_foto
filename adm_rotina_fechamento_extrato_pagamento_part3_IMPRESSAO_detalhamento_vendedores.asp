<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/classe_fechamento.asp" -->
<%

Function pedidosFormandosFormatado(COD_PEDIDO, conexao_aberta)

	dim rsFormandos, retorno 

	if not conexao_aberta then
		AbrirConexao()
	end if

	Set rsFormandos = objConexao.Execute("SELECT COD_FORMANDO, NOME_FORMANDO ,ALBUM, DESCRICAO FROM VW_PEDIDOS_FORMANDOS WHERE COD_PEDIDO="&COD_PEDIDO)
	if not rsFormandos.eof then
		retorno = "<table width='100%' border='0' cellspacing='0' cellpadding='2'>"
		While not rsFormandos.Eof
			retorno = retorno & "<tr>"
			retorno = retorno & "<td width='20%'><div align='left'>"&digito(rsFormandos("COD_FORMANDO"))&"</div></td>"
			retorno = retorno & "<td width='3%'>/</td>"
			retorno = retorno & "<td width='8%'>"&rsFormandos("ALBUM")&"</td>"
			retorno = retorno & "<td width='3%'>/</td>"
			retorno = retorno & "<td width='66%'><div align='left'>"&rsFormandos("DESCRICAO")&"</div></td>"
			retorno = retorno & "</tr>"
		rsFormandos.MoveNext()
		Wend
		retorno = retorno & "</table>"
	end if
	rsFormandos.Close
	Set rsFormandos = Nothing

	if not conexao_aberta then
		fecharConexao()
	end if

pedidosFormandosFormatado = retorno
end Function

Function pedidosFormataFormaPagamento(tipo_pagamento)

	If int(FORMA_PAGAMENTO) = 1 Then
		pedidosFormataFormaPagamento = "&agrave; vista" 
	Elseif int(FORMA_PAGAMENTO) = 2 Then
		pedidosFormataFormaPagamento = "&agrave; prazo"
	End If

end Function



Dim COD_VENDEDOR, ESCOLHA_EXTRATO, COD_FECHAMENTO_PERIODO, NOME_VENDEDOR, TOTAL_BOLETO
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


Function retiraZero2(CAMPO)
	Dim temp
	temp = CAMPO
	If temp = 0 Then
		temp = " - "
	End If	
	
	retiraZero2 = temp
	
End Function

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="all">
<title>Extrato dos Vendedores</title>


<style type="text/css" media="all">
<!--
.folha2 {
    page-break-after: always;
}
-->
</style><style>
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
@page { size:A4 landscape; margin:10mm; }
@media print {
  body { background:white; color:#111; font-size:9.5pt; }
  .cap-document { max-width:none; margin:0; padding:3mm 5mm; border:0; }
  .cap-doc-logo { width:32mm; background:transparent; filter:brightness(0); }
  .cap-doc-header { padding-bottom:2mm; margin-bottom:3mm; }
  .cap-doc-header h1 { font-size:13pt; }
  .cap-doc-toolbar, .cap-boleto-filter, .cap-boleto-edit, .cap-boleto-action { display:none !important; }
  .cap-doc-scroll { overflow:visible; }
  .cap-doc-scroll table { max-width:100%; }
  .cap-doc-scroll th, .cap-doc-scroll td { padding:4px 3px; }
  .extract-table > tbody > tr, .extract-total { break-inside:avoid; }
  .cap-doc-scroll thead { display:table-header-group; }
}

.extract-person { margin:0 0 28px; }
.extract-print-seller, .extract-print-total { display:none; }
.extract-person[hidden] { display:none !important; }
.extract-navigation { display:flex; flex-wrap:wrap; align-items:center; gap:8px 12px; margin:0 0 18px; }
.extract-navigation[hidden] { display:none; }
.extract-navigation select { max-width:100%; padding:7px; border:1px solid #d8d0e1; border-radius:6px; background:#fff; color:inherit; font:inherit; }
.extract-person h2 { margin:0 0 12px; font:700 clamp(16px,2.5vw,23px)/1.3 system-ui,Arial,sans-serif; overflow-wrap:anywhere; }
.extract-period { display:flex; flex-wrap:wrap; gap:12px 28px; margin:0 0 18px; padding:12px 16px; background:#f5f2f8; border-radius:8px; }
.extract-period dt { font-size:.88em; font-weight:650; color:#514960; }
.extract-period dd { margin:3px 0 0; font-weight:650; }
.extract-person h3 { margin:18px 0 7px; font:700 15px/1.35 system-ui,Arial,sans-serif; }
.extract-table-scroll { overflow-x:auto; border:1px solid #e4dfeb; border-radius:8px; }
.extract-table-scroll:focus-visible { outline:2px solid #b20b61; outline-offset:2px; }
.extract-table { width:100%; min-width:940px; font-variant-numeric:tabular-nums; }
.extract-table[data-commission="6"] { min-width:600px; }
.cap-doc-scroll .extract-table th, .cap-doc-scroll .extract-table td { border:0; border-bottom:1px solid #e4dfeb; padding:6px; }
.extract-table th { text-align:center; }
.extract-table .extract-bands th { background:#eae4f0; }
.extract-table tbody > tr:nth-child(even) { background:#faf9fc; }
.extract-table tbody > tr.extract-subtotal { background:#f0edf5; }
.extract-table tr > :nth-last-child(2) { white-space:nowrap; }
.extract-table thead tr:last-child > :nth-last-child(2),
.extract-table tbody > tr > :nth-last-child(2) { color:#761349; font-weight:700; }
.extract-table tr > :last-child { overflow-wrap:anywhere; }
.extract-table tbody td table td { padding:2px 4px; border:0; }
.extract-total { background:#eee8f3; color:#521739; }
.cap-doc-scroll .extract-total td { padding:12px; }
@media print {
  .extract-navigation { display:none; }
  .extract-person { margin:0; }
  .extract-person:not([hidden]) ~ .extract-person:not([hidden]) { break-before:page; }
  .extract-person > br { display:none; }
  .extract-person.has-print-total .extract-total { display:none; }
  .extract-print-total { display:block; margin-left:auto; }
  .extract-print-total dd { font-size:10pt; }
  .extract-print-seller { display:table-row; }
  .extract-table .extract-print-seller th { text-align:left; font-weight:700; }
  .extract-person.has-print-total > table:not(:has(tr:not(.extract-total))) { display:none; }
  .extract-period { break-inside:avoid; break-after:avoid; }
  .extract-period { padding:1mm 2mm; margin:0 0 1.5mm; gap:1mm 5mm; }
  .extract-period dd { margin-top:0; }
  .extract-person h2 { font-size:11pt; margin:0 0 1.5mm; }
  .extract-person h3 { font:700 10pt/1.2 system-ui,Arial,sans-serif; margin:1.5mm 0 .7mm; }
  .extract-table .extract-subtotal-group { break-inside:avoid; page-break-inside:avoid; }
  .extract-table .extract-last-sale { break-after:avoid; }
  .cap-doc-scroll .extract-table tbody td table td { padding:0 1px; border:0; }
  .extract-table tbody td table td:nth-child(-n+4) { white-space:nowrap; }
  .extract-person h2, .extract-person h3 { break-after:avoid; }
  .extract-table-scroll { overflow:visible; border-radius:0; break-inside:auto; }
  .extract-table, .extract-table[data-commission="6"] { min-width:0; font-size:8pt; line-height:1.2; }
  .cap-doc-scroll .extract-table td, .cap-doc-scroll .extract-table th { padding:1px 3px; }
  .cap-doc-scroll .extract-total td { padding:2mm 3mm; }
}
</style>
</head>

<body class="cap-report-page">
<main class="cap-document" data-cap-report>
<%If Request.QueryString("modal") <> "1" Then%><div class="cap-doc-toolbar"><button class="cap-doc-button cap-doc-primary" type="button" onclick="window.print()">Imprimir / Salvar PDF</button></div><%End If%>
<header class="cap-doc-header"><h1>Extrato de vendedores e coordenadores</h1><img class="cap-doc-logo" src="images/logoparafundopreto.png" alt="Stillos"></header>
<div class="extract-navigation" id="extract-navigation" hidden>
  <label for="extract-seller">Filtrar por vendedor ou coordenador</label>
  <select id="extract-seller">
    <option value="">Todos os vendedores e coordenadores</option>
  </select>
</div>
<div class="cap-doc-scroll">

<%
AbrirConexao()
'//////////////////////////////////////////////////////////////

VERIFICA_COMISSAO = 0

'oBS.: Nos casos dos vendedores coordenadores 7 e 14 os mesmos dever�o aparecer mesmo que n�o tenham venda

'Set rsMaster = objConexao.Execute("SELECT * FROM (SELECT "&_
'	"TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR ,"&_
'	"TB_VENDEDORES.NOME_VENDEDOR "&_
'	"FROM TB_VENDEDORES INNER JOIN TB_FECHAMENTO_PAGAMENTO ON TB_VENDEDORES.COD_VENDEDOR = TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR "&_
'	"WHERE (TB_FECHAMENTO_PAGAMENTO.COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR <> 27) OR TB_VENDEDORES.COD_VENDEDOR IN (SELECT cod_vendedor FROM TB_FECHAMENTO_PENDENCIAS_PAGAMENTO WHERE cod_fechamento_periodo = '"&COD_FECHAMENTO_PERIODO&"') or TB_VENDEDORES.COD_VENDEDOR IN(7,14,48) "&_
'	"GROUP BY TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR ORDER BY TB_VENDEDORES.NOME_VENDEDOR UNION ALL "&_
'	"SELECT TB_FECHAMENTO_PAGAMENTO_BOLETOS.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR FROM TB_VENDEDORES INNER JOIN TB_FECHAMENTO_PAGAMENTO_BOLETOS ON TB_VENDEDORES.COD_VENDEDOR = TB_FECHAMENTO_PAGAMENTO_BOLETOS.COD_VENDEDOR WHERE (TB_FECHAMENTO_PAGAMENTO_BOLETOS.COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&") GROUP BY TB_FECHAMENTO_PAGAMENTO_BOLETOS.COD_VENDEDOR ORDER BY NOME_VENDEDOR) AS teste GROUP BY teste.COD_VENDEDOR")


Dim filtroExVendedores
filtroExVendedores = ""
If Len(idsExVendedoresFechamento) > 0 Then
    filtroExVendedores = " WHERE teste.COD_VENDEDOR NOT IN (" & idsExVendedoresFechamento & ")"
End If
Set rsMaster = objConexao.Execute("SELECT * FROM ( SELECT  TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR , TB_VENDEDORES.NOME_VENDEDOR FROM TB_VENDEDORES INNER JOIN TB_FECHAMENTO_PAGAMENTO ON TB_VENDEDORES.COD_VENDEDOR = TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR WHERE (TB_FECHAMENTO_PAGAMENTO.COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR <> 27) OR TB_VENDEDORES.COD_VENDEDOR  IN (SELECT cod_vendedor FROM TB_FECHAMENTO_PENDENCIAS_PAGAMENTO WHERE cod_fechamento_periodo = '"&COD_FECHAMENTO_PERIODO&"') or TB_VENDEDORES.COD_VENDEDOR IN(7,14,48) UNION ALL SELECT TB_FECHAMENTO_PAGAMENTO_BOLETOS.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR FROM TB_VENDEDORES INNER JOIN TB_FECHAMENTO_PAGAMENTO_BOLETOS ON TB_VENDEDORES.COD_VENDEDOR = TB_FECHAMENTO_PAGAMENTO_BOLETOS.COD_VENDEDOR WHERE (TB_FECHAMENTO_PAGAMENTO_BOLETOS.COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&")) AS teste" & filtroExVendedores & " GROUP BY teste.COD_VENDEDOR ORDER BY teste.NOME_VENDEDOR")






If not rsMaster.Eof Then
	While not rsMaster.Eof
		COD_VENDEDOR 	= rsMaster("COD_VENDEDOR")
		NOME_VENDEDOR 	= rsMaster("NOME_VENDEDOR")
		%>

<section class="extract-person" aria-labelledby="vendedor-<%=COD_VENDEDOR%>">
  <h2 tabindex="-1" id="vendedor-<%=COD_VENDEDOR%>"><%=Server.HTMLEncode(Ucase(NOME_VENDEDOR & ""))%></h2>
  <dl class="extract-period">
    <div><dt>Pagamento</dt><dd><%=ArrumaData(data_pagamento)%></dd></div>
    <div><dt>Per&iacute;odo de vendas</dt><dd><%=data_inicio%> a <%=data_termino%></dd></div>
    <div><dt>Emitido em</dt><dd><%=ArrumaData(Now())%></dd></div>
  </dl>
				<%
				'1 - VENDAS NORMAIS
				'Set rs = objConexao.Execute("SELECT * FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND COD_VENDEDOR = "&COD_VENDEDOR&" AND TIPO_VENDA = 1",TOTAL_PEDIDOS_TIPO_VENDA_1)
				Set rs = objConexao.Execute("SELECT TB_FECHAMENTO_PAGAMENTO.*, (select sum(valor_un * quantidade) from TB_PEDIDOS_PRODUTOS b where b.cod_pedido=TB_FECHAMENTO_PAGAMENTO.cod_pedido and cod_produto=13) as frete FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND COD_VENDEDOR = "&COD_VENDEDOR&" AND TIPO_VENDA = 1",TOTAL_PEDIDOS_TIPO_VENDA_1)
				frt = 0
				XVALOR_TOTAL_COMISSAO = CCur(0)
				If not rs.Eof Then
					%>
					<h3>Vendas normais</h3>
<div class="extract-table-scroll" tabindex="0" role="region" aria-label="Vendas normais">
<table class="extract-table" aria-label="Vendas normais" data-commission="14">
<thead><tr class="extract-bands"><th colspan="4" scope="colgroup">Pedido</th><th colspan="4" scope="colgroup">Recebimento</th><th colspan="3" scope="colgroup">Itens</th><th colspan="3" scope="colgroup">Valores</th><th colspan="1" scope="colgroup">Formandos</th></tr><tr>
						<th scope="col"><div align="center"><strong>C&oacute;d. Pedido </strong></div></th>
						<th scope="col"><div align="center"><strong>Data da venda </strong></div></th>
						<th scope="col"><div align="center"><strong>Tipo Venda </strong></div></th>
						<th scope="col"><div align="center"><strong>Forma Pag.</strong></div></th>
						<th scope="col"><div align="center"><strong>Dinheiro</strong></div></th>
						<th scope="col"><div align="center"><strong>Cheque</strong></div></th>
						<th scope="col"><div align="center"><strong>Cart&atilde;o</strong></div></th>
						<th scope="col"><div align="center"><strong>Boleto</strong></div></th>
						<th scope="col"><div align="center"><strong>&Aacute;lbuns</strong></div></th>
						<th scope="col"><div align="center"><strong>DVDs</strong></div></th>
						<th scope="col"><div align="center"><strong>Foto-Telas e Trios </strong></div></th>
						<th scope="col"><div align="center"><strong>Ajuda de Custo </strong></div></th>
						<th scope="col"><div align="center"><strong>Frete </strong></div></th>
						<th scope="col"><div align="center"><strong>Comiss&atilde;o</strong></div></th>
						<th scope="col"><div align="center"><strong>Formandos e institui&ccedil;&atilde;o</strong></div></th>
					</tr></thead><tbody>
					<%
					While not rs.Eof
						COD_PEDIDO			= rs("COD_PEDIDO")
						DATA_VENDA			= formatDateTime(rs("DATA_VENDA"))
						FORMA_PAGAMENTO		= rs("FORMA_PAGAMENTO")
						VALOR_AJUDA_CUSTO	= formatNumber(0&rs("VALOR_AJUDA_CUSTO"))/1
						COD_GRUPO_VENDEDORES= rs("COD_GRUPO_VENDEDORES")
						TOTAL_DINHEIRO		= formatNumber(0&rs("TOTAL_DINHEIRO"))/1
						TOTAL_CHEQUE		= formatNumber(0&rs("TOTAL_CHEQUE"))/1
						TOTAL_CARTAO		= formatNumber(0&rs("TOTAL_CARTAO"))/1
						TOTAL_BOLETO 		= formatNumber(0&rs("TOTAL_BOLETO"))/1
						TOTAL_RECEBIMENTO	= formatNumber(0&rs("TOTAL_RECEBIMENTO"))/1
						FRETE				= formatNumber(0&rs("frete"))/1
						QTD_ALBUNS			= rs("QTD_ALBUNS")
						QTD_DVDS			= rs("QTD_DVDS")
						QTD_FOTO_TELAS_TRIOS= rs("QTD_FOTO_TELAS_TRIOS")
						VALOR_TOTAL_COMISSAO = 0 : If Not IsNull(rs("VALOR_TOTAL_COMISSAO")) And rs("VALOR_TOTAL_COMISSAO") <> "" Then VALOR_TOTAL_COMISSAO = CDbl(rs("VALOR_TOTAL_COMISSAO"))
						
						frt = frt + FRETE
						%>
						<tr>
							<td><div align="center"><%=COD_PEDIDO%></div></td>
							<td><div align="center"><%=ArrumaData(DATA_VENDA)%></div></td>
							<td>
									<div align="center">
										<%
										'=================== AJUSTE PARA PRÉ-VENDA ===================
Dim rsPreVenda, pre_venda_status, total_boleto_valor, tipo_venda_texto
    Set rsPreVenda = objConexao.Execute("SELECT PRE_VENDA FROM TB_PEDIDOS WHERE COD_PEDIDO = " & COD_PEDIDO)

    If Not rsPreVenda.Eof Then
        pre_venda_status = rsPreVenda("PRE_VENDA")
        
        ' Define texto amigável
        If UCase(Trim(pre_venda_status)) = "S" Then
            tipo_venda_texto = "PRÉ-VENDA"
        Else
            tipo_venda_texto = "NORMAL"
        End If
        
        ' Garante que TOTAL_BOLETO é numérico
        If IsNull(rs("TOTAL_BOLETO")) Or Trim(rs("TOTAL_BOLETO")) = "" Then
            total_boleto_valor = 0
        Else
            total_boleto_valor = CDbl(rs("TOTAL_BOLETO"))
        End If

        ' Verifica se é pré-venda com boleto → zera comissão
        If UCase(Trim(pre_venda_status)) = "S" And total_boleto_valor > 0 Then
            VALOR_TOTAL_COMISSAO = 0
        End If
    End If

    rsPreVenda.Close
    Set rsPreVenda = Nothing

    ' Exibe o tipo de venda formatado
    Response.Write tipo_venda_texto
										%>
									</div>
								</td>

							<td><div align="center"><%=pedidosFormataFormaPagamento(FORMA_PAGAMENTO)%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(TOTAL_DINHEIRO))%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(TOTAL_CHEQUE))%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(TOTAL_CARTAO))%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(TOTAL_BOLETO))%></div></td>
							<td><div align="center"><%=retiraZero2(QTD_ALBUNS)%></div></td>
							<td><div align="center"><%=retiraZero2(QTD_DVDS)%></div></td>
							<td><div align="center"><%=retiraZero2(QTD_FOTO_TELAS_TRIOS)%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(VALOR_AJUDA_CUSTO))%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(FRETE))%></div></td>
							<% XVALOR_TOTAL_COMISSAO = XVALOR_TOTAL_COMISSAO + CCur(FormatNumber(VALOR_TOTAL_COMISSAO, 2)) %>
							<td><div align="right"><%=retiraZero2(formatNumber(VALOR_TOTAL_COMISSAO))%></div></td>
							<td><div align="center"><%=pedidosFormandosFormatado(COD_PEDIDO, true)%></div></td>
						</tr>
						<%
					rs.MoveNext()
					Wend
					
					'TOTAL 1 - VENDAS NORMAIS
Set rsTotal = objConexao.Execute("SELECT COUNT(*) AS XPEDIDOS, SUM(TOTAL_DINHEIRO) AS XTOTAL_DINHEIRO, SUM(TOTAL_BOLETO) AS XTOTAL_BOLETO, SUM(TOTAL_CHEQUE) AS XTOTAL_CHEQUE, SUM(TOTAL_CARTAO) AS XTOTAL_CARTAO, SUM(TOTAL_RECEBIMENTO) AS XTOTAL_RECEBIMENTO, SUM(QTD_ALBUNS) AS XQTD_ALBUNS, SUM(QTD_DVDS) AS XQTD_DVDS, SUM(QTD_FOTO_TELAS_TRIOS) AS XQTD_FOTO_TELAS_TRIOS, SUM(VALOR_AJUDA_CUSTO) AS XVALOR_AJUDA_CUSTO, SUM(VALOR_TOTAL_COMISSAO) AS XVALOR_TOTAL_COMISSAO FROM TB_FECHAMENTO_PAGAMENTO fp WHERE fp.COD_VENDEDOR = " & COD_VENDEDOR & " AND fp.TIPO_VENDA = 1 AND fp.COD_FECHAMENTO_PERIODO = " & COD_FECHAMENTO_PERIODO & " AND fp.COD_PEDIDO NOT IN (SELECT COD_PEDIDO FROM TB_PEDIDOS WHERE PRE_VENDA = 'S' AND TOTAL_BOLETO > 0)")
					If not rsTotal.Eof Then
						XPEDIDOS				= rsTotal("XPEDIDOS")
						XTOTAL_DINHEIRO			= formatNumber(0&rsTotal("XTOTAL_DINHEIRO"))
						XTOTAL_CHEQUE			= formatNumber(0&rsTotal("XTOTAL_CHEQUE"))
						XTOTAL_CARTAO			= formatNumber(0&rsTotal("XTOTAL_CARTAO"))
						XTOTAL_BOLETO			= formatNumber(0&rsTotal("XTOTAL_BOLETO"))
						XTOTAL_RECEBIMENTO 		= formatNumber(0&rsTotal("XTOTAL_RECEBIMENTO"))
						XQTD_ALBUNS				= rsTotal("XQTD_ALBUNS")
						XQTD_DVDS				= rsTotal("XQTD_DVDS")
						XQTD_FOTO_TELAS_TRIOS	= rsTotal("XQTD_FOTO_TELAS_TRIOS")
						XVALOR_AJUDA_CUSTO		= formatNumber(0&rsTotal("XVALOR_AJUDA_CUSTO"))
						
					End If
					rsTotal.Close
					Set rsTotal = Nothing
					%>
					<tr class="extract-subtotal">
						<td><div align="center"><strong><%=XPEDIDOS%></strong></div></td>
						<td><div align="center"></div></td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XTOTAL_DINHEIRO))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XTOTAL_CHEQUE))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XTOTAL_CARTAO))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XTOTAL_BOLETO))%></strong></div></td>
						<td><div align="center"><strong><%=XQTD_ALBUNS%></strong></div></td>
						<td><div align="center"><strong><%=XQTD_DVDS%></strong></div></td>
						<td><div align="center"><strong><%=XQTD_FOTO_TELAS_TRIOS%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XVALOR_AJUDA_CUSTO))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(frt))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XVALOR_TOTAL_COMISSAO))%></strong></div></td>
						<td>&nbsp;</td>
					</tr>
					</tbody></table>
</div>
         		<br />
					<%		
				End If
				rs.Close
				Set rs = Nothing
				
				'2 - VENDAS REPASSE
				'Set rs = objConexao.Execute("SELECT * FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND COD_VENDEDOR = "&COD_VENDEDOR&" AND TIPO_VENDA = 2",TOTAL_PEDIDOS_TIPO_VENDA_2)
				Set rs = objConexao.Execute("SELECT TB_FECHAMENTO_PAGAMENTO.*, (select sum(valor_un * quantidade) from TB_PEDIDOS_PRODUTOS b where b.cod_pedido=TB_FECHAMENTO_PAGAMENTO.cod_pedido and cod_produto=13) as frete FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND COD_VENDEDOR = "&COD_VENDEDOR&" AND TIPO_VENDA = 2",TOTAL_PEDIDOS_TIPO_VENDA_2)
				frt = 0
				If not rs.Eof Then
					%>
   	  			<h3>Vendas de repasse</h3>
<div class="extract-table-scroll" tabindex="0" role="region" aria-label="Vendas de repasse">
<table class="extract-table" aria-label="Vendas de repasse" data-commission="13">
<thead><tr class="extract-bands"><th colspan="4" scope="colgroup">Pedido</th><th colspan="3" scope="colgroup">Recebimento</th><th colspan="3" scope="colgroup">Itens</th><th colspan="3" scope="colgroup">Valores</th><th colspan="1" scope="colgroup">Formandos</th></tr><tr>
						<th scope="col"><div align="center"><strong>C&oacute;d. Pedido </strong></div></th>
						<th scope="col"><div align="center"><strong>Data da venda </strong></div></th>
						<th scope="col"><div align="center"><strong>Tipo Venda </strong></div></th>
						<th scope="col"><div align="center"><strong>Forma Pag.</strong></div></th>
						<th scope="col"><div align="center"><strong>Dinheiro</strong></div></th>
						<th scope="col"><div align="center"><strong>Cheque</strong></div></th>
						<th scope="col"><div align="center"><strong>Cart&atilde;o</strong></div></th>
						<th scope="col"><div align="center"><strong>&aacute;lbuns</strong></div></th>
						<th scope="col"><div align="center"><strong>DVDs</strong></div></th>
						<th scope="col"><div align="center"><strong>Foto-Telas e Trios </strong></div></th>
						<th scope="col"><div align="center"><strong>Ajuda de Custo </strong></div></th>
						<th scope="col"><div align="center"><strong>Frete </strong></div></th>
						<th scope="col"><div align="center"><strong>Comiss&atilde;o</strong></div></th>
						<th scope="col"><div align="center"><strong>Formandos e institui&ccedil;&atilde;o</strong></div></th>
					</tr></thead><tbody>
					<%
					While not rs.Eof
						COD_PEDIDO				= rs("COD_PEDIDO")
						DATA_VENDA				= formatDateTime(rs("DATA_VENDA"))
						FORMA_PAGAMENTO		= rs("FORMA_PAGAMENTO")
						VALOR_AJUDA_CUSTO		= formatNumber(0&rs("VALOR_AJUDA_CUSTO"))/1
						COD_GRUPO_VENDEDORES = rs("COD_GRUPO_VENDEDORES")
						TOTAL_DINHEIRO			= formatNumber(0&rs("TOTAL_DINHEIRO"))/1
						TOTAL_CHEQUE			= formatNumber(0&rs("TOTAL_CHEQUE"))/1
						TOTAL_CARTAO			= formatNumber(0&rs("TOTAL_CARTAO"))/1
						TOTAL_RECEBIMENTO		= formatNumber(0&rs("TOTAL_RECEBIMENTO"))/1
						FRETE					= formatNumber(0&rs("frete"))/1
						QTD_ALBUNS				= rs("QTD_ALBUNS")
						QTD_DVDS					= rs("QTD_DVDS")
						QTD_FOTO_TELAS_TRIOS	= rs("QTD_FOTO_TELAS_TRIOS")
						VALOR_TOTAL_COMISSAO	= rs("VALOR_TOTAL_COMISSAO")
						frt = frt + FRETE
						%>
						<tr>
							<td><div align="center"><%=COD_PEDIDO%></div></td>
							<td><div align="center"><%=ArrumaData(DATA_VENDA)%></div></td>
							<td><div align="center">REPASSE</div></td>
							<td><div align="center"><%=pedidosFormataFormaPagamento(FORMA_PAGAMENTO)%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(TOTAL_DINHEIRO))%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(TOTAL_CHEQUE))%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(TOTAL_CARTAO))%></div></td>
							<td><div align="center"><%=retiraZero2(QTD_ALBUNS)%></div></td>
							<td><div align="center"><%=retiraZero2(QTD_DVDS)%></div></td>
							<td><div align="center"><%=retiraZero2(QTD_FOTO_TELAS_TRIOS)%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(VALOR_AJUDA_CUSTO))%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(FRETE))%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(VALOR_TOTAL_COMISSAO))%></div></td>
							<td><div align="center"><%=pedidosFormandosFormatado(COD_PEDIDO, true)%></div></td>
						</tr>
						<%
					rs.MoveNext()
					Wend

					'2 - VENDAS REPASSE TOTAL
Set rsTotal = objConexao.Execute("SELECT COUNT(*) AS XPEDIDOS, SUM(TOTAL_DINHEIRO) AS XTOTAL_DINHEIRO, SUM(TOTAL_BOLETO) AS XTOTAL_BOLETO, SUM(TOTAL_CHEQUE) AS XTOTAL_CHEQUE, SUM(TOTAL_CARTAO) AS XTOTAL_CARTAO, SUM(TOTAL_RECEBIMENTO) AS XTOTAL_RECEBIMENTO, SUM(QTD_ALBUNS) AS XQTD_ALBUNS, SUM(QTD_DVDS) AS XQTD_DVDS, SUM(QTD_FOTO_TELAS_TRIOS) AS XQTD_FOTO_TELAS_TRIOS, SUM(VALOR_AJUDA_CUSTO) AS XVALOR_AJUDA_CUSTO, SUM(VALOR_TOTAL_COMISSAO) AS XVALOR_TOTAL_COMISSAO FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_VENDEDOR = " & COD_VENDEDOR & " AND TIPO_VENDA = 2 AND COD_FECHAMENTO_PERIODO = " & COD_FECHAMENTO_PERIODO & " AND COD_PEDIDO NOT IN (SELECT COD_PEDIDO FROM TB_PEDIDOS WHERE PRE_VENDA = 'S' AND TOTAL_BOLETO > 0)")
					If not rsTotal.Eof Then
						XPEDIDOS						= rsTotal("XPEDIDOS")
						XTOTAL_DINHEIRO			= formatNumber(0&rsTotal("XTOTAL_DINHEIRO"))
						XTOTAL_CHEQUE				= formatNumber(0&rsTotal("XTOTAL_CHEQUE"))
						XTOTAL_CARTAO				= formatNumber(0&rsTotal("XTOTAL_CARTAO"))
						XTOTAL_RECEBIMENTO 		= formatNumber(0&rsTotal("XTOTAL_RECEBIMENTO"))
						XQTD_ALBUNS					= rsTotal("XQTD_ALBUNS")
						XQTD_DVDS					= rsTotal("XQTD_DVDS")
						XQTD_FOTO_TELAS_TRIOS	= rsTotal("XQTD_FOTO_TELAS_TRIOS")
						XVALOR_AJUDA_CUSTO		= formatNumber(0&rsTotal("XVALOR_AJUDA_CUSTO"))
						XVALOR_TOTAL_COMISSAO	= formatNumber(0&rsTotal("XVALOR_TOTAL_COMISSAO"))
					End If
					rsTotal.Close
					Set rsTotal = Nothing
					%>
					<tr class="extract-subtotal">
						<td><div align="center"><strong><%=XPEDIDOS%></strong></div></td>
						<td><div align="center"></div></td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XTOTAL_DINHEIRO))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XTOTAL_CHEQUE))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XTOTAL_CARTAO))%></strong></div></td>
						<td><div align="center"><strong><%=XQTD_ALBUNS%></strong></div></td>
						<td><div align="center"><strong><%=XQTD_DVDS%></strong></div></td>
						<td><div align="center"><strong><%=XQTD_FOTO_TELAS_TRIOS%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XVALOR_AJUDA_CUSTO))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(frt))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XVALOR_TOTAL_COMISSAO))%></strong></div></td>
						<td>&nbsp;</td>
					</tr>
					</tbody></table>
</div>

					<br />
					<%
				End If
				rs.Close
				Set rs = Nothing

				'3 - VENDAS REPASSE (GRUPO ESPECIAL)
				'Set rs = objConexao.Execute("SELECT * FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND COD_VENDEDOR = "&COD_VENDEDOR&" AND TIPO_VENDA = 3",TOTAL_PEDIDOS_TIPO_VENDA_3)
				Set rs = objConexao.Execute("SELECT TB_FECHAMENTO_PAGAMENTO.*, (select sum(valor_un * quantidade) from TB_PEDIDOS_PRODUTOS b where b.cod_pedido=TB_FECHAMENTO_PAGAMENTO.cod_pedido and cod_produto=13) as frete FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND COD_VENDEDOR = "&COD_VENDEDOR&" AND TIPO_VENDA = 3",TOTAL_PEDIDOS_TIPO_VENDA_3)
				frt = 0
				If not rs.Eof Then
					%>
					<h3>Repasses do grupo especial</h3>
<div class="extract-table-scroll" tabindex="0" role="region" aria-label="Repasses do grupo especial">
<table class="extract-table" aria-label="Repasses do grupo especial" data-commission="13">
<thead><tr class="extract-bands"><th colspan="4" scope="colgroup">Pedido</th><th colspan="3" scope="colgroup">Recebimento</th><th colspan="3" scope="colgroup">Itens</th><th colspan="3" scope="colgroup">Valores</th><th colspan="1" scope="colgroup">Formandos</th></tr><tr>
						<th scope="col"><div align="center"><strong>C&oacute;d. Pedido </strong></div></th>
						<th scope="col"><div align="center"><strong>Data da venda </strong></div></th>
						<th scope="col"><div align="center"><strong>Tipo Venda </strong></div></th>
						<th scope="col"><div align="center"><strong>Forma Pag.</strong></div></th>
						<th scope="col"><div align="center"><strong>Dinheiro</strong></div></th>
						<th scope="col"><div align="center"><strong>Cheque</strong></div></th>
						<th scope="col"><div align="center"><strong>Cart&atilde;o</strong></div></th>
						<th scope="col"><div align="center"><strong>&aacute;lbuns</strong></div></th>
						<th scope="col"><div align="center"><strong>DVDs</strong></div></th>
						<th scope="col"><div align="center"><strong>Foto-Telas e Trios </strong></div></th>
						<th scope="col"><div align="center"><strong>Ajuda de Custo </strong></div></th>
						<th scope="col"><div align="center"><strong>Frete </strong></div></th>
						<th scope="col"><div align="center"><strong>Comiss&atilde;o</strong></div></th>
						<th scope="col"><div align="center"><strong>Formandos e institui&ccedil;&atilde;o</strong></div></th>
					</tr></thead><tbody>
					<%
					While not rs.Eof
						COD_PEDIDO				= rs("COD_PEDIDO")
						DATA_VENDA				= formatDateTime(rs("DATA_VENDA"))
						FORMA_PAGAMENTO		= rs("FORMA_PAGAMENTO")
						VALOR_AJUDA_CUSTO		= formatNumber(0&rs("VALOR_AJUDA_CUSTO"))/1
						COD_GRUPO_VENDEDORES = rs("COD_GRUPO_VENDEDORES")
						TOTAL_DINHEIRO			= formatNumber(0&rs("TOTAL_DINHEIRO"))/1
						TOTAL_CHEQUE			= formatNumber(0&rs("TOTAL_CHEQUE"))/1
						TOTAL_CARTAO			= formatNumber(0&rs("TOTAL_CARTAO"))/1
						TOTAL_RECEBIMENTO		= formatNumber(0&rs("TOTAL_RECEBIMENTO"))/1
						FRETE					= formatNumber(0&rs("frete"))/1
						QTD_ALBUNS				= rs("QTD_ALBUNS")
						QTD_DVDS					= rs("QTD_DVDS")
						QTD_FOTO_TELAS_TRIOS = rs("QTD_FOTO_TELAS_TRIOS")
						VALOR_TOTAL_COMISSAO = rs("VALOR_TOTAL_COMISSAO")
						frt = frt + FRETE
						%>
						<tr>
							<td><div align="center"><%=COD_PEDIDO%></div></td>
							<td><div align="center"><%=ArrumaData(DATA_VENDA)%></div></td>
							<td><div align="center">REPASSE (GRUPO ESPECIAL)</div></td>
							<td><div align="center"><%=pedidosFormataFormaPagamento(FORMA_PAGAMENTO)%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(TOTAL_DINHEIRO))%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(TOTAL_CHEQUE))%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(TOTAL_CARTAO))%></div></td>
							<td><div align="center"><%=retiraZero2(QTD_ALBUNS)%></div></td>
							<td><div align="center"><%=retiraZero2(QTD_DVDS)%></div></td>
							<td><div align="center"><%=retiraZero2(QTD_FOTO_TELAS_TRIOS)%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(VALOR_AJUDA_CUSTO))%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(FRETE))%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(VALOR_TOTAL_COMISSAO))%></div></td>
							<td><div align="center"><%=pedidosFormandosFormatado(COD_PEDIDO, true)%></div></td>
						</tr>
						<%
					rs.MoveNext()
					Wend

					'3 - VENDAS REPASSE (GRUPO ESPECIAL) TOTAL
Set rsTotal = objConexao.Execute("SELECT COUNT(*) AS XPEDIDOS, SUM(TOTAL_DINHEIRO) AS XTOTAL_DINHEIRO, SUM(TOTAL_BOLETO) AS XTOTAL_BOLETO, SUM(TOTAL_CHEQUE) AS XTOTAL_CHEQUE, SUM(TOTAL_CARTAO) AS XTOTAL_CARTAO, SUM(TOTAL_RECEBIMENTO) AS XTOTAL_RECEBIMENTO, SUM(QTD_ALBUNS) AS XQTD_ALBUNS, SUM(QTD_DVDS) AS XQTD_DVDS, SUM(QTD_FOTO_TELAS_TRIOS) AS XQTD_FOTO_TELAS_TRIOS, SUM(VALOR_AJUDA_CUSTO) AS XVALOR_AJUDA_CUSTO, SUM(VALOR_TOTAL_COMISSAO) AS XVALOR_TOTAL_COMISSAO FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_VENDEDOR = " & COD_VENDEDOR & " AND TIPO_VENDA = 3 AND COD_FECHAMENTO_PERIODO = " & COD_FECHAMENTO_PERIODO & " AND COD_PEDIDO NOT IN (SELECT COD_PEDIDO FROM TB_PEDIDOS WHERE PRE_VENDA = 'S' AND TOTAL_BOLETO > 0)")
					If not rsTotal.Eof Then
						XPEDIDOS						= rsTotal("XPEDIDOS")
						XTOTAL_DINHEIRO			= formatNumber(0&rsTotal("XTOTAL_DINHEIRO"))
						XTOTAL_CHEQUE				= formatNumber(0&rsTotal("XTOTAL_CHEQUE"))
						XTOTAL_CARTAO				= formatNumber(0&rsTotal("XTOTAL_CARTAO"))
						XTOTAL_RECEBIMENTO 		= formatNumber(0&rsTotal("XTOTAL_RECEBIMENTO"))
						XQTD_ALBUNS					= rsTotal("XQTD_ALBUNS")
						XQTD_DVDS					= rsTotal("XQTD_DVDS")
						XQTD_FOTO_TELAS_TRIOS	= rsTotal("XQTD_FOTO_TELAS_TRIOS")
						XVALOR_AJUDA_CUSTO		= formatNumber(0&rsTotal("XVALOR_AJUDA_CUSTO"))
						XVALOR_TOTAL_COMISSAO	= formatNumber(0&rsTotal("XVALOR_TOTAL_COMISSAO"))
					End If
					rsTotal.Close
					Set rsTotal = Nothing
					%>

					<tr class="extract-subtotal">
						<td><div align="center"><strong><%=XPEDIDOS%></strong></div></td>
						<td><div align="center"></div></td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XTOTAL_DINHEIRO))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XTOTAL_CHEQUE))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XTOTAL_CARTAO))%></strong></div></td>
						<td><div align="center"><strong><%=XQTD_ALBUNS%></strong></div></td>
						<td><div align="center"><strong><%=XQTD_DVDS%></strong></div></td>
						<td><div align="center"><strong><%=XQTD_FOTO_TELAS_TRIOS%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XVALOR_AJUDA_CUSTO))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(frt))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XVALOR_TOTAL_COMISSAO))%></strong></div></td>
						<td>&nbsp;</td>
					</tr>		
					</tbody></table>
</div>
            	<%		
				End If
				rs.Close
				Set rs = Nothing	

				'4 - PAGAMENTO DE BOLETOS
				set rsBoleto = objConexao.execute("SELECT * FROM TB_FECHAMENTO_PAGAMENTO_BOLETOS WHERE COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND COD_VENDEDOR = "&COD_VENDEDOR)
				if not rsBoleto.eof then
					%>
					<h3>Pagamento de boletos</h3>
<div class="extract-table-scroll" tabindex="0" role="region" aria-label="Pagamento de boletos">
<table class="extract-table" aria-label="Pagamento de boletos" data-commission="6">
<thead><tr>
						<th scope="col"><div align="center"><strong>C&oacute;d. Pedido</strong></div></th>
						<th scope="col"><div align="center"><strong>Parcela </strong></div></th>
						<th scope="col"><div align="center"><strong>Valor</strong></div></th>
						<th scope="col"><div align="center"><strong>Data Pagamento</strong></div></th>
						<th scope="col"><div align="center"><strong>Data Baixa</strong></div></th>
						<th scope="col"><div align="center"><strong>Comiss&atilde;o</strong></div></th>
						<th scope="col"><div align="center"><strong>Formandos e institui&ccedil;&atilde;o</strong></div></th>
					</tr></thead><tbody>
					<% 
					while not rsBoleto.eof
						%>
						<tr>
							<td width="5%" height="30"><div align="center"><%=rsBoleto("COD_PEDIDO")%></div></td>
							<td width="7%"><div align="center"><%=rsBoleto("PARCELA")%></div></td>
							<td width="5%"><div align="center"><%=ArrumaNumero(rsBoleto("VALOR"),"-",2)%></div></td>
							<td width="8%"><div align="center"><%=ArrumaData(rsBoleto("data_pagamento"))%></div></td>
							<td width="8%"><div align="center"><%=ArrumaData(rsBoleto("data_baixa"))%></div></td>
							<td width="5%"><div align="center"><%=ArrumaNumero(rsBoleto("valor_comissao_vendedor"),"-",2)%></div></td>
							<td><div align="center"><%=pedidosFormandosFormatado(rsBoleto("COD_PEDIDO"), true)%></div></td>
						</tr>
						<%
					rsBoleto.MoveNext()
					Wend				
					
					'4 - PAGAMENTO DE BOLETOS TOTAL
					%>
					</tbody></table>
</div>
					<br />
					<%
				end if
				rsBoleto.Close
				set rsBoleto = Nothing

				
				%>
				<table width="100%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
				<%
				'TOTAIS

				Dim TOTAL_AJUDA_COMISSAO, TOTAL_METAS_GERAL, TOTAL_PENDENCIAS
				TOTAL_AJUDA_COMISSAO = 0
				TOTAL_METAS_GERAL    = 0
				TOTAL_PENDENCIAS     = 0

				'///////////////////////////////////////////////////////////////

				Dim total_por_vend_PEDIDOS, total_por_vend_TOTAL_DINHEIRO, total_por_vend_TOTAL_CHEQUE, total_por_vend_TOTAL_CARTAO, total_por_vend_TOTAL_RECEBIMENTO, total_por_vend_QTD_ALBUNS, total_por_vend_VALOR_AJUDA_CUSTO, total_por_vend_TOTAL_COMISSAO

				total_por_vend_PEDIDOS           = 0
				total_por_vend_TOTAL_DINHEIRO    = 0
				total_por_vend_TOTAL_CHEQUE      = 0
				total_por_vend_TOTAL_CARTAO	     = 0
				total_por_vend_TOTAL_RECEBIMENTO = 0
				total_por_vend_QTD_ALBUNS        = 0
				total_por_vend_TOTAL_COMISSAO    = 0
				total_por_vend_VALOR_AJUDA_CUSTO = 0

				If int(TOTAL_PEDIDOS_TIPO_VENDA_3) > 0 Then	
					Dim TOTAL_METAS_1, TOTAL_METAS_2
					TOTAL_METAS_1	= 0
					TOTAL_METAS_2	= 0
				End If
			
				If TOTAL_METAS_GERAL > 0 Then
					%>
					<tr bgcolor="whitesmoke">
						<td class="textoMedio"><div align="left"><strong class="textoMedio">TOTAL METAS</strong></div></td>
						<td width="81%"><div align="right"><%=fechamentoRetornaValorRecebidoVendedorTipo(3,cod_vendedor,cod_fechamento_periodo,true)%></div></td>
					</tr>
					<%
				End If

				Set rsPendencias = objConexao.Execute("SELECT * FROM TB_FECHAMENTO_PENDENCIAS_PAGAMENTO WHERE COD_VENDEDOR = "&COD_VENDEDOR&" AND COD_FECHAMENTO_PERIODO ="&COD_FECHAMENTO_PERIODO)
				If not rsPendencias.Eof Then
					Dim SINAL
					%>
					<tr bgcolor="whitesmoke">    
						<td class="textoMedio"><div align="left"><strong>PEND&Ecirc;NCIAS</strong></div></td>
						<td>
							<table bordercolor="#CCCCCC" width="100%" border="1" cellspacing="0" cellpadding="3">
							<tr>
								<td width="61%"><div align="center"><strong>Descri&ccedil;&atilde;o</strong></div></td>
								<td width="25%"><div align="center"><strong>Tipo </strong></div></td>
								<td width="14%"><div align="center"><strong>Valor</strong></div></td>
							</tr>
							<%
							While not rsPendencias.Eof
								Dim VALOR_PENDENCIA
								VALOR_PENDENCIA = formatNumber(0&rsPendencias("VALOR"))
								%>
									<tr>
										<td><%=rsPendencias("MOTIVO")%></td>
										<td>
											<%
												If rsPendencias("TIPO")  = "A"  Then
													Response.Write("ACR�SCIMO")
													SINAL = "+"
													TOTAL_PENDENCIAS = (TOTAL_PENDENCIAS + VALOR_PENDENCIA)/1
												Else
													TOTAL_PENDENCIAS = (TOTAL_PENDENCIAS - VALOR_PENDENCIA)/1
													Response.Write("DECR�CIMO")
													SINAL = "-"
												End If
											%>
										</td>
										<td><div align="right"><strong><%="("&SINAL&") "&formatNumber(rsPendencias("VALOR"))%></strong></div></td>
									</tr>
								<%
							rsPendencias.MoveNext()
							Wend
							%>
							</table></td>
						</tr>
					<%
				End If
				rsPendencias.Close
				Set rsPendencias = Nothing
					%>
				
				<tr class="extract-total">
					<td width="19%" class="textoMedio extract-total-label"><div align="left"><span class="textoMaior"><strong>TOTAL A RECEBER</strong></span><strong><br /></strong></div></td>
<td><div align="right" class="textoMaior"><strong><% Set rsTotalReceber = objConexao.Execute("SELECT IFNULL(SUM(fp.VALOR_TOTAL_COMISSAO),0) AS TOTAL_A_RECEBER FROM TB_FECHAMENTO_PAGAMENTO fp LEFT JOIN TB_PEDIDOS p ON p.COD_PEDIDO = fp.COD_PEDIDO WHERE fp.COD_VENDEDOR=" & COD_VENDEDOR & " AND fp.COD_FECHAMENTO_PERIODO=" & COD_FECHAMENTO_PERIODO & " AND NOT (p.PRE_VENDA='S' AND fp.TOTAL_BOLETO > 0)") : If Not rsTotalReceber.Eof Then Response.Write(formatCurrency(rsTotalReceber("TOTAL_A_RECEBER"))) Else Response.Write(formatCurrency(0)) : End If : rsTotalReceber.Close : Set rsTotalReceber = Nothing %></strong></div></td>
				</tr>	
				</table>
		
		</section>

	<%
	rsMaster.MoveNext()
	Wend
Else
%>
<br />
<br />
<br />
<br />
<P align="center">N&atilde;o h&aacute; dados para o per&iacute;odo selecionado</P>
<%
	End If	
	rsMaster.Close
	Set rsMaster = Nothing
	FecharConexao()
%>
</div></main>
<script>
(function () {
  document.querySelectorAll('.extract-subtotal').forEach(function (subtotal) {
    var lastSale = subtotal.previousElementSibling;
    if (!lastSale) return;

    var rows = subtotal.parentNode;
    var subtotalGroup = rows.parentNode.createTBody();
    subtotalGroup.className = 'extract-subtotal-group';
    lastSale.classList.add('extract-last-sale');
    subtotalGroup.appendChild(lastSale);
    subtotalGroup.appendChild(subtotal);
  });

  var people = document.querySelectorAll('.extract-person');
  people.forEach(function (person) {
    var seller = person.querySelector('h2').textContent.trim();
    var period = person.querySelector('.extract-period');
    var total = person.querySelector('.extract-total td:last-child');
    if (period && total && total.textContent.trim()) {
      var summary = document.createElement('div');
      var label = document.createElement('dt');
      var amount = document.createElement('dd');
      summary.className = 'extract-print-total';
      label.textContent = 'Total a receber';
      amount.textContent = total.textContent.trim();
      summary.appendChild(label);
      summary.appendChild(amount);
      period.appendChild(summary);
      person.classList.add('has-print-total');
    }

    person.querySelectorAll('.extract-table thead').forEach(function (head) {
      var columns = head.rows[head.rows.length - 1].cells.length;
      var row = head.insertRow(0);
      var title = document.createElement('th');
      row.className = 'extract-print-seller';
      title.colSpan = columns;
      title.textContent = seller + ' | ' + head.parentNode.getAttribute('aria-label');
      row.appendChild(title);
    });
  });
  if (people.length < 2) return;

  var navigation = document.getElementById('extract-navigation');
  var select = document.getElementById('extract-seller');
  people.forEach(function (person) {
    var heading = person.querySelector('h2');
    select.add(new Option(heading.textContent.trim(), heading.id));
  });

  function filterPeople() {
    var selected = select.value;
    people.forEach(function (person) {
      person.hidden = selected !== '' && person.getAttribute('aria-labelledby') !== selected;
    });
  }

  select.addEventListener('change', filterPeople);
  filterPeople();
  navigation.hidden = false;
}());
</script>
</body>
</html>



