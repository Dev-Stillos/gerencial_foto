<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Valida��o especial para verificar se possui sub-valida��o espec�fica
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '18' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se Não encontrar ent�o redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
		
	rs.Close
	Set rs=Nothing
	
FecharConexao()


Dim DATA_VENDA_INICIO, DATA_VENDA_FIM
Dim sqlQuery
DATA_VENDA_INICIO = Request.QueryString("DATA_VENDA_INICIO")
DATA_VENDA_FIM = Request.QueryString("DATA_VENDA_FIM")

Dim rankModal, rankQuery
rankModal = (Request.QueryString("modal") = "1")
rankQuery = "data_venda_inicio=" & Server.URLEncode(DATA_VENDA_INICIO) & "&data_venda_fim=" & Server.URLEncode(DATA_VENDA_FIM)

If DATA_VENDA_INICIO = Empty or DATA_VENDA_FIM = Empty Then
    If rankModal Then
        Response.Status = "400 Bad Request"
        Response.Write "Informe o per&iacute;odo do ranking."
        Response.End
    End If
	Response.Redirect("adm_relatorios.asp?msg=Informe uma data de venda de início e outra de fim para gerar o relatório.")
End If

If rankModal Then rankQuery = rankQuery & "&modal=1"

sqlQuery = "P.DATA_VENDA BETWEEN " & sqlData(DATA_VENDA_INICIO) & " AND " & sqlData(DATA_VENDA_FIM)

%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Ranking de representantes comerciais</title>
<style>
* { box-sizing:border-box; }
body { margin:0; background:#f6f7fb; color:#242738; font:14px/1.45 system-ui,Arial,sans-serif; }
.rank-report { max-width:1200px; margin:16px auto; padding:clamp(12px,3vw,28px); border:1px solid #e4dfeb; border-radius:12px; background:#fff; }
.rank-header { display:flex; justify-content:space-between; align-items:center; gap:12px; padding-bottom:12px; border-bottom:2px solid #b20b61; }
h1 { margin:0; font-size:clamp(18px,3vw,28px); }
.rank-period { margin:4px 0 0; color:#514960; }
.rank-logo { width:clamp(90px,18vw,150px); height:auto; background:#292b40; border-radius:7px; }
.rank-tools { display:flex; flex-wrap:wrap; gap:8px; margin:0 0 14px; }
a { color:#861451; }
.rank-tools a,.rank-tools button,.rank-action a { display:inline-block; padding:6px 10px; border:1px solid #e4dfeb; border-radius:6px; background:#fff; color:#861451; text-decoration:none; font:inherit; cursor:pointer; }
.rank-tools button { background:#b20b61; color:#fff; }
h2 { margin:22px 0 8px; font-size:20px; }
h3 { margin:15px 0 6px; font-size:16px; }
.rank-tables { overflow-x:auto; }
.rank-table { width:100%; border-collapse:collapse; font-variant-numeric:tabular-nums; }
.rank-table th,.rank-table td { padding:6px; border-bottom:1px solid #e4dfeb; }
.rank-table thead { background:#eee9f4; }
.rank-table tbody > tr:nth-child(even) { background:#faf9fc; }
.rank-table tbody > tr[bgcolor] { background:#eee9f4; font-weight:700; }
.rank-table tbody > tr > td:first-child { white-space:nowrap; }
.rank-action { text-align:center; }
.rank-detail > td { white-space:normal !important; padding:12px; background:#f5f2f8; }
.rank-empty { padding:12px; background:#f5f2f8; border-radius:8px; }
@media(max-width:520px) {
  .rank-table th,.rank-table td { padding:5px 3px; }
  .rank-detail .rank-table { min-width:420px; }
}
.rank-jump { display:flex; gap:8px; flex-wrap:wrap; margin:14px 0; }
.rank-jump a { padding:5px 9px; border:1px solid #e4dfeb; border-radius:6px; color:#861451; text-decoration:none; }
@page { size:A4 portrait; margin:10mm; }
@media print {
  body { background:#fff; color:#111; font-size:9pt; }
  .rank-report { margin:0; max-width:none; padding:2mm 3mm; border:0; }
  .rank-tools,.rank-jump,.rank-action,.rank-detail > td > div:last-child { display:none; }
  .rank-header,h2,h3 { break-after:avoid; }
  .rank-logo { width:30mm; background:transparent; filter:brightness(0); }
  .rank-tables { overflow:visible; }
  .rank-table thead { display:table-header-group; }
  .rank-table tr:not(.rank-detail) { break-inside:avoid; }
  .rank-table th,.rank-table td { padding:1mm; }
}

[hidden] { display:none !important; }
.rank-overview-head { display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:10px; margin:16px 0 6px; }
.rank-overview-head h2 { margin:0; }
.rank-overview-head select { display:block; max-width:100%; min-height:36px; margin-top:3px; padding:4px 8px; border:1px solid #d8d0e1; border-radius:6px; background:#fff; color:inherit; font:inherit; }
.rank-scope { margin:0 0 10px; color:#514960; }
.rank-kpis { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; margin:0 0 12px; }
.rank-kpis > div,.rank-chart { border:1px solid #e4dfeb; border-radius:10px; padding:14px; }
.rank-kpis dt { color:#514960; font-size:.88em; }
.rank-kpis dd { margin:5px 0 0; font-size:1.2em; font-weight:700; overflow-wrap:anywhere; }
.rank-charts { display:grid; grid-template-columns:minmax(0,2fr) minmax(0,1fr); gap:12px; }
.rank-chart h3 { margin:0 0 10px; }
.rank-chart ul { list-style:none; padding:0; margin:0; }
.rank-chart li + li { margin-top:9px; }
.rank-bar-label { display:flex; justify-content:space-between; gap:8px; font-size:.9em; }
.rank-bar-label strong { white-space:nowrap; }
.rank-track { height:7px; border-radius:7px; margin-top:3px; background:#eee9f4; overflow:hidden; }
.rank-track span { display:block; height:100%; background:#a31962; }
.rank-donut { display:grid; place-items:center; width:120px; aspect-ratio:1; margin:10px auto; border-radius:50%; background:conic-gradient(#a31962 var(--rank-share,0%),#eee9f4 0); }
.rank-donut strong { display:grid; place-items:center; width:78px; aspect-ratio:1; background:white; border-radius:50%; }
#rank-share-text { text-align:center; margin:8px 0 0; }
.rank-filtered .rank-table:not(.rank-detail .rank-table) > tbody > tr[bgcolor] { display:none; }
@media(max-width:740px) {
  .rank-kpis { grid-template-columns:repeat(2,minmax(0,1fr)); }
  .rank-charts { grid-template-columns:1fr; }
}
@media print {
  .rank-charts,.rank-overview-head label { display:none; }
  .rank-kpis { break-inside:avoid; }
  .rank-kpis > div { padding:2mm; }
}

.rank-section { margin-top:16px; padding:clamp(10px,2vw,18px); border:1px solid #e4dfeb; border-radius:10px; }
.rank-section > h2 { margin:0 0 10px; }
.rank-product-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(min(100%,340px),1fr)); gap:12px; align-items:start; }
.rank-product { min-width:0; padding:12px; border:1px solid #e4dfeb; border-radius:8px; }
.rank-product h3 { margin:0 0 8px; overflow-wrap:anywhere; }
.rank-orders-dialog { width:min(1120px,calc(100% - 20px)); max-height:calc(100dvh - 20px); padding:0; border:1px solid #e4dfeb; border-radius:12px; color:inherit; }
.rank-orders-dialog[open] { display:flex; flex-direction:column; }
.rank-orders-dialog::backdrop { background:#24273899; }
.rank-orders-dialog > header { display:flex; align-items:center; justify-content:space-between; gap:12px; padding:12px 16px; border-bottom:1px solid #e4dfeb; }
.rank-orders-dialog h2,.rank-orders-context { margin:0; }
.rank-orders-context { color:#514960; }
.rank-orders-dialog button { padding:6px 10px; border:1px solid #e4dfeb; border-radius:6px; background:#fff; color:#861451; font:inherit; cursor:pointer; }
.rank-orders-content { overflow:auto; padding:12px 16px; }
.rank-orders-content .rank-table { min-width:480px; }
@media print {
  .rank-orders-dialog { display:none !important; }
  .rank-section,.rank-product { border:0; padding:0; }
  .rank-section + .rank-section { break-before:page; margin-top:0; }
  .rank-product { break-inside:avoid; }
  .rank-product + .rank-product { margin-top:4mm; }
  .rank-product-grid { display:block; }
}

.rank-quantity { margin:0 0 12px; }
.rank-quantity .rank-kpis { grid-template-columns:repeat(auto-fit,minmax(min(100%,135px),1fr)); }
.rank-quantity h4 { margin:0 0 10px; font-size:1em; }
.rank-quantity-share { text-align:center; margin:8px 0 0; }
.rank-product-grid { grid-template-columns:minmax(0,1fr); }
@media print { .rank-quantity { display:none; } }
</style>
</head>
<body>
<main class="rank-report" data-cap-report>
<%If Not rankModal Then%>
<nav class="rank-tools" aria-label="A&ccedil;&otilde;es do relat&oacute;rio">
  <a href="adm_relatorios.asp">Voltar aos relat&oacute;rios</a>
  <button type="button" onclick="window.print()">Imprimir / Salvar PDF</button>
</nav>
<%End If%>
<header class="rank-header">
  <div><h1>Ranking de representantes comerciais</h1>
  <p class="rank-period">Vendas de <%=ArrumaData(DATA_VENDA_INICIO)%> a <%=ArrumaData(DATA_VENDA_FIM)%></p></div>
  <img class="rank-logo" src="images/logoparafundopreto.png" alt="Stillos">
</header>
<section class="rank-dashboard" hidden aria-labelledby="rank-overview">
  <div class="rank-overview-head">
    <h2 id="rank-overview">Vis&atilde;o geral</h2>
    <label>Vendedor <select id="rank-seller"><option value="">Todos os vendedores</option></select></label>
  </div>
  <p id="rank-scope" class="rank-scope" role="status"></p>
  <dl class="rank-kpis">
    <div><dt>Arrecada&ccedil;&atilde;o</dt><dd id="rank-kpi-revenue"></dd></div>
    <div><dt>Vendedores com arrecada&ccedil;&atilde;o registrada</dt><dd id="rank-kpi-sellers"></dd></div>
    <div><dt>&Aacute;lbuns vendidos</dt><dd id="rank-kpi-albums"></dd></div>
    <div><dt id="rank-position-label">L&iacute;der em arrecada&ccedil;&atilde;o</dt><dd id="rank-kpi-position"></dd></div>
  </dl>
  <div class="rank-charts">
    <section class="rank-chart" aria-labelledby="rank-bars-title">
      <h3 id="rank-bars-title">Maiores arrecada&ccedil;&otilde;es</h3>
      <ul id="rank-bars"></ul>
    </section>
    <section class="rank-chart" aria-labelledby="rank-share-title">
      <h3 id="rank-share-title">Participa&ccedil;&atilde;o na arrecada&ccedil;&atilde;o do per&iacute;odo</h3>
      <div class="rank-donut" aria-hidden="true"><strong id="rank-percent"></strong></div>
      <p id="rank-share-text"></p>
    </section>
  </div>
</section>
<nav class="rank-jump" aria-label="Se&ccedil;&otilde;es do ranking" hidden>
  <a href="#rank-overview">Vis&atilde;o geral</a>
</nav>
<div class="rank-tables">
<%Dim rs, sql, rs2, sql2,sql3, asdasd,asd, AlbunsTotais
		 
		 AlbunsTotais = 0

		 
		  	AbrirConexao()		  
 		
		sql2 =  "SELECT " &_ 
        "COD_VENDEDOR, " &_ 
        "NOME_VENDEDOR, " &_ 
        "(SUM(IFNULL(PAGAMENTO_EM_DINHEIRO,0)) + " &_ 
        "SUM(IFNULL(XTOTAL_CHEQUE,0)) + " &_ 
        "SUM(IFNULL(XTOTAL_CARTAO,0))) AS TOTAL_RECEBIDO_GERAL " &_ 
        "FROM VW_EXTRATO_TOTAL_RECEBIDO_COMPLETO P " &_ 
        "WHERE " & sqlQuery & " " &_ 
        "GROUP BY COD_VENDEDOR " &_ 
        "ORDER BY TOTAL_RECEBIDO_GERAL DESC, COD_VENDEDOR"

		



	'Response.Write(sql2)
					
			Set rs2 = objConexao.Execute(sql2)
			Dim num_p, VendaTotal
			VendaTotal = 0
			num_p = 0%>
<%If rs2.Eof then%>
<p class="rank-empty" role="status">Nenhuma arrecada&ccedil;&atilde;o encontrada no per&iacute;odo.</p>
<%rs2.Close : Set rs2 = Nothing
Else%>
	 
	     
	      	  <section class="rank-section" aria-labelledby="rank-revenue"><h2 id="rank-revenue">Arrecada&ccedil;&atilde;o</h2>
<p class="rank-empty" id="rank-empty-revenue" hidden>Sem registros para o vendedor selecionado nesta se&ccedil;&atilde;o.</p>
	      	 		<table class="rank-table"><thead><tr bgcolor="whitesmoke">
                <th><div align="center"><span class="MensagemFicha"><strong>P</strong></span><strong><span class="MensagemFicha">osi&ccedil;&atilde;o</span></strong></div></th>
                <th><div align="left" class="MensagemFicha">
                  <div align="center"><strong>Representante Comercial</strong></div>
                </div></th>
                <th class="rank-action"><div align="center"><strong>Listar Pedidos</strong></div></th>
                <th width="17%">
                  
                  <div align="center" class="MensagemFicha"><strong>Total</strong></div></th></tr></thead><tbody>
		     
			<%While not rs2.Eof
			num_p = num_p + 1
			    %>
                     
				<tr data-rank-kind="revenue" data-seller="<%=rs2("COD_VENDEDOR")%>" data-amount="<%=FormatNumber(rs2("TOTAL_RECEBIDO_GERAL"), 2, -1, 0, 0)%>">
		          <td width="10%"><div align="center"><%=num_p%>&ordm;</div></td>

                  <td width="36%"><div align="left"><%=Server.HTMLEncode(Ucase(rs2("NOME_VENDEDOR") & ""))%></div></td>
<td class="rank-action"><a href="adm_relatorios_ranking_vendedores_completo.asp?<%=Server.HTMLEncode(rankQuery)%>&amp;comando=listarPedidos&amp;cod_vendedor=<%=rs2("COD_VENDEDOR")%>#rank-orders">Ver pedidos do per&iacute;odo</a></td>
                  <td><div align="right" style="padding-right:8px">
				<%
				subValores = CCur(rs2("TOTAL_RECEBIDO_GERAL"))
				Response.Write(FormatNumber(subValores, 2))
				VendaTotal = VendaTotal + subValores
				%>
				</div></td>
              </tr>
                  <%
				  If Request.QueryString("comando") = "listarPedidos" and Request.QueryString("COD_VENDEDOR") = CStr(rs2("COD_VENDEDOR")) Then
				  %>
<tr class="rank-detail" id="rank-orders" data-order-seller="<%=rs2("COD_VENDEDOR")%>"><td colspan="4"><table class="rank-table"><thead><tr bgcolor="whitesmoke">
    <th width="15%"><div align="center"><strong>C&oacute;d. Pedido</strong></div></th>
	<th width="10%"><div align="center"><strong>Tipo Venda</strong></div></th>
   <!-- <th width="15%"><div align="center"><strong>Ajuda de Custo</strong></div></th>-->
    <th width="20%"><div align="center"><strong>Valor Total Pedido</strong></div></th>
    <th width="50%"><div align="center"><strong>Institui&ccedil;&atilde;o de ensino</strong></div></th>
</tr></thead><tbody>

  <%
Dim totalPedidosVendedor
totalPedidosVendedor = 0

 Set rs3 = objConexao.Execute("SELECT " & _
    "P.COD_PEDIDO, " & _
    "PE.TIPO_VENDA, " & _
    "PE.PRE_VENDA, " & _
    "P.VALOR_AJUDA_CUSTO, " & _
    "(IFNULL(P.PAGAMENTO_EM_DINHEIRO,0) + IFNULL(P.XTOTAL_CHEQUE,0) + IFNULL(P.XTOTAL_CARTAO,0)) AS TOTAL_RECEBIDO_GERAL, " & _
    "GROUP_CONCAT(DISTINCT E.DESCRICAO ORDER BY E.DESCRICAO SEPARATOR ' / ') AS INSTITUICAO_ENSINO " & _
    "FROM VW_EXTRATO_TOTAL_RECEBIDO_COMPLETO P " & _
    "LEFT JOIN TB_PEDIDOS_PRODUTOS PP ON P.COD_PEDIDO = PP.COD_PEDIDO " & _
    "LEFT JOIN TB_FICHAS_IDENTIFICACAO F ON PP.COD_FORMANDO = F.COD_FORMANDO " & _
    "LEFT JOIN TB_ENTIDADES E ON F.COD_CONTRATO = E.COD_CONTRATO " & _
    "LEFT JOIN TB_PEDIDOS PE ON PE.COD_PEDIDO = P.COD_PEDIDO " & _
    "WHERE " & sqlQuery & _
    " AND P.COD_VENDEDOR = " & rs2("COD_VENDEDOR") & " " & _
    "GROUP BY P.COD_PEDIDO, PE.TIPO_VENDA, PE.PRE_VENDA, P.VALOR_AJUDA_CUSTO, P.PAGAMENTO_EM_DINHEIRO, P.XTOTAL_CHEQUE, P.XTOTAL_CARTAO " & _
    "ORDER BY P.COD_PEDIDO ASC")
  
  	While not rs3.Eof

	totalPedidosVendedor = totalPedidosVendedor + 1

	TOTAL_RECEBIMENTO_POR_VENDEDOR = formatNumber(0&rs3("TOTAL_RECEBIDO_GERAL"))
  %>
  <tr>
    <td><div align="center"><a <%If rankModal Then%> target="_blank" rel="noopener"<%End If%> href="lca_busca_detalhamento_pedido.asp?cod_pedido=<%=rs3("COD_PEDIDO")%>"><%=rs3("COD_PEDIDO")%></a></div></td>
	<td><div align="center">
        <%
            Dim tp
            If rs3("TIPO_VENDA") = 1 AND rs3("PRE_VENDA") = "N" Then 
                tp = "NORMAL"
            ElseIf rs3("TIPO_VENDA") = 1 AND rs3("PRE_VENDA") = "S" Then
                tp = "PRÉ-VENDA"
            ElseIf rs3("TIPO_VENDA") = 2 Then
                tp = "REPASSE"
            Else
                tp = "-"
            End If

            Response.Write(tp)
        %>
    </div></td>
    <!--<td><div align="right"><%=formatCurrency(rs3("VALOR_AJUDA_CUSTO"))%></div></td>-->
    <td><div align="center"><%=formatCurrency(TOTAL_RECEBIMENTO_POR_VENDEDOR)%></div></td>
	<td><div align="center"><strong><%=Server.HTMLEncode(rs3("INSTITUICAO_ENSINO") & "")%></strong></div></td>
  </tr>
  <%
  rs3.MoveNext()
  Wend
  
  rs3.Close
  Set rs3=Nothing
  %>
  <tr>
    <td bgcolor="whitesmoke" colspan="4" align="center"><strong>Total de pedidos: <%= totalPedidosVendedor %></strong></td>
</tr>
</tbody></table>
               <DIV align="center"><a href="adm_relatorios_ranking_vendedores_completo.asp?<%=Server.HTMLEncode(rankQuery)%>">Ocultar pedidos</a></DIV></td></tr>

<%End If%>
					 
					 
          <%
		  	rs2.MoveNext()
			Wend
			
			rs2.Close
			Set rs2 = Nothing
		  %> 
		  
		  				  <tr bgcolor="whitesmoke"><th class="rank-revenue-total" colspan="3" scope="row">Total arrecadado</th><td><div align="right"><strong><%=formatCurrency(VendaTotal)%></strong></div></td>
	    </tr></tbody></table></section>
<%End If%>
		  <section class="rank-section" aria-labelledby="rank-albums"><h2 id="rank-albums">&Aacute;lbuns vendidos</h2>
<p class="rank-empty" id="rank-empty-albums" hidden>Sem registros para o vendedor selecionado nesta se&ccedil;&atilde;o.</p>
	      	 		<table class="rank-table"><thead><tr bgcolor="whitesmoke">
                <th><div align="center"><span class="MensagemFicha"><strong>P</strong></span><strong><span class="MensagemFicha">osi&ccedil;&atilde;o</span></strong></div></th>
                <th><div align="left" class="MensagemFicha">
                  <div align="center"><strong>Representante Comercial</strong></div>
                </div></th>
                <th width="25%">
                  
                  <div align="center" class="MensagemFicha"><strong>Total</strong></div></th></tr></thead><tbody>
		      <%	 
	
		sql3 =	"SELECT TB_PEDIDOS.COD_VENDEDOR, MAX(TB_VENDEDORES.NOME_VENDEDOR) AS NOME_VENDEDOR, SUM( VW_QTD_FORMANDOS_POR_PEDIDO.QTD_FORMANDOS ) AS TOTAL_ALBUNS " & _
        "FROM VW_QTD_FORMANDOS_POR_PEDIDO " & _
        "INNER JOIN TB_PEDIDOS ON VW_QTD_FORMANDOS_POR_PEDIDO.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO " & _
        "LEFT JOIN TB_VENDEDORES ON TB_VENDEDORES.COD_VENDEDOR = TB_PEDIDOS.COD_VENDEDOR " & _
        "WHERE " & Replace(sqlQuery, "P.", "TB_PEDIDOS.") & _
        " AND TB_PEDIDOS.FINALIZA_CORRECAO = 1 " & _
        "GROUP BY TB_PEDIDOS.COD_VENDEDOR " & _
        "ORDER BY TOTAL_ALBUNS DESC, TB_PEDIDOS.COD_VENDEDOR"

			'Response.Write(sql3)
					
			Set rs3 = objConexao.Execute(sql3)
			Dim num_p2
			num_p2 = 0
			While not rs3.Eof
			num_p2 = num_p2 + 1
			    %>
                     
				<tr data-rank-kind="albums" data-seller="<%=rs3("COD_VENDEDOR")%>" data-amount="<%=FormatNumber(rs3("TOTAL_ALBUNS"), 2, -1, 0, 0)%>">
		          <td width="15%"><div align="center"><%=num_p2%>&ordm;</div></td>
                  <td width="60%"><div align="left"><%=Server.HTMLEncode(UCase(rs3("NOME_VENDEDOR") & ""))%></div></td>
                  <td><div align="center"><%=rs3("TOTAL_ALBUNS")%>
				  <%
				  subAlbuns = CLng(rs3("TOTAL_ALBUNS"))
				  AlbunsTotais = AlbunsTotais + subAlbuns
				  
				  %>
				  </div>
                  <div align="right"></div></td>
              </tr>
					 
					 
          <%
		  	rs3.MoveNext()
			Wend
			
			rs3.Close
			Set rs3=Nothing

			
		  %>  
		  	<tr bgcolor="whitesmoke">
		  				    <td align="center" colspan="2">TOTAL:</td>
                            <td><div align="center"><strong>
                              <%=FormatNumber(AlbunsTotais, 0)%>
                            </strong></div></td>
		  	</tr></tbody></table></section>
<%
AlbunsTotais = 0


%>

    
            
            
            
            
            <%
			sql = "SELECT TB_PEDIDOS_PRODUTOS.COD_PRODUTO, TB_PRODUTOS.NOME_PRODUTO, " & _
      "SUM(TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS TOTAL_QUANTIDADE_GERAL, " & _
      "SUM(TB_PEDIDOS_PRODUTOS.VALOR_UN * TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS VALOR_TOTAL " & _
      "FROM TB_PEDIDOS " & _
      "INNER JOIN TB_PEDIDOS_PRODUTOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO " & _
      "INNER JOIN TB_PRODUTOS ON TB_PEDIDOS_PRODUTOS.COD_PRODUTO = TB_PRODUTOS.COD_PRODUTO " & _
      "WHERE " & Replace(sqlQuery, "P.", "TB_PEDIDOS.") & _
      " AND TB_PEDIDOS.FINALIZA_CORRECAO = 1 " & _
      "GROUP BY TB_PEDIDOS_PRODUTOS.COD_PRODUTO " & _
      "ORDER BY TOTAL_QUANTIDADE_GERAL DESC"

			'Response.Write(sql)
			
			Set rs=objConexao.Execute(sql,QTD_PRODUTOS)
			
			
			
			
				If not rs.Eof Then
				
							%>
<section class="rank-section rank-products" aria-labelledby="rank-products"><h2 id="rank-products">Produtos</h2><div class="rank-product-grid">
<%
While not rs.Eof
																
								%>
                                 <section class="rank-product"><h3><%=Server.HTMLEncode(Ucase(rs("NOME_PRODUTO") & ""))%></h3>
                                    <table class="rank-table"><thead><tr bgcolor="whitesmoke">
                    						<th align="center"><strong>Posi&ccedil;&atilde;o</strong></th>
                    						<th align="center"><strong>Representante Comercial</strong></th>
                    						<th width="25%" align="center"><strong>Total</strong></th></tr></thead><tbody>
                  <%	 
       			
				i = 0
				total = 0
	   
	   			sql = "SELECT TB_PEDIDOS.COD_VENDEDOR AS COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR, " & _
      "SUM(TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS TOTAL_QUANTIDADE, " & _
      "SUM(TB_PEDIDOS_PRODUTOS.VALOR_UN * TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS VALOR_TOTAL " & _
      "FROM TB_PEDIDOS " & _
      "INNER JOIN TB_PEDIDOS_PRODUTOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO " & _
      "INNER JOIN TB_PRODUTOS ON TB_PEDIDOS_PRODUTOS.COD_PRODUTO = TB_PRODUTOS.COD_PRODUTO " & _
      "INNER JOIN TB_VENDEDORES ON TB_PEDIDOS.COD_VENDEDOR = TB_VENDEDORES.COD_VENDEDOR " & _
      "WHERE " & Replace(sqlQuery, "P.", "TB_PEDIDOS.") & _
      " AND TB_PEDIDOS.FINALIZA_CORRECAO = 1 " & _
      " AND TB_PEDIDOS_PRODUTOS.COD_PRODUTO = " & rs("COD_PRODUTO") & _
      " GROUP BY TB_PEDIDOS.COD_VENDEDOR " & _
      " ORDER BY TOTAL_QUANTIDADE DESC, TB_PEDIDOS.COD_VENDEDOR"

				   
				                                
                Set rs2 = objConexao.Execute(sql)

				
                While not rs2.Eof
               		
					i = i + 1
                    subtotal = formatNumber(0&rs2("TOTAL_QUANTIDADE"),0)/1
                    total =  (total+subtotal)/1
                    %>
                         
                    <tr data-rank-kind="product" data-seller="<%=rs2("COD_VENDEDOR")%>" data-amount="<%=FormatNumber(subtotal, 2, -1, 0, 0)%>">
                      <td width="15%"><div align="center"><%=i%>&ordm;</div></td>
                      <td width="60%"><div align="left"><%=Server.HTMLEncode(Ucase(rs2("NOME_VENDEDOR") & ""))%></div></td>
                      <td><div align="center"><%=subtotal%>
                      </div>
                      <div align="right"></div></td>
                  </tr>
                         
                         
				  <%                
                rs2.MoveNext()
                Wend
                
                rs2.Close
                Set rs2=Nothing
              %>  
                <tr bgcolor="whitesmoke">
                                <td align="center" colspan="2">TOTAL:</td>
                                <td><div align="center"><strong>
                                  <%Response.Write(formatNumber(total,0))%>
                                </strong></div></td>
                </tr></tbody></table></section>
                                
                                <%
								
								
							
							rs.MoveNext()
							Wend
				
%></div></section><%
				End if
			
			
			rs.Close
			Set rs=Nothing
			
			%>
            
            
            
<%
			FecharConexao()
%>
</div>
</main>
<template id="rank-quantity-template">
  <div class="rank-quantity">
    <dl class="rank-kpis">
      <div><dt data-metric="total-label"></dt><dd data-metric="total"></dd></div>
      <div><dt>Vendedores com registros</dt><dd data-metric="sellers"></dd></div>
      <div><dt data-metric="position-label"></dt><dd data-metric="position"></dd></div>
    </dl>
    <div class="rank-charts">
      <section class="rank-chart"><h4 data-metric="bars-title"></h4><ul data-metric="bars"></ul></section>
      <section class="rank-chart"><h4 data-metric="share-title"></h4>
        <div class="rank-donut" aria-hidden="true"><strong data-metric="percent"></strong></div>
        <p class="rank-quantity-share" data-metric="share"></p>
      </section>
    </div>
  </div>
</template>
<dialog class="rank-orders-dialog" aria-labelledby="rank-orders-title">
  <header>
    <div><h2 id="rank-orders-title">Pedidos do per&iacute;odo</h2><p class="rank-orders-context"></p></div>
    <form method="dialog"><button autofocus>Fechar</button></form>
  </header>
  <div class="rank-orders-content" aria-live="polite"></div>
</dialog>

<script>
(function () {
    var dashboard = document.querySelector('.rank-dashboard');
    var report = dashboard.closest('main');
    var select = document.getElementById('rank-seller');
    var rows = Array.from(report.querySelectorAll('[data-rank-kind]'));
    var revenue = rows.filter(function (row) { return row.dataset.rankKind === 'revenue'; });
    var albums = rows.filter(function (row) { return row.dataset.rankKind === 'albums'; });
    var detail = document.getElementById('rank-orders');
    var names = new Map();
    var money = new Intl.NumberFormat('pt-BR', {style:'currency',currency:'BRL'});
    var integer = new Intl.NumberFormat('pt-BR');
    var percent = new Intl.NumberFormat('pt-BR', {style:'percent',maximumFractionDigits:1});
    var orderLinks = report.querySelectorAll('.rank-action a, #rank-orders > td > div > a');

    function amount(row) {
        return Number(row.dataset.amount.replace(',', '.'));
    }

    function sum(list) {
        return list.reduce(function (total, row) { return total + amount(row); }, 0);
    }

    rows.forEach(function (row) {
        names.set(row.dataset.seller, row.cells[1].textContent.trim());
    });
    Array.from(names.entries()).sort(function (a,b) {
        return a[1].localeCompare(b[1], 'pt-BR');
    }).forEach(function (entry) { select.add(new Option(entry[1],entry[0])); });

    function renderBars(bars, top, max, formatter) {
        bars.replaceChildren();
        top.forEach(function (row) {
            var item = document.createElement('li');
            var label = document.createElement('div');
            var value = document.createElement('strong');
            var track = document.createElement('div');
            var fill = document.createElement('span');
            label.className = 'rank-bar-label';
            label.append(names.get(row.dataset.seller));
            value.textContent = formatter.format(amount(row));
            label.append(value);
            track.className = 'rank-track';
            track.setAttribute('aria-hidden','true');
            fill.style.width = (max ? Math.max(0,amount(row)) / max * 100 : 0) + '%';
            track.append(fill);
            item.append(label,track);
            bars.append(item);
        });
        if (!top.length) {
            var empty = document.createElement('li');
            empty.textContent = 'Sem registros para exibir.';
            bars.append(empty);
        }
    }

    var quantityTemplate = document.getElementById('rank-quantity-template');
    var quantityPanels = [];
    var albumHeading = document.getElementById('rank-albums');
    if (albumHeading) {
        quantityPanels.push({heading:albumHeading, rows:albums, unit:'\u00c1lbuns vendidos'});
    }
    report.querySelectorAll('.rank-product').forEach(function (product) {
        quantityPanels.push({
            heading:product.querySelector('h3'),
            rows:Array.from(product.querySelectorAll('[data-rank-kind="product"]')),
            unit:'Unidades vendidas'
        });
    });
    quantityPanels.forEach(function (panel) {
        panel.view = quantityTemplate.content.firstElementChild.cloneNode(true);
        panel.heading.after(panel.view);
    });

    function renderQuantities(seller) {
        quantityPanels.forEach(function (panel) {
            var visible = panel.rows.filter(function (row) {
                return !seller || row.dataset.seller === seller;
            });
            var leader = visible[0];
            var total = sum(panel.rows);
            var share = leader && total > 0 ? amount(leader) / total : 0;
            var fields = panel.view.querySelectorAll('[data-metric]');
            var metrics = {
                'total-label':panel.unit,
                total:integer.format(sum(visible)),
                sellers:integer.format(new Set(visible.map(function (row) { return row.dataset.seller; })).size),
                'position-label':seller ? 'Posi\u00e7\u00e3o no ranking' : 'L\u00edder em quantidade',
                position:leader ? (seller ? leader.cells[0].textContent.trim() : names.get(leader.dataset.seller)) : '\u2014',
                'bars-title':seller ? 'Quantidade do vendedor' : 'Maiores quantidades',
                'share-title':'Participa\u00e7\u00e3o no total do per\u00edodo',
                percent:leader && total > 0 ? percent.format(share) : '\u2014',
                share:leader && total > 0 ? percent.format(share) + ' \u00b7 ' + names.get(leader.dataset.seller) : 'Sem participa\u00e7\u00e3o calcul\u00e1vel.'
            };
            fields.forEach(function (field) {
                if (field.dataset.metric !== 'bars') field.textContent = metrics[field.dataset.metric];
            });
            var max = panel.rows.length ? Math.max(0, amount(panel.rows[0])) : 0;
            renderBars(panel.view.querySelector('[data-metric="bars"]'), visible.slice(0,5), max, integer);
            panel.view.querySelector('.rank-donut').style.setProperty('--rank-share', Math.max(0,Math.min(100,share * 100)) + '%');
        });
    }

    function render() {
        var seller = select.value;
        var name = names.get(seller);
        var visibleRevenue = revenue.filter(function (row) { return !seller || row.dataset.seller === seller; });
        var visibleAlbums = albums.filter(function (row) { return !seller || row.dataset.seller === seller; });
        var leader = visibleRevenue[0];
        var total = sum(revenue);
        var share = total > 0 && leader ? amount(leader) / total : 0;
        report.classList.toggle('rank-filtered', Boolean(seller));
        rows.forEach(function (row) { row.hidden = Boolean(seller && row.dataset.seller !== seller); });
        [['revenue', visibleRevenue], ['albums', visibleAlbums]].forEach(function (section) {
            var empty = document.getElementById('rank-empty-' + section[0]);
            if (!empty) return;
            empty.hidden = !seller || section[1].length > 0;
            empty.nextElementSibling.hidden = !empty.hidden;
        });
        if (detail) detail.hidden = Boolean(seller && detail.dataset.orderSeller !== seller);
        report.querySelectorAll('.rank-product').forEach(function (product) {
            product.hidden = Boolean(seller && !product.querySelector('[data-rank-kind]:not([hidden])'));
        });
        var products = document.getElementById('rank-products');
        if (products) products.closest('.rank-products').hidden = Boolean(seller && !report.querySelector('.rank-product:not([hidden])'));
        sectionNav.querySelectorAll('a').forEach(function (link) {
            var heading = document.getElementById(link.hash.slice(1));
            link.hidden = Boolean(heading && heading.closest('[hidden]'));
        });
        renderQuantities(seller);
        orderLinks.forEach(function (link) {
            var url = new URL(link.href);
            if (seller) url.searchParams.set('filtro_vendedor',seller);
            else url.searchParams.delete('filtro_vendedor');
            link.href = url.href;
        });

        document.getElementById('rank-scope').textContent = seller ? name + ' \u2022 Posi\u00e7\u00f5es no ranking geral do per\u00edodo' : 'Todos os vendedores do per\u00edodo';
        document.getElementById('rank-kpi-revenue').textContent = money.format(sum(visibleRevenue));
        document.getElementById('rank-kpi-sellers').textContent = integer.format(visibleRevenue.length);
        document.getElementById('rank-kpi-albums').textContent = integer.format(sum(visibleAlbums));
        document.getElementById('rank-position-label').textContent = seller ? 'Posi\u00e7\u00e3o em arrecada\u00e7\u00e3o' : 'L\u00edder em arrecada\u00e7\u00e3o';
        document.getElementById('rank-kpi-position').textContent = leader ? (seller ? leader.cells[0].textContent.trim() : names.get(leader.dataset.seller)) : '\u2014';

        var bars = document.getElementById('rank-bars');
        var top = visibleRevenue.slice(0,5);
        var max = revenue.length ? Math.max(0,amount(revenue[0])) : 0;
        document.getElementById('rank-bars-title').textContent = seller ? 'Arrecada\u00e7\u00e3o do vendedor' : 'Maiores arrecada\u00e7\u00f5es';
        renderBars(bars, top, max, money);
        var shareText = document.getElementById('rank-share-text');
        shareText.textContent = leader && total > 0 ? percent.format(share) + ' \u00b7 ' + names.get(leader.dataset.seller) : 'Sem participa\u00e7\u00e3o calcul\u00e1vel.';
        document.getElementById('rank-percent').textContent = leader && total > 0 ? percent.format(share) : '\u2014';
        dashboard.querySelector('.rank-donut').style.setProperty('--rank-share', Math.max(0,Math.min(100,share * 100)) + '%');
    }

    var revenueTotal = report.querySelector('.rank-revenue-total');
    var orderDetailCell = detail && detail.cells[0];
    function setPrintColumns(printing) {
        if (revenueTotal) revenueTotal.colSpan = printing ? 2 : 3;
        if (orderDetailCell) orderDetailCell.colSpan = printing ? 3 : 4;
    }
    window.addEventListener('beforeprint', function () { setPrintColumns(true); });
    window.addEventListener('afterprint', function () { setPrintColumns(false); });

    var ordersDialog = document.querySelector('.rank-orders-dialog');
    var ordersContext = ordersDialog.querySelector('.rank-orders-context');
    var ordersContent = ordersDialog.querySelector('.rank-orders-content');
    var orderRequest;
    var orderTrigger;

    report.addEventListener('click', async function (event) {
        var link = event.target.closest('.rank-action a');
        if (!link || event.ctrlKey || event.metaKey || event.shiftKey || event.altKey) return;
        event.preventDefault();
        if (orderRequest) orderRequest.abort();
        var request = new AbortController();
        orderRequest = request;
        orderTrigger = link;
        var seller = link.closest('[data-seller]').dataset.seller;
        ordersContext.textContent = names.get(seller) + ' | ' + report.querySelector('.rank-period').textContent.trim();
        ordersContent.textContent = 'Carregando pedidos...';
        if (!ordersDialog.open) ordersDialog.showModal();

        try {
            var response = await fetch(link.href, {signal:request.signal});
            if (!response.ok) throw new Error('Falha ao carregar pedidos');
            var page = new DOMParser().parseFromString(await response.text(),'text/html');
            var detail = page.getElementById('rank-orders');
            var table = detail && detail.dataset.orderSeller === seller && detail.querySelector('table');
            if (!table) throw new Error('Pedidos indisponiveis');
            if (orderRequest !== request || request.signal.aborted) return;
            table.querySelectorAll('a').forEach(function (order) {
                order.target = '_blank';
                order.rel = 'noopener';
            });
            ordersContent.replaceChildren(table);
        } catch (error) {
            if (request.signal.aborted || orderRequest !== request) return;
            var fallback = document.createElement('a');
            fallback.href = link.href;
            fallback.textContent = 'Abrir listagem de pedidos';
            ordersContent.replaceChildren('N\u00e3o foi poss\u00edvel carregar os pedidos. ',fallback);
        }
    });

    ordersDialog.addEventListener('close', function () {
        if (orderRequest) orderRequest.abort();
        if (orderTrigger) orderTrigger.focus();
    });

    if (!rows.length) return;
    var sectionNav = report.querySelector('.rank-jump');
    report.querySelectorAll('.rank-section > h2[id]').forEach(function (heading) {
        var link = document.createElement('a');
        link.href = '#' + heading.id;
        link.textContent = heading.textContent;
        sectionNav.append(link);
    });
    sectionNav.hidden = false;
    var initial = new URLSearchParams(location.search).get('filtro_vendedor');
    if (names.has(initial)) select.value = initial;
    select.addEventListener('change',render);
    dashboard.hidden = false;
    render();
}());
</script>
</body>
</html>
