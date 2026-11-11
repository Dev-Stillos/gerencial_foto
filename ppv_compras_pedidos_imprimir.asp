<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<!-- #include file="files/includes/functions_local.asp" -->
<%
Dim cod_pedido_compra, numero_pedido_compra, cod_fornecedor, nome_fantasia, aprovado_cod_usuario, data_aprovado
Dim qtd_itens_total

cod_pedido_compra = Request.QueryString("c")

Dim modoModalCompra, codigoCompraValido
modoModalCompra = (Request.QueryString("modal") = "1")
Set codigoCompraValido = New RegExp
codigoCompraValido.Pattern = "^[0-9]{1,10}$"
If Not codigoCompraValido.Test(cod_pedido_compra & "") Then
    Response.Status = "400 Bad Request"
    Response.Write "C&oacute;digo de pedido inv&aacute;lido."
    Response.End
End If

AbrirConexao()

	Set rs=objConexao.Execute("SELECT * FROM TB_PEDIDOS_COMPRAS WHERE cod_pedido_compra = "&cod_pedido_compra)
			
			If not rs.Eof Then
				
				numero_pedido_compra=rs("numero_pedido_compra")
				cod_fornecedor=rs("cod_fornecedor")
				aprovado_cod_usuario=rs("aprovado_cod_usuario")
				data_aprovado=rs("data_aprovado")
				valor_total_produtos = pedidoCompraTotais(cod_pedido_compra,1,true)
				contato=rs("contato")
				pedido_manual = "S"
				
				
				Set rs2=objConexao.Execute("SELECT nome_fantasia FROM TB_FORNECEDORES WHERE cod_fornecedor = '"&cod_fornecedor&"'")
					
					If not rs2.Eof Then
						nome_fantasia = rs2("nome_fantasia")
					End If
				
				rs2.Close
				Set rs2=Nothing
							
			End If
	

If rs.Eof Then
    rs.Close
    Set rs = Nothing
    FecharConexao()
    Response.Status = "404 Not Found"
    Response.Write "Pedido de compra n&atilde;o encontrado."
    Response.End
End If
rs.Close
Set rs = Nothing

Dim subtotal, total
subtotal = 0
total = 0
%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Pedido de compra <%=Server.HTMLEncode(numero_pedido_compra & "")%> | Stillos</title>
<style>
* { box-sizing:border-box; }
body { margin:0; background:#f6f7fb; color:#242738; font:14px/1.45 system-ui,Arial,sans-serif; }
.report-toolbar { max-width:1080px; margin:16px auto; padding:0 16px; text-align:right; }
.report-toolbar button { background:#b20b61; color:white; border:0; border-radius:7px; padding:10px 14px; font:inherit; cursor:pointer; }
.report { max-width:1080px; margin:16px auto; padding:clamp(14px,4vw,30px); background:#fff; border:1px solid #e4dfeb; border-radius:12px; }
.report-header { display:flex; justify-content:space-between; align-items:center; gap:12px; border-bottom:2px solid #b20b61; padding-bottom:12px; }
.report-header img { width:60px; height:auto; }
h1 { font-size:clamp(19px,3vw,29px); margin:0; }
h2 { font-size:1.05rem; margin:0 0 8px; }
.report-meta { display:grid; grid-template-columns:2fr 1fr 1fr; gap:10px; padding:12px; background:#f7f5fa; border-radius:8px; }
dt { color:#57566a; font-size:.85em; font-weight:600; }
dd { margin:2px 0 0; overflow-wrap:anywhere; }
.report-group { margin:18px 0; }
.report-table { width:100%; border-collapse:collapse; font-size:.92em; }
.report-table th, .report-table td { padding:7px 6px; border-bottom:1px solid #e4dfeb; text-align:left; vertical-align:top; overflow-wrap:anywhere; }
.report-table thead th { background:#f0edf6; }
.report-table .amount { text-align:right; white-space:nowrap; }
.report-totals { display:flex; justify-content:flex-end; flex-wrap:wrap; gap:12px; padding:12px; border:1px solid #e4dfeb; border-radius:8px; }
.report-totals > div { min-width:135px; }
.report-totals dd { font-size:1.14em; font-weight:700; }
.report-signatures { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:14px; margin-top:26px; }
.report-signatures > div { border-top:1px solid #9794a0; padding-top:6px; }
@media screen and (max-width:620px) {
  .report { margin:0; border:0; }
  .report-meta, .report-signatures { grid-template-columns:1fr; }
  .report-table, .report-table tbody, .report-table tfoot { display:block; }
  .report-table thead { position:absolute; width:1px; height:1px; overflow:hidden; clip-path:inset(50%); }
  .report-table tr { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); padding:6px 0; border-bottom:1px solid #e4dfeb; }
  .report-table td, .report-table tfoot th { display:block; border:0; text-align:left; }
  .report-table td::before { content:attr(data-label); display:block; font-size:.8em; color:#57566a; font-weight:650; }
}
@page { size:A4 portrait; margin:12mm; }
@media print {
  body { background:#fff; color:#111; font-size:9pt; }
  .report-toolbar { display:none; }
  .report { max-width:none; margin:0; padding:0; border:0; }
  .report-table thead { display:table-header-group; }
  .report-table tr, .report-signatures, .report-totals { break-inside:avoid; }
  h2 { break-after:avoid; }
}
</style>
</head>
<body>
<%If Not modoModalCompra Then%>
<div class="report-toolbar"><button type="button" onclick="window.print()">Imprimir / Salvar PDF</button></div>
<%End If%>
<main class="report" data-purchase-report>
<header class="report-header">
    <h1>Pedido de compra n&ordm; <%=Server.HTMLEncode(numero_pedido_compra & "")%></h1>
    <img src="images/layout/logomarca_foto_reportagens.jpg" alt="Stillos">
</header>
<dl class="report-meta">
    <div><dt>Fornecedor</dt><dd><%=Server.HTMLEncode(UCase(nome_fantasia & ""))%></dd></div>
    <div><dt>Data do pedido</dt><dd><%=ArrumaData(data_aprovado)%></dd></div>
    <div><dt>Contato</dt><dd><%=Server.HTMLEncode(contato & "")%></dd></div>
</dl>
<%
if  moduloPPVpedidoManual(cod_pedido_compra, "", true) Then
										set rs = objConexao.execute("SELECT	SUM(quantidade_solicitada), SUM(quantidade_solicitada*valor_custo_solicitado) FROM VW_TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra = "&sqlNumero(cod_pedido_compra))
											if not rs.Eof Then
												qtd_itens_total = rs(0)
												total = rs(1)
											end if
										rs.close
										set rs = nothing


										set rs = objConexao.execute("SELECT *, quantidade_solicitada * valor_custo_solicitado as valor_custo_total FROM VW_TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra = "&sqlNumero(cod_pedido_compra))
%>
<section class="report-group">
<h2>Itens do pedido manual</h2>
<table class="report-table" role="table">
<thead><tr><th>Produto</th><th class="amount">Qtd.</th><th class="amount">Valor un.</th><th class="amount">Valor total</th><th>Observa&ccedil;&atilde;o</th></tr></thead>
<tbody>
<%While Not rs.Eof%>
<tr>
<td data-label="Produto"><%=Server.HTMLEncode(rs("nome_produto") & "")%></td>
<td data-label="Qtd." class="amount"><%=rs("quantidade_solicitada")%></td>
<td data-label="Valor un." class="amount">R$ <%=ArrumaNumero(rs("valor_custo_solicitado"),0,2)%></td>
<td data-label="Valor total" class="amount">R$ <%=ArrumaNumero(rs("valor_custo_total"),0,2)%></td>
<td data-label="Observa&ccedil;&atilde;o"><%=Server.HTMLEncode(rs("obs") & "")%></td>
</tr>
<%rs.MoveNext
Wend
rs.Close
Set rs = Nothing
%>
</tbody></table>
</section>
<%Else
    Set rs=objConexao.Execute("SELECT VW_TB_ORDENS_SERVICO.COD_CONTRATO, VW_TB_ORDENS_SERVICO.ENTIDADE_ENSINO FROM VW_TB_ORDENS_SERVICO INNER JOIN VW_TB_PEDIDOS_COMPRAS_PRODUTOS on VW_TB_ORDENS_SERVICO.COD_ORDEM_SERVICO = VW_TB_PEDIDOS_COMPRAS_PRODUTOS.COD_ORDEM_SERVICO WHERE VW_TB_PEDIDOS_COMPRAS_PRODUTOS.cod_pedido_compra = "&sqlTexto(cod_pedido_compra)&" GROUP BY VW_TB_ORDENS_SERVICO.COD_CONTRATO ORDER BY VW_TB_ORDENS_SERVICO.ENTIDADE_ENSINO ASC")
    While Not rs.Eof
%>
<section class="report-group">
<h2><%=Server.HTMLEncode(UCase(rs("ENTIDADE_ENSINO") & ""))%></h2>
<table class="report-table" role="table">
<thead><tr><th>OS</th><th>Formando</th><th>Produto</th><th class="amount">Qtd.</th><th>Entrega</th><th class="amount">Valor un.</th><th class="amount">Valor total</th></tr></thead>
<tbody>
<%
    subtotal = 0
    Set rs2=objConexao.Execute("SELECT VW_TB_ORDENS_SERVICO.NOME_FORMANDO, VW_TB_PEDIDOS_COMPRAS_PRODUTOS.* FROM VW_TB_ORDENS_SERVICO right JOIN VW_TB_PEDIDOS_COMPRAS_PRODUTOS on VW_TB_ORDENS_SERVICO.COD_ORDEM_SERVICO = VW_TB_PEDIDOS_COMPRAS_PRODUTOS.COD_ORDEM_SERVICO WHERE VW_TB_ORDENS_SERVICO.COD_CONTRATO = "&sqlTexto(rs("COD_CONTRATO"))&" AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.cod_pedido_compra = "&cod_pedido_compra&" ORDER BY VW_TB_ORDENS_SERVICO.ENTIDADE_ENSINO ASC")
    While Not rs2.Eof
        subtotal = (subtotal+formatNumber(0&rs2("valor_total_item_solicitado")))/1
        qtd_itens_total = qtd_itens_total+formatNumber(0&rs2("quantidade_solicitada"))/1
%>
<tr>
<td data-label="OS"><%=moduloPPVOSMascara(rs2("cod_ordem_servico"),"",true)%></td>
<td data-label="Formando"><%=Server.HTMLEncode(rs2("nome_formando") & "")%></td>
<td data-label="Produto"><%=Server.HTMLEncode(rs2("nome_produto") & "")%></td>
<td data-label="Qtd." class="amount"><%=rs2("quantidade_solicitada")%></td>
<td data-label="Entrega"><%=ArrumaData(rs2("data_previsao_entrega"))%></td>
<td data-label="Valor un." class="amount"><%=formatNumber(rs2("valor_custo_solicitado"),2)%></td>
<td data-label="Valor total" class="amount"><%=formatNumber(rs2("valor_total_item_solicitado"),2)%></td>
</tr>
<%
    rs2.MoveNext
    Wend
    rs2.Close
    Set rs2 = Nothing
    total = (total+subtotal)/1
%>
</tbody>
<tfoot><tr><th colspan="6">Subtotal</th><td class="amount"><%=formatNumber(subtotal,2)%></td></tr></tfoot>
</table>
</section>
<%
    rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
End If
%>
<dl class="report-totals">
    <div><dt>Quantidade de itens</dt><dd><%=ArrumaNumero(qtd_itens_total,0,0)%></dd></div>
    <div><dt>Valor total</dt><dd>R$ <%=ArrumaNumero(total,0,2)%></dd></div>
</dl>
<dl class="report-signatures">
    <div><dt>Fornecedor</dt><dd><%=Server.HTMLEncode(UCase(nome_fantasia & ""))%></dd></div>
    <div><dt>Solicitado por</dt><dd><%=nomeUsuario(Session("cod_usuario"),true)%></dd></div>
    <div><dt>Aprovado por</dt><dd><%=nomeUsuario(aprovado_cod_usuario,true)%></dd></div>
</dl>
</main>
<%FecharConexao()%>
<%If Not modoModalCompra Then%><script>window.addEventListener('load', function () { window.print(); });</script><%End If%>
</body>
</html>
