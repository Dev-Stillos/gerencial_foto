<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/classe_fechamento.asp" -->
<%
Dim COD_FECHAMENTO_PERIODO
COD_FECHAMENTO_PERIODO	= Request.QueryString("COD_FECHAMENTO_PERIODO")


'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
Dim data_pagamento, data_inicio, data_termino

AbrirConexao()

	Set rs=objConexao.Execute("Select * FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE cod_fechamento_periodo = "&cod_fechamento_periodo)
		If not rs.Eof Then
			data_pagamento	= ArrumaData(rs("data_pagamento"))
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
<title>Folha Geral</title>
</link>
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
<body onLoad="window.print()">
<div align="center"><strong>Reportagens Fotográfica - Resumo Folha de Pagamento</strong><br>
<strong>Pagamento de <%=data_pagamento%> - Vendas de <%=data_inicio%> à <%=data_termino%></strong><br>
<strong style="color:#C00">PROVIS&Oacute;RIO - FECHAMENTO N&Atilde;O FINALIZADO</strong>

</div>


<table width="100%" border="0" cellspacing="0" cellpadding="3">
  <tr>
    <td width="33%"><div align="center" class="style8">Consultor</div></td>
    <td width="14%"><div align="center" class="style8">Comiss&atilde;o</div></td>
    <td width="19%"><div align="center" class="style8">Ajuda de Custo</div></td>
    <td width="15%"><div align="center" class="style8">Pend&ecirc;ncias</div></td>
    <td width="19%"><div align="center" class="style8">Total a Receber</div></td>
  </tr>

<%
AbrirConexao()


Dim rsTotos
Dim NOME_VENDEDOR, COD_VENDEDOR
Dim TOTAL_GERAL_FOLHA_PAGAMENTO, TOTAL_A_RECEBER_POR_CONSULTOR
Dim TOTAL_COMISSAO, TOTAL_AJUDA_CUSTO, TOTAL_PENDENCIA_PAGAMENTO
TOTAL_A_RECEBER_POR_CONSULTOR = 0
TOTAL_GERAL_FOLHA_PAGAMENTO = 0

			'Obs.: Nos casos dos vendedores coordenadores 7 e 14 os mesmos dever�o aparecer mesmo que não tenham venda
				
			sql = "SELECT TB_PEDIDOS.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR FROM TB_VENDEDORES LEFT JOIN TB_PEDIDOS ON (TB_VENDEDORES.COD_VENDEDOR = TB_PEDIDOS.COD_VENDEDOR) WHERE TB_PEDIDOS.finaliza_correcao = 1 AND  TB_PEDIDOS.cod_fechamento_periodo = "&cod_fechamento_periodo&" AND TB_VENDEDORES.COD_VENDEDOR <> 27 GROUP BY TB_PEDIDOS.COD_VENDEDOR ORDER BY TB_VENDEDORES.NOME_VENDEDOR "	
																
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
    <td class="style9" style="border-bottom:solid 1px #999999"><%=NOME_VENDEDOR%></td>
    <td class="style9" style="border-bottom:solid 1px #999999"><div align="right"><%=formatNumber(TOTAL_COMISSAO)%></div></td>
    <td class="style9" style="border-bottom:solid 1px #999999"><div align="right"><%=formatNumber(TOTAL_AJUDA_CUSTO)%></div></td>
    <td class="style9" style="border-bottom:solid 1px #999999"><div align="right"><%=formatNumber(TOTAL_PENDENCIA_PAGAMENTO)%></div></td>
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
    <td colspan="4" style="border-top:solid 1px #333333">TOTAL:</td>
    <td style="border-top:solid 1px #333333"><div align="right"><strong><%=formatCurrency(TOTAL_GERAL_FOLHA_PAGAMENTO)%></strong></div></td>
  </tr>
</table>

<div style="font-size:12px; text-align:center"><em>Observações:<br>
Para agilizar o processamento as colunas Total Coordenador e Total Metas não serão calculadas. Vendas conferidas pelo dpto administrativo fotográfico.</em><br>
</div>
</body>
</html>



