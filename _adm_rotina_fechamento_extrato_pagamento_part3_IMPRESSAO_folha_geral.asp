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
<strong>Pagamento de <%=data_pagamento%> - Vendas de <%=data_inicio%> à <%=data_termino%></strong></div>

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

			'Obs.: Nos casos dos vendedores coordenadores 7 e 14 os mesmos deverão aparecer mesmo que não tenham venda
					
			sql = "SELECT "&_
											"TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR ,"&_
											"TB_FECHAMENTO_PAGAMENTO.COD_GRUPO_VENDEDORES,"&_
											"TB_VENDEDORES.NOME_VENDEDOR "&_
											"FROM TB_VENDEDORES INNER JOIN TB_FECHAMENTO_PAGAMENTO ON TB_VENDEDORES.COD_VENDEDOR = TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR "&_
											"WHERE TB_FECHAMENTO_PAGAMENTO.COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR <> 27 OR TB_VENDEDORES.COD_VENDEDOR IN (SELECT cod_vendedor FROM TB_FECHAMENTO_PENDENCIAS_PAGAMENTO WHERE cod_fechamento_periodo = '"&COD_FECHAMENTO_PERIODO&"') or TB_VENDEDORES.COD_VENDEDOR IN(7,14,48) "&_
											"GROUP BY TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR ORDER BY NOME_VENDEDOR"
											
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
    <td class="style9" style="border-bottom:solid 1px #999999"><div align="right"><%=formatNumber(0&fechamentoRetornaValorRecebidoVendedorTipo(1,cod_vendedor,cod_fechamento_periodo,true))%></div></td>
    <td class="style9" style="border-bottom:solid 1px #999999"><div align="right"><%=formatNumber(0&fechamentoRetornaValorRecebidoVendedorTipo(2,cod_vendedor,cod_fechamento_periodo,true))%></div></td>
    <td class="style9" style="border-bottom:solid 1px #999999"><div align="right"><%=fechamentoRetornaValorRecebidoVendedorTipo(5,cod_vendedor,cod_fechamento_periodo,true)%></div></td>
    <td class="style9" style="border-bottom:solid 1px #999999"><div align="right"><%=formatNumber(0&fechamentoRetornaValorRecebidoVendedorTipo(4,cod_vendedor,cod_fechamento_periodo,true))%></div></td>
    <td class="style9" style="border-bottom:solid 1px #999999"><div align="right"><%=formatNumber(0&fechamentoRetornaValorRecebidoVendedorTipo(3,cod_vendedor,cod_fechamento_periodo,true))%></div></td>
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
    <td colspan="6" style="border-top:solid 1px #333333"> </td>
    <td style="border-top:solid 1px #333333"><div align="right"><strong><%=formatCurrency(TOTAL_GERAL_FOLHA_PAGAMENTO)%></strong></div></td>
  </tr>
</table>
</body>
</html>



