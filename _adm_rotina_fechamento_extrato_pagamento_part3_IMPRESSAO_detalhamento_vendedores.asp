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
		pedidosFormataFormaPagamento = "� vista" 
	Elseif int(FORMA_PAGAMENTO) = 2 Then
		pedidosFormataFormaPagamento = "� prazo"
	End If

end Function



Dim COD_VENDEDOR, ESCOLHA_EXTRATO, COD_FECHAMENTO_PERIODO, NOME_VENDEDOR, TOTAL_BOLETO
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
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="all"></link>
<title>Extrato dos Vendedores</title>


<style type="text/css" media="all">
<!--
.folha2 {
    page-break-after: always;
}
-->
</style></head>

<body bgcolor="#FFFFFF" style="background-color:#FFFFFF">
<%
AbrirConexao()
'//////////////////////////////////////////////////////////////

VERIFICA_COMISSAO = 0

'oBS.: Nos casos dos vendedores coordenadores 7 e 14 os mesmos dever�o aparecer mesmo que n�o tenham venda
Set rsMaster = objConexao.Execute("SELECT "&_
									"TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR ,"&_
									"TB_VENDEDORES.NOME_VENDEDOR "&_
									"FROM TB_VENDEDORES INNER JOIN TB_FECHAMENTO_PAGAMENTO ON TB_VENDEDORES.COD_VENDEDOR = TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR "&_
									"WHERE (TB_FECHAMENTO_PAGAMENTO.COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR <> 27) OR TB_VENDEDORES.COD_VENDEDOR IN (SELECT cod_vendedor FROM TB_FECHAMENTO_PENDENCIAS_PAGAMENTO WHERE cod_fechamento_periodo = '"&COD_FECHAMENTO_PERIODO&"') or TB_VENDEDORES.COD_VENDEDOR IN(7,14,48) "&_
									"GROUP BY TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR ORDER BY NOME_VENDEDOR")
If not rsMaster.Eof Then
	While not rsMaster.Eof
		COD_VENDEDOR 	= rsMaster("COD_VENDEDOR")
		NOME_VENDEDOR 	= rsMaster("NOME_VENDEDOR")
		%>

				<table style="margin-top:10px" border="0" cellpadding="5" cellspacing="0" align="center" width="100%">
				<tr>
					<td><div align="center"><span class="textoMaior"><strong>Data do Pagamento <%=ArrumaData(data_pagamento)%> - Per�odo de Vendas de <%=data_inicio%> � <%=data_termino%></strong></span></div></td>
				</tr>
				</table>
				<table style="margin-top:10px" cellpadding="5" cellspacing="0" align="center" width="100%">
				<tr>
					<td><div align="left"><span class="textoMedio"><strong><%=Ucase(NOME_VENDEDOR)%></strong></span></div></td>
				</tr>
				</table>
				<%
				'1 - VENDAS NORMAIS
				Set rs = objConexao.Execute("SELECT * FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND COD_VENDEDOR = "&COD_VENDEDOR&" AND TIPO_VENDA = 1",TOTAL_PEDIDOS_TIPO_VENDA_1)
				If not rs.Eof Then
					%>
					<table width="100%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
					<tr bgcolor="whitesmoke">
						<td height="0" colspan="13"><div align="left"><strong>1 - VENDAS NORMAIS</strong></div></td>
					</tr>	
					<tr bgcolor="whitesmoke">
						<td width="5%" height="30"><div align="center"><strong>C�d. Pedido </strong></div></td>
						<td width="7%"><div align="center"><strong>Data Venda </strong></div></td>
						<td width="5%"><div align="center"><strong>Forma Pag.</strong></div></td>
						<td width="8%"><div align="center"><strong>Dinheiro</strong></div></td>
						<td width="8%"><div align="center"><strong>Cheque</strong></div></td>
						<td width="8%"><div align="center"><strong>Cart�o</strong></div></td>
						<td width="8%"><div align="center"><strong>Boleto</strong></div></td>
						<td width="5%"><div align="center"><strong>�lbuns</strong></div></td>
						<td width="5%"><div align="center"><strong>DVDs</strong></div></td>
						<td width="6%"><div align="center"><strong>Foto-Telas e Trios </strong></div></td>
						<td width="3%"><div align="center"><strong>Ajuda de Custo </strong></div></td>
						<td width="4%"><div align="center"><strong>Comiss&atilde;o</strong></div></td>
						<td width="36%"><div align="center"><strong>C�d. Formando(s) / �lbum / Institui��o de Ensino </strong></div></td>
					</tr>
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
						QTD_ALBUNS			= rs("QTD_ALBUNS")
						QTD_DVDS			= rs("QTD_DVDS")
						QTD_FOTO_TELAS_TRIOS= rs("QTD_FOTO_TELAS_TRIOS")
						VALOR_TOTAL_COMISSAO= formatNumber(0&rs("VALOR_TOTAL_COMISSAO"))/1
						%>
						<tr>
							<td><div align="center"><%=COD_PEDIDO%></div></td>
							<td><div align="center"><%=ArrumaData(DATA_VENDA)%></div></td>
							<td><div align="center"><%=pedidosFormataFormaPagamento(FORMA_PAGAMENTO)%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(TOTAL_DINHEIRO))%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(TOTAL_CHEQUE))%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(TOTAL_CARTAO))%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(TOTAL_BOLETO))%></div></td>
							<td><div align="center"><%=retiraZero2(QTD_ALBUNS)%></div></td>
							<td><div align="center"><%=retiraZero2(QTD_DVDS)%></div></td>
							<td><div align="center"><%=retiraZero2(QTD_FOTO_TELAS_TRIOS)%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(VALOR_AJUDA_CUSTO))%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(VALOR_TOTAL_COMISSAO))%></div></td>
							<td><div align="center"><%=pedidosFormandosFormatado(COD_PEDIDO, true)%></div></td>
						</tr>
						<%
					rs.MoveNext()
					Wend
					
					'TOTAL 1 - VENDAS NORMAIS
					Set rsTotal = objConexao.Execute("SELECT COUNT(*) AS XPEDIDOS, SUM(TOTAL_DINHEIRO) AS XTOTAL_DINHEIRO, SUM(TOTAL_BOLETO) AS XTOTAL_BOLETO, SUM(TOTAL_CHEQUE)AS XTOTAL_CHEQUE, SUM(TOTAL_CARTAO)AS XTOTAL_CARTAO, SUM(TOTAL_RECEBIMENTO) AS XTOTAL_RECEBIMENTO, SUM(QTD_ALBUNS) AS XQTD_ALBUNS, SUM(QTD_DVDS) AS XQTD_DVDS, SUM(QTD_FOTO_TELAS_TRIOS) AS XQTD_FOTO_TELAS_TRIOS , SUM(VALOR_AJUDA_CUSTO) AS XVALOR_AJUDA_CUSTO, SUM(VALOR_TOTAL_COMISSAO) AS XVALOR_TOTAL_COMISSAO FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_VENDEDOR = "&COD_VENDEDOR&" AND TIPO_VENDA = 1 AND COD_FECHAMENTO_PERIODO ="&COD_FECHAMENTO_PERIODO)
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
						XVALOR_TOTAL_COMISSAO   = formatNumber(0&rsTotal("XVALOR_TOTAL_COMISSAO"))/1
					End If
					rsTotal.Close
					Set rsTotal = Nothing
					%>
					<tr bgcolor="whitesmoke">
						<td><div align="center"><strong><%=XPEDIDOS%></strong></div></td>
						<td><div align="center"></div></td>
						<td>&nbsp;</td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XTOTAL_DINHEIRO))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XTOTAL_CHEQUE))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XTOTAL_CARTAO))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XTOTAL_BOLETO))%></strong></div></td>
						<td><div align="center"><strong><%=XQTD_ALBUNS%></strong></div></td>
						<td><div align="center"><strong><%=XQTD_DVDS%></strong></div></td>
						<td><div align="center"><strong><%=XQTD_FOTO_TELAS_TRIOS%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XVALOR_AJUDA_CUSTO))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XVALOR_TOTAL_COMISSAO))%></strong></div></td>
						<td>&nbsp;</td>
					</tr>
					</table>
         		<br />
					<%		
				End If
				rs.Close
				Set rs = Nothing
				
				'2 - VENDAS REPASSE
				Set rs = objConexao.Execute("SELECT * FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND COD_VENDEDOR = "&COD_VENDEDOR&" AND TIPO_VENDA = 2",TOTAL_PEDIDOS_TIPO_VENDA_2)
				If not rs.Eof Then
					%>
   	  			<table width="100%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
					<tr bgcolor="whitesmoke">
						<td height="0" colspan="13"><div align="left"><strong>2 - VENDAS REPASSE</strong></div></td>
					</tr>
					<tr bgcolor="whitesmoke">
						<td width="5%" height="30"><div align="center"><strong>C�d. Pedido </strong></div></td>
						<td width="7%"><div align="center"><strong>Data Venda </strong></div></td>
						<td width="5%"><div align="center"><strong>Forma Pag.</strong></div></td>
						<td width="8%"><div align="center"><strong>Dinheiro</strong></div></td>
						<td width="8%"><div align="center"><strong>Cheque</strong></div></td>
						<td width="8%"><div align="center"><strong>Cart�o</strong></div></td>
						<td width="5%"><div align="center"><strong>�lbuns</strong></div></td>
						<td width="5%"><div align="center"><strong>DVDs</strong></div></td>
						<td width="6%"><div align="center"><strong>Foto-Telas e Trios </strong></div></td>
						<td width="3%"><div align="center"><strong>Ajuda de Custo </strong></div></td>
						<td width="4%"><div align="center"><strong>Comiss&atilde;o</strong></div></td>
						<td width="36%"><div align="center"><strong>C�d. Formando(s) / �lbum / Institui��o de Ensino </strong></div></td>
					</tr>
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
						QTD_ALBUNS				= rs("QTD_ALBUNS")
						QTD_DVDS					= rs("QTD_DVDS")
						QTD_FOTO_TELAS_TRIOS	= rs("QTD_FOTO_TELAS_TRIOS")
						VALOR_TOTAL_COMISSAO	= rs("VALOR_TOTAL_COMISSAO")
						%>
						<tr>
							<td><div align="center"><%=COD_PEDIDO%></div></td>
							<td><div align="center"><%=ArrumaData(DATA_VENDA)%></div></td>
							<td><div align="center"><%=pedidosFormataFormaPagamento(FORMA_PAGAMENTO)%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(TOTAL_DINHEIRO))%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(TOTAL_CHEQUE))%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(TOTAL_CARTAO))%></div></td>
							<td><div align="center"><%=retiraZero2(QTD_ALBUNS)%></div></td>
							<td><div align="center"><%=retiraZero2(QTD_DVDS)%></div></td>
							<td><div align="center"><%=retiraZero2(QTD_FOTO_TELAS_TRIOS)%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(VALOR_AJUDA_CUSTO))%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(VALOR_TOTAL_COMISSAO))%></div></td>
							<td><div align="center"><%=pedidosFormandosFormatado(COD_PEDIDO, true)%></div></td>
						</tr>
						<%
					rs.MoveNext()
					Wend

					'2 - VENDAS REPASSE TOTAL
					Set rsTotal = objConexao.Execute("SELECT COUNT(*) AS XPEDIDOS, SUM(TOTAL_DINHEIRO) AS XTOTAL_DINHEIRO, SUM(TOTAL_CHEQUE)AS XTOTAL_CHEQUE, SUM(TOTAL_CARTAO)AS XTOTAL_CARTAO, SUM(TOTAL_RECEBIMENTO) AS XTOTAL_RECEBIMENTO, SUM(QTD_ALBUNS) AS XQTD_ALBUNS, SUM(QTD_DVDS) AS XQTD_DVDS, SUM(QTD_FOTO_TELAS_TRIOS) AS XQTD_FOTO_TELAS_TRIOS , SUM(VALOR_AJUDA_CUSTO) AS XVALOR_AJUDA_CUSTO, SUM(VALOR_TOTAL_COMISSAO) AS XVALOR_TOTAL_COMISSAO FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_VENDEDOR = "&COD_VENDEDOR&" AND TIPO_VENDA = 2 AND COD_FECHAMENTO_PERIODO ="&COD_FECHAMENTO_PERIODO)
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
					<tr bgcolor="whitesmoke">
						<td><div align="center"><strong><%=XPEDIDOS%></strong></div></td>
						<td><div align="center"></div></td>
						<td>&nbsp;</td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XTOTAL_DINHEIRO))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XTOTAL_CHEQUE))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XTOTAL_CARTAO))%></strong></div></td>
						<td><div align="center"><strong><%=XQTD_ALBUNS%></strong></div></td>
						<td><div align="center"><strong><%=XQTD_DVDS%></strong></div></td>
						<td><div align="center"><strong><%=XQTD_FOTO_TELAS_TRIOS%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XVALOR_AJUDA_CUSTO))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XVALOR_TOTAL_COMISSAO))%></strong></div></td>
						<td>&nbsp;</td>
					</tr>
					</table>

					<br />
					<%
				End If
				rs.Close
				Set rs = Nothing

				'3 - VENDAS REPASSE (GRUPO ESPECIAL)
				Set rs = objConexao.Execute("SELECT * FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND COD_VENDEDOR = "&COD_VENDEDOR&" AND TIPO_VENDA = 3",TOTAL_PEDIDOS_TIPO_VENDA_3)
				If not rs.Eof Then
					%>
					<table width="100%" border="1" align="center"	cellpadding="5" style="border-collapse:collapse">
					<tr bgcolor="whitesmoke">
						<td height="0" colspan="13"><div align="left"><strong>3 - VENDAS REPASSE</strong> <strong>GRUPO ESPECIAL</strong></div></td>
					</tr>
					<tr bgcolor="whitesmoke">
						<td width="5%" height="30"><div align="center"><strong>C�d. Pedido </strong></div></td>
						<td width="7%"><div align="center"><strong>Data Venda </strong></div></td>
						<td width="5%"><div align="center"><strong>Forma Pag.</strong></div></td>
						<td width="8%"><div align="center"><strong>Dinheiro</strong></div></td>
						<td width="8%"><div align="center"><strong>Cheque</strong></div></td>
						<td width="8%"><div align="center"><strong>Cart�o</strong></div></td>
						<td width="5%"><div align="center"><strong>�lbuns</strong></div></td>
						<td width="5%"><div align="center"><strong>DVDs</strong></div></td>
						<td width="6%"><div align="center"><strong>Foto-Telas e Trios </strong></div></td>
						<td width="3%"><div align="center"><strong>Ajuda de Custo </strong></div></td>
						<td width="4%"><div align="center"><strong>Comiss&atilde;o</strong></div></td>
						<td width="36%"><div align="center"><strong>C�d. Formando(s) / �lbum / Institui��o de Ensino </strong></div></td>
					</tr>
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
						QTD_ALBUNS				= rs("QTD_ALBUNS")
						QTD_DVDS					= rs("QTD_DVDS")
						QTD_FOTO_TELAS_TRIOS = rs("QTD_FOTO_TELAS_TRIOS")
						VALOR_TOTAL_COMISSAO = rs("VALOR_TOTAL_COMISSAO")
						%>
						<tr>
							<td><div align="center"><%=COD_PEDIDO%></div></td>
							<td><div align="center"><%=ArrumaData(DATA_VENDA)%></div></td>
							<td><div align="center"><%=pedidosFormataFormaPagamento(FORMA_PAGAMENTO)%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(TOTAL_DINHEIRO))%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(TOTAL_CHEQUE))%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(TOTAL_CARTAO))%></div></td>
							<td><div align="center"><%=retiraZero2(QTD_ALBUNS)%></div></td>
							<td><div align="center"><%=retiraZero2(QTD_DVDS)%></div></td>
							<td><div align="center"><%=retiraZero2(QTD_FOTO_TELAS_TRIOS)%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(VALOR_AJUDA_CUSTO))%></div></td>
							<td><div align="right"><%=retiraZero2(formatNumber(VALOR_TOTAL_COMISSAO))%></div></td>
							<td><div align="center"><%=pedidosFormandosFormatado(COD_PEDIDO, true)%></div></td>
						</tr>
						<%
					rs.MoveNext()
					Wend

					'3 - VENDAS REPASSE (GRUPO ESPECIAL) TOTAL
					Set rsTotal = objConexao.Execute("SELECT COUNT(*) AS XPEDIDOS, SUM(TOTAL_DINHEIRO) AS XTOTAL_DINHEIRO, SUM(TOTAL_CHEQUE)AS XTOTAL_CHEQUE, SUM(TOTAL_CARTAO)AS XTOTAL_CARTAO, SUM(TOTAL_RECEBIMENTO) AS XTOTAL_RECEBIMENTO, SUM(QTD_ALBUNS) AS XQTD_ALBUNS, SUM(QTD_DVDS) AS XQTD_DVDS, SUM(QTD_FOTO_TELAS_TRIOS) AS XQTD_FOTO_TELAS_TRIOS , SUM(VALOR_AJUDA_CUSTO) AS XVALOR_AJUDA_CUSTO, SUM(VALOR_TOTAL_COMISSAO) AS XVALOR_TOTAL_COMISSAO FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_VENDEDOR = "&COD_VENDEDOR&" AND TIPO_VENDA = 3 AND COD_FECHAMENTO_PERIODO ="&COD_FECHAMENTO_PERIODO)
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

					<tr bgcolor="whitesmoke">
						<td><div align="center"><strong><%=XPEDIDOS%></strong></div></td>
						<td><div align="center"></div></td>
						<td>&nbsp;</td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XTOTAL_DINHEIRO))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XTOTAL_CHEQUE))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XTOTAL_CARTAO))%></strong></div></td>
						<td><div align="center"><strong><%=XQTD_ALBUNS%></strong></div></td>
						<td><div align="center"><strong><%=XQTD_DVDS%></strong></div></td>
						<td><div align="center"><strong><%=XQTD_FOTO_TELAS_TRIOS%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XVALOR_AJUDA_CUSTO))%></strong></div></td>
						<td><div align="right"><strong><%=retiraZero2(formatNumber(XVALOR_TOTAL_COMISSAO))%></strong></div></td>
						<td>&nbsp;</td>
					</tr>		
					</table>
            	<%		
				End If
				rs.Close
				Set rs = Nothing	

				'4 - PAGAMENTO DE BOLETOS
				set rsBoleto = objConexao.execute("SELECT * FROM TB_FECHAMENTO_PAGAMENTO_BOLETOS WHERE COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND COD_VENDEDOR = "&COD_VENDEDOR)
				if not rsBoleto.eof then
					%>
					<table width="100%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
					<tr bgcolor="whitesmoke">
                	<td height="0" colspan="13"><div align="left"><strong>4 - PAGAMENTO DE BOLETOS</strong></div></td>
	    			</tr>
					<tr bgcolor="whitesmoke">
						<td width="5%" height="30"><div align="center"><strong>C�d. Pedido</strong></div></td>
						<td width="7%"><div align="center"><strong>Parcela </strong></div></td>
						<td width="5%"><div align="center"><strong>Valor</strong></div></td>
						<td width="8%"><div align="center"><strong>Data Pagamento</strong></div></td>
						<td width="8%"><div align="center"><strong>Data Baixa</strong></div></td>
						<td width="5%"><div align="center"><strong>Comiss&atilde;o</strong></div></td>
						<td width="36%"><div align="center"><strong>C�d. Formando(s) / �lbum / Institui��o de Ensino </strong></div></td>
					</tr>
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
							<td><div align="center"><%=pedidosFormandosFormatado(COD_PEDIDO, true)%></div></td>
						</tr>
						<%
					rsBoleto.MoveNext()
					Wend				
					
					'4 - PAGAMENTO DE BOLETOS TOTAL
					%>
					</table>
					<br />
					<%
				end if
				rsBoleto.Close
				set rsBoleto = Nothing

				'5 - DETALHAMENTO COORDENADOR
				'Somente para M�rio e Renato (cordenadores)
				If int(cod_vendedor) = 14 or int(cod_vendedor) = 7 or int(cod_vendedor) = 48 Then	
					Dim percentual_coodenadores
					percentual_coodenadores = 0

					'Set rs=objConexao.Execute(" SELECT * FROM TB_FECHAMENTO_COMISSIONAMENTO_COORDENADORES WHERE DATA_INICIO <= "&sqlData(data_pagamento)&" ORDER BY DATA_INICIO DESC LIMIT 1")
					'If not rs.Eof Then
					''	percentual_coodenadores = formatNumber(0&rs("PERCENTUAL"),0)
					'End If
					'rs.Close
					'Set rs=Nothing

					If int(COD_VENDEDOR) = 7 Then
						cod_grupo_vendedores = 1
					Elseif  int(COD_VENDEDOR) = 14 Then
						cod_grupo_vendedores = 2
					Elseif  int(COD_VENDEDOR) = 48 Then
						cod_grupo_vendedores = 3
					End If

					'Zerar contadores
					XTOTAL_RECEBIMENTO    = 0
					XPEDIDOS              = 0
					XCOMISSAO_COORDENADOR = 0

					Set rs = objConexao.Execute("SELECT DATA_VENDA, TOTAL_RECEBIMENTO, COD_PEDIDO, NOME_VENDEDOR FROM VW_TB_FECHAMENTO_PAGAMENTO WHERE COD_GRUPO_VENDEDORES = "&cod_grupo_vendedores&" AND COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO)
					If not rs.Eof Then
						%>
						<table width="100%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
						<tr bgcolor="whitesmoke">
							<td height="0" colspan="6"><div align="left"><strong>5 - DETALHAMENTO COORDENADOR</strong><strong></strong></div></td>
						</tr>	
						<tr bgcolor="whitesmoke">
							<td width="6%" height="30"><div align="center"><strong>C�d. Pedido </strong></div></td>
							<td width="11%"><div align="center"><strong>Data Venda </strong></div></td>
							<td width="31%"><div align="center"><strong>Vendedor</strong></div></td>
							<td width="9%"><div align="center"><strong>Valor Total</strong></div></td>
							<td width="10%"><div align="center"><strong>Comiss&atilde;o</strong></div></td>
							<td width="33%"><div align="center"><strong>C�d. Formando(s) / �lbum / Institui��o de Ensino </strong></div></td>
						</tr>
						<%
						While not rs.Eof
							COD_PEDIDO            = rs("COD_PEDIDO")
							DATA_VENDA            = formatDateTime(rs("DATA_VENDA"))
							

							'calculamos a comissao do pedido com base em todos os produtos individalmente
							COMISSAO_COORDENADOR = fechamentoRetornaValorComissaoCoordenadorPedido(rs("COD_PEDIDO"),1,true)
							TOTAL_RECEBIMENTO = cdbl(rs("TOTAL_RECEBIMENTO"))


							'TOTAL_RECEBIMENTO 	 = formatNumber(0&rs("TOTAL_RECEBIMENTO"))/1
							'COMISSAO_COORDENADOR  = (TOTAL_RECEBIMENTO/100)*percentual_coodenadores
							XTOTAL_RECEBIMENTO	 = (XTOTAL_RECEBIMENTO+TOTAL_RECEBIMENTO)/1
							XPEDIDOS              = (XPEDIDOS+1)/1
							XCOMISSAO_COORDENADOR = (XCOMISSAO_COORDENADOR+COMISSAO_COORDENADOR)/1
							

							%>
							<tr>
								<td><div align="center"><%=COD_PEDIDO%></div></td>
								<td><div align="center"><%=ArrumaData(DATA_VENDA)%></div></td>
								<td><div align="left"><%=rs("NOME_VENDEDOR")%></div></td>
								<td><div align="right"><%=retiraZero2(formatNumber(TOTAL_RECEBIMENTO,2))%></div></td>
								<td><div align="right"><%=retiraZero2(formatNumber(COMISSAO_COORDENADOR,2))%></div></td>
								<td><div align="center"><%=pedidosFormandosFormatado(COD_PEDIDO, true)%></div></td>
							</tr>
							<%
						rs.MoveNext()
						Wend

						'5 - DETALHAMENTO COORDENADOR TOTAL
						%>	
						<tr bgcolor="whitesmoke">
							<td><div align="center"><strong><%=XPEDIDOS%></strong></div></td>
							<td><div align="center"></div></td>
							<td>&mbsp;</td>
							<td><div align="right"><strong><%=retiraZero2(formatNumber(XTOTAL_RECEBIMENTO))%></strong></div></td>
							<td><div align="right"><strong><%=retiraZero2(formatNumber(XCOMISSAO_COORDENADOR))%></strong></div></td>
							<td>&mbsp;</td>
						</tr>
						</table>
						<br />
						<%
					End If
					rs.Close
					Set rs = Nothing

					dim comissao_vendedor
					'verificamos boletos de cordenadores
					set rsBoletoCod = objConexao.Execute("SELECT PEDBOL.*, VND.NOME_VENDEDOR FROM TB_FECHAMENTO_PAGAMENTO_BOLETOS PEDBOL INNER JOIN TB_VENDEDORES VND ON PEDBOL.cod_vendedor = VND.cod_vendedor WHERE  PEDBOL.cod_grupo_vendedores  = "&cod_grupo_vendedores&" AND PEDBOL.COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO)
					if not rsBoletoCod.Eof then
						%>
						<table width="100%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
						<tr bgcolor="whitesmoke">
							<td height="0" colspan="9"><div align="left"><strong>6 - DETALHAMENTO DE BOLETOS DO COORDENADOR</strong><strong></strong></div></td>
						</tr>	
						<tr bgcolor="whitesmoke">
							<td width="5%" height="30"><div align="center"><strong>C�d. Pedido</strong></div></td>
							<td width="7%"><div align="center"><strong>Parcela </strong></div></td>
							<td width="5%"><div align="center"><strong>Valor</strong></div></td>
							<td width="8%"><div align="center"><strong>Data Pagamento</strong></div></td>
							<td width="8%"><div align="center"><strong>Data Baixa</strong></div></td>
							<td width="5%"><div align="center"><strong>Comiss&atilde;o</strong></div></td>
							<td width="5%"><div align="center"><strong>Nome Vendedor</strong></div></td>
							<td width="36%"><div align="center"><strong>C�d. Formando(s) / �lbum / Institui��o de Ensino </strong></div></td>
						</tr>
						<%
						COMISSAO_COORDENADOR  = 0
						XCOMISSAO_COORDENADOR = 0
						XTOTAL_RECEBIMENTO    = 0
						XPEDIDOS = 0
						while not rsBoletoCod.eof
							COMISSAO_COORDENADOR  = cdbl(0&rsBoletoCod("Valor")) * (percentual_coodenadores/100)
							XCOMISSAO_COORDENADOR = (XCOMISSAO_COORDENADOR+COMISSAO_COORDENADOR)/1
							XTOTAL_RECEBIMENTO    = cdbl(0&rsBoletoCod("Valor")) + XTOTAL_RECEBIMENTO
							XPEDIDOS              = XPEDIDOS + 1
						   %>
							<tr bgcolor="#ffffff">
								<td width="5%"><div align="center"><%=rsBoletoCod("COD_PEDIDO")%></div></td>
								<td width="7%"><div align="center"><%=rsBoletoCod("parcela")%></div></td>
								<td width="5%"><div align="right"><%=ArrumaNumero(rsBoletoCod("valor"),"-",2)	%></div></td>
								<td width="8%"><div align="center"><%=ArrumaData(rsBoletoCod("data_pagamento"))%></div></td>
								<td width="8%"><div align="center"><%=ArrumaData(rsBoletoCod("data_baixa"))%></div></td>
								<td width="5%"><div align="right"><%=ArrumaNumero(COMISSAO_COORDENADOR,"-",2)%></div></td>
								<td width="5%"><div align="center"><%=rsBoletoCod("NOME_VENDEDOR")%></div></td>
								<td><div align="center"><%=pedidosFormandosFormatado(COD_PEDIDO, true)%></div></td>
							</tr>
						   <%
						rsBoletoCod.MoveNext
						Wend
						%>
						<tr bgcolor="whitesmoke">
							<td><div align="center"><strong><%=XPEDIDOS%></strong></div></td>
							<td><div align="center">-</div></td>
							<td><div align="right"><strong><%=retiraZero2(formatNumber(XTOTAL_RECEBIMENTO))%></strong></div></td>
							<td><div align="center">-</div></td>
							<td><div align="center">-</div></td>
							<td><div align="right"><strong><%=retiraZero2(formatNumber(XCOMISSAO_COORDENADOR))%></strong></div></td>
							<td><div align="center"><strong>-</strong></div></td>
							<td>&nbsp;</td>
						</tr>
						</table>
						<br />
						<%
					end if
					rsBoletoCod.Close
					Set rsBoletoCod = nothing
				End If
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
						<td class="textoMedio"><div align="left"><strong>PEND�NCIAS</strong></div></td>
						<td>
							<table bordercolor="#CCCCCC" width="100%" border="1" cellspacing="0" cellpadding="3">
							<tr>
								<td width="61%"><div align="center"><strong>Descri��o</strong></div></td>
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
													Response.Write("DECR�SCIMO")
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
  
				'Somente para M�rio e Renato
				If int(cod_vendedor) = 14 or int(cod_vendedor) = 7 Then			
					%>
					<tr bgcolor="whitesmoke">
						<td class="textoMedio"><div align="left"><strong>TOTAL COORDENADOR</strong></div></td>
						<td><div align="right"><%=formatCurrency(fechamentoRetornaValorRecebidoVendedorTipo(4,cod_vendedor,cod_fechamento_periodo,true))%></div></td>
					</tr>
					<%
				End If
					%>
				
				<tr bgcolor="whitesmoke">
					<td width="19%" class="textoMedio" bgcolor="whitesmoke"><div align="left"><span class="textoMaior"><strong>TOTAL A RECEBER</strong></span><strong><br /></strong></div></td>
					<td><div align="right" class="textoMaior"><strong><%=formatCurrency(fechamentoRetornaValorRecebidoVendedor(COD_VENDEDOR,cod_fechamento_periodo,true))%></strong></div></td>
				</tr>	
				</table>
		
		<p><!-- pagebreak --></p>
		<p style="page-break-after:always;"></p>
	<%
	rsMaster.MoveNext()
	Wend
Else
%>
<br />
<br />
<br />
<br />
<P align="center">N�o h� dados para o per�do selecionado</P>
<%
	End If	
	rsMaster.Close
	Set rsMaster = Nothing
	FecharConexao()
%>
</body>
</html>



