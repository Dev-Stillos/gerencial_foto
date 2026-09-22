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

'Set rsMaster = objConexao.Execute("SELECT * FROM (SELECT "&_
'	"TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR ,"&_
'	"TB_VENDEDORES.NOME_VENDEDOR "&_
'	"FROM TB_VENDEDORES INNER JOIN TB_FECHAMENTO_PAGAMENTO ON TB_VENDEDORES.COD_VENDEDOR = TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR "&_
'	"WHERE (TB_FECHAMENTO_PAGAMENTO.COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR <> 27) OR TB_VENDEDORES.COD_VENDEDOR IN (SELECT cod_vendedor FROM TB_FECHAMENTO_PENDENCIAS_PAGAMENTO WHERE cod_fechamento_periodo = '"&COD_FECHAMENTO_PERIODO&"') or TB_VENDEDORES.COD_VENDEDOR IN(7,14,48) "&_
'	"GROUP BY TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR ORDER BY TB_VENDEDORES.NOME_VENDEDOR UNION ALL "&_
'	"SELECT TB_FECHAMENTO_PAGAMENTO_BOLETOS.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR FROM TB_VENDEDORES INNER JOIN TB_FECHAMENTO_PAGAMENTO_BOLETOS ON TB_VENDEDORES.COD_VENDEDOR = TB_FECHAMENTO_PAGAMENTO_BOLETOS.COD_VENDEDOR WHERE (TB_FECHAMENTO_PAGAMENTO_BOLETOS.COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&") GROUP BY TB_FECHAMENTO_PAGAMENTO_BOLETOS.COD_VENDEDOR ORDER BY NOME_VENDEDOR) AS teste GROUP BY teste.COD_VENDEDOR")

Set rsMaster = objConexao.Execute("SELECT * FROM ( SELECT  TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR , TB_VENDEDORES.NOME_VENDEDOR FROM TB_VENDEDORES INNER JOIN TB_FECHAMENTO_PAGAMENTO ON TB_VENDEDORES.COD_VENDEDOR = TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR WHERE (TB_FECHAMENTO_PAGAMENTO.COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR <> 27) OR TB_VENDEDORES.COD_VENDEDOR  IN (SELECT cod_vendedor FROM TB_FECHAMENTO_PENDENCIAS_PAGAMENTO WHERE cod_fechamento_periodo = '"&COD_FECHAMENTO_PERIODO&"') or TB_VENDEDORES.COD_VENDEDOR IN(7,14,48) UNION ALL SELECT TB_FECHAMENTO_PAGAMENTO_BOLETOS.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR FROM TB_VENDEDORES INNER JOIN TB_FECHAMENTO_PAGAMENTO_BOLETOS ON TB_VENDEDORES.COD_VENDEDOR = TB_FECHAMENTO_PAGAMENTO_BOLETOS.COD_VENDEDOR WHERE (TB_FECHAMENTO_PAGAMENTO_BOLETOS.COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&")) AS teste GROUP BY teste.COD_VENDEDOR ORDER BY teste.NOME_VENDEDOR")






If not rsMaster.Eof Then
	While not rsMaster.Eof
		COD_VENDEDOR 	= rsMaster("COD_VENDEDOR")
		NOME_VENDEDOR 	= rsMaster("NOME_VENDEDOR")
		%>

				<table style="margin-top:10px" border="0" cellpadding="5" cellspacing="0" align="center" width="100%">
				<tr>
					<td><div align="center"><span class="textoMaior"><strong>Data do Pagamento <%=ArrumaData(data_pagamento)%> - Per&iacute;odo de Vendas de <%=data_inicio%> &agrave; <%=data_termino%> - Impresso em <%=ArrumaData(Now())%></strong></span></div></td>
				</tr>
				</table>
				<table style="margin-top:10px" cellpadding="5" cellspacing="0" align="center" width="100%">
				<tr>
					<td><div align="left"><span class="textoMedio"><strong><%=Ucase(NOME_VENDEDOR)%></strong></span></div></td>
				</tr>
				</table>
				<%
				'1 - VENDAS NORMAIS
				'Set rs = objConexao.Execute("SELECT * FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND COD_VENDEDOR = "&COD_VENDEDOR&" AND TIPO_VENDA = 1",TOTAL_PEDIDOS_TIPO_VENDA_1)
				Set rs = objConexao.Execute("SELECT TB_FECHAMENTO_PAGAMENTO.*, (select sum(valor_un * quantidade) from TB_PEDIDOS_PRODUTOS b where b.cod_pedido=TB_FECHAMENTO_PAGAMENTO.cod_pedido and cod_produto=13) as frete FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND COD_VENDEDOR = "&COD_VENDEDOR&" AND TIPO_VENDA = 1",TOTAL_PEDIDOS_TIPO_VENDA_1)
				frt = 0
				If not rs.Eof Then
					%>
					<table width="100%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
					<tr bgcolor="whitesmoke">
						<td height="0" colspan="13"><div align="left"><strong>1 - VENDAS NORMAIS</strong></div></td>
					</tr>	
					<tr bgcolor="whitesmoke">
						<td width="5%" height="30"><div align="center"><strong>C&oacute;d. Pedido </strong></div></td>
						<td width="7%"><div align="center"><strong>Data Venda </strong></div></td>
						<td width="5%"><div align="center"><strong>Tipo Venda </strong></div></td>
						<td width="5%"><div align="center"><strong>Forma Pag.</strong></div></td>
						<td width="8%"><div align="center"><strong>Dinheiro</strong></div></td>
						<td width="8%"><div align="center"><strong>Cheque</strong></div></td>
						<td width="8%"><div align="center"><strong>Cart&atilde;o</strong></div></td>
						<td width="8%"><div align="center"><strong>Boleto</strong></div></td>
						<td width="5%"><div align="center"><strong>&aacute;buns</strong></div></td>
						<td width="5%"><div align="center"><strong>DVDs</strong></div></td>
						<td width="6%"><div align="center"><strong>Foto-Telas e Trios </strong></div></td>
						<td width="3%"><div align="center"><strong>Ajuda de Custo </strong></div></td>
						<td width="3%"><div align="center"><strong>Frete </strong></div></td>
						<td width="4%"><div align="center"><strong>Comiss&atilde;o</strong></div></td>
						<td width="31%"><div align="center"><strong>C&oacute;d. Formando(s) / &aacute;lbum / Institui&ccedil;&atilde;o de Ensino </strong></div></td>
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
						
						XVALOR_TOTAL_COMISSAO = 0 : If Not IsNull(rsTotal("XVALOR_TOTAL_COMISSAO")) And IsNumeric(rsTotal("XVALOR_TOTAL_COMISSAO")) Then XVALOR_TOTAL_COMISSAO = CDbl(rsTotal("XVALOR_TOTAL_COMISSAO"))
					End If
					rsTotal.Close
					Set rsTotal = Nothing
					%>
					<tr bgcolor="whitesmoke">
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
					</table>
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
   	  			<table width="100%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
					<tr bgcolor="whitesmoke">
						<td height="0" colspan="13"><div align="left"><strong>2 - VENDAS REPASSE</strong></div></td>
					</tr>
					<tr bgcolor="whitesmoke">
						<td width="5%" height="30"><div align="center"><strong>C&oacute;d. Pedido </strong></div></td>
						<td width="7%"><div align="center"><strong>Data Venda </strong></div></td>
						<td width="5%"><div align="center"><strong>Tipo Venda </strong></div></td>
						<td width="5%"><div align="center"><strong>Forma Pag.</strong></div></td>
						<td width="8%"><div align="center"><strong>Dinheiro</strong></div></td>
						<td width="8%"><div align="center"><strong>Cheque</strong></div></td>
						<td width="8%"><div align="center"><strong>Cart&atilde;o</strong></div></td>
						<td width="5%"><div align="center"><strong>&aacute;lbuns</strong></div></td>
						<td width="5%"><div align="center"><strong>DVDs</strong></div></td>
						<td width="6%"><div align="center"><strong>Foto-Telas e Trios </strong></div></td>
						<td width="3%"><div align="center"><strong>Ajuda de Custo </strong></div></td>
						<td width="3%"><div align="center"><strong>Frete </strong></div></td>
						<td width="4%"><div align="center"><strong>Comiss&atilde;o</strong></div></td>
						<td width="31%"><div align="center"><strong>C&oacute;d. Formando(s) / &aacute;lbum / Institui&ccedil;&atilde;o de Ensino </strong></div></td>
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
					<tr bgcolor="whitesmoke">
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
					</table>

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
					<table width="100%" border="1" align="center"	cellpadding="5" style="border-collapse:collapse">
					<tr bgcolor="whitesmoke">
						<td height="0" colspan="13"><div align="left"><strong>3 - VENDAS REPASSE</strong> <strong>GRUPO ESPECIAL</strong></div></td>
					</tr>
					<tr bgcolor="whitesmoke">
						<td width="5%" height="30"><div align="center"><strong>C&oacute;d. Pedido </strong></div></td>
						<td width="7%"><div align="center"><strong>Data Venda </strong></div></td>
						<td width="5%"><div align="center"><strong>Tipo Venda </strong></div></td>
						<td width="5%"><div align="center"><strong>Forma Pag.</strong></div></td>
						<td width="8%"><div align="center"><strong>Dinheiro</strong></div></td>
						<td width="8%"><div align="center"><strong>Cheque</strong></div></td>
						<td width="8%"><div align="center"><strong>Cart&atilde;o</strong></div></td>
						<td width="5%"><div align="center"><strong>&aacute;lbuns</strong></div></td>
						<td width="5%"><div align="center"><strong>DVDs</strong></div></td>
						<td width="6%"><div align="center"><strong>Foto-Telas e Trios </strong></div></td>
						<td width="3%"><div align="center"><strong>Ajuda de Custo </strong></div></td>
						<td width="3%"><div align="center"><strong>Frete </strong></div></td>
						<td width="4%"><div align="center"><strong>Comiss&atilde;o</strong></div></td>
						<td width="31%"><div align="center"><strong>C&oacute;d. Formando(s) / &aacute;lbum / Institui&ccedil;&atilde;o de Ensino </strong></div></td>
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

					<tr bgcolor="whitesmoke">
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
							<td><div align="center"><%=pedidosFormandosFormatado(rsBoleto("COD_PEDIDO"), true)%></div></td>
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
				
				<tr bgcolor="whitesmoke">
					<td width="19%" class="textoMedio" bgcolor="whitesmoke"><div align="left"><span class="textoMaior"><strong>TOTAL A RECEBER</strong></span><strong><br /></strong></div></td>
<td><div align="right" class="textoMaior"><strong><% Set rsTotalReceber = objConexao.Execute("SELECT IFNULL(SUM(fp.VALOR_TOTAL_COMISSAO),0) AS TOTAL_A_RECEBER FROM TB_FECHAMENTO_PAGAMENTO fp LEFT JOIN TB_PEDIDOS p ON p.COD_PEDIDO = fp.COD_PEDIDO WHERE fp.COD_VENDEDOR=" & COD_VENDEDOR & " AND fp.COD_FECHAMENTO_PERIODO=" & COD_FECHAMENTO_PERIODO & " AND NOT (p.PRE_VENDA='S' AND fp.TOTAL_BOLETO > 0)") : If Not rsTotalReceber.Eof Then Response.Write(formatCurrency(rsTotalReceber("TOTAL_A_RECEBER"))) Else Response.Write(formatCurrency(0)) : End If : rsTotalReceber.Close : Set rsTotalReceber = Nothing %></strong></div></td>
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



