<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/classe_fechamento.asp" -->
<%
Server.ScriptTimeout = 2240
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<%
Dim VENDEDOR_RECEBE_EM, COD_FECHAMENTO_PERIODO
VENDEDOR_RECEBE_EM = Request.QueryString("VENDEDOR_RECEBE_EM")
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta HTTP-EQUIV = "Refresh" CONTENT = "0; URL = adm_rotina_fechamento_extrato_pagamento_part2_tipo2.asp?VENDEDOR_RECEBE_EM=<%=VENDEDOR_RECEBE_EM%>">
<link type="text/css" rel="stylesheet" href="css/base.css" media="all"></link>
<title>Gerador de Pagamento</title>
<style type="text/css">
<!--
body {
	margin-left: 10px;
	margin-top: 10px;
	margin-right: 10px;
	margin-bottom: 10px;
	background-color: #FFFFFF;
}
.style1 {
	color: #0000CC;
	font-style: italic;
}
-->
</style>
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
</head>
<body>
<%
Dim QTD_PEDIDOS_TOTAL, QTD_PEDIDOS_PROCESSADOS, QTD_PEDIDOS_NAO_PROCESSADOS, PERCENTUAL_FINALIZADOS_BARRA
QTD_PEDIDOS_TOTAL = 0
QTD_PEDIDOS_PROCESSADOS = 0
QTD_PEDIDOS_NAO_PROCESSADOS = 0
PERCENTUAL_FINALIZADOS_BARRA = 0

AbrirConexao()

'////////////////////////////
'Busca o campo COD_FECHAMENTO_PERIODO de acordo com a data do pagamento
Set rs=objConexao.Execute("SELECT COD_FECHAMENTO_PERIODO FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS where data_pagamento = "&sqlData(VENDEDOR_RECEBE_EM))
	
	If not rs.Eof Then
		COD_FECHAMENTO_PERIODO = rs(0)
	End If
	
rs.Close
Set rs=Nothing


sql = "SELECT p.COD_PEDIDO, p.COD_VENDEDOR, p.DATA_VENDA, p.VALOR_AJUDA_CUSTO, p.FORMA_PAGAMENTO, p.TIPO_VENDA, p.COD_GRUPO_VENDEDORES, p.PAGAMENTO_EM_DINHEIRO, p.pre_venda, p.tipo_venda AS tipo_venda_pedido, p.forma_pagamento AS forma_pgto_pedido, " &_
      "COALESCE(b.XTOTAL_BOLETO,0) AS TOTAL_BOLETO, " &_
      "COALESCE(c.XTOTAL_CHEQUE,0) AS TOTAL_CHEQUE, " &_
      "COALESCE(t.XTOTAL_CARTAO,0) AS TOTAL_CARTAO, " &_
      "COALESCE(f.QTD_FORMANDOS,0) AS QTD_ALBUNS, " &_
      "COALESCE(d.QTD_DVD,0) + COALESCE(db.QTD_DVDS_BRINDES,0) AS QTD_DVDS, " &_
      "COALESCE(ft.XTOTAL_FTO_TELAS_AGREGADOS,0) + COALESCE(fv.QTD_TOTAL,0) AS QTD_FOTO_TELAS_TRIOS, " &_
      "COALESCE((SELECT SUM(VALOR_UN * QUANTIDADE) FROM TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO = p.COD_PEDIDO AND COD_PRODUTO = 13),0) AS VALOR_FRETE " &_
      "FROM TB_PEDIDOS p " &_
      "LEFT JOIN (SELECT COD_PEDIDO, SUM(XTOTAL_BOLETO) AS XTOTAL_BOLETO FROM VW_PEDIDOS_BOLETO_TOTAL GROUP BY COD_PEDIDO) b ON b.COD_PEDIDO = p.COD_PEDIDO " &_
      "LEFT JOIN (SELECT COD_PEDIDO, SUM(XTOTAL_CHEQUE) AS XTOTAL_CHEQUE FROM VW_PEDIDOS_CHEQUE_TOTAL GROUP BY COD_PEDIDO) c ON c.COD_PEDIDO = p.COD_PEDIDO " &_
      "LEFT JOIN (SELECT COD_PEDIDO, SUM(XTOTAL_CARTAO) AS XTOTAL_CARTAO FROM VW_PEDIDOS_CARTAO_TOTAL GROUP BY COD_PEDIDO) t ON t.COD_PEDIDO = p.COD_PEDIDO " &_
      "LEFT JOIN (SELECT COD_PEDIDO, SUM(QTD_FORMANDOS) AS QTD_FORMANDOS FROM VW_QTD_FORMANDOS_POR_PEDIDO GROUP BY COD_PEDIDO) f ON f.COD_PEDIDO = p.COD_PEDIDO " &_
      "LEFT JOIN (SELECT COD_PEDIDO, SUM(QTD_DVD) AS QTD_DVD FROM VW_PEDIDOS_PRODUTOS_TOTAIS_DVD_VENDA_NORMAL GROUP BY COD_PEDIDO) d ON d.COD_PEDIDO = p.COD_PEDIDO " &_
      "LEFT JOIN (SELECT COD_PEDIDO, SUM(QTD_DVDS_BRINDES) AS QTD_DVDS_BRINDES FROM VW_PEDIDOS_PRODUTOS_TOTAIS_DVD_BRINDES GROUP BY COD_PEDIDO) db ON db.COD_PEDIDO = p.COD_PEDIDO " &_
      "LEFT JOIN (SELECT COD_PEDIDO, SUM(XTOTAL_FTO_TELAS_AGREGADOS) AS XTOTAL_FTO_TELAS_AGREGADOS FROM VW_TOTAL_PARCIAL_FOTO_TELAS_TRIOS_PRODUTOS_AGREGADOS GROUP BY COD_PEDIDO) ft ON ft.COD_PEDIDO = p.COD_PEDIDO " &_
      "LEFT JOIN (SELECT COD_PEDIDO, SUM(QTD_TOTAL) AS QTD_TOTAL FROM VW_TOTAL_PARCIAL_FOTO_TELAS_TRIOS_PRODUTOS_VEND_A_PARTE GROUP BY COD_PEDIDO) fv ON fv.COD_PEDIDO = p.COD_PEDIDO " &_
      "WHERE p.FINALIZA_CORRECAO = 1 AND p.PAGAMENTO_PROCESSADO = 2 AND p.VENDEDOR_RECEBE_EM = "&sqlData(VENDEDOR_RECEBE_EM)&" ORDER BY p.COD_VENDEDOR LIMIT 1"

'Response.Write(SQL)
Set rs=objConexao.Execute(SQL)
				
Dim COD_VENDEDOR, DATA_VENDA,VALOR_AJUDA_CUSTO,FORMA_PAGAMENTO, TIPO_VENDA, COD_GRUPO_VENDEDORES, VALOR_TOTAL_COMISSAO, VALOR_TOTAL_COMISSAO_BOLETOS
				
If not rs.Eof Then				
%>
<h1 align="center" class="style1">AGUARDE PROCESSANDO PAGAMENTO!</h1>
<p align="center">
<%
Set rs2=objConexao.Execute("SELECT COUNT(PAGAMENTO_PROCESSADO) AS QTD FROM TB_PEDIDOS WHERE VENDEDOR_RECEBE_EM = "&sqlData(VENDEDOR_RECEBE_EM))
	If not rs2.Eof Then
		QTD_PEDIDOS_TOTAL = formatNumber(0&rs2(0),0)
	End If
rs2.Close
Set rs2=Nothing

Set rs2=objConexao.Execute("SELECT COUNT(PAGAMENTO_PROCESSADO) AS QTD FROM TB_PEDIDOS WHERE PAGAMENTO_PROCESSADO = 1 and  VENDEDOR_RECEBE_EM = "&sqlData(VENDEDOR_RECEBE_EM))
	If not rs2.Eof Then
		QTD_PEDIDOS_PROCESSADOS = formatNumber(0&rs2(0),0)
	End If
rs2.Close
Set rs2=Nothing

QTD_PEDIDOS_NAO_PROCESSADOS = (QTD_PEDIDOS_TOTAL-QTD_PEDIDOS_PROCESSADOS)/1

If QTD_PEDIDOS_PROCESSADOS > 1 Then
	PERCENTUAL_FINALIZADOS_BARRA = (100*QTD_PEDIDOS_PROCESSADOS)/QTD_PEDIDOS_TOTAL
Else
	PERCENTUAL_FINALIZADOS_BARRA = 0
End If

%>
<strong><%=QTD_PEDIDOS_PROCESSADOS%>/<%=QTD_PEDIDOS_TOTAL%></strong>
<table width="300" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FF0000">
  <tr>
    <td>
    <div style="background-color:#000099; border:solid 1px #0099FF; width:<%=formatNumber(PERCENTUAL_FINALIZADOS_BARRA,0)%>%">
  &nbsp;
    </div></td>
  </tr>
</table>


</p>
<table width="100%"  border="1" align="center"
					cellpadding="5" 
					style="border-collapse:collapse">
				<tr bgcolor="whitesmoke">	
                <td width="6%" height="30"><div align="center"><strong>Cód. Vendedor</strong></div></td>
			    <td width="5%"><div align="center"><strong>Cód. Pedido </strong></div></td>
			    <td width="7%"><div align="center"><strong>Data Venda </strong></div></td>
			    <td width="5%"><div align="center"><strong>Dinheiro</strong></div></td>
			    <td width="6%"><div align="center"><strong>Cheque</strong></div></td>
			    <td width="4%"><div align="center"><strong>Cartão</strong></div></td>
			    <td width="10%"><div align="center"><strong>TOTAL RECEBIDO</strong></div></td>
			    <td width="9%"><div align="center"><strong>Álbuns</strong></div></td>
			    <td width="7%"><div align="center"><strong>DVDs</strong></div></td>
                <td width="8%"><div align="center"><strong>Foto-Telas e Trios </strong></div></td>
                <td width="9%"><div align="center"><strong>Ajuda de Custo </strong></div></td>
                <td width="8%"><div align="center"><strong>Forma Pagamento</strong></div></td>
                <td width="8%"><div align="center"><strong>Tipo Venda</strong></div></td>
                <td width="4%"><div align="center"><strong>Cod Grupo Especial</strong></div></td>
				<td width="4%"><div align="center"><strong>Valor Total de Comiss&atilde;o</strong></div></td>
				</tr>
				<%				
				While not rs.Eof
								
					COD_PEDIDO			= rs("COD_PEDIDO")
					COD_VENDEDOR		= rs("COD_VENDEDOR")
					DATA_VENDA			= formatDateTime(rs("DATA_VENDA"))
					FORMA_PAGAMENTO		= rs("FORMA_PAGAMENTO")
					VALOR_AJUDA_CUSTO	= rs("VALOR_AJUDA_CUSTO") 
					TIPO_VENDA			= rs("TIPO_VENDA")
					COD_GRUPO_VENDEDORES= rs("COD_GRUPO_VENDEDORES")
					TOTAL_DINHEIRO		= formatNumber(0&rs("PAGAMENTO_EM_DINHEIRO"))/1
					TOTAL_CHEQUE		= formatNumber(0&rs("TOTAL_CHEQUE"))/1
					TOTAL_CARTAO		= formatNumber(0&rs("TOTAL_CARTAO"))/1
					TOTAL_BOLETO		= formatNumber(0&rs("TOTAL_BOLETO"))/1
					TOTAL_RECEBIMENTO	= formatNumber(TOTAL_CHEQUE+TOTAL_CARTAO+TOTAL_BOLETO+TOTAL_DINHEIRO)/1
					QTD_ALBUNS			= formatNumber(0&rs("QTD_ALBUNS"),0)
					QTD_DVDS			= formatNumber(0&rs("QTD_DVDS"),0)
					QTD_FOTO_TELAS_TRIOS= formatNumber(0&rs("QTD_FOTO_TELAS_TRIOS"),0)
					
					Dim pre_venda_pedido, tipo_venda_pedido, forma_pgto_pedido, valor_frete, base_comissao, perc_comissao
					pre_venda_pedido = Trim(rs("pre_venda"))
					tipo_venda_pedido = rs("tipo_venda_pedido")
					forma_pgto_pedido = rs("forma_pgto_pedido")
					valor_frete = CDbl(0 & rs("VALOR_FRETE"))
					base_comissao = CDbl(0 & TOTAL_RECEBIMENTO) - valor_frete
					
					Select Case True
					    Case pre_venda_pedido = "S" And forma_pgto_pedido = 1
					        perc_comissao = 0.12
					    Case pre_venda_pedido = "S" And forma_pgto_pedido = 2
					        perc_comissao = 0.10
					    Case pre_venda_pedido = "N" And tipo_venda_pedido = 1 And forma_pgto_pedido = 1
					        perc_comissao = 0.15
					    Case pre_venda_pedido = "N" And tipo_venda_pedido = 1 And forma_pgto_pedido = 2
					        perc_comissao = 0.13
					    Case pre_venda_pedido = "N" And tipo_venda_pedido = 2 And forma_pgto_pedido = 1
					        perc_comissao = 0.25
					    Case pre_venda_pedido = "N" And tipo_venda_pedido = 2 And forma_pgto_pedido = 2
					        perc_comissao = 0.23
					    Case Else
					        perc_comissao = 0
					End Select
					
					VALOR_TOTAL_COMISSAO = Round(base_comissao * perc_comissao, 2)
			    <td><div align="right"><%=TOTAL_BOLETO%></div></td>		    
			    <td><div align="right"><%=TOTAL_RECEBIMENTO%></div></td>
			    <td><div align="center"><%=QTD_ALBUNS%></div></td>
			    <td><div align="center"><%=QTD_DVDS%></div></td>
			    <td><div align="center"><%=QTD_FOTO_TELAS_TRIOS%></div></td>
			    <td><div align="right"><%=VALOR_AJUDA_CUSTO%></div></td>
			    <td><div align="center"><%=FORMA_PAGAMENTO%></div></td>
					<td><div align="center"><%=TIPO_VENDA%></div></td>
					<td><div align="center"><%=COD_GRUPO_VENDEDORES%>
                    </div></td>
					<td><div align="center">
				      <%=VALOR_TOTAL_COMISSAO%></div></td>
					</tr>

                <%
				sql = "INSERT INTO TB_FECHAMENTO_PAGAMENTO (COD_PEDIDO, VENDEDOR_RECEBE_EM, COD_VENDEDOR, DATA_VENDA, FORMA_PAGAMENTO, TOTAL_DINHEIRO, TOTAL_CHEQUE, TOTAL_CARTAO, TOTAL_BOLETO, TOTAL_RECEBIMENTO, QTD_ALBUNS, QTD_DVDS, QTD_FOTO_TELAS_TRIOS, VALOR_AJUDA_CUSTO, TIPO_VENDA, COD_GRUPO_VENDEDORES, VALOR_TOTAL_COMISSAO,COD_FECHAMENTO_PERIODO) VALUES ("&_
				COD_PEDIDO&","&_
				sqlData(VENDEDOR_RECEBE_EM)&","&_
				COD_VENDEDOR&","&_
				sqlData(DATA_VENDA)&","&_
				FORMA_PAGAMENTO&","&_
				"'"&sqlNumero(TOTAL_DINHEIRO)&"',"&_
				"'"&sqlNumero(TOTAL_CHEQUE)&"',"&_
				"'"&sqlNumero(TOTAL_CARTAO)&"',"&_
				"'"&sqlNumero(TOTAL_BOLETO)&"',"&_				
				"'"&sqlNumero(TOTAL_RECEBIMENTO)&"',"&_
				"'"&sqlNumero(QTD_ALBUNS)&"',"&_
				"'"&sqlNumero(QTD_DVDS)&"',"&_
				"'"&sqlNumero(QTD_FOTO_TELAS_TRIOS)&"',"&_
				"'"&sqlNumero(VALOR_AJUDA_CUSTO)&"',"&_
				TIPO_VENDA&","&_
				COD_GRUPO_VENDEDORES&","&_
				"'"&sqlNumero(VALOR_TOTAL_COMISSAO)&"',"&COD_FECHAMENTO_PERIODO&")"
				
				'rESPONSE.Write(SQL)
				
				objConexao.Execute(sql)
				
				objConexao.Execute("UPDATE TB_PEDIDOS SET PAGAMENTO_PROCESSADO= '1' WHERE COD_PEDIDO = "&COD_PEDIDO&" LIMIT 1")
				
				'Response.Write(sql&"<BR>")
					
								
				rs.MoveNext()
				Wend
				%>
			</table>
<br />
<br />

<%Else
	
	

	'PROCESSAMOS BOLETOS BAIXADOS COM PAGAMENTO PARA O FECHAMENTO
	sql_insert_boletos = "INSERT INTO TB_FECHAMENTO_PAGAMENTO_BOLETOS (cod_pedido, cod_boleto, cod_vendedor, valor, valor_pago, valor_comissao_vendedor, data_vencimento, data_pagamento, data_baixa, cod_fechamento_periodo, data_processamento, parcela, cod_grupo_vendedores, vendedor_recebe_em) SELECT pedbol.cod_pedido, pedbol.cod_boleto, pedbol.cod_vendedor, pedbol.valor, pedbol.valor_pago, pedbol.valor_comissao_vendedor, pedbol.data_vencimento, pedbol.data_pagamento, pedbol.data_baixa, "&sqlNumero(cod_fechamento_periodo)&",  now(), pedbol.parcela, vnd.cod_grupo_vendedores, pedbol.vendedor_recebe_em  FROM TB_PEDIDOS_BOLETOS pedbol inner join TB_VENDEDORES vnd on pedbol.COD_VENDEDOR = vnd.COD_VENDEDOR WHERE vendedor_recebe_em = "&sqlData(vendedor_recebe_em)&" and cod_fechamento_periodo is null and data_baixa is not null and pagamento_processado = 0"


	'FLAGAMOS OS BOLETOS SEELCIONADOS
	sql_update_boletos = "UPDATE TB_PEDIDOS_BOLETOS SET cod_fechamento_periodo = "&sqlNumero(cod_fechamento_periodo)&" , pagamento_processado = 1 WHERE vendedor_recebe_em = "&sqlData(vendedor_recebe_em)&" and cod_fechamento_periodo is null and data_baixa is not null and pagamento_processado = 0"

	'executamos
	objConexao.Execute(sql_insert_boletos)
	objConexao.Execute(sql_update_boletos)


	Set rs2=objConexao.Execute("SELECT * FROM TB_FECHAMENTO_PAGAMENTO_CABECALHO where VENDEDOR_RECEBE_EM = "&sqlData(VENDEDOR_RECEBE_EM))
			
			If rs2.Eof Then
				objConexao.Execute("INSERT INTO TB_FECHAMENTO_PAGAMENTO_CABECALHO (VENDEDOR_RECEBE_EM,PAGAMENTO_GERADO_POR) VALUES ("&sqlData(VENDEDOR_RECEBE_EM)&", '"&Session("cod_usuario")&"')")

				objConexao.Execute("UPDATE TB_FECHAMENTO_PAGAMENTO_PERIODOS SET data_processado_pagamento = NOW() WHERE data_pagamento = "&sqlData(VENDEDOR_RECEBE_EM))
			End If
	rs2.Close	
	Set rs2=Nothing
	

''	Response.Write sql_insert_boletos
''	Response.Write sql_update_boletos
'Response.end
	Response.redirect("cap_fechamentos.asp?codmsg=1&vendedor_recebe_em="&vendedor_recebe_em)
	
End If

rs.Close
Set rs=Nothing

FecharConexao()
%>
</body>
</html>



