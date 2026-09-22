<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
 
Response.Buffer = False
Dim cod_fechamento_periodo, vendedor_recebe_em, autorizacao_diretoria_para_finalizar
cod_fechamento_periodo = Request.QueryString("cod_fechamento_periodo")
	
	AbrirConexao()

	Set rs2=objConexao.Execute("SELECT data_pagamento, autorizacao_diretoria_para_finalizar FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE cod_fechamento_periodo = "&cod_fechamento_periodo)
								
		If not rs2.Eof Then
			vendedor_recebe_em = formatDateTime(rs2(0))
			autorizacao_diretoria_para_finalizar = Cstr(""&rs2("autorizacao_diretoria_para_finalizar"))
		End If
								
	rs2.Close
	Set rs2=Nothing
	FecharConexao()	
	
Dim VERIFICA_ERRO
VERIFICA_ERRO = 2
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="all"></link>
<title>Gerar Fechamento</title>
<style type="text/css">

.style1 {color: #ff007f}
.style2 {color: #ff007f
	font-size: 16px;
	font-weight: bold;
}
body {
	background-color: white;
	margin-left: 10px;
	margin-top: 10px;
	margin-right: 10px;
	margin-bottom: 10px;
}

</style>
</head>
<body>
<h2>
  
<%
AbrirConexao()
%>
</h2>
<h2 class="style1" style="color: #ff007f !important;">1) Verifica se h&aacute; pedidos n&atilde;o  processados para recebimento anterior a este pagamento</h2>

<%
sql = "SELECT COD_PEDIDO, DATA_VENDA, DATA_CADASTRO, VENDEDOR_RECEBE_EM FROM TB_PEDIDOS WHERE VENDEDOR_RECEBE_EM < "&sqlData(VENDEDOR_RECEBE_EM)&" AND FINALIZA_CORRECAO = 1 AND PAGAMENTO_PROCESSADO = 2"
'Response.Write(sql)
Set rs=objConexao.Execute(sql)

If not rs.Eof Then

VERIFICA_ERRO = 1
%>
<table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
        <td><div align="center"><strong>C&oacute;d. Pedido</strong></div></td>
        <td><div align="center"><strong>Data Venda</strong></div></td>
        <td><div align="center"><strong>Data Cadastro</strong></div></td>
        <td><div align="center"><strong>Vendedor Recebe em:</strong></div></td>
  </tr>
    <%
    While not rs.Eof
    %> 
      <tr>
        <td><div align="center"><a href="lca_busca_detalhamento_pedido.asp?cod_pedido=<%=rs("COD_PEDIDO")%>" target="_blank"><%=rs("COD_PEDIDO")%></a></div></td>
        <td><div align="center"><%=rs("DATA_VENDA")%></div></td>
        <td><div align="center"><%=rs("DATA_CADASTRO")%></div></td>
        <td><div align="center"><%=rs("VENDEDOR_RECEBE_EM")%></div></td>
      </tr>
    <%
    rs.MoveNext()
    Wend
    %>
</table>
<br />
<div style="border-bottom:solid 1px #CCCCCC">
<table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>H&aacute; pedidos n&atilde;o processados programados para um recebimento anterior a este pagamento. Corrija antes de continuar</strong></td>
    <td width="8%"><img src="images/icons/icon_red.jpg" width="36" height="39" /></td>
  </tr>
</table>
</div>
<%
Else
%>
<div style="border-bottom:solid 1px #CCCCCC">
<table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>Todos os pedidos anteriores a este fechamento j&aacute; foram processados.</strong></td>
    <td width="8%"><img src="images/icons/icon_green.jpg" width="36" height="39" /></td>
  </tr>
</table>
</div>
<%
End If

rs.Close
Set rs=Nothing
%>
<br />
<br />

<h2 class="style1" style="color: #ff007f !important;">2) Verifica se h&aacute; pedidos que n&atilde;o h&aacute; TIPO VENDA, FORMA_PAGAMENTO, COD_GRUPO_VENDEDOR, COD_FECHAMENTO_PERIODO para este pagamento</h2>

<%
sql = "SELECT COD_PEDIDO, DATA_VENDA, DATA_CADASTRO, VENDEDOR_RECEBE_EM FROM TB_PEDIDOS WHERE VENDEDOR_RECEBE_EM = "&sqlData(VENDEDOR_RECEBE_EM)&" AND FINALIZA_CORRECAO = 1 AND (TIPO_VENDA = 0 or cod_fechamento_periodo = 0)"
'Response.Write(sql)
Set rs=objConexao.Execute(sql)

If not rs.Eof Then

VERIFICA_ERRO = 1
%>
<table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
        <td><div align="center"><strong>C&oacute;d. Pedido</strong></div></td>
        <td><div align="center"><strong>Data Venda</strong></div></td>
        <td><div align="center"><strong>Data Cadastro</strong></div></td>
        <td><div align="center"><strong>Vendedor Recebe em:</strong></div></td>
  </tr>
    <%
    While not rs.Eof
    %> 
      <tr>
        <td><div align="center"><%=rs("COD_PEDIDO")%></div></td>
        <td><div align="center"><%=rs("DATA_VENDA")%></div></td>
        <td><div align="center"><%=rs("DATA_CADASTRO")%></div></td>
        <td><div align="center"><%=rs("VENDEDOR_RECEBE_EM")%></div></td>
      </tr>
    <%
    rs.MoveNext()
    Wend
    %>
</table>
    <br />
    <div style="border-bottom:solid 1px #CCCCCC">
    <table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>Informe o Administrador do Sistema a Falha</strong></td>
    <td width="8%"><img src="images/icons/icon_red.jpg" width="36" height="39" /></td>
  </tr>
</table>
</div>
<%
Else
%>
<div style="border-bottom:solid 1px #CCCCCC">
<table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>N&atilde;o h&aacute; pedidos com estes campos em branco</strong></td>
    <td width="8%"><img src="images/icons/icon_green.jpg" width="36" height="39" /></td>
  </tr>
</table>
</div>
<%
End If

rs.Close
Set rs=Nothing
%>
<br />
<br />
<h2 class="style1" style="color: #ff007f !important;">3) Verifica se h&aacute; pedidos j&aacute; processados para pagamentos futuros</h2>
<%
sql = "SELECT COD_PEDIDO, DATA_VENDA, DATA_CADASTRO, VENDEDOR_RECEBE_EM FROM TB_PEDIDOS WHERE VENDEDOR_RECEBE_EM >= "&sqlData(VENDEDOR_RECEBE_EM)&" AND PAGAMENTO_PROCESSADO = 1"
'Response.Write(sql)
Set rs=objConexao.Execute(sql)

If not rs.Eof Then

VERIFICA_ERRO = 1
%>
<table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
        <td width="15%"><div align="center"><strong>C&oacute;d. Pedido</strong></div></td>
        <td width="18%"><div align="center"><strong>Data Venda</strong></div></td>
        <td width="21%"><div align="center"><strong>Data Cadastro</strong></div></td>
        <td width="21%"><div align="center"><strong>Data Processamento</strong></div></td>
        <td width="25%"><div align="center"><strong>Vendedor Recebe em:</strong></div></td>
      </tr>
    <%
    While not rs.Eof
    %> 
      <tr>
        <td><div align="center"><%=rs("COD_PEDIDO")%></div></td>
        <td><div align="center"><%=rs("DATA_VENDA")%></div></td>
        <td><div align="center"><%=rs("DATA_CADASTRO")%></div></td>
        <td>
			<div align="center">
			  <%
			Set rs2=objCOnexao.Execute("SELECT DATA_PROCESSAMENTO FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_PEDIDO = "&rs("COD_PEDIDO"))
					
					If not rs2.Eof Then
							Response.Write(rs2("DATA_PROCESSAMENTO"))
					Else
							Response.Write("N&Atilde;O PROCESSADO DE FATO")
							
					End If
			
			rs2.Close
			Set rs2=Nothing
			%>
			  
		  </div></td>
        <td><div align="center"><%=rs("VENDEDOR_RECEBE_EM")%></div></td>
      </tr>
    <%
    rs.MoveNext()
    Wend
    %>
</table>
    <br />
    <div style="border-bottom:solid 1px #CCCCCC">
    <table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>H&aacute; pedidos j&aacute; processados para um pagamento futuro a este, n&atilde;o &eacute; poss&iacute;vel continuar</strong></td>
    <td width="8%"><img src="images/icons/icon_red.jpg" width="36" height="39" /></td>
  </tr>
</table>
</div>
<%
Else
%>
<div style="border-bottom:solid 1px #CCCCCC">
    <table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>N&atilde;o h&aacute; pedidos futuros j&aacute; processados</strong></td>
    <td width="8%"><img src="images/icons/icon_green.jpg" width="36" height="39" /></td>
  </tr>
</table>
</div>
<%
End If

rs.Close
Set rs=Nothing
%><br />

<h2 class="style1" style="color: #ff007f !important;">4) Verifica se h&aacute; produtos neste fechamento que n&atilde;o ser&atilde;o processados, pois n&atilde;o possuem comiss&atilde;o cadastrada de todos os produtos que tiveram sa&iacute;da deste fechamento</h2>
<%
Dim qtd_tabelas, VERIFICA_ERRO_tipo4, temp
VERIFICA_ERRO_tipo4 = 2
qtd_tabelas = 0
Set rs=objConexao.Execute("SELECT COUNT(*) as qtd FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES")

If not rs.Eof Then
	qtd_tabelas = formatNumber(0&rs(0),0)/1
End If

rs.Close
Set rs=Nothing

qtd_tabelas = qtd_tabelas/1

sql = "SELECT TB_PRODUTOS.COD_PRODUTO, TB_PRODUTOS.NOME_PRODUTO, COUNT(TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES_PRODUTOS.cod_tabela_comissao_vendedores) AS QTD_TABELAS FROM TB_PRODUTOS LEFT JOIN TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES_PRODUTOS ON TB_PRODUTOS.COD_PRODUTO = TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES_PRODUTOS.COD_PRODUTO WHERE TB_PRODUTOS.PAGAMENTO_COMISSAO = 'S' AND  TB_PRODUTOS.COD_PRODUTO IN (SELECT  cod_produto FROM TB_PEDIDOS_PRODUTOS INNER JOIN TB_PEDIDOS ON TB_PEDIDOS_PRODUTOS.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO WHERE   TB_PEDIDOS.vendedor_recebe_em = "&sqlData(VENDEDOR_RECEBE_EM)&") GROUP BY TB_PRODUTOS.COD_PRODUTO"

Set rs=objConexao.Execute(sql)

If not rs.Eof Then
%>

<table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
        <td width="27%"><div align="center"><strong>C&oacute;d. Produto</strong></div></td>
        <td width="30%"><div align="center"><strong>Produto</strong></div></td>
        <td width="43%"><div align="center"><strong>N&atilde;o Cadastrado na(s) tabela(s)</strong></div></td>
  </tr>
    <%
    While not rs.Eof
		
		temp = formatNumber(0&rs("QTD_TABELAS"),0)
		temp = temp/1
		
		If temp < qtd_tabelas Then
		VERIFICA_ERRO = 1
		VERIFICA_ERRO_tipo4 = 1
    %> 
      <tr>
        <td><div align="center"><%=rs("COD_PRODUTO")%></div></td>
        <td><div align="center"><%=rs("NOME_PRODUTO")%></div></td>
        <td><div align="center">PRODUTO N&Atilde;O CADASTRADO EM TODAS AS TABELAS</div></td>
      </tr>
    <%
		End If
	
    rs.MoveNext()
    Wend
    %>
</table>
    <br />
    
<%
End If

rs.Close
Set rs=Nothing
%>
<%If VERIFICA_ERRO_tipo4 = 1 Then%>
<div style="border-bottom:solid 1px #CCCCCC">
    <table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>H&aacute; produtos que n&atilde;o est&atilde;o cadastrados em todas as tabelas de comissionamento de vendedores.</strong></td>
    <td width="8%"><img src="images/icons/icon_red.jpg" width="36" height="39" /></td>
  </tr>
</table>
</div>
<%Else%>
<div style="border-bottom:solid 1px #CCCCCC">
    <table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>Todos os produtos est&atilde;o cadastrados em todas as tabelas de pre&ccedil;os para este pagamento.</strong></td>
    <td width="8%"><img src="images/icons/icon_green.jpg" width="36" height="39" /></td>
  </tr>
</table>
</div>
<%End If%>
<br />
<br />
<h2 class="style1" style="color: #ff007f !important;">5) Verifica se h&aacute; pedidos deste pagamento que n&atilde;o foram conferidos pelo diretoria </h2>
<%
sql = "SELECT * FROM TB_PEDIDOS WHERE FINALIZA_CORRECAO = 1 AND FINALIZA_DIRETORIA = 2 AND VENDEDOR_RECEBE_EM = "&sqlData(VENDEDOR_RECEBE_EM)
'Response.Write(sql)
Set rs=objConexao.Execute(sql)

If not rs.Eof Then
	
	If autorizacao_diretoria_para_finalizar = "N" Then
		VERIFICA_ERRO = 1
	End If
%>
<table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
        <td width="36%"><div align="center"><strong>C&oacute;d. Pedido</strong></div></td>
        <td width="28%"><div align="center"><strong>Data Venda</strong></div></td>
        <td width="36%"><div align="center"><strong>Data desceu p/ Diretoria</strong></div></td>
  </tr>
    <%
    While not rs.Eof
    %> 
      <tr>
        <td><div align="center"><a href="lca_busca_detalhamento_pedido.asp?cod_pedido=<%=rs("COD_PEDIDO")%>" target="_blank"><%=rs("COD_PEDIDO")%></a></div></td>
        <td><div align="center"><%=rs("DATA_VENDA")%></div></td>
        <td><div align="center"><%=rs("DATA_DESCEU_DIRETORIA")%></div></td>
      </tr>
    <%
    rs.MoveNext()
    Wend
    %>
</table>
<br />
    <div style="border-bottom:solid 1px #CCCCCC">
    <table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>H&aacute; pedidos que n&atilde;o foram conferidos pela diretoria para este pagamento.</strong>
    	
    </td>
    <td width="8%"><img src="images/icons/icon_red.jpg" width="36" height="39" /></td>
  </tr>
</table>

	<%
		If autorizacao_diretoria_para_finalizar = "N" Then
		%>
    		<div class="msgGrandeErro">A diretoria N&Atilde;O aprovou continuar este pagamento sem finalizar a corre&ccedil;&atilde;o de todos os pedidos</div>
    	<%
		Else
		%>
        	<div class="msgGrandeSucesso">A diretoria aprovou continuar este pagamento sem finalizar a corre&ccedil;&atilde;o de todos os pedidos</div>
        <%End If%>

</div>
<%
Else
%>
<div style="border-bottom:solid 1px #CCCCCC">
    <table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>Todos os pedidos para este fechamento foram conferidos pela Diretoria</strong></td>
    <td width="8%"><img src="images/icons/icon_green.jpg" width="36" height="39" /></td>
  </tr>
</table>
</div>
<%
End If

rs.Close
Set rs=Nothing
%><br />

<h2 style="color: #ff007f !important;">6) Pedidos para os pr&oacute;ximos pagamento (ATEN&Ccedil;&Atilde;O PARA DATAS EXORBITANTES)</h2>

<%
sql = "SELECT COD_PEDIDO, DATA_VENDA, DATA_CADASTRO, COD_VENDEDOR, VENDEDOR_RECEBE_EM FROM TB_PEDIDOS WHERE VENDEDOR_RECEBE_EM > "&sqlData(VENDEDOR_RECEBE_EM)&" AND FINALIZA_CORRECAO = 1 AND DATA_VENDA < "&sqlData(VENDEDOR_RECEBE_EM)&" ORDER BY VENDEDOR_RECEBE_EM DESC"
'Response.Write(sql)
Set rs=objConexao.Execute(sql)


If not rs.Eof Then
%>
<table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
        <td width="17%"><div align="center"><strong>C&oacute;d. Pedido</strong></div></td>
        <td width="17%"><div align="center"><strong>Nome Vendedor</strong></div></td>
        <td width="14%"><div align="center"><strong>Data Venda</strong></div></td>
        <td width="21%"><div align="center"><strong>Data Cadastro</strong></div></td>
        <td width="31%"><div align="center"><strong>Vendedor Recebe em:</strong></div></td>
      </tr>
    <%
    While not rs.Eof
    %> 
      <tr>
        <td><div align="center"><a href="lca_busca_detalhamento_pedido.asp?cod_pedido=<%=rs("COD_PEDIDO")%>" target="_blank"><%=rs("COD_PEDIDO")%></a></div></td>
        <td>
			<div align="left">
			  <%
			Set rs2=objConexao.Execute("SELECT NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR = "&rs("COD_VENDEDOR"))
				
				If not rs2.Eof Then	
					Response.Write(rs2("NOME_VENDEDOR"))		
				End If
				
			rs2.Close
			Set rs2=Nothing
			%>
			  
	        </div></td>
        <td><div align="center"><%=rs("DATA_VENDA")%></div></td>
        <td><div align="center"><%=rs("DATA_CADASTRO")%></div></td>
        <td><div align="center"><%=rs("VENDEDOR_RECEBE_EM")%></div></td>
      </tr>
    <%
    rs.MoveNext()
    Wend
    %>
</table>
<%
Else
%>
<div style="border-bottom:solid 1px #CCCCCC">
<table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>N&atilde;o h&aacute; pedidos programados para pagamentos futuros</strong></td>
    <td width="8%"><img src="images/icons/icon_green.jpg" width="36" height="39" /></td>
  </tr>
</table>
</div>
<%
End If

rs.Close
Set rs=Nothing
%>
<br />
<h2 style="color: #ff007f !important;">7) Listagem de Pend&ecirc;ncias de Pagamento para este fechamento</h2>
<%
sql = "SELECT TB_FECHAMENTO_PENDENCIAS_PAGAMENTO.*, TB_VENDEDORES.NOME_VENDEDOR, TB_ADM_USUARIOS.NOME FROM TB_VENDEDORES INNER JOIN TB_FECHAMENTO_PENDENCIAS_PAGAMENTO ON TB_VENDEDORES.COD_VENDEDOR = TB_FECHAMENTO_PENDENCIAS_PAGAMENTO.COD_VENDEDOR INNER JOIN TB_ADM_USUARIOS ON TB_FECHAMENTO_PENDENCIAS_PAGAMENTO.CADASTRADO_POR = TB_ADM_USUARIOS.COD_USUARIO WHERE TB_FECHAMENTO_PENDENCIAS_PAGAMENTO.cod_fechamento_periodo = "&cod_fechamento_periodo&" ORDER BY VENDEDOR_RECEBE_EM DESC"
'Response.Write(sql)
Set rs=objConexao.Execute(sql)


If not rs.Eof Then
%>
<table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="27%"><div align="center"><strong>Vendedor</strong></div></td>
            <td width="10%"><div align="center"><strong>Tipo</strong></div></td>
            <td width="12%"><div align="center"><strong>Valor</strong></div></td>
            <td width="33%"><div align="center"><strong>Motivo</strong></div></td>
            <td width="18%"><div align="center"><strong>Data do Cadastro</strong></div>
     </td>
  </tr>

		  		 <%
	 
		While not rs.Eof
			
				%>
				<tr>
							<td><div align="left"><%=rs("nome_vendedor")%></div></td>
							<td align="center">
<%
							If Ucase(Cstr(rs("TIPO"))) = "A" Then
								Response.Write("Acr�scimo (+)")
							Else
								Response.Write("Decr�scimo (-)")
							End If
							%>                            </td>
							<td><div align="right"><%=FormatCurrency(rs("VALOR"))%></div></td>
							<td><div align="left"><%=rs("MOTIVO")%></div></td>
							<td><div align="center"><%=ArrumaData(rs("DATA_CADASTRO"))%> / <%=rs("NOME")%> </div></td>
				</tr>
		        <%
			rs.MoveNext()
			Wend	
      
		  %>
</table>
<%Else%>
<table width="600" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="92%" class="textoMedio"><strong>N&atilde;o h&aacute; pend&ecirc;ncias para este fechamento</strong></td>
    <td width="8%"><img src="images/icons/icon_green.jpg" width="36" height="39" /></td>
  </tr>
</table>
<%
End If

rs.Close
Set rs=Nothing
%>
<br /><br />
<!--
<div style="border-bottom:solid 1px #CCCCCC; background-color:#FFFFCC">

<form name="formulario_padrao" action="adm_rotina_fechamento_extrato_pagamento_part2_tipo2.asp" style="margin:0px">
<input type="hidden" name="VENDEDOR_RECEBE_EM" value="<%=VENDEDOR_RECEBE_EM%>" />

'FORMA DE PAGAMENTO NOVA N�O � MAIS PERMITIDO ALTERA��O DO DIA DE PAGAMENTO
<table width="600" border="0" cellpadding="5" cellspacing="0">
  <tr>
    <td width="52%" class="textoMedio"><strong>Escolha o dia do pagamento dos recibos</strong></td>
    <td width="48%">
  		<select name="DIA_VENCIMENTO" class="inputTamanho1">
			<option value="10">10</option>
        	<option value="11">11</option>
            <option value="12">12</option>
        	<option value="13">13</option>
            <option value="14">14</option>
        	<option value="15" selected="selected">15</option>
            <option value="16">16</option>
            <option value="17">17</option>
            <option value="18">18</option>
            <option value="19">19</option>
            <option value="20">20</option>
        </select>
    </td>
  </tr>
</table>

</form>
</div>
-->
<br />
<br />
<h2>&nbsp;</h2>

<table width="100%"  border="1" align="center" cellpadding="5" style="border-collapse:collapse">	
                    
					
                    <tr>
               			<td colspan="15"><div align="center" style="border-bottom:solid 2PX #CCCCCC; margin-bottom:15PX;"><strong>QTD. TOTAL DE PEDIDOS N&Atilde;O PROCESSADOS:</strong> <span class="style2">
               			 <%
							Set rs2=objConexao.Execute("SELECT COUNT(*) AS QTD_PEDIDOS FROM TB_PEDIDOS WHERE FINALIZA_CORRECAO = 1 AND PAGAMENTO_PROCESSADO = 2 AND VENDEDOR_RECEBE_EM = "&sqlData(VENDEDOR_RECEBE_EM))
				
								If not rs2.Eof Then
									Response.Write(rs2("QTD_PEDIDOS"))
									temp = formatNumber(0&rs2(0),0)
								End If
				
							rs2.Close
							Set rs2=Nothing
						%>
                	</span></div>
                  <div align="center"><strong>QTD. TOTAL DE PEDIDOS TOTAL:</strong> <span class="style2">
                 	 <%
						Set rs2=objConexao.Execute("SELECT COUNT(*) AS QTD_PEDIDOS FROM TB_PEDIDOS WHERE FINALIZA_CORRECAO = 1 AND VENDEDOR_RECEBE_EM = "&sqlData(VENDEDOR_RECEBE_EM))
						
							If not rs2.Eof Then
								Response.Write(rs2("QTD_PEDIDOS"))
							
							End If
						
						rs2.Close
						Set rs2=Nothing
						%>
                </span></div></td>
  </tr>
                 
                 	<tr>
    <td height="31" colspan="15" align="center">
        <% If int(VERIFICA_ERRO) = 2 and int(temp) > 0 Then %>

            <button id="btnProcessar" style="padding: 10px 20px; background-color: #ff007f; color: white; border: none; border-radius: 5px; cursor: pointer;">
                Processar Pagamento
            </button>

            <div id="statusProcessamento" style="margin-top: 10px;"></div>

        <% Else %>
            <strong> N&Atilde;O &Eacute; POSS&Iacute;VEL GERAR PAGAMENTO. VERIFIQUE ERROS ACIMA ANTES DE CONTINUAR!</strong>
        <% End If %>
    </td>
</tr>
			</table>
<br />
<br />     
          
<%
FecharConexao()
%>
<script>
document.addEventListener('DOMContentLoaded', function() {
    const btn = document.getElementById('btnProcessar');
    if (btn) {
        btn.addEventListener('click', function() {
            const statusDiv = document.getElementById('statusProcessamento');
            statusDiv.innerHTML = 'Processando...';

            fetch('adm_rotina_fechamento_extrato_pagamento_part2_tipo2.asp?COD_FECHAMENTO_PERIODO=<%=cod_fechamento_periodo%>&VENDEDOR_RECEBE_EM=<%=vendedor_recebe_em%>')
                .then(response => response.text())
                .then(data => {
                    statusDiv.innerHTML = data; // exibe "Processamento concluído com sucesso!"
                })
                .catch(error => {
                    statusDiv.innerHTML = 'Erro no processamento: ' + error;
                });
        });
    }
});
</script>
</body>
</html>



