<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Validação especial para verificar se possui sub-validação específica
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '72' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se não encontrar então redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
		
	rs.Close
	Set rs=Nothing
	
FecharConexao()


Dim MES_INICIO, ANO_INICIO, MES_FIM, ANO_FIM
Dim QTD_DIAS, QTD_MESES, PERIODO, PERIODO_INICIAL, PERIODO_FINAL, DESMEMBRAR
Dim SUBTOTAL_CARTAO, SUBTOTAL_CHEQUE, SUBTOTAL_DINHEIRO, SUBTOTAL
DIm SUBVENDTOTAL_CARTAO, SUBVENDTOTAL_CHEQUE, SUBVENDTOTAL_DINHEIRO, SUBVENDTOTAL
Dim TOTAL_CARTAO, TOTAL_CHEQUE, TOTAL_DINHEIRO, TOTAL

SUBTOTAL_CARTAO = 0
SUBTOTAL_CHEQUE = 0
SUBTOTAL_DINHEIRO = 0
SUBTOTAL = 0
SUBVENDTOTAL_CARTAO = 0
SUBVENDTOTAL_CHEQUE = 0
SUBVENDTOTAL_DINHEIRO = 0
SUBVENDTOTAL = 0
TOTAL_CARTAO = 0
TOTAL_CHEQUE = 0
TOTAL_DINHEIRO = 0
TOTAL = 0

DESMEMBRAR = Request.QueryString("desmembrar")
MES_INICIO = Request.QueryString("MES_INICIO")
ANO_INICIO = Request.QueryString("ANO_INICIO")
MES_FIM = Request.QueryString("MES_FIM")
ANO_FIM = Request.QueryString("ANO_FIM") 
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Versão 1.2</title>
</head>
	<script language="javascript">
		function mOvr(src, cor)
			{
			if (!src.contains(event.fromElement))
				{
					src.style.cursor = 'hand';
					src.bgColor = cor;
				}
			}
		
		function mOut(src, cor)
			{
				if (!src.contains(event.toElement))
					{
						src.style.cursor = 'default';
						src.bgColor = cor;
			
					}
			}
				
			function abre(url)
			{    
						//if(event.srcElement.tagName=='TD')
						//	{
						//		src.children.tags('A')[0].click(); 
						//	}
			location.href = url;
		
			}	
	</script>
<body>
<table cellpadding="9px" cellspacing="0" width="100%">
<tr>
<td>
<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td>
		<!-- #include file="includes/topo.asp" -->
<table width="100%" style="border-bottom:#999999 solid dotted 1px" background="images/layout/fundo_menu_azul.jpg"border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
	<!-- #include file="includes/menu.asp" -->
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="10">
  <tr>
  <td>
   <%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 2
  BTN_FONTE = 1
  BTN_IMPRIMIR_LINK = "#"
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  <%
  AbrirConexao()
  %>
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">
	  <h2>Relatórios - Total Vendas por Forma de Pagamento</h2>
      <!--
	  <%if Request.QueryString("desmembrar") <> 1 then%>
	  <div class="subMenu">Sub-Menu: <a href="?desmembrar=1&MES_INICIO=<%=MES_INICIO%>&ANO_INICIO=<%=ANO_INICIO%>&filtro2=+&MES_FIM=<%=MES_FIM%>&ANO_FIM=<%=ANO_FIM%>">Desmembrar por vendedor</a></div>
	  <%else%>
	  <div class="subMenu">Sub-Menu: <a href="?desmembrar=2&MES_INICIO=<%=MES_INICIO%>&ANO_INICIO=<%=ANO_INICIO%>&filtro2=+&MES_FIM=<%=MES_FIM%>&ANO_FIM=<%=ANO_FIM%>">Agrupar por mês</a></div>
	  <%end if%>
      -->
	  
	  <%
		PERIODO_INICIAL = formatDateTime("01"&"/"&MES_INICIO&"/"&ANO_INICIO)
		PERIODO_FINAL = formatDateTime("01"&"/"&MES_FIM&"/"&ANO_FIM)
		PERIODO_FINAL = formatDateTime(QuantidadeDiasMes(PERIODO_FINAL)&"/"&MES_FIM&"/"&ANO_FIM)
		
		If (DateDiff("d",PERIODO_INICIAL,PERIODO_FINAL)) < 0 then
			Response.Redirect("adm_relatorios.asp?msg=Você selecionou uma data inicial menor que a data final!")
		End If			 
			
		PERIODO = PERIODO_INICIAL
		QTD_MESES = DateDiff("m",PERIODO_INICIAL,PERIODO_FINAL)
		'Response.Write(QTD_MESES)
		%>
		<h3 align="center"><%Response.Write(Ucase(MonthName(MES_INICIO))&" de "&ANO_INICIO)%> à <%Response.Write(Ucase(MonthName(MES_FIM)))&" de "&ANO_FIM%></h3>
		<div id="comboTexto">
        <table width="98%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
        <tr bgcolor="whitesmoke">
              <td width="24%"  bgcolor="whitesmoke"><div align="center"><strong>Ano </strong><strong>/ Mês </strong></div></td>
              <td bgcolor="#FCF8BC"><div align="center"><strong>Cartão de Crédito</strong></div></td>
              <td  bgcolor="#D6FED8"><div align="center"><strong>Cheque</strong></div>                <div align="center"></div></td>
              <td  bgcolor="#D3E3FE"><div align="center"><strong>Dinheiro</strong></div>                <div align="center"></div></td>
              <td  bgcolor="whitesmoke"><div align="center"><strong> Total Vendas</strong></div></td>
            </tr>	

        <%	
		PERIODO = formatDateTime("01"&"/"&MES_INICIO&"/"&ANO_INICIO)
			
		For i=0 to qtd_meses
			
			PERIODO_INICIAL = PERIODO
			PERIODO_FINAL = formatDateTime(QuantidadeDiasMes(PERIODO_INICIAL)&"/"&MONTH(PERIODO_INICIAL)&"/"&YEAR(PERIODO_INICIAL))
			
			SUBTOTAL_CARTAO = 0
			SUBTOTAL_CHEQUE = 0
			SUBTOTAL_DINHEIRO = 0
			SUBTOTAL = 0
				
			sql = "SELECT SUM(IFNULL(PAGAMENTO_EM_DINHEIRO,0)) as dinheiro, SUM(IFNULL(XTOTAL_CHEQUE,0)) as cheque, SUM(IFNULL (XTOTAL_CARTAO,0)) as cartao, SUM(IFNULL (XTOTAL,0)) as total FROM VW_PEDIDOS_TOTAL_RECEBIDO WHERE cod_pedido IN (SELECT cod_pedido FROM TB_PEDIDOS WHERE FINALIZA_CORRECAO = 1 AND DATA_VENDA BETWEEN "&sqlData(PERIODO_INICIAL)&" AND "&sqlData(PERIODO_FINAL)&")"
			'Response.Write(sql)
			Set rs=objConexao.Execute(sql)
				If not rs.Eof Then
					SUBTOTAL_CARTAO = formatNumber(0&rs("cartao"),2)
					SUBTOTAL_CHEQUE = formatNumber(0&rs("cheque"),2)
					SUBTOTAL_DINHEIRO = formatNumber(0&rs("dinheiro"),2)
					SUBTOTAL = formatNumber(0&rs("total"),2)
				End If

			rs.Close
			Set rs=Nothing
			
			TOTAL_CARTAO = (TOTAL_CARTAO+SUBTOTAL_CARTAO)/1
			TOTAL_CHEQUE = (TOTAL_CHEQUE+SUBTOTAL_CHEQUE)/1
			TOTAL_DINHEIRO = (TOTAL_DINHEIRO+SUBTOTAL_DINHEIRO)/1
			TOTAL = (TOTAL+SUBTOTAL)/1

		%>
				<tr>
              			<td><div align="left"><%=year(PERIODO)&" / "&Ucase(MonthName(Month(PERIODO)))%></div></td>
              			<td width="18%"><div align="right"><%=formatNumber(SUBTOTAL_CARTAO,2)%></div></td>
              			<td width="20%"><div align="right"><%=formatNumber(SUBTOTAL_CHEQUE,2)%></div></td>
				 		<td width="18%"><div align="right"><%=formatNumber(SUBTOTAL_DINHEIRO,2)%></div></td>
				 		<td width="20%"><div align="right"><strong><%=formatNumber(SUBTOTAL,2)%></strong></div></td>
				 	</tr>
        <%
		PERIODO = DateAdd("m",1,PERIODO)
		
		Next
		%>
          <tr bgcolor="#C4E0FF">
                  <td bgcolor="whitesmoke"><h3 align="center" style="margin:0px">Totais</h3></td>
                  <td bgcolor="whitesmoke"><div align="right" class="textoMedio"><strong><%=formatNumber(TOTAL_CARTAO,2)%></strong></div></td>
                  <td bgcolor="whitesmoke"><div align="right" class="textoMedio"><strong><%=formatNumber(TOTAL_CHEQUE,2)%></strong></div></td>
                  <td bgcolor="whitesmoke"><div align="right" class="textoMedio"><strong><%=formatNumber(TOTAL_DINHEIRO,2)%></strong></div></td>
                  <td bgcolor="whitesmoke"><div align="right" class="textoMedio"><strong><%=formatNumber(TOTAL,2)%></strong></div></td>
              </tr>
		 	
       </table>
		</div>
		

		<br />
		<strong style="color:#900">Observação: Valores totais da venda e não desmembradas de acordo com parcelas/vencimento.</strong>
        
<!--
<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="98%" border="0" align="center" cellpadding="5" cellspacing="0">
          <tr>
            <td><p><strong>Taxas das Operadoras de Cartão de Crédito<br />
            </strong></p>
              <p> <strong>REDECARD (MASTERCARD)</strong><br />
                À vista – 2,90%<br />
Parcelado de 2 à 6 – 3,90%<br />
Parcelado de 7 à 12 – 4,40%<br />
<br />
<strong>VISA</strong><br />
VISA ELECTRON - Débito – 2%<br />
VISA - Crédito – 1X – 2,70%<br />
VISA – Crédito – 2 à 12 – 3,70%</p></td>
            </tr>
        </table>
-->
		<p align="center">
		<input class="btnTopo" onclick="abre('#')" type="button" name="btnTopo2" value=" " />
		</p>		</td>
    </tr>
</table>

 
  
  </td>
    </tr>
</table>

  
  </td>
  </tr>
</table>
<!-- #include file="includes/rodape.asp" -->
<%FecharConexao()%>
  </td>
</tr>
</table>
</body>
</html>






