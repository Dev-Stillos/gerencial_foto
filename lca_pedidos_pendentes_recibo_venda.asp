<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim pedidos, cod_vendedor, nome_vendedor, qtd_pedidos

pedidos			= Request.QueryString("pedidos")
cod_vendedor	= Request.QueryString("cod_vendedor")

AbrirConexao()

	Set rs=objConexao.Execute("Select nome_vendedor FROM TB_VENDEDORES where cod_vendedor='"&cod_vendedor&"'")
			
			If not rs.Eof Then
				
				nome_vendedor = rs("nome_vendedor")
				
			End If
			
	rs.Close
	Set rs=Nothing
	
FecharConexao()

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="print"></link>
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<style type="text/css" media="print">
<!--
.folha {
    page-break-after: always;
}
.btnFinalizaProducao{	background-image:url(images/layout/btnFinalizaProducao.jpg);
	cursor:pointer;
	width:140px;
	height:22px;
	border:0px;
	}
body,td,th {
	font-size: 10px;
}
.style1 {font-size: 12px}
-->
</style>

<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
    <style type="text/css">
<!--
.style1 {font-weight: bold}
.style2 {font-size: 10px}
-->
    </style>
</head>

<body onload="window.print()" style="background:#FFFFFF">

<table class="folha" width="100%" align="center" border="0" cellpadding="0" cellspacing="0">
<tr>
<td>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0">
      <tr>
        <td valign="top"><h2 style="color:#000000; border-bottom:solid #000 3px; padding-bottom:2px">Recibo de Entrega de Vendas</h2>
          <br />
               <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="90%"><table width="100%" border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td width="245" class="textoMedio">Data da entrada das vendas:</td>
                  <td width="494"><strong class="textoMedio"><%=ArrumaData(Date())%></strong></td>
                </tr>
                <tr>
                  <td class="textoMedio">Representante Comercial:</td>
                  <td class="textoMedio"><strong><%=NOME_VENDEDOR%></strong></td>
                </tr>
              </table></td>
              <td width="10%"><div align="center"></div></td>
            </tr>
          </table></td>
        </tr>
      <tr>
        <td>   
		       <table width="100%" border="0" align="center" 	cellpadding="2" cellspacing="0" style="border-collapse:collapse">
			<tr bgcolor="whitesmoke">
                    <td width="9%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><table width="100%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td class="style2 textoMenor textoMenor"><div align="center"><strong>C&oacute;d Pedido</strong> </div>
                        <div align="center"></div>                        <div align="right"></div></td>
                      </tr>
                    </table></td>
                    <td width="22%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style2 textoMenor textoMenor"><strong>C&oacute;d. Formando(s) / Formando(s)</strong> </div></td>
                    <td width="4%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style2 textoMenor textoMenor"><strong>Lan&ccedil;. do Pedido no sistema</strong></div></td>
                    <td width="4%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style2 textoMenor textoMenor"><strong>Total Recebido</strong></div></td>
			</tr>
		
						
						<%
						AbrirConexao()
						pedidos = Replace(pedidos," ","")
						pedidos = Split(pedidos,",")
						
						Dim DATA_CADASTRO
						
						For i = 0 to Ubound(pedidos)
						
						Set rs=objConexao.Execute("Select DATA_CADASTRO FROM TB_PEDIDOS WHERE COD_PEDIDO = '"&pedidos(i)&"'")
							If not rs.Eof Then
								DATA_CADASTRO = rs("DATA_CADASTRO")
							End If
						
						rs.Close
						Set rs=Nothing
						%>
          					<tr>
                    <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333; border-left:solid 0.25px #333333"><div align="center" class="style2 textoMenor textoMenor">
          
                      
					      <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><div align="center"><strong><%=pedidos(i)%></strong></div>      <div align="center"></div></td>
    </tr>
</table>
						
						
						
                        </div></td>
                    <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333; border-left:solid 0.25px #333333"><div align="left" class="style2 textoMenor textoMenor">
                        <%
			sql = 	"SELECT TB_PEDIDOS_FORMANDOS.COD_FORMANDO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO "&_
					"FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON "&_
					"TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO "&_
					"WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&pedidos(i)&" ORDER BY NOME_FORMANDO"
			Set rs3 = objConexao.Execute(sql)
				
		
				While not rs3.Eof
								
				%>
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="23%"><div align="center"><%=digito(rs3("COD_FORMANDO"))%></div></td>
                            <td width="57%"><div align="left"><%=rs3("NOME_FORMANDO")%></div></td>
                          </tr>
                        </table>
                      <%
				rs3.MoveNext()
				Wend
				
			rs3.Close
			Set rs3 = Nothing
			%>
                    </div></td>
                    <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333; border-left:solid 0.25px #333333"><div align="center" class="style2 textoMenor textoMenor"><%=ArrumaData(DATA_CADASTRO)%></div></td>
                    <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333; border-left:solid 0.25px #333333"><div align="center"><span class="style2"><span class="textoMenor"><span class="textoMenor"></span></span></span><%
			Set rs4 = objConexao.Execute("SELECT * FROM VW_PEDIDOS_TOTAL_RECEBIDO WHERE COD_PEDIDO="&pedidos(i))
			Dim PAGAMENTO_EM_DINHEIRO, XTOTAL_CHEQUE, XTOTAL_CARTAO, TOTAL_PAGO
			
			TOTAL_PAGO = 0
				
				If not rs4.Eof Then
				
					PAGAMENTO_EM_DINHEIRO	= formatNumber(0&rs4("PAGAMENTO_EM_DINHEIRO"))/1
					XTOTAL_CHEQUE			= formatNumber(0&rs4("XTOTAL_CHEQUE"))/1
					XTOTAL_CARTAO			= formatNumber(0&rs4("XTOTAL_CARTAO"))/1
					TOTAL_PAGO				= (PAGAMENTO_EM_DINHEIRO + XTOTAL_CHEQUE + XTOTAL_CARTAO)/1
					
					Response.Write(formatNumber(TOTAL_PAGO))

				End If
			
			rs4.Close()
			Set rs4 = Nothing
			%></div></td>
   					</tr>
   					<%
						Next
						
						FecharConexao()
						%>
						</table>
	        <p>Qtd de vendas entregues: <strong><%=Ubound(pedidos)+1%></strong><br />
	        </p>
	      <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
           <tr>
             <td style="border-bottom:solid 1px #000000"><p>&nbsp;</p>
               <p>&nbsp;</p></td>
           </tr>
           <tr>
             <td><div align="center"><span class="textoMedio"><span class="style1">REPRESENTANTE COMERCIAL</span><br />
                   <span class="style1">  <strong><%=UCase(NOME_VENDEDOR)%></strong></span></span></div></td>
           </tr>
         </table>
		   <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
           <tr>
             <td style="border-bottom:solid 1px #000000"><p>&nbsp;</p>
               <p>&nbsp;</p></td>
           </tr>
           <tr>
             <td><div align="center"><span class="textoMedio">RECEBIDO POR<strong><br />
               <%=Ucase(Session("NOME"))%></strong></span></div></td>
           </tr>
         </table>
		<!-- #include file="includes/rodape_impressao.asp" -->
	
		 </td>
        </tr>


    </table></td>
  </tr>
</table>
</td>
</tr>
</table>

</body>
</html>	