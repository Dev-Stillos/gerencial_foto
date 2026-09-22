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

If DATA_VENDA_INICIO = Empty or DATA_VENDA_FIM = Empty Then
	Response.Redirect("adm_relatorios.asp?msg=Informe uma data de venda de in�cio e outra de fim para gerar o relat�rio.")
End If

sqlQuery = "DATA_VENDA BETWEEN "&sqlData(DATA_VENDA_INICIO)&" AND "&sqlData(DATA_VENDA_FIM)


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="all"></link>
<title>StillosWEB - Sistema corporativo - Versão 1.2</title>
</head>
<style type="text/css" media="all">
<!--
.folha {
    page-break-after: always;
}
-->
</style>
<body style="background-color:#FFFFFF" onload="window.print();">
 

	  	
      
	     <%Dim rs, sql, rs2, sql2,sql3, asdasd,asd, AlbunsTotais
		 
		 AlbunsTotais = 0

		 
		  	AbrirConexao()		  
 		
		sql2 =  "SELECT "&_
				"COD_VENDEDOR, "&_
				"NOME_VENDEDOR, "&_
				"(SUM(IFNULL(PAGAMENTO_EM_DINHEIRO,0)) + "&_
				"SUM(IFNULL(XTOTAL_CHEQUE,0)) + "&_
				"SUM(IFNULL(XTOTAL_CARTAO,0))) "&_
				"AS TOTAL_RECEBIDO_GERAL "&_
				"FROM VW_EXTRATO_TOTAL_RECEBIDO_COMPLETO WHERE "&_
				sqlQuery&_
				"GROUP BY COD_VENDEDOR "&_
				"ORDER BY TOTAL_RECEBIDO_GERAL DESC "
		



	'Response.Write(sql2)
					
			Set rs2 = objConexao.Execute(sql2)
			Dim num_p, VendaTotal
			VendaTotal = 0
			num_p = 0%>
<%If rs2.Eof then
Response.Redirect("adm_relatorios.asp?msg=Não há dados para o ranking de vendedores do período selecionado, selecione outro período!")%> 	

		
<%else%>
        	
<div class="folha"> 
      	 	    <h3 align="center"><strong> Ranking por  arrecadação - <%=ArrumaData(DATA_VENDA_INICIO)%> à <%=ArrumaData(DATA_VENDA_FIM)%></strong> </h3>
	  <table width="80%" 
		      border="1" align="center"
					cellpadding="5" 
					style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
                <td><div align="center"><span class="MensagemFicha"><strong>P</strong></span><strong><span class="MensagemFicha">osição</span></strong></div></td>
                <td><div align="left" class="MensagemFicha">
                  <div align="center"><strong>Representante Comercial</strong></div>
                </div></td>
        </tr>
		     
			<%While not rs2.Eof
			num_p = num_p + 1
			    %>
                     
				<tr>
		          <td width="10%"><div align="center"><%=num_p%>°</div></td>

                  <td><div align="left"><%=Ucase(rs2("NOME_VENDEDOR"))%></div></td>
        </tr>
					 
					 
          <%
		  	rs2.MoveNext()
			Wend
			
			rs2.Close
			Set rs2 = Nothing
		  %>
  </table>
  <br />

   <div align="center">Impresso em: <strong><%=now()%></strong></div>

  
  
</div>
        <div class="folha">
		  <div align="center">
	      	  <h3><strong><br />Ranking por número de álbuns vendidos - <%=ArrumaData(DATA_VENDA_INICIO)%> à <%=ArrumaData(DATA_VENDA_FIM)%></strong> </h3>
		  </div>
	      	 		<table width="80%" 
		      border="1" align="center"
					cellpadding="5" 
					style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
                <td><div align="center"><span class="MensagemFicha"><strong>P</strong></span><strong><span class="MensagemFicha">osição</span></strong></div></td>
                <td><div align="left" class="MensagemFicha">
                  <div align="center"><strong>Representante Comercial</strong></div>
                </div></td>
                <td width="25%">
                  
                  <div align="center" class="MensagemFicha"><strong>Total</strong></div></td></tr>
		      <%	 
	
		sql3 =	"SELECT TB_PEDIDOS.COD_VENDEDOR, SUM( VW_QTD_FORMANDOS_POR_PEDIDO.QTD_FORMANDOS ) AS TOTAL_ALBUNS "&_
				"FROM VW_QTD_FORMANDOS_POR_PEDIDO "&_
				"INNER JOIN TB_PEDIDOS ON VW_QTD_FORMANDOS_POR_PEDIDO.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO "&_
				"WHERE TB_PEDIDOS."&sqlQuery&_
				" AND FINALIZA_CORRECAO =1 "&_
				" GROUP BY COD_VENDEDOR "&_
				" ORDER BY TOTAL_ALBUNS DESC  "

			'Response.Write(sql3)
					
			Set rs3 = objConexao.Execute(sql3)
			Dim num_p2
			num_p2 = 0
			While not rs3.Eof
			num_p2 = num_p2 + 1
			    %>
                     
				<tr>
		          <td width="15%"><div align="center"><%=num_p2%>°</div></td>
                  <td width="60%"><div align="left"><%
				  Dim rsVendedorNome
				  Set rsVendedorNome = objConexao.Execute("SELECT NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR = "&rs3("COD_VENDEDOR"))
				  			
							Response.Write(UCASE(rsVendedorNome("NOME_VENDEDOR")))
							
				  rsVendedorNome.Close
				  Set rsVendedorNome = Nothing
				  %></a></div></td>
                  <td><div align="center"><%=rs3("TOTAL_ALBUNS")%>
				  <%
				  subAlbuns = formatNumber(rs3("TOTAL_ALBUNS"))
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
                              <%AlbunsTotais = Replace(AlbunsTotais,",","")%>
                              <%Response.Write(AlbunsTotais)%>
                            </strong></div></td>
		  	</table>
            <br />

             <div align="center">Impresso em: <strong><%=now()%></strong></div>
            </div> 
            
<%
AlbunsTotais = 0
end if


%>

    
            
            
            
            
            <%
			sql = "SELECT TB_PEDIDOS_PRODUTOS.COD_PRODUTO, TB_PRODUTOS.NOME_PRODUTO, SUM(TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS TOTAL_QUANTIDADE_GERAL, SUM(TB_PEDIDOS_PRODUTOS.VALOR_UN*TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS VALOR_TOTAL FROM TB_PEDIDOS  INNER JOIN TB_PEDIDOS_PRODUTOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO INNER JOIN TB_PRODUTOS ON TB_PEDIDOS_PRODUTOS.COD_PRODUTO = TB_PRODUTOS.COD_PRODUTO WHERE "&sqlQuery&" and TB_PEDIDOS.FINALIZA_CORRECAO = 1 GROUP BY TB_PEDIDOS_PRODUTOS.COD_PRODUTO ORDER BY TOTAL_QUANTIDADE_GERAL DESC"
			
			'Response.Write(sql)
			
			Set rs=objConexao.Execute(sql,QTD_PRODUTOS)
			
			
			
			
				If not rs.Eof Then
				
							While not rs.Eof
																
								%>
                                <div class="folha">
                                 <div align="center">
	      	 						<h3><strong>Ranking de <%=Ucase(rs("NOME_PRODUTO"))%> - <%=ArrumaData(DATA_VENDA_INICIO)%> à <%=ArrumaData(DATA_VENDA_FIM)%></strong></h3></div>
                                    <table width="80%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">   
                                    	<tr bgcolor="whitesmoke">
                    						<td align="center"><strong>Posição</strong></td>
                    						<td align="center"><strong>Representante Comercial</strong></td>
                    						<td width="25%" align="center"><strong>Total</strong></td></tr>
                  <%	 
       			
				i = 0
				total = 0
	   
	   			sql = "SELECT TB_VENDEDORES.NOME_VENDEDOR, SUM(TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS TOTAL_QUANTIDADE, SUM(TB_PEDIDOS_PRODUTOS.VALOR_UN*TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS VALOR_TOTAL FROM TB_PEDIDOS INNER JOIN TB_PEDIDOS_PRODUTOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO INNER JOIN TB_PRODUTOS ON TB_PEDIDOS_PRODUTOS.COD_PRODUTO = TB_PRODUTOS.COD_PRODUTO INNER JOIN TB_VENDEDORES ON TB_PEDIDOS.COD_VENDEDOR = TB_VENDEDORES.COD_VENDEDOR WHERE "&sqlQuery&" and TB_PEDIDOS.FINALIZA_CORRECAO = 1 AND TB_PEDIDOS_PRODUTOS.COD_PRODUTO = "&rs("COD_PRODUTO")&"  GROUP BY TB_PEDIDOS.COD_VENDEDOR ORDER BY  TOTAL_QUANTIDADE DESC"
				   
				   
				                                
                Set rs2 = objConexao.Execute(sql)

				
                While not rs2.Eof
               		
					i = i + 1
                    subtotal = formatNumber(0&rs2("TOTAL_QUANTIDADE"),0)/1
                    total =  (total+subtotal)/1
                    %>
                         
                    <tr>
                      <td width="15%"><div align="center"><%=i%>°</div></td>
                      <td width="60%"><div align="left"><%=Ucase(rs2("NOME_VENDEDOR"))%></div></td>
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
                </table>
                                	
                                      <br />
  	Obs.: Somente produtos vendidos são contabilizados.<br />
    <div align="center">Impresso em: <strong><%=now()%></strong></div>
    
    			</div>
    
                                <%
								
								
							
							rs.MoveNext()
							Wend
				
				End if
			
			
			rs.Close
			Set rs=Nothing
			
			FecharConexao()
%>
</body>
</html>
