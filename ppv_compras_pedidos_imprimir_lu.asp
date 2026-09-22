<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<!-- #include file="files/includes/functions_local.asp" -->
<%
Dim cod_pedido_compra, numero_pedido_compra, cod_fornecedor, nome_fantasia, aprovado_cod_usuario, data_aprovado
Dim qtd_itens_total

cod_pedido_compra = Request.QueryString("c")


AbrirConexao()

	Set rs=objConexao.Execute("SELECT * FROM TB_PEDIDOS_COMPRAS WHERE cod_pedido_compra = "&cod_pedido_compra)
			
			If not rs.Eof Then
				
				numero_pedido_compra=rs("numero_pedido_compra")
				cod_fornecedor=rs("cod_fornecedor")
				aprovado_cod_usuario=rs("aprovado_cod_usuario")
				data_aprovado=rs("data_aprovado")
				valor_total_produtos = pedidoCompraTotais(cod_pedido_compra,1,true)
				contato=rs("contato")
				
				
				Set rs2=objConexao.Execute("SELECT nome_fantasia FROM TB_FORNECEDORES WHERE cod_fornecedor = '"&cod_fornecedor&"'")
					
					If not rs2.Eof Then
						nome_fantasia = rs2("nome_fantasia")
					End If
				
				rs2.Close
				Set rs2=Nothing
							
			End If
	
	rs.Close
	Set rs=Nothing
	
'Será fechado no final da tela a conexão


%>

   <style type="text/css" media="all">
	body{
		margin:0px;
		font-family:Arial, Helvetica, sans-serif;
		color:#000;
		font-size:11px;	
	}
	.folha{
	}
	.recibo{
		border:1px solid #000; 	
	}
	.imagemRecibo{
		position:relative;
		float:left;
		width:799px;
		height:560px;
	}
	.imagemDentro{
		position:absolute;
		top:0px;
		left:0px;
		z-index:0;
	}
	.imagemConteudo{
	z-index:100;
	position:absolute;
	top:20px;
	left:18px;
	width: 760px;
	height: 382px;
	}
	.numero_protocolo{
		padding-top:50px; font-size:23px; text-align:center;	
	}
    </style>
</head>
<script>
function imprimeFecha(){
	window.print();
}
</script>
<%
Dim subtotal, total
subtotal = 0
total = 0
%>
<body onLoad="imprimeFecha()" style="background:#FFFFFF">
	<table class="folha" width="100%" align="center" border="0" cellpadding="0" cellspacing="10">
					<tr>
					<td>
					<table width="100%" border="0" align="center" cellpadding="10" cellspacing="0" bgcolor="#FFFFFF">
					  <tr>
						<td><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0">
						  <tr>
							<td valign="top"><h2 style="color:#000000; border-bottom:solid #000 3px; padding-bottom:2px">Recibo de Entrega de Produtos</h2>
						
							  <table width="100%" border="0" cellspacing="0" cellpadding="0">
								<tr>
								  <td width="82%"><table width="100%" border="0" cellspacing="0" cellpadding="2">
									<tr>
									  <td width="142"><span style="font-size:13px">Fornecedor</span></td>
									  <td width="623" style="font-size:13px"><%=Ucase(nome_fantasia)%></td>
									</tr>
									<tr>
									  <td style="font-size:13px">Data da Entrega</td>
									  <td style="font-size:13px"><%=ArrumaData(date())%></td>
								    </tr>
     
								  </table></td>
								  <td width="18%"><div align="center"><img src="images/layout/logomarca_foto_reportagens.jpg" width="92" height="93" /></div></td>
								</tr>
							  </table></td>
							</tr>
						  <tr>
							<td> 
                            
                            <%
							'Agrupa por instituição de ensino
							Set rs=objConexao.Execute("SELECT VW_TB_ORDENS_SERVICO.COD_CONTRATO, VW_TB_ORDENS_SERVICO.ENTIDADE_ENSINO FROM VW_TB_ORDENS_SERVICO INNER JOIN VW_TB_PEDIDOS_COMPRAS_PRODUTOS on VW_TB_ORDENS_SERVICO.COD_ORDEM_SERVICO = VW_TB_PEDIDOS_COMPRAS_PRODUTOS.COD_ORDEM_SERVICO WHERE VW_TB_PEDIDOS_COMPRAS_PRODUTOS.cod_pedido_compra = "&sqlTexto(cod_pedido_compra)&" GROUP BY VW_TB_ORDENS_SERVICO.COD_CONTRATO ORDER BY VW_TB_ORDENS_SERVICO.ENTIDADE_ENSINO ASC")
							
								While not rs.Eof%>
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr>
				  <th colspan="8" style="padding-top:5px;"><div align="left" style="font-size:12px"><%=Ucase(rs("ENTIDADE_ENSINO"))%></div></th>
				  </tr>
		
				<tr style="font-size:11px">
				
				  <td width="13%"><div align="center"><strong>C&oacute;d. Pedido de Compra</strong></div></td>
				  <td width="29%"><div align="center"><strong>Formando</strong></div></td>
				  <td width="26%"><div align="center"><strong>Produto</strong></div></td>
				  <td width="8%"><div align="center"><strong>Qtd. Entregue</strong></div></td>
				  <td width="9%"><div align="center"><strong>Valor Un.</strong></div></td>
				  <td width="15%" colspan="2"><div align="center"><strong>Valor Total</strong></div></td>
				  
				  </tr>
        	   	  <%
				subtotal = 0
				Set rs2=objConexao.Execute("SELECT VW_TB_ORDENS_SERVICO.NOME_FORMANDO, VW_TB_PEDIDOS_COMPRAS_PRODUTOS.* FROM VW_TB_ORDENS_SERVICO INNER JOIN VW_TB_PEDIDOS_COMPRAS_PRODUTOS on VW_TB_ORDENS_SERVICO.COD_ORDEM_SERVICO = VW_TB_PEDIDOS_COMPRAS_PRODUTOS.COD_ORDEM_SERVICO WHERE VW_TB_ORDENS_SERVICO.COD_CONTRATO = "&sqlTexto(rs("COD_CONTRATO"))&" AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.cod_pedido_compra = "&cod_pedido_compra&" ORDER BY VW_TB_ORDENS_SERVICO.ENTIDADE_ENSINO ASC")
					
					While not rs2.Eof
						
					subtotal = (subtotal+formatNumber(0&rs2("valor_total_item_solicitado")))/1
					
					qtd_itens_total = qtd_itens_total+formatNumber(0&rs2("quantidade_solicitada"))/1
					%>
                	<tr>
                      <td><div align="center" style="font-size:10px"><%=moduloPPVOSMascara(rs2("cod_ordem_servico"),"",true)%></div></td>
                      <td><div align="center" style="font-size:10px">
                        <div align="left"><%=rs2("nome_formando")%></div>
                      </div></td>
                      <td><div align="center" style="font-size:9px"><%=rs2("nome_produto")%></div></td>
                      <td><div align="center" style="font-size:10px"><%=rs2("quantidade_solicitada")%></div></td>
                      <td><div align="right" style="font-size:10px"><%=formatNumber(rs2("valor_custo_solicitado"),2)%></div></td>
                      <td colspan="2" style="font-size:10px"><div align="right"><%=formatNumber(rs2("valor_total_item_solicitado"),2)%></div></td>
				 	</tr>
                  	<%
				  	rs2.MoveNext()
				  	Wend
				  
				  	rs2.close
					Set rs2=nothing
					
					total = (total+subtotal)/1
					
				  %>
               
			<tr>
            	<th colspan="8" style="border-top:1px solid #000; font-size:12px;"><div align="right"><%=formatNumber(subtotal,2)%></div></th>
              	</tr>
</table>
                            <%
								rs.MoveNext()
								Wend
							
							rs.close
							Set rs=nothing
							%>
                            <br>


                            <table width="100%" cellpadding="0" cellspacing="0">
                            <tr>
            	<th style="font-size:15px;"><div align="left">Qtd. Itens
            	  </div>
          	  </div></th>
            	<th style="font-size:15px;"><div align="right"><%=formatNumber(qtd_itens_total,0)%></div></th>
              	            </tr>
                            <tr>
            	<th style="border-top:2px solid #000; font-size:15px;"><div align="left">Valor Total
            	  </div>
          	  </div></th>
            	<th style="border-top:2px solid #000; font-size:15px;"><div align="right"><%=formatNumber(total,2)%></div></th>
              	            </tr>
                            </table>

							 <table width="350" border="0" align="center" cellpadding="0" cellspacing="0">
							   <tr>
								 <td style="border-bottom:solid 1px #000000"><p> </p>
								   <p> </p></td>
							   </tr>
							   <tr>
								 <td><div align="center"><span style="font-size:11px"><strong><%=Ucase(nomeUsuario(31,true))%></strong></span></div></td>
							   </tr>
							 </table>
						<br />	
					<div style="padding-top:30px">		 
					
					</div>
							<p></td>
							</tr>
					
					
						</table></td>
					  </tr>
					</table>
					</td>
					</tr>
					</table>
<%FecharConexao()%>
</body>
</html>



