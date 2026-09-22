<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<!-- #include file="files/includes/functions_local.asp" -->
<%
sub segundaViaPeidoComrpa(cod_pedido_compra, conexao_aberta)

	if not conexao_aberta Then
		AbrirConexao()
	end if

	objConexao.execute("DELETE FROM TB_PEDIDOS_COMPRAS_PRODUTOS_TEMPORARIA WHERE cod_usuario = "&sqlNumero(session("cod_usuario")))

	objConexao.Execute("INSERT INTO TB_PEDIDOS_COMPRAS_PRODUTOS_TEMPORARIA (cod_pedido_compra_produto, cod_usuario ) SELECT cod_pedido_compra_produto, "&session("cod_usuario")&" FROM TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra = "&sqlNumero(cod_pedido_compra))

	if not conexao_aberta Then
		FecharConexao()
	end if	

end sub


Dim cod_pedido_compra, numero_pedido_compra, cod_fornecedor, nome_fantasia, aprovado_cod_usuario, data_aprovado
Dim qtd_itens_total, c

cod_pedido_compra = ""&request.queryString("c")


AbrirConexao()
	if cod_pedido_compra = "" then 
		'Busca o primeiro fornecedor
		Set rs=objConexao.Execute("SELECT * FROM VW_TB_PEDIDOS_COMPRAS_PRODUTOS_TEMPORARIA WHERE cod_usuario_temporario = "&Session("cod_usuario")&" LIMIT 1")
				
				If not rs.Eof Then
					cod_pedido_compra = rs("cod_pedido_compra")
					nome_fantasia = Ucase(rs("nome_fantasia"))
				End If
		
		rs.Close
		Set rs=Nothing
	else
		call segundaViaPeidoComrpa(cod_pedido_compra, true)
	end if
	
	Set rs=objConexao.Execute("SELECT * FROM TB_PEDIDOS_COMPRAS WHERE cod_pedido_compra = "&cod_pedido_compra)
		
		If not rs.Eof Then
			
			numero_pedido_compra=rs("numero_pedido_compra")
			cod_fornecedor=rs("cod_fornecedor")
			aprovado_cod_usuario=rs("aprovado_cod_usuario")
			data_aprovado=rs("data_aprovado")
			contato=rs("contato")
			pedido_manual = "S"

			Set rs2=objConexao.Execute("SELECT nome_fantasia FROM TB_FORNECEDORES WHERE cod_fornecedor = '"&cod_fornecedor&"'")
				
				If not rs2.Eof Then
					nome_fantasia = rs2("nome_fantasia")
				End If
			
			rs2.Close
			Set rs2=Nothing

		End If

	rs.Close
	Set rs=Nothing

'Ser� fechado no final da tela a conex�o


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
							<td>
								<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0">
						  			<tr>
										<td valign="top">
											<h2 style="color:#000000; border-bottom:solid #000 3px; padding-bottom:2px">
												Recibo de Entrega de Produtos - N° <%=numero_pedido_compra%> 
											</h2>
							  				<table width="100%" border="0" cellspacing="0" cellpadding="0">
												<tr>
								  					<td width="82%">
								  						<table width="100%" border="0" cellspacing="0" cellpadding="2">
															<tr>
																<td width="142"><span style="font-size:13px">Fornecedor</span></td>
																<td width="623" style="font-size:13px"><%=Ucase(nome_fantasia)%></td>
															</tr>
															<tr>
									  							<td style="font-size:13px">Data da Entrega</td>
									  							<td style="font-size:13px"><%=ArrumaData(date())%></td>
															</tr>
								  						</table>
								  					</td>
								  					<td width="18%">
								  						<div align="center">
								  							<img src="images/layout/logomarca_foto_reportagens.jpg" width="92" height="93" />
								  						</div>
								  					</td>
												</tr>
							  				</table>
							  			</td>
									</tr>
						  			<tr>
										<td> 
                            			<%
										if  moduloPPVpedidoManual(cod_pedido_compra, "", true) Then
										set rs = objConexao.execute("SELECT	SUM(quantidade_solicitada), SUM(quantidade_solicitada*valor_custo_solicitado) FROM VW_TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra = "&sqlNumero(cod_pedido_compra))
											if not rs.Eof Then
												qtd_itens_total = rs(0)
												total = rs(1)
											end if
										rs.close
										set rs = nothing

										set rs = objConexao.execute("SELECT *, quantidade_solicitada * valor_custo_solicitado as valor_custo_total FROM VW_TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra = "&sqlNumero(cod_pedido_compra))
											if not rs.Eof Then
												%>
													<table border="0" cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<th><center>Produto</center></th>
															<th><center>Quantidade</center></th>
															<th><center>Valor Un.</center></th>
															<th><center>Valor Total</center></th>
															<th><center>Observação</center></th>
														</tr>
													
												<%
												While not rs.Eof
													%>
													<tr>
														<td><%=rs("nome_produto")%></td>
														<td align="center"><%=rs("quantidade_solicitada")%></td>
														<td align="right">R$ <%=ArrumaNumero(rs("valor_custo_solicitado"),0,2)%></td>
														<td align="right">R$ <%=ArrumaNumero(rs("valor_custo_total"),0,2)%></td>
														<td align="right"><%=rs("obs")%></td>
													</tr>
													<%
												rs.MoveNext
												Wend
												%></table><%
											end if 
										rs.close
										set rs = nothing
										%>
										<br /><br />
									<br />
										<table width="100%" cellpadding="0" cellspacing="0">
			                            <tr>
											<th style="font-size:15px;">
												<div align="left">Qtd. Itens</div>
			          	  					</th>
			            					<th style="font-size:15px;">
			            						<div align="right"><%=ArrumaNumero(qtd_itens_total,0,0)%></div>
			            					</th>
			              	            </tr>
			                            <tr>
			            					<th style="border-top:2px solid #000; font-size:15px;">
			            						<div align="left">Valor Total</div>
			          	  					</th>
			            					<th style="border-top:2px solid #000; font-size:15px;">
			            						<div align="right"><%=ArrumaNumero(total,0,2)%></div>
			            					</th>
			              	            </tr>
									</table>

									<table width="350" border="0" align="center" cellpadding="0" cellspacing="0">
										<tr>
											<td style="border-bottom:solid 1px #000000"><p></p><p></p></td>
										</tr>
										<tr>
											<td>
												<div align="center">
													<span style="font-size:11px"><strong><%=nomeUsuario(Session("cod_usuario"),true)%></strong></span>
												</div>
											</td>
										</tr>
									</table>
<div style="padding-top:30px">
										
									</div>
										<%
									else

										'Agrupa por institui��o de ensino
										Set rs=objConexao.Execute("SELECT VW_TB_PEDIDOS_COMPRAS_PRODUTOS_TEMPORARIA.*, VW_TB_ORDENS_SERVICO.NOME_FORMANDO, VW_TB_ORDENS_SERVICO.COD_CONTRATO, VW_TB_ORDENS_SERVICO.ENTIDADE_ENSINO, VW_TB_ORDENS_SERVICO.cod_formando FROM VW_TB_ORDENS_SERVICO INNER JOIN VW_TB_PEDIDOS_COMPRAS_PRODUTOS_TEMPORARIA ON VW_TB_ORDENS_SERVICO.cod_ordem_servico = VW_TB_PEDIDOS_COMPRAS_PRODUTOS_TEMPORARIA.cod_ordem_servico WHERE cod_usuario_temporario = "&Session("cod_usuario")&" GROUP BY VW_TB_ORDENS_SERVICO.COD_CONTRATO ORDER BY VW_TB_ORDENS_SERVICO.ENTIDADE_ENSINO ASC")
										While not rs.Eof%>
											<table border="0" cellpadding="0" cellspacing="0" width="100%">
												<tr>
				  									<th colspan="7" style="padding-top:5px;">
				  										<div align="left" style="font-size:12px">
				  											<%=Ucase(rs("ENTIDADE_ENSINO"))%>
				  										</div>
				  									</th>
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
											Set rs2=objConexao.Execute("SELECT VW_TB_PEDIDOS_COMPRAS_PRODUTOS_TEMPORARIA.*, VW_TB_ORDENS_SERVICO.NOME_FORMANDO, VW_TB_ORDENS_SERVICO.COD_CONTRATO, VW_TB_ORDENS_SERVICO.ENTIDADE_ENSINO, VW_TB_ORDENS_SERVICO.cod_formando FROM VW_TB_ORDENS_SERVICO INNER JOIN VW_TB_PEDIDOS_COMPRAS_PRODUTOS_TEMPORARIA ON VW_TB_ORDENS_SERVICO.cod_ordem_servico = VW_TB_PEDIDOS_COMPRAS_PRODUTOS_TEMPORARIA.cod_ordem_servico WHERE cod_usuario_temporario = "&Session("cod_usuario")&" and VW_TB_ORDENS_SERVICO.COD_CONTRATO = "&sqlTexto(rs("COD_CONTRATO"))&" ORDER BY VW_TB_ORDENS_SERVICO.NOME_FORMANDO ASC")
												While not rs2.Eof
												
												subtotal = (subtotal+formatNumber(0&rs2("valor_total_item_solicitado")))/1
												qtd_itens_total = qtd_itens_total+formatNumber(0&rs2("quantidade_entregue"))/1
											%>
                        						<tr>
                          							<td><div align="center" style="font-size:10px"><%=rs2("numero_pedido_compra")%></div></td>
                          							<td><div align="left" style="font-size:10px"><%=rs2("nome_formando")%></div></td>
                          							<td><div align="center" style="font-size:10px"><%=rs2("nome_produto")%></div></td>
                          							<td><div align="center" style="font-size:10px"><%=rs2("quantidade_entregue")%></div></td>
                          							<td><div align="right" style="font-size:10px"><%=formatNumber(rs2("valor_custo_solicitado"),2)%></div></td>
                          							<td colspan="2"><div align="right" style="font-size:10px"><%=formatNumber(rs2("valor_total_item_solicitado"),2)%></div></td>
                        						</tr>
                  							<%
				  								rs2.MoveNext()
				  								Wend
				  							rs2.close
											Set rs2=nothing
											total = (total+subtotal)/1
				  							%>
               
											<tr>
            									<th colspan="7" style="border-top:1px solid #000; font-size:12px;"><div align="right"><%=formatNumber(subtotal,2)%></div></th>
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
									<th><div align="left"  style="font-size:15px;">Qtd. Itens</div></th>
									<th><div align="right" style="font-size:15px;"><%=formatNumber(qtd_itens_total,0)%></div></th>
								</tr>
								<tr>
									<th style="border-top:2px solid #000; font-size:15px;"><div align="left">Valor Total</div></th>
									<th style="border-top:2px solid #000; font-size:15px;"><div align="right"><%=formatNumber(total,2)%></div></th>
								</tr>
							</table>
							<table width="350" border="0" align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td style="border-bottom:solid 1px #000000"><p></p><p></p></td>
								</tr>
								<tr>
									<td>
										<div align="center">
											<span style="font-size:11px"><strong><%=Ucase(nomeUsuario(Session("cod_usuario"),true))%></strong></span>
										</div>
									</td>
								</tr>
							</table>

							<%end if%>
						
							<br />	
							<div style="padding-top:30px"></div>
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



