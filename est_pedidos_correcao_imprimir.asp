<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->

<!-- #include file="includes/classe_modulo_ppv.asp" -->
<!-- #include file="files/includes/functions_local.asp" -->
<%
Dim cod_pedido_compra, numero_pedido_compra, cod_fornecedor, nome_fantasia, aprovado_cod_usuario, data_aprovado, obs
Dim qtd_itens_total, corrigido, aux_text

cod_pedido_compra = Request.QueryString("c")

AbrirConexao()	

	Set rs=objConexao.Execute("SELECT * FROM tb_est_pedidos WHERE cod_pedido = "&cod_pedido_compra)

	If not rs.Eof Then
		corrigido=rs("corrigido")
		numero_pedido_compra=rs("cod_pedido")
		cod_fornecedor=rs("cod_fornecedor")
		cod_usuario=rs("cod_usuario")
		data_pedido=rs("data_pedido")
		data_previsao=rs("data_previsao")
		data_entrega=rs("data_entrega")

		obs = rs("obs_correcao")
		pedido_manual = "S"

		Set rs2=objConexao.Execute("SELECT nome_fantasia FROM TB_FORNECEDORES WHERE cod_fornecedor = '"&cod_fornecedor&"'")

		If not rs2.Eof Then
			nome_fantasia = rs2("nome_fantasia")
		End If

		if corrigido = "S" then
			aux_text = "_corrigido"
		else
			corrigido = "N"
			aux_text = ""
		end if

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
						<td>
							<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0">
								<tr>
									<td valign="top"><h2 style="color:#000000; border-bottom:solid #000 3px; padding-bottom:2px">
									<% if corrigido <> "S" then%>
										Pedido de Compra <%else%>Recibo <%end if%>- Nº <%=numero_pedido_compra%> / EST</h2>
										<table width="100%" border="0" cellspacing="0" cellpadding="0">
											<tr>
												<td width="82%">
													<table width="100%" border="0" cellspacing="0" cellpadding="2">
														<tr>
															<td width="142"><span style="font-size:13px">Fornecedor</span></td>
															<td width="623" style="font-size:13px"><%=Ucase(nome_fantasia)%></td>
														</tr>
														<tr>
															<td style="font-size:13px">Data do Pedido</td>
															<td style="font-size:13px"><%=ArrumaData(data_pedido)%></td>
														</tr>
															<tr>
															<td style="font-size:13px">Data entrega</td>
															<td style="font-size:13px"><%=ArrumaData(data_entrega)%></td>
														</tr>
														<tr>
															<td style="font-size:13px">Observação</td>
															<td style="font-size:13px"><%=obs%></td>
														</tr>
													</table>
												</td>
												<td width="18%">
													<div align="center"><img src="images/layout/logomarca_foto_reportagens.jpg" width="92" height="93" /></div>
												</td>
											</tr>
										</table>
							  		</td>
								</tr>
						  		<tr>
									<td>
                            		<%
                            		'Seção de produtos dos pedidos de  compra
                            			dim a
                            			if corrigido = "S" Then
                            				a = "_corrigida"
                            			end if

										set rs = objConexao.execute("SELECT	SUM(Quantidade"&a&"), SUM(valor_total"&aux_text&") FROM VW_EST_ITENS_PEDIDOS WHERE cod_pedido = "&sqlNumero(cod_pedido_compra))
											if not rs.Eof Then
												qtd_itens_total = rs(0)
												total = rs(1)
											end if
										rs.close
										set rs = nothing

										set rs = objConexao.execute("SELECT nome_produto, Quantidade_corrigida as Quantidade_corrigido, valor_corrigido, valor_total_corrigido, valor_total, valor, Quantidade FROM VW_EST_ITENS_PEDIDOS WHERE cod_pedido = "&sqlNumero(cod_pedido_compra))
											if not rs.Eof Then
												%>
													<table border="0" cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<th><center>Produto</center></th>
															<th><center>Quantidade</center></th>
															<th><center>Valor Un.</center></th>
															<th><center>Valor Total</center></th>
														</tr>
													
												<%
												While not rs.Eof
													%>
													<tr>
														<td><%=rs("nome_produto")%></td>
														<td align="center"><%=rs("Quantidade"&aux_text)%></td>
														<td align="right">R$ <%=ArrumaNumero(rs("valor"&aux_text),0,2)%></td>
														<td align="right">R$ <%=ArrumaNumero(rs("valor_total"&aux_text),0,2)%></td>
														<td align="right"><%'=rs("obs")%></td>
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
													<%if Request.QueryString("d") <> "" then %>
													<span style="font-size:11px"><strong><%=Ucase(nome_fantasia)%></strong></span>
													<% else %>
													<span style="font-size:11px"><strong><%=Ucase(nomeUsuario(Session("cod_usuario"),true))%></strong></span>
													<% end if%>
												</div>
											</td>
										</tr>
									</table>
									<div style="padding-top:30px">
										<table width="400" border="0" align="right" cellpadding="0" cellspacing="0">
								  			<tr>
												<td width="140">
													<div align="right" style="font-size:11px;">
														<strong>Solicitado por:</strong><br />
													</div>
												</td>
												<td width="160" style="border-bottom:1px #000000 solid;font-size:11px;">
													<%=nomeUsuario(Session("cod_usuario"),true)%>
												</td>
								  			</tr>
										</table>
									</div>
								</td>
							</tr>
						</table>
					</td>
	  			</tr>
	</table>
	</td>
	</tr>
	</table>
<%FecharConexao()%>
</body>
</html>