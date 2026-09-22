<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim rs, rs2, rs3, rs3_2, rs4, rs4_1, sql
Dim COD_VENDEDOR_INICIAL, DATA_SAIDA_INICIAL, QTD_FORMANDOS_INICIAL, PRAZO_DIAS_CORRIDOS_INICIAL, LOTE
Dim NOME_VENDEDOR, COD_PRODUCAO, DESCRICAO, QTD_ALBUNS, QTD_PROD_AGREGADOS, PRODUTOS_AGREDADOS
Dim ENCADERNACOES_EXTRAS
QTD_ALBUNS = 0
QTD_PROD_AGREGADOS = 0

COD_PRODUCAO = Request.QueryString("COD_PRODUCAO")
LOTE = Request.QueryString("LOTE")

AbrirConexao()
'1) Verifica se realmente a produção está finalizada
Set rs = objConexao.Execute("SELECT DESCRICAO, PRODUTOS_AGREDADOS FROM TB_PRODUCAO WHERE COD_PRODUCAO="&COD_PRODUCAO&" AND PRODUCAO_FINALIZADA= 1 LIMIT 1")

	If not rs.Eof then
		DESCRICAO = rs("DESCRICAO")
		PRODUTOS_AGREDADOS = formatNumber(0&rs("PRODUTOS_AGREDADOS"),0)
	else
		Response.Redirect("dip_producoes.asp?msg=Produção número: "&COD_PRODUCAO&" não está finalizada!")
	End If
	
rs.Close()
Set rs = Nothing

'2) Verifica se � segunda via ou primeira via
Set rs = objConexao.Execute("SELECT * FROM TB_PRODUCAO_LOTES WHERE COD_PRODUCAO="&COD_PRODUCAO&" AND LOTE = "&LOTE&" LIMIT 1")

	If not rs.Eof then
	
		DATA_SAIDA_INICIAL = ""&rs("DATA_SAIDA_INICIAL")
		QTD_FORMANDOS_INICIAL = rs("QTD_FORMANDOS_INICIAL")
		PRAZO_DIAS_CORRIDOS_INICIAL = rs("PRAZO_DIAS_CORRIDOS_INICIAL")
		COD_VENDEDOR_INICIAL = formatNumber(0&rs("COD_VENDEDOR_INICIAL"),0)

		
		
		if int(COD_VENDEDOR_INICIAL) = 0 Then
			'Pega da querystring
			COD_VENDEDOR_INICIAL = Request.QueryString("COD_VENDEDOR_INICIAL")
		End If
 		
	else
		Response.Redirect("dip_producoes.asp?msg=produção lote : "&LOTE&" não encontrado!")
	End If
	
rs.Close()
Set rs = Nothing

'///////////////////////////////////////////////////
'/Se não tiver data de sa�da significa que � para imprimir a 1via e dar saida para o vendedor

If DATA_SAIDA_INICIAL = "" Then
DATA_SAIDA_INICIAL = Now()
'1a via dever� fazer os updates e lan�ar a data de sa�da no A) lote, B) lotes_formandos e C) movimentacoes
		
		'A)----------------
		objConexao.Execute("UPDATE TB_PRODUCAO_LOTES SET DATA_SAIDA_INICIAL = NOW( ) WHERE COD_PRODUCAO = '"&COD_PRODUCAO&"' AND LOTE = '"&LOTE&"' AND DATA_SAIDA_INICIAL IS NULL LIMIT 1")
		
		
		Dim DATA_PRAZO
		
		Set rs=objConexao.Execute("SELECT cod_formando FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = '"&COD_PRODUCAO&"' AND LOTE = '"&LOTE&"'")
		
		DATA_PRAZO = DateAdd("d",PRAZO_DIAS_CORRIDOS_INICIAL,DATA_SAIDA_INICIAL)
		
			If not rs.Eof Then
				While not rs.Eof
					
					'B)----------------
					'Somente os que não tiverem data de sa�da inicial
					objConexao.Execute("UPDATE TB_PRODUCAO_LOTES_FORMANDOS SET COD_VENDEDOR = '"&COD_VENDEDOR_INICIAL&"',DATA_SAIDA = "&sqlData(DATA_SAIDA_INICIAL)&", DATA_PRAZO = "&sqlData(DATA_PRAZO)&" WHERE COD_FORMANDO = "&rs("COD_FORMANDO")&" and DATA_SAIDA IS NULL")
					
					'C)----------------
					'Somente os que não tiverem data de sa�da inicial e nao tiverem vendedor ainda
					objConexao.Execute("UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET COD_VENDEDOR = '"&COD_VENDEDOR_INICIAL&"',DATA_SAIDA = "&sqlData(DATA_SAIDA_INICIAL)&", DATA_PREVISTA_RETORNO = "&sqlData(DATA_PRAZO)&"  WHERE COD_FORMANDO ="&rs("COD_FORMANDO")&" and DATA_SAIDA IS NULL and COD_VENDEDOR = 0")
					
				rs.MoveNext()
				Wend

			End If		

		rs.Close
		Set rs=Nothing
			
		
End If

'//////////////////
'Nome do vendedor
Set rs2 = objConexao.Execute("SELECT NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR="&COD_VENDEDOR_INICIAL)
										
		If not rs2.Eof then
													
			NOME_VENDEDOR = rs2("NOME_VENDEDOR")			
									
		End If
		

rs2.Close
Set rs2 = Nothing
'//////////////////
%>



<!-- include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="print"></link>
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<style type="text/css">
<!--
.btnFinalizaProducao{	background-image:url(images/layout/btnFinalizaProducao.jpg);
	cursor:pointer;
	width:140px;
	height:22px;
	border:0px;
}
.folha {
   page-break-after: always;
}
.style1 {color:#CC0000; font-weight: bold;}
-->
</style>

<title>StillosWEB - Sistema corporativo - Versão 1.2</title>
</head>
<body style="background:#FFFFFF">

<!--INICIO DO LAÇO POR VENDEDOR-->			
						   <%
						   '1)------------------------------
							AbrirConexao()
							%>	
					<table class="folha" width="100%" align="center" border="0" cellpadding="0" cellspacing="10">
					<tr>
					<td>
					<table width="100%" border="0" align="center" cellpadding="10" cellspacing="0" bgcolor="#FFFFFF">
					  <tr>
						<td><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0">
						  <tr>
							<td valign="top"><h2 style="color:#000000; border-bottom:solid #000 3px; padding-bottom:2px">Liberação de Produtos por produção </h2>
							  <br />
							  <table width="100%" border="0" cellspacing="0" cellpadding="0">
								<tr>
								  <td width="82%"><table width="100%" border="0" cellspacing="0" cellpadding="2">
									<tr>
									  <td width="142"><span class="textoMedio">Produção:</span><br /></td>
									  <td width="623"><strong class="textoMedio"><%=COD_PRODUCAO%> - <%=Ucase(DESCRICAO)%>|Lote: <%=LOTE%></strong></td>
									</tr>
									<tr>
									  <td class="textoMedio">Data e hora: <strong class="textoMedio">
										
									  </strong></td>
									  <td class="textoMedio"><div align="left"><%=DATA_SAIDA_INICIAL%></div></td>
									</tr>
									<tr>
									  <td class="textoMedio">Repres. Comer.:</td>
									  <td class="textoMedio"><strong><%=NOME_VENDEDOR%></strong></td>
									</tr>
                                    <tr>
									  <td class="textoMedio">Prazo</td>
									  <td class="textoMedio"><%=PRAZO_DIAS_CORRIDOS_INICIAL%> dias corridos (até <%=DateAdd("d",PRAZO_DIAS_CORRIDOS_INICIAL,ArrumaData(Year(DATA_SAIDA_INICIAL)&"-"&Month(DATA_SAIDA_INICIAL)&"-"&Day(DATA_SAIDA_INICIAL)))%>)</td>
									</tr>
								  </table></td>
								  <td width="18%"><div align="center"><img src="images/layout/logomarca_foto_reportagens.jpg" width="92" height="93" /></div></td>
								</tr>
							  </table></td>
							</tr>
						  <tr>
							<td>          <table width="100%" border="1" align="center" 	cellpadding="1" cellspacing="0" style="border-collapse:collapse">
					<tr bgcolor="whitesmoke">
						   
								<td width="7%"><div align="center"><strong>Cód. Formando </strong></div></td>
					  <td width="4%"><div align="center"><strong>Álbum</strong></div></td>
					  <td width="4%"><div align="center"><strong>Álbum Digital</strong></div></td>
					  <td width="27%"><div align="center"><strong>Nome</strong></div></td>
					  <td width="26%"><div align="center"><strong>Entidade de Ensino </strong></div></td>
					  <td width="4%"><div align="center"><strong>Inicial</strong></div></td>
					  <td width="6%"><div align="center"><strong>Comissão</strong></div></td>
					  <td width="6%"><div align="center"><strong>Brinde</strong></div></td>
								<%If int(PRODUTOS_AGREDADOS) = 1 Then%>
                      <td width="13%"><div align="center"><strong>Produto Agregado </strong></div></td>
                                <%End If%>
							  </tr>
							  
							  <%

							  'Instru��es:
							  '1 Seleciona todas as fichas desta produção
							  '2 Dentro da ficha verifica nas Movimenta��es se está ficha � deste vendedor
							  '3 Caso for, verifica se tem brinde ou produto agregado
							  '4 Faz o update caso for 1� via
							  
							  '1
							  Set rs4 = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO="&COD_PRODUCAO&" and LOTE = "&LOTE&" ORDER BY CONVERT(ALBUM,SIGNED)")
							  
					
							  
							  While not rs4.Eof
									
									'2
									Set rs4_1 = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_FORMANDO="&rs4("COD_FORMANDO")&" AND COD_VENDEDOR="&COD_VENDEDOR_INICIAL&" LIMIT 1")
											
											If not rs4_1.Eof then
											'*Significa que � deste vendedor
													
											'Armazena vari�veis
											'Dim COD_FORMANDO, ALBUM, ALBUM_DIGITAL, INICIAL, NOME_FORMANDO, COMISSAO, BRINDE, PRODUTO_AGREGADO
											Dim DATA_SAIDA
											
											'Contador de �lbuns
											QTD_ALBUNS = QTD_ALBUNS + 1
											
											COD_CONTRATO			= rs4("COD_CONTRATO")
											COD_FORMANDO 			= rs4("COD_FORMANDO")
											ALBUM 					= rs4("ALBUM")
											ALBUM_DIGITAL			= rs4("ALBUM_DIGITAL")
											ENCADERNACOES_EXTRAS	= formatNumber(0&rs4("ENCADERNACOES_EXTRAS"),0)
											INICIAL 				= rs4("INICIAL")
											NOME_FORMANDO 			= rs4("NOME_FORMANDO")
											
												'Ser�o validados
												COMISSAO = rs4("COMISSAO")
												BRINDE = rs4("BRINDE")
												PRODUTO_AGREGADO 	= rs4("PRODUTO_AGREGADO")					

													
													'3
													If COMISSAO = 1 Then
														COMISSAO = "SIM"
													Else
														COMISSAO = "NÃO"								
													End If
													
												%>
														<tr>
								<td><div align="center" class="textoMenor textoMenor"><%=digito(COD_FORMANDO)%></div></td>
								<td><div align="center" class="textoMenor textoMenor"><%=ALBUM%><%If ENCADERNACOES_EXTRAS = 1 Then%> 
								    <strong>D</strong> 
							      <%End If%></div></td>
								<td><div align="center" class="textoMenor textoMenor">
									<% 
										' Se vier vazio ou 0, força exibir SIM
										If ALBUM_DIGITAL = "" Or ALBUM_DIGITAL = 0 Then
											Response.Write("SIM")
										Else
											' Se quiser, pode mostrar "SIM" ou "NÃO" dependendo do valor
											If ALBUM_DIGITAL = 1 Then
												Response.Write("SIM")
											Else
												Response.Write("NÃO")
											End If
										End If
									%>
								</div></td>

								<td><span class="textoMenor textoMenor"><%=NOME_FORMANDO%></span></td>
								<td><span class="textoMenor textoMenor">
								  <%
								Dim rsTurma
								Set rsTurma = objConexao.Execute("SELECT DESCRICAO FROM TB_ENTIDADES WHERE COD_CONTRATO="&sqlTexto(COD_CONTRATO))
									
									If not rsTurma.Eof Then
										Response.Write(rsTurma("DESCRICAO"))
									Else%>
								  <span class="style1">Falha na entidade de ensino não encontrada, informe ao depto.TI.</span></span>
								  <span class="textoMenor textoMenor">
									<%End If
								
								rsTurma.Close
								Set rsTurma = Nothing
								%>
								  </span></td>
								<td><div align="center" class="textoMenor textoMenor"><%=INICIAL%></div></td>
								<td><div align="center" class="textoMenor textoMenor"><%=COMISSAO%></div></td>
								<td style="padding:4px">
										<span class="textoMenor textoMenor">
										<%
								If BRINDE = 1 Then
											
										Dim rsBrinde
										
										sql = "SELECT TB_PRODUTOS.NOME_PRODUTO, Sum( TB_FICHAS_IDENTIFICACAO_BRINDES.QUANTIDADE ) AS QUANTIDADE"&_
																	" FROM TB_PRODUTOS"&_
																	" INNER JOIN TB_FICHAS_IDENTIFICACAO_BRINDES ON TB_PRODUTOS.COD_PRODUTO = TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO"&_
																	" WHERE TB_FICHAS_IDENTIFICACAO_BRINDES.COD_FORMANDO = "&COD_FORMANDO&_
																	" GROUP BY TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO ORDER BY TB_PRODUTOS.NOME_PRODUTO"
										
										
										Set rsBrinde = objConexao.Execute(sql)
					
										If not rsBrinde.Eof then%>
										</span>
										<table width="115" border="1" align="center" 	cellpadding="0" cellspacing="0" style="border-collapse:collapse">
												
											  <%While not rsBrinde.Eof%>
											  <tr>
												<td width="12%" class="textoMenor textoMenor"><div align="center"><%=rsBrinde("QUANTIDADE")%></div></td>
												<td width="81%" class="textoMenor textoMenor"><div align="center"><%=rsBrinde("NOME_PRODUTO")%></div></td>
											  </tr>
											  <%
											  rsBrinde.MoveNext()
											  Wend
											  %>
											</table>
										<span class="textoMenor textoMenor">
										<%Else
										if rs4("COD_PRODUCAO_desconto") <> "" then
											set rsDesconto = objConexao.Execute("SELECT * FROM TB_PRODUCAO_DESCONTOS WHERE COD_PRODUCAO_desconto ="&rs4("COD_PRODUCAO_desconto"))

											if not rsDesconto.Eof Then
												Response.write rsDesconto("DESCRICAO")
											end if

										else
										%>
											<span class="style1">Brinde não encontrado, contate o administrador do sistema!</span></span>
										<%
										end if
										%>
										


										<span class="textoMenor textoMenor">
										<%End if	
											
										rsBrinde.Close()		
										Set rsBrinde = Nothing
								Else
													
									Response.Write("<div align='center'>NÃO</div>")
								
								End If
								
								%>			
										</span></td>
                                        <%If int(PRODUTOS_AGREDADOS) = 1 Then%>
								<td><span class="textoMenor textoMenor">
								  <%
													
							
								If PRODUTO_AGREGADO = 1 Then
											
										Dim rsProd
										sql = "SELECT TB_PRODUTOS.NOME_PRODUTO, Sum( TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.QUANTIDADE ) AS QUANTIDADE"&_
																	" FROM TB_PRODUTOS"&_
																	" INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO"&_
																	" WHERE TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO = "&COD_FORMANDO&_
																	" GROUP BY TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO ORDER BY TB_PRODUTOS.NOME_PRODUTO"
																	'Response.Write(sql)
										Set rsProd = objConexao.Execute(sql)
					
										If not rsProd.Eof then%>
											</span>
								  <table width="115" border="1" align="center" 	cellpadding="0" cellspacing="0" style="border-collapse:collapse">
											  <%While not rsProd.Eof
											  'Contador de Produtos agregados
											  QTD_PROD_AGREGADOS = QTD_PROD_AGREGADOS + formatNumber(0&rsProd("QUANTIDADE"),0)
											  
											  %>
											  <tr>
												<td width="12%" class="textoMenor textoMenor"><div align="center"><%=rsProd("QUANTIDADE")%></div></td>
												<td width="81%" class="textoMenor textoMenor"><div align="center"><%=rsProd("NOME_PRODUTO")%></div></td>
											  </tr>
											  <%
											  rsProd.MoveNext()
											  Wend
											  %>
											</table>
										<span class="textoMenor textoMenor">
										<%Else%>
										<span class="style1">Produto agregado  não encontrado, contate o administrador do sistema!</span></span>
										<span class="textoMenor textoMenor">
										<%End if	
											
										rsProd.Close()		
										Set rsProd = Nothing
								Else
													
									Response.Write("<div align='center'>NÃO</div>")
								
								End If
								
								%>			
										</span></td>
                                        <%End If%>
							  </tr>
														
																				<%						
												
											
											End If
									
									
									rs4_1.Close()
									Set rs4_1 = Nothing
					
							  rs4.MoveNext()
							  Wend
							  
							  %>
							
							  <%
							  rs4.Close()
							  Set rs4 = Nothing
							  
					  
							  %>
							
							</table>
						
							<%If QTD_PROD_AGREGADOS > 0 then
							'Com produtos agregados
							%>
							 <p align="justify">EU <strong><%=Ucase(NOME_VENDEDOR)%></strong>, ESTOU RETIRANDO DA SEDE DA STILLO'S REPORTAGENS FOTOGRÁFICAS EM <strong><%=ArrumaData(DATA_SAIDA_INICIAL)%></strong> A QUANTIA DE <strong><%=QTD_ALBUNS%></strong> ÁLBUNS DIGITAIS <strong><%=QTD_PROD_AGREGADOS%></strong> PRODUTOS AGREGADOS, E ACIMA RELACIONADOS COM SUAS RESPECTIVAS FICHAS DE IDENTIFICAÇÃO E ESTOU CIENTE QUE DEVEREI DEVOLVER TODO MATERIAL PARA STILLO'S NO PRAZO MÁXIMO DE <%=PRAZO_DIAS_CORRIDOS_INICIAL%> DIAS (CORRIDOS) A CONTAR DA DATA DE HOJE.</p>
							 <%
							else
							'Normal somente �lbuns
							%>
							<p align="justify">
							EU <strong><%=Ucase(NOME_VENDEDOR)%></strong>, ESTOU RETIRANDO DA SEDE DA STILLO'S REPORTAGENS FOTOGRÁFICAS EM <strong><%=ArrumaData(formatDateTime(DATA_SAIDA_INICIAL))%></strong> A QUANTIA DE <strong><%=QTD_ALBUNS%></strong> ÁLBUNS DIGITAIS, E ACIMA RELACIONADOS COM SUAS RESPECTIVAS FICHAS DE IDENTIFICAÇÃO E ESTOU CIENTE QUE DEVEREI DEVOLVER TODO MATERIAL PARA STILLO'S NO PRAZO MÁXIMO DE <%=PRAZO_DIAS_CORRIDOS_INICIAL%> DIAS (CORRIDOS) A CONTAR DA DATA DE HOJE.		   </p>
							   <%End If%>
							 
							 <table width="350" border="0" align="center" cellpadding="0" cellspacing="0">
							   <tr>
								 <td style="border-bottom:solid 1px #000000"><p></p>
								   <p></p></td>
							   </tr>
							   <tr>
								 <td><div align="center"><span class="textoMedio"><%=NOME_VENDEDOR%></span></div></td>
							   </tr>
							 </table>
						<br />	
					<div style="padding-top:30px">		 
					<table width="300" border="0" align="right" cellpadding="0" cellspacing="0">
					  <tr>
						<td width="140"><div align="right"><strong>Separado por:</strong><br />
						</div></td>
						<td width="160" style="border-bottom:1px #000000 solid"></td>
					  </tr>
					  <tr>
						<td><div style="padding-top:10px">
						  <div align="right"><strong>Conferido Recepção por:</strong></div>
						</div></td>
						<td style="border-bottom:1px #000000 solid"></td>
					  </tr>
					</table>
					</div>
							<p></td>
							</tr>
					
					
						</table></td>
					  </tr>
					</table>
					</td>
					</tr>
					</table>
					<%
					QTD_PROD_AGREGADOS = 0
					QTD_ALBUNS = 0
					%>			
					</td>
					</tr>
					</table>	
					<!--FIM COUNTEÚDO LAÇO-->
		 		
		<%				  		
		FecharConexao()
		%>
<!--FIM DO LAÇO POR VENDEDOR-->
</body>
</html>



