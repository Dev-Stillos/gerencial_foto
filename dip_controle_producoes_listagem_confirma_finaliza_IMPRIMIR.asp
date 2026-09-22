<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim rs, rs2, rs3, rs3_2, rs4, sql
Dim COD_VENDEDOR, NOME_VENDEDOR, COD_PRODUCAO, DESCRICAO
COD_VENDEDOR = Request.QueryString("COD_VENDEDOR")
COD_PRODUCAO = Request.QueryString("COD_PRODUCAO")

AbrirConexao()
	Set rs=objConexao.Execute("SELECT NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR = "&COD_VENDEDOR)
		
		If not rs.Eof Then
			NOME_VENDEDOR = rs("NOME_VENDEDOR")
		End If
		
	rs.Close
	Set rs=Nothing

	Set rs=objConexao.Execute("SELECT DESCRICAO FROM TB_PRODUCAO WHERE COD_PRODUCAO = "&COD_PRODUCAO)
		
		If not rs.Eof Then
			DESCRICAO = rs("DESCRICAO")
		End If
		
	rs.Close
	Set rs=Nothing
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

<!--INICÍCIO DO LAÇO POR VENDEDOR-->			
						   <%
						   '1)------------------------------
							AbrirConexao()
							%>	
					<table class="folha" width="100%" align="center" border="0" cellpadding="0" cellspacing="10">
					<tr>
					<td>
					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
					  <tr>
						<td><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0">
						  <tr>
							<td valign="top"><h2 style="color:#000000; border-bottom:solid #000 3px; padding-bottom:2px">Listagem de Produção</h2>
							  <br />
							  <table width="100%" border="0" cellspacing="0" cellpadding="0">
								<tr>
								  <td width="82%"><table width="100%" border="0" cellspacing="0" cellpadding="2">
									<tr>
									  <td width="142"><span class="textoMedio">Produção:</span><br /></td>
									  <td width="623"><strong class="textoMedio"><%=COD_PRODUCAO%> - <%=Ucase(DESCRICAO)%> </strong></td>
									</tr>
									<tr>
									  <td class="textoMedio">Data e hora: <strong class="textoMedio">
										
									  </strong></td>
									  <td class="textoMedio"><div align="left"><strong><%=Now()%></strong></div></td>
									</tr>
									<tr>
									  <td class="textoMedio">Repres. Comer.:</td>
									  <td class="textoMedio"><strong><%=NOME_VENDEDOR%></strong></td>
									</tr>
                                    <tr>
									  <td class="textoMedio">Gerado por:</td>
									  <td class="textoMedio"><strong><%=Ucase(Session("NOME"))%></strong></td>
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
					  <td width="27%"><div align="center"><strong>Nome</strong></div></td>
					  <td width="21%"><div align="center"><strong>Entidade de Ensino </strong></div></td>
					  <td width="7%"><div align="center"><strong>Comissão</strong></div></td>
					  <td width="10%"><div align="center"><strong>Brinde</strong></div></td>
					  <td width="24%"><div align="center"><strong>Renovação ou Devolução</strong></div></td>
					          </tr>
							  
							  <%
							  Set rs4 = objConexao.Execute("SELECT TB_PRODUCAO_LOTES_FORMANDOS_TEMPORARIA.*,VW_TB_PRODUCAO_LOTES_FORMANDOS.* FROM VW_TB_PRODUCAO_LOTES_FORMANDOS inner join TB_PRODUCAO_LOTES_FORMANDOS_TEMPORARIA on VW_TB_PRODUCAO_LOTES_FORMANDOS.cod_formando = TB_PRODUCAO_LOTES_FORMANDOS_TEMPORARIA.cod_formando where TB_PRODUCAO_LOTES_FORMANDOS_TEMPORARIA.cod_usuario = "&Session("cod_usuario")&" order by nome_formando")
							  

							  While not rs4.Eof
									
													
											COD_CONTRATO			= rs4("COD_CONTRATO")
											COD_FORMANDO 			= rs4("COD_FORMANDO")
											ALBUM 					= rs4("ALBUM")
											ENCADERNACOES_EXTRAS	= formatNumber(0&rs4("ENCADERNACOES_EXTRAS"),0)
											INICIAL 				= rs4("INICIAL")
											NOME_FORMANDO 			= rs4("NOME_FORMANDO")
												'Serão validados
												COMISSAO = rs4("COMISSAO")
												BRINDE = rs4("BRINDE")
												PRODUTO_AGREGADO 	= rs4("PRODUTO_AGREGADO")					

													
													'3
													If COMISSAO = 1 Then
														COMISSAO = "Sim"
													Else
														COMISSAO = "Não"								
													End If
													
												%>
														<tr>
								<td><div align="center" class="textoMenor textoMenor"><%=digito(COD_FORMANDO)%></div></td>
								<td><div align="center" class="textoMenor textoMenor"><%=ALBUM%><%If ENCADERNACOES_EXTRAS = 1 Then%> 
								    <strong>D</strong> 
							      <%End If%></div></td>
								<td style="font-size:10px;"><span class="textoMenor textoMenor"><%=NOME_FORMANDO%></span></td>
								<td style="font-size:10px;"><%=rs4("DESCRICAO")%></td>
								<td><div align="center"  style="font-size:10px;" class="textoMenor textoMenor"><%=COMISSAO%></div></td>
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
										<table  width="80" border="1" align="center" 	cellpadding="0" cellspacing="0" style="border-collapse:collapse; font-size:10px">
												
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
										<%Else%>
										<span class="style1">Brinde não encontrado, contate o administrador do sistema!</span></span>
										<span class="textoMenor textoMenor">
										<%End if	
											
										rsBrinde.Close()		
										Set rsBrinde = Nothing
								Else
													
									Response.Write("<div align='center'>Não</div>")
								
								End If
								
								%>			
										</span></td>
							    <td style="font-size:10px;"><div align="center">
								<%
								If int(rs4("RENOVAR_DEVOLVER")) = 1 Then
								'/Renovar
									 
										Response.Write("Renovação para mais "&rs4("PERIODO_RENOVAR")&" dias<br />(até "&ArrumaData(dateAdd("d",rs4("PERIODO_RENOVAR"),Date))&")")
								
								ElseIf int(rs4("RENOVAR_DEVOLVER")) = 2 Then
								'Devolver ficha
										Response.Write("Devolução de ficha*")
								End If
								%>
                                
                                </div></td>
						      </tr>
														
							<%
							  rs4.MoveNext()
							  Wend
							  
							  %>
							
							  <%
							  rs4.Close()
							  Set rs4 = Nothing
							  
					  
							  %>
							
							</table>
						
							
							 <p align="justify">* Devolução de Fichas: É obrigatório a devolução do álbum e demais itens do formando no estoque.</p>
							 <table width="350" border="0" align="center" cellpadding="0" cellspacing="0">
							   <tr>
								 <td style="border-bottom:solid 1px #000000"><p> </p>
								   <p> </p></td>
							   </tr>
							   <tr>
								 <td><div align="center"><span class="textoMedio"><%=NOME_VENDEDOR%></span></div></td>
							   </tr>
							 </table>
                             
                             <br />
						<p></td>
						  </tr>
					
					
						</table></td>
					  </tr>
					</table>
					</td>
					</tr>
					</table>
					
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



