<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<%
  Dim cod_contrato, qtd_por_contrato, quebra_linha
  Dim DATA_RETORNO, NOME_VENDEDOR
  
  cod_contrato = ""
  qtd_por_contrato = 0
  quebra_linha = chr(13)
  AbrirConexao()
  
  If 1=10 Then
  	Response.AddHeader "Refresh", "15" 
	%>
    	<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
</head>
<body>
    
	  Set rs=objConexao.Execute("SELECT TB_FICHAS_IDENTIFICACAO.COD_FORMANDO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, TB_FICHAS_IDENTIFICACAO.COD_CONTRATO,TB_FICHAS_IDENTIFICACAO.ANO_CONCLUSAO,TB_FICHAS_IDENTIFICACAO.SEMESTRE_CONCLUSAO,TB_FICHAS_IDENTIFICACAO.ALBUM, TB_FICHAS_IDENTIFICACAO.INICIAL,  TB_ENTIDADES.DESCRICAO FROM TB_ENTIDADES INNER JOIN  TB_FICHAS_IDENTIFICACAO ON TB_ENTIDADES.COD_CONTRATO = TB_FICHAS_IDENTIFICACAO.COD_CONTRATO WHERE TB_FICHAS_IDENTIFICACAO.VENDIDO = '2' AND TB_ENTIDADES.ANO IN(2009,2010) AND TB_FICHAS_IDENTIFICACAO.COD_FORMANDO NOT IN(SELECT COD_FORMANDO FROM TEMP) ORDER BY TB_ENTIDADES.ANO, TB_FICHAS_IDENTIFICACAO.COD_CONTRATO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO LIMIT 500")
	 
			While not rs.Eof
												
					If Cstr(""&cod_contrato) <> Cstr(""&rs("cod_contrato")) Then
						If qtd_por_contrato > 0 Then
							'Response.Write("Total deste contrato: "&qtd_por_contrato&quebra_linha&quebra_linha)
						End If
						
						cod_contrato = rs("cod_contrato")
						'Response.Write("<strong>"&rs("descricao")&"</strong>"&quebra_linha)
						qtd_por_contrato = 0
					End If
					
					DATA_RETORNO = ""
					NOME_VENDEDOR = ""
					
					Set rs2=objConexao.Execute("SELECT TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_RETORNO, TB_VENDEDORES.NOME_VENDEDOR FROM  TB_VENDEDORES INNER JOIN TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ON (TB_VENDEDORES.COD_VENDEDOR = TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR) WHERE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_FORMANDO = "&rs("COD_FORMANDO")&" AND TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.TIPO = 0 ORDER BY TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_MOVIMENTACAO DESC LIMIT 1")
						
						If not rs2.Eof Then
							DATA_RETORNO = ArrumaData(rs2("DATA_RETORNO"))
							NOME_VENDEDOR = rs2("NOME_VENDEDOR")
						End If
						
					rs2.Close
					Set rs2=Nothing
					
					If DATA_RETORNO <> "" Then
						'Response.Write(rs("NOME_FORMANDO")&chr(9)&rs("ALBUM")&chr(9)&rs("INICIAL")&chr(9)&DATA_RETORNO&chr(9)&NOME_VENDEDOR&quebra_linha)
						qtd_por_contrato = qtd_por_contrato+1
					End If
					
					objConexao.Execute("INSERT INTO TEMP VALUES("&sqlTexto(rs("COD_FORMANDO"))&","&sqlTexto(rs("NOME_FORMANDO"))&","&sqlTexto(rs("ALBUM"))&","&sqlTexto(rs("INICIAL"))&","&sqlTexto(rs("COD_CONTRATO"))&","&sqlTexto(rs("DESCRICAO"))&","&sqlTexto(NOME_VENDEDOR)&","&sqlData(DATA_RETORNO)&")")
				
			rs.MoveNext()
			Wend
	  
	  rs.Close  
	  Set rs=Nothing
    </body>
</html>	
    <%
  Else
  
  	  Response.ContentType="application/xls"
  
	  Set rs=objConexao.Execute("SELECT * FROM TEMP")
		 
				While not rs.Eof
													
						If Cstr(""&cod_contrato) <> Cstr(""&rs("cod_contrato")) Then
							If qtd_por_contrato > 0 Then
								Response.Write("Total deste contrato: "&qtd_por_contrato&quebra_linha&quebra_linha)
							End If
							
							cod_contrato = rs("cod_contrato")
							Response.Write(Ucase(rs("descricao"))&quebra_linha)
							qtd_por_contrato = 0
						End If
					
						DATA_RETORNO = Cstr(""&ArrumaData(rs("DATA_RETORNO")))
						NOME_VENDEDOR = Ucase(""&rs("NOME_VENDEDOR"))
						
						If DATA_RETORNO <> "" Then
							Response.Write(Ucase(rs("NOME_FORMANDO"))&chr(9)&rs("COD_FORMANDO")&chr(9)&rs("ALBUM")&chr(9)&rs("INICIAL")&chr(9)&DATA_RETORNO&chr(9)&NOME_VENDEDOR&quebra_linha)
							qtd_por_contrato = qtd_por_contrato+1
						End If
						
					
					
				rs.MoveNext()
				Wend
		  
		  rs.Close  
		  Set rs=Nothing
	  
  End If
  
  FecharConexao()
  %>
