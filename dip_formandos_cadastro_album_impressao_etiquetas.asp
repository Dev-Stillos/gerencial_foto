<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim cod_formando, redirecionar_volta, imprimir_agora
Dim NOME_FORMANDO, NOME_CURSO, ENCADERNACOES_EXTRAS, SEMESTRE_CONCLUSAO, ANO_CONCLUSAO, ALBUM, COD_CONTRATO
cod_formando = Request.QueryString("cod_formando")
redirecionar_volta = Request.QueryString("redit")
imprimir_agora	   = formatNumber(0&Request.QueryString("ia"),0)


If imprimir_agora = 1 Then
	redirecionar_volta = "dip_busca_detalhamento_formando.asp?cod_formando="&cod_formando&"&msg=Etiqueta Agendada com sucesso!"
Else	
	redirecionar_volta = "dip_formandos.asp?msg=Etiqueta(s) agendada(s) com sucesso!"
End If


Dim rs, rs2
AbrirConexao()
Set rs = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO="&cod_formando)

		If not rs.Eof Then
			NOME_FORMANDO			= rs("NOME_FORMANDO")
			NOME_CURSO				= Ucase(Left(rs("NOME_CURSO"),16)) 
			SEMESTRE_CONCLUSAO		= rs("SEMESTRE_CONCLUSAO") 
			ANO_CONCLUSAO			= rs("ANO_CONCLUSAO") 
			ALBUM					= rs("ALBUM") 
			COD_CONTRATO			= rs("COD_CONTRATO")
			ENCADERNACOES_EXTRAS 	= formatNumber(0&rs("ENCADERNACOES_EXTRAS"),0)
		End If

rs.Close
Set rs = Nothing

Dim HABILITA_IMPRESSAO_ALBUM, HABILITA_IMPRESSAO_ALBUM_DUPLO
HABILITA_IMPRESSAO_ALBUM		= 2
HABILITA_IMPRESSAO_ALBUM_DUPLO	= 2

Set rs2 = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_FORMANDO ="&cod_formando)
	
	If not rs2.Eof Then
			
			While not rs2.Eof
					
					Dim TIPO
					TIPO = formatNumber(0&rs2("TIPO"),0)
					
					If TIPO = 0 Then
					
						HABILITA_IMPRESSAO_ALBUM	= 1
						
						
						If ENCADERNACOES_EXTRAS = 1 Then
						
							HABILITA_IMPRESSAO_ALBUM_DUPLO	= 1
						
						End If
					
					End If
					
					
			rs2.MoveNext()
			Wend
	
	
	Else
			HABILITA_IMPRESSAO_ALBUM			= 2
			HABILITA_IMPRESSAO_ALBUM_DUPLO		= 2
			
			
	End if
	
	
rs2.Close
Set rs2 = Nothing

FecharConexao()




If Request.Form("Grava3") <> Empty Then

	AbrirConexao()


	Dim ETIQUETA_ALBUM, ETIQUETA_ALBUM_DUPLO, ETIQUETA_FICHA
	
	ETIQUETA_ALBUM			= formatNumber(0&Request.Form("ETIQUETA_ALBUM"),0)
	ETIQUETA_ALBUM_DUPLO	= formatNumber(0&Request.Form("ETIQUETA_ALBUM_DUPLO"),0)
	ETIQUETA_FICHA			= formatNumber(0&Request.Form("ETIQUETA_FICHA"),0)
	IMPRESSAO				= formatNumber(0&Request.Form("IMPRESSAO"),0)
	
	If ETIQUETA_ALBUM <> 0 Then
	
		If Len (ALBUM) = 1 Then
			ALBUM = "000"&ALBUM 
		ElseIf Len(ALBUM) = 2 Then
			ALBUM = "00"&ALBUM 
		ElseIf Len(ALBUM) = 3 Then
			ALBUM = "0"&ALBUM 
		End If
	
		NOME_FORMANDO	= UCASE(Left(NOME_FORMANDO,23))
		
		
		
		Dim rsTurmaForm
				
		Set rsTurmaForm = objConexao.Execute("SELECT DESCRICAO FROM TB_ENTIDADES WHERE COD_CONTRATO='"&COD_CONTRATO&"'")
						
						If not rsTurmaForm.Eof Then
						'Caso encontre coloca o nome da faculdade correspondente.
							TURMA = rsTurmaForm("DESCRICAO")
							'Correção do nome
							TURMA = Replace(TURMA,COD_CONTRATO&"    -   ","")
							TURMA = Replace(TURMA,COD_CONTRATO&"    -  ","")
							TURMA = Replace(TURMA,COD_CONTRATO&"   -   ","")
							TURMA = Replace(TURMA,COD_CONTRATO&"   - ","")
							TURMA = Replace(TURMA,COD_CONTRATO&" - ","")
							TURMA = Replace(TURMA,COD_CONTRATO&"- ","")
							TURMA = Replace(TURMA,COD_CONTRATO&" -","")
							TURMA = Replace(TURMA,COD_CONTRATO&"-","")
							TURMA = Replace(TURMA,COD_CONTRATO&" / ","")
							TURMA = Replace(TURMA,COD_CONTRATO,"")
							
							'2)Limpa o ano, semestre
							For i=2000 to 2015
							
							TURMA = Replace(TURMA,"-  /  J/"&i,"")
							TURMA = Replace(TURMA,"- /  J/"&i,"")
							TURMA = Replace(TURMA,"-  / J/"&i,"")
							TURMA = Replace(TURMA,"-/ J/"&i,"")
							TURMA = Replace(TURMA,"-/J/"&i,"")
							TURMA = Replace(TURMA,"-J/"&i,"")
							TURMA = Replace(TURMA,"J - /"&i,"")
							TURMA = Replace(TURMA,"/ J - "&i,"")
							TURMA = Replace(TURMA,"D/"&i,"")
							TURMA = Replace(TURMA,"J - D "&i,"")
							TURMA = Replace(TURMA,"J- /"&i,"")
							TURMA = Replace(TURMA,"-/ J"&i,"")
							TURMA = Replace(TURMA," - / J- "&i,"")
							TURMA = Replace(TURMA,"-  /  "&i,"")
							TURMA = Replace(TURMA,"- /  "&i,"")
							TURMA = Replace(TURMA,"J-/"&i,"")
							TURMA = Replace(TURMA,"-/"&i,"")
							TURMA = Replace(TURMA,"-/"&i,"")
							TURMA = Replace(TURMA,"-"&i,"")
							TURMA = Replace(TURMA,"J/"&i,"")
							TURMA = Replace(TURMA,"J"&i,"")
							TURMA = Replace(TURMA," J ","")
							TURMA = Replace(TURMA," D ","")
							TURMA = Replace(TURMA,i,"")
							
							Next
							
							TURMA = Replace(TURMA," - "," ")
							TURMA = Replace(TURMA,"- "," ")
							TURMA = Replace(TURMA,"-"," ")
							TURMA = Replace(TURMA,"/"," ")
							TURMA = Trim(TURMA)
							
							TURMA = Left(TURMA,23)
						End If
				
				rsTurmaForm.Close
				Set rsTurmaForm = Nothing
																		
		
		'0 por que é álbum
		LINHA1 = chr(34)&COD_FORMANDO&"0"&chr(34)
		LINHA2 = chr(34)&NOME_FORMANDO&chr(34)
		LINHA3 = chr(34)& Left(TURMA,16)&" - "&COD_CONTRATO&chr(34)
		LINHA4 = chr(34)&NOME_CURSO&" - "&SEMESTRE_CONCLUSAO&"/"&Right(ANO_CONCLUSAO,2)&chr(34)
		LINHA5 = chr(34)&"ALBUM:"&ALBUM&" - "&ArrumaData(Date())&chr(34)
		LINHA6 = chr(34)&"ALBUM"&chr(34)
											
		sql="INSERT INTO TB_FICHAS_IDENTIFICACAO_ETIQUETAS_TEMPORARIA ("&_
			"COD_USUARIO ,"&_
			"COD_FORMANDO ,"&_
			"LINHA1 ,"&_
			"LINHA2 ,"&_
			"LINHA3 ,"&_
			"LINHA4 ,"&_
			"LINHA5 ,"&_
			"LINHA6 ,"&_
			"AGENDADA "&_
			")"&_
			"VALUES ("&_
			Session("COD_USUARIO")&", "&COD_FORMANDO&", "&sqlTexto(LINHA1)&", "&sqlTexto(LINHA2)&", "&sqlTexto(LINHA3)&", "&sqlTexto(LINHA4)&", "&sqlTexto(LINHA5)&", "&sqlTexto(LINHA6)&_
			",1)"
								
			If ETIQUETA_ALBUM = 1 Then
				'1X
				objConexao.Execute(sql)
			Elseif ETIQUETA_ALBUM = 2 Then
				'2X
				objConexao.Execute(sql)
				objConexao.Execute(sql)
			End If
			
	End If
	
	
	
	If ETIQUETA_ALBUM_DUPLO <> 0 Then
	
	
	If Len (ALBUM) = 1 Then
			ALBUM = "000"&ALBUM 
		ElseIf Len(ALBUM) = 2 Then
			ALBUM = "00"&ALBUM 
		ElseIf Len(ALBUM) = 3 Then
			ALBUM = "0"&ALBUM 
		End If
	
		NOME_FORMANDO	= UCASE(Left(NOME_FORMANDO,23))
		
		
		
				
		Set rsTurmaForm = objConexao.Execute("SELECT DESCRICAO FROM TB_ENTIDADES WHERE COD_CONTRATO='"&COD_CONTRATO&"'")
						
						If not rsTurmaForm.Eof Then
						'Caso encontre coloca o nome da faculdade correspondente.
							TURMA = rsTurmaForm("DESCRICAO")
							'Correção do nome
							TURMA = Replace(TURMA,COD_CONTRATO&"    -   ","")
							TURMA = Replace(TURMA,COD_CONTRATO&"    -  ","")
							TURMA = Replace(TURMA,COD_CONTRATO&"   -   ","")
							TURMA = Replace(TURMA,COD_CONTRATO&"   - ","")
							TURMA = Replace(TURMA,COD_CONTRATO&" - ","")
							TURMA = Replace(TURMA,COD_CONTRATO&"- ","")
							TURMA = Replace(TURMA,COD_CONTRATO&" -","")
							TURMA = Replace(TURMA,COD_CONTRATO&"-","")
							TURMA = Replace(TURMA,COD_CONTRATO&" / ","")
							TURMA = Replace(TURMA,COD_CONTRATO,"")
							
							'2)Limpa o ano, semestre
							For i=2000 to 2015
							
							TURMA = Replace(TURMA,"-  /  J/"&i,"")
							TURMA = Replace(TURMA,"- /  J/"&i,"")
							TURMA = Replace(TURMA,"-  / J/"&i,"")
							TURMA = Replace(TURMA,"-/ J/"&i,"")
							TURMA = Replace(TURMA,"-/J/"&i,"")
							TURMA = Replace(TURMA,"-J/"&i,"")
							TURMA = Replace(TURMA,"J - /"&i,"")
							TURMA = Replace(TURMA,"/ J - "&i,"")
							TURMA = Replace(TURMA,"D/"&i,"")
							TURMA = Replace(TURMA,"J - D "&i,"")
							TURMA = Replace(TURMA,"J- /"&i,"")
							TURMA = Replace(TURMA,"-/ J"&i,"")
							TURMA = Replace(TURMA," - / J- "&i,"")
							TURMA = Replace(TURMA,"-  /  "&i,"")
							TURMA = Replace(TURMA,"- /  "&i,"")
							TURMA = Replace(TURMA,"J-/"&i,"")
							TURMA = Replace(TURMA,"-/"&i,"")
							TURMA = Replace(TURMA,"-/"&i,"")
							TURMA = Replace(TURMA,"-"&i,"")
							TURMA = Replace(TURMA,"J/"&i,"")
							TURMA = Replace(TURMA,"J"&i,"")
							TURMA = Replace(TURMA," J ","")
							TURMA = Replace(TURMA," D ","")
							TURMA = Replace(TURMA,i,"")
							
							Next
							
							TURMA = Replace(TURMA," - "," ")
							TURMA = Replace(TURMA,"- "," ")
							TURMA = Replace(TURMA,"-"," ")
							TURMA = Replace(TURMA,"/"," ")
							TURMA = Trim(TURMA)
							
							TURMA = Left(TURMA,23)
						End If
				
				rsTurmaForm.Close
				Set rsTurmaForm = Nothing
		
		
		
		
																				
		
									  '0 por que é álbum
		LINHA1 = chr(34)&COD_FORMANDO&"0"&chr(34)
		LINHA2 = chr(34)&NOME_FORMANDO&chr(34)
		LINHA3 = chr(34)& Left(TURMA,16)&" - "&COD_CONTRATO&chr(34)
		LINHA4 = chr(34)&NOME_CURSO&" - "&SEMESTRE_CONCLUSAO&"/"&Right(ANO_CONCLUSAO,2)&chr(34)
		LINHA5 = chr(34)&"ALBUM:"&ALBUM&"A -"&ArrumaData(Date())&chr(34)
		LINHA6 = chr(34)&"ALBUM DUPLO"&chr(34)
											
		sql="INSERT INTO TB_FICHAS_IDENTIFICACAO_ETIQUETAS_TEMPORARIA ("&_
			"COD_USUARIO ,"&_
			"COD_FORMANDO ,"&_
			"LINHA1 ,"&_
			"LINHA2 ,"&_
			"LINHA3 ,"&_
			"LINHA4 ,"&_
			"LINHA5 ,"&_
			"LINHA6 ,"&_
			"AGENDADA "&_
			")"&_
			"VALUES ("&_
			Session("COD_USUARIO")&", "&COD_FORMANDO&", "&sqlTexto(LINHA1)&", "&sqlTexto(LINHA2)&", "&sqlTexto(LINHA3)&", "&sqlTexto(LINHA4)&", "&sqlTexto(LINHA5)&", "&sqlTexto(LINHA6)&_
			",1)"
			
			If ETIQUETA_ALBUM_DUPLO = 1 Then
				'1X
				objConexao.Execute(sql)
			Elseif ETIQUETA_ALBUM_DUPLO = 2 Then
				'2X
				objConexao.Execute(sql)
				objConexao.Execute(sql)			
			End If
	
	End If
	
	If ETIQUETA_FICHA <> 0 Then
			
		
		If Len (ALBUM) = 1 Then
			ALBUM = "000"&ALBUM 
		ElseIf Len(ALBUM) = 2 Then
			ALBUM = "00"&ALBUM 
		ElseIf Len(ALBUM) = 3 Then
			ALBUM = "0"&ALBUM 
		End If
	
		NOME_FORMANDO	= UCASE(Left(NOME_FORMANDO,23))
		
		
		
		
				
		Set rsTurmaForm = objConexao.Execute("SELECT DESCRICAO FROM TB_ENTIDADES WHERE COD_CONTRATO='"&COD_CONTRATO&"'")
						
						If not rsTurmaForm.Eof Then
						'Caso encontre coloca o nome da faculdade correspondente.
							TURMA = rsTurmaForm("DESCRICAO")
							'Correção do nome
							TURMA = Replace(TURMA,COD_CONTRATO&"    -   ","")
							TURMA = Replace(TURMA,COD_CONTRATO&"    -  ","")
							TURMA = Replace(TURMA,COD_CONTRATO&"   -   ","")
							TURMA = Replace(TURMA,COD_CONTRATO&"   - ","")
							TURMA = Replace(TURMA,COD_CONTRATO&" - ","")
							TURMA = Replace(TURMA,COD_CONTRATO&"- ","")
							TURMA = Replace(TURMA,COD_CONTRATO&" -","")
							TURMA = Replace(TURMA,COD_CONTRATO&"-","")
							TURMA = Replace(TURMA,COD_CONTRATO&" / ","")
							TURMA = Replace(TURMA,COD_CONTRATO,"")
							
							'2)Limpa o ano, semestre
							For i=2000 to 2015
							
							TURMA = Replace(TURMA,"-  /  J/"&i,"")
							TURMA = Replace(TURMA,"- /  J/"&i,"")
							TURMA = Replace(TURMA,"-  / J/"&i,"")
							TURMA = Replace(TURMA,"-/ J/"&i,"")
							TURMA = Replace(TURMA,"-/J/"&i,"")
							TURMA = Replace(TURMA,"-J/"&i,"")
							TURMA = Replace(TURMA,"J - /"&i,"")
							TURMA = Replace(TURMA,"/ J - "&i,"")
							TURMA = Replace(TURMA,"D/"&i,"")
							TURMA = Replace(TURMA,"J - D "&i,"")
							TURMA = Replace(TURMA,"J- /"&i,"")
							TURMA = Replace(TURMA,"-/ J"&i,"")
							TURMA = Replace(TURMA," - / J- "&i,"")
							TURMA = Replace(TURMA,"-  /  "&i,"")
							TURMA = Replace(TURMA,"- /  "&i,"")
							TURMA = Replace(TURMA,"J-/"&i,"")
							TURMA = Replace(TURMA,"-/"&i,"")
							TURMA = Replace(TURMA,"-/"&i,"")
							TURMA = Replace(TURMA,"-"&i,"")
							TURMA = Replace(TURMA,"J/"&i,"")
							TURMA = Replace(TURMA,"J"&i,"")
							TURMA = Replace(TURMA," J ","")
							TURMA = Replace(TURMA," D ","")
							TURMA = Replace(TURMA,i,"")
							
							Next
							
							TURMA = Replace(TURMA," - "," ")
							TURMA = Replace(TURMA,"- "," ")
							TURMA = Replace(TURMA,"-"," ")
							TURMA = Replace(TURMA,"/"," ")
							TURMA = Trim(TURMA)
							
							TURMA = Left(TURMA,23)
						End If
				
				rsTurmaForm.Close
				Set rsTurmaForm = Nothing
		
		
		
		
																				
		
		'1 por que é FICHA
		LINHA1 = chr(34)&COD_FORMANDO&"1"&chr(34)
		LINHA2 = chr(34)&NOME_FORMANDO&chr(34)
		LINHA3 = chr(34)& Left(TURMA,16)&" - "&COD_CONTRATO&chr(34)
		LINHA4 = chr(34)&NOME_CURSO&" - "&SEMESTRE_CONCLUSAO&"/"&Right(ANO_CONCLUSAO,2)&chr(34)
		LINHA5 = chr(34)&"ALBUM:"&ALBUM&" - "&ArrumaData(Date())&chr(34)
		LINHA6 = chr(34)&"FICHA"&chr(34)
											
		sql="INSERT INTO TB_FICHAS_IDENTIFICACAO_ETIQUETAS_TEMPORARIA ("&_
			"COD_USUARIO ,"&_
			"COD_FORMANDO ,"&_
			"LINHA1 ,"&_
			"LINHA2 ,"&_
			"LINHA3 ,"&_
			"LINHA4 ,"&_
			"LINHA5 ,"&_
			"LINHA6 ,"&_
			"AGENDADA "&_
			")"&_
			"VALUES ("&_
			Session("COD_USUARIO")&", "&COD_FORMANDO&", "&sqlTexto(LINHA1)&", "&sqlTexto(LINHA2)&", "&sqlTexto(LINHA3)&", "&sqlTexto(LINHA4)&", "&sqlTexto(LINHA5)&", "&sqlTexto(LINHA6)&_
			",1)"

			objConexao.Execute(sql)
			
			
			
	End If
	
	'////////////////////////////
	'ETIQUETAS AGREGADOS
	'1) Verifica quais são os agregados
	'2) Verifica quais estão selecionados
	'3) Imprime cada qual de acordo
	
	
	'1)
			sql =" SELECT TB_PRODUTOS.NOME_PRODUTO, TB_PRODUTOS.COD_PRODUTO, TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.TIPO, TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO FROM "&_
				 " TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO WHERE "&_
				 " TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO="&COD_FORMANDO&_
				 " ORDER BY TIPO"

					
					Set rsProd = objConexao.Execute(sql)
					
					While not rsProd.Eof
					
					'2)
					Dim COD_PRODUTO_AGREGADO, COD_PRODUTO_AGREGADO_SELECIONADO
					COD_PRODUTO_AGREGADO 				= rsProd("COD_PRODUTO")
					COD_PRODUTO_AGREGADO_SELECIONADO	= formatNumber(0&Request.Form("ETIQUETA_PRODUTO_AGREGADO_"&COD_PRODUTO_AGREGADO),0)
					
						'3)
						If COD_PRODUTO_AGREGADO_SELECIONADO = 1 Then
						
							If Len (ALBUM) = 1 Then
								ALBUM = "000"&ALBUM 
							ElseIf Len(ALBUM) = 2 Then
								ALBUM = "00"&ALBUM 
							ElseIf Len(ALBUM) = 3 Then
								ALBUM = "0"&ALBUM 
							End If
					
							NOME_FORMANDO	= UCASE(Left(NOME_FORMANDO,23))
						
						
						
						
								
							Set rsTurmaForm = objConexao.Execute("SELECT DESCRICAO FROM TB_ENTIDADES WHERE COD_CONTRATO="&sqlTexto(COD_CONTRATO))
										
							If not rsTurmaForm.Eof Then
										'Caso encontre coloca o nome da faculdade correspondente.
											TURMA = rsTurmaForm("DESCRICAO")
											'Correção do nome
											TURMA = Replace(TURMA,COD_CONTRATO&"    -   ","")
											TURMA = Replace(TURMA,COD_CONTRATO&"    -  ","")
											TURMA = Replace(TURMA,COD_CONTRATO&"   -   ","")
											TURMA = Replace(TURMA,COD_CONTRATO&"   - ","")
											TURMA = Replace(TURMA,COD_CONTRATO&" - ","")
											TURMA = Replace(TURMA,COD_CONTRATO&"- ","")
											TURMA = Replace(TURMA,COD_CONTRATO&" -","")
											TURMA = Replace(TURMA,COD_CONTRATO&"-","")
											TURMA = Replace(TURMA,COD_CONTRATO&" / ","")
											TURMA = Replace(TURMA,COD_CONTRATO,"")
											
											'2)Limpa o ano, semestre
											For i=2000 to 2015
											
											TURMA = Replace(TURMA,"-  /  J/"&i,"")
											TURMA = Replace(TURMA,"- /  J/"&i,"")
											TURMA = Replace(TURMA,"-  / J/"&i,"")
											TURMA = Replace(TURMA,"-/ J/"&i,"")
											TURMA = Replace(TURMA,"-/J/"&i,"")
											TURMA = Replace(TURMA,"-J/"&i,"")
											TURMA = Replace(TURMA,"J - /"&i,"")
											TURMA = Replace(TURMA,"/ J - "&i,"")
											TURMA = Replace(TURMA,"D/"&i,"")
											TURMA = Replace(TURMA,"J - D "&i,"")
											TURMA = Replace(TURMA,"J- /"&i,"")
											TURMA = Replace(TURMA,"-/ J"&i,"")
											TURMA = Replace(TURMA," - / J- "&i,"")
											TURMA = Replace(TURMA,"-  /  "&i,"")
											TURMA = Replace(TURMA,"- /  "&i,"")
											TURMA = Replace(TURMA,"J-/"&i,"")
											TURMA = Replace(TURMA,"-/"&i,"")
											TURMA = Replace(TURMA,"-/"&i,"")
											TURMA = Replace(TURMA,"-"&i,"")
											TURMA = Replace(TURMA,"J/"&i,"")
											TURMA = Replace(TURMA,"J"&i,"")
											TURMA = Replace(TURMA," J ","")
											TURMA = Replace(TURMA," D ","")
											TURMA = Replace(TURMA,i,"")
											
											Next
											
											TURMA = Replace(TURMA," - "," ")
											TURMA = Replace(TURMA,"- "," ")
											TURMA = Replace(TURMA,"-"," ")
											TURMA = Replace(TURMA,"/"," ")
											TURMA = Trim(TURMA)
											
											TURMA = Left(TURMA,23)
										End If
								
								rsTurmaForm.Close
								Set rsTurmaForm = Nothing
						
						
						
						
																								
						LINHA1 = chr(34)&COD_FORMANDO&rsProd("TIPO")&chr(34)
						LINHA2 = chr(34)&NOME_FORMANDO&chr(34)
						LINHA3 = chr(34)& Left(TURMA,16)&" - "&COD_CONTRATO&chr(34)
						LINHA4 = chr(34)&NOME_CURSO&" - "&SEMESTRE_CONCLUSAO&"/"&Right(ANO_CONCLUSAO,2)&chr(34)
						LINHA5 = chr(34)&"ALBUM:"&ALBUM&" - "&ArrumaData(Date())&chr(34)
						LINHA6 = chr(34)&UCASE(Left(rsProd("NOME_PRODUTO"),9))&chr(34)
															
						sql="INSERT INTO TB_FICHAS_IDENTIFICACAO_ETIQUETAS_TEMPORARIA ("&_
							"COD_USUARIO ,"&_
							"COD_FORMANDO ,"&_
							"LINHA1 ,"&_
							"LINHA2 ,"&_
							"LINHA3 ,"&_
							"LINHA4 ,"&_
							"LINHA5 ,"&_
							"LINHA6 ,"&_
							"AGENDADA "&_
							")"&_
							"VALUES ("&_
							Session("COD_USUARIO")&", "&COD_FORMANDO&", "&sqlTexto(LINHA1)&", "&sqlTexto(LINHA2)&", "&sqlTexto(LINHA3)&", "&sqlTexto(LINHA4)&", "&sqlTexto(LINHA5)&", "&sqlTexto(LINHA6)&_
							",1)"
				
							objConexao.Execute(sql)
						
						End If
						
					rsProd.MoveNext()
					Wend
			
					rsProd.Close
					Set rsProd=Nothing
					
	'////////////////////////////
	
	
	FecharConexao()
	
	'Verifica se agenda ou imprime na hora
	If IMPRESSAO = 1 Then
	'Agendada
	Response.Redirect(redirecionar_volta)
	
	Elseif IMPRESSAO = 2 Then
	'Imprimir agora
	
	Dim voltar
	voltar = "dip_formandos_cadastro_album_impressao_etiquetas_imprimir.asp?cod_formando="&cod_formando&"&redit="&redirecionar_volta
	Response.Redirect(voltar)
	End If
	
	

End If
 	


%>
<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="iso-8859-1">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Impress&atilde;o de etiquetas | Portal Stillos</title>
  <link rel="stylesheet" href="css/vendor/bootstrap.min.css">
  <link rel="stylesheet" href="css/dashboard.css">
  <style>
    .labels-page .dashboard-nav a.active:not([aria-current="page"]) { background:transparent; color:#d4d5e1; }
    .labels-page .dashboard-nav a.active:hover { background:#37394f; color:#fff; }
    .labels-page .dashboard-nav a[aria-current="page"] { background:#b20b61; color:#fff; font-weight:600; }
    .labels-page .labels-panel { height:auto; max-width:960px; margin-inline:auto; padding:clamp(18px, 3vw, 32px); }
    .labels-panel h2 { margin-bottom:14px; }
    .labels-person { display:grid; grid-template-columns:minmax(0, 2fr) minmax(0, 1fr); gap:12px 24px; margin:0 0 24px; padding:16px; border-radius:8px; background:#f7f5fa; }
    .labels-person dt { color:#53566b; font-size:12px; margin-bottom:3px; }
    .labels-person dd { font-weight:650; margin:0; overflow-wrap:anywhere; }
    .labels-group { min-width:0; margin:0 0 18px; padding:0; border:0; }
    .labels-group legend, .labels-mode label { float:none; font-size:14px; font-weight:650; margin-bottom:8px; }
    .labels-choices { display:flex; flex-wrap:wrap; gap:8px; }
    .labels-choice { display:flex; align-items:center; gap:9px; padding:10px 13px; border:1px solid #dddde7; border-radius:8px; cursor:pointer; overflow-wrap:anywhere; }
    .labels-choice:has(input:checked) { border-color:#b20b61; background:#fdf1f7; }
    .labels-choice .form-check-input { flex:none; margin:0; }
    .labels-choice .form-check-input:checked { background-color:#b20b61; border-color:#b20b61; }
    .labels-mode .form-select { max-width:280px; }
    .labels-actions { display:flex; justify-content:flex-end; margin-top:24px; padding-top:16px; border-top:1px solid var(--line); }
    @media(max-width:480px) {
      .labels-person { grid-template-columns:1fr; }
      .labels-choices { flex-direction:column; }
      .labels-mode .form-select, .labels-actions .btn { max-width:none; width:100%; }
    }
  </style>
</head>
<body class="labels-page">
<a class="skip-link" href="#conteudo">Ir para o conte&uacute;do</a>
<div class="dashboard-layout">
  <aside class="dashboard-sidebar" id="menuPrincipal">
    <a class="dashboard-brand" href="adm_index.asp"><img src="images/logoparafundopreto.png" alt="Stillos" width="180" /><span>GERENCIAL FOTOS</span></a>
    <div class="nav-caption">&Aacute;REA DE TRABALHO</div>
    <!-- #include file="includes/dashboard_menu.asp" -->
  </aside>
  <div class="dashboard-workspace">
    <header class="dashboard-topbar">
      <button class="btn menu-toggle" id="ppvMenuToggle" type="button" aria-controls="menuPrincipal" aria-expanded="false">Menu</button>
      <div class="topbar-title">Portal Stillos <span>/ Impress&atilde;o de etiquetas</span></div>
      <div class="topbar-actions"><a class="user-name" href="adm_log_usuarios.asp"><%=Server.HTMLEncode(Session("NOME") & "")%></a><a href="manual_sistema.asp" target="_blank" rel="noopener">Manual</a><a class="btn btn-outline-secondary btn-sm" href="adm_sair.asp">Sair</a></div>
    </header>
    <main class="dashboard-main ppv-main" id="conteudo">
      <div class="dashboard-heading">
        <div>
          <h1>Impress&atilde;o de etiquetas</h1>
          <p>Selecione as etiquetas e escolha quando imprimir.</p>
        </div>
        <%If imprimir_agora = 1 Then%>
        <a class="btn btn-outline-secondary" href="dip_busca_detalhamento_formando.asp?cod_formando=<%=Server.URLEncode(COD_FORMANDO & "")%>">Voltar &agrave; ficha</a>
        <%Else%>
        <a class="btn btn-outline-secondary" href="dip_formandos.asp">Voltar aos formandos</a>
        <%End If%>
      </div>
      <div class="search-panel"><strong>Busca no portal</strong><!-- #include file="includes/dashboard_busca.asp" --></div>
      <!-- #include file="includes/dashboard_avisos.asp" -->
      <%If Len(Request.QueryString("msg") & "") > 0 Then%>
      <p class="alert alert-info" role="status"><%=Server.HTMLEncode(Request.QueryString("msg") & "")%></p>
      <%End If%>
      <section class="dashboard-panel labels-panel" aria-labelledby="etiquetas-titulo">
        <h2 id="etiquetas-titulo">Etiquetas de identifica&ccedil;&atilde;o</h2>
        <dl class="labels-person">
          <div><dt>Formando</dt><dd><%=Server.HTMLEncode(NOME_FORMANDO & "")%></dd></div>
          <div><dt>Curso</dt><dd><%=Server.HTMLEncode(NOME_CURSO & "")%></dd></div>
        </dl>
        <form name="CAD_FORMANDOS_3" method="post" action="dip_formandos_cadastro_album_impressao_etiquetas.asp?ia=<%=imprimir_agora%>&amp;cod_formando=<%=COD_FORMANDO%>">
        <%If HABILITA_IMPRESSAO_ALBUM = 1 Then%>
        <fieldset class="labels-group">
          <legend>Etiqueta de &aacute;lbum</legend>
          <div class="labels-choices">
            <label class="labels-choice"><input class="form-check-input" type="radio" name="ETIQUETA_ALBUM" value="1" checked> 1 etiqueta</label>
            <label class="labels-choice"><input class="form-check-input" type="radio" name="ETIQUETA_ALBUM" value="2"> 2 etiquetas</label>
            <label class="labels-choice"><input class="form-check-input" type="radio" name="ETIQUETA_ALBUM" value="0"> Nenhuma</label>
          </div>
        </fieldset>
        <%End If%>
        <%If HABILITA_IMPRESSAO_ALBUM_DUPLO = 1 Then%>
        <fieldset class="labels-group">
          <legend>Etiqueta de &aacute;lbum duplo</legend>
          <div class="labels-choices">
            <label class="labels-choice"><input class="form-check-input" type="radio" name="ETIQUETA_ALBUM_DUPLO" value="1" checked> 1 etiqueta</label>
            <label class="labels-choice"><input class="form-check-input" type="radio" name="ETIQUETA_ALBUM_DUPLO" value="2"> 2 etiquetas</label>
            <label class="labels-choice"><input class="form-check-input" type="radio" name="ETIQUETA_ALBUM_DUPLO" value="0"> Nenhuma</label>
          </div>
        </fieldset>
        <%End If%>
        <fieldset class="labels-group">
          <legend>Ficha de identifica&ccedil;&atilde;o e produtos</legend>
          <div class="labels-choices">
            <label class="labels-choice"><input class="form-check-input" type="checkbox" name="ETIQUETA_FICHA" value="1"> Ficha de identifica&ccedil;&atilde;o &middot; 1 etiqueta</label>
<%
AbrirConexao()
											
										Dim rsProd
										sql = "SELECT TB_PRODUTOS.COD_PRODUTO, TB_PRODUTOS.NOME_PRODUTO, Sum( TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.QUANTIDADE ) AS QUANTIDADE"&_
											  " FROM TB_PRODUTOS"&_
											 " INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO"&_
											" WHERE TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO = "&COD_FORMANDO&_
											" GROUP BY TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO ORDER BY TB_PRODUTOS.NOME_PRODUTO"
																	'Response.Write(sql)
										Set rsProd = objConexao.Execute(sql)
					
										While not rsProd.Eof%>
            <label class="labels-choice"><input class="form-check-input" type="checkbox" name="ETIQUETA_PRODUTO_AGREGADO_<%=rsProd("COD_PRODUTO")%>" value="1"> <%=Server.HTMLEncode(rsProd("NOME_PRODUTO") & "")%> &middot; 1 etiqueta</label>
<%
										rsProd.MoveNext()
										Wend
										
									rsProd.Close()		
									Set rsProd = Nothing
FecharConexao()%>
          </div>
        </fieldset>
        <div class="labels-mode">
          <label for="labels-print-mode">Modo de impress&atilde;o</label>
          <select class="form-select" id="labels-print-mode" name="IMPRESSAO">
            <option value="1" selected>Agendada</option>
            <option value="2">Imprimir agora</option>
          </select>
        </div>
        <div class="labels-actions">
          <button class="btn btn-primary" type="submit" name="Grava3" value=" ">Avan&ccedil;ar</button>
        </div>
        </form>
      </section>
    </main>
    <footer class="dashboard-footer">&copy; <%=Year(Date)%> Grupo Stillos <span>Sistema administrativo interno</span></footer>
  </div>
</div>
<script>
(function () {
    var toggle = document.getElementById('ppvMenuToggle');
    var menu = document.getElementById('menuPrincipal');
    var previous = menu.querySelector('[aria-current="page"]');
    var current = menu.querySelector('a[href="dip_formandos.asp"]');
    if (previous) previous.removeAttribute('aria-current');
    if (current) current.setAttribute('aria-current', 'page');
    toggle.addEventListener('click', function () {
        var expanded = menu.classList.toggle('is-open');
        toggle.setAttribute('aria-expanded', String(expanded));
    });
    document.addEventListener('keydown', function (event) {
        if (event.key === 'Escape' && menu.classList.contains('is-open')) {
            menu.classList.remove('is-open');
            toggle.setAttribute('aria-expanded', 'false');
            toggle.focus();
        }
    });
})();
</script>
</body>
</html>
