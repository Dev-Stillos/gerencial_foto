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
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
	<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>
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
	</script>
	
	<script type="text/javascript" language="javascript">
	<!--AJUDA F1 E CTRL PARA BUSCA RÁPIDA-->
		
	function abrir(theURL,winName,features) 
	{ 
	  window.open(theURL,winName,features);
	}

		
	function abre(url)
	{    
				//if(event.srcElement.tagName=='TD')
				//	{
				//		src.children.tags('A')[0].click(); 
				//	}
	location.href = url;

	}	
		
		
	

	function buscaIni(){
	d = document.buscaForm
	d.busca.focus()
	d.busca.select();
	}
	function mudacorBusca(CAMPO){
	CAMPO.style.backgroundColor = '#F5ECBA'
	}
	function mudacorBusca2(CAMPO){
	CAMPO.style.backgroundColor = '#FFFFFF'
	}

 function Verificar()
 {
  var tecla=window.event.keyCode;
  if (tecla==17) {
    buscaIni();
	}
 }
 


	<!--AJUDA F1 E CTRL PARA BUSCA RÁPIDA-->
	</script>
		<!-- CALENDÁRIO SCRIPTS
	<script type="text/javascript">
	DHTMLSuite.include("calendar");
	</script>
	<script type="text/javascript">
	var calendarObjForForm = new DHTMLSuite.calendar({minuteDropDownInterval:10,numberOfRowsInHourDropDown:5,callbackFunctionOnDayClick:'getDateFromCalendar',isDragable:true,displayTimeBar:false});
	
		 
	calendarObjForForm.setCallbackFunctionOnClose('myOtherFunction');
	
	function myOtherFunction()
	{
		
		
	}
	function pickDate(buttonObj,inputObject)
	{
		calendarObjForForm.setCalendarPositionByHTMLElement(inputObject,0,inputObject.offsetHeight+2);	// Position the calendar right below the form input
		//calendarObjForForm.setInitialDateFromInput(inputObject,'yyyy-mm-dd hh:ii');	// Specify that the calendar should set it's initial date from the value of the input field.
		
		calendarObjForForm.setInitialDateFromInput(inputObject,'dd/mm/yyyy');	// Specify that the calendar should set it's initial date from the value of the input field.
		
		calendarObjForForm.addHtmlElementReference('myDate',inputObject);	// Adding a reference to this element so that I can pick it up in the getDateFromCalendar below(myInput is a unique key)
		if(calendarObjForForm.isVisible()){
			calendarObjForForm.hide();
		}else{
			calendarObjForForm.resetViewDisplayedMonth();	// This line resets the view back to the inital display, i.e. it displays the inital month and not the month it displayed the last time it was open.
			calendarObjForForm.display();
		}		
	}	
	/* inputArray is an associative array with the properties
	year
	month
	day
	hour
	minute
	calendarRef - Reference to the DHTMLSuite.calendar object.
	*/
	function getDateFromCalendar(inputArray)
	{
		var references = calendarObjForForm.getHtmlElementReferences(); // Get back reference to form field.
		references.myDate.value = inputArray.day + '/' + inputArray.month + '/' + inputArray.year
//		references.myDate.value = inputArray.year + '-' + inputArray.month + '-' + inputArray.day + ' ' + inputArray.hour + ':' + inputArray.minute;
		calendarObjForForm.hide();	
		
	}	
	</script>
	-->
	<script language="javascript">
function YNconfirm(txt){
return typeof(suporteVBscript)=="undefined"?confirm(txt):VBconfirm(txt)==6
}
</script>
<script language="vbscript">
suporteVBscript=1
function VBconfirm(mensagem)

 VBconfirm=msgbox(mensagem,vbyesno)
end function

</script>
	<script language="javascript">
	function validaForm(){
	d = document.CAD_FORMANDOS_2
		if(d.ALBUM.value == ""){
		d.ALBUM.focus();
		alert('Campo álbum não digitado!');
		return false;
		}
	
		if(d.INICIAL.value == ""){
		d.INICIAL.focus();
		alert('Campo inicial não digitado!');
		return false;
		}
	
		if(d.ALBUM_EM_POSSE.value == 3 && d.FICHAS_EM_POSSE.value == 27){
		d.ALBUM_EM_POSSE.focus();
		alert('Você informou que o álbum está com o vendedor, mas não informou qual vendedor é!');
		return false;
		}
	
	}
	</script>
	
	<script language="javascript">
		function popUp(campo,URL,campoSelect,msgErro){
			
			if(document.CAD_FORMANDOS_2[campo].value==1){
				
				if (campoSelect !== 'no'){
					//Brinde = "Sim"
					document.CAD_FORMANDOS_2[campoSelect].options[1].selected = "true";
				}
			
				win=window.open(URL,"","width=350,height=220,scrollbars=yes")
				var text = "Se a janela não abrir desbloquei os Pop-ups!";
				if(win == null) { alert(text); return; }
				win.moveTo(350,150);
					
			}
			else {
				if (msgErro==1){alert('Para visualizar os brindes é necessário alterar a opção Brinde para Sim.')}
			}
		
		}
		
		
		function fecharPopUp(){
			if(win){
				win.close();
			}
		}
		
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
			
		function focuInic(){
		document.CAD_FORMANDOS_3.Grava3.focus();
		}	
	</script>
</head>

<body onload="focuInic()">

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
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">

	  <h2>DIP - Impress&atilde;o de Etiqueta de Identifica&ccedil;&atilde;o </h2>
	   <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	  <h3 align="center">Selecione as etiquetas que deseja imprimir  </h3>
		  <div id="comboTexto">
		  	<form name="CAD_FORMANDOS_3" method="post" action="dip_formandos_cadastro_album_impressao_etiquetas.asp?ia=<%=imprimir_agora%>&cod_formando=<%=COD_FORMANDO%>"> 
          	<table width="95%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
		        	<tr>
            <td width="24%" bgcolor="#CCCCCC"><strong>Nome Formando: </strong></td>
            <td width="76%" bgcolor="#CCCCCC"><strong><%=NOME_FORMANDO%></strong></td>
            </tr>
					<tr>
				<td width="24%" bgcolor="#CCCCCC"><strong>Curso:</strong></td>
            <td width="76%" bgcolor="#CCCCCC"><strong><%=NOME_CURSO%></strong></td>
            </tr>
			
			<tr>
            <td colspan="2"><div align="center"><strong>Impress&atilde;o de Etiqueta </strong></div></td>
            </tr>
			
          
		 <%If HABILITA_IMPRESSAO_ALBUM = 1 Then%>
		  <tr>
            <td> Etiqueta de &Aacute;lbum: </td>
            <td><table width="300" border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td width="17%">1 Etiqueta</td>
                <td width="8%"><div align="center">
                  <input class="inputRadio" checked="checked" name="ETIQUETA_ALBUM" type="radio" value="1" />
                </div></td>
                <td width="24%"><div align="center">2 Etiquetas </div></td>
                <td width="12%">
                  <div align="left">
                    <input  class="inputRadio" name="ETIQUETA_ALBUM" type="radio" value="2" />
                    </div></td>
                <td width="17%">Nenhuma</td>
                <td width="22%"><input class="inputRadio" name="ETIQUETA_ALBUM" type="radio" value="0" /></td>
              </tr>
            </table></td>
            </tr>
		<%End If%>		
			
		  <%If HABILITA_IMPRESSAO_ALBUM_DUPLO = 1 Then%>
          <tr>
            <td height="38"> Etiqueta de &Aacute;lbum Duplo:</td>
            <td><table width="300" border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td width="17%">1 Etiqueta</td>
                <td width="8%"><div align="center">
                  <input checked="checked" class="inputRadio"  name="ETIQUETA_ALBUM_DUPLO" type="radio" value="1" />
                </div></td>
                <td width="24%"><div align="center">2 Etiquetas </div></td>
                <td width="12%">
                  <div align="left">
                    <input class="inputRadio" name="ETIQUETA_ALBUM_DUPLO" type="radio" value="2" />
                    </div></td>
                <td width="17%">Nenhuma</td>
                <td width="22%"><input class="inputRadio" name="ETIQUETA_ALBUM_DUPLO" type="radio" value="0" /></td>
              </tr>
            </table></td>
            </tr>		
		  <%End If%>
          
		  <tr>
            <td>Ficha de Identifica&ccedil;&atilde;o:</td>
            <td>1 Etiqueta
              <input class="inputRadio" type="checkbox" name="ETIQUETA_FICHA" value="1" /></td>
            </tr>
			
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
									  <tr>
            <td><%=rsProd("NOME_PRODUTO")%>:</td>
            <td>1 Etiqueta
              <input class="inputRadio" type="checkbox" name="ETIQUETA_PRODUTO_AGREGADO_<%=rsProd("COD_PRODUTO")%>" value="1" /></td>
            </tr>	
									<%
										rsProd.MoveNext()
										Wend
										
									rsProd.Close()		
									Set rsProd = Nothing
FecharConexao()
									%>	
			<tr>
            <td>Impress&atilde;o:</td>
            <td><select name="IMPRESSAO">
				<option value="1" selected="selected">AGENDADA</option>
				<option value="2">IMPRIMIR AGORA</option>
            </select>
            </td>
            </tr>
			
          <tr>
            <td colspan="2"><div align="center">
              <input class="btnAvancar" type="submit" name="Grava3" value=" " />
            </div></td>
            </tr>
		         </table>
		  	</form>
		  </div>
		
		  <p>&nbsp;</p>



				
		</td>

    </tr>
</table>

 
  
  </td>
    </tr>
</table>

  
  </td>
  </tr>
</table>
<!-- #include file="includes/rodape.asp" -->

  </td>
</tr>
</table>
</body>
</html>
