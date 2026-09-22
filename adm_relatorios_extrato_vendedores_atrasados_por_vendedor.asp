<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Validação especial para verificar se possui sub-validação específica
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '15' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se não encontrar então redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
		
	rs.Close
	Set rs=Nothing
	
FecharConexao()



Dim COD_VENDEDOR_BLOQUEADO, COM_COPIA, SOMENTE_VENDEDORES_EM_ATRASO
COD_VENDEDOR_BLOQUEADO = formatNumber(0&Request.QueryString("COD_VENDEDOR_BLOQUEADO"),0)
COM_COPIA			   = formatNumber(0&Request.QueryString("COM_COPIA"),0)
SOMENTE_VENDEDORES_EM_ATRASO = Cstr(""&Request.QueryString("SOMENTE_VENDEDORES_EM_ATRASO"))



If SOMENTE_VENDEDORES_EM_ATRASO = "" AND Request.QueryString("COD_VENDEDOR") <> Empty Then
'Significa que foi solicitado à partir da página 'adm_relatorios.asp', sendo que foi solicitado todos os itens em posse do vendedor, exceto fichas.
	SOMENTE_VENDEDORES_EM_ATRASO = "N"
	COD_VENDEDOR_BLOQUEADO = Request.QueryString("COD_VENDEDOR")
End If

If COD_VENDEDOR_BLOQUEADO = 0 Then
	Response.Redirect("adm_relatorios_vendedores_bloqueados.asp")
End If

	'1º Parte Cria os arquivos vazios e da nome as futuros arquivos
	'//////////////////////////////////////////////////////////////////////////////////////////////////
	
		Dim COD_ALEATORIO
		COD_ALEATORIO			= Now()
		COD_ALEATORIO			= Replace(COD_ALEATORIO,":","")
		COD_ALEATORIO			= Replace(COD_ALEATORIO,",","")
		COD_ALEATORIO			= Replace(COD_ALEATORIO,"/","")
		COD_ALEATORIO			= Replace(COD_ALEATORIO," ","")
		NOME_ARQUIVO_MASTER 	= "master2_"&COD_ALEATORIO&"_"&Session("COD_USUARIO")
		NOME_ARQUIVO_INDIVIDUAL = "individual2_"&COD_ALEATORIO&"_"&Session("COD_USUARIO")
		NOME_ARQUIVO_HIPER		= "hiper2_"&COD_ALEATORIO&"_"&Session("COD_USUARIO")

		SET AspPdf = Server.CreateObject("Persits.Pdf")
		SET Doc = AspPdf.CreateDocument
		SET Page = Doc.Pages.Add(216, 280)
		'Define os parâmetros de alinhamento: X, Y do canto superior esquerdo ao lado inferior direito, altura, largura e tamanho da fonte.
		Filename = Doc.Save( "e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf") 
		SET AspPdf = Nothing
		SET Doc = Nothing
		SET Page = Nothing
		SET Font = Nothing
		SET Param = Nothing
		
		SET AspPdf = Server.CreateObject("Persits.Pdf")
		SET Doc = AspPdf.CreateDocument
		SET Page = Doc.Pages.Add( 216, 280 )
		'Define os parâmetros de alinhamento: X, Y do canto superior esquerdo ao lado inferior direito, altura, largura e tamanho da fonte.
		Filename = Doc.Save( "e:\home\stillos\web\gerencial_foto\documentos\hiper.pdf")
		SET AspPdf = Nothing
		SET Doc = Nothing
		SET Page = Nothing
		SET Font = Nothing
		SET Param = Nothing

'/////////////
'Faz a média de dias
AbrirConexao()

Dim sql, registros, qtd_paginas, data_inicial, registros_totais,qtd_dias_devedores_padrao

' Variáveis fora de uso
'qtd_dias_devedores_padrao = -46
'data_inicial = dateAdd("d",qtd_dias_devedores_padrao,Now())
'data_inicial = Year(data_inicial)&"-"&Month(data_inicial)&"-"&Day(data_inicial)
 		
'Forma Antiga
'sql="SELECT (SUM((DATEDIFF(NOW(),DATA_SAIDA))"&qtd_dias_devedores_padrao&")/COUNT(*))  AS QTD_DIAS_FORA  FROM VW_MOVIMENTACOES_COMPLETA WHERE DATA_RETORNO IS NULL AND TIPO <> 1 AND DATA_SAIDA <= '"&data_inicial&" 23:59:59'  AND COD_VENDEDOR = "&COD_VENDEDOR_BLOQUEADO
		
	
	
	'Forma Nova
	If SOMENTE_VENDEDORES_EM_ATRASO = "N" Then
		sqlMedia = "SELECT AVG(DIAS_ATRASADOS) AS MEDIA FROM VW_MOVIMENTACOES_COMPLETA WHERE TIPO <> 1 AND DATA_RETORNO IS NULL AND COD_VENDEDOR = "&COD_VENDEDOR_BLOQUEADO
		sql="SELECT * FROM VW_MOVIMENTACOES_COMPLETA WHERE TIPO <> 1 AND COD_VENDEDOR = "&COD_VENDEDOR_BLOQUEADO&" AND  DATA_RETORNO IS NULL ORDER BY NOME_FORMANDO"
	Else
		sqlMedia = "SELECT AVG(DIAS_ATRASADOS) AS MEDIA FROM VW_MOVIMENTACOES_COMPLETA WHERE TIPO <> 1 AND DATA_RETORNO IS NULL AND DIAS_ATRASADOS >= 0 AND COD_VENDEDOR = "&COD_VENDEDOR_BLOQUEADO
		sql="SELECT * FROM VW_MOVIMENTACOES_COMPLETA WHERE TIPO <> 1 AND COD_VENDEDOR = "&COD_VENDEDOR_BLOQUEADO&" AND  DATA_RETORNO IS NULL AND DIAS_ATRASADOS >= 0 ORDER BY NOME_FORMANDO"
	End If
	
	Set rs2 = objConexao.Execute(sqlMedia)
		
		If not rs2.Eof Then
			MEDIA_GERAL = formatNumber(rs2("MEDIA"),0)
		End If
				
	rs2.Close
	Set rs2 = Nothing

'/////////////
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
</head>
<body>
<%
'Forma antiga
'sql="SELECT * FROM VW_MOVIMENTACOES_COMPLETA WHERE DATA_RETORNO IS NULL AND TIPO <> 1 AND DATA_SAIDA <= '"&data_inicial&" 23:59:59' AND COD_VENDEDOR = "&COD_VENDEDOR_BLOQUEADO&" ORDER BY NOME_FORMANDO"	
 
 Set rs = objConexao.Execute(sql,registros) 
 	
	registros_totais = registros
	'registros = 71
 	qtd_paginas = int(registros/70)
	resto = registros Mod 70
	
	If qtd_paginas > 70 Then
		qtd_paginas = 1
	End If
	
	If resto > 0 Then
		qtd_paginas = qtd_paginas + 1
	End If
	
	For i=0 to qtd_paginas-1
	
		TOTAL_PAGINA = 0
		'Abri o PDF	
		'////////////////////////
		Set PDF = Server.CreateObject("Persits.PDF")

		' Open an existing document, form W-9
		Set Doc = PDF.OpenDocument("e:\home\stillos\web\gerencial_foto\documentos\impressao_relatorio_itens_atrasados.pdf")

		' Obtain Page object for page 1
		Set Page = Doc.Pages(1)


		' Create font object
		Set Font = Doc.Fonts("Arial")
		Set Param = PDF.CreateParam("html=true")
		
		'////////////////////////
		'
		Dim NOME_VENDEDOR, DATA_EMISSAO, PRAZO_NORMAL, TOTAL_PAGINA, HORAS, MINUTOS
				
		HORAS			= Hour(Now())
		MINUTOS			= Minute(Now())
		
		If HORAS < 10 Then
				HORAS = "0"&HORAS
		End If
		
		If MINUTOS < 10 Then
				MINUTOS = "0"&MINUTOS
		End If
		
		NOME_VENDEDOR	= "<b>"&Ucase(rs("NOME_VENDEDOR"))&"</b>"
		DATA_EMISSAO	= "<b>"&ArrumaData(date()) &" "&HORAS&"h"&MINUTOS&"</b>"
		
		
		PRAZO_NORMAL	= "<b>45 dias</b>"
		
		Param.Add("x=18, y=789, size=9")
		Page.Canvas.DrawText NOME_VENDEDOR, Param, Font
		
		Param.Add("x=270, y=789, size=9")
		Page.Canvas.DrawText DATA_EMISSAO, Param, Font
		
		Param.Add("x=360, y=789, size=9")
		Page.Canvas.DrawText "Variável", Param, Font
		
		Param.Add("x=495, y=789, size=9")
		Page.Canvas.DrawText "<b>"&registros_totais&"</b>", Param, Font
		
		Param.Add("x=424, y=789, size=9")
		Page.Canvas.DrawText "<b>"&MEDIA_GERAL&" dia(s)</b>", Param, Font
		
		Param.Add("x=540, y=20, size=7")
		Page.Canvas.DrawText "Pág. "&i+1&" de "&qtd_paginas, Param, Font
		
		
		
		
		
			For ii=1 to 70
				registros = registros - 1
				
				
				If registros => 0 Then
				'///////////////
						TOTAL_PAGINA = TOTAL_PAGINA + 1
				
						Dim COD_FORMANDO, ALBUM, NOME_FORMANDO, DESCRICAO, TIPO, DATA_SAIDA, DIAS_ATRASADOS
						
						
						TIPO = formatNumber(0&rs("TIPO"),0)
						DATA_SAIDA = rs("DATA_SAIDA")
						'/////
						
						ENCADERNACOES_EXTRAS = formatNumber(0&rs("ENCADERNACOES_EXTRAS"),0)		

						If TIPO = 0 Then
								
								If ENCADERNACOES_EXTRAS = 1 Then
									TIPO = "ÁLBUM DUPLO"
								Else
									TIPO = "ÁLBUM"
								End If
						
						Elseif TIPO = 1 Then
								TIPO = "FICHA"
						Elseif TIPO > 1 Then
								'Verifica qual o produto agregado
									Dim rsProdAgrePend
									sql = 	"SELECT TB_PRODUTOS.NOME_PRODUTO "&_
											"FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = "&_
											"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO WHERE "&_
											"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO = "&rs("COD_FORMANDO")&" AND "&_
											"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.TIPO = "&TIPO&" LIMIT 1"
									Set rsProdAgrePend = objConexao.Execute(sql)
										If not rsProdAgrePend.Eof Then
									
											TIPO = rsProdAgrePend("NOME_PRODUTO")		
									
										End If
									rsProdAgrePend.Close
									Set rsProdAgrePend = Nothing
								
							End If 	
						
						
						
						
						'/////
						DIAS_ATRASADOS = rs("DIAS_ATRASADOS")
						
						
						y = 751 - (ii*9)
						Param.Add("x=20, y="&y&", size=7")
						Page.Canvas.DrawText digito(rs("COD_FORMANDO")), Param, Font
						
						Param.Add("x=61, y="&y&", size=7")
						Page.Canvas.DrawText rs("ALBUM"), Param, Font
						
						Param.Add("x=85, y="&y&", size=7")
						Page.Canvas.DrawText Ucase(Left(rs("NOME_FORMANDO"),28)), Param, Font
						
						Param.Add("x=218, y="&y&", size=7")
						Page.Canvas.DrawText Ucase(Left(rs("DESCRICAO"),35)), Param, Font
						
						TIPO = Replace(TIPO,"Trio (2 porta-retrados + 1 foto-Tela 20X30)","TRIO")
						Param.Add("x=380, y="&y&", size=7")					
						Page.Canvas.DrawText Ucase(TIPO), Param, Font
						
						Param.Add("x=452, y="&y&", size=7")					
						Page.Canvas.DrawText arrumadata(DATA_SAIDA), Param, Font
						
						Param.Add("x=542, y="&y+1&", size=8")					
						Page.Canvas.DrawText DIAS_ATRASADOS, Param, Font
						
						
						
						
				rs.MoveNext()
				'///////////////
				End If
						
						
			Next
				
			'Total por Página
			Param.Add("x=557, y=105, size=8")
			Page.Canvas.DrawText TOTAL_PAGINA, Param, Font	
				
			'Fecha o arquivo PDF
				'////////////////////////
				FileName = Doc.Save("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_INDIVIDUAL&".pdf",true)
				Set Page = Nothing
				Set Doc = Nothing
				Set Pdf = Nothing
				
				'////////////
				'Cópia (2 vias)
				Dim ABRIR_ARQUIVO, QTD_PAGINAS_PDF
				
				IF COM_COPIA = 1 THEN
				
					
					ABRIR_ARQUIVO = "e:\home\stillos\web\gerencial_foto\documentos\hiper_copiados"&NOME_ARQUIVO_INDIVIDUAL&".pdf"
					QTD_PAGINAS_PDF   = 2
					'Para copiar é necessário:
					'1) Duplicar o arquivo individual FSO
					'2) Concatenar os PDF (2 individuais)
					'3) Concatenar o PDF com o MASTER e dar continuide
				
						Set objFSO = CreateObject("Scripting.FileSystemObject")

				objFSO.CopyFile "e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_INDIVIDUAL&".pdf","e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_INDIVIDUAL&"2.pdf"
				
				Set objFSO = Nothing
				
				Set Pdf = Server.CreateObject("Persits.PDF")
								
				Set a = Pdf.OpenDocument("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_INDIVIDUAL&".pdf")
				Set b = Pdf.OpenDocument("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_INDIVIDUAL&"2.pdf") 
				
				a.AppendDocument b
				FileName = a.Save("e:\home\stillos\web\gerencial_foto\documentos\hiper_copiados"&NOME_ARQUIVO_INDIVIDUAL&".pdf",true)			
				
				
				Set a = Nothing
				Set b = Nothing
				Set Pdf = Nothing
				ELSE
				'Apenas 1 via
					ABRIR_ARQUIVO = "e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_INDIVIDUAL&".pdf"				
					QTD_PAGINAS_PDF	  = 1
				END IF
				
				'////////////
				
				
				Set PDF = Server.CreateObject("Persits.PDF")
				
				
			
				'Abri o master e o acabado de ser criado
				Set DocMaster 		= PDF.OpenDocument("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf")				
				Set DocIndividual 	= PDF.OpenDocument(ABRIR_ARQUIVO)
		
				
				Set Page = DocIndividual.Pages(QTD_PAGINAS_PDF)
				'Concatena os 2 arquivos
				DocMaster.AppendDocument DocIndividual
			
				'Salva os arquivos
				FileName = DocMaster.Save("e:\home\stillos\web\gerencial_foto\documentos\hiper.pdf",true)
				'FileName = DocMaster.SaveHttp("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf")
				
				'Fechar os Arquivos
				Set DocMaster		= Nothing
				Set DocIndividual 	= Nothing
				Set PDF = Nothing
				
				'Deleta o Master
				Set fso = CreateObject("Scripting.FileSystemObject") 
					If fso.FileExists("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf") then
						fso.DeleteFile("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf")
					End if
				Set fso = Nothing
				
				'Renomeia o Hiper para Master
				set fso = CreateObject("Scripting.FileSystemObject") 
				fso.MoveFile "e:\home\stillos\web\gerencial_foto\documentos\hiper.pdf", "e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf"
				set fso = Nothing
	
		Response.Write("---------------------------------------------")
		Response.Write("<br />")
	Next
	 
 rs.Close
 Set rs = Nothing
 
FecharConexao()	

'///////////////////////////////////////////////////////

		'Retira a primeira página e exibe na tela
		Set Pdf = Server.CreateObject("Persits.Pdf") 
		Set Doc = Pdf.OpenDocument("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf")
	
		Doc.Pages.Remove 1
		'FileName = Doc.SaveHttp("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf")
		Doc.SaveHttp "attachment;filename=extrato_por_vendedor.pdf"
	
		Set Doc = Nothing
		Set Pdf = Nothing
		'////////////////////////////

'///////////////////////////////////////////////////////
'Exclui arquivos remanecentes
set fso = CreateObject("Scripting.FileSystemObject") 

If fso.FileExists("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_INDIVIDUAL&".pdf") then
fso.DeleteFile("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_INDIVIDUAL&".pdf")
end if

If fso.FileExists("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_INDIVIDUAL&"2.pdf") then
fso.DeleteFile("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_INDIVIDUAL&"2.pdf")
end if

If fso.FileExists("e:\home\stillos\web\gerencial_foto\documentos\hiper_copiados"&NOME_ARQUIVO_INDIVIDUAL&".pdf") then
fso.DeleteFile("e:\home\stillos\web\gerencial_foto\documentos\hiper_copiados"&NOME_ARQUIVO_INDIVIDUAL&".pdf")
end if

		
If fso.FileExists("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf") then
fso.DeleteFile("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf")
end if
		
Set fso = Nothing		
%>
</body>
</html>












