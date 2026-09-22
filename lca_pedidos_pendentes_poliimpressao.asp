<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/impressao_pedido_PDF.asp" -->
<%
Function in_array(element, arr)
  in_array = False
  For i=0 To Ubound(arr)
     If Trim(arr(i)) = Trim(element) Then
        in_array = True
        Exit Function      
     End If
  Next
End Function

Server.ScriptTimeout = 3600

AbrirConexao()
Dim COD_PEDIDO_RS, COD_PEDIDO_QS, QTD_IMPRESSOES, QTD_LINHAS, IMPRESSAO_A_OU_B
Dim QTD_PDFS, INDIVIDUAL
QTD_PDFS       = 0
QTD_IMPRESSOES = 0
INDIVIDUAL     = 0

IMPRESSAO_A_OU_B = "A"


'////////////////////////////////////////////////////////////////////////
'Verifica os pedidos a serem impressos e grava na base temporária
count_i = 0
Set rs = objConexao.Execute("SELECT COD_PEDIDO, PEDIDO_EXTRA FROM TB_PEDIDOS WHERE FINALIZA_CORRECAO = 1 AND DATA_DESCEU_DIRETORIA IS NULL")
		
		While not rs.Eof
		
			COD_PEDIDO_RS = rs("COD_PEDIDO")
			COD_PEDIDO_QS = Request.Form("PEDIDO_"&COD_PEDIDO_RS)

			If int(COD_PEDIDO_QS) = 1 Then
				
				QTD_IMPRESSOES = QTD_IMPRESSOES + 1
				objConexao.Execute("INSERT INTO TB_PEDIDOS_IMPRESSAO_MASSA (COD_USUARIO,COD_PEDIDO,PEDIDO_EXTRA)VALUES ('"&Session("COD_USUARIO")&"', '"&COD_PEDIDO_RS&"','"&rs("PEDIDO_EXTRA")&"')")
				
			End If
			count_i = count_i + 1
		rs.MoveNext()
		Wend

rs.Close
Set rs = Nothing

Dim qtdImpressoes, metadeImpressa, contaImpresso, metade, metadeB, metadeA, faltaImprimir

qtdImpressoes = 0
contaImpresso = 0
metadeB = 0
metadeA = 0
faltaImprimir = 0

'///////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////
Response.Write("Qtd Impressões: "&QTD_IMPRESSOES)
Response.Write("<br />")
Response.Write("<br />")

sql = 	" SELECT TB_PEDIDOS_IMPRESSAO_MASSA.COD_PEDIDO, VW_PEDIDOS_COMPLETO.NOME_FORMANDO, TB_VENDEDORES.NOME_VENDEDOR "&_
		" FROM TB_PEDIDOS_IMPRESSAO_MASSA "&_
		" LEFT JOIN VW_PEDIDOS_COMPLETO ON TB_PEDIDOS_IMPRESSAO_MASSA.COD_PEDIDO = VW_PEDIDOS_COMPLETO.COD_PEDIDO "&_
		" INNER JOIN TB_PEDIDOS ON VW_PEDIDOS_COMPLETO.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO "&_
		" INNER JOIN TB_VENDEDORES ON TB_PEDIDOS.COD_VENDEDOR = TB_VENDEDORES.COD_VENDEDOR "&_
		" WHERE TB_PEDIDOS_IMPRESSAO_MASSA.COD_USUARIO="&Session("COD_USUARIO")&_
		" GROUP BY VW_PEDIDOS_COMPLETO.COD_PEDIDO ORDER BY TB_VENDEDORES.NOME_VENDEDOR, VW_PEDIDOS_COMPLETO.NOME_FORMANDO"	

Set rs = objConexao.Execute(sql)

If not rs.Eof Then
		
		
	If QTD_IMPRESSOES > 1 Then
	'Significa mais que 1 impressão, por tanto utilizará pelo menos 1 folha dupla
	QTD_LINHAS 	= int(QTD_IMPRESSOES/2)
	INDIVIDUAL	= QTD_IMPRESSOES Mod 2
	
	
	'1º Parte Cria os arquivos vazios e da nome as futuros arquivos
	'//////////////////////////////////////////////////////////////////////////////////////////////////
	
		Dim COD_ALEATORIO
		COD_ALEATORIO			= Now()
		COD_ALEATORIO			= Replace(COD_ALEATORIO,":","")
		COD_ALEATORIO			= Replace(COD_ALEATORIO,",","")
		COD_ALEATORIO			= Replace(COD_ALEATORIO,"/","")
		COD_ALEATORIO			= Replace(COD_ALEATORIO," ","")
		NOME_ARQUIVO_MASTER 	= "master_"&COD_ALEATORIO&"_"&Session("COD_USUARIO")
		NOME_ARQUIVO_INDIVIDUAL = "individual_"&COD_ALEATORIO&"_"&Session("COD_USUARIO")
		NOME_ARQUIVO_HIPER		= "hiper_"&COD_ALEATORIO&"_"&Session("COD_USUARIO")

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
		
		'Variáveis
		'/////
		Dim COD_PEDIDO, DATA_VENDA, NOME_VENDEDOR, REPASSE, VENDEDOR_RECEBE_EM, PEDIDO_EXTRA
		Dim TOTAL_EM_PRODUTOS, TOTAL_RECEBIDO, TOTAL_CARTAO, TOTAL_CHEQUE, TOTAL_DINHEIRO, TOTAL_CREDITO
		
		
	'//////////////////////////////////////////////////////////////////////////////////////////////////	
	For i=0 to QTD_LINHAS-1	
	'Abri o PDF	
		'////////////////////////
		Set PDF = Server.CreateObject("Persits.PDF")

		' Open an existing document, form W-9
		Set Doc = PDF.OpenDocument("e:\home\stillos\web\gerencial_foto\documentos\impressao_pedidos_duplo_v3.pdf")

		' Obtain Page object for page 1
		Set Page = Doc.Pages(1)


		' Create font object
		Set Font = Doc.Fonts("Arial")
		Set Param = PDF.CreateParam("html=true")
		
		'////////////////////////
		
			For ii=0 to 1
				'2 colunas
			
				If 	QTD_IMPRESSOES => 1 Then
				'É duplo
				
					If IMPRESSAO_A_OU_B = "A" Then
					'Impressão topo					
						IMPRESSAO_A_OU_B = "B"					
						
						'1º FORMANDO!!!
						'//////////////////////////////////////////////////////////////////////////////////////////////////////
						call impressaoPedidoPDF(rs("COD_PEDIDO"),"A")			
					
					Elseif IMPRESSAO_A_OU_B = "B" Then
					
						IMPRESSAO_A_OU_B = "A"
						
						'2º FORMANDO!!!
						'//////////////////////////////////////////////////////////////////////////////////////////////////////
						call impressaoPedidoPDF(rs("COD_PEDIDO"),"B")
					
					End If
								 
				End If
						
				QTD_IMPRESSOES = QTD_IMPRESSOES - 1
					
			rs.MoveNext()	
			Next
			
		'Fecha o arquivo PDF duplo	
		'////////////////////////
		FileName = Doc.Save("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_INDIVIDUAL&".pdf",true)
		Set Page = Nothing
		Set Doc = Nothing
		Set Pdf = Nothing
		
		Dim DocMaster, DocIndividual
		Set PDF = Server.CreateObject("Persits.PDF")
	
		'Abri o master e o acabado de ser criado
		Set DocMaster 		= PDF.OpenDocument("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf")
		Set DocIndividual 	= PDF.OpenDocument("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_INDIVIDUAL&".pdf")
		Set Page = DocIndividual.Pages(1)
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
		set fso = CreateObject("Scripting.FileSystemObject") 
		if fso.FileExists("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf") then
		fso.DeleteFile("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf")
		end if
		Set fso = Nothing
		
		'Renomeia o Hiper para Master
		set fso = CreateObject("Scripting.FileSystemObject") 
		fso.MoveFile "e:\home\stillos\web\gerencial_foto\documentos\hiper.pdf", "e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf"
		set fso = Nothing
		
		'/////////////////////////
		If QTD_IMPRESSOES = 1 and INDIVIDUAL = 1 Then
											
				Set PDF = Server.CreateObject("Persits.PDF")

				' Open an existing document, form W-9
				Set Doc = PDF.OpenDocument("e:\home\stillos\web\gerencial_foto\documentos\impressao_pedidos_individual_v3.pdf")
		
				' Obtain Page object for page 1
				Set Page = Doc.Pages(1)
		
				' Create font object
				Set Font = Doc.Fonts("Arial")
				Set Param = PDF.CreateParam("html=true")
				
				'1º FORMANDO REMANECENTE!!!
				'//////////////////////////////////////////////////////////////////////////////////////////////////////
				call impressaoPedidoPDF(rs("COD_PEDIDO"),"A")	

				'//////////////////////////////////////////////////////////////////////////////////////////////////////
				
				'Fecha o arquivo PDF duplo	
				'////////////////////////
				FileName = Doc.Save("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_INDIVIDUAL&".pdf",true)
				Set Page = Nothing
				Set Doc = Nothing
				Set Pdf = Nothing
				
				Set PDF = Server.CreateObject("Persits.PDF")
			
				'Abri o master e o acabado de ser criado
				Set DocMaster 		= PDF.OpenDocument("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf")
				Set DocIndividual 	= PDF.OpenDocument("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_INDIVIDUAL&".pdf")
				Set Page = DocIndividual.Pages(1)
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
				set fso = CreateObject("Scripting.FileSystemObject") 
				if fso.FileExists("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf") then
				fso.DeleteFile("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf")
				end if
				Set fso = Nothing
				
				'Renomeia o Hiper para Master
				set fso = CreateObject("Scripting.FileSystemObject") 
				fso.MoveFile "e:\home\stillos\web\gerencial_foto\documentos\hiper.pdf", "e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf"
				set fso = Nothing
				
				rs.MoveNext()
		End If
			
	Next
	
	'#######################################################################################
	
		'Retira a primeira página e exibe na tela
		Set Pdf = Server.CreateObject("Persits.Pdf") 
		Set Doc = Pdf.OpenDocument("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf")
		
	
		Doc.Pages.Remove 1
		Doc.SaveHttp "attachment;filename=pedido_multiplo_impressao.pdf"
				
		Set Doc = Nothing
		Set Pdf = Nothing
		'////////////////////////////
		
	
Else
'Significa que apesar de ter sido utilizado o metodo em massa só há uma impressão agendada
	Dim COD_PEDIDO_INDIVIDUAL, IMPRIMIR_SOMENTE_UMA
	IMPRIMIR_SOMENTE_UMA = 1
	Set rs2 = objConexao.Execute("SELECT COD_PEDIDO FROM TB_PEDIDOS_IMPRESSAO_MASSA WHERE COD_USUARIO="&Session("COD_USUARIO")&" LIMIT 1")
			
			If not rs2.Eof Then
				COD_PEDIDO_INDIVIDUAL = rs2("COD_PEDIDO")
			End If
			
	rs2.Close
	Set rs2 = Nothing

End If
		
		
		
End If

rs.Close
Set rs = Nothing
'////////////////////////////////////////////////////////////////////////
objConexao.Execute("DELETE FROM TB_PEDIDOS_IMPRESSAO_MASSA WHERE COD_USUARIO="&Session("COD_USUARIO"))


'Exclui arquivos remanecentes
set fso = CreateObject("Scripting.FileSystemObject") 
If fso.FileExists("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_INDIVIDUAL&".pdf") then
fso.DeleteFile("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_INDIVIDUAL&".pdf")
end if
		
If fso.FileExists("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf") then
fso.DeleteFile("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf")
end if
		
Set fso = Nothing

FecharConexao()

'Somente 1 etiqueta 
If IMPRIMIR_SOMENTE_UMA = 1 Then
	Response.Redirect("lca_pedidos_pendentes_uniimpressao.asp?cod_pedido="&COD_PEDIDO_INDIVIDUAL)
End If
%>