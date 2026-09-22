<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<%Response.charset="utf-8"%>
 
<%
'Para tirar a validação da página basta tirar o ""#"" ou coloca-lo para habilitar antes do include
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>

<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
<script language="javascript">
function carregaEtiquetas(){
	document.etiquetas.submit();
}
</script>
</head>

<body onload="carregaEtiquetas()">
 
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px "#"999999">
	  <%
	  Dim QTD_AGREGADOS_TIPOS(20)
	  Dim rs9, PRODUCAO_FINALIZADA, DESCRICAO, VERIFICA_FINALIZACAO, CONTEUDO_TXT, LINHA1, LINHA2, LINHA3, LINHA4, LINHA5
	  
	  LINHA_COMPLETA_PREENCHIDA = "Sim"
	  
	  COLUNA1 = 0
	  COLUNA2 = 0
	  
	  VERIFICA_FINALIZACAO = 1

	  Dim imprimir, quantidade, i, zpl

imprimir = Request("imprimir")
quantidade = Request("quantidade")

If imprimir = "sim" Then
    Response.ContentType = "text/plain"
    Response.AddHeader "Content-Disposition", "inline; filename=etiquetas.zpl"

    zpl = ""
    For i = 1 To CInt(quantidade)
        zpl = zpl & "^XA" & vbCrLf
        zpl = zpl & "^FO50,50^ADN,36,20^FDPORTAL STILLO'S - Etiqueta " & i & "^FS" & vbCrLf
        zpl = zpl & "^FO50,100^BY2" & vbCrLf
        zpl = zpl & "^BCN,100,Y,N,N" & vbCrLf
        zpl = zpl & "^FD1234567890" & i & "^FS" & vbCrLf
        zpl = zpl & "^XZ" & vbCrLf
    Next

    Response.Write zpl
    Response.End
End If
	  
	  AbrirConexao()
	  'AbrirConexao2()
		
	  'Limpa a base temporária
	  objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_ETIQUETAS_TEMPORARIA WHERE COD_USUARIO="&Session("COD_USUARIO")&" AND AGENDADA=2")
	  
	  Dim rsQ
			  Set rs9 = objConexao.Execute("SELECT PRODUCAO_FINALIZADA, DESCRICAO FROM TB_PRODUCAO WHERE COD_PRODUCAO="&Request.QueryString("cod_prod"))

					If not rs9.Eof then
					
						PRODUCAO_FINALIZADA = rs9("PRODUCAO_FINALIZADA")
						DESCRICAO = rs9("DESCRICAO") 
					End If					

			  rs9.Close()
	  		  Set rs9 = Nothing 
	 

	 If PRODUCAO_FINALIZADA  = 1 then
	   
	      Dim rs, rs2, rs3, rs18
		  Dim QTD_ALBUNS, FICHAS_VENDA, QTD_AGREGADOS, GASTO_DE_ETIQUETAS, COLUNAS, COD_CONTRATO
		  GASTO_DE_ETIQUETAS = 0

		  <!--ETIQUETAS SELECIONADAS OU NÃO-->

' --- Recebe parâmetros
cod_prod  = Request.QueryString("cod_prod")
formandos = Trim(Request.QueryString("formandos"))

sql = "SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO=" & cod_prod

If Len(formandos) > 0 Then
    arr = Split(formandos, ",")
    sql = sql & " AND COD_FORMANDO IN ("
    For i = 0 To UBound(arr)
        item = Trim(arr(i))
        item = Replace(item, "'", "")
        item = Replace(item, """", "")
        sql = sql & item
        If i < UBound(arr) Then sql = sql & ","
    Next
    sql = sql & ")"
End If

sql = sql & " ORDER BY CONVERT(ALBUM,SIGNED)"

Set rs = objConexao.Execute(sql)




		  
		  If not rs.Eof then 
		  
		  Dim TURMA, COD_FORMANDO, NOME_CURSO, ALBUM, NOME_FORMANDO, ANO_CONCLUSAO, SEMESTRE_CONCLUSAO, ERRO, COD_SEM_DIGITO, PRODUTO_AGREGADO, ENCADERNACOES_EXTRAS, LOTE
		  ERRO = 2
			While not rs.Eof		
				
				
				COD_FORMANDO			= rs("COD_FORMANDO")
				NOME_CURSO				= Ucase(Left(rs("NOME_CURSO"),16))
				ALBUM					= rs("ALBUM")
				ANO_CONCLUSAO			= Right(rs("ANO_CONCLUSAO"),2)
				SEMESTRE_CONCLUSAO		= UCASE(rs("SEMESTRE_CONCLUSAO"))
				COD_SEM_DIGITO			= Mid(COD_FORMANDO,1,Len(COD_FORMANDO)-1)
				PRODUTO_AGREGADO		= rs("PRODUTO_AGREGADO")
				COD_CONTRATO			= rs("COD_CONTRATO")
				ENCADERNACOES_EXTRAS	= formatNumber(0&rs("ENCADERNACOES_EXTRAS"),0)
				LOTE					= formatNumber(0&rs("LOTE"),0)
				
				ALBUM = Replace(ALBUM,".","")
				ALBUM = Replace(ALBUM,",","")
				
				<!--BLOQUEIO DOS ALBUNS 1000-->
				
				
				'Nome da Turma
				'///////////////////////////////
				Dim rsTurmaForm
				
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
							For i=2002 to 2020
							
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
						Else
						'Caso não encontre coloca o nome da descrição da Produção, por precaução.
							TURMA = UCASE(Left(Request.QueryString("desc"),23))
						End If
				
				rsTurmaForm.Close
				Set rsTurmaForm = Nothing
				

				
				  
				
				'///////////////////////////////
				
				If Len (ALBUM) = 1 Then
					ALBUM = "000"&ALBUM 
				ElseIf Len(ALBUM) = 2 Then
					ALBUM = "00"&ALBUM 
				ElseIf Len(ALBUM) = 3 Then
					ALBUM = "0"&ALBUM 
				End If
				
				If Len(LOTE) = 1 Then
					LOTE = "0"&LOTE
				End If
				
				NOME_FORMANDO	= Left(Trim(rs("NOME_FORMANDO")),23)
					
					QTD_AGREGADOS_TIPOS(0) = "ALBUM"
					QTD_AGREGADOS_TIPOS(1) = "FICHA"	
				
				'Response.Write("Formando: "&COD_FORMANDO&" Agregado: "&PRODUTO_AGREGADO&"<br />")
				QTD_AGREGADOS = 0
				If PRODUTO_AGREGADO = 1 Then
					
					sql = " SELECT TB_PRODUTOS.NOME_PRODUTO, TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.TIPO, TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO FROM "&_
					" TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO WHERE "&_
					" TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO="&COD_FORMANDO&_
				    " ORDER BY TIPO"
					
					
					Set rs18 = objConexao.Execute(sql,QTD_AGREGADOS)
					'Response.Write("Agregados:"&QTD_AGREGADOS&"<br />")				
					QTD_AGREGADOS_TIPOS(2) = ""
					QTD_AGREGADOS_TIPOS(3) = ""
					QTD_AGREGADOS_TIPOS(4) = ""
					QTD_AGREGADOS_TIPOS(5) = ""
					QTD_AGREGADOS_TIPOS(6) = ""
					QTD_AGREGADOS_TIPOS(7) = ""
					QTD_AGREGADOS_TIPOS(8) = ""
					QTD_AGREGADOS_TIPOS(9) = ""
								
								While not rs18.Eof
										
										QTD_AGREGADOS_TIPOS(rs18("TIPO")) = rs18("NOME_PRODUTO")
										'Response.Write("Tipo ="&QTD_AGREGADOS_TIPOS(rs18("TIPO"))&"<br />")
								rs18.MoveNext()
								Wend
					
					rs18.Close
					Set rs18 = Nothing			
					
				End If
				
					
				'3 = 1 para álbum, 1 caixa do álbum, 1 ficha + AGREGADO(S)
				
				GASTO_DE_ETIQUETAS = 0
				GASTO_DE_ETIQUETAS = formatNumber(3 + QTD_AGREGADOS,0)
				
				
				
				
				Dim NUMERO_ETIQUETA, DESCRICAO_BREVE
				<!-- MUDAR O NUMERO_ETIQUETA AQUI, SENDO O FINAL 0 OU 1 -->
				NUMERO_ETIQUETA = 1
				GASTO_DE_ETIQUETAS = 1
				'Response.Write("Etiquetas:"&GASTO_DE_ETIQUETAS&"<br />")
				
				'/////////////////
				'Antes de fazer o "FOR" com a qtd de etiquetas verifica se é álbum com encadernação dupla, se for grava na base.
				
				If ENCADERNACOES_EXTRAS > 0 Then
									
									
									Dim LINHA_3_CORRIGIDA_CONTRATO, TIPO
									
									LINHA_3_CORRIGIDA_CONTRATO = Left(TURMA,16)&" - "&COD_CONTRATO
									'Response.Write(LINHA_3_CORRIGIDA_CONTRATO)
									'Response.Write("<br>")
									'Response.Write(TURMA)
									
									TIPO = 0
									
									LINHA1 = chr(34)&COD_FORMANDO&"1"&chr(34)
									LINHA2 = chr(34)&NOME_FORMANDO&chr(34)
									LINHA3 = chr(34)&LINHA_3_CORRIGIDA_CONTRATO&chr(34)
									LINHA4 = chr(34)&NOME_CURSO&" - "&SEMESTRE_CONCLUSAO&"/"&ANO_CONCLUSAO&chr(34)
									LINHA5 = chr(34)&"ALBUM:"&ALBUM&"A -"&ArrumaData(Date())&chr(34)
									LINHA6 = chr(34)&"ALBUM DUPLO"&chr(34)
									
									
									flagImprimir = 0
									If Len(formandos) > 0 Then
										' Se o formando atual estiver na lista recebida pela querystring, marca como 1
										If InStr(1, "," & formandos & ",", "," & CStr(COD_FORMANDO) & ",") > 0 Then
											flagImprimir = 1
										End If
									End If

									sql="INSERT INTO TB_FICHAS_IDENTIFICACAO_ETIQUETAS_TEMPORARIA ("&_
										"COD_USUARIO ,"&_
										"COD_FORMANDO ,"&_
										"LINHA1 ,"&_
										"LINHA2 ,"&_
										"LINHA3 ,"&_
										"LINHA4 ,"&_
										"LINHA5 ,"&_
										"LINHA6 ,"&_
										"ALBUM_INDICE ,"&_
										"TIPO ,"&_
										"IMPRIMIR_ETIQUETA"&_
										")"&_
										"VALUES ("&_
										Session("COD_USUARIO")&", "&COD_FORMANDO&", "&sqlTexto(LINHA1)&", "&sqlTexto(LINHA2)&", "&sqlTexto(LINHA3)&", "&sqlTexto(LINHA4)&", "&sqlTexto(LINHA5)&", "&sqlTexto(LINHA6)&", "&formatNumber(ALBUM,0)&", "&NUMERO_ETIQUETA&", "&flagImprimir&_
    									")"

								'2 Por serem 2 cópias    
								objConexao.Execute(sql)
								objConexao.Execute(sql)

				
				End If
				
				
				'/////////////////
				
				
				' --- Etiqueta 1 (FICHA - FINAL 1)
LINHA_3_CORRIGIDA_CONTRATO = Left(TURMA,16)&" - "&COD_CONTRATO
LINHA1 = chr(34)&COD_FORMANDO & "1" & chr(34)
LINHA2 = chr(34)&NOME_FORMANDO&chr(34)
LINHA3 = chr(34)&LINHA_3_CORRIGIDA_CONTRATO&chr(34)
LINHA4 = chr(34)&NOME_CURSO&" - "&SEMESTRE_CONCLUSAO&"/"&ANO_CONCLUSAO&chr(34)
LINHA5 = chr(34)&"ALBUM:"&ALBUM&" - "&ArrumaData(Date())&chr(34)
LINHA6 = chr(34)&"FICHA"&chr(34)

flagImprimir = 0
If Len(formandos) > 0 Then
    If InStr(1, "," & formandos & ",", "," & CStr(COD_FORMANDO) & "-1,") > 0 Then
        flagImprimir = 1
    End If
End If

sql="INSERT INTO TB_FICHAS_IDENTIFICACAO_ETIQUETAS_TEMPORARIA ("&_
    "COD_USUARIO ,COD_FORMANDO ,LINHA1 ,LINHA2 ,LINHA3 ,LINHA4 ,LINHA5 ,LINHA6 ,ALBUM_INDICE ,TIPO ,IMPRIMIR_ETIQUETA)"&_
    " VALUES ("&Session("COD_USUARIO")&", "&COD_FORMANDO&", "&sqlTexto(LINHA1)&", "&sqlTexto(LINHA2)&", "&sqlTexto(LINHA3)&", "&sqlTexto(LINHA4)&", "&sqlTexto(LINHA5)&", "&sqlTexto(LINHA6)&", "&formatNumber(ALBUM,0)&", 1, "&flagImprimir&")"
objConexao.Execute(sql)


' === Controle do segundo tipo de etiqueta ===
Dim tipoEtiqueta, tipoSolicitado
tipoSolicitado = LCase(Trim(Request("tipo")))

Select Case tipoSolicitado
    Case "ficha_album"
        tipoEtiqueta = 0
        LINHA1 = chr(34)&COD_FORMANDO & "0" & chr(34)
        LINHA6 = chr(34)&"ÁLBUM"&chr(34)

    Case "ficha_digital"
        tipoEtiqueta = 2
        LINHA1 = chr(34)&COD_FORMANDO & "2" & chr(34)
        LINHA6 = chr(34)&"DIGITAL"&chr(34)

    Case Else
        tipoEtiqueta = -1  ' não gera segunda etiqueta
End Select


If tipoEtiqueta <> -1 Then
    flagImprimir = 0
    If Len(formandos) > 0 Then
        If InStr(1, "," & formandos & ",", "," & CStr(COD_FORMANDO) & "-" & tipoEtiqueta & ",") > 0 Then
            flagImprimir = 1
        End If
    End If

    sql="INSERT INTO TB_FICHAS_IDENTIFICACAO_ETIQUETAS_TEMPORARIA ("&_
        "COD_USUARIO ,COD_FORMANDO ,LINHA1 ,LINHA2 ,LINHA3 ,LINHA4 ,LINHA5 ,LINHA6 ,ALBUM_INDICE ,TIPO ,IMPRIMIR_ETIQUETA)"&_
        " VALUES ("&Session("COD_USUARIO")&", "&COD_FORMANDO&", "&sqlTexto(LINHA1)&", "&sqlTexto(LINHA2)&", "&sqlTexto(LINHA3)&", "&sqlTexto(LINHA4)&", "&sqlTexto(LINHA5)&", "&sqlTexto(LINHA6)&", "&formatNumber(ALBUM,0)&", "&tipoEtiqueta&", "&flagImprimir&")"
    objConexao.Execute(sql)
End If




				
				

			
				
				'//////////////////////////////////////////////////
							DESCRICAO_BREVE = ""

							 ' <-- fecha o IF do bloqueio do álbum 1000
			rs.MoveNext()
			Wend
		  
		  Else
		  
		  Response.Write("Não há formandos para impressão de etiquetas!")
		  
		  End If 


	 Else
	 	
		Response.Write("Produção não finalizada, não poderá imprimir etiquetas!")
		
	End If 

   %>
   <!-- <form name="etiquetas" action="http://localhost/etiquetas/dip_producoes_fichas_etiquetas_localhost.asp?cod_prod=<%=Request.QueryString("cod_prod")%>" method="post"> -->
   <!-- <textarea name="ETIQUETAS_PRODUCAO" cols="70" rows="30"> -->
   <%
Dim codUsuario
codUsuario = 0
If IsNumeric(Session("COD_USUARIO")) Then
    codUsuario = CLng(Session("COD_USUARIO"))
End If

sql = "SELECT COD_USUARIO,LINHA1,LINHA2,LINHA3,LINHA4,LINHA5,LINHA6, COD_FORMANDO,AGENDADA,ALBUM_INDICE,TIPO " & _
      "FROM TB_FICHAS_IDENTIFICACAO_ETIQUETAS_TEMPORARIA " & _
      "WHERE COD_USUARIO=" & codUsuario & " AND AGENDADA=2"

' Se houver algum marcado, imprime só os marcados
Set rsCheck = objConexao.Execute("SELECT COUNT(*) AS QTD FROM TB_FICHAS_IDENTIFICACAO_ETIQUETAS_TEMPORARIA WHERE COD_USUARIO=" & codUsuario & " AND AGENDADA=2 AND IMPRIMIR_ETIQUETA=1")

Dim qtdMarcados
qtdMarcados = 0
If Not rsCheck.Eof Then
    qtdMarcados = CLng(rsCheck("QTD") & "")
End If

If qtdMarcados > 0 Then
    sql = sql & " AND IMPRIMIR_ETIQUETA=1"
End If

rsCheck.Close
Set rsCheck = Nothing

sql = sql & " GROUP BY COD_USUARIO,LINHA1,LINHA2,LINHA3,LINHA4,LINHA5,LINHA6,COD_FORMANDO,AGENDADA,ALBUM_INDICE,TIPO ORDER BY TIPO,ALBUM_INDICE ASC"




   
   Set rsEtiquetasLinhas = objConexao.Execute(sql,qtdEtiquetasImpressas)
   		
		qtdLinhasEtiquetas = int(qtdEtiquetasImpressas / 2) + qtdEtiquetasImpressas Mod 2
		
   		'Linhas
		For t=0 to qtdLinhasEtiquetas-1
			'Cabeçalho
			Response.Write "@quebra_de_linha"
			Response.Write "@quebra_de_linha"
			Response.Write "@quebra_de_linha"
			Response.Write "Q240,024"&"@quebra_de_linha"
			Response.Write "q832"&"@quebra_de_linha"
			Response.Write "rN"&"@quebra_de_linha"
			Response.Write "S4"&"@quebra_de_linha"
			Response.Write "D7"&"@quebra_de_linha"
			Response.Write "ZT"&"@quebra_de_linha"
			Response.Write "JB"&"@quebra_de_linha"
			Response.Write "OD"&"@quebra_de_linha"
			Response.Write "R16,0"&"@quebra_de_linha"
			Response.Write "N"&"@quebra_de_linha"		
				
				'Colunas
				For tt=0 to 1
					qtdEtiquetasImpressas = qtdEtiquetasImpressas - 1
					
					If qtdEtiquetasImpressas => 0 Then
					'Corrige erro para somente uma etiqueta
						If tt = 0 Then
						
						'LINHA1 CÓDIGO FORMANDO
						'LINHA2 NOME FORMANDO
						'LINHA3 NOME TURMA - SEMESTRE CONCLUSÃO / ANO CONCLUSAO / CONTRATO
						'LINHA4 NOME CURSO
						'LINHA5 ALBUM E DATA IMPRESSA
						'LINHA6 DESCRICAO DO QUE FOI IMPRESSO				
						
						'PADRÃO 50 X 30 mm
							Response.Write "B775,120,2,1,4,12,61,B,"&rsEtiquetasLinhas("LINHA1")&"@quebra_de_linha"
							Response.Write "A775,139,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA2")&"@quebra_de_linha"
							Response.Write "A775,200,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA3")&"@quebra_de_linha"
							Response.Write "A775,180,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA4")&"@quebra_de_linha"
							Response.Write "A775,160,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA5")&"@quebra_de_linha"
							Response.Write "A591,48,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA6")&"@quebra_de_linha"
							
							

						
						'PADRÃO 50 X 27 mm
						'	Response.Write "B352,94,2,1,4,12,61,B,"&rsEtiquetasLinhas("LINHA1")&"@quebra_de_linha"
						'	Response.Write "A352,114,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA2")&"@quebra_de_linha"
						'	Response.Write "A352,175,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA3")&"@quebra_de_linha"
						'	Response.Write "A352,155,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA4")&"@quebra_de_linha"
						'	Response.Write "A352,135,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA5")&"@quebra_de_linha"
						'	Response.Write "A176,24,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA6")&"@quebra_de_linha"
						Else
						
						'PADRÃO 50 X 30 mm
							Response.Write "B359,120,2,1,4,12,61,B,"&rsEtiquetasLinhas("LINHA1")&"@quebra_de_linha"
							Response.Write "A359,139,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA2")&"@quebra_de_linha"
							Response.Write "A359,200,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA3")&"@quebra_de_linha"
							Response.Write "A359,180,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA4")&"@quebra_de_linha"
							Response.Write "A359,160,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA5")&"@quebra_de_linha"
							Response.Write "A176,48,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA6")&"@quebra_de_linha"
									
						'PADRÃO 50 X 27 mm
						'	Response.Write "B779,94,2,1,4,12,61,B,"&rsEtiquetasLinhas("LINHA1")&"@quebra_de_linha"
						'	Response.Write "A779,114,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA2")&"@quebra_de_linha"
						'	Response.Write "A779,175,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA3")&"@quebra_de_linha"
						'	Response.Write "A779,155,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA4")&"@quebra_de_linha"	
						'	Response.Write "A779,135,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA5")&"@quebra_de_linha"
						'	Response.Write "A575,24,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA6")&"@quebra_de_linha"
						End If
					
						rsEtiquetasLinhas.MoveNext()
					
					End If
					
				Next
				
			
			'Rodapé
			Response.Write "P1"&"@quebra_de_linha"
			Response.Write "@quebra_de_linha"
			Response.Write "@quebra_de_linha"
			Response.Write "@quebra_de_linha"
		Next
		
   

   rsEtiquetasLinhas.Close
   Set rsEtiquetasLinhas = Nothing
   'Limpa a temporária novamente
   objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_ETIQUETAS_TEMPORARIA WHERE COD_USUARIO="&Session("COD_USUARIO")&" AND AGENDADA=2")
'FecharConexao2()	
FecharConexao()
%>
</textarea>
</form>
		</td>
    </tr>
</table>  </td>
    </tr>
</table>  </td>
  </tr>
</table>
 </td>
</tr>
</table>
<script>
	const link = document.createElement("a");
	let content = document.body.innerText;
	const lines = content.split("@quebra_de_linha").join("\r\n"); // Quebra as linhas no caractere "@quebra_de_linha"
	const file = new Blob([lines], { type: 'text/plain' });
	link.href = URL.createObjectURL(file);
	link.download = "etiquetas.txt";
	link.click();
	URL.revokeObjectURL(link.href);
</script>

</body>
</html>