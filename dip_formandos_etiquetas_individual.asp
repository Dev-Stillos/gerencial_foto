<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<%Response.charset="utf-8"%>

<%
'Para tirar a validação da página basta tirar o ""#"" ou coloca-lo para habilitar antes do include
%>
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
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
   <form name="etiquetas" action="http://localhost/etiquetas/dip_producoes_fichas_etiquetas_localhost.asp?cod_prod=<%=Request.QueryString("cod_prod")%>" method="post">
   <textarea name="ETIQUETAS_PRODUCAO" cols="70" rows="30">
   <%
 AbrirConexao()
 'AbrirConexao2()
   sql = "SELECT * FROM TB_FICHAS_IDENTIFICACAO_ETIQUETAS_TEMPORARIA WHERE COD_USUARIO="&Session("COD_USUARIO")&" AND AGENDADA=1  AND COD_FORMANDO = "&Request.QueryString("cod_formando")&" ORDER BY LINHA2 DESC, LINHA1 DESC"
   
   Set rsEtiquetasLinhas = objConexao.Execute(sql,qtdEtiquetasImpressas)
   		
		qtdLinhasEtiquetas = int(qtdEtiquetasImpressas / 2) + qtdEtiquetasImpressas Mod 2
		
   		'Linhas
		For t=0 to qtdLinhasEtiquetas-1
			'Cabeçalho
			Response.Write chr(10)
			Response.Write chr(10)
			Response.Write chr(10)
			'FORMA ANTIGA 50 X 27 mm
			Response.Write "Q216,019"&chr(10)
			
			'FORMA NOVA 50 X 30 mm
			'Response.Write "Q240,024"&chr(10)
			Response.Write "q832"&chr(10)
			Response.Write "rN"&chr(10)
			Response.Write "S4"&chr(10)
			Response.Write "D7"&chr(10)
			Response.Write "ZT"&chr(10)
			Response.Write "JB"&chr(10)
			Response.Write "OD"&chr(10)
			Response.Write "R16,0"&chr(10)
			Response.Write "N"&chr(10)		
				
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
						'	Response.Write "B775,120,2,1,4,12,61,B,"&rsEtiquetasLinhas("LINHA1")&chr(10)
						'	Response.Write "A775,139,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA2")&chr(10)
						'	Response.Write "A775,200,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA3")&chr(10)
						'	Response.Write "A775,180,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA4")&chr(10)
						'	Response.Write "A775,160,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA5")&chr(10)
						'	Response.Write "A591,48,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA6")&chr(10)
							
							

						
						'PADRÃO 50 X 27 mm
							Response.Write "B352,94,2,1,4,12,61,B,"&rsEtiquetasLinhas("LINHA1")&chr(10)
							Response.Write "A352,114,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA2")&chr(10)
							Response.Write "A352,175,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA3")&chr(10)
							Response.Write "A352,155,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA4")&chr(10)
							Response.Write "A352,135,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA5")&chr(10)
							Response.Write "A176,24,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA6")&chr(10)
						Else
						
						'PADRÃO 50 X 30 mm
						'	Response.Write "B359,120,2,1,4,12,61,B,"&rsEtiquetasLinhas("LINHA1")&chr(10)
						'	Response.Write "A359,139,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA2")&chr(10)
						'	Response.Write "A359,200,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA3")&chr(10)
						'	Response.Write "A359,180,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA4")&chr(10)
						'	Response.Write "A359,160,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA5")&chr(10)
						'	Response.Write "A176,48,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA6")&chr(10)
									
						'PADRÃO 50 X 27 mm
							Response.Write "B779,94,2,1,4,12,61,B,"&rsEtiquetasLinhas("LINHA1")&chr(10)
							Response.Write "A779,114,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA2")&chr(10)
							Response.Write "A779,175,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA3")&chr(10)
							Response.Write "A779,155,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA4")&chr(10)	
							Response.Write "A779,135,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA5")&chr(10)
							Response.Write "A575,24,2,3,1,1,N,"&rsEtiquetasLinhas("LINHA6")&chr(10)
						End If
					
						rsEtiquetasLinhas.MoveNext()
					
					End If
					
				Next
				
			
			'Rodapé
			Response.Write "P1"&chr(10)
			Response.Write chr(10)
			Response.Write chr(10)
			Response.Write chr(10)
		Next
		
   

   rsEtiquetasLinhas.Close
   Set rsEtiquetasLinhas = Nothing
   'Limpa a temporária novamente
  
'FecharConexao2()

objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_ETIQUETAS_TEMPORARIA WHERE COD_FORMANDO="&Request.QueryString("cod_formando")&" AND AGENDADA=1")	
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
</body>
</html>