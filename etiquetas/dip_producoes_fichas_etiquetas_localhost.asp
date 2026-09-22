<%Response.charset="utf-8"%>

<%
'Para tirar a validação da página basta tirar o ""#"" ou coloca-lo para habilitar antes do include
%>
<!-- include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>

<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
<script language="JavaScript"> 
function carregaArquivo(){
	win = location.href ="IMPRIMIR_ETIQUETAS.bat";
} 
</script> 
<style type="text/css">
<!--
body,td,th {
	font-family: Trebuchet MS, Tahoma, Arial;
}
-->
</style></head>

<body onLoad="carregaArquivo()"> 
 
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px "#"999999">

<h2 align="center"><a href="IMPRIMIR_ETIQUETAS.bat">Caso o arquivo não abrir automaticamente Clique aqui para imprimir</a></h2>
<% 	
Dim ETIQUETAS_PRODUCAO, linhas_grava
ETIQUETAS_PRODUCAO = Request.Form("ETIQUETAS_PRODUCAO")
'If ETIQUETAS_PRODUCAO <> Empty Then

'Instância o objeto FSO. 
Set Fso = CreateObject("Scripting.FileSystemObject") 

'Utilizando a propriedade CreateTextFile para criar o arquivo. 
Set Cria = Fso.CreateTextFile("c:\etiquetas\IMPRESSAO_ARQUIVO.txt",True) 
		
		linhas_grava = Split(ETIQUETAS_PRODUCAO,chr(10))
   
   		'Linhas
		For i=0 to UBound(linhas_grava)
			Cria.WriteLine linhas_grava(i)
		Next
		

'Fecha o arquivo criado. 
Cria.Close 

'Destruíndo objetos 
Set Fso = nothing 
Set Cria = nothing 

'End If
%>

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