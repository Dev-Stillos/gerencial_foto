<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body>
<%
'1) Descobri o número de IP real da placa que está sendo utilizada
'2) Limpa e deixa somente o número de IP
'3) Com este número de IP localiza o MAC Adress
'4) Limpa o mesmo
'5) Retorna com o nome da máquina e número de MAC Adress

Dim IP_USER, MAC_ADRESS, NOME_COMPUTADOR
Dim comando, cod_usuario
comando = Request.QueryString("comando")
cod_usuario = Request.QueryString("cod_usuario")
Dim wshell, wshell2, intReturn, intReturn2 

set wshell2 = server.createobject("wscript.shell") 
intReturn2 = wshell2.run("c:/etiquetas/ip.bat",0,true) 
'Response.Write intReturn
set wshell2 = nothing

set wshell = server.createobject("wscript.shell") 
intReturn = wshell.run("c:/etiquetas/mac.bat",0,true) 
'Response.Write intReturn
set wshell = nothing

%>
<% 
Const forReading = 1, forWriting = 2, forAppending = 8 
Set Fso2 = CreateObject("Scripting.FileSystemObject") 
Set ObjFile = Fso2.GetFile("c:/etiquetas/ip.txt")
Set objStream = ObjFile.OpenAsTextStream(forReading,false) 
Dim linha_ip_txt_cont, posicaoDoisPontos2
linha_ip_txt_cont = 0
While Not ObjStream.AtEndOfStream 
	linha_ip_txt_cont = linha_ip_txt_cont + 1
	Dim LINHA_IP_TXT, posicaoEspaco
	LINHA_IP_TXT = ObjStream.Readline

	If linha_ip_txt_cont = 2 Then
			posicaoEspaco = inStr(LINHA_IP_TXT,"-")
			LINHA_IP_TXT = Left(LINHA_IP_TXT,posicaoEspaco)
			LINHA_IP_TXT = Replace(LINHA_IP_TXT," ","")
			LINHA_IP_TXT = Replace(LINHA_IP_TXT,"-","")			
			LINHA_IP_TXT = Replace(LINHA_IP_TXT,chr(13),"")	
			LINHA_IP_TXT = Replace(LINHA_IP_TXT,chr(10),"")	
			posicaoDoisPontos2 = inStr(LINHA_IP_TXT,":")
			LINHA_IP_TXT = Mid(LINHA_IP_TXT,posicaoDoisPontos2+1,Len(LINHA_IP_TXT))
			LINHA_IP_TXT = Replace(LINHA_IP_TXT," ","")
			IP_USER = LINHA_IP_TXT
	End If			
	'Response.Write(LINHA_IP_TXT&"<BR />")

Wend  
objStream.Close 

'Destruíndo objetos 
Set Fso2 = nothing 
Set objStream = nothing 


'O parâmetro forReading é utilizado somente para leitura. 


'Instância o objeto FSO. 
Set Fso = CreateObject("Scripting.FileSystemObject") 

'O método GetFile retorna o arquivo existente. 
Set ObjFile = Fso.GetFile("c:/etiquetas/mac.txt") 

'O método OpenAsTextStream abre o arquivo como sendo um fluxo de texto. 
Set objStream = ObjFile.OpenAsTextStream(forReading,false) 

'Looping para varrer todas as linha do arquivo. 
Dim posicaoDoisPontos, CONTEUDO_LINHA, LINHA_IP, linha_contador
linha_contador = 0
While Not ObjStream.AtEndOfStream 


CONTEUDO_LINHA = ObjStream.Readline 
'O método Readline lê uma linha de um arquivo 
'Response.Write "Antes: "&CONTEUDO_LINHA
'Limpar conteúdo

posicaoDoisPontos = inStr(CONTEUDO_LINHA,":")
'CONTEUDO_LINHA = ObjStream.Readline
'Response.Write(posicaoDoisPontos)
If posicaoDoisPontos > 0 Then
linha_contador = linha_contador+1
CONTEUDO_LINHA = Trim(Mid(CONTEUDO_LINHA,posicaoDoisPontos+2,Len(CONTEUDO_LINHA)))
CONTEUDO_LINHA = Replace(CONTEUDO_LINHA," ","")
CONTEUDO_LINHA = Replace(CONTEUDO_LINHA,chr(13),"")
CONTEUDO_LINHA = Replace(CONTEUDO_LINHA,chr(10),"")

'Response.Write(CONTEUDO_LINHA)
	'Após limpar busca a linha que possui o IP
	'Response.Write("Linha:"&CONTEUDO_LINHA&"<br />")
	If linha_contador = 1 Then
		NOME_COMPUTADOR = CONTEUDO_LINHA
	End If
	'Response.Write("IP IPCONFIG: "&CONTEUDO_LINHA&" IP BUSCADO: "&IP_USER&"<br />")
	If CONTEUDO_LINHA = IP_USER Then
		LINHA_IP = linha_contador
	End If
End If
'CONTEUDO_LINHA = Mid(CONTEUDO_LINHA,posicaoDoisPontos,200)
'Response.Write("<br />")
'Response.Write("Depois:"&CONTEUDO_LINHA&"<br />")

Wend 
Response.Write("Linha ="&LINHA_IP)

'Fecha o arquivo. 
objStream.Close 

'Destruíndo objetos 
Set Fso = nothing 
Set objStream = nothing 
%> 

<%
'//////RODANOVAMENTE ENTRETANTO VAI NA LINHA CORRETA
'O parâmetro forReading é utilizado somente para leitura. 

'Instância o objeto FSO. 
Set Fso = CreateObject("Scripting.FileSystemObject") 

'O método GetFile retorna o arquivo existente. 
Set ObjFile = Fso.GetFile("c:/etiquetas/mac.txt") 

'O método OpenAsTextStream abre o arquivo como sendo um fluxo de texto. 
Set objStream = ObjFile.OpenAsTextStream(forReading,false) 

'Looping para varrer todas as linha do arquivo. 
linha_contador = 0
While Not ObjStream.AtEndOfStream 


CONTEUDO_LINHA = ObjStream.Readline 
'O método Readline lê uma linha de um arquivo 
'Response.Write "Antes: "&CONTEUDO_LINHA
'Limpar conteúdo

posicaoDoisPontos = inStr(CONTEUDO_LINHA,":")
'CONTEUDO_LINHA = ObjStream.Readline
'Response.Write(posicaoDoisPontos)
	If posicaoDoisPontos > 0 Then
	linha_contador = linha_contador+1
	
	
	CONTEUDO_LINHA = Trim(Mid(CONTEUDO_LINHA,posicaoDoisPontos+2,Len(CONTEUDO_LINHA)))
	CONTEUDO_LINHA = Replace(CONTEUDO_LINHA," ","")
	CONTEUDO_LINHA = Replace(CONTEUDO_LINHA,chr(13),"")
	CONTEUDO_LINHA = Replace(CONTEUDO_LINHA,chr(10),"")
	'Response.Write(CONTEUDO_LINHA)
	
	
	If linha_contador = LINHA_IP-3  Then
		MAC_ADRESS = CONTEUDO_LINHA
	End If
End If
'CONTEUDO_LINHA = Mid(CONTEUDO_LINHA,posicaoDoisPontos,200)
'Response.Write("<br />")
'Response.Write("Depois:"&CONTEUDO_LINHA&"<br />")

Wend 

'Fecha o arquivo. 
objStream.Close 

'Destruíndo objetos 
Set Fso = nothing 
Set objStream = nothing 

If comando = "incluir" Then
	Response.Redirect("https://stillos.com.br/gerencial_foto/adm_usuarios_pcs.asp?cod_usuario="&cod_usuario&"&DESCRICAO_PC="&NOME_COMPUTADOR&"&CHAVE_VALIDACAO="&MAC_ADRESS)
End If

If comando = "validar" Then
	Response.Redirect("https://stillos.com.br/gerencial_foto/index.asp?&DESCRICAO_PC="&NOME_COMPUTADOR&"&CHAVE_VALIDACAO="&MAC_ADRESS)
End If
%> 

Nome do PC: <%=NOME_COMPUTADOR%><br />
Este é o seu MAC! <%=MAC_ADRESS%>
</body>
</html>
