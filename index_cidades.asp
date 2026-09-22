<%@ Language=VBScript Codepage=65001 %>
<%



if request.querystring("u") = 1 then
	url = "http://cidades.ibge.gov.br/xtras/uf.php?lang=&coduf=17"&request.querystring("p")
else
	url = "http://cidades.ibge.gov.br/xtras/perfil.php?lang=&codmun="&request.querystring("p")
end if


response.Write(request.querystring("u")  & request.querystring("p") &"<hr>")

Dim xmlhttp
Set xmlhttp = CreateObject("Msxml2.ServerXMLHTTP")
xmlhttp.setTimeouts 5000, 5000, 5000, 5000
xmlhttp.open "GET", url, False
xmlhttp.send ""
Dim pagina

response.AddHeader "Content-Type", "text/html;charset=UTF-8"
Response.CodePage = 65001
Response.Charset = "UTF-8"

pagina = xmlhttp.responseText
Response.Write pagina
Set xmlhttp = Nothing

%>