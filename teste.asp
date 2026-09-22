<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<%

dim a 
a = getConfiguracoes(false, array("periodo_renovacao_padrao"))



response.write("item" & a("periodo_renovacao_padrao"))

Response.Write("<p>Key values:</p>")

b=a.Keys
for i=0 to a.Count-1
  Response.Write(b(i))
  Response.Write("<br>")
next
%>