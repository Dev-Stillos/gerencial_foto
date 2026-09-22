<%

Set d=Server.CreateObject("Scripting.Dictionary")
d.Add "renovacao","temp"

response.write(d.item("renovacao"))

response.write("<br>")
response.write("<br>")

d.item("renovacao") = "temp222"

response.write(d.item("renovacao"))

%>