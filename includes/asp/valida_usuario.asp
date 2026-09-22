<%

VALIDA_PAG  = Request.ServerVariables("URL")
SCRIPT_NAME = Request.ServerVariables("SCRIPT_NAME")
COD_USUARIO = session("COD_USUARIO")
'session.Abandon()
if COD_USUARIO=empyt then
%>

<p id="Erro" class="login_erro"> Sua sessão expirou.<br> Para sua segurança será necessário realizar o login novamente para continuar utilizando o sistema.
<a href="index.asp" > Clique aqui para ir para a pagina de login</a>
</p>

<%
    response.end
end if


%>
