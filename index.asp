<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<%
 
if Request.ServerVariables("HTTPS") = "on"  then
	Response.Redirect("http://stillos.com.br/gerencial_foto")
end if 

sub atuIp()
     '
   'dim url, objSrvHTTP, ip
   
   'url = "http://ddns.stillos.com.br/stillosg.php?a="&now()
   
   'Set objSrvHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
   'objSrvHTTP.setTimeouts 5000, 5000, 5000, 5000
  ' objSrvHTTP.Open "GET", url, False
   'objSrvHTTP.Send()
   
  ' ip=objSrvHTTP.responseText
   
  ' Set objSrvHTTP = Nothing
  
  ' objConexao.execute("UPDATE TB_SISTEMAS_IPS SET ip_habilitado = '"&ip&"' WHERE descricao = 'Gr�fica (n�o alterar)'")

   'url = "http://ddns.stillos.com.br/stillos.php?a="&now()
   
   'Set objSrvHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
   'objSrvHTTP.setTimeouts 5000, 5000, 5000, 5000
   'objSrvHTTP.Open "GET", url, False
   'objSrvHTTP.Send()
   
   'ip=objSrvHTTP.responseText
   
   'Set objSrvHTTP = Nothing
  '179.153.202.98
   'objConexao.execute("UPDATE TB_SISTEMAS_IPS SET ip_habilitado = '"&ip&"' WHERE 	descricao = 'Stillos (n�o alterar)'")
   'objConexao.execute("UPDATE TB_SISTEMAS_IPS SET ip_habilitado = '187.38.53.48' WHERE 	descricao = 'Stillos (n�o alterar)'")
  '
end sub


Dim mensagem

If Request.Form("USUARIO")<> Empty Then
	
			AbrirConexao()
			Dim US, SE, sql, rs, ERRO_NAO_LOGAR, COD_ERRO, PERMITIR_USO_EXTERNO
			US = Request.Form("USUARIO")
			SE = Request.Form("SENHA")
			
			US = Replace(US,"'","''")
			US = Replace(US," ","")
			SE = Replace(SE,"'","''")
			SE = Replace(SE," ","")
					
			')Valida��es
			'1) Verifica nome de usu�rio e senha
			'2) Verifica se possui acesso externo, se n�o tiver verifica o IP da rede

            'Antes de verificar atualizamos o IP da grafica de acordo com o DNS
            call atuIp()
			
			sql = " SELECT TB_ADM_USUARIOS.COD_PERFIL,TB_ADM_USUARIOS.BLOQUEADO, TB_ADM_USUARIOS.PERMITIR_USO_EXTERNO, TB_ADM_USUARIOS.COD_USUARIO, TB_ADM_USUARIOS.NOME, TB_ADM_USUARIOS.LOGIN, TB_ADM_USUARIOS.SENHA,"&_
				"	TB_ADM_USUARIOS_DEPARTAMENTOS.NOME_DEPTO "&_
				" FROM TB_ADM_USUARIOS_DEPARTAMENTOS "&_
				" INNER JOIN TB_ADM_USUARIOS ON TB_ADM_USUARIOS_DEPARTAMENTOS.COD_DEPTO = TB_ADM_USUARIOS.COD_DEPTO"&_
				" WHERE LOGIN = '"&US&"' AND SENHA = '"&SE&"' AND TB_ADM_USUARIOS.BLOQUEADO=2"
				  
			Set rs = objConexao.execute(sql)
		
			If Not rs.eof Then
				Session("COD_PERFIL")	 = rs("COD_PERFIL")
				Session("COD_USUARIO")	 = rs("COD_USUARIO")
				Session("NOME")			 = rs("NOME")
				Session("DEPARTAMENTO")	 = rs("NOME_DEPTO")
				PERMITIR_USO_EXTERNO	 = rs("PERMITIR_USO_EXTERNO")

				'///////////////////////////
				'2)==>
				'===================== LIBERAR TODOS OS IPs TEMPORARIAMENTE =====================
' If int(PERMITIR_USO_EXTERNO) = 2 Then
'     'Verifica o IP                  
'     sql = "SELECT cod_sistema_ip FROM TB_SISTEMAS_IPS WHERE ip_habilitado = "&sqlTexto(Request.ServerVariables("REMOTE_ADDR"))
'     Response.Write(sql)
'     Set rsIp = objConexao.Execute(sql)

'     If rsIp.Eof Then
'         'Significa que não encontrou nenhum IP válido
'         ERRO_NAO_LOGAR = 2
'         COD_ERRO 	   = 2						
'     End If

'     rsIp.Close
'     Set rsIp = Nothing	
' End If
'===================== FIM LIBERAÇÃO DE IPs =====================================
	
				'///////////////////////////	
					
			Else
					'1)
					ERRO_NAO_LOGAR = 1
					COD_ERRO 	   = 1
						 
			End If
				

			rs.Close
			Set rs=Nothing 
			
			FecharConexao()	
			
			If ERRO_NAO_LOGAR = 1 Then
				Session.Abandon()
				Response.Redirect("index.asp?msg="& Server.URLencode("Senha e/ou usu&aacute;rio inv&aacute;lido(s)! Se não lembrar chame o T.I 1") )
			ElseIf ERRO_NAO_LOGAR = 2 Then
				Session.Abandon()
				Response.Redirect("index.asp?msg="& Server.URLencode("Aten&ccedil;&atilde;o seu login n&atilde;o permite acesso para este n&uacute;mero de IP<br />Se voc&ecirc; estiver usando a pr&oacute;pria rede da Stillo's entre em contato com o(a) respons&aacute;vel pelo sistema na empresa</strong>!"))
			Else
				'Caso contr�rio permite o acesso e grava o log
				'Log do usu�rio
				'Conta a quantidade de acessos e guarda o �ltimo acesso
				Dim rsLog, qtdAcessos, ACESSO_EM
					AbrirConexao()	
						Set rsLog = objConexao.Execute("SELECT * FROM TB_ADM_LOG_USUARIOS WHERE COD_USUARIO="&Session("COD_USUARIO")&" ORDER BY ACESSO_EM DESC",qtdAcessos)	
							
							If not rsLog.Eof then
							
								Session("ACESSO_EM") = rsLog("ACESSO_EM")
								Session("QTDACESSOS") = qtdAcessos+1
							End If
							
						rsLog.Close()
						Set rsLog = Nothing 
							
						'Insere um novo log
						objConexao.Execute("INSERT INTO TB_ADM_LOG_USUARIOS (ACESSO_EM, COD_USUARIO) VALUES (CURRENT_TIMESTAMP , '"&Session("COD_USUARIO")&"')")

					FecharConexao()
					Session("GerencialAvisoModernizacaoPendente") = True
					Response.Redirect("adm_index.asp")
			End If
End if	
	
%>


<html>
<head>
<meta name="robots" content="noindex, nofollow">
<title> Portal Stillo's - Login</title>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<style>
    body {
        font-family: Verdana, Arial, sans-serif;
        margin: 0;
        padding: 0;
        background-color: #10151d;
        color: white;
        display: flex;
        align-items: center;
        justify-content: center;
        height: 100vh;
    }

    .login-container {
        background-color: #1c1f2b;
        padding: 70px;
        border-radius: 30px;
        box-shadow: 0 0 20px rgba(0, 0, 0, 0.5);
        width: 350px;
        text-align: center;
    }

    .login-container img {
        width: 240px;
        margin-bottom: 20px;
    }

    .login-container h2 {
        margin-bottom: 25px;
        font-size: 20px;
    }

    .login-container input[type="text"],
    .login-container input[type="password"] {
        width: 100%;
        padding: 10px;
        margin-bottom: 15px;
        border: none;
        border-radius: 8px;
        background-color: #2d2f3b;
        color: #ffffff;
    }

    .login-container input[type="submit"] {
        width: 90%;
        padding: 8px;
        background-color: #ff007f;
        border: none;
        border-radius: 8px;
        color: white;
        font-weight: bold;
        cursor: pointer;
        margin-top: 10px;
    }

    .login-container p {
        font-size: 12px;
        margin-top: 15px;
        color: #ccc;
    }
</style>
<script language="javascript">
function ini(){
	document.login.USUARIO.focus();
}
function validaForm(){
    var d = document.login;
	if (d.USUARIO.value == ""){
		alert('Digite o nome de usuário!');
		d.USUARIO.focus();
		return false;
	}
	else if (d.SENHA.value == ""){
		alert('Digite a senha de acesso!');
		d.SENHA.focus();
		return false;
	}
	else {
		return true;
	}
}
</script>
</head>
<body onLoad="ini()">
<form method="post" name="login" onSubmit="return validaForm()">
  <div class="login-container">
	<p style="color: #ff4d6d; font-size: 13px; margin-top: 10px;">
		<%=mensagem%><%=Request.QueryString("msg")%>
	  </p>
    <img src="images/logoparafundopreto.png" alt="Logo Stillos" />
    <h2>Portal Stillos</h2>

    <input type="text" name="USUARIO" placeholder="Usu&aacute;rio" />
    <input type="password" name="SENHA" placeholder="Senha" />
    <input type="submit" value="Entrar" />

    <p>Este é um software particular e o uso indevido por pessoas não autorizadas pode acarretar em punições previstas por lei.<br />
	</p>
	<p>		

    Para segurança o I.P. do seu computador foi registrado: <%=Request.ServerVariables("remote_addr")%> - <%=Now()%>.</p>
  </div>
</form>
</body>
</html>

