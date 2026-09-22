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
  
  ' objConexao.execute("UPDATE TB_SISTEMAS_IPS SET ip_habilitado = '"&ip&"' WHERE descricao = 'Gráfica (não alterar)'")

   'url = "http://ddns.stillos.com.br/stillos.php?a="&now()
   
   'Set objSrvHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
   'objSrvHTTP.setTimeouts 5000, 5000, 5000, 5000
   'objSrvHTTP.Open "GET", url, False
   'objSrvHTTP.Send()
   
   'ip=objSrvHTTP.responseText
   
   'Set objSrvHTTP = Nothing
  '179.153.202.98
   'objConexao.execute("UPDATE TB_SISTEMAS_IPS SET ip_habilitado = '"&ip&"' WHERE 	descricao = 'Stillos (não alterar)'")
   'objConexao.execute("UPDATE TB_SISTEMAS_IPS SET ip_habilitado = '187.38.53.48' WHERE 	descricao = 'Stillos (não alterar)'")
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
					
			')Validações
			'1) Verifica nome de usuário e senha
			'2) Verifica se possui acesso externo, se não tiver verifica o IP da rede

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
				If int(PERMITIR_USO_EXTERNO) = 2 Then
                 
                    
					'Verifica o IP                  
					sql = "SELECT cod_sistema_ip FROM TB_SISTEMAS_IPS WHERE ip_habilitado = "&sqlTexto(Request.ServerVariables("REMOTE_ADDR"))
					Response.Write(sql)
					Set rsIp = objConexao.Execute(sql)

						If rsIp.Eof Then
						'Significa que não encontrou nenhum IP válido
								ERRO_NAO_LOGAR = 2
								COD_ERRO 	   = 2						
						End If
						
					rsIp.Close
					Set rsIp = Nothing	
				End If		
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
				Response.Redirect("index.asp?msg="& Server.URLencode("Senha e/ou usuário inválido(s)! 1") )
			ElseIf ERRO_NAO_LOGAR = 2 Then
				Session.Abandon()
				Response.Redirect("index.asp?msg="& Server.URLencode("Atenção seu login não permite acesso para este número de IP<br />Se você estiver usando a própria rede da Stillo's entre em contato com o(a) responsável pelo sistema na empresa</strong>!"))
			Else
				'Caso contrário permite o acesso e grava o log
				'Log do usuário
				'Conta a quantidade de acessos e guarda o último acesso
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
					Response.Redirect("adm_index.asp")
			End If
End if	
	
%>

<html>
<head>
<title>...::: &Aacute;rea Restrita :::...</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
/* Caracteristicas gerais */
BODY
{
	font-size: 12px;
	font-family: Verdana, Arial;
	margin:0px;
	text-align:center;
	background-color: #F5F5F5;
}

TBODY
{
	font-size:12px;
}

#pagina
{
	margin-top:10px;
	width:800px;
	text-align:left;
	background-color: White;
	border:1px solid black;
}

#cabecalho
{
	filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='#4B92D9', endColorStr='#CEDFF6', gradientType='1');
	margin-bottom: 0px;
	height: 80px;
	float: left;
	width: 100%;
}

#menu
{
	width: 160px;
	list-style-type: square;
	float:left;
	padding:10px;
	margin-bottom:0px;
}

#conteudo
{
	padding-top:0px;
	padding-left:20px;
	BORDER:1PX SOLID WHITE;
	float: left;
	width: 525px;
	position: relative;
	padding-right: 20px;
	padding-bottom: 20px;
}


#rodape
{
	background-color:white;
	height:40px;
	border-top: 1px solid #CCCCCC;
	padding:10px;
	font-size:9px;
	margin:0px;
	clear:both;

}

#rodape H6
{
	margin:0px;
	padding:0px;
	font-size: 10px;
	
}


#cabecalho H1
{
	padding-right: 10px;
	padding-left: 10px;
	font-size: 16px;
	padding-bottom: 00px;
	margin: 0px;
	color: #ddeeff;
	padding-top: 10px;
	letter-spacing: -1px;
}

#cabecalho H2
{
	padding-right: 10px;
	padding-left: 10px;
	font-size: 12px;
	padding-bottom: 10px;
	margin: 0px;
	color: #ddeeff;
	padding-top: 0px;
	letter-spacing: -1px;
}


#conteudo H2
{
	font-size: 14px;
	color: #001644;
	letter-spacing:-1px;
	margin-bottom:30px;
	border-bottom:1px solid #CCCCCC;
}
#menu  ul
{
	list-style:none;
	margin: 0px;
	padding: 0px;

}

#menu LI
{
	margin: 3px;
	
}

#menu A
{
	color:Black;
	font-size:10px;
	text-decoration:none;
	
}

#menu A:Hover
{
	background-color:CCCCCC;
		
	
}



.Mensagem
{
	width: 100%;
	height: 25px;
	background-color: slategray;
	padding-top:5px;
	padding-right:10px;
	text-align:right;
	font-size:10px;
	color:White;
	
	filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='#CEDFF6', endColorStr='#4B92D9', gradientType='1');
}

.MensagemFicha
{
	font-size:13px;
}


#menu H2
{
	font-size:12px;
	margin-top: 10px;
	margin-right: 0px;
	margin-bottom: 10px;
	margin-left: 0px;

}
#tb_principal{
	border: 1px solid #CCCCCC;
}
#red{
color:#CC0000;
}
-->
</style>
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<style type="text/css">
<!--
.style2 {
	font-size: 12px;
	padding: 10px;
}
.style3 {
	color: #CC0000;
	font-size: 13px;
}
.style4 {
	font-size: 11px;
	color: #333333;
}
-->
</style>
<script language="javascript">
function ini(){
	
	document.login.USUARIO.focus();

}

function validaForm(){

d = document.login

	if (d.USUARIO.value == ""){
		
		alert('Digite o nome de usuário!')
		d.USUARIO.focus();
		return false;
		
	}
	else if (d.SENHA.value == ""){
		
		alert('Digite a senha de acesso!')
		d.SENHA.focus();
		return false;
	}	
	else
	{
		return true;
	}

}
</script>
</head>
<body onLoad="ini()">

	 <table cellpadding="9px" cellspacing="0" width="100%">
<tr>
<td>
<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td>
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="24%" background="images/layout/fundo_menu.jpg"><img src="images/layout/topo_esquerda.jpg" width="288" height="58" border="0" /></td>
        <td width="75%" background="images/layout/fundo_menu.jpg" valign="top">
		<form name="buscaForm" action="busca.asp" style="margin:0px">
		</form>
		
		</td>
        <td width="1%"><img src="images/layout/topo_direita.jpg" width="15" height="58" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" height="284" cellpadding="0" cellspacing="0">
  <tr>
<td valign="top" bgcolor="#f5f5f5">

 <form method="post" name="login" onSubmit="return validaForm()">
        <p>&nbsp;</p>
<table style="background-image:url(images/layout/cadeado_index.jpg); background-repeat:no-repeat; background-position:right" width="500" border="0" align="center" cellpadding="0" cellspacing="5">
              <tr>
                <td colspan="2"><div align="center" class="style2"> &Aacute;rea administrativa</div></td>
              </tr>
              <tr>
                <td width="129"><div align="right"><strong>Usu&aacute;rio:</strong></div></td>
                <td width="356"><label>
                  <input style="border:1px #999999 solid; width:210px; padding:4px; height:25px; font-size:13px;" type="text" name="USUARIO">
                  </label>                </td>
              </tr>
              <tr>
                <td><div align="right"><strong>Senha:</strong></div></td>
                <td><input style="border:1px #999999 solid; width:210px; padding:4px; height:25px; font-size:13px;" type="password" name="SENHA"></td>
              </tr>
              <tr>
                <td colspan="2"><div align="center">
                    <label><input name="ENTRAR" class="btnEntrar" type="submit" value=" " onClick="return validaForm()">
                    </label>
                </div></td>
              </tr>
              <tr>
                <td colspan="2"><div align="center" style="padding-top:5px">
                  <p class="style3">&nbsp;</p>
                  <p class="style3"><%=mensagem%><%=Request.QueryString("msg")%></p>
                </div></td>
              </tr>
            </table>
			  <p align="center" class="style4">Este é um software particular e o uso indevido por pessoas não autorizadas pode acarretar em punições previstas por lei.
<br />
Para segurança o I.P. do seu computador foi registrado: <%=Request.ServerVariables("remote_addr")%> - <%=Now()%>.</p>
			  <p></p>
			  <p>&nbsp;</p>
			  <p>&nbsp;</p>
			  <p>
			  	
			  </p>
 </form>
  </td>
 </tr>
     </table>
</td>
  </tr>
</table>
</td>
</tr>
</table>
</body>
</html>
