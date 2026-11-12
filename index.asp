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


<%
Dim portalMensagem, portalEntidade
portalMensagem = Server.HTMLEncode(Trim(Request.QueryString("msg") & ""))
' Existing redirects send HTML entities in these messages.
For Each portalEntidade In Array("aacute", "agrave", "atilde", "ccedil", "eacute", "iacute", "oacute", "otilde", "uacute", "Aacute")
    portalMensagem = Replace(portalMensagem, "&amp;" & portalEntidade & ";", "&" & portalEntidade & ";")
Next
portalMensagem = Replace(portalMensagem, "&lt;br /&gt;", " ")
%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="robots" content="noindex,nofollow">
<title>Portal Stillos</title>
<style>
:root {
  --ink:#242236; --muted:#5d5b70; --night:#26253a; --line:#e3e0ea;
  --accent:#a90b60; --accent-dark:#89074d; --canvas:#f7f6fa;
}
* { box-sizing:border-box; }
html { scroll-padding-top:20px; }
body { margin:0; background:var(--canvas); color:var(--ink); font:15px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif; }
a { color:inherit; }
.wrap { width:min(1180px,calc(100% - 56px)); margin-inline:auto; }
.masthead { color:#fff; background:#222135; }
.masthead .wrap { min-height:73px; display:flex; align-items:center; justify-content:space-between; gap:20px; }
.brand { display:flex; align-items:center; gap:19px; font-weight:650; }
.brand img { display:block; width:125px; height:auto; }
.brand span { padding-left:19px; border-left:1px solid #5b586e; }
.masthead nav { display:flex; align-items:center; gap:22px; font-size:13px; }
.masthead nav a { color:#ece8f4; text-decoration:none; }
.masthead nav a:hover { text-decoration:underline; }
.hero { position:relative; overflow:hidden; isolation:isolate; color:#fff; background:linear-gradient(120deg,#2c2843 5%,#40263f 62%,#68294e); }
.hero::before {
  content:""; position:absolute; inset:0; z-index:-1; pointer-events:none;
  background:radial-gradient(ellipse at 70% 92%,#de5f936b,transparent 56%),
    repeating-radial-gradient(ellipse at 83% 122%,transparent 0 23px,#f8c9eb20 24px 25px,transparent 26px 35px);
}
.hero .wrap { display:grid; grid-template-columns:minmax(0,1fr) minmax(270px,360px); align-items:center; gap:clamp(24px,6vw,78px); padding-block:clamp(28px,5vw,62px); }
.hero-copy { max-width:600px; }
.hero-copy p:first-child { font-size:12px; font-weight:700; letter-spacing:.14em; text-transform:uppercase; color:#f7bfde; margin:0 0 13px; }
h1,h2,h3,p { margin:0; }
h1 { font-size:clamp(36px,6vw,68px); line-height:1.02; letter-spacing:-.06em; }
.hero-copy .intro { color:#f1e4ee; font-size:clamp(15px,1.8vw,18px); margin:17px 0 22px; max-width:34ch; }
.hero-copy a { display:inline-flex; gap:12px; align-items:center; min-height:42px; border:1px solid #f3c1df80; border-radius:8px; padding:8px 14px; text-decoration:none; font-weight:650; }
.hero-copy a:hover { background:#fff2; }
.login { background:#fff; color:var(--ink); border-radius:16px; border:1px solid #fff; box-shadow:0 17px 45px #1b142f4d; padding:clamp(19px,3vw,28px); }
.login h2 { font-size:23px; letter-spacing:-.025em; }
.login p { color:var(--muted); margin:4px 0 0; font-size:14px; }
.login form { display:grid; gap:13px; margin-top:16px; }
.login label { display:grid; gap:5px; font-size:13px; font-weight:650; }
.login input { min-height:42px; min-width:0; width:100%; border:1px solid #d6d2e1; border-radius:8px; padding:8px 10px; color:var(--ink); font:inherit; background:#fff; }
.login button { min-height:43px; border:0; border-radius:8px; background:var(--accent); color:#fff; cursor:pointer; font:inherit; font-weight:700; }
.login button:hover { background:var(--accent-dark); }
.login .error { padding:9px 11px; margin-top:10px; border-radius:7px; color:#7d1044; background:#fdebf4; overflow-wrap:anywhere; }
.directory { padding-block:30px 36px; }
.section + .section { margin-top:27px; }
.section h2 { font-size:22px; letter-spacing:-.025em; margin-bottom:12px; }
.cards { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:14px; }
.card { display:flex; flex-direction:column; align-items:flex-start; gap:8px; min-height:146px; padding:17px; border:1px solid var(--line); border-radius:13px; background:#fff; text-decoration:none; overflow-wrap:anywhere; }
a.card:hover { border-color:#c188ae; box-shadow:0 5px 16px #46203912; }
.card-top { display:flex; width:100%; align-items:center; justify-content:space-between; gap:9px; }
.card-mark { display:inline-grid; place-items:center; flex:none; width:37px; height:37px; border-radius:10px; color:var(--accent); background:#f8e8f2; font-size:20px; font-weight:700; }
.card:nth-child(2) .card-mark { color:#3e527c; background:#ecf0f9; }
.card:nth-child(3) .card-mark { color:#6b4381; background:#f2eaf8; }
.card-arrow { color:var(--accent); font-size:20px; }
.card h3 { font-size:17px; line-height:1.25; }
.card p { color:var(--muted); font-size:13px; }
.status { font-size:11px; font-weight:700; color:#5d526b; background:#f2eff6; border-radius:20px; padding:3px 8px; }
.footer { border-top:1px solid var(--line); }
.footer .wrap { padding-block:15px; display:flex; justify-content:space-between; gap:10px; flex-wrap:wrap; color:var(--muted); font-size:13px; }
.footer a { color:var(--accent); }
:focus-visible { outline:3px solid #d76aa7; outline-offset:3px; }
@media (max-width:820px) {
  .wrap { width:calc(100% - 36px); }
  .hero .wrap { gap:25px; }
  .cards { grid-template-columns:repeat(2,minmax(0,1fr)); }
}
@media (max-width:610px) {
  .masthead .wrap { min-height:62px; }
  .brand { gap:11px; }
  .brand img { width:96px; }
  .brand span { padding-left:11px; font-size:13px; }
  .masthead nav { gap:10px; }
  .masthead nav a:last-child { display:none; }
  .hero .wrap,.cards { grid-template-columns:minmax(0,1fr); }
  .hero-copy .intro { margin-bottom:13px; }
  .card { min-height:0; }
}
</style>
</head>
<body>
<header class="masthead">
  <div class="wrap">
    <div class="brand"><img src="images/logoparafundopreto.png" alt="Stillos"><span>Portal Stillos</span></div>
    <nav aria-label="Navega&ccedil;&atilde;o do portal">
      <a href="#sistemas">Sistemas</a><a href="https://suporte.stillos.net/">Suporte</a>
    </nav>
  </div>
</header>
<main>
  <section class="hero" aria-labelledby="portal-heading">
    <div class="wrap">
      <div class="hero-copy">
        <p>Intranet do Grupo Stillos</p>
        <h1 id="portal-heading">Sua central<br>de trabalho</h1>
        <p class="intro">Acesse os sistemas do grupo para trabalhar com fotos, pedidos, suporte e rotinas administrativas.</p>
        <a href="#sistemas">Ver sistemas <span aria-hidden="true">&darr;</span></a>
      </div>
      <section class="login" aria-labelledby="login-title">
        <h2 id="login-title">Gerencial Fotos</h2>
        <p>Entre com suas credenciais do Fotos.</p>
        <%If Len(portalMensagem) > 0 Then%>
        <p class="error" role="alert"><%=portalMensagem%></p>
        <%End If%>
        <form method="post" name="login">
          <label for="portal-usuario">Usu&aacute;rio
            <input type="text" id="portal-usuario" name="USUARIO" autocomplete="username" required>
          </label>
          <label for="portal-senha">Senha
            <input type="password" id="portal-senha" name="SENHA" autocomplete="current-password" required>
          </label>
          <button type="submit">Entrar no Fotos</button>
        </form>
      </section>
    </div>
  </section>
  <div class="wrap directory" id="sistemas">
    <section class="section" aria-labelledby="available-title">
      <h2 id="available-title">Acesso r&aacute;pido</h2>
      <div class="cards">
        <a class="card" href="/loja/">
          <div class="card-top"><span class="card-mark" aria-hidden="true">&#9636;</span><span class="card-arrow" aria-hidden="true">&nearr;</span></div>
          <h3>Loja Stillos</h3><p>Lance pedidos e acompanhe suas vendas.</p>
        </a>
        <a class="card" href="https://suporte.stillos.net/">
          <div class="card-top"><span class="card-mark" aria-hidden="true">?</span><span class="card-arrow" aria-hidden="true">&nearr;</span></div>
          <h3>Suporte Stillos</h3><p>Abra chamados e acompanhe os atendimentos.</p>
        </a>
        <a class="card" href="https://sistema.stillos.net/">
          <div class="card-top"><span class="card-mark" aria-hidden="true">&#9638;</span><span class="card-arrow" aria-hidden="true">&nearr;</span></div>
          <h3>Posses e Administra&ccedil;&atilde;o</h3><p>Acesse o sistema de posses e rotinas administrativas.</p>
        </a>
      </div>
    </section>
    <section class="section" aria-labelledby="products-title">
      <h2 id="products-title">Outros sistemas Stillos</h2>
      <div class="cards">
        <article class="card">
          <div class="card-top"><span class="card-mark" aria-hidden="true">&#9635;</span><span class="status">Em breve</span></div>
          <h3>Stillos Estoque</h3><p>Controle de estoque.</p>
        </article>
        <article class="card">
          <div class="card-top"><span class="card-mark" aria-hidden="true">&#9707;</span><span class="status">Link a definir</span></div>
          <h3>Stillos GradFlow Studio</h3><p>Diagrama&ccedil;&atilde;o de &aacute;lbuns de formatura.</p>
        </article>
        <article class="card">
          <div class="card-top"><span class="card-mark" aria-hidden="true">&#9673;</span><span class="status">Link a definir</span></div>
          <h3>Stillos Gov360</h3><p>Governan&ccedil;a, rastreabilidade e acompanhamento de projetos e implanta&ccedil;&otilde;es.</p>
        </article>
      </div>
    </section>
  </div>
</main>
<footer class="footer"><div class="wrap"><span>&copy; <%=Year(Date())%> Grupo Stillos</span><a href="https://suporte.stillos.net/">Precisa de ajuda?</a></div></footer>
</body>
</html>
