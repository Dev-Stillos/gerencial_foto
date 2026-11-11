<%
Dim COD_USUARIO_INTERNO_LOGADO, COD_PERFIL_LOGADO, ENTRAR_Zz5487iO8Tttt, VALIDA_PAG, COD_PAGINA_USUARIO, sqlADM, rsADM, sqlADM2,rsADM2, ADM_HABILITADA_PARA_TODOS_AAS
Dim redirecionaUsuario
redirecionaUsuario = "/gerencial_foto"
COD_USUARIO_INTERNO_LOGADO 	= Session("COD_USUARIO")
COD_PERFIL_LOGADO			= Session("COD_PERFIL")

VALIDA_PAG = Request.ServerVariables("URL")
VALIDA_PAG = Replace(VALIDA_PAG,"/loja/","")
VALIDA_PAG = Replace(VALIDA_PAG,"/loja2/","")
VALIDA_PAG = Replace(VALIDA_PAG,"/gerencial_foto/","")

'Verifica se o usuário está logado
If COD_USUARIO_INTERNO_LOGADO = Empty Or COD_PERFIL_LOGADO = Empty Then
Session.Abandon()
Response.Redirect(redirecionaUsuario)		
end if

AbrirConexao() 
'Verifica se a página solicitada possui cadastrada
sqlADM2 = "SELECT * FROM TB_ADM_PAGINAS_HABILITADAS WHERE CAMINHO_VALIDA = '"&VALIDA_PAG&"' "
Set rsADM2 = objConexao.execute(sqlADM2)

'VALIDA A PAGINA
If Not rsADM2.eof Then
	COD_PAGINA_USUARIO = formatNumber(0&rsADM2("COD_PAGINA"),0)
	ENTRAR_Zz5487iO8Tttt = True
	ADM_HABILITADA_PARA_TODOS_AAS = rsADM2("habilitada_para_todos")
Else
    rsADM2.Close
    Set rsADM2 = Nothing
    FecharConexao()
    ExibirAcessoNegadoGerencial
End If
rsADM2.Close
Set rsADM2=Nothing

'VALIDA O USUÁRIO
if ENTRAR_Zz5487iO8Tttt = True Then
	sqlADM = "SELECT * FROM TB_ADM_PERFIS_PAGINAS_HABILITADAS WHERE COD_PERFIL = " &COD_PERFIL_LOGADO& " AND COD_PAGINA = "&COD_PAGINA_USUARIO
	
	Set rsADM = objConexao.execute(sqlADM)
		If Not rsADM.eof Then
			ENTRAR_Zz5487iO8Tttt = True
			
			'Verifica se há sub-validações para página
			
		sql="SELECT "&_
			"TB_ADM_PAGINAS_HABILITADAS_SUB_VALIDACOES.COD_SUB_VALIDACAO, "&_
			"TB_ADM_PAGINAS_HABILITADAS_SUB_VALIDACOES.NOME_SUB_VALIDACAO "&_
			"FROM TB_ADM_PAGINAS_HABILITADAS_SUB_VALIDACOES INNER JOIN "&_
			"TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES ON "&_
			"TB_ADM_PAGINAS_HABILITADAS_SUB_VALIDACOES.COD_SUB_VALIDACAO = TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES.COD_SUB_VALIDACAO WHERE "&_
			"TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES.COD_PERFIL="&COD_PERFIL_LOGADO&" AND TB_ADM_PAGINAS_HABILITADAS_SUB_VALIDACOES.COD_PAGINA="&COD_PAGINA_USUARIO
					
				Set rsAdmSubValidacoes = objConexao.Execute(sql)
				If not rsAdmSubValidacoes.Eof Then
				Dim sub_validacoes_paginas
					
					While not rsAdmSubValidacoes.Eof
					
						sub_validacoes_paginas = sub_validacoes_paginas&"#"&rsAdmSubValidacoes("COD_SUB_VALIDACAO")
     
					rsAdmSubValidacoes.MoveNext()
					Wend
				
				End If
				
			 rsAdmSubValidacoes.Close
			 Set rsAdmSubValidacoes = Nothing
		
			
		ElseIf int(ADM_HABILITADA_PARA_TODOS_AAS) = 2 Then 
		'Se ela não for habilitada para todos então bloqueia
            rsADM.Close
            Set rsADM = Nothing
            FecharConexao()
            ExibirAcessoNegadoGerencial

		end if

	rsADM.Close
	Set rsADM=Nothing

Else

    FecharConexao()
    ExibirAcessoNegadoGerencial

end if 
FecharConexao()


Sub ExibirAcessoNegadoGerencial()
    Response.Clear
    Response.Status = "403 Forbidden"
    Response.CacheControl = "no-cache"
    Response.ContentType = "text/html"
    Response.Write "<!doctype html><html lang='pt-BR'><head>"
    Response.Write "<meta name='viewport' content='width=device-width, initial-scale=1'>"
    Response.Write "<title>Acesso restrito | Gerencial Foto</title>"
    Response.Write "<style>"
    Response.Write "body{margin:0;font:1rem system-ui,sans-serif;background:#f6f3f7;color:#29243a}"
    Response.Write ".fundo{box-sizing:border-box;min-height:100vh;display:grid;place-items:center;padding:20px}"
    Response.Write ".aviso{box-sizing:border-box;width:100%;max-width:420px;background:white;border:1px solid #e6dfeb;border-radius:16px;padding:28px;box-shadow:0 12px 36px #29243a1a}"
    Response.Write "h1{font-size:1.4rem;margin:0 0 12px}p{line-height:1.5;margin:0 0 20px}"
    Response.Write "a{display:inline-block;background:#9e125c;color:white;text-decoration:none;border-radius:8px;padding:10px 14px}a:focus-visible{outline:3px solid #29243a;outline-offset:3px}"
    Response.Write "</style></head><body><main class='fundo'>"
    Response.Write "<section class='aviso' role='alertdialog' aria-modal='true' aria-labelledby='titulo' aria-describedby='mensagem'>"
    Response.Write "<h1 id='titulo'>Acesso restrito</h1>"
    Response.Write "<p id='mensagem'>Seu perfil n&atilde;o tem permiss&atilde;o para acessar esta p&aacute;gina.</p>"
    Response.Write "<a href='/gerencial_foto/adm_index.asp' id='voltar'>Voltar ao painel</a></section></main>"
    Response.Write "<script>document.getElementById('voltar').focus();</script></body></html>"
    Response.End
End Sub
%>
