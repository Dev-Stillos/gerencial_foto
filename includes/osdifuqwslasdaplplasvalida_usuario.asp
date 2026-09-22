<%
Dim COD_USUARIO_INTERNO_LOGADO, COD_PERFIL_LOGADO, ENTRAR_Zz5487iO8Tttt, VALIDA_PAG, COD_PAGINA_USUARIO, sqlADM, rsADM, sqlADM2,rsADM2, ADM_HABILITADA_PARA_TODOS_AAS
Dim redirecionaUsuario
redirecionaUsuario = "https://stillos.com.br/gerencial_foto"
COD_USUARIO_INTERNO_LOGADO 	= Session("COD_USUARIO")
COD_PERFIL_LOGADO			= Session("COD_PERFIL")

VALIDA_PAG = Request.ServerVariables("URL")
VALIDA_PAG = Replace(VALIDA_PAG,"/loja/","")
VALIDA_PAG = Replace(VALIDA_PAG,"/loja2/","")
VALIDA_PAG = Replace(VALIDA_PAG,"/gerencial_foto/","")

'Verifica se o usuário está logado
if COD_USUARIO_INTERNO_LOGADO=Empty OR VALIDA_PAG=Empty then
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
else 
		Session.Abandon()
		Response.Redirect(redirecionaUsuario)	
			
end if
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
			Session.Abandon()
			Response.Redirect(redirecionaUsuario)	

		end if

	rsADM.Close
	Set rsADM=Nothing

Else

	Session.Abandon()
	Response.Redirect(redirecionaUsuario)	

end if 
FecharConexao()


%>
