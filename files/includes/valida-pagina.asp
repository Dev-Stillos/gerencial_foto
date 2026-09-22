<%
'Está validação é individual por cliente, é a mais correta pois possui a validação por página para cada usuário/perfil

Dim sub_validacao_cod_grupo_acesso_pagina 'esta validação se for utiliza faz com que seja verificado as sub-validações padrões de outras páginas e ignora as da própria página
Dim sub_validacoes_paginas 'esta variavel armazena todas as outras sub_validações se houver
Dim sub_validacoes_paginas_relacionadas 'esta variavel verifica as sub-validações que possuem páginas relaciondas, verificando todas as sub-validações relaciondas a página e quais o usuário possui acesso
Dim usuario_podeVisualizar, usuario_podeIncluir, usuario_podeAlterar, usuario_podeExcluir


	
	AbrirConexao(cod_base)	
		
		'1) Verifica se a página habilitada possui cadastro, e se este cadastro é habilitado para todos
		Set rs=objConexao(cod_base).Execute("Select habilitada_para_todos,titulo_pagina_cliente,sub_validacao_cod_grupo_acesso_pagina  from VW_PADRAO_ACESSO_SISTEMA_GRUPOS_PAGINAS_VALIDACAO WHERE cod_grupo_acesso_pagina ='"&cod_grupo_acesso_pagina&"'")
				
				If not rs.Eof Then
										
					sub_validacao_cod_grupo_acesso_pagina = Cstr(""&rs("sub_validacao_cod_grupo_acesso_pagina"))
					
					
					
					If int(rs("habilitada_para_todos")) = 2  and Session("cod_usuario") <> 1 Then
						'Significa que possui cadastro válido e portando deverá ser validada
							Set rs2=objConexao(cod_base).Execute("Select * From PADRAO_USUARIOS_VALIDACAO_PAGINAS Where  cod_grupo_acesso_pagina = '"&cod_grupo_acesso_pagina&"' and (cod_usuario="&Session("cod_usuario")&" or cod_perfil = '"&PADRAO_SISTEMA_cod_perfil&"')")
								
								If not rs2.Eof Then
									'Está cadastrado e habilitado, basta verificar possíveis sub-validações padrões e outras sub-validações						
									
									'Verifica se irá buscar as sub-validações na mesma página ou em outra página
									If Cstr(sub_validacao_cod_grupo_acesso_pagina) = "" or Cstr(sub_validacao_cod_grupo_acesso_pagina) = "0"  Then
										'Por padrão verifica todas as sub-validações padrões
										
										usuario_podeVisualizar	= formatNumber(0&rs2("visualizar"),0)
										usuario_podeIncluir		= formatNumber(0&rs2("incluir"),0)
										usuario_podeAlterar		= formatNumber(0&rs2("alterar"),0)
										usuario_podeExcluir		= formatNumber(0&rs2("excluir"),0)
																				
									Else
										'Em outra página as sub-validações
										Set rs3=objConexao(cod_base).Execute("Select * From PADRAO_USUARIOS_VALIDACAO_PAGINAS Where cod_grupo_acesso_pagina = '"&sub_validacao_cod_grupo_acesso_pagina&"' and (cod_usuario="&Session("cod_usuario")&" or cod_perfil = '"&PADRAO_SISTEMA_cod_perfil&"')")
											
											If not rs3.Eof Then
												
												usuario_podeVisualizar	= formatNumber(0&rs3("visualizar"),0)
												usuario_podeIncluir		= formatNumber(0&rs3("incluir"),0)
												usuario_podeAlterar		= formatNumber(0&rs3("alterar"),0)
												usuario_podeExcluir		= formatNumber(0&rs3("excluir"),0)
											Else
												'Não possui acesso a página relacionada, por tanto fecha tudo
												
												usuario_podeVisualizar	= 2
												usuario_podeIncluir		= 2
												usuario_podeAlterar		= 2
												usuario_podeExcluir		= 2
											End If
										
										rs3.Close
										Set rs3=Nothing
									End If
									
									'Agora verifica se a página possui sub_validações									
									Set rsAdmSubValidacoes = objConexao(cod_base).Execute("SELECT * FROM VW_PADRAO_USUARIOS_VALIDACAO_PAGINAS_SUBS_COD_GRUPO_ACESSO Where cod_grupo_acesso_pagina = '"&cod_grupo_acesso_pagina&"' and  (cod_usuario="&Session("cod_usuario")&" or cod_perfil = '"&PADRAO_SISTEMA_cod_perfil&"')")
										If not rsAdmSubValidacoes.Eof Then
										
											While not rsAdmSubValidacoes.Eof
												
												sub_validacoes_paginas = sub_validacoes_paginas&"#"&rsAdmSubValidacoes("cod_pagina_validacao_sub_validacao")
							 
											rsAdmSubValidacoes.MoveNext()
											Wend
										
										End If
										
									 rsAdmSubValidacoes.Close
									 Set rsAdmSubValidacoes = Nothing
									
								Else
									'Não possui acesso a página
									Response.Redirect("inicio.asp?msg=Você não possui acesso a esta página!&cod_pagina="&cod_grupo_acesso_pagina)					
								End If
								
	
							rs2.Close
							Set rs2=Nothing
		
					Elseif int(rs("habilitada_para_todos")) = 2 and int(Session("cod_usuario")) = 1 Then
												
						'////////////////////////////////////////////
						'Login OXNET liberado para todas as sub-validações padrões
						'Lembrando que sempre poderá ter somente 1(um) login OXNET 
						'padrão este deverá ser sempre o número 1 para funcionar.
								usuario_podeVisualizar	= 1
								usuario_podeIncluir		= 1
								usuario_podeAlterar		= 1
								usuario_podeExcluir		= 1
								
						'Verifica se há sub-validações para está página se houver habilita todas
								Set rsAdmSubValidacoes = objConexao(cod_base).Execute("SELECT cod_pagina_validacao_sub_validacao FROM PADRAO_ACESSO_SISTEMA_GRUPOS_PAGINAS_VALIDACAO_SUBS WHERE cod_grupo_acesso_pagina = '"&cod_grupo_acesso_pagina&"'")
								
										If not rsAdmSubValidacoes.Eof Then
										
											While not rsAdmSubValidacoes.Eof
												
												sub_validacoes_paginas = sub_validacoes_paginas&"#"&rsAdmSubValidacoes("cod_pagina_validacao_sub_validacao")
							 
											rsAdmSubValidacoes.MoveNext()
											Wend
										
										End If
										
									 rsAdmSubValidacoes.Close
									 Set rsAdmSubValidacoes = Nothing				
								
						'////////////////////////////////////////////
					Else
											
						Set rsSubRel = objConexao(cod_base).Execute("SELECT * FROM PADRAO_ACESSO_SISTEMA_GRUPOS_PAGINAS_VALIDACAO_SUBS WHERE relaciona_cod_grupo_acesso_pagina = '"&cod_grupo_acesso_pagina&"'")
							
						While not rsSubRel.Eof
						'Se encontrar sub-validações relacionadas com está pagina verifica se o usuário possui acesso
						
						
									Set rsAdmSubValidacoes = objConexao(cod_base).Execute("SELECT cod_pagina_validacao_sub_validacao FROM PADRAO_USUARIOS_VALIDACAO_PAGINAS_SUBS WHERE (cod_usuario = "&Session("cod_usuario")&" or cod_perfil = '"&PADRAO_SISTEMA_cod_perfil&"') and cod_pagina_validacao_sub_validacao = '"&rsSubRel("cod_pagina_validacao_sub_validacao")&"'")
							
										'Verifica validação à validação se o usuário possui a sub-validação
										'relacionada a página, se encontrar mais que 1 não tem problema, pois
										'a checagem final é feita manual
										If not rsAdmSubValidacoes.Eof or int(Session("cod_usuario")) = 1 Then
										
												sub_validacoes_paginas_relacionadas = sub_validacoes_paginas_relacionadas&"#"&rsSubRel("cod_pagina_validacao_sub_validacao")
												
										End If
										
									 rsAdmSubValidacoes.Close
									 Set rsAdmSubValidacoes = Nothing			
																	
						rsSubRel.MoveNext()
						Wend
							
					
						rsSubRel.Close
						Set rsSubRel = Nothing
					End If
				
				Else
				
					Response.Redirect("inicio.asp?msg=Está página está inacessível, informe a OXNET a falha OX078#: "&cod_grupo_acesso_pagina)
					
				
				End If
			
			
			
		rs.Close
		Set rs=Nothing
%>

