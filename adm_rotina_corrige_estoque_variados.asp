<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
AbrirConexao()
	
	'1) Verificação de álbuns duplos, que deveriam estar zerados
	'////////////////////////////////////////////////////////////////
		
	Set rs=objConexao.Execute("SELECT * FROM VW_CORRIGE_DUPLICIDADE_MOVIMENTACOES where qtd > 1")

		While not rs.Eof
						
			'Verifica se todas as duplicidades possuem data de retorno
			Set rs2=objConexao.Execute("SELECT cod_formando FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES Where tipo = "&rs("tipo")&" and data_saida = "&sqlData(rs("data_saida"))&" and cod_formando = "&rs("cod_formando")&" and data_retorno is null")
				
				
				If rs2.Eof Then
				'Não encontrou nenhum do mesmo formando,tipo,data_saida que possui data de retorno vazia, por tanto pode excluir 1 com data de retorno preenchido
					sql = "DELETE FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE tipo = "&rs("tipo")&" and data_saida = "&sqlData(rs("data_saida"))&" and cod_formando = "&rs("cod_formando")&" limit 1"
					'Response.Write(sql&"<br />")
					 objConexao.Execute(sql)
				Else
				'Caso contrário remove somente 1 sem data de retorno
					 sql = "DELETE FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE tipo = "&rs("tipo")&" and data_saida = "&sqlData(rs("data_saida"))&" and data_retorno is null and cod_formando = "&rs("cod_formando")&" limit 1"
					'Response.Write(sql&"<br />")
					 objConexao.Execute(sql)
				End if
								
			rs2.Close
			Set rs2=Nothing			
			
				  
		rs.MoveNext()	
		Wend
	
	rs.Close
	Set rs=Nothing
	
	
	
FecharConexao()
%>