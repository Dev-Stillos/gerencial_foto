<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<%
'A idéia desta rotina é importar todas as entidades do SQL para a Base da Foto
AbrirConexao()
AbrirConexaoSQL6()

Dim COD_CONTRATO, DESCRICAO, SEMESTRE, ANO

	Set rs=objConexaoSQL6.Execute("Select * from TB_CONTRATOS")
			
			If not rs.Eof Then
				
				While not rs.Eof
					
					COD_CONTRATO = rs("FD_CODIG_001")
					DESCRICAO	 = rs("FD_NOME_002")
					SEMESTRE	 = rs("FD_MEREV_002")
					ANO			 = int(2000+rs("FD_ANREV_002"))
					
					If int(SEMESTRE) < 10 Then
						SEMESTRE = "J"
					Else
					'10, 11 e 12 são considerados dezembro
						SEMESTRE = "D"
					End If
					
					Set rs2 = objConexao.Execute("SELECT COD_CONTRATO FROM TB_ENTIDADES WHERE COD_CONTRATO = '"&COD_CONTRATO&"'")
							
							If rs2.Eof Then
							'Não existe cadastrar
									
								sql = "INSERT INTO TB_ENTIDADES (COD_CONTRATO, DESCRICAO, SEMESTRE, ANO) VALUES ('"&COD_CONTRATO&"', '"&DESCRICAO&"', '"&SEMESTRE&"', '"&ANO&"')"
								objConexao.Execute(sql)									
							End If
					
					rs2.Close
					Set rs2 = Nothing
					
				rs.MoveNext()
				Wend
				
			End If
	
	rs.Close
	Set rs=Nothing


FecharConexaoSQL6()		
FecharConexao()
%>