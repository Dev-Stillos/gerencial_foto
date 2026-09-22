<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim COD_FORMANDO, SEMESTRE_CONCLUSAO, ANO_CONCLUSAO
Dim msg
Dim comando


COD_FORMANDO = Request.QueryString("cod_formando")
comando = Request.QueryString("comando")


If comando = "alterar" Then
	
	AbrirConexao()
	
		NOME_FORMANDO 		= Request.Form("NOME_FORMANDO")
		COD_CONTRATO 		= Request.Form("COD_CONTRATO")
		COD_CURSO 			= Request.Form("COD_CURSO")
		
		
		'--------------------------------------------
		'Alteração realizada em: 06/08/2012
		'Agora não poderá ser separado o ano e semestre do formando, deverá obdecer a entidade selecionada
		
		'///NOVO	
		Set rs=objConexao.Execute("SELECT * FROM TB_ENTIDADES WHERE COD_CONTRATO = "&sqlTexto(COD_CONTRATO))
				
			If not rs.Eof Then
				SEMESTRE_CONCLUSAO	= Ucase(""&rs("SEMESTRE"))
				ANO_CONCLUSAO = rs("ANO")
			End If
			
		rs.Close
		Set rs=Nothing
			
		'///DESABILITADO
			'ANO_CONCLUSAO		=Request.Form("ANO_CONCLUSAO")
			'SEMESTRE_CONCLUSAO	=Ucase(Request.Form("SEMESTRE_CONCLUSAO"))
			
		'--------------------------------------------
		
		INICIAL 			= Request.Form("INICIAL")
		COMISSAO 			= Request.Form("COMISSAO")
		BRINDE 				= Request.Form("BRINDE")
		ALBUM				= Request.Form("ALBUM")
		
		Set rs=objConexao.Execute("SELECT Nome_Curso FROM TB_FORMANDOS_IMP WHERE Codigo_Curso = "&COD_CURSO&" limit 1")
			If not rs.Eof Then
				NOME_CURSO = rs("Nome_Curso")
			End If
		
		rs.Close
		Set rs=Nothing
		
		objConexao.Execute("UPDATE TB_FICHAS_IDENTIFICACAO SET COD_CONTRATO = '"&COD_CONTRATO&"',COD_CURSO = '"&COD_CURSO&"',NOME_CURSO = "&sqlTexto(NOME_CURSO)&",NOME_FORMANDO = '"&NOME_FORMANDO&"', ANO_CONCLUSAO = '"&ANO_CONCLUSAO&"',SEMESTRE_CONCLUSAO = '"&SEMESTRE_CONCLUSAO&"',INICIAL = '"&INICIAL&"',ALBUM = '"&ALBUM&"', BRINDE = '"&BRINDE&"',COMISSAO = '"&COMISSAO&"' WHERE COD_FORMANDO ="&COD_FORMANDO&" LIMIT 1")
	
	'Se escolher brinde "NÃO" apaga todos os brindes do formando se houver
		If int(BRINDE) = 2 Then
			objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_BRINDES WHERE COD_FORMANDO = '"&COD_FORMANDO&"'")
		End If
		
		Dim msg_log
		msg_log = "Formando(a) alterado"
		
		objConexao.Execute("INSERT INTO TB_ADM_LOG_ALTERACOES (COD_USUARIO,cod_formando,DESCRICAO) VALUES ("&Session("cod_usuario")&",'"&COD_FORMANDO&"',"&sqlTexto(msg_log)&")")
	
	FecharConexao()
	
	Response.Redirect("dip_busca_detalhamento_formando.asp?cod_formando="&COD_FORMANDO&"&msg=Formando alterado com sucesso!")
	
End If
%>



