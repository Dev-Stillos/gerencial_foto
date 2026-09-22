<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim COD_VENDEDOR, NOME_VENDEDOR, COD_PRODUCAO, DESCRICAO
COD_VENDEDOR = Request.QueryString("COD_VENDEDOR")
COD_PRODUCAO = Request.QueryString("COD_PRODUCAO")

AbrirConexao()
	Set rs=objConexao.Execute("SELECT NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR = "&COD_VENDEDOR)
		
		If not rs.Eof Then
			NOME_VENDEDOR = rs("NOME_VENDEDOR")
		End If
		
	rs.Close
	Set rs=Nothing

	Set rs=objConexao.Execute("SELECT DESCRICAO FROM TB_PRODUCAO WHERE COD_PRODUCAO = "&COD_PRODUCAO)
		
	If not rs.Eof Then
		DESCRICAO = rs("DESCRICAO")
	End If
		
	rs.Close
	Set rs=Nothing

	'/////////////////////////////////////////////////////////////////////////
	'1) Limpa a base temporária aqui, pois se gerar somente no recibo poderá dar problema na impressão e eles perderem o recibo
	objConexao.Execute("DELETE FROM TB_PRODUCAO_LOTES_FORMANDOS_TEMPORARIA WHERE COD_USUARIO = '"&Session("COD_USUARIO")&"'")	
	
	Dim RENOVAR_DEVOLVER, MOTIVO_RENOVAR_DEVOLVER, PERIODO_RENOVAR, PODE_GRAVAR
	PERIODO_RENOVAR = 0
	RENOVAR_DEVOLVER = 0
	MOTIVO_RENOVAR_DEVOLVER = ""

	AbrirConexao()
	
	'Agora seleciona somente os nao vendidos
	Set rs=objConexao.Execute("SELECT * FROM VW_TB_PRODUCAO_LOTES_FORMANDOS WHERE COD_VENDEDOR = "&COD_VENDEDOR&" and COD_PRODUCAO = "&COD_PRODUCAO&" and STATUS = 3 ORDER BY NOME_FORMANDO",qtd_formandos)
	
	While not rs.Eof 
		
		PODE_GRAVAR = true
		
		PERIODO_RENOVAR = INT(formatNumber(0&Request.Form("PERIODO_RENOVAR_"&rs("COD_FORMANDO")),0))
		RENOVAR_DEVOLVER = INT(formatNumber(0&Request.Form("RENOVAR_DEVOLVER_"&rs("COD_FORMANDO")),0))
		MOTIVO_RENOVAR_DEVOLVER = ""&Request.Form("MOTIVO_RENOVAR_DEVOLVER_"&rs("COD_FORMANDO"))
		
		If RENOVAR_DEVOLVER > 0 Then
		
		
			'//////////////////
			'2 Gravação na base temporária
			objConexao.Execute("INSERT INTO TB_PRODUCAO_LOTES_FORMANDOS_TEMPORARIA (COD_FORMANDO,RENOVAR_DEVOLVER,PERIODO_RENOVAR ,MOTIVO_RENOVAR_DEVOLVER,COD_USUARIO) VALUES ('"&rs("COD_FORMANDO")&"','"&RENOVAR_DEVOLVER&"','"&PERIODO_RENOVAR&"',"&sqlTexto(MOTIVO_RENOVAR_DEVOLVER)&","&Session("COD_USUARIO")&")")
			
			If 	RENOVAR_DEVOLVER = 1 Then
			
				'/////////////////////
				'RENOVAR
				
				'1) Pega a data do prazo e aumenta para uma nova data com base no parametro de dias de renovação e na data adual
					objConexao.Execute("UPDATE TB_PRODUCAO_LOTES_FORMANDOS SET DATA_PRAZO = "&sqlData(dateAdd("d",PERIODO_RENOVAR,Date))&" WHERE COD_FORMANDO ='"&rs("COD_FORMANDO")&"' LIMIT 1")
				
				'2) Altera nas movimentações o novo prazo somente das movimentações que estão em aberto
					objConexao.Execute("UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET DATA_PREVISTA_RETORNO = "&sqlData(dateAdd("d",PERIODO_RENOVAR,Date))&" WHERE COD_FORMANDO ='"&rs("COD_FORMANDO")&"' AND DATA_RETORNO IS NULL")				
								
				'3) Insere o motivo da renovação na ficha do formando, sem necessidade de contato
					objConexao.Execute("INSERT INTO TB_FICHAS_IDENTIFICACAO_HISTORICO (COD_FORMANDO, GERADO_POR, MENSAGEM, COD_VENDEDOR) VALUES ('"&rs("COD_FORMANDO")&"','"&Session("COD_USUARIO")&"',"&sqlTexto(MOTIVO_RENOVAR_DEVOLVER)&", "&sqlNumero(0&COD_VENDEDOR)&")")

				'Obs.: Flag cobrar permanece como 2 (Não cobrar)			
				'/////////////////////
			Elseif RENOVAR_DEVOLVER = 2 Then
			
				'/////////////////////
				'DEVOLVER
				
				'1) Tira a responsabilidade do vendedor da TB_LOTES_FORMANDO
				objConexao.Execute("UPDATE TB_PRODUCAO_LOTES_FORMANDOS SET COD_VENDEDOR = 0 WHERE COD_FORMANDO ='"&rs("COD_FORMANDO")&"' LIMIT 1")

				'2) Da baixa na entrada da ficha na TB_MOVIMENTACOES
				objConexao.Execute("UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET DATA_RETORNO = NOW() WHERE COD_FORMANDO ='"&rs("COD_FORMANDO")&"' AND TIPO = 1 LIMIT 1")

				'3) Marcar na TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES que o prazo de retorno passa a ser hoje para todos os itens exceto a ficha e exceto os que já tiverem sido devolvidos
				objConexao.Execute("UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET DATA_PREVISTA_RETORNO = NOW() WHERE COD_FORMANDO ='"&rs("COD_FORMANDO")&"' AND TIPO <> 1 AND DATA_RETORNO IS NULL")

				'4) Insere o motivo da renovação na ficha do formando, com necessidade de contato
				objConexao.Execute("INSERT INTO TB_FICHAS_IDENTIFICACAO_HISTORICO (COD_FORMANDO,GERADO_POR,MENSAGEM, COD_VENDEDOR) VALUES ('"&rs("COD_FORMANDO")&"','"&Session("COD_USUARIO")&"',"&sqlTexto(MOTIVO_RENOVAR_DEVOLVER)&", "&sqlNumero(0&COD_VENDEDOR)&")")	

				objConexao.Execute("INSERT INTO TB_FICHAS_IDENTIFICACAO_RETORNAR_CONTATO (COD_FORMANDO,COD_USUARIO) VALUES ("&rs("COD_FORMANDO")&",0)")
				'/////////////////////
			
			End If
			
			'### Limpa o motivo temporário do formando
			objConexao.Execute("DELETE FROM TB_PRODUCAO_LOTES_FORMANDOS_MOTIVO_TEMPORARIA WHERE COD_FORMANDO = "&rs("COD_FORMANDO"))
  	End If
  
  rs.MoveNext()
  Wend
  
		  
		  
	rs.Close
	Set rs=Nothing
	FecharConexao()

	Response.Redirect("dip_controle_producoes_listagem_confirma_finaliza_imprime_recibo.asp?cod_vendedor="&COD_VENDEDOR&"&cod_producao="&COD_PRODUCAO)
		 
%>



