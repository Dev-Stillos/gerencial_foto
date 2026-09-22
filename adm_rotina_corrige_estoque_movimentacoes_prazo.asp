<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
	'ROTINA CRIADA EM 07/08/2012 - JOSÉ HENRIQUE
	'Verifica se possui movimentações que não foram devolvidas e também não há prazo para retorno
	Dim COD_FORMANDO, DATA_PREVISTA_RETORNO
	
	AbrirConexao()
	
	'1) Seleciona fichas nestas condições e que já tenham data de sáida
	Set rs=objConexao.Execute("SELECT COD_MOVIMENTACAO, COD_FORMANDO, DATA_SAIDA, COD_VENDEDOR, TIPO FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE DATA_RETORNO IS NULL AND DATA_PREVISTA_RETORNO IS NULL AND DATA_SAIDA IS NOT NULL LIMIT 2000")
		
		While not rs.Eof
		
			COD_FORMANDO = rs("COD_FORMANDO")
			
			'2) Verifica se há lote para o formando, que já tenho saído
			Set rs2=objConexao.Execute("SELECT DATA_PRAZO FROM TB_PRODUCAO_LOTES_FORMANDOS WHERE DATA_SAIDA IS NOT NULL AND  COD_FORMANDO = "&COD_FORMANDO)
				
				If not rs2.Eof Then
				'Existe pega o prazo do lote
					DATA_PREVISTA_RETORNO = formatDateTime(rs2(0))			
				Else
				'Não existe pega o prazo da saída e soma 45 dias
					DATA_PREVISTA_RETORNO = formatDateTime(rs("DATA_SAIDA"))
					DATA_PREVISTA_RETORNO = DateAdd("d",45,DATA_PREVISTA_RETORNO)
				End If
			
			rs2.Close
			Set rs2=Nothing
		
			'3) Atualiza a data prevista para retorno
			sql = "UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET DATA_PREVISTA_RETORNO = "&sqlData(DATA_PREVISTA_RETORNO)&" WHERE COD_MOVIMENTACAO = "&rs("COD_MOVIMENTACAO")
			'Response.Write(sql&"<br />")
			objConexao.Execute(sql)
		
		rs.MoveNext()
		Wend					
		
								
	rs.Close
	Set rs=Nothing

	FecharConexao()

%>
<!--
<html>
<head>
<meta http-equiv="refresh" content="5">
<title></title>
</head>
<body>
</body>
</html> 
-->