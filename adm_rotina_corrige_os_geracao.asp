<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<%
  AbrirConexao()
  
  Dim data_base
  data_base = DateAdd("d",-10,date())
  
   'Verifica se esta rotina já foi processada
	Set rsVerifica=objConexao.Execute("SELECT * FROM TB_ADM_ROTINAS WHERE rotina = 'adm_rotina_corrige_os_geracao.asp' AND DAY(data_processada) = '"&day(date())&"' AND Month(data_processada) = '"&Month(date())&"' AND Year(data_processada) = '"&Year(date())&"'")
		
		If rsVerifica.Eof Then
		'Não foi processada na data base
			
			objConexao.Execute("INSERT INTO TB_ADM_ROTINAS (rotina) VALUES ('adm_rotina_corrige_os_geracao.asp')")
			
			Set rs=objConexao.Execute("SELECT cod_pedido FROM TB_PEDIDOS WHERE DATA_VENDA >= '"&year(data_base)&"-"&month(data_base)&"-01' AND FINALIZA_CORRECAO = 1 AND FINALIZA_DIRETORIA = 1",qtdRegistros)
				
				While not rs.Eof
					
					'retorna = moduloPPVgerarOrdensServico(rs("cod_pedido"),true)
					'Response.Write(rs("cod_pedido")&" - "&retorna&"<br />")
					
				rs.MoveNext()
				Wend
		  
			rs.Close
			Set rs=Nothing 		
				
		
		End If
		
	rsVerifica.Close
	Set rsVerifica=Nothing
	
	FecharConexao()
%>
