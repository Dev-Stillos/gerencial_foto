<%
			 AbrirConexao()
		If int(Session("COD_PERFIL")) = 5 Then	
			
		 
		  Dim rsRetContato, qtdRetornar
		  qtdRetornar= 0
		  
		  Set rsRetContato=objConexao.Execute("Select count(*) as qtd FROM TB_FICHAS_IDENTIFICACAO_RETORNAR_CONTATO where COD_USUARIO = '"&Session("cod_usuario")&"' or COD_USUARIO = 0")
		  		
				If not rsRetContato.Eof Then
					qtdRetornar = formatNumber(0&rsRetContato("qtd"),0)
					
					If qtdRetornar > 0 Then
					%>
                  	<a class="session-notice" href="dip_formandos_entrar_em_contato.asp">Entrar em contato(<%=qtdRetornar%>)</a>
            
                      
                    
                <%
					End If
				End If
		  
		  rsRetContato.Close
		  Set rsRetContato=Nothing
		  
		  End If
		  
		  Dim rsEtiquetasAgendadas,qtdAgendada
		  qtdAgendada = 0
		  Set rsEtiquetasAgendadas = objConexao.Execute("SELECT COD_USUARIO FROM TB_FICHAS_IDENTIFICACAO_ETIQUETAS_TEMPORARIA WHERE COD_USUARIO="&Session("COD_USUARIO")&" AND AGENDADA=1",qtdAgendada)
			
			If int(qtdAgendada) > 0 Then
			%>
            <a class="session-notice" href="adm_etiquetas_agendadas.asp">Etiquetas Agendadas(<%=qtdAgendada%>)</a>
            
            <%End If
			
			rsEtiquetasAgendadas.Close
			 Set rsEtiquetasAgendadas = Nothing
			 
			 FecharConexao()
			%>