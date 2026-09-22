
 	  <%
 	  Response.Buffer=true
 	  
Dim Iterations, StartTime, ElapsedTime, AverageTimePerIteration, X

Iterations = 100
StartTime = Timer


				Set objConexao = Server.CreateObject("ADODB.Connection")
				'objConexao.Open("DRIVER={MySQL ODBC 5.1 Driver};SERVER=186.202.185.112;PORT=3306;DATABASE=stillos3;USER=stillos3;PASSWORD=c_m@lc46_@@;OPTION=3;")
				'objConexao.Open("DRIVER={MySQL ODBC 5.1 Driver};SERVER=raposa;PORT=3306;DATABASE=Reportagem;USER=stillos;PASSWORD=stillos;OPTION=3;")
				objConexao.Open("DRIVER={MySQL ODBC 5.1 Driver};SERVER=mysql03.stillos.com.br;PORT=3306;DATABASE=stillos3;USER=stillos3;PASSWORD=c_m@lc46_@@;PREFETCH=5;")
							Dim rs, rs2, rs3, sql,qtd_registros
		
							sql = "SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE NOW() = NOW()"

							Set rs=objConexao.Execute(sql,qtd_registros)
							while not rs.eof
									'response.write(rs(0)&"<br>")
							rs.MoveNext()
						 Wend
						 rs.Close()
						 Set rs = Nothing
							response.write(qtd_registros)
			objConexao.Close
   Set objConexao=nothing
   

ElapsedTime = Timer - StartTime
AverageTimePerIteration = ElapsedTime / Iterations
response.write("<br>")
response.write(AverageTimePerIteration)

Response.Flush
%>
		            
