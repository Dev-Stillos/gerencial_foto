<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Server.ScriptTimeout = 500000
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
</head>
<body>
<%
Dim objConexao12

Sub AbrirConexao12()
	Set objConexao12 = Server.CreateObject("ADODB.Connection")
''	objConexao12.Open("Provider=SQLOLEDB.1;Password=Q@yE2v6_Ta56j;Persist Security Info=True;User ID=stillos5;Initial Catalog=stillos5;Data Source=sqlserver02.stillos.com.br")
	
End sub

Sub FecharConexao12()
   objConexao12.Close
   Set objConexao12=nothing
End sub

AbrirConexao()

Dim rs, rs2, rs3, sql
Dim strConn, conn

Dim objFSO,oInStream,sRows,arrRows  
Dim sFileName  
Dim parte : parte = 0
Dim qtd_por_vez : qtd_por_vez = 1000
 

If Request.QueryString("finalizado") <> Empty Then
	Response.Write("Importação finalizada!!!")
Else
	If Request.QueryString("parte") <> Empty Then
		parte = int(0&Request.QueryString("parte"))/1
	End If 
	
	 
	sFileName = "e:\home\stillos\dados\foto.csv"  
	'*** Create Object ***'  
	Set objFSO = CreateObject("Scripting.FileSystemObject")  
	  
	'*** Check Exist Files ***'  
	If Not objFSO.FileExists(sFileName) Then  
		Response.write("File not found.")  
	Else  
	 
	 	Dim Contrato, Turma, Codigo_Formando, Nome_Formando, Codigo_Curso, Nome_Curso, Periodo, FONE_RES, FONE_COM_CEL, ENDERECO, Bairro, Municipio, Estado, Cep, Semestre, Ano, cpf
	  
	'*** Open Files ***'  
		Set oInStream = objFSO.OpenTextFile(sFileName,1,False)  
		  
			
				If parte = 0 Then
					inicio = 0
				Else
					inicio = qtd_por_vez*parte	
				End If
			
				Response.Write("Parte atual: "&inicio&":"&(inicio+qtd_por_vez)/1&"<br />")
			
				If parte > 0 Then	
					For i=0 to inicio-1
						If not oInStream.AtEndOfStream Then	
						'Pula linhas
							'Response.Write(oInStream.line&"<br />")
							oInStream.SkipLine
							
						End If
					Next
				End If
				
			
				For i=inicio to ((inicio+qtd_por_vez)-1)/1
	
					If not oInStream.AtEndOfStream Then
						
						arrRows = Split(Replace(oInStream.readLine,"""",""),";") 
												
						Contrato = Trim(""&arrRows(0))
						Turma = Trim(""&arrRows(1))
						Codigo_Formando = Trim(""&arrRows(2))
						Nome_Formando = Trim(""&arrRows(3))
						Codigo_Curso = Trim(""&arrRows(4))
						Nome_Curso = Trim(""&arrRows(5))
						Periodo = Trim(""&arrRows(6))
						FONE_RES = Trim(""&arrRows(7))
						FONE_COM_CEL = Trim(""&arrRows(8))
						ENDERECO = Trim(""&arrRows(9))
						Bairro = Trim(""&arrRows(10))
						Municipio = Trim(""&arrRows(11))
						Estado = Trim(""&arrRows(12))
						Cep = Trim(""&arrRows(13))
						cpf = Trim(""&arrRows(14))
						'response.write ("AA")
						Response.Write(Contrato&"<br />")
						
						'2/3) Verifica se o formando já não foi importado para a tabela de importaÃ§Ã£o para foto
							Set rs2=objConexao.Execute("SELECT codigo_formando FROM TB_FORMANDOS_IMP WHERE Codigo_Formando = '"&Codigo_Formando&"'")
						
							If rs2.Eof Then
						
								SEMESTRE_CONCLUSAO = ""
								ANO_CONCLUSAO = ""
								
								'Verifica o contrato e semestre na tabela entidades, pois não poderá ser diferente da entidade
								Set rs3=objConexao.Execute("SELECT SEMESTRE, ANO FROM TB_ENTIDADES WHERE COD_CONTRATO = "&Contrato)
										
									If not rs3.Eof Then
										SEMESTRE_CONCLUSAO	= Ucase(""&rs3("SEMESTRE"))
										ANO_CONCLUSAO = rs3("ANO")
									End If
									
								rs3.Close
								Set rs3=Nothing
								
								
								sql = "INSERT INTO TB_FORMANDOS_IMP (Contrato, Turma, Codigo_Formando, Nome_Formando, Codigo_Curso, Nome_Curso, Periodo, `FONE-RES`, `FONE-COM-CEL`, ENDERECO, Bairro, Municipio, Estado, Cep, Semestre, Ano,cpf) VALUES ("&sqlTexto(Contrato)&","&sqlTexto(Turma)&","&sqlTexto(Codigo_Formando)&","&sqlTexto(Nome_Formando)&","&sqlTexto(Codigo_Curso)&","&sqlTexto(Nome_Curso)&","&sqlTexto(Periodo)&","&sqlTexto(FONE_RES)&","&sqlTexto(FONE_COM_CEL)&","&sqlTexto(ENDERECO)&","&sqlTexto(Bairro)&","&sqlTexto(Municipio)&","&sqlTexto(Estado)&","&sqlTexto(Cep)&","&sqlTexto(SEMESTRE_CONCLUSAO)&","&sqlTexto(ANO_CONCLUSAO)&","&sqlTexto(cpf)&")"
								
								'Response.Write(sql&"<br />")
								objConexao.Execute(sql)
							
							End If
							
							rs2.Close
							Set rs2=Nothing
						
						Else
							Response.Redirect("adm_rotina_importa_formandos_formatura.asp?finalizado=SIM")
						End If
					
					
				Next
		
		'Do Until oInStream.AtEndOfStream 
		'Loop  
		
		oInStream.Close()  
		Set oInStream = Nothing  
			
		Response.AddHeader "Refresh", "5;URL="&"adm_rotina_importa_formandos_formatura.asp?parte="&(parte+1)/1
			
		
	End If  
End If

FecharConexao()



If 1=2 Then

AbrirConexao12()
AbrirConexao()


Dim SEMESTRE_CONCLUSAO, ANO_CONCLUSAO

 
	'1/3) Seleciona todos os formandos da base da formatura atualizada
	Set rs=objConexao12.Execute("SELECT TOP 250 * FROM TB_FORMANDOS_FOTO_EXP WHERE importado = 'N' order by contrato desc")
	
	If not rs.Eof Then
	
		While not rs.Eof
				
				'2/3) Verifica se o formando jÃ¡ nÃ£o foi importado para a tabela de importaÃ§Ã£o para foto
				Set rs2=objConexao.Execute("SELECT Codigo_Formando FROM TB_FORMANDOS_IMP WHERE Codigo_Formando = '"&rs("Codigo_formando")&"'")
					
					'NÃ£o existe insere
					If rs2.Eof Then
						
						SEMESTRE_CONCLUSAO = ""
						ANO_CONCLUSAO = ""
						
						'Verifica o contrato e semestre na tabela entidades, pois não poderá ser diferente da entidade
						Set rs3=objConexao.Execute("SELECT SEMESTRE, ANO FROM TB_ENTIDADES WHERE COD_CONTRATO = "&rs("Contrato"))
								
							If not rs3.Eof Then
								SEMESTRE_CONCLUSAO	= Ucase(""&rs3("SEMESTRE"))
								ANO_CONCLUSAO = rs3("ANO")
							End If
							
						rs3.Close
						Set rs3=Nothing
										
												
						'Insere, pois nÃ£o existe na base da foto
						sql = "INSERT INTO TB_FORMANDOS_IMP (Contrato, Turma, Codigo_Formando, Nome_Formando, Codigo_Curso, Nome_Curso, Periodo, `FONE-RES`, `FONE-COM-CEL`, ENDERECO, Bairro, Municipio, Estado, Cep, Semestre, Ano,cpf) VALUES ("&sqlTexto(rs("Contrato"))&","&sqlTexto(rs("Turma"))&","&sqlTexto(rs("Codigo_Formando"))&","&sqlTexto(rs("Nome_Formando"))&","&sqlTexto(rs("Codigo_Curso"))&","&sqlTexto(rs("Nome_Curso"))&","&sqlTexto(rs("Periodo"))&","&sqlTexto(rs("Fone_Residencial"))&","&sqlTexto(rs("Fone_Celular"))&","&sqlTexto(rs("ENDERECO"))&","&sqlTexto(rs("Bairro"))&","&sqlTexto(rs("Municipio"))&","&sqlTexto(rs("Estado"))&","&sqlTexto(rs("Cep"))&","&sqlTexto(SEMESTRE_CONCLUSAO)&","&sqlTexto(ANO_CONCLUSAO)&","&sqlTexto(rs("cpf"))&")"
						
						'Response.Write(sql&"<br />")
						objConexao.Execute(sql)
	
					Else
					'Atualiza o CPF
						sql = "UPDATE TB_FORMANDOS_IMP SET cpf = "&sqlTexto(rs("cpf"))&" WHERE Codigo_Formando = "&rs("Codigo_formando")
						objConexao.Execute(sql)
						
	
					End If
					
					'3/3) Atualiza como importado, para nÃ£o carregar a mesma pÃ¡gina, mesmo que nÃ£o inseriu
					objConexao12.Execute("UPDATE TB_FORMANDOS_FOTO_EXP SET importado = 'S' WHERE Codigo_formando = '"&rs("Codigo_formando")&"'")
					
				rs2.Close
				Set rs2=Nothing
			
			
		rs.MoveNext()
		Wend
	
	Else
		
		Response.Write("NÃ£o hÃ¡ dados a serem inseridos!")
		
	End If	
	
	rs.Close
	Set rs=Nothing

FecharConexao()
FecharConexao12()

End If
%>
</body>
</html>

