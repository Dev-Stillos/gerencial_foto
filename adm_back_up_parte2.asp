<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>BackUP - Foto</title>
<style type="text/css">
<!--
.style1 {
	font-size: 36px;
	font-weight: bold;
}
.style4{
	color:#33CC00;
	font-weight: bold;
}
body,td,th {
	font-family: Arial, Helvetica, sans-serif;
}
body {
	background-color: #E9E9E9;
}
-->
</style>
</head>
<body>
<%
Dim CONTINUA_BACKUP
CONTINUA_BACKUP = 1

Dim COD_BACKUP, NOME_PASTA, NOME_BASE, QTD_LINHAS_POR_VEZ,TEMPO_REATUALIZAR
					
QTD_LINHAS_POR_VEZ	= 4000
'Em segundos
TEMPO_REATUALIZAR	= 5
COD_BACKUP = formatNumber(0&Request.QueryString("COD_BACKUP"),0)
NOME_PASTA = Request.QueryString("NOME_PASTA")
NOME_BASE  = Request.QueryString("NOME_BASE")


If COD_BACKUP = 0 Then
	Response.Redirect("http://stillos.com.br/gerencial_foto?msg=Falha no backup!")
End If

'////////////////////////////////////
'CONEXÃO COM A BASE AUXILIAR
Dim objConexaoBACK

Sub AbrirConexaoBACK()
Set objConexaoBACK = Server.CreateObject("ADODB.Connection")
	objConexaoBACK.Open("DRIVER={MySQL ODBC 5.1 Driver};SERVER=187.45.196.164;PORT=3306;DATABASE=stillos4;USER=stillos4;PASSWORD=fca123fca123@;")
End Sub


'Sub Rotina que Fecha a Conexao
Sub FecharConexaoBACK()
   objConexaoBACK.Close
   Set objConexaoBACK=nothing
End Sub


'////////////////////////////////////
%>



<%
'/////////////////
'2ª parte registros importacao e criação do TXT
QUEBRA_LINHA = "$####$"

AbrirConexaoBACK()
AbrirConexao()



'///////////
'TELE VISUALIZAÇÃO DO BACKUP

Set rsTela = objConexaoBACK.Execute("SELECT NOME_TABELA,IMPORTADA,IMPORTADA_ESTRUTURA  FROM TB_BACKUP_TABELAS_TEMPORARIA WHERE COD_BACKUP = "&COD_BACKUP)%>
	
	<%
	If not rsTela.Eof Then
	%>
	
	<table width="80%" border="0" align="center" cellpadding="4" cellspacing="0" style="border:dashed #999999 1px">
	  <tr>
    <td colspan="2" style="border-bottom:solid 3px #666666"><div align="center" class="style1">BACKUP <%=COD_BACKUP%> / <%=UCASE(NOME_BASE)%></div>      <div align="center"></div></td>
    </tr>
	  <tr>
    <td width="696" bgcolor="#CCCCCC"><div align="center"><strong>Tabela </strong></div></td>
    <td width="271" bgcolor="#CCCCCC"><div align="center"><strong>Status</strong></div></td>
    </tr>
	<%
	While not rsTela.Eof
	
		Dim STATUS_TABLE,STATUS_TABLE2, MSG_STATUS
		STATUS_TABLE = formatNumber(0&rsTela("IMPORTADA"),0)
		STATUS_TABLE2 = formatNumber(0&rsTela("IMPORTADA_ESTRUTURA"),0)
	
		If STATUS_TABLE = 1 and STATUS_TABLE2 = 1 Then
		MSG_STATUS = "<span class='style4'>Importada</span>"
	Elseif 	STATUS_TABLE = 2 and STATUS_TABLE2 = 1 Then
		MSG_STATUS = "<a name='ANCORA'></a>Estrutura importada, importando registros"
	Elseif 	STATUS_TABLE = 2 and STATUS_TABLE2 = 2 Then
		MSG_STATUS = "Não importada"
	End If
		
%>

	
  <tr>
    <td bgcolor="#FFFFFF" style="border-bottom: solid 2px #CCCCCC"><div align="left"><%=rsTela("NOME_TABELA")%></div></td>
    <td bgcolor="#FFFFFF" style="border-bottom: solid 2px #CCCCCC">
    <div align="center"><%=MSG_STATUS%></div></td>
    </tr>
<%

	rsTela.MoveNext()
	Wend
	%>
	</table>
	<%
	End If
	
rsTela.Close
Set rsTela = Nothing
''///////////

	
			Set rs = objConexaoBACK.Execute("SELECT * FROM TB_BACKUP_TABELAS_TEMPORARIA WHERE IMPORTADA = 2 AND COD_BACKUP = "&COD_BACKUP&" LIMIT 1")			
					
					If not rs.Eof Then
					
					Dim NOME_TABELA, QTD_REGISTROS_TOTAL, TIPO_IMPORTACAO, IMPORTADA_ESTRUTURA
					Dim ENGINE, AUTO_INCREMENT, CRIADA_ORIGINALMENTE, ULTIMA_ALTERACAO, COLLATION, COMENTARIO
					Dim LINHA_INICIO, LINHA_FIM
					
					NOME_TABELA			= rs("NOME_TABELA")
					QTD_REGISTROS_TOTAL	= formatNumber(0&rs("QTD_REGISTROS_TOTAL"),0)
					TIPO_IMPORTACAO		= formatNumber(0&rs("TIPO_IMPORTACAO"),0)
					IMPORTADA_ESTRUTURA	= formatNumber(0&rs("IMPORTADA_ESTRUTURA"),0)
					ENGINE				= ""&rs("ENGINE")
					AUTO_INCREMENT		= formatNumber(0&rs("AUTO_INCREMENT"),0)
					CRIADA_ORIGINALMENTE= ""&rs("CRIADA_ORIGINALMENTE")
					ULTIMA_ALTERACAO	= ""&rs("ULTIMA_ALTERACAO")
					COLLATION			= ""&rs("COLLATION")
					COMENTARIO			= ""&rs("COMENTARIO")
					LINHA_INICIO		= formatNumber(0&rs("LINHA_INICIO"),0)
					LINHA_FIM			= formatNumber(0&rs("LINHA_FIM"),0)
					
					
					If CRIADA_ORIGINALMENTE <> "" Then
				
				'Response.Write("---"&CRIADA_ORIGINALMENTE&"<br />")
				CRIADA_ORIGINALMENTE  = Year(CRIADA_ORIGINALMENTE)&"-"&Month(CRIADA_ORIGINALMENTE)&"-"&Day(CRIADA_ORIGINALMENTE)&" "&Hour(CRIADA_ORIGINALMENTE)&":"&Minute(CRIADA_ORIGINALMENTE)&":"&Second(CRIADA_ORIGINALMENTE)
				
			End If
			
			If ULTIMA_ALTERACAO <> "" Then
				'Response.Write("---"&ULTIMA_ALTERACAO&"<br />")
				ULTIMA_ALTERACAO  = Year(ULTIMA_ALTERACAO)&"-"&Month(ULTIMA_ALTERACAO)&"-"&Day(ULTIMA_ALTERACAO)&" "&Hour(ULTIMA_ALTERACAO)&":"&Minute(ULTIMA_ALTERACAO)&":"&Second(ULTIMA_ALTERACAO)				
			End If
						
					If IMPORTADA_ESTRUTURA = 2 Then
					'Importa a estrutura
							
							'COLUNAS///////////////////////////////////////////////////////////////////
							Dim ESTRUTURA_COMPLETA
							
							ESTRUTURA_COMPLETA = "CREATE TABLE "&NOME_TABELA&" ( "&QUEBRA_LINHA
	
							
							Dim x, qtdChavesMult
							x = 0
							qtdChavesMult = 0
							
							Dim NOME_CAMPO_AUTO_INCREMENTE, CHAVE_PRIMARIA, CHAVES_MULTIPLAS
							
							CHAVE_PRIMARIA		= ""
							CHAVES_MULTIPLAS	= ""
							
							
							Set rs4 = objConexao.Execute("SHOW COLUMNS FROM "&NOME_TABELA)
									
									
									While not rs4.Eof
									x = x + 1
									
										Dim NOME_CAMPO, TIPO_CAMPO, NULLO, CHAVE, PADRAO, ALGO_EXTRA
								
										NOME_CAMPO	= ""&rs4(0)
										TIPO_CAMPO	= ""&rs4(1)
										NULLO		= ""&rs4(2)
										CHAVE		= Ucase(""&rs4(3))
										PADRAO		= ""&rs4(4)
										ALGO_EXTRA	= Ucase(""&rs4(5))
											
										'///
										If NULLO = "NO" Then
											NULLO = "NOT NULL"
										Else
											NULLO = "NULL"
										End If
										
										'///
										If ALGO_EXTRA = "AUTO_INCREMENT" Then
												NOME_CAMPO_AUTO_INCREMENTE = NOME_CAMPO
										Else										
												ALGO_EXTRA = ""
										End If
										
										'///
										If PADRAO <> "" Then												
												PADRAO = "default "&PADRAO
										End If
											
										'///
										If CHAVE = "PRI" Then
											CHAVE_PRIMARIA = ", PRIMARY KEY  ("&NOME_CAMPO&")"&QUEBRA_LINHA
										Elseif CHAVE = "MUL" Then
											qtdChavesMult = qtdChavesMult + 1
											
											If qtdChavesMult = 1 Then
											'Sem virgula no início
											'INICÍO
												CHAVES_MULTIPLAS = "KEY "&NOME_CAMPO&" ("&NOME_CAMPO&")"
											Else
											'Com virgula no início
												CHAVES_MULTIPLAS = CHAVES_MULTIPLAS&", KEY "&NOME_CAMPO&" ("&NOME_CAMPO&")"
											End If
											
											
										End If
																					
										If x = 1 Then
											'Sem virgula no início
											ESTRUTURA_COMPLETA = ESTRUTURA_COMPLETA & NOME_CAMPO&" "&TIPO_CAMPO&" "&NULLO&" "&Lcase(ALGO_EXTRA)&" "&PADRAO&QUEBRA_LINHA
											Else
											'Com virgula no início
											ESTRUTURA_COMPLETA = ESTRUTURA_COMPLETA &", "&NOME_CAMPO&" "&TIPO_CAMPO&" "&NULLO&" "&Lcase(ALGO_EXTRA)&" "&PADRAO&QUEBRA_LINHA	
										End If
									
									rs4.MoveNext()
									Wend
								
								
	
							
							Set rs4 = Nothing
							
							'/////////////////////////////////
							'CHAVES
							If CHAVE_PRIMARIA <> "" And CHAVES_MULTIPLAS <> "" Then
								'Chave primária e indíces
								ESTRUTURA_COMPLETA = ESTRUTURA_COMPLETA&CHAVE_PRIMARIA&", "&CHAVES_MULTIPLAS
								
							Elseif CHAVE_PRIMARIA = "" And CHAVES_MULTIPLAS <> "" Then
								'Somente indíces
								ESTRUTURA_COMPLETA = ESTRUTURA_COMPLETA&","&CHAVES_MULTIPLAS
								
							Elseif CHAVE_PRIMARIA <> "" And CHAVES_MULTIPLAS = "" Then
								'Somente chave primária
								ESTRUTURA_COMPLETA = ESTRUTURA_COMPLETA&CHAVE_PRIMARIA
							End If
							
							 
							'///////////////////////////////////////////////////////////
							'FINALIZAÇÃO ESTRUTURA
							ESTRUTURA_COMPLETA = ESTRUTURA_COMPLETA &QUEBRA_LINHA& ") ENGINE="&ENGINE
							
							If COLLATION <> "" Then
								ESTRUTURA_COMPLETA = ESTRUTURA_COMPLETA &" DEFAULT CHARSET=latin1 COLLATE="&COLLATION
							End If
							
							If AUTO_INCREMENT > 0 Then
								ESTRUTURA_COMPLETA = ESTRUTURA_COMPLETA &" AUTO_INCREMENT= "&AUTO_INCREMENT
							End If
							
							If COMENTARIO <> "" Then
								ESTRUTURA_COMPLETA = ESTRUTURA_COMPLETA &" COMMENT= '"&COMENTARIO&"'"
							End If
													
							ESTRUTURA_COMPLETA = ESTRUTURA_COMPLETA &";"							
												
							
							'TXT///////////////////////////////////////////////////////////////////////
							Set Fso = CreateObject("Scripting.FileSystemObject") 
							Set Cria = Fso.CreateTextFile("e:\home\stillos\backups\"&NOME_PASTA&"\TABLES\"&NOME_TABELA&".txt",True) 
							 
							
								Cria.WriteLine "-- Backup realizado em: "&Now() 
								Cria.WriteLine "-- Qtd Registros: "&QTD_REGISTROS_TOTAL
								Cria.WriteLine "-- Base de dados: "&NOME_BASE
								Cria.WriteLine "-- Tabela: "&NOME_TABELA
								
									If TIPO_IMPORTACAO = 1 Then
										Cria.WriteLine "-- Tipo Importação: Estrutura e Dados"
									Elseif TIPO_IMPORTACAO = 2 Then
										Cria.WriteLine "-- Tipo Importação: Somente Estrutura"
									End If
								
								
								Cria.WriteLine "-- ------------------------------------------------------"
								Cria.WriteLine "-- Dados CSV"
								Cria.WriteLine "-- BASE_DE_DADOS,TABELA,QTD_REGISTROS,DATA_CRIACAO_TABELA,DATA_ULTIMA_MOVIMENTACAO,TIPO_IMPORTACAO"
								Cria.WriteLine "-- "&NOME_BASE&","&NOME_TABELA&","&QTD_REGISTROS_TOTAL&","&CRIADA_ORIGINALMENTE&","&ULTIMA_ALTERACAO&","&TIPO_IMPORTACAO 
								Cria.WriteLine "-- ------------------------------------------------------"
								Cria.WriteLine ""
								Cria.WriteLine ""
								Cria.WriteLine "-- ESTRUTURA:"
								Cria.WriteLine ""
								Cria.WriteLine ""
								
								
								'Quebra as linhas da estrutura
								Dim linhas_grava
								linhas_grava = Split(ESTRUTURA_COMPLETA,"$####$")
								
								For i=0 to UBound(linhas_grava)
									Cria.WriteLine linhas_grava(i)
								Next
								
								Cria.WriteLine ""
								Cria.WriteLine ""
								Cria.WriteLine ""

							 
							'Fecha o arquivo criado. 
							Cria.Close 
							 
							'Destruíndo objetos 
							Set Fso = nothing 
							Set Cria = nothing 
							'TXT///////////////////////////////////////////////////////////////////////	
							
							'FINAL ESTRUTURA
							
							
							If TIPO_IMPORTACAO = 2 or QTD_REGISTROS_TOTAL = 0 Then
							'ATENÇÃO SE FOR SOMENTE ESTRUTURA FINALIZA AQUI O BACKUP DA TABELA
								sql = "UPDATE TB_BACKUP_TABELAS_TEMPORARIA SET IMPORTADA_ESTRUTURA = '1', IMPORTADA = '1' WHERE NOME_TABELA ='"&NOME_TABELA&"' AND COD_BACKUP = "&COD_BACKUP  
							Else
							'APÓS ESTRUTURA RODA DE NOVO PARA OS REGISTROS
								sql = "UPDATE TB_BACKUP_TABELAS_TEMPORARIA SET IMPORTADA_ESTRUTURA = '1' WHERE NOME_TABELA ='"&NOME_TABELA&"' AND COD_BACKUP = "&COD_BACKUP  				
							End If
												
							
							
							objConexaoBACK.Execute(sql)							
							
					'Fim estrutura	
					Elseif IMPORTADA_ESTRUTURA = 1 Then
					'Importação dos registros
							'REGISTROS///////////////////////////////////////////////////////////////////////
							'Início
							Dim QTD_COLUNAS, VALUES_INSERT
							QTD_COLUNAS = 0
							VALUES_INSERT = ""
							Set rs4 = objConexao.Execute("SHOW COLUMNS FROM "&NOME_TABELA,QTD_COLUNAS)
								
									
								
							rs4.Close
							Set rs4 = Nothing
							

							
							'1) Atualizada o número de linhas
							'2) Faz o for com select em relação as linhsa
							'3) Insere no arquivo txt
							'4) Verifica quando for a última linha fecha a atualização e vai para a próxima tabela
							
							Dim LINHA_INICIO_PROX, LINHA_FIM_PROX
							
							If LINHA_FIM = 0 Then
								LINHA_INICIO_PROX	= 0
							Else
								LINHA_INICIO_PROX	= (LINHA_FIM+1)/1
							End If
																			
							
							LINHA_FIM_PROX		= (LINHA_FIM+QTD_LINHAS_POR_VEZ)/1	
							
							'1)--------
							'Ou seja a linha inicio do próximo será a linha fim do anterior mais 1 e a linha fim do próximo será a linha fim do anterior + a quantidade
							'de linhas por vez
							sql =	" UPDATE TB_BACKUP_TABELAS_TEMPORARIA SET LINHA_INICIO = "&LINHA_INICIO_PROX&", LINHA_FIM = "&LINHA_FIM_PROX&_
									" WHERE NOME_TABELA ='"&NOME_TABELA&"' AND COD_BACKUP = "&COD_BACKUP
							
							'Response.Write(sql)
							objConexaoBACK.Execute(sql)

							'2)--------
							Dim qtdRegistrosReais
							qtdRegistrosReais = 0
							
							sql = "SELECT * FROM "&NOME_TABELA&" LIMIT "&LINHA_INICIO_PROX&","&QTD_LINHAS_POR_VEZ
							'Response.Write(sql)							
							Set rsR = objConexao.Execute(sql,qtdRegistrosReais)
									
									If int(qtdRegistrosReais) > 0 Then
																		'Ainda a linhas a serem inseridas
									
										'Abre o TXT
										'O parâmetro forReading é utilizado somente para leitura. 
										Const forReading = 1, forWriting = 2, forAppending = 8
 										Set Fso = CreateObject("Scripting.FileSystemObject")
										Set ArqTxt = Fso.OpenTextFile("e:\home\stillos\backups\"&NOME_PASTA&"\TABLES\"&NOME_TABELA&".txt",8,True)
											
											If LINHA_INICIO_PROX = 0 Then
												ArqTxt.WriteLine("-- REGISTROS:")
												ArqTxt.WriteLine("")
												ArqTxt.WriteLine("")
											End If
											
											Dim CAMPO_INSERE
											
												For a=0 to QTD_LINHAS_POR_VEZ
									
													If not rsR.Eof Then												
																													
															
															For l=0 to QTD_COLUNAS-1
															
															'3)--------
															CAMPO_INSERE = rsR(l)
															
															'Verifica se o campo é Data, Número ou Texto normal
															
															If IsDate(CAMPO_INSERE)	Then
																	
																	Dim DD,MM,AA,HH,MN,SS
																	DD = Day(CAMPO_INSERE)
																	MM = Month(CAMPO_INSERE)
																	AA = Year(CAMPO_INSERE)
																	HH = Hour(CAMPO_INSERE)
																	MN = Minute(CAMPO_INSERE)
																	SS = Second(CAMPO_INSERE)
																	
																	CAMPO_INSERE = "'"&AA&"-"&MM&"-"&DD&" "&HH&":"&MN&":"&SS&"'"
																	
																	
																																																		
															ElseIf isNumeric(CAMPO_INSERE) Then
																	
																	CAMPO_INSERE = sqlNumero(CAMPO_INSERE)
															Else
																	CAMPO_INSERE = sqlTexto(CAMPO_INSERE)
															End If
															
																If l=0 Then
																	VALUES_INSERT = CAMPO_INSERE
																Else
																	VALUES_INSERT = VALUES_INSERT&","&CAMPO_INSERE
																End If
															Next
															
															
															ArqTxt.WriteLine("INSERT INTO "&NOME_TABELA&" VALUES ("&VALUES_INSERT&");")
													
														rsR.MoveNext()
																									
													End If
												
												Next
		
																				
										ArqTxt.Close
 										Set ArqTxt = nothing									
										
										
									Elseif int(qtdRegistrosReais) = 0 Then									
										'Acabou fecha o TXT a Tabela e vai para a próxima tabela depois que atualizar novamente
										'4)--------
										sql = "UPDATE TB_BACKUP_TABELAS_TEMPORARIA SET IMPORTADA = '1' WHERE NOME_TABELA ='"&NOME_TABELA&"' AND COD_BACKUP = "&COD_BACKUP  			
										objConexaoBACK.Execute(sql)						
																		
									End If
							
							rsR.Close
							Set rsR = Nothing
					
		
							'Fim			
							'REGISTROS///////////////////////////////////////////////////////////////////////
						
										
						End If
										
						CONTINUA_BACKUP = 1
						
					Else
						
						CONTINUA_BACKUP = 2					
						'/////////////////////////////////////////
								'Limpa a base temporária
								objConexaoBACK.Execute("DELETE FROM TB_BACKUP_TABELAS_TEMPORARIA WHERE COD_BACKUP="&COD_BACKUP)
						'/////////////////////////////////////////
			
									

					End If
		rs.Close
		Set rs = Nothing
		

FecharConexao()
FecharConexaoBACK()
%>


<%If CONTINUA_BACKUP = 1 Then%>
	<script language="javascript">
	var intervalo = window.setInterval("redirecionarPag()",<%=TEMPO_REATUALIZAR*1000%>);
		function redirecionarPag(){
		<%If Request.ServerVariables("HTTPS") = "on" Then%>
			location.href='https://stillos.com.br/gerencial_foto/adm_back_up_parte2.asp?COD_BACKUP=<%=COD_BACKUP%>&NOME_PASTA=<%=NOME_PASTA%>&NOME_BASE=<%=NOME_BASE%>';
		<%Else%>
			location.href='https://stillos.com.br/gerencial_foto/adm_back_up_parte2.asp?COD_BACKUP=<%=COD_BACKUP%>&NOME_PASTA=<%=NOME_PASTA%>&NOME_BASE=<%=NOME_BASE%>';
		<%End If%>
		}
	</script>
<%Else 
		AbrirConexaoBACK()
		'BAIXA COM A DATA DO TÉRMINO DO BACKUP
			Set rs=objConexaoBACK.Execute("SELECT COD_BACKUP FROM TB_BACKUPS WHERE COD_BACKUP = "&COD_BACKUP&" AND DATA_TERMINO IS NULL")
			
			If not rs.eof Then
				objConexaoBACK.Execute("UPDATE TB_BACKUPS SET DATA_TERMINO = NOW() WHERE COD_BACKUP ="&COD_BACKUP&" LIMIT 1")	
			End If
			
			rs.Close
			Set rs=Nothing
		FecharConexaoBACK()
	
	Response.Redirect("adm_back_up_parte3.asp?COD_BACKUP="&COD_BACKUP&"&NOME_BASE="&NOME_BASE)	
End If%>
</body>
</html>



