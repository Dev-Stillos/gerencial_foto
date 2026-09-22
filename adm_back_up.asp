<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<%
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

Dim rs, sql
Dim NOME_BASE, COD_BACKUP
NOME_BASE = "stillos3"


AbrirConexaoBACK()
AbrirConexao()

	'IMPORTANTE - ATUALIZA ESTRUTURA DAS VIEWS!!!!!!!!!!!!!!!!!!!
	objConexao.Execute("call pro_backup_views('1','BK','EXECUTE_DA_MESMA_FORMA')")
	'-----------------------------------------------------------
	
	objConexaoBACK.Execute("INSERT INTO TB_BACKUPS(DATA,NOME_BASE) VALUES (CURRENT_TIMESTAMP, '"&NOME_BASE&"')")

	Set rsI = objConexaoBACK.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_BACKUPS LIMIT 1")
			
			If not rsI.Eof Then
				
					COD_BACKUP = rsI("ULTIMO_ID")
				
			End If
			
	rsI.Close
	Set rsI = Nothing
	
	'/////////////////////////////
	
	
	'Cria a pasta do backup

			Dim DIA, MES
			
			DIA = Day(Now)
			MES = Month(Now)
			
			If Len(DIA) = 1 Then
				DIA = "0"&DIA
			End If
			
			If Len(MES) = 1 Then
				MES = "0"&MES
			End If
			
			
			If len(COD_BACKUP) = 5 Then
				COD_BACKUP = "0"&COD_BACKUP
			Elseif len(COD_BACKUP) = 4 Then
				COD_BACKUP = "00"&COD_BACKUP
			Elseif len(COD_BACKUP) = 3 Then
				COD_BACKUP = "000"&COD_BACKUP
			Elseif len(COD_BACKUP) = 2 Then
				COD_BACKUP = "0000"&COD_BACKUP
			Elseif len(COD_BACKUP) = 1 Then
				COD_BACKUP = "00000"&COD_BACKUP
			End If
			
			NOME_PASTA = Ucase("BACKUP_"&COD_BACKUP&"_"&NOME_BASE&"_"&DIA&"_"&MES&"_"&Year(Now))
			NOME_PASTA = Replace(NOME_PASTA," ","")
			
			
			'Pasta inicial do Backup
			Dim fs,f
			set fs=Server.CreateObject("Scripting.FileSystemObject")
			set f=fs.CreateFolder("e:\home\stillos\backups\"&NOME_PASTA)
			set f=fs.CreateFolder("e:\home\stillos\backups\"&NOME_PASTA&"\TABLES")
			set f=fs.CreateFolder("e:\home\stillos\backups\"&NOME_PASTA&"\VIEWS")
			set f=fs.CreateFolder("e:\home\stillos\backups\"&NOME_PASTA&"\PROGRAMA")
			'/////////////////////////////
				'ARQUIVOS DO SITE PASTA "PROGRAMA"
				fs.CopyFolder "e:\home\stillos\Web\gerencial_foto","e:\home\stillos\backups\"&NOME_PASTA&"\PROGRAMA\"
				'fs.CopyFolder "e:\home\stillos\Web\gerencial_foto\*","e:\home\stillos\backups\"&NOME_PASTA&"\PROGRAMA\"
			'/////////////////////////////
			set f=nothing
			set fs=nothing
	
	
			'///////////////////////////////
			'ARQUIVO INSTRUÇÕES DO BACK UP
								Set Fso = CreateObject("Scripting.FileSystemObject")	
								set Cria = Fso.CreateTextFile("e:\home\stillos\backups\"&NOME_PASTA&"\IMPORTANTE_INSTRUCOES.txt",True) 
							 
							
										Cria.WriteLine "-- Backup realizado em: "&Now() 
										Cria.WriteLine "-- Base de dados: "&NOME_BASE
										Cria.WriteLine "-- View: "&NOME_TABELA																					
										Cria.WriteLine "-- ------------------------------------------------------"
										Cria.WriteLine ""
										Cria.WriteLine ""
										Cria.WriteLine "--------- Instruções do Back-Up"
										Cria.WriteLine ""
										Cria.WriteLine "Os arquivos do backup foram dividos em 3 pastas:"
										Cria.WriteLine ""
										Cria.WriteLine ""
										Cria.WriteLine ""
										Cria.WriteLine "PROGRAMA:"
										Cria.WriteLine ""
										Cria.WriteLine "	Contém todos os arquivos do sistema ASP, jpg, css, etc..."
										Cria.WriteLine "	Esta pasta deverá ser colocada na raiz da web do servidor com o nome gerencial_foto"
										Cria.WriteLine "	Caso tem ocorrido a transferência de hospedagem deverá ser alterado também o caminho fisíco"
										Cria.WriteLine "	destes arquivos e também o desbloqueio para leitura (mais informações consulte o manual do sistema sobre isso)"
										Cria.WriteLine ""
										Cria.WriteLine "	1º Passo caso necessário backup dos arquivos, publique da mesma forma"
										Cria.WriteLine ""
										Cria.WriteLine ""
										Cria.WriteLine "TABLES:"
										Cria.WriteLine ""
										Cria.WriteLine "	Contém todas as tabelas e respectivos registros,chaves, index, etc... Atuais de acordo com a data do backup"
										Cria.WriteLine ""	
										Cria.WriteLine "	2º Passo caso necessário backup da estrutura execute as sqls"
										Cria.WriteLine ""
										Cria.WriteLine "VIEWS:"
										Cria.WriteLine "	Contém as views do software já devidamente ajustada para qualquer base de dados Mysql que contenha as tabelas das"
										Cria.WriteLine "	respectivas Views"
										Cria.WriteLine ""
										Cria.WriteLine "	3º Após a inclusão de todas as tabelas execute as sqls das VIEWS"
										Cria.WriteLine ""
										Cria.WriteLine ""
										Cria.WriteLine "-----------OBSERVAÇÃO IMPORTANTE-----------"
										Cria.WriteLine ""
										Cria.WriteLine "Crie esta procedure abaixo para o próximo backUp:"
										Cria.WriteLine ""
										Cria.WriteLine ""
										Cria.WriteLine "--Início procedure"
										Cria.WriteLine ""
										Cria.WriteLine "DROP PROCEDURE pro_backup_views//"
										Cria.WriteLine "CREATE PROCEDURE pro_backup_views( "
										Cria.WriteLine "	IN var_id TINYINT,"
										Cria.WriteLine "	   var_uf CHAR(2),"
										Cria.WriteLine "	   var_nome CHAR(30))"
										Cria.WriteLine "BEGIN"
										Cria.WriteLine "        TRUNCATE TB_BACK_UP_VIEWS;"
										Cria.WriteLine "	INSERT INTO TB_BACK_UP_VIEWS SELECT TABLE_NAME,VIEW_DEFINITION FROM VW_BACKUP_VIEWS; "
										Cria.WriteLine "END"
										Cria.WriteLine "--Fim procedure"
										Cria.WriteLine ""
										Cria.WriteLine ""
										Cria.WriteLine "/////////////////////////////////////////////////////////////////////////////////////"
										Cria.WriteLine ""
										Cria.WriteLine ""
										Cria.WriteLine "-- Arquivos para efetuar o BackUp:"
										Cria.WriteLine ""
										Cria.WriteLine "Os arquivos ""adm_back_up.asp"",""adm_back_up_parte2.asp"" e ""adm_back_up_parte3.asp""  que se encontram na pasta ""PROGRAMAS/gerencial_foto"" que efetuam o backUp"
										Cria.WriteLine ""
										Cria.WriteLine ""
										Cria.WriteLine "-- Tabelas que controlam o BackUp:"
										Cria.WriteLine ""
										Cria.WriteLine "CREATE TABLE TB_BACKUPS ("								
										Cria.WriteLine "  COD_BACKUP int(5) NOT NULL AUTO_INCREMENT,"
										Cria.WriteLine "  DATA timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,"
										Cria.WriteLine "  DATA_TERMINO datetime DEFAULT NULL,"
										Cria.WriteLine "  NOME_BASE varchar(255) COLLATE latin1_general_ci NOT NULL,"
										Cria.WriteLine "  PRIMARY KEY (COD_BACKUP)"
										Cria.WriteLine ") ENGINE=MyISAM  DEFAULT CHARSET=latin1 COLLATE=latin1_general_ci AUTO_INCREMENT=0;"
										Cria.WriteLine ""
										Cria.WriteLine ""
										Cria.WriteLine "CREATE TABLE TB_BACKUP_TABELAS_TEMPORARIA ("
										Cria.WriteLine "COD_BACKUP int(10) NOT NULL,"
										Cria.WriteLine "NOME_TABELA varchar(255) COLLATE latin1_general_ci NOT NULL,"
										Cria.WriteLine "ENGINE varchar(255) COLLATE latin1_general_ci NOT NULL,"
										Cria.WriteLine "AUTO_INCREMENT int(10) NOT NULL,"
										Cria.WriteLine "CRIADA_ORIGINALMENTE datetime NOT NULL,"
										Cria.WriteLine "ULTIMA_ALTERACAO datetime NOT NULL,"
										Cria.WriteLine "COLLATION varchar(255) COLLATE latin1_general_ci NOT NULL,"
										Cria.WriteLine "COMENTARIO varchar(255) COLLATE latin1_general_ci NOT NULL,"
										Cria.WriteLine "QTD_REGISTROS_TOTAL int(10) NOT NULL,"
										Cria.WriteLine "TIPO_IMPORTACAO smallint(1) NOT NULL,"
										Cria.WriteLine "IMPORTADA smallint(1) NOT NULL DEFAULT '2',"
										Cria.WriteLine "IMPORTADA_ESTRUTURA smallint(1) NOT NULL DEFAULT '2',"
										Cria.WriteLine "LINHA_INICIO int(15) NOT NULL DEFAULT '0',"
										Cria.WriteLine "LINHA_FIM int(15) NOT NULL DEFAULT '0',"
										Cria.WriteLine "KEY COD_BACKUP (COD_BACKUP,IMPORTADA)"
										Cria.WriteLine ") ENGINE=MyISAM DEFAULT CHARSET=latin1 COLLATE=latin1_general_ci;"
										Cria.WriteLine ""
										Cria.WriteLine ""
										Cria.WriteLine "/////////////////////////////////////////////////////////////////////////////////////"
										
										Cria.Close 
										Set Fso = nothing 
										Set Cria = nothing 
			
	
					'/////////////////////////////
					

	Set rs = objConexao.Execute("SHOW TABLE STATUS")%>
	
	<table width="100%" border="0" cellspacing="0" cellpadding="0">		
			
			<%While not rs.Eof
			
			Dim NOME_TABELA, QTD_REGISTROS_TOTAL, TIPO_IMPORTACAO
			Dim ENGINE, AUTO_INCREMENT, CRIADA_ORIGINALMENTE, ULTIMA_ALTERACAO, COLLATION, COMENTARIO
			
			
			NOME_TABELA			= rs(0)
			ENGINE				= ""&rs(1)
			AUTO_INCREMENT		= sqlNumero(formatNumber(0&rs(10),0)/1)
			CRIADA_ORIGINALMENTE= ""&rs(11)
			ULTIMA_ALTERACAO	= ""&rs(12)
			COLLATION			= ""&rs(14)
			COMENTARIO			= ""&rs(17)
			QTD_REGISTROS_TOTAL	= rs(4)
			
			
			
			If CRIADA_ORIGINALMENTE <> "" Then
				
				'Response.Write("---"&CRIADA_ORIGINALMENTE&"<br />")
				CRIADA_ORIGINALMENTE  = Year(CRIADA_ORIGINALMENTE)&"-"&Month(CRIADA_ORIGINALMENTE)&"-"&Day(CRIADA_ORIGINALMENTE)&" "&Hour(CRIADA_ORIGINALMENTE)&":"&Minute(CRIADA_ORIGINALMENTE)&":"&Second(CRIADA_ORIGINALMENTE)
				
			End If
			
			If ULTIMA_ALTERACAO <> "" Then
				'Response.Write("---"&ULTIMA_ALTERACAO&"<br />")
				
				ULTIMA_ALTERACAO  = Year(ULTIMA_ALTERACAO)&"-"&Month(ULTIMA_ALTERACAO)&"-"&Day(ULTIMA_ALTERACAO)&" "&Hour(ULTIMA_ALTERACAO)&":"&Minute(ULTIMA_ALTERACAO)&":"&Second(ULTIMA_ALTERACAO)				
			End If
			

			TIPO_IMPORTACAO = Right(NOME_TABELA,10)
			
			If TIPO_IMPORTACAO  = "TEMPORARIA" Then
				TIPO_IMPORTACAO = 2
			Else
				TIPO_IMPORTACAO = 1
			End If	
				
			
			
			If ENGINE <>  "" Then
			'BackUp das Bases de dados.
				
				
								
sql="INSERT INTO TB_BACKUP_TABELAS_TEMPORARIA (COD_BACKUP,NOME_TABELA,ENGINE,AUTO_INCREMENT,CRIADA_ORIGINALMENTE,ULTIMA_ALTERACAO,COLLATION,COMENTARIO,QTD_REGISTROS_TOTAL,TIPO_IMPORTACAO)VALUES("&COD_BACKUP&", '"&NOME_TABELA&"', '"&ENGINE&"', "&AUTO_INCREMENT&", '"&CRIADA_ORIGINALMENTE&"', '"&ULTIMA_ALTERACAO&"', '"&COLLATION&"', '"&COMENTARIO&"', "&QTD_REGISTROS_TOTAL&", "&TIPO_IMPORTACAO&")"
						
						
						Response.Write(sql)												
						objConexaoBACK.Execute(sql)
									
			Else
			'CRIA A VIEW BACKUP
						
						sql = "SELECT * FROM TB_BACK_UP_VIEWS WHERE NOME_VIEW = '"&NOME_TABELA&"'"
						
						Set rsView = objConexao.Execute(sql)
									
									If not rsView.Eof Then
										
										
										Dim VIEW_TEXT
								
									
										VIEW_TEXT = UCase(rsView(1))									
										VIEW_TEXT = Replace(VIEW_TEXT,"`","")
										VIEW_TEXT = Replace(VIEW_TEXT,"'","")
										VIEW_TEXT = Replace(VIEW_TEXT,"´","")
										VIEW_TEXT = Replace(VIEW_TEXT,"*","")
										VIEW_TEXT = Replace(VIEW_TEXT,"/","")
										VIEW_TEXT = Replace(VIEW_TEXT,"\","")
										VIEW_TEXT = Replace(VIEW_TEXT,Ucase(NOME_BASE)&".","")
										VIEW_TEXT = Replace(VIEW_TEXT,Ucase(NOME_BASE),"")
										VIEW_TEXT = Replace(VIEW_TEXT,"ALGORITHM=UNDEFINED","CREATE VIEW AS "&rsView(0))
										
										
																				
										'TXT VIEW
										Set Fso = CreateObject("Scripting.FileSystemObject")	
										Set Cria = Fso.CreateTextFile("e:\home\stillos\backups\"&NOME_PASTA&"\VIEWS\"&NOME_TABELA&".txt",True) 
							 
							
										Cria.WriteLine "-- Backup realizado em: "&Now() 
										Cria.WriteLine "-- Base de dados: "&NOME_BASE
										Cria.WriteLine "-- View: "&NOME_TABELA																					
										Cria.WriteLine "-- ------------------------------------------------------"
										Cria.WriteLine ""
										Cria.WriteLine ""
										Cria.WriteLine "-- ESTRUTURA VIEW:"
										Cria.WriteLine ""
										Cria.WriteLine ""
										Cria.WriteLine VIEW_TEXT&";"
								
									Cria.Close 
									Set Fso = nothing 
									Set Cria = nothing 
									'TXT///////////////////////////////////////////////////////////////////////	
											
																			
									End If
						
						
						rsView.Close
						Set rsView = Nothing
			
			
			%>		
			</table>
			
<%	
			End If
			
			rs.MoveNext()
			Wend
					
rs.Close
Set rs = Nothing

FecharConexaoBACK()
FecharConexao()

Response.Redirect("adm_back_up_parte2.asp?COD_BACKUP="&COD_BACKUP&"&NOME_PASTA="&NOME_PASTA&"&NOME_BASE="&NOME_BASE)

%>



