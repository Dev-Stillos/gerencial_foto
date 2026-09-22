<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<%
Dim COD_PRODUCAO
COD_PRODUCAO = Request.QueryString("cod_prod")
Dim DESCRICAO


AbrirConexao()
	Dim rsDes
	Set rsDes = objConexao.Execute("SELECT DESCRICAO FROM TB_PRODUCAO WHERE COD_PRODUCAO="&COD_PRODUCAO&" LIMIT 1")
	
		DESCRICAO=rsDes("DESCRICAO")
		
	rsDes.Close()
	Set rsDes = Nothing

Dim COD_CONTRATO1, COD_CONTRATO2, COD_CONTRATO3, COD_CONTRATO4, COD_CONTRATO5, COD_CONTRATO6
'1) Resgastada osCOD_CONTRATOs, informados
'2) Consulta osCOD_CONTRATOs na base e verifica se n�o h� formandos com oCOD_CONTRATO informado

COD_CONTRATO1 = ""&Request.Form("COD_CONTRATO1")
COD_CONTRATO2 = ""&Request.Form("COD_CONTRATO2")
COD_CONTRATO3 = ""&Request.Form("COD_CONTRATO3")
COD_CONTRATO4 = ""&Request.Form("COD_CONTRATO4")
COD_CONTRATO5 = ""&Request.Form("COD_CONTRATO5")
COD_CONTRATO6 = ""&Request.Form("COD_CONTRATO6")


Dim rsValida1, rsValida2, sql, rsCOD_CONTRATO, posicaoInicial

	sql = 	"SELECT * FROM TB_FORMANDOS_IMP WHERE "
			
			
			If Len(COD_CONTRATO1) > 1 Then				
				'Verifica se h� registro com esteCOD_CONTRATO j� em algum formando
						Dim rsVerificaImporta
						Set rsVerificaImporta = objConexao.Execute("SELECT COD_CONTRATO, COD_PRODUCAO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_CONTRATO ="&COD_CONTRATO1&" LIMIT 1")
						
							If rsVerificaImporta.Eof Then
								sql = sql&"CONTRATO = "&COD_CONTRATO1
							Else
								Dim arta
								arta = formatNumber(0&rsVerificaImporta("COD_PRODUCAO"),0)
								If arta = COD_PRODUCAO Then
									sql = sql&"CONTRATO = "&COD_CONTRATO1							
								Else
									Response.Redirect("dip_producoes_importar_formandos.asp?cod_prod="&COD_PRODUCAO&"&msg=Desculpe no momento da grava��o foi detectado que h� outra produ��o comCOD_CONTRATO selecionado, por gentileza inicie o processo novamente.")							
								End If
							End If
						
						rsVerificaImporta.Close()
						Set rsVerificaImporta = Nothing
			End If
			
				If Len(COD_CONTRATO2) > 1 Then				
				'Verifica se h� registro com esteCOD_CONTRATO j� em algum formando
						Set rsVerificaImporta = objConexao.Execute("SELECT COD_CONTRATO, COD_PRODUCAO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_CONTRATO ="&COD_CONTRATO2&" LIMIT 1")
						
							If rsVerificaImporta.Eof Then
								sql = sql&" OR CONTRATO = "&COD_CONTRATO2
							Else
								Response.Redirect("dip_producoes_importar_formandos.asp?cod_prod="&COD_PRODUCAO&"&msg=Desculpe no momento da grava��o foi detectado que h� outra produ��o comCOD_CONTRATO selecionado, por gentileza inicie o processo novamente.")
								
							End If
						
						rsVerificaImporta.Close()
						Set rsVerificaImporta = Nothing
			End If
			
			
			If Len(COD_CONTRATO3) > 1 Then				
				'Verifica se h� registro com esteCOD_CONTRATO j� em algum formando
						Set rsVerificaImporta = objConexao.Execute("SELECT COD_CONTRATO, COD_PRODUCAO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_CONTRATO ="&COD_CONTRATO3&" LIMIT 1")
						
							If rsVerificaImporta.Eof Then
								sql = sql&" OR CONTRATO = "&COD_CONTRATO3
							Else
								Response.Redirect("dip_producoes_importar_formandos.asp?cod_prod="&COD_PRODUCAO&"&msg=Desculpe no momento da grava��o foi detectado que h� outra produ��o comCOD_CONTRATO selecionado, por gentileza inicie o processo novamente.")
								
							End If
						
						rsVerificaImporta.Close()
						Set rsVerificaImporta = Nothing
			End If
			
			If Len(COD_CONTRATO4) > 1 Then				
				'Verifica se h� registro com esteCOD_CONTRATO j� em algum formando
						Set rsVerificaImporta = objConexao.Execute("SELECT COD_CONTRATO, COD_PRODUCAO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_CONTRATO ='"&COD_CONTRATO4&"' LIMIT 1")
						
							If rsVerificaImporta.Eof Then
								sql = sql&" OR CONTRATO = "&COD_CONTRATO4
							Else
								Response.Redirect("dip_producoes_importar_formandos.asp?cod_prod="&COD_PRODUCAO&"&msg=Desculpe no momento da grava��o foi detectado que h� outra produ��o comCOD_CONTRATO selecionado, por gentileza inicie o processo novamente.")
								
							End If
						
						rsVerificaImporta.Close()
						Set rsVerificaImporta = Nothing
			End If
			
			If Len(COD_CONTRATO5) > 1 Then				
				'Verifica se h� registro com esteCOD_CONTRATO j� em algum formando
						Set rsVerificaImporta = objConexao.Execute("SELECT COD_CONTRATO, COD_PRODUCAO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_CONTRATO ='"&COD_CONTRATO5&"' LIMIT 1")
						
							If rsVerificaImporta.Eof Then
								sql = sql&" OR CONTRATO = "&COD_CONTRATO5
							Else
								Response.Redirect("dip_producoes_importar_formandos.asp?cod_prod="&COD_PRODUCAO&"&msg=Desculpe no momento da grava��o foi detectado que h� outra produ��o comCOD_CONTRATO selecionado, por gentileza inicie o processo novamente.")
								
							End If
						
						rsVerificaImporta.Close()
						Set rsVerificaImporta = Nothing
			End If
			
			If Len(COD_CONTRATO6) > 1 Then				
				'Verifica se h� registro com esteCOD_CONTRATO j� em algum formando
						Set rsVerificaImporta = objConexao.Execute("SELECT COD_CONTRATO, COD_PRODUCAO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_CONTRATO ="&COD_CONTRATO6&" LIMIT 1")
						
							If rsVerificaImporta.Eof Then
								sql = sql&" OR CONTRATO = "&COD_CONTRATO6
							Else
								Response.Redirect("dip_producoes_importar_formandos.asp?cod_prod="&COD_PRODUCAO&"&msg=Desculpe no momento da grava��o foi detectado que h� outra produ��o comCOD_CONTRATO selecionado, por gentileza inicie o processo novamente.")
								
							End If
						
						rsVerificaImporta.Close()
						Set rsVerificaImporta = Nothing
			End If
			
			
			'Corrige por causa do primeiroCOD_CONTRATO
			sql = Replace(sql," WHERE  OR"," WHERE ")
			sql = Replace(sql," WHERE OR"," WHERE ")	
				
			'ESTE C�DIGO FOI UTILIZADO PARA GERAR UMA SEGUNDA PARTE DE UMA PRODU��O DA GV SOMENTE COM OS FORMANDOS NAO UTILIZADOS NA OUTRA
			'sql = "SELECT * FROM TB_FORMANDOS_IMP WHERE contrato = 2285 AND CONCAT(codigo_formando,'0') NOT IN(SELECT COD_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_CONTRATO = 2285)"
	
	Set rsValida1 = objConexao.Execute(sql)
		
			If not rsValida1.Eof Then
			'H� registros vamos para 2� valida��o
				Dim CONTA_IMPORTADOS, CONTA_NAO_IMPORTADOS
				
				CONTA_IMPORTADOS		= 0
				CONTA_NAO_IMPORTADOS	= 0
				
				'PARTE1
				Dim COD_FORMANDO, COD_CONTRATO, COD_CURSO, NOME_CURSO, NOME_FORMANDO, ANO_CONCLUSAO, SEMESTRE_CONCLUSAO
				'PARTE2
				Dim ENDERECO, CEP, FONE_RES, FONE_CEL, BAIRRO, CIDADE, ESTADO
				
				While not rsValida1.Eof					
				
				'Tb 1
																   '0 devido a importa��o da formatura
				COD_FORMANDO		= rsValida1("CODIGO_FORMANDO")&"0"
				COD_CONTRATO		= rsValida1("CONTRATO")
				COD_CURSO			= rsValida1("CODIGO_CURSO")
				NOME_CURSO			= rsValida1("NOME_CURSO")
				NOME_FORMANDO		= ""&rsValida1("NOME_FORMANDO")
				'Corre��o do nome do formando
				posicaoInicial = inStr(NOME_FORMANDO,"(")
				If posicaoInicial > 0 Then
					NOME_FORMANDO = Left(NOME_FORMANDO,posicaoInicial-1)
				End If

				ANO_CONCLUSAO		= rsValida1("ANO")
				SEMESTRE_CONCLUSAO	= rsValida1("SEMESTRE")
				'------------------------------------
				
				'Tb2
				ENDERECO			= rsValida1("ENDERECO")
				CEP					= rsValida1("CEP")
				CEP					= Replace(CEP,"-","")
				FONE_RES			= rsValida1("FONE-RES")
				FONE_CEL			= rsValida1("FONE-COM-CEL")
				BAIRRO				= rsValida1("BAIRRO")
				CIDADE				= rsValida1("MUNICIPIO")
				ESTADO				= rsValida1("ESTADO")
				'------------------------------------
				
				'AGORA BASTA VERIFICAR SE H� REGISTRO DO COD. FORMANDO ANTES DE DAR O INSERT!!!!!!
				Set rsValida2 = objConexao.Execute("SELECT COD_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO="&COD_FORMANDO&" LIMIT 1")
					
						If rsValida2.Eof Then
						'Grava pois n�o h� registro
						CONTA_IMPORTADOS = CONTA_IMPORTADOS + 1	
									
						sql = 	"INSERT INTO TB_FICHAS_IDENTIFICACAO ("&_
								"COD_FORMANDO , "&_
								"COD_PRODUCAO , "&_
								"COD_CONTRATO , "&_
								"COD_CURSO , "&_
								"NOME_CURSO , "&_
								"NOME_FORMANDO , "&_
								"ANO_CONCLUSAO , "&_
								"SEMESTRE_CONCLUSAO "&_
								")"&_
								"VALUES ("&_
								sqltexto(COD_FORMANDO)&", "&_
								sqltexto(COD_PRODUCAO)&", "&_
								COD_CONTRATO&", "&_
								COD_CURSO&", "&_
								sqltexto(NOME_CURSO)&", "&_
								sqltexto(NOME_FORMANDO)&", "&_
								sqltexto(ANO_CONCLUSAO)&", "&_
								sqltexto(Ucase(SEMESTRE_CONCLUSAO))&_							
								")"			
								'Tb 1
								objConexao.Execute(sql)
								'Response.Write(sql&"<br /><br /><br /><br />")
								'-------------
								
						sql = 	"INSERT INTO TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO ("&_
								"COD_FORMANDO ,"&_
								"ENDERECO ,"&_
								"CEP ,"&_
								"FONE_RES ,"&_
								"FONE_CEL ,"&_
								"BAIRRO ,"&_
								"CIDADE ,"&_
								"ESTADO ,"&_
								"CADASTRADO_POR "&_
								")"&_
								"VALUES ("&_
								COD_FORMANDO&", "&sqlTexto(ENDERECO)&", '"&CEP&"', '"&FONE_RES&"', '"&FONE_CEL&"', "&sqlTexto(BAIRRO)&", "&sqlTexto(CIDADE)&", "&_
								sqlTexto(ESTADO)&", "&Session("COD_USUARIO")&_
								")"
								
								'Tb 2
								objConexao.Execute(sql)
								'Response.Write(sql&"<br /><br /><br /><br />")
								'-------------		
									
						Else
						'N�o grava
						CONTA_NAO_IMPORTADOS = CONTA_NAO_IMPORTADOS + 1											
						End If
				
				rsValida2.Close()
				Set rsValida2 = Nothing
				
				rsValida1.MoveNext()
				Wend
						
			Else
			
				Response.Redirect("dip_producoes.asp?msg=Não foi possível importar formandos, pois não foi encontrado formandos relacionados aosCOD_CONTRATOs informados para produção "&COD_PRODUCAO&"-"&DESCRICAO&"!")
			
			
			End If
			
	rsValida1.Close()
	Set rsValida1 = Nothing



FecharConexao()

Response.Redirect("dip_producoes.asp?cod_prod="&COD_PRODUCAO&"&msg="&CONTA_IMPORTADOS&" formandos importados com sucesso!")
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
</head>

<body>
</body>
</html>