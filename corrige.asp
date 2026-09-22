<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<!-- #include file="includes/classe_fechamento.asp" -->
<%
abrirConexao()

'response.end

'IMPORTACAO MACKENZIE
dim cod_colacao_importacao

DIm COD_FORMANDO, SISTEMA_NOVO, COD_PRODUCAO,COD_CONTRATO, COD_CURSO, NOME_CURSO, NOME_FORMANDO,ANO_CONCLUSAO, SEMESTRE_CONCLUSAO, ALBUM, INICIAL,RETIRADO, BRINDE, COMISSAO, PRODUTO_AGREGADO

Dim ENDERECO, CEP, FONE_RES, FONE_CEL, BAIRRO, CIDADE, ESTADO

'PARAMETROS GERAIS 
'569 - MACKENZIE ADMINISTRAÇÃO DEZEMBRO 2015 EVENTO DO DIA 24/02/2016 - 118F - 22
'568 - MACKENZIE PUBLICIDADE PROPAGANDA DEZEMBRO 2015 EVENTO DO DIA 16/02/2016 - 117F - 17
'567 - MACKENZIE PUBLICIDADE MARKETING DEZEMBRO 2015 EVENTO DO DIA 15/02/2016 - 116F - 16
'--------------------------------------------------------------------------------------------------
'571 - MACKENZIE COMEX DEZEMBRO 2015 EVENTO DO DIA 19/02/2016 - 120F - 20
'572 - MACKENZIE ADM-COMEX DEZEMBRO 2015 EVENTO DO DIA 18/02/2016 - 119F - 19
'--------------------------------------------------------------------------------------------------
'574 - MACKENZIE ADM DEZEMBRO 2015-EVENTO DO DIA 17/02/2016 - 121F - 18
'575 - MACKENZIE ADM DEZEMBRO 2015 EVENTO DO DIA 23/02/2016 - 122F - 21
'--------------------------------------------------------------------------------------------------
'603 - MACKENZIE DIREITO JULHO 2016 EVENTO DO DIA 14/07/2016 - 131F - 29
'602 - MACKENZIE DIREITO JULHO 2016 EVENTO DO DIA 12/07/2016 - 130F - 27 
'--------------------------------------------------------------------------------------------------
'604 - MACKENZIE ENGENHARIA JULHO 2016 EVENTO DO DIA 21/07/2016 - 155F - 33
'605 - MACKENZIE ENGENHARIA JULHO 2016 EVENTO DO DIA 20/07/2016 - 156F - 32
'--------------------------------------------------------------------------------------------------
'611 - MACKENZIE DIREITO CAMPINAS TURMA JULHO 2016 EVENTO DO DIA 19/07/2016 - 157F - 31
'612 - MACKENZIE DESIGN TURMA JULHO 2016 EVENTO DO DIA 03/08/2016 - 158F - 40
'613 - MACKENZIE DIREITO CAMPINAS TURMA JULHO 2016 EVENTO DO DIA 19/0/7/2016 - 159F - 38

'--------------------------------------------------------------------------------------------------
'615 - MACKENZIE-ARQUITETURA-URBANISMO-JULHO 2016 EVENTO DO DIA 02/08/2016 - 160F - 39
'616 - MACKENZIE DIREITO T-U-TURMA JULHO 2016 EVENTO DO DIA 15/07/2016 - 161F - 30
'617 - MACKENZIE JORNALISMO-LETRAS-JULHO 2016-EVENTO DO DIA 04/08/2016 - 162F - 41
'618 - MACKENZIE PSICOLOGIA-JULHO 2016-EVENTO DO DIA 28/07/2016 - 163F - 37
'619 - MACKENZIE DIREITO C-D=JULHO 2016 EVENTO DO DIA 13/07/2016 - 164 - 28
'710	Mackenzie eng.prod/eletrica  - evento do dia 20/07/17 	196F	79	2017	JULHO
'711	Mackenzie eng.civil          - evento do dia 18/07/17	194F	77	2017	JULHO
'712	Mackenzie eng.mec/quimica    - evento do dia 19/07/17 	195F	78	2017	JULHO
'713	Mackenzie filo/ped/teo       - evento do dia 21/07/17	197F	80	2017	JULHO
'714	Mackenzie desing             - evento do dia 24/07/17 	198F	81	2017	JULHO
'715	Mackenzie arq e ubarnismo    - evento do dia 25/07/17	199F	82	2017	JULHO
'716	Mackenzie direito campinas   - evento do dia 12/07/17	202F	85	2017	JULHO
'---- 20/10/2017
'721	Mackenzie arq.urbanismo-vespertino - evento do dia 26/07/17	200F	83	2017	JULHO
'722	Mackenzie adm campinas             - evento do dia 27/07/17	201F	84	2017	JULHO
'723	Mackenzie c.comp/sist/mat          - evento do dia 01/08/17	203F	86	2017	JULHO
'724	Mackenzie                          - evento do dia 02/08/17 204F	87	2017	JULHO
'725	Mackenzie pp mkt                   - evento do dia 08/08/17 207F	90	2017	JULHO
'---- 07/11/2017
'727	Mackenzie pp criaçaõ/mkt/mat       - evento do dia 09/08/17 208F	91	2017	JULHO
'726	Mackenzie psicologia               - evento do dia 03/08/17	205F	88	2017	JULHO
'729	Mackenzie jornalismo/letras        - evento do dia 04/08/17 206F	89	2017	JULHO
'----- 21/11/2017
'730	Mackenzie comercio exterior        - evento do dia 10/08/17	209F	92	2017	JULHO		
'731	Mackenzie cie. cont. - cie. eco.   - evento do dia 11/08/17	210F	93	2017	JULHO		
'732	Mackenzie adm                      - evento do dia 15/08/17	211F	94	2017	JULHO		
'733	Mackenzie adm e comex              - evento do dia 16/08/17	212F	95	2017	JULHO		
'----- 24/11/2017
'734	Mackenzie direito                  - evento do dia 11/07/17	190F	73	2017	JULHO		
'735	Mackenzie direito                  - evento do dia 12/07/17	191F	74	2017	JULHO		
'736	Mackenzie direito                  - evento do dia 13/07/17	192F	75	2017	JULHO		
'737	Mackenzie direito                  - evento do dia 14/07/17	193F	76	2017	JULHO		

'836    Mackenzie direito                  - evento do dia 31/07/2018 cod 836 287F ok
'837    Mackenzie direito                  - evento do dia 01/08/2018 cod 837 290F ok
'838    Mackenzie direito                  - evento do dia 02/08/2018 cod 838 291F
'839    Mackenzie direito                  - evento do dia 03/08/2018 cod 839 292F

'Mackenzie filo-ped-teo turma julho 2018 evento do dia 09/08/2018 cod 844
'Mackenzie cbio-farm-fisio-nut-tec-gast turma julho 2018 evento do dia 14/08/2018 cod 843
'Mackenzie design turma julho 2018 evento do dia 17/08/2018

'-----------------------------------------------------------------------

SISTEMA_NOVO = 1
COD_PRODUCAO = 912
COD_CONTRATO = "354F"
COD_CURSO = 0
'NOME_CURSO = "ENGENHARIA"
ANO_CONCLUSAO = "2018"
SEMESTRE_CONCLUSAO = "D"
ALBUM = 0
INICIAL = 0
cod_colacao_importacao = 162

'-----------------------------------------------------------------------

dim rsMack, sql, sqlEndereco, rsCodNew

Set rsMack = objConexao.Execute("SELECT * FROM stillos3.TB_MACK_FORMANDOS where numero_identificacao is not null and numero_identificacao <> ''  and cod_mack_colacao = "&cod_colacao_importacao)
while not rsMack.Eof
	

	'GERAMOS O CODIGO DO FORMANDO	
	objConexao.Execute("INSERT INTO TB_FICHAS_IDENTIFICACAO_CONTAGEM_NUMERO_FORMANDO_CADASTRADO (COD_FORMANDO_CADASTRADO) VALUES (NULL)")
	Set rsCodNew = objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_FICHAS_IDENTIFICACAO_CONTAGEM_NUMERO_FORMANDO_CADASTRADO LIMIT 1")
	if not rsCodNew.Eof then
		COD_FORMANDO = rsCodNew("ULTIMO_ID")
	End if
	rsCodNew.Close
	Set rsCodNew = Nothing


	'SEPARAMOS AS INFORMACOES
	'FORMANDO
	NOME_FORMANDO = replace(rsMack("nome"),"'"," ")
	ALBUM = "0"
	INICIAL = "0"
	RETIRADO = ""
	BRINDE = ""
	COMISSAO = ""
	PRODUTO_AGREGADO = ""
	NOME_CURSO = ucase(rsMack("curso"))

	'ENDERECO
	ENDERECO = replace(rsMack("residencia_endereco"),"'"," ")
	CEP = rsMack("residencia_cep")
	FONE_RES = rsMack("telefone")
	FONE_CEL = rsMack("celular")
	BAIRRO = replace(rsMack("residencia_bairro"),"'"," ")
	CIDADE = replace(rsMack("residencia_cidade"),"'"," ")
	ESTADO = replace(rsMack("residencia_estado"),"'"," ")

	'INSERT DO FORMANDO
	sql = "INSERT INTO TB_FICHAS_IDENTIFICACAO ("&_
			"COD_FORMANDO ,"&_
			"SISTEMA_NOVO ,"&_
			"COD_PRODUCAO ,"&_
			"COD_CONTRATO ,"&_
			"COD_CURSO ,"&_
			"NOME_CURSO ,"&_
			"NOME_FORMANDO ,"&_
			"ANO_CONCLUSAO ,"&_
			"SEMESTRE_CONCLUSAO ,"&_
			"ALBUM ,"&_
			"INICIAL ,"&_
			"RETIRADO ,"&_
			"BRINDE ,"&_
			"COMISSAO ,"&_
			"PRODUTO_AGREGADO "&_	
			")"&_
			"VALUES ("&_
			COD_FORMANDO&"1, "&SISTEMA_NOVO&", "&COD_PRODUCAO&", '"&COD_CONTRATO&"', '"&COD_CURSO&"', '"&NOME_CURSO&"', "&sqlTexto(NOME_FORMANDO)&", "&ANO_CONCLUSAO&", '"&_
			SEMESTRE_CONCLUSAO&"', '"&ALBUM&"', '"&INICIAL&"', '"&RETIRADO&"', '"&BRINDE&"', '"&COMISSAO&"', '"&PRODUTO_AGREGADO&"')"
	response.write(sql&"<br>")
	objConexao.Execute(sql)


	'INSERT DO ENDERECO DO FORMANDO
	sqlEndereco =	"INSERT INTO TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO("&_
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
					COD_FORMANDO&"1, "&sqlTexto(ENDERECO)&", '"&CEP&"', '"&FONE_RES&"', '"&FONE_CEL&"', '"&BAIRRO&"', "&sqlTexto(CIDADE)&", "&sqlTexto(ESTADO)&", "&"1"&_
					")"
	response.write(sqlEndereco&"<br><hr>")
	objConexao.Execute(sqlEndereco)


rsMack.movenext()
Wend
rsMack.close
Set rsMack = nothing


FecharConexao()

%>