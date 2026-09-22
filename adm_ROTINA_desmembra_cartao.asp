<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
</head>
<body>
<%
Dim rs, sql, qtd_parcelas
qtd_parcelas = 0
sql = 	"SELECT "&_
		"TB_PEDIDOS.COD_PEDIDO, "&_
		"TB_PEDIDOS_CARTAO.COD_PEDIDO_CARTAO, "&_
		"TB_PEDIDOS_CARTAO.CARTAO, "&_
		"TB_PEDIDOS_CARTAO.VALOR_TOTAL, "&_
		"TB_PEDIDOS_CARTAO.PARCELAS, "&_
		"TB_PEDIDOS_CARTAO.NOME_TITULAR, "&_
		"TB_PEDIDOS_CARTAO.GRAU_PARENTESCO, "&_
		"TB_PEDIDOS_CARTAO.AUTENTICACAO, "&_
		"TB_PEDIDOS_CARTAO.DOCUMENTO, "&_
		"TB_PEDIDOS_CARTAO.DATA_AUTENTICACAO, "&_
		"TB_PEDIDOS_CARTAO.VERIFICADO, "&_
		"TB_PEDIDOS_CARTAO.DESMEMBRADO, "&_
		"TB_PEDIDOS_CARTAO.DESMEMBRADO_EM "&_
		"FROM TB_PEDIDOS INNER JOIN TB_PEDIDOS_CARTAO ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_CARTAO.COD_PEDIDO WHERE TB_PEDIDOS.FINALIZA_CORRECAO = 1 AND TB_PEDIDOS_CARTAO.DESMEMBRADO = 2"

AbrirConexao()

'Verifica se há alguma parcela = 0 e corrigi para 1
objConexao.Execute("UPDATE TB_PEDIDOS_CARTAO SET PARCELAS = '1' WHERE PARCELAS = 0")

Set rs = objConexao.Execute(sql)

While not rs.Eof
	Dim DATA_DOCUMENTO, TAXA, COD_OPERADORA, VALOR_PREVISTO, PARCELAS, sql2, TIPO_CARTAO
	DATA_DOCUMENTO = rs("DATA_AUTENTICACAO")
	qtd_parcelas = qtd_parcelas + formatNumber(rs("PARCELAS"))
	PARCELAS = formatNumber(0&rs("PARCELAS"),0)
	TAXA = 0
	
	'OPERADORAS:
	'1 VISA
	'2 MASTERCARD(REDECARD)
	'3 VISA ELECTRON
	
	'TAXAS VIGENTES/////////////////////////////////////
	
	'REDECARD (MASTERCARD)
	'À vista – 2,90%
	'Parcelado de 2 à 6 – 3,90%
	'Parcelado de 7 à 12 – 4,40%
		
	'VISA
	'VISA ELECTRON - Débito – 2%
	'VISA - Crédito – 1X – 2,70%
	'VISA – Crédito – 2 à 12 – 3,70%
	
	'///////////////////////////////////////////////////

	'SELECT * FROM `TB_TAXAS_CARTAO` WHERE OPERADORA = 2 AND TIPO_CARTAO=1 AND DATA_INICIO <= '2010-9-29' ORDER BY `DATA_INICIO` DESC LIMIT 1
	'1) É necessário incluir um campo para o tipo no cartão no lançamento Débio ou Crédito para o vendedor
	'2) Na correção da Ligia colocar tbm o campo Crédito ou Débito validando se for débito permite apenas 1 parcela
	'3) Depois verificar
	
	
	
	If rs("CARTAO") = "MASTERCARD" Then
			
		If PARCELAS = 1 Then
			TAXA = 2.9		
		Elseif PARCELAS >= 2 and PARCELAS <= 6  Then 
			TAXA = 3.9
		Elseif PARCELAS > 7 Then
			TAXA = 4.4
		End If
		
		COD_OPERADORA = 2
		
	Elseif rs("CARTAO") = "VISA" Then
	'Crédito
		COD_OPERADORA = 1
	
		If PARCELAS = 1 then
		'Taxa à vista
			TAXA = 2.7
				
		ElseIf PARCELAS > 2 Then
		
		'Taxa à prazo
			TAXA = 3.7
			
		End If
	
	Elseif rs("CARTAO") = "VISA ELECTRON" Then
	'Débito
			COD_OPERADORA = 3
			TAXA = 2
	End If
	
	VALOR_PREVISTO 			= 0	
	VALOR_PREVISTO 			= formatNumber(0&rs("VALOR_TOTAL"))
	'Response.Write(PARCELAS&"<br /><br /><br />")
	VALOR_NORMAL_PARCELA 	= formatNumber(VALOR_PREVISTO/PARCELAS,2)
	VALOR_PREVISTO 			= formatNumber((VALOR_PREVISTO - (VALOR_PREVISTO/100 * TAXA))/PARCELAS,2)
	
	
	For i = 0 to PARCELAS-1

	DATA_DOCUMENTO = DateAdd("d",31,DATA_DOCUMENTO)
		
	sql2= 	"INSERT INTO TB_PEDIDOS_CARTAO_DESMEMBRADO ( "&_
			"COD_PEDIDO_CARTAO , "&_
			"COD_PEDIDO , "&_
			"cod_cartao , "&_
			"documento , "&_
			"autenticacao , "&_
			"data_documento , "&_
			"parcela , "&_
			"valor_parcela , "&_
			"valor_previsto , "&_
			"dt_vencimento , "&_
			"taxa ) "&_
			"VALUES ( "&_
"'"&rs("COD_PEDIDO_CARTAO")&"', '"&rs("COD_PEDIDO")&"', '"&COD_OPERADORA&"','"&rs("DOCUMENTO")&"', '"&rs("AUTENTICACAO")&"', "&sqlData(rs("DATA_AUTENTICACAO"))&", "&_
		"'"&i+1&"','"&sqlNumero(VALOR_NORMAL_PARCELA)&"', '"&sqlNumero(VALOR_PREVISTO)&"', "&sqlData(DATA_DOCUMENTO)&", '"&sqlNumero(TAXA)&_
		"')"
		'Insere na base
		objConexao.Execute(sql2)
		'Response.Write(sql2&"<br /><br />")		
	Next
		'Da o update informando que foi desmembrado
		objConexao.Execute("UPDATE TB_PEDIDOS_CARTAO SET DESMEMBRADO = '1', DESMEMBRADO_EM = NOW() WHERE COD_PEDIDO_CARTAO = "&rs("COD_PEDIDO_CARTAO")&" LIMIT 1")
		
rs.MoveNext()
Wend

If rs.Eof then

Response.Write("Não há registros a serem desmembrados!")

End if 

Response.Write("Qtd parcelas geradas: "&qtd_parcelas)
rs.Close
Set rs = Nothing
FecharConexao()

Dim MES_INICIO, MES_INICIO2
Dim ANO_INICIO, ANO_INICIO2

MES_INICIO		= Request.QueryString("MES_INICIO")
MES_INICIO2		= Request.QueryString("MES_INICIO2")
ANO_INICIO		= Request.QueryString("ANO_INICIO")
ANO_INICIO2		= Request.QueryString("ANO_INICIO2")

Response.Redirect("adm_relatorios_previsao_recebimento.asp?MES_INICIO="&MES_INICIO&"&ANO_INICIO="&ANO_INICIO&"&MES_INICIO2="&MES_INICIO2&"&ANO_INICIO2="&ANO_INICIO2&"&qtdPARCELAS="&qtd_parcelas)
%>

</body>
</html>



