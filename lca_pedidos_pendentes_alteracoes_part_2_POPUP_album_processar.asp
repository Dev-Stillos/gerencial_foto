<!-- #include virtual="gerencial_foto/includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim COD_PEDIDO
COD_PEDIDO = Request.QueryString("cod_pedido")

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; chars44=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/principal.css" media="screen"></link>
<title>Loja Virtual - Stillo's Produções e Eventos</title>
<noscript>
  <meta http-equiv="Refresh" content="1;url=http://stillos.com.br/gerencial_foto/noscript.asp">
</noscript>
<style type="text/css">
<!--
.style3 {
	font-size: 13px;
	color: #003366;
	font-weight: bold;
}
-->
</style>
</head>
<body>
<table id="tb_p" align="center" cellpadding="0" cellspacing="0">
<tr>
<td>&nbsp;</td>
</tr>

<tr>
<td align="center">
<div>
  <p align="center" style="padding-top:5px;">&nbsp;</p>
  <p align="left">
   <%
  '1) Seleciona os códigos dos formandos do pedido
  '2) Seleciona os dados inseridos nas fichas
  AbrirConexao()
  
  Dim sql, sql2, rs, rs5, rs6
  
  sql = "SELECT TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO, TB_FICHAS_IDENTIFICACAO.ALBUM, TB_FICHAS_IDENTIFICACAO.INICIAL, TB_FICHAS_IDENTIFICACAO.BRINDE, "&_
		"TB_FICHAS_IDENTIFICACAO.PRODUTO_AGREGADO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, TB_FICHAS_IDENTIFICACAO.COMISSAO, "&_
		"TB_PEDIDOS_FORMANDOS.COD_FORMANDO "&_
		"FROM TB_FICHAS_IDENTIFICACAO INNER JOIN TB_PEDIDOS_FORMANDOS ON "&_
		"TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = TB_PEDIDOS_FORMANDOS.COD_FORMANDO "&_
		"WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = '"&COD_PEDIDO&"' AND TB_PEDIDOS_FORMANDOS.COD_FORMANDO="&Request.QueryString("cod_for")&" LIMIT 1"
  		
  Set rs = objConexao.Execute(sql)
	Dim QTD_BRINDE
	
		While not rs.Eof
  		Dim FATURADO, RETIRADO, COD_FOTO_PADRAO, VALOR_UNITARIO, QUANTIDADE_BRINDE, NOME_FORMANDO
		'Grava somente a quantidade que foi faturada
		NOME_FORMANDO = rs("NOME_FORMANDO")
		FATURADO = formatNumber(0&Request.Form("QTD_FATURADA_"&rs("COD_FORMANDO")),0)
		RETIRADO = formatNumber(0&Request.Form("QTD_RETIRADA_"&rs("COD_FORMANDO")),0)
		COD_FOTO_PADRAO = Request.Form("COD_FOTO_PADRAO_"&rs("COD_FORMANDO"))
		'JAMAIS COLOCAR UM FORMAT NUMBER EM CAMPOS Q SERÃO GRAVADOS NA BASE DE DADOS!!!
		VALOR_UNITARIO = Request.Form("VALOR_UNITARIO_"&rs("COD_FORMANDO"))
		QUANTIDADE_BRINDE = formatNumber(0&Request.Form("QUANTIDADE_BRINDE_"&rs("COD_FORMANDO")),0)
		'//////////////////////////////////////////////////////////////////////		
		'ATENÇÃO!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
		'Produtos que não podem ser excluídos:
		'1) Álbuns		> TB_FICHAS_IDENTIFICACAO
		'2) Brindes		> TB_FICHAS_IDENTIFICACAO_BRINDES
		'3) Prod. Agre	> TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS 
		
		'NESTAS TABELAS É REALIZADO SOMENTE O UPDATE E NADA EXCLUSÃO
		'//////////////////////////////////////////////////////////////////////
					
  		'1---------------------------------------------------------
		'Grava o álbum
		sql2 = 	"UPDATE TB_FICHAS_IDENTIFICACAO SET "&_
				"FATURADO = "&FATURADO&", "&_
				"VALOR_UN = '"&sqlNumero(VALOR_UNITARIO)&"', "&_
				"RETIRADO = "&RETIRADO&" WHERE COD_FORMANDO = "&rs("COD_FORMANDO")&" LIMIT 1"
				objConexao.Execute(sql2)
				
		'2 - Será feito somente na conclusão, dando o update de vendido
		'3 - Na próxima tela a informação se foi vendido e qual o valor se foi vendido
	%>
			
	
<br />&nbsp;
  <%
  		rs.MoveNext()
		Wend
		
		rs.Close()
		Set rs = Nothing
  FecharConexao()
  'Redirecionar para próxima etapa inclusão de produtos agregados a venda, juntamente com outros brindes já lançados  
  
 
  Response.Redirect("lca_pedidos_pendentes_alteracoes_part_2.asp?cod_pedido="&COD_PEDIDO&"&msg=Álbum do(a) formando(a) "&NOME_FORMANDO&" alterado com sucesso!")
  %>
<p align="center" style="margin:0px">&nbsp;</p>

  </p>
</div></td>
</tr>
</table>
</body>
</html>