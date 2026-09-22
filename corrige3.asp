<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<%
problema
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<style type="text/css">
<!--
body,td,th {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 11px;
}
-->
</style></head>
<body>
<table width="100%" border="1" cellspacing="0" cellpadding="5">
  <tr>
    <td width="11%" bgcolor="#CCCCCC"><div align="center"><strong>Cód Formando</strong></div></td>
    <td width="35%" bgcolor="#CCCCCC"><div align="center"><strong>Produto</strong></div></td>
    <td width="36%" bgcolor="#CCCCCC"><div align="center"><strong>Data Saída</strong></div></td>
    <td bgcolor="#CCCCCC"><div align="center"><strong>Data Retorno</strong></div></td>
  </tr>
  <%
  AbrirConexao()
  
  Dim COD_PRODUCAO, COD_VENDEDOR, LOTE, DATA_SAIDA, DATA_PRAZO, QTD_FORMANDOS_INICIAL, STATUS_W
  COD_PRODUCAO = 184
  LOTE = 0
  
  '1) Verifica qual o vendedor correto responsável
  '2) Verifica se está vendido
  '3) Verifica se há pedido
  '4) Sabendo número de vendedores
  
  
  
  '/////////////////
  'LIMPA BASE TEMPORÁRIA E TABELA DE LOTES CABEÇALHO
  objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA WHERE COD_USUARIO ="&Session("COD_USUARIO"))
  objConexao.Execute("DELETE FROM TB_PRODUCAO_LOTES WHERE COD_PRODUCAO = "&COD_PRODUCAO)
  
  '1)--- Verifica quem é o VENDEDOR responsável pelo formando para isso seleciona todos os formandos
	Set rs=objConexao.Execute("SELECT COD_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = "&COD_PRODUCAO)
  		
		While not rs.Eof
				
				'Verifica o primeiro vendedor do formando
				Set rs2=objConexao.Execute("Select COD_VENDEDOR, DATA_SAIDA FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES Where COD_FORMANDO = "&rs("COD_FORMANDO")&" and tipo = 0 order by DATA_RETORNO desc limit 1")
					If not rs2.Eof Then
			
						COD_VENDEDOR = rs2("COD_VENDEDOR")
						DATA_SAIDA = rs2("DATA_SAIDA")
					Else
			
						COD_VENDEDOR = 0
						DATA_SAIDA = Null					
			
					End If
				rs2.Close
				Set rs2=Nothing
						
				'Insere na base temporária
				objConexao.Execute("INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA (COD_FORMANDO ,COD_ULTIMO_VENDEDOR_FICHA,COD_USUARIO,DATA_SAIDA) VALUES ("&rs("COD_FORMANDO")&","&COD_VENDEDOR&","&Session("COD_USUARIO")&","&sqlData(DATA_SAIDA)&")")
		
		rs.MoveNext()
		Wend
  
	rs.Close
	Set rs=Nothing 
	
	'/////////////////////////////////////////////////////////////////////////////////////
	
	'Seleciona os vendedores da produção
	Set rs=objConexao.Execute("SELECT COD_ULTIMO_VENDEDOR_FICHA, DATA_SAIDA FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA WHERE COD_USUARIO = 1 GROUP BY COD_ULTIMO_VENDEDOR_FICHA ORDER BY COD_ULTIMO_VENDEDOR_FICHA")
  		
		While not rs.Eof
		
				COD_VENDEDOR = rs("COD_ULTIMO_VENDEDOR_FICHA")
				DATA_SAIDA = rs("DATA_SAIDA")
				DATA_PRAZO = DateAdd("d",15,DATA_SAIDA)
				LOTE = LOTE + 1
							
				'Seleciona todos os formandos deste vendedor
				Set rs2=objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA WHERE COD_USUARIO = "&Session("COD_USUARIO")&" and  COD_ULTIMO_VENDEDOR_FICHA="&COD_VENDEDOR,QTD_FORMANDOS_INICIAL)
					If not rs2.Eof Then
					
						Response.Write("<br /><br />///////////////////////////////////////////////////////////<br />")
						Response.Write("VENDEDOR: "&COD_VENDEDOR&" LOTE "&LOTE&" SAÍDA: "&DATA_SAIDA)
						'////////////////////
						'Grava o lote cabeçalho
						
						objConexao.Execute("INSERT INTO TB_PRODUCAO_LOTES (COD_PRODUCAO , LOTE ,COD_VENDEDOR_INICIAL ,DATA_SAIDA_INICIAL ,QTD_FORMANDOS_INICIAL ,PRAZO_DIAS_CORRIDOS_INICIAL)VALUES ('"&COD_PRODUCAO&"','"&LOTE&"','"&COD_VENDEDOR&"',"&sqlData(DATA_SAIDA)&",'"&QTD_FORMANDOS_INICIAL&"',15)")
						
						'////////////////////
						While not rs2.Eof
								
								
								Set rs3=objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO = "&rs2("COD_FORMANDO"))
								
								If not rs3.Eof Then
															
									Response.Write("<BR/>      "&rs2("COD_FORMANDO"))
								
									'1//// Atribui o valor do lote
										objConexao.Execute("UPDATE TB_FICHAS_IDENTIFICACAO SET LOTE = '"&LOTE&"' WHERE COD_FORMANDO = '"&rs2("COD_FORMANDO")&"' LIMIT 1")
									'2//// Atribui o valor das movimentações
										'2.1//// Delete o formando do lote individual
											objConexao.Execute("DELETE FROM TB_PRODUCAO_LOTES_FORMANDOS WHERE COD_FORMANDO = '"&rs2("COD_FORMANDO")&"'")	
						

										If int(rs3("VENDIDO")) = 1 Then
											STATUS_W = 1
										Else
										
										'Verifica se tem pedido
											Set rs4=objConexao.Execute("SELECT * FROM TB_PRODUCAO_LOTES_FORMANDOS WHERE COD_FORMANDO = "&rs2("COD_FORMANDO"))
												If not rs4.Eof Then
												'Existe pedido nao conferido
														STATUS_W = 2
												Else
														STATUS_W = 3
												End if
											
											rs4.Close
											Set rs4=Nothing
										
										End if
									
								objConexao.Execute("INSERT INTO TB_PRODUCAO_LOTES_FORMANDOS (COD_FORMANDO, COD_VENDEDOR , DATA_SAIDA , DATA_PRAZO ,STATUS)VALUES ('"&rs2("COD_FORMANDO")&"',  '"&COD_VENDEDOR&"',  "&sqlData(DATA_SAIDA)&",  "&sqlData(DATA_PRAZO)&",  '"&STATUS_W&"')")
								
								
								End If
								
								rs3.Close
								Set rs3=Nothing
								
							
						rs2.MoveNext()
						Wend
						
					End If
				rs2.Close
				Set rs2=Nothing
				
		
		rs.MoveNext()
		Wend
  
	rs.Close
	Set rs=Nothing 
%>
</table>

</body>
</html>