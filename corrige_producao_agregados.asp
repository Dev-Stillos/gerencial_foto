<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<%
'###########################################################
'RODAR UMA ÚNICA VEZ PARA NÃO AUMENTAR MUITO O INDICE DAS TABELAS
'EM CASO DE FALHA PODERÁ REPETIR SEM PROBLEMAS
'Verifique se há produtos lançados no cabeçalho da produção, se estiver alguma diferença lance os produtos que deverão ser replicados para os formandos

'INSTRUÇÕES:
'1) Somente para produções finalizadas e com movimentações dos vendedores lançadas
'2) Somente para produções com produtos agregados no cabeçalho da produção
'3) Lance os produtos corretos no cabeçalho da produção
'###########################################################
COD_PRODUCAO = 142

Dim CONTINUAR
CONTINUAR = "SIM"

AbrirConexao()
	
	Set rs=objConexao.Execute("SELECT COD_PRODUCAO FROM TB_PRODUCAO WHERE COD_PRODUCAO = "&COD_PRODUCAO&" AND PRODUCAO_FINALIZADA = 1 AND PRODUTOS_AGREDADOS = 1")
			
			If not rs.Eof Then
				Set rs2=objConexao.Execute("SELECT * FROM TB_PRODUCAO_PRODUTOS_AGREGADOS WHERE COD_PRODUCAO ="&COD_PRODUCAO)
						
						If rs2.Eof Then
						'Não há produtos lançados
							CONTINUAR = "NÃO"
							
						End If
						
				rs2.Close
				Set rs2=Nothing
			Else
				'Não está finalizada ou não tem produtos agregados
				CONTINUAR = "NÃO"
			End If
			
	rs.Close
	Set rs=Nothing
	
FecharConexao()

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
  <%
  If CONTINUAR = "SIM" Then
  
 	 AbrirConexao()
 	
	'1) LIMPEZA DE PRODUTOS AGREGADOS, CASO HOUVER ALGUM LANÇADO PARA PRODUÇÃO
	'	DEIXA APENAS OS PRODUTOS AGREGADOS DO CABEÇALHO DA PRODUÇÃO
		
		'1.1) Limpa fichas identificação para todos com produtos agregados
		 objConexao.Execute("UPDATE TB_FICHAS_IDENTIFICACAO SET PRODUTO_AGREGADO = '1' WHERE COD_PRODUCAO ="&COD_PRODUCAO)
		
		'1.2) Limpa fichas identificação para todos com produtos agregados
		 objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS WHERE COD_PRODUCAO = "&COD_PRODUCAO)
		
		'1.3) Limpa fichas identificação movimentações
		Set rs=objConexao.Execute("SELECT COD_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = '"&COD_PRODUCAO&"'")
				
				While not rs.Eof
						
						'Deleta as movimentações do formando
						objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_FORMANDO = "&rs("COD_FORMANDO")&" AND TIPO > 1")
				
				rs.MoveNext()
				Wend
		
		rs.Close
		Set rs=Nothing
  
  '--Fim limpeza e correção de formandos		
  '//////////////////////////////////////////////////
  
  '2) LANÇAMENTO DOS ITENS AGREGADOS DO CABEÇALHO DA COMISSÃO
  		
		'2.1) Seleciona todos os formandos da produção para iniciar os lançamentos
		Set rs=objConexao.Execute("SELECT COD_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = '"&COD_PRODUCAO&"'")
		
		While not rs.Eof
		
			'2.1.1) Verifica quais os produtos agregados da produção
			Set rsProdAgre = objConexao.Execute("SELECT COD_PRODUCAO, COD_PRODUTO,  Sum(QUANTIDADE) as  QUANTIDADE  FROM TB_PRODUCAO_PRODUTOS_AGREGADOS WHERE COD_PRODUCAO="&COD_PRODUCAO&" GROUP BY COD_PRODUTO")
										
							Dim agregadoNumero
							agregadoNumero = 1
							While not rsProdAgre.Eof
							agregadoNumero = agregadoNumero + 1
							
							'2.1.1.1) Lança nas fichas individuais os produtos agregados
							 objConexao.Execute("INSERT INTO TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ( "&_
												"TIPO, "&_
												"COD_PRODUCAO , "&_
												"COD_FORMANDO , "&_
												"COD_PRODUTO , "&_
												"QUANTIDADE, CADASTRADO_POR) "&_
												"VALUES ("&_
												agregadoNumero&", "&COD_PRODUCAO&", "&rs("COD_FORMANDO")&", "&rsProdAgre("COD_PRODUTO")&", "&rsProdAgre("QUANTIDADE")&",'9999')")
							
							
							'2.1.1.2) Verifica qual o vendedor da ficha para lançar igual
							Dim COD_VENDEDOR, DATA_SAIDA
							COD_VENDEDOR	= ""
							DATA_SAIDA		= ""
							'Pega o vendedor da ficha inicial por precaução
							Set rsVendedor = objConexao.Execute("SELECT COD_VENDEDOR, DATA_SAIDA FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_FORMANDO="&rs("COD_FORMANDO")&" AND TIPO=1")
										
										If not rsVendedor.Eof Then
												COD_VENDEDOR 	= rsVendedor("COD_VENDEDOR")
												DATA_SAIDA		= rsVendedor("DATA_SAIDA")
										End If
										
							rsVendedor.Close
							Set rsVendedor = Nothing
																		
							'2.1.1.3) Lança nas movimentações dos formandos os agregados
							 objConexao.Execute("INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ("&_
												"DATA_SAIDA ,"&_
												"COD_FORMANDO ,"&_
												"TIPO ,"&_
												"COD_VENDEDOR ,"&_
												"CADASTRADO_POR"&_
												")"&_
												"VALUES ("&_
												sqlData(DATA_SAIDA)&","&rs("COD_FORMANDO")&","&agregadoNumero&", "&COD_VENDEDOR&", '9999')")						
						
							rsProdAgre.MoveNext()
							Wend
							
						
			rsProdAgre.Close()
			Set rsProdAgre = Nothing
					
		
		rs.MoveNext()
  		Wend
  
  rs.Close  
  Set rs=Nothing
  
  FecharConexao()
  
  Response.Write("Pronto!<br />Correção de produtos agregados finalizado com sucesso!")
  
  Else
  	Response.Write("ATENÇÃO NÃO FOI POSSÍVEL EFETUAR A CORREÇÃO DE PRODUTOS AGREGADOS VERIFIQUE AS INSTRUÇÕES!")
  End If
  %>
</body>
</html>