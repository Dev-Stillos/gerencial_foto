<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Validação especial para verificar se possui sub-validação específica
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '44' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se não encontrar então redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
		
	rs.Close
	Set rs=Nothing
	
FecharConexao()

'/////////////////////////////////////////////////////////////////////////////////////////////

Dim DATA_INICIO, DATA_FIM, COD_PRODUTO_DETALHADO
Dim nome_fornecedor, valor_produto, imprimir_nome_formando
Dim TOTAL_PRODUTOS, VALOR_TOTAL_PAGO
Dim diferenca

DATA_INICIO 			= ""&Request.Form("DATA_INICIO")
DATA_FIM 				= ""&Request.Form("DATA_FIM")
COD_PRODUTO_DETALHADO	= ""&Request.Form("detalhar_produto")
nome_fornecedor			= Ucase(""&Request.Form("nome_fornecedor"))
valor_produto			= formatNumber(0&Request.Form("valor_produto"))
imprimir_nome_formando	= Cstr(""&Request.Form("imprimir_nome_formando"))




If DATA_INICIO = "" OR DATA_FIM = "" Then
	Response.Redirect("adm_relatorios.asp?msg=Atenção preencha a data de início e fim para gerar o relatório de venda de produtos")
Else

	If isDate(DATA_INICIO) and isDate(DATA_FIM) Then
		DATA_INICIO		= formatDateTime(DATA_INICIO)
		DATA_FIM		= formatDateTime(DATA_FIM)
		diferenca		= dateDiff("d",DATA_INICIO,DATA_FIM)
	Else
		Response.Redirect("adm_relatorios.asp?msg=Atenção formato de data inválido preencha no formato DD/MM/AAAA")
	
	End If	
		
End If

If diferenca < 0 Then
Response.Redirect("adm_relatorios.asp?msg=Atenção data de início é maior qua a data fim do intervalo de datas.")
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="all"></link>
<title>StillosWEB - Sistema corporativo - Versão 1.2</title>
<style type="text/css">
<!--
body {
	background-color: #FFFFFF;
}
-->
</style></head>

<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
  <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top">
	
	  <h2 align="center">Recibo de Fornecedor</h2>
	  <%
		AbrirConexao()		
		Dim rs, sql
		
		Set rs=objConexao.Execute("SELECT TB_PEDIDOS_PRODUTOS.COD_PRODUTO, SUM(TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS QTD_POR_PRODUTO, TB_PRODUTOS.NOME_PRODUTO FROM TB_PEDIDOS INNER JOIN TB_PEDIDOS_PRODUTOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO INNER JOIN TB_PRODUTOS on TB_PEDIDOS_PRODUTOS.COD_PRODUTO = TB_PRODUTOS.COD_PRODUTO WHERE TB_PEDIDOS.FINALIZA_CORRECAO = 1 AND  TB_PEDIDOS.DATA_VENDA BETWEEN "&sqlData(DATA_INICIO)&" AND "&sqlData(DATA_FIM)&" AND TB_PEDIDOS_PRODUTOS.COD_PRODUTO = '"&COD_PRODUTO_DETALHADO&"' GROUP BY TB_PEDIDOS_PRODUTOS.COD_PRODUTO ORDER BY TB_PRODUTOS.NOME_PRODUTO")
		
		If not rs.Eof Then
		
			TOTAL_PRODUTOS = formatNumber(0&rs("QTD_POR_PRODUTO"))
			If valor_produto <> 0 Then
				VALOR_TOTAL_PAGO = (TOTAL_PRODUTOS*valor_produto)/1
			End If
		%>
	  <table width="100%" 
		      border="0" align="center"
					cellpadding="5" 
					style="border-collapse:collapse; margin-top:10px;">              
          <tr>
 
            
            <td width="39%" class="textoMedio"><strong><%=rs("NOME_PRODUTO")%></strong></td>
            <td width="40%" class="textoMedio"><div align="center"><strong><%=ArrumaData(DATA_INICIO)%> &agrave; <%=ArrumaData(DATA_FIM)%></strong></div></td>
            <td width="21%" class="textoMedio"><div align="right"><strong><%=rs("QTD_POR_PRODUTO")%></strong></div></td>
          </tr>
          <tr>
            <td colspan="3" valign="top" class="textoMedio">
                	<%
					SQL = "SELECT TB_PEDIDOS_PRODUTOS.COD_PRODUTO, SUM(TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS QTD_POR_CONTRATO, TB_FICHAS_IDENTIFICACAO.COD_CONTRATO, TB_ENTIDADES.DESCRICAO FROM TB_PEDIDOS INNER JOIN TB_PEDIDOS_PRODUTOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO INNER JOIN TB_FICHAS_IDENTIFICACAO ON TB_PEDIDOS_PRODUTOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO INNER JOIN TB_ENTIDADES ON TB_FICHAS_IDENTIFICACAO.COD_CONTRATO = TB_ENTIDADES.COD_CONTRATO  WHERE TB_PEDIDOS.FINALIZA_CORRECAO = 1 AND TB_PEDIDOS.DATA_VENDA BETWEEN "&sqlData(DATA_INICIO)&" AND "&sqlData(DATA_FIM)&" AND  TB_PEDIDOS_PRODUTOS.COD_PRODUTO = '"&rs("COD_PRODUTO")&"' GROUP BY TB_FICHAS_IDENTIFICACAO.COD_CONTRATO,TB_PEDIDOS_PRODUTOS.COD_PRODUTO"
					Set rs2=objConexao.Execute(SQL)
						While not rs2.Eof
					%>
                	<table width="100%" border="0" cellspacing="0" cellpadding="2">
                	  <tr>
                	    <td width="79%" style="border-bottom:solid 1px #666666"><strong>
                	      <%=rs2("DESCRICAO")%>
                	      
                	      </strong></td>
      <td width="21%" style="border-bottom:solid 1px #666666"><div align="right"><strong><%=rs2("QTD_POR_CONTRATO")%></strong></div></td>
    </tr>
                    </table>
                	<%
							Set rs3=objConexao.Execute("SELECT TB_PEDIDOS_PRODUTOS.*, SUM(TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS QTD_POR_FORMANDO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, TB_PEDIDOS.COD_VENDEDOR, TB_PEDIDOS.DATA_VENDA FROM TB_PEDIDOS INNER JOIN TB_PEDIDOS_PRODUTOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO INNER JOIN TB_FICHAS_IDENTIFICACAO ON TB_PEDIDOS_PRODUTOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO WHERE TB_PEDIDOS.DATA_VENDA BETWEEN "&sqlData(DATA_INICIO)&" AND "&sqlData(DATA_FIM)&" AND COD_PRODUTO = '"&rs("COD_PRODUTO")&"' AND TB_FICHAS_IDENTIFICACAO.COD_CONTRATO = '"&rs2("COD_CONTRATO")&"' AND TB_PEDIDOS.FINALIZA_CORRECAO = 1 GROUP BY COD_FORMANDO")
							%>
                   			 <table width="100%" border="0" cellspacing="0" cellpadding="0">
                             	<%While not rs3.Eof%>
  <tr>
    <td style="padding-left:20px;"><div style="border-bottom:solid 1px #CCCCCC; padding:2px;"><%=digito(rs3("COD_FORMANDO"))%>
    <%If imprimir_nome_formando = "S" Then%>
    	&nbsp;&nbsp;<%=rs3("NOME_FORMANDO")%>
    <%End If%>
    </div></td>
    <td width="16%"><div style="border-bottom:solid 1px #CCCCCC; padding:2px;" align="right"><%=rs3("QTD_POR_FORMANDO")%></div></td>
    </tr>
    							<%								
								rs3.MoveNext()
								Wend
								%>
                             </table>
                   			 <%
							rs3.Close
							Set rs3=Nothing%>
                      	     <br />
                     <%
					 rs2.MoveNext()
					 Wend
					 
					 rs2.Close
					 Set rs2=Nothing
					 %>                </td>
            </tr>
        </table>
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="0">
  <tr>
    <td width="31%" class="textoMedio" style="border-bottom:solid 1px #CCCCCC">Total Produtos: </td>
    <td width="69%" class="textoMedio" style="border-bottom:solid 1px #CCCCCC"><div align="right"><%=formatNumber(TOTAL_PRODUTOS,0)%></div></td>
    </tr>
  <tr>
    <td class="textoMedio" style="border-bottom:solid 1px #CCCCCC">Valor pago por Produto:</td>
    <td class="textoMedio" style="border-bottom:solid 1px #CCCCCC"><div align="right"><%=formatNumber(valor_produto,2)%></div></td>
    </tr>
  <tr>
    <td class="textoMedio"  style="border-bottom:solid 1px #CCCCCC">Valor Total:</td>
    <td class="textoMedio"  style="border-bottom:solid 1px #CCCCCC"><div align="right"><%=formatNumber(VALOR_TOTAL_PAGO,2)%></div></td>
    </tr>
  <tr>
    <td class="textoMedio"  style="border-bottom:solid 1px #CCCCCC">Fornecedor:</td>
    <td class="textoMedio"  style="border-bottom:solid 1px #CCCCCC"><div align="right"><%=Ucase(nome_fornecedor)%></div></td>
    </tr>
  <tr>
    <td class="textoMedio"  style="border-bottom:solid 1px #CCCCCC">Data:</td>
    <td class="textoMedio"  style="border-bottom:solid 1px #CCCCCC"><div align="right"><%=ArrumaData(NOW)%></div></td>
    </tr>
  <tr>
    <td colspan="2" class="textoMedio"><div style="width:350px; margin:auto; border-bottom:dotted 1px #666666; padding-top:20px;">&nbsp;</div></td>
    </tr>
  <tr>
    <td colspan="2" class="textoMedio"><div align="center">Assinatura do fornecedor</div></td>
  </tr>
</table>

        
<%Else%>
        		
                <div style="padding:50px; font-size:16px;" align="center">Não foi encontrado resultados para os parametros informados de <strong><%=ArrumaData(DATA_INICIO)%></strong> à <strong><%=ArrumaData(DATA_FIM)%></strong></div>
                
        <%End If
		
		rs.Close
		Set rs=Nothing	
			
		FecharConexao()
		%>
		   
		  </td>
    </tr>
</table>

 
  
  </td>
    </tr>
</table>
</body>
</html>






