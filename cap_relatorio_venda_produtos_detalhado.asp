<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 2
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
  Dim titulo_pagina
  titulo_pagina = "Boletim Informativo"
%>
<%
Dim calendario_necessario
calendario_necessario = "NÃO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
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
Dim diferenca
DATA_INICIO = ""&Request.QueryString("DATA_INI")
DATA_FIM = ""&Request.QueryString("DATA_FIM")
COD_PRODUTO_DETALHADO = Request.QueryString("detalhar_produto")

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
	<form method="post" target="_blank" name="impressao" action="cap_relatorio_venda_produtos_detalhado_imprimir.asp">
    <input type="hidden" name="DATA_INICIO" value="<%=DATA_INICIO%>" />
    <input type="hidden" name="DATA_FIM" value="<%=DATA_FIM%>" />
    <input type="hidden" name="detalhar_produto" value="<%=COD_PRODUTO_DETALHADO%>" />
    
  	<fieldset style="padding:10px"> 
  		<legend>Opções de Impressão do Recibo</legend>
        <table width="100%" border="0" cellspacing="0" cellpadding="4">
  <tr>
    <td width="22%"><strong>Nome do Fornecedor:</strong></td>
    <td width="78%"><input type="text" class="inputTamanho3" name="nome_fornecedor" /></td>
  </tr>
  <tr>
    <td><strong>Valor Pago por Produto:</strong></td>
    <td><input type="text" class="inputTamanho1 padrao_valor" name="valor_produto" /></td>
  </tr>
  <tr>
    <td><strong>Imprimir nome de Formando:</strong></td>
    <td><table width="200" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="20">
          <input class="inputRadio" checked="checked" type="radio" name="imprimir_nome_formando" value="S" />        </td>
        <td width="41">Sim</td>
        <td width="21"><input class="inputRadio" type="radio" name="imprimir_nome_formando" value="N" /></td>
        <td width="118">Não</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="2"><div align="center"><a href="javascript:document.impressao.submit()" class="linkMaster" style="color:#CC0000">Clique aqui para imprimir o recibo</a></div></td>
    </tr>
</table>

        
  	</fieldset>
    </form>
    <br />
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">
	
	  <h2>Relatórios - Venda de Produtos por Per&iacute;odo<br />	    
          <br />
      </h2>
	  <div align="center" class="msg"><%=Request.QueryString("msg")%></div>
	  <%
		AbrirConexao()		
		Dim rs, sql
		
		Set rs=objConexao.Execute("SELECT TB_PEDIDOS_PRODUTOS.COD_PRODUTO, SUM(TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS QTD_POR_PRODUTO, TB_PRODUTOS.NOME_PRODUTO FROM TB_PEDIDOS INNER JOIN TB_PEDIDOS_PRODUTOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO INNER JOIN TB_PRODUTOS on TB_PEDIDOS_PRODUTOS.COD_PRODUTO = TB_PRODUTOS.COD_PRODUTO WHERE TB_PEDIDOS.FINALIZA_CORRECAO = 1 AND  TB_PEDIDOS.DATA_VENDA BETWEEN "&sqlData(DATA_INICIO)&" AND "&sqlData(DATA_FIM)&" AND TB_PEDIDOS_PRODUTOS.COD_PRODUTO = '"&COD_PRODUTO_DETALHADO&"' GROUP BY TB_PEDIDOS_PRODUTOS.COD_PRODUTO ORDER BY TB_PRODUTOS.NOME_PRODUTO")
		
		If not rs.Eof Then%>
        <span class="textoMaior"><strong><%=ArrumaData(DATA_INICIO)%> à <%=ArrumaData(DATA_FIM)%></strong></span>
        <table width="100%" 
		      border="0" align="center"
					cellpadding="5" 
					style="border-collapse:collapse; margin-top:10px;">              
          <tr>
 
            
            <td width="79%" class="textoMaior"><strong><%=rs("NOME_PRODUTO")%></strong></td>
            <td width="21%" class="textoMaior"><div align="right"><%=rs("QTD_POR_PRODUTO")%></div></td>
          </tr>
          <tr>
            <td colspan="2" valign="top" class="textoMedio">
                	<%
					SQL = "SELECT TB_PEDIDOS_PRODUTOS.COD_PRODUTO, SUM(TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS QTD_POR_CONTRATO, TB_FICHAS_IDENTIFICACAO.COD_CONTRATO, TB_ENTIDADES.DESCRICAO FROM TB_PEDIDOS INNER JOIN TB_PEDIDOS_PRODUTOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO INNER JOIN TB_FICHAS_IDENTIFICACAO ON TB_PEDIDOS_PRODUTOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO INNER JOIN TB_ENTIDADES ON TB_FICHAS_IDENTIFICACAO.COD_CONTRATO = TB_ENTIDADES.COD_CONTRATO  WHERE TB_PEDIDOS.FINALIZA_CORRECAO = 1 AND TB_PEDIDOS.DATA_VENDA BETWEEN "&sqlData(DATA_INICIO)&" AND "&sqlData(DATA_FIM)&" AND  TB_PEDIDOS_PRODUTOS.COD_PRODUTO = '"&rs("COD_PRODUTO")&"' GROUP BY TB_FICHAS_IDENTIFICACAO.COD_CONTRATO,TB_PEDIDOS_PRODUTOS.COD_PRODUTO"
					Set rs2=objConexao.Execute(SQL)
						While not rs2.Eof
					%>
                	<table width="100%" border="0" cellspacing="0" cellpadding="2">
                	  <tr>
                	    <td width="79%" bgcolor="#EFEFEF"><strong>
                	      <%=rs2("DESCRICAO")%>
                	      
                	      </strong></td>
      <td width="21%" bgcolor="#EFEFEF"><div align="right"><strong><%=rs2("QTD_POR_CONTRATO")%></strong></div></td>
    </tr>
                    </table>
                	<%
							Set rs3=objConexao.Execute("SELECT TB_PEDIDOS_PRODUTOS.*, SUM(TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS QTD_POR_FORMANDO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, TB_PEDIDOS.COD_VENDEDOR, TB_PEDIDOS.DATA_VENDA FROM TB_PEDIDOS INNER JOIN TB_PEDIDOS_PRODUTOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO INNER JOIN TB_FICHAS_IDENTIFICACAO ON TB_PEDIDOS_PRODUTOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO WHERE TB_PEDIDOS.DATA_VENDA BETWEEN "&sqlData(DATA_INICIO)&" AND "&sqlData(DATA_FIM)&" AND COD_PRODUTO = '"&rs("COD_PRODUTO")&"' AND TB_FICHAS_IDENTIFICACAO.COD_CONTRATO = '"&rs2("COD_CONTRATO")&"' AND TB_PEDIDOS.FINALIZA_CORRECAO = 1 GROUP BY COD_FORMANDO")
							%>
                   			 <table width="100%" border="0" cellspacing="0" cellpadding="2">
                             	<%While not rs3.Eof%>
  <tr>
    <td width="58%" style="border-bottom:solid 1px #CCCCCC"><a href="dip_busca_detalhamento_formando.asp?cod_formando=<%=rs3("COD_FORMANDO")%>" target="_blank"><%=digito(rs3("COD_FORMANDO"))%></a>&nbsp;&nbsp;<%=rs3("NOME_FORMANDO")%></td>
    <td width="26%" style="border-bottom:solid 1px #CCCCCC"><div align="left"><a href="lca_busca_detalhamento_pedido.asp?cod_pedido=<%=rs3("COD_PEDIDO")%>" target="_blank"><%=rs3("COD_PEDIDO")%></a></div></td>
    <td width="16%" style="border-bottom:solid 1px #CCCCCC"><div align="right"><%=rs3("QTD_POR_FORMANDO")%></div></td>
    </tr>
    							<%								
								rs3.MoveNext()
								Wend
								%>
                             </table>
                   			 <%
							rs3.Close
							Set rs3=Nothing%>
                      	     <hr />
                     <%
					 rs2.MoveNext()
					 Wend
					 
					 rs2.Close
					 Set rs2=Nothing
					 %>        
                             </td>
            </tr>
        </table>
        <p>&nbsp;</p>
        <%Else%>
        		
                <div style="padding:50px; font-size:16px;" align="center">Não foi encontrado resultados para os parametros informados de <strong><%=ArrumaData(DATA_INICIO)%></strong> à <strong><%=ArrumaData(DATA_FIM)%></strong></div>
                
        <%End If
		
		rs.Close
		Set rs=Nothing	
			
		FecharConexao()
		%>
		   
		    

		<div align="center"><br />
            <br />
            <input class="btnTopo" type="button" onclick="abre('#')" name="Submit2" value=" " />
            <br />
            <br />
            <br />
		  </div></td>
    </tr>
</table>
<!-- #include file="files/includes/miolo-inferior.asp" -->





