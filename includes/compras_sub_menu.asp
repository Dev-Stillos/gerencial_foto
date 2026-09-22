<div class="subMenu">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="4%">Submenu:</td>
    <td width="96%"><%AbrirConexao()%>
   					<a href="ppv_compras_pedidos.asp">Compras</a> |    
                    
                    <%If validaUsuario(2,,61,true) Then%>
                        <a href="ppv_compras_pedidos_formulario.asp">Novo Pedido de Compra (manual) | </a>
                    <%Else%>
                        <a href="javascript:alert('Voc&ecirc; n&atilde;o possui acesso a esta p&aacute;gina')">Novo Pedido de Compra (manual) | </a>
                    <%End If%>

                    <%If validaUsuario(2,,61,true) Then%>
                    	<a href="ppv_fornecedores.asp">Fornecedores</a>
                    <%Else%>
	                  	<a href="javascript:alert('Voc&ecirc; n&atilde;o possui acesso a esta p&aacute;gina')">Fornecedores</a>
					<%End If%>
                     | 
                    <%If validaUsuario(2,,63,true) Then%>
                   	   <a href="ppv_formas_recebimento.asp">Formas de Recebimento</a>
                    <%Else%>
	                   <a href="javascript:alert('Voc&ecirc; n&atilde;o possui acesso a esta p&aacute;gina')">Formas de Recebimento</a>
                    <%End If%>
                     | 
                    <%If validaUsuario(2,,62,true) Then%>
                    	<a href="ppv_formas_pagamento.asp">Formas de Pagamento</a>
                    <%Else%>
	                  	<a href="javascript:alert('Voc&ecirc; n&atilde;o possui acesso a esta p&aacute;gina')">Formas de Pagamento</a>
                    <%End If%>

                    <br />

                    <%If validaUsuario(2,,60,true) Then%>
                    	<a href="ppv_ordens_servico_facilitador_geracao_pedidos_compra.asp">Gera&ccedil;&atilde;o de Pedidos de Compra com origem em Ordens de Servi&ccedil;o</a>
                    <%Else%>
	                  	<a href="javascript:alert('Voc&ecirc; n&atilde;o possui acesso a esta p&aacute;gina')">Gera&ccedil;&atilde;o de Pedidos de Compra com origem em Ordens de Servi&ccedil;o</a>
                    <%End If%>

                    <%'------------------------------------------------------%>
                     | 
                    <%If validaUsuario(2,,60,true) Then%>
                    	<a href="ppv_compras_pedidos_produtos_solicitados.asp">Produtos Solicitados</a>
                    <%Else%>
	                  	<a href="javascript:alert('Voc&ecirc; n&atilde;o possui acesso a esta p&aacute;gina')">Produtos Solicitados</a>
                    <%End If%>
                 </td>
  </tr>
  </table>
</div>