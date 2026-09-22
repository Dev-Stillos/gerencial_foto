<div class="subMenu">Submenu:
    <% AbrirConexao() %>
    <a href="ppv_ordens_servico.asp">Ordens de Servi&ccedil;o</a>
    	 | 
    <%If validaUsuario(2,,54,true) Then%>
        <a href="ppv_produtos_etapas.asp">Produtos Etapas</a>
    <%Else%>    
     	<a href="javascript:alert('Voc&ecirc; n&atilde;o possui acesso a esta p&aacute;gina')">Produtos Etapas</a>
    <%End If%>
    <!--
     | 
    <%If acesso_inclusao_exclusao_os_manual Then%>
    	<a href="ppv_ordens_servico_formulario.asp?comando=incluir">Cadastrar OS Manual Sem Origem</a>
    <%Else%>    
      	<a href="javascript:alert('Voc� n�o possui acesso a esta p�gina')">Cadastrar OS Manual Sem Origem</a>				
    <%End If%>
     -->
     |
    <%If validaUsuario(2,,52,true) Then%>
    	   <a href="ppv_ordens_servico_triagem_inicial.asp">Triagem Inicial</a>
    <%Else%>    
       <a href="javascript:alert('Voc&ecirc; n&atilde;o possui acesso a esta p&aacute;gina')">Triagem Inicial</a>
    <%End If%>
     | 
    <%If validaUsuario(2,,53,true) Then%>
    	<a href="ppv_ordens_servico_triagem_despacho.asp">Triagem Despacho</a>| 
    	<a href="ppv_ordens_servico_triagem_entrega.asp">Triagem Entrega</a>                    	
    <%Else%>    
      	<a href="javascript:alert('Voc&ecirc; n&atilde;o possui acesso a esta p&aacute;gina')">Triagem Despacho</a>| 
    	<a href="javascript:alert('Voc&ecirc; n&atilde;o possui acesso a esta p&aacute;gina')">Triagem Entrega</a>                    		                  	
    <%End If%>
     | 
    <%If validaUsuario(2,,58,true) Then%>
    	<a href="ppv_ordens_servico_os_manual_origem_vendas.asp">Inser&ccedil;&atilde;o de OS Manual com Origem em Pedido</a>
    <%Else%>    
      	<a href="javascript:alert('Voc&ecirc; n&atilde;o possui acesso a esta p&aacute;gina')">Inser&ccedil;&atilde;o de OS Manual com Origem em Pedido</a>
    <%End If%>                   
     | 
    <%If validaUsuario(2,,53,true) Then%>
        <a href="ppv_ordens_servico_triagem_entrega_correios.asp">Retorno Correios</a>
    <%Else%>
        <a href="javascript:alert('Voc&ecirc; n&atilde;o possui acesso a esta p&aacute;gina')">Retorno Correios</a>
    <%End If%>
    </div>
<% fecharConexao() %>