<div class="subMenu">Submenu:
    <% AbrirConexao() %>
    <a href="ppv_ordens_servico.asp">Ordens de Servi&ccedil;o</a>
    	 | 
    <%If validaUsuario(2,,54,true) Then%>
        <a href="ppv_produtos_etapas.asp">Produtos Etapas</a>
    <%Else%>    
     	<button type="button" class="ppv-access-trigger" aria-haspopup="dialog">Produtos Etapas</button>
    <%End If%>
    <!--
     | 
    <%If acesso_inclusao_exclusao_os_manual Then%>
    	<a href="ppv_ordens_servico_formulario.asp?comando=incluir">Cadastrar OS Manual Sem Origem</a>
    <%Else%>    
      	<button type="button" class="ppv-access-trigger" aria-haspopup="dialog">Cadastrar OS Manual Sem Origem</button>				
    <%End If%>
     -->
     |
    <%If validaUsuario(2,,52,true) Then%>
    	   <a href="ppv_ordens_servico_triagem_inicial.asp">Triagem Inicial</a>
    <%Else%>    
       <button type="button" class="ppv-access-trigger" aria-haspopup="dialog">Triagem Inicial</button>
    <%End If%>
     | 
    <%If validaUsuario(2,,53,true) Then%>
    	<a href="ppv_ordens_servico_triagem_despacho.asp">Triagem Despacho</a>| 
    	<a href="ppv_ordens_servico_triagem_entrega.asp">Triagem Entrega</a>                    	
    <%Else%>    
      	<button type="button" class="ppv-access-trigger" aria-haspopup="dialog">Triagem Despacho</button>| 
    	<button type="button" class="ppv-access-trigger" aria-haspopup="dialog">Triagem Entrega</button>                    		                  	
    <%End If%>
     | 
    <%If validaUsuario(2,,58,true) Then%>
    	<a href="ppv_ordens_servico_os_manual_origem_vendas.asp">Inser&ccedil;&atilde;o de OS Manual com Origem em Pedido</a>
    <%Else%>    
      	<button type="button" class="ppv-access-trigger" aria-haspopup="dialog">Inser&ccedil;&atilde;o de OS Manual com Origem em Pedido</button>
    <%End If%>                   
     | 
    <%If validaUsuario(2,,53,true) Then%>
        <a href="ppv_ordens_servico_triagem_entrega_correios.asp">Retorno Correios</a>
    <%Else%>
        <button type="button" class="ppv-access-trigger" aria-haspopup="dialog">Retorno Correios</button>
    <%End If%>
    </div>
<% fecharConexao() %>
<style>
.subMenu .ppv-access-trigger {
    appearance:none; display:inline; padding:0; border:0; background:none;
    color:inherit; font:inherit; text-decoration:underline; cursor:pointer;
}
.subMenu .ppv-access-trigger:hover { color:#88134f; }
.ppv-access-dialog {
    box-sizing:border-box; width:min(440px, calc(100% - 32px)); padding:24px;
    border:1px solid #e4dfeb; border-radius:14px; color:#242738;
    background:#fff; box-shadow:0 18px 60px #24273833; font:1rem/1.5 system-ui,sans-serif;
}
.ppv-access-dialog::backdrop { background:#24273899; }
.ppv-access-dialog h2 { margin:0; color:#242738; font:700 1.2rem/1.3 system-ui,sans-serif; }
.ppv-access-dialog p { margin:10px 0 20px; }
.ppv-access-dialog form { display:flex; justify-content:flex-end; margin:0; }
.ppv-access-dialog button {
    padding:9px 17px; border:0; border-radius:7px; color:#fff; background:#b20b61;
    font:600 .875rem/1.5 system-ui,sans-serif; cursor:pointer;
}
.ppv-access-dialog button:hover { background:#88134f; }
</style>
<dialog class="ppv-access-dialog" id="ppvAccessDialog" aria-labelledby="ppvAccessTitle" aria-describedby="ppvAccessMessage">
    <h2 id="ppvAccessTitle">Acesso restrito</h2>
    <p id="ppvAccessMessage">Seu perfil n&atilde;o tem permiss&atilde;o para acessar esta p&aacute;gina.</p>
    <form method="dialog"><button autofocus>Entendi</button></form>
</dialog>
<script>
(function () {
    var dialog = document.getElementById('ppvAccessDialog');
    document.querySelectorAll('.subMenu .ppv-access-trigger').forEach(function (button) {
        button.addEventListener('click', function () {
            dialog.showModal();
        });
    });
})();
</script>
