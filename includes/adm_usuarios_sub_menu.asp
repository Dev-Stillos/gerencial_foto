<div class="subMenu">Sub-Menu: <a href="adm_usuarios.asp">Listagem de Usu&aacute;rios e Vendedores</a> | <a href="adm_usuarios_incluir.asp">Incluir Usu&aacute;rio do Sistema</a> | <a href="adm_vendedores_formulario.asp">Incluir Vendedor</a> | <a href="adm_perfis.asp">Perfis de Acesso</a> | <a href="adm_perfis_formulario.asp">Cadastrar Perfil</a> | <a href="adm_configuracoes_gerais.asp">Configura&ccedil;&atilde;o de Acesso</a>
<%
If Session("cod_usuario") = 1 Then
%>
 | <a href="adm_restrito_paginas.asp">Valida&ccedil;&atilde;o de P&aacute;ginas</a>
<%
End If
%>
</div>