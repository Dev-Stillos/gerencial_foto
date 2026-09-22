<div class="subMenu">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td width="4%">Submenu:</td>
			<td width="96%">
			<%AbrirConexao()%>
				<a href="mck_formandos.asp">Formandos</a> | 
				<a href="mck_pedidos.asp">Pedidos</a>
				<%'If 'validaUsuario(2,,61,true) Then%>
				<% if session("cod_usuario") = 1 Then 	%>
					 | <a href="mck_contatos.asp">Tipos de contato (somente administrador) </a>
				<%Else%>
					<!--<a href="javascript:alert('Você não possui acesso a esta página')">Novo Pedido de Compra (manual) | </a>-->
				<%End If%>
			<%fecharConexao()%>
			</td>
		</tr>
	</table>
</div>