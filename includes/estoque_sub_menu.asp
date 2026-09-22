<div class="subMenu">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td width="4%">Submenu:</td>
			<td width="96%">
			<%AbrirConexao()%>
				<a href="est_resumo.asp">Resumo Estoque</a> | 
				<a href="est_movimentacao_manual.asp">Movimentações</a> | 
				<a href="est_relatorios.asp">Relatários</a> | 
				<a href="est_renovacao_produtos.asp">Renovação de Produtos</a>
				<%'If validaUsuario(2,,61,true) Then%>
					<!--<a href="ppv_compras_pedidos_formulario.asp">Novo Pedido de Compra (manual) | </a>-->
				<%'Else%>
					<!--<a href="javascript:alert('Voc� n�o possui acesso a esta p�gina')">Novo Pedido de Compra (manual) | </a>-->
				<%'End If%>
			<%fecharConexao()%>
			</td>
		</tr>
	</table>
</div>