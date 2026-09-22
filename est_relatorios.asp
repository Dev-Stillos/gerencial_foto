<%
Dim calendario_necessario, BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
calendario_necessario = "SIM"
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 1
BTN_IMPRIMIR_LINK = "#"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<h2>Relat&oacuterios</h2>
<!-- #include file="includes/estoque_sub_menu.asp" -->
<h2><%=titulo_pagina%></h2><br />

<div class="subMenu" style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC">
	<form action="est_relatorios_itens_vendedores.asp" name="FORMULARIO_FILTRO" method="get" target="_blank">
		<table cellpadding="4">
			<tr>
				<td colspan="2" align="center"><h3>Itens com Vendedores</h3></td>
			</tr>
			<tr>
				<td align="left"><strong>Vendedor</strong></td>
				<td>
					<%=padraoFormularioSelect("cod_vendedor","","inputTamanho2","","","Todos", "","SELECT cod_vendedor, nome_vendedor FROM TB_VENDEDORES	 order by nome_vendedor", "incluir", false, false, false)%>
				</td>
			</tr>
			<tr>
				<td align="left"><strong>Produto</strong></td>
				<td>
					<%=padraoFormularioSelect("cod_produto","","inputTamanho2","","","Todos", "","SELECT cod_produto, nome_produto FROM TB_PRODUTOS order by nome_produto", "incluir", false, false, false)%>
				</td>
			</tr>
			<tr>
				<td><strong>Filtrar periodo por</strong></td>
				<td>
					<table>
						<tr>
							<td><input type="radio" class="padraoradio" name="periodo_parametro" value="devolucao" id='campo_periodo_id_d' checked=""></strong></td>
							<td><label for='campo_periodo_id_d'>Data Devolu&ccedil&atildeo</label></td>
						</tr>
						<tr>
							<td><input type="radio" class="padraoradio" name="periodo_parametro" value="retirada" id='campo_periodo_id_r'></td>
							<td><label for='campo_periodo_id_r'>Data Retirada</label></td>
						</tr>
					</table>

				</td>
			</tr>
			<tr>
				<td align="left"><strong>Per&iacuteodo</strong></td>
				<td>
					<input type="text" class="inputTamanho1 padrao_data" name="periodo_inicio">&agrave<input type="text" class="inputTamanho1 padrao_data" name="periodo_fim">
				</td>
			</tr>

			<tr>
				<td colspan="2" align="center">
					<input type="submit" class="btnFiltrarAmarelo" name="filtro" value=" ">
				</td>
			</tr>
		</table>
	</form>
</div>
<br />
<div class="subMenu" style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC">
	<form action="est_extrato.asp" name="FORMULARIO_FILTRO" method="get" target="_blank">
		<table cellpadding="4">
			<tr>
				<td colspan="2" align="center"><h3>Extrato de Produto</h3></td>
			</tr>
			<tr>
				<td align="left"><strong>Produtos</strong></td>
				<td>
					<%=padraoFormularioSelect("cod_produto","","inputTamanho2","","","Todos", "","SELECT cod_produto, nome_produto FROM TB_PRODUTOS order by nome_produto", "incluir", false, false, false)%>
				</td>
			</tr>
			<tr>
				<td align="left"><strong>Per&iacuteodo</strong></td>
				<td>
					<input type="text" class="inputTamanho1 padrao_data" name="periodo_inicio">&agrave<input type="text" class="inputTamanho1 padrao_data" name="periodo_fim">
				</td>
			</tr>
			<tr>
				<td colspan="2" align="center">
					<input type="submit" class="btnFiltrarAmarelo" name="filtro" value=" ">
				</td>
			</tr>
		</table>
	</form>
</div>
<br />
<!--
<div class="subMenu" style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC">
	<form action="est_relatorios_." name="FORMULARIO_FILTRO" method="get">
		<table cellpadding="4">
			<tr>
				<td colspan="2" align="center"><h3>Posi&ccedil&atildeo Geral</h3></td>
			</tr>
			<tr>
				<td align="left"><strong>Vendedor</strong></td>
				<td>
					<%=padraoFormularioSelect("cod_produto","","inputTamanho2","","","Todos", "","SELECT cod_produto, nome_produto FROM TB_PRODUTOS order by nome_produto", "incluir", false, false, false)%>
				</td>
			</tr>
			<tr>
				<td align="left"><strong>Per&iacuteodo</strong></td>
				<td>
					<input type="text" class="inputTamanho1 padrao_data" name="periodo_inicio">&agrave<input type="text" class="inputTamanho1 padrao_data" name="periodo_fim">
				</td>
			</tr>
			<tr>
				<td colspan="2" align="center">
					<input type="submit" class="btnFiltrarAmarelo" name="filtro" value=" ">
				</td>
			</tr>
		</table>
	</form>
</div>

<br />

<div class="subMenu" style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC">
	<form action="est_relatorios_." name="FORMULARIO_FILTRO" method="get">
		<table cellpadding="4">
			<tr>
				<td colspan="2" align="center"><h3>Entrada X Sa&iacuteda</h3></td>
			</tr>
			<tr>
				<td align="left"><strong>Vendedor</strong></td>
				<td>
					<%=padraoFormularioSelect("cod_produto","","inputTamanho2","","","Todos", "","SELECT cod_produto, nome_produto FROM TB_PRODUTOS order by nome_produto", "incluir", false, false, false)%>
				</td>
			</tr>
			<tr>
				<td align="left"><strong>Per&iacuteodo</strong></td>
				<td>
					<input type="text" class="inputTamanho1 padrao_data" name="periodo_inicio">&agrave<input type="text" class="inputTamanho1 padrao_data" name="periodo_fim">
				</td>
			</tr>
			<tr>
				<td colspan="2" align="center">
					<input type="submit" class="btnFiltrarAmarelo" name="filtro" value=" ">
				</td>
			</tr>
		</table>
	</form>
</div>
-->
<br />
<br />
<br />
<!-- #include file="files/includes/miolo-inferior.asp" -->
</body>
</html>