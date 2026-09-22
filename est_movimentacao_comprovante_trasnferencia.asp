<%
Dim calendario_necessario, BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
calendario_necessario = "SIM"
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 1
BTN_IMPRIMIR_LINK = "#"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->

<h2>Estoque</h2>

<%
	AbrirConexao()
	
	Dim sql, sqlWhere, sqlOrderBy
	sql = "SELECT * FROM vw_est_movimentacao_transferencias "
	sqlOrderBy = " ORDER BY data_transferencia desc"

	sql = sql&sqlWhere&sqlOrderBy
	%>
	<h2><%=titulo_pagina%></h2><br />

	<%
	set rs = objConexao.execute(sql)
	if not rs.Eof then
	%>
	<table width="99%" border="1" align="center" cellpadding="5" style="border-collapse:collapse; text-align: right;">
		<tr bgcolor="whitesmoke">
			<td width="7%"><div align="center" style="font-weight: bold">C&oacute;d. Pedido Compra</div></td>
			<td width="20%"><div align="center" style="font-weight: bold">De:</div></td>
			<td width="20%"><div align="center"><strong>Para:</strong></div></td>
			<td width="8%"><div align="center"><strong>Data:</strong></div></td>
			<td width="10%" align="center"><div align="center"><strong>Quantidade - Itens:</strong></div></td>
			<td width="5%" align="center">&nbsp;</td>
		</tr>
		<%
		while not rs.Eof
		%>
			<tr>
  				<td><div align="center"><%=rs("cod_transferencia")%></div></td>
  				<td><div align="left"><%=rs("local_armazenamento_origem")%></div></td>
  				<td><div align="left"><%=rs("local_armazenamento_destino")%></div></td>
				<td><div align="center"><%=ArrumaData(left(rs("data_transferencia"),10))%></div></td>
				<td><div align="center">
				<%
				set rs2 = objConexao.execute("SELECT tb_est_itens_movimentacao.* , `TB_PRODUTOS`.`NOME_PRODUTO` FROM tb_est_itens_movimentacao JOIN `stillos3`.`TB_PRODUTOS` ON `TB_PRODUTOS`.`COD_PRODUTO` = `tb_est_itens_movimentacao`.`cod_item` WHERE tb_est_itens_movimentacao.cod_movimentacao = "&sqlNumero(rs("cod_movimentacao_destino")))
				while not rs2.Eof 
					response.write("(" & rs2("Quantidade") & ") "& rs2("NOME_PRODUTO"))
				rs2.movenext()
				Wend
				rs2.close
				set rs2 = nothing

				%>
				</div></td>
 
				<td>
					<table border="0" align="center"> 
						<tr>
							<td>
								<a target='_blank' href="est_transtefencia_comprovante.asp?cod_transferencia=<%=rs("cod_transferencia")%>">
									<img src="images/icons/icon_visualizar.png" alt="Visualizar Registro" title="Visualizar Registro" border="0" />
								</a>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		<%
		rs.movenext
		Wend
	%>
		</table>
	<%
	Else
	%>
    <div align="center" class="msgGrande">
    	Nenhuma Transfer&ecircncia encontrada!</div>
    <%End if

	rs.Close
	Set rs=Nothing

	FecharConexao()
	%>
<!-- #include file="files/includes/miolo-inferior.asp" -->
</body>
</html>