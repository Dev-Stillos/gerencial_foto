<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/classe_fechamento.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<style type="text/css">
<!--
body,td,th {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 11px;
}
-->
</style></head>
<body>
<%
	AbrirConexao()
		
		Set rs = Objconexao.Execute("SELECT * FROM TB_MACK_PEDIDOS")
		while not rs.Eof 

			'SE FOR IGUAL A UMA PARCELA VERIFICAMOS SE HÁ DESCONTO
			if int(rs("qtd_parcelas")) = 1 then

				Set rs2 = Objconexao.Execute("SELECT desconto_avista FROM TB_MACK_FORMAS_PAGAMENTO where cod_mack_forma_pagamento = "&sqlNumero(rs("cod_mack_forma_pagamento")))
				if not rs2.eof then
					desconto = (100-int(0&rs2("desconto_avista")))/100
				end if
				rs2.close
				set rs2 = nothing
	
				if desconto <> 0 then
					'Objconexao.execute("UPDATE TB_MACK_PEDIDO_PRODUTOS SET valor_unitario = valor_unitario * "&sqlNumero(desconto)" WHERE cod_mack_pedido = "&sqlNumero(rs2("cod_mack_pedido")))

					response.write("<br>UPDATE TB_MACK_PEDIDOS_PRODUTOS SET valor_unitario = valor_unitario * "&sqlNumero(desconto)&" WHERE cod_mack_pedido = "&sqlNumero(rs("cod_mack_pedido")))
				end if
			end if

		rs.movenext
		wend
		rs.close
		set rs = nothing




		'set rs = Objconexao.Execute("SELECT COD_PEDIDO FROM TB_PEDIDOS WHERE FINALIZA_CORRECAO = 1 AND FINALIZA_DIRETORIA = 2 and DATA_CORRIGIDO_DEP_FOTO > '2015-09-27' order by DATA_CORRIGIDO_DEP_FOTO")
		'if not rs.Eof then
		''	while not rs.Eof
		''		a = moduloPPVgerarOrdensServico(rs("COD_PEDIDO"),true)
		''		response.write a
		''	rs.movenext
		''	wend
		'end if
		'rs.close
		'set rs = nothing
	FecharConexao()
%>
</body>
</html>