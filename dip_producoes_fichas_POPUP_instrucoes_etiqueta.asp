<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Cadastro de Produto em pedido</title>
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<style type="text/css">
<!--
body {
	background-color: #FFFFFF;
}
-->
</style></head>
<body>
<form name="dinheiro" style="margin:0px" action="?cod_pedido=<%=COD_PEDIDO%>&comando=<%=COMANDO%>" target="_parent" onsubmit="return verificaForm()" method="post">
<br />
 <table width="402" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
              <td width="142">
                <div align="center"><strong>Digito Final</strong> </div></td>
              <td width="246"><div align="center"><strong>Produto</strong></div></td>
            </tr>
			<tr>
              <td width="142">
                <div align="center">0</div></td>
              <td width="246"><div align="center">Álbum (Padrão) </div></td>
            </tr>
			<tr>
			  <td><div align="center">1</div></td>
			  <td><div align="center">Ficha (Padrão) </div></td>
    </tr>
			<%
			Dim rs, sql
			AbrirConexao()
			sql = 	" SELECT TB_PRODUTOS.NOME_PRODUTO, TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.TIPO FROM "&_
					" TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO WHERE "&_
					" TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUCAO="&Request.QueryString("cod_prod")&_
				    " GROUP BY TIPO ORDER BY TIPO"			
			Set rs = objConexao.Execute(sql)
			If not rs.Eof Then
				While not rs.Eof
			%>
			<tr>
			  <td><div align="center"><%=rs("TIPO")%></div></td>
			  <td><div align="center"><%=rs("NOME_PRODUTO")%></div></td>
    		</tr>
			<%
			rs.MoveNext()
			Wend
			
			End If
			rs.Close
			Set rs = Nothing
			
			FecharConexao()
			%>
  </table>
</form>
</body>
</html>



