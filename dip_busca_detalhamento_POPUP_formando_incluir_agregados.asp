<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Para tirar a validação da página basta tirar o "#" ou coloca-lo para habilitar antes do include
Dim COD_FORMANDO
COD_FORMANDO = Request.QueryString("cod_form")


If Request.Form("btnGravar") <> Empty then
'Inclui
	'1) Inclui na TABELA de AGREGADOS
	'2) Incluir na TABELA MOVIMENTAÇÕES
	'3) FAZ UPDATE COM PRODUTO_AGREGADO
	
	'Disponível somente FOTO-TELA e SLIDE SHOW
	
	AbrirConexao()
	
	If Request.Form("COD_PRODUTO") <> "" then
	
	COD_FORMANDO = Request.Form("COD_FORMANDO")
	Response.Write("Cód Formando:"&COD_FORMANDO&"<br />")
	
	Dim COD_PRODUTO_INCLUIR, TIPO
	COD_PRODUTO_INCLUIR = formatNumber(0&Request.Form("COD_PRODUTO"),0)
	'SOMENTE 2 PRODUTOS HABILITADOS
	If int(COD_PRODUTO_INCLUIR) = 9 Then
		'Slide Show
		TIPO = 2
	Elseif int(COD_PRODUTO_INCLUIR) = 16 Then 
		'Foto-tela
		TIPO = 3
	End If
	
	'1)
	sql = 	"INSERT INTO TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ("&_
			"TIPO ,"&_
			"COD_PRODUCAO ,"&_
			"COD_FORMANDO ,"&_
			"COD_PRODUTO ,"&_
			"QUANTIDADE ,"&_
			"CADASTRADO_POR "&_
			") VALUES ( "&_
			TIPO&",'0',"&Request.Form("COD_FORMANDO")&", "&COD_PRODUTO_INCLUIR&", 1,"&Session("COD_USUARIO")&_
			")"
			
			
			'Response.Write(sql)
			objConexao.Execute(sql)
	
	'ANTES É NECESSÁRIO SABER QUAL O VENDEDOR
	''/////
			Dim rsVendedor, COD_VENDEDOR
			sql = "SELECT * FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_FORMANDO ="&Request.Form("COD_FORMANDO")&" AND TIPO = 1 ORDER BY DATA_SAIDA DESC LIMIT 1"
			'Response.Write(sql)
			Set rsVendedor = objConexao.Execute(sql)
					
					If not rsVendedor.Eof Then
							COD_VENDEDOR = rsVendedor("COD_VENDEDOR")			
					End If
					
			
			rsVendedor.Close
			Set rsVendedor = Nothing
	''/////
	
	'2)
		sql="INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ("&_
		"COD_FORMANDO ,"&_
		"TIPO ,"&_
		"COD_VENDEDOR ,"&_
		"DATA_SAIDA ,"&_
		"DATA_RETORNO ,"&_
		"CADASTRADO_POR "&_
		")"&_
		"VALUES ("&_
		Request.Form("COD_FORMANDO")&", "&TIPO&", "&COD_VENDEDOR&",NOW( ),NOW( ),"&Session("COD_USUARIO")&")"
	
		Response.Write(sql)
		objConexao.Execute(sql)
	

		'3)	
			objConexao.Execute("UPDATE TB_FICHAS_IDENTIFICACAO SET PRODUTO_AGREGADO = '1' WHERE COD_FORMANDO = "&Request.Form("COD_FORMANDO")&" LIMIT 1")
				

	
	FecharConexao()	
	Response.Redirect("dip_busca_detalhamento_POPUP_formando_incluir_agregados.asp?cod_form="&Request.Form("COD_FORMANDO"))
	
	End If

End If




%>
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<style type="text/css">
<!--
body {
	background-color: #FFFFFF;
}
-->
</style>
<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td bgcolor="#FFFFFF"><h1 style="margin:0px; padding:0px">DIP - Formandos / Produto agregado </h1>
          <div align="left">Cód. Formando.: <strong><%=digito(COD_FORMANDO)%></strong>		 </div>
	<form name="FORMULARIO"  method="post" action="dip_busca_detalhamento_POPUP_formando_incluir_agregados.asp?cod_formando=<%=COD_FORMANDO%>&comando=incluir" style="margin:0px">
		   <table width="320" border="0" align="center" cellpadding="3" cellspacing="0" style="margin-top:5px; border-bottom:2px #3366CC solid; border-top:#3366CC solid 2px">
             <tr>
               <td><div align="center">Incluir Produto Agregado</div>
			   <input type="hidden" name="COD_FORMANDO" value="<%=COD_FORMANDO%>" /> </td>
               </tr>
			   <tr>
                 <td><div align="center">
                   <div align="center">
                     <select onclick="sair('sim')" name="COD_PRODUTO" onchange="sair('sim')" onfocus="sair('sim')">
                       <option value="0">---</option>
                       <%
				 AbrirConexao()
				 Dim rsProdutos, sql
				 
				 
	
				 	SQL = "SELECT * FROM TB_PRODUTOS WHERE COD_PRODUTO = 9 OR COD_PRODUTO = 16"
				 				 				 
				 
				 Set rsProdutos = objConexao.Execute(SQL)
				 
				 While not rsProdutos.Eof
				 %>
                       <option value="<%=rsProdutos("COD_PRODUTO")%>"><%=rsProdutos("NOME_PRODUTO")%></option>
                       <%
				 rsProdutos.MoveNext()
				 Wend
				 %>
                       <%
				rsProdutos.Close()
				Set rsProdutos = Nothing
				
				
				
				
				
				FecharConexao()
				%>
                     </select>
                   </div>
                 </div></td>
                </tr>
             <tr>
               <td><div align="center">
                 
                 <input class="btnGravar" onClick="anula()" type="submit" name="btnGravar" value=" " />
                 
               </div></td>
               </tr>
           </table>
	</form>
		 <%
		 AbrirConexao()
		 	 Dim rs
			 sql = "SELECT TB_PRODUTOS.NOME_PRODUTO, Sum( TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.QUANTIDADE ) AS QUANTIDADE"&_
																	" FROM TB_PRODUTOS"&_
																	" INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO"&_
																	" WHERE TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO = "&COD_FORMANDO&_
																	" GROUP BY TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO ORDER BY TB_PRODUTOS.NOME_PRODUTO"
			 
			 	 
			 Set rs = objConexao.Execute(sql)
		 
		 	 If not rs.Eof then
		 %>
		 <table width="98%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse; margin-top:6px">
           <tr bgcolor="whitesmoke">
             <td width="66%"><div align="center"><strong>Produto Agregado </strong></div></td>
             </tr>
           <%While not rs.Eof%>
           <tr>
             <td><div align="left"><%=rs("NOME_PRODUTO")%></div></td>
             </tr>
           <%
			rs.MoveNext()
			Wend
			%>
         </table>
		 <%
		  	 End If
			
			
			  rs.Close()
			  Set rs = Nothing
		  FecharConexao()
		  %>		  </td>
      </tr>
    </table></td>
  </tr>
</table>


</body>
</html>



