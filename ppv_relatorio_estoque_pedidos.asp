<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<%
If not validaUsuario(2,,77,false) Then
	Response.Redirect("inicio.asp")
End If

Dim data_inicial : data_inicial = Cstr(""&Request.QueryString("data_inicial"))
Dim data_final : data_final = Cstr(""&Request.QueryString("data_final"))
Dim cod_pedido : cod_pedido = ""
Dim status_pedido :  status_pedido = ""
Dim tipo_relatorio : tipo_relatorio = Cstr(""&Request.QueryString("tipo_relatorio"))


cod_pedido = Cstr(""&Request.QueryString("cod_pedido"))
status_pedido = Cstr(""&Request.QueryString("status_pedido"))



If status_pedido <> "0" Then
	status_pedido = " AND status = '"&status_pedido&"' "
else
	status_pedido = ""
End If

If data_inicial = "" Then
	data_inicial = date()
End If
data_inicial = Year(data_inicial)&"-"&Month(data_inicial)&"-"&Day(data_inicial)&" 00:00:00"


If data_final = "" Then
	data_final = date()
End If
data_final = Year(data_final)&"-"&Month(data_final)&"-"&Day(data_final)&" 00:00:00"

sql_filtro = " WHERE data_pedido BETWEEN '"&data_inicial&"' AND '"&data_final&"' "&status_pedido&cod_pedido

If cod_pedido <> "" Then
	cod_pedido = "  cod_pedido = '"&cod_pedido&"' "
	sql_filtro = " WHERE "&cod_pedido

else
	cod_pedido = ""
End If

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="all"></link>
    <title>Relatório de Despacho</title>
	
	
	<style>
	
<%
	if tipo_relatorio  = "S" then
%>
	.hide_line{
		display:none;
	}

<%
	else
%>
	.hide_valor_total{
		display: none;
	}
<%
	end if


%>
		.desc_pedido{
			background-color: #F6F6F6;
		}
		.desc_pedido_total{
			background-color: #F6F6F6;
		}		
	</style>
</head>

<body style="background:#FFF">
<!-- End code for the left panel -->
<table cellpadding="0px" cellspacing="0" bgcolor="#FFFFFF" width="100%">
<tr>
<td>
<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="10">
      <tr>
        <td>
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td valign="top">
                <div align="center"><img src="images/layout/logomarca_foto_reportagens.jpg" width="92" height="93" />
                </div>
                <h3 align="center">
						Relatório de Pedidos (Pré-venda)<br />               
                </h3>

<%
AbrirConexao()

Dim qtd_registros	, cont		
sql = "SELECT A.cod_pedido, C.nome_fantasia,A.status,A.data_pedido, A.data_entrega, (SELECT SUM(B.valor) FROM tb_est_itens_pedidos B WHERE B.cod_pedido = A.cod_pedido ) AS valor, "&_
                    " CAST((SELECT GROUP_CONCAT(CONCAT(D.nome_produto,'::', C.quantidade, '::',C.valor) separator '[[#]]"&_
                    "' ) FROM tb_est_itens_pedidos C "&_
                    "INNER JOIN TB_PRODUTOS D on D.COD_PRODUTO = C.cod_item "&_
                    "WHERE A.cod_pedido = C.cod_pedido ) AS CHAR(100000) CHARACTER SET Latin1 ) as itens "&_
	  "FROM tb_est_pedidos A  INNER JOIN TB_FORNECEDORES C ON A.cod_fornecedor = C.COD_FORNECEDOR "&sql_filtro

'response.Write(sql)
cont = 1
Set rs=objConexao.Execute(sql,qtd_registros)
valor_total = 0

If not rs.Eof Then%>
                <table width="95%" cellspacing="0" border="0" align="center" cellpadding="5" style="border-collapse: collapse;">
						<tr style="background-color:#DDD;color: #424242;">
						  <td width="5%"><div align="center"><strong>#</strong></div></td>
						  <td><div align="left"><strong>Fornecedor</strong><em></em> </div></td>
						  <td><div align="left"><strong>Status</strong><em></em> </div></td>						  
						  <td><div align="left"><strong>Data do Pedido</strong></div></td>
						  <td><div align="left"><strong>Data Entrega</strong></div></td>
						  <td class="hide_valor_total"><div align="left"><strong>Valor total</strong></div></td>

						  
						</tr>
<%
While not rs.Eof
	status = rs("status")
	if status = "EA" then	
		status = "Em Andamento"		
	elseif status = "CN" then	
		status = "Cancelado"
	else	
		status = "Concluído"
	end if
	
	valor_pedido = FormatCurrency(rs("valor"))
%>
                  <tr style="background-color:#E8EFFC;color: #424242;">
                    <td style="border-bottom:solid 1px #CCC"><div align="center"><%=rs("cod_pedido")%></div></td>                  
                    <td style="border-bottom:solid 1px #CCC"><div align="left"><%=rs("nome_fantasia")%></div></td>
                    <td style="border-bottom:solid 1px #CCC"><div align="left"><%=status%></div></td>
                    <td style="border-bottom:solid 1px #CCC"><div align="left"><%=ArrumaData(rs("data_pedido"))%></div></td>				
                    <td style="border-bottom:solid 1px #CCC"><div align="left"><%=ArrumaData(rs("data_entrega"))%></div></td>
					<td style="border-bottom:solid 1px #CCC" class="hide_valor_total"><div align="left"><%=valor_pedido%></div></td>

				  </tr>
                  <tr class="hide_line" bgcolor="">
					<td colspan=2></td>                    
					<td class="desc_pedido"><div align="left"><strong>Nome do Item </strong></div></td>
					<td class="desc_pedido"><div align="left"><strong>Quantidade </strong></div></td>
					<td class="desc_pedido"><div align="left"><strong>Valor </strong></div></td>															
	                
				  </tr>
	<%
	
		lista_itens = split(rs("itens"),"[[#]]")
		for each item in lista_itens
			item_composto = split(item,"::")
			nome_item = item_composto(0)
			quantidade = item_composto(1)
			valor = FormatCurrency(item_composto(2))			
		
	%>
				<tr class="hide_line" >
				  <td colspan=2></td>
				  <td class="desc_pedido" ><div align="left"><%=nome_item%> </div></td>
				  <td class="desc_pedido"><div align="left"><%=quantidade%></div></td>
				  <td class="desc_pedido"><div align="left"> <%=valor%></div></td>								
			    </tr>
	<%
			next
			
	%>
				<tr class="hide_line" bgcolor="">
				  <td colspan=2></td>
				  <td colspan=0 class="desc_pedido_total"></td>			  
				  <td class="desc_pedido_total"><div align="left"><strong>Total do pedido</strong></div></td>
				  <td class="desc_pedido_total"><div align="left"><strong><%=valor_pedido%></strong></div></td>
			    </tr>
				<tr class="hide_line" bgcolor="">
				  <td colspan=5></td>
			    </tr>
<%
valor_total = valor_total + rs("valor")
cont = cint(cont) + 1
rs.MoveNext()
Wend
%>
				  <tr ><td colspan=5></td></tr>
           		  <tr>
					<td colspan="0" style="border-bottom:solid 1px #CCC; font-size:22px;">
					<div align="left">
							<small>Totais</small>
						</div>
					</td>					
				  
					<td colspan="2" style="border-bottom:solid 1px #CCC; font-size:22px;"></td>
                    <td colspan="0" style="border-bottom:solid 1px #CCC; font-size:22px;">
						<div align="left">
							<small>Pedidos: <%=qtd_registros%></small>
						</div>
					</td>
                    <td colspan="0" style="border-bottom:solid 1px #CCC; font-size:22px;">
						<div align="left">
							<small><%=formatcurrency(valor_total)%></small>
						</div>
					</td>					
                  </tr>
				  

				  

		
<br />
<br />
		<tr>
			<td colspan="5">
                Impresso em: <%=Now()%> <br>
                Lançamentos do dia <%=ArrumaData(data_inicial)%> ao <%=ArrumaData(data_final)%>
<%
End If
rs.Close()
Set rs = Nothing
FecharConexao()
%>
                
                <p align="center">&nbsp;</p>
                </td>

              </tr>
		</table>
          
          
          
          </td>
        </tr>
</table>

  
  </td>
  </tr>
</table></td>
</tr>
</table>
</body>
</html>
