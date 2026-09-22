<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<%
If not validaUsuario(2,,77,false) Then
	Response.Redirect("inicio.asp")
End If

Dim data_inicial : data_inicial = Cstr(""&Request.QueryString("data_inicial"))
Dim data_final : data_final = Cstr(""&Request.QueryString("data_final"))

If data_inicial = "" Then
	data_inicial = date()
End If

If data_final = "" Then
	data_final = date()
End If
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="all"></link>
<title>Relatório de Despacho</title>
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
                <h3 align="center">Relat&oacute;rio de Sa&iacute;da de produtos produzidos p&oacute;s-venda - Despacho<br />
                 <%=ArrumaData(data_inicial)%> a <%=ArrumaData(data_final)%>
                </h3>
&nbsp;</div>
                <%
			AbrirConexao()
			
			Dim qtd_registros			
			sql = "SELECT cod_ordem_servico, NOME_FORMANDO,ENTIDADE_ENSINO,data_despachado,NOME_PRODUTO,prazo_geral,quantidade,cod_usuario_despachou,despacho_transportador  FROM VW_TB_ORDENS_SERVICO WHERE status_ordem_servico = 3 AND data_despachado BETWEEN "&sqlData(data_inicial&" 00:00:00")&" AND "&sqlData(data_final&" 23:23:59")&" ORDER BY data_despachado ASC"
			
			'rESPONSE.Write(sql)
			Set rs=objConexao.Execute(sql,qtd_registros)
			
			
			If not rs.Eof Then%>
                <table width="95%" cellspacing="0" border="0" align="center" 	cellpadding="5">
  <tr bgcolor="whitesmoke">
    <td width="7%"><div align="center"><strong>OS</strong></div></td>
    <td width="22%"><div align="center"><strong>Formando</strong><em></em>
    </div></td>
    <td width="18%"><div align="center"><strong>Institui&ccedil;&atilde;o de Ensino</strong></div></td>
    <td width="19%"><div align="center"><strong>Produto - Qtd.</strong></div></td>
    <td width="7%"><div align="center"><strong>Prazo Geral</strong></div></td>
    <td width="10%"><div align="center"><strong>Transportador</strong></div></td>
    <td width="9%"><div align="center"><strong>Resp. Despacho</strong></div></td>
    <td width="8%"><div align="center"><strong>Data Despacho</strong></div></td>
    </tr>
                  <%
		  	While not rs.Eof
		  %>
                  <tr>
                    <td style="border-bottom:solid 1px #CCC"><div align="center"><%=moduloPPVOSMascara(rs("cod_ordem_servico"),"",true)%></div></td>
                    <td style="border-bottom:solid 1px #CCC"><div align="left"><%=rs("nome_formando")%></div></td>
                    <td style="border-bottom:solid 1px #CCC"><div align="left"><%=rs("ENTIDADE_ENSINO")%></div></td>
                    <td style="border-bottom:solid 1px #CCC"><div align="left"><%=rs("nome_produto")%> - <%=rs("quantidade")%></div></td>
                    <td style="border-bottom:solid 1px #CCC"><div align="center"><%=ArrumaData(rs("prazo_geral"))%></div></td>
                    <td style="border-bottom:solid 1px #CCC"><div align="left"><%=rs("despacho_transportador")%></div></td>
                    <td style="border-bottom:solid 1px #CCC"><div align="left"><%=nomeUsuario(rs("cod_usuario_despachou"),true)%></div></td>
                    <td style="border-bottom:solid 1px #CCC"><div align="center"><%=ArrumaData(rs("data_despachado"))%>
                 
                    </div></td>
                    </tr>
                  <%
				rs.MoveNext()
				Wend
		  %>
           			<tr>
                    	<td colspan="8" style="border-bottom:solid 1px #CCC; font-size:22px;"><div align="right"><strong>Total &nbsp;&nbsp;&nbsp;<%=qtd_registros%></strong></div></td>
                    </tr>
                  </table>
                <br />
<br />

                Impresso em: <%=Now()%>
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