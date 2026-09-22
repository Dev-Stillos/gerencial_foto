<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<%
If not validaUsuario(2,,77,false) Then
  Response.Redirect("inicio.asp")
End If

Dim cods_pedido
Dim data_inicial : data_inicial = Cstr(""&Request.QueryString("data_inicial"))
Dim data_final : data_final = Cstr(""&Request.QueryString("data_final"))

If data_inicial = "" Then
  data_inicial = date()
End If

If data_final = "" Then
  data_final = date()
End If

Dim datas : datas = " ENT.data BETWEEN "&sqlData(data_inicial&" 00:00:00")&" AND "&sqlData(data_final&" 23:23:59 ")
Dim datas2 : datas2 = " data_entregue BETWEEN "&sqlData(data_inicial&" 00:00:00")&" AND "&sqlData(data_final&" 23:23:59 ")
Dim datas3 : datas3 =  " data_entregue BETWEEN "&sqlData(data_inicial&" 00:00:00")&" AND "&sqlData(data_final&" 23:23:59 ")

Dim data_inicial_cadastro : data_inicial_cadastro = Cstr(""&Request.QueryString("data_inicial_cadastro"))
Dim data_final_cadastro : data_final_cadastro = Cstr(""&Request.QueryString("data_final_cadastro"))

if data_inicial_cadastro <> "" and data_final_cadastro <> "" Then
  datas =  " data_lancamento BETWEEN "&sqlData(data_inicial_cadastro&" 00:00:00")&" AND "&sqlData(data_final_cadastro&" 23:23:59 ")
  datas2 =  " data_cadastro BETWEEN "&sqlData(data_inicial_cadastro&" 00:00:00")&" AND "&sqlData(data_final_cadastro&" 23:23:59 ")
  datas3 =  " A.data_cadastro BETWEEN "&sqlData(data_inicial_cadastro&" 00:00:00")&" AND "&sqlData(data_final_cadastro&" 23:23:59 ")
end if

despacho_transportador = request.QueryString("despacho_transportador")
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
                <h3 align="center">Relat&oacute;rio de Sa&iacute;da de produtos produzidos p&oacute;s-venda - Entregues<br />
                </h3>
                <%
    AbrirConexao()

		Dim qtd_registros	, cont

   '' sql = "SELECT DISTINCT VW_TB_ORDENS_SERVICO.cod_pedido, NOME_FORMANDO,despacho_transportador, (SELECT GROUP_CONCAT(CONCAT(cast(IF(IFNULL(B.quantidade_entregue,0) > 0,B.quantidade_entregue,A.quantidade) AS CHAR), ' ' , A.NOME_PRODUTO) SEPARATOR ' + ') FROM VW_TB_ORDENS_SERVICO A INNER JOIN TB_PEDIDOS_ITENS_ENTREGUES B ON B.cod_ordem_servico = A.cod_ordem_servico WHERE A.NOME_FORMANDO = VW_TB_ORDENS_SERVICO.NOME_FORMANDO and B.data = TB_PEDIDOS_ITENS_ENTREGUES.data) as itens, TB_PEDIDOS_ITENS_ENTREGUES.data as data_entrega FROM VW_TB_ORDENS_SERVICO INNER JOIN TB_PEDIDOS_ITENS_ENTREGUES ON TB_PEDIDOS_ITENS_ENTREGUES.cod_ordem_servico = VW_TB_ORDENS_SERVICO.cod_ordem_servico  WHERE status_ordem_servico = 3 AND "&datas&" AND TB_PEDIDOS_ITENS_ENTREGUES.LACAMENTO_DE_CORRECAO is null AND despacho_transportador = "&sqlTexto(despacho_transportador)&" ORDER BY  NOME_FORMANDO"


    sql="SELECT ENT.LACAMENTO_DE_CORRECAO, OS.cod_pedido, ENT.data as data_entrega, ENT.data_lancamento, OS.despacho_transportador FROM stillos3.TB_PEDIDOS_ITENS_ENTREGUES AS ENT INNER JOIN TB_ORDENS_SERVICO AS OS ON OS.cod_ordem_servico = ENT.cod_ordem_servico INNER JOIN TB_PEDIDOS_FORMANDOS AS FS ON FS.COD_PEDIDO = OS.COD_PEDIDO WHERE "&datas&" AND OS.STATUS_ORDEM_SERVICO = 4 AND despacho_transportador = "&sqlTexto(despacho_transportador)&" AND ENT.LACAMENTO_DE_CORRECAO is null GROUP BY OS.cod_pedido, ENT.data, ENT.data_lancamento, OS.despacho_transportador"

    Dim cabecalho
		cont = 1
		Set rs=objConexao.Execute(sql,qtd_registros)
		If not rs.Eof Then
      cabecalho = true


      %>
        <table width="95%" cellspacing="0" border="0" align="center" 	cellpadding="5">
          <tr bgcolor="whitesmoke">
            <td width="5%"><div align="center"><strong>#</strong></div></td>
            <td width="5%"><div align="center"><strong>Pedido</strong></div></td>
            <td width="30%"><div align="left"><strong>Formando</strong><em></em> </div></td>
            <td width="50%"><div align="left"><strong>Itens</strong></div></td>
            <td width="8%"><div align="center"><strong>Data Entrega</strong></div></td>
          </tr>
          <% While not rs.Eof %>
          <tr>
            <td style="border-bottom:solid 1px #CCC"><div align="center"><%=cont%></div></td>
            <td style="border-bottom:solid 1px #CCC"><div align="center"><%=rs("cod_pedido")%></div></td>
            <td style="border-bottom:solid 1px #CCC"><div align="left">
            <%
              Set rs2 = objConexao.Execute("SELECT NOME_FORMANDO FROM TB_PEDIDOS_FORMANDOS PF INNER JOIN TB_FICHAS_IDENTIFICACAO IDT ON  IDT.cod_formando = PF.cod_formando WHERE cod_pedido = "&sqlNumero(rs("cod_pedido") ) )
              while not rs2.Eof
                response.write(rs2(0)&"<br />")
              rs2.MoveNext
              Wend
              rs2.Close
              set rs2 = nothing
            %>
            </div></td>
            <td style="border-bottom:solid 1px #CCC"><div align="left"><%
            set rs2 = objConexao.Execute("select OS.nome_produto, cast(IF(IFNULL(ENT.quantidade_entregue,0) > 0, ENT.quantidade_entregue,OS.quantidade) AS CHAR) as quantidade, OS.cod_pedido from TB_PEDIDOS_ITENS_ENTREGUES ENT inner join VW_TB_ORDENS_SERVICO OS on OS.cod_ordem_servico = ENT.cod_ordem_servico where OS.cod_pedido =  "&sqlNumero(rs("cod_pedido")))
            while not rs2.Eof
              response.write rs2("quantidade") & " " &rs2("nome_produto") & ", "
            rs2.MoveNext
            Wend
            rs2.Close
            set rs2 = nothing


            set rs2 = objConexao.Execute("SELECT  * from TB_PEDIDOS_ITENS_ENTREGUES_MANUAL WHERE cod_ordem_servico = "&sqlNumero(rs("cod_pedido")) &" AND data_entregue ="& sqlData(rs("data_entrega")))
            if not rs2.Eof Then
              cods_pedido = cods_pedido & rs("cod_pedido") & ","
              response.write(" Entrada Manual: ")
              while not rs2.Eof
                response.write(rs2("descricao")&", ")
              rs2.MoveNext()
              Wend
            end if
            rs2.Close
            set rs2 = nothing
            %>
            </div></td>
            <td style="border-bottom:solid 1px #CCC"><div align="center"><%=ArrumaData(rs("data_entrega"))%></div></td>
          </tr>
        <%
        cont = int(cont) + 1
				rs.MoveNext()
				Wend
    End If

    rs.Close()
    Set rs = Nothing
    cods_pedido = cods_pedido & "0,"
    'SELECIONAMOS TODOS OS PEDIDOS ENTREGUES MANUALMENTE
    sql = "SELECT TB_PEDIDOS_ITENS_ENTREGUES_MANUAL.id, TB_PEDIDOS_ITENS_ENTREGUES_MANUAL.cod_ordem_servico, TB_PEDIDOS_ITENS_ENTREGUES_MANUAL.data_entregue,TB_PEDIDOS_ITENS_ENTREGUES_MANUAL.data_cadastro, TB_PEDIDOS_ITENS_ENTREGUES_MANUAL.descricao,  group_concat(trim(VW_PEDIDOS_FORMANDOS.NOME_FORMANDO)) as NOME_FORMANDO FROM TB_PEDIDOS_ITENS_ENTREGUES_MANUAL JOIN VW_PEDIDOS_FORMANDOS on TB_PEDIDOS_ITENS_ENTREGUES_MANUAL.cod_ordem_servico = VW_PEDIDOS_FORMANDOS.cod_pedido where TB_PEDIDOS_ITENS_ENTREGUES_MANUAL.cod_ordem_servico not in ( "&left(cods_pedido,len(cods_pedido) - 1)&" ) and "&datas2&" GROUP BY TB_PEDIDOS_ITENS_ENTREGUES_MANUAL.id, TB_PEDIDOS_ITENS_ENTREGUES_MANUAL.cod_ordem_servico,  TB_PEDIDOS_ITENS_ENTREGUES_MANUAL.data_entregue, TB_PEDIDOS_ITENS_ENTREGUES_MANUAL.data_cadastro, TB_PEDIDOS_ITENS_ENTREGUES_MANUAL.descricao ORDER BY NOME_FORMANDO;"

    Set rs=objConexao.Execute(sql,qtd_registros2)
    If not rs.Eof Then
      if not cabecalho then
      %>
        <table width="95%" cellspacing="0" border="0" align="center"  cellpadding="5">
          <tr bgcolor="whitesmoke">
            <td width="5%"><div align="center"><strong>#</strong></div></td>
            <td width="5%"><div align="center"><strong>Pedido</strong></div></td>
            <td width="30%"><div align="left"><strong>Formando</strong><em></em> </div></td>
            <td width="50%"><div align="left"><strong>Itens</strong></div></td>
            <td width="8%"><div align="center"><strong>Data Entrega</strong></div></td>
          </tr>
      <%
      end if
      while not rs.Eof
      %>
<tr>
<td style="border-bottom:solid 1px #CCC"><div align="center"><%=cont%></div></td>
<td style="border-bottom:solid 1px #CCC"><div align="center"><%=rs("cod_ordem_servico")%></div></td>
<td style="border-bottom:solid 1px #CCC"><div align="left"><%=rs("nome_formando")%></div></td>
<td style="border-bottom:solid 1px #CCC"><div align="left"><%=rs("descricao")%></div></td>
<td style="border-bottom:solid 1px #CCC"><div align="center"><%=ArrumaData(rs("data_entregue"))%></div></td>
</tr>
        <%
      cont = int(cont) + 1
      rs.MoveNext
      Wend
    end if
    rs.Close()
    Set rs = Nothing


    sql = "SELECT A.id, A.tia, A.data_entregue, A.data_cadastro, A.descricao , B.nome  from TB_PEDIDOS_ITENS_ENTREGUES_MANUAL_MACK  A INNER JOIN TB_MACK_FORMANDOS B on A.tia = B.tia where "&datas3&" group by  A.id, A.tia, A.data_entregue, A.data_cadastro, A.descricao , B.nome  "

    Set rs=objConexao.Execute(sql,qtd_registros2)
    If not rs.Eof Then
      if not cabecalho then
      %>
        <table width="95%" cellspacing="0" border="0" align="center"  cellpadding="5">
          <tr bgcolor="whitesmoke">
            <td width="5%"><div align="center"><strong>#</strong></div></td>
            <td width="5%"><div align="center"><strong>TIA</strong></div></td>
            <td width="30%"><div align="left"><strong>Formando</strong><em></em> </div></td>
            <td width="50%"><div align="left"><strong>Itens</strong></div></td>
            <td width="8%"><div align="center"><strong>Data Entrega</strong></div></td>
          </tr>
      <%
      end if
      while not rs.Eof
      %>
<tr>
<td style="border-bottom:solid 1px #CCC"><div align="center"><%=cont%></div></td>
<td style="border-bottom:solid 1px #CCC"><div align="center"><%=rs("tia")%></div></td>
<td style="border-bottom:solid 1px #CCC"><div align="left"><%=rs("nome")%></div></td>
<td style="border-bottom:solid 1px #CCC"><div align="left"><%=rs("descricao")%></div></td>
<td style="border-bottom:solid 1px #CCC"><div align="center"><%=ArrumaData(rs("data_entregue"))%><strong></strong></div></td>
</tr>
        <%
      cont = int(cont) + 1
      rs.MoveNext
      Wend
    end if
    rs.Close()
    Set rs = Nothing


    FecharConexao()
	  %>
    <tr>
      <td colspan="8" style="border-bottom:solid 1px #CCC; font-size:22px;"><div align="right"><strong>Total &nbsp;&nbsp;&nbsp;<%=cont-1%></strong></div></td>
    </tr>
    </table>
    <br />
    <br />
    Impresso em: <%=Now()%> <br>
    Lançamentos do dia <%=ArrumaData(data_inicial)%> ao <%=ArrumaData(data_final)%>
		<%

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
