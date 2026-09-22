<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim COD_PEDIDO, COD_PEDIDO_CHEQUE, cod_pedido_cheque_historico
Dim historico
Dim comando

comando = Request.QueryString("comando")

if comando = "" then
  comando = Request.form("comando")
  if comando = "" then
    comando = "incluir"
  end if
end if
if comando = "visualizar" then
  comando = "incluir"
end if

campo1 = Request.form("campo1")

COD_PEDIDO  = Request.QueryString("cod_pedido")
COD_PEDIDO_CHEQUE   = Request.QueryString("cod_pedido_cheque")



'SELECIONAMOS A FUNÇÃO A SER EXECUTADA'

if comando = "incluir" AND campo1  = "" then
  'nada
elseif comando = "incluir" AND campo1 = "1" then 
  incluir()
elseif comando = "excluir" AND campo1 = "" then
  cod_pedido_cheque_historico = Request.QueryString("cod_pedido_cheque_historico")
  lerregistro()
elseif comando = "excluir" AND campo1 = "1" then
  deletar()
end if

sub lerformulario ()
  historico = Request.form("historico")
  COD_PEDIDO = Request.form("COD_PEDIDO")
  COD_PEDIDO_CHEQUE = Request.form("COD_PEDIDO_CHEQUE")
  cod_pedido_cheque_historico = Request.form("cod_pedido_cheque_historico")
end sub

sub lerregistro()
  AbrirConexao()

  set rs = objConexao.Execute("SELECT historico,COD_PEDIDO_CHEQUE FROM TB_PEDIDOS_CHEQUE_HISTORICO WHERE cod_pedido_cheque_historico = "&sqlNumero(cod_pedido_cheque_historico))
  if not rs.Eof then
    historico = rs("historico")
    COD_PEDIDO_CHEQUE = rs("COD_PEDIDO_CHEQUE")  
  end if
  rs.close
  set rs = Nothing

  set rs = objConexao.Execute("SELECT COD_PEDIDO FROM TB_PEDIDOS_CHEQUE WHERE cod_pedido_cheque = "&sqlNumero(cod_pedido_cheque))
  if not rs.Eof then
    COD_PEDIDO = rs("COD_PEDIDO")
  end if
  rs.close
  set rs = Nothing

  FecharConexao()

end sub

sub incluir()
  lerformulario()

  AbrirConexao()

    sql = "INSERT INTO TB_PEDIDOS_CHEQUE_HISTORICO ( cod_pedido_cheque, historico, cod_usuario) VALUES ("&sqlNumero(cod_pedido_cheque)&", "&sqlTexto(historico)&", "&sqlNumero(session("cod_usuario"))&" )"

  objConexao.Execute(sql) 
  FecharConexao()
  Response.redirect("lca_pedidos_pendentes_alteracoes_part_3_POPUP_cheque_historico.asp?COD_PEDIDO="&COD_PEDIDO&"&cod_pedido_cheque="&cod_pedido_cheque)
end sub

sub deletar()
  lerformulario()

  sql = "DELETE FROM TB_PEDIDOS_CHEQUE_HISTORICO WHERE cod_pedido_cheque_historico = "&sqlNumero(cod_pedido_cheque_historico)

  AbrirConexao()
    'checar validação 
    objConexao.Execute(sql)
  FecharConexao()
  Response.redirect("lca_pedidos_pendentes_alteracoes_part_3_POPUP_cheque_historico.asp?COD_PEDIDO="&COD_PEDIDO&"&cod_pedido_cheque="&cod_pedido_cheque)
end sub



%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <title>Cadastro de Produto em pedido</title>
    <link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
    <style type="text/css">
    <!--
      body {background-color: #FFFFFF;}
    -->
    </style>
    <script language="vbscript">
      suporteVBscript=1
      function VBconfirm(mensagem)

      VBconfirm=msgbox(mensagem,vbyesno)
      end function
    </script>
  </head>
<body>
<%if comando <> "visualizar" then%>
  <a href="javascript:parent.history.go(-1)"><img alt="Voltar" class="botoes_icones" src="images/icons/icon_voltar.jpg" width="32" height="28" border="0"></a>
<%end if%>
  <% if comando = "excluir" then%>
    <div class="msgGrandeErro">
      <p>Você tem certeza que deseja excluir este registro?</p>
    </div>
  <%end if%>
    <form name="FORM_CHEQUE" action="lca_pedidos_pendentes_alteracoes_part_3_POPUP_cheque_historico.asp" method="post">
      <table width="100%" border="0" align="center" cellpadding="2" cellspacing="2">
        <tr>
          <td colspan="2"><div align="center"><strong>Histórico</strong></div></td>
          <td colspan="2"><div align="center">&nbsp</div></td>
        </tr>
        <tr>
          <td colspan="2"><div align="center">
            <textarea rows='4' name='historico' class='inputTamanho3' maxlength="500"><%=historico%></textarea>
          </td>
          <td colspan="2"><div align="center">
            <a href="javascript:document.FORM_CHEQUE.submit()">
              <%if comando = "excluir" Then%>
                <img src="images/layout/btnExcluir.jpg" />
              <%else%>
                <img src="images/layout/btnGravar.jpg" />
              <%end if%>
            </a>
          </div></td>
        </tr>
      </table>
    <input type='hidden' name='campo1' value="1" />  
    <input type='hidden' name='comando' value="<%=comando%>" />
    <input type='hidden' name='cod_pedido' value="<%=cod_pedido%>" />
    <input type='hidden' name='cod_pedido_cheque' value="<%=cod_pedido_cheque%>" />
    <input type='hidden' name='cod_pedido_cheque_historico' value="<%=cod_pedido_cheque_historico%>" />
    </form>
    <%
    AbrirConexao()
      Dim i : i = qtd
      Set rs = objConexao.Execute("SELECT * FROM TB_PEDIDOS_CHEQUE_HISTORICO WHERE cod_pedido_cheque = "&sqlNumero(cod_pedido_cheque)&" ORDER BY data_cadastro DESC", qtd)
      i = qtd + 1
      if not rs.Eof then
        %>
        <table width="100%" border="1" align="center" cellpadding="2" cellspacing="2" style="border-collapse:collapse">
          <tr>
            <th colspan="2"><div align="center">#</div></th>
            <th colspan="2"><div align="center">Histórico</div></th>
            <th colspan="2"><div align="center">Usuário</div></th>
            <th colspan="2"><div align="center">Data cadastro</div></th>
            <% if validaUsuario(2,"",80,true)  then %>
            <th colspan="2"><div align="center">&nbsp</div></th>
            <%end if%>
          </tr>
        <%
        while not rs.Eof
        i = i - 1
          %>
          <tr>
            <td colspan="2"><div align="center"><strong><%=i%></strong></div></td>
            <td colspan="2"><div align="center"><pre><%=rs("historico")%></pre></div></td>
            <td colspan="2"><div align="center"><%=nomeUsuario(rs("cod_usuario"),true)%></div></td>
            <td colspan="2"><div align="center"><%=arrumaData(rs("data_cadastro"))%></div></td>
            <% if validaUsuario(2,"",80,true)  then %>
            <td colspan="2"><div align="center">
              <a href="lca_pedidos_pendentes_alteracoes_part_3_POPUP_cheque_historico.asp?comando=excluir&cod_pedido_cheque_historico=<%=rs(0)%>">
                <img alt="Excluir" src="images/icons/icon_excluir.png" border="0">
              </a>
            </div></td>
            <% end if %>
          </tr>
          <%
        rs.movenext()
        Wend
      
        %>
        </table>
        <%
      else
        Response.write("<br><br><center><strong>Nenhum Histórico encontrado</strong></center>")
      end if
      rs.close
      set rs = Nothing
      FecharConexao()
    %>
  </body>
</html>