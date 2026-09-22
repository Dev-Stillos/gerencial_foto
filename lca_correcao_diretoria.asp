<%
Dim calendario_necessario, BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK, script_valida_form_master
calendario_necessario = "NAO"
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 1
BTN_IMPRIMIR_LINK = "#"

script_valida_form_master = "NAO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->

<h2>LCA - Corre&ccedil;&atilde;o Diretoria </h2>

  <div class="subMenu">
    Sub-Menu: 
    <a href="lca_correcao_diretoria_lista_erros.asp">
      Verificar fichas marcadas com erros
    </a> | 
    <a href="adm_relatorios_fichas_autenticadas_diretoria.asp?filtro1=+&DATA_INI=01%2F04%2F2009&DATA_FIM=30%2F04%2F2009">
      &Uacute;ltimas fichas baixadas
    </a> | 
    <a href="lca_pedidos_excluidos.asp">
      Pedidos Exclu&iacute;dos
    </a>
  </div>

  <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
  <%
    Dim rs, rs2

    AbrirConexao()
      Set rs = objConexao.Execute("SELECT * FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE cod_fechamento_periodo in (SELECT cod_fechamento_periodo FROM TB_PEDIDOS WHERE finaliza_correcao = 1 and finaliza_diretoria = 2)  order by data_termino desc", qtdRegistros)

      If rs.Eof then %>
        <p align="center">N&atilde;o h&aacute; fichas pendentes para verifica&ccedil;&atilde;o!</p>
      <% else %>
        <h3 align="center">Pagamentos abaixo com fichas ainda n&atilde;o conferidas (forma nova): </h3>
      <%
      Set rs2=objConexao.Execute("SELECT COUNT(*) as qtd FROM TB_PEDIDOS WHERE FINALIZA_CORRECAO = 1 AND DATA_DESCEU_DIRETORIA is null")
      If not rs2.Eof Then
        temp = formatNumber(0&rs2("qtd"),0)
        If int(temp) > 0 Then
          %>
          <div class="msgGrande">
            H&Aacute; <%=rs2(0)%> pedido(s) que foi(foram) conferido(s), mas que n&atilde;o desceram at&eacute; o momento para confer&ecirc;ncia da diretoria
          </div>
          <%
        End If
      End If
      rs2.Close	
      Set rs2=Nothing
  %>

  <table width="95%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
    <tr bgcolor="whitesmoke">
      <td><div align="center"><strong>Per&iacute;odo de Vendas</strong></div></td>
      <td width="19%"><div align="center"><strong>Data limite de Lan&ccedil;amento no sistema</strong></div></td>
      <td width="10%"><div align="center"><strong>Data que ser&aacute; ou foi o pagamento</strong></div></td>
      <td width="10%"><div align="center"><strong>Vendas que n&atilde;o chegaram na diretoria</strong></div></td>
      <td width="21%"><div align="center"><strong>Autoriza&ccedil;&atilde;o para finalizar fechamento sem autenticar todos os pedidos</strong></div></td>
      <td width="7%">&nbsp;</td>
    </tr>
		<% While not rs.Eof %>
      <tr>
        <td><div align="center" class="textoMedio"><%=ArrumaData(rs("data_inicio"))%> &eacute; <%=ArrumaData(rs("data_termino"))%></div></td>
        <td><div align="center"><%=ArrumaData(rs("data_termino_lancamento_vendas"))%></div></td>
        <td><div align="center"><%=ArrumaData(rs("data_pagamento"))%></div></td>
        <td><div align="center">
          <%
          Set rs2=objConexao.Execute("SELECT COUNT(*) as qtd FROM TB_PEDIDOS WHERE FINALIZA_CORRECAO = 1 AND DATA_DESCEU_DIRETORIA is null AND cod_fechamento_periodo = "&rs("cod_fechamento_periodo"))
          If not rs2.Eof Then
            Response.Write(rs2(0))
          End If
          rs2.Close
          Set rs2=Nothing
          %>
        </div></td>
        <td><div align="center">
          <%
          If rs("autorizacao_diretoria_para_finalizar") = "S" Then
            Response.Write("Sim")
          Else
            Response.Write("N&atilde;o")
          End if
          %>
        </div></td>
        <td>
          <div align="center">
            <a href="lca_correcao_diretoria_mensal.asp?cod_fechamento_periodo=<%=rs("cod_fechamento_periodo")%>">
              <strong>Listar Pedidos</strong>
            </a>
          </div>
        </td>
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
			
  Set rs2=objConexao.Execute("SELECT COD_PEDIDO FROM TB_PEDIDOS WHERE FINALIZA_CORRECAO = 1 AND FINALIZA_DIRETORIA = 2 AND cod_fechamento_periodo = 0")

	If not rs2.Eof Then
    Response.Write("Confira os pedido(s) da forma antiga de pagamento que precisam ser corrigidos pelo administrador do sistema para que sejam visualizadas para corre��o:<br />")

    While not rs2.Eof
      Response.Write("<a href='https://stillos.com.br/gerencial_foto/lca_busca_detalhamento_pedido.asp?cod_pedido="&rs2("cod_pedido")&"' target='_blank'>"&rs2("cod_pedido")&"</a><br />")
    rs2.MoveNext()
    Wend

    Response.Write("Para baixar estes pedidos ser� necess�rio enviar e-mail para josehenriquepl@gmail.com informando quais os pedidos que ser�o autenticados.")
  End If
  
  rs2.Close	
  Set rs2=Nothing
  
  FecharConexao()
  %>
  <p align="center">
    <input class="btnTopo" onclick="abre('#')" type="button" name="btnTopo2" value=" " />
  </p>		

<!-- #include file="files/includes/miolo-inferior.asp" -->