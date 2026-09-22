
<%
Dim calendario_necessario, BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
calendario_necessario = "N�O"
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 1
BTN_IMPRIMIR_LINK = "#"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%
Dim comando, TXT_DEVOLUCAO, ERROR_LINHA, pronto, CONT_ITENS
pronto = Request.QueryString("pronto")
ERROR_LINHA = 1
comando = Request.QueryString("comando")
TXT_DEVOLUCAO = Request.Form("TXT_DEVOLUCAO")
TXT_DEVOLUCAO = Replace(TXT_DEVOLUCAO," ","")
CONT_ITENS = 0

%>
<h2>CAP - Pagamento de Consultores - Comissionamento de Coordenadores</h2>

<div class="msgGrandeSucesso" style="display:none" id="sucesso_msg"></div>
<div class="msgGrandeErro" style="display:none"  id="erro_msg"></div>
<div class="msg" align="center">
  <p><%=Request.QueryString("msg")%></p>
</div>

<div class="subMenu">
  Sub-Menu: <a href="cap_fechamentos_comissionamento_coordenadores_formulario.asp">Novo Comissionamento de Coordenadores</a>
</div>

<h3 align="center">Rela&ccedil;&atilde;o de Comissionamentos de Coordenadores</h3>

<table width="99%" border="1" align="center" cellpadding="5" style="border-collapse:collapse; text-align: center;">
  <tr>
    <th>Data In&iacute;cio da vigencia da comiss&atilde;o</th>
    <th>Data Final da vig&ecirc;ncia da comiss&atilde;o</th>
    <th>Percentual de Comiss&atilde;o sobre as vendas</th>
    <th></th>
  </tr>
  <%
    abrirConexao()
    set rs = objConexao.execute("SELECT ano_inicio, ano_fim, percentual, COD_COMISSIONAMENTO_COORDENADORES FROM stillos3.TB_FECHAMENTO_COMISSIONAMENTO_COORDENADORES")
    while not rs.Eof 
    %>
      <tr>
        <td><%=rs("ano_inicio").value%></td>
        <td><%=rs("ano_fim").value%></td>
        <td><%=formatNumber(rs("percentual").value, 2)%></td>
        <td>
          <a href="cap_fechamentos_comissionamento_coordenadores_formulario.asp?comando=excluir&COD_COMISSIONAMENTO_COORDENADORES=<%=rs("COD_COMISSIONAMENTO_COORDENADORES").value%>">
            <img src="images/icons/icon_excluir.png" width="16" height="16" border="0">
          </a>
        </td>
      </tr>
    <%

    rs.movenext()
    wend
    rs.close()
    set rs = nothing
    FecharConexao() 
  %>
</table>
<!-- #include file="files/includes/miolo-inferior.asp" -->