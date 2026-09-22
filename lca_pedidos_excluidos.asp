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
<style>

.menu_paginacaocao a, .menu_paginacaocao span {
  position:relative;
  display: block;
  float:left;
}
.tabela_listagem th{
  text-align: center;
  font-weight: bold;
}
.tabela_listagem td:last-child {
  text-align: center;
}
h2 {
  margin-bottom: 15px !important;
}
</style>
<h2>LCA - Pedidos Excluídos</h2>

<%

dim pagina, limit, offset
dim rs, sql

limit = 50
pagina = int(request.queryString("pagina"))

if pagina = empty or pagina = 0 then 
  pagina = 1
end if
offset = (pagina - 1) * limit

abrirConexao()
sql = "SELECT a.*, b.nome_vendedor, c.nome FROM LOG_TB_PEDIDOS AS a INNER JOIN TB_VENDEDORES AS b ON a.cod_vendedor = b.cod_vendedor INNER JOIN TB_ADM_USUARIOS AS c on a.cod_usuario_excluiu = c.cod_usuario ORDER BY COD_PEDIDO DESC LIMIT " &limit& " OFFSET " &offset
set rs = objConexao.execute(sql)
if not rs.Eof then
  %><table width="95%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse" class='tabela_listagem'>
    <tr bgcolor="whitesmoke">
      <th>Cód. Pedido</th>
      <th>Representante</th>
      <th>Data excluído</th>
      <th>Excluído Por</th>
      <!--<th> </th>-->
    </tr><%

  while not rs.Eof
    %><tr>
      <td><%=rs("cod_pedido").value%></td>
      <td><%=rs("nome_vendedor").value%></td>
      <td><%=rs("data_excluido").value%></td>
      <td><%=rs("nome").value%></td>
      <!--<td>
        <a href='lca_pedidos_excluidos_formulario.asp?cod_pedido=<%=rs("cod_pedido").value%>'>
          <img src="images/icons/icon_visualizar.png" style="cursor: pointer;" alt="Visualizar registro">
        </a>
      </td>-->
    </tr><%
  rs.movenext()
  wend 
  %></table><%
else
  %> <h3>Não há pedidos Excluídos</h3> <%
end if
rs.close()
set rs = nothing


sql = "SELECT count(*) as total FROM LOG_TB_PEDIDOS" 
set rs = objConexao.execute(sql)
if not rs.Eof then
  total = int(rs("total").value&"")
end if 
rs.close
set rs = nothing
qtd_paginas = int(total / limit + .5)

i = 1
%>
<div class='menu_paginacaocao'>
<%
while i <= qtd_paginas 
  if i = pagina then
   %><span><%=i%></span><%
  else
  %><a class="left" href='?pagina=<%=i%>'><%=i%></a><%
  end if
i = i + 1
wend

%>
</div>
<%

fecharConexao()
%>
<!-- #include file="files/includes/miolo-inferior.asp" -->