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


<%
dim rs , cod_pedido

abrirConexao()

cod_pedido = request.queryString("cod_pedido")
set rs = objConexao("SELECT a.*, b.nome_vendedor FROM LOG_TB_PEDIDOS AS a INNER JOIN TB_VENDEDORES AS b ON a.cod_vendedor = b.cod_vendedor WHERE cod_pedido = "&sqlNumero(cod_pedido))
if not rs.eof then
  
end if 
rs.close()
set rs =nothing

fecharConexao()
%>

<style>
h2 { margin-bottom: 15px !important; }
</style>
<h2>LCA - Pedidos Excluídos</h2>





</div>
<%
fecharConexao()
%>
<!-- #include file="files/includes/miolo-inferior.asp" -->