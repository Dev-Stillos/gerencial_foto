<!-- epv_renovacao_album_imprimir.asp -->
<!-- #include file="files/includes/miolo-superior.asp" -->
<%
AbrirConexao()
Set rs = objConexao.Execute("SELECT R.COD_RENOVACAO_TEMP, R.COD_FORMANDO, R.TIPO, R.PERIODO_RENOVACAO, R.DATA_CRIACAO, F.FD_NOME_004 
                            FROM TB_FICHAS_IDENTIFICACAO_RENOVACAO_TEMPORARIA R
                            LEFT JOIN TB_FORMANDOS F ON F.FD_CODIG_004 = R.COD_FORMANDO
                            WHERE R.PODE_IMPRIMIR = 1")
%>

<h2>Registros Prontos para Impressão</h2>

<table border="1" width="100%">
  <tr>
    <th>Código</th>
    <th>Nome do Formando</th>
    <th>Tipo</th>
    <th>Período de Renovação (dias)</th>
    <th>Data Criação</th>
  </tr>
<% Do While Not rs.EOF %>
  <tr>
    <td><%= rs("COD_RENOVACAO_TEMP") %></td>
    <td><%= rs("FD_NOME_004") %></td>
    <td><%= rs("TIPO") %></td>
    <td><%= rs("PERIODO_RENOVACAO") %></td>
    <td><%= rs("DATA_CRIACAO") %></td>
  </tr>
<% rs.MoveNext : Loop
rs.Close : Set rs = Nothing
FecharConexao()
%>
</table>

<br>
<div style="text-align:center;">
  <button onclick="window.print();">Imprimir Página</button>
  <a href="epv_renovacao_album.asp">Voltar</a>
</div>

<!-- #include file="files/includes/miolo-inferior.asp" -->
