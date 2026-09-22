<meta http-equiv="cache-control" content="no-cache" />
<!-- #include file="../../includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="../../includes/asp/valida_usuario.asp" -->
<!-- #include file="functions.asp" -->
<% Response.Charset="ISO-8859-1" %>
<div id="listas" style="margin-bottom:50px;">
<%

    AbrirConexao()
    dim resultSet
    dim sql
    dim rows
    
%>

	<h1 class="titulo_pagina" > Projetos futuros </h1>
	<table id="lista_itens" class="tabela projetos_futuros">
	<tr>
		 <th> Nome </th>
		 <th> Descrição </th>    
	</tr>

		<%
			 
			sql = "SELECT nome,descricao FROM tb_ple_projetos WHERE status = 0"
			set resultSet = objConexao.Execute(sql)
			while not resultSet.eof 
			    nome = resultSet("nome")
			    descricao = resultSet("descricao")          

				        
		%>
					 <tr>
						  <td>  <%=nome%> </td>
						  <td>  <%=descricao%> </td>           
					 </tr>  
		<%
			 resultSet.MoveNext()
			wend
		%>
	</table>

	<table  class="tabela tabela_legenda" style="display:none;">
		<tr>
			<th colspan=2> Legenda </th>
		</tr>
		<tr>
			<td> (C)</td>
			<td> Concluído </td>	
		</tr>
		<tr>
			<td> (I)</td>
			<td> Incompleto </td>	
		</tr>	
	</table>

</div>
