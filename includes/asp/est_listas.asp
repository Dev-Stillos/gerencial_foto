<meta http-equiv="cache-control" content="no-cache" />
<!-- #include file="../../includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="../../includes/asp/valida_usuario.asp" -->
<!-- #include file="functions.asp" -->
<% Response.Charset="ISO-8859-1" %>
<div id="listas">
<%

    AbrirConexao()
    dim resultSet
    dim sql
    dim rows
    
if(request.queryString("tipo") = "itens") then%>
<h1 class="titulo_pagina" > Lista de Itens </h1>
<table id="lista_itens" class="tabela">
<tr>
    <th> Nome </th>
    <th> Categoria </th>
    <th> Tipo </th>
    <th> Descri&ccedil&atildeo </th>
    <th colspan="3" class="botoes_tabela"> </th>
</tr>

<%
    
            sql = "SELECT A.cod_item, A.nome as nome_item, A.descricao, A.tipo, B.nome as nome_categoria FROM tb_est_itens A INNER JOIN tb_est_categoria B on A.cod_categoria = B.cod_categoria"
            set resultSet = objConexao.Execute(sql)
            while not resultSet.eof 
                cod_item = resultSet("cod_item")
                nome = resultSet("nome_item")
                categoria = resultSet("nome_categoria")
                tipo = resultSet("tipo")
                descricao = resultSet("descricao")
                
                if tipo = "C" then
                    tipo = "Composto"
                else
                    tipo = "Simples"
                end if
                

            %>
                <tr>
                    <td>  <%=nome%> </td>
                    <td>  <%=categoria%> </td>
                    <td>  <%=tipo%> </td>   
                    <td>  <%=descricao%> </td>  

                    <td class="botao_tabela" onclick="visualizarItem(<%=cod_item%>)"><img src="images/icons/icon_visualizar.png" alt="Visualizar Registro" title="Visualizar Registro" border="0"></td>
                    <td class="botao_tabela" onclick="alterarItem(<%=cod_item%>)"><img src="images/icons/icon_alterar.png" alt="Alterar Registro" title="Alterar Registro" width="16" height="16" border="0"></td>
                    <td class="botao_tabela" onclick="excluirItem(<%=cod_item%>)"><img src="images/icons/icon_excluir.png" alt="Excluir Registro" title="Excluir Registro" border="0"></td>
                </tr>  
            <%
                resultSet.MoveNext()
             wend
            %>
</table>

<% end if%>

<%if(request.queryString("tipo") = "categorias") then%>

    <h1 class="titulo_pagina" > Lista de Categorias </h1>
    <table id="lista_categorias" class="tabela">
    <tr>
        <th> Nome </th>
        <th> Descri&ccedil&atildeo </th>        
        <th colspan="3" class="botoes_tabela"> </th>
    </tr>


<%
            sql = "SELECT cod_categoria, nome, descricao FROM tb_est_categoria"
            set resultSet = objConexao.Execute(sql)
            while not resultSet.eof 
                cod_categoria = resultSet("cod_categoria")
                nome = resultSet("nome")
                descricao = resultSet("descricao")
            %>
                <tr>
                    <td>  <%=nome%> </td>
                    <td>  <%=descricao%> </td>                    

                    <td class="botao_tabela" onclick="visualizarCategoria(<%=cod_categoria%>)"><img src="images/icons/icon_visualizar.png" alt="Visualizar Registro" title="Visualizar Registro" border="0"></td>
                    <td class="botao_tabela" onclick="alterarCategoria(<%=cod_categoria%>)"><img src="images/icons/icon_alterar.png" alt="Alterar Registro" title="Alterar Registro" width="16" height="16" border="0"></td>
                    <td class="botao_tabela" onclick="excluirCategoria(<%=cod_categoria%>)"><img src="images/icons/icon_excluir.png" alt="Excluir Registro" title="Excluir Registro" border="0"></td>
                </tr>  
            <%
                resultSet.MoveNext()
             wend
            %>
</table>

<% end if%>


<%if(request.queryString("tipo") = "fornecedores") then%>

    <h1 class="titulo_pagina" > Lista de fornecedores </h1>
    <table id="lista_categorias" class="tabela">
    <tr>
        <th> Nome </th>
        <th> Descri&ccedil&atildeo </th>        
        <th colspan="3" class="botoes_tabela"> </th>
    </tr>


<%
            sql = "SELECT cod_fornecedor, nome, descricao FROM tb_est_fornecedores"
            set resultSet = objConexao.Execute(sql)
            while not resultSet.eof 
                cod_fornecedor = resultSet("cod_fornecedor")
                nome = resultSet("nome")
                descricao = resultSet("descricao")
            %>
                <tr>
                    <td>  <%=nome%> </td>
                    <td>  <%=descricao%> </td>                    

                    <td class="botao_tabela" onclick="visualizarFornecedor(<%=cod_fornecedor%>)"><img src="images/icons/icon_visualizar.png" alt="Visualizar Registro" title="Visualizar Registro" border="0"></td>
                    <td class="botao_tabela" onclick="alterarFornecedor(<%=cod_fornecedor%>)"><img src="images/icons/icon_alterar.png" alt="Alterar Registro" title="Alterar Registro" width="16" height="16" border="0"></td>
                    <td class="botao_tabela" onclick="excluirFornecedor(<%=cod_fornecedor%>)"><img src="images/icons/icon_excluir.png" alt="Excluir Registro" title="Excluir Registro" border="0"></td>
                </tr>  
            <%
                resultSet.MoveNext()
             wend
            %>
</table>

<% end if%>


<%if(request.queryString("tipo") = "locais") then%>

    <h1 class="titulo_pagina" > Lista de Locais de Armazenamento </h1>
     Listar: <select name="select_listar">
            <%
               select_listar = ""
               select_listar = request.queryString("select_listar")
               todos_selected = ""
               locais_selected = ""
               posses_selected = ""
               
               if(select_listar = "*") then
                  todos_selected = "selected"
               elseif(select_listar = "L")then
                  locais_selected = "selected"
               elseif(select_listar = "P")then               
                  posses_selected = "selected"
               else
                  todos_selected = "selected"
                  select_listar = "*"
               end if
            
            %>
                      <option value="*" <%=todos_selected%>>Todos</option>
                      <option value="L" <%=locais_selected%>>Locais</option>
                      <option value="P" <%=posses_selected%>>Em posse (Vendedores)</option>
                </select>
        <br><br>
        
    <table id="lista_locais" class="tabela">
    <tr>
        <th> Nome </th>
        <th> Descri&ccedil&atildeo </th>        
        <th colspan="3" class="botoes_tabela"> </th>
    </tr>


<%

            if(select_listar = "*") then
             sql = "SELECT cod_local_armazenamento, nome, descricao FROM tb_est_local_armazenamento"
            elseif(select_listar = "L")then
             sql = "SELECT cod_local_armazenamento, nome, descricao FROM tb_est_local_armazenamento WHERE cod_local_armazenamento NOT IN (SELECT cod_local_armazenamento FROM tb_est_rel_vendedor_local)"
            else
             sql = "SELECT cod_local_armazenamento, nome, descricao FROM tb_est_local_armazenamento  WHERE cod_local_armazenamento IN (SELECT cod_local_armazenamento FROM tb_est_rel_vendedor_local)"
            end if

            set resultSet = objConexao.Execute(sql)
            while not resultSet.eof 
                cod_local_armazenamento = resultSet("cod_local_armazenamento")
                nome = resultSet("nome")
                descricao = resultSet("descricao")
            %>
                <tr>
                    <td>  <%=nome%> </td>
                    <td>  <%=descricao%> </td>                    

                    <td class="botao_tabela" onclick="visualizarLocal(<%=cod_local_armazenamento%>)"><img src="images/icons/icon_visualizar.png" alt="Visualizar Registro" title="Visualizar Registro" border="0"></td>
                    <td class="botao_tabela" onclick="alterarLocal(<%=cod_local_armazenamento%>)"><img src="images/icons/icon_alterar.png" alt="Alterar Registro" title="Alterar Registro" width="16" height="16" border="0"></td>
                    <td class="botao_tabela" onclick="excluirLocal(<%=cod_local_armazenamento%>)"><img src="images/icons/icon_excluir.png" alt="Excluir Registro" title="Excluir Registro" border="0"></td>
                </tr>  
            <%
                resultSet.MoveNext()
             wend
            %>
</table>

<% end if%>


<%if(request.queryString("tipo") = "rel_item") then%>

    <h1 class="titulo_pagina" > Lista de rela&ccedil&atildeo de itens - Produ&ccedil&atildeo X Estoque </h1>
    <table id="lista_locais" class="tabela">
    <tr>
        <th> Item Produ&ccedil&atildeo </th>
        <th> Item Estoque </th>
        <th> local de Armazenamento </th>   
        <th> Descri&ccedil&atildeo </th>        
        <th colspan="1" class="botoes_tabela"> </th>
    </tr>


<%
            sql = "SELECT A.cod_rel_item, B.nome AS nome_e, C.NOME_PRODUTO AS nome_p , A.descricao, D.nome as local FROM tb_est_rel_item A INNER JOIN tb_est_itens B ON A.cod_item_e = B.cod_item INNER JOIN TB_PRODUTOS C ON C.COD_PRODUTO = A.cod_item_p INNER JOIN tb_est_local_armazenamento D ON D.cod_local_armazenamento = A.cod_local;"
            set resultSet = objConexao.Execute(sql)
            while not resultSet.eof
                cod_rel_item = resultSet("cod_rel_item")
                nome_e = resultSet("nome_e")
                nome_p = resultSet("nome_p")
                local = resultSet("local")
                descricao = resultSet("descricao")
            %>
                <tr>
                    <td>  <%=nome_p%> </td>
                    <td>  <%=nome_e%> </td>
                    <td>  <%=local%> </td>                    
                    <td>  <%=descricao%> </td>                    

                    <td class="botao_tabela" onclick="excluirRelItem(<%=cod_rel_item%>)"><img src="images/icons/icon_excluir.png" alt="Excluir Registro" title="Excluir Registro" border="0"></td>
                </tr>  
            <%
                resultSet.MoveNext()
             wend
            %>
</table>

<% end if%>


<%if(request.queryString("tipo") = "rel_item_venda") then%>

    <h1 class="titulo_pagina" > Lista de rela&ccedil&atildeo de itens - Estoque X Venda</h1>
    <table id="lista_locais" class="tabela">
    <tr>
        <th> Item Venda </th>
        <th> Item Estoque </th>
        <th> local de Armazenamento </th>      
        <th> Descri&ccedil&atildeo </th>        
        <th colspan="1" class="botoes_tabela"> </th>
    </tr>


<%
            sql = "SELECT A.cod_rel_item_venda, B.nome AS nome_e, C.NOME_PRODUTO AS nome_p , A.descricao, D.nome as local  FROM tb_est_rel_item_venda A INNER JOIN tb_est_itens B ON A.cod_item_e = B.cod_item INNER JOIN TB_PRODUTOS C ON C.COD_PRODUTO = A.cod_item_p INNER JOIN tb_est_local_armazenamento D ON D.cod_local_armazenamento = A.cod_local;"
            set resultSet = objConexao.Execute(sql)
            while not resultSet.eof 
                cod_rel_item_venda = resultSet("cod_rel_item_venda")
                nome_e = resultSet("nome_e")
                nome_p = resultSet("nome_p") 
                local = resultSet("local")            

                descricao = resultSet("descricao")
            %>
                <tr>
                    <td>  <%=nome_p%> </td>
                    <td>  <%=nome_e%> </td>   
                    <td>  <%=local%> </td>                
                    <td>  <%=descricao%> </td>                    

                    <td class="botao_tabela" onclick="excluirRelItemVenda(<%=cod_rel_item_venda%>)"><img src="images/icons/icon_excluir.png" alt="Excluir Registro" title="Excluir Registro" border="0"></td>
                </tr>  
            <%
                resultSet.MoveNext()
             wend
            %>
</table>

<% end if%>


<%if(request.queryString("tipo") = "movimentacoes") then%>


    <%      
            if (isEmpty(request.queryString("mes")) or request.queryString("mes") = "undefined") then
                mes = month(now)
            else
                mes = request.queryString("mes")
            end if
            
            if (isEmpty(request.queryString("ano"))  or request.queryString("ano") = "undefined" ) then
               ano = year(now)
            else
                ano = request.queryString("ano")
            end if
            
    
    %>
    <h1 class="titulo_pagina" > Movimenta&ccedil&otildees do Estoque </h1>
        Ano: <select name="ano">
                <option value="*">Todos</option>
                <%  
                    selected = ""
                    for i=2014 to 2017 %>
                    <% 
                        selected = "" 
                        if isnumeric(ano) then
                            if clng(ano) = i then
                                selected = "selected" 
                            end if
                        end if
                    %>
                    <option <%=selected%> value="<%=i%>"><%=i%></option>
                <% next %>
            </select>
        M&ecircs: <select name="mes">
                <option value="*">Todos</option>
                <%  for i=1 to 12 %>
                    <% 
                        selected = "" 
                        if isnumeric(mes) then
                            if clng(mes) = i then
                                selected = "selected" 
                            end if
                        end if
                    %>
                    <option <%=selected%> value="<%=i%>"><%=i%></option>
                <% next %>
            </select>
        <br><br>
    <table id="lista_locais" class="tabela">
    <tr>
        <th class="coluna_cod" > Cod. </th>  
        <th class="coluna_cod" > Respons&aacutevel </th>  
        
        <th class="coluna_data" > Data </th>   
        <th class="coluna_tipo"> Tipo </th>
        <th class="coluna_local"> Em posse de </th>  
        <th class="coluna_descricao"> Descri&ccedil&atildeo </th>            
        <th> Itens </th>          

    </tr>


<%
            filtro = ""     
            if isNumeric(mes) then
                mes = " month(A.data) = "&mes
            else
                mes = ""
            end if
            
            
            if isNumeric(ano) then
                if len(mes) > 1 then
                   ano = " AND year(A.data) = "&ano
                else
                    ano = " year(A.data) = "&ano
                end if
            else
                ano = ""
            end if
            
            if len(ano) > 1 or len(mes) > 1 then
                filtro = " WHERE "&mes &"  "&ano
            else
                filtro = " "
            end if
     
            
            


            
            sql = "SELECT F.login as login, A.cod_movimentacao, A.tipo, A.descricao, B.nome, A.data, "&_
                  "CAST((SELECT GROUP_CONCAT(CONCAT(CONVERT(D.nome_produto USING utf8),', Qtd: ', C.quantidade) separator '<br>') FROM tb_est_itens_movimentacao C "&_
                  " INNER JOIN TB_PRODUTOS D on D.COD_PRODUTO = C.cod_item "&_
                  " WHERE A.cod_movimentacao = C.cod_movimentacao  ) AS CHAR) as itens  "&_
                  "FROM tb_est_movimentacao A "&_ 
                  "INNER JOIN tb_est_local_armazenamento B ON A.cod_local_armazenamento = B.cod_local_armazenamento "&_
                  "INNER JOIN TB_ADM_USUARIOS F on F.COD_USUARIO = A.cod_usuario "&_
                  ""&filtro&_
                  " ORDER BY A.cod_movimentacao DESC"

	
            set resultSet = objConexao.Execute(sql)
            cont = 0
            while not resultSet.eof 
		          cont = cont + 1
                cod_movimentacao = resultSet("cod_movimentacao")
                tipo = resultSet("tipo")
                local = resultSet("nome")
                usuario = resultSet("login")
                descricao = resultSet("descricao")
                itens = resultSet("itens")
                data = resultSet("Data")
                
                if tipo = "IN" then tipo = "Entrada" else if tipo = "OUT" then tipo = "Sa&iacuteda" else tipo = "Especial" end if
            %>
                <tr> 
                    <td class="coluna_cod" >  <%=cod_movimentacao%> </td>
                   <td class="coluna_cod" >  <%=usuario%> </td>                      
                    <td class="coluna_data" >  <%=data%> </td>  
                    <td class="coluna_tipo" >  <%=tipo%> </td>
                    <td class="coluna_local" >  <%=local%> </td>                    
                    <td class="coluna_descricao">  <%=descricao%> </td>         
                    <td> <textarea style="display:none"><%=itens%></textarea>
                    	<table>
		                 <%
		                 	if len(itens) > 0 then
		                 		itens_array = split(itens,"<br>")

		                 		
		                 		for each y in itens_array
		                 			   x = split(y,", Qtd:")
		                 			   
		                 				response.write("<tr>")
											response.write("<td>" & x(0) & "</td>")
	                 				
    										response.write("<td>" & x(1) & "</td>")
    										response.write("</tr>")
    										
 									next
	                 		end if
		                 %> 
		                 </table>
                    </td>  

                  
                </tr>  
            <%
                resultSet.MoveNext()
             wend
             
             if cont = 0 then
             	response.write("<tr> <td colspan=7> <h1> N&atildeo foram encontrados registros neste periodo </h1></td></tr>")
          	 end if
            %>
</table>

<table id="instrucoes" class="tabela">
	<tr> <th> Instru&ccedil&atildees </th> <tr>
	<tr>
		<td>
			<p>Nesta lista est&atildeo as Movimenta&ccedil&otildees dos produtos em estoque, podendo esta ser do tipo: "Entrada" e "Sa&iacuteda".<br>
			H&aacute tamb&eacutem um ter&ccedileiro tipo, este que ocorre quando feita a transfer&ecircncia entre locais de armazenamento. <br>
				Nesta situa&ccedil&atildeo ser&aacute lan&ccedilado duas movimenta&ccedil&otildees com os mesmos itens sendo uma do tipo "Sa&iacuteda", do local de origem, e a outra do tipo "Entrada", no local de destino.</p>
		</td>
	</tr>

</table>
<% end if%>




<%if(request.queryString("tipo") = "estoque") then%>

    <h1 class="titulo_pagina" > Itens do Estoque </h1>
         Listar: <select name="select_listar">
            <%               
               select_listar = ""
               select_listar = request.queryString("select_listar")
               todos_selected = ""
               locais_selected = ""
               posses_selected = ""
               
               if(select_listar = "*") then
                  todos_selected = "selected"
               elseif(select_listar = "L")then
                  locais_selected = "selected"
               elseif(select_listar = "P")then               
                  posses_selected = "selected"
               else
                  todos_selected = "selected"
                  select_listar = "*"
               end if
            
            %>
                      <option value="*" <%=todos_selected%>>Todos</option>
                      <option value="L" <%=locais_selected%>>Locais</option>
                      <option value="P" <%=posses_selected%>>Em posse (Vendedores)</option>
                </select>
        <br><br>
        
        
    <table id="lista_locais" class="tabela">
    <tr>
        <th> Item </th>
        <th> Local </th>           
        <th> Quantidade </th>          

    </tr>


<%



            
            sql = "call atualizarEstoque()"
            set resultSet = objConexao.Execute(sql)
            
            if ( select_listar = "*") then

            sql = "SELECT B.nome_produto as item, C.nome as local, A.quantia  "&_
                  "FROM tb_est_resumo A inner join TB_PRODUTOS B on A.cod_item = B.COD_PRODUTO "&_ 
                  "inner join tb_est_local_armazenamento C on C.cod_local_armazenamento = A.cod_local_armazenamento "&_
                  "  "&_
                  " ORDER BY B.nome_produto, C.nome"
                  
            elseif(select_listar = "L") then
                 sql = "SELECT B.nome_produto as item, C.nome as local, A.quantia  "&_
                  "FROM tb_est_resumo A inner join TB_PRODUTOS B on A.cod_item = B.COD_PRODUTO "&_ 
                  "inner join tb_est_local_armazenamento C on C.cod_local_armazenamento = A.cod_local_armazenamento "&_
                  " WHERE A.cod_local_armazenamento NOT IN (SELECT cod_local_armazenamento FROM tb_est_rel_vendedor_local)"&_
                  " ORDER BY B.nome_produto, C.nome"
                  
            else
                  sql = "SELECT B.nome_produto as item, C.nome as local, A.quantia  "&_
                  "FROM tb_est_resumo A inner join TB_PRODUTOS B on A.cod_item = B.COD_PRODUTO "&_ 
                  "inner join tb_est_local_armazenamento C on C.cod_local_armazenamento = A.cod_local_armazenamento "&_
                  " WHERE A.cod_local_armazenamento IN (SELECT cod_local_armazenamento FROM tb_est_rel_vendedor_local)"&_
                  " ORDER BY B.nome_produto, C.nome"
            end if
            
            
            set resultSet = objConexao.Execute(sql)
            old_local = ""
            row = ""
            row_cont = 1
            last_row_cont = 0
            while not resultSet.eof 
                nome = resultSet("item")
                local = resultSet("local")
                quantia = resultSet("quantia")
                
 
                if old_nome = nome then
                  row = row & "<tr><td>  "& local &" </td><td>  " & quantia & " </td></tr>"
                  row_cont = row_cont + 1
					 else
					 	row = replace(row,"ROW_P",row_cont)
					 	last_row_cont = row_cont
            		row_cont = 1
					   row = row & "<tr><td rowspan=ROW_P> " & nome & " </td><td>  " & local & " </td><td>  " & quantia & " </td></tr>"
					 	old_nome = nome
                end if
              
                resultSet.MoveNext()
             wend
				 	row = replace(row,"ROW_P",last_row_cont+1)
               response.write(row)
                row = ""
            %>
</table>
<table id="instrucoes" class="tabela">
	<tr> <th> Instru&ccedil&atildees </th> <tr>
	<tr>
		<td>
			<p>Nesta lista est&atildeo todos os itens do estoque agrupados por locais.</p>
		</td>
	</tr>

</table>
<% end if%>


<%if(request.queryString("tipo") = "pedidos") then%>


    <%          
            if (isEmpty(request.queryString("mes")) or request.queryString("mes") = "undefined") then
                'mes = month(now)
                mes = "*"
            else
                mes = request.queryString("mes")
            end if
            
            if (isEmpty(request.queryString("ano"))  or request.queryString("ano") = "undefined" ) then
                'ano = year(now)
                ano = "*"
            else
                ano = request.queryString("ano")
            end if
            
    
    %>
    <h1 class="titulo_pagina" > Pedidos </h1>
        Ano: <select name="ano">
                <option value="*">Todos</option>
                <%  
                    selected = ""
                    for i=2014 to 2017 %>
                    <% 
                        selected = "" 
                        if isnumeric(ano) then
                            if clng(ano) = i then
                                selected = "selected" 
                            end if
                        end if
                    %>
                    <option <%=selected%> value="<%=i%>"><%=i%></option>
                <% next %>
            </select>
        M&ecircs: <select name="mes">
                <option value="*">Todos</option>
                <%  for i=1 to 12 %>
                    <% 
                        selected = "" 
                        if isnumeric(mes) then
                            if clng(mes) = i then
                                selected = "selected" 
                            end if
                        end if
                    %>
                    <option <%=selected%> value="<%=i%>"><%=i%></option>
                <% next %>
            </select>
        <br><br>
    <table id="lista_locais" class="tabela">
    <tr>
        <th class="coluna_cod" > Cod. </th>  
        <th class="coluna_cod" > Respons&aacutevel </th>        
        <th class="coluna_status" > Status </th>          
        <th class="coluna_data" > Data do pedido </th>   
        <th class="coluna_data" > Previs&atildeo de Entrega</th>           
        <th class="coluna_data" > Data de Entrega </th>                
        <th class="coluna_descricao"> Descri&ccedil&atildeo </th>            
        <th> Itens </th>          
        <th colspan="3" class="botoes_tabela"> </th>
    </tr>
<%
            filtro = ""     
            if isNumeric(mes) then
                mes = " month(A.data_pedido) = "&mes
            else
                mes = ""
            end if
            
            
            if isNumeric(ano) then
                if len(mes) > 1 then
                   ano = " AND year(A.data_pedido) = "&ano
                else
                    ano = " year(A.data_pedido) = "&ano
                end if
            else
                ano = ""
            end if
            
            if len(ano) > 1 or len(mes) > 1 then
                filtro = " WHERE "&mes &"  "&ano
            else
                filtro = " "
            end if
            
            sql = "SELECT F.LOGIN as login ,A.cod_pedido, A.corrigido, A.descricao, A.data_pedido,A.data_previsao,A.data_entrega,A.status,"&_
                    " CAST((SELECT GROUP_CONCAT(CONCAT(CONVERT(D.nome_produto USING utf8),', Qtd: ', C.quantidade) separator '<br>"&_
                    "' ) FROM tb_est_itens_pedidos C "&_
                    "INNER JOIN TB_PRODUTOS D on D.COD_PRODUTO = C.cod_item "&_
                    "WHERE A.cod_pedido = C.cod_pedido ) AS CHAR) as itens "&_
                "FROM tb_est_pedidos A "&_
                "INNER JOIN TB_ADM_USUARIOS F on F.COD_USUARIO = A.cod_usuario"&_
                " "&filtro&_
                " ORDER BY A.cod_pedido DESC"
     
            set resultSet = objConexao.Execute(sql)
            while not resultSet.eof 
                cod_pedido = resultSet("cod_pedido") 
                usuario = resultSet("login")
                descricao = resultSet("descricao")
                itens = resultSet("itens")
                status = resultSet("status")
                conferido = resultSet("corrigido")
                data_pedido = resultSet("data_pedido")
                data_previsao = resultSet("data_previsao")
                data_entrega = resultSet("data_entrega")
                Dim temp_status : temp_status = status
                if status = "CD" then imgbaixa = "baixa_ativo.png"  else imgbaixa = "baixa_inativo.png" end if
                if status = "EA" then imgalterar = "icon_alterar.png"  else imgalterar = "icon_alterar_block.png" end if                
                if status = "EA" then status = "Em Andamento" else if status = "CN" then status = "Cancelado" else status = "Conclu&iacutedo" end if
            %>
                <tr> 
                    <td class="coluna_cod" >  <%=cod_pedido%> </td>
                   <td class="coluna_cod" >  <%=usuario%> </td>    
                    <td class="coluna_status" >   <%=status%> </td>  
                    <td class="coluna_data" >  <%=data_pedido%> </td>  
                    <td class="coluna_data" >  <%=data_previsao%> </td>  
                    <td class="coluna_data" >  <%=data_entrega%> </td>          
                    <td class="coluna_descricao">  <%=descricao%> </td>         
                    <td> <textarea style="display:none"><%=itens%></textarea>
                    	<table>
                        <%
                        if len(itens) > 0 then
                            itens_array = split(itens,"<br>")

                            for each y in itens_array
                                x = split(y,", Qtd:")
                                response.write("<tr>")
                                response.write("<td>" & x(0) & "</td>")
                                response.write("<td>" & x(1) & "</td>")
                                response.write("</tr>")
                            next
                        end if
                        %> 
                        </table>
                    </td>   
                    <td class="botao_tabela">
                        <% 
                        if temp_status = "CD" then
                            if conferido&"" <> "S" Then
                            %>
                            <center>
                                <a href='est_pedidos_correcao.asp?cod_pedido=<%=cod_pedido%>' target='_self'>
                                    Conferir
                                </a>
                                </br>
                                <a target="_blank" href='est_pedidos_correcao_imprimir.asp?d=1&c=<%=cod_pedido%>'><img src='images/icons/icon_imprimir.png' /></a>  
                            </center>
                            <%
                            else
                                response.write("<center>Conferido<br /><a target='_blank' href='est_pedidos_correcao_imprimir.asp?c="&cod_pedido&"'><img src='images/icons/icon_imprimir.png' /></a></center>")
                            end if
                        else
                        %>
                            <center>Aguardando conclus&atildeo</center><br />
                            <a target="_blank" href='est_pedidos_correcao_imprimir.asp?d=1&c=<%=cod_pedido%>'><img src='images/icons/icon_imprimir.png' /></a>  
                        <% end if %>
                    </td>

                    <td class="botao_tabela" onclick="visualizarPedido(<%=cod_pedido%>)"><img src="images/icons/icon_visualizar.png" alt="Visualizar Registro" title="Visualizar Registro" border="0"></td>

                    <td class="botao_tabela" onclick="alterarPedido(<%=cod_pedido%>,'<%=status%>')"><img src="images/icons/<%=imgalterar%>" alt="Alterar Registro" title="Alterar Registro" width="16" height="16" border="0"></td>
            </tr>  
            <%
            resultSet.MoveNext()
            wend
            %>
</table>

<table id="instrucoes" class="tabela">
	<tr> <th> Instru&ccedil&atildees </th> <tr>
	<tr>
		<td>
			<p>Nesta lista est&atildeo os pedidos j&aacute registrados.<br>
			Oferencendo a op&ccedil&atildeo de edi&ccedil&atildeo, quando o pedido n&atildeo esta marcado como "Conclu&iacutedo" ou "Cancelado", e a op&ccedil&atildeo de visualizar os detalhes do pedido. </p>
		</td>
	</tr>

</table>
<% end if%>


</div>
