<!-- #include file="../../includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="../../includes/asp/valida_usuario.asp" -->
<!-- #include file="functions.asp" -->

<form name="cadastro_rel_item" class="formulario cadastro_rel_item">
    
    <h1 class="titulo_pagina" > Relação Item - Estoque X Produção </h1>
    <table border="1">
         <tr>
            
            <td class="coluna_formulario_label">
                <label> Item ( produção )  </label>
            </td>        
            <td class="coluna_formulario_fields">
                <select name="cod_item_p">
                    <option value="0">Selecione um Item</option>
                    <%                    
                        Response.CharSet = "ISO-8859-1"
                        AbrirConexao()
                        dim resultSet
                        set resultSet = objConexao.Execute("SELECT COD_PRODUTO, NOME_PRODUTO FROM TB_PRODUTOS")
                                
                                While not resultSet.Eof
                                    
                                    cod_produto  = resultSet("COD_PRODUTO")
                                    nome_produto = resultSet("NOME_PRODUTO")

                    %>
                        <option value="<%=cod_produto%>"><%=nome_produto%></option>
                    <%
                                resultSet.MoveNext()
                                Wend
                                
                        resultSet.Close
                        Set resultSet=Nothing
       
                    %>
                </select>
            </td>
            
        </tr>    
                <tr>
            
            <td class="coluna_formulario_label">
                <label> Item ( Estoque ) </label>
            </td>        
            <td class="coluna_formulario_fields">
                <select name="cod_item_e">
                    <option value="0">Selecione um Item </option>
                    <%                    
               

                        set resultSet = objConexao.Execute("select cod_item, nome from tb_est_itens")
                                
                                While not resultSet.Eof
                                    
                                    cod_item  = resultSet("cod_item")
                                    nome = ConvertFromUTF8(resultSet("nome"))

                    %>
                        <option value="<%=cod_item%>"><%=nome%></option>
                    <%
                                resultSet.MoveNext()
                                Wend
                                
                        resultSet.Close
                        Set resultSet=Nothing
                        
                    %>
                </select>
            </td>
            
        </tr>   
        
                <tr>
            
            <td class="coluna_formulario_label">
                <label> Local de Armazenamento </label>
            </td>        
            <td class="coluna_formulario_fields">
                <select name="cod_local">
                    <option value="0">Selecione um local </option>
                    <%                    
               

                        set resultSet = objConexao.Execute("select cod_local_armazenamento, nome from tb_est_local_armazenamento")
                                
                                While not resultSet.Eof
                                    
                                    cod_local_armazenamento  = resultSet("cod_local_armazenamento")
                                    nome = ConvertFromUTF8(resultSet("nome"))

                    %>
                        <option value="<%=cod_local_armazenamento%>"><%=nome%></option>
                    <%
                                resultSet.MoveNext()
                                Wend
                                
                        resultSet.Close
                        Set resultSet=Nothing
                        FecharConexao()
                    %>
                </select>
            </td>
            
        </tr> 
        
        <tr>
             <td class="coluna_formulario_label">
                <label> Descrição </label>
            </td>    
            <td class="coluna_formulario_fields">
                <textarea name="descricao"></textarea><br>
            </td>
        </tr>    
        <tr>
            <td  colspan="2" class="formulario_botoes">

                <span id="salvar_rel_item"> Salvar </span>
                <span id="cancelar_rel_item"> Cancelar </span>                
                
            </td>
        </tr>  
    </table>    
</form>