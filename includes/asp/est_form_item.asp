<!-- #include file="../../includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="../../includes/asp/valida_usuario.asp" -->
<!-- #include file="functions.asp" -->

<form name="cadastro_item" class="formulario cadastro_item">
    
    <h1 class="titulo_pagina" > Cadastro de item </h1>
    <table border="1">
        <tr>
            <input type="hidden" name="cod_item" value="0"/>
            <td class="coluna_formulario_label">
                <label> Nome </label>
            </td>        
            <td class="coluna_formulario_fields">
                <input type="text" name="nome"/>
            </td>
            
        </tr>    
        <tr>
            
            <td class="coluna_formulario_label">
                <label> Categoria </label>
            </td>        
            <td class="coluna_formulario_fields">
                <select name="cod_categoria">
                    <option value="0">Selecione uma categoria </option>
                    <%                    
                        Response.CharSet = "ISO-8859-1"
                        AbrirConexao()
                        dim resultSet
                        set resultSet = objConexao.Execute("select cod_categoria, nome from tb_est_categoria")
                                
                                While not resultSet.Eof
                                    
                                    cod_categoria  = resultSet("cod_categoria")
                                    nome_categoria = ConvertFromUTF8(resultSet("nome"))

                    %>
                        <option value="<%=cod_categoria%>"><%=nome_categoria%></option>
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
        <tr class="hide">
        
            <td class="coluna_formulario_label">
                <label> Tipo </label>
            </td>  
             <td class="coluna_formulario_fields">
                <select name="tipo">
                    <option value="0">Selecione um tipo </option>
                    <option value="S">Simples </option>
                    <option value="C">Composto </option>
                </select>
            </td>
            
        </tr>    
        
        <tr id="form_sub_item" class="hide">
        
            <td class="coluna_formulario_label">
                <label> Sub Itens</label>
            </td>  
             <td class="coluna_formulario_fields">
                <select name="sub_itens" class="sub_itens">
                    <option value="0">Selecione um item </option>
                    <%                    
                        AbrirConexao()
                        set resultSet = objConexao.Execute("select cod_item, nome from tb_est_itens WHERE tipo = 'S'")
                                
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
                        FecharConexao()
                    %>
                </select>
                <input type="number" name="quantidade_sub_item" class="quantidade_sub_item" value="0"/>
                <span class="botao" id="add_sub_item" onclick="addSubItem()">Adicionar</span>
                <div id="lista_sub_itens">                    
                     <input type="hidden" name="list_sub_item" />               
                </div>
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

                <span id="salvar_item"> Salvar </span>
                <span id="cancelar_item"> Cancelar </span>                
                
            </td>
        </tr>  
    </table>    
</form>
