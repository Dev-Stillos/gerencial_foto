<!-- #include file="../../includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="../../includes/asp/valida_usuario.asp" -->
<!-- #include file="../../includes/asp/functions.asp" -->
<% Response.Charset="ISO-8859-1" %>
<form name="lancamento_movimentacao" class="lancamento_movimentacao formulario">
    <input type="hidden" name="cod_movimentacao" value="0"/>
    <h1 class="titulo_pagina" > Lançamento de movimentação </h1>
    <table border="1">
    
        <tr>
            <td class="coluna_formulario_label">
                <label> Tipo </label>
            </td>  
             <td class="coluna_formulario_fields">
                <select name="tipo">
                    <option value="0">Selecione um tipo </option>
                    <option value="IN">Entrada </option>
                    <option value="OUT">Saída </option>
                    <option value="TRA">Transferência</option>
                </select>
            </td>
            
        </tr>   
        
        
         <tr id="" class="local_transferencia">
        
            <td class="coluna_formulario_label">
                <label class="local_origem"> Local de Armazenamento</label>
            </td>  
             <td class="coluna_formulario_fields">
                <select name="local_origem" class="sub_itens">
                    <option value="0">Selecione um local </option>
                    <%                    
                        AbrirConexao()
                        set resultSet = objConexao.Execute("select cod_local_armazenamento, nome from tb_est_local_armazenamento")
                                
                                While not resultSet.Eof
                                    
                                    cod_local_armazenamento  = resultSet("cod_local_armazenamento")
                                    nome = resultSet("nome")

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
        
        
         <tr id="" class="">
        
            <td class="coluna_formulario_label">
                <label class="local_destino"> Destino</label>
            </td>  
             <td class="coluna_formulario_fields">
                <select name="local" class="sub_itens">
                    <option value="0">Selecione um local </option>
                    <option value="-1"> Vendedor </option>
                    <%                    
                        AbrirConexao()
                        set resultSet = objConexao.Execute("select cod_local_armazenamento, nome from tb_est_local_armazenamento WHERE  cod_local_armazenamento NOT IN (SELECT cod_local_armazenamento FROM tb_est_rel_vendedor_local)")
                                
                                While not resultSet.Eof
                                    
                                    cod_local_armazenamento  = resultSet("cod_local_armazenamento")
                                    nome = resultSet("nome")

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
        
        <tr id="" class="hide select_vendedor">
        
            <td class="coluna_formulario_label">
                <label class=""> Vendedor</label>
            </td>  
             <td class="coluna_formulario_fields">
                <select name="vendedor" class="sub_itens">
                    <option value="0">Selecione um vendedor </option>

                    <%                    
                        AbrirConexao()
                        set resultSet = objConexao.Execute("SELECT COD_VENDEDOR, NOME_VENDEDOR FROM TB_VENDEDORES WHERE ATIVADO = 1 ORDER BY NOME_VENDEDOR;")
                                
                                While not resultSet.Eof
                                    
                                    cod_vendedor = resultSet("COD_VENDEDOR")
                                    nome = resultSet("NOME_VENDEDOR")
                                    response.write("<option value="&cod_vendedor&">"&nome&"</option>")

                                resultSet.MoveNext()
                                Wend
                                
                        resultSet.Close
                        Set resultSet=Nothing
                        FecharConexao()
                    %>
                </select>
             </td>
            
        </tr> 
        
         <tr id="form_sub_item" class="">
        
            <td class="coluna_formulario_label">
                <label> Sub Itens</label>
            </td>  
             <td class="coluna_formulario_fields">
                <select name="sub_itens" class="sub_itens">
                    <option value="0">Selecione um item </option>
                    <%                    
                        AbrirConexao()
                        set resultSet = objConexao.Execute("select cod_produto as cod_item, nome_produto as nome from TB_PRODUTOS ORDER BY nome_produto")
                                
                                While not resultSet.Eof
                                    
                                    cod_item  = resultSet("cod_item")
                                    nome = resultSet("nome")

                    %>
                        <option value="<%=cod_item%>"><%=resultSet("nome")%></option>
                    <%
                                resultSet.MoveNext()
                                Wend
                                
                        resultSet.Close
                        Set resultSet=Nothing
                        FecharConexao()
                    %>
                </select>
                <input type="number" name="quantidade_sub_item" class="quantidade_sub_item" value="0"/>
                <span class="botao" id="add_sub_item" onclick="addSubItemLancamento()">Adicionar</span>
                <div id="lista_sub_itens">  
                	<table>
                	
                	</table>                  
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

                <span id="salvar_lancamento_movimentacao"> Salvar </span>
                <span id="cancelar_lancamento_movimentacao"> Cancelar </span>                
                
            </td>
        </tr>  
    </table> 
    
    
    <table id="instrucoes" class="tabela">
	<tr> <th> Instruções </th> <tr>
	<tr>
		<td>
			<p>Utilize esta tela para realizar lançamentos manuais dos tipos: "Entrada", "Saída" e "Transferência".<br>
			Além do registro das movimentações, é nesta tela que será feita a transferência dos itens para os vendedores.</p>
			<br>
			<p>
				<b>A movimentação automática dos itens do estoque podera ser feita de duas formas:</b><br>
				1° Ao finalizar a produção dos itens dos pedidos.<br>
				2° Ao baixar a venda para autenticação da diretoria.<br>
				Obs: Será feita a baixa dos itens em posse do vendedor de acordo com seu pedido.
				<br><br>
				A forma da movimentação automática será definida pelo produto, neste poderá ser marcado como "Baixa por venda", "Baixa por produção" ou "Não realizar baixa automática".
			</p>
			
		</td>
	</tr>

</table>   
</form>
