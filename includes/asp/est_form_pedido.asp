<!-- #include file="../../includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="../../includes/asp/valida_usuario.asp" -->
<% Response.Charset="ISO-8859-1" %>
<form name="cadastro_pedido" class="cadastro_pedido formulario">
    <input type="hidden" name="cod_pedido" value="0"/>
    <h1 class="titulo_pagina" > Cadastro de pedido </h1>
    <table border="1">
    
        <tr>
            <td class="coluna_formulario_label">
                <label> Status </label>
            </td>  
             <td class="coluna_formulario_fields">
                <select name="status">
                    <option value="0">Selecione um status </option>
                    <option value="EA"> Em andamento </option>
                    <option value="CD"> Conclu&iacutedo </option>
                    <option value="CN"> Cancelado </option>
                </select>
            </td>
            
        </tr>   
        
         <tr id="" class="">
        
            <td class="coluna_formulario_label">
                <label> Fornecedor</label>
            </td>  
             <td class="coluna_formulario_fields">
                <select name="fornecedor" class="sub_itens">
                    <option value="0">Selecione um fornecedor </option>
                    <%                    
                        AbrirConexao()
                        set resultSet = objConexao.Execute("select cod_fornecedor, nome_fantasia from TB_FORNECEDORES")
                                
                                While not resultSet.Eof
                                    
                                    cod_fornecedor  = resultSet("cod_fornecedor")
                                    nome = resultSet("nome_fantasia")

                    %>
                        <option value="<%=cod_fornecedor%>"><%=nome%></option>
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
        
        
         <tr id="form_sub_item" class="">
        
            <td class="coluna_formulario_label">
                <label> Sub Itens</label>
            </td>  
             <td class="coluna_formulario_fields">
                <select name="sub_itens" class="sub_itens" title="Selecione um item" >
                    <option value="0">Selecione um item </option>
                    <%                    
                        AbrirConexao()
                        set resultSet = objConexao.Execute("select COD_PRODUTO, nome_produto from TB_PRODUTOS ORDER BY nome_produto")
                                
                                While not resultSet.Eof
                                    
                                    cod_item  = resultSet("COD_PRODUTO")
                                    nome = resultSet("nome_produto")

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
                <input type="number" name="quantidade_sub_item" class="quantidade_sub_item" title="Quantia do item selecionado." value="0"/>
                <span class="moeda">R$</span> <input type="text" name="valor" class="valor_item" title="Valor total da quantia especificada." placeholder="1.000,00"/>                
                <span class="botao" id="add_sub_item" onclick="addSubItemPedido()">Adicionar</span>
                <div id="lista_sub_itens">   	
                		
                		<table>
                		</table>
                     <input type="hidden" name="list_sub_item" />               
                </div>
            </td>
            
        </tr> 
        
        
        <tr class="data_pedido">
             <td class="coluna_formulario_label">
                <label> Data do Pedido </label>
            </td>    
            <td class="coluna_formulario_fields">
                <input type="text" name="data_pedido" class="date_field"/><br>
            </td>
        </tr>    
        
         <tr>
             <td class="coluna_formulario_label">
                <label> Previs&atildeo de Entrega </label>
            </td>    
            <td class="coluna_formulario_fields">
                <input type="text" name="data_previsao" class="date_field datepicker"/><br>
            </td>
        </tr>    
        
        <tr>
             <td class="coluna_formulario_label">
                <label> Entregue em </label>
            </td>    
            <td class="coluna_formulario_fields">
                <input type="text" name="data_entrega" class="date_field datepicker"/><br>
            </td>
        </tr> 

        <tr>
             <td class="coluna_formulario_label">
                <label> Descri&ccedil&atildeo </label>
            </td>    
            <td class="coluna_formulario_fields">
                <textarea name="descricao"></textarea><br>
            </td>
        </tr>    
        <tr>
            <td  colspan="2" class="formulario_botoes">

                <span id="salvar_pedido"> Salvar </span>
                <span id="cancelar_pedido"> Cancelar </span>                
                
            </td>
        </tr>  
    </table>    
    
        
<table id="instrucoes" class="tabela">
	<tr> <th> Instru&ccedil&otildees </th> <tr>
	<tr>
		<td>
			<p>Utilize esta tela para realizar lan�amentos de pedidos.<br>
			Ao adicionar itens no pedido, ser� solicitado a quantia(quantitativa) do item e tamb�m o valor (total da quantia selecionada).<br><br>		
			
			Ao selecionar os status "Conclu&iacutedo" ou "Cancelado" o pedido n�o poder� ser alterado.
			Quando selecionado o status "Conclu&iacutedo", ser� realizada automaticamente a entrada, dos itens selecionados, no local "Estoque Padr&atildeo". </p>
			<p>
				<b>A movimenta��o autom�tica dos itens do estoque podera ser feita de duas formas:</b><br>
				1� Ao finalizar a produ��o dos itens dos pedidos.<br>
				2� Ao baixar a venda para autentica��o da diretoria.<br>
				Obs: Ser� feita a baixa dos itens em posse do vendedor de acordo com seu pedido.
				<br><br>
				A forma da movimenta��o autom�tica ser� definida pelo produto, neste poder� ser marcado como "Baixa por venda", "Baixa por produ��o" ou "N�o realiza baixa autom�tica".
			</p>
			
		</td>
	</tr>
</table>   

</form>
