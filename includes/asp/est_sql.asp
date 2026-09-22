<!-- #include file="../../includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="../../includes/asp/valida_usuario.asp" -->
<!-- #include file="../../includes/asp_json.asp" -->
<!-- #include file="functions.asp" -->

<%                    
    Response.CharSet = "ISO-8859-1"
    AbrirConexao()
    dim resultSet
    dim sql
    dim rows
    sql_log = ""    
            
    if(request("form") = "cadastro_item") then'Formulário de cadastro de itens
        if(request("acao") = "1") then
            sql = "INSERT INTO tb_est_itens VALUES(0,'"&request("nome")&"',"&request("cod_categoria")&",'"&request("descricao")&"','"&request("tipo")&"',now());"
            sql_log = sql_log&"<br>"&sql
            objConexao.Execute sql, rows
            if (rows>=1) then
                sql = "SELECT MAX(cod_item) AS ID FROM tb_est_itens"
                sql_log = sql_log&"<br>"&sql
                set resultSet = objConexao.Execute(sql)
                cod_item_pai = resultSet("ID")                
                
                if (request("tipo") = "C") then

                    sub_itens = request("list_sub_item")
                    sub_itens = split(sub_itens,";")
                    
                    for each item_composto in sub_itens
                        if (len(item_composto) > 3) then
                            item_composto = replace(item_composto,"[[","")
                            item_composto = replace(item_composto,"]]","")
                            item_composto = split(item_composto,":")
                            cod_item = item_composto(0)
                            qtd = item_composto(1)
                            sql = "INSERT INTO tb_est_item_composto VALUES (0,"&cod_item_pai&","&cod_item&","&qtd&");" 
                            sql_log = sql_log&"<br>"&sql  
                            objConexao.Execute sql, rows                            
                        
                        end if           
                    next
                end if
                
                response.write("Registro cadastrado com sucesso")
                
            else
                response.write("Erro ao cadastrar o registro")
            end if
            

                    
        elseif(request("acao") = "2" and isNumeric(request("cod_item"))) then
            

            sql = "UPDATE tb_est_itens set nome = '"&request("nome")&"', cod_categoria = "&request("cod_categoria")&", descricao = '"&request("descricao")&"', tipo = '"&request("tipo")&"' WHERE cod_item = "&request("cod_item")&""
            sql_log = sql_log&"<br>"&sql
            objConexao.Execute sql, rows
          
            if (rows>=1) then
                sql = "DELETE FROM tb_est_item_composto WHERE cod_item_pai = "&request("cod_item")
                sql_log = sql_log&"<br>"&sql
                set resultSet = objConexao.Execute(sql)
                cod_item_pai = request("cod_item")             
                
                if (request("tipo") = "C") then

                    sub_itens = request("list_sub_item")
                    sub_itens = split(sub_itens,";")
                    
                    for each item_composto in sub_itens
                        if (len(item_composto) > 3) then
                            item_composto = replace(item_composto,"[[","")
                            item_composto = replace(item_composto,"]]","")
                            item_composto = split(item_composto,":")
                            cod_item = item_composto(0)
                            qtd = item_composto(1)
                            sql = "INSERT INTO tb_est_item_composto VALUES (0,"&cod_item_pai&","&cod_item&","&qtd&");"
                            sql_log = sql_log&"<br>"&sql   
                            objConexao.Execute sql, rows                            
                        
                        end if           
                    next
                end if
                
                response.write("Registro alterado com sucesso")
            else
                response.write("Erro ao alterar")
            end if
                    
        elseif(request("acao") = "3" and isNumeric(request("cod_item"))) then
            
            sql = "SELECT DISTINCT B.nome FROM tb_est_item_composto A INNER JOIN tb_est_itens B ON A.cod_item_pai = B.cod_item WHERE A.cod_item = "&request("cod_item")&""
            sql_log = sql_log&"<br>"&sql
            set resultSet = objConexao.Execute(sql)
            listaItens = ""
            delete = 0
            while not resultSet.eof 
            
                listaItens = listaItens&" <li> "&resultSet("nome")
                delete = 1
                resultSet.MoveNext()
            wend
            
            sql = "SELECT cod_movimentacao FROM tb_est_itens_movimentacao WHERE cod_item = "&request("cod_item")&""
            sql_log = sql_log&"<br>"&sql
            set resultSet = objConexao.Execute(sql)
            while not resultSet.eof 
                delete = 2
                resultSet.MoveNext()
            wend
             
             
            sql = "SELECT cod_pedido FROM tb_est_itens_pedidos WHERE cod_item = "&request("cod_item")&""
            sql_log = sql_log&"<br>"&sql
            set resultSet = objConexao.Execute(sql)
            while not resultSet.eof 
                delete = 3
                resultSet.MoveNext()
            wend
                        
            if (delete = 0) then
            sql = "DELETE FROM tb_est_item_composto WHERE cod_item_pai = "&request("cod_item")&""
            sql_log = sql_log&"<br>"&sql
            objConexao.Execute sql, rows
            
            sql = "DELETE FROM tb_est_itens WHERE cod_item = "&request("cod_item")&""
            sql_log = sql_log&"<br>"&sql
            objConexao.Execute sql, rows     
            
                    if (rows>=1) then
                        response.write("Registro removido com sucesso")
                    else
                        response.write("Erro ao remover o registro")
                      
                    end if
                    
            elseif (delete = 2) then
            
                response.write("Não foi possível remover o registro. Movimentações Existentes.")

            elseif (delete = 3) then
            
                response.write("Não foi possível remover o registro. Pedidos Existentes.")

            else
            
                response.write("Não foi possível remover o registro. <br>Os seguintes itens estão relacionados: <br><br>"&ConvertFromUTF8(listaItens))
            end if       
                    
        elseif(request("acao") = "4" and isNumeric(request("cod_item"))) then

            Dim item
            Set item = jsObject()

            
            sql = "SELECT A.cod_item, A.descricao, A.tipo , A.nome as nome_item, B.cod_categoria FROM tb_est_itens A INNER JOIN tb_est_categoria B on A.cod_categoria = B.cod_categoria WHERE cod_item = "&request("cod_item")&""
            sql_log = sql_log&"<br>"&sql
            set resultSet = objConexao.Execute(sql)
            if not resultSet.eof then
                cod_item = resultSet("cod_item")
                nome_item = resultSet("nome_item")
                cod_categoria = resultSet("cod_categoria")
                descricao = resultSet("descricao")
                tipo = resultSet("tipo")                
                

            end if

            lista_sub_item = ""
            sql = "select A.cod_item, A.quantidade, B.nome from tb_est_item_composto A inner join tb_est_itens B on A.cod_item = B.cod_item WHERE A.cod_item_pai = "&request("cod_item")&""
            sql_log = sql_log&"<br>"&sql
            set resultSet = objConexao.Execute(sql)
                    
                    While not resultSet.Eof
                        cod = resultSet("cod_item")
                        nome = resultSet("nome")
                        qtd = resultSet("quantidade")
                        lista_sub_item = lista_sub_item&"[["&cod&"::"&nome&"::"&qtd&"]]"                     

                    resultSet.MoveNext()
                    Wend
                    
            resultSet.Close
            Set resultSet=Nothing                  
                    

            item("cod_item") = cod_item
            item("nome_item") = ConvertFromUTF8(nome_item)
            item("cod_categoria") = cod_categoria
            item("descricao") = ConvertFromUTF8(descricao)
            item("tipo") = ConvertFromUTF8(tipo)
            item("lista_sub_item") = ConvertFromUTF8(lista_sub_item)
            

            item.Flush
        
        else
            
            response.write("Comando não reconhecido")
            
        end if       
    
    end if    
    
    
    
    
    
    if(request("form") = "cadastro_categoria") then'Formulário de cadastro de itens
        if(request("acao") = "1") then
            
   
            sql = "INSERT INTO tb_est_categoria VALUES(0,'"&request("nome")&"','"&request("descricao")&"',now())"
            sql_log = sql_log&"<br>"&sql
            objConexao.Execute sql, rows
                    if (rows>=1) then
                        response.write("Registro cadastrado com sucesso")
                    else
                        response.write("Erro ao cadastrar o registro")
                    end if
                    
        elseif(request("acao") = "2" and isNumeric(request("cod_categoria"))) then
            

            sql = "UPDATE tb_est_categoria set nome = '"&request("nome")&"', descricao = '"&request("descricao")&"' WHERE cod_categoria = "&request("cod_categoria")&""
            sql_log = sql_log&"<br>"&sql
            objConexao.Execute sql, rows
                    if (rows>=1) then
                        response.write("Registro alterado com sucesso")
                    else
                        response.write("Erro ao alterar")
                    end if
                    
        elseif(request("acao") = "3" and isNumeric(request("cod_categoria"))) then
            

            sql = "SELECT nome FROM tb_est_itens WHERE cod_categoria = "&request("cod_categoria")&""
            sql_log = sql_log&"<br>"&sql
            set resultSet = objConexao.Execute(sql)
            listaItens = ""
            delete = 0
            while not resultSet.eof 
            
                listaItens = listaItens&" <li>"&resultSet("nome")
                delete = 1
                resultSet.MoveNext()
            wend
                        
            if (delete = 0) then
            
            sql = "DELETE FROM tb_est_categoria WHERE cod_categoria = "&request("cod_categoria")&""
            sql_log = sql_log&"<br>"&sql
            objConexao.Execute sql, rows
                    if (rows>=1) then
                        response.write("Registro removido com sucesso")
                    else
                        response.write("Erro ao remover o registro")
                    end if
            else
                response.write("Não foi possivel remover o registro.<br>Os seguintes itens estão relacionados: <br><br>"&ConvertFromUTF8(listaItens))
            end if           
                    
        elseif(request("acao") = "4" and isNumeric(request("cod_categoria"))) then

            Dim categoria
            Set categoria = jsObject()

            
            sql = "SELECT cod_categoria, nome , descricao FROM tb_est_categoria WHERE cod_categoria = "&request("cod_categoria")&""
            sql_log = sql_log&"<br>"&sql
            set resultSet = objConexao.Execute(sql)
            if not resultSet.eof then
                cod_categoria = resultSet("cod_categoria")
                nome = resultSet("nome")
                descricao = resultSet("descricao")
              
                

             end if
             

            categoria("cod_categoria") = cod_categoria
            categoria("nome") = ConvertFromUTF8(nome)
            categoria("descricao") = ConvertFromUTF8(descricao)
            

            categoria.Flush
        
        else
            
            response.write("Comando não reconhecido")
            
        end if       
        
        
        

        
    end if    
   

    if(request("form") = "cadastro_rel_item") then'Formulário de cadastro de itens
        if(request("acao") = "1") then
            
            sql = "SELECT cod_item_p FROM tb_est_rel_item WHERE cod_item_p = "&request("cod_item_p")&" AND cod_item_e = "&request("cod_item_e")&""
            sql_log = sql_log&"<br>"&sql
            objConexao.Execute sql, rows
            
            if (rows<1) then
                sql = "INSERT INTO tb_est_rel_item VALUES(0,"&request("cod_item_p")&","&request("cod_item_e")&","&request("cod_local")&",'"&request("descricao")&"')"
                sql_log = sql_log&"<br>"&sql
                objConexao.Execute sql, rows
                        if (rows>=1) then
                            response.write("Registro cadastrado com sucesso")
                        else
                            response.write("Erro ao cadastrar o registro")
                        end if
            else
                response.write("Erro: Esta relação já foi cadastrada anteriormente.")
            end if
        
        elseif(request("acao") = "3" and isNumeric(request("cod_rel_item"))) then
            

            sql = "DELETE FROM tb_est_rel_item WHERE cod_rel_item = "&request("cod_rel_item")&""
            sql_log = sql_log&"<br>"&sql
            objConexao.Execute sql, rows
                    if (rows>=1) then
                        response.write("Registro removido com sucesso")
                    else
                        response.write("Erro ao remover o registro")
                    end if       

        else
            
            response.write("Comando não reconhecido")
            
        end if       

    end if  
    
    
    
    
    
    
    
    if(request("form") = "cadastro_rel_item_venda") then'Formulário de cadastro de itens relacionados à venda
        if(request("acao") = "1") then
            
            sql = "SELECT cod_item_p FROM tb_est_rel_item_venda WHERE cod_item_p = "&request("cod_item_p")&" AND cod_item_e = "&request("cod_item_e")&""
            sql_log = sql_log&"<br>"&sql
            objConexao.Execute sql, rows         
            sql = "SELECT cod_item_p FROM tb_est_rel_item WHERE cod_item_p = "&request("cod_item_p")&" AND cod_item_e = "&request("cod_item_e")&""
            sql_log = sql_log&"<br>"&sql
            objConexao.Execute sql, rows2
            if (rows<1 and rows2<1) then
                sql = "INSERT INTO tb_est_rel_item_venda VALUES(0,"&request("cod_item_p")&","&request("cod_item_e")&","&request("cod_local")&",'"&request("descricao")&"')"
                sql_log = sql_log&"<br>"&sql

                objConexao.Execute sql, rows
                        if (rows>=1) then
                            response.write("Registro cadastrado com sucesso")
                        else
                            response.write("Erro ao cadastrar o registro")
                        end if
            elseif (rows>1) then
                response.write("Erro: Esta relação já foi cadastrada anteriormente.")
            else
                response.write("Erro: Esta relação já foi cadastrada anteriormente (Relações Itens - produção).")
            end if
        
        elseif(request("acao") = "3" and isNumeric(request("cod_rel_item_venda"))) then
            

            sql = "DELETE FROM tb_est_rel_item_venda WHERE cod_rel_item_venda = "&request("cod_rel_item_venda")&""
            sql_log = sql_log&"<br>"&sql
            
            objConexao.Execute sql, rows
                    if (rows>=1) then
                        response.write("Registro removido com sucesso")
                    else
                        response.write("Erro ao remover o registro")
                    end if       

        else
            
            response.write("Comando não reconhecido")
            
        end if       

    end if  
    
    

    
    
    if(request("form") = "cadastro_fornecedor") then'Formulário de cadastro de itens
        if(request("acao") = "1") then
            
   
            sql = "INSERT INTO tb_est_fornecedores VALUES(0,'"&request("nome")&"','"&request("descricao")&"')"
            sql_log = sql_log&"<br>"&sql
            
            objConexao.Execute sql, rows
                    if (rows>=1) then
                        response.write("Registro cadastrado com sucesso")
                    else
                        response.write("Erro ao cadastrar o registro")
                    end if
                    
        elseif(request("acao") = "2" and isNumeric(request("cod_fornecedor"))) then
            

            sql = "UPDATE tb_est_fornecedores set nome = '"&request("nome")&"', descricao = '"&request("descricao")&"' WHERE cod_fornecedor = "&request("cod_fornecedor")&""
            sql_log = sql_log&"<br>"&sql
            
            objConexao.Execute sql, rows
                    if (rows>=1) then
                        response.write("Registro alterado com sucesso")
                    else
                        response.write("Erro ao alterar")
                    end if
                    
        elseif(request("acao") = "3" and isNumeric(request("cod_fornecedor"))) then
            

            'sql = "SELECT nome FROM tb_est_itens WHERE cod_fornecedor = "&request("cod_fornecedor")&""
            'set resultSet = objConexao.Execute(sql)
            'listaItens = ""
            'delete = 0
            'while not resultSet.eof 
            
            '    listaItens = listaItens&", "&resultSet("nome")
            '    delete = 1
            '    resultSet.MoveNext()
            'wend
                        
            'if (delete = 0) then
         
            
            sql = "DELETE FROM tb_est_fornecedores WHERE cod_fornecedor = "&request("cod_fornecedor")&""
            sql_log = sql_log&"<br>"&sql
            
            objConexao.Execute sql, rows
                    if (rows>=1) then
                        response.write("Registro removido com sucesso")
                    else
                        response.write("Erro ao remover o registro")
                    end if
            'else
            '    response.write("Não foi possivel remover o registro. Os seguintes itens estão relacionados: "&listaItens)
            'end if           

        elseif(request("acao") = "4" and isNumeric(request("cod_fornecedor"))) then

            Dim fornecedor
            Set fornecedor = jsObject()

            
            sql = "SELECT cod_fornecedor, nome , descricao FROM tb_est_fornecedores WHERE cod_fornecedor = "&request("cod_fornecedor")&""
            sql_log = sql_log&"<br>"&sql
            
            set resultSet = objConexao.Execute(sql)
            if not resultSet.eof then
                cod_fornecedor = resultSet("cod_fornecedor")
                nome = resultSet("nome")
                descricao = resultSet("descricao")
            end if

            fornecedor("cod_fornecedor") = cod_fornecedor
            fornecedor("nome") = ConvertFromUTF8(nome)
            fornecedor("descricao") = ConvertFromUTF8(descricao)
            fornecedor.Flush
        else
            response.write("Comando não reconhecido")
        end if
    end if
    
    
    if(request("form") = "cadastro_local") then'Formulário de cadastro de itens
        if(request("acao") = "1") then
            
   
            sql = "INSERT INTO tb_est_local_armazenamento VALUES(0,'"&request("nome")&"','"&request("descricao")&"',now())"
            sql_log = sql_log&"<br>"&sql
            
            objConexao.Execute sql, rows
                    if (rows>=1) then
                        response.write("Registro cadastrado com sucesso")
                    else
                        response.write("Erro ao cadastrar o registro")
                    end if
                    
        elseif(request("acao") = "2" and isNumeric(request("cod_local_armazenamento"))) then
            

            sql = "UPDATE tb_est_local_armazenamento set nome = '"&request("nome")&"', descricao = '"&request("descricao")&"' WHERE cod_local_armazenamento = "&request("cod_local_armazenamento")&""
            sql_log = sql_log&"<br>"&sql
            
            objConexao.Execute sql, rows
                    if (rows>=1) then
                        response.write("Registro alterado com sucesso")
                    else
                        response.write("Erro ao alterar")
                    end if
                    
        elseif(request("acao") = "3" and isNumeric(request("cod_local_armazenamento"))) then
            

            sql = "SELECT cod_movimentacao FROM tb_est_movimentacao WHERE cod_local_armazenamento = "&request("cod_local_armazenamento")&""
            sql_log = sql_log&"<br>"&sql
            
            set resultSet = objConexao.Execute(sql)
            listaItens = ""
            delete = 0
            while not resultSet.eof 
            
                listaItens = listaItens&""&resultSet("cod_movimentacao")&", "
                delete = 1
                resultSet.MoveNext()
            wend
                        
            if (delete = 0) then
            
            sql = "DELETE FROM tb_est_local_armazenamento WHERE cod_local_armazenamento = "&request("cod_local_armazenamento")&""
            sql_log = sql_log&"<br>"&sql
            
            objConexao.Execute sql, rows
                    if (rows>=1) then
                        response.write("Registro removido com sucesso")
                    else
                        response.write("Erro ao remover o registro")
                    end if
            else
                response.write("Não foi possível remover o registro.<br>As seguintes movimentações estão relacionadas: "&listaItens)
            end if           
                    
        elseif(request("acao") = "4" and isNumeric(request("cod_local_armazenamento"))) then

            Dim local
            Set local = jsObject()

            sql = "SELECT cod_local_armazenamento, nome , descricao FROM tb_est_local_armazenamento WHERE cod_local_armazenamento = "&request("cod_local_armazenamento")&""
            sql_log = sql_log&"<br>"&sql
            
            set resultSet = objConexao.Execute(sql)
            if not resultSet.eof then
                cod_local_armazenamento = resultSet("cod_local_armazenamento")
                nome = resultSet("nome")
                descricao = resultSet("descricao")
            end if

            local("cod_local_armazenamento") = cod_local_armazenamento
            local("nome") = ConvertFromUTF8(nome)
            local("descricao") = ConvertFromUTF8(descricao)
            local.Flush
        
        else
            response.write("Comando não reconhecido")
            
        end if
    end if        
    
    
    
    if(request("form") = "lancamento_movimentacao") then'Formulário de lançamento de movimentações.
    
        if(request("acao") = "1") then
                
                cod_usuario = Session("COD_USUARIO")                               
                if (request("local") < 0) then

                    sql = "SELECT A.cod_local_armazenamento AS ID FROM tb_est_local_armazenamento A INNER JOIN tb_est_rel_vendedor_local B ON "&_
                          " A.cod_local_armazenamento = B.cod_local_armazenamento WHERE B.cod_vendedor = "&request("vendedor")
                    sql_log = sql_log&"<br>"&sql
                    set resultSet = objConexao.Execute(sql)                    
                    if NOT resultSet.EOF then 

                      cod_local = resultSet("ID")
                      local = cod_local
                      
                    else   
                       sql = "INSERT INTO tb_est_local_armazenamento VALUES (0,"&_
                       "(SELECT NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR = "&request("vendedor")&" LIMIT 1),'vendedor',now())"
                       sql_log = sql_log&"<br>"&sql
                       
                       objConexao.Execute sql, rows
                       if (rows>=1) then
                           sql = "SELECT MAX(cod_local_armazenamento) AS ID FROM tb_est_local_armazenamento"
                           sql_log = sql_log&"<br>"&sql
                           
                           set resultSet = objConexao.Execute(sql)
                           cod_local = resultSet("ID")
                           sql = "INSERT INTO tb_est_rel_vendedor_local VALUES (0,"&cod_local&","&request("vendedor")&",'');"
                           sql_log = sql_log&"<br>"&sql
                           
                           objConexao.Execute sql, rows
                           if (rows>=1) then
                              local = cod_local
                              
                           else
                              response.write("Ocorreu um erro ao tentar realizar sua solicitação.\n Tente novamente.") 
                              Response.End                       
                           end if
                       else
                           response.write("Ocorreu um erro ao tentar realizar sua solicitação.\n Tente novamente.")
                           Response.End
                       end if               
                    end if  
                else
                    local = request("local") 
                  
                end if

                if(request("tipo") = "TRA") then
                    sql = "INSERT INTO tb_est_movimentacao VALUES(0,'OUT','"&ConvertToUTF(request("descricao"))&"','"&request("local_origem")&"',now(),"&cod_usuario&");"
                    sql_log = sql_log&"<br>"&sql

                    objConexao.Execute sql, rows
                    sql = "SELECT MAX(cod_movimentacao) AS ID FROM tb_est_movimentacao"
                    sql_log = sql_log&"<br>"&sql

                    set resultSet = objConexao.Execute(sql)
                    cod_movimentacao_out = resultSet("ID")                           

                    sql = "INSERT INTO tb_est_movimentacao VALUES(0,'IN','"&ConvertToUTF(request("descricao"))&"','"&local&"',now(),"&cod_usuario&");"
                    sql_log = sql_log&"<br>"&sql

                    objConexao.Execute sql, rows
                    sql = "SELECT MAX(cod_movimentacao) AS ID FROM tb_est_movimentacao"
                    sql_log = sql_log&"<br>"&sql

                    set resultSet = objConexao.Execute(sql)
                    cod_movimentacao_in = resultSet("ID")   

                    sub_itens = request("list_sub_item")
                    sub_itens = split(sub_itens,";")

                    'GRAVAMOS A MOVIMENTAÇÂO DE TRANSFERENCIA EM UMA TABELA A PARTE PARA MANTERMOS O HISTORICO
                    sql = "INSERT INTO tb_est_movimentacao_transferencias values (0, "&sqlNumero(cod_movimentacao_out)&","&sqlNumero(cod_movimentacao_in)&");"
                    objConexao.Execute sql, rows
                    sql = ""

                    for each item_composto in sub_itens
                        if (len(item_composto) > 3) then
                            item_composto = replace(item_composto,"[[","")
                            item_composto = replace(item_composto,"]]","")
                            item_composto = split(item_composto,":")
                            cod_item = item_composto(0)
                            qtd = item_composto(1)
                            sql = "SELECT COD_PRODUTO FROM TB_PRODUTOS WHERE COD_PRODUTO = "&cod_item
                            sql_log = sql_log&"<br>" & sql
                            set resultSet = objConexao.Execute(sql)
                            While not resultSet.Eof
                                cod = resultSet("COD_PRODUTO")

                                sql = "INSERT INTO tb_est_itens_movimentacao VALUES (0,"&cod_movimentacao_out&","&cod&","&qtd&");"  
                                sql_log = sql_log&"<br>"&sql
                                objConexao.Execute sql, rows  

                                sql = "INSERT INTO tb_est_itens_movimentacao VALUES (0,"&cod_movimentacao_in&","&cod&","&qtd&");"  
                                sql_log = sql_log & "<br>" & sql
                                objConexao.Execute sql, rows

                            resultSet.MoveNext()
                            Wend
                        end if
                    next
                    
                    sql = "call atualizarEstoque()"
                    sql_log = sql_log & "<br>" & sql
                    
                    set resultSet = objConexao.Execute(sql)
     
                    response.write("Movimentação cadastrada com sucesso")
                
                else

                    sql = "INSERT INTO tb_est_movimentacao VALUES(0,'"&request("tipo")&"','"&ConvertToUTF(request("descricao"))&"','"&local&"',now(),"&cod_usuario&");"
                    sql_log = sql_log&"<br>"&sql
                    
                    objConexao.Execute sql, rows
                    if (rows>=1) then
                        sql = "SELECT MAX(cod_movimentacao) AS ID FROM tb_est_movimentacao"
                        sql_log = sql_log&"<br>"&sql
                        
                        set resultSet = objConexao.Execute(sql)
                        cod_movimentacao = resultSet("ID")                

                        sub_itens = request("list_sub_item")
                        sub_itens = split(sub_itens,";")
                        
                        for each item_composto in sub_itens
                            if (len(item_composto) > 3) then
                                item_composto = replace(item_composto,"[[","")
                                item_composto = replace(item_composto,"]]","")
                                item_composto = split(item_composto,":")
                                cod_item = item_composto(0)
                                qtd = item_composto(1)
                                sql = "SELECT COD_PRODUTO FROM TB_PRODUTOS WHERE COD_PRODUTO = "&cod_item
                                sql_log = sql_log&"<br>"&sql
                                set resultSet = objConexao.Execute(sql)
        
                                While not resultSet.Eof
                                    cod = resultSet("COD_PRODUTO")
                                    sql = "INSERT INTO tb_est_itens_movimentacao VALUES (0,"&cod_movimentacao&","&cod&","&qtd&");"
                                    sql_log = sql_log&"<br>"&sql
                                    objConexao.Execute(sql)
                                    resultSet.MoveNext()
                                Wend
                            end if
                        next

                        sql = "call atualizarEstoque()"
                        sql_log = sql_log&"<br>"&sql
                        set resultSet = objConexao.Execute(sql)
         
                        response.write("Movimentação cadastrada com sucesso")
                    else
                        response.write("Erro ao cadastrar a movimentação")
                    end if
                End if

        elseif(request("acao") = "2" and isNumeric(request("cod_item"))) then
            

            sql = "UPDATE tb_est_itens set nome = '"&request("nome")&"', cod_categoria = "&request("cod_categoria")&", descricao = '"&request("descricao")&"', tipo = '"&request("tipo")&"' WHERE cod_item = "&request("cod_item")&""
            sql_log = sql_log&"<br>"&sql
            
            objConexao.Execute sql, rows
          
            if (rows>=1) then
                sql = "DELETE FROM tb_est_item_composto WHERE cod_item_pai = "&request("cod_item")
                sql_log = sql_log&"<br>"&sql
                
                set resultSet = objConexao.Execute(sql)
                cod_item_pai = request("cod_item")             
                
                if (request("tipo") = "C") then

                    sub_itens = request("list_sub_item")
                    sub_itens = split(sub_itens,";")
                    
                    for each item_composto in sub_itens
                        if (len(item_composto) > 3) then
                            item_composto = replace(item_composto,"[[","")
                            item_composto = replace(item_composto,"]]","")
                            item_composto = split(item_composto,":")
                            cod_item = item_composto(0)
                            qtd = item_composto(1)
                            sql = "INSERT INTO tb_est_item_composto VALUES (0,"&cod_item_pai&","&cod_item&","&qtd&");"   
                            sql_log = sql_log&"<br>"&sql

                            objConexao.Execute sql, rows
                        end if           
                    next
                end if
                
                response.write("Registro alterado com sucesso")
            else
                response.write("Erro ao alterar")
            end if
                    
        elseif(request("acao") = "3" and isNumeric(request("cod_item"))) then
            sql = "SELECT DISTINCT B.nome FROM tb_est_item_composto A INNER JOIN tb_est_itens B ON A.cod_item_pai = B.cod_item WHERE A.cod_item = "&request("cod_item")&""
            sql_log = sql_log&"<br>"&sql
            
            set resultSet = objConexao.Execute(sql)
            listaItens = ""
            delete = 0
            while not resultSet.eof 
            
                listaItens = listaItens&", "&resultSet("nome")
                delete = 1
                resultSet.MoveNext()
            wend
            
            sql = "SELECT cod_movimentacao FROM tb_est_movimentacao WHERE cod_item = "&request("cod_item")&""
            sql_log = sql_log&"<br>"&sql

            set resultSet = objConexao.Execute(sql)
            while not resultSet.eof 
                delete = 2
                resultSet.MoveNext()
            wend

            if (delete = 0) then
            sql = "DELETE FROM tb_est_item_composto WHERE cod_item_pai = "&request("cod_item")&""
            sql_log = sql_log&"<br>"&sql
            
            objConexao.Execute sql, rows
            
            sql = "DELETE FROM tb_est_itens WHERE cod_item = "&request("cod_item")&""
            sql_log = sql_log&"<br>"&sql
            
            objConexao.Execute sql, rows     
            
                if (rows>=1) then
                    response.write("Registro removido com sucesso")
                else
                    response.write("Erro ao remover o registro")
                  
                end if

            elseif (delete = 2) then
                response.write("Não foi possível remover o registro. Movimentações Existentes.")

            else
                response.write("Não foi possível remover o registro. Os seguintes itens estão relacionados: "&listaItens)
            end if

        elseif(request("acao") = "4" and isNumeric(request("cod_item"))) then

            Dim item_lancamento
            Set item_lancamento = jsObject()

            sql = "SELECT A.cod_item, A.descricao, A.tipo , A.nome as nome_item, B.cod_categoria FROM tb_est_itens A INNER JOIN tb_est_categoria B on A.cod_categoria = B.cod_categoria WHERE cod_item = "&request("cod_item")&""
            sql_log = sql_log&"<br>"&sql
            
            set resultSet = objConexao.Execute(sql)
            if not resultSet.eof then
                cod_item = resultSet("cod_item")
                nome = resultSet("nome_item")
                cod_categoria = resultSet("cod_categoria")
                descricao = resultSet("descricao")
                tipo = resultSet("tipo")                

            end if

            lista_sub_item = ""
            sql = "select A.cod_item, A.quantidade, B.nome from tb_est_item_composto A inner join tb_est_itens B on A.cod_item = B.cod_item WHERE A.cod_item_pai = "&request("cod_item")&""
            sql_log = sql_log&"<br>"&sql
            set resultSet = objConexao.Execute(sql)

                While not resultSet.Eof
                    cod = resultSet("cod_item")
                    nome = resultSet("nome")
                    qtd = resultSet("quantidade")
                    lista_sub_item = lista_sub_item&"[["&cod&"::"&nome&"::"&qtd&"]]"
                resultSet.MoveNext()
                Wend

            resultSet.Close
            Set resultSet=Nothing

            item_lancamento("cod_item") = cod_item
            item_lancamento("nome_item") = ConvertFromUTF8(nome)
            item_lancamento("cod_categoria") = cod_categoria
            item_lancamento("descricao") = ConvertFromUTF8(descricao)
            item_lancamento("tipo") = ConvertFromUTF8(tipo)
            item_lancamento("lista_sub_item") = ConvertFromUTF8(lista_sub_item)
            item_lancamento.Flush

        elseif(request("acao") = "5" and isNumeric(request("cod_local")) and isNumeric(request("cod_item"))) then

            data = ""
            verif = 0
                if (request("cod_local") < 0) then
                sql = "SELECT A.cod_local_armazenamento AS ID FROM tb_est_local_armazenamento A INNER JOIN tb_est_rel_vendedor_local B ON "&_
                      " A.cod_local_armazenamento = B.cod_local_armazenamento WHERE B.cod_vendedor = "&request("vendedor")
                sql_log = sql_log&"<br>"&sql
                
                set resultSet = objConexao.Execute(sql)                    
                if NOT resultSet.EOF then 
                    cod_local = resultSet("ID")
                    local = cod_local

                else   
                   sql = "INSERT INTO tb_est_local_armazenamento VALUES (0,"&_
                   "(SELECT NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR = "&request("vendedor")&" LIMIT 1),'vendedor',now())"
                   sql_log = sql_log&"<br>"&sql
                   
                   objConexao.Execute sql, rows
                   if (rows>=1) then
                       sql = "SELECT MAX(cod_local_armazenamento) AS ID FROM tb_est_local_armazenamento"
                       sql_log = sql_log&"<br>"&sql
                       
                       set resultSet = objConexao.Execute(sql)
                       cod_local = resultSet("ID")
                       sql = "INSERT INTO tb_est_rel_vendedor_local VALUES (0,"&cod_local&","&request("vendedor")&",'');"
                       sql_log = sql_log&"<br>"&sql
                       
                       objConexao.Execute sql, rows
                       if (rows>=1) then
                          local = cod_local
                          
                       else
                          response.write("Ocorreu um erro ao tentar realizar sua solicitação.\n Tente novamente.") 
                          Response.End                       
                       end if
                   else
                       response.write("Ocorreu um erro ao tentar realizar sua solicitação.\n Tente novamente.")
                       Response.End
                   end if               
                end if  
            else
                local = request("cod_local") 

            end if
                
            sql = "SELECT COD_PRODUTO  FROM TB_PRODUTOS WHERE COD_PRODUTO = "&request("cod_item")    
            sql_log = sql_log&"<br>"&sql
            set resultSet = objConexao.Execute(sql)
            While not resultSet.Eof
                cod = resultSet("COD_PRODUTO")

                sql = "SELECT cod_item,cod_local_armazenamento,quantia FROM tb_est_resumo WHERE quantia >= "&request("quantia")&"  AND cod_item ="&request("cod_item")&" AND cod_local_armazenamento = "&local

                sql_log = sql_log&"<br>"&sql
                set resultSet2 = objConexao.Execute(sql)

                While not resultSet2.Eof
                    verif = 1
                resultSet2.MoveNext()
                Wend

                resultSet2.Close
                Set resultSet2=Nothing

            resultSet.MoveNext()
            Wend

            resultSet.Close
            Set resultSet=Nothing 

            if verif = 1 then
                data = ""
            else
                data = "Esta quantia não esta disponível neste local de estoque.<br>Consulte a quantia disponível ou tente a partir de outro local de armazenamento.<br>"
            end if
            response.write(data)

        else
            response.write("Comando não reconhecido")
            
        end if
    end if   
    
    
    
    if(request("form") = "cadastro_pedido") then'Formulário de lançamento de movimentações
        if(request("acao") = "1") then
            cod_usuario = Session("COD_USUARIO")
            
            data_previsao_array = split(request("data_previsao"),"/")
            data_previsao = data_previsao_array(2)&"-"&data_previsao_array(1)&"-"&data_previsao_array(0)
            
            if len(request("data_entrega"))>9 then
                data_entrega_array = split(request("data_entrega"),"/")
                data_entrega = "'"&data_entrega_array(2)&"-"&data_entrega_array(1)&"-"&data_entrega_array(0)&"'"
            else
                data_entrega = "NULL"    
            end if
            
            sql = "INSERT INTO tb_est_pedidos VALUES(0,"&request("fornecedor")&","&cod_usuario&",'"&request("status")&"',now(),'"&data_previsao&"',"&data_entrega&",'"&ConvertToUTF(request("descricao"))&"',NULL,NULL);"
            sql_log = sql_log&"<br>"&sql
            

            objConexao.Execute sql, rows

            
            'if request("status") = "CD" then
            '   COMENTADO POIS A ENTRADA NO ENTOQUE É FEITA NO MOMENTO DA CORRECAO
            '   sql = "SELECT MAX(cod_pedido) AS ID FROM tb_est_pedidos"
            '   sql_log = sql_log&"<br>"&sql

            '   set resultSet = objConexao.Execute(sql)
            '   cod_pedido = resultSet("ID") 
            '   cod_usuario = Session("COD_USUARIO")
            '   sql = "INSERT INTO tb_est_movimentacao VALUES(0,'IN',N'Pedido Cod: "&cod_pedido&" ',1,now(),"&cod_usuario&");"
            '   sql_log = sql_log&"<br>"&sql
                
            '   objConexao.Execute sql, rows
            '   sql = "SELECT MAX(cod_movimentacao) AS ID FROM tb_est_movimentacao"
            '   sql_log = sql_log&"<br>"&sql
                
            '   set resultSet = objConexao.Execute(sql)
            '   cod_movimentacao = resultSet("ID")   
 
            '   sql = "INSERT INTO tb_est_movimento_pedido VALUES(0,"&cod_movimentacao&","&cod_pedido&");"
            '   sql_log = sql_log&"<br>"&sql

            '   objConexao.Execute sql, rows
            'end if
            
            if (rows>=1) then
                sql = "SELECT MAX(cod_pedido) AS ID FROM tb_est_pedidos"
                sql_log = sql_log&"<br>"&sql
                
                set resultSet = objConexao.Execute(sql)
                cod_pedido = resultSet("ID")                

                sub_itens = request("list_sub_item")
                sub_itens = split(sub_itens,";")
                
                for each item_composto in sub_itens
                    if (len(item_composto) > 3) then
                        item_composto = replace(item_composto,"[[","")
                        item_composto = replace(item_composto,"]]","")
                        item_composto = split(item_composto,":")
                        cod_item = item_composto(0)
                        qtd = item_composto(1)
                        val = item_composto(2)       
                        val = replace(val,".","")                        
                        val = replace(val,",",".")
                        
                        sql = "INSERT INTO tb_est_itens_pedidos VALUES (0,"&cod_pedido&","&cod_item&","&qtd&","&val&",0, 0);"  
                        sql_log = sql_log&"<br>"&sql
                        
                        objConexao.Execute sql 

                        'if request("status") = "CD" then
                            'sql = "INSERT INTO tb_est_itens_movimentacao VALUES (0,"&cod_movimentacao&","&cod_item&","&qtd&");"   
                            'sql_log = sql_log&"<br>"&sql
                            
                            'objConexao.Execute sql, rows                         
                        'end if                         
                    
                    end if           
                next
                 
                 
                'if request("status") = "CD" then            
                    'sql = "call atualizarEstoque()"
                    'sql_log = sql_log&"<br>"&sql
                    
                    'set resultSet = objConexao.Execute(sql)
                'end if
                
                response.write("Pedido cadastrado com sucesso")
                
            else
                response.write("Erro ao cadastrar o pedido")
            end if
            

                    
        elseif(request("acao") = "2" and isNumeric(request("cod_pedido"))) then
            
            data_previsao_array = split(request("data_previsao"),"/")
            data_previsao = data_previsao_array(2)&"-"&data_previsao_array(1)&"-"&data_previsao_array(0)
            
            if len(request("data_entrega"))>9 then
                data_entrega_array = split(request("data_entrega"),"/")
                data_entrega = "'"&data_entrega_array(2)&"-"&data_entrega_array(1)&"-"&data_entrega_array(0)&"'"
            else
                data_entrega = "NULL"    
            end if
            
            sql = "UPDATE tb_est_pedidos set cod_fornecedor = "&request("fornecedor")&", status = '"&request("status")&"', descricao = '"&ConvertToUTF(request("descricao"))&"', data_previsao = '"&data_previsao&"', data_entrega = "&data_entrega&" WHERE cod_pedido = "&request("cod_pedido")&""
            sql_log = sql_log&"<br>"&sql
            
        
            objConexao.Execute sql, rows
            'if request("status") = "CD" then
                'cod_usuario = Session("COD_USUARIO")
                'sql = "INSERT INTO tb_est_movimentacao VALUES(0,'IN',N'Pedido Cod: "&request("cod_pedido")&" ',1,now(),"&cod_usuario&");"
                'sql_log = sql_log&"<br>"&sql
                
                'objConexao.Execute sql, rows
                'sql = "SELECT MAX(cod_movimentacao) AS ID FROM tb_est_movimentacao"
                'sql_log = sql_log&"<br>"&sql
                
                'set resultSet = objConexao.Execute(sql)
                'cod_movimentacao = resultSet("ID")     
                'cod_pedido = request("cod_pedido")
                'sql = "INSERT INTO tb_est_movimento_pedido VALUES(0,"&cod_movimentacao&","&cod_pedido&");"
                'sql_log = sql_log&"<br>"&sql
                
                'objConexao.Execute sql, rows
            'end if
  
            if (rows>=1) then
                sql = "DELETE FROM tb_est_itens_pedidos WHERE cod_pedido = "&request("cod_pedido")
                sql_log = sql_log&"<br>"&sql
                
                set resultSet = objConexao.Execute(sql)
                cod_pedido = request("cod_pedido")             
                        
                sub_itens = request("list_sub_item")
                sub_itens = split(sub_itens,";")
                
                for each item_composto in sub_itens
                    if (len(item_composto) > 3) then
                        item_composto = replace(item_composto,"[[","")
                        item_composto = replace(item_composto,"]]","")
                        item_composto = split(item_composto,":")
                        cod_item = item_composto(0)
                        qtd = item_composto(1)
                        val = item_composto(2)       
                        val = replace(val,".","")                        
                        val = replace(val,",",".")
                        
                        sql = "INSERT INTO tb_est_itens_pedidos VALUES (0,"&cod_pedido&","&cod_item&","&qtd&","&val&",0, 0);"  
                        sql_log = sql_log&"<br>"&sql
                        
                        objConexao.Execute sql, rows   

                        'if request("status") = "CD" then
                            'sql = "INSERT INTO tb_est_itens_movimentacao VALUES (0,"&cod_movimentacao&","&cod_item&","&qtd&");"   
                            'sql_log = sql_log&"<br>"&sql
                            
                            'objConexao.Execute sql, rows                         
                        'end if                        
                    
                    end if           
                next
                
                'if request("status") = "CD" then            
                    'sql = "call atualizarEstoque()"
                    'sql_log = sql_log&"<br>"&sql
                    
                    'set resultSet = objConexao.Execute(sql)
                'end if
                
                response.write("Registro alterado com sucesso")
            else
                response.write("Erro ao alterar")
            end if
                    
       
        elseif(request("acao") = "4" and isNumeric(request("cod_pedido"))) then

            Dim pedido
            Set pedido = jsObject()

            
            sql = "SELECT A.cod_pedido, A.descricao, A.status, A.data_pedido, A.data_previsao, A.data_entrega, A.cod_fornecedor FROM tb_est_pedidos A WHERE cod_pedido = "&request("cod_pedido")&""
            sql_log = sql_log&"<br>"&sql
            
            set resultSet = objConexao.Execute(sql)
            if not resultSet.eof then
                cod_pedido = resultSet("cod_pedido")
                status = resultSet("status")
                data_pedido = resultSet("data_pedido")
                data_previsao = resultSet("data_previsao")
                data_entrega = resultSet("data_entrega")               
                cod_fornecedor = resultSet("cod_fornecedor")
                descricao = resultSet("descricao")
            
                

            end if

            lista_sub_item = ""
            sql = "select A.cod_item, A.quantidade,A.valor, B.nome_produto as nome from tb_est_itens_pedidos A inner join TB_PRODUTOS B on A.cod_item = B.COD_PRODUTO WHERE A.cod_pedido = "&request("cod_pedido")&""
            sql_log = sql_log&"<br>"&sql
            set resultSet = objConexao.Execute(sql)
                    
                    While not resultSet.Eof
                        cod = resultSet("cod_item")
                        nome = resultSet("nome")
                        qtd = resultSet("quantidade")
                        val = resultSet("valor")
                        
                        lista_sub_item = lista_sub_item&"[["&cod&"::"&nome&"::"&qtd&"::"&val&"]]"                     

                    resultSet.MoveNext()
                    Wend
                    
            resultSet.Close
            Set resultSet=Nothing                  
                    

            pedido("cod_pedido") = cod_pedido
            pedido("nome_item") = nome
            pedido("fornecedor") = cod_fornecedor
            pedido("descricao") = descricao
            pedido("status") = status
            pedido("data_pedido") = data_pedido
            pedido("data_previsao") = data_previsao
            pedido("data_entrega") = data_entrega                        
            pedido("lista_sub_item") = lista_sub_item
            

            pedido.Flush

        else
            
            response.write("Comando não reconhecido")
            
        end if       
    
    end if        
    


    
    ' Envio de E-mail - Homologação/Debug
    para = "luiz@stillos.com.br"
    copia = ""
    copiaOculta = ""
    assunto = "Sistema de Reportagem - Ações"
    mensagem = "Log:<br>"&_
      " Sql: "&sql_log&_
      "<br> Form: "&request("form")&_  
      "<br> Ação: "&request("acao")&_               
      "<br> Usuário(cod): "&Session("NOME")&" ("&Session("COD_USUARIO")&")"&_  
      "<br> Departamento: "&Session("DEPARTAMENTO")&_      
      "<br> Data: "&Now() 
      mensagem = replace(mensagem,"'","\")                  
      sql = "INSERT INTO tb_est_log_acao(id,cod_usuario,nome_usuario,nome_departamento,form,acao,descricao,data)"&_
    "VALUES(0,"&Session("COD_USUARIO")&",'"&Session("NOME")&"','"&Session("DEPARTAMENTO")&"','"&request("form")&"','"&request("acao")&"','"&mensagem&"',now())"
    objConexao.Execute sql, rows   
    on error resume next  
    if (rows<1) then
       enviarEmail para,copia,copiaOculta,"Erro ao salvar Log",mensagem
       If err <> 0 Then
       
       end if
    end if
    if(Session("COD_USUARIO") <> 1)then
       enviarEmail para,copia,copiaOculta,assunto,mensagem
       If err <> 0 Then
       
       end if
    end if
    FecharConexao()

%>
               
               
