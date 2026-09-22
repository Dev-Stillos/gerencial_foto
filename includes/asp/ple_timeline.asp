
<!-- #include file="../xqrtopqqetyynmkgfdconexao.asp" -->
<%
        Response.CharSet = "ISO-8859-1"
        AbrirConexao()
        projeto = ""
        set resultSet = objConexao.Execute("select cod_projeto, nome, descricao, inicio, termino, status from tb_ple_projetos WHERE status NOT IN (0) order by inicio")
                
        While not resultSet.Eof
                
            cod_projeto  = resultSet("cod_projeto")
            nome = resultSet("nome")
            descricao = resultSet("descricao")
            inicio  = resultSet("inicio")
            termino = resultSet("termino")			
	
            projeto = projeto & "[["&nome&"::"&descricao&"::"&inicio&"::"&termino&"]]"
        resultSet.MoveNext()
        Wend
        response.write(projeto)       
        resultSet.Close
        Set resultSet=Nothing
        FecharConexao()
                        
%>
