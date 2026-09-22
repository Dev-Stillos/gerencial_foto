<?php include("functions.php"); ?>
<?php include("connection.php"); ?>
<?php

    $return = null;
    openConnection();
    switch($objeto)
    {
    
        /* Listagem de informações
         * Rotinas: SELECT
         * Retorno: JSON
         **/
        case "listar-movimentacoes":
            
            $sql = "SELECT A.cod_movimentacao,F.login as login, A.data, A.tipo, B.nome ,  
                        CAST((
                            SELECT GROUP_CONCAT(CONCAT(D.nome_produto,', Qtd: ', C.quantidade) separator '::')
                            FROM tb_est_itens_movimentacao C 
                            INNER JOIN TB_PRODUTOS D on D.COD_PRODUTO = C.cod_item 
                            WHERE A.cod_movimentacao = C.cod_movimentacao  ) AS CHAR(100000) CHARACTER SET Latin1 
                            ) as itens
                            , A.descricao
                  FROM tb_est_movimentacao A 
                  INNER JOIN tb_est_local_armazenamento B ON A.cod_local_armazenamento = B.cod_local_armazenamento 
                  INNER JOIN TB_ADM_USUARIOS F on F.COD_USUARIO = A.cod_usuario 
                  ORDER BY A.cod_movimentacao DESC";
                   
               
            $resultset = mysql_query($sql);
            while($linha = mysql_fetch_array($resultset,MYSQL_NUM)){
                foreach ($linha as $key => $value){
                        $linha[$key] = utf8_encode($value);
                        if($value == "IN")$linha[$key] =  "Entrada";
                        if($value == "OUT")$linha[$key] =  "Saída";                              
                }
                for($x = 1;$x < 100;$x++)$resultset_array[] = $linha;
                
            }
            $resultset_array = array_reverse($resultset_array);
            $resultset_array[] = array("Codigo","Responsavel","Data","Tipo","Em posse do","Itens","Descricao");
            $resultset_array = array_reverse($resultset_array);

            
            $return = json_encode($resultset_array);     
            break;

        default:
            $return = json_encode(array("error"=> "Objeto nao localizado."));
            break;
    
    }   

    echo $return;
?>