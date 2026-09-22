DROP TRIGGER IF EXISTS baixa_automatica_autenticacao;
DELIMITER $$    

CREATE TRIGGER baixa_automatica_autenticacao AFTER UPDATE ON TB_PEDIDOS FOR EACH ROW 
BEGIN
    DECLARE existBefore INT;
    SET @CODIGO_PEDIDO = new.cod_pedido;
    SELECT COUNT(A.id) INTO existBefore FROM tb_est_baixa_automatica_producao_log A WHERE A.cod_pedido = @CODIGO_PEDIDO AND A.tipo = 'venda';

    IF old.DATA_DESCEU_DIRETORIA <> new.DATA_DESCEU_DIRETORIA  AND existBefore = 0 THEN
        DROP TEMPORARY TABLE IF EXISTS TB_ProdutosByPedido;
        CREATE TEMPORARY TABLE IF NOT EXISTS TB_ProdutosByPedido(
            cod_item INT,
            cod_local INT,
            quantidade INT
        );
        SET @COD_FORMANDO := 0;
        SELECT COD_FORMANDO INTO @COD_FORMANDO FROM TB_PEDIDOS_FORMANDOS WHERE COD_PEDIDO = @CODIGO_PEDIDO LIMIT 1;
        INSERT INTO  TB_ProdutosByPedido (cod_item,cod_local,quantidade)  SELECT A.cod_item_e, A.cod_local,B.quantidade FROM tb_est_rel_item_venda A INNER JOIN TB_FICHAS_IDENTIFICACAO_BRINDES B ON B.COD_PRODUTO = A.cod_item_p WHERE B.COD_FORMANDO = @COD_FORMANDO AND VENDIDO = 1;
        INSERT INTO  TB_ProdutosByPedido (cod_item,cod_local,quantidade)  SELECT A.cod_item_e, A.cod_local,B.quantidade FROM tb_est_rel_item_venda A INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS B ON B.COD_PRODUTO = A.cod_item_p WHERE B.COD_FORMANDO = @COD_FORMANDO AND VENDIDO = 1;
        INSERT INTO  TB_ProdutosByPedido (cod_item,cod_local,quantidade)  SELECT A.cod_item_e, A.cod_local,B.quantidade FROM tb_est_rel_item_venda A INNER JOIN TB_PEDIDOS_PRODUTOS B ON B.COD_PRODUTO = A.cod_item_p WHERE B.COD_PEDIDO = @CODIGO_PEDIDO;
        CALL baixa_auto_proc (@CODIGO_PEDIDO,'venda');
    END IF;
   
END$$
DELIMITER ;


DROP TRIGGER baixa_automatica_producao;
/* create triggers */

DELIMITER $$
CREATE TRIGGER baixa_automatica_producao AFTER UPDATE ON TB_ORDENS_SERVICO FOR EACH ROW 
BEGIN
    DECLARE cod_pedido INT;
    DECLARE quantProd INT;
    DECLARE quantProdDesp INT;
    DECLARE existBefore INT;
    SET @CODIGO_PEDIDO = new.cod_pedido;
    
    SELECT COUNT(A.cod_ordem_servico) INTO quantProd FROM TB_ORDENS_SERVICO A WHERE A.cod_pedido = @CODIGO_PEDIDO;
    SELECT COUNT(A.cod_ordem_servico) INTO quantProdDesp FROM TB_ORDENS_SERVICO A WHERE A.cod_pedido = @CODIGO_PEDIDO AND A.status_ordem_servico = 2;
    SELECT COUNT(A.id) INTO existBefore FROM tb_est_baixa_automatica_producao_log A WHERE A.cod_pedido = @CODIGO_PEDIDO AND A.tipo = 'pedido';
    
    IF quantProd = quantProdDesp AND existBefore = 0 AND new.status_ordem_servico = 2 THEN      

        DROP TEMPORARY TABLE IF EXISTS TB_ProdutosByPedido;
        CREATE TEMPORARY TABLE IF NOT EXISTS TB_ProdutosByPedido(
            cod_item INT,
            cod_local INT,
            quantidade INT
        );
        SET @COD_FORMANDO := 0;
        SELECT COD_FORMANDO INTO @COD_FORMANDO FROM TB_PEDIDOS_FORMANDOS WHERE COD_PEDIDO = @CODIGO_PEDIDO LIMIT 1;
        INSERT INTO  TB_ProdutosByPedido (cod_item,cod_local,quantidade)  SELECT A.cod_item_e, A.cod_local,B.quantidade FROM tb_est_rel_item A INNER JOIN TB_FICHAS_IDENTIFICACAO_BRINDES B ON B.COD_PRODUTO = A.cod_item_p WHERE B.COD_FORMANDO = @COD_FORMANDO AND VENDIDO = 1;
        INSERT INTO  TB_ProdutosByPedido (cod_item,cod_local,quantidade)  SELECT A.cod_item_e, A.cod_local,B.quantidade FROM tb_est_rel_item A INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS B ON B.COD_PRODUTO = A.cod_item_p WHERE B.COD_FORMANDO = @COD_FORMANDO AND VENDIDO = 1;
        INSERT INTO  TB_ProdutosByPedido (cod_item,cod_local,quantidade)  SELECT A.cod_item_e, A.cod_local,B.quantidade FROM tb_est_rel_item A INNER JOIN TB_PEDIDOS_PRODUTOS B ON B.COD_PRODUTO = A.cod_item_p WHERE B.COD_PEDIDO = @CODIGO_PEDIDO;
        CALL baixa_auto_proc (@CODIGO_PEDIDO,'producao');
    END IF;
END$$
DELIMITER ;

DROP PROCEDURE baixa_auto_proc;

delimiter $$

CREATE PROCEDURE `baixa_auto_proc`(COD_PEDIDO INT,TIPO VARCHAR(255))
BEGIN
	DECLARE done INT DEFAULT FALSE;
	DECLARE cod_item,quantidade,cod_local,quantidade_item, quantidade_item_composto, cod_movimentacao, cod_item_pai,cod_item_filho INT;	
	DECLARE cur_itens CURSOR FOR SELECT A.cod_item, A.cod_local,A.quantidade FROM TB_ProdutosByPedido A;
	DECLARE cur_sub_itens CURSOR FOR SELECT A.cod_item_pai, A.cod_item , A.quantidade FROM tb_est_item_composto A;
	DECLARE CONTINUE HANDLER FOR NOT FOUND SET done = TRUE;
	SET @COD_USUARIO = (SELECT COD_USUARIO FROM TB_ADM_USUARIOS WHERE NOME = 'Sistema' AND SENHA = '-------' AND LOGIN = 'sistema' LIMIT 0,1);
    SET @COD_LOCAL_ANTIGO = 0;
	OPEN cur_itens;
	READ_LOOP: LOOP
		FETCH cur_itens INTO cod_item, cod_local, quantidade;
		IF done THEN
			LEAVE READ_LOOP;
		END IF;
        IF @COD_LOCAL_ANTIGO <> cod_local THEN
            INSERT INTO tb_est_movimentacao VALUES(0,'OUT',CONCAT('Baixa automatica - Tipo:',tipo,' :  - Pedido: ',cod_pedido),cod_local,NOW(),@COD_USUARIO);
            SET @COD_MOVIMENTACAO = LAST_INSERT_ID();
            SET @COD_LOCAL_ANTIGO = cod_local;
            INSERT INTO tb_est_baixa_automatica_producao_log (cod_pedido,cod_movimentacao,data,tipo) VALUES (COD_PEDIDO,@COD_MOVIMENTACAO, NOW(),TIPO);
        END IF;  

		IF (SELECT A.tipo FROM tb_est_itens A WHERE A.cod_item = cod_item) = 'S' THEN
			BEGIN
				INSERT INTO tb_est_itens_movimentacao VALUES (0,@COD_MOVIMENTACAO,cod_item,quantidade,NULL);
			END;
		ELSE
			BEGIN

				OPEN cur_sub_itens;
				READ_LOOP2: LOOP
					FETCH cur_sub_itens INTO cod_item_pai, cod_item_filho, quantidade_item;
					IF done THEN
						LEAVE READ_LOOP;
					END IF;
					IF cod_item_pai = cod_item THEN
						INSERT INTO tb_est_itens_movimentacao VALUES (0,@COD_MOVIMENTACAO,cod_item_filho,(quantidade_item*quantidade),NULL);
					END IF;
				END LOOP READ_LOOP2;
				CLOSE cur_sub_itens;

			END;
		END IF;

	END LOOP READ_LOOP;
	CLOSE cur_itens;
END$$




DELIMITER ;




-- Test's Scripts
-- SELECT * FROM TB_ORDENS_SERVICO WHERE COD_PEDIDO = '30867';
-- SELECT * FROM TB_PEDIDOS WHERE COD_PEDIDO = 30867 ORDER BY COD_PEDIDO DESC;
-- SELECT COD_FORMANDO FROM TB_PEDIDOS_FORMANDOS WHERE COD_PEDIDO = 30867 LIMIT 1;
-- UPDATE `TB_PEDIDOS` SET DATA_DESCEU_DIRETORIA='2014-04-11' WHERE COD_PEDIDO=30867 AND SISTEMA_NOVO = 1 LIMIT 1 ;
-- UPDATE `TB_PEDIDOS` SET `SISTEMA_NOVO`= 1 WHERE COD_PEDIDO=30867;
-- SELECT * FROM tb_est_rel_item_venda;

-- SELECT * FROM tb_est_baixa_automatica_producao_log;
-- SELECT * FROM testeLog;
-- SELECT * FROM tb_est_rel_item_venda;
-- SELECT * FROM tb_est_movimentacao;
-- SELECT * FROM TB_ORDENS_SERVICO WHERE COD_PEDIDO = '30867';
