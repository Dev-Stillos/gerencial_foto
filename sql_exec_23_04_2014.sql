delimiter $$

CREATE TABLE `tb_est_baixa_automatica_producao_log` (
  `id` INT(11) NOT NULL AUTO_INCREMENT,
  `cod_pedido` INT(11) DEFAULT NULL,
  `cod_movimentacao` INT(11) DEFAULT NULL,
  `data` DATETIME DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=INNODB DEFAULT CHARSET=LATIN1$$



delimiter $$

CREATE TABLE `tb_est_rel_item_venda` (
  `cod_rel_item_venda` INT(11) NOT NULL AUTO_INCREMENT,
  `cod_item_p` INT(11) NOT NULL,
  `cod_item_e` INT(11) NOT NULL,
  `descricao` VARCHAR(4000) DEFAULT NULL,
  PRIMARY KEY (`cod_rel_item_venda`)
) ENGINE=INNODB DEFAULT CHARSET=LATIN1$$


delimiter $$

CREATE TABLE `tb_est_rel_vendedor_local` (
  `cod_rel_vendedor_local` INT(11) NOT NULL AUTO_INCREMENT,
  `cod_local_armazenamento` INT(11) DEFAULT NULL,
  `cod_vendedor` INT(11) NOT NULL,
  `nome_vendedor` VARCHAR(255) DEFAULT NULL,
  PRIMARY KEY (`cod_rel_vendedor_local`)
) ENGINE=INNODB DEFAULT CHARSET=LATIN1$$


delimiter $$

CREATE PROCEDURE `baixa_auto_producao_proc`(cod_pedido INT,tipo VARCHAR(255))
BEGIN
	DECLARE done INT DEFAULT FALSE;
	DECLARE cod_item,quantidade,cod_local,quantidade_item, quantidade_item_composto, cod_movimentacao, cod_item_pai,cod_item_filho INT;	
	DECLARE cur_itens CURSOR FOR SELECT A.cod_item_e, A.cod_local,B.quantidade FROM tb_est_rel_item A INNER JOIN TB_ORDENS_SERVICO B ON B.cod_produto = A.cod_item_p WHERE B.cod_pedido = @COD_PEDIDO;
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
            INSERT INTO tb_est_movimentacao VALUES(0,'OUT',CONCAT('Baixa automatica -',tipo,' : ',cod_pedido),cod_local,NOW(),@COD_USUARIO);
            SET @COD_MOVIMENTACAO = LAST_INSERT_ID();
            SET @COD_LOCAL_ANTIGO = cod_local;
            INSERT INTO tb_est_baixa_automatica_producao_log (cod_pedido,cod_movimentacao,data) VALUES (@COD_PEDIDO,@COD_MOVIMENTACAO, NOW());
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



delimiter $$

CREATE TRIGGER baixa_automatica_producao AFTER UPDATE ON TB_ORDENS_SERVICO FOR EACH ROW 
BEGIN
    DECLARE cod_pedido INT;
    SET @COD_PEDIDO = new.cod_pedido;

    IF (SELECT COUNT(A.cod_ordem_servico) FROM TB_ORDENS_SERVICO A WHERE A.cod_pedido = cod_pedido) 
        =
       (SELECT COUNT(A.cod_ordem_servico) FROM TB_ORDENS_SERVICO A WHERE A.cod_pedido = cod_pedido AND status_ordem_servico = 2)
        AND
       (SELECT COUNT(A.id) FROM tb_est_baixa_automatica_producao_log A WHERE A.cod_pedido = @COD_PEDIDO) = 0
    THEN        
        CALL baixa_auto_producao_proc (@COD_PEDIDO,'pedido');
    END IF;
END$$
DELIMITER ;