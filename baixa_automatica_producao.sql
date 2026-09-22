
DROP TABLE tb_est_baixa_automatica_producao_log;
CREATE TABLE tb_est_baixa_automatica_producao_log(
 id int auto_increment,
 cod_pedido INT,
 cod_movimentacao INT,
 data DATETIME,
 primary key(id)
);

DROP TRIGGER baixa_automatica_producao;
DELIMITER $$
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


DROP PROCEDURE baixa_auto_producao_proc;
DELIMITER $$
CREATE PROCEDURE `baixa_auto_producao_proc`(cod_pedido INT,tipo varchar(255))
BEGIN
	DECLARE done INT DEFAULT FALSE;
	DECLARE cod_item,quantidade,cod_local,quantidade_item, quantidade_item_composto, cod_movimentacao, cod_item_pai,cod_item_filho INT;	
	DECLARE cur_itens cursor for select A.cod_item_e, A.cod_local,B.quantidade from tb_est_rel_item A INNER JOIN TB_ORDENS_SERVICO B ON B.cod_produto = A.cod_item_p WHERE B.cod_pedido = @cod_pedido;
	DECLARE cur_sub_itens cursor for select A.cod_item_pai, A.cod_item , A.quantidade FROM tb_est_item_composto A;
	DECLARE CONTINUE HANDLER FOR NOT FOUND SET done = TRUE;
	SET @cod_usuario = (SELECT COD_USUARIO FROM TB_ADM_USUARIOS WHERE NOME = 'Sistema' AND SENHA = '-------' AND LOGIN = 'sistema' LIMIT 0,1);
    SET @COD_LOCAL_ANTIGO = 0;

	open cur_itens;
	read_loop: LOOP
		fetch cur_itens into cod_item, cod_local, quantidade;
		if done then
			leave read_loop;
		end if;
        if @COD_LOCAL_ANTIGO <> cod_local then
            INSERT INTO tb_est_movimentacao VALUES(0,'OUT',CONCAT('Baixa automatica -',tipo,' : ',cod_pedido),cod_local,now(),@cod_usuario);
            SET @cod_movimentacao = LAST_INSERT_ID();
            SET @COD_LOCAL_ANTIGO = cod_local;
            INSERT INTO tb_est_baixa_automatica_producao_log (cod_pedido,cod_movimentacao,data) VALUES (@cod_pedido,@cod_movimentacao, now());
        end if;  

		IF (SELECT A.tipo FROM tb_est_itens A WHERE A.cod_item = cod_item) = 'S' THEN
			BEGIN
				INSERT INTO tb_est_itens_movimentacao VALUES (0,@cod_movimentacao,cod_item,quantidade,NULL);
			END;
		ELSE
			BEGIN

				open cur_sub_itens;
				read_loop2: LOOP
					fetch cur_sub_itens into cod_item_pai, cod_item_filho, quantidade_item;
					if done then
						leave read_loop;
					end if;
					IF cod_item_pai = cod_item THEN
						INSERT INTO tb_est_itens_movimentacao VALUES (0,@cod_movimentacao,cod_item_filho,(quantidade_item*quantidade),NULL);
					END IF;
				end loop read_loop2;
				close cur_sub_itens;

			END;
		END IF;

	end loop read_loop;
	close cur_itens;
end