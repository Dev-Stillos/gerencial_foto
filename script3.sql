CREATE DATABASE learn;
CREATE TABLE t(
	cod_t int auto_increment not null,
	name varchar(255),
	primary key(cod_t)
);

INSERT INTO t VALUES (0,'Luiz esta ');
UPDATE t SET name = 'Luiz' WHERE cod_t = 1;
SELECT * FROM t;

DROP TRIGGER baixa_auto_tgr;
CREATE TRIGGER baixa_auto_tgr BEFORE UPDATE ON t 
FOR EACH ROW
CALL baixa_auto_proc(NEW.cod_t);



DELIMITER $$
DROP PROCEDURE IF EXISTS baixa_auto_proc$$
CREATE PROCEDURE baixa_auto_proc (cod_item_p INT, quantidade INT)
BEGIN
	DECLARE done INT DEFAULT FALSE;
	DECLARE cod_item, cod_local,quantidade_item, quantidade_item_composto, cod_movimentacao INT;	
	DECLARE cur_itens cursor for select A.cod_item_e, A.cod_local from tb_est_rel_item A WHERE A.cod_item_p = cod_item_p;
	DECLARE cur_sub_itens cursor for select A.cod_item FROM tb_est_item_composto;
	DECLARE CONTINUE HANDLER FOR NOT FOUND SET done = TRUE;

	open cur_itens;
	read_loop: LOOP
		fetch cur_itens into cod_item, cod_local;
		if done then
			leave read_loop;
		end if;
		INSERT INTO tb_est_movimentacao VALUES(0,'OUT','baixa automatica',cod_local,now(),0);
        SET cod_movimentacao = LAST_INSERT_ID();
           
		IF (SELECT A.tipo FROM tb_est_itens A WHERE A.cod_item = cod_item) = 'S' THEN
			BEGIN





			END;
		ELSE
			BEGIN






			END;


		SET quantidade_in =  (SELECT sum(A.quantidade) FROM tb_est_itens_movimentacao A INNER JOIN tb_est_movimentacao B ON A.cod_movimentacao = B.cod_movimentacao WHERE B.tipo = 'IN'  and A.cod_item = cod and B.cod_local_armazenamento = cod_local) + 0;
		SET quantidade_out = (SELECT sum(A.quantidade) FROM tb_est_itens_movimentacao A INNER JOIN tb_est_movimentacao B ON A.cod_movimentacao = B.cod_movimentacao WHERE B.tipo = 'OUT' and A.cod_item = cod and B.cod_local_armazenamento = cod_local) + 0;

		if quantidade_in is null then set quantidade_in = 0; end if;
		if quantidade_out is null then set quantidade_out = 0; end if;
		set quantidade_sum := quantidade_in - quantidade_out;
		INSERT INTO tb_est_resumo VALUES (0,cod,cod_local,quantidade_sum,now());

	end loop read_loop;

	close cur_itens;

end$$
DELIMITER ;



