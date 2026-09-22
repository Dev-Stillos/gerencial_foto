DELIMITER $$

DROP PROCEDURE IF EXISTS atualizarEstoque$$

CREATE PROCEDURE atualizarEstoque()
begin
	DECLARE done INT DEFAULT FALSE;

	declare cod , cod_local int; 
	declare nome varchar(255);	
	declare quantidade_out, quantidade_in, quantidade_sum int;
	declare cur_itens cursor for select DISTINCT A.cod_item, C.cod_local_armazenamento from tb_est_itens A INNER JOIN tb_est_itens_movimentacao B on A.cod_item = B.cod_item inner join tb_est_movimentacao C on B.cod_movimentacao = C.cod_movimentacao;

	DECLARE CONTINUE HANDLER FOR NOT FOUND SET done = TRUE;

	DELETE FROM tb_est_resumo where cod_resumo <> 0;

	open cur_itens;
	read_loop: LOOP
		fetch cur_itens into cod, cod_local;
		if done then
			leave read_loop;
		end if;
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

call atualizarEstoque();
SELECT * FROM tb_est_resumo;