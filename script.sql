USE stillos_local;
/* 
####################################
Preparando Ambiente
####################################

SELECT * FROM tb_adm_modulos;
INSERT INTO tb_adm_modulos VALUES(8,8,'PTI - Planejamento TI','PTI','Módulo destinado a divulgação da timeline de t=projetos da TI');
INSERT INTO tb_adm_modulos VALUES(9,9,'EST - Estoque (Pré-Venda)','EST','Módulo destinado ao controle de estoque interno.');
UPDATE tb_am_modulos SET NOME_MODULO = 'PLE - Planejamento estratégico', SIGLA_MODULO = 'PLE', DESCRICAO = 'Módulo destinado ao planejamento de atividades em timeline' WHERE COD_MODULO = 8;
SELECT * FROM tb_adm_paginas_habilitadas WHERE COD_PAGINA = 280;
UPDATE tb_adm_paginas_habilitadas SET COD_MODULO = 9 WHERE COD_PAGINA = 280;
UPDATE tb_adm_paginas_habilitadas  SET COD_MODULO = 8, CAMINHO_VALIDA = 'ple_projetos.asp' WHERE COD_PAGINA = 279;
SELECT * FROM
SELECT * FROM TB_ADM_PERFIS_PAGINAS_HABILITADAS WHERE COD_PERFIL

*/

/* Criando as tabelas do Módulo EST */
DROP TABLE tb_est_movimento_pedido;
DROP TABLE tb_est_rel_item;
DROP TABLE tb_est_resumo;
DROP TABLE tb_est_itens_pedidos;
DROP TABLE tb_est_itens_movimentacao;
DROP TABLE tb_est_item_composto;

DROP TABLE tb_est_movimentacao;
DROP TABLE tb_est_pedidos;
DROP TABLE tb_est_fornecedores;

DROP TABLE tb_est_itens;
DROP TABLE tb_est_categoria;
DROP TABLE tb_est_local_armazenamento;

DROP TABLE tb_ple_projetos;

DELETE FROM TB_ADM_USUARIOS WHERE NOME = 'Sistema' AND SENHA = '-------' AND LOGIN = 'sistema';

CREATE TABLE tb_est_categoria(
	cod_categoria int auto_increment not null,
	nome varchar(255) null,
	descricao varchar(4000) null,
	data datetime not null,
	primary key (cod_categoria)
);
INSERT INTO tb_est_categoria VALUES(0,'padrão', '' ,now());
SELECT * FROM tb_est_categoria;

CREATE TABLE tb_est_local_armazenamento(
	cod_local_armazenamento int auto_increment not null,
	nome varchar(255) null,
	descricao varchar(400) null,
	data datetime not null,
	primary key(cod_local_armazenamento)
);
INSERT INTO tb_est_local_armazenamento VALUES(0,'Pedidos','',now());
SELECT * FROM tb_est_local_armazenamento;

CREATE TABLE tb_est_itens(
	cod_item int auto_increment not null,
	nome varchar(255) null,
	cod_categoria int not null,
	descricao varchar(4000) null,
	tipo varchar(2) not null,
	data datetime not null,	
	primary key (cod_item),
	foreign key (cod_categoria)  references tb_est_categoria(cod_categoria)
);
SELECT A.cod_item, A.nome as nome_item, B.nome as nome_categoria FROM tb_est_itens A INNER JOIN tb_est_categoria B on A.cod_categoria = B.cod_categoria;

CREATE TABLE tb_est_item_composto(
	cod_item_composto int auto_increment not null,
	cod_item_pai int not null,
	cod_item int not null,
	quantidade int not null,
	primary key (cod_item_composto),
	foreign key (cod_item) references tb_est_itens(cod_item),
	foreign key (cod_item_pai) references tb_est_itens(cod_item)
);
SELECT * FROM tb_est_item_composto;

CREATE TABLE tb_est_movimentacao(
	cod_movimentacao int auto_increment not null,
	tipo varchar(3) not null,
	descricao varchar(4000) null,
	cod_local_armazenamento int not null,
	data datetime not null,
	cod_usuario int not null,
	primary key (cod_movimentacao),
	foreign key (cod_local_armazenamento) references tb_est_local_armazenamento (cod_local_armazenamento)
);
SELECT * FROM tb_est_movimentacao;

CREATE TABLE tb_est_itens_movimentacao(
	cod_itens_movimentacao int auto_increment not null,
	cod_movimentacao int not null,
	cod_item int not null,
	quantidade int not null,
	cod_item_pai int null,
	primary key (cod_itens_movimentacao),
	foreign key (cod_item) references tb_est_itens (cod_item),
	foreign key (cod_movimentacao) references tb_est_movimentacao (cod_movimentacao)
);
SELECT * FROM tb_est_itens_movimentacao;

CREATE TABLE tb_est_resumo(
	cod_resumo int auto_increment not null,
	cod_item int not null,
	cod_local_armazenamento int not null,
	quantia int not null,
	data datetime not null,
	foreign key (cod_item) references tb_est_itens(cod_item),
	foreign key (cod_local_armazenamento) references tb_est_local_armazenamento(cod_local_armazenamento),
	primary key(cod_resumo)
);
SELECT * FROM tb_est_resumo;

CREATE TABLE tb_est_fornecedores(
	cod_fornecedor int auto_increment not null,
	nome varchar(255) null,
	descricao VARCHAR(4000) null,
	primary key(cod_fornecedor)
);
SELECT * FROM tb_est_fornecedores;

CREATE TABLE tb_est_pedidos(
	cod_pedido int auto_increment not null,
	cod_fornecedor int not null,
	cod_usuario int not null,
	status varchar(3) not null,
	data_pedido date not null,
	data_previsao date,
	data_entrega date,
	descricao varchar(4000),
	primary key(cod_pedido),
	foreign key (cod_fornecedor) references tb_est_fornecedores(cod_fornecedor)
);
SELECT * FROM tb_est_pedidos;

CREATE TABLE tb_est_itens_pedidos(
	cod_itens_pedidos int auto_increment not null,
	cod_pedido int not null,
	cod_item int not null,
	quantidade int not null,
	valor float not null,
	primary key(cod_itens_pedidos),
	foreign key (cod_pedido) references tb_est_pedidos(cod_pedido),
	foreign key (cod_item) references tb_est_itens(cod_item)
);
SELECT * FROM tb_est_itens_pedidos;
	


CREATE TABLE tb_est_rel_item(
	cod_rel_item int AUTO_INCREMENT not null,
	cod_item_p int not null,
	cod_item_e int not null,
	cod_local int not null,
	descricao varchar(4000),
	primary key (cod_rel_item),
	foreign key (cod_item_e) references tb_est_itens(cod_item),
	foreign key (cod_local) references tb_est_local_armazenamento(cod_local_armazenamento)
);
SELECT A.cod_rel_item, B.nome, C.NOME_PRODUTO, D.nome as local FROM tb_est_rel_item A INNER JOIN tb_est_itens B ON A.cod_item_e = B.cod_item INNER JOIN TB_PRODUTOS C ON C.COD_PRODUTO = A.cod_item_p INNER JOIN tb_est_local_armazenamento D ON D.cod_local_armazenamento = A.cod_local;

CREATE TABLE tb_est_movimento_pedido(
	cod_movimento_pedido int AUTO_INCREMENT not null,
	cod_movimentacao int not null,
	cod_pedido int not null,
	primary key(cod_movimento_pedido),
	foreign key (cod_movimentacao) references tb_est_movimentacao(cod_movimentacao),
	foreign key (cod_pedido) references tb_est_pedidos (cod_pedido)
);
SELECT * FROM tb_est_movimento_pedido;

CREATE TABLE tb_ple_projetos(
	cod_projeto int auto_increment not null,
	nome varchar(255),
	descricao varchar(4000),
	inicio date,
	termino date,
	primary key(cod_projeto)
);
SELECT * FROM  tb_ple_projetos;

INSERT INTO `tb_ple_projetos` (`cod_projeto`,`nome`,`descricao`,`inicio`,`termino`) VALUES (1,'Solicitações','','2014-03-10','2015-06-30');
INSERT INTO `tb_ple_projetos` (`cod_projeto`,`nome`,`descricao`,`inicio`,`termino`) VALUES (2,'Controle de Estoque','','2014-03-17','2014-04-28');
INSERT INTO `tb_ple_projetos` (`cod_projeto`,`nome`,`descricao`,`inicio`,`termino`) VALUES (3,'Gestão de contratos','','2014-03-31','2014-05-10');



/* Procedures */
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

/*
call atualizarEstoque();
SELECT * FROM tb_est_resumo;
SELECT * FROM TB_ADM_USUARIOS;
SELECT * FROM tb_est_rel_item;
SELECT * FROM tb_est_item_composto;
SELECT * FROM tb_est_movimentacao;
CALL baixa_auto_proc (3,1);
*/


/*
Criação do Trigger para baixa automatica

DROP TRIGGER baixa_auto_tgr;
CREATE TRIGGER baixa_auto_tgr BEFORE UPDATE ON t 
FOR EACH ROW
IF NEW.STATUS = 3 THEN
	CALL baixa_auto_proc(NEW.cod_t);
END IF;

*/


INSERT INTO `stillos_local`.`TB_ADM_USUARIOS` (`COD_USUARIO`, `NOME`, `LOGIN`, `SENHA`, `COD_DEPTO`, `BLOQUEADO`, `COD_PERFIL`, `PERMITIR_USO_EXTERNO`) VALUES (0, 'Sistema', 'sistema', '-------', '1', '1', '0', '2');

DELIMITER $$
DROP PROCEDURE IF EXISTS baixa_auto_proc$$
CREATE PROCEDURE baixa_auto_proc (cod_item_p INT, quantidade INT)
BEGIN
	DECLARE done INT DEFAULT FALSE;
	DECLARE cod_item, cod_local,quantidade_item, quantidade_item_composto, cod_movimentacao, cod_item_pai,cod_item_filho, cod_usuario INT;	
	DECLARE cur_itens cursor for select A.cod_item_e, A.cod_local from tb_est_rel_item A WHERE A.cod_item_p = cod_item_p;
	DECLARE cur_sub_itens cursor for select A.cod_item_pai, A.cod_item , A.quantidade FROM tb_est_item_composto A;
	DECLARE CONTINUE HANDLER FOR NOT FOUND SET done = TRUE;
	SET cod_usuario = (SELECT COD_USUARIO FROM TB_ADM_USUARIOS WHERE NOME = 'Sistema' AND SENHA = '-------' AND LOGIN = 'sistema');
	open cur_itens;
	read_loop: LOOP
		fetch cur_itens into cod_item, cod_local;
		if done then
			leave read_loop;
		end if;
		INSERT INTO tb_est_movimentacao VALUES(0,'OUT','baixa automatica',cod_local,now(),cod_usuario);
        SET cod_movimentacao = LAST_INSERT_ID();
           
		IF (SELECT A.tipo FROM tb_est_itens A WHERE A.cod_item = cod_item) = 'S' THEN
			BEGIN
				INSERT INTO tb_est_itens_movimentacao VALUES (0,cod_movimentacao,cod_item,quantidade,NULL);
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
						INSERT INTO tb_est_itens_movimentacao VALUES (0,cod_movimentacao,cod_item_filho,(quantidade_item*quantidade),NULL);
					END IF;
				end loop read_loop2;
				close cur_sub_itens;

			END;
		END IF;

	end loop read_loop;

	close cur_itens;

end$$
DELIMITER ;

/* Calculo - Estoque 
SELECT 
DISTINCT A.cod_item, sum(A.quantidade) 
FROM tb_est_itens_movimentacao A 
INNER JOIN tb_est_itens B ON A.cod_item = B.cod_item 
INNER JOIN tb_est_movimentacao C ON A.cod_movimentacao = C.cod_movimentacao  
WHERE C.tipo = "IN";
*/