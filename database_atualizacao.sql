/*Removendo Tabelas */
DROP TABLE `tb_est_baixa_automatica_producao_log`;
DROP TABLE `tb_est_categoria`;
DROP TABLE `tb_est_fornecedores`;
DROP TABLE `tb_est_item_composto`;
DROP TABLE `tb_est_itens`;
DROP TABLE `tb_est_itens_movimentacao`;
DROP TABLE `tb_est_itens_pedidos`;
DROP TABLE `tb_est_local_armazenamento`;
DROP TABLE `tb_est_movimentacao`;
DROP TABLE `tb_est_movimento_pedido`;
DROP TABLE `tb_est_pedidos`;
DROP TABLE `tb_est_rel_item`;
DROP TABLE `tb_est_rel_item_venda`;
DROP TABLE `tb_est_rel_vendedor_local`;
DROP TABLE `tb_est_resumo`;
DROP TABLE `tb_ple_projetos`;
DROP TABLE `tb_est_log_acao`;


/*Criando Tabelas */


CREATE TABLE tb_est_log_acao(
    id INT AUTO_INCREMENT,
    cod_usuario INT,
    nome_usuario VARCHAR(255),
    nome_departamento VARCHAR(255),
    form VARCHAR(255),
    acao INT,
    descricao VARCHAR(4000),
    data date,
    primary key(id)
);


CREATE TABLE `tb_est_baixa_automatica_producao_log` (
  `id` INT(11) NOT NULL AUTO_INCREMENT,
  `cod_pedido` INT(11) DEFAULT NULL,
  `cod_movimentacao` INT(11) DEFAULT NULL,
  `data` DATETIME DEFAULT NULL,
  `tipo` varchar(255) NOT NULL,
  PRIMARY KEY (`id`)
);

CREATE TABLE `tb_est_categoria` (
  `cod_categoria` INT(11) NOT NULL AUTO_INCREMENT,
  `nome` VARCHAR(255) DEFAULT NULL,
  `descricao` VARCHAR(4000) DEFAULT NULL,
  `data` DATETIME NOT NULL,
  PRIMARY KEY (`cod_categoria`)
);
INSERT INTO tb_est_categoria VALUES (0,'padrão','',now()),(0,'virtual','',now());
CREATE TABLE `tb_est_fornecedores` (
  `cod_fornecedor` INT(11) NOT NULL AUTO_INCREMENT,
  `nome` VARCHAR(255) DEFAULT NULL,
  `descricao` VARCHAR(4000) DEFAULT NULL,
  PRIMARY KEY (`cod_fornecedor`)
);

CREATE TABLE `tb_est_item_composto` (
  `cod_item_composto` INT(11) NOT NULL AUTO_INCREMENT,
  `cod_item_pai` INT(11) NOT NULL,
  `cod_item` INT(11) NOT NULL,
  `quantidade` INT(11) NOT NULL,
  PRIMARY KEY (`cod_item_composto`),
  KEY `cod_item` (`cod_item`),
  KEY `cod_item_pai` (`cod_item_pai`)
);

CREATE TABLE `tb_est_itens` (
  `cod_item` INT(11) NOT NULL AUTO_INCREMENT,
  `nome` VARCHAR(255) DEFAULT NULL,
  `cod_categoria` INT(11) NOT NULL,
  `descricao` VARCHAR(4000) DEFAULT NULL,
  `tipo` VARCHAR(2) NOT NULL,
  `data` DATETIME NOT NULL,
  PRIMARY KEY (`cod_item`),
  KEY `cod_categoria` (`cod_categoria`)
);

CREATE TABLE `tb_est_itens_movimentacao` (
  `cod_itens_movimentacao` INT(11) NOT NULL AUTO_INCREMENT,
  `cod_movimentacao` INT(11) NOT NULL,
  `cod_item` INT(11) NOT NULL,
  `quantidade` INT(11) NOT NULL,
  `cod_item_pai` INT(11) DEFAULT NULL,
  PRIMARY KEY (`cod_itens_movimentacao`),
  KEY `cod_item` (`cod_item`),
  KEY `cod_movimentacao` (`cod_movimentacao`)
);

CREATE TABLE `tb_est_itens_pedidos` (
  `cod_itens_pedidos` INT(11) NOT NULL AUTO_INCREMENT,
  `cod_pedido` INT(11) NOT NULL,
  `cod_item` INT(11) NOT NULL,
  `quantidade` INT(11) NOT NULL,
  `valor` FLOAT NOT NULL,
  PRIMARY KEY (`cod_itens_pedidos`),
  KEY `cod_pedido` (`cod_pedido`),
  KEY `cod_item` (`cod_item`)
);

CREATE TABLE `tb_est_local_armazenamento` (
  `cod_local_armazenamento` INT(11) UNSIGNED NOT NULL AUTO_INCREMENT,
  `nome` VARCHAR(255) DEFAULT NULL,
  `descricao` VARCHAR(400) DEFAULT NULL,
  `data` DATETIME NOT NULL,
  PRIMARY KEY (`cod_local_armazenamento`)
);
INSERT INTO tb_est_local_armazenamento VALUES (0,'pedido','',now());


CREATE TABLE `tb_est_movimentacao` (
  `cod_movimentacao` INT(11) NOT NULL AUTO_INCREMENT,
  `tipo` VARCHAR(3) NOT NULL,
  `descricao` VARCHAR(4000) DEFAULT NULL,
  `cod_local_armazenamento` INT(11) NOT NULL,
  `data` DATETIME NOT NULL,
  `cod_usuario` INT(11) NOT NULL,
  PRIMARY KEY (`cod_movimentacao`),
  KEY `cod_local_armazenamento` (`cod_local_armazenamento`)
);

CREATE TABLE `tb_est_movimento_pedido` (
  `cod_movimento_pedido` INT(11) NOT NULL AUTO_INCREMENT,
  `cod_movimentacao` INT(11) NOT NULL,
  `cod_pedido` INT(11) NOT NULL,
  PRIMARY KEY (`cod_movimento_pedido`),
  KEY `cod_movimentacao` (`cod_movimentacao`),
  KEY `cod_pedido` (`cod_pedido`)
);

CREATE TABLE `tb_est_pedidos` (
  `cod_pedido` INT(11) NOT NULL AUTO_INCREMENT,
  `cod_fornecedor` INT(11) NOT NULL,
  `cod_usuario` INT(11) NOT NULL,
  `status` VARCHAR(3) NOT NULL,
  `data_pedido` DATE NOT NULL,
  `data_previsao` DATE DEFAULT NULL,
  `data_entrega` DATE DEFAULT NULL,
  `descricao` VARCHAR(4000) DEFAULT NULL,
  PRIMARY KEY (`cod_pedido`),
  KEY `cod_fornecedor` (`cod_fornecedor`)
);

CREATE TABLE `tb_est_rel_item` (
  `cod_rel_item` INT(11) NOT NULL AUTO_INCREMENT,
  `cod_item_p` INT(11) NOT NULL,
  `cod_item_e` INT(11) NOT NULL,
  `cod_local` INT(11) NOT NULL,
  `descricao` VARCHAR(4000) DEFAULT NULL,
  PRIMARY KEY (`cod_rel_item`),
  KEY `cod_item_e` (`cod_item_e`),
  KEY `cod_local` (`cod_local`)
);

CREATE TABLE `tb_est_rel_item_venda` (
  `cod_rel_item_venda` INT(11) NOT NULL AUTO_INCREMENT,
  `cod_item_p` INT(11) NOT NULL,
  `cod_item_e` INT(11) NOT NULL,
  `cod_local` INT(11) NOT NULL,
  `descricao` VARCHAR(4000) DEFAULT NULL,
  PRIMARY KEY (`cod_rel_item_venda`),
  KEY cod_local_armazenamento (`cod_local`)
);

CREATE TABLE `tb_est_rel_vendedor_local` (
  `cod_rel_vendedor_local` INT(11) NOT NULL AUTO_INCREMENT,
  `cod_local_armazenamento` INT(11) DEFAULT NULL,
  `cod_vendedor` INT(11) NOT NULL,
  `nome_vendedor` VARCHAR(255) DEFAULT NULL,
  PRIMARY KEY (`cod_rel_vendedor_local`)
);

CREATE TABLE `tb_est_resumo` (
  `cod_resumo` INT(11) NOT NULL AUTO_INCREMENT,
  `cod_item` INT(11) NOT NULL,
  `cod_local_armazenamento` INT(11) NOT NULL,
  `quantia` INT(11) NOT NULL,
  `data` DATETIME NOT NULL,
  PRIMARY KEY (`cod_resumo`),
  KEY `cod_item` (`cod_item`),
  KEY `cod_local_armazenamento` (`cod_local_armazenamento`)
);

CREATE TABLE `tb_ple_projetos` (
  `cod_projeto` INT(11) NOT NULL AUTO_INCREMENT,
  `nome` VARCHAR(255) DEFAULT NULL,
  `descricao` VARCHAR(4000) DEFAULT NULL,
  `inicio` DATE DEFAULT NULL,
  `termino` DATE DEFAULT NULL,
  PRIMARY KEY (`cod_projeto`)
);
INSERT INTO tb_ple_projetos VALUES (0,'Solicitações','','2014-03-10','2015-06-30'), (0,'Controle de Estoque','','2014-03-17','2014-04-28'), (0,'Gestão de contratos','','2014-03-31','2014-05-10');

/* Drop Procedures */

DROP PROCEDURE IF EXISTS atualizarEstoque;

/* Create Procedures */
delimiter $$

CREATE PROCEDURE `atualizarEstoque`()
BEGIN
	DECLARE done INT DEFAULT FALSE;

	DECLARE cod , cod_local INT; 
	DECLARE nome VARCHAR(255);	
	DECLARE quantidade_out, quantidade_in, quantidade_sum INT;
	DECLARE cur_itens CURSOR FOR SELECT DISTINCT A.cod_item, C.cod_local_armazenamento FROM tb_est_itens A INNER JOIN tb_est_itens_movimentacao B ON A.cod_item = B.cod_item INNER JOIN tb_est_movimentacao C ON B.cod_movimentacao = C.cod_movimentacao;

	DECLARE CONTINUE HANDLER FOR NOT FOUND SET done = TRUE;

	DELETE FROM tb_est_resumo WHERE cod_resumo >= 0;
	OPEN cur_itens;
	READ_LOOP: LOOP
		FETCH cur_itens INTO cod, cod_local;
		IF done THEN
			LEAVE READ_LOOP;
		END IF;
		SET quantidade_in =  (SELECT SUM(A.quantidade) FROM tb_est_itens_movimentacao A INNER JOIN tb_est_movimentacao B ON A.cod_movimentacao = B.cod_movimentacao WHERE B.tipo = 'IN'  AND A.cod_item = cod AND B.cod_local_armazenamento = cod_local) + 0;
		SET quantidade_out = (SELECT SUM(A.quantidade) FROM tb_est_itens_movimentacao A INNER JOIN tb_est_movimentacao B ON A.cod_movimentacao = B.cod_movimentacao WHERE B.tipo = 'OUT' AND A.cod_item = cod AND B.cod_local_armazenamento = cod_local) + 0;

		IF quantidade_in IS NULL THEN SET quantidade_in = 0; END IF;
		IF quantidade_out IS NULL THEN SET quantidade_out = 0; END IF;
		SET quantidade_sum := quantidade_in - quantidade_out;
        IF quantidade_sum <> 0 THEN
            INSERT INTO tb_est_resumo VALUES (NULL,cod,cod_local,quantidade_sum,NOW());
        END IF;
	END LOOP READ_LOOP;

	CLOSE cur_itens;
END$$






