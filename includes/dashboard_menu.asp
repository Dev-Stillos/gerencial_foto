<%
AbrirConexao()
Dim MENU_1, MENU_2, MENU_3, MENU_4, MENU_5, MENU_6, MENU_7, MENU_8, MENU_9, MENU_10, MENU_11, MENU_12, MENU_13, MENU_14, MENU_15, MENU_16, MENU_17, MENU_18, MENU_19, MENU_20, MENU_21, MENU_22, MENU_23, MENU_24 
MENU_1	= 2
MENU_2	= 2
MENU_3	= 2
MENU_4	= 2
MENU_5	= 2
MENU_6	= 2
MENU_7	= 2
MENU_8	= 2
MENU_9	= 2
MENU_10	= 2
MENU_11 = 2
MENU_12 = 2
MENU_13	= 2
MENU_14 = 2
MENU_15 = 2
MENU_16 = 2
MENU_17 = 2
MENU_18 = 2
MENU_19 = 2
MENU_20 = 2
MENU_21 = 2
MENU_22 = 2
MENU_23 = 2
MENU_24 = 2
MENU_25 = 2


Dim rsVerificadorMenu,COD_PAGINA_VALIDADA_MENU_VALIDA
'SQL_CACHE
Set rsVerificadorMenu = objConexao.Execute("SELECT * FROM TB_ADM_PERFIS_PAGINAS_HABILITADAS WHERE COD_PERFIL="&Session("COD_PERFIL"))

		While not rsVerificadorMenu.Eof

			COD_PAGINA_VALIDADA_MENU_VALIDA = rsVerificadorMenu("COD_PAGINA")

			If COD_PAGINA_VALIDADA_MENU_VALIDA = 11 Then
			'Produ??es
				MENU_1 = 1
			End If

			If COD_PAGINA_VALIDADA_MENU_VALIDA = 8 Then
			'Formandos
				MENU_2 = 1
			End If

			If COD_PAGINA_VALIDADA_MENU_VALIDA = 26 Then
			'Produtos e Valores
				MENU_3 = 1
			End If

			If COD_PAGINA_VALIDADA_MENU_VALIDA = 55 Then
			'Relat?rios
				MENU_4 = 1
			End If

			If COD_PAGINA_VALIDADA_MENU_VALIDA = 36 Then
			'Pedidos Pendentes
				MENU_5 = 1
			End If

			If COD_PAGINA_VALIDADA_MENU_VALIDA = 32 Then
			'Devolu??o e Sa?da
				MENU_6 = 1
			End If

			If COD_PAGINA_VALIDADA_MENU_VALIDA = 7 Then
			'Controle de Usu?rios
				MENU_7 = 1
			End If

			If COD_PAGINA_VALIDADA_MENU_VALIDA = 64 Then
			'Espa?o de Armazenamento
				MENU_8 = 1
			End If

			If COD_PAGINA_VALIDADA_MENU_VALIDA = 72 Then
			'Corre??o diretoria
				MENU_9 = 1
			End If

			If COD_PAGINA_VALIDADA_MENU_VALIDA = 134 Then
			'Movimenta??o de Repasse
				MENU_10 = 1
			End If

			If COD_PAGINA_VALIDADA_MENU_VALIDA = 116 Then
			'Controle de DVD's
				MENU_11 = 1
			End If

			If COD_PAGINA_VALIDADA_SELECT_VALIDA = 150 Then
			'Repasses
				MENU_12 = 1
			End If

			If COD_PAGINA_VALIDADA_MENU_VALIDA = 167 Then
			'Consulta de Cheques
				MENU_13 = 1
			End If

			If COD_PAGINA_VALIDADA_MENU_VALIDA = 275 Then
			'Controle de Produ??es
				MENU_14 = 1
			End If

			If COD_PAGINA_VALIDADA_MENU_VALIDA = 201 Then
			'Fechamentos
				MENU_15 = 1
			End If

			If COD_PAGINA_VALIDADA_MENU_VALIDA = 217 Then
			'Aprova??es Pendentes
				MENU_16 = 1
			End If

			If COD_PAGINA_VALIDADA_MENU_VALIDA = 253 Then
			'Produ??o P?s-Venda //p?gina ordens_servico.asp
				MENU_17 = 1
			End If

			If COD_PAGINA_VALIDADA_MENU_VALIDA = 241 Then
			'Pedidos de Compras
				MENU_18 = 1
			End If

            If COD_PAGINA_VALIDADA_MENU_VALIDA = 279 Then
			'Estoque
				MENU_19 = 1
			End If

            If COD_PAGINA_VALIDADA_MENU_VALIDA = 280 Then
			'Projetos
				MENU_20 = 1
			End If

			 If COD_PAGINA_VALIDADA_MENU_VALIDA = 326 Then
			'Projetos
				MENU_21 = 1
			End If

			'COBRAN?A
			If COD_PAGINA_VALIDADA_MENU_VALIDA = 360 Then
			'Projetos
				MENU_22 = 1
			End If

			If COD_PAGINA_VALIDADA_MENU_VALIDA = 381 Then
			'Projetos
				MENU_23 = 1
			End If

      If COD_PAGINA_VALIDADA_MENU_VALIDA = 388 Then
			'Projetos
				MENU_24 = 1
			End If
			
			If COD_PAGINA_VALIDADA_MENU_VALIDA = 1984 Then
    			'Renovação de Álbum
    				MENU_25 = 1
			End If


		rsVerificadorMenu.MoveNext()
		Wend

rsVerificadorMenu.Close
Set rsVerificadorMenu=Nothing
FecharConexao()
%>
<nav class="dashboard-nav" aria-label="Navega&ccedil;&atilde;o principal"><a class="active" href="adm_index.asp" aria-current="page">Vis&atilde;o geral</a>
<%If MENU_1 = 1 Then%><a href="dip_producoes.asp">Produ&ccedil;&otilde;es</a><%End If%>
<%If MENU_11 = 1 Then%><a href="cap_producoes.asp">Controle de DVD's</a><%End If%>
<%If MENU_9 = 1 Then%><a href="lca_correcao_diretoria.asp">Autenticar Fichas</a><%End If%>
<%If MENU_2 = 1 Then%><a href="dip_formandos.asp">Cadastro de Formandos</a><%End If%>
<%If MENU_3 = 1 Then%><a href="dip_produtos.asp">Produtos e Valores</a><%End If%>
<%If MENU_4 = 1 Then%><a href="adm_relatorios.asp">Relat&oacute;rios</a><%End If%>
<%If MENU_5 = 1 Then%><a href="lca_pedidos_pendentes.asp">Pedidos Pendentes</a><%End If%>
<%If MENU_6 = 1 Then%><a href="epv_estoque.asp">Devolu&ccedil;&atilde;o e Sa&iacute;da </a><%End If%>
<%If MENU_8 = 1 Then%><a href="epv_espaco_de_armazenamento_salas.asp">Armazenamento</a><%End If%>
<%If MENU_10 = 1 Then%><a href="dip_repasse_inicio.asp">Transfer&ecirc;ncia de Itens</a><%End If%>
<%If MENU_7 = 1 Then%><a href="adm_usuarios.asp">Controle de Usu&aacute;rios</a><%End If%>
<%If MENU_12 = 1 Then%><a href="rep_inicio.asp">Controle de Repasses</a><%End If%>
<%If MENU_14 = 1 Then%><a href="dip_controle_producoes.asp">Controle de Produ&ccedil;&otilde;es</a><%End If%>
<%If MENU_16 = 1 Then%><a href="adm_aprovacoes_pendentes.asp">Aprova&ccedil;&otilde;es Pendentes</a><%End If%>
<%If MENU_17 = 1 Then%><a href="ppv_ordens_servico.asp">Produ&ccedil;&atilde;o P&oacute;s-Venda</a><%End If%>
<%If MENU_18 = 1 Then%><a href="ppv_compras_pedidos.asp">Compras</a><%End If%>
<%If MENU_15 = 1 Then%><a href="cap_fechamentos.asp">Fechamentos</a><%End If%>
<%If MENU_19 = 1 Then%><a href="est_estoque.asp">Estoque Interno (Produtos)</a><%End If%>
<%If MENU_19 = 1 Then%><a href="est_resumo.asp">Resumo do estoque</a><%End If%>
<%If MENU_20 = 1 Then%><a href="ple_projetos.asp">Planejamento Estrat&eacute;gico</a><%End If%>
<%If MENU_21 = 1 Then%><a href="mck_formandos.asp">Mackenzie</a><%End If%>
<%If MENU_22 = 1 Then%><a href="cob_pedidos_boletos.asp">Cobran&ccedil;a</a><%End If%>
<%If MENU_23 = 1 Then%><a href="adm_rotinas_configuracoes.asp">Configura&ccedil;&otilde;es</a><%End If%>
<%If MENU_24 = 1 Then%><a href="epv_renovacao.asp?comando=nova">Renova&ccedil;&atilde;o em geral</a><%End If%>
<%If MENU_25 = 1 Then%><a href="adm_controle_no_extrato_comissao_vendedores_boletos.asp">&Aacute;lbuns Digitais</a><%End If%>
</nav>