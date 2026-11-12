<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="files/includes/functions_local.asp" -->
<%
Dim caminho_pagina, calendario_necessario
caminho_pagina = "ppv_compras_pedidos_produtos_solicitados.asp"
calendario_necessario = "SIM"
%>
<!-- #include file="includes/classe_modulo_ppv.asp" -->
    <%
	AbrirConexao()
	Dim acesso_entrada_pedido_compra, acesso_excluir_entrada_errada, validacao_para_entrada_qualquer_pedido
	acesso_entrada_pedido_compra = validaUsuario(2,,66,true)
	acesso_excluir_entrada_errada = validaUsuario(2,,69,true)
	validacao_para_entrada_qualquer_pedido = validaUsuario(2,,74,true)
	 
	
	Dim sql, sqlWhere, sqlOrderBy
	
	sql = "SELECT VW_TB_ORDENS_SERVICO.NOME_FORMANDO, VW_TB_ORDENS_SERVICO.ENTIDADE_ENSINO, VW_TB_PEDIDOS_COMPRAS_PRODUTOS.* FROM VW_TB_PEDIDOS_COMPRAS_PRODUTOS LEFT JOIN VW_TB_ORDENS_SERVICO on VW_TB_ORDENS_SERVICO.COD_ORDEM_SERVICO = VW_TB_PEDIDOS_COMPRAS_PRODUTOS.COD_ORDEM_SERVICO WHERE VW_TB_PEDIDOS_COMPRAS_PRODUTOS.status_pedido_compra <> 0"
	'sql = "SELECT * FROM VW_TB_PEDIDOS_COMPRAS_PRODUTOS WHERE status_pedido_compra <> 0 "
	sqlOrderBy =  " order by  VW_TB_PEDIDOS_COMPRAS_PRODUTOS.data_previsao_entrega asc"
	
	Dim status_produto, diferenca_prazo_entrega
	
	'/////////////
	'Filtro
		
		Dim filtro_status_produto
		Dim filtro_cod_fornecedor
		Dim filtro_NUMERO_PEDIDO_COMPRAS
		Dim filtro_COD_CONTRATO
		Dim filtro_ordernar_por
		Dim b_DATA_PREVISAO_ENTREGA_INICIO
		Dim b_DATA_PREVISAO_ENTREGA_FIM
		Dim FILTRO_COD_PRODUTO

		
		filtro_NUMERO_PEDIDO_COMPRAS	=	Cstr(""&Request.QueryString("NUMERO_PEDIDO_COMPRAS"))
		filtro_COD_CONTRATO				=	Cstr(""&Request.QueryString("COD_CONTRATO"))
		filtro_status_produto			=	Cstr(""&Request.QueryString("status_produto"))
		filtro_cod_fornecedor			=	Cstr(""&Request.QueryString("cod_fornecedor"))
		filtro_ordernar_por				=	Cstr(""&Request.QueryString("ordernar_por"))
		b_DATA_PREVISAO_ENTREGA_INICIO	=	Request.QueryString("DATA_PREVISAO_ENTREGA_INICIO")
		b_DATA_PREVISAO_ENTREGA_FIM		=	Request.QueryString("DATA_PREVISAO_ENTREGA_FIM")
		FILTRO_COD_PRODUTO				=	Cstr(""&Request.QueryString("COD_PRODUTO"))
		
		If Not validacao_para_entrada_qualquer_pedido Then
			sqlWhere = sqlWhere&" AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.cadastro_cod_usuario = '"&Session("cod_usuario")&"'"
		End If
		
		If Request.QueryString("filtro") <> "" Then
			
			If filtro_ordernar_por <> "" Then
				If filtro_ordernar_por = "1" Then
					sqlOrderBy =  " order by  VW_TB_PEDIDOS_COMPRAS_PRODUTOS.data_previsao_entrega asc"
				ElseIf filtro_ordernar_por = "2" Then
					sqlOrderBy =  " order by  VW_TB_ORDENS_SERVICO.NOME_FORMANDO asc"
				ElseIf filtro_ordernar_por = "3" Then
					sqlOrderBy =  " order by  VW_TB_ORDENS_SERVICO.ENTIDADE_ENSINO asc"
				End If
				
			End If
			
			If filtro_NUMERO_PEDIDO_COMPRAS <> "" Then
			'Filtro preferencial, limpa os demais
				
				filtro_cod_fornecedor = ""
				filtro_cadastro_cod_usuario = ""
				b_DATA_PREVISAO_ENTREGA_INICIO = ""
				b_DATA_PREVISAO_ENTREGA_FIM	= ""
				filtro_COD_CONTRATO = ""
				FILTRO_COD_PRODUTO = ""
				sqlWhere = " AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.numero_pedido_compra = '"&filtro_NUMERO_PEDIDO_COMPRAS&"'"
				
				If Not validacao_para_entrada_qualquer_pedido Then
					sqlWhere = sqlWhere&" AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.cadastro_cod_usuario = '"&Session("cod_usuario")&"'"
				End If
				
				If filtro_status_produto <> "" Then
					sqlWhere = sqlWhere&" AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.status_produto = '"&filtro_status_produto&"'"
				End If
			
			Else
			
				If filtro_COD_CONTRATO <> "" Then
					sqlWhere = sqlWhere&" AND VW_TB_ORDENS_SERVICO.COD_CONTRATO = '"&filtro_COD_CONTRATO&"'"
				End If
		
				If filtro_status_produto <> "" Then
					sqlWhere = sqlWhere&" AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.status_produto = '"&filtro_status_produto&"'"
				End If
				
				if filtro_cod_fornecedor <> "" Then
					sqlWhere = sqlWhere&" AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.cod_fornecedor = '"&filtro_cod_fornecedor&"'"
				End if
				
				if filtro_cadastro_cod_usuario <> "" Then
					sqlWhere = sqlWhere&" AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.cadastro_cod_usuario = '"&filtro_cadastro_cod_usuario&"'"
				End if	
	
				if FILTRO_COD_PRODUTO <> "" Then
					sqlWhere = sqlWhere&" AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.cod_produto = '"&FILTRO_COD_PRODUTO&"'"
				End if			
				
				
				If b_DATA_PREVISAO_ENTREGA_INICIO <> Empty or b_DATA_PREVISAO_ENTREGA_FIM <> Empty Then
							
							
							If b_DATA_PREVISAO_ENTREGA_INICIO <> Empty Then
								b_DATA_PREVISAO_ENTREGA_INICIO = Day(b_DATA_PREVISAO_ENTREGA_INICIO)&"/"&Month(b_DATA_PREVISAO_ENTREGA_INICIO)&"/"&Year(b_DATA_PREVISAO_ENTREGA_INICIO)&" 00:00:00"
							End If
		
							If b_DATA_PREVISAO_ENTREGA_FIM <> Empty Then
								b_DATA_PREVISAO_ENTREGA_FIM = Day(b_DATA_PREVISAO_ENTREGA_FIM)&"/"&Month(b_DATA_PREVISAO_ENTREGA_FIM)&"/"&Year(b_DATA_PREVISAO_ENTREGA_FIM)&"  23:59:59"
							End If					
																	
							If b_DATA_PREVISAO_ENTREGA_INICIO <> Empty and b_DATA_PREVISAO_ENTREGA_FIM <> Empty Then
								sqlWhere = sqlWhere&" AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.DATA_PREVISAO_ENTREGA BETWEEN "&sqlData(b_DATA_PREVISAO_ENTREGA_INICIO)&" AND "&sqlData(b_DATA_PREVISAO_ENTREGA_FIM)
							Elseif b_DATA_PREVISAO_ENTREGA_INICIO = Empty and b_DATA_PREVISAO_ENTREGA_FIM <> Empty Then
								sqlWhere = sqlWhere&" AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.DATA_PREVISAO_ENTREGA <= "&sqlData(b_DATA_PREVISAO_ENTREGA_FIM)
							Elseif b_DATA_PREVISAO_ENTREGA_INICIO <> Empty and b_DATA_PREVISAO_ENTREGA_FIM = Empty Then
								sqlWhere = sqlWhere&" AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.DATA_PREVISAO_ENTREGA >= "&sqlData(b_DATA_PREVISAO_ENTREGA_INICIO)
							End If
					
							If b_DATA_PREVISAO_ENTREGA_FIM <> Empty Then
							'Corrigi data
								b_DATA_PREVISAO_ENTREGA_FIM = Day(b_DATA_PREVISAO_ENTREGA_FIM)&"/"&Month(b_DATA_PREVISAO_ENTREGA_FIM)&"/"&Year(b_DATA_PREVISAO_ENTREGA_FIM)&"  00:00:00"
							End If	
					
				End If	
			
			
			End If

			
		Else
				'Padrão filtrado como somente pendentes		
				filtro_status_produto = "P"
				sqlWhere = " AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.status_produto = 'P' "
		End If	
		



'Junta tudo
sql = sql&sqlWhere&sqlOrderBy

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
If Request.Form("campo1") <> Empty and acesso_entrada_pedido_compra Then
		
	'Limpa a base temporária para o usuário
	objConexao.Execute("DELETE FROM TB_PEDIDOS_COMPRAS_PRODUTOS_TEMPORARIA WHERE cod_usuario = "&sqlTexto(Session("cod_usuario")))
	
	Set rs=objConexao.Execute("SELECT VW_TB_PEDIDOS_COMPRAS_PRODUTOS.* FROM VW_TB_ORDENS_SERVICO INNER JOIN VW_TB_PEDIDOS_COMPRAS_PRODUTOS on VW_TB_ORDENS_SERVICO.COD_ORDEM_SERVICO = VW_TB_PEDIDOS_COMPRAS_PRODUTOS.COD_ORDEM_SERVICO WHERE VW_TB_PEDIDOS_COMPRAS_PRODUTOS.status_pedido_compra <> 0 AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.status_produto = 'P'")
		
		While not rs.Eof
				
				cod_pedido_compra_produto = rs("cod_pedido_compra_produto")
				cod_ordem_servico_etapa = Cstr(""&rs("cod_ordem_servico_etapa"))
				cod_ordem_servico = Cstr(""&rs("cod_ordem_servico"))
				cod_pedido_compra = rs("cod_pedido_compra")
				
				If Cstr(""&Request.Form("cod_pedido_compra_produto_"&cod_pedido_compra_produto))  = "S" Then
					
					'//////
					'Insere na base de dados temporária para gerar o recibo futuramente
					objConexao.Execute("INSERT INTO TB_PEDIDOS_COMPRAS_PRODUTOS_TEMPORARIA (cod_pedido_compra_produto,cod_usuario) VALUES("&sqlTexto(cod_pedido_compra_produto)&","&sqlTexto(Session("cod_usuario"))&")")
					'//////
					
					
					'1) Baixa o item na compra
					objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS_PRODUTOS SET status_produto = 'E', quantidade_entregue = "&sqlNumero(rs("quantidade_solicitada"))&" WHERE cod_pedido_compra_produto = "&sqlNumero(cod_pedido_compra_produto)) 
					
					'2) Verifica se não há mais itens para serem baixados do pedido
					Set rs2=objConexao.Execute("SELECT cod_pedido_compra FROM TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra = "&sqlTexto(cod_pedido_compra)&" and status_produto = 'P'")
						
						If rs2.Eof Then
							'Finaliza o Pedido de Compra se não encontrar nenhum produto pendente
							objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS SET STATUS_PEDIDO_COMPRA = 3 WHERE cod_pedido_compra = "&sqlTexto(cod_pedido_compra))
														
						End If
						
					rs2.Close
					Set rs2=Nothing
					
					'3) Finaliza etapa da OS
					If cod_ordem_servico_etapa <> "" Then
						
						'Verifica a qtd_etapas
						Set rs3=objConexao.Execute("Select count(*) as qtd_etapas From TB_ORDENS_SERVICO_ETAPAS where cod_ordem_servico = "&cod_ordem_servico&" group by cod_ordem_servico")
													
							If not rs3.Eof Then
								qtd_etapas = formatNumber(0&rs3("qtd_etapas"),0)
							End If
													
						rs3.Close
						Set rs3=Nothing
						
						'Verifica qual o número da etapa atual
						Set rs3=objConexao.Execute("Select numero_etapa From VW_TB_ORDENS_SERVICO where cod_ordem_servico = "&cod_ordem_servico)
					
							If not rs3.Eof Then
								numero_etapa = formatNumber(0&rs3("numero_etapa"),0)
							End If
													
						rs3.Close
						Set rs3=Nothing

						If int(qtd_etapas) = int(numero_etapa) Then
						'Deverá finalizar OS, pois é a última etapa
							sql = "UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_termino = NOW() WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&sqlTexto(cod_ordem_servico_etapa)
							objConexao.Execute(sql)
							sql = "UPDATE TB_ORDENS_SERVICO SET status_ordem_servico = 2, data_finalizado = NOW(), cod_usuario_finalizou = "&Session("cod_usuario")&" WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)
							objConexao.Execute(sql)
						Else
						'Finaliza a etapa atual				
							sql = "UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_termino = NOW() WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&sqlTexto(cod_ordem_servico_etapa)		
							objConexao.Execute(sql)
							'Inicia a próxima Etapa
							Set rs3=objConexao.Execute("SELECT cod_ordem_servico_etapa FROM TB_ORDENS_SERVICO_ETAPAS WHERE cod_ordem_servico = "&cod_ordem_servico&" and numero_etapa IN(SELECT (numero_etapa+1) FROM VW_TB_ORDENS_SERVICO WHERE cod_ordem_servico = "&cod_ordem_servico&")")
												
									If not rs3.Eof Then
										sql = "UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_inicio = NOW() WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&rs3("cod_ordem_servico_etapa")
										objConexao.Execute(sql)
										sql = "UPDATE TB_ORDENS_SERVICO SET cod_ordem_servico_etapa_atual = "&rs3("cod_ordem_servico_etapa")&" WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)
										objConexao.Execute(sql)
									End If
												
							rs3.Close
							Set rs3=Nothing
										
						End If
						
					
					End If
					
					
				End If
				
		rs.MoveNext()
		Wend
		
	rs.Close
	Set rs=Nothing
	
	
	FecharConexao()
	
	Response.Redirect("ppv_compras_pedidos_produtos_solicitados.asp?codmsg=1")


End If
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


%>
<%FecharConexao()%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Produtos solicitados | Portal Stillos</title>
<link rel="stylesheet" href="css/vendor/bootstrap.min.css">
<link rel="stylesheet" href="css/dashboard.css">
<link rel="stylesheet" href="css/modal-message.css">
<!-- #include file="files/includes/scripts.asp" -->
<style>
.sol-page .dashboard-nav a.active:not([aria-current]) { background:transparent; color:#d4d5e1; }
.sol-page .dashboard-nav a[aria-current="page"] { background:#b20b61; color:#fff; font-weight:600; }
.sol-page .sol-panel { height:auto; margin-bottom:16px; padding:clamp(14px,3vw,26px); }
.sol-muted { color:#57566a; }
.sol-filters { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:14px; margin:14px 0; }
.sol-filters label { display:block; font-weight:650; margin-bottom:5px; }
.sol-date, .sol-actions { display:flex; align-items:center; flex-wrap:wrap; gap:8px; }
.sol-date { flex-wrap:nowrap; }
.sol-date input { min-width:0; }
.sol-date button { padding:8px; }
.sol-heading { display:flex; justify-content:space-between; flex-wrap:wrap; gap:10px; margin-bottom:14px; }
.sol-table { width:100%; margin:0; }
.sol-table th, .sol-table td { padding:10px 8px; vertical-align:top; }
.sol-table th { background:#f4f1f7; color:#42364e; }
.sol-table tr:has(.sol-select:checked) > td { background:#f5fbf7; }
.sol-table small { display:block; color:#57566a; }
.sol-link { display:block; margin-top:4px; padding:0; border:0; background:none; color:#88134f; text-decoration:underline; font:inherit; cursor:pointer; }
.sol-item-details { margin-top:5px; }
.sol-item-details summary { color:#88134f; cursor:pointer; }
.sol-item-details dl { margin:8px 0 0; overflow-wrap:anywhere; }
.sol-item-details dd { margin:0 0 6px; }
.sol-status { display:inline-block; padding:4px 7px; border-radius:5px; background:#f0edf8; color:#51436e; font-weight:650; }
.sol-ok { background:#e7f3ec; color:#286143; }
.sol-today { background:#fdf2db; color:#795419; }
.sol-late { background:#fce9ee; color:#8e2045; }
.sol-pagination { justify-content:center; margin-top:14px; }
.sol-review { position:sticky; bottom:0; z-index:2; display:flex; flex-wrap:wrap; align-items:center; justify-content:flex-end; gap:8px; padding:10px; margin-top:14px; border:1px solid #e4dfeb; border-radius:8px; background:#fff; }
.sol-review p { flex-basis:100%; color:#8e2045; margin:0; }
.sol-dialog { width:min(480px,calc(100% - 24px)); padding:24px; border:1px solid #e4dfeb; border-radius:14px; color:var(--ink); }
.sol-dialog::backdrop { background:#24273899; }
.sol-dialog .sol-actions { justify-content:flex-end; }
@media(max-width:1050px) {
  .sol-filters { grid-template-columns:repeat(2,minmax(0,1fr)); }
  .sol-table, .sol-table tbody { display:block; }
  .sol-table thead { position:absolute; width:1px; height:1px; overflow:hidden; clip-path:inset(50%); }
  .sol-table tbody tr { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); margin-bottom:10px; border:1px solid #e4dfeb; border-radius:8px; }
  .sol-table td { display:block; min-width:0; border:0; overflow-wrap:anywhere; }
  .sol-table td:nth-child(3) { grid-column:1/-1; background:#f7f5fa; }
  .sol-table td::before { content:attr(data-label); display:block; color:#57566a; font-size:.85em; font-weight:650; }
}
@media(max-width:480px) { .sol-filters { grid-template-columns:1fr; } }
</style>
</head>
<body class="sol-page">
<a class="skip-link" href="#conteudo">Ir para o conte&uacute;do</a>
<div class="dashboard-layout">
  <aside class="dashboard-sidebar" id="menuPrincipal">
    <a class="dashboard-brand" href="adm_index.asp"><img src="images/logoparafundopreto.png" alt="Stillos" width="180" /><span>GERENCIAL FOTOS</span></a>
    <div class="nav-caption">&Aacute;REA DE TRABALHO</div>
    <!-- #include file="includes/dashboard_menu.asp" -->
  </aside>
  <div class="dashboard-workspace">
    <header class="dashboard-topbar">
      <button class="btn menu-toggle" id="ppvMenuToggle" type="button" aria-controls="menuPrincipal" aria-expanded="false">Menu</button>
      <div class="topbar-title">Portal Stillos <span>/ Compras</span></div>
      <div class="topbar-actions"><a class="user-name" href="adm_log_usuarios.asp"><%=Server.HTMLEncode(Session("NOME") & "")%></a><a href="manual_sistema.asp" target="_blank" rel="noopener">Manual</a><a class="btn btn-outline-secondary btn-sm" href="adm_sair.asp">Sair</a></div>
    </header>
    <main class="dashboard-main ppv-main" id="conteudo">
      <div class="dashboard-heading">
        <div><h1>Produtos solicitados</h1><p>Consulte os itens dos pedidos de compra e registre os recebimentos.</p></div>
        <a href="adm_index.asp" class="btn btn-outline-secondary">Voltar ao painel</a>
      </div>
      <div class="search-panel"><div><strong>Busca no portal</strong></div><!-- #include file="includes/dashboard_busca.asp" --></div>
      <div class="session-notices"><!-- #include file="includes/dashboard_avisos.asp" --></div>
<!-- #include file="includes/compras_sub_menu.asp" -->
<section class="dashboard-panel sol-panel" aria-labelledby="sol-filter-title">
<h2 id="sol-filter-title">Filtrar itens</h2>
<p class="sol-muted">Informe um fornecedor ou o n&uacute;mero do pedido de compra para listar os itens.</p>
<form action="ppv_compras_pedidos_produtos_solicitados.asp" name="FORMULARIO_FILTRO" method="get">
<input type="hidden" value="1" name="filtro">
<div class="sol-filters">
<div><label for="sol-order">N&uacute;mero do pedido</label><input class="form-control" id="sol-order" inputmode="numeric" name="NUMERO_PEDIDO_COMPRAS" value="<%=Server.HTMLEncode(filtro_NUMERO_PEDIDO_COMPRAS & "")%>"></div>
<div><label for="sol-contract">C&oacute;digo do contrato</label><input class="form-control" id="sol-contract" name="COD_CONTRATO" value="<%=Server.HTMLEncode(filtro_COD_CONTRATO & "")%>"></div>
<div><label for="sol-ordernar_por">Ordenar por</label><select class="form-select" id="sol-ordernar_por" name="ordernar_por">
      <option <%If Cstr(filtro_ordernar_por) = "1" Then%> selected="selected" <%End If%> value="1">Data de Entrega (Crescente)</option>
      <option <%If Cstr(filtro_ordernar_por) = "2" Then%> selected="selected" <%End If%> value="2">Formando</option>
      <option <%If Cstr(filtro_ordernar_por) = "3" Then%> selected="selected" <%End If%> value="3">Institui&ccedil;&atilde;o de Ensino</option>
    </select></div><div><label for="sol-status_produto">Status do produto</label><select class="form-select" id="sol-status_produto" name="status_produto">
      <option <%If Cstr(filtro_status_produto) = "" Then%> selected="selected" <%End If%> value="">Todos</option>
      <option <%If Cstr(filtro_status_produto) = "P" Then%> selected="selected" <%End If%> value="P">Pendentes</option>
      <option <%If Cstr(filtro_status_produto) = "E" Then%> selected="selected" <%End If%> value="E">Entregues</option>
      <option <%If Cstr(filtro_status_produto) = "C" Then%> selected="selected" <%End If%> value="C">Cancelados</option>
    </select></div>
<%If validacao_para_entrada_qualquer_pedido Then%>
<div><label for="sol-CADASTRO_COD_USUARIO">Respons&aacute;vel pelo pedido</label><select class="form-select" id="sol-CADASTRO_COD_USUARIO" name="CADASTRO_COD_USUARIO">
      <option value="" <%If Cstr(filtro_cadastro_cod_usuario) = "" Then%> selected="selected" <%End If%>>Todos</option>
      <%                                          
                                            Set rs3=objConexao.Execute("SELECT COD_USUARIO, NOME FROM TB_ADM_USUARIOS WHERE COD_USUARIO IN(SELECT cadastro_cod_usuario FROM TB_PEDIDOS_COMPRAS WHERE status_pedido_compra > 0 GROUP BY cadastro_cod_usuario) ORDER BY NOME ASC")
 
 											If not rs3.Eof Then
 
                                            	While not rs3.Eof%>
      <option value="<%=rs3("COD_USUARIO")%>" <%If Cstr(filtro_cadastro_cod_usuario) = Cstr(rs3("cod_usuario")) Then%> selected="selected" <%End If%>><%=rs3("NOME")%></option>
      <%
                	                            rs3.MoveNext()
                    	                        Wend	
                                            
											End If
                                        
                                            rs3.Close
                                            Set rs3=Nothing
                                        %>
    </select></div>
<%Else%>
<div><strong>Respons&aacute;vel pelo pedido</strong><p class="mb-0"><%=nomeUsuario(Session("cod_usuario"),true)%></p></div>
<%End If%>
<div><label for="sol-COD_FORNECEDOR">Fornecedor</label><select class="form-select" id="sol-COD_FORNECEDOR" name="COD_FORNECEDOR">
      <option value="" <%If Cstr(filtro_cod_fornecedor) = "" Then%> selected="selected" <%End If%>>Todos</option>
      <%                                          
                                            Set rs3=objConexao.Execute("SELECT cod_fornecedor, NOME_FANTASIA FROM TB_FORNECEDORES WHERE TB_FORNECEDORES.ATIVADO = 1 order by TB_FORNECEDORES.NOME_FANTASIA asc")
 
 											If not rs3.Eof Then
 
                                            	While not rs3.Eof%>
      <option value="<%=rs3("cod_fornecedor")%>" <%If Cstr(filtro_cod_fornecedor) = Cstr(rs3("cod_fornecedor")) Then%> selected="selected" <%End If%>><%=rs3("NOME_FANTASIA")%></option>
      <%
                	                            rs3.MoveNext()
                    	                        Wend	
                                            
											End If
                                        
                                            rs3.Close
                                            Set rs3=Nothing
                                        %>
    </select></div><div><label for="sol-COD_PRODUTO">Produto</label><select class="form-select" id="sol-COD_PRODUTO" name="COD_PRODUTO">
          <option value=''>Selecione</option>
          <%
Set rsE3=objConexao.Execute("SELECT COD_PRODUTO, NOME_PRODUTO FROM TB_PRODUTOS WHERE ATIVO = 1 AND gera_ordem_servico_prazo_cliente > 0 ORDER BY NOME_PRODUTO")
While not rsE3.Eof 
'Padrao
Dim selRelPadrao3, verificador3
verificador3 = ""&rsE3(0)

If CStr(verificador3) = CStr(filtro_COD_PRODUTO)  Then%>
          <option selected='selected' value="<%=rsE3(0)%>"><%=rsE3(1)%></option>
          <%Else%>
          <option value="<%=rsE3(0)%>"><%=rsE3(1)%></option>
          <%End if%>
          <%rsE3.MoveNext()
Wend
rsE3.Close
Set rsE3=Nothing
%>
        </select></div><div><label for="sol-DATA_PREVISAO_ENTREGA_INICIO">Previs&atilde;o de entrega a partir de</label><div class="sol-date"><input class='form-control' id='sol-DATA_PREVISAO_ENTREGA_INICIO' inputmode='numeric' placeholder='dd/mm/aaaa' maxlength='10' onkeydown='mascara(this,data);' onkeypress='mascara(this,data);' onkeyup='mascara(this,data);' type='text' name='DATA_PREVISAO_ENTREGA_INICIO' value="<%=ArrumaData(b_DATA_PREVISAO_ENTREGA_INICIO)%>" onblur='validarData(this)' /><button class="btn btn-outline-secondary" type="button" onclick="pickDate(this,document.FORMULARIO_FILTRO.DATA_PREVISAO_ENTREGA_INICIO)" aria-label="Abrir calend&aacute;rio"><img alt="" width="19" height="17" src="images/icons/icon_calendario.jpg"></button></div></div><div><label for="sol-DATA_PREVISAO_ENTREGA_FIM">Previs&atilde;o de entrega at&eacute;</label><div class="sol-date"><input class='form-control' id='sol-DATA_PREVISAO_ENTREGA_FIM' inputmode='numeric' placeholder='dd/mm/aaaa' maxlength='10' onkeydown='mascara(this,data);' onkeypress='mascara(this,data);' onkeyup='mascara(this,data);' type='text' name='DATA_PREVISAO_ENTREGA_FIM' value="<%=ArrumaData(b_DATA_PREVISAO_ENTREGA_FIM)%>" onblur='validarData(this)' /><button class="btn btn-outline-secondary" type="button" onclick="pickDate(this,document.FORMULARIO_FILTRO.DATA_PREVISAO_ENTREGA_FIM)" aria-label="Abrir calend&aacute;rio"><img alt="" width="19" height="17" src="images/icons/icon_calendario.jpg"></button></div></div>
</div>
<div class="sol-actions"><button class="btn btn-primary" type="submit" name="filtro1" value="1">Aplicar filtros</button><a class="btn btn-outline-secondary" href="ppv_compras_pedidos_produtos_solicitados.asp">Limpar filtros</a></div>
</form>
</section>
<%
'/////////////////////////////////////////////////////////////////////////////////
	'Verifica se o último recibo está disponível
	Dim ultimo_recibo_disponivel
	ultimo_recibo_disponivel = false
	
		Set rs2=objConexao.Execute("SELECT * FROM VW_TB_PEDIDOS_COMPRAS_PRODUTOS_TEMPORARIA WHERE cod_usuario_temporario = "&Session("cod_usuario")&" LIMIT 1")
			
			If not rs2.Eof Then
				ultimo_recibo_disponivel = true
			End If
			
		rs2.Close
		Set rs2=Nothing
	
If Request.QueryString("codmsg") = "1" And Request.QueryString("sucesso") = "" Then
%><p class="alert alert-success" role="status">Entrada de itens realizada com sucesso.</p><%
End If
If ultimo_recibo_disponivel Then
%>
<div class="dashboard-panel sol-panel sol-actions"><a class="btn btn-outline-secondary" target="_blank" rel="noopener" href="ppv_compras_pedidos_produtos_solicitados_recibo_entrega.asp">Imprimir &uacute;ltimo recibo de entrega</a></div>
<%End If%>
    <%
	If Request.QueryString("cod_fornecedor") <> Empty or Request.QueryString("NUMERO_PEDIDO_COMPRAS") <> Empty Then
	%>
    <%

	'Response.Write(sql)
	Set rs=server.CreateObject("ADODB.recordset")
	rs.CursorLocation=3  'Cursor no cliente
	rs.PageSize=200
	rs.Open sql, objConexao
	Dim paginacaocao
	paginacaocao = formatNumber(0&Request.QueryString("paginacao"),0)
	
	If not rs.Eof Then
	
		If paginacaocao > 0 Then
			If paginacaocao > rs.PageCount Then paginacaocao = rs.PageCount
			rs.AbsolutePage=paginacaocao
		Else
			paginacaocao = 1
			rs.AbsolutePage=1
		End If	
%>


<section class="dashboard-panel sol-panel" aria-labelledby="sol-results-title">
<div class="sol-heading"><div><h2 id="sol-results-title">Itens encontrados</h2><span class="sol-muted"><%=rs.RecordCount%> itens</span></div><span>P&aacute;gina <%=rs.AbsolutePage%> de <%=rs.PageCount%></span></div>
<form id="sol-entry" action="ppv_compras_pedidos_produtos_solicitados.asp" name="FORMULARIO_PADRAO" method="post">
<input type="hidden" name="campo1" value="S">
<table class="table sol-table" role="table" aria-labelledby="sol-results-title">
<thead role="rowgroup"><tr role="row"><th scope="col" role="columnheader">Entrada</th><th scope="col" role="columnheader">Pedido</th><th scope="col" role="columnheader">Produto e fornecedor</th><th scope="col" role="columnheader">Qtd.</th><th scope="col" role="columnheader">Previs&atilde;o</th><th scope="col" role="columnheader">Status</th></tr></thead>
<tbody role="rowgroup">
<%
Dim dias_atrasados, acesso_cancelamento
acesso_cancelamento = validaUsuario(2,,65,true)


Dim usuario_pode_imprimir
usuario_pode_imprimir = validaUsuario(2,,64,true)

Dim x
For x=1 To rs.PageSize
If Not rs.Eof Then
	status_produto = Cstr(""&rs("status_produto"))
	dias_atrasados = formatNumber(rs("dias_atrasados"))
%>
<tr role="row">
<td role="cell" data-label="Entrada">
  <%If status_produto = "P" And acesso_entrada_pedido_compra Then%>
    <input type="checkbox" name="cod_pedido_compra_produto_<%=rs("cod_pedido_compra_produto")%>" value="S" class="form-check-input sol-select" aria-label="Selecionar item do pedido <%=rs("numero_pedido_compra")%>">
  <%Else%>&mdash;<%End If%>
</td>
<td role="cell" data-label="Pedido">
  <strong>#<%=formatNumber(rs("numero_pedido_compra"),0)%></strong>
  <%If usuario_pode_imprimir Then%>
  <%If comprasPaginasPermitidas.Exists("ppv_compras_pedidos_imprimir.asp") Then%>
    <a class="sol-link" href="ppv_compras_pedidos_imprimir.asp?c=<%=rs("cod_pedido_compra")%>" target="_blank" rel="noopener">Imprimir</a>
  <%Else%><button type="button" class="btn btn-outline-secondary btn-sm compras-access-trigger">Imprimir</button><%End If%>
  <%End If%>
</td>
<td role="cell" data-label="Produto e fornecedor"><strong><%=Server.HTMLEncode(rs("nome_produto") & "")%></strong><small><%=Server.HTMLEncode(rs("razao_social") & "")%></small>
<details class="sol-item-details"><summary>Ver detalhes</summary>
    <dl>
      <dt>Ordem de servi&ccedil;o</dt><dd><%=moduloPPVOSMascara(rs("cod_ordem_servico"),"",true)%></dd>
      <dt>Respons&aacute;vel</dt><dd><%
	Set rs2=objConexao.Execute("SELECT cod_usuario_responsavel FROM  VW_TB_ORDENS_SERVICO WHERE cod_ordem_servico = "&sqlTexto(rs("cod_ordem_servico")))
		
		If not rs2.Eof Then
			Response.Write(ucase(nomeUsuario(rs2("cod_usuario_responsavel"),true)))
		End If
		
	rs2.Close
	Set rs2=Nothing
	%></dd>
      <dt>Formando</dt><dd><%=Server.HTMLEncode(rs("nome_formando") & "")%></dd>
      <dt>Institui&ccedil;&atilde;o de ensino</dt><dd><%=Server.HTMLEncode(rs("ENTIDADE_ENSINO") & "")%></dd>
      <dt>Observa&ccedil;&atilde;o</dt><dd><%=Server.HTMLEncode(rs("obs") & "")%></dd>
    </dl>
</details>
</td>
<td role="cell" data-label="Quantidade"><%=ArrumaNumero(rs("quantidade_solicitada"),"-",0)%></td>
<td role="cell" data-label="Previs&atilde;o"><%=ArrumaData(rs("data_previsao_entrega"))%></td>
<td role="cell" data-label="Status">
<div id="containerProdutoPedidoCompra_<%=base64Encode(rs("cod_pedido_compra_produto"))%>">
    <%If status_produto = "E" Then%><span class="sol-status sol-ok">Entregue</span>
      <%If acesso_excluir_entrada_errada Then%><button class="sol-link" type="button" onclick="confirmaExclusaoEntradaErrada('<%=base64Encode(rs("cod_pedido_compra_produto"))%>')">Excluir entrada</button><%End If%>
    <%ElseIf status_produto = "P" And dias_atrasados = 0 Then%><span class="sol-status sol-today">Vence hoje</span>
    <%ElseIf status_produto = "P" And dias_atrasados < 1 Then%><span class="sol-status sol-ok">Em dia</span>
    <%ElseIf status_produto = "P" And dias_atrasados > 0 Then%><span class="sol-status sol-late">Em atraso</span>
    <%ElseIf status_produto = "C" Then%><span class="sol-status">Cancelado</span><%End If%>
</div>
</td>
</tr>
<%rs.MoveNext()
End If
Next%>

</tbody></table>
<%If acesso_entrada_pedido_compra Then%>
<div class="sol-review"><span id="sol-count" role="status">0 itens selecionados</span><button class="btn btn-primary" type="submit" aria-haspopup="dialog">Revisar entrada</button><p id="sol-feedback" role="alert" hidden></p></div>
<%End If%>
</form>
<%If rs.PageCount > 1 Then
    Dim solUrl, solParam
    solUrl = "ppv_compras_pedidos_produtos_solicitados.asp?"
    For Each solParam In Array("filtro","filtro1","NUMERO_PEDIDO_COMPRAS","COD_CONTRATO","ordernar_por","status_produto","CADASTRO_COD_USUARIO","COD_FORNECEDOR","COD_PRODUTO","DATA_PREVISAO_ENTREGA_INICIO","DATA_PREVISAO_ENTREGA_FIM")
        If Len(Request.QueryString(solParam) & "") > 0 Then solUrl = solUrl & solParam & "=" & Server.URLEncode(Request.QueryString(solParam) & "") & "&"
    Next
    solUrl = Server.HTMLEncode(solUrl & "paginacao=")
%>
<nav class="sol-actions sol-pagination" aria-label="Paginar itens">
  <%If paginacaocao > 1 Then%><a class="btn btn-outline-secondary btn-sm" href="<%=solUrl & (paginacaocao-1)%>">Anterior</a><%End If%>
  <span>P&aacute;gina <%=paginacaocao%> de <%=rs.PageCount%></span>
  <%If paginacaocao < rs.PageCount Then%><a class="btn btn-outline-secondary btn-sm" href="<%=solUrl & (paginacaocao+1)%>">Pr&oacute;xima</a><%End If%>
</nav>
<%End If%>
</section>
<%Else%>
<section class="dashboard-panel sol-panel" role="status"><h2>Nenhum item encontrado</h2><p class="mb-0">Confira os filtros para buscar outros itens.</p></section>
<%End If
rs.Close
Set rs = Nothing
End If
FecharConexao()
%>
<div id="tsZXucAdca" hidden></div>
<dialog id="sol-confirm" class="sol-dialog" aria-labelledby="sol-confirm-title" aria-describedby="sol-confirm-message">
  <h2 id="sol-confirm-title">Confirmar opera&ccedil;&atilde;o</h2>
  <p id="sol-confirm-message"></p>
  <form method="dialog" class="sol-actions">
    <button class="btn btn-outline-secondary" value="cancelar" autofocus>Cancelar</button>
    <button class="btn btn-primary" value="confirmar">Confirmar</button>
  </form>
</dialog>
    </main>
    <footer class="dashboard-footer">&copy; <%=Year(Date)%> Grupo Stillos <span>Sistema administrativo interno</span></footer>
  </div>
</div>
<script>
(function () {
    var toggle = document.getElementById('ppvMenuToggle');
    var menu = document.getElementById('menuPrincipal');
    var previous = menu.querySelector('[aria-current="page"]');
    var current = menu.querySelector('a[href="ppv_compras_pedidos.asp"]');
    if (previous) previous.removeAttribute('aria-current');
    if (current) current.setAttribute('aria-current', 'page');
    toggle.addEventListener('click', function () {
        var expanded = menu.classList.toggle('is-open');
        toggle.setAttribute('aria-expanded', String(expanded));
    });
    document.addEventListener('keydown', function (event) {
        if (event.key === 'Escape' && menu.classList.contains('is-open')) {
            menu.classList.remove('is-open');
            toggle.setAttribute('aria-expanded', 'false');
            toggle.focus();
        }
    });
})();
function CarregaAjax(div,ts,p0,p1,p2,p3){
	//Função padrao para utilização de ajax nesta página.
	var caminho = "ppv_compras_pedidos_ajax.asp?ts="+ts+"&valor="+p0+'&valor1='+p1+'&valor2='+p2+'&valor3='+p3;
	//alert(caminho);
	ajax_loadContent(div,caminho,false,false,false);	
}


var solPendingAction;
function solConfirm(title, message, callback) {
    var dialog = document.getElementById('sol-confirm');
    document.getElementById('sol-confirm-title').textContent = title;
    document.getElementById('sol-confirm-message').textContent = message;
    solPendingAction = callback;
    dialog.returnValue = '';
    dialog.showModal();
}
function confirmaExclusaoEntradaErrada(id) {
    solConfirm('Excluir entrada?', 'O item voltará a ficar pendente de recebimento.', function () {
        var url = 'ppv_compras_pedidos_ajax.asp?ts=cee&valor=' + encodeURIComponent(id);
        ajax_loadContent('containerProdutoPedidoCompra_' + id, url, false, false, false);
    });
}
(function () {
    var dialog = document.getElementById('sol-confirm');
    dialog.addEventListener('close', function () {
        var action = solPendingAction;
        solPendingAction = null;
        if (dialog.returnValue === 'confirmar' && action) action();
    });
    var form = document.getElementById('sol-entry');
    var counter = document.getElementById('sol-count');
    if (!form || !counter) return;
    var feedback = document.getElementById('sol-feedback');
    function selected() {
        return form.querySelectorAll('.sol-select:checked').length;
    }
    form.addEventListener('change', function () {
        var count = selected();
        counter.textContent = count + (count === 1 ? ' item selecionado' : ' itens selecionados');
        feedback.hidden = true;
    });
    form.addEventListener('submit', function (event) {
        event.preventDefault();
        var count = selected();
        if (!count) {
            feedback.textContent = 'Selecione pelo menos um item pendente.';
            feedback.hidden = false;
            return;
        }
        solConfirm('Confirmar entrada?', 'Itens selecionados para recebimento: ' + count + '.', function () { form.submit(); });
    });
})();
</script>
</body>
</html>
