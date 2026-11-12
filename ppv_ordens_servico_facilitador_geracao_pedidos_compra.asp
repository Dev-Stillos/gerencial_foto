<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<%
Dim efetuar_transacao
efetuar_transacao = Cstr(""&Request.Form("campo1"))	

Dim busca_cod_produto

Dim sql, sqlWhere, sqlOrderBy, sqlGroupBy

'Padr�o Where
sqlWhere = sqlWhere&" and status_ordem_servico = '1' and necessita_gerar_pedido_compra = 'S' and pedido_compra_gerado IS NULL"

'Padr�o Order By 
sqlOrderBy = "cod_contrato, nome_formando, numero_ordem_servico desc, numero_ordem_servico_indice asc"

If Request.QueryString("filtro1") <> Empty Then
		
	Dim filtro_COD_CONTRATO
	Dim filtro_NOME_FORMANDO
	Dim filtro_COD_PRODUTO
	Dim filtro_ORDERNAR_POR
	
	filtro_COD_CONTRATO 	= Cstr(""&Request.QueryString("COD_CONTRATO"))
	filtro_NOME_FORMANDO 	= Cstr(""&Request.QueryString("NOME_FORMANDO"))
	filtro_COD_PRODUTO  	= Cstr(""&Request.QueryString("COD_PRODUTO"))
	filtro_ORDERNAR_POR  	= Cstr(""&Request.QueryString("ORDERNAR_POR"))
	
		If filtro_ordernar_por <> "" Then
			If filtro_ordernar_por = "2" Then
				sqlOrderBy = "nome_formando, numero_ordem_servico desc, numero_ordem_servico_indice asc"
			ElseIf filtro_ordernar_por = "3" Then				
				sqlOrderBy = "cod_contrato, nome_formando, numero_ordem_servico desc, numero_ordem_servico_indice asc"
			End If	
			
		End If
	
	If filtro_COD_PRODUTO <> "" Then
		sqlWhere = sqlWhere&" and cod_produto = "&filtro_COD_PRODUTO
	End If
	
	If filtro_COD_CONTRATO <> "" Then
			sqlWhere = sqlWhere&" AND COD_CONTRATO = '"&filtro_COD_CONTRATO&"'"
	End If
			
	If filtro_NOME_FORMANDO <> "" Then
			sqlWhere = sqlWhere&" AND NOME_FORMANDO LIKE '%"&filtro_NOME_FORMANDO&"%'"
	End If	
			

End If
'/////////////////////
'Padroniza sqlWhere

	If sqlWhere <> "" Then
		'Retirar o "AND "
		sqlWhere = Mid(sqlWhere,5,1000)
	End If


'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
If efetuar_transacao = "S" Then
	
	AbrirConexao()
		
		Dim cod_fornecedor, ultima_etapa, cod_ordem_servico, cod_ordem_servico_etapa
		
		'//////
		'Limpa base tempor�ria
		objConexao.Execute("DELETE FROM TB_PEDIDOS_COMPRAS_TEMPORARIA WHERE cod_usuario = "&sqlTexto(Session("cod_usuario")))	
		'//////		

		'1� ETAPA
		'/////////////////////////////////////////////////////////////////////		
		'1) Selecionar todas OS para triagem
		Set rs=objConexao.Execute("SELECT cod_produto, nome_produto, sum(quantidade) as qtd_necessario_comprar, nome_produto FROM VW_TB_ORDENS_SERVICO WHERE "&sqlWhere&" GROUP BY cod_produto order by nome_produto asc")
			If not rs.Eof then
				While not rs.Eof
						
					cod_produto = rs("cod_produto")				
					cod_fornecedor = Cstr(""&Request.Form("cod_fornecedor_"&cod_produto))
					
					'Response.Write(rs("nome_produto")&"<br />")
					
					'Verifica se foi selecionado algum fornecedor
					If cod_fornecedor <> "" Then
						
						sql = "SELECT * FROM VW_TB_ORDENS_SERVICO WHERE cod_produto = "&cod_produto&" and "&sqlWhere&" order by numero_ordem_servico desc, numero_ordem_servico_indice asc" 						
						'Seleciona todas as OS do produto atual
						'Response.Write(sql)
						Set rs2=objConexao.Execute(sql)
							
							If not rs2.Eof Then
								
								While not rs2.Eof
										
										cod_ordem_servico = rs2("cod_ordem_servico")
										ordem_servico_selecionada = Cstr(""&Request.Form("cod_ordem_servico_"&cod_ordem_servico))
										quantidade = rs2("quantidade")
										
										If ordem_servico_selecionada = "S" Then
											
											'Insere na base tempor�ria
											
												'Verifica se � �ltima etapa
												Set rs3=objConexao.Execute("Select count(*) as qtd_etapas From TB_ORDENS_SERVICO_ETAPAS where cod_ordem_servico = "&cod_ordem_servico&" group by cod_ordem_servico")
													
													If not rs3.Eof Then
														qtd_etapas = formatNumber(0&rs3("qtd_etapas"),0)
													End If
													
												rs3.Close
												Set rs3=Nothing
												
												if int(qtd_etapas) = int(rs2("numero_etapa")) Then
													ultima_etapa = "S"
												Else
													ultima_etapa = "N"
												End If
											
												sql = "INSERT INTO TB_PEDIDOS_COMPRAS_TEMPORARIA (cod_ordem_servico, cod_produto, cod_fornecedor, quantidade, cod_usuario,prazo_em_dias,cod_ordem_servico_etapa,ultima_etapa) VALUES ("&sqlTexto(cod_ordem_servico)&","&sqlTexto(cod_produto)&","&sqlTexto(cod_fornecedor)&","&sqlTexto(quantidade)&","&sqlTexto(Session("cod_usuario"))&","&rs2("prazo_em_dias")&","&rs2("cod_ordem_servico_etapa_atual")&","&sqlTexto(ultima_etapa)&")"
												
												objConexao.Execute(sql)
														
										End If
										
								rs2.MoveNext()
								Wend
								
							End If
							
						rs2.Close
						Set rs2=Nothing
											
					End If
									
				rs.MoveNext()
				Wend
				
			End if
			
		rs.close
		Set rs=nothing
		
		'2� ETAPA
		'/////////////////////////////////////////////////////////////////////
			
			Dim cod_pedido_compra, valor_custo_solicitado, data_previsao_entrega
			cod_pedido_compra = ""
			
			'Lista os fornecedores agrupados
			Set rs=objConexao.Execute("SELECT cod_fornecedor FROM TB_PEDIDOS_COMPRAS_TEMPORARIA WHERE cod_usuario = "&sqlTexto(Session("cod_usuario"))&" GROUP BY cod_fornecedor")
				If not rs.Eof then
				
					While not rs.Eof
					
						cod_fornecedor = rs("cod_fornecedor")		
						
						'1) Faz o insert do Pedido de Compras
						'//////
						objConexao.Execute("INSERT INTO TB_PEDIDOS_COMPRAS (cod_fornecedor,cadastro_cod_usuario,data_aprovado,aprovado_cod_usuario,status_pedido_compra,cod_usuario,modificado) VALUES ("&sqlTexto(cod_fornecedor)&","&sqlTexto(Session("cod_usuario"))&",NOW(),"&sqlTexto(Session("cod_usuario"))&",2,"&sqlTexto(Session("cod_usuario"))&",NOW())")
						'Obs: Status 2 = aprovado e n�o enviado para fornecedor
						'Busca a chave criada
						Set rs2=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_PEDIDOS_COMPRAS LIMIT 1")
							If not rs2.Eof Then
								cod_pedido_compra = rs2("ULTIMO_ID")
							End if 
						rs2.Close
						Set rs2=Nothing																
						
						'Atribui o n�mero do pedido de compras
						Set rs2=objConexao.Execute("SELECT IFNULL(MAX(numero_pedido_compra),0)+1 as numero_pedido_compra FROM TB_PEDIDOS_COMPRAS WHERE status_pedido_compra <> 0")
							If not rs2.Eof Then
								objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS SET numero_pedido_compra = "&sqlNumero(rs2(0))&" WHERE TB_PEDIDOS_COMPRAS.cod_pedido_compra ="&sqlTexto(cod_pedido_compra))
							End if 
						rs2.Close
						Set rs2=Nothing		
						'//////
						
						
						'###////////////////////////////////////////
						'### - N�O AGRUPA MAIS OS PRODUTOS
'						'2) Agrupa os produtos do fornededor
'						sql = "SELECT cod_produto, sum(quantidade) as quantidade, prazo_em_dias FROM TB_PEDIDOS_COMPRAS_TEMPORARIA WHERE cod_usuario = "&sqlTexto(Session("cod_usuario"))&" and cod_fornecedor = "&cod_fornecedor&" GROUP BY cod_produto"
'						'Response.Write(sql)
'						Set rs2=objConexao.Execute(sql)				
'								While not rs2.Eof
'									
'									cod_produto = rs2("cod_produto")
'									'3) Insere no pedido de compra produto
'										
'										'///
'										'Valor de custo do produto do fornecedor
'										valor_custo_solicitado  = 0
'										Set rs3=objConexao.Execute("SELECT valor_custo FROM TB_FORNECEDOR_PRODUTOS WHERE cod_fornecedor = "&sqlTexto(cod_fornecedor)&" AND cod_produto = "&sqlTexto(cod_produto))
'											
'											If not rs3.Eof Then
'												valor_custo_solicitado = rs3(0)
'											End If
'										
'										rs3.Close	
'										Set rs3=Nothing
'																			
'										'Previs�o de Entrega 
'										data_previsao_entrega = DateAdd("d",rs2("prazo_em_dias"),date())
'										'///
'									
'										objConexao.Execute("INSERT INTO TB_PEDIDOS_COMPRAS_PRODUTOS (cod_pedido_compra, cod_produto, quantidade_solicitada, valor_custo_solicitado, data_previsao_entrega, status_produto, cod_usuario, modificado) VALUES ("&sqlTexto(cod_pedido_compra)&",'"&cod_produto&"',"&sqlNumero(rs2("quantidade"))&","&sqlNumero(valor_custo_solicitado)&","&sqlData(data_previsao_entrega)&",'P',"&sqlTexto(Session("cod_usuario"))&",NOW())")
'									
'								rs2.MoveNext()
'								Wend
'						
'						rs2.Close
'						Set rs2=Nothing
						
						'###////////////////////////////////////////
						
						'4) Listar todos os produtos do fornecedor SEM Group By
						Set rs2=objConexao.Execute("SELECT * FROM TB_PEDIDOS_COMPRAS_TEMPORARIA WHERE cod_usuario = "&sqlTexto(Session("cod_usuario"))&" and cod_fornecedor = "&cod_fornecedor)				
								While not rs2.Eof
									
									cod_ordem_servico = rs2("cod_ordem_servico")
									cod_ordem_servico_etapa = rs2("cod_ordem_servico_etapa")									
									cod_produto = rs2("cod_produto")
										
'										'///
'										'Valor de custo do produto do fornecedor
										valor_custo_solicitado  = 0
										Set rs3=objConexao.Execute("SELECT valor_custo FROM TB_FORNECEDOR_PRODUTOS WHERE cod_fornecedor = "&sqlTexto(cod_fornecedor)&" AND cod_produto = "&sqlTexto(cod_produto))
											
											If not rs3.Eof Then
												valor_custo_solicitado = rs3(0)
											End If
										
										rs3.Close	
										Set rs3=Nothing
'																			
'										'Previs�o de Entrega 
										data_previsao_entrega = DateAdd("d",rs2("prazo_em_dias"),date())
										'///
									
										objConexao.Execute("INSERT INTO TB_PEDIDOS_COMPRAS_PRODUTOS (cod_pedido_compra, cod_produto, quantidade_solicitada, valor_custo_solicitado, data_previsao_entrega, cod_usuario, modificado,cod_ordem_servico_etapa) VALUES ("&sqlTexto(cod_pedido_compra)&",'"&cod_produto&"',"&sqlNumero(rs2("quantidade"))&","&sqlNumero(valor_custo_solicitado)&","&sqlData(data_previsao_entrega)&","&sqlTexto(Session("cod_usuario"))&",NOW(),"&sqlTexto(cod_ordem_servico_etapa)&")")
									
										objConexao.Execute("UPDATE TB_ORDENS_SERVICO_ETAPAS SET pedido_compra_gerado = "&sqlNumero(cod_pedido_compra)&" WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&sqlTexto(cod_ordem_servico_etapa))
									
									'////
									'N�O FINALIZA OS, SOMENTE QUANDO FOR ENTREGAR
									
									If Cstr(rs2("ultima_etapa")) = "S" Then
									'Dever� finalizar OS
											'sql = "UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_termino = NOW() WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&sqlTexto(cod_ordem_servico_etapa)
											'objConexao.Execute(sql)
											'sql = "UPDATE TB_ORDENS_SERVICO SET status_ordem_servico = 2, data_finalizado = NOW(), cod_usuario_finalizou = "&Session("cod_usuario")&" WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)
											'objConexao.Execute(sql)
									Else
										
											'Finaliza a etapa atual
											'sql = "UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_termino = NOW() WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&sqlTexto(cod_ordem_servico_etapa)		
											'objConexao.Execute(sql)
											'Inicia a pr�xima Etapa
											'Set rs3=objConexao.Execute("SELECT cod_ordem_servico_etapa FROM TB_ORDENS_SERVICO_ETAPAS WHERE cod_ordem_servico = "&cod_ordem_servico&" and numero_etapa IN(SELECT (numero_etapa+1) FROM VW_TB_ORDENS_SERVICO WHERE cod_ordem_servico = "&cod_ordem_servico&")")
												
											'	If not rs3.Eof Then
											'		sql = "UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_inicio = NOW() WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&rs3("cod_ordem_servico_etapa")
											'		objConexao.Execute(sql)
											'		sql = "UPDATE TB_ORDENS_SERVICO SET cod_ordem_servico_etapa_atual = "&rs3("cod_ordem_servico_etapa")&" WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)
											'		objConexao.Execute(sql)
											'	End If
												
											'rs3.Close
											'Set rs3=Nothing
										
									End If		
									
									'///							
									
								rs2.MoveNext()
								Wend
						
						rs2.Close
						Set rs2=Nothing						
			
					rs.MoveNext()
					Wend
				
				End if
			
			rs.close
			Set rs=nothing		
		
	FecharConexao()
	Response.Redirect("ppv_compras_pedidos.asp?codmsg=601")
	
End If
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

%>
<%Response.Buffer = False%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Gerar pedidos de compra por OS | Portal Stillos</title>
<link rel="stylesheet" href="css/vendor/bootstrap.min.css">
<link rel="stylesheet" href="css/dashboard.css">
<style>
.fac-page .dashboard-nav a.active:not([aria-current]) { background:transparent; color:#d4d5e1; }
.fac-page .dashboard-nav a[aria-current="page"] { background:#b20b61; color:#fff; font-weight:600; }
.fac-page .fac-panel { height:auto; margin-bottom:16px; padding:clamp(14px,3vw,24px); }
.fac-filters { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; margin:14px 0; }
.fac-filters label, .fac-supplier-field label { display:block; font-weight:650; }
.fac-filters .form-select, .fac-filters .form-control, .fac-supplier { margin-top:5px; }
.fac-actions { display:flex; flex-wrap:wrap; gap:8px; }
.fac-product summary { cursor:pointer; color:#242738; }
.fac-product summary strong { margin-right:6px; font-size:1.08em; }
.fac-count { display:inline-block; margin:4px 2px; padding:4px 8px; border-radius:6px; background:#f3edf6; color:#51436e; font-weight:650; }
.fac-supplier-field { max-width:420px; margin:14px 0; }
.fac-table { margin:0; }
.fac-table th, .fac-table td { padding:9px 7px; vertical-align:top; }
.fac-table th { background:#f4f1f7; color:#42364e; }
.fac-table tr:has(.fac-order-check:checked) > td { background:#f5fbf7; }
.fac-person small { display:block; color:#57566a; }
.fac-deadline-alert { display:inline-block; padding:3px 7px; border-radius:5px; background:#fce9ee; color:#8e2045; font-weight:650; }
.fac-review-bar { position:sticky; bottom:0; z-index:2; display:flex; flex-wrap:wrap; align-items:center; justify-content:flex-end; gap:8px; padding:12px; border:1px solid #e4dfeb; border-radius:10px; background:#fff; }
.fac-feedback { flex-basis:100%; margin:0; color:#8e2045; }
.fac-dialog { width:min(600px,calc(100% - 24px)); max-height:calc(100dvh - 24px); padding:24px; border:1px solid #e4dfeb; border-radius:14px; color:var(--ink); }
.fac-dialog::backdrop { background:#24273899; }
.fac-dialog .fac-actions { justify-content:flex-end; }
@media(max-width:1100px) {
  .fac-filters { grid-template-columns:repeat(2,minmax(0,1fr)); }
  .fac-table, .fac-table tbody { display:block; }
  .fac-table thead { position:absolute; width:1px; height:1px; overflow:hidden; clip-path:inset(50%); }
  .fac-table tbody tr { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); border-bottom:1px solid #e4dfeb; }
  .fac-table td { display:block; min-width:0; border:0; overflow-wrap:anywhere; }
  .fac-table td:nth-child(3), .fac-table td:last-child { grid-column:1/-1; }
  .fac-table td::before { content:attr(data-label); display:block; color:#57566a; font-size:.85em; font-weight:650; }
}
@media(max-width:480px) { .fac-filters { grid-template-columns:1fr; } }
</style>
</head>
<body class="fac-page">
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
        <div><h1>Gerar pedidos de compra por OS</h1><p>Selecione as ordens que aguardam compra e confira o fornecedor de cada produto.</p></div>
        <a href="ppv_compras_pedidos.asp" class="btn btn-outline-secondary">Voltar aos pedidos</a>
      </div>
      <div class="search-panel"><div><strong>Busca no portal</strong></div><!-- #include file="includes/dashboard_busca.asp" --></div>
      <div class="session-notices"><!-- #include file="includes/dashboard_avisos.asp" --></div>

<!-- #include file="includes/compras_sub_menu.asp" -->
<section class="dashboard-panel fac-panel" aria-labelledby="fac-filter-title">
  <h2 id="fac-filter-title">Filtrar ordens</h2>
  <form method="get" action="ppv_ordens_servico_facilitador_geracao_pedidos_compra.asp">
    <input type="hidden" name="filtro1" value="1">
    <div class="fac-filters">
      <div><label for="fac-product">Produto</label><select name="COD_PRODUTO" id="fac-product" class="form-select">
		          <option value=''>Selecione</option>
		          <%
				  AbrirConexao()
				  
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

FecharConexao()
%>
		          </select></div>
      <div><label for="fac-contract">C&oacute;digo do contrato</label><input class="form-control" id="fac-contract" name="COD_CONTRATO" value="<%=Server.HTMLEncode(filtro_COD_CONTRATO & "")%>"></div>
      <div><label for="fac-name">Nome do formando</label><input class="form-control" id="fac-name" name="NOME_FORMANDO" value="<%=Server.HTMLEncode(filtro_NOME_FORMANDO & "")%>"></div>
      <div><label for="fac-sort">Ordenar por</label><select name="ordernar_por" id="fac-sort" class="form-select">
		          <option <%If Cstr(filtro_ordernar_por) = "2" Then%> selected="selected" <%End If%> value="2">Formando</option>
		          <option <%If Cstr(filtro_ordernar_por) = "3" or Cstr(filtro_ordernar_por) = "" Then%> selected="selected" <%End If%> value="3">Institui&ccedil;&atilde;o de Ensino</option>
		          </select></div>
    </div>
    <div class="fac-actions">
      <button class="btn btn-primary" name="filtro2" value="1" type="submit">Aplicar filtros</button>
      <a class="btn btn-outline-secondary" href="ppv_ordens_servico_facilitador_geracao_pedidos_compra.asp">Limpar filtros</a>
    </div>
  </form>
</section>
<%		
		Dim temp
		'Neste momento corrige o sqlWhere somente para o primeiro rs
		If sqlWhere <> "" Then
			temp = "WHERE "
		End If
		
		sql = "SELECT cod_produto, nome_produto, sum(quantidade) as qtd_necessario_comprar FROM VW_TB_ORDENS_SERVICO "&temp&sqlWhere&" GROUP BY cod_produto order by nome_produto asc"
		
		'Response.Write(sql)
		
		AbrirConexao()
		
		Set rs=objConexao.Execute(sql)
		
		If not rs.Eof Then
		%>

<p class="alert alert-warning">Ao gerar a compra, a etapa pendente das OS inclu&iacute;das ser&aacute; finalizada.</p>
<form id="fac-orders" name="FORMULARIO_PADRAO" action="ppv_ordens_servico_facilitador_geracao_pedidos_compra.asp" method="post">
  <input name="campo1" value="S" type="hidden">
  <%    'Neste momento que j� passou o group by corrige o sqlWhere para o rs2
    If sqlWhere <> "" Then
        sqlWhere = " AND "&sqlWhere
    End If
    
    While not rs.Eof
    %>
 

  <details class="dashboard-panel fac-panel fac-product">
    <summary>
      <strong><%=Server.HTMLEncode(rs("nome_produto") & "")%></strong>
      <span class="fac-count"><%=rs("qtd_necessario_comprar")%> unidades pendentes</span>
      <span class="fac-count"><output class="fac-quantity">0</output> selecionadas</span>
    </summary>
    <div class="fac-supplier-field"><label>Fornecedor e custo unit&aacute;rio
    <select class="form-select fac-supplier" name="cod_fornecedor_<%=rs("cod_produto")%>">
                                    <%
                                            'Seleciona todos os funcion&aacute;rios ativos, exceto o usu&aacute;rio atual.
                                            Set rs3=objConexao.Execute("SELECT TB_FORNECEDORES.NOME_FANTASIA, TB_FORNECEDOR_PRODUTOS.* FROM  TB_FORNECEDOR_PRODUTOS INNER JOIN TB_FORNECEDORES ON  TB_FORNECEDOR_PRODUTOS.cod_fornecedor = TB_FORNECEDORES.cod_fornecedor WHERE TB_FORNECEDORES.ATIVADO = 1 and TB_FORNECEDOR_PRODUTOS.cod_produto = "&sqlTexto(rs("cod_produto"))&" order by TB_FORNECEDOR_PRODUTOS.valor_custo asc")
 
 											If not rs3.Eof Then
 
                                            	While not rs3.Eof%>
		    	                                <option value="<%=rs3("cod_fornecedor")%>"><%=Server.HTMLEncode(rs3("NOME_FANTASIA") & "")%> - <%=formatCurrency(rs3("valor_custo"))%></option>
        	                            		<%
                	                            rs3.MoveNext()
                    	                        Wend	
                                            
											Else		    
											%>
											<option value="">Nenhum fornecedor cadastrado</option>
											<%
											End If
                                        
                                            rs3.Close
                                            Set rs3=Nothing
                                        %>
                                  </select>
    </label></div>
    <input type="hidden" class="fac-quantity-input" value="0" name="qtd_solicitada_<%=rs("cod_produto")%>">
    <table class="table fac-table" role="table" aria-label="Ordens do produto">
      <thead role="rowgroup"><tr role="row"><th scope="col" role="columnheader">Selecionar</th><th scope="col" role="columnheader">OS</th><th scope="col" role="columnheader">Formando</th><th scope="col" role="columnheader">Qtd.</th><th scope="col" role="columnheader">In&iacute;cio da etapa</th><th scope="col" role="columnheader">Prazo final</th><th scope="col" role="columnheader">Respons&aacute;vel</th></tr></thead>
      <tbody role="rowgroup">
      <%	 
		   sql = "SELECT * FROM VW_TB_ORDENS_SERVICO WHERE cod_produto = "&rs("cod_produto")&sqlWhere&" ORDER BY "&sqlOrderBy 
           'Response.Write(sql)
           Set rs2=objConexao.Execute(sql)
           
           While not rs2.Eof
           %>
           		<tr role="row">
  <td role="cell" data-label="Selecionar"><input type="checkbox" class="form-check-input fac-order-check" value="S" name="cod_ordem_servico_<%=rs2("cod_ordem_servico")%>" data-quantity="<%=rs2("quantidade")%>" aria-label="Selecionar OS <%=rs2("cod_ordem_servico")%>"></td>
  <td role="cell" data-label="OS"><strong><%=moduloPPVOSMascara(rs2("cod_ordem_servico"),"",true)%></strong></td>
  <td role="cell" data-label="Formando"><%
						sql = 	" SELECT "&_
								" TB_PEDIDOS_FORMANDOS.COD_FORMANDO, "&_
								" TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, "&_
								" TB_FICHAS_IDENTIFICACAO.COD_CONTRATO, "&_
								" TB_ENTIDADES.DESCRICAO, "&_
								" TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO, "&_
								" TB_FICHAS_IDENTIFICACAO.NOME_CURSO "&_
								" FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON "&_
								" TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO "&_
								" INNER JOIN TB_ENTIDADES ON(TB_FICHAS_IDENTIFICACAO.COD_CONTRATO = TB_ENTIDADES.COD_CONTRATO)"&_
								" WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&rs2("cod_pedido")&" ORDER BY NOME_FORMANDO"
											
						           Set rs3=objConexao.Execute(sql)
								   %>
											
                                    	<%
                                    	While not rs3.Eof
										%>
                                        	<div class="fac-person"><strong><%=Server.HTMLEncode(rs3("NOME_FORMANDO") & "")%></strong><small><%=Server.HTMLEncode(rs3("DESCRICAO") & "")%></small></div>
                						<%
										rs3.MoveNext()
										Wend
										
										%>
                                        
                                        <%
									
									rs3.Close
									Set rs3=Nothing
									%></td>
  <td role="cell" data-label="Qtd."><%=rs2("quantidade")%></td>
  <td role="cell" data-label="In&iacute;cio da etapa"><%=ArrumaData(rs2("data_inicio"))%></td>
  <td role="cell" data-label="Prazo final">
    <%If IsDate(rs2("prazo_geral")) Then%>
    <span class="fac-deadline<%If DateDiff("d",now(),rs2("prazo_geral")) < 10 Then%> fac-deadline-alert<%End If%>"><%=ArrumaData(rs2("prazo_geral"))%></span>
    <%Else%>&mdash;<%End If%>
  </td>
  <td role="cell" data-label="Respons&aacute;vel"><%=nomeUsuario(rs2("cod_usuario_responsavel"),true)%></td>
</tr>
<%
    rs2.MoveNext()
    Wend
    rs2.Close
    Set rs2 = Nothing
%>

      </tbody>
    </table>
  </details>
  <%rs.MoveNext()
    Wend%>
  <div class="fac-review-bar">
    <span id="fac-selection" role="status">0 OS selecionadas</span>
    <button class="btn btn-primary" type="submit" aria-haspopup="dialog">Revisar compra</button>
    <p class="fac-feedback" id="fac-feedback" role="alert" hidden></p>
  </div>
</form>
<dialog class="fac-dialog" id="fac-confirm" aria-labelledby="fac-confirm-title">
  <h2 id="fac-confirm-title">Confirmar gera&ccedil;&atilde;o?</h2>
  <p>A gera&ccedil;&atilde;o finaliza a etapa pendente das ordens inclu&iacute;das.</p>
  <ul id="fac-review-list"></ul>
  <form method="dialog" class="fac-actions">
    <button class="btn btn-outline-secondary" value="cancelar" autofocus>Cancelar</button>
    <button class="btn btn-primary" value="gerar">Gerar pedidos</button>
  </form>
</dialog>
<%Else%>
<section class="dashboard-panel fac-panel" role="status"><h2>Nenhuma OS aguardando compra</h2><p class="mb-0">Confira os filtros para buscar outros itens.</p></section>
<%End If
rs.Close
Set rs = Nothing
FecharConexao()
%>
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

(function () {
    var form = document.getElementById('fac-orders');
    if (!form) return;
    var dialog = document.getElementById('fac-confirm');
    var feedback = document.getElementById('fac-feedback');
    var selection = document.getElementById('fac-selection');
    var list = document.getElementById('fac-review-list');
    var products = form.querySelectorAll('.fac-product');

    function updateProduct(product) {
        var checked = product.querySelectorAll('.fac-order-check:checked');
        var units = Array.from(checked).reduce(function (sum, input) {
            return sum + (Number.parseInt(input.dataset.quantity, 10) || 0);
        }, 0);
        product.querySelector('.fac-quantity').value = units;
        product.querySelector('.fac-quantity-input').value = units;
        return checked.length;
    }

    function updateSelection() {
        var count = form.querySelectorAll('.fac-order-check:checked').length;
        selection.textContent = count + (count === 1 ? ' OS selecionada' : ' OS selecionadas');
    }

    products.forEach(updateProduct);
    updateSelection();
    form.addEventListener('change', function (event) {
        if (event.target.matches('.fac-order-check')) {
            updateProduct(event.target.closest('.fac-product'));
            updateSelection();
        }
        feedback.hidden = true;
    });

    form.addEventListener('submit', function (event) {
        event.preventDefault();
        list.replaceChildren();
        for (var product of products) {
            if (!updateProduct(product)) continue;
            var supplier = product.querySelector('.fac-supplier');
            if (!supplier.value) {
                product.open = true;
                feedback.textContent = 'Selecione um fornecedor para cada produto com OS selecionadas.';
                feedback.hidden = false;
                supplier.focus();
                return;
            }
            var item = document.createElement('li');
            item.textContent = product.querySelector('summary strong').textContent.trim() + ' \u2014 ' +
                product.querySelector('.fac-quantity-input').value + ' un. \u2014 ' +
                supplier.options[supplier.selectedIndex].text.trim();
            list.appendChild(item);
        }
        if (!list.children.length) {
            feedback.textContent = 'Selecione pelo menos uma OS para continuar.';
            feedback.hidden = false;
            return;
        }
        feedback.hidden = true;
        dialog.returnValue = '';
        dialog.showModal();
    });

    dialog.addEventListener('close', function () {
        if (dialog.returnValue === 'gerar') form.submit();
    });
})();
</script>
</body>
</html>
