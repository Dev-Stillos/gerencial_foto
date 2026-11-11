<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="files/includes/functions_local.asp" -->
<%
Dim caminho_pagina, calendario_necessario
caminho_pagina = "ppv_compras_pedidos.asp"
calendario_necessario = "SIM"
%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Pedidos de compra | Portal Stillos</title>
<link rel="stylesheet" href="css/vendor/bootstrap.min.css">
<link rel="stylesheet" href="css/dashboard.css">
<!-- #include file="files/includes/scripts.asp" -->
<style>
.purchase-page .dashboard-nav a.active:not([aria-current]) { background:transparent; color:#d4d5e1; }
.purchase-page .dashboard-nav a[aria-current="page"] { background:#b20b61; color:#fff; font-weight:600; }
.purchase-page .purchase-panel { height:auto; margin-bottom:18px; padding:clamp(16px,3vw,26px); }
.purchase-filters { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:14px; margin:14px 0; }
.purchase-filters label { display:block; font-weight:650; margin-bottom:5px; }
.purchase-date, .purchase-actions { display:flex; align-items:center; flex-wrap:wrap; gap:8px; }
.purchase-date { flex-wrap:nowrap; }
.purchase-date input { min-width:0; }
.purchase-date button { padding:8px; }
.purchase-results-heading { display:flex; align-items:center; justify-content:space-between; gap:10px; flex-wrap:wrap; margin-bottom:14px; }
.purchase-results-heading span { color:#57566a; }
.purchase-table-scroll { max-width:100%; overflow-x:auto; }
.purchase-table { width:max-content; min-width:100%; margin-bottom:0; }
.purchase-table th, .purchase-table td { padding:12px 10px; vertical-align:middle; overflow-wrap:normal; word-break:normal; }
.purchase-table th { white-space:nowrap; }
.purchase-table td:is(:nth-child(1), :nth-child(2), :nth-child(6), :nth-child(7)) { white-space:nowrap; }
.purchase-table td:nth-child(3) { max-width:260px; min-width:140px; }
.purchase-table td:nth-child(9) { max-width:210px; }
.purchase-table td:nth-child(10) .purchase-status { white-space:nowrap; }
.purchase-table thead th { background:#f4f1f7; color:#42364e; }
.purchase-table td:first-child { font-weight:700; }
.purchase-manual { display:block; color:#65546e; font-weight:600; }
.purchase-status { display:inline-block; padding:3px 8px; border-radius:6px; background:#f0edf8; color:#51436e; font-weight:650; }
.purchase-status.is-delivered, .purchase-status.is-on-time { background:#e7f3ec; color:#286143; }
.purchase-status.is-late { background:#fce9ee; color:#8e2045; }
.purchase-row-actions { gap:4px; }
.purchase-action { background:#fff; cursor:pointer; font-family:inherit; display:inline-block; padding:4px 6px; border:1px solid #e4dfeb; border-radius:5px; text-decoration:none; font-size:.87em; color:#88134f; }
.purchase-action:hover { background:#f7edf3; color:#88134f; }
.purchase-pagination { display:flex; justify-content:center; align-items:center; flex-wrap:wrap; gap:8px; margin-top:18px; padding-top:16px; border-top:1px solid #e4dfeb; }
.purchase-pagination span { color:#57566a; }
.purchase-row-actions .purchase-action { white-space:nowrap; }
@media(max-width:1200px) {
  .purchase-table, .purchase-table tbody { display:block; width:100%; }
  .purchase-table thead { position:absolute; width:1px; height:1px; overflow:hidden; clip-path:inset(50%); }
  .purchase-row { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); margin-bottom:14px; border:1px solid #e4dfeb; border-radius:10px; overflow:hidden; }
  .purchase-row td:nth-child(1), .purchase-row td:nth-child(9) { background:#f4f1f7; padding:12px; }
  .purchase-row td:nth-child(1) { grid-row:1; grid-column:1; }
  .purchase-row td:nth-child(9) { grid-row:1; grid-column:2; }
  .purchase-row td:nth-child(2), .purchase-row td:nth-child(3), .purchase-row td:nth-child(6) { grid-column:1/-1; }
  .purchase-row td:nth-child(3), .purchase-row td:nth-child(9) { max-width:none; min-width:0; }
  .purchase-row td:nth-child(2) > div, .purchase-row td:nth-child(6) > div { overflow-x:auto; }
  .purchase-row td:nth-child(7) { white-space:normal; }
  .purchase-row td:nth-child(2) { font-weight:700; }
  .purchase-row td:nth-child(4), .purchase-row td:nth-child(5) { background:#faf9fc; }
  .purchase-row td > div[align] { text-align:left; }
  .purchase-row td:last-child { padding:10px; }
  .purchase-row .purchase-action { display:inline-flex; align-items:center; min-height:36px; padding:6px 9px; }
  .purchase-row td { display:block; border:0; }
  .purchase-row td::before { content:attr(data-label); display:block; color:#66586e; font-size:.85em; font-weight:700; }
  .purchase-row td:last-child { grid-column:1/-1; border-top:1px solid #e4dfeb; }
}
@media(max-width:640px) {
  .purchase-filters { grid-template-columns:repeat(2,minmax(0,1fr)); }
}
@media(max-width:400px) {
  .purchase-filters { grid-template-columns:minmax(0,1fr); }
}
</style>
</head>
<body class="purchase-page">
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
        <div><h1>Pedidos de compra</h1><p>Acompanhe fornecedores, recebimentos e prazos dos pedidos.</p></div>
        <a href="adm_index.asp" class="btn btn-outline-secondary">Voltar ao painel</a>
      </div>
      <div class="search-panel"><div><strong>Busca no portal</strong></div><!-- #include file="includes/dashboard_busca.asp" --></div>
      <div class="session-notices"><!-- #include file="includes/dashboard_avisos.asp" --></div>

<!-- #include file="includes/compras_sub_menu.asp" -->
<%
	 
sub segundaViaRecibo(cod_pedido_compra, conexao_aberta)

	if not conexao_aberta Then
		AbrirConexao()
	end if

	objConexao.Execute("DELETE FROM TB_PEDIDOS_COMPRAS_PRODUTOS_TEMPORARIA WHERE cod_usuario = "&sqlNumero(session("cod_usuario")))

	objConexao.Execute("INSERT INTO TB_PEDIDOS_COMPRAS_PRODUTOS_TEMPORARIA (cod_pedido_compra_produto, cod_usuario) SELECT cod_pedido_compra_produto, "&session("cod_usuario")&" FROM TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra = "&sqlNumero(cod_pedido_compra) &" AND status_produto = 'E' " )

	if not conexao_aberta Then
		fecharConexao()
	end if


	Response.Redirect("ppv_compras_pedidos_produtos_solicitados_recibo_entrega.asp")


end sub

dim cod_pedido_compra_recibo, acao_recibo

cod_pedido_compra_recibo = Request.QueryString("cod_pedido_compra_recibo")
acao_recibo = Request.QueryString("acao_recibo")

if cod_pedido_compra_recibo <> "" and acao_recibo = "segunda_via_recibo" and validaUsuario(2,,83,true) then 
	call segundaViaRecibo(cod_pedido_compra_recibo, false)
end if

	AbrirConexao()
	
	Dim validacao_para_listar_qualquer_pedido
	validacao_para_listar_qualquer_pedido = validaUsuario(2,,73,true)
	'/////////////////////////////////////////////////////////////
	'Rotina de corre&ccedil;&atilde;o de pedidos pendentes
	
		'1) Lista dos os pedidos de compras pendentes
		Set rs=objConexao.Execute("SELECT * FROM TB_PEDIDOS_COMPRAS WHERE status_pedido_compra = 2")
			
			While not rs.Eof
				
				'2) Verifica se h&aacute; itens pendentes realmente
				Set rs2=objConexao.Execute("SELECT * FROM TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra = "&rs("cod_pedido_compra")&" AND status_produto = 'P'")
					
					If not rs2.Eof Then
						
						pode_finalizar = true
						
						While not rs2.Eof
							
							If Cstr(""&rs2("cod_ordem_servico_etapa")) <> "" Then
							
								'3) Verifica se o item pendente pode ser excluido, pois a OS foi excluida
								Set rs3=objConexao.Execute("SELECT * FROM TB_ORDENS_SERVICO_ETAPAS WHERE cod_ordem_servico_etapa = "&rs2("cod_ordem_servico_etapa"))
									
									If not rs3.Eof Then
									'Encontrou ent&atilde;o n&atilde;o pode excluir, pois esta pendente e h&aacute; relacionamento
										pode_finalizar = false
									Else
									'Cancela, pois OS n&atilde;o existe mais.
										objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS_PRODUTOS SET status_produto = 'C' WHERE cod_pedido_compra_produto = "&rs2("cod_pedido_compra_produto"))
									End If
								
								rs3.Close
								Set rs3=Nothing
							
							Else
								pode_finalizar = false
							End If
							
						rs2.MoveNext()
						Wend
						
						If pode_finalizar Then
							'Response.Write("Baixa o pedido de compra => "&rs("numero_pedido_compra")&" OS n&atilde;o existe mais<br />")		
							objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS set status_pedido_compra = 3 WHERE cod_pedido_compra = "&rs("cod_pedido_compra"))
							
							
						End If
					
					Else
						
						'Baixa o pedido de compra
						'Response.Write("Baixa o pedido de compra => "&rs("numero_pedido_compra")&"<br />")
						objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS set status_pedido_compra = 3 WHERE cod_pedido_compra = "&rs("cod_pedido_compra"))
						
					End If
				
				rs2.Close
				Set rs2=Nothing
				
			rs.MoveNext()
			Wend
		
		rs.Close
		Set rs=Nothing
	
	'/////////////////////////////////////////////////////////////
	
	Dim sql, sqlWhere, sqlOrderBy
	sql = "SELECT * FROM VW_TB_PEDIDOS_COMPRAS WHERE status_pedido_compra > 0 "
	sqlOrderBy = " order by numero_pedido_compra desc"
	
	Dim status_pedido_compra, diferenca_prazo_entrega

	'/////////////
	'Filtro
		Dim filtro_status_pedido_compra
		Dim filtro_cod_fornecedor
		Dim filtro_cadastro_cod_usuario
		Dim filtro_NUMERO_PEDIDO_COMPRAS
		Dim b_DATA_CADASTRO_INICIO
		Dim b_DATA_CADASTRO_FIM
		
		filtro_NUMERO_PEDIDO_COMPRAS	= Cstr(""&Request.QueryString("NUMERO_PEDIDO_COMPRAS"))
		filtro_status_pedido_compra		= Cstr(""&Request.QueryString("status_pedido_compra"))
		filtro_cod_fornecedor			= Cstr(""&Request.QueryString("cod_fornecedor"))
		filtro_cadastro_cod_usuario		= Cstr(""&Request.QueryString("cadastro_cod_usuario"))
		b_DATA_CADASTRO_INICIO			=	Request.QueryString("DATA_CADASTRO_INICIO")
		b_DATA_CADASTRO_FIM				=	Request.QueryString("DATA_CADASTRO_FIM")
		
		If Not validacao_para_listar_qualquer_pedido Then
			sqlWhere = sqlWhere&" AND cadastro_cod_usuario = '"&Session("cod_usuario")&"'"
		End If
		
		If Request.QueryString("filtro") <> "" Then
		
			If filtro_NUMERO_PEDIDO_COMPRAS <> "" Then
			'Filtro preferencial, limpa os demais
				
				filtro_cod_fornecedor = ""
				filtro_cadastro_cod_usuario = ""
				b_DATA_CADASTRO_INICIO = ""
				b_DATA_CADASTRO_FIM	= ""
				
					sqlWhere = sqlWhere&" AND numero_pedido_compra = '"&filtro_NUMERO_PEDIDO_COMPRAS&"'"
					
					If filtro_status_pedido_compra <> "" Then
						sqlWhere = sqlWhere&" AND status_pedido_compra = '"&filtro_status_pedido_compra&"'"
					End If
			
			Else
		
				If filtro_status_pedido_compra <> "" Then
					sqlWhere = sqlWhere&" AND status_pedido_compra = '"&filtro_status_pedido_compra&"'"
				End If
				
				if filtro_cod_fornecedor <> "" Then
					sqlWhere = sqlWhere&" AND cod_fornecedor = '"&filtro_cod_fornecedor&"'"
				End if
				
				if filtro_cadastro_cod_usuario <> "" Then
					sqlWhere = sqlWhere&" AND cadastro_cod_usuario = '"&filtro_cadastro_cod_usuario&"'"
				End if			
				
				
				If b_DATA_CADASTRO_INICIO <> Empty or b_DATA_CADASTRO_FIM <> Empty Then
							
							
							If b_DATA_CADASTRO_INICIO <> Empty Then
								b_DATA_CADASTRO_INICIO = Day(b_DATA_CADASTRO_INICIO)&"/"&Month(b_DATA_CADASTRO_INICIO)&"/"&Year(b_DATA_CADASTRO_INICIO)&" 00:00:00"
							End If
		
							If b_DATA_CADASTRO_FIM <> Empty Then
								b_DATA_CADASTRO_FIM = Day(b_DATA_CADASTRO_FIM)&"/"&Month(b_DATA_CADASTRO_FIM)&"/"&Year(b_DATA_CADASTRO_FIM)&"  23:59:59"
							End If					
																	
							If b_DATA_CADASTRO_INICIO <> Empty and b_DATA_CADASTRO_FIM <> Empty Then
								sqlWhere = sqlWhere&" AND DATA_CADASTRO BETWEEN "&sqlData(b_DATA_CADASTRO_INICIO)&" AND "&sqlData(b_DATA_CADASTRO_FIM)
							Elseif b_DATA_CADASTRO_INICIO = Empty and b_DATA_CADASTRO_FIM <> Empty Then
								sqlWhere = sqlWhere&" AND DATA_CADASTRO <= "&sqlData(b_DATA_CADASTRO_FIM)
							Elseif b_DATA_CADASTRO_INICIO <> Empty and b_DATA_CADASTRO_FIM = Empty Then
								sqlWhere = sqlWhere&" AND DATA_CADASTRO >= "&sqlData(b_DATA_CADASTRO_INICIO)
							End If
					
							If b_DATA_CADASTRO_FIM <> Empty Then
							'Corrigi data
								b_DATA_CADASTRO_FIM = Day(b_DATA_CADASTRO_FIM)&"/"&Month(b_DATA_CADASTRO_FIM)&"/"&Year(b_DATA_CADASTRO_FIM)&"  00:00:00"
							End If	
					
				End if
			
			
			End If
			
	End If	
	sql = sql&sqlWhere&sqlOrderBy
	%>
<section class="dashboard-panel purchase-panel" aria-labelledby="purchase-filter-title">
<h2 id="purchase-filter-title">Filtrar pedidos</h2>
<form action="ppv_compras_pedidos.asp" name="FORMULARIO_FILTRO" method="get">
  <input type="hidden" value="1" name="filtro">
  <div class="purchase-filters">
    <div><label for="purchase-number">N&uacute;mero do pedido</label><input id="purchase-number" class="form-control" inputmode="numeric" name="NUMERO_PEDIDO_COMPRAS" value="<%=Server.HTMLEncode(filtro_NUMERO_PEDIDO_COMPRAS)%>"></div>
    <div><%If validacao_para_listar_qualquer_pedido Then%><label for="purchase-cadastro_cod_usuario">Respons&aacute;vel pelo pedido</label><select class="form-select" id="purchase-cadastro_cod_usuario" name="CADASTRO_COD_USUARIO">
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
    </select><%Else%><strong>Respons&aacute;vel pelo pedido</strong><p><%=nomeUsuario(Session("cod_usuario"),true)%></p><%End If%></div>
    <div><label for="purchase-status_pedido_compra">Status</label><select class="form-select" id="purchase-status_pedido_compra" name="status_pedido_compra">
      <option <%If Cstr(filtro_status_pedido_compra) = "" Then%> selected="selected" <%End If%> value="">Todos</option>
      <%
		Dim temp
		For i=1 to 4
			If i=4 Then
				temp = Cstr("99")
			Else
				temp = Cstr(i)
			End If

		pedidoCompraStatus(i)
		%>
      <option <%If Cstr(filtro_status_pedido_compra) = temp Then%> selected="selected" <%End If%> value="<%=temp%>"><%=pedidoCompraStatus(temp)%></option>
      <%

		Next
    	%>
    </select></div>
    <div><label for="purchase-cod_fornecedor">Fornecedor</label><select class="form-select" id="purchase-cod_fornecedor" name="COD_FORNECEDOR">
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
    </select></div>
    <div><label for="purchase-inicio">Cadastro a partir de</label><div class="purchase-date"><input class='form-control' id='purchase-inicio' inputmode='numeric' placeholder='dd/mm/aaaa' maxlength='10' onkeydown='mascara(this,data);' onkeypress='mascara(this,data);' onkeyup='mascara(this,data);' type='text' name='DATA_CADASTRO_INICIO' value="<%=ArrumaData(b_DATA_CADASTRO_INICIO)%>" onblur='validarData(this)' /><button class="btn btn-outline-secondary" type="button" onclick="pickDate(this,document.FORMULARIO_FILTRO.DATA_CADASTRO_INICIO)" aria-label="Abrir calend&aacute;rio da data inicial"><img alt="" width="19" height="17" src="images/icons/icon_calendario.jpg"></button></div></div>
    <div><label for="purchase-fim">Cadastro at&eacute;</label><div class="purchase-date"><input class='form-control' id='purchase-fim' inputmode='numeric' placeholder='dd/mm/aaaa' maxlength='10' onkeydown='mascara(this,data);' onkeypress='mascara(this,data);' onkeyup='mascara(this,data);' type='text' name='DATA_CADASTRO_FIM' value="<%=ArrumaData(b_DATA_CADASTRO_FIM)%>" onblur='validarData(this)' /><button class="btn btn-outline-secondary" type="button" onclick="pickDate(this,document.FORMULARIO_FILTRO.DATA_CADASTRO_FIM)" aria-label="Abrir calend&aacute;rio da data final"><img alt="" width="19" height="17" src="images/icons/icon_calendario.jpg"></button></div></div>
  </div>
  <div class="purchase-actions"><button class="btn btn-primary" type="submit" name="filtro1" value="1">Aplicar filtros</button><a class="btn btn-outline-secondary" href="ppv_compras_pedidos.asp">Limpar filtros</a></div>
</form>
</section>
<%
Dim comprasAcaoNegada
comprasAcaoNegada = False
Select Case Request.QueryString("codmsg") & ""
    Case "900", "901", "902", "903"
        comprasAcaoNegada = True
End Select
If comprasAcaoNegada Then
%>
<span id="purchase-access-denied" hidden></span>
<noscript><p role="status">Seu perfil n&atilde;o tem permiss&atilde;o para realizar esta a&ccedil;&atilde;o.</p></noscript>
<%End If%>
<%
	
	Set rs=server.CreateObject("ADODB.recordset")
	rs.CursorLocation=3  'Cursor no cliente
	rs.PageSize=50
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
	
    <%
	Dim sucesso, codmsg
	codmsg	= CStr(""&Request.QueryString("codmsg"))
	sucesso = CStr(""&Request.QueryString("sucesso"))
		
	If sucesso = "" And codmsg <> "" And Not comprasAcaoNegada Then%>
    	<div align="center" class="alert alert-success" role="status">
        	<%
			If codmsg = "1" Then
				Response.Write("Pedido de Compra inclu&iacute;do com sucesso!<br />ATEN&Ccedil;&Atilde;O IMPRIMA 3 C&oacute;PIAS, SENDO 1 PARA VOC�, OUTRA PARA O FORNECEDOR E A �LTIMA PARA SER ENTREGUE PARA O DPTO. FINANCEIRO JUNTO COM O BOLETO.")
			Elseif codmsg = "2" Then
				Response.Write("Pedido de Compra alterado com sucesso!")
			Elseif codmsg = "3" Then	
				Response.Write("Pedido de Compra exclu&iacute;do com sucesso!")
			Elseif codmsg = "202" or codmsg = "203" or codmsg = "2099" or codmsg = "2022" Then		
			'Mudan&ccedil;a de status com resposta dinamica
				Response.Write("Pedido de Compra alterado o status para """&pedidoCompraStatus(Mid(codmsg,3,Len(codmsg)))&""" com sucesso!")
			Elseif codmsg = "601" Then
				Response.Write("Pedidos de Compra gerados com sucesso!")
			End If	
			%>
</div>
    <%ElseIf sucesso <> "" And codmsg <> "" And Not comprasAcaoNegada Then%>
        <div align="center" class="alert alert-warning" role="status">
			<%
			If codmsg = "4" Then
				Response.Write("Pedido de Compra n&atilde;o alterado para n&atilde;o haver duplicidade de informa&ccedil;&atilde;o!<BR>O conte&uacute;do do campo """&Server.HTMLEncode(Request.QueryString("chaveUnica") & "")&""" j&aacute; possui uma informa&ccedil;&atilde;o igual em outro Pedido de Compra.<br />")
			Elseif codmsg = "5" Then
				Response.Write("Pedido de Compra n&atilde;o exclu&iacute;do, pois h&aacute; Or&ccedil;amentos/Pedidos/NF que dependem ou relacionam com este vendedor.<BR>Se n&atilde;o desejar ter mais este Pedido de Compra inative ao inv&eacute;s de excluir.<br />")
			
			Elseif codmsg = "300" Then
				Response.Write("Pedido de Compra N&atilde;O alterado status, <br />pois voc&ecirc; n&atilde;o possui acesso ou o pedido j&aacute; foi mudado para o status que voc&ecirc; est&aacute; tentando mudar.")
			End If					
			%>
</div>
<%End If%>
<section class="dashboard-panel purchase-panel" aria-labelledby="purchase-results-title">
<div class="purchase-results-heading">
  <div><h2 id="purchase-results-title">Pedidos encontrados</h2><span><%=rs.RecordCount%> pedidos</span></div>
  <span>P&aacute;gina <strong><%=rs.AbsolutePage%></strong> de <%=rs.PageCount%></span>
</div>
<div class="purchase-table-scroll" role="region" aria-labelledby="purchase-results-title" tabindex="0">
<table class="table purchase-table" role="table" aria-labelledby="purchase-results-title">
<thead role="rowgroup"><tr role="row">
<th role="columnheader" scope="col">Pedido</th>
<th role="columnheader" scope="col">Fornecedor</th>
<th role="columnheader" scope="col">Produto</th>
<th role="columnheader" scope="col">Qtd. pendente</th>
<th role="columnheader" scope="col">Qtd. entregue</th>
<th role="columnheader" scope="col">Respons&aacute;vel</th>
<th role="columnheader" scope="col">Emiss&atilde;o</th>
<th role="columnheader" scope="col">Dias em atraso</th>
<th role="columnheader" scope="col">Status</th>
<th role="columnheader" scope="col">Entrega</th>
<th role="columnheader" scope="col">A&ccedil;&otilde;es</th>
</tr></thead>
<tbody role="rowgroup">
<%
Dim qtd_pendente, qtd_entregue, qtd_atrasados
Dim usuario_pode_imprimir
usuario_pode_imprimir = validaUsuario(1,100,0,true)

Dim x
For x=1 To rs.PageSize
If Not rs.Eof Then
	
	status_pedido_compra = cstr(rs("status_pedido_compra"))
	status_pedido_compra = int(status_pedido_compra)
	qtd_pendente = pedidoCompraTotais(rs("cod_pedido_compra"),5,true)
	qtd_entregue = pedidoCompraTotais(rs("cod_pedido_compra"),4,true)
	qtd_atrasados= pedidoCompraTotais(rs("cod_pedido_compra"),7,true)
%>
<tr class="purchase-row" role="row">
  <td role="cell" data-label="Pedido"><div align="center"><%=rs("numero_pedido_compra")%>
  <% if rs("pedido_manual") = "S" Then%>
  	<small class="purchase-manual">Manual</small>
  <% end if %>
  </div></td>
  <td role="cell" data-label="Fornecedor"><div align="left"><%=rs("nome_fantasia")%></div></td>
  <td role="cell" data-label="Produto"><div align="left">
  <%
  Set rs2=objConexao.Execute("SELECT nome_produto FROM VW_TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra = "&rs("cod_pedido_compra")&" GROUP BY nome_produto ORDER BY nome_produto ASC")
  	If not rs2.Eof Then
		While not rs2.Eof
			Response.Write(rs2("nome_produto")&"&nbsp;&nbsp;")
		rs2.MoveNext
		Wend
	End If
  rs2.Close
  Set rs2=Nothing
  %>
  </div></td>
  <td role="cell" data-label="Qtd. pendente"><div align="center"><%=qtd_pendente%></div></td>
  <td role="cell" data-label="Qtd. entregue"><div align="center"><%=qtd_entregue%></div></td>
  <td role="cell" data-label="Respons&aacute;vel"><div align="left"> <%=nomeUsuario(rs("cadastro_cod_usuario"),true)%> </div></td>
  <td role="cell" data-label="Emiss&atilde;o"><div align="center"><%=ArrumaData(rs("data_cadastro"))%></div></td>
  <td role="cell" data-label="Dias em atraso"><div align="center">
  <%
  If qtd_atrasados > 0 Then
	  Set rs2=objConexao.Execute("SELECT MIN(data_previsao_entrega) AS data_previsao_entrega FROM VW_TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra = "&rs("cod_pedido_compra"))
		If not rs2.Eof Then
			Response.Write(DateDiff("d",rs2("data_previsao_entrega"),date()))
		End If
	  rs2.Close
	  Set rs2=Nothing
  Else
  	Response.Write("-")
  End If
  %>
  
  </div></td>
  <td role="cell" data-label="Status"><span class="purchase-status"><%=pedidoCompraStatus(status_pedido_compra)%></span></td>
  <td role="cell" data-label="Entrega">
  <%If qtd_pendente = 0 Then%>
    <span class="purchase-status is-delivered">Sem pend&ecirc;ncias</span>
  <%ElseIf qtd_atrasados > 0 Then%>
    <span class="purchase-status is-late">Em atraso</span>
  <%Else%>
    <span class="purchase-status is-on-time">Em dia</span>
  <%End If%>
  </td>
  <td role="cell" data-label="A&ccedil;&otilde;es"><div class="purchase-actions purchase-row-actions"><%
		 'Impress&atilde;o somente se o status do pedido de compra for Enviado para o fornecedor, pois neste caso j&aacute; foi aprovado e mudado o status para Enviado para fornecedor para for&ccedil;ar o usu&aacute;rio a mudar o status para Enviado para Fornecedor e ter como controlar a data que foi enviado.
		If status_pedido_compra > 1  Then%>
         <%If comprasPaginasPermitidas.Exists("ppv_compras_pedidos_imprimir.asp") Then%><a class="purchase-action" href="ppv_compras_pedidos_imprimir.asp?c=<%=rs("cod_pedido_compra")%>" target="_blank" rel="noopener">
          Pedido (2&ordf; via)</a><%Else%><button class="purchase-action compras-access-trigger" type="button" aria-haspopup="dialog">Pedido (2&ordf; via)</button><%End If%>
          <%
		 End If

		 If status_pedido_compra > 2 and validaUsuario(2,,83,true) Then%>
         <a class="purchase-action" href="ppv_compras_pedidos.asp?cod_pedido_compra_recibo=<%=rs("cod_pedido_compra")%>&acao_recibo=segunda_via_recibo" target="_blank" rel="noopener">
          Recibo (2&ordf; via)</a>
          <%
		 End If
		  %>
          
       
            <%If comprasPaginasPermitidas.Exists("ppv_compras_pedidos_formulario.asp") Then%><a class="purchase-action" href="ppv_compras_pedidos_formulario.asp?comando=visualizar&cod_pedido_compra=<%=rs("cod_pedido_compra")%>">
          Ver
          </a><%Else%><button class="purchase-action compras-access-trigger" type="button" aria-haspopup="dialog">Ver</button><%End If%>
        
        
        <%
  'Para alterar ou excluir somentes nos seguintes status
  'Se n&atilde;o estiver Enviado para fornecedor (status = 3)
  'Depois &eacute; permitido somente o cancelamento do pedido de compras completo ou cancelamento individual por item.
  %>
        
        <%If int(status_pedido_compra) < 3 and int(status_pedido_compra) <> 99 Then%>
			
                    
                        <%If comprasPaginasPermitidas.Exists("ppv_compras_pedidos_formulario.asp") Then%><a class="purchase-action" href="ppv_compras_pedidos_formulario.asp?comando=alterar&cod_pedido_compra=<%=rs("cod_pedido_compra")%>">
              Editar</a><%Else%><button class="purchase-action compras-access-trigger" type="button" aria-haspopup="dialog">Editar</button><%End If%> 	 	
       
			<%If  int(status_pedido_compra) <> 99 Then%>
              <%If comprasPaginasPermitidas.Exists("ppv_compras_pedidos_formulario.asp") Then%><a class="purchase-action" href="ppv_compras_pedidos_formulario.asp?comando=excluir&cod_pedido_compra=<%=rs("cod_pedido_compra")%>">
              Excluir
              </a><%Else%><button class="purchase-action compras-access-trigger" type="button" aria-haspopup="dialog">Excluir</button><%End If%>
          	 <%End If%>
        <%End If%>
        </div></td>
  </tr>
<%
rs.MoveNext()
End If
Next
%>
</tbody>
</table>
</div>
<%If rs.PageCount > 1 Then
    Dim comprasLinkPagina, comprasParametroPagina
    comprasLinkPagina = "ppv_compras_pedidos.asp?"
    For Each comprasParametroPagina In Array("filtro", "filtro1", "NUMERO_PEDIDO_COMPRAS", "CADASTRO_COD_USUARIO", "status_pedido_compra", "COD_FORNECEDOR", "DATA_CADASTRO_INICIO", "DATA_CADASTRO_FIM")
        If Len(Request.QueryString(comprasParametroPagina) & "") > 0 Then
            comprasLinkPagina = comprasLinkPagina & comprasParametroPagina & "=" & Server.URLEncode(Request.QueryString(comprasParametroPagina) & "") & "&"
        End If
    Next
    comprasLinkPagina = Server.HTMLEncode(comprasLinkPagina & "paginacao=")
%>
<nav class="purchase-pagination" aria-label="P&aacute;ginas dos pedidos">
    <%If paginacaocao > 1 Then%>
      <a class="btn btn-outline-secondary btn-sm" href="<%=comprasLinkPagina%>1">Primeira</a>
      <a class="btn btn-outline-secondary btn-sm" href="<%=comprasLinkPagina & (paginacaocao - 1)%>">Anterior</a>
    <%End If%>
    <span>P&aacute;gina <strong><%=paginacaocao%></strong> de <%=rs.PageCount%></span>
    <%If paginacaocao < rs.PageCount Then%>
      <a class="btn btn-outline-secondary btn-sm" href="<%=comprasLinkPagina & (paginacaocao + 1)%>">Pr&oacute;xima</a>
      <a class="btn btn-outline-secondary btn-sm" href="<%=comprasLinkPagina & rs.PageCount%>">&Uacute;ltima</a>
    <%End If%>
</nav>
<%End If%>
</section>
<%Else%>
<section class="dashboard-panel purchase-panel" role="status"><h2>Nenhum pedido encontrado</h2><p class="mb-0">Confira os filtros para buscar outros pedidos.</p></section>
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
</script>
</body>
</html>
