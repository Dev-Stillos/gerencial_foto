<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%




'EXCLUS�O DE FECHAMENTO COMPLETO
	 If Cstr(inStr(sub_validacoes_paginas,"#42")) <> "0" and Request.QueryString("comando") = "excluir" Then
		
		AbrirConexao()
		
		'1) Deleta o cabe�alho		
		sql = "DELETE FROM TB_FECHAMENTO_PAGAMENTO_CABECALHO WHERE VENDEDOR_RECEBE_EM = "&sqlData(Request.QueryString("VENDEDOR_RECEBE_EM"))
		
		'Response.Write(sql&"<br />")
		objConexao.Execute(sql)
		
		
		'2) Deleta Tabela de Fechamento	
		sql = "DELETE FROM TB_FECHAMENTO_PAGAMENTO WHERE VENDEDOR_RECEBE_EM = "&sqlData(Request.QueryString("VENDEDOR_RECEBE_EM"))
		
		'Response.Write(sql&"<br />")
		objConexao.Execute(sql)
		
		'3) Update em Pedidos como n�o processados
		sql = "UPDATE TB_PEDIDOS SET PAGAMENTO_PROCESSADO = 2 WHERE VENDEDOR_RECEBE_EM = "&sqlData(Request.QueryString("VENDEDOR_RECEBE_EM"))
		
		
		'Response.Write(sql&"<br />")
		objConexao.Execute(sql)
		
		'4) Ciclos de fechamentos
		sql = "UPDATE TB_FECHAMENTO_PAGAMENTO_PERIODOS SET data_processado_pagamento = NULL WHERE data_pagamento = "&sqlData(Request.QueryString("VENDEDOR_RECEBE_EM"))

		'Response.Write(sql&"<br />")
		objConexao.Execute(sql)



		objConexao.Execute("DELETE FROM TB_FECHAMENTO_PAGAMENTO_BOLETOS WHERE vendedor_recebe_em = "&sqlData(Request.QueryString("VENDEDOR_RECEBE_EM")) )
		objConexao.Execute("UPDATE TB_PEDIDOS_BOLETOS SET cod_fechamento_periodo = Null , pagamento_processado = 0 WHERE vendedor_recebe_em = "&sqlData(Request.QueryString("VENDEDOR_RECEBE_EM")) )

		FecharConexao()
		
		Response.Redirect("cap_fechamentos.asp?codmsg=3")
		
	 End If
%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Fechamentos | Portal Stillos</title>
<link rel="stylesheet" href="css/vendor/bootstrap.min.css">
<link rel="stylesheet" href="css/dashboard.css">
<style>
.cap-page .dashboard-nav a.active:not([aria-current]) { background:transparent; color:#d4d5e1; }
.cap-page .dashboard-nav a[aria-current="page"] { background:#b20b61; color:#fff; font-weight:600; }
.cap-page .cap-panel { height:auto; margin-bottom:16px; padding:clamp(14px,3vw,26px); }
.cap-overview h2 { margin-bottom:10px; }
.cap-metrics { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; margin:0 0 16px; }
.cap-metric { min-width:0; padding:16px; border:1px solid #e4dfeb; border-radius:12px; background:#fff; }
.cap-metric-primary { background:#faf0f6; border-color:#ebd3e1; }
.cap-metric dt { color:#57566a; font-size:.87em; font-weight:650; }
.cap-metric dd { margin:6px 0 0; font-size:clamp(19px,2vw,27px); line-height:1.15; font-weight:700; font-variant-numeric:tabular-nums; }
.cap-metric small { display:block; margin-top:6px; color:#57566a; }
.cap-tools { display:grid; grid-template-columns:minmax(0,1fr) minmax(0,1.2fr); gap:16px; }
.cap-tools h2 { margin-bottom:10px; }
.cap-tools p, .cap-muted { color:#57566a; }
.cap-tools p { margin:6px 0 14px; }
.cap-actions, .cap-heading { display:flex; align-items:center; flex-wrap:wrap; gap:8px; }
.cap-actions .btn { white-space:normal; }
.cap-heading { justify-content:space-between; margin-bottom:14px; }
.cap-heading form { margin:0; }
.cap-period-field { flex:1 1 155px; max-width:230px; }
.cap-tools form { align-items:flex-end; }
.cap-period-field label, .cap-date-fields label { display:block; font-weight:650; margin-bottom:4px; }
.cap-links { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:8px; margin-top:12px; }
.cap-links a { padding:8px 10px; border:1px solid #e4dfeb; border-radius:7px; color:#88134f; text-decoration:none; }
.cap-links a:hover { background:#faf0f6; }
.cap-table { width:100%; margin:0; }
.cap-table th, .cap-table td { padding:10px; vertical-align:middle; overflow-wrap:anywhere; }
.cap-table th { background:#f4f1f7; color:#42364e; }
.cap-details { display:inline-flex; flex-direction:column; align-items:flex-start; gap:2px; background:#faf0f6; color:#88134f; border-color:#e8c4d8; }
.cap-details:hover { background:#f3dce9; color:#741044; border-color:#c889ac; }
.cap-pagination { justify-content:center; margin-top:14px; }
.cap-help summary { font-weight:700; cursor:pointer; }
.cap-help p { margin:10px 0 0; }
.cap-dialog { box-sizing:border-box; width:min(480px,calc(100% - 24px)); max-height:calc(100dvh - 24px); padding:22px; border:1px solid #e4dfeb; border-radius:14px; color:var(--ink); }
.cap-reports-dialog { width:min(720px,calc(100% - 24px)); }
.cap-dialog::backdrop { background:#24273899; }
.cap-dialog p { margin:10px 0; }
.cap-report-meta { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:10px; margin:0 0 18px; padding:12px; border-radius:8px; background:#f7f5fa; }
.cap-report-meta dt { font-size:.85em; color:#57566a; }
.cap-report-meta dd { margin:4px 0 0; overflow-wrap:anywhere; }
.cap-modal-section { margin:0 0 8px; font-size:1rem; font-weight:700; }
.cap-report-grid { display:grid; gap:8px; }
.cap-report .btn, .cap-report-grid > .btn { display:flex; justify-content:space-between; align-items:center; gap:8px; width:100%; text-align:left; white-space:normal; }
.cap-report-grid .cap-muted { display:block; padding:8px 12px; border-radius:7px; background:#f7f5fa; }
.cap-delete-area { margin-top:16px; padding-top:12px; border-top:1px solid #e4dfeb; }
.cap-date-fields { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:10px; margin-top:12px; }
.cap-confirm-actions { justify-content:flex-end; margin-top:14px; }
@media(max-width:980px) { .cap-tools { grid-template-columns:minmax(0,1fr); gap:0; } .cap-metrics { grid-template-columns:repeat(2,minmax(0,1fr)); } }
@media(max-width:700px) {
  .cap-table, .cap-table tbody { display:block; }
  .cap-table thead { position:absolute; width:1px; height:1px; overflow:hidden; clip-path:inset(50%); }
  .cap-row { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); margin-bottom:10px; border:1px solid #e4dfeb; border-radius:9px; }
  .cap-row td { display:block; min-width:0; border:0; }
  .cap-row td:first-child { grid-column:1/-1; background:#faf8fc; }
  .cap-row td::before { content:attr(data-label); display:block; color:#57566a; font-size:.85em; font-weight:650; }
}
@media(max-width:440px) { .cap-links, .cap-report-meta, .cap-date-fields { grid-template-columns:minmax(0,1fr); } .cap-metric { padding:10px; } .cap-dialog { padding:14px; } }
</style>
</head>
<body class="cap-page">
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
      <div class="topbar-title">Portal Stillos <span>/ Fechamentos</span></div>
      <div class="topbar-actions"><a class="user-name" href="adm_log_usuarios.asp"><%=Server.HTMLEncode(Session("NOME") & "")%></a><a href="manual_sistema.asp" target="_blank" rel="noopener">Manual</a><a class="btn btn-outline-secondary btn-sm" href="adm_sair.asp">Sair</a></div>
    </header>
    <main class="dashboard-main ppv-main" id="conteudo">
      <div class="dashboard-heading">
        <div><p class="eyebrow">PAGAMENTO DE VENDEDORES</p><h1>Fechamentos</h1><p>Consulte os per&iacute;odos processados e acesse os relat&oacute;rios de pagamento.</p></div>
        <a href="adm_index.asp" class="btn btn-outline-secondary">Voltar ao painel</a>
      </div>
      <div class="search-panel"><div><strong>Busca no portal</strong></div><!-- #include file="includes/dashboard_busca.asp" --></div>
      <div class="session-notices"><!-- #include file="includes/dashboard_avisos.asp" --></div>

<%
    AbrirConexao()
    Dim cod_fechamento_periodo, data_pagamento_nao_processado
    Set rs2=objConexao.Execute("SELECT cod_fechamento_periodo, data_pagamento FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE data_processado_pagamento IS NULL ORDER BY data_pagamento ASC ")
    
    Dim capPeriodOptions, capPendingCount
    capPeriodOptions = ""
    capPendingCount = 0
    While Not rs2.Eof
        If capPendingCount = 0 Then
            cod_fechamento_periodo = rs2(0)
            data_pagamento_nao_processado = rs2(1)
        End If
        capPendingCount = capPendingCount + 1
        capPeriodOptions = capPeriodOptions & "<option value=""" & Server.HTMLEncode(rs2(0) & "") & """>" & Server.HTMLEncode(ArrumaData(rs2(1)) & "") & "</option>"
        rs2.MoveNext
    Wend
    rs2.Close
    Set rs2 = Nothing
    FecharConexao()
%>


	  	  <%
		  AbrirConexao()
		  
		  
		  '////////////////////////////////////////////////////
		  'Valida��o de relat�rios
		  
		  Dim acesso_1, acesso_2, acesso_3, acesso_4
		  acesso_1 = false
		  acesso_2 = false
		  acesso_3 = false
		  acesso_4 = false
		  
		  Set rs=objConexao.Execute("SELECT * FROM TB_ADM_PERFIS_PAGINAS_HABILITADAS WHERE COD_PAGINA in(163,164,168,165) and COD_PERFIL = '"&Session("cod_perfil")&"'")
		  		
				While not rs.Eof
					
					If int(rs("COD_PAGINA")) = 163 Then
					'Impress�o de Extrato de Coordenadores
						acesso_1 = true
					Elseif int(rs("COD_PAGINA")) = 164 Then
					'Impress�o de Extrato de Vendedores
						acesso_2 = true
					Elseif int(rs("COD_PAGINA")) = 168 Then
					'Impress�o do Resumo da Folha de Pagamento
						acesso_3 = true
					Elseif int(rs("COD_PAGINA")) = 165 Then
					'Impress�o de Recibo
						acesso_4 = true
					End If
				
				rs.MoveNext()
				Wend
							  
		  rs.Close
		  Set rs=Nothing
		  '////////////////////////////////////////////////////
		  
		  Dim rs, sql
		  sql = "SELECT VENDEDOR_RECEBE_EM,DATA_PROCESSAMENTO, count(*) as QTD_PEDIDOS, COD_FECHAMENTO_PERIODO FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_FECHAMENTO_PERIODO > 0 group by VENDEDOR_RECEBE_EM order by VENDEDOR_RECEBE_EM desc"  
		  
		  
		  	Set rs=server.CreateObject("ADODB.recordset")
			rs.CursorLocation=3  'Cursor no cliente
			rs.PageSize=12
			rs.Open sql, objConexao		  
		  
		  Dim capLatestDate, capLatestOrders, capHistoryCount
          capHistoryCount = rs.RecordCount
          If Not rs.Eof Then
              capLatestDate = ArrumaData(rs("VENDEDOR_RECEBE_EM")) & ""
              capLatestOrders = rs("QTD_PEDIDOS")
          End If
          %>
<section class="cap-overview" aria-labelledby="cap-overview-title">
  <h2 id="cap-overview-title">Resumo dos fechamentos</h2>
  <dl class="cap-metrics">
    <div class="cap-metric cap-metric-primary">
      <dt>Pr&oacute;ximo pagamento pendente</dt>
      <dd><%If capPendingCount > 0 Then%><%=ArrumaData(data_pagamento_nao_processado)%><%Else%>&mdash;<%End If%></dd>
      <small><%=capPendingCount%> <%If capPendingCount = 1 Then%>per&iacute;odo a processar<%Else%>per&iacute;odos a processar<%End If%></small>
    </div>
    <div class="cap-metric">
      <dt>Fechamento mais recente</dt>
      <dd><%If Len(capLatestDate) > 0 Then%><%=capLatestDate%><%Else%>&mdash;<%End If%></dd>
      <small>Data de pagamento</small>
    </div>
    <div class="cap-metric">
      <dt>Pedidos no &uacute;ltimo fechamento</dt>
      <dd><%If Len(capLatestDate) > 0 Then%><%=capLatestOrders%><%Else%>&mdash;<%End If%></dd>
    </div>
    <div class="cap-metric">
      <dt>Fechamentos no hist&oacute;rico</dt>
      <dd><%=capHistoryCount%></dd>
    </div>
  </dl>
</section>
<div class="cap-tools">
  <section class="dashboard-panel cap-panel" aria-labelledby="cap-preview-title">
    <h2 id="cap-preview-title">Previs&atilde;o da folha de pagamento</h2>
    <p>Consulte a previs&atilde;o de um per&iacute;odo ainda n&atilde;o processado.</p>
    <form class="cap-actions" name="EXTRATO_APROVEITAMENTO" action="cap_fechamentos_IMPRESSAO_folha_geral.asp" target="_blank" method="get">
      <div class="cap-period-field"><label for="cap-period">Data de pagamento</label><select id="cap-period" class="form-select" name="cod_fechamento_periodo" required<%If Len(cod_fechamento_periodo & "") = 0 Then%> disabled<%End If%>>
        <%If capPendingCount = 0 Then%><option value="">Nenhum per&iacute;odo pendente</option><%Else%><%=capPeriodOptions%><%End If%>
      </select></div>
      <button class="btn btn-primary" type="submit"<%If Len(cod_fechamento_periodo & "") = 0 Then%> disabled<%End If%>>Abrir previs&atilde;o</button>
    </form>
  </section>
  <section class="dashboard-panel cap-panel" aria-labelledby="cap-routines-title">
    <h2 id="cap-routines-title">Rotinas do fechamento</h2>
    
    <div class="cap-actions">
      <%If Len(cod_fechamento_periodo & "") > 0 Then%>
      <a class="btn btn-primary" target="_blank" rel="noopener" href="adm_rotina_fechamento_extrato_pagamento.asp?cod_fechamento_periodo=<%=cod_fechamento_periodo%>">Gerar fechamento de <%=ArrumaData(data_pagamento_nao_processado)%></a>
      <%End If%>
      <button class="btn btn-outline-secondary" id="cap-ranking-open" type="button" aria-haspopup="dialog">Ranking de vendedores</button>
    </div>
    <nav class="cap-links" aria-label="Cadastros e pend&ecirc;ncias">
      <a href="cap_fechamentos_periodos.asp">Per&iacute;odos de vendas</a>
      <a href="cap_fechamentos_tabelas_comissao_vendedores.asp">Comiss&atilde;o de vendedores</a>
      <a href="cap_fechamentos_comissionamento.asp">Comiss&atilde;o de coordenadores</a>
      <a href="cap_fechamentos_pendencias.asp">Pend&ecirc;ncias de pagamento</a>
    </nav>
  </section>
</div>
<%If Len(Request.QueryString("msg") & "") > 0 Then%>
<p class="alert alert-info" role="status"><%=Server.HTMLEncode(Request.QueryString("msg") & "")%></p>
<%End If
Dim sucesso, codmsg
codmsg = Request.QueryString("codmsg") & ""
sucesso = Request.QueryString("sucesso") & ""
If sucesso = "" And codmsg = "1" Then%>
<p class="alert alert-success" role="status">Fechamento processado. Os recibos est&atilde;o dispon&iacute;veis na lista.</p>
<%ElseIf sucesso = "" And codmsg = "3" Then%>
<p class="alert alert-success" role="status">Fechamento exclu&iacute;do com sucesso.</p>
<%ElseIf sucesso <> "" And codmsg = "41" Then%>
<p class="alert alert-warning" role="status">O pagamento do m&ecirc;s de refer&ecirc;ncia j&aacute; foi conclu&iacute;do. A pend&ecirc;ncia n&atilde;o pode ser alterada ou exclu&iacute;da.</p>
<%End If%>
<%
          If not rs.Eof then
		  
		  	Dim PAGINA_ATUAL
			PAGINA_ATUAL = formatNumber(0&Request.QueryString("pagina"),0)
			
			If PAGINA_ATUAL > 0 Then
					If PAGINA_ATUAL > rs.PageCount Then PAGINA_ATUAL = rs.PageCount
					rs.AbsolutePage=PAGINA_ATUAL
			Else
				PAGINA_ATUAL = 1
				rs.AbsolutePage=1
			End If	
		  
		  %>

<section class="dashboard-panel cap-panel" aria-labelledby="cap-history-title">
  <div class="cap-heading"><div><h2 id="cap-history-title">Hist&oacute;rico de fechamentos</h2></div><span>P&aacute;gina <%=rs.AbsolutePage%> de <%=rs.PageCount%></span></div>
  <table class="table cap-table" role="table" aria-labelledby="cap-history-title">
    <thead role="rowgroup"><tr role="row"><th scope="col" role="columnheader">Pagamento</th><th scope="col" role="columnheader">Processamento</th><th scope="col" role="columnheader">Gerado por</th><th scope="col" role="columnheader">Pedidos</th></tr></thead>
    <tbody role="rowgroup">
    <%
	 
				 
				  
		 Dim x, primeiro
		 primeiro = true
		 
		 For x=1 To rs.PageSize
		    If Not rs.Eof Then
			
				PAGAMENTO_GERADO_POR = ""
                Set rs2=objConexao.Execute("SELECT TB_FECHAMENTO_PAGAMENTO_CABECALHO.*, TB_ADM_USUARIOS.NOME FROM TB_ADM_USUARIOS INNER JOIN TB_FECHAMENTO_PAGAMENTO_CABECALHO ON TB_ADM_USUARIOS.COD_USUARIO = TB_FECHAMENTO_PAGAMENTO_CABECALHO.PAGAMENTO_GERADO_POR WHERE VENDEDOR_RECEBE_EM = "&sqlData(rs(0)))
				
					If not rs2.Eof Then
						
						DIA_VENCIMENTO = rs2("DIA_VENCIMENTO")
						PAGAMENTO_GERADO_POR = rs2("NOME")
						
					End If
				
				rs2.Close			
				Set rs2=Nothing
				%>
      <tr role="row" class="cap-row">
        <td role="cell" data-label="Pagamento">
          <button type="button" class="btn btn-outline-secondary btn-sm cap-details" aria-haspopup="dialog" data-period="<%=Server.HTMLEncode(ArrumaData(rs("VENDEDOR_RECEBE_EM")) & "")%>"><strong><%=ArrumaData(rs("VENDEDOR_RECEBE_EM"))%></strong><small>Ver relat&oacute;rios &rarr;</small></button>
          <template>
            <dl class="cap-report-meta">
              <div><dt>Processamento</dt><dd><%=ArrumaData(rs("DATA_PROCESSAMENTO"))%></dd></div>
              <div><dt>Gerado por</dt><dd><%=Server.HTMLEncode(PAGAMENTO_GERADO_POR & "")%></dd></div>
              <div><dt>Pedidos</dt><dd><%=rs("QTD_PEDIDOS")%></dd></div>
            </dl>
            <h3 class="cap-modal-section">Relat&oacute;rios do per&iacute;odo</h3>
            <div class="cap-report-grid"><div class="cap-report"><%If acesso_4 Then%>
                            	<a class="btn btn-outline-secondary" target="_blank" rel="noopener" href="adm_rotina_fechamento_extrato_pagamento_part3_IMPRESSAO_recibo_extenso_agrupado.asp?COD_FECHAMENTO_PERIODO=<%=rs("COD_FECHAMENTO_PERIODO")%>"><span>Recibos</span><span aria-hidden="true">&#8599;</span></a>
                            <%Else
							Response.Write("<span class=""cap-muted"">Recibos &middot; acesso restrito</span>")
							
							End If
							%></div>
                                
				            
				              <div class="cap-report"><%If acesso_2 Then%>
				              <a class="btn btn-outline-secondary" target="_blank" rel="noopener" href="adm_rotina_fechamento_extrato_pagamento_part3_IMPRESSAO_detalhamento_vendedores.asp?COD_FECHAMENTO_PERIODO=<%=rs("COD_FECHAMENTO_PERIODO")%>"><span>Extrato de vendedores e coordenadores</span><span aria-hidden="true">&#8599;</span></a>
				              <%Else
							Response.Write("<span class=""cap-muted"">Extrato de vendedores &middot; acesso restrito</span>")
							
							End If
							%></div>
				              
				            
				              <div class="cap-report"><%If acesso_3 Then%>
				              <a class="btn btn-outline-secondary" target="_blank" rel="noopener" href="adm_rotina_fechamento_extrato_pagamento_part3_IMPRESSAO_folha_geral.asp?COD_FECHAMENTO_PERIODO=<%=rs("COD_FECHAMENTO_PERIODO")%>"><span>Folha Geral</span><span aria-hidden="true">&#8599;</span></a><%Else
							Response.Write("<span class=""cap-muted"">Folha geral &middot; acesso restrito</span>")
							
							End If
							%></div>
				           
							
								<% 
								' --- Mostra o link apenas para usuários autorizados ---
								 If Session("COD_USUARIO") = 57 Or Session("COD_USUARIO") = 16 Or Session("COD_USUARIO") = 22 Or Session("COD_USUARIO") = 42 Or Session("COD_USUARIO") = 53 Or Session("COD_USUARIO") = 1 Then
									Dim rsBoleto, temBoleto
									Set rsBoleto = objConexao.Execute("SELECT SUM(TOTAL_BOLETO) as TOTAL FROM TB_FECHAMENTO_PAGAMENTO WHERE VENDEDOR_RECEBE_EM = " & sqlData(rs("VENDEDOR_RECEBE_EM")) & " AND TOTAL_BOLETO > 0")
									temBoleto = False
									If Not rsBoleto.Eof Then
										If Not IsNull(rsBoleto("TOTAL")) Then
											If CDbl(rsBoleto("TOTAL")) > 0 Then temBoleto = True
										End If
									End If
									rsBoleto.Close
									Set rsBoleto = Nothing
									If temBoleto Then
								%>
									<a href="adm_controle_no_extrato_comissao_vendedores_boletos.asp?COD_FECHAMENTO_PERIODO=<%=rs("COD_FECHAMENTO_PERIODO")%>" 
									class="btn btn-outline-secondary" target="_blank" rel="noopener"><span>Comiss&atilde;o por boleto</span><span aria-hidden="true">&#8599;</span></a>
								<% Else %>
									<span class="cap-muted">Sem pagamentos por boleto</span>
								<% 
									End If
							End If	
								%>
							
							
				            </div>
            <%If Cstr(inStr(sub_validacoes_paginas,"#42")) <> "0" And primeiro Then
                primeiro = False
            %>
            <div class="cap-delete-area"><button class="btn btn-outline-danger btn-sm" type="button" data-delete-period="<%=Server.HTMLEncode(rs(0) & "")%>">Excluir fechamento</button></div>
            <%End If%>
          </template>
        </td>
        <td role="cell" data-label="Processamento"><%=ArrumaData(rs("DATA_PROCESSAMENTO"))%></td>
        <td role="cell" data-label="Gerado por"><%=Server.HTMLEncode(PAGAMENTO_GERADO_POR & "")%></td>
        <td role="cell" data-label="Pedidos"><strong><%=rs("QTD_PEDIDOS")%></strong></td>
      </tr>
      <%rs.MoveNext
      End If
      Next%>
    </tbody>
  </table>
  <%If rs.PageCount > 1 Then%>
  <nav class="cap-actions cap-pagination" aria-label="Paginar fechamentos">
    <%If PAGINA_ATUAL > 1 Then%><a class="btn btn-outline-secondary btn-sm" href="?pagina=1">Primeira</a><a class="btn btn-outline-secondary btn-sm" href="?pagina=<%=PAGINA_ATUAL - 1%>">Anterior</a><%End If%>
    <span>P&aacute;gina <%=PAGINA_ATUAL%> de <%=rs.PageCount%></span>
    <%If PAGINA_ATUAL < rs.PageCount Then%><a class="btn btn-outline-secondary btn-sm" href="?pagina=<%=PAGINA_ATUAL + 1%>">Pr&oacute;xima</a><a class="btn btn-outline-secondary btn-sm" href="?pagina=<%=rs.PageCount%>">&Uacute;ltima</a><%End If%>
  </nav>
  <%End If%>
</section>
<%Else%>
<section class="dashboard-panel cap-panel" role="status"><h2>Nenhum fechamento processado</h2></section>
<%End If
rs.Close
Set rs = Nothing
FecharConexao()
%>
<details class="dashboard-panel cap-panel cap-help">
  <summary>Instru&ccedil;&otilde;es de impress&atilde;o</summary>
  <p>Para recibos e folha geral, desative os cabe&ccedil;alhos e rodap&eacute;s do navegador e configure margens de 1,5 cm.</p>
  <p>Para extratos de vendedores, copie o conte&uacute;do para o <a href="documentos/modelo_extrato_pagamento.doc" target="_blank" rel="noopener">modelo de extrato no Word</a> e ajuste as quebras de p&aacute;gina por vendedor. O extrato legado foi preparado para Internet Explorer.</p>
</details>

<dialog class="cap-dialog cap-reports-dialog" id="cap-reports" aria-labelledby="cap-reports-title">
  <header class="cap-heading"><h2 id="cap-reports-title">Relat&oacute;rios do fechamento</h2><form method="dialog"><button class="btn btn-outline-secondary btn-sm" autofocus>Fechar</button></form></header>
  <div id="cap-reports-content"></div>
</dialog>
<dialog class="cap-dialog" id="cap-delete" aria-labelledby="cap-delete-title" aria-describedby="cap-delete-message">
  <h2 id="cap-delete-title">Excluir fechamento?</h2>
  <p id="cap-delete-message"></p>
  <form method="dialog" class="cap-actions cap-confirm-actions">
    <button class="btn btn-outline-secondary" value="cancelar" autofocus>Cancelar</button>
    <button class="btn btn-danger" value="excluir">Excluir fechamento</button>
  </form>
</dialog>
<dialog class="cap-dialog" id="cap-ranking" aria-labelledby="cap-ranking-title">
  <h2 id="cap-ranking-title">Ranking de vendedores</h2>
  <p>Selecione o intervalo das vendas para abrir o relat&oacute;rio.</p>
  <form method="get" action="adm_relatorios_ranking_vendedores_completo.asp" id="cap-ranking-form">
    <div class="cap-date-fields">
      <div><label for="cap-ranking-from">Data inicial</label><input class="form-control" id="cap-ranking-from" type="date" name="data_venda_inicio" required></div>
      <div><label for="cap-ranking-to">Data final</label><input class="form-control" id="cap-ranking-to" type="date" name="data_venda_fim" required></div>
    </div>
    <div class="cap-actions cap-confirm-actions">
      <button class="btn btn-outline-secondary" type="submit" formmethod="dialog" formnovalidate>Cancelar</button>
      <button class="btn btn-primary" type="submit">Abrir ranking</button>
    </div>
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
    var current = menu.querySelector('a[href="cap_fechamentos.asp"]');
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
    var reports = document.getElementById('cap-reports');
    var title = document.getElementById('cap-reports-title');
    var content = document.getElementById('cap-reports-content');
    var deletion = document.getElementById('cap-delete');
    var ranking = document.getElementById('cap-ranking');
    var rankFrom = document.getElementById('cap-ranking-from');
    var rankTo = document.getElementById('cap-ranking-to');
    var deletePeriod;

    document.querySelectorAll('.cap-details').forEach(function (button) {
        button.addEventListener('click', function () {
            title.textContent = 'Fechamento de ' + button.dataset.period;
            content.replaceChildren(button.nextElementSibling.content.cloneNode(true));
            reports.showModal();
        });
    });
    content.addEventListener('click', function (event) {
        var button = event.target.closest('[data-delete-period]');
        if (!button) return;
        deletePeriod = button.dataset.deletePeriod;
        document.getElementById('cap-delete-message').textContent =
            title.textContent + '. Os pedidos do período voltarão a ficar sem processamento de pagamento.';
        deletion.returnValue = '';
        deletion.showModal();
    });
    deletion.addEventListener('close', function () {
        if (deletion.returnValue === 'excluir' && deletePeriod) {
            location.assign('cap_fechamentos.asp?comando=excluir&vendedor_recebe_em=' + encodeURIComponent(deletePeriod));
        }
        deletePeriod = '';
    });
    document.getElementById('cap-ranking-open').addEventListener('click', function () {
        ranking.showModal();
        rankFrom.focus();
    });
    rankFrom.addEventListener('input', function () { rankTo.min = rankFrom.value; });
})();
</script>
</body>
</html>
