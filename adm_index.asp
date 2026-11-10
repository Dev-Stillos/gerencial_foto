<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Valida somente o usu�rio
COD_USUARIO_INTERNO_LOGADO 	= Session("COD_USUARIO")
COD_PERFIL_LOGADO			= Session("COD_PERFIL")

Function DashboardNumero(nome, padrao, minimo, maximo)
    Dim valor
    valor = Trim(Request.QueryString(nome) & "")
    DashboardNumero = padrao
    If Len(valor) = 0 Or Len(valor) > 4 Then Exit Function
    If Not IsNumeric(valor) Then Exit Function
    valor = CDbl(valor)
    If valor < minimo Or valor > maximo Then Exit Function
    If valor <> Fix(valor) Then Exit Function
    DashboardNumero = CInt(valor)
End Function
Dim MES_INICIO, ANO_INICIO, MES_FIM, ANO_FIM
MES_INICIO = DashboardNumero("MES_INICIO", Month(Date), 1, 12)
ANO_INICIO = DashboardNumero("ANO_INICIO", Year(Date), 2008, Year(Date)+1)
MES_FIM = DashboardNumero("MES_FIM", Month(Date), 1, 12)
ANO_FIM = DashboardNumero("ANO_FIM", Year(Date), 2008, Year(Date)+1)

'Verifica se o usu�rio est� logado
if COD_USUARIO_INTERNO_LOGADO=Empty OR COD_PERFIL_LOGADO=Empty then
Session.Abandon()
Response.Redirect("index.asp")		
end if

'//ROTINA QUE VERIFICA SE H� PEDIDOS QUE N�O POSSUEM ORDEM DE SERVI�O GERADA POR ALGUM MOTIVO.
Server.Execute("adm_rotina_corrige_os_geracao.asp")

AbrirConexao()
'//ROTINA QUE VERIFICA SE H� OS FINALIZADAS, POR�M COM ETAPAS COM DATA EM ABERTO
objConexao.Execute("UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_termino = NOW( ) WHERE data_termino IS NULL AND cod_ordem_servico IN (SELECT cod_ordem_servico FROM TB_ORDENS_SERVICO WHERE status_ordem_servico = 3)")

'//ROTINA QUE CORRIGE SE H� FORMANDOS VENDIDOS E QUE CONSTAM COMO N�O ENTREGUES
Set rs=objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE TIPO = 0 AND DATA_RETORNO IS NULL AND COD_FORMANDO IN (SELECT COD_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE VENDIDO = 1)")

	While not rs.Eof		
		Set rs2=objConexao.Execute("SELECT DATA_VENDA FROM TB_PEDIDOS WHERE FINALIZA_CORRECAO = 1 AND COD_PEDIDO IN (SELECT cod_pedido FROM TB_PEDIDOS_FORMANDOS WHERE COD_FORMANDO = "&rs("COD_FORMANDO")&")")

			If not rs2.Eof Then
				sql = "UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET DATA_RETORNO = "&sqlData(rs2("DATA_VENDA"))&", COBRAR_DEVOLUCAO = 2 WHERE COD_MOVIMENTACAO = "&rs("COD_MOVIMENTACAO")
				objConexao.Execute(sql)
			End If
		rs2.Close	
		Set rs2=Nothing
	rs.MoveNext()
	Wend
	
rs.Close
Set rs=Nothing
'//


'//ROTINA QUE CORRIGE SE H� ITENS ANTIGOS QUE CONTAM EM ABERTO, POR�M O MESMO ITEM MAIS RECENTE EST� DEVOLVIDO. COMO � UM ALGORITMO MUITO IMPORTANTO OS REGISTROS ALTERADOS EST�O SENDO MARCARDOS.
AbrirConexao()

Set rs=objConexao.Execute("SELECT COD_MOVIMENTACAO AS a, COD_FORMANDO b, TIPO AS c, DATA_SAIDA AS d, DATA_RETORNO AS e FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES z WHERE COBRAR_DEVOLUCAO = 1 AND DATA_RETORNO IS NULL AND DATA_SAIDA < (SELECT MAX(DATA_SAIDA) FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_FORMANDO = z.COD_FORMANDO AND TIPO = z.TIPO) GROUP BY COD_FORMANDO, TIPO ORDER BY COD_FORMANDO ASC")

	While not rs.Eof
		
		'Atualiza o retorno com a pr�pria data de sa�da
		sql = "UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET DATA_RETORNO = "&sqlData(rs("d"))&", COBRAR_DEVOLUCAO = 2, CORRECAO_ALGORITMO_DATA = "&sqlData(now())&" WHERE COD_MOVIMENTACAO = "&rs("a")
		objConexao.Execute(sql)
		
	rs.MoveNext()
	Wend
	
rs.Close
Set rs=Nothing
'//


'//ROTINA QUE CORRIGE SE H� ITENS COMPRADOS, POR�M N�O H� MAIS OS PARA O ITEM E CANCELA
'//SOMENTE SE O PRODUTO N�O FOR ORIUNDO DE PEDIDO MANUAL'
AbrirConexao()

'Set rs=objConexao.Execute("SELECT a.cod_pedido_compra_produto, b.cod_ordem_servico FROM TB_PEDIDOS_COMPRAS_PRODUTOS a LEFT JOIN  TB_ORDENS_SERVICO_ETAPAS b ON (a.cod_ordem_servico_etapa = b.cod_ordem_servico_etapa) WHERE b.cod_ordem_servico IS NULL AND a.pedido_manual <> 'S' ")

''	While not rs.Eof

		'Atualiza o retorno com a pr�pria data de sa�da
''		sql = "UPDATE TB_PEDIDOS_COMPRAS_PRODUTOS SET status_produto = 'C', obs='Cancelado, pois OS n�o existe mais' WHERE cod_pedido_compra_produto = "&rs("cod_pedido_compra_produto")
''		objConexao.Execute(sql)
''		
		'N�o se preocupa com status principal, pois este dever� ser cancelado manualmente.
		
''	rs.MoveNext()
''	Wend
	
'rs.Close
'Set rs=Nothing

'//ROTINA QUE CORRIGE SE H� ITENS ANTIGOS QUE CONTAM EM ABERTO, POR�M O MESMO ITEM MAIS RECENTE EST� DEVOLVIDO. COMO � UM ALGORITMO MUITO IMPORTANTE OS REGISTROS ALTERADOS EST�O SENDO MARCARDOS.

AbrirConexao()

Set rs=objConexao.Execute("SELECT COD_MOVIMENTACAO AS a, COD_FORMANDO b, TIPO AS c, DATA_SAIDA AS d, DATA_RETORNO AS e FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES z WHERE COBRAR_DEVOLUCAO = 1 AND DATA_RETORNO IS NULL AND DATA_SAIDA < (SELECT MAX(DATA_SAIDA) FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_FORMANDO = z.COD_FORMANDO AND TIPO = z.TIPO) GROUP BY COD_FORMANDO, TIPO ORDER BY COD_FORMANDO ASC")

	While not rs.Eof
		
		'Atualiza o retorno com a pr�pria data de sa�da
		sql = "UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET DATA_RETORNO = "&sqlData(rs("d"))&", COBRAR_DEVOLUCAO = 2, CORRECAO_ALGORITMO_DATA = "&sqlData(now())&" WHERE COD_MOVIMENTACAO = "&rs("a")
		objConexao.Execute(sql)

	rs.MoveNext()
	Wend
	
rs.Close
Set rs=Nothing
'//


'//ROTINA QUE CORRIGE SE H� OS's PENDENTES, POR�M DEVERIAM ESTAR FINALIZADAS, UMA VEZ QUE AS TODAS AS ETAPAS J� EST�O FINALIZADAS.
AbrirConexao()

'Set rs=objConexao.Execute("SELECT cod_ordem_servico, (SELECT COUNT(*) as qtd_etapas FROM TB_ORDENS_SERVICO_ETAPAS b WHERE b.cod_ordem_servico = a.cod_ordem_servico  GROUP BY b.cod_ordem_servico) as qtd_etapas_total, (SELECT COUNT(*) as qtd_etapas FROM TB_ORDENS_SERVICO_ETAPAS c WHERE  c.cod_ordem_servico = a.cod_ordem_servico AND c.data_termino IS NOT NULL GROUP BY c.cod_ordem_servico) as qtd_etapas_pendentes  FROM TB_ORDENS_SERVICO a  WHERE a.data_finalizado IS NULL AND (SELECT COUNT(*) as qtd_etapas FROM TB_ORDENS_SERVICO_ETAPAS b WHERE b.cod_ordem_servico = a.cod_ordem_servico  GROUP BY b.cod_ordem_servico) = (SELECT COUNT(*) as qtd_etapas FROM TB_ORDENS_SERVICO_ETAPAS c WHERE  c.cod_ordem_servico = a.cod_ordem_servico AND c.data_termino IS NOT NULL GROUP BY c.cod_ordem_servico)")

	'While not rs.Eof
		
		'Atualiza o retorno com a pr�pria data de sa�da	
		'sql = "UPDATE TB_ORDENS_SERVICO SET status_ordem_servico = 2, data_finalizado = NOW(), cod_usuario_finalizou = "&Session("cod_usuario")&" WHERE cod_ordem_servico = "&sqlTexto(rs("cod_ordem_servico"))
		'objConexao.Execute(sql)
		
		
	'rs.MoveNext()
	'Wend
	
'rs.Close
'Set rs=Nothing
'//



'//ROTINA QUE VERIFICA SE H� PEDIDOS DE COMPRAS FINALIZADOS = 3, por�m na realidade possuem itens em aberto
Set rs=objConexao.Execute("SELECT cod_pedido_compra FROM VW_TB_PEDIDOS_COMPRAS_PRODUTOS WHERE status_produto = 'P' AND status_pedido_compra = 3 GROUP BY cod_pedido_compra")

	While not rs.Eof
		
		'Atualiza o retorno com a pr�pria data de sa�da	
		sql = "UPDATE TB_PEDIDOS_COMPRAS SET status_pedido_compra = 2 WHERE cod_pedido_compra = "&sqlTexto(rs("cod_pedido_compra"))
		objConexao.Execute(sql)		
		
	rs.MoveNext()
	Wend
	
rs.Close
Set rs=Nothing


'//ROTINA PARA CORRIGIR STATUS DE PEDIDOS DE COMPRAS
objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS_PRODUTOS SET status_produto = 'P' WHERE quantidade_entregue IS NULL AND (status_produto IS NULL OR status_produto = '0')")	
objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS_PRODUTOS SET status_produto = 'E' WHERE quantidade_entregue IS NOT NULL AND (status_produto IS NULL OR status_produto = '0') AND status_produto <> 'P'")

FecharConexao()


'Server.Execute("adm_ROTINA_importa_entidades.asp")


%>
<%
Dim dashboardFinanceiro, dashboardErroPeriodo
Dim TOTAL_CARTAO, TOTAL_CHEQUE, TOTAL_BOLETO, TOTAL_DINHEIRO, TOTAL
Dim PERIODO_INICIAL, PERIODO_FINAL, dashboardValores, dashboardNomes, dashboardCores, dashboardIndice, dashboardPercentual
dashboardFinanceiro = (Session("COD_USUARIO") = 1 Or Session("COD_USUARIO") = 16)
dashboardErroPeriodo = ""
PERIODO_INICIAL = DateSerial(ANO_INICIO, MES_INICIO, 1)
PERIODO_FINAL = DateSerial(ANO_FIM, MES_FIM + 1, 1)
If PERIODO_INICIAL >= PERIODO_FINAL Then dashboardErroPeriodo = "Escolha um per&iacute;odo final igual ou posterior ao inicial."
TOTAL_CARTAO = 0 : TOTAL_CHEQUE = 0 : TOTAL_BOLETO = 0 : TOTAL_DINHEIRO = 0 : TOTAL = 0
If dashboardFinanceiro And dashboardErroPeriodo = "" Then
    AbrirConexao()
    sql = "SELECT SUM(IFNULL(PAGAMENTO_EM_DINHEIRO,0)) AS dinheiro, SUM(IFNULL(XTOTAL_CHEQUE,0)) AS cheque, SUM(IFNULL(XTOTAL_BOLETO,0)) AS boleto, SUM(IFNULL(XTOTAL_CARTAO,0)) AS cartao, SUM(IFNULL(XTOTAL,0)) AS total FROM VW_PEDIDOS_TOTAL_RECEBIDO WHERE cod_pedido IN (SELECT cod_pedido FROM TB_PEDIDOS WHERE FINALIZA_CORRECAO = 1 AND DATA_VENDA >= " & sqlData(PERIODO_INICIAL) & " AND DATA_VENDA < " & sqlData(PERIODO_FINAL) & ")"
    Set rs = objConexao.Execute(sql)
    If Not rs.EOF Then
        If Not IsNull(rs("dinheiro")) Then TOTAL_DINHEIRO = CDbl(rs("dinheiro"))
        If Not IsNull(rs("cheque")) Then TOTAL_CHEQUE = CDbl(rs("cheque"))
        If Not IsNull(rs("boleto")) Then TOTAL_BOLETO = CDbl(rs("boleto"))
        If Not IsNull(rs("cartao")) Then TOTAL_CARTAO = CDbl(rs("cartao"))
        If Not IsNull(rs("total")) Then TOTAL = CDbl(rs("total"))
    End If
    rs.Close
    Set rs = Nothing
    FecharConexao()
End If
dashboardValores = Array(TOTAL_CARTAO, TOTAL_DINHEIRO, TOTAL_BOLETO, TOTAL_CHEQUE)
dashboardNomes = Array("Cart&atilde;o de cr&eacute;dito", "Dinheiro", "Boleto", "Cheque")
dashboardCores = Array("pink", "purple", "blue", "amber")
%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Vis&atilde;o geral | Stillos Fotos</title>
<link rel="stylesheet" href="css/vendor/bootstrap.min.css">
<link rel="stylesheet" href="css/dashboard.css">
</head>
<body>
<a class="skip-link" href="#conteudo">Ir para o conte&uacute;do</a>
<div class="dashboard-layout">
<aside class="dashboard-sidebar" id="menuPrincipal">
<a class="dashboard-brand" href="adm_index.asp"><img src="images/logoparafundopreto.png" alt="Stillos" width="180"><span>GERENCIAL FOTOS</span></a>
<div class="nav-caption">&Aacute;REA DE TRABALHO</div>
<!-- #include file="includes/dashboard_menu.asp" -->
<div class="sidebar-footer">Stillos &middot; Produ&ccedil;&otilde;es e eventos</div>
</aside>
<div class="dashboard-workspace">
<header class="dashboard-topbar">
<button class="btn menu-toggle" type="button" aria-controls="menuPrincipal" aria-expanded="false">Menu</button>
<div class="topbar-title">Portal Stillos <span>/ Painel administrativo</span></div>
<div class="topbar-actions"><a href="manual_sistema.asp" target="_blank" rel="noopener">Manual</a><span class="user-name"><%=Server.HTMLEncode(Session("NOME") & "")%></span><a class="btn btn-outline-secondary btn-sm" href="adm_sair.asp">Sair</a></div>
</header>
<main id="conteudo" class="dashboard-main">
<div class="dashboard-heading"><div><p class="eyebrow">SEU DIA NO GERENCIAL</p><h1>Vis&atilde;o geral</h1><p>Bem-vindo(a), <strong><%=Server.HTMLEncode(Session("NOME") & "")%></strong>. Acompanhe os resultados e acesse suas ferramentas.</p></div><span class="date-chip"><%=Day(Date)%> de <%=MonthName(Month(Date))%> de <%=Year(Date)%></span></div>
<section class="search-panel" aria-label="Busca no sistema"><div><strong>Encontre o que precisa</strong><p>Consulte pedidos, formandos e produ&ccedil;&otilde;es.</p></div><!-- #include file="includes/dashboard_busca.asp" --></section>
<%If Request.QueryString("codmsg") = "1" Then%><div class="alert alert-warning" role="alert">Voc&ecirc; n&atilde;o possui acesso &agrave; altera&ccedil;&atilde;o de pedidos ap&oacute;s baixa da Diretoria. Solicite a Diretoria para alterar o pedido.</div><%End If%>
<%If dashboardFinanceiro Then%>
<section class="filter-panel" aria-labelledby="tituloResultados"><div><h2 id="tituloResultados">Resumo de vendas</h2><p>Vendas conclu&iacute;das no per&iacute;odo selecionado</p></div>
<form method="get" action="adm_index.asp" class="period-filter">
<%Dim dashboardCampos, dashboardRotulos, dashboardSelecionados
dashboardCampos = Array("MES_INICIO", "ANO_INICIO", "MES_FIM", "ANO_FIM")
dashboardRotulos = Array("M&ecirc;s inicial", "Ano inicial", "M&ecirc;s final", "Ano final")
dashboardSelecionados = Array(MES_INICIO, ANO_INICIO, MES_FIM, ANO_FIM)
For dashboardIndice = 0 To 3%>
<div><label for="<%=dashboardCampos(dashboardIndice)%>"><%=dashboardRotulos(dashboardIndice)%></label><select class="form-select" id="<%=dashboardCampos(dashboardIndice)%>" name="<%=dashboardCampos(dashboardIndice)%>">
<%Dim dashboardMin, dashboardMax
dashboardMin = 1 : dashboardMax = 12
If dashboardIndice = 1 Or dashboardIndice = 3 Then
    dashboardMin = 2008 : dashboardMax = Year(Date)+1
End If
For i = dashboardMin To dashboardMax
    selected = ""
    If i = dashboardSelecionados(dashboardIndice) Then selected = " selected"
%><option value="<%=i%>"<%=selected%>><%If dashboardIndice = 0 Or dashboardIndice = 2 Then%><%=MonthName(i)%><%Else%><%=i%><%End If%></option><%Next%>
</select></div><%Next%><button class="btn btn-primary" type="submit">Aplicar filtro</button></form></section>
<%If dashboardErroPeriodo <> "" Then%><div class="alert alert-warning" role="alert"><%=dashboardErroPeriodo%></div><%Else%>
<div class="period-caption"><%=MonthName(MES_INICIO)%> de <%=ANO_INICIO%> &mdash; <%=MonthName(MES_FIM)%> de <%=ANO_FIM%></div>
<div class="row g-3 mb-4">
<div class="col-12 col-xl-4"><article class="metric-card metric-total"><span class="metric-label">Total geral</span><strong>R$ <%=FormatNumber(TOTAL,2)%></strong><span>Pagamentos vinculados &agrave;s vendas conclu&iacute;das</span><div class="metric-decoration" aria-hidden="true">&#8599;</div></article></div>
<%For dashboardIndice = 0 To 1%><div class="col-12 col-sm-6 col-xl-4"><article class="metric-card"><span class="metric-icon <%=dashboardCores(dashboardIndice)%>" aria-hidden="true"><%If dashboardIndice = 0 Then%>&#9635;<%Else%>$<%End If%></span><span class="metric-label"><%=dashboardNomes(dashboardIndice)%></span><strong>R$ <%=FormatNumber(dashboardValores(dashboardIndice),2)%></strong><span>Valor registrado no per&iacute;odo</span></article></div><%Next%>
</div>
<div class="row g-4"><div class="col-12 col-xl-7"><section class="dashboard-panel"><div class="panel-heading"><h2>Distribui&ccedil;&atilde;o dos pagamentos</h2><span class="subtle-badge">4 modalidades</span></div><p class="panel-description">Participa&ccedil;&atilde;o de cada forma no total geral.</p>
<%If TOTAL = 0 Then%><p class="empty-state">Nenhum valor registrado neste per&iacute;odo.</p><%End If%>
<%For dashboardIndice = 0 To 3
dashboardPercentual = 0
If TOTAL > 0 Then dashboardPercentual = dashboardValores(dashboardIndice) / TOTAL * 100
%><div class="payment-row"><div class="payment-label"><span><i class="dot <%=dashboardCores(dashboardIndice)%>"></i><%=dashboardNomes(dashboardIndice)%></span><strong>R$ <%=FormatNumber(dashboardValores(dashboardIndice),2)%></strong></div><div class="payment-bar" aria-hidden="true"><span class="<%=dashboardCores(dashboardIndice)%>" style="width:<%=Replace(CStr(Round(dashboardPercentual,2)),",",".")%>%"></span></div><span class="payment-percentage"><%=FormatNumber(dashboardPercentual,1)%>% do total</span></div><%Next%>
</section></div><div class="col-12 col-xl-5"><section class="dashboard-panel"><div class="panel-heading"><h2>Valores por modalidade</h2></div><p class="panel-description">Detalhamento financeiro do per&iacute;odo.</p><div class="table-responsive"><table class="table payment-table"><thead><tr><th scope="col">Forma de pagamento</th><th scope="col" class="text-end">Valor</th></tr></thead><tbody><%For dashboardIndice = 0 To 3%><tr><td><%=dashboardNomes(dashboardIndice)%></td><td class="text-end">R$ <%=FormatNumber(dashboardValores(dashboardIndice),2)%></td></tr><%Next%></tbody><tfoot><tr><th scope="row">Total geral</th><td class="text-end">R$ <%=FormatNumber(TOTAL,2)%></td></tr></tfoot></table></div></section></div></div>
<%End If%><%End If%>
<section class="dashboard-panel shortcuts"><div class="panel-heading"><h2>Acesso r&aacute;pido</h2><span class="panel-description">Ferramentas do seu perfil</span></div><div class="row g-3">
<%If MENU_5 = 1 Then%><div class="col-12 col-md-6 col-xl-3"><a class="shortcut-card" href="lca_pedidos_pendentes.asp"><span class="shortcut-symbol">&#9636;</span><strong>Pedidos pendentes</strong><span>Acompanhar lan&ccedil;amentos &#8594;</span></a></div><%End If%>
<%If MENU_1 = 1 Then%><div class="col-12 col-md-6 col-xl-3"><a class="shortcut-card" href="dip_producoes.asp"><span class="shortcut-symbol">&#9673;</span><strong>Produ&ccedil;&otilde;es</strong><span>Consultar produ&ccedil;&otilde;es &#8594;</span></a></div><%End If%>
<%If MENU_15 = 1 Then%><div class="col-12 col-md-6 col-xl-3"><a class="shortcut-card" href="cap_fechamentos.asp"><span class="shortcut-symbol">&#10003;</span><strong>Fechamentos</strong><span>Acessar per&iacute;odos e recibos &#8594;</span></a></div><%End If%>
<%If MENU_4 = 1 Then%><div class="col-12 col-md-6 col-xl-3"><a class="shortcut-card" href="adm_relatorios.asp"><span class="shortcut-symbol">&#8599;</span><strong>Relat&oacute;rios</strong><span>Consultar resultados &#8594;</span></a></div><%End If%>
</div><p class="shortcut-hint">Todas as op&ccedil;&otilde;es dispon&iacute;veis para seu perfil est&atilde;o no menu.</p></section>
<div class="session-notices"><!-- #include file="includes/dashboard_avisos.asp" --></div>
<section class="session-panel"><div><strong><%=Server.HTMLEncode(Session("DEPARTAMENTO") & "")%></strong><p><%If Session("ACESSO_EM") <> Empty Then%>&Uacute;ltimo acesso em <%=Server.HTMLEncode(Session("ACESSO_EM") & "")%> &middot; <%=Server.HTMLEncode(Session("QTDACESSOS") & "")%> acessos<%Else%>Este &eacute; seu primeiro acesso.<%End If%></p></div><%If Session("COD_PERFIL") = 12 Then%><a href="ple_projetos.asp" class="btn btn-outline-secondary">Visualizar projetos de TI</a><%End If%></section>
</main><footer class="dashboard-footer">&copy; <%=Year(Date)%> Grupo Stillos <span>Sistema administrativo interno</span></footer>
</div></div>
<script>
document.querySelector('.menu-toggle').addEventListener('click', function () {
    var opened = this.getAttribute('aria-expanded') !== 'true';
    this.setAttribute('aria-expanded', String(opened));
    document.getElementById('menuPrincipal').classList.toggle('is-open', opened);
});
document.addEventListener('keydown', function (event) {
    if (event.key === 'Escape') {
        document.getElementById('menuPrincipal').classList.remove('is-open');
        document.querySelector('.menu-toggle').setAttribute('aria-expanded', 'false');
    }
});
</script>
</body></html>
