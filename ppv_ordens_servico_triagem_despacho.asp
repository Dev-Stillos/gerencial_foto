<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="files/includes/functions_local.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<%
Dim caminho_pagina
caminho_pagina = Replace(Request.ServerVariables("URL"), "/gerencial_foto/", "")
Dim efetuar_despacho, despacho_transportador
efetuar_despacho = Cstr(""&Request.Form("campo1"))	
despacho_transportador = Cstr(""&Request.Form("despacho_transportador"))	

If efetuar_despacho = "S" And despacho_transportador <> "" Then

	AbrirConexao()
		'///
		'Limpar base tempor&aacuteria
			objConexao.Execute("DELETE FROM TB_ORDENS_SERVICO_TEMPORARIA WHERE cod_usuario = "&Session("cod_usuario"))
		'///
		'1) Selecionar todas OS para triagem
		Set rs=objConexao.Execute("SELECT cod_ordem_servico, cod_pedido FROM TB_ORDENS_SERVICO WHERE status_ordem_servico = 2 ORDER BY cod_pedido asc")
			If not rs.Eof then

				While not rs.Eof
					
					'3) Verifica se o respons&aacutevel foi selecionado
					cod_pedido = rs("cod_pedido")			
					cod_ordem_servico = rs("cod_ordem_servico")	
						
					despachar = Cstr(""&Request.Form("despachar_"&cod_pedido))
					despachar_item = Cstr(""&Request.Form("item_"&cod_ordem_servico))
					cod_ordem_servico = rs("cod_ordem_servico")
					
					If despachar_item = "S" Then
						
						'1) Altera status 
						sql = "UPDATE TB_ORDENS_SERVICO SET cod_usuario_despachou = "&Session("cod_usuario")&",	data_despachado = NOW(), despacho_transportador = "&sqlTexto(despacho_transportador)&", status_ordem_servico = 3 WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" AND status_ordem_servico = 2 AND data_despachado IS NULL"
						objConexao.Execute(sql)
						
						'2) Grava na base temporaria para emissao do recibo
						sql = "INSERT INTO TB_ORDENS_SERVICO_TEMPORARIA (cod_ordem_servico,cod_usuario) VALUES ('"&cod_ordem_servico&"', '"&Session("cod_usuario")&"')"
						objConexao.Execute(sql)						
					End If
					
								
				rs.MoveNext()
				Wend
				
			End if
			
		rs.close
		Set rs=nothing
		
	FecharConexao()
	
	Response.Redirect("ppv_ordens_servico.asp?codmsg=301")
	
End If
	
	
Dim busca_cod_ordem_servico, busca_nome_formando, busca_cod_contrato, busca_cod_formando, busca_finalizados_parcialmente, busca_cod_usuario_responsavel

Dim sql, sqlWhere, sqlOrderBy, sqlGroupBy

sqlOrderBy = "cod_contrato, nome_formando"
busca_cod_ordem_servico = Cstr(""&Request.QueryString("cod_ordem_servico")) 'N&atildeo est&aacute sendo utilizado
busca_nome_formando = Cstr(""&Request.QueryString("nome_formando"))
busca_cod_contrato = Cstr(""&Request.QueryString("cod_contrato"))
busca_cod_formando = Cstr(""&Request.QueryString("cod_formando"))
busca_finalizados_parcialmente = Cstr(""&Request.QueryString("finalizados_parcialmente"))
busca_cod_usuario_responsavel = Cstr(""&Request.QueryString("cod_usuario_responsavel"))

'Padr&atildeo
Dim validacao_listar_os_todos_usuarios, valida_para_despacho_qualquer_usuario
validacao_listar_os_todos_usuarios = validaUsuario(2,,57,false)
valida_para_despacho_qualquer_usuario = validaUsuario(2,,76,false)

If not valida_para_despacho_qualquer_usuario Then
	sqlWhere = sqlWhere&" and cod_usuario_responsavel = "&sqlTexto(Session("cod_usuario"))	
End If

'Por padr&atildeo lista somente os pedidos que est&atildeo totalmente prontos para despacho
If Request.QueryString("filtro1") <> Empty Then
		If busca_cod_ordem_servico <> "" Then
		
		Else

			If busca_finalizados_parcialmente <> "S" Then
			'Se a opcao de finalizados parcialmente estiver desabilidata ira filtrar somente TOTALMENTE finalizados para despacho
				sqlWhere = sqlWhere&" AND cod_ordem_servico NOT IN (SELECT cod_ordem_servico FROM TB_ORDENS_SERVICO WHERE status_ordem_servico IN(0,1))"
			End If
					
			If busca_cod_usuario_responsavel <> "" Then
				sqlWhere = sqlWhere&" and cod_usuario_responsavel = "&sqlTexto(busca_cod_usuario_responsavel)	
				sql_usuario = " and cod_usuario_responsavel = "&sqlTexto(busca_cod_usuario_responsavel)	
			End If
			
			If busca_nome_formando <> "" Then
				sqlWhere = sqlWhere&" and nome_formando LIKE '%"&busca_nome_formando&"%'"	
			End If

			If busca_cod_contrato <> "" Then
				sqlWhere = sqlWhere&" and cod_contrato = "&sqlTexto(busca_cod_contrato)			
			End If

			If busca_cod_formando <> "" Then
				temp = Right(busca_cod_formando,1)
				
				If Cstr(""&temp) <> "1" Then
					'Response.Redirect("ppv_ordens_servico_triagem_despacho.asp?codmsg=701&sucesso=2")
				Else
					sqlWhere = sqlWhere&" and cod_formando = "&sqlTexto(Mid(busca_cod_formando,1,Len(busca_cod_formando)-1))
				End if
		
			End If					
			
		End If
Else
	sqlWhere = sqlWhere&" AND cod_ordem_servico NOT IN (SELECT cod_ordem_servico FROM TB_ORDENS_SERVICO WHERE status_ordem_servico IN(0,1))"

End If

Dim acesso_inclusao_exclusao_os_manual
acesso_inclusao_exclusao_os_manual = validaUsuario(2,,51,false)
%>

<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Triagem de despacho | Portal Stillos</title>
<link rel="stylesheet" href="css/vendor/bootstrap.min.css">
<link rel="stylesheet" href="css/dashboard.css">
<style>
.dispatch-page .dashboard-nav a.active:not([aria-current]) { background:transparent; color:#d4d5e1; }
.dispatch-page .dashboard-nav a[aria-current="page"] { background:#b20b61; color:#fff; font-weight:600; }
.dispatch-page .ppv-module { margin-bottom:20px; }
.dispatch-page .subMenu { padding:12px 16px; background:#fff; border:1px solid var(--line); border-radius:10px; line-height:2; }
.dispatch-page .subMenu [aria-current="page"] { color:#88134f; font-weight:700; }
.dispatch-page .dispatch-panel { height:auto; margin-bottom:18px; padding:clamp(16px,3vw,26px); }
.dispatch-panel p { margin:6px 0 0; color:#55576b; }
.dispatch-tools { display:flex; flex-wrap:wrap; justify-content:space-between; align-items:center; gap:14px; }
.dispatch-actions { display:flex; align-items:center; flex-wrap:wrap; gap:8px; }
.dispatch-actions .btn { white-space:normal; }
.dispatch-filter { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:14px; margin:14px 0; }
.dispatch-filter label, .dispatch-label, .dispatch-select label { display:block; font-weight:650; }
.dispatch-filter input, .dispatch-select select { margin-top:5px; }
.dispatch-select select { display:block; width:100%; min-height:40px; padding:6px 10px; border:1px solid #dfe0e9; border-radius:7px; background:#fff; color:var(--ink); font:inherit; }
.dispatch-owner { padding:10px 0 0; overflow-wrap:anywhere; }
.dispatch-filter small { display:block; margin-top:4px; color:#55576b; }
.dispatch-check { display:inline-flex; align-items:center; gap:8px; cursor:pointer; }
.dispatch-check input { flex:none; margin:0; }
.dispatch-filter-panel .dispatch-actions { margin-top:14px; }
.dispatch-carrier { max-width:310px; margin:16px 0; }
.dispatch-table { margin:16px 0 0; }
.dispatch-table th, .dispatch-table td { padding:10px; vertical-align:top; overflow-wrap:anywhere; }
.dispatch-table thead th, .dispatch-table .dispatch-group th { background:#f4f1f7; color:#42364e; }
.dispatch-person { display:flex; flex-wrap:wrap; gap:4px 12px; margin-top:5px; font-weight:400; }
.dispatch-status { display:inline-block; padding:3px 7px; border-radius:5px; background:#f0edf8; color:#51436e; font-weight:650; }
.dispatch-status.is-delivered { background:#e7f3ec; color:#286143; }
.dispatch-status.is-open { background:#fce9ee; color:#8e2045; }
.dispatch-submit { justify-content:flex-end; margin-top:16px; }
.dispatch-dialog { box-sizing:border-box; width:min(460px,calc(100% - 24px)); padding:24px; border:1px solid var(--line); border-radius:14px; color:var(--ink); }
.dispatch-dialog::backdrop { background:#24273899; }
.dispatch-dialog .dispatch-actions { justify-content:flex-end; margin-top:14px; }
@media(max-width:1100px) {
  .dispatch-filter { grid-template-columns:repeat(2,minmax(0,1fr)); }
}
@media screen and (max-width:900px) {
  .dispatch-table, .dispatch-table tbody { display:block; }
  .dispatch-table thead { position:absolute; width:1px; height:1px; overflow:hidden; clip-path:inset(50%); }
  .dispatch-group, .dispatch-group th { display:block; }
  .dispatch-order { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); border-bottom:1px solid var(--line); }
  .dispatch-order td { display:block; border:0; }
  .dispatch-order td::before { content:attr(data-label); display:block; color:#66586e; font-size:.85em; font-weight:700; }
}
@media(max-width:480px) {
  .dispatch-filter { grid-template-columns:minmax(0,1fr); }
  .dispatch-tools .dispatch-actions, .dispatch-tools .dispatch-actions .btn { width:100%; }
}
@media print {
  .dispatch-page .dashboard-layout { display:block; }
  .dispatch-page .dashboard-sidebar, .dispatch-page .dashboard-topbar, .dispatch-page .ppv-module,
  .dispatch-page .search-panel, .dispatch-page .session-notices, .dispatch-page .dashboard-heading .btn,
  .dispatch-page .dashboard-footer, .dispatch-page .dispatch-tools[aria-label], .dispatch-filter-panel,
  .dispatch-print, .dispatch-carrier, .dispatch-submit, .dispatch-dialog, .dispatch-check { display:none !important; }
  .dispatch-page .dashboard-main, .dispatch-page .dispatch-panel { padding:0; border:0; }
}
</style>
</head>
<body class="dispatch-page">
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
      <div class="topbar-title">Portal Stillos <span>/ Produ&ccedil;&atilde;o p&oacute;s-venda</span></div>
      <div class="topbar-actions"><a class="user-name" href="adm_log_usuarios.asp"><%=Server.HTMLEncode(Session("NOME") & "")%></a><a href="manual_sistema.asp" target="_blank" rel="noopener">Manual</a><a class="btn btn-outline-secondary btn-sm" href="adm_sair.asp">Sair</a></div>
    </header>
    <main class="dashboard-main ppv-main" id="conteudo">
      <div class="dashboard-heading">
        <div><p class="eyebrow">PRODU&Ccedil;&Atilde;O P&Oacute;S-VENDA</p><h1>Triagem de despacho</h1><p>Selecione as ordens e informe como os produtos ser&atilde;o enviados ao cliente.</p></div>
        <a href="ppv_ordens_servico.asp" class="btn btn-outline-secondary">Voltar &agrave;s ordens</a>
      </div>
      <div class="search-panel"><div><strong>Busca no portal</strong></div><!-- #include file="includes/dashboard_busca.asp" --></div>
      <div class="session-notices"><!-- #include file="includes/dashboard_avisos.asp" --></div>
      <nav class="ppv-module" aria-label="Produ&ccedil;&atilde;o p&oacute;s-venda">
        <!-- #include file="includes/classe_modulo_ppv_menu.asp" -->
      </nav>
      <section class="dashboard-panel dispatch-panel dispatch-tools" aria-label="Recibos e pedidos despachados">
        <div><h2>Recibos e pedidos</h2><p>Consulte o &uacute;ltimo recibo ou altere um despacho j&aacute; registrado.</p></div>
        <div class="dispatch-actions">
          <a class="btn btn-outline-secondary" href="ppv_ordens_servico_triagem_despacho_imprimir.asp" target="_blank" rel="noopener">Reimprimir &uacute;ltimo recibo</a>
          <a class="btn btn-outline-secondary" href="ppv_ordens_servico_triagem_despacho_alterar.asp">Alterar pedido despachado</a>
        </div>
      </section>
      <section class="dashboard-panel dispatch-panel dispatch-filter-panel" aria-labelledby="dispatch-filter-title">
        <h2 id="dispatch-filter-title">Filtrar ordens</h2>
        <form method="get" action="ppv_ordens_servico_triagem_despacho.asp">
          <input type="hidden" name="filtro1" value="1">
          <div class="dispatch-filter">
            <div><label for="dispatch-name">Nome do formando</label><input class="form-control" id="dispatch-name" name="nome_formando" value="<%=Server.HTMLEncode(busca_nome_formando)%>"></div>
            <div><label for="dispatch-student">C&oacute;digo do formando (ficha)</label><input class="form-control" id="dispatch-student" name="cod_formando" aria-describedby="dispatch-student-help" value="<%=Server.HTMLEncode(busca_cod_formando)%>"><small id="dispatch-student-help">Informe o c&oacute;digo com 1 no final.</small></div>
            <div><label for="dispatch-contract">C&oacute;digo do contrato</label><input class="form-control" id="dispatch-contract" name="cod_contrato" value="<%=Server.HTMLEncode(busca_cod_contrato)%>"></div>
            <div class="dispatch-select">
              <%If valida_para_despacho_qualquer_usuario Then%>
                <label>Respons&aacute;vel pela OS
<%=PadraoFormularioSelect("cod_usuario_responsavel","","inputTamanho2","","","Todos",busca_cod_usuario_responsavel,"SELECT cod_usuario, nome FROM TB_ADM_USUARIOS WHERE COD_USUARIO IN(SELECT cod_usuario_responsavel FROM VW_TB_ORDENS_SERVICO WHERE status_ordem_servico = 2 GROUP BY cod_usuario_responsavel) ORDER BY NOME", "incluir", false, false, false )%>
                </label>
              <%Else%>
                <span class="dispatch-label">Respons&aacute;vel pela OS</span>
                <div class="dispatch-owner"><%=Server.HTMLEncode(nomeUsuario(Session("cod_usuario"),false) & "")%></div>
              <%End If%>
            </div>
          </div>
          <label class="dispatch-check"><input class="form-check-input" type="checkbox" value="S" name="finalizados_parcialmente"<%If busca_finalizados_parcialmente = "S" Then%> checked<%End If%>> Incluir pedidos finalizados parcialmente</label>
          <div class="dispatch-actions">
            <button class="btn btn-primary" type="submit">Aplicar filtros</button>
            <a class="btn btn-outline-secondary" href="ppv_ordens_servico_triagem_despacho.asp">Limpar filtros</a>
          </div>
        </form>
      </section>
<%						
sql = "SELECT cod_pedido FROM VW_TB_ORDENS_SERVICO WHERE status_ordem_servico = 2 "&sqlWhere&" GROUP BY cod_pedido ORDER BY cod_contrato, nome_formando LIMIT 100"

AbrirConexao()

Set rs=objConexao.Execute(sql)
If not rs.Eof Then

'Verifica a qtd de OS de acordo com o filtro
Set rs2=objConexao.Execute("SELECT COUNT(*) as qtd FROM VW_TB_ORDENS_SERVICO WHERE status_ordem_servico = 2 "&sqlWhere)
If not rs2.Eof Then
qtdOSs = rs2(0)
End If
rs2.Close
Set rs2=Nothing
%>

<section class="dashboard-panel dispatch-panel" aria-labelledby="dispatch-orders-title">
  <div class="dispatch-tools">
    <div><h2 id="dispatch-orders-title">Ordens para despacho</h2><p>At&eacute; 100 pedidos por consulta, em ordem de contrato e formando.</p></div>
    <button class="btn btn-outline-secondary dispatch-print" type="button" onclick="window.print()">Imprimir lista</button>
  </div>
  <form id="dispatch-form" name="DESPACHO_CLIENTE" action="ppv_ordens_servico_triagem_despacho.asp" method="post">
    <input name="campo1" value="S" type="hidden">
    <div class="dispatch-select dispatch-carrier"><label>Transportador
      <%=PadraoFormularioSelect("despacho_transportador","","inputTamanho3","","","Selecione",letra,"PADRAO:TRANSPORTADORA#TRANSPORTADORA|CORREIOS#CORREIOS|RETIRADA#RETIRADA|STILLOS#STILLOS|STILLOS#STILLOS|PANDEMIA#PANDEMIA|LOGISTICA#LOGISTICA|ESPECIAL#ESPECIAL", "incluir", false, false, false )%>
    </label></div>
    <p id="dispatch-feedback" class="alert alert-warning" role="status" hidden></p>
    <table class="table dispatch-table" role="table" aria-labelledby="dispatch-orders-title">
      <thead role="rowgroup"><tr role="row">
        <th scope="col" role="columnheader">OS</th><th scope="col" role="columnheader">Produto</th>
        <th scope="col" role="columnheader">Finaliza&ccedil;&atilde;o</th><th scope="col" role="columnheader">Prazo geral</th>
        <th scope="col" role="columnheader">Dias at&eacute; o prazo</th><th scope="col" role="columnheader">Despacho</th>
      </tr></thead>
      <tbody role="rowgroup">
		<%
		Dim numero_etapa, descricao, cod_usuario_responsavel, nome_usuario_responsavel, prazo_em_dias, data_inicio, prazo_limite, diferenca, qtd_etapas
		While not rs.Eof
			Set rs2=objConexao.Execute("SELECT * FROM TB_ORDENS_SERVICO WHERE cod_pedido = "&rs("cod_pedido")&" and data_finalizado IS NULL and status_ordem_servico NOT IN (0,99)")
			If not rs2.Eof Then
				permitido = false
			Else
				permitido = true							
			End If
			rs2.Close
			Set rs2=Nothing
			%>
            <tr class="dispatch-group" role="row"><th colspan="6" role="cell">
              <strong>Pedido #<%=rs("cod_pedido")%></strong>
              <%
						sql = " SELECT "&_
						" TB_PEDIDOS_FORMANDOS.COD_FORMANDO, "&_
						" TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, "&_
						" TB_FICHAS_IDENTIFICACAO.COD_CONTRATO, "&_
						" TB_ENTIDADES.DESCRICAO, "&_
						" TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO, "&_
						" TB_FICHAS_IDENTIFICACAO.NOME_CURSO "&_
						" FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON "&_
						" TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO "&_
						" INNER JOIN TB_ENTIDADES ON(TB_FICHAS_IDENTIFICACAO.COD_CONTRATO = TB_ENTIDADES.COD_CONTRATO)"&_
						" WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&rs("cod_pedido")&" ORDER BY NOME_FORMANDO"

						Set rs2=objConexao.Execute(sql)
                        While Not rs2.Eof
              %>
              <div class="dispatch-person"><strong><%=Server.HTMLEncode(rs2("NOME_FORMANDO") & "")%></strong><span><%=Server.HTMLEncode(rs2("DESCRICAO") & "")%></span></div>
              <%
                            rs2.MoveNext()
                        Wend
                        rs2.Close
                        Set rs2 = Nothing
              %>
              <%If permitido Then%><input hidden type="checkbox" value="S" name="despachar_<%=rs("cod_pedido")%>"><%End If%>
            </th></tr>
			<%
			sql = "SELECT VW_TB_ORDENS_SERVICO.*, A.data  FROM VW_TB_ORDENS_SERVICO LEFT JOIN TB_PEDIDOS_ITENS_ENTREGUES A ON A.cod_ordem_servico = VW_TB_ORDENS_SERVICO.cod_ordem_servico WHERE cod_pedido = "&rs("cod_pedido")&" and status_ordem_servico NOT IN (0,99) "&sql_usuario&" order by numero_ordem_servico desc, numero_ordem_servico_indice asc"
			'Response.Write(sql)
			Set rs2=objConexao.Execute(sql)
			While not rs2.Eof

				data_finalizado = rs2("data_finalizado")
			
				Set rs3=objConexao.Execute("Select count(*) as qtd_etapas From TB_ORDENS_SERVICO_ETAPAS where cod_ordem_servico = "&rs2("cod_ordem_servico")&" group by cod_ordem_servico")
				If not rs3.Eof Then
					qtd_etapas = formatNumber(0&rs3("qtd_etapas"),0)
				End If
				rs3.Close
				Set rs3=Nothing

			%>                   
			<tr class="dispatch-order" role="row">
                <td role="cell" data-label="OS"><strong><%=moduloPPVOSMascara(rs2("cod_ordem_servico"),"",true)%></strong></td>
                <td role="cell" data-label="Produto"><%=Server.HTMLEncode(rs2("nome_produto") & "")%></td>
                <td role="cell" data-label="Finaliza&ccedil;&atilde;o"><%If IsDate(data_finalizado) Then%><%=ArrumaData(data_finalizado)%><%Else%><span class="dispatch-status is-open">Em aberto</span><%End If%></td>
                <td role="cell" data-label="Prazo geral"><%=ArrumaData(DateAdd("d",rs2("prazo_cliente_total"),rs2("data_base_prazo_cliente_inicio")))%></td>
                <td role="cell" data-label="Dias at&eacute; o prazo"><%=dateDiff("d",date(),DateAdd("d",rs2("prazo_cliente_total"),rs2("data_base_prazo_cliente_inicio")))%></td>
                <td role="cell" data-label="Despacho">
                    <%If len(rs2("data")) > 3 Then%>
                        <span class="dispatch-status is-delivered">Entregue</span>
                    <%elseIf len(rs2("data_despachado")) > 3 Then%>
                        <span class="dispatch-status">Despachado</span>
                    <%Else%>
                        <label class="dispatch-check"><input type="checkbox" value="S" name="item_<%=rs2("cod_ordem_servico")%>" class="form-check-input"> Selecionar</label>
                    <%End If%>
                </td>
            </tr>

			<%
			rs2.MoveNext()
			Wend

			rs2.Close
			Set rs2=Nothing

		rs.MoveNext()
		Wend
	%>
      </tbody>
    </table>
    <div class="dispatch-actions dispatch-submit">
      <span id="dispatch-count" role="status">0 OS selecionadas</span>
      <button class="btn btn-primary" type="submit" aria-haspopup="dialog">Revisar despacho</button>
    </div>
  </form>
</section>
<dialog id="dispatch-confirm" class="dispatch-dialog" aria-labelledby="dispatch-confirm-title">
  <h2 id="dispatch-confirm-title">Confirmar despacho?</h2>
  <p id="dispatch-confirm-message"></p>
  <form class="dispatch-actions" method="dialog">
    <button class="btn btn-outline-secondary" value="cancelar" autofocus>Cancelar</button>
    <button class="btn btn-primary" value="confirmar">Confirmar despacho</button>
  </form>
</dialog>

<%Else%>
<section class="dashboard-panel dispatch-panel" role="status">
  <h2>Nenhuma OS para despacho</h2>
  <p class="mb-0">Confira os filtros para buscar outros pedidos.</p>
</section>
<%End if

rs.Close
Set rs=Nothing

FecharConexao()
%>    </main>
    <footer class="dashboard-footer">&copy; <%=Year(Date)%> Grupo Stillos <span>Sistema administrativo interno</span></footer>
  </div>
</div>
<script>
(function () {
    var toggle = document.getElementById('ppvMenuToggle');
    var menu = document.getElementById('menuPrincipal');
    var previous = menu.querySelector('[aria-current="page"]');
    var current = menu.querySelector('a[href="ppv_ordens_servico.asp"]');
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
    var current = document.querySelector('.ppv-module a[href="ppv_ordens_servico_triagem_despacho.asp"]');
    if (current) current.setAttribute('aria-current', 'page');

    var form = document.getElementById('dispatch-form');
    if (!form) return;
    var dialog = document.getElementById('dispatch-confirm');
    var feedback = document.getElementById('dispatch-feedback');
    var carrier = form.querySelector('select[name="despacho_transportador" i]');
    var counter = document.getElementById('dispatch-count');
    var items = form.querySelectorAll('input[name^="item_"]');
    function selectionCount() {
        return form.querySelectorAll('input[name^="item_"]:checked').length;
    }
    function selectionText(count) {
        return count + (count === 1 ? ' OS selecionada' : ' OS selecionadas');
    }
    form.addEventListener('change', function () {
        counter.textContent = selectionText(selectionCount());
        feedback.hidden = true;
    });
    form.addEventListener('submit', function (event) {
        event.preventDefault();
        var count = selectionCount();
        if (!carrier || !carrier.value.trim() || !count) {
            feedback.textContent = !carrier || !carrier.value.trim()
                ? 'Selecione o transportador para continuar.'
                : 'Selecione pelo menos uma ordem para despacho.';
            feedback.hidden = false;
            if (!count && carrier && carrier.value.trim()) {
                if (items.length) items[0].focus();
            } else if (carrier) {
                carrier.focus();
            }
            return;
        }
        var carrierName = carrier.options[carrier.selectedIndex].text.trim();
        document.getElementById('dispatch-confirm-message').textContent =
            selectionText(count) + '. Transportador: ' + carrierName + '.';
        dialog.returnValue = '';
        dialog.showModal();
    });
    dialog.addEventListener('close', function () {
        if (dialog.returnValue === 'confirmar') form.submit();
    });
})();
</script>
</body>
</html>
