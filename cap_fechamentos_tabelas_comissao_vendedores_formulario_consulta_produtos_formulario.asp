<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<%

'Valida��o especial para verificar se possui acesso a p�gina m�e de formul�rio
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS where COD_PAGINA = '228' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se n�o encontrar ent�o redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
	rs.Close
	Set rs=Nothing
	
FecharConexao()


'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 11/1/2012 12:13:51

'Declara vari�veis para utiliza��o do sistema
Dim id, campo1, comando, cod_tabela_comissao_vendedores, actionformulario

actionformulario = "cap_fechamentos_tabelas_comissao_vendedores_formulario_consulta_produtos_formulario.asp"

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
cod_tabela_comissao_vendedores = Request.QueryString("cod_tabela_comissao_vendedores")
Dim editorPattern, editorTableName
Set editorPattern = New RegExp
editorPattern.Pattern = "^[1-9][0-9]{0,9}$"
If Not editorPattern.Test(cod_tabela_comissao_vendedores) Then
    Response.Redirect "cap_fechamentos_tabelas_comissao_vendedores.asp"
End If
Set editorPattern = Nothing
AbrirConexao()
Set rs = objConexao.Execute("SELECT descricao FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES WHERE cod_tabela_comissao_vendedores = " & cod_tabela_comissao_vendedores)
If rs.Eof Then
    rs.Close
    Set rs = Nothing
    FecharConexao()
    Response.Redirect "cap_fechamentos_tabelas_comissao_vendedores.asp"
End If
editorTableName = rs("descricao") & ""
rs.Close
Set rs = Nothing
FecharConexao()

'Declara vari�veis
Dim cod_produto, tipo_pagamento_normal_avista, valor_pagamento_normal_avista, tipo_pagamento_normal_aprazo, valor_pagamento_normal_aprazo, tipo_pagamento_repasse_avista, valor_pagamento_repasse_avista, tipo_pagamento_repasse_aprazo, valor_pagamento_repasse_aprazo, cod_usuario, modificado



If campo1 <> Empty Then
	
	'Padr�o
	cod_usuario = Session("cod_usuario")
	modificado = now()


	AbrirConexao()
				Set rs=objConexao.Execute("SELECT cod_produto, nome_produto FROM TB_PRODUTOS WHERE ativo = 1 and PAGAMENTO_COMISSAO = 'S' order by order_menu")
					
					If not rs.Eof Then
					
						While not rs.Eof
						
						cod_produto = rs("cod_produto") 
					
						'///////////////////////////////////////////////////////////							
								
						tipo_pagamento_normal_avista = Request.Form("tipo_pagamento_normal_avista_"&cod_produto)
						valor_pagamento_normal_avista = "0"&Request.Form("valor_pagamento_normal_avista_"&cod_produto)
						tipo_pagamento_normal_aprazo = Request.Form("tipo_pagamento_normal_aprazo_"&cod_produto)
						valor_pagamento_normal_aprazo = "0"&Request.Form("valor_pagamento_normal_aprazo_"&cod_produto)
						tipo_pagamento_repasse_avista = Request.Form("tipo_pagamento_repasse_avista_"&cod_produto)
						valor_pagamento_repasse_avista = "0"&Request.Form("valor_pagamento_repasse_avista_"&cod_produto)
						tipo_pagamento_repasse_aprazo = Request.Form("tipo_pagamento_repasse_aprazo_"&cod_produto)
						valor_pagamento_repasse_aprazo = "0"&Request.Form("valor_pagamento_repasse_aprazo_"&cod_produto)
						
						'///////////////////////////////////////////////////////////
						'Verifica se possui o produto cadastrado nesta tabela de venda
						'Para os comandos de incluir ou alterar obs.: N�O H� EXCLUIR!!! POIS PARA OS PRODUTOS QUE
						'FOREM INATIVADOS N�O SER�O LISTADOS E TAMB�M OS PRODUTOS EXCLUIDOS SER�O EXCLUIDOS NESTA
						'TABELA ABAIXO TAMB�M
						Set rs2=objConexao.Execute("SELECT * FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES_PRODUTOS WHERE cod_tabela_comissao_vendedores = "&cod_tabela_comissao_vendedores&" and cod_produto ="&cod_produto)
					
							If not rs2.Eof Then
								'J� existe por tanto altera!								
								objConexao.Execute("UPDATE TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES_PRODUTOS SET tipo_pagamento_normal_avista = "&sqlTexto(tipo_pagamento_normal_avista)&", valor_pagamento_normal_avista = '"&sqlNumero(valor_pagamento_normal_avista)&"', tipo_pagamento_normal_aprazo = "&sqlTexto(tipo_pagamento_normal_aprazo)&", valor_pagamento_normal_aprazo = '"&sqlNumero(valor_pagamento_normal_aprazo)&"', tipo_pagamento_repasse_avista = "&sqlTexto(tipo_pagamento_repasse_avista)&", valor_pagamento_repasse_avista = '"&sqlNumero(valor_pagamento_repasse_avista)&"', tipo_pagamento_repasse_aprazo = "&sqlTexto(tipo_pagamento_repasse_aprazo)&", valor_pagamento_repasse_aprazo = '"&sqlNumero(valor_pagamento_repasse_aprazo)&"', cod_usuario = "&sqlTexto(cod_usuario)&", modificado = "&sqlData(modificado)&" WHERE cod_produto = "&cod_produto&" and cod_tabela_comissao_vendedores = "&cod_tabela_comissao_vendedores)
								
								
							Else
								'N�o existe incluir								
								objConexao.Execute("INSERT INTO TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES_PRODUTOS (cod_produto, cod_tabela_comissao_vendedores, tipo_pagamento_normal_avista, valor_pagamento_normal_avista, tipo_pagamento_normal_aprazo, valor_pagamento_normal_aprazo, tipo_pagamento_repasse_avista, valor_pagamento_repasse_avista, tipo_pagamento_repasse_aprazo, valor_pagamento_repasse_aprazo, cod_usuario, modificado) VALUES ("&sqlTexto(cod_produto)&", "&sqlTexto(cod_tabela_comissao_vendedores)&", "&sqlTexto(tipo_pagamento_normal_avista)&", '"&sqlNumero(valor_pagamento_normal_avista)&"', "&sqlTexto(tipo_pagamento_normal_aprazo)&", '"&sqlNumero(valor_pagamento_normal_aprazo)&"', "&sqlTexto(tipo_pagamento_repasse_avista)&", '"&sqlNumero(valor_pagamento_repasse_avista)&"', "&sqlTexto(tipo_pagamento_repasse_aprazo)&", "&sqlNumero(valor_pagamento_repasse_aprazo)&", "&sqlTexto(cod_usuario)&", "&sqlData(modificado)&")")
													
							End If
						
						rs2.Close
						Set rs2=Nothing
						

						rs.MoveNext()
						Wend
						
						
						'Limpa todos os produtos inativos desta tabela ou que n�o � para pagar comiss�o
						objConexao.Execute("DELETE FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES_PRODUTOS WHERE cod_tabela_comissao_vendedores = "&cod_tabela_comissao_vendedores&" AND cod_produto IN(SELECT cod_produto FROM TB_PRODUTOS WHERE ativo = 2 or PAGAMENTO_COMISSAO = 'N')")					
					
					End If
				
				rs.Close
				Set rs=Nothing
	FecharConexao()

	Response.Redirect("cap_fechamentos_tabelas_comissao_vendedores.asp?codmsg=101")
	
End If

'----------------------------------------------------------------------------------------

%>

<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Tabelas de comissionamento | Portal Stillos</title>
<link rel="stylesheet" href="css/vendor/bootstrap.min.css">
<link rel="stylesheet" href="css/dashboard.css">
<style>
.cap-metrics { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; margin:0 0 16px; }
.cap-metric { min-width:0; padding:16px; border:1px solid #e4dfeb; border-radius:12px; background:#fff; }
.cap-metric dt { color:#57566a; font-size:.87em; font-weight:650; }
.cap-metric dd { margin:6px 0 0; font-size:clamp(19px,2vw,27px); line-height:1.15; font-weight:700; font-variant-numeric:tabular-nums; }
.cap-metric small { display:block; margin-top:6px; color:#57566a; }
.cap-overview-charts { display:grid; grid-template-columns:minmax(0,2fr) minmax(0,1fr); gap:12px; margin:0 0 18px; }
.cap-chart { min-width:0; padding:16px; border:1px solid #e4dfeb; border-radius:12px; background:#fff; }
.cap-chart h3 { margin:0 0 4px; color:#242738; font:700 1.05em/1.3 system-ui,sans-serif; }
.cap-chart p { margin:0 0 12px; color:#57566a; font-size:.9em; }
.cap-chart-bars { list-style:none; margin:0; padding:0; }
.cap-chart-bars li + li { margin-top:8px; }
.cap-chart-bars li > div:first-child { display:flex; justify-content:space-between; gap:8px; font-variant-numeric:tabular-nums; }
.cap-chart-track { height:7px; overflow:hidden; margin-top:3px; border-radius:7px; background:#eee9f4; }
.cap-chart-track span { display:block; height:100%; background:#a31962; }
.cap-chart-donut { display:grid; place-items:center; width:124px; aspect-ratio:1; margin:14px auto 10px; border-radius:50%; background:conic-gradient(#a31962 var(--cap-share),#eee9f4 0); }
.cap-chart-donut strong { display:grid; place-items:center; width:82px; aspect-ratio:1; border-radius:50%; background:#fff; }
.cap-chart .cap-chart-share { margin:0; text-align:center; }
.period-overview { margin-bottom:20px; }
.period-overview > h2 { margin-bottom:10px; }
.period-actions,.period-paging,.period-filter { display:flex; flex-wrap:wrap; align-items:center; gap:8px; }
.period-panel { height:auto; }
.period-panel header { display:flex; flex-wrap:wrap; justify-content:space-between; gap:12px; align-items:end; margin-bottom:12px; }
.period-scroll { overflow-x:auto; border:1px solid #e4dfeb; border-radius:9px; }
.period-table { width:100%; min-width:760px; margin:0; border-collapse:collapse; font-variant-numeric:tabular-nums; }
.period-table th,.period-table td { padding:9px; border-bottom:1px solid #eeeaf2; text-align:left; vertical-align:middle; }
.period-table th { background:#f1edf6; }
.period-table tbody tr:nth-child(even) { background:#faf9fc; }
.period-date { white-space:nowrap; }
.period-meta { display:block; margin-top:3px; font-size:.9em; color:#53566b; }
.period-badge { display:inline-block; padding:3px 7px; border-radius:6px; font-weight:650; background:#faf0dd; color:#754600; }
.period-done { background:#e5f2ea; color:#275e42; }
.period-paging { justify-content:center; margin-top:14px; }
.cap-chart-donut { margin-left:0; }
@media(max-width:850px) { .cap-metrics { grid-template-columns:repeat(2,minmax(0,1fr)); } }
@media(max-width:460px) { .cap-metrics { grid-template-columns:minmax(0,1fr); } }

.commission-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(min(100%,365px),1fr)); gap:14px; align-items:start; }
.commission-card { height:auto; min-width:0; }
.commission-card h3 { margin:0 0 4px; font-size:1.1em; overflow-wrap:anywhere; }
.commission-meta { color:#53566b; font-size:.9em; margin:0 0 12px; }
.commission-card dl { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:8px; margin:0 0 12px; }
.commission-card dl > div { border:1px solid #e4dfeb; border-radius:8px; padding:8px; }
.commission-card dt { font-size:.85em; font-weight:500; }
.commission-card dd { font-size:1.15em; font-weight:700; margin:4px 0 0; }
.commission-card .period-actions { margin-top:12px; }
.commission-chart { margin-bottom:16px; }
@media(max-width:460px) { .commission-card dl { grid-template-columns:1fr; } }

.commission-preview { width:min(1240px,calc(100% - 24px)); max-height:calc(100dvh - 24px); padding:0; border:1px solid #e4dfeb; border-radius:12px; color:#242738; }
.commission-preview[open] { display:flex; flex-direction:column; }
.commission-preview::backdrop { background:#24273899; }
.commission-preview-header { display:flex; flex-wrap:wrap; align-items:center; justify-content:space-between; gap:10px; padding:12px 16px; border-bottom:1px solid #e4dfeb; }
.commission-preview-frame { width:100%; height:80dvh; min-height:0; border:0; }
.commission-preview-status { margin:0; padding:12px 16px; }
.commission-preview [hidden] { display:none !important; }

.editor-card { height:auto; }
.editor-toolbar { display:flex; flex-wrap:wrap; align-items:end; justify-content:space-between; gap:10px; margin-bottom:12px; }
.editor-toolbar label { display:block; font-weight:650; margin-bottom:4px; }
.editor-toolbar input { min-width:min(280px,100%); }
.editor-note { color:#53566b; margin-bottom:12px; }
.editor-scroll { overflow-x:auto; border:1px solid #e4dfeb; border-radius:9px; }
.editor-table { width:100%; border-collapse:collapse; font-variant-numeric:tabular-nums; }
.editor-table th,.editor-table td { border-bottom:1px solid #eeeaf2; padding:8px; vertical-align:middle; }
.editor-table thead th { background:#f1edf6; text-align:center; }
.editor-table thead th:first-child[rowspan] { text-align:left; }
.editor-table tbody th { min-width:140px; font-weight:650; overflow-wrap:anywhere; }
.editor-table tbody tr:nth-child(even) { background:#faf9fc; }
.editor-value { display:flex; gap:6px; align-items:center; }
.editor-value select { width:64px; flex:none; }
.editor-value input { min-width:64px; }
.editor-label { display:none; }
.editor-toolbar [hidden],.editor-table [hidden] { display:none !important; }
@media(max-width:1050px) {
  .editor-table,.editor-table tbody { display:block; }
  .editor-table thead { position:absolute; width:1px; height:1px; clip-path:inset(50%); overflow:hidden; }
  .editor-table tbody tr { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); padding:8px; }
  .editor-table tbody th { grid-column:1/-1; }
  .editor-label { display:block; margin-bottom:4px; font-size:.85em; color:#53566b; }
}
@media(max-width:390px) { .editor-table tbody tr { grid-template-columns:1fr; } }
</style>
<!-- #include file="files/includes/scripts.asp" -->
</head>
<body>
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
      <div class="topbar-title">Portal Stillos <span>/ Tabelas de comissionamento</span></div>
      <div class="topbar-actions"><a class="user-name" href="adm_log_usuarios.asp"><%=Server.HTMLEncode(Session("NOME") & "")%></a><a href="manual_sistema.asp" target="_blank" rel="noopener">Manual</a><a class="btn btn-outline-secondary btn-sm" href="adm_sair.asp">Sair</a></div>
    </header>
    <main class="dashboard-main ppv-main" id="conteudo">



      <header class="dashboard-heading">
        <div><p class="eyebrow">FECHAMENTOS</p><h1>Editar comiss&otilde;es por produto</h1><p><%=Server.HTMLEncode(editorTableName)%> &middot; Tabela <%=cod_tabela_comissao_vendedores%></p></div>
        <div class="period-actions">
          <a class="btn btn-outline-secondary" href="cap_fechamentos_tabelas_comissao_vendedores.asp">Voltar &agrave;s tabelas</a>
          <a class="btn btn-outline-secondary" data-commission-print href="cap_fechamentos_tabelas_comissao_vendedores_formulario_consulta_produtos_formulario_imprimir.asp?cod_tabela_comissao_vendedores=<%=cod_tabela_comissao_vendedores%>">Imprimir tabela</a>
        </div>
      </header>
      <div class="search-panel"><strong>Busca no portal</strong>
        <!-- #include file="includes/dashboard_busca.asp" -->
      </div>
      <div class="session-notices">
        <!-- #include file="includes/dashboard_avisos.asp" -->
      </div>
      <section class="dashboard-panel editor-card" aria-labelledby="editor-title">
        <h2 id="editor-title" class="mb-3">Produtos da tabela</h2>
        <p class="editor-note">Use % para percentual ou R$ para valor fixo. Ao preencher Normal &agrave; vista, o sistema copia o valor para as condi&ccedil;&otilde;es vazias do mesmo produto.</p>
        <form action="<%=actionformulario%>?comando=<%=Server.URLEncode(comando & "")%>&amp;cod_tabela_comissao_vendedores=<%=cod_tabela_comissao_vendedores%>" name="FORMULARIO_PADRAO" method="post">
          <input type="hidden" value="1" name="campo1">
          <%
          AbrirConexao()
          Set rs=objConexao.Execute("SELECT cod_produto, nome_produto FROM stillos3.TB_PRODUTOS WHERE ativo = 1 and PAGAMENTO_COMISSAO = 'S' order by order_menu")
          If Not rs.Eof Then%>
          <div class="editor-toolbar">
            <div><label for="editor-search">Buscar produto</label><input class="form-control" type="search" id="editor-search" placeholder="Nome ou c&oacute;digo"></div>
            <button class="btn btn-primary" type="submit">Salvar comiss&otilde;es</button>
          </div>
          <p id="editor-count" class="editor-note" role="status"></p>
          <p id="editor-empty" class="empty-state" hidden>Nenhum produto encontrado.</p>
          <div class="editor-scroll" role="region" aria-label="Comiss&otilde;es por produto" tabindex="0">
          <table class="editor-table">
            <thead>
              <tr><th rowspan="2" scope="col">Produto</th><th colspan="2" scope="colgroup">Vendas normais</th><th colspan="2" scope="colgroup">Vendas de repasse</th></tr>
              <tr><th scope="col">&Agrave; vista</th><th scope="col">A prazo</th><th scope="col">&Agrave; vista</th><th scope="col">A prazo</th></tr>
            </thead>
            <tbody>
            <%While Not rs.Eof
cod_produto = rs("cod_produto") 
					
					'///////////////////////////////////////////////////////////
					'Verifica se possui o produto cadastrado nesta tabela de venda
					Set rs2=objConexao.Execute("SELECT * FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES_PRODUTOS WHERE cod_tabela_comissao_vendedores = "&cod_tabela_comissao_vendedores&" and cod_produto ="&cod_produto)
				
						If not rs2.Eof Then
							tipo_pagamento_normal_avista=trim(rs2("tipo_pagamento_normal_avista"))
							valor_pagamento_normal_avista= ArrumaNumero(rs2("valor_pagamento_normal_avista"),"",2)
							tipo_pagamento_normal_aprazo=trim(rs2("tipo_pagamento_normal_aprazo"))
							valor_pagamento_normal_aprazo= ArrumaNumero(rs2("valor_pagamento_normal_aprazo"),"",2)
							tipo_pagamento_repasse_avista=trim(rs2("tipo_pagamento_repasse_avista"))
							valor_pagamento_repasse_avista= ArrumaNumero(rs2("valor_pagamento_repasse_avista"),"",2)
							tipo_pagamento_repasse_aprazo=trim(rs2("tipo_pagamento_repasse_aprazo"))
							valor_pagamento_repasse_aprazo= ArrumaNumero(rs2("valor_pagamento_repasse_aprazo"),"",2)

							
						Else
							tipo_pagamento_normal_avista=""
							valor_pagamento_normal_avista=""
							tipo_pagamento_normal_aprazo=""
							valor_pagamento_normal_aprazo=""
							tipo_pagamento_repasse_avista=""
							valor_pagamento_repasse_avista=""
							tipo_pagamento_repasse_aprazo=""
							valor_pagamento_repasse_aprazo=""					
						End If
					
					rs2.Close
					Set rs2=Nothing
            %>
              <tr data-editor-product="<%=cod_produto%>">
                <th scope="row"><%=Server.HTMLEncode(rs("nome_produto") & "")%></th>
                <td><span class="editor-label">Normal &agrave; vista</span><div class="editor-value">
                  <select class="form-select" name="TIPO_PAGAMENTO_NORMAL_AVISTA_<%=cod_produto%>" aria-label="Tipo: Normal &agrave; vista, produto <%=cod_produto%>">
                    <option value="F"<%If TIPO_PAGAMENTO_NORMAL_AVISTA & "" <> "V" Then%> selected<%End If%>>R$</option>
                    <option value="V"<%If TIPO_PAGAMENTO_NORMAL_AVISTA & "" = "V" Then%> selected<%End If%>>%</option>
                  </select>
                  <input class="form-control padrao_valor" inputmode="decimal" name="VALOR_PAGAMENTO_NORMAL_AVISTA_<%=cod_produto%>" value="<%=Server.HTMLEncode(VALOR_PAGAMENTO_NORMAL_AVISTA & "")%>" maxlength="200" aria-label="Valor: Normal &agrave; vista, produto <%=cod_produto%>" onblur="replicarValor('<%=cod_produto%>')">
                </div></td>
                <td><span class="editor-label">Normal a prazo</span><div class="editor-value">
                  <select class="form-select" name="TIPO_PAGAMENTO_NORMAL_APRAZO_<%=cod_produto%>" aria-label="Tipo: Normal a prazo, produto <%=cod_produto%>">
                    <option value="F"<%If TIPO_PAGAMENTO_NORMAL_APRAZO & "" <> "V" Then%> selected<%End If%>>R$</option>
                    <option value="V"<%If TIPO_PAGAMENTO_NORMAL_APRAZO & "" = "V" Then%> selected<%End If%>>%</option>
                  </select>
                  <input class="form-control padrao_valor" inputmode="decimal" name="VALOR_PAGAMENTO_NORMAL_APRAZO_<%=cod_produto%>" value="<%=Server.HTMLEncode(VALOR_PAGAMENTO_NORMAL_APRAZO & "")%>" maxlength="200" aria-label="Valor: Normal a prazo, produto <%=cod_produto%>">
                </div></td>
                <td><span class="editor-label">Repasse &agrave; vista</span><div class="editor-value">
                  <select class="form-select" name="TIPO_PAGAMENTO_REPASSE_AVISTA_<%=cod_produto%>" aria-label="Tipo: Repasse &agrave; vista, produto <%=cod_produto%>">
                    <option value="F"<%If TIPO_PAGAMENTO_REPASSE_AVISTA & "" <> "V" Then%> selected<%End If%>>R$</option>
                    <option value="V"<%If TIPO_PAGAMENTO_REPASSE_AVISTA & "" = "V" Then%> selected<%End If%>>%</option>
                  </select>
                  <input class="form-control padrao_valor" inputmode="decimal" name="VALOR_PAGAMENTO_REPASSE_AVISTA_<%=cod_produto%>" value="<%=Server.HTMLEncode(VALOR_PAGAMENTO_REPASSE_AVISTA & "")%>" maxlength="200" aria-label="Valor: Repasse &agrave; vista, produto <%=cod_produto%>">
                </div></td>
                <td><span class="editor-label">Repasse a prazo</span><div class="editor-value">
                  <select class="form-select" name="TIPO_PAGAMENTO_REPASSE_APRAZO_<%=cod_produto%>" aria-label="Tipo: Repasse a prazo, produto <%=cod_produto%>">
                    <option value="F"<%If TIPO_PAGAMENTO_REPASSE_APRAZO & "" <> "V" Then%> selected<%End If%>>R$</option>
                    <option value="V"<%If TIPO_PAGAMENTO_REPASSE_APRAZO & "" = "V" Then%> selected<%End If%>>%</option>
                  </select>
                  <input class="form-control padrao_valor" inputmode="decimal" name="VALOR_PAGAMENTO_REPASSE_APRAZO_<%=cod_produto%>" value="<%=Server.HTMLEncode(VALOR_PAGAMENTO_REPASSE_APRAZO & "")%>" maxlength="200" aria-label="Valor: Repasse a prazo, produto <%=cod_produto%>">
                </div></td>

              </tr>
            <%rs.MoveNext
            Wend%>
            </tbody>
          </table>
          </div>
          <%Else%>
          <p class="empty-state">Nenhum produto ativo com comiss&atilde;o encontrado.</p>
          <%End If
          rs.Close
          Set rs = Nothing
          FecharConexao()%>
        </form>
      </section>
    </main>
    <footer class="dashboard-footer">&copy; <%=Year(Date)%> Grupo Stillos <span>Sistema administrativo interno</span></footer>
  </div>
</div>
<script type="text/javascript">
function replicarValor(cod_produto){
	var d = document.FORMULARIO_PADRAO;
	var tipo_pagamento_normal_avista, valor_pagamento_normal_avista, tipo_pagamento_normal_aprazo, valor_pagamento_normal_aprazo, tipo_pagamento_repasse_avista, valor_pagamento_repasse_avista, tipo_pagamento_repasse_aprazo, valor_pagamento_repasse_aprazo;
	
	tipo_pagamento_normal_avista = d['TIPO_PAGAMENTO_NORMAL_AVISTA_'+cod_produto].value;
	valor_pagamento_normal_avista = d['VALOR_PAGAMENTO_NORMAL_AVISTA_'+cod_produto].value;
	tipo_pagamento_normal_aprazo = d['TIPO_PAGAMENTO_NORMAL_APRAZO_'+cod_produto].value;
	valor_pagamento_normal_aprazo = d['VALOR_PAGAMENTO_NORMAL_APRAZO_'+cod_produto].value;
	tipo_pagamento_repasse_avista = d['TIPO_PAGAMENTO_REPASSE_AVISTA_'+cod_produto].value;
	valor_pagamento_repasse_avista = d['VALOR_PAGAMENTO_REPASSE_AVISTA_'+cod_produto].value;
	tipo_pagamento_repasse_aprazo = d['TIPO_PAGAMENTO_REPASSE_APRAZO_'+cod_produto].value;
	valor_pagamento_repasse_aprazo = d['VALOR_PAGAMENTO_REPASSE_APRAZO_'+cod_produto].value;
	
	if(valor_pagamento_normal_avista !=='' && valor_pagamento_normal_avista !=='0,00'){
		
		if(valor_pagamento_normal_aprazo == ''){
		//Se n�o tiver valor altera para o valor � vista venda normal
			d['TIPO_PAGAMENTO_NORMAL_APRAZO_'+cod_produto].value = tipo_pagamento_normal_avista;
			d['VALOR_PAGAMENTO_NORMAL_APRAZO_'+cod_produto].value = valor_pagamento_normal_avista;
		}
		
		if(valor_pagamento_repasse_avista == ''){
		//Se n�o tiver valor altera para o valor � vista venda normal
			d['TIPO_PAGAMENTO_REPASSE_AVISTA_'+cod_produto].value = tipo_pagamento_normal_avista;
			d['VALOR_PAGAMENTO_REPASSE_AVISTA_'+cod_produto].value = valor_pagamento_normal_avista;	
		}
	
	
		if(valor_pagamento_repasse_aprazo == ''){
		//Se n�o tiver valor altera para o valor � vista venda normal
			d['TIPO_PAGAMENTO_REPASSE_APRAZO_'+cod_produto].value = tipo_pagamento_normal_avista;
			d['VALOR_PAGAMENTO_REPASSE_APRAZO_'+cod_produto].value = valor_pagamento_normal_avista;	
		}		

	}


	
}
</script>
<!-- #include file="files/includes/scripts_rodape.asp" -->
<dialog class="commission-preview" id="commission-preview" aria-labelledby="commission-preview-title">
  <header class="commission-preview-header">
    <h2 id="commission-preview-title">Pr&eacute;via da tabela de comiss&otilde;es</h2>
    <div class="period-actions">
      <button class="btn btn-primary btn-sm" type="button" id="commission-print-button" disabled>Imprimir / Salvar PDF</button>
      <form method="dialog"><button class="btn btn-outline-secondary btn-sm" autofocus>Fechar</button></form>
    </div>
  </header>
  <p class="commission-preview-status" id="commission-preview-status" role="status" hidden></p>
  <iframe class="commission-preview-frame" id="commission-preview-frame" title="Pr&eacute;via da tabela de comiss&otilde;es" hidden></iframe>
</dialog>
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
    var dialog = document.getElementById('commission-preview');
    var frame = document.getElementById('commission-preview-frame');
    var status = document.getElementById('commission-preview-status');
    var printButton = document.getElementById('commission-print-button');
    var trigger;

    document.addEventListener('click', function (event) {
        var link = event.target.closest('[data-commission-print]');
        if (!link || typeof dialog.showModal !== 'function' || event.ctrlKey || event.metaKey || event.shiftKey || event.altKey) return;
        event.preventDefault();
        trigger = link;
        var url = new URL(link.href);
        url.searchParams.set('modal', '1');
        status.textContent = 'Carregando tabela...';
        status.hidden = false;
        frame.hidden = false;
        printButton.disabled = true;
        frame.src = url.href;
        dialog.showModal();
    });
    frame.addEventListener('load', function () {
        if (!dialog.open || !frame.hasAttribute('src')) return;
        var preview = frame.contentDocument;
        if (preview && preview.URL === 'about:blank') return;
        var ready = preview && preview.querySelector('[data-commission-report]');
        printButton.disabled = !ready;
        status.hidden = Boolean(ready);
        if (!ready) status.textContent = 'A pr\u00e9via da tabela n\u00e3o est\u00e1 dispon\u00edvel.';
    });
    printButton.addEventListener('click', function () {
        if (!printButton.disabled) frame.contentWindow.print();
    });
    dialog.addEventListener('close', function () {
        frame.hidden = true;
        frame.removeAttribute('src');
        printButton.disabled = true;
        if (trigger) trigger.focus();
    });
})();

(function () {
    var search = document.getElementById('editor-search');
    if (!search) return;
    var rows = Array.from(document.querySelectorAll('[data-editor-product]'));
    var count = document.getElementById('editor-count');
    var empty = document.getElementById('editor-empty');
    function normalize(value) {
        return value.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
    }
    function filterProducts() {
        var term = normalize(search.value);
        var visible = 0;
        rows.forEach(function (row) {
            row.hidden = !normalize(row.cells[0].textContent + ' ' + row.dataset.editorProduct).includes(term);
            if (!row.hidden) visible += 1;
        });
        count.textContent = visible + ' de ' + rows.length + ' produtos';
        empty.hidden = visible !== 0;
    }
    search.addEventListener('input', filterProducts);
    filterProducts();
})();
</script>
</body>
</html>
