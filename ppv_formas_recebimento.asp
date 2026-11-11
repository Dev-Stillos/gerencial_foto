<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Formas de Recebimento | Portal Stillos</title>
<link rel="stylesheet" href="css/vendor/bootstrap.min.css">
<link rel="stylesheet" href="css/dashboard.css">
<style>
.payment-method-page .dashboard-nav a.active:not([aria-current]) { background:transparent; color:#d4d5e1; }
.payment-method-page .dashboard-nav a[aria-current="page"] { background:#b20b61; color:#fff; font-weight:600; }
.methods-panel { height:auto; padding:clamp(14px,3vw,26px); }
.methods-heading, .methods-actions, .methods-pagination { display:flex; align-items:center; flex-wrap:wrap; gap:8px; }
.methods-heading { justify-content:space-between; margin-bottom:14px; }
.methods-heading p { margin:5px 0 0; color:#57566a; }
.methods-table { margin:0; }
.methods-table th, .methods-table td { padding:12px 10px; vertical-align:middle; overflow-wrap:anywhere; }
.methods-table thead th { background:#f4f1f7; color:#42364e; }
.methods-table td:first-child { font-weight:700; white-space:nowrap; }
.methods-status { display:inline-block; padding:4px 8px; border-radius:6px; background:#f0edf8; color:#51436e; font-weight:650; }
.methods-status.active { background:#e7f3ec; color:#286143; }
.methods-action { background:#fff; color:#88134f; border:1px solid #e4dfeb; }
.methods-action:hover { background:#f7edf3; color:#741044; }
.methods-pagination { justify-content:center; margin-top:16px; }
@media(max-width:1050px) {
  .methods-table, .methods-table tbody { display:block; }
  .methods-table thead { position:absolute; width:1px; height:1px; overflow:hidden; clip-path:inset(50%); }
  .methods-table tr { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); margin-bottom:10px; border:1px solid #e4dfeb; border-radius:9px; }
  .methods-table td { display:block; min-width:0; border:0; }
  .methods-table td::before { content:attr(data-label); display:block; margin-bottom:4px; color:#57566a; font-size:.85em; font-weight:650; }
  .methods-table td:nth-child(2), .methods-table td:last-child { grid-column:1/-1; }
  .methods-table td:last-child { border-top:1px solid #e4dfeb; }
}
</style>
</head>
<body class="payment-method-page">
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
        <div><h1>Formas de Recebimento</h1></div>
        <a href="ppv_compras_pedidos.asp" class="btn btn-outline-secondary">Voltar aos pedidos</a>
      </div>
      <div class="search-panel"><div><strong>Busca no portal</strong></div><!-- #include file="includes/dashboard_busca.asp" --></div>
      <div class="session-notices"><!-- #include file="includes/dashboard_avisos.asp" --></div>

<!-- #include file="includes/compras_sub_menu.asp" -->
<%
Dim formasFormularioPermitido, formasPermissao, formasConsulta, formasPerfil
formasFormularioPermitido = False
If IsNumeric(Session("COD_PERFIL") & "") Then
    formasPerfil = CLng(Session("COD_PERFIL"))
    formasConsulta = "SELECT p.COD_PAGINA FROM TB_ADM_PAGINAS_HABILITADAS p " & _
        "WHERE p.CAMINHO_VALIDA = 'ppv_formas_recebimento_formulario.asp' " & _
        "AND (COALESCE(p.HABILITADA_PARA_TODOS, 0) <> 2 OR EXISTS (SELECT 1 FROM TB_ADM_PERFIS_PAGINAS_HABILITADAS a " & _
        "WHERE a.COD_PAGINA = p.COD_PAGINA AND a.COD_PERFIL = " & formasPerfil & "))"
    Set formasPermissao = objConexao.Execute(formasConsulta)
    formasFormularioPermitido = Not formasPermissao.Eof
    formasPermissao.Close
    Set formasPermissao = Nothing
End If

Dim sucesso, codmsg
codmsg = Request.QueryString("codmsg") & ""
sucesso = Request.QueryString("sucesso") & ""
Select Case codmsg
    Case "900", "901", "902", "903"
%>
<span id="purchase-access-denied" hidden></span>
<noscript><p role="status">Seu perfil n&atilde;o tem permiss&atilde;o para realizar esta a&ccedil;&atilde;o.</p></noscript>
<%Case "1", "2", "3"
    If sucesso = "" Then
        Dim formasAcao
        Select Case codmsg
            Case "1": formasAcao = "inclu&iacute;da"
            Case "2": formasAcao = "alterada"
            Case "3": formasAcao = "exclu&iacute;da"
        End Select
%><p class="alert alert-success" role="status">Forma de recebimento <%=formasAcao%> com sucesso.</p>
<%End If
Case "4"%><p class="alert alert-warning" role="status">J&aacute; existe outro cadastro com o mesmo valor no campo <strong><%=Server.HTMLEncode(Request.QueryString("chaveUnica") & "")%></strong>.</p>
<%Case "5"%><p class="alert alert-warning" role="status">Este cadastro tem registros vinculados. Se precisar deixar de us&aacute;-lo, altere a situa&ccedil;&atilde;o para Inativa.</p>
<%End Select

sql = "SELECT * FROM TB_FORMAS_RECEBIMENTO order by descricao"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3
rs.PageSize = 50
rs.Open sql, objConexao
Dim paginacaocao, formasPaginaNumero
paginacaocao = 1
Set formasPaginaNumero = New RegExp
formasPaginaNumero.Pattern = "^[0-9]{1,8}$"
If formasPaginaNumero.Test(Request.QueryString("paginacao") & "") Then
    paginacaocao = CLng(Request.QueryString("paginacao"))
    If paginacaocao < 1 Then paginacaocao = 1
End If
If Not rs.Eof Then
    If paginacaocao > rs.PageCount Then paginacaocao = rs.PageCount
    rs.AbsolutePage = paginacaocao
End If
%>
<section class="dashboard-panel methods-panel" aria-labelledby="methods-title">
  <div class="methods-heading">
    <div><h2 id="methods-title">Cadastros</h2><p><%=rs.RecordCount%> <%If rs.RecordCount = 1 Then%>registro<%Else%>registros<%End If%></p></div>
    <%If formasFormularioPermitido Then%>
    <a class="btn btn-primary" href="ppv_formas_recebimento_formulario.asp?comando=incluir">Incluir forma de recebimento</a>
    <%Else%>
    <button class="btn btn-primary compras-access-trigger" type="button" aria-haspopup="dialog">Incluir forma de recebimento</button>
    <%End If%>
  </div>
<%If rs.Eof Then%>
  <p class="mb-0">Nenhuma forma de recebimento cadastrada.</p>
<%Else%>
  <table class="table methods-table" role="table" aria-labelledby="methods-title">
    <thead role="rowgroup"><tr role="row">
      <th scope="col" role="columnheader">C&oacute;digo</th>
      <th scope="col" role="columnheader">Descri&ccedil;&atilde;o</th>
      <th scope="col" role="columnheader">&Uacute;ltima modifica&ccedil;&atilde;o</th>
      <th scope="col" role="columnheader">Situa&ccedil;&atilde;o</th>
      <th scope="col" role="columnheader">A&ccedil;&otilde;es</th>

    </tr></thead>
    <tbody role="rowgroup">
    <%Dim x
    For x = 1 To rs.PageSize
        If rs.Eof Then Exit For
    %>
    <tr role="row">
      <td role="cell" data-label="C&oacute;digo"><%=formatNumber(rs(0),0)%></td>
      <td role="cell" data-label="Descri&ccedil;&atilde;o"><%=Server.HTMLEncode(rs("descricao") & "")%></td>
      <td role="cell" data-label="&Uacute;ltima modifica&ccedil;&atilde;o"><%=ArrumaData(rs("modificado"))%></td>
      <td role="cell" data-label="Situa&ccedil;&atilde;o"><%Select Case UCase(Trim(rs("situacao") & ""))
Case "A"%><span class="methods-status active">Ativa</span>
<%Case "I"%><span class="methods-status">Inativa</span>
<%Case Else%><span class="methods-status"><%=Server.HTMLEncode(rs("situacao") & "")%></span>
<%End Select%></td>
      <td role="cell" data-label="A&ccedil;&otilde;es"><div class="methods-actions">
      <%If formasFormularioPermitido Then%><a class="btn btn-sm methods-action" href="ppv_formas_recebimento_formulario.asp?comando=visualizar&amp;cod_forma_recebimento=<%=Server.URLEncode(rs("cod_forma_recebimento") & "")%>">Ver</a>
      <%Else%><button class="btn btn-sm methods-action compras-access-trigger" type="button" aria-haspopup="dialog">Ver</button><%End If%>
      <%If formasFormularioPermitido Then%><a class="btn btn-sm methods-action" href="ppv_formas_recebimento_formulario.asp?comando=alterar&amp;cod_forma_recebimento=<%=Server.URLEncode(rs("cod_forma_recebimento") & "")%>">Editar</a>
      <%Else%><button class="btn btn-sm methods-action compras-access-trigger" type="button" aria-haspopup="dialog">Editar</button><%End If%>
      <%If formasFormularioPermitido Then%><a class="btn btn-sm methods-action" href="ppv_formas_recebimento_formulario.asp?comando=excluir&amp;cod_forma_recebimento=<%=Server.URLEncode(rs("cod_forma_recebimento") & "")%>">Excluir</a>
      <%Else%><button class="btn btn-sm methods-action compras-access-trigger" type="button" aria-haspopup="dialog">Excluir</button><%End If%></div></td>

    </tr>
    <%rs.MoveNext
    Next%>
    </tbody>
  </table>
  <%If rs.PageCount > 1 Then%>
  <nav class="methods-pagination" aria-label="P&aacute;ginas dos cadastros">
    <%If paginacaocao > 1 Then%><a class="btn btn-outline-secondary btn-sm" href="ppv_formas_recebimento.asp?paginacao=<%=paginacaocao - 1%>">Anterior</a><%End If%>
    <span>P&aacute;gina <%=paginacaocao%> de <%=rs.PageCount%></span>
    <%If paginacaocao < rs.PageCount Then%><a class="btn btn-outline-secondary btn-sm" href="ppv_formas_recebimento.asp?paginacao=<%=paginacaocao + 1%>">Pr&oacute;xima</a><%End If%>
  </nav>
  <%End If%>
<%End If%>
</section>
<%
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
