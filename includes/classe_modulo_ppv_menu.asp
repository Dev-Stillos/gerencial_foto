<%
Dim ppvPaginasPermitidas, ppvRegistroPermissao, ppvConsultaPermissao, ppvPerfilMenu
Dim ppvPaginaAtualMenu
ppvPaginaAtualMenu = LCase(Request.ServerVariables("URL") & "")
ppvPaginaAtualMenu = Mid(ppvPaginaAtualMenu, InStrRev(ppvPaginaAtualMenu, "/") + 1)
Select Case ppvPaginaAtualMenu
    Case "ppv_ordens_servico_formulario.asp"
        ppvPaginaAtualMenu = "ppv_ordens_servico.asp"
    Case "ppv_ordens_servico_triagem_despacho_alterar.asp"
        ppvPaginaAtualMenu = "ppv_ordens_servico_triagem_despacho.asp"
End Select

Set ppvPaginasPermitidas = Server.CreateObject("Scripting.Dictionary")
ppvPaginasPermitidas.CompareMode = 1

If Len(Session("COD_USUARIO") & "") > 0 And IsNumeric(Session("COD_PERFIL") & "") Then
    ppvPerfilMenu = CLng(Session("COD_PERFIL"))
    AbrirConexao()
    ppvConsultaPermissao = "SELECT p.CAMINHO_VALIDA FROM TB_ADM_PAGINAS_HABILITADAS p " & _
        "WHERE p.CAMINHO_VALIDA IN ('ppv_ordens_servico.asp', 'ppv_produtos_etapas.asp', 'ppv_ordens_servico_triagem_inicial.asp', 'ppv_ordens_servico_triagem_despacho.asp', 'ppv_ordens_servico_triagem_entrega.asp', 'ppv_ordens_servico_os_manual_origem_vendas.asp', 'ppv_ordens_servico_triagem_entrega_correios.asp') " & _
        "AND (COALESCE(p.HABILITADA_PARA_TODOS, 0) <> 2 OR EXISTS (" & _
        "SELECT 1 FROM TB_ADM_PERFIS_PAGINAS_HABILITADAS a WHERE a.COD_PAGINA = p.COD_PAGINA " & _
        "AND a.COD_PERFIL = " & ppvPerfilMenu & "))"
    Set ppvRegistroPermissao = objConexao.Execute(ppvConsultaPermissao)
    While Not ppvRegistroPermissao.Eof
        ppvPaginasPermitidas(ppvRegistroPermissao("CAMINHO_VALIDA") & "") = True
        ppvRegistroPermissao.MoveNext()
    Wend
    ppvRegistroPermissao.Close
    Set ppvRegistroPermissao = Nothing
%>
<div class="ppv-shared-menu" role="group" aria-label="Submenu de Produ&ccedil;&atilde;o P&oacute;s-Venda">
    <%If ppvPaginasPermitidas.Exists("ppv_ordens_servico.asp") Then%>
    <a class="ppv-shared-item" href="ppv_ordens_servico.asp"<%If ppvPaginaAtualMenu = "ppv_ordens_servico.asp" Then%> aria-current="page"<%End If%>>Ordens de Servi&ccedil;o</a>
    <%Else%>
    <button type="button" class="ppv-shared-item ppv-access-trigger" aria-haspopup="dialog">Ordens de Servi&ccedil;o</button>
    <%End If%>
    <%If ppvPaginasPermitidas.Exists("ppv_produtos_etapas.asp") And validaUsuario(2,,54,true) Then%>
    <a class="ppv-shared-item" href="ppv_produtos_etapas.asp"<%If ppvPaginaAtualMenu = "ppv_produtos_etapas.asp" Then%> aria-current="page"<%End If%>>Produtos Etapas</a>
    <%Else%>
    <button type="button" class="ppv-shared-item ppv-access-trigger" aria-haspopup="dialog">Produtos Etapas</button>
    <%End If%>
    <%If ppvPaginasPermitidas.Exists("ppv_ordens_servico_triagem_inicial.asp") And validaUsuario(2,,52,true) Then%>
    <a class="ppv-shared-item" href="ppv_ordens_servico_triagem_inicial.asp"<%If ppvPaginaAtualMenu = "ppv_ordens_servico_triagem_inicial.asp" Then%> aria-current="page"<%End If%>>Triagem Inicial</a>
    <%Else%>
    <button type="button" class="ppv-shared-item ppv-access-trigger" aria-haspopup="dialog">Triagem Inicial</button>
    <%End If%>
    <%If ppvPaginasPermitidas.Exists("ppv_ordens_servico_triagem_despacho.asp") And validaUsuario(2,,53,true) Then%>
    <a class="ppv-shared-item" href="ppv_ordens_servico_triagem_despacho.asp"<%If ppvPaginaAtualMenu = "ppv_ordens_servico_triagem_despacho.asp" Then%> aria-current="page"<%End If%>>Triagem Despacho</a>
    <%Else%>
    <button type="button" class="ppv-shared-item ppv-access-trigger" aria-haspopup="dialog">Triagem Despacho</button>
    <%End If%>
    <%If ppvPaginasPermitidas.Exists("ppv_ordens_servico_triagem_entrega.asp") And validaUsuario(2,,53,true) Then%>
    <a class="ppv-shared-item" href="ppv_ordens_servico_triagem_entrega.asp"<%If ppvPaginaAtualMenu = "ppv_ordens_servico_triagem_entrega.asp" Then%> aria-current="page"<%End If%>>Triagem Entrega</a>
    <%Else%>
    <button type="button" class="ppv-shared-item ppv-access-trigger" aria-haspopup="dialog">Triagem Entrega</button>
    <%End If%>
    <%If ppvPaginasPermitidas.Exists("ppv_ordens_servico_os_manual_origem_vendas.asp") And validaUsuario(2,,58,true) Then%>
    <a class="ppv-shared-item" href="ppv_ordens_servico_os_manual_origem_vendas.asp"<%If ppvPaginaAtualMenu = "ppv_ordens_servico_os_manual_origem_vendas.asp" Then%> aria-current="page"<%End If%>>Inser&ccedil;&atilde;o de OS Manual com Origem em Pedido</a>
    <%Else%>
    <button type="button" class="ppv-shared-item ppv-access-trigger" aria-haspopup="dialog">Inser&ccedil;&atilde;o de OS Manual com Origem em Pedido</button>
    <%End If%>
    <%If ppvPaginasPermitidas.Exists("ppv_ordens_servico_triagem_entrega_correios.asp") And validaUsuario(2,,53,true) Then%>
    <a class="ppv-shared-item" href="ppv_ordens_servico_triagem_entrega_correios.asp"<%If ppvPaginaAtualMenu = "ppv_ordens_servico_triagem_entrega_correios.asp" Then%> aria-current="page"<%End If%>>Retorno Correios</a>
    <%Else%>
    <button type="button" class="ppv-shared-item ppv-access-trigger" aria-haspopup="dialog">Retorno Correios</button>
    <%End If%>
</div>
<%
    FecharConexao()
End If
Set ppvPaginasPermitidas = Nothing
%>
<style>
.dashboard-main .ppv-module:has(> .ppv-shared-menu) { margin-bottom:0; }
.ppv-shared-menu {
    box-sizing:border-box; display:flex; flex-wrap:wrap; gap:6px;
    margin:0 0 20px; padding:10px; border:1px solid #e4dfeb;
    border-radius:12px; background:#fff; font:600 14px/1.35 system-ui,sans-serif;
}
.ppv-shared-menu .ppv-shared-item {
    appearance:none; box-sizing:border-box; display:inline-flex; align-items:center;
    min-height:42px; max-width:100%; margin:0; padding:8px 12px;
    border:1px solid transparent; border-radius:8px; background:transparent;
    color:#3e4057; font:inherit; text-align:left; text-decoration:none;
    white-space:normal; overflow-wrap:anywhere; cursor:pointer;
}
.ppv-shared-menu .ppv-shared-item:hover { background:#f7edf3; color:#88134f; text-decoration:none; }
.ppv-shared-menu .ppv-shared-item[aria-current="page"] { background:#b20b61; color:#fff; }
.ppv-shared-menu .ppv-access-trigger { color:#66677b; }
.ppv-shared-menu .ppv-shared-item:focus-visible { outline:2px solid #88134f; outline-offset:2px; }
@media(max-width:640px) {
    .ppv-shared-menu { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); }
}
@media(max-width:360px) {
    .ppv-shared-menu { grid-template-columns:minmax(0,1fr); }
}
.ppv-access-dialog {
    box-sizing:border-box; width:min(440px, calc(100% - 32px)); padding:24px;
    border:1px solid #e4dfeb; border-radius:14px; color:#242738;
    background:#fff; box-shadow:0 18px 60px #24273833; font:1rem/1.5 system-ui,sans-serif;
}
.ppv-access-dialog::backdrop { background:#24273899; }
.ppv-access-dialog h2 { margin:0; color:#242738; font:700 1.2rem/1.3 system-ui,sans-serif; }
.ppv-access-dialog p { margin:10px 0 20px; }
.ppv-access-dialog form { display:flex; justify-content:flex-end; margin:0; }
.ppv-access-dialog button {
    padding:9px 17px; border:0; border-radius:7px; color:#fff; background:#b20b61;
    font:600 .875rem/1.5 system-ui,sans-serif; cursor:pointer;
}
.ppv-access-dialog button:hover { background:#88134f; }
</style>
<dialog class="ppv-access-dialog" id="ppvAccessDialog" aria-labelledby="ppvAccessTitle" aria-describedby="ppvAccessMessage">
    <h2 id="ppvAccessTitle">Acesso restrito</h2>
    <p id="ppvAccessMessage">Seu perfil n&atilde;o tem permiss&atilde;o para acessar esta p&aacute;gina.</p>
    <form method="dialog"><button autofocus>Entendi</button></form>
</dialog>
<script>
(function () {
    var dialog = document.getElementById('ppvAccessDialog');
    document.querySelectorAll('.ppv-shared-menu .ppv-access-trigger').forEach(function (button) {
        button.addEventListener('click', function () {
            dialog.showModal();
        });
    });
})();
</script>
