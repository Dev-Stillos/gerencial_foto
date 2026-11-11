<%
Dim ppvPaginasPermitidas, ppvRegistroPermissao, ppvConsultaPermissao, ppvPerfilMenu
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
<div class="subMenu">Submenu:
    <%If ppvPaginasPermitidas.Exists("ppv_ordens_servico.asp") Then%>
    <a href="ppv_ordens_servico.asp">Ordens de Servi&ccedil;o</a>
    <%Else%>
    <button type="button" class="ppv-access-trigger" aria-haspopup="dialog">Ordens de Servi&ccedil;o</button>
    <%End If%> |
    <%If ppvPaginasPermitidas.Exists("ppv_produtos_etapas.asp") And validaUsuario(2,,54,true) Then%>
    <a href="ppv_produtos_etapas.asp">Produtos Etapas</a>
    <%Else%>
    <button type="button" class="ppv-access-trigger" aria-haspopup="dialog">Produtos Etapas</button>
    <%End If%> |
    <%If ppvPaginasPermitidas.Exists("ppv_ordens_servico_triagem_inicial.asp") And validaUsuario(2,,52,true) Then%>
    <a href="ppv_ordens_servico_triagem_inicial.asp">Triagem Inicial</a>
    <%Else%>
    <button type="button" class="ppv-access-trigger" aria-haspopup="dialog">Triagem Inicial</button>
    <%End If%> |
    <%If ppvPaginasPermitidas.Exists("ppv_ordens_servico_triagem_despacho.asp") And validaUsuario(2,,53,true) Then%>
    <a href="ppv_ordens_servico_triagem_despacho.asp">Triagem Despacho</a>
    <%Else%>
    <button type="button" class="ppv-access-trigger" aria-haspopup="dialog">Triagem Despacho</button>
    <%End If%> |
    <%If ppvPaginasPermitidas.Exists("ppv_ordens_servico_triagem_entrega.asp") And validaUsuario(2,,53,true) Then%>
    <a href="ppv_ordens_servico_triagem_entrega.asp">Triagem Entrega</a>
    <%Else%>
    <button type="button" class="ppv-access-trigger" aria-haspopup="dialog">Triagem Entrega</button>
    <%End If%> |
    <%If ppvPaginasPermitidas.Exists("ppv_ordens_servico_os_manual_origem_vendas.asp") And validaUsuario(2,,58,true) Then%>
    <a href="ppv_ordens_servico_os_manual_origem_vendas.asp">Inser&ccedil;&atilde;o de OS Manual com Origem em Pedido</a>
    <%Else%>
    <button type="button" class="ppv-access-trigger" aria-haspopup="dialog">Inser&ccedil;&atilde;o de OS Manual com Origem em Pedido</button>
    <%End If%> |
    <%If ppvPaginasPermitidas.Exists("ppv_ordens_servico_triagem_entrega_correios.asp") And validaUsuario(2,,53,true) Then%>
    <a href="ppv_ordens_servico_triagem_entrega_correios.asp">Retorno Correios</a>
    <%Else%>
    <button type="button" class="ppv-access-trigger" aria-haspopup="dialog">Retorno Correios</button>
    <%End If%>
</div>
<%
    FecharConexao()
End If
Set ppvPaginasPermitidas = Nothing
%>
<style>
.subMenu .ppv-access-trigger {
    appearance:none; display:inline; padding:0; border:0; background:none;
    color:inherit; font:inherit; text-decoration:underline; cursor:pointer;
}
.subMenu .ppv-access-trigger:hover { color:#88134f; }
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
    document.querySelectorAll('.subMenu .ppv-access-trigger').forEach(function (button) {
        button.addEventListener('click', function () {
            dialog.showModal();
        });
    });
})();
</script>
