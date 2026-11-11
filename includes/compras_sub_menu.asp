<%
Dim comprasPaginaAtual
comprasPaginaAtual = LCase(Request.ServerVariables("URL") & "")
comprasPaginaAtual = Mid(comprasPaginaAtual, InStrRev(comprasPaginaAtual, "/") + 1)
AbrirConexao()
%>
<div class="compras-shared-menu" role="navigation" aria-label="Compras">
    <a class="compras-shared-item" href="ppv_compras_pedidos.asp"<%If comprasPaginaAtual = "ppv_compras_pedidos.asp" Then%> aria-current="page"<%End If%>>Compras</a>
    <%If validaUsuario(2,,61,true) Then%>
    <a class="compras-shared-item" href="ppv_compras_pedidos_formulario.asp"<%If comprasPaginaAtual = "ppv_compras_pedidos_formulario.asp" And Len(Request.QueryString("comando") & "") = 0 Then%> aria-current="page"<%End If%>>Novo Pedido de Compra (manual)</a>
    <%Else%>
    <button type="button" class="compras-shared-item compras-access-trigger" aria-haspopup="dialog">Novo Pedido de Compra (manual)</button>
    <%End If%>
    <%If validaUsuario(2,,61,true) Then%>
    <a class="compras-shared-item" href="ppv_fornecedores.asp"<%If comprasPaginaAtual = "ppv_fornecedores.asp" Then%> aria-current="page"<%End If%>>Fornecedores</a>
    <%Else%>
    <button type="button" class="compras-shared-item compras-access-trigger" aria-haspopup="dialog">Fornecedores</button>
    <%End If%>
    <%If validaUsuario(2,,63,true) Then%>
    <a class="compras-shared-item" href="ppv_formas_recebimento.asp"<%If comprasPaginaAtual = "ppv_formas_recebimento.asp" Then%> aria-current="page"<%End If%>>Formas de Recebimento</a>
    <%Else%>
    <button type="button" class="compras-shared-item compras-access-trigger" aria-haspopup="dialog">Formas de Recebimento</button>
    <%End If%>
    <%If validaUsuario(2,,62,true) Then%>
    <a class="compras-shared-item" href="ppv_formas_pagamento.asp"<%If comprasPaginaAtual = "ppv_formas_pagamento.asp" Then%> aria-current="page"<%End If%>>Formas de Pagamento</a>
    <%Else%>
    <button type="button" class="compras-shared-item compras-access-trigger" aria-haspopup="dialog">Formas de Pagamento</button>
    <%End If%>
    <%If validaUsuario(2,,60,true) Then%>
    <a class="compras-shared-item" href="ppv_ordens_servico_facilitador_geracao_pedidos_compra.asp"<%If comprasPaginaAtual = "ppv_ordens_servico_facilitador_geracao_pedidos_compra.asp" Then%> aria-current="page"<%End If%>>Gera&ccedil;&atilde;o de Pedidos de Compra com origem em Ordens de Servi&ccedil;o</a>
    <%Else%>
    <button type="button" class="compras-shared-item compras-access-trigger" aria-haspopup="dialog">Gera&ccedil;&atilde;o de Pedidos de Compra com origem em Ordens de Servi&ccedil;o</button>
    <%End If%>
    <%If validaUsuario(2,,60,true) Then%>
    <a class="compras-shared-item" href="ppv_compras_pedidos_produtos_solicitados.asp"<%If comprasPaginaAtual = "ppv_compras_pedidos_produtos_solicitados.asp" Then%> aria-current="page"<%End If%>>Produtos Solicitados</a>
    <%Else%>
    <button type="button" class="compras-shared-item compras-access-trigger" aria-haspopup="dialog">Produtos Solicitados</button>
    <%End If%>
</div>
<style>
.compras-shared-menu {
    box-sizing:border-box; display:flex; flex-wrap:wrap; gap:6px;
    margin:0 0 20px; padding:10px; border:1px solid #e4dfeb;
    border-radius:12px; background:#fff; font:600 14px/1.35 system-ui,sans-serif;
}
.compras-shared-menu .compras-shared-item {
    appearance:none; box-sizing:border-box; display:inline-flex; align-items:center;
    min-height:42px; max-width:100%; margin:0; padding:8px 12px;
    border:1px solid transparent; border-radius:8px; background:transparent;
    color:#3e4057; font:inherit; text-align:left; text-decoration:none;
    white-space:normal; overflow-wrap:anywhere; cursor:pointer;
}
.compras-shared-menu .compras-shared-item:hover { background:#f7edf3; color:#88134f; text-decoration:none; }
.compras-shared-menu .compras-shared-item[aria-current="page"] { background:#b20b61; color:#fff; }
.compras-shared-menu .compras-access-trigger { color:#66677b; }
.compras-shared-menu .compras-shared-item:focus-visible { outline:2px solid #88134f; outline-offset:2px; }
@media(max-width:640px) {
    .compras-shared-menu { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); }
}
@media(max-width:360px) {
    .compras-shared-menu { grid-template-columns:minmax(0,1fr); }
}
.compras-access-dialog {
    box-sizing:border-box; width:min(440px, calc(100% - 32px)); padding:24px;
    border:1px solid #e4dfeb; border-radius:14px; color:#242738;
    background:#fff; box-shadow:0 18px 60px #24273833; font:1rem/1.5 system-ui,sans-serif;
}
.compras-access-dialog::backdrop { background:#24273899; }
.compras-access-dialog h2 { margin:0; color:#242738; font:700 1.2rem/1.3 system-ui,sans-serif; }
.compras-access-dialog p { margin:10px 0 20px; }
.compras-access-dialog form { display:flex; justify-content:flex-end; margin:0; }
.compras-access-dialog button {
    padding:9px 17px; border:0; border-radius:7px; color:#fff; background:#b20b61;
    font:600 .875rem/1.5 system-ui,sans-serif; cursor:pointer;
}
.compras-access-dialog button:hover { background:#88134f; }
</style>
<dialog class="compras-access-dialog" id="comprasAccessDialog" aria-labelledby="comprasAccessTitle" aria-describedby="comprasAccessMessage">
    <h2 id="comprasAccessTitle">Acesso restrito</h2>
    <p id="comprasAccessMessage">Seu perfil n&atilde;o tem permiss&atilde;o para acessar esta p&aacute;gina.</p>
    <form method="dialog"><button autofocus>Entendi</button></form>
</dialog>
<script>
(function () {
    var dialog = document.getElementById('comprasAccessDialog');
    document.querySelectorAll('.compras-shared-menu .compras-access-trigger').forEach(function (button) {
        button.addEventListener('click', function () {
            dialog.showModal();
        });
    });
})();
</script>
