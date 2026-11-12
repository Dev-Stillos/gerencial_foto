<div align="center" class="textoCinza">
    <a href="#" class="linkCinza">Inicio</a> |
    <a class="linkCinza" href="#" onclick="buscaIni()">Busca</a> |
    <a href="#" class="linkCinza">Topo</a><br />
    StillosWeb sistema administrativo interno - Desenvolvido pelo departamento de T.I.
    2008 - <%=Year(now())%>
</div>

<script>
(function () {
    function ajustarDropdowns() {
        document.querySelectorAll('select').forEach(function (campo) {
            Array.from(campo.options).forEach(function (opcao) {
                var textoCompleto = opcao.text;

                if (textoCompleto.length > 80) {
                    opcao.title = textoCompleto;
                    opcao.text = textoCompleto.slice(0, 80) + '...';
                }
            });
        });
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', ajustarDropdowns);
    } else {
        ajustarDropdowns();
    }
})();
</script>