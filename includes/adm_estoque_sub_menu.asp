
<% if perfil = 1 then %>
 <div class="subMenu">
       <table width="100%" border="0" cellspacing="0" cellpadding="3">
            <tr>
                <td width="6%" rowspan="2">Sub-Menu: </td>
                <td width="94%">
                    <a href="javascript:listarMovimentacoes()">Ultimas Entradas/Sa&iacutedas</a> | 
                    <a href="javascript:listaItensEstoque()">Itens em estoque</a> |
                    <a href="javascript:lancarMovimentacao()">Registrar Entrada/Sa&iacuteda</a>  |
                    <a href="javascript:cadastrarPedidos()">Novo pedido</a>  |                          
                    <a href="javascript:listarPedidos()">Listar pedidos</a> |
                    <a href="est_movimentacao_comprovante_trasnferencia.asp">Comprovante de Transfer&ecircncia</a>
                </td>
            </tr>
        </table>
</div>
<%else if perfil = 2 then %>

 <div class="subMenu">
       <table width="100%" border="0" cellspacing="0" cellpadding="3">
            <tr>
                <td width="6%" rowspan="2">Sub-Menu: </td>
                <td width="94%">
                    <a href="javascript:listarMovimentacoes()">Ultimas Entradas/Sa&iacutedas</a> | 
                    <a href="javascript:listaItensEstoque()">Itens em estoque</a> |
                    <a href="javascript:lancarMovimentacao()">Registrar Entrada/Sa&iacuteda</a>  |                
                </td>
            </tr>
        </table>
</div>

<%end if %>
<%end if %>


<!-- Antigo 
<div class="subMenu">
       <table width="100%" border="0" cellspacing="0" cellpadding="3">
            <tr>
                <td width="6%" rowspan="2">Sub-Menu: </td>
                <td width="94%">
                    <a href="javascript:listarMovimentacoes()">Ultimas Entradas/Sa&iacutedas</a> | 
                    <a href="javascript:listaItensEstoque()">Itens em estoque</a> |
                    <a href="javascript:lancarMovimentacao()">Registrar Entrada/Sa&iacuteda</a>  |

                    <a href="javascript:cadastrarPedidos()">Novo pedido</a>                     
                    
                    <br><br>
                    <a href="javascript:listarItens()">Listar itens</a> | 
                    <a href="javascript:listarCategorias()">Listar categorias</a> | 
                    <a href="javascript:listarLocais()">Listar locais</a> |
                    <a href="javascript:listarFornecedores()">Listar fornecedores</a> | 
                    <a href="javascript:listarPedidos()">Listar pedidos</a> | 
                    <a href="javascript:listarRelItens()">Listar Rela&ccedil&otildees Item - produ&ccedil&atildeo</a> |  
                    <a href="javascript:listarRelItensVenda()">Listar Rela&ccedil&otildees Item - venda</a> |                                           
                       <br><br>
                    <a href="javascript:cadastrarItens()">Cadastrar item</a> |  
                    <a href="javascript:cadastrarCategorias()">Cadastrar categoria</a> |
                    <a href="javascript:cadastrarLocais()">Cadastrar locais</a> |
                    <a href="javascript:cadastrarFornecedores()">Cadastrar fornecedor</a> |    
                    <a href="javascript:cadastrarRelItens()">Cadastrar Rela&ccedil&atildeo Item - produ&ccedil&atildeo</a> |  
                    <a href="javascript:cadastrarRelItensVenda()">Cadastrar Rela&ccedil&atildeo Item - venda</a> |                       
                </td>
            </tr>
        </table>
</div>
-->
