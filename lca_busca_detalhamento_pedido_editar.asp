<%
'=======================================================================================
' PÁGINA: lca_busca_detalhamento_pedido_editar.asp
' DESCRIÇÃO: Formulário para editar itens de um pedido (quantidades e valores)
' AUTOR: Sistema
' DATA: 2026-01-23
' SEGURANÇA: Validação de permissões, transações de banco, proteção de dados
'=======================================================================================

Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE
BTN_VOLTAR = 1
BTN_IMPRIMIR = 0
BTN_FONTE = 2

Dim cod_pedido, acao
cod_pedido = Request.QueryString("cod_pedido")
acao = Request.QueryString("acao")

' Verificar se está chegando do formulário de processamento
Dim msg_sucesso, msg_erro
msg_sucesso = Request.QueryString("msg_sucesso")
msg_erro = Request.QueryString("msg_erro")

%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->

<%
AbrirConexao()

' ======================= VALIDAÇÕES =======================

' 1. Validar permissão do usuário
Dim VALIDA_SUB22
VALIDA_SUB22 = formatNumber(0&inStr(sub_validacoes_paginas,"#22"),0)

If VALIDA_SUB22 = 0 Then
    Response.Write("<div class='msg' style='color:red; text-align:center;'><strong>❌ Você não possui permissão para editar pedidos.</strong></div>")
    FecharConexao()
    Response.End()
End If

' 2. Validar se o código do pedido existe
If cod_pedido = "" Or Not IsNumeric(cod_pedido) Then
    Response.Write("<div class='msg' style='color:red; text-align:center;'><strong>❌ Pedido inválido.</strong></div>")
    FecharConexao()
    Response.End()
End If

' 3. Carregar dados do pedido
Dim PAGAMENTO_PROCESSADO, FINALIZA_CORRECAO, PEDIDO_EXTRA, NOME_VENDEDOR, DATA_VENDA
Dim sql, rs, rs_itens

sql = "SELECT TB_PEDIDOS.PAGAMENTO_PROCESSADO, TB_PEDIDOS.FINALIZA_CORRECAO, TB_PEDIDOS.PEDIDO_EXTRA, "&_
      "TB_VENDEDORES.NOME_VENDEDOR, TB_PEDIDOS.DATA_VENDA " &_
      "FROM TB_PEDIDOS INNER JOIN TB_VENDEDORES ON TB_VENDEDORES.COD_VENDEDOR = TB_PEDIDOS.COD_VENDEDOR " &_
      "WHERE COD_PEDIDO="&cod_pedido

Set rs = objConexao.Execute(sql)

If rs.Eof Then
    Response.Write("<div class='msg' style='color:red; text-align:center;'><strong>❌ Pedido não encontrado no banco de dados.</strong></div>")
    FecharConexao()
    Response.End()
End If

PAGAMENTO_PROCESSADO = formatNumber(0&rs("PAGAMENTO_PROCESSADO"),0)
FINALIZA_CORRECAO = formatNumber(0&rs("FINALIZA_CORRECAO"),0)
PEDIDO_EXTRA = formatNumber(0&rs("PEDIDO_EXTRA"),0)
NOME_VENDEDOR = rs("NOME_VENDEDOR")
DATA_VENDA = rs("DATA_VENDA")

rs.Close()
Set rs = Nothing

' 4. Proteger se pagamento já foi processado
If int(PAGAMENTO_PROCESSADO) = 1 Then
    Response.Write("<div class='msg' style='color:red; text-align:center; padding:20px;'>" &_
                   "<strong>⚠️ Este pedido já teve o pagamento processado!</strong><br/>" &_
                   "Edições podem afetar a comissão do vendedor. Contate a administração." &_
                   "</div>")
    FecharConexao()
    Response.End()
End If

%>

<style type="text/css">
.tabela-edicao {
    width: 100%;
    border-collapse: collapse;
    margin: 15px 0;
}

.tabela-edicao th {
    background-color: #f0f0f0;
    padding: 10px;
    text-align: center;
    border: 1px solid #999;
    font-weight: bold;
}

.tabela-edicao td {
    padding: 8px;
    border: 1px solid #ddd;
}

.tabela-edicao input[type="text"],
.tabela-edicao input[type="number"] {
    width: 90%;
    padding: 5px;
    border: 1px solid #ccc;
    border-radius: 3px;
}

.valor-original {
    color: #999;
    font-size: 12px;
    font-style: italic;
}

.btn-grupo {
    margin: 20px 0;
    text-align: center;
}

.btn-grupo a,
.btn-grupo input[type="submit"] {
    padding: 10px 20px;
    margin: 0 5px;
    border: none;
    cursor: pointer;
    font-size: 14px;
    border-radius: 4px;
    text-decoration: none;
    display: inline-block;
}

.btn-salvar {
    background-color: #4CAF50;
    color: white;
}

.btn-salvar:hover {
    background-color: #45a049;
}

.btn-cancelar {
    background-color: #f44336;
    color: white;
    text-decoration: none;
}

.btn-cancelar:hover {
    background-color: #da190b;
}

.secao-titulo {
    background-color: #f9f9f9;
    padding: 10px;
    margin-top: 20px;
    font-weight: bold;
    border-left: 4px solid #2196F3;
}

.info-pedido {
    background-color: #e3f2fd;
    padding: 15px;
    border-radius: 4px;
    margin: 15px 0;
}

.msg-sucesso {
    background-color: #d4edda;
    color: #155724;
    padding: 12px;
    border-radius: 4px;
    margin: 15px 0;
    text-align: center;
}

.msg-erro {
    background-color: #f8d7da;
    color: #721c24;
    padding: 12px;
    border-radius: 4px;
    margin: 15px 0;
    text-align: center;
}
</style>

<table width="100%" border="0" cellspacing="0" cellpadding="5">
<tr>
    <td valign="top" style="border:dotted 1px #999999">

        <h2>✏️ Editar Itens do Pedido nº <%=COD_PEDIDO%></h2>

        <% If msg_sucesso <> "" Then %>
            <div class="msg-sucesso">✅ <%=msg_sucesso%></div>
        <% End If %>

        <% If msg_erro <> "" Then %>
            <div class="msg-erro">❌ <%=msg_erro%></div>
        <% End If %>

        <!-- INFORMAÇÕES DO PEDIDO -->
        <div class="info-pedido">
            <strong>Vendedor:</strong> <%=NOME_VENDEDOR%> &nbsp;|&nbsp;
            <strong>Data:</strong> <%=ArrumaData(DATA_VENDA)%> &nbsp;|&nbsp;
            <strong>Tipo:</strong> <%If int(PEDIDO_EXTRA) = 2 Then%>Extra<%Else%>Normal<%End If%>
        </div>

        <!-- FORMULÁRIO DE EDIÇÃO -->
        <form name="FORM_EDITAR_ITENS" method="post" action="lca_busca_detalhamento_pedido_processar_edicao.asp">

            <input type="hidden" name="cod_pedido" value="<%=cod_pedido%>" />
            <input type="hidden" name="acao" value="salvar" />

            <%
            ' ======================= CARREGAR ITENS PARA EDIÇÃO =======================

            ' Se é PEDIDO_EXTRA = 2, carrega itens de TB_FICHAS_IDENTIFICACAO
            If int(PEDIDO_EXTRA) = 2 Then
                
                sql = "SELECT " &_
                      "TB_FICHAS_IDENTIFICACAO.COD_FORMANDO, " &_
                      "TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, " &_
                      "TB_FICHAS_IDENTIFICACAO.INICIAL, " &_
                      "TB_FICHAS_IDENTIFICACAO.FATURADO, " &_
                      "TB_FICHAS_IDENTIFICACAO.VALOR_UN, " &_
                      "TB_FICHAS_IDENTIFICACAO.ALBUM, " &_
                      "TB_FICHAS_IDENTIFICACAO.NOME_CURSO " &_
                      "FROM TB_PEDIDOS_FORMANDOS " &_
                      "INNER JOIN TB_FICHAS_IDENTIFICACAO ON TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO " &_
                      "WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&cod_pedido&" " &_
                      "ORDER BY TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO"

                Set rs_itens = objConexao.Execute(sql)

                If Not rs_itens.Eof Then
            %>

            <div class="secao-titulo">📷 Álbuns / Formandos</div>

            <table class="tabela-edicao">
                <tr>
                    <th>Formando</th>
                    <th>Curso</th>
                    <th>Álbum</th>
                    <th>Qtd. Inicial</th>
                    <th>Qtd. Faturada</th>
                    <th>Valor Unitário</th>
                </tr>

                <%
                While Not rs_itens.Eof
                %>

                <tr>
                    <td><%=rs_itens("NOME_FORMANDO")%></td>
                    <td><%=rs_itens("NOME_CURSO")%></td>
                    <td><%=rs_itens("ALBUM")%></td>
                    
                    <!-- Qtd. Inicial -->
                    <td>
                        <input type="number" name="inicial_<%=rs_itens("COD_FORMANDO")%>" 
                               value="<%=rs_itens("INICIAL")%>" min="0" />
                        <div class="valor-original">(atual: <%=rs_itens("INICIAL")%>)</div>
                    </td>

                    <!-- Qtd. Faturada -->
                    <td>
                        <input type="number" name="faturado_<%=rs_itens("COD_FORMANDO")%>" 
                               value="<%=rs_itens("FATURADO")%>" min="0" />
                        <div class="valor-original">(atual: <%=rs_itens("FATURADO")%>)</div>
                    </td>

                    <!-- Valor Unitário -->
                    <td>
                        <input type="text" name="valor_un_<%=rs_itens("COD_FORMANDO")%>" 
                               value="<%=formatNumber(rs_itens("VALOR_UN"))%>" />
                        <div class="valor-original">(atual: R$ <%=formatNumber(rs_itens("VALOR_UN"))%>)</div>
                    </td>

                    <!-- Hidden fields para comparação no processamento -->
                    <input type="hidden" name="cod_formando_list" value="<%=rs_itens("COD_FORMANDO")%>" />
                    <input type="hidden" name="inicial_original_<%=rs_itens("COD_FORMANDO")%>" 
                           value="<%=rs_itens("INICIAL")%>" />
                    <input type="hidden" name="faturado_original_<%=rs_itens("COD_FORMANDO")%>" 
                           value="<%=rs_itens("FATURADO")%>" />
                    <input type="hidden" name="valor_un_original_<%=rs_itens("COD_FORMANDO")%>" 
                           value="<%=rs_itens("VALOR_UN")%>" />
                </tr>

                <%
                rs_itens.MoveNext()
                Wend
                %>

            </table>

            <%
                rs_itens.Close()
                Set rs_itens = Nothing

                End If ' If Not rs_itens.Eof

            End If ' If int(PEDIDO_EXTRA) = 2

            ' ======================= PRODUTOS À PARTE =======================

            sql = "SELECT TB_PEDIDOS_PRODUTOS.COD_PEDIDO_PRODUTO, " &_
                  "TB_PRODUTOS.NOME_PRODUTO, " &_
                  "TB_PEDIDOS_PRODUTOS.QUANTIDADE, " &_
                  "TB_PEDIDOS_PRODUTOS.VALOR_UN " &_
                  "FROM TB_PEDIDOS_PRODUTOS " &_
                  "INNER JOIN TB_PRODUTOS ON TB_PRODUTOS.COD_PRODUTO = TB_PEDIDOS_PRODUTOS.COD_PRODUTO " &_
                  "WHERE TB_PEDIDOS_PRODUTOS.COD_PEDIDO = "&cod_pedido&" " &_
                  "ORDER BY TB_PRODUTOS.NOME_PRODUTO"

            Set rs_itens = objConexao.Execute(sql)

            If Not rs_itens.Eof Then
            %>

            <div class="secao-titulo">📦 Produtos À Parte</div>

            <table class="tabela-edicao">
                <tr>
                    <th>Produto</th>
                    <th>Quantidade</th>
                    <th>Valor Unitário</th>
                    <th>Valor Total</th>
                </tr>

                <%
                While Not rs_itens.Eof
                    Dim valor_total_prod
                    valor_total_prod = rs_itens("QUANTIDADE") * rs_itens("VALOR_UN")
                %>

                <tr>
                    <td><%=rs_itens("NOME_PRODUTO")%></td>
                    
                    <!-- Quantidade -->
                    <td>
                        <input type="number" name="qtd_prod_<%=rs_itens("COD_PEDIDO_PRODUTO")%>" 
                               value="<%=rs_itens("QUANTIDADE")%>" min="0" />
                        <div class="valor-original">(atual: <%=rs_itens("QUANTIDADE")%>)</div>
                    </td>

                    <!-- Valor Unitário -->
                    <td>
                        <input type="text" name="valor_prod_<%=rs_itens("COD_PEDIDO_PRODUTO")%>" 
                               value="<%=formatNumber(rs_itens("VALOR_UN"))%>" />
                        <div class="valor-original">(atual: R$ <%=formatNumber(rs_itens("VALOR_UN"))%>)</div>
                    </td>

                    <!-- Valor Total (apenas visualização) -->
                    <td>
                        <strong>R$ <%=formatNumber(valor_total_prod)%></strong>
                    </td>

                    <!-- Hidden fields -->
                    <input type="hidden" name="cod_pedido_produto_list" value="<%=rs_itens("COD_PEDIDO_PRODUTO")%>" />
                    <input type="hidden" name="qtd_prod_original_<%=rs_itens("COD_PEDIDO_PRODUTO")%>" 
                           value="<%=rs_itens("QUANTIDADE")%>" />
                    <input type="hidden" name="valor_prod_original_<%=rs_itens("COD_PEDIDO_PRODUTO")%>" 
                           value="<%=rs_itens("VALOR_UN")%>" />
                </tr>

                <%
                rs_itens.MoveNext()
                Wend
                %>

            </table>

            <%
                rs_itens.Close()
                Set rs_itens = Nothing

            End If ' If Not rs_itens.Eof

            %>

            <!-- BOTÕES DE AÇÃO -->
            <div class="btn-grupo">
                <input type="submit" class="btn-salvar" value="💾 Salvar Alterações" />
                <a href="lca_busca_detalhamento_pedido.asp?cod_pedido=<%=cod_pedido%>" class="btn-cancelar">❌ Cancelar</a>
            </div>

        </form>

        <!-- INFORMAÇÕES ADICIONAIS -->
        <div style="background-color: #fff3cd; padding: 15px; margin-top: 20px; border-radius: 4px; border-left: 4px solid #ff9800;">
            <strong>⚠️ Avisos Importantes:</strong>
            <ul style="margin: 10px 0;">
                <li>As alterações serão gravadas imediatamente no banco de dados</li>
                <li>Um registro de auditoria será criado com data, hora e usuário responsável</li>
                <li>Os totais serão recalculados automaticamente</li>
                <li>Alterações após pagamento processado afetarão a comissão do vendedor</li>
            </ul>
        </div>

    </td>
</tr>
</table>

<%
FecharConexao()
%>

<!-- #include file="files/includes/miolo-inferior.asp" -->
