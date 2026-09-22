<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<%
Dim cod_pedido, parcela
cod_pedido = Request.QueryString("cod_pedido")
parcela = Request.QueryString("parcela")

AbrirConexao()

Dim rs, rsVend, nome_vendedor, cod_vendedor, valor, valor_comissao, data_vencimento

' Buscar dados da parcela específica
Set rs = objConexao.Execute("SELECT cod_pedido, cod_vendedor, parcela, valor, valor_comissao_vendedor, data_vencimento FROM TB_PEDIDOS_BOLETOS WHERE cod_pedido = " & cod_pedido & " AND parcela = " & parcela)

If Not rs.Eof Then
    cod_vendedor = rs("cod_vendedor")
    valor = rs("valor")
    valor_comissao = rs("valor_comissao_vendedor")
    data_vencimento = rs("data_vencimento")

    ' Busca nome do vendedor
    Set rsVend = objConexao.Execute("SELECT nome_vendedor FROM TB_VENDEDORES WHERE cod_vendedor = " & cod_vendedor)
    If Not rsVend.Eof Then
        nome_vendedor = rsVend("nome_vendedor")
    Else
        nome_vendedor = "NÃO ENCONTRADO"
    End If
    rsVend.Close
    Set rsVend = Nothing
End If

' Data de hoje formatada
Dim dia, mes, ano
dia = Day(Date())
mes = MonthName(Month(Date()))
ano = Year(Date())
%>
<html>
<head>
<meta charset="utf-8">
<title>Recibo Pedido <%=cod_pedido%> - Parcela <%=parcela%></title>
<style>
body {
    font-family: Arial;
    font-size: 16px;
    margin: 60px;
    color: #000;
}
h2 {
    text-align: center;
    border-bottom: 1px solid #000;
    padding-bottom: 5px;
    width: 200px;
    margin: 0 auto 30px auto;
}
p { text-align: justify; line-height: 1.6; }
.detalhes { margin-top: 20px; line-height: 1.8; }
.assinatura { margin-top: 60px; text-align: left; }
</style>
</head>
<body onload="window.print()">

<h2>RECIBO DE COMISSÃO</h2>

<p>
Eu <strong><%=(nome_vendedor)%></strong> recebi da empresa: <strong>Stillo's Reportagens Fotográficas S/C Ltda.</strong>, 
a importância de <strong>R$ <%=FormatNumber(valor_comissao,2)%></strong> 
(de reais) referente à <strong>comissão nas vendas</strong> do pedido nº 
<strong><%=cod_pedido%></strong> (parcela <%=parcela%>).
</p>

<p>Sem nada a reclamar, assino.</p>

<div class="detalhes">
   <!--<strong>Valor da Parcela:</strong> R$ <%=FormatNumber(valor,2)%><br>-->
   <!--<strong>Data de Vencimento:</strong> <%=Day(data_vencimento)%>/<%=Month(data_vencimento)%>/<%=Year(data_vencimento)%><br>-->
    <strong>Valor da Comissão:</strong> R$ <%=FormatNumber(valor_comissao,2)%>
</div>

<p style="margin-top:30px;">
<strong>São Paulo, <%=dia%> de <%=mes%> de <%=ano%>.</strong>
</p>

<div class="assinatura">
<p>-----------------------------------------------<br>
<em><%=UCase(nome_vendedor)%></em></p>
</div>

</body>
</html>
<%
rs.Close
Set rs = Nothing
FecharConexao()
%>
