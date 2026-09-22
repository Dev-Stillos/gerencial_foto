<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<%
Dim cod_pedido, comissao

cod_pedido = Request.QueryString("cod_pedido")
comissao   = Request.QueryString("comissao")

If cod_pedido <> "" And comissao <> "" Then
    AbrirConexao()
    sql = "UPDATE TB_FICHAS_IDENTIFICACAO SET COMISSAO=" & comissao & " WHERE COD_FORMANDO=" & cod_pedido
    objConexao.Execute(sql)
    FecharConexao()
    
    Response.Redirect "lca_busca_detalhamento_pedido.asp?cod_pedido=" & cod_pedido & "&msg=Comissão atualizada com sucesso!"
Else
    Response.Write "Parâmetros inválidos."
End If
%>
