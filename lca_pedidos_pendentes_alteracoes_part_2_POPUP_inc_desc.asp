<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<%
' Captura o cod_pedido da URL
Dim COD_PEDIDO
COD_PEDIDO = Trim(Request.QueryString("cod_pedido"))

' Validação de segurança
If Not IsNumeric(COD_PEDIDO) Or COD_PEDIDO = "" Then
    Response.Write "<div style='color:red; text-align:center;'>Erro: Código do pedido inválido.</div>"
    Response.End
End If

' Declaração e execução da conexão e query
Dim conn, rsFormandos
Set conn = AbrirConexaoRapida()

On Error Resume Next
Set rsFormandos = conn.Execute("SELECT F.NOME_FORMANDO, PF.COD_FORMANDO FROM TB_PEDIDOS_FORMANDOS PF INNER JOIN TB_FICHAS_IDENTIFICACAO F ON PF.COD_FORMANDO = F.COD_FORMANDO WHERE PF.COD_PEDIDO = " & COD_PEDIDO)
If Err.Number <> 0 Then
    Response.Write "<div style='color:red; text-align:center;'>Erro na consulta SQL: " & Err.Description & "</div>"
    Set rsFormandos = Nothing
End If
On Error GoTo 0
%>

<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>Inclusão de Desconto</title>
<link rel="stylesheet" type="text/css" href="css/base.css" />
<style>
body {
    background-color: #FFFFFF;
    font-family: Verdana;
    font-size: 12px;
}
input[type="text"], select {
    padding: 5px;
    width: 140px;
    text-align: center;
    font-size: 12px;
}
</style>
<script>
function validaFormDesconto() {
    var d = document.form_desc;
    if (d.COD_FORMANDO.value == "") {
        alert("Selecione o formando.");
        d.COD_FORMANDO.focus();
        return false;
    }
    if (d.DESC_ALBUM.value == "" || isNaN(d.DESC_ALBUM.value)) {
        alert("Informe o desconto do álbum (%).");
        d.DESC_ALBUM.focus();
        return false;
    }
    if (d.DESC_PRODUTO.value == "" || isNaN(d.DESC_PRODUTO.value)) {
        alert("Informe o desconto dos produtos vendidos à parte (%).");
        d.DESC_PRODUTO.focus();
        return false;
    }
    return true;
}
</script>
</head>
<body>
<form method="post" name="form_desc" action="lca_pedidos_pendentes_alteracoes_part_2.asp?cod_pedido=<%=COD_PEDIDO%>" onsubmit="return validaFormDesconto()" target="_parent">
    <table width="100%" border="1" cellpadding="5" style="border-collapse:collapse;" align="center">


        <tr>
            <td align="center"><strong>Formando</strong></td>
            <td align="center"><strong>Desc% - Álbum</strong></td>
            <td align="center"><strong>Desc% - Produtos vendidos à parte</strong></td>
        </tr>
        <tr>
            <td align="center">
                <select name="COD_FORMANDO">
                    <option value="">Selecione</option>
                    <%
                    If Not rsFormandos Is Nothing Then
                        Do While Not rsFormandos.Eof
                    %>
                        <option value="<%=rsFormandos("COD_FORMANDO")%>"><%=rsFormandos("NOME_FORMANDO")%></option>
                    <%
                        rsFormandos.MoveNext
                        Loop
                    Else
                        Response.Write "<option disabled>Erro ao carregar formandos</option>"
                    End If
                    %>
                </select>
            </td>
            <td align="center">
                <input type="text" name="DESC_ALBUM" maxlength="5" placeholder="Ex: 10" />
            </td>
            <td align="center">
                <input type="text" name="DESC_PRODUTO" maxlength="5" placeholder="Ex: 5" />
            </td>
        </tr>
        <tr>
            </td>
            <tr>
    <td colspan="3" align="center">
        <input type="submit" name="GRAVAR_DESCONTO" value="Gravar Desconto"
            style="padding: 8px 16px; font-weight: bold; color: white; background-color: #ff007f; border: none; border-radius: 6px; cursor: pointer;" />
    </td>
</tr>
        </tr>
    </table>
</form>
<%
' Fechamento seguro
If Not rsFormandos Is Nothing Then
    rsFormandos.Close
    Set rsFormandos = Nothing
End If

If Not conn Is Nothing Then
    conn.Close
    Set conn = Nothing
End If
%>
</body>
</html>
