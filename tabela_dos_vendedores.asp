<%
' ==========================
' CONEXÃO AO SQL SERVER
' ==========================
Dim conn, rs, sql, filtro_contrato, filtro_turma

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLOLEDB;Data Source=SPOTEC0001;Initial Catalog=stillos;User Id=usuario_admin_web;Password=F9289*#s;"

filtro_contrato = Request("filtro_contrato")
filtro_turma     = Request("filtro_turma")

sql = "SELECT CODIGO_CONTRATO, TURMA, VALOR_TOTAL, PARCELAS, VALOR_PARCELA, DATA_INICIAL " & _
      "FROM _GERA_TABELA_DE_PRECO WHERE 1=1 "

If filtro_contrato <> "" Then
    sql = sql & " AND CODIGO_CONTRATO = '" & filtro_contrato & "' "
End If

If filtro_turma <> "" Then
    sql = sql & " AND TURMA LIKE '%" & filtro_turma & "%' "
End If

sql = sql & " ORDER BY CODIGO_CONTRATO, DATA_INICIAL"

Set rs = conn.Execute(sql)
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="UTF-8">
<title>Tabela de Preços</title>

<style>
body {
    font-family: Arial;
    background:#383838;
    padding:20px;
    color:white;
}

/* Cards de filtros */
.card {
    background:#ff0095;
    padding:20px;
    border-radius:12px;
    margin-bottom:15px;
}

/* Inputs */
.card input, 
.card select {
    width:100%;
    padding:10px;
    margin-top:5px;
    border:0;
    border-radius:8px;
}

/* Botão */
.btnFiltro {
    margin-top:10px;
    width:100%;
    background:#222;
    padding:12px;
    border-radius:8px;
    border:0;
    color:white;
    cursor:pointer;
}

.table {
    width:100%;
    margin-top:15px;
    border-collapse: collapse;
    background:white;
    color:black;
    border-radius:8px;
    overflow:hidden;
}

.table th {
    background:#ff0095;
    padding:10px;
    color:white;
}

.table td {
    padding:10px;
    border-top:1px solid #ccc;
}

.table tr:nth-child(even) {
    background:#f6f6f6;
}
</style>
</head>
<body>

<h2>Tabela de Preços</h2>

<form method="GET">

<div class="card">
    <label><strong>Código do contrato</strong></label>
    <input type="text" name="filtro_contrato" value="<%=filtro_contrato%>">
</div>

<div class="card">
    <label><strong>Turma</strong></label>
    <input type="text" name="filtro_turma" value="<%=filtro_turma%>">
</div>

<button class="btnFiltro">Buscar</button>
</form>

<table class="table">
    <tr>
        <th>Contrato</th>
        <th>Turma</th>
        <th>Data</th>
        <th>Valor</th>
        <th>X</th>
        <th>Valor Parcela</th>
    </tr>

<%
If Not rs.EOF Then
    Do While Not rs.EOF
%>
    <tr>
        <td><%=rs("CODIGO_CONTRATO")%></td>
        <td><%=rs("TURMA")%></td>
        <td><%=Day(rs("DATA_INICIAL"))%>/<%=Month(rs("DATA_INICIAL"))%>/<%=Year(rs("DATA_INICIAL"))%></td>
        <td>R$ <%=FormatNumber(rs("VALOR_TOTAL"), 2)%></td>
        <td><%=rs("PARCELAS")%></td>
        <td>R$ <%=FormatNumber(rs("VALOR_PARCELA"), 2)%></td>
    </tr>
<%
        rs.MoveNext
    Loop
Else
%>
    <tr>
        <td colspan="6" align="center">Nenhum resultado encontrado</td>
    </tr>
<%
End If

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>
</table>

</body>
</html>
