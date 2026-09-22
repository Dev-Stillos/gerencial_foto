<%@ Language="VBScript" %>
<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'==============================================================================
' RELATÓRIO: FORMULÁRIO GOOGLE vs SISTEMA
' Sistema: Stillos
' Data: 27/01/2026
' Descrição: Exibe relatório consolidado com opção de exportar para Excel
'==============================================================================

On Error Resume Next

' Verificar se o usuário tem permissão (ajuste conforme seu sistema)
If Trim(session("vUSER_ID")) = "" Then
    Response.Redirect "index.asp"
    Response.End
End If

Dim objConn, objRs, strSQL, filtroTurma
Dim contador, valorCampo

' Parâmetros de filtro
filtroTurma = Trim(Request("filtro_turma"))

' Criar conexão com banco de dados
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open session("conn_string")

' Query para extrair dados (sem dependência de View)
strSQL = "SELECT " & _
    "dbo.TBL_DEPARA_CONTRATO.CODIGO_CONTRATO, " & _
    "dbo.TBL_DEPARA_CONTRATO.NRO_CONTROLE AS TURMA, " & _
    "dbo.TBL_CLIENTE_NRO_CONTROLE.ID_CLIENTE AS COD_FORMANDO, " & _
    "UPPER(dbo.TBL_CLIENTES_DADOS.NOME_CLIENTE) AS NOME_SISTEMA, " & _
    "UPPER(dbo.TB_FORMULARIO_GOOGLE.nome) AS NOME_FORMULARIO_GOOGLE, " & _
    "dbo.TB_FORMULARIO_GOOGLE.cpf AS CPF_FORMULARIO, " & _
    "dbo.TBL_CLIENTES_DADOS.CPFCNPJ_CLIENTE AS CPF_SISTEMA, " & _
    "dbo.TB_FORMULARIO_GOOGLE.data AS DATA_FORMULARIO " & _
    "FROM dbo.TBL_DEPARA_CONTRATO " & _
    "INNER JOIN dbo.TBL_CONTRATO ON dbo.TBL_DEPARA_CONTRATO.NRO_CONTROLE = dbo.TBL_CONTRATO.NRO_CONTROLE " & _
    "INNER JOIN dbo.TBL_CLIENTE_NRO_CONTROLE ON dbo.TBL_CONTRATO.NRO_CONTROLE = dbo.TBL_CLIENTE_NRO_CONTROLE.NRO_CONTROLE " & _
    "INNER JOIN dbo.TBL_CLIENTES_DADOS ON dbo.TBL_CLIENTE_NRO_CONTROLE.ID_CLIENTE = dbo.TBL_CLIENTES_DADOS.ID_CLIENTE " & _
    "INNER JOIN dbo.TB_FORMULARIO_GOOGLE ON dbo.TBL_CLIENTES_DADOS.CPFCNPJ_CLIENTE = dbo.TB_FORMULARIO_GOOGLE.cpf " & _
    "WHERE dbo.TBL_CLIENTES_DADOS.CPFCNPJ_CLIENTE <> '000.000.000-00'"

If filtroTurma <> "" Then
    strSQL = strSQL & " AND dbo.TBL_DEPARA_CONTRATO.NRO_CONTROLE LIKE '%" & Replace(filtroTurma, "'", "''") & "%'"
End If

strSQL = strSQL & " ORDER BY TURMA, NOME_SISTEMA"

Set objRs = objConn.Execute(strSQL)
%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Relatório: Formulário Google vs Sistema</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f5f5f5;
            color: #333;
        }
        
        .container {
            max-width: 1400px;
            margin: 20px auto;
            padding: 20px;
            background-color: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        
        h1 {
            color: #366092;
            margin-bottom: 10px;
            font-size: 24px;
        }
        
        .info-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            padding: 15px;
            background-color: #f9f9f9;
            border-left: 4px solid #366092;
        }
        
        .info-text {
            font-size: 14px;
            color: #666;
        }
        
        .info-text strong {
            color: #333;
        }
        
        .filtros {
            background-color: #f0f0f0;
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 20px;
        }
        
        .filtros form {
            display: flex;
            gap: 10px;
            align-items: flex-end;
            flex-wrap: wrap;
        }
        
        .form-group {
            display: flex;
            flex-direction: column;
        }
        
        .form-group label {
            font-weight: bold;
            margin-bottom: 5px;
            font-size: 13px;
        }
        
        .form-group input {
            padding: 8px 12px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 13px;
        }
        
        .botoes {
            display: flex;
            gap: 10px;
        }
        
        button {
            padding: 8px 16px;
            border: none;
            border-radius: 4px;
            font-size: 13px;
            font-weight: bold;
            cursor: pointer;
            transition: background-color 0.3s;
        }
        
        .btn-buscar {
            background-color: #4CAF50;
            color: white;
        }
        
        .btn-buscar:hover {
            background-color: #45a049;
        }
        
        .btn-limpar {
            background-color: #999;
            color: white;
        }
        
        .btn-limpar:hover {
            background-color: #888;
        }
        
        .btn-excel {
            background-color: #366092;
            color: white;
            text-decoration: none;
            display: inline-block;
        }
        
        .btn-excel:hover {
            background-color: #254570;
        }
        
        .acoes {
            display: flex;
            gap: 10px;
            margin-bottom: 20px;
            align-items: center;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }
        
        table th {
            background-color: #366092;
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: bold;
            border: 1px solid #245570;
        }
        
        table td {
            padding: 10px 12px;
            border: 1px solid #ddd;
        }
        
        table tbody tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        
        table tbody tr:hover {
            background-color: #f0f0f0;
        }
        
        .texto-center {
            text-align: center;
        }
        
        .texto-direita {
            text-align: right;
        }
        
        .footer {
            margin-top: 20px;
            padding-top: 15px;
            border-top: 1px solid #ddd;
            font-size: 12px;
            color: #999;
            text-align: right;
        }
        
        .mensagem-vazia {
            padding: 20px;
            text-align: center;
            color: #999;
            font-size: 14px;
        }
        
        .destaque-diferenca {
            background-color: #fff3cd;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 Relatório: Formulário Google vs Sistema Stillos</h1>
        
        <div class="info-header">
            <div class="info-text">
                <strong>Data de Extração:</strong> <%=Format(Now(), "dd/mm/yyyy HH:mm:ss")%>
            </div>
            <div class="info-text">
                <strong>Usuário:</strong> <%=session("vUSER_NOME")%>
            </div>
        </div>
        
        <!-- FILTROS -->
        <div class="filtros">
            <form method="GET" action="">
                <div class="form-group">
                    <label for="filtro_turma">Filtrar por Turma:</label>
                    <input type="text" id="filtro_turma" name="filtro_turma" value="<%=filtroTurma%>" placeholder="Ex: 01, 02, etc">
                </div>
                <div class="botoes">
                    <button type="submit" class="btn-buscar">🔍 Buscar</button>
                    <button type="button" class="btn-limpar" onclick="window.location=''">✖ Limpar</button>
                </div>
            </form>
        </div>
        
        <!-- AÇÕES -->
        <div class="acoes">
            <a href="adm_export_formulario_google_excel.asp<%If filtroTurma <> "" Then%>?filtro_turma=<%=Server.URLEncode(filtroTurma)%><%End If%>" class="btn-excel" target="_blank">
                📥 Exportar para Excel
            </a>
        </div>
        
        <!-- TABELA DE DADOS -->
        <%
        If objRs.EOF Then
        %>
            <div class="mensagem-vazia">
                ℹ️ Nenhum registro encontrado com os critérios de busca informados.
            </div>
        <%
        Else
        %>
            <table>
                <thead>
                    <tr>
                        <th>Código Contrato</th>
                        <th>Turma</th>
                        <th>Cod. Formando</th>
                        <th>Nome (Sistema)</th>
                        <th>Nome (Formulário)</th>
                        <th>CPF (Formulário)</th>
                        <th>CPF (Sistema)</th>
                        <th>Data Formulário</th>
                    </tr>
                </thead>
                <tbody>
        <%
            contador = 0
            Do While Not objRs.EOF
                contador = contador + 1
                
                Dim diferenca
                diferenca = ""
                If UCase(objRs("NOME_SISTEMA")) <> UCase(objRs("NOME_FORMULARIO_GOOGLE")) Then
                    diferenca = " destaque-diferenca"
                End If
        %>
                    <tr<%=diferenca%>>
                        <td><%=objRs("CODIGO_CONTRATO")%></td>
                        <td><%=objRs("TURMA")%></td>
                        <td class="texto-direita"><%=objRs("COD_FORMANDO")%></td>
                        <td><%=objRs("NOME_SISTEMA")%></td>
                        <td><%=objRs("NOME_FORMULARIO_GOOGLE")%></td>
                        <td><%=objRs("CPF_FORMULARIO")%></td>
                        <td><%=objRs("CPF_SISTEMA")%></td>
                        <td class="texto-center"><%=IIf(IsNull(objRs("DATA_FORMULARIO")) Or objRs("DATA_FORMULARIO") = "", "-", Format(CDate(objRs("DATA_FORMULARIO")), "dd/mm/yyyy"))%></td>
                    </tr>
        <%
                objRs.MoveNext
            Loop
        %>
                </tbody>
            </table>
            
            <div class="footer">
                <strong>Total de Registros:</strong> <%=contador%>
            </div>
        <%
        End If
        
        ' Fechar conexão
        objRs.Close
        objConn.Close
        Set objRs = Nothing
        Set objConn = Nothing
        %>
    </div>
</body>
</html>
