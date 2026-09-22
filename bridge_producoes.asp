<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<%
' ##############################################################################
' BRIDGE API - Produções Stillos (V2)
' Este arquivo fornece dados em JSON para o novo portal.
' Acesso seguro via chave de API.
' ##############################################################################

Response.ContentType = "application/json"
Response.Charset = "UTF-8"

' CHAVE DE SEGURANÇA - Mantenha esta chave em segredo
Dim MY_API_KEY : MY_API_KEY = "Stillos_Secure_Bridge_2026"
' Em ASP Classic, headers personalizados chegam via HTTP_NOME_DO_HEADER
Dim receivedKey : receivedKey = Request.ServerVariables("HTTP_X_API_KEY")
If receivedKey = "" Then receivedKey = Request.QueryString("key")

' Validação de segurança
If receivedKey <> MY_API_KEY Then
    Response.Status = "403 Forbidden"
    Response.Write "{""error"": ""Acesso negado""}"
    Response.End
End If

' Filtros
Dim busca_nome, busca_codigo, busca_contrato, filtro_status, pagina, tamanho_pagina
busca_nome = Trim(Request("busca_nome"))
busca_codigo = Trim(Request("busca_codigo"))
busca_contrato = Trim(Request("busca_contrato"))
filtro_status = 0&Request("filtro") ' 1=Finalizadas, 2=Não Finalizadas, 3=Digital
pagina = 0&Request("pagina")
If pagina = 0 Then pagina = 1
tamanho_pagina = 100 

AbrirConexao()

Dim sql, rs, rsSub, json
sql = "SELECT COD_PRODUCAO, DESCRICAO, DATA_FINALIZADA, PRODUCAO_FINALIZADA, PRODUTOS_AGREDADOS FROM TB_PRODUCAO p WHERE 1=1"

If filtro_status = 1 Then
    sql = sql & " AND p.PRODUCAO_FINALIZADA = 1"
ElseIf filtro_status = 2 Then
    sql = sql & " AND p.PRODUCAO_FINALIZADA = 2"
ElseIf filtro_status = 3 Then
    sql = sql & " AND EXISTS (SELECT 1 FROM TB_FICHAS_IDENTIFICACAO f WHERE f.COD_PRODUCAO = p.COD_PRODUCAO AND f.ALBUM_DIGITAL = 1)"
End If

If busca_nome <> "" Then
    sql = sql & " AND p.DESCRICAO LIKE '%" & Replace(busca_nome, "'", "''") & "%'"
End If

If busca_codigo <> "" Then
    sql = sql & " AND p.COD_PRODUCAO = '" & Replace(busca_codigo, "'", "''") & "'"
End If

If busca_contrato <> "" Then
    sql = sql & " AND EXISTS (SELECT 1 FROM TB_FICHAS_IDENTIFICACAO f2 WHERE f2.COD_PRODUCAO = p.COD_PRODUCAO AND f2.COD_CONTRATO LIKE '%" & Replace(busca_contrato, "'", "''") & "%')"
End If

sql = sql & " ORDER BY COD_PRODUCAO DESC LIMIT " & (pagina - 1) * tamanho_pagina & ", " & tamanho_pagina

Set rs = objConexao.Execute(sql)

json = "["

While Not rs.EOF
    Dim codProd : codProd = rs("COD_PRODUCAO")
    
    ' Sub-query: Contratos
    Dim contratos : contratos = ""
    Set rsSub = objConexao.Execute("SELECT DISTINCT(COD_CONTRATO) FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = " & codProd)
    While Not rsSub.EOF
        contratos = contratos & rsSub("COD_CONTRATO") & ", "
        rsSub.MoveNext
    Wend
    If Len(contratos) > 2 Then contratos = Left(contratos, Len(contratos) - 2)
    rsSub.Close
    
    ' Sub-query: Qtd Lotes
    Dim qtdLotes : qtdLotes = 0
    Set rsSub = objConexao.Execute("SELECT COUNT(*) FROM TB_PRODUCAO_LOTES WHERE COD_PRODUCAO = " & codProd)
    If Not rsSub.EOF Then qtdLotes = rsSub(0)
    rsSub.Close
    
    ' Sub-query: Álbuns, Fotos, Digital
    Dim qtdAlbuns, totalFotos, isDigital
    Set rsSub = objConexao.Execute("SELECT COUNT(*) as qtd, SUM(INICIAL) as fotos, MAX(ALBUM_DIGITAL) as digital FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = " & codProd)
    If Not rsSub.EOF Then
        qtdAlbuns = 0&rsSub("qtd")
        totalFotos = 0&rsSub("fotos")
        isDigital = (0&rsSub("digital") = 1)
    End If
    rsSub.Close

    Dim desc : desc = rs("DESCRICAO")
    desc = Replace(desc, """", "\""")
    desc = Replace(desc, vbCrLf, " ")

    json = json & "{"
    json = json & """id"": """ & codProd & ""","
    json = json & """nome"": """ & desc & ""","
    json = json & """contrato"": """ & contratos & ""","
    json = json & """dataFinalizada"": """ & rs("DATA_FINALIZADA") & ""","
    json = json & """qtdLotes"": " & qtdLotes & ","
    json = json & """qtdAlbuns"": " & qtdAlbuns & ","
    json = json & """totalFotos"": " & totalFotos & ","
    json = json & """isDigital"": " & LCase(CStr(isDigital)) & ","
    json = json & """originalStatus"": " & rs("PRODUCAO_FINALIZADA")
    json = json & "}"

    rs.MoveNext
    If Not rs.EOF Then json = json & ","
Wend

json = json & "]"

Response.Write json

rs.Close
Set rs = Nothing
FecharConexao()
%>
