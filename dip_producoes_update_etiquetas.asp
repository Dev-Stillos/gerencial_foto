<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<%
Response.Charset = "utf-8"

Dim codUsuario, codForm, status
codUsuario = 0
If IsNumeric(Session("COD_USUARIO")) Then codUsuario = CLng(Session("COD_USUARIO"))

codForm = Request("cod_form")
status  = Request("status")

If codUsuario > 0 And IsNumeric(codForm) And IsNumeric(status) Then
    AbrirConexao()

    ' Tenta atualizar
    sql = "UPDATE TB_FICHAS_IDENTIFICACAO_ETIQUETAS_TEMPORARIA " & _
          "SET IMPRIMIR_ETIQUETA=" & status & " " & _
          "WHERE COD_USUARIO=" & codUsuario & " AND COD_FORMANDO=" & codForm & " AND AGENDADA=2"
    objConexao.Execute(sql)

    ' Verifica se update afetou algo
    Set rs = objConexao.Execute("SELECT ROW_COUNT() AS AFETADOS")
    If rs("AFETADOS") = 0 Then
        ' Faz insert se não existe
        sql = "INSERT INTO TB_FICHAS_IDENTIFICACAO_ETIQUETAS_TEMPORARIA " & _
              "(COD_USUARIO, COD_FORMANDO, IMPRIMIR_ETIQUETA, AGENDADA) " & _
              "VALUES (" & codUsuario & "," & codForm & "," & status & ",2)"
        objConexao.Execute(sql)
    End If
    rs.Close
    Set rs = Nothing
End If

Response.End
%>
