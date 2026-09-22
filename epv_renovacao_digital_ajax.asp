<%
'=========================================
' AJAX para Renovação Digital e Fichas
'=========================================
COD_USUARIO = sqlNumero(Session("COD_USUARIO"))
ts = Request.QueryString("ts")

If ts = "listar" Then
    AbrirConexao()
    Set rs = objConexao.Execute("SELECT * FROM TB_RENOVACAO_ITENS_DIGITAIS WHERE COD_USUARIO=" & COD_USUARIO & " ORDER BY DATA_CRIACAO DESC")
    Do While Not rs.Eof
        Response.Write "<div>" & rs("COD_FORMANDO") & " - Tipo: " & rs("TIPO") & " - Prazo: " & rs("DATA_PRAZO") & _
                       " <button onclick='removerItemRenovacao(" & rs("COD_FORMANDO") & ")'>Remover</button></div>"
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    FecharConexao()
End If

If ts = "incluir" Then
    codFormando = sqlNumero(Request.QueryString("p1"))
    tipo = 2 ' álbum digital
    periodoRenovacao = 30
    dataPrazo = sqlData(DateAdd("d", periodoRenovacao, Date()))
    AbrirConexao()
    objConexao.Execute "INSERT INTO TB_RENOVACAO_ITENS_DIGITAIS (COD_FORMANDO, COD_USUARIO, TIPO, PERIODO_RENOVACAO, DATA_PRAZO) " & _
                       "VALUES (" & codFormando & ", " & COD_USUARIO & ", " & tipo & ", " & periodoRenovacao & ", " & dataPrazo & ")"
    FecharConexao()
    Response.Write "Item " & codFormando & " adicionado com sucesso!"
End If

If ts = "remover" Then
    codFormando = sqlNumero(Request.QueryString("p1"))
    AbrirConexao()
    objConexao.Execute "DELETE FROM TB_RENOVACAO_ITENS_DIGITAIS WHERE COD_USUARIO=" & COD_USUARIO & " AND COD_FORMANDO=" & codFormando
    FecharConexao()
    Response.Write "Item " & codFormando & " removido!"
End If
%>
