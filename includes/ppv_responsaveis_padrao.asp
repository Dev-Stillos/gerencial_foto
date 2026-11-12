<%
Function PPVPodeConfigurarResponsaveis()
    PPVPodeConfigurarResponsaveis = False
    If Len(Trim(Session("cod_usuario") & "")) = 0 Then Exit Function
    Select Case CStr(Session("cod_perfil") & "")
        Case "3", "12", "42"
            PPVPodeConfigurarResponsaveis = True
    End Select
End Function

Function PPVUsuarioElegivel(codigo)
    Dim formato, cadastro
    PPVUsuarioElegivel = False
    Set formato = New RegExp
    formato.Pattern = "^[0-9]{1,9}$"
    If Not formato.Test(codigo) Then Exit Function
    Set cadastro = objConexao.Execute("SELECT COD_USUARIO FROM TB_ADM_USUARIOS WHERE BLOQUEADO = 2 AND COD_USUARIO > 1 AND COD_USUARIO = " & CLng(codigo))
    PPVUsuarioElegivel = Not cadastro.EOF
    cadastro.Close
End Function

Function PPVResponsavelPadrao(tipo)
    Dim coluna, cadastro
    PPVResponsavelPadrao = ""
    Select Case UCase(Trim(tipo & ""))
        Case "S": coluna = "COD_USUARIO_PRE"
        Case "N": coluna = "COD_USUARIO_POS"
        Case Else: Exit Function
    End Select
    Set cadastro = objConexao.Execute("SELECT u.COD_USUARIO FROM TB_PPV_RESPONSAVEIS_PADRAO c INNER JOIN TB_ADM_USUARIOS u ON u.COD_USUARIO = c." & coluna & " WHERE c.ID = 1 AND u.BLOQUEADO = 2 AND u.COD_USUARIO > 1")
    If Not cadastro.EOF Then PPVResponsavelPadrao = CStr(cadastro("COD_USUARIO"))
    cadastro.Close
End Function
%>
