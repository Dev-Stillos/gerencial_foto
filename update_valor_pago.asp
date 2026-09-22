<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<%
Dim codPedido, parcela, valorRaw, fechamento, sql
Dim integerPart, decimalPart, parts, sepPos, valorParaSQL

codPedido   = Trim(Request.Form("COD_PEDIDO"))
parcela     = Trim(Request.Form("PARCELA"))
valorRaw    = Trim(Request.Form("VALOR_PAGO"))
fechamento  = Trim(Request.Form("COD_FECHAMENTO_PERIODO"))

' Segurança básica
If Not IsNumeric(codPedido) Or Not IsNumeric(parcela) Then
    Response.Write("Parâmetros inválidos.")
    Response.End
End If

codPedido = CLng(codPedido)
parcela   = CLng(parcela)

' Normaliza espaços
valorRaw = Replace(valorRaw, " ", "")

' Se veio vazio, define 0
If valorRaw = "" Then
    integerPart = "0"
    decimalPart = "00"
Else
    ' Caso 1: existe vírgula -> formato BR esperado (1.234,56 ou 303,61)
    If InStr(valorRaw, ",") > 0 Then
        parts = Split(valorRaw, ",")
        integerPart = Replace(parts(0), ".", "")    ' remove pontos de milhar
        If UBound(parts) >= 1 Then
            decimalPart = parts(1)
        Else
            decimalPart = ""
        End If
        ' garante apenas dígitos
        integerPart = RegExpKeepDigits(integerPart)
        decimalPart = RegExpKeepDigits(decimalPart)
        decimalPart = Left(decimalPart & "00", 2)   ' garante 2 casas (padded ou truncated)
        
    ' Caso 2: não tem vírgula mas pode ter ponto como decimal (303.61) ou pontos de milhar (1.234)
    ElseIf InStr(valorRaw, ".") > 0 Then
        sepPos = InStrRev(valorRaw, ".")
        ' se há um ponto e existem exatamente 2 dígitos depois dele, consideramos ponto como separador decimal
        If Len(valorRaw) - sepPos = 2 Then
            integerPart = Replace(Left(valorRaw, sepPos-1), ".", "")
            decimalPart = Mid(valorRaw, sepPos+1)
            integerPart = RegExpKeepDigits(integerPart)
            decimalPart = RegExpKeepDigits(decimalPart)
            decimalPart = Left(decimalPart & "00", 2)
        Else
            ' pontos são de milhar, sem parte decimal
            integerPart = Replace(valorRaw, ".", "")
            integerPart = RegExpKeepDigits(integerPart)
            decimalPart = "00"
        End If
    Else
        ' nenhum separador — inteiro puro
        integerPart = RegExpKeepDigits(valorRaw)
        decimalPart = "00"
    End If

    ' segurança final: se integerPart vazio -> zero
    If integerPart = "" Then integerPart = "0"
End If

' monta string no formato americano para SQL: inteiro.decimal
valorParaSQL = integerPart & "." & decimalPart

' Executa update
AbrirConexao()

If CDbl(valorParaSQL) > 0 Then
    ' Define a data de pagamento como a data atual
    sql = "UPDATE TB_PEDIDOS_BOLETOS SET VALOR_PAGO = " & valorParaSQL & _
          ", DATA_PAGAMENTO = Now() " & _
          "WHERE COD_PEDIDO = " & codPedido & " AND PARCELA = " & parcela
Else
    ' Se o valor for zerado, limpa a data
    sql = "UPDATE TB_PEDIDOS_BOLETOS SET VALOR_PAGO = 0, DATA_PAGAMENTO = NULL " & _
          "WHERE COD_PEDIDO = " & codPedido & " AND PARCELA = " & parcela
End If

On Error Resume Next
objConexao.Execute sql
If Err.Number <> 0 Then
    Response.Write("Erro ao atualizar: " & Err.Description)
    Err.Clear
End If
On Error GoTo 0

FecharConexao()

' Redireciona
If IsNumeric(fechamento) Then
    Response.Redirect "adm_controle_no_extrato_comissao_vendedores_boletos.asp?COD_FECHAMENTO_PERIODO=" & CLng(fechamento)
Else
    Response.Redirect "adm_controle_no_extrato_comissao_vendedores_boletos.asp"
End If

' ============================
' Função auxiliar: mantém só dígitos (VBScript)
' ============================
Function RegExpKeepDigits(s)
    Dim re, m
    Set re = New RegExp
    re.Global = True
    re.Pattern = "\D"   ' tudo que não é dígito
    RegExpKeepDigits = re.Replace(s, "")
    Set re = Nothing
End Function
%>

