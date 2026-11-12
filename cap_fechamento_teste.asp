<% Response.Buffer = True %>
<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Response.AddHeader "Cache-Control", "no-store"

Dim ftPeriodo, ftRegex, ftConexao, ftRs, ftErro
Dim ftData, ftAutorizacao, ftProcessado, ftElegiveis, ftDiretoria, ftPedidos, ftBoletos, ftComissao, ftBoletosElegiveis
ftPeriodo = Trim(Request.QueryString("periodo") & "")
If ftPeriodo = "" Then ftPeriodo = "394"
Set ftRegex = New RegExp
ftRegex.Pattern = "^[1-9][0-9]{0,8}$"
If Not ftRegex.Test(ftPeriodo) Then
    ftErro = "Informe um c&oacute;digo de per&iacute;odo v&aacute;lido."
Else
    Set ftConexao = Server.CreateObject("ADODB.Connection")
    On Error Resume Next
    ftConexao.Open "DSN=StillosFechamentoTeste"
    If Err.Number <> 0 Then ftErro = "Configure o DSN StillosFechamentoTeste no IIS."
    Err.Clear
    On Error GoTo 0
    If ftErro = "" Then
        Set ftRs = ftConexao.Execute("SELECT DATABASE()")
        If LCase(ftRs(0).Value & "") <> "stillos3_fechamento_teste" Then
            ftErro = "Conex&atilde;o recusada. Use o banco stillos3_fechamento_teste."
        End If
        ftRs.Close
    End If
    If ftErro = "" Then
        Set ftRs = ftConexao.Execute("SELECT DATE_FORMAT(data_pagamento,'%Y-%m-%d') AS data_iso, autorizacao_diretoria_para_finalizar, data_processado_pagamento FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE cod_fechamento_periodo=" & ftPeriodo)
        If ftRs.Eof Then
            ftErro = "Per&iacute;odo indispon&iacute;vel na c&oacute;pia."
        Else
            ftData = ftRs("data_iso").Value & ""
            ftAutorizacao = UCase(Trim(ftRs("autorizacao_diretoria_para_finalizar").Value & ""))
            ftProcessado = ftRs("data_processado_pagamento").Value & ""
        End If
        ftRs.Close
        If ftData = "" And ftErro = "" Then ftErro = "Per&iacute;odo sem data de pagamento."
    End If
    If ftErro = "" Then
        Set ftRs = ftConexao.Execute("SELECT IFNULL(SUM(FINALIZA_CORRECAO=1 AND PAGAMENTO_PROCESSADO=2),0),IFNULL(SUM(FINALIZA_CORRECAO=1 AND FINALIZA_DIRETORIA=2),0) FROM TB_PEDIDOS WHERE VENDEDOR_RECEBE_EM='" & ftData & "'")
        ftElegiveis = ftRs(0).Value
        ftDiretoria = ftRs(1).Value
        ftRs.Close
        Set ftRs = ftConexao.Execute("SELECT COUNT(*),IFNULL(SUM(VALOR_TOTAL_COMISSAO),0) FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_FECHAMENTO_PERIODO=" & ftPeriodo)
        ftPedidos = ftRs(0).Value
        ftComissao = ftRs(1).Value
        ftRs.Close
        Set ftRs = ftConexao.Execute("SELECT COUNT(*) FROM TB_FECHAMENTO_PAGAMENTO_BOLETOS WHERE COD_FECHAMENTO_PERIODO=" & ftPeriodo)
        ftBoletos = ftRs(0).Value
        ftRs.Close
        Set ftRs = ftConexao.Execute("SELECT COUNT(*) FROM TB_PEDIDOS_BOLETOS WHERE vendedor_recebe_em='" & ftData & "' AND cod_fechamento_periodo IS NULL AND data_baixa IS NOT NULL AND pagamento_processado=0")
        ftBoletosElegiveis = ftRs(0).Value
        ftRs.Close
    End If
    If ftConexao.State = 1 Then ftConexao.Close
End If
%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Fechamento teste | Stillos</title>
<style>
body {margin:0;background:#f4f0f4;color:#2e2530;font:16px/1.5 Arial,sans-serif}
main {max-width:960px;margin:24px auto;background:#fff;border-radius:14px;padding:24px}
h1 {margin:4px 0} .selo {color:#a9125c;font-weight:bold}
form {display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin:18px 0}
input,button {font:inherit;padding:7px 10px} button {background:#a9125c;color:#fff;border:0;border-radius:7px}
dl {display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:10px}
dl div {background:#fcf6fa;border:1px solid #eadde7;border-radius:8px;padding:12px}
dd {font-weight:bold;margin:4px 0 0}
.aviso {padding:12px;background:#fff2df;border-radius:8px}
a {color:#922052}
</style>
</head>
<body><main>
<a href="cap_fechamentos.asp">Voltar aos fechamentos</a>
<p class="selo">Ambiente de teste &middot; somente leitura</p>
<h1>Fechamento teste</h1>
<p>Consulte o per&iacute;odo na c&oacute;pia isolada do banco.</p>
<form method="get">
<label for="periodo">C&oacute;digo do per&iacute;odo</label>
<input id="periodo" name="periodo" inputmode="numeric" pattern="[1-9][0-9]{0,8}" value="<%=Server.HTMLEncode(ftPeriodo)%>" required>
<button type="submit">Conferir per&iacute;odo</button>
</form>
<% If ftErro <> "" Then %>
<p class="aviso" role="alert"><%=ftErro%></p>
<% Else %>
<h2>Pagamento de <%=Right(ftData,2)%>/<%=Mid(ftData,6,2)%>/<%=Left(ftData,4)%></h2>
<dl>
<div><dt>Pedidos eleg&iacute;veis pela regra do processador</dt><dd><%=ftElegiveis%></dd></div>
<div><dt>Pedidos sem confer&ecirc;ncia da diretoria</dt><dd><%=ftDiretoria%></dd></div>
<div><dt>Pedidos j&aacute; gravados no fechamento</dt><dd><%=ftPedidos%></dd></div>
<div><dt>Comiss&atilde;o j&aacute; gravada em pedidos</dt><dd><%=FormatCurrency(ftComissao)%></dd></div>
<div><dt>Boletos eleg&iacute;veis</dt><dd><%=ftBoletosElegiveis%></dd></div>
<div><dt>Boletos j&aacute; gravados</dt><dd><%=ftBoletos%></dd></div>
</dl>
<% If CLng(ftDiretoria) > 0 And ftAutorizacao <> "S" Then %>
<p class="aviso">H&aacute; pedidos sem confer&ecirc;ncia e o per&iacute;odo n&atilde;o tem autoriza&ccedil;&atilde;o expl&iacute;cita (S).</p>
<% End If %>
<% If ftProcessado <> "" Then %><p>Processamento registrado: <%=Server.HTMLEncode(ftProcessado)%>.</p><% End If %>
<% End If %>
</main></body></html>
