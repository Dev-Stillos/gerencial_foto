<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/ppv_responsaveis_padrao.asp" -->
<%
Response.CodePage = 65001
Response.Charset = "UTF-8"
Response.CacheControl = "no-cache"
If Not PPVPodeConfigurarResponsaveis() Then
    Response.Status = "403 Forbidden"
    Response.Write "Acesso restrito."
    Response.End
End If

Function H(valor)
    H = Server.HTMLEncode(valor & "")
End Function

Function LotePPV(tipo, transferir, anterior)
    Dim regra
    regra = " FROM TB_ORDENS_SERVICO_ETAPAS e " & _
        "INNER JOIN TB_ORDENS_SERVICO o ON o.COD_ORDEM_SERVICO = e.COD_ORDEM_SERVICO " & _
        "INNER JOIN TB_PEDIDOS p ON p.COD_PEDIDO = o.COD_PEDIDO " & _
        "WHERE o.STATUS_ORDEM_SERVICO = 1 AND e.DATA_TERMINO IS NULL AND UPPER(TRIM(p.PRE_VENDA)) = '" & tipo & "'"
    If transferir Then
        regra = regra & " AND e.COD_USUARIO_RESPONSAVEL = " & CLng(anterior)
    Else
        regra = regra & " AND o.COD_ORDEM_SERVICO_ETAPA_ATUAL = e.COD_ORDEM_SERVICO_ETAPA " & _
            "AND LOWER(TRIM(e.DESCRICAO)) = 'triagem' AND COALESCE(e.COD_USUARIO_RESPONSAVEL,0) = 0"
    End If
    LotePPV = regra
End Function

Function QuantidadePPV(regra)
    Dim linhas
    Set linhas = objConexao.Execute("SELECT CAST(COUNT(*) AS CHAR) AS qtd" & regra)
    QuantidadePPV = CLng(linhas.Fields("qtd").Value)
    linhas.Close
End Function

Function GravarPPVLote(regra, usuario, ByRef total)
    Dim posWhere, sql, afetados
    posWhere = InStr(regra, "WHERE ")
    sql = "UPDATE " & Trim(Mid(regra, 7, posWhere - 7)) & _
        " SET e.COD_USUARIO_RESPONSAVEL = " & CLng(usuario) & " " & Mid(regra, posWhere)
    GravarPPVLote = ExecutarPPV(sql, afetados)
    If GravarPPVLote Then total = total + CLng(afetados)
End Function

Function ExecutarPPV(sql, ByRef afetados)
    On Error Resume Next
    Err.Clear
    objConexao.Execute sql, afetados, 128
    ExecutarPPV = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Dim ppvEmModal, ppvDestino
ppvEmModal = (Request.QueryString("modal") = "1")
ppvDestino = "ppv_responsaveis_padrao.asp"
If ppvEmModal Then ppvDestino = ppvDestino & "?modal=1"

Dim config, usuarios, codigo, nome, atualPre, atualPos, novoPre, novoPos, acao, erro, mensagem
Dim mostrarModal, lote, transferencia, transfPre, transfPos, contPre, contPos, contTransfPre, contTransfPos
Dim lockPPV, linhasLock, totalAlterado, gravou, afetados, token
AbrirConexao()
If Len(Session("PPVConfigToken") & "") = 0 Then
    Set linhasLock = objConexao.Execute("SELECT UUID() AS token")
    token = CStr(linhasLock("token"))
    linhasLock.Close
    Session("PPVConfigToken") = token
End If
acao = ""
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Request.Form("token") <> Session("PPVConfigToken") Then
        FecharConexao()
        Response.Status = "403 Forbidden"
        Response.End
    End If
    acao = Request.Form("acao")
End If

lockPPV = False
If acao = "salvar" Then
    Set linhasLock = objConexao.Execute("SELECT CAST(GET_LOCK('stillos3_ppv_responsaveis_padrao', 5) AS CHAR) AS obteve")
    If Not linhasLock.EOF Then lockPPV = (Trim(linhasLock.Fields("obteve").Value & "") = "1")
    linhasLock.Close
    If Not lockPPV Then erro = "Outra altera&ccedil;&atilde;o est&aacute; em andamento. Revise novamente."
End If

atualPre = "0"
atualPos = "0"
Set config = objConexao.Execute("SELECT COD_USUARIO_PRE, COD_USUARIO_POS FROM TB_PPV_RESPONSAVEIS_PADRAO WHERE ID = 1")
If Not config.EOF Then
    atualPre = CStr(config("COD_USUARIO_PRE"))
    atualPos = CStr(config("COD_USUARIO_POS"))
End If
config.Close
novoPre = atualPre
novoPos = atualPos

If acao = "revisar" Or acao = "salvar" Then
    novoPre = Trim(Request.Form("pre"))
    novoPos = Trim(Request.Form("pos"))
    If Not PPVUsuarioElegivel(novoPre) Or Not PPVUsuarioElegivel(novoPos) Then
        erro = "Selecione um funcion&aacute;rio ativo para cada fluxo."
    Else
        novoPre = CStr(CLng(novoPre))
        novoPos = CStr(CLng(novoPos))
        transfPre = (atualPre <> "0" And atualPre <> novoPre)
        transfPos = (atualPos <> "0" And atualPos <> novoPos)
        If acao = "revisar" Then
            Session("PPVRevisao") = novoPre & ":" & novoPos & ":" & atualPre & ":" & atualPos
            contPre = QuantidadePPV(LotePPV("S", False, 0))
            contPos = QuantidadePPV(LotePPV("N", False, 0))
            contTransfPre = 0
            contTransfPos = 0
            If transfPre Then contTransfPre = QuantidadePPV(LotePPV("S", True, atualPre))
            If transfPos Then contTransfPos = QuantidadePPV(LotePPV("N", True, atualPos))
            mostrarModal = True
        ElseIf Len(erro) = 0 Then
            If Session("PPVRevisao") <> novoPre & ":" & novoPos & ":" & atualPre & ":" & atualPos Then
                erro = "A configura&ccedil;&atilde;o mudou ou a revis&atilde;o expirou. Revise novamente."
            Else
                lote = (Request.Form("lote") = "1")
                transferencia = (Request.Form("transferir") = "1")
                totalAlterado = 0
                gravou = True
                If transferencia And transfPre Then gravou = GravarPPVLote(LotePPV("S", True, atualPre), novoPre, totalAlterado)
                If gravou And transferencia And transfPos Then gravou = GravarPPVLote(LotePPV("N", True, atualPos), novoPos, totalAlterado)
                If gravou And lote Then gravou = GravarPPVLote(LotePPV("S", False, 0), novoPre, totalAlterado)
                If gravou And lote Then gravou = GravarPPVLote(LotePPV("N", False, 0), novoPos, totalAlterado)
                If gravou Then
                    gravou = ExecutarPPV("INSERT INTO TB_PPV_RESPONSAVEIS_PADRAO (ID,COD_USUARIO_PRE,COD_USUARIO_POS,ALTERADO_POR,ALTERADO_EM) VALUES (1," & CLng(novoPre) & "," & CLng(novoPos) & "," & sqlNumero(Session("cod_usuario")) & ",NOW()) ON DUPLICATE KEY UPDATE COD_USUARIO_PRE=VALUES(COD_USUARIO_PRE),COD_USUARIO_POS=VALUES(COD_USUARIO_POS),ALTERADO_POR=VALUES(ALTERADO_POR),ALTERADO_EM=VALUES(ALTERADO_EM)", afetados)
                End If
                Session("PPVRevisao") = ""
                If gravou Then
                    mensagem = "Respons&aacute;veis salvos. Etapas atribu&iacute;das ou transferidas: " & totalAlterado & "."
                Else
                    erro = "A grava&ccedil;&atilde;o foi interrompida. Etapas confirmadas nesta tentativa: " & totalAlterado & ". Revise as quantidades e confirme novamente para processar o restante."
                End If
            End If
        End If
    End If
End If
If lockPPV Then
    Set linhasLock = objConexao.Execute("SELECT RELEASE_LOCK('stillos3_ppv_responsaveis_padrao')")
    linhasLock.Close
End If

Dim opcoesPre, opcoesPos, nomePre, nomePos, selecionadoPre, selecionadoPos
Set usuarios = objConexao.Execute("SELECT COD_USUARIO,NOME FROM TB_ADM_USUARIOS WHERE BLOQUEADO=2 AND COD_USUARIO > 1 ORDER BY NOME, COD_USUARIO")
Do While Not usuarios.EOF
    codigo = CStr(usuarios("COD_USUARIO"))
    nome = H(usuarios("NOME"))
    selecionadoPre = ""
    selecionadoPos = ""
    If codigo = novoPre Then
        selecionadoPre = " selected"
        nomePre = nome
    End If
    If codigo = novoPos Then
        selecionadoPos = " selected"
        nomePos = nome
    End If
    opcoesPre = opcoesPre & "<option value=""" & codigo & """" & selecionadoPre & ">" & nome & "</option>"
    opcoesPos = opcoesPos & "<option value=""" & codigo & """" & selecionadoPos & ">" & nome & "</option>"
    usuarios.MoveNext
Loop
usuarios.Close
FecharConexao()
%>
<!doctype html>
<html lang="pt-BR"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Respons&aacute;veis padr&atilde;o | Portal Stillos</title>
<link rel="stylesheet" href="css/vendor/bootstrap.min.css"><link rel="stylesheet" href="css/dashboard.css">
<style>
body { background:#f5f4f8; }
main { width:min(760px,calc(100% - 32px)); margin:32px auto; }
.ppv-config,dialog { background:#fff; color:#282638; border:1px solid #ddd6e4; border-radius:14px; padding:clamp(18px,4vw,32px); }
dialog { width:min(600px,calc(100% - 32px)); }
dialog::backdrop { background:#202033a0; }
.ppv-fields { display:grid; gap:16px; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); margin:20px 0; }
.ppv-fields label { display:grid; gap:6px; font-weight:600; }
.ppv-choice { display:flex; align-items:baseline; gap:8px; margin:14px 0; }
.ppv-actions { display:flex; gap:8px; flex-wrap:wrap; margin-top:20px; }
body.ppv-embutido { background:#fff; }
.ppv-embutido main { width:100%; margin:0; }
.ppv-embutido .ppv-config { padding:8px 4px; border:0; }
</style></head><body<%If ppvEmModal Then%> class="ppv-embutido"<%End If%>>
<main<%If totalAlterado > 0 Then%> data-ppv-fila-alterada<%End If%>>
<%If Not ppvEmModal Then%><p><a href="ppv_ordens_servico_triagem_inicial.asp">Voltar &agrave; Triagem Inicial</a></p><%End If%>
<section class="ppv-config"<%If ppvEmModal And mostrarModal Then%> hidden<%End If%>>
<%If Not ppvEmModal Then%><h1>Respons&aacute;veis padr&atilde;o</h1><%End If%>
<p>Escolha quem receber&aacute; as novas OS de cada tipo de venda.</p>
<%If Len(erro) > 0 Then%><p class="alert alert-warning" role="alert"><%=erro%></p><%End If%>
<%If Len(mensagem) > 0 Then%><p class="alert alert-success" role="status"><%=mensagem%></p><%End If%>
<form method="post" action="<%=H(ppvDestino)%>">
<input type="hidden" name="token" value="<%=H(Session("PPVConfigToken"))%>">
<div class="ppv-fields">
<label>Pr&eacute;-venda<select name="pre" class="form-select" required><option value="">Selecione</option><%=opcoesPre%></select></label>
<label>P&oacute;s-venda<select name="pos" class="form-select" required><option value="">Selecione</option><%=opcoesPos%></select></label>
</div>
<button class="btn btn-primary" name="acao" value="revisar">Revisar altera&ccedil;&otilde;es</button>
</form>
</section>
<%If mostrarModal Then%>
<%If ppvEmModal Then%>
<section class="ppv-config" aria-labelledby="ppv-confirmar-titulo">
<%Else%>
<dialog id="ppv-confirmar" aria-labelledby="ppv-confirmar-titulo">
<%End If%>
<h2 id="ppv-confirmar-titulo">Confirmar respons&aacute;veis padr&atilde;o</h2>
<p>As novas OS eleg&iacute;veis receber&atilde;o os respons&aacute;veis abaixo.</p>
<ul><li>Pr&eacute;-venda: <strong><%=nomePre%></strong></li><li>P&oacute;s-venda: <strong><%=nomePos%></strong></li></ul>
<form method="post" action="<%=H(ppvDestino)%>">
<input type="hidden" name="token" value="<%=H(Session("PPVConfigToken"))%>">
<input type="hidden" name="pre" value="<%=H(novoPre)%>"><input type="hidden" name="pos" value="<%=H(novoPos)%>">
<label class="ppv-choice"><input type="checkbox" name="lote" value="1"><span>Atribuir OS sem respons&aacute;vel na Triagem Inicial (<%=contPre%> Pr&eacute;-venda e <%=contPos%> P&oacute;s-venda). Elas continuar&atilde;o com a triagem pendente.</span></label>
<%If transfPre Or transfPos Then%>
<label class="ppv-choice"><input type="checkbox" name="transferir" value="1"><span>Transferir etapas abertas do respons&aacute;vel anterior (<%=contTransfPre%> etapas de Pr&eacute;-venda e <%=contTransfPos%> de P&oacute;s-venda).</span></label>
<%End If%>
<div class="ppv-actions">
<button class="btn btn-primary" name="acao" value="salvar">Salvar respons&aacute;veis</button>
<%If ppvEmModal Then%>
<button class="btn btn-outline-secondary" type="button" onclick="this.closest('section').hidden=true; const selecao=document.querySelector('main > section[hidden]'); selecao.hidden=false; selecao.querySelector('select').focus()">Voltar &agrave; sele&ccedil;&atilde;o</button>
<%Else%>
<button class="btn btn-outline-secondary" type="button" id="ppv-cancelar">Cancelar</button>
<%End If%>
</div>
</form>
<%If ppvEmModal Then%>
</section>
<script>document.querySelector('[name="acao"][value="salvar"]').focus();</script>
<%Else%>
</dialog>
<script>
const ppvConfirmar = document.getElementById('ppv-confirmar');
if (typeof ppvConfirmar.showModal === 'function') ppvConfirmar.showModal();
else ppvConfirmar.setAttribute('open', '');
document.getElementById('ppv-cancelar').onclick = function () {
    if (typeof ppvConfirmar.close === 'function') ppvConfirmar.close();
    else ppvConfirmar.removeAttribute('open');
};
</script>
<%End If%>
<%End If%>
</main></body></html>
