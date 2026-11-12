<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim cod_pendencia, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir
Dim COD_VENDEDOR, TIPO, VALOR, MOTIVO, CADASTRADO_POR, cod_fechamento_periodo
Dim pendingFormTitle, pendingFormReadOnly, pendingFormError, pendingFormPattern
Dim pendingFormState, pendingFormVendor, pendingFormPayment, pendingFormOriginalVendor
Dim pendingFormPeriods, pendingFormHasPeriods
redirecionar_incluir = "cap_fechamentos_pendencias.asp"
redirecionar_alterar = redirecionar_incluir
redirecionar_excluir = redirecionar_incluir
comando = Request.QueryString("comando")
campo1 = Request.Form("campo1")
Select Case comando
    Case "", "incluir"
        comando = "incluir"
        pendingFormTitle = "Cadastrar pend&ecirc;ncia de pagamento"
    Case "alterar"
        pendingFormTitle = "Editar pend&ecirc;ncia de pagamento"
    Case "excluir"
        pendingFormTitle = "Excluir pend&ecirc;ncia de pagamento"
    Case "visualizar"
        pendingFormTitle = "Detalhes da pend&ecirc;ncia"
    Case Else
        Response.Redirect redirecionar_incluir
End Select
pendingFormReadOnly = (comando = "visualizar" Or comando = "excluir")
Set pendingFormPattern = New RegExp
pendingFormPattern.Pattern = "^[1-9][0-9]{0,9}$"

If comando <> "incluir" Then
    If campo1 = "" Then
        cod_pendencia = Request.QueryString("cod_pendencia")
    Else
        cod_pendencia = Request.Form("cod_pendencia")
    End If
    If Not pendingFormPattern.Test(cod_pendencia) Then Response.Redirect redirecionar_incluir
    lerRegistro()
    pendingFormOriginalVendor = COD_VENDEDOR
    If comando <> "visualizar" And pendingFormState <> 0 Then
        comando = "visualizar"
        pendingFormReadOnly = True
        pendingFormTitle = "Detalhes da pend&ecirc;ncia"
        If pendingFormState = 1 Then
            pendingFormError = "O pagamento de refer&ecirc;ncia j&aacute; foi processado."
        Else
            pendingFormError = "N&atilde;o foi poss&iacute;vel confirmar a situa&ccedil;&atilde;o do pagamento de refer&ecirc;ncia."
        End If
    End If
End If

If campo1 <> "" Then
    If comando = "excluir" Then
        excluirRegistro(redirecionar_excluir)
    ElseIf Not pendingFormReadOnly Then
        lerFormulario()
        If Not pendingFormPattern.Test(COD_VENDEDOR) Then
            pendingFormError = "Selecione um vendedor."
        ElseIf TIPO <> "A" And TIPO <> "D" Then
            pendingFormError = "Selecione o tipo de ajuste."
        ElseIf Len(Trim(MOTIVO)) = 0 Or Len(MOTIVO) > 255 Then
            pendingFormError = "Informe o motivo em at&eacute; 255 caracteres."
        ElseIf Not pendingFormPattern.Test(cod_fechamento_periodo) Then
            pendingFormError = "Selecione um pagamento em aberto."
        Else
            AbrirConexao()
            Set rs2 = objConexao.Execute("SELECT cod_fechamento_periodo FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE data_processado_pagamento IS NULL AND cod_fechamento_periodo = " & cod_fechamento_periodo)
            If rs2.Eof Then pendingFormError = "O pagamento selecionado j&aacute; foi processado ou est&aacute; indispon&iacute;vel."
            rs2.Close
            Set rs2 = Nothing
            If Len(pendingFormError & "") = 0 Then
                Set rs2 = objConexao.Execute("SELECT cod_vendedor FROM TB_VENDEDORES WHERE ativado = 1 AND lote = 2 AND cod_vendedor = " & COD_VENDEDOR)
                If rs2.Eof And COD_VENDEDOR <> pendingFormOriginalVendor Then pendingFormError = "Selecione um vendedor dispon&iacute;vel."
                rs2.Close
                Set rs2 = Nothing
            End If
            FecharConexao()
        End If
        If Len(pendingFormError & "") = 0 Then
            If comando = "incluir" Then
                incluirRegistro(redirecionar_incluir)
            Else
                alterarRegistro(redirecionar_alterar)
            End If
        End If
    End If
End If
Set pendingFormPattern = Nothing
Sub lerFormulario()

COD_VENDEDOR = Request.Form("COD_VENDEDOR")
TIPO = Request.Form("TIPO")
VALOR = Request.Form("VALOR")
MOTIVO = Request.Form("MOTIVO")
CADASTRADO_POR = Session("cod_usuario")
cod_fechamento_periodo = Request.Form("cod_fechamento_periodo")

End sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT d.*, v.nome_vendedor, p.data_pagamento AS pagamento_periodo, " & _
        "CASE WHEN p.cod_fechamento_periodo IS NOT NULL THEN CASE WHEN p.data_processado_pagamento IS NULL THEN 0 ELSE 1 END " & _
        "WHEN d.vendedor_recebe_em IS NOT NULL AND corte.primeiro_pendente IS NOT NULL THEN CASE WHEN DATE(d.vendedor_recebe_em) <= DATE(corte.primeiro_pendente) THEN 1 ELSE 0 END ELSE 2 END AS situacao_pendencia " & _
        "FROM TB_FECHAMENTO_PENDENCIAS_PAGAMENTO d LEFT JOIN TB_VENDEDORES v ON v.cod_vendedor = d.cod_vendedor " & _
        "LEFT JOIN TB_FECHAMENTO_PAGAMENTO_PERIODOS p ON p.cod_fechamento_periodo = d.cod_fechamento_periodo " & _
        "CROSS JOIN (SELECT MIN(data_pagamento) AS primeiro_pendente FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE data_processado_pagamento IS NULL) corte WHERE d.cod_pendencia = " & cod_pendencia)

				If rs.Eof Then
                    rs.Close
                    Set rs = Nothing
                    FecharConexao()
                    Response.Redirect "cap_fechamentos_pendencias.asp"
                End If
                pendingFormState = CLng(rs("situacao_pendencia"))
                pendingFormVendor = rs("nome_vendedor") & ""
                pendingFormPayment = rs("pagamento_periodo")
                If Not IsDate(pendingFormPayment) Then pendingFormPayment = rs("vendedor_recebe_em")
                COD_VENDEDOR=trim(rs("COD_VENDEDOR") & "")
				TIPO=trim(rs("TIPO"))
				VALOR= ArrumaNumero(rs("VALOR"),"",2)
				MOTIVO=trim(rs("MOTIVO"))
				CADASTRADO_POR=trim(rs("CADASTRADO_POR"))
				cod_fechamento_periodo=trim(rs("cod_fechamento_periodo"))

		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)

	AbrirConexao()
		objConexao.Execute("INSERT INTO TB_FECHAMENTO_PENDENCIAS_PAGAMENTO (COD_VENDEDOR, TIPO, VALOR, MOTIVO, CADASTRADO_POR, cod_fechamento_periodo) VALUES ("&sqlNumero(COD_VENDEDOR)&", "&sqlTexto(TIPO)&", "&sqlNumero(VALOR)&", "&sqlTexto(MOTIVO)&", "&sqlTexto(CADASTRADO_POR)&", "&sqlTexto(cod_fechamento_periodo)&")")

		'Encontrar chave criada
		Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_FECHAMENTO_PENDENCIAS_PAGAMENTO LIMIT 1")
			If not rs.Eof Then
				cod_pendencia = rs("ULTIMO_ID")
			End if 
		rs.Close
		Set rs=Nothing


	FecharConexao()
	Response.Redirect(redirecionar&"?cod_pendencia="&cod_pendencia&"&codmsg=1")
End sub

Sub alterarRegistro(redirecionar)
	
	AbrirConexao()
		objConexao.Execute("UPDATE TB_FECHAMENTO_PENDENCIAS_PAGAMENTO SET COD_VENDEDOR = "&sqlNumero(COD_VENDEDOR)&", TIPO = "&sqlTexto(TIPO)&", VALOR = "&sqlNumero(VALOR)&", MOTIVO = "&sqlTexto(MOTIVO)&", cod_fechamento_periodo = "&sqlTexto(cod_fechamento_periodo)&" WHERE cod_pendencia = "&cod_pendencia)
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_pendencia="&cod_pendencia&"&codmsg=2")
End sub

Sub excluirRegistro(redirecionar)

	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_FECHAMENTO_PENDENCIAS_PAGAMENTO WHERE cod_pendencia = "&cod_pendencia)
	FecharConexao()
	
	Response.Redirect(redirecionar&"?cod_pendencia="&cod_pendencia&"&codmsg=3")
End sub
%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Pend&ecirc;ncia de pagamento | Portal Stillos</title>
<link rel="stylesheet" href="css/vendor/bootstrap.min.css">
<link rel="stylesheet" href="css/dashboard.css">
<!-- #include file="files/includes/scripts.asp" -->
<style>
.pending-form { height:auto; max-width:820px; }
.pending-form h2 { margin-bottom:14px; }
.pending-fields { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:14px; }
.pending-wide { grid-column:1/-1; }
.pending-fields label { display:block; font-weight:650; margin-bottom:4px; }
.pending-fields textarea { min-height:105px; }
.pending-count { display:block; text-align:right; color:#53566b; }
.pending-actions { display:flex; flex-wrap:wrap; gap:8px; margin-top:16px; }
.pending-reference { padding:12px; background:#f5f2f8; border-radius:8px; }
.pending-details { display:grid; gap:10px; margin:0; }
.pending-details dd { margin:2px 0 0; overflow-wrap:anywhere; white-space:pre-wrap; }
@media(max-width:600px) { .pending-fields { grid-template-columns:1fr; } }
</style>
</head>
<body>
<a class="skip-link" href="#conteudo">Ir para o conte&uacute;do</a>
<div class="dashboard-layout">
  <aside class="dashboard-sidebar" id="menuPrincipal">
    <a class="dashboard-brand" href="adm_index.asp"><img src="images/logoparafundopreto.png" alt="Stillos" width="180" /><span>GERENCIAL FOTOS</span></a>
    <div class="nav-caption">&Aacute;REA DE TRABALHO</div>
    <!-- #include file="includes/dashboard_menu.asp" -->
  </aside>
  <div class="dashboard-workspace">
    <header class="dashboard-topbar">
      <button class="btn menu-toggle" id="ppvMenuToggle" type="button" aria-controls="menuPrincipal" aria-expanded="false">Menu</button>
      <div class="topbar-title">Portal Stillos <span>/ Pend&ecirc;ncias de pagamento</span></div>
      <div class="topbar-actions"><a class="user-name" href="adm_log_usuarios.asp"><%=Server.HTMLEncode(Session("NOME") & "")%></a><a href="manual_sistema.asp" target="_blank" rel="noopener">Manual</a><a class="btn btn-outline-secondary btn-sm" href="adm_sair.asp">Sair</a></div>
    </header>
    <main class="dashboard-main ppv-main" id="conteudo">



      <header class="dashboard-heading">
        <div><p class="eyebrow">FECHAMENTOS</p><h1><%=pendingFormTitle%></h1></div>
        <a class="btn btn-outline-secondary" href="cap_fechamentos_pendencias.asp">Voltar &agrave;s pend&ecirc;ncias</a>
      </header>
      <section class="dashboard-panel pending-form" aria-labelledby="pending-form-heading">
        <h2 id="pending-form-heading">Dados da pend&ecirc;ncia<%If comando <> "incluir" Then%> #<%=cod_pendencia%><%End If%></h2>
        <%If Len(pendingFormError & "") > 0 Then%><p class="alert alert-danger" role="alert"><%=pendingFormError%></p><%End If%>
        <%If comando = "excluir" Then%><p class="alert alert-warning">Confirme a exclus&atilde;o desta pend&ecirc;ncia.</p><%End If%>
        <form action="cap_fechamentos_pendencias_formulario.asp?comando=<%=comando%>" name="FORMULARIO_PADRAO" method="post">
          <input type="hidden" name="campo1" value="1">
          <input type="hidden" name="COD_PENDENCIA" value="<%=cod_pendencia%>">
          <%If pendingFormReadOnly Then%>
          <dl class="pending-details">
            <div><dt>Vendedor</dt><dd><%=Server.HTMLEncode(pendingFormVendor)%></dd></div>
            <div><dt>Tipo de ajuste</dt><dd><%If TIPO = "A" Then%>Acr&eacute;scimo<%ElseIf TIPO = "D" Then%>Decr&eacute;scimo<%Else%>&mdash;<%End If%></dd></div>
            <div><dt>Valor</dt><dd>R$ <%=Server.HTMLEncode(VALOR & "")%></dd></div>
            <div><dt>Motivo</dt><dd><%=Server.HTMLEncode(MOTIVO & "")%></dd></div>
            <div><dt>Pagamento</dt><dd><%If IsDate(pendingFormPayment) Then%><%=ArrumaData(pendingFormPayment)%><%Else%>&mdash;<%End If%></dd></div>
          </dl>
          <%Else%>
          <div class="pending-fields">
            <div class="pending-wide"><label for="pending-vendor">Vendedor *</label>
              <select id="pending-vendor" name="COD_VENDEDOR" class="form-select" required>
                <option value="">Selecione</option>
                <%AbrirConexao()
                Dim pendingVendorSql
                pendingVendorSql = "SELECT cod_vendedor, nome_vendedor FROM TB_VENDEDORES WHERE (ATIVADO = 1 AND lote = 2)"
                If Len(pendingFormOriginalVendor & "") > 0 Then pendingVendorSql = pendingVendorSql & " OR cod_vendedor = " & sqlNumero(pendingFormOriginalVendor)
                Set rs2 = objConexao.Execute(pendingVendorSql & " ORDER BY nome_vendedor ASC")
                While Not rs2.Eof%>
                  <option value="<%=rs2(0)%>"<%If CStr(rs2(0)) = CStr(COD_VENDEDOR & "") Then%> selected<%End If%>><%=Server.HTMLEncode(rs2(1) & "")%></option>
                <%rs2.MoveNext
                Wend
                rs2.Close
                Set rs2 = Nothing%>
              </select>
            </div>
            <div><label for="pending-type">Tipo de ajuste *</label>
              <select id="pending-type" name="TIPO" class="form-select" required>
                <option value="">Selecione</option>
                <option value="A"<%If TIPO = "A" Then%> selected<%End If%>>Acr&eacute;scimo (+)</option>
                <option value="D"<%If TIPO = "D" Then%> selected<%End If%>>Decr&eacute;scimo (-)</option>
              </select>
            </div>
            <div><label for="pending-value">Valor (R$)</label><input id="pending-value" name="VALOR" class="form-control padrao_valor" inputmode="decimal" maxlength="200" value="<%=Server.HTMLEncode(VALOR & "")%>"></div>
            <div class="pending-wide"><label for="pending-reason">Motivo *</label><textarea id="pending-reason" name="MOTIVO" class="form-control" maxlength="255" required aria-describedby="pending-count"><%=Server.HTMLEncode(MOTIVO & "")%></textarea><small id="pending-count" class="pending-count">At&eacute; 255 caracteres</small></div>
            <div class="pending-wide pending-reference"><label for="pending-payment">Agendar para pagamento *</label>
              <%Set rs2 = objConexao.Execute("SELECT cod_fechamento_periodo, data_pagamento FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE data_processado_pagamento IS NULL ORDER BY data_pagamento ASC")
              pendingFormHasPeriods = Not rs2.Eof%>
              <select id="pending-payment" name="cod_fechamento_periodo" class="form-select" aria-describedby="pending-payment-note" required<%If Not pendingFormHasPeriods Then%> disabled<%End If%>>
              <%If pendingFormHasPeriods Then
                  If comando = "alterar" Or Len(pendingFormError & "") > 0 Then%><option value="">Selecione</option><%End If
                  While Not rs2.Eof%>
                    <option value="<%=rs2(0)%>"<%If CStr(rs2(0)) = CStr(cod_fechamento_periodo & "") Then%> selected<%End If%>><%=ArrumaData(rs2(1))%></option>
                  <%rs2.MoveNext
                  Wend
                Else%><option value="">Nenhum pagamento em aberto</option><%End If
                rs2.Close
                Set rs2 = Nothing
                FecharConexao()%>
              </select>
              <p id="pending-payment-note" class="panel-description">A lista mostra os per&iacute;odos ainda n&atilde;o processados.</p>
            </div>
          </div>
          <%End If%>
          <div class="pending-actions">
            <%If comando = "excluir" Then%><button type="submit" class="btn btn-danger">Excluir pend&ecirc;ncia</button>
            <%ElseIf Not pendingFormReadOnly Then%><button type="submit" class="btn btn-primary"<%If Not pendingFormHasPeriods Then%> disabled<%End If%>>Salvar pend&ecirc;ncia</button><%End If%>
            <a class="btn btn-outline-secondary" href="cap_fechamentos_pendencias.asp"><%If comando = "visualizar" Then%>Voltar<%Else%>Cancelar<%End If%></a>
          </div>
        </form>
      </section>
    </main>
    <footer class="dashboard-footer">&copy; <%=Year(Date)%> Grupo Stillos <span>Sistema administrativo interno</span></footer>
  </div>
</div>
<!-- #include file="files/includes/scripts_rodape.asp" -->
<script>
(function () {
    var toggle = document.getElementById('ppvMenuToggle');
    var menu = document.getElementById('menuPrincipal');
    var previous = menu.querySelector('[aria-current="page"]');
    var current = menu.querySelector('a[href="cap_fechamentos.asp"]');
    if (previous) previous.removeAttribute('aria-current');
    if (current) current.setAttribute('aria-current', 'page');
    toggle.addEventListener('click', function () {
        var expanded = menu.classList.toggle('is-open');
        toggle.setAttribute('aria-expanded', String(expanded));
    });
    document.addEventListener('keydown', function (event) {
        if (event.key === 'Escape' && menu.classList.contains('is-open')) {
            menu.classList.remove('is-open');
            toggle.setAttribute('aria-expanded', 'false');
            toggle.focus();
        }
    });
})();


(function () {
    var reason = document.getElementById('pending-reason');
    if (!reason) return;
    var count = document.getElementById('pending-count');
    function updateCount() {
        count.textContent = reason.value.length + ' / 255 caracteres';
    }
    reason.addEventListener('input', updateCount);
    updateCount();
})();
</script>
</body>
</html>
