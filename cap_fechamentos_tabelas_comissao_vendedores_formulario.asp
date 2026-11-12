<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="files/includes/functions_local.asp" -->
<%

'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 11/1/2012 11:19:30

'Declara vari�veis para utiliza��o do sistema
Dim cod_tabela_comissao_vendedores, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

redirecionar_incluir	= "cap_fechamentos_tabelas_comissao_vendedores_formulario_consulta_produtos_formulario.asp"
redirecionar_alterar	= "cap_fechamentos_tabelas_comissao_vendedores.asp"
redirecionar_excluir	= "cap_fechamentos_tabelas_comissao_vendedores.asp"
redirecionar_validacao	= "cap_fechamentos_tabelas_comissao_vendedores.asp"

actionformulario = "cap_fechamentos_tabelas_comissao_vendedores_formulario.asp"

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
'Declara vari�veis
Dim descricao, cod_usuario, modificado, ANO_FIM,ANO_INICIO, tipo
Dim commissionTitle, commissionError, commissionReadOnly, commissionIdPattern
Select Case comando
    Case "", "incluir"
        comando = "incluir"
        commissionTitle = "Cadastrar tabela de comiss&otilde;es"
    Case "alterar"
        commissionTitle = "Editar tabela de comiss&otilde;es"
    Case "excluir"
        commissionTitle = "Excluir tabela de comiss&otilde;es"
    Case "visualizar"
        commissionTitle = "Detalhes da tabela de comiss&otilde;es"
    Case Else
        Response.Redirect "cap_fechamentos_tabelas_comissao_vendedores.asp"
End Select
commissionReadOnly = (comando = "visualizar" Or comando = "excluir")

If comando <> "incluir" Then
    If campo1 = Empty Then
        cod_tabela_comissao_vendedores = Request.QueryString("cod_tabela_comissao_vendedores")
    Else
        cod_tabela_comissao_vendedores = Request.Form("cod_tabela_comissao_vendedores")
    End If
    Set commissionIdPattern = New RegExp
    commissionIdPattern.Pattern = "^[1-9][0-9]{0,9}$"
    If Not commissionIdPattern.Test(cod_tabela_comissao_vendedores) Then
        Response.Redirect "cap_fechamentos_tabelas_comissao_vendedores.asp"
    End If
    Set commissionIdPattern = Nothing
End If


If campo1 <> Empty And Not commissionReadOnly Then
    lerFormulario()
    If Len(Trim(descricao & "")) = 0 Or Len(descricao & "") > 255 Then
        commissionError = "Informe um nome para a tabela, com at&eacute; 255 caracteres."
    ElseIf tipo <> "P" And tipo <> "R" Then
        commissionError = "Selecione o tipo da tabela."
    End If
    If Len(commissionError & "") = 0 Then
        If comando = "incluir" Then
            incluirRegistro(redirecionar_incluir)
        Else
            alterarRegistro(redirecionar_alterar)
        End If
    End If
ElseIf comando = "excluir" And campo1 <> Empty Then
    excluirRegistro(redirecionar_excluir)
ElseIf comando <> "incluir" Then
    lerRegistro()
End If

'Fun��es

Sub lerFormulario()

cod_tabela_comissao_vendedores = Request.Form("cod_tabela_comissao_vendedores")
descricao = Request.Form("descricao")
cod_usuario = Session("cod_usuario")
modificado = now()
ANO_FIM=Request.Form("ANO_FIM")
ANO_INICIO=Request.Form("ANO_INICIO")
tipo = UCase(Trim(Request.Form("tipo")))
If tipo = "P" Then
    ANO_INICIO = ""
    ANO_FIM = ""
End If

End sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES WHERE cod_tabela_comissao_vendedores = "&cod_tabela_comissao_vendedores)
        If rs.Eof Then
            rs.Close
            Set rs = Nothing
            FecharConexao()
            Response.Redirect "cap_fechamentos_tabelas_comissao_vendedores.asp"
        End If


				descricao=trim(rs("descricao"))
				cod_usuario=trim(rs("cod_usuario"))
				modificado=trim(rs("modificado"))
				ANO_FIM=trim(rs("ANO_FIM"))
				ANO_INICIO=trim(rs("ANO_INICIO"))
				tipo=UCase(Trim(rs("tipo") & ""))


		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)	

	Dim chavesUnicasDuplicadas,camposUnicos
	call chaveUnica ("TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES","descricao",descricao,redirecionar&"?cod_tabela_comissao_vendedores="&cod_tabela_comissao_vendedores&"&codmsg=4&sucesso=2&chaveUnica=",comando,"cod_tabela_comissao_vendedores",cod_tabela_comissao_vendedores)

	AbrirConexao()
		objConexao.Execute("INSERT INTO TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES (descricao, cod_usuario, modificado, ano_fim, ano_inicio, tipo) VALUES ("&sqlTexto(descricao)&", "&sqlTexto(cod_usuario)&", "&sqlData(modificado)&", "&sqlTexto(ano_fim)&",  "&sqlTexto(ano_inicio)&", "&sqlTexto(tipo)&")")

		'Encontrar chave criada
		Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES LIMIT 1")
			If not rs.Eof Then
				cod_tabela_comissao_vendedores = rs("ULTIMO_ID")
			End if 
		rs.Close
		Set rs=Nothing


	FecharConexao()
	Response.Redirect(redirecionar&"?cod_tabela_comissao_vendedores="&cod_tabela_comissao_vendedores&"&codmsg=101")
End sub

Sub alterarRegistro(redirecionar)
		Dim chavesUnicasDuplicadas,camposUnicos
		call chaveUnica ("TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES","descricao",descricao,redirecionar&"?cod_tabela_comissao_vendedores="&cod_tabela_comissao_vendedores&"&codmsg=4&sucesso=2&chaveUnica=",comando,"cod_tabela_comissao_vendedores",cod_tabela_comissao_vendedores)


	AbrirConexao()
		objConexao.Execute("UPDATE TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES SET descricao = "&sqlTexto(descricao)&", cod_usuario = "&sqlTexto(cod_usuario)&", modificado = "&sqlData(modificado)&" , ano_fim = "&sqlTexto(ano_fim)&", ano_inicio = "&sqlTexto(ANO_INICIO)&", tipo = "&sqlTexto(tipo)&" WHERE cod_tabela_comissao_vendedores = "&cod_tabela_comissao_vendedores)
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_tabela_comissao_vendedores="&cod_tabela_comissao_vendedores&"&codmsg=2")
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()

	call verificaAntesExcluir(cod_tabela_comissao_vendedores,"TB_VENDEDORES","cod_tabela_comissao_vendedores",redirecionar_excluir&"?sucesso=2&codmsg=5")

	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES WHERE cod_tabela_comissao_vendedores = "&cod_tabela_comissao_vendedores)
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_tabela_comissao_vendedores="&cod_tabela_comissao_vendedores&"&codmsg=3")
End sub


'----------------------------------------------------------------------------------------

%>


<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Tabela de comiss&otilde;es | Portal Stillos</title>
<link rel="stylesheet" href="css/vendor/bootstrap.min.css">
<link rel="stylesheet" href="css/dashboard.css">
<style>
.commission-form-card { height:auto; max-width:760px; }
.commission-form-card h2 { margin-bottom:12px; }
.commission-form-fields { display:grid; gap:14px; }
.commission-form-fields label { display:block; margin-bottom:4px; font-weight:650; }
.commission-form-code { display:inline-block; margin:0 0 12px; padding:4px 8px; border-radius:6px; background:#f1edf6; }
.commission-years { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:12px; }
.commission-years[hidden] { display:none; }
.commission-form-actions { display:flex; flex-wrap:wrap; gap:8px; margin-top:16px; }
.commission-details { display:grid; gap:8px; margin:0; }
.commission-details dd { margin:0; overflow-wrap:anywhere; }
@media(max-width:460px) { .commission-years { grid-template-columns:1fr; } }
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
      <div class="topbar-title">Portal Stillos <span>/ Tabelas de comissionamento</span></div>
      <div class="topbar-actions"><a class="user-name" href="adm_log_usuarios.asp"><%=Server.HTMLEncode(Session("NOME") & "")%></a><a href="manual_sistema.asp" target="_blank" rel="noopener">Manual</a><a class="btn btn-outline-secondary btn-sm" href="adm_sair.asp">Sair</a></div>
    </header>
    <main class="dashboard-main ppv-main" id="conteudo">


      <header class="dashboard-heading">
        <div><p class="eyebrow">FECHAMENTOS</p><h1><%=commissionTitle%></h1></div>
        <a class="btn btn-outline-secondary" href="cap_fechamentos_tabelas_comissao_vendedores.asp">Voltar &agrave;s tabelas</a>
      </header>
      <div class="search-panel"><strong>Busca no portal</strong>
        <!-- #include file="includes/dashboard_busca.asp" -->
      </div>
      <div class="session-notices">
        <!-- #include file="includes/dashboard_avisos.asp" -->
      </div>
      <section class="dashboard-panel commission-form-card" aria-labelledby="commission-form-heading">
        <h2 id="commission-form-heading">Dados da tabela</h2>
        <%If comando <> "incluir" Then%><p class="commission-form-code">C&oacute;digo <strong><%=Server.HTMLEncode(cod_tabela_comissao_vendedores & "")%></strong></p><%End If%>
        <%If Len(commissionError & "") > 0 Then%><p class="alert alert-danger" role="alert"><%=commissionError%></p><%End If%>
        <%If comando = "excluir" Then%><p class="alert alert-warning">Confirme a exclus&atilde;o desta tabela. Tabelas vinculadas a vendedores est&atilde;o protegidas contra exclus&atilde;o.</p><%End If%>
        <form action="<%=actionformulario%>?comando=<%=comando%>" name="FORMULARIO_PADRAO" method="post">
          <input type="hidden" name="campo1" value="1">
          <input type="hidden" name="COD_TABELA_COMISSAO_VENDEDORES" value="<%=Server.HTMLEncode(cod_tabela_comissao_vendedores & "")%>">
          <%If commissionReadOnly Then%>
          <dl class="commission-details">
            <div><dt>Nome da tabela</dt><dd><%=Server.HTMLEncode(descricao & "")%></dd></div>
            <div><dt>Tipo</dt><dd><%If tipo = "R" Then%>Repasse<%ElseIf tipo = "P" Then%>Padr&atilde;o<%Else%>N&atilde;o informado<%End If%></dd></div>
            <%If tipo = "R" Then%>
            <div><dt>Ano inicial</dt><dd><%=Server.HTMLEncode(ANO_INICIO & "")%></dd></div>
            <div><dt>Ano final</dt><dd><%=Server.HTMLEncode(ANO_FIM & "")%></dd></div>
            <%End If%>
          </dl>
          <%Else%>
          <div class="commission-form-fields">
            <div><label for="commission-name">Nome da tabela</label><input class="form-control" id="commission-name" name="DESCRICAO" value="<%=Server.HTMLEncode(descricao & "")%>" maxlength="255" required></div>
            <div><label for="tipo_id">Tipo</label><select class="form-select" id="tipo_id" name="TIPO" required>
              <option value="">Selecione</option>
              <option value="P"<%If tipo = "P" Then%> selected<%End If%>>Padr&atilde;o</option>
              <option value="R"<%If tipo = "R" Then%> selected<%End If%>>Repasse</option>
            </select></div>
            <div class="commission-years" id="commission-years"<%If tipo <> "R" Then%> hidden<%End If%>>
              <div><label for="ANO_INICIO_id">Ano inicial</label><input class="form-control" id="ANO_INICIO_id" name="ANO_INICIO" inputmode="numeric" maxlength="4" value="<%=Server.HTMLEncode(ANO_INICIO & "")%>"></div>
              <div><label for="ANO_FIM_id">Ano final</label><input class="form-control" id="ANO_FIM_id" name="ANO_FIM" inputmode="numeric" maxlength="4" value="<%=Server.HTMLEncode(ANO_FIM & "")%>"></div>
            </div>
          </div>
          <%End If%>
          <div class="commission-form-actions">
            <%If comando = "excluir" Then%><button type="submit" class="btn btn-danger">Excluir tabela</button>
            <%ElseIf Not commissionReadOnly Then%><button type="submit" class="btn btn-primary">Salvar tabela</button><%End If%>
            <a class="btn btn-outline-secondary" href="cap_fechamentos_tabelas_comissao_vendedores.asp"><%If comando = "visualizar" Then%>Voltar<%Else%>Cancelar<%End If%></a>
          </div>
        </form>
      </section>
    </main>
    <footer class="dashboard-footer">&copy; <%=Year(Date)%> Grupo Stillos <span>Sistema administrativo interno</span></footer>
  </div>
</div>
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
    var type = document.getElementById('tipo_id');
    if (!type) return;
    var years = document.getElementById('commission-years');
    function updateYears() {
        years.hidden = type.value !== 'R';
        years.querySelectorAll('input').forEach(function (input) {
            input.disabled = years.hidden;
        });
    }
    type.addEventListener('change', updateYears);
    updateYears();
})();
</script>
</body>
</html>
