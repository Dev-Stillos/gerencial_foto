<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 1
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
  Dim titulo_pagina
  titulo_pagina = "CAP - Formulário de Períodos de Vendas"
%>			
<%
Dim calendario_necessario
calendario_necessario = "NÃO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%

'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 11/1/2012 11:19:30

'Declara vari�veis para utiliza��o do sistema
Dim cod_tabela_comissao_vendedores, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

redirecionar_incluir	= "cap_fechamentos_tabelas_comissao_vendedores_formulario_consulta_produtos_formulario.asp"
redirecionar_alterar	= "cap_fechamentos_tabelas_comissao_vendedores.asp"
redirecionar_excluir	= "cap_fechamentos_tabelas_comissao_vendedores.asp"
redirecionar_validacao	= "cap_fechamentos_tabelas_comissao_vendedores.asp"
msgAoExcluir			= "Tem certeza que deseja excluir esta Tabela de Comissionamento?<br /><br />Se houver vendedores com esta tabela, a mesma NÃO ser� exclu�da"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
'Declara vari�veis
Dim descricao, cod_usuario, modificado, ANO_FIM,ANO_INICIO, tipo

If comando = Empty Then
comando = "incluir"
titulo_pagina = titulo_pagina&" / Inclusão"
End if


If comando = "incluir" and campo1 <> Empty Then
'Incluir novo registro
incluirRegistro(redirecionar_incluir)
titulo_pagina = titulo_pagina&" / Inclusão"
Elseif comando = "alterar" and campo1 = Empty Then
'Leitura inicial para poss�vel Alteração
cod_tabela_comissao_vendedores = Request.QueryString("cod_tabela_comissao_vendedores")
lerRegistro()
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "alterar" and campo1 <> Empty Then
'Gravar Alteração
cod_tabela_comissao_vendedores = Request.Form("cod_tabela_comissao_vendedores")
alterarRegistro(redirecionar_alterar)
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 
'Leitura inicial para poss�vel Exclusão
cod_tabela_comissao_vendedores = Request.QueryString("cod_tabela_comissao_vendedores")
lerRegistro()
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
'Exclui registro
cod_tabela_comissao_vendedores = Request.Form("cod_tabela_comissao_vendedores")
excluirRegistro(redirecionar_excluir)
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
'Visualiza registro
cod_tabela_comissao_vendedores = Request.QueryString("cod_tabela_comissao_vendedores")
lerRegistro()
titulo_pagina = titulo_pagina&" / Visualização"
End If


'Fun��es

Sub lerFormulario()

cod_tabela_comissao_vendedores = Request.Form("cod_tabela_comissao_vendedores")
descricao = Request.Form("descricao")
cod_usuario = Session("cod_usuario")
modificado = now()
ANO_FIM=Request.Form("ANO_FIM")
ANO_INICIO=Request.Form("ANO_INICIO")
tipo=Request.Form("tipo")

End sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES WHERE cod_tabela_comissao_vendedores = "&cod_tabela_comissao_vendedores)

				descricao=trim(rs("descricao"))
				cod_usuario=trim(rs("cod_usuario"))
				modificado=trim(rs("modificado"))
				ANO_FIM=trim(rs("ANO_FIM"))
				ANO_INICIO=trim(rs("ANO_INICIO"))
				tipo=trim(rs("tipo"))&""


		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)
	lerFormulario()	

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
	lerFormulario()
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

<h2><%=titulo_pagina%></h2><br />

<%If comando = "excluir" Then%>
<div class='msgExcluir'><%=msgAoExcluir%></div>
<%End If%>

<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post'>
<%if comando = "incluir" Then%>
<div class="msgGrande">Após o cadastro você será direcionado para cadastrar as comissões dos produto desta tabela.</div>
<%End If%>

<table width='90%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
<%If comando <> "incluir" Then %><tr>
<td width='45%'><strong>Cód. Tabela de Comissionamento de Vendedor </strong></td>
<td width='55%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<span class='textoMaior'><strong><%=COD_TABELA_COMISSAO_VENDEDORES%></strong></span><%Else%>
<%If not isNull(COD_TABELA_COMISSAO_VENDEDORES) Then%>
<%=COD_TABELA_COMISSAO_VENDEDORES%>
<%End If%>
<%End If%>
</div></td>
</tr>
<%End If%>

<tr>
	<td width='45%'><strong>Nome da Tabela *</strong></td>
	<td width='55%'>
		<div align='left'>
			<%If comando <> "visualizar" and comando <> "excluir"  Then%>
				<input name='DESCRICAO' value="<%=DESCRICAO%>" class='inputTamanho3' type='text' maxlength='255' />
			<%Else
				If not isNull(DESCRICAO) Then
					response.write(DESCRICAO)
				End If
			End If%>
		</div>
	</td>
</tr>


<tr>
	<td width='45%'><strong>Tipo *</strong></td>
	<td width='55%'>
		<div align='left'>

			<%=padraoFormularioSelect("TIPO","tipo_id","inputTamanho1","","", "Selecione", tipo,"PADRAO:R#Repasse|P#Padrão", comando, false, false, false)%>

		
		</div>
	</td>
</tr>

<tr class='campos_repasse <% if tipo <> "R" or tipo = empty then %> hidden <% end if %>'>
	<td width='45%'><strong>Ano Início *</strong></td>
	<td width='55%'>
		<div align='left'>
			<%If comando <> "visualizar" and comando <> "excluir"  Then%>
				<input name='ANO_INICIO' id='ANO_INICIO_id' value="<%=ANO_INICIO%>" class='inputTamanho1' type='text' maxlength='4' />
			<%Else
				If not isNull(ANO_INICIO) Then
					response.write(ANO_INICIO)
				End If
			End If%>
		</div>
	</td>
</tr>

<tr class='campos_repasse <% if tipo <> "R" or tipo = empty  then %> hidden <% end if %>'>
	<td width='45%'><strong>Ano Fim *</strong></td>
	<td width='55%'>
		<div align='left'>
			<%If comando <> "visualizar" and comando <> "excluir"  Then%>
				<input name='ANO_FIM' id='ANO_FIM_id' value="<%=ANO_FIM%>" class='inputTamanho1' type='text' maxlength='4' />
			<%Else
				If not isNull(ANO_FIM) Then
					response.write(ANO_FIM) 
				End If
			End If%>
		</div>
	</td>
</tr>




<%If comando <> "visualizar" Then%>
<tr>
<td colspan='2' align='center'><div align='left'><strong>* - Campos obrigatórios</strong><br /></div></td>
</tr>
<tr>
  <td colspan='2' align='center'><a href="javascript:salvar();"><img border="0" src="images/layout/btnGravar.jpg" /></a></td>
</tr>
<%End If%>
</table>


<script>
	$(function(){
		$("#tipo_id").change(function(){
			if ($(this).val() == 'R') 
			{
				$(".campos_repasse").removeClass("hidden");
			}
			else
			{
				$(".campos_repasse").addClass("hidden");
				$("#ANO_INICIO_id").val("");
				$("#ANO_FIM_id").val("");
			}
		})
	});
</script>


<input type='hidden' value='1' name='campo1'  />
<input type='hidden' name='COD_TABELA_COMISSAO_VENDEDORES' value="<%=COD_TABELA_COMISSAO_VENDEDORES%>" /></form>
<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('DESCRICAO#Nome da Tabela#1',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->