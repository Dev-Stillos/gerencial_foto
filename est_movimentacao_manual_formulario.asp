<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 1
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
  Dim titulo_pagina
  titulo_pagina = "Produtos"
%>
<%
Dim calendario_necessario
calendario_necessario = "N&AtildeO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<!-- #include file="includes/estoque_sub_menu.asp" -->
<%

'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 10/10/2011 16:04:06

'Declara vari&aacuteveis para utiliza&ccedil&atildeo do sistema
Dim cod_movimentacao, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

redirecionar_incluir	= "est_movimentacao_manual.asp"
redirecionar_alterar	= "est_movimentacao_manual.asp"
redirecionar_excluir	= "est_movimentacao_manual.asp"
redirecionar_validacao	= "est_movimentacao_manual.asp"
msgAoExcluir			= "Tem certeza que deseja excluir esta movimenta&ccedil&atildeo?"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
'Declara vari&aacuteveis
Dim tipo, cod_produto, quantidade, descricao

If comando = Empty Then
comando = "incluir"
titulo_pagina = titulo_pagina&" / Inclus&atildeo"
End if

If comando = "incluir" and campo1 <> Empty Then
	'Incluir novo registro
	incluirRegistro(redirecionar_incluir)
	titulo_pagina = titulo_pagina&" / Inclus&atildeo"
Elseif comando = "alterar" and campo1 = Empty Then
	'Leitura inicial para poss&iacutevel altera&ccedil&atildeo
	cod_movimentacao = Request.QueryString("cod_movimentacao")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / Altera&ccedil&atildeo"
Elseif 	comando = "alterar" and campo1 <> Empty Then
	'Gravar altera&ccedil&atildeo
	cod_movimentacao = Request.Form("cod_movimentacao")
	alterarRegistro(redirecionar_alterar)
	titulo_pagina = titulo_pagina&" / Altera&ccedil&atildeo"
Elseif 	comando = "excluir" and campo1 = Empty Then 
	'Leitura inicial para poss&iacutevel exclus&atildeo
	cod_movimentacao = Request.QueryString("cod_movimentacao")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / Exclus&atildeo"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
	'Exclui registro
	cod_movimentacao = Request.Form("cod_movimentacao")
	excluirRegistro(redirecionar_excluir)
	titulo_pagina = titulo_pagina&" / Exclus&atildeo"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
	'Visualiza registro
	cod_movimentacao = Request.QueryString("cod_movimentacao")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / Visualiza&ccedil&atildeo"
End If


'Fun&ccedil&otildees

Sub lerFormulario()

	COD_MOVIMENTACAO = Request.Form("COD_MOVIMENTACAO")
	COD_PRODUTO = Request.form("COD_PRODUTO")
	quantidade = Request.form("quantidade")
	TIPO = Request.form("tipo")
	descricao = Request.form("descricao")

End sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM TB_ESTOQUE_MOVIMENTACOES WHERE cod_movimentacao = "&cod_movimentacao)

			quantidade = rs("quantidade")
			COD_PRODUTO = rs("COD_PRODUTO")
			TIPO = rs("tipo")
			descricao = rs("descricao")

		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("INSERT INTO TB_ESTOQUE_MOVIMENTACOES (tipo, quantidade, COD_PRODUTO, descricao) VALUES ("&sqltexto(tipo)&", "&sqlNumero(quantidade)&","&sqlNumero(COD_PRODUTO)&", "&sqltexto(descricao)&") ")

		'Encontrar chave criada
		Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_ESTOQUE_MOVIMENTACOES LIMIT 1")
			If not rs.Eof Then
				cod_movimentacao = rs("ULTIMO_ID")
			End if 
		rs.Close
		Set rs=Nothing

	FecharConexao()
	Response.Redirect(redirecionar&"?cod_movimentacao="&cod_movimentacao&"&codmsg=1")
End sub

Sub alterarRegistro(redirecionar)
	
	Response.Redirect(redirecionar&"?cod_movimentacao="&cod_movimentacao&"&codmsg=2")
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		if  estoqueVerificaManual(cod_movimentacao, true) then
			objConexao.Execute("DELETE FROM TB_ESTOQUE_MOVIMENTACOES WHERE cod_movimentacao = "&cod_movimentacao)
		end if
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_movimentacao="&cod_movimentacao&"&codmsg=3")
End sub

'----------------------------------------------------------------------------------------
%>

<h2><%=titulo_pagina%></h2><br />

<%If comando = "excluir" Then%>
	<div class='msgExcluir'><%=msgAoExcluir%></div>
<%End If%>

<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post'>
	<table width='90%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
		<%If comando <> "incluir" Then %><tr>
		<td width='27%'><strong>C&oacuted Movimenta&ccedil&atildeo </strong></td>
			<td width='73%'>
				<div align='left'>
				<% If comando <> "visualizar" and comando <> "excluir"  Then%>
					<span class='textoMaior'><strong><%=COD_MOVIMENTACAO%></strong></span><%Else%>
					<%
					If not isNull(COD_MOVIMENTACAO) Then
						response.write(COD_MOVIMENTACAO)
					End If
				End If %>
				</div>
			</td>
		</tr>
		<%end if%>
		<tr>
			<td width='27%'><strong>Tipo *</strong></td>
			<td width='73%'>
				<div align='left'>
					<%=padraoFormularioSelect("tipo", "", "inputTamanho1","","","SEM_SELECIONE", TIPO,"PADRAO:E#Entrada|S#Sa&iacuteda",comando, false, false, false) %>
				</div>
			</td>
		</tr>
		<tr>
			<td width='27%'><strong>Produto *</strong></td>
			<td width='73%'>
				<div align='left'>
					<%=padraoFormularioSelect("COD_PRODUTO", "", "inputTamanho2","","","Selecione", COD_PRODUTO,"SELECT COD_PRODUTO, NOME_PRODUTO FROM TB_PRODUTOS ORDER BY NOME_PRODUTO ",comando, false, false, false) %>
				</div>
			</td>
		</tr>
		<tr>
			<td width='27%'><strong>Quantidade *</strong></td>
			<td width='73%'>
				<div align='left'>
				<%If comando <> "visualizar" and comando <> "excluir"  Then%>
					<input name='QUANTIDADE' value="<%=QUANTIDADE%>" class='inputTamanho2 padrao_numero' type='text' maxlength='250' />
				<% Else 
					If not isNull(QUANTIDADE) Then
						response.write(QUANTIDADE)
					End If
				End If %>
				</div>
			</td>
		</tr>
		<tr>
			<td width='27%'><strong>Descri&ccedil&atildeo *</strong></td>
			<td width='73%'>
				<div align='left'>
				<%If comando <> "visualizar" and comando <> "excluir"  Then%>
					<input name='DESCRICAO' class='inputTamanho2' value="<%=DESCRICAO%>" type='text' />
				<%
				Else
					If not isNull(DESCRICAO) Then
						response.write(DESCRICAO)
					End If
				End If %>
				</div>
			</td>
		</tr>
	
		<%If comando <> "visualizar" Then%>
		<tr>
			<td colspan='2' align='center'><div align='left'><strong>* - Campos obrigat&oacuterios</strong><br /></div></td>
		</tr>
		<tr>
			<td colspan='2' align='center'><a href="javascript:salvar();"><img src="images/layout/btnGravar.jpg" /></a></td>
		</tr>
	<%End If%>
	</table>
<input type='hidden' value='1' name='campo1'  />
<input type='hidden' name='COD_MOVIMENTACAO' value="<%=COD_MOVIMENTACAO%>" />&aacute</form>
<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('TIPO#Tipo#2|COD_PRODUTO#Produto#2|QUANTIDADE#Quantidade#1|DESCRICAO#Descri&ccedil&atildeo#1',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>
<!-- #include file="files/includes/miolo-inferior.asp" -->