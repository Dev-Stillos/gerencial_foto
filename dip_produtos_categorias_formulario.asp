<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 1
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
  Dim titulo_pagina
  titulo_pagina = "Categorias de Produtos"
%>
<%
Dim calendario_necessario
calendario_necessario = "NÃO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%

'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 10/10/2011 16:44:00

'Declara variáveis para utilização do sistema
Dim cod_categoria, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

redirecionar_incluir	= "dip_produtos_categorias.asp"
redirecionar_alterar	= "dip_produtos_categorias.asp"
redirecionar_excluir	= "dip_produtos_categorias.asp"
redirecionar_validacao	= "dip_produtos_categorias.asp"
msgAoExcluir			= "Tem certeza que deseja excluir esta Categoria de Produto?<br /><br />Você poderá inativa-la ao invés de exclui-la. Se houver Produtos com esta categoria, a mesma não será excluída"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
'Declara variáveis
Dim NOME_CATEGORIA

If comando = Empty Then
comando = "incluir"
titulo_pagina = titulo_pagina&" / Inclusão"
End if

If comando = "incluir" and campo1 <> Empty Then
'Incluir novo registro
incluirRegistro(redirecionar_incluir)
titulo_pagina = titulo_pagina&" / Inclusão"
Elseif comando = "alterar" and campo1 = Empty Then
'Leitura inicial para possível alteração
cod_categoria = Request.QueryString("cod_categoria")
lerRegistro()
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "alterar" and campo1 <> Empty Then
'Gravar alteração
cod_categoria = Request.Form("cod_categoria")
alterarRegistro(redirecionar_alterar)
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 
'Leitura inicial para possível exclusão
cod_categoria = Request.QueryString("cod_categoria")
lerRegistro()
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
'Exclui registro
cod_categoria = Request.Form("cod_categoria")
excluirRegistro(redirecionar_excluir)
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
'Visualiza registro
cod_categoria = Request.QueryString("cod_categoria")
lerRegistro()
titulo_pagina = titulo_pagina&" / Visualização"
End If


'Funções

Sub lerFormulario()

COD_CATEGORIA = Request.Form("COD_CATEGORIA")
NOME_CATEGORIA = Request.Form("NOME_CATEGORIA")

End sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM TB_PRODUTOS_CATEGORIAS WHERE cod_categoria = "&cod_categoria)

				NOME_CATEGORIA=trim(rs("NOME_CATEGORIA"))

		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)
	lerFormulario()

		Dim chavesUnicasDuplicadas,camposUnicos
		call chaveUnica ("TB_PRODUTOS_CATEGORIAS","NOME_CATEGORIA",NOME_CATEGORIA,redirecionar&"?cod_categoria="&cod_categoria&"&codmsg=4&sucesso=2&chaveUnica=",comando,"cod_categoria",cod_categoria)

	AbrirConexao()
		objConexao.Execute("INSERT INTO TB_PRODUTOS_CATEGORIAS (COD_CATEGORIA, NOME_CATEGORIA) VALUES ("&sqlTexto(COD_CATEGORIA)&", "&sqlTexto(NOME_CATEGORIA)&")")

		'Encontrar chave criada
		Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_PRODUTOS_CATEGORIAS LIMIT 1")
			If not rs.Eof Then
				cod_categoria = rs("ULTIMO_ID")
			End if 
		rs.Close
		Set rs=Nothing


	FecharConexao()
	Response.Redirect(redirecionar&"?cod_categoria="&cod_categoria&"&codmsg=1")
End sub

Sub alterarRegistro(redirecionar)
	lerFormulario()
		Dim chavesUnicasDuplicadas,camposUnicos
		call chaveUnica ("TB_PRODUTOS_CATEGORIAS","NOME_CATEGORIA",NOME_CATEGORIA,redirecionar&"?cod_categoria="&cod_categoria&"&codmsg=4&sucesso=2&chaveUnica=",comando,"cod_categoria",cod_categoria)


	AbrirConexao()
		objConexao.Execute("UPDATE TB_PRODUTOS_CATEGORIAS SET NOME_CATEGORIA = "&sqlTexto(NOME_CATEGORIA)&" WHERE cod_categoria = "&cod_categoria)
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_categoria="&cod_categoria&"&codmsg=2")
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()
		
		call verificaAntesExcluir(cod_categoria,"TB_PRODUTOS","cod_categoria",redirecionar_excluir&"?sucesso=2&codmsg=5")

	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_PRODUTOS_CATEGORIAS WHERE cod_categoria = "&cod_categoria)
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_categoria="&cod_categoria&"&codmsg=3")
End sub


'----------------------------------------------------------------------------------------

%>

<h2><%=titulo_pagina%></h2><br />

<%If comando = "excluir" Then%>
<script type="text/javascript">
setTimeout("parent.submenu.trocarPorGravarporExclusao()",1)
</script>
<div class='msgExcluir'><%=msgAoExcluir%></div>
<%Elseif comando = "visualizar" Then%>
<script type="text/javascript">
setTimeout("parent.submenu.comandoVisualizar()",1)
</script>
<%End If%>

<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post'>
<table width='90%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
<%If comando <> "incluir" Then %><tr>
<td width='27%'><strong>Cod. Categoria </strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
	<strong><%=COD_CATEGORIA%></strong>
<%Else%>
<%If not isNull(COD_CATEGORIA) Then%>
<%=COD_CATEGORIA%>
<%End If%>
<%End If%>
</div></td>
</tr>
<%End If%><tr>
<td width='27%'><strong>Nome da Categoria *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='NOME_CATEGORIA' value="<%=NOME_CATEGORIA%>" class='inputTamanho2' type='text' maxlength='250' />
<%Else%>
<%If not isNull(NOME_CATEGORIA) Then%>
<%=NOME_CATEGORIA%>
<%End If%>
<%End If%>
</div></td>
</tr>
<%If comando <> "visualizar" Then%>

<tr>
<td colspan='2' align='center'><div align='left'><strong>* - Campos obrigatórios</strong><br /></div></td>
</tr>
<tr>
  <td colspan='2' align='center'><a href="javascript:salvar();"><img src="images/layout/btnGravar.jpg" /></a></td>
</tr>
<%End If%>
</table>
<input type='hidden' value='1' name='campo1'  />
<input type='hidden' name='COD_CATEGORIA' value="<%=COD_CATEGORIA%>" /> </form>
<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('NOME_CATEGORIA#Nome da Categoria#1',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->
