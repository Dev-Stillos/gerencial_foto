<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 2
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
  Dim titulo_pagina
  titulo_pagina = "Produtos | Formulário de Tabela de Preços"
%>
<%
Dim calendario_necessario
calendario_necessario = "SIM"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%

'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 13/10/2011 11:48:10

'Declara variáveis para utilização do sistema
Dim COD_PRODUTO, nome_produto, cod_valor, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario


COD_PRODUTO = Request.QueryString("COD_PRODUTO")

AbrirConexao()
	
			Set rs2=objConexao.Execute("Select NOME_PRODUTO From TB_PRODUTOS where COD_PRODUTO = '"&COD_PRODUTO&"'")
				
				If not rs2.Eof Then
					nome_produto = rs2(0)
				End If
				
			rs2.Close
			Set rs2=Nothing
	
FecharConexao()

redirecionar_incluir	= "dip_produtos_tabela.asp"
redirecionar_alterar	= "dip_produtos_tabela.asp"
redirecionar_excluir	= "dip_produtos_tabela.asp"
redirecionar_validacao	= "dip_produtos_tabela.asp"
msgAoExcluir			= "Tem certeza que deseja excluir este Preço deste Produto?"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
'Declara variáveis
Dim VALIDO_ATE, VALOR_A_VISTA, VALOR_3X, VALOR_PARCELADO, CADASTRADO_POR, GRUPO_VENDEDORES

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
cod_valor = Request.QueryString("cod_valor")
lerRegistro()
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "alterar" and campo1 <> Empty Then
'Gravar alteração
cod_valor = Request.Form("cod_valor")
alterarRegistro(redirecionar_alterar)
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 
'Leitura inicial para possível exclusão
cod_valor = Request.QueryString("cod_valor")
lerRegistro()
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
'Exclui registro
cod_valor = Request.Form("cod_valor")
excluirRegistro(redirecionar_excluir)
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
'Visualiza registro
cod_valor = Request.QueryString("cod_valor")
lerRegistro()
titulo_pagina = titulo_pagina&" / Visualização"
End If


'Funções

Sub lerFormulario()

COD_VALOR = Request.Form("COD_VALOR")
VALIDO_ATE = Request.Form("VALIDO_ATE")
VALOR_A_VISTA = Request.Form("VALOR_A_VISTA")
VALOR_3X = Request.Form("VALOR_3X")
VALOR_PARCELADO = Request.Form("VALOR_PARCELADO")
CADASTRADO_POR = Request.Form("CADASTRADO_POR")
GRUPO_VENDEDORES = Request.Form("GRUPO_VENDEDORES")
End sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM TB_PRODUTOS_VALORES WHERE cod_valor = "&cod_valor)

				VALIDO_ATE=trim(rs("VALIDO_ATE"))
				VALOR_A_VISTA= ArrumaNumero(rs("VALOR_A_VISTA"),"",2)
				VALOR_3X= ArrumaNumero(rs("VALOR_3X"),"",2)
				VALOR_PARCELADO= ArrumaNumero(rs("VALOR_PARCELADO"),"",2)
				CADASTRADO_POR=trim(rs("CADASTRADO_POR"))
				GRUPO_VENDEDORES=rs("GRUPO_VENDEDORES")

		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("INSERT INTO TB_PRODUTOS_VALORES (GRUPO_VENDEDORES,COD_PRODUTO, VALIDO_ATE, VALOR_A_VISTA, VALOR_3X, VALOR_PARCELADO, CADASTRADO_POR) VALUES ("&sqlTexto(GRUPO_VENDEDORES)&","&sqlTexto(COD_PRODUTO)&", "&sqlData(VALIDO_ATE)&", "&sqlNumero(VALOR_A_VISTA)&", "&sqlNumero(VALOR_3X)&", "&sqlNumero(VALOR_PARCELADO)&", "&sqlNumero(CADASTRADO_POR)&")")

		'Encontrar chave criada
		Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_PRODUTOS_VALORES LIMIT 1")
			If not rs.Eof Then
				cod_valor = rs("ULTIMO_ID")
			End if 
		rs.Close
		Set rs=Nothing


	FecharConexao()
	Response.Redirect(redirecionar&"?cod_valor="&cod_valor&"&codmsg=1&COD_PRODUTO="&COD_PRODUTO)
End sub

Sub alterarRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("UPDATE TB_PRODUTOS_VALORES SET GRUPO_VENDEDORES = "&sqlTexto(GRUPO_VENDEDORES)&", VALIDO_ATE = "&sqlData(VALIDO_ATE)&", VALOR_A_VISTA = "&sqlNumero(VALOR_A_VISTA)&", VALOR_3X = "&sqlNumero(VALOR_3X)&", VALOR_PARCELADO = "&sqlNumero(VALOR_PARCELADO)&", CADASTRADO_POR = "&sqlNumero(CADASTRADO_POR)&" WHERE cod_valor = "&cod_valor)
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_valor="&cod_valor&"&codmsg=2&COD_PRODUTO="&COD_PRODUTO)
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_PRODUTOS_VALORES WHERE cod_valor = "&cod_valor)
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_valor="&cod_valor&"&codmsg=3&COD_PRODUTO="&COD_PRODUTO)
End sub


'----------------------------------------------------------------------------------------

%>

<h2><%=titulo_pagina%></h2><br />

<%If comando = "excluir" Then%>
<div class='msgExcluir'><%=msgAoExcluir%></div>
<%End If%>

<form action='<%=actionformulario%>?comando=<%=comando%>&cod_produto=<%=cod_produto%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post'>
<table width='90%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
<tr>
<td width='27%'><span style="font-weight: bold">Produto</span></td>
<td width='73%'><div align="left" class="textoMedio"><strong><%=nome_produto%></strong></div></td>
</tr>
<tr>
<td width='27%'><span style="font-weight: bold">Destinado aos vendedores *</span></td>
<td width='73%'><div align="left">
<select name="GRUPO_VENDEDORES" class="inputTamanho2">
<option value="" <%If comando = "incluir" Then%> selected="selected" <%End If%>>Selecione</option>
<option value="0" <%IF INT(GRUPO_VENDEDORES) = 0 Then%> selected="selected"<%End If%>>Vendedores Comum</option>
<option value="1" <%IF INT(GRUPO_VENDEDORES) = 1 Then%> selected="selected"<%End If%>>Vendedores Grupo Repasse</option>
</select>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Preço válido até *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<table width='135' border='0' cellspacing='0' cellpadding='0'>
<tr>
<td width='102'><input class='campoData' maxlength='10' onKeyDown='mascara(this,data);' onKeyPress='mascara(this,data);' onKeyUp='mascara(this,data);' type='text' name='VALIDO_ATE' value="<%If(ArrumaData(VALIDO_ATE)) = "-" Then Response.Write("") Else Response.Write(ArrumaData(VALIDO_ATE)) End if%>" onblur='validarData(this)' /></td>
<td width='298'>&nbsp;</td>
</tr>
</table>
<%Else%>
<%If not isNull(VALIDO_ATE) Then%>
<%=VALIDO_ATE%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Valor à vista *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='VALOR_A_VISTA' value="<%=VALOR_A_VISTA%>" class='inputTamanho1 padrao_valor' type='text' maxlength='200' />
<%Else%>
<%If not isNull(VALOR_A_VISTA) Then%>
<%=VALOR_A_VISTA%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Valor em 3X *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='VALOR_3X' value="<%=VALOR_3X%>" class='inputTamanho1 padrao_valor' type='text' maxlength='200' />
<%Else%>
<%If not isNull(VALOR_3X) Then%>
<%=VALOR_3X%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Valor parcelado (acima de 3X) *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='VALOR_PARCELADO' value="<%=VALOR_PARCELADO%>" class='inputTamanho1 padrao_valor' type='text' maxlength='200' />
<%Else%>
<%If not isNull(VALOR_PARCELADO) Then%>
<%=VALOR_PARCELADO%>
<%End If%>
<%End If%>
</div></td>
</tr>
<%If comando <> "visualizar" Then%>
<tr>
  <td colspan='2' align='center'><a href="javascript:salvar();"><img src="images/layout/btnGravar.jpg" /></a></td>
</tr>
<tr>
<td colspan='2' align='center'><div align='left'><strong>* - Campos obrigatórios</strong><br /></div></td>
</tr>
<%End If%>
</table>
<input type='hidden' value='1' name='campo1'  />
<input type='hidden' name='COD_VALOR' value="<%=COD_VALOR%>" /> 
<input type='hidden' name='CADASTRADO_POR' value="<%=Session("COD_USUARIO")%>" /> </form>
<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('GRUPO_VENDEDORES#Preço válidado para qual tipo de vendedores#2|VALIDO_ATE#Preço válido até#1|VALOR_A_VISTA#Valor à vista#1|VALOR_3X#Valor em 3X#1|VALOR_PARCELADO#Valor parcelado#1',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->
