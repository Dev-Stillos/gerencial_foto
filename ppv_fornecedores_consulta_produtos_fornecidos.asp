				
<%
'data in�cio formatado, data t�rmino, t�tulo da p�gina
'call manutencao(ArrumaData(date(),"")& "�s 08h00",ArrumaData(date(),"")& "�s 12h00","")
%>
<%
Dim calendario_necessario,titulo_pagina
calendario_necessario = "NÃO"
titulo_pagina = "Produtos Fornecidos"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%

Dim cod_fornecedor
cod_fornecedor = request.QueryString("cod_fornecedor")

'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 26/6/2012 17:53:05

'Declara vari�veis para utiliza��o do sistema
Dim id, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

redirecionar_incluir	= "ppv_fornecedores_consulta_produtos_fornecidos.asp"
redirecionar_alterar	= "ppv_fornecedores_consulta_produtos_fornecidos.asp"
redirecionar_excluir	= "ppv_fornecedores_consulta_produtos_fornecidos.asp"
redirecionar_validacao	= "ppv_fornecedores_consulta_produtos_fornecidos.asp"
msgAoExcluir			= "Tem certeza que deseja excluir este produto fornecido?<br /><br />Voc� poder� inativa-lo ao inv�s de exclui-lo. Se houver registros com este, produto mesmo NÃO ser� exclu�do"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
'Declara vari�veis
Dim cod_produto, valor_custo, cod_usuario, modificado

'If comando = Empty Then
'comando = "incluir"
'titulo_pagina = titulo_pagina&" / Inclusão"
'End if


'#########################################################################
'Valida��o de sub-lida��es padr�es com formul�rio
'#########################################################################

If comando = "incluir" and campo1 <> Empty Then
'Incluir novo registro
incluirRegistro(redirecionar_incluir)
titulo_pagina = titulo_pagina&" / Inclusão"
Elseif comando = "alterar" and campo1 = Empty Then
'Leitura inicial para poss�vel Alteração
id = Request.QueryString("id")
lerRegistro()
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "alterar" and campo1 <> Empty Then
'Gravar Alteração
id = Request.Form("id")
alterarRegistro(redirecionar_alterar)
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 
'Leitura inicial para poss�vel Exclusão
id = Request.QueryString("id")
lerRegistro()
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
'Exclui registro
id = Request.Form("id")
excluirRegistro(redirecionar_excluir)
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
'Visualiza registro
id = Request.QueryString("id")
lerRegistro()
titulo_pagina = titulo_pagina&" / Visualização"
End If


'Fun��es

Sub lerFormulario()

id = Request.Form("id")
cod_fornecedor = Request.Form("cod_fornecedor")
cod_produto = Request.Form("cod_produto")
valor_custo = Request.Form("valor_custo")
cod_usuario = Session("cod_usuario")
modificado = now()

End sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM TB_FORNECEDOR_PRODUTOS WHERE id = "&id)

				cod_fornecedor=trim(rs("cod_fornecedor"))
				cod_produto=trim(rs("cod_produto"))
				valor_custo= ArrumaNumero(rs("valor_custo"),"",2)
				cod_usuario=trim(rs("cod_usuario"))
				modificado=trim(rs("modificado"))

		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("INSERT INTO TB_FORNECEDOR_PRODUTOS (cod_fornecedor, cod_produto, valor_custo, cod_usuario, modificado) VALUES ("&sqlNumero(cod_fornecedor)&", "&sqlNumero(cod_produto)&", "&sqlNumero(valor_custo)&", "&sqlTexto(cod_usuario)&", "&sqlData(modificado)&")")

		'Encontrar chave criada
		Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_FORNECEDOR_PRODUTOS LIMIT 1")
			If not rs.Eof Then
				id = rs("ULTIMO_ID")
			End if 
		rs.Close
		Set rs=Nothing

		'Grava log
		objConexao.Execute("INSERT INTO TB_FORNECEDOR_PRODUTOS_LOG (id_origem,valor_custo,cod_usuario) VALUES ("&sqlTexto(id)&", "&sqlNumero(valor_custo)&", "&sqlTexto(cod_usuario)&")")

	FecharConexao()
	Response.Redirect(redirecionar&"?cod_fornecedor="&cod_fornecedor&"&id="&id&"&codmsg=1")
End sub

Sub alterarRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("UPDATE TB_FORNECEDOR_PRODUTOS SET cod_fornecedor = "&sqlNumero(cod_fornecedor)&", cod_produto = "&sqlNumero(cod_produto)&", valor_custo = "&sqlNumero(valor_custo)&", cod_usuario = "&sqlTexto(cod_usuario)&", modificado = "&sqlData(modificado)&" WHERE id = "&id)
		
		'Grava log
		objConexao.Execute("INSERT INTO TB_FORNECEDOR_PRODUTOS_LOG (id_origem,valor_custo,cod_usuario) VALUES ("&sqlTexto(id)&", "&sqlNumero(valor_custo)&", "&sqlTexto(cod_usuario)&")")
		
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_fornecedor="&cod_fornecedor&"&id="&id&"&codmsg=2")
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()
		'call verificaAntesExcluir(,id,"TABELASRELACIONADAS;TABELASRELACIONADAS","id",redirecionar_excluir&"?sucesso=2&codmsg=5")

	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_FORNECEDOR_PRODUTOS WHERE id = "&sqlTexto(id))
		objConexao.Execute("DELETE FROM TB_FORNECEDOR_PRODUTOS_LOG WHERE id_origem = "&sqlTexto(id))
		
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_fornecedor="&cod_fornecedor&"&id="&id&"&codmsg=3")
End sub


'----------------------------------------------------------------------------------------

%>

<h2><%=titulo_pagina%></h2><br />
<div class="subMenu">Sub-Menu: <a href="ppv_fornecedores_consulta_produtos_fornecidos.asp?comando=incluir&cod_fornecedor=<%=cod_fornecedor%>">Incluir produto para este forncedor</a> | <a href="ppv_fornecedores.asp">Voltar para Fornecedores</a></div>


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
<fieldset>
<legend>Produtos fornecidos por este forncedor</legend>


<%
AbrirConexao()
Set rs = Objconexao.Execute("SELECT * FROM TB_FORNECEDOR_PRODUTOS WHERE cod_fornecedor = "&sqlnumero(cod_fornecedor))
If not rs.Eof Then
%>
<table width='90%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
<tr bgcolor="whitesmoke">
  <td width="17%"><div align="center"><strong>Cód. Produto</strong></div></td>
  <td width="48%"><div align="center"><strong>Nome Produto</strong></div></td>
  <td width="13%"><div align="center"><strong>Valor de custo atual</strong></div></td>
  <td width="13%"><div align="center"><strong>Respons&aacute;vel</strong></div></td>
  <td width="9%"><div align="center"><strong>&nbsp;</strong></div></td>
</tr>
<%
	While not rs.Eof
	%>
        <tr>
          <td width="17%"><div align="center"><%=rs("cod_produto")%></div></td>
          <td width="48%"><div align="left"><%
		  	Set rs2 = Objconexao.Execute("SELECT nome_produto FROM TB_PRODUTOS  WHERE cod_PRODUTO = "&rs("cod_produto"))
           	If not rs2.Eof Then
            	response.Write(rs2(0))
            End if
            rs2.close()
			Set rs2 = nothing
			%></div></td>
          <td width="13%"><div align="right"><%=ArrumaNumero(rs("valor_custo"),"",2)%></div></td>
          <td width="13%"><%=nomeUsuario(rs("cod_usuario"),true)%></td>
          <td width="9%"><div align="center">
            <table border="0" align="center" cellpadding="3" cellspacing="0"> 
              <tr>
                <td>
                  <a href="ppv_fornecedores_consulta_produtos_fornecidos.asp?cod_fornecedor=<%=cod_fornecedor%>&comando=visualizar&id=<%=rs("id")%>">
                    <img src="images/icons/icon_visualizar.png" alt="Visualizar Registro" title="Visualizar Registro" border="0" />
                    </a>
                  </td>
                <td>
                  <a href="ppv_fornecedores_consulta_produtos_fornecidos.asp?cod_fornecedor=<%=cod_fornecedor%>&comando=alterar&id=<%=rs("id")%>">
                    <img src="images/icons/icon_alterar.png" alt="Alterar Registro" title="Alterar Registro" width="16" height="16" border="0" /></a>  </td>
                <td>  <a href="ppv_fornecedores_consulta_produtos_fornecidos.asp?cod_fornecedor=<%=cod_fornecedor%>&comando=excluir&id=<%=rs("id")%>">
                  <img src="images/icons/icon_excluir.png" alt="Excluir Registro" title="Excluir Registro" border="0" />
                  </a></td>
                </tr>
              </table>
            
          </div></td>
        </tr>
	<%
	rs.movenext
	Wend
	%></table><%
	Else
	%>
<div align="center" class="msgGrande">
    Nenhum produto cadastrado para este Fornecedor !<br />
</div>
	<%
End if
rs.close()
Set rs = nothing
FecharConexao()
%>
</fieldset>
<%if comando <> "" Then%>
<fieldset>
<legend>Descri&ccedil;&atilde;o do Produto</legend>
<form action='<%=actionformulario%>?comando=<%=comando%>&cod_forncedor=<%=cod_forncedor%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post'>
<table width='90%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
<%If comando <> "incluir" Then %><tr>
<td width='27%'><strong>ID </strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<%Else%>
<%If not isNull(ID) Then%>
<%=ID%>
<%End If%>
<%End If%>
</div></td>
</tr>
<%End If%><tr>
<td width='27%'><strong>Cód. Produto *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<select name='COD_PRODUTO' class='inputTamanho3'>
<option value=''>Selecione</option>
<%AbrirConexao()
Set rsE3=objConexao.Execute("SELECT COD_PRODUTO, NOME_PRODUTO FROM `TB_PRODUTOS` WHERE `ATIVO` = 1 ORDER BY NOME_PRODUTO ")
While not rsE3.Eof 
'Padrao
Dim selRelPadrao3, verificador3
verificador3 = ""&rsE3(0)

selRelPadrao3 = "" 

If COD_PRODUTO = Empty or isNull(COD_PRODUTO) Then
COD_PRODUTO = selRelPadrao3
End if

If CStr(verificador3) = CStr(COD_PRODUTO)  Then%>
<option selected='selected' value="<%=rsE3(0)%>"><%=rsE3(1)%></option>
<%Else%>
<option value="<%=rsE3(0)%>"><%=rsE3(1)%></option>
<%End if%>
<%rsE3.MoveNext()
Wend
rsE3.Close
Set rsE3=Nothing
FecharConexao()%>

</select>
<%Else%>
<%If not isNull(COD_PRODUTO) Then%>
<%AbrirConexao()
Set rsE3=objConexao.Execute("SELECT COD_PRODUTO, NOME_PRODUTO FROM `TB_PRODUTOS` WHERE COD_PRODUTO='"&COD_PRODUTO&"' AND  `ATIVO` = 1 ORDER BY NOME_PRODUTO ")
If not rsE3.Eof Then%>
<%=rsE3(1)%>
<%End if%>
<%rsE3.Close
Set rsE3=Nothing
FecharConexao()%>

<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Valor Custo *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='VALOR_CUSTO' value="<%=VALOR_CUSTO%>" class='inputTamanho3 padrao_valor' type='text' maxlength='200' />
<%Else%>
<%If not isNull(VALOR_CUSTO) Then%>
<%=VALOR_CUSTO%>
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
<input type='hidden' name='ID' value="<%=ID%>" /> 
<input type='hidden' name='COD_FORNECEDOR' value="<%=COD_FORNECEDOR%>" /> </form>
</fieldset>

<%If comando <> "incluir" Then%>
    <fieldset>
    	<legend>Histórico de Alteração de Preços</legend>
        <table width='90%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
        <tr bgcolor="whitesmoke">
        <td width='27%'><div align="center"><strong>Data</strong></div></td>
        <td width='36%'><div align="center"><strong>Valor custo</strong></div></td>
        <td width='37%'><div align="center"><strong>Respons&aacute;vel</strong></div></td>
        </tr>
        <%
		AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM  TB_FORNECEDOR_PRODUTOS_LOG WHERE ID_ORIGEM = "&sqlTexto(ID)&" order by id desc")
			While not rs.Eof
		%>
        <tr>
          <td><div align="center"><%=ArrumaData(rs("data_cadastro"))%></div></td>
          <td><div align="right"><%=formatNumber(rs("valor_custo"),2)%></div></td>
          <td><%=nomeUsuario(rs("cod_usuario"),true)%></td>
        </tr>
       	<%
			rs.MoveNext
			Wend
		rs.Close
		Set rs=Nothing
		
		FecharConexao()
		%>
        
        </table>
    </fieldset>
<%End If%>

<%end if%>
<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('COD_PRODUTO#Cód. Produto#2|VALOR_CUSTO#Valor Custo#1',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->
