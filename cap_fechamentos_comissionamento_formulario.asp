<%

Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
BTN_VOLTAR = 1
BTN_IMPRIMIR = 1
BTN_FONTE = 2
BTN_IMPRIMIR_LINK = "#"
  
Dim titulo_pagina
titulo_pagina = "CAP - Fechamento - Comissões sobre Vendas Formulário"

Dim calendario_necessario
calendario_necessario = "SIM"

%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%

'Declara variáveis para utilização do sistema
Dim cod_comissionamento, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

redirecionar_incluir	= "cap_fechamentos_comissionamento.asp"
redirecionar_alterar	= "cap_fechamentos_comissionamento.asp"
redirecionar_excluir	= "cap_fechamentos_comissionamento.asp"
redirecionar_validacao	= "cap_fechamentos_comissionamento.asp"
msgAoExcluir			= "Tem certeza que deseja excluir este comissionamento? Obs. Se estiver sendo utilizado em algum fechamento, o mesmo não será excluído"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
'Declara variáveis
Dim TIPO_VENDA, FORMA_PAGAMENTO, PERCENTUAL, DATA_INICIO

If comando = Empty Then
comando = "incluir"

titulo_pagina = titulo_pagina&" / Inclusão"
End if


If comando = "incluir" and campo1 <> Empty Then
'Incluir novo registro
incluirRegistro(redirecionar_incluir)
titulo_pagina = titulo_pagina&" / Inclusão"
Elseif 	comando = "excluir" and campo1 = Empty Then 
'Leitura inicial para possível exclusão
cod_comissionamento = Request.QueryString("cod_comissionamento")
lerRegistro()
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
'Exclui registro
cod_comissionamento = Request.Form("cod_comissionamento")
excluirRegistro(redirecionar_excluir)
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
'Visualiza registro
cod_comissionamento = Request.QueryString("cod_comissionamento")
lerRegistro()
titulo_pagina = titulo_pagina&" / Visualização"
End If


'Funções

Sub lerFormulario()

COD_COMISSIONAMENTO = Request.Form("COD_COMISSIONAMENTO")
TIPO_VENDA = Request.Form("TIPO_VENDA")
FORMA_PAGAMENTO = Request.Form("FORMA_PAGAMENTO")
PERCENTUAL = Request.Form("PERCENTUAL")
DATA_INICIO = Request.Form("DATA_INICIO")

End sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM TB_FECHAMENTO_COMISSIONAMENTO_PADRAO WHERE cod_comissionamento = "&cod_comissionamento)

				TIPO_VENDA=trim(rs("TIPO_VENDA"))
				FORMA_PAGAMENTO=trim(rs("FORMA_PAGAMENTO"))
				PERCENTUAL=trim(rs("PERCENTUAL"))
				DATA_INICIO=trim(rs("DATA_INICIO"))

		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("INSERT INTO TB_FECHAMENTO_COMISSIONAMENTO_PADRAO (COD_COMISSIONAMENTO, TIPO_VENDA, FORMA_PAGAMENTO, PERCENTUAL, DATA_INICIO) VALUES ("&sqlTexto(COD_COMISSIONAMENTO)&", "&sqlTexto(TIPO_VENDA)&", "&sqlTexto(FORMA_PAGAMENTO)&", "&sqlNumero(PERCENTUAL)&", "&sqlData(DATA_INICIO)&")")

		'Encontrar chave criada
		Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_FECHAMENTO_COMISSIONAMENTO_PADRAO LIMIT 1")
			If not rs.Eof Then
				cod_comissionamento = rs("ULTIMO_ID")
			End if 
		rs.Close
		Set rs=Nothing


	FecharConexao()
	Response.Redirect(redirecionar&"?cod_comissionamento="&cod_comissionamento&"&codmsg=1")
End sub

'COMANDO ALTERAR INATIVADO, POIS ESTÁ TELA  NÃO PERMITE ALTERAÇÃO, SOMENTE INCLUSÃO OU EXCLUSÃO
Sub alterarRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("UPDATE TB_FECHAMENTO_COMISSIONAMENTO_PADRAO SET TIPO_VENDA = "&sqlTexto(TIPO_VENDA)&", FORMA_PAGAMENTO = "&sqlTexto(FORMA_PAGAMENTO)&", PERCENTUAL = "&sqlNumero(PERCENTUAL)&", DATA_INICIO = "&sqlData(DATA_INICIO)&" WHERE cod_comissionamento = "&cod_comissionamento)
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_comissionamento="&cod_comissionamento&"&codmsg=2")
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()
		
	call verificaAntesExcluir(cod_comissionamento,"TB_FECHAMENTO_PAGAMENTO","cod_comissionamento",redirecionar_excluir&"?sucesso=2&codmsg=5")

	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_FECHAMENTO_COMISSIONAMENTO_PADRAO WHERE cod_comissionamento = "&cod_comissionamento)
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_comissionamento="&cod_comissionamento&"&codmsg=3")
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
<td width='27%'><strong>Cód Comissionamento </strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<span class='textoMaior'><strong><%=COD_COMISSIONAMENTO%></strong></span><%Else%>
<%If not isNull(COD_COMISSIONAMENTO) Then%>
<%=COD_COMISSIONAMENTO%>
<%End If%>
<%End If%>
</div></td>
</tr>
<%End If%><tr>
<td width='27%'><strong>Tipo *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<select name='TIPO_VENDA' class='inputTamanho2'>
<option value=''>Selecione</option>
<%'Padrao
Dim selRelPadrao2
selRelPadrao2 = "" 

If TIPO_VENDA = Empty or isNull(TIPO_VENDA) Then
TIPO_VENDA = selRelPadrao2
End if%>

<%If CStr(TIPO_VENDA) =  "1" Then
%><option selected='selected' value="1">Padrão</option>
<%Else%>
<option value="1">Padrão</option>
<%End if%>
<%If CStr(TIPO_VENDA) =  "2" Then
%><option selected='selected' value="2">Repasse</option>
<%Else%>
<option value="2">Repasse</option>
<%End if%>
<%If CStr(TIPO_VENDA) =  "3" Then
%><option selected='selected' value="3">Repasse Grupo Especial</option>
<%Else%>
<option value="3">Repasse Grupo Especial</option>
<%End if%>
</select>
<%Else%>
<%If not isNull(TIPO_VENDA) Then%>
<%If CStr(TIPO_VENDA) =  "1" Then%>
Padrão<%End if%>
<%If CStr(TIPO_VENDA) =  "2" Then%>
Repasse<%End if%>
<%If CStr(TIPO_VENDA) =  "3" Then%>
Repasse Grupo Especial<%End if%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Forma de Pagamento *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<select name='FORMA_PAGAMENTO' class='inputTamanho1'>
<option value=''>Selecione</option>
<%'Padrao
Dim selRelPadrao3
selRelPadrao3 = "" 

If FORMA_PAGAMENTO = Empty or isNull(FORMA_PAGAMENTO) Then
FORMA_PAGAMENTO = selRelPadrao3
End if%>

<%If CStr(FORMA_PAGAMENTO) =  "1" Then
%><option selected='selected' value="1">À vista</option>
<%Else%>
<option value="1">À vista</option>
<%End if%>
<%If CStr(FORMA_PAGAMENTO) =  "2" Then
%><option selected='selected' value="2">À prazo</option>
<%Else%>
<option value="2">À prazo</option>
<%End if%>
</select>
<%Else%>
<%If not isNull(FORMA_PAGAMENTO) Then%>
<%If CStr(FORMA_PAGAMENTO) =  "1" Then%>
À vista<%End if%>
<%If CStr(FORMA_PAGAMENTO) =  "2" Then%>
À prazo<%End if%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Percentual sobre as vendas *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='PERCENTUAL' maxlength='5,2' onkeypress='mascara(this,soNumeros)' class='inputTamanho1' value="<%=PERCENTUAL%>" type='text' />
<%Else%>
<%If not isNull(PERCENTUAL) Then%>
<%=PERCENTUAL%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Data início do novo comissionamento</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<%=Date%>
<br />
Observação: Se você desejar cadastrar um novo comissionamento sempre faça com antecedência de pelo menos 30 dias ao próximo pagamento, pois caso contrário as comissões cadastradas ficarão somente para o pagamento futuro a data de início. Ex.: Data de início 12/08/11, será válida somente para o pagamento de 09/2011, pois no mês de agosto já foi.

<%Else%>
<%If not isNull(DATA_INICIO) Then%>
<%=DATA_INICIO%>
<%End If%>
<%End If%>
</div></td>
</tr>
<%If comando <> "visualizar" Then%>
<tr>
<td colspan='2' align='center'><div align='left'>
  <p><strong>* - Campos obrigatórios</strong></p>
  <p align="center"><a href="javascript:salvar();"><img border="0" src="images/layout/btnGravar.jpg" /></a>
  </p>
</div></td>
</tr>
<%End If%>
</table>
<input type='hidden' value='1' name='campo1'  />
<input type='hidden' name='COD_COMISSIONAMENTO' value="<%=COD_COMISSIONAMENTO%>" /> 
<input class='campoData' maxlength='12' onkeyup='mascara(this,data);' type='hidden' name='DATA_INICIO' value="<%=Replace(ArrumaData(Date()),"-","")%>"  />
</form>
<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('TIPO_VENDA#Tipo#2|FORMA_PAGAMENTO#Forma de Pagamento#2|PERCENTUAL#Percentual sobre as vendas#1',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->
