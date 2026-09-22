<%

Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 2
BTN_IMPRIMIR_LINK = "#"

Dim titulo_pagina
titulo_pagina = "CAP - Fechamento - Comissões sobre Vendas dos Coordenadores Formulário"

Dim calendario_necessario : calendario_necessario = "NÃO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%

Dim cod_comissionamento_coordenadores, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

redirecionar_incluir	= "cap_fechamentos_comissionamento.asp"
redirecionar_alterar	= "cap_fechamentos_comissionamento.asp"
redirecionar_excluir	= "cap_fechamentos_comissionamento.asp"
redirecionar_validacao	= "cap_fechamentos_comissionamento.asp"
msgAoExcluir			= "Tem certeza que deseja excluir este Comissionamento de Coordenador de Repasse?"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
'Declara variáveis
Dim PERCENTUAL, DATA_INICIO

If comando = Empty Then
comando = "incluir"
titulo_pagina = titulo_pagina&" / Inclusão"
DATA_INICIO = Date()
End if


If comando = "incluir" and campo1 <> Empty Then
'Incluir novo registro
incluirRegistro(redirecionar_incluir)
titulo_pagina = titulo_pagina&" / Inclusão"
Elseif 	comando = "excluir" and campo1 = Empty Then 
'Leitura inicial para possível exclusão
cod_comissionamento_coordenadores = Request.QueryString("cod_comissionamento_coordenadores")
lerRegistro()
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
'Exclui registro
cod_comissionamento_coordenadores = Request.Form("cod_comissionamento_coordenadores")
excluirRegistro(redirecionar_excluir)
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
'Visualiza registro
cod_comissionamento_coordenadores = Request.QueryString("cod_comissionamento_coordenadores")
lerRegistro()
titulo_pagina = titulo_pagina&" / Visualização"
End If


'Funções

Sub lerFormulario()

PERCENTUAL = Request.Form("PERCENTUAL")
DATA_INICIO = Request.Form("DATA_INICIO")

End sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM TB_FECHAMENTO_COMISSIONAMENTO_COORDENADORES WHERE cod_comissionamento_coordenadores = "&cod_comissionamento_coordenadores)

				PERCENTUAL=trim(rs("PERCENTUAL"))
				DATA_INICIO=trim(rs("DATA_INICIO"))

		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("INSERT INTO TB_FECHAMENTO_COMISSIONAMENTO_COORDENADORES (PERCENTUAL, DATA_INICIO) VALUES ("&sqlNumero(PERCENTUAL)&", "&sqlData(DATA_INICIO)&")")

		'Encontrar chave criada
		Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_FECHAMENTO_COMISSIONAMENTO_COORDENADORES LIMIT 1")
			If not rs.Eof Then
				cod_comissionamento_coordenadores = rs("ULTIMO_ID")
			End if 
		rs.Close
		Set rs=Nothing
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_comissionamento_coordenadores="&cod_comissionamento_coordenadores&"&codmsg=1")
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()
	
	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_FECHAMENTO_COMISSIONAMENTO_COORDENADORES WHERE cod_comissionamento_coordenadores = "&cod_comissionamento_coordenadores)
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_comissionamento_coordenadores="&cod_comissionamento_coordenadores&"&codmsg=3")
End sub

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
<td width='27%'><strong>ID </strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<span class='textoMaior'><strong><%=COD_COMISSIONAMENTO_COORDENADORES%></strong></span><%Else%>
<%If not isNull(COD_COMISSIONAMENTO_COORDENADORES) Then%>
<%=COD_COMISSIONAMENTO_COORDENADORES%>
<%End If%>
<%End If%>
</div></td>
</tr>
<%End If%><tr>
<td width='27%'><strong>Percentual sobre as vendas *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='PERCENTUAL' maxlength='2' onkeypress='mascara(this,soNumeros)' class='inputTamanho1' value="<%=PERCENTUAL%>" type='text' />
<%Else%>
<%If not isNull(PERCENTUAL) Then%>
<%=PERCENTUAL%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Data do In&iacute;cio</strong></td>
<td width='73%'><%=DATA_INICIO%><br />
Observação: Se você desejar cadastrar um novo comissionamento sempre faça com antecedência de pelo menos 30 dias ao próximo pagamento, pois caso contrário as comissões cadastradas ficarão somente para o pagamento futuro a data de início. Ex.: Data de início 12/08/11, será válida somente para o pagamento de 09/2011, pois no mês de agosto já foi.</td>
</tr>
<%If comando <> "visualizar" Then%>
<tr>
<td colspan='2' align='center'><div align='left'><strong>* - Campos obrigatórios</strong>  <p align="center"><a href="javascript:salvar();"><img border="0" src="images/layout/btnGravar.jpg" /></a>
  </p></div></td>
</tr>
<%End If%>
</table>
<input type='hidden' value='1' name='campo1'  />
<input type='hidden' name='COD_COMISSIONAMENTO_COORDENADORES' value="<%=COD_COMISSIONAMENTO_COORDENADORES%>" /> <input type='hidden' name='DATA_INICIO' value="<%=DATA_INICIO%>" /> </form>
<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('PERCENTUAL#Percentual sobre as vendas#1',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->
