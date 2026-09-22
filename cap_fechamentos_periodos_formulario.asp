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
'Criado OXWebDeveloper em: 10/1/2012 19:56:47

'Declara variáveis para utilização do sistema
Dim cod_fechamento_periodo, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

redirecionar_incluir	= "cap_fechamentos_periodos.asp"
redirecionar_alterar	= "cap_fechamentos_periodos.asp"
redirecionar_excluir	= "cap_fechamentos_periodos.asp"
redirecionar_validacao	= "cap_fechamentos_periodos.asp"
msgAoExcluir			= "Tem certeza que deseja excluir este Período de Vendas?Se este período já estiver em uso o mesmo não será excluido."
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
'Declara variáveis
Dim data_inicio, data_termino, data_termino_lancamento_vendas, data_pagamento, cod_usuario, modificado

If comando = Empty Then
comando = "incluir"
titulo_pagina = titulo_pagina&" / Inclusão"
End if

'COMANDO DE ALTERAÇÃO DESABILITADO!!! PARA NÃO GERAR INCONSISTÊNCIAS

If comando = "incluir" and campo1 <> Empty Then
'Incluir novo registro
incluirRegistro(redirecionar_incluir)
titulo_pagina = titulo_pagina&" / Inclusão"
'Elseif comando = "alterar" and campo1 = Empty Then
'Leitura inicial para possível alteração
'cod_fechamento_periodo = Request.QueryString("cod_fechamento_periodo")
'lerRegistro()
'titulo_pagina = titulo_pagina&" / Alteração"
'Elseif 	comando = "alterar" and campo1 <> Empty Then
'Gravar alteração
'cod_fechamento_periodo = Request.Form("cod_fechamento_periodo")
'alterarRegistro(redirecionar_alterar)
'titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 
'Leitura inicial para possível exclusão
cod_fechamento_periodo = Request.QueryString("cod_fechamento_periodo")
lerRegistro()
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
'Exclui registro
cod_fechamento_periodo = Request.Form("cod_fechamento_periodo")
excluirRegistro(redirecionar_excluir)
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
'Visualiza registro
cod_fechamento_periodo = Request.QueryString("cod_fechamento_periodo")
lerRegistro()
titulo_pagina = titulo_pagina&" / Visualização"
End If


'Funções

Sub lerFormulario()

cod_fechamento_periodo = Request.Form("cod_fechamento_periodo")
data_inicio = Request.Form("data_inicio")
data_termino = Request.Form("data_termino")
data_termino_lancamento_vendas = Request.Form("data_termino_lancamento_vendas")
data_pagamento = Request.Form("data_pagamento")
cod_usuario = Session("cod_usuario")
modificado = now()

End sub

Sub lerRegistro()
	AbrirConexao
		Set rs=objConexao.Execute("SELECT * FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE cod_fechamento_periodo = "&cod_fechamento_periodo)

				data_inicio=trim(rs("data_inicio"))
				data_termino=trim(rs("data_termino"))
				data_termino_lancamento_vendas=trim(rs("data_termino_lancamento_vendas"))
				data_pagamento=trim(rs("data_pagamento"))
				cod_usuario=trim(rs("cod_usuario"))
				modificado=trim(rs("modificado"))


		rs.Close
		Set rs=Nothing
	FecharConexao
End sub

Sub incluirRegistro(redirecionar)
	lerFormulario()
	
	
	'1) Verifica se há um outro período de venda com mesma data de pagamento, pois não é permitido devido aos extratos em pdf que já vão com a data de pagamento e se houver mais que um no mesmo dia dará inconsistência no sistema
	
	Dim chavesUnicasDuplicadas,camposUnicos
	call chaveUnica ("TB_FECHAMENTO_PAGAMENTO_PERIODOS","data_pagamento",Replace(sqlData(data_pagamento),"'",""),redirecionar&"?cod_fechamento_periodo="&cod_fechamento_periodo&"&codmsg=4&sucesso=2&chaveUnica=",comando,"cod_fechamento_periodo",cod_fechamento_periodo)

	AbrirConexao()
	
		'///////////////////////////////////
		'Outras Validações
		
			'2) Verifica se a data início não é maior que a data fim
			Dim temp
			temp = 0
			temp = DateDiff("d",formatDateTime(data_inicio),formatDateTime(data_termino))
			If int(temp) < 0 Then
				Response.Redirect(redirecionar&"?codmsg=401&sucesso=2")
			End If
			
			'3) Verifica se a data término é menor que a data término de lançamento de vendas
			temp = DateDiff("d",formatDateTime(data_termino),formatDateTime(data_termino_lancamento_vendas))
			If int(temp) < 0 Then
				Response.Redirect(redirecionar&"?codmsg=402&sucesso=2")
			End If	

			'4) Verifica se a data máxima de lançamento de vendas não é maior que a data de pagamento
			temp = DateDiff("d",formatDateTime(data_termino_lancamento_vendas),formatDateTime(data_pagamento))
			If int(temp) < 0 Then
				Response.Redirect(redirecionar&"?codmsg=403&sucesso=2")
			End If	

			'5) Verifica o período informado não está em um intervalo de outro período de vendas
			Set rs=objConexao.Execute("SELECT cod_fechamento_periodo FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE  (data_inicio between "&sqlData(data_inicio)&" and "&sqlData(data_termino)&") or (data_termino between "&sqlData(data_inicio)&" and "&sqlData(data_termino)&")")
				
				If not rs.Eof Then
					Response.Redirect(redirecionar&"?codmsg=404&sucesso=2")
				End If
			
			rs.Close
			Set rs=Nothing	


			If int(temp) < 0 Then
				Response.Redirect(redirecionar&"?codmsg=403&sucesso=2")
			End If				
			
			'6) Verifica se as datas de início e fim fazem parte de meses diferentes não é permitido devido a inconsistências na finalização do pedido para verificação qual data de recebimento do vendedor
			If (month(data_inicio) <> month(data_termino)) or (year(data_inicio) <> year(data_termino)) Then
				Response.Redirect(redirecionar&"?codmsg=405&sucesso=2")			
			End If
			
		
		'///////////////////////////////////
	
	
		objConexao.Execute("INSERT INTO TB_FECHAMENTO_PAGAMENTO_PERIODOS (data_inicio, data_termino, data_termino_lancamento_vendas, data_pagamento, cod_usuario, modificado) VALUES ("&sqlData(data_inicio)&", "&sqlData(data_termino)&", "&sqlData(data_termino_lancamento_vendas)&", "&sqlData(data_pagamento)&", "&sqlTexto(cod_usuario)&", "&sqlData(modificado)&")")


	FecharConexao
	Response.Redirect(redirecionar&"?cod_fechamento_periodo="&cod_fechamento_periodo&"&codmsg=1")
End sub

Sub alterarRegistro(redirecionar)
	'////
	'Desabilitado devido inconsistências
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()
		
		call verificaAntesExcluir(cod_fechamento_periodo,"TB_PEDIDOS","cod_fechamento_periodo",redirecionar_excluir&"?sucesso=2&codmsg=5")

	AbrirConexao
		objConexao.Execute("DELETE FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE cod_fechamento_periodo = "&cod_fechamento_periodo)
	FecharConexao
	Response.Redirect(redirecionar&"?cod_fechamento_periodo="&cod_fechamento_periodo&"&codmsg=3")
End sub


'----------------------------------------------------------------------------------------

%>

<h2><%=titulo_pagina%></h2><br />

<%If comando = "excluir" Then%>
<div class='msgExcluir'><%=msgAoExcluir%></div>
<%End If%>

<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post'>
<table width='90%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
<%If comando <> "incluir" Then %>
<tr>
<td width='27%'><strong>Cód. Fechamento Período </strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<span class='textoMaior'><strong><%=COD_FECHAMENTO_PERIODO%></strong></span><%Else%>
<%If not isNull(COD_FECHAMENTO_PERIODO) Then%>
<%=COD_FECHAMENTO_PERIODO%>
<%End If%>
<%End If%>
</div></td>
</tr>
<%Else%>
<tr>
<td colspan="2"><div align="center"><em>Aten&ccedil;&atilde;o n&atilde;o &eacute; permitido a altera&ccedil;&atilde;o do Per&iacute;odo de Vendas ap&oacute;s a inclus&atilde;o apenas &eacute; poss&iacute;vel excluir o per&iacute;odo*<br />
  * Exclus&atilde;o permitida somente se n&atilde;o houver pedidos utilizando este per&iacute;odo de vendas para pagamento</em></div></td>
</tr>
<%End If%><tr>
<td width='27%'><strong>Data Início *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<table width='135' border='0' cellspacing='0' cellpadding='0'>
<tr>
<td width='102'><input class='campoData padrao_data' maxlength='10'  type='text' name='DATA_INICIO' value="<%=ArrumaData(DATA_INICIO)%>" onblur='validarData(this)' /></td><td width='298'><div align='center'></div></td>
</tr>
</table>
<%Else%>
<%If not isNull(DATA_INICIO) Then%>
<%=DATA_INICIO%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Data Término *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<table width='135' border='0' cellspacing='0' cellpadding='0'>
<tr>
<td width='102'><input class='campoData padrao_data' maxlength='10'  type='text' name='DATA_TERMINO' value="<%=ArrumaData(DATA_TERMINO)%>" onblur='validarData(this)' /></td><td width='298'><div align='center'></div></td>
</tr>
</table>
<%Else%>
<%If not isNull(DATA_TERMINO) Then%>
<%=DATA_TERMINO%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Data Término Lançamento de Vendas *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<table width='135' border='0' cellspacing='0' cellpadding='0'>
<tr>
<td width='102'><input class='campoData padrao_data' maxlength='10'  type='text' name='DATA_TERMINO_LANCAMENTO_VENDAS' value="<%=ArrumaData(DATA_TERMINO_LANCAMENTO_VENDAS)%>" onblur='validarData(this)' /></td><td width='298'><div align='center'></div></td>
</tr>
</table>
<%Else%>
<%If not isNull(DATA_TERMINO_LANCAMENTO_VENDAS) Then%>
<%=DATA_TERMINO_LANCAMENTO_VENDAS%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Data Pagamento *</strong><br />
(É permitido apenas 1 pagamento por dia)</td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<table width='135' border='0' cellspacing='0' cellpadding='0'>
<tr>
<td width='102'><input class='campoData padrao_data' maxlength='10'  type='text' name='DATA_PAGAMENTO' value="<%=ArrumaData(DATA_PAGAMENTO)%>" onblur='validarData(this)' /></td><td width='298'><div align='center'></div></td>
</tr>
</table>
<%Else%>
<%If not isNull(DATA_PAGAMENTO) Then%>
<%=DATA_PAGAMENTO%>
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
<input type='hidden' name='COD_FECHAMENTO_PERIODO' value="<%=COD_FECHAMENTO_PERIODO%>" /> </form>
<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('DATA_INICIO#Data Início#1|DATA_TERMINO#Data Término#1|DATA_TERMINO_LANCAMENTO_VENDAS#Data Término Lançamento de Vendas#1',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->
