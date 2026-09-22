<%
	Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
	BTN_VOLTAR = 1
	BTN_IMPRIMIR = 1
	BTN_FONTE = 2
	BTN_IMPRIMIR_LINK = "#"
  
	Dim titulo_pagina
	titulo_pagina = "Produtos"

	Dim calendario_necessario
	calendario_necessario = "Não"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%

'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 10/10/2011 16:04:06

'Declara variáveis para utilização do sistema
Dim cod_produto, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

redirecionar_incluir	= "dip_produtos.asp"
redirecionar_alterar	= "dip_produtos.asp"
redirecionar_excluir	= "dip_produtos.asp"
redirecionar_validacao	= "dip_produtos.asp"
msgAoExcluir			= "Tem certeza que deseja excluir este produto?<br /><br />Você poderá inativa-lo ao invés de exclui-lo. Se houver pedidos com este produto, o mesmo Não será excluído"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
'Declara variáveis
Dim SAIDA_AUTOMATICA ,PERMITE_NEGATIVO , ORDER_MENU, NOME_PRODUTO, COD_CATEGORIA, ATIVO, HABILITA_SAIDA_FORNECEDOR, PAGAMENTO_COMISSAO, gera_ordem_servico_prazo_cliente, instrucoes_para_producao_vendedor, NECESSITA_FOTOS

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
	cod_produto = Request.QueryString("cod_produto")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / alteração"
Elseif 	comando = "alterar" and campo1 <> Empty Then
	'Gravar alteração
	cod_produto = Request.Form("cod_produto")
	alterarRegistro(redirecionar_alterar)
	titulo_pagina = titulo_pagina&" / alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 
	'Leitura inicial para possível exclusão
	cod_produto = Request.QueryString("cod_produto")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
	'Exclui registro
	cod_produto = Request.Form("cod_produto")
	excluirRegistro(redirecionar_excluir)
	titulo_pagina = titulo_pagina&" / exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
	'Visualiza registro
	cod_produto = Request.QueryString("cod_produto")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / Visualização"
End If


'Funções

Sub lerFormulario()

	COD_PRODUTO = Request.Form("COD_PRODUTO")
	ORDER_MENU = Request.Form("ORDER_MENU")
	NOME_PRODUTO = Request.Form("NOME_PRODUTO")
	COD_CATEGORIA = Request.Form("COD_CATEGORIA")
	ATIVO = Request.Form("ATIVO")
	HABILITA_SAIDA_FORNECEDOR = Request.Form("HABILITA_SAIDA_FORNECEDOR")
	PAGAMENTO_COMISSAO = Request.Form("PAGAMENTO_COMISSAO")
	gera_ordem_servico_prazo_cliente = Request.Form("gera_ordem_servico_prazo_cliente")
	instrucoes_para_producao_vendedor = Request.Form("instrucoes_para_producao_vendedor")
	SAIDA_AUTOMATICA  = Request.Form("SAIDA_AUTOMATICA")
	PERMITE_NEGATIVO  = Request.Form("PERMITE_NEGATIVO")
	NECESSITA_FOTOS = Request.Form("NECESSITA_FOTOS")

End sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM TB_PRODUTOS WHERE cod_produto = "&cod_produto)

			ORDER_MENU=trim(rs("ORDER_MENU"))
			NOME_PRODUTO=trim(rs("NOME_PRODUTO"))
			COD_CATEGORIA=trim(rs("COD_CATEGORIA"))
			ATIVO=trim(rs("ATIVO"))
			SAIDA_AUTOMATICA  = trim(rs("SAIDA_AUTOMATICA"))
			PERMITE_NEGATIVO  = trim(rs("PERMITE_NEGATIVO"))			
			HABILITA_SAIDA_FORNECEDOR=trim(rs("HABILITA_SAIDA_FORNECEDOR"))
			PAGAMENTO_COMISSAO=trim(rs("PAGAMENTO_COMISSAO"))
			gera_ordem_servico_prazo_cliente=trim(rs("gera_ordem_servico_prazo_cliente"))
			instrucoes_para_producao_vendedor=trim(rs("instrucoes_para_producao_vendedor"))
			NECESSITA_FOTOS=trim(rs("NECESSITA_FOTOS"))

		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("INSERT INTO TB_PRODUTOS (ORDER_MENU, NOME_PRODUTO, COD_CATEGORIA, ATIVO, HABILITA_SAIDA_FORNECEDOR, PAGAMENTO_COMISSAO, gera_ordem_servico_prazo_cliente, instrucoes_para_producao_vendedor,SAIDA_AUTOMATICA,PERMITE_NEGATIVO, NECESSITA_FOTOS) VALUES ("&sqlTexto(ORDER_MENU)&", "&sqlTexto(NOME_PRODUTO)&", "&sqlTexto(COD_CATEGORIA)&", "&sqlTexto(ATIVO)&", "&sqlTexto(HABILITA_SAIDA_FORNECEDOR)&", "&sqlTexto(PAGAMENTO_COMISSAO)&", "&sqlTexto(gera_ordem_servico_prazo_cliente)&", "&sqlTexto(instrucoes_para_producao_vendedor)&","&sqlTexto(SAIDA_AUTOMATICA)&","&sqlTexto(PERMITE_NEGATIVO)&", "&sqlNumero(NECESSITA_FOTOS)&")")

		'Encontrar chave criada
		Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_PRODUTOS LIMIT 1")
			If not rs.Eof Then
				cod_produto = rs("ULTIMO_ID")
			End if 
		rs.Close
		Set rs=Nothing

		dim cod_produtos_filho
		cod_produtos_filho = Request.form("produtos")
		cod_produtos_filho = split(cod_produtos_filho, ",")
		
		objConexao.Execute("DELETE FROM TB_PRODUTOS_VINCULADOS WHERE cod_produto_pai = "&sqlNumero(cod_produto))

		for each p in cod_produtos_filho
			objConexao.Execute("INSERT INTO TB_PRODUTOS_VINCULADOS (cod_produto_pai, cod_produto_filho, quantidade) VALUES ("&sqlNumero(cod_produto)&", "&sqlNumero(p)&", "&sqlNumero(Request.form("produto_"&p))&") ")
		next

	FecharConexao()
	Response.Redirect(redirecionar&"?cod_produto="&cod_produto&"&codmsg=1")
End sub

Sub alterarRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("UPDATE TB_PRODUTOS SET COD_PRODUTO = "&sqlNumero(COD_PRODUTO)&", ORDER_MENU = "&sqlTexto(ORDER_MENU)&", NOME_PRODUTO = "&sqlTexto(NOME_PRODUTO)&", COD_CATEGORIA = "&sqlTexto(COD_CATEGORIA)&", ATIVO = "&sqlTexto(ATIVO)&", HABILITA_SAIDA_FORNECEDOR = "&sqlTexto(HABILITA_SAIDA_FORNECEDOR)&", PAGAMENTO_COMISSAO = "&sqlTexto(PAGAMENTO_COMISSAO)&", gera_ordem_servico_prazo_cliente = "&sqlTexto(gera_ordem_servico_prazo_cliente)&", instrucoes_para_producao_vendedor = "&sqlTexto(instrucoes_para_producao_vendedor)&" ,SAIDA_AUTOMATICA ="&sqlTexto(SAIDA_AUTOMATICA)&",PERMITE_NEGATIVO ="&sqlTexto(PERMITE_NEGATIVO)&", NECESSITA_FOTOS = "&sqlNumero(NECESSITA_FOTOS)&" WHERE cod_produto = "&cod_produto)

		dim cod_produtos_filho
		cod_produtos_filho = Request.form("produtos")
		cod_produtos_filho = split(cod_produtos_filho, ",")
		
		objConexao.Execute("DELETE FROM TB_PRODUTOS_VINCULADOS WHERE cod_produto_pai = "&sqlNumero(cod_produto))

		for each p in cod_produtos_filho
			objConexao.Execute("INSERT INTO TB_PRODUTOS_VINCULADOS (cod_produto_pai, cod_produto_filho, quantidade) VALUES ("&sqlNumero(cod_produto)&", "&sqlNumero(p)&", "&sqlNumero(Request.form("produto_"&p))&") ")
		next
	FecharConexao()

	Response.Redirect(redirecionar&"?cod_produto="&cod_produto&"&codmsg=2")
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()		
	call verificaAntesExcluir(cod_produto,"TB_PEDIDOS_PRODUTOS","cod_produto",redirecionar_excluir&"?sucesso=2&codmsg=5")

	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_PRODUTOS_VINCULADOS WHERE cod_produto_pai = "&cod_produto)
		objConexao.Execute("DELETE FROM TB_PRODUTOS WHERE cod_produto = "&cod_produto)
		objConexao.Execute("DELETE FROM TB_PRODUTOS_VALORES WHERE cod_produto = "&cod_produto)
		objConexao.Execute("DELETE FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES_PRODUTOS WHERE cod_produto = "&cod_produto)
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_produto="&cod_produto&"&codmsg=3")
End sub

'----------------------------------------------------------------------------------------
%>

<h2><%=titulo_pagina%></h2><br />

<%If comando = "excluir" Then%>
	<div class='msgExcluir'><%=msgAoExcluir%></div>
<%End If%>

<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post'>
<fieldset>
<legend>Informações Gerais</legend>
	<table width='90%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
		<%If comando <> "incluir" Then %><tr>
		<td width='27%'><strong>Cod Produto </strong></td>
			<td width='73%'>
				<div align='left'>
				<% If comando <> "visualizar" and comando <> "excluir"  Then%>
					<span class='textoMaior'><strong><%=COD_PRODUTO%></strong></span><%Else%>
					<%
					If not isNull(COD_PRODUTO) Then
						response.write(COD_PRODUTO)
					End If
				End If %>
				</div>
			</td>
		</tr>
		<%end if%>
		
		<tr>
			<td width='27%'><strong>Nome do Produto *</strong></td>
			<td width='73%'>
				<div align='left'>
				<%If comando <> "visualizar" and comando <> "excluir"  Then%>
					<input name='NOME_PRODUTO' value="<%=NOME_PRODUTO%>" class='inputTamanho3' type='text' maxlength='250' />
				<% Else 
					If not isNull(NOME_PRODUTO) Then
						response.write(NOME_PRODUTO)
					End If
				End If %>
				</div>
			</td>
		</tr>
		<tr>
			<td width='27%'><strong>Ordem no Menu </strong></td>
			<td width='73%'>
				<div align='left'>
				<%If comando <> "visualizar" and comando <> "excluir"  Then%>
					<input name='ORDER_MENU' maxlength='3' onkeypress='mascara(this,soNumeros)' class='inputTamanho3' value="<%=ORDER_MENU%>" type='text' />
				<%
				Else
					If not isNull(ORDER_MENU) Then
						response.write(ORDER_MENU)
					End If
				End If %>
				</div>
			</td>
		</tr>
		<tr>
			<td width='27%'><strong>Categoria </strong></td>
			<td width='73%'>
				<div align='left'>
					<%=padraoFormularioSelect("COD_CATEGORIA", "", "inputTamanho3","","","Selecione", COD_CATEGORIA,"SELECT COD_CATEGORIA, NOME_CATEGORIA FROM TB_PRODUTOS_CATEGORIAS ORDER BY NOME_CATEGORIA ASC",comando, false, false, false) %>
				</div>
			</td>
		</tr>
		<tr>
			<td width='27%'><strong>Produto necessita anexar foto *</strong></td>
			<td width='73%'>
				<div align='left'>
				<%
				'selRelPadrao7 = "1" 
				If NECESSITA_FOTOS = Empty or isNull(NECESSITA_FOTOS) Then
					NECESSITA_FOTOS = "1"
				End if
				response.write(padraoFormularioSelect("NECESSITA_FOTOS", "", "inputTamanho1","","","Selecione", NECESSITA_FOTOS,"PADRAO:1#Sim|2#Não",comando, false, false, false))
				%>
				</div>
			</td>
		</tr>
		<tr>
			<td width='27%'><strong>Permitir pagamento de comissão *</strong></td>
			<td width='73%'>
				<div align='left'>
				<%
				'selRelPadrao7 = "1" 
				If PAGAMENTO_COMISSAO = Empty or isNull(PAGAMENTO_COMISSAO) Then
					PAGAMENTO_COMISSAO = "S"
				End if
				response.write(padraoFormularioSelect("PAGAMENTO_COMISSAO", "", "inputTamanho1","","","Selecione", PAGAMENTO_COMISSAO,"PADRAO:S#Sim|S#Não",comando, false, false, false))
				%>
				</div>
			</td>
		</tr>
		<tr>
			<td width='27%'><strong>Ativo *</strong></td>
			<td width='73%'>
				<div align='left'>
				<%
				'selRelPadrao7 = "1" 
				If ATIVO = Empty or isNull(ATIVO) Then
					ATIVO = "1"
				End if
				response.write(padraoFormularioSelect("ATIVO", "", "inputTamanho1","","","Selecione", ATIVO,"PADRAO:1#Sim|2#Não",comando, false, false, false))
				%>
				</div>
			</td>
		</tr>
		</table>
</fieldset>
<fieldset>
	<legend>Dados de Produção</legend>
	<table width='90%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
		<tr>
			<td width='27%'><strong>Produto tem saída do estoque para Fornecedor *</strong></td>
			<td width='73%'>
				<div align='left'>
				<%
				'selRelPadrao7 = "1" 
				If HABILITA_SAIDA_FORNECEDOR = Empty or isNull(HABILITA_SAIDA_FORNECEDOR) Then
					HABILITA_SAIDA_FORNECEDOR = "1"
				End if
				response.write(padraoFormularioSelect("HABILITA_SAIDA_FORNECEDOR", "", "inputTamanho1","","","Selecione", HABILITA_SAIDA_FORNECEDOR,"PADRAO:1#Sim|2#Não",comando, false, false, false))
				%>
				</div>
			</td>
		</tr>
		<tr>
			<td width='27%'><strong>Prazo Ordem de serviço </strong><br />
				<em>(Se Não preenchido ou dor igual a 0 este produto Não gera OS)</em></td>
			<td width='73%'>
				<div align='left'>
				<%If comando <> "visualizar" and comando <> "excluir"  Then%>
					<input name='GERA_ORDEM_SERVICO_PRAZO_CLIENTE' maxlength='3' onkeypress='mascara(this,soNumeros)' class='inputTamanho3' value="<%=GERA_ORDEM_SERVICO_PRAZO_CLIENTE%>" type='text' />
				<% Else
					If not isNull(GERA_ORDEM_SERVICO_PRAZO_CLIENTE) Then
						response.write(GERA_ORDEM_SERVICO_PRAZO_CLIENTE)
					end If
				End If %>
				</div>
			</td>
		</tr>
		<tr>
			<td width='27%'><strong>Instruções para o vendedor referente a produ&ccedil;&atilde;o do produto<br /></strong><em>(Preencha somente se o campo Prazo Ordem de Servi&ccedil;o for maior que 0(zero).) </em></td>
			<td width='73%'>
				<div align='left'>
				<%If comando <> "visualizar" and comando <> "excluir"  Then%>
					<table style='width:420px;' border='0' cellspacing='0' cellpadding='0'>
						<tr>
							<td width='57'><textarea rows='5' name='INSTRUCOES_PARA_PRODUCAO_VENDEDOR' class='inputTamanho3' onkeydown='textCounter(this.form.INSTRUCOES_PARA_PRODUCAO_VENDEDOR,this.form.charRestantes_INSTRUCOES_PARA_PRODUCAO_VENDEDOR,300);' onkeyup='textCounter(this.form.INSTRUCOES_PARA_PRODUCAO_VENDEDOR,this.form.charRestantes_INSTRUCOES_PARA_PRODUCAO_VENDEDOR,300);'><%=INSTRUCOES_PARA_PRODUCAO_VENDEDOR%></textarea></td>
						</tr>
						<tr>
							<td>
								<table width='50' border='0' align='right' cellpadding='0' cellspacing='0'>
									<tr>
										<td>
											<div align='center'>
												<input class='inputContaCaracteres' tabindex='10000' type='text' name='charRestantes_INSTRUCOES_PARA_PRODUCAO_VENDEDOR' value="<%=300- Len(INSTRUCOES_PARA_PRODUCAO_VENDEDOR)%>" />
											</div>
										</td>
									</tr>
								</table>
							</td>
						</tr>
					</table>
				<%Else
					If not isNull(INSTRUCOES_PARA_PRODUCAO_VENDEDOR) Then
						response.write(Replace(INSTRUCOES_PARA_PRODUCAO_VENDEDOR,Chr(13),"<br />"))
					End If
				End If %>
				</div>
			</td>
		</tr>
		
	</table>
	</fieldset>
	<fieldset>
	<legend>Parametros do Estoque</legend>
	<table width='90%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
		<tr>
			<td width='27%'><strong>Saída Automatica</strong></td>
			<td width='73%'>
				<div align='left'>
				<%If comando <> "visualizar" and comando <> "excluir"  Then%>
					<select name='SAIDA_AUTOMATICA' class='inputTamanho2'>
						<%
						'Padrao
						If SAIDA_AUTOMATICA = Empty or isNull(SAIDA_AUTOMATICA) Then
							SAIDA_AUTOMATICA = "0"
						End if
						
						SAIDA_AUTOMATICA0 = ""
						SAIDA_AUTOMATICA1 = ""
						SAIDA_AUTOMATICA2 = ""				

						If CStr(SAIDA_AUTOMATICA) =  "0" Then
							SAIDA_AUTOMATICA0 = "selected='selected'"
						Elseif CStr(SAIDA_AUTOMATICA) = "1" Then
							SAIDA_AUTOMATICA1 = "selected='selected'"	
						Elseif CStr(SAIDA_AUTOMATICA) = "2" Then
							SAIDA_AUTOMATICA2 = "selected='selected'"								
						End if
						%>
						<option value='0' <%=SAIDA_AUTOMATICA0%>>Não</option>
						<option value='1' <%=SAIDA_AUTOMATICA1%>>Por Produção</option>
						<option value='2' <%=SAIDA_AUTOMATICA2%>>Por venda</option>
					</select>
				<%Else
					If not isNull(SAIDA_AUTOMATICA) Then
						If CStr(SAIDA_AUTOMATICA) =  "0" Then%>
							Não
						<%Elseif CStr(SAIDA_AUTOMATICA) =  "1" Then%>
							Por Produção
						<%Elseif CStr(SAIDA_AUTOMATICA) =  "2" Then%>
							Por venda					
						<%End if
					End If
				End If %>
				</div>
			</td>
		</tr>
		<tr class="trEstoqueDeSaida" <%if SAIDA_AUTOMATICA <> "2" then %>style="display:none;" <%end if%>>
			<td><strong>Saida do estoque do: </strong></td>
			<td>
				<div align="left">
					<%=padraoFormularioSelect("estoque_saida","","inputTamanho2","","", "SEM_SELECIONE", estoque_saida,"PADRAO:V#Vendedor|P#Estoque Padrão",comando, false, false, false)%>
				</div>
			</td>
		</tr>
		<tr>
			<td width='27%'><strong>Permite Negativo</strong></td>
			<td width='73%'>
				<div align='left'>
				<%If comando <> "visualizar" and comando <> "excluir" Then%>
					<select name='PERMITE_NEGATIVO' class='inputTamanho1'>
					<%
						'Padrao
						If PERMITE_NEGATIVO = Empty or isNull(PERMITE_NEGATIVO) Then
							PERMITE_NEGATIVO = "0"
						End if
						
						PERMITE_NEGATIVO0 = ""	
						PERMITE_NEGATIVO1 = ""	
					
						If CStr(PERMITE_NEGATIVO) =  "0" Then
							PERMITE_NEGATIVO0 = "selected"								
						ElseIf CStr(PERMITE_NEGATIVO) =  "1" Then
							PERMITE_NEGATIVO1 = "selected"	
						End if
						%>
						<option value='0' <%=PERMITE_NEGATIVO0%>>Não</option>
						<option value='1' <%=PERMITE_NEGATIVO1%>>Sim</option>
					</select>
				<% Else
					If not isNull(PERMITE_NEGATIVO) Then%>	
						<%If CStr(PERMITE_NEGATIVO) =  "1" Then%>
							Sim
						<%Else%>
							Não
						<%End if
					End If
				End If %>
			</div>
		</td>
	</tr>
	<% if comando <> "incluir" then  %>
	<tr>
		<td colspan="2" align="center"><h4>Produtos Vinculados</h4></td>
	</tr>
	<tr>
		<td colspan="2">
			<table width="100%" border="0
			" align="center" cellpadding="5" style="border-collapse:collapse;border:#999999 dashed 1px; padding:2px; text-align:center; background:#FFFFCC; width: 800px;">
				<tr>
					<td align="center"><strong>Produto</strong></td>
					<td align="center"><strong>Qunatidade</strong></td>
					<td align="center" rowspan="2">
						<img src="images/layout/btnIncluir.jpg" id='adicionar_produto' class='padrao_link' />
					</td>
				</tr>
				<tr>
					<td align="center">
						<%=padraoFormularioSelect("cod_produto_incluir","cod_produto_id","inputTamanho3", "", "","Selecione", cod_produto_incluir, "SELECT cod_produto, NOME_PRODUTO FROM TB_PRODUTOS where ativo = 1", "incluir", false, false, false) %>
					</td>
					<td align="center">
						<input type="text" class="padrao_numero inputTamanho1" id="quantidade_id">		
					</td>
				</tr>
			</table>

			<div id="lista" style="margin-top:10px;">
				<table width="90%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
					<tr>
						<td align="center"><strong>Produto</strong></td>
						<td align="center"><strong>Quantidade</strong></td>
						<td align="center"><strong></strong></td>
					</tr>

					<%
					'Buscamos produtos já vinculados


					AbrirConexao()
					'response.write "SELECT * FROM VW_TB_PRODUTOS_VINCULADOS WHERE cod_produto_pai = "&cod_produto
					set rs = objConexao.execute("SELECT * FROM VW_TB_PRODUTOS_VINCULADOS WHERE cod_produto_pai = "&cod_produto)
					if not rs.Eof then
						while not rs.Eof
						%>
						<tr id='tr_<%=rs("cod_produto_filho")%>'>
							<td><%=rs("NOME_PRODUTO_FILHO")%><input type='hidden' value='<%=rs("cod_produto_filho")%>' name='produtos' /></td>
							<td><%=rs("quantidade")%><input type='hidden' value='<%=rs("quantidade")%>' name='produto_<%=rs("cod_produto_filho")%>' /></td>
							<td>
								<table width='100%'>
								<tr>
									<td align='center'>
										<img src='images/icons/icon_excluir.png' alt='excluir Registro' title='excluir Registro' border='0' class='remover_produto' cp='<%=rs("cod_produto_filho")%>' />
									</td>
								</tr>
							</table>
							</td>
						</tr>
						<% 
						rs.movenext
						Wend
						display_vazio = " style='display:none;' "
					end if
					rs.Close
					set rs = Nothing
					FecharConexao()
					%>
					<tr id='aviso_vazio' <%=display_vazio%>>
						<td colspan="3" align="center"> 
							<strong>Sem produtos Vinculados</strong>
						</td>
					</tr>
				</table>

		</td>
	</tr>

	<%
	end if
	If comando <> "visualizar" Then%>
	<tr>
		<td colspan='2' align='center'><div align='left'><strong>* - Campos obrigatórios</strong><br /></div></td>
	</tr>
	<tr>
		<td colspan='2' align='center'><a href="javascript:salvar();"><img src="images/layout/btnGravar.jpg" /></a></td>
	</tr>
	<%End If%>
</table>
</fieldset>
<input type='hidden' value='1' name='campo1'  />
<input type='hidden' name='COD_PRODUTO' value="<%=COD_PRODUTO%>" />�</form>
<script type="text/javascript">
	$(function(){
		$('select[name=SAIDA_AUTOMATICA]').change(function(){
			if ($(this).val() == 1)
				$('.trEstoqueDeSaida').css('display','none');
			else
				$('.trEstoqueDeSaida').css('display','table-row');

		});
	});
</script>
<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('NOME_PRODUTO#Nome do Produto#1|ATIVO#Ativo#2|HABILITA_SAIDA_FORNECEDOR#Produto tem saída do estoque para Fornecedor#2|PAGAMENTO_COMISSAO#Permitir pagamento de comissão#2',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>
<script type="text/javascript">
	$(function(){
		$("select[name=COD_VENDEDOR]").val("");
		function mostrarAviso()
		{
			if($("#lista > table > tbody > tr").length > 2)
				$("#aviso_vazio").css("display", "none");
			else
				$("#aviso_vazio").css("display", "table-row");

			return true
		}
		
		$("select[name=COD_VENDEDOR]").change(function(){
			if($("select[name=COD_VENDEDOR]").val() == "")
				$("#produtosInc").css("display", "none");
			else
				$("#produtosInc").css("display", "block");

		});

		$(document).on("click", ".remover_produto", function() {		
			$("#tr_"+$(this).attr('cp')).remove();
			mostrarAviso()
		});

 		$("#adicionar_produto").click(function(){
 		//verificar se o produto já está na lista se estiver exibe alerta

 		var cp = $("#cod_produto_id").val();
 		var NOME_PRODUTO = $("#cod_produto_id option[value="+cp+"]").text();
 		var qtd = $("#quantidade_id").val();

 		if( $("#tr_"+cp).length > 0)
 		{
 			alert('O produto selecionado já está na lista!');
 		}
 		else //se Não existir
 		{
 			var ultimo = $("#lista > table > tbody:nth-last-child(1)");

 			var html =  "<tr id='tr_"+cp+"'>"+
 							"<td>"+NOME_PRODUTO+"<input type='hidden' value='"+cp+"' name='produtos' /></td>"+
 							"<td>"+qtd+"<input type='hidden' value='"+qtd+"' name='produto_"+cp+"' /></td>"+
 							"<td>"+
 								"<table width='100%'>"+
									"<tr>"+
										"<td align='center'>"+
											"<img src='images/icons/icon_excluir.png' alt='excluir Registro' title='excluir Registro' border='0' class='remover_produto' cp='"+cp+"' />"+
										"</td>"+
									"</tr>"+
								"</table>"+
 							"</td>"+
 						"</tr>";
 			ultimo.append(html);	
 			mostrarAviso()
		}

 		});
	});
</script>
<!-- #include file="files/includes/miolo-inferior.asp" -->