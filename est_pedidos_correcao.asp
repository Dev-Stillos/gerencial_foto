<%
'data início formatado, data término, título da página
'call manutencao(ArrumaData(date(),"")& "às 08h00",ArrumaData(date(),"")& "às 12h00","")




Dim calendario_necessario, BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
calendario_necessario = "SIM"
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 1
BTN_IMPRIMIR_LINK = "#"
titulo_pagina = "Correção de Pedidos - Formulário"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%

'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 17/12/2011 20:14:45

'Declara variáveis para utilização do sistema
Dim cod_pedido_compra, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

redirecionar_incluir	= "est_estoque.asp"
redirecionar_alterar	= "est_estoque.asp"
redirecionar_excluir	= "est_estoque.asp"
redirecionar_validacao	= "est_estoque.asp"
msgAoExcluir			= "Tem certeza que deseja excluir este Pedido de Compra?"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
'Declara variáveis

Dim cod_pedido 
Dim cod_fornecedor, cod_usuario, status_pedido, data_pedido, data_previsao, data_entrega, descricao
Dim fornecedor, usuario, obs_correcao

If comando = Empty Then
	comando = "incluir"
	titulo_pagina = titulo_pagina&" / Inclusão"
End if

'#########################################################################
'Validação de sub-lidações padrões com formulário
'#########################################################################
If comando = "incluir" and campo1 = Empty Then
	cod_pedido = ""&request.QueryString("cod_pedido")
	lerRegistro()

	if corrigido = "S" and status_pedido = "CD" then
		response.redirect("est_estoque.asp?c=1")
	end if

ElseIf comando = "incluir" and campo1 <> Empty Then
	'Incluir novo registro
	incluirRegistro(redirecionar_incluir)
	titulo_pagina = titulo_pagina&" / Inclusão"

'Elseif comando = "alterar" and campo1 = Empty Then
'Elseif 	comando = "alterar" and campo1 <> Empty Then
'Elseif 	comando = "excluir" and campo1 = Empty Then 
'Elseif 	comando = "excluir" and campo1 <> Empty Then 
'Elseif 	comando = "visualizar" and campo1 = Empty Then 
End If

'Funções
Sub lerFormulario()
	cod_usuario = Session("COD_USUARIO")
	cod_pedido  = request.form("cod_pedido")
	obs_correcao = request.form("obs_correcao")
End sub

sub prt(str)
	response.write(str&"<br />")
end sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT tb_est_pedidos.cod_pedido, tb_est_pedidos.corrigido, tb_est_pedidos.cod_fornecedor, tb_est_pedidos.cod_usuario, tb_est_pedidos.status, tb_est_pedidos.data_pedido, tb_est_pedidos.data_previsao, tb_est_pedidos.data_entrega, tb_est_pedidos.descricao, tb_est_pedidos.obs_correcao, TB_ADM_USUARIOS.NOME, TB_FORNECEDORES.NOME_FANTASIA FROM stillos3.tb_est_pedidos JOIN stillos3.TB_ADM_USUARIOS ON TB_ADM_USUARIOS.COD_USUARIO = tb_est_pedidos.cod_usuario JOIN stillos3.TB_FORNECEDORES ON TB_FORNECEDORES.COD_FORNECEDOR = tb_est_pedidos.cod_fornecedor WHERE cod_pedido = "&sqlNumero(cod_pedido))

			cod_fornecedor = rs("cod_fornecedor")
			cod_usuario = rs("cod_usuario")
			status_pedido = rs("status")
			data_pedido = rs("data_pedido")
			data_previsao = rs("data_previsao")
			data_entrega = rs("data_entrega")
			descricao = trim(replace(rs("descricao"),"	",""))
			fornecedor = rs("NOME_FANTASIA")
			usuario = rs("NOME")
			obs_correcao = rs("obs_correcao")
			corrigido = rs("corrigido")&""

		rs.Close
		Set rs=Nothing

	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		Dim cod_itens_pedidos, quantidade_corrigida, valor_corrgido

		'Atualizamos a tabela de pedidos
		objConexao.execute("UPDATE tb_est_pedidos SET obs_correcao = "&sqlTexto(obs_correcao)&", corrigido = 'S' WHERE cod_pedido = "&sqlNumero(cod_pedido))


		'atualizamos os itens do pedido
		Set rs = objConexao.execute("SELECT cod_itens_pedidos, cod_pedido, cod_item, quantidade, valor FROM `tb_est_itens_pedidos` WHERE cod_pedido =  "&sqlNumero(cod_pedido))
		while not rs.Eof
			cod_itens_pedidos = rs("cod_itens_pedidos")
			quantidade_corrigida = request.form(cod_itens_pedidos&"_qtd")
			valor_corrigido = request.form(cod_itens_pedidos&"_val")

			objConexao.Execute("UPDATE tb_est_itens_pedidos SET quantidade_corrigida = "&sqlNumero(quantidade_corrigida)&" ,valor_corrigido = "&sqlNumero(valor_corrigido) &" WHERE cod_itens_pedidos = " & sqlNumero(cod_itens_pedidos))

		rs.movenext()
		wend
		rs.Close
		set rs = nothing

 		'Inserimos na tabela de movimentações
		objConexao.Execute("INSERT INTO tb_est_movimentacao VALUES(0,'IN','Pedido Cod: "&cod_pedido&" ',1,now(),"&cod_usuario&");")	

		'encontramos o codigo gerado
		set rs = objConexao.Execute("SELECT MAX(cod_movimentacao) AS ID FROM tb_est_movimentacao")
		if not rs.Eof then
			cod_movimentacao = rs("ID")
		end if
		rs.Close
		set rs = nothing

		objConexao.Execute("INSERT INTO tb_est_movimento_pedido VALUES(0,"&cod_movimentacao&","&cod_pedido&")")

		Set rs = objConexao.execute("SELECT  cod_item, quantidade_corrigida from tb_est_itens_pedidos WHERE cod_pedido =  "&sqlNumero(cod_pedido))
		while not rs.Eof
			objConexao.Execute("INSERT INTO tb_est_itens_movimentacao VALUES (0,"&cod_movimentacao&","&rs("cod_item")&","&rs("quantidade_corrigida")&");")

		rs.movenext()
		wend
		rs.Close
		set rs = nothing
    	
    	'Chamamos a função de atualizar o estoque
    	objConexao.Execute("call atualizarEstoque()")
		

	FecharConexao()
	Response.Redirect(redirecionar&"?cod_pedido_compra="&cod_pedido_compra&"&codmsg=1")
End sub

Sub alterarRegistro(redirecionar)

End sub

Sub excluirRegistro(redirecionar)
	
End sub
'----------------------------------------------------------------------------------------

function trataStatusPedido(c)
	Dim retorno
	Select Case c
		Case "CD"
			retorno = "Concluído"
		Case "EA"
			retorno = "Em Andamento"
		Case "CN"
			retorno = "Cancelado" 
	End Select
trataStatusPedido = retorno
end function

%>
<style type="text/css">
.padrao_numero, .padrao_valor{text-align: right;}
</style>
<script language="javascript" type="text/javascript"></script>
<script type="text/javascript"></script>

<h2><%=titulo_pagina%></h2>
<br />
<%If comando = "excluir" Then%>
<div class='msgExcluir'><%=msgAoExcluir%></div>
<%End If%>

<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post' style="margin:0px">      

<fieldset>
	<legend>Dados do Pedido</legend>
	<div class="fieldset">
		<table width='95%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
			<tr>
				<td width='27%'><strong>Status</strong></td>
				<td width='78%'>
					<div align='left'>
						<%=trataStatusPedido(status_pedido)%>
					</div>
				</td>
			</tr>
			<tr> 
				<td><strong>Fornecedor:</strong></td>
				<td><%=fornecedor%></td>
			</tr>
			<tr> 
				<td><strong>Data do Pedido:</strong></td>
				<td><%=arrumaData(data_pedido)%></td>
			</tr>
			<tr> 
				<td><strong>Previsão de entrega:</strong></td>
				<td><%=arrumaData(data_previsao)%></td>
			</tr>
			<tr>
				<td><strong>Entregue em:</strong></td>
				<td><%=arrumaData(data_entrega)%></td>
			</tr>
			<tr>
				<td><strong>Descrição:</strong></td>
				<td><%=descricao%></td>
			</tr>
			<tr>
				<td width='27%'><strong>Observação </strong></td>
				<td width='78%'>
					<div align='left'>
					<%If comando <> "visualizar" and comando <> "excluir"  Then%>
					<table style='width:420px;' border='0' cellspacing='0' cellpadding='0'>
						<tr>
							<td width='57'>
								<textarea rows='2' name='obs_correcao' class='inputTamanho3' onkeydown='textCounter(this.form.obs_correcao,this.form.charRestantes_OBS,500);' onkeyup='textCounter(this.form.obs_correcao,this.form.charRestantes_OBS,500);'><%=obs_correcao%></textarea>
							</td>
						</tr>
						<tr>
							<td>
								<table width='50' border='0' align='right' cellpadding='0' cellspacing='0'>
									<tr>
										<td>
											<div align='center'>
												<input class='inputContaCaracteres' tabindex='10000' type='text' name='charRestantes_OBS' value="<%=500- Len(obs_correcao)%>" />
											</div>
										</td>
									</tr>
								</table>
							</td>
						</tr>
					</table>
					<%
					Else
						If not isNull(obs_correcao) Then
							response.write(Replace(obs_correcao,Chr(13),"<br />"))
						End If
					End If
				    %>
				    </div>
				</td>
			</tr>
		</table>
	</div>
</fieldset>

<%
AbrirConexao()
set rs = objConexao.execute("SELECT `tb_est_itens_pedidos`.`cod_itens_pedidos`, `tb_est_itens_pedidos`.`cod_pedido`, `tb_est_itens_pedidos`.`cod_item`, `tb_est_itens_pedidos`.`quantidade`, `tb_est_itens_pedidos`.`valor`, `TB_PRODUTOS`.`NOME_PRODUTO` FROM `stillos3`.`tb_est_itens_pedidos` JOIN `stillos3`.`TB_PRODUTOS` ON `TB_PRODUTOS`.`COD_PRODUTO` = `tb_est_itens_pedidos`.`cod_item` WHERE cod_pedido =  "&sqlNumero(cod_pedido))

if not rs.Eof then 
%>
<fieldset class="fieldset">
	<legend>Produtos</legend>
	<div class="fieldset">
		<table width='95%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
			<tr>
				<th>#</th>
				<th>Produto</th>
				<th>Qtd. Pedida</th>
				<th>valor inicial</th>
				<th>Qtd. Entregue</th>
				<th>valor Entregue</th>
			</tr>
			<%
			Dim i
			while not rs.Eof
			i = i + 1
			%>
			<tr>
				<td align="center"><%=i%></td>
				<td align="left"><%=rs("NOME_PRODUTO")%></td>
				<td align="right"><%=rs("quantidade")%></td>
				<td align="right">R$ <%=arrumaNumero(rs("valor"),0,2)%></td>
				<td align="center">
					<input type="text" name='<%=rs("cod_itens_pedidos")%>_qtd' value='<%=rs("quantidade")%>' class='inputTamanho1 padrao_numero'>
				</td>
				<td align="center">
					<input type="text" name='<%=rs("cod_itens_pedidos")%>_val' value='<%=arrumaNumero(rs("valor"),0,2)%>' class='inputTamanho1 padrao_valor'>
				</td>
			</tr>
			<%
			rs.movenext
			Wend
			%>
		</table>
	</div>
</fieldset>
<%
End if 

rs.Close
set rs = nothing
FecharConexao()
%>

<%if comando = "incluir" then %>
	<center>
		<a href="javascript:salvar();"><img src="images/layout/btnGravar.jpg"></a>
<%end if %>

<input type='hidden' value='1' name='campo1'  />
<input type='hidden' name='COD_PEDIDO' value="<%=COD_PEDIDO%>" />
</form>


<script type='text/javascript'>
function validaFormMaster() 
{
	return true
}
</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->
