<%
Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
BTN_VOLTAR = 1
BTN_IMPRIMIR = 1
BTN_FONTE = 2
BTN_IMPRIMIR_LINK = "#"

Dim titulo_pagina
titulo_pagina = "Triagem Entrega | Transpostadora"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<%

dim cod_pedido, cod_ordem_servico, campo1, justificativa,despacho_transportador,msg
cod_pedido = request.form("cod_pedido")
cod_ordem_servico = request.form("cod_ordem_servico")
despacho_transportador = request.form("despacho_transportador")
justificativa = request.form("justificativa")
campo1 = request.form("campo1")
msg = request.querystring("msg")

if campo1 = "1" and despacho_transportador <> "" and cod_ordem_servico <> "" then

	if justificativa = "" then
		response.redirect("ppv_ordens_servico_triagem_despacho_alterar.asp?msg=justificativa inválida")
	end if

	abrirConexao()

	Set rs = objConexao.Execute("SELECT * FROM TB_ORDENS_SERVICO WHERE cod_ordem_servico in ("&cod_ordem_servico&")")
	if not rs.Eof then
		while not rs.Eof

			sql = "UPDATE TB_ORDENS_SERVICO SET despacho_transportador = "&sqlTexto(despacho_transportador)&" WHERE cod_ordem_servico = "&sqlNumero(rs("cod_ordem_servico"))

			objConexao.Execute(sql)

			sql = "INSERT INTO TB_ORDENS_SERVICO_HISTORICO (cod_ordem_servico_etapa, cod_ordem_servico, cod_usuario, data_cadastro, descricao) SELECT MAX(cod_ordem_servico_etapa), cod_ordem_servico, "&Session("cod_usuario")&", NOW(), 'A transportadora do pedido foi alterada para "&despacho_transportador&" motivo: "&justificativa&"' FROM TB_ORDENS_SERVICO_ETAPAS WHERE cod_ordem_servico = "&sqlNumero(rs("cod_ordem_servico"))

			objConexao.Execute(sql)

		rs.MoveNext
		Wend 
	end if
	rs.close
	set rs = nothing

	fecharConexao()
	response.redirect("lca_busca_detalhamento_pedido.asp?cod_pedido="&cod_pedido)
end if

%>
<h2>PPV - Despacho para Cliente | Alteração</h2>
<br />
<!-- #include file="includes/classe_modulo_ppv_menu.asp" -->
<br />
<br />

<%
if msg <> "" then
	%>
	<div class="msg">
		<%=msg%>
	</div>
	<%
end if
%>

<form action='<%=actionformulario%>' name='FORMULARIO_PADRAO' method='post'>
<%if campo1 = "" then%>
	<table width="600" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
	    <tr bgcolor="whitesmoke">
			<td>
				<h4>Informe o pedido de compras</h4>
			</td>
		</tr>
		<tr>
			<td>
				<div align="center">
					<input type="text" class="inputTamanho3 padrao_numero" name="cod_pedido" />
				</div>
			</td>
		</tr>
		<tr>
			<td>
				<div align="center">
					<a href="javascript:document.FORMULARIO_PADRAO.submit();">
						<img src="images/layout/btnGravar.jpg" />
					</a>
				</div>
			</td>
		</tr>
	</table>
<% elseif campo1 = "1" and cod_pedido <> "" then %>
	<table width="600" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
		
			<%
			abrirConexao()

			if validaUsuario(2,,80,true) then
				sql = "SELECT * FROM VW_TB_ORDENS_SERVICO WHERE cod_pedido = "&sqlNumero(cod_pedido)&"  AND status_ordem_servico = 3"
			else
				sql = "SELECT * FROM VW_TB_ORDENS_SERVICO WHERE cod_pedido = "&sqlNumero(cod_pedido)&" AND cod_ordem_servico not in (SELECT cod_ordem_servico FROM TB_PEDIDOS_ITENS_ENTREGUES) and status_ordem_servico = 3"
			end if

			Set rs = objConexao.Execute(sql)
			if not rs.Eof then
			%>
			<tr>
				<th>
					Formando:
				</th>
				<td colspan="3 ">
					<%=rs("nome_formando")%>
				</td>
			</tr>
			<tr>
				<th>
					Instituição:
				</th>
				<td colspan="3">
					<%=rs("entidade_ensino")%> &nbsp;&nbsp;&nbsp;Curso: <%=rs("nome_curso")%>
				</td>
			</tr>
			<tr>
				<th>
					Transportador:
				</th>
				<td colspan="3">
				<%=PadraoFormularioSelect("despacho_transportador","","inputTamanho3","","","SEM_SELECIONE",letra,"PADRAO:TRANSPORTADORA#TRANSPORTADORA|CORREIOS#CORREIOS|RETIRADA#RETIRADA|STILLOS#STILLOS|PANDEMIA#PANDEMIA|LOGISTICA#LOGISTICA|ESPECIAL#ESPECIAL", "incluir", false, false, false )%>
				</td>
			</tr>
			<tr>
				<th>
					Justificativa:
				</th>
				<td colspan="3">
					<textarea name="justificativa" class='inputTamanho3'></textarea>
				</td>
			</tr>
			<tr bgcolor="whitesmoke">
				<th width="110px">Ordem Serviço</th>
				<th width="60px">Quantidade</th>
				<th colspan="2">Envio / Produto</th>
			</tr>
			<%
				while not rs.Eof
					%>
					<tr>
						<td align="center"><%=moduloPPVOSMascara(rs("cod_ordem_servico"),"",true)%></td>
						<td align="center"><%=rs("quantidade")%></td>
						<td><%=rs("despacho_transportador")%> / <%=rs("nome_produto")%></td>
						<td align="center" width="20px"><input type="checkbox" name="cod_ordem_servico" value="<%=rs("cod_ordem_servico")%>" class='inputRadio'></td>
					</tr>
					<%
				rs.MoveNext
				Wend
				%>
				<input type="hidden" name="cod_pedido" value="<%=cod_pedido%>">
				<tr>
					<td colspan="4">
						<div align="center">
							<a href="javascript:document.FORMULARIO_PADRAO.submit();">
								<img src="images/layout/btnGravar.jpg" />
							</a>
						</div>
					</td>
				</tr>
						<%
			else
				%>
				<tr>
					<td>
						<div class="msgGrandeErro">
							Nenhuma OS Encontrada verifique se a OS já foi despachada ou entregue.
						</div>
					</td>
				</tr>
				<%
			end if
			rs.close
			set rs = nothing
			fecharConexao()
			%>
	</table>
<%end if%>
<br />
<br />
<input type="hidden" name="campo1" value="1" />
</form>
<!-- #include file="files/includes/miolo-inferior.asp" -->