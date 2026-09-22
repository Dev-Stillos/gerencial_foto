<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 1
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
  Dim titulo_pagina
  titulo_pagina = "Formandos Identificados"

  Dim calendario_necessario
  calendario_necessario = "NÃO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%
'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 10/10/2011 16:04:06

'Declara vari�veis para utiliza��o do sistema
Dim campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario, cod_pedido

Dim rsProdutosTotal, QTD_EM_PRODUTOS, XTOTAL_AGREGADOS, XTOTAL_ALBUNS, XTOTAL_PRODUTOS_A_PARTE, XTOTAL_DESCONTO, QTD_EM_PAGAMENTO

	QTD_EM_PRODUTOS = 0
	XTOTAL_AGREGADOS = 0
	XTOTAL_ALBUNS = 0
	XTOTAL_PRODUTOS_A_PARTE = 0
	XTOTAL_DESCONTO = 0
	QTD_EM_PAGAMENTO = 0

redirecionar_incluir	= "cob_pedidos_boletos_formulario.asp"
redirecionar_alterar	= "cob_pedidos_boletos.asp"
redirecionar_excluir	= "cob_pedidos_boletos.asp"
redirecionar_validacao	= "cob_pedidos_boletos.asp"
msgAoExcluir			= "Tem certeza que deseja excluir este formando?"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")

Dim nome_vendedor, data_venda, valor_pedido

Dim sacado_nome, sacado_cpf, sacado_email


If comando = Empty Then
comando = "incluir"
titulo_pagina = titulo_pagina&" / Incluir"
End if

If comando = "incluir" and campo1 <> Empty Then
	'Incluir novo registro
	incluirRegistro(redirecionar_incluir)
	titulo_pagina = titulo_pagina&" / inclusão"

ElseIf comando = "incluir" and campo1 = Empty Then
	'lemos o registro da lista de provaveis formandos	
	lerIdentificacao()
	titulo_pagina = titulo_pagina&" / inclusão"

Elseif comando = "alterar" and campo1 = Empty Then
	'Leitura inicial para possível Alteração
	cod_pedido = Request.QueryString("cod_pedido")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "alterar" and campo1 <> Empty Then
	'Gravar Alteração
	cod_pedido = Request.Form("cod_pedido")
	alterarRegistro(redirecionar_alterar)
	titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 
	'Leitura inicial para possível exclusão
	cod_pedido = Request.QueryString("cod_pedido")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
	'Exclui registro
	cod_pedido = Request.Form("cod_pedido")
	excluirRegistro(redirecionar_excluir)
	titulo_pagina = titulo_pagina&" / exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
	'Visualiza registro
	cod_pedido = Request.QueryString("cod_pedido")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / Visualização"
End If



Sub lerFormulario()

	

End sub


Sub lerRegistro()
	
	'buscamos dados do pedido
	dim rs
	AbrirConexao()
	Set rs=objConexao.Execute("SELECT cod_sacado,  nome_vendedor, data_venda FROM TB_PEDIDOS ped inner join TB_VENDEDORES vnd ON ped.cod_vendedor = vnd.cod_vendedor WHERE cod_pedido = "& cod_pedido)
	if not rs.eof then
		cod_sacado = rs("cod_sacado")
		nome_vendedor = rs("nome_vendedor")
		data_venda = rs("data_venda")
	''	valor_pedido = rs("valor_pedido")
	end if
	rs.Close
	Set rs=Nothing

	'buscamos dados do sacado
	Set rs=objConexao.Execute("SELECT * FROM TB_SACADOS WHERE cod_sacado = "&cod_sacado)
	if not rs.eof then
		sacado_nome = rs("nome")
		sacado_cpf = rs("cpf")
		sacado_email = rs("email")
		
	end if
	rs.Close
	Set rs=Nothing


	Set rsProdutosTotal = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_AGREGADOS WHERE COD_PEDIDO="&COD_PEDIDO)
	if not rsProdutosTotal.Eof Then
		XTOTAL_AGREGADOS = cdbl(0&rsProdutosTotal("XTOTAL_AGREGADOS"))
	End If
	rsProdutosTotal.Close()
	Set rsProdutosTotal = Nothing
	
	Set rsProdutosTotal = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_ALBUNS WHERE COD_PEDIDO="&COD_PEDIDO)
	if not rsProdutosTotal.Eof Then
		XTOTAL_ALBUNS = cdbl(0&rsProdutosTotal("XTOTAL_ALBUNS"))
	End If
	rsProdutosTotal.Close()
	Set rsProdutosTotal = Nothing
	
	Set rsProdutosTotal = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_PRODUTOS_A_PARTE WHERE COD_PEDIDO="&COD_PEDIDO)
	if not rsProdutosTotal.Eof Then
		XTOTAL_PRODUTOS_A_PARTE = cdbl(0&rsProdutosTotal("XTOTAL_PRODUTOS_A_PARTE"))
	End If
	rsProdutosTotal.Close()
	Set rsProdutosTotal = Nothing

	set rsProdutosTotal = objConexao.Execute("select SUM(valor_total_com_desconto) as valor_total_com_desconto from(SELECT SUM(faturado * (valor_un-(valor_un * (1-(ifnull(aliquota_desconto_foto,0)/100))))) as valor_total_com_desconto FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO in (SELECT cod_formando FROM TB_PEDIDOS_FORMANDOS WHERE cod_pedido =  "&cod_pedido&") union all SELECT SUM(quantidade * (valor_un-(valor_un * (1-(ifnull(aliquota,0)/100))))) as valor_total_com_desconto  FROM TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO = "&cod_pedido&") as total_desconto")
	if not rsProdutosTotal.Eof then
		XTOTAL_DESCONTO = cdbl(0&rsProdutosTotal(0))
	end if
	rsProdutosTotal.Close
	set rsProdutosTotal = Nothing
	QTD_EM_PRODUTOS = cdbl(XTOTAL_AGREGADOS + XTOTAL_ALBUNS + XTOTAL_PRODUTOS_A_PARTE - XTOTAL_DESCONTO)

	FecharConexao()

End sub

Sub incluirRegistro(redirecionar)

	lerFormulario()

	AbrirConexao()

	
	FecharConexao()
	Response.Redirect(redirecionar&"?comando=alterar&cod_pedido="&cod_pedido&"&codmsg=1&guia=2")
End sub

Sub alterarRegistro(redirecionar)
	
	lerFormulario()

	AbrirConexao()


	objConexao.execute(sql)

	FecharConexao()

	Response.Redirect(redirecionar&"?cod_pedido="&cod_pedido&"&codmsg=2")
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_pedido="&cod_pedido&"&codmsg=3")
End sub

function totaisFormaPagamento(cod_pedido, tipo, conexao_aberta)

	dim sql, retorno

	if not conexao_aberta then
		AbrirConexao()
	end if


	sql = "SELECT [CAMPO] FROM VW_PEDIDOS_TOTAL_RECEBIDO WHERE COD_PEDIDO = "&sqlNumero(cod_pedido)

	if tipo = "cheque" then
		sql = replace(sql, "[CAMPO]", "XTOTAL_CHEQUE")

	elseif tipo = "dinheiro" then
		sql = replace(sql, "[CAMPO]", "PAGAMENTO_EM_DINHEIRO")

	elseif tipo = "cartao" then
		sql = replace(sql, "[CAMPO]", "XTOTAL_CARTAO")

	elseif tipo = "boleto" then
		sql = replace(sql, "[CAMPO]", "XTOTAL_BOLETO")

	elseif tipo = "total" then
		sql = replace(sql, "[CAMPO]", "ifnull(XTOTAL_BOLETO,0) + ifnull(XTOTAL_CHEQUE,0) + ifnull(PAGAMENTO_EM_DINHEIRO,0) + ifnull(XTOTAL_CARTAO,0)")

	end if

	'response.write sql

	set rs = objConexao.execute(sql)
	if not rs.Eof then
		retorno = rs(0)
	end if
	rs.Close
	set rs = Nothing

	if not conexao_aberta then
		FecharConexao()
	end if

	totaisFormaPagamento = arrumanumero(0&retorno,"0,00",2)
end function

function totaisBoleto(cod_pedido, tipo, conexao_aberta)

	dim sql, retorno

	if not conexao_aberta then
		AbrirConexao()
	end if

	'avencer
	'vencidas
	'pagas
	'total

	if tipo = "pagas" then
		sql = "SELECT sum(valor_pago_efetivo) FROM TB_PEDIDOS_BOLETOS WHERE data_pagamento_efetivo is not null and cod_pedido = "&sqlNumero(cod_pedido)

	elseif tipo = "vencidas" then
		sql = "SELECT sum(valor) FROM TB_PEDIDOS_BOLETOS WHERE data_pagamento_efetivo is null AND data_vencimento < NOW() and cod_pedido = "&sqlNumero(cod_pedido)

	elseif tipo = "avencer" then
		sql = "SELECT sum(valor) FROM TB_PEDIDOS_BOLETOS WHERE data_pagamento_efetivo is null AND data_vencimento > NOW() and cod_pedido = "&sqlNumero(cod_pedido)

	elseif tipo = "total" then
		sql = "SELECT sum(valor) FROM TB_PEDIDOS_BOLETOS WHERE cod_pedido = "&sqlNumero(cod_pedido)

	end if

	set rs = objConexao.execute(sql)
	if not rs.Eof then
		retorno = rs(0)
	end if
	rs.Close
	set rs = Nothing

	if not conexao_aberta then
		FecharConexao()
	end if

	totaisBoleto = retorno
end function

'----------------------------------------------------------------------------------------
%>
<style>
.bloqueio_tela{
	width: 100%;
	height: 100%;
	background-color: rgba(0,0,0,0.3	);
	position: fixed;
	top: 0px;
	left: 0px;
	z-index: 1000;
	background-position:  center center;
	background-repeat: no-repeat;
	background-image: url( images/layout/preloader_grande.gif );
}
/* body.bloquedo{ overflow: hidden; } */
</style>



<h2><%=titulo_pagina%></h2><br />
<!-- #include file="includes/sub_menu_cob.asp" -->
<br />
<%If comando = "excluir" Then%>
	<div class='msgExcluir'><%=msgAoExcluir%></div>
<%End If%>

<style type="text/css">
	fieldset{
		width: 820px;
		margin:10px auto 5px auto ;
	}
</style>
<table align="center" width='800px' >
	<tr>
		<td>
			<!--<button class="btn btn-sm btn-primary">Enviar Cobran�a</button> -->
		</td>
	</tr>
</table>


<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO'  method='post'>
	<fieldset>
		<legend>Pedido</legend>
		<table width='800px' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
			<tr>
				<td width='27%'><strong>Código pedido</strong></td>
				<td width='73%'>
					<div align='left'>
						<%=cod_pedido%>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>Vendedor *</strong></td>
				<td width='73%'>
					<div align='left'>
						<%=nome_vendedor%>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>Data venda *</strong></td>
				<td width='73%'>
					<div align='left'>
						<%=arrumaData(data_venda)%>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>Valor Pedido*</strong></td>
				<td width='73%'>
					<div align='left'>
						R$ <%=arrumanumero(QTD_EM_PRODUTOS, "0,00", 2)%>
					
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>formas pagamento*</strong></td>
				<td width='73%'>
					<div align='left'>
						Dinheiro: R$ <%=totaisFormaPagamento(cod_pedido, "dinheiro", false)%> <br> 
						Cartão: R$ <%=totaisFormaPagamento(cod_pedido, "cartao", false)%> <br>
						Boleto: R$ <%=totaisFormaPagamento(cod_pedido, "boleto", false)%> <br>
						Cheque: R$ <%=totaisFormaPagamento(cod_pedido, "cartao", false)%> 
					</div>
				</td>
			</tr>
		</table>
	</fieldset>

	<fieldset>
		<legend>Sacado</legend>
		<table width='800px' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
			<tr>
				<td width='27%'><strong>Nome: *</strong></td>
				<td width='73%'>
					<div align='left'>
						<%=sacado_nome%>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>CPF: *</strong></td>
				<td width='73%'>
					<div align='left'>
						<%=sacado_CPF%>
					</div>
				</td>
			</tr>
			<tr>
				<td width='27%'><strong>E-Mail: *</strong></td>
				<td width='73%'>
					<div align='left'>
						<%=sacado_email%>
					</div>
				</td>
			</tr>			
		</table>
	</fieldset>
	<fieldset>
		<legend>Boletos</legend>
	<button class="btn btn-sm btn-primary" id='reenviarboletos'>Enviar boletos</button>
		<div class="subMenu" style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC; margin-bottom:15px; width: 500px; margin-left:auto; margin-right:auto;">
			<table cellpadding="4" align="center" id='form_adicionar_parcela'>
				<tr>
					<th>Parcela</th>
					<th>Vencimento</th>
					<th>Valor</th>
					<th></th>
				</tr>
				<tr>
					<td>
						<input type="text" id="parcela_id" style="text-align:center; width:50px;  height: 27px; " value="" maxlength="2">
					</td>
					<td>
						<input type="text" class="padrao_data" id="vencimento_id" style="text-align:center;  width:75px; height: 27px; ">
					</td>
					<td>
						<input type="text" class="padrao_valor" id="valor_id" style="text-align:right;   width:120px;  height: 27px; ">
					</td>
					<td align="center">
						<input type="button" class='btn btn-sm btn-success' value='Adicionar' id='botao_adicionar'>
						<input type="button" class='btn btn-sm btn-danger' value='Cancelar' id='botao_cancelar'>
					</td>
				</tr>
			</table>
		
			<table cellpadding="2" align="center" class='hidden' id='form_baixa_manual'>
				<tr>
					<td colspan="2"><h4>Baixa de Parcela Manual</h4></td>
				</tr>
				<tr>
					<td ><strong>Parcela:</strong> </td> <td><span id="parcela_baixa"></span></td> 
				</tr>
				<tr>
					<td ><strong>Vencimento:</strong> </td> <td> <span id="vencimento_baixa"></span></td>
				</tr>
				<tr>
					<td ><strong>Valor R$</strong> </td> <td> <span id="valor_baixa"></span></td>
				</tr>
				<tr>
					<td ><strong>Valor Pago:</strong> </td> <td> <input type="text" id='baixar_txt_valor_pago' style="width:100px; text-align:right; " class="padrao_valor" /></td>
				</tr>
				<tr>
					<td ><strong>Data Pagamento:</strong> </td> <td> <input type="text"  id='baixar_txt_data_pagamento' style="width:100px; text-align:center; " class='padrao_data' /></td>
				</tr>
				<tr>
					<td colspan="2">
						<strong>Motivo:</strong>
					</td>
				</tr>	
				<tr>
					<td colspan="2">
						<strong>
							<textarea name="" id="" cols="26" rows="5" id='baixar_txt_motivo'></textarea>
						</strong>
					</td>
				</tr>	
				<tr><td colspan="2">&nbsp; <input type="hidden" id="baixar_txt_cod_boleto">	</td></tr>
				<tr>
					<td colspan="2" align="center">
						<input type="button" class='btn btn-sm btn-success' value='Baixar' id='confirma_baixar_boleto'>
						<input type="button" class='btn btn-sm btn-danger' value='Cancelar' id='cancelar_baixar_boleto'>
					</td>
				</tr>
			
			</table>
		</div>

		<table width='800px' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
		<tr bgcolor="whitesmoke">
			<th>#</th>
			<th>Vencimento</th>
			<th>Pagamento</th>
			<th>Baixa</th>
			<th>Valor</th>
			<th>Valor Pago</th>
			<th>Valor Comissão</th>
			<th width="64px"> 
			<a href="javascript:void(0)" class="marcar_todos">Marcar todos</a>
			</th>
		</tr>
		
		<%
		AbrirConexao()

		set rs = objConexao.Execute("SELECT * FROM TB_PEDIDOS_BOLETOS WHERE cod_pedido = "&sqlNumero(cod_pedido)&" order by parcela")
		while not rs.Eof
			%>
				<tr>
					<td align="center">
						<input type="hidden" id='parcela_<%=rs("cod_boleto")%>' value='<%=rs("parcela")%>'>
						<%=rs("parcela")%>
					</td>
					<td align="center">
						<input type="hidden" id='vencimento_<%=rs("cod_boleto")%>' value='<%=arrumadata(rs("data_vencimento"))%>'>
						<%=arrumadata(rs("data_vencimento"))%>
					</td>
					<td align="center">
						<%=arrumaData(rs("data_pagamento_efetivo"))%>
					</td>
					<td align="center">
						<%=arrumaData(rs("data_baixa"))%>
					</td>
					<td align="right">
						<input type="hidden" id='valor_<%=rs("cod_boleto")%>' value='<%=arrumanumero(rs("valor"),"0,00",2)%>'>
						R$ <%=arrumanumero(rs("valor"),"0,00",2)%>
					</td>
						<td align="right">R$ <%=arrumanumero(rs("valor_pago_efetivo"), "0,00",2)%></td>
						<td align="right">R$ <%=arrumanumero(rs("valor_comissao_vendedor"), "0,00",2)%></td>
					<td>
					<% if ""&rs("data_pagamento_efetivo") = "" then%>
						<input type="checkbox" value='<%=rs("cod_boleto")%>' class="cod_boleto_enviar">

						<img src="images/icons/icon_excluir.png" style="cursor: pointer;" alt="Excluir registro" class= 'excluir_boleto' cod-pedido='<%=rs("cod_pedido")%>' cod-boleto='<%=rs("cod_boleto")%>'>

						<img src="images/icons/baixa_ativo.png" width="16px" height="16px" style="cursor: pointer;" alt="Excluir registro" cod-pedido='<%=rs("cod_pedido")%>' cod-boleto='<%=rs("cod_boleto")%>' class='baixar_boleto' 	>
					<%end if%>
					</td>
				</tr>
			<%
			rs.movenext	
		Wend
		rs.Close
		set rs = Nothing

		FecharConexao()
		%>
		</table>
		<br>
			
		<table width='800px'>
			<tr>
				<td width="50%">
					<h2 style='color:#009933 !important;'>Total Pago</h2>
				</td>
				<td width="50%" align="right">
					<h2 style='color:#009933 !important;'>R$ <%=arrumaNumero(totaisBoleto(cod_pedido, "pagas", false),"0,00",2)%></h2>
				</td>
			</tr>
			<tr>
				<td>
					<h2 style='color:#cc0000 !important;'>Total Vencido</h2>
				</td>
				<td width="50%" align="right">
					<h2 style='color:#cc0000 !important;'>R$ <%=arrumaNumero(totaisBoleto(cod_pedido, "vencidas", false),"0,00",2)%></h2>
				</td>
			</tr>
			<tr>
				<td>
					<h2 style='color:#0066cc !important;'>Total A vencer</h2>
				</td>
				<td width="50%" align="right">
					<h2 style='color:#0066cc !important;'>R$ <%=arrumaNumero(totaisBoleto(cod_pedido, "avencer", false),"0,00",2)%></h2>
				</td>
			</tr>
			<tr>
				<td>
					<h2 style='color:#000 !important;'>Total</h2>
				</td>
				<td width="50%" align="right">
					<h2 style='color:#000 !important;'>R$ <%= arrumaNumero(totaisBoleto(cod_pedido, "total", false),"0,00",2)%></h2>
				</td>
			</tr>
		</table>

	</fieldset>
	<input type='hidden' value='1' name='campo1' />
	<input type='hidden' id='cod_pedido_id' name='cod_pedido' value="<%=cod_pedido%>" />
</form>
<script type="text/javascript">
		function bloqueioTela(){
			$(".bloqueio_tela").removeClass("hidden");
			$("body").addClass("bloquedo");
		}

		function desbloqueioTela(){
			$(".bloqueio_tela").addClass("hidden");
			$("body").removeClass("bloquedo");
		}
	$(function(){

		var parcelas = {
			url : "ajax_boletos.asp",
			assincrono : true,



			incluir : function(cod_pedido, parcela, vencimento, valor)
			{
				var request = $.ajax({
					url: this.url,
					type: "POST",
					async: this.assincrono,
					data: { 
						tp : "incluirBoleto", 
						p0 : cod_pedido,
						p1 : parcela,
						p2 : vencimento,
						p3 : valor
					},
					beforeSend : function () {bloqueioTela();}, 
					dataType: "text"
				});
				request.done(function( msg ) {
					alert(msg.split("|")[1]);
				});
				request.fail(function( jqXHR, textStatus ) {
					alert(textStatus);
				});
				request.always(function(){
					desbloqueioTela();
					location.reload();	
				});
			},
			excluir : function(cod_boleto)
			{
				var request = $.ajax({
					url: this.url,
					type: "POST",
					async: this.assincrono,
					data: { 
						tp : "excluirBoleto", 
						p0 : cod_boleto
					},
					beforeSend : function () {bloqueioTela();}, 
					dataType: "text"
				});
				request.done(function( msg ) {
					alert( msg.split("|")[1]);
				});
				request.fail(function( jqXHR, textStatus ) {
					alert(textStatus);
				});
				request.always(function(){
					desbloqueioTela();
					location.reload();	
				});
			},
			baixar : function(cod_boleto, valor_pago, data_pagamento, motivo)
			{
				var request = $.ajax({
					url: this.url,
					type: "POST",
					async: this.assincrono,
					data: { 
						tp : "baixarBoleto", 
						p0 : cod_boleto,
						p1 : valor_pago,
						p2 : data_pagamento,
						p3 : motivo,
					},
					beforeSend : function () {bloqueioTela();}, 
					dataType: "text"
				});
				request.done(function( msg ) {
					alert(msg);
				});
				request.fail(function( jqXHR, textStatus ) {
					alert(textStatus);
				});
				request.always(function(){
					desbloqueioTela();
					location.reload();	
				});
			}, 
			enviarBoletos : function (cod_boletos)
			{
				var request = $.ajax({
					url: "gerarBoletoPdf.php",
					type: "GET",
					async: this.assincrono,
					data: { 
						tipo : "boletos", 
						cod : cod_boletos,
					},
					beforeSend : function () {bloqueioTela();}, 
					dataType: "JSON"
				});
				request.done(function( msg ) {
					console.log(msg);
					alert(msg.msg);
				});
				request.fail(function( jqXHR, textStatus ) {
					alert(textStatus);
				});
				request.always(function(){
					desbloqueioTela();
				});
			}
		}

		$(".marcar_todos").click(function(){
			$('input[type=checkbox]').attr('checked', true);
		});

		$("#botao_adicionar").click(function(){
			var cod_pedido = $("#cod_pedido_id").val();
			var parcela = $("#parcela_id").val();
			var vencimento = $("#vencimento_id").val();
			var valor = $("#valor_id").val();

 			parcelas.incluir(cod_pedido, parcela, vencimento, valor);
		});

		$("#botao_cancelar").click(function(){
			$("#vencimento_id").val("");
			$("#parcela_id").val("");
			$("#valor_id").val("");
		});

		$(".excluir_boleto").click(function(){
			if (confirm("Tem certeza que deseja excluir esta parcela?"))
				parcelas.excluir($(this).attr("cod-boleto"));
		});

		$(".baixar_boleto").click(function(){
			var cod_boleto = $(this).attr("cod-boleto");
			var parcela = $("#parcela_"+cod_boleto).val() ;
			var vencimento = $("#vencimento_"+cod_boleto).val();
			var valor = $("#valor_"+cod_boleto).val();

			$("#parcela_baixa").text(parcela);
			$("#vencimento_baixa").text(vencimento);
			$("#valor_baixa").text(valor);
			$("#baixar_txt_cod_boleto").val(cod_boleto);

			$("#form_baixa_manual").removeClass("hidden");
			$("#form_adicionar_parcela").addClass("hidden");
		});

		$("#cancelar_baixar_boleto").click(function(){
			$("#form_baixa_manual").addClass("hidden");
			$("#form_adicionar_parcela").removeClass("hidden");
		});

		$("#reenviarboletos").click(function(event ){
			event.preventDefault();

			var a ='';
			$(".cod_boleto_enviar:checked").each(function(){
				a += $(this).val()+',';
			});
			a = a.substr(0, a.length - 1);			
			
			if (a.length > 0) 
				parcelas.enviarBoletos(a);
			else
				alert('Nenhum boleto Selecionado');

		});

		$("#confirma_baixar_boleto").click(function(){

			var cod_boleto = $("#baixar_txt_cod_boleto").val(); 
			var motivo = $("#baixar_txt_motivo").val(); 
			var data_pagamento = $("#baixar_txt_data_pagamento").val(); 
			var valor_pago = $("#baixar_txt_valor_pago").val(); 
			if (confirm("Tem certeza que deseja baixar esta parcela?"))
				parcelas.baixar(cod_boleto, valor_pago, data_pagamento, motivo);
		});	
	});
</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->