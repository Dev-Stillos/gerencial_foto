<%
'data in�cio formatado, data t�rmino, t�tulo da p�gina
'call manutencao(ArrumaData(date(),"")& "�s 08h00",ArrumaData(date(),"")& "�s 12h00","")

Dim calendario_necessario, BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
calendario_necessario = "SIM"
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 1
BTN_IMPRIMIR_LINK = "#"
titulo_pagina = "Compras - Formulário"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<%

'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 17/12/2011 20:14:45

'Declara variáveis para utilização do sistema
Dim cod_pedido_compra, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

redirecionar_incluir	= "ppv_compras_pedidos.asp"
redirecionar_alterar	= "ppv_compras_pedidos.asp"
redirecionar_excluir	= "ppv_compras_pedidos.asp"
redirecionar_validacao	= "ppv_compras_pedidos.asp"
msgAoExcluir			= "Tem certeza que deseja excluir este Pedido de Compra?"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
'Declara variáveis
Dim numero_pedido_compra, cod_empresa, cod_fornecedor, frete_por_conta, frete_por_conta_acrescentar_valor_total, frete, entrega_ou_retira, cod_forma_pagamento, cod_forma_recebimento, contato, obs, cadastro_cod_usuario, data_aprovado, aprovado_cod_usuario, data_enviado_fornecedor, enviado_cod_usuario, status_pedido_compra, cancelado_motivo, cod_usuario, modificado

'Secundárias
Dim razao_social 'Do fornecedor usado somente para leitura
Dim valor_total_produtos
valor_total_produtos = 0

If comando = Empty Then
comando = "incluir"
titulo_pagina = titulo_pagina&" / Inclusão"
End if

'#########################################################################
'Valida��o de sub-lida��es padr�es com formul�rio
'#########################################################################
If comando = "incluir" and campo1 = Empty Then
	AbrirConexao()
		'1) Cria uma chave do pedido de compra para inserir nas guias adicionais
		'Obs NÃO cria um n�mero de pedido de compra que será gerado somente ap�s a grava��o e finaliza��o da inserá�o do pedido de compra completo
		objConexao.Execute("INSERT INTO TB_PEDIDOS_COMPRAS (cod_usuario,cadastro_cod_usuario, pedido_manual) VALUES ("&sqlNumero(Session("cod_usuario"))&","&sqlNumero(Session("cod_usuario"))&", 'S')")
		
		'Encontrar chave criada
		Set rs=objConexao.Execute("SELECT @@IDENTITY AS 'ULTIMO_ID'")
			If not rs.Eof Then
				'Continua o lançamento, mas agora já com chave criada, sendo que o incluir na verdade será o alterar.
				'Obs.: NÃO redirecionei com o comando alterar, pois neste caso quem tiver acesso a incluir e NÃO tiver a alterar NÃO conseguiria nem incluir o pedido, pois o comando já seria trocado para Alteração
				cod_pedido_compra = rs("ULTIMO_ID")
			Else
				Response.Redirect("inicio.asp?msg=Desculpe houve uma falha técnica na geração do pedido de compra, se o problema persistir informe o Administrador do sistema.")
			End if 
		rs.Close
		Set rs=Nothing
			
		'Por enquanto o status � 0 (zero), ou seja em cadastro.
		status_pedido_compra = 0
		
	FecharConexao()
ElseIf comando = "incluir" and campo1 <> Empty Then
'Incluir novo registro
incluirRegistro(redirecionar_incluir)
titulo_pagina = titulo_pagina&" / Inclusão"
Elseif comando = "alterar" and campo1 = Empty Then
'Leitura inicial para poss�vel Alteração
cod_pedido_compra = Request.QueryString("cod_pedido_compra")
lerRegistro()
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "alterar" and campo1 <> Empty Then
'Gravar Alteração
cod_pedido_compra = Request.Form("cod_pedido_compra")
alterarRegistro(redirecionar_alterar)
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 
'Leitura inicial para poss�vel Exclusão
cod_pedido_compra = Request.QueryString("cod_pedido_compra")
lerRegistro()
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
'Exclui registro
cod_pedido_compra = Request.Form("cod_pedido_compra")
excluirRegistro(redirecionar_excluir)
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
'Visualiza registro
cod_pedido_compra = Request.QueryString("cod_pedido_compra")
lerRegistro()
titulo_pagina = titulo_pagina&" / Visualização"
End If


'Fun��es

Sub lerFormulario()

cod_pedido_compra = Request.Form("cod_pedido_compra")
numero_pedido_compra = Request.Form("numero_pedido_compra")
cod_empresa = Request.Form("cod_empresa")
cod_fornecedor = Request.Form("cod_fornecedor")
frete_por_conta = Request.Form("frete_por_conta")
frete_por_conta_acrescentar_valor_total = Request.Form("frete_por_conta_acrescentar_valor_total")
frete = Request.Form("frete")
entrega_ou_retira = Request.Form("entrega_ou_retira")
cod_forma_pagamento = Request.Form("cod_forma_pagamento")
cod_forma_recebimento = Request.Form("cod_forma_recebimento")
contato = Request.Form("contato")
obs = Request.Form("obs")
data_cadastro = Request.Form("data_cadastro")
cadastro_cod_usuario = Request.Form("cadastro_cod_usuario")
data_aprovado = Request.Form("data_aprovado")
aprovado_cod_usuario = Request.Form("aprovado_cod_usuario")
data_enviado_fornecedor = Request.Form("data_enviado_fornecedor")
enviado_cod_usuario = Request.Form("enviado_cod_usuario")
status_pedido_compra = Request.Form("status_pedido_compra")
cancelado_motivo = Request.Form("cancelado_motivo")
cod_usuario = Session("cod_usuario")
modificado = now()

End sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM VW_TB_PEDIDOS_COMPRAS WHERE cod_pedido_compra = "&cod_pedido_compra)

			numero_pedido_compra=trim(rs("numero_pedido_compra"))
'				cod_empresa=trim(rs("cod_empresa"))
			cod_fornecedor=trim(rs("cod_fornecedor"))
			frete_por_conta=trim(rs("frete_por_conta"))
			frete_por_conta_acrescentar_valor_total=trim(rs("frete_por_conta_acrescentar_valor_total"))
			frete= ArrumaNumero(rs("frete"),"",2)
			entrega_ou_retira=trim(rs("entrega_ou_retira"))
			cod_forma_pagamento=trim(rs("cod_forma_pagamento"))
			cod_forma_recebimento=trim(rs("cod_forma_recebimento"))
			contato=trim(rs("contato"))
			obs=trim(rs("obs"))
			data_cadastro=trim(rs("data_cadastro"))
			cadastro_cod_usuario=trim(rs("cadastro_cod_usuario"))
			data_aprovado=trim(rs("data_aprovado"))
			aprovado_cod_usuario=trim(rs("aprovado_cod_usuario"))
			data_enviado_fornecedor=trim(rs("data_enviado_fornecedor"))
			enviado_cod_usuario=trim(rs("enviado_cod_usuario"))
			status_pedido_compra=trim(rs("status_pedido_compra"))
			cancelado_motivo=trim(rs("cancelado_motivo"))
			cod_usuario=trim(rs("cod_usuario"))
			modificado=trim(rs("modificado"))

			'//Busca o valor total em produtos solicitados
			valor_total_produtos = pedidoCompraTotais(cod_pedido_compra,1,true)
				
			'//Denomina��o do fornecedor
			Set rs2=objConexao.Execute("SELECT razao_social FROM TB_FORNECEDORES WHERE cod_fornecedor = '"&cod_fornecedor&"'")
				
				If not rs2.Eof Then
					razao_social = rs2("razao_social")
				End If
			
			rs2.Close
			Set rs2=Nothing
				
				
		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
				
		'////////////////////////////
		'No momento da Inclusão verifica se o usu�rio possui valida��o para aprova��o do pedido de compra se tiver o status do pedido ao inv�s de 1 j� pula para 2, pois o pr�prio usu�rio que possui acesso a aprova��o est� cadastrando um pedido de compra
		'If validaUsuario(2,3,true) Then
		'Possui valida��o
		''	status_pedido_compra = 2 'Aprovado, mas NÃO enviado para fornecedor
			
			'Chama a fun��o que faz o update para mudar o status do pedido de compra
		''	Dim mudanca_status_realiza 'neste caso será realizada sem precisar exibir mensagem de erro caso NÃO ocorra
		''	mudanca_status_realiza = pedidoCompraMudancaStatus(cod_pedido_compra,2,"",true)
			
		'Else
		'NÃO possui
			status_pedido_compra = 1 'Aguardando aprova��o
		'End If
		'////////////////////////////
		
		
		'ATENÇÃO NÃO � necess�rio fazer um where somente dos pedidos de compra com status superiores a 1 (Em cadastro) pois se estiver com este status NÃO h� numero de pedido de compra gerado e por tanto NÃO ir� contar.
		Set rs=objConexao.Execute("SELECT max(numero_pedido_compra) + 1 as novo FROM TB_PEDIDOS_COMPRAS ")
			
			If not rs.Eof Then
				numero_pedido_compra = rs(0)
			End If
			
		rs.Close
		Set rs=Nothing	
		
		sql = "UPDATE TB_PEDIDOS_COMPRAS SET" &_ 
					" numero_pedido_compra = "&sqlNumero(numero_pedido_compra)&_
					", cod_fornecedor = "&sqlNumero(cod_fornecedor)&_
					", obs = "&sqlTexto(obs)&_
					", cadastro_cod_usuario = "&sqlTexto(cod_usuario)&_
					", cod_usuario = "&sqlNumero(cod_usuario)&_
					", modificado = "&sqlData(modificado)&_
					", status_pedido_compra = '"&status_pedido_compra&_
					"' WHERE cod_pedido_compra = "&cod_pedido_compra

		'response.write sql
		'Na Inclusão na realidade faz um update e gera o n�mere do pedido de compra
		objConexao.Execute(sql)

	FecharConexao()
	Response.Redirect(redirecionar&"?cod_pedido_compra="&cod_pedido_compra&"&codmsg=1")
End sub

Sub alterarRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		'NÃO altera o n�mero de pedido e nem o status pelo UPDATE
		objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS SET cod_fornecedor = "&sqlNumero(cod_fornecedor)&", frete = "&sqlNumero(frete)&" , cod_forma_pagamento = "&sqlTexto(cod_forma_pagamento)&", cod_forma_recebimento = "&sqlTexto(cod_forma_recebimento)&", contato = "&sqlTexto(contato)&", obs = "&sqlTexto(obs)&", cod_usuario = "&sqlNumero(cod_usuario)&", modificado = "&sqlData(modificado)&" WHERE cod_pedido_compra = "&cod_pedido_compra)
	
	
		'Verifica o status_pedido_compra do pedido de compra
		Set rs=objConexao.Execute("SELECT status_pedido_compra FROM TB_PEDIDOS_COMPRAS where cod_pedido_compra = "&cod_pedido_compra)
				If not rs.Eof Then
					status_pedido_compra = int(rs(0))
				End If
		rs.Close
		Set rs=Nothing
	
	FecharConexao()	
	Response.Redirect(redirecionar&"?cod_pedido_compra="&cod_pedido_compra&"&codmsg=2")
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()
		'call verificaAntesExcluir(,cod_pedido_compra,"TABELASRELACIONADAS;TABELASRELACIONADAS","cod_pedido_compra",redirecionar_excluir&"?sucesso=2&codmsg=5")

	AbrirConexao()
	
		'Volta os status das etapas como pendentes deste pedido de compra
		objConexao.Execute("UPDATE TB_ORDENS_SERVICO_ETAPAS SET pedido_compra_gerado = NULL where pedido_compra_gerado = "&cod_pedido_compra)
		objConexao.Execute("DELETE FROM TB_PEDIDOS_COMPRAS WHERE cod_pedido_compra = "&cod_pedido_compra)
		objConexao.Execute("DELETE FROM TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra = "&cod_pedido_compra)
		
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_pedido_compra="&cod_pedido_compra&"&codmsg=3")
End sub

'----------------------------------------------------------------------------------------

%>
<script language="javascript" type="text/javascript">
function atualizaValor(div,tipo)
{
	var d = document.FORMULARIO_PADRAO
	var d2 = document
	if(tipo == 1)
	{
		var caminho = "compras_pedidos_formulario_consulta_ajax.asp?tipo=altualizaValor&valor6="+d.produto_pedido_compras_d.value+'&valor2='+d.produto_pedido_compras_c.value+'&valor3='+d.produto_pedido_compras_e.value+'&valor4='+d.produto_pedido_compras_e2.value+'&valor5='+d.produto_pedido_compras_e3.value+'&valor7=itensComPedidoValorUnit';
	}
	else
	{
		var caminho = "compras_pedidos_formulario_consulta_ajax.asp?tipo=altualizaValor&valor6="+d2.getElementById('valor_custo_id').value+'&valor2='+d2.getElementById('retorna_quantidade_id').value+'&valor3='+d2.getElementById('aliquota_ipi_id').value+'&valor4='+d2.getElementById('aliquota_substituicao_tributaria_id').value+'&valor5='+d2.getElementById('percentual_desconto_id').value+'&valor7=itensSemPedidoValorUnit';
	}
	//alert(caminho);
	ajax_loadContent(div,caminho,false,false,false);	
}
</script>

<script type="text/javascript">
function pedidoCompraAtualizaValorTotalProdutos(cod_pedido_compra){
	var caminho = 'compras_pedidos_formulario_consulta_ajax.asp?tipo=vtp&valor='+cod_pedido_compra+'&valor2=<%=status_pedido_compra%>';
	ajax_loadContent('ajax_VALOR_TOTAL_PRODUTOS',caminho,false,false,false);	
}

function listarContatos (campo){
	
	var temp = "";
	temp = document.FORMULARIO_PADRAO[campo].value;
	
	if(campo == 'COD_FORNECEDOR'){
		if(temp == ''){
			alert('Informe o código do fornecedor para buscar os contatos cadastrados do mesmo!');
		}else{
			var caminho = "buscaParaCampo.asp?tipo=CONTATO_FORNECEDOR_"+temp+"&campo=CONTATO&formulario=FORMULARIO_PADRAO&st=S";
			//alert(caminho);
			displayMessage(caminho);				
		}
	}
}


function pedidoCompraConfirmaMudancaStatus(a,b,c){
<%
''a = cod_pedido_compra com base encode 64
''b = status atual
''c = novo status
%>
	var msg = "", caminho = "", m = "";
	var d = document.FORMULARIO_PADRAO;
	
	if(c==2){
		msg = "Tem certeza que deseja mudar o status para Aprovado?";	
	}else if(c==3){
		msg = "Tem certeza que deseja mudar o status para Enviado para fornecedor?\nObs.: NÃO esqueça que você deverá enviar o pedido de compra via e-mail para o fornecedor.";		
	}else if(c==99){		
		msg = "ATENÇÃO PARA CANCELAMENTO DE PEDIDO DE COMPRA!\nTem certeza que deseja cancelar o Pedido de Compra e seus respectivos itens?\nObs.: Se houver compras relacionadas com este pedido de compras as mesmas NÃO serão canceladas.";
		m = d.cancelado_motivo.value;
	}else if(c==22){
		msg = "Tem certeza que deseja voltar o Status para NÃO Enviado para fornecedor?\nObs.: Desta forma você poderá alterar o Pedido de Compra novamente, mas posteriormente deverá enviar novamente para o fornecedor.";	
	}
	
	if (msg !== "") {
		if(YNconfirm(msg)){
			caminho = 'compras_pedidos_formulario_consulta_ajax.asp?tipo=ms&valor='+a+'&valor2='+b+'&valor3='+c+'&valor4='+m;
			ajax_processa(caminho,true);
		}
	}
}

function validaCancelamento(){
	var d = document.FORMULARIO_PADRAO.cancelado_motivo;
	if(d.value !== ""){
		pedidoCompraConfirmaMudancaStatus('<%=base64Encode(cod_pedido_compra)%>','<%=status_pedido_compra%>',99);
	}else{
		alert('Preencha o motivo do cancelamento antes de gravar!');
		d.focus();
	}

}

function consultaUltimoPreco(){
	var cod_produto = document.FORMULARIO_PADRAO.COD_PRODUTO_combo_1.value;
	if(cod_produto){
			
		var caminho = 'compras_pedidos_formulario_consulta_ajax.asp?tipo=cuf&valor=<%=cod_pedido_compra%>&valor2='+cod_produto;	

		document.getElementById("containerConsultaUltimoFornecedor").innerHTML = "<br /><strong>Aguarde carregando informações da última compra deste produto</strong>"
		ajax_loadContent('containerConsultaUltimoFornecedor',caminho,false,false,false);	
	}
}

var envio_em_andamento;

$(function(){

	$("#aprovar_pedido").click(function(event){
 		event.preventDefault();

		var campos = [
					{name:"ts", value:"aprovar_pedido"},
					{name:"valor", value:$("input[name=COD_PEDIDO_COMPRA]").val()}
				];

		var msg = "ATENÇÃO \n\n você deseja aprovar este pedido de compra, após esta ação NÃO será mais possível voltar."

		if (!envio_em_andamento && confirm(msg))
		{
			var request = $.ajax({
				url: "ppv_compras_pedidos_ajax.asp",
				method: "GET",
				data: campos,
				dataType: "html",
				beforeSend: function(){
					envio_em_andamento = true;
				}
			});

			request.done(function(msg) {
	            alert("Pedido aprovado com sucesso");
	            location.reload();
	            envio_em_andamento = false;
	        });
	         
	        request.fail(function( jqXHR, textStatus ) {
				alert("ATENÇÃO \n houve um erro na aprovação do pedido, informe o dpto de TI");
				envio_em_andamento = false;
	        });
		}
	});

	$("#finaliza_pedido").click(function(event){
		event.preventDefault();
		var campos = [
					{name:"ts", value:"finalizar_pedido"},
					{name:"valor", value:$("input[name=COD_PEDIDO_COMPRA]").val()}
				];

		var msg = "ATENÇÃO \n\n você deseja finalizar este pedido de compra, todos os itens serão marcados como entregues e  esta ação NÃO poderá ser desfeita."

		if (!envio_em_andamento && confirm(msg))
		{
			var request = $.ajax({
				url: "ppv_compras_pedidos_ajax.asp",
				method: "GET",
				data: campos,
				dataType: "html",
				beforeSend: function(){
					envio_em_andamento = true;
				},
				error: function(jqXHR, textStatus){
					alert("ATENÇÃO \n houve um erro na finalizaçãoo do pedido, informe o dpto de TI (caue) " + textStatus );
				}
			});

			request.done(function(msg) {
	            alert("Pedido aprovado com sucesso");
	            location.reload();
	            envio_em_andamento = false;
	        });
	         
	        request.fail(function( jqXHR, textStatus ) {
				//alert("ATENÇÃO \n houve um erro na finalização do pedido, informe o dpto de TI " + textStatus );
				location.reload();
				envio_em_andamento = false;
	        });
		}
	});

	$("#gerar_recibo").click(function(event){
		event.preventDefault();
		var campos = [
					{name:"ts", value:"gerar_recibo"},
					{name:"valor", value:$("input[name=COD_PEDIDO_COMPRA]").val()}
				];

		if (!envio_em_andamento)
		{
			var request = $.ajax({
				url: "ppv_compras_pedidos_ajax.asp",
				method: "GET",
				data: campos,
				dataType: "html",
				beforeSend: function(){
					envio_em_andamento = true;
				}
			});

			request.done(function(msg) {
	            envio_em_andamento = false;
	            window.open("ppv_compras_pedidos_produtos_solicitados_recibo_entrega.asp",'_blank');
	        });
	         
	        request.fail(function( jqXHR, textStatus ) {
				envio_em_andamento = false;
				location.reload();
	        });
		}
	});

});


</script>

<h2><%=titulo_pagina%></h2><br />
<!-- #include file="includes/compras_sub_menu.asp" -->
<%If comando = "excluir" Then%>
<div class='msgExcluir'><%=msgAoExcluir%></div>
<%End If%>

<%		
Dim qtd_tabs : qtd_tabs = 2
%>
<div id="tabsJ">
	<ul>
		<li id="liTab1" class="current"><a href="javascript:tabsMostra('tab1',<%=qtd_tabs%>,1);"><span>Dados do Pedido de Compra</span></a></li>
		<li id="liTab2"><a href="javascript:tabsMostra('tab2',<%=qtd_tabs%>,2);"><span>Produtos</span></a></li>
	</ul>
</div>

<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post' style="margin:0px">
<div id="tabsCaixa">
	<div id="tab1"><br />&nbsp;
	<%
		'Padroniza e coloca o cod_pedido_compra criptografado
		temp = Base64Encode(cod_pedido_compra) 

	if moduloPPVpedidoManual(cod_pedido_compra,"", false) Then
	%>
	<fieldset>
		<legend>Ações</legend>
		<table border="0">
			<tr>
				<% 'Aprovar pedido 
				if status_pedido_compra = 1 then
				%>
				<td>
					<button id='aprovar_pedido' class="acao">
						Aprovar Pedido
					</button>
				</td>
				<%
				'Baixa 
				elseif status_pedido_compra = 2 then
				%>
				<td>
					<button id='finaliza_pedido' class="acao">
						Finalizar Pedido
					</button>
				</td>
				<%
				'Gerar recibo 
				elseif status_pedido_compra = 3 then
				%>
				<td>
					<button id='gerar_recibo' class="acao">
						Imprimir Recibo
					</button>
				</td>
				<% end if %>
			</tr>
		</table>
	</fieldset>
	<% end if %>
	<fieldset>
	<legend>Dados Básicos</legend>
		<div class="fieldset">
		<table width='95%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
			<tr>
				<td width='27%'><strong>Status</strong></td>
				<td width='78%'><div align='left'><%=status_pedido_compra%> - <%=pedidoCompraStatus(status_pedido_compra)%></div></td>
			</tr>
			<%If comando <> "incluir" Then %>
			<tr>
				<td width='27%'><strong>N&uacute;mero Pedido de Compra</strong></td>
				<td width='78%'><div align='left' class="textoMedio"><strong><%=NUMERO_PEDIDO_COMPRA%></strong></div></td>
			</tr>
			<%
			'Abri a conexao para NÃO ter que fazer isso 4X devido a busca pelos usuários das validações
			AbrirConexao()
			%>
			<tr>
				<td width='27%'><strong>Data Cadastro </strong></td>
				<td width='78%'><div align='left'><%=ArrumaData(DATA_CADASTRO)%>&nbsp;<%=nomeUsuario(cadastro_cod_usuario,true)%></div></td>
			</tr>
			<tr>
				<td width='27%'><strong>Data Aprovado </strong></td>
				<td width='78%'><div align='left'><%=ArrumaData(DATA_APROVADO)%>&nbsp;<%=nomeUsuario(aprovado_cod_usuario,true)%></div></td></tr>
			<tr>
				<td width='27%'><strong>Data Enviado para Fornecedor</strong></td>
				<td width='78%'><div align='left'><%=ArrumaData(DATA_ENVIADO_FORNECEDOR)%>&nbsp;<%=nomeUsuario(enviado_cod_usuario,true)%></div></td>
			</tr>
			<%
			FecharConexao()
			End If%>
			
			<tr> 
  				<td><strong>Valor Total em Produtos</strong></td>
  				<td>
    				<input type='hidden' name='VALOR_TOTAL_PRODUTOS' value="<%=VALOR_TOTAL_PRODUTOS%>" />
    				<div class="camposNaoAlteraveisAjax textoMedio" id="ajax_VALOR_TOTAL_PRODUTOS"><%=VALOR_TOTAL_PRODUTOS%>&nbsp;</div>
    			</td>
			</tr>
			<tr>
  				<td width='27%'><strong>Observação </strong></td>
  				<td width='78%'>
    				<div align='left'>
      				<%If comando <> "visualizar" and comando <> "excluir"  Then%>
      					<table style='width:420px;' border='0' cellspacing='0' cellpadding='0'>
        				<tr>
          					<td width='57'><textarea rows='2' name='OBS' class='inputTamanho3' onkeydown='textCounter(this.form.OBS,this.form.charRestantes_OBS,500);' onkeyup='textCounter(this.form.OBS,this.form.charRestantes_OBS,500);'><%=OBS%></textarea></td>
          				</tr>
        				<tr>
          					<td><table width='50' border='0' align='right' cellpadding='0' cellspacing='0'>
            					<tr>
              						<td><div align='center'>
                					<input class='inputContaCaracteres' tabindex='10000' type='text' name='charRestantes_OBS' value="<%=500- Len(OBS)%>" />
                					</div></td>
              					</tr>
            				</table></td>
          				</tr>
        				</table>
      				<% Else
      					If not isNull(OBS) Then
      						response.write(Replace(OBS,Chr(13),"<br />"))
      					End If
      				End If %>
					</div>
				</td>
			</tr>
		</table>
		</div>
	</fieldset>
	<fieldset  class="fieldset">
		<legend>Dados do Fornecedor</legend>
		<div class="fieldset">
			<table width='95%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>		
				<tr>
					<td width='27%'><strong>Fornecedor *</strong></td>
					<td width='73%'><div align='left'>
					<% If  comando = "incluir" Then
		  				Response.Write(PadraoFormularioSelect("COD_FORNECEDOR","cof_fornecedor","inputTamanho3","","","Selecione",COD_FORNECEDOR,"SELECT cod_fornecedor, razao_social  FROM TB_FORNECEDORES  ORDER BY razao_social",comando,false,false,false))
		  			Else
						Response.Write(PadraoFormularioSelect("COD_FORNECEDOR","cof_fornecedor","inputTamanho3","","","Selecione",COD_FORNECEDOR,"SELECT cod_fornecedor, razao_social  FROM TB_FORNECEDORES  ORDER BY razao_social","visualizar",false,false,false)&" <input type='hidden' name='COD_FORMA_PAGAMENTO' value='"&COD_FORMA_PAGAMENTO&"' />")
					End If %>
					</div></td>
				</tr>
				<tr>
  					<td width='27%'><strong>Contato </strong></td>
  					<td width='73%'>
  						<div align='left'>
  							<%If comando <> "visualizar" and comando <> "excluir"  Then%>
								<table width="265" border="0" cellpadding="0" cellspacing="0">
  									<tr>
  										<td>
  											<input name='CONTATO' value="<%=CONTATO%>" class='inputTamanho2' type='text' maxlength='200' />
  										</td>
  										<td>&nbsp;</td>
  									</tr>
  								</table>
  							<%Else
  								If not isNull(CONTATO) Then
  									response.write(CONTATO)
  								End If
  							End If%>
  						</div>
  					</td>
				</tr>
			</table>
		</div>
	</fieldset>

	<%If comando <> "visualizar" Then%>
		<table width='90%' border='0' align='center' cellpadding='5' style='border-collapse:collapse'>
			<tr>
				<td colspan='2' align='center'><div align='left'><strong>* - Campos obrigatórios</strong><br /></div></td>
			</tr>
			<tr>
				<td colspan='2' align='center'><a href="javascript:salvar();"><img src="images/layout/btnGravar.jpg" /></a></td>
			</tr>
		</table>
	<%End If%>
	<input type='hidden' value='1' name='campo1'  />
	<input type='hidden' name='COD_PEDIDO_COMPRA' value="<%=COD_PEDIDO_COMPRA%>" /> <input type='hidden' name='NUMERO_PEDIDO_COMPRA' value="<%=NUMERO_PEDIDO_COMPRA%>" /> 
	</div>
    <div id="tab2" class="tabs">
		<br />
		<div align="center">
			<strong>
				<a href="ppv_compras_pedidos_formulario_consulta_produtos.asp?cod_pedido_compra=<%=cod_pedido_compra%> &ic=S" target="_blank">Imprimir esta tela</a>
			</strong>
		</div>
		<%if comando = "incluir" then%>
			<div class='msgGrande'>
				Grave o pedido antes de incluir produtos.
			</div>			
		<%else%>
			<iframe src="ppv_compras_pedidos_formulario_consulta_produtos.asp?comando_externo=<%=comando%>&cod_pedido_compra=<%=cod_pedido_compra%>" frameborder="0" width="99%" height="350px;"></iframe>
		<%end if %>
   </div>
</div>
</form>

<script type='text/javascript'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
	tabsMostra('tab1',<%=qtd_tabs%>,1);//Volta para guia 1 antes de continuar
	var total_produtos = new Number(document.FORMULARIO_PADRAO.VALOR_TOTAL_PRODUTOS.value);

	//if(total_produtos < 1){		
	//alert('ATENÇÃO antes de continuar é necessário que você lance pelo menos um item no pedido de compras.');
	//}else{
		if(validaForm('COD_FORNECEDOR#Cód. Fornecedor#1',document.FORMULARIO_PADRAO)){return true}else{return false}
	//}
}
<%Else%>
function validaFormMaster() {
	return true;
}
<%End If%>
</script>
<!-- #include file="files/includes/miolo-inferior.asp" -->