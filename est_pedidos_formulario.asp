<%
'data início formatado, data término, título da página
'call manutencao(ArrumaData(date(),"")& "às 08h00",ArrumaData(date(),"")& "às 12h00","")

Dim calendario_necessario, BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
calendario_necessario = "SIM"
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 1
BTN_IMPRIMIR_LINK = "#"
titulo_pagina = "Compras - Formulário"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
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
'Validação de sub-lidações padrões com formulário
'#########################################################################
If comando = "incluir" and campo1 = Empty Then
AbrirConexao()
			'1) Cria uma chave do pedido de compra para inserir nas guias adicionais
			'Obs não cria um número de pedido de compra que será gerado somente após a gravação e finalização da inserção do pedido de compra completo
			objConexao.Execute("INSERT INTO TB_PEDIDOS_COMPRAS (cod_usuario,cadastro_cod_usuario) VALUES ("&sqlNumero(Session("cod_usuario"))&","&sqlNumero(Session("cod_usuario"))&")")
			
					'Encontrar chave criada
					Set rs=objConexao.Execute("SELECT @@IDENTITY AS 'ULTIMO_ID'")
						If not rs.Eof Then
							'Continua o lançamento, mas agora já com chave criada, sendo que o incluir na verdade será o alterar.
							'Obs.: Não redirecionei com o comando alterar, pois neste caso quem tiver acesso a incluir e não tiver a alterar não conseguiria nem incluir o pedido, pois o comando já seria trocado para alteração
							cod_pedido_compra = rs("ULTIMO_ID")
						Else
							Response.Redirect("inicio.asp?msg=Desculpe houve uma falha técnica na geração do pedido de compra, se o problema persistir informe a OXNET.")
						End if 
					rs.Close
					Set rs=Nothing
					
					'Por enquanto o status é 0 (zero), ou seja em cadastro.
					status_pedido_compra = 0
			
		FecharConexao()
ElseIf comando = "incluir" and campo1 <> Empty Then
'Incluir novo registro
incluirRegistro(redirecionar_incluir)
titulo_pagina = titulo_pagina&" / Inclusão"
Elseif comando = "alterar" and campo1 = Empty Then
'Leitura inicial para possível alteração
cod_pedido_compra = Request.QueryString("cod_pedido_compra")
lerRegistro()
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "alterar" and campo1 <> Empty Then
'Gravar alteração
cod_pedido_compra = Request.Form("cod_pedido_compra")
alterarRegistro(redirecionar_alterar)
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 
'Leitura inicial para possível exclusão
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


'Funções

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
					
				'//Denominação do fornecedor
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
		'No momento da inclusão verifica se o usuário possui validação para aprovação do pedido de compra se tiver o status do pedido ao invés de 1 já pula para 2, pois o próprio usuário que possui acesso a aprovação está cadastrando um pedido de compra
		If validaUsuario(2,3,true) Then
		'Possui validação
			status_pedido_compra = 2 'Aprovado, mas não enviado para fornecedor
			
			'Chama a função que faz o update para mudar o status do pedido de compra
			Dim mudanca_status_realiza 'neste caso será realizada sem precisar exibir mensagem de erro caso não ocorra
			mudanca_status_realiza = pedidoCompraMudancaStatus(cod_pedido_compra,2,"",true)
			
		Else
		'Não possui
			status_pedido_compra = 1 'Aguardando aprovação
		End If
		'////////////////////////////
		
		
		'Atenção não é necessário fazer um where somente dos pedidos de compra com status superiores a 1 (Em cadastro) pois se estiver com este status não há numero de pedido de compra gerado e por tanto não irá contar.
		Set rs=objConexao.Execute("SELECT (ISNULL(max(numero_pedido_compra),0) + 1) as novo FROM TB_PedidoCompras")
			
			If not rs.Eof Then
				numero_pedido_compra = rs(0)
			End If
			
		rs.Close
		Set rs=Nothing	
		
		'Na inclusão na realidade faz um update e gera o númere do pedido de compra
		objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS SET numero_pedido_compra = "&sqlNumero(numero_pedido_compra)&", cod_empresa = "&sqlTexto(cod_empresa)&", cod_fornecedor = "&sqlNumero(cod_fornecedor)&", frete_por_conta = "&sqlTexto(frete_por_conta)&", frete_por_conta_acrescentar_valor_total = "&sqlTexto(frete_por_conta_acrescentar_valor_total)&", frete = "&sqlNumero(frete)&", entrega_ou_retira = "&sqlTexto(entrega_ou_retira)&", cod_forma_pagamento = "&sqlTexto(cod_forma_pagamento)&", cod_forma_recebimento = "&sqlTexto(cod_forma_recebimento)&", contato = "&sqlTexto(contato)&", obs = "&sqlTexto(obs)&", cadastro_cod_usuario = "&sqlTexto(cod_usuario)&", cod_usuario = "&sqlNumero(cod_usuario)&", modificado = "&sqlData(modificado)&", status_pedido_compra = '"&status_pedido_compra&"' WHERE cod_pedido_compra = "&cod_pedido_compra)
			

	FecharConexao()
	Response.Redirect(redirecionar&"?cod_pedido_compra="&cod_pedido_compra&"&codmsg=1")
End sub

Sub alterarRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		'Não altera o número de pedido e nem o status pelo UPDATE
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
'a = cod_pedido_compra com base encode 64
'b = status atual
'c = novo status
%>
	var msg = "", caminho = "", m = "";
	var d = document.FORMULARIO_PADRAO;
	
	if(c==2){
		msg = "Tem certeza que deseja mudar o status para Aprovado?";	
	}else if(c==3){
		msg = "Tem certeza que deseja mudar o status para Enviado para fornecedor?\nObs.: Não esqueça que você deverá enviar o pedido de compra via e-mail para o fornecedor.";		
	}else if(c==99){		
		msg = "ATENÇÃO PARA CANCELAMENTO DE PEDIDO DE COMPRA!\nTem certeza que deseja cancelar o Pedido de Compra e seus respectivos itens?\nObs.: Se houver compras relacionadas com este pedido de compras as mesmas NÃO serão canceladas.";
		m = d.cancelado_motivo.value;
	}else if(c==22){
		msg = "Tem certeza que deseja voltar o Status para Não Enviado para fornecedor?\nObs.: Desta forma você poderá alterar o Pedido de Compra novamente, mas posteriormente deverá enviar novamente para o fornecedor.";	
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
		//alert(caminho);
		document.getElementById("containerConsultaUltimoFornecedor").innerHTML = "<br /><strong>Aguarde carregando informações da última compra deste produto</strong>"
		ajax_loadContent('containerConsultaUltimoFornecedor',caminho,false,false,false);	
	}
}

</script>

<h2><%=titulo_pagina%></h2><br />
<!-- #include file="includes/compras_sub_menu.asp" -->
<%If comando = "excluir" Then%>
<div class='msgExcluir'><%=msgAoExcluir%></div>
<%End If%>

<%		
Dim qtd_tabs
qtd_tabs = 2
%>
         <div id="tabsJ">
             <ul>
				<li id="liTab1" class="current"><a href="javascript:tabsMostra('tab1',<%=qtd_tabs%>,1);"><span>Dados do Pedido de Compra</span></a>
                </li>
                <li id="liTab2">
               		<a href="javascript:tabsMostra('tab2',<%=qtd_tabs%>,2);"><span>Produtos</span></a>
                </li>
                
             </ul>
 		 </div>
  <form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post' style="margin:0px">      
<div id="tabsCaixa">
	<div id="tab1"><br />&nbsp;

<%
'Padroniza e coloca o cod_pedido_compra criptografado
temp = Base64Encode(cod_pedido_compra)
%>

<fieldset>
<legend>Dados Básicos</legend>
	<div class="fieldset">
<table width='95%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
<tr>
<td width='27%'><strong>Status</strong></td>
<td width='78%'>
<div align='left'>
<%=status_pedido_compra%> - <%=pedidoCompraStatus(status_pedido_compra)%></div></td>
</tr>
<%If comando <> "incluir" Then %>
<tr>
<td width='27%'><strong>N&uacute;mero Pedido de Compra</strong></td>
<td width='78%'>
<div align='left' class="textoMedio">
<strong><%=NUMERO_PEDIDO_COMPRA%></strong></div></td>
</tr>
<%
	'Abri a conexao para não ter que fazer isso 4X devido a busca pelos usuários das validações
	AbrirConexao()
%>
<tr>
<td width='27%'><strong>Data Cadastro </strong></td>
<td width='78%'>
<div align='left'>
<%=ArrumaData(DATA_CADASTRO)%>&nbsp;
<%=nomeUsuario(cadastro_cod_usuario,true)%></div></td>
</tr>
<tr>
<td width='27%'><strong>Data Aprovado </strong></td>
<td width='78%'>
<div align='left'>
<%=ArrumaData(DATA_APROVADO)%>&nbsp;
<%=nomeUsuario(aprovado_cod_usuario,true)%></div></td>
</tr>
<tr>
<td width='27%'><strong>Data Enviado para Fornecedor</strong></td>
<td width='78%'>
<div align='left'>
<%=ArrumaData(DATA_ENVIADO_FORNECEDOR)%>&nbsp;
<%=nomeUsuario(enviado_cod_usuario,true)%></div></td>
</tr>
	<%
	FecharConexao()
	%>
<%End If%>
<tr> 
  <td><strong>Valor Total em Produtos</strong></td>
  <td>
    <input type='hidden' name='VALOR_TOTAL_PRODUTOS' value="<%=VALOR_TOTAL_PRODUTOS%>" />
    <div class="camposNaoAlteraveisAjax textoMedio" id="ajax_VALOR_TOTAL_PRODUTOS"><%=VALOR_TOTAL_PRODUTOS%>&nbsp;</div></td>
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
      <%Else%>
      <%If not isNull(OBS) Then%>
      <%=Replace(OBS,Chr(13),"<br />")%>
      <%End If%>
      <%End If%>
      </div></td>
</tr>
</table>
	</div>
</fieldset>

<fieldset  class="fieldset">
<legend>Dados do Fornecedor</legend>
	<div class="fieldset">
<table width='95%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
<%If comando <> "incluir" Then %>
<%End If%><tr>
<td width='27%'><strong>Fornecedor *</strong></td>
<td width='73%'>
<div align='left'>

<%If comando <> "visualizar" and comando <> "excluir" and comando = "incluir" Then
	  		Response.Write(PadraoFormularioSelect("COD_FORNECEDOR","cof_fornecedor","inputTamanho3","","","Selecione",COD_FORNECEDOR,"SELECT cod_fornecedor, razao_social  FROM TB_FORNECEDORES  ORDER BY razao_social",comando,false,false,false))
	  Else
	  		Response.Write(PadraoFormularioSelect("COD_FORNECEDOR","cof_fornecedor","inputTamanho3","","","Selecione",COD_FORNECEDOR,"SELECT cod_fornecedor, razao_social  FROM TB_FORNECEDORES  ORDER BY razao_social","visualizar",false,false,false)&" <input type='hidden' name='COD_FORMA_PAGAMENTO' value='"&COD_FORMA_PAGAMENTO&"' />")
	  End If
	  %>
</div></td>
</tr>
<tr>
  <td width='27%'><strong>Valor de Frete</strong></td>
  <td width='73%'>
    <div align='left'>
      <%If comando <> "visualizar" and comando <> "excluir"  Then%>
      <input name='FRETE' value="<%=FRETE%>" class='inputTamanho1 padrao_valor' type='text' maxlength='200' />
      <div <%If Cstr(""&FRETE_POR_CONTA_ACRESCENTAR_VALOR_TOTAL) <> "S" Then%> class="ocultarCompleto" <%End If%> id="container_FRETE_POR_CONTA_ACRESCENTAR_VALOR_TOTAL">
        <table width="163" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="132">Acrescentar no valor total</td>
            <td width="31">
              <%If comando <> "visualizar" and comando <> "excluir"  Then%>
              
              <input class="inputRadio" type="checkbox" name="FRETE_POR_CONTA_ACRESCENTAR_VALOR_TOTAL" value="S" <% If Cstr(""&FRETE_POR_CONTA_ACRESCENTAR_VALOR_TOTAL) = "S" Then%> checked="checked" <%End If%> />
              <%Elseif Cstr(""&FRETE_POR_CONTA_ACRESCENTAR_VALOR_TOTAL) = "S" Then
	  	Response.Write("Sim")
	  Else
	  	Response.Write("Não")	   
      End If%>
              
              
              </td>
            </tr>
  </table>
        </div>
      
      <%Else%>
      <%If not isNull(FRETE) Then%>
      <%=FRETE%>
      <%End If%>
      <%End If%>
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
  <td>&nbsp;</td></tr>
  </table>
  <%Else%>
  <%If not isNull(CONTATO) Then%>
  <%=CONTATO%>
  <%End If%>
  <%End If%>
  </div></td>
</tr>
</table>
	</div>
</fieldset>

<fieldset  class="fieldset">
<legend>Dados da Nota Fiscal / Forma de Pagamento</legend>
	<div class="fieldset">
<table width='95%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
<tr>
<td width='27%'><strong>Forma de Pagamento *</strong></td>
<td width='73%'>
<div align='left'>
	<%If comando <> "visualizar" and comando <> "excluir" Then
	  		Response.Write(PadraoFormularioSelect("cod_forma_pagamento","cod_forma_pagamento","inputTamanho3","","","Selecione",COD_FORMA_PAGAMENTO,"SELECT cod_forma_pagamento, descricao  FROM TB_FORMAS_PAGAMENTO WHERE situacao = 'A' ORDER BY descricao",comando,false,false,false))
	  Else
	  		Response.Write(PadraoFormularioSelect("cod_forma_pagamento","cod_forma_pagamento","inputTamanho3","","","Selecione",COD_FORMA_PAGAMENTO,"SELECT cod_forma_pagamento, descricao  FROM TB_FORMAS_PAGAMENTO WHERE situacao = 'A' ORDER BY descricao","visualizar",false,false,false)&" <input type='hidden' name='COD_FORMA_PAGAMENTO' value='"&COD_FORMA_PAGAMENTO&"' />")
	  End If
	  %>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Forma de Recebimento *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<select name='COD_FORMA_RECEBIMENTO' class='inputTamanho3'>
<option value=''>Selecione</option>
<%AbrirConexao()
Set rsE15=objConexao.Execute("SELECT cod_forma_recebimento, descricao FROM TB_FORMAS_RECEBIMENTO WHERE situacao = 'A' ORDER BY descricao")
While not rsE15.Eof 
'Padrao
Dim selRelPadrao15, verificador15
verificador15 = ""&rsE15(0)

selRelPadrao15 = "3" 

If COD_FORMA_RECEBIMENTO = Empty or isNull(COD_FORMA_RECEBIMENTO) Then
COD_FORMA_RECEBIMENTO = selRelPadrao15
End if

If CStr(verificador15) = CStr(COD_FORMA_RECEBIMENTO)  Then%>
<option selected='selected' value="<%=rsE15(0)%>"><%=rsE15(1)%></option>
<%Else%>
<option value="<%=rsE15(0)%>"><%=rsE15(1)%></option>
<%End if%>
<%rsE15.MoveNext()
Wend
rsE15.Close
Set rsE15=Nothing
FecharConexao()%>
</select>
<%Else%>
<%If not isNull(COD_FORMA_RECEBIMENTO) Then%>
<%AbrirConexao()
Set rsE15=objConexao.Execute("SELECT cod_forma_recebimento, descricao FROM TB_FORMAS_RECEBIMENTO WHERE  COD_FORMA_RECEBIMENTO='"&COD_FORMA_RECEBIMENTO&"' ORDER BY descricao")
If not rsE15.Eof Then%>
<%=rsE15(1)%>
<%End if%>
<%rsE15.Close
Set rsE15=Nothing
FecharConexao()%>

<%End If%>
<%End If%>
</div></td>
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
    <div id="tab2" class="tabs"><br />
    	<div align="center"><strong><a href="ppv_compras_pedidos_formulario_consulta_produtos.asp?cod_pedido_compra=<%=cod_pedido_compra%>&ic=S" target="_blank">Imprimir esta tela</a></strong></div>
		<iframe src="ppv_compras_pedidos_formulario_consulta_produtos.asp?cod_pedido_compra=<%=cod_pedido_compra%>" frameborder="0" width="99%" height="350px;"></iframe>
   </div>
          
</div>
</form>


<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>

function validaFormMaster() {
	tabsMostra('tab1',<%=qtd_tabs%>,1);//Volta para guia 1 antes de continuar
	var total_produtos = new Number(document.FORMULARIO_PADRAO.VALOR_TOTAL_PRODUTOS.value);
		
	//if(total_produtos < 1){		
	//	alert('Atenção antes de continuar é necessário que você lance pelo menos um item no pedido de compras.');
	//}else{
		if(validaForm('COD_FORNECEDOR#Cód. Fornecedor#1|COD_FORMA_PAGAMENTO#Forma de Pagamento#2|COD_FORMA_RECEBIMENTO#Condição de Pagamento#2',document.FORMULARIO_PADRAO)){return true}else{return false}
	//}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->
