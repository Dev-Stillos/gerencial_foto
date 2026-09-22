<%
'######################################################################
'INSTRUÇÕES CORREÇÃO DE PEDIDOS
'1) Lançamento da DATA da VENDa // Muito importante, pois é calculado dados como quando o vendedores receberá
'2) Correção de Produtos Lançados e Calculo se o álbum é repasse ou não, caso for será concedito um acréscimo na comissão totalizando 10% do valor do recibo, independente se o pedido não conter somente o álbum
'3) Correção de Forma de Pagamento
'4) Inserção de Ajuda de Custo, Visualização de Quando o vendedor receberá, Observação
'5) Conclusão confirmação e finalização
'######################################################################
%>
<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/classe_modulo_cob.asp" -->
<%
Dim COD_PEDIDO, PEDIDO_EXTRA
COD_PEDIDO = Request.QueryString("cod_pedido")

Dim VERIFICA_SE_HA_PENDENCIAS
'Inicia sem pendência
VERIFICA_SE_HA_PENDENCIAS = 1
%>
<%
'//////////////////////
'Daqui para baixo não foi alterado
Dim COD_VENDEDOR
COD_VENDEDOR = 4
TOTAL_PAGO_EM_PRODUTOS = 0

Dim chave, msg, novo, TOTAL_GERAL, TOTAL_CARTAO, TOTAL_CHEQUE, PARCIAL_CHEQUE, PARCIAL_CARTAO, ARREDONDAMENTO
Dim comando, fp, id

'TB_FOTO_VENDA_CARTAO
Dim VALOR, PARCELAS, NOME_TITULAR, GRAU_PARENTESCO, AUTENTICACAO, DOCUMENTO, DATA_AUTENTICACAO, DATA_DESCEU_DIRETORIA

msg = Request.QueryString("msg")
comando = Request.QueryString("comando")
'----------------------------------------------------
If Request.Form("EM_DINHEIRO_NOVO") <> Empty then

AbrirConexao()
	objConexao.Execute("UPDATE TB_PEDIDOS SET PAGAMENTO_EM_DINHEIRO = '"&sqlNumero(Request.Form("EM_DINHEIRO_NOVO"))&"' WHERE COD_PEDIDO = "&COD_PEDIDO&" LIMIT 1")

FecharConexao()
	
	Response.Redirect("?cod_pedido="&COD_PEDIDO&"&n=n&msg=Dinheiro incluído com sucesso!")
		
End if
'Dinheiro
'///////////////////
	
	'Excluir
	If comando = "excluir" and Request.QueryString("dinheiro") = "sim" then
	
	AbrirConexao()
	objConexao.Execute("UPDATE TB_PEDIDOS SET PAGAMENTO_EM_DINHEIRO = '0' WHERE COD_PEDIDO = "&COD_PEDIDO&" LIMIT 1")
	FecharConexao()
	Response.Redirect("lca_pedidos_pendentes_alteracoes_part_3.asp?cod_pedido="&COD_PEDIDO&"&msg=Dinheiro excluído com sucesso!")
	
	end if
	
'////////////////////////
Dim cod_pedido_cartao
cod_pedido_cartao = Request.QueryString("cod_pedido_cartao")
'exclusão
	If comando = "excluir" and cod_pedido_cartao <> Empty then

	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_PEDIDOS_CARTAO WHERE COD_PEDIDO_CARTAO = "&cod_pedido_cartao&" LIMIT 1")
	FecharConexao()
	Response.Redirect("lca_pedidos_pendentes_alteracoes_part_3.asp?cod_pedido="&COD_PEDIDO&"&n=n&msg=Cartão de Crédito excluído com sucesso!") 

	end if

	'verificamos se pode exclir boleto
	if comando = "excluirboleto" and COD_PEDIDO <> "" and true then

	AbrirConexao()

		set rsBoletos = objConexao.Execute("SELECT * FROM TB_PEDIDOS WHERE COD_PEDIDO = "&sqlNumero(COD_PEDIDO))
		if not rsBoletos.Eof then
			if rsBoletos("FINALIZA_DIRETORIA") = 2 and rsBoletos("FINALIZA_CORRECAO") = 2 then
				
				set rsBoletos2 = objConexao.execute("select cod_sacado from TB_PEDIDOS_BOLETOS where COD_PEDIDO = "&sqlNumero(COD_PEDIDO) &" limit 1")
				if not rsBoletos2.eof then
					objConexao.Execute("UPDATE TB_PEDIDOS set cod_sacado = "&sqlNumero(rsBoletos2("cod_sacado"))&" WHERE COD_PEDIDO = "&sqlNumero(COD_PEDIDO))
				end if
				rsBoletos2.close
				set rsBoletos2 = Nothing
				Response.write "TB_PEDIDOS"

				objConexao.Execute("DELETE FROM TB_PEDIDOS_BOLETOS WHERE COD_PEDIDO = "&sqlNumero(COD_PEDIDO))

			end if
		end if
		rsBoletos.close
		set rsBoletos = Nothing

	FecharConexao()



	end if

	'Cheque
	Dim cod_pedido_cheque
	cod_pedido_cheque = Request.QueryString("cod_pedido_cheque")
	'exclusão
	If comando = "excluir" and cod_pedido_cheque <> Empty then
	
	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_PEDIDOS_CHEQUE_HISTORICO WHERE COD_PEDIDO_CHEQUE = "&cod_pedido_cheque&" ")
		objConexao.Execute("DELETE FROM TB_PEDIDOS_CHEQUE WHERE COD_PEDIDO_CHEQUE = "&cod_pedido_cheque&" LIMIT 1")
	FecharConexao()
	Response.Redirect("lca_pedidos_pendentes_alteracoes_part_3.asp?cod_pedido="&COD_PEDIDO&"&n=n&msg=Cheque excluído com sucesso!")
	
	end if


'----------------------------------------------------

	Dim rsVALOR_TOTAL, VALOR_TOTAL, PAGAMENTO_EM_DINHEIRO, XTOTAL_CHEQUE, XTOTAL_CARTAO, XTOTAL_BOLETO

	AbrirConexao()

		Set rsVALOR_TOTAL = objConexao.Execute("SELECT PAGAMENTO_EM_DINHEIRO, XTOTAL_CHEQUE, XTOTAL_CARTAO, XTOTAL_BOLETO FROM VW_PEDIDOS_TOTAL_RECEBIDO WHERE COD_PEDIDO = "&COD_PEDIDO&" LIMIT 1")
		
			If not rsVALOR_TOTAL.Eof Then
				
				PAGAMENTO_EM_DINHEIRO = formatNumber(0&rsVALOR_TOTAL("PAGAMENTO_EM_DINHEIRO"))
				XTOTAL_CHEQUE = formatNumber(0&rsVALOR_TOTAL("XTOTAL_CHEQUE"))
				XTOTAL_CARTAO = formatNumber(0&rsVALOR_TOTAL("XTOTAL_CARTAO"))
				XTOTAL_BOLETO = formatNumber(0&rsVALOR_TOTAL("XTOTAL_BOLETO"))
				VALOR_TOTAL = (PAGAMENTO_EM_DINHEIRO/1 + XTOTAL_CHEQUE/1 +  XTOTAL_CARTAO/1 + XTOTAL_BOLETO/1)

			End IF
		rsVALOR_TOTAL.Close()
		Set rsVALOR_TOTAL = Nothing
	FecharConexao()


'////////////////////////////////////////////////////////////////////////////////////////////
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT FINALIZA_DIRETORIA, DATA_DESCEU_DIRETORIA FROM TB_PEDIDOS WHERE COD_PEDIDO="&COD_PEDIDO)
			If not rs.Eof Then
				FINALIZA_DIRETORIA = rs("FINALIZA_DIRETORIA")
				DATA_DESCEU_DIRETORIA = Cstr(""&rs("DATA_DESCEU_DIRETORIA"))
			End If	
		rs.Close
		Set rs=Nothing
	FecharConexao()
	
	If DATA_DESCEU_DIRETORIA <> "" and Session("cod_usuario") <> 16 Then
		Response.Redirect("adm_index.asp?codmsg=1")
	End If
'////////////////////////////////////////////////////////////////////////////////////////////
						
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
<style type="text/css">
	.hidden{
		display: none;
	}
</style>
<style type="text/css">
	/* CSS for the demo. CSS needed for the scripts are loaded dynamically by the scripts */
	#mainContainer{
		width:500px;
		margin:0 auto;
		margin-top:10px;
		border:1px double #000;
		padding:3px;
		
	}
	
	/* Alternative layout for the message box - used by one of the messages */
	
	.modalDialog_contentDiv_error{
		border:3px solid #FFF;	
		padding:2px;
		z-index:100;/* Always needed	*/
		position:absolute;	/* Always needed	*/
		background-color:#F00;	/* White background color for the message */
		color:#FFF;
	}
	.modalDialog_contentDiv_error a{
		color:#FFF;
	}
	h1{
		margin-top:0px;
	}

	</style>	
	<script type="text/javascript">
	function verify(ver){
		
		if(ver){
			// Confirmed message, i.e. clicked on "Yes"
			alert('Message confirmed');
		}else{
			// Clicked on "No"
			alert('Message not confirmed');
		}
	}
	</script>
	<link rel="stylesheet" href="css/modal-message.css" type="text/css">

	<script type="text/javascript" src="js/ajax.js"></script>
	<script type="text/javascript" src="js/modal-message.js"></script>
	<script type="text/javascript" src="js/ajax-dynamic-content.js"></script>
	<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>
	
	<link type="text/css" rel="stylesheet" href="css/font-awesome.min.css"></link>
<link type="text/css" rel="stylesheet" href="js/font-awesome-animation.min.css"></link>
	
	<script type="text/javascript">
messageObj = new DHTML_modalMessage();	// We only create one object of this class
messageObj.setShadowOffset(5);	// Large shadow


function displayMessage(url)
{
	messageObj.setSource(url);
	messageObj.setCssClassMessageBox(false);
	messageObj.setSize(550,450);
	messageObj.setShadowDivVisible(true);	// Enable shadow for these boxes
	messageObj.display();
}

function displayStaticMessage(messageContent,cssClass,altura)
{
	messageObj.setHtmlContent(messageContent);
	messageObj.setSize(550,altura);
	messageObj.setCssClassMessageBox(cssClass);
	messageObj.setSource(false);	// no html source since we want to use a static message here.
	messageObj.setShadowDivVisible(false);	// Disable shadow for these boxes	
	messageObj.display();
	
	
}

function closeMessage()
{
	messageObj.close();	
}


</script>
	<script language="javascript">
		function mOvr(src, cor)
			{
			if (!src.contains(event.fromElement))
					{
						src.style.cursor = 'hand';
						src.bgColor = cor;
					}
			}
		
		function mOut(src, cor)
			{
				if (!src.contains(event.toElement))
					{
						src.style.cursor = 'default';
						src.bgColor = cor;
			
					}
			}
	</script>
	
	<script type="text/javascript" language="javascript">
	<!--AJUDA F1 E CTRL PARA BUSCA RÁPIDA-->
		
	function abrir(theURL,winName,features) 
	{ 
	  window.open(theURL,winName,features);
	}

		
	function abre(url)
	{    
				//if(event.srcElement.tagName=='TD')
				//	{
				//		src.children.tags('A')[0].click(); 
				//	}
	location.href = url;

	}	
		
		
	

	function buscaIni(){
	d = document.buscaForm
	d.busca.focus()
	d.busca.select();
	}
	function mudacorBusca(CAMPO){
	CAMPO.style.backgroundColor = '#F5ECBA'
	}
	function mudacorBusca2(CAMPO){
	CAMPO.style.backgroundColor = '#FFFFFF'
	}

 function Verificar()
 {
  var tecla=window.event.keyCode;
  if (tecla==17) {
    buscaIni();
	}
 }
 


	<!--AJUDA F1 E CTRL PARA BUSCA RÁPIDA-->
	</script>
		<!-- CALENDÁRIO SCRIPTS
	<script type="text/javascript">
	DHTMLSuite.include("calendar");
	</script>
	<script type="text/javascript">
	var calendarObjForForm = new DHTMLSuite.calendar({minuteDropDownInterval:10,numberOfRowsInHourDropDown:5,callbackFunctionOnDayClick:'getDateFromCalendar',isDragable:true,displayTimeBar:false});
	
		 
	calendarObjForForm.setCallbackFunctionOnClose('myOtherFunction');
	
	function myOtherFunction()
	{
		
		
	}
	function pickDate(buttonObj,inputObject)
	{
		calendarObjForForm.setCalendarPositionByHTMLElement(inputObject,0,inputObject.offsetHeight+2);	// Position the calendar right below the form input
		//calendarObjForForm.setInitialDateFromInput(inputObject,'yyyy-mm-dd hh:ii');	// Specify that the calendar should set it's initial date from the value of the input field.
		
		calendarObjForForm.setInitialDateFromInput(inputObject,'dd/mm/yyyy');	// Specify that the calendar should set it's initial date from the value of the input field.
		
		calendarObjForForm.addHtmlElementReference('myDate',inputObject);	// Adding a reference to this element so that I can pick it up in the getDateFromCalendar below(myInput is a unique key)
		if(calendarObjForForm.isVisible()){
			calendarObjForForm.hide();
		}else{
			calendarObjForForm.resetViewDisplayedMonth();	// This line resets the view back to the inital display, i.e. it displays the inital month and not the month it displayed the last time it was open.
			calendarObjForForm.display();
		}		
	}	
	/* inputArray is an associative array with the properties
	year
	month
	day
	hour
	minute
	calendarRef - Reference to the DHTMLSuite.calendar object.
	*/
	function getDateFromCalendar(inputArray)
	{
		var references = calendarObjForForm.getHtmlElementReferences(); // Get back reference to form field.
		references.myDate.value = inputArray.day + '/' + inputArray.month + '/' + inputArray.year
//		references.myDate.value = inputArray.year + '-' + inputArray.month + '-' + inputArray.day + ' ' + inputArray.hour + ':' + inputArray.minute;
		calendarObjForForm.hide();	
		
	}	
	</script>
	-->
<script>
function confirma() {
	
	if(YNconfirm("Mensagem")){
			document.dados.submit();			
	}	
}
</script>
<script language="javascript">
function YNconfirm(txt){
return typeof(suporteVBscript)=="undefined"?confirm(txt):VBconfirm(txt)==6
}
</script>
<script language="vbscript">
suporteVBscript=1
function VBconfirm(mensagem)

 VBconfirm=msgbox(mensagem,vbyesno)
end function

</script>
<script>
function confirmaProdAparte(PRODUTO_A_PARTE){

	if(YNconfirm("Tem certeza que deseja excluir este produto?")){
			location.href = "?cod_pedido=<%=COD_PEDIDO%>&prod_excluir_a_parte="+PRODUTO_A_PARTE
	}

}
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  if (selObj.value !== ""){
  	eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  	if (restore) selObj.selectedIndex=0;
  }
}

function deletaCheque(COD_PEDIDO_CHEQUE){
	if(YNconfirm("Tem certeza que deseja excluir este cheque?")){
			location.href = "?cod_pedido=<%=COD_PEDIDO%>&cod_pedido_cheque="+COD_PEDIDO_CHEQUE+"&comando=excluir"
	}
}
function deletaCartao(COD_PEDIDO_CARTAO){
	if(YNconfirm("Tem certeza que deseja excluir este cartão de crédito?")){
			location.href = "?cod_pedido=<%=COD_PEDIDO%>&cod_pedido_cartao="+COD_PEDIDO_CARTAO+"&comando=excluir"
	}
}
function deletaDinheiro(){
	if(YNconfirm("Tem certeza que deseja excluir a forma de pagamento dinheiro?")){
			location.href = "?cod_pedido=<%=COD_PEDIDO%>&dinheiro=sim&comando=excluir"
	}
}
function alteraCheque(COD_PEDIDO_CHEQUE){
	var caminho
	caminho = "<table cellpadding=\'0\' cellspacing=\'0\' border=\'0\' width=\'100%\'><tr><td><table cellpadding=\'0\' cellspacing=\'0\' width=\'100%\'><tr><td width=\'70%\'><h1>Cheque pré-datado</h1></td><td align=\'right\'><a href=\'#\' onclick=\'closeMessage();return false\'>Fechar Janela</a></td></tr></table></td></tr><tr><td><iframe frameborder=\'0\' scrolling=\'no\' height=\'380px\' width=\'99%\' src=\'lca_pedidos_pendentes_alteracoes_part_3_POPUP_cheque.asp?cod_pedido=<%=COD_PEDIDO%>&cod_pedido_cheque="+COD_PEDIDO_CHEQUE+"\'></iframe></td></tr></table>"
	displayStaticMessage(caminho,false,420)		
}

function alteraCartao(COD_PEDIDO_CARTAO){
	var caminho
	caminho = "<table cellpadding=\'0\' cellspacing=\'0\' border=\'0\' width=\'100%\'><tr><td><table cellpadding=\'0\' cellspacing=\'0\' width=\'100%\'><tr><td width=\'70%\'><h1>Cartão de Crédito</h1></td><td align=\'right\'><a href=\'#\' onclick=\'closeMessage();return false\'>Fechar Janela</a></td></tr></table></td></tr><tr><td><iframe frameborder=\'0\' scrolling=\'no\' height=\'300px\' width=\'99%\' src=\'lca_pedidos_pendentes_alteracoes_part_3_POPUP_cartao.asp?cod_pedido=<%=COD_PEDIDO%>&cod_pedido_cartao="+COD_PEDIDO_CARTAO+"&comando=alterar\'></iframe></td></tr></table>"
	displayStaticMessage(caminho,false,350)	
}
function alteraDinheiro(valor_dinheiro){
	var caminho
	caminho = "<table cellpadding=\'0\' cellspacing=\'0\' border=\'0\' width=\'100%\'><tr><td><table cellpadding=\'0\' cellspacing=\'0\' width=\'100%\'><tr><td width=\'70%\'><h1>Dinheiro</h1></td><td align=\'right\'><a href=\'#\' onclick=\'closeMessage();return false\'>Fechar Janela</a></td></tr></table></td></tr><tr><td><iframe frameborder=\'0\' scrolling=\'no\' height=\'120px\' width=\'99%\' src=\'lca_pedidos_pendentes_alteracoes_part_3_POPUP_dinheiro.asp?cod_pedido=<%=COD_PEDIDO%>&valor_em_dinheiro="+valor_dinheiro+"&comando=alterar\'></iframe></td></tr></table>"
	displayStaticMessage(caminho,false,180)	
}

function adicionarFormaPagamento(valor){
	if(valor == 1){
	//Dinheiro
		var caminho
		caminho = "<table cellpadding=\'0\' cellspacing=\'0\' border=\'0\' width=\'100%\'><tr><td><table cellpadding=\'0\' cellspacing=\'0\' width=\'100%\'><tr><td width=\'70%\'><h1>Dinheiro</h1></td><td align=\'right\'><a href=\'#\' onclick=\'closeMessage();return false\'>Fechar Janela</a></td></tr></table></td></tr><tr><td><iframe frameborder=\'0\' scrolling=\'no\' height=\'120px\' width=\'99%\' src=\'lca_pedidos_pendentes_alteracoes_part_3_POPUP_dinheiro.asp?cod_pedido=<%=COD_PEDIDO%>&valor_em_dinheiro=0&comando=incluir\'></iframe></td></tr></table>"
		
			displayStaticMessage(caminho,false,180)	
			//Volta para opção zero
			document.addpagamento.ADD_FORMA_PAGAMENTO.options[0].selected = "true";
		
	}else if (valor == 2) {
	//Cheque
			var caminho
			caminho = "<table cellpadding=\'0\' cellspacing=\'0\' border=\'0\' width=\'100%\'><tr><td><table cellpadding=\'0\' cellspacing=\'0\' width=\'100%\'><tr><td width=\'70%\'><h1>Cheque pré-datado</h1></td><td align=\'right\'><a href=\'#\' onclick=\'closeMessage();return false\'>Fechar Janela</a></td></tr></table></td></tr><tr><td><iframe frameborder=\'0\' scrolling=\'yes\' height=\'370px\' width=\'99%\' src=\'lca_pedidos_pendentes_alteracoes_part_3_POPUP_cheque_inclusao.asp?cod_pedido=<%=COD_PEDIDO%>&comando=incluir\'></iframe></td></tr></table>"
			displayStaticMessage(caminho,false,400)	
			//Volta para opção zero
			document.addpagamento.ADD_FORMA_PAGAMENTO.options[0].selected = "true";
		
	
	}else if (valor == 3) {
	//Cartão
			var caminho
			caminho = "<table cellpadding=\'0\' cellspacing=\'0\' border=\'0\' width=\'100%\'><tr><td><table cellpadding=\'0\' cellspacing=\'0\' width=\'100%\'><tr><td width=\'70%\'><h1>Cartão de Crédito</h1></td><td align=\'right\'><a href=\'#\' onclick=\'closeMessage();return false\'>Fechar Janela</a></td></tr></table></td></tr><tr><td><iframe frameborder=\'0\' scrolling=\'no\' height=\'300px\' width=\'99%\' src=\'lca_pedidos_pendentes_alteracoes_part_3_POPUP_cartao.asp?cod_pedido=<%=COD_PEDIDO%>&comando=incluir\'></iframe></td></tr></table>"
			displayStaticMessage(caminho,false,350)	
			//Volta para opção zero
			document.addpagamento.ADD_FORMA_PAGAMENTO.options[0].selected = "true";
	}
	else if (valor == 4)
	{

		var caminho
		caminho = "<table cellpadding=\'0\' cellspacing=\'0\' border=\'0\' width=\'100%\'><tr><td><table cellpadding=\'0\' cellspacing=\'0\' width=\'100%\'><tr><td width=\'70%\'><h1>Boleto Bancário</h1></td><td align=\'right\'><a href=\'#\' onclick=\'closeMessage();return false\'>Fechar Janela</a></td></tr></table></td></tr><tr><td><iframe frameborder=\'0\' scrolling=\'yes\' height=\'300px\' width=\'99%\' src=\'lca_pedidos_pendentes_alteracoes_part_3_POPUP_boleto.asp?cod_pedido=<%=COD_PEDIDO%>&comando=incluir\'></iframe></td></tr></table>"
		displayStaticMessage(caminho,false,350)	
		//Volta para opção zero
		document.addpagamento.ADD_FORMA_PAGAMENTO.options[0].selected = "true";
	}
}

</script>
</head>

<body onKeyDown="javascript:Verificar()">

<table cellpadding="9px" cellspacing="0" width="100%">
<tr>
<td>
<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td>
		<!-- #include file="includes/topo.asp" -->
<table width="100%" style="border-bottom:#999999 solid dotted 1px" background="images/layout/fundo_menu_azul.jpg"border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
		<!-- #include file="includes/menu.asp" -->
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="10">
  <tr>
  <td>
   <%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 2
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"

AbrirConexao()
  Dim PAGAMENTO_EM_BOLETO
  'VERIFICAMOS SE JÁ EXISTEM BOLETOS PARA ESTE PEDIDO
  Set rsBoletos = objConexao.Execute("SELECT * FROM TB_PEDIDOS_BOLETOS WHERE COD_PEDIDO = "&sqlNumero(COD_PEDIDO))
  if not rsBoletos.Eof then
  	PAGAMENTO_EM_BOLETO = 1
  else
  	PAGAMENTO_EM_BOLETO = 0
  end if
  rsBoletos.close
  set rsBoletos = Nothing
FecharConexao()

  %>
  <!-- #include file="includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">

	  <h2>LCA - Pedidos Pendentes / Corre&ccedil;&atilde;o</h2>
	  <div class="subMenu">
	    <table width="600" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="10%">Sub-Menu: </td>
            <td width="39%"><div align="right" class="textoMenor">
              <div align="center"><span class="textoMedio">Adicionar Forma de Pagamento</span>: </div>
            </div></td>
            <td width="51%"><div align="center">
              <form name="addpagamento" id="form1" style="margin:0px">
                <div align="left">
                  <select name="ADD_FORMA_PAGAMENTO" onchange="adicionarFormaPagamento(this.value)">
                              <option selected="selected" value="">---</option>
                              <%if PAGAMENTO_EM_DINHEIRO =< 0 then%>
						      <option value="1">Dinheiro</option>
						      <%end if%>

                              <option value="2">Cheque</option>
                              <option value="3">Cart&atilde;o de Cr&eacute;dito</option>
							  <%if PAGAMENTO_EM_BOLETO =< 0 then%>
                              	<option value="4">Boleto</option>
                              <%end if%>
                  </select>
                  </div>
              </form>
              </div></td>
          </tr>
        </table>
	 </div>
	  <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	  
		  	<h3 align="center">Corre&ccedil;&atilde;o de Pedido - Forma de Pagamento </h3>
			<table width="98%" align="center" 	cellpadding="0" cellspacing="3">
			<tr>
            <td width="12%"><div align="left"><strong>Vendedor</strong>:</div></td>
            <td width="88%"><%
			AbrirConexao()
			'Nome do vendedor
			Dim rsVendedorQ
			Set rsVendedorQ = objConexao.Execute("SELECT TB_VENDEDORES.NOME_VENDEDOR,TB_PEDIDOS.PEDIDO_EXTRA FROM TB_VENDEDORES INNER JOIN TB_PEDIDOS ON TB_VENDEDORES.COD_VENDEDOR = TB_PEDIDOS.COD_VENDEDOR WHERE TB_PEDIDOS.COD_PEDIDO="&COD_PEDIDO&" LIMIT 1")
				
				If not rsVendedorQ.Eof Then
					PEDIDO_EXTRA = rsVendedorQ("PEDIDO_EXTRA")
					Response.Write(rsVendedorQ("NOME_VENDEDOR"))
				End If
			
			rsVendedorQ.Close
			Set rsVendedorQ = Nothing
			
			%></td>
            </tr>
			<tr>
			  <td><strong>Formando(s):</strong></td>
			  <td><%
			sql = 	"SELECT TB_PEDIDOS_FORMANDOS.COD_FORMANDO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO "&_
					"FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON "&_
					"TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO "&_
					"WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&COD_PEDIDO&" ORDER BY NOME_FORMANDO"
			Set rs3 = objConexao.Execute(sql)
				
				While not rs3.Eof
				%>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				  <tr>
					<td><div align="left"><%=rs3("NOME_FORMANDO")%> (<%=digito(rs3("COD_FORMANDO"))%>) </div></td>
					</tr>
				</table>
				<%
				rs3.MoveNext()
				Wend
				
			rs3.Close
			Set rs3 = Nothing
			FecharConexao()
			%></td>
			  </tr>
			</table>
		  	<p align="center">
<%if PAGAMENTO_EM_DINHEIRO > 0 then%>
					<form name="alterarDinheiro" method="post">
                  <table width="98%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
					<tr bgcolor="#E4E4E4">
                        <td colspan="2"><div align="center" class="textoMedio"><strong>Dinheiro</strong></div>                          
                          <div align="center"></div>                          <div align="center"></div>                          <div align="center"></div>                          <div align="center"></div>                          <div align="center"></div></td>
                      </tr>             
					<tr>
						<td><div align="left"><table width="210" border="0" cellspacing="0" cellpadding="0">
						  <tr>
							<td width="62%"><div align="left" class="textoMedio">Total em Dinheiro:</div></td>
							<td width="38%"><div align="right" class="textoMedio"><strong><%=formatNumber(PAGAMENTO_EM_DINHEIRO)%></strong></div></td>
						  </tr>
						</table>
						  </div></td>
						<td width="6%"><div align="center"><a href="javascript:void(0)" onclick="alteraDinheiro('<%=formatNumber(PAGAMENTO_EM_DINHEIRO)%>')"><img alt="Alterar Dinheiro" src="images/icons/icon_alterar.png" border="0" /></a> &nbsp;<a href="javascript:void(0)" onclick="deletaDinheiro()"><img alt="Excluir Dinheiro" src="images/icons/icon_excluir.png" border="0" /> </a> </div></td>
						    </tr>
                    </table>
					</form>
					<%end if%>
					<%
					Dim rsCartao
					
					AbrirConexao()
					
					Set rsCartao = objConexao.Execute("SELECT * FROM TB_PEDIDOS_CARTAO WHERE COD_PEDIDO="&COD_PEDIDO&" ORDER BY NOME_TITULAR") 
					
				
					if Not rsCartao.Eof then
					
					%>
<form method="post" style="margin:0px">
  <table width="98%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
					<tr bgcolor="#E4E4E4">
                        <td colspan="8"><div align="center" class="textoMedio"><strong>Cart&atilde;o de Cr&eacute;dito/D&eacute;bito</strong></div>                          
                          <div align="center"></div>                          <div align="center"></div>                          <div align="center"></div>                          <div align="center"></div>                          <div align="center"></div></td>
                      </tr>
                      <tr>
                        <td width="25%"><div align="center"><strong>Nome Titular </strong></div></td>
                        <td width="15%"><div align="center"><strong>Cart&atilde;o</strong></div></td>
                        <td width="13%"><div align="center"><strong>Autentica&ccedil;&atilde;o </strong></div></td>
                        <td width="11%"><div align="center"><strong>Documento</strong></div></td>
                        <td width="10%"><div align="center"><strong>Parcelas </strong></div></td>
                        <td width="11%"><div align="center"><strong>Valor Total </strong></div></td>
                        <td width="9%"><div align="center"><strong>Status</strong></div></td>
                        <td width="6%"><div align="center"><strong>Alterar</strong></div></td>
                      </tr>
                      <%
					  While not rsCartao.Eof
					  %>
                      <tr>
                        <td><div align="left"><%=rsCartao("NOME_TITULAR")%></div></td>
                        <td><div align="center"><%=rsCartao("CARTAO")%></div></td>
                        <td><div align="center"><%=rsCartao("AUTENTICACAO")%></div></td>
                        <td><div align="center"><%=rsCartao("DOCUMENTO")%></div></td>
                        <td><div align="center"><%=rsCartao("PARCELAS")%></div></td>
                        <td><div align="right"><%=formatNumber(rsCartao("VALOR_TOTAL"))%>
                              <%
							  PARCIAL_CARTAO=0
						PARCIAL_CARTAO=formatNumber(rsCartao("VALOR_TOTAL"))
						TOTAL_CARTAO=(TOTAL_CARTAO+PARCIAL_CARTAO)/1
						PARCIAL_CARTAO=0						
						%>
                        </div></td>
                        <td><div align="center">
						<%If rsCartao("VERIFICADO") = 1 Then%>
							<img src="images/icons/icon_balao_green.gif" width="17" height="16" />
						<%Else
							VERIFICA_SE_HA_PENDENCIAS = 2
							%>
							<img src="images/icons/icon_balao_red.gif" width="17" height="16" />
						<%End If%>
						</div></td>
                        <td><div align="center"><a href="javascript:void(0)" onclick="alteraCartao(<%=rsCartao("COD_PEDIDO_CARTAO")%>)"><img alt="Alterar Cartão" src="images/icons/icon_alterar.png" border="0" /></a> &nbsp;<a href="javascript:void(0)" onclick="deletaCartao(<%=rsCartao("COD_PEDIDO_CARTAO")%>)"><img alt="Excluir Cartão" src="images/icons/icon_excluir.png" border="0" /> </a> </div></td>
                      </tr>
                      <%
					  rsCartao.MoveNext
					  Wend
					  
					  %>
                      <tr>
                        <td colspan="8">			
						<div align="left">
						<table width="210" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="62%"><div align="left" class="textoMedio">Total em Cart&atilde;o:</div></td>
    <td width="38%"><div align="right" class="textoMedio"><strong><%=formatNumber(TOTAL_CARTAO)%></strong></div></td>
  </tr>
</table>  </div>                       
                          </td>
                        </tr>
                    </table>
		</form>
					<%
					end if
					
					rsCartao.Close
					Set rsCartao=Nothing
					
					if TOTAL_CARTAO > 0 then 
					TOTAL_GERAL = TOTAL_GERAL + formatNumber(TOTAL_CARTAO)
					end if  
					  
					Dim rsCheque
					Set rsCheque = objConexao.Execute("SELECT * FROM TB_PEDIDOS_CHEQUE WHERE COD_PEDIDO="&COD_PEDIDO&" ORDER BY BOM_PARA") 
					if Not rsCheque.Eof then
					%>
					<form method="post">
					<table width="98%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
					<tr bgcolor="#E4E4E4">
                        <td colspan="8"><div align="center" class="textoMedio"><strong>Cheque</strong></div>                          <div align="center"></div>                          <div align="center"></div>                          <div align="center"></div>                          <div align="center"></div>                          <div align="center"></div></td>
                      </tr>
                      <tr>
                        <td width="25%"><div align="center"><strong>Nome Titular </strong></div></td>
                        <td width="15%"><div align="center"><strong>CPF</strong></div></td>
                        <td width="13%"><div align="center"><strong>C. Corrente </strong></div></td>
                        <td width="11%"><div align="center"><strong>N&ordm; Cheque </strong></div></td>
                        <td width="10%"><div align="center"><strong>Bom Para</strong></div></td>
                        <td width="11%"><div align="center"><strong>Valor Cheque </strong></div></td>
                        <td width="9%"><div align="center"><strong>Status</strong></div></td>
                        <td width="6%"><div align="center"><strong>Alterar</strong></div></td>
                      </tr>
                      <%
					  TOTAL_CHEQUE = 0
					  Dim idParcial2, qtdChequesa
					  idParcial2 = 0
					  qtdChequesa = 0
					  While Not rsCheque.Eof
					  qtdChequesa = qtdChequesa + 1
					  idParcial2 = formatNumber(rsCheque("COD_PEDIDO_CHEQUE"))
					  %>
                      <tr>
                        <td><div align="left"><%=rsCheque("NOME")%></div></td>
                        <td><div align="center"><%=mascaraCPF(rsCheque("CPF"))%></div></td>
                        <td><div align="center"><%=rsCheque("CONTA_CORRENTE")%></div></td>
                        <td><div align="center"><%=rsCheque("NUMERO_CHEQUE")%></div></td>
                        <td><div align="center"><%=ArrumaData(rsCheque("BOM_PARA"))%></div></td>
                        <td><div align="right"><%=FormatNumber(rsCheque("VALOR"))%><%
						PARCIAL_CHEQUE = formatNumber(rsCheque("VALOR"))
						TOTAL_CHEQUE=TOTAL_CHEQUE+formatNumber(PARCIAL_CHEQUE)
						PARCIAL_CHEQUE = ""
						%></div></td>
                        <td><div align="center">
						<%If rsCheque("VERIFICADO") = 1 Then%>
							<img src="images/icons/icon_balao_green.gif" width="17" height="16" />
						<%Else
							VERIFICA_SE_HA_PENDENCIAS = 2
						%>
							<img src="images/icons/icon_balao_red.gif" width="17" height="16" />
						<%End If%>
						</div></td>
                        <td><div align="center"><a href="javascript:void(0)" onclick="alteraCheque(<%=rsCheque("COD_PEDIDO_CHEQUE")%>)"><img alt="Alterar Cheque" src="images/icons/icon_alterar.png" border="0" /></a> &nbsp;<a href="javascript:void(0)" onclick="deletaCheque(<%=rsCheque("COD_PEDIDO_CHEQUE")%>)"><img alt="Excluir Cheque" src="images/icons/icon_excluir.png" border="0" /> </a> </div></td>
                      </tr>
                      <%
					  rsCheque.MoveNext
					  Wend
					  
					  %>
                      <tr>
                        <td colspan="8">
                        <div align="left">
						<table width="400" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="33%"><div align="left" class="textoMedio">Total em Cheque:</div></td>
    <td width="20%"><div align="right" class="textoMedio"><strong><%=formatNumber(TOTAL_CHEQUE)%></strong></div></td>
    <td width="47%">&nbsp;&nbsp; |&nbsp; &nbsp;<span class="textoMedio">Qtd. cheque(s): <strong><%=qtdChequesa%></strong></span></td>
  </tr>
</table>
						</div>
						
						</td>
                        </tr>
                    </table>

				</form>
					<%
					  rsCheque.Close
					  Set rsCheque = Nothing
					  end if
					  

					if TOTAL_CHEQUE > 0 then 
					TOTAL_GERAL = TOTAL_GERAL + formatNumber(TOTAL_CHEQUE)
					end if
					  
						'PAGAMENTO EM BOLETO
					sql = "SELECT nome, email, cpf, GRAU_PARENTESCO, TB_SACADOS_ENDERECOS.* , TB_PEDIDOS.cod_pedido FROM TB_SACADOS INNER JOIN TB_SACADOS_ENDERECOS on TB_SACADOS.cod_sacado = TB_SACADOS_ENDERECOS.cod_sacado INNER JOIN TB_PEDIDOS on TB_PEDIDOS.cod_sacado = TB_SACADOS.cod_sacado where TB_PEDIDOS.cod_pedido = "&sqlNumero(COD_PEDIDO)
					set rs = objConexao.Execute(sql)
					if not rs.Eof then
					%> 
					<from method='POST' action='#'>
						<table width="98%" border="1" align="center" cellpadding="5" style="border-collapse:collapse; margin-top:10px;">
							<tr bgcolor="#E4E4E4"><td colspan="6"><div align="center" class="textoMedio"><strong>Boleto</strong></div></td></tr>
							<tr>
								<td width="10%"><div align="center"><strong>Sacado:</strong></div></td>
								<td>
									<font face='Courier'>
									<strong>Nome:</strong><%=rs("nome")%> <strong>CPF:</strong><%=mascaraCPF(rs("cpf"))%> <br />
									<strong>E-mail:</strong><%=rs("email")%> <br />
									<strong>Endereço:</strong><%=rs("endereco")%>, <%=rs("Numero")%> - <%=rs("complemento")%> <br>
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <%=rs("bairro")%>, <%=rs("cidade")%> / <%=rs("estado")%>
									</font>
								</td>
								<td rowspan="2" width="150px">
									<center>
										<a href="javascript:void(0)" class='excluir_boletos_sacados'>Excluir boletos e sacado</a> <br /><br />
										<a href="javascript:void(0)" class='alterar_boletos'>Alterar boletos </a> <br /><br />
										<a href="javascript:void(0)" class='alterar_sacados'>Alterar infromações do Sacado</a><br /><br />
									</center>
								</td>
							</tr>
							<tr>
								<td width="25%"><div align="center"><strong>Boletos:</strong></div></td>
								<td>
									<font face='Courier' size='2'>
										Total:
										<%
											dim qtd_parcelas_boletos
											dim valor_total_boletos

											set rs2 = objConexao.Execute("select count(parcela) as qtd_parcelas , sum(valor) as total from TB_PEDIDOS_BOLETOS where cod_pedido = "&sqlNumero(COD_PEDIDO))
											if Not rs2.eof then
												qtd_parcelas_boletos = rs2("qtd_parcelas")
												valor_total_boletos	= rs2("total")
											end if
											rs2.close
											set rs2 = Nothing

											%>
											<br />	
											R$ <%=arrumaNumero(valor_total_boletos,0,2) %> em <%=qtd_parcelas_boletos%>x
											<br /><br />


											Parcelas:	
											<%

											dim quebra :  quebra = round(int(qtd_parcelas_boletos&"" )/2,0)
											dim ct_linha : ct_linha = 0	

											set rs2 = objConexao.execute("SELECT * FROM TB_PEDIDOS_BOLETOS WHERE cod_pedido = "&sqlNumero(cod_pedido))
											While not rs2.Eof

												if ct_linha = 0 then
													%> <table width="100%"> <tr'><td width='200px;'  valign="top"><%
												end if

												if ct_linha = quebra then
													%></td><td valign="top"><%
												end if
													Response.write(preencheCaracter(rs2("parcela"),"0",2)&" "&ArrumaData(rs2("data_vencimento"))&" R$"&arrumaNumero(rs2("valor"),"",2)&"<br>")

												if ct_linha = int(qtd_parcelas_boletos&"") then
													%></td></tr></table><%
												end if 

												ct_linha = ct_linha + 1

											rs2.MoveNext
											Wend
											rs2.Close
											set rs2 = Nothing

										%>
									</font>
								</td>
							</tr>
							
						</table>




						</form>
					<%
					end if
					rs.Close
					set rs = Nothing
					%>

					</tr>
					<tr>
								<td colspan="3">
									<div class="excluir_confirmacao hidden" >
										
										<div class="msgGrandeErro">
											Deseja excluir o boleto e o sacado atrelado a esse pedido
										</div>
					
										<center>
											<button id="confirma_exclusao">Excluir</button> 
											<button class="cancelar">Cancelar</button>
										</center>
									</div>
									<div class="alterar_boletos_form hidden">
										<table width="600px%" border="1" align="center" cellpadding="3" cellspacing="3" style="border-collapse:collapse;">
										<tr>
										<td width="189"><div align="left">Primeiro Vencimento:</div></td>
										<td width="360">
										  <div align="left">
										    <select name='dia_vencimento' required="required" id='dia_vencimento_id' class='valida_vencimento validacao_etapa_1 vencimento' nome-campo='dia vencimento'>
										      <option value=""></option>
										      <option value="10">10</option>
										      <option value="25">25</option>
										    </select>
										    /
										    <select name='mes_vencimento' required="required" id='mes_vencimento_id' class='valida_vencimento validacao_etapa_1 vencimento' nome-campo='mês vencimento' >
										      <option value=""></option>
										      <option value="1">Janeiro</option>
										      <option value="2">Fevereiro</option>
										      <option value="3">Março</option>
										      <option value="4">Abril</option>
										      <option value="5">Maio</option>
										      <option value="6">Junho</option>
										      <option value="7">Julho</option>
										      <option value="8">Agosto</option>
										      <option value="9">Setembro</option>
										      <option value="10">Outubro</option>
										      <option value="11">Novembro</option>
										      <option value="12">Dezembro</option>
										    </select>
										    /
										    <select name='ano_vencimento' required="required" id='ano_vencimento_id' class='valida_vencimento validacao_etapa_1 vencimento' nome-campo='ano vencimento'>
										      <option value=""></option>
										      <option value="2016">2016</option>
										      <option value="2017">2017</option>
										    </select>
										  </div>
										</td>
										</tr>
										<tr>
										<td width="189"><div align="left">Quant. Parcelas:</div></td>
										<td width="360"><div align="left"><%=parcelasBoleto(cod_pedido, "html", false)%></div></td>
										</tr>
										</table>
										<br />
										<center>

										<button id="confirma_alterarcao_boleto">Alterar</button>

										<button class="cancelar">Cancelar</button>

										</center>
									</div>
									<div class="alterar_sacados_form hidden">
									<br />

	<table width="600px%" border="1" align="center" cellpadding="3" cellspacing="3" style="border-collapse:collapse;">
  <tr>
    <td><div align="left">Sacado:</div></td>
    <td>
      <div align="left">
      	<%
      	AbrirConexao()
      	set rs = objConexao.Execute("select cod_sacado from TB_PEDIDOS_BOLETOS where cod_pedido = "&sqlNumero(cod_pedido)&" limit 1")
      	if not rs.eof then
      		Response.write ("<input type='hidden' value='"&rs(0)&"' id='cod_sacado_id' />") 
      	end if
      	rs.close
      	set rs = Nothing
      	FecharConexao()
      	%>
        <select id='cod_formando_id' style="margin-bottom:5px; width:293px;" class="validacao_etapa_2 sacado" required="required"  nome-campo='Parcelas' >
          <option value="">Selecione o nome do Sacado</option>
          <%
          AbrirConexao()
          sql = "SELECT TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO, TB_FICHAS_IDENTIFICACAO.ALBUM, TB_FICHAS_IDENTIFICACAO.INICIAL, TB_FICHAS_IDENTIFICACAO.BRINDE, TB_FICHAS_IDENTIFICACAO.PRODUTO_AGREGADO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, TB_FICHAS_IDENTIFICACAO.COMISSAO, TB_PEDIDOS_FORMANDOS.COD_FORMANDO FROM TB_FICHAS_IDENTIFICACAO INNER JOIN TB_PEDIDOS_FORMANDOS ON TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = TB_PEDIDOS_FORMANDOS.COD_FORMANDO "&_
            "WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&COD_PEDIDO
          Set rsFormsCor = objConexao.Execute(sql)
          While not rsFormsCor.Eof
          %>
            <option value='<%=rsFormsCor("COD_FORMANDO")%>'><%=rsFormsCor("NOME_FORMANDO")%></option>
          <%
          rsFormsCor.MoveNext()
          Wend
          rsFormsCor.Close
          Set rsFormsCor = Nothing 
          FecharConexao()
          %>
          <option value="#outro">Outro...</option>
        </select>
        <input name="NOME_OUTRO" id='nome_formando_id' type="text" size="40" maxlength="80" class="hidden required_form sacado" nome-campo='Nome Formando' />
      </div>
    </td>
  </tr>
  <tr>
    <td width="189"><div align="left">CPF*:</div></td>
    <td width="360"><div align="left"><input name="CPF" type="text" id='cpf_id'  size="40" maxlength="80" class="validacao_etapa_2 padrao_cpf sacado" required="required"  nome-campo='CPF'/></div></td>
  </tr>
  
  <tr class='oculta_formando hidden'>
    <td><div align="left">Se for de terceiro informe o grau de parentesco:</div></td>
    <td><div align="left">
      <input name="GRAU_PARENTESCO" id='grau_parentesco_id' type="text" size="40" maxlength="50"  />
    </div></td>
  </tr>

  <tr></tr>
    <td height="30"><div align="left">E-mail*:</div></td>
    <td><div align="left">
      
      <input name="email" type="email" id='email_id' required="required" size="40" maxlength="80"  class="validacao_etapa_2 sacado" nome-campo='E-mail' />
    </div></td>
  </tr>
  </table> <br />


										<table width="600px%" class="oculta_formando hidden" border="1" align="center" cellpadding="3" cellspacing="3" style="border-collapse:collapse;">
										  <tr>
										    <td height="30"><div align="left">CEP:</div></td>
										    <td>
										      <input name="CEP" type="text" size="28" class="padrao_cep" id='cep_id' maxlength="80" class="validacao_etapa_3 required_form"  nome-campo='CPF' />
										<input type="button" id='buscar_cep_id' value="Busar CEP" />

										      <input name="CEP" type="hidden" id='cep_ultima_pesquisa' maxlength="80"  />
										      &nbsp;&nbsp;<i class="fa fa-refresh faa-spin animated hidden " id='carregandoa' style="font-size:16px;"></i>
										    </td>
										  </tr>

										  <tr>
										    <td width="189"><div align="left">Endereço*:</div></td>
										    <td width="360"><input name="ENDERECO" id='endereco_id' type="text" size="40" maxlength="255" class="validacao_etapa_3 required_form sacado" nome-campo='Endereço'/></td>
										  </tr>

										  <tr>
										    <td width="189"><div align="left">Número*:</div></td>
										    <td width="360"><input name="NUMERO" id='numero_id' type="number" maxlength="80" min='1' class="validacao_etapa_3 required_form sacado"  nome-campo='Número' /></td>
										  </tr>


										  <tr>
										    <td width="189"><div align="left">Complemento:</div></td>
										    <td width="360"><input name="COMPLEMENTO" id="complemento_id" type="text" size="40" maxlength="80" nome-campo='Complemento' /></td>
										  </tr>

										  <tr>
										    <td width="189"><div align="left">Bairro*:</div></td>
										    <td width="360"><input name="BAIRRO" id='bairro_id' type="text" size="40" maxlength="80" class="validacao_etapa_3 required_form sacado" nome-campo='Bairro' /></td>
										  </tr>

										  <tr>
										    <td width="189"><div align="left">Cidade*:</div></td>
										    <td width="360"><input name="CIDADE" id='cidade_id' type="text" size="40" maxlength="80" class="validacao_etapa_3 required_form sacado" nome-campo='Cidade' /></td>
										  </tr>

										  <tr>
										    <td width="189"><div align="left">Estado*:</div></td>
										    <td width="360"><input name="ESTADO" id='estado_id' type="text" size="40" maxlength="80" class="validacao_etapa_3 required_form sacado" nome-campo='Estado' /></td>
										  </tr>
										</table> <br />
										<br />
		
<center>
											
										<button id='confirma_alterarcao_sacado'>Alterar</button>
										
										<button class='cancelar'>Cancelar</button>

</center>


									</div>
								</td>
							</tr>
					</table>
					<br />
					<%If int(PEDIDO_EXTRA) = 2 Then%>
	                   	 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr>
    <td width="279"><div align="left"><strong class="textoMaior">Total em Produtos:</strong></div></td>
    <td><div align="left"><span class="textoMaior"><strong>
      <%
						Dim rsProdutosTotal, QTD_EM_PRODUTOS, XTOTAL_AGREGADOS, XTOTAL_ALBUNS, XTOTAL_PRODUTOS_A_PARTE
						AbrirConexao()
						
						QTD_EM_PRODUTOS = 0
						XTOTAL_AGREGADOS = 0
						XTOTAL_ALBUNS = 0
						XTOTAL_PRODUTOS_A_PARTE = 0
						XTOTAL_DESCONTO = 0
						
						'VW_PEDIDOS_PRODUTOS_TOTAL_AGREGADOS
						Set rsProdutosTotal = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_AGREGADOS WHERE COD_PEDIDO="&COD_PEDIDO)
						if not rsProdutosTotal.Eof Then
							XTOTAL_AGREGADOS = formatNumber(0&rsProdutosTotal("XTOTAL_AGREGADOS"))/1
						End If
						rsProdutosTotal.Close()
						Set rsProdutosTotal = Nothing
						
						'VW_PEDIDOS_PRODUTOS_TOTAL_ALBUNS
						Set rsProdutosTotal = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_ALBUNS WHERE COD_PEDIDO="&COD_PEDIDO)
						if not rsProdutosTotal.Eof Then
							XTOTAL_ALBUNS = formatNumber(0&rsProdutosTotal("XTOTAL_ALBUNS"))/1
						End If
						rsProdutosTotal.Close()
						Set rsProdutosTotal = Nothing
						
						'VW_PEDIDOS_PRODUTOS_TOTAL_PRODUTOS_A_PARTE
						Set rsProdutosTotal = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_PRODUTOS_A_PARTE WHERE COD_PEDIDO="&COD_PEDIDO)
						if not rsProdutosTotal.Eof Then
							XTOTAL_PRODUTOS_A_PARTE = formatNumber(0&rsProdutosTotal("XTOTAL_PRODUTOS_A_PARTE"))/1
						End If
						rsProdutosTotal.Close()
						Set rsProdutosTotal = Nothing

						set rsProdutosTotal = objConexao.Execute("select SUM(valor_total_com_desconto) as valor_total_com_desconto from(SELECT SUM(faturado * (valor_un-(valor_un * (1-(ifnull(aliquota_desconto_foto,0)/100))))) as valor_total_com_desconto FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO in (SELECT cod_formando FROM TB_PEDIDOS_FORMANDOS WHERE cod_pedido =  "&cod_pedido&") union all SELECT SUM(quantidade * (valor_un-(valor_un * (1-(ifnull(aliquota,0)/100))))) as valor_total_com_desconto  FROM TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO = "&cod_pedido&") as total_desconto")
						if not rsProdutosTotal.Eof then
							XTOTAL_DESCONTO = formatCurrency(0&rsProdutosTotal(0))/1
						end if
						rsProdutosTotal.Close
						set rsProdutosTotal = Nothing

						QTD_EM_PRODUTOS = formatCurrency(XTOTAL_AGREGADOS + XTOTAL_ALBUNS + XTOTAL_PRODUTOS_A_PARTE)
	
						%>
      					<%=formatCurrency(QTD_EM_PRODUTOS - XTOTAL_DESCONTO)%> <small>(<%=QTD_EM_PRODUTOS & " - "& formatCurrency(XTOTAL_DESCONTO)%>)</small></strong></span></div></td>
  </tr>
  <tr>
    <td class="textoMaior"><div align="left"><strong>Total em Forma de Pagamento:</strong></div></td>
    <td class="textoMaior"><div align="left"><strong><%=formatCurrency(VALOR_TOTAL)%></strong></div></td>
  </tr>
  <%
  	'Crédito de formando, quando há algum saldo residual por algum motivo vindo da formatura.
  	Dim TOTAL_CREDITO
  	sql =	" SELECT SUM(IFNULL(TB_FICHAS_IDENTIFICACAO.CREDITO_EM_VALOR,0)) AS TOTAL_CREDITO "&_
			" FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON "&_
			" TB_PEDIDOS_FORMANDOS.COD_FORMANDO =  "&_
			" TB_FICHAS_IDENTIFICACAO.COD_FORMANDO WHERE  "&_
			" TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&COD_PEDIDO&" GROUP BY TB_PEDIDOS_FORMANDOS.COD_PEDIDO"
	Set rsCredito = objConexao.Execute(sql) 
	If not rsCredito.Eof Then
		TOTAL_CREDITO = formatNumber(0&rsCredito("TOTAL_CREDITO"))
	End If
	rsCredito.Close
	Set rsCredito = Nothing

  If TOTAL_CREDITO > 0 Then %>
  <tr>
    <td class="textoMaior"><div align="left"><strong>Total Cr&eacute;dito da formatura:</strong></div></td>
    <td class="textoMaior"><div align="left"><strong><%=formatCurrency(TOTAL_CREDITO)%></strong></div></td>
  </tr>
  <tr bgcolor="#E2E2E2">
    <td bgcolor="#E2E2E2" class="textoMaior"><div align="left"><strong>Total em Pagamento + Cr&eacute;dito:</strong></div></td>
    <td class="textoMaior"><div align="left"><strong><%=formatCurrency(VALOR_TOTAL+TOTAL_CREDITO)%></strong></div></td>
  </tr>
  <%
  End If
  
FecharConexao()
  %>
</table>
					<%Else%>
                    	 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr>
    <td width="279"><div align="left"><strong class="textoMaior">Total em Produtos:</strong></div></td>
    <td><div align="left"><span class="textoMaior"><strong>
      <%
						AbrirConexao()
						
						QTD_EM_PRODUTOS = 0
						
							Set rsProdutosTotal = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_PRODUTOS WHERE COD_PEDIDO="&COD_PEDIDO)
							
								if not rsProdutosTotal.Eof Then
									

									XTOTAL_PRODUTOS_A_PARTE = formatNumber(0&rsProdutosTotal("XTOTAL_PRODUTOS_A_PARTE"))/1
									
									QTD_EM_PRODUTOS = formatCurrency(XTOTAL_PRODUTOS_A_PARTE)
									
								Else
									
									QTD_EM_PRODUTOS = 0
									
								End If
							
							
							rsProdutosTotal.Close()
							Set rsProdutosTotal = Nothing
							
							set rsProdutosTotal = objConexao.Execute("select SUM(valor_total_com_desconto) as valor_total_com_desconto from(SELECT CASE WHEN "&PEDIDO_EXTRA&" = 2 THEN SUM(faturado * (valor_un-(valor_un * (1-(ifnull(aliquota_desconto_foto,0)/100))))) ELSE '0' END as valor_total_com_desconto FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO in (SELECT cod_formando FROM TB_PEDIDOS_FORMANDOS WHERE cod_pedido =  "&cod_pedido&") union all SELECT SUM(quantidade * (valor_un-(valor_un * (1-(ifnull(aliquota,0)/100))))) as valor_total_com_desconto  FROM TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO = "&cod_pedido&") as total_desconto")
							if not rsProdutosTotal.Eof then
								XTOTAL_DESCONTO = formatCurrency(0&rsProdutosTotal(0))/1
							end if
							rsProdutosTotal.Close
							set rsProdutosTotal = Nothing


						%>
      <%=formatCurrency(QTD_EM_PRODUTOS - XTOTAL_DESCONTO)%></strong></span></div></td>
  </tr>
  <tr>
    <td class="textoMaior"><div align="left"><strong>Total em Forma de Pagamento:</strong></div></td>
    <td class="textoMaior"><div align="left"><strong><%=formatCurrency(VALOR_TOTAL)%></strong></div></td>
  </tr>
 <%
  
FecharConexao()
  %>
</table>
                    <%End If%>

<div align="center">
                        <p>
						<form action="lca_pedidos_pendentes_alteracoes_part_4.asp?cod_pedido=<%=COD_PEDIDO%>" method="post">
                          <input type="button" class="btnVoltar" onclick="abre('https://stillos.com.br/gerencial_foto/lca_pedidos_pendentes_alteracoes_part_2.asp?cod_pedido=<%=COD_PEDIDO%>')" name="aaa2" value=" " /> 
                          &nbsp;
						  <%If VERIFICA_SE_HA_PENDENCIAS = 1 Then%>
                          <input type="submit" class="btnAvancar" value=" " />
						  <input name="VALIDADOR" value="1" type="hidden" />
                          <%Else%>
                          <input name="VALIDADOR" value="2" type="hidden" />
						   <input type="button" class="btnAvancar" onclick="alert('Atenção há itens não verificados, favor verificar antes de continuar!')"  value=" " />
						   <%End IF%>
					<input type="hidden" name='codpedido' id='cod_pedido_boleto' value="<%=cod_pedido%>">
					</form>
                        </p>
</div>
					
</td>

    </tr>
</table>

 
  
  </td>
    </tr>
</table>

  
  </td>
  </tr>
</table>
<!-- #include file="includes/rodape.asp" -->

  </td>
</tr>
</table>

<script type="text/javascript" src='js/jquery-mask.js'></script>
<script>
	$(function(){
		function validaFormulario(tipo)
	    {
	      var validado = true;
	      var campos = '';
	      //alert("."+tipo+"[required=required]")
	      $("."+tipo+"[required=required]").each(function(){
	        if($(this).val() == "")
	        {
	          validado = false;
	          campos += $(this).attr("nome-campo") +   "\n";
	        }
	      }); 
	      if (campos.length > 0) {
	        alert(campos);
	      }
	      
	      return validado;
	    }

		//botoes
		$(".excluir_boletos_sacados").click(function(){
			$(".excluir_confirmacao").removeClass("hidden");
			$(".alterar_boletos_form").addClass("hidden");
			$(".alterar_sacados_form").addClass("hidden");
		});

		$(".alterar_boletos").click(function(){
			$(".excluir_confirmacao").addClass("hidden");
			$(".alterar_boletos_form").removeClass("hidden");
			$(".alterar_sacados_form").addClass("hidden");
		});

		$(".alterar_sacados").click(function(){
			$(".excluir_confirmacao").addClass("hidden");
			$(".alterar_boletos_form").addClass("hidden");
			$(".alterar_sacados_form").removeClass("hidden");
		});

		$(".cancelar").click(function(){
			$(".excluir_confirmacao").addClass("hidden");
			$(".alterar_boletos_form").addClass("hidden");
			$(".alterar_sacados_form").addClass("hidden");
		});
		//fim botoes

		$("#confirma_exclusao").click(function(){
			var request = $.ajax({
			  url: "ajax_boletos.asp",
			  type: "POST",
			  data: { tp : "excluirboletosSacados", p1 : $("#cod_pedido_boleto").val() },
			  dataType: "html",
		  	  beforeSend : function(){
			  	disableInputs();
			  }
			});
			 
			request.done(function( msg ) {
		    	enableInputs();
			    location.reload();
			});
			 
			request.fail(function( jqXHR, textStatus ) {
			 	$(".cancelar").trigger("click");
			 	enableInputs();
			 	alert("houve um erro ao tentar executar a operação");	
			});
		});

		$("#confirma_alterarcao_boleto").click(function(){
			if (validaFormulario("vencimento")) 
			{ 
				var request = $.ajax({
				  url: "ajax_boletos.asp",
				  type: "POST",
				  data: { tp : "alteraParcelamento", p1 : $("#cod_pedido_boleto").val(), p2 : $("#dia_vencimento_id").val() , p3 : $("#mes_vencimento_id").val() , p4 : $("#ano_vencimento_id").val() , p5 : $("select[name=qtd_parcelas]").val()  },
				  dataType: "html",
				  beforeSend : function(){
				  	disableInputs();
				  }
				});
				 
				request.done(function( msg ) {
					enableInputs();
				    location.reload();
				});
				 
				request.fail(function( jqXHR, textStatus ) {
				 	$(".cancelar").trigger("click");
				 	alert("houve um erro ao tentar executar a operação");
				 	enableInputs();
				});
			}
		});

		$("#confirma_alterarcao_sacado").click(function(){
			if (validaFormulario("sacado")) 
			{ 
				var request = $.ajax({
				  url: "ajax_boletos.asp",
				  type: "POST",
				  data: { tp : "alteraSacado", 
				  			p1 : $("#cod_sacado_id").val(), 
				  			p2 : $("#cod_formando_id").val(), 
				  			p3 : $("#nome_formando_id").val(), 
				  			p4 : $("#cpf_id").val(), 
				  			p5 : $("#email_id").val(), 
				  			p6 : $("#cep_id").val(), 
				  			p7 : $("#endereco_id").val(), 
				  			p8 : $("#numero_id").val(), 
				  			p9 : $("#bairro_id").val(), 
				  			p10 : $("#cidade_id").val() , 
				  			p11 : $("#estado_id").val() , 
				  			p12 : $("#grau_parentesco_id").val(),
				  		},
				  dataType: "html",
				  beforeSend : function(){
				  	disableInputs();
				  }
				});
				 
				request.done(function( msg ) {
				    location.reload();
				    enableInputs();
				});
				 
				request.fail(function( jqXHR, textStatus ) {
				 	$(".cancelar").trigger("click");
				 	enableInputs();
				 	alert("houve um erro ao tentar executar a operação");
				});

			}
		});

	    function disableInputs(){
	      $("input").prop('disabled', true);
	      $("select").prop('disabled', true);
	    }

	    function enableInputs(){
	      $("input").prop('disabled', false);
	      $("select").prop('disabled', false);
	    }

		function validaVencimento(){
	      var dia = $("#dia_vencimento_id").val();
	      var mes = $("#mes_vencimento_id").val();
	      var ano = $("#ano_vencimento_id").val();

	      console.log('dia : '+ dia);
	      console.log('mes : '+ mes);
	      console.log('ano : '+ ano);

	      if (dia != '' && mes != '' && ano != '')
	      {
	        //validamos via AJAX
	        var request = $.ajax({
	          url: "ajax_boletos.asp",
	          type: "POST",
	          data: { tp : 'valida_vencimento', p1 : dia, p2 : mes, p3 : ano },
	          dataType: "text",
	          beforeSend : function(){
	            $("#carregando_vencimento").removeClass("hidden");
	            disableInputs();
	          }
	        });

	        request.done(function( msg ) {
	          if(msg.toLowerCase() == 'false')
	          { 
	            alert('Data de vencimento inválida');
	            $(".valida_vencimento").val('');
	          }
	          $("#carregando_vencimento").addClass("hidden");
	          enableInputs();
	          $("select[name=qtd_parcelas]").focus();
	          $("#vencimento_validado").val(msg.toLowerCase());
	        });

	        request.fail(function( jqXHR, textStatus ) {
	          $("#carregando_vencimento").addClass("hidden");
	          enableInputs();
	        });
	      }
	    }
		$(".valida_vencimento").blur(function(){validaVencimento()});

		 //--------------------------------------------------------------
    // CEP
    //--------------------------------------------------------------
    $('#buscar_cep_id').click(function(){
      var cep_ultima_pesquisa = $("#cep_ultima_pesquisa").val();

      if (cep_ultima_pesquisa != $("#cep_id").val())
      {
        $("#cep_ultima_pesquisa").val($("#cep_id").val());
        var request = $.ajax({
          url: "busca_cep.php",
          type: "POST",
          data: { cep : $('#cep_id').val() },
          dataType: "text",
          beforeSend : function(){
            $("#carregandoa").removeClass("hidden");
            $("#buscar_cep_id").addClass("hidden");
            disableInputs();
          }
        });

        request.done(function( msg ) {
          if(msg == '')
          {
            alert('CEP Não localizado');
            $("#carregandoa").addClass("hidden");
            $("#buscar_cep_id").removeClass("hidden");
            enableInputs();
            $('input[name=ENDERECO]').focus();
          }
          else
          {
            $("#carregandoa").addClass("hidden");
            $("#buscar_cep_id").removeClass("hidden");
            enableInputs();
            cep = JSON.parse(msg);
            $('input[name=ENDERECO]').val(cep['logradouro']);
            $('input[name=BAIRRO]').val(cep['bairro']);
            $('input[name=CIDADE]').val(cep['cidade']);
            $('input[name=ESTADO]').val(cep['estado']);
            $('input[name=NUMERO]').focus();
          }
        });

        request.fail(function( jqXHR, textStatus ) {
          $("#buscar_cep_id").removeClass("hidden");
          $("#carregandoa").addClass("hidden");
          enableInputs();
        });
      }
    });

    $("#cod_formando_id").change(function(){
      if ($(this).val() == "#outro" ) {
        $("#nome_formando_id").removeClass('hidden');
        $(".oculta_formando").removeClass('hidden');

       $(".required_form").attr("required","");
      }
      else
      {
        $("#nome_formando_id").addClass('hidden');
        $(".oculta_formando").addClass('hidden');

        
         $(".required_form").removeAttr("required","");
      }
    });
  


    //--------------------------------------------------------------
    //PADRAO
    //--------------------------------------------------------------
    $(".padrao_cep").mask("99999-999"); 
    $(".padrao_data").mask("99/99/9999" );  
    $(".padrao_cpf").mask("999.999.999-99" );

    console.log('mascaras carregadas');



    $('.padrao_cpf').blur(function(){
        var cpfv;
        var erro = new Array();
        var d = new Array();
        var s = new Number;
        var c = 10;
        var s = 0;
        
        cpfv = $(this).val();
        //cpfv = cpfv.replace(/\D/,'');
        cpfv = cpfv.replace('-','');
        cpfv = cpfv.replace('_','');
        cpfv = cpfv.replace('.','');
        cpfv = cpfv.replace('.','');
        
        
      if (cpfv !== '')
      {
        
        if (cpfv.length === 11)
        {
          if (cpfv == "00000000000" || cpfv == "11111111111" || cpfv == "22222222222" || cpfv == "33333333333" || cpfv == "44444444444" || cpfv == "55555555555" || cpfv == "66666666666" || cpfv == "77777777777" || cpfv == "88888888888" || cpfv == "99999999999") {
              erro[0] = true;
              erro[1] = "Número de CPF inválido!";
          }
          
          for(var i=0; i<9; i++){
            d[i] = cpfv.charAt(i);
            s += (d[i] * c);
            c--;
          }   
          
          x = s % 11
          if(x < 2){
            d[9] = 0
          }else{
            d[9] = 11-x
          }
          
          s = 0;
          c = 11;
          for(y=0; y<10; y++) 
          {
            s += (d[y] * c--);
          }
          
          x = s % 11      
          if(x < 2){
            d[10] = 0;
          }else{
            d[10] = 11-x;
          }
    
          if((cpfv.charAt(9) != d[9]) || (cpfv.charAt(10) != d[10])){
            erro[0] = true;
            erro[1] = "Número de CPF inválido!";
          }
          
          if (erro[0] === true)
          {
            alert(erro[1]);
            $(this).val('');    
          }
        }
        else
        {
          alert("Número de CPF inválido!");
        }
        
      }
      });
    

	});
</script>


</body>
</html>