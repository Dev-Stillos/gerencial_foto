<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'######################################################################
'INSTRU��ES CORRE��O DE PEDIDOS
'1) Lan�amento da DATA da VENDa // Muito importante, pois � calculado dados como quando o vendedores receber�
'2) Corre��o de Produtos Lan�ados e Calculo se o �lbum � repasse ou n�o, caso for ser� concedito um acr�scimo na comiss�o totalizando 10% do valor do recibo, independente se o pedido n�o conter somente o �lbum
'3) Corre��o de Forma de Pagamento
'4) Inser��o de Ajuda de Custo, Visualiza��o de Quando o vendedor receber�, Observa��o
'5) Conclus�o confirma��o e finaliza��o
'######################################################################
%>
<%
Dim COD_PEDIDO, DATA_VENDA, PEDIDO_EXTRA, QTD_PARCELAS, COD_GRUPO_VENDEDORES, COD_VENDEDOR, FINALIZA_DIRETORIA, DATA_DESCEU_DIRETORIA
COD_PEDIDO = Request.QueryString("cod_pedido")

	'///////////
	'DATA DA VENDA
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT DATA_VENDA,PEDIDO_EXTRA, QTD_PARCELAS,COD_VENDEDOR, FINALIZA_DIRETORIA, DATA_DESCEU_DIRETORIA FROM TB_PEDIDOS WHERE COD_PEDIDO="&COD_PEDIDO)
				
				If not rs.Eof Then
					DATA_VENDA=rs("DATA_VENDA")
					PEDIDO_EXTRA=rs("PEDIDO_EXTRA")
					QTD_PARCELAS=formatNumber(0&rs("QTD_PARCELAS"),0)
					COD_VENDEDOR=rs("COD_VENDEDOR")
					FINALIZA_DIRETORIA = rs("FINALIZA_DIRETORIA")
					DATA_DESCEU_DIRETORIA = Cstr(""&rs("DATA_DESCEU_DIRETORIA"))
				End If
				
		rs.Close
		Set rs=Nothing
		
		Set rs=objConexao.Execute("SELECT COD_GRUPO_VENDEDORES FROM TB_VENDEDORES WHERE COD_VENDEDOR="&COD_VENDEDOR)
				
				If not rs.Eof Then
					COD_GRUPO_VENDEDORES = rs("COD_GRUPO_VENDEDORES")
				End If
				
		rs.Close
		Set rs=Nothing
		
	FecharConexao()
	'///////////

'Para tirar a valida��o da p�gina basta tirar o "#" ou coloca-lo para habilitar antes do include
'///////////////////////////////////////////
'Grava??o de descontos recebidos do popup de desconto
If Request.Form("GRAVAR_DESCONTO") <> "" Then
    Dim codFormando, desconto_album, desconto_produto
    codFormando = Request.Form("COD_FORMANDO")
    desconto_album = Replace(Trim(Request.Form("DESC_ALBUM")), ",", ".")
    desconto_produto = Replace(Trim(Request.Form("DESC_PRODUTO")), ",", ".")

    If codFormando <> "" And IsNumeric(desconto_album) And IsNumeric(desconto_produto) Then
        AbrirConexao()

        ' Atualiza o desconto de ?lbum para o formando selecionado
        objConexao.Execute "UPDATE TB_FICHAS_IDENTIFICACAO SET aliquota_desconto_foto = " & desconto_album & _
                           " WHERE COD_FORMANDO = " & codFormando

        ' Atualiza o desconto dos produtos vendidos ? parte apenas para o formando no pedido
        objConexao.Execute "UPDATE TB_PEDIDOS_PRODUTOS SET aliquota = " & desconto_produto & _
                           " WHERE COD_PEDIDO = " & COD_PEDIDO & " AND COD_FORMANDO = " & codFormando

        FecharConexao()
        Response.Redirect "lca_pedidos_pendentes_alteracoes_part_2.asp?cod_pedido=" & COD_PEDIDO & "&msg=Desconto salvo para o formando!"
    Else
        Response.Redirect "lca_pedidos_pendentes_alteracoes_part_2.asp?cod_pedido=" & COD_PEDIDO & "&msg=Erro: Preencha todos os campos corretamente."
    End If
End If


'Grava lan�amentos � parte
If Request.Form("FORM_COD_FORMANDO") <> Empty Then
	'Gravar produtos � parte
	Dim FORM_COD_FORMANDO, FORM_COD_PROD, FORM_QTD_VENDIDA, FORM_VALOR_UN
	FORM_COD_FORMANDO	= Request.Form("FORM_COD_FORMANDO")
	FORM_COD_PROD		= Request.Form("FORM_COD_PROD") 
	FORM_QTD_VENDIDA	= Request.Form("FORM_QTD_VENDIDA") 
	FORM_VALOR_UN		= Request.Form("FORM_VALOR_UN")
	
	AbrirConexao()
	objConexao.Execute( "INSERT INTO TB_PEDIDOS_PRODUTOS ( "&_
						"COD_PEDIDO , "&_
						"COD_FORMANDO , "&_
						"COD_PRODUTO , "&_
						"VALOR_UN , "&_
						"QUANTIDADE "&_
						") "&_
						"VALUES ("&_
						COD_PEDIDO&","&FORM_COD_FORMANDO&", "&FORM_COD_PROD&", '"&sqlNumero(FORM_VALOR_UN)&"', "&FORM_QTD_VENDIDA&_
						")")
	
	FecharConexao()
	Response.Redirect("lca_pedidos_pendentes_alteracoes_part_2.asp?cod_pedido="&COD_PEDIDO&"&msg=Produto Lan&ccedil;ado com sucesso!")
End If

'/////////////////////////////////////////////////////
'Excluir produto extra
If Request.QueryString("prod_excluir_a_parte") <> Empty then
	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO_PRODUTO = "&Request.QueryString("prod_excluir_a_parte")&" LIMIT 1")
	FecharConexao()
	Response.Redirect("lca_pedidos_pendentes_alteracoes_part_2.asp?cod_pedido="&COD_PEDIDO&"&msg=Produto exclu&iacute;do com sucesso!")
End If
'//////////////////////////////////////////////////////////////////////////
'Produto agregado, salvo individualmente
If Request.QueryString("vendido") <> Empty Then
'Grava o produto agregado
Dim valor_un_a, vendido_a
'Caso for vendido igual a 2 ent�o faz o update com 0 no valor unit�rio

vendido_a	= Request.QueryString("vendido")

	If vendido_a = 1 Then
		valor_un_a  = sqlNumero(Request.QueryString("valor_un"))
	Else
		valor_un_a  = 0
	End If



	AbrirConexao()
		objConexao.Execute( " UPDATE TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS SET "&_
							" VALOR_UN = '"&valor_un_a&"', "&_
							" VENDIDO  = "&vendido_a&" WHERE COD_PRODUTO_AGREGADO = "&Request.QueryString("id")&" LIMIT 1")
	FecharConexao()
	Response.Redirect("lca_pedidos_pendentes_alteracoes_part_2.asp?cod_pedido="&COD_PEDIDO&"&msg=Produto agregado salvo com sucesso!")
End If

If DATA_DESCEU_DIRETORIA <> "" and Session("cod_usuario") <> 16 Then
	Response.Redirect("adm_index.asp?codmsg=1")
End If
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>Pedidos Pendentes</title>
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
	
	
	
	<script type="text/javascript">
messageObj = new DHTML_modalMessage();	// We only create one object of this class
messageObj.setShadowOffset(5);	// Large shadow


function displayMessage(url)
{
	messageObj.setSource(url);
	messageObj.setCssClassMessageBox(false);
	messageObj.setSize(400,200);
	messageObj.setShadowDivVisible(true);	// Enable shadow for these boxes
	messageObj.display();
}

function displayStaticMessage(messageContent,cssClass)
{
	messageObj.setHtmlContent(messageContent);
	messageObj.setSize(670,200);
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
	<!--AJUDA F1 E CTRL PARA BUSCA R�PIDA-->
		
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
		<!-- CALEND�RIO SCRIPTS
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
 function confirmaAgregados(COD_FORMANDO,I,COD_AGREGADO){
  		d = document["AGREGADOS_"+COD_FORMANDO]
 		if(d["AGREGADOS_VENDIDO_"+I+"_"+COD_FORMANDO].value == ""){
			alert('� necess�rio selecionar se foi vendido o produto agregado!');
		   d["AGREGADOS_VENDIDO_"+I+"_"+COD_FORMANDO].focus();
		   return false;
		}
		if( d["AGREGADOS_VENDIDO_"+I+"_"+COD_FORMANDO].value == 1  && (d["AGREGADOS_VALOR_UN_"+I+"_"+COD_FORMANDO].value == "" || d["AGREGADOS_VALOR_UN_"+I+"_"+COD_FORMANDO].value == 0)){
			alert('� necess�rio informar o valor vendido do produto agregado!');
		   d["AGREGADOS_VALOR_UN_"+I+"_"+COD_FORMANDO].focus();
			return false;
		}
		
		var valor_un = d["AGREGADOS_VALOR_UN_"+I+"_"+COD_FORMANDO].value
		var vendido  = d["AGREGADOS_VENDIDO_"+I+"_"+COD_FORMANDO].value
			
		location.href = "?cod_pedido=<%=COD_PEDIDO%>&valor_un="+valor_un+"&vendido="+vendido+"&id="+COD_AGREGADO
		return true;
		
 		
 }
function resetarAlbum(cod_for){
	var caminho
	caminho = "<table cellpadding=\'0\' cellspacing=\'0\' border=\'0\' width=\'100%\'><tr><td><table cellpadding=\'0\' cellspacing=\'0\' width=\'100%\'><tr><td width=\'70%\'><h1>Altera&ccedil;&atilde;o de &Aacute;lbum</h1></td><td align=\'right\'><a href=\'#\' onclick=\'closeMessage();return false\'>Fechar Janela</a></td></tr></table></td></tr><tr><td><iframe frameborder=\'0\' scrolling=\'no\' width=\'99%\' src=\'lca_pedidos_pendentes_alteracoes_part_2_POPUP_album.asp?cod_pedido=<%=COD_PEDIDO%>&cod_for="+cod_for+"\'></iframe></td></tr></table>";
	displayStaticMessage(caminho,false)		
}
</script>
</head>
<body>

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
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">

<% If validaUsuario(2,"",85,false) Then %>
<div style="margin-bottom: 20px; display: flex; justify-content: center;">
  <div style="position: relative;">
    <button onclick="toggleDropdown()" style="background-color:#ff007f; color:white; padding: 10px 20px; border:none; border-radius: 6px; cursor:pointer; display:flex; align-items:center; gap:10px;">
      <strong>LCA - Pedidos Pendentes / Corre&ccedil;&atilde;o</strong> | Sub-Menu
      <span style="font-size:12px;">&#9660;</span>
    </button>

    <!-- SUBMENU -->
    <div id="submenuDropdown" style="display:none; position:absolute; background-color:#fff; width:100%; box-shadow:0px 8px 16px rgba(0,0,0,0.2); z-index:1000; border-radius:6px; text-align:center;">
      
      <!-- Bot?o Cadastrar Produto -->
      <a href="#"
         style="display:inline-block; background-color:#ff007f; color:white; padding:10px 20px; margin:10px auto; text-decoration:none; border-radius:6px; font-weight:bold;"
         onclick="abrirJanelaProduto(<%=COD_PEDIDO%>); return false;">
        Cadastrar Produto
      </a>

      <!-- Bot?o Cadastrar Desconto -->
      <a href="#"
         style="display:inline-block; background-color:#ff007f; color:white; padding:10px 20px; margin:10px auto; text-decoration:none; border-radius:6px; font-weight:bold;"
         onclick="abrirJanelaDesconto(<%=COD_PEDIDO%>); return false;">
        Cadastrar Desconto
      </a>

    </div>
  </div>
</div>
<% End If %>

<!-- SCRIPT para Submenu e Modais -->
<script>
function toggleDropdown() {
  var menu = document.getElementById("submenuDropdown");
  menu.style.display = (menu.style.display === "block") ? "none" : "block";
}

function abrirJanelaProduto(codPedido) {
  const html = `
    <table cellpadding='0' cellspacing='0' border='0' width='100%'>
      <tr>
        <td>
          <table cellpadding='0' cellspacing='0' width='100%'>
            <tr>
              <td width='70%'><h1>Inclus&atilde;o de Produto</h1></td>
              <td align='right'>
                <a href='#' onclick='closeMessage();return false'>Fechar Janela</a>
              </td>
            </tr>
          </table>
        </td>
      </tr>
      <tr>
        <td>
          <iframe frameborder='0' scrolling='auto' width='100%' height='150px'
            src='lca_pedidos_pendentes_alteracoes_part_2_POPUP_inc_prod.asp?cod_pedido=${codPedido}'>
          </iframe>
        </td>
      </tr>
    </table>
  `;
  displayStaticMessage(html, false);
}

function abrirJanelaDesconto(codPedido) {
  const html = `
    <table cellpadding='0' cellspacing='0' border='0' width='100%'>
      <tr>
        <td>
          <table cellpadding='0' cellspacing='0' width='100%'>
            <tr>
              <td width='70%'><h1>Inclus&atilde;o de Desconto</h1></td>
              <td align='right'>
                <a href='#' onclick='closeMessage();return false'>Fechar Janela</a>
              </td>
            </tr>
          </table>
        </td>
      </tr>
      <tr>
        <td>
          <iframe frameborder='0' scrolling='auto' width='100%' height='150px'
            src='lca_pedidos_pendentes_alteracoes_part_2_POPUP_inc_desc.asp?cod_pedido=${codPedido}'>
          </iframe>
        </td>
      </tr>
    </table>
  `;
  displayStaticMessage(html, false);
}
</script>



<div class="msg" align="center"><%=Request.QueryString("msg")%></div>

	  	  
		  	<h3 align="center">Corre&ccedil;&atilde;o de Pedido - Produtos Cadastrados </h3>
			<table width="98%" align="center" 	cellpadding="0" cellspacing="3">
			<tr>
            <td width="12%"><div align="left"><strong>Vendedor</strong>:</div></td>
            <td width="88%"><%
			AbrirConexao()
			'Nome do vendedor
			Dim rsVendedorQ
			Set rsVendedorQ = objConexao.Execute("SELECT TB_VENDEDORES.NOME_VENDEDOR FROM TB_VENDEDORES INNER JOIN TB_PEDIDOS ON TB_VENDEDORES.COD_VENDEDOR = TB_PEDIDOS.COD_VENDEDOR WHERE TB_PEDIDOS.COD_PEDIDO="&COD_PEDIDO&" LIMIT 1")
				
				If not rsVendedorQ.Eof Then
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
			<br />
  <script language="javascript">
  
  function validaFormAparte(){
  d = document.prod_parte
	 if (d.FORM_COD_FORMANDO.value == ""){
		alert('Selecione o formando correspondente ao produto vendido!')
	 	d.FORM_COD_FORMANDO.focus();
		return false;
	 }
	 if (d.FORM_COD_PROD.value == ""){
		alert('Selecione o produto do pedido!')
	 	d.FORM_COD_PROD.focus();
		return false;
	 }
	 //alert(qtd_vendida);
	 if (d.FORM_QTD_VENDIDA.value == "" || d.FORM_QTD_VENDIDA.value == 0 || d.FORM_QTD_VENDIDA.value < 0){
		alert('Digite a quantidade vendida!')
	 	d.FORM_QTD_VENDIDA.focus();
		return false;
	 }
	  if (d.FORM_VALOR_UN.value == "" || d.FORM_VALOR_UN.value == 0 || d.FORM_VALOR_UN.value < 0){
		alert('Digite o valor do produto!')
	 	d.FORM_VALOR_UN.focus();
		return false;
	 }  
  }
  </script>
			
			
			
			<%
			  '1) Seleciona os c�digos dos formandos do pedido
			  '2) Seleciona os dados inseridos nas fichas
			  AbrirConexao()
			  
			  Dim sql, rs
			  
		sql = " SELECT TB_FICHAS_IDENTIFICACAO.COD_CONTRATO,TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO, TB_FICHAS_IDENTIFICACAO.CREDITO_EM_VALOR, TB_FICHAS_IDENTIFICACAO.CREDITO_MOTIVO, TB_FICHAS_IDENTIFICACAO.CREDITO_DATA_INSERIDO, TB_FICHAS_IDENTIFICACAO.FATURADO, TB_FICHAS_IDENTIFICACAO.VALOR_UN, TB_FICHAS_IDENTIFICACAO.ALBUM, TB_FICHAS_IDENTIFICACAO.INICIAL, TB_FICHAS_IDENTIFICACAO.BRINDE, TB_FICHAS_IDENTIFICACAO.aliquota_desconto_foto, "&_
					" TB_FICHAS_IDENTIFICACAO.PRODUTO_AGREGADO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, TB_FICHAS_IDENTIFICACAO.COMISSAO, "&_
					" TB_PEDIDOS_FORMANDOS.COD_FORMANDO "&_
					" FROM TB_FICHAS_IDENTIFICACAO INNER JOIN TB_PEDIDOS_FORMANDOS ON "&_
					" TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = TB_PEDIDOS_FORMANDOS.COD_FORMANDO "&_
					" WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&COD_PEDIDO
					
			  Set rs = objConexao.Execute(sql)
				Dim QTD_BRINDE, QTD_RETIRADA, CREDITO_EM_VALOR
				
													
					While not rs.Eof
					QTD_BRINDE			= 0
					QTD_RETIRADA		= 0
					CREDITO_EM_VALOR	= formatNumber(0&rs("CREDITO_EM_VALOR"))
					Dim rsProducao, rsBrindesAlbum
				
				'Consulta a produ��o desta ficha para verificar qual o c�digo padr�o do �lbum
				Set rsProducao = objConexao.Execute("SELECT COD_FOTO_PADRAO FROM TB_PRODUCAO WHERE COD_PRODUCAO="&rs("COD_PRODUCAO"))
						If not rsProducao.Eof Then				
						COD_FOTO_PADRAO = rsProducao("COD_FOTO_PADRAO")
						Else
						COD_FOTO_PADRAO = 2
						End If
				rsProducao.Close()
				Set rsProducao = Nothing
				
				'Consulta a quantidade de brinde, agora com a foto padr�o do �lbum
				If COD_FOTO_PADRAO > 0 Then
					Set rsBrindesAlbum = objConexao.Execute("SELECT Sum(QUANTIDADE) as QUANTIDADE FROM TB_FICHAS_IDENTIFICACAO_BRINDES WHERE COD_FORMANDO="&rs("COD_FORMANDO")&" AND COD_PRODUTO="&COD_FOTO_PADRAO&" GROUP BY COD_PRODUTO")
				
					If not rsBrindesAlbum.Eof then
						QTD_BRINDE = formatNumber(0&rsBrindesAlbum("QUANTIDADE"),0)
					Else
						QTD_BRINDE = 0
					End If
					
				
				rsBrindesAlbum.Close()
				Set rsBrindesAlbum = Nothing
				End If				
			%>  
			<table width="98%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="#E4E4E4">
    <td width="100%">
	  <div align="left" class="textoMedio" style="border-bottom:1px #999999 solid"><strong><%=rs("NOME_FORMANDO")%>
	      <%If rs("COMISSAO") = 1 Then %>
	      - Membro de Comiss&atilde;o de Formatura
          <%end if%>
	    </strong></div>
	  <div align="center" class="style3" style="padding:3px">
	    <div align="left">C&oacute;d. Form: <%=digito(rs("COD_FORMANDO"))%> - &Aacute;lbum: <%=rs("ALBUM")%></div>
	  </div></td>
    </tr>
<%If int(PEDIDO_EXTRA) = 2 Then%>
		<%If CREDITO_EM_VALOR > 0 Then%>
		<tr bgcolor="whitesmoke">
  <td><div align="center"><strong>Cr&eacute;dito da Formatura </strong></div></td>
</tr>
<tr>
  <td><table width="98%" border="1" align="center" 	cellpadding="2" cellspacing="2" style="border-collapse:collapse">
    <tr>
      <td width="21%"><div align="center"><strong>Valor Cr&eacute;dito </strong></div></td>
      <td width="53%"><div align="center"><strong>Motivo</strong></div></td>
      <td width="26%"><div align="center"><strong>Data inserido no sistema </strong></div></td>
      </tr>
    
    <tr>
      <td><div align="center"><%=formatCurrency(CREDITO_EM_VALOR)%></div></td>
      <td><div align="left"><%=rs("CREDITO_MOTIVO")%></div></td>
      <td><div align="center"><%=rs("CREDITO_DATA_INSERIDO")%></div></td>
      </tr>
  </table></td>
</tr>
	<%End If%>

<tr bgcolor="whitesmoke">
    <td><div align="center"><strong>Dados do &Aacute;lbum </strong></div></td>
    </tr>
<tr>
  <td><table width="98%" border="1" align="center" cellpadding="2" cellspacing="2" style="border-collapse:collapse">
    <tr>
      <td width="8%"><div align="center"><strong>Qtd. Inicial</strong></div></td>
      <td width="7%"><div align="center"><strong>Qtd. Brinde</strong></div></td>
      <td width="7%"><div align="center"><strong>Qtd. Faturada</strong></div></td>
      <td width="7%"><div align="center"><strong>Qtd. Retirada </strong></div></td>
      <td width="9%"><div align="center"><strong>Valor Un.</strong></div></td>
      <td width="7%"><div align="center"><strong>Valor Total</strong></div></td>
      <td width="7%"><div align="center"><strong>Desc(%) <img src="../gerencial_foto/images/icons/icon_alterar.png" style="vertical-align:middle; cursor:pointer;" alt="Editar">.</strong></div></td>
      <td width="7%"><div align="center"><strong>Desc(un).</strong></div></td>
      <td width="7%"><div align="center"><strong>Desc(total).</strong></div></td>
      <td width="7%"><div align="center"><strong>Valor Un. Desc.</strong></div></td>
      <td width="7%"><div align="center"><strong>Valor Total. Desc.</strong></div></td>
      <td width="8%"><div align="center"><strong>Valor Padr&atilde;o</strong></div></td>
      <td width="8%"><div align="center"><strong>Status</strong></div></td>
      <td width="3%"><div align="center"><strong>Altera&ccedil;&otilde;es</strong></div></td>
    </tr>
    <tr>
      <td><div align="center"><%=rs("INICIAL")%></div></td>
      <td><div align="center"><%=QTD_BRINDE%></div></td>
      <td><div align="center">
		  <%
			  Dim  rsVALOR_UN, rsFATURADO, rsVALOR_TOTAL
			  rsVALOR_UN = formatNumber(0&rs("VALOR_UN"))
			  rsFATURADO = formatNumber(0&rs("FATURADO"),0)
			  rsVALOR_TOTAL = rsVALOR_UN * rsFATURADO
			  
			  QTD_RETIRADA = (0&rs("INICIAL")) - (QTD_BRINDE/1 + rsFATURADO/1)
		  %>
      <%If rsFATURADO > 0 then %>
	  	<%=rsFATURADO%>
	  <%Else%>
	  	---
	  <%End If%>
	  </div></td>
      <td><div align="center"><%=QTD_RETIRADA%></div></td>
      <td><div align="center">
		  <%If rsVALOR_UN > 0 then %>
		  	<%=formatCurrency(rsVALOR_UN)%>
		  <%Else%>
		  	---
		  <%End If%>
	  </div></td>
      <td><div align="center">

        <%If rsVALOR_TOTAL > 0 then %>
			<%=formatCurrency(rsVALOR_TOTAL)%>
        <%Else%>
			---
		<%End If%>
	  </div></td>


	<td align="center"><%=rs("aliquota_desconto_foto")%>(%)</td>
	<%
		dim xdesconto_un, xdesconto_total, xaliquota
		xaliquota = formatNumber(0&rs("aliquota_desconto_foto"))
		xdesconto_un = formatCurrency(rsVALOR_UN) - (formatCurrency(rsVALOR_UN) * (1-(xaliquota/100)))
		xdesconto_total = xdesconto_un * int(0&rs("faturado"))
	%>
	<td align="center"><%=formatCurrency(xdesconto_un)%></td>
	<td align="center"><%=formatCurrency(xdesconto_total)%></td>
	<td align="center"><%= formatCurrency(rsVALOR_UN-xdesconto_un) %></td>
	<td align="center"><%= formatCurrency(rsVALOR_TOTAL-xdesconto_total) %></td>

      <td bgcolor="#FFFF99">
	    <div align="center">
	      <%
	  valor_padrao = 0	  
				  
	  If QTD_PARCELAS > 0 Then
	  	valor_padrao = valorProdutoGeral(COD_FOTO_PADRAO,rs("COD_CONTRATO"),QTD_PARCELAS,true,COD_GRUPO_VENDEDORES)
		
		if(valor_padrao) > 0 Then
			Response.Write(formatCurrency(valor_padrao))
		Else
			Response.Write("Produto n&atilde;o possui pre&ccedil;o padr&atilde;o ou boletim informativo")
		End If
		
	  Else
	   Response.Write("N&atilde;o &aacute; poss&iacute;vel verificar, pois o n&uacute;mero de parcelas &eacute; 0.")
	  End If
	  %>
	        </div></td>
      <td bgcolor="#FFFF99"><div align="center">
      <%
	  If valor_padrao > 0 and rsVALOR_UN > 0 Then
	  		
			diferenca = (rsVALOR_UN-valor_padrao)/1
			diferenca = formatNumber(diferenca)
			
			If diferenca >= 0 Then
				Response.Write("Correto")
			Else
				Response.Write("Valor inferior a tabela, n&atilde;o ser&aacute; poss&aacute;vel finalizar")
			End If
		
		
	  Else
	  	Response.Write("N&atilde;o &eacute; poss&iacute;vel verificar")
	  End If
	  %>
      </div></td>
      <td><div align="center">
	  <% if  validaUsuario(2,"",85, true) Then %>
	  <a href="javascript:resetarAlbum('<%=rs("COD_FORMANDO")%>')"><img src="../gerencial_foto/images/icons/icon_alterar.png" width="16" height="16" border="0" /></a>
	  <% end if %>
	  </div></td>
    </tr>

  </table></td>
</tr>
<%IF COD_FOTO_PADRAO > 0 Then%>
	<%
Dim rsBrindes

Set rsBrindes = objConexao.Execute( " SELECT Sum(TB_FICHAS_IDENTIFICACAO_BRINDES.QUANTIDADE) as QUANTIDADE, "&_
	" TB_PRODUTOS.NOME_PRODUTO FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_BRINDES "&_
	" ON TB_PRODUTOS.COD_PRODUTO = TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO "&_
	" WHERE TB_FICHAS_IDENTIFICACAO_BRINDES.COD_FORMANDO="&rs("COD_FORMANDO")&" AND TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO <> "& COD_FOTO_PADRAO&_
	" GROUP BY TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO ")
If not rsBrindes.Eof Then
%>
<tr bgcolor="whitesmoke">
  <td><div align="center"><strong>Outros Brindes</strong></div></td>
</tr>
<tr>
  <td><table width="98%" border="1" align="center" 	cellpadding="2" cellspacing="2" style="border-collapse:collapse">
    <tr>
      <td width="56%"><div align="center"><strong>Produto</strong></div></td>
      <td width="44%"><div align="center"><strong>Qtd</strong></div></td>
      </tr>
	   <%
	   While not rsBrindes.Eof
	   %>
    	<tr>
      <td><div align="left"><%=rsBrindes("NOME_PRODUTO")%></div></td>
      <td><div align="center"><%=rsBrindes("QUANTIDADE")%></div></td>
      </tr>
	   <%
	   rsBrindes.MoveNext()
	   Wend
	   %>
  </table></td>
</tr>
<%
End If

rsBrindes.Close()
Set rsBrindes = Nothing
%>
<%End If%>
<%
Dim rsAgregados
Set rsAgregados = objConexao.Execute(	" SELECT TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO, TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO_AGREGADO, TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.TIPO, "&_
										" Sum(TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.QUANTIDADE) as QUANTIDADE, "&_
										" TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.VENDIDO, TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.VALOR_UN, "&_
										" TB_PRODUTOS.NOME_PRODUTO FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS "&_
										" ON TB_PRODUTOS.COD_PRODUTO = TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO "&_
										" WHERE TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO="&rs("COD_FORMANDO")&_
										" GROUP BY TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO ")

If not rsAgregados.Eof Then
%>
<tr bgcolor="whitesmoke">
  <td><div align="center"><strong>Produtos Agregados com sa&iacute;da de produ&ccedil;&atilde;o</strong><br />
  </div></td>
</tr>
<tr>  <td>	
		<form style="margin:0px" method="post" name="AGREGADOS_<%=rs("COD_FORMANDO")%>">
			<table width="98%" border="1" align="center" cellpadding="2" cellspacing="2" style="border-collapse:collapse">
    <tr>
      <td width="35%"><div align="center"><strong>Produto</strong></div></td>
      <td width="20%"><div align="center"><strong>Qtd</strong>.</div></td>
      <td width="10%"><div align="center"><strong>Vendido</strong></div></td>
      <td width="10%"><div align="center"><strong>Valor Unit&aacute;rio </strong></div></td>
      <td width="8%"><div align="center"><strong>Valor Padr&atilde;o</strong></div></td>
      <td width="8%"><div align="center"><strong>Status</strong></div></td>
      <td width="9%"><div align="center"><strong>Altera&ccedil;&otilde;es</strong></div></td>
    </tr>
	<%
	Dim contProdAgregado, qteCod, qtaUi, FOI_BRINDE
	FOI_BRINDE = 2
	contProdAgregado = 1
	qteCod = formatNumber(0&Request.QueryString("prod_agre_alterar"),0)
	qtaUi  = formatNumber(0&rsAgregados("COD_PRODUTO_AGREGADO"),0)
	While not rsAgregados.Eof
		FOI_BRINDE = 2	
		'/////////////////
		'VERIFICA SE O PRODUTO AGREGADO N�O POSSUI O MESMO COMO BRINDE
		Set rs12=objConexao.Execute("SELECT COD_BRINDE FROM TB_FICHAS_IDENTIFICACAO_BRINDES WHERE COD_FORMANDO = "&rs("COD_FORMANDO")&" AND COD_PRODUTO="&rsAgregados("COD_PRODUTO"))
					
					If not rs12.Eof Then
					'&eacute BRINDE - SOMENTE ENTREGA
						FOI_BRINDE = 1
						'///////////////////////////////////////
							'1) Verifica se n�o estiver como baixado, baixa nas movimenta��es
							'2) Verifica se n�o estiver como Vendido + Foi Brinde faz os updates
							
							'1//
							sql =	" UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET DATA_RETORNO = "&sqlData(DATA_VENDA)&_
									" WHERE COD_FORMANDO = "&rs("COD_FORMANDO")&" AND TIPO = "&rsAgregados("TIPO")&" AND DATA_RETORNO IS NULL"
									
							objConexao.Execute(sql)	
							
							'2//
							'N�o tem problema fazer mais que uma vez o mesmo Update para caso consultar novamnte a mesma p�gina!
							sql =	" UPDATE TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS SET VENDIDO = 1, FOI_BRINDE = 1, VALOR_UN=0 WHERE "&_
									" COD_PRODUTO_AGREGADO = "&rsAgregados("COD_PRODUTO_AGREGADO")

							objConexao.Execute(sql)
						'///////////////////////////////////////
						
						
					End If
					
		rs12.Close
		Set rs12=Nothing			
					
		'/////////////////
		
		contProdAgregado = contProdAgregado + 1
	
	If FOI_BRINDE <> 1 Then
	%>
    <tr <%If qteCod = qtaUi Then%> bgcolor="#FFFF99" <%End If%>>
      <td><div align="left"><%=rsAgregados("NOME_PRODUTO")%></div></td>
      <td><div align="center"><%=rsAgregados("QUANTIDADE")%></div></td>
      <td><div align="center">
	  <%If rsAgregados("VENDIDO") = 3 then
	  CHAVE_AVANCAR = 2
	  %>
          <select name="AGREGADOS_VENDIDO_<%=contProdAgregado%>_<%=rs("COD_FORMANDO")%>">
            <option>---</option>
            <option value="1">Sim</option>
            <option value="2">N&atilde;o</option>
          </select>
      <%else%>
	  		<%If Request.QueryString("prod_agre_alterar") <> Empty Then
			CHAVE_AVANCAR = 2
			%>
          	<select name="AGREGADOS_VENDIDO_<%=contProdAgregado%>_<%=rs("COD_FORMANDO")%>">
				<option>---</option>
				<option <%If rsAgregados("VENDIDO") = 1 Then%> selected="selected"<%End If%> value="1">Sim</option>
				<option <%If rsAgregados("VENDIDO") = 2 Then%> selected="selected"<%End If%> value="2">N&atilde;o</option>
         	 </select>
			<%Else%>
				<%If rsAgregados("VENDIDO") = 1 Then%>
					Sim
				<%Else%>
					N&atilde;o
				<%End If%>	
	  		<%End If%>
	  <%End If%>
	  </div></td>
      <td>
	  <div align="center">
	  <%If Request.QueryString("prod_agre_alterar") <> Empty Then%>
			<input value="<%=retiraZero(formatNumber(0&rsAgregados("VALOR_UN")))%>" name="AGREGADOS_VALOR_UN_<%=contProdAgregado%>_<%=rs("COD_FORMANDO")%>" type="text" size="6" maxlength="5" />	
	  <%Else%>
			 <%If rsAgregados("VENDIDO") = 3 then%>
				  <input name="AGREGADOS_VALOR_UN_<%=contProdAgregado%>_<%=rs("COD_FORMANDO")%>" type="text" size="6" maxlength="5" />
			 	 <%elseif rsAgregados("VENDIDO") = 2 then%>
				  ---
			 	<%else%>
			  	<%=formatCurrency(0&rsAgregados("VALOR_UN"))%>
			  	<%End If%>
	  <%End If
	  
	  
	  %>
      </div>	  </td>
      <td bgcolor="#FFFF99"><div align="center">
	      <%
	  valor_padrao = 0	  
		  
	  If QTD_PARCELAS > 0 Then
	  	valor_padrao = valorProdutoGeral(rsAgregados("COD_PRODUTO_AGREGADO"),rs("COD_CONTRATO"),QTD_PARCELAS,true,COD_GRUPO_VENDEDORES)
		
		if(valor_padrao) > 0 Then
			Response.Write(formatCurrency(valor_padrao))
		Else
			Response.Write("Produto n&atilde;o possui pre&ccedil;o padr&atilde;o ou boletim informativo")
		End If
		
	  Else
	   Response.Write("N&atilde;o &eacute; poss&iacute;vel verificar, pois o n&uacute;mero de parcelas &eacute; 0.")
	  End If
	  %>
	        </div></td>
      <td bgcolor="#FFFF99"><div align="center">
      <%
	  valor_produto_agregado = formatNumber(0&rsAgregados("VALOR_UN"))
	  If valor_padrao > 0 and valor_produto_agregado > 0 Then
	  		
			diferenca = (rsAgregados("VALOR_UN")-valor_padrao)/1
			diferenca = formatNumber(diferenca)
			
			If diferenca >= 0 Then
				Response.Write("Correto")
			Else
				Response.Write("Valor inferior a tabela, n&atilde;o ser&aacute; poss&iacute;vel finalizar")
			End If
		
		
	  Else
	  	Response.Write("N&atilde;o &eacute; poss&iacute;vel verificar")
	  End If
	  %>
      </div></td>
      <td>
	  <%If rsAgregados("VENDIDO") = 3 Then%>
	  <div align="center">
	  <input name="SALVER_PROD" class="btnGravar" type="button" onclick="confirmaAgregados(<%=rs("COD_FORMANDO")%>,<%=contProdAgregado%>,<%=rsAgregados("COD_PRODUTO_AGREGADO")%>)" value=" " />
	  </div>
	 
	    <%Else%>
		 <div align="center">
	    <%If Request.QueryString("prod_agre_alterar") <> Empty Then%>
	    <input name="SALVER_PROD" class="btnGravar" type="button" onclick="confirmaAgregados(<%=rs("COD_FORMANDO")%>,<%=contProdAgregado%>,<%=rsAgregados("COD_PRODUTO_AGREGADO")%>)" value=" " />
	        <%Else%>
	          <a href="?cod_pedido=<%=COD_PEDIDO%>&prod_agre_alterar=<%=rsAgregados("COD_PRODUTO_AGREGADO")%>">
	               <img src="../gerencial_foto/images/icons/icon_alterar.png" width="16" height="16" border="0" />	                  </a>
	        <%End If%>
	    
	              <%End If%>	  
	    </div></td>
    </tr>
	<%
	Else
	'MOSTRA QUE FOI SOMENTE PROD. AGREGADO ENTREGUE COMO BRINDE
	%>
    <tr>
      <td><div align="left"><%=rsAgregados("NOME_PRODUTO")%></div></td>
      <td><div align="center"><%=rsAgregados("QUANTIDADE")%></div></td>
      <td><div align="center">N&atilde;o - produto entregue como brinde</div></td>
      <td><div align="center">R$ 0,00</div></td>
      <td bgcolor="#FFFF99"><div align="center">
        <%
	  valor_padrao = 0	  
		  
	  If QTD_PARCELAS > 0 Then
	  	valor_padrao = valorProdutoGeral(rsAgregados("COD_PRODUTO_AGREGADO"),rs("COD_CONTRATO"),QTD_PARCELAS,true,COD_GRUPO_VENDEDORES)
		
		if(valor_padrao) > 0 Then
			Response.Write(formatCurrency(valor_padrao))
		Else
			Response.Write("Produto n&atilde;o possui pre&ccedil;o padr&atilde;o ou boletim informativo")
		End If
		
	  Else
	   Response.Write("N&atilde;o &eacute; poss&iacute;vel verificar, pois o n&uacute;mero de parcelas &eacute; 0.")
	  End If
	  %>
      </div></td>
      <td bgcolor="#FFFF99"><div align="center">
        <%
	  If valor_padrao > 0 and CINT(rsAgregados("VALOR_UN")) > 0 Then
	  		
			diferenca = (rsAgregados("VALOR_UN")-valor_padrao)/1
			diferenca = formatNumber(diferenca)
			
			If diferenca >= 0 Then
				Response.Write("Correto")
			Else
				Response.Write("Valor inferior a tabela, n&atilde;o ser&aacute; poss&iacute;vel finalizar")
			End If
		
		
	  Else
	  	Response.Write("N&atilde;o verificado")
	  End If
	  %>
      </div></td>
      <td><div align="center">-</div></td>
    </tr>
    
	<%
	End If
	
	rsAgregados.MoveNext()
	Wend
	%>
  </table>
  		</form>
  </td>
</tr>	
		<%End If

rsAgregados.Close()
Set rsAgregados = Nothing

%>

<%
'//Fim se n&atilde;o for pedido extra, se for mostra somente os produtos vendidos a parte
End If

Dim rsProdutosAparte
Set rsProdutosAparte = objConexao.Execute(" SELECT TB_PEDIDOS_PRODUTOS.COD_PRODUTO,TB_PEDIDOS_PRODUTOS.COD_PEDIDO_PRODUTO, TB_PRODUTOS.NOME_PRODUTO, TB_PRODUTOS.necessita_fotos, TB_PEDIDOS_PRODUTOS.VALOR_UN,TB_PEDIDOS_PRODUTOS.aliquota, "&_
										  " TB_PEDIDOS_PRODUTOS.QUANTIDADE FROM TB_PRODUTOS INNER JOIN "&_
										  " TB_PEDIDOS_PRODUTOS ON TB_PRODUTOS.COD_PRODUTO = TB_PEDIDOS_PRODUTOS.COD_PRODUTO"&_
										  " WHERE COD_PEDIDO="&COD_PEDIDO&" AND COD_FORMANDO="&rs("COD_FORMANDO")&" ORDER BY NOME_PRODUTO")
If not rsProdutosAparte.Eof Then
%>
<tr bgcolor="whitesmoke">
  <td>
  <div align="center">
  <%If int(PEDIDO_EXTRA) = 2 Then%>
	  <strong>Produtos Vendidos &aacute; parte</strong>
  <%Else%>
  	  <strong>Produtos do Pedido Extra</strong>
  <%End If%>
  </div></td>
</tr>
<tr>
  <td><table width="98%" border="1" align="center" 	cellpadding="2" cellspacing="2" style="border-collapse:collapse">
    <tr>
      <td width="30%"><div align="center"><strong>Produto</strong></div></td>
      <td width="8%"><div align="center"><strong>Qtd.</strong></div></td>
      <td width="8%"><div align="center"><strong>Valor Unit&aacute;rio</strong></div></td>
      <td width="8%"><div align="center"><strong>Valor Total </strong></div></td>
      <td width="8%"><div align="center"><strong>Desc.(%)<img src="../gerencial_foto/images/icons/icon_alterar.png" style="vertical-align:middle; cursor:pointer;" alt="Editar">.</strong></div></td>
      <td width="8%"><div align="center"><strong>Desc.(un.)</strong></div></td>
      <td width="8%"><div align="center"><strong>Desc.(total)</strong></div></td>
      <td width="8%"><div align="center"><strong>Valor Un. Desc.</strong></div></td>
      <td width="8%"><div align="center"><strong>Valor Total Desc.</strong></div></td>
      <td width="8%" bgcolor="#FFFF99"><div align="center"><strong>Valor Padr&atilde;o</strong></div></td>
      <td width="8%" bgcolor="#FFFF99"><div align="center"><strong>Status</strong></div></td>
      <td width="5%"><div align="center"><strong>Excluir</strong></div></td>
    </tr>
    <%While not rsProdutosAparte.Eof%>
	<tr>
      <td><div align="left">
        <%=rsProdutosAparte("NOME_PRODUTO")%>
      	<%
      		if rsProdutosAparte("necessita_fotos") = 1 then

      			set rsFotos = objConexao.Execute("SELECT * FROM TB_PEDIDOS_PRODUTOS_ARQUIVOS WHERE COD_PEDIDO_PRODUTO = "&rsProdutosAparte("COD_PEDIDO_PRODUTO"))

      			if not rsFotos.Eof then
      				Response.Write ("<br /><br />Arquivo(s):")
	      			While not rsFotos.Eof
	      				Response.Write ("<br><a href='lca_busca_detalhamento_pedido_download_arquivo.asp?p0="&rsFotos("id")&"' target='_blank'>-"&rsFotos("descricao")&"</a>")
	      			rsFotos.MoveNext
	      			Wend
	      		end if
      			rsFotos.close
      			set rsFotos = Nothing
      		end if
      	%>
      </div></td>
      <td><div align="center"><%=rsProdutosAparte("QUANTIDADE")%></div></td>
      <td><div align="center"><%=formatCurrency(rsProdutosAparte("VALOR_UN"))%></div></td>
      <td><div align="center">
	  <%
	  Dim tot1, tot2
	 tot1 = formatNumber(0&rsProdutosAparte("QUANTIDADE"),0)
     tot2 = 0&rsProdutosAparte("VALOR_UN")
	  %>
	  <%=formatCurrency(tot1*tot2)%></div></td>
	  <td><div align="center"><%=rsProdutosAparte("aliquota")%>(%)</div></td>
		<%
		xaliquota = formatNumber(0&rsProdutosAparte("aliquota"))
		xdesconto_un = formatCurrency(formatNumber(rsProdutosAparte("VALOR_UN")) - (formatCurrency(rsProdutosAparte("VALOR_UN")) * (1-(xaliquota/100))))
		xdesconto_total = formatCurrency(xdesconto_un * int(rsProdutosAparte("QUANTIDADE")))
		%>

	  <td><div align="center"><%=xdesconto_un%></div></td>
	  <td><div align="center"><%=xdesconto_total%></div></td>
	  <td><div align="center"><%=formatCurrency(formatNumber(rsProdutosAparte("VALOR_UN")) - formatNumber(xdesconto_un))%></div></td>
      <td><div align="center"><strong><%=formatCurrency(formatNumber(rsProdutosAparte("VALOR_UN")) * int(rsProdutosAparte("QUANTIDADE")) - formatCurrency(xdesconto_total)) %></strong></div></td> 
      <td bgcolor="#FFFF99"><div align="center">
        <%
	  valor_padrao = 0	  
		  
	  If QTD_PARCELAS > 0 Then
	  	valor_padrao = valorProdutoGeral(rsProdutosAparte("COD_PRODUTO"),rs("COD_CONTRATO"),QTD_PARCELAS,true,COD_GRUPO_VENDEDORES)
		
		if(valor_padrao) > 0 Then
			Response.Write(formatCurrency(valor_padrao))
		Else
			Response.Write("Produto n&atilde;o possui pre&ccedil;o padr&atilde;o ou boletim informativo")
		End If
		
	  Else
	   Response.Write("N&atilde;o &eacute; poss&iacute;vel verificar, pois o n&uacute;mero de parcelas &eacute; 0.")
	  End If
	  %>
      </div></td>
      <td bgcolor="#FFFF99"><div align="center">
        <%
		valorProduto = formatNumber(0&rsProdutosAparte("VALOR_UN"))
	  If valor_padrao > 0 and valorProduto  > 0 Then
	  		
			diferenca = (valorProduto-valor_padrao)/1
			diferenca = formatNumber(diferenca)
			
			If diferenca >= 0 Then
				Response.Write("Correto")
			Else
				Response.Write("Valor inferior a tabela, n&atilde;o ser&aacute; poss&iacute;vel finalizar")
			End If
		
		
	  Else
	  	Response.Write("N&atilde;o &eacute; poss&iacute;vel verificar")
	  End If
	  %>
      </div></td>
      <td><div align="center">
	  <% If  validaUsuario(2,"",85, true) Then %>
	  
	  <a href="javascript:Void(0)" onclick="confirmaProdAparte(<%=rsProdutosAparte("COD_PEDIDO_PRODUTO")%>)" tabindex="98777"><img src="../gerencial_foto/images/icons/icon_excluir.png" width="16" height="16" border="0" /></a>
	  
	  <% end if%>
	  </div></td>
    </tr>
	<%
	rsProdutosAparte.MoveNext()
	Wend
	%>
  </table></td>
</tr>
<%
End If

rsProdutosAparte.Close()
Set rsProdutosAparte = Nothing
%>
</table>
			<br />

			<%
			rs.MoveNext()
			Wend
				
			rs.Close()
			Set rs = Nothing
		  	FecharConexao()
  			%>
			<%If int(PEDIDO_EXTRA) = 2 Then%>			
            <table width="98%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr>
    <td width="197"><div align="left"><strong class="textoMaior">Total em Produtos:</strong></div></td>
    <td><div align="left"><span class="textoMaior"><strong>
     <%
						Dim rsProdutosTotal, QTD_EM_PRODUTOS, XTOTAL_AGREGADOS, XTOTAL_ALBUNS, XTOTAL_PRODUTOS_A_PARTE, XTOTAL_DESCONTO
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
							
												
							QTD_EM_PRODUTOS = formatCurrency(XTOTAL_AGREGADOS + XTOTAL_ALBUNS + XTOTAL_PRODUTOS_A_PARTE)
						%>
      <%=QTD_EM_PRODUTOS%></strong></span></div></td>
  </tr>

  <tr>
    <td width="197"><div align="left"><strong class="textoMaior">Total em Descontos:</strong></div></td>
    <td><div align="left"><span class="textoMaior"><strong>
     <%
						
						'VW_PEDIDOS_PRODUTOS_TOTAL_AGREGADOS
						set rsProdutosTotal = objConexao.Execute("select SUM(valor_total_com_desconto) as valor_total_com_desconto from(SELECT SUM(faturado * (valor_un-(valor_un * (1-(ifnull(aliquota_desconto_foto,0)/100))))) as valor_total_com_desconto FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO in (SELECT cod_formando FROM TB_PEDIDOS_FORMANDOS WHERE cod_pedido =  "&cod_pedido&") union all SELECT SUM(quantidade * (valor_un-(valor_un * (1-(ifnull(aliquota,0)/100))))) as valor_total_com_desconto  FROM TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO = "&cod_pedido&") as total_desconto")
						if not rsProdutosTotal.Eof then
							XTOTAL_DESCONTO = formatCurrency(rsProdutosTotal(0))
						end if
						rsProdutosTotal.Close
						set rsProdutosTotal = Nothing					
												
							
						%>
     - R$ <%=ARRUMANUMERO(XTOTAL_DESCONTO,"",2) %></strong></span></div></td>
  </tr>
  <tr>
    <td width="197"><div align="left"><strong class="textoMaior">Total em Descontos:</strong></div></td>
    <td><div align="left"><span class="textoMaior"><strong>
      R$ <%=ARRUMANUMERO(QTD_EM_PRODUTOS-XTOTAL_DESCONTO,"",2) %></strong></span></div></td>
  </tr>


</table>
			<%Else%>
			<table width="98%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr>
    <td width="197"><div align="left"><strong class="textoMaior">Total Produtos:</strong></div></td>
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
						
							

					
						%>
      <%=QTD_EM_PRODUTOS%></strong></span></div></td>
  </tr>
  <tr>
    <td width="197"><div align="left"><strong class="textoMaior">Total em Descontos:</strong></div></td>
    <td><div align="left"><span class="textoMaior"><strong>
     <%
						
						'VW_PEDIDOS_PRODUTOS_TOTAL_AGREGADOS
						set rsProdutosTotal = objConexao.Execute("select SUM(valor_total_com_desconto) as valor_total_com_desconto from(SELECT SUM(faturado * (valor_un-(valor_un * (1-(ifnull(aliquota_desconto_foto,0)/100))))) as valor_total_com_desconto FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO in (SELECT cod_formando FROM TB_PEDIDOS_FORMANDOS WHERE cod_pedido =  "&cod_pedido&") union all SELECT SUM(quantidade * (valor_un-(valor_un * (1-(ifnull(aliquota,0)/100))))) as valor_total_com_desconto  FROM TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO = "&cod_pedido&") as total_desconto")
						if not rsProdutosTotal.Eof then
							XTOTAL_DESCONTO = formatCurrency(rsProdutosTotal(0))
						end if
						rsProdutosTotal.Close
						set rsProdutosTotal = Nothing					
					
						FecharConexao()
							
						%>
     - R$ <%=ARRUMANUMERO(XTOTAL_DESCONTO,"",2) %></strong></span></div></td>
  </tr>
   <tr>
    <td width="197"><div align="left"><strong class="textoMaior">Total Final:</strong></div></td>
    <td><div align="left"><span class="textoMaior"><strong>
    	
   R$ <%=ARRUMANUMERO(QTD_EM_PRODUTOS-XTOTAL_DESCONTO,"",2) %>

    </strong></span></div></td>
    </tr>

</table>
			<%End If%>            
			<p align="center" class="textoMaior"> <input type="button" class="btnVoltar" onclick="abre('https://stillos.com.br/gerencial_foto/lca_pedidos_pendentes_alteracoes_part_1.asp?cod_pedido=<%=COD_PEDIDO%>')" name="aaa2" value=" " /> 
			  <input name="bot" class="btnAvancar" type="button" value=" " onclick="abre('lca_pedidos_pendentes_alteracoes_part_3.asp?cod_pedido=<%=COD_PEDIDO%>')" />
			</p>
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
</body>
</html>



