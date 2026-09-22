<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<%
Dim COD_FORMANDO
Dim msg
Dim comando
Dim VALIDA_SUB7, VALIDA_SUB9 
VALIDA_SUB7 = formatNumber(0&inStr(sub_validacoes_paginas,"#7"),0)
VALIDA_SUB9	= formatNumber(0&inStr(sub_validacoes_paginas,"#9"),0)
COD_FORMANDO = Request.QueryString("cod_formando")
msg = Request.QueryString("msg")
comando = Request.QueryString("comando")

AbrirConexao()
		'PARTE1
		Dim COD_CONTRATO, SISTEMA_NOVO, NOME_CURSO, NOME_FORMANDO, ANO_CONCLUSAO, SEMESTRE_CONCLUSAO, ALBUM, INICIAL, BRINDE, COMISSAO, PRODUTO_AGREGADO, VENDIDO, ENCADERNACOES_EXTRAS
		'PARTE2
		Dim ENDERECO, CEP, FONE_RES, FONE_CEL, BAIRRO, CIDADE, ESTADO
				
		sql = "SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO="&COD_FORMANDO
		
		Set rs=objConexao.Execute(sql)
		
		If not rs.Eof Then
				
				ENCADERNACOES_EXTRAS = formatNumber(0&rs("ENCADERNACOES_EXTRAS"),0)
				COD_CONTRATO		= rs("COD_CONTRATO")
				SISTEMA_NOVO		= formatNumber(0&rs("SISTEMA_NOVO"),0)
				NOME_CURSO			= rs("NOME_CURSO")
				NOME_FORMANDO		= rs("NOME_FORMANDO")
				ANO_CONCLUSAO		= rs("ANO_CONCLUSAO")
				SEMESTRE_CONCLUSAO	= rs("SEMESTRE_CONCLUSAO")
				ALBUM				= rs("ALBUM")
				INICIAL				= rs("INICIAL")
				BRINDE				= rs("BRINDE")
				COMISSAO			= rs("COMISSAO")
				PRODUTO_AGREGADO	= rs("PRODUTO_AGREGADO")
				VENDIDO				= rs("VENDIDO")
				
				If SEMESTRE_CONCLUSAO = "J" Then
					SEMESTRE_CONCLUSAO = "JULHO"
				Else
					SEMESTRE_CONCLUSAO = "DEZEMBRO"
				End If
							
		End If
				
		
		rs.Close
		Set rs=Nothing
			
		If Request.Form("BTN_DECLARA_ALBUM") <> Empty Then
			
			Dim ESCOLHA_ALBUM
			
			ESCOLHA_ALBUM = Request.Form("ESCOLHA")
			If ESCOLHA_ALBUM = 1 Then
			'Vai para o vendedor, �ltimo titular da ficha
				sql = 	"INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ( "&_
						"COD_FORMANDO ,"&_
						"TIPO ,"&_
						"COD_VENDEDOR ,"&_
						"DATA_SAIDA ,"&_
						"DATA_RETORNO ,"&_
						"CADASTRADO_POR "&_
						")"&_
						"VALUES ("&_
						COD_FORMANDO&", 0, "&Request.Form("COD_VENDEDOR")&", NOW(), NULL, "&Session("COD_USUARIO")&_
						")"
				
				'Response.Write(sql)
				objConexao.Execute(sql)	
			Else
			'Permanece no estoque
				sql = 	"INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ( "&_
						"COD_FORMANDO ,"&_
						"TIPO ,"&_
						"COD_VENDEDOR ,"&_
						"DATA_SAIDA ,"&_
						"DATA_RETORNO ,"&_
						"CADASTRADO_POR "&_
						")"&_
						"VALUES ("&_
						COD_FORMANDO&", 0, '"&Request.Form("COD_VENDEDOR")&"', NOW(), NOW(), "&Session("COD_USUARIO")&_
						")"
				
				Response.Write(sql)
				objConexao.Execute(sql)
			End If			
			
			Response.Redirect("dip_busca_detalhamento_formando.asp?cod_formando="&COD_FORMANDO&"&msg=álbum verificado com sucesso!")
			
		End If
	
				
							
FecharConexao()%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>Detalhes do Formando</title>
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
messageObj.setShadowOffset(5);	// Large shadowfunction displayMessage(url)
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
}</script>
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
		
		
	var panelWidth = 113;	// Width of help panel	
	var slideSpeed = 15;		// Higher = quicker slide
	var slideTimer = 10;	// Lower = quicker slide
	var slideActive = true;	// Slide active ?
	var initBodyMargin = 0;	// Left or top margin of your <body> tag (left if panel is at the left, top if panel is on the top)
	var pushMainContentOnSlide = false;	// Push your main content to the right when sliding
	var panelPosition = 1; 	// 0 = left , 1 = top
	
	/*	Don't change these values */
	var slideLeftPanelObj=false;
	var slideInProgress = false;	
	var startScrollPos = false;
	var panelVisible = false;
	function initSlideLeftPanel(expandOnly)
	{
		if(slideInProgress)return;
		if(!slideLeftPanelObj){
			if(document.getElementById('dhtmlgoodies_leftPanel')){	// Object exists in HTML code?
				slideLeftPanelObj = document.getElementById('dhtmlgoodies_leftPanel');
				if(panelPosition == 1)slideLeftPanelObj.style.width = '100%';
			}else{	// Object doesn't exist -> Create <div> dynamically
				slideLeftPanelObj = document.createElement('DIV');
				slideLeftPanelObj.id = 'dhtmlgoodies_leftPanel';
				slideLeftPanelObj.style.display='none';
				document.body.appendChild(slideLeftPanelObj);
			}
			
			if(panelPosition == 1){
				slideLeftPanelObj.style.top = "-" + panelWidth + 'px';
				slideLeftPanelObj.style.left = '0px';	
				slideLeftPanelObj.style.height = panelWidth + 'px';			
			}else{
				slideLeftPanelObj.style.left = "-" + panelWidth + 'px';
				slideLeftPanelObj.style.top = '0px';
				slideLeftPanelObj.style.width = panelWidth + 'px';
			}
			

			if(!document.all || navigator.userAgent.indexOf('Opera')>=0)slideLeftPanelObj.style.position = 'fixed';;
		}	
		
		if(panelPosition == 0){
			if(document.documentElement.clientHeight){
				slideLeftPanelObj.style.height = document.documentElement.clientHeight + 'px';
			}else if(document.body.clientHeight){
				slideLeftPanelObj.style.height = document.body.clientHeight + 'px';
			}
			var leftPos = slideLeftPanelObj.style.left.replace(/[^0-9\-]/g,'')/1;
		}else{
			if(document.documentElement.clientWidth){
				slideLeftPanelObj.style.width = document.documentElement.clientWidth + 'px';
			}else if(document.body.clientHeight){
				slideLeftPanelObj.style.width = document.body.clientWidth + 'px';
			}
			var leftPos = slideLeftPanelObj.style.top.replace(/[^0-9\-]/g,'')/1;			
			
			
		}
		slideLeftPanelObj.style.display='block';
		
		if(panelPosition==1)
			startScrollPos = Math.max(document.body.scrollTop,document.documentElement.scrollTop);
		else
			startScrollPos = Math.max(document.body.scrollLeft,document.documentElement.scrollLeft);
		if(leftPos<(0+startScrollPos)){
			if(slideActive){
				slideLeftPanel(slideSpeed);	
			
			}else{
				document.body.style.marginLeft = panelWidth + 'px';
				slideLeftPanelObj.style.left = '0px';
			}
		}else{
			if(expandOnly)return;
			if(slideActive){		
				slideLeftPanel(slideSpeed*-1);
			}else{
				if(panelPosition == 0){
					if(pushMainContentOnSlide)document.body.style.marginLeft =  initBodyMargin + 'px';
					slideLeftPanelObj.style.left = (panelWidth*-1) + 'px';	
				}else{
					if(pushMainContentOnSlide)document.body.style.marginTop =  initBodyMargin + 'px';
					slideLeftPanelObj.style.top = (panelWidth*-1) + 'px';						
				}			
			}
		}	
		
		if(navigator.userAgent.indexOf('MSIE')>=0 && navigator.userAgent.indexOf('Opera')<0){
			window.onscroll = repositionHelpDiv;
		
			repositionHelpDiv();
		}
		window.onresize = resizeLeftPanel;
		
	}
	
	function resizeLeftPanel()
	{
		if(panelPosition == 0){
			if(document.documentElement.clientHeight){
				slideLeftPanelObj.style.height = document.documentElement.clientHeight + 'px';
			}else if(document.body.clientHeight){
				slideLeftPanelObj.style.height = document.body.clientHeight + 'px';
			}		
		}else{
			if(document.documentElement.clientWidth){
				slideLeftPanelObj.style.width = document.documentElement.clientWidth + 'px';
			}else if(document.body.clientWidth){
				slideLeftPanelObj.style.width = document.body.clientWidth + 'px';
			}	
		}
	}
	
	function slideLeftPanel(slideSpeed){
		slideInProgress =true;
		var scrollValue = 0;
		if(panelPosition==1)
			var leftPos = slideLeftPanelObj.style.top.replace(/[^0-9\-]/g,'')/1;
		else
			var leftPos = slideLeftPanelObj.style.left.replace(/[^0-9\-]/g,'')/1;
			
		leftPos+=slideSpeed;
		okToSlide = true;
		if(slideSpeed<0){
			if(leftPos < ((panelWidth*-1) + startScrollPos)){
				leftPos = (panelWidth*-1) + startScrollPos;	
				okToSlide=false;
			}
		}
		if(slideSpeed>0){
			if(leftPos > (0 + startScrollPos)){
				leftPos = 0 + startScrollPos;
				okToSlide = false;
			}			
		}
		
		
		if(panelPosition==0){
			slideLeftPanelObj.style.left = leftPos + startScrollPos + 'px';
			if(pushMainContentOnSlide)document.body.style.marginLeft = leftPos - startScrollPos + panelWidth + 'px';
		}else{
			slideLeftPanelObj.style.top = leftPos + 'px';
			if(pushMainContentOnSlide)document.body.style.marginTop = leftPos - startScrollPos + panelWidth + 'px';			
			
		}
		if(okToSlide)setTimeout('slideLeftPanel(' + slideSpeed + ')',slideTimer); else {
			slideInProgress = false;
			if(slideSpeed>0)panelVisible=true; else panelVisible = false;
		}
		
	}
	
	
	function repositionHelpDiv()
	{
		if(panelPosition==0){
			var maxValue = Math.max(document.body.scrollTop,document.documentElement.scrollTop);
			slideLeftPanelObj.style.top = maxValue;
		}else{
			var maxValue = Math.max(document.body.scrollLeft,document.documentElement.scrollLeft);
			slideLeftPanelObj.style.left = maxValue;	
			var maxTop = Math.max(document.body.scrollTop,document.documentElement.scrollTop);
			if(!slideInProgress)slideLeftPanelObj.style.top = (maxTop - (panelVisible?0:panelWidth)) + 'px'; 		
		}
	}
	
	function cancelEvent()
	{
		return false;
	}
	function keyboardShowLeftPanel()
	{
			initSlideLeftPanel();
			return false;	
	
	}
	
	function leftPanelKeyboardEvent(e)
	{
		if(document.all)return;
		
		if(e.keyCode==112){
			initSlideLeftPanel();
			return false;
		}
	
		
			
	}
	
	function setLeftPanelContent(text)
	{
		document.getElementById('leftPanelContent').innerHTML = text;
		initSlideLeftPanel(true);
		
	}
	if(!document.all)document.documentElement.onkeypress = leftPanelKeyboardEvent;
	document.documentElement.onhelp  = keyboardShowLeftPanel;

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
	<!--AJUDA F1 E CTRL PARA BUSCA R�PIDA-->
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
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  if (selObj.value !== ""){
  	eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  	if (restore) selObj.selectedIndex=0;
  }
}
function popUp(URL){
	win=window.open(URL,"","width=350,height=220,scrollbars=yes")
	var text = "Se a janela n�o abrir desbloquei os Pop-ups!";
	if(win == null) { alert(text); return; }
	win.moveTo(350,150);
}

function validaForm(){
d = document.ALBUM_ENCONTRADO
	if (!d.ESCOLHA[0].checked && !d.ESCOLHA[1].checked){
		alert('Aten��o, escolha uma das op��es antes de continuar!')
		return false;
	}

}

</script>
</head>

<body>
<!-- Code for the left panel -->
<div id="dhtmlgoodies_leftPanel">
	<a class="closeLink" href="#" onclick="initSlideLeftPanel();return false" style="font-size:15px">Fechar</a>
	<div id="leftPanelContent">
	<!-- This is the content -->
	<div>

<table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
  <td width="1%"><img src="images/icons/icon_help.jpg" /></td>
    <td width="95%" align="left"><strong style="font-size:16px">Ajuda do Sistema</strong><br />
<strong>	Esta é a tela inicial do sistema, verifique o número de acessos ao sistema e a data da última utilização.<br />
</strong>
	</td>
  </tr>
</table>

	</div>
	<!-- End content -->
	</div>
</div>
<!-- End code for the left panel -->
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
  BTN_IMPRIMIR_LINK = " href=""javascript:popUp('lca_busca_detalhamento_pedido_IMPRIMIR.asp?cod_pedido="&COD_PEDIDO&""" )"">"
  %>
  <!-- #include file="includes/sub_menu.asp" -->

  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">

	  <h2>DIP - Busca / Detalhe de Formando / Declaração de Álbum </h2>
      <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	  
		  	<h3 align="center">Detalhes do formando n&ordm; <%=digito(COD_FORMANDO)%></h3>
				<table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
  <tr>
            <td width="17%"><div align="left"><strong>Nome Formando:</strong></div></td>
            <td width="83%"><%=NOME_FORMANDO%></td>
            </tr>
			<tr>
			  <td><div align="left"><strong>Instituição de Ensino: </strong></div></td>
			  <td>
			  <%
			  AbrirConexao()
			  
			  Set rsEntidade = objConexao.Execute("SELECT * FROM TB_ENTIDADES WHERE COD_CONTRATO='"&COD_CONTRATO&"'")
			  
			  		If not rsEntidade.Eof Then
						Response.Write(rsEntidade("DESCRICAO"))
					End If
			  
			  rsEntidade.Close
			  Set rsEntidade = Nothing

			  FecharConexao()
			  %>			  </td>
			</tr>
			<tr>
			  <td><div align="left"><strong>Curso: </strong></div></td>
			  <td><%=NOME_CURSO%></td>
			</tr>
			<tr>
			  <td><div align="left"><strong>Semestre / Ano de Conclusão: </strong></div></td>
			  <td><%=SEMESTRE_CONCLUSAO%> / <%=ANO_CONCLUSAO%></td>
			  </tr>
			<tr>
			  <td><div align="left"><strong>Álbum: </strong></div></td>
			  <td><%=ALBUM%></td>
			  </tr>
			<tr>
			  <td><div align="left"><strong>Inicial</strong>:</div></td>
			  <td><%=INICIAL%></td>
			  </tr>
			  
			  <%IF SISTEMA_NOVO = 1 Then%>
			  <%End If%>
			<%If VALIDA_SUB9 <> 0 Then%>  
					<%
					AbrirConexao()
					sql = "SELECT "&_
								"TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.TIPO, "&_
								"TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR, "&_
								"TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_SAIDA, "&_
								"TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_RETORNO, "&_
								"TB_VENDEDORES.NOME_VENDEDOR  "&_
								"FROM TB_VENDEDORES INNER JOIN TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ON "&_
								"TB_VENDEDORES.COD_VENDEDOR = TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR "&_
								"WHERE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_FORMANDO="&COD_FORMANDO&" AND TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.TIPO = 0"
					
					Set rsMovimentacoes2 = objConexao.Execute(sql)
					
					If rsMovimentacoes2.Eof Then 

					Else
					
						Response.Redirect("dip_busca_detalhamento_formando.asp?cod_formando="&COD_FORMANDO&"&msg=Formando j� possui álbum declarado!")
					
					End If
								
					rsMovimentacoes2.Close
					Set rsMovimentacoes2 = Nothing
					
					FecharConexao()
					%>
			<%End If%>
			</table>
			
			    <h3 align="center">Escolha o destino do Álbum encontrado </h3>
				<form onsubmit="return validaForm()" method="post" name="ALBUM_ENCONTRADO" action="dip_busca_detalhamento_formando_cadastro_album_finalizacao_cadastro_ficha.asp?cod_formando=<%=COD_FORMANDO%>" style="margin:0px">
			    <table width="98%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
                 <tr bgcolor="#E2E2E2">
                    <td colspan="2"><div align="center"><strong>Responsável a partir de agora pelo Álbum </strong></div></td>
                    <td width="71%">
               
                      <div align="center"><strong>Descrição</strong></div></td>
                  </tr>
				  <%
						AbrirConexao()
							
							sql = "SELECT TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.TIPO, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_SAIDA, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_RETORNO FROM TB_VENDEDORES INNER JOIN TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ON TB_VENDEDORES.COD_VENDEDOR = TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR WHERE COD_FORMANDO = "&COD_FORMANDO&" AND TIPO = 1 ORDER BY DATA_SAIDA DESC LIMIT 1"
							'Response.Write(sql)
							Set rs70 = objConexao.Execute(sql)
							
								If not rs70.Eof Then
									nome_vendedor = rs70("NOME_VENDEDOR")
									Dim COD_VENDEDOR
									COD_VENDEDOR = rs70("COD_VENDEDOR")
								End If
							
							rs70.Close
							Set rs70 = Nothing
						
						FecharConexao()
						%>
			<%If COD_VENDEDOR <> 27 and COD_VENDEDOR > 0 Then%>
				  <tr bgcolor="#FAFEC2">
                    <td width="3%" class="textoMedio"><div align="center">
                      <input class="inputRadio" name="ESCOLHA" type="radio" value="1" />
                    </div></td>
                    <td width="26%" bgcolor="#FAFEC2" class="textoMedio">
					<strong><%=nome_vendedor%></strong>
					</td>
                    <td width="71%" bgcolor="#FAFEC2" class="textoMedio">O Álbum não está no Estoque, está com o vendedor<br />
                             -Depois você poderá dar <strong>entrada</strong> no estoque desta forma </td>
                  </tr>				<%End If%>
                  <tr bgcolor="#F8E28F">
                    <td><div align="center"><span class="textoMedio">
                      <input class="inputRadio" checked="checked" name="ESCOLHA" type="radio" value="2" />
                    </span></div></td>					<%
					'06/12/2012
					'Efetua uma corre��o para quando o vendedor for igual a 0.
					If COD_VENDEDOR = 0 Then
						
						
						
						
					
						COD_VENDEDOR = 27
					End If
					%>
                    <td bgcolor="#F8E28F"><strong class="textoMedio">ESTOQUE<input type="hidden" name="COD_VENDEDOR" value="<%=COD_VENDEDOR%>" /></strong></td>
                    <td bgcolor="#F8E28F" class="textoMedio">O álbum foi encontrado no estoque, por tanto não ficará pendente para o vendedor <br />
                             -Depois você poderá dar <strong>saída</strong> do estoque desta forma</td>
                  </tr>
                  <tr>
                    <td colspan="3"><div align="center">
                      <input type="submit" name="BTN_DECLARA_ALBUM" class="btnGravar" value=" " />
                    </div></td>
                    </tr>
                  <%IF SISTEMA_NOVO = 1 Then%>
                  <%End If%>
                  <%If VALIDA_SUB9 <> 0 Then%>
                  <%
					AbrirConexao()
					sql = "SELECT "&_
								"TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.TIPO, "&_
								"TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR, "&_
								"TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_SAIDA, "&_
								"TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_RETORNO, "&_
								"TB_VENDEDORES.NOME_VENDEDOR  "&_
								"FROM TB_VENDEDORES INNER JOIN TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ON "&_
								"TB_VENDEDORES.COD_VENDEDOR = TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR "&_
								"WHERE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_FORMANDO="&COD_FORMANDO&" AND TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.TIPO = 0"
					
					Set rsMovimentacoes2 = objConexao.Execute(sql)
					
					If rsMovimentacoes2.Eof Then 
					%>

                  <%
					End If
								
					rsMovimentacoes2.Close
					Set rsMovimentacoes2 = Nothing
					
					FecharConexao()
					%>
                  <%End If%>
                </table>			    
				</form>
	
			    <p align="center"> </p></td>
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
