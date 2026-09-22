<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<%
Dim COD_PRODUCAO
COD_PRODUCAO = int(Request.QueryString("cod_prod"))
Dim DESCRICAO
Dim COD_CONTRATO1, COD_CONTRATO2, COD_CONTRATO3, COD_CONTRATO4, COD_CONTRATO5, COD_CONTRATO6
Dim rsCOD_CONTRATO
Dim TURMA1, TURMA2, TURMA3, TURMA4, TURMA5, TURMA6
Dim VALIDA_1, VALIDA_2, VALIDA_3, VALIDA_4, VALIDA_5, VALIDA_6 


'Valida = 3, significa que foi preenchido o contrato informado.
'Valida = 2, significa que n�o foi preenchido o contrato
'Valida = 1, significa que est� ok e foi encontrato
'Valida = 4, sgignifica que foi encontrado, entretando j� h� outra produ��o que utiliza este contrato
VALIDA_1 = 3
VALIDA_2 = 2
VALIDA_3 = 2
VALIDA_4 = 2
VALIDA_5 = 2
VALIDA_6 = 2

COD_CONTRATO1 = ""&Request.Form("COD_CONTRATO1")
COD_CONTRATO2 = ""&Request.Form("COD_CONTRATO2")
COD_CONTRATO3 = ""&Request.Form("COD_CONTRATO3")
COD_CONTRATO4 = ""&Request.Form("COD_CONTRATO4")
COD_CONTRATO5 = ""&Request.Form("COD_CONTRATO5")
COD_CONTRATO6 = ""&Request.Form("COD_CONTRATO6")

AbrirConexao()
	Dim rsDes
  response.write("SELECT DESCRICAO FROM TB_PRODUCAO WHERE COD_PRODUCAO="&COD_PRODUCAO&" LIMIT 1")
	Set rsDes = objConexao.Execute("SELECT DESCRICAO FROM TB_PRODUCAO WHERE COD_PRODUCAO="&COD_PRODUCAO&" LIMIT 1")
	
		DESCRICAO=rsDes("DESCRICAO")
		
	rsDes.Close()
	Set rsDes = Nothing



Dim rsValida1, sql, qtdContratos
qtdContratos = 1
	sql = 	"SELECT DISTINCT CONTRATO, TURMA FROM TB_FORMANDOS_IMP WHERE CONTRATO = "&sqlTexto(COD_CONTRATO1)
			
			'ACCESS FDP!!!
			If COD_CONTRATO2 <> Empty Then
				sql = sql&" OR CONTRATO = "&sqlTexto(COD_CONTRATO2)
				qtdContratos = qtdContratos + 1
				VALIDA_2 = 3
			End If
			
			If COD_CONTRATO3 <> Empty Then
				sql = sql&" OR CONTRATO = "&sqlTexto(COD_CONTRATO3)
				qtdContratos = qtdContratos + 1
				VALIDA_3 = 3
			End If
			
			If COD_CONTRATO4 <> Empty Then
				sql = sql&" OR CONTRATO = "&sqlTexto(COD_CONTRATO4)
				qtdContratos = qtdContratos + 1
				VALIDA_4 = 3
			End If
			
			If COD_CONTRATO5 <> Empty Then
				sql = sql&" OR CONTRATO = "&sqlTexto(COD_CONTRATO5)
				qtdContratos = qtdContratos + 1
				VALIDA_5 = 3
			End If
			
			If COD_CONTRATO6 <> Empty Then
				sql = sql&" OR CONTRATO = "&sqlTexto(COD_CONTRATO6)
				qtdContratos = qtdContratos + 1
				VALIDA_6 = 3
			End If
		
	
	
	Set rsValida1 = objConexao.Execute(sql)
			
			If not rsValida1.Eof Then		
				While not rsValida1.Eof
					'AGORA BASTA VERIFICAR SE H� REGISTRO DO COD. FORMANDO ANTES DE DAR O INSERT!!!!!!
					
					rsCOD_CONTRATO = ""&rsValida1("CONTRATO")
					If rsCOD_CONTRATO = COD_CONTRATO1 Then
						VALIDA_1 	= 1
						TURMA1		= rsValida1("TURMA")
						
						'Verifica se h� registro com este contrato j� em algum formando
						Dim rsVerificaImporta
						Set rsVerificaImporta = objConexao.Execute("SELECT COD_CONTRATO, COD_PRODUCAO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_CONTRATO ='"&rsCOD_CONTRATO&"' LIMIT 1")
						
							If not rsVerificaImporta.Eof Then
							Dim arpa
							arpa = formatNumber(0&rsVerificaImporta("COD_PRODUCAO"),0)
								If arpa = COD_PRODUCAO Then
											VALIDA_1	= 44
								Else
											VALIDA_1	= 4
											TURMA1		= TURMA1&" - J� utilizado na produ��o c�d.:"&rsVerificaImporta("COD_PRODUCAO")
											
								
								End If
											
											
							End If
						
						rsVerificaImporta.Close()
						Set rsVerificaImporta = Nothing
						
					End If
					
					If rsCOD_CONTRATO = COD_CONTRATO2 Then
						VALIDA_2 	= 1
						TURMA2		= rsValida1("TURMA")
						
						'Verifica se h� registro com este contrato j� em algum formando
		
						Set rsVerificaImporta = objConexao.Execute("SELECT COD_CONTRATO, COD_PRODUCAO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_CONTRATO ='"&rsCOD_CONTRATO&"' LIMIT 1")
						
							If not rsVerificaImporta.Eof Then
								VALIDA_2	= 4
								TURMA2		= TURMA1&" - J� utilizado na produ��o c�d.:"&rsVerificaImporta("COD_PRODUCAO")
							End If
						
						rsVerificaImporta.Close()
						Set rsVerificaImporta = Nothing
					End If
					
					If rsCOD_CONTRATO = COD_CONTRATO3 Then
						VALIDA_3 	= 1
						TURMA3		= rsValida1("TURMA")
						
						'Verifica se h� registro com este contrato j� em algum formando
						Set rsVerificaImporta = objConexao.Execute("SELECT COD_CONTRATO, COD_PRODUCAO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_CONTRATO ="&rsCOD_CONTRATO&" LIMIT 1")
						
							If not rsVerificaImporta.Eof Then
								VALIDA_3	= 4
								TURMA3		= TURMA1&" - J� utilizado na produ��o c�d.:"&rsVerificaImporta("COD_PRODUCAO")
							End If
						
						rsVerificaImporta.Close()
						Set rsVerificaImporta = Nothing
					End If
	
					If rsCOD_CONTRATO = COD_CONTRATO4 Then
						VALIDA_4 	= 1
						TURMA4		= rsValida1("TURMA")
						
						'Verifica se h� registro com este contrato j� em algum formando
						Set rsVerificaImporta = objConexao.Execute("SELECT COD_CONTRATO, COD_PRODUCAO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_CONTRATO ="&rsCOD_CONTRATO&" LIMIT 1")
						
							If not rsVerificaImporta.Eof Then
								VALIDA_4	= 4
								TURMA4		= TURMA1&" - J� utilizado na produ��o c�d.:"&rsVerificaImporta("COD_PRODUCAO")
							End If
						
						rsVerificaImporta.Close()
						Set rsVerificaImporta = Nothing
					End If
					
					If rsCOD_CONTRATO = COD_CONTRATO5 Then
						VALIDA_5 	= 1
						TURMA5		= rsValida1("TURMA")
						
						'Verifica se h� registro com este contrato j� em algum formando
						Set rsVerificaImporta = objConexao.Execute("SELECT COD_CONTRATO, COD_PRODUCAO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_CONTRATO ='"&rsCOD_CONTRATO&"' LIMIT 1")
						
							If not rsVerificaImporta.Eof Then
								VALIDA_5	= 4
								TURMA5		= TURMA1&" - J� utilizado na produ��o c�d.:"&rsVerificaImporta("COD_PRODUCAO")
							End If
						
						rsVerificaImporta.Close()
						Set rsVerificaImporta = Nothing
					End If
					
					If rsCOD_CONTRATO = COD_CONTRATO6 Then
						VALIDA_6 	= 1
						TURMA6		= rsValida1("TURMA")
						
						'Verifica se h� registro com este contrato j� em algum formando
						Set rsVerificaImporta = objConexao.Execute("SELECT COD_CONTRATO, COD_PRODUCAO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_CONTRATO ="&rsCOD_CONTRATO&" LIMIT 1")
						
							If not rsVerificaImporta.Eof Then
								VALIDA_6	= 4
								TURMA6		= TURMA1&" - J� utilizado na produ��o c�d.:"&rsVerificaImporta("COD_PRODUCAO")
							End If
						
						rsVerificaImporta.Close()
						Set rsVerificaImporta = Nothing
					End If				
					
				
				rsValida1.MoveNext()
				Wend			
			Else
			
				Response.Redirect("dip_producoes.asp?msg=N�o foi poss�vel importar formandos, pois n�o foi encontrado formandos relacionados aos contratos informados para produ��o "&COD_PRODUCAO&"-"&DESCRICAO&"!")
			
			
			End If
			
	rsValida1.Close()
	Set rsValida1 = Nothing



FecharConexao()

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
	<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>
	
	<script language="javascript">
		function popUp(pagina){
		
		
		}
		
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
				//		src.children.tags('A')0.click(); 
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
			var leftPos = slideLeftPanelObj.style.left.replace(/^0-9\-/g,'')/1;
		}else{
			if(document.documentElement.clientWidth){
				slideLeftPanelObj.style.width = document.documentElement.clientWidth + 'px';
			}else if(document.body.clientHeight){
				slideLeftPanelObj.style.width = document.body.clientWidth + 'px';
			}
			var leftPos = slideLeftPanelObj.style.top.replace(/^0-9\-/g,'')/1;			
			
			
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
			var leftPos = slideLeftPanelObj.style.top.replace(/^0-9\-/g,'')/1;
		else
			var leftPos = slideLeftPanelObj.style.left.replace(/^0-9\-/g,'')/1;
			
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
	
	if(YNconfirm("Tem certeza que deseja continuar?\n*Os contratos com pend�ncias n�o ser�o importados")){
			document.valida.submit();
			return true;
	}else{
		return false;
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
<strong>Esta � a tela inicial do sistema, verifique o n�mero de acessos ao sistema e a data da �ltima utiliza��o.<br />
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
  BTN_FONTE = 1
  BTN_IMPRIMIR_LINK = "#"
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">

	  <h2>DIP - Produ&ccedil;&otilde;es / Importa&ccedil;&atilde;o por Contrato </h2>
	  
	  <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
		  <h3 align="center">Informe os contratos para importa&ccedil;&atilde;o da produ&ccedil;&atilde;o:<br />
		    
		    <%=COD_PRODUCAO%> / <%=DESCRICAO%>		  </h3>
		 	<p align="left">
		 	<form name="valida" method="post" action="dip_producoes_importar_formandos_processar.asp?cod_prod=<%=COD_PRODUCAO%>&comando=incluir"  onsubmit="return confirma()">
			<table width="90%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse; margin-top:6px">
    	  	<%If VALIDA_1 <> 2 Then%>
			   	<tr>
             <td width="11%"><div align="center"><strong>*Contrato 1: </strong></div></td>
             <td width="35%">
			   <div align="left">
			     <%If VALIDA_1 = 1 or VALIDA_1 = 44 Then%>
			      <%=TURMA1%>
			     <input value="<%=COD_CONTRATO1%>" name="COD_CONTRATO1" type="hidden"  size="8" maxlength="4" />
			       <%Else%>
				   <%=COD_CONTRATO1%>
				   <%End IF%>
			       </div></td>
             <td width="54%">
               <table width="220" border="0" cellspacing="0" cellpadding="0">
                 <%If VALIDA_1 = 1 or VALIDA_1 = 44 Then%>
				 <tr>
                   <td width="11%"><img src="images/icons/icon_green.gif" width="16" height="15" /></td>
                   <td width="89%"><strong>Encontrado!</strong></td>
                 </tr>
				 <%Elseif VALIDA_1 = 3 Then%>
              	   <tr>
                   <td><img src="images/icons/icon_red.gif" width="17" height="15" /></td>
                   <td><span class="msg">N&atilde;o encontrado!</span></td>
                 </tr>
				 <%Elseif VALIDA_1 = 4 Then%>
              	   <tr>
              	     <td><img src="images/icons/icon_red.gif" width="17" height="15" /></td>
              	     <td class="msg">Consta em outra produ&ccedil;&atilde;o! </td>
            	     </tr>
				 <%End If%>
               </table></td>
			   	</tr>
	        <%End If%>
		  	<%If VALIDA_2 <> 2 Then%>  
		   		<tr>
             <td><div align="center">Contrato 2: </div></td>
             <td>
			   <div align="left">
			     <%If VALIDA_2 = 1 Then%>
			      <%=TURMA2%>
			     <input value="<%=COD_CONTRATO2%>" name="COD_CONTRATO2" type="hidden"  size="8" maxlength="4" />
			       	<%Else%>
				   <%=COD_CONTRATO2%>
				   <%End IF%>
			       </div></td>
             <td><table width="220" border="0" cellspacing="0" cellpadding="0">
                 <%If VALIDA_2 = 1 Then%>
				 <tr>
                   <td width="11%"><img src="images/icons/icon_green.gif" width="16" height="15" /></td>
                   <td width="89%"><strong>Encontrado!</strong></td>
                 </tr>
				 <%Elseif VALIDA_2 = 3 Then%>
              	   <tr>
                   <td><img src="images/icons/icon_red.gif" width="17" height="15" /></td>
                   <td><span class="msg">N&atilde;o encontrado!</span></td>
                 </tr>
				 <%Elseif VALIDA_2 = 4 Then%>
              	   <tr>
              	     <td><img src="images/icons/icon_red.gif" width="17" height="15" /></td>
              	     <td class="msg">Consta em outra produ&ccedil;&atilde;o! </td>
            	     </tr>
				 <%End If%>
               </table></td>
		   		</tr>
			<%End If%>
			<%If VALIDA_3 <> 2 Then%>
				<tr>
             <td><div align="center">Contrato 3: </div></td>
             <td>
               <div align="left">
                 <%If VALIDA_3 = 1 Then%>
                  <%=TURMA3%>
                 <input value="<%=COD_CONTRATO3%>" name="COD_CONTRATO3" type="hidden"  size="8" maxlength="4" />
                   			       <%Else%>
				   <%=COD_CONTRATO3%>
				   <%End IF%>
                 </div></td>
             <td><table width="220" border="0" cellspacing="0" cellpadding="0">
                 <%If VALIDA_3 = 1 Then%>
				 <tr>
                   <td width="11%"><img src="images/icons/icon_green.gif" width="16" height="15" /></td>
                   <td width="89%"><strong>Encontrado!</strong></td>
                 </tr>
				 <%Elseif VALIDA_3 = 3 Then%>
              	   <tr>
                   <td><img src="images/icons/icon_red.gif" width="17" height="15" /></td>
                   <td><span class="msg">N&atilde;o encontrado!</span></td>
                 </tr>
				 <%Elseif VALIDA_3 = 4 Then%>
              	   <tr>
              	     <td><img src="images/icons/icon_red.gif" width="17" height="15" /></td>
              	     <td class="msg">Consta em outra produ&ccedil;&atilde;o! </td>
            	     </tr>
				 <%End If%>
               </table></td>
				</tr>
 			<%End If%>
			<%If VALIDA_4 <> 2 Then%>	
				 <tr>
             <td><div align="center">Contrato 4: </div></td>
             <td>
               <div align="left">
                 <%If VALIDA_4 = 1 Then%>
                  <%=TURMA4%>
                 <input value="<%=COD_CONTRATO4%>" name="COD_CONTRATO4" type="hidden"  size="8" maxlength="4" />
                   			       <%Else%>
				   <%=COD_CONTRATO4%>
				   <%End IF%>
                 </div></td>
             <td><table width="220" border="0" cellspacing="0" cellpadding="0">
                 <%If VALIDA_4 = 1 Then%>
				 <tr>
                   <td width="11%"><img src="images/icons/icon_green.gif" width="16" height="15" /></td>
                   <td width="89%"><strong>Encontrado!</strong></td>
                 </tr>
				 <%Elseif VALIDA_4 = 3 Then%>
              	   <tr>
                   <td><img src="images/icons/icon_red.gif" width="17" height="15" /></td>
                   <td><span class="msg">N&atilde;o encontrado!</span></td>
                 </tr>
				 <%Elseif VALIDA_4 = 4 Then%>
              	   <tr>
              	     <td><img src="images/icons/icon_red.gif" width="17" height="15" /></td>
              	     <td class="msg">Consta em outra produ&ccedil;&atilde;o! </td>
            	     </tr>
				 <%End If%>
               </table></td>
				 </tr>
			 <%End If%>
			<%If VALIDA_5 <> 2 Then%>	
				 <tr>
             <td><div align="center">Contrato 5: </div></td>
             <td>
               <div align="left">
                 <%If VALIDA_5 = 1 Then%>
                  <%=TURMA5%>
                 <input value="<%=COD_CONTRATO5%>" name="COD_CONTRATO5" type="hidden"  size="8" maxlength="4" />
                 			       <%Else%>
				   <%=COD_CONTRATO5%>
				   <%End IF%>
                 </div></td>
             <td><table width="220" border="0" cellspacing="0" cellpadding="0">
                 <%If VALIDA_5 = 1 Then%>
				 <tr>
                   <td width="11%"><img src="images/icons/icon_green.gif" width="16" height="15" /></td>
                   <td width="89%"><strong>Encontrado!</strong></td>
                 </tr>
				 <%Elseif VALIDA_5 = 3 Then%>
              	   <tr>
                   <td><img src="images/icons/icon_red.gif" width="17" height="15" /></td>
                   <td><span class="msg">N&atilde;o encontrado!</span></td>
                 </tr>
				 <%Elseif VALIDA_5 = 4 Then%>
              	   <tr>
              	     <td><img src="images/icons/icon_red.gif" width="17" height="15" /></td>
              	     <td class="msg">Consta em outra produ&ccedil;&atilde;o! </td>
            	     </tr>
				 <%End If%>
               </table></td>
				 </tr>
 			<%End If%>
			<%If VALIDA_6 <> 2 Then%>	
				<tr>
				 <td><div align="center">Contrato 6: </div></td>
				 <td>
				   <div align="left">
				     <%If VALIDA_6 = 1 Then%>
				          <%=TURMA6%>
				     <input value="<%=COD_CONTRATO6%>" name="COD_CONTRATO6" type="hidden"  size="8" maxlength="4" />
				        			       <%Else%>
				   <%=COD_CONTRATO6%>
				   <%End IF%>
				     </div></td>
				 <td><table width="220" border="0" cellspacing="0" cellpadding="0">
                 <%If VALIDA_6 = 1 Then%>
				 <tr>
                   <td width="11%"><img src="images/icons/icon_green.gif" width="16" height="15" /></td>
                   <td width="89%"><strong>Encontrado!</strong></td>
                 </tr>
				 <%Elseif VALIDA_6 = 3 Then%>
              	   <tr>
                   <td><img src="images/icons/icon_red.gif" width="17" height="15" /></td>
                   <td><span class="msg">N&atilde;o encontrado!</span></td>
                 </tr>
				 <%Elseif VALIDA_6 = 4 Then%>
              	   <tr>
              	     <td><img src="images/icons/icon_red.gif" width="17" height="15" /></td>
              	     <td class="msg">Consta em outra produ&ccedil;&atilde;o! </td>
            	     </tr>
				 <%End If%>
               </table></td>
				</tr>
          	<%End If%>
                    <tr>
             <td colspan="3"><div align="center">
               <p align="left">* Informe pelo menos o primeiro contrato</p>
               <p>
                 <input class="btnVoltar" onclick="history.go(-1)"  type="button" name="GRAVAR22" value=" " />
                 &nbsp; 
                 <input class="btnAvancar"  type="submit" name="GRAVAR2" value=" " />
                 </p>
             </div></td>
             </tr>
		 </table>
		 	</form>	
		 </td>
      </tr>
    </table>
  </p></td>
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