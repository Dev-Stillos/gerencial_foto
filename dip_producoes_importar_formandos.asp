<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<%
Dim COD_PRODUCAO
COD_PRODUCAO = Request.QueryString("cod_prod")
Dim DESCRICAO
Dim COD_CONTRATO1, COD_CONTRATO2, COD_CONTRATO3, COD_CONTRATO4, COD_CONTRATO5, COD_CONTRATO6

COD_CONTRATO1 = Request.QueryString("COD_CONTRATO1")
COD_CONTRATO2 = Request.QueryString("COD_CONTRATO2")
COD_CONTRATO3 = Request.QueryString("COD_CONTRATO3")
COD_CONTRATO4 = Request.QueryString("COD_CONTRATO4")
COD_CONTRATO5 = Request.QueryString("COD_CONTRATO5")
COD_CONTRATO6 = Request.QueryString("COD_CONTRATO6")

AbrirConexao()
	Dim rsDes
	Set rsDes = objConexao.Execute("SELECT DESCRICAO FROM TB_PRODUCAO WHERE COD_PRODUCAO="&COD_PRODUCAO&" LIMIT 1")
	
		DESCRICAO=rsDes("DESCRICAO")
		
	rsDes.Close()
	Set rsDes = Nothing
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
	
	<script language="javascript">
	

	function ValidaForm(){
	d = document.valida
		if (d.COD_CONTRATO1.value == ""){
			d.COD_CONTRATO1.focus();
			alert('Preencha pelo menos o contrato 1');
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
<strong>Esta é a tela inicial do sistema, verifique o nímero de acessos ao sistema e a data da última utilização.<br />
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
		 	<form name="valida" method="post" action="dip_producoes_importar_formandos_part_2.asp?cod_prod=<%=COD_PRODUCAO%>"  onsubmit="return ValidaForm()">
			<table width="90%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse; margin-top:6px">
           <tr>
             <td width="20%"><div align="center"><strong>*Contrato 1: </strong></div></td>
             <td width="80%"><input name="COD_CONTRATO1" type="text" size="8" maxlength="5" /></td>
             </tr>
           <tr>
             <td><div align="center">Contrato 2: </div></td>
             <td><input type="text" size="8" maxlength="5" name="COD_CONTRATO2" /></td>
             </tr>
			 <tr>
             <td><div align="center">Contrato 3: </div></td>
             <td><input type="text" size="8" maxlength="5" name="COD_CONTRATO3" /></td>
             </tr>
			 <tr>
             <td><div align="center">Contrato 4: </div></td>
             <td><input type="text" size="8" maxlength="5" name="COD_CONTRATO4" /></td>
             </tr>
			 <tr>
             <td><div align="center">Contrato 5: </div></td>
             <td><input type="text" size="8" maxlength="5" name="COD_CONTRATO5" /></td>
             </tr>
			 <tr>
             <td><div align="center">Contrato 6: </div></td>
             <td><input type="text" size="8" maxlength="5" name="COD_CONTRATO6" /></td>
             </tr>
         
                    <tr>
             <td colspan="2"><div align="center">
               <p align="left">* Informe pelo menos o primeiro contrato</p>
               <p>
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