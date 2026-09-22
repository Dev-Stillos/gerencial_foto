<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim cod_formando, redirecionar_volta, imprimir_agora
Dim NOME_FORMANDO, NOME_CURSO, ENCADERNACOES_EXTRAS, SEMESTRE_CONCLUSAO, ANO_CONCLUSAO, ALBUM, COD_CONTRATO
cod_formando = Request.QueryString("cod_formando")
redirecionar_volta = Request.QueryString("redit")


Dim rs, rs2
AbrirConexao()
Set rs = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO="&cod_formando)

		If not rs.Eof Then
			NOME_FORMANDO			= rs("NOME_FORMANDO")
			NOME_CURSO				= rs("NOME_CURSO") 
			SEMESTRE_CONCLUSAO		= rs("SEMESTRE_CONCLUSAO") 
			ANO_CONCLUSAO			= rs("ANO_CONCLUSAO") 
			ALBUM					= rs("ALBUM") 
			COD_CONTRATO			= rs("COD_CONTRATO")
			ENCADERNACOES_EXTRAS 	= formatNumber(0&rs("ENCADERNACOES_EXTRAS"),0)
		End If

rs.Close
Set rs = Nothing

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
	<script language="javascript">
	function validaForm(){
	d = document.CAD_FORMANDOS_2
		if(d.ALBUM.value == ""){
		d.ALBUM.focus();
		alert('Campo álbum não digitado!');
		return false;
		}
	
		if(d.INICIAL.value == ""){
		d.INICIAL.focus();
		alert('Campo inicial não digitado!');
		return false;
		}
	
		if(d.ALBUM_EM_POSSE.value == 3 && d.FICHAS_EM_POSSE.value == 27){
		d.ALBUM_EM_POSSE.focus();
		alert('Você informou que o álbum está com o vendedor, mas não informou qual vendedor é!');
		return false;
		}
	
	}
	</script>
	
	<script language="javascript">
		function popUp(campo,URL,campoSelect,msgErro){
			
			if(document.CAD_FORMANDOS_2[campo].value==1){
				
				if (campoSelect !== 'no'){
					//Brinde = "Sim"
					document.CAD_FORMANDOS_2[campoSelect].options[1].selected = "true";
				}
			
				win=window.open(URL,"","width=350,height=220,scrollbars=yes")
				var text = "Se a janela não abrir desbloquei os Pop-ups!";
				if(win == null) { alert(text); return; }
				win.moveTo(350,150);
					
			}
			else {
				if (msgErro==1){alert('Para visualizar os brindes é necessário alterar a opção Brinde para Sim.')}
			}
		
		}
		
		
		function fecharPopUp(){
			if(win){
				win.close();
			}
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
			
		function focuInic(){
					abrir('dip_formandos_etiquetas_individual.asp?cod_formando=<%=COD_FORMANDO%>','teste','width=200,height=200,scrollbars=no')
		}	
	</script>
</head>

<body onload="focuInic()">

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

	  <h2>DIP - Impress&atilde;o de Etiqueta de Identifica&ccedil;&atilde;o </h2>
	   <div class="msg" align="center"><%=Request.QueryString("msg")%></div>

		  <div id="comboTexto">
		    <table width="95%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
		        	<tr>
            <td width="24%" bgcolor="#CCCCCC"><strong>Nome Formando: </strong></td>
            <td width="76%" bgcolor="#CCCCCC"><strong><%=NOME_FORMANDO%></strong></td>
            </tr>
					<tr>
				<td width="24%" bgcolor="#CCCCCC"><strong>Curso:</strong></td>
            <td width="76%" bgcolor="#CCCCCC"><strong><%=NOME_CURSO%></strong></td>
            </tr>
			
			<tr>
            <td colspan="2"><div align="center">
              <p><strong>Impress&atilde;o de Etiqueta </strong></p>
			  
              <p><a href="javascript:abrir('dip_formandos_etiquetas_individual.asp?cod_formando=<%=COD_FORMANDO%>','teste','width=200,height=200,scrollbars=no')" class="linkMaster">CASO A JANELA N&Atilde;O ABRIR CLIQUE AQUI PARA IMPRIMIR</a></p>
              <p>&nbsp;</p>
              <p><a href="<%=redirecionar_volta%>"><strong>VOLTAR</strong></a></p>
              <p>&nbsp;</p>
            </div></td>
            </tr>
			
		         </table>
	
		  </div>
		
		  <p>&nbsp;</p>



				
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