<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
If Request.QueryString("limpar_agendadas") = "sim" Then

AbrirConexao()
 	objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_ETIQUETAS_TEMPORARIA WHERE COD_USUARIO="&Session("COD_USUARIO")&" AND AGENDADA=1")
FecharConexao()
	Response.Redirect("adm_etiquetas_agendadas.asp?msg=Etiquetas Agendadas deletadas com sucesso!")
End If


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
	<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>
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
	
<!--calendário-->
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
		
		calendarObjForForm.setCalendarPositionByHTMLElement(inputObject,-60,inputObject.offsetHeight-175);	// Position the calendar right below the form input
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
<!--calendário-->

<!--valida form-->
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
	function confirmarExclusaoAgendamento(){
		if(YNconfirm('Tem certeza que deseja limpar as etiquetas agendadas pra impressão\nCertifique-se que foram impressão com sucesso antes de continuar.')){
			abre('?limpar_agendadas=sim')
		}
	}
</script>

<!--valida form-->
</head>

<body onKeyDown="javascript:Verificar()">
<!-- Code for the left panel -->
<div id="dhtmlgoodies_leftPanel">
	<a class="closeLink" href="#" onclick="initSlideLeftPanel();return false" style="font-size:15px">Fechar</a>
	<div id="leftPanelContent">
	<!-- This is the content -->
	<div>

<table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
  <td width="1%" valign="top"><img src="images/icons/icon_help.jpg" /></td>
    <td width="95%" align="left"><strong style="font-size:16px">Ajuda do Sistema</strong><br />
<strong>Esta é a tela Log do usuário <%=Session("NOME")%>, é possível verificar as datas e horários de acesso ao sistema e a quantidade de acessos também.<br />
<%If Session("COD_USUARIO") =< 2 then %>Para verificar as fichas baixadas utilize os filtros: "últimos registros" ou "por data."<br /><br />
Lembrete: A data inicial pré-selecionada é correspondente ao mês e ano atual.<%End If%>

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
	
	  <h2>Agendamento de Etiquetas</h2>
	   <div class="subMenu">Sub-Menu: 
       <a href="javascript:abrir('dip_formandos_etiquetas.asp','Etiquetas','width=200,height=200,scrollbars=no')">Imprimir Etiquesta Agendadas</a>  |  <a href="javascript:confirmarExclusaoAgendamento()">Limpar Etiquetas Agendadas</a></div>
	  <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	 
		<div id="comboTexto">
         
		  <%
		  AbrirConexao()
		  Set rsEtiquetasAgendadas = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO_ETIQUETAS_TEMPORARIA WHERE COD_USUARIO="&Session("COD_USUARIO")&" AND AGENDADA=1",qtdAgendada)
		  if not rsEtiquetasAgendadas.Eof Then%>
           <h3 align="center">Etiquetas Agendadas</h3>
           <table width="70%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
          <tr>
            <td><div align="center"><strong>C&oacute;d. Formando</strong></div></td>
            <td width="25%"><div align="center"><strong>Data gravada</strong></div></td>
          </tr>
          <%
		  While not rsEtiquetasAgendadas.Eof
		  %>
          <tr>
            <td><div align="center"><a href="dip_busca_detalhamento_formando.asp?cod_formando=<%=rsEtiquetasAgendadas("COD_FORMANDO")%>"><%=rsEtiquetasAgendadas("COD_FORMANDO")%></a></div></td>
            <td width="25%"><div align="center"><%=rsEtiquetasAgendadas("DATA_CADASTRO")%></div></td>
          </tr>
          <%
		  rsEtiquetasAgendadas.MoveNext()
		  Wend
		  %>
                    <tr>
            <td colspan="2"><div align="center"><strong>Qtd. total de etiquetas:  <%=qtdAgendada%></strong></div></td>
          </tr>
        </table>
          <p>
            <%
		Else
		%>
                </p>
          <p>&nbsp;</p>
          <p>&nbsp;</p>
          <p align="center" class="textoMedio"><strong><em>Não há etiquetas agendadas!</em></strong></p>
          <p align="center" class="textoMedio">&nbsp;</p>
          <p align="center" class="textoMedio">&nbsp;</p>
          <p>
                <%
		End If
		
		rsEtiquetasAgendadas.Close
		Set rsEtiquetasAgendadas=Nothing
		FecharConexao()
		%>
            </p>
		</div>
		
        <p align="center">
		  <input class="btnTopo" onclick="abre('#')" type="button" name="btnTopo2" value=" " />
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