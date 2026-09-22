<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Validação especial para verificar se possui sub-validação específica
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '15' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se não encontrar então redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
		
	rs.Close
	Set rs=Nothing
	
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
<!--
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

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->

function cadastraVendedor(valor){
	alert(valor);
}
</script>
		
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
	function popUp(URL){
	win=window.open(URL,"","width=550,height=350,scrollbars=yes")
	var text = "Se a janela não abrir desbloquei os Pop-ups!";
	if(win == null) { alert(text); return; }
	win.moveTo(350,150);
}	
	</script></head>

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
  BTN_IMPRIMIR = 1
  BTN_FONTE = 1
  
   BTN_IMPRIMIR_LINK = "href=""javascript:popUp('adm_relatorios_extrato_vendedores_IMPRIMIR.asp?COD_VENDEDOR="&COD_VENDEDOR&"&DATA_INI="&DATA_INI&"&DATA_FIM="&DATA_FIM&"&FILTRO_ALBUNS="&FILTRO_ALBUNS&"&FILTRO_AGREGADOS="&FILTRO_AGREGADOS&"&CHEQ_DEVOLVIDOS="&CHEQ_DEVOLVIDOS&"')"""
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">
	
	  <h2>Relat&oacute;rios - Atraso de Devolu&ccedil;&atilde;o </h2>
	  <div align="center" class="msg"><%=Request.QueryString("msg")%></div>
	   <%
		AbrirConexao()

		   Dim data_inicial
		   data_inicial = dateAdd("d",-45,Now())
		   data_inicial = Year(data_inicial)&"-"&Month(data_inicial)&"-"&Day(data_inicial)
				
				sql = "SELECT COD_VENDEDOR, NOME_VENDEDOR, (SUM( (DATEDIFF( NOW( ) , DATA_SAIDA ) ) -45) / COUNT( * ) ) AS MEDIA_DIAS_FORA, COUNT( * ) AS ITENS_PENDENTES FROM VW_MOVIMENTACOES_COMPLETA WHERE DATA_RETORNO IS NULL AND TIPO <> 1 AND DATA_SAIDA <= '"&data_inicial&" 23:59:59' AND COD_VENDEDOR <> 27 GROUP BY COD_VENDEDOR ORDER BY NOME_VENDEDOR"

		'Elseif CHEQ_DEVOLVIDOS = 3 Then
		'Devolvidos e Não devolvidos
			'Não faz nada!!!
		
		Set rs = objConexao.Execute(sql)
		
		If not rs.Eof Then
		%>

	    <h3 align="center">Relat&oacute;rio de Representantes Comerciais com itens  pendentes por mais de 45 dias</h3>
	  
	  <table width="100%" 
		      border="1" align="center"
					cellpadding="5" 
					style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
                <td width="63%" height="28"><div align="center"><span class="MensagemFicha"><strong>Representante Comercial </strong></span><strong><span class="MensagemFicha"></span></strong></div></td>
                <td colspan="2"><div align="center"><strong>Imprimir</strong></div></td>
                <td width="22%"><div align="left" class="MensagemFicha">
                  <div align="center" class="textoMedio"><strong>Qtd. Itens Pendentes </strong></div>
                </div></td>
                </tr>
		<%
		Dim QTD_GERAL
		QTD_GERAL = 0
		While not rs.Eof
		QTD_GERAL = QTD_GERAL + formatNumber(0&rs("ITENS_PENDENTES"),0)
		%>     
				<tr>
		          <td><%=rs("NOME_VENDEDOR")%></td>
                  <td width="8%"><div align="center"><a href="adm_relatorios_extrato_vendedores_atrasados_por_vendedor_html.asp?COD_VENDEDOR_BLOQUEADO=<%=rs("COD_VENDEDOR")%>&COM_COPIA=2" target="_blank"> 1 via</a></div></td>
                  <td width="7%"><div align="center"><a href="adm_relatorios_extrato_vendedores_atrasados_por_vendedor_html.asp?COD_VENDEDOR_BLOQUEADO=<%=rs("COD_VENDEDOR")%>&COM_COPIA=1" target="_blank">2 vias</a></div></td>
                  <td bgcolor="#FDDBDB"><div align="center" class="textoMedio"><strong><%'=formatNumber(rs("MEDIA_DIAS_FORA"),0)%><%=formatNumber(rs("ITENS_PENDENTES"),0)%></strong></div></td>
                  </tr>
					 
			<%
			rs.MoveNext()
			Wend
			
			rs.Close
			Set rs = Nothing
			
			FecharConexao()
			%>		 
        
		  	<tr bgcolor="whitesmoke">
		  				    <td colspan="4"><div align="center"><strong>Quantidade Total de Itens Pendentes: <%=QTD_GERAL%></strong></div></td>
                        <tr>
                          <td colspan="4"><strong>Observa&ccedil;&otilde;es:</strong><br />
Itens n&atilde;o devolvidos, mas j&aacute; lan&ccedil;ados pelo sistema pelo representante comercial<strong> constam</strong> como atrasados mesmo que n&atilde;o foram corrigidos pelo departamento fotogr&aacute;fico. <br />
Caso o item for vendido sua data retorno ser&aacute; a mesma data do lan&ccedil;amento do sistema pelo representante comercial, entretanto esta data  &eacute; lan&ccedil;ada no final da corre&ccedil;&atilde;o efetuada pelo depto. de foto e reportagens.</td>
                        </table>

		  <p></p>
		    <%Else%>
		   <div align="center">
           <br /><br /><br /><br />
		  <p><span class="textoMedio"><strong>Não há dados para os filtros selecionados!</strong></span></p>
          <br /><br /><br /><br />
		</div>
		    <%End If%>
	  
	  
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