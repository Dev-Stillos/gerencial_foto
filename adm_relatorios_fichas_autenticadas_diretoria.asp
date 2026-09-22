<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Validação especial para verificar se possui sub-validação específica
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '13' and cod_perfil = '"&Session("cod_perfil")&"'")
		
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
<title>StillosWEB - Sistema corporativo - Versão 1.2</title>
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
	</script>

        <style type="text/css">
<!--
.style1 {
	color: #009900;
	font-weight: bold;
}
.style2 {
	color: #CC0000;
	font-weight: bold;
}
-->
        </style>
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
  BTN_FONTE = 1
  BTN_IMPRIMIR_LINK = "#"
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">
	
	  <h2>Relatórios - Fichas autenticadas pela Diretoria </h2>
	  <div align="center" class="msg"><%=Request.QueryString("msg")%></div>
		<br />
		<br />
		<%
		
		Dim sql, CONSULTAR
		sql = "SELECT * FROM VW_PEDIDOS_AUTENTICADOS_DIRETORIA "
		CONSULTAR = "N"
		
		
		If Request.QueryString("filtro1") <> Empty then
		CONSULTAR = "S"	
			sql=sql&" LIMIT 100"
		
		elseif Request.QueryString("filtro2") <> Empty then
		
				Dim DATA_INI, DATA_FIM
				DATA_INI = Request.QueryString("DATA_INI")
				DATA_INI = Year(DATA_INI)&"-"&Month(DATA_INI)&"-"&Day(DATA_INI)&" 00:00:00"
				DATA_FIM = Request.QueryString("DATA_FIM")
				DATA_FIM = Year(DATA_FIM)&"-"&Month(DATA_FIM)&"-"&Day(DATA_FIM)&" 23:59:59"
				
				If DATA_FIM < DATA_INI then
				CONSULTAR = "N"
				Dim msgErro
					msgErro = "Atenção você selecionou uma data inicial maior que a data final<br />Data selecionada:<br /> Início: "&Request.QueryString("DATA_INI")&" Final: "&Request.QueryString("DATA_FIM")
				else
					msgErro = "Você selecionou um período que não há registros:<br /> Início: "&Request.QueryString("DATA_INI")&" Final: "&Request.QueryString("DATA_FIM")			
					sql=sql&" Where DATA_CORRIGIDO_DIRETORIA BETWEEN '" & DATA_INI & "' AND '" & DATA_FIM&"'"
				CONSULTAR = "S"
				End If
		End If
		%>
		<div align="center" id="comboTexto2">
		<%
		  AbrirConexao()
		  Dim rsAcessoLog2, qtdRegistros
		  'Response.Write(sql)
		  Set rsAcessoLog2 = objConexao.Execute(sql,qtdRegistros)
		%>

		<%IF CONSULTAR = "S" and qtdRegistros > 0 THEN %>
		<table width="80%" 
		      border="1" align="center"
					cellpadding="5" 
					style="border-collapse:collapse">
									<tr>
            <td colspan="3"><div align="center"><strong>Qtd. total de registros: <%=qtdRegistros%></strong></div></td>
          </tr>
<tr bgcolor="whitesmoke">
            <td width="35%"><div align="center"><strong>Data / Hora da Autenticação</strong> (Ordem decrescente) </div></td>
            <td width="14%"><div align="center"><strong>Cód Pedido </strong></div></td>
            <td width="51%"><div align="center"><strong>Vendedor</strong></div></td>
            </tr>
		 <%
		  While not rsAcessoLog2.Eof
		  %>
          <tr>
            <td><%=ArrumaData(rsAcessoLog2("DATA_CORRIGIDO_DIRETORIA"))%></td>
            <td><div align="center"><%=rsAcessoLog2("COD_PEDIDO")%></div></td>
            <td><div align="left"><%=rsAcessoLog2("NOME_VENDEDOR")%></div></td>
            </tr>
		  <%
		  rsAcessoLog2.MoveNext()
		  Wend
	      
		  rsAcessoLog2.Close()
		  Set rsAcessoLog2 = Nothing
		  FecharConexao()
		  %>
          <tr bgcolor="whitesmoke">
            <td colspan="3"><div align="center"><strong>Qtd. total de registros: <%=qtdRegistros%></strong></div></td>
          </tr>
        </table>
		<p>
		  <input class="btnTopo" onclick="abre('#')" type="button" name="btnTopo2" value=" " />
		  <%else%>
		    <%Response.Redirect("adm_relatorios.asp?msg="&msgErro)%>
			  </p>
		</div>
		<%end if%>
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



