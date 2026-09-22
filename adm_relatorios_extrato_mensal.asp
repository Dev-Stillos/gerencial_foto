<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Validação especial para verificar se possui sub-validação específica
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '43' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se não encontrar então redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
		
	rs.Close
	Set rs=Nothing
	
FecharConexao()

'/////////////////////////////////////////////////////////////////////////////////////////////

Dim PERIODO_REFERENCIA
PERIODO_REFERENCIA = "01/"&Request.QueryString("MES_INICIO")&"/"&Request.QueryString("ANO_INICIO")

If (PERIODO_REFERENCIA) <> "" Then
	PERIODO_REFERENCIA = formatDateTime(PERIODO_REFERENCIA)
End If

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="all"></link>
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
	function popUp(URL){
	win=window.open(URL,"","width=550,height=350,scrollbars=yes")
	var text = "Se a janela não abrir desbloquei os Pop-ups!";
	if(win == null) { alert(text); return; }
	win.moveTo(350,150);
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
  BTN_FONTE = 1
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">
	
	  <h2>Relatórios - Extrato Mensal de Produtos<br />
	    
	    
          <br />
      </h2>
	  <div align="center" class="msg"><%=Request.QueryString("msg")%></div>
	  <%
		AbrirConexao()		
		Dim rs, sql, i, produtos(100,5), QTD_PRODUTOS, TOTAL_VENDIDO
		TOTAL_VENDIDO = 0
		%>
        <span class="textoMaior"><strong><%=Ucase(MonthName(Month(PERIODO_REFERENCIA)))%> - <%=year(PERIODO_REFERENCIA)%></strong></span>
        <table width="100%" 
		      border="1" align="center"
					cellpadding="5" 
					style="border-collapse:collapse">
          <tr bgcolor="whitesmoke">
            <td width="20%" height="28"><div align="center"><strong>Fotos vendidas</strong>*</div></td>
            <td width="68%"><div align="center"><strong>Venda de Produtos </strong></div>                </td>
            <td width="12%"><div align="center"><strong>Valor Total Produtos**</strong></div></td>
          </tr>
          <%	
		'/////////////////////////////////////////////
		'FOTOS
		Dim TOTAL_GERAL_FOTOS_FATURADAS, TOTAL_FATURADO_ARRECADADO
		TOTAL_GERAL_FOTOS_FATURADAS = 0
		TOTAL_FATURADO_ARRECADADO = 0

			Set rs5=objConexao.Execute("SELECT SUM(FATURADO) AS FATURADO, SUM(FATURADO*VALOR_UN) AS TOTAL_FATURADO_ARRECADADO FROM VW_PEDIDOS_COMPLETO WHERE month(DATA_VENDA) = "&Month(PERIODO_REFERENCIA)&" AND year(DATA_VENDA) = "&Year(PERIODO_REFERENCIA)&" and FINALIZA_CORRECAO = 1")
				
				If not rs5.Eof Then
					
					TOTAL_GERAL_FOTOS_FATURADAS = formatNumber(0&rs5("FATURADO"),0)
					TOTAL_FATURADO_ARRECADADO = formatNumber(0&rs5("TOTAL_FATURADO_ARRECADADO"),0)
					
					TOTAL_VENDIDO = (TOTAL_VENDIDO+TOTAL_FATURADO_ARRECADADO)/1
					
				End If
				
			rs5.Close
			Set rs5=Nothing
		
		'/////////////////////////////////////////////
		
		%>
          <tr>
            <td><p>Total Fotos: <strong><%=TOTAL_GERAL_FOTOS_FATURADAS%></strong></p>
                <p>Valor Total Fotos: <strong><%=formatCurrency(TOTAL_FATURADO_ARRECADADO)%></strong></p></td>
            <td ><%
				  Set rs4=objConexao.Execute("SELECT TB_PEDIDOS_PRODUTOS.COD_PRODUTO, TB_PRODUTOS.NOME_PRODUTO, SUM(TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS QUANTIDADE, SUM(TB_PEDIDOS_PRODUTOS.VALOR_UN*TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS VALOR_TOTAL FROM TB_PEDIDOS  INNER JOIN TB_PEDIDOS_PRODUTOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO INNER JOIN TB_PRODUTOS ON TB_PEDIDOS_PRODUTOS.COD_PRODUTO = TB_PRODUTOS.COD_PRODUTO WHERE month(TB_PEDIDOS.DATA_VENDA) = "&Month(PERIODO_REFERENCIA)&" AND  year(TB_PEDIDOS.DATA_VENDA) = "&Year(PERIODO_REFERENCIA)&" and TB_PEDIDOS.FINALIZA_CORRECAO = 1 GROUP BY TB_PEDIDOS_PRODUTOS.COD_PRODUTO ",QTD_PRODUTOS)
				  
				 If not rs4.Eof Then
				 %>
                <table style="border-collapse:collapse; font-size:11px" width="100%" 
		      border="1" align="center"
					cellpadding="2" 
					>
                  <tr>
                    <td width="47%"><div align="center"><strong>Descri&ccedil;&atilde;o</strong></div></td>
                    <td width="28%"><div align="center"><strong>Quantidade</strong></div></td>
                    <td width="25%"><div align="center"><strong>Sub-Total</strong>**</div></td>
                  </tr>
                  <%
				 dIM VALOR_TOTAL 
				While not rs4.Eof
				
						VALOR_TOTAL = formatNumber(0&rs4("VALOR_TOTAL"))
						TOTAL_VENDIDO = TOTAL_VENDIDO + VALOR_TOTAL
				%>
                  <tr bgcolor="">
                    <td><div align="left"><%=rs4("NOME_PRODUTO")%></div></td>
                    <td><div align="center"><%=rs4("QUANTIDADE")%></div></td>
                    <td><div align="right"><%=formatCurrency(VALOR_TOTAL)%></div></td>
                  </tr>
                  <%	
						i=i+1
					rs4.MoveNext()
					Wend
					
				  rs4.Close
				  Set rs4=Nothing
				  %>
                </table>
              <%End If%>
              &nbsp; </td>
            <td><div align="right" class="textoMedio"><%=formatCurrency(TOTAL_VENDIDO)%></div></td>
          </tr>
        </table>
        <p>
        <strong>Observações:</strong><br /><br />
        1) Somente vendas já corrigidas pelo dpto. fotográfico aparecem nesta relação<br /><br />
        * Fotos j&aacute; produzidas e que foram vendidas, as outras fotos na coluna venda de produtos s&atilde;o fotos vendidas &agrave; parte.<br /><br />
          ** Valores baseados no valor declarado do produto no momento da venda, porém o valor faturado geralmente é arredondado por tanto gerando uma diferença entre Valor Total dos Produtos e Total Arrecadado.          </p>
		    <%
			FecharConexao()
			%>
		   
		    

		<div align="center"><br />
            <br />
            <input class="btnTopo" type="button" onclick="abre('#')" name="Submit2" value=" " />
            <br />
            <br />
            <br />
		  </div></td>
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






