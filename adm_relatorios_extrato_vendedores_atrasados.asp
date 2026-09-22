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


<!--
'*** Instruções
'1) Relaciona todos os vendedores
'2) Mostra somentes os vendedores que possuem pedidos pendentes
-->

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
function popUp(URL){
	win=window.open(URL,"","width=550,height=350,scrollbars=yes")
	var text = "Se a janela não abrir desbloquei os Pop-ups!";
	if(win == null) { alert(text); return; }
	win.moveTo(350,150);
}	
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
  BTN_IMPRIMIR = 1
  BTN_FONTE = 1
  BTN_IMPRIMIR_LINK = "href=""javascript:popUp('adm_relatorios_extrato_vendedores_atrasados_IMPRIMIR.asp')"""
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">

	  <h2>Relatório de Produtos Atrasados - Todos os vendedores </h2>
	  <h3 align="center">Vendedores com produtos  a mais de 30 dias (exceto fichas) </h3>
		    
		      <div align="center">
		        <%
		   AbrirConexao()

		   'Instruções:
		   Dim data_inicial
		   data_inicial = dateAdd("d",-30,Now())
		   data_inicial = Year(data_inicial)&"-"&Month(data_inicial)&"-"&Day(data_inicial)
		   
		   sql ="SELECT * FROM VW_MOVIMENTACOES_COMPLETA "&_
		   		"WHERE FINALIZA_VENDEDOR <> 1 AND DATA_RETORNO IS NULL AND TIPO <> 1 AND DATA_SAIDA <= '"&data_inicial&" 23:59:59' ORDER BY NOME_VENDEDOR"
		   Set rs4 = objConexao.Execute(sql,qtdItensAtrasados)
	%>		
	  <table width="98%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
    <td colspan="7"><div align="center" class="msg"><strong>Quantidade de produtos atrasados:</strong> <span class="textoMaior"><%=qtdItensAtrasados%></span> </div></td>
    </tr>
<tr bgcolor="whitesmoke">
    <td width="20%"><div align="center"><strong>Nome Vendedor </strong></div></td>
    <td width="20%"><div align="center"><strong>Nome Formando </strong></div></td>
    <td width="27%"><div align="center"><strong>Instituição de Ensino</strong> </div></td>
    <td width="6%"><div align="center"><strong>Álbum</strong></div></td>
    <td width="8%"><div align="center"><strong>Item Pendênte </strong></div></td>
    <td width="13%"><div align="center"><strong>Data Saída </strong></div></td>
    <td width="6%"><div align="center"><strong>Qtd de dias fora </strong></div></td>
  </tr>
	<%While not rs4.Eof
	
	Dim COD_FORMANDO, NOME_VENDEDOR, NOME_FORMANDO, DESCRICAO, ALBUM, TIPO, DATA_SAIDA, ENCADERNACOES_EXTRAS
	Dim QTD_DIAS_FORA
	
	COD_FORMANDO			= rs4("COD_FORMANDO")
	NOME_VENDEDOR			= rs4("NOME_VENDEDOR")
	NOME_FORMANDO			= rs4("NOME_FORMANDO")
	DESCRICAO				= rs4("DESCRICAO")	
	ALBUM					= rs4("ALBUM")
	TIPO					= formatNumber(0&rs4("TIPO"),0)
	DATA_SAIDA				= rs4("DATA_SAIDA")
	ENCADERNACOES_EXTRAS	= formatNumber(0&rs4("ENCADERNACOES_EXTRAS"),0)
	
	If TIPO = 0 Then
		
		If ENCADERNACOES_EXTRAS = 1 Then
			TIPO = "Álbum Duplo"
		Else
			TIPO = "Álbum"
		End If
		
	Elseif TIPO > 1 Then
		'Verifica qual o produto agregado
			Dim rsProdAgrePend
			sql = 	"SELECT TB_PRODUTOS.NOME_PRODUTO "&_
					"FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = "&_
					"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO WHERE "&_
					"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO = "&COD_FORMANDO&" AND "&_
					"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.TIPO = "&TIPO&" LIMIT 1"
			Set rsProdAgrePend = objConexao.Execute(sql)
				If not rsProdAgrePend.Eof Then
				%>
				<%TIPO = rsProdAgrePend("NOME_PRODUTO")%>			
			    <%
				End If
			rsProdAgrePend.Close
			Set rsProdAgrePend = Nothing
		
	End If 	
	
	QTD_DIAS_FORA = dateDiff("d",DATA_SAIDA,Now())
	%>
<tr>
  <td><%=NOME_VENDEDOR%></td>
  <td><%=NOME_FORMANDO%></td>
  <td><%=DESCRICAO%></td>
  <td><div align="center"><%=ALBUM%></div></td>
  <td><div align="center"><%=TIPO%></div></td>
  <td><div align="center"><%=ArrumaData(DATA_SAIDA)%></div></td>
  <td><div align="center" class="textoMedio"><strong><%=QTD_DIAS_FORA%></strong></div></td>
</tr>

    	
		   <%
		  	rs4.MoveNext()
			Wend
		  
%>
</table>
<%
		  FecharConexao()		  
		  %>
		          </div>
		      <p align="center">
		  <input class="btnTopo" onclick="abre('#')" type="button" name="btnTopo2" value=" " />
		</p>		</td>

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



