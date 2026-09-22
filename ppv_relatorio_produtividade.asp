<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Validação especial para verificar se possui sub-validação específica
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '59' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se não encontrar então redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
		
	rs.Close
	Set rs=Nothing
	
FecharConexao()

'/////////////////////////////////////////////////////////////////////////////////////////////

Dim DATA_INICIO, DATA_FIM
Dim diferenca
DATA_INICIO = ""&Request.QueryString("DATA_INI")
DATA_FIM = ""&Request.QueryString("DATA_FIM")

If DATA_INICIO = "" OR DATA_FIM = "" Then
	Response.Redirect("adm_relatorios.asp?msg=Atenção preencha a data de início e fim para gerar o relatório de venda de produtos")
Else

	If isDate(DATA_INICIO) and isDate(DATA_FIM) Then
		DATA_INICIO		= formatDateTime(DATA_INICIO)
		DATA_FIM		= formatDateTime(DATA_FIM)
		diferenca		= dateDiff("d",DATA_INICIO,DATA_FIM)
	Else
		Response.Redirect("adm_relatorios.asp?msg=Atenção formato de data inválido preencha no formato DD/MM/AAAA")
	
	End If	
		
End If

If diferenca < 0 Then
Response.Redirect("adm_relatorios.asp?msg=Atenção data de início é maior qua a data fim do intervalo de datas.")
End If


'Response.Write(">>>>>>>>>>>>> "&diferenca)

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
	
	  <h2>Relatório de Pordutividade de Funcion&aacute;rios por OS<br />
	    
	    
          <br />
      </h2>
	  <div align="center" class="msg"><%=Request.QueryString("msg")%></div>
	  <%
		AbrirConexao()		
		Dim rs0, sql, rs, sqlWhere
		Dim qtd_fora_prazo, qtd_dentro_prazo, qtd_total
		 
				
		Set rs0=objConexao.Execute("SELECT cod_usuario_responsavel FROM `TB_ORDENS_SERVICO_ETAPAS` WHERE data_inicio BETWEEN "&sqlData(DateAdd("d", -1, DATA_INICIO))&" AND "&sqlData(DateAdd("d", 1, DATA_FIM))&" GROUP BY cod_usuario_responsavel ")
		If not rs0.Eof Then
		%>
        <span class="textoMaior"><strong><%=ArrumaData(DATA_INICIO)%> à <%=ArrumaData(DATA_FIM)%></strong></span>
        <table width="100%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">              
          	<tr bgcolor="whitesmoke"> 
            	<td width="40%" rowspan="2"><div align="center"><strong>Usu&aacute;rio</strong></div></td>
            	<td colspan="3"><div align="center"><strong>Quantidade de O.S.</strong></div></td>
          	</tr>
            <tr bgcolor="whitesmoke">
            	<td width="22%"><div align="center"><strong>Dentro do Prazo</strong></div></td>
				<td width="21%"><div align="center"><strong>Excederam o prazo</strong></div></td>
            	<td width="17%"><div align="center"><strong>Total</strong></div></td>
        	</tr>
		<%
		While not rs0.Eof
			sqlWhere = " data_inicio BETWEEN "&sqlData(DateAdd("d", -1, DATA_INICIO))&" AND "&sqlData(DateAdd("d", 1, DATA_FIM))&" AND cod_usuario_responsavel = "&sqlNumero(rs0("cod_usuario_responsavel"))&" "
		%>
          <tr>
            <td class="textoMedio" valign="top"><div align="left"><%=Ucase(nomeUsuario(rs0(0),True))%></div></td>
            <td class="textoMedio" valign="top"><div align="center"><strong><%
			Set rs = Objconexao.Execute("SELECT COUNT(*) FROM `TB_ORDENS_SERVICO_ETAPAS` WHERE (DATEDIFF(`data_termino`,`data_inicio`) >= 0 or DATEDIFF(`data_termino`,`data_inicio`) = NULL) AND ("&sqlWhere&") GROUP BY cod_usuario_responsavel")
			
				If not rs.Eof Then
					qtd_dentro_prazo = formatNumber(0&rs(0),0)
					response.Write(rs(0))
				Else
					qtd_dentro_prazo = 0
					response.Write("0")
				End if
			rs.close()
			Set rs = nothing
			%></strong></div></td>
           	<td width="21%"><div align="center">
            <%
			Set rs = Objconexao.Execute("SELECT COUNT(*) FROM `TB_ORDENS_SERVICO_ETAPAS` WHERE "&sqlWhere&" GROUP BY cod_usuario_responsavel")
				If not rs.Eof Then
					qtd_total = formatNumber(0&rs(0),0)
				Else
					qtd_total = 0
				End if
			rs.close()
			Set rs = nothing
			
			qtd_fora_prazo = (qtd_total-qtd_dentro_prazo)/1
			Response.Write(qtd_fora_prazo)
			%>
            </div></td>            
            <td width="17%"><div align="center">
            <%
			Set rs = Objconexao.Execute("SELECT COUNT(*) FROM `TB_ORDENS_SERVICO_ETAPAS` WHERE "&sqlWhere&" GROUP BY cod_usuario_responsavel")
				If not rs.Eof Then
					response.Write(rs(0))
				Else
					response.Write("0")
				End if
			rs.close()
			Set rs = nothing
			%>
            </strong></div></td>
          </tr>
        <%
		rs0.MoveNext()
		Wend
		
		%>
        </table>
        <p>&nbsp;</p>
        <%Else%>
        		
                <div style="padding:50px; font-size:16px;" align="center">Não foi encontrado resultados para os parametros informados de <strong><%=ArrumaData(DATA_INICIO)%></strong> à <strong><%=ArrumaData(DATA_FIM)%></strong></div>
                
        <%End If
		
		rs0.Close
		Set rs0=Nothing	
			
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






