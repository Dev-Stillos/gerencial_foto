<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Valida��o especial para verificar se possui acesso a p�gina m�e de formul�rio
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS where COD_PAGINA = '224' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se n�o encontrar ent�o redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
		
	rs.Close
	Set rs=Nothing
	
FecharConexao()
'/////////////////////////////////////////////////////////////////////////////////////////////
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="all"></link>
<title>Per&iacute;odos de Vendas</title>
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
</head>

<body style="background:#FFF">
<!-- End code for the left panel -->
<table cellpadding="0px" cellspacing="0" bgcolor="#FFFFFF" width="100%">
<tr>
<td>
<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="10">
      <tr>
        <td>
          
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td valign="top">
                <div align="center"><img src="images/layout/logomarca_foto_reportagens.jpg" width="92" height="93" />
                </div>
                <h3 align="center">Per&iacute;odos de Vendas</h3>
                <div align="center"><em>Listagem de períodos cujos pagamentos não foram processados</em><br />
&nbsp;</div>
                <%
			AbrirConexao()
			Dim rsContratos, qtdContratos, CONTRATOS_ULT, qtdProdutos
			
			sql = "SELECT * FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE data_processado_pagamento is null  order by data_termino desc"
			
			Set rs=objConexao.Execute(sql)
			
			
			If not rs.Eof Then%>
                <table width="95%" cellspacing="0" border="0" align="center" 	cellpadding="5">
  <tr bgcolor="whitesmoke">
    <td width="12%"><div align="center"><strong>Cód. Período de Venda</strong></div></td>
    <td width="45%"><div align="center"><strong>Per&iacute;odo</strong>
    <em><br />
    (Data da Venda)</em>
    </div></td>
    <td width="16%"><div align="center"><strong>Data limite de Lan&ccedil;amento</strong> <strong>no sistema</strong></div></td>
    <td width="13%"><div align="center"><strong>Data Pagamento</strong></div></td>
    <td width="14%"><div align="center"><strong>Cadastrado por</strong></div></td>
    </tr>
                  <%
		  	While not rs.Eof
		  %>
                  <tr>
                    <td style="border-bottom:solid 1px #CCC"><div align="center"><%=rs(0)%></div></td>
                    <td style="border-bottom:solid 1px #CCC"><div align="center" class="textoMedio"><strong><%=ArrumaData(rs("data_inicio"))%> à <%=ArrumaData(rs("data_termino"))%></strong></div></td>
                    <td style="border-bottom:solid 1px #CCC"><div align="center"><%=ArrumaData(rs("data_termino_lancamento_vendas"))%></div></td>
                    <td style="border-bottom:solid 1px #CCC"><div align="center"><%=ArrumaData(rs("data_pagamento"))%></div></td>
                    <td style="border-bottom:solid 1px #CCC"><div align="center">
                      <%	
			Set rs2=objConexao.Execute("Select nome From TB_ADM_USUARIOS where COD_USUARIO = '"&rs("cod_usuario")&"'")
				
				If not rs2.Eof Then
					Response.Write(rs2(0))
				End If
				
			rs2.Close
			Set rs2=Nothing
			%>
                    </div></td>
                    </tr>
                  <%
				rs.MoveNext()
				Wend
		  %>
                  </table>
                <br />
<br />

                Impresso em: <%=Now()%>
			<%
			  End If
			  
			  	  rs.Close()
				  Set rs = Nothing
				  FecharConexao()
			  %>
                
                <p align="center">&nbsp;</p>
                </td>
              
              </tr>
  </table>
          
          
          
          </td>
        </tr>
</table>

  
  </td>
  </tr>
</table></td>
</tr>
</table>
</body>
</html>