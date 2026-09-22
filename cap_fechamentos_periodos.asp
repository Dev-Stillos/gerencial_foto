<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>Per&iacute;odo de Vendas</title>
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
  BTN_IMPRIMIR_LINK = "<a href='cap_fechamentos_periodos_imprimir.asp' target='_blank'"
  %>
  <!-- #include file="includes/sub_menu.asp" -->
   
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">
 
	  <h2>CAP - Pagamento de Consultores - Per&iacute;odos de Vendas</h2>
	  <div class="subMenu">Sub-Menu: <a href="cap_fechamentos.asp">Voltar para Fechamentos</a> | 
       <a href="cap_fechamentos_periodos_formulario.asp">Cadastrar novo Per&iacute;odo de Vendas</a></div>
      
	  <%
	   Response.Write("<br />")
	Dim sucesso, codmsg
	codmsg	= CStr(""&Request.QueryString("codmsg"))
	sucesso = CStr(""&Request.QueryString("sucesso"))
		
	If sucesso = "" and codmsg <> "" Then%>
    	<div align="center" class="msgGrandeSucesso">
        	<%
			If codmsg = "1" Then
				Response.Write("Per&iacute;odo de Venda inclu&iacute;do com sucesso!")
			Elseif codmsg = "2" Then
				Response.Write("Per&iacute;odo de Venda alterado com sucesso!")
			Elseif codmsg = "3" Then	
				Response.Write("Per&iacute;odo de Venda exclu&iacute;do com sucesso!")
			End If	
			%>
            
        
        </div>
    <%ElseIf sucesso <> "" and codmsg <> "" Then%>
        <div align="center" class="msgGrandeErro">
			<%
			If codmsg = "3" Then
				Response.Write("Per&iacute;odo de Vendas n&atilde;o exclu&iacute;do, pois h&aacute; pedidos com este per&iacute;odo!")
			ElseIf codmsg = "4" Then
				Response.Write("Per&iacute;odo de Vendas n&atilde;o inclu&iacute;do, pois &eacute; permitido somente 1 pagamento por dia, devido os recibos de venda serem gerados com a data do pagamento!")
			
			'/////////
			ElseIf codmsg = "401" Then
				Response.Write("Data in&iacute;cio maior que Data T&eacute;rmino do Per&iacute;odo de Venda N&Atilde;O inclu&iacute;do!")
			ElseIf codmsg = "402" Then
				Response.Write("Data t&eacute;rmino de venda maior que data limite de lan&ccedil;amento de Venda N&Atilde;O inclu&iacute;do!")
			ElseIf codmsg = "403" Then
				Response.Write("Data do pagamento n&atilde;o pode ser inferior a data m&aacute;xima de lan&ccedil;amento de vendas N&Atilde;O inclu&iacute;do!")
			ElseIf codmsg = "404" Then
				Response.Write("Voc&ecirc; est&aacute; tentando inserir um per&iacute;odo, por&eacute;m j&aacute; existe um outro Per&iacute;odo de Venda dentro do Per&iacute;odo informado, n&atilde;o &eacute; permitido misturar as datas dos per&iacute;odos de venda. Por tanto per&iacute;odo N&Atilde;O inclu&iacutedo!")
			ElseIf codmsg = "405" Then
				Response.Write("Voc&ecirc; est&aacute; tentando inserir um per&iacute;odo com meses diferentes n&atilde;o &eacute; permitido data in&iacute;cio e data t&eacute;rmino terem meses diferentes. Por tanto per&iacute;odo N&Atilde;O inclu&iacute;do!")
			End If	
			%>
</div>
	<%End If%>
    
    
			
            <%
			AbrirConexao()
			Dim rsContratos, qtdContratos, CONTRATOS_ULT, qtdProdutos
			
			sql = "SELECT * FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS order by data_termino desc"
			
			Set rs=server.CreateObject("ADODB.recordset")
			rs.CursorLocation=3  'Cursor no cliente
			rs.PageSize=50
			rs.Open sql, objConexao
		  
		  	Dim PAGINA_ATUAL
			PAGINA_ATUAL = formatNumber(0&Request.QueryString("pagina"),0)
			
			
			If not rs.Eof Then%>
               <table width="95%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="11%"><div align="center"><strong>C&oacute;d. Per&iacute;odo de Venda</strong></div></td>
            <td width="12%"><div align="center"><strong>Data In&iacute;cio Vendas</strong></div></td>
            <td width="12%"><div align="center"><strong>Data T&eacute;rmino Vendas</strong></div></td>
            <td width="10%"><div align="center"><strong>Data limite de Lan&ccedil;amento</strong> <strong>no sistema</strong></div></td>
            <td width="11%"><div align="center"><strong>Data que ser&aacute; ou foi o pagamento</strong></div></td>
            <td width="17%"><div align="center"><strong>Data do Processado</strong> <strong>do Pagamento</strong></div></td>
            <td width="10%"><div align="center"><strong>Cadastrado por</strong></div></td>
            <td width="10%"><div align="center" style="font-weight: bold">Data da &Uacute;ltima Intera&ccedil;&atilde;o</div></td>
            <td width="7%"><div align="center"><strong> Excluir</strong></div></td>
</tr>
		  <%
		  	Dim x
		 	For x=1 To rs.PageSize
		    If Not rs.Eof Then
		  %>
           <tr>
            <td><div align="center"><%=rs(0)%></div></td>
            <td><div align="center"><%=ArrumaData(rs("data_inicio"))%></div></td>
            <td>
           	  <div align="center"><%=ArrumaData(rs("data_termino"))%>&nbsp;</div></td>
            <td><div align="center"><%=ArrumaData(rs("data_termino_lancamento_vendas"))%></div></td>
            <td><div align="center"><%=ArrumaData(rs("data_pagamento"))%></div></td>
            <td><div align="center"><%=ArrumaData(rs("data_processado_pagamento"))%></div></td>
            <td><div align="center">
              <%	
			Set rs2=objConexao.Execute("Select nome From TB_ADM_USUARIOS where COD_USUARIO = '"&rs("cod_usuario")&"'")
				
				If not rs2.Eof Then
					Response.Write(rs2(0))
				End If
				
			rs2.Close
			Set rs2=Nothing
			%>
            </div></td>
            <td><div align="center"><%=ArrumaData(rs("modificado"))%></div></td>
            <td>
              <div align="center">
                    <a href="cap_fechamentos_periodos_formulario.asp?cod_fechamento_periodo=<%=rs("cod_fechamento_periodo")%>&comando=excluir">
                          <img alt="Excluir" src="images/icons/icon_excluir.png" border="0" /></a>                </div></td>
           </tr>
		  <%
				rs.MoveNext()
				End If 
				Next
		  %>
        </table>
        
        <br />
<br />
  <strong>Dica</strong>: Clique no &iacute;cone de impress&atilde;o para imprimir esta listagem de uma forma mais elegante.
        		
                		        <%If rs.PageCount > 1 Then %>
					<br />
                  
                    
					<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td>
						<div class="menu_paginacao">
					Ir para a P&aacute;gina:
					
					<%
					Dim aae
					aae = 0
					For x=1 To rs.PageCount 
					aae = formatNumber(0&x,0)
					
						If int(PAGINA_ATUAL) <> int(aae) Then
						'Response.Write(PAGINA_ATUAL&":"&x&"<br />")
						%>
							<a href="?pagina=<%=x%>&f=<%=FILTRO_F%>"><%=x%></a>
						<%Else%>
							<%=x%>
						<%End If%>
						
					<%Next%>
					</div>
					</td>
                      </tr>
                    </table>
				<%End if%>
        
        
		      <%else%>
              
                     
                     <div align="center">
                      <div class="msgGrande">N&atilde;o h&aacute; per&iacute;odos de vendas cadastrados!<br />
Para incluir clique no link &quot;Cadastrar novo  Per&iacute;odo de Vendas&quot; acima.</div>

                     </div> 
 
                    
		      <%end if
			  
			  
			  	  rs.Close()
				  Set rs = Nothing
				  FecharConexao()
			  %>
		     
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