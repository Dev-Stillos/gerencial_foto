<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>Pend&ecirc;ncias de Pagamento</title>
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
<style type="text/css">
<!--
.menu_paginacao{
	float: left;
	width: 300px;
	position: relative;
}
.menu_paginacao div{
	position: relative;
	float: left;
	height: 5px;
	width: 10px;
	border: 1px solid #666666;
	padding: 2px;
}
.menu_paginacao a{
	font-size:12px;
	text-decoration:none;
	color: #990000;
}

.menu_paginacao a:link{
	font-size:12px;
	text-decoration:none;
}

.menu_paginacao a:visited{
	font-size:12px;
	text-decoration:none;
}

.menu_paginacao a:active{
	font-size:12px;
	text-decoration:none;
}

.menu_paginacao a:hover{
	font-size:12px;
	text-decoration:none;
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

	  <h2>CAP - Pagamento de Consultores - Pend&ecirc;ncias de Pagamento</h2>
	   <div class="subMenu">Sub-Menu: <a href="cap_fechamentos.asp">Voltar para Fechamentos</a> | 
       <a href="cap_fechamentos_pendencias_formulario.asp">Cadastrar nova Pend&ecirc;ncia de Pagamento</a></div>
	  <%
	Dim sucesso, codmsg
	codmsg	= CStr(""&Request.QueryString("codmsg"))
	sucesso = CStr(""&Request.QueryString("sucesso"))
		
	If sucesso = "" and codmsg <> "" Then%>
    	<div align="center" class="msgGrandeSucesso">
        	<%
			If codmsg = "1" Then
				Response.Write("Pend�ncia de Pagamento inclu�da com sucesso!")
			Elseif codmsg = "2" Then
				Response.Write("Pend�ncia de Pagamento alterada com sucesso!")
			Elseif codmsg = "3" Then	
				Response.Write("Pend�ncia de Pagamento exclu�da com sucesso!")
			End If	
			%>
            
        
        </div>
    <%ElseIf sucesso <> "" and codmsg <> "" Then%>
        <div align="center" class="msgGrandeErro">
			<%
			If codmsg = "41" Then
				Response.Write("Voc&ecirc; est&eacute; tentando alterar/excluir uma pend&ecirc;ncia de pagamento cujo pagamento do m&ecirc;s de refer&ecirc;ncia j&aacute; foi conclu&iacute;do e por tanto a pend&ecirc;ncia n&atilde;o seria realizada.")	
			End If	
			%>
</div>
	<%End If%>
     
     
	  	  <%
		  AbrirConexao()
		  
		  Dim rs, rs2, sql
		  Dim pendencia_processada 'Vari�vel que verifica se o pagamento j� foi processado e por tanto a pendencia j� foi realizada
		  
				Set rs2=objConexao.Execute("SELECT data_pagamento FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE data_processado_pagamento IS NULL ORDER BY data_pagamento ASC ")
								
					If not rs2.Eof Then
						ULTIMO_PAGAMENTO_PROCESSADO = formatDateTime(rs2(0))
					End If
								
				rs2.Close
				Set rs2=Nothing
					  
		  
		  
		  	sql = "SELECT TB_FECHAMENTO_PENDENCIAS_PAGAMENTO.*, TB_VENDEDORES.NOME_VENDEDOR, TB_ADM_USUARIOS.NOME FROM TB_VENDEDORES INNER JOIN TB_FECHAMENTO_PENDENCIAS_PAGAMENTO ON TB_VENDEDORES.COD_VENDEDOR = TB_FECHAMENTO_PENDENCIAS_PAGAMENTO.COD_VENDEDOR INNER JOIN TB_ADM_USUARIOS ON TB_FECHAMENTO_PENDENCIAS_PAGAMENTO.CADASTRADO_POR = TB_ADM_USUARIOS.COD_USUARIO ORDER BY DATA_CADASTRO DESC"  
		  
		  
		  	Set rs=server.CreateObject("ADODB.recordset")
			rs.CursorLocation=3  'Cursor no cliente
			rs.PageSize=20
			rs.Open sql, objConexao
		  
		  	Dim PAGINA_ATUAL
			PAGINA_ATUAL = formatNumber(0&Request.QueryString("pagina"),0)
			
			If PAGINA_ATUAL > 0 Then
					rs.AbsolutePage=PAGINA_ATUAL
			Else
				PAGINA_ATUAL = 1
				rs.AbsolutePage=1
			End If		
		  
		  
		  If not rs.Eof then
		  %>
		  <h3 align="center">Rela&ccedil;&atilde;o de Pend&ecirc;ncias de Pagamentos</h3>
		  <div id="comboTexto">
		  <table width="95%" border="0" align="center">
  <tr>
    <td width="47%"> 
P&aacute;gina <strong><%=rs.AbsolutePage%></strong> de <%=rs.PageCount%><br /></td>
    <td width="53%"><div align="right">Total <strong><%=rs.RecordCount%></strong> fechamentos</div></td>
  </tr>
</table>
		  
          <table width="95%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="13%"><div align="center"><strong>Pagamento de</strong></div></td>
            <td width="18%"><div align="center"><strong>Vendedor</strong></div></td>
            <td width="8%"><div align="center"><strong>Tipo</strong></div></td>
            <td width="7%"><div align="center"><strong>Valor</strong></div></td>
            <td width="27%"><div align="center"><strong>Motivo</strong></div></td>
            <td width="11%"><div align="center"><strong>Data do Cadastro</strong></div></td>
            <td width="9%"><div align="center"><strong>Pend&ecirc;ncia Processada</strong></div></td>
            <td width="7%"><p>&nbsp;</p>
              <p>&nbsp;</p></td>
</tr>

		  		 <%
	 
				 
				  
		 Dim x
		 For x=1 To rs.PageSize
		    If Not rs.Eof Then
			
				Set rs2=objConexao.Execute("SELECT TB_FECHAMENTO_PAGAMENTO_CABECALHO.*, TB_ADM_USUARIOS.NOME FROM TB_ADM_USUARIOS INNER JOIN TB_FECHAMENTO_PAGAMENTO_CABECALHO ON TB_ADM_USUARIOS.COD_USUARIO = TB_FECHAMENTO_PAGAMENTO_CABECALHO.PAGAMENTO_GERADO_POR WHERE VENDEDOR_RECEBE_EM = "&sqlData(rs(0)))
				
					If not rs2.Eof Then
						
						DIA_VENCIMENTO = rs2("DIA_VENCIMENTO")
						PAGAMENTO_GERADO_POR = rs2("NOME")
						
					End If
				
				rs2.Close			
				Set rs2=Nothing
				%>
				<tr>
							<td>
							<%
							If isDate(rs("VENDEDOR_RECEBE_EM")) Then
							'Para pend�ncias antigas
							%>
                            	<%=Ucase(MonthName(Month(rs("VENDEDOR_RECEBE_EM"))))%> / <%=Year(rs("VENDEDOR_RECEBE_EM"))%>
                            <%
							Else
								Set rs2=objConexao.Execute("SELECT data_pagamento FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE cod_fechamento_periodo = "&rs("cod_fechamento_periodo"))
								
									If not rs2.Eof Then
										Response.Write(ArrumaData(rs2(0)))
									End If
									
								rs2.Close
								Set rs2=Nothing
							End If
							%>
                            </td>
							<td><div align="left"><%=rs("nome_vendedor")%></div></td>
							<td align="center">
							<%
							If Ucase(Cstr(rs("TIPO"))) = "A" Then
								Response.Write("Acr&eacute;scimo (+)")
							Else
								Response.Write("Decr&eacute;scimo (-)")
							End If
							%>
                            </td>
							<td><div align="right"><%=FormatCurrency(rs("VALOR"))%></div></td>
							<td><div align="left"><%=rs("MOTIVO")%></div></td>
							<td><div align="center"><%=ArrumaData(rs("DATA_CADASTRO"))%> / <%=rs("NOME")%> </div></td>
							<td><div align="center">
                            <%							
							If int(dateDiff("d",ULTIMO_PAGAMENTO_PROCESSADO,rs("VENDEDOR_RECEBE_EM"))) <= 0 Then
								pendencia_processada = true
								Response.Write("Sim")
							Else
								pendencia_processada = false
								Response.Write("N&atilde;o")
							End If
							%>
                            </div></td>
				            <td><div align="center">
                            <%If not pendencia_processada Then%>
                            	<a href="cap_fechamentos_pendencias_formulario.asp?comando=alterar&cod_pendencia=<%=rs("cod_pendencia")%>"><img src="images/icons/icon_alterar.png" border="0" /></a>
                                &nbsp;
                                <a href="cap_fechamentos_pendencias_formulario.asp?comando=excluir&cod_pendencia=<%=rs("cod_pendencia")%>"><img src="images/icons/icon_excluir.png" width="16" height="16" border="0" /></a>
                            <%End If%>
                            </div></td>
				</tr>
		        <%
				rs.MoveNext
				End If 
				Next	
      
		  %>
        </table>
		
		
		      <div align="left">
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
							<a href="?pagina=<%=x%>"><%=x%></a>
						<%Else%>
							<%=x%>
						<%End If%>
						
					<%Next%>
					</div>
					</td>
                      </tr>
                    </table>
					
					
					

				
				<%End if
		
		Else
		
		End If      
		
		
 		rs.Close()
		Set rs = Nothing
		FecharConexao()
		%>
		  <br /><br />
		      </div>
		  </div>
		  <p><strong>Observa&ccedil;&atilde;o:</strong><br />&Eacute; permitido a exclus&atilde;o/altera&ccedil;&atilde;o de pend&ecirc;ncia de pagamento somente se o pagamento do m&ecirc;s correspondente n&atilde;o tiver sido processado.</p>
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