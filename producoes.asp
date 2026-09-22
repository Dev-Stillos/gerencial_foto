<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
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

	  <h2>DIP - Produ&ccedil;&otilde;es</h2>
	   <div class="subMenu">Sub-Menu: <a href="dip_producoes_alteracoes.asp?comando=incluir">Incluir Produ&ccedil;&atilde;o</a>  
	     Filtrar:	<a href="?f=1">Somente Finalizadas</a> | <a href="?f=2">Somente N&atilde;o Finalizadas</a> | <a href="?">Limpar Filtro</a></div>
	   <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	  <%
		  AbrirConexao()
		  Dim rs, rs2, rs3, sql
		  Dim QTD_ALBUNS, FICHAS_VENDA, FILTRO_F
		  
		  FILTRO_F = formatNumber(0&Request.QueryString("f"),0)
		  
		  If FILTRO_F  = 1 then
		  'Somente Finalizadas
		  	sql = "SELECT * FROM TB_PRODUCAO WHERE PRODUCAO_FINALIZADA = 1 ORDER BY COD_PRODUCAO DESC"
		  
		  Elseif FILTRO_F = 2 then
		  'Somente Não Finalizadas
		  	sql = "SELECT * FROM TB_PRODUCAO WHERE PRODUCAO_FINALIZADA = 2 ORDER BY COD_PRODUCAO DESC"
		  Else
		  'Todas
		  	sql = "SELECT * FROM TB_PRODUCAO ORDER BY COD_PRODUCAO DESC"
		  End If
		  
		  
		  
		  	Set rs=server.CreateObject("ADODB.recordset")
			rs.CursorLocation=3  'Cursor no cliente
			rs.PageSize=500
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
		  <h3 align="center">Produ&ccedil;&otilde;es </h3>
		  <div id="comboTexto">
		  <table width="95%" border="0" align="center">
  <tr>
    <td width="47%"> 
Página <strong><%=rs.AbsolutePage%></strong> de <%=rs.PageCount%><br /></td>
    <td width="53%"><div align="right">Total <strong><%=rs.RecordCount%></strong> produ&ccedil;&otilde;es</div></td>
  </tr>
</table>
		  
          <table width="95%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="15%"><div align="center"><strong>C&oacute;d. Produ&ccedil;&atilde;o</strong></div></td>
            <td width="47%"><div align="center"><strong>Produ&ccedil;&atilde;o</strong></div></td>
            <td width="19%"><div align="center"><strong>Contrato(s)</strong></div></td>
            <td width="19%"><div align="center"><strong>Data Sa&iacute;da </strong></div></td>
            </tr>
		  		 <%  
		 Dim x
		 For x=1 To rs.PageSize
		    If Not rs.Eof Then
			
				Dim DESCRICAO, PRODUTOS_AGREDADOS, COD_PRODUCAO
				DESCRICAO = rs("DESCRICAO")
				PRODUTOS_AGREDADOS = rs("PRODUTOS_AGREDADOS")
				COD_PRODUCAO = rs("COD_PRODUCAO")
				FICHAS_COMPLETAS = 0
		 		%>
				<tr
						  <%
						  Dim cod_prod_Query, cod_prod_rs
						  cod_prod_Query	= formatNumber(0&Request.QueryString("cod_prod"),0)
						  cod_prod_rs		= formatNumber(0&rs("COD_PRODUCAO"),0)
						
						  if cod_prod_Query = cod_prod_rs then%> bgcolor="#FFFF99"<%end if%>>
							<td><div align="center"><%=COD_PRODUCAO%></div></td>
							<td><%=Ucase(DESCRICAO)%></td>
							<td>
							<%
							Dim rsContratos, qtdContratos, CONTRATOS_ULT
							
							Set rsContratos = objConexao.Execute("SELECT DISTINCT(COD_CONTRATO) AS COD_CONTRATO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO="&rs("COD_PRODUCAO")&" ORDER BY COD_CONTRATO ASC",qtdContratos)
							
							If not rsContratos.Eof then
							
								If qtdContratos > 1 then
							
									CONTRATOS_ULT = ""
									For i = 0 to qtdContratos-1
										
										If i = qtdContratos-1 then
										
											CONTRATOS_ULT = CONTRATOS_ULT&rsContratos("COD_CONTRATO")
										
										else
										
											CONTRATOS_ULT = CONTRATOS_ULT&rsContratos("COD_CONTRATO")&", "
										
										End If
									
									rsContratos.MoveNext()
									Next
									
									Response.Write(CONTRATOS_ULT)
									
								else
									
									Response.Write(rsContratos("COD_CONTRATO"))
								
								End If
								
									
									
							else
							
							Response.Write("-")
							
							End If
							
							rsContratos.Close()
							Set rsContratos = Nothing
							
							'Dim CONTRATOS_TRATAMENTO
							'CONTRATOS_TRATAMENTO = Left(CONTRATOS_ULT,Len(CONTRATOS_ULT)-2)&Replace(Right(CONTRATOS_ULT,2),", ","")
							'Response.Write(CONTRATOS_TRATAMENTO)
							
							%>			</td>
							<td><div align="center"><%=ArrumaData(rs("DATA_SAIDA"))%>			</div></td>
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
					Ir para a página:
					
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
					
					
					

				
				<%End if
		
		Else
		
		End If      
		
		
 		rs.Close()
		Set rs = Nothing
		FecharConexao()
		%>
		            
		      </div>
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