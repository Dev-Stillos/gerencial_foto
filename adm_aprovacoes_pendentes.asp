<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
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
	--></head>

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

	  <h2>ADM - Aprova&ccedil;&otilde;es Pendentes e Recentes</h2>
	  <%
	   Response.Write("<br />")
	Dim sucesso, codmsg
	codmsg	= CStr(""&Request.QueryString("codmsg"))
	sucesso = CStr(""&Request.QueryString("sucesso"))
		
	If sucesso = "" and codmsg <> "" Then%>
    	<div align="center" class="msgGrandeSucesso">
        	<%
			If codmsg = "1" Then
			Elseif codmsg = "2" Then
				Response.Write("Aprova��o/Reprova��o gravada com sucesso!")
			Elseif codmsg = "3" Then	
			End If	
			%>
            
        
        </div>
    <%ElseIf sucesso <> "" and codmsg <> "" Then%>
        <div align="center" class="msgGrandeErro">
			<%
			If codmsg = "4" Then
			Elseif codmsg = "5" Then		
			End If	
			%>
</div>
	<%End If%>
			
            <%
			AbrirConexao()
			sql = "SELECT TB_APROVACOES.*, TB_APROVACOES_CATEGORIA.descricao_geral, TB_APROVACOES_CATEGORIA.descricao_interna_sobre_parametro, TB_APROVACOES_CATEGORIA.link_consulta_para_responsavel FROM TB_APROVACOES_CATEGORIA INNER JOIN TB_APROVACOES ON TB_APROVACOES_CATEGORIA.COD_CATEGORIA = TB_APROVACOES.COD_CATEGORIA where TB_APROVACOES.status = 0 ORDER BY TB_APROVACOES.data_cadastro DESC"
			
			Set rs=server.CreateObject("ADODB.recordset")
			rs.CursorLocation=3  'Cursor no cliente
			rs.PageSize=50
			rs.Open sql, objConexao
		  
		  	Dim PAGINA_ATUAL
			PAGINA_ATUAL = formatNumber(0&Request.QueryString("pagina"),0)
			
			
			If not rs.Eof Then%>
            	<h3 align="center">Aprova&ccedil;&atilde;es Pendentes</h3>
               <table width="95%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="30%"><div align="center"><strong>Tipo de Aprova&ccedil;&atilde;o</strong></div></td>
            <td width="22%"><div align="center"><strong>Solicitado por</strong></div></td>
            <td width="17%"><div align="center"><strong>Data Solicita&ccedil;&atilde;o</strong></div></td>
            <td width="17%"><strong>Respons&aacute;vel pela aprova&ccedil;&atilde;o/reprova&ccedil;&atilde;o</strong></td>
            <td width="14%"><div align="center"><strong>Aprovar/Reprovar</strong></div></td>
</tr>
		  <%
		  	Dim x
		 	For x=1 To rs.PageSize
		    If Not rs.Eof Then
		  %>
           <tr>
            <td><div align="left"><%=rs("descricao_geral")%></div></td>
            <td><div align="left">
<%
									Set rs2=objConexao.Execute("Select nome FROM TB_ADM_USUARIOS WHERE COD_USUARIO = "&rs("solicitante_cod_usuario"))
										
										if not rs2.Eof Then
											Response.Write(rs2(0))
										End if
									
									rs2.Close
									Set rs2=Nothing
									%>
            
            </div></td>
            <td><div align="center"><%=ArrumaData(rs("data_cadastro"))%></div></td>
            <td class="textoMedio"><div align="left">
              <%
									Set rs2=objConexao.Execute("Select nome FROM TB_ADM_USUARIOS WHERE COD_USUARIO = "&rs("responsavel_cod_usuario"))
										
										if not rs2.Eof Then
											Response.Write(rs2(0))
										End if
									
									rs2.Close
									Set rs2=Nothing
									%>
            </div></td>
            <td>
            	<div align="center">
            	  <%If int(Session("cod_usuario")) = int(rs("responsavel_cod_usuario")) Then%>
            	  <a href="adm_aprovacoes_pendentes_formulario.asp?cod_aprovacao=<%=rs("cod_aprovacao")%>">Aprovar/Reprovar</a>
            	      <%End If%>
          	  </div></td>
           </tr>
		  <%
				rs.MoveNext()
				End If 
				Next
		  %>
        </table>
        		
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
                    <p class="textoMedio">
                      <em>N�o h� aprova��es pendentes!</em>	</p>
                
                    <p>&nbsp;</p>
                    <p>	   
                     </p>
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