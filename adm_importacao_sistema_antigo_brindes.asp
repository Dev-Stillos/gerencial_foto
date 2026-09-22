<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<%
Dim redirecionar_ok
Dim objConexao3
redirecionar_ok = formatNumber(0&Request.QueryString("red"),0)

'Sub Rotina que Abre a Conexao

Sub AbrirConexao3()
Set objConexao3 = Server.CreateObject("ADODB.Connection")
objConexao3.Open = "Driver={MySQL ODBC 3.51 Driver}; SERVER=200.234.202.106; DATABASE=stillos2; UID=stillos2; PWD=fca123fca123@;"
End Sub


'Sub Rotina que Fecha a Conexao
Sub FecharConexao3()
   objConexao3.Close
   Set objConexao3=nothing
End Sub

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
	<!--AJUDA F1 E CTRL PARA BUSCA RÁPIDA-->
		
	function abrir(theURL,winName,features) 
	{ 
	  window.open(theURL,winName,features);
	}

		
	function abre(url)
	{    
				//if(event.srcElement.tagName=='TD')
				//	{
				//		src.children.tags('A')0.click(); 
				//	}
	location.href = url;

	}	
		
		
	var panelWidth = 113;	// Width of help panel	
	var slideSpeed = 15;		// Higher = quicker slide
	var slideTimer = 10;	// Lower = quicker slide
	var slideActive = true;	// Slide active ?
	var initBodyMargin = 0;	// Left or top margin of your <body> tag (left if panel is at the left, top if panel is on the top)
	var pushMainContentOnSlide = false;	// Push your main content to the right when sliding
	var panelPosition = 1; 	// 0 = left , 1 = top
	
	/*	Don't change these values */
	var slideLeftPanelObj=false;
	var slideInProgress = false;	
	var startScrollPos = false;
	var panelVisible = false;
	function initSlideLeftPanel(expandOnly)
	{
		if(slideInProgress)return;
		if(!slideLeftPanelObj){
			if(document.getElementById('dhtmlgoodies_leftPanel')){	// Object exists in HTML code?
				slideLeftPanelObj = document.getElementById('dhtmlgoodies_leftPanel');
				if(panelPosition == 1)slideLeftPanelObj.style.width = '100%';
			}else{	// Object doesn't exist -> Create <div> dynamically
				slideLeftPanelObj = document.createElement('DIV');
				slideLeftPanelObj.id = 'dhtmlgoodies_leftPanel';
				slideLeftPanelObj.style.display='none';
				document.body.appendChild(slideLeftPanelObj);
			}
			
			if(panelPosition == 1){
				slideLeftPanelObj.style.top = "-" + panelWidth + 'px';
				slideLeftPanelObj.style.left = '0px';	
				slideLeftPanelObj.style.height = panelWidth + 'px';			
			}else{
				slideLeftPanelObj.style.left = "-" + panelWidth + 'px';
				slideLeftPanelObj.style.top = '0px';
				slideLeftPanelObj.style.width = panelWidth + 'px';
			}
			

			if(!document.all || navigator.userAgent.indexOf('Opera')>=0)slideLeftPanelObj.style.position = 'fixed';;
		}	
		
		if(panelPosition == 0){
			if(document.documentElement.clientHeight){
				slideLeftPanelObj.style.height = document.documentElement.clientHeight + 'px';
			}else if(document.body.clientHeight){
				slideLeftPanelObj.style.height = document.body.clientHeight + 'px';
			}
			var leftPos = slideLeftPanelObj.style.left.replace(/^0-9\-/g,'')/1;
		}else{
			if(document.documentElement.clientWidth){
				slideLeftPanelObj.style.width = document.documentElement.clientWidth + 'px';
			}else if(document.body.clientHeight){
				slideLeftPanelObj.style.width = document.body.clientWidth + 'px';
			}
			var leftPos = slideLeftPanelObj.style.top.replace(/^0-9\-/g,'')/1;			
			
			
		}
		slideLeftPanelObj.style.display='block';
		
		if(panelPosition==1)
			startScrollPos = Math.max(document.body.scrollTop,document.documentElement.scrollTop);
		else
			startScrollPos = Math.max(document.body.scrollLeft,document.documentElement.scrollLeft);
		if(leftPos<(0+startScrollPos)){
			if(slideActive){
				slideLeftPanel(slideSpeed);	
			
			}else{
				document.body.style.marginLeft = panelWidth + 'px';
				slideLeftPanelObj.style.left = '0px';
			}
		}else{
			if(expandOnly)return;
			if(slideActive){		
				slideLeftPanel(slideSpeed*-1);
			}else{
				if(panelPosition == 0){
					if(pushMainContentOnSlide)document.body.style.marginLeft =  initBodyMargin + 'px';
					slideLeftPanelObj.style.left = (panelWidth*-1) + 'px';	
				}else{
					if(pushMainContentOnSlide)document.body.style.marginTop =  initBodyMargin + 'px';
					slideLeftPanelObj.style.top = (panelWidth*-1) + 'px';						
				}			
			}
		}	
		
		if(navigator.userAgent.indexOf('MSIE')>=0 && navigator.userAgent.indexOf('Opera')<0){
			window.onscroll = repositionHelpDiv;
		
			repositionHelpDiv();
		}
		window.onresize = resizeLeftPanel;
		
	}
	
	function resizeLeftPanel()
	{
		if(panelPosition == 0){
			if(document.documentElement.clientHeight){
				slideLeftPanelObj.style.height = document.documentElement.clientHeight + 'px';
			}else if(document.body.clientHeight){
				slideLeftPanelObj.style.height = document.body.clientHeight + 'px';
			}		
		}else{
			if(document.documentElement.clientWidth){
				slideLeftPanelObj.style.width = document.documentElement.clientWidth + 'px';
			}else if(document.body.clientWidth){
				slideLeftPanelObj.style.width = document.body.clientWidth + 'px';
			}	
		}
	}
	
	function slideLeftPanel(slideSpeed){
		slideInProgress =true;
		var scrollValue = 0;
		if(panelPosition==1)
			var leftPos = slideLeftPanelObj.style.top.replace(/^0-9\-/g,'')/1;
		else
			var leftPos = slideLeftPanelObj.style.left.replace(/^0-9\-/g,'')/1;
			
		leftPos+=slideSpeed;
		okToSlide = true;
		if(slideSpeed<0){
			if(leftPos < ((panelWidth*-1) + startScrollPos)){
				leftPos = (panelWidth*-1) + startScrollPos;	
				okToSlide=false;
			}
		}
		if(slideSpeed>0){
			if(leftPos > (0 + startScrollPos)){
				leftPos = 0 + startScrollPos;
				okToSlide = false;
			}			
		}
		
		
		if(panelPosition==0){
			slideLeftPanelObj.style.left = leftPos + startScrollPos + 'px';
			if(pushMainContentOnSlide)document.body.style.marginLeft = leftPos - startScrollPos + panelWidth + 'px';
		}else{
			slideLeftPanelObj.style.top = leftPos + 'px';
			if(pushMainContentOnSlide)document.body.style.marginTop = leftPos - startScrollPos + panelWidth + 'px';			
			
		}
		if(okToSlide)setTimeout('slideLeftPanel(' + slideSpeed + ')',slideTimer); else {
			slideInProgress = false;
			if(slideSpeed>0)panelVisible=true; else panelVisible = false;
		}
		
	}
	
	
	function repositionHelpDiv()
	{
		if(panelPosition==0){
			var maxValue = Math.max(document.body.scrollTop,document.documentElement.scrollTop);
			slideLeftPanelObj.style.top = maxValue;
		}else{
			var maxValue = Math.max(document.body.scrollLeft,document.documentElement.scrollLeft);
			slideLeftPanelObj.style.left = maxValue;	
			var maxTop = Math.max(document.body.scrollTop,document.documentElement.scrollTop);
			if(!slideInProgress)slideLeftPanelObj.style.top = (maxTop - (panelVisible?0:panelWidth)) + 'px'; 		
		}
	}
	
	function cancelEvent()
	{
		return false;
	}
	function keyboardShowLeftPanel()
	{
			initSlideLeftPanel();
			return false;	
	
	}
	
	function leftPanelKeyboardEvent(e)
	{
		if(document.all)return;
		
		if(e.keyCode==112){
			initSlideLeftPanel();
			return false;
		}
	
		
			
	}
	
	function setLeftPanelContent(text)
	{
		document.getElementById('leftPanelContent').innerHTML = text;
		initSlideLeftPanel(true);
		
	}
	if(!document.all)document.documentElement.onkeypress = leftPanelKeyboardEvent;
	document.documentElement.onhelp  = keyboardShowLeftPanel;

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
  BTN_IMPRIMIR_LINK = "#"
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">

	  <h2>ADM - Importação Sistema Antigo</h2>
	   <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	  
          <%
		  AbrirConexao()
		  AbrirConexao3()
		  
		  '1) Seleciona todos os pedidos que já foram finalizados pela Ligia e ainda não foram exportados
		  		'Para os lançamentos antigos deverão ser primeiro verificados no sistema antigo, para depois serem autenticados no novo
				'O sistema antigo só permitira a correção não será possível baixar mais do sistema antigo, para baixar para diretoria deverá ser do novo
		  '2) 
		  
		  Dim rs, rs2, rs3, rs4, rs5, sql
		
															  
		  '///////////////////////////////////////////////////////////////////////////////////////////////////
		  '/////////////////////
		 	'	--VERIFICA SE HÁ PRODUTOS COMO "OUTROS"
			'	SELECT 
			'	TB_PEDIDOS.COD_PEDIDO,
			'	TB_PEDIDOS.SISTEMA_NOVO,
			'	TB_PEDIDOS_PRODUTOS.COD_PRODUTO 
			'	FROM TB_PEDIDOS INNER JOIN TB_PEDIDOS_PRODUTOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO
			'	WHERE 
			'	TB_PEDIDOS.SISTEMA_NOVO = 2 AND (TB_PEDIDOS_PRODUTOS.COD_PRODUTO = 12)
		  
		  '/////////
		  	'Brindes pendêntes
			'"SELECT COUNT(*) AS QTD_PENDENTES FROM BRINDES_PENDENTES"
		  '/////////
		  
		  'IMPORTAÇÃO DE BRINDES!!!!!!!!!
		  '1) SISTEMA NOVO: Padronizar as FOTOS dos pedidos antigos com código único "2", para que o brinde possa ser contado corretamente.
		  '2) SISTEMA ANTIGO: Verifica se há pedidos com brindes não importados.
		  '3) SISTEMA ANTIGO: Destes pedidos com brindes e não exportados verifica qual o brinde e dependendo do brinde exportado somente o brinde para o novo
		  '4) SISTEMA ANTIGO: Sempre padronizar as FOTOS dos pedidos antigos como código "2"
		  '5) SISTEMA ANTIGO: Faz update no antigo informando que o brinde foi importado
		  '6) SISTEMA ANTIGO: Contabiliza os brindes importados e os brindes não importados
		  '7) SISTEMA ANTIGO: Importar manualmente brindes com pedidos com mais de 1 formando
		 'Somente pedidos individuais por enquanto

		 Dim VALIDACAO_DA_VEZ, ROTINA_QTD_BRINDES_FOTO
		 VALIDACAO_DA_VEZ = "01 ÁLBUM COM 30 FOTOS + DVD"
		 ROTINA_QTD_BRINDES_FOTO = FormatNumber(0&Request.QueryString("a"),0)
		 
		  
		' sql = "SELECT * FROM TB_FOTO_VENDA WHERE (BRINDES2 = '' or BRINDES2 = '0') AND BRINDE_EXPORTADO = 2 AND EXPORTADO_SISTEMA_NOVO = 1 AND QTD_FORMANDOS = 0 AND BRINDES LIKE '%"&VALIDACAO_DA_VEZ&"%'"
		
		sql  = "SELECT * FROM TB_FOTO_VENDA WHERE COD_PEDIDO = "&Request.QueryString("p")&" LIMIT 1"
		  
		'  	sql = "SELECT * FROM TB_FOTO_VENDA WHERE (BRINDES2 = '' or BRINDES2 = '0') AND BRINDE_EXPORTADO = 2 AND EXPORTADO_SISTEMA_NOVO = 1 AND QTD_FORMANDOS = 0 AND BRINDES = '"&VALIDACAO_DA_VEZ&"'"
		  
		  Response.Write("<br>"&sql)
		  Set rs=objConexao3.Execute(sql,qtdRegistros)
		  
		  Response.Write("<p align='center'>"&qtdRegistros&"</p>")
		  Dim BRINDES_IMPORTADOS, BRINDES_NAO_IMPORTADOS, NAO_CONTINUA
		  BRINDES_IMPORTADOS = 0
		  BRINDES_NAO_IMPORTADOS = 0
		  NAO_CONTINUA = 2
		  	
				While not rs.Eof 
		  								
					Response.Write("INÍCIO ------////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////")
					Response.Write("<br />")
									
						Dim ROTINA_COD_PEDIDO, ROTINA_BRINDES
						Dim ROTINA_COD_FORMANDO
						Dim ROTINA_INICIAL, ROTINA_FATURADO, ROTINA_RETIRADO
						Dim ROTINA_TOTAL_BASE_1
						ROTINA_TOTAL_BASE_1 = 0
						
								
						ROTINA_COD_PEDIDO	= rs("COD_PEDIDO")
						ROTINA_BRINDES		= rs("BRINDES")
						'ROTINA_BRINDES		= Replace(ROTINA_BRINDES,".","")
						
						') Acha o código do formando
						Set rs2 = objConexao.Execute("SELECT * FROM TB_PEDIDOS_FORMANDOS WHERE COD_PEDIDO="&ROTINA_COD_PEDIDO&" LIMIT 1")
								
								If not rs2.Eof Then
										
									ROTINA_COD_FORMANDO = rs2("COD_FORMANDO")
									
									Set rs3=objConexao.Execute("SELECT INICIAL,FATURADO,RETIRADO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO="&ROTINA_COD_FORMANDO&" LIMIT 1")
										
										If not rs3.Eof Then
												
										ROTINA_INICIAL		= formatNumber(0&rs3("INICIAL"),0)/1
										ROTINA_FATURADO		= formatNumber(0&rs3("FATURADO"),0)/1
										ROTINA_RETIRADO		= formatNumber(0&rs3("RETIRADO"),0)/1
										Response.Write("Inicial: "&ROTINA_INICIAL)
										Response.Write("<br />")
										Response.Write("Faturado: "&ROTINA_FATURADO)
										Response.Write("<br />")
										Response.Write("Retirado: "&ROTINA_RETIRADO)
										Else
										
										NAO_CONTINUA = 1
										
										End If
										
									rs3.Close
									Set rs3 = Nothing
									

									
								Else
								
									Response.Write("ERRO CÓD DO FORMANDO NÃO ENCONTRADO PEDIDO: "&ROTINA_COD_PEDIDO)
									Response.Write("<br />")
									
								End If
								
								rs2.Close
								Set rs2 = Nothing				 
						
						IF  NAO_CONTINUA = 2 Then
						
							'//////////////////////////////////////////////////////////////////////////////
												
						'If ROTINA_BRINDES = VALIDACAO_DA_VEZ Then
						IF 1 = 1 Then
						'pronta	
						
						'Antes de inserir verificase:
						'1) FATURADO + RETIRADO + 35 = INICIAL 
								
								ROTINA_TOTAL_BASE_1 = (ROTINA_FATURADO+ROTINA_RETIRADO+ROTINA_QTD_BRINDES_FOTO)/1
								Response.Write("<br />----------------")
								Response.Write("Fotos: "&ROTINA_TOTAL_BASE_1&" Inicial: "&ROTINA_INICIAL)
								Response.Write("<br /><br />")
								If Int(ROTINA_TOTAL_BASE_1) = Int(ROTINA_INICIAL) Then
										
										
																			'INSERE OS PRODUTOS:
											'35 FOTOS
											'1 DVD
											
											sql = 	"INSERT INTO TB_FICHAS_IDENTIFICACAO_BRINDES ("&_
													"COD_FORMANDO ,"&_
													"COD_PRODUTO ,"&_
													"QUANTIDADE ,"&_
													"CADASTRADO_POR ,"&_
													"VENDIDO "&_
													")"&_
													"VALUES ("&_
													ROTINA_COD_FORMANDO&", '2', "&ROTINA_QTD_BRINDES_FOTO&", '9999', '1')"


											objConexao.Execute(sql)
											'Response.Write(sql)
											
											sql = 	"INSERT INTO TB_FICHAS_IDENTIFICACAO_BRINDES ("&_
													"COD_FORMANDO ,"&_
													"COD_PRODUTO ,"&_
													"QUANTIDADE ,"&_
													"CADASTRADO_POR ,"&_
													"VENDIDO "&_
													")"&_
													"VALUES ("&_
													ROTINA_COD_FORMANDO&", '1', '1', '9999', '1')"
												
											'objConexao.Execute(sql)		
																						
											
											
											
											objConexao3.Execute("UPDATE TB_FOTO_VENDA SET BRINDE_EXPORTADO = '1', DATA_BRINDE_EXPORTADO = NOW() WHERE COD_PEDIDO ="&ROTINA_COD_PEDIDO&" LIMIT 1")
											
											BRINDES_IMPORTADOS = BRINDES_IMPORTADOS + 1
											Response.Write(ROTINA_COD_PEDIDO&" IMPORTADO!")
											Response.Write("<br />")
								
								
						'/////////////////////////

						Else
								
									BRINDES_NAO_IMPORTADOS = BRINDES_NAO_IMPORTADOS + 1
									Response.Write("<br />************")
									Response.Write("Cód Pedido: "&ROTINA_COD_PEDIDO&" Brinde: "&ROTINA_BRINDES&" NÃO IMPORTADO QTD. INICIAL NÃO BATE!!!")	
						End If








											
						Else
	
								BRINDES_NAO_IMPORTADOS = BRINDES_NAO_IMPORTADOS + 1
								Response.Write("Cód Pedido: "&ROTINA_COD_PEDIDO&" Brinde: "&ROTINA_BRINDES&" NÃO IMPORTADO!!!")
								Response.Write("<br />")
						End If
						
						Elseif NAO_CONTINUA = 1 Then
						
							Response.Write("PROBLEMAAAAAAAAAAAA CÓD PEDIDO: "&ROTINA_COD_PEDIDO&" com código do formando "&ROTINA_COD_FORMANDO&"!")
										
						End If
					Response.Write("<br />")	
					Response.Write("FIM ------////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////")
					Response.Write("<br />")	
					Response.Write("<br />")	
					Response.Write("<br />")
							
				rs.MoveNext()
				Wend
				
		  Set rs = Nothing
		  %>
		  <br />
<br />
<br />
<br />
<br />

		  Total Importados: <%=BRINDES_IMPORTADOS%><br />
  		  Total Não Importados: <%=BRINDES_NAO_IMPORTADOS%><br />

		  
		  
		  <%
		   
		  
		  
		  
		  fecharConexao3()
		  fecharConexao()
		  
		  If redirecionar_ok = 1 Then
		  	Response.Redirect("lca_pedidos_pendentes.asp?importado")
		  End If
		  %>
          
          <p align="center"> </p>
		  <p align="center"> </p>		
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



