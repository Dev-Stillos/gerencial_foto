<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim redirecionar_ok
Dim objConexao3
redirecionar_ok = formatNumber(0&Request.QueryString("red"),0)

'Sub Rotina que Abre a Conexao

Sub AbrirConexao3()
Set objConexao3 = Server.CreateObject("ADODB.Connection")
objConexao3.Open("DRIVER={MySQL ODBC 5.1 Driver};SERVER=200.234.202.106;PORT=3306;DATABASE=stillos2;USER=stillos2;PASSWORD=fca123fca123@;")
'Conexão Antiga	
'objConexao3.Open = "Driver={MySQL ODBC 3.51 Driver}; SERVER=200.234.202.106; DATABASE=stillos2; UID=stillos2; PWD=fca123fca123@;"
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

	  <h2>ADM - Importa&ccedil;&atilde;o Sistema Antigo</h2>
	   <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	  
          <%
		  AbrirConexao()
		  AbrirConexao3()
		  
		  '1) Seleciona todos os pedidos que já foram finalizados pela Ligia e ainda não foram exportados
		  		'Para os lançamentos antigos deverão ser primeiro verificados no sistema antigo, para depois serem autenticados no novo
				'O sistema antigo só permitira a correção não será possível baixar mais do sistema antigo, para baixar para diretoria deverá ser do novo
		  '2) 
		  
		  Dim rs, rs2, rs3, rs4, rs5, sql
		  Dim qtdParaImportacao
		  Dim msg
		  Dim ERRO_NAO_PODERA_IMPORTAR, CONTA_NAO_IMPORTADOS, CONTA_IMPORTADOS, msg_erro_descricao
		  Dim COD_FORMANDO_1_VALIDADO, COD_FORMANDO_2_VALIDADO, COD_FORMANDO_3_VALIDADO, COD_FORMANDO_4_VALIDADO
		  
		  Set rs=objConexao3.Execute("SELECT * FROM TB_FOTO_VENDA WHERE FINALIZA_CORRECAO=1 AND EXPORTADO_SISTEMA_NOVO=2 LIMIT 1000",qtdParaImportacao)
		  
		  	If not rs.Eof Then
			
				
				CONTA_NAO_IMPORTADOS		= 0
				CONTA_IMPORTADOS			= 0
				
				While not rs.Eof
					ERRO_NAO_PODERA_IMPORTAR 	= 2
					'/////////
					'1 = Aguardando verificação
					'2 = Verificado existe
					'3 = Verificado não encontrado Access e nem no sistema novo como formando cadastro individual
					'4 = Verificado não encontrado Access, mas encontrado no sistema novo como formando cadastro individual pelo Depto. de Foto
					
					COD_FORMANDO_1_VALIDADO		= 1
					COD_FORMANDO_2_VALIDADO		= 1
					COD_FORMANDO_3_VALIDADO		= 1
					COD_FORMANDO_4_VALIDADO		= 1
					'/////////
					
					Dim COD_PEDIDO, NUMERO_PEDIDO, CODIGO_VENDEDOR, ENTIDADE_DE_ENSINO, CURSO, CURSO2, CURSO3, CURSO4, DATA
					Dim COD_CURSO, COD_CURSO2, COD_CURSO3, COD_CURSO4, ANO, SEMESTRE, ALBUM, ALBUM2, ALBUM3, ALBUM4, EM_DINHEIRO
					Dim INICIAL, INICIAL2, INICIAL3, INICIAL4, FATURADO, FATURADO2, FATURADO3, FATURADO4, RETIRADO, RETIRADO2, RETIRADO3, RETIRADO4
					Dim BRINDES, BRINDES2, BRINDES3, BRINDES4, FINALIZA_VENDEDOR, FINALIZA_CORRECAO, FINALIZA_DEP_FOTO, FINALIZA_DIRETORIA
					Dim OBSERVACAO, COD_FORMANDO, COD_FORMANDO2, COD_FORMANDO3, COD_FORMANDO4, COD_CONTRATO, QTD_FORMANDOS, VALOR_AJUDA_CUSTO
					Dim REPASSE, COD_MES_BAIXA, COD_ANO_BAIXA, DATA_AUT, DATA_DESC_DIR, ERRO_LISTA, DATA_BAIXADO_DIRETORIA, HORA_BAIXADO_DIRETORIA
					
					'//
					'Tb 3 Mysql 3
					'PARTE1
					Dim rs3_COD_FORMANDO, rs3_COD_CONTRATO, rs3_COD_CURSO, rs3_NOME_CURSO, rs3_NOME_FORMANDO, rs3_ANO_CONCLUSAO, rs3_SEMESTRE_CONCLUSAO
					'PARTE2
					Dim rs3_ENDERECO, rs3_CEP, rs3_FONE_RES, rs3_FONE_CEL, rs3_BAIRRO, rs3_CIDADE, rs3_ESTADO
					
										
					COD_PEDIDO					= rs("COD_PEDIDO")
                    NUMERO_PEDIDO				= rs("NUMERO_PEDIDO")
                    CODIGO_VENDEDOR				= rs("CODIGO_VENDEDOR")
					DATA						= rs("DATA")
                    ENTIDADE_DE_ENSINO			= rs("ENTIDADE_DE_ENSINO")
                    CURSO						= rs("CURSO")
                    CURSO2						= rs("CURSO2")
                    CURSO3						= rs("CURSO3")
                    CURSO4						= rs("CURSO4")
					COD_CURSO					= rs("COD_CURSO")
                    COD_CURSO2					= rs("COD_CURSO2")
                    COD_CURSO3					= rs("COD_CURSO3")
                    COD_CURSO4					= rs("COD_CURSO4")
                    ANO							= rs("ANO")
                    SEMESTRE					= rs("SEMESTRE")
                    ALBUM						= rs("ALBUM")
                    ALBUM2						= rs("ALBUM2")
                    ALBUM3						= rs("ALBUM3")
                    ALBUM4						= rs("ALBUM4")
                    EM_DINHEIRO					= rs("EM_DINHEIRO")
					INICIAL						= rs("INICIAL")
                    INICIAL2					= rs("INICIAL2")
                    INICIAL3					= rs("INICIAL3")
                    INICIAL4					= rs("INICIAL4")
                    FATURADO					= rs("FATURADO")
                    FATURADO2					= rs("FATURADO2")
                    FATURADO3					= rs("FATURADO3")
                    FATURADO4					= rs("FATURADO4")
                    RETIRADO					= rs("RETIRADO")
                    RETIRADO2					= rs("RETIRADO2")
                    RETIRADO3					= rs("RETIRADO3")
                    RETIRADO4					= rs("RETIRADO4")
					BRINDES						= rs("BRINDES")
                    BRINDES2					= rs("BRINDES2")
                    BRINDES3					= rs("BRINDES3")
                    BRINDES4					= rs("BRINDES4")
                    FINALIZA_VENDEDOR			= rs("FINALIZA_VENDEDOR")
                    FINALIZA_CORRECAO			= rs("FINALIZA_CORRECAO")
                    FINALIZA_DEP_FOTO			= rs("FINALIZA_DEP_FOTO")
                    FINALIZA_DIRETORIA			= rs("FINALIZA_DIRETORIA")
					OBSERVACAO					= rs("OBSERVACAO")
                    COD_FORMANDO				= rs("COD_FORMANDO")
                    COD_FORMANDO2				= rs("COD_FORMANDO2")
                    COD_FORMANDO3				= rs("COD_FORMANDO3")
                    COD_FORMANDO4				= rs("COD_FORMANDO4")
                    COD_CONTRATO				= rs("COD_CONTRATO")
                    QTD_FORMANDOS				= rs("QTD_FORMANDOS")
                    VALOR_AJUDA_CUSTO			= rs("VALOR_AJUDA_CUSTO")
					REPASSE						= rs("REPASSE")
                    COD_MES_BAIXA				= rs("COD_MES_BAIXA")
                    COD_ANO_BAIXA				= rs("COD_ANO_BAIXA")
                    DATA_AUT					= rs("DATA_AUT")
                    DATA_DESC_DIR				= rs("DATA_DESC_DIR")
                    ERRO_LISTA					= rs("ERRO_LISTA")
                    DATA_BAIXADO_DIRETORIA		= rs("DATA_BAIXADO_DIRETORIA")
                    HORA_BAIXADO_DIRETORIA		= rs("HORA_BAIXADO_DIRETORIA")
					
					'///////////////////////////////////////////////////////
					'Com todos os dados da tabela principal verifica os códigos dos formandos com a tabela no Access se não encontrar de acordo com o número
					'do código do formando para importação desta linha e passa para próxima
					sql = "SELECT * FROM TB_FORMANDOS_IMP  WHERE "
					
					
					If Len(COD_FORMANDO) > 2 Then
							sql = sql & "CODIGO_FORMANDO= "&COD_FORMANDO
							Set rs2 = objConexao.execute(sql)
							
								If not rs2.Eof Then
										
										'Verifica se já existe o formando
										sql = "SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO="&COD_FORMANDO&"0"
										'Response.Write(sql&"<br />")
										Set rs5 = objConexao.Execute(sql)
												
												If not rs5.Eof Then
													msg_erro_descricao = "Código do formando DUPLICADO!"
													ERRO_NAO_PODERA_IMPORTAR = 1
												Else
													COD_FORMANDO_1_VALIDADO = 2
												End if
												
										rs5.Close
										Set rs5 = Nothing
										
								Else
										
									'Verifica se foi um formando importado do sistema antigo	
									'//////////////////////////////////////////
									Dim COD_FORMANDO_ANTIGO_1, IMPORTACAO_DIRETA_PEDIDO_1, METODO_B
									COD_FORMANDO_ANTIGO_1 	= COD_FORMANDO
									METODO_B				= 2
									COD_FORMANDO_ANTIGO_1 	= Replace(COD_FORMANDO_ANTIGO_1,"0000","")
									
									'1) Verifica se este formando consta no novo
									Dim rs7789, rs7790
									
									Set rs7789 = objConexao.Execute("SELECT COD_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO="&COD_FORMANDO_ANTIGO_1)
									
										'2) Caso existir verifica se não tem pedido relacionado
										If not rs7789.Eof Then
												
												Set rs7790 = objConexao.Execute("SELECT COD_PEDIDO FROM TB_PEDIDOS_FORMANDOS WHERE COD_FORMANDO ="&COD_FORMANDO_ANTIGO_1)
													If not rs7790.Eof Then
													'3) Caso já existir bloqueia e não deixa importar
														msg_erro_descricao = "Pedido já migrado! "&COD_FORMANDO_ANTIGO_1
														ERRO_NAO_PODERA_IMPORTAR = 1
													Else
													'-4) Grava o FORMANDO_PEDIDO
															METODO_B = 1
															sql="INSERT INTO TB_PEDIDOS_FORMANDOS ("&_
																"COD_PEDIDO ,"&_
																"COD_FORMANDO)"&_ 
																"VALUES("&COD_PEDIDO&", "&COD_FORMANDO_ANTIGO_1&")"
															
																objConexao.Execute(sql)
													
													End If 
												
												
												rs7790.Close
												Set rs7790 = Nothing
																	
										
										Else
										'Código não encontrado
												msg_erro_descricao = "NÃO ENCONTRADO NO SISTEMA! "&COD_FORMANDO_ANTIGO_1
												ERRO_NAO_PODERA_IMPORTAR = 1
										
										End IF
									
									
									rs7789.Close
									Set rs7789 = Nothing
																		
																									

									'//////////////////////////////////////////
								
								End If
								
							rs2.Close
							Set rs2 = Nothing
					Else
					
						
						msg_erro_descricao = "Código do formando não encontrado"
						ERRO_NAO_PODERA_IMPORTAR = 1
					
					
					End If
					
					If Len(COD_FORMANDO2) > 2 Then
							sql = "SELECT * FROM TB_FORMANDOS_IMP  WHERE "
							sql = sql & "CODIGO_FORMANDO= "&COD_FORMANDO
							Set rs2 = objConexao.execute(sql)
							
								If not rs2.Eof Then
									'Verifica se já existe o formando
										Set rs5 = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO="&COD_FORMANDO2&"0")
												
												If not rs5.Eof Then
													msg_erro_descricao = "Código do formando DUPLICADO!"
													ERRO_NAO_PODERA_IMPORTAR = 1
												Else
													
													COD_FORMANDO_2_VALIDADO = 2
												End if
												
										rs5.Close
										Set rs5 = Nothing
								Else
									'Verifica se foi um formando importado do sistema antigo	
																		'//////////////////////////////////////////
																		Dim COD_FORMANDO_ANTIGO_2, IMPORTACAO_DIRETA_PEDIDO_2
																		COD_FORMANDO_ANTIGO_2 = COD_FORMANDO
																		COD_FORMANDO_ANTIGO_2 = Replace(COD_FORMANDO_ANTIGO_2,"0000","")
																		
																		'1) Verifica se este formando consta no novo
																		
																		
																		Set rs7789 = objConexao.Execute("SELECT COD_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO="&COD_FORMANDO_ANTIGO_1)
																		
																			'2) Caso existir verifica se não tem pedido relacionado
																			If not rs7789.Eof Then
																					
																					Set rs7790 = objConexao.Execute("SELECT COD_PEDIDO FROM TB_PEDIDOS_FORMANDOS WHERE COD_FORMANDO ="&COD_FORMANDO_ANTIGO_1)
																						If not rs7790.Eof Then
																						'3) Caso já existir bloqueia e não deixa importar
																							msg_erro_descricao = "Pedido já migrado! "&COD_FORMANDO_ANTIGO_2
																							ERRO_NAO_PODERA_IMPORTAR = 1
																						Else
																						'-4) Grava o FORMANDO_PEDIDO
																						
																								sql="INSERT INTO TB_PEDIDOS_FORMANDOS ("&_
																									"COD_PEDIDO ,"&_
																									"COD_FORMANDO)"&_ 
																									"VALUES("&COD_PEDIDO&", "&COD_FORMANDO_ANTIGO_2&")"
																								
																									objConexao.Execute(sql)
																						
																						End If 
																					
																					
																					rs7790.Close
																					Set rs7790 = Nothing
																										
																			
																			Else
																			'Código não encontrado
																					msg_erro_descricao = "NÃO ENCONTRADO NO SISTEMA! "&COD_FORMANDO_ANTIGO_2
																					ERRO_NAO_PODERA_IMPORTAR = 1
																			
																			End IF
																		
																		
																		rs7789.Close
																		Set rs7789 = Nothing
																											
																																		
									
																		'//////////////////////////////////////////
								End If
								
							rs2.Close
							Set rs2 = Nothing
					End If
					
					
					'//////////////////////////////////////////////////////////////////// 
					'REPETIR ESTE PROCESSO PARA TODOS OS CÓDIGOS FORMANDOS 2,3 e 4 ABAIXO. 
					If Len(COD_FORMANDO3) > 2 Then
							sql = "SELECT * FROM TB_FORMANDOS_IMP  WHERE "
							sql = sql & "CODIGO_FORMANDO= "&COD_FORMANDO
							Set rs2 = objConexao.execute(sql)
							
								If not rs2.Eof Then
									'Verifica se já existe o formando
										Set rs5 = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO="&COD_FORMANDO3&"0")
												
												If not rs5.Eof Then
													msg_erro_descricao = "Código do formando DUPLICADO!"
													ERRO_NAO_PODERA_IMPORTAR = 1
												Else
													COD_FORMANDO_3_VALIDADO = 2
												End if
												
										rs5.Close
										Set rs5 = Nothing
								Else
									msg_erro_descricao = "Código do formando não encontrado"
									ERRO_NAO_PODERA_IMPORTAR = 1
								End If
								
							rs2.Close
							Set rs2 = Nothing
					End If
					
					If Len(COD_FORMANDO4) > 2 Then
							sql = "SELECT * FROM TB_FORMANDOS_IMP  WHERE "
							sql = sql & "CODIGO_FORMANDO= "&COD_FORMANDO
							Set rs2 = objConexao.execute(sql)
							
								If not rs2.Eof Then
									'Verifica se já existe o formando
										Set rs5 = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO="&COD_FORMANDO4&"0")
												
												If not rs5.Eof Then
													msg_erro_descricao = "Código do formando DUPLICADO!"
													ERRO_NAO_PODERA_IMPORTAR = 1
												Else
													COD_FORMANDO_4_VALIDADO = 2
												End if
												
										rs5.Close
										Set rs5 = Nothing
								Else
									msg_erro_descricao = "Código do formando não encontrado"
									ERRO_NAO_PODERA_IMPORTAR = 1
								End If
								
							rs2.Close
							Set rs2 = Nothing
					End If
					
					
					
					
					
					If ERRO_NAO_PODERA_IMPORTAR =1 Then
						CONTA_NAO_IMPORTADOS = CONTA_NAO_IMPORTADOS + 1
						Response.Write("<br />Pedido do sistema antigo número:"&rs("cod_pedido")&" não poderá ser importado, Erro:"&msg_erro_descricao&" !")
					Else
						CONTA_IMPORTADOS = CONTA_IMPORTADOS + 1
						'Avança somente se não tive nenhum erro
					
					
									'ATENÇÃO COMEÇAR IMPORTAÇÃO SOMENTE APÓS VALIDAÇÃO DE COD_FORMANDO_1_VALIDADO, 2, 3 e 4
									'PARA NÃO CORRER O RISCO DE UM FORMANDO EXISTIR E OUTRO NÃO E ACABAR IMPORTADO MESMO ASSIM
									'/////////////////////////////////////////////////////////////////
									'Verificado é possível importar!
									'1) Capta os dados na tb Access
									'2) Faz a validação de importação (limpa nomes etc. e coloca digito 0)
									'3) Insere na tabela TB_FICHAS_IDENTIFICACAO com os dados do Access
									'4) Insere o pedido na TB_PEDIDOS e seus respectivos formandos na TB_PEDIDOS_FORMANDOS
									'5) Verifica as formas de pagamento e insere cada uma na respectivas tabelas de Cartão e Cheque
									'6) Os produtos que estiverem lançados como "Outros" serão lançados com o código do produto 12
									'	e posteriormente deverão ser corrigidos
									'7) Após inserir a FICHA, PEDIDO, VALIDAÇÃO D.ANDREA, ou seja todos os campos importados
										'volta na tb antiga e faz update com a data da importação e EXPORTADO = 1
										'atenção fazer isto somente após os 4 códigos de formandos
										
										
										
									'Observações:	
									'- Futuramente deverá verificar quais eram de comissão de formatura destes pedidos e cadastrar seus respectivos brindes.
									'- Corrigir os códigos de formando que estiverem incorretos
									'- Os formandos que não tiverem código deverão ser cadastrados no sistema novo, com o número de formando novo serão lançados no
									   'sistema antigo e na importação o sistema verificará primeiro no Access caso não encontre tenta na FICHAS_IDENTIFICACAO tb nova
									   'desta forma importa o pedido da antiga
									  
									'//////////////////////////////////////////////////////////////////// 
									'REPETIR ESTE PROCESSO PARA TODOS OS CÓDIGOS FORMANDOS 1, 2,3 e 4 ABAIXO.
									'FICHAS DESMEMBRADAS!!!!!!!!
								
									COD_FORMANDO_1_VALIDADO 	= 2
														IMPORTACAO_DIRETA_PEDIDO_1 	= 1
									
									If COD_FORMANDO_1_VALIDADO = 2 Then
									'Significa que está correto pode começar o processo!
										sql = "SELECT * FROM TB_FORMANDOS_IMP  WHERE "
										sql = sql & "CODIGO_FORMANDO= "&COD_FORMANDO
										Set rs3 = objConexao.execute(sql)
										
											If not rs3.Eof Then						
													
													
													'Tb 1
													'0 devido a importação da formatura
													rs3_COD_FORMANDO		= rs3("CODIGO_FORMANDO")&"0"
													rs3_COD_CONTRATO		= rs3("CONTRATO")
													rs3_COD_CURSO			= rs3("CODIGO_CURSO")
													rs3_NOME_CURSO			= rs3("NOME_CURSO")
													rs3_NOME_FORMANDO		= ""&rs3("NOME_FORMANDO")
													'Correção do nome do formando
													posicaoInicial = inStr(rs3_NOME_FORMANDO,"(")
													If posicaoInicial > 0 Then
														rs3_NOME_FORMANDO = Trim(Left(rs3_NOME_FORMANDO,posicaoInicial-1))
													End If

													rs3_ANO_CONCLUSAO		= rs3("ANO")
													rs3_SEMESTRE_CONCLUSAO	= rs3("SEMESTRE")
													'------------------------------------
				
													'Tb2
													rs3_ENDERECO			= rs3("ENDERECO")
													rs3_CEP					= rs3("CEP")
													rs3_CEP					= Replace(rs3_CEP,"-","")
													rs3_FONE_RES			= rs3("FONE-RES")
													rs3_FONE_CEL			= rs3("FONE-COM-CEL")
													rs3_BAIRRO				= rs3("BAIRRO")
													rs3_CIDADE				= rs3("MUNICIPIO")
													rs3_ESTADO				= rs3("ESTADO")
													'------------------------------------
				
												
															sql = 	"INSERT INTO TB_FICHAS_IDENTIFICACAO ("&_
															"COD_FORMANDO , "&_
															"SISTEMA_NOVO , "&_
															"COD_PRODUCAO , "&_
															"COD_CONTRATO , "&_
															"COD_CURSO , "&_
															"NOME_CURSO , "&_
															"NOME_FORMANDO , "&_
															"ANO_CONCLUSAO , "&_
															"SEMESTRE_CONCLUSAO , "&_
															"ALBUM , "&_
															"INICIAL , "&_
															"FATURADO , "&_
															"RETIRADO , "&_
															"VENDIDO , "&_
															"REPASSE"&_
															")"&_
															"VALUES ("&_
															sqltexto(rs3_COD_FORMANDO)&", "&_
															"2, "&_
															"0, "&_
															sqlTexto(rs3_COD_CONTRATO)&", "&_
															rs3_COD_CURSO&", "&_
															sqltexto(rs3_NOME_CURSO)&", "&_
															sqltexto(rs3_NOME_FORMANDO)&", "&_
															sqltexto(rs3_ANO_CONCLUSAO)&", "&_
															sqltexto(Ucase(rs3_SEMESTRE_CONCLUSAO))&", "&_
															ALBUM&", "&_
															INICIAL&", "&_
															FATURADO&", "&_
															RETIRADO&", "&_
															FINALIZA_CORRECAO&", "&_
															REPASSE&_
															")"			
															'Tb 1
															objConexao.Execute(sql)
															Response.Write(sql&"<br /><br /><br /><br />")
															'-------------
								
															sql = 	"INSERT INTO TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO ("&_
																	"COD_FORMANDO ,"&_
																	"ENDERECO ,"&_
																	"CEP ,"&_
																	"FONE_RES ,"&_
																	"FONE_CEL ,"&_
																	"BAIRRO ,"&_
																	"CIDADE ,"&_
																	"ESTADO ,"&_
																	"CADASTRADO_POR "&_
																	")"&_
																	"VALUES ("&_
																	rs3_COD_FORMANDO&", "&sqlTexto(rs3_ENDERECO)&", '"&rs3_CEP&"', '"&rs3_FONE_RES&"', '"&rs3_FONE_CEL&"', "&sqlTexto(rs3_BAIRRO)&", "&sqlTexto(rs3_CIDADE)&", "&_
																	sqlTexto(rs3_ESTADO)&", "&Session("COD_USUARIO")&_
																	")"
																	
																	'Tb 2
																	Response.Write(sql&"<br /><br /><br /><br />")
																	objConexao.Execute(sql)
																	
																	'-------------																	
																	sql="INSERT INTO TB_PEDIDOS_FORMANDOS ("&_
																	"COD_PEDIDO ,"&_
																	"COD_FORMANDO)"&_ 
																	"VALUES("&COD_PEDIDO&", "&rs3_COD_FORMANDO&")"
																	objConexao.Execute(sql)
																	Response.Write(sql&"<br /><br /><br /><br />")
													
											End If
											
										rs3.Close
										Set rs3 = Nothing									
									End If
									
																	If COD_FORMANDO_2_VALIDADO = 2 Then
									'Significa que está correto pode começar o processo!
										sql = "SELECT * FROM TB_FORMANDOS_IMP  WHERE "
										sql = sql & "CODIGO_FORMANDO= "&COD_FORMANDO2
										Set rs3 = objConexao.execute(sql)
										
											If not rs3.Eof Then						
													
													
													'Tb 1
													'0 devido a importação da formatura
													rs3_COD_FORMANDO		= rs3("CODIGO_FORMANDO")&"0"
													rs3_COD_CONTRATO		= rs3("CONTRATO")
													rs3_COD_CURSO			= rs3("CODIGO_CURSO")
													rs3_NOME_CURSO			= rs3("NOME_CURSO")
													rs3_NOME_FORMANDO		= ""&rs3("NOME_FORMANDO")
													'Correção do nome do formando
													posicaoInicial = inStr(rs3_NOME_FORMANDO,"(")
													If posicaoInicial > 0 Then
														rs3_NOME_FORMANDO = Trim(Left(rs3_NOME_FORMANDO,posicaoInicial-1))
													End If

													rs3_ANO_CONCLUSAO		= rs3("ANO")
													rs3_SEMESTRE_CONCLUSAO	= rs3("SEMESTRE")
													'------------------------------------
				
													'Tb2
													rs3_ENDERECO			= Replace(rs3("ENDERECO"),"'","")
													rs3_CEP					= Replace(rs3("CEP"),"'","")
													rs3_CEP					= Replace(rs3_CEP,"-","")
													rs3_FONE_RES			= Replace(rs3("FONE-RES"),"'","")
													'rs3_FONE_CEL			= Replace(rs3("FONE-COM-CEL"),"'","")
													rs3_BAIRRO				= Replace(rs3("BAIRRO"),"'","")
													rs3_BAIRRO				= Replace(rs3_BAIRRO,"\","")
													rs3_CIDADE				= Replace(rs3("MUNICIPIO"),"'","")
													rs3_CIDADE				= Replace(rs3_CIDADE,"\","")
													rs3_ESTADO				= Replace(rs3("ESTADO"),"'","")
													rs3_ESTADO				= Replace(rs3_ESTADO,"\","")
													
													'------------------------------------
				
												
															sql = 	"INSERT INTO TB_FICHAS_IDENTIFICACAO ("&_
															"COD_FORMANDO , "&_
															"SISTEMA_NOVO , "&_
															"COD_PRODUCAO , "&_
															"COD_CONTRATO , "&_
															"COD_CURSO , "&_
															"NOME_CURSO , "&_
															"NOME_FORMANDO , "&_
															"ANO_CONCLUSAO , "&_
															"SEMESTRE_CONCLUSAO , "&_
															"ALBUM , "&_
															"INICIAL , "&_
															"FATURADO , "&_
															"RETIRADO , "&_
															"VENDIDO , "&_
															"REPASSE"&_
															")"&_
															"VALUES ("&_
															sqltexto(rs3_COD_FORMANDO)&", "&_
															"2, "&_
															"0, "&_
															sqlTexto(rs3_COD_CONTRATO)&", "&_
															rs3_COD_CURSO&", "&_
															sqltexto(rs3_NOME_CURSO)&", "&_
															sqltexto(rs3_NOME_FORMANDO)&", "&_
															sqltexto(rs3_ANO_CONCLUSAO)&", "&_
															sqltexto(Ucase(rs3_SEMESTRE_CONCLUSAO))&", "&_
															ALBUM2&", "&_
															INICIAL2&", "&_
															FATURADO2&", "&_
															RETIRADO2&", "&_
															FINALIZA_CORRECAO&", "&_
															REPASSE&_			
															")"					
															'Tb 1
															objConexao.Execute(sql)
															Response.Write(sql&"<br /><br /><br /><br />")
															'-------------
								
															sql = 	"INSERT INTO TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO ("&_
																	"COD_FORMANDO ,"&_
																	"ENDERECO ,"&_
																	"CEP ,"&_
																	"FONE_RES ,"&_
																	"FONE_CEL ,"&_
																	"BAIRRO ,"&_
																	"CIDADE ,"&_
																	"ESTADO ,"&_
																	"CADASTRADO_POR "&_
																	")"&_
																	"VALUES ("&_
																	rs3_COD_FORMANDO&", "&sqlTexto(rs3_ENDERECO)&", '"&rs3_CEP&"', '"&rs3_FONE_RES&"', '"&rs3_FONE_CEL&"', "&sqlTexto(rs3_BAIRRO)&", "&sqlTexto(rs3_CIDADE)&", "&_
																	sqlTexto(rs3_ESTADO)&", "&Session("COD_USUARIO")&_
																	")"
																	
																	'Tb 2
																	objConexao.Execute(sql)
																	Response.Write(sql&"<br /><br /><br /><br />")
																	'-------------																	
																	sql="INSERT INTO TB_PEDIDOS_FORMANDOS ("&_
																	"COD_PEDIDO ,"&_
																	"COD_FORMANDO)"&_ 
																	"VALUES("&COD_PEDIDO&", "&rs3_COD_FORMANDO&")"
																	objConexao.Execute(sql)
																	Response.Write(sql&"<br /><br /><br /><br />")
													
											End If
											
										rs3.Close
										Set rs3 = Nothing									
									End If
									
									If COD_FORMANDO_3_VALIDADO = 2 Then
									'Significa que está correto pode começar o processo!
										sql = "SELECT * FROM TB_FORMANDOS_IMP  WHERE "
										sql = sql & "CODIGO_FORMANDO= "&COD_FORMANDO3
										Set rs3 = objConexao.execute(sql)
										
											If not rs3.Eof Then						
													
													
													'Tb 1
													'0 devido a importação da formatura
													rs3_COD_FORMANDO		= rs3("CODIGO_FORMANDO")&"0"
													rs3_COD_CONTRATO		= rs3("CONTRATO")
													rs3_COD_CURSO			= rs3("CODIGO_CURSO")
													rs3_NOME_CURSO			= rs3("NOME_CURSO")
													rs3_NOME_FORMANDO		= ""&rs3("NOME_FORMANDO")
													'Correção do nome do formando
													posicaoInicial = inStr(rs3_NOME_FORMANDO,"(")
													If posicaoInicial > 0 Then
														rs3_NOME_FORMANDO = Trim(Left(rs3_NOME_FORMANDO,posicaoInicial-1))
													End If

													rs3_ANO_CONCLUSAO		= rs3("ANO")
													rs3_SEMESTRE_CONCLUSAO	= rs3("SEMESTRE")
													'------------------------------------
				
													'Tb2
													rs3_ENDERECO			= rs3("ENDERECO")
													rs3_CEP					= rs3("CEP")
													rs3_CEP					= Replace(rs3_CEP,"-","")
													rs3_FONE_RES			= rs3("FONE-RES")
													rs3_FONE_CEL			= rs3("FONE-COM-CEL")
													rs3_BAIRRO				= rs3("BAIRRO")
													rs3_CIDADE				= rs3("MUNICIPIO")
													rs3_ESTADO				= rs3("ESTADO")
													'------------------------------------
				
												
															sql = 	"INSERT INTO TB_FICHAS_IDENTIFICACAO ("&_
															"COD_FORMANDO , "&_
															"SISTEMA_NOVO , "&_
															"COD_PRODUCAO , "&_
															"COD_CONTRATO , "&_
															"COD_CURSO , "&_
															"NOME_CURSO , "&_
															"NOME_FORMANDO , "&_
															"ANO_CONCLUSAO , "&_
															"SEMESTRE_CONCLUSAO , "&_
															"ALBUM , "&_
															"INICIAL , "&_
															"FATURADO , "&_
															"RETIRADO , "&_
															"VENDIDO , "&_
															"REPASSE"&_
															")"&_
															"VALUES ("&_
															sqltexto(rs3_COD_FORMANDO)&", "&_
															"2, "&_
															"0, "&_
															rs3_COD_CONTRATO&", "&_
															rs3_COD_CURSO&", "&_
															sqltexto(rs3_NOME_CURSO)&", "&_
															sqltexto(rs3_NOME_FORMANDO)&", "&_
															sqltexto(rs3_ANO_CONCLUSAO)&", "&_
															sqltexto(Ucase(rs3_SEMESTRE_CONCLUSAO))&", "&_
															ALBUM3&", "&_
															INICIAL3&", "&_
															FATURADO3&", "&_
															RETIRADO3&", "&_
															FINALIZA_CORRECAO&", "&_
															REPASSE&_				
															")"				
															'Tb 1
															objConexao.Execute(sql)
															Response.Write(sql&"<br /><br /><br /><br />")
															'-------------
								
															sql = 	"INSERT INTO TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO ("&_
																	"COD_FORMANDO ,"&_
																	"ENDERECO ,"&_
																	"CEP ,"&_
																	"FONE_RES ,"&_
																	"FONE_CEL ,"&_
																	"BAIRRO ,"&_
																	"CIDADE ,"&_
																	"ESTADO ,"&_
																	"CADASTRADO_POR "&_
																	")"&_
																	"VALUES ("&_
																	rs3_COD_FORMANDO&", "&sqlTexto(rs3_ENDERECO)&", '"&rs3_CEP&"', '"&rs3_FONE_RES&"', '"&rs3_FONE_CEL&"', "&sqlTexto(rs3_BAIRRO)&", "&sqlTexto(rs3_CIDADE)&", "&_
																	sqlTexto(rs3_ESTADO)&", "&Session("COD_USUARIO")&_
																	")"
																	
																	'Tb 2
																	objConexao.Execute(sql)
																	Response.Write(sql&"<br /><br /><br /><br />")
																	'-------------																	
																	sql="INSERT INTO TB_PEDIDOS_FORMANDOS ("&_
																	"COD_PEDIDO ,"&_
																	"COD_FORMANDO)"&_ 
																	"VALUES("&COD_PEDIDO&", "&rs3_COD_FORMANDO&")"
																	objConexao.Execute(sql)
																	Response.Write(sql&"<br /><br /><br /><br />")
													
											End If
											
										rs3.Close
										Set rs3 = Nothing									
									End If
									
									If COD_FORMANDO_4_VALIDADO = 2 Then
									'Significa que está correto pode começar o processo!
										sql = "SELECT * FROM TB_FORMANDOS_IMP  WHERE "
										sql = sql & "CODIGO_FORMANDO= "&COD_FORMANDO4
										Set rs3 = objConexao.execute(sql)
										
											If not rs3.Eof Then						
													
													
													'Tb 1
													'0 devido a importação da formatura
													rs3_COD_FORMANDO		= rs3("CODIGO_FORMANDO")&"0"
													rs3_COD_CONTRATO		= rs3("CONTRATO")
													rs3_COD_CURSO			= rs3("CODIGO_CURSO")
													rs3_NOME_CURSO			= rs3("NOME_CURSO")
													rs3_NOME_FORMANDO		= ""&rs3("NOME_FORMANDO")
													'Correção do nome do formando
													posicaoInicial = inStr(rs3_NOME_FORMANDO,"(")
													If posicaoInicial > 0 Then
														rs3_NOME_FORMANDO = Trim(Left(rs3_NOME_FORMANDO,posicaoInicial-1))
													End If

													rs3_ANO_CONCLUSAO		= rs3("ANO")
													rs3_SEMESTRE_CONCLUSAO	= rs3("SEMESTRE")
													'------------------------------------
				
													'Tb2
													rs3_ENDERECO			= rs3("ENDERECO")
													rs3_CEP					= rs3("CEP")
													rs3_CEP					= Replace(rs3_CEP,"-","")
													rs3_FONE_RES			= rs3("FONE-RES")
													rs3_FONE_CEL			= rs3("FONE-COM-CEL")
													rs3_BAIRRO				= rs3("BAIRRO")
													rs3_CIDADE				= rs3("MUNICIPIO")
													rs3_ESTADO				= rs3("ESTADO")
													'------------------------------------
				
												
															sql = 	"INSERT INTO TB_FICHAS_IDENTIFICACAO ("&_
															"COD_FORMANDO , "&_
															"SISTEMA_NOVO , "&_
															"COD_PRODUCAO , "&_
															"COD_CONTRATO , "&_
															"COD_CURSO , "&_
															"NOME_CURSO , "&_
															"NOME_FORMANDO , "&_
															"ANO_CONCLUSAO , "&_
															"SEMESTRE_CONCLUSAO , "&_
															"ALBUM , "&_
															"INICIAL , "&_
															"FATURADO , "&_
															"RETIRADO , "&_
															"VENDIDO , "&_
															"REPASSE"&_
															")"&_
															"VALUES ("&_
															sqltexto(rs3_COD_FORMANDO)&", "&_
															"2, "&_
															"0, "&_
															rs3_COD_CONTRATO&", "&_
															rs3_COD_CURSO&", "&_
															sqltexto(rs3_NOME_CURSO)&", "&_
															sqltexto(rs3_NOME_FORMANDO)&", "&_
															sqltexto(rs3_ANO_CONCLUSAO)&", "&_
															sqltexto(Ucase(rs3_SEMESTRE_CONCLUSAO))&", "&_
															ALBUM4&", "&_
															INICIAL4&", "&_
															FATURADO4&", "&_
															RETIRADO4&", "&_
															FINALIZA_CORRECAO&", "&_
															REPASSE&_				
															")"				
															'Tb 1
															objConexao.Execute(sql)
															Response.Write(sql&"<br /><br /><br /><br />")
															'-------------
								
															sql = 	"INSERT INTO TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO ("&_
																	"COD_FORMANDO ,"&_
																	"ENDERECO ,"&_
																	"CEP ,"&_
																	"FONE_RES ,"&_
																	"FONE_CEL ,"&_
																	"BAIRRO ,"&_
																	"CIDADE ,"&_
																	"ESTADO ,"&_
																	"CADASTRADO_POR "&_
																	")"&_
																	"VALUES ("&_
																	rs3_COD_FORMANDO&", "&sqlTexto(rs3_ENDERECO)&", '"&rs3_CEP&"', '"&rs3_FONE_RES&"', '"&rs3_FONE_CEL&"', "&sqlTexto(rs3_BAIRRO)&", "&sqlTexto(rs3_CIDADE)&", "&_
																	sqlTexto(rs3_ESTADO)&", "&Session("COD_USUARIO")&_
																	")"
																	
																	'Tb 2
																	objConexao.Execute(sql)
																	Response.Write(sql&"<br /><br /><br /><br />")
																	'-------------																	
																	sql="INSERT INTO TB_PEDIDOS_FORMANDOS ("&_
																	"COD_PEDIDO ,"&_
																	"COD_FORMANDO)"&_ 
																	"VALUES("&COD_PEDIDO&", "&rs3_COD_FORMANDO&")"
																	objConexao.Execute(sql)
																	Response.Write(sql&"<br /><br /><br /><br />")
											End If
											
										rs3.Close
										Set rs3 = Nothing									
									End If
									
									'/////////////////////////////////////////////////////////////////////					
								'Correções:
								'///
								If DATA_DESC_DIR = "1900-1-1" Then
									DATA_DESC_DIR = "NULL"
								End If
								
								If Len(DATA_BAIXADO_DIRETORIA) > 2 Then
									DATA_BAIXADO_DIRETORIA = "'"&Year(DATA_BAIXADO_DIRETORIA)&"-"&Month(DATA_BAIXADO_DIRETORIA)&"-"&Day(DATA_BAIXADO_DIRETORIA)&" "&HORA_BAIXADO_DIRETORIA&"'"			
								Else
									DATA_BAIXADO_DIRETORIA = "NULL"
								End If
								
								Dim VENDEDOR_RECEBEU_EM
								VENDEDOR_RECEBEU_EM = COD_ANO_BAIXA&"-"&COD_MES_BAIXA&"-01"
								
								'///
								'If Len(BRINDES) > 2 Then
								'OBSERVACAO = OBSERVACAO&" - Brinde Formando 1("&BRINDES&")"
								'End If
								
								'If Len(BRINDES2) > 2 Then
								'OBSERVACAO = OBSERVACAO&" - Brinde Formando 2("&BRINDES2&")"
								'End If
								
									'LANÇAMENTO DO PEDIDO
								sql="INSERT INTO TB_PEDIDOS ("&_
									"COD_PEDIDO ,"&_
									"SISTEMA_NOVO ,"&_
									"DATA_VENDA ,"&_
									"COD_VENDEDOR ,"&_
									"DATA_CADASTRO ,"&_
									"FINALIZA_VENDEDOR ,"&_
									"FINALIZA_CORRECAO ,"&_
									"FINALIZA_DIRETORIA ,"&_
									"DATA_CORRIGIDO_DEP_FOTO ,"&_
									"DATA_DESCEU_DIRETORIA ,"&_
									"DATA_CORRIGIDO_DIRETORIA ,"&_
									"OBSERVACAO ,"&_
									"VALOR_AJUDA_CUSTO ,"&_
									"VENDEDOR_RECEBE_EM ,"&_
									"PAGAMENTO_EM_DINHEIRO "&_
									")"&_
									"VALUES ("&_
									COD_PEDIDO&",2,"&sqlData(DATA)&", "&CODIGO_VENDEDOR&", "&sqlData(DATA)&",'"&FINALIZA_VENDEDOR&"', '"&FINALIZA_CORRECAO&"', '"&FINALIZA_DIRETORIA&"', "&sqlData(DATA_DESC_DIR)&", "&sqlData(DATA_DESC_DIR)&", "&DATA_BAIXADO_DIRETORIA&", "&sqltexto(OBSERVACAO)&", '"&sqlNumero(VALOR_AJUDA_CUSTO)&"', "&sqlData(VENDEDOR_RECEBEU_EM)&", '"&sqlNumero(EM_DINHEIRO)&"')"
									
									objConexao.Execute(sql)
									Response.Write(sql&"<br /><br />")
									
									'Observação os formandos foram inseridos anteriormente na TB_PEDIDOS_FORMANDOS
									'Neste momento há as fichas, o pedido lançado, entretanto falta a forma de pagamento e os produtos
									
									'1) TB_FOTO_VENDA_CARTAO
									'2) TB_FOTO_VENDA_CHEQUES
									'3) TB_FOTO_PRODUTOS
									
									Set rs3_1 = objConexao3.Execute("SELECT * FROM TB_FOTO_VENDA_CARTAO WHERE NUMERO_PEDIDO="&NUMERO_PEDIDO)
												
												While not rs3_1.Eof
												'Existe pagamentos em cartão insere na tb nova
													
													sql="INSERT INTO TB_PEDIDOS_CARTAO ("&_
														"COD_PEDIDO ,"&_
														"CARTAO ,"&_
														"VALOR_TOTAL ,"&_
														"PARCELAS ,"&_
														"NOME_TITULAR ,"&_
														"GRAU_PARENTESCO ,"&_
														"AUTENTICACAO ,"&_
														"DOCUMENTO ,"&_
														"DATA_AUTENTICACAO ,"&_
														"VERIFICADO "&_
														")"&_
														"VALUES ("&_
														COD_PEDIDO&",'"&rs3_1("CARTAO")&"','"&sqlNumero(rs3_1("VALOR"))&"','"&rs3_1("PARCELAS")&"',"&sqlTexto(Ucase(""&rs3_1("NOME_TITULAR")))&","&sqlTexto(rs3_1("GRAU_PARENTESCO"))&_
														",'"&rs3_1("AUTENTICACAO")&"','"&rs3_1("DOCUMENTO")&"',"&sqlData(rs3_1("DATA_AUTENTICACAO"))&","&rs3_1("VERIFICADO")&_
														")"
														
														
														Response.Write(sql&"<br /><br />")
														objConexao.Execute(sql)
												rs3_1.MoveNext()
												Wend
									
									rs3_1.Close
									Set rs3_1 = Nothing
									
									
									Set rs3_2 = objConexao3.Execute("SELECT * FROM TB_FOTO_VENDA_CHEQUES WHERE NUMERO_PEDIDO="&NUMERO_PEDIDO)
												
												While not rs3_2.Eof
													Dim CPF_CHEQUE
													CPF_CHEQUE = rs3_2("CPF")
													CPF_CHEQUE = Replace(CPF_CHEQUE,"'","")
														
													sql="INSERT INTO TB_PEDIDOS_CHEQUE ("&_
														"COD_PEDIDO ,"&_
														"BANCO ,"&_
														"AGENCIA ,"&_
														"CONTA_CORRENTE ,"&_
														"NUMERO_CHEQUE ,"&_
														"VALOR ,"&_
														"BOM_PARA ,"&_
														"NOME ,"&_
														"CPF ,"&_
														"SE_TERCEIRO ,"&_
														"FONE ,"&_
														"VERIFICADO "&_
														")"&_
														"VALUES ("&_
														COD_PEDIDO&", '"&rs3_2("BANCO")&"', '"&rs3_2("AGENCIA")&"', '"&rs3_2("CONTA_CORRENTE")&"', '"&rs3_2("NUMERO_CHEQUE")&_
														"', '"&sqlNumero(rs3_2("VALOR"))&"', "&sqlData(rs3_2("BOM_PARA"))&", "&sqlTexto(UCASE(rs3_2("NOME")))&", '"&CPF_CHEQUE&"', '"&rs3_2("SE_TERCEIRO")&"', '"&rs3_2("FONE")&"', '"&rs3_2("VERIFICADO")&_
														"')"	
														Response.Write(sql&"<br /><br />")
														objConexao.Execute(sql)
														
												rs3_2.MoveNext()
												Wend
									
									rs3_2.Close
									Set rs3_2 = Nothing
									
									Set rs3_3 = objConexao3.Execute("SELECT * FROM TB_FOTO_PRODUTOS WHERE NUMERO_PEDIDO="&NUMERO_PEDIDO)
												
												While not rs3_3.Eof
													
													'Posteriormente corrigir os códigos de produtos "12"
																										
													
													If METODO_B <> 1 Then
													sql="INSERT INTO TB_PEDIDOS_PRODUTOS ("&_
														"COD_PEDIDO ,"&_
														"COD_FORMANDO ,"&_
														"COD_PRODUTO ,"&_
														"VALOR_UN ,"&_
														"QUANTIDADE "&_
														")"&_
														"VALUES ("&_
									COD_PEDIDO&", "&COD_FORMANDO&"0, "&rs3_3("PRODUTO")&", '"&sqlNumero(rs3_3("VALOR"))&"', '"&sqlNumero(rs3_3("QUANTIDADE"))&_
														"')"
													
													Else
													
													sql="INSERT INTO TB_PEDIDOS_PRODUTOS ("&_
														"COD_PEDIDO ,"&_
														"COD_FORMANDO ,"&_
														"COD_PRODUTO ,"&_
														"VALOR_UN ,"&_
														"QUANTIDADE "&_
														")"&_
														"VALUES ("&_
									COD_PEDIDO&", "&COD_FORMANDO_ANTIGO_1&", "&rs3_3("PRODUTO")&", '"&sqlNumero(rs3_3("VALOR"))&"', '"&sqlNumero(rs3_3("QUANTIDADE"))&_
														"')"
														
													End If	
														
														
														
														Response.Write(sql&"<br /><br />")
														objConexao.Execute(sql)
														
												rs3_3.MoveNext()
												Wend
									
									rs3_3.Close
									Set rs3_3 = Nothing
						
						


									sql = 	"UPDATE TB_FOTO_VENDA SET EXPORTADO_SISTEMA_NOVO = '1', "&_
											"DATA_EXPORTADO = NOW( ) WHERE COD_PEDIDO ="&COD_PEDIDO&" LIMIT 1"
									
									'Faz o UPTADE que foi importado e encerra o processo!
									objConexao3.execute(sql)
									Response.Write(sql&"<br />_________________________________________________<br />")
									
					End If
					
					Response.Write("<br />")
					
					
					
					
					
					
					
					
					
					
				rs.MoveNext()
				Wend
				Response.Write("<br />")
				Response.Write("<br />")
				Response.Write("<br />")
				Response.Write("Não importados: "&CONTA_NAO_IMPORTADOS)
				Response.Write("<br />")
				Response.Write("Importados: "&CONTA_IMPORTADOS)
				Response.Write("<br />")
				Response.Write("<br />")
			Else
			
				msg="Não dados a serem importados !"
			
			End If
		  
		  rs.Close		  
		  Set rs = Nothing
		  
%>

		  Total Importados: <%=BRINDES_IMPORTADOS%><br />
  		  Total Não Importados: <%=BRINDES_NAO_IMPORTADOS%><br />

		  
		  
		  <%
		   
		  
		  
		  
		  fecharConexao3()
		  fecharConexao()
		  
		  If redirecionar_ok = 1 Then
		  	Response.Redirect("lca_pedidos_pendentes.asp?importado")
		  End If
		  %>
          
          <p align="center">&nbsp;</p>
		  <p align="center">&nbsp;</p>		
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