<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim COD_FORMANDO
Dim msg
Dim comando
Dim VALIDA_SUB7, VALIDA_SUB9, VALIDA_SUB24, VALIDA_SUB25, VALIDA_SUB26, VALIDA_SUB28, VALIDA_SUB68
VALIDA_SUB7 =  formatNumber(0&inStr(sub_validacoes_paginas,"#7"),0)
VALIDA_SUB9	=  formatNumber(0&inStr(sub_validacoes_paginas,"#9"),0)
VALIDA_SUB24 = formatNumber(0&inStr(sub_validacoes_paginas,"#24"),0)
VALIDA_SUB25 = formatNumber(0&inStr(sub_validacoes_paginas,"#25"),0)
VALIDA_SUB26 = formatNumber(0&inStr(sub_validacoes_paginas,"#26"),0)
VALIDA_SUB28 = formatNumber(0&inStr(sub_validacoes_paginas,"#28"),0)
VALIDA_SUB68 = formatNumber(0&inStr(sub_validacoes_paginas,"#68"),0)

COD_FORMANDO = Request.QueryString("cod_formando")
msg = Request.QueryString("msg")
comando = Request.QueryString("comando")

AbrirConexao()
		'PARTE1
		Dim COD_CONTRATO, SISTEMA_NOVO, NOME_CURSO, NOME_FORMANDO, CPF, ANO_CONCLUSAO, SEMESTRE_CONCLUSAO, ALBUM, INICIAL, BRINDE, COMISSAO, PRODUTO_AGREGADO, VENDIDO, ENCADERNACOES_EXTRAS, COD_PRODUCAO, COD_PRODUCAO2,	CREDITO_EM_VALOR, CREDITO_MOTIVO, CREDITO_DATA_INSERIDO
		'PARTE2
		Dim ENDERECO, CEP, FONE_RES, FONE_CEL, BAIRRO, CIDADE, ESTADO, graduateReportInstitution, graduateReportGifts, graduateReportProducts
				
		sql = "SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO="&COD_FORMANDO
		
		Set rs=objConexao.Execute(sql)
		
		If not rs.Eof Then
				
				ENCADERNACOES_EXTRAS = formatNumber(0&rs("ENCADERNACOES_EXTRAS"),0)
				COD_CONTRATO		= rs("COD_CONTRATO")
				SISTEMA_NOVO		= formatNumber(0&rs("SISTEMA_NOVO"),0)
				NOME_CURSO			= rs("NOME_CURSO")
				NOME_FORMANDO		= rs("NOME_FORMANDO")
				CPF					= rs("CPF")
				EMAIL_FORMANDO		= rs("EMAIL_FORMANDO")
				ANO_CONCLUSAO		= rs("ANO_CONCLUSAO")
				SEMESTRE_CONCLUSAO	= rs("SEMESTRE_CONCLUSAO")
				ALBUM				= rs("ALBUM")
				ALBUM_DIGITAL		= rs("ALBUM_DIGITAL")
				NUMERO_PASTA_GERENCIA = rs("NUMERO_PASTA_GERENCIA")
				NUMERO_PASTA_VENDEDOR = rs("NUMERO_PASTA_VENDEDOR")
				INICIAL				= rs("INICIAL")
				BRINDE				= rs("BRINDE")
				COMISSAO			= formatNumber(0&rs("COMISSAO"),0)
				PRODUTO_AGREGADO	= rs("PRODUTO_AGREGADO")
				VENDIDO				= rs("VENDIDO")
				COD_PRODUCAO		= formatNumber(0&rs("COD_PRODUCAO"),0)
				COD_PRODUCAO2		= rs("COD_PRODUCAO")
				CREDITO_EM_VALOR	= formatNumber(0&rs("CREDITO_EM_VALOR"))
				CREDITO_MOTIVO		= ""&rs("CREDITO_MOTIVO")
				CREDITO_DATA_INSERIDO = ""&rs("CREDITO_DATA_INSERIDO")

				If SEMESTRE_CONCLUSAO = "J" Then
					SEMESTRE_CONCLUSAO = "JULHO"
				Else
					SEMESTRE_CONCLUSAO = "DEZEMBRO"
				End If
							
		End If
				
		
		rs.Close
		Set rs=Nothing
		
		Dim qtdEnderecos
		Set rs2 = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO WHERE COD_FORMANDO="&COD_FORMANDO&" ORDER BY DATA_CADASTRO DESC",qtdEnderecos)
		
		If not rs2.Eof Then
		
				ENDERECO			= rs2("ENDERECO")
				CEP					= rs2("CEP")
				FONE_RES			= rs2("FONE_RES")
				FONE_CEL			= rs2("FONE_CEL")
				BAIRRO				= rs2("BAIRRO")
				CIDADE				= rs2("CIDADE")
				ESTADO				= rs2("ESTADO")
				
		End If
		
		rs2.Close
		Set rs2 = Nothing		
		
FecharConexao()


'//////////////

Dim mensagem, retornar_contato

If Request.QueryString("comando") = "gravar_msg" Then
	
		retornar_contato = formatNumber(0&Request.Form("retornar_contato"),0)
		mensagem = ""&Request.Form("mensagem")
		
		If mensagem <> "" Then
		
			AbrirConexao()
		
				objConexao.Execute("INSERT INTO TB_FICHAS_IDENTIFICACAO_HISTORICO (COD_FORMANDO,GERADO_POR,MENSAGEM) VALUES ('"&cod_formando&"','"&Session("cod_usuario")&"',"&sqlTexto(mensagem)&")")
		
					If int(retornar_contato) = 1 Then
			
						objConexao.Execute("INSERT INTO TB_FICHAS_IDENTIFICACAO_RETORNAR_CONTATO (COD_FORMANDO,COD_USUARIO) VALUES ('"&COD_FORMANDO&"','"&Session("cod_usuario")&"')")
			
					End If
		
			FecharConexao()
	
			Response.Redirect("dip_busca_detalhamento_formando.asp?cod_formando="&cod_formando&"&msg=Histórico de mensagem gravada com sucesso!")
		
		End If
End If


If Request.QueryString("comando") = "voltar_local_provisorio" Then
		
		AbrirConexao()
			objConexao.Execute("UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET COD_CAIXA_ARMAZENADO = '0' WHERE COD_MOVIMENTACAO = "&Request.QueryString("id")&" AND COD_FORMANDO = "&cod_formando)
		FecharConexao()
	
		Response.Redirect("dip_busca_detalhamento_formando.asp?cod_formando="&cod_formando&"&msg=Local provisório retornado com sucesso!")
		
End If
'///////////////

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Detalhes do formando | Portal Stillos</title>
<link rel="stylesheet" href="css/vendor/bootstrap.min.css" />
<link rel="stylesheet" href="css/dashboard.css" />
<style type="text/css">
	/* CSS for the demo. CSS needed for the scripts are loaded dynamically by the scripts */
	#mainContainer{
		width:500px;
		margin:0 auto;
		margin-top:10px;
		border:1px double #000;
		padding:3px;
		
	}
	
	/* Alternative layout for the message box - used by one of the messages */
	
	.modalDialog_contentDiv_error{
		border:3px solid #FFF;	
		padding:2px;
		z-index:100;/* Always needed	*/
		position:absolute;	/* Always needed	*/
		background-color:#F00;	/* White background color for the message */
		color:#FFF;
	}
	.modalDialog_contentDiv_error a{
		color:#FFF;
	}
	h1{
		margin-top:0px;
	}

	</style>	
	<script type="text/javascript">
	function verify(ver){
		
		if(ver){
			// Confirmed message, i.e. clicked on "Yes"
			alert('Message confirmed');
		}else{
			// Clicked on "No"
			alert('Message not confirmed');
		}
	}
	</script>
	<link rel="stylesheet" href="css/modal-message.css" type="text/css">
	<script type="text/javascript" src="js/ajax.js"></script>
	<script type="text/javascript" src="js/modal-message.js"></script>
	<script type="text/javascript" src="js/ajax-dynamic-content.js"></script>
	<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>

	<script type="text/javascript">
messageObj = new DHTML_modalMessage();	// We only create one object of this class
messageObj.setShadowOffset(5);	// Large shadow


function displayMessage(url)
{
	messageObj.setSource(url);
	messageObj.setCssClassMessageBox(false);
	messageObj.setSize(550,450);
	messageObj.setShadowDivVisible(true);	// Enable shadow for these boxes
	messageObj.display();
}

function displayStaticMessage(messageContent,cssClass,altura)
{
	messageObj.setHtmlContent(messageContent);
	messageObj.setSize(550,altura);
	messageObj.setCssClassMessageBox(cssClass);
	messageObj.setSource(false);	// no html source since we want to use a static message here.
	messageObj.setShadowDivVisible(false);	// Disable shadow for these boxes	
	messageObj.display();
	
	
}

function closeMessage()
{
	messageObj.close();	
}


</script>
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
	<!--AJUDA F1 E CTRL PARA BUSCA R?PIDA-->
		
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
 


	<!--AJUDA F1 E CTRL PARA BUSCA R?PIDA-->
	</script>
		<!-- CALEND?RIO SCRIPTS
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
<script>
function confirma() {
	
	if(YNconfirm("Mensagem")){
			document.dados.submit();			
	}	
}

function confirmaVoltarLocalProvisorio(id){
	if(YNconfirm("Tem certeza que deseja voltar para o local provis&oacute;rio de armazenamento?")){
		abre('dip_busca_detalhamento_formando.asp?cod_formando=<%=cod_formando%>&comando=voltar_local_provisorio&id='+id);
	}		
}

</script>
<script language="javascript">
function YNconfirm(txt){
return typeof(suporteVBscript)=="undefined"?confirm(txt):VBconfirm(txt)==6
}
</script>
<script language="vbscript">
suporteVBscript=1
function VBconfirm(mensagem)

 VBconfirm=msgbox(mensagem,vbyesno)
end function

</script>
<script>
function confirmaProdAparte(PRODUTO_A_PARTE){

	if(YNconfirm("Tem certeza que deseja excluir este produto?")){
			location.href = "?cod_pedido=<%=COD_PEDIDO%>&prod_excluir_a_parte="+PRODUTO_A_PARTE
	}

}
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  if (selObj.value !== ""){
  	eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  	if (restore) selObj.selectedIndex=0;
  }
}
function popUp(URL){
	win=window.open(URL,"","width=350,height=220,scrollbars=yes")
	var text = "Se a janela n&atilde;o abrir desbloquei os Pop-ups!";
	if(win == null) { alert(text); return; }
	win.moveTo(350,150);
}
</script>
<style type="text/css">
<!--
.style1 {font-weight: bold}
-->
</style>
<style>
.graduate-page .graduate-content { min-width:0; color:#292738; }
.graduate-page .graduate-heading { align-items:flex-start; }
.graduate-page .graduate-heading h1 { margin:0; padding:0; color:#242738; font:700 clamp(23px,3vw,34px)/1.18 system-ui,sans-serif; overflow-wrap:anywhere; }
.graduate-page .graduate-heading { padding:20px; border:1px solid #e4dfeb; border-radius:12px; background:#fff; }
.graduate-page .graduate-heading .eyebrow { margin:0 0 6px; }
.graduate-page .graduate-actions, .graduate-page .graduate-jump { display:flex; flex-wrap:wrap; gap:8px; }
.graduate-page .graduate-jump { margin:0 0 20px; }
.graduate-page .graduate-jump a { padding:6px 12px; border:1px solid #e4dfeb; border-radius:8px; background:#fff; color:#6d1749; font-weight:650; text-decoration:none; }
.graduate-page .graduate-content table[border="1"] { border:1px solid #e4dfeb; background:#fff; }
.graduate-page .graduate-content table[border="1"] > tbody > tr > td { padding:8px 10px; border-color:#eeeaf2; overflow-wrap:anywhere; }
.graduate-page .graduate-content tr[bgcolor="whitesmoke"] > td { background:#f1edf6; color:#3e344e; }
.graduate-page .graduate-content .graduate-summary { height:auto; margin:0; }
.graduate-page .graduate-content .graduate-summary h2, .graduate-page .graduate-content h2.graduate-section-heading { color:#292738; font:700 clamp(17px,2vw,23px)/1.3 system-ui,sans-serif; margin:0 0 12px; }
.graduate-page .graduate-panel { height:auto; margin-bottom:20px; }
.graduate-page .graduate-panel h2 { scroll-margin-top:18px; }
.graduate-page .graduate-summary table { width:100%; border-collapse:collapse; }
.graduate-page .graduate-summary > table > tbody > tr > td:first-child:not([colspan]) { background:#f8f6fa; width:26%; }
.graduate-page .graduate-summary > table > tbody > tr > td:nth-child(2) { width:74%; }
.graduate-page .graduate-summary > table > tbody > tr > td { border-bottom:1px solid #eeeaf2; }
.graduate-page .graduate-scroll { max-width:100%; overflow-x:auto; }
.graduate-page .graduate-scroll table { min-width:680px; }
.graduate-page .graduate-content :is(input,select,textarea,iframe,.btnGeral,table) { max-width:100%; }
.graduate-page .graduate-content .btnGeral { display:inline-block; padding:7px 10px; margin:3px 0; border:1px solid #ddd5e5; border-radius:7px; background:#fff; color:#7c174e; font-weight:650; white-space:normal; }
.graduate-page .graduate-content a:focus-visible { outline:2px solid #9e125c; outline-offset:2px; }
.graduate-page .dashboard-nav a.active:not([aria-current="page"]) { background:transparent; color:#d4d5e1; }
.graduate-page .dashboard-nav a[aria-current="page"] { background:#b20b61; color:#fff; font-weight:600; }
@media (max-width:560px) {
    .graduate-page .graduate-summary > table,
    .graduate-page .graduate-summary > table > tbody,
    .graduate-page .graduate-summary > table > tbody > tr,
    .graduate-page .graduate-summary > table > tbody > tr > td { display:block; width:100%; }
    .graduate-page .graduate-summary > table > tbody > tr > td:first-child:not([colspan]),
    .graduate-page .graduate-summary > table > tbody > tr > td:nth-child(2) { width:100%; }
    .graduate-page .graduate-content { overflow-wrap:anywhere; }
}
.graduate-page .graduate-content textarea[name="mensagem"] { width:100%; min-height:86px; padding:8px; border:1px solid #ded5e5; border-radius:8px; background:#fff; }
.graduate-page .graduate-compose { padding:14px; margin-bottom:16px; border:1px solid #e4dfeb; border-radius:10px; background:#faf8fc; }
.graduate-page .graduate-compose label[for="graduateMessage"] { display:block; margin-bottom:7px; font-weight:700; }
.graduate-page .graduate-compose-actions { display:flex; flex-wrap:wrap; align-items:center; justify-content:space-between; gap:10px; margin-top:8px; }
.graduate-page .graduate-compose-actions label { display:flex; align-items:center; gap:6px; margin:0; }
.graduate-page .graduate-content .graduate-submit { width:auto; min-width:78px; height:auto; }
.graduate-page .graduate-panel .graduate-scroll table { min-width:560px; }
.graduate-page .graduate-panel .graduate-scroll table > tbody > tr:first-child > td { background:#f1edf6; }
.graduate-page .graduate-panel .graduate-scroll table > tbody > tr > td { border-bottom:1px solid #eeeaf2; }
.graduate-page .graduate-movement-table { width:100%; border-collapse:collapse; text-align:left; }
.graduate-page .graduate-movement-table th,
.graduate-page .graduate-movement-table td { padding:12px 14px; border:0; border-bottom:1px solid #eeeaf2; vertical-align:top; text-align:left; overflow-wrap:anywhere; }
.graduate-page .graduate-movement-table th { background:#f1edf6; color:#3e344e; font-weight:700; }
.graduate-page .graduate-movement-table .graduate-movement-previous > td { background:#faf8fc; }
.graduate-page .graduate-movement-item strong { display:block; }
.graduate-page .graduate-movement-tag { display:inline-block; margin-top:3px; padding:3px 8px; border-radius:6px; background:#eee8f4; color:#564468; font-size:.85em; font-weight:700; }
.graduate-page .graduate-movement-pending { background:#f9e7ec; color:#872047; }
.graduate-page .graduate-movement-label { display:none; }
@media (max-width:900px) {
    .graduate-page .graduate-movement-table,
    .graduate-page .graduate-movement-table tbody { display:block; }
    .graduate-page .graduate-movement-table thead { position:absolute; width:1px; height:1px; overflow:hidden; clip-path:inset(50%); }
    .graduate-page .graduate-movement-table tbody > tr { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); margin:0 0 10px; border:1px solid #e4dfeb; border-radius:10px; overflow:hidden; }
    .graduate-page .graduate-movement-table tbody > tr:last-child { margin-bottom:0; }
    .graduate-page .graduate-movement-table td { display:block; width:auto; border:0; }
    .graduate-page .graduate-movement-table td:first-child { grid-column:1 / -1; border-bottom:1px solid #eeeaf2; }
    .graduate-page .graduate-movement-label { display:block; margin-bottom:4px; color:#564468; font-size:.85em; font-weight:700; }
}
@media (max-width:420px) {
    .graduate-page .graduate-movement-table tbody > tr { grid-template-columns:minmax(0,1fr); }
}
.graduate-page .graduate-content .graduate-summary { margin-bottom:20px; }
.graduate-page .session-notices a[href="adm_etiquetas_agendadas.asp"],
.graduate-page .graduate-label-trigger { display:inline-block; border:1px solid #ead6e2; border-radius:7px; padding:6px 11px; background:#fff; color:#88134f; font-weight:650; text-decoration:none; }
.graduate-page .session-notices a[href="adm_etiquetas_agendadas.asp"] { margin-bottom:10px; }
.graduate-label-dialog { width:min(960px, calc(100% - 24px)); max-height:calc(100dvh - 24px); padding:0; border:1px solid #e4dfeb; border-radius:14px; background:#fff; color:#242738; box-shadow:0 18px 60px #24273833; }
.graduate-label-dialog::backdrop { background:#24273899; }
.graduate-label-dialog-header { display:flex; justify-content:space-between; align-items:center; gap:12px; padding:14px 18px; border-bottom:1px solid #e4dfeb; }
.graduate-label-dialog-header h2 { margin:0; font:700 17px/1.3 system-ui,sans-serif; }
.graduate-label-dialog-header form { margin:0; }
.graduate-label-frame { display:block; width:100%; height:min(680px, calc(100dvh - 105px)); border:0; }
.graduate-label-frame[hidden] { display:none; }
.graduate-label-loading { padding:18px; margin:0; }

.graduate-report-actions { display:flex; flex-wrap:wrap; gap:8px; align-items:center; }
.graduate-report-dialog { width:min(820px,calc(100% - 24px)); }
.graduate-report-preview { padding:16px 20px; }
.graduate-report-options { display:flex; gap:8px; align-items:center; margin:0 0 12px; }
.graduate-report { color:#242738; background:#fff; overflow-wrap:anywhere; }
.graduate-report-header { display:flex; align-items:center; justify-content:space-between; gap:12px; border-bottom:2px solid #b20b61; padding-bottom:10px; }
.graduate-report-header h2 { margin:0; font:700 clamp(17px,3vw,23px)/1.25 system-ui,sans-serif; }
.graduate-report-logo { width:clamp(88px,20vw,140px); height:auto; border-radius:6px; background:#292b40; }
.graduate-report-name { margin:14px 0 8px; font-size:1.15em; font-weight:700; }
.graduate-report-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:8px; margin:0 0 12px; }
.graduate-report-grid > div { border:1px solid #e4dfeb; border-radius:6px; padding:7px 9px; }
.graduate-report-grid dt { font-size:.88em; color:#53566b; }
.graduate-report-grid dd { margin:2px 0 0; font-weight:600; }
.graduate-report-items { margin:0; padding-left:18px; }
.graduate-report-wide { grid-column:1/-1; }
.graduate-report-contact h3 { font-size:1em; margin:0 0 6px; }
#graduateReportPrint,[hidden].graduate-report-contact { display:none; }
@page graduate-summary { size:A4 portrait; margin:12mm; }
@media print {
  body.graduate-printing > :not(#graduateReportPrint) { display:none !important; }
  body.graduate-printing #graduateReportPrint { display:block; page:graduate-summary; }
  body.graduate-printing { margin:0; background:#fff; }
  #graduateReportPrint .graduate-report { color:#111; font:10pt/1.35 system-ui,sans-serif; }
  #graduateReportPrint .graduate-report-logo { width:32mm; background:transparent; filter:brightness(0); }
  #graduateReportPrint .graduate-report-header { border-color:#444; }
  #graduateReportPrint .graduate-report-grid > div { break-inside:avoid; }
  #graduateReportPrint h2,#graduateReportPrint h3 { break-after:avoid; }
}
</style>
</head>

<body class="graduate-page">
<a class="skip-link" href="#conteudo">Ir para o conte&uacute;do</a>
<div class="dashboard-layout">
  <aside class="dashboard-sidebar" id="menuPrincipal">
    <a class="dashboard-brand" href="adm_index.asp"><img src="images/logoparafundopreto.png" alt="Stillos" width="180" /><span>GERENCIAL FOTOS</span></a>
    <div class="nav-caption">&Aacute;REA DE TRABALHO</div>
    <!-- #include file="includes/dashboard_menu.asp" -->
  </aside>
  <div class="dashboard-workspace">
    <header class="dashboard-topbar">
      <button class="btn menu-toggle" id="ppvMenuToggle" type="button" aria-controls="menuPrincipal" aria-expanded="false">Menu</button>
      <div class="topbar-title">Portal Stillos <span>/ Detalhes do formando</span></div>
      <div class="topbar-actions"><a class="user-name" href="adm_log_usuarios.asp"><%=Server.HTMLEncode(Session("NOME") & "")%></a><a href="manual_sistema.asp" target="_blank" rel="noopener">Manual</a><a class="btn btn-outline-secondary btn-sm" href="adm_sair.asp">Sair</a></div>
    </header>
    <main class="dashboard-main ppv-main" id="conteudo">
      <div class="graduate-content">
        <div class="dashboard-heading graduate-heading">
          <div>
            <p class="eyebrow">FORMANDOS</p>
            <h1><%=Server.HTMLEncode(NOME_FORMANDO & "")%></h1>
            <p class="panel-description">Ficha n&ordm; <%=digito(COD_FORMANDO)%></p>
          </div>
          <div class="graduate-report-actions">
            <%If Len(NOME_FORMANDO & "") > 0 Then%><button class="btn btn-primary" id="graduateReportOpen" type="button" aria-haspopup="dialog">Resumo para impress&atilde;o</button><%End If%>
            <a class="btn btn-outline-secondary" href="adm_index.asp" id="orderBack">Voltar</a>
          </div>
        </div>
        <div class="search-panel"><div><strong>Busca no portal</strong></div><!-- #include file="includes/dashboard_busca.asp" --></div>
        <div class="session-notices"><!-- #include file="includes/dashboard_avisos.asp" --></div>
        <nav class="graduate-jump" aria-label="Se&ccedil;&otilde;es da ficha">
          <a href="#dados_formando">Dados do formando</a>
          <%If VALIDA_SUB7 <> 0 or VENDIDO = 1 Then%><a href="#endereco_formando">Endere&ccedil;o e contato</a><%End If%>
          <a href="#historico_formando">Hist&oacute;rico</a>
          <a href="#movimentacoes_formando">Movimenta&ccedil;&otilde;es</a>
        </nav>
        <%If qtdEnderecos > 1 and VALIDA_SUB7 Then%>
          <div class="subMenu"><a href="#">Visualizar endere&ccedil;o(s) anterior(es)</a></div>
        <%End If%>
        <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
        <section class="dashboard-panel graduate-summary" aria-labelledby="dados_formando">
          <h2 id="dados_formando">Dados do formando</h2>
          <table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
  		<tr>
            <td width="17%"><div align="left"><strong>Nome Formando:</strong></div></td>
            <td width="83%"><%=NOME_FORMANDO%></td>
        </tr>
		<tr>
    		<td><div align="left"><strong>CPF:</strong></div></td>
    		<td><%=CPF%></td>
		</tr>
        <tr>
          <td><div align="left"><strong>Produ&ccedil;&atilde;o: </strong></div></td>
        <%
			  AbrirConexao()
			  
			  
			  
			  
			  Set rsProducao = objConexao.Execute("SELECT * FROM TB_PRODUCAO WHERE COD_PRODUCAO='"&COD_PRODUCAO2&"'")
			  
			  		If not rsProducao.Eof Then
						descricao_producao = COD_PRODUCAO&" - "&rsProducao("DESCRICAO")
					End If
			  
			  rsProducao.Close
			  Set rsProducao = Nothing

			  FecharConexao()
			  %>  
          <td><div align="left"><%=descricao_producao%></div></td>
          
        </tr>

			<tr>
			  <td><div align="left"><strong>Institui&ccedil;&atilde;o de Ensino: </strong></div></td>
			  <td>
			  <%
			  AbrirConexao()
			  
			  Set rsEntidade = objConexao.Execute("SELECT * FROM TB_ENTIDADES WHERE COD_CONTRATO='"&COD_CONTRATO&"'")
			  
			  		If not rsEntidade.Eof Then
						graduateReportInstitution = rsEntidade("DESCRICAO")
                        Response.Write(rsEntidade("DESCRICAO"))
					End If
			  
			  rsEntidade.Close
			  Set rsEntidade = Nothing

			  FecharConexao()
			  %>			  </td>
			</tr>
			<tr>
			  <td><div align="left"><strong>Curso: </strong></div></td>
			  <td><%=NOME_CURSO%></td>
			</tr>
			<tr>
			  <td><div align="left"><strong>Semestre / Ano de Conclus&atilde;o: </strong></div></td>
			  <td><%=SEMESTRE_CONCLUSAO%> / <%=ANO_CONCLUSAO%></td>
			  </tr>
			<tr>
			  <td><div align="left"><strong>&Aacute;lbum: </strong></div></td>
			  <td><%=ALBUM%><%If ENCADERNACOES_EXTRAS = 1 Then%> DUPLO <%End If%></td>
			  </tr>
			  <td><div align="left"><strong>Álbum Digital: </strong></div></td>
				<td>
					<% If ALBUM_DIGITAL = 1 Then %>
						Sim
					<% Else %>
						Não
					<% End If %>
				</td>
			  <tr>
			  <td><div align="left"><strong>N&uacute;mero Pasta <br>(cadastrado pela ger&ecirc;ncia):</strong></div></td>
			  <td><%=NUMERO_PASTA_GERENCIA%></td>
			  </tr>
			  <tr>
			  <td><div align="left"><strong>N&uacute;mero Pasta <br>(cadastrado pelo vendedor):</strong></div>
			  <td><%=NUMERO_PASTA_VENDEDOR%></td>
			  </tr>
			<tr>
			  <td><div align="left"><strong>Inicial</strong>:</div></td>
			  <td><%=INICIAL%></td>
			  </tr>
              <tr>
			  <td><div align="left"><strong>Comiss&atilde;o:</strong></div></td>
			  <td><%=booleano(COMISSAO)%></td>
			  </tr>
			<tr>
			  <td><div align="left"><strong>Brinde: </strong></div></td>
			  <td>
			 <%
								If BRINDE = 1 Then
								AbrirConexao()			
								Dim rsBrinde
										
										sql = "SELECT TB_PRODUTOS.NOME_PRODUTO, Sum( TB_FICHAS_IDENTIFICACAO_BRINDES.QUANTIDADE ) AS QUANTIDADE"&_
																	" FROM TB_PRODUTOS"&_
																	" INNER JOIN TB_FICHAS_IDENTIFICACAO_BRINDES ON TB_PRODUTOS.COD_PRODUTO = TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO"&_
																	" WHERE TB_FICHAS_IDENTIFICACAO_BRINDES.COD_FORMANDO = "&COD_FORMANDO&_
																	" GROUP BY TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO ORDER BY TB_PRODUTOS.NOME_PRODUTO"
										
										
										Set rsBrinde = objConexao.Execute(sql)
					
										If not rsBrinde.Eof then%>
										</span>
										<table width="115" border="1" align="left" 	cellpadding="2" cellspacing="0" style="border-collapse:collapse">
												
											  <%While not rsBrinde.Eof
                                              graduateReportGifts = graduateReportGifts & "<li>" & Server.HTMLEncode(rsBrinde("QUANTIDADE") & " x " & rsBrinde("NOME_PRODUTO")) & "</li>"

											  
											  LINHAS_A_MENOS = LINHAS_A_MENOS+1
											  %>
											  <tr>
												<td width="12%" class="textoMenor textoMenor"><div align="center"><%=rsBrinde("QUANTIDADE")%></div></td>
												<td width="81%" class="textoMenor textoMenor"><div align="center"><%=rsBrinde("NOME_PRODUTO")%></div></td>
											  </tr>
											  <%
											  rsBrinde.MoveNext()
											  Wend
											  %>
							  </table>	
										<%Else%>
										<span class="textoMenor textoMenor">
										<span class="style1">Brinde n&atilde;o encontrado, contate o administrador do sistema!</span>										</span>
										
										<%End if	
											
										rsBrinde.Close()		
								
										Set rsBrinde = Nothing
								FecharConexao()
								Else
													
									Response.Write("N&atilde;o")
								
								End If
								%>			  </td>
			  </tr>
			
			<tr>
			  <td><div align="left"><strong>Produto Agregado:</strong></div></td>
			  <td>
			   <%
													
							
								If PRODUTO_AGREGADO = 1 Then
									AbrirConexao()		
										Dim rsProd
										sql = "SELECT TB_PRODUTOS.NOME_PRODUTO, Sum( TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.QUANTIDADE ) AS QUANTIDADE"&_
																	" FROM TB_PRODUTOS"&_
																	" INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO"&_
																	" WHERE TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO = "&COD_FORMANDO&_
																	" GROUP BY TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO ORDER BY TB_PRODUTOS.NOME_PRODUTO"
																	'Response.Write(sql)
										Set rsProd = objConexao.Execute(sql)
					
										If not rsProd.Eof then%>
											
								<table width="115" border="1" align="left" 	cellpadding="2" cellspacing="0" style="border-collapse:collapse">
											  <%While not rsProd.Eof
                                              graduateReportProducts = graduateReportProducts & "<li>" & Server.HTMLEncode(rsProd("QUANTIDADE") & " x " & rsProd("NOME_PRODUTO")) & "</li>"

											  'Contador de Produtos agregados
											  QTD_PROD_AGREGADOS = QTD_PROD_AGREGADOS + formatNumber(0&rsProd("QUANTIDADE"),0)
											  
											  %>
											  <tr>
												<td width="12%" class="textoMenor textoMenor"><div align="center"><%=rsProd("QUANTIDADE")%></div></td>
												<td width="81%" class="textoMenor textoMenor"><div align="center"><%=rsProd("NOME_PRODUTO")%></div></td>
											  </tr>
											  <%
											  rsProd.MoveNext()
											  Wend
											  %>
							  </table>
										
										<%Else%>
										<span class="textoMenor textoMenor">
										<span class="style1">Produto agregado  n&atilde;o encontrado, contate o administrador do sistema!</span>										</span>
										<%End if	
											
										rsProd.Close()		
										Set rsProd = Nothing
										FecharConexao()
								Else
													
									Response.Write("N&atilde;o")
								
								End If
								
								%>			  </td>
			  </tr>
			
			  
			  
			 <tr bgcolor="whitesmoke">
			  <td><div align="left"><strong>Vendido:</strong></div></td>
			  <td><%=booleano(VENDIDO)%></td>
			  </tr>
			  			  <tr bgcolor="whitesmoke">
			    <td><div align="left"><strong>Sistema novo:</strong> </div></td>
			    <td><%=booleano(SISTEMA_NOVO)%></td>
			    </tr>
			  
			  <%IF SISTEMA_NOVO = 1 Then%>
			   <tr bgcolor="whitesmoke">
			  <td><div align="left"><strong>Impress&atilde;o de etiqueta:</strong></div></td>
			  <td><a class="graduate-label-trigger" aria-haspopup="dialog" href="dip_formandos_cadastro_album_impressao_etiquetas.asp?ia=1&amp;cod_formando=<%=COD_FORMANDO%>">Imprimir etiqueta</a>
			  <%If COD_PRODUCAO = 0 and VALIDA_SUB26 <> 0  Then%> | <a href="javascript: popUp('dip_busca_detalhamento_POPUP_formando_incluir_agregados.asp?cod_form=<%=COD_FORMANDO%>','no',1)">Incluir produtos Agregados</a></td><%End If%>
			  </tr>
			  <%End If%>
			  
			  <%If VALIDA_SUB24 <> 0 and SISTEMA_NOVO = 1 Then%>  
			   <tr bgcolor="whitesmoke">
			  <td><div align="left"><strong>Corre&ccedil;&atilde;o de Formando:</strong></div></td>
<td><a href="dip_busca_detalhamento_formando_correcao_formando.asp?cod_formando=<%=COD_FORMANDO%>">Corre&ccedil;&atilde;o de Formando</a>

</td></tr>
			  <%End If%>
			  <tr>
			  <td><div align="left"><strong> Cr&eacute;dito em valor: </strong></div></td>
			<td>
            <%If VALIDA_SUB28 <> 0 Then%>		  
				<a href="javascript: popUp('dip_busca_detalhamento_formando_POPUP_credito_valor.asp?cod_formando=<%=COD_FORMANDO%>','no',1)">Cr&eacute;dito em Valor</a>&nbsp;
			    <%End If%>
		  <%If CREDITO_EM_VALOR = 0 Then%>
				N&atilde;o h&aacute; cr&eacute;dito em valor para este formando.  
		  <%Else%>	
				Cr&eacute;dito de <%=formatCurrency(CREDITO_EM_VALOR)%> - Motivo: <%=CREDITO_MOTIVO%>
		  <%End If%>
			</td>
			  </tr>
			  

			  <%If VALIDA_SUB25 <> 0 and COMISSAO = 2 and COD_PRODUCAO = 0 Then%>  
			  <tr bgcolor="#F9D0D0">
			  <td><div align="left"><strong>Alterar para comiss&atilde;o :</strong></div></td>
<td><a href="javascript: popUp('dip_producoes_fichas_POPUP_brindes.asp?cod_form=<%=COD_FORMANDO%>&cod_prod=0&altComissao=1','no',1)">Incluir brinde e alterar para comiss&atilde;o</a> </td>
			  </tr>
			  <%End If%>
			  
			  
			  
			<%If VALIDA_SUB9 <> 0 and SISTEMA_NOVO = 1 Then%>  
					<%
					AbrirConexao()
					sql = "SELECT "&_
								"TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.TIPO, "&_
								"TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR, "&_
								"TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_SAIDA, "&_
								"TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_RETORNO, "&_
								"TB_VENDEDORES.NOME_VENDEDOR  "&_
								"FROM TB_VENDEDORES INNER JOIN TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ON "&_
								"TB_VENDEDORES.COD_VENDEDOR = TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR "&_
								"WHERE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_FORMANDO="&COD_FORMANDO&" AND TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.TIPO = 0"
					
					Set rsMovimentacoes2 = objConexao.Execute(sql)
					
					If rsMovimentacoes2.Eof Then 
					%>  
					  <tr bgcolor="#F9D0D0">
						 <td><strong>&Aacute;lbum encontrado:</strong></td>
						 <td>
						 Este formando foi cadastrado, entretanto n&atilde;o foi declarado a exist&ecirc;ncia do &Aacute;lbum, <a href="dip_busca_detalhamento_formando_cadastro_album_finalizacao_cadastro_ficha.asp?cod_formando=<%=COD_FORMANDO%>"><strong>clique aqui</strong></a> para declarar.				 </td>
						 </tr>
					<%
					End If
								
					rsMovimentacoes2.Close
					Set rsMovimentacoes2 = Nothing
					
					FecharConexao()
					%>
			<%End If%>
            
            <%
					AbrirConexao()
					sql = "SELECT TB_ADM_LOG_ALTERACOES.*, TB_ADM_USUARIOS.NOME FROM TB_ADM_USUARIOS INNER JOIN TB_ADM_LOG_ALTERACOES ON TB_ADM_USUARIOS.COD_USUARIO = TB_ADM_LOG_ALTERACOES.COD_USUARIO WHERE TB_ADM_LOG_ALTERACOES.cod_formando = '"&cod_formando&"' order by TB_ADM_LOG_ALTERACOES.data_cadastrado desc"
					
					'Response.Write(sql)
					
					Set rs = objConexao.Execute(sql)
					
					If not rs.Eof Then 
					
					
					%>  
					  <tr bgcolor="#F9D0D0">
						 <td><strong>Log de corre&ccedil;&otilde;es:</strong></td>
						 <td>
                         		<table width="450" border="1" align="left" 	cellpadding="2" cellspacing="0" style="border-collapse:collapse">
                                <tr>
                                <td width="152"><div align="center"><strong>Respons&aacute;vel</strong></div></td>
                                <td width="209"><div align="center"><strong>Descri&ccedil;&atilde;o</strong></div></td>
                                <td width="69"><div align="center"><strong>Data</strong></div></td>
                                </tr>
                                <%While not rs.Eof%>
                                <tr>
                                  <td><%=rs("NOME")%></td>
                                  <td><%=rs("DESCRICAO")%></td>
                                  <td><%=rs("DATA_CADASTRADO")%></td>
                                </tr>
                                <%
								rs.MoveNext()
								Wend
								%>
                              </table>
                                
                                
                          </td>
					</tr>
					<%
					End If
								
					rs.Close
					Set rs = Nothing 
					
					FecharConexao()
					%>
            
			</table>
			
			</section>
            <%If VALIDA_SUB7 <> 0  or VENDIDO = 1 Then%>
            <section class="dashboard-panel graduate-summary" aria-labelledby="endereco_formando">
              <h2 id="endereco_formando">Endere&ccedil;o e contato</h2>
				<table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
  <tr>
            <td width="17%"><div align="left"><strong>Endere&ccedil;o:</strong></div></td>
            <td width="83%"><%=ENDERECO%></td>
            </tr>
			<tr>
			  <td><div align="left"><strong>CEP: </strong></div></td>
			  <td><%=CEP%></td>
			</tr>
			<tr>
			  <td><div align="left"><strong>Bairro:</strong></div></td>
			  <td><%=BAIRRO%>&nbsp;</td>
			</tr>
			<tr>
			  <td><div align="left"><strong>Cidade:</strong></div></td>
			  <td><%=CIDADE%>&nbsp;</td>
			  </tr>
			<tr>
			  <td><div align="left"><strong>Estado:</strong></div></td>
			  <td><%=ESTADO%>&nbsp;</td>
			  </tr>
			<tr>
			  <td><div align="left"><strong>Telefone  Residencial</strong>:</div></td>
			  <td><%=FONE_RES%>&nbsp;</td>
			  </tr>
			<tr>
			  <td><div align="left"><strong>Telefone Celular:</strong></div></td>
			  <td><%=FONE_CEL%>&nbsp;</td>
			  </tr>
			  <tr>
			  <td><div align="left"><strong>E-mail:</strong></div></td>
			  <td><%=EMAIL_FORMANDO%>&nbsp;</td>
			  </tr>
			   
			  </tr>
			</table>
			</section>
            <%End If%>
            <section class="dashboard-panel graduate-panel" aria-labelledby="historico_formando">
            <h2 class="graduate-section-heading" id="historico_formando">Hist&oacute;rico</h2>
           <form action="dip_busca_detalhamento_formando.asp?cod_formando=<%=COD_FORMANDO%>&comando=gravar_msg" style="margin:0px" method="post">
            <div class="graduate-compose">
                <label for="graduateMessage">Anota&ccedil;&atilde;o do formando</label>
                <textarea class="inputTamanho3" id="graduateMessage" name="mensagem" rows="3"></textarea>
                <div class="graduate-compose-actions">
                    <label for="checkbox"><input checked="checked" class="inputRadio" type="checkbox" value="1" name="retornar_contato" id="checkbox" /> Marcar para retornar contato</label>
                    <button type="submit" class="btn btn-primary graduate-submit">Gravar</button>
                </div>
            </div>
            <div class="graduate-scroll" role="region" aria-label="Hist&oacute;rico" tabindex="0">
<table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
                  <tr>
                    <td width="12%"><div align="center"><strong>Gravado por</strong></div></td>
                    <td width="67%"><div align="center"><strong>Mensagem</strong></div></td>
                    <td width="21%"><div align="center">
                        <div align="center"><strong>Data Cadastro</strong></div>
                      </div></td>
                  </tr>

				  <%
				  AbrirConexao()
				  
				  sql = "SELECT * FROM TB_FICHAS_IDENTIFICACAO_HISTORICO WHERE COD_FORMANDO = '"&COD_FORMANDO&"' ORDER BY ID DESC"
				 
				  Set rsMovimentacoesMaster = objConexao.Execute(sql)
				  
				  While not rsMovimentacoesMaster.Eof
				  	
				  		

						  %>
                 		 <tr>
                    <td><%
						Set rs4=objConexao.Execute("SELECT NOME FROM TB_ADM_USUARIOS WHERE COD_USUARIO = "&sqlTexto(rsMovimentacoesMaster("GERADO_POR")))	
							
							If not rs4.Eof Then
							
									Response.Write(Server.HTMLEncode(rs4("nome") & ""))							
							End If
							
						
						rs4.Close
						Set rs4=Nothing
						%></td>
                    <td><div align="left"><%=rsMovimentacoesMaster("MENSAGEM")%></div></td>
                    <td><%=rsMovimentacoesMaster("DATA_CADASTRO")%></td>
                    </tr>
			  
				  <%
				  rsMovimentacoesMaster.MoveNext()
				  Wend
				  
				  rsMovimentacoesMaster.Close
				  Set rsMovimentacoesMaster= Nothing
				  
				  FecharConexao()
				  %>
                </table></div>
                </form>
            </section>
            <section class="dashboard-panel graduate-panel" aria-labelledby="movimentacoes_formando">
            <h2 class="graduate-section-heading" id="movimentacoes_formando">Movimenta&ccedil;&otilde;es</h2>
				<%If SISTEMA_NOVO = 1 Then%>
                <div class="graduate-movements">
<table class="graduate-movement-table" role="table" aria-labelledby="movimentacoes_formando">
    <thead role="rowgroup"><tr role="row">
        <th role="columnheader" scope="col">Item</th>
        <th role="columnheader" scope="col">Vendedor respons&aacute;vel</th>
        <th role="columnheader" scope="col">Data de sa&iacute;da</th>
        <th role="columnheader" scope="col">Data de retorno</th>
        <th role="columnheader" scope="col">Local de armazenagem</th>
    </tr></thead>
    <tbody role="rowgroup">
				  <%
				  AbrirConexao()
				  
				  sql = "SELECT "&_
						"TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.TIPO "&_
						"FROM TB_VENDEDORES INNER JOIN TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ON "&_
						"TB_VENDEDORES.COD_VENDEDOR = TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR "&_
						"WHERE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_FORMANDO="&COD_FORMANDO&" GROUP BY TIPO ORDER BY TIPO ASC, DATA_SAIDA DESC"
				 
				  Set rsMovimentacoesMaster = objConexao.Execute(sql)
                  If rsMovimentacoesMaster.Eof Then
                  %>
                    <tr role="row"><td role="cell" colspan="5">Nenhuma movimenta&ccedil;&atilde;o registrada.</td></tr>
                  <%
                  End If
				  
				  While not rsMovimentacoesMaster.Eof
				  		
						
						 sql = "SELECT "&_
						"TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.TIPO, "&_
						"TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR, "&_
						"TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_MOVIMENTACAO, "&_
						"TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_CAIXA_ARMAZENADO, "&_
						"TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_SAIDA, "&_
						"TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_RETORNO, "&_
						"TB_VENDEDORES.NOME_VENDEDOR  "&_
						"FROM TB_VENDEDORES INNER JOIN TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ON "&_
						"TB_VENDEDORES.COD_VENDEDOR = TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR "&_
						"WHERE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_FORMANDO="&COD_FORMANDO&" AND TIPO ="&rsMovimentacoesMaster("TIPO")&" ORDER BY TIPO ASC, DATA_SAIDA DESC"
						
						 Set rsMovimentacoes = objConexao.Execute(sql)
						
						 iMovimentacoes = 0
						 						 
						 While not rsMovimentacoes.Eof
				  		
 						 iMovimentacoes = iMovimentacoes + 1
						  
						  If iMovimentacoes = 1 Then
						  %>
                    <tr role="row">
                    <td role="cell" class="graduate-movement-item">
                        <%
                     Dim TIPO, movementItemLabel
                  movementItemLabel = "Produto agregado"
                  TIPO = formatNumber(0&rsMovimentacoes("TIPO"),0)
				  
				  If TIPO = 1 Then
                    movementItemLabel = "Ficha"
				  Elseif  TIPO = 0 Then
						
						If ENCADERNACOES_EXTRAS = 1 Then
                            movementItemLabel = "&Aacute;lbum Duplo"
						Else
                            movementItemLabel = "&Aacute;lbum"
						End If
						Elseif TIPO = 2 Then
                            movementItemLabel = "Álbum Digital"
				  Else		
						'Produto agregado, consultar qual &eacute;.
						Dim rsProdAgrePend
						sql = 	"SELECT TB_PRODUTOS.NOME_PRODUTO "&_
						"FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = "&_
						"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO WHERE "&_
						"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO = "&COD_FORMANDO&" AND "&_
						"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.TIPO = "&TIPO&" LIMIT 1"
				
				Set rsProdAgrePend = objConexao.Execute(sql)
				If not rsProdAgrePend.Eof Then
                    movementItemLabel = rsProdAgrePend("NOME_PRODUTO")
				End If
				rsProdAgrePend.Close
				Set rsProdAgrePend = Nothing
						
						
				  End If
					%>
                        <strong><%=movementItemLabel%></strong></td>
					<%
					Dim DATA_SAIDA, DATA_RETORNO
					DATA_SAIDA		= ""&rsMovimentacoes("DATA_SAIDA")
					DATA_RETORNO	= ""&rsMovimentacoes("DATA_RETORNO")
					%>
                    <td role="cell"><span class="graduate-movement-label">Vendedor respons&aacute;vel</span><%=rsMovimentacoes("NOME_VENDEDOR")%></td>
                    <td role="cell"><span class="graduate-movement-label">Data de sa&iacute;da</span><%=DATA_SAIDA%></td>
                    <td role="cell"><span class="graduate-movement-label">Data de retorno</span>
					<%If Len(DATA_RETORNO) > 1 Then%>
							<%=rsMovimentacoes("DATA_RETORNO")%>
					<%Else%>
                        <span class="graduate-movement-tag graduate-movement-pending">Item n&atilde;o devolvido</span>
					<%End If%>					</td>
                    <td role="cell"><span class="graduate-movement-label">Local de armazenagem</span>
                    <%If Len(DATA_RETORNO) > 1 Then%>                    
                    	<%If int(rsMovimentacoes("COD_CAIXA_ARMAZENADO")) = 0 Then%>
                        	Local Provis&oacute;rio de Armazenamento
                       	<%Else%>
                        	<%
							Set rsCaixa = objConexao.Execute("SELECT * FROM TB_LOCAL_ARMAZENAGEM_SALA_CAIXA WHERE COD_CAIXA_ARMAZENADO = "&rsMovimentacoes("COD_CAIXA_ARMAZENADO"))
							
							If not rsCaixa.Eof Then
							%>
                        	Sala <%=rsCaixa("COD_SALA_ARMAZENADO")%> - Coluna <%=rsCaixa("COLUNA")%> - Caixa <%=estoqueNumeroCaixa(rsMovimentacoes("COD_CAIXA_ARMAZENADO"),true)%>
                            	<%If VALIDA_SUB68 <> 0 Then%>
                            	<br />
								<a href="javascript:confirmaVoltarLocalProvisorio('<%=rsMovimentacoes("COD_MOVIMENTACAO")%>');">Voltar para Local provis&oacute;rio de armazenamento</a>
                            	<%
								End If
							End If
							
							rsCaixa.Close
							Set rsCaixa=Nothing
							%>
                            
                        <%End If%>
                    <%Else%>
                    -    
                    <%End If%>
                    </td>
                  </tr>
				          <%Else%>
                         <tr role="row" class="graduate-movement-previous">
                    <td role="cell" class="graduate-movement-item"><strong><%=movementItemLabel%></strong> <span class="graduate-movement-tag">Anterior</span></td>
					<%
					DATA_SAIDA		= ""&rsMovimentacoes("DATA_SAIDA")
					DATA_RETORNO	= ""&rsMovimentacoes("DATA_RETORNO")
					%>
                    <td role="cell"><span class="graduate-movement-label">Vendedor respons&aacute;vel</span><%=rsMovimentacoes("NOME_VENDEDOR")%></td>
                    <td role="cell"><span class="graduate-movement-label">Data de sa&iacute;da</span><%=DATA_SAIDA%></td>
                    <td role="cell"><span class="graduate-movement-label">Data de retorno</span>
					<%If Len(DATA_RETORNO) > 1 Then%>
							<%=rsMovimentacoes("DATA_RETORNO")%>
					<%Else%>
                        <span class="graduate-movement-tag graduate-movement-pending">Item n&atilde;o devolvido</span>
					<%End If%>					</td>
                    <td role="cell"><span class="graduate-movement-label">Local de armazenagem</span>
                    <%If Len(DATA_RETORNO) > 1 Then%>                    
                    	<%If int(rsMovimentacoes("COD_CAIXA_ARMAZENADO")) = 0 Then%>
                        	Local Provis&oacute;rio de Armazenamento
                       	<%Else%>
                        	<%
							Set rsCaixa = objConexao.Execute("SELECT * FROM TB_LOCAL_ARMAZENAGEM_SALA_CAIXA WHERE COD_CAIXA_ARMAZENADO = "&rsMovimentacoes("COD_CAIXA_ARMAZENADO"))
							
							If not rsCaixa.Eof Then
							%>
                        	Sala <%=rsCaixa("COD_SALA_ARMAZENADO")%> - Coluna <%=rsCaixa("COLUNA")%> - Caixa <%=estoqueNumeroCaixa(rsMovimentacoes("COD_CAIXA_ARMAZENADO"),true)%>
                            <%End If
							
							rsCaixa.Close
							Set rsCaixa=Nothing
							%>
                            
                        <%End If%>
                    <%Else%>
                    -    
                    <%End If%>
                    </td>
                  </tr>
                  		  <%End If%>
				 		 <%
				 		 rsMovimentacoes.Movenext()
				  		 Wend
				  
						 rsMovimentacoes.Close
						 Set rsMovimentacoes = Nothing
				  		 %>
				  
				  <%
				  rsMovimentacoesMaster.MoveNext()
				  Wend
				  
				  rsMovimentacoesMaster.Close
				  Set rsMovimentacoesMaster= Nothing
				  
				  FecharConexao()
				  %>
                </tbody>
                </table></div>
                <%Else%>
				   <div align="center">
	  	          <span class="msg"><strong>Aten&ccedil;&atilde;o, as movimenta&ccedil;&otilde;es s&oacute; est&atilde;o dispon&iacute;veis para o sistema novo! </strong>	  	          </span>
				  </div>
	  	          <%End If%>
	  	          
	  	      
            </section>
      </div>
      <%If Len(NOME_FORMANDO & "") > 0 Then%>
      <dialog class="graduate-label-dialog graduate-report-dialog" id="graduateReportDialog" aria-labelledby="graduateReportTitle">
        <header class="graduate-label-dialog-header">
          <h2 id="graduateReportTitle">Pr&eacute;via do resumo</h2>
          <div class="graduate-report-actions">
            <button class="btn btn-primary btn-sm" id="graduateReportPrintButton" type="button">Imprimir / Salvar PDF</button>
            <form method="dialog"><button class="btn btn-outline-secondary btn-sm" autofocus>Fechar</button></form>
          </div>
        </header>
        <div class="graduate-report-preview">
          <%If VALIDA_SUB7 <> 0 Or VENDIDO = 1 Then%>
          <label class="graduate-report-options"><input type="checkbox" id="graduateReportIncludeContact"> Incluir telefone e e-mail</label>
          <%End If%>
          <article class="graduate-report">
            <header class="graduate-report-header">
              <h2>Resumo da ficha do formando</h2>
              <img class="graduate-report-logo" src="images/logoparafundopreto.png" alt="Stillos">
            </header>
            <p class="graduate-report-name"><%=Server.HTMLEncode(NOME_FORMANDO & "")%></p>
            <dl class="graduate-report-grid">
              <div><dt>Ficha</dt><dd><%=Server.HTMLEncode(digito(COD_FORMANDO) & "")%></dd></div>
              <div><dt>Contrato</dt><dd><%=Server.HTMLEncode(COD_CONTRATO & "")%></dd></div>
              <div><dt>CPF</dt><dd><%If Len(Trim(CPF & "")) > 0 Then%><%=Server.HTMLEncode(CPF & "")%><%Else%>&mdash;<%End If%></dd></div>
              <div class="graduate-report-wide"><dt>Institui&ccedil;&atilde;o de ensino</dt><dd><%=Server.HTMLEncode(graduateReportInstitution & "")%></dd></div>
              <div class="graduate-report-wide"><dt>Produ&ccedil;&atilde;o</dt><dd><%=Server.HTMLEncode(descricao_producao & "")%></dd></div>
              <div><dt>Curso</dt><dd><%=Server.HTMLEncode(NOME_CURSO & "")%></dd></div>
              <div><dt>Conclus&atilde;o</dt><dd><%=Server.HTMLEncode(SEMESTRE_CONCLUSAO & " / " & ANO_CONCLUSAO)%></dd></div>
              <div><dt>&Aacute;lbum</dt><dd><%=Server.HTMLEncode(ALBUM & "")%><%If ENCADERNACOES_EXTRAS = 1 Then%> (duplo)<%End If%></dd></div>
              <div><dt>&Aacute;lbum digital</dt><dd><%If ALBUM_DIGITAL = 1 Then%>Sim<%Else%>N&atilde;o<%End If%></dd></div>
              <div><dt>Quantidade de fotos (inicial)</dt><dd><%=Server.HTMLEncode(INICIAL & "")%></dd></div>
              <div><dt>N&uacute;mero da pasta (ger&ecirc;ncia)</dt><dd><%=Server.HTMLEncode(NUMERO_PASTA_GERENCIA & "")%></dd></div>
              <div><dt>N&uacute;mero da pasta (vendedor)</dt><dd><%=Server.HTMLEncode(NUMERO_PASTA_VENDEDOR & "")%></dd></div>
              <div><dt>Comiss&atilde;o</dt><dd><%=booleano(COMISSAO)%></dd></div>
              <div><dt>Brinde</dt><dd><%If Len(graduateReportGifts & "") > 0 Then%><ul class="graduate-report-items"><%=graduateReportGifts%></ul><%ElseIf BRINDE = 1 Then%>Brinde n&atilde;o encontrado<%Else%>N&atilde;o<%End If%></dd></div>
              <div><dt>Produto agregado</dt><dd><%If Len(graduateReportProducts & "") > 0 Then%><ul class="graduate-report-items"><%=graduateReportProducts%></ul><%ElseIf PRODUTO_AGREGADO = 1 Then%>Produto agregado n&atilde;o encontrado<%Else%>N&atilde;o<%End If%></dd></div>
              <div><dt>Vendido</dt><dd><%=booleano(VENDIDO)%></dd></div>
              <div><dt>Emiss&atilde;o</dt><dd><%=Server.HTMLEncode(ArrumaData(Date) & "")%></dd></div>
            </dl>
            <%If VALIDA_SUB7 <> 0 Or VENDIDO = 1 Then%>
            <section class="graduate-report-contact" hidden>
              <h3>Contato</h3>
              <dl class="graduate-report-grid">
                <div><dt>Celular</dt><dd><%=Server.HTMLEncode(FONE_CEL & "")%></dd></div>
                <div><dt>Telefone residencial</dt><dd><%=Server.HTMLEncode(FONE_RES & "")%></dd></div>
                <div class="graduate-report-wide"><dt>E-mail</dt><dd><%=Server.HTMLEncode(EMAIL_FORMANDO & "")%></dd></div>
              </dl>
            </section>
            <%End If%>
          </article>
        </div>
      </dialog>
      <%End If%>
      <dialog class="graduate-label-dialog" id="graduateLabelDialog" aria-labelledby="graduateLabelTitle">
        <header class="graduate-label-dialog-header">
          <h2 id="graduateLabelTitle">Etiquetas</h2>
          <form method="dialog"><button class="btn btn-outline-secondary btn-sm" autofocus>Fechar</button></form>
        </header>
        <p class="graduate-label-loading" id="graduateLabelLoading" role="status" hidden>Carregando&hellip;</p>
        <iframe class="graduate-label-frame" id="graduateLabelFrame" title="Etiquetas" hidden></iframe>
      </dialog>
    </main>
    <footer class="dashboard-footer">&copy; <%=Year(Date)%> Grupo Stillos <span>Sistema administrativo interno</span></footer>
  </div>
</div>
<script>
(function () {
    var toggle = document.getElementById('ppvMenuToggle');
    var menu = document.getElementById('menuPrincipal');
    var previous = menu.querySelector('[aria-current="page"]');
    var current = menu.querySelector('a[href="dip_formandos.asp"]');
    var back = document.getElementById('orderBack');
    back.addEventListener('click', function (event) {
        if (document.referrer && new URL(document.referrer).origin === window.location.origin) {
            event.preventDefault();
            window.history.back();
        }
    });
    if (previous) previous.removeAttribute('aria-current');
    if (current) current.setAttribute('aria-current', 'page');
    toggle.addEventListener('click', function () {
        var expanded = menu.classList.toggle('is-open');
        toggle.setAttribute('aria-expanded', String(expanded));
    });
    document.addEventListener('keydown', function (event) {
        if (event.key === 'Escape' && menu.classList.contains('is-open')) {
            menu.classList.remove('is-open');
            toggle.setAttribute('aria-expanded', 'false');
            toggle.focus();
        }
    });
})();
(function () {
    var dialog = document.getElementById('graduateLabelDialog');
    if (typeof dialog.showModal !== 'function') return;
    var frame = document.getElementById('graduateLabelFrame');
    var title = document.getElementById('graduateLabelTitle');
    var loading = document.getElementById('graduateLabelLoading');
    var queueLink = document.querySelector('.session-notices a[href="adm_etiquetas_agendadas.asp"]');
    var printLink = document.querySelector('.graduate-label-trigger');

    function openLabels(event) {
        if (event.ctrlKey || event.metaKey || event.shiftKey || event.altKey) return;
        event.preventDefault();
        title.textContent = this === queueLink ? 'Etiquetas agendadas' : 'Impress\u00e3o de etiquetas';
        frame.title = title.textContent;
        frame.hidden = false;
        frame.style.visibility = 'hidden';
        loading.hidden = false;
        frame.src = this.href;
        dialog.showModal();
    }

    if (queueLink) {
        queueLink.setAttribute('aria-haspopup', 'dialog');
        queueLink.addEventListener('click', openLabels);
    }
    if (printLink) printLink.addEventListener('click', openLabels);

    frame.addEventListener('load', function () {
        if (!dialog.open) return;
        try {
            var child = frame.contentDocument;
            if (!child || child.URL === 'about:blank') return;
            if (new URL(child.URL).pathname === window.location.pathname) {
                window.location.assign(child.URL);
                return;
            }
            if (child.body.matches('.queue-page, .labels-page')) {
                var styles = child.createElement('style');
                styles.textContent = '.dashboard-sidebar,.dashboard-topbar,.dashboard-footer,.dashboard-heading,.search-panel,.skip-link{display:none!important}.dashboard-main{padding:16px!important}.dashboard-layout{min-height:0!important}';
                child.head.appendChild(styles);
                if (queueLink && child.body.classList.contains('queue-page') && child.querySelector('.queue-panel .empty-state')) {
                    queueLink.parentNode.removeChild(queueLink);
                    queueLink = null;
                }
            }
            child.addEventListener('keydown', function (event) {
                if (event.key === 'Escape' && !child.querySelector('dialog[open]')) dialog.close();
            });
        } finally {
            loading.hidden = true;
            frame.style.visibility = '';
        }
    });
    dialog.addEventListener('close', function () {
        frame.hidden = true;
        frame.removeAttribute('src');
    });
})();
</script>
<div id="graduateReportPrint"></div>
<script>
(function () {
    var trigger = document.getElementById('graduateReportOpen');
    if (!trigger) return;

    var dialog = document.getElementById('graduateReportDialog');
    var article = dialog.querySelector('.graduate-report');
    var contact = dialog.querySelector('.graduate-report-contact');
    var includeContact = document.getElementById('graduateReportIncludeContact');
    var output = document.getElementById('graduateReportPrint');

    trigger.addEventListener('click', function () {
        dialog.showModal();
    });
    if (includeContact) {
        includeContact.addEventListener('change', function () {
            contact.hidden = !includeContact.checked;
        });
    }
    document.getElementById('graduateReportPrintButton').addEventListener('click', function () {
        output.replaceChildren(article.cloneNode(true));
        document.body.classList.add('graduate-printing');
        window.print();
    });
    function clearPrint() {
        document.body.classList.remove('graduate-printing');
        output.replaceChildren();
    }
    window.addEventListener('afterprint', clearPrint);
    dialog.addEventListener('close', function () {
        clearPrint();
        trigger.focus();
    });
}());
</script>
</body>
</html>
