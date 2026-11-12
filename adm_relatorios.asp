<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim VALIDA_SUB10, VALIDA_SUB11, VALIDA_SUB72, VALIDA_SUB13, VALIDA_SUB14, VALIDA_SUB15, VALIDA_SUB16
Dim VALIDA_SUB18, VALIDA_SUB27, VALIDA_SUB29, VALIDA_SUB30, VALIDA_SUB31, VALIDA_SUB41, VALIDA_SUB43
Dim VALIDA_SUB44, VALIDA_SUB45, VALIDA_SUB49, VALIDA_SUB59, VALIDA_SUB77, VALIDA_SUB78, VALIDA_SUB88
'VALIDA_SUB79


'Ranking Vendedores
VALIDA_SUB10 = formatNumber(0&inStr(sub_validacoes_paginas,"#10"),0)

'Ranking Vendedores - Completo 
VALIDA_SUB18 = formatNumber(0&inStr(sub_validacoes_paginas,"#18"),0)

'Produtos Vendidos
VALIDA_SUB11 = formatNumber(0&inStr(sub_validacoes_paginas,"#11"),0)

'Previs?o de Recebimento
VALIDA_SUB72 = formatNumber(0&inStr(sub_validacoes_paginas,"#72"),0)

'Fichas autenticadas pela diretoria
VALIDA_SUB13 = formatNumber(0&inStr(sub_validacoes_paginas,"#13"),0)

'Extrato de Vendedor
VALIDA_SUB15 = formatNumber(0&inStr(sub_validacoes_paginas,"#15"),0)

'Relat?rio de Vendas por Produ??o
VALIDA_SUB16 = formatNumber(0&inStr(sub_validacoes_paginas,"#16"),0)

'Relat?rio Foto Nota - Fechamento para Pagamento
VALIDA_SUB29 = formatNumber(0&inStr(sub_validacoes_paginas,"#29"),0)

'Relat?rio de Sa?da para Fornecedor
VALIDA_SUB30 = formatNumber(0&inStr(sub_validacoes_paginas,"#30"),0)

'Extrato de produ??o com lotes
VALIDA_SUB31 = formatNumber(0&inStr(sub_validacoes_paginas,"#31"),0)

'Relat?rio de Cheques
VALIDA_SUB41 = formatNumber(0&inStr(sub_validacoes_paginas,"#41"),0)

'Relat?rio Mensal de Produtos
VALIDA_SUB43 = formatNumber(0&inStr(sub_validacoes_paginas,"#43"),0)

'Relat?rio de Vendas de Produtos Completo
VALIDA_SUB44 = formatNumber(0&inStr(sub_validacoes_paginas,"#44"),0)

'Relat?rio de Vendas do Grupo Repasse - Ranking
VALIDA_SUB45 = formatNumber(0&inStr(sub_validacoes_paginas,"#45"),0)

'Relat?rio de Formandos por Produ??o
VALIDA_SUB49 = formatNumber(0&inStr(sub_validacoes_paginas,"#49"),0)

'Relat?rio de Produtividade de Funcion?rios por OS
VALIDA_SUB59 = formatNumber(0&inStr(sub_validacoes_paginas,"#59"),0)

'Relat?rio de Despacho
VALIDA_SUB77 = formatNumber(0&inStr(sub_validacoes_paginas,"#77"),0)
'Relat?rio de Entrega
'VALIDA_SUB78 = formatNumber(0&inStr(sub_validacoes_paginas,"#78"),0)

'Relat?rio de Compras
VALIDA_SUB78 = formatNumber(0&inStr(sub_validacoes_paginas,"#78"),0)

'Relat?rio Venda Bolsa e Encaderna??o por Fornecedor
VALIDA_SUB88 = formatNumber(0&inStr(sub_validacoes_paginas,"#88"),0)

%>

<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="windows-1252">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Relat&oacute;rios | Portal Stillos</title>
<link rel="stylesheet" href="css/base.css" media="screen">
<link rel="stylesheet" href="css/vendor/bootstrap.min.css">
<link rel="stylesheet" href="css/dashboard.css">
		<!-- #include virtual="gerencial_foto/files/includes/scripts.asp" -->
		<!-- #include virtual="gerencial_foto/files/includes/functions_local.asp" -->

		<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>

		<script type="application/javascript">

			function trocaContainerRelatorioItensPosseVendedor(campo){
				
				if(campo){

					$("#container_itens_em_posse_vendedor_todos").hide();	
					$("#container_itens_em_posse_vendedor_atrasados").show();

				}else{

					$("#container_itens_em_posse_vendedor_todos").show();	
					$("#container_itens_em_posse_vendedor_atrasados").hide();

				}
					
			}

		</script>

		<script language="JavaScript" type="text/javascript">
			// JavaScript Document

			//adiciona mascara de cnpj
			function MascaraCNPJ(cnpj){

				if(mascaraInteiro(cnpj)==false){
					event.returnValue = false;
				}  

				return formataCampo(cnpj, '00.000.000/0000-00', event);
			}

			//adiciona mascara de cep
			function MascaraCep(cep){
				
				if(mascaraInteiro(cep)==false){
					event.returnValue = false;
				} 

				return formataCampo(cep, '00.000-000', event);
			}

			//adiciona mascara de data
			function MascaraData(data){

				if(mascaraInteiro(data)==false){
					event.returnValue = false;
				}  

				return formataCampo(data, '00/00/0000', event);
			}

			//adiciona mascara ao telefone
			function MascaraTelefone(tel){

				if(mascaraInteiro(tel)==false){
					event.returnValue = false;
				}  

				return formataCampo(tel, '(00) 0000-0000', event);
			}

			//adiciona mascara ao CPF
			function MascaraCPF(cpf){

				if(mascaraInteiro(cpf)==false){
					event.returnValue = false;
				}   

				return formataCampo(cpf, '000.000.000-00', event);
			}

			//valida telefone
			function ValidaTelefone(tel){

				exp = /\(\d{2}\)\ \d{4}\-\d{4}/

				if(!exp.test(tel.value))
					alert('Numero de Telefone Invalido!');
			}

			//valida CEP
			function ValidaCep(cep){

				exp = /\d{2}\.\d{3}\-\d{3}/

				if(!exp.test(cep.value))
					alert('Numero de Cep Invalido!');        
			}

			//valida data
			function validarData(campo){

				var expReg = /^(([0-2]\d|[3][0-1])\/([0]\d|[1][0-2])\/[1-2][0-9]\d{2})$/;
				var msgErro = 'Formato inválido de data.';

				if ((campo.value.match(expReg)) && (campo.value!='')){

					var dia = campo.value.substring(0,2);
					var mes = campo.value.substring(3,5);
					var ano = campo.value.substring(6,10);

					if(mes==4 || mes==6 || mes==9 || mes==11 && dia > 30){
						
						campo.focus();
						alert("Dia incorreto ! O mês especificado contém no máximo 30 dias.");
						return false;

					} else{
				
						if(ano%4!=0 && mes==2 && dia>28){
				
			campo.focus();
			alert("Data incorreta! O mês especificado contém no máximo 28 dias.");
			return false;
						
						} else{
			
			if(ano%4==0 && mes==2 && dia>29){
			campo.focus();
			alert("Data incorreta! O mês especificado contém no máximo 29 dias.");
			return false;
				
			} else{ 
			
			return true;
			}
						}
					}

				} else {
					
					alert(msgErro);
					campo.focus();
					return false;
				}
			}


			//valida o CPF digitado
			function ValidarCPF(Objcpf){

				var cpf = Objcpf.value;
				exp = /\.|\-/g
				cpf = cpf.toString().replace( exp, "" ); 
				var digitoDigitado = eval(cpf.charAt(9)+cpf.charAt(10));
				var soma1=0, soma2=0;
				var vlr =11;
				
				for(i=0;i<9;i++){

					soma1+=eval(cpf.charAt(i)*(vlr-1));
					soma2+=eval(cpf.charAt(i)*vlr);
					vlr--;

				}  

				soma1 = (((soma1*10)%11)==10 ? 0:((soma1*10)%11));
				soma2=(((soma2+(2*soma1))*10)%11);
				
				var digitoGerado=(soma1*10)+soma2;

				if(digitoGerado!=digitoDigitado)    
					alert('CPF Invalido!');        
			}

			//valida numero inteiro com mascara
			function mascaraInteiro(){

				if (event.keyCode < 48 || event.keyCode > 57){
					event.returnValue = false;
					return false;
				}

				return true;
			}

			//valida o CNPJ digitado
			function ValidarCNPJ(ObjCnpj){

				var cnpj = ObjCnpj.value;
				var valida = new Array(6,5,4,3,2,9,8,7,6,5,4,3,2);
				var dig1= new Number;
				var dig2= new Number;
				
				exp = /\.|\-|\//g
				cnpj = cnpj.toString().replace( exp, "" ); 
				var digito = new Number(eval(cnpj.charAt(12)+cnpj.charAt(13)));
					
				for(i = 0; i<valida.length; i++){

					dig1 += (i>0? (cnpj.charAt(i-1)*valida[i]):0);    
					dig2 += cnpj.charAt(i)*valida[i]; 

				}
				
				dig1 = (((dig1%11)<2)? 0:(11-(dig1%11)));
				dig2 = (((dig2%11)<2)? 0:(11-(dig2%11)));
				
				if(((dig1*10)+dig2) != digito)    
					alert('CNPJ Invalido!');
					
			}

			//formata de forma generica os campos
			function formataCampo(campo, Mascara, evento) { 

				var boleanoMascara; 				
				var Digitato = evento.keyCode;
				exp = /\-|\.|\/|\(|\)| /g
				campoSoNumeros = campo.value.toString().replace( exp, "" ); 
			
				var posicaoCampo = 0;     
				var NovoValorCampo="";
				var TamanhoMascara = campoSoNumeros.length;; 
				
				if (Digitato != 8) { // backspace 

					for(i=0; i<= TamanhoMascara; i++) {

						boleanoMascara  = ((Mascara.charAt(i) == "-") || (Mascara.charAt(i) == ".")
			|| (Mascara.charAt(i) == "/")) 

						boleanoMascara  = boleanoMascara || ((Mascara.charAt(i) == "(") 
			|| (Mascara.charAt(i) == ")") || (Mascara.charAt(i) == " ")) 

						if (boleanoMascara) { 

			NovoValorCampo += Mascara.charAt(i); 
			TamanhoMascara++;

						}else { 

			NovoValorCampo += campoSoNumeros.charAt(posicaoCampo); 
			posicaoCampo++; 

						}            
					}   

					campo.value = NovoValorCampo;
					return true; 

				}else { 

					return true; 

				}
			}
		</script> 

		<script language="javascript">

			//Formata n?mero tipo moeda usando o evento onKeyDown
			function Limpar(valor, validos) {
				
				// retira caracteres invalidos da string
				var result = "";
				var aux;

				for (var i=0; i < valor.length; i++) {
					
					ux = validos.indexOf(valor.substring(i, i+1));
					
					if (aux>=0) {
						result += aux;
					}
				}

				return result;
			}

			function Formata(campo,tammax,teclapres,decimal) {
				
				var tecla = teclapres.keyCode;
				vr = Limpar(campo.value,"0123456789");
				tam = vr.length;
				dec=decimal

				if (tam < tammax && tecla != 8){
					tam = vr.length + 1;
				}

				if (tecla == 8 ){
					tam = tam - 1;
				}

				if (tecla == 8 || tecla >= 48 && tecla <= 57 || tecla >= 96 && tecla <= 105){

					if (tam <= dec){
						campo.value = vr;
					}

					if ((tam > dec) && (tam <= 5)){
						campo.value = vr.substr( 0, tam - 2 ) + "," + vr.substr( tam - dec, tam );
					}

					if ((tam >= 6) && (tam <= 8)){
						campo.value = vr.substr( 0, tam - 5 ) + "." + vr.substr( tam - 5, 3 ) + "," + vr.substr( tam - dec, tam ); 
					}

					if ((tam >= 9) && (tam <= 11)){
						campo.value = vr.substr( 0, tam - 8 ) + "." + vr.substr( tam - 8, 3 ) + "." + vr.substr( tam - 5, 3 ) + "," + vr.substr( tam - dec, tam );
					}

					if ((tam >= 12) && (tam <= 14)){
						campo.value = vr.substr( 0, tam - 11 ) + "." + vr.substr( tam - 11, 3 ) + "." + vr.substr( tam - 8, 3 ) + "." + vr.substr( tam - 5, 3 ) + "," + vr.substr( tam - dec, tam );
					}
					if ((tam >= 15) && (tam <= 17)){
						campo.value = vr.substr( 0, tam - 14 ) + "." + vr.substr( tam - 14, 3 ) + "." + vr.substr( tam - 11, 3 ) + "." + vr.substr( tam - 8, 3 ) + "." + vr.substr( tam - 5, 3 ) + "," + vr.substr( tam - 2, tam );
					}
				} 
			}

		</script>

		<script language="javascript">
		
			function mOvr(src, cor){

				if (!src.contains(event.fromElement)){
					src.style.cursor = 'hand';
					src.bgColor = cor;
				}
			}
			
			function mOut(src, cor){

				if (!src.contains(event.toElement)){					
					src.style.cursor = 'default';
					src.bgColor = cor;				
				}
			}

		</script>
		
		<script type="text/javascript" language="javascript">


			function abrir(theURL,winName,features){ 
				window.open(theURL,winName,features);
			}

			
			function abre(url){    
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

			function Verificar(){
				var tecla=window.event.keyCode;
				
				if (tecla==17){
					buscaIni();
				}
			}


			function MM_jumpMenu(targ,selObj,restore){ //v3.0
			eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
			if (restore) selObj.selectedIndex=0;
			}
			

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
		
			function myOtherFunction(){}
			
			function pickDate(buttonObj,inputObject){

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

			function getDateFromCalendar(inputArray){
				
				var references = calendarObjForForm.getHtmlElementReferences(); // Get back reference to form field.
				references.myDate.value = inputArray.day + '/' + inputArray.month + '/' + inputArray.year
				//references.myDate.value = inputArray.year + '-' + inputArray.month + '-' + inputArray.day + ' ' + inputArray.hour + ':' + inputArray.minute;
				calendarObjForForm.hide();	
			
			}	

			function validaFormExtrato(){
				
				d = document.EXTRATO_MOVIMENTACOES

				if(d.COD_VENDEDOR.value == 0){
					
					alert('Selecione o vendedor para o Extrato de Vendedor - Movimentações')
					d.COD_VENDEDOR.focus();
					return false;

				} else {
					return true;
				}
			}
			
            function validaForm7() {
                var production = document.forms.EXTRATO_APROVEITAMENTO.elements.COD_PRODUCAO;
                if (production.value !== "0" && production.value !== "") return true;
                var notice = document.getElementById("report-production-notice");
                if (!notice.open) notice.showModal();
                return false;
            }
		
		</script>


<style>
.dashboard-search .input-filtro { width:auto; min-width:0; }
.reports-search { padding:16px 20px; margin-bottom:16px; }
.reports-search label { display:block; font-weight:650; margin-bottom:6px; }
.reports-search input { max-width:420px; }
.reports-status:empty { display:none; }
.reports-catalog { display:grid; gap:20px; }
.reports-card { min-width:0; padding:24px; border:1px solid #e4dfeb; border-radius:12px; background:#fff; }
.reports-catalog .reports-card > h2 { margin:0 !important; padding-bottom:18px; color:#242738; font-size:1.16rem; }
.reports-card p { color:#53566b; margin:0 0 12px; }
.reports-vendor { width:min(100%,400px); min-width:0; }
.reports-production { flex:0 1 560px; min-width:0; max-width:100%; }
.reports-card .reports-fields select.form-select { width:100%; min-width:0; padding-right:2.5rem; background-color:#fff; }
.reports-production select { text-overflow:ellipsis; }
.reports-production [hidden] { display:none !important; }
.reports-picker > summary { padding:9px 12px; border:1px solid #d9d5e3; border-radius:8px; background:#fff; cursor:pointer; overflow-wrap:anywhere; }
.reports-picker > summary:focus-visible,.reports-picker-list button:focus-visible { outline:2px solid #b20b61; outline-offset:-2px; }
.reports-picker-panel { margin-top:6px; padding:10px; border:1px solid #e4dfeb; border-radius:8px; background:#fff; }
.reports-picker-panel .form-control { width:100%; min-width:0; }
.reports-picker-list { max-height:min(18rem,42vh); overflow:auto; margin:8px 0 0; padding:0; list-style:none; }
.reports-picker-list button { display:block; width:100%; padding:9px 8px; border:0; border-radius:5px; background:#fff; color:#242738; text-align:left; font:inherit; white-space:normal; overflow-wrap:anywhere; }
.reports-picker-list button:hover,.reports-picker-list button[aria-pressed="true"] { background:#f5edf3; }
.reports-picker-panel .reports-picker-status { margin:8px 0 0; }
.reports-picker-status:empty { display:none; }
.reports-field-name { display:block; font-weight:600; margin-bottom:8px; }
.reports-fields .reports-check { display:flex; align-items:center; gap:8px; min-height:38px; font-weight:500; }
.reports-quick { margin-top:14px; padding-top:12px; border-top:1px solid #eeeaf2; }
.report-notice { width:min(420px,calc(100% - 32px)); padding:22px; border:1px solid #e4dfeb; border-radius:12px; color:#242738; }
.report-notice::backdrop { background:#24273899; }
.report-notice h2 { margin:0 0 8px; color:#242738; }
.report-notice form { display:flex; justify-content:flex-end; margin-top:16px; }
.reports-fields { display:flex; gap:14px 16px; flex-wrap:wrap; align-items:end; }
.reports-fields label { display:grid; gap:8px; min-width:130px; font-weight:600; }
.reports-compact label { flex:0 1 152px; min-width:0; }
.reports-compact :is(.reports-date,.form-select) { width:100%; min-height:40px; }
#RELATORIO_RETIRADA input[type="date"] { appearance:auto; }
#RELATORIO_RETIRADA input[type="date"]::-webkit-calendar-picker-indicator {
    display:block;
    visibility:visible;
    opacity:1;
    cursor:pointer;
}
.reports-compact .form-select { padding-right:2.5rem; }
@media (max-width:600px) {
    .reports-compact label { flex:1 1 130px; }
}

.reports-filter-scroll { overflow-x:auto; }
.reports-filter-scroll:focus-visible { outline:2px solid #b20b61; outline-offset:2px; }
.reports-filter-scroll > table,
.reports-filter-scroll > form > table { width:100% !important; background:none !important; border:0 !important; margin:0 !important; }
.reports-filter-scroll td { padding:4px; text-align:left; }
.reports-filter-scroll td > div[align] { text-align:left; }
.reports-filter-scroll td:has(.btn) > div[align] { text-align:right; }
.reports-card :where(select,input:not([type=hidden],[type=radio],[type=checkbox],[type=button])) {
    min-height:38px; max-width:100%; padding:6px 8px; border:1px solid #d9d5e3;
    border-radius:7px; background-color:#fff; color:#242738; font:inherit;
}
.reports-card .campoData { width:114px; }
.reports-card :is(input,select):focus { outline:2px solid #b20b61; outline-offset:1px; }
@media(max-width:600px) {
    .reports-card { padding:18px; }
    .reports-catalog .reports-card > h2 { padding-bottom:16px; }
    .reports-fields label { flex:1 1 125px; }
    .reports-fields .btn { width:100%; }
}

.report-preview { width:min(1640px,calc(100% - 32px)); height:calc(100dvh - 32px); max-width:none; max-height:none; padding:0; border:1px solid #e4dfeb; border-radius:16px; color:#242738; }
.report-preview[open] { display:flex; flex-direction:column; }
.report-preview::backdrop { background:#24273899; }
.report-preview > header { display:flex; flex-wrap:wrap; align-items:center; justify-content:space-between; gap:8px; padding:12px 16px; border-bottom:1px solid #e4dfeb; }
.report-preview > header h2 { margin:0; color:#242738; font-size:1.15rem; }
.report-preview-actions .btn-primary { --bs-btn-disabled-bg:#ede8f0; --bs-btn-disabled-border-color:#ede8f0; --bs-btn-disabled-color:#595368; --bs-btn-disabled-opacity:1; }
.report-preview-actions { display:flex; flex-wrap:wrap; gap:8px; }
.report-preview iframe { width:100%; min-height:0; flex:1; border:0; background:#f6f7fb; }
.report-preview-status { padding:16px; margin:0; color:#53566b; }
.report-preview-status[aria-busy="true"]::before { content:""; display:inline-block; width:1em; height:1em; margin-right:.55em; vertical-align:-.12em; border:2px solid #e4dfeb; border-top-color:#b20b61; border-radius:50%; animation:report-loading .75s linear infinite; }
@keyframes report-loading { to { transform:rotate(360deg); } }
@media(prefers-reduced-motion:reduce) { .report-preview-status::before { animation:none !important; } }
body:has(.report-preview[open]) { overflow:hidden; }
@media(max-width:600px) {
  .report-preview { width:calc(100% - 12px); height:calc(100dvh - 12px); border-radius:10px; }
  .report-preview > header { padding:10px; }
}
</style>
</head>
<body>
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
      <div class="topbar-title">Portal Stillos <span>/ Relat&oacute;rios</span></div>
      <div class="topbar-actions"><a class="user-name" href="adm_log_usuarios.asp"><%=Server.HTMLEncode(Session("NOME") & "")%></a><a href="manual_sistema.asp" target="_blank" rel="noopener">Manual</a><a class="btn btn-outline-secondary btn-sm" href="adm_sair.asp">Sair</a></div>
    </header>
    <main class="dashboard-main" id="conteudo">

      <header class="dashboard-heading">
        <div><h1>Relat&oacute;rios</h1><p>Escolha um relat&oacute;rio e preencha os filtros para consultar os resultados.</p></div>
        <a class="btn btn-outline-secondary" href="adm_index.asp">Voltar ao painel</a>
      </header>
      <div class="search-panel"><strong>Busca no portal</strong>
      <!-- #include file="includes/dashboard_busca.asp" -->
      </div>
      <div class="session-notices">
      <!-- #include file="includes/dashboard_avisos.asp" -->
      </div>
      <%If Len(Request.QueryString("msg")) > 0 Then%>
      <div class="alert alert-info" role="status"><%=Server.HTMLEncode(Request.QueryString("msg") & "")%></div>
      <%End If%>
      <div class="dashboard-panel reports-search" id="reports-search-panel" hidden>
        <label for="reports-search">Buscar relat&oacute;rio</label>
        <input class="form-control" id="reports-search" type="search" placeholder="Digite o nome do relat&oacute;rio">
      </div>
      <p class="reports-status" id="reports-status" role="status"></p>
      <div class="reports-catalog">
<%
			Dim dataInicialForm, dataFinalForm, qtdDiasMesArray(12), AnoBi
			dataInicialForm = Request.QueryString("DATA_INI")
			dataFinalForm = Request.QueryString("DATA_FIM")
					
			If dataInicialForm = Empty or dataFinalForm = Empty then

			AnoBi = Year(Now) Mod 4

			IF AnoBi = 0 then

			AnoBi = 29

			else

			AnoBi = 28

			End If
			
			qtdDiasMesArray(1) = 31'Jan
			qtdDiasMesArray(2) = AnoBi
			qtdDiasMesArray(3) = 31'Mar
			qtdDiasMesArray(4) = 30'Abr
			qtdDiasMesArray(5) = 31'Mai
			qtdDiasMesArray(6) = 30'Jun
			qtdDiasMesArray(7) = 31'Jul
			qtdDiasMesArray(8) = 31'Ago
			qtdDiasMesArray(9) = 30'Set
			qtdDiasMesArray(10) = 31'Out
			qtdDiasMesArray(11) = 30'Nov
			qtdDiasMesArray(12) = 31'Dez
			
			dataInicialForm = "01/"&Month(Now)&"/"&Year(Now)
			dataFinalForm = qtdDiasMesArray(Month(Now))&"/"&Month(Now)&"/"&Year(Now)
			
			End If
			
			%>

			<!--
			<%If VALIDA_SUB10 <> 0 Then%>
			<h3 align="center">Ranking de Vendedores</h3>
			<form style="margin:0px" name="FILTRO" action="adm_relatorios_ranking_vendedores.asp">
			<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
			<tr>
			<td width="12%"><div align="right"><strong>M&ecirc;s / Ano:</strong> </div></td>
			<td width="88%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
			<tr>
			<td width="18%"><div align="center">
			<select name="mesUtilizar" style="font-size:13px">
			<option value="01" <%If Month(Now) = 1 then%> selected="selected"<%end if%>>Janeiro</option>
			<option value="02" <%If Month(Now) = 2 then%> selected="selected"<%end if%>>Fevereiro</option>
			<option value="03" <%If Month(Now) = 3 then%> selected="selected"<%end if%>>Mar&ccedil;o</option>
			<option value="04" <%If Month(Now) = 4 then%> selected="selected"<%end if%>>Abril</option>
			<option value="05" <%If Month(Now) = 5 then%> selected="selected"<%end if%>>Maio</option>
			<option value="06" <%If Month(Now) = 6 then%> selected="selected"<%end if%>>Junho</option>
			<option value="07" <%If Month(Now) = 7 then%> selected="selected"<%end if%>>Julho</option>
			<option value="08" <%If Month(Now) = 8 then%> selected="selected"<%end if%>>Agosto</option>
			<option value="09" <%If Month(Now) = 9 then%> selected="selected"<%end if%>>Setembro</option>
			<option value="10" <%If Month(Now) = 10 then%> selected="selected"<%end if%>>Outubro</option>
			<option value="11" <%If Month(Now) = 11 then%> selected="selected"<%end if%>>Novembro</option>
			<option value="12" <%If Month(Now) = 12 then%> selected="selected"<%end if%>>Dezembro</option>
			</select>
			</div></td>
			<td width="8%"><div align="center">Ano</div></td>
			<td width="11%"><div align="center">
			<select name="anoUtilizar" style="font-size:13px">
			<option value="<%=Year(Now)-2%>"><%=Year(Now)-2%></option>
			<option value="<%=Year(Now)-1%>"><%=Year(Now)-1%></option>
			<option selected="selected" value="<%=Year(Now)%>"><%=Year(Now)%></option>
			<option value="<%=Year(Now)+1%>"><%=Year(Now)+1%></option>
			</select>
			</div></td>
			<td width="11%">&nbsp;</td>
			<td width="52%"><div align="right">
			<button class="btn btn-primary" type="submit" name="filtro22" value=" ">Abrir relat&oacute;rio</button>
			</div></td>
			</tr>
			</table></td>
			</tr>
			</table>
			</form>
			<%End If%>
			-->
				
			<%If VALIDA_SUB18 <> 0 Then%>
			<section class="reports-card" aria-labelledby="report-1">
<h2 id="report-1">Ranking de representantes comerciais</h2>
<p>Compare o desempenho dos representantes pela data da venda.</p><form method="get" name="RELATORIO_RANKING_VENDEDORES" action="adm_relatorios_ranking_vendedores_completo.asp" data-report-preview style="margin:0px">
<div class="reports-fields">
  <label for="reports-ranking-start">Data inicial
    <input class="form-control" id="reports-ranking-start" type="date" name="DATA_VENDA_INICIO" required>
  </label>
  <label for="reports-ranking-end">Data final
    <input class="form-control" id="reports-ranking-end" type="date" name="DATA_VENDA_FIM" required>
  </label>
  <button class="btn btn-primary" type="submit" name="filtro3" value=" ">Abrir relat&oacute;rio</button>
</div>
</form>
</section>
			<%End If%>					
				
			<%If VALIDA_SUB72 <> 0 Then%>
			
			<section class="reports-card" aria-labelledby="report-2">
<h2 id="report-2">Vendas mensais por forma de pagamento</h2>
<form style="margin:0px" name="FILTRO" action="adm_relatorios_previsao_recebimento.asp" data-report-preview>
<div class="reports-fields"><label>M&ecirc;s inicial<select name="MES_INICIO" style="font-size:13px">
			
			<option value="01" <%If Month(Now) = 1 then%> selected="selected"<%end if%>>Janeiro</option>
			
			<option value="02" <%If Month(Now) = 2 then%> selected="selected"<%end if%>>Fevereiro</option>
			
			<option value="03" <%If Month(Now) = 3 then%> selected="selected"<%end if%>>Mar&ccedil;o</option>
			
			<option value="04" <%If Month(Now) = 4 then%> selected="selected"<%end if%>>Abril</option>
			
			<option value="05" <%If Month(Now) = 5 then%> selected="selected"<%end if%>>Maio</option>
			
			<option value="06" <%If Month(Now) = 6 then%> selected="selected"<%end if%>>Junho</option>
			
			<option value="07" <%If Month(Now) = 7 then%> selected="selected"<%end if%>>Julho</option>
			
			<option value="08" <%If Month(Now) = 8 then%> selected="selected"<%end if%>>Agosto</option>
			
			<option value="09" <%If Month(Now) = 9 then%> selected="selected"<%end if%>>Setembro</option>
			
			<option value="10" <%If Month(Now) = 10 then%> selected="selected"<%end if%>>Outubro</option>
			
			<option value="11" <%If Month(Now) = 11 then%> selected="selected"<%end if%>>Novembro</option>
			
			<option value="12"<%If Month(Now) = 12 then%> selected="selected"<%end if%>>Dezembro</option>

			</select></label><label>Ano inicial<select name="ANO_INICIO" style="font-size:13px">
			
			<option value="<%=Year(Now)-3%>"><%=Year(Now)-3%></option>
			
			<option value="<%=Year(Now)-2%>"><%=Year(Now)-2%></option>
			
			<option value="<%=Year(Now)-1%>" selected="selected"><%=Year(Now)-1%></option>
			
			<option value="<%=Year(Now)%>"><%=Year(Now)%></option>
			
			</select></label><label>M&ecirc;s final<select name="MES_FIM" style="font-size:13px">
			
			<option value="01" <%If Month(Now) = 1 then%> selected="selected"<%end if%>>Janeiro</option>
			
			<option value="02" <%If Month(Now) = 2 then%> selected="selected"<%end if%>>Fevereiro</option>
			
			<option value="03" <%If Month(Now) = 3 then%> selected="selected"<%end if%>>Mar&ccedil;o</option>
			
			<option value="04" <%If Month(Now) = 4 then%> selected="selected"<%end if%>>Abril</option>
			
			<option value="05" <%If Month(Now) = 5 then%> selected="selected"<%end if%>>Maio</option>
			
			<option value="06" <%If Month(Now) = 6 then%> selected="selected"<%end if%>>Junho</option>
			
			<option value="07" <%If Month(Now) = 7 then%> selected="selected"<%end if%>>Julho</option>
			
			<option value="08" <%If Month(Now) = 8 then%> selected="selected"<%end if%>>Agosto</option>
			
			<option value="09" <%If Month(Now) = 9 then%> selected="selected"<%end if%>>Setembro</option>
			
			<option value="10" <%If Month(Now) = 10 then%> selected="selected"<%end if%>>Outubro</option>
			
			<option value="11" <%If Month(Now) = 11 then%> selected="selected"<%end if%>>Novembro</option>
			
			<option value="12" <%If Month(Now) = 12 then%> selected="selected"<%end if%>>Dezembro</option>
			
			</select></label><label>Ano final<select name="ANO_FIM" style="font-size:13px">
			
			<option value="<%=Year(Now)-3%>"><%=Year(Now)-3%></option>
			
			<option value="<%=Year(Now)-2%>"><%=Year(Now)-2%></option>
			
			<option value="<%=Year(Now)-1%>"><%=Year(Now)-1%></option>
			
			<option value="<%=Year(Now)%>" selected="selected"><%=Year(Now)%></option>

			</select></label><button class="btn btn-primary" type="submit" name="filtro2" value=" ">Abrir relat&oacute;rio</button></div>
</form>
</section>

			<%End If%>			
				
				
			<%If VALIDA_SUB29 <> 0 Then%>
			
			<section class="reports-card" aria-labelledby="report-3">
<h2 id="report-3">Relat&oacute;rio de notas</h2>
<form style="margin:0px" action="adm_relatorios_n_mensal.asp" data-report-preview target="_blank">
<div class="reports-fields"><label>M&ecirc;s<select name="MES_INICIO" style="font-size:13px">
			
			<option value="01" <%If Month(Now) = 1 then%> selected="selected"<%end if%>>Janeiro</option>
			
			<option value="02" <%If Month(Now) = 2 then%> selected="selected"<%end if%>>Fevereiro</option>
			
			<option value="03" <%If Month(Now) = 3 then%> selected="selected"<%end if%>>Mar&ccedil;o</option>
			
			<option value="04" <%If Month(Now) = 4 then%> selected="selected"<%end if%>>Abril</option>
			
			<option value="05" <%If Month(Now) = 5 then%> selected="selected"<%end if%>>Maio</option>
			
			<option value="06" <%If Month(Now) = 6 then%> selected="selected"<%end if%>>Junho</option>
			
			<option value="07" <%If Month(Now) = 7 then%> selected="selected"<%end if%>>Julho</option>
			
			<option value="08" <%If Month(Now) = 8 then%> selected="selected"<%end if%>>Agosto</option>
			
			<option value="09" <%If Month(Now) = 9 then%> selected="selected"<%end if%>>Setembro</option>
			
			<option value="10" <%If Month(Now) = 10 then%> selected="selected"<%end if%>>Outubro</option>
			
			<option value="11" <%If Month(Now) = 11 then%> selected="selected"<%end if%>>Novembro</option>
			
			<option value="12"<%If Month(Now) = 12 then%> selected="selected"<%end if%>>Dezembro</option>
			
			</select></label><label>Ano<select name="ANO_INICIO" style="font-size:13px">
			
			<option value="<%=Year(Now)-3%>"><%=Year(Now)-3%></option>
			
			<option value="<%=Year(Now)-2%>"><%=Year(Now)-2%></option>
			
			<option value="<%=Year(Now)-1%>"><%=Year(Now)-1%></option>
			
			<option selected="selected"  value="<%=Year(Now)%>"><%=Year(Now)%></option>
			
			<option value="<%=Year(Now)+2%>"><%=Year(Now)+1%></option>
			
			<option value="<%=Year(Now)+2%>"><%=Year(Now)+2%></option>
			</select></label><button class="btn btn-primary" type="submit" name="filtro2" value=" ">Abrir relat&oacute;rio</button></div>
</form>
</section>

			<%End If%>	
				
				
			<%If VALIDA_SUB13 <> 0 Then%>

			<section class="reports-card" aria-labelledby="report-4">
<h2 id="report-4">Fichas autenticadas pela Diretoria</h2>

<form method="get" action="adm_relatorios_fichas_autenticadas_diretoria.asp" data-report-preview>
  <input type="hidden" name="filtro2" value="1">
  <%Dim fichasDataInicio, fichasDataFim
    fichasDataFim = Date()
    fichasDataInicio = DateAdd("m", -1, fichasDataFim)%>
  <div class="reports-fields" style="margin-top:12px">
    <label>Data inicial
      <input class="form-control" type="date" name="DATA_INI" required value="<%=Year(fichasDataInicio) & "-" & Right("0" & Month(fichasDataInicio), 2) & "-" & Right("0" & Day(fichasDataInicio), 2)%>">
    </label>
    <label>Data final
      <input class="form-control" type="date" name="DATA_FIM" required value="<%=Year(fichasDataFim) & "-" & Right("0" & Month(fichasDataFim), 2) & "-" & Right("0" & Day(fichasDataFim), 2)%>">
    </label>
    <button class="btn btn-primary" type="submit">Consultar per&iacute;odo</button>
  </div>
</form>
<form class="reports-quick" method="get" action="adm_relatorios_fichas_autenticadas_diretoria.asp" data-report-preview>
  <input type="hidden" name="filtro1" value="1">
  <button class="btn btn-outline-secondary" type="submit">Ver &uacute;ltimos 100 registros</button>
</form>
</section>

			<%end if%>
				
			<%
			dataFinalForm 		=	dataInicialForm
			dataInicialForm		=	DateAdd("m",-1,dataFinalForm)
			%>

			<%If VALIDA_SUB15 <> 0 Then%>

			<section class="reports-card" aria-labelledby="report-5">
<h2 id="report-5">&Aacute;lbuns em posse do vendedor</h2>
<form method="get" name="EXTRATO_BLOQUEADOS" action="adm_relatorios_extrato_vendedores_atrasados_por_vendedor_html3.asp" data-report-preview>
  <div class="reports-fields">
    <div class="reports-vendor">
      <span class="reports-field-name">Vendedor</span>
<%AbrirConexao()%>                
			
			<div id="container_itens_em_posse_vendedor_todos">
			
			<%
			Set rsVendedor = objConexao.Execute("SELECT COD_VENDEDOR, NOME_VENDEDOR FROM VW_MOVIMENTACOES_COMPLETA WHERE TIPO = 0 AND DATA_SAIDA IS NOT NULL AND DATA_RETORNO IS NULL GROUP BY COD_VENDEDOR ORDER BY NOME_VENDEDOR")
			
			If not rsVendedor.Eof Then%>
			
			<select name="COD_VENDEDOR" class="form-select" aria-label="Vendedor">
			
			<% 
			While not rsVendedor.Eof
			%>

			<option value="<%=rsVendedor("COD_VENDEDOR")%>"><%=Server.HTMLEncode(rsVendedor("NOME_VENDEDOR") & "")%></option>
			
			<%
			rsVendedor.MoveNext()
			Wend
			%>
			
			</select>

			<%Else%>

			Não há vendedores com itens em posse.

			<%End If
			
			rsVendedor.Close()
			Set rsVendedor = Nothing
			%>

			</div> 
			
			<div id="container_itens_em_posse_vendedor_atrasados" style="display:none">
			
			<%
			Set rsVendedor = objConexao.Execute("SELECT COD_VENDEDOR, NOME_VENDEDOR FROM VW_MOVIMENTACOES_COMPLETA WHERE TIPO = 0 AND DATA_SAIDA IS NOT NULL AND DATA_RETORNO IS NULL AND DIAS_ATRASADOS >= 0 GROUP BY COD_VENDEDOR ORDER BY NOME_VENDEDOR")
			
			If not rsVendedor.Eof Then%>
			
			<select name="COD_VENDEDOR_BLOQUEADO" class="form-select" aria-label="Vendedor">
			
			<%
			While not rsVendedor.Eof
			%>
			
			<option value="<%=rsVendedor("COD_VENDEDOR")%>"><%=Server.HTMLEncode(rsVendedor("NOME_VENDEDOR") & "")%></option>
			
			<%
			rsVendedor.MoveNext()
			Wend
			%>
			
			</select>
			
			<%Else%>
			
			N&atilde;o h&aacute; vendedores com itens em atraso.
			
			<%End If
			rsVendedor.Close()
			Set rsVendedor = Nothing
			%>

			</div>

			<%
			FecharConexao()
			%>
			
			
    </div>
    <label class="reports-check"><input type="checkbox" onclick="trocaContainerRelatorioItensPosseVendedor(this.checked)" name="SOMENTE_VENDEDORES_EM_ATRASO" value="S"> Somente &aacute;lbuns em atraso</label>
    <button class="btn btn-primary" type="submit" name="filtro216" value=" ">Abrir relat&oacute;rio</button>
  </div>
</form>
</section>

			<%end if%>
				
			<%If VALIDA_SUB31 <> 0 Then%>

			<section class="reports-card" aria-labelledby="report-6">
<h2 id="report-6">Extrato de aproveitamento de Produ&ccedil;&atilde;o</h2>
<form name="EXTRATO_APROVEITAMENTO" action="adm_relatorios_extrato_por_producao_e_vendedor.asp" method="get" data-report-preview style="margin:0px" onsubmit="return validaForm7()">
  <div class="reports-fields"><div class="reports-production">
<label for="COD_PRODUCAO" id="production-native-label">Produ&ccedil;&atilde;o</label>
<select name="COD_PRODUCAO" id="COD_PRODUCAO" class="form-select">
			
			<option selected="selected" value="0">Selecione uma produ&ccedil;&atilde;o</option>

			<%
			AbrirConexao()
			Set rsProducoes = objConexao.Execute("SELECT COD_PRODUCAO, DESCRICAO FROM TB_PRODUCAO WHERE PRODUCAO_FINALIZADA = 1 ORDER BY cod_producao desc")
			While not rsProducoes.Eof
			%>
			
			<option value="<%=rsProducoes("COD_PRODUCAO")%>"><%=rsProducoes("cod_producao")%> - <%=Server.HTMLEncode(rsProducoes("DESCRICAO") & "")%></option>

			<%
			rsProducoes.MoveNext()
			Wend
			
			rsProducoes.Close
			Set rsProducoes = Nothing
			%>
			
			</select>
<details class="reports-picker" id="production-picker" hidden>
  <summary id="production-picker-summary" aria-labelledby="production-picker-label production-picker-choice"><span id="production-picker-label">Produ&ccedil;&atilde;o</span><span aria-hidden="true"> &middot; </span><span id="production-picker-choice">Selecione uma produ&ccedil;&atilde;o</span></summary>
  <div class="reports-picker-panel">
    <label for="production-picker-search">Buscar por c&oacute;digo ou nome</label>
    <input type="search" class="form-control" id="production-picker-search" autocomplete="off">
    <p class="reports-picker-status" id="production-picker-status" role="status"></p>
    <ul class="reports-picker-list" id="production-picker-list" aria-label="Produ&ccedil;&otilde;es"></ul>
  </div>
</details>
</div>
<button class="btn btn-primary" type="submit" name="filtro242" value=" ">Abrir relat&oacute;rio</button>
  </div>
</form>
</section>
			
			
			<section class="reports-card" aria-labelledby="report-7">
<h2 id="report-7">Retirada de &aacute;lbuns por per&iacute;odo</h2>
<form id="RELATORIO_RETIRADA" name="RELATORIO_RETIRADA" action="adm_relatorios_extrato_por_producao_e_vendedor_ano.asp" data-report-preview method="get" style="margin:0px" >
  <%Dim retiradaDataInicial, retiradaDataFinal
    retiradaDataFinal = Date()
    retiradaDataInicial = DateAdd("m", -12, retiradaDataFinal)%>
  <div class="reports-fields reports-compact">
    <label>Data inicial
      <input class="form-control reports-date" type="date" name="DATA_INICIO" required value="<%=Year(retiradaDataInicial) & "-" & Right("0" & Month(retiradaDataInicial), 2) & "-" & Right("0" & Day(retiradaDataInicial), 2)%>">
    </label>
    <label>Data final
      <input class="form-control reports-date" type="date" name="DATA_TERMINO" required value="<%=Year(retiradaDataFinal) & "-" & Right("0" & Month(retiradaDataFinal), 2) & "-" & Right("0" & Day(retiradaDataFinal), 2)%>">
    </label>
    <button class="btn btn-primary" type="submit" name="filtro242" value=" ">Abrir relat&oacute;rio</button>
  </div>
</form>
</section>		
			
			<%end if%>
				
			<%If VALIDA_SUB43 <> 0 Then%>

			<section class="reports-card" aria-labelledby="report-8">
<h2 id="report-8">Relat&oacute;rio Mensal de Produtos Vendidos - Valor Vendido</h2>
<form name="RELATORIO_MENSAL_PRODUTOS" action="adm_relatorios_extrato_mensal.asp" data-report-preview method="get" style="margin:0px">
  <div class="reports-fields reports-compact">
    <label>M&ecirc;s
      <select name="MES_INICIO" class="form-select">

			<option value="01" <%If Month(Now) = 1 then%> selected="selected"<%end if%>>Janeiro</option>
			
			<option value="02" <%If Month(Now) = 2 then%> selected="selected"<%end if%>>Fevereiro</option>
			
			<option value="03" <%If Month(Now) = 3 then%> selected="selected"<%end if%>>Mar&ccedil;o</option>
			
			<option value="04" <%If Month(Now) = 4 then%> selected="selected"<%end if%>>Abril</option>
			
			<option value="05" <%If Month(Now) = 5 then%> selected="selected"<%end if%>>Maio</option>
			
			<option value="06" <%If Month(Now) = 6 then%> selected="selected"<%end if%>>Junho</option>
			
			<option value="07" <%If Month(Now) = 7 then%> selected="selected"<%end if%>>Julho</option>
			
			<option value="08" <%If Month(Now) = 8 then%> selected="selected"<%end if%>>Agosto</option>
			
			<option value="09" <%If Month(Now) = 9 then%> selected="selected"<%end if%>>Setembro</option>
			
			<option value="10" <%If Month(Now) = 10 then%> selected="selected"<%end if%>>Outubro</option>
			
			<option value="11" <%If Month(Now) = 11 then%> selected="selected"<%end if%>>Novembro</option>
			
			<option value="12"<%If Month(Now) = 12 then%> selected="selected"<%end if%>>Dezembro</option>

			</select>
    </label>
    <label>Ano
      <select name="ANO_INICIO" class="form-select">
			
			<option value="<%=Year(Now)-3%>"><%=Year(Now)-3%></option>
			
			<option value="<%=Year(Now)-2%>"><%=Year(Now)-2%></option>
			
			<option value="<%=Year(Now)-1%>"><%=Year(Now)-1%></option>
			
			<option selected="selected"  value="<%=Year(Now)%>"><%=Year(Now)%></option>
			
			<option value="<%=Year(Now)+1%>"><%=Year(Now)+1%></option>
			
			<option value="<%=Year(Now)+2%>"><%=Year(Now)+2%></option>

			</select>
    </label>
    <button class="btn btn-primary" type="submit" name="filtro2" value=" ">Abrir relat&oacute;rio</button>
  </div>
</form>
</section>

			<%end if%>
				
				
			<%If VALIDA_SUB44 <> 0 Then%>

			<section class="reports-card" aria-labelledby="report-9">
<h2 id="report-9">Relat&oacute;rio de Venda de Produtos para Fornecedor</h2>
<form method="get" name="RELATORIO_PRODUTOS" action="cap_relatorio_venda_produtos.asp" data-report-preview style="margin:0px">
  <div class="reports-fields reports-compact">
    <label>Data inicial
      <input class="form-control reports-date" type="date" name="DATA_INI" required value="<%=Year(DateAdd("m",-1,Date())) & "-" & Right("0" & Month(DateAdd("m",-1,Date())), 2) & "-" & Right("0" & Day(DateAdd("m",-1,Date())), 2)%>">
    </label>
    <label>Data final
      <input class="form-control reports-date" type="date" name="DATA_FIM" required value="<%=Year(Date()) & "-" & Right("0" & Month(Date()), 2) & "-" & Right("0" & Day(Date()), 2)%>">
    </label>
    <button class="btn btn-primary" type="submit" name="filtro3" value=" ">Abrir relat&oacute;rio</button>
  </div>
</form>
</section>

			<%end if%>

			<%If VALIDA_SUB88 <> 0 Then%>

			<section class="reports-card" aria-labelledby="report-10">
<h2 id="report-10">Relat&oacute;rio de Venda de Bolsa e Encaderna&ccedil;&atilde;o por Vendedor</h2>
<form method="get" name="RELATORIO_PRODUTOS" action="cap_relatorio_venda_bolsa_encadernacao.asp" style="margin:0px">
  <div class="reports-fields reports-compact">
    <label>Data inicial
      <input class="form-control reports-date" type="text" onfocus="pickDate(this,this);" maxlength="10" name="DATA_INI" value="<%=ArrumaData(DateAdd("m",-1,Date()))%>" onclick="" />
    </label>
    <label>Data final
      <input maxlength="10" class="form-control reports-date" type="text" onfocus="pickDate(this,this);" name="DATA_FIM" value="<%=ArrumaData(Date())%>" onclick="" />
    </label>
    <button class="btn btn-primary" type="submit" name="filtro3" value=" ">Abrir relat&oacute;rio</button>
  </div>
</form>
</section>

			<%end if%>
				
			<%If VALIDA_SUB45 <> 0 Then%>

			<section class="reports-card" aria-labelledby="report-11">
<h2 id="report-11">Relat&oacute;rio de Vendas do Grupo Repasse - Ranking</h2>
<div class="reports-filter-scroll" tabindex="0" role="group" aria-label="Filtros">

			<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
			
			<tr>

			<td width="12%">
			
			<div align="right">
			
			<strong>M&ecirc;s / Ano</strong>
			
			</div>
			
			</td>

			<td width="88%">

			<form action="adm_relatorios_vendas_repasse.asp" method="get" style="margin:0px">					
			<table width="100%" border="0" cellspacing="0" cellpadding="3">

			<tr>
			
			<td width="18%">
			
			<div align="center">
			<select name="MES" style="font-size:13px">
			
			<option value="01" <%If Month(Now) = 1 then%> selected="selected"<%end if%>>Janeiro</option>
			
			<option value="02" <%If Month(Now) = 2 then%> selected="selected"<%end if%>>Fevereiro</option>
			
			<option value="03" <%If Month(Now) = 3 then%> selected="selected"<%end if%>>Mar&ccedil;o</option>
			
			<option value="04" <%If Month(Now) = 4 then%> selected="selected"<%end if%>>Abril</option>
			
			<option value="05" <%If Month(Now) = 5 then%> selected="selected"<%end if%>>Maio</option>
			
			<option value="06" <%If Month(Now) = 6 then%> selected="selected"<%end if%>>Junho</option>
			
			<option value="07" <%If Month(Now) = 7 then%> selected="selected"<%end if%>>Julho</option>
			
			<option value="08" <%If Month(Now) = 8 then%> selected="selected"<%end if%>>Agosto</option>
			
			<option value="09" <%If Month(Now) = 9 then%> selected="selected"<%end if%>>Setembro</option>
			
			<option value="10" <%If Month(Now) = 10 then%> selected="selected"<%end if%>>Outubro</option>
			
			<option value="11" <%If Month(Now) = 11 then%> selected="selected"<%end if%>>Novembro</option>
			
			<option value="12"<%If Month(Now) = 12 then%> selected="selected"<%end if%>>Dezembro</option>

			</select>

			</div>

			</td>

			<td width="8%">
			
			<div align="center">Ano</div>
			
			</td>

			<td width="11%">
			
			<div align="center">

			<select name="ANO" style="font-size:13px">
			
			<option value="<%=Year(Now)-3%>"><%=Year(Now)-3%></option>
			
			<option value="<%=Year(Now)-2%>"><%=Year(Now)-2%></option>
			
			<option value="<%=Year(Now)-1%>"><%=Year(Now)-1%></option>
			
			<option selected="selected"  value="<%=Year(Now)%>"><%=Year(Now)%></option>
			
			<option value="<%=Year(Now)+2%>"><%=Year(Now)+1%></option>
			
			<option value="<%=Year(Now)+2%>"><%=Year(Now)+2%></option>

			</select>
			
			</div>
			
			</td>

			<td width="11%">&nbsp;</td>

			<td width="52%" rowspan="2">
			
			<div align="right">
			
			<button class="btn btn-primary" type="submit" name="filtro2" value=" ">Abrir relat&oacute;rio</button>
			
			</div>
			
			</td>

			</tr>
			
			</table>
			
			</form>
			
			</td>
			
			</tr>
			
			</table></div>
</section>
			
			<%end if%>
			
			<%If VALIDA_SUB30 <> 0 Then%>
			
			<section class="reports-card" aria-labelledby="report-12">
<h2 id="report-12">Relat&oacute;rio de Sa&iacute;da de Produtos do Estoque para Fornecedor</h2>
<div class="reports-filter-scroll" tabindex="0" role="group" aria-label="Filtros">
			
			<form method="get" action="adm_relatorios_saida_fornecedores.asp" style="margin:0px">
			
			<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
			
			<tr>

			<td width="12%">&nbsp;</td>

			<td width="88%">
			
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			
			<tr>
			
			<td width="34%">
			
			<div align="center">
			
			<div align="left">Fornecedor</div>

			</div>
			
			</td>

			<td>M&ecirc;s</td>

			<td>Ano</td>

			<td width="23%" rowspan="2">
			
			<div align="right">
			
			<button class="btn btn-primary" type="submit" name="filtro2" value=" ">Abrir relat&oacute;rio</button>
			
			</div>
			
			</td>

			</tr>
			
			<tr valign="top">
			
			<td>
			
			<select name="COD_FORNECEDOR">
			
			<%
	  																							AbrirConexao()
	  
	  																							Set rsVendedor = objConexao.Execute("SELECT * FROM TB_FORNECEDORES ORDER BY NOME_FANTASIA")	  
	  
	  																						%>
                  
				  																			<option value="0" selected="selected">TODOS OS FORNECEDORES</option>
				  
				  																			<%While not rsVendedor.Eof %>
                  											
			  								<option value="<%=rsVendedor("COD_FORNECEDOR")%>"><%=rsVendedor("NOME_FANTASIA")%></option>
                  																			
			<%
			rsVendedor.MoveNext()
			Wend%>

			</select>
			
			</td>
			
			<td width="17%">
			
			<select name="MES" style="font-size:13px">

			<option value="01" <%If Month(Now) = 1 then%> selected="selected"<%end if%>>Janeiro</option>
			
			<option value="02" <%If Month(Now) = 2 then%> selected="selected"<%end if%>>Fevereiro</option>
			
			<option value="03" <%If Month(Now) = 3 then%> selected="selected"<%end if%>>Mar&ccedil;o</option>
			
			<option value="04" <%If Month(Now) = 4 then%> selected="selected"<%end if%>>Abril</option>
			
			<option value="05" <%If Month(Now) = 5 then%> selected="selected"<%end if%>>Maio</option>
			
			<option value="06" <%If Month(Now) = 6 then%> selected="selected"<%end if%>>Junho</option>
			
			<option value="07" <%If Month(Now) = 7 then%> selected="selected"<%end if%>>Julho</option>
			
			<option value="08" <%If Month(Now) = 8 then%> selected="selected"<%end if%>>Agosto</option>
			
			<option value="09" <%If Month(Now) = 9 then%> selected="selected"<%end if%>>Setembro</option>
			
			<option value="10" <%If Month(Now) = 10 then%> selected="selected"<%end if%>>Outubro</option>
			
			<option value="11" <%If Month(Now) = 11 then%> selected="selected"<%end if%>>Novembro</option>
			
			<option value="12" <%If Month(Now) = 12 then%> selected="selected"<%end if%>>Dezembro</option>
			
			</select>
			
			</td>

			<td width="26%">
			
			<div align="left">
			
			<input name="limit" value="0" type="hidden" />
			
			<select name="ANO" style="font-size:13px">
			
			<option value="<%=Year(Now)-2%>"><%=Year(Now)-2%></option>
			
			<option value="<%=Year(Now)-1%>"><%=Year(Now)-1%></option>
			
			<option selected="selected" value="<%=Year(Now)%>"><%=Year(Now)%></option>
			
			<option value="<%=Year(Now)+1%>"><%=Year(Now)+1%></option>

			</select>
			
			</div>
			
			</td>

			</tr>

			</table>

			</td>

			</tr>

			</table>

			</form></div>
</section>
			
			<%end if%>
				
				
			<%If VALIDA_SUB41 <> 0 Then%>

			<section class="reports-card" aria-labelledby="report-13">
<h2 id="report-13">Consulta de Cheques</h2>
<div class="reports-filter-scroll" tabindex="0" role="group" aria-label="Filtros">

			<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
			
			<tr>
			
			<td width="96%">
			
			<form name="FORM_PADRAO" action="cap_cheques_consulta.asp" style="margin:0px" onsubmit="return validaForm()">
			
			<table width="100%" border="0" cellspacing="0" cellpadding="0">

			<tr>
			<td width="80%">
			
			<strong>Escolha 1 ou mais par&acirc;metros abaixo para efetuar a busca</strong>
			
			</td>


			<td width="20%" rowspan="2">
			
			<div align="right">
			
			<button class="btn btn-primary" type="submit" tabindex="7" name="filtro1" value=" ">Abrir relat&oacute;rio</button>
			
			</div>
			
			</td>

			</tr>

			<tr>

			<td>
			
			<table width="100%" border="0" cellpadding="2" cellspacing="0">
			
			<tr>
			
			<td width="78">
			
			<strong>Banco</strong>
			
			</td>
			
			<td width="108">
			
			<input maxlength="30" tabindex="1" type="text" value="" class="inputTamanho1" name="BANCO" />
			
			</td>
			
			<td width="84">
			
			<strong>Ag&ecirc;ncia</strong>
			
			</td>
			
			<td width="102">
			
			<input tabindex="2" type="text" value="" class="inputTamanho1" name="AGENCIA" />
			
			</td>
			
			<td width="91">
			
			<strong>Contato Corrente</strong>
			
			</td>
			
			<td width="124">
			
			<input tabindex="3" type="text" value="" class="inputTamanho1" name="CONTA_CORRENTE" />
			
			</td>

			</tr>
			
			<tr>

			<td>
			
			<strong>N&uacute;mero do Cheque</strong>
			
			</td>

			<td>
			
			<input tabindex="4" type="text" value="" class="inputTamanho1" name="NUMERO_CHEQUE" />
			
			</td>

			<td>
			
			<strong>Valor exato</strong>
			
			</td>

			<td>
			
			<input tabindex="5" type="text" onKeydown="Formata(this,20,event,2)" value="" class="inputTamanho1" name="VALOR" />
			
			</td>
			<td>
			
			<strong>Bom para</strong>
			
			</td>

			<td>
			
			<input class="campoData" type="text" onkeypress="MascaraData(FORM_PADRAO.BOM_PARA);" maxlength="10" tabindex="6" value="" name="BOM_PARA" /> 
			(dd/mm/aaaa)
			
			</td>

			</tr>

			</table>

			</td>

			</tr>

			</table>

			</form>

			</td>

			</tr>

			</table></div>
</section>

			<%End If%>
				
			<%If VALIDA_SUB49 <> 0 Then%>
						
			<section class="reports-card" aria-labelledby="report-14">
<h2 id="report-14">Relat&oacute;rio de Formandos por Produ&ccedil;&atilde;o</h2>
<div class="reports-filter-scroll" tabindex="0" role="group" aria-label="Filtros">
					
			<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
		
			<tr>
			
			<td width="12%">
			
			<div align="right">
			
			<strong>Filtro  por produ&ccedil;&atilde;o </strong>
			
			</div>
			
			</td>

			<td width="88%">

			<form name="RELATORIO_FORMANDOS_PRODUCAO" action="adm_relatorios_formandos_producao.asp" target="_blank" method="get" style="margin:0px"> 
			
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			
			<tr>
			
			<td width="77%">
			
			<div align="left">
			
			<select name="COD_PRODUCAO" id="COD_PRODUCAO">
			
			<option selected="selected" value="0">SELECIONE A PRODU&Ccedil;&Atilde;O DESEJADA</option>

			<%
			AbrirConexao()
			Set rsProducoes = objConexao.Execute("SELECT COD_PRODUCAO, DESCRICAO FROM TB_PRODUCAO WHERE PRODUCAO_FINALIZADA = 1 ORDER BY cod_producao desc")
			While not rsProducoes.Eof
			%>
			
			<option value="<%=rsProducoes("COD_PRODUCAO")%>"><%=rsProducoes("cod_producao")%> - <%=rsProducoes("DESCRICAO")%></option>

			<%
			rsProducoes.MoveNext()
			Wend
			
			rsProducoes.Close
			Set rsProducoes = Nothing
			%>
			
			</select>
			
			</div>
			
			</td>

			<td width="23%">
			
			<div align="right">
			
			<button class="btn btn-primary" type="submit" name="filtro242" value=" ">Abrir relat&oacute;rio</button>
			</div>
			
			</td>
			</tr>

			</table>

			</form>
			
			</td>
			
			</tr>
			
			</table></div>
</section>
			
			<%end if%>

			<%If VALIDA_SUB77 <> 0 Then%>
			
			<section class="reports-card" aria-labelledby="report-15">
<h2 id="report-15">Relat&oacute;rio de Sa&iacute;da de produtos produzidos p&oacute;s-venda - Despacho</h2>
<div class="reports-filter-scroll" tabindex="0" role="group" aria-label="Filtros">

			<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
			
			<tr>

			<td width="13%">
			
			<div align="right">
			
			<strong>Data do despacho: </strong>
			
			</div>
			
			</td>

			<td width="87%">

			<form method="get" name="RELATORIO_DESPACHO" target="_blank" action="adm_relatorios_ordens_servico_despacho_imprimir.asp" style="margin:0px">
			
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			
			<tr>
			
			<td>
			
			<table width="316" border="0" cellspacing="0" cellpadding="3">

			<tr>
			
			<td width="34%">
			
			<div align="center">
			
			<input class="campoData" type="text" onfocus="pickDate(this,this);" maxlength="10" name="DATA_INICIAL" value="<%=ArrumaData(DateAdd("m",-1,Date()))%>" onclick="" />
			
			</div>
			
			</td>

			<td width="5%">
			
			<div align="center">&agrave;</div>
			
			</td>
			
			<td width="34%">
			
			<div align="center">
			
			<input maxlength="10" class="campoData" type="text" onfocus="pickDate(this,this);" name="DATA_FINAL" value="<%=ArrumaData(Date())%>" onclick="" />
			
			</div>
			
			</td>
			
			<td width="27%">&nbsp;</td>
			
			</tr>

			</table>
			
			</td>

			<td>
			
			<div align="right">
			
			
			<button class="btn btn-primary" type="submit" name="filtro3" value=" ">Abrir relat&oacute;rio</button>
			
			</div>
			
			</td>

			</tr>

			</table>

			</form>
			
			</td>
			
			</tr>

			</table></div>
</section>

			<section class="reports-card" aria-labelledby="report-16">
<h2 id="report-16">Relat&oacute;rio de Sa&iacute;da de produtos produzidos p&oacute;s-venda - Entregues</h2>
<div class="reports-filter-scroll" tabindex="0" role="group" aria-label="Filtros">
			
			<form method="get" name="RELATORIO_ENTREGA" target="_blank" action="adm_relatorios_ordens_servico_entrega_imprimir.asp" style="margin:0px">
			
			<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
			
			<tr>
			
			<td width="13%">
			
			<div align="right">
			
			<strong>Data do lan&ccedil;amento: </strong>
			
			</div>
			
			</td>
			
			<td width="87%">
			
			<table width="316" border="0" cellspacing="0" cellpadding="3">
			
			<tr>
			
			<td width="34%">
			
			<div align="center">
			
			<input class="campoData" type="text" onfocus="pickDate(this,this);" maxlength="10" name="DATA_INICIAL_CADASTRO" value="" onclick="" />
			
			</div>
			
			</td>
			
			<td width="5%">
			
			<div align="center">&agrave;</div>
			
			</td>

			<td width="34%">
			
			<div align="center">
			
			<input maxlength="10" class="campoData" type="text" onfocus="pickDate(this,this);" name="DATA_FINAL_CADASTRO" value="" onclick="" />
			
			</div>
			
			</td>
			
			<td width="27%">&nbsp;</td>

			</tr>
			
			</table>
			
			</td>

			</tr>

			<!--</tr>-->

			<tr>
			
			<td width="13%">
			
			<div align="right">
			
			<strong>Data da entrega: </strong>
			
			</div>
			
			</td>

			<td width="87%">
			
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			
			<tr>
			
			<td>
			
			<table width="316" border="0" cellspacing="0" cellpadding="3">

			<tr>

			<td width="34%">
			
			<div align="center">
			
			<input class="campoData" type="text" onfocus="pickDate(this,this);" maxlength="10" name="DATA_INICIAL" value="<%=ArrumaData(DateAdd("m",-1,Date()))%>" onclick="" />
			
			</div>
			
			</td>

			<td width="5%">
			
			<div align="center">&agrave;</div>
			
			</td>

			<td width="34%">
			
			<div align="center">

			<input maxlength="10" class="campoData" type="text" onfocus="pickDate(this,this);" name="DATA_FINAL" value="<%=ArrumaData(Date())%>" onclick="" />
			
			</div>
			
			</td>

			<td width="27%">&nbsp;</td>

			</tr>
			
			</table>
			
			</td>

			<td>
			
			<div align="left">
			
			<strong>Transportadora: </strong>
			
			</div>
			
			</td>

			<td>
			
			<div align="left">
			
			<%=PadraoFormularioSelect("despacho_transportador","","inputTamanho1","","","Todos",letra,"PADRAO:TRANSPORTADORA#TRANSPORTADORA|CORREIOS#CORREIOS|RETIRADA#RETIRADA|STILLOS#STILLOS|Antonio Felipe#Antonio Felipe", "incluir", false, false, false )%>

			</div>

			</td>

			<td>
			
			<div align="right">
			
			<button class="btn btn-primary" type="submit" name="filtro3" value=" ">Abrir relat&oacute;rio</button>
			
			</div>
			
			</td>
			
			</tr>
			
			</table>
			
			</td>
			
			</tr>
			
			</table>
			
			</form></div>
</section>	

			<%END If

			If VALIDA_SUB49 <> 0 Then %>

			<section class="reports-card" aria-labelledby="report-17">
<h2 id="report-17">Relat&oacute;rio de Produtividade de Funcion&aacute;rios por OS</h2>
<div class="reports-filter-scroll" tabindex="0" role="group" aria-label="Filtros">

			<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
			
			<tr>
			
			<td width="13%">
			
			<div align="right">
			
			<strong>Escolha o per&iacute;odo: </strong>
			
			</div>
			
			</td>										
			
			<td width="87%">
			
			<form method="get" name="RELATORIO_PRODUTOS" action="ppv_relatorio_produtividade.asp" style="margin:0px">
			
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			
			<tr>
			
			<td>
			
			<table width="316" border="0" cellspacing="0" cellpadding="3">
			
			<tr>
			
			<td width="34%">
			
			<div align="center">
			
			<input class="campoData" type="text" onfocus="pickDate(this,this);" maxlength="10" name="DATA_INI" value="<%=ArrumaData(DateAdd("m",-1,Date()))%>" onclick="" />
			
			</div>
			
			</td>
			
			<td width="5%">
			
			<div align="center">&agrave;</div>
			
			</td>

			<td width="34%">
			
			<div align="center">
			
			<input maxlength="10" class="campoData" type="text" onfocus="pickDate(this,this);" name="DATA_FIM" value="<%=ArrumaData(Date())%>" onclick="" />
			
			</div>
			
			</td>
			
			<td width="27%">&nbsp;</td>
			</tr>
			
			</table>
			
			</td>

			<td>
			
			<div align="right">

			<button class="btn btn-primary" type="submit" name="filtro3" value=" ">Abrir relat&oacute;rio</button>

			</div>
			
			</td>

			</tr>

			</table>

			</form>
			
			</td>
			
			</tr>
			
			</table></div>
</section>

			<%END If%>
				
			<!-- Relat?rios -->

			<%If VALIDA_SUB78 <> 0 Then%>

			<section class="reports-card" aria-labelledby="report-18">
<h2 id="report-18">Relat&oacute;rio de Pedidos (Estoque Pr&eacute;-Venda)</h2>
<div class="reports-filter-scroll" tabindex="0" role="group" aria-label="Filtros">
			
			<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="80%" border="0" 	align="center" cellpadding="5" cellspacing="0">
			
			<tr>
			
			<td width="87%">
			
			<form target="_BLANK" method="get" name="RELATORIO_ESTOQUE_PEDIDOS" action="ppv_relatorio_estoque_pedidos.asp" style="margin:0px">
			
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			
			<tr>
			
			<td>
			
			<table width="700" border="0" cellspacing="0" cellpadding="3">
			
			<tr>
			
			<td width="">
			
			<div align="left">
			
			<strong>Codigo Pedido: </strong>
			
			</div>

			</td>

			<td>

			<div align="center">
			
			<input  style="width: 100px;padding: 2px;" type="text" maxlength="10" name="cod_pedido" value="" onclick="" />
			
			</div>

			</td>

			<td width="">

			<div align="right">

			<strong>Status do pedido: </strong>

			</div>

			</td>

			<td>
			
			<div align="center">
			
			<select name="status_pedido">
			
			<option value="0">Todos </option>
			
			<option value="EA"> Em andamento </option>
			
			<option value="CD"> Concluído </option>
			
			<option value="CN"> Cancelado </option>

			</select>

			</div>
			
			</td>

			<td >

			<div align="right">
			
			<strong>Tipo de relat&oacute;rio: </strong>

			</div>

			</td>

			<td>

			<div align="center">
			
			<select name="tipo_relatorio" title="Relat&oacute;rio simples exibe apenas os pedidos, j&aacute; o detalhado exibe tamb&eacute;m seu itens.">
			
			<option value="S">Simples </option>
			
			<option value="D"> Detalhado </option>

			</select>

			</div>

			</td>									
			</tr>					
			
			<tr>
			<td width="">

			<div align="left" title="Periodo no qual foi feito a solicita&ccedil;&atilde;o do pedido.">
			
			<strong>Escolha o per&iacute;odo: </strong>

			</div>

			</td>
			
			<td>
			<div align="center">
			
			<input class="campoData" type="text" onfocus="pickDate(this,this);" maxlength="10" name="data_inicial" value="<%=ArrumaData(DateAdd("m",-1,Date()))%>" onclick="" />
			
			</div>

			</td>

			<td>
			
			<div align="center">&agrave;</div>

			</td>

			<td>
			
			<div align="center">
			
			<input maxlength="10" class="campoData" type="text" onfocus="pickDate(this,this);" name="data_final" value="<%=ArrumaData(Date())%>" onclick="" />
			
			</div>

			</td>

			<td></td>

			<td></td>
			
			</tr>
			
			</table>
			
			</td>

			<td>
			<div align="right">

			<button class="btn btn-primary" type="submit" name="filtro3" value=" ">Abrir relat&oacute;rio</button>

			</div>

			</td>
			
			</tr>
			
			</table>

			</form>

			</td>
			
			</tr>

			</table></div>
</section>

			<%END If%>


			


			<%If VALIDA_SUB78 <> 0 and 1= 2 Then%>
			<!--
			<h3 align="center">Extrato Estoque Interno (Produtos)</h3>

			<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
			<tr>
			<td width="87%">
			<form target="_BLANK" method="get" name="RELATORIO_ESTOQUE_PEDIDOS" action="est_extrato.asp" style="margin:0px">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr>
			<td>
			<table width="700" border="0" cellspacing="0" cellpadding="3">
			<tr>
			<td>
			<div align="left">
			<strong>Produto*: </strong>
			</div>
			</td>
			<td colspan="5">
			<div align="left">
			<%=PadraoFormularioSelect("f_cod_produto","","inputTamanho3","","","Selecione","","SELECT cod_produto, nome_produto FROM TB_PRODUTOS WHERE ativo = 1 ORDER BY ORDER_MENU","incluir",false,false,true)%>
			</select>
			</div>
			</td>									
			</tr>					
			<tr>
			<td width="120px;">
			<div align="left" title="Periodo no qual foi feita a moveimenta??o">
			<strong>Escolha o per&iacute;odo: </strong>
			</div>
			</td>
			<td width="120px;">
			<div align="left">
			<input class="campoData" type="text" onfocus="pickDate(this,this);" maxlength="10" name="f_periodo_inicio" value="<%=ArrumaData(DateAdd("m",-1,Date()))%>" onclick="" />
			</div>
			</td>
			<td width="15px;">
			<div align="center">
			&agrave;
			</div>
			</td>
			<td width="120px;">
			<div align="center">
			<input maxlength="10" class="campoData" type="text" onfocus="pickDate(this,this);" name="f_periodo_fim" value="<%=ArrumaData(Date())%>" onclick="" />
			</div>
			</td>
			<td width="120px;">
			<div align="left" title="Periodo no qual foi feita a moveimenta??o">
			<strong>Em posse de: </strong>
			</div>
			</td>
			<td>
			<%=PadraoFormularioSelect("f_cod_local_armazenamento","","inputTamanho2","","","Selecione","","SELECT cod_local_armazenamento, nome FROM`tb_est_local_armazenamento","incluir",false,false,true)%>

			</td>
			</tr>
			</table>
			</td>
			<td>
			<div align="right">
			<button class="btn btn-primary" type="submit" name="filtro3" value=" ">Abrir relat&oacute;rio</button>
			</div>
			</td>
			</tr>
			</table>
			<input type="hidden" name="filtro" value="1">
			</form>
			</td>
			</tr>
			</table>
			-->
			<%END If%>


			<% if validaUsuario(1, 371,"", false) then %>

			<section class="reports-card" aria-labelledby="report-19">
<h2 id="report-19">Relat&oacute;rio para retirada de produ&ccedil;&atilde;o</h2>
<div class="reports-filter-scroll" tabindex="0" role="group" aria-label="Filtros">
			
			<table style="border:#999999 dashed 1px; padding:1px; background:#FFFFCC; width: 80%; margin: 0 auto; box-sizing: border-box;" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr>
    <td width="100%">
      <form target="_BLANK" method="get" action="adm_relatorios_vendedores_historico.asp" style="margin:0px">
        
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			
			<tr>
			
			<td>
			
			<table width="700" border="0" cellspacing="0" cellpadding="3">
			
			<tr>

			<td width="">

			<div align="left" title="Periodo no qual foi feito a solicitação do pedido.">
			
			<strong>Escolha o per&iacute;odo: </strong>
			
			</div>

			</td>

			<td>
			<div align="center">
			
			<input class="campoData" type="text" maxlength="10" name="a" value="<%=ArrumaData(DateAdd("m",-1,Date()))%>" onclick="" />
			
			</div>

			</td>

			<td>

			<div align="center">&agrave;</div>

			</td>

			<td>

			<div align="center">
			
			<input maxlength="10" class="padrao_data" type="text"  name="b" value="<%=ArrumaData(Date())%>" onclick="" />

			</div>

			</td>

			<td>
			
			<%=padraoFormularioSelect("c","","","","","SEM_SELECIONE","","PADRAO:false#Sem valores|true#Com Valores","incluir", false, false, false)%>

			</td>

			</tr>

			</table>
			
			</td>

			<td>
			<div align="right">
			
			<button class="btn btn-primary" type="submit" name="filtro3" value=" ">Abrir relat&oacute;rio</button>
			
			</div>

			</td>
			
			</tr>
			
			</table>
			
			</form>
			
			</td>
			
			</tr>
			
			</table></div>
</section>

			<% end if %>


			<section class="reports-card" aria-labelledby="report-20">
<h2 id="report-20">Relat&oacute;rio de hist&oacute;rico de vendedores por produ&ccedil;&atilde;o</h2>
<div class="reports-filter-scroll" tabindex="0" role="group" aria-label="Filtros">
			
			<table style="border:#999999 dashed 1px; padding:1px; max-width: 100%; box-sizing: border-box; background:#FFFFCC" width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
			
			<tr>

			<td width="100%">

			<form target="_BLANK" method="get" name="" action="adm_relatorios_vendedores_historico.asp" style="margin:0px">
			
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			
			<tr>

			<td>
			
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			
			<tr>

			<td>

			<div align="center">&agrave;</div>

			</td>

			<td>
			<div align="center">
			
			<%=padraoFormularioSelect("P","","","","","SEM_SELECIONE","","SELECT COD_PRODUCAO, CAST(CONCAT(COD_PRODUCAO ,' - ', DESCRICAO) AS CHAR) AS DESCRICAO FROM TB_PRODUCAO ORDER BY COD_PRODUCAO DESC","incluir", false, false, false)%>
			
			</div>

			</td>

			<td>

			<%=padraoFormularioSelect("v","","","","","SEM_SELECIONE","","SELECT COD_VENDEDOR, NOME_VENDEDOR FROM TB_VENDEDORES WHERE ATIVADO = 1 AND cod_tabela_comissao_vendedores = 1 ORDER BY 2","incluir", false, false, false) %>

			</td>

			</tr>

			</table>

			</td>

			<td>

			<div align="right">
			
			<button class="btn btn-primary" type="submit" name="filtro3" value=" ">Abrir relat&oacute;rio</button>

			</div>

			</td>
			</tr>
			
			</table>

			</form>

			</td>

			</tr>

			</table></div>
</section>

			
      </div>
    </main>
    <footer class="dashboard-footer">&copy; <%=Year(Date)%> Grupo Stillos <span>Sistema administrativo interno</span></footer>
  </div>
</div>
<dialog class="report-notice" id="report-production-notice" aria-labelledby="report-notice-title" aria-describedby="report-notice-message">
  <h2 id="report-notice-title">Selecione uma produ&ccedil;&atilde;o</h2>
  <p id="report-notice-message">Escolha a produ&ccedil;&atilde;o para abrir o extrato de aproveitamento.</p>
  <form method="dialog"><button class="btn btn-primary" autofocus>Entendi</button></form>
</dialog>
<script>
(function () {
    var form = document.forms.EXTRATO_APROVEITAMENTO;
    if (!form) return;
    var production = form.elements.COD_PRODUCAO;
    var picker = document.getElementById('production-picker');
    var summary = document.getElementById('production-picker-summary');
    var choice = document.getElementById('production-picker-choice');
    var search = document.getElementById('production-picker-search');
    var list = document.getElementById('production-picker-list');
    var status = document.getElementById('production-picker-status');

    function normalize(text) {
        return text.toLocaleLowerCase('pt-BR').normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    }

    var entries = Array.from(production.options).slice(1).map(function (option) {
        var item = document.createElement('li');
        var button = document.createElement('button');
        button.type = 'button';
        button.value = option.value;
        button.textContent = option.textContent.trim();
        item.appendChild(button);
        list.appendChild(item);
        return { item:item, button:button, text:normalize(button.textContent) };
    });

    function showSelection() {
        choice.textContent = production.options[production.selectedIndex].textContent.trim();
        entries.forEach(function (entry) {
            entry.button.setAttribute('aria-pressed', String(entry.button.value === production.value));
        });
    }

    function filterProductions() {
        var query = normalize(search.value.trim());
        var visible = 0;
        entries.forEach(function (entry) {
            entry.item.hidden = !entry.text.includes(query);
            if (!entry.item.hidden) visible++;
        });
        status.textContent = visible ? '' : 'Nenhuma produ\u00e7\u00e3o encontrada.';
        list.scrollTop = 0;
    }

    list.addEventListener('click', function (event) {
        var button = event.target.closest('button');
        if (!button) return;
        production.value = button.value;
        showSelection();
        picker.open = false;
        summary.focus();
    });
    search.addEventListener('input', filterProductions);
    search.addEventListener('keydown', function (event) {
        if (event.key === 'Enter') event.preventDefault();
        if (event.key === 'ArrowDown') {
            var first = list.querySelector('li:not([hidden]) button');
            if (!first) return;
            event.preventDefault();
            first.focus();
        }
    });
    picker.addEventListener('toggle', function () {
        if (!picker.open) return;
        search.value = '';
        filterProductions();
        search.focus();
    });
    picker.addEventListener('keydown', function (event) {
        if (event.key !== 'Escape' || !picker.open) return;
        picker.open = false;
        summary.focus();
    });
    document.getElementById('report-production-notice').addEventListener('close', function () {
        if (!picker.open) picker.open = true;
        else search.focus();
    });
    window.addEventListener('pageshow', showSelection);
    showSelection();
    production.hidden = true;
    document.getElementById('production-native-label').hidden = true;
    picker.hidden = false;
})();
</script>
<dialog class="report-preview" id="report-preview" aria-labelledby="report-preview-title">
  <header>
    <h2 id="report-preview-title">Pr&eacute;via do relat&oacute;rio</h2>
    <div class="report-preview-actions">
      <button class="btn btn-primary" id="report-preview-print" type="button" disabled>Imprimir / Salvar PDF</button>
      <form method="dialog"><button class="btn btn-outline-secondary" autofocus>Fechar</button></form>
    </div>
  </header>
  <p class="report-preview-status" role="status" id="report-preview-status" hidden></p>
  <iframe id="report-preview-frame" title="Pr&eacute;via do relat&oacute;rio" hidden></iframe>
</dialog>
<script>
(function () {
    var rankingStart = document.getElementById('reports-ranking-start');
    var rankingEnd = document.getElementById('reports-ranking-end');
    if (rankingStart && rankingEnd) {
        function syncRankingDates() {
            rankingEnd.min = rankingStart.value;
        }
        rankingStart.addEventListener('input', syncRankingDates);
        window.addEventListener('pageshow', syncRankingDates);
        syncRankingDates();
    }

    var modal = document.getElementById('report-preview');
    var frame = document.getElementById('report-preview-frame');
    var title = document.getElementById('report-preview-title');
    var status = document.getElementById('report-preview-status');
    var print = document.getElementById('report-preview-print');
    var trigger;

    if (typeof modal.showModal !== 'function') return;

    document.addEventListener('submit', function (event) {
        var form = event.target;
        if (!form.matches('[data-report-preview]') || event.defaultPrevented) return;
        var url = new URL(form.action);
        var data = new FormData(form, event.submitter);
        data.set('modal', '1');
        url.search = new URLSearchParams(data).toString();
        event.preventDefault();
        trigger = event.submitter || document.activeElement;
        title.textContent = form.closest('.reports-card').querySelector('h2').textContent;
        frame.title = title.textContent;
        status.textContent = 'Carregando relat\u00f3rio...';
        status.hidden = false;
        status.setAttribute('aria-busy', 'true');
        frame.hidden = true;
        print.disabled = true;
        frame.src = url.href;
        if (!modal.open) modal.showModal();
    });

    frame.addEventListener('load', function () {
        var preview = frame.contentDocument;
        if (!modal.open || (preview && preview.URL === 'about:blank')) return;
        var report = preview && preview.querySelector('[data-cap-report], [data-report-preview-content]');
        frame.hidden = !report;
        print.disabled = !report || Boolean(report.querySelector('[data-report-error]'));
        status.removeAttribute('aria-busy');
        status.hidden = Boolean(report);
        if (!report) status.textContent = 'N\u00e3o foi poss\u00edvel abrir a pr\u00e9via. Confira os filtros ou tente novamente.';
        if (report) {
            preview.addEventListener('keydown', function (event) {
                if (event.key === 'Escape' && !event.defaultPrevented && !preview.querySelector('dialog[open]')) modal.close();
            });
        }
    });

    print.addEventListener('click', function () {
        if (print.disabled) return;
        var preview = frame.contentDocument;
        if (preview && preview.querySelector('[data-supplier-detail-print]')) {
            var url = new URL(preview.URL);
            url.searchParams.delete('modal');
            url.searchParams.set('imprimir', '1');
            window.open(url.href, '_blank', 'noopener');
            return;
        }
        frame.contentWindow.print();
    });

    modal.addEventListener('close', function () {
        frame.hidden = true;
        frame.removeAttribute('src');
        print.disabled = true;
        if (trigger) trigger.focus();
    });
})();
</script>
<script>
(function () {
    var toggle = document.getElementById('ppvMenuToggle');
    var menu = document.getElementById('menuPrincipal');
    var previous = menu.querySelector('[aria-current="page"]');
    var current = menu.querySelector('a[href="adm_relatorios.asp"]');
    if (previous) {
        previous.removeAttribute('aria-current');
        previous.classList.remove('active');
    }
    if (current) {
        current.setAttribute('aria-current', 'page');
        current.classList.add('active');
    }
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
    var search = document.getElementById('reports-search');
    var cards = Array.from(document.querySelectorAll('.reports-card'));
    var status = document.getElementById('reports-status');
    if (!cards.length) {
        status.textContent = 'Nenhum relatório disponível para o seu perfil.';
        return;
    }
    function normalize(value) {
        return value.toLocaleLowerCase('pt-BR').normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    }
    var titles = cards.map(function (card) { return normalize(card.querySelector('h2').textContent); });
    document.getElementById('reports-search-panel').hidden = false;
    search.addEventListener('input', function () {
        var terms = normalize(search.value.trim()).split(/\s+/);
        var shown = 0;
        cards.forEach(function (card, index) {
            card.hidden = !terms.every(function (term) { return titles[index].includes(term); });
            if (!card.hidden) shown++;
        });
        status.textContent = shown ? '' : 'Nenhum relatório encontrado.';
    });
})();
</script>
</body>
</html>
