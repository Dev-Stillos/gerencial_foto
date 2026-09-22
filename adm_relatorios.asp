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

'Previs�o de Recebimento
VALIDA_SUB72 = formatNumber(0&inStr(sub_validacoes_paginas,"#72"),0)

'Fichas autenticadas pela diretoria
VALIDA_SUB13 = formatNumber(0&inStr(sub_validacoes_paginas,"#13"),0)

'Extrato de Vendedor
VALIDA_SUB15 = formatNumber(0&inStr(sub_validacoes_paginas,"#15"),0)

'Relat�rio de Vendas por Produ��o
VALIDA_SUB16 = formatNumber(0&inStr(sub_validacoes_paginas,"#16"),0)

'Relat�rio Foto Nota - Fechamento para Pagamento
VALIDA_SUB29 = formatNumber(0&inStr(sub_validacoes_paginas,"#29"),0)

'Relat�rio de Sa�da para Fornecedor
VALIDA_SUB30 = formatNumber(0&inStr(sub_validacoes_paginas,"#30"),0)

'Extrato de produ��o com lotes
VALIDA_SUB31 = formatNumber(0&inStr(sub_validacoes_paginas,"#31"),0)

'Relat�rio de Cheques
VALIDA_SUB41 = formatNumber(0&inStr(sub_validacoes_paginas,"#41"),0)

'Relat�rio Mensal de Produtos
VALIDA_SUB43 = formatNumber(0&inStr(sub_validacoes_paginas,"#43"),0)

'Relat�rio de Vendas de Produtos Completo
VALIDA_SUB44 = formatNumber(0&inStr(sub_validacoes_paginas,"#44"),0)

'Relat�rio de Vendas do Grupo Repasse - Ranking
VALIDA_SUB45 = formatNumber(0&inStr(sub_validacoes_paginas,"#45"),0)

'Relat�rio de Formandos por Produ��o
VALIDA_SUB49 = formatNumber(0&inStr(sub_validacoes_paginas,"#49"),0)

'Relat�rio de Produtividade de Funcion�rios por OS
VALIDA_SUB59 = formatNumber(0&inStr(sub_validacoes_paginas,"#59"),0)

'Relat�rio de Despacho
VALIDA_SUB77 = formatNumber(0&inStr(sub_validacoes_paginas,"#77"),0)
'Relat�rio de Entrega
'VALIDA_SUB78 = formatNumber(0&inStr(sub_validacoes_paginas,"#78"),0)

'Relat�rio de Compras
VALIDA_SUB78 = formatNumber(0&inStr(sub_validacoes_paginas,"#78"),0)

'Relat�rio Venda Bolsa e Encaderna��o por Fornecedor
VALIDA_SUB88 = formatNumber(0&inStr(sub_validacoes_paginas,"#88"),0)

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

	<head>
		
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

		<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>

		<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>

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

			//Formata n�mero tipo moeda usando o evento onKeyDown
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
	
			AJUDA F1 E CTRL PARA BUSCA RÁPIDA-->
			
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
			
			AJUDA F1 E CTRL PARA BUSCA RÁPIDA

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
			
			function validaForm7(){
			d = document.EXTRATO_APROVEITAMENTO
				if(d.COD_PRODUCAO.value == 0){
					alert('Selecione a produção desejada!')
					d.COD_PRODUCAO.focus();
					return false;
				}else {
					return true;
				}
			}
		
		</script>

	</head>

	<body onKeyDown="javascript:Verificar()">

		<table cellpadding="1px" cellspacing="0" width="100%">
		
			<tr>
		
				<td>
		
					<table width="100%" align="center" border="0" cellpadding="10" cellspacing="0" bgcolor="#FFFFFF">
						
						<tr>
							
							<td>
								
								<!-- #include file="includes/topo.asp" -->
								 <style>
									/* Sobrescreve a largura só nessa página */
									.input-filtro {
										width: 100px !important;
									}
									</style>

								
								<table width="100%" style="border-bottom:#999999" cellspacing="0" cellpadding="0">
								
									<tr>

										<td><!-- #include file="includes/menu.asp" --></td>
									
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
												
														<h2>Relat&oacute;rios </h2>
														
														<div align="center" class="msg">
																
															<%=Request.QueryString("msg")%>
															
														</div>
															
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
																<input class="btnFiltrarAmarelo" type="submit" name="filtro22" value=" " />
																</div></td>
															</tr>
															</table></td>
														</tr>
														</table>
														</form>
														<%End If%>
														-->
				
														<%If VALIDA_SUB18 <> 0 Then%>
															<h3 align="center">Ranking de Vendedores - Completo </h3>

															<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
					
																<tr>

																	<td width="13%">
																		
																		<div align="right">
																		
																			<strong>Escolha o per&iacute;odo: </strong>
																		
																		</div>
																	
																	</td>

																	<td width="87%">
																		
																		<form method="get" name="RELATORIO_RANKING_VENDEDORES" action="adm_relatorios_ranking_vendedores_completo.asp" style="margin:0px">
																		
																			<table width="100%" border="0" cellspacing="0" cellpadding="0">

																				<tr>

																					<td>
																					
																						<table width="438" border="0" cellspacing="0" cellpadding="3">

																							<tr>
																								
																								<td width="25%">
																								
																									<div align="center">
																									
																										<input class="campoData" type="text" onfocus="pickDate(this,document.RELATORIO_RANKING_VENDEDORES.DATA_VENDA_INICIO);" maxlength="10" name="DATA_VENDA_INICIO" value="" onclick="" />
																									
																									</div>
																								
																								</td>
																							
																								<td width="3%">
																									
																									<div align="center">&agrave;</div>
																								
																								</td>

																								<td width="25%">
																								
																									<div align="center">
																										
																										<input maxlength="10" class="campoData" type="text" onfocus="pickDate(this,document.RELATORIO_RANKING_VENDEDORES.DATA_VENDA_FIM);" name="DATA_VENDA_FIM" value="" />
																									</div>
																								
																								</td>
																								
																								<td width="47%">
																								
																									<em>Filtro pela data da venda dos pedidos</em>
																								
																								</td>
																							
																							</tr>
																						
																						</table>
																					</td>
																					
																					<td>
																					
																						<div align="right">
																							
																							<input class="btnFiltrarAmarelo" type="submit" name="filtro3" value=" " />
																						</div>
																					
																					</td>

																				</tr>

																			</table>

																		</form>	
																	
																	</td>

																</tr>

																<tr>

																	<td colspan="2"></td>

																</tr>

															</table>
														<%End If%>					
				
														<%If VALIDA_SUB72 <> 0 Then%>
													
															<h3 align="center">Relat&aacute;rio de Vendas Mensal por Forma de Pagamento</h3>
															
															<form style="margin:0px" name="FILTRO" action="adm_relatorios_previsao_recebimento.asp">
															
																<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
																	
																	<tr>
																		
																		<td width="12%">
																		
																			<div align="right">
																			
																				<strong>In&iacute;cio / T&eacute;rmino:</strong>
																			
																			</div>
																			
																		</td>
																		
																		<td width="88%">
																		
																			<table width="100%" border="0" cellspacing="0" cellpadding="3">
																			
																				<tr>

																					<td width="18%">
																					
																						<div align="center">

																							<select name="MES_INICIO" style="font-size:13px">
																							
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
																							
																							<select name="ANO_INICIO" style="font-size:13px">
																								
																								<option value="<%=Year(Now)-3%>"><%=Year(Now)-3%></option>
																								
																								<option value="<%=Year(Now)-2%>"><%=Year(Now)-2%></option>
																								
																								<option value="<%=Year(Now)-1%>" selected="selected"><%=Year(Now)-1%></option>
																								
																								<option value="<%=Year(Now)%>"><%=Year(Now)%></option>
																							
																							</select>
																						
																						</div>
																					
																					</td>

																					<td width="11%">&nbsp;</td>

																					<td width="52%" rowspan="2">
																					
																						<div align="right">
																							
																							<input class="btnFiltrarAmarelo" onclick="Abre('#')" type="submit" name="filtro2" value=" " />
																						
																						</div>
																					
																					</td>
																				
																				</tr>

																				<tr>

																					<td>
																					
																						<div align="center">
																							
																							<select name="MES_FIM" style="font-size:13px">
																						
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
																						
																						</div>
																					
																					</td>

																					<td>
																					
																						<div align="center">Ano</div>
																					
																					</td>

																					<td>
																					
																						<div align="center">

																							<select name="ANO_FIM" style="font-size:13px">
																								
																								<option value="<%=Year(Now)-3%>"><%=Year(Now)-3%></option>
																								
																								<option value="<%=Year(Now)-2%>"><%=Year(Now)-2%></option>
																								
																								<option value="<%=Year(Now)-1%>"><%=Year(Now)-1%></option>
																								
																								<option value="<%=Year(Now)%>" selected="selected"><%=Year(Now)%></option>

																							</select>

																						</div>
																					
																					</td>

																					<td>&nbsp;</td>

																				</tr>

																			</table>
																		
																		</td>
																	
																	</tr>
																
																</table>
															
															</form>

														<%End If%>			
				
				
														<%If VALIDA_SUB29 <> 0 Then%>
													
															<h3 align="center">Relat&aacute;rio de Notas</h3>
															
															<form style="margin:0px" action="adm_relatorios_n_mensal.asp" target="_blank">
																
																<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
																	
																	<tr>
																		
																		<td width="12%">
																		
																			<div align="right">
																			
																				<strong>In&iacute;cio / T&eacute;rmino:</strong>
																			
																			</div>
																		
																		</td>
															
																		<td width="88%">
																		
																			<table width="100%" border="0" cellspacing="0" cellpadding="3">
																				
																				<tr>

																					<td width="18%">
																					
																						<div align="center">
																						
																							<select name="MES_INICIO" style="font-size:13px">
																						
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

																							<select name="ANO_INICIO" style="font-size:13px">
																								
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
																							
																							<input class="btnFiltrarAmarelo" onclick="Abre('#')" type="submit" name="filtro2" value=" " />
																						
																						</div>
																					
																					</td>

																				</tr>
																			
																			</table>
																		
																		</td>
																	
																	</tr>
																
																</table>
															
															</form>

														<%End If%>	
				
				
														<%If VALIDA_SUB13 <> 0 Then%>

															<h3 align="center">Fichas autenticadas  pela Diretoria</h3>
															
															<form style="margin:0px" name="FILTRO_FICHAS" action="adm_relatorios_fichas_autenticadas_diretoria.asp">
																	
																<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
																	
																	<tr>
																		
																		<td width="12%" style="border-bottom: dashed 1px #999999">
																		
																			<p align="right"><strong>Filtro R&aacute;pido:</strong></p>
																		
																		</td>
																		
																		<td width="88%" style="border-bottom: dashed 1px #999999">
																		
																			<table width="100%" border="0" cellspacing="0" cellpadding="3">
																				
																				<tr>
																					<td width="43%">
																					
																						<div align="left">&nbsp; &Uacute;ltimos 100 registros </div>
																						
																						<div align="center"></div>
																					
																					</td>

																					<td width="57%">
																					
																						<div align="right">
																							
																							<input class="btnFiltrarAmarelo" onclick="Abre('#')" type="submit" name="filtro1" value=" " />
																						
																						</div>
																						
																					</td>

																				</tr>

																			</table>
																		
																		</td>
																	
																	</tr>
																
																	<tr>
																		<td>
																		
																			<div align="right">
																			
																				<strong>Filtrar por data espec&iacute;fica:</strong>
																				
																			</div>
																			
																		</td>

																		<td>
																		
																			<table width="100%" border="0" cellspacing="0" cellpadding="3">
																		
																				<tr>
																					
																					<td width="17%">
																				
																						<div align="center">
																				
																							<input class="campoData" type="text" onfocus="pickDate(this,document.FILTRO_FICHAS.DATA_INI);" name="DATA_INI" value="<%=ArrumaData(DateAdd("m",-1,Date()))%>" onclick="" />
																						
																						</div>
																					
																					</td>
																					
																					<td width="5%">
																						
																						<div align="center">&agrave;</div>
																					
																					</td>
																			
																					<td width="21%">
																					
																						<div align="center">
																							
																							<input class="campoData" type="text" onfocus="pickDate(this,document.FILTRO_FICHAS.DATA_FIM);" name="DATA_FIM" value="<%=ArrumaData(Date())%>" onclick="" />
																						
																						</div>
																						
																					</td>

																					<td width="57%">
																					
																						<div align="right">
																							
																							<input class="btnFiltrarAmarelo" onclick="Abre('#')" type="submit" name="filtro2" value=" " />
																						
																						</div>
																					
																					</td>
																				
																				</tr>
																			
																			</table>
																		
																		</td>

																	</tr>

																</table>

															</form>

														<%end if%>
				
														<%
															dataFinalForm 		=	dataInicialForm
															dataInicialForm		=	DateAdd("m",-1,dataFinalForm)
														%>

														<%If VALIDA_SUB15 <> 0 Then%>

															<h3 align="center" style="margin-bottom:5px;">Relat&oacute;rio de Itens em posse do Vendedor</h3>
																
															<div align="center">

																<strong>(N&atilde;o contabiliza fichas, somente &aacute;lbuns, entre outros fabricados com anteced&ecirc;ncia como: Foto-telas, &Aacute;lbuns-digitais e entre outros.)</strong>
																
															</div>

															<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="80%" border="0" align="center" cellpadding="5" cellspacing="0">

																<tr>
																	<td width="12%">
																	
																		<div align="right">
																		
																			<strong>Escolha o vendedor:</strong>
																		
																		</div>
																	
																	</td>

																	<td width="88%">
																		
																		<form method="get" name="EXTRATO_BLOQUEADOS" action="adm_relatorios_extrato_vendedores_atrasados_por_vendedor_html3.asp" target="_blank" style="margin:0px">
																			
																			<table width="100%" border="0" cellspacing="0" cellpadding="3">
																				
																				<tr>
																					
																					<td width="43%">
																			
																						<div align="center">
																							
																							<div align="left">Vendedor</div>
																					
																						</div>
																					
																					</td>
																					
																					<td>Somente vendedores com itens em atraso</td>

																					<td width="14%" rowspan="2">
																					
																						<div align="right">
																							
																							<input class="btnFiltrarAmarelo" type="submit" name="filtro216" value=" "/>
																						
																						</div>
																					
																					</td>
																				
																				</tr>
																				
																				<tr>
																					
																					<td>
																					
																						<%AbrirConexao()%>                
																						
																						<div id="container_itens_em_posse_vendedor_todos">
																							
																							<%
																								Set rsVendedor = objConexao.Execute("SELECT COD_VENDEDOR, NOME_VENDEDOR FROM VW_MOVIMENTACOES_COMPLETA WHERE TIPO <> 1 AND DATA_RETORNO IS NULL AND DIAS_ATRASADOS >= 0 GROUP BY COD_VENDEDOR ORDER BY NOME_VENDEDOR")
																								
																							If not rsVendedor.Eof Then%>
																								
																								<select name="COD_VENDEDOR" class="inputTamanho3">
																								
																									<% 
																										While not rsVendedor.Eof
																									%>

																										<option value="<%=rsVendedor("COD_VENDEDOR")%>"><%=trocarAcento(rsVendedor("NOME_VENDEDOR"))%></option>
																									
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
																								Set rsVendedor = objConexao.Execute("SELECT COD_VENDEDOR, NOME_VENDEDOR FROM VW_MOVIMENTACOES_COMPLETA WHERE TIPO <> 1 AND DATA_RETORNO IS NULL GROUP BY COD_VENDEDOR ORDER BY NOME_VENDEDOR")
																								
																							If not rsVendedor.Eof Then%>
																								
																								<select name="COD_VENDEDOR_BLOQUEADO" class="inputTamanho3">
																									
																									<%
																										While not rsVendedor.Eof
																									%>
																									
																									<option value="<%=rsVendedor("COD_VENDEDOR")%>"><%=trocarAcento(rsVendedor("NOME_VENDEDOR"))%></option>
																									
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
																					
																					</td>
																					
																					<td>
																					
																						<input class="inputRadio" type="checkbox" onclick="trocaContainerRelatorioItensPosseVendedor(this.checked)" name="SOMENTE_VENDEDORES_EM_ATRASO" value="S" />
																					
																					</td>
																				
																				</tr>
																			
																			</table>
																		
																		</form>	
																		
																	</td>

																</tr>

															</table>

														<%end if%>
				
														<%If VALIDA_SUB31 <> 0 Then%>

															<h3 align="center">Extrato de aproveitamento de Produ&ccedil;&atilde;o</h3>

															<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
													
																<tr>
																	
																	<td width="12%">
																	
																		<div align="right">
																		
																			<strong>Filtro  por produ&ccedil;&atilde;o e vendedor </strong>
																		
																		</div>
																	
																	</td>
															
																	<td width="88%">
																		
																		<form name="EXTRATO_APROVEITAMENTO" action="adm_relatorios_extrato_por_producao_e_vendedor.asp" method="get" style="margin:0px" onsubmit="return validaForm7()"> 
																			
																			<table width="100%" border="0" cellspacing="0" cellpadding="3">
																			
																				<tr>
																					
																					<td width="77%">
																					
																						<div align="left">
																							
																							<select name="COD_PRODUCAO" id="COD_PRODUCAO" class="inputTamanho3">
																								
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
																							
																							<input class="btnFiltrarAmarelo" type="submit" name="filtro242" value=" " />
																						
																						</div>
																					
																					</td>
																				
																				</tr>
																			
																			</table>
																		
																		</form>
																	
																	</td>
																
																</tr>
															
															</table>
														
														
															<h3 align="center">Relat&oacute;rios - Retirada de album por periodo</h3>

															<form id="RELATORIO_RETIRADA" name="RELATORIO_RETIRADA" action="adm_relatorios_extrato_por_producao_e_vendedor_ano.asp" method="get" style="margin:0px" > 

																<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
																	
																	<tr>
																		
																		<td width="10%">
																		
																			<div align="left">
																			
																				<strong>Selecione o periodo </strong>
																			
																			</div>
																		
																		</td>

																		<td width="80%" colspan="2">
																			
																			<table width="100%" border="0" cellspacing="0" cellpadding="3">
																				
																				<tr>
																					
																					<td width="10%">
																				
																						<div align="left">						  
																							<input class="campoData" type="text" name="DATA_INICIO" onfocus="pickDate(this,this);" value="<%=ArrumaData(DateAdd("m",-12,Date()))%>" onclick="" />

																						</div>
																					
																					</td>
																			
																					<td width="5%">
																					
																						<div align="center">&agrave;</div>
																					
																					</td>

																					<td width="10%">

																						<div align="left">

																							<input class="campoData" type="text" name="DATA_TERMINO" onfocus="pickDate(this,this;" value="<%=ArrumaData(Date())%>" onclick="" />

																						</div>

																					</td>
																					
																					<td width="75%"></td>
																			
																				</tr>
																	
																			</table>
																		
																		</td>
																
																		<td width="10%">

																			<div align="right">

																				<input class="btnFiltrarAmarelo" type="submit" name="filtro242" value=" " />

																			</div>

																		</td>

																	</tr>

																</table>

															</form>		
															
														<%end if%>
				
														<%If VALIDA_SUB43 <> 0 Then%>

															<h3 align="center">Relat&oacute;rio Mensal de Produtos Vendidos - Valor Vendido</h3>

															<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
															
																<tr>

																	<td width="12%">
																	
																		<div align="right">
																		
																			<strong>M&ecirc;s / Ano</strong>
																			
																		</div>
																	
																	</td>

																	<td width="88%">

																		<form name="RELATORIO_MENSAL_PRODUTOS" action="adm_relatorios_extrato_mensal.asp" method="get" style="margin:0px"> 
																			
																			<table width="100%" border="0" cellspacing="0" cellpadding="3">

																				<tr>
																					
																					<td width="18%">
																					
																						<div align="center">

																							<select name="MES_INICIO" style="font-size:13px">

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

																							<select name="ANO_INICIO" style="font-size:13px">
																								
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
																							
																							<input class="btnFiltrarAmarelo" onclick="Abre('#')" type="submit" name="filtro2" value=" " />
																						
																						</div>
																					
																					</td>

																				</tr>
																			
																			</table>

																		</form>
																	
																	</td>
																
																</tr>
															
															</table>

														<%end if%>
				
				
														<%If VALIDA_SUB44 <> 0 Then%>

															<h3 align="center">Relat&oacute;rio de Venda de Produtos para Fornecedor</h3>

															<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
															
																<tr>

																	<td width="13%">
																	
																		<div align="right">
																		
																			<strong>Escolha o per&iacute;odo: </strong>
																		
																		</div>
																	
																	</td>

																	<td width="87%">
																		
																		<form method="get" name="RELATORIO_PRODUTOS" action="cap_relatorio_venda_produtos.asp" style="margin:0px">

																			<table width="100%" border="0" cellspacing="0" cellpadding="0">
																				
																				<tr>

																					<td>
																					
																						<table width="316" border="0" cellspacing="0" cellpadding="3">

																							<tr>
																								
																								<td width="34%">
																								
																									<div align="center">
																										
																										<input class="campoData" type="text" onfocus="pickDate(this,document.RELATORIO_PRODUTOS.DATA_INI);" maxlength="10" name="DATA_INI" value="<%=ArrumaData(DateAdd("m",-1,Date()))%>" onclick="" />
																									
																									</div>
																								
																								</td>
																								
																								<td width="5%">
																								
																									<div align="center">&agrave;</div>
																								
																								</td>
																					
																								<td width="34%">
																								
																									<div align="center">
																											
																										<input maxlength="10" class="campoData" type="text" onfocus="pickDate(this,document.RELATORIO_PRODUTOS.DATA_FIM);" name="DATA_FIM" value="<%=ArrumaData(Date())%>" onclick="" />
																								
																									</div>
																								
																								</td>

																								<td width="27%">&nbsp;</td>

																							</tr>
																						
																						</table>
																					
																					</td>

																					<td>
																					
																						<div align="right">
																							
																							<input class="btnFiltrarAmarelo" type="submit" name="filtro3" value=" " />
																						
																						</div>
																					
																					</td>

																				</tr>

																			</table>
																	
																		</form>	

																	</td>

																</tr>

															</table>

														<%end if%>

														<%If VALIDA_SUB88 <> 0 Then%>

															<h3 align="center">Relat&oacute;rio de Venda de Bolsa e Encaderna&ccedil;&atilde;o por Vendedor</h3>

															<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
															
																<tr>

																	<td width="13%">
																	
																		<div align="right">
																		
																			<strong>Escolha o per&iacute;odo: </strong>
																		
																		</div>
																	
																	</td>

																	<td width="87%">
																		
																		<form method="get" name="RELATORIO_PRODUTOS" action="cap_relatorio_venda_bolsa_encadernacao.asp" style="margin:0px">

																			<table width="100%" border="0" cellspacing="0" cellpadding="0">
																				
																				<tr>

																					<td>
																					
																						<table width="316" border="0" cellspacing="0" cellpadding="3">

																							<tr>
																								
																								<td width="34%">
																								
																									<div align="center">
																										
																										<input class="campoData" type="text" onfocus="pickDate(this,document.RELATORIO_PRODUTOS.DATA_INI);" maxlength="10" name="DATA_INI" value="<%=ArrumaData(DateAdd("m",-1,Date()))%>" onclick="" />
																									
																									</div>
																								
																								</td>
																								
																								<td width="5%">
																								
																									<div align="center">&agrave;</div>
																								
																								</td>
																					
																								<td width="34%">
																								
																									<div align="center">
																											
																										<input maxlength="10" class="campoData" type="text" onfocus="pickDate(this,document.RELATORIO_PRODUTOS.DATA_FIM);" name="DATA_FIM" value="<%=ArrumaData(Date())%>" onclick="" />
																								
																									</div>
																								
																								</td>

																								<td width="27%">&nbsp;</td>

																							</tr>
																						
																						</table>
																					
																					</td>

																					<td>
																					
																						<div align="right">
																							
																							<input class="btnFiltrarAmarelo" type="submit" name="filtro3" value=" " />
																						
																						</div>
																					
																					</td>

																				</tr>

																			</table>
																	
																		</form>	

																	</td>

																</tr>

															</table>

														<%end if%>
				
														<%If VALIDA_SUB45 <> 0 Then%>

															<h3 align="center">Relat&oacute;rio de Vendas do Grupo Repasse - Ranking</h3>

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
																							
																							<input class="btnFiltrarAmarelo" onclick="Abre('#')" type="submit" name="filtro2" value=" " />
																						
																						</div>
																					
																					</td>

																				</tr>
																			
																			</table>
																		
																		</form>
																	
																	</td>
																
																</tr>
															
															</table>
														
														<%end if%>
																
														<%If VALIDA_SUB30 <> 0 Then%>
															
															<h3 align="center">Relat&oacute;rio de Sa&iacute;da de Produtos do Estoque para Fornecedor </h3>
															
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
																					
																							<input class="btnFiltrarAmarelo" onclick="Abre('#')" type="submit" name="filtro2" value=" " />
																						
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

															</form>
															
														<%end if%>
				
				
														<%If VALIDA_SUB41 <> 0 Then%>

															<h3 align="center">Consulta de Cheques</h3>

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
																							
																							<input class="btnFiltrarAmarelo" type="submit" tabindex="7" name="filtro1" value=" " />
																						
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

															</table>

														<%End If%>
				
														<%If VALIDA_SUB49 <> 0 Then%>
						
															<h3 align="center">Relat&oacute;rio de Formandos por Produ&ccedil;&atilde;o</h3>
					
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
																							
																							<input class="btnFiltrarAmarelo" type="submit" name="filtro242" value=" " />
																						</div>
																					
																					</td>
																				</tr>

																			</table>

																		</form>
																	
																	</td>
																
																</tr>
															
															</table>
														
														<%end if%>

														<%If VALIDA_SUB77 <> 0 Then%>
															
															<h3 align="center">Relat&oacute;rio de Sa&iacute;da de produtos produzidos p&oacute;s-venda - Despacho</h3>

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
																										
																										<input class="campoData" type="text" onfocus="pickDate(this,document.RELATORIO_DESPACHO.DATA_INICIAL);" maxlength="10" name="DATA_INICIAL" value="<%=ArrumaData(DateAdd("m",-1,Date()))%>" onclick="" />
																									
																									</div>
																								
																								</td>

																								<td width="5%">
																							
																									<div align="center">&agrave;</div>
																								
																								</td>
																								
																								<td width="34%">
																								
																									<div align="center">
																										
																										<input maxlength="10" class="campoData" type="text" onfocus="pickDate(this,document.RELATORIO_DESPACHO.DATA_FINAL);" name="DATA_FINAL" value="<%=ArrumaData(Date())%>" onclick="" />
																									
																									</div>
																								
																								</td>
																								
																								<td width="27%">&nbsp;</td>
																							
																							</tr>

																						</table>
																					
																					</td>

																					<td>
																					
																						<div align="right">
																					
																					
																							<input class="btnFiltrarAmarelo" type="submit" name="filtro3" value=" " />
																						
																						</div>
																					
																					</td>

																				</tr>

																			</table>

																		</form>
																	
																	</td>
																	
																</tr>

															</table>

															<h3 align="center">Relat&oacute;rio de Sa&iacute;da de produtos produzidos p&oacute;s-venda - Entregues</h3>
															
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
																						
																							<input class="campoData" type="text" onfocus="pickDate(this,document.RELATORIO_ENTREGA.DATA_INICIAL);" maxlength="10" name="DATA_INICIAL_CADASTRO" value="" onclick="" />
																						
																						</div>
																					
																					</td>
																		
																					<td width="5%">
																					
																						<div align="center">&agrave;</div>
																					
																					</td>

																					<td width="34%">
																					
																						<div align="center">
																						
																							<input maxlength="10" class="campoData" type="text" onfocus="pickDate(this,document.RELATORIO_ENTREGA.DATA_FINAL);" name="DATA_FINAL_CADASTRO" value="" onclick="" />
																						
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
																								
																										<input class="campoData" type="text" onfocus="pickDate(this,document.RELATORIO_ENTREGA.DATA_INICIAL);" maxlength="10" name="DATA_INICIAL" value="<%=ArrumaData(DateAdd("m",-1,Date()))%>" onclick="" />
																								
																									</div>
																							
																								</td>

																								<td width="5%">
																									
																									<div align="center">&agrave;</div>
																									
																								</td>

																								<td width="34%">
																								
																									<div align="center">

																										<input maxlength="10" class="campoData" type="text" onfocus="pickDate(this,document.RELATORIO_ENTREGA.DATA_FINAL);" name="DATA_FINAL" value="<%=ArrumaData(Date())%>" onclick="" />
																									
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
																							
																							<input class="btnFiltrarAmarelo" type="submit" name="filtro3" value=" " />
																						
																						</div>
																						
																					</td>
																				
																				</tr>
																		
																			</table>
																		
																		</td>
																	
																	</tr>
																
																</table>
																
															</form>	

														<%END If

														If VALIDA_SUB49 <> 0 Then %>

															<h3 align="center">Relat&oacute;rio de Produtividade de Funcion&aacute;rios por OS</h3>

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
																										
																										<input class="campoData" type="text" onfocus="pickDate(this,document.RELATORIO_PRODUTOS.DATA_INI);" maxlength="10" name="DATA_INI" value="<%=ArrumaData(DateAdd("m",-1,Date()))%>" onclick="" />
																									
																									</div>
																								
																								</td>
																								
																								<td width="5%">
																								
																									<div align="center">&agrave;</div>
																								
																								</td>

																								<td width="34%">
																								
																									<div align="center">
																										
																										<input maxlength="10" class="campoData" type="text" onfocus="pickDate(this,document.RELATORIO_PRODUTOS.DATA_FIM);" name="DATA_FIM" value="<%=ArrumaData(Date())%>" onclick="" />
																									
																									</div>
																								
																								</td>
																		
																								<td width="27%">&nbsp;</td>
																							</tr>
																						
																						</table>
																					
																					</td>

																					<td>
																					
																						<div align="right">

																							<input class="btnFiltrarAmarelo" type="submit" name="filtro3" value=" " />

																						</div>
																					
																					</td>

																				</tr>

																			</table>

																		</form>
																	
																	</td>
																
																</tr>
																
															</table>

														<%END If%>
				
														<!-- Relat�rios -->

														<%If VALIDA_SUB78 <> 0 Then%>

															<h3 align="center">Relat&oacute;rio de Pedidos (Estoque Pr&eacute;-Venda)</h3>
															
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
																										
																										<input class="campoData" type="text" onfocus="pickDate(this,document.RELATORIO_ESTOQUE_PEDIDOS.data_inicial);" maxlength="10" name="data_inicial" value="<%=ArrumaData(DateAdd("m",-1,Date()))%>" onclick="" />
																									
																									</div>

																								</td>

																								<td>
																								
																									<div align="center">&agrave;</div>

																								</td>

																								<td>
																									
																									<div align="center">
																										
																										<input maxlength="10" class="campoData" type="text" onfocus="pickDate(this,document.RELATORIO_ESTOQUE_PEDIDOS.data_final);" name="data_final" value="<%=ArrumaData(Date())%>" onclick="" />
																									
																									</div>

																								</td>

																								<td></td>

																								<td></td>
																							
																							</tr>
																						
																						</table>
																					
																					</td>

																					<td>
																						<div align="right">

																							<input class="btnFiltrarAmarelo" type="submit" name="filtro3" value=" " />

																						</div>

																					</td>
																				
																				</tr>
																			
																			</table>

																		</form>

																	</td>
																
																</tr>

															</table>

														<%END If%>


														<%If VALIDA_SUB78 <> 0 Then%>

															<h3 align="center">Relat&oacute;rio de Pedidos (Estoque Pr&eacute;-Venda)</h3>
														
															<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
																
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
																										
																										<select name="tipo_relatorio" title="Relatório simples exibe apenas os pedidos, já o detalhado exibe também seu itens.">
																											
																											<option value="S">Simples </option>
																											
																											<option value="D"> Detalhado </option>
																										
																										</select>
																									
																									</div>

																								</td>									
																							</tr>	

																							<tr>
																								
																								<td width="">
																									
																									<div align="left" title="Periodo no qual foi feito a solicitação do pedido.">
																										
																										<strong>Escolha o per&iacute;odo: </strong>
																									
																									</div>
																								
																								</td>
																		
																								<td>
																									
																									<div align="center">
																										
																										<input class="campoData" type="text" onfocus="pickDate(this,document.RELATORIO_ESTOQUE_PEDIDOS.data_inicial);" maxlength="10" name="data_inicial" value="<%=ArrumaData(DateAdd("m",-1,Date()))%>" onclick="" />
																									
																									</div>
																								
																								</td>
																								
																								<td>
																									
																									<div align="center">&agrave;</div>
																									
																								</td>
																								
																								<td>
																									
																									<div align="center">
																													
																										<input maxlength="10" class="campoData" type="text" onfocus="pickDate(this,document.RELATORIO_ESTOQUE_PEDIDOS.data_final);" name="data_final" value="<%=ArrumaData(Date())%>" onclick="" />
																									
																									</div>

																								</td>

																								<td></td>

																								<td></td>

																							</tr>

																						</table>
																					</td>

																					<td>

																						<div align="right">

																							<input class="btnFiltrarAmarelo" type="submit" name="filtro3" value=" " />

																						</div>

																					</td>

																				</tr>

																			</table>

																		</form>
																	</td>

																</tr>

															</table>

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
																								<div align="left" title="Periodo no qual foi feita a moveimenta��o">
																									<strong>Escolha o per&iacute;odo: </strong>
																								</div>
																							</td>
																							<td width="120px;">
																								<div align="left">
																									<input class="campoData" type="text" onfocus="pickDate(this,document.RELATORIO_ESTOQUE_PEDIDOS.data_inicial);" maxlength="10" name="f_periodo_inicio" value="<%=ArrumaData(DateAdd("m",-1,Date()))%>" onclick="" />
																								</div>
																							</td>
																							<td width="15px;">
																								<div align="center">
																									&agrave;
																								</div>
																							</td>
																							<td width="120px;">
																								<div align="center">
																									<input maxlength="10" class="campoData" type="text" onfocus="pickDate(this,document.RELATORIO_ESTOQUE_PEDIDOS.data_final);" name="f_periodo_fim" value="<%=ArrumaData(Date())%>" onclick="" />
																								</div>
																							</td>
																							<td width="120px;">
																								<div align="left" title="Periodo no qual foi feita a moveimenta��o">
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
																						<input class="btnFiltrarAmarelo" type="submit" name="filtro3" value=" " />
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

															<h3 align="center">Relat&oacute;rio para retirada de produ&ccedil;&atilde;o</h3>
														
															<table style="border:#999999 dashed 1px; padding:1px; background:#FFFFCC; width: 80%; margin: 0 auto; box-sizing: border-box;" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr>
    <td width="100%">
      <form target="_BLANK" method="get" action="adm_relatorios_vendedores_historico.asp" style="margin:0px">
        <style>
          select {
            width: 100%;
            max-width: 100%;
            box-sizing: border-box;
          }
        </style>
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
																						
																							<input class="btnFiltrarAmarelo" type="submit" name="filtro3" value=" " />
																						
																						</div>

																					</td>
																				
																				</tr>
																			
																			</table>
																		
																		</form>
																	
																	</td>
																
																</tr>
															
															</table>

														<% end if %>


														<h3 align="center">Relat&oacute;rio de hist&oacute;rico de vendedores por produ&ccedil;&atilde;o </h3>
															
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
																						
																						<input class="btnFiltrarAmarelo" type="submit" name="filtro3" value=" " />

																					</div>

																				</td>
																			</tr>
																		
																		</table>

																	</form>

																</td>

															</tr>

														</table>

														<!-- fim relat�rios -->        
														<p align="center">&nbsp;</p>

														<br />
														<br />
														<br />
														<br />
														<br />
														<br />
														<br />
														<br />

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
