<script type="text/javascript" src="js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="js/jquery.price_format.1.6.js"></script>
<script type="text/javascript" src="js/jquery.maskedinput-1.3.min.js"></script>
<script type="text/javascript" src="js/chosen.jquery.min.js"></script>
<script type="text/javascript" src="files/js/scrips_calendario/dhtmlSuite-common.js"></script>
<script type="text/javascript" src="js/modal-message.js"></script>
<script type="text/javascript" src="js/ajax-dynamic-content.js"></script>
<script type="text/javascript" src="js/ajax.js"></script>
<script type="text/javascript" src="js/AC_RunActiveContent.js"></script>



 
<!--###################   MENSAGEM DHTML   ###################--> 
<script language="javascript" type="text/javascript"> 
//Iniciar de DHTML Mensagem 
messageObj = new DHTML_modalMessage();
messageObj.setShadowOffset(5);

function displayMessage(url)
{
	
	messageObj.setSource(url);
	messageObj.setCssClassMessageBox(false);
	messageObj.setSize(500,280); //Não alterar tamanho!!!
	messageObj.setShadowDivVisible(false);	// Enable shadow for these boxes
	messageObj.display();
}

function paginacaoIrParaPagina(caminho,pagina_anterior,pagina_nova,qtd_paginas){	
	
	var pag = "?paginacao=";
	var caminho_novo;
	var a = new Number(pagina_nova);
	var b = new Number(qtd_paginas);
	if(pagina_anterior !== pagina_nova){
		if(a > b){ 
			caminho_novo = pag+qtd_paginas+caminho;
		}else{
			caminho_novo = pag+pagina_nova+caminho;
		}
		//displayCarregando();
		abre(caminho_novo);
		//setTimeout("abre("+caminho_novo+")",500);
	}
	
}


function displayCarregando()
{
	exibeOculta('ajaxCarregandoContainer','exibir');
}

function displayCarregandoClose()
{
	exibeOculta('ajaxCarregandoContainer','ocultar');
}

function displayStaticMessage(messageContent,cssClass)
{
	messageObj.setHtmlContent(messageContent);
	messageObj.setSize(500,280); //Não alterar tamanho!!!
	messageObj.setCssClassMessageBox(cssClass);
	messageObj.setSource(false);	// no html source since we want to use a static message here.
	messageObj.setShadowDivVisible(false);	// Disable shadow for these boxes	
	messageObj.display();
	
	
}


function closeMessage()
{
	messageObj.close();	
}

function info(mensagem){
//Página de informações rápidas.
	alert('INFORMAÇÕES!\n'+mensagem);
}



</script>
<!--##########################################################-->

<!--###################   INICIO MASCARAS   ###################-->
<!--MOEDA-->
<script language="javascript" type="text/javascript">
function formatarMoeda(fld, e) {   
var milSep = "."
var decSep = ","
var sep = 0;   
var key = '';   
var i = j = 0;   
var len = len2 = 0;   
var strCheck = '0123456789';   
var aux = aux2 = '';  	
var whichCode = (window.event)?event.keyCode:e.which; 
//alert(whichCode)
if (whichCode == 13 || whichCode == 8 || whichCode == 0 || whichCode == 9) return true;  // Enter   
key = String.fromCharCode(whichCode);  // recebe o valor da chave vinda da chave do código  
if (strCheck.indexOf(key) == -1) return false;  // Chave não válida  
len = fld.value.length;   
for(i = 0; i < len; i++)   
if ((fld.value.charAt(i) != '0') && (fld.value.charAt(i) != decSep)) break;   
aux = '';   
for(i = 0; i < len; i++)   
if (strCheck.indexOf(fld.value.charAt(i))!=-1) aux += fld.value.charAt(i);   
aux += key;   
len = aux.length;   
if (len == 0) fld.value = '';   
if (len == 1) fld.value = '0'+ decSep + '0' + aux;   
if (len == 2) fld.value = '0'+ decSep + aux;   
if (len > 2) {   
aux2 = '';   
for (j = 0, i = len - 3; i >= 0; i--) {   
if (j == 3) {   
aux2 += milSep;   
j = 0;   
}   
aux2 += aux.charAt(i);   
j++;   
}   
fld.value = '';   
len2 = aux2.length;   
for (i = len2 - 1; i >= 0; i--)   
fld.value += aux2.charAt(i);   
fld.value += decSep + aux.substr(len - 2, len);   
}   
return false;   
}  
</script>
<!--FIM MOEDA-->

<%
If calendario_necessario = "SIM" Then
%>
<!--CALENDÁRIO-->
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
		references.myDate.focus();
	}	
</script>
<!--FIM CALENDÁRIO-->
<%End If%>
<!--DEMAIS MASCARAS-->
<script language="JavaScript">
// JavaScript Document
//adiciona mascara de cnpj
function MascaraCNPJ(cnpj){
    if(mascaraInteiro(cnpj)==false){
        event.returnValue = false;
    }    
    return formataCampo(cnpj, '000.000.000/0000-00', event);
}

//adiciona mascara de cep
function MascaraCep(cep){
        if(mascaraInteiro(cep)==false){
        event.returnValue = false;
    }    
    return formataCampo(cep, '00000-000', event);
}

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
        alert('Número de Telefone Inválido!');
}

//valida CEP
function ValidaCep(cep){
    exp = /\d{2}\\d{3}\-\d{3}/
    if(!exp.test(cep.value))
        alert('Numero de Cep Inválido!');        
}

//valida data
function validarData(campo){
	if (campo.value !== "" && campo.value !== "__/__/____"){

var ano_bissexto;
var expReg = /^(([0-2]\d|[3][0-1])\/([0]\d|[1][0-2])\/[1-2][0-9]\d{2})$/;
var msgErro = 'Formato inválido de data.\nInsira novamente';
if ((campo.value.match(expReg)) && (campo.value!='')){
var dia = campo.value.substring(0,2);
var mes = campo.value.substring(3,5);
var ano = campo.value.substring(6,10);

if (((!(ano%4)) && (ano%100)) || (!(ano%400))){
	ano_bissexto = true;
}else{
	ano_bissexto = false;
}

if((mes==4 || mes==6 || mes==9 || mes==11) && dia > 30){
campo.focus();
alert("Dia incorreto ! O mês especificado contém no máximo 30 dias.\nInsira novamente");
campo.value = "";
return false;
} else{
if(!ano_bissexto && mes==2 && dia>28){
campo.focus();
alert("Data incorreta! O mês especificado contém no máximo 28 dias.\nInsira novamente");
campo.value = "";
return false;
} else{
if(ano_bissexto && mes==2 && dia>29){
campo.focus();
alert("Data incorreta! O mês especificado contém no máximo 29 dias.\nInsira novamente");
campo.value = "";
return false;
} else{ 
return true;
}}}} else {
alert(msgErro);
campo.focus();
campo.value = "";
return false;
}
	}
}

//valida o CPF digitado
function ValidarCPF(Objcpf){

	var txt = Objcpf.value;

	if(Objcpf.value !== ''){
	//Valida pela quantidade de caracteres
		if(txt.length !== 14){
			Objcpf.value = '';
			Objcpf.focus();
			alert('Preencha o campo '+Objcpf.name+' por completo!');
		}
	}

	/*
	INATIVADO ATÉ CORREÇÃO DO MESMO
    var cpf = Objcpf.value;
	
	if(cpf !== ""){
	
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
    if(digitoGerado!=digitoDigitado){    
		alert('CPF Inválido\nInsira novamente.');
		Objcpf.value = "";
	}
	
	}
	*/  
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
	
	if(cnpj !== "" && cnpj !== "__.___.___/____-__"){
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
    
    if(((dig1*10)+dig2) != digito){
		alert('CNPJ Inválido\nInsira novamente.');
		ObjCnpj.value = "";
	}
	
	}
	
		
		   
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
<!--FIM MASCARAS-->
<script>

function mascara(o,f){
    v_obj=o
    v_fun=f
    setTimeout("execmascara()",1)
}

function execmascara(){
    v_obj.value=v_fun(v_obj.value)
}

function leech(v){
    v=v.replace(/o/gi,"0")
    v=v.replace(/i/gi,"1")
    v=v.replace(/z/gi,"2")
    v=v.replace(/e/gi,"3")
    v=v.replace(/a/gi,"4")
    v=v.replace(/s/gi,"5")
    v=v.replace(/t/gi,"7")
    return v
}

function soNumeros(v){
    return v.replace(/\D/g,"")
}

function soLetras(e){
	var tecla=(window.event)?event.keyCode:e.which;
	if (tecla == 9 || tecla ==0 || tecla ==64){
		return true;
	}else{		
		if((tecla > 64 && tecla < 91)||(tecla > 96 && tecla < 123)) return true;
		else{
		if (tecla != 8) return false;
		else return true;
		}
	}
}

function soNumerosLetras(e){
	var tecla=(window.event)?event.keyCode:e.which;
	if (tecla == 9 || tecla ==0){
		return true;
	}else{
		if((tecla > 47 && tecla < 58)||(tecla > 64 && tecla < 91)||(tecla > 96 && tecla < 123) || tecla == 45 || tecla == 95) return true;
		else{
		if (tecla != 8) return false;
		else return true;
		}
	}
}

function telefone(v){
    v=v.replace(/\D/g,"")
    v=v.replace(/^(\d\d)(\d)/g,"($1) $2") 
    v=v.replace(/(\d{4})(\d)/,"$1-$2")   
    return v
}

function cpf(v){

    v=v.replace(/\D/g,"")                   
    v=v.replace(/(\d{3})(\d)/,"$1.$2")      
    v=v.replace(/(\d{3})(\d)/,"$1.$2")      
                                            
    v=v.replace(/(\d{3})(\d{1,2})$/,"$1-$2")
    return v
}

function cep(v){
    v=v.replace(/\D/g,"")               
    v=v.replace(/^(\d{5})(\d)/,"$1-$2") 
    return v
}

function cnpj(v){
    v=v.replace(/\D/g,"")                           
    v=v.replace(/^(\d{2})(\d)/,"$1.$2")             
    v=v.replace(/^(\d{2})\.(\d{3})(\d)/,"$1.$2.$3") 
    v=v.replace(/\.(\d{3})(\d)/,".$1/$2")         
    v=v.replace(/(\d{4})(\d)/,"$1-$2")              
    return v
}
function data(v){
        v=v.replace(/\D/g,"") 
        v=v.replace(/(\d{2})(\d)/,"$1/$2") 
        v=v.replace(/(\d{2})(\d)/,"$1/$2") 
        return v
}
function data2(v){
        v=v.replace(/\D/g,"") 
        v=v.replace(/(\d{2})(\d)/,"$1/$2")
		v=v.replace(/(\d{6})(\d)/,"$1/$2") 
        return v
}

</script>
<!--DEMAIS MASCARAS-->
<!--###################   FIM MASCARAS   ###################-->
<!--###################   VARIADOS   ###################-->
<script language="javascript" type="text/javascript">
function listaCidades(uf,cod_cidade,div,coluna){
	if(uf !== ""){
		var caminho = "files/ajax/cidades.asp?uf="+uf+"&cod_cidade="+cod_cidade+"&coluna="+coluna;
		ajax_loadContent(div,caminho,false,false,false);
	}
}

function focusJanela(){
	window.focus();
}

function naoNullo(campo,e){
	if(campo.value !== ""){
		eval(e);
	}
}

function exibeOcultaDinamico(div)
	{
		var o = document.getElementById(div);
		
		
		if(o){
			if (o.style.display == "none" || o.style.display == "")
				{
					o.style.visibility = "visible";
					o.style.display = "inline";
				}
				else
				{
					o.style.visibility = "hidden";
					o.style.display = "none";
				}
		}	
}
	
function exibeOculta(div,acao){

	var o = document.getElementById(div);
	
	if(o){
		if(acao == "ocultar"){
			//Oculta
				o.style.visibility = "hidden";
				o.style.display = "none";
		}else{
			//Exibe
				o.style.visibility = "visible";
				o.style.display = "inline";		
		}
	}
}

//Exemplo: no checkbox habilitaComboBox('COLUNA_DATA_NASCIMENTO','FILTRO_MES')
function habilitaComboBox(obj,combo)
{
	if (document.getElementById(obj).checked)
	{
		document.getElementById(combo).disabled = "";
	}
	else
	{
		document.getElementById(combo).disabled = "disabled";
		document.getElementById(combo)[0].selected = true;
	}
}

function tabsMostra(idTabConteudo,qtdTabs,numeroTabAtual){
	 for (i = 1; i < qtdTabs+1; i++) {
		document.getElementById('tab'+i).style.visibility = "hidden";
		document.getElementById('tab'+i).style.display = "none";
		document.getElementById("liTab"+i).className="";	
	 }
		
	document.getElementById("liTab"+numeroTabAtual).className="current";
	document.getElementById(idTabConteudo).style.visibility = "visible";
	document.getElementById(idTabConteudo).style.display = "inline";						
}

function mOverOut(campo,cor,tipo){

	if (tipo == "over"){
		campo.style.cursor = "pointer";
	}else{
		campo.style.cursor = "";
	}
	campo.style.backgroundColor = cor	
}

function botaoConteudo(campo,tipo,codigo){
	
	var e = document.getElementById("botaoConteudoE_"+codigo);
	var c = document.getElementById("botaoConteudoC_"+codigo);
	var d = document.getElementById("botaoConteudoD_"+codigo);
	
	if (tipo == "over"){
		campo.style.cursor = "pointer";
		
		e.style.backgroundImage ="url(images/layout/botaoConteudoLateralEsquerdaHover.png)";
		c.style.backgroundImage ="url(images/layout/botaoConteudoLateralCentroHover.png)";
		d.style.backgroundImage ="url(images/layout/botaoConteudoLateralDireitaHover.png)";
		
	}else{
		campo.style.cursor = "";
		e.style.backgroundImage ="url(images/layout/botaoConteudoLateralEsquerda.png)";
		c.style.backgroundImage ="url(images/layout/botaoConteudoLateralCentro.png)";
		d.style.backgroundImage ="url(images/layout/botaoConteudoLateralDireita.png)";
	}	
}

function focusCampo(campo){
	campo.focus();
}

function limpaAbre(url){
	
	if(url !== ''){
		parent.location.href = url;
	}
}

function abrir(theURL,winName,features) 
{ 
  window.open(theURL,winName,features);
}

function abre(url){
	document.location = url;
}

function abreNovaJanela(url,nomeJanela,parametros){	
	if (nomeJanela == "" && parametros == ""){
		window.open(url); //_blank
	}else{
		window.open(url,nomeJanela,parametros);
	}
}

function limparCampo(form,name){
	document[form][name].value = '';
}

function limparRadio (form,name){
	var i;
	var qtd = 0;
	if(document[form]){
		if(document[form][name]){
			qtd = qtd+(document[form][name].length);
			if(qtd > 0){
				for(i=0;i<qtd;i++){
					document[form][name][i].checked = false;
				}
			}
		}
	}
}

function abreConteudo(pagina,submenu,tituloPagina){
	//parent.topo.mudarTitulo(tituloPagina);
	//parent.topo.document.getElementById('tituloPaginas').innerHTML = tituloPagina;
	parent.conteudo.document.location = pagina;
}
</script>
<!--###################    VALIDAÇÕES    ###################-->
<script language="javascript" type="text/javascript">
function textCounter(campo, countcampo, maxlimit){
    	if (campo.value.length > maxlimit) {
		    campo.value = campo.value.substring(0, maxlimit);
    	}
    	else {
            countcampo.value = maxlimit - campo.value.length;
    	}
 }
 
function validaForm(campos_validados,formulario_validacao) {

	var str = campos_validados;
	var formulario = formulario_validacao;
	var d = formulario;
    var arrayPrincipal = str.split('|');
	var radio, r, rvalidado;
	var msg = "";
	var classOriginal;
	var temp;
		
    for (i = 0; i < arrayPrincipal.length; i++) {
                        subArray = arrayPrincipal[i].split('#');                         
                        for (t = 0; t < subArray.length; t+=3) {
						
							if(subArray[t+2] < 3){
								temp = new String(formulario[subArray[t]].value);
									if(temp == ""){
											msg += "   - "+subArray[t+1]+'\n';
											
											$(formulario[subArray[t]]).css("border","1px #CC0000 solid");
											$(formulario[subArray[t]]).css("background-color","#FECBCB");
											$(formulario[subArray[t]]).addClass("formNaoValidado");
									}else{
											$(formulario[subArray[t]]).css("border","1px solid #666666");
											$(formulario[subArray[t]]).css("background-color","#F4F4F4");
											$(formulario[subArray[t]]).removeClass("formNaoValidado");											
									}
							}else if(subArray[t+2] == 3){
								radio = formulario[subArray[t]].length;
								rvalidado = false;
								for(r=0;r<radio;r++){
									if(formulario[subArray[t]][r].checked){
										rvalidado = true;
									}
								}
								if(!rvalidado){
										msg += "   - "+subArray[t+1]+'\n';
								}								
							}
                        }//End for
     }
	 if(msg == ""){
	 	return true
	 }else{
		msg = msg+" ";
	 	alert('Verifique o(s) campo(s) obrigatórios(s) antes de continuar:\n\n'+msg);
	 	return false
	 } 
}
</script>
<!--CONFIRMAÇÃO SIM/NÃO-->
<script language="javascript">
/*Exemplo
if(YNconfirm(msg)){

}
*/
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
        