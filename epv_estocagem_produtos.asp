<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim COD_SALA_ARMAZENADO, COD_SALA_ARMAZENADO_record_set
COD_SALA_ARMAZENADO = formatNumber(0&Request.QueryString("COD_SALA_ARMAZENADO"),0)
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
<!--
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

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->

function cadastraVendedor(valor){
	alert(valor);
}
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
<script language="javascript">
function focoCx(campo){
	campo.style.backgroundColor = '#FFFF66'
}
function blurCx(campo){
	campo.style.backgroundColor = '';
}
function carregaCampo(){
<%If COD_SALA_ARMAZENADO > 0 Then%>
	document.FORMULARIO.COD_CAIXA_ARMAZENADO.focus();
<%else%>	
	document.FORMULARIO.COD_SALA_ARMAZENADO.focus();
<%End If%>
}
</script>
<script>
function confirmaDevolucao(nomeVendedor) {
	var msg
	msg = "Tem certeza que deseja gravar a devolução destes itens para o vendedor(a): "+nomeVendedor
	if(YNconfirm(msg)){
			return true;
			//document.DEVOLUCAO.submit();			
	}else{
			return false;
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

<!--INICIO MASCARAS-->
<script language="JavaScript">
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
function ValidaData(data){
    exp = /\d{2}\/\d{2}\/\d{4}/
    if(!exp.test(data.value))
        alert('Data Invalida!');            
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
    v.replace(/\D/g,"")               
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

function selecionaCaixa(CAMPO){
	if (CAMPO.value > 0){
		document.location.href = "epv_estocagem_produtos.asp?COD_SALA_ARMAZENADO="+CAMPO.value;
	}else{
		document.location.href = "epv_estocagem_produtos.asp";
	}

}
function validaForm(){
d = document.FORMULARIO
	if (d.COD_SALA_ARMAZENADO.value == 0){
		d.COD_SALA_ARMAZENADO.focus();
		alert('Selecione a sala');
		return false;
	}else if (d.COD_CAIXA_ARMAZENADO.value == 0){
		d.COD_CAIXA_ARMAZENADO.focus();
		alert('Selecione a caixa');
		return false;
	}else {
		return true;
	}
}
</script>
<body onload="carregaCampo()">

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

	  <h2>EPV - Estoque / Estocagem de Produtos</h2>

		<h3 align="center">Estocagem de produtos</h3>
        <div class="msg" align="center">
	    <p><%=Request.QueryString("msg")%></p>
	  </div>
		<form action="epv_estocagem_produtos_itens.asp" name="FORMULARIO" onsubmit="return validaForm()">
		  <div align="center">
						  <p>
								 
						  </p>		
						  <table width="98%" border="0" cellspacing="0" cellpadding="5">
                            <tr>
                              <td width="51%"><div align="right"><strong>Selecione a sala: </strong></div></td>
                              <td width="49%"><select name="COD_SALA_ARMAZENADO" onchange="selecionaCaixa(this)">
                              <option value="0">Selecione a sala...</option>
                                 	<%
									AbrirConexao()
									Set rs = objConexao.Execute("SELECT ANDAR FROM TB_LOCAL_ARMAZENAGEM_SALA GROUP BY ANDAR ORDER BY ANDAR")
									
									Dim ANDAR
									
									While not rs.Eof
									
									ANDAR = formatNumber(0&rs("ANDAR"),0)
									
											  If ANDAR = 1 Then
												ANDAR = "(1) - Sub-Solo"
											  Elseif ANDAR = 2 Then
												ANDAR = "(2) - Térreo"
											  Elseif ANDAR = 3 Then
												ANDAR = "(3) - 1º Andar"
											  End If
									%>
                                    <optgroup label="<%=ANDAR%>">
               						 <%
									Set rs2 = objConexao.Execute("SELECT * FROM TB_LOCAL_ARMAZENAGEM_SALA WHERE ANDAR='"&rs("ANDAR")&"' ORDER BY NOME_SALA_ARMAZENADO")
									While not rs2.Eof
									COD_SALA_ARMAZENADO_record_set = formatNumber(0&rs2("COD_SALA_ARMAZENADO"),0)
									%>
									  <option <%if COD_SALA_ARMAZENADO_record_set = COD_SALA_ARMAZENADO Then%> selected="selected" <%End If%> value="<%=rs2("COD_SALA_ARMAZENADO")%>"><%=rs2("COD_SALA_ARMAZENADO")%> - <%=rs2("NOME_SALA_ARMAZENADO")%></option>
			 						<%
									rs2.MoveNext()
									Wend
									
									rs2.Close
									Set rs2=Nothing
									%>
                                    
                                     </optgroup>
									<%
									rs.MoveNext()
									Wend
									
									FecharConexao()
									%>
                                            </select></td>
                              </tr>
                               <%If COD_SALA_ARMAZENADO > 0 Then%>
                            <tr>
                              <td><div align="right"><strong>Selecione a caixa</strong></div></td>
                              <td>
							 
                              	<select name="COD_CAIXA_ARMAZENADO">
                                <option value="0">Selecione a caixa...</option>
                                <%
								AbrirConexao()
								
								Set rs = objConexao.Execute("SELECT COLUNA FROM TB_LOCAL_ARMAZENAGEM_SALA_CAIXA WHERE COD_SALA_ARMAZENADO = "&COD_SALA_ARMAZENADO&" GROUP BY COLUNA ORDER BY COLUNA")
												
									While not rs.Eof
									%>
								    <optgroup label="Coluna <%=rs("COLUNA")%>">
                                    <%
										 Set rs2=objConexao.Execute("SELECT * FROM TB_LOCAL_ARMAZENAGEM_SALA_CAIXA WHERE COD_SALA_ARMAZENADO = "&COD_SALA_ARMAZENADO&" AND COLUNA = '"&rs("COLUNA")&"' ORDER BY NUMERO_CAIXA ASC")
										 While not rs2.Eof
										 %>
                                    <option value="<%=rs2("COD_CAIXA_ARMAZENADO")%>"><%=rs2("NUMERO_CAIXA")%></option>
                                        
                                        <%
										rs2.MoveNext()
										Wend
										
										rs2.Close
										Set rs2=Nothing
										
										%>
                                    </optgroup>
                                        
                                	<%
									rs.MoveNext()
									Wend
									
								rs.Close
								Set rs=Nothing
								FecharConexao()
								%>
                                </select></td>
                              </tr>
                            <tr>
                              <td colspan="2"><div align="center"> <br />
                              <input type="hidden" name="comando" value="nova" />
                              <input class="btnAvancar" name="btnAvancar" type="submit" value="" />
                              </div></td>
                              </tr>
                              <%End If%>
                          </table>
						  <p>&nbsp;</p>
                           <p>&nbsp;</p>
                            <p>&nbsp;</p>
						</div>
		</form>	
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