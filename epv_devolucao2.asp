<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!--
'*** Instruções
'1) 1 Grava na temporária os códigos informados e abaixo informa de quem corresponde
'2) Após informar todos clica em avançar para concluir devolução

-->
<%
Dim comando, TXT_DEVOLUCAO, ERROR_LINHA, pronto, CONT_ITENS
pronto = Request.QueryString("pronto")
ERROR_LINHA = 1
comando = Request.QueryString("comando")
TXT_DEVOLUCAO = Request.Form("TXT_DEVOLUCAO")
TXT_DEVOLUCAO = Replace(TXT_DEVOLUCAO," ","")
CONT_ITENS = 0

If comando = "nova" Then
'Apaga tudo da base do determinado usuário
AbrirConexao()
	objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_DEVOLUCAO_TEMPORARIA WHERE COD_USUARIO = "&Session("COD_USUARIO"))
FecharConexao()
End If

If comando = "excluir" Then
	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_DEVOLUCAO_TEMPORARIA WHERE COD_ETIQUETA = "&Request.QueryString("cod_etiqueta"))
		Response.Redirect("epv_devolucao.asp")
	FecharConexao()
End If

If comando = "efetuar_devolucao" Then
'Faz os UPDATES com a data de hoje,
		AbrirConexao()
		Dim rsUp, COD_FORMANDO_UP, TIPO_UP
		Dim COD_MOVIMENTACAO_UP
		
		Set rsUp = objConexao.Execute("SELECT COD_ETIQUETA FROM TB_FICHAS_IDENTIFICACAO_DEVOLUCAO_TEMPORARIA WHERE COD_USUARIO="&Session("COD_USUARIO")&" GROUP BY COD_ETIQUETA")

			
		
			While not rsUP.Eof
			COD_MOVIMENTACAO_UP = 0
			TIPO_UP				= Right(rsUP("COD_ETIQUETA"),1)
			COD_FORMANDO_UP 	= Mid(rsUP("COD_ETIQUETA"),1,Len(rsUP("COD_ETIQUETA"))-1)
			COD_MOVIMENTACAO_UP	= Request.Form("COD_MOVIMENTACAO_"&COD_FORMANDO_UP&"_"&TIPO_UP)
				 
				 sql = "UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET DATA_RETORNO = NOW( ) WHERE "&_
									"COD_MOVIMENTACAO ="&COD_MOVIMENTACAO_UP&" LIMIT 1"
									'Response.Write(sql)
				 objConexao.Execute(sql)
			
			rsUp.MoveNext()
			Wend
		
		rsUp.Close()
		Set rsUp = Nothing

	Response.Redirect("epv_devolucao_imprimir.asp")
End If
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

function qtdItensContainer(qtd){
	document.getElementById('qtdItens').innerHTML = '<strong>Itens selecionados: <span class="textoMaior">'+qtd+'</span></strong>';
}

</script>
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
  
  <!-- include file="includes/estoque_bip.asp" -->
	
       <script type="text/javascript">
		function beep(tag)
        {
        	return document.getElementById(tag);
        }
		
		function gravarItem(){
			var etiqueta = $("#TXT_DEVOLUCAO").val();
			var caminho = 'epv_movimentacoes_estoque_ajax.asp?ts=incluir&p1='+etiqueta;
			ajax_processa(caminho,false);
			
		}
		
		function carregaListagem(){
			var caminho = 'epv_movimentacoes_estoque_ajax.asp?ts=listar';
			ajax_loadContent('containerListagem',caminho,false,false,false);
			$("#TXT_DEVOLUCAO").focus();
		}
		
		
		function displayLeitura(tipo,msg){
			$('#sucesso_msg').hide();
			$('#erro_msg').hide();
			
			if(tipo == 'erro'){
				$('#erro_msg').show();
				$('#erro_msg').html(msg)	
			}else{
				$('#sucesso_msg').show();
				$('#sucesso_msg').html(msg);
				beep('beep_general').play(1);	
			}
			
			$("#TXT_DEVOLUCAO").val('');		
		}
        </script>
        <object id="beep_general" class="flash" type="application/x-shockwave-flash" data="flash/flashbeep_general.swf" width="1" height="1">
			<param name="movie" value="flash/flashbeep_general.swf" />
		</object>
  
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">

	  <h2>EPV - Devolu&ccedil;&atilde;o de Produtos </h2>
	  <div class="subMenu">
Sub-Menu: <a href="epv_devolucao.asp?comando=nova">Iniciar devolu&ccedil;&atilde;o de Vendedor</a> | <a href="epv_saida_inicio.asp?comando=nova">Sa&iacute;da de &Aacute;lbuns do Estoque</a></div>
	  <div class="msg" align="center">
	    <p><%=Request.QueryString("msg")%></p>
	  </div>
      
      	  	  
		
			  <table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td style="padding:5px;">
    <div class="msgGrandeSucesso" style="display:none" id="sucesso_msg"></div>
    <div class="msgGrandeErro" style="display:none"  id="erro_msg"></div>
    <div align="center">
      <input onkeypress="mascara(this,soNumeros)" maxlength="8" onfocus="focoCx(this)" onblur="blurCx(this)" style="height:50px; width:200px; font-size:34px; border:" type="text" id="TXT_DEVOLUCAO" />
    </div></td>
    </tr>
  <tr>
    <td><div align="center">
      <input class="btnGravar " type="button" onclick="gravarItem()" id="botaoGravar" value=" " />
      <br />
<br />
    </div></td>
    </tr>
</table>

			
		<%
		AbrirConexao()
		%>	
		<form action="epv_devolucao.asp?comando=efetuar_devolucao" method="post" name="DEVOLUCAO_ITENS" style="margin:0px" onsubmit="return confirmaDevolucao('<%=NOME_VENDEDOR%>')">
			<div id="containerListagem"></div>
            
            
		</form>
						
			                
						
		<p>&nbsp;</p>			
				
						
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

<script type="text/javascript">
$(document).ready(function() {
	carregaListagem();
});

</script>

<!-- #include file="files/includes/scripts_rodape.asp" -->
<!-- #include file="includes/rodape.asp" -->

  </td>
</tr>
</table>
</body>
</html>