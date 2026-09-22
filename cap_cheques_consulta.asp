<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Validação especial para verificar se possui sub-validação específica
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '41' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se não encontrar então redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
		
	rs.Close
	Set rs=Nothing
	
FecharConexao()




Dim sql, rs
sql				= ""

Dim BANCO, AGENCIA, CONTA_CORRENTE, NUMERO_CHEQUE, VALOR, BOM_PARA

If Request.QueryString("filtro1") <> Empty Then	

	BANCO			= formatNumber(0&Request.QueryString("BANCO"),0)
	AGENCIA			= ""&Request.QueryString("AGENCIA")
	CONTA_CORRENTE	= ""&Request.QueryString("CONTA_CORRENTE")
	NUMERO_CHEQUE	= ""&Request.QueryString("NUMERO_CHEQUE")
	VALOR			= formatNumber(0&Request.QueryString("VALOR"))
	BOM_PARA		= ""&Request.QueryString("BOM_PARA")


	If int(BANCO) > 0 Then
			sql = sql & " AND BANCO = "&BANCO
	
	Else
	
		BANCO = ""
	
	End If
	
	If AGENCIA <> "" Then
			sql = sql & " AND AGENCIA = "&sqlTexto(AGENCIA)
	End If
	
	If CONTA_CORRENTE <> "" Then
			sql = sql & " AND CONTA_CORRENTE = "&sqlTexto(CONTA_CORRENTE)
	End If
	
	If NUMERO_CHEQUE <> "" Then
			sql = sql & " AND NUMERO_CHEQUE = "&sqlTexto(NUMERO_CHEQUE)
	End If
	
	If int(VALOR) > 0 Then
			sql = sql & " AND VALOR = '"&sqlNumero(VALOR)&"'"
	
	Else
	
		VALOR = ""
	
	End If
	
	If BOM_PARA <> "" Then
			sql = sql & " AND BOM_PARA = "&sqlData(formatDateTime(BOM_PARA))
	End If
	
		
	sql = Mid(sql,6,200)
	
	sql = "SELECT * FROM TB_PEDIDOS_CHEQUE WHERE "&sql&" AND VERIFICADO = 1 ORDER BY COD_PEDIDO, NOME, BOM_PARA"
	
	'Response.Write(sql)
End If


%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Versão 1.2</title>
	<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>
	<script type="text/javascript" src="js/modal-message.js"></script>
<link rel="stylesheet" href="css/modal-message.css" type="text/css">
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
	<script language="javascript">
	function validaForm(){
	d = document.FORM_PADRAO
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
}}}} else {
alert(msgErro);
campo.focus();
return false;
}}


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
//Formata número tipo moeda usando o evento onKeyDown
function Limpar(valor, validos) {
// retira caracteres invalidos da string
var result = "";
var aux;
for (var i=0; i < valor.length; i++) {
aux = validos.indexOf(valor.substring(i, i+1));
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

if (tam < tammax && tecla != 8){ tam = vr.length + 1 ; }

if (tecla == 8 )
{ tam = tam - 1 ; }

if ( tecla == 8 || tecla >= 48 && tecla <= 57 || tecla >= 96 && tecla <= 105 )
{

if ( tam <= dec )
{ campo.value = vr ; }

if ( (tam > dec) && (tam <= 5) ){
campo.value = vr.substr( 0, tam - 2 ) + "," + vr.substr( tam - dec, tam ) ; }
if ( (tam >= 6) && (tam <= 8) ){
campo.value = vr.substr( 0, tam - 5 ) + "." + vr.substr( tam - 5, 3 ) + "," + vr.substr( tam - dec, tam ) ; 
}
if ( (tam >= 9) && (tam <= 11) ){
campo.value = vr.substr( 0, tam - 8 ) + "." + vr.substr( tam - 8, 3 ) + "." + vr.substr( tam - 5, 3 ) + "," + vr.substr( tam - dec, tam ) ; }
if ( (tam >= 12) && (tam <= 14) ){
campo.value = vr.substr( 0, tam - 11 ) + "." + vr.substr( tam - 11, 3 ) + "." + vr.substr( tam - 8, 3 ) + "." + vr.substr( tam - 5, 3 ) + "," + vr.substr( tam - dec, tam ) ; }
if ( (tam >= 15) && (tam <= 17) ){
campo.value = vr.substr( 0, tam - 14 ) + "." + vr.substr( tam - 14, 3 ) + "." + vr.substr( tam - 11, 3 ) + "." + vr.substr( tam - 8, 3 ) + "." + vr.substr( tam - 5, 3 ) + "," + vr.substr( tam - 2, tam ) ;}
} 



}
messageObj = new DHTML_modalMessage();	// We only create one object of this class
messageObj.setShadowOffset(5);	// Large shadow

function displayStaticMessage(messageContent,cssClass,altura)
{
	messageObj.setHtmlContent(messageContent);
	messageObj.setSize(550,altura);
	messageObj.setCssClassMessageBox(cssClass);
	messageObj.setSource(false);	// no html source since we want to use a static message here.
	messageObj.setShadowDivVisible(false);	// Disable shadow for these boxes	
	messageObj.display();
	
	
}

function historicoCheque(COD_PEDIDO,COD_PEDIDO_CHEQUE){
	var caminho
	caminho = "<table cellpadding=\'0\' cellspacing=\'0\' border=\'0\' width=\'100%\'><tr><td><table cellpadding=\'0\' cellspacing=\'0\' width=\'100%\'><tr><td width=\'70%\'><h1>Cheque pré-datado</h1></td><td align=\'right\'><a href=\'#\' onclick=\'closeMessage();return false\'>Fechar Janela</a></td></tr></table></td></tr><tr><td><iframe frameborder=\'0\' scrolling=\'yes\' height=\'380px\' width=\'99%\' src=\'lca_pedidos_pendentes_alteracoes_part_3_POPUP_cheque_historico.asp?comando=visualizar&COD_PEDIDO="+COD_PEDIDO+"&cod_pedido_cheque="+COD_PEDIDO_CHEQUE+"\'></iframe></td></tr></table>"
	displayStaticMessage(caminho,false,420)		
}
function closeMessage()
{
	messageObj.close();	
}
</script>

</head>

<body onload="selecionaBusca()">

<!-- #include file="includes/topo.asp" -->
<table cellpadding="9px" cellspacing="0" width="100%">
<tr>
<td>
<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td><table width="100%" style="border-bottom:#999999 solid dotted 1px" background="images/layout/fundo_menu_azul.jpg"border="0" cellspacing="0" cellpadding="0">
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

	  <h2>CAP - Consulta de Cheques<br />
	    <br />
	  </h2>
	   <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="100%" border="0" align="center" cellpadding="5" cellspacing="0">
		  <tr>
         
            <td width="96%">
            <form name="FORM_PADRAO" action="cap_cheques_consulta.asp" style="margin:0px" onsubmit="return validaForm()">
	            <table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td width="87%" class="textoMedio"><strong>Escolha 1 ou mais parâmetros abaixo para efetuar a busca</strong></td>
                <td width="13%" rowspan="2"><div align="right">
                  <input class="btnFiltrarAmarelo" type="submit" tabindex="7" name="filtro1" value=" " />
                </div></td>
              </tr>
              <tr>
                <td><table width="675" border="0" cellpadding="2" cellspacing="0">
                  <tr>
                    <td width="102">Banco</td>
                    <td width="130"><input maxlength="30" tabindex="1" type="text" value="<%=BANCO%>" class="inputTamanho1" name="BANCO" /></td>
                    <td width="76">Agência</td>
                    <td width="126"><input tabindex="2" type="text" value="<%=AGENCIA%>" class="inputTamanho1" name="AGENCIA" /></td>
                    <td width="107">Contato Corrente</td>
                    <td width="110"><input tabindex="3" type="text" value="<%=CONTA_CORRENTE%>" class="inputTamanho1" name="CONTA_CORRENTE" /></td>
                  </tr>
                  <tr>
                    <td>Número do Cheque</td>
                    <td><input tabindex="4" type="text" value="<%=NUMERO_CHEQUE%>" class="inputTamanho1" name="NUMERO_CHEQUE" /></td>
                    <td>Valor exato</td>
                    <td><input tabindex="5" type="text" onKeydown="Formata(this,20,event,2)" value="<%=VALOR%>" class="inputTamanho1" name="VALOR" /></td>
                    <td>Bom para</td>
                    <td><input class="campoData" type="text" onkeypress="MascaraData(FORM_PADRAO.BOM_PARA);" maxlength="10" tabindex="6" value="<%=BOM_PARA%>" name="BOM_PARA" /></td>
                  </tr>
                </table>
                  </td>
              </tr>
            </table>
            </form>
            </td>

          </tr>
        </table>


		  
		  <%
		  
If Request.QueryString("filtro1") <> Empty Then		  
		  		  
		  AbrirConexao()
		  Set rs = objConexao.Execute(sql,qtdRegistrosBusca)
		  
		  
		  
		  
		  
		  
		  If not rs.Eof Then
		  
		  
		  
		  
		  %>
		   <h3 align="center">Foram  encontrados os seguintes registros para a busca: </h3>
			<table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="7%"><div align="center"><strong>Cód Pedido</strong></div></td>
            <td width="16%"><div align="center"><strong>Formando(s)</strong></div></td>
            <td width="17%"><div align="center"><strong>Nome do Titular</strong></div></td>
            <td width="9%"><div align="center"><strong>Se for terceiro</strong></div></td>
            <td width="9%"><div align="center"><strong>Fone</strong></div></td>
            <td width="5%"><div align="center"><strong>Banco</strong></div></td>
            <td width="6%"><div align="center"><strong>Agência</strong></div></td>
            <td width="10%"><div align="center"><strong>Conta Corrente</strong></div></td>
            <td width="7%"><div align="center"><strong>Número Cheque</strong></div></td>
            <td><div align="center"><strong>Valor do Cheque</strong></div></td>
            <td width="7%"><div align="center"><strong>Historico</strong></div></td>
            <td width="7%"><div align="center"><strong>Bom para</strong></div></td>
            </tr>
		  <%
		  While not rs.Eof
	  	  %>
          <tr>
            <td><div align="center"><a href="lca_busca_detalhamento_pedido.asp?cod_pedido=<%=rs("COD_PEDIDO")%>" title="Clique aqui para detalhar o pedido"><%=rs("COD_PEDIDO")%></a></div></td>
<td>

	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<%
	Set rs2=objConexao.Execute("SELECT NOME_FORMANDO, COD_CONTRATO FROM VW_PEDIDOS_FORMANDOS WHERE COD_PEDIDO = "&rs("COD_PEDIDO"))
	
	While not rs2.Eof
	%>
		  <tr>
    <td width="67%"><div align="left"><%=rs2("NOME_FORMANDO")%></div></td>
    <td width="33%"><div align="center"><%=rs2("COD_CONTRATO")%></div></td>
  </tr>
        
	
	<%
    rs2.MoveNext()
	Wend
	
	rs2.Close
	Set rs2=Nothing
	%>
    </table>
    </td>
<td><%=rs("NOME")%></td>
<td><%=rs("SE_TERCEIRO")%></td>
<td><%=rs("FONE")%></td>
<td><%=rs("BANCO")%></td>
<td><%=rs("AGENCIA")%></td>
<td><%=rs("CONTA_CORRENTE")%></td>
<td><%=rs("NUMERO_CHEQUE")%></td>
            <td width="7%"><%=formatCurrency(rs("VALOR"))%></td>
            
            <td class=""><div align="center"><strong>

            <%
            Set rs3 = objConexao.Execute("SELECT count(*) FROM TB_PEDIDOS_CHEQUE_HISTORICO WHERE COD_PEDIDO_cheque = "&sqlNumero(rs(0)))
            if not rs3.Eof Then
            	qtdHistorico = rs3(0)
            End if
            rs3.Close
            set rs3 = Nothing

            Response.Write("<a href='javascript:historicoCheque("&rs("COD_PEDIDO")&","&rs("COD_PEDIDO_CHEQUE")&")'> Visualizar ("&qtdHistorico&") </a>")

            %>
            </strong></div></td>
<td class="textoMedio"><div align="center"><strong><%=rs("BOM_PARA")%></strong></div></td>
            </tr>
          
		  <%
		  rs.MoveNext()
		  Wend
	      
		    
		  %>
		  <tr bgcolor="whitesmoke">
            <td colspan="12"><div align="center"><strong>Quantidade de Registros Encontrados: <%=qtdRegistrosBusca%></strong></div></td>
            </tr>
        </table>
		  </div>
		  <%Else%>
		  	      <div align="center">
			  <p><em>Não foi encontrado registros para a busca informada!</em>	</p>
		    <p> </p>
		    <p>	 
			 </div>  
			<%End If%>
		 <%
		 rs.Close()
		  Set rs = Nothing
		  FecharConexao()
		  
		  End If
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



