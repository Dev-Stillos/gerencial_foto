<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!--
'*** Instruções
'1) Grava somente a se tiver sido lançado a ficha do formandos
'2) Após a gravação do cód. de identificação verifica o que o formando possui em aberto
'3) Verifica se o formando já não está na responsábilidade do vendedor que está retirando
-->
<%
Dim comando, COD_BARRAS, ERROR_LINHA, pronto, COD_VENDEDOR, CONT_ITENS
Dim COD_FORMANDO_COD_BARRAS, TIPO_COD_BARRAS, POSSUI_PEDIDO
Dim redirecionar, redirecionar_msg

set config = getConfiguracoes(false, array("periodo_renovacao_padrao"))

redirecionar = "dip_repasse.asp"

'Contador e verificador de erros
'///
CONT_ITENS = 0
ERROR_LINHA = 1
POSSUI_PEDIDO = 0
'///

comando = Request.QueryString("comando")
COD_BARRAS = ""&Request.Form("COD_BARRAS")
COD_BARRAS = Replace(COD_BARRAS," ","")
COD_VENDEDOR = Request.QueryString("COD_VENDEDOR")
Dim NOME_VENDEDOR

AbrirConexao()
	Set rs = objConexao.Execute("SELECT NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR = "&COD_VENDEDOR)
			
			If not rs.Eof Then
				
				NOME_VENDEDOR = rs("NOME_VENDEDOR")
				
			End If
						
	rs.Close
	Set rs = Nothing
FecharConexao()




If comando = "nova" Then
'Apaga tudo da base do determinado usuário
AbrirConexao()
	objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_DEVOLUCAO_TEMPORARIA WHERE COD_USUARIO = "&Session("COD_USUARIO"))
FecharConexao()
End If

If comando = "transferencia" and COD_BARRAS <> "" Then
	
	'1 - Verifica se é ficha de identificação
	'2 - Verifica se o cód. do formando existe
	
	COD_FORMANDO_COD_BARRAS	= COD_BARRAS
	COD_FORMANDO_COD_BARRAS	= Mid(COD_FORMANDO_COD_BARRAS,1,Len(COD_FORMANDO_COD_BARRAS)-1)
	TIPO_COD_BARRAS			= Right(COD_BARRAS,1)
	
	
  If TIPO_COD_BARRAS = 1 Then
	'É ficha de identificação pode continuar
		'2//
		AbrirConexao()
      
			
			sql = "SELECT COD_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO = "&COD_FORMANDO_COD_BARRAS
			Set rs = objConexao.Execute(sql)
				
					If not rs.Eof Then
					'Existe o formando pode inserir na tabela temporária
						PRAZO = formatNumber(0&Request.Form("PRAZO_"&COD_FORMANDO_COD_BARRAS),0)
						PRAZO_NECESSARIO = Request.Form("PRAZO_NECESSARIO_"&COD_FORMANDO_COD_BARRAS)
						
						If int(PRAZO) = 0 Then
							PRAZO = config("periodo_renovacao_padrao")
						End If
					
						sql =	" INSERT INTO TB_FICHAS_IDENTIFICACAO_DEVOLUCAO_TEMPORARIA (COD_USUARIO,COD_ETIQUETA,PRAZO,PRAZO_NECESSARIO, COD_FORMANDO) "&_
								" VALUES ("&Session("COD_USUARIO")&","&COD_BARRAS&",'"&PRAZO&"','"&PRAZO_NECESSARIO&"', '"&COD_FORMANDO_COD_BARRAS&"')"
								
						objConexao.Execute(sql)
						redirecionar_msg = "Código inserido com sucesso!<br />"
						Response.Redirect(redirecionar&"?cod_vendedor="&COD_VENDEDOR&"&comando=continua&msg="&redirecionar_msg)	
						
						
					Else
					'Não existe retorna o erro
						redirecionar_msg = "Código de formando não encontrado!<br />"
						redirecionar_msg = redirecionar_msg & "Cód inserido: <strong>"&COD_BARRAS&"</strong>."
						Response.Redirect(redirecionar&"?cod_vendedor="&COD_VENDEDOR&"&comando=continua&msg="&redirecionar_msg)				
					End If
				
			rs.Close
			Set rs = Nothing
		
		
		FecharConexao()
		
		
	Else
		'1//
		redirecionar_msg = "Atenção você não inseriu uma ficha de identificação, somente fichas são válidas para transferência de repasses!<br />"
		redirecionar_msg = redirecionar_msg & "Cód inserido: <strong>"&COD_BARRAS&"</strong>."
		Response.Redirect(redirecionar&"?cod_vendedor="&COD_VENDEDOR&"&comando=continua&msg="&redirecionar_msg)
	End If
	
	
End If


If comando = "excluir" Then
	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_DEVOLUCAO_TEMPORARIA WHERE COD_ETIQUETA = "&Request.QueryString("cod_etiqueta"))
		Response.Redirect("dip_repasse.asp?cod_vendedor="&COD_VENDEDOR&"&msg=Exclusão efetuada com sucesso!")
	FecharConexao()
End If
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
document.DEVOLUCAO.COD_BARRAS.focus();
}
</script>
<script>
function confirmaDevolucao() {
	var msg
	
	d = document.DEVOLUCAO_ITENS
	if (d.COD_VENDEDOR.value == ""){
		d.COD_VENDEDOR.focus();
		alert('Escolha o vendedor que está retirando os produtos!');
		return false;
	}else{
		msg = "Tem certeza que deseja gravar a transferência destes itens para o vendedor '<%=NOME_VENDEDOR%>'?"
		if(YNconfirm(msg)){
				return true;
				//document.DEVOLUCAO.submit();			
		}else{
				return false;
		}
	
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

	  <h2>DIP - Transfer&ecirc;ncia de Itens</h2>
<div class="msg" align="center"><p><%=Request.QueryString("msg")%></p></div>
	  	  
		  	<form action="dip_repasse.asp?comando=transferencia&cod_vendedor=<%=COD_VENDEDOR%>" name="DEVOLUCAO" method="post" style="margin:0px">
			  <table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><div align="center">
      <input onkeypress="mascara(this,soNumeros)" maxlength="8" onfocus="focoCx(this)" onblur="blurCx(this)" style="height:50px; width:200px; font-size:34px;" type="text" name="COD_BARRAS" />
    </div></td>
    </tr>
  <tr>
    <td><div align="center">
      <input class="btnGravar" type="submit" name="Submit2" value=" " />
    </div></td>
    </tr>
</table>
			</form>

<%

AbrirConexao()

sql = "SELECT COD_ETIQUETA FROM TB_FICHAS_IDENTIFICACAO_DEVOLUCAO_TEMPORARIA WHERE COD_USUARIO="&Session("COD_USUARIO")&" GROUP BY COD_ETIQUETA"
Set rs=objConexao.Execute(sql)

If not rs.Eof Then
%>
		<h3 align="center"> Transfer&ecirc;ncia de Itens<br />
		Para o vendedor(a): <%=NOME_VENDEDOR%></h3>
		<form action="dip_repasse_processa.asp?cod_vendedor=<%=COD_VENDEDOR%>&comando=processar" method="post" name="DEVOLUCAO_ITENS" style="margin:0px" onsubmit="return confirmaDevolucao('<%=NOME_VENDEDOR%>')">
			<table width="98%" border="1" align="center" 	cellpadding="2" style="border-collapse:collapse">
			<tr bgcolor="whitesmoke">
       
            <td width="14%"><div align="center"><strong>C&oacute;d. Formando </strong></div></td>
            <td width="18%"><div align="center"><strong>Nome</strong></div></td>
            <td width="7%"><div align="center"><strong>Curso</strong></div></td>
            <td width="7%"><div align="center"><strong>Repasse*</strong></div></td>
            <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td colspan="3"><div align="center">
                    <div align="center"><strong>Transfer&ecirc;ncia</strong></div>
                </div></td>
              </tr>
              <tr>
                <td width="27%"><div align="center"><strong>Item</strong></div></td>
                <td width="39%"><div align="center"><strong>&Uacute;ltmo Resp.</strong></div></td>
                <td width="34%"><div align="center"><strong>Item em aberto</strong></div></td>
              </tr>
            </table></td>
            <td width="5%"><div align="center"><strong>Status</strong></div></td>
			<td width="5%"><strong>Prazo</strong><a href="javascript:alert('Se este formando estiver cadastrado com número de lote, então é necessário informar o prazo para prestação de contas.')"><img src="images/icons/icon_help_mini.jpg" width="15" height="11" border="0" /></a></td>
			<td width="8%"><div align="center"><strong>Excluir da Transfer&ecirc;ncia </strong></div></td>
			</tr>
			
			<%
			Dim COD_FORMANDO, NOME_FORMANDO, NOME_CURSO, SEMESTRE_CONCLUSAO, ANO_CONCLUSAO, ENCADERNACOES_EXTRAS
			Dim TRANS_DATA_CONCLUSAO, TRANS_DATA_VIRA_REPASSE, TRANS_DATA_ATUAL, DIFERENCA_DATA
			Dim TIPO, ITEM_TRANSFERENCIA, DATA_RETORNO
			Dim COD_VENDEDOR_FICHA, DATA_RETORNO_FICHA
			Dim PRAZO_NECESSARIO
			
			While not rs.Eof
				
				COD_FORMANDO = Mid(rs("COD_ETIQUETA"),1,Len(rs("COD_ETIQUETA"))-1)
				
				'1 - Verifica qual é o nome formando e nome do curso
				'///
				Set rs2 = objConexao.Execute("SELECT NOME_FORMANDO, NOME_CURSO, SEMESTRE_CONCLUSAO, ANO_CONCLUSAO, ENCADERNACOES_EXTRAS FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO = "&COD_FORMANDO)	
					
						If not rs2.Eof Then
							ENCADERNACOES_EXTRAS= formatNumber(0&rs2("ENCADERNACOES_EXTRAS"),0)
							NOME_FORMANDO		= rs2("NOME_FORMANDO")
							NOME_CURSO			= rs2("NOME_CURSO")
							SEMESTRE_CONCLUSAO	= rs2("SEMESTRE_CONCLUSAO")
							ANO_CONCLUSAO		= rs2("ANO_CONCLUSAO")
							
							'Verifica repasse
							'////////////
							'***Baseado na data atual
							
							'Ex. 1 Repasse: Conclusão faculdade DEZEMBRO DE 2005 virará repasse em DEZEMBRO DE 2008.
							'Ex. 2 Repasse: Conclusão faculdade JULHO DE 2006 virará repasse em JULHO DE 2009.
					
							If SEMESTRE_CONCLUSAO = "J" Then
								SEMESTRE_CONCLUSAO = 7
							Elseif SEMESTRE_CONCLUSAO = "D" Then
								SEMESTRE_CONCLUSAO = 12
							End If
					
							TRANS_DATA_CONCLUSAO	= formatDateTime(ANO_CONCLUSAO&"-"&SEMESTRE_CONCLUSAO&"-1")
							TRANS_DATA_VIRA_REPASSE	= DateAdd("yyyy",3,TRANS_DATA_CONCLUSAO)
							TRANS_DATA_ATUAL 		= formatDateTime(Year(Now)&"-"&Month(Now)&"-1")
									
							'Para ser repasse tem de ser no mínimo 3 anos de diferença
							'Verifica se já ha tempo mínimo para virar repasse
							DIFERENCA_DATA = DateDiff("d",TRANS_DATA_VIRA_REPASSE,TRANS_DATA_ATUAL)
							'////////////
							
							
						End If
					
				rs2.Close
				Set rs2 = Nothing
				'///
			CONT_ITENS = CONT_ITENS + 1
			%>
				<tr>
                              <td><div align="center"><%=digito(COD_FORMANDO)%></div></td>
                              <td><%=NOME_FORMANDO%></td>
            <td><%=NOME_CURSO%></td>
            <td>
            <div align="center">
            <%If DIFERENCA_DATA => 0 Then%>
            	SIM
            <%Else%>
            	NÃO
            <%End If%>
            </div></td>
            <td>
            	<table width="100%" border="1" cellpadding="2" cellspacing="0" bgcolor="#F3F3F3" style="border-collapse:collapse">
			<%
			'Transferência 
			'/////////////////
				'1 Verifica quais o produtos que o formando possui
				'2 Verifica se há algum produto em aberto, se tiver será transferido para o novo vendedor
				
				'Zerando variaveis
				ITEM_TRANSFERENCIA = ""
				COD_VENDEDOR_FICHA = 0
				DATA_RETORNO_FICHA = ""
				PRAZO_NECESSARIO = ""
				
				
				sql = "SELECT TIPO FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE  COD_FORMANDO = "&COD_FORMANDO&" GROUP BY TIPO ORDER BY TIPO"
				Set rs2 = objConexao.Execute(sql)
						
						While not rs2.Eof
							
							'Verifica dos tipo quais estão em aberto	
							'////
								sql =	" SELECT TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_SAIDA, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_RETORNO, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR  FROM  "&_
										" TB_VENDEDORES INNER JOIN TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ON (TB_VENDEDORES.COD_VENDEDOR = TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR) "&_
										"  WHERE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.TIPO = "&rs2("TIPO")&" AND TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_FORMANDO = "&COD_FORMANDO&_
										"  ORDER BY COD_MOVIMENTACAO DESC LIMIT 1"
								Set rs3=objConexao.Execute(sql)
								
									If not rs3.Eof Then
									
									
									'Exibe o item e último resposável e verifica qual é o tipo
										DATA_RETORNO = ""&rs3("DATA_RETORNO")
										
										TIPO = formatNumber(0&rs2("TIPO"),0)
										
										If TIPO = 0 and ENCADERNACOES_EXTRAS = 0 Then
										'Álbum
											ITEM_TRANSFERENCIA = "ÁLBUM"
										Elseif TIPO = 0 and ENCADERNACOES_EXTRAS = 1 Then
										'Álbum Duplo
											ITEM_TRANSFERENCIA = "ÁLBUM DUPLO"
										Elseif TIPO = 1 Then
										'Ficha de Identificação
											ITEM_TRANSFERENCIA = "FICHA"
											COD_VENDEDOR_FICHA = formatNumber(0&rs3("COD_VENDEDOR"),0)
											DATA_RETORNO_FICHA = rs3("DATA_RETORNO")
										Else
										'Produto agregado
											sql = 	"SELECT TB_PRODUTOS.NOME_PRODUTO "&_
													"FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = "&_
													"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO WHERE "&_
													"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO = "&COD_FORMANDO&" AND "&_
													"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.TIPO = "&TIPO&" LIMIT 1"
				
											Set rsProdAgrePend = objConexao.Execute(sql)
											If not rsProdAgrePend.Eof Then
											
												ITEM_TRANSFERENCIA = rsProdAgrePend("NOME_PRODUTO")					
											
											End If
											rsProdAgrePend.Close
											Set rsProdAgrePend = Nothing
										
											
										End If
									%>
										
  <tr>
    <td width="27%"><%=ITEM_TRANSFERENCIA%></td>
    <td width="39%"><%=rs3("NOME_VENDEDOR")%></td>
    <td width="34%"><div align="center">
		<%If DATA_RETORNO <> "" Then%>
    		N&Atilde;O
        <%Elseif DATA_RETORNO = "" and TIPO <> 1 Then%>
           	<span class="msg">SIM - Respons&aacute;vel ser&aacute; o vendedor que est&aacute; rebendo a transfer&ecirc;ncia</span>
        <%Elseif DATA_RETORNO = "" and TIPO = 1 Then %>
        	SIM
		<%End If%>
    </div></td>
  </tr>

                                    <%
									End If							
								Set rs3=Nothing
							'////	
							
							rs2.MoveNext()
							Wend				
										
				rs2.Close
				Set rs2 = Nothing
					
			'/////////////////		
			%></table>           	</td>
            <td>
            <div align="center">
            <table width="98%" border="0" cellspacing="0" cellpadding="0">
  							<%
							'Se o cód dos vendedores forem iguais e a ficha estiver em aberto não permite continuar pois
							'o mesmo haverá duplicidade de itens em aberto												
							If int(COD_VENDEDOR_FICHA) = int(COD_VENDEDOR) AND isNull(DATA_RETORNO_FICHA) Then
							ERROR_LINHA = 2
							%>
							<tr>
								<td><img src="images/icons/icon_red.gif" width="17" height="15" /></td>
							  	<td>Ficha consta em aberto para o mesmo vendedor da transferência.</td>
							</tr>
                            <%
								
							Else 
								
								'Verifica se possui pedido para o formando, se tiver não faz a transferência
								'///
								sql = "SELECT COD_PEDIDO FROM VW_PEDIDOS_COMPLETO WHERE COD_FORMANDO = "&COD_FORMANDO&" AND FINALIZA_CORRECAO = 1"
								Set rsPedido = objConexao.Execute(sql)					
									If not rsPedido.Eof Then
										POSSUI_PEDIDO = formatNumber(0&rsPedido("COD_PEDIDO"),0)
									Else
										POSSUI_PEDIDO = 0									
									End If
							
								rsPedido.Close
								Set rsPedido = Nothing
								'///
								
								
								'///
								If INT(POSSUI_PEDIDO) > 0 Then
								ERROR_LINHA = 2
								%>
                            	<tr>
									<td><img src="images/icons/icon_red.gif" width="17" height="15" /></td>
							  		<td>Possui pedido n&uacute;mero <%=POSSUI_PEDIDO%>.</td>
								</tr>
								<%Else%>
									<tr>
                            		  <td colspan="2"><div align="center"><img src="images/icons/icon_green.gif" width="16" height="15" /></div></td>
                            		  </tr>
								<%End If%>
                          <%End If%>
                          </table>
            </div></td>
          	              <td>
          	                <div align="center">
                            <%							
							Set rs2=objConexao.Execute("SELECT COD_FORMANDO FROM TB_PRODUCAO_LOTES_FORMANDOS WHERE cod_formando = "&COD_FORMANDO)
							If not rs2.Eof Then
								PRAZO_NECESSARIO = "S"
							%>
          	                  <input class="inputTamanho1" value="<%=config("periodo_renovacao_padrao")%>" style="width:40px;" type="text" name="PRAZO_<%=COD_FORMANDO%>" />
                            <%
							Else
								PRAZO_NECESSARIO = "N"
							%>
							Prazo padrão
                            <%
                            End If
							
							rs2.Close
							Set rs2=Nothing
							%>
                            <input type="hidden" name="PRAZO_NECESSARIO_<%=COD_FORMANDO%>" value="<%=PRAZO_NECESSARIO%>" />
        	                  </div></td>
          	              <td><div align="center"><a href="?comando=excluir&cod_etiqueta=<%=rs("COD_ETIQUETA")%>&cod_vendedor=<%=COD_VENDEDOR%>"><img src="images/icons/icon_excluir.png" width="16" height="16" border="0" /></a>
						  </div></td>
          	</tr>
			<%
			rs.MoveNext()
			Wend
			%>
			
            
            <tr bgcolor="#CCCCCC">
          					  <td colspan="8" class="textoMedio" style="padding:10px"><div align="center"><strong>Itens selecionados: <span class="textoMaior"><%=CONT_ITENS%></span></strong> - <strong>M&aacute;x 40 por vez</strong></div></td>
       					    </tr>
						</table>
			<br />
						<strong>Observa&ccedil;&otilde;es:</strong><br />
						*Repasse calculado com base na data atual<br />
						<div align="center">
						  <p>
						  <%If ERROR_LINHA <> 2 and int(CONT_ITENS)  < 41 Then%>
						  </p>		
						  <table width="98%" border="0" cellspacing="0" cellpadding="5">
                            <tr>
                              <td><div align="center">
                                  <input class="btnAvancar" name="btnAvancar" type="submit" value="" />
                                  <input name="COD_VENDEDOR" type="hidden" value="<%=COD_VENDEDOR%>" />
                              </div></td>
                            </tr>
                          </table>
						  <p>
               			<%Elseif ERROR_LINHA = 2 Then%>
						    <span class="msg">Atenção não será possível avançar, pois h&aacute; itens com status em vermelho.<br />Corrija para continuar.</span>						
						    <%Elseif int(CONT_ITENS)  > 40 Then%>
						    <span class="msg">Atenção não será possível avançar, pois h&aacute; há mais de 60 fichas lançadas exclua até atingir o limite.<br />Corrija para continuar.</span><%End If%>
						    </p>
						</div>
		</form>
						
			             
<%
End If

rs.Close
Set rs=Nothing
FecharConexao()
%>
<p align="center">&nbsp;</p>		</td>
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
