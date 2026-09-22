 <!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!--
'*** Instruções
'1) 1 Grava na temporária os códigos informados e abaixo informa de quem corresponde
'2) Após informar todos clica em avançar para concluir devolução

-->
<%
Dim comando, TXT_SAIDA, ERROR_LINHA, pronto, COD_VENDEDOR, CONT_ITENS
CONT_ITENS = 0
pronto = Request.QueryString("pronto")
ERROR_LINHA = 1
comando = Request.QueryString("comando")
TXT_SAIDA = Request.Form("TXT_SAIDA")
TXT_SAIDA = Replace(TXT_SAIDA," ","")
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


If comando = "devolver" and TXT_SAIDA <> Empty Then
AbrirConexao()
			Dim VALIDA_TIPO, VALIDA_COD_FORMANDO
			VALIDA_TIPO			= Right(TXT_SAIDA,1)
			VALIDA_COD_FORMANDO	= Mid(TXT_SAIDA,1,Len(TXT_SAIDA)-1)
					
					sql = "SELECT TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.TIPO, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_SAIDA, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_RETORNO FROM TB_VENDEDORES INNER JOIN TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ON TB_VENDEDORES.COD_VENDEDOR = TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR WHERE COD_FORMANDO = '"&VALIDA_COD_FORMANDO&"' AND TIPO = '"&VALIDA_TIPO&"' ORDER BY DATA_SAIDA DESC LIMIT 1"
			'Response.Write(sql)
		Set rsVerificaDig = objConexao.Execute(sql)
		
			If not rsVerificaDig.Eof Then
			'Pode gravar, pois o número é válido
				
				objConexao.Execute(	"INSERT INTO TB_FICHAS_IDENTIFICACAO_SAIDA_TEMPORARIA (COD_USUARIO,COD_ETIQUETA)"&_
							"VALUES ('"&Session("COD_USUARIO")&"', "&TXT_SAIDA&")")
				Response.Redirect("epv_saida.asp?pronto=sim&cod_vendedor="&COD_VENDEDOR)
			
			Else
			'Não grava número inválido
				
				Response.Redirect("epv_saida.asp?cod_vendedor="&COD_VENDEDOR&"&pronto=sim&msg=Número de etiqueta: <strong>"&TXT_SAIDA&"</strong> inválida!")	
			
			End if	
		
		rsVerificaDig.Close
		Set rsVerificaDig = Nothing
		
		
FecharConexao()
End If

If comando = "nova" Then
'Apaga tudo da base do determinado usuário
AbrirConexao()
	objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_SAIDA_TEMPORARIA WHERE COD_USUARIO = "&Session("COD_USUARIO"))
FecharConexao()
End If

If comando = "excluir" Then
	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_SAIDA_TEMPORARIA WHERE COD_ETIQUETA = "&Request.QueryString("cod_etiqueta"))
		Response.Redirect("epv_saida.asp?cod_vendedor="&COD_VENDEDOR)
	FecharConexao()
End If

If comando = "efetuar_devolucao" Then
'Faz os UPDATES com a data de hoje,
		AbrirConexao()
set config = getConfiguracoes(true, array("periodo_renovacao_padrao") )

		Dim rsUp, COD_FORMANDO_UP, TIPO_UP, ULTIMO_VENDEDOR_ETIQUETA
		Dim COD_VENDEDOR_STILLOS, DATA_RETORNO_STILLOS, COD_MOVIMENTACAO_STILLOS
		Dim COD_VENDED0R_UP, DIAS_PRAZO_PADRAO
		DIAS_PRAZO_PADRAO = config("periodo_renovacao_padrao") 'Dias é o padrão para devolução quando a saída é do estoque
		
		
		Set rsUp = objConexao.Execute("SELECT COD_ETIQUETA FROM TB_FICHAS_IDENTIFICACAO_SAIDA_TEMPORARIA WHERE COD_USUARIO="&Session("COD_USUARIO")&" GROUP BY COD_ETIQUETA")

			COD_VENDEDOR_UP		=Request.Form("COD_VENDEDOR")
			Dim INSERIR_FICHA
			
			While not rsUP.Eof
			
			INSERIR_FICHA = 1	
			TIPO_UP				= Right(rsUP("COD_ETIQUETA"),1)
			COD_FORMANDO_UP 	= Mid(rsUP("COD_ETIQUETA"),1,Len(rsUP("COD_ETIQUETA"))-1)
			COD_MOVIMENTACAO_UP	= Request.Form("COD_MOVIMENTACAO_"&COD_FORMANDO_UP&"_"&TIPO_UP)
			ULTIMO_VENDEDOR_ETIQUETA= formatNumber(0&Request.Form("ULTIMO_VENDEDOR_"&rsUP("COD_ETIQUETA")),0)
			DATA_PREVISTA_RETORNO = DateAdd("d",DIAS_PRAZO_PADRAO,Date())
		
			'//////////////////
			'Verifica se estiver em nome da Stillo's a FICHA DE IDENTIFICACAO automaticamente faz a transferência para o vendedor
				If ULTIMO_VENDEDOR_ETIQUETA = 27 Then
				'Verifica a ficha do formando como está
					sql =	" SELECT COD_MOVIMENTACAO, COD_VENDEDOR, DATA_RETORNO FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE "&_
							" COD_FORMANDO = "&COD_FORMANDO_UP&" AND TIPO = 1 ORDER BY COD_MOVIMENTACAO DESC LIMIT 1"						

				Set rs5 = objConexao.Execute(sql)
						
						If not rs5.Eof Then
							COD_MOVIMENTACAO_STILLOS=rs5("COD_MOVIMENTACAO")
							COD_VENDEDOR_STILLOS=formatNumber(0&rs5("COD_VENDEDOR"),0)
							DATA_RETORNO_STILLOS=""&rs5("DATA_RETORNO")
						End If
					
					rs5.Close
					Set rs5 = Nothing
					
					'Verifica se realmente o último era da Stillo's para não criar duplicidade e executa a transferência
					If COD_VENDEDOR_STILLOS = 27 Then
						
						'Se tiver em aberto para Stillo's então fecha primeiro na Stillo's e cria novo registro
						If DATA_RETORNO_STILLOS = "" Then						
							
							sql = " UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET DATA_RETORNO = NOW( ), TRANSFERIDO = '1' "&_
								  " WHERE COD_MOVIMENTACAO = "&COD_MOVIMENTACAO_STILLOS&" LIMIT 1"			
							objConexao.Execute(sql)
							
						End If
						
							sql= " INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES (COD_FORMANDO, TIPO, COD_VENDEDOR, DATA_SAIDA, CADASTRADO_POR, TRANSFERIDO) "&_ 
								 " VALUES ("&COD_FORMANDO_UP&","&TIPO_UP&","&COD_VENDEDOR_UP&",NOW(),"&Session("COD_USUARIO")&",2)"	
								 
							 INSERIR_FICHA = 2					
							objConexao.Execute(sql)
						
					End If
				End If
					
	
			
				'//////////////////
				If INSERIR_FICHA <> 2 Then 
					sql =	"INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ("&_
							"COD_FORMANDO ,"&_
							"TIPO ,"&_
							"COD_VENDEDOR ,"&_
							"DATA_SAIDA ,"&_
							"DATA_PREVISTA_RETORNO ,"&_
							"CADASTRADO_POR "&_
							")"&_
							"VALUES ("&_
							COD_FORMANDO_UP&", "&TIPO_UP&", "&COD_VENDEDOR_UP&", NOW(),"&sqlData(DATA_PREVISTA_RETORNO)&", "&Session("COD_USUARIO")&")"	 
				 
					objConexao.Execute(sql)

				End If
        
        objConexao.Execute("UPDATE TB_PRODUCAO_LOTES_FORMANDOS SET COD_VENDEDOR = "&COD_VENDEDOR_UP&", DATA_PRAZO = "&sqlData(DATA_PREVISTA_RETORNO)&" WHERE COD_FORMANDO ="&sqlNumero(COD_FORMANDO_UP)&" LIMIT 1")
			
			rsUp.MoveNext()
			Wend
		
		rsUp.Close()
		Set rsUp = Nothing

	Response.Redirect("epv_saida_imprimir.asp")
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
document.DEVOLUCAO.TXT_SAIDA.focus();
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
		msg = "Tem certeza que deseja gravar a saída destes itens do estoque para o vendedor selecionado ao lado?"
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

	  <h2>EPV - Sa&iacute;da de Produtos </h2>
<div class="subMenu">
	  Sub-Menu: <a href="epv_devolucao.asp?comando=nova">Iniciar devolu&ccedil;&atilde;o de Vendedor</a> | <a href="epv_saida_inicio.asp?comando=nova">Sa&iacute;da de &Aacute;lbuns do Estoque</a> | <a href="dip_producoes_transferencias_entre_vendedores.asp">Transferir Item entre vendedores</a></div>
	  <div class="msg" align="center">
	    <p><%=Request.QueryString("msg")%></p>
	  </div>
	  	  
		  	<form action="epv_saida.asp?comando=devolver&cod_vendedor=<%=COD_VENDEDOR%>" name="DEVOLUCAO" method="post" style="margin:0px">
			  <table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><div align="center">
      <input onkeypress="mascara(this,soNumeros)" maxlength="8" onfocus="focoCx(this)" onblur="blurCx(this)" style="height:50px; width:200px; font-size:34px;" type="text" name="TXT_SAIDA" />
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
		
		'1) Verifica se há pendencias de devolução
		'2) Verifica seleciona todas as fichas
		Dim rs, rs2, rsVendedorVerifica
		
		Set rs = objConexao.Execute("SELECT COD_ETIQUETA FROM TB_FICHAS_IDENTIFICACAO_SAIDA_TEMPORARIA WHERE COD_USUARIO="&Session("COD_USUARIO")&" GROUP BY COD_ETIQUETA")
		
		
		If not rs.Eof Then
			TIPO = Right(rs("COD_ETIQUETA"),1)				
			COD_FORMANDO = Mid(rs("COD_ETIQUETA"),1,Len(rs("COD_ETIQUETA"))-1)
			'Seleciona a última devolução deste cód_formando
			
			sql = "SELECT TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.TIPO, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_SAIDA, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_RETORNO FROM TB_VENDEDORES INNER JOIN TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ON TB_VENDEDORES.COD_VENDEDOR = TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR WHERE COD_FORMANDO = '"&COD_FORMANDO&"' AND TIPO = '"&TIPO&"' ORDER BY DATA_SAIDA DESC LIMIT 1"
			
			
			
			'Response.Write(sql)
		%>	
		<h3 align="center">Sa&iacute;da   de Produtos<br />
		Para o vendedor(a): <%=NOME_VENDEDOR%></h3>
		<form action="epv_saida.asp?comando=efetuar_devolucao&cod_vendedor=<%=COD_VENDEDOR%>" method="post" name="DEVOLUCAO_ITENS" style="margin:0px" onsubmit="return confirmaDevolucao('<%=NOME_VENDEDOR%>')">
			<table width="98%" border="1" align="center" 	cellpadding="2" style="border-collapse:collapse">
			<tr bgcolor="whitesmoke">
       
            <td width="6%"><div align="center"><strong>Retirada n&ordm; </strong></div></td>
            <td width="9%"><div align="center"><strong>C&oacute;d. Formando </strong></div></td>
            <td width="21%"><div align="center"><strong>Nome</strong></div></td>
            <td width="17%"><div align="center"><strong>Curso</strong></div></td>
            <td width="11%"><div align="center">
              <div align="center"><strong>Item Retirado </strong></div>
            </div></td>
            <td width="8%"><div align="center"><strong>Data do &uacute;ltimo retorno </strong></div></td>
            <td width="14%"><div align="center"><strong>&Uacute;ltimo Vendedor </strong></div></td>
            <td width="10%"><div align="center"><strong>Status</strong></div></td>
			<td width="4%"><div align="center"><strong>Excluir da Sa&iacute;da </strong></div></td>
			</tr>
		
						
						<%

						While not rs.Eof
						'*Significa possui diversos formandos
								Dim COD_FORMANDO, TIPO, NOME_VENDEDOR_NEW, COD_VENDEDOR_NEW, DATA_RETORNO, COD_MOVIMENTACAO, viaRetirada, NOME_CURSO, ENCADERNACOES_EXTRAS
								TIPO = Right(rs("COD_ETIQUETA"),1)
								COD_FORMANDO = Mid(rs("COD_ETIQUETA"),1,Len(rs("COD_ETIQUETA"))-1)
								viaRetirada = 0
						Set rs2 = objConexao.Execute("SELECT * FROM VW_DEVOLUCAO_PRODUTOS WHERE COD_FORMANDO="&COD_FORMANDO&" AND TIPO="&TIPO&" ORDER BY TIPO, DATA_SAIDA DESC",viaRetirada)
						
						If not rs2.Eof Then
						  	NOME_VENDEDOR_NEW		= rs2("NOME_VENDEDOR")
							COD_MOVIMENTACAO		= rs2("COD_MOVIMENTACAO")
							COD_VENDEDOR_NEW		= rs2("COD_VENDEDOR")
							ALBUM					= rs2("ALBUM")
							DATA_RETORNO			= ""&rs2("DATA_RETORNO")
							NOME_FORMANDO			= rs2("NOME_FORMANDO")
							DATA_SAIDA				= rs2("DATA_SAIDA")
							NOME_CURSO				= rs2("NOME_CURSO")	
							ENCADERNACOES_EXTRAS	= formatNumber(0&rs2("ENCADERNACOES_EXTRAS"),0)
							CONT_ITENS				= CONT_ITENS + 1
						%>
          					<tr>
            <td><div align="center" class="textoMedio"><strong><%=viaRetirada+1%></strong></div></td>
            <td><div align="center"><%=digito(COD_FORMANDO)%></div></td>
            <td><div align="left"><%=NOME_FORMANDO%></div></td>
            <td><div align="center"><%=NOME_CURSO%></div></td>
            <td><div align="center" class="textoMedio"><strong>
			<%If TIPO = 0 Then%>
			<%If ENCADERNACOES_EXTRAS = 1 Then%>
					&Aacute;lbum Duplo
				<%Else%>
					&Aacute;lbum
				<%End If%>
			<%Elseif TIPO = 1 Then%>
			Ficha
			<%Else
			'Produto agregado, consultar qual é.
			Dim rsProdAgrePend
			sql = 	"SELECT TB_PRODUTOS.NOME_PRODUTO "&_
					"FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = "&_
					"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO WHERE "&_
					"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO = "&COD_FORMANDO&" AND "&_
					"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.TIPO = "&TIPO&" LIMIT 1"
			Set rsProdAgrePend = objConexao.Execute(sql)
				If not rsProdAgrePend.Eof Then
				%>
				<%=rsProdAgrePend("NOME_PRODUTO")%>					
			<%
				End If
			rsProdAgrePend.Close
			Set rsProdAgrePend = Nothing
			End If%>
			</strong></div></td>
            <td><div align="center"><%=DATA_RETORNO%></div></td>
						  <td>
						  
                          <%If int(COD_VENDEDOR_NEW) = 27 Then%> 
                          Atenção caso a FICHA estiver em nome da Stillo's como resp. A mesma passará a ser da responsabilidade do vendedor que está retirando. Transferência automatica de ficha.
                          <%Else%>
						  <%=NOME_VENDEDOR_NEW%>
						  <%End if%>
						  <input type="hidden" name="ULTIMO_VENDEDOR_<%=COD_FORMANDO%><%=TIPO%>" value="<%=COD_VENDEDOR_NEW%>" /></td>
						  <td>
						  <table width="98%" border="0" cellspacing="0" cellpadding="0">
                            
							<%
							Dim sqla
				sqla = 	" SELECT TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_RETORNO, TB_VENDEDORES.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR "&_
						" FROM TB_VENDEDORES "&_
						" INNER JOIN TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ON TB_VENDEDORES.COD_VENDEDOR = TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR "&_
						" WHERE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.TIPO = 1 "&_
						" AND COD_FORMANDO ="&COD_FORMANDO&_
						" ORDER BY TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_SAIDA DESC "&_
						" LIMIT 1"
	
							Set rsVerificaFicha = objConexao.Execute(sqla)
							
								If not rsVerificaFicha.Eof Then
									Dim COD_VENDEDOR_FICHA, DATA_RETORNO_FICHA, NOME_VENDEDOR_FICHA
									COD_VENDEDOR_FICHA = rsVerificaFicha("COD_VENDEDOR")
									DATA_RETORNO_FICHA = rsVerificaFicha("DATA_RETORNO")
									NOME_VENDEDOR_FICHA = rsVerificaFicha("NOME_VENDEDOR")
								End If
							
							rsVerificaFicha.Close
							Set rsVerificaFicha = Nothing
							%>
							
							<%If int(COD_VENDEDOR_FICHA) <> 27 AND isNull(DATA_RETORNO_FICHA) AND int(COD_VENDEDOR_FICHA) <> int(COD_VENDEDOR) Then
							ERROR_LINHA = 2
							%>
							<tr>
								<td><img src="images/icons/icon_red.gif" width="17" height="15" /></td>
							  	<td>Ficha consta como aberta para o vendedor: <%=NOME_VENDEDOR_FICHA%></td>
							</tr>
							<%Else%>
							
							
                            <%If Len(DATA_RETORNO) > 3 Then%>
							<tr>
                              <td colspan="2"><div align="center"><img src="images/icons/icon_green.gif" width="16" height="15" /></div></td>
                              </tr>
	
							<%
							Else
							ERROR_LINHA = 2
							%>
													<tr>
                              <td><img src="images/icons/icon_red.gif" width="17" height="15" />
                                <div align="center"></div></td>
                              <td>Item n&atilde;o consta como devolvido</td>
								  </tr>
							<%
							End If
							
						End If
						%>
                          </table></td>
          	              <td><div align="center"><a href="?comando=excluir&cod_etiqueta=<%=COD_FORMANDO%><%=TIPO%>&cod_vendedor=<%=COD_VENDEDOR%>"><img src="images/icons/icon_excluir.png" width="16" height="16" border="0" /></a>
						  <input name="COD_MOVIMENTACAO_<%=COD_FORMANDO%>_<%=TIPO%>" value="<%=COD_MOVIMENTACAO%>" type="hidden" />
						  </div></td>
          	</tr>
          				<%
						End If
							
						rs2.Close
						Set rs2 = Nothing
							
						rs.MoveNext()
						Wend
						%>
						<tr bgcolor="#CCCCCC">
          					  <td colspan="9" class="textoMedio" style="padding:10px"><div align="center"><strong>Itens selecionados: <span class="textoMaior"><%=CONT_ITENS%></span></strong></div></td>
       					    </tr>
						</table>
						<div align="center">
						  <p>
						  <%If ERROR_LINHA <> 2 Then%>
					 
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
						    
						    <%Else%>
						    <span class="msg">Atenção não será possível avançar, pois h&aacute; itens com status em vermelho.<br />Corrija para continuar.</span>
						      <%End If%>
						    </p>
						</div>
		</form>
						
			                
						
		<p>&nbsp;</p>			
		
		<%End If
		
		rs.Close
		Set rs = Nothing
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
