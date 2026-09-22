<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim COD_FORNECEDOR, NOME_RESPONSAVEL_SAIDA, RG_RESPONSAVEL_SAIDA


If Session("COD_FORNECEDOR") = Empty Then
	
	Session("COD_FORNECEDOR") 			= Request.Form("COD_FORNECEDOR")
	Session("NOME_RESPONSAVEL_SAIDA") 	= Request.Form("NOME_RESPONSAVEL_SAIDA")
	Session("RG_RESPONSAVEL_SAIDA") 	= Request.Form("RG_RESPONSAVEL_SAIDA")
	
	AbrirConexao()
		
		Set rs=objConexao.Execute("SELECT * FROM TB_FORNECEDORES WHERE COD_FORNECEDOR = "&Session("COD_FORNECEDOR"))
			
			If not rs.Eof Then
				
				Session("NOME_FANTASIA") = rs("NOME_FANTASIA")	
				
			End If
		
		
		rs.Close
		Set rs=Nothing
		
	FecharConexao()
		
End If


%>


<!--

-->
<%
Dim comando, TXT_SAIDA, ERROR_LINHA, pronto, CONT_ITENS
CONT_ITENS = 0
ERROR_LINHA = 1
pronto = Request.QueryString("pronto")
comando = Request.QueryString("comando")
TXT_SAIDA = Request.Form("TXT_SAIDA")
TXT_SAIDA = Replace(TXT_SAIDA," ","")


If comando = "devolver" and TXT_SAIDA <> Empty Then
AbrirConexao()
			Dim VALIDA_TIPO, VALIDA_COD_FORMANDO, msgError
			VALIDA_TIPO			= Right(TXT_SAIDA,1)
			VALIDA_COD_FORMANDO	= Mid(TXT_SAIDA,1,Len(TXT_SAIDA)-1)
			msgError			= ""
			
			If int(VALIDA_TIPO) < 2 Then
				msgError = "Atenção este produto não poderá sair para o Fornecedor (Álbum ou ficha de identificação)."
			Else
			
				Set rs=objConexao.Execute("SELECT COD_MOVIMENTACAO FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_FORMANDO = "&VALIDA_COD_FORMANDO&" AND TIPO = "&VALIDA_TIPO&" ORDER BY DATA_SAIDA DESC LIMIT 1")
					
					If not rs.Eof Then
						
						objConexao.Execute( " INSERT INTO TB_FORNECEDORES_SAIDA_ITENS_TEMPORARIA(CADASTRADO_POR,COD_ETIQUETA,COD_FORMANDO,COD_MOVIMENTACAO)" &_
											" VALUES ('"&Session("COD_USUARIO")&"', '"&TXT_SAIDA&"', '"&VALIDA_COD_FORMANDO&"', '"&rs("COD_MOVIMENTACAO")&"')") 
											
											msgError = "Item cadastrado!"
					Else
						
						msgError = "Atenção número de etiqueta inválida!"
						
					End If
					
				
				rs.Close()
				Set rs=Nothing
				
			End If	
			
	
			Response.Redirect("epv_saida_fornecedores_itens.asp?pronto=sim&msg="&msgError)	
			

		
		
FecharConexao()
End If

If comando = "nova" Then
'Apaga tudo da base do determinado usu�rio
	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_FORNECEDORES_SAIDA_ITENS_TEMPORARIA WHERE CADASTRADO_POR = "&Session("COD_USUARIO"))
	FecharConexao()
End If

If comando = "excluir" Then
	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_FORNECEDORES_SAIDA_ITENS_TEMPORARIA WHERE COD_ETIQUETA = "&Request.QueryString("cod_etiqueta"))
		Response.Redirect("epv_saida_fornecedores_itens.asp")
	FecharConexao()
End If

If comando = "efetuar_saida" Then
AbrirConexao()
	Dim GRAVA_COD_FORMANDO , GRAVA_COD_MOVIMENTACAO , GRAVA_COD_PRODUTO, GRAVA_TIPO
	
	objConexao.Execute("INSERT INTO TB_FORNECEDORES_SAIDA (COD_FORNECEDOR,NOME_RESPONSAVEL_SAIDA,RG_RESPONSAVEL_SAIDA,CADASTRADO_POR) VALUES ('"&Session("COD_FORNECEDOR")&"', "&sqlTexto(Session("NOME_RESPONSAVEL_SAIDA"))&", "&sqlTexto(Session("RG_RESPONSAVEL_SAIDA"))&", '"&Session("COD_USUARIO")&"')")
	
	
	Dim COD_FORNECEDOR_SAIDA
	Set rsI = objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_FORNECEDORES_SAIDA LIMIT 1")
			
			If not rsI.Eof Then
				
					COD_FORNECEDOR_SAIDA = rsI("ULTIMO_ID")
				
			End If
			
	rsI.Close
	Set rsI = Nothing
	

	
	Set rs = objConexao.Execute("SELECT * FROM TB_FORNECEDORES_SAIDA_ITENS_TEMPORARIA WHERE CADASTRADO_POR="&Session("COD_USUARIO")&" GROUP BY COD_ETIQUETA ORDER BY DATA_CADASTRO")
			
			While not rs.Eof
			
			GRAVA_COD_FORMANDO		= rs("COD_FORMANDO")
			GRAVA_COD_MOVIMENTACAO	= rs("COD_MOVIMENTACAO")
			GRAVA_TIPO				= Right(rs("COD_ETIQUETA"),1)
				
							sql = 		"SELECT TB_PRODUTOS.NOME_PRODUTO, TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO "&_
										"FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = "&_
										"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO WHERE "&_
										"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO = "&GRAVA_COD_FORMANDO&" AND "&_
										"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.TIPO = "&GRAVA_TIPO&" LIMIT 1"
						
						Set rsProdAgrePend = objConexao.Execute(sql)
						If not rsProdAgrePend.Eof Then
						
						GRAVA_COD_PRODUTO = rsProdAgrePend("COD_PRODUTO")
	
						End If
					rsProdAgrePend.Close
					Set rsProdAgrePend = Nothing
					
					'Insere na base
					objConexao.Execute("INSERT INTO TB_FORNECEDORES_SAIDA_ITENS(COD_FORNECEDOR_SAIDA,COD_FORMANDO,COD_PRODUTO,COD_MOVIMENTACAO) VALUES ('"&COD_FORNECEDOR_SAIDA&"', '"&GRAVA_COD_FORMANDO&"', '"&GRAVA_COD_PRODUTO&"', '"&GRAVA_COD_MOVIMENTACAO&"')")
				
			rs.MoveNext()
			Wend
	
	rs.Close
	Set rs = Nothing
	
	
	'Limpa antes de continuar
	Session("COD_FORNECEDOR") 			= Empty
	Session("NOME_RESPONSAVEL_SAIDA") 	= Empty
	Session("RG_RESPONSAVEL_SAIDA") 	= Empty
	Session("NOME_FANTASIA")			= Empty
	
		
	objConexao.Execute("DELETE FROM TB_FORNECEDORES_SAIDA_ITENS_TEMPORARIA WHERE CADASTRADO_POR = "&Session("COD_USUARIO"))
	
FecharConexao()

Response.Redirect("epv_saida_fornecedores_itens_IMPRIMIR.asp?COD_FORNECEDOR_SAIDA="&COD_FORNECEDOR_SAIDA)
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
<!--AJUDA F1 E CTRL PARA BUSCA R�PIDA-->
		
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
 


	<!--AJUDA F1 E CTRL PARA BUSCA R�PIDA-->

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->

function cadastraVendedor(valor){
	alert(valor);
}
</script>
		<!-- CALEND�RIO SCRIPTS
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
function confirmaDevolucao(NOME) {
	var msg = "Tem certeza que deseja gravar a saída destes itens do estoque para o fornecedor: "+NOME+"?";
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

	  <h2>EPV - Sa&iacute;da de Produtos para Fornecedor</h2>
      <div class="msg" align="center">
	    <p><%=Request.QueryString("msg")%></p>
	  </div>
	  	  
		  	<form action="epv_saida_fornecedores_itens.asp?comando=devolver" name="DEVOLUCAO" method="post" style="margin:0px">
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
		
		Dim rs, rs2, rsVendedorVerifica
		
		Set rs = objConexao.Execute("SELECT * FROM TB_FORNECEDORES_SAIDA_ITENS_TEMPORARIA WHERE CADASTRADO_POR="&Session("COD_USUARIO")&" GROUP BY COD_ETIQUETA ORDER BY DATA_CADASTRO",CONT_ITENS)
		
		
		If not rs.Eof Then
		%>	
		<h3 align="center">Sa&iacute;da   de Produtos<br />
		Para o fornecedor: <%=Session("NOME_FANTASIA")%>
        </h3>
        <div align="center" class="textoMedio">Responsável pela retirada: <strong><%=Session("NOME_RESPONSAVEL_SAIDA")%></strong><br>RG: <%=Session("RG_RESPONSAVEL_SAIDA")%><br />
<br /></div>
		<form action="epv_saida_fornecedores_itens.asp?comando=efetuar_saida" method="post" name="DEVOLUCAO_ITENS" style="margin:0px" onsubmit="return confirmaDevolucao('<%=Session("NOME_FANTASIA")%>')">
			<table width="98%" border="1" align="center" 	cellpadding="2" style="border-collapse:collapse">
			<tr bgcolor="whitesmoke">
       
            <td><div align="center"><strong>Produto </strong></div></td>
            <td width="21%"><div align="center"><strong>C&oacute;d. Formando</strong></div></td>
            <td width="31%"><div align="center"><strong>Institui&ccedil;&atilde;o de Ensino</strong></div></td>
            <td width="29%"><div align="center"><strong>Status</strong></div></td>
			<td width="4%"><div align="center"><strong>Excluir da Sa&iacute;da </strong></div></td>
			</tr>
		
						
						<%
						
						While not rs.Eof
						
							'Significa possui diversos formandos
							Dim COD_ETIQUETA, COD_FORMANDO, TIPO, COD_MOVIMENTACAO, COD_CONTRATO, DESCRICAO, COD_PRODUTO, HABILITA_SAIDA_FORNECEDOR, DATA_RETORNO
							COD_ETIQUETA				= rs("COD_ETIQUETA")
							TIPO 						= Right(rs("COD_ETIQUETA"),1)
							COD_FORMANDO 				= rs("COD_FORMANDO")
							COD_MOVIMENTACAO 			= rs("COD_MOVIMENTACAO")
							HABILITA_SAIDA_FORNECEDOR	= 0
							COD_PRODUTO					= 0
							
							
							Set rs2 = objConexao.Execute("SELECT COD_CONTRATO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO ="&COD_FORMANDO)
									
									If not rs2.Eof Then
										
										COD_CONTRATO = rs2("COD_CONTRATO")
										
									End If
							
							rs2.Close
							Set rs2 = Nothing 
							
							
							Set rs2 = objConexao.Execute("SELECT DESCRICAO FROM TB_ENTIDADES WHERE COD_CONTRATO ="&sqlTexto(COD_CONTRATO))
									
									If not rs2.Eof Then
										
										DESCRICAO = rs2("DESCRICAO")
										
									End If
							
							rs2.Close
							Set rs2 = Nothing
							
							%>
								<tr>
				<td><div align="center" class="textoMedio"><strong>
				<%
				'Produto agregado, consultar qual &eacute;.
				Dim rsProdAgrePend
				sql = 	"SELECT TB_PRODUTOS.NOME_PRODUTO, TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO "&_
						"FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = "&_
						"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO WHERE "&_
						"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO = "&COD_FORMANDO&" AND "&_
						"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.TIPO = "&TIPO&" LIMIT 1"
				Set rsProdAgrePend = objConexao.Execute(sql)
					If not rsProdAgrePend.Eof Then
					
					COD_PRODUTO = rsProdAgrePend("COD_PRODUTO")
					%>
				  <%=rsProdAgrePend("NOME_PRODUTO")%>
				  <%
					End If
				rsProdAgrePend.Close
				Set rsProdAgrePend = Nothing
				%>
				</strong></div></td>
				<td><div align="left"><%=digito(COD_FORMANDO)%></div></td>
				<td><div align="left"><%=COD_CONTRATO%> - <%=DESCRICAO%></div></td>
				<td>
							  <table width="98%" border="0" cellspacing="0" cellpadding="3">
								
								<%
								Set rs2=objConexao.Execute("SELECT COD_FORNECEDOR_SAIDA FROM TB_FORNECEDORES_SAIDA_ITENS WHERE COD_MOVIMENTACAO = "&COD_MOVIMENTACAO)
								
								If not rs2.Eof Then
									ERROR_LINHA = 2
								%>
								<tr>
									<td><img src="images/icons/icon_red.gif" width="17" height="15" /></td>
									<td>Item j&aacute; consta com sa&iacute;da para fornecedor. N&uacute;mero sa&iacute;da: <strong><%=rs2("COD_FORNECEDOR_SAIDA")%></strong></td>
								</tr>
								<%
								End If
								
								rs2.Close
								Set rs2=Nothing
								%>
								
								
								<%
								Set rs2=objConexao.Execute("SELECT DATA_RETORNO FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_MOVIMENTACAO ="&COD_MOVIMENTACAO)
									If not rs2.Eof Then
										DATA_RETORNO = ""&rs2("DATA_RETORNO")
									End If
								
									If DATA_RETORNO = "" Then
										ERROR_LINHA = 2
								%>
										<tr>
									<td><img src="images/icons/icon_red.gif" width="17" height="15" /></td>
									<td>Item n&atilde;o pode ser retirado por que n&atilde;o consta como devolvido no estoque de a entrada antes de continuar</td>
								</tr>
								<%	
									End If
									
								rs2.Close
								Set rs2=Nothing
								%>
								
								
								<%
								Set rs2 = objConexao.Execute("SELECT HABILITA_SAIDA_FORNECEDOR FROM TB_PRODUTOS WHERE COD_PRODUTO = "&COD_PRODUTO)
									
									If not rs2.Eof Then
										
										HABILITA_SAIDA_FORNECEDOR = rs2("HABILITA_SAIDA_FORNECEDOR")
										
									End If
							
								rs2.Close
								Set rs2 = Nothing
								
								
								If int(HABILITA_SAIDA_FORNECEDOR) = 2 Then
									ERROR_LINHA = 2	
								%>
								<tr>
								  <td><img src="images/icons/icon_red.gif" width="17" height="15" />
									<div align="center"></div></td>
								  <td>Item n&atilde;o &eacute; considerado pelo sistema como necess&aacute;rio sa&iacute;da para fornecedor. Caso o mesmo estiver correto entre em contato com o Administrador do sistema.</td>
									  </tr>
								<%End If%>
								<%
								If int(ERROR_LINHA = 1) Then
								%>	  
								<tr>
								  <td colspan="2"><div align="center"><img src="images/icons/icon_green.gif" width="16" height="15" /></div></td>
								  </tr>
                                 <%End If%>
							  </table></td>
							  <td><div align="center"><a href="?comando=excluir&cod_etiqueta=<%=COD_ETIQUETA%>"><img src="images/icons/icon_excluir.png" width="16" height="16" border="0" /></a>
							  <input name="COD_MOVIMENTACAO_<%=COD_FORMANDO%>_<%=TIPO%>" value="<%=COD_MOVIMENTACAO%>" type="hidden" />
							  </div></td>
				</tr>
							<%							
						rs.MoveNext()
						Wend
						%>
						<tr bgcolor="#CCCCCC">
          					  <td colspan="5" class="textoMedio" style="padding:10px"><div align="center"><strong>Itens selecionados: <span class="textoMaior"><%=CONT_ITENS%></span></strong></div></td>
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