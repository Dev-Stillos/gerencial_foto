<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim COD_SALA_ARMAZENADO, COD_CAIXA_ARMAZENADO

COD_SALA_ARMAZENADO= int(0&Request.QueryString("COD_SALA_ARMAZENADO"))
COD_CAIXA_ARMAZENADO= int(0&Request.QueryString("COD_CAIXA_ARMAZENADO"))


if COD_SALA_ARMAZENADO = 0 or COD_CAIXA_ARMAZENADO = 0 Then
	Response.Redirect("epv_estocagem_produtos.asp?msg=Desculpe houve uma falha, informe ao depto. de informática.")
End If


Dim comando, TXT_DEVOLUCAO, ERROR_LINHA, pronto, CONT_ITENS
pronto = Request.QueryString("pronto")
ERROR_LINHA = 1
comando = Request.QueryString("comando")
TXT_DEVOLUCAO = Request.Form("TXT_DEVOLUCAO")
TXT_DEVOLUCAO = Replace(TXT_DEVOLUCAO," ","")
CONT_ITENS = 0

If comando = "devolver" and TXT_DEVOLUCAO <> Empty Then
AbrirConexao()
			Dim VALIDA_TIPO, VALIDA_COD_FORMANDO
			VALIDA_TIPO			= Right(TXT_DEVOLUCAO,1)
			VALIDA_COD_FORMANDO	= Mid(TXT_DEVOLUCAO,1,Len(TXT_DEVOLUCAO)-1)
					
					sql = "SELECT TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.TIPO, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_SAIDA, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_RETORNO FROM TB_VENDEDORES INNER JOIN TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ON TB_VENDEDORES.COD_VENDEDOR = TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR WHERE COD_FORMANDO = '"&VALIDA_COD_FORMANDO&"' AND TIPO = '"&VALIDA_TIPO&"' ORDER BY DATA_SAIDA DESC LIMIT 1"
			'Response.Write(sql)
		Set rsVerificaDig = objConexao.Execute(sql)
		
			If not rsVerificaDig.Eof Then

			'Pode gravar, pois o número é válido
				
				objConexao.Execute(	"INSERT INTO TB_FICHAS_IDENTIFICACAO_DEVOLUCAO_TEMPORARIA (COD_USUARIO,COD_ETIQUETA, COD_FORMANDO)"&_
							"VALUES ('"&Session("COD_USUARIO")&"', "&TXT_DEVOLUCAO&",'"&VALIDA_COD_FORMANDO&"')")
				Response.Redirect("epv_estocagem_produtos_itens.asp?pronto=sim&COD_SALA_ARMAZENADO="&COD_SALA_ARMAZENADO&"&COD_CAIXA_ARMAZENADO="&COD_CAIXA_ARMAZENADO)
			
			Else
			'Não grava número inválido
				
				Response.Redirect("epv_estocagem_produtos_itens.asp?pronto=sim&msg=Número de etiqueta: <strong>"&TXT_DEVOLUCAO&"</strong> inválida!&COD_SALA_ARMAZENADO="&COD_SALA_ARMAZENADO&"&COD_CAIXA_ARMAZENADO="&COD_CAIXA_ARMAZENADO)	
			
			End if	
		
		rsVerificaDig.Close
		Set rsVerificaDig = Nothing
		
		
FecharConexao()
End If

If comando = "nova" Then
'Apaga tudo da base do determinado usuário
AbrirConexao()
	objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_DEVOLUCAO_TEMPORARIA WHERE COD_USUARIO = "&Session("COD_USUARIO"))
FecharConexao()
End If

If comando = "excluir" Then
	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_DEVOLUCAO_TEMPORARIA WHERE COD_ETIQUETA = "&Request.QueryString("cod_etiqueta"))
		Response.Redirect("epv_estocagem_produtos_itens.asp?COD_SALA_ARMAZENADO="&COD_SALA_ARMAZENADO&"&COD_CAIXA_ARMAZENADO="&COD_CAIXA_ARMAZENADO)
	FecharConexao()
End If

If comando = "efetuar_estocagem" Then

		AbrirConexao()
		Dim rsUp, COD_FORMANDO_UP, TIPO_UP
		Dim COD_MOVIMENTACAO_UP, TRANSFERIDO_UP, COD_CAIXA_ARMAZENADO_UP
		
		Set rsUp = objConexao.Execute("SELECT COD_ETIQUETA FROM TB_FICHAS_IDENTIFICACAO_DEVOLUCAO_TEMPORARIA WHERE COD_USUARIO="&Session("COD_USUARIO")&" GROUP BY COD_ETIQUETA")

			
		
			While not rsUP.Eof
			COD_MOVIMENTACAO_UP 	= 0
			TIPO_UP					= Right(rsUP("COD_ETIQUETA"),1)
			COD_FORMANDO_UP 		= Mid(rsUP("COD_ETIQUETA"),1,Len(rsUP("COD_ETIQUETA"))-1)
			COD_MOVIMENTACAO_UP		= Request.Form("COD_MOVIMENTACAO_"&COD_FORMANDO_UP&"_"&TIPO_UP)
							
							'Atualiza informando a caixa atual
							'Obs.: Caso for uma movimentação interna, não fará um novo registro para não poluir a ficha do produto
							      'apenas fará o Update para a nova caixa
								  
							sql = "UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET COD_CAIXA_ARMAZENADO = "&COD_CAIXA_ARMAZENADO&" WHERE "&_
								  "COD_MOVIMENTACAO ="&COD_MOVIMENTACAO_UP&" LIMIT 1"
					
					
				 objConexao.Execute(sql)
			
			rsUp.MoveNext()
			Wend
		
		rsUp.Close()
		Set rsUp = Nothing
		
		Set rs=objConexao.Execute("SELECT numero_caixa FROM TB_LOCAL_ARMAZENAGEM_SALA_CAIXA WHERE COD_CAIXA_ARMAZENADO = "&COD_CAIXA_ARMAZENADO)
			
			If not rs.EOf Then
				numero_caixa = rs("numero_caixa")
			End If
			
		rs.Close
		Set rs=Nothing

	FecharConexao()

	Response.Redirect("epv_espaco_de_armazenamento_salas.asp?msg=Estocagem efetuada com sucesso para caixa número: <strong>"&numero_caixa&"</strong>!")
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
document.DEVOLUCAO.TXT_DEVOLUCAO.focus();
}
</script>
<script>
function confirmaEstocagem(nomeVendedor) {
	var msg
	msg = "Tem certeza que deseja gravar está estocagem?"
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

	  <h2>EPV - Estoque / Estocagem de Produtos</h2>
	  <div class="msg" align="center">
	    <p><%=Request.QueryString("msg")%></p>
	  </div>
	  	  
		  	<form action="epv_estocagem_produtos_itens.asp?comando=devolver&COD_CAIXA_ARMAZENADO=<%=COD_CAIXA_ARMAZENADO%>&COD_SALA_ARMAZENADO=<%Response.Write(COD_SALA_ARMAZENADO)%>" name="DEVOLUCAO" method="post" style="margin:0px">
			  <table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><div align="center">
      <input onkeypress="mascara(this,soNumeros)" maxlength="8" onfocus="focoCx(this)" onblur="blurCx(this)" style="height:50px; width:200px; font-size:34px;" type="text" name="TXT_DEVOLUCAO" />
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
		
		Set rs = objConexao.Execute("SELECT COD_ETIQUETA FROM TB_FICHAS_IDENTIFICACAO_DEVOLUCAO_TEMPORARIA WHERE COD_USUARIO="&Session("COD_USUARIO")&" GROUP BY COD_ETIQUETA")
		
		
		If not rs.Eof Then
			TIPO = Right(rs("COD_ETIQUETA"),1)				
			COD_FORMANDO = Mid(rs("COD_ETIQUETA"),1,Len(rs("COD_ETIQUETA"))-1)
			'Seleciona a última devolução deste cód_formando
			
			sql = "SELECT TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.TIPO, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_SAIDA, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_RETORNO FROM TB_VENDEDORES INNER JOIN TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ON TB_VENDEDORES.COD_VENDEDOR = TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR WHERE COD_FORMANDO = '"&COD_FORMANDO&"' AND TIPO = '"&TIPO&"' ORDER BY DATA_SAIDA DESC LIMIT 1"
			'Response.Write(sql)
		Set rsVendedorVerifica = objConexao.Execute(sql)
		
			If not rsVendedorVerifica.Eof Then
				Dim NOME_VENDEDOR, COD_VENDEDOR
				NOME_VENDEDOR 	= rsVendedorVerifica("NOME_VENDEDOR")
				COD_VENDEDOR 	= rsVendedorVerifica("COD_VENDEDOR")
			End if	
		
		rsVendedorVerifica.Close
		Set rsVendedorVerifica = Nothing
		COD_FORMANDO = ""
		%>	
		<h3 align="center">Rela&ccedil;&atilde;o de Itens que ser&atilde;o armazenados:<br />
		</h3>
		<form action="epv_estocagem_produtos_itens.asp?comando=efetuar_estocagem&COD_CAIXA_ARMAZENADO=<%=COD_CAIXA_ARMAZENADO%>&COD_SALA_ARMAZENADO=<%=COD_SALA_ARMAZENADO%>" method="post" name="DEVOLUCAO_ITENS" style="margin:0px" onsubmit="return confirmaEstocagem()">
			<table width="95%" border="1" align="center" 	cellpadding="2" style="border-collapse:collapse">
			<tr bgcolor="whitesmoke">
       
            <td width="15%"><div align="center"><strong>C&oacute;d. Formando </strong></div></td>
            <td width="39%"><div align="center"><strong>Nome</strong></div></td>
            <td width="11%"><div align="center"><strong>N&uacute;mero do &Aacute;lbum</strong></div></td>
            <td width="15%"><div align="center"><strong>Item</strong></div></td>
            <td width="14%"><div align="center"><strong>Status</strong></div></td>
			<td width="6%"><div align="center"><strong>Excluir da Rela&ccedil;&atilde;o </strong></div></td>
			</tr>
		
						
						<%
						Dim ITEM_VENDIDO
						While not rs.Eof
						'*Significa possui diversos formandos
								Dim COD_FORMANDO, TIPO, NOME_VENDEDOR_NEW, COD_VENDEDOR_NEW, DATA_RETORNO, COD_MOVIMENTACAO, ENCADERNACOES_EXTRAS
								TIPO = Right(rs("COD_ETIQUETA"),1)
								COD_FORMANDO = Mid(rs("COD_ETIQUETA"),1,Len(rs("COD_ETIQUETA"))-1)
						
						Set rs2 = objConexao.Execute("SELECT * FROM VW_DEVOLUCAO_PRODUTOS WHERE COD_FORMANDO="&COD_FORMANDO&" AND TIPO="&TIPO&" ORDER BY TIPO, DATA_SAIDA DESC LIMIT 1")
						DATA_RETORNO = ""
						If not rs2.Eof Then
							ITEM_VENDIDO			= 2
						
							COD_MOVIMENTACAO		= rs2("COD_MOVIMENTACAO")
							COD_VENDEDOR_NEW		= rs2("COD_VENDEDOR")
							ALBUM					= rs2("ALBUM")
							NOME_FORMANDO			= rs2("NOME_FORMANDO")
							TRANSFERIDO				= formatNumber(0&rs2("TRANSFERIDO"),0)
							COD_CAIXA_ARMAZENADO	= formatNumber(0&rs2("COD_CAIXA_ARMAZENADO"),0)
							DATA_RETORNO			= ""&rs2("DATA_RETORNO")
							CONT_ITENS				= CONT_ITENS + 1
							ENCADERNACOES_EXTRAS	= rs2("ENCADERNACOES_EXTRAS")
							
							'Verifica se está vendido
							If int(TIPO) = 0  Then
							
									Set rsVen = objConexao.Execute("SELECT VENDIDO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO ="&COD_FORMANDO)
											
											If not rsVen.Eof Then
												
												if int(rsVen("VENDIDO")) = 1 Then
													ITEM_VENDIDO = 1
												End If
												
											End If
											
									rsVen.Close
									Set rsVen = Nothing
							
							Elseif int(TIPO) > 1 Then
									
							sql = 	"SELECT TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.VENDIDO "&_
									"FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = "&_
									"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO WHERE "&_
									"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO = "&COD_FORMANDO&" AND "&_
									"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.TIPO = "&TIPO&" LIMIT 1"
							Set rsProdAgrePend = objConexao.Execute(sql)
							If not rsProdAgrePend.Eof Then
							VENDIDO = formatNumber(0&rsProdAgrePend("VENDIDO"),0)
									If VENDIDO = 1 Then
										ITEM_VENDIDO = 1
									End If
							End If
							
							rsProdAgrePend.Close
							Set rsProdAgrePend = Nothing
	
							End If
							
							
							
						%>
          					<tr>
            <td><div align="center"><%=digito(COD_FORMANDO)%></div></td>
            <td><div align="left"><%=NOME_FORMANDO%></div></td>
            <td><div align="center"><%=ALBUM%></div>              <div align="center"></div>              <div align="center"></div></td>
            <td><div align="center" class="textoMedio"><strong>
			<%If int(TIPO) = 0 Then%>
				<%If int(ENCADERNACOES_EXTRAS) = 1 Then%>
					&Aacute;lbum Duplo
				<%Else%>
					&Aacute;lbum
				<%End If%>
			<%Elseif int(TIPO) = 1 Then%>
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
				<span style="color:#FF0000"><%=rsProdAgrePend("NOME_PRODUTO")%>					</span>
			<%
				End If
			rsProdAgrePend.Close
			Set rsProdAgrePend = Nothing
			End If%>
			</strong></div></td>
            <td><table width="98%" border="0" cellspacing="0" cellpadding="0">
                            
							<%If Len(DATA_RETORNO) < 3 Then
							  ERROR_LINHA = 2								
							%>
                            <tr>
                              <td width="13%"><img src="images/icons/icon_red.gif" width="17" height="15" /></td>
                              <td width="87%">Item n&atilde;o consta no estoque</td>
                            </tr>
                            <%Elseif TIPO = 1 Then
							  ERROR_LINHA = 2
							%>
                            <tr>
                              <td width="13%"><img src="images/icons/icon_red.gif" width="17" height="15" /></td>
                              <td width="87%">Não é possível armazenar ficha de identificação</td>
                            </tr>
          					<%Elseif TRANSFERIDO = 4 Then
							  ERROR_LINHA = 2	
							%>
							<tr>
                              <td><img src="images/icons/icon_red.gif" width="17" height="15" /></td>
                              <td>Item consta como retirado para coleta</td>
                            </tr>
                            <%Elseif ITEM_VENDIDO = 1 Then
							  ERROR_LINHA = 2	
							%>
                            <tr>
                              <td><img src="images/icons/icon_red.gif" width="17" height="15" /></td>
                              <td>Item consta como vendido, por tanto não está no estoque</td>
                            </tr>
                            
							<%Elseif COD_CAIXA_ARMAZENADO > 0 Then%>
	                       <tr>
                              <td><div align="center">
							  <img src="images/icons/icon_green.gif" width="16" height="15" /></div></td>
                              <td>Transfer&ecirc;ncia interna</td>
    					  </tr>
                            <%Else%>
                            <tr>
                              <td colspan="2"><div align="center">
							  <img src="images/icons/icon_green.gif" width="16" height="15" /></div></td>
                             </tr>
							<%End If%>
	
                          </table></td>
          	              <td><div align="center">
                          <a href="epv_estocagem_produtos_itens.asp?comando=excluir&cod_etiqueta=<%=COD_FORMANDO&TIPO%>&COD_CAIXA_ARMAZENADO=<%=Request.QueryString("COD_CAIXA_ARMAZENADO")%>&COD_SALA_ARMAZENADO=<%=Request.QueryString("COD_SALA_ARMAZENADO")%>"><img src="images/icons/icon_excluir.png" width="16" height="16" border="0" /></a>
						  <input name="COD_MOVIMENTACAO_<%=COD_FORMANDO%>_<%=TIPO%>" value="<%=COD_MOVIMENTACAO%>" type="hidden" />
                          <input name="COD_CAIXA_ARMAZENADO_<%=COD_FORMANDO%>_<%=TIPO%>" value="<%=COD_CAIXA_ARMAZENADO%>" type="hidden" />
                          <input name="TRANSFERIDO_<%=COD_FORMANDO%>_<%=TIPO%>" value="<%=TRANSFERIDO%>" type="hidden" />
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
          					  <td colspan="6" class="textoMedio" style="padding:10px"><div align="center"><strong>Itens selecionados: <span class="textoMaior"><%=CONT_ITENS%></span></strong></div></td>
       					    </tr>
						</table>
						<div align="center">
						  <p>
						  <%If ERROR_LINHA <> 2 Then%>
						  <input name="COD_VENDEDOR" value="<%=COD_VENDEDOR%>" type="hidden" />
						  <input name="NOME_VENDEDOR" value="<%=NOME_VENDEDOR%>" type="hidden" />
						  <input class="btnAvancar" name="btnAvancar" type="submit" value="" />
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
