<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>
<script language="JavaScript" type="text/javascript">
// JavaScript Document
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

//adiciona mascara de cpf
function MascaraCPF(cpf){
    if(mascaraInteiro(cpf)==false){
        event.returnValue = false;
    }    
    return formataCampo(cpf, '000.000.000-00', event);
}

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
var msgErro = 'Formato inv&aacutelido de data.';
if ((campo.value.match(expReg)) && (campo.value!='')){
var dia = campo.value.substring(0,2);
var mes = campo.value.substring(3,5);
var ano = campo.value.substring(6,10);
if(mes==4 || mes==6 || mes==9 || mes==11 && dia > 30){
campo.focus();
alert("Dia incorreto ! O m&ecircs especificado cont&eacutem no m&aacuteximo 30 dias.");
return false;
} else{
if(ano%4!=0 && mes==2 && dia>28){
campo.focus();
alert("Data incorreta! O m&ecircs especificado cont&eacutem no m&aacuteximo 28 dias.");
return false;
} else{
if(ano%4==0 && mes==2 && dia>29){
campo.focus();
alert("Data incorreta! O m&ecircs especificado cont&eacutem no m&aacuteximo 29 dias.");
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
//Formata n&uacutemero tipo moeda usando o evento onKeyDown
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
</script>
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
	
	<script type="text/javascript">	
	
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
	
	function exibirBuscaComplemento(valor){
		if (valor == 1){	
			exibeOculta('comboComplementoBusca','exibir'); 
		}else{
			exibeOculta('comboComplementoBusca','ocultar'); 
		}
	
	}
	
	</script>
    
    
    </head>
</head>

<body onKeyDown="javascript:Verificar()">

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
	
	  
 
 <h2>Fornecedores</h2><br />
 <div class="subMenu">Sub-Menu: <%If (session("cod_usuario") = 1 OR session("cod_usuario") = 16) Then%> <a href="ppv_fornecedores_formulario.asp?comando=incluir">Incluir Fornecedor</a> <% End If %> | <a href="ppv_compras_pedidos.asp">Voltar para Compras</a></div>
<%
AbrirConexao()
	
	sql = "SELECT * FROM TB_FORNECEDORES  order by razao_social"
	

	Set rs=server.CreateObject("ADODB.recordset")
	rs.CursorLocation=3  'Cursor no cliente
	rs.PageSize=50
	rs.Open sql, objConexao
	
	Dim paginacaocao
	paginacaocao = formatNumber(0&Request.QueryString("paginacao"),0)
	
	If not rs.Eof Then
	
		If paginacaocao > 0 Then
			rs.AbsolutePage=paginacaocao
		Else
			paginacaocao = 1
			rs.AbsolutePage=1
		End If	
%>
	
    <%
	Dim sucesso, codmsg
	codmsg	= CStr(""&Request.QueryString("codmsg"))
	sucesso = CStr(""&Request.QueryString("sucesso"))
		
	If sucesso = "" and codmsg <> "" Then%>
    	<div align="center" class="msgGrandeSucesso">
        	<%
			If codmsg = "1" Then
				Response.Write("Fornecedor inclu&iacutedo com sucesso!")
			Elseif codmsg = "2" Then
				Response.Write("Fornecedor alterado com sucesso!")
			Elseif codmsg = "3" Then	
				Response.Write("Fornecedor exclu&iacutedo com sucesso!")
			End If	
			%>
</div>
<%ElseIf sucesso <> "" and codmsg <> "" Then%>
        <div align="center" class="msgGrandeErro">
			<%
			If codmsg = "4" Then
				Response.Write("Fornecedor n&atildeo alterado para n&atildeo haver duplicidade de informa&ccedil&atildeo!<BR>O conte&uacutedo do campo """&Request.QueryString("chaveUnica")&""" j&aacute possui uma informa&ccedil&atildeo igual em outro Fornecedor.<br />")
			Elseif codmsg = "5" Then
				Response.Write("Fornecedor n&atildeo exclu&iacutedo, pois h&aacute Or&ccedilamentos/Pedidos/NF que dependem ou relacionam com este vendedor.<BR>Se n&atildeo desejar ter mais este Fornecedor inative ao inv&eacutes de excluir.<br />")
			
			Elseif codmsg = "900" Then
				Response.Write("Voc&ecirc n&atildeo possui acesso a VISUALIZA&Ccedil&AtildeO!")
			Elseif codmsg = "901" Then
				Response.Write("Voc&ecirc n&atildeo possui acesso a INCLUS&AtildeO!")
			Elseif codmsg = "902" Then
				Response.Write("Voc&ecirc n&atildeo possui acesso a ALTERA&Ccedil&AtildeO!")	
			Elseif codmsg = "903" Then
				Response.Write("Voc&ecirc n&atildeo possui acesso a EXCLUS&AtildeO!")
			End If					
			%>
</div>
<%End If%>
<br />

<table width="99%" border="0" align="center">
  <tr>
    <td width="47%"> 
P&aacutegina <strong><%=rs.AbsolutePage%></strong> de <%=rs.PageCount%>
</td>
    <td width="53%"><div align="right">Total <strong><%=rs.RecordCount%></strong> fornecedores</div></td>
  </tr>
</table>
<table width="99%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
<td width="7%" rowspan="2"><div align="center" style="font-weight: bold">C&oacute;d. Forn.</div></td>
<td width="31%" rowspan="2"><div align="center" style="font-weight: bold">Raz&atilde;o Social</div></td>
<td colspan="2"><div align="center"><strong>Contato</strong></div></td>

<td width="10%" rowspan="2"><div align="center" style="font-weight: bold">Data Cadastro</div></td>
<td width="17%" rowspan="2"><div align="center" style="font-weight: bold">Situa&ccedil;&atilde;o</div></td>
<td width="9%" rowspan="2"><div align="center" style="font-weight: bold"></div></td>
<td width="7%" rowspan="2">&nbsp;</td>
</tr>
<tr bgcolor="whitesmoke">
  <td width="12%"><div align="center"><strong>Nome</strong></div></td>
  <td width="7%"><div align="center"><strong>Fone</strong></div></td>
</tr>
<%
Dim x
For x=1 To rs.PageSize
If Not rs.Eof Then
%>
<tr>
  <td><div align="center"><%=formatNumber(rs(0),0)%></div></td>
  <td><div align="left"><%=rs("razao_social")%></div></td>
  <td><div align="left">
    <%=rs("contato")%>
  </div></td>
  <td><div align="left"> <%=padraoTelefone(rs("telefone"))%>	
  </div></td>

  <td><div align="center"><%=ArrumaData(rs("data_cadastro"))%></div></td>
  <td><div align="center"><%=booleano(rs("ativado"))%></div></td>
  
  <!--'alteracao para somente admin e andrea ter acesso-->

  <td width="9%"><div align="center"><%If (session("cod_usuario") = 1 OR session("cod_usuario") = 16) Then%><a href="ppv_fornecedores_consulta_produtos_fornecidos.asp?cod_fornecedor=<%=rs("cod_fornecedor")%>">Produtos Fornecidos</a><% End If %></div></td>
  <td>
  
  <table border="0" align="center" cellpadding="3" cellspacing="0"> 
  <tr>
      <td>    <a href="ppv_fornecedores_formulario.asp?comando=visualizar&cod_fornecedor=<%=rs("cod_fornecedor")%>">
  <img src="images/icons/icon_visualizar.png" alt="Visualizar Registro" title="Visualizar Registro" border="0" />
  	</a></td>
    <td><%If (session("cod_usuario") = 1 OR session("cod_usuario") = 16) Then%>
    <a href="ppv_fornecedores_formulario.asp?comando=alterar&cod_fornecedor=<%=rs("cod_fornecedor")%>">
  <img src="images/icons/icon_alterar.png" alt="Alterar Registro" title="Alterar Registro" width="16" height="16" border="0" /></a><% End If %>  </td>
    <td><%If (session("cod_usuario") = 1 OR session("cod_usuario") = 16) Then%>  <a href="ppv_fornecedores_formulario.asp?comando=excluir&cod_fornecedor=<%=rs("cod_fornecedor")%>">
  <img src="images/icons/icon_excluir.png" alt="Excluir Registro" title="Excluir Registro" border="0" />
  </a><% End If%></td>
  </tr>
</table>  </td>
  </tr>
<%
rs.MoveNext()
End If
Next
%>
<%If rs.PageCount > 1 Then %>                                  
     <tr bgcolor="whitesmoke">
      <td colspan="9">
      	   <%call paginacao(paginacaocao,rs.PageCount)%>      </td>
    </tr>   
<%End if%>
</table>	
<%
	Else
	%>
    <div align="center" class="msgGrande">
    	Nenhum Fornecedor encontrado!
    <br />
	</div>
    <%End if
	
	rs.Close
	Set rs=Nothing
	
	FecharConexao()
	%>
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