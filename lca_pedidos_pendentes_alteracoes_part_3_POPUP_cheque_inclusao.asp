<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim COD_PEDIDO
COD_PEDIDO = Request.QueryString("cod_pedido")
'GRAVA DIVERSOS CHEQUES //////////////////////
'1ª Parte
If Request.Form("DIVERSOS_CHEQUES") <> Empty then

Session("QTD_CHEQUES")			=	Request.Form("QTD_CHEQUES")
Session("NOME_TP2")				=	Request.Form("NOME_TP2")
Session("SE_TERCEIRO_TP2")		=	Request.Form("SE_TERCEIRO_TP2")
Session("FONE_TP2")				=	Request.Form("FONE_TP2")
Session("CPF_TP2")				=	Request.Form("CPF_TP2")
Session("BANCO_TP2")			=	Request.Form("BANCO_TP2")
Session("AGENCIA_TP2")			=	Request.Form("AGENCIA_TP2")
Session("CONTA_CORRENTE_TP2")	=	Request.Form("CONTA_CORRENTE_TP2")
Session("CPF_TP2")				=	Replace(Session("CPF_TP2"),".","")
Session("CPF_TP2")				=	Replace(Session("CPF_TP2"),"-","")
Session("HISTORICO")			=	Request.Form("HISTORICO")
'PADRÕES
Session("VALOR_TP2")			=	Request.Form("VALOR_TP2")
Session("NUM_PRIMEIRO_CHEQUE")	= 	Request.Form("NUM_PRIMEIRO_CHEQUE")

Dim ANO_PRIMEIRO_CHEQUEqqr

If Len(ANO_PRIMEIRO_CHEQUEqqr) = 2 then
	ANO_PRIMEIRO_CHEQUEqqr = "20"&Request.Form("ANO_PRIMEIRO_CHEQUE")
else
	ANO_PRIMEIRO_CHEQUEqqr =Request.Form("ANO_PRIMEIRO_CHEQUE") 
end if

Session("DATA_INICIAL") 		= ANO_PRIMEIRO_CHEQUEqqr&"-"&Request.Form("MES_PRIMEIRO_CHEQUE")&"-"&Request.Form("DIA_PRIMEIRO_CHEQUE") 
Session("PREFIXO_PRIMEIRO_CHEQUE")=	Ucase(Request.Form("PREFIXO_PRIMEIRO_CHEQUE"))


end if


'------------------------------------------------
'GRAVA CHEQUE INDIVIDUAL//////////
if Request.Form("GRAVAR_CHEQUE_INDIVIDUAL")	<> Empty then

Dim sql, BANCO, AGENCIA, CONTA_CORRENTE, NUMERO_CHEQUE, VALOR_IN, BOM_PARA_DIA, BOM_PARA_MES, BOM_PARA_ANO, NOME, CPF, SE_TERCEIRO, FONE, HISTORICO

BANCO=Request.Form("BANCO")
AGENCIA=Request.Form("AGENCIA")
CONTA_CORRENTE=Request.Form("CONTA_CORRENTE")
NUMERO_CHEQUE=Request.Form("NUMERO_CHEQUE")
VALOR_IN=Request.Form("VALOR")
BOM_PARA_DIA=Request.Form("BOM_PARA_DIA")
BOM_PARA_MES=Request.Form("BOM_PARA_MES")
BOM_PARA_ANO=Request.Form("BOM_PARA_ANO")
NOME=Ucase(Request.Form("NOME"))
CPF=mascaraCPF(Request.Form("CPF"))
SE_TERCEIRO=Ucase(Request.Form("SE_TERCEIRO"))
FONE=Request.Form("FONE")
HISTORICO=Request.Form("HISTORICO")

sql="INSERT INTO TB_PEDIDOS_CHEQUE ( "&_
	"COD_PEDIDO , "&_
	"BANCO , "&_
	"AGENCIA , "&_
	"CONTA_CORRENTE , "&_
	"NUMERO_CHEQUE , "&_
	"VALOR , "&_
	"BOM_PARA , "&_
	"NOME , "&_
	"CPF , "&_
	"SE_TERCEIRO , "&_
	"FONE, "&_
	"VERIFICADO "&_
	" ) VALUES ( "&_
	COD_PEDIDO&", "&BANCO&", "&sqlTexto(AGENCIA)&", "&sqlTexto(CONTA_CORRENTE)&", "&sqlTexto(NUMERO_CHEQUE)&", '"&sqlNumero(VALOR_IN)&"', "&sqlTexto(BOM_PARA_ANO&"-"&BOM_PARA_MES&"-"&BOM_PARA_DIA)&", "&sqlTexto(Ucase(NOME))&", "&sqlTexto(CPF)&", "&sqlTexto(Ucase(SE_TERCEIRO))&", "&sqlTexto(FONE)&", 1)"

	AbrirConexao()
		objConexao.Execute(sql)

		'Set rs = objConexao.Execute("SELECT MAX(cod_pedido_cheque) as cod_pedido_cheque WHERE COD_PEDIDO = "'&sqlNumero(COD_PEDIDO)) 
		'if not rs.Eof then
		''	sql = "INSERT INTO TB_PEDIDOS_CHEQUE_HISTORICO ( cod_pedido_cheque, historico, cod_usuario) VALUES ("&sqlNumero(rs("cod_pedido_cheque"))&", "&sqlTexto(HISTORICO)&", "&sqlNumero(session("cod_usuario"))&" )"
		'end if
		'rs.close
		'set rs = nothing

		'objConexao.Execute(sql)

	FecharConexao()
	Response.Redirect("lca_pedidos_pendentes_alteracoes_part_3.asp?cod_pedido="&COD_PEDIDO&"&msg=Cheque incluído com sucesso!")

end if

if Request.Form("GRAVAR_DIVERSOS_CHEQUES_FINAL") <> Empty then
Dim sql2

AbrirConexao()
	For i=0 to Session("QTD_CHEQUES")-1
	
	sql2="INSERT INTO TB_PEDIDOS_CHEQUE ( "&_
		"COD_PEDIDO , "&_
		"BANCO , "&_
		"AGENCIA , "&_
		"CONTA_CORRENTE , "&_
		"NUMERO_CHEQUE , "&_
		"VALOR , "&_
		"BOM_PARA , "&_
		"NOME , "&_
		"CPF , "&_
		"SE_TERCEIRO , "&_
		"FONE, "&_
		"VERIFICADO "&_
		" ) VALUES ( "&_
		COD_PEDIDO&", "&Session("BANCO_TP2")&", "&sqlTexto(Session("AGENCIA_TP2"))&", "&sqlTexto(Session("CONTA_CORRENTE_TP2"))&", "&sqlTexto(Request.Form("NUM_CHEQUES_"&i+1))&", "&sqlNumero(Request.Form("VALOR_CHEQUE_"&i+1))&", "&sqlTexto(Request.Form("BOM_PARA_ANO_"&i+1)&"-"&Request.Form("BOM_PARA_MES_"&i+1)&"-"&Request.Form("BOM_PARA_DIA_"&i+1))&", "&sqlTexto(Ucase(Session("NOME_TP2")))&", "&sqlTexto(Session("CPF_TP2"))&", "&sqlTexto(Ucase(Session("SE_TERCEIRO_TP2")))&", "&sqlTexto(Session("FONE_TP2"))&", 1)"
	
	objConexao.Execute(sql2)
	set rs2 = objConexao.Execute("SELECT MAX(cod_pedido_cheque) as cod_pedido_cheque FROM TB_PEDIDOS_CHEQUE WHERE COD_PEDIDO = "&sqlNumero(COD_PEDIDO)) 
	if not rs2.Eof then
		sql2 = "INSERT INTO TB_PEDIDOS_CHEQUE_HISTORICO ( cod_pedido_cheque, historico, cod_usuario) VALUES ("&sqlNumero(rs2("cod_pedido_cheque"))&", "&sqlTexto(SESSION("HISTORICO"))&", "&sqlNumero(session("cod_usuario"))&" )"
	end if
	rs2.close
	set rs2 = nothing

	objConexao.Execute(sql2)

	sql2 = ""
	Next
FecharConexao()
Response.Redirect("lca_pedidos_pendentes_alteracoes_part_3.asp?cod_pedido="&COD_PEDIDO&"&msg=Cheques incluídos com sucesso!")

end if
'------------------------------------------------


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Cadastro de Produto em pedido</title>
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<style type="text/css">
<!--
body {
	background-color: #FFFFFF;
}
-->
</style></head>
<script>
function validaForm(){

		d = document.Cheque1
		if (d.NOME.value == ""){
		alert("Digite o nome do titular do cheque!");	
		d.NOME.focus();
		return false;
		}
		if (d.FONE.value == ""){
		alert("Digite o telefone do titular do cheque!");	
		d.FONE.focus();
		return false;
		}		
		if (d.BANCO.value == ""){
		alert("Digite o número do banco!");	
		d.BANCO.focus();
		return false;
		}		
		if (d.AGENCIA.value == ""){
		alert("Digite o número da agência!");	
		d.AGENCIA.focus();
		return false;
		}
		if (d.CONTA_CORRENTE.value == ""){
		alert("Digite o nome do formando!");	
		d.CONTA_CORRENTE.focus();
		return false;
		}
		if (d.NUMERO_CHEQUE.value == ""){
		alert("Digite o número do cheque!");	
		d.NUMERO_CHEQUE.focus();
		return false;
		}		
		if (d.VALOR.value == ""){
		alert("Digite o valor do cheque!");	
		d.VALOR.focus();
		return false;
		}	
		
		if (d.NUMERO_CHEQUE.value == ""){
		alert("Digite o número do cheque");	
		d.NUMERO_CHEQUE.focus();
		return false;
		}	
		if (d.BOM_PARA_DIA.value == "" || d.BOM_PARA_MES.value == ""  || d.BOM_PARA_ANO.value == ""){
		alert("Digite a data completa do campo Bom Para!");	
		d.BOM_PARA_DIA.focus();
		return false;
		}	

}
function validaForm2(){

		d = document.addcheques_verios
		if (d.NOME_TP2.value == ""){
		alert("Digite o nome do titular do cheque!");	
		d.NOME_TP2.focus();
		return false;
		}
		
		if (d.FONE_TP2.value == ""){
		alert("Digite o telefone do titular do cheque!");	
		d.FONE_TP2.focus();
		return false;
		}
		
		if (d.BANCO_TP2.value == ""){
		alert("Digite o banco!");	
		d.BANCO_TP2.focus();
		return false;
		}
		
		if (d.AGENCIA_TP2.value == ""){
		alert("Digite a agência!");	
		d.AGENCIA_TP2.focus();
		return false;
		}
		if (d.CONTA_CORRENTE_TP2.value == ""){
		alert("Digite a conta corrente!");	
		d.CONTA_CORRENTE_TP2.focus();
		return false;
		}		
		if (d.NUM_PRIMEIRO_CHEQUE.value == ""){
		alert("Digite o número do 1º Cheque!");	
		d.NUM_PRIMEIRO_CHEQUE.focus();
		return false;
		}	
		if (d.DIA_PRIMEIRO_CHEQUE.value == "" || d.MES_PRIMEIRO_CHEQUE.value == "" || d.ANO_PRIMEIRO_CHEQUE.value == ""){
		alert("Digite a data completa do 1º Cheque!");	
		d.DIA_PRIMEIRO_CHEQUE.focus();
		return false;
		}	
		if (d.VALOR_TP2.value == ""){
		alert("Digite o valor do cheque!");	
		d.VALOR_TP2.focus();
		return false;
		}
}
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

function YNconfirm(txt){
return typeof(suporteVBscript)=="undefined"?confirm(txt):VBconfirm(txt)==6
}

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


<script language="vbscript">
suporteVBscript=1
function VBconfirm(mensagem)

 VBconfirm=msgbox(mensagem,vbyesno)
end function

</script>
<script>
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

//Formata número tipo moeda usando o evento onKeyDown

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
<body>

<!--
<form name="cartao" style="margin:0px" action="?cod_pedido=<%=COD_PEDIDO%>&cod_pedido_cartao=<%=COD_PEDIDO_CARTAO%>&comando=<%=COMANDO%>" target="_parent" onsubmit="return verificaForm()" method="post">-->
<br />
<table width="448" border="0" align="center" cellpadding="2" cellspacing="2">
			  
            <tr>
              <td>
		  <table width="100%" align="center" cellpadding="0" cellspacing="0" id="tb_p">

<tr>
<td width="100%" align="center">
<div id="tb_principal">
  <div id="linha_preta"></div>
  <%If Request.Form("QTD_CHEQUES") > 1 then%>
 <br /> 
 <span class="msg">Todos os valores preenchidos podem ser alterados!</span>
		<form method="post" action="lca_pedidos_pendentes_alteracoes_part_3_POPUP_cheque_inclusao.asp?cod_pedido=<%=COD_PEDIDO%>" target="_parent">
                    <table width="100%" border="0" align="center" cellpadding="3" cellspacing="3">
                    <tr>
                        <td width="69" height="29">&nbsp;</td>
                      <td width="60"><div align="center"><strong>N&uacute;mero:</strong></div></td>
                      <td width="154"><div align="center"><strong>Bom para:</strong></div></td>
                      <td width="116"><div align="center"><strong>Valor:</strong></div></td>
                      </tr>
					 			 
					  <%
					  Dim DIA_YY, MES_YY, ANO_YY, DATA_COMPLETA_YY, DIAS_SOMADOS, MES_ATUAL_VERIFICA_ANO_BISSEXTO
					  DATA_COMPLETA_YY = Cdate(Session("DATA_INICIAL"))
					  DIAS_SOMADOS = 0
					  For bi = 0 to Request.Form("QTD_CHEQUES")-1  %>
					  <tr>
                        <td width="69" height="29"><em><%=bi+1%>&ordm; Cheque</em></td>
                        <td width="60"><div align="center">
<input name="NUM_CHEQUES_<%=bi+1%>" type="text" value="<%=Session("PREFIXO_PRIMEIRO_CHEQUE")&(Session("NUM_PRIMEIRO_CHEQUE")+bi)%>" size="7" maxlength="25" />
                        </div></td>
                        <td width="154">
                          <div align="center">
						  <%
						DIA_YY = Day(DATA_COMPLETA_YY)
						MES_YY = Month(DATA_COMPLETA_YY)  
						ANO_YY = Year(DATA_COMPLETA_YY)
						'Response.Write(DATA_COMPLETA_YY)
						
						MES_ATUAL_VERIFICA_ANO_BISSEXTO	= month(DATA_COMPLETA_YY)
						
						if MES_ATUAL_VERIFICA_ANO_BISSEXTO = 2 then
							
							if year(DATA_COMPLETA_YY) Mod 4 = 0 then
								DATA_COMPLETA_YY = DATA_COMPLETA_YY + 29
							else
								DATA_COMPLETA_YY = DATA_COMPLETA_YY + 28
							end if							
						
						else
						
							DATA_COMPLETA_YY = DATA_COMPLETA_YY + QuantidadeDiasMes(DATA_COMPLETA_YY)
						
						end if
					
						  
						  %>
<input name="BOM_PARA_DIA_<%=bi+1%>" type="text" value="<%=Day(Session("DATA_INICIAL"))%>" size="2" maxlength="2" /> 
/ 
<input name="BOM_PARA_MES_<%=bi+1%>" type="text" value="<%=MES_YY%>" size="2" maxlength="2" /> 
/ 
<input name="BOM_PARA_ANO_<%=bi+1%>" type="text" value="<%=ANO_YY%>" size="4" maxlength="4" />
                        </div></td>
                        <td width="116"><div align="center">R$ 
                          <input name="VALOR_CHEQUE_<%=bi+1%>" type="text" onKeydown="Formata(this,20,event,2)" value="<%=Session("VALOR_TP2")%>" size="10" maxlength="35" />
                        </div></td>
                      </tr>
                      <%Next%>
					</table>
                    <p align="center">
                      <input class="btnGravar" type="submit" name="GRAVAR_DIVERSOS_CHEQUES_FINAL" value=" " />
				</p>
		</form>
<%else%>
				<%if Len(Request.Form("QTD_CHEQUES_TITULAR")) => 1 then%>
			
			<%if Request.Form("QTD_CHEQUES_TITULAR") = 1 then%>	
				<form method="post" action="lca_pedidos_pendentes_alteracoes_part_3_POPUP_cheque_inclusao.asp?cod_pedido=<%=COD_PEDIDO%>" name="Cheque1" target="_parent">
	
				 <table width="100%" border="0" align="center" cellpadding="2" cellspacing="2">
             <tr>
              <td width="178">
              <div align="left"><strong>Nome Titular:</strong></div>             </td>
              <td width="246"><div align="left">
                <input name="NOME" type="text" size="40" maxlength="80" />
              </div></td>
            </tr>
			<tr>
              <td>
              <div align="left"><strong>Se for de terceiro:</strong></div>              </td>
              <td><div align="left">
                <input name="SE_TERCEIRO" type="text" size="40" maxlength="15" />
              </div></td>
            </tr>
			<tr>
              <td>
              <div align="left"><strong>Fone:</strong></div>            </td>
              <td><div align="left">
                <input name="FONE" type="text" size="40" maxlength="20" />
              </div></td>
            </tr>
			<tr>
              <td>
              <div align="left"><strong>CPF do Titular:</strong></div>              </td>
              <td><div align="left">
                <input name="CPF" type="text" onKeyPress="MascaraCPF(this)" size="40" maxlength="14" />
              </div></td>
            </tr>
            <tr>
              <td><div align="left"><strong>Banco (C&oacute;d.):</strong></div></td>
              <td><div align="left">
                <input name="BANCO" type="text" size="40" maxlength="8" />
              </div></td>
            </tr>
            <tr>
              <td><div align="left"><strong>Ag&ecirc;ncia:</strong></div></td>
              <td><div align="left">
                <input name="AGENCIA" type="text" size="40" maxlength="35" />
              </div></td>
            </tr>
			            <tr>
              <td><div align="left"><strong>Conta Corrente:</strong></div></td>
              <td><div align="left">
                <input name="CONTA_CORRENTE" type="text" size="40" maxlength="40" />
              </div></td>
            </tr>
			<tr>
              <td><div align="left"><strong>N&uacute;mero do Cheque:</strong></div></td>
              <td><div align="left">
                <input name="NUMERO_CHEQUE" type="text" size="40" maxlength="40" />
              </div></td>
            </tr>
            <tr>
              <td><div align="left"><strong>Valor:</strong></div></td>
              <td><div align="left">
                <input name="VALOR" onKeydown="Formata(this,20,event,2)" type="text" size="40" maxlength="40" />
              </div></td>
            </tr>
			<tr>
              <td><div align="left"><strong>Bom para:</strong></div></td>
              <td><div align="left">
                  <input name="BOM_PARA_DIA" type="text" size="2" maxlength="2" />
              /
              <input name="BOM_PARA_MES" type="text" size="2" maxlength="2" />
              /
              <input name="BOM_PARA_ANO" type="text" size="4" maxlength="4" />
              </div></td>
			</tr>
			<tr>
              <td><div align="left"><strong>Histórtico:</strong></div></td>
              <td><div align="left">
                   <textarea rows='2' name='historico' class='inputTamanho3' maxlength="500">
                   	
                   </textarea>
               </td>
			</tr>
			  
            <tr>
              <td colspan="2"><div align="center"><span style="padding-top:10px">
             <input name="button3" tabindex="9887" class="btnVoltar" type="button" onClick="history.go(-1)" value=" " />
             </span>
                  <input type="submit" name="GRAVAR_CHEQUE_INDIVIDUAL" class="btnGravar" onclick="return validaForm()" value=" " />
              </div></td>
            </tr>
            <tr>
              <td colspan="2">&nbsp;</td>
            </tr>
  </table>
</form>
			<%else%>
				<form method="post" name="addcheques_verios">
	
				 <table width="100%" border="0" align="center" cellpadding="3" cellspacing="3">
           <tr>
              <td width="171">
              <div align="left"><strong>Quant. de Cheques:</strong></div>              </td>
              <td width="246"><div align="left">
                <input name="QTD_CHEQUES" value="<%=Request.Form("QTD_CHEQUES_TITULAR")%>" type="hidden" />
              <%=Request.Form("QTD_CHEQUES_TITULAR")%></div></td>
           </tr>
		     <tr>
              <td width="171">
              <div align="left"><strong>Nome Titular:</strong></div>              </td>
              <td width="246"><div align="left">
                <input name="NOME_TP2" type="text" size="40" maxlength="40" />
              </div></td>
            </tr>
			<tr>
              <td>
              <div align="left"><strong>Se for de terceiro:</strong></div>              </td>
              <td><div align="left">
                <input name="SE_TERCEIRO_TP2" type="text" size="40" maxlength="10" />
              </div></td>
            </tr>
			<tr>
              <td>
              <div align="left"><strong>Fone:</strong></div>             </td>
              <td><div align="left">
                <input name="FONE_TP2" type="text" size="40" maxlength="15" />
              </div></td>
            </tr>
			<tr>
              <td>
              <div align="left"><strong>CPF do Titular:</strong></div>              </td>
              <td><div align="left">
                <input name="CPF_TP2" type="text" onKeyPress="MascaraCPF(this)" size="40" maxlength="14" />
              </div></td>
            </tr>
            <tr>
              <td><div align="left"><strong>Banco (C&oacute;d.):</strong></div></td>
              <td><div align="left">
                <input name="BANCO_TP2" type="text" size="40" />
              </div></td>
            </tr>
            <tr>
              <td><div align="left"><strong>Ag&ecirc;ncia:</strong></div></td>
              <td><div align="left">
                <input name="AGENCIA_TP2" type="text" size="40" maxlength="15" />
              </div></td>
            </tr>
			<tr>
              <td><div align="left"><strong>Conta Corrente:</strong></div></td>
              <td><div align="left">
                <input name="CONTA_CORRENTE_TP2" type="text" size="40" maxlength="15" />
              </div></td>
            </tr>
            <tr>
              <td><div align="left"><strong>N&uacute;mero do 1&ordm; Cheque:</strong></div></td>
              <td><div align="left">
                <input onKeyPress="mascara(this,soNumeros)" name="NUM_PRIMEIRO_CHEQUE" type="text" size="5" maxlength="15" />
             &nbsp;<span class="style7">  (Somente n&uacute;meros</span><span class="style7">)</span></div></td>
            </tr>
			<tr>
              <td><div align="left"><strong>Bom para do 1&ordm; Cheque:</strong></div></td>
              <td><div align="left">
                <input onKeyPress="mascara(this,soNumeros)" name="DIA_PRIMEIRO_CHEQUE" type="text" size="2" maxlength="2" />
               / 
             <input onKeyPress="mascara(this,soNumeros)" name="MES_PRIMEIRO_CHEQUE" type="text" size="2" maxlength="2" />
               / 
               <input onKeyPress="mascara(this,soNumeros)" name="ANO_PRIMEIRO_CHEQUE" type="text" size="4" maxlength="4" />
              </div></td>
            </tr>
            <tr>
              <td><div align="left"><strong>Histórtico:</strong></div></td>
              <td><div align="left">
                   <textarea rows='2' name='historico' class='inputTamanho2' maxlength="500"></textarea>
               </td>
			</tr>
			<tr>
              <td><div align="left"><strong>Valor por Cheque:</strong></div></td>
              <td><div align="left">
                <input name="VALOR_TP2" onKeydown="Formata(this,20,event,2)" type="text"  size="40" maxlength="15" />
              </div></td>
            </tr>
			  
            <tr>
              <td colspan="2"><div align="center" style="padding-top:10px">&nbsp;
                <input name="button22" tabindex="9887" class="btnVoltar" type="button" onclick="history.go(-1)" value=" " />
                <input type="submit" class="btnAvancar" name="DIVERSOS_CHEQUES" onclick="return validaForm2()" value=" " />
              </div></td>
            </tr>
  </table>
</form>		
			<%end if%>
<%else%>
					<form method="post" name="cheque1">
					  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="3"><tr>
              <td colspan="2" valign="middle"><div align="left">
                <p>&nbsp;</p>
                <p>&nbsp;</p>
                <p>&nbsp;</p>
              </div></td>
              </tr>
            <tr>
              <td width="357" valign="middle"><div align="right">
                     <p><strong>Quantidade de Cheques do Mesmo Titular</strong></p>
              </div></td>
              <td width="189" valign="middle"><div align="left">
                <select name="QTD_CHEQUES_TITULAR">
                  <option selected="selected" value="">---</option>
                  <option value="1">1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                  <option value="10">10</option>
                  <option value="11">11</option>
                  <option value="12">12</option>
				  <option value="13">13</option>
				  <option value="14">14</option>
				   <option value="15">15</option>
				<option value="16">16</option>
				<option value="17">17</option>
				<option value="18">18</option>
				<option value="19">19</option>
				<option value="20">20</option>
                </select>
              </div></td>
            </tr>
			  
            <tr>
              <td colspan="2"><div align="center">
                <p>&nbsp;</p>
                <p><span style="padding-top:10px">
                  <input name="button2" tabindex="9887" class="btnVoltar" type="button" onclick="history.go(-1)" value=" " />
                  </span>
                    <input type="submit" name="AVANCA_CHEQUE1" class="btnAvancar" value=" " />
                  &nbsp;</p>
              </div></td>
            </tr>
          </table>
		</form>
                    <p>
<%end if%>
<%end if%>
</p>
</div></td>
</tr>
<tr>
  <td align="center">&nbsp;</td>
</tr>
</table>
		  
		  
</td>
            </tr>
</table>
</body>
</html>
