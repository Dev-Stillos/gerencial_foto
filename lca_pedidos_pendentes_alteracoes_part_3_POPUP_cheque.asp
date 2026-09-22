<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim COD_PEDIDO, COD_PEDIDO_CHEQUE
COD_PEDIDO 			= Request.QueryString("cod_pedido")
COD_PEDIDO_CHEQUE 	= Request.QueryString("cod_pedido_cheque")

Dim NOME, BANCO, AGENCIA, CONTA_CORRENTE, NUMERO_CHEQUE, VALOR, BOM_PARA, CPF, SE_TERCEIRO, FONE

If Request.Form("NOME") <> Empty Then

	NOME				= Ucase(Request.Form("NOME"))
	BANCO				= Request.Form("BANCO")
	AGENCIA				= Request.Form("AGENCIA")	
	CONTA_CORRENTE		= Request.Form("CONTA_CORRENTE")
	NUMERO_CHEQUE		= Request.Form("NUMERO_CHEQUE")
	VALOR				= Request.Form("VALOR")
	BOM_PARA			= Request.Form("BOM_PARA_ANO")&"-"&Request.Form("BOM_PARA_MES")&"-"&Request.Form("BOM_PARA_DIA")
	CPF					= mascaraCPF(Request.Form("CPF"))
	SE_TERCEIRO			= Ucase(Request.Form("SE_TERCEIRO"))
	FONE				= Request.Form("FONE")
	FONE				= Replace(FONE,"-","")
	FONE				= Replace(FONE,"(","")
	FONE				= Replace(FONE,")","")
	FONE				= Replace(FONE,".","")
	
	AbrirConexao()
	
	Dim sql
	
	sql  = 	"UPDATE TB_PEDIDOS_CHEQUE SET BANCO = "&BANCO&", "&_
			"AGENCIA = "&sqlTexto(AGENCIA)&", "&_
			"CONTA_CORRENTE = '"&CONTA_CORRENTE&"', "&_
			"NUMERO_CHEQUE = '"&NUMERO_CHEQUE&"', "&_
			"VALOR = '"&sqlNumero(VALOR)&"', "&_
			"BOM_PARA = '"&BOM_PARA&"', "&_
			"NOME = "&sqlTexto(NOME)&", "&_
			"CPF = '"&CPF&"', "&_
			"SE_TERCEIRO = "&sqlTexto(SE_TERCEIRO)&", "&_
			"FONE = '"&FONE&"', "&_
			"VERIFICADO = '1' WHERE COD_PEDIDO_CHEQUE="&COD_PEDIDO_CHEQUE&" LIMIT 1"
	
			'Response.Write(sql)
	objConexao.Execute(sql)
	
	Response.Redirect("lca_pedidos_pendentes_alteracoes_part_3.asp?cod_pedido="&COD_PEDIDO&"&msg=Cheque alterado com sucesso!")
	FecharConexao()
End If
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
function verificaForm(){
d = document.cheque

	if (d.NOME.value == ""){
		alert('Preencha o campo NOME!');
		d.NOME.focus();
		return false;
	}
	if (d.CPF.value == ""){
		alert('Preencha o campo CPF!');
		d.CPF.focus();
		return false;
	}
	if (d.BANCO.value == ""){
		alert('Preencha o campo BANCO!');
		d.BANCO.focus();
		return false;
	}
	if (d.AGENCIA.value == ""){
		alert('Preencha o campo AGENCIA!');
		d.AGENCIA.focus();
		return false;
	}
	if (d.CONTA_CORRENTE.value == ""){
		alert('Preencha o campo CONTA_CORRENTE!');
		d.CONTA_CORRENTE.focus();
		return false;
	}
	if (d.NUMERO_CHEQUE.value == ""){
		alert('Preencha o campo NUMERO_CHEQUE!');
		d.NUMERO_CHEQUE.focus();
		return false;
	}
	if (d.VALOR.value == ""){
		alert('Preencha o campo VALOR!');
		d.VALOR.focus();
		return false;
	}
	if (d.BOM_PARA_DIA.value == "" || d.BOM_PARA_MES.value == "" || d.BOM_PARA_ANO.value == ""){
		alert('Preencha a data completa dos campos BOM PARA!');
		d.BOM_PARA_DIA.focus();
		return false;
	}
	return true;
}

function focusIni(){
	document.cheque.NOME.focus();
}
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
<body onload="focusIni()">
<form name="cheque" style="margin:0px" action="?cod_pedido=<%=COD_PEDIDO%>&cod_pedido_cheque=<%=COD_PEDIDO_CHEQUE%>" target="_parent" onsubmit="return verificaForm()" method="post">
<br />
<%
AbrirConexao()
Dim rs
Set rs = objConexao.Execute("SELECT * FROM TB_PEDIDOS_CHEQUE WHERE COD_PEDIDO_CHEQUE="&COD_PEDIDO_CHEQUE&" LIMIT 1")

If not rs.Eof Then
%>
<table width="402" border="0" align="center" cellpadding="2" cellspacing="2">
            <tr>
              <td colspan="2">
                <div align="center" class="msg"><%=Request.QueryString("msg")%></div></td>
            </tr>
			 <tr>
              <td width="142">
              <div align="left"><strong>Nome Titular:</strong></div>             </td>
              <td width="246"><div align="left">
                <input name="NOME" value="<%=Ucase(rs("NOME"))%>" type="text" size="40" maxlength="80" />
              </div></td>
            </tr>
			<tr>
              <td>
              <div align="left"><strong>Se for de terceiro:</strong></div>              </td>
              <td><div align="left">
                <input name="SE_TERCEIRO" value="<%=Ucase(rs("SE_TERCEIRO"))%>" type="text" size="40" maxlength="15" />
              </div></td>
            </tr>
			<tr>
              <td>
              <div align="left"><strong>Fone:</strong></div>            </td>
              <td><div align="left">
                <input name="FONE" value="<%=Ucase(rs("FONE"))%>" type="text" size="40" maxlength="20" />
              </div></td>
            </tr>
			<tr>
              <td>
              <div align="left"><strong>CPF do Titular:</strong></div>              </td>
              <td><div align="left">
                <input name="CPF" value="<%=mascaraCPF(rs("CPF"))%>" type="text" onKeyPress="MascaraCPF(this)" size="40" maxlength="14" />
              </div></td>
            </tr>
            <tr>
              <td><div align="left"><strong>Banco (C&oacute;d.):</strong></div></td>
              <td><div align="left">
                <input name="BANCO" onkeypress="mascara(this,soNumeros)" value="<%=rs("BANCO")%>" type="text" size="40" maxlength="8" />
              </div></td>
            </tr>
            <tr>
              <td><div align="left"><strong>Ag&ecirc;ncia:</strong></div></td>
              <td><div align="left">
                <input name="AGENCIA" value="<%=rs("AGENCIA")%>" type="text" size="40" maxlength="35" />
              </div></td>
            </tr>
			            <tr>
              <td><div align="left"><strong>Conta Corrente:</strong></div></td>
              <td><div align="left">
                <input name="CONTA_CORRENTE" value="<%=rs("CONTA_CORRENTE")%>" type="text" size="40" maxlength="40" />
              </div></td>
            </tr>
			<tr>
              <td><div align="left"><strong>N&uacute;mero do Cheque:</strong></div></td>
              <td><div align="left">
                <input name="NUMERO_CHEQUE" onkeypress="mascara(this,soNumeros)" value="<%=rs("NUMERO_CHEQUE")%>" type="text" size="40" maxlength="40" />
              </div></td>
            </tr>
            <tr>
              <td><div align="left"><strong>Valor:</strong></div></td>
              <td><div align="left">
                <input name="VALOR" onKeydown="Formata(this,20,event,2)" value="<%=formatarNumero(rs("VALOR"))%>" type="text" size="40" maxlength="40" />
              </div></td>
            </tr>
			<tr>
              <td><div align="left"><strong>Bom para:</strong></div></td>
              <td><div align="left">
              <input name="BOM_PARA_DIA" onkeypress="mascara(this,soNumeros)" value="<%=Day(rs("BOM_PARA"))%>" type="text" size="2" maxlength="2" />
              /
              <input name="BOM_PARA_MES" onkeypress="mascara(this,soNumeros)" value="<%=Month(rs("BOM_PARA"))%>" type="text" size="2" maxlength="2" />
              /
              <input name="BOM_PARA_ANO" onkeypress="mascara(this,soNumeros)" value="<%=Year(rs("BOM_PARA"))%>" type="text" size="4" maxlength="4" />
              </div></td>
			</tr>
      <tr>
        <td><div align="left"><strong>Histórico:</strong></div></td>
        <td><div align="left">
          <%
          dim qtd_historico
          Set rs2 = objConexao.Execute("SELECT count(*) FROM TB_PEDIDOS_CHEQUE_HISTORICO WHERE cod_pedido_cheque = "&sqlNumero(cod_pedido_cheque))
          if not rs2.Eof Then
            qtd_historico = rs2(0)
          end if
          rs2.close()
          Set rs2 = Nothing
          %>

          <a href="lca_pedidos_pendentes_alteracoes_part_3_POPUP_cheque_historico.asp?COD_PEDIDO_CHEQUE=<%=COD_PEDIDO_CHEQUE%>&COD_PEDIDO=<%=COD_PEDIDO%>">
            Visualizar (<%=qtd_historico%>)
          </a>
        </div></td>
      </tr>
            <tr>
              <td colspan="2"><div align="center">
                <p style="margin:0px">&nbsp;</p>
                <p style="margin:0px">
                  <input name="btnAvancar9" class="btnGravar" type="submit" id="btnAvancar" value=" " />
                  </p>
                <p style="margin:0px">&nbsp;</p>
              </div></td>
            </tr>
  </table>
<%
End If

rs.Close
Set rs = Nothing
FecharConexao()
%>
</form>
</body>
</html>
