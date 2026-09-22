<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim sql
Dim COD_SALA_ARMAZENADO, NOME_SALA_ARMAZENADO, COD_CAIXA_ARMAZENADO, DESCRICAO_CAIXA

COD_SALA_ARMAZENADO		= Request.QueryString("cod_sala_armazenado")
NOME_SALA_ARMAZENADO	= Request.QueryString("nome_sala_armazenado")
DESCRICAO_CAIXA			= Request.Form("DESCRICAO_CAIXA")

If DESCRICAO_CAIXA <> Empty Then
	AbrirConexao()
		sql = 	"INSERT INTO TB_LOCAL_ARMAZENAGEM_SALA_CAIXA (COD_SALA_ARMAZENADO,DESCRICAO_CAIXA) "&_
				"VALUES ("&COD_SALA_ARMAZENADO&", "&sqlTexto(DESCRICAO_CAIXA)&")"
		objConexao.Execute(sql)
	FecharConexao()
	
	Response.Redirect("epv_espaco_de_armazenamento_salas.asp?cod_sala_armazenado="&COD_SALA_ARMAZENADO&"&nome_sala_armazenado="&NOME_SALA_ARMAZENADO&"&msg=Caixa "&DESCRICAO_CAIXA&" incluída com sucesso!")
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
d = document.form_popUp

		if (d.DESCRICAO_CAIXA.value == ""){
		alert("Digite a descrição da caixa!");	
		return false;
		}else{
		return true;	
		}
		
}
function focusIni(){
	document.form_popUp.DESCRICAO_CAIXA.focus();
}
</script>

<script LANGUAGE="JavaScript">
<!-- Script que formata Valores em reais ao digitar -->

<!-- Início da Função FormataReais --> 
function FormataReais(fld, milSep, decSep, e) {
var sep = 0;
var key = '';
var i = j = 0;
var len = len2 = 0;
var strCheck = '0123456789';
var aux = aux2 = '';
var whichCode = (window.Event) ? e.which : e.keyCode;
if (whichCode == 13) return true;
key = String.fromCharCode(whichCode);  // Valor para o código da Chave
if (strCheck.indexOf(key) == -1) return false;  // Chave inválida
len = fld.value.length;
for(i = 0; i < len; i++)
if ((fld.value.charAt(i) != '0') && (fld.value.charAt(i) != decSep)) break;
aux = '';
for(; i < len; i++)
if (strCheck.indexOf(fld.value.charAt(i))!=-1) aux += fld.value.charAt(i);
aux += key;
len = aux.length;
if (len == 0) fld.value = '';
if (len == 1) fld.value = '0'+ decSep + '0' + aux;
if (len == 2) fld.value = '0'+ decSep + aux;
if (len > 2) {
aux2 = '';
for (j = 0, i = len - 3; i >= 0; i--) {
if (j == 3) {
aux2 += milSep;
j = 0;
}
aux2 += aux.charAt(i);
j++;
}
fld.value = '';
len2 = aux2.length;
for (i = len2 - 1; i >= 0; i--)
fld.value += aux2.charAt(i);
fld.value += decSep + aux.substr(len - 2, len);
}
return false;
}
//Fim da Função FormataReais -->
</script>
<body onload="focusIni()">
<form name="form_popUp" style="margin:0px" action="?cod_sala_armazenado=<%=COD_SALA_ARMAZENADO%>&nome_sala_armazenado=<%=NOME_SALA_ARMAZENADO%>" target="_parent" onsubmit="return verificaForm()" method="post">
<br />
<table width="402" border="0" align="center" cellpadding="2" cellspacing="2">
             <tr>
              <td width="142">
              <div align="left"><strong>Sala:</strong></div>             </td>
              <td width="246"><%=NOME_SALA_ARMAZENADO%></td>
    </tr>
			<tr>
              <td>
              <div align="left"><strong>Descri&ccedil;&atilde;o Caixa:</strong></div>              </td>
              <td><div align="left">
                <label>
                <input type="text" name="DESCRICAO_CAIXA" id="textfield" class="inputTamanho2" />
                </label>
              </div></td>
            </tr>
			  
            <tr>
              <td colspan="2"><div align="center">
                <p style="margin:0px">&nbsp;</p>
                <p style="margin:0px">
                  <input name="btnAvancar10" class="btnGravar" type="submit" id="btnAvancar" value=" " />
                  </p>
                <p style="margin:0px">&nbsp;</p>
              </div></td>
            </tr>
    </table>
</form>
</body>
</html>