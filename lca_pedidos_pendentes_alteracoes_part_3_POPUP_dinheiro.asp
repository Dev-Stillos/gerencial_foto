<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim COD_PEDIDO, comando, sql
COD_PEDIDO 			= Request.QueryString("cod_pedido")
comando				= Request.QueryString("comando")


Dim PAGAMENTO_EM_DINHEIRO

If Request.Form("PAGAMENTO_EM_DINHEIRO") <> Empty and comando = "incluir" Then
	incluirBanco()
Elseif Request.Form("PAGAMENTO_EM_DINHEIRO") <> Empty and comando = "alterar" Then
	alteraBanco()
End iF
	
Sub incluirBanco()

	PAGAMENTO_EM_DINHEIRO = Request.Form("PAGAMENTO_EM_DINHEIRO")
	
			AbrirConexao()
				objConexao.Execute("UPDATE TB_PEDIDOS SET PAGAMENTO_EM_DINHEIRO = '"&sqlNumero(PAGAMENTO_EM_DINHEIRO)&"' WHERE COD_PEDIDO = "&COD_PEDIDO&" LIMIT 1")
			FecharConexao()
			
			Response.Redirect("lca_pedidos_pendentes_alteracoes_part_3.asp?cod_pedido="&COD_PEDIDO&"&msg=Dinheiro incluído com sucesso!")	
End Sub

Sub alteraBanco()
		
			PAGAMENTO_EM_DINHEIRO = Request.Form("PAGAMENTO_EM_DINHEIRO")
			
			AbrirConexao()
				objConexao.Execute("UPDATE TB_PEDIDOS SET PAGAMENTO_EM_DINHEIRO = '"&sqlNumero(PAGAMENTO_EM_DINHEIRO)&"' WHERE COD_PEDIDO = "&COD_PEDIDO&" LIMIT 1")
			FecharConexao()

			Response.Redirect("lca_pedidos_pendentes_alteracoes_part_3.asp?cod_pedido="&COD_PEDIDO&"&msg=Dinheiro alterado com sucesso!")	

End Sub

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
d = document.dinheiro


		if (d.PAGAMENTO_EM_DINHEIRO.value == "" || d.PAGAMENTO_EM_DINHEIRO.value == "0,00"){
			alert("Digite o valor em dinheiro!");
			d.PAGAMENTO_EM_DINHEIRO.focus();	
			return false;
		}else{
			return true;	
		}
		
	
}
function focusIni(){
	document.dinheiro.PAGAMENTO_EM_DINHEIRO.focus();
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
<form name="dinheiro" style="margin:0px" action="?cod_pedido=<%=COD_PEDIDO%>&comando=<%=COMANDO%>" target="_parent" onsubmit="return verificaForm()" method="post">
<br />
<table width="402" border="0" align="center" cellpadding="2" cellspacing="2">
             <tr>
              <td width="142">
              <div align="left"><strong>Valor em Dinheiro </strong></div>             </td>
              <td width="246"><div align="left">
                <input onKeydown="Formata(this,20,event,2)" name="PAGAMENTO_EM_DINHEIRO" value="<%=formatarNumero(Request.QueryString("valor_em_dinheiro"))%>" type="text" size="40" maxlength="80" />
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
