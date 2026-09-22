<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim COD_PEDIDO, COD_PEDIDO_CARTAO, comando, sql
COD_PEDIDO 			= Request.QueryString("cod_pedido")
COD_PEDIDO_CARTAO 	= Request.QueryString("cod_pedido_cartao")
comando				= Request.QueryString("comando")


Dim CARTAO, VALOR_TOTAL, PARCELAS, NOME_TITULAR, GRAU_PARENTESCO, AUTENTICACAO, DOCUMENTO, DATA_AUTENTICACAO
DATA_AUTENTICACAO = ""

If Request.Form("btnAvancar10") <> Empty and comando = "incluir" Then
	incluirBanco()
Elseif Request.Form("btnAvancar10") <> Empty and comando = "alterar" Then
	alteraBanco()
Elseif Request.Form("btnAvancar10") = Empty and comando = "alterar" Then
	lerBanco()
End iF
	

Sub lerBanco()

	Dim rsCartao
	
	AbrirConexao()
	
	Set rsCartao = objConexao.Execute("SELECT * FROM TB_PEDIDOS_CARTAO WHERE COD_PEDIDO_CARTAO="&COD_PEDIDO_CARTAO&" LIMIT 1")
	
		If not rsCartao.Eof Then
			
			CARTAO				= rsCartao("CARTAO")
			VALOR_TOTAL			= rsCartao("VALOR_TOTAL")
			PARCELAS			= rsCartao("PARCELAS")
			NOME_TITULAR		= rsCartao("NOME_TITULAR")
			GRAU_PARENTESCO		= rsCartao("GRAU_PARENTESCO")
			AUTENTICACAO		= rsCartao("AUTENTICACAO")
			DOCUMENTO			= rsCartao("DOCUMENTO")
			DATA_AUTENTICACAO	= rsCartao("DATA_AUTENTICACAO")
			
		End If	
	
	
	rsCartao.Close
	Set rsCartao = Nothing
	
	FecharConexao()
	
End Sub


Sub lerForm()
	
	CARTAO				= Request.Form("CARTAO")
	VALOR_TOTAL			= Request.Form("VALOR_TOTAL")
	PARCELAS			= Request.Form("PARCELAS")
	NOME_TITULAR		= Ucase(Request.Form("NOME_TITULAR"))
	GRAU_PARENTESCO		= Ucase(Request.Form("GRAU_PARENTESCO"))
	AUTENTICACAO		= Request.Form("AUTENTICACAO")
	DOCUMENTO			= Request.Form("DOCUMENTO")
	DATA_AUTENTICACAO	= Request.Form("DATA_AUTENTICACAO_ANO")&"-"&Request.Form("DATA_AUTENTICACAO_MES")&"-"&Request.Form("DATA_AUTENTICACAO_DIA")
	
End Sub

Sub incluirBanco()
	lerForm()
	sql = 	"INSERT INTO TB_PEDIDOS_CARTAO ("&_
			"COD_PEDIDO, "&_
			"CARTAO, "&_
			"VALOR_TOTAL, "&_
			"PARCELAS, "&_
			"NOME_TITULAR, "&_
			"GRAU_PARENTESCO, "&_
			"AUTENTICACAO ,"&_
			"DOCUMENTO ,"&_
			"DATA_AUTENTICACAO ,"&_
			"VERIFICADO "&_
			")"&_
			"VALUES ("&_
			COD_PEDIDO&","&sqlTexto(CARTAO)&",'"&sqlNumero(VALOR_TOTAL)&"','"&PARCELAS&"',"&sqlTexto(NOME_TITULAR)&","&sqlTexto(GRAU_PARENTESCO)&",'"&AUTENTICACAO&_
			"','"&DOCUMENTO&"','"&DATA_AUTENTICACAO&"',1)"
			
			AbrirConexao()
				objConexao.Execute(sql)
			FecharConexao()
			
			Response.Redirect("lca_pedidos_pendentes_alteracoes_part_3.asp?cod_pedido="&COD_PEDIDO&"&msg=Cartão incluído com sucesso!")	
End Sub

Sub alteraBanco()
		
			lerForm()
			
			sql =	"UPDATE TB_PEDIDOS_CARTAO SET "&_
					"CARTAO = 			"&sqlTexto(CARTAO)&", "&_
					"VALOR_TOTAL = 		'"&sqlNumero(VALOR_TOTAL)&"', "&_
					"PARCELAS = 		'"&PARCELAS&"', "&_
					"NOME_TITULAR = 	"&sqlTexto(NOME_TITULAR)&", "&_
					"GRAU_PARENTESCO = 	"&sqlTexto(GRAU_PARENTESCO)&", "&_
					"AUTENTICACAO = 	"&sqlTexto(AUTENTICACAO)&", "&_
					"DOCUMENTO = 		"&sqlTexto(DOCUMENTO)&", "&_
					"VERIFICADO = 		1, "&_
					"DATA_AUTENTICACAO = '"&DATA_AUTENTICACAO&"' WHERE COD_PEDIDO_CARTAO = "&COD_PEDIDO_CARTAO&" LIMIT 1"
			
			AbrirConexao()
				objConexao.Execute(sql)
			FecharConexao()
			
			Response.Redirect("lca_pedidos_pendentes_alteracoes_part_3.asp?cod_pedido="&COD_PEDIDO&"&msg=Cartão alterado com sucesso!")	

End Sub

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Cadastro de Produto em pedido</title>
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<style type="text/css">

body {
	background-color: #FFFFFF;
}

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
d = document.cartao


		if (!d.CARTAO[0].checked && !d.CARTAO[1].checked && !d.CARTAO[2].checked){
		alert("Escolha a operadora do Cartão!");	
		return false;
		}
		
		if (d.VALOR_TOTAL.value == ""){
		alert("Digite o valor total!");	
		d.VALOR.focus();
		return false;
		}
		
		if (d.PARCELAS.value == ""){
		alert("Digite a quantidade de parcelas!");	
		d.PARCELAS.focus();
		return false;
		}
		
		if (d.NOME_TITULAR.value == ""){
		alert("Digite o nome do titular do Cartão!");	
		d.NOME_TITULAR.focus();
		return false;
		}
		
		if (d.AUTENTICACAO.value == ""){
		alert("Digite o número da autenticaçao!");	
		d.AUTENTICACAO.focus();
		return false;
		}
		
		if (d.DOCUMENTO.value == ""){
		alert("Digite o número do documento!");	
		d.DOCUMENTO.focus();
		return false;
		}
		
		if (d.DATA_AUTENTICACAO_DIA.value == "" || d.DATA_AUTENTICACAO_MES.value == "" || d.DATA_AUTENTICACAO_ANO.value == ""){
		alert("Digite a data completa da autenticaçao");	
		d.DATA_AUTENTICACAO_DIA.focus();
		return false;
		}
	return true;
}
function focusIni(){
	document.cartao.NOME_TITULAR.focus();
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

//Formata n�mero tipo moeda usando o evento onKeyDown

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
<form name="cartao" style="margin:0px" action="?cod_pedido=<%=COD_PEDIDO%>&cod_pedido_cartao=<%=COD_PEDIDO_CARTAO%>&comando=<%=COMANDO%>" target="_parent" onsubmit="return verificaForm()" method="post">
<br />
<table width="402" border="0" align="center" cellpadding="2" cellspacing="2">
             <tr>
              <td width="142">
              <div align="left"><strong>Nome Titular:</strong></div>             </td>
              <td width="246"><div align="left">
                <input name="NOME_TITULAR" value="<%=Ucase(NOME_TITULAR)%>" type="text" size="40" maxlength="80" />
              </div></td>
            </tr>
			<tr>
              <td>
              <div align="left"><strong>Se for de terceiro:</strong></div>              </td>
              <td><div align="left">
                <input name="GRAU_PARENTESCO" value="<%=Ucase(GRAU_PARENTESCO)%>" type="text" size="40" maxlength="15" />
              </div></td>
            </tr>
			<tr>
              <td>
              <div align="left"><strong>Cart&atilde;o</strong></div>            </td>
              <td><div align="left">
                <table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="20">
            <input name="CARTAO" id="cartao1" class="inputRadio" type="radio" <%if CARTAO = "VISA" then%> checked="checked"<%end if%> value="VISA" />
            </td>
            <td width="39"><label for="cartao1">VISA</label></td>
            <td width="20"><input id="cartao2" class="inputRadio" name="CARTAO" <%if CARTAO = "VISA ELECTRON" then%> checked="checked"<%end if%> type="radio" value="VISA ELECTRON" /></td>
            <td width="103"><label for="cartao2">VISA ELECTRON</label></td>
            <td width="20"><input id="cartao3" class="inputRadio" name="CARTAO" <%if CARTAO = "MASTERCARD" then%> checked="checked"<%end if%> type="radio" value="MASTERCARD" /></td>
            <td><label for="cartao3">MASTERCARD</label></td>
            </tr>
      </table>
              </div></td>
            </tr>
			<tr>
              <td>
              <div align="left"><strong>Parcelas</strong></div>              </td>
              <td><div align="left">
	  			<% If  validaUsuario(2,"",85, false) Then %>
                	<input name="PARCELAS" onkeypress="mascara(this,soNumeros)" value="<%=PARCELAS%>" type="text" size="40" maxlength="20" />
				<% else %>
					<%=PARCELAS%>
					<input name="PARCELAS" value="<%=PARCELAS%>" type="hidden" size="40" maxlength="20" />
				<% end if %>
              </div></td>
            </tr>
            <tr>
              <td><div align="left"><strong>Valor Total</strong></div></td>
              <td><div align="left">
				<% If  validaUsuario(2,"",85, false) Then %>
                	<input name="VALOR_TOTAL" value="<%=formatarNumero(VALOR_TOTAL)%>" onKeydown="Formata(this,20,event,2)" type="text" size="40" maxlength="8" />                
				<% else %>
					<%=VALOR_TOTAL%>
					<input name="VALOR_TOTAL" value="<%=formatarNumero(VALOR_TOTAL)%>" type="hidden" size="40" maxlength="20" />
				<% end if %>
              </div></td>
            </tr>
            <tr bgcolor="#FFFF99">
              <td bgcolor="#FFFF99"><div align="left"><strong>Autentita&ccedil;&atilde;o</strong></div></td>
              <td><div align="left">
                <input name="AUTENTICACAO" onkeypress="mascara(this,soNumeros)" value="<%=AUTENTICACAO%>" type="text" size="40" maxlength="35" />
              </div></td>
            </tr>
			<tr bgcolor="#FFFF99">
              <td><div align="left"><strong>Documento</strong></div></td>
              <td><div align="left">
                <input name="DOCUMENTO" onkeypress="mascara(this,soNumeros)" value="<%=DOCUMENTO%>" type="text" size="40" maxlength="40" />
              </div></td>
            </tr>
			<tr bgcolor="#FFFF99">
              <td><div align="left"><strong>Data da Autentica&ccedil;&atilde;o </strong></div></td>
              <td><div align="left">
			  <%If DATA_AUTENTICACAO <> "" Then%>
              <input name="DATA_AUTENTICACAO_DIA" onkeypress="mascara(this,soNumeros)" value="<%=Day(DATA_AUTENTICACAO)%>" type="text" size="2" maxlength="2" />
              /
              <input name="DATA_AUTENTICACAO_MES" onkeypress="mascara(this,soNumeros)" value="<%=Month(DATA_AUTENTICACAO)%>" type="text" size="2" maxlength="2" />
              /
              <input name="DATA_AUTENTICACAO_ANO" onkeypress="mascara(this,soNumeros)" value="<%=Year(DATA_AUTENTICACAO)%>" type="text" size="4" maxlength="4" />
			  <%Else%>
			    <input name="DATA_AUTENTICACAO_DIA" onkeypress="mascara(this,soNumeros)"  type="text" size="2" maxlength="2" />
              /
              <input name="DATA_AUTENTICACAO_MES" onkeypress="mascara(this,soNumeros)"  type="text" size="2" maxlength="2" />
              /
              <input name="DATA_AUTENTICACAO_ANO" onkeypress="mascara(this,soNumeros)" type="text" size="4" maxlength="4" />
			  <%End If%>
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
