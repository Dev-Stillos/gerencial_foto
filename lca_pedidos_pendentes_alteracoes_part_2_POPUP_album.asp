<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim COD_PEDIDO
COD_PEDIDO = Request.QueryString("cod_pedido")
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

function confirma() {		
	d = document.produtos

	if (d.QTD_FATURADA_<%=Request.QueryString("cod_for")%>.value == ""){
		alert("Digite a quantidade faturada!\n* Caso n�o tiver digite 0 (zero).");
		d.QTD_FATURADA_<%=Request.QueryString("cod_for")%>.focus();
		return false;
	}
	
	if (d.VALOR_UNITARIO_<%=Request.QueryString("cod_for")%>.value == ""){
		
		alert("Digite o valor unit�rio vendido!\n* Caso n�o tiver digite 0 (zero).");
		d.VALOR_UNITARIO_<%=Request.QueryString("cod_for")%>.focus();
		return false;
	} else{
		if(d.QTD_FATURADA_<%=Request.QueryString("cod_for")%>.value !== "" && d.VALOR_UNITARIO_<%=Request.QueryString("cod_for")%>.value == 0){
			
			if(d.INICIAL_<%=Request.QueryString("cod_for")%>.value == d.QTD_RETIRADA_<%=Request.QueryString("cod_for")%>.value){	
				alert("Aten��o!\nN�o � poss�vel continuar, pois a quantidade inicial � igual a quantidade\nfaturada, desta forma o �lbum inteiro seria descartado, corriga antes de continuar.");
				return false;
			}
				
		}
		else if (d.QTD_FATURADA_<%=Request.QueryString("cod_for")%>.value/1 > 0 && d.VALOR_UNITARIO_<%=Request.QueryString("cod_for")%>.value == 0 ){
				alert("Aten��o!\nO campo valor unit�rio s� aceita zero quando a Qtd. Faturada � igual a 0 (zero)."); 
				d.QTD_FATURADA_<%=Request.QueryString("cod_for")%>.focus();
				return false;
				
		}	
	}
	
	if (d.QTD_FATURADA_<%=Request.QueryString("cod_for")%>.value/1 == 0 && d.VALOR_UNITARIO_<%=Request.QueryString("cod_for")%>.value > 0 ){
		alert("Aten��o!\nVoc� informou que n�o h� fotos faturadas, mas informou o valor da foto, por gentileza corriga!"); 
		d.VALOR_UNITARIO_<%=Request.QueryString("cod_for")%>.focus();
		return false;
				
	}
	
	if (d.QTD_FATURADA_<%=Request.QueryString("cod_for")%>.value == "" && d.VALOR_UNITARIO_<%=Request.QueryString("cod_for")%>.value !== ""){
		alert("Aten��o!\nVoc� informou que n�o h� fotos faturadas, mas informou o valor da foto, por gentileza corriga!"); 
		d.VALOR_UNITARIO_<%=Request.QueryString("cod_for")%>.focus();
		return false;
				
	}
	
	if(YNconfirm("Tem certeza que deseja avan�ar?")){
			return true;
	}else{
			return false;
	}			
}

</script>
	<!--
VB Confirm SIM/N�O
-->
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
<script language="javascript">
function YNconfirm(txt){
return typeof(suporteVBscript)=="undefined"?confirm(txt):VBconfirm(txt)==6
}

function qtdRetiraFocus(CAMPO,TIPO){
	
	if (TIPO == 1){
		//Qtd Retirada
		alert('Aten��o!\nA quantidade retirada � igual a \"Qtd. Inicial - (Qtd. Brinde + Qtd.Faturada)\"\nn�o � poss�vel alterar manualmente este valor!')
	}
	else if (TIPO == 2){
		//Qtd Brinde
		alert('Aten��o!\nA quantidade inicial � pr�viamente determinada, caso o valor estiver incorreto entre em contato com o Depto. Fotogr�fico.')
	}
	else if (TIPO == 3){
		//Qtd Brinde
		alert('Aten��o!\nA quantidade do brinde � pr�viamente determinada, caso o valor estiver incorreto entre em contato com o Depto. Fotogr�fico.')
	}
	
CAMPO.blur();

}

function qtdRetirada(qtd_inicial, qtd_brinde, qtd_faturada, COD_FORMANDO){
	d = document.produtos;
	qtd_brinde +=qtd_faturada/1;
	var valorRetirada = (qtd_inicial/1) - (qtd_brinde/1)

	if (valorRetirada < 0) {
		alert('Aten��o n�o lance c�pias aqui, somente o que foi vendido do �lbum.');
		d["QTD_FATURADA_"+COD_FORMANDO].value = "";
		d["QTD_FATURADA_"+COD_FORMANDO].focus();
	}
	else{
		d["QTD_RETIRADA_"+COD_FORMANDO].value = valorRetirada;
	}
	
	
}
</script>



<script language="vbscript">
suporteVBscript=1
function VBconfirm(mensagem)

 VBconfirm=msgbox(mensagem,vbyesno)
end function

</script>
<body>
<form name="produtos" action="lca_pedidos_pendentes_alteracoes_part_2_POPUP_album_processar.asp?cod_pedido=<%=COD_PEDIDO%>&cod_for=<%=Request.QueryString("cod_for")%>" target="_parent" onsubmit="return confirma()" method="post">
<br />
  <%
  '1) Seleciona os c�digos dos formandos do pedido
  '2) Seleciona os dados inseridos nas fichas
  AbrirConexao()
  
  Dim sql, rs
  
  sql = "SELECT TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO, TB_FICHAS_IDENTIFICACAO.ALBUM, TB_FICHAS_IDENTIFICACAO.INICIAL, TB_FICHAS_IDENTIFICACAO.BRINDE, "&_
		"TB_FICHAS_IDENTIFICACAO.PRODUTO_AGREGADO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, TB_FICHAS_IDENTIFICACAO.COMISSAO, "&_
		"TB_PEDIDOS_FORMANDOS.COD_FORMANDO "&_
		"FROM TB_FICHAS_IDENTIFICACAO INNER JOIN TB_PEDIDOS_FORMANDOS ON "&_
		"TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = TB_PEDIDOS_FORMANDOS.COD_FORMANDO "&_
		"WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&COD_PEDIDO&" AND TB_PEDIDOS_FORMANDOS.COD_FORMANDO="&Request.QueryString("cod_for")&" LIMIT 1"
  Set rs = objConexao.Execute(sql)
	Dim QTD_BRINDE
	
										
		While not rs.Eof
  		QTD_BRINDE = 0
		
		Dim rsProducao
	
	'Consulta a produ��o desta ficha para verificar qual o c�digo padr�o do �lbum
	Set rsProducao = objConexao.Execute("SELECT COD_FOTO_PADRAO FROM TB_PRODUCAO WHERE COD_PRODUCAO="&rs("COD_PRODUCAO"))
							
							If not rsProducao.Eof Then
								COD_FOTO_PADRAO = rsProducao("COD_FOTO_PADRAO")
							Else
								COD_FOTO_PADRAO = 2								
							End If
							
										
	rsProducao.Close()
	Set rsProducao = Nothing
  %>
  
  <table width="98%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
    <td colspan="5"><div align="center" class="style3"><strong><%=rs("NOME_FORMANDO")%>
          <%If rs("COMISSAO") = 1 Then %>
          - Membro de Comiss&atilde;o de Formatura
          <%end if%>
    </strong></div></td>
    </tr>
<tr bgcolor="whitesmoke">
    <td width="21%"><div align="center"><strong>Qtd. Inicial</strong></div></td>
    <td width="21%"><div align="center"><strong>Qtd. Brinde</strong> </div></td>
    <td width="20%"><div align="center"><strong>Qtd. Faturada</strong></div></td>
    <td width="17%"><div align="center"><strong>Qtd. Retirada</strong></div></td>
    <td width="21%"><div align="center"><strong>Valor Unit&aacute;rio </strong></div></td>
</tr>
								<tr>
								<td><div align="center">
                                <input style="background:#EBEBEB; color:#000000" tabindex="98874" onfocus="qtdRetiraFocus(this,2)" name="INICIAL_<%=rs("COD_FORMANDO")%>" type="text" value="<%=rs("INICIAL")%>" size="3" maxlength="5" />
								   <input name="COD_FOTO_PADRAO_<%=rs("COD_FORMANDO")%>" type="hidden" value="<%=COD_FOTO_PADRAO%>" />
                                </div></td>
								<td><div align="center">
								
								
								
								<%If rs("BRINDE") = 1 then
									Dim rsBrindes, COD_FOTO_PADRAO
									
	
									
									Set rsBrindes = objConexao.Execute("SELECT Sum(QUANTIDADE) as QUANTIDADE FROM "&_
									"TB_FICHAS_IDENTIFICACAO_BRINDES WHERE COD_PRODUTO = '"&COD_FOTO_PADRAO&"' AND COD_FORMANDO = "&rs("COD_FORMANDO"))
									
										If not rsBrindes.Eof Then
											QTD_BRINDE = rsBrindes("QUANTIDADE")%>
											
											 <input style="background:#EBEBEB; color:#000000" tabindex="98871" onfocus="qtdRetiraFocus(this,3)" name="QUANTIDADE_BRINDE_<%=rs("COD_FORMANDO")%>" type="text" value="<%=QTD_BRINDE%>" size="3" maxlength="5" />
										<%Else
											'Seguran�a para quando n�o encontrar a foto do �lbum que o formando teve de brinde
											'As vezes este erro pode ter sido ocasionado por um cadastro de outra foto que n�o � padr�o
											QTD_BRINDE = 0
											Response.Write("N&atilde;o foi poss&iacute;vel encontrar o brinde deste formando, por gentileza entre em contato com o Depto. T.I. ")	%>	
												 <input style="background:#EBEBEB; color:#000000" name="QUANTIDADE_BRINDE_<%=rs("COD_FORMANDO")%>" type="hidden" value="<%=QTD_BRINDE%>" size="3" maxlength="5" />									
	<%									End If
									
									rsBrindes.Close()
									Set rsBrindes = Nothing
								%>
									
								     <%else
								 	QTD_BRINDE = 0
								 %>
								  N&atilde;o
								   <input style="background:#EBEBEB; color:#000000" name="QUANTIDADE_BRINDE_<%=rs("COD_FORMANDO")%>" type="hidden" value="<%=QTD_BRINDE%>" size="3" maxlength="5" />		
								 <%End If%>
								</div></td>
								<td><div align="center">
								  <input name="QTD_FATURADA_<%=rs("COD_FORMANDO")%>" onblur="qtdRetirada(<%=rs("INICIAL")%>,<%=QTD_BRINDE%>,this.value,<%=rs("COD_FORMANDO")%>)" type="text" size="3" maxlength="3" />
								  </div></td>
									<td><div align="center">
								
						<input tabindex="99999" onfocus="qtdRetiraFocus(this,1)" style="background:#EBEBEB; color:#000000" name="QTD_RETIRADA_<%=rs("COD_FORMANDO")%>" type="text" size="3" maxlength="3" />
									  </div></td>
									<td><div align="center">
									  <input onKeydown="Formata(this,20,event,2)" name="VALOR_UNITARIO_<%=rs("COD_FORMANDO")%>" type="text" size="8" maxlength="15" />
									  </div></td>
		  </tr>
</table>
<br />&nbsp;
  <%
  		rs.MoveNext()
		Wend
		
		rs.Close()
		Set rs = Nothing
  FecharConexao()
  %>
  <input name="COD_PEDIDO" type="hidden" value="<%=Request.QueryString("COD_PEDIDO")%>" />
  <p align="center" style="margin:0px"><input name="btnAvancar" class="btnGravar" type="submit" id="btnAvancar" value=" " /></p>
</form>
</body>
</html>
