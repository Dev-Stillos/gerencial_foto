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
</style>
 <script language="javascript">
  
  function validaFormAparte(){
  d = document.prod_parte
	 if (d.FORM_COD_FORMANDO.value == ""){
		alert('Selecione o formando correspondente ao produto vendido!')
	 	d.FORM_COD_FORMANDO.focus();
		return false;
	 }
	 if (d.FORM_COD_PROD.value == ""){
		alert('Selecione o produto do pedido!')
	 	d.FORM_COD_PROD.focus();
		return false;
	 }
	 //alert(qtd_vendida);
	 if (d.FORM_QTD_VENDIDA.value == "" || d.FORM_QTD_VENDIDA.value == 0 || d.FORM_QTD_VENDIDA.value < 0){
		alert('Digite a quantidade vendida!')
	 	d.FORM_QTD_VENDIDA.focus();
		return false;
	 }
	  if (d.FORM_VALOR_UN.value == "" || d.FORM_VALOR_UN.value == 0 || d.FORM_VALOR_UN.value < 0){
		alert('Digite o valor do produto!')
	 	d.FORM_VALOR_UN.focus();
		return false;
	 }  
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











</head>
<body>
  <form method="post" name="prod_parte" target="_parent" action="lca_pedidos_pendentes_alteracoes_part_2.asp?cod_pedido=<%=COD_PEDIDO%>" onsubmit="return validaFormAparte()">					
  <%AbrirConexao()%>
  <%Dim rss1, rss2, qtdFormandos%>
	  <table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
      <td colspan="4"><div align="center"><strong>Lan&ccedil;amentos de Produtos &agrave; parte </strong></div></td>
      </tr>
    <tr>
      <td width="22%"><div align="center"><strong>Formando</strong></div></td>
      <td width="23%"><div align="center"><strong>Produto</strong></div></td>
      <td width="27%"><div align="center"><strong>Quantidade</strong></div></td>
      <td width="28%"><div align="center"><strong>Valor Unit&aacute;rio </strong></div></td>
    </tr>
    <tr>
      <td><select name="FORM_COD_FORMANDO">
	  <%
	  Set rss1 = objConexao.Execute("SELECT TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, TB_PEDIDOS_FORMANDOS.COD_FORMANDO "&_
									"FROM TB_FICHAS_IDENTIFICACAO "&_
									"INNER JOIN TB_PEDIDOS_FORMANDOS ON TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = TB_PEDIDOS_FORMANDOS.COD_FORMANDO "&_
									"WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO="&COD_PEDIDO&" ORDER BY TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO",qtdFormandos)
	  
	  If qtdFormandos > 1 then
	  %>
     <option value="">Selecione o formando</option>
      	<%While not rss1.Eof%>
			<option value="<%=rss1("COD_FORMANDO")%>"><%=Mid(rss1("NOME_FORMANDO"),1,20)%></option>
		<%
		rss1.MoveNext()
		Wend
		%>
	  <%else%>
	 		<option value="<%=rss1("COD_FORMANDO")%>"><%=Mid(rss1("NOME_FORMANDO"),1,20)%></option>
	 <%End If%>
	  </select>
	  <%
	  rss1.Close()
	  Set rss1 = Nothing
	  %>      </td>
      <td>
	    <select name="FORM_COD_PROD" style="width:240px;">
        <option value="">Selecione o produto</option>
		<%
		Set rss2 = objConexao.Execute("SELECT * FROM TB_PRODUTOS WHERE ATIVO = 1 ORDER BY COD_CATEGORIA, NOME_PRODUTO")
			While not rss2.Eof
		%>
			<option value="<%=rss2("COD_PRODUTO")%>"><%=rss2("NOME_PRODUTO")%></option>
		<%
			rss2.MoveNext()
			Wend
		%>
		</select>
		<%
		rss2.Close()
		Set rss2 = Nothing
		%>        </td>
      <td><div align="center">
        <input onkeypress="mascara(this,soNumeros)" name="FORM_QTD_VENDIDA" type="text" size="4" maxlength="4" />
      </div></td>
      <td><div align="center">
        <input name="FORM_VALOR_UN" type="text" size="6" value="" onKeydown="Formata(this,20,event,2)" maxlength="15" />
      </div></td>
    </tr>
    <tr>
      <td colspan="4"><div align="center">
        <input type="submit" name="GRAVAR_PRODUTO"  class="btnGravar" value=" " />
      </div></td>
    </tr>
  </table>
  <%FecharConexao()%>
  </form>
</body>
</html>
