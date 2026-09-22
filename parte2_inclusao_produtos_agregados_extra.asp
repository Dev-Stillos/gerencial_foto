<!-- #include virtual="gerencial_foto/includes/xqrtopqqetyynmkgfdconexao.asp" -->
<%
Dim COD_VENDEDOR, COD_PEDIDO, CHAVE_AVANCAR, PEDIDO_EXTRA
'Chave AVANCAR nada mais é que o controle se há produtos agregados pendentes, se houver não permite que o usuário avance.
'Inicia válido
CHAVE_AVANCAR = 1
COD_VENDEDOR = Session("COD_VENDEDOR")
COD_PEDIDO = criptografar(Request.QueryString("cod_pedido"))

If Session("COD_VENDEDOR") = Empty Then
Session.Abandon()
Response.Redirect("index.asp")
End if

'Verifica se o pedido não está finalizado
AbrirConexao()

Dim rsPedido

	Set rsPedido = objConexao.Execute("SELECT FINALIZA_VENDEDOR, PEDIDO_EXTRA FROM TB_PEDIDOS WHERE COD_PEDIDO="&COD_PEDIDO)
	
		If not rsPedido.Eof then
			If rsPedido("FINALIZA_VENDEDOR") = 1 Then
				Session.Abandon()
				Response.Redirect("index.asp")
			End If	
			
			If int(rsPedido("PEDIDO_EXTRA")) = 2 Then 
				Session.Abandon()
				Response.Redirect("index.asp?msg=Pedido Extra/normal incorreto!")
			End If		
		Else
			Session.Abandon()
			Response.Redirect("index.asp")
			
		End If
	
	rsPedido.Close()
	Set rsPedido = Nothing	

FecharConexao()
'--------------------------------------

'///////////////////////////////////////////
'Grava lançamentos à parte
If Request.Form("FORM_COD_FORMANDO") <> Empty Then
	'Gravar produtos à parte
	Dim FORM_COD_FORMANDO, FORM_COD_PROD, FORM_QTD_VENDIDA, FORM_VALOR_UN
	FORM_COD_FORMANDO	= Request.Form("FORM_COD_FORMANDO")
	FORM_COD_PROD		= Request.Form("FORM_COD_PROD") 
	FORM_QTD_VENDIDA	= Request.Form("FORM_QTD_VENDIDA") 
	FORM_VALOR_UN		= Request.Form("FORM_VALOR_UN")
	
	AbrirConexao()
	objConexao.Execute( "INSERT INTO TB_PEDIDOS_PRODUTOS ( "&_
						"COD_PEDIDO , "&_
						"COD_FORMANDO , "&_
						"COD_PRODUTO , "&_
						"VALOR_UN , "&_
						"QUANTIDADE "&_
						") "&_
						"VALUES ("&_
						COD_PEDIDO&","&FORM_COD_FORMANDO&", "&FORM_COD_PROD&", '"&sqlNumero(FORM_VALOR_UN)&"', "&FORM_QTD_VENDIDA&_
						")")
	
	FecharConexao()
	Response.Redirect("parte2_inclusao_produtos_agregados_extra.asp?cod_pedido="&criptografar(COD_PEDIDO)&"&msg=Produto Lançado com sucesso!")
End If

'/////////////////////////////////////////////////////
'Excluir produto extra
If Request.QueryString("prod_excluir_a_parte") <> Empty then
	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO_PRODUTO = "&Request.QueryString("prod_excluir_a_parte")&" LIMIT 1")
	FecharConexao()
	Response.Redirect("parte2_inclusao_produtos_agregados_extra.asp?cod_pedido="&criptografar(COD_PEDIDO)&"&msg=Produto excluído com sucesso!")
End If
'//////////////////////////////////////////////////////////////////////////
'Produto agregado, salvo individualmente
If Request.QueryString("vendido") <> Empty Then
'Grava o produto agregado
Dim valor_un_a, vendido_a
'Caso for vendido igual a 2 então faz o update com 0 no valor unitário

vendido_a	= Request.QueryString("vendido")

	If vendido_a = 1 Then
		valor_un_a  = sqlNumero(Request.QueryString("valor_un"))
	Else
		valor_un_a  = 0
	End If



	AbrirConexao()
		objConexao.Execute( " UPDATE TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS SET "&_
							" VALOR_UN = '"&valor_un_a&"', "&_
							" VENDIDO  = "&vendido_a&" WHERE COD_PRODUTO_AGREGADO = "&Request.QueryString("id")&" LIMIT 1")
	FecharConexao()
	Response.Redirect("parte2_inclusao_produtos_agregados_extra.asp?cod_pedido="&criptografar(COD_PEDIDO)&"&msg=Produto agregado salvo com sucesso!")
End If


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/principal.css" media="screen"></link>
<title>Loja Virtual - Stillo's Produções e Eventos</title><noscript>
  <meta http-equiv="Refresh" content="1;url=http://stillos.com.br/gerencial_foto/noscript.asp">
</noscript>
<style type="text/css">
<!--
.style3 {
	font-size: 13px;
	color: #003366;
	font-weight: bold;
}
-->
</style>
</head>
<script>
function abre(src, url)
{    
				//if(event.srcElement.tagName=='TD')
				//	{
				//		src.children.tags('A')[0].click(); 
				//	}
			location.href = url;

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

function confirma() {		
	d = document.produtos
	<%
	AbrirConexao()
	Dim rsJav
	
	Set rsJav = objConexao.Execute("SELECT COD_FORMANDO FROM TB_PEDIDOS_FORMANDOS WHERE COD_PEDIDO="&COD_PEDIDO)
	
	While not rsJav.Eof%>
	
	if (d.QTD_FATURADA_<%=rsJav("COD_FORMANDO")%>.value == ""){
		alert("Digite a quantidade faturada!\n* Caso não tiver digite 0 (zero).");
		d.QTD_FATURADA_<%=rsJav("COD_FORMANDO")%>.focus();
		return false;
	}
	
	if (d.VALOR_UNITARIO_<%=rsJav("COD_FORMANDO")%>.value == ""){
		
		alert("Digite o valor unitário vendido!\n* Caso não tiver digite 0 (zero).");
		d.VALOR_UNITARIO_<%=rsJav("COD_FORMANDO")%>.focus();
		return false;
	} else{
		if(d.QTD_FATURADA_<%=rsJav("COD_FORMANDO")%>.value !== "" && d.VALOR_UNITARIO_<%=rsJav("COD_FORMANDO")%>.value == 0){
			
			if(d.INICIAL_<%=rsJav("COD_FORMANDO")%>.value == d.QTD_RETIRADA_<%=rsJav("COD_FORMANDO")%>.value){	
				alert("Atenção!\nNão é possível continuar, pois a quantidade inicial é igual a quantidade\nfaturada, desta forma o álbum inteiro seria descartado, corriga antes de continuar.");
				return false;
			}
				
		}
		else if (d.QTD_FATURADA_<%=rsJav("COD_FORMANDO")%>.value/1 > 0 && d.VALOR_UNITARIO_<%=rsJav("COD_FORMANDO")%>.value == 0 ){
				alert("Atenção!\nO campo valor unitário só aceita zero quando a Qtd. Faturada é igual a 0 (zero)."); 
				d.QTD_FATURADA_<%=rsJav("COD_FORMANDO")%>.focus();
				return false;
				
		}	
	}
	
	if (d.QTD_FATURADA_<%=rsJav("COD_FORMANDO")%>.value/1 == 0 && d.VALOR_UNITARIO_<%=rsJav("COD_FORMANDO")%>.value > 0 ){
		alert("Atenção!\nVocê informou que não há fotos faturadas, mas informou o valor da foto, por gentileza corriga!"); 
		d.VALOR_UNITARIO_<%=rsJav("COD_FORMANDO")%>.focus();
		return false;
				
	}
	
	<%rsJav.MoveNext()
	Wend
	
	rsJav.Close()
	Set rsJav = Nothing
	FecharConexao()
	%>
	
	if(YNconfirm("Você não poderá voltar esta etapa, tem certeza que deseja avançar?")){
			document.produtos.submit();
			return true;
	}else{
			return false;
	}
}

</script>
	<!--
VB Confirm SIM/NÃO
-->

<script language="javascript">
function YNconfirm(txt){
return typeof(suporteVBscript)=="undefined"?confirm(txt):VBconfirm(txt)==6
}

function qtdRetiraFocus(CAMPO,TIPO){
	
	if (TIPO == 1){
		//Qtd Retirada
		alert('Atenção!\nA quantidade retirada é igual a \"Qtd. Inicial - (Qtd. Brinde + Qtd.Faturada)\"\nnão é possível alterar manualmente este valor!')
	}
	else if (TIPO == 2){
		//Qtd Brinde
		alert('Atenção!\nA quantidade inicial é préviamente determinada, caso o valor estiver incorreto entre em contato com o Depto. Fotográfico.')
	}
	else if (TIPO == 3){
		//Qtd Brinde
		alert('Atenção!\nA quantidade do brinde é préviamente determinada, caso o valor estiver incorreto entre em contato com o Depto. Fotográfico.')
	}
	
CAMPO.blur();

}

function qtdRetirada(qtd_inicial, qtd_brinde, qtd_faturada, COD_FORMANDO){
	d = document.produtos;
	qtd_brinde +=qtd_faturada/1;
	var valorRetirada = (qtd_inicial/1) - (qtd_brinde/1)

	if (valorRetirada < 0) {
		alert('Atenção não lance cópias aqui, somente o que foi vendido do álbum.');
		d["QTD_FATURADA_"+COD_FORMANDO].value = "";
		d["QTD_FATURADA_"+COD_FORMANDO].focus();
	}
	else{
		d["QTD_RETIRADA_"+COD_FORMANDO].value = valorRetirada;
	}
	
	
}

 function confirmaAgregados(COD_FORMANDO,I,COD_AGREGADO){
  		d = document["AGREGADOS_"+COD_FORMANDO]
 		if(d["AGREGADOS_VENDIDO_"+I+"_"+COD_FORMANDO].value == ""){
			alert('É necessário selecionar se foi vendido o produto agregado!');
		   d["AGREGADOS_VENDIDO_"+I+"_"+COD_FORMANDO].focus();
		   return false;
		}
		if( d["AGREGADOS_VENDIDO_"+I+"_"+COD_FORMANDO].value == 1  && (d["AGREGADOS_VALOR_UN_"+I+"_"+COD_FORMANDO].value == "" || d["AGREGADOS_VALOR_UN_"+I+"_"+COD_FORMANDO].value == 0)){
			alert('É necessário informar o valor vendido do produto agregado!');
		   d["AGREGADOS_VALOR_UN_"+I+"_"+COD_FORMANDO].focus();
			return false;
		}
		
		if(YNconfirm("Tem certeza que deseja salvar!")){
			var valor_un = d["AGREGADOS_VALOR_UN_"+I+"_"+COD_FORMANDO].value
			var vendido  = d["AGREGADOS_VENDIDO_"+I+"_"+COD_FORMANDO].value
			
			location.href = "?cod_pedido=<%=criptografar(COD_PEDIDO)%>&valor_un="+valor_un+"&vendido="+vendido+"&id="+COD_AGREGADO
			return true;
		}
 		
 }
  

</script>



<script language="vbscript">
suporteVBscript=1
function VBconfirm(mensagem)

 VBconfirm=msgbox(mensagem,vbyesno)
end function

</script>

<script>
function confirmaProdAparte(PRODUTO_A_PARTE){

	if(YNconfirm("Tem certeza que deseja excluir este produto?")){
			location.href = "?cod_pedido=<%=criptografar(COD_PEDIDO)%>&prod_excluir_a_parte="+PRODUTO_A_PARTE
	}

}

</script>


<body>
<table id="tb_p" align="center" cellpadding="0" cellspacing="0">
<tr>
<td><img src="imagens/cabecalho.jpg" width="688" height="109" /></td>
</tr>
<tr>
<td>
  <div align="center">

    <p><img style="pa" src="imagens/2_passo.jpg" width="688" height="80" /></p>
  </div></td>
</tr>
<tr>
<td align="center">
<div>
  <p align="center" class="style2" style="padding-top:5px;">PEDIDO EXTRA<br />
    Lan&ccedil;amento de Produtos Vendidos</p>
  <br />
  <span class="red"><strong><%=Request.QueryString("msg")%></strong></span>
  <br />
  <form method="post" name="prod_parte" action="parte2_inclusao_produtos_agregados_extra.asp?cod_pedido=<%=criptografar(COD_PEDIDO)%>" onsubmit="return validaFormAparte()">
  <%AbrirConexao()%>
  <%Dim rss1, rss2, qtdFormandos%>
	  <table width="98%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
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
			<option value="<%=rss1("COD_FORMANDO")%>"><%=rss1("NOME_FORMANDO")%></option>
		<%
		rss1.MoveNext()
		Wend
		%>
	  <%else%>
	 	<option value="<%=rss1("COD_FORMANDO")%>"><%=rss1("NOME_FORMANDO")%></option>
	 <%End If%>
	  </select>
	  <%
	  rss1.Close()
	  Set rss1 = Nothing
	  %>      </td>
      <td>
	    <select name="FORM_COD_PROD">
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
        <input name="FORM_VALOR_UN" type="text" size="6" maxlength="7" />
      </div></td>
    </tr>
    <tr>
      <td colspan="4"><div align="center">
        <input type="submit" name="GRAVAR_PRODUTO" value="GRAVAR PRODUTO" />
        <input type="hidden" name="COD_PEDIDO" value="<%=criptografar(COD_PEDIDO)%>" />
      </div></td>
      </tr>
  </table>
      <%FecharConexao()%>
  </form>
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
  <p align="center">  
   <%
  '1) Seleciona os códigos dos formandos do pedido
  '2) Seleciona os dados inseridos nas fichas
  AbrirConexao()
  
  Dim sql, rs
  
  sql = "SELECT TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO, TB_FICHAS_IDENTIFICACAO.CREDITO_EM_VALOR, TB_FICHAS_IDENTIFICACAO.CREDITO_MOTIVO, TB_FICHAS_IDENTIFICACAO.CREDITO_DATA_INSERIDO, TB_FICHAS_IDENTIFICACAO.FATURADO, TB_FICHAS_IDENTIFICACAO.VALOR_UN, TB_FICHAS_IDENTIFICACAO.ALBUM, TB_FICHAS_IDENTIFICACAO.INICIAL, TB_FICHAS_IDENTIFICACAO.BRINDE, "&_
		"TB_FICHAS_IDENTIFICACAO.PRODUTO_AGREGADO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, TB_FICHAS_IDENTIFICACAO.COMISSAO, "&_
		"TB_PEDIDOS_FORMANDOS.COD_FORMANDO "&_
		"FROM TB_FICHAS_IDENTIFICACAO INNER JOIN TB_PEDIDOS_FORMANDOS ON "&_
		"TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = TB_PEDIDOS_FORMANDOS.COD_FORMANDO "&_
		"WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&COD_PEDIDO
  		
  Set rs = objConexao.Execute(sql)
	Dim QTD_BRINDE, QTD_RETIRADA, CREDITO_EM_VALOR
	
										
		While not rs.Eof
  		QTD_BRINDE 			= 0
		QTD_RETIRADA 		= 0
		CREDITO_EM_VALOR	= formatNumber(0&rs("CREDITO_EM_VALOR"))
		
		Dim rsProducao, rsBrindesAlbum
	
	'Consulta a produção desta ficha para verificar qual o código padrão do álbum
	Set rsProducao = objConexao.Execute("SELECT COD_FOTO_PADRAO FROM TB_PRODUCAO WHERE COD_PRODUCAO="&rs("COD_PRODUCAO"))
			If not rsProducao.Eof Then
			COD_FOTO_PADRAO = formatNumber(0&rsProducao("COD_FOTO_PADRAO"),0)
			Else
			COD_FOTO_PADRAO = 2			
			End If
	rsProducao.Close()
	Set rsProducao = Nothing
	
	'Consulta a quantidade de brinde, agora com a foto padrão do álbum
	if COD_FOTO_PADRAO <> 0 Then
	
	Set rsBrindesAlbum = objConexao.Execute("SELECT Sum(QUANTIDADE) as QUANTIDADE FROM TB_FICHAS_IDENTIFICACAO_BRINDES WHERE COD_FORMANDO="&rs("COD_FORMANDO")&" AND COD_PRODUTO="&COD_FOTO_PADRAO&" GROUP BY COD_PRODUTO")
	
		If not rsBrindesAlbum.Eof then
			QTD_BRINDE = formatNumber(0&rsBrindesAlbum("QUANTIDADE"),0)
		Else
			QTD_BRINDE = 0
		End If
		
	
	rsBrindesAlbum.Close()
	Set rsBrindesAlbum = Nothing
	
	Else
		
		QTD_BRINDE = 0
	
	End If
  %>
  <table width="98%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
    <td width="100%">
	  <div align="left" class="style3" style="border-bottom:1px #999999 solid"><%=rs("NOME_FORMANDO")%>
	    <%If rs("COMISSAO") = 1 Then %>
	    - Membro de Comissão de Formatura
	    <%end if%>
	    </div>
	  <div align="center" class="style3">
	    <div align="left">Cód. Form: <%=digito(rs("COD_FORMANDO"))%> - Álbum: <%=rs("ALBUM")%></div>
	  </div></td>
    </tr>
<%
Dim rsProdutosAparte, QTD_PROD_A_PARTE
QTD_PROD_A_PARTE = 0
Set rsProdutosAparte = objConexao.Execute(" SELECT TB_PEDIDOS_PRODUTOS.COD_PEDIDO_PRODUTO, TB_PRODUTOS.NOME_PRODUTO, TB_PEDIDOS_PRODUTOS.VALOR_UN, "&_
										  " TB_PEDIDOS_PRODUTOS.QUANTIDADE FROM TB_PRODUTOS INNER JOIN "&_
										  " TB_PEDIDOS_PRODUTOS ON TB_PRODUTOS.COD_PRODUTO = TB_PEDIDOS_PRODUTOS.COD_PRODUTO"&_
										  " WHERE COD_PEDIDO="&COD_PEDIDO&" AND COD_FORMANDO="&rs("COD_FORMANDO")&" ORDER BY NOME_PRODUTO")
If not rsProdutosAparte.Eof Then
QTD_PROD_A_PARTE = 1
%>
<tr bgcolor="whitesmoke">
  <td><div align="center"><strong>Pedido Extra</strong></div></td>
</tr>
<tr>
  <td><table width="98%" border="1" align="center" 	cellpadding="2" cellspacing="2" style="border-collapse:collapse">
    <tr>
      <td width="35%"><div align="center"><strong>Produto</strong></div></td>
      <td width="21%"><div align="center"><strong>Qtd.</strong></div></td>
      <td width="17%"><div align="center"><strong>Valor Unit&aacute;rio </strong></div></td>
      <td width="16%"><div align="center"><strong>Valor Total</strong></div></td>
      <td width="11%"><div align="center"><strong>Excluir</strong></div></td>
    </tr>
    <%While not rsProdutosAparte.Eof%>
	<tr>
      <td><div align="left"><%=rsProdutosAparte("NOME_PRODUTO")%></div></td>
      <td><div align="center"><%=rsProdutosAparte("QUANTIDADE")%></div></td>
      <td><div align="center"><%=formatCurrency(rsProdutosAparte("VALOR_UN"))%></div></td>
      <td><div align="center">
	  <%
	  Dim tot1, tot2
	 tot1 = formatNumber(0&rsProdutosAparte("QUANTIDADE"),0)
     tot2 = 0&rsProdutosAparte("VALOR_UN")
	  %>
	  <%=formatCurrency(tot1*tot2)%></div></td>
      <td><div align="center"><a href="javascript:Void(0)" onclick="confirmaProdAparte(<%=rsProdutosAparte("COD_PEDIDO_PRODUTO")%>)" tabindex="98777"><img src="../gerencial_foto/images/icons/icon_excluir.png" width="16" height="16" border="0" /></a></div></td>
    </tr>
	<%
	rsProdutosAparte.MoveNext()
	Wend
	%>
  </table></td>
</tr>
<%Else
QTD_PROD_A_PARTE = 0	

End If
rsProdutosAparte.Close()
Set rsProdutosAparte = Nothing
%>
</table>

  <div align="left">
    <p><br />
      &nbsp;
      <%
  		rs.MoveNext()
		Wend
		
		rs.Close()
		Set rs = Nothing
  FecharConexao()
  %>
      <span class="style2">Total pago em produtos: <%
						Dim rsProdutosTotal, XTOTAL_PRODUTOS_A_PARTE
						AbrirConexao()
						
						QTD_EM_PRODUTOS = 0
						
							Set rsProdutosTotal = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_PRODUTOS WHERE COD_PEDIDO="&COD_PEDIDO)
							
								if not rsProdutosTotal.Eof Then
									
									XTOTAL_PRODUTOS_A_PARTE = formatNumber(0&rsProdutosTotal("XTOTAL_PRODUTOS_A_PARTE"))/1
									
									QTD_EM_PRODUTOS = formatCurrency(XTOTAL_PRODUTOS_A_PARTE)
									
								Else
									
									QTD_EM_PRODUTOS = 0
									
								End If
							
							
							rsProdutosTotal.Close()
							Set rsProdutosTotal = Nothing
							
						FecharConexao()
						%>
						<%=QTD_EM_PRODUTOS%></span> </p>
	</div>
  <p align="center" style="margin:0px">
  <%If QTD_PROD_A_PARTE = 1 Then%>
  <input name="btnAvancar" type="button" id="btnAvancar" onclick="abre('','pag3_forma_pagamento.asp?cod_pedido=<%=criptografar(COD_PEDIDO)%>')" value="Avan&ccedil;ar" />
  <%Else%>
  <input name="btnAvancar" type="submit" onclick="alert('Não é possível avançar pois não há produtos lançados')" id="btnAvancar" value="Avan&ccedil;ar" />
  <%End If%>
  <br />
  <br />
  </p>
  </p>
</div></td>
</tr>
</table>
<!-- #include file="rodape.asp" -->
</body>
</html>