<!-- #include virtual="gerencial_foto/includes/xqrtopqqetyynmkgfdconexao.asp" -->
<%
sub registroArquivo(COD_PRODUTO, COD_PEDIDO, COD_PEDIDO_PRODUTO)

	Dim rs, quantidade, necessita_fotos, NOME_PRODUTO
	'A conexao deve obrigatoriamente aberta
	'Verificamos se o produto precisa de arquivos
	'Response.write "SELECT necessita_fotos, NOME_PRODUTO FROM TB_PRODUTOS WHERE COD_PRODUTO = "&sqlNumero(COD_PRODUTO)
	Set rs = objConexao.Execute("SELECT necessita_fotos, NOME_PRODUTO FROM TB_PRODUTOS WHERE COD_PRODUTO = "&sqlNumero(COD_PRODUTO) )
	if not rs.Eof then
		necessita_fotos = int(0&rs(0))
		NOME_PRODUTO = rs("NOME_PRODUTO")
	end if
	rs.close
	set rs = Nothing

	if necessita_fotos = 1 then

		'contador de ordem dos arquivos
		Dim count_prod 

		'Verificamos a quantidade de produtos para definirmos a quantidade de itens
		set rs = objConexao.Execute("SELECT QUANTIDADE FROM TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO = "&sqlNumero(cod_pedido) &" AND cod_pedido_produto = "&sqlNumero(cod_pedido_produto))
		if not rs.Eof then
			for count_prod = 1 to rs("quantidade")

				if COD_PRODUTO = 7 then
					'inserimos os produtos na tabela de itens
					objConexao.Execute("INSERT INTO TB_PEDIDOS_PRODUTOS_ARQUIVOS (descricao,COD_PEDIDO, COD_PEDIDO_PRODUTO, ordem, nome_arquivo) VALUES ('Tela 20x30',"&sqlNumero(cod_pedido)&", "&sqlNumero(cod_pedido_produto)&", "&sqlNumero(count_prod)&",'') ")

					'inserimos os produtos na tabela de itens
					objConexao.Execute("INSERT INTO TB_PEDIDOS_PRODUTOS_ARQUIVOS (descricao,COD_PEDIDO, COD_PEDIDO_PRODUTO, ordem, nome_arquivo) VALUES ('prt 10x15',"&sqlNumero(cod_pedido)&", "&sqlNumero(cod_pedido_produto)&", "&sqlNumero(count_prod)&",'') ")

					'inserimos os produtos na tabela de itens
					objConexao.Execute("INSERT INTO TB_PEDIDOS_PRODUTOS_ARQUIVOS (descricao,COD_PEDIDO, COD_PEDIDO_PRODUTO, ordem, nome_arquivo) VALUES ('prt 10x15',"&sqlNumero(cod_pedido)&", "&sqlNumero(cod_pedido_produto)&", "&sqlNumero(count_prod)&",'') ")
				else 
					'set rs2 = objConexao.execute("SELECT NOME_PRODUTO FROM TB_PRODUTOS WHERE COD_PRODUTO = "&sqlNumero("cod_produto"))
					'if not rs2.eof then
						'inserimos os produtos na tabela de itens
						objConexao.Execute("INSERT INTO TB_PEDIDOS_PRODUTOS_ARQUIVOS (descricao, COD_PEDIDO, COD_PEDIDO_PRODUTO, ordem, nome_arquivo) VALUES ("&sqlTexto(NOME_PRODUTO & " - "& count_prod )&", "&sqlNumero(cod_pedido)&", "&sqlNumero(cod_pedido_produto)&", "&sqlNumero(count_prod)&",'') ")
					'end if
					'rs2.close
					'set rs2 = nothing
				end if
			next
		end if 
		rs.close
		set rs = Nothing

	end if

end sub

sub apagarArquivo(cod_pedido_produto)
		
	'A CONEXAO DEVE ESTAR OBRIGATORIAMENTE ABERTA
	Dim rs, aux, fso, arqvo

	set rs = objConexao.Execute("SELECT * FROM TB_PEDIDOS_PRODUTOS_ARQUIVOS WHERE cod_pedido_produto = "&sqlNumero(cod_pedido_produto))
	while not rs.Eof
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		'a variavel caminho_dados vem do arquivo de conexao
		deletarArquivo(caminho_dados&"pedidos_produtos_anexos\"&rs("nome_arquivo")&"."&rs("extencao"))

	rs.MoveNext
	Wend
	rs.close
	set rs = Nothing

	objConexao.Execute("DELETE FROM TB_PEDIDOS_PRODUTOS_ARQUIVOS WHERE cod_pedido_produto = "&sqlNumero(cod_pedido_produto))

end sub


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
			
			If int(rsPedido("PEDIDO_EXTRA")) = 1 Then
				Session.Abandon()
				Response.Redirect("index.asp?msg=Pedido Normal/Extra incorreto!")
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

	'RESGATAMOS O CODIGO DO ULTIMO PRODUTO INCLUIDO
	set rsUltimoProdu = objConexao.Execute("SELECT max(cod_pedido_produto) as cod_pedido_produto FROM TB_PEDIDOS_PRODUTOS WHERE cod_pedido = "&sqlNumero(COD_PEDIDO))
	if not rsUltimoProdu.Eof then
		call registroArquivo(FORM_COD_PROD, COD_PEDIDO, rsUltimoProdu(0) )
	end if
	rsUltimoProdu.close
	set rsUltimoProdu = Nothing
	
	FecharConexao()
	Response.Redirect("parte2_inclusao_produtos_agregados.asp?cod_pedido="&criptografar(COD_PEDIDO)&"&msg=Produto Lançado com sucesso!")
End If

'/////////////////////////////////////////////////////
'Excluir produto extra
If Request.QueryString("prod_excluir_a_parte") <> Empty then
	AbrirConexao()

		call apagarArquivo(Request.QueryString("prod_excluir_a_parte"))

		objConexao.Execute("DELETE FROM TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO_PRODUTO = "&Request.QueryString("prod_excluir_a_parte")&" LIMIT 1")

	FecharConexao()
	Response.Redirect("parte2_inclusao_produtos_agregados.asp?cod_pedido="&criptografar(COD_PEDIDO)&"&msg=Produto excluído com sucesso!")
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
	Response.Redirect("parte2_inclusao_produtos_agregados.asp?cod_pedido="&criptografar(COD_PEDIDO)&"&msg=Produto agregado salvo com sucesso!")
End If

'INCLUIDO JOAO
'20/05/2015

if request.form("incluir_pasta") <> "" then

	Dim COD_FORMANDO, numero_pasta
	COD_FORMANDO = request.form("COD_FORMANDO")
	numero_pasta = request.form("numero_pasta")

	AbrirConexao()

		if numero_pasta <> "" and COD_FORMANDO <> "" then

			'Response.write ("UPDATE TB_FICHAS_IDENTIFICACAO SET numero_pasta_vendedor = "&numero_pasta&" WHERE COD_FORMANDO = "&COD_FORMANDO)

			objConexao.execute("UPDATE TB_FICHAS_IDENTIFICACAO SET numero_pasta_vendedor = "&numero_pasta&" WHERE COD_FORMANDO = "&COD_FORMANDO)

		end if

	FecharConexao()

	Response.Redirect("parte2_inclusao_produtos_agregados.asp?cod_pedido="&criptografar(COD_PEDIDO)&"&msg=Produto agregado salvo com sucesso!")

end if

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<script type="text/javascript" src='js/jquery-1.11.2.min.js'></script>
<script type="text/javascript" src='js/bootstrap.min.js'></script>
<script type="text/javascript" src='js/chosen.jquery.min.js'></script>
<link rel="stylesheet" type="text/css" href="css/chosen.min.css" />
<link rel="stylesheet" type="text/css" href="css/bootstrap.min.css" />
<link type="text/css" rel="stylesheet" href="css/principal.css"></link>
<title>Loja Virtual - Stillo's Produções e Eventos</title><noscript>
	<meta http-equiv="Refresh" content="1;url=http://stillos.com.br/gerencial_foto/noscript.asp">
</noscript>


<script type="text/javascript">
	$(function(){
		 $(".chosen-select").chosen({no_results_text: "Nehum produto encontrado"}); 
		});

</script>

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
	
	alert('confirma');

	if (d.QTD_FATURADA_<%=rsJav("COD_FORMANDO")%>.value == ""){
		alert("Digite a quantidade faturada!\n* Caso não tiver digite 0 (zero).");
		d.QTD_FATURADA_<%=rsJav("COD_FORMANDO")%>.focus();
		return false;
	}
	else
		{alert('campo não encontrado');}
	
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

<script type="text/javascript">
	function modalArquivo(cod_pedido, COD_PEDIDO_PRODUTO)
	{
		$("#iframeArquivo").prop("src", "/loja/upload.asp?c="+cod_pedido+"&cp="+COD_PEDIDO_PRODUTO);
		$("#modalArquivoDiv").modal('show');
	}
	$(function(){
		$("#modalArquivoDiv").on('hide.bs.modal', function(e){
	 		location.reload();
	  	});
	});
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
  <p align="center" class="style2" style="padding-top:5px;">Lan&ccedil;amento de Produtos Vendidos - Produtos Extras </p>
  <table style="border:dashed #ccc 1px; background-color:#FDF4D5" width="98%" border="0" cellspacing="0" cellpadding="5">
    <tr>
      <td><div align="left"><strong>Instru&ccedil;&otilde;es para o lan&ccedil;amento de produtos:</strong><br />
        1&ordm; Lan&ccedil;amento dos &aacute;lbuns dos formandos<br />
        2&ordm; Lan&ccedil;amento de produtos extras (C&oacute;pias, DVD, Trio, entre outros)</div></td>
    </tr>
  </table>
  <br />
  <span class="red"><strong><%=Request.QueryString("msg")%></strong></span>
  <br />
  <form method="post" name="prod_parte" onsubmit="return validaFormAparte()">
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
	  dim COD_CONTRATO, sqlll
	  sqlll = "SELECT TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, TB_FICHAS_IDENTIFICACAO.COD_CONTRATO , TB_PEDIDOS_FORMANDOS.COD_FORMANDO "&_
			"FROM TB_FICHAS_IDENTIFICACAO "&_
			"INNER JOIN TB_PEDIDOS_FORMANDOS ON TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = TB_PEDIDOS_FORMANDOS.COD_FORMANDO "&_
			"WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO="&COD_PEDIDO&" ORDER BY TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO"
	 Response.write sqlll

	  Set rss1 = objConexao.Execute(sqlll,qtdFormandos)
	  
	  If qtdFormandos > 1 then
	  %>
     <option value="">Selecione o formando</option>
      	<%While not rss1.Eof 



      	 COD_CONTRATO = COD_CONTRATO & sqlTexto(rss1("COD_CONTRATO")) &","
      	%>
			<option value="<%=rss1("COD_FORMANDO")%>"><%=rss1("NOME_FORMANDO")%><%response.write (rss1("COD_CONTRATO")&"ssss")%></option>
		<%
		rss1.MoveNext()
		Wend
		%>
	  <%else

 COD_CONTRATO = COD_CONTRATO & sqlTexto(rss1("COD_CONTRATO"))&","
	  %>
	 	<option value="<%=rss1("COD_FORMANDO")%>"><%=rss1("NOME_FORMANDO")%></option>
	 <%End If%>
	  </select>
	  <%
	  rss1.Close()
	  Set rss1 = Nothing

COD_CONTRATO = mid(COD_CONTRATO&"",1,len(COD_CONTRATO)-1)

	  %>      </td>
      <td>
	    <select name="FORM_COD_PROD" style="width:250px;" class="chosen-select">
        <option value="">Selecione o produto</option>
		<%


		'response.write "SELECT * FROM TB_PRODUTOS WHERE ATIVO = 1 AND cod_produto in (SELECT cod_produto FROM stillos3.VW_TB_BOLETIM_INFORMATIVO_PRODUTOS where COD_BOLETIM_INFORMATIVO in (select COD_BOLETIM_INFORMATIVO from VW_TB_BOLETIM_INFORMATIVO_CONTRATOS where cod_contrato in ("&COD_CONTRATO&") )) ORDER BY  NOME_PRODUTO"

		Set rss2 = objConexao.Execute("SELECT * FROM TB_PRODUTOS WHERE ATIVO = 1 AND cod_produto in (SELECT cod_produto FROM stillos3.VW_TB_BOLETIM_INFORMATIVO_PRODUTOS where COD_BOLETIM_INFORMATIVO in (select COD_BOLETIM_INFORMATIVO from VW_TB_BOLETIM_INFORMATIVO_CONTRATOS where cod_contrato in ("&COD_CONTRATO&") )) or cod_produto = 40 ORDER BY  NOME_PRODUTO")
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

	dim pode_continuar
pode_continuar = true

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
		<%If CREDITO_EM_VALOR > 0 Then%>
	<tr bgcolor="whitesmoke">
  <td><div align="center"><strong>Cr&eacute;dito da Formatura </strong></div></td>
</tr>
<tr>
  <td><table width="98%" border="1" align="center" 	cellpadding="2" cellspacing="2" style="border-collapse:collapse">
    <tr>
      <td width="21%"><div align="center"><strong>Valor Cr&eacute;dito </strong></div></td>
      <td width="53%"><div align="center"><strong>Motivo</strong></div></td>
      <td width="26%"><div align="center"><strong>Data inserido no sistema </strong></div></td>
      </tr>

    <tr>
      <td><div align="center"><%=formatCurrency(CREDITO_EM_VALOR)%></div></td>
      <td><div align="left"><%=rs("CREDITO_MOTIVO")%></div></td>
      <td><div align="center"><%=rs("CREDITO_DATA_INSERIDO")%></div></td>
      </tr>

  </table></td>
</tr>
	<%End If%>
<tr bgcolor="whitesmoke">
    <td><div align="center"><strong>Dados do &Aacute;lbum </strong></div></td>
    </tr>
<tr>
  <td><table width="98%" border="1" align="center" cellpadding="2" cellspacing="2" style="border-collapse:collapse">
    <tr>
      <td width="16%"><div align="center"><strong>Qtd. Inicial</strong></div></td>
      <td width="15%"><div align="center"><strong>Qtd. Brinde</strong></div></td>
      <td width="15%"><div align="center"><strong>Qtd. Faturada</strong></div></td>
      <td width="14%"><div align="center"><strong>Qtd. Retirada </strong></div></td>
      <td width="19%"><div align="center"><strong>Valor Un.</strong></div></td>
      <td width="21%"><div align="center"><strong>Valor Total</strong></div></td>
      </tr>
    <tr>
      <td><div align="center"><%=rs("INICIAL")%></div></td>
      <td><div align="center"><%=QTD_BRINDE%></div></td>
      <td><div align="center">
	   <%
	  Dim  rsVALOR_UN, rsFATURADO, rsVALOR_TOTAL
	  rsVALOR_UN = formatNumber(0&rs("VALOR_UN"))
	  rsFATURADO = formatNumber(0&rs("FATURADO"),0)
	  rsVALOR_TOTAL = rsVALOR_UN * rsFATURADO
	  
	  QTD_RETIRADA = (0&rs("INICIAL")) - (QTD_BRINDE/1 + rsFATURADO/1)
	  %>
	  
      <%If rsFATURADO > 0 then %>
	  	<%=rsFATURADO%>
	  <%Else%>
	  	---
	  <%End If%>
	  </div></td>
      <td><div align="center"><%=QTD_RETIRADA%></div></td>
      <td><div align="center">
	  <%If rsVALOR_UN > 0 then %>
	  	<%=formatCurrency(rsVALOR_UN)%>
	  <%Else%>
	  	---
	  <%End If%>
	  </div></td>
      <td><div align="center">

        <%If rsVALOR_TOTAL > 0 then 
			
		%>
			<%=formatCurrency(rsVALOR_TOTAL)%>
        <%Else%>
		---
		<%End If%>
</div></td>
    </tr>

  </table></td>
</tr>
<%
Dim rsBrindes

Set rsBrindes = objConexao.Execute( " SELECT Sum(TB_FICHAS_IDENTIFICACAO_BRINDES.QUANTIDADE) as QUANTIDADE, "&_
	" TB_PRODUTOS.NOME_PRODUTO, TB_PRODUTOS.necessita_fotos FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_BRINDES "&_
	" ON TB_PRODUTOS.COD_PRODUTO = TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO "&_
	" WHERE TB_FICHAS_IDENTIFICACAO_BRINDES.COD_FORMANDO="&rs("COD_FORMANDO")&" AND TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO <> "& COD_FOTO_PADRAO&_
	" GROUP BY TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO ")
If not rsBrindes.Eof Then
%>
<tr bgcolor="whitesmoke">
  <td><div align="center"><strong>Outros Brindes</strong></div></td>
</tr>
<tr>
  <td><table width="98%" border="1" align="center" 	cellpadding="2" cellspacing="2" style="border-collapse:collapse">
    <tr>
      <td width="56%"><div align="center"><strong>Produto</strong></div></td>
      <td width="44%"><div align="center"><strong>Qtd</strong></div></td>
      
      </tr>
	   <%
	   While not rsBrindes.Eof
	   %>
    	<tr>
      <td><div align="left"><%=rsBrindes("NOME_PRODUTO")%></div></td>
      <td><div align="center"><%=rsBrindes("QUANTIDADE")%></div></td>
      </tr>
	   <%
	   rsBrindes.MoveNext()
	   Wend
	   %>
  </table></td>
</tr>
<%
End If

rsBrindes.Close()
Set rsBrindes = Nothing
%>
<%
Dim rsAgregados
Set rsAgregados = objConexao.Execute(	" SELECT TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO, TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO_AGREGADO, "&_
										" Sum(TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.QUANTIDADE) as QUANTIDADE, "&_
										" TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.VENDIDO, TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.VALOR_UN, "&_
										" TB_PRODUTOS.NOME_PRODUTO FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS "&_
										" ON TB_PRODUTOS.COD_PRODUTO = TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO "&_
										" WHERE TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO="&rs("COD_FORMANDO")&_
										" GROUP BY TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO ")

If not rsAgregados.Eof Then
%>
<tr bgcolor="whitesmoke">
  <td><div align="center"><strong>Produtos Agregados com sa&iacute;da de produ&ccedil;&atilde;o</strong><br />
  </div></td>
</tr>
<tr>  <td>	
		<form style="margin:0px" method="post" name="AGREGADOS_<%=rs("COD_FORMANDO")%>">
			<table width="98%" border="1" align="center" 	cellpadding="2" cellspacing="2" style="border-collapse:collapse">
    <tr>
      <td width="35%"><div align="center"><strong>Produto</strong></div></td>
      <td width="21%"><div align="center"><strong>Qtd</strong>.</div></td>
      <td width="17%"><div align="center"><strong>Vendido</strong></div></td>
      <td width="16%"><div align="center"><strong>Valor Unit&aacute;rio </strong></div></td>
      <td width="11%"><div align="center"><strong>Altera&ccedil;&otilde;es</strong></div></td>
    </tr>
	<%
	Dim contProdAgregado, qteCod, qtaUi, FOI_BRINDE
	contProdAgregado = 1
	qteCod = formatNumber(0&Request.QueryString("prod_agre_alterar"),0)
	qtaUi  = formatNumber(0&rsAgregados("COD_PRODUTO_AGREGADO"),0)
	While not rsAgregados.Eof
	FOI_BRINDE = 2		
			
		'/////////////////
		'VERIFICA SE O PRODUTO AGREGADO NÃO POSSUI O MESMO COMO BRINDE
		
		Set rs12=objConexao.Execute("SELECT COD_BRINDE FROM TB_FICHAS_IDENTIFICACAO_BRINDES WHERE COD_FORMANDO = "&rs("COD_FORMANDO")&" AND COD_PRODUTO="&rsAgregados("COD_PRODUTO"))
		If not rs12.Eof Then
		'É BRINDE - SOMENTE ENTREGA
			FOI_BRINDE = 1
		End If
		rs12.Close
		Set rs12=Nothing
					
		'/////////////////
			
		contProdAgregado = contProdAgregado + 1
		
	
	If FOI_BRINDE <> 1 Then
	%>
    
    <tr <%If qteCod = qtaUi Then%> bgcolor="#CCCCCC" <%End If%>>
      <td><div align="left"><%=rsAgregados("NOME_PRODUTO")%></div></td>
      <td><div align="center"><%=rsAgregados("QUANTIDADE")%></div></td>
      <td><div align="center">
	  <%If rsAgregados("VENDIDO") = 3 then
	  CHAVE_AVANCAR = 2
	  %>
          <select name="AGREGADOS_VENDIDO_<%=contProdAgregado%>_<%=rs("COD_FORMANDO")%>">
            <option>---</option>
            <option value="1">Sim</option>
            <option value="2">N&atilde;o</option>
          </select>
      <%else%>
	  		<%If Request.QueryString("prod_agre_alterar") <> Empty Then
			CHAVE_AVANCAR = 2
			%>
          	<select name="AGREGADOS_VENDIDO_<%=contProdAgregado%>_<%=rs("COD_FORMANDO")%>">
				<option>---</option>
				<option <%If rsAgregados("VENDIDO") = 1 Then%> selected="selected"<%End If%> value="1">Sim</option>
				<option <%If rsAgregados("VENDIDO") = 2 Then%> selected="selected"<%End If%> value="2">N&atilde;o</option>
         	 </select>
			<%Else%>
				<%If rsAgregados("VENDIDO") = 1 Then%>
					Sim
				<%Else%>
					Não
				<%End If%>	
	  		<%End If%>
	  <%End If%>
	  </div></td>
      <td>
	  <div align="center">
	  <%If Request.QueryString("prod_agre_alterar") <> Empty Then%>
			<input value="<%=retiraZero(formatNumber(0&rsAgregados("VALOR_UN")))%>" name="AGREGADOS_VALOR_UN_<%=contProdAgregado%>_<%=rs("COD_FORMANDO")%>" type="text" size="6" maxlength="5" />	
	  <%Else%>
			 <%If rsAgregados("VENDIDO") = 3 then%>
				  <input name="AGREGADOS_VALOR_UN_<%=contProdAgregado%>_<%=rs("COD_FORMANDO")%>" type="text" size="6" maxlength="5" />
			 	 <%elseif rsAgregados("VENDIDO") = 2 then%>
				  ---
			 	<%else%>
			  	<%=formatCurrency(0&rsAgregados("VALOR_UN"))%>
			  	<%End If%>
	  <%End If
	  
	  
	  %>
      </div>
	  </td>
      <td>
	  <%If rsAgregados("VENDIDO") = 3 Then%>
	  <div align="center">
	  <input name="SALVER_PROD" type="button" onclick="confirmaAgregados(<%=rs("COD_FORMANDO")%>,<%=contProdAgregado%>,<%=rsAgregados("COD_PRODUTO_AGREGADO")%>)" value="Salvar" />
	  </div>
	 
	    <%Else%>
		 <div align="center">
	    <%If Request.QueryString("prod_agre_alterar") <> Empty Then%>
	    <input name="SALVER_PROD" type="button" onclick="confirmaAgregados(<%=rs("COD_FORMANDO")%>,<%=contProdAgregado%>,<%=rsAgregados("COD_PRODUTO_AGREGADO")%>)" value="Salvar" />
	        <%Else%>
	          <a href="?cod_pedido=<%=criptografar(COD_PEDIDO)%>&prod_agre_alterar=<%=rsAgregados("COD_PRODUTO_AGREGADO")%>">
	               <img src="../gerencial_foto/images/icons/icon_alterar.png" width="16" height="16" border="0" />
	                  </a>
	        <%End If%>
	    
	              <%End If%>	  
	    </div></td>
    </tr>
    
    <%
	Else
	'MOSTRA QUE FOI SOMENTE BRINDE
	%>
    <tr>
      <td><div align="left"><%=rsAgregados("NOME_PRODUTO")%></div></td>
      <td><div align="center"><%=rsAgregados("QUANTIDADE")%></div></td>
      <td><div align="center">N&atilde;o - produto entregue como brinde</div></td>
      <td><div align="center">R$ 0,00</div></td>
      <td><div align="center">-</div></td>
    </tr>
    
	<%
	End If
	
	rsAgregados.MoveNext()
	Wend
	%>
  </table>
  		</form>
  </td>
</tr>	
<%End If

rsAgregados.Close()
Set rsAgregados = Nothing

%>

<%
Dim rsProdutosAparte
Set rsProdutosAparte = objConexao.Execute(" SELECT TB_PEDIDOS_PRODUTOS.COD_PEDIDO_PRODUTO, TB_PRODUTOS.NOME_PRODUTO,TB_PRODUTOS.COD_PRODUTO, TB_PEDIDOS_PRODUTOS.VALOR_UN, TB_PRODUTOS.necessita_fotos, "&_
										  " TB_PEDIDOS_PRODUTOS.QUANTIDADE FROM TB_PRODUTOS INNER JOIN "&_
										  " TB_PEDIDOS_PRODUTOS ON TB_PRODUTOS.COD_PRODUTO = TB_PEDIDOS_PRODUTOS.COD_PRODUTO"&_
										  " WHERE COD_PEDIDO="&COD_PEDIDO&" AND COD_FORMANDO="&rs("COD_FORMANDO")&" ORDER BY NOME_PRODUTO")
If not rsProdutosAparte.Eof Then
%>
<tr bgcolor="whitesmoke">
  <td><div align="center"><strong>Produtos Vendidos &agrave; parte</strong></div></td>
</tr>
<tr>
  <td><table width="98%" border="1" align="center" 	cellpadding="2" cellspacing="2" style="border-collapse:collapse">
    <tr>
      <td width="34%"><div align="center"><strong>Produto</strong></div></td>
      <td width="21%"><div align="center"><strong>Qtd.</strong></div></td>
      <td width="17%"><div align="center"><strong>Valor Unit&aacute;rio </strong></div></td>
      <td width="16%"><div align="center"><strong>Valor Total</strong></div></td>
      <td width="7%"><div align="center"><strong>Excluir</strong></div></td>
      <td width="4%"><div align="center"><strong>Pendência</strong></div></td>
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
       <td>
      	<div align="center">
			<%
			if rsProdutosAparte("necessita_fotos") = 1 then 

				dim quantidade_enviado
		
				Set rsFoto = objConexao.execute("SELECT count(*) FROM TB_PEDIDOS_PRODUTOS_ARQUIVOS WHERE nome_arquivo <> '' AND cod_pedido_produto = "&sqlNumero(rsProdutosAparte("COD_PEDIDO_PRODUTO")))
				if not rsFoto.Eof then
					quantidade_enviado = int(0&rsFoto(0))
				end if
				rsFoto.close
				set rsFoto = Nothing

				if quantidade_enviado = rsProdutosAparte("QUANTIDADE") then
				%>
					<img src="imagens/upload_ok.png" style="cursor:pointer;" width="16" height="15" onclick="modalArquivo(<%=COD_PEDIDO%>,<%=rsProdutosAparte("COD_PEDIDO_PRODUTO")%>)">
				<% 
				elseif (quantidade_enviado < rsProdutosAparte("QUANTIDADE") ) and quantidade_enviado > 0 then
					pode_continuar = false
				%>
					<img src="imagens/upload_warning.png" style="cursor:pointer;" width="16" height="15" onclick="modalArquivo(<%=COD_PEDIDO%>,<%=rsProdutosAparte("COD_PEDIDO_PRODUTO")%>)">
				<%
				elseif quantidade_enviado = 0 then
					pode_continuar = false
				%>
					<img src="imagens/upload.png" style="cursor:pointer;" width="16" height="15" onclick="modalArquivo(<%=COD_PEDIDO%>,<%=rsProdutosAparte("COD_PEDIDO_PRODUTO")%>)">
				<%
				elseif (quantidade_enviado >  rsProdutosAparte("QUANTIDADE") ) and rsProdutosAparte("COD_PRODUTO") = 7 then
					pode_continuar = true
				%>
					<img src="imagens/upload_ok.png" style="cursor:pointer;" width="16" height="15" onclick="modalArquivo(<%=COD_PEDIDO%>,<%=rsProdutosAparte("COD_PEDIDO_PRODUTO")%>)">
				<%
				else

					pode_continuar = false
					response.write("Um erro inesperado ocorreu, entre em contato com o departamento de TI.")

					msgg = "quantidade_enviado: "&quantidade_enviado&"<br />"
					msgg = msgg & "quantidade: "&rsProdutosAparte("QUANTIDADE") &"<br />"

					call enviarEmail("joao@stillos.com.br","","","Erro na loja","erro no envio de imagens a loja: <br> "&msgg)
					'if rsProdutosAparte("COD_PRODUTO") = 7 then
						'pode_continuar = true
						'<!--<img src="imagens/upload_ok.png" style="cursor:pointer;" width="16" height="15" onclick="modalArquivo(<%=COD_PEDIDO ''=rsProdutosAparte("COD_PEDIDO_PRODUTO"))">
					'end if
				end if
			else 
				'pode_continuar = true
			%>
				<img src="imagens/icon_green.gif" width="16" height="15">
			<% end if %>
      	</div>
      </td>
    </tr>
	<%
	rsProdutosAparte.MoveNext()
	Wend
	%>
  </table></td>
</tr>
<%
else
pode_continuar = true
End If

rsProdutosAparte.Close()
Set rsProdutosAparte = Nothing

dim numero_pasta_vendedor
Set rsPasta = objConexao.Execute("SELECT numero_pasta_vendedor FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO = "&rs("COD_FORMANDO"))
if not rsPasta.eof then
	
	numero_pasta_vendedor = rsPasta(0)

end if
rsPasta.close
set rsPasta = Nothing

if numero_pasta_vendedor <> "" and pode_continuar then
	pode_continuar = true
else
	pode_continuar = false
end if

%>
<tr bgcolor="whitesmoke">
  <td><div align="center"><strong>Número da pasta:</strong></div></td>
</tr>
<tr>
	<td align="center">
		<form method="post" action="?cod_pedido=<%=criptografar(cod_pedido)%>">
			<input type="text" value="<%=numero_pasta_vendedor%>" name="numero_pasta" style="width:100px; height:30px; text-align:center;" maxlength="20">
			

			<input type="hidden" name="cod_formando" value="<%=rs("COD_FORMANDO")%>">
			<input type="hidden" name="incluir_pasta" value='S'>
			<input type="submit" value="GRAVAR PASTA">
		</form>
	</td>
</tr>
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
      <span class="style2">Total pago em produtos: 
      <%
		Dim rsProdutosTotal, QTD_EM_PRODUTOS, XTOTAL_AGREGADOS, XTOTAL_ALBUNS, XTOTAL_PRODUTOS_A_PARTE
		AbrirConexao()
		
		QTD_EM_PRODUTOS = 0
		XTOTAL_AGREGADOS = 0
		XTOTAL_ALBUNS = 0
		XTOTAL_PRODUTOS_A_PARTE = 0
		
			'VW_PEDIDOS_PRODUTOS_TOTAL_AGREGADOS
			Set rsProdutosTotal = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_AGREGADOS WHERE COD_PEDIDO="&COD_PEDIDO)
			
				if not rsProdutosTotal.Eof Then
					
					XTOTAL_AGREGADOS = formatNumber(0&rsProdutosTotal("XTOTAL_AGREGADOS"))/1
				
				End If
			
			
			rsProdutosTotal.Close()
			Set rsProdutosTotal = Nothing
			
			'VW_PEDIDOS_PRODUTOS_TOTAL_ALBUNS
			Set rsProdutosTotal = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_ALBUNS WHERE COD_PEDIDO="&COD_PEDIDO)
			
				if not rsProdutosTotal.Eof Then
					
					XTOTAL_ALBUNS = formatNumber(0&rsProdutosTotal("XTOTAL_ALBUNS"))/1
				
				End If
			
			rsProdutosTotal.Close()
			Set rsProdutosTotal = Nothing
			
			'VW_PEDIDOS_PRODUTOS_TOTAL_PRODUTOS_A_PARTE
			Set rsProdutosTotal = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_PRODUTOS_A_PARTE WHERE COD_PEDIDO="&COD_PEDIDO)
			
				if not rsProdutosTotal.Eof Then
					
					XTOTAL_PRODUTOS_A_PARTE = formatNumber(0&rsProdutosTotal("XTOTAL_PRODUTOS_A_PARTE"))/1
				
				End If

			rsProdutosTotal.Close()
			Set rsProdutosTotal = Nothing
			
								
			QTD_EM_PRODUTOS = formatCurrency(XTOTAL_AGREGADOS + XTOTAL_ALBUNS + XTOTAL_PRODUTOS_A_PARTE)

		%>
		<%=QTD_EM_PRODUTOS%></span>
		</p>
	</div>
  <p align="center" style="margin:0px">
  <%If CHAVE_AVANCAR = 1 and pode_continuar Then%>
  	<input name="btnAvancar" type="button" id="btnAvancar" onclick="abre('','pag3_desconto_forma_pagamento.asp?cod_pedido=<%=criptografar(COD_PEDIDO)%>')" value="Avan&ccedil;ar" />
  <%Else%>
  	<input name="btnAvancar" type="submit" onclick="alert('Não é possível avançar pois há produtos agregados que não foram salvos, ou com pendências')" id="btnAvancar" value="Avan&ccedil;ar" />
  <%End If%>
  <br />
  <br />
  </p>
  </p>
</div></td>
</tr>
</table>
<!-- #include file="rodape.asp" -->
<div class="modal fade" id='modalArquivoDiv'>
  <div class="modal-dialog">
    <div class="modal-content">
      <div class="modal-header">
        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
        <h4 class="modal-title">Upload de foto</h4>
      </div>
      <div class="modal-body">
        <iframe src="/loja/upload.asp" frameborder='0' width="100%" height="300px" id='iframeArquivo'></iframe>
      </div>
      <div class="modal-footer">
      </div>
    </div><!-- /.modal-content -->
  </div><!-- /.modal-dialog -->
</div><!-- /.modal -->
</body>
</html>