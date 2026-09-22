<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include virtual="gerencial_foto/files/includes/functions_local.asp" -->

<%
Dim COD_FORMANDO, COD_PRODUCAO, altComissao, brinde
COD_FORMANDO = Request.QueryString("cod_form")
COD_PRODUCAO = Request.QueryString("cod_prod")
altComissao  = formatNumber(0&Request.QueryString("altComissao"),0)
brinde = Request.QueryString("brinde")

If Request.Form("btnGravar") <> Empty then

	AbrirConexao()

	If Request.Form("COD_PRODUTO") <> "" and Request.Form("QUANTIDADE") <> "" then

		objConexao.Execute( "INSERT INTO TB_FICHAS_IDENTIFICACAO_BRINDES("&_
							"COD_FORMANDO ,"&_
							"COD_PRODUTO ,"&_
							"QUANTIDADE ,"&_
							"CADASTRADO_POR)"&_
							"VALUES ("&_
							"'"&COD_FORMANDO&"', '"&Request.Form("COD_PRODUTO")&"', '"&Request.Form("QUANTIDADE")&"', '"&Session("COD_USUARIO")&"')")	
		
		dim altComissao2
		altComissao2 = formatNumber(0&Request.Form("altComissao2"))

		If altComissao2  = 1 Then
			objConexao.Execute("UPDATE TB_FICHAS_IDENTIFICACAO SET BRINDE = '1', COMISSAO = '1' WHERE COD_FORMANDO = "&COD_FORMANDO&" LIMIT 1")
		End If

		Response.Redirect("dip_producoes_fichas_POPUP_brindes.asp?cod_form="&COD_FORMANDO&"&cod_prod="&COD_PRODUCAO&"&brinde="&brinde)
	End If

	FecharConexao()
End If

If Request.QueryString("comando") = "excluir" then
'Exclui direto

	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_BRINDES WHERE COD_FORMANDO="&COD_FORMANDO&" AND COD_PRODUTO= "&Request.QueryString("id"))
	FecharConexao()
	
	Response.Redirect("dip_producoes_fichas_POPUP_brindes.asp?cod_form="&COD_FORMANDO&"&cod_prod="&COD_PRODUCAO&"&brinde=1")

End If

if Request.QueryString("comando") = "aplicar_desconto" then

	AbrirConexao()

		dim aplicar_desconto, cod_producao_desconto

		aplicar_desconto = Request.form("aplicar_desconto")
		'COD_FORMANDO = Request.form("COD_FORMANDO")
		cod_producao_desconto = Request.form("cod_producao_desconto")

		if COD_FORMANDO <> "" and cod_producao_desconto <> "" then
			if aplicar_desconto = "S" then
				objConexao.Execute("UPDATE TB_FICHAS_IDENTIFICACAO SET cod_producao_desconto =  "&cod_producao_desconto&" WHERE COD_FORMANDO = "&sqlNumero(COD_FORMANDO))
			Else
				objConexao.Execute("UPDATE TB_FICHAS_IDENTIFICACAO SET cod_producao_desconto = 0 WHERE COD_FORMANDO = "&sqlNumero(COD_FORMANDO))
			end if
		end if

	FecharConexao()

	Response.Redirect("dip_producoes_fichas_POPUP_brindes.asp?cod_form="&COD_FORMANDO&"&cod_prod="&COD_PRODUCAO&"&brinde=1")

end if

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
	<link rel="stylesheet" type="text/css" href="css/bootstrap.min.css" />
	<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
	<script type="text/javascript" src='js/bootstrap.min.js'></script>
	<style type="text/css">
	<!--
	body {
		background-color: #FFFFFF;
	}
	table{
		border-collapse: collapse;
	}
	-->
	</style>
	<script type="text/javascript">
	function tabsMostra(idTabConteudo,qtdTabs,numeroTabAtual){
		for (i = 1; i < qtdTabs+1; i++) {
			document.getElementById('tab'+i).style.visibility = "hidden";
			document.getElementById('tab'+i).style.display = "none";
			document.getElementById("liTab"+i).className="";	
		}
			
		document.getElementById("liTab"+numeroTabAtual).className="current";
		document.getElementById(idTabConteudo).style.visibility = "visible";
		document.getElementById(idTabConteudo).style.display = "inline";						
	}
	</script>

<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr><td>

<%
Dim qtd_tabs
qtd_tabs = 2
%>

<div id="tabsJ">
	<ul>
		<li id="liTab1" class="current">
			<a href="javascript:tabsMostra('tab1',<%=qtd_tabs%>,1);">
				<span>Brindes</span>
			</a>
		</li>
		<li id="liTab2">
			<a href="javascript:tabsMostra('tab2',<%=qtd_tabs%>,2);">
				<span>Descontos</span>
			</a>
		</li>
	</ul>
</div>
<div id="tabsCaixa" style="height:350px; overflow:auto;">
	<div id="tab1">

	<br />&nbsp;

	<table width="100%" border="0" cellspacing="0" cellpadding="5">
    	<tr>
        	<td bgcolor="#FFFFFF">
				<div align="left">Cód. Formando.: <strong><%=digito(COD_FORMANDO)%></strong>
			</div>
		<% If Request.QueryString("brinde") = 1 Then %>
		<form name="FORMULARIO" method="post" action="dip_producoes_fichas_POPUP_brindes.asp?cod_form=<%=COD_FORMANDO%>&cod_prod=<%=COD_PRODUCAO%>&comando=incluir&brinde=<%=Request.QueryString("brinde")%>" style="margin:0px">

			<table width="100%" border="0" align="center" cellpadding="3" cellspacing="0" style="margin-top:5px; border-bottom:2px #3366CC solid; border-top:#3366CC solid 2px">
            <tr>
				<td width="24%"><div align="right">Incluir Brinde: </div></td>
				<td colspan="2">
				<select style="width:100%" name="COD_PRODUTO">
					<option value="0">---</option>
					<%
					AbrirConexao()
					Dim rsProdutos, sql

					'Busca foto padrão
					Dim rsPadrao, COD_FOTO_PADRAO, NOME_FOTO_PADRAO
					
					sql = "SELECT TB_PRODUTOS.NOME_PRODUTO, TB_PRODUCAO.COD_FOTO_PADRAO FROM TB_PRODUTOS INNER JOIN TB_PRODUCAO ON TB_PRODUTOS.COD_PRODUTO = TB_PRODUCAO.COD_FOTO_PADRAO WHERE TB_PRODUCAO.COD_PRODUCAO = "&COD_PRODUCAO
					Set rsPadrao = objConexao.Execute(sql)

						If not rsPadrao.Eof then
							COD_FOTO_PADRAO = rsPadrao("COD_FOTO_PADRAO")
							NOME_FOTO_PADRAO =rsPadrao("NOME_PRODUTO")
						Else
							COD_FOTO_PADRAO = 2
							NOME_FOTO_PADRAO = "FOTO"
						End If

					rsPadrao.Close()
					Set rsPadrao = Nothing
					 
					If int(COD_FOTO_PADRAO) = 2 Then
						SQL = "SELECT COD_PRODUTO, NOME_PRODUTO FROM TB_PRODUTOS WHERE ATIVO = '1' OR COD_PRODUTO = '"&COD_FOTO_PADRAO&"' ORDER BY NOME_PRODUTO"
					Else
						SQL = "SELECT COD_PRODUTO, NOME_PRODUTO FROM TB_PRODUTOS WHERE ATIVO = '1' ORDER BY NOME_PRODUTO"
					End If

					Set rsProdutos = objConexao.Execute(SQL) 
					While not rsProdutos.Eof
					%>
						<option value="<%=rsProdutos("COD_PRODUTO")%>"><%=rsProdutos("NOME_PRODUTO")%> <%If int(0&COD_FOTO_PADRAO) = rsProdutos("COD_PRODUTO") Then Response.Write(" (Foto padrão do álbum)") End If%></option>
					<%
					rsProdutos.MoveNext()
					Wend				 
					rsProdutos.Close()
					Set rsProdutos = Nothing
					
					FecharConexao()
					%>
				</select>
				</td>
			</tr>
			<tr>
				<td>
					<div align="right">Qtd.: </div>
				</td>
				<td width="17%">
					<input style="width:35px" type="text" name="QUANTIDADE" />
				</td>
				<td width="59%">
					<div align="center">
						<input class="btnGravar" type="submit" name="btnGravar" value=" " />
					</div>
				</td>
			</tr>
			<tr>
				<td colspan="3">*Foto padrão do álbum: <strong><%=NOME_FOTO_PADRAO%></strong>
					<input type="hidden" name="COD_FOTO_PADRAO" value="<%=COD_FOTO_PADRAO%>" />	
					<input type="hidden" name="altComissao2" value="<%=altComissao%>" />
					<br />
					<strong>
						Para lan&ccedil;ar fotos de brinde, lance pela mesma da produ&ccedil;&atilde;o.
					</strong>
				</td>
			</tr>
			</table>
		</form>
		<%Else%>
			<div class="msg" align="center">
				Para incluir um item � necess�rio selecionar a op��o SIM em brindes
			</div>
		<% End If

		AbrirConexao()
		Dim rs
		sql = "SELECT TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO, TB_PRODUTOS.NOME_PRODUTO, TB_FICHAS_IDENTIFICACAO_BRINDES.COD_BRINDE,  Sum(TB_FICHAS_IDENTIFICACAO_BRINDES.QUANTIDADE) as QUANTIDADE FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_BRINDES ON TB_PRODUTOS.COD_PRODUTO = TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO WHERE TB_FICHAS_IDENTIFICACAO_BRINDES.COD_FORMANDO="&COD_FORMANDO&" GROUP BY TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO"

		Set rs = objConexao.Execute(sql)

		If not rs.Eof  then
		%>
		<table width="98%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse; margin-top:6px">
			<tr bgcolor="whitesmoke">
				<td width="66%"><div align="center"><strong>Brinde</strong></div></td>
				<td width="18%"><div align="center"><strong>Qtd.</strong></div></td>
				<td width="16%"><div align="center"><strong>Excluir</strong></div></td>
			</tr>
			<%While not rs.Eof%>
			<tr>
				<td><div align="left"><%=rs("NOME_PRODUTO")%></div></td>
				<td><div align="center"><%=rs("QUANTIDADE")%></div></td>
				<td><div align="center"><a onClick="anula()" href="?cod_form=<%=COD_FORMANDO%>&cod_prod=<%=COD_PRODUCAO%>&id=<%=rs("COD_PRODUTO")%>&comando=excluir"><img src="images/icons/icon_red.gif" width="17" height="15" border="0" /></a></div></td>
			</tr>
			<%
			rs.MoveNext()
			Wend
		%>
		</table>
		<%
		End If
		
		rs.Close()
		Set rs = Nothing
		FecharConexao()
		%>	
	</td>
</tr>
</table>
</div>
<div id='tab2' class="tabs">
<br>
<%
AbrirConexao()
	dim cod_contrato

	Set rs = objConexao.Execute("SELECT cod_contrato FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO = "&sqlNumero(COD_FORMANDO))
	if not rs.Eof then
		cod_contrato = rs(0)
	end if
	rs.Close
	set rs = Nothing

	Dim descricao_desconto, contratos_desconto, qtd_produtos, count_prod

	Set rs = objConexao.Execute("SELECT * FROM VW_TB_PRODUCAO_DESCONTOS WHERE contratos like '%"&cod_contrato&"%' AND COD_PRODUCAO = "&sqlNumero(COD_PRODUCAO))
	if not rs.Eof  Then
		cod_producao_desconto = rs("cod_producao_desconto")
		descricao_desconto = rs("descricao")
		contratos_desconto = rs("contratos")
		qtd_produtos = rs("qtd_produtos")
	end if
	rs.Close
	set rs = Nothing

	if descricao_desconto <> "" then
		set rs = objConexao.Execute("SELECT cod_producao_desconto FROM TB_FICHAS_IDENTIFICACAO WHERE cod_formando = "&sqlNumero(COD_FORMANDO))
		if not rs.Eof then
			cod_producao_desconto_formando = rs(0)
		end if
		rs.Close
		set rs = Nothing

		if cod_producao_desconto_formando > "0" then
			aplicar_desconto = "S"
		Else
			aplicar_desconto = "N"
		end if

		%>
		<form action="?comando=aplicar_desconto&cod_form=<%=COD_FORMANDO%>&cod_prod=<%=COD_PRODUCAO%>" method="post">
			<div class="alert alert-warning">
				<strong>
					Existe um desconto para este contrato, deseja aplicar ao formando?
				</strong>
				<br />
				<table width="98%" align="center">
					<tr>
						<td>
							<%=padraoFormularioSelect("APLICAR_DESCONTO","","inputtamanho2","","","Selecione",aplicar_desconto, "PADRAO:S#Sim|N#N�o", "incluir", false, false, false) %>
						</td>
						<td>
							<input type="submit" value=" GRAVAR ">
						</td>
					</tr>
				</table>
			</div>
			<input type="hidden" name="cod_formando" value="<%=request.QueryString("COD_FORMANDO")%>">
			<input type="hidden" name="cod_producao_desconto" value="<%=cod_producao_desconto%>">
		</form>
		<table width="98%" border="1" align="center" cellpadding="4">
			<tr>
				<td width="100px;"><strong>Descri��o:</strong></td>
				<td><%=descricao_desconto%></td>
			</tr>
			<tr>
				<td><strong>Contratos:</strong></td>
				<td><%=contratos_desconto%></td>
			</tr>
			<tr>
				<td><strong>Qtd. Produtos:</strong></td>
				<td><%=qtd_produtos%></td>
			</tr>
		</table>
		<br>
		<%
		count_prod = 1
		Set rs = objConexao.Execute("SELECT * FROM VW_TB_PRODUCAO_DESCONTOS_PRODUTOS WHERE cod_producao_desconto = "&sqlNumero(cod_producao_desconto))
		if not rs.Eof then
		%>
		<table width="98%" border="1" align="center" cellpadding="4">
			<tr bgcolor="whitesmoke">
				<th colspan="3">Produtos com desconto</th>
			</tr>
			<tr bgcolor="whitesmoke">
				<th>#</th>
				<th>Produto</th>
				<th>Desconto</th>
			</tr>
			<% while not rs.Eof	%>			
			<tr>
				<td align="center"><%=count_prod%></td>
				<td><%=rs("NOME_PRODUTO")%></td>
				<td align="right"><%=rs("aliquota")%>%</td>
			</tr>
			<%
			count_prod = count_prod + 1
			rs.MoveNext
			wend
			%>
		</table>
		<%
		end if
		rs.Close
		set rs = Nothing
	Else
		%>
			<div class='alert alert-warning'>
				<strong>Nenhum desconto relacionado a este contrato.</strong>
			</div>
		<%	
	end if
FecharConexao()
%>
</div>

</div>
</td></tr>
</table>

</body>
</html>
