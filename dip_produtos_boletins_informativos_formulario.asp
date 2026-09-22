<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 2
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
  Dim titulo_pagina
  titulo_pagina = "Boletim Informativo"
%>
<%
Dim calendario_necessario
calendario_necessario = "Não"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%

'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 15/10/2011 15:55:46

'Declara variáveis para utilização do sistema
Dim cod_boletim_informativo, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario, comandoAuxiliar

redirecionar_incluir	= "dip_produtos_boletins_informativos_formulario.asp"
redirecionar_alterar	= "dip_produtos_boletins_informativos.asp"
redirecionar_excluir	= "dip_produtos_boletins_informativos.asp"
redirecionar_validacao	= "dip_produtos_boletins_informativos.asp"
msgAoExcluir			= "Tem certeza que deseja excluir este Boletim Informativo?<br /><br />Os contratos que faziam parte deste boletim informativo passaram a adotar como preço base os preços da tabela padrão em todos os produtos e os pedidos que já foram processados com o preço deste boletim informativo Não seráo alterados."
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
comandoAuxiliar = Request.Form("comandoAuxiliar")
'Declara variáveis
Dim OBS, DATA_PROXIMA_ATUALIZACAO_MONETARIA, CADASTRADO_POR

If comando = Empty Then
comando = "incluir"
titulo_pagina = titulo_pagina&" / Inclusão"
End if

If comando = "incluir" and campo1 <> Empty Then
'Incluir novo registro
incluirRegistro(redirecionar_incluir)
titulo_pagina = titulo_pagina&" / Inclusão"
Elseif comando = "alterar" and campo1 = Empty Then
'Leitura inicial para possível alteração
cod_boletim_informativo = Request.QueryString("cod_boletim_informativo")
lerRegistro()
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "alterar" and campo1 <> Empty Then
'Gravar alteração
cod_boletim_informativo = Request.Form("cod_boletim_informativo")
alterarRegistro(redirecionar_alterar)
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 
'Leitura inicial para possível exclusão
cod_boletim_informativo = Request.QueryString("cod_boletim_informativo")
lerRegistro()
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
'Exclui registro
cod_boletim_informativo = Request.Form("cod_boletim_informativo")
excluirRegistro(redirecionar_excluir)
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
'Visualiza registro
cod_boletim_informativo = Request.QueryString("cod_boletim_informativo")
lerRegistro()
titulo_pagina = titulo_pagina&" / Visualização"
End If


'Funções

Sub lerFormulario()

OBS = Request.Form("OBS")
DATA_PROXIMA_ATUALIZACAO_MONETARIA = Request.Form("DATA_PROXIMA_ATUALIZACAO_MONETARIA")
CADASTRADO_POR = Session("cod_usuario")

End sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM TB_BOLETIM_INFORMATIVO WHERE cod_boletim_informativo = "&cod_boletim_informativo)

				OBS=trim(rs("OBS"))
				DATA_PROXIMA_ATUALIZACAO_MONETARIA=trim(rs("DATA_PROXIMA_ATUALIZACAO_MONETARIA"))
				CADASTRADO_POR=trim(rs("CADASTRADO_POR"))

		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("INSERT INTO TB_BOLETIM_INFORMATIVO (OBS, DATA_PROXIMA_ATUALIZACAO_MONETARIA, CADASTRADO_POR) VALUES ("&sqlTexto(OBS)&", "&sqlData(DATA_PROXIMA_ATUALIZACAO_MONETARIA)&", "&sqlNumero(CADASTRADO_POR)&")")

		'Encontrar chave criada
		Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_BOLETIM_INFORMATIVO LIMIT 1")
			If not rs.Eof Then
				cod_boletim_informativo = rs("ULTIMO_ID")
			End if 
		rs.Close
		Set rs=Nothing
	FecharConexao()
	
	if comandoAuxiliar <> "" Then
		executaComandoAuxiliar(cod_boletim_informativo)
	End If
	
	Response.Redirect(redirecionar&"?cod_boletim_informativo="&cod_boletim_informativo&"&comando=alterar")
End sub

Sub alterarRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("UPDATE TB_BOLETIM_INFORMATIVO SET OBS = "&sqlTexto(OBS)&", DATA_PROXIMA_ATUALIZACAO_MONETARIA = "&sqlData(DATA_PROXIMA_ATUALIZACAO_MONETARIA)&", CADASTRADO_POR = "&sqlNumero(CADASTRADO_POR)&" WHERE cod_boletim_informativo = "&cod_boletim_informativo)
	FecharConexao()
	
	if comandoAuxiliar <> "" Then
		executaComandoAuxiliar(cod_boletim_informativo)
	End If
	
	Response.Redirect(redirecionar&"?cod_boletim_informativo="&cod_boletim_informativo&"&codmsg=2")
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		'Exclui cabeçalho
		objConexao.Execute("DELETE FROM TB_BOLETIM_INFORMATIVO WHERE cod_boletim_informativo = "&cod_boletim_informativo)
		'Exclui filhas
		objConexao.Execute("DELETE FROM TB_BOLETIM_INFORMATIVO_CONTRATOS WHERE cod_boletim_informativo = "&cod_boletim_informativo)
		objConexao.Execute("DELETE FROM TB_BOLETIM_INFORMATIVO_PRODUTOS WHERE cod_boletim_informativo = "&cod_boletim_informativo)
		
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_boletim_informativo="&cod_boletim_informativo&"&codmsg=3")
End sub

Sub executaComandoAuxiliar(cod_boletim_informativo)
	Dim COD_CONTRATO, COD_PRODUTO
	Dim VALOR_A_VISTA,VALOR_3X, VALOR_PARCELADO
	
	COD_CONTRATO = Request.Form("COD_CONTRATO_auxiliar")
	COD_PRODUTO = Request.Form("COD_PRODUTO_auxiliar")
	VALOR_A_VISTA = Request.Form("VALOR_A_VISTA")
	VALOR_PARCELADO = Request.Form("VALOR_PARCELADO")
	VALOR_3X = Request.Form("VALOR_3X")
	
	AbrirConexao()
		
		if comandoAuxiliar = "incluirContrato" Then
			
			'1) Verifica se o contrato já está cadastrado para algum Boletim Informativo, inclusive este.
			sql = "SELECT COD_CONTRATO FROM TB_BOLETIM_INFORMATIVO_CONTRATOS WHERE COD_CONTRATO = '"&COD_CONTRATO&"'"
			
			'Response.Redirect("adm_index.asp?msg="&sql)
			Set rs=objConexao.Execute(sql)
				
				if not rs.Eof Then
					'Já utilizado
					Response.Redirect("dip_produtos_boletins_informativos_formulario.asp?comando=alterar&cod_boletim_informativo="&cod_boletim_informativo&"&codmsg=34&sucesso=2")
					
				Else
					'Verifica se o contrato existe
					Set rs2=objConexao.Execute("SELECT COD_CONTRATO FROM TB_ENTIDADES WHERE COD_CONTRATO = '"&COD_CONTRATO&"'")	
					
						if not rs2.Eof Then
							'Existe, inclui no boletim informativo
							objConexao.Execute("INSERT INTO TB_BOLETIM_INFORMATIVO_CONTRATOS (COD_BOLETIM_INFORMATIVO,COD_CONTRATO) VALUES ('"&cod_boletim_informativo&"', '"&COD_CONTRATO&"')")
							
							Response.Redirect("dip_produtos_boletins_informativos_formulario.asp?comando=alterar&cod_boletim_informativo="&cod_boletim_informativo&"&codmsg=31")
							
						Else
							'Não existe
							Response.Redirect("dip_produtos_boletins_informativos_formulario.asp?comando=alterar&cod_boletim_informativo="&cod_boletim_informativo&"&codmsg=35&sucesso=2")
						End If
				
					rs2.Close
					Set rs2=Nothing
				
				End If
				
			rs.Close
			Set rs=Nothing
			
		Elseif comandoAuxiliar = "excluirContrato" Then
			'Excluir contrato
			objConexao.Execute("DELETE FROM TB_BOLETIM_INFORMATIVO_CONTRATOS WHERE COD_CONTRATO = '"&COD_CONTRATO&"' limit 1")	
			Response.Redirect("dip_produtos_boletins_informativos_formulario.asp?comando=alterar&cod_boletim_informativo="&cod_boletim_informativo&"&codmsg=33")		
		Elseif comandoAuxiliar = "incluirProduto" Then

			'1) Verifica se o produto já está cadastrado para este Boletim Informativo.
			Set rs=objConexao.Execute("SELECT cod_produto FROM TB_BOLETIM_INFORMATIVO_PRODUTOS WHERE COD_PRODUTO = '"&COD_PRODUTO&"' AND COD_BOLETIM_INFORMATIVO = '"&cod_boletim_informativo&"'")
				
				if not rs.Eof Then
					'já utilizado
					Response.Redirect("dip_produtos_boletins_informativos_formulario.asp?comando=alterar&cod_boletim_informativo="&cod_boletim_informativo&"&codmsg=24&sucesso=2")
					
				Else
					
					'Insere o produto
					objConexao.Execute("INSERT INTO TB_BOLETIM_INFORMATIVO_PRODUTOS (COD_BOLETIM_INFORMATIVO,COD_PRODUTO,VALOR_A_VISTA,VALOR_PARCELADO,VALOR_3X) VALUES ('"&COD_BOLETIM_INFORMATIVO&"', '"&COD_PRODUTO&"',"&sqlNumero(VALOR_A_VISTA)&","&sqlNumero(VALOR_PARCELADO)&","&sqlNumero(VALOR_3X)&")")
							
					Response.Redirect("dip_produtos_boletins_informativos_formulario.asp?comando=alterar&cod_boletim_informativo="&cod_boletim_informativo&"&codmsg=21")
				
				
				End If
				
			rs.Close
			Set rs=Nothing
			
		
		Elseif comandoAuxiliar = "excluirProduto" Then		
			
			'Excluir contrato
			objConexao.Execute("DELETE FROM TB_BOLETIM_INFORMATIVO_PRODUTOS WHERE COD_BOLETIM_INFORMATIVO = '"&COD_BOLETIM_INFORMATIVO&"' AND COD_PRODUTO = '"&COD_PRODUTO&"' limit 1")	
			Response.Redirect("dip_produtos_boletins_informativos_formulario.asp?comando=alterar&cod_boletim_informativo="&cod_boletim_informativo&"&codmsg=23")
			
		End If
		
		
	FecharConexao()
	
End Sub
'----------------------------------------------------------------------------------------

%>
<script type="text/javascript">
function gravarBoletim(tipo,comandoAuxiliar){

	var d = document.FORMULARIO_PADRAO;

	if(tipo == 1){
		//Formulário Padrão
		salvar();
	}else if(tipo == 2){
		//Contratos
		if(comandoAuxiliar == "incluirContrato" && d.COD_CONTRATO_auxiliar.value == ""){
			alert('Prenche o número do contrato antes de continuar!');
			d.COD_CONTRATO_auxiliar.focus();
		}else{
			if(validaFormMaster()){
				d.comandoAuxiliar.value = comandoAuxiliar;
				d.submit();
			}
		}	
	}else if(tipo == 3){
		//Formulário Padrão
		if(comandoAuxiliar == "incluirProduto" && (d.COD_PRODUTO_INCLUIR.value == "" || d.VALOR_A_VISTA.value == "" || d.VALOR_PARCELADO.value == "" || d.VALOR_3X.value == "")){
			alert('Para incluir um produto é necessário, selecionar o produto e colocar o valor à vista e a prazo.');
			d.COD_PRODUTO_INCLUIR.focus();
		}else{
			if(validaFormMaster()){
				d.COD_PRODUTO_auxiliar.value = d.COD_PRODUTO_INCLUIR.value;
				d.comandoAuxiliar.value = comandoAuxiliar;
				d.submit();
			}
		}
	}
}

function excluirContrato(cod_contrato){

	var d = document.FORMULARIO_PADRAO;	
	var msg = "Tem certeza que deseja excluir este cod. contrato: "+cod_contrato+"?";
	if( YNconfirm(msg)){
		if(validaFormMaster()){
			d.comandoAuxiliar.value = "excluirContrato";
			d.COD_CONTRATO_auxiliar.value = cod_contrato;
			d.submit();
		}
	}
}

function excluirProduto(cod_produto,nome_produto){

	var d = document.FORMULARIO_PADRAO;	
	var msg = "Tem certeza que deseja excluir este produto: "+nome_produto+"?";
	if( YNconfirm(msg)){
		if(validaFormMaster()){
			d.comandoAuxiliar.value = "excluirProduto";
			d.COD_PRODUTO_auxiliar.value = cod_produto;
			d.submit();
		}
	}
}
</script>


<h2><%=titulo_pagina%></h2><br />
	<div class="subMenu">Sub-menu: <a href="dip_produtos_boletins_informativos.asp">Voltar para Listagem de Boletins</a></div>
<%
	   Response.Write("<br />")
	Dim sucesso, codmsg
	codmsg	= CStr(""&Request.QueryString("codmsg"))
	sucesso = CStr(""&Request.QueryString("sucesso"))
		
	If sucesso = "" and codmsg <> "" Then%>
    	<div align="center" class="msgGrandeSucesso">
        	<%
			If codmsg = "1" Then
				Response.Write("Boletim Informativo incluído com sucesso!")
			Elseif codmsg = "2" Then
				Response.Write("Boletim Informativo alterado com sucesso!")
			Elseif codmsg = "3" Then	
				Response.Write("Boletim Informativo excluído com sucesso!")
			Elseif codmsg = "21" Then	
				Response.Write("Produto incluído com sucesso!")
			Elseif codmsg = "23" Then	
				Response.Write("Produto excluído com sucesso!")
			Elseif codmsg = "31" Then	
				Response.Write("Contrato incluído com sucesso!")
			Elseif codmsg = "33" Then	
				Response.Write("Contrato excluído com sucesso!")
			End If	
			%>
            
        
        </div>
    <%ElseIf sucesso <> "" and codmsg <> "" Then%>
        <div align="center" class="msgGrandeErro">
			<%
			If codmsg = "24" Then
				Response.Write("Produto já utilizado neste mesmo Boletim Informativo, se desejar alterar exclua primeiro o produto e inclua novamente com os preços alterados.")
			ElseIf codmsg = "34" Then
				Response.Write("Cod. Contrato já utilizado em Boletim Informativo. Atenção é permitido a utilização de apenas uma vez do contrato")
			Elseif codmsg = "35" Then	
				Response.Write("Cod. Contrato Não encontrado no sistema!")	
			End If	
			%>
</div>
	<%End If%>



<%If comando = "excluir" Then%>
<div class='msgExcluir'><%=msgAoExcluir%></div>
<%End If%>
<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post'>
<fieldset style="padding:10px; margin:10px">
<legend>Dados B&aacute;sicos do  Boletim Informativo</legend>
<table width='90%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
<%If comando <> "incluir" Then %><tr>
<td width='28%'><strong>cod. Boletim Informativo </strong></td>
<td width='72%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<span class='textoMaior'><strong><%=COD_BOLETIM_INFORMATIVO%></strong></span><%Else%>
<%If not isNull(COD_BOLETIM_INFORMATIVO) Then%>
<%=COD_BOLETIM_INFORMATIVO%>
<%End If%>
<%End If%>
</div></td>
</tr>
<%End If%><tr>
<td width='28%'><p><strong>Observação Complemento</strong><br />(Esta Observação será colocada na folha de impressão do Boletim Informativo)</p>  </td>
<td width='72%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<table style='width:420px;' border='0' cellspacing='0' cellpadding='0'>
<tr>
<td width='57'><textarea rows='3' name='OBS' class='inputTamanho3' onkeydown='textCounter(this.form.OBS,this.form.charRestantes_OBS,300);' onkeyup='textCounter(this.form.OBS,this.form.charRestantes_OBS,300);'><%=OBS%></textarea></td>
</tr>
<tr>
<td><table width='50' border='0' align='right' cellpadding='0' cellspacing='0'>
<tr>
<td><div align='center'>
<input class='inputContaCaracteres' tabindex='10000' type='text' style="width:40px; border:solid 1px #FFFFFF; background-color:#FFFFFF" name='charRestantes_OBS' value="<%=300- Len(OBS)%>" />
</div></td>
</tr>
</table></td>
</tr>
</table>
<%Else%>
<%If not isNull(OBS) Then%>
<%=Replace(OBS,Chr(13),"<br />")%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='28%'><strong>Data Próxima Atualização Monet&aacute;ria *</strong></td>
<td width='72%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<table width='135' border='0' cellspacing='0' cellpadding='0'>
<tr>
<td width='102'><input class='campoData' maxlength='10' onKeyDown='mascara(this,data);' onKeyPress='mascara(this,data);' onKeyUp='mascara(this,data);' type='text' name='DATA_PROXIMA_ATUALIZACAO_MONETARIA' value="<%If(ArrumaData(DATA_PROXIMA_ATUALIZACAO_MONETARIA)) = "-" Then Response.Write("") Else Response.Write(ArrumaData(DATA_PROXIMA_ATUALIZACAO_MONETARIA))  End If%>" onblur='validarData(this)' /></td>
<td width='298'>&nbsp;</td>
</tr>
</table>
<%Else%>
<%If not isNull(DATA_PROXIMA_ATUALIZACAO_MONETARIA) Then%>
<%=DATA_PROXIMA_ATUALIZACAO_MONETARIA%>
<%End If%>
<%End If%>
</div></td>
</tr>


<%If comando <> "visualizar" Then%>
<tr>
  <td colspan='2' align='center'><div align="center"><a href="javascript:gravarBoletim(1,'');"><img border="0" src="images/layout/btnGravar.jpg" /></a></div></td>
</tr>
<%End If%>
</table>
</fieldset>
<br />
<fieldset style="padding:10px; margin:10px">
<legend>Contratos do Boletim Informativo</legend>
<table width='90%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
<tr>
<td width='14%'><span style="font-weight: bold">Incluir Contrato:</span></td>
<td width='15%'><input type="text" class="inputTamanho1" name="COD_CONTRATO_auxiliar" /></td>
<td width='71%'><div align="left"><a href="javascript:gravarBoletim(2,'incluirContrato');"><img border="0" src="images/layout/btnIncluir.jpg" /></a></div></td>
</tr>
</table>
<BR />
	<%
	AbrirConexao()
	
	Set rs=objConexao.Execute("SELECT * FROM TB_BOLETIM_INFORMATIVO_CONTRATOS WHERE COD_BOLETIM_INFORMATIVO = '"&COD_BOLETIM_INFORMATIVO&"'")
		If not rs.Eof Then
	%>
        <table width='90%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
        <tr>
        <td width='60%'><div align="center" style="font-weight: bold">Contratos</div></td>
        <td width='40%'><div align="center" style="font-weight: bold">Excluir</div></td>
        </tr>
        <%While not rs.Eof%>
        <tr>
          <td>
            <div align="left">
              <%
				Set rs2=objConexao.Execute("SELECT DESCRICAO FROM TB_ENTIDADES WHERE COD_CONTRATO = '"&rs("COD_CONTRATO")&"'")
					
					If not rs2.Eof Then
						Response.Write(rs2(0))
					End If
					
				rs2.Close
				Set rs2=Nothing
				%>
                </div></td>
          <td><div align="center"><a href="javascript:excluirContrato('<%=rs("COD_CONTRATO")%>')"> <img alt="Excluir" src="images/icons/icon_excluir.png" border="0" /></a> </div></td>
        </tr>
        <%
		rs.MoveNext()
		Wend
		
		%>
        </table>
	<%Else%>
        <div align="center">Não há contratos lançados para este Boletim informativo!</div>
    <%
	End If
	
	rs.Close
	Set rs=Nothing
	%>
    
    

</fieldset>
<br />
<fieldset style="padding:10px; margin:10px">
<legend>Produtos do Boletim Informativo</legend>
<table width='90%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
<tr>
<td width='14%'><div align="center"><span style="font-weight: bold">Incluir Produto:</span></div></td>
<td width="86%">
  <div align="left">
    <INPUT type="hidden" name="COD_PRODUTO_auxiliar" />
    <select name="COD_PRODUTO_INCLUIR" class="inputTamanho3">
      <option value="">Selecione</option>
        <%
Set rs=objConexao.Execute("SELECT COD_PRODUTO, NOME_PRODUTO FROM TB_PRODUTOS WHERE ATIVO = 1 ORDER BY NOME_PRODUTO")
	While not rs.Eof
%>
      <option value="<%=rs(0)%>"><%=rs(1)%></option>
        <%
	rs.MoveNext()
	Wend

rs.Close
Set rs=nothing
%>
    </select>
  </div></td>
</tr>
<tr>
  <td><div align="center" style="font-weight: bold">Valor &agrave; vista</div></td>
  <td><input name='VALOR_A_VISTA' value="" class='inputTamanho1 padrao_valor' type='text' maxlength='200' /></td>
  </tr>
<tr>
  <td><div align="center" style="font-weight: bold">Valor em 3X</div></td>
  <td><input name='VALOR_3X' value="" class='inputTamanho1 padrao_valor' type='text' maxlength='200' /></td>
  </tr>
<tr>
  <td><div align="center" style="font-weight: bold">Valor parcelado</div></td>
  <td><input name='VALOR_PARCELADO' value="" class='inputTamanho1 padrao_valor' type='text' maxlength='200' /></td>
  </tr>
<tr>
  <td colspan="2"><div align="left"><a href="javascript:gravarBoletim(3,'incluirProduto');"><img border="0" src="images/layout/btnIncluir.jpg" /></a></div></td>
  </tr>
</table>
<BR />
	<%
	Set rs=objConexao.Execute("SELECT * FROM VW_TB_BOLETIM_INFORMATIVO_PRODUTOS WHERE COD_BOLETIM_INFORMATIVO = '"&COD_BOLETIM_INFORMATIVO&"' ORDER BY NOME_PRODUTO ASC")
		If not rs.Eof Then
	%>

<table width='90%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
<tr>
<td width='29%'><div align="center" style="font-weight: bold">Produtos</div></td>
<td width='21%'><div align="center" style="font-weight: bold">Valor &agrave; vista</div></td>
<td width='20%'><div align="center" style="font-weight: bold">Valor em 3X</div></td>
<td width='17%'><div align="center" style="font-weight: bold">Valor parcelado</div></td>
<td width='13%'><div align="center" style="font-weight: bold">Excluir</div></td>
</tr>
<%
While not rs.Eof
%>
<tr>
  <td><div align="left"><%=rs("NOME_PRODUTO")%></div></td>
  <td><div align="right"><%=formatCurrency(rs("VALOR_A_VISTA"))%></div></td>
  <td><div align="right"><%=formatCurrency(rs("VALOR_3X"))%></div></td>
  <td><div align="right"><%=formatCurrency(rs("VALOR_PARCELADO"))%></div></td>
  <td><div align="center"><a href="javascript:excluirProduto('<%=rs("COD_PRODUTO")%>','<%=rs("NOME_PRODUTO")%>')"> <img alt="Excluir" src="images/icons/icon_excluir.png" border="0" /></a> </div></td>
</tr>
<%
rs.MoveNext()
Wend
%>
</table>

	<%Else%>
    	
        <div align="center">Não há produtos lançados para este Boletim informativo!</div>
    
    <%
	End If
	
	rs.Close
	Set rs=Nothing
	
	FecharConexao()
	%>

</fieldset>
<input type='hidden' value='1' name='campo1'  />
<input type='hidden' value='' name='comandoAuxiliar'  />
<input type='hidden' name='COD_BOLETIM_INFORMATIVO' value="<%=COD_BOLETIM_INFORMATIVO%>" /> <input type='hidden' name='CADASTRADO_POR' value="<%=CADASTRADO_POR%>" /> </form>
<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('DATA_PROXIMA_ATUALIZACAO_MONETARIA#Data Próxima Atualização#1',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->