<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<%
Dim comando, botao, ACAO, ID_NEW
Dim COD_PRODUCAO, DESCRICAO, DATA_SAIDA, DATA_SAIDA_DIA, DATA_SAIDA_MES, DATA_SAIDA_ANO, PRODUTOS_AGREDADOS, PRODUCAO_FINALIZADA, ENCADERNACAO_ESPECIAL, NOME_PRODUTO_PADRAO, COD_FOTO_PADRAO

ACAO = Request.Form("ACAO")
comando = Request.QueryString("comando")
COD_PRODUCAO = Request.QueryString("cod_prod")

If comando = "incluir" and ACAO = "incluir" Then
	IncluirRegistro()
elseif comando = "alterar" and ACAO = "alterar" Then
	AlterarRegistro()
'Somente quando clicar no bot�o alterar o a vari�vel do form ACAO ser� carregada caso contrario � Empty
elseif comando = "alterar" and ACAO = "" Then
	LerRegistro()
End If

'/////////////////////////////////////
'Rotinas
'

Sub LerRegistro()
	
	AbrirConexao()
	Dim rs, sql
	sql = 	" SELECT "&_
			" TB_PRODUCAO.COD_FOTO_PADRAO, "&_
			" TB_PRODUCAO.DESCRICAO, "&_
			" TB_PRODUCAO.PRODUTOS_AGREDADOS, "&_
			" TB_PRODUCAO.PRODUCAO_FINALIZADA, "&_
			" TB_PRODUCAO.ENCADERNACAO_ESPECIAL,"&_
			" TB_PRODUTOS.NOME_PRODUTO "&_
			" FROM TB_PRODUTOS INNER JOIN TB_PRODUCAO ON TB_PRODUTOS.COD_PRODUTO = TB_PRODUCAO.COD_FOTO_PADRAO"&_
			" WHERE TB_PRODUCAO.COD_PRODUCAO="&COD_PRODUCAO
	
		Set rs = objConexao.Execute(sql)
			
			DESCRICAO 				= Ucase(rs("DESCRICAO"))
			PRODUTOS_AGREDADOS		= rs("PRODUTOS_AGREDADOS")
			PRODUCAO_FINALIZADA		= rs("PRODUCAO_FINALIZADA")
			ENCADERNACAO_ESPECIAL	= rs("ENCADERNACAO_ESPECIAL")
			NOME_PRODUTO_PADRAO		= rs("NOME_PRODUTO")
			
		rs.Close()
		Set rs = Nothing
	FecharConexao()
	
End Sub


Sub LerFormulario()
	
	DESCRICAO = Ucase(Request.Form("DESCRICAO"))
	PRODUTOS_AGREDADOS = Request.Form("PRODUTOS_AGREDADOS")
	ENCADERNACAO_ESPECIAL = Request.Form("ENCADERNACAO_ESPECIAL")
	COD_FOTO_PADRAO = Request.Form("COD_PRODUTO")

End Sub


Sub AlterarRegistro()

	LerFormulario()
	Dim rs, sql

	sql = 	"UPDATE TB_PRODUCAO SET "&_
			"DESCRICAO = "&sqlTexto(Ucase(DESCRICAO))&_
			"WHERE COD_PRODUCAO = "&Request.Form("COD_PRODUCAO")&" LIMIT 1"

	AbrirConexao()
		
		objConexao.Execute(sql)
		Response.Redirect("dip_producoes.asp?msg=Produ&ccedil;&atilde;o alterada com sucesso!&cod_prod="&COD_PRODUCAO)

	FecharConexao()

End Sub


Sub IncluirRegistro()
	AbrirConexao()
	LerFormulario()
	Dim rs, sql, ultimo_codigo
	
	'Padr�o para novas produ��es n�o incluir produtos agredados.
	PRODUTOS_AGREDADOS = 2
	'///////
	
	sql = "INSERT INTO TB_PRODUCAO (DESCRICAO, PRODUTOS_AGREDADOS, ENCADERNACAO_ESPECIAL, COD_FOTO_PADRAO)"&_
		  "VALUES ("&_
		  sqlTexto(DESCRICAO)&", "&PRODUTOS_AGREDADOS&", "&ENCADERNACAO_ESPECIAL&","&COD_FOTO_PADRAO&")"
		   ' Response.Write(sql)
		objConexao.Execute(sql)
		'Verifica o �ltimo id registrado
		Set rs = objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_PRODUCAO LIMIT 1")
			If not rs.Eof then
								
				ID_NEW  = rs("ULTIMO_ID")
											
			End If
			
		rs.Close()	
		Set rs = Nothing	
		

	FecharConexao()
		
	PRODUTOS_AGREDADOS = formatNumber(PRODUTOS_AGREDADOS,0)
	Dim redirecionar
	
	If PRODUTOS_AGREDADOS = 1 and comando = "incluir"then
	'Registra os produtos agregados, para depois importar os formandos
		redirecionar = "dip_producoes_produtos_agregados.asp?cod_prod="&ID_NEW
		Response.Redirect(redirecionar)
		Else
	'Vai direto para tela de importa��o de formando
		redirecionar = "dip_producoes_importar_formandos.asp?cod_prod="&ID_NEW
		Response.Redirect(redirecionar)
		End If
	
		
End Sub

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
	<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>
	
	<script language="javascript">	
		function mOvr(src, cor)
			{
			if (!src.contains(event.fromElement))
					{
						src.style.cursor = 'hand';
						src.bgColor = cor;
					}
			}
		
		function mOut(src, cor)
			{
				if (!src.contains(event.toElement))
					{
						src.style.cursor = 'default';
						src.bgColor = cor;
			
					}
			}
	</script>
	
	<script type="text/javascript" language="javascript">
	<!--AJUDA F1 E CTRL PARA BUSCA R�PIDA-->
		
	function abrir(theURL,winName,features) 
	{ 
	  window.open(theURL,winName,features);
	}

		
	function abre(url)
	{    
				//if(event.srcElement.tagName=='TD')
				//	{
				//		src.children.tags('A')[0].click(); 
				//	}
	location.href = url;

	}	
		
		
	

	function buscaIni(){
	d = document.buscaForm
	d.busca.focus()
	d.busca.select();
	}
	function mudacorBusca(CAMPO){
	CAMPO.style.backgroundColor = '#F5ECBA'
	}
	function mudacorBusca2(CAMPO){
	CAMPO.style.backgroundColor = '#FFFFFF'
	}

 function Verificar()
 {
  var tecla=window.event.keyCode;
  if (tecla==17) {
    buscaIni();
	}
 }
 


	<!--AJUDA F1 E CTRL PARA BUSCA RÁPIDA-->
	</script>
		<!-- CALEND�RIO SCRIPTS
	<script type="text/javascript">
	DHTMLSuite.include("calendar");
	</script>
	<script type="text/javascript">
	var calendarObjForForm = new DHTMLSuite.calendar({minuteDropDownInterval:10,numberOfRowsInHourDropDown:5,callbackFunctionOnDayClick:'getDateFromCalendar',isDragable:true,displayTimeBar:false});
	
		 
	calendarObjForForm.setCallbackFunctionOnClose('myOtherFunction');
	
	function myOtherFunction()
	{
		
		
	}
	function pickDate(buttonObj,inputObject)
	{
		calendarObjForForm.setCalendarPositionByHTMLElement(inputObject,0,inputObject.offsetHeight+2);	// Position the calendar right below the form input
		//calendarObjForForm.setInitialDateFromInput(inputObject,'yyyy-mm-dd hh:ii');	// Specify that the calendar should set it's initial date from the value of the input field.
		
		calendarObjForForm.setInitialDateFromInput(inputObject,'dd/mm/yyyy');	// Specify that the calendar should set it's initial date from the value of the input field.
		
		calendarObjForForm.addHtmlElementReference('myDate',inputObject);	// Adding a reference to this element so that I can pick it up in the getDateFromCalendar below(myInput is a unique key)
		if(calendarObjForForm.isVisible()){
			calendarObjForForm.hide();
		}else{
			calendarObjForForm.resetViewDisplayedMonth();	// This line resets the view back to the inital display, i.e. it displays the inital month and not the month it displayed the last time it was open.
			calendarObjForForm.display();
		}		
	}	
	/* inputArray is an associative array with the properties
	year
	month
	day
	hour
	minute
	calendarRef - Reference to the DHTMLSuite.calendar object.
	*/
	function getDateFromCalendar(inputArray)
	{
		var references = calendarObjForForm.getHtmlElementReferences(); // Get back reference to form field.
		references.myDate.value = inputArray.day + '/' + inputArray.month + '/' + inputArray.year
//		references.myDate.value = inputArray.year + '-' + inputArray.month + '-' + inputArray.day + ' ' + inputArray.hour + ':' + inputArray.minute;
		calendarObjForForm.hide();	
		
	}	
	</script>
	-->
	
	<script language="javascript">
	
	function validaForm(){
	d = document.producao
	var erro = 2
		<%If comando = "incluir" then%>
		if (d.COD_PRODUTO.value == ""){
			alert("Escolha a foto padrão da produção.");
			d.COD_PRODUTO.focus();
			erro = 1
			return false;
		}
		<%End If%>	
		
		if (d.DESCRICAO.value == ""){
			alert("Preencha o campo Descrição!");
			producao.DESCRICAO.focus();
			erro = 1
			return false;
		}			
		if (erro == 2){
		
			var pergunta = confirm('Tem certeza que deseja gravar esta produção?')
			if (pergunta==true){
				return true;
			}
			else{
				return false;
			}
			
		}
	
	}

	</script>
</head>

<body>
<!-- Code for the left panel -->
<div id="dhtmlgoodies_leftPanel">
	<a class="closeLink" href="#" onclick="initSlideLeftPanel();return false" style="font-size:15px">Fechar</a>
	<div id="leftPanelContent">
	<!-- This is the content -->
	<div>

<table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
  <td width="1%"><img src="images/icons/icon_help.jpg" /></td>
    <td width="95%" align="left"><strong style="font-size:16px">Ajuda do Sistema</strong><br />
<strong>Esta &eacute; a tela inicial do sistema, verifique o n&uacute;mero de acessos ao sistema e a data da &uacute;ltima utiliza&ccedil;&atilde;o.<br />
</strong>
	</td>
  </tr>
</table>

	</div>
	<!-- End content -->
	</div>
</div>
<!-- End code for the left panel -->
<table cellpadding="9px" cellspacing="0" width="100%">
<tr>
<td>
<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td>
		<!-- #include file="includes/topo.asp" -->
<table width="100%" style="border-bottom:#999999 solid dotted 1px" background="images/layout/fundo_menu_azul.jpg"border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
	<!-- #include file="includes/menu.asp" -->
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="10">
  <tr>
  <td>
   <%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 2
  BTN_FONTE = 1
  BTN_IMPRIMIR_LINK = "#"
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">

	  <h2>DIP - Produ&ccedil;&otilde;es / Cadastro e
	  Altera&ccedil;&otilde;es</h2>
	  
<%If PRODUCAO_FINALIZADA = 2 or comando = "incluir" then%>
	  <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
		  <h3 align="center">Produ&ccedil;&atilde;o</h3>
		  <form method="post" action="dip_producoes_alteracoes.asp?comando=<%=comando%>&cod_prod=<%=COD_PRODUCAO%>" name="producao">
		  <div id="comboTexto">
          <table width="90%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr>
            <td width="16%"><strong>Descri&ccedil;&atilde;o:<a href="#"><img src="images/icons/icone_interrogacao.jpg" width="17" height="16" border="0" onclick="alert('Este � o nome amig�vel da produ��o, utilize o padr�o existente.')" /></a></strong></td>
            <td width="84%"><input class="inputTamanho3" type="text" name="DESCRICAO" value="<%=DESCRICAO%>" />
              <input type="hidden" name="ACAO" value="<%=comando%>" /> <input type="hidden" name="COD_PRODUCAO" value="<%=COD_PRODUCAO%>" /></td>
            </tr>
<tr>
  <td><strong>Produtos Agregados:</strong><strong><a href="#"><img src="images/icons/icone_interrogacao.jpg" width="17" height="16" border="0" onclick="alert('Caso houver produto(s) agregado(s), estes dever�o ser informados no cadastro da produ��o, para posteriormente informar individualmente se o formando tem ou n�o o produto agregado')" /></a></strong> </td>
  <td>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
	<%If comando = "incluir" then %>
    Op&ccedil;&atilde;o desabilidata de produtos agregados
	N&atilde;o<%'Op&ccedil;&atilde;o desabilitada o valor atualmente esta fixo em 2 (N&Atilde;O)%>
  *Poder&aacute; ser cadastrado uma &uacute;nica vez no cadastro da produ&ccedil;&atilde;o.
  <%else
  	
		If PRODUTOS_AGREDADOS = 1 then
			Response.Write("Sim")
  		Else
			Response.Write("N&atilde;o")
		End If
  End If
  %>  </td>
    <td> 
	
	
	 <%If PRODUTOS_AGREDADOS = 1 Then
  Dim rsProAgre
  AbrirConexao()
  
  	Set rsProAgre = objConexao.Execute("SELECT TB_PRODUTOS.NOME_PRODUTO, TB_PRODUCAO_PRODUTOS_AGREGADOS.QUANTIDADE"&_
										" FROM TB_PRODUTOS "&_
										" INNER JOIN TB_PRODUCAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = TB_PRODUCAO_PRODUTOS_AGREGADOS.COD_PRODUTO "&_
										" WHERE TB_PRODUCAO_PRODUTOS_AGREGADOS.COD_PRODUCAO="&COD_PRODUCAO)
  	
	If not rsProAgre.Eof then%>
	
	    <table width="90%" border="1" align="center" 	cellpadding="3" style="border-collapse:collapse">
  <tr bgcolor="whitesmoke">
    <td colspan="2"><div align="center" class="textoMedio"><strong>Padr&atilde;o:</strong></div></td>
    </tr>
  <tr>
    <td width="67%"><div align="center">Produto</div></td>
    <td width="33%"><div align="center">Quant.</div></td>
  </tr>
	
		<%While not rsProAgre.Eof%>
		
		  <tr>
			<td><%=rsProAgre("NOME_PRODUTO")%></td>
			<td><div align="center"><%=rsProAgre("QUANTIDADE")%></div></td>
 		  </tr>
		

		
		<%rsProAgre.MoveNext()
		Wend%>
    	</table>
	<%Else%>
		
		<span class="msg">N&atilde;o h&aacute; produtos agregados cadastrados, para cadastrar <a class="linkMaster" href="#">clique aqui.</a></span>
	    <%End If
	
	rsProAgre.Close()
  	Set rsProAgre = Nothing
  FecharConexao()
  
  
  %>
  
  <%End If%>   </td>
  </tr>
</table>  </td>
</tr>
<tr>
  <td><strong>Encaderna&ccedil;&atilde;o especial:</strong></td>
  <td><%If comando = "incluir" then %>
	<select name="ENCADERNACAO_ESPECIAL">
    <option selected="selected" value="2">N&atilde;o</option>
    <option value="1">Sim</option>
  </select>
  *Poder&aacute; ser cadastrado uma &uacute;nica vez no cadastro da produ&ccedil;&atilde;o.
  <%else
  	
		If ENCADERNACAO_ESPECIAL = 1 then
			Response.Write("Sim")
  		Else
			Response.Write("N&atilde;o")
		End If
  End If
  %>  </td>
</tr>
<tr>
  <td><strong>Foto padr&atilde;o do &aacute;lbum: </strong></td>
  <td>
  <%If comando = "incluir" then %>
  <select name="COD_PRODUTO">
    <option value="">---</option>
    <%
				 AbrirConexao()
				 Dim rsProdutos, sql
				 
				 Set rsProdutos = objConexao.Execute("SELECT COD_PRODUTO, NOME_PRODUTO FROM TB_PRODUTOS WHERE (COD_CATEGORIA = 1 AND TIPO = 1 AND ATIVO = 1) OR COD_PRODUTO = 2 ORDER BY NOME_PRODUTO")
				 
				 While not rsProdutos.Eof
				 %>
    <option value="<%=rsProdutos("COD_PRODUTO")%>"><%=rsProdutos("NOME_PRODUTO")%></option>
    <%
				 rsProdutos.MoveNext()
				 Wend
				 %>
    <%
				rsProdutos.Close()
				Set rsProdutos = Nothing	
				FecharConexao()
				%>
  </select>
  <%else%>
  
  	<%=NOME_PRODUTO_PADRAO%>
  
  <%End If%>  </td>
</tr>
<tr>
  <td colspan="2"><div align="center">

    <input class="btnAvancar" type="submit" name="GRAVAR" value=" " onclick="return validaForm()" />

  </div></td>
</tr>
        </table>
		</div>
		  </form>
		  <p align="center">&nbsp;</p>		
		
          <div align="center">
            <p>
              <%Else%>
              </p>
            <p>&nbsp;</p>
            <p class="textoMedio">Aten&ccedil;&atilde;o	est&aacute; produ&ccedil;&atilde;o est&aacute; finalizada e portanto n&atilde;o poder&aacute; ser alterada por motivos de seguran&ccedil;a. </p>
            <p>&nbsp;</p>
            <p>             
              <%End If%>		
            </p>
          </div></td>

    </tr>
</table>

 
  
  </td>
    </tr>
</table>

  
  </td>
  </tr>
</table>
<!-- #include file="includes/rodape.asp" -->

  </td>
</tr>
</table>
</body>
</html>