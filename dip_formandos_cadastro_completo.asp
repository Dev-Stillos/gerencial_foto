<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim COD_PRODUCAO, COD_FORMANDO, COMANDO, SISTEMA_NOVO_QUESTIONAR, ESCOLHA_PRODUCAO
COD_PRODUCAO 			= Request.QueryString("cod_prod")
COD_FORMANDO 			= Request.QueryString("cod_formando")
COMANDO 				= Request.QueryString("comando")

If COD_PRODUCAO = Empty or COD_PRODUCAO = 0 Then
	COD_PRODUCAO			= 0
	SISTEMA_NOVO_QUESTIONAR	= 1 
	ESCOLHA_PRODUCAO		= 2
Else
	AbrirConexao()
		Set rsProducao = objConexao.Execute("SELECT PRODUTOS_AGREDADOS FROM TB_PRODUCAO WHERE COD_PRODUCAO="&COD_PRODUCAO)
				
				If rsProducao("PRODUTOS_AGREDADOS") = 1 Then
					ESCOLHA_PRODUCAO = 1
				Else
					ESCOLHA_PRODUCAO = 2
				End If
		rsProducao.Close
		Set rsProducao = Nothing
	FecharConexao()
	
	
End If

If COMANDO = "novo" Then
	'Obtem um novo c�digo de formando
	Dim rsCodNew, sql
	AbrirConexao()
	sql = 	"INSERT INTO TB_FICHAS_IDENTIFICACAO_CONTAGEM_NUMERO_FORMANDO_CADASTRADO (COD_FORMANDO_CADASTRADO) "&_
			"VALUES (NULL)"
			
	objConexao.Execute(sql)		
			
	Set rsCodNew = objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_FICHAS_IDENTIFICACAO_CONTAGEM_NUMERO_FORMANDO_CADASTRADO LIMIT 1")
	
			COD_FORMANDO = rsCodNew("ULTIMO_ID")
	
	rsCodNew.Close
	Set rsCodNew = Nothing
	FecharConexao()
	
	Response.Redirect("dip_formandos_cadastro_completo.asp?comando=incluir&cod_prod="&COD_PRODUCAO&"&cod_formando="&COD_FORMANDO)
End If

If comando ="incluir2" and Request.Form("btnGravar") <> Empty Then
AbrirConexao()
	'TB1
	Dim SISTEMA_NOVO, COD_CONTRATO, COD_CURSO, NOME_CURSO, NOME_FORMANDO, ANO_CONCLUSAO, SEMESTRE_CONCLUSAO, ALBUM, INICIAL, BRINDE, COMISSAO, PRODUTO_AGREGADO
	
	SISTEMA_NOVO		=""&Request.Form("SISTEMA_NOVO")
	
	If SISTEMA_NOVO = Empty or SISTEMA_NOVO = "" Then
		SISTEMA_NOVO = 1
	End If
	
	COD_CONTRATO		=Request.Form("COD_CONTRATO")
	COD_CURSO			=Request.Form("COD_CURSO")
	
	Dim rsCursoa
	Set rsCursoa = objConexao.Execute("SELECT * FROM TB_FORMANDOS_IMP WHERE CODIGO_CURSO="&COD_CURSO)
			
			If not rsCursoa.Eof Then 
				NOME_CURSO = rsCursoa("NOME_CURSO")
			End If
				
	rsCursoa.Close
	Set rsCursoa = Nothing
	

	NOME_FORMANDO		=TrocarAcento(Ucase(Request.Form("NOME_FORMANDO")))
	
	'--------------------------------------------
	'Altera��o realizada em: 06/08/2012
	'Agora n�o poder� ser separado o ano e semestre do formando, dever� obdecer a entidade selecionada
	
	'///NOVO	
	Set rs=objConexao.Execute("SELECT * FROM  TB_ENTIDADES WHERE COD_CONTRATO = "&sqlTexto(COD_CONTRATO))
			
		If not rs.Eof Then
			ANO_CONCLUSAO = rs("ANO")
			SEMESTRE_CONCLUSAO	= Ucase(""&rs("SEMESTRE"))
		End If
		
	rs.Close
	Set rs=Nothing
		
	'///DESABILITADO
		'ANO_CONCLUSAO		=Request.Form("ANO_CONCLUSAO")
		'SEMESTRE_CONCLUSAO	=Ucase(Request.Form("SEMESTRE_CONCLUSAO"))
		
	'--------------------------------------------
	
	
	ALBUM				=formatNumber(0&Request.Form("ALBUM"),0)
	INICIAL				=formatNumber(0&Request.Form("INICIAL"),0)
	BRINDE				=formatNumber(0&Request.Form("BRINDE"),0)
	COMISSAO			=formatNumber(0&Request.Form("COMISSAO"),0)
	PRODUTO_AGREGADO	=""&Request.Form("PRODUTO_AGREGADO")
	
	If PRODUTO_AGREGADO = Empty or PRODUTO_AGREGADO = "" Then
		PRODUTO_AGREGADO = 2				
	Elseif PRODUTO_AGREGADO = 1 Then 
		'Seleciona o produto agregado para inclus�o
		Dim rsAgregado
		Set rsAgregado = objConexao.Execute("SELECT * FROM TB_PRODUCAO_PRODUTOS_AGREGADOS WHERE COD_PRODUCAO="&COD_PRODUCAO)
				dim qtdAgregado
				qtdAgregado = 1
				While not rsAgregado.Eof
					qtdAgregado = qtdAgregado+1
					sql = 	"INSERT INTO TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ("&_
							"TIPO ,"&_
							"COD_PRODUCAO ,"&_
							"COD_FORMANDO ,"&_
							"COD_PRODUTO ,"&_
							"QUANTIDADE ,"&_
							"CADASTRADO_POR "&_
							")"&_
							"VALUES ("&_
							qtdAgregado&", '"&COD_PRODUCAO&"', '"&COD_FORMANDO&"1', "&rsAgregado("COD_PRODUTO")&", "&rsAgregado("QUANTIDADE")&", "&Session("COD_USUARIO")&_
							")"
				Response.Write(sql&"<br>"&"<br>")
				objConexao.Execute(sql)
				
				rsAgregado.MoveNext()
				Wend
				
		rsAgregado.Close
		Set rsAgregado=Nothing
	End If
	'TB2
	Dim ENDERECO, CEP, EMAIL_FORMANDO, FONE_RES, FONE_CEL, BAIRRO, CIDADE, ESTADO 
	
	ENDERECO			=TrocarAcento(Ucase(Request.Form("ENDERECO")))
	CEP					=Request.Form("CEP")
	CEP					=Replace(CEP,"-","")
	CEP					=Replace(CEP," ","")
	CEP					=Replace(CEP,".","")
	EMAIL_FORMANDO 		=Request.Form("EMAIL_FORMANDO")
	FONE_RES			=Request.Form("FONE_RES")
	FONE_CEL			=Request.Form("FONE_CEL")
	BAIRRO				=TrocarAcento(Ucase(Request.Form("BAIRRO")))
	CIDADE				=TrocarAcento(Ucase(Request.Form("CIDADE")))
	ESTADO 				=Request.Form("ESTADO")
	
	
	
	
	sql=	"INSERT INTO TB_FICHAS_IDENTIFICACAO ("&_
			"COD_FORMANDO ,"&_
			"SISTEMA_NOVO ,"&_
			"COD_PRODUCAO ,"&_
			"COD_CONTRATO ,"&_
			"COD_CURSO ,"&_
			"NOME_CURSO ,"&_
			"EMAIL_FORMANDO ,"&_
			"CPF ,"&_
			"NOME_FORMANDO ,"&_
			"ANO_CONCLUSAO ,"&_
			"SEMESTRE_CONCLUSAO ,"&_
			"ALBUM ,"&_
			"INICIAL ,"&_
			"RETIRADO ,"&_
			"BRINDE ,"&_
			"COMISSAO ,"&_
			"PRODUTO_AGREGADO "&_	
			")"&_
			"VALUES ("&_
			COD_FORMANDO&"1, "&SISTEMA_NOVO&", "&COD_PRODUCAO&", '"&COD_CONTRATO&"', '"&COD_CURSO&"', '"&NOME_CURSO&"', '"&EMAIL_FORMANDO&"', '"&CPF&"','"&NOME_FORMANDO&"', "&ANO_CONCLUSAO&", '"&_
			SEMESTRE_CONCLUSAO&"', '"&ALBUM&"', '"&INICIAL&"', '"&RETIRADO&"', '"&BRINDE&"', '"&COMISSAO&"', "&PRODUTO_AGREGADO&_
			")"
			
			'Response.Write(sql&"<br>"&"<br>")
			objConexao.Execute(sql)
			
	sql=	"INSERT INTO TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO("&_
			"COD_FORMANDO ,"&_
			"ENDERECO ,"&_
			"CEP ,"&_
			"FONE_RES ,"&_
			"FONE_CEL ,"&_
			"BAIRRO ,"&_
			"CIDADE ,"&_
			"ESTADO ,"&_
			"CADASTRADO_POR "&_
			")"&_
			"VALUES ("&_
			COD_FORMANDO&"1, '"&ENDERECO&"', '"&CEP&"', '"&FONE_RES&"', '"&FONE_CEL&"', '"&BAIRRO&"', '"&CIDADE&"', '"&ESTADO&"', "&Session("COD_USUARIO")&_
			")"
			
			'Response.Write(sql&"<br>"&"<br>")
			objConexao.Execute(sql)
			
FecharConexao()
			If SISTEMA_NOVO = 2 Then
				Response.Redirect("dip_formandos_cadastro_album_sistema_antigo.asp?cod_formando="&COD_FORMANDO&"1")
			Elseif COD_PRODUCAO = 0 Then
				Response.Redirect("dip_formandos_cadastro_album.asp?cod_formando="&COD_FORMANDO&"1")
			Elseif COD_PRODUCAO > 0 Then
				Response.Redirect("dip_producoes_fichas.asp?cod_prod="&COD_PRODUCAO&"&prod_agre="&PRODUTO_AGREGADO&"&msg=Formando incluído com sucesso!")
			End If
			
End If
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
	d = document.CAD_FORMANDO
	if (d.NOME_FORMANDO.value == ""){
		alert('Campo Nome do Formando em branco!')
		d.NOME_FORMANDO.focus();
		return false;
	}
	if (d.NOME_FORMANDO.value == ""){
			alert('Campo Nome do Formando em branco!')
			d.NOME_FORMANDO.focus();
			return false;
		}
		if (d.COD_CONTRATO.value == ""){
			alert('Campo Contrato em branco!')
			d.COD_CONTRATO.focus();
			return false;
		}
	if (d.COD_CURSO.value == ""){
			alert('Campo Curso em branco!')
			d.COD_CURSO.focus();
			return false;
	}<%If COD_PRODUCAO = 0 or COD_PRODUCAO  = Empty Then%>
	if (d.ALBUM.value == ""){
			alert('Campo Nome do Formando em branco!')
			d.ALBUM.focus();
			return false;
		}	
	if (d.INICIAL.value == ""){
			alert('Campo Inicial em branco!')
			d.INICIAL.focus();
			return false;
		}
		<%End IF%>
	}
	</script>
	<script language="javascript">
		function popUp(campo,URL,campoSelect,msgErro){
			
			if(document.CAD_FORMANDO[campo].value==1){
				
				if (campoSelect !== 'no'){
					//Brinde = "Sim"
					document.CAD_FORMANDO[campoSelect].options[1].selected = "true";
				}
			
				win=window.open(URL,"","width=350,height=220,scrollbars=yes")
				var text = "Se a janela não abrir desbloquei os Pop-ups!";
				if(win == null) { alert(text); return; }
				win.moveTo(350,150);
					
			}
			else {
				if (msgErro==1){alert('Para visualizar os brindes é necessário alterar a opção Brinde para Sim.')}
			}
		
		}
		
		
		function fecharPopUp(){
			if(win){
				win.close();
			}
		}
		
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
</head>

<body>

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

	  <h2>DIP - Formandos Cadastro Completo </h2>
	   <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	  
		  <h3 align="center">Cadastro de Usu&aacute;rio </h3>
		  <form method="post" name="CAD_FORMANDO" onsubmit="return validaForm()" action="dip_formandos_cadastro_completo.asp?comando=incluir2&cod_formando=<%=COD_FORMANDO%>&cod_prod=<%=COD_PRODUCAO%>">
		  <table width="100%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
			<%If SISTEMA_NOVO_QUESTIONAR = 1 Then%>
		 		<tr>
            <td width="20%"><div align="left"><strong>Cadastro para o sistema novo:</strong></div></td>
            <td width="80%"><select name="SISTEMA_NOVO">
  <option selected="selected" value="1">SIM</option>
    <option value="2">NÃO</option>
  </select></td>
            </tr>
			<%End If%> 
<tr>
            <td width="20%"><div align="left"><strong>*Nome Formando:</strong></div></td>
            <td width="80%"><input class="inputTamanho2" type="text" name="NOME_FORMANDO" /></td>
</tr>

<tr>
  <td><div align="left"><strong>*C&oacute;d. Contrato:</strong></div></td>
  <td>
  <select name="COD_CONTRATO">
   <option selected="selected">---</option>
  <%
  AbrirConexao()
  Set rsContratos = objConexao.Execute("SELECT COD_CONTRATO, DESCRICAO FROM TB_ENTIDADES WHERE DESCRICAO IS NOT NULL ORDER BY COD_CONTRATO")
  
  While not rsContratos.Eof
  %>
  <option value="<%=rsContratos("COD_CONTRATO")%>"><%=rsContratos("DESCRICAO")%></option>
  <%
  rsContratos.MoveNext()
  Wend
  rsContratos.Close
  Set rsContratos = Nothing
  FecharConexao()
  %>
  </select>  </td>
</tr>
<tr>
  <td><div align="left"><strong>*Nome Curso: </strong></div></td>
  <td>
    <select name="COD_CURSO">
      <option selected="selected">---</option>
      <%
  AbrirConexao()
  Set rsContratos = objConexao.Execute("SELECT DISTINCT CODIGO_CURSO, NOME_CURSO FROM TB_FORMANDOS_IMP WHERE NOME_CURSO IS NOT NULL ORDER BY NOME_CURSO")
  
  While not rsContratos.Eof
  %>
      <option value="<%=rsContratos("CODIGO_CURSO")%>"><%=rsContratos("NOME_CURSO")%></option>
      <%
  rsContratos.MoveNext()
  Wend
  rsContratos.Close
  Set rsContratos = Nothing
  FecharConexao()
  %>
      </select>  </td>
</tr>
<tr>
  <td><div align="left"><strong>Semestre / Ano de Conclus&atilde;o: </strong></div></td>
  <td>
  <!--
  Desabilitado em 06/08/2012 - Agora não é permitido ano/semestre diferente da entidade
  <select name="SEMESTRE_CONCLUSAO">
  <option selected="selected">---</option>
  <option value="J">JUL</option>
    <option value="D">DEZ</option>
  </select>
   &nbsp; /  
    <select name="ANO_CONCLUSAO">
   		<option>---</option>
	<%For i=1995 to 2035%>
		<option value="<%=i%>"><%=i%></option>
	<%Next%>
   </select>
   -->
   Será de acordo com o contrato selecionado</td>
</tr>


<tr>
  <td><div align="left"><strong>Endere&ccedil;o:</strong></div></td>
  <td><input class="inputTamanho2" type="text" name="ENDERECO" /></td>
</tr>
<tr>
  <td><div align="left"><strong>CEP:</strong></div></td>
  <td><input class="inputTamanho2" type="text" name="CEP" /></td>
</tr>
<tr>
    <td align="left">E-mail:</td>
    <td> <input type="text" name="EMAIL_FORMANDO" id="EMAIL_FORMANDO" 
        value="<%=EMAIL_FORMANDO%>" size="50" maxlength="255">
    </td>
</tr>
<tr>
  <td><div align="left"><strong>Fone Res: </strong></div></td>
  <td><input class="inputTamanho2" type="text" name="FONE_RES" /></td>
</tr>
<tr>
  <td><div align="left"><strong>Fone Cel: </strong></div></td>
  <td><input class="inputTamanho2" type="text" name="FONE_CEL" /></td>
</tr>
<tr>
  <td><div align="left"><strong>Bairro:</strong></div></td>
  <td><input class="inputTamanho2" type="text" name="BAIRRO" /></td>
</tr>
<tr>
  <td><div align="left"><strong>Cidade:</strong></div></td>
 	 <td>
	 <input class="inputTamanho2" type="text" name="CIDADE" />	</td>
</tr>
<tr>
  <td><div align="left"><strong>Estado:</strong></div></td>
  <td><select id="ESTADO" name="ESTADO">
              <option value="AC">Acre</option>
              <option value="AL">Alagoas</option>
              <option value="AP">Amap&aacute;</option>
              <option value="AM">Amazonas</option>
              <option value="BA">Bahia</option>
              <option value="CE">Cear&aacute;</option>
              <option value="DF">Distrito Federal</option>
              <option value="ES">Esp&iacute;rito Santo</option>
              <option value="GO">Goi&aacute;s</option>
              <option value="MA">Maranh&atilde;o</option>
              <option value="MT">Mato Grosso</option>
              <option value="MS">Mato Grosso do Sul</option>
              <option value="MG">Minas Gerais</option>
              <option value="PB">Para&iacute;ba</option>
              <option value="PR">Paran&aacute;</option>
              <option value="PA">Par&aacute;</option>
              <option value="PE">Pernambuco</option>
              <option value="PI">Piau&iacute;</option>
              <option value="RJ">Rio de Janeiro</option>
              <option value="RN">Rio Grande do Norte</option>
              <option value="RS">Rio Grande do   Sul</option>
              <option value="RO">Rond&ocirc;nia</option>
              <option value="RR">Roraima</option>
              <option value="SC">Santa Catarina</option>
              <option  selected="selected" value="SP">S&atilde;o Paulo</option>
              <option value="SE">Sergipe</option>
              <option value="TO">Tocantins</option>
          </select></td>
</tr>
<tr>
            <td width="20%"><div align="left"><strong>CPF:</strong></div></td>
            <td width="80%"><input class="inputTamanho2" type="text" name="CPF" /></td>
</tr>
<tr>
  <td>&nbsp; Os que estão com (*) são campos obrigatórios, o restante é opcional!⚠️ </td>
  <td><input type="submit" class="btnGravar" name="btnGravar" value=" " /></td>
</tr>
        </table>
		</form>
          <h3 align="center">&nbsp;</h3>
		  <p align="center">&nbsp;</p>
		  <p align="center">&nbsp;</p>		
		</td>

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