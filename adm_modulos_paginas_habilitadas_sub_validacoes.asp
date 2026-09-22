<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Páginas especiais somente para o administrador master do sistema
'//////////////
If Session("cod_usuario") <> 1 Then
	Session.Abandon()
	Response.Redirect("index.asp")
End If
'//////////////
%>
<%
Dim COD_PAGINA, NOME_PAGINA, comando
Dim NOME_SUB_VALIDACAO, COD_SUB_VALIDACAO, COD_SUB_VALIDACAO_RS
COD_PAGINA			= Request.QueryString("cod_pagina")
NOME_PAGINA			= Request.QueryString("nome_pagina")
comando				= Request.QueryString("comando")
COD_SUB_VALIDACAO	= formatNumber(0&Request.QueryString("cod_sub_validacao"),0)

If comando="incluir" Then
	AbrirConexao()
	
	NOME_SUB_VALIDACAO = Request.Form("NOME_SUB_VALIDACAO")
	
	objConexao.Execute("INSERT INTO TB_ADM_PAGINAS_HABILITADAS_SUB_VALIDACOES (COD_PAGINA,NOME_SUB_VALIDACAO) VALUES ("&COD_PAGINA&", "&sqlTexto(NOME_SUB_VALIDACAO)&")")
	FecharConexao()
	
	Response.Redirect("adm_modulos_paginas_habilitadas_sub_validacoes.asp?msg=Sub-validação incluída com sucesso!&cod_pagina="&COD_PAGINA&"&nome_pagina="&NOME_PAGINA)
	
End If

If comando = "excluir" Then
	AbrirConexao()
	
	COD_SUB_VALIDACAO = Request.QueryString("cod_sub_validacao")
	
		objConexao.Execute("DELETE FROM TB_ADM_PAGINAS_HABILITADAS_SUB_VALIDACOES WHERE COD_SUB_VALIDACAO = "&COD_SUB_VALIDACAO&" LIMIT 1")
		
	FecharConexao()
	
	
End If

If comando = "alterar" and Request.Form("btnGravar") <> Empty Then
	AbrirConexao()
	NOME_SUB_VALIDACAO = Request.Form("NOME_SUB_VALIDACAO")
	
		objConexao.Execute("UPDATE TB_ADM_PAGINAS_HABILITADAS_SUB_VALIDACOES SET NOME_SUB_VALIDACAO = "&sqlTexto(NOME_SUB_VALIDACAO)&" WHERE COD_SUB_VALIDACAO = "&Request.Form("COD_SUB_VALIDACAO_RS")&" LIMIT 1")
	FecharConexao()
	Response.Redirect("adm_modulos_paginas_habilitadas_sub_validacoes.asp?msg=Sub-validação alterada com sucesso!&cod_pagina="&COD_PAGINA&"&nome_pagina="&NOME_PAGINA)
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Versão 1.2</title>
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
	<!--AJUDA F1 E CTRL PARA BUSCA RÁPIDA-->
		
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
		<!-- CALENDÁRIO SCRIPTS
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
		if(document.FORMULARIO.NOME_SUB_VALIDACAO.value == ""){
			document.FORMULARIO.NOME_SUB_VALIDACAO.focus();
			alert('Informe o nome da nova sub-validação!');
			return false;
		}
	}
	function confirmaExclusao(ID){
	var pergunta = confirm('Tem certeza que deseja excluir esta sub-validação?')
	var caminho	 = "adm_modulos_paginas_habilitadas_sub_validacoes.asp?comando=excluir&cod_pagina=<%=COD_PAGINA%>&nome_pagina=<%=NOME_PAGINA%>&cod_sub_validacao="+ID
			if (pergunta==true){
				abre(caminho)
			}
	}
	</script>
</head>

<body onKeyDown="javascript:Verificar()">

<table cellpadding="9px" cellspacing="0" width="100%">
<tr>
<td>
<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td>
		<!-- #include file="includes/topo.asp" -->
<table width="100%" style="border-bottom:#999999 solid dotted 1px" background="images/layout/fundo_menu_azul.jpg"border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><!-- #include file="includes/menu.asp" --></td>
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

	  <h2>ADM - Módulos / Páginas Habilitadas / Sub-Validações</h2>
      <div class="subMenu">Sub-menu: <a href="adm_restrito_paginas.asp">Listar Páginas</a> | <a href="adm_modulos_paginas_habilitadas_cadastrar.asp">Cadastrar Páginas</a> | <a href="?cod_pagina=<%=COD_PAGINA%>&nome_pagina=<%=NOME_PAGINA%>&comando=novo">Incluir sub-validação para está página</a></div>

	  <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	  <h3 align="center">Sub-Validações da Página "<%=NOME_PAGINA%>"</h3>
		  <%If comando <> "novo" Then%>
		            <%
		   AbrirConexao()
		   Dim rsSubValidacoes
			Set rsSubValidacoes = objConexao.Execute("SELECT * FROM TB_ADM_PAGINAS_HABILITADAS_SUB_VALIDACOES WHERE COD_PAGINA="&COD_PAGINA,qtdSubValidacoes)
				If not rsSubValidacoes.Eof Then%>
          <form name="FORMULARIO" action="?comando=alterar&cod_pagina=<%=COD_PAGINA%>&nome_pagina=<%=NOME_PAGINA%>" method="post" style="margin:0px" onsubmit="return validaForm()">
          <table width="70%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="21%"><div align="center"><strong>Cód Sub-Validação </strong></div></td>
            <td width="68%"><div align="center"><strong>Nome Sub-Validação</strong></div></td>
            <td width="11%"><div align="center"><strong>Alterar</strong></div></td>
</tr>
		 <%
				While not rsSubValidacoes.Eof
				COD_SUB_VALIDACAO_RS = formatNumber(0&rsSubValidacoes("COD_SUB_VALIDACAO"),0)
				%>
          <tr>
            <td>
			  <div align="center"><%=rsSubValidacoes("COD_SUB_VALIDACAO")%>
			      </div></td>
            <td><%If comando = "alterar" and COD_SUB_VALIDACAO_RS = COD_SUB_VALIDACAO Then%>
				<div align="center">
					<input class="inputTamanho3" value="<%=rsSubValidacoes("NOME_SUB_VALIDACAO")%>" type="text" name="NOME_SUB_VALIDACAO" />
				</div>
			<%Else%>
				<div align="left">
					<%=rsSubValidacoes("NOME_SUB_VALIDACAO")%>				</div>
			<%End If%>			 </td>
            <td><div align="center">
			<%If comando = "alterar" and COD_SUB_VALIDACAO_RS = COD_SUB_VALIDACAO Then%>
				<input class="btnGravar" type="submit" name="btnGravar" value=" " />
				<input type="hidden" name="COD_SUB_VALIDACAO_RS" value="<%=COD_SUB_VALIDACAO_RS%>" />
			<%Else%>
			<a href="?cod_pagina=<%=COD_PAGINA%>&nome_pagina=<%=NOME_PAGINA%>&comando=alterar&cod_sub_validacao=<%=rsSubValidacoes("COD_SUB_VALIDACAO")%>"><img src="images/icons/icon_alterar.png" width="16" height="16" border="0" /></a>	
				   <a href="javascript:confirmaExclusao(<%=rsSubValidacoes("COD_SUB_VALIDACAO")%>)"><img src="images/icons/icon_excluir.png" width="16" height="16" border="0" /></a>
			<%End If%>
			</div></td>
          </tr>
          		  <%
				rsSubValidacoes.MoveNext
				Wend%>
        </table>   
          </form>
            <%Else
			Response.Redirect("adm_modulos_paginas_habilitadas_sub_validacoes.asp?comando=novo&cod_pagina="&COD_PAGINA&"&nome_pagina="&NOME_PAGINA)
		End If
		rsSubValidacoes.Close
		Set rsSubValidacoes = Nothing
		FecharConexao()
		%>
              <%Elseif comando = "novo" Then%>
            <div align="center">
			<form name="FORMULARIO" action="?comando=incluir&cod_pagina=<%=COD_PAGINA%>&nome_pagina=<%=NOME_PAGINA%>" method="post" style="margin:0px" onsubmit="return validaForm()">
				             <table width="70%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
                <td><div align="center"><strong>Incluir nova Sub-Validação</strong> </div></td>
              </tr>
              <tr>
                <td><div align="center">
                  <input class="inputTamanho3" type="text" name="NOME_SUB_VALIDACAO" />
                </div></td>
              </tr>
              <tr>
                <td><div align="center">
                  <input class="btnGravar" type="submit" name="btnGravar" value=" " />
                </div></td>
              </tr>
            </table>
			</form>
                      </div>
			<%End If%>

          <p align="center"> </p>
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



