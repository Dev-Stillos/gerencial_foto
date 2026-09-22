<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim COD_FORMANDO
Dim msg
Dim comando


COD_FORMANDO = Request.QueryString("cod_formando")
msg = Request.QueryString("msg")
comando = Request.QueryString("comando")

AbrirConexao()
		'PARTE1
		Dim COD_CONTRATO, SISTEMA_NOVO, NOME_CURSO, NOME_FORMANDO, ANO_CONCLUSAO, SEMESTRE_CONCLUSAO, ALBUM, INICIAL, BRINDE, COMISSAO, PRODUTO_AGREGADO, VENDIDO, ENCADERNACOES_EXTRAS, COD_PRODUCAO,	CREDITO_EM_VALOR, CREDITO_MOTIVO, CREDITO_DATA_INSERIDO, COD_CURSO
		'PARTE2
		Dim ENDERECO, CEP, FONE_RES, FONE_CEL, BAIRRO, CIDADE, ESTADO
				
		sql = "SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO="&COD_FORMANDO
		
		Set rs=objConexao.Execute(sql)
		
		If not rs.Eof Then
				
				ENCADERNACOES_EXTRAS = formatNumber(0&rs("ENCADERNACOES_EXTRAS"),0)
				COD_CONTRATO		= rs("COD_CONTRATO")
				SISTEMA_NOVO		= formatNumber(0&rs("SISTEMA_NOVO"),0)
				NOME_CURSO			= rs("NOME_CURSO")
				COD_CURSO			= rs("COD_CURSO")
				NOME_FORMANDO		= rs("NOME_FORMANDO")
				ANO_CONCLUSAO		= rs("ANO_CONCLUSAO")
				SEMESTRE_CONCLUSAO	= rs("SEMESTRE_CONCLUSAO")
				ALBUM				= rs("ALBUM")
				INICIAL				= rs("INICIAL")
				BRINDE				= rs("BRINDE")
				COMISSAO			= formatNumber(0&rs("COMISSAO"),0)
				PRODUTO_AGREGADO	= rs("PRODUTO_AGREGADO")
				VENDIDO				= rs("VENDIDO")
				COD_PRODUCAO		= formatNumber(0&rs("COD_PRODUCAO"),0)
				CREDITO_EM_VALOR	= formatNumber(0&rs("CREDITO_EM_VALOR"))
				CREDITO_MOTIVO		= ""&rs("CREDITO_MOTIVO")
				CREDITO_DATA_INSERIDO = ""&rs("CREDITO_DATA_INSERIDO")

				If SEMESTRE_CONCLUSAO = "J" Then
					SEMESTRE_CONCLUSAO = "JULHO"
				Else
					SEMESTRE_CONCLUSAO = "DEZEMBRO"
				End If
							
		End If
				
		
		rs.Close
		Set rs=Nothing
		
		Dim qtdEnderecos
		Set rs2 = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO WHERE COD_FORMANDO="&COD_FORMANDO&" ORDER BY DATA_CADASTRO DESC LIMIT 1",qtdEnderecos)
		
		If not rs2.Eof Then
		
				ENDERECO			= rs2("ENDERECO")
				CEP					= rs2("CEP")
				FONE_RES			= rs2("FONE_RES")
				FONE_CEL			= rs2("FONE_CEL")
				BAIRRO				= rs2("BAIRRO")
				CIDADE				= rs2("CIDADE")
				ESTADO				= rs2("ESTADO")
		End If
		
		rs2.Close
		Set rs2 = Nothing		
		
FecharConexao()


'//////////////
'///////////////

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Vers�o 1.2</title>
<style type="text/css">
	/* CSS for the demo. CSS needed for the scripts are loaded dynamically by the scripts */
	#mainContainer{
		width:500px;
		margin:0 auto;
		margin-top:10px;
		border:1px double #000;
		padding:3px;
		
	}
	
	/* Alternative layout for the message box - used by one of the messages */
	
	.modalDialog_contentDiv_error{
		border:3px solid #FFF;	
		padding:2px;
		z-index:100;/* Always needed	*/
		position:absolute;	/* Always needed	*/
		background-color:#F00;	/* White background color for the message */
		color:#FFF;
	}
	.modalDialog_contentDiv_error a{
		color:#FFF;
	}
	h1{
		margin-top:0px;
	}

	</style>	
<script type="text/javascript">
	function verify(ver){
		
		if(ver){
			// Confirmed message, i.e. clicked on "Yes"
			alert('Message confirmed');
		}else{
			// Clicked on "No"
			alert('Message not confirmed');
		}
	}
	</script>
	<link rel="stylesheet" href="css/modal-message.css" type="text/css">
<script type="text/javascript" src="js/ajax.js"></script>
<script type="text/javascript" src="js/modal-message.js"></script>
<script type="text/javascript" src="js/ajax-dynamic-content.js"></script>
<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>
	
	
	
<script type="text/javascript">
messageObj = new DHTML_modalMessage();	// We only create one object of this class
messageObj.setShadowOffset(5);	// Large shadow


function displayMessage(url)
{
	messageObj.setSource(url);
	messageObj.setCssClassMessageBox(false);
	messageObj.setSize(550,450);
	messageObj.setShadowDivVisible(true);	// Enable shadow for these boxes
	messageObj.display();
}

function displayStaticMessage(messageContent,cssClass,altura)
{
	messageObj.setHtmlContent(messageContent);
	messageObj.setSize(550,altura);
	messageObj.setCssClassMessageBox(cssClass);
	messageObj.setSource(false);	// no html source since we want to use a static message here.
	messageObj.setShadowDivVisible(false);	// Disable shadow for these boxes	
	messageObj.display();
	
	
}

function closeMessage()
{
	messageObj.close();	
}


</script>
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
 


	<!--AJUDA F1 E CTRL PARA BUSCA R�PIDA-->
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
<script>
function confirma() {
	
	if(YNconfirm("Mensagem")){
			document.dados.submit();			
	}	
}
</script>
<script language="javascript">
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
<script>
function confirmaProdAparte(PRODUTO_A_PARTE){

	if(YNconfirm("Tem certeza que deseja excluir este produto?")){
			location.href = "?cod_pedido=<%=COD_PEDIDO%>&prod_excluir_a_parte="+PRODUTO_A_PARTE
	}

}
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  if (selObj.value !== ""){
  	eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  	if (restore) selObj.selectedIndex=0;
  }
}
function popUp(URL){
	win=window.open(URL,"","width=350,height=220,scrollbars=yes")
	var text = "Se a janela n&atilde;o abrir desbloquei os Pop-ups!";
	if(win == null) { alert(text); return; }
	win.moveTo(350,150);
}


function validaForm(){
	d = document.CAD_FORMANDO;
	if(d.NOME_FORMANDO.value == "" || d.COD_CONTRATO.value == "" || d.COD_CURSO.value == "" || d.INICIAL.value == "" || d.COMISSAO.value == "" || d.BRINDE.value == ""){	
		alert('Preencha todos os campos obrigat&oacute;rios!');
		return false;
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
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = " href=""javascript:popUp('lca_busca_detalhamento_pedido_IMPRIMIR.asp?cod_pedido="&COD_PEDIDO&""" )"">"
  %>
  <!-- #include file="includes/sub_menu.asp" -->

  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">

	  <h2>DIP - Corre&ccedil;&atilde;o de Formando</h2>
      <div class="msgExcluir">Aten&ccedil;&atilde;o!<br /> 
      <%=Session("NOME")%>, ao corrigir o formando ser&aacute; gravado no Log de Acesso que voc&ecirc; efetuou altera&ccedil;&atilde;o no cadastro do formando.
      </div>
      <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	  
		  	<h3 align="center">Corre&ccedil;&atilde;o do formando n&ordm; <%=digito(COD_FORMANDO)%></h3>
				<form method="post" name="CAD_FORMANDO" onsubmit="return validaForm()" action="dip_busca_detalhamento_formando_correcao_formando_processa.asp?comando=alterar&cod_formando=<%=COD_FORMANDO%>">
                <fieldset style="padding:10px;">
            <legend>Dados Pessoais</legend><br />
		  <table width="100%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
			<%If SISTEMA_NOVO_QUESTIONAR = 1 Then%>
			<%End If%> 
<tr>
            <td width="21%"><div align="left"><strong>Nome Formando *:</strong></div></td>
            <td width="79%"><input class="inputTamanho3" value="<%=NOME_FORMANDO%>" type="text" name="NOME_FORMANDO" /></td>
            </tr>
<tr>
  <td><div align="left"><strong>C&oacute;d. Contrato *:</strong></div></td>
  <td>
  <select name="COD_CONTRATO">
   <option value="">---</option>
  <%
  AbrirConexao()
  Set rsContratos = objConexao.Execute("SELECT COD_CONTRATO, DESCRICAO FROM TB_ENTIDADES WHERE DESCRICAO IS NOT NULL ORDER BY COD_CONTRATO")
  
  While not rsContratos.Eof
  
  	If Cstr(COD_CONTRATO) = Cstr(rsContratos("COD_CONTRATO")) Then%>
    <option selected="selected" value="<%=rsContratos("COD_CONTRATO")%>"><%=rsContratos("DESCRICAO")%></option>
    <%
	Else
  %>
  <option value="<%=rsContratos("COD_CONTRATO")%>"><%=rsContratos("DESCRICAO")%></option>
  <%
  	End If
  
  rsContratos.MoveNext()
  Wend
  rsContratos.Close
  Set rsContratos = Nothing
  FecharConexao()
  %>
  </select>  </td>
</tr>
<tr>
  <td><div align="left"><strong>Curso *:</strong></div></td>
  <td>
  <select name="COD_CURSO">
   <option value="">---</option>
  <%
  AbrirConexao()
  Set rsContratos = objConexao.Execute("SELECT DISTINCT CODIGO_CURSO, NOME_CURSO FROM TB_FORMANDOS_IMP WHERE NOME_CURSO IS NOT NULL ORDER BY NOME_CURSO")
  
  While not rsContratos.Eof
		If int (COD_CURSO) = int(rsContratos("CODIGO_CURSO")) Then%>
   <option value="<%=rsContratos("CODIGO_CURSO")%>" selected="selected"><%=rsContratos("NOME_CURSO")%></option>     
  <%Else%>      
  <option value="<%=rsContratos("CODIGO_CURSO")%>"><%=rsContratos("NOME_CURSO")%></option>
  <%
  End If
  
  rsContratos.MoveNext()
  Wend
  rsContratos.Close
  Set rsContratos = Nothing
  FecharConexao()
  %>
  </select>  </td>
</tr>
<tr>
  <td><div align="left"><strong>Semestre / Ano de Conclus&atilde;o *:</strong></div></td>
  <td>
  
  <!--
  Desabilitado em 06/08/2012 - Agora n�o � permitido ano/semestre diferente da entidade
  
  <select name="SEMESTRE_CONCLUSAO">
  <option value="">---</option>
  <option <%If SEMESTRE_CONCLUSAO = "JULHO" Then %> selected="selected" <%End If%> value="J">JUL</option>
    <option <%If SEMESTRE_CONCLUSAO = "DEZEMBRO" Then %> selected="selected" <%End If%> value="D">DEZ</option>
  </select>
   � /  
    <select name="ANO_CONCLUSAO">
   		<option value=""></option>
	<%For i=1990 to 2040
		
		if int(ANO_CONCLUSAO) = int(i) Then
	%>
    	<option selected="selected" value="<%=i%>"><%=i%></option>
    <%Else%>
		<option value="<%=i%>"><%=i%></option>
	<%
	End If
	Next%>
   </select>
   -->
   Ser&aacute; de acordo com o contrato selecionado
   </td>
</tr>
        </table>
        </fieldset>
        <br /><br />
            <fieldset style="padding:10px;">
            <legend>Informa&ccedil;&otilde;es do &Aacute;lbum</legend><br />
		  <table width="100%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
			<%If SISTEMA_NOVO_QUESTIONAR = 1 Then%>
			<%End If%>


<tr>
            <td width="20%"><strong>N&uacute;mero do &Aacute;lbum *:</strong></td>
            <td width="80%"><input name="ALBUM" value="<%=ALBUM%>" type="text" size="10" maxlength="10" /></td>
            </tr>
<tr>
            <td width="20%"><strong>Qtd. Inicial *:</strong></td>
            <td width="80%"><input name="INICIAL" value="<%=INICIAL%>" type="text" size="10" maxlength="10" /></td>
            </tr>

           <!--BLOQUEIA ALTERACAO DE FORMANDO: SOMENTE O ADMINISTRADOR PODE EDITAR O TIPO COMISSAO E SE TEM BRINDES-->
           <%If Session("COD_USUARIO") = 1 then %>
          <tr>
            <td><strong>Comiss&atilde;o *:</strong></td>
            <td><select name="COMISSAO">
                <option value="2" <%If int(COMISSAO) = 2 or COMISSAO = Empty Then %> selected="selected" <%End If%>>N&Atilde;O</option>
                <option value="1" <%If int(COMISSAO) = 1 Then %> selected="selected" <%End If%>>SIM</option>
              </select> 
              </td>
            </tr>
          <tr>
            <td><strong>Brinde *:</strong></td>
            <td><select name="BRINDE">
                <option value="2" <%If int(COMISSAO) = 2 or COMISSAO = Empty Then %> selected="selected"<%End If%>>N&Atilde;O</option>
                <option value="1" <%If int(COMISSAO) = 1 Then %>selected="selected"<%End If%>>SIM</option>
              </select> 
              <a href="javascript:popUp('dip_producoes_fichas_POPUP_brindes.asp?cod_form=<%=cod_formando%>&cod_prod=<%=COD_PRODUCAO%>&brinde='+document.CAD_FORMANDO.BRINDE.value)">Listar produtos brindes</a></td>
            </tr>
           <%End If %>
        </table>
        </fieldset>
                <p><br />
                    <!--
			<fieldset style="padding:10px;">
            <legend>Endere�o Atual</legend><br />
		 <table width="100%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
        <tr>
  <td width="20%"><div align="left"><strong>Endere�o:</strong></div></td>
  <td width="80%"><input class="inputTamanho3" value="<%=ENDERECO%>" type="text" name="ENDERECO" /></td>
</tr>
<tr>
  <td><div align="left"><strong>CEP:</strong></div></td>
  <td><input class="inputTamanho3" value="<%=CEP%>" type="text" name="CEP" /></td>
</tr>
<tr>
  <td><div align="left"><strong>Fone Res.: </strong></div></td>
  <td><input class="inputTamanho3" value="<%=FONE_RES%>" type="text" name="FONE_RES" /></td>
</tr>
<tr>
  <td><div align="left"><strong>Fone Cel.: </strong></div></td>
  <td><input class="inputTamanho3" value="<%=FONE_CEL%>" type="text" name="FONE_CEL" /></td>
</tr>
<tr>
  <td><div align="left"><strong>Bairro:</strong></div></td>
  <td><input class="inputTamanho3" value="<%=BAIRRO%>" type="text" name="BAIRRO" /></td>
</tr>
<tr>
  <td><div align="left"><strong>Cidade:</strong></div></td>
 	 <td>
	 <input class="inputTamanho3" value="<%=CIDADE%>" type="text" name="CIDADE" />	</td>
</tr>
<tr>
  <td><div align="left"><strong>Estado:</strong></div></td>
  <td><input name="ESTADO" type="text" class="inputTamanho1" value="<%=ESTADO%>" size="2" maxlength="2" /></td>
</tr>
        </table>
        </fieldset>
        -->
                  
                  
                         <strong>* Campos obrigat&oacute;rios</strong><br />
                  </p>
                <p align="center">
                  <input type="submit" value=" " class="btnGravar" />
                </p>
                <p align="center">                  Obs.: O lote poder&aacute; ser alterado somente antes de finalizar a produ&ccedil;&atilde;o e o para alterar para um novo endere&ccedil;o do formando clica na op&ccedil;&atilde;o Cadastrar novo Endere&ccedil;o na p&aacute;gina anterior. <br />
                                                </p>
                <p align="center"></p>
<p align="center"></p>

  </form>

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



