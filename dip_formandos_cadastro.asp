<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<%
'Para tirar a validação da página basta tirar o "#" ou coloca-lo para habilitar antes do include
Dim CODIGO_FORMANDO, CADASTRO, FINALIZADO_CADASTRO
CODIGO_FORMANDO = Request.QueryString("cod_formando")
CADASTRO		= Request.QueryString("cadastro")
FINALIZADO_CADASTRO=2

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

	  <h2>DIP - Formandos Cadastro Individual </h2>
	   <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	  <%
		  If CODIGO_FORMANDO <> Empty and CADASTRO = "importado" Then
		  
		  	AbrirConexao()
		  	
		 	Dim rs, rs2, rs3, sql
	 	
				sql = "SELECT * FROM TB_FORMANDOS_IMP WHERE CODIGO_FORMANDO ="&CODIGO_FORMANDO
	
		  Set rs = objConexao.Execute(sql)
		  
		  If not rs.Eof then
		  		
					'Verifica se o formando realmente não está cadastrado
					Dim rsVerificaImportado
					Set rsVerificaImportado = objConexao.Execute("SELECT COD_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO="&CODIGO_FORMANDO&"0 ")
				
					If rsVerificaImportado.Eof Then
					
					'Importa o formando
					Dim COD_FORMANDO, COD_CONTRATO, COD_CURSO, NOME_CURSO, NOME_FORMANDO, ANO_CONCLUSAO, SEMESTRE_CONCLUSAO, TURMA
					'PARTE2
					Dim ENDERECO, CEP, FONE_RES, FONE_CEL, BAIRRO, CIDADE, ESTADO
					
					'Tb 1
					'0 devido a importação da formatura
					COD_FORMANDO		= rs("CODIGO_FORMANDO")&"0"
					TURMA 				= rs("TURMA")
					COD_CONTRATO		= rs("CONTRATO")
					COD_CURSO			= rs("CODIGO_CURSO")
					NOME_CURSO			= rs("NOME_CURSO")
					NOME_FORMANDO		= ""&rs("NOME_FORMANDO")
					'Correção do nome do formando
					posicaoInicial = inStr(NOME_FORMANDO,"(")
					If posicaoInicial > 0 Then
						NOME_FORMANDO = Left(NOME_FORMANDO,posicaoInicial-1)
					End If
	
					ANO_CONCLUSAO		= rs("ANO")
					SEMESTRE_CONCLUSAO	= rs("SEMESTRE")
					'------------------------------------
					'Tb2
					ENDERECO			= rs("ENDERECO")
					CEP					= rs("CEP")
					CEP					= Replace(CEP,"-","")
					FONE_RES			= rs("FONE-RES")
					FONE_CEL			= rs("FONE-COM-CEL")
					BAIRRO				= rs("BAIRRO")
					CIDADE				= rs("MUNICIPIO")
					ESTADO				= rs("ESTADO")
							
					sql = 	"INSERT INTO TB_FICHAS_IDENTIFICACAO ("&_
								"COD_FORMANDO , "&_
								"COD_PRODUCAO , "&_
								"COD_CONTRATO , "&_
								"COD_CURSO , "&_
								"NOME_CURSO , "&_
								"NOME_FORMANDO , "&_
								"ANO_CONCLUSAO , "&_
								"SEMESTRE_CONCLUSAO "&_
								")"&_
								"VALUES ("&_
								sqltexto(COD_FORMANDO)&", "&_
								"0, "&_
								COD_CONTRATO&", "&_
								COD_CURSO&", "&_
								sqltexto(NOME_CURSO)&", "&_
								sqltexto(NOME_FORMANDO)&", "&_
								sqltexto(ANO_CONCLUSAO)&", "&_
								sqltexto(Ucase(SEMESTRE_CONCLUSAO))&_							
								")"			
								'Tb 1
								objConexao.Execute(sql)
								
								Response.Write(sql&"<br />")
								
								sql = 	"INSERT INTO TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO ("&_
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
								COD_FORMANDO&", "&sqlTexto(ENDERECO)&", '"&CEP&"', '"&FONE_RES&"', '"&FONE_CEL&"', "&sqlTexto(BAIRRO)&", "&sqlTexto(CIDADE)&", "&_
								sqlTexto(ESTADO)&", '9999'"&_
								")"
								
								'Tb 2
								objConexao.Execute(sql)
								
								FINALIZADO_CADASTRO=1
								
								
					
					End If
	 					
				
		  		
				Else
					Response.Redirect("dip_formandos.asp?msg=Desculpe não foi possível importar o formando, pois consta como já cadastrado cod.:"&CODIGO_FORMANDO)
				End IF
				
				%>
		  <%
  		  rs.Close
		  Set rs=Nothing

		  
		  FecharConexao()
		  		  
		  End If
		  
'////////////////////////////////////////
		  If FINALIZADO_CADASTRO = 1 Then
		  	Response.Redirect("dip_formandos_cadastro_album.asp?cod_formando="&CODIGO_FORMANDO&"0")
		  End If

		  %>
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



