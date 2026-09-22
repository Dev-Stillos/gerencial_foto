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
	--></head>

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

	  <h2>ADM - M&oacute;dulos / P&aacute;ginas Habilitadas / Cadastrar P&aacute;ginas</h2>
      <br />
<div class="subMenu">Sub-menu: <a href="adm_restrito_paginas.asp">Listar Páginas</a> | <a href="adm_modulos_paginas_habilitadas_cadastrar.asp">Cadastrar Páginas</a></div>
<br />
	   <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	<%
Dim MODULO, NOME_PAGINA, CAMINHO_PAGINA, COD_MODULO
Set fsobj = Server.CreateObject("Scripting.FileSystemObject")
'Cria o objeto FSO que será utilizado para a operação.

pasta = "e:\home\stillos\web\gerencial_foto\"
'Seta a variavel pasta com o caminho para a pasta que desejamos listar seu conteúdo.

Set objpasta = fsobj.GetFolder(pasta)
'Cria o objeto objpasta para manipularmos o conteúdo da referida pasta.

response.write "<b>Listando conteúdo de: "& pasta &"</b><br><br>"



'A rotina abaixo lista todos os arquivos da pasta e exibe os detalhes do mesmo.
AbrirConexao()
Dim rs
contar = 0
For Each objFile in objpasta.Files


CAMINHO_PAGINA = objFile.Path
CAMINHO_PAGINA = Replace(CAMINHO_PAGINA,"E:\home\stillos\web\gerencial_foto\","")

MODULO = Ucase(Left(CAMINHO_PAGINA,3))

NOME_PAGINA = CAMINHO_PAGINA
NOME_PAGINA = Replace(NOME_PAGINA,Lcase(MODULO)&"_","")
NOME_PAGINA = Replace(NOME_PAGINA,"_"," ")
NOME_PAGINA = Replace(NOME_PAGINA,".asp","")
NOME_PAGINA = Ucase(NOME_PAGINA)

'1) Verifica se é alguma página dos módulos cadastrados
If MODULO = "ADM" or MODULO = "DIP" or MODULO = "EPV" or MODULO = "LCA" or MODULO = "CAP" or MODULO = "REP" or MODULO = "PPV" or MODULO = "EST" or MODULO = "PLE" or MODULO = "MCK" or MODULO = "COB" Then
	contar = contar + 1
	'2) Verifica se a página já está cadastrada
	Set rs = objConexao.Execute("SELECT * FROM TB_ADM_PAGINAS_HABILITADAS WHERE CAMINHO_VALIDA = "&sqlTexto(CAMINHO_PAGINA))
	If not rs.Eof Then
		'Pág já cadastrada
		Response.Write contar & ". Mód.: "& MODULO & " | Pág.: "& CAMINHO_PAGINA & "Nome Página:"&NOME_PAGINA &" já cadastrada!<br>"
	Else
		'Pág. não cadastrada, cadastra agora
		If MODULO = "ADM" Then
			COD_MODULO = 5
		Elseif MODULO = "DIP" Then
			COD_MODULO = 1
		Elseif MODULO = "EPV" Then
			COD_MODULO = 3	
		Elseif MODULO = "LCA" Then
			COD_MODULO = 2			
		Elseif MODULO = "CAP" Then
			COD_MODULO = 4
		Elseif MODULO = "REP" Then
			COD_MODULO = 6
		Elseif MODULO = "PPV" Then
			COD_MODULO = 7		
        Elseif MODULO = "EST" Then
			COD_MODULO = 8
		Elseif MODULO = "PLE" Then
			COD_MODULO = 9
		Elseif MODULO = "MCK" Then
			COD_MODULO = 10
		Elseif MODULO = "COB" Then
			COD_MODULO = 12
		End If
			
		objConexao.Execute("INSERT INTO TB_ADM_PAGINAS_HABILITADAS (NOME_PAGINA ,CAMINHO_VALIDA ,COD_MODULO) VALUES ('"&NOME_PAGINA&"', '"&CAMINHO_PAGINA&"', '"&COD_MODULO&"')")
			
		Response.Write contar & ". Mód.: "& MODULO & " | Pág.: "& CAMINHO_PAGINA & "Nome Página:"&NOME_PAGINA &" cadastrada agora!<br>"
	End If 
End If

Next

FecharConexao()
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