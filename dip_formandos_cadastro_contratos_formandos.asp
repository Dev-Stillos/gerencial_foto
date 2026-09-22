<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
	<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>
	<script language="javascript">
		function confirmaExclucaoPerfil(cod_perfil,nome_perfil){
			var msg = "Tem certeza que deseja excluir o perfil "+nome_perfil+"?\nAtenção: Se houver usuários utilizando este perfil o mesmo não será excluído!";
			if(YNconfirm(msg)){
				document.location = "adm_perfis.asp?comando=excluir&cod_perfil="+cod_perfil;
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
	
    <script language="javascript">
/*Exemplo
if(YNconfirm(msg)){

}
*/
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

	  <h2>DIP - Contrato somente Foto</h2>
        <div class="subMenu">Sub-Menu: <a href="dip_formandos_cadastro_contratos_formandos_formulario.asp?teste">Cadastrar novo contrato somente foto</a> | <a href="dip_formandos.asp">Voltar para Cadastro de Formandos</a></div>
        <br />
       
	   <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
       	  	<%
	   Response.Write("<br />")
	Dim sucesso, codmsg
	codmsg	= CStr(""&Request.QueryString("codmsg"))
	sucesso = CStr(""&Request.QueryString("sucesso"))
		
	If sucesso = "" and codmsg <> "" Then%>
    	<div align="center" class="msgGrandeSucesso">
        	<%
			If codmsg = "1" Then
				Response.Write("Contrato somente foto incluído com sucesso!")
			Elseif codmsg = "2" Then
				Response.Write("Contrato somente foto alterado com sucesso!")
			Elseif codmsg = "3" Then	
				Response.Write("Contrato somente foto excluído com sucesso!")
			End If	
			%>
            
        
        </div>
    <%ElseIf sucesso <> "" and codmsg <> "" Then%>
        <div align="center" class="msgGrandeErro">
			<%
			If codmsg = "4" Then
				Response.Write("Contrato somente foto não alterado para não haver duplicidade de informação!<BR>O conteúdo do campo """&Request.QueryString("chaveUnica")&""" já possui uma informação igual em outro Contrato somente foto.<br />")
			Elseif codmsg = "5" Then
				Response.Write("Contrato somente foto não excluído, pois há formandos que dependem deste Contrato.<BR>Se não desejar ter mais este Contrato exclua primeiro todos os formandos relacionados e depois exclua o formando.<br />")
			
			End If	
			%>
</div>
	<%End If%>
            
              <%
		  AbrirConexao()
		  Dim rs, rs2, rs3, sql
		 
		  sql = "SELECT * FROM TB_ENTIDADES WHERE SOMENTE_FOTO = 1 ORDER BY COD_CONTRATO"
	  
		  Set rs = objConexao.Execute(sql)
		  
		  		If not rs.Eof Then
		  %>
       
       			<table width="90%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
   		 <td width="13%"><div align="center"><strong>Contrato</strong></div></td>
  		 <td width="45%"><div align="center"><strong>Descri&ccedil;&atilde;o</strong></div></td>
  		 <td width="20%"><div align="center"><strong>Semestre</strong></div></td>
  		 <td width="14%"><div align="center"><strong>Ano</strong></div></td>
  		 <td width="8%">&nbsp;</td>
  		</tr>
		  <h3 align="center">Contratos somente foto cadastrados</h3>
          <%While not rs.Eof %>

  		<tr>
  		  <td><div align="center"><strong><%=rs("cod_contrato")%></strong></div></td>
  		  <td><div align="left"><%=rs("DESCRICAO")%></div></td>
  		  <td>
		    <div align="center">
		      <%
			If Cstr(rs("SEMESTRE")) = "J" Then  
				Response.Write("Julho")
			Elseif Cstr(rs("SEMESTRE")) = "D" Then
				Response.Write("Dezembro")
			End If
		  %>
		        </div></td>
  		  <td><div align="center"><%=rs("ano")%></div></td>
  		  <td><div align="center"><a href="dip_formandos_cadastro_contratos_formandos_formulario.asp?COD_CONTRATO=<%=rs("COD_CONTRATO")%>&comando=alterar"><img alt="Alterar Cart&atilde;o" src="images/icons/icon_alterar.png" border="0" /></a>&nbsp;&nbsp;&nbsp;<a href="dip_formandos_cadastro_contratos_formandos_formulario.asp?COD_CONTRATO=<%=rs("COD_CONTRATO")%>&comando=excluir"><img alt="Alterar Cart&atilde;o" src="images/icons/icon_excluir.png" border="0" /> </a></div></td>
		  </tr>
     

        <%
		 rs.MoveNext()
		 Wend
		 %>

		
        <%Else%>
        
        
       <div align="center"><em><strong> Não há nenhum contrato somente foto cadastrado no momento</strong></em></div>
        
			

        <%
		End if
		
		rs.Close
		Set rs = Nothing
		
		FecharConexao()
		%>
        		</table>     
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