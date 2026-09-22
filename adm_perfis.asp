<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
If Request.QueryString("comando") = "excluir" Then

	call verificaAntesExcluir(Request.QueryString("cod_perfil"),"TB_ADM_USUARIOS","cod_perfil","adm_perfis.asp?msg=Não foi possível excluir o perfil, devido usuário(s) do sistema estarem utilizando o mesmo.")
	
	AbrirConexao()
		
		objConexao.Execute("DELETE FROM TB_ADM_PERFIS WHERE COD_PERFIL = "&Request.QueryString("cod_perfil"))
		objConexao.Execute("DELETE FROM TB_ADM_PERFIS_PAGINAS_HABILITADAS WHERE COD_PERFIL = "&Request.QueryString("cod_perfil"))
		objConexao.Execute("DELETE FROM TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES WHERE COD_PERFIL = "&Request.QueryString("cod_perfil"))
		
	FecharConexao()
	
	Response.Redirect("adm_perfis.asp?msg=Perfil excluído com sucesso!")
	
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

	  <h2>ADM - M&oacute;dulos / P&aacute;ginas Habilitadas</h2>
	   
        <!-- #include file="includes/adm_usuarios_sub_menu.asp" -->
        <br />
       
	   <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
       		<table width="90%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
   		 <td width="93%"><div align="center"><strong>Perfil</strong></div></td>
  		 <td width="7%">&nbsp;</td>
  		</tr>
	  	  <%
		  AbrirConexao()
		  Dim rs, rs2, rs3, sql
		 
		  sql = "SELECT * FROM TB_ADM_PERFIS ORDER BY NOME_PERFIL"
	  
		  Set rs = objConexao.Execute(sql)
		  %>
		  <h3 align="center">Perfis Cadastrados</h3>
          <%While not rs.Eof %>

  		<tr>
  		  <td><strong><%=rs("NOME_PERFIL")%></strong></td>
  		  <td><div align="center"><a href="adm_perfis_formulario.asp?cod_perfil=<%=rs("COD_PERFIL")%>&amp;nome_perf=<%=rs("NOME_PERFIL")%>&comando=alterar"><img alt="Alterar Cart&atilde;o" src="images/icons/icon_alterar.png" border="0" /></a>&nbsp;&nbsp;&nbsp;<a href="javascript:confirmaExclucaoPerfil(<%=rs("COD_PERFIL")%>,'<%=rs("NOME_PERFIL")%>');"><img alt="Alterar Cart&atilde;o" src="images/icons/icon_excluir.png" border="0" /> </a></div></td>
		  </tr>
     

        <%
		 rs.MoveNext()
		 Wend
		  rs.Close
		  Set rs = Nothing
		 %>

        <%
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