<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<%
Dim BUSCAR_TERMO, TIPO_BUSCA
BUSCAR_TERMO = TrocarAcento(""&Request.QueryString("BUSCAR_TERMO"))
TIPO_BUSCA		= formatNumber(0&Request.QueryString("TIPO_BUSCA"),0)

If Request.QueryString("limpar_agendadas") = "sim" Then

AbrirConexao()
 	objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_ETIQUETAS_TEMPORARIA WHERE COD_USUARIO="&Session("COD_USUARIO")&" AND AGENDADA=1")
FecharConexao()
	Response.Redirect("dip_formandos.asp?msg=Etiquetas agendadas excluídas do agendamento com sucesso!")
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
	<script language="javascript">
	function validaForm(cod_formando,nomeFormando){
	d = document.CAD_FORMANDOS
	var acao
	acao = "dip_formandos_cadastro.asp?cod_formando="+cod_formando+"&cadastro=importado"
		var valida = 'Tem certeza que deseja cadastrar o formando(a) '+nomeFormando
		
			if(YNconfirm(valida)){
				d.action = acao;
				d.submit();	
			}
	}
	
	function confirmarExclusaoAgendamento(){
		if(YNconfirm('Tem certeza que deseja limpar as etiquetas agendadas pra impressão\nCertifique-se que foram impressão com sucesso antes de continuar.')){
			abre('?limpar_agendadas=sim')
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

	  <h2>DIP - Formandos - Fichas com erro </h2>
	   <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	  <h3 align="center">Relação de Fichas com Erro </h3>
		  <div id="comboTexto">
		  	<form name="CAD_FORMANDOS" method="post"> 
          	<table width="95%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="11%"><div align="center"><strong>Contrato</strong></div></td>
            <td width="31%"><div align="center"><strong>Instituição de Ensino </strong></div></td>
            <td width="15%"><div align="center"><strong>Formando</strong></div></td>
            <td width="16%"><div align="center"><strong>Álbum</strong></div></td>
            <td width="14%"><div align="center"><strong>Curso </strong></div></td>
            </tr>
		  <%
		  AbrirConexao()
		  Set rs = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE ALBUM = 0 AND COD_PRODUCAO = 0",FICHAS_ERROS)
		  
		  While not rs.Eof
		    
		
	  %>
          <tr>
            <td><div align="center"><%=rs("COD_CONTRATO")%></div></td>
            <td><%
			Set rsTurma = objConexao.Execute("SELECT DESCRICAO FROM TB_ENTIDADES WHERE COD_CONTRATO='"&rs("COD_CONTRATO")&"'")
				
				If not rsTurma.Eof Then
					
					Response.Write(rsTurma("DESCRICAO"))
				
				End If
			rsTurma.Close
			Set rsTurma = Nothing
			%></td>
            <td><%=rs("NOME_FORMANDO")%> (<%=rs("COD_FORMANDO")%>)</td>
            <td><div align="center"><a href="dip_formandos_cadastro_album.asp?cod_formando=<%=rs("COD_FORMANDO")%>"><strong>Corrigir album</strong></a></div></td>
            <td><%=rs("NOME_CURSO")%></td>
            </tr>
		  <%
		  rs.MoveNext()
		  Wend
	      
		  rs.Close()
		  Set rs = Nothing
		  FecharConexao()
		  %>
        </table>
		  	</form>
		  </div>
		
		  <p align="center">
		  <input class="btnTopo" onclick="abre('#')" type="button" name="btnTopo2" value=" " />
		</p>		</td>
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



