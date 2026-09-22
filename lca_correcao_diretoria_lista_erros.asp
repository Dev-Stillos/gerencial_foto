<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim comando, COD_PEDIDO_QS, COD_PEDIDO_RS, COD_PEDIDO_FORM
comando			= Request.QueryString("comando")
COD_PEDIDO_QS	= formatNumber(0&Request.QueryString("COD_PEDIDO"),0)


If comando = "alterar" and Request.Form("SALVAR") <> Empty Then
	COD_PEDIDO_FORM	= Request.Form("COD_PEDIDO")
	AbrirConexao()
	objConexao.Execute("UPDATE TB_PEDIDOS SET ERRO_LISTA = "&sqlTexto(Request.Form("ERRO_LISTA"))&" WHERE COD_PEDIDO = "&COD_PEDIDO_FORM&" LIMIT 1")
	Response.Redirect("lca_correcao_diretoria_lista_erros.asp?msg=Erro alterado com sucesso!")
	FecharConexao()
End If

If comando = "deletar" Then
	AbrirConexao()
	objConexao.Execute("UPDATE TB_PEDIDOS SET ERRO_LISTA = NULL WHERE COD_PEDIDO = "&COD_PEDIDO_QS&" LIMIT 1")
	Response.Redirect("lca_correcao_diretoria_lista_erros.asp?msg=Erro excluído com sucesso!")
	FecharConexao()
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
<!--
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

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->

function cadastraVendedor(valor){
	alert(valor);
}
</script>
<script>
function confirma(codigoPedido) {
	
	if(YNconfirm("Tem certeza que deseja excluir este erro?")){
		var urlRedirecionar = "lca_correcao_diretoria_lista_erros.asp?COD_PEDIDO="+codigoPedido+"&comando=deletar"
		abre(urlRedirecionar)		
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

	  <h2>LCA - Corre&ccedil;&atilde;o Diretoria / Lista de Erros</h2>
	   <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
			</p>
	   <p align="center">&nbsp;</p>
       <%
	   AbrirConexao()
	   
	   Set rs = objConexao.Execute("SELECT COD_PEDIDO, ERRO_LISTA FROM TB_PEDIDOS WHERE ERRO_LISTA IS NOT NULL ORDER BY COD_PEDIDO")
	   
	   If not rs.Eof Then
	   

	   %>
       <form method="post" style="margin:0px">
	<table width="80%" 
		      border="1" align="center"
					cellpadding="5" bgcolor="#FFFFFF" 
					style="border-collapse:collapse">
  <tr bgcolor="whitesmoke">  
             <td width="18%"><div align="center"><strong>Cod. Pedido</strong></div></td>
             <td width="64%"><div align="center"><strong>Erro relatado</strong></div></td>
             <td width="9%"><div align="center"><strong>Alterar</strong></div></td>
             <td width="9%"><div align="center"><strong>Excluir da Lista de Erro</strong></div></td>
  </tr>
           <%
		   While not rs.Eof
		   COD_PEDIDO_RS	= formatNumber(0&rs("COD_PEDIDO"),0)
		   %>
           <tr <%If comando = "alterar" and COD_PEDIDO_QS = COD_PEDIDO_RS Then%> bgcolor="#FFFF99"<%End If%>>
             <td>
			 <%=rs("COD_PEDIDO")%>
             
             </td>
             <td>
             <%If comando = "alterar" and COD_PEDIDO_QS = COD_PEDIDO_RS Then%>
             	<input type="hidden" value="<%=rs("COD_PEDIDO")%>" name="COD_PEDIDO" />
				 <textarea name="ERRO_LISTA" cols="50" rows="2"><%=rs("ERRO_LISTA")%></textarea>
			 <%Else%>
				 <%=rs("ERRO_LISTA")%>
			 <%End If%> 
             </td>
             <td><div align="center">
             <%If comando = "alterar" and COD_PEDIDO_QS = COD_PEDIDO_RS Then%>
             	<input name="SALVAR" class="btnGravar" value=" " type="submit" />
             <%Else%>
             	<a href="?comando=alterar&COD_PEDIDO=<%=rs("COD_PEDIDO")%>"><img src="images/icons/icon_alterar.png" width="16" height="16" border="0" /></a>
             <%End If%>
             </div></td>
             <td><div align="center"><a href="javascript:confirma(<%=rs("COD_PEDIDO")%>)"><img src="images/icons/icon_excluir.png" width="16" height="16" border="0" /></a></div></td>
           </tr>
           <%
		   rs.MoveNext()
		   Wend
		   %>
         </table>
         </form>
		
        <div align="center">
          <p>
            <%Else%>
            </p>
          <p><em>Não há erros declarados!</em></p>
          <p>
              <%
		End If
		
		rs.Close
		Set rs = Nothing
		FecharConexao()
		%>
            </p>
        </div>
        <h3 align="center">&nbsp;</h3>
	     <p>
	       
	     <p align="center">
		  <input class="btnTopo" onclick="abre('#')" type="button" name="btnTopo2" value=" " />
		</p>		
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