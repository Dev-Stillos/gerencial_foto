<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
If Session("TENTATIVAS") => 3 then
'Bloqueia usu�rio
AbrirConexao()
objConexao.Execute("UPDATE TB_ADM_USUARIOS SET BLOQUEADO = '1' WHERE COD_USUARIO = "&Session("COD_USUARIO")&" LIMIT 1")
FecharConexao()
Response.Redirect("adm_sair.asp")
End If


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
	<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>
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
	
<!--calend�rio-->
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
		
		calendarObjForForm.setCalendarPositionByHTMLElement(inputObject,-60,inputObject.offsetHeight-175);// Position the calendar right below the form input
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
<!--calend�rio-->

<!--valida form-->
<!--valida form-->
</head>

<body onKeyDown="javascript:Verificar()">
<!-- Code for the left panel -->
<div id="dhtmlgoodies_leftPanel">
	<a class="closeLink" href="#" onclick="initSlideLeftPanel();return false" style="font-size:15px">Fechar</a>
	<div id="leftPanelContent">
	<!-- This is the content -->
	<div>

<table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
  <td width="1%" valign="top"><img src="images/icons/icon_help.jpg" /></td>
    <td width="95%" align="left"><strong style="font-size:16px">Ajuda do Sistema</strong><br />
<strong>Esta � a tela Log do usu�rio <%=Session("NOME")%>, � poss�vel verificar as datas e hor�rios de acesso ao sistema e a quantidade de acessos tamb�m.<br />
<%If Session("COD_USUARIO") =< 2 then %>Para verificar as fichas baixadas utilize os filtros: "�ltimos registros" ou "por data."<br /><br />
Lembrete: A data inicial pr�-selecionada � correspondente ao m�s e ano atual.<%End If%>

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
	
	  <h2>Log de usu&aacute;rio </h2>
	   <div class="subMenu">Sub-Menu: <a href="?comando=alterar">Alterar Senha de Acesso</a></div>
	  <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  <p><strong>Nome: </strong><%=SESSION("NOME")%> <br />
	 <strong>Departamento: </strong><%=Session("DEPARTAMENTO")%> </p>
	  	  <%If Request.QueryString("comando") = "alterar" then
		  		
				
				If Request.Form("btnGravar") <> Empty then
					AbrirConexao()
					Dim rsUsers
					Set rsUsers = objConexao.Execute("SELECT * FROM TB_ADM_USUARIOS WHERE COD_USUARIO = '"&_
													 Session("COD_USUARIO")&"' LIMIT 1")
				
						If not rsUsers.Eof then
						
							If Request.Form("SENHA_ATUAL") = rsUsers("SENHA") then
							'Verifica se as senhas novas s�o iguais
								If Request.Form("NOVA_SENHA") = Request.Form("CONFIRMA_SENHA") then
								'Altera senha
									objConexao.Execute("UPDATE TB_ADM_USUARIOS SET SENHA = "&sqlTexto(Request.Form("NOVA_SENHA"))&_
									" WHERE COD_USUARIO = '"&Session("COD_USUARIO")&"' LIMIT 1")	
									Session("TENTATIVAS") = 0
									Response.Redirect("adm_log_usuarios.asp?msg=Senha alterada com sucesso!")
															
								else
								'N�o altera
								Session("TENTATIVAS") = Session("TENTATIVAS") + 1
								Response.Redirect("adm_log_usuarios.asp?msg=Senha n�o alterada, pois a Senha digitada nova � diferente da senha confirmada!<BR />"&_
								Session("TENTATIVAS")&" de 3 tentativas.")
								
								End If
								
							Else
							'N�o altera senha
							Session("TENTATIVAS") = Session("TENTATIVAS") + 1
							Response.Redirect("adm_log_usuarios.asp?msg=Senha n�o alterada, pois a Senha Atual digitada � incorreta!<BR />"&_
							Session("TENTATIVAS")&" de 3 tentativas")
							End If
							
							
						
						End If
					
					rsUsers.Close
					Set rsUsers = Nothing
					FecharConexao()
				End If
				
		  		  
		  %>
			<form style="margin:0px" name="ALTERAR_SENHA" method="post">
			<table width="275" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr>
    <td width="84"><strong>Senha Atual: </strong></td>
    <td width="191"><input name="SENHA_ATUAL" value="" type="password" size="30" maxlength="50" /></td>
  </tr>
  <tr>
    <td><strong>Nova Senha:</strong> </td>
    <td><input name="NOVA_SENHA" type="password" size="30" maxlength="50" /></td>
  </tr>
  <tr>
    <td><strong>Confirma&ccedil;&atilde;o:</strong> </td>
    <td><input name="CONFIRMA_SENHA" type="password" size="30" /></td>
  </tr>
  <tr>
    <td colspan="2"><div align="center">
      <input class="btnVoltar" onclick="history.go(-1)" type="button" name="btnVoltar" value=" " />   
      <input class="btnGravar" type="submit" name="btnGravar" value=" " />
    </div></td>
    </tr>
</table>
			</form>

		  <%else%>
	  
	  <h3 align="center">Acesso ao sistema </h3>
		<div id="comboTexto">
          <table width="80%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td><div align="left"><strong>Data</strong> <strong>/ Hora</strong> <%iF Session("QTDACESSOS") => 5 THEN %>
              (&uacute;ltimos 20 acessos)
              <%end if%> </div></td>
          </tr>
		  <%
		  AbrirConexao()
		  Dim rsAcessoLog
		  Set rsAcessoLog = objConexao.Execute("SELECT * FROM TB_ADM_LOG_USUARIOS WHERE COD_USUARIO="&Session("COD_USUARIO")&" ORDER BY ACESSO_EM DESC LIMIT 20")
		  While not rsAcessoLog.Eof
		  %>
          <tr>
            <td><%=rsAcessoLog("ACESSO_EM")%></td>
          </tr>
		  <%
		  rsAcessoLog.MoveNext()
		  Wend
	      
		  rsAcessoLog.Close()
		  Set rsAcessoLog = Nothing
		  FecharConexao()
		  %>
          <tr>
            <td><strong>Qtd. total de acessos: <%=Session("QTDACESSOS")%></strong></td>
          </tr>
        </table>
		</div>
		
<%end if%>
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