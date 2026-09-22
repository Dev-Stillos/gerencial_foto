<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim comando, campo1
comando = Request.QueryString("comando")
campo1 = Request.Form("campo1")

Dim NOME, LOGIN, SENHA, COD_DEPTO, BLOQUEADO, COD_PERFIL, PERMITIR_USO_EXTERNO

If comando = Empty Then
comando = "incluir"
titulo_pagina = titulo_pagina&" / Inclus�o"
End if


If comando = "incluir" and campo1 <> Empty Then
'Incluir novo registro
incluirRegistro()
titulo_pagina = titulo_pagina&" / Inclus�o"
Elseif comando = "alterar" and campo1 = Empty Then
'Leitura inicial para poss�vel altera��o
cod_usuario = Request.QueryString("cod_usuario")
lerRegistro()
titulo_pagina = titulo_pagina&" / Altera��o"
Elseif 	comando = "alterar" and campo1 <> Empty Then
'Gravar altera��o
cod_usuario = Request.Form("cod_usuario")
alterarRegistro()
titulo_pagina = titulo_pagina&" / Altera��o"
Elseif 	comando = "excluir" and campo1 = Empty Then 
'Leitura inicial para poss�vel exclus�o
cod_usuario = Request.QueryString("cod_usuario")
lerRegistro()
titulo_pagina = titulo_pagina&" / Exclus�o"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
'Exclui registro
cod_usuario = Request.Form("cod_usuario")
excluirRegistro()
titulo_pagina = titulo_pagina&" / Exclus�o"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
'Visualiza registro
cod_usuario = Request.QueryString("cod_usuario")
lerRegistro()
titulo_pagina = titulo_pagina&" / Visualiza��o"
End If

'////////////////////////////////////////////////////
'Valida��o para NÃO alterar o login da D. Andrea ou do Administrativo, somente pelos pr�prios fazer isso.
If int(Session("cod_usuario")) <> 16 and int(Session("cod_usuario")) <> 1 and (int(cod_usuario) = 16 or int(cod_usuario) = 1) Then
	'NÃO � a D. Andrea nem o Administrador e desejar alterar o login da D. Andrea ou do Administrador
	Session.Abandon()
	Response.Redirect("index.asp")
End If



Sub lerFormulario()

	NOME		= Request.Form("NOME")
	LOGIN		= Request.Form("LOGIN")
	SENHA		= Request.Form("SENHA")
	COD_DEPTO	= Request.Form("COD_DEPTO")
	BLOQUEADO	= Request.Form("BLOQUEADO")
	COD_PERFIL	= Request.Form("COD_PERFIL")
	COD_USUARIO	= Request.Form("COD_USUARIO")
	PERMITIR_USO_EXTERNO = Request.Form("PERMITIR_USO_EXTERNO")

End sub

Sub lerRegistro()
	Dim rsAl
	AbrirConexao()
	
	Set rsAl  = objConexao.Execute("SELECT * FROM TB_ADM_USUARIOS WHERE COD_USUARIO="&Request.QueryString("cod_usuario")&" AND COD_USUARIO <> 1")
		
		If not rsAl.Eof Then
			NOME = rsAl("NOME")
			LOGIN = rsAl("LOGIN")
			SENHA = rsAl("SENHA")
			COD_DEPTO = rsAl("COD_DEPTO")
			BLOQUEADO = rsAl("BLOQUEADO")
			COD_PERFIL = rsAl("COD_PERFIL")
			PERMITIR_USO_EXTERNO = rsAl("PERMITIR_USO_EXTERNO")
		End If
	
	
	rsAl.Close
	Set rsAl = Nothing
	FecharConexao()
	
	iF Request.QueryString("cod_usuario") = 1 tHEN
		Response.Redirect("adm_usuarios_incluir.asp?comando=incluir")
	End IF
End sub

Sub incluirRegistro()

	lerFormulario()

		Dim sql
sql = 	"INSERT INTO TB_ADM_USUARIOS ("&_
		"NOME ,"&_
		"LOGIN ,"&_
		"SENHA ,"&_
		"COD_DEPTO ,"&_
		"COD_PERFIL ,"&_
		"PERMITIR_USO_EXTERNO "&_
		")"&_
		"VALUES ("&_
		sqlTexto(Request.Form("NOME"))&", "&sqlTexto(Request.Form("LOGIN"))&", "&sqlTexto(Request.Form("SENHA"))&", "&sqlTexto(Request.Form("COD_DEPTO"))&", "&sqlTexto(Request.Form("COD_PERFIL"))&",'"&Request.Form("PERMITIR_USO_EXTERNO")&"')"
		
		
		Dim chavesUnicasDuplicadas,camposUnicos
			call chaveUnica ("TB_ADM_USUARIOS","LOGIN",LOGIN,"adm_usuarios.asp?codmsg=44&sucesso=2&chaveUnica=",comando,"cod_usuario",cod_usuario)
		
		AbrirConexao()
			
			objConexao.Execute(sql)
			
			Dim rsId, ULTIMO_ID
			Set rsId = objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_ADM_USUARIOS LIMIT 1")
				
					ULTIMO_ID = rsId("ULTIMO_ID")
			
			rsId.Close
			Set rsId = Nothing
		FecharConexao()
		
		Response.Redirect("adm_usuarios.asp?msg=Usuário "&NOME&" incluído com sucesso!")
End sub

Sub alterarRegistro()

	lerFormulario()
	
	sql = 	"UPDATE TB_ADM_USUARIOS SET "&_
			"NOME = "&sqlTexto(NOME)&","&_
			"LOGIN = "&sqlTexto(LOGIN)&","&_
			"SENHA = "&sqlTexto(SENHA)&","&_
			"COD_DEPTO = "&sqlTexto(COD_DEPTO)&","&_
			"PERMITIR_USO_EXTERNO = "&sqlTexto(PERMITIR_USO_EXTERNO)&","&_
			"BLOQUEADO = "&sqlTexto(BLOQUEADO)&","&_
			"COD_PERFIL = "&sqlTexto(COD_PERFIL)&" WHERE COD_USUARIO = "&COD_USUARIO&" LIMIT 1"

	call chaveUnica ("TB_ADM_USUARIOS","LOGIN",LOGIN,"adm_usuarios.asp?codmsg=44&sucesso=2&chaveUnica=",comando,"cod_usuario",cod_usuario)

	AbrirConexao()
		objConexao.Execute(sql)
	FecharConexao()
	Response.Redirect("adm_usuarios.asp?msg=Usuário "&NOME&" alterado com sucesso!")

End sub

Sub excluirRegistro()
	
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_ADM_USUARIOS WHERE COD_USUARIO = '"&COD_USUARIO&"' AND COD_USUARIO <> 1")
	FecharConexao()
	Response.Redirect("adm_usuarios.asp?msg=Usuário "&NOME&" excluído com sucesso!")
	
End sub


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
	d = document.USUARIOS
		if(d.NOME.value == ""){
			alert('Campo NOME em branco!')
			d.NOME.focus();
			return false;
		}
		if(d.COD_DEPTO.value == ""){
			alert('Campo DEPTO não selecionado!')
			d.COD_DEPTO.focus();
			return false;
		}
		if(d.LOGIN.value == ""){
			alert('Campo LOGIN em branco!')
			d.LOGIN.focus();
			return false;
		}
		if(d.SENHA.value == ""){
			alert('Campo SENHA em branco!')
			d.SENHA.focus();
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
  BTN_FONTE = 1
  BTN_IMPRIMIR_LINK = "#"
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">

	  <h2>ADM - Usu&aacute;rios / Incluir</h2>
      
      <%
	  If comando = "excluir" Then
	  %>
      	<div class="msgExcluir">
      		Tem certeza que deseja excluir este Usuário do sistema?<br />Você poderá bloquear seu acesso ao invés de excluir.
      	</div>
      <%
	  End If
	  %>
	  <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
		  <h3 align="center">Dados do Usu&aacute;rio</h3>
          <form style="margin:0px" method="post" action="adm_usuarios_incluir.asp?comando=<%=comando%>" name="USUARIOS" onsubmit="return validaForm()">
		<table width="95%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
<tr>
            <td width="20%"><div align="left"><strong>Nome</strong></div></td>
            <td width="80%">
              <div align="left">
                <input type="text" value="<%=NOME%>" name="NOME" class="inputTamanho3" id="textfield" />
                  </div></td>
            </tr>
          <tr>
            <td><div align="left"><strong>Departamento</strong></div></td>
            <td><div align="left">
              <select name="COD_DEPTO" class="inputTamanho2">
              <option>Selecione</option>
			  <%
			  AbrirConexao()
			  Set rsDeptos = objConexao.Execute("SELECT * FROM TB_ADM_USUARIOS_DEPARTAMENTOS ORDER BY NOME_DEPTO")
			  While not rsDeptos.Eof
			  Dim cod_depto2
			  cod_depto2 = rsDeptos("COD_DEPTO")
			  %>
			  <option <%If cod_depto2 = COD_DEPTO Then %> selected="selected" <%End If%> value="<%=rsDeptos("COD_DEPTO")%>"><%=rsDeptos("NOME_DEPTO")%></option>
              <%
			  rsDeptos.MoveNext()
			  Wend
			  
			  rsDeptos.Close
			  Set rsDeptos = Nothing
			  
			  %>
              </select>
              </div></td>
            </tr>
            
          <tr>
            <td><div align="left"><strong>Perfil</strong> <strong>de Acesso</strong></div></td>
            <td><div align="left"><select name="COD_PERFIL" class="inputTamanho2">
              <option value="0" selected="selected">Selecione</option>
			  <%
			  Set rsPerfis = objConexao.Execute("SELECT * FROM TB_ADM_PERFIS ORDER BY NOME_PERFIL")
			  While not rsPerfis.Eof
			  Dim cod_perfil2
			  cod_perfil2 = rsPerfis("COD_PERFIL")
			  %>
			  <option <%If cod_perfil2 = COD_PERFIL Then %> selected="selected"<%End If%>  value="<%=rsPerfis("COD_PERFIL")%>"><%=rsPerfis("NOME_PERFIL")%></option>
              <%
			  rsPerfis.MoveNext()
			  Wend
			  
			  rsPerfis.Close
			  Set rsPerfis=Nothing
			  
			  FecharConexao()
			  %>
              </select></div></td>
            </tr>
            <tr>
            <td><div align="left"><strong>Permitir acesso fora da rede</strong> <strong>Stillos</strong>. (se o acesso ao sistema estiver bloqueado, ente campo ser&aacute; ignorado.)</div></td>
            <td><div align="left">
              <label>
              <select name="PERMITIR_USO_EXTERNO" class="inputTamanho1">
              	<option <%If PERMITIR_USO_EXTERNO = 1 Then%> selected="selected"<%End If%> value="1">SIM</option>
				<option <%If PERMITIR_USO_EXTERNO <> 1  Then%> selected="selected"<%End If%>value="2">NÃO</option>
              </select>
              </label>
            </div></td>
            </tr>
          <tr>
            <td><div align="left"><strong>Login</strong></div></td>
            <td><div align="left">
              <input type="text" name="LOGIN" value="<%=LOGIN%>" class="inputTamanho2" id="textfield2" />
            </div></td>
            </tr>
          <tr>
            <td><div align="left"><strong>Senha</strong></div></td>
            <td><div align="left">
              <input type="text" name="SENHA" value="<%=SENHA%>" class="inputTamanho2" id="textfield3" />
            </div></td>
            </tr>
            
             <%If Request.QueryString("comando") = "alterar" Then%>
             <tr>
            <td><div align="left"><strong>Bloquear acesso</strong> <strong>ao sistema</strong></div></td>
            <td><div align="left">
              <label>
              <select name="BLOQUEADO" class="inputTamanho1">
              	<option <%If BLOQUEADO = 1 Then%> selected="selected"<%End If%> value="1">SIM</option>
				<option <%If BLOQUEADO = 2 Then%> selected="selected"<%End If%>value="2">NÃO</option>
              </select>
              </label>
            </div></td>
            </tr>
            <%End If%>
          <tr>
            <td>&nbsp;</td>
            <td>
            <input type='hidden' value='1' name='campo1'  />
            <%If comando = "alterar" Then%>
            	<input type="hidden" name="COD_USUARIO" value="<%=Request.QueryString("cod_usuario")%>" />
	            <input class="btnGravar" type="submit" name="btnAlterar" id="button" value=" " /> 
            <%Elseif comando = "incluir" Then%>
    	         <input class="btnGravar" type="submit" name="btnIncluir" id="button" value=" " /> 
            <%Elseif comando = "excluir" Then%>
            	<input type="hidden" name="COD_USUARIO" value="<%=Request.QueryString("cod_usuario")%>" />
            	  <input class="btnExcluir" type="submit" name="btnExcluir" id="button" value=" " />
			<%End If%>
            </td>
          </tr>
		        </table>
           </form>
        <div align="center"></div>
        <h3 align="center">&nbsp;</h3>
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