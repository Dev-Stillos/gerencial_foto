<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>Usu&aacute;rios & Configura&ccedil;&otilde;es</title>
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

	  <h2>ADM - Usu&aacute;rios &amp; Configura&ccedil;&otilde;es</h2>
       <%
	   Response.Write("<br />")
	Dim sucesso, codmsg
	codmsg	= CStr(""&Request.QueryString("codmsg"))
	sucesso = CStr(""&Request.QueryString("sucesso"))
		
	If sucesso = "" and codmsg <> "" Then%>
    	<div align="center" class="msgGrandeSucesso">
        	<%
			If codmsg = "1" Then
				Response.Write("Vendedor inclu&iacute;do com sucesso!")
			Elseif codmsg = "2" Then
				Response.Write("Vendedor alterado com sucesso!")
			Elseif codmsg = "3" Then	
				Response.Write("Vendedor exclu&iacute;do com sucesso!")
			End If	
			%>
            
        
        </div>
    <%ElseIf sucesso <> "" and codmsg <> "" Then%>
        <div align="center" class="msgGrandeErro">
			<%
			If codmsg = "4" Then
				Response.Write("Vendedor n&atilde;o alterado para n&atilde;o haver duplicidade de informa&ccedil;&atilde;o!<BR>O conte&uacute;do do campo """&Request.QueryString("chaveUnica")&""" j&aacute; possui uma informa&ccedil;&atilde;o igual em outro Vendedor.<br />")
			Elseif codmsg = "5" Then
				Response.Write("Vendedor n&atilde;o exclu&iacute;do, pois h&aacute; pedidos que dependem deste Vendedor.<BR>Se n&atilde;o desejar ter mais este Vendedor inative ao inv&iacute;s de excluir.<br />")
				
			'////
			'Usu�rios
			ElseIf codmsg = "44" Then
				Response.Write("Usu&aacute;rio n&atilde;o alterado para n&atilde;o haver duplicidade de informa&ccedi;&atilde;o!<BR>O conte&uacute;do do campo """&Request.QueryString("chaveUnica")&""" j&aacute; possui uma informa&ccedil;&atilde;o igual em outro Usu&aacute;rio.<br />")
			
			End If	
			%>
</div>
	<%End If%>
      
      

	   <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
       	<%
		If validaUsuario(1,53,"",false)	Then
		%>
        <!-- #include file="includes/adm_usuarios_sub_menu.asp" -->
        
	  	  <%
		  AbrirConexao()
		  Dim rs, rs2, rs3, sql
		 
		  sql = "SELECT TB_ADM_USUARIOS.BLOQUEADO,TB_ADM_USUARIOS.PERMITIR_USO_EXTERNO,TB_ADM_USUARIOS.COD_USUARIO, TB_ADM_USUARIOS.NOME, TB_ADM_USUARIOS.COD_PERFIL,TB_ADM_USUARIOS.LOGIN, TB_ADM_USUARIOS_DEPARTAMENTOS.NOME_DEPTO FROM TB_ADM_USUARIOS_DEPARTAMENTOS INNER JOIN TB_ADM_USUARIOS ON TB_ADM_USUARIOS_DEPARTAMENTOS.COD_DEPTO = TB_ADM_USUARIOS.COD_DEPTO WHERE COD_USUARIO <> 1 and BLOQUEADO = 2"
	  
		  Set rs = objConexao.Execute(sql)
		  
		  If not rs.Eof then
		  %>
		  <h3 align="center">Usu&aacute;rios Internos:</h3>
		<table width="95%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="20%"><div align="center"><strong>Usu&aacute;rio</strong></div></td>
            <td width="21%"><div align="center"><strong>Departamento</strong></div></td>
            <td width="13%"><div align="center"><strong>Login</strong></div></td>
            <td width="14%"><div align="center"><strong>Perfil</strong></div></td>
            <td width="15%"><div align="center"><strong>&Uacute;ltimo Acesso:</strong></div></td>
            <td><div align="center"><strong>Possui acesso Externo</strong></div></td>
            <td width="5%">&nbsp;</td>
            </tr>
		  <%
		  Dim COD_USUARIO, NOME, LOGIN, NOME_DEPTO, ACESSO_EM, COD_PERFIL
		  
		  While not rs.Eof
		  
		  COD_PERFIL	= formatNumber(0&rs("COD_PERFIL"),0)
		  COD_USUARIO 	= rs("COD_USUARIO")
		  NOME			= rs("NOME")	
		  LOGIN 		= rs("LOGIN")
		  NOME_DEPTO	= rs("NOME_DEPTO")
		  
		  Set rs2 = objConexao.Execute("SELECT ACESSO_EM FROM TB_ADM_LOG_USUARIOS WHERE COD_USUARIO ="&COD_USUARIO&" ORDER BY ACESSO_EM DESC LIMIT 1")
		  
		  	If not rs2.Eof Then
				ACESSO_EM = rs2("ACESSO_EM")
			Else
				ACESSO_EM = "N&atilde;o houve o primeiro acesso"
			End If
		  
		  rs2.Close
		  Set rs2 = Nothing
		  %>
          <tr>
            <td><div align="left"><%=NOME%></div></td>
            <td><div align="center"><%=NOME_DEPTO%></div></td>
            <td><div align="center"><%=LOGIN%></div></td>
            <td><div align="center"><%
			If COD_PERFIL <> 0 Then
				Set rsPerfil = objConexao.Execute("SELECT * FROM TB_ADM_PERFIS WHERE COD_PERFIL="&COD_PERFIL)
					Response.Write(rsPerfil("NOME_PERFIL"))
				rsPerfil.Close	
				Set rsPerfil = Nothing
			Else
				Response.Write("N&atilde;o h&aacute; perfil selecionado")
			End If
			%></div></td>
            <td><div align="center"><%=ACESSO_EM%></div></td>
            <td><div align="center">
              <%
			If int(rs("PERMITIR_USO_EXTERNO")) = 1 Then
			%>
              Sim
              <%Else%>
              N&atilde;o
              <%
			End If
			%>
            </div></td>
            <td><div align="center">
              <%
			If int(rs("cod_usuario")) = 16 Then
				If (int(Session("cod_usuario")) = 16 or int(Session("cod_usuario")) = 1) Then%>
              <a href="adm_usuarios_incluir.asp?cod_usuario=<%=COD_USUARIO%>&comando=alterar"><img alt="Alterar Usu&aacute;rio" src="images/icons/icon_alterar.png" border="0" /></a>&nbsp;&nbsp;<a href="adm_usuarios_incluir.asp?cod_usuario=<%=COD_USUARIO%>&amp;comando=excluir"><img alt="Excluir usu&aacute;rio" src="images/icons/icon_excluir.png" border="0" /></a>
              <%
				Else
					Response.Write("Altera&ccedil;&atilde;o somente pela diretoria.")
				End if
			Else%>
              <a href="adm_usuarios_incluir.asp?cod_usuario=<%=COD_USUARIO%>&comando=alterar"><img alt="Alterar Usu&aacute;rio" src="images/icons/icon_alterar.png" border="0" /></a>&nbsp;&nbsp;<a href="adm_usuarios_incluir.asp?cod_usuario=<%=COD_USUARIO%>&amp;comando=excluir"><img alt="Excluir usu&aacute;rio" src="images/icons/icon_excluir.png" border="0" /></a>
              <%End If%>
              
              
              
            </div></td>
          </tr>
		  <%
		  rs.MoveNext()
		  Wend
	      %>
        </table>
        <div align="center">
          <%Else%>
          N&atilde;o h&aacute; usu&aacute;rios cadastrados!
          <%End IF%>
           <%
		  rs.Close()
		  Set rs = Nothing
		  
		  FecharConexao()
		  %>
       
        <%
		Else
		%>
       	<div class="subMenu"><a href="adm_vendedores_formulario.asp">Incluir Vendedor</a></div>
        <%
		End If
		%>    
        </div>
        <h3 align="center">Vendedores:</h3>
		<%
		AbrirConexao()
		Set rs = objConexao.Execute("SELECT * FROM TB_VENDEDORES WHERE ATIVADO = 1 AND COD_VENDEDOR > 1 ORDER BY NOME_VENDEDOR")
		
		If not rs.Eof Then
		%>
        <table width="95%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="33%"><div align="center"><strong>Representante Comercial</strong></div></td>
            <td width="19%"><div align="center"><strong>Login</strong></div></td>
            <td width="20%"><div align="center"><strong>Grupo de Repasse</strong></div></td>
            <td width="23%"><div align="center"><strong>Tabela de Comissionamento</strong></div></td>
            <td width="5%">&nbsp;</td>
</tr>
		  <%
		  Dim COD_VENDEDOR, NOME_VENDEDOR, ATIVADO

		  
		  While not rs.Eof
		  
		  COD_VENDEDOR	= rs("COD_VENDEDOR")
		  NOME_VENDEDOR = rs("NOME_VENDEDOR")
		  LOGIN 		= rs("LOGIN")
		  SENHA			= rs("SENHA")
		  ATIVADO 		= booleano(rs("ATIVADO"))
		  %>
          <tr <%If int(rs("ATIVADO")) = 2 Then%> bgcolor="#999999" <%End If%>>
            <td><div align="left"><%=NOME_VENDEDOR%></div></td>
            <td><div align="center"><%=LOGIN%></div></td>
            <td><div align="center">
            <%
			If rs("COD_GRUPO_VENDEDORES") = 1 Then
				Response.Write("Renato")
			Elseif rs("COD_GRUPO_VENDEDORES") = 2 Then
				Response.Write("Mario")
			Elseif rs("COD_GRUPO_VENDEDORES") = 3 Then
				Response.Write("Paulo Mantuam")
			Else
				Response.Write("-")
			End if
			%>
            </div></td>
            <td><div align="left">
              <%Set rsE10=objConexao.Execute("SELECT COD_TABELA_COMISSAO_VENDEDORES, DESCRICAO FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES WHERE COD_TABELA_COMISSAO_VENDEDORES='"&rs("COD_TABELA_COMISSAO_VENDEDORES")&"' ORDER BY DESCRICAO ASC")
If not rsE10.Eof Then%>
              <%=rsE10(1)%>
              <%End if%>
              <%rsE10.Close

Set rsE10=Nothing
%>
            </div></td>
            <td><div align="center"><a href="adm_vendedores_formulario.asp?cod_vendedor=<%=COD_VENDEDOR%>&amp;comando=alterar"><img alt="Alterar Usu&aacute;rio" src="images/icons/icon_alterar.png" border="0" /></a>&nbsp;&nbsp;<a href="adm_vendedores_formulario.asp?cod_vendedor=<%=COD_VENDEDOR%>&amp;comando=excluir"><img alt="Excluir usu&aacute;rio" src="images/icons/icon_excluir.png" border="0" /></a></div></td>
          </tr>
		  <%
		  rs.MoveNext()
		  Wend
	      %>
        </table>
        <%End If%>
        <%
		rs.Close
		Set rs = Nothing
		FecharConexao()
		%>
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