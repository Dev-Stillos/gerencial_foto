<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim cod_perfil, comando, campo1
cod_perfil = Request.QueryString("cod_perfil")
comando = Request.QueryString("comando")
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
	-->
</head>

<body onKeyDown="javascript:Verificar()">
<!-- Code for the left panel -->
<div id="dhtmlgoodies_leftPanel">
	<a class="closeLink" href="#" onClick="initSlideLeftPanel();return false" style="font-size:15px">Fechar</a>
	<div id="leftPanelContent">
	<!-- This is the content -->
	<div>

<table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
  <td width="1%"><img src="images/icons/icon_help.jpg" /></td>
    <td width="95%" align="left"><strong style="font-size:16px">Ajuda do Sistema</strong><br />
<strong>	Esta é a tela inicial do sistema, verifique o número de acessos ao sistema e a data da última utilização.<br />
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
<form action="adm_perfis_formulario.asp?cod_perfil=<%=COD_PERFIL%>&comando=<%=comando%>" onsubmit="return validaForm()" method="post" style="margin:0px;">
	  <h2>ADM - M&oacute;dulos / Perfis / Cadastro</h2>
	   <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	  <h3 align="center">Cadastro de Perfil</h3>
	  	  <table width="400" border="0" align="center" cellpadding="2" cellspacing="2">
            <tr>
              <td width="22%"><strong>Nome Perfil:</strong></td>
              <td width="78%">
                <input class="inputTamanho2" type="text" name="NOME_PERFIL" value="<%=Request.QueryString("nome_perf")%>" id="textfield" />
             </td>
            </tr>
          </table>
	  	
          <table width="95%" border="0" align="center" cellpadding="3" style="border-collapse:collapse">
<%

Dim habilitada_para_todos_PRIMEIRA
Dim verifica_formulario

AbrirConexao()

'////////////////////////////////////
		If comando = "incluir" Then
		
			'Grava o nome do perfil na Base e gerar um cód de perfil
			objConexao.Execute("INSERT INTO TB_ADM_PERFIS (NOME_PERFIL) VALUES ("&sqlTexto(Request.Form("NOME_PERFIL"))&")")
			
			'Verifica o código gerado
			Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_ADM_PERFIS LIMIT 1")
				If not rs.Eof Then
					cod_perfil = rs("ULTIMO_ID")
				End if 
			rs.Close
			Set rs=Nothing
		
		Else
		
			objConexao.Execute("UPDATE TB_ADM_PERFIS SET NOME_PERFIL = "&sqlTexto(Request.Form("NOME_PERFIL"))&" WHERE COD_PERFIL ="&COD_PERFIL&" LIMIT 1")
		
		End If
'////////////////////////////////////	

Set rsMaster = objConexao.Execute("SELECT TB_ADM_PAGINAS_HABILITADAS.*, TB_ADM_MODULOS.* from TB_ADM_MODULOS inner join TB_ADM_PAGINAS_HABILITADAS on TB_ADM_MODULOS.cod_modulo = TB_ADM_PAGINAS_HABILITADAS.cod_modulo  where pag_filha_cod_grupo_acesso_pagina = 0  group by TB_ADM_MODULOS.cod_modulo order by TB_ADM_MODULOS.order_MENU,TB_ADM_PAGINAS_HABILITADAS.caminho_valida")

	While not rsMaster.Eof
	
				COD_MODULO = ""&rsMaster("COD_MODULO")
				NOME_MODULO = rsMaster("NOME_MODULO")
				
				if Cstr(COD_MODULO) = "" Then
					'Não possuem grupo
					NOME_MODULO = "Variados"
				End if
%>
  <%	
	If "2" = "2" Then
	
	'Seleciona todas que não são filhas e não são únicas 
	sql =	"SELECT TB_ADM_PAGINAS_HABILITADAS.*, TB_ADM_MODULOS.* from TB_ADM_MODULOS inner join TB_ADM_PAGINAS_HABILITADAS on TB_ADM_MODULOS.cod_modulo = TB_ADM_PAGINAS_HABILITADAS.cod_modulo  where pag_filha_cod_grupo_acesso_pagina = 0 and TB_ADM_PAGINAS_HABILITADAS.cod_modulo = "&COD_MODULO&" order by TB_ADM_MODULOS.order_MENU,TB_ADM_PAGINAS_HABILITADAS.nome_pagina"


	Set rs=objConexao.Execute(sql)
			
	If not rs.Eof Then
%>
<tr bgcolor="#99FDC4">
 <td colspan="3" bgcolor="#ffffff" style="padding-top:30px"><div class="textoMedio" style="padding:5px;border-bottom:solid 2px #CCCCCC; background-color:#DDECFF"><strong><%=NOME_MODULO%></strong></div></td>
</tr>

  <%
		While not rs.Eof
		
		acesso_pagina = 2
		

		
 %>
  <%If CStr(rs("habilitada_para_todos")) = "2" Then
  
			NOME_PAGINA = rs("NOME_PAGINA")
 		 	COD_PAGINA = rs("COD_PAGINA")
		 
			'///////////////////////////
			'Valida
			verifica_formulario = int(Request.Form("valida_"&COD_PAGINA))
			
			If verifica_formulario = 1 Then
			'Selecionado
				
				'1) Verifica se já exite, se já existir não faz nada, caso contrário se não existir então grava
				Set rsSubPadrao = objConexao.Execute("SELECT * FROM TB_ADM_PERFIS_PAGINAS_HABILITADAS Where cod_perfil = '"&cod_perfil&"' and COD_PAGINA = '"&COD_PAGINA&"'")
					
					If rsSubPadrao.Eof Then
						objConexao.Execute( "INSERT INTO TB_ADM_PERFIS_PAGINAS_HABILITADAS (COD_PERFIL,COD_PAGINA) "&_
											"VALUES ( "&_
											COD_PERFIL&", "&COD_PAGINA&_
											")")
					End If
			
				rsSubPadrao.Close
				Set rsSubPadrao = Nothing
					
			Else
			'Não selecionado
			
				'1) Verifica se já exite, se já existir EXCLUIR, caso contrário se não existir então não faz nada			
				Set rsSubPadrao = objConexao.Execute("SELECT * FROM TB_ADM_PERFIS_PAGINAS_HABILITADAS Where cod_perfil = '"&cod_perfil&"' and COD_PAGINA = '"&COD_PAGINA&"'")
					
					If not rsSubPadrao.Eof Then
						objConexao.Execute("DELETE FROM TB_ADM_PERFIS_PAGINAS_HABILITADAS WHERE COD_PERFIL = "&COD_PERFIL&" AND COD_PAGINA="&COD_PAGINA&" LIMIT 1")												
					End If
			
				rsSubPadrao.Close
				Set rsSubPadrao = Nothing
			
			End If
			'///////////////////////////
			
			
  %>
  <tr>
    <td width="21" bgcolor="#FFFFFF">
    <div align="right">

      <input name="valida_<%=COD_PAGINA%>" type="checkbox" class="inputRadio" value="1"  <%If acesso_pagina = 1 Then%> checked="checked" <%End If%> />
      <input type="hidden" value="<%=acesso_pagina%>" name="valida_antiga_<%=COD_PAGINA%>" />
    </div></td>
    <td>&nbsp;</td>
    <td><div align="left" style="font-size:13px"><strong><%=NOME_PAGINA%></strong></div></td>
    </tr>
	<%
	'Seleciona as sub-validações das páginas PAI
	
	
	Set rs2=objConexao.Execute("Select * From TB_ADM_PAGINAS_HABILITADAS_SUB_VALIDACOES Where COD_PAGINA = '"&rs("COD_PAGINA")&"' order by NOME_SUB_VALIDACAO")
		If not rs2.Eof Then
			While not rs2.Eof
			
			COD_SUB_VALIDACAO = rs2("COD_SUB_VALIDACAO")
			acesso_pagina_sub = 2
			
			
			'///////////////////////////
			'Valida
			verifica_formulario = int(Request.Form("valida_sub_"&COD_SUB_VALIDACAO))
			
			If verifica_formulario = 1 Then
			'Selecionado
				
				'1) Verifica se já exite, se já existir não faz nada, caso contrário se não existir então grava
				Set rsSubValida = objConexao.Execute("Select * From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES Where COD_SUB_VALIDACAO = "&COD_SUB_VALIDACAO&" and cod_perfil = '"&cod_perfil&"'")
			
					If rsSubValida.Eof then
						objConexao.Execute(	"INSERT INTO TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES (COD_PERFIL,COD_SUB_VALIDACAO) "&_
						"VALUES ( "&_
						COD_PERFIL&", "&COD_SUB_VALIDACAO&_
						")")
					End If
				
				rsSubValida.Close
				Set rsSubValida = Nothing	
				
					
			Else
			'Não selecionado
			
				'1) Verifica se já exite, se já existir EXCLUIR, caso contrário se não existir então não faz nada
				Set rsSubValida = objConexao.Execute("Select * From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES Where COD_SUB_VALIDACAO = "&COD_SUB_VALIDACAO&" and cod_perfil = '"&cod_perfil&"'")
			
					If not rsSubValida.Eof then
						objConexao.Execute("DELETE FROM TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES WHERE COD_PERFIL = "&COD_PERFIL&" AND COD_SUB_VALIDACAO="&COD_SUB_VALIDACAO&" LIMIT 1")
					End If
				
				rsSubValida.Close
				Set rsSubValida = Nothing	
				
			
			End If
			'///////////////////////////
			
			
			
			
	%>
     <tr bgcolor="#f2f2f2">
    <td bgcolor="#FFFFFF">&nbsp;</td>
    <td bgcolor="#FFFFFF">&nbsp;</td>
    <td bgcolor="#f2f2f2"><div align="left">
      <table width="400" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="20"><input type="checkbox" name="valida_sub_<%=COD_SUB_VALIDACAO%>" <%If acesso_pagina_sub = 1 Then%> checked="checked" <%End If%> class="inputRadio" value="1" />
            <input type="hidden" name="valida_sub_antiga_<%=COD_SUB_VALIDACAO%>" value="<%=acesso_pagina_sub%>" />            </td>
            <td width="380"><span style="font-weight: bold"><%=rs2("NOME_SUB_VALIDACAO")%></span></td>
          </tr>
        </table>
      </div></td>
    </tr>
  	<%
			rs2.MoveNext()
			Wend
		End if
	rs2.Close
	Set rs2=Nothing
	%>
  	<%
	'Seleciona as filhas desta página
	sql = "SELECT TB_ADM_PAGINAS_HABILITADAS.*, TB_ADM_MODULOS.* from TB_ADM_MODULOS inner join TB_ADM_PAGINAS_HABILITADAS on TB_ADM_MODULOS.cod_modulo = TB_ADM_PAGINAS_HABILITADAS.cod_modulo  where pag_filha_cod_grupo_acesso_pagina = "&rs("COD_PAGINA")&" order by TB_ADM_MODULOS.order_MENU,TB_ADM_PAGINAS_HABILITADAS.caminho_valida"
	Set rs2=objConexao.Execute(sql)
		
		If not rs2.Eof Then
			While not rs2.Eof
			
			acesso_pagina = 2
			
			COD_PAGINA = rs2("COD_PAGINA")
	%>
  <tr bgcolor="#f2f2f2">
    <td bgcolor="#FFFFFF">&nbsp;</td>
    <td bgcolor="#FFFFFF">&nbsp;</td>
    <td bgcolor="#f2f2f2"><div align="left">
      <table width="400" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="20"><%

			'///////////////////////////
			'Valida
			verifica_formulario = int(Request.Form("valida_"&COD_PAGINA))
			
			If verifica_formulario = 1 Then
			'Selecionado
				
				'1) Verifica se já exite, se já existir não faz nada, caso contrário se não existir então grava
				Set rsSubPadrao = objConexao.Execute("SELECT * FROM TB_ADM_PERFIS_PAGINAS_HABILITADAS Where cod_perfil = '"&cod_perfil&"' and COD_PAGINA = '"&COD_PAGINA&"'")
					
					If rsSubPadrao.Eof Then
						objConexao.Execute( "INSERT INTO TB_ADM_PERFIS_PAGINAS_HABILITADAS (COD_PERFIL,COD_PAGINA) "&_
											"VALUES ( "&_
											COD_PERFIL&", "&COD_PAGINA&_
											")")
					End If
			
				rsSubPadrao.Close
				Set rsSubPadrao = Nothing
					
			Else
			'Não selecionado
			
				'1) Verifica se já exite, se já existir EXCLUIR, caso contrário se não existir então não faz nada			
				Set rsSubPadrao = objConexao.Execute("SELECT * FROM TB_ADM_PERFIS_PAGINAS_HABILITADAS Where cod_perfil = '"&cod_perfil&"' and COD_PAGINA = '"&COD_PAGINA&"'")
					
					If not rsSubPadrao.Eof Then
						objConexao.Execute("DELETE FROM TB_ADM_PERFIS_PAGINAS_HABILITADAS WHERE COD_PERFIL = "&COD_PERFIL&" AND COD_PAGINA="&COD_PAGINA&" LIMIT 1")												
					End If
			
				rsSubPadrao.Close
				Set rsSubPadrao = Nothing
			
			End If
			'///////////////////////////
	
	%>
      <input name="valida_<%=COD_PAGINA%>" type="checkbox" class="inputRadio" value="1"  <%If acesso_pagina = 1 Then%> checked="checked" <%End If%> />
      <input type="hidden" value="<%=acesso_pagina%>" name="valida_antiga_<%=COD_PAGINA%>" /></td>
              <td width="380"><strong><%=rs2("NOME_PAGINA")%></strong></td>
            </tr>
          </table>
           
    </div></td>
    </tr>
  			
    <%
            'Seleciona as sub-validações das páginas FILHA		
			Set rs3=objConexao.Execute("Select * From TB_ADM_PAGINAS_HABILITADAS_SUB_VALIDACOES Where COD_PAGINA = '"&COD_PAGINA&"' order by NOME_SUB_VALIDACAO")
		
			
			If not rs3.Eof Then
				While not rs3.Eof
				
				COD_SUB_VALIDACAO = rs3("COD_SUB_VALIDACAO")
				acesso_pagina_sub = 2
			
			
				'///////////////////////////
				'Valida
				verifica_formulario = int(Request.Form("valida_sub_"&COD_SUB_VALIDACAO))
				
				If verifica_formulario = 1 Then
				'Selecionado
					
					'1) Verifica se já exite, se já existir não faz nada, caso contrário se não existir então grava
					Set rsSubValida = objConexao.Execute("Select * From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES Where COD_SUB_VALIDACAO = "&COD_SUB_VALIDACAO&" and cod_perfil = '"&cod_perfil&"'")
				
						If rsSubValida.Eof then
							objConexao.Execute(	"INSERT INTO TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES (COD_PERFIL,COD_SUB_VALIDACAO) "&_
							"VALUES ( "&_
							COD_PERFIL&", "&COD_SUB_VALIDACAO&_
							")")
						End If
					
					rsSubValida.Close
					Set rsSubValida = Nothing	
					
						
				Else
				'Não selecionado
				
					'1) Verifica se já exite, se já existir EXCLUIR, caso contrário se não existir então não faz nada
					Set rsSubValida = objConexao.Execute("Select * From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES Where COD_SUB_VALIDACAO = "&COD_SUB_VALIDACAO&" and cod_perfil = '"&cod_perfil&"'")
				
						If not rsSubValida.Eof then
							objConexao.Execute("DELETE FROM TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES WHERE COD_PERFIL = "&COD_PERFIL&" AND COD_SUB_VALIDACAO="&COD_SUB_VALIDACAO&" LIMIT 1")
						End If
					
					rsSubValida.Close
					Set rsSubValida = Nothing	
					
				
				End If
				'///////////////////////////
				
				
				
			%>	
				<tr bgcolor="#99FDC4">
		<td bgcolor="#FFFFFF">&nbsp;</td>
		<td width="12" bgcolor="#FFFFFF">&nbsp;</td>
		<td bgcolor="#E8E8E8"><table width="600" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="22">&nbsp;</td>
              <td width="24"><div align="center">
                <input type="checkbox" name="valida_sub_<%=COD_SUB_VALIDACAO%>" <%If int(acesso_pagina_sub) = 1 Then%> checked="checked" <%End If%> class="inputRadio" value="1" />
              </div></td>
              <td width="554"><span style="font-weight: bold"><%=rs3("NOME_SUB_VALIDACAO")%></span></td>
            </tr>
                  </table>		  <div align="left"></div></td>
	  </tr>
			
			<%	
				rs3.MoveNext()
				Wend
			End If
  			
			rs3.Close
			Set rs3=Nothing
			
			rs2.MoveNext()
			Wend
		
			
		
		End if
	rs2.Close
	Set rs2=Nothing
	%>
  <tr bgcolor="#99FDC4">
 <td colspan="3" bgcolor="#FFFFFF"><div style="border-bottom:solid 1px #ccc; margin-left:22px; height:2px; font-size:1px; color:#FFFFFF">1</div> </td>
 </tr>
  <%End If%>
 
  <%
		rs.MoveNext()
		Wend
%>

<%
	End if
	
	rs.Close
	Set rs=Nothing
	
	End if
	
rsMaster.MoveNext()
Wend

rsMaster.Close
Set rsMaster = Nothing

FecharConexao()

Response.Redirect("adm_perfis.asp?msg=Perfil incluído/alterado com sucesso!&cod_perfil="&cod_perfil)
%>
</table>
             
          <p align="center">
         
            <input class="btnGravar" type="submit" name="button" id="button" value=" " />
       
          </p>
		  <p align="center">&nbsp;</p>		
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