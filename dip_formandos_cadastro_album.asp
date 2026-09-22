<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim cod_formando, vendedor_preselecionado
Dim NOME_FORMANDO, NOME_CURSO
vendedor_preselecionado = formatNumber(0&Session("vendedor_preselecionado"),0)

cod_formando = Request.QueryString("cod_formando")
Dim rs
AbrirConexao()
Set rs = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO="&cod_formando)

		If not rs.Eof Then
			NOME_FORMANDO	= rs("NOME_FORMANDO")
			NOME_CURSO		= rs("NOME_CURSO") 
		End If

rs.Close
Set rs = Nothing
FecharConexao()

If Request.Form("gravarPre") <> Empty Then
	Session("vendedor_preselecionado") = Request.Form("FICHAS_EM_POSSE")
	Response.Redirect("dip_formandos_cadastro_album.asp?cod_formando="&cod_formando)
End If


If Request.Form("Grava2") <> Empty Then
	
	Dim ALBUM, ALBUM_DUPLO, INICIAL, COMISSAO, BRINDE, FICHAS_EM_POSSE, ALBUM_EM_POSSE	
	
	ALBUM			=	Request.Form("ALBUM")
	ALBUM_DUPLO		=	Request.Form("ALBUM_DUPLO")	
	INICIAL			=	Request.Form("INICIAL")	
	COMISSAO		=	Request.Form("COMISSAO")	
	BRINDE			=	Request.Form("BRINDE")	
	FICHAS_EM_POSSE	=	formatNumber(0&Request.Form("FICHAS_EM_POSSE"),0)	
	ALBUM_EM_POSSE	=	formatNumber(0&Request.Form("ALBUM_EM_POSSE"),0)
	
	sql = 	" UPDATE TB_FICHAS_IDENTIFICACAO SET ALBUM = "&ALBUM&","&_
			" ENCADERNACOES_EXTRAS = "&ALBUM_DUPLO&","&_
			" INICIAL = "&INICIAL&","&_
			" BRINDE = "&BRINDE&","&_
			" COMISSAO = "&COMISSAO&" WHERE COD_FORMANDO ="&COD_FORMANDO&" LIMIT 1"
	
	AbrirConexao()

	
	'1� Parte atualiza ficha	
	Response.Write("<br /><br /><br />")
	Response.Write(sql)
	objConexao.Execute(sql)
	
	'2� Parte verifica se j� existe o álbum, se exitir de fato grava o mesmo nas movimenta��es
	
	'/B�sico, primeiro insere a ficha que está em posse ou do vendedor ou da Stillo's
	Dim COMPLETA_DATA_FINAL
	
	If FICHAS_EM_POSSE = 27 Then
	'Se estiver com a Stillos atualiza data para a mesma da sa�da
	COMPLETA_DATA_FINAL = "NOW( )"
	Else
	'Caso contr�rio se estiver com vendedor NÃO dar� data do retorno
	COMPLETA_DATA_FINAL = "NULL"
	End If
	
	sql = 	" INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES( "&_
			" COD_FORMANDO ,"&_
			" TIPO ,"&_
			" COD_VENDEDOR ,"&_
			" DATA_SAIDA ,"&_
			" DATA_RETORNO ,"&_
			" CADASTRADO_POR "&_
			" )"&_
			" VALUES ("&_
			 cod_formando&", '1', "&FICHAS_EM_POSSE&", NOW( ), "&COMPLETA_DATA_FINAL&", "&Session("COD_USUARIO")&_
			")"
	
			Response.Write("<br /><br /><br />")
			Response.Write(sql)
			objConexao.Execute(sql)
	
	'/Agora sim verifica se existe álbum
	
	If ALBUM_EM_POSSE <> 1  Then
				
			If ALBUM_EM_POSSE = 2 Then
			
				COMPLETA_DATA_FINAL = "NOW( )"
				
			Elseif ALBUM_EM_POSSE = 3 Then
			
				COMPLETA_DATA_FINAL = "NULL"								
				
			End If
		
			sql = 	" INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES( "&_
			" COD_FORMANDO ,"&_
			" TIPO ,"&_
			" COD_VENDEDOR ,"&_
			" DATA_SAIDA ,"&_
			" DATA_RETORNO ,"&_
			" CADASTRADO_POR "&_
			" )"&_
			" VALUES ("&_
			 cod_formando&", '0', "&FICHAS_EM_POSSE&", NOW( ), "&COMPLETA_DATA_FINAL&", "&Session("COD_USUARIO")&_
			")"	
			Response.Write("<br /><br /><br />")
			Response.Write(sql)
			objConexao.Execute(sql)
	
	End If
	
	
	FecharConexao()
	Response.Redirect("dip_formandos_cadastro_album_impressao_etiquetas.asp?cod_formando="&cod_formando)
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
	function validaForm(){
	d = document.CAD_FORMANDOS_2
		if(d.ALBUM.value == ""){
		d.ALBUM.focus();
		alert('Campo álbum NÃO digitado!');
		return false;
		}
	
		if(d.INICIAL.value == ""){
		d.INICIAL.focus();
		alert('Campo inicial NÃO digitado!');
		return false;
		}
	
		if(d.ALBUM_EM_POSSE.value == 3 && d.FICHAS_EM_POSSE.value == 27){
		d.ALBUM_EM_POSSE.focus();
		alert('Você informou que o álbum está com o vendedor, mas NÃO informou qual vendedor é!');
		return false;
		}		
	}
	</script>
	
	<script language="javascript">
		function popUp(campo,URL,campoSelect,msgErro){
			
			if(document.CAD_FORMANDOS_2[campo].value==1){
				
				if (campoSelect !== 'no'){
					//Brinde = "Sim"
					document.CAD_FORMANDOS_2[campoSelect].options[1].selected = "true";
				}
			
				win=window.open(URL,"","width=350,height=220,scrollbars=yes")
				var text = "Se a janela NÃO abrir desbloquei os Pop-ups!";
				if(win == null) { alert(text); return; }
				win.moveTo(350,150);
					
			}
			else {
				if (msgErro==1){alert('Para visualizar os brindes é necessário alterar a opção Brinde para Sim.')}
			}
		
		}
		
		
		function fecharPopUp(){
			if(win){
				win.close();
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
			
		function focuInic(){
		document.CAD_FORMANDOS_2.ALBUM.focus();
		}	
	</script>
</head>

<body onload="focuInic()">

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
	 
	  <div class="subMenu">
	   <form action="?cod_formando=<%=COD_FORMANDO%>" method="post" style="margin:0px">
	     <table width="100%" border="0" cellspacing="0" cellpadding="0">
           <tr>
             <td width="7%">Sub-Menu:</td>
             <td width="19%" class="textoMedio">Pr&eacute;-selecionar o vendedor: </td>
             <td width="29%" class="textoMedio"><div align="left">
                 <select name="FICHAS_EM_POSSE">
                   <option selected="selected" value="27">STILLO'S</option>
                                      <%
							  AbrirConexao()
								  Dim rsVendedor
								  Set rsVendedor = objConexao.Execute("SELECT * FROM TB_VENDEDORES WHERE COD_VENDEDOR <> 1 AND COD_VENDEDOR <> 27 ORDER BY NOME_VENDEDOR")
								  %>
                                      <%While not rsVendedor.Eof %>
                                           <option value="<%=rsVendedor("COD_VENDEDOR")%>"><%=TrocarAcento(rsVendedor("NOME_VENDEDOR"))%></option>
                                      <%
									rsVendedor.MoveNext()
									Wend
									
									rsVendedor.Close()
									Set rsVendedor = Nothing
									
									FecharConexao()
									%>
                 </select>
                 &nbsp;</div></td>
             <td width="45%" class="textoMedio"><input type="submit" class="btnGravar" name="gravarPre" value=" " /></td>
           </tr>
         </table>
	   </form>
	   </div>

	   </form>
	   <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	  <h3 align="center">Informe os outros dados:  </h3>
		  <div id="comboTexto">
		  	<form name="CAD_FORMANDOS_2" method="post" onsubmit="return validaForm()" action="dip_formandos_cadastro_album.asp?cod_formando=<%=COD_FORMANDO%>"> 
          	<table width="95%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
		        	<tr>
            <td width="17%" bgcolor="#CCCCCC"><strong>Nome Formando: </strong></td>
            <td width="83%" bgcolor="#CCCCCC"><strong><%=NOME_FORMANDO%></strong></td>
            </tr>
					<tr>
				<td width="17%" bgcolor="#CCCCCC"><strong>Curso:</strong></td>
            <td width="83%" bgcolor="#CCCCCC"><strong><%=NOME_CURSO%></strong></td>
            </tr>
          <tr>
            <td>N&ordm; &Aacute;lbum / &Aacute;lbum Duplo:</td>
            <td><input name="ALBUM" type="text" size="4" maxlength="4" /> 
              / 
                <select name="ALBUM_DUPLO">
                  <option value="0" selected="selected">N&Atilde;O</option>
                  <option value="1">SIM</option>
                </select></td>
            </tr>
          <tr>
            <td>Qtd. Inicial:</td>
            <td><input name="INICIAL" type="text" size="4" maxlength="4" /></td>
            </tr>
          <tr>
            <td>Comiss&atilde;o:</td>
            <td><select name="COMISSAO" onchange="popUp('COMISSAO','dip_producoes_fichas_POPUP_brindes.asp?cod_form=<%=cod_formando%>&cod_prod=0','BRINDE',2)">
                <option value="2" selected="selected">NÃO</option>
                <option value="1">SIM</option>
              </select></td>
            </tr>
          <tr>
            <td>Brinde:</td>
            <td><select name="BRINDE" onchange="popUp('BRINDE','dip_producoes_fichas_POPUP_brindes.asp?cod_form=<%=cod_formando%>&cod_prod=0','no',2)">
                <option value="2" selected="selected">NÃO</option>
                <option value="1">SIM</option>
              </select></td>
            </tr>
          <tr>
            <td>Ficha em Posse de: </td>
            <td>
			<%If vendedor_preselecionado <> 0 Then
				 AbrirConexao()
				 Dim rsvendPre
				 	Set rsvendPre =	objConexao.Execute("SELECT NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR="&vendedor_preselecionado)
					
						If not rsvendPre.Eof Then
							Response.Write(rsvendPre("NOME_VENDEDOR"))
						End If
					
					rsvendPre.Close
					Set rsvendPre =	Nothing
				 FecharConexao()
				%>
				<input type="hidden" name="FICHAS_EM_POSSE" value="<%=vendedor_preselecionado%>" />

			<%Else%>
				<select name="FICHAS_EM_POSSE">
			<option selected="selected" value="27">STILLO'S</option>
									<%
							  AbrirConexao()
								   Set rsVendedor = objConexao.Execute("SELECT * FROM TB_VENDEDORES WHERE COD_VENDEDOR <> 1 AND COD_VENDEDOR <> 27 AND ATIVADO='1' ORDER BY NOME_VENDEDOR")
								  %>
									<%While not rsVendedor.Eof %>
									<option value="<%=rsVendedor("COD_VENDEDOR")%>"><%=TrocarAcento(rsVendedor("NOME_VENDEDOR"))%></option>
									<%
									rsVendedor.MoveNext()
									Wend
									
									rsVendedor.Close()
									Set rsVendedor = Nothing
									
									FecharConexao()
									%>
								  </select> 
				(para maior segurança pré-selecione um vendedor na barra acima)
			    <%End If%>					  
			</td>
          </tr>
          <tr>
            <td>&Aacute;lbum em Posse de: </td>
            <td><select name="ALBUM_EM_POSSE">
			<option selected="selected" value="1">NÃO SE SABE NO MOMENTO</option>
			<option value="2">ESTOQUE, ÁLBUM JÁ ENCONTRADO</option>
			<option value="3">VENDEDOR ACIMA RELACIONADO</option>
            </select>            </td>
          </tr>
          <tr>
            <td colspan="2"><div align="center">
              <input class="btnAvancar" type="submit" name="Grava2" value=" " />
            </div></td>
            </tr>
		         </table>
		  	</form>
		  </div>
		
		  <p>&nbsp;</p>
		  		  <p>&nbsp;</p>
				  		  <p align="center" class="textoMaior">&nbsp;</p>
						  		  <p>&nbsp;</p>
								  		  <p>&nbsp;</p>
		                             
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