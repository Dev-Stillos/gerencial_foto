<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<%
'Para tirar a valida��o da p�gina basta tirar o "#" ou coloca-lo para habilitar antes do include
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
<title>Cadastro de Formandos</title>
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
	
function comboFocus(id){
document.getElementById(id).style.fontWeight='bold';
//document.getElementById(id).style.borderStyle='solid';
//document.getElementById(id).style.borderWidth='2px';
document.getElementById(id).style.backgroundColor='#FACFCF';
}

function comboBlur(id,ultimo){
document.getElementById(id).style.fontWeight="";
document.getElementById(id).style.borderColor="";
document.getElementById(id).style.borderStyle="";
document.getElementById(id).style.borderWidth="";
document.getElementById(id).style.backgroundColor="";

	if (ultimo == 'ultimo'){
		document.getElementById('btn_combo_1').focus();
	}

}

function Setas(idParaCima,idParaBaixo)
{
	var tecla = window.event.keyCode;
	//alert(tecla);
	
	//alert(String.fromCharCode(tecla));
	if (tecla == 81){
		document.getElementById(idParaCima).focus();
	}else if (tecla == 65){
		document.getElementById(idParaBaixo).focus();
	}

}
function loadInicio(){
	<%If BUSCAR_TERMO = "" Then%>
		document.BUSCA_QW.BUSCAR_TERMO.focus();
	<%Elseif BUSCAR_TERMO <> "" Then%>
		document.getElementById('btn_combo_1').focus();
	<%End If%>
	
}
	</script><style type="text/css">
<!--
#botnsAvanca:Hover{
border:solid 2px #CC0000;
padding:5PX;
}
-->
    </style></head>

<body onload="loadInicio()">

<table cellpadding="9px" cellspacing="0" width="100%">
<tr>
<td>
<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td><!-- #include file="includes/topo.asp" -->
      <table width="100%" style="border-bottom:#999999 solid dotted 1px" background="images/layout/fundo_menu_azul.jpg"border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
	<!-- #include file="includes/menu.asp" -->	  </td>
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

	  <h2>DIP - Formandos Cadastro Individual</h2>
	   <div class="subMenu">
	   <form action="dip_formandos.asp" name="BUSCA_QW" style="margin:0px">
	     <table width="100%" border="0" cellspacing="0" cellpadding="0">
           <tr>
             <td width="7%">Sub-Menu:</td>
             <td width="14%" class="textoMedio"> Buscar para cadastro: </td>
             <td width="24%"><input class="inputTamanho2" type="text" name="BUSCAR_TERMO" /></td>
             <td width="12%">
               <div align="left">
                 <select name="TIPO_BUSCA">
				 <option selected="selected" value="1">FORMANDO</option>
				 <option value="2">FACULDADE</option>
				 <option value="3">CONTRATO</option>
                 </select>
                 </div></td>
             <td width="10%"><input class="btnFiltrar" type="submit" name="Submit2" value=" " /></td>
			 <%
			 AbrirConexao()
			 Dim rs877, FICHAS_ERROS
			 Set rs877 = objConexao.Execute("SELECT COD_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE ALBUM = 0 AND COD_PRODUCAO = 0",FICHAS_ERROS)
			 Set rs877 = Nothing
			 FecharConexao()
			 %>
             <td width="33%">&nbsp;</td>
           </tr>
           <tr>
             <td colspan="6"><br /><a href="dip_formandos_cadastro_contratos_formandos.asp">Contratos somente Foto</a> | <a href="dip_formandos_cadastro_completo.asp?comando=novo">Cadastro de Formando Completo</a> | <a href="dip_formando_albuns_com_erro.asp">Fichas com Erros (<%=FICHAS_ERROS%>)</a></td>
             </tr>
         </table>
	   </form>
	   </div>
	   <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
       <div align="center">
	  	    
	  	    
	  	          <%
		  If BUSCAR_TERMO <> Empty Then
		  
		  	AbrirConexao()
		  	
		 	Dim rs, rs2, rs3, sql

		 	If TIPO_BUSCA = 1 Then
			
				sql = "SELECT CONTRATO, TURMA, CODIGO_FORMANDO, NOME_FORMANDO, NOME_CURSO FROM TB_FORMANDOS_IMP WHERE NOME_FORMANDO LIKE '%"&BUSCAR_TERMO&"%' ORDER BY NOME_FORMANDO" 	
				'Response.Write(sql)
		  	Elseif TIPO_BUSCA = 2 Then
				sql = "SELECT CONTRATO, TURMA, CODIGO_FORMANDO, NOME_FORMANDO, NOME_CURSO FROM TB_FORMANDOS_IMP WHERE TURMA LIKE '%"&BUSCAR_TERMO&"%' ORDER BY NOME_FORMANDO"			
			Else
				sql = "SELECT CONTRATO, TURMA, CODIGO_FORMANDO, NOME_FORMANDO, NOME_CURSO FROM TB_FORMANDOS_IMP WHERE CONTRATO ="&BUSCAR_TERMO&" ORDER BY NOME_FORMANDO"
			End If
		  Set rs = objConexao.Execute(sql,qtdRegistros)
		  
		  If not rs.Eof then
		  %>
	  	    </div>
	  	  <h3 align="center">Foram  encontrados os seguintes registros:
	  	  
</h3>
Observação<br />
		  Utilize o teclado para movimentar os registros: precione <strong>Q</strong>
 para ir para cima e <strong>A</strong> para baixo.     
 <div id="comboTexto">
		  	<form name="CAD_FORMANDOS" method="post"> 
          	<table width="95%" border="0" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="11%" style="border-left: #CCCCCC solid 1px; border-bottom: #CCCCCC solid 1px; border-top: #CCCCCC solid 1px"><div align="center"><strong>Contrato</strong></div></td>
            <td width="31%" style="border-left: #CCCCCC solid 1px; border-bottom: #CCCCCC solid 1px; border-top: #CCCCCC solid 1px"><div align="center"><strong>Instituição de Ensino </strong></div></td>
            <td width="31%" style="border-left: #CCCCCC solid 1px; border-bottom: #CCCCCC solid 1px; border-top: #CCCCCC solid 1px"><div align="center"><strong>Formando</strong></div></td>
            <td width="14%" style="border-left: #CCCCCC solid 1px; border-bottom: #CCCCCC solid 1px; border-top: #CCCCCC solid 1px"><div align="center"><strong>Curso </strong></div></td>
            <td width="13%" style="border: #CCCCCC solid 1px"><div align="center"><strong>Cadastrar</strong></div></td>
            </tr>
			</table>
		  <%
  		Dim combo
		combo = 0
		  
		  While not rs.Eof
		  
		  combo = combo + 1
		  Dim CONTRATO, TURMA, CODIGO_FORMANDO, NOME_FORMANDO, NOME_CURSO
		  Dim CADASTRADO
		  CADASTRADO		= 2
		  CONTRATO			= rs("CONTRATO")	
		  TURMA				= rs("TURMA")	
		  CODIGO_FORMANDO	= rs("CODIGO_FORMANDO")	
		  NOME_FORMANDO		= rs("NOME_FORMANDO")
		  Dim posicaoInicial
		  
		  posicaoInicial = inStr(NOME_FORMANDO,"(")
		  If posicaoInicial > 0 Then
					NOME_FORMANDO = Left(NOME_FORMANDO,posicaoInicial-1)
		  End If
		  	
		  NOME_CURSO		= rs("NOME_CURSO")	
		  
	  Dim rsVerificaImportado
	  Set rsVerificaImportado = objConexao.Execute("SELECT COD_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO="&CODIGO_FORMANDO&"0")
	  		
				If not rsVerificaImportado.Eof Then
						
						CADASTRADO = 1
						
				End If
			
	  rsVerificaImportado.Close
	  Set rsVerificaImportado = Nothing
	  %>
	  	        	<div>
					<table id="combo_<%=combo%>" width="95%" border="0" align="center" 	cellpadding="5" style="border-collapse:collapse">
          <tr>
            <td width="11%" style="border-left: #CCCCCC solid 1px; border-bottom: #CCCCCC solid 1px"><div align="center"><%=CONTRATO%></div></td>
            <td width="31%" style="border-left: #CCCCCC solid 1px; border-bottom: #CCCCCC solid 1px"><%=TURMA%></td>
            <td width="31%" style="border-left: #CCCCCC solid 1px; border-bottom: #CCCCCC solid 1px"><%=NOME_FORMANDO%> (<%=CODIGO_FORMANDO%>)</td>
            <td width="14%" style="border-left: #CCCCCC solid 1px; border-bottom: #CCCCCC solid 1px"><%=NOME_CURSO%></td>
            <td width="13%" style="border-left: #CCCCCC solid 1px; border-right: #CCCCCC solid 1px; border-bottom: #CCCCCC solid 1px"><div align="center">
              <%
			  Dim IRPARACIMA, IRPARABAIXO, JA_CADASTRADO_CIMA, JA_CADASTRADO_BAIXO
			  %>
			  <%If combo = 1 Then
			  
			  	IRPARACIMA			= qtdRegistros
				IRPARABAIXO			= combo+1

				
			  Elseif combo = qtdRegistros Then
			  	
				IRPARACIMA			= combo-1
				IRPARABAIXO			= 1

			  Else
			  
				IRPARACIMA			= combo-1
  				IRPARABAIXO			= combo+1

			  
			  End If
			  %>
			  <%If CADASTRADO = 1 Then%>
			 <a href="#" onblur="comboBlur('combo_<%=combo%>','')" id="btn_combo_<%=combo%>" onfocus="comboFocus('combo_<%=combo%>')" onkeydown="return Setas('btn_combo_<%=IRPARACIMA%>','btn_combo_<%=IRPARABAIXO%>')" >
              <img  src="images/icons/icon_green.gif" border="0" width="16" height="15" />
			  </a>
			  <br />
			  
              <span class="textoMenor">(Já cadastrado)</span>
			  <%Else%>
			  
			  <a href="#" onblur="comboBlur('combo_<%=combo%>','')" id="btn_combo_<%=combo%>" onfocus="comboFocus('combo_<%=combo%>')" onclick="return validaForm('<%=CODIGO_FORMANDO%>','<%=NOME_FORMANDO%>')" onkeydown="return Setas('btn_combo_<%=IRPARACIMA%>','btn_combo_<%=IRPARABAIXO%>')" ><img src="images/layout/btnAvancar.jpg" border="0" /></a>
             
              <%End If%>
            </div></td>
            </tr>
			</table>
					</div>
		  <%
		  rs.MoveNext()
		  Wend
	      
		  rs.Close()
		  Set rs = Nothing
		  FecharConexao()
		  %>
    <table width="95%" border="0" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr>
            <td style="border-top: #CCCCCC solid 1px"></td>
            </tr>
			</table>
		  	</form>
		  </div>
		
		  <div align="center">
		    <p>
		      <%else%>
		      <em>Não foi encontrado registros para a busca informada!</em>	</p>
		    <p></p>
		    <p>	      
		      <%end if%>
		      </p>
		  </div>
		  <%Else%>
		  <p></p>
		  		  <p></p>
				  		  <p align="center" class="textoMaior">Para iniciar o cadastro do formando utilize a busca do sub-menu!</p>
						  		  <p></p>
								  		  <p></p>
		                             
		                                    <%End If%>
	
		                                  
		  <p align="center">
		  <input class="btnTopo" onclick="abre('#')" tabindex="999999" type="button" name="btnTopo2" value=" " />
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



