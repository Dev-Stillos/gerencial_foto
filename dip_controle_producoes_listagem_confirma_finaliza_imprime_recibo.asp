<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim COD_VENDEDOR, NOME_VENDEDOR, COD_PRODUCAO, DESCRICAO
COD_VENDEDOR = Request.QueryString("COD_VENDEDOR")
COD_PRODUCAO = Request.QueryString("COD_PRODUCAO")

AbrirConexao()
	Set rs=objConexao.Execute("SELECT NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR = "&COD_VENDEDOR)
		
		If not rs.Eof Then
			NOME_VENDEDOR = rs("NOME_VENDEDOR")
		End If
		
	rs.Close
	Set rs=Nothing

	Set rs=objConexao.Execute("SELECT DESCRICAO FROM TB_PRODUCAO WHERE COD_PRODUCAO = "&COD_PRODUCAO)

		If not rs.Eof Then
			DESCRICAO = rs("DESCRICAO")
		End If

	rs.Close
	Set rs=Nothing
FecharConexao()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Vers�o 1.2</title>
	<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>
	<script type="text/javascript">
function confirma(cod_formando,nomeFormando,comando){
var mensagem

	if (comando == "alterar"){
		mensagem = "ATENÇÃO ALTERAÇÃO: \nDeseja resetar a ficha de: "+nomeFormando+" deseja continuar?\n*A ficha voltará na forma original."	
	}	
	else {
		mensagem = "ATENÇÃO EXCLUSÃO: \nSe você continuar poderá perder as alterações feitas nesta página,\nalem de resetar a ficha do formando: "+nomeFormando+"  deseja continuar?"
	}

	var resposta=confirm(mensagem);
		if (resposta==true) {
			
			if (comando == "excluir"){
				//Nova confirma��o, Exclui
				var resposta2=confirm("Deseja realmente EXCLUIR?")
					if (resposta2==true){
						document.producao["MODIFICACOES_"+cod_formando].value = "2";
						document.producao.submit();
					}	
			}
			else{
			//Altera
				document.producao["MODIFICACOES_"+cod_formando].value = "1"
				document.producao.submit();
			}
			
		}
}
</script>

	
	<script language="javascript">
		function popUp(campo,URL,campoSelect,msgErro){
			
			if(document.producao[campo].value==1){
				
				if (campoSelect !== 'no'){
					//Brinde = "Sim"
					document.producao[campoSelect].options[1].selected = "true";
				}
			
				win=window.open(URL,"","width=350,height=220,scrollbars=yes")
				var text = "Se a janela não abrir desbloquei os Pop-ups!";
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
	
	<script>

function mascara(o,f){
    v_obj=o
    v_fun=f
    setTimeout("execmascara()",1)
}

function execmascara(){
    v_obj.value=v_fun(v_obj.value)
}

function leech(v){
    v=v.replace(/o/gi,"0")
    v=v.replace(/i/gi,"1")
    v=v.replace(/z/gi,"2")
    v=v.replace(/e/gi,"3")
    v=v.replace(/a/gi,"4")
    v=v.replace(/s/gi,"5")
    v=v.replace(/t/gi,"7")
    return v
}

function soNumeros(v){
    return v.replace(/\D/g,"")
}

function telefone(v){
    v=v.replace(/\D/g,"")
    v=v.replace(/^(\d\d)(\d)/g,"($1) $2") 
    v=v.replace(/(\d{4})(\d)/,"$1-$2")   
    return v
}

function cpf(v){
    v=v.replace(/\D/g,"")                   
    v=v.replace(/(\d{3})(\d)/,"$1.$2")      
    v=v.replace(/(\d{3})(\d)/,"$1.$2")      
                                            
    v=v.replace(/(\d{3})(\d{1,2})$/,"$1-$2")
    return v
}

function cep(v){
    v.replace(/\D/g,"")               
    v=v.replace(/^(\d{5})(\d)/,"$1-$2") 
    return v
}

function cnpj(v){
    v=v.replace(/\D/g,"")                           
    v=v.replace(/^(\d{2})(\d)/,"$1.$2")             
    v=v.replace(/^(\d{2})\.(\d{3})(\d)/,"$1.$2.$3") 
    v=v.replace(/\.(\d{3})(\d)/,".$1/$2")         
    v=v.replace(/(\d{4})(\d)/,"$1-$2")              
    return v
}

function focusFormando(){
	document.producao.<%="ALBUM_"&Request.QueryString("cod_form")%>.focus()
}
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
	function confirmaFinalizarProducao(){
	
	var pergunta = confirm('Tem certeza que deseja finalizar está produção?\nObs.: Esta produção não poderá ser alterada posteriormente somente com senha especial.')
		if(pergunta==true){
		abre('?cod_prod=<%=Request.QueryString("cod_prod")%>&comando=finalizar_producao')
		}
	}
	function confirmaExclusao(){
	var perguntaq = confirm('EXCLUIR produção:\nATENÇÃO caso queira continuar lembre-se que todas as fichas(se houver) serão excluídas completamente, deseja continuar?')
		if (perguntaq==true){
			var perguntadeNovo = confirm("CONFIRMA EXCLUSÃO DESTA produção E SUAS RESPECTIVAS FICHAS?")
			if (perguntadeNovo==true){
			abre('?cod_prod=<%=Request.QueryString("cod_prod")%>&comando=excluir_producao')
			}
		}
	
	}
	
	function limparSelecao(cod_formando){
		var d=document.producoes;
		
		d["MOTIVO_RENOVAR_DEVOLVER_"+cod_formando].value = "";
		d["RENOVAR_DEVOLVER_"+cod_formando][0].checked = false;
		d["RENOVAR_DEVOLVER_"+cod_formando][1].checked = false;
	}	
	</script>
</head>

<body>

<table cellpadding="9px" cellspacing="0" width="100%">
<tr>
<td>
  <!-- #include file="includes/topo.asp" -->
  <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td><table width="100%" style="border-bottom:#999999 solid dotted 1px" background="images/layout/fundo_menu_azul.jpg"border="0" cellspacing="0" cellpadding="0">
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
  BTN_VOLTAR = 2
  BTN_IMPRIMIR = 2
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr> 
    <td valign="top" style="border:dotted 1px #999999">
	
	  <h2>Controle de Produções/Lotes - Recibo</h2>
	  <br />
	<div class="subMenu">Sub-menu: <a href="dip_controle_producoes_listagem_producoes.asp?cod_vendedor=<%=COD_VENDEDOR%>">Voltar para relação de Produções do vendedor(a)</a></div>

	<p>
	  <br />
	</p>
	<div align="center" class="textoMedio"><strong>Processo finalizado, imprima o recibo clicando no link abaixo:<br />
	</strong></div>
	              </div>
	  </p>
	<p></p>
	<p align="center"><a href="dip_controle_producoes_listagem_confirma_finaliza_IMPRIMIR.asp?cod_vendedor=<%=COD_VENDEDOR%>&COD_PRODUCAO=<%=COD_PRODUCAO%>" target="_blank" class="linkMaster">IMPRIMIR RECIBO DO CONSULTOR</a></p>
		<p></p>
	  		<p align="center"></p>
	  		<p align="center"></p>
	  		<p align="center"></p>
	  		<p align="center"></p>
	    </td>
    </tr>
</table>

  
  
  
  </td>
    </tr>
</table>  </td>
  </tr>
</table>
<!-- #include file="includes/rodape.asp" -->

  </td>
</tr>
</table>
</body>
</html>



