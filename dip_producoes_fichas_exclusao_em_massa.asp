<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
COD_FORMANDO_QS = Request.QueryString("COD_FORMANDO_QS")
NOME_FORMANDO_QS = Request.QueryString("NOME_FORMANDO_QS")

If Request.QueryString("comando") = "finalizar_producao" then
'FINALIZAR PRODUÇÃO
	AbrirConexao()
	
		objConexao.Execute("UPDATE TB_PRODUCAO SET PRODUCAO_FINALIZADA = '1' WHERE COD_PRODUCAO ="&Request.QueryString("cod_prod")&" LIMIT 1 ")
				
	FecharConexao()
	
	Response.Redirect("dip_producoes.asp?msg=Produção finalizada com sucesso!&cod_prod="&Request.QueryString("cod_prod"))
	
End If

If Request.QueryString("comando") = "excluir_producao" then
'Antes de excluir verifica novamente se a produção não está finalizada para poder excluir

	Dim rsProdq, PRODUCAO_FINALIZADA_verf
	AbrirConexao()
	
		Set rsProdq = objConexao.Execute("SELECT PRODUCAO_FINALIZADA FROM TB_PRODUCAO WHERE COD_PRODUCAO="&Request.QueryString("cod_prod"))
		
			PRODUCAO_FINALIZADA_verf = rsProdq("PRODUCAO_FINALIZADA")
		
		rsProdq.Close()
		Set rsProdq = Nothing
	
	
	If PRODUCAO_FINALIZADA_verf = 2 then
		
		'
	Dim rsqq
	
	Set rsqq = objConexao.Execute("SELECT COD_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO="&Request.QueryString("cod_prod"))
	
	If not rsqq.Eof then
	'Há diversas fichas
		
		While not rsqq.Eof
		'Iniciando dos auxiliares para trás, para finalizar excluindo as fichas e a produção em si.
				'1) Exclui brindes se houver.
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_BRINDES WHERE COD_FORMANDO="&rsqq("COD_FORMANDO"))
				
				
				'2) Exclui movimentação se houver.
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_FORMANDO="&rsqq("COD_FORMANDO"))
				
				'3) Exclui novo endereço se houver;
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO WHERE COD_FORMANDO="&rsqq("COD_FORMANDO"))
				
				'4) Exclui produtos agregados se houver;
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS WHERE COD_FORMANDO="&rsqq("COD_FORMANDO"))	
							
				'5) Exclui tabela temporária se houver;
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_TEMPORARIA WHERE COD_FORMANDO="&rsqq("COD_FORMANDO"))			
		
		rsqq.MoveNext()
		Wend
		
		rsqq.Close()
		Set rsqq = Nothing
		
		'6) Exclui todas as fichas
		objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO="&Request.QueryString("cod_prod"))
		
		'7) Exclui os produtos agregados da produção
		objConexao.Execute("DELETE FROM TB_PRODUCAO_PRODUTOS_AGREGADOS WHERE COD_PRODUCAO="&Request.QueryString("cod_prod"))
		
		'8) Finalmente exclui a produção
		objConexao.Execute("DELETE FROM TB_PRODUCAO WHERE COD_PRODUCAO="&Request.QueryString("cod_prod")&" LIMIT 1")
		
		'9) Por precaução exclui novamente a TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS, entretanto pelo código da produção
		objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS WHERE COD_FORMANDO="&Request.QueryString("cod_prod"))	
		
		Response.Redirect("dip_producoes.asp?msg=Produção excluída com sucesso!")
			
	Else
	'Não há fichas, por tanto basta excluir a produção em si.
	
		objConexao.Execute("DELETE FROM TB_PRODUCAO WHERE COD_PRODUCAO = "&Request.QueryString("cod_prod")&" LIMIT 1")
		Response.Redirect("dip_producoes.asp?msg=Produção excluída com sucesso!")
	
		
	End If	
					
	elseif PRODUCAO_FINALIZADA_verf = 1 then
	
		Response.Redirect("dip_producoes.asp?msg=Produção NÃO EXCLUÍDA, pois a mesma já está finalizada e não poderá ser excluída&cod_prod="&Request.QueryString("cod_prod"))
	
	End If
	
	FecharConexao()
	
	
End If


%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<style type="text/css">
<!--
.btnFinalizaProducao{	background-image:url(images/layout/btnFinalizaProducao.jpg);
	cursor:pointer;
	width:140px;
	height:22px;
	border:0px;
	}
-->
	/* CSS for the demo. CSS needed for the scripts are loaded dynamically by the scripts */
	#mainContainer{
		width:500px;
		margin:0 auto;
		margin-top:10px;
		border:1px double #000;
		padding:3px;
		
	}
	
	/* Alternative layout for the message box - used by one of the messages */
	
	.modalDialog_contentDiv_error{
		border:3px solid #FFF;	
		padding:2px;
		z-index:100;/* Always needed	*/
		position:absolute;	/* Always needed	*/
		background-color:#F00;	/* White background color for the message */
		color:#FFF;
	}
	.modalDialog_contentDiv_error a{
		color:#FFF;
	}
	h1{
		margin-top:0px;
	}
	#scroll {
	height:400px;
	background-color:#FFFFFF;
	overflow:auto;
	padding: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
	margin-left: 0px;
	}
	#exclusaoemMassa{
	display: none;
	visibility: hidden;
	}
	</style>
<link rel="stylesheet" href="css/modal-message.css" type="text/css">
<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
	<script type="text/javascript" src="js/ajax.js"></script>
	<script type="text/javascript" src="js/modal-message.js"></script>
	<script type="text/javascript" src="js/ajax-dynamic-content.js"></script>
	<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>
	<script type="text/javascript">
messageObj = new DHTML_modalMessage();	// We only create one object of this class
messageObj.setShadowOffset(5);	// Large shadow


function displayMessage(url)
{
	messageObj.setSource(url);
	messageObj.setCssClassMessageBox(false);
	messageObj.setSize(550,450);
	messageObj.setShadowDivVisible(true);	// Enable shadow for these boxes
	messageObj.display();
}

function displayStaticMessage(messageContent,cssClass,altura)
{
	messageObj.setHtmlContent(messageContent);
	messageObj.setSize(550,altura);
	messageObj.setCssClassMessageBox(cssClass);
	messageObj.setSource(false);	// no html source since we want to use a static message here.
	messageObj.setShadowDivVisible(false);	// Disable shadow for these boxes	
	messageObj.display();
	
	
}

function closeMessage()
{
	messageObj.close();	
}


</script>
	<script language="javascript">
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
				//Nova confirmação, Exclui
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

function confirmaFinalizado(cod_formando,nomeFormando){
	var mensagem = "ATENÇÃO ALTERAÇÃO APÓS FINALIZAÇÃO: \nDeseja resetar a ficha de: "+nomeFormando+" e voltar a forma original?"
	if(YNconfirm(mensagem)){
		location.href = 'dip_producoes.asp?cod_form='+
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
	var win27
	win27=window.open(theURL,winName,features);
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


   
      function pegarPosicaoX(objeto)
   
      {
   
          var atual_left = 0;
   
          if (objeto.offsetParent) {
   
              while (objeto.offsetParent) {
   
                  atual_left += objeto.offsetLeft
   
                  objeto = objeto.offsetParent;
   
              }
   
          }
 
          return atual_left;
 
      }


      function pegarPosicaoY(objeto)

      {
 
          var atual_top = 0;

          if (objeto.offsetParent) {

              while (objeto.offsetParent) {

                  atual_top += objeto.offsetTop
 
                  objeto = objeto.offsetParent;

              }

          }

          return atual_top;

      }
	function centralizarnaTela(ALTURA,LARGURA){
		var screenX = screen.availWidth;  
		var screenY = screen.availHeight;  
		var h = ALTURA;  
		var w = LARGURA;  
		
		//window.scrollTop((screenX) - (w / 2), (screenY / 2) - (h / 2));
		//document.producao.<%="ALBUM_"&Request.QueryString("cod_form")%>.focus();
		//d = document.producao.<%="ALBUM_"&Request.QueryString("cod_form")%>;
		//d.scrollTop(-600);
		//scrollTop((screenX) - (w / 2), (screenY / 2) - (h / 2));
	}

function focusFormando(){
	document.producao.<%="ALBUM_"&Request.QueryString("cod_form")%>.focus();

	d = document.producao.<%="ALBUM_"&Request.QueryString("cod_form")%>;
	//d.scrollIntoView();
	//centralizarnaTela(pegarPosicaoX(d),pegarPosicaoY(d));
}
</script>
<!--INICIO VALIDAÇÃO SIM/NÃO-->
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
<!--FIM VALIDAÇÃO SIM/NÃO-->	
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
	var perguntaq = confirm('EXCLUIR PRODUÇÃO:\nAtenção caso queira continuar lembre-se que todas as fichas(se houver) serão excluídas completamente, deseja continuar?')
		if (perguntaq==true){
			var perguntadeNovo = confirm("CONFIRMA EXCLUSÃO DESTA PRODUÇÃO E SUAS RESPECTIVAS FICHAS?")
			if (perguntadeNovo==true){
			abre('?cod_prod=<%=Request.QueryString("cod_prod")%>&comando=excluir_producao')
			}
		}
	
	}
	</script>
	
<script language="javascript">

function deletaFormando(nomeFormando,codFormando){
	var mensagem = "Tem certeza que deseja excluir a ficha do(a) formando(a) "+nomeFormando+"?"

	if(YNconfirm(mensagem)){
		document.producao["MODIFICACOES_"+codFormando].value = "2";
		document.producao.submit();
	}
}

function alterarFormando(nomeFormando,codFormando){
	var mensagem = "Tem certeza que deseja resetar a ficha do(a) formando(a) "+nomeFormando+"?\n*Desta maneira a ficha volta a forma inicial."

	if(YNconfirm(mensagem)){
		document.producao["MODIFICACOES_"+codFormando].value = "1";
		document.producao.submit();
	}
}
function instrucoesEtiquetas(){
	var caminho
	caminho = "<table cellpadding=\'0\' cellspacing=\'0\' border=\'0\' width=\'100%\'><tr><td><table cellpadding=\'0\' cellspacing=\'0\' width=\'100%\'><tr><td width=\'70%\'><h1>Instruções para colagem da Etiqueta</h1></td><td align=\'right\'><a href=\'#\' onclick=\'closeMessage();return false\'>Fechar Janela</a></td></tr></table></td></tr><tr><td><iframe frameborder=\'0\' scrolling=\'no\' height=\'380px\' width=\'99%\' src=\'dip_producoes_fichas_POPUP_instrucoes_etiqueta.asp?cod_prod=<%=Request.QueryString("cod_prod")%>\'></iframe></td></tr></table>"
	displayStaticMessage(caminho,false,420)		
}
function alterarFormandoFinalizado(NOME_FORMANDO,COD_FORMANDO){
	var caminho = "<table cellpadding=\'0\' cellspacing=\'0\' border=\'0\' width=\'100%\'><tr><td><table cellpadding=\'0\' cellspacing=\'0\' width=\'100%\'><tr><td width=\'70%\'><h1>Senha especial para liberação de Alteração</h1></td><td align=\'right\'><a href=\'#\' onclick=\'closeMessage();return false\'>Fechar Janela</a></td></tr></table></td></tr><tr><td><iframe frameborder=\'0\' scrolling=\'no\' height=\'175px\' width=\'99%\' src=\'dip_producoes_fichas_POPUP_confirmacao_alteracao_senha_especial.asp?cod_producao=<%=Request.QueryString("cod_prod")%>&prod_agregado=<%=Request.QueryString("prod_agre")%>&cod_formando="+COD_FORMANDO+"&nome_formando="+NOME_FORMANDO+"\'></iframe></td></tr></table>"
	displayStaticMessage(caminho,false,200)
}
</script>
	
	
<script language="javascript">
function YNconfirm(txt){
return typeof(suporteVBscript)=="undefined"?confirm(txt):VBconfirm(txt)==6
}
<%If Request.QueryString("comando") = "focuBotao" Then%>
function focoBotaoCarregar(){
alterarFormandoFinalizado('<%=NOME_FORMANDO_QS%>','<%=COD_FORMANDO_QS%>')
}
<%End If%>
</script>
<script language="vbscript">
suporteVBscript=1
function VBconfirm(mensagem)

 VBconfirm=msgbox(mensagem,vbyesno)
end function

</script>
<script>
function exclusaoSeleciona(COD_FORMANDO){

	var d = document.producao["EXCLUSAO_EM_MASSA_"+COD_FORMANDO]

	if(d.checked){
		document.getElementById('excluirCelula_'+COD_FORMANDO).style.backgroundColor ='#FFA6A6';
	}else {
		document.getElementById('excluirCelula_'+COD_FORMANDO).style.backgroundColor ='';
	}
}

function excluirRegistrosemMassa(){
	if(YNconfirm('Atenção tem certeza que deseja excluir os registros selecionados?\nSALVE OS REGISTROS ANTES DE CONTINUAR!')){
		document.producao.action='dip_producoes_fichas_exclusao_em_massa.asp';
		document.producao.submit();
	}
}
</script>
</head>

<body <%If Request.QueryString("cod_form") <> Empty then%>
	onload="focusFormando()" 
<%Elseif Request.QueryString("comando") = "focuBotao" Then
	Response.Write("onload='focoBotaoCarregar()'")
%><%End If%>>

<table cellpadding="9px" cellspacing="0" width="100%">
<tr>
<td>
  <!-- #include file="includes/topo.asp" -->
  <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td><table width="100%" style="border-bottom:#999999 solid dotted 1px" background="images/layout/fundo_menu_azul.jpg"border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><!-- #include file="includes/menu.asp" --></td>
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

	  <h2>DIP - Produ&ccedil;&otilde;es / &Aacute;lbuns </h2>
	  <h3 align="center">Formandos da Produ&ccedil;&atilde;o:<br /><%=Request.QueryString("cod_prod")%> - <%=DESCRICAO%></h3>
		  


		  <%
		  AbrirConexao()
		  
		  Dim COD_PRODUCAO, PRODUTOS_AGREGADOS, COD_FORMANDO, EXCLUSAO_EM_MASSA, COUNT_EXCLUIDOS
		  COUNT_EXCLUIDOS = 0
		  COD_PRODUCAO = Request.QueryString("cod_prod")	
		  PRODUTOS_AGREGADOS = Request.QueryString("prod_agre")
		  
		  Set rs = objConexao.Execute("SELECT COD_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = "&COD_PRODUCAO)	
					  

		  While not rs.Eof
		  
		  
		  COD_FORMANDO = rs("COD_FORMANDO")
		  EXCLUSAO_EM_MASSA = formatNumber(0&Request.Form("EXCLUSAO_EM_MASSA_"&COD_FORMANDO),0)
		  
		  If EXCLUSAO_EM_MASSA = 1 Then
		  	
			'1) Exclui Ficha de Identificação
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO="&COD_FORMANDO&" LIMIT 1")
			
				'2) Exclui brindes se houver.
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_BRINDES WHERE COD_FORMANDO="&COD_FORMANDO)
				
				
				'3) Exclui movimentação se houver.
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_FORMANDO="&COD_FORMANDO)
				
				'4) Exclui novo endereço se houver;
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO  WHERE COD_FORMANDO="&COD_FORMANDO)
				
				'5) Exclui produtos agregados se houver;
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS WHERE COD_FORMANDO="&COD_FORMANDO)	
				
				'6) Exclui tabela temporária se houver;
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_TEMPORARIA WHERE COD_FORMANDO="&COD_FORMANDO)	
			
			COUNT_EXCLUIDOS = COUNT_EXCLUIDOS + 1
		  End If
		
		
		  rs.MoveNext()
		  Wend
	      
		  rs.Close()
		  Set rs = Nothing
		  FecharConexao()
		  
		  Response.Redirect("dip_producoes_fichas.asp?cod_prod="&COD_PRODUCAO&"&prod_agre="&PRODUTOS_AGREGADOS&"&msg=Foram excluídos "&COUNT_EXCLUIDOS&" registros!")
		  %>



		  <p align="center">&nbsp;</p>		</td>
    </tr>
</table>  </td>
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