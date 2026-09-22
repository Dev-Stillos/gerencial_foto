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
<title>StillosWEB - Sistema corporativo - Versão 1.2</title>
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
		document.producao.action='dip_producoes_fichas_exclusao_em_massa.asp?cod_prod=<%=Request.QueryString("cod_prod")%>&prod_agre=<%=Request.QueryString("prod_agre")%>';
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

	  <h2>DIP - Produções / Álbuns </h2>
	  <%
	  Dim rs9, PRODUCAO_FINALIZADA, DESCRICAO, VERIFICA_FINALIZACAO
	  
	  VERIFICA_FINALIZACAO = 1
	  
	  AbrirConexao()
			  Set rs9 = objConexao.Execute("SELECT PRODUCAO_FINALIZADA, DESCRICAO FROM TB_PRODUCAO WHERE COD_PRODUCAO="&Request.QueryString("cod_prod"))

					If not rs9.Eof then
					
						PRODUCAO_FINALIZADA = rs9("PRODUCAO_FINALIZADA")
						DESCRICAO = rs9("DESCRICAO") 
					End If					

			  rs9.Close()
	  		  Set rs9 = Nothing 
	  FecharConexao()

	  %>
	  
	   <div class="subMenu">Sub-Menu:
	   <%If PRODUCAO_FINALIZADA  = 2 then%>
		   <a href="dip_producoes_alteracoes.asp?cod_prod=<%=Request.QueryString("cod_prod")%>&comando=alterar">Alterar esta produção</a> | 
	   
		   <a href="javascript:void(0)" onclick="confirmaExclusao()">Excluir esta produção</a> |
		   <a href="javascript:void(0)" onclick="ExclusaoMassa()">Exclusão em massa</a> |
		   <a href="dip_producoes_importar_formandos.asp?cod_prod=<%=Request.QueryString("cod_prod")%>">Importar Formandos</a> | 
		   <a href="dip_formandos_cadastro_completo.asp?cod_prod=<%=Request.QueryString("cod_prod")%>&comando=novo">Cadastrar Formando</a>


		   
	       <%End If%>
	   <%If PRODUCAO_FINALIZADA  = 1 then%>
	   		 <a href="dip_producoes_fichas_lotes.asp?cod_prod=<%=Request.QueryString("cod_prod")%>">Lotes por vendedor</a> | <a href="javascript:Void(0)" onclick="abrir('dip_producoes_fichas_etiquetas2.asp?cod_prod=<%=Request.QueryString("cod_prod")%>&desc=<%=Left(DESCRICAO,23)%>','teste','width=200,height=200,scrollbars=no')">Imprimir Etiquetas</a> | <a href="javascript:Void(0)" onclick="instrucoesEtiquetas()">Instruções para colagem de Etiqueta</a>
	         <%End If%>
			 
		| <a href="dip_producoes_fichas_relacao_por_vendedor.asp?cod_prod=<%=Request.QueryString("cod_prod")%>&comando=alterar">Relacionar por vendedor (somente cadastrados)</a></div>
	    
	   

	   <%If PRODUCAO_FINALIZADA  = 1 then%>
	  <div>
	    <table width="98%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="4%"><img src="images/icons/lista_exclamation.jpg" width="44" height="42" /></td>
            <td width="96%"><span class="msg"><strong class="msg">Atenção</strong>!</span><br />
              Esta produção está finalizada poderá ser alterada somente com senha especial. </td>
          </tr>
        </table>
	  </div>
	  <%End If%> 
	 	   
		   
	       <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	  <%
		  AbrirConexao()
		  'AbrirConexao2()
		  
		  Dim rs, rs2, rs3
		  Dim QTD_ALBUNS, FICHAS_VENDA

		  Set rs = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO="&Request.QueryString("cod_prod")&" ORDER BY NOME_FORMANDO")
		  
		  If not rs.Eof then
		  %>
		  <h3 align="center">Formandos da Produção:<br /><%=Request.QueryString("cod_prod")%> - <%=DESCRICAO%></h3>
		  <form action="dip_producoes_fichas_confirma2.asp?cod_prod=<%=Request.QueryString("cod_prod")%>&prod_agre=<%=Request.QueryString("prod_agre")%>" method="post" name="producao">
		  
          <table width="98%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="33%"><div align="center"><strong>Nome Formando </strong></div></td>
            <td width="10%"><div align="center"><strong>Cód. Form. </strong></div></td>
            <td width="9%"><div align="center"><strong>Contrato</strong></div></td>
            <td width="9%"><div align="center"><strong>Álbum / Duplo </strong></div></td>
            <td width="6%"><div align="center"><strong>Inicial</strong></div></td>
            <td width="7%"><div align="center"><strong>Comissão</strong></div></td>
            <td width="8%"><div align="center"><strong>Brinde</strong></div></td>
            <td width="6%"><div align="center"><strong>Prod. Agreg. </strong></div></td>
            <td width="6%"><div align="center"><strong>Lote</strong></div></td>
			<td width="6%"> </td>
</tr>

		  <%
		  While not rs.Eof
		  
		  Dim INICIAL, COMISSAO, PRODUTO_AGREGADO, VENDIDO, ENCADERNACOES_EXTRAS, LOTE
		  
		  INICIAL = 0&rs("INICIAL")
		  INICIAL = formatNumber(INICIAL,0)
		  VENDIDO = formatNumber(0&rs("VENDIDO"),0)
		  
		  COMISSAO =  0&rs("COMISSAO")
		  COMISSAO = formatNumber(COMISSAO,0)
		  
		  PRODUTO_AGREGADO = 0&rs("PRODUTO_AGREGADO")
		  PRODUTO_AGREGADO = formatNumber(PRODUTO_AGREGADO,0)
		  ENCADERNACOES_EXTRAS = formatNumber(0&rs("ENCADERNACOES_EXTRAS"),0)
		  
		  LOTE = rs("LOTE")
		  
		  '/VISUALIZA A GRAVAÇÃO TEMPORÁRIA NA PÁG. CONFIRMAÇÃO NA PRÓX. TELA
		  Dim rsTemporaria, rsEntidadeEnsino
		  Dim TEMP_ALBUM, TEMP_INICIAL, TEMP_COMISSAO, TEMP_BRINDE, TEMP_COD_VENDEDOR, TEMP_PROD_AGREGADO, TEMP_ENCADERNACOES_EXTRAS, TEMP_LOTE
		  
		  TEMP_ALBUM = 0
		  TEMP_ENCADERNACOES_EXTRAS = 0
		  TEMP_INICIAL = 0
		  TEMP_COMISSAO = 0
		  TEMP_BRINDE = 0
		  TEMP_COD_VENDEDOR = 0
		  TEMP_PROD_AGREGADO = 0
		  TEMP_LOTE = 0
		  
		  Set rsTemporaria = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO_TEMPORARIA WHERE COD_FORMANDO="&rs("COD_FORMANDO"))
		  
		  	if not rsTemporaria.Eof then
			
				  TEMP_ALBUM				= rsTemporaria("ALBUM")
				  TEMP_INICIAL				= rsTemporaria("INICIAL")
				  TEMP_COMISSAO 			= rsTemporaria("COMISSAO")
				  TEMP_BRINDE			 	= rsTemporaria("BRINDE")
				  'TEMP_COD_VENDEDOR		 	= rsTemporaria("COD_VENDEDOR") 'Não é mais necessário
				  TEMP_PROD_AGREGADO	 	= rsTemporaria("PROD_AGREGADO")
				  TEMP_ENCADERNACOES_EXTRAS	= rsTemporaria("ENCADERNACOES_EXTRAS")
				  TEMP_LOTE					= rsTemporaria("LOTE")
				
			End If
		  	
		  
		  rsTemporaria.Close()
		  Set rsTemporaria = Nothing
		  
		  
		  Dim COD_FORM_STRING, COD_RS_FORMANDO
		  COD_FORM_STRING = 0&Request.QueryString("cod_form")
		  COD_FORM_STRING = formatNumber(COD_FORM_STRING,0)
		  COD_RS_FORMANDO = 0&rs("COD_FORMANDO")
		  COD_RS_FORMANDO = formatNumber(COD_RS_FORMANDO,0)
		  %>
          <tr <%If COD_FORM_STRING = COD_RS_FORMANDO then%> bgcolor="#FFFF99"<%End If%> id="excluirCelula_<%=rs("COD_FORMANDO")%>">
            <td><%=rs("NOME_FORMANDO")%></td>
            <td><%=digito(rs("COD_FORMANDO"))%></td>
            <td>
				<div align="center"><%=rs("COD_CONTRATO")%>				      </div></td>
            <td><div align="center">
			<%If INICIAL > 0 then%>
					
					<%=rs("ALBUM")%>
					<%If ENCADERNACOES_EXTRAS = 1 Then %>
						<strong> DUPLO</strong>
					<%End If%>
			
			<%Else%>  
					<%
					VERIFICA_FINALIZACAO = 2
						
					If  TEMP_ALBUM <= 0 Then
						TEMP_ALBUM = ""
					End If
					
					IF TEMP_INICIAL <= 0 Then
						TEMP_INICIAL = ""
					End If
					
					IF PRODUCAO_FINALIZADA = 1 Then
						TEMP_ALBUM = ""
						TEMP_INICIAL = ""
					End If
						%>
			 <input value="<%=TEMP_ALBUM%>" style="width:40px" onkeypress="mascara(this,soNumeros)" type="text" name="ALBUM_<%=rs("COD_FORMANDO")%>" />
             <input value="1" class="inputRadio" <%If TEMP_ENCADERNACOES_EXTRAS = 1 Then%> checked="checked" <%End If%> type="checkbox" name="ENCADERNACOES_EXTRAS_<%=rs("COD_FORMANDO")%>" />
			 <%End If%>
            </div></td>
            <td><div align="center">
             			<%If INICIAL > 0 then%>
				<%=INICIAL%>
			<%Else%>
              <input value="<%=TEMP_INICIAL%>" onkeypress="mascara(this,soNumeros)" style="width:40px" type="text" name="INICIAL_<%=rs("COD_FORMANDO")%>" />
			<%End If%>
            </div></td>
            <td><div align="center">
			<%If INICIAL > 0 then%>
				<%If COMISSAO = 1 then%>
				Sim
				<%Else%>
				Não
				<%End If%>
			<%Else%>
			    <select name="COMISSAO_<%=rs("COD_FORMANDO")%>" onchange="popUp('COMISSAO_<%=rs("COD_FORMANDO")%>','dip_producoes_fichas_POPUP_brindes.asp?cod_form=<%=rs("COD_FORMANDO")%>&cod_prod=<%=rs("COD_PRODUCAO")%>','BRINDE_<%=rs("COD_FORMANDO")%>',2)">
                <option value="2" <%If TEMP_COMISSAO <> 1 then%>selected="selected"<%End If%>>NÃO</option>
                <option value="1" <%If TEMP_COMISSAO = 1 then%> selected="selected"<%End If%>>SIM</option>
              </select>
			<%End if%> 
			  
            </div></td>
            <td><div align="center">
    			<%If INICIAL > 0 then%>
				<%If COMISSAO = 1 then%>
				Sim
				<%Else%>
				Não
				<%End If%>
			<%Else%>
			
					<table border="0" align="center" cellpadding="0" cellspacing="0">
					  <tr>
						<td width="68%"> <select name="BRINDE_<%=rs("COD_FORMANDO")%>" onchange="popUp('BRINDE_<%=rs("COD_FORMANDO")%>','dip_producoes_fichas_POPUP_brindes.asp?cod_form=<%=rs("COD_FORMANDO")%>&cod_prod=<%=rs("COD_PRODUCAO")%>','no',2)">
                <option value="2" <%If TEMP_BRINDE <> 1 then%>selected="selected"<%End If%>>NÃO</option>
                <option value="1" <%If TEMP_BRINDE = 1 then%>selected="selected"<%End If%>>SIM</option>
              </select></td>
						<td width="32%">
						<div align="center" id="listaBrindes"><a href="#" tabindex="9999"><img onclick="popUp('BRINDE_<%=rs("COD_FORMANDO")%>','dip_producoes_fichas_POPUP_brindes.asp?cod_form=<%=rs("COD_FORMANDO")%>&cod_prod=<%=rs("COD_PRODUCAO")%>','no',1)" src="images/icons/icon_bloco_notas.gif" width="24" height="29" border="0" /></a>						</div></td>
					  </tr>
					</table>

			
		
			    
			    <%End if%> 
            </div></td>
            <td><div align="center">
			
			<%If int(Request.QueryString("prod_agre")) = 1 then%>
				<%If int(INICIAL) > 0 then%>
					<%If int(PRODUTO_AGREGADO) = 1 then%>
					Sim
					<%Else%>
					Não
					<%End If%>
				<%Else%>
   				  <select name="PRODUTO_AGREGADO_<%=rs("COD_FORMANDO")%>">
					<option value="2" <%If int(TEMP_PROD_AGREGADO) = 2 then%>selected="selected"<%End If%>>NÃO</option>
				  	<option value="1" <%If int(TEMP_PROD_AGREGADO) = 1 or TEMP_PROD_AGREGADO = 0 then%>selected="selected"<%End If%>>SIM</option>
				  </select>
				<%End if%> 
			<%else%>
			<input name="PRODUTO_AGREGADO_<%=rs("COD_FORMANDO")%>" type="hidden" value="2" />
			Não
			<%End If%>
			
            </div></td>
            <td><div align="center">
             			<%If INICIAL > 0 then%>
				<%=LOTE%>
			<%Else%>
              <input value="<%=TEMP_LOTE%>" onkeypress="mascara(this,soNumeros)" style="width:40px" type="text" name="LOTE_<%=rs("COD_FORMANDO")%>" />
			<%End If%>
            </div></td>
             
			<td>
			<%If PRODUCAO_FINALIZADA = 2 THEN%>
				<table width="50" border="0" align="center" cellpadding="2" cellspacing="0">
              <tr>
                <td width="25"><div align="center">
				<%If INICIAL > 0 then%>
                <a href="javascript:alterarFormando('<%=rs("NOME_FORMANDO")%>','<%=rs("COD_FORMANDO")%>')" tabindex="98777"><img src="images/icons/icon_alterar.png" width="16" height="16" border="0" />				</a>
               <%End If%>
			    </div></td>
                <td width="25">
				<div align="center">
				
				<div style="border: solid 1px #999999; padding:2px">			  
				  <a href="javascript:deletaFormando('<%=rs("NOME_FORMANDO")%>','<%=rs("COD_FORMANDO")%>')" tabindex="999999">
                  <img src="images/icons/icon_excluir.png" border="0" />				  </a>
				<input name="EXCLUSAO_EM_MASSA_<%=rs("COD_FORMANDO")%>" onclick="javascript:exclusaoSeleciona('<%=rs("COD_FORMANDO")%>')" class="inputRadio" type="checkbox" value="1" />
				</div>
                </div>
				<input style="width:35px" name="MODIFICACOES_<%=rs("COD_FORMANDO")%>" value="0" type="hidden" />				</td>
              </tr>
            </table>	
			 <%Elseif PRODUCAO_FINALIZADA = 1 AND VENDIDO = 2 AND INICIAL > 0 Then%>
			    <div align="center">
                <a id="BTN_<%=rs("COD_FORMANDO")%>" href="javascript:alterarFormandoFinalizado('<%=rs("NOME_FORMANDO")%>','<%=rs("COD_FORMANDO")%>')" tabindex="98777">
				<img src="images/icons/icon_alterar_block.png" width="16" height="18" border="0"  />				</a>            </div>
			    <div align="center">
			      <%Elseif PRODUCAO_FINALIZADA = 1 AND VENDIDO = 2 AND INICIAL = 0 Then%>
			            <input style="width:35px" name="MODIFICACOES_<%=rs("COD_FORMANDO")%>" value="9" type="hidden" />
			            <%Elseif PRODUCAO_FINALIZADA = 1 AND VENDIDO = 1 Then%>
  Álbum vendido e corrigido
  <%End If%>			
                </div></td>
			</tr>
          <%
		  rs.MoveNext()
		  Wend
	      
		  rs.Close()
		  Set rs = Nothing
  		  'FecharConexao2()
		  FecharConexao()
		  %>
		   <tr>
            <td colspan="8"><div align="center">
			<%If VERIFICA_FINALIZACAO = 1 and PRODUCAO_FINALIZADA = 2 then%>
              <input class="btnFinalizaProducao" onclick="confirmaFinalizarProducao()" type="button" name="Submit22" value=" " />
                
             <%Elseif VERIFICA_FINALIZACAO = 2 then%>
			  <input class="btnAvancar" type="submit" name="Submit2" value=" " />
            	<%End If%>


			</div></td>
            <td colspan="2">
			<%If PRODUCAO_FINALIZADA = 2 Then%>
			<div align="center"><a href="javascript:excluirRegistrosemMassa()">Excluir Registros Selecionados </a></div>
			<%End If%>			</td>
            </tr>
        </table>
		  </form>


		  <div align="center">
		    <p>
		      <%else%>
		      <em>Não há formando importados</em>!</p>
		    <p><a href="dip_producoes_importar_formandos.asp?cod_prod=<%=Request.QueryString("cod_prod")%>">Deseja importar agora?</a></p>
		    <p> </p>
		    <p>	      
		      <%end if%>
		      </p>
		  </div>
		  Observações:<br />
		  Digito 0 = Formando importado da formatura<br />
		  Digito 1 = Formando cadastrado pelo depto. de foto
		  <p align="center">
		  <input class="btnTopo" onclick="abre('#')" type="button" name="btnTopo2" value=" " />
		</p>		</td>
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



