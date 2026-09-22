<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<%

Function inArray(needle, haystack)
  dim i 
  inArray = false
  for i = 0 to Ubound(haystack)
    if haystack(i) = needle then
      inArray = true
      exit function
    end If
  next
End Function

Dim COD_PRODUCAO, comando
COD_PRODUCAO = Request.QueryString("cod_prod")

If Request.QueryString("comando") = "efetuar_alteracao" Then

	AbrirConexao()
	Dim PRAZO_DIAS_CORRIDOS_INICIAL, COD_VENDEDOR_INICIAL

  Set rs=objConexao.Execute("SELECT * FROM VW_TB_PRODUCAO_LOTES WHERE COD_PRODUCAO = "&COD_PRODUCAO)
  
    If not rs.Eof Then
      While not rs.Eof
        PRAZO_DIAS_CORRIDOS_INICIAL	= Request.Form("PRAZO_DIAS_CORRIDOS_INICIAL_"&rs("LOTE"))
        COD_VENDEDOR_INICIAL = int(Request.Form("COD_VENDEDOR_INICIAL_"&rs("LOTE")))
		
        If int(COD_VENDEDOR_INICIAL) > 0 Then
		
          objConexao.Execute("UPDATE TB_PRODUCAO_LOTES SET COD_VENDEDOR_INICIAL = "&sqlNumero(COD_VENDEDOR_INICIAL)&",PRAZO_DIAS_CORRIDOS_INICIAL = '"&PRAZO_DIAS_CORRIDOS_INICIAL&"' WHERE COD_PRODUCAO ="&COD_PRODUCAO&" AND LOTE = "&rs("LOTE")&" AND (COD_VENDEDOR_INICIAL IS NULL or COD_VENDEDOR_INICIAL = 0)LIMIT 1")
        End If
      rs.MoveNext()
      Wend
    End If
  rs.Close
  Set rs=Nothing
  
  Response.Redirect("dip_producoes_fichas_lotes.asp?cod_prod="&COD_PRODUCAO&"&msg=Lotes separados para vendedores com sucesso!")	
	FecharConexao()

ElseIf Request.QueryString("comando") = "excluir_de_lote" Then
	
	AbrirConexao()
	
		PRAZO_DIAS_CORRIDOS_INICIAL	= Request.Form("PRAZO_DIAS_CORRIDOS_INICIAL_"&Request.QueryString("lote"))
		COD_VENDEDOR_INICIAL = formatNumber(0&Request.Form("COD_VENDEDOR_INICIAL_"&Request.QueryString("lote")),0)
				
		'Atualiza somente se a data de saída inicial for null.	
		objConexao.Execute("UPDATE TB_PRODUCAO_LOTES SET COD_VENDEDOR_INICIAL = NULL,PRAZO_DIAS_CORRIDOS_INICIAL = 0 WHERE COD_PRODUCAO ="&COD_PRODUCAO&" AND LOTE = "&Request.QueryString("lote")&" AND DATA_SAIDA_INICIAL IS NULL LIMIT 1")
					
		
		Response.Redirect("dip_producoes_fichas_lotes.asp?cod_prod="&COD_PRODUCAO&"&msg=Vendedor excluído do lote com sucesso!")
	
	FecharConexao()


End If

Function selectPadraoVendedores()

Dim retorno, albuns_atrasados, autoriza_bloqueio, diferenca_dias
Dim rsVendedor, rs, i, ii, temp, temp2
Dim vendedores_desbloqueados, vendedores_desbloqueados_producoes, vendedor_liberado_manualmente
	
'Solicitação dia 20/09/13
vendedores_desbloqueados_producoes = array(397,398,400)
vendedores_desbloqueados = array(1,4,7,8,10,11,12,14,16,17,18,19,21,23,27,37,41,47,48,75,96,97,98,104,105,106,108,109,110,111)
'data_atual = date()
'diferenca_dias = DateDiff("d",date(),"2013-12-20")
'autoriza_bloqueio = false
'If diferenca_dias < 0 Then
'	autoriza_bloqueio = true
'End If
	
autoriza_bloqueio = true 
AbrirConexao()

set dicVendedoresPrazos = Server.CreateObject("Scripting.Dictionary")

Set rs2=objConexao.Execute("SELECT COUNT(*) AS ALBUNS_PENDENTES, COD_VENDEDOR FROM VW_MOVIMENTACOES_COMPLETA WHERE TIPO NOT IN (1, 2) AND DATA_RETORNO IS NULL AND DIAS_ATRASADOS > 0 AND COD_VENDEDOR IN(SELECT COD_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR <> 1 AND ATIVADO = 1 AND LOTE <> 1) GROUP BY  COD_VENDEDOR")
while not rs2.Eof
  call dicVendedoresPrazos.add(rs2("COD_VENDEDOR").value, rs2("ALBUNS_PENDENTES").value)
rs2.MoveNext()
wend						
rs2.Close
Set rs2=Nothing


Set rsVendedor = objConexao.Execute("SELECT COD_VENDEDOR, NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR <> 1 AND ATIVADO = 1 AND LOTE <> 1 ORDER BY NOME_VENDEDOR")

retorno = "<option selected='selected' value='0'>Selecione um vendedor para este lote</option>"
While not rsVendedor.Eof

  if (dicVendedoresPrazos.exists(rsVendedor("COD_VENDEDOR").value) ) then 
    albuns_atrasados = int(dicVendedoresPrazos(rsVendedor("COD_VENDEDOR").value)&"")
  else 
    albuns_atrasados = 0
  end if
  
  vendedor_liberado_manualmente = false

  '/Verifica exceções
  if (inArray(int(cod_producao), vendedores_desbloqueados_producoes)) then 
    if (inArray(int(rsVendedor("COD_VENDEDOR").value), vendedores_desbloqueados) ) then
      vendedor_liberado_manualmente = true
    end if
  end if

  if albuns_atrasados = 0 or Not autoriza_bloqueio or vendedor_liberado_manualmente then
    'Disponível para novas produções.
    retorno = retorno & "<option value='"&rsVendedor("COD_VENDEDOR")&"'>"&Ucase(TrocarAcento(rsVendedor("NOME_VENDEDOR")))
    if vendedor_liberado_manualmente Then
      retorno = retorno & " (LIBERADO MANUALMENTE)"
    end if
    retorno = retorno &"</option>"
  else
    retorno = retorno & "<option disabled='disabled' value='"&rsVendedor("COD_VENDEDOR")&"'>"&Ucase(TrocarAcento(rsVendedor("NOME_VENDEDOR")))&" (BLOQUEADO QTD.: "&albuns_atrasados&")</option>"    
  end If

  rsVendedor.MoveNext()
  wend
rsVendedor.close()
set rsVendedor = nothing

FecharConexao()

selectPadraoVendedores = retorno	
end function

Dim selectPadraoVendedoresVariavel
selectPadraoVendedoresVariavel = selectPadraoVendedores()
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
<title>Lotes dos Vendedores</title>
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
function confirmaFinalizado(){
	var mensagem = "TEM CERTEZA QUE DESEJA EFETUAR A SEPARAÇÃO DE LOTES PARA OS VENDEDORES SELECIONADOS?\nnão SERÁ POSSÍVEL ALTERAR DEPOIS O LOTE INICIAL PARA CADA VENDEDOR."
	if(YNconfirm(mensagem)){
		document.formulario_padrao.submit();
	}
}

function confirmaSaidaporVendedor(nomeVendedor,cod_vendedor_inicial, lote){
	var mensagem = "TEM CERTEZA QUE DESEJA IMPRIMIR O RECIBO DE saída DO VENDEDOR(A) "+nomeVendedor+"?\nSE não TIVER DATA DE saída O SISTEMA IRÁ COLOCAR A DATA E HORA ATUAL."
	if(YNconfirm(mensagem)){
		var caminho = "dip_producoes_fichas_lotes_IMPRIMIR_recibo.asp?cod_producao=<%=COD_PRODUCAO%>&cod_vendedor_inicial="+cod_vendedor_inicial+"&lote="+lote;
		openInNewWindow(caminho);
		setTimeout("confirmaImpressao()",2000);
		
	}
}

function confirmaImpressao (){
	//if(YNconfirm('Impressão do recibo efetuada?')){
		window.location.reload();
	//}
}

function confirmaExclusaoVendedor(nomeVendedor,cod_vendedor_inicial, lote){
	var mensagem = "TEM CERTEZA QUE DESEJA EXCLUIR ESTE VENDEDOR(A) "+nomeVendedor+" DO LOTE "+lote+"?\nObs.: É possível fazer isso somente se o recibo não estiver gerado, depois deverá ser feito somente através de transferência."
	if(YNconfirm(mensagem)){
		var caminho = "dip_producoes_fichas_lotes.asp?comando=excluir_de_lote&cod_prod=<%=COD_PRODUCAO%>&cod_vendedor_inicial="+cod_vendedor_inicial+"&lote="+lote;
		abre(caminho);
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
	
	
	function openInNewWindow(url) {
	// Change "_blank" to something like "newWindow" to load all links in the same new window
	var newWindow = window.open(url, '_blank');
	newWindow.focus();
	return false;
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
<!--INICIO VALIDA��O SIM/não-->
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
<!--FIM VALIDA��O SIM/não-->	
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

	  <h2>DIP - produções / Separação por Lotes</h2>
	  <%
	  Dim rs9, PRODUCAO_FINALIZADA, DESCRICAO, VERIFICA_FINALIZACAO
	  
	  VERIFICA_FINALIZACAO = 1
	  
	  AbrirConexao()
			  Set rs9 = objConexao.Execute("SELECT PRODUCAO_FINALIZADA, DESCRICAO FROM TB_PRODUCAO WHERE COD_PRODUCAO="&COD_PRODUCAO)

					If not rs9.Eof then
					
						PRODUCAO_FINALIZADA = rs9("PRODUCAO_FINALIZADA")
						DESCRICAO = rs9("DESCRICAO") 
					End If					

			  rs9.Close()
	  		  Set rs9 = Nothing 

	  %>	 	   
		   
	       <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
		  <h3 align="center">Lotes da produção:<br />
		    <%=COD_PRODUCAO%> - <%=DESCRICAO%></h3>
<%
Set rs=objConexao.Execute("SELECT * FROM VW_TB_PRODUCAO_LOTES WHERE COD_PRODUCAO = "&COD_PRODUCAO)

			If not rs.Eof Then
				
%>
		  <form name="formulario_padrao" action="dip_producoes_fichas_lotes.asp?cod_prod=<%=COD_PRODUCAO%>&comando=efetuar_alteracao" method="post" style="margin:0px;">
          <table width="98%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="14%"><div align="center"><strong>Lote</strong></div></td>
            <td width="13%"><div align="center"><strong>Qtd. de Formandos</strong></div></td>
            <td width="31%"><div align="center"><strong> Vendedor</strong></div></td>
            <td width="14%"><div align="center"><strong>Data de saída</strong></div></td>
            <td width="15%"><div align="center"><strong>Prazo em dias corridos</strong></div></td>
            <td width="13%"><div align="center"><strong>Imprimir Recibo de saída</strong></div></td>
</tr>
            <%While not rs.Eof%>
            <tr>
              <td width="14%"><div align="center" class="textoMaior"><strong><%=rs("LOTE")%></strong></div></td>
              <td width="13%"><div align="center"><%=rs("QTD_FORMANDOS_INICIAL")%></div></td>
              <td width="31%">
              <div align="left">
              <%If int(rs("COD_VENDEDOR_INICIAL")) > 0 Then%>
			  	<%=rs("NOME_VENDEDOR")%>
              	<%If IsNull(rs("DATA_SAIDA_INICIAL")) Then %>
              	<a href="javascript:confirmaExclusaoVendedor('<%=rs("nome_vendedor")%>',<%=rs("cod_vendedor_inicial")%>,'<%=rs("LOTE")%>');"><img border="0" src="images/icons/icon_excluir.png" /></a>
                <%End If%>
                
              <%Else%>
              		<select name="COD_VENDEDOR_INICIAL_<%=rs("LOTE")%>" style="font-size:12px">
						<%=selectPadraoVendedoresVariavel%>	
					</select>
              <%End If%>
              </div>              </td>
              <td width="14%"><div align="center"><%=ArrumaData(rs("DATA_SAIDA_INICIAL"))%></div></td>
              <td width="15%">
                            <div align="center">
                              <%
			  If int(rs("COD_VENDEDOR_INICIAL")) > 0 Then
			  
			  				If rs("DATA_SAIDA_INICIAL") <> Empty Then
			  %>
                                  <%=rs("PRAZO_DIAS_CORRIDOS_INICIAL")%> (Até <%=DateAdd("d",rs("PRAZO_DIAS_CORRIDOS_INICIAL"),rs("DATA_SAIDA_INICIAL"))%>)
                            <%Else%>
                            
                            <%=rs("PRAZO_DIAS_CORRIDOS_INICIAL")%>
                            
                            <%End If%>      
                                  
                              <%Else%>
                                  <input type="text" value="30" class="inputTamanho1" name="PRAZO_DIAS_CORRIDOS_INICIAL_<%=rs("LOTE")%>" />
<%End If%>
                                          </div></td>
              <td width="13%"><div align="center">
              
              <%
			  If int(rs("COD_VENDEDOR_INICIAL")) > 0 Then
			  %>
    		<a href="javascript:confirmaSaidaporVendedor('<%=rs("nome_vendedor")%>',<%=rs("cod_vendedor_inicial")%>,'<%=rs("LOTE")%>');">
     
              <img border="0" src="images/icons/icon_print.jpg" width="32" height="28" />              </a>
              <%Else%>
              Falta vendedor
              <a href="javascript:alert('Para poder imprimir é necessário selecionar o vendedor e gravar.')"><img src="images/icons/icon_help_mini.jpg" width="15" height="11" border="0" /></a>
              <%End If%>
              </div></td>
            </tr>
			<%
			rs.MoveNext()
			wend
			%>
		
		   <tr>
            <td colspan="6"><div align="center">

			  <input class="btnGravar" type="button" onclick="return confirmaFinalizado()" name="Submit2" value=" " />


			</div></td>
            </tr>
            <tr>
              <td colspan="6"><div align="center">
                <p></p>
                <p><a href="dip_producoes_fichas_recibos_vendedores_comissao_IMPRIMIR.asp?cod_producao=<%=COD_PRODUCAO%>" target="_blank" class="linkMaster">Imprimir comprovantes de comissão de formatura</a></p>
              </div></td>
            </tr>
        </table>
        </form>
		<br />
     <strong>   <span class="textoMedio">Observações:<br />
     </span></strong><span class="textoMedio"> Se você selecionar um vendedor para um lote e clicar em avançar, automaticamente todos os formandos deste lote serão da responsabilidade do vendedor.<br />
     Este recibo deverá ser utilizado somente a liberação inicial.<br />
     É possível alterar o vendedor do lote desde que o recibo não seja alterado. </span>
     <div align="center">
		    <p>
		      <%else%>
              <br /><br /><br />
		      <em>não há lotes lançados para esta produção</em>  
               <br /><br /><br />   
		      <%end if%>
		      </p>
		  </div>
<%


fecharConexao()
%>
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



