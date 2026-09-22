<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<%
'Valida somente o usuário
COD_USUARIO_INTERNO_LOGADO 	= Session("COD_USUARIO")
COD_PERFIL_LOGADO			= Session("COD_PERFIL")

'Verifica se o usuário está logado
if COD_USUARIO_INTERNO_LOGADO=Empty OR COD_PERFIL_LOGADO=Empty then
	Session.Abandon()
	Response.Redirect("index.asp")		
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
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
	var perguntaq = confirm('EXCLUIR PRODUÇÃO:\nAtenção caso queira continuar lembre-se que todas as fichas(se houver) serão excluídas completamente, deseja continuar?')
		if (perguntaq==true){
			var perguntadeNovo = confirm("CONFIRMA EXCLUSÃO DESTA PRODUÇÃO E SUAS RESPECTIVAS FICHAS?")
			if (perguntadeNovo==true){
			abre('?cod_prod=<%=Request.QueryString("cod_prod")%>&comando=excluir_producao')
			}
		}
	
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
	
	  <h2>Controle de Projetos Pendentes</h2>
	  <p>&nbsp;</p>
	  <table width="100%" border="1" cellpadding="5" cellspacing="0">
              <col width="64" />
              <col width="473" />
              <col width="251" />
              <col width="124" span="2" />
              <col width="101" />
              <col width="197" />
              <col width="123" />
              <tr height="17" bgcolor="#CCCCCC">
                <td height="20" width="38"><div align="center"><strong>ORDEM</strong></div></td>
                <td width="329"><div align="center"><strong>PROJETO</strong></div></td>
                <td width="96"><div align="center"><strong>HIST&Oacute;RICO</strong></div></td>
                <td width="73"><div align="center"><strong>TEMPO</strong></div></td>
                <td width="50"><div align="center"><strong>STATUS</strong></div></td>
                <td width="64"><div align="center"><strong>% FINALIZADO</strong></div></td>
                <td width="39"><div align="center"><strong>DATA &Iacute;NICIO</strong></div></td>
                <td width="77"><div align="center"><strong>DATA &Uacute;LT.    MODIFICA&Ccedil;&Atilde;O</strong></div></td>
                <td width="54"><div align="center">
                  <div align="center"><strong>DATA T&Eacute;RMINO</strong></div>
                  <strong></strong></div></td>
              </tr>
              <!--
              <tr height="17">
                <td height="17" width="38"><div align="center"></div></td>
                <td width="202">Manuten&ccedil;&atilde;o completa no sistema</td>
                <td width="174">Reestrutura&ccedil;&atilde;o de todas as p&aacute;ginas para possibilitar o cadastro de usu&aacute;rios no sistema, perfis de acesso, al&eacute;m de outras configura&ccedil;&otilde;es parametrizadas tamb&eacute;m. Foi realizado uma manuten&ccedil;&atilde;o em todas as p&aacute;ginas do sistema melhorando a seguran&ccedil;a exclu&iacute;ndo todas as p&aacute;ginas que estavam inutilizadas.</td>
                <td width="56"><p align="center">&nbsp;</p>                  </td>
                <td><div align="center">Finalizado (30 horas)</div></td>
                <td><div align="center">100%</div></td>
                <td>30/06/2011</td>
                <td><div align="center">-</div></td>
                <td><div align="center">29/07/2011</div></td>
              </tr>
              
              <tr height="17">
                <td height="17" width="38"><div align="center">1</div></td>
                <td width="202">Reestrutura&ccedil;&atilde;o    Produ&ccedil;&atilde;o Sa&iacute;da. </td>
                <td width="174">15/06.<br />
                  Solicitado diversas altera&ccedil;&otilde;es e relat&oacute;rios no projeto, prazo alterado e informado via e-mail<br /><br />17/06<br />
                  Solicitado por &Eacute;lcio o cadastrado do motivo da devolução ou renovação online, para execu&ccedil;&atilde;o ser&aacute; necess&aacute;rio acr&eacute;scimo de 6 dias &uacute;teis.</td>
                <td width="56"><p align="center"><a href="projetos_pendentes_detalhes.asp?cod_projeto=1">Exibir detalhes</a></p>                  </td>
                <td><div align="center">Finalizado</div></td>
                <td><div align="center">100%</div></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
             
              <tr height="17">
                <td height="17" width="38"><div align="center">2</div></td>
                <td width="202">Boletim Informativo - Cadastro de Pre&ccedil;os nas produ&ccedil;&otilde;es e bloqueio de finaliza&ccedil;&atilde;o de pedidos, cujos produtos estiverem com pre&ccedil;os de venda inferiores ao acordado no boletim informativo</td>
                <td width="174">&nbsp;</td>
                <td width="56">&nbsp;</td>
                <td><div align="center">Finalizado</div></td>
                <td>&nbsp;</td>
                <td>13/06/2011</td>
                <td><div align="center">10/10/2011</div></td>
                <td>&nbsp;</td>
              </tr>
              <tr height="17">
                <td height="17" width="38"><div align="center">3</div></td>
                <td width="202">Relat&oacute;rio    por per&iacute;odo de todos produtos vendidos</td>
                <td width="174">&nbsp;</td>
                <td width="56">&nbsp;</td>
                <td><div align="center">Finalizado</div></td>
                <td>&nbsp;</td>
                <td>08/06/2011</td>
                <td><div align="center">13/06/2011</div></td>
                <td>&nbsp;</td>
              </tr>
               -->
              <tr height="51">
                <td height="51" width="38"><div align="center">*</div></td>
                <td width="329">Estoque    interno de produtos (Ex.: Qtd. De foto-telas compradas, fotos compradas) com    entrada de Nota Fiscal e baixa de estoque somente se vendido</td>
                <td width="96"><div align="center"></div></td>
                <td width="73"><div align="center">N&atilde;o efetuada an&aacute;lise</div></td>
                <td><div align="center">N&atilde;o aprovado</div></td>
                <td><div align="center"></div></td>
                <td>&nbsp;</td>
                <td><div align="center"></div></td>
                <td><div align="center"></div></td>
              </tr>
              
              <tr height="34">
                <td height="34" width="38"><div align="center">*</div></td>
                <td width="329">Cart&atilde;o    de cr&eacute;dito confrontar com extrato de operadora se os valores foram realmente    creditados</td>
                <td width="96"><div align="center"></div></td>
                <td width="73"><div align="center">N&atilde;o efetuada an&aacute;lise</div></td>
                <td><div align="center">N&atilde;o aprovado</div></td>
                <td><div align="center"></div></td>
                <td>&nbsp;</td>
                <td><div align="center"></div></td>
                <td><div align="center"></div></td>
              </tr>
              <tr height="34">
                <td height="34" width="38"><div align="center">*</div></td>
                <td width="329">M&oacute;dulo    controle de cheques, autentica&ccedil;&atilde;o do n&uacute;mero do cheque. Sa&iacute;da para fornecedor,    SPC controle, entre outros.</td>
                <td width="96"><div align="center"></div></td>
                <td width="73"><div align="center">N&atilde;o efetuada an&aacute;lise</div></td>
                <td><div align="center">N&atilde;o aprovado</div></td>
                <td><div align="center"></div></td>
                <td>&nbsp;</td>
                <td><div align="center"></div></td>
                <td><div align="center"></div></td>
              </tr>
              <!--
              <tr height="17">
                <td height="17" width="38"><div align="center">*</div></td>
                <td width="202">Modulo    venda DVD's - Gerencial Foto</td>
                <td width="174">&nbsp;</td>
                <td width="56">&nbsp;</td>
                <td><div align="center">Finalizado</div></td>
                <td><div align="center"></div></td>
                <td>&nbsp;</td>
                <td><div align="center"></div></td>
                <td><div align="center"></div></td>
              </tr>
              -->
              <tr height="17">
                <td height="17" width="38"><div align="center">*</div></td>
                <td width="329">Desmembrar    Encaderna&ccedil;&atilde;o do &aacute;lbum</td>
                <td width="96"><div align="center"></div></td>
                <td width="73"><div align="center">N&atilde;o efetuada an&aacute;lise</div></td>
                <td><div align="center">N&atilde;o aprovado</div></td>
                <td><div align="center"></div></td>
                <td>&nbsp;</td>
                <td><div align="center"></div></td>
                <td><div align="center"></div></td>
              </tr>
              <!--
              <tr height="17">
                <td height="17" width="38"><div align="center">*</div></td>
                <td width="202">Calend&aacute;rio    de vendas - Ex.: Maio de 2011 de 01 &agrave; 01 de junho</td>
                <td width="174">&nbsp;</td>
                <td width="56">&nbsp;</td>
                <td><div align="center">Finalizado (8,5)</div></td>
                <td><div align="center"></div></td>
                <td>&nbsp;</td>
                <td><div align="center"></div></td>
                <td><div align="center"></div></td>
              </tr>
              
              <tr height="17">
                <td height="17" width="38"><div align="center">*</div></td>
                <td width="202">Altera&ccedil;&atilde;o    foto - Melhoramento de Relat&oacute;rio de devolu&ccedil;&atilde;o e sa&iacute;da</td>
                <td width="174">&nbsp;</td>
                <td width="56">&nbsp;</td>
                <td><div align="center">Finalizado (6,5 horas)</div></td>
                <td><div align="center"></div></td>
                <td>&nbsp;</td>
                <td><div align="center"></div></td>
                <td><div align="center"></div></td>
              </tr>
              -->
              <tr height="17">
                <td height="17" width="38"><div align="center">*</div></td>
                <td width="329">Modulo    Produ&ccedil;&atilde;o Inicio - Gerencial Foto - Controle desde a solicita&ccedil;&atilde;o dos profissionais para o evento at&eacute; pagamento de Freelancer</td>
                <td width="96"><div align="center"></div></td>
                <td width="73"><div align="center">N&atilde;o efetuada an&aacute;lise</div></td>
                <td><div align="center">N&atilde;o aprovado</div></td>
                <td><div align="center"></div></td>
                <td>&nbsp;</td>
                <td><div align="center"></div></td>
                <td><div align="center"></div></td>
              </tr>
              <!--
              <tr height="51">
                <td height="51" width="38"><div align="center">*</div></td>
                <td width="202">Boletim    informativo com valores individuais por contrato. Verificar melhor o cr&eacute;dito    da formatura para autentica&ccedil;&atilde;o da diretoria. Ordem de Cr&eacute;dito para formatura    da cobran&ccedil;a</td>
                <td width="174">&nbsp;</td>
                <td width="56">&nbsp;</td>
                <td><div align="center">Finalizado (15 horas)</div></td>
                <td><div align="center"></div></td>
                <td>&nbsp;</td>
                <td><div align="center"></div></td>
                <td><div align="center"></div></td>
              </tr>
              <tr height="34">
                <td height="34" width="38"><div align="center">*</div></td>
                <td width="202">Rotina    sistema foto - limpa pedidos n&atilde;o finalizados pelos vendedores de 2 dias para    tr&aacute;s</td>
                <td width="174">&nbsp;</td>
                <td width="56">&nbsp;</td>
                <td><div align="center">Finalizado</div></td>
                <td><div align="center"></div></td>
                <td>&nbsp;</td>
                <td><div align="center"></div></td>
                <td><div align="center"></div></td>
              </tr>
              <tr height="85">
                <td height="85" width="38"><div align="center">*</div></td>
                <td width="202">Sistema    Foto / Alterar parte de sa&iacute;da do estoque para Stillo&rsquo;s trocar para Outro,    este outro dever&aacute; ser digitado o nome da Outra pessoa que est&aacute; retirando o    &aacute;lbum e motivo que o mesmo esta saindo para esta pessoa. Exemplo de    utiliza&ccedil;&atilde;o disso quando algum vendedor da formatura solicita o &aacute;lbum</td>
                <td width="174">&nbsp;</td>
                <td width="56">&nbsp;</td>
                <td><div align="center">Finalizado</div></td>
                <td><div align="center"></div></td>
                <td>&nbsp;</td>
                <td><div align="center"></div></td>
                <td><div align="center"></div></td>
              </tr>
              -->
              <tr height="17">
                <td height="17" width="38"><div align="center">*</div></td>
                <td width="329">Troca    de mensagens online entre Vendedores e Ligia/Almerinda.</td>
                <td width="96"><div align="center"></div></td>
                <td width="73"><div align="center">N&atilde;o efetuada an&aacute;lise</div></td>
                <td><div align="center">N&atilde;o aprovado</div></td>
                <td><div align="center"></div></td>
                <td>&nbsp;</td>
                <td><div align="center"></div></td>
                <td><div align="center"></div></td>
              </tr>
              <tr height="17">
                <td height="17"><div align="center">*</div></td>
                <td><p>Relat&oacute;rio de vendas (mesmo antes de finalizar o m&ecirc;s) para  Ligia e Almerinda utilizarem e terem o controle de quantas vendas j&aacute; foram  conferidas antes do pagamento realizado<br />
                </p></td>
                <td><div align="center">Solicitado em 15/06/11</div></td>
                <td><div align="center">N&atilde;o efetuada an&aacute;lise</div></td>
                <td><div align="center">N&atilde;o aprovado</div></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <!--
              <tr height="17">
                <td height="17"><div align="center">*</div></td>
                <td> Relat&oacute;rio  com layout diferenciado em colunas Ex.: Nome do Vendedor, Qtd de &aacute;lbuns  vendidos, qtd. de DVDs vendidos, qtd de Posters vendidos etc...</td>
                <td>Solicitado em 15/06/11</td>
                <td>&nbsp;</td>
                <td><div align="center">Finalizado (10 horas)</div></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
            
              
              <tr height="17">
                <td height="17"><div align="center">*</div></td>
                <td>Relat&oacute;rio Extrato de Produ&ccedil;&atilde;o (mesmo sem corre&ccedil;&atilde;o), disponibilizar a quantidade de lotes que ainda n&atilde;o foram retirados</td>
                <td>Solicitado em 16/06/11</td>
                <td>&nbsp;</td>
                <td><div align="center">Finalizado (4 horas)</div></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr height="17">
                <td height="17"><div align="center">*</div></td>
                <td>Cria&ccedil;&atilde;o de m&oacute;dulo para controle e organiza&ccedil;&atilde;o de contratos somente foto, que n&atilde;o existe contrato na formatura. Haver&aacute; altera&ccedil;&atilde;o desde o lan&ccedil;amento do formando at&eacute; a etiqueta que ter&aacute; uma marca&ccedil;&atilde;o especial informando que &eacute; um contrato somente foto para n&atilde;o ocacionar duplicidade de numera&ccedil;&atilde;o com os contratos da formatura.</td>
                <td>Solicitado em 22/06/11</td>
                <td>&nbsp;</td>
                <td><div align="center">Finalizado (12 horas)</div></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td><div align="center">09/08/2011</div></td>
              </tr>
              <tr height="17">
                <td height="17"><div align="center">*</div></td>
                <td>Relat&oacute;rio Pr&eacute;vio com os Rankings de vendas</td>
                <td>Solicitado em 30/06/11</td>
                <td>&nbsp;</td>
                <td><div align="center">Finalizado<br />(2 horas)</div></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td><div align="center">30/06/2011</div></td>
              </tr>
              <tr height="17">
                <td height="17"><div align="center">*</div></td>
                <td>Revis&atilde;o e reestrutura&ccedil;&atilde;o da transfer&ecirc;ncia de Fichas de Identifica&ccedil;&atilde;o, agora ser&aacute; necess&aacute;rio lan&ccedil;ar o prazo de cada item se o formando estiver cadastrado com lote. Desta forma o novo vendedor que estar&aacute; recebendo dever&aacute; prestar conta depois. O recibo foi alterado tamb&eacute;m para informar o prazo individualmente dos itens</td>
                <td>Solicitado em 20/06/11</td>
                <td>&nbsp;</td>
                <td><div align="center">Finalizado<br />
                  (6 horas)</div></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td><div align="center">01/07/2011</div></td>
              </tr>
              <tr height="17">
                <td height="17"><div align="center">*</div></td>
                <td>Listagem de pedidos por vendedor, para os pr&oacute;ximos pagamentos.</td>
                <td>Solicitado em 30/06/11</td>
                <td>&nbsp;</td>
                <td><div align="center">Finalizado<br />
                  (1,5 horas)</div></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <!--
              <tr height="17">
                <td height="17">&nbsp;</td>
                <td><p><br />
                  Controle de Projetos melhorar telas para eles possam ter acesso o que est&aacute; ocorrendo nos projetos</p>
                  <p>&Aacute;rea do Vendedor al&eacute;m do motivo melhorar visualiza&ccedil;&atilde;o das telas e incluir outros recursos como: Comunica&ccedil;&atilde;o entre Stillos/Consultores e Busca de Dados de Formandos</p>
                  <p>Relat&oacute;rios na tela inicial da D. Andrea / &Eacute;lcio como vendas, devolu&ccedil;&atilde;o, resumo de tudo que &eacute; relevante para diretoria etc.</p>
                  <p>Relat&oacute;rio de Cart&atilde;o de Cr&eacute;dito como existe para o Cheque (1 hora)</p>
                  <p>Verificar como est&aacute; sendo armazenado no estoque os produtos e se precisa de alguma melhoria.</p>
                  <p>Manual completo do sistema (3 dias - 18 horas)</p>
                  <p> Login com seguran&ccedil;a adicional at&eacute; 3 tentativas depois &eacute; bloqueado o acesso.</p>                  </td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
             
              <tr height="17">
                <td height="17">&nbsp;</td>
                <td>Sistema de Aprova&ccedil;&otilde;es para diversos casos como: Finaliza&ccedil;&atilde;o do Fechamento; Vendas com pre&ccedil;os dos produtos inferiores ao Boletim Informativo, Exclus&atilde;o de Formandos ap&oacute;s finalizada produ&ccedil;&atilde;o, entre outras aprova&ccedil;&otilde;es que possam surgir</td>
                <td>Solicitado em 16/08/11</td>
                <td>&nbsp;</td>
                <td><div align="center">Finalizado<br />
                  (10 horas)
                </div></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
                 -->
              <tr height="17">
                <td height="17"><div align="center">*</div></td>
                <td>Relat&oacute;rio de Resumo de Problemas (verificar quais os principais problemas e melhor layout quando for desenvolver)</td>
                <td><div align="center">Solicitado em 16/08/11</div></td>
                <td><div align="center">N&atilde;o efetuada an&aacute;lise</div></td>
                <td><div align="center">N&atilde;o aprovado</div></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <!--
              <tr height="17">
                <td height="17">&nbsp;</td>
                <td>Relat&oacute;rio Mensal Consolidado de Todos os itens vendidos<br />
                  Neste relat&oacute;rio dever&aacute; constar o que foi vendido no m&ecirc;s em produtos e tamb&eacute;m um detalhamento por item.</td>
                <td>Solicitado em 20/09/11</td>
                <td>&nbsp;</td>
                <td><div align="center">Finalizado<br />(4,5 horas)</div></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>27/09/11</td>
              </tr>
              -->
              <tr height="17">
                <td height="17"><div align="center">*</div></td>
                <td>Ligia/Luciana Felix solicitam a inclus&atilde;o da forma de pagamento dep&oacute;sito nos pedidos e no recibo</td>
                <td><div align="center">Solicitado em 09/12/11</div></td>
                <td><div align="center">N&atilde;o efetuada an&aacute;lise</div></td>
                <td><div align="center">N&atilde;o aprovado</div></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <!--
              <tr height="17">
                <td height="17"><div align="center">*</div></td>
                <td><p>M&oacute;dulo novo Procedimento para Pagamento de Consultores</p>
                  <p>Novo procedimento dever&aacute; ser criado um calend&aacute;rio para lan&ccedil;amento  das vendas onde &eacute; definido o per&iacute;odo de vendas no m&ecirc;s. Podendo dividir um m&ecirc;s em  2 pagamentos por exemplo.</p>
                  <p>O sistema dever&aacute; calcular o dia do pagamento conforme os par&acirc;metros  da venda como data venda, data do lan&ccedil;amento, valor de entrada do pedido, entre  outros.</p>
                  <p>Com o novo m&oacute;dulo n&atilde;o ser&aacute; permitido (exceto com autoriza&ccedil;&atilde;o)  a gera&ccedil;&atilde;o do pagamento sem antes finalizar a corre&ccedil;&atilde;o da diretoria tamb&eacute;m.</p>
                  <p>Ser&aacute; realizado uma reestrutura&ccedil;&atilde;o completa no sistema envolvendo  todas as telas onde aparece a data de pagamento do consultar antes era M&Ecirc;S/ANO agora  ser&aacute; DIA/M&Ecirc;S/ANO. Outro ponto &eacute; a gera&ccedil;&atilde;o dos recibos que ser&aacute; totalmente diferenciada.</p></td>
                <td>Solicitado em 14/11/11</td>
                <td>&nbsp;</td>
                <td><div align="center">Finalizado (120 horas)</div></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
             
              <tr height="17">
                <td height="17">&nbsp;</td>
                <td>Habilitar todas as p&aacute;ginas do sistema para o  perfil Diretoria </td>
                <td>Solicitado em 27/02/12</td>
                <td>&nbsp;</td>
                <td>Finalizado</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr height="17">
                <td height="17">&nbsp;</td>
                <td>Formando 291725-0 foi renovado como teste,  precisa voltar ao normal sem renova&ccedil;&atilde;o </td>
                <td>Solicitado em 27/02/12</td>
                <td>&nbsp;</td>
                <td>Finalizado</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr height="17">
                <td height="17">&nbsp;</td>
                <td>Erro ao finalizar a corre&ccedil;&atilde;o pela Diretoria de  todas as fichas ainda o fechamento fica pendente. Verifiquei e n&atilde;o detectei o  erro, pois em todos os casos havia fichas pendentes. Se o problema persistir  por favor me informar</td>
                <td>Solicitado em 27/02/12</td>
                <td>&nbsp;</td>
                <td>Finalizado</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr height="17">
                <td height="17">&nbsp;</td>
                <td>Resumo Folha Geral pr&eacute;via (antes de processar o pagamento)</td>
                <td>Solicitado em 27/02/12</td>
                <td>&nbsp;</td>
                <td><p>Finalizo em 03/03 (2,5 horas)</p></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
             
              <tr height="17">
                <td height="17"><div align="center">*</div></td>
                <td>M&oacute;dulo Controle de Produ&ccedil;&atilde;o de Produtos Vendidos<br />
                  <br /></td>
                <td>Solicitado em 27/02/12</td>
                <td><div align="center">N&atilde;o efetuada an&aacute;lise</div></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
                -->
              <tr height="17">
                <td height="17"><div align="center">*</div></td>
                <td><p>Adapta&ccedil;&atilde;o para inclus&atilde;o de desconto em venda para Comiss&atilde;o de Formatura. Esta altera&ccedil;&atilde;o implica em mudan&ccedil;a nos seguintes Formul&aacute;rios:<br />
                  1) Lan&ccedil;amento do vendedor para mostrar o valor do desconto referente a comiss&atilde;o de formatura no pedido total. Quando for mais que um formando no mesmo pedido e se os 2 tiverem desconto ir&aacute; somar os descontos. Revisar regras com diretoria<br />
                2) Corre&ccedil;&atilde;o de pedido com informa&ccedil;&atilde;o do desconto no final da venda para abater no valor total da venda<br />
                3) Recibo de Comiss&atilde;o de Formatura e lan&ccedil;amento de desconto no cadastro da produ&ccedil;&atilde;o<br />
                4) Impress&atilde;o do Recebibo com Campo &quot;Desconto Valor da Venda&quot;<br />
                </p>                  </td>
                <td><div align="center">Solicitado em 03/07/12</div></td>
                <td><div align="center">4,0 horas t&eacute;cnicas 
                  (com desconto de 2 horas t&eacute;cnicas devido a atraso de projeto 76/2012</div></td>
                <td><div align="center">N&atilde;o aprovado</div></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr height="17">
                <td height="17">&nbsp;</td>
                <td>Solicita&ccedil;&atilde;o de Ligia para Recibo de Vendedor de Entrega de Vendas:<br />
                  <br />
                  <div>&quot;O recibo &eacute; emitido ao vendedor no ato da entrega da   venda, por&eacute;m s&oacute; se pode conferir a venda ap&oacute;s imprimir este extrato, s&oacute; que as   vezes o vendedor faz o lan&ccedil;amento incorreto e gera valores que n&atilde;o os reais que   ir&atilde;o ficar para comissionamento.</div>
                    <div>&nbsp;</div>
                    <div>Gostaria que o extrato s&oacute; fosse emitido ap&oacute;s a   finaliza&ccedil;&atilde;o da corre&ccedil;&atilde;o dos caixas e que l&aacute; pudesse mostrar qual ser&aacute; o valor   estimado de comiss&atilde;o de cada venda entregue e conferida.</div>
                    <div>&nbsp;</div>
                  <div>Qualquer d&uacute;vida a disposi&ccedil;&atilde;o.&quot;</div></td><td><div align="center">Solicitado em<br />
                    03/07/12</div></td>
                <td><div align="center">1,5 horas t&eacute;cnicas</div></td>
                <td><div align="center">N&atilde;o aprovado</div></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr height="17">
                <td height="17"><div align="center">*</div></td>
                <td>Bloqueio de Produ&ccedil;&atilde;o &agrave; partir de 01/08 para novas produ&ccedil;&otilde;es dos vendedores que n&atilde;o prestarem contas e renovarem a permiss&atilde;o de fichas<br />
                  <br />
                  Sistema ir&aacute; verificar se houver mais que 1 item em aberto sem justificativa n&atilde;o ir&aacute; permitir a sele&ccedil;&atilde;o do vendedor para distribui&ccedil;&atilde;o de lotes.</td>
                <td><div align="center">Solicitado em<br />
                  02/07/12</div></td>
                <td><div align="center">1,5 horas t&eacute;cnicas</div></td>
                <td><div align="center">N&atilde;o aprovado</div></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr height="17">
                <td height="17"><div align="center">*</div></td>
                <td>Devolu&ccedil;&atilde;o do &aacute;lbum, somente junto com a ficha de identifica&ccedil;&atilde;o. Confirmar informa&ccedil;&atilde;o com D. Andrea</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
            </table>
		    <p align="center">&nbsp;</p>
		<p>
	  </p>
	  		<p align="center">&nbsp;</p>
		<p align="center">&nbsp;</p>
		<p align="center">&nbsp;</p>
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