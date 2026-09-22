<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/classe_fechamento.asp" -->

<%
Dim VALIDA_SUB46 'Verifica se o usuário possui acesso a somente listagem ou a correção mesmo
VALIDA_SUB46 = formatNumber(0&inStr(sub_validacoes_paginas,"#46"),0)


Dim cod_fechamento_periodo, COD_VENDEDOR, NOME_VENDEDOR
Dim rs, rs2, rs3, rsNomeVend, sql, sqlBase

Dim PRODUTOS, DINHEIRO, CARTAO, CHEQUE
Dim SUB_TOTAL_EM_PRODUTOS, SUB_TOTAL_EM_DINHEIRO, SUB_TOTAL_EM_CARTAO, SUB_TOTAL_EM_CHEQUE
Dim TOTAL_EM_PRODUTOS, TOTAL_EM_DINHEIRO, TOTAL_EM_CARTAO, TOTAL_EM_CHEQUE
Dim valor_estimado_comissao, temp

TOTAL_EM_PRODUTOS	= 0
TOTAL_EM_DINHEIRO	= 0
TOTAL_EM_CARTAO		= 0
TOTAL_EM_CHEQUE		= 0

COD_VENDEDOR	= Request.QueryString("COD_VENDEDOR")
cod_fechamento_periodo = Request.QueryString("cod_fechamento_periodo")


AbrirConexao()

	Set rsNomeVend = objConexao.Execute("SELECT NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR="&COD_VENDEDOR)
	
		If not rsNomeVend.Eof Then
				
				NOME_VENDEDOR = rsNomeVend("NOME_VENDEDOR")
				
		End If
	
	rsNomeVend.Close
	Set rsNomeVend = Nothing
	
FecharConexao()

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
Dim data_inicio, data_termino, data_termino_lancamento_vendas, data_pagamento, autorizacao_diretoria_para_finalizar

AbrirConexao()

	Set rs=objConexao.Execute("Select * FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE cod_fechamento_periodo = "&cod_fechamento_periodo)
		If not rs.Eof Then
			data_inicio = ArrumaData(rs("data_inicio"))
			data_termino = ArrumaData(rs("data_termino"))
			data_termino_lancamento_vendas = ArrumaData(rs("data_termino_lancamento_vendas"))
			data_pagamento = ArrumaData(rs("data_pagamento"))
			autorizacao_diretoria_para_finalizar = Cstr(""&rs("autorizacao_diretoria_para_finalizar"))
		End If
	rs.Close
	Set rs=Nothing

FecharConexao()
'////////////////////////////////////////////////////////////////////////////////////

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Versão 1.2</title>
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
	<script language="javascript">
	function remote(COD_PEDIDO){
/*Credit: JavaScript Kit www.javascriptkit.com more JavaScripts here.*/
win=window.open("lca_correcao_diretoria_autenticacao_POPUP_lista_erro.asp?cod_pedido="+COD_PEDIDO,"","width=329,height=180,scrollbars=no")
//interceptacao de erro na abertura da janela
text = "Se a janela não abrir desbloquei os Pop-ups!";
if(win == null) { alert(text); return; }
//fim
win.moveTo(350,150);
//win2.creator=self;
}
	
function mOvr21(src, COMB)
			{
			
			e = document.vendas[COMB].checked;
			//if (!e) {
			
			if (!src.contains(event.fromElement))
					{
						//src.style.cursor = 'hand';
						src.bgColor = '#FFCC66';
						//src.style.color = '#000000'
					}
			
			
			}
		function mOut21(src, cor, COMB)
			{
				
				e = document.vendas[COMB].checked
				if (!e) {
					if (!src.contains(event.toElement))
					{
						src.style.cursor = 'default';
						src.bgColor = '#FFFFFF'
						//src.style.color = '#CC0000'
					}
				}else{
					src.bgColor = '#CDFCC5'
				}
			}

function mOut212 (src) {
			if (!src.contains(event.fromElement))
					{
						src.style.cursor = 'hand';
						src.bgColor = '#CCCCCC';
						//src.style.color = '#000000'
					}
					
}

function mOut213 (src) {
		
			if (!src.contains(event.toElement))
					{
						src.style.cursor = 'default';
						src.bgColor = '#FCFBC2'
						//src.style.color = '#CC0000'
					}

}
function mClk(COMB, src)
			{    
				//if(event.srcElement.tagName=='TD')
				//	{
				//		src.children.tags('A')[0].click(); 
				//	}!d.SEMESTRE[0].checked &&
			d = document.vendas[COMB]
			
			if (d.checked) {	
			document.vendas[COMB].checked = false;
			//src.bgColor = '#cccccc';		
			}
			else {
			document.vendas[COMB].checked = 'checked';	
			}
			
}

function verifica(COMBO,COMBO2){
	
	//alert(COMBO.checked)
		if (COMBO.checked) {
			
			COMBO.checked = false;
		//	document.vendas[COMBO2].checked = false;
			//COMBO.checked = true;
			//alert('ok')
		}else {
		
		COMBO.checked = 'checked';
		}	
}
function verifica2(COMBO,COMBO2){
	
	//alert(COMBO.checked)
		if (COMBO) {
			
			document.vendas[COMBO2].checked = false;
			//COMBO.checked = true;
			//alert('ok')
		}	
}

function mClk2(src, url)
			{    
				//if(event.srcElement.tagName=='TD')
				//	{
				//		src.children.tags('A')[0].click(); 
				//	}
			location.href = url;

			}
			
	</script>
	
	<script type="text/javascript" language="javascript">
<!--
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

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->

function cadastraVendedor(valor){
	alert(valor);
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
</head>

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

	  <h2>LCA - Correção Diretoria / Mensal </h2>
      
      <%If VALIDA_SUB46 = 0 Then%>
      <br /><br />
      <div class="msgGrande">Você possui acesso somente a consulta nesta página.</div>
      <%End If%>
      
	   <!--<div class="subMenu">Sub-Menu: <a href="filtro_baixas_pag_erros.asp">Verificar fichas marcadas com erros</a> | <a href="relatorios_fichas_baixadas.asp?filtro1=+">Últimas fichas baixadas</a> </div>-->
	   <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
		
						<h3 align="center">Fichas pendentes para o pagamento   do(a) vendedor(a) <%=NOME_VENDEDOR%></h3>
                        <table width="100%" 
		      border="1" align="center"
					cellpadding="3" 
					style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
                <td width="33%"><div align="center"><strong>Per&iacute;odo de Vendas</strong><br />
                    <em>Data da Venda</em></div></td>
                <td width="36%"><div align="center"><strong>Data M&aacute;xima Permita para o Lan&ccedil;amento</strong> <strong>no sistema</strong></div></td>
                <td width="31%"><div align="center"><strong>Data Prevista do Pagamento</strong></div></td>
                </tr>
<tr>
  <td class="textoMedio"><div align="center"><%=data_inicio%> à <%=data_termino%></div></td>
  <td><div align="center"><%=data_termino_lancamento_vendas%></div></td>
  <td><div align="center"><%=data_pagamento%></div></td>
  </tr>
           
           
                        </table>
                        <br />
<br />

                        
                        	<form name="vendas" method="post" action="lca_correcao_diretoria_autenticacao_confirmacao.asp?COD_VENDEDOR=<%=COD_VENDEDOR%>&cod_fechamento_periodo=<%=cod_fechamento_periodo%>">
          <%
		  
	  sqlBase = " SELECT TB_PEDIDOS.COD_PEDIDO, TB_PEDIDOS.CONFERIDO_POR_COD_USUARIO, TB_PEDIDOS.PEDIDO_EXTRA, TB_PEDIDOS.DATA_VENDA, TB_PEDIDOS.ERRO_LISTA, TB_PEDIDOS_FORMANDOS.COD_FORMANDO, "&_
	 			" TB_PEDIDOS.FINALIZA_DIRETORIA, TB_PEDIDOS.DATA_CORRIGIDO_DIRETORIA, "&_
			" TB_FICHAS_IDENTIFICACAO.COD_CONTRATO "&_
				" FROM TB_PEDIDOS "&_
				" LEFT JOIN TB_PEDIDOS_FORMANDOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_FORMANDOS.COD_PEDIDO "&_
				" LEFT JOIN TB_FICHAS_IDENTIFICACAO ON TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO WHERE "&_
				" TB_PEDIDOS.COD_VENDEDOR="&COD_VENDEDOR&" AND TB_PEDIDOS.cod_fechamento_periodo= "&cod_fechamento_periodo&" AND TB_PEDIDOS.DATA_DESCEU_DIRETORIA IS NOT NULL GROUP BY TB_PEDIDOS.COD_PEDIDO ORDER BY TB_FICHAS_IDENTIFICACAO.COD_CONTRATO"
		
		
		  AbrirConexao()
		  'AbrirConexao2()
		  'sql = sqlBase&" GROUP BY TB_FICHAS_IDENTIFICACAO.COD_CONTRATO ORDER BY COD_CONTRATO"
		  'Response.Write(sql)
		  sql = " SELECT * FROM TB_PEDIDOS WHERE "&_
		        " COD_VENDEDOR="&COD_VENDEDOR&" AND TB_PEDIDOS.cod_fechamento_periodo = '"&cod_fechamento_periodo&"' AND DATA_DESCEU_DIRETORIA IS NOT NULL"
		 
		 'response.write(sql)
		  
		  		SUB_TOTAL_EM_DINHEIRO	= 0
				SUB_TOTAL_EM_CARTAO		= 0
				SUB_TOTAL_EM_CHEQUE		= 0
	  
				  
			'Busca o nome do primerO
			Set rs2 = objConexao.Execute(sqlBase)%>
					
					
					<table width="100%" 
		      border="1" align="center"
					cellpadding="5" bgcolor="#FFFFFF" 
					style="border-collapse:collapse">
  <tr bgcolor="#DEEBFE">   
    <td width="4%"> </td>
		      <td width="4%" class="style6"><div align="center"><strong>Erro </strong></div></td>
		      <td width="6%"><div align="center"><strong>Cód. Pedido </strong></div></td>
		      <td width="19%"><div align="center"><strong>Instituição de Ensino </strong></div></td>
		      <td width="4%"><div align="center"><strong>Álbum</strong></div></td>
		      <td width="17%"><div align="center" class="style6"><strong>Formando</strong></div></td>
		      <td width="4%"><div align="center" class="style6"><strong>Inic.</strong></div></td>
		      <td width="5%"><div align="center" class="style6"><strong>Fatur.</strong></div></td>
		      <td width="5%"><div align="center" class="style6"><strong>Retir.</strong></div></td>
		      <td width="8%"><div align="center" class="style6"><strong> Dinheiro </strong></div></td>
		      <td width="7%"><div align="center" class="style6"><strong>Cartão </strong></div></td>
		      <td width="9%"><div align="center" class="style6"><strong>Cheque </strong></div></td>
		      <td width="9%"><div align="center" class="style6"><strong>Boleto </strong></div></td>
		      <td width="4%"><div align="center" class="style6"><strong>Total  Recebido </strong></div></td>
	          <td width="2%"><div align="center"><strong>Conferido por</strong></div></td>
	          <td width="2%"><div align="center"><strong>Valor comissão <a href="javascript:alert('Valor estimado se o fechamento fosse hoje e não houvesse alterações no pedido e na tabela de comissionamento.\nObs.: Valor poderá ser alterado.')"><img src="images/icons/icon_help_mini.jpg" width="15" height="11" border="0" /></a> </strong></div></td>
  </tr>
		
			<%
			
			
			Dim FINALIZA_DIRETORIA_RS
			While not rs2.Eof
			FINALIZA_DIRETORIA_RS = formatNumber(rs2("FINALIZA_DIRETORIA"))
			%>
			
	
			

  <tr id="cor_<%=rs2("COD_PEDIDO")%>" style="color:#<%If isNull(rs2("ERRO_LISTA")) Then Response.Write("000000") else Response.Write("FF0000")end if%>" <%if FINALIZA_DIRETORIA_RS = 2 then%>
  onMouseOver="mOvr21(this,'COMBO_<%=rs2("COD_PEDIDO")%>')" onMouseOut="mOut21(this,'#ffffff','COMBO_<%=rs2("COD_PEDIDO")%>')"
  <%end if%>>
    <td id="borderBaixo">
	
	  <div align="center">
      
	    <%
	if FINALIZA_DIRETORIA_RS = 1 then%>
	       <img height="20" src="http://www.stillos.com.br/painel/imagens/green.png" width="23" alt="<%=rs2("DATA_CORRIGIDO_DIRETORIA")%>">
	        <%else%>
            	<%If VALIDA_SUB46 <> 0 Then%>
	        <input class="inputRadio" id="COMBO_<%=rs2("COD_PEDIDO")%>" type="checkbox" name="COMBO_<%=rs2("COD_PEDIDO")%>" value="1">
       			 <%Else%>
                  -
      			<%End If%>          	
			<%end if%>	

	    </div></td>
    <td id="borderBaixo">
    <div align="center">
    	<%If VALIDA_SUB46 <> 0 Then%>
    	<a href="javascript:remote(<%=rs2("COD_PEDIDO")%>)"><img border="0" src="images/icons/lista.png" width="26" height="29"></a>
      	<input name="teste_<%=rs2("COD_PEDIDO")%>" type="hidden" size="2" maxlength="2">
      	<%Else%>
         - 
        <%End If%>
      
      </div></td>
    <td id="borderBaixo"><div align="center"><a href="lca_busca_detalhamento_pedido.asp?cod_pedido=<%=rs2("COD_PEDIDO")%>"><%=rs2("COD_PEDIDO")%><%If int(rs2("PEDIDO_EXTRA")) = 1 Then%><br />
(Pedido Extra)<%End If%></a></div></td>
    <td colspan="6" id="borderBaixo"><div align="left">
			<%
			sql = 	"SELECT TB_FICHAS_IDENTIFICACAO.INICIAL, TB_FICHAS_IDENTIFICACAO.FATURADO, TB_FICHAS_IDENTIFICACAO.RETIRADO, "&_
					"TB_FICHAS_IDENTIFICACAO.NOME_CURSO, TB_PEDIDOS_FORMANDOS.COD_FORMANDO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, TB_FICHAS_IDENTIFICACAO.ALBUM,  TB_FICHAS_IDENTIFICACAO.COD_CONTRATO "&_
					"FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON "&_
					"TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO "&_
					"WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&rs2("COD_PEDIDO")&" ORDER BY NOME_FORMANDO"
					
					'Response.Write(sql)
			Set rs3 = objConexao.Execute(sql)
				
		
				While not rs3.Eof
				
				%>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				  <tr>
					<!--<td width="23%"><div align="center"><%=digito(rs3("COD_FORMANDO"))%></div></td>-->
					<td width="35%"><div align="left">
					<%
					Set rsNomeFacul = objConexao.Execute("SELECT DESCRICAO FROM TB_ENTIDADES WHERE COD_CONTRATO = '"&rs3("COD_CONTRATO")&"'")
							
							If not rsNomeFacul.Eof Then
								Response.Write(rsNomeFacul("DESCRICAO"))
							Else
								Response.Write(rs3("COD_CONTRATO"))
							End If
							
					
					rsNomeFacul.Close
					Set rsNomeFacul = Nothing
					
					%></div></td>
					<td width="8%"><%=rs3("ALBUM")%></td>
					<td width="34%">					  
					  
					    
					    <div align="left"><%=rs3("NOME_FORMANDO")%></div>					    <div align="left"></div></td>
					<td width="7%"><div align="center"><%=rs3("INICIAL")%></div></td>
				          <td width="9%"><div align="center"><%=rs3("FATURADO")%></div></td>
				          <td width="7%"><div align="center"><%=rs3("RETIRADO")%></div></td>
				  </tr>
				</table>
				<%
				rs3.MoveNext()
				Wend
				
			rs3.Close
			Set rs3 = Nothing
			%>
			</div></td>
    <td id="borderBaixo"><div align="right">
     <%
   Set rsPag = objConexao.Execute("SELECT * FROM VW_PEDIDOS_TOTAL_RECEBIDO WHERE COD_PEDIDO="&rs2("COD_PEDIDO"))
   			
			If not rsPag.Eof Then
				
				DINHEIRO	= formatNumber(0&rsPag("PAGAMENTO_EM_DINHEIRO"))/1
				CARTAO		= formatNumber(0&rsPag("XTOTAL_CARTAO"))/1
				CHEQUE		= formatNumber(0&rsPag("XTOTAL_CHEQUE"))/1
				BOLETO		= formatNumber(0&rsPag("XTOTAL_BOLETO"))/1
				
				SUB_TOTAL_EM_DINHEIRO	= SUB_TOTAL_EM_DINHEIRO+DINHEIRO
				SUB_TOTAL_EM_CARTAO		= SUB_TOTAL_EM_CARTAO+CARTAO
				SUB_TOTAL_EM_CHEQUE		= SUB_TOTAL_EM_CHEQUE+CHEQUE
				SUB_TOTAL_EM_BOLETO		= SUB_TOTAL_EM_BOLETO+BOLETO

				TOTAL_EM_DINHEIRO		= TOTAL_EM_DINHEIRO+DINHEIRO
				TOTAL_EM_CARTAO			= TOTAL_EM_CARTAO+CARTAO
				TOTAL_EM_CHEQUE			= TOTAL_EM_CHEQUE+CHEQUE
				TOTAL_EM_BOLETO			= TOTAL_EM_BOLETO+BOLETO
				
			End If
			
   rsPag.Close
   Set rsPag = Nothing
   %>
   <%=formatNumber(DINHEIRO)%>
   </div></td>
   <td id="borderBaixo"><div align="right"><%=formatNumber(CARTAO)%></div></td>
   <td id="borderBaixo"><div align="right"><%=formatNumber(CHEQUE)%></div></td>
   <td id="borderBaixo"><div align="right"><%=formatNumber(CHEQUE)%></div></td>
   <td id="borderBaixo"><div align="right"><%=formatNumber(DINHEIRO+CARTAO+CHEQUE+BOLETO)%></div></td>
   <td id="borderBaixo"><div align="center">
   	<% 				
	
	Set rs4=objConexao.Execute("SELECT NOME FROM TB_ADM_USUARIOS WHERE COD_USUARIO = "&sqlTexto(rs2("CONFERIDO_POR_COD_USUARIO")))	

	If not rs4.Eof Then
		Response.Write(Left(rs4("nome"),5))
	Else
		Response.Write("Não disp.")
	End If
	rs4.Close
	Set rs4=Nothing
	
	%>
   </div></td>
   <td id="borderBaixo"><div align="right">
     <%
   temp = formatNumber(0&fechamentoRetornaValorComissaoPedido(rs2("cod_pedido"),1,true))
   Response.Write(temp)
   valor_estimado_comissao = (valor_estimado_comissao+temp)/1
   %>
   </div></td>
  </tr>
  		
            <%
			rs2.MoveNext()
		    Wend
			
			rs2.Close
			Set rs2=Nothing%>
		   <tr>
    <td colspan="9" bgcolor="#E8E8E8"> </td>
    <td bgcolor="#E8E8E8"><div align="right"><%=formatNumber(SUB_TOTAL_EM_DINHEIRO)%></div></td>
    <td bgcolor="#E8E8E8"><div align="right"><%=formatNumber(SUB_TOTAL_EM_CARTAO)%></div></td>
    <td bgcolor="#E8E8E8"><div align="right"><%=formatNumber(SUB_TOTAL_EM_CHEQUE)%></div></td>
    <td bgcolor="#E8E8E8"><div align="right"><%=formatNumber(SUB_TOTAL_EM_BOLETO)%></div></td>
    <td bgcolor="#E8E8E8"><div align="right"><%=formatNumber(SUB_TOTAL_EM_DINHEIRO+SUB_TOTAL_EM_CARTAO+SUB_TOTAL_EM_CHEQUE+SUB_TOTAL_EM_BOLETO)%></div></td>
    <td bgcolor="#E8E8E8"> </td>
    <td bgcolor="#E8E8E8"><div align="right"><%=formatNumber(valor_estimado_comissao)%></div></td>
		   </tr>
	      </table>

		          <div align="center">
		<%
		  
		    
		  
		  
		  
		  
		  'FecharConexao2()
		  FecharConexao()
		  
		  %> 
                      <br />
                      <%If VALIDA_SUB46 <> 0 Then%>
		              <input class="btnAvancar" type="submit" name="Submit" value=" " />
                      <%End If%>
		              <br /><br />
		              <br />
          </div>
		            <table style="border:solid #CCCCCC 1px" width="100%" border="0" align="left" cellpadding="3" cellspacing="0"><tr>
      <td bgcolor="#E7EFFE" id="borderBaixo"><div align="left"><strong>Total Recebido:</strong> </div></td>
            <td bgcolor="#E7EFFE"><h3 align="left" style="margin:0px"><%=formatNumber(TOTAL_EM_DINHEIRO+TOTAL_EM_CHEQUE+TOTAL_EM_CARTAO,2)%></h3>			</td>
            </tr>
		
	<tr>
	      <td width="12%" bgcolor="#E7EFFE">
        <div align="left"><strong>    <em>-Total  em Dinheiro:</em></strong></div></td>
            <td width="88%" bgcolor="#E7EFFE">
              <div align="left"><strong> <%=formatNumber(TOTAL_EM_DINHEIRO,2)%> </strong></div>
          </td>
            </tr>

		<tr>
	
      <td width="12%" bgcolor="#E7EFFE">
        <div align="left"><strong>    <em>-Total  em Cartão: </em></strong></div>
     </td>
            <td width="88%" bgcolor="#E7EFFE">
              <div align="left"><strong> <%=formatNumber(TOTAL_EM_CARTAO,2)%> </strong></div>
           </td>
            </tr>
		<tr>
	
      <td width="12%" bgcolor="#E7EFFE">
        <div align="left"><strong>    <em>-Total  em Cheque:</em> </strong></div>
     </td>
            <td width="88%" bgcolor="#E7EFFE">
              <div align="left"><strong> <%=formatNumber(TOTAL_EM_CHEQUE,2)%> </strong></div>
    
			</td>
            </tr>
    </table>

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



