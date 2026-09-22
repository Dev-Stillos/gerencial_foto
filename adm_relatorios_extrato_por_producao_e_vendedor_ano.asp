<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Session.Timeout=5
'Validação especial para verificar se possui sub-validação específica
'AbrirConexao()
	
'	Set rs=objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '31' and cod_perfil = '"&Session("cod_perfil")&"'")
		
'		If rs.Eof Then
		'Se não encontrar então redireciona
'			Session.Abandon()
''			Response.Redirect("index.asp")
'		End If
		
'	rs.Close
'	Set rs=Nothing
	
'FecharConexao()


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
	function popUp(URL){
	win=window.open(URL,"","width=550,height=350,scrollbars=yes")
	var text = "Se a janela não abrir desbloquei os Pop-ups!";
	if(win == null) { alert(text); return; }
	win.moveTo(350,150);
}	
	</script>
    <style type="text/css">
<!--
.style1 {color: #006600}
.style2 {color: #CC0000}
.style3 {font-size: 10px}
-->
.text-center{
	text-align: center;
}
.text-right{
	text-align: right;
}
.text-left{
	text-align: left;
}
    </style>
</head>

<body>
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

							  Dim INICIO, TERMINO, DATA_INICIO, DATA_TERMINO
													DATA_INICIO = Request.QueryString("DATA_INICIO")
													DATA_INICIO = Split(DATA_INICIO,"/")
													INICIO = DATA_INICIO(2) & "-" & DATA_INICIO(1) & "-" & DATA_INICIO(0)
													DATA_TERMINO = Request.QueryString("DATA_TERMINO")
													DATA_TERMINO = Split(DATA_TERMINO,"/")
													TERMINO = DATA_TERMINO(2) & "-" & DATA_TERMINO(1) & "-" & DATA_TERMINO(0)
							%>
							  <!-- #include file="includes/sub_menu.asp" -->
						  
								<table width="100%" border="0" cellspacing="0" cellpadding="5" style="border:dotted 1px #999999">
									  <tr>
										<td valign="top" >
										  <h2>Relatórios - Saída de album por periodo (<%=arrumaData(INICIO) %> à <%=arrumaData(TERMINO)%>)<br /><br /></h2>
										  <div align="center" class="msg"><%=Request.QueryString("msg")%></div>
										</td>			
									   </tr>
									<tr>
										<td valign="top">										  
											<table width="90%" border="1" align="center" cellpadding="5"  style="border-collapse:collapse">
												<thead>
													<tr bgcolor="whitesmoke">
														<th class="text-left"> Nome </th>
														<th class="text-center">Qtd. Retirada </th>
														<th class="text-center">Qtd. Vendida </th>
														<th class="text-center">Total Vendido</th>												
														<th class="text-center">% Aproveitamento</th>
														<th class="text-center">Dvd</th>
														<th class="text-center">Trio</th>
														<th class="text-center">Slide Show</th>
														<th class="text-center">Foto Tela</th>
														<th class="text-center">Revista</th>												
														<th class="text-center">Mini Album </th>
														<th class="text-center">Porta retrato Digital</th>
														<th class="text-center">Bolsa</th>
														<th class="text-center">Nº de álbuns Total/período</th>
														<th class="text-center">Valor Arrecadado total/período</th>

														
														

													</tr>
												</thead>
												<tbody>
												  <%
													AbrirConexao()		
													Dim count
													Dim rs_principal,rs_1,rs_2,rs_3,rs_4,rs_5, sql,sql_1,sql_2,sql_3,sql_4,sql_5, i

													'response.write(INICIO & "--" & TERMINO)
													'INICIO = "2015-01-01"
													'TERMINO = "2015-02-23"

													i = 0
													sql = "SELECT NOME_VENDEDOR, TB_VENDEDORES.COD_VENDEDOR, SUM(QTD_FORMANDOS_INICIAL) AS QTD_RETIRADO "&_
														  "FROM TB_PRODUCAO_LOTES INNER JOIN TB_VENDEDORES ON TB_VENDEDORES.COD_VENDEDOR = TB_PRODUCAO_LOTES.COD_VENDEDOR_INICIAL "&_
														  "WHERE  DATA_SAIDA_INICIAL >= '"&INICIO&"' and  DATA_SAIDA_INICIAL <= '"&TERMINO&"'  AND TB_VENDEDORES.COD_VENDEDOR <> 27 GROUP BY NOME_VENDEDOR"
														  
													'sql_1 = "SELECT sum(TB_PEDIDOS_PRODUTOS.VALOR_UN * TB_PEDIDOS_PRODUTOS.QUANTIDADE) as valor "&_
													'		"FROM TB_PEDIDOS_FORMANDOS "&_
													'		"INNER JOIN TB_PEDIDOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_FORMANDOS.COD_PEDIDO "&_
													'		"INNER JOIN TB_PRODUCAO_LOTES_FORMANDOS ON TB_PRODUCAO_LOTES_FORMANDOS.COD_FORMANDO = TB_PEDIDOS_FORMANDOS.COD_FORMANDO "&_		
													'		"INNER JOIN TB_PEDIDOS_PRODUTOS ON TB_PEDIDOS_PRODUTOS.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO "&_
													'		"WHERE TB_PEDIDOS.FINALIZA_CORRECAO = 1 AND TB_PRODUCAO_LOTES_FORMANDOS.DATA_SAIDA > '"&INICIO&"' and TB_PRODUCAO_LOTES_FORMANDOS.DATA_SAIDA < '" &TERMINO&"' "&_
													'		"AND TB_PRODUCAO_LOTES_FORMANDOS.COD_VENDEDOR = {}"
															
													sql_1 = "SELECT (SUM(IFNULL(PAGAMENTO_EM_DINHEIRO,0)) + SUM(IFNULL(XTOTAL_CHEQUE,0))+ SUM(IFNULL(XTOTAL_BOLETO,0)) + SUM(IFNULL(XTOTAL_CARTAO,0))) AS VALOR_VENDIDO FROM VW_EXTRATO_TOTAL_RECEBIDO_COMPLETO WHERE DATA_VENDA BETWEEN '"&INICIO&" 0:0:0' AND '"&TERMINO&" 0:0:0' AND COD_PEDIDO IN ( SELECT COD_PEDIDO FROM TB_PEDIDOS_FORMANDOS WHERE COD_FORMANDO IN (SELECT COD_FORMANDO FROM TB_PRODUCAO_LOTES_FORMANDOS WHERE DATA_SAIDA BETWEEN '"&INICIO&" 0:0:0' AND '"&TERMINO&" 0:0:0' AND TB_PRODUCAO_LOTES_FORMANDOS.COD_VENDEDOR = COD_VENDEDOR)) AND COD_VENDEDOR = {}"													
															
													'response.write(sql_1)
													'sql_2 = "SELECT COUNT(*) AS QTD_VENDIDO,sum(TB_FICHAS_IDENTIFICACAO.FATURADO * TB_FICHAS_IDENTIFICACAO.VALOR_UN) as valor "&_
													'	  "FROM TB_PRODUCAO_LOTES_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = TB_PRODUCAO_LOTES_FORMANDOS.COD_FORMANDO "&_
													'	  "WHERE  DATA_SAIDA >= '"&INICIO&"' and  DATA_SAIDA <= '"&TERMINO&"' AND VENDIDO = 1 AND TB_PRODUCAO_LOTES_FORMANDOS.COD_VENDEDOR = {}"
														  
													sql_2 = "SELECT sum(VW_QTD_FORMANDOS_POR_PEDIDO.QTD_FORMANDOS) AS QTD_VENDIDO FROM VW_QTD_FORMANDOS_POR_PEDIDO INNER JOIN TB_PEDIDOS ON VW_QTD_FORMANDOS_POR_PEDIDO.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO  WHERE DATA_VENDA BETWEEN '"&INICIO&" 0:0:0' AND '"&TERMINO&" 0:0:0' AND TB_PEDIDOS.COD_PEDIDO IN ( SELECT COD_PEDIDO FROM TB_PEDIDOS_FORMANDOS WHERE COD_FORMANDO IN (SELECT COD_FORMANDO FROM TB_PRODUCAO_LOTES_FORMANDOS WHERE DATA_SAIDA BETWEEN '"&INICIO&" 0:0:0' AND '"&TERMINO&" 0:0:0'))  AND FINALIZA_CORRECAO = 1 AND  TB_PEDIDOS.COD_VENDEDOR = {} "												  
														  
													sql_3 = "SELECT TB_PRODUTOS.COD_PRODUTO, NOME_PRODUTO, SUM(QUANTIDADE) as QUANTIDADE "&_
															"FROM TB_PEDIDOS_FORMANDOS "&_
															"INNER JOIN TB_PEDIDOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_FORMANDOS.COD_PEDIDO  "&_
															"INNER JOIN TB_PRODUCAO_LOTES_FORMANDOS ON TB_PRODUCAO_LOTES_FORMANDOS.COD_FORMANDO = TB_PEDIDOS_FORMANDOS.COD_FORMANDO  "&_
															"INNER JOIN TB_PEDIDOS_PRODUTOS ON TB_PEDIDOS_PRODUTOS.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO "&_
															"INNER JOIN TB_PRODUTOS ON TB_PRODUTOS.COD_PRODUTO = TB_PEDIDOS_PRODUTOS.COD_PRODUTO "&_
															"WHERE TB_PEDIDOS.DATA_VENDA IS NOT NULL AND TB_PRODUCAO_LOTES_FORMANDOS.DATA_SAIDA >= '"&INICIO&"' "&_
															"AND TB_PRODUCAO_LOTES_FORMANDOS.DATA_SAIDA <= '"&TERMINO&"' AND TB_PRODUTOS.COD_PRODUTO IN (41,35,34,1,16,31,7,9) "&_
															"AND TB_PRODUCAO_LOTES_FORMANDOS.COD_VENDEDOR = {} "&_													
															" GROUP BY TB_PRODUTOS.COD_PRODUTO, NOME_PRODUTO ORDER BY TB_PRODUTOS.COD_PRODUTO"
															
													sql_4 = "SELECT (SUM(IFNULL(PAGAMENTO_EM_DINHEIRO,0)) + SUM(IFNULL(XTOTAL_CHEQUE,0)) + SUM(IFNULL(XTOTAL_BOLETO,0)) + SUM(IFNULL(XTOTAL_CARTAO,0))) AS TOTAL_RECEBIDO_GERAL FROM VW_EXTRATO_TOTAL_RECEBIDO_COMPLETO WHERE DATA_VENDA BETWEEN '"&INICIO&" 0:0:0' AND '"&TERMINO&" 0:0:0' AND COD_VENDEDOR = {}"
													
													sql_5 = "SELECT SUM( VW_QTD_FORMANDOS_POR_PEDIDO.QTD_FORMANDOS ) AS TOTAL_ALBUNS_GERAL FROM VW_QTD_FORMANDOS_POR_PEDIDO INNER JOIN TB_PEDIDOS ON VW_QTD_FORMANDOS_POR_PEDIDO.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO WHERE TB_PEDIDOS.DATA_VENDA BETWEEN '"&INICIO&" 0:0:0' AND '"&TERMINO&" 0:0:0' AND FINALIZA_CORRECAO =1 AND COD_VENDEDOR = {};"
													
													
															
													Dim QTD_RETIRADO,NOME_VENDEDOR,COD_VENDEDOR,QTD_VENDIDO,APROVEITAMENTO, COD_FORMANDO_LIST, VALOR_VENDIDO,TOTAL_RECEBIDO_GERAL,TOTAL_ALBUNS_GERAL
													
													Set rs_principal=objConexao.Execute(sql)													
														while not rs_principal.Eof 						
															COD_VENDEDOR = rs_principal("COD_VENDEDOR")
															NOME_VENDEDOR = rs_principal("NOME_VENDEDOR")
															QTD_RETIRADO = cint(rs_principal("QTD_RETIRADO"))

															VALOR_VENDIDO = 0
															Set rs_1=objConexao.Execute(replace(sql_1,"{}",COD_VENDEDOR))													
															if not rs_1.Eof	then	
																if rs_1("VALOR_VENDIDO") <> "null" then
																	VALOR_VENDIDO = cdbl(rs_1("VALOR_VENDIDO"))
																end if
															End if
															
															QTD_VENDIDO = 0
															APROVEITAMENTO = 0
															
															Set rs_2=objConexao.Execute(replace(sql_2,"{}",COD_VENDEDOR))													
															if not rs_2.Eof	then					
																if rs_2("QTD_VENDIDO") <> "null" then
																	QTD_VENDIDO = cdbl(rs_2("QTD_VENDIDO"))
																end if
															End if	

															Set rs_4=objConexao.Execute(replace(sql_4,"{}",COD_VENDEDOR))													
															if not rs_4.Eof	then	
																TOTAL_RECEBIDO_GERAL = 0
																if rs_4("TOTAL_RECEBIDO_GERAL") <> "null" then
																	TOTAL_RECEBIDO_GERAL = cdbl(rs_4("TOTAL_RECEBIDO_GERAL"))
																end if
															End if	

															Set rs_5=objConexao.Execute(replace(sql_5,"{}",COD_VENDEDOR))													
															if not rs_5.Eof	then
																TOTAL_ALBUNS_GERAL = 0													
																if rs_5("TOTAL_ALBUNS_GERAL") <> "null" then
																	TOTAL_ALBUNS_GERAL = cdbl(rs_5("TOTAL_ALBUNS_GERAL"))
																end if
															End if	

															reDim produtos(7)

															Set rs_3=objConexao.Execute(replace(sql_3,"{}",COD_VENDEDOR))				
															while not rs_3.Eof 	
															
																if rs_3("COD_PRODUTO") = 1 then'DVD
																	produtos(0) = rs_3("QUANTIDADE")	
																end if
																if rs_3("COD_PRODUTO") = 7 then'Trio
																	produtos(1) = rs_3("QUANTIDADE")
																end if
																if rs_3("COD_PRODUTO") = 9 then'Slide
																	produtos(2) = rs_3("QUANTIDADE")
																end if
																if rs_3("COD_PRODUTO") = 16 then'Foto Tela
																	produtos(3) = rs_3("QUANTIDADE")
																end if
																if rs_3("COD_PRODUTO") = 31 then'revista
																	produtos(4) = rs_3("QUANTIDADE")
																end if
																if rs_3("COD_PRODUTO") = 34 then'mini Album
																	produtos(5) = rs_3("QUANTIDADE")
																end if	
																if rs_3("COD_PRODUTO") = 35 then'porta retrato
																	produtos(6) = rs_3("QUANTIDADE")
																end if		
																if rs_3("COD_PRODUTO") = 41 then'bolsa
																	produtos(7) = rs_3("QUANTIDADE")
																end if																
																
																rs_3.MoveNext()
															wend	
															
															APROVEITAMENTO = QTD_VENDIDO * 100 / QTD_RETIRADO 
																APROVEITAMENTO = formatNumber(cdbl(APROVEITAMENTO))

																
															for i = 0 to 7
																if produtos(i) = "" then
																	produtos(i) = 0
																end if
															next
															
												%>
								
													<tr>
														<td class="text-left"><%=NOME_VENDEDOR%> </td>
														<td class="text-center"><%=QTD_RETIRADO%> </td>
														<td class="text-center"><%=QTD_VENDIDO%> </td>
														<td class="text-right"><%=FormatCurrency(VALOR_VENDIDO)%> </td>												
														<td class="text-center"><%=APROVEITAMENTO%> </td>
														<%
															for i = 0 to 7											
														%>
															<td class="text-center"> <%=produtos(i)%> </td>
														<%
															next
														%>
														<td class="text-right"><%=TOTAL_ALBUNS_GERAL%> </td>												
														<td class="text-right"><%=FormatCurrency(TOTAL_RECEBIDO_GERAL)%> </td>												

														
													</tr>

												<%																			
															rs_principal.MoveNext()
														wend
															
														rs_principal.Close
														Set rs_principal=Nothing
													FecharConexao()
													%>
												</tbody>											
											</table>
											<br><br>
								</td>
							</tr>

						</table>
						<!-- #include file="includes/rodape.asp" -->

				  </td>
				</tr>
			</table>
		</td>
	</tr>
</body>
</html>





