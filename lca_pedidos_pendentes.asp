<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!--
'*** Instru��es
'1) Relaciona todos os vendedores
'2) Mostra somentes os vendedores que possuem pedidos pendentes
-->
<%
Dim msg, cod_vendedor
msg = Request.QueryString("msg")

If Request.Form("btnV") <>	Empty Then

Dim PEDIDO_CHEGADA_VALIDA, PEDIDOS_RECIBO, COD_VENDEDOR_RECIBO
	AbrirConexao()
	Set rs=objConexao.Execute("SELECT COD_PEDIDO FROM TB_PEDIDOS WHERE DATA_DESCEU_DIRETORIA IS NULL AND DATA_CHEGADA IS NULL AND FINALIZA_CORRECAO <> 3 AND FINALIZA_VENDEDOR = 1")
		
		PEDIDOS_RECIBO = ""
		COD_VENDEDOR_RECIBO = Request.QueryString("cod_vendedor")
		
		While not rs.Eof
			
			PEDIDO_CHEGADA_VALIDA = formatNumber(0&Request.Form("PEDIDO_CHEGADA_"&rs("COD_PEDIDO")),0)
			
			'Response.Write(Request.Form("PEDIDO_CHEGADA_"&rs("COD_PEDIDO"))&"<br>")
			
			If int(PEDIDO_CHEGADA_VALIDA) = 1 Then
				
				sql = "UPDATE TB_PEDIDOS SET DATA_CHEGADA = NOW( ) WHERE COD_PEDIDO ="&rs("COD_PEDIDO")&" LIMIT 1"
				objConexao.Execute(sql)
				
				PEDIDOS_RECIBO = PEDIDOS_RECIBO&", "&rs("COD_PEDIDO")
				
			End If
			
		rs.MoveNext()
		Wend
		
	rs.Close
	Set rs=Nothing
	FecharConexao()
		
	msg ="Entrada de pedidos efetuada com sucesso!"
	PEDIDOS_RECIBO = Mid(PEDIDOS_RECIBO,3,Len(PEDIDOS_RECIBO))
	'Response.Redirect("lca_pedidos_pendentes.asp?msg=")
End If

If Request.QueryString("comando") = "limparChegada" Then
	
	AbrirConexao()
		objConexao.Execute("UPDATE TB_PEDIDOS SET DATA_CHEGADA = NULL WHERE COD_PEDIDO ='"&Request.QueryString("COD_PEDIDO")&"' LIMIT 1")
	FecharConexao()	
	
	Response.Redirect("lca_pedidos_pendentes.asp?msg=Chegada limpa com sucesso!")
End If

AbrirConexao()
Dim valida_permissao_para_alteracao_pedidos, valida_entrada_venda_vendedores, valida_pagina_baixa_diretoria
valida_permissao_para_alteracao_pedidos = validaUsuario(1,37,,true)
valida_entrada_venda_vendedores = validaUsuario(2,,70,true)
valida_pagina_baixa_diretoria = validaUsuario(1,66,,true)

FecharConexao()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
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
	document.location.href = url;

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
 


<!-- AJUDA F1 E CTRL PARA BUSCA RÁPIDA -->

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->

function cadastraVendedor(valor){
	alert(valor);
}

function popUp(URL){
	win=window.open(URL,"","width=350,height=220,scrollbars=yes")
	var text = "Se a janela n&atilde;o abrir desbloquei os Pop-ups!";
	if(win == null) { alert(text); return; }
	win.moveTo(350,150);
}
</script>

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
function confirmaChegada(cod_vendedor){
	if(YNconfirm("Tem certeza que deseja dar entrada nas vendas selecionadas e imprimir o recibo de entrega de venda para este vendedor?")){
		document.poliimpressao.target="_parent";
		document.poliimpressao.action="lca_pedidos_pendentes.asp?cod_vendedor="+cod_vendedor;
		document.poliimpressao.submit();
	}	

}

function limparChegada(COD_PEDIDO){
	if(YNconfirm("Tem certeza que deseja limpar a chegada do pedido "+COD_PEDIDO+"?")){
		abre("?cod_pedido="+COD_PEDIDO+"&comando=limparChegada")
	}	
}
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
        <style type="text/css">
<!--
.style1 {
	color: #009900;
	font-weight: bold;
}
.style2 {
	color: #CC0000;
	font-weight: bold;
}
.listaVendedores .linha:hover{
	background: #ccc;

}
-->
        </style>
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

	  <h2>LCA - Pedidos Pendentes </h2>
      <%If valida_pagina_baixa_diretoria Then%>
	  		<div class="subMenu">
	  			Sub-Menu: <a href="lca_pedidos_pendentes_baixar_para_diretoria.asp">Baixar Fichas</a>
            </div>
	  <%End If%>
       <div class="msg" align="center"><%=msg%></div>
	  	  <%
		  If Len(PEDIDOS_RECIBO) > 3 Then
		  %>
    		<div align="center"><a href="lca_pedidos_pendentes_recibo_venda.asp?pedidos=<%=PEDIDOS_RECIBO%>&cod_vendedor=<%=COD_VENDEDOR_RECIBO%>" target="_blank" class="linkMaster">IMPRIMIR RECIBO DE ENTRADA DE VENDA</a></div>
          <%
		  End If
		  %>      
		  	<h3 align="center">Pedidos Pendentes  </h3>
			<%
			AbrirConexao()

		Dim vendedorId,COUNT_PEDIDOS

		vendedorId = Request.QueryString("vid")
		codVendedor = Request.QueryString("cod_vendedor")
		COUNT_PEDIDOS = 0 

		If vendedorId = "" and codVendedor = "" then
			Dim rs5, sql1, pp

			sql1 = "SELECT TB_VENDEDORES.COD_VENDEDOR,TB_VENDEDORES.NOME_VENDEDOR, COUNT(DISTINCT VW_PEDIDOS_COMPLETO.COD_PEDIDO) AS PENDENETES,VW_PEDIDOS_COMPLETO.COD_PEDIDO FROM VW_PEDIDOS_COMPLETO JOIN TB_VENDEDORES ON TB_VENDEDORES.COD_VENDEDOR = VW_PEDIDOS_COMPLETO.COD_VENDEDOR WHERE DATA_DESCEU_DIRETORIA IS NULL AND FINALIZA_CORRECAO <> 3 AND  FINALIZA_VENDEDOR = 1 GROUP BY VW_PEDIDOS_COMPLETO.COD_VENDEDOR ORDER BY TB_VENDEDORES.NOME_VENDEDOR"


			Set rs5 = objConexao.Execute(sql1)
			
		%>
			
			    <table width="100%" border="1" class="listaVendedores" align="center" cellpadding="5" style="border-collapse:collapse">
                  <tr bgcolor="whitesmoke">
                  	<td><strong>Vendedor</strong></td>
                  	<td><strong>Pedidos Pendentes</strong></td>
                  	<td></td>
				  </tr>

				  
					<%
					While not rs5.Eof 
					%>

					<tr class="linha">
					<td><%=Ucase(rs5("NOME_VENDEDOR"))%></td>
                  	<td><%=rs5("PENDENETES")%></td>
                  	<td>
                  		<a href="lca_pedidos_pendentes.asp?vid=<%=rs5("cod_vendedor")%>">
                  			Visualizar
                  		</a>
                  	</td>
				
					</tr>
					<%
					pp = CInt(rs5("PENDENETES"))
					COUNT_PEDIDOS = COUNT_PEDIDOS + pp 
					rs5.MoveNext()
					Wend

					rs5.Close
					Set rs5 = Nothing
					%>
			</table>
		<%
			

		Else
		  
		  Dim rs, rs2, rs3, rs4, sql
		  '1) Consulta todos os vendedores, por�m s� mostra na tela os que possuem pedidos pendentes
		  If vendedorId = "all" then
		 	sql = "SELECT COD_VENDEDOR, NOME_VENDEDOR FROM TB_VENDEDORES ORDER BY NOME_VENDEDOR"
		  Elseif codVendedor <> "" then
			sql = "SELECT COD_VENDEDOR, NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR = "&codVendedor&" ORDER BY NOME_VENDEDOR"
		  Else
			sql = "SELECT COD_VENDEDOR, NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR = "&vendedorId&" ORDER BY NOME_VENDEDOR"
		  End If

		  Set rs = objConexao.Execute(sql)
			Dim COUNT_PENDENTES, COUNT_FINALIZADOS 
			COUNT_PENDENTES		= 0
			COUNT_FINALIZADOS	= 0
		  	
			While not rs.Eof 
		  	
			  
		 		 	'2) Verifica se h&aacute; registros pendentes do vendedor acima
			'		Filtros: Vendedor, N&atilde;o desceu para diretoria (Pendentes) e Pedidos Ativos
			sql = "SELECT * FROM VW_PEDIDOS_COMPLETO WHERE COD_VENDEDOR="&rs("COD_VENDEDOR")&" AND DATA_DESCEU_DIRETORIA IS NULL AND FINALIZA_CORRECAO <> 3 AND FINALIZA_VENDEDOR = 1 GROUP BY COD_PEDIDO ORDER BY NOME_FORMANDO"	  	
			Set rs2 = objConexao.Execute(sql)
			
			If not rs2.Eof Then
					
			DIM COUNT_PENDENTES_POR_VENDEDOR, COUNT_FINALIZADOS_POR_VENDEDOR
			COUNT_PENDENTES_POR_VENDEDOR 	= 0
			COUNT_FINALIZADOS_POR_VENDEDOR 	= 0
			%>
			<form name="poliimpressao" method="post" action="lca_pedidos_pendentes_poliimpressao.asp" target="_blank" style="margin:0px">
				<table style="margin-top:10px" cellpadding="5" cellspacing="0" align="center" width="100%">
			<tr>
			<td>
		      <div align="left"><span class="textoMedio"><strong><%=Ucase(rs("NOME_VENDEDOR"))%></strong></span></div></td>
			</tr>
			</table>
			    <table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
                  <tr bgcolor="whitesmoke">
                    <td width="9%"><table width="100%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><div align="center"><strong>C&oacute;d Pedido</strong> </div>
                        <div align="center"></div></td>
                        <td><div align="right">
										
						<a href="javascript:selecaoVendedor_<%=rs("COD_VENDEDOR")%>()"><img src="images/icons/icon_print.gif" width="20" height="19" border="0" /></a>
						<input type="hidden" name="VERIFICA_SELECIONADO_<%=rs("COD_VENDEDOR")%>" value="0" />
						
						</div></td>
                      </tr>
                    </table></td>
                    <td width="22%"><div align="center"><strong>C&oacute;d. Formando(s) / Formando(s)</strong> </div></td>
                    <td width="5%"><div align="center"><strong>Form. Pag.</strong></div></td>
                    <td width="8%"><div align="center"><strong>Tipo Venda</strong></div></td>
                    <td width="8%"><div align="center"><strong>Ajuda de Custo</strong></div></td>
                    <td width="6%"><div align="center"><strong>Grupo Especial</strong></div></td>
                    <td width="8%"><div align="center"><strong>Total Recebido </strong></div></td>
                    <td width="4%"><div align="center"><strong>Lan&ccedil;. do Pedido</strong></div></td>
                    <td width="10%"><div align="center">
                    <table width="100" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td width="76%"><div align="center"><strong>Chegou em:</strong></div>
                          <div align="center"></div></td>
                        <td width="24%"><div align="right">
										
						<a href="javascript:selecaoVendedorChegada_<%=rs("COD_VENDEDOR")%>()"><img src="images/icons/icon_setas.gif" width="22" height="15" border="0" /></a>
						<input type="hidden" name="VERIFICA_SELECIONADO_CHEGADA_<%=rs("COD_VENDEDOR")%>" value="0" />
						
						</div></td>
                      </tr>
                    </table>
                    </div></td>
                    <td width="4%"><div align="center"><strong>Vend. recebe em: </strong></div></td>
                    <td width="5%"><div align="center"><strong>Conferido por</strong></div></td>
                    <td width="6%"><div align="center"><strong>Status</strong></div></td>
                    <td width="5%"><div align="center"><strong>Corrigir Pedido</strong></div></td>
                  </tr>
                  <%
				Dim JAVA_CHECKED_TRUE, JAVA_CHECKED_FALSE, JAVA_CHEGADA_CHECKED_TRUE, JAVA_CHEGADA_CHECKED_FALSE
				JAVA_CHECKED_TRUE	= ""
				JAVA_CHECKED_FALSE	= ""
				JAVA_CHEGADA_CHECKED_TRUE 	= ""
				JAVA_CHEGADA_CHECKED_FALSE 	= ""
				  
				  While not rs2.Eof
				  
				  DATA_CHEGADA = ""&rs2("DATA_CHEGADA")
				  %>
                  <tr>
                    <td><div align="center">
                        <%If int(rs2("FINALIZA_CORRECAO")) = 1 Then%>
                      
						     <table width="100%" border="0" cellspacing="0" cellpadding="0">
							  <tr>
							    <td width="72%"><div align="center"><a href="lca_busca_detalhamento_pedido.asp?cod_pedido=<%=rs2("COD_PEDIDO")%>"><%=rs2("COD_PEDIDO")%></a><%If int(rs2("PEDIDO_EXTRA")) = 1 Then%>&nbsp;<em class="textoMenor">(Extra)</em><%End If%></div></td>
							    <td width="28%"><div align="center">
							      <input class="inputRadio" type="checkbox" name="PEDIDO_<%=rs2("COD_PEDIDO")%>" value="1" />		
							    </div></td>
							  </tr>
							</table>
						
						
						
                        <%Else%>
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
							  <tr>
							    <td width="72%"><div align="center"><%=rs2("COD_PEDIDO")%><%If int(rs2("PEDIDO_EXTRA")) = 1 Then%>&nbsp;<em class="textoMenor">(Extra)</em><%End If%></div></td>
							    <td width="28%">
									<div align="center">
							&nbsp;		</div>	</td>
							  </tr>
							</table>
	                        <%End If%>
	                    </div></td>
	                    <td><div align="left">
                        <%
						sql = 	"SELECT TB_PEDIDOS_FORMANDOS.COD_FORMANDO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO "&_
								"FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON "&_
								"TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO "&_
								"WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&rs2("COD_PEDIDO")&" ORDER BY NOME_FORMANDO"
						Set rs3 = objConexao.Execute(sql)
							
					
						While not rs3.Eof
											
						%>
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="23%"><div align="center"> <a href="dip_busca_detalhamento_formando.asp?cod_formando=<%=rs3("COD_FORMANDO")%>"> <%=digito(rs3("COD_FORMANDO"))%> </a> </div></td>
                            <td width="57%"><div align="left"><%=rs3("NOME_FORMANDO")%></div></td>
                          </tr>
                        </table>
                      <%
				rs3.MoveNext()
				Wend
				
			rs3.Close
			Set rs3 = Nothing
			%>
                    </div></td>
                    <td>
                    <%
					If int(rs2("FINALIZA_CORRECAO")) = 1 Then
					
						If int(rs2("FORMA_PAGAMENTO")) = 1 Then
							Response.Write("&Agrave; vista")
						Elseif int(rs2("FORMA_PAGAMENTO")) = 2 Then
							Response.Write("&Agrave; prazo")					
						End If
					Else
					%>
                    <em class="textoMenor">N&atilde;o dispon&iacute;vel</em>
                    <%
					End If
					%>                    </td>
                    <td> <%
					If int(rs2("FINALIZA_CORRECAO")) = 1 Then
					
						If int(rs2("TIPO_VENDA")) = 1 Then
							Response.Write("Normal")
						Elseif int(rs2("TIPO_VENDA")) = 2 Then
							Response.Write("Repasse")
						Elseif int(rs2("TIPO_VENDA")) = 3 Then
							Response.Write("Repasse Especial")
						End If
					Else
					%>
                    <em class="textoMenor">N&atilde;o dispon&iacute;vel</em>
                    <%
					End If
					%></td>
                    <td align="right"><%
					If int(rs2("FINALIZA_CORRECAO")) = 1 Then
						VALOR_AJUDA_CUSTO = formatNumber(0&rs2("VALOR_AJUDA_CUSTO"))
						'VALOR_AJUDA_CUSTO = formatCurrency(VALOR_AJUDA_CUSTO)
						Response.Write(VALOR_AJUDA_CUSTO)
					Else
					%>
                    <em class="textoMenor">N&atilde;o dispon&iacute;vel</em>
                    <%
					End If
					%></td>
                    <td>  <%
					If int(rs2("FINALIZA_CORRECAO")) = 1 Then
					
						If int(rs2("COD_GRUPO_VENDEDORES")) = 1 Then
							Response.Write("Renato")
						Elseif int(rs2("COD_GRUPO_VENDEDORES")) = 2 Then
							Response.Write("M�rio")	
						Elseif int(rs2("COD_GRUPO_VENDEDORES")) = 3 Then
							Response.Write("Paulo M.")	
						Elseif int(rs2("COD_GRUPO_VENDEDORES")) = 4 Then
              Response.Write("Vicente")	
						Elseif int(rs2("COD_GRUPO_VENDEDORES")) = 0 Then
							Response.Write("N&atilde;o h&aacute;")	
						End If
					
					Else
					%>
                    <em class="textoMenor">N&atilde;o dispon&iacute;vel</em>
                    <%
					End If
					%></td>
                    <td><div align="right">
                      <%
         objConexao.Execute("SET SQL_BIG_SELECTS=1")
			Set rs4 = objConexao.Execute("SELECT * FROM VW_PEDIDOS_TOTAL_RECEBIDO WHERE COD_PEDIDO="&rs2("COD_PEDIDO"))
			Dim PAGAMENTO_EM_DINHEIRO, XTOTAL_CHEQUE, XTOTAL_CARTAO, XTOTAL_BOLETO, TOTAL_PAGO
			
			TOTAL_PAGO = 0
				
				If not rs4.Eof Then
				
					PAGAMENTO_EM_DINHEIRO	= formatNumber(0&rs4("PAGAMENTO_EM_DINHEIRO"))/1
					XTOTAL_CHEQUE			= formatNumber(0&rs4("XTOTAL_CHEQUE"))/1
					XTOTAL_CARTAO			= formatNumber(0&rs4("XTOTAL_CARTAO"))/1
					XTOTAL_BOLETO			= cdbl(0&rs4("XTOTAL_BOLETO"))/1
					TOTAL_PAGO				= (PAGAMENTO_EM_DINHEIRO + XTOTAL_CHEQUE + XTOTAL_CARTAO + XTOTAL_BOLETO)/1
					
					Response.Write(formatNumber(TOTAL_PAGO))

				End If
			
			rs4.Close()
			Set rs4 = Nothing
			%>
                    </div></td>
                    <td><div align="center"><%=ArrumaData(rs2("DATA_CADASTRO"))%></div></td>
                    <td><div align="center">
                    <%If Len(DATA_CHEGADA) > 3 Then%>
					<%=ArrumaData(DATA_CHEGADA)%>
                    	<%If int(rs2("FINALIZA_CORRECAO")) = 2 and valida_entrada_venda_vendedores Then%>
                        <a href="javascript:limparChegada(<%=rs2("COD_PEDIDO")%>)"><img src="images/icons/icon_red.gif" border="0" /></a>
                        <%End If%>
                    <%Else%>
                    <input class="inputRadio" type="checkbox" name="PEDIDO_CHEGADA_<%=rs2("COD_PEDIDO")%>" value="1" />
                    <%End If%>
                    </div></td>
                    <td><div align="center">
                        <%If rs2("VENDEDOR_RECEBE_EM") <> "" Then%>
                        <%=arrumaData(rs2("VENDEDOR_RECEBE_EM"))%>
                        <%Else%>
                      -
                      <%End iF%>
                    </div></td>
                    <td><%
					If int(rs2("FINALIZA_CORRECAO")) = 1 Then
					
						Set rs4=objConexao.Execute("SELECT NOME FROM TB_ADM_USUARIOS WHERE COD_USUARIO = "&sqlTexto(rs2("CONFERIDO_POR_COD_USUARIO")))	
							
							If not rs4.Eof Then
							
									Response.Write(Left(rs4("nome"),3))
							Else
							
									Response.Write("-")
							
							End If
							
						
						rs4.Close
						Set rs4=Nothing
					
					Else
					%>
                    <em class="textoMenor">N&atilde;o dispon&iacute;vel</em>
                    <%
					End If
					%></td>
                    <td><div align="center">
                        <%If rs2("FINALIZA_CORRECAO") = 1 Then
						COUNT_FINALIZADOS_POR_VENDEDOR = COUNT_FINALIZADOS_POR_VENDEDOR + 1
						
						%>
                        <img src="images/icons/icon_balao_green.gif" width="17" height="16" title="<%=ArrumaData(rs2("DATA_CORRIGIDO_DEP_FOTO"))%>" alt="<%=ArrumaData(rs2("DATA_CORRIGIDO_DEP_FOTO"))%>" />
                        <%Else
						
						COUNT_PENDENTES_POR_VENDEDOR = COUNT_PENDENTES_POR_VENDEDOR + 1	
									
								'Verifica se est&aacute; aguardando aprova&ccedil;&atilde;o			
								Set rs3=objConexao.Execute("SELECT * FROM TB_APROVACOES WHERE parametro = "&sqlTexto(rs2("COD_PEDIDO")))
								
									If not rs3.Eof Then
									
											If int(rs3("status")) = 1 Then%>
                               	  <img src="images/icons/icon_aprovado.jpg" alt="Aprovado, finalize a venda" title="Aprovado, finalize a venda" />
                                            
                                  <%Elseif int(rs3("status")) = 2 Then%>
                               	  <img src="images/icons/icon_reprovado.jpg" alt="Reprovado" title="Reprovado" />
                                            <%Else%>
                               	  <img src="images/icons/icon_aguardando_aprovacao.jpg" alt="Aguardando Aprova��o" title="Aguardando Aprova��o" />
                                  <%
											End if
									Else%>
                               	  <img src="images/icons/icon_balao_red.gif" width="17" height="16" />
                                  <%
									End If
								
								rs3.Close
								Set rs3=Nothing
								%>
                        <%End If%>
                    </div></td>
                    <td><div align="center">
                    <%If Len(DATA_CHEGADA) > 3 Then%>
                    	<%If valida_permissao_para_alteracao_pedidos Then%>
                    		<a href="lca_pedidos_pendentes_alteracoes_part_1.asp?cod_pedido=<%=rs2("COD_PEDIDO")%>"><img src="images/icons/icon_alterar.png" width="16" height="16" border="0" /></a>
                        <%Else%>
                    	 -
                        <%End If%>
                    <%Else%>
                    <em class="textoMenor">N&atilde;o permitido</em>
                    <a href="javascript:alert('Pedido ainda chegou na empresa, d&aacute; entrada na venda antes de corrigir')">
<img src="images/icons/icon_help_mini.jpg" border="0"  />                    </a>
                    <%End If%>
                    </div></td>
                  </tr>
                <%
				
				If int(rs2("FINALIZA_CORRECAO")) = 1 Then
				JAVA_CHECKED_TRUE	= JAVA_CHECKED_TRUE&" d.PEDIDO_"&rs2("COD_PEDIDO")&".checked=""checked""; "
				JAVA_CHECKED_FALSE	= JAVA_CHECKED_FALSE&" d.PEDIDO_"&rs2("COD_PEDIDO")&".checked=""""; "
				End If
				
				If Len(DATA_CHEGADA) < 3 Then
				JAVA_CHEGADA_CHECKED_TRUE	= JAVA_CHEGADA_CHECKED_TRUE&" d.PEDIDO_CHEGADA_"&rs2("COD_PEDIDO")&".checked=""checked""; "
				JAVA_CHEGADA_CHECKED_FALSE	= JAVA_CHEGADA_CHECKED_FALSE&" d.PEDIDO_CHEGADA_"&rs2("COD_PEDIDO")&".checked=""""; "
				End If
				'
				rs2.MoveNext()
				Wend
				%>
				<script language="javascript">
						function selecaoVendedor_<%=rs("COD_VENDEDOR")%>(){
							d = document.poliimpressao
							if(d.VERIFICA_SELECIONADO_<%=rs("COD_VENDEDOR")%>.value == 0){
								d.VERIFICA_SELECIONADO_<%=rs("COD_VENDEDOR")%>.value = 1;
								 selecionartodasVendedor_<%=rs("COD_VENDEDOR")%>();
							}else{
								d.VERIFICA_SELECIONADO_<%=rs("COD_VENDEDOR")%>.value = 0;
								tirarselecaodeTodasVendedor_<%=rs("COD_VENDEDOR")%>()
							}
						}
				
						function selecionartodasVendedor_<%=rs("COD_VENDEDOR")%>(){
							d = document.poliimpressao		
							<%=JAVA_CHECKED_TRUE%>
						}
						
						function tirarselecaodeTodasVendedor_<%=rs("COD_VENDEDOR")%>(){
							d = document.poliimpressao
							<%=JAVA_CHECKED_FALSE%>
						}
						
						//DATA DE CHEGADA
						
						function selecaoVendedorChegada_<%=rs("COD_VENDEDOR")%>(){
							d = document.poliimpressao
							if(d.VERIFICA_SELECIONADO_CHEGADA_<%=rs("COD_VENDEDOR")%>.value == 0){
								d.VERIFICA_SELECIONADO_CHEGADA_<%=rs("COD_VENDEDOR")%>.value = 1;
								 selecionartodasVendedorChegada_<%=rs("COD_VENDEDOR")%>();
							}else{
								d.VERIFICA_SELECIONADO_CHEGADA_<%=rs("COD_VENDEDOR")%>.value = 0;
								tirarselecaodeTodasVendedorChegada_<%=rs("COD_VENDEDOR")%>()
							}
						}
						
						
						function selecionartodasVendedorChegada_<%=rs("COD_VENDEDOR")%>(){
							d = document.poliimpressao		
							<%=JAVA_CHEGADA_CHECKED_TRUE%>
						}
						
						function tirarselecaodeTodasVendedorChegada_<%=rs("COD_VENDEDOR")%>(){
							d = document.poliimpressao
							<%=JAVA_CHEGADA_CHECKED_FALSE%>
						}
						
					</script>
                </table>
			    <table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
			<tr bgcolor="whitesmoke">
			<td>
			<%
			COUNT_PENDENTES		= COUNT_PENDENTES + COUNT_PENDENTES_POR_VENDEDOR
			COUNT_FINALIZADOS	= COUNT_FINALIZADOS + COUNT_FINALIZADOS_POR_VENDEDOR
			COUNT_PEDIDOS = COUNT_FINALIZADOS + COUNT_PENDENTES
			%>
            
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="40%"><table width="400" border="0" cellspacing="0" cellpadding="3">
     <%If valida_entrada_venda_vendedores Then%>
      <tr>
        <th width="22" scope="col"><img src="images/icons/icon_setas.gif" width="22" height="15" border="0" /></th>
        <th width="366" scope="col"><div align="left"><a href="javascript:confirmaChegada(<%=rs("COD_VENDEDOR")%>)"><strong>Dar entrada nas vendas e imprimir recibo</strong></a></div></th>
      </tr>
     <%End If%> 
     
    
      <tr>
        <th scope="col"><img src="images/icons/icon_print.gif" alt="" width="20" height="19" border="0" /></th>
        <th scope="col"><div align="left"><a href="javascript:document.poliimpressao.submit();"><strong>Imprimir Pedidos Selecionados</strong></a></div></th>
      </tr>
    </table>
      </td>
    <td width="60%"><div align="right">Total Pend&ecirc;ntes: <span class="style2"><%=COUNT_PENDENTES_POR_VENDEDOR%></span>
		        <br />
		        Total Finalizados: <span class="style1"><%=COUNT_FINALIZADOS_POR_VENDEDOR%></span> </div></td>
  </tr>
</table>

		      </td>
			</tr>
			</table>
			
			<%
			End If
			
			rs2.Close
			Set rs2 = Nothing
			%>
			<%
		 	rs.MoveNext()
		  	Wend
	      	
		  rs.Close()
		  Set rs = Nothing
		 End If 
		  FecharConexao()
		  %>
     <p class="textoMaior"> 
     	<%If vendedorId <> "" or codVendedor <> "" then%>
     	<strong>
     		Total Pendentes:&nbsp; <span class="style2"><%=COUNT_PENDENTES%></span><br />
 			Total Finalizados:</strong>&nbsp;<strong class="style1"><%=COUNT_FINALIZADOS%></strong> <br />
 			<strong> Total Pedidos:&nbsp; <%=COUNT_PEDIDOS%> <strong> 
 		<%Else%>
 			<strong> Total Pedidos:&nbsp; <%=COUNT_PEDIDOS%> <a href="lca_pedidos_pendentes.asp?vid=all">Visualizar Todos</a> 
 		<%End If%>
 	</p>
				<div align="left"><br />
				  <table width="385" border="0" cellspacing="0" cellpadding="2">
                       <tr>
                      <td colspan="2" style="background-color:#EAEAEA; padding:5px;"><div align="left"><strong>Legenda de &Iacute;cones:</strong></div></td>
                      </tr>
                    
                    <tr>
                      <td width="25"><div align="center"><img src="images/icons/icon_balao_green.gif" width="17" height="16" /></div></td>
                      <td width="352">Venda Finalizada</td>
                    </tr>
                    <tr>
                      <td><div align="center"><img src="images/icons/icon_balao_red.gif" width="17" height="16" /></div></td>
                      <td>Venda N&atilde;o corrigida</td>
                    </tr>
                    <tr>
                      <td><div align="center"><img src="images/icons/icon_aprovado.jpg" alt="Aprovado, finalize a venda" title="Aprovado, finalize a venda" /></div></td>
                      <td>Venda com valor inferior a tabela, mas aprovada</td>
                    </tr>
                    <tr>
                      <td><div align="center"><img src="images/icons/icon_reprovado.jpg" alt="Reprovado" title="Reprovado" /></div></td>
                      <td>Venda com valor inferior a tabela, mas reprovada</td>
                    </tr>
                    <tr>
                      <td><div align="center"><img src="images/icons/icon_aguardando_aprovacao.jpg" alt="Aguardando Aprova&ccedil;&atilde;o" title="Aguardando Aprova&ccedil;&atilde;o" /></div></td>
                      <td>Venda com valor inferior a tabela aguardando aprova&ccedil;&atilde;o</td>
                    </tr>
                  </table>
				</div>
               <p align="center">     
		  <input class="btnTopo" onclick="abre('#')" type="button" name="btnTopo2" value=" " />
		</p>		
		<input type="hidden" name="btnV" value="1" />
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
