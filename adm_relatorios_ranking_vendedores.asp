<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'///////////////////////////////
'Validação especial para verificar se possui sub-validação específica
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '10' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se não encontrar então redireciona
			Session.Abandon()
			Response.Redirect("index.asp?falha=1")
		End If
		
	rs.Close
	Set rs=Nothing
	
FecharConexao()



Dim VENDEDOR_RECEBE_EM
VENDEDOR_RECEBE_EM = Request.QueryString("anoUtilizar")&"-"&Request.QueryString("mesUtilizar")&"-01"
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
	win=window.open(URL,"","width=350,height=220,scrollbars=yes")
	var text = "Se a janela não abrir desbloquei os Pop-ups!";
	if(win == null) { alert(text); return; }
	win.moveTo(350,150);
	}
	</script>

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
  BTN_IMPRIMIR = 1
  BTN_FONTE = 1
  BTN_IMPRIMIR_LINK = "href=""javascript:popUp('adm_relatorios_ranking_vendedores_IMPRIMIR.asp?mesUtilizar="&Request.QueryString("mesUtilizar")&"&anoUtilizar="&Request.QueryString("anoUtilizar")&"')"""
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">
	
	  <h2>Relatórios - Ranking de Vendedores </h2>
	  <div align="center" class="msg">
	    <p><%=Request.QueryString("msg")%></p>
	  </div>
		<div id="comboTexto">
	  
	     <%Dim rs, sql, rs2, sql2,sql3, asdasd,asd, AlbunsTotais
		 
		 AlbunsTotais = 0

		 
		  	AbrirConexao()		  
 		
		sql2 =  "SELECT "&_
				"NOME_VENDEDOR, "&_
				"(SUM(IFNULL(PAGAMENTO_EM_DINHEIRO,0)) + "&_
				"SUM(IFNULL(XTOTAL_CHEQUE,0)) + "&_
				"SUM(IFNULL(XTOTAL_CARTAO,0))) "&_
				"AS TOTAL_RECEBIDO_GERAL "&_
				"FROM VW_EXTRATO_TOTAL_RECEBIDO_COMPLETO WHERE "&_
				"VENDEDOR_RECEBE_EM = '"&VENDEDOR_RECEBE_EM&"' "&_
				"GROUP BY COD_VENDEDOR "&_
				"ORDER BY TOTAL_RECEBIDO_GERAL DESC "
		



	'Response.Write(sql2)
					
			Set rs2 = objConexao.Execute(sql2)
			Dim num_p, VendaTotal
			VendaTotal = 0
			num_p = 0%>
<%If rs2.Eof then
Response.Redirect("adm_relatorios.asp?msg=Não há dados para o ranking de vendedores do período selecionado: "&MonthName(Month(VENDEDOR_RECEBE_EM))&"/"&Year(VENDEDOR_RECEBE_EM)&", selecione outro período!")%> 	

		
<%else%>
	 
	     
	      	  <h3 align="center"><strong> Ranking por  arrecadação</strong></h3>
	      	 		<table width="80%" 
		      border="1" align="center"
					cellpadding="5" 
					style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
                <td><div align="center"><span class="MensagemFicha"><strong>P</strong></span><strong><span class="MensagemFicha">osição</span></strong></div></td>
                <td><div align="left" class="MensagemFicha">
                  <div align="center"><strong>Vendedor</strong></div>
                </div></td>
                </tr>
		     
			<%While not rs2.Eof
			num_p = num_p + 1
			    %>
                     
				<tr>
		          <td width="15%"><div align="center"><%=num_p%>º</div></td>

                  <td><div align="left"><%=Ucase(rs2("NOME_VENDEDOR"))%></div></td>
                  </tr>
					 
					 
          <%
		  	rs2.MoveNext()
			Wend
			
			rs2.Close
			Set rs2 = Nothing
		  %>
	      	 		</table>
		  <div align="center">
	      	  <h3><strong><br />Ranking por número de álbuns vendidos </strong> </h3>
		  </div>
	      	 		<table width="80%" 
		      border="1" align="center"
					cellpadding="5" 
					style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
                <td><div align="center"><span class="MensagemFicha"><strong>P</strong></span><strong><span class="MensagemFicha">osição</span></strong></div></td>
                <td><div align="left" class="MensagemFicha">
                  <div align="center"><strong>Vendedor</strong></div>
                </div></td>
                <td width="25%">
                  
                  <div align="center" class="MensagemFicha"><strong>Total</strong></div></td></tr>
		      <%	 
	
		sql3 =	"SELECT TB_PEDIDOS.COD_VENDEDOR, SUM( VW_QTD_FORMANDOS_POR_PEDIDO.QTD_FORMANDOS ) AS TOTAL_ALBUNS "&_
				"FROM VW_QTD_FORMANDOS_POR_PEDIDO "&_
				"INNER JOIN TB_PEDIDOS ON VW_QTD_FORMANDOS_POR_PEDIDO.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO "&_
				"WHERE VENDEDOR_RECEBE_EM = '"&VENDEDOR_RECEBE_EM&"'"&_
				"AND FINALIZA_CORRECAO =1 "&_
				"GROUP BY COD_VENDEDOR "&_
				"ORDER BY TOTAL_ALBUNS DESC  "

	
					
			Set rs3 = objConexao.Execute(sql3)
			Dim num_p2
			num_p2 = 0
			While not rs3.Eof
			num_p2 = num_p2 + 1
			    %>
                     
				<tr>
		          <td width="15%"><div align="center"><%=num_p2%>º</div></td>
                  <td width="60%"><div align="left"><%
				  Dim rsVendedorNome
				  Set rsVendedorNome = objConexao.Execute("SELECT NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR = "&rs3("COD_VENDEDOR"))
				  			
							Response.Write(UCASE(rsVendedorNome("NOME_VENDEDOR")))
							
				  rsVendedorNome.Close
				  Set rsVendedorNome = Nothing
				  %></a></div></td>
                  <td><div align="center"><%=rs3("TOTAL_ALBUNS")%>
				  <%
				  subAlbuns = formatNumber(rs3("TOTAL_ALBUNS"))
				  AlbunsTotais = AlbunsTotais + subAlbuns
				  
				  %>
				  </div>
                  <div align="right"></div></td>
              </tr>
					 
					 
          <%
		  	rs3.MoveNext()
			Wend
			
			rs3.Close
			Set rs3=Nothing

			
		  %>  
		  	<tr bgcolor="whitesmoke">
		  				    <td colspan="2"> </td>
                            <td><div align="center"><strong>
                              <%AlbunsTotais = Replace(AlbunsTotais,",","")%>
                              <%Response.Write(AlbunsTotais)%>
                            </strong></div></td>
		  	</table>
<%
AlbunsTotais = 0
end if


%>
<div align="center">
	      	  <h3><strong><br />
	      	  Ranking por número de DVDs vendidos </strong> </h3>
		  </div>
	      	 		<table width="80%" 
		      border="1" align="center"
					cellpadding="5" 
					style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
                <td><div align="center"><span class="MensagemFicha"><strong>P</strong></span><strong><span class="MensagemFicha">osição</span></strong></div></td>
                <td><div align="left" class="MensagemFicha">
                  <div align="center"><strong>Vendedor</strong></div>
                </div></td>
                <td width="25%">
                  
                  <div align="center" class="MensagemFicha"><strong>Total</strong></div></td></tr>
		      <%	 
	'DVDS
		sql3 =	"SELECT TB_PEDIDOS.COD_VENDEDOR, SUM(VW_PEDIDOS_PRODUTOS_TOTAIS_DVD.XTOTAL_DVD) AS TOTAL_DVD "&_
				"FROM TB_PEDIDOS "&_
				"LEFT JOIN VW_PEDIDOS_PRODUTOS_TOTAIS_DVD ON TB_PEDIDOS.COD_PEDIDO = VW_PEDIDOS_PRODUTOS_TOTAIS_DVD.COD_PEDIDO "&_
				"WHERE VENDEDOR_RECEBE_EM = '"&VENDEDOR_RECEBE_EM&"' "&_
				"GROUP BY COD_VENDEDOR "&_
				"ORDER BY TOTAL_DVD DESC "
				
						'Response.Write(sql3)	
			Set rs3 = objConexao.Execute(sql3)
			num_p2 = 0
			While not rs3.Eof
			num_p2 = num_p2 + 1
			    
				subAlbuns = formatNumber(0&rs3("TOTAL_DVD"))
				If subAlbuns > 0 Then
				%>
                     
				<tr>
		          <td width="15%"><div align="center"><%=num_p2%>º</div></td>
                  <td width="60%"><div align="left"><%
				  Set rsVendedorNome = objConexao.Execute("SELECT NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR = "&rs3("COD_VENDEDOR"))
				  			
							Response.Write(UCASE(rsVendedorNome("NOME_VENDEDOR")))
							
				  rsVendedorNome.Close
				  Set rsVendedorNome = Nothing
				  %></a></div></td>
                  <td><div align="center"><%=rs3("TOTAL_DVD")%>
				  <%
				  
				  AlbunsTotais = (AlbunsTotais + subAlbuns)/1
				  
				  %>
				  </div>
                  <div align="right"></div></td>
              </tr>
					 
					 
          <%
		  	End If
		  	
		  	rs3.MoveNext()
			Wend
			
			rs3.Close
			Set rs3=Nothing
		  
		 
		  %>  
		  	<tr bgcolor="whitesmoke">
		  				    <td colspan="2"> </td>
                            <td><div align="center"><strong>
                              <%AlbunsTotais = Replace(AlbunsTotais,",","")%>
                              <%Response.Write(AlbunsTotais)
							   AlbunsTotais = 0
							  %>
                            </strong></div></td>
		  	</table>


 <div align="center">
	      	  <h3><strong><br />
	      	  Ranking por número de Foto-Telas vendidas</strong> (Foto-Telas + Trios)</h3>
		  </div>
	      	 		<table width="80%" 
		      border="1" align="center"
					cellpadding="5" 
					style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
                <td><div align="center"><span class="MensagemFicha"><strong>P</strong></span><strong><span class="MensagemFicha">osição</span></strong></div></td>
                <td><div align="left" class="MensagemFicha">
                  <div align="center"><strong>Vendedor</strong></div>
                </div></td>
                <td width="25%">
                  
                  <div align="center" class="MensagemFicha"><strong>Total</strong></div></td></tr>
		      <%	 
	'Foto-tela + trio
		sql3 =	"SELECT TB_PEDIDOS.COD_VENDEDOR, SUM(VW_TOTAL_FOTO_TELAS_TRIOS.QTD_TOTAL) AS XTOTAL_FOTO_TELAS_E_TRIOS "&_
				"FROM TB_PEDIDOS "&_
				"LEFT JOIN VW_TOTAL_FOTO_TELAS_TRIOS ON TB_PEDIDOS.COD_PEDIDO = VW_TOTAL_FOTO_TELAS_TRIOS.COD_PEDIDO "&_
				"WHERE VENDEDOR_RECEBE_EM = '"&VENDEDOR_RECEBE_EM&"' "&_
				"GROUP BY TB_PEDIDOS.COD_VENDEDOR "&_
				"ORDER BY XTOTAL_FOTO_TELAS_E_TRIOS DESC "
							'Response.Write(sql3)
			Set rs3 = objConexao.Execute(sql3)
			num_p2 = 0
			While not rs3.Eof
			num_p2 = num_p2 + 1
			    
				subAlbuns = formatNumber(0&rs3("XTOTAL_FOTO_TELAS_E_TRIOS"))
				If subAlbuns > 0 Then
				%>
                     
				<tr>
		          <td width="15%"><div align="center"><%=num_p2%>º</div></td>
                  <td width="60%"><div align="left"><%
				  Set rsVendedorNome = objConexao.Execute("SELECT NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR = "&rs3("COD_VENDEDOR"))
				  			
							Response.Write(UCASE(rsVendedorNome("NOME_VENDEDOR")))
							
				  rsVendedorNome.Close
				  Set rsVendedorNome = Nothing
				  %></a></div></td>
                  <td><div align="center"><%=rs3("XTOTAL_FOTO_TELAS_E_TRIOS")%>
				  <%
				  
				  AlbunsTotais = (AlbunsTotais + subAlbuns)/1
				  
				  %>
				  </div>
                  <div align="right"></div></td>
              </tr>
					 
					 
          <%
		  	End If
		  	
		  	rs3.MoveNext()
			Wend
			
			rs3.Close
			Set rs3=Nothing
		  %>  
		  	<tr bgcolor="whitesmoke">
		  				    <td colspan="2"> </td>
                            <td><div align="center"><strong>
                              <%AlbunsTotais = Replace(AlbunsTotais,",","")%>
                              <%Response.Write(AlbunsTotais)%>
                            </strong></div></td>
		  	</table>
<%
			FecharConexao()
%>
	  <DIV align="center">
	    <p>
	      <input class="btnTopo" onclick="abre('#')" type="button" name="btnTopo2" value=" " />
	    </p>
	    </DIV>
	  </div>
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






