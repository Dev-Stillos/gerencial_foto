<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Valida��o especial para verificar se possui sub-valida��o espec�fica
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '18' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se Não encontrar ent�o redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
		
	rs.Close
	Set rs=Nothing
	
FecharConexao()


Dim DATA_VENDA_INICIO, DATA_VENDA_FIM
Dim sqlQuery
DATA_VENDA_INICIO = Request.QueryString("DATA_VENDA_INICIO")
DATA_VENDA_FIM = Request.QueryString("DATA_VENDA_FIM")

If DATA_VENDA_INICIO = Empty or DATA_VENDA_FIM = Empty Then
	Response.Redirect("adm_relatorios.asp?msg=Informe uma data de venda de início e outra de fim para gerar o relatório.")
End If

sqlQuery = "P.DATA_VENDA BETWEEN " & sqlData(DATA_VENDA_INICIO) & " AND " & sqlData(DATA_VENDA_FIM)

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Vers�o 1.2</title>
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
   BTN_IMPRIMIR_LINK = " target='_blank' href=""adm_relatorios_ranking_vendedores_completo_impressao.asp?DATA_VENDA_INICIO="&Request.QueryString("DATA_VENDA_INICIO")&"&DATA_VENDA_FIM="&Request.QueryString("DATA_VENDA_FIM")&""""
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">
	
	  <h2>Relatórios - Ranking de <strong>Representantes Comerciais</strong> - <%=ArrumaData(data_venda_inicio)%> à <%=ArrumaData(data_venda_fim)%></h2>
      <div><em>Filtro por data das vendas</em></div>
      
	  <div align="center" class="msg">
	    <p><%=Request.QueryString("msg")%></p>
	  </div>
		<div id="comboTexto">
	  	
      
	     <%Dim rs, sql, rs2, sql2,sql3, asdasd,asd, AlbunsTotais
		 
		 AlbunsTotais = 0

		 
		  	AbrirConexao()		  
 		
		sql2 =  "SELECT " &_ 
        "COD_VENDEDOR, " &_ 
        "NOME_VENDEDOR, " &_ 
        "(SUM(IFNULL(PAGAMENTO_EM_DINHEIRO,0)) + " &_ 
        "SUM(IFNULL(XTOTAL_CHEQUE,0)) + " &_ 
        "SUM(IFNULL(XTOTAL_CARTAO,0))) AS TOTAL_RECEBIDO_GERAL " &_ 
        "FROM VW_EXTRATO_TOTAL_RECEBIDO_COMPLETO P " &_ 
        "WHERE " & sqlQuery & " " &_ 
        "GROUP BY COD_VENDEDOR " &_ 
        "ORDER BY TOTAL_RECEBIDO_GERAL DESC"

		



	'Response.Write(sql2)
					
			Set rs2 = objConexao.Execute(sql2)
			Dim num_p, VendaTotal
			VendaTotal = 0
			num_p = 0%>
<%If rs2.Eof then
Response.Redirect("adm_relatorios.asp?msg=Não há dados para o ranking de vendedores do período selecionado, selecione outro período!")%> 	

		
<%else%>
	 
	     
	      	  <h3 align="center"><strong> Ranking por  arrecadação</strong> </h3>
	      	 		<table width="90%" 
		      border="1" align="center"
					cellpadding="5" 
					style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
                <td><div align="center"><span class="MensagemFicha"><strong>P</strong></span><strong><span class="MensagemFicha">osição</span></strong></div></td>
                <td><div align="left" class="MensagemFicha">
                  <div align="center"><strong>Representante Comercial</strong></div>
                </div></td>
                <td><div align="center"><strong>Listar Pedidos</strong></div></td>
                <td width="17%">
                  
                  <div align="center" class="MensagemFicha"><strong>Total</strong></div></td></tr>
		     
			<%While not rs2.Eof
			num_p = num_p + 1
			    %>
                     
				<tr>
		          <td width="10%"><div align="center"><%=num_p%>°</div></td>

                  <td width="36%"><div align="left"><%=Ucase(rs2("NOME_VENDEDOR"))%></div></td>
                  <td width="37%"><div align="center"><div align="center">
                  <%
				  If Request.QueryString("comando") = "listarPedidos" and int(0&Request.QueryString("COD_VENDEDOR")) = int(rs2("COD_VENDEDOR")) Then
				  %>
                  <table width="100%" 
		      border="1" align="center"
					cellpadding="2" 
					style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
    <td width="15%"><div align="center"><strong>Cód. Pedido</strong></div></td>
	<td width="10%"><div align="center"><strong>Tipo Venda</strong></div></td>
   <!-- <td width="15%"><div align="center"><strong>Ajuda de Custo</strong></div></td>-->
    <td width="20%"><div align="center"><strong>Valor Total Pedido</strong></div></td>
    <td width="50%"><div align="center"><strong>Instituição de Ensino</strong></div></td>
</tr>

  <%
Dim totalPedidosVendedor
totalPedidosVendedor = 0

 Set rs3 = objConexao.Execute("SELECT " & _
    "P.COD_PEDIDO, " & _
    "PE.TIPO_VENDA, " & _
    "PE.PRE_VENDA, " & _
    "P.VALOR_AJUDA_CUSTO, " & _
    "(IFNULL(P.PAGAMENTO_EM_DINHEIRO,0) + IFNULL(P.XTOTAL_CHEQUE,0) + IFNULL(P.XTOTAL_CARTAO,0)) AS TOTAL_RECEBIDO_GERAL, " & _
    "E.DESCRICAO AS INSTITUICAO_ENSINO " & _
    "FROM VW_EXTRATO_TOTAL_RECEBIDO_COMPLETO P " & _
    "LEFT JOIN TB_PEDIDOS_PRODUTOS PP ON P.COD_PEDIDO = PP.COD_PEDIDO " & _
    "LEFT JOIN TB_FICHAS_IDENTIFICACAO F ON PP.COD_FORMANDO = F.COD_FORMANDO " & _
    "LEFT JOIN TB_ENTIDADES E ON F.COD_CONTRATO = E.COD_CONTRATO " & _
    "LEFT JOIN TB_PEDIDOS PE ON PE.COD_PEDIDO = P.COD_PEDIDO " & _
    "WHERE " & sqlQuery & _
    "AND P.COD_VENDEDOR = " & rs2("COD_VENDEDOR") & " " & _
    "GROUP BY P.COD_PEDIDO, PE.TIPO_VENDA, PE.PRE_VENDA, P.VALOR_AJUDA_CUSTO, E.DESCRICAO, P.PAGAMENTO_EM_DINHEIRO, P.XTOTAL_CHEQUE, P.XTOTAL_CARTAO " & _
    "ORDER BY P.COD_PEDIDO ASC")
  
  	While not rs3.Eof

	totalPedidosVendedor = totalPedidosVendedor + 1

	TOTAL_RECEBIMENTO_POR_VENDEDOR = formatNumber(0&rs3("TOTAL_RECEBIDO_GERAL"))
  %>
  <tr>
    <td><div align="center"><a href="lca_busca_detalhamento_pedido.asp?cod_pedido=<%=rs3("COD_PEDIDO")%>"><%=rs3("COD_PEDIDO")%></a></div></td>
	<td><div align="center">
        <%
            Dim tp
            If rs3("TIPO_VENDA") = 1 AND rs3("PRE_VENDA") = "N" Then 
                tp = "NORMAL"
            ElseIf rs3("TIPO_VENDA") = 1 AND rs3("PRE_VENDA") = "S" Then
                tp = "PRÉ-VENDA"
            ElseIf rs3("TIPO_VENDA") = 2 Then
                tp = "REPASSE"
            Else
                tp = "-"
            End If

            Response.Write(tp)
        %>
    </div></td>
    <!--<td><div align="right"><%=formatCurrency(rs3("VALOR_AJUDA_CUSTO"))%></div></td>-->
    <td><div align="center"><%=formatCurrency(TOTAL_RECEBIMENTO_POR_VENDEDOR)%></div></td>
	<td><div align="center"><strong><%=rs3("INSTITUICAO_ENSINO")%></strong></div></td>
  </tr>
  <%
  rs3.MoveNext()
  Wend
  
  rs3.Close
  Set rs3=Nothing
  %>
  <tr>
    <td bgcolor="whitesmoke" colspan="4" align="center"><strong>Total de pedidos: <%= totalPedidosVendedor %></strong></td>
</tr>
</table>
               <DIV align="center"><a href="adm_relatorios_ranking_vendedores_completo.asp?data_venda_inicio=<%=DATA_VENDA_INICIO%>&data_venda_fim=<%=DATA_VENDA_FIM%>">Limpar Listagem</a></DIV>
                  <%Else%>
                  <a href="adm_relatorios_ranking_vendedores_completo.asp?data_venda_inicio=<%=DATA_VENDA_INICIO%>&data_venda_fim=<%=DATA_VENDA_FIM%>&comando=listarPedidos&cod_vendedor=<%=rs2("COD_VENDEDOR")%>">Listar pedidos para este mês</a>
                  <%End If%>
                  </div></div></td>
                  <td><div align="right" style="padding-right:8px">
				<%
				subValores = formatNumber(rs2("TOTAL_RECEBIDO_GERAL"))
				Response.Write(subValores)
				VendaTotal = VendaTotal + subValores
				%>
				</div></td>
              </tr>
					 
					 
          <%
		  	rs2.MoveNext()
			Wend
			
			rs2.Close
			Set rs2 = Nothing
		  %> 
		  
		  				  <tr bgcolor="whitesmoke"><td colspan="4"><div align="right"><strong><%=formatCurrency(VendaTotal)%></strong></div></td>
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
                  <div align="center"><strong>Representante Comercial</strong></div>
                </div></td>
                <td width="25%">
                  
                  <div align="center" class="MensagemFicha"><strong>Total</strong></div></td></tr>
		      <%	 
	
		sql3 =	"SELECT TB_PEDIDOS.COD_VENDEDOR, SUM( VW_QTD_FORMANDOS_POR_PEDIDO.QTD_FORMANDOS ) AS TOTAL_ALBUNS " & _
        "FROM VW_QTD_FORMANDOS_POR_PEDIDO " & _
        "INNER JOIN TB_PEDIDOS ON VW_QTD_FORMANDOS_POR_PEDIDO.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO " & _
        "WHERE " & Replace(sqlQuery, "P.", "TB_PEDIDOS.") & _
        " AND TB_PEDIDOS.FINALIZA_CORRECAO = 1 " & _
        "GROUP BY COD_VENDEDOR " & _
        "ORDER BY TOTAL_ALBUNS DESC"

			'Response.Write(sql3)
					
			Set rs3 = objConexao.Execute(sql3)
			Dim num_p2
			num_p2 = 0
			While not rs3.Eof
			num_p2 = num_p2 + 1
			    %>
                     
				<tr>
		          <td width="15%"><div align="center"><%=num_p2%>°</div></td>
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
		  				    <td align="center" colspan="2">TOTAL:</td>
                            <td><div align="center"><strong>
                              <%AlbunsTotais = Replace(AlbunsTotais,",","")%>
                              <%Response.Write(AlbunsTotais)%>
                            </strong></div></td>
		  	</table>
<%
AlbunsTotais = 0
end if


%>

    
            
            
            
            
            <%
			sql = "SELECT TB_PEDIDOS_PRODUTOS.COD_PRODUTO, TB_PRODUTOS.NOME_PRODUTO, " & _
      "SUM(TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS TOTAL_QUANTIDADE_GERAL, " & _
      "SUM(TB_PEDIDOS_PRODUTOS.VALOR_UN * TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS VALOR_TOTAL " & _
      "FROM TB_PEDIDOS " & _
      "INNER JOIN TB_PEDIDOS_PRODUTOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO " & _
      "INNER JOIN TB_PRODUTOS ON TB_PEDIDOS_PRODUTOS.COD_PRODUTO = TB_PRODUTOS.COD_PRODUTO " & _
      "WHERE " & Replace(sqlQuery, "P.", "TB_PEDIDOS.") & _
      " AND TB_PEDIDOS.FINALIZA_CORRECAO = 1 " & _
      "GROUP BY TB_PEDIDOS_PRODUTOS.COD_PRODUTO " & _
      "ORDER BY TOTAL_QUANTIDADE_GERAL DESC"

			'Response.Write(sql)
			
			Set rs=objConexao.Execute(sql,QTD_PRODUTOS)
			
			
			
			
				If not rs.Eof Then
				
							While not rs.Eof
																
								%>
                                 <div align="center">
	      	 						<h3><strong>Ranking de <%=Ucase(rs("NOME_PRODUTO"))%> - <%=ArrumaData(DATA_VENDA_INICIO)%> à <%=ArrumaData(DATA_VENDA_FIM)%></strong></h3></div>
                                    <table width="80%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">   
                                    	<tr bgcolor="whitesmoke">
                    						<td align="center"><strong>Posição</strong></td>
                    						<td align="center"><strong>Representante Comercial</strong></td>
                    						<td width="25%" align="center"><strong>Total</strong></td></tr>
                  <%	 
       			
				i = 0
				total = 0
	   
	   			sql = "SELECT TB_VENDEDORES.NOME_VENDEDOR, " & _
      "SUM(TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS TOTAL_QUANTIDADE, " & _
      "SUM(TB_PEDIDOS_PRODUTOS.VALOR_UN * TB_PEDIDOS_PRODUTOS.QUANTIDADE) AS VALOR_TOTAL " & _
      "FROM TB_PEDIDOS " & _
      "INNER JOIN TB_PEDIDOS_PRODUTOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO " & _
      "INNER JOIN TB_PRODUTOS ON TB_PEDIDOS_PRODUTOS.COD_PRODUTO = TB_PRODUTOS.COD_PRODUTO " & _
      "INNER JOIN TB_VENDEDORES ON TB_PEDIDOS.COD_VENDEDOR = TB_VENDEDORES.COD_VENDEDOR " & _
      "WHERE " & Replace(sqlQuery, "P.", "TB_PEDIDOS.") & _
      " AND TB_PEDIDOS.FINALIZA_CORRECAO = 1 " & _
      " AND TB_PEDIDOS_PRODUTOS.COD_PRODUTO = " & rs("COD_PRODUTO") & _
      " GROUP BY TB_PEDIDOS.COD_VENDEDOR " & _
      " ORDER BY TOTAL_QUANTIDADE DESC"

				   
				                                
                Set rs2 = objConexao.Execute(sql)

				
                While not rs2.Eof
               		
					i = i + 1
                    subtotal = formatNumber(0&rs2("TOTAL_QUANTIDADE"),0)/1
                    total =  (total+subtotal)/1
                    %>
                         
                    <tr>
                      <td width="15%"><div align="center"><%=i%>°</div></td>
                      <td width="60%"><div align="left"><%=Ucase(rs2("NOME_VENDEDOR"))%></div></td>
                      <td><div align="center"><%=subtotal%>
                      </div>
                      <div align="right"></div></td>
                  </tr>
                         
                         
				  <%                
                rs2.MoveNext()
                Wend
                
                rs2.Close
                Set rs2=Nothing
              %>  
                <tr bgcolor="whitesmoke">
                                <td align="center" colspan="2">TOTAL:</td>
                                <td><div align="center"><strong>
                                  <%Response.Write(formatNumber(total,0))%>
                                </strong></div></td>
                </table>
                                
                                <%
								
								
							
							rs.MoveNext()
							Wend
				
				End if
			
			
			rs.Close
			Set rs=Nothing
			
			%>
            
            
            
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









