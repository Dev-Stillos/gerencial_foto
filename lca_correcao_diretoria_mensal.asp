<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim cod_fechamento_periodo, comando
cod_fechamento_periodo = Request.QueryString("cod_fechamento_periodo")

comando = Request.QueryString("comando")


Dim VALIDA_SUB48 'Verifica se o usu�rio possui acesso a liberar fechamento sem finalizar todas as vendas
VALIDA_SUB48 = formatNumber(0&inStr(sub_validacoes_paginas,"#48"),0)

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

	If comando = "alterarAutorizacao" Then
		
		AbrirConexao()
			objConexao.Execute("UPDATE TB_FECHAMENTO_PAGAMENTO_PERIODOS SET autorizacao_diretoria_para_finalizar = '"&Request.QueryString("autorizacao_diretoria_para_finalizar")&"' WHERE cod_fechamento_periodo = "&Request.QueryString("cod_fechamento_periodo"))
		FecharConexao()		
		
		Response.Redirect("lca_correcao_diretoria_mensal.asp?cod_fechamento_periodo="&Request.QueryString("cod_fechamento_periodo")&"&codmsg=1")
	
	End If
	
'/////////////////////////////////////////////////////////////////////////////////////


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>Correção Diretoria Mensal</title>
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
		
		
	

	function buscaIni2(){
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
 


	<!--AJUDA F1 E CTRL PARA BUSCA R�PIDA-->

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->

function cadastraVendedor(valor){
	alert(valor);
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
  BTN_IMPRIMIR_LINK = "#"
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">

	  <h2>LCA - Corre&ccedil;&atilde;o Diretoria / Mensal </h2>
      
      <%If Request.QueryString("msg") <> Empty Then%>
      	<div class="msgGrandeSucesso" align="center"><%=Request.QueryString("msg")%></div>
      <%End If%>
	   
       
		
						<h3 align="center">Rela&ccedil;&atilde;o de Pedidos do Pagamento para Autoriza&ccedil;&atilde;o</h3>
			           
                       <%If Cstr(""&Request.QueryString("codmsg")) = "1" Then%>
                           <div class="msgGrandeSucesso">
                           	Autorização de processar pagamento alterada com sucesso!
                           </div>
                       <%End If%>
                       
                         <table width="80%" 
		      border="1" align="center"
					cellpadding="3" 
					style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
                <td width="21%"><div align="center"><strong>Per&iacute;odo de Vendas</strong><br />
                    <em>Data da Venda</em></div></td>
                <td width="20%"><div align="center"><strong>Data M&aacute;xima Permita para o Lan&ccedil;amento</strong> <strong>no sistema</strong></div></td>
                <td width="12%"><div align="center"><strong>Data Prevista do Pagamento</strong></div></td>
                <td width="47%" ><div align="center"><strong>Data do Processamento do Pagamento</strong></div></td>
	            </tr>
<tr>
  <td class="textoMedio"><div align="center"><%=data_inicio%> à <%=data_termino%></div></td>
  <td><div align="center"><%=data_termino_lancamento_vendas%></div></td>
  <td><div align="center"><%=data_pagamento%></div></td>
  <td><div align="center">
    <%If len(data_processado_pagamento) > 2 Then%>
    <%=data_processado_pagamento%>
    <%Else%>
    <em><strong>Não foi realizado o pagamento</strong></em><br />
    </div>
    <%If VALIDA_SUB48 <> 0 Then%>
    <table bgcolor="#FFFFCC" width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td colspan="2"><div align="center">Autorização para processar pagamento deste período sem a finalização de todas as autenticações de pedidos pela diretoria</div></td>
        </tr>
      <tr>
        <td width="50%"><div align="right">
          <form name="gravar_autorizacao">
            <input type="hidden" name="cod_fechamento_periodo" value="<%=cod_fechamento_periodo%>" />
            <input type="hidden" name="comando" value="alterarAutorizacao" />
            <select name="autorizacao_diretoria_para_finalizar">
              <option value="N" <%If autorizacao_diretoria_para_finalizar = "N" Then %> selected="selected" <%End If%>>NÃO</option>
              <option value="S" <%If autorizacao_diretoria_para_finalizar = "S" Then %> selected="selected" <%End If%>>SIM</option>
              </select>
            </form>
          </div></td>
        <td width="50%"><a href="javascript:document.gravar_autorizacao.submit()"><img border="0" src="images/layout/btnGravar.jpg" width="77" height="24" /></a></td>
        </tr>
  </table>
    <%End If%>
    
    
    <%End If%>
  </td>
  </tr>
           
           
                        </table><br />
<br />

                       
                       
                       
                        <table width="80%" 
		      border="1" align="center"
					cellpadding="3" 
					style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
                <td width="63%"><div align="left" class="MensagemFicha">
                  <div align="center"><strong>Vendedor</strong></div>
                </div></td>
                <td width="12%">
                  
                  <div align="center" class="MensagemFicha"><strong>Qtd. N&atilde;o Corrigida</strong> </div></td>
                <td width="11%" class="MensagemFicha"><div align="center"><strong>Qtd.</strong> <strong>Corrigida</strong> </div></td>
	            <td width="14%" class="MensagemFicha"><div align="center"><strong>Total Vendas </strong></div></td>
                </tr>
		      <%Dim rs, sql, rs2, sql2, COD_VENDEDOR_QUERY, COD_VENDEDOR_RS

		 
		  	AbrirConexao()		  
 	
		
		sql2 = " SELECT TB_PEDIDOS.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR "&_
			" FROM TB_VENDEDORES"&_
			" INNER JOIN TB_PEDIDOS ON TB_VENDEDORES.COD_VENDEDOR = TB_PEDIDOS.COD_VENDEDOR"&_	
			" WHERE TB_PEDIDOS.cod_fechamento_periodo = "&cod_fechamento_periodo&" AND TB_PEDIDOS.FINALIZA_CORRECAO = 1 AND DATA_DESCEU_DIRETORIA IS NOT NULL GROUP BY TB_PEDIDOS.COD_VENDEDOR ORDER BY TB_VENDEDORES.NOME_VENDEDOR"

	
					
			Set rs2 = objConexao.Execute(sql2)
			
			While not rs2.Eof
		  	COD_VENDEDOR_RS		= FormatNumber(0&rs2("COD_VENDEDOR"),0)
		    COD_VENDEDOR_QUERY 	= FormatNumber(0&Request.QueryString("COD_VENDEDOR"),0)
			
			%>
                     
				<tr 
				<%
				Dim corr
				
				if corr = "#FFFFFF" then
				corr = "#E9E6E6"
				else
				corr = "#FFFFFF"
				end if
				
				%>
				<%if COD_VENDEDOR_RS = COD_VENDEDOR_QUERY then%>
				<%Response.Write("bgcolor='#FFFF99'")%>
				<%else%>
                	bgcolor="<%=corr%>"<%end if%>>
		          <td><div align="left">
<a href="lca_correcao_diretoria_autenticacao.asp?COD_VENDEDOR=<%=rs2("COD_VENDEDOR")%>&cod_fechamento_periodo=<%=cod_fechamento_periodo%>"><%=Ucase(rs2("NOME_VENDEDOR"))%></a>
				  </div></td>
                  <td><div align="center" style="height:15px; background-color:#FA9A9A; padding-top:2px"><strong>
				 <%
				  Dim rsContVerificado2, qtdPedidosTotais2
				  
				  Set rsContVerificado2 = objConexao.Execute("SELECT COUNT(*) AS XCONTA_LINHA FROM TB_PEDIDOS WHERE FINALIZA_CORRECAO = 1 AND DATA_DESCEU_DIRETORIA IS NOT NULL and FINALIZA_DIRETORIA = 2  AND cod_fechamento_periodo = '"&cod_fechamento_periodo&"' AND COD_VENDEDOR = "&rs2("COD_VENDEDOR"))
				  
				  	Response.Write(rsContVerificado2("XCONTA_LINHA"))
				  qtdB = 0
				  qtdB = formatNumber(0&rsContVerificado2("XCONTA_LINHA"))/1
				  qtdA = qtdA + qtdB  
				  rsContVerificado2.Close
				  Set rsContVerificado2 = Nothing
		
	
				  %>
				  
				  
				  </strong></div></td>
                  <td><div align="center" style="height:15px; background-color:#00FF66; padding-top:2px"><strong>
				 <%
				  Dim rsContVerificado11, qtdPedidosTotais11
				  
				  sql = "SELECT COUNT(*) AS XCONTA_LINHA FROM TB_PEDIDOS WHERE  FINALIZA_DIRETORIA = 1 AND cod_fechamento_periodo = "&cod_fechamento_periodo&"  and  COD_VENDEDOR = "&rs2("COD_VENDEDOR")
				  
				 
				  
				  Set rsContVerificado11 = objConexao.Execute(sql)
				  
				  Response.Write(rsContVerificado11("XCONTA_LINHA"))
					
				  qtdB1 = 0
				  qtdB1 = formatNumber(0&rsContVerificado11("XCONTA_LINHA"))/1
				  qtdA1 = qtdA1 + formatNumber(qtdB1)/1
				  'Response.Write(qtdB1)
				  rsContVerificado11.Close
				  Set rsContVerificado11 = Nothing
				  %>
				  </strong></div>				  </td>
			      <td><div align="center"><strong><%=qtdB+qtdB1%></strong></div></td>
				  </tr>
					 
					 
          <%
		  	rs2.MoveNext()
			Wend
			
			rs2.Close
			Set rs2=Nothing
			FecharConexao()
			
		  %>  
		  
		  <tr>
		  <td>&nbsp;</td>
		  <td>
		    <div align="center"><strong><%=qtdA%></strong></div></td>
		  		  <td>
		            <div align="center"><strong><%=qtdA1%>           </strong></div></td>
		          <td><div align="center"><strong><%=qtdA+qtdA1%></strong></div></td>
		          </tr>  
        </table>
           
			<p align="center">
		  <input class="btnTopo" onclick="abre('#')" type="button" name="btnTopo2" value=" " />
		</p>		</td>
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