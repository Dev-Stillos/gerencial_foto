<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim cod_fechamento_periodo, COD_VENDEDOR, NOME_VENDEDOR
Dim rs, rs3, rsNomeVend, sql, sqlBase

Dim PRODUTOS, DINHEIRO, CARTAO, CHEQUE
Dim SUB_TOTAL_EM_PRODUTOS, SUB_TOTAL_EM_DINHEIRO, SUB_TOTAL_EM_CARTAO, SUB_TOTAL_EM_CHEQUE
Dim TOTAL_EM_PRODUTOS, TOTAL_EM_DINHEIRO, TOTAL_EM_CARTAO, TOTAL_EM_CHEQUE

TOTAL_EM_PRODUTOS	= 0
TOTAL_EM_DINHEIRO	= 0
TOTAL_EM_CARTAO		= 0
TOTAL_EM_CHEQUE		= 0

COD_VENDEDOR	= Request.QueryString("COD_VENDEDOR")
ANO_VENDA 		= Request.QueryString("ANO_VENDA")
MES_VENDA 		= Request.QueryString("MES_VENDA")
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

	  <h2>LCA - Corre&ccedil;&atilde;o Diretoria / Confirma&ccedil;&atilde;o de Baixa </h2>
	   <h3 align="center">Fichas pendentes para o pagamento   do(a) vendedor(a) <%=NOME_VENDEDOR%></h3>
       <table width="80%" 
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

					    <div class="msg" align="center">Tem certeza que deseja baixar estes pedidos?</div>
						
						
						<form name="vendas" method="post" action="lca_correcao_diretoria_autenticacao_confirmacao_processar.asp?COD_VENDEDOR=<%=COD_VENDEDOR%>&cod_fechamento_periodo=<%=cod_fechamento_periodo%>">
          <%
		 Dim contaRegistrosConfirma
		 contaRegistrosConfirma = 0 
sqlBase = " SELECT TB_PEDIDOS.COD_PEDIDO, TB_PEDIDOS.DATA_VENDA, TB_PEDIDOS.ERRO_LISTA, TB_PEDIDOS_FORMANDOS.COD_FORMANDO, "&_
	 			" TB_PEDIDOS.FINALIZA_DIRETORIA, TB_PEDIDOS.DATA_CORRIGIDO_DIRETORIA, "&_
			" TB_FICHAS_IDENTIFICACAO.COD_CONTRATO "&_
				" FROM TB_PEDIDOS "&_
				" LEFT JOIN TB_PEDIDOS_FORMANDOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_FORMANDOS.COD_PEDIDO "&_
				" LEFT JOIN TB_FICHAS_IDENTIFICACAO ON TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO WHERE "&_
				" TB_PEDIDOS.COD_VENDEDOR="&COD_VENDEDOR&" AND TB_PEDIDOS.cod_fechamento_periodo = '"&cod_fechamento_periodo&"' AND TB_PEDIDOS.DATA_DESCEU_DIRETORIA IS NOT NULL GROUP BY TB_PEDIDOS.COD_PEDIDO ORDER BY TB_FICHAS_IDENTIFICACAO.COD_CONTRATO"
		
		
		  AbrirConexao()
		  'AbrirConexao2()
		  'sql = sqlBase & " GROUP BY TB_PEDIDOS.COD_PEDIDO ORDER BY COD_CONTRATO"
		  'Response.Write(sql)
		  
		  Set rs = objConexao.Execute(sqlBase)
		 %>
					
					
					<table width="80%" 
		      border="1" align="center"
					cellpadding="5" bgcolor="#FFFFFF" 
					style="border-collapse:collapse">
  <tr bgcolor="whitesmoke">   
    <td width="9%"><div align="center"><strong>C&oacute;d. Pedido </strong></div></td>
		      <td width="10%"><div align="center"><strong>&Aacute;lbum</strong></div></td>
		      <td width="27%"><div align="center" class="style6"><strong>Formando</strong></div></td>
		      <td width="27%"><div align="center"><strong>Curso</strong></div></td>
		      <td width="27%"><div align="center"><strong>Institui&ccedil;&atilde;o de Ensino </strong></div></td>
  </tr>
		
			<%
			While not rs.Eof
			%>

<%
If Request.Form("COMBO_"&rs("COD_PEDIDO")) = 1 Then
contaRegistrosConfirma = contaRegistrosConfirma + 1
%>
  <tr>
    <td id="borderBaixo"><div align="center" class="textoMedio"><strong>
	<%=rs("COD_PEDIDO")%>
	<input type="hidden" name="COMBO_<%=rs("COD_PEDIDO")%>" value="1" />
	</strong></div></td>
    <td colspan="4" id="borderBaixo"><div align="left">
			<%
			sql = 	"SELECT TB_FICHAS_IDENTIFICACAO.INICIAL, TB_FICHAS_IDENTIFICACAO.FATURADO, TB_FICHAS_IDENTIFICACAO.RETIRADO, "&_
					"TB_FICHAS_IDENTIFICACAO.NOME_CURSO, TB_PEDIDOS_FORMANDOS.COD_FORMANDO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, TB_FICHAS_IDENTIFICACAO.ALBUM,  TB_FICHAS_IDENTIFICACAO.COD_CONTRATO "&_
					"FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON "&_
					"TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO "&_
					"WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&rs("COD_PEDIDO")&" ORDER BY NOME_FORMANDO"
					
					'Response.Write(sql)
			Set rs3 = objConexao.Execute(sql)
				
		
				While not rs3.Eof
				
				%>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				  <tr>
					<!--<td width="23%"><div align="center"><%=digito(rs3("COD_FORMANDO"))%></div></td>-->
					<td width="10%"><div align="center"><%=rs3("ALBUM")%></div></td>
					<td width="31%">					  
					  
					    
					    <div align="left"><%=rs3("NOME_FORMANDO")%></div></td>
					<td width="30%"><div align="left"><%=rs3("NOME_CURSO")%></div>					  <div align="center"></div>					  <div align="center"></div>					  <div align="center"></div></td>
					<td width="29%"><%
					Dim rs6
					Set rs6 = objConexao.Execute("SELECT DESCRICAO FROM TB_ENTIDADES WHERE COD_CONTRATO= "&sqlTexto(rs3("COD_CONTRATO")))
						
						If not rs6.Eof Then
							Response.Write(rs6("DESCRICAO"))
						Else
							Response.Write(rs("COD_CONTRATO"))
						End If
						
					rs6.Close
					Set rs6 = Nothing
					%></td>
				  </tr>
				</table>
				<%
				rs3.MoveNext()
				Wend
				
			rs3.Close
			Set rs3 = Nothing
			%>
			</div></td>
    </tr>
<%End If%>  		
            <%
			rs.MoveNext()
		    Wend
			
			rs.Close
			Set rs=Nothing%>
		   <tr>
    <td colspan="5" bgcolor="#E8E8E8"><div align="center" class="textoMedio"><strong>Total de Pedidos Verificados: <%=contaRegistrosConfirma%></strong></div></td>
    </tr>
	      </table>

		            <div align="center">
		<%		  
		  'FecharConexao2()
		  FecharConexao()
		  
		  %> 
                      <br />
		              <input class="btnVoltar" type="button" onclick="javascript:history.go(-1)" name="Submit2" value=" " />
		              &nbsp;&nbsp;&nbsp; 
		              <input class="btnAvancar" type="submit" name="Submit" value=" " />
		              <br /><br />
		              <br />
          </div>
		            <br />
		            <br />   
</form>
						
			            <p align="center">&nbsp;</p>		</td>
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