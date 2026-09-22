<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Valida��o especial para verificar se possui sub-valida��o espec�fica
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '29' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se n�o encontrar ent�o redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
		
	rs.Close
	Set rs=Nothing
	
FecharConexao()








Dim MES_INICIO, ANO_INICIO, MES_FIM, ANO_FIM
Dim DATA_INICIO, DATA_FIM
MES_INICIO	= Request.QueryString("MES_INICIO")
ANO_INICIO	= Request.QueryString("ANO_INICIO")
'Somente mensal
MES_FIM	= Request.QueryString("MES_INICIO")
ANO_FIM	= Request.QueryString("ANO_INICIO")
'MES_FIM		= Request.QueryString("MES_FIM")
'ANO_FIM		= Request.QueryString("ANO_FIM")

DATA_INICIO	= ANO_INICIO&"-"&MES_INICIO&"-01"
DATA_FIM	= ANO_FIM&"-"&MES_FIM&"-01"
DATA_FIM	= DateAdd("m",1,DATA_FIM)
DATA_FIM	= DateAdd("d",-1,DATA_FIM)
%>
<%
'Variavel de Conexao
Dim objConexao3

'Sub Rotina que Abre a Conexao

Sub AbrirConexao3()
Set objConexao3 = Server.CreateObject("ADODB.Connection")

'Forma nova
objConexao3.Open("Provider=SQLOLEDB.1;Password=WyHoZ1Ijln4!g_;Persist Security Info=True;User ID=stillos6;Initial Catalog=stillos6;Data Source=sqlserver03.stillos.com.br")

End Sub

'Sub Rotina que Fecha a Conexao
Sub FecharConexao3()
   objConexao3.Close
   Set objConexao3=nothing
End Sub
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="all"></link>
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
	var text = "Se a janela n�o abrir desbloquei os Pop-ups!";
	if(win == null) { alert(text); return; }
	win.moveTo(350,150);
}	
	</script>
    <style type="text/css">
<!--
.style1 {font-size: 13px}
body {
	background-color: #FFFFFF;
}
-->
    </style>
</head>

<body>
<!--<strong><%=ArrumaData(DATA_INICIO)%></strong> a <strong><%=ArrumaData(DATA_FIM)%> - <%=Ucase(MOnthName(month(DATA_FIM)))%></strong><br />-->
	  <%
		AbrirConexao()
		AbrirConexao3()
		
		Dim rs, sql
			  
		sql = "SELECT COD_PEDIDO, XTOTAL_CHEQUE, XTOTAL_CARTAO, PAGAMENTO_EM_DINHEIRO, DATA_VENDA FROM VW_FECHAMENTO_GERAL WHERE DATA_VENDA BETWEEN "&sqlData(DATA_INICIO)&" AND "&sqlData(DATA_FIM)&" ORDER BY DATA_VENDA"

		
		Set rs = objConexao.Execute(sql, qtdRegistros)
		
		If not rs.Eof Then

		%>
      <table width="100%" border="1" align="center"	cellpadding="5" style="border-collapse:collapse">

<%
		Dim COD_PEDIDO
		Dim TOTAL_DINHEIRO, TOTAL_CHEQUE, TOTAL_CARTAO, TOTAL_FORMA_PAGAMENTO
		Dim TOTAL_GERAL_PAG
		Dim TOTAL_GERAL_DINHEIRO, TOTAL_GERAL_CHEQUE, TOTAL_GERAL_CARTAO
		
		TOTAL_GERAL_DINHEIRO	= 0	
		TOTAL_GERAL_CHEQUE		= 0
		TOTAL_GERAL_CARTAO		= 0
		
		While not rs.Eof
		
		
		COD_PEDIDO = rs("COD_PEDIDO")
		'Forma de Pagamento
				
					TOTAL_DINHEIRO	= formatNumber(0&rs("PAGAMENTO_EM_DINHEIRO"))/1
					TOTAL_CHEQUE	= formatNumber(0&rs("XTOTAL_CHEQUE"))/1
					TOTAL_CARTAO	= formatNumber(0&rs("XTOTAL_CARTAO"))/1
					TOTAL_FORMA_PAGAMENTO = (TOTAL_DINHEIRO + TOTAL_CHEQUE + TOTAL_CARTAO)
					TOTAL_GERAL_DINHEIRO = (TOTAL_GERAL_DINHEIRO + TOTAL_DINHEIRO)/1
					TOTAL_GERAL_CHEQUE = (TOTAL_GERAL_CHEQUE + TOTAL_CHEQUE)/1
					TOTAL_GERAL_CARTAO = (TOTAL_GERAL_CARTAO + TOTAL_CARTAO)/1
					TOTAL_GERAL_PAG = (TOTAL_GERAL_PAG + TOTAL_FORMA_PAGAMENTO)/1
		%>
        
					<tr>
		          <td><div align="center"><%=rs("DATA_VENDA")%></div></td>
		          <td>
                    <div align="center"><%=COD_PEDIDO%></div></td>
                  <td width="4%"><div align="right"><%=formatNumber(TOTAL_DINHEIRO)%></div></td>
                  <td width="4%"><div align="right"><%=formatNumber(TOTAL_CHEQUE)%></div></td>
                  <td width="4%"><div align="right"><%=formatNumber(TOTAL_CARTAO)%></div></td>
                  <td width="5%"><div align="right"><%=formatNumber(TOTAL_FORMA_PAGAMENTO)%></div></td>
                  <td width="12%">
                  
                  <table width="100%" border="1" align="center"	cellpadding="3" style="border-collapse:collapse">
                   <%
				   '1) �lbum
				   '2) Prod. Agregado
				   '3) Produtos Extras
				   %>
                    
                    <%
					Set rs3=objConexao.Execute("SELECT COD_FORMANDO FROM TB_PEDIDOS_FORMANDOS WHERE COD_PEDIDO="&COD_PEDIDO)
					
					If not rs3.Eof Then
					
					While not rs3.Eof
					
						Set rs4=objConexao.Execute("SELECT FATURADO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO = "&rs3("COD_FORMANDO"))
						
						If not rs4.Eof Then
						
						Dim FATURADO
						FATURADO = formatNumber(0&rs4("FATURADO"),0)
						
								If FATURADO > 0 Then
						%>
        	            <tr>
            	          <td width="55%">Fotos</td>
                	      <td width="45%"><%=FATURADO%></td>
                    	</tr>
	                    <%
						End If
						
								Set rsAgregados = objConexao.Execute(	" SELECT TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO, TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO_AGREGADO, TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.TIPO, "&_
												" Sum(TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.QUANTIDADE) as QUANTIDADE, "&_
												" TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.VENDIDO, TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.VALOR_UN, "&_
												" TB_PRODUTOS.NOME_PRODUTO FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS "&_
												" ON TB_PRODUTOS.COD_PRODUTO = TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO "&_
												" WHERE TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO="&rs3("COD_FORMANDO")&" AND TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.VENDIDO = 1"&_
												" GROUP BY TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO ")
								
								
								If not rsAgregados.Eof Then
								
								While not rsAgregados.Eof
								%>
								<tr>
								  <td width="55%"><%=rsAgregados("NOME_PRODUTO")%></td>
								  <td width="45%"><%=rsAgregados("QUANTIDADE")%></td>
								</tr>
								
								<%
								rsAgregados.MoveNext()
								Wend
								
								rsAgregados.Close
								Set rsAgregados = Nothing
								
								End If
						
								
								Set rsProdutosAparte = objConexao.Execute(" SELECT TB_PEDIDOS_PRODUTOS.COD_PEDIDO_PRODUTO, TB_PRODUTOS.NOME_PRODUTO, TB_PEDIDOS_PRODUTOS.VALOR_UN, "&_
										  " TB_PEDIDOS_PRODUTOS.QUANTIDADE FROM TB_PRODUTOS INNER JOIN "&_
										  " TB_PEDIDOS_PRODUTOS ON TB_PRODUTOS.COD_PRODUTO = TB_PEDIDOS_PRODUTOS.COD_PRODUTO"&_
										  " WHERE COD_PEDIDO="&COD_PEDIDO&" AND COD_FORMANDO="&rs3("COD_FORMANDO")&" ORDER BY NOME_PRODUTO")

								If not rsProdutosAparte.Eof Then
											
											While not rsProdutosAparte.Eof
												
												%>
                                                <tr>
												  <td width="55%"><%=rsProdutosAparte("NOME_PRODUTO")%></td>
												  <td width="45%"><%=rsProdutosAparte("QUANTIDADE")%></td>
												</tr>
                                                <%
												
											rsProdutosAparte.MoveNext()
											Wend
											
								
								rsProdutosAparte.Close
								Set rsProdutosAparte = Nothing

								End If
								
						End If
						
						
						rs4.Close
						Set rs4=Nothing
					
					
					rs3.MoveNext()
					Wend
					
					End If
					
					rs3.Close
					Set rs3=Nothing
					%>
                  </table></td>
                  <td width="63%" colspan="2"><table width="100%" border="1" align="center"	cellpadding="3" style="border-collapse:collapse">
                    <%
					Set rs2=objConexao.Execute("SELECT COD_FORMANDO FROM TB_PEDIDOS_FORMANDOS WHERE COD_PEDIDO="&COD_PEDIDO)
					
					Dim NOME_FORMANDO, ENDERECO, CEP, FONE_RES, BAIRRO, CIDADE, ESTADO
					
					While not rs2.Eof
					
					Set rs3=objConexao.Execute("SELECT NOME_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO ="&rs2("COD_FORMANDO"))
						
						If not rs3.Eof Then
							NOME_FORMANDO = rs3("NOME_FORMANDO")
						End If
							
					rs3.Close
					Set rs3=Nothing
					
					Set rs3=objConexao.Execute("SELECT * FROM  TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO WHERE COD_FORMANDO = "&rs2("COD_FORMANDO")&" ORDER BY DATA_CADASTRO DESC LIMIT 1")
						
						If not rs3.Eof Then
							ENDERECO		= rs3("ENDERECO")
							CEP				= rs3("CEP")
							FONE_RES		= rs3("FONE_RES")
							BAIRRO			= rs3("BAIRRO")
							CIDADE			= rs3("CIDADE")
							ESTADO			= rs3("ESTADO")
						End If
						
					rs3.Close
					Set rs3=Nothing
					%>
                    <tr>
                      <td width="14%"><div align="left"><%=NOME_FORMANDO%></div></td>
                      <td width="8%">
                      <%
						Dim codFormando
						codFormando = Mid(rs2("COD_FORMANDO"), 1, Len(rs2("COD_FORMANDO")) - 1)

						If IsNumeric(codFormando) Then
						Set rs6 = objConexao3.Execute("SELECT FD_CPF_004 FROM TB_FORMANDOS WHERE FD_CODIG_004 = " & codFormando)

						If Not rs6.Eof Then
							Response.Write(rs6("FD_CPF_004"))
						End If

						rs6.Close
						Set rs6 = Nothing
						End If
						%>                      </td>
                      <td width="13%"><%=FONE_RES%></td>
                      <td width="19%"><%=ENDERECO%></td>
                      <td width="9%"><%=CEP%></td>
                      <td width="14%"><%=BAIRRO%></td>
                      <td width="17%"><%=CIDADE%></td>
                      <td width="6%"><%=ESTADO%></td>
                    </tr>
                    <%
					rs2.MoveNext()
					Wend
					
					rs2.Close
					Set rs2=Nothing
					%>
                  </table></td>
					</tr>
								 
		<%
			
			rs.MoveNext()
			Wend
			
			rs.Close
			Set rs = Nothing
			
			FecharConexao()
			FecharConexao3()
			%>
            <tr>
					  <td colspan="9">Total em cart&atilde;o: <%=formatCurrency(TOTAL_GERAL_CARTAO)%><br />
					    Total em dinheiro: <%=formatCurrency(TOTAL_GERAL_DINHEIRO)%><br />
					    Total em cheque: <%=formatCurrency(TOTAL_GERAL_CHEQUE)%><br />
					    <br />
					    <span class="style1">Total Geral: <strong><%=formatCurrency(TOTAL_GERAL_PAG)%></strong></span></td>
					  </tr>
			</table>
		
		  <p><br />
		  </p>
		    <%Else%>
		   <div align="center">
           <br /><br /><br /><br />
		  <p><span class="textoMedio"><strong>N�o h� dados para os filtros selecionados!</strong></span></p>
          <br /><br /><br /><br />
		</div>
		    <%End If%>

</body>
</html>



