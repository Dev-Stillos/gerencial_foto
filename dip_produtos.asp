<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>Produtos</title>
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
 


	<!--AJUDA F1 E CTRL PARA BUSCA RáPIDA-->
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

	  <h2>DIP - Produtos e Valores </h2>
	  <!-- #include file="includes/adm_produtos_sub_menu.asp" --> 
	  <%
	   Response.Write("<br />")
	Dim sucesso, codmsg
	codmsg	= CStr(""&Request.QueryString("codmsg"))
	sucesso = CStr(""&Request.QueryString("sucesso"))
		
	If sucesso = "" and codmsg <> "" Then%>
    	<div align="center" class="msgGrandeSucesso">
        	<%
			If codmsg = "1" Then
				Response.Write("Produto incluÍdo com sucesso!")
			Elseif codmsg = "2" Then
				Response.Write("Produto alterado com sucesso!")
			Elseif codmsg = "3" Then	
				Response.Write("Produto excluÍdo com sucesso!")
			End If	
			%>
            
        
        </div>
    <%ElseIf sucesso <> "" and codmsg <> "" Then%>
        <div align="center" class="msgGrandeErro">
			<%
			If codmsg = "4" Then
				Response.Write("Produto Não alterado para Não haver duplicidade de informação!<BR>O conteúdo do campo """&Request.QueryString("chaveUnica")&""" já possui uma informação igual em outro Produto.<br />")
			Elseif codmsg = "5" Then
				Response.Write("Produto Não excluído, pois há pedidos que dependem deste Produto.<BR>Se Não desejar ter mais este Produto inative ao invés de excluir.<br />")
							
			End If	
			%>
</div>
	<%End If%>
      
      
	  	  <%
		  AbrirConexao()
		  Dim rs, rs2, rs3, sql

		  sql = "SELECT TB_PRODUTOS.*,TB_PRODUTOS_CATEGORIAS.NOME_CATEGORIA FROM TB_PRODUTOS LEFT JOIN TB_PRODUTOS_CATEGORIAS ON TB_PRODUTOS.COD_CATEGORIA = TB_PRODUTOS_CATEGORIAS.COD_CATEGORIA ORDER BY COD_CATEGORIA, NOME_PRODUTO"
		   
		  Set rs = objConexao.Execute(sql)
		  
		  If not rs.Eof then
		  %>
		  <h3 align="center">Produtos</h3>
		  <div id="comboTexto">
          <table width="95%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="33%"><div align="center"><strong>Produto</strong></div></td>
            <td width="28%"><div align="center"><strong>Categoria</strong></div></td>
            <td width="6%"><div align="center"><strong>Ativo</strong></div></td>
            <td width="11%"><div align="center"><strong>Baixa Automática</strong></div></td>            
            <td width="12%"><div align="center"><strong>Consultar Pre&ccedil;os</strong></div></td>
            <td width="10%"><div align="center"><strong>Alterar / Excluir</strong></div></td>
			
</tr>
		  <%
		  While not rs.Eof
		  	 SAIDA_AUTOMATICA = rs("SAIDA_AUTOMATICA")
			'Dim DATA_ATUAL
		  	'Dim COD_PRODUTO, NOME_PRODUTO, COD_CATEGORIA, TIPO
		  	'DATA_ATUAL 		= Date()
			'COD_PRODUTO 	= rs("COD_PRODUTO")
		  	'NOME_PRODUTO 	= rs("NOME_PRODUTO")
		  	'COD_CATEGORIA	= rs("COD_CATEGORIA")
		  	'TIPO			= rs("TIPO")
			
			'Verifica se h� valores dentro do per�odo vigente
			'Dim sql1
			
			'sql1 = "SELECT * FROM TB_PRODUTOS_VALORES WHERE COD_PRODUTO ="&_
			'COD_PRODUTO&" AND VALIDO_ATE >="&sqlData(DATA_ATUAL)
			
			'Set rs2 = objConexao.Execute(sql1)
			
				'If not rs2.Eof then
				
					'Dim VALIDO_ATE, VALOR_A_VISTA, VALOR_PARCELADO
					
					'VALIDO_ATE 		= rs2("VALIDO_ATE")
					'VALOR_A_VISTA	= formatCurrency(rs2("VALOR_A_VISTA"))
					'VALOR_PARCELADO	= formatCurrency(rs2("VALOR_PARCELADO"))
				
				'Else
				
					'VALIDO_ATE		= "Não h� pre�o cadastrado e/ou o pre�o cadastrado Não est� vigente"
					'VALOR_A_VISTA	= "-"
					'VALOR_PARCELADO	= "-"
				
				'End If
			
			'rs2.Close()
		  	'Set rs2 = Nothing
		  
		  %>
           <tr <%if int(rs("ATIVO")) = 2 Then%> bgcolor="#E4E4E4" <%End If%>>
            <td><div align="left"><%=rs("NOME_PRODUTO")%></div></td>
            <td><div align="left"><%=rs("NOME_CATEGORIA")%>&nbsp;</div></td>
            <td><div align="center"><%if int(rs("ATIVO")) = 1 Then%>Sim
              <%Else%>
              Não
              <%End If%>
            </div></td>
        		 <td>
     		 		<div align="center">
					<%
					If CStr(SAIDA_AUTOMATICA) =  "0" Then%>
					Não
					<%Elseif CStr(SAIDA_AUTOMATICA) =  "1" Then%>
					Por produção
					<%Elseif CStr(SAIDA_AUTOMATICA) =  "2" Then%>
					Por venda					
					<%End if%>
           	 	</div>
            </td>            
            <td><div align="center"><a href="dip_produtos_tabela.asp?cod_produto=<%=rs("COD_PRODUTO")%>">Pre&ccedil;os deste produto</a></div></td>
			<% If Session("COD_USUARIO") <> 110 Then %>
            <td>
              <div align="center"><a href="dip_produtos_formulario.asp?cod_produto=<%=rs("COD_PRODUTO")%>&comando=alterar">
                  <img alt="Alterar" src="images/icons/icon_alterar.png" border="0" /></a>&nbsp;&nbsp;
                    <a href="dip_produtos_formulario.asp?cod_produto=<%=rs("COD_PRODUTO")%>&comando=excluir">
                          <img alt="Excluir" src="images/icons/icon_excluir.png" border="0" /></a>
                </div></td>
				<% End If %>
           </tr>
		  <%
		  rs.MoveNext()
		  Wend
	      
		  %>
        </table>
        <br /><br />
        <strong>Obs.:</strong> Linhas na cor cinza são produtos inativos.
        
		</div>
		
		  <div align="center">
		    <p>
		      <%else%>
		      <em>Não há produtos cadastrados! </em>	</p>
		    <p>&nbsp;</p>
		    <p>	      
		      <%end if
			  
			  
			  	  rs.Close()
				  Set rs = Nothing
				  FecharConexao()
			  %>
		      </p>
		  </div>
		  <p align="center">
		  <input class="btnTopo" onclick="abre('#')" type="button" name="btnTopo2" value=" " />
		</p>		
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
