<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim cod_produto, nome_produto, cod_valor_vigente
cod_produto = Request.QueryString("cod_produto")

AbrirConexao()
	
			Set rs2=objConexao.Execute("Select NOME_PRODUTO From TB_PRODUTOS where COD_PRODUTO = '"&COD_PRODUTO&"'")
				
				If not rs2.Eof Then
					nome_produto = rs2(0)
				End If
				
			rs2.Close
			Set rs2=Nothing
			
FecharConexao()
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
	  <div class="subMenu">Sub-menu: <a href="dip_produtos.asp">Voltar para Listagem de Produtos</a> | <a href="dip_produtos_tabela_formulario.asp?cod_produto=<%=cod_produto%>&comando=incluir">Incluir novo Preço para este produto</a></div>
	  <%
	   Response.Write("<br />")
	Dim sucesso, codmsg
	codmsg	= CStr(""&Request.QueryString("codmsg"))
	sucesso = CStr(""&Request.QueryString("sucesso"))
		
	If sucesso = "" and codmsg <> "" Then%>
    	<div align="center" class="msgGrandeSucesso">
        	<%
			If codmsg = "1" Then
				Response.Write("Preço incluído com sucesso!")
			Elseif codmsg = "2" Then
				Response.Write("Preço alterado com sucesso!")
			Elseif codmsg = "3" Then	
				Response.Write("Preço excluído com sucesso!")
			End If	
			%>
            
        
        </div>
    <%ElseIf sucesso <> "" and codmsg <> "" Then%>
        <div align="center" class="msgGrandeErro">
			<%
			If codmsg = "4" Then
				Response.Write("Produto não alterado para não haver duplicidade de informação!<BR>O conteúdo do campo """&Request.QueryString("chaveUnica")&""" já possui uma informação igual em outro Produto.<br />")
			Elseif codmsg = "5" Then
				Response.Write("Produto não exclu�do, pois há pedidos que dependem deste Produto.<BR>Se não desejar ter mais este Produto inative ao invés de excluir.<br />")
							
			End If	
			%>
</div>
	<%End If%>
      
        <h3 align="center">Tabela de Pre&ccedil;os do Produto &quot;<strong><%=nome_produto%></strong>&quot;<br />Vendedores Comum</h3>
	  	  <%
		  AbrirConexao()
		  
		  'Fun��o que retorna o c�d. do produto vigente, se for igual a 0 quer dizer que não produto vigente
		  cod_valor_vigente = valorProdutoVigente(cod_produto,true,0)
		  
		  Dim rs, rs2, rs3, sql

		  sql = "SELECT * FROM TB_PRODUTOS_VALORES WHERE COD_PRODUTO = '"&cod_produto&"' AND GRUPO_VENDEDORES = 0 order by VALIDO_ATE desc"
		   
		  Set rs = objConexao.Execute(sql)
		  
		  If not rs.Eof then
		  %>
			
            <%
			If int(cod_valor_vigente) = 0 Then
			%>
        	<div class="msgExcluir">Atenção este produto não possui Preço vigente, cadastre um Preço com data de validade igual ou superior a data atual, pois caso contrário o produto não poderá ser lançado nas vendas sem Preço vigente.</div>
            <br />
        	<%
			End if
			%>
		  <div id="comboTexto">
          <table width="95%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="17%"><div align="center"><strong>Valor</strong> <strong>&agrave; vista</strong></div></td>
            <td width="14%"><div align="center"><strong>Valor em 3X</strong></div></td>
            <td width="15%"><div align="center"><strong>Valor</strong> <strong>parcelado</strong></div></td>
            <td width="13%"><div align="center">
              <div align="center"><strong>Valido at&eacute;</strong></div>
            </div></td>
            <td width="15%"><div align="center"><strong>Data Cadastro</strong></div></td>
            <td width="16%"><div align="center"><strong>Cadastrado por</strong></div></td>
            <td width="10%"><div align="center"><strong>Alterar / Excluir</strong></div></td>
</tr>
		  <%
		  While not rs.Eof		  
		  %>
           <tr <%If int(cod_valor_vigente) = int(rs("cod_valor")) Then%> bgcolor="#FFFF99" <%End If%>>
            <td><div align="center"><%=formatCurrency(rs("VALOR_A_VISTA"))%></div></td>
            <td><div align="center"><%=formatCurrency(rs("VALOR_3X"))%></div></td>
            <td><div align="center"><%=formatCurrency(rs("VALOR_PARCELADO"))%></div></td>
            <td><div align="center" class="textoMedio"><strong><%=ArrumaData(rs("VALIDO_ATE"))%></strong></div></td>
            <td><div align="center"><%=ArrumaData(rs("DATA_CADASTRO"))%></div></td>
            <td>
			
			      <div align="center">
			        <%
					
					
					
			Set rs2=objConexao.Execute("Select nome From TB_ADM_USUARIOS where COD_USUARIO = '"&rs("CADASTRADO_POR")&"'")
				
				If not rs2.Eof Then
					Response.Write(rs2(0))
				End If
				
			rs2.Close
			Set rs2=Nothing
			%>
			        </div></td>
            <td>
              <div align="center">
              <a href="dip_produtos_tabela_formulario.asp?COD_VALOR=<%=rs("COD_VALOR")%>&comando=alterar&cod_produto=<%=cod_produto%>">
                  <img alt="Alterar" src="images/icons/icon_alterar.png" border="0" /></a>&nbsp;&nbsp;
              <a href="dip_produtos_tabela_formulario.asp?COD_VALOR=<%=rs("COD_VALOR")%>&comando=excluir&cod_produto=<%=cod_produto%>">
                          <img alt="Excluir" src="images/icons/icon_excluir.png" border="0" /></a>                </div></td>
           </tr>
		  <%
		  rs.MoveNext()
		  Wend
	      
		  %>
        </table>
        </div>
		
		  
		    <p>
		      <%else%>
              <div align="center">
		      <em class="textoMedio">não há pre&ccedil;os cadastrados para este produto e para os vendedores comum,<br />
		      clique no sub-menu acima em &quot;Incluir novo pre&ccedil;o para este produto&quot;</em>
		      <p>&nbsp;</p>
		    <p>	      
             </div>
		      <%end if
			  
			  
			  	  rs.Close()
				  Set rs = Nothing
				  FecharConexao()
			  %>
		      </p>
		 
         <BR />
         <HR />
         <h3 align="center">Tabela de Pre&ccedil;os do Produto &quot;<strong><%=nome_produto%></strong>&quot;<br />
         Vendedores Repasse</h3>
	  	  <%
		  AbrirConexao()
		  
		  'Fun��o que retorna o c�d. do produto vigente, se for igual a 0 quer dizer que não produto vigente
		  cod_valor_vigente = valorProdutoVigente(cod_produto,true,1)

		  sql = "SELECT * FROM TB_PRODUTOS_VALORES WHERE COD_PRODUTO = '"&cod_produto&"' AND GRUPO_VENDEDORES = 1 order by VALIDO_ATE desc"
		   
		  Set rs = objConexao.Execute(sql)
		  
		  If not rs.Eof then
		  %>
			
            <%
			If int(cod_valor_vigente) = 0 Then
			%>
        	<div class="msgExcluir">Atenção este produto não possui Preço vigente, cadastre um Preço com data de validade igual ou superior a data atual, pois caso contrário o produto não poderá ser lançado nas vendas sem Preço vigente.</div>
            <br />
        	<%
			End if
			%>
		  <div id="comboTexto">
          <table width="95%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="17%"><div align="center"><strong>Valor</strong> <strong>&agrave; vista</strong></div></td>
            <td width="14%"><div align="center"><strong>Valor em 3X</strong></div></td>
            <td width="15%"><div align="center"><strong>Valor</strong> <strong>parcelado</strong></div></td>
            <td width="13%"><div align="center">
              <div align="center"><strong>Valido at&eacute;</strong></div>
            </div></td>
            <td width="15%"><div align="center"><strong>Data Cadastro</strong></div></td>
            <td width="16%"><div align="center"><strong>Cadastrado por</strong></div></td>
            <td width="10%"><div align="center"><strong>Alterar / Excluir</strong></div></td>
</tr>
		  <%
		  While not rs.Eof		  
		  %>
           <tr <%If int(cod_valor_vigente) = int(rs("cod_valor")) Then%> bgcolor="#FFFF99" <%End If%>>
            <td><div align="center"><%=formatCurrency(rs("VALOR_A_VISTA"))%></div></td>
            <td><div align="center"><%=formatCurrency(rs("VALOR_3X"))%></div></td>
            <td><div align="center"><%=formatCurrency(rs("VALOR_PARCELADO"))%></div></td>
            <td><div align="center" class="textoMedio"><strong><%=ArrumaData(rs("VALIDO_ATE"))%></strong></div></td>
            <td><div align="center"><%=ArrumaData(rs("DATA_CADASTRO"))%></div></td>
            <td>
			
			      <div align="center">
			        <%
					
					
					
			Set rs2=objConexao.Execute("Select nome From TB_ADM_USUARIOS where COD_USUARIO = '"&rs("CADASTRADO_POR")&"'")
				
				If not rs2.Eof Then
					Response.Write(rs2(0))
				End If
				
			rs2.Close
			Set rs2=Nothing
			%>
			        </div></td>
            <td>
              <div align="center">
              <a href="dip_produtos_tabela_formulario.asp?COD_VALOR=<%=rs("COD_VALOR")%>&comando=alterar&cod_produto=<%=cod_produto%>">
                  <img alt="Alterar" src="images/icons/icon_alterar.png" border="0" /></a>&nbsp;&nbsp;
              <a href="dip_produtos_tabela_formulario.asp?COD_VALOR=<%=rs("COD_VALOR")%>&comando=excluir&cod_produto=<%=cod_produto%>">
                          <img alt="Excluir" src="images/icons/icon_excluir.png" border="0" /></a>                </div></td>
           </tr>
		  <%
		  rs.MoveNext()
		  Wend
	      
		  %>
        </table>
       </div>
		
		  
		    <p>
		      <%else%>
              <div align="center">
		      <em class="textoMedio">Não há pre&ccedil;os cadastrados para este produto e para os vendedores de repasse,<br />
		      clique no sub-menu acima em &quot;Incluir novo pre&ccedil;o para este produto&quot;</em>
		      <p>&nbsp;</p>
		    <p>	      
             </div>
		      <%end if
			  
			  
			  	  rs.Close()
				  Set rs = Nothing
				  FecharConexao()
			  %>
		      </p>
		 
         <BR><BR>
         <strong>Obs.:</strong> Linhas na cor amarela significa pre&ccedil;o que o sistema estar&aacute; utilizando no momento, exceto para boletins informativos com este produto.
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