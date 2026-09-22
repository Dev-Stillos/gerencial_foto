<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<%
Dim comando, botao, ACAO
Dim CODIGO_PRODUCAO, DESCRICAO, DATA_SAIDA, DATA_SAIDA_DIA, DATA_SAIDA_MES, DATA_SAIDA_ANO, PRODUTOS_AGREDADOS, PRODUCAO_FINALIZADA, ENCADERNACAO_ESPECIAL

comando = Request.QueryString("comando")
CODIGO_PRODUCAO = Request.QueryString("cod_prod")

If Request.Form("GRAVAR") <> Empty then

	Dim QUANTIDADE, COD_PRODUTO
	
	COD_PRODUTO = 0&Request.Form("COD_PRODUTO")
	QUANTIDADE  = 0&Request.Form("QUANTIDADE")
	
	COD_PRODUTO = formatNumber(COD_PRODUTO,0)
	QUANTIDADE	= formatNumber(QUANTIDADE,0)
	
	Response.Write(COD_PRODUTO)
	If COD_PRODUTO <> 0 and QUANTIDADE <> 0 then
	'Grava o produto agregado
	
		AbrirConexao()
		
		sql = "INSERT INTO TB_PRODUCAO_PRODUTOS_AGREGADOS ("&_
								"COD_PRODUCAO , "&_
								"COD_PRODUTO , "&_
								"QUANTIDADE , "&_
								"CADASTRADO_POR ) "&_
								"VALUES ("&_
								CODIGO_PRODUCAO&", "&COD_PRODUTO&", "&QUANTIDADE&", "&Session("COD_USUARIO")&")"
		
			
		
		
		objConexao.Execute(sql)
		
		'Response.Write(sql)
		
		FecharConexao()
		
	
	End If
	
	Response.Redirect("dip_producoes_produtos_agregados.asp?cod_prod="&CODIGO_PRODUCAO)

End If

If Request.QueryString("comando") = "excluir" then

	AbrirConexao()
	
		objConexao.Execute("DELETE FROM TB_PRODUCAO_PRODUTOS_AGREGADOS WHERE COD_PRODUCAO="&CODIGO_PRODUCAO&" AND COD_PRODUTO="&Request.QueryString("cod_produto"))
	
	FecharConexao()
	
	Response.Redirect("dip_producoes_produtos_agregados.asp?cod_prod="&CODIGO_PRODUCAO)

End If	
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
	<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>
	
	<script language="javascript">
		function popUp(pagina){
		
		
		}
		
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
	
	<script language="javascript">
	
	function validaForm(){
	d = document.producao
	
		if (producao.DESCRICAO.value == ""){
			alert("Preencha o campo Descrição!");
			producao.DESCRICAO.focus();
			return false;
		}
		
		if (producao.DATA_SAIDA_DIA.value == "" || producao.DATA_SAIDA_MES.value == "" || producao.DATA_SAIDA_ANO.value == ""){
			alert("Preencha a data completa da saída para venda.");
			producao.DATA_SAIDA_DIA.focus();
			return false;
		}		
	
	}
	
	</script>
</head>

<body>
<!-- Code for the left panel -->
<div id="dhtmlgoodies_leftPanel">
	<a class="closeLink" href="#" onclick="initSlideLeftPanel();return false" style="font-size:15px">Fechar</a>
	<div id="leftPanelContent">
	<!-- This is the content -->
	<div>

<table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
  <td width="1%"><img src="images/icons/icon_help.jpg" /></td>
    <td width="95%" align="left"><strong style="font-size:16px">Ajuda do Sistema</strong><br />
<strong>Esta é a tela inicial do sistema, verifique o número de acessos ao sistema e a data da última utilização.<br />
</strong>
	</td>
  </tr>
</table>

	</div>
	<!-- End content -->
	</div>
</div>
<!-- End code for the left panel -->
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

	  <h2>DIP - Produ&ccedil;&otilde;es / Cadastro e
	  Altera&ccedil;&otilde;es</h2>
	  
	  <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
		  <h3 align="center">Produ&ccedil;&atilde;o - Inclus&atilde;o de Produtos Agregados</h3>
		 	<form name="FORMULARIO"  method="post" action="dip_producoes_produtos_agregados.asp?cod_prod=<%=CODIGO_PRODUCAO%>&comando=incluir" style="margin:0px">
		  <div id="comboTexto">
          <table width="90%" border="0" align="center" 	cellpadding="5">
<tr>
            <td width="9%"><strong>Produto:</strong></td>
            <td><select name="COD_PRODUTO">
                 <option value="0">---</option>
				 <%
				 AbrirConexao()
				 Dim rsProdutos, sql
				 
				 Set rsProdutos = objConexao.Execute("SELECT COD_PRODUTO, NOME_PRODUTO FROM TB_PRODUTOS WHERE ATIVO = 1 ORDER BY ORDER_MENU")
				 
				 While not rsProdutos.Eof
				 %>
					 <option value="<%=rsProdutos("COD_PRODUTO")%>"><%=rsProdutos("NOME_PRODUTO")%></option>
				 <%
				 rsProdutos.MoveNext()
				 Wend
				 %>
						 
             
               	<%
				rsProdutos.Close()
				Set rsProdutos = Nothing
				
				'Busca foto padrão
				Dim rsPadrao, COD_FOTO_PADRAO, NOME_FOTO_PADRAO
				
				sql =  	"SELECT TB_PRODUTOS.NOME_PRODUTO, TB_PRODUCAO.COD_FOTO_PADRAO "&_
						"FROM TB_PRODUTOS "&_
						"INNER JOIN TB_PRODUCAO ON TB_PRODUTOS.COD_PRODUTO = TB_PRODUCAO.COD_FOTO_PADRAO "&_
						"WHERE TB_PRODUCAO.COD_PRODUCAO = "&CODIGO_PRODUCAO
				
				Set rsPadrao = objConexao.Execute(sql)
				
					If not rsPadrao.Eof then
					
						COD_FOTO_PADRAO = rsPadrao("COD_FOTO_PADRAO")
						NOME_FOTO_PADRAO =  rsPadrao("NOME_PRODUTO")
					End If
					
				
				rsPadrao.Close()
				Set rsPadrao = Nothing
				
				
				
				FecharConexao()
				%>
			   </select></td>
            </tr>
<tr>
  <td colspan="2"><input style="width:35px" type="hidden" value="1" name="QUANTIDADE" />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input class="btnGravar" type="submit" name="GRAVAR" value=" " onclick="return validaForm()" /></td>
  </tr>
        </table>
		</div>
		  </form>
		  <p align="left"> <%
		 AbrirConexao()
		 	 Dim rs
			 sql =  " SELECT TB_PRODUCAO_PRODUTOS_AGREGADOS.COD_PRODUCAO_AGREGADO, TB_PRODUCAO_PRODUTOS_AGREGADOS.COD_PRODUTO,  Sum(TB_PRODUCAO_PRODUTOS_AGREGADOS.QUANTIDADE) as QUANTIDADE,  TB_PRODUTOS.NOME_PRODUTO"&_
					" FROM TB_PRODUTOS INNER JOIN TB_PRODUCAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = TB_PRODUCAO_PRODUTOS_AGREGADOS.COD_PRODUTO"&_
					" WHERE TB_PRODUCAO_PRODUTOS_AGREGADOS.COD_PRODUCAO="&CODIGO_PRODUCAO&" GROUP BY TB_PRODUCAO_PRODUTOS_AGREGADOS.COD_PRODUTO ORDER BY TB_PRODUTOS.NOME_PRODUTO"
		 
			 Set rs = objConexao.Execute(sql)
		 	 If not rs.Eof then
		 %>
		 <table width="90%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse; margin-top:6px">
           <tr bgcolor="whitesmoke">
             <td width="66%"><div align="center"><strong>Produtos Agregados </strong></div></td>
             <td width="18%"><div align="center"><strong>Qtd.</strong></div></td>
             <td width="16%"><div align="center"><strong>Excluir</strong></div></td>
           </tr>
           <%While not rs.Eof%>
           <tr>
             <td><div align="left"><%=rs("NOME_PRODUTO")%></div></td>
             <td><div align="center"><%=rs("QUANTIDADE")%></div></td>
             <td><div align="center"><a onClick="anula()" href="?cod_prod=<%=CODIGO_PRODUCAO%>&id=<%=rs("COD_PRODUCAO_AGREGADO")%>&comando=excluir&cod_produto=<%=rs("COD_PRODUTO")%>"><img src="images/icons/icon_excluir.png" width="16" height="16" border="0" /></a></div></td>
           </tr>
           <%
			rs.MoveNext()
			Wend
			%>
                    <tr>
             <td colspan="3"><div align="center">
               <p>&nbsp;                 </p>
               <p>
                 <input class="btnAvancar" type="button" onclick="abre('dip_producoes_importar_formandos.asp?cod_prod=<%=Request.QueryString("cod_prod")%>')" name="GRAVAR2" value=" " />
                 </p>
             </div></td>
             </tr>
		 </table>
		 <%
		  	 End If
			
			
			  rs.Close()
			  Set rs = Nothing
		  FecharConexao()
		  %>		  </td>
      </tr>
    </table></p></td>
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