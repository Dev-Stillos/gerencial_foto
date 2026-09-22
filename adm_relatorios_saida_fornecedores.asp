<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Validação especial para verificar se possui sub-validação específica
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '30' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se não encontrar então redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
		
	rs.Close
	Set rs=Nothing
	
FecharConexao()


Dim COD_FORNECEDOR, MES, ANO, TITULO, NOME_FORNECEDOR
Dim COD_FORNECEDOR_SAIDA_qs
Dim sql, rs, rs2, rs3

COD_FORNECEDOR			= formatNumber(0&Request.QueryString("COD_FORNECEDOR"),0)
MES						= ""&Request.QueryString("MES")
ANO						= ""&Request.QueryString("ANO")
COD_FORNECEDOR_SAIDA_qs	= formatNumber(0&""&Request.QueryString("COD_FORNECEDOR_SAIDA"),0)

If int(COD_FORNECEDOR) = 0 Then
	sql = "SELECT * FROM TB_FORNECEDORES_SAIDA WHERE MONTH(DATA_SAIDA) = '"&MES&"' AND YEAR(DATA_SAIDA) = '"&ANO&"' ORDER BY DATA_SAIDA ASC"
	TITULO = "Retiradas de "&Ucase(MonthName(MES))&"/"&ANO
Else
	sql = "SELECT * FROM TB_FORNECEDORES_SAIDA WHERE MONTH(DATA_SAIDA) = '"&MES&"' AND YEAR(DATA_SAIDA) = '"&ANO&"' AND COD_FORNECEDOR = "&COD_FORNECEDOR&" ORDER BY DATA_SAIDA ASC"
	
	
		AbrirConexao()
		Set rs=objConexao.Execute("SELECT NOME_FANTASIA FROM TB_FORNECEDORES WHERE COD_FORNECEDOR = "&COD_FORNECEDOR)
			
			If not rs.Eof Then
				TITULO = "Retiradas do fornecedor "&Ucase(rs("NOME_FANTASIA"))&" de "&Ucase(MonthName(MES))&"/"&ANO
			End If
		
		rs.Close
		Set rs=Nothing
		FecharConexao()
		
	
End If
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
	</script>
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
  <%
  AbrirConexao()
  %>
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">
	  <h2>Relatórios - Sáida de Produtos para Fornecedor </h2>
 
	  
		<h3 align="center"><%=TITULO%></h3>
		<div id="comboTexto">
        <%
		Set rs=objConexao.Execute(sql)
		
		If not rs.Eof Then
		%>
        
		  <table width="98%" 
		      border="1" align="center"
					cellpadding="5" 
					style="border-collapse:collapse">
            		<tr bgcolor="whitesmoke">
              <td width="9%" bgcolor="whitesmoke"><div align="center"><strong>Data Saída</strong><strong> </strong></div></td>
        <td width="20%" bgcolor="whitesmoke"><div align="center"><strong>Empresa</strong></div></td>
              <td width="29%" bgcolor="whitesmoke"><div align="center"><strong>Nome / RG - Resposável</strong> <strong>da Empresa</strong></div></td>
              <td width="8%" bgcolor="whitesmoke"><div align="center"><strong>Qtd de Itens retirados</strong></div></td>
              <td width="12%" bgcolor="whitesmoke"><div align="center"><strong>Detalhar Itens Retirados</strong></div></td>
              <td width="12%" bgcolor="whitesmoke"><div align="center"><strong>Resp. Stillos</strong></div></td>
              <td width="10%" bgcolor="whitesmoke"><div align="center"><strong>Imprimir 2ª via de recibo</strong></div></td>
            		</tr>
                    <%
					Dim QTD_TOTAL_GERAL
					
					QTD_TOTAL_GERAL = 0
					
					Dim CADASTRADO_POR_rs, DATA_SAIDA_rs, COD_FORNECEDOR_rs, NOME_FANTASIA_rs, NOME_RESPONSAVEL_SAIDA_rs, RG_RESPONSAVEL_SAIDA_rs, QTD_TOTAL_RETIRADO_rs, NOME_RESPONSAVEL_STILLOS_rs				
					
					While not rs.Eof
					
					DATA_SAIDA_rs				= ArrumaData(formatDateTime(""&rs("DATA_SAIDA")))
					CADASTRADO_POR_rs			= rs("CADASTRADO_POR")
					COD_FORNECEDOR_SAIDA_rs		= rs("COD_FORNECEDOR_SAIDA")
					COD_FORNECEDOR_rs			= rs("COD_FORNECEDOR")
					NOME_RESPONSAVEL_SAIDA_rs	= rs("NOME_RESPONSAVEL_SAIDA")
					RG_RESPONSAVEL_SAIDA_rs		= rs("RG_RESPONSAVEL_SAIDA")
					
						
						Set rs2=objConexao.Execute("SELECT NOME_FANTASIA FROM TB_FORNECEDORES WHERE COD_FORNECEDOR = "&COD_FORNECEDOR_rs)
								
								If not rs2.Eof Then
									NOME_FANTASIA_rs = rs2("NOME_FANTASIA")
								End If
						
						rs2.Close
						Set rs2=Nothing
						
						
						QTD_TOTAL_RETIRADO_rs = 0
						
						Set rs2=objConexao.Execute("SELECT COUNT(*) AS QTD_ITENS FROM TB_FORNECEDORES_SAIDA_ITENS WHERE  COD_FORNECEDOR_SAIDA = "&COD_FORNECEDOR_SAIDA_rs)
								
								If not rs2.Eof Then
									QTD_TOTAL_RETIRADO_rs = formatNumber(0&rs2("QTD_ITENS"),0)
								End If
						
						rs2.Close
						Set rs2=Nothing
						
						
						Set rs2=objConexao.Execute("SELECT NOME FROM TB_ADM_USUARIOS WHERE COD_USUARIO = '"&CADASTRADO_POR_rs&"'")
								
								If not rs2.Eof Then
									NOME_RESPONSAVEL_STILLOS_rs = rs2("NOME")
								End If
						
						rs2.Close
						Set rs2=Nothing
						
					
					
					
					
					QTD_TOTAL_GERAL = (QTD_TOTAL_GERAL + QTD_TOTAL_RETIRADO_rs)/1
					%>
                    <tr>
              <td width="9%"><div align="center"><%=DATA_SAIDA_rs%></div></td>
              <td width="20%"><%=NOME_FANTASIA_rs%></td>
              <td width="29%"><%=NOME_RESPONSAVEL_SAIDA_rs%><br /><%=RG_RESPONSAVEL_SAIDA_rs%></td>
              <td width="8%"><div align="center"><%=QTD_TOTAL_RETIRADO_rs%></div></td>
              <td width="12%"><div align="center">
              <%If int(COD_FORNECEDOR_SAIDA_qs) = int(COD_FORNECEDOR_SAIDA_rs) Then%>
              <a href="adm_relatorios_saida_fornecedores.asp?COD_FORNECEDOR=<%=COD_FORNECEDOR%>&MES=<%=MES%>&ANO=<%=ANO%>">Recolher esta saída</a>
              <%Else%>
              <a href="adm_relatorios_saida_fornecedores.asp?COD_FORNECEDOR=<%=COD_FORNECEDOR%>&MES=<%=MES%>&ANO=<%=ANO%>&COD_FORNECEDOR_SAIDA=<%=COD_FORNECEDOR_SAIDA_rs%>">Detalhar esta saída</a>
              <%End If%>
              
              </div></td>
              <td width="12%"><%=NOME_RESPONSAVEL_STILLOS_rs%></td>
              <td width="10%"><div align="center"><a href="epv_saida_fornecedores_itens_POPUP_IMPRIMIR.asp?COD_FORNECEDOR_SAIDA=<%=COD_FORNECEDOR_SAIDA_rs%>" target="_blank">Imprimir</a></div></td>
            		</tr>
                    
                    	<%If int(COD_FORNECEDOR_SAIDA_qs) = int(COD_FORNECEDOR_SAIDA_rs) Then%>
                    		
                            <%
							Set rs3=objConexao.Execute("SELECT TB_FORNECEDORES_SAIDA_ITENS.COD_FORMANDO, TB_FORNECEDORES_SAIDA_ITENS.COD_MOVIMENTACAO, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_RETORNO, TB_FICHAS_IDENTIFICACAO.COD_CONTRATO, TB_ENTIDADES.DESCRICAO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO FROM TB_FICHAS_IDENTIFICACAO INNER JOIN TB_FORNECEDORES_SAIDA_ITENS ON TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = TB_FORNECEDORES_SAIDA_ITENS.COD_FORMANDO INNER JOIN TB_ENTIDADES ON TB_FICHAS_IDENTIFICACAO.COD_CONTRATO = TB_ENTIDADES.COD_CONTRATO INNER JOIN TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ON TB_FORNECEDORES_SAIDA_ITENS.COD_MOVIMENTACAO = TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_MOVIMENTACAO WHERE TB_FORNECEDORES_SAIDA_ITENS.COD_FORNECEDOR_SAIDA = "&COD_FORNECEDOR_SAIDA_rs&" ORDER BY TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO")
							%>
							
							
		                    	<tr bgcolor="#F3F3F3">
              <td valign="top"><div align="right"><img src="images/icons/icon_filha.jpg" width="18" height="15" /></div></td>
              <td colspan="6">
              
              <table width="98%" 
		      border="1" align="center"
					cellpadding="5" 
					style="border-collapse:collapse">
            		<tr bgcolor="whitesmoke">
                  <th width="12%" scope="col">Cód Formando</th>
                  <th width="29%" bgcolor="whitesmoke" scope="col">Nome do Formando</th>
                  <th width="43%" bgcolor="whitesmoke" scope="col">Contrato - Faculdade</th>
                  <th width="16%" bgcolor="whitesmoke" scope="col">Data Retornou para Estoque</th>
                </tr>
                <%
							While not rs3.Eof
							%>
                			<tr bgcolor="whitesmoke">
                  <td><%=rs3("COD_FORMANDO")%></td>
                  <td bgcolor="whitesmoke"><%=rs3("NOME_FORMANDO")%></td>
                  <td bgcolor="whitesmoke"><%=rs3("DESCRICAO")%></td>
                  <td bgcolor="whitesmoke"><%=ArrumaData(formatDateTime(rs3("DATA_RETORNO")))%></td>
                </tr>
                			<%
							rs3.MoveNext()
							Wend%>	
              </table></td>
              </tr>
                    								
							
							rs3.Close
							Set rs3=Nothing
							%>
                            
                            
                        <%End If%>
                    
            		<%
					rs.MoveNext()
					Wend
					
					%>   
                    
                    <tr bgcolor="#CCCCCC">
              <td colspan="6"><h3>Total Itens Retirados</h3></td>
              <td width="10%"><div align="center" class="textoMaior"><strong>
              <%=formatNumber(QTD_TOTAL_GERAL,0)%>
              </strong></div></td>
                    </tr>
          </table>
          
        <%Else%>
        	<p align="center"><br />
<br />
<em>Não foram encontrados registros para os filtros selecionados!</em><br />
<br />
<br />
<br />
</p>
        
        <%End If
		
		rs.Close
		Set rs=Nothing
		%>  
		</div>
		
		
		


		<br />
		<br />
<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="98%" border="0" align="center" cellpadding="5" cellspacing="0">
          <tr>
            <td><p><strong>Produtos habilitados para saída de fornecedores<br />
            </strong></p>
        
			<%
			Set rs3=objConexao.Execute("SELECT NOME_PRODUTO FROM TB_PRODUTOS WHERE ATIVO = 1 AND HABILITA_SAIDA_FORNECEDOR = 1 ORDER BY NOME_PRODUTO")
			
				While not rs3.Eof
			%>
              <strong><%=rs3("NOME_PRODUTO")%></strong><BR />

            <%
				rs3.MoveNext()
				Wend
				
			rs3.Close
			Set rs3=Nothing
			%> 
              </td>
           
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
<%FecharConexao()%>
  </td>
</tr>
</table>
</body>
</html>






