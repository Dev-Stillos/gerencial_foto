<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
sql = "SELECT * FROM TB_BOLETIM_INFORMATIVO ORDER BY DATA_CADASTRO DESC"

If Request.QueryString("COD_CONTRATO") <> Empty Then
	sql = "SELECT TB_BOLETIM_INFORMATIVO.* FROM TB_BOLETIM_INFORMATIVO_CONTRATOS INNER JOIN TB_BOLETIM_INFORMATIVO ON TB_BOLETIM_INFORMATIVO_CONTRATOS.COD_BOLETIM_INFORMATIVO = TB_BOLETIM_INFORMATIVO.COD_BOLETIM_INFORMATIVO WHERE TB_BOLETIM_INFORMATIVO_CONTRATOS.COD_CONTRATO = '"&Request.QueryString("COD_CONTRATO")&"' GROUP BY TB_BOLETIM_INFORMATIVO.COD_BOLETIM_INFORMATIVO "
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

	  <h2>DIP - Produtos e Valores - Boletins Informativos</h2>
	  <!-- #include file="includes/adm_produtos_sub_menu.asp" --> 
      <form style="margin:0px" name="FILTRO_BOLETINS" action="dip_produtos_boletins_informativos.asp">
		<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC; margin:3px" width="99%" border="0" align="center" cellpadding="5" cellspacing="0">
          <tr>
            <td><table width="670" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td width="14%"><div align="left"><strong>Filtrar por<BR />
                C&oacute;d. contrato:</strong></div></td>
                <td width="65%">
                    <div align="left">
                      <select name="COD_CONTRATO" class="inputTamanho3">
                      	<option value="">Todos</option>
                      	<%
						AbrirConexao()
						Set rs = objConexao.Execute("SELECT COD_CONTRATO, DESCRICAO FROM TB_ENTIDADES ORDER BY COD_CONTRATO")
							While not rs.Eof
						%>
                      		<option <%If Cstr(""&Request.QueryString("COD_CONTRATO")) = Cstr(""&rs("COD_CONTRATO")) Then%> selected="selected"<%End If%> value="<%=rs("COD_CONTRATO")%>"><%=rs("DESCRICAO")%></option>
                        <%
							rs.MoveNext()
							Wend
						
						rs.Close
						Set rs=Nothing
						FecharConexao()
						%>
                      </select>
                      
                      
                        </div></td>
                <td width="21%"><div align="left">
                    <input class="btnFiltrarAmarelo" onclick="Abre('#')" type="submit" name="filtro2" value=" " />
                </div></td>
              </tr>
            </table></td>
            </tr>
        </table>
		</form>
      
	  <%
	   Response.Write("<br />")
	Dim sucesso, codmsg
	codmsg	= CStr(""&Request.QueryString("codmsg"))
	sucesso = CStr(""&Request.QueryString("sucesso"))
		
	If sucesso = "" and codmsg <> "" Then%>
    	<div align="center" class="msgGrandeSucesso">
        	<%
			If codmsg = "1" Then
				Response.Write("Boletim Informativo incluído com sucesso!")
			Elseif codmsg = "2" Then
				Response.Write("Boletim Informativo alterado com sucesso!")
			Elseif codmsg = "3" Then	
				Response.Write("Boletim Informativo excluído com sucesso!")
			End If	
			%>
            
        
        </div>
    <%ElseIf sucesso <> "" and codmsg <> "" Then%>
        <div align="center" class="msgGrandeErro">
			<%
			If codmsg = "4" Then
			Elseif codmsg = "5" Then		
			End If	
			%>
</div>
	<%End If%>
    
    
			
            <%
			AbrirConexao()
			Dim rsContratos, qtdContratos, CONTRATOS_ULT, qtdProdutos
			
			
			
			Set rs=server.CreateObject("ADODB.recordset")
			rs.CursorLocation=3  'Cursor no cliente
			rs.PageSize=50
			rs.Open sql, objConexao
		  
		  	Dim PAGINA_ATUAL
			PAGINA_ATUAL = formatNumber(0&Request.QueryString("pagina"),0)
			
			If PAGINA_ATUAL > 0 Then
					rs.AbsolutePage=PAGINA_ATUAL
			Else
				PAGINA_ATUAL = 1
				rs.AbsolutePage=1
			End If
			
			If not rs.Eof Then%>
               <table width="95%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="11%"><div align="center"><strong>C&oacute;d. Boletim Informativo</strong></div></td>
            <td width="12%"><div align="center"><strong>Contrato(s)</strong></div></td>
            <td width="10%"><div align="center"><strong>Qtd. Produtos</strong></div></td>
            <td width="12%"><div align="center"><strong>Data do Cadastro</strong></div></td>
            <td width="18%"><div align="center"><strong>Data da Pr&oacute;xima<br />corre&ccedil;&atilde;o monet&aacute;ria</strong></div></td>
            <td width="10%"><div align="center"><strong>Status</strong></div></td>
            <td width="9%"><div align="center"><strong>Cadastrado por</strong></div></td>
            <td width="10%"><div align="center"><strong>Imprimir para assinaturas</strong></div></td>
            <td width="8%"><div align="center"><strong>Alterar / Excluir</strong></div></td>
</tr>
		  <%
		  	Dim x
		 	For x=1 To rs.PageSize
		    If Not rs.Eof Then
		  %>
           <tr>
            <td><div align="center"><%=rs("COD_BOLETIM_INFORMATIVO")%></div></td>
            <td><div align="left">
            <%
				
				qtdContratos = 0
				CONTRATOS_ULT = ""
							
				Set rsContratos = objConexao.Execute("SELECT DISTINCT(COD_CONTRATO) AS COD_CONTRATO, DESCRICAO FROM VW_TB_BOLETIM_INFORMATIVO_CONTRATOS WHERE COD_BOLETIM_INFORMATIVO="&rs("COD_BOLETIM_INFORMATIVO")&" ORDER BY COD_CONTRATO ASC",qtdContratos)
							
							If not rsContratos.Eof then
															
								
								If qtdContratos > 1 then
							
									CONTRATOS_ULT = ""
									For i = 0 to qtdContratos-1
										
										If i = qtdContratos-1 then
										
											CONTRATOS_ULT = CONTRATOS_ULT&rsContratos("DESCRICAO")
										
										else
										
											CONTRATOS_ULT = CONTRATOS_ULT&rsContratos("DESCRICAO")&"<br />"
										
										End If
									
									rsContratos.MoveNext()
									Next
								
									
								else
									CONTRATOS_ULT = rsContratos("DESCRICAO")
									
								
								End If
			
								CONTRATOS_ULT = Replace(CONTRATOS_ULT,Request.QueryString("COD_CONTRATO"),"<span style='background-color:#FFFF99; padding:1px; font-weight:bold;'>"&Request.QueryString("COD_CONTRATO")&"</span>")
								
								Response.Write(CONTRATOS_ULT)
								
							else
							
							Response.Write("-")
							
							End If
							
							rsContratos.Close()
							Set rsContratos = Nothing
			%>
            &nbsp;</div></td>
            <td>
           	  <div align="center">
           	    <%
			qtdProdutos = 0
			
			Set rs2=objConexao.Execute("SELECT ID FROM VW_TB_BOLETIM_INFORMATIVO_PRODUTOS WHERE COD_BOLETIM_INFORMATIVO = '"&rs("COD_BOLETIM_INFORMATIVO")&"'",qtdProdutos)
				
				Response.Write(qtdProdutos)
				
			rs2.Close
			Set rs2=Nothing
			%>
           	    &nbsp;</div></td>
            <td><div align="center"><%=ArrumaData(rs("DATA_CADASTRO"))%></div></td>
            <td><div align="center"><%=ArrumaData(rs("DATA_PROXIMA_ATUALIZACAO_MONETARIA"))%></div></td>
            <td><div align="center">
            <%
			If int(rs("status")) = 1 Then
				Response.Write("Em utilização")
			Else%>
				<div style="color:#CC0000;"><strong>Preço desatualizado é necessário atualizar os Preços</strong>.</div>
                <%
			End If
			%>
            </div></td>
            <td>
            
            
              <div align="left">
                <%	
			Set rs2=objConexao.Execute("Select nome From TB_ADM_USUARIOS where COD_USUARIO = '"&rs("CADASTRADO_POR")&"'")
				
				If not rs2.Eof Then
					Response.Write(rs2(0))
				End If
				
			rs2.Close
			Set rs2=Nothing
			%>
              </div></td>
            <td><div align="center"><a href="dip_produtos_boletins_informativos_impressao.asp?cod_boletim_informativo=<%=rs("COD_BOLETIM_INFORMATIVO")%>" target="_blank">Imprimir</a></div></td>
            <td>
              <div align="center"><a href="dip_produtos_boletins_informativos_formulario.asp?cod_boletim_informativo=<%=rs("COD_BOLETIM_INFORMATIVO")%>&comando=alterar">
                  <img alt="Alterar" src="images/icons/icon_alterar.png" border="0" /></a>&nbsp;&nbsp;
                    <a href="dip_produtos_boletins_informativos_formulario.asp?cod_boletim_informativo=<%=rs("COD_BOLETIM_INFORMATIVO")%>&comando=excluir">
                          <img alt="Excluir" src="images/icons/icon_excluir.png" border="0" /></a>                </div></td>
           </tr>
		  <%
				rs.MoveNext()
				End If 
				Next
		  %>
        </table>
        		
                		        <%If rs.PageCount > 1 Then %>
					<br />
					<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td>
						<div class="menu_paginacao">
					Ir para a página:
					
					<%
					Dim aae
					aae = 0
					For x=1 To rs.PageCount 
					aae = formatNumber(0&x,0)
					
						If int(PAGINA_ATUAL) <> int(aae) Then
						'Response.Write(PAGINA_ATUAL&":"&x&"<br />")
						%>
							<a href="?pagina=<%=x%>&f=<%=FILTRO_F%>"><%=x%></a>
						<%Else%>
							<%=x%>
						<%End If%>
						
					<%Next%>
					</div>
					</td>
                      </tr>
                    </table>
				<%End if%>
        
        
		      <%else%>
              
              		<%If Request.QueryString("COD_CONTRATO") <> Empty Then%>
                      <div align="center">
                    <p class="textoMedio">
                      <em>Não foi encontrado boletim informativo para o contrato <strong>"<%=Request.QueryString("COD_CONTRATO")%>"</strong></em>	</p>
                    <p class="textoMedio">&nbsp;</p>
                    </div>
                     <%eLSE%>
                     
                     <div align="center">
                    <p class="textoMedio">
                      <em>Não há boletins informativos cadastrados! </em>	</p>
                    <p class="textoMedio">Para incluir clique no link &quot;Incluir Boletim Informativo&quot; acima.</p>
                    <p>&nbsp;</p>
                    <p>	   
                     </p>
                     </div> 
                     <%eND IF%>
                    
		      <%end if
			  
			  
			  	  rs.Close()
				  Set rs = Nothing
				  FecharConexao()
			  %>
		     
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