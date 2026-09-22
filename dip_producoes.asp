<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>


<title>Produ&ccedil;&otilde;es</title>
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
	<!--AJUDA F1 E CTRL PARA BUSCA RPIDA-->
		
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
 


	<!--AJUDA F1 E CTRL PARA BUSCA RPIDA-->
	</script>
		<!-- CALENDRIO SCRIPTS
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
<style type="text/css">
<!--
.menu_paginacao{
	float: left;
	width: 300px;
	position: relative;
	color: #2f2f44
}
.menu_paginacao div{
	position: relative;
	float: left;
	height: 5px;
	width: 10px;
	border: 1px solid #666666;
	padding: 2px;
}
.menu_paginacao a{
	font-size:12px;
	text-decoration:none;
	color: #2f2f44;
}

.menu_paginacao a:link{
	font-size:12px;
	text-decoration:none;
}

.menu_paginacao a:visited{
	font-size:12px;
	text-decoration:none;
}

.menu_paginacao a:active{
	font-size:12px;
	text-decoration:none;
}

.menu_paginacao a:hover{
	font-size:12px;
	text-decoration:none;
	font-weight: bold;
}
-->
</style>


<%
Dim filtro, busca_nome, busca_codigo, busca_contrato
filtro = ""
busca_nome = Trim(Request("busca_nome"))
busca_codigo = Trim(Request("busca_codigo"))
busca_contrato = Trim(Request("busca_contrato"))
%>

</head>

<body onKeyDown="javascript:Verificar()">

<table cellpadding="9px" cellspacing="0" width="100%">
<tr>
<td>
<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td>
		<!-- #include file="includes/topo.asp" -->
<table width="100%" style="border-bottom:#999999 solid dotted 1px" background="#ff007f" border="0" cellspacing="0" cellpadding="0">
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

	  
		
<div style="margin-bottom: 20px; display: flex; justify-content: center;">
  <div style="position: relative; display: inline-block;">
    <button type="button" onclick="toggleDropdown()" 
       style="background-color:#ff007f; color:white; padding: 10px 20px; border:none; border-radius: 6px; cursor:pointer; display:flex; align-items:center; gap:10px;">
      <strong>DIP - Produ&ccedil;&otilde;es</strong> | Sub-Menu <span style="font-size:12px;">&#9660;</span>
    </button>
    <div id="submenuDropdown" 
         style="visibility: hidden; opacity: 0; transition: opacity 0.3s ease; position:absolute; background-color:#fff; min-width:250px; box-shadow:0px 8px 16px rgba(0,0,0,0.2); border:1px solid #ccc; border-radius:6px; z-index:1;">
      <a href="adm_ROTINA_gravacao_contratos.asp" style="color:#ff007f; padding:10px 15px; display:block; text-decoration:none;">Incluir Produ&ccedil;&atilde;o</a>
      <a href="?f=1" style="color:#ff007f; padding:10px 15px; display:block; text-decoration:none;">Somente Finalizadas</a>
      <a href="?f=2" style="color:#ff007f; padding:10px 15px; display:block; text-decoration:none;">Somente Não Finalizadas</a>
	  <a href="?f=3" style="color:#ff007f; padding:10px 15px; display:block; text-decoration:none;">Somente Álbum Digital 📸</a> <!--IMPLEMNETADO 16/09/2025-->
      <a href="?" style="color:#ff007f; padding:10px 15px; display:block; text-decoration:none;">Limpar Filtro</a>
      <a href="dip_lista_antigos.asp" style="color:#ff007f; padding:10px 15px; display:block; text-decoration:none;">Produ&ccedil;&otilde;es anteriores a 2007</a>
    </div>
  </div>
</div>

<script>
function toggleDropdown() {
  var dropdown = document.getElementById("submenuDropdown");
  if (dropdown.style.visibility === "visible") {
    dropdown.style.visibility = "hidden";
    dropdown.style.opacity = "0";
  } else {
    dropdown.style.visibility = "visible";
    dropdown.style.opacity = "1";
  }
}
</script>






</script>
<form method="get" action="dip_producoes.asp" style="display:flex; justify-content:center; align-items:center; gap:10px; margin-bottom: 15px; margin-top: 10px;">
    <input type="text" name="busca_nome" placeholder="Nome da produção" value="<%=Request("busca_nome")%>" 
        style="padding: 8px; border: 1px solid #ccc; border-radius: 6px; width: 250px;">
    <input type="text" name="busca_codigo" placeholder="Código da produção" value="<%=Request("busca_codigo")%>" 
        style="padding: 8px; border: 1px solid #ccc; border-radius: 6px; width: 180px;">
	<input type="text" name="busca_contrato" placeholder="Contrato" value="<%=Request("busca_contrato")%>"
    	style="padding: 8px; border: 1px solid #ccc; border-radius: 6px; width: 180px;">
    <input type="submit" value="Buscar" 
        style="background-color:#ff007f; color:white; border:none; padding:8px 16px; border-radius:6px; cursor:pointer;">
    <a href="dip_producoes.asp" style="padding: 8px 12px; background: #ccc; border-radius: 6px; text-decoration: none; color: black;">Limpar</a>
</form>
<div class="msg" align="center"><%=Request.QueryString("msg")%></div>

	   <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	  <%
		  AbrirConexao()
		  Dim rs, rs2, rs3, sql
		  Dim QTD_ALBUNS, FICHAS_VENDA, FILTRO_F
		  
		  FILTRO_F = formatNumber(0&Request.QueryString("f"),0)
		  
		  If FILTRO_F = 1 Then
    sql = "SELECT p.* FROM TB_PRODUCAO p WHERE p.PRODUCAO_FINALIZADA = 1"
ElseIf FILTRO_F = 2 Then
    sql = "SELECT p.* FROM TB_PRODUCAO p WHERE p.PRODUCAO_FINALIZADA = 2"
ElseIf FILTRO_F = 3 Then
    sql = "SELECT p.* " & _
          "FROM TB_PRODUCAO p " & _
          "WHERE EXISTS (SELECT 1 FROM TB_FICHAS_IDENTIFICACAO f " & _
          "              WHERE f.COD_PRODUCAO = p.COD_PRODUCAO " & _
          "                AND f.ALBUM_DIGITAL = 1)"
Else
    sql = "SELECT p.* FROM TB_PRODUCAO p WHERE 1=1"
End If


' Aplica o filtro por nome e código, se fornecidos
If busca_nome <> "" Then
    sql = sql & " AND p.DESCRICAO LIKE '%" & Replace(busca_nome, "'", "''") & "%'"
End If

If busca_codigo <> "" Then
    sql = sql & " AND p.COD_PRODUCAO = '" & Replace(busca_codigo, "'", "''") & "'"
End If

If busca_contrato <> "" Then
    sql = sql & " AND EXISTS (" & _
                " SELECT 1 " & _
                " FROM TB_FICHAS_IDENTIFICACAO f " & _
                " WHERE f.COD_PRODUCAO = p.COD_PRODUCAO " & _
                " AND f.COD_CONTRATO LIKE '%" & Replace(busca_contrato, "'", "''") & "%'" & _
                ")"
End If







sql = sql & " ORDER BY COD_PRODUCAO DESC"	  
		  	Set rs=server.CreateObject("ADODB.recordset")
			rs.CursorLocation=3  'Cursor no cliente
			rs.PageSize=10
			rs.Open sql, objConexao
		  
		  	Dim PAGINA_ATUAL
			PAGINA_ATUAL = formatNumber(0&Request.QueryString("pagina"),0)
			
						If Not (rs.EOF Or rs.BOF) Then
				If PAGINA_ATUAL > 0 Then
					rs.AbsolutePage = PAGINA_ATUAL
				Else
					PAGINA_ATUAL = 1
					rs.AbsolutePage = 1
				End If

				' Agora é seguro acessar os dados
			%>
				<div id="comboTexto">
				<table width="95%" border="0" align="center">
				<tr>
				</tr>
				</table>
			<%
			Else
				Response.Write "<div align='center' style='color:red;'>Nenhuma produ&ccedil;&atilde;o encontrada com os crit&eacute;rios informados.</div>"
			End If
	
		  
		  
		  If not rs.Eof then
		  %>
		  <h3 align="center">Produ&ccedil;&otilde;es </h3>
		  <div id="comboTexto">
		  <table width="95%" border="0" align="center">
  <tr>
    <td width="47%"> 
P&aacute;gina <strong><%=rs.AbsolutePage%></strong> de <%=rs.PageCount%><br /></td>
    <td width="53%"><div align="right">Total <strong><%=rs.RecordCount%></strong> produ&ccedil;&otilde;es</div></td>
  </tr>
</table>
		  
          <table width="95%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr style="background-color:#333353;; color:white !important;">
            <td width="6%"><div align="center"><strong>C&oacute;d. Produ&ccedil;&atilde;o</strong></div></td>
            <td width="24%"><div align="center"><strong>Produ&ccedil;&atilde;o</strong></div></td>
            <td width="7%"><div align="center"><strong>Contrato(s)</strong></div></td>
            <td width="8%"><div align="center"><strong>Data Finalizada</strong></div></td>
            <td width="5%"><div align="center"><strong>Qtd. Lotes</strong></div></td>
            <td width="7%"><div align="center"><strong>Produto(s) Agregado(s) </strong></div></td>
            <td width="9%"><div align="center"><strong>Qtd. &Aacute;lbuns / M&iacute;n. M&aacute;x N&uacute;m. </strong></div></td>
            <td width="7%"><div align="center"><strong>Total Fotos </strong></div></td>
            <td width="7%"><div align="center"><strong>Formandos</strong></div></td>
            <td width="7%"><div align="center"><strong>Produ&ccedil;&atilde;o Finalizada </strong></div></td>
			<td width="8%"><div align="center"><strong>Álbum Digital 📸</strong></div></td>

</tr>
		  		 <%
	 
				 
				  
		 Dim x
		 For x=1 To rs.PageSize
		    If Not rs.Eof Then
			
				Dim DESCRICAO, PRODUTOS_AGREDADOS
				DESCRICAO = rs("DESCRICAO")
				PRODUTOS_AGREDADOS = rs("PRODUTOS_AGREDADOS")
				FICHAS_COMPLETAS = 0
		 		%>
				<tr
						  <%
						  Dim cod_prod_Query, cod_prod_rs
						  cod_prod_Query	= formatNumber(0&Request.QueryString("cod_prod"),0)
						  cod_prod_rs		= formatNumber(0&rs("COD_PRODUCAO"),0)
						
						  if cod_prod_Query = cod_prod_rs then%> bgcolor="#FFFF99"<%end if%>>
							<td><div align="center"><strong><%=cod_prod_rs%></strong></div></td>
							<td><%=Ucase(DESCRICAO)%></td>
							<td>
							<%
							Dim rsContratos, qtdContratos, CONTRATOS_ULT
							
							Set rsContratos = objConexao.Execute("SELECT DISTINCT(COD_CONTRATO) AS COD_CONTRATO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO="&rs("COD_PRODUCAO")&" ORDER BY COD_CONTRATO ASC",qtdContratos)
							
							If not rsContratos.Eof then
							
								If qtdContratos > 1 then
							
									CONTRATOS_ULT = ""
									For i = 0 to qtdContratos-1
										
										If i = qtdContratos-1 then
										
											CONTRATOS_ULT = CONTRATOS_ULT&rsContratos("COD_CONTRATO")
										
										else
										
											CONTRATOS_ULT = CONTRATOS_ULT&rsContratos("COD_CONTRATO")&", "
										
										End If
									
									rsContratos.MoveNext()
									Next
									
									Response.Write(CONTRATOS_ULT)
									
								else
									
									Response.Write(rsContratos("COD_CONTRATO"))
								
								End If
								
									
									
							else
							
							Response.Write("-")
							
							End If
							
							rsContratos.Close()
							Set rsContratos = Nothing							
							%>			</td>
							<td><div align="center">
                            <%=ArrumaData(rs("DATA_FINALIZADA"))%>
                            </div></td>
							<td><div align="center">
							  <%
							Set rs4=objConexao.Execute("SELECT count(*) as qtd_lotes FROM TB_PRODUCAO_LOTES WHERE COD_PRODUCAO = "&rs("COD_PRODUCAO"))
								If not rs4.Eof Then
									
									Response.Write(rs4("qtd_lotes"))
								
								End If
								
								
							rs4.Close
							Set rs4=Nothing
							%>
							  </div></td>
							<td><div align="center">
							  <%If PRODUTOS_AGREDADOS = 1 Then%>
							  Sim
							  <%Else%>
							  N&atilde;o
							  <%End If%>
							</div></td>
							<td>
							
								<div align="center">
								  <%
									Set rs2 = objConexao.Execute("SELECT COUNT(*) AS QTD_ALBUNS FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = "&rs("COD_PRODUCAO"))		
										
										QTD_ALBUNS = 0&rs2("QTD_ALBUNS")
										Response.Write("Qtd.: "&formatNumber(QTD_ALBUNS,0))
									
									rs2.Close()
									Set rs2 = Nothing
									
									Set rs2 = objConexao.Execute("SELECT MIN(CONVERT(ALBUM,SIGNED)) AS MENOR, MAX(CONVERT(ALBUM,SIGNED)) AS MAIOR FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = "&rs("COD_PRODUCAO")&" GROUP BY COD_PRODUCAO")
										
										If not rs2.Eof Then
										Dim MENOR_ALBUM, MAIOR_ALBUM
											MENOR_ALBUM = rs2("MENOR")
											MAIOR_ALBUM = rs2("MAIOR")
										End If
										
									
									rs2.Close()
									Set rs2 = Nothing
									
									
								  %>
							   &nbsp; <br />
							   Menor: <%=MENOR_ALBUM%> Maior: <%=MAIOR_ALBUM%> </div></td>
							<td>
							  <div align="center">
								<%
							Set rs87 = objConexao.Execute("SELECT SUM(INICIAL) AS TOTAL_INICIAL FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = "&rs("COD_PRODUCAO")&" GROUP BY COD_PRODUCAO ")
									If not rs87.Eof Then
										Response.Write(formatNumber(0&rs87("TOTAL_INICIAL"),0))
									End If
									
							rs87.Close
							Set rs87 = Nothing
							%>
								</div></td>
							<td><div align="center"><a href="dip_producoes_fichas.asp?cod_prod=<%=rs("COD_PRODUCAO")%>&prod_agre=<%=PRODUTOS_AGREDADOS%>"><img src="images/layout/icon_listas.png" width="16" height="16" border="0" /></a></div></td>
							<td><div align="center">
							<%If rs("PRODUCAO_FINALIZADA") = 1 then  %>
							<img src="images/icons/icon_balao_green.gif" width="17" height="16" />
							<%else%>
							<img src="images/icons/icon_balao_red.gif" width="17" height="16" />
							<%End If%>
							</div></td>

							<!--Álbum Digital-->

							<td><div align="center">
								<%
								Dim rsAlbDig, sqlAlbDig, codProd
								codProd = rs("COD_PRODUCAO")

								' Se não for numérico, transforma em 0 para não quebrar
								If IsNull(codProd) Or codProd = "" Then
									codProd = 0
								End If

								sqlAlbDig = "SELECT COUNT(*) AS QTD " & _
											"FROM TB_FICHAS_IDENTIFICACAO " & _
											"WHERE COD_PRODUCAO = " & CLng(codProd) & " " & _
											"AND ALBUM_DIGITAL = 1"

								Set rsAlbDig = objConexao.Execute(sqlAlbDig)

								If Not rsAlbDig.Eof Then
								Dim qtdAlbDig
								qtdAlbDig = rsAlbDig("QTD")

								' Se vier Null, vira 0
								If IsNull(qtdAlbDig) Or qtdAlbDig = "" Then
									qtdAlbDig = 0
								End If

								If CLng(qtdAlbDig) > 0 Then
									Response.Write("SIM 📔")
								Else
									Response.Write("NÃO")
								End If
							End If


								rsAlbDig.Close
								Set rsAlbDig = Nothing
								%>
							</div></td>


						  </tr>
		        <%
				rs.MoveNext
				End If 
				Next	
      
		  %>
                </table>

<div style="text-align: center; margin-top: 20px;">
    <% If rs.PageCount > 1 Then %>
        <div style="display: inline-flex; align-items: center; gap: 10px; flex-wrap: wrap; background: #ff007f; padding: 10px 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
            <label for="paginaSelect" style="font-weight: bold; color: White;">Ir para a P&aacute;gina:</label>

            <select id="paginaSelect" 
                    onchange="location.href='?pagina=' + this.value + '&f=<%=FILTRO_F%>'"
                    style="padding: 8px 12px; border-radius: 6px; border: 1px solid #ccc; background-color: #fff; cursor: pointer; font-weight: bold; color: #ff007f;">

                <% 
                Dim aae
                For x = 1 To rs.PageCount
                    aae = formatNumber(0 & x, 0)
                %>
                    <option value="<%=x%>" <% If int(PAGINA_ATUAL) = int(aae) Then %>selected<% End If %>>
                        P&aacute;gina <%=x%>
                    </option>
                <% Next %>
            </select>
        </div>
    <% End If %>
</div>


        <%
		Else
		
		End If      
		
		
 		rs.Close()
		Set rs = Nothing
		FecharConexao()
		%>
		            
		      </div>
		  </div>
		
		  <p align="center">
		  <button class="btnTopo" onclick="window.scrollTo({top: 0, behavior: 'smooth'});"></button>
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