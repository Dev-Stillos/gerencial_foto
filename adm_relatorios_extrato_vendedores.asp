<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Validação especial para verificar se possui sub-validação específica
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '15' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se não encontrar então redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
		
	rs.Close
	Set rs=Nothing
	
FecharConexao()



Dim COD_VENDEDOR, DATA_INI, DATA_FIM, TIPO, ENCADERNACOES_EXTRAS, NOME_VENDEDOR
Dim FILTRO_ALBUNS, FILTRO_FICHAS, FILTRO_AGREGADOS, CHEQ_DEVOLVIDOS, FILTRO_VENDIDOS
Dim PRIMEIRO_ACESSO
COD_VENDEDOR	= Request.QueryString("COD_VENDEDOR")
DATA_INI		= ArrumaData(Request.QueryString("DATA_INI"))
DATA_FIM		= ArrumaData(Request.QueryString("DATA_FIM"))

PRIMEIRO_ACESSO = formatNumber(0&Request.QueryString("pa"),0)

FILTRO_ALBUNS		= formatNumber(0&Request.QueryString("FILTRO_ALBUNS"),0)
FILTRO_FICHAS		= formatNumber(0&Request.QueryString("FILTRO_FICHAS"),0)
FILTRO_AGREGADOS	= formatNumber(0&Request.QueryString("FILTRO_AGREGADOS"),0)
FILTRO_VENDIDOS		= formatNumber(0&Request.QueryString("FILTRO_VENDIDOS"),0)
CHEQ_DEVOLVIDOS		= formatNumber(0&Request.QueryString("CHEQ_DEVOLVIDOS"),0)

'Deixa todos os filtros habilitados
If PRIMEIRO_ACESSO = 1 Then
	FILTRO_ALBUNS		= 1
	FILTRO_FICHAS		= 0
	FILTRO_AGREGADOS	= 1
	CHEQ_DEVOLVIDOS		= 4
End If

AbrirConexao()

	Set rsV = objConexao.Execute("SELECT NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR="&COD_VENDEDOR)
				
				If not rsV.Eof Then
					NOME_VENDEDOR = rsV("NOME_VENDEDOR")
				End If
				
	rsV.Close
	Set rsV = Nothing
FecharConexao()

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
	function popUp(URL){
	win=window.open(URL,"","width=550,height=350,scrollbars=yes")
	var text = "Se a janela não abrir desbloquei os Pop-ups!";
	if(win == null) { alert(text); return; }
	win.moveTo(350,150);
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
  BTN_IMPRIMIR = 1
  BTN_FONTE = 1
  
   BTN_IMPRIMIR_LINK = "href=""javascript:popUp('adm_relatorios_extrato_vendedores_IMPRIMIR.asp?COD_VENDEDOR="&COD_VENDEDOR&"&DATA_INI="&DATA_INI&"&DATA_FIM="&DATA_FIM&"&FILTRO_ALBUNS="&FILTRO_ALBUNS&"&FILTRO_AGREGADOS="&FILTRO_AGREGADOS&"&CHEQ_DEVOLVIDOS="&CHEQ_DEVOLVIDOS&"')"""
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">
	
	  <h2>Relatórios - Extrato de Movimentações </h2>
	  <div align="center" class="msg"><%=Request.QueryString("msg")%></div>
	  <p><span class="textoMedio"><strong>Representante Comercial: <%=NOME_VENDEDOR%><br />
</strong></span>
<strong>
Filtro:  <%=DATA_INI%> à <%=DATA_FIM%><br />
</strong>
<br />

<form style="margin:0px">
<input name="COD_VENDEDOR" value="<%=COD_VENDEDOR%>"  type="hidden" />
<input name="DATA_INI"  value="<%=DATA_INI%>" type="hidden" />
<input name="DATA_FIM"  value="<%=DATA_FIM%>" type="hidden" />
<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="100%" border="0" align="center" cellpadding="5" cellspacing="0">
		  <tr>
           
            <td width="96%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td width="87%"><table align="left" width="75" border="0" cellspacing="0" cellpadding="3">
                  <tr>
                    <td width="28"><div align="center">
                        <input <%If FILTRO_ALBUNS = 1 Then%> checked="checked" <%End If%> class="inputRadio" type="checkbox" name="FILTRO_ALBUNS" value="1" />
                    </div></td>
                    <td width="272"><div align="left">Álbuns</div></td>
                  </tr>
                </table>
                <table align="left" width="80" border="0" cellspacing="0" cellpadding="3">
                  <tr>
                    <td width="28"><div align="center">
                        <input <%If FILTRO_FICHAS = 1 Then%> checked="checked" <%End If%> class="inputRadio" type="checkbox" name="FILTRO_FICHAS" value="1" />
                    </div></td>
                    <td width="272"><div align="left">Fichas</div></td>
                  </tr>
                </table>
                <table align="left" width="130" border="0" cellspacing="0" cellpadding="3">
                  <tr>
                    <td width="28"><div align="center">
                        <input <%If FILTRO_AGREGADOS = 1 Then%> checked="checked" <%End If%> class="inputRadio" type="checkbox" name="FILTRO_AGREGADOS" value="1" />
                    </div></td>
                    <td width="272"><div align="left">Prod. Agregados </div></td>
                  </tr>
                </table>
                </td>
                <td width="13%" rowspan="2"><div align="right">
                  <input class="btnFiltrarAmarelo" type="submit" name="filtro1" value=" " />
                </div></td>
              </tr>
              <tr>
                <td><table align="left" width="200" border="0" cellspacing="0" cellpadding="3">
                  <tr>
                    <td width="20"><div align="center">
                      <input <%If CHEQ_DEVOLVIDOS = 3 Then%> checked="checked" <%End If%> class="inputRadio" name="CHEQ_DEVOLVIDOS" type="radio" value="3" />
                    </div></td>
                    <td width="168"><div align="left">Devolvidos e Não Devolvidos </div></td>
                  </tr>
                </table>
                <table align="left" width="100" border="0" cellspacing="0" cellpadding="3">
                  <tr>
                    <td width="28"><div align="center">
                      <input <%If CHEQ_DEVOLVIDOS = 1 Then%> checked="checked" <%End If%> class="inputRadio" name="CHEQ_DEVOLVIDOS" type="radio" value="1" />
                    </div></td>
                    <td width="272"><div align="left">Devolvidos</div></td>
                  </tr>
                </table>
                  <table align="left" width="120" border="0" cellspacing="0" cellpadding="3">
                  <tr>
                    <td width="28"><div align="center"><input <%If CHEQ_DEVOLVIDOS = 2 Then%> checked="checked" <%End If%> class="inputRadio" name="CHEQ_DEVOLVIDOS" type="radio" value="2" /></div></td>
                    <td width="272"><div align="left"> Não Devolvidos</div></td>
                  </tr>
                </table>
                  <table align="left" width="200" border="0" cellspacing="0" cellpadding="3">
                  <tr>
                    <td width="20"><div align="center"><input <%If CHEQ_DEVOLVIDOS = 4 Then%> checked="checked" <%End If%> class="inputRadio" name="CHEQ_DEVOLVIDOS" type="radio" value="4" /></div></td>
                    <td width="168"><div align="left">Atrasados </div></td>
                  </tr>
                </table></td>
                </tr>
            </table></td>
          </tr>
        </table>
</form>

      <p>
	  <%
		AbrirConexao()
		DATA_INI = Year(DATA_INI)&"-"&Month(DATA_INI)&"-"&Day(DATA_INI)&" 00:00:00"
		DATA_FIM = Year(DATA_FIM)&"-"&Month(DATA_FIM)&"-"&Day(DATA_FIM)&" 23:59:59"
		
		Dim rs, sql
		sql = 	"SELECT * FROM VW_MOVIMENTACOES_COMPLETA WHERE COD_VENDEDOR = "&COD_VENDEDOR&" AND "&_
				" (DATA_SAIDA BETWEEN '"&DATA_INI&"' AND '"&DATA_FIM&"'"&_
				" OR DATA_RETORNO BETWEEN '"&DATA_INI&"' AND '"&DATA_FIM&"' OR DATA_RETORNO IS NULL)"
		
		'Response.Write(sql)
		If FILTRO_ALBUNS = 0 Then
			sql = sql&" AND TIPO <> 0"
		End If
		
		If FILTRO_FICHAS = 0 Then
			sql = sql&" AND TIPO <> 1"
		End If
		
		If FILTRO_AGREGADOS = 0 Then
			sql = sql&" AND TIPO <> 2 AND TIPO <> 3 AND TIPO <> 4 AND TIPO <> 5 AND TIPO <> 6"  
		End If
		
		If CHEQ_DEVOLVIDOS = 1 Then
		'Somente devolvidos
				sql = sql&" AND DATA_RETORNO IS NOT NULL"
		Elseif CHEQ_DEVOLVIDOS = 2 Then
		'Somente não devolvidos
				sql = sql&" AND DATA_RETORNO IS NULL AND FINALIZA_VENDEDOR <> 1"
		'---
		Elseif CHEQ_DEVOLVIDOS = 4 Then
	
		   Dim data_inicial
		   data_inicial = dateAdd("d",-30,Now())
		   data_inicial = Year(data_inicial)&"-"&Month(data_inicial)&"-"&Day(data_inicial)
		
				sql = sql&" AND DATA_RETORNO IS NULL AND FINALIZA_VENDEDOR <> 1 AND DATA_SAIDA <= '"&data_inicial&" 23:59:59'"
		'Elseif CHEQ_DEVOLVIDOS = 3 Then
		'Devolvidos e Não devolvidos
			'Não faz nada!!!
		End If
		
		'Response.Write(sql)
		Set rs = objConexao.Execute(sql, qtdRegistros)
		
		If not rs.Eof Then
		%>
      <table width="100%" 
		      border="1" align="center"
					cellpadding="5" 
					style="border-collapse:collapse">
					<tr bgcolor="whitesmoke">
		  				    <td colspan="8"><div align="center"><strong>Quantidade Registros: <%=qtdRegistros%></strong></div></td>
                        </tr>
<tr bgcolor="whitesmoke">
                <td width="26%" height="28"><div align="center"><span class="MensagemFicha"><strong>Formando</strong></span><strong><span class="MensagemFicha"></span></strong></div></td>
                <td width="26%"><div align="center"><strong>Instituição de Ensino</strong></div></td>
                <td width="8%"><div align="center"><strong>Item </strong></div></td>
                <td width="11%"><div align="left" class="MensagemFicha">
                  <div align="center"><strong>Data Saída </strong></div>
                </div></td>
                <td width="11%"><div align="center" class="MensagemFicha"><strong>Data Retorno</strong></div></td>
                <td width="6%"><div align="center"><strong>Qtd de dias fora </strong></div></td>
                <td width="6%"><div align="center"><strong>Vendido</strong></div></td>
                <td width="6%"><div align="center"><strong>Status</strong></div></td>
</tr>
		<%
		
		While not rs.Eof
		
		%>     
				<tr>
		          <td><a href="dip_busca_detalhamento_formando.asp?cod_formando=<%=rs("COD_FORMANDO")%>"><%=digito(rs("COD_FORMANDO"))%></a> - <%=rs("NOME_FORMANDO")%></td>
                  <td><%=rs("DESCRICAO")%></td>
                  <td><%
				  TIPO					= rs("TIPO")
				  ENCADERNACOES_EXTRAS	= rs("ENCADERNACOES_EXTRAS")
				  If TIPO = 0 Then
		
						If ENCADERNACOES_EXTRAS = 1 Then
							Response.Write("Álbum Duplo")
						Else
							Response.Write("Álbum")
						End If
						
				  Elseif TIPO = 1 Then
				  		Response.Write("Ficha")
				  Elseif TIPO > 1 Then
				  		
						'Produto agregado, consultar qual é.
			sql = 	"SELECT TB_PRODUTOS.NOME_PRODUTO "&_
					"FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = "&_
					"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO WHERE "&_
					"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO = "&rs("COD_FORMANDO")&" AND "&_
					"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.TIPO = "&TIPO&" LIMIT 1"
			Set rsProdAgrePend = objConexao.Execute(sql)
				If not rsProdAgrePend.Eof Then
				%>
				<%=rsProdAgrePend("NOME_PRODUTO")%>					
			    <%
				End If
			rsProdAgrePend.Close
			Set rsProdAgrePend = Nothing
						
				  End If
				  
				  %>				  </td>
                  <td><%=arrumaData(rs("DATA_SAIDA"))%></td>
                  <td><%=arrumaData(rs("DATA_RETORNO"))%></td>
                  <td>
				   <div align="center">
				  <%
				  Dim QTD_DIAS_FORA_LINHA, DATA_SAIDA, DATA_RETORNO
				  
				  DATA_SAIDA	=	formatDateTime(rs("DATA_SAIDA"))
				  DATA_RETORNO	=	""&rs("DATA_RETORNO")
				  Dim DATA_RETORNO_2
				  DATA_RETORNO_2 = ""&rs("DATA_RETORNO")
				  
				  If DATA_RETORNO = "" Then
				  		DATA_RETORNO = now()
				  End If
				  
				  QTD_DIAS_FORA_LINHA = dateDiff("d",DATA_SAIDA,DATA_RETORNO)
				  Response.Write(QTD_DIAS_FORA_LINHA)
				  %>
				  </div></td>
                  <td>
				  <%
				  Dim VENDIDO
				  VENDIDO = formatNumber(0&rs("VENDIDO"),0)
				  %>
				  	<div align="center">
				  		<%
						If VENDIDO = 2 Then
						'Antes de exibir na tela verifica se não está pelo menos lançado no sistema
							Dim rsVera
							Set rsVera = objConexao.Execute("SELECT * FROM TB_PEDIDOS_FORMANDOS WHERE COD_FORMANDO ="&rs("COD_FORMANDO"))
							
									If not rsVera.Eof Then
										'Significa que existe pedido, porém não foi conferido pelo depto. de foto
										Response.Write("Sim, mas ainda não conferido")
										VENDIDO = 1
									Else
										Response.Write(booleano(rs("VENDIDO")))
									End If
							
							
							rsVera.Close
							Set rsVera = Nothing
						
							Else
							
								Response.Write("Sim")
						
							End If
						
						%>
							
					</div></td>
                  <td>
				  <div align="center">
				  <%If QTD_DIAS_FORA_LINHA => 30 and VENDIDO = 2 and DATA_RETORNO_2 = "" Then%>
				  	<img src="images/icons/icon_balao_red.gif" width="17" height="16" />	
				  <%Elseif QTD_DIAS_FORA_LINHA => 25 and VENDIDO = 2 and DATA_RETORNO_2 = "" Then   %>
				  	<img src="images/icons/icon_balao_yellon.gif" width="17" height="16" />
				  <%Else%>
				  	<img src="images/icons/icon_balao_green.gif" width="17" height="16" />
				  <%End If%>
				  				  </div></td>
				</tr>
					 
			<%
			rs.MoveNext()
			Wend
			
			rs.Close
			Set rs = Nothing
			
			FecharConexao()
			%>		 
        
		  	<tr bgcolor="whitesmoke">
		  				    <td colspan="8"><div align="center"><strong>Quantidade Registros: <%=qtdRegistros%></strong></div></td>
                        <tr>
                          <td colspan="8"><strong>Legenda</strong>:<br />
                            <img src="images/icons/icon_balao_green.gif" width="17" height="16" /> - Item em dia<br />
                            <img src="images/icons/icon_balao_yellon.gif" width="17" height="16" /> - Item próximo de atraso<br />
                            <img src="images/icons/icon_balao_red.gif" width="17" height="16" /> - Item atrasado<br />
                            <br />
                            <br />
                            <strong>Observações:</strong><br />
Itens não devolvidos, mas já lançados pelo sistema pelo representante comercial <strong>não</strong> constam como atrasados mesmo que não foram corrigidos pelo departamento fotográfico. <br />
Caso o item for vendido sua data retorno será a mesma data do lançamento do sistema pelo representante comercial, entretanto esta data  é lançada no final da correção efetuada pelo depto. de foto e reportagens.</td>
                        </table>
		
		  <p>
           </p>
		    <%Else%>
		   <div align="center">
           <br /><br /><br /><br />
		  <p><span class="textoMedio"><strong>Não há dados para os filtros selecionados!</strong></span></p>
          <br /><br /><br /><br />
		</div>
		    <%End If%>
		   
		    

		<div align="center"><br />
            <br />
            <input class="btnTopo" type="button" onclick="abre('#')" name="Submit2" value=" " />
            <br />
            <br />
            <br />
		  
		  </div></td>

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



