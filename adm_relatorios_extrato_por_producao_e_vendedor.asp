<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Session.Timeout=5
'Valida��o especial para verificar se possui sub-valida��o espec�fica
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '31' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se Não encontrar ent�o redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
		
	rs.Close
	Set rs=Nothing
	
FecharConexao()

Dim COD_PRODUCAO
COD_PRODUCAO = Request.QueryString("COD_PRODUCAO")

AbrirConexao()

'1) Deleta TB TEMPOR�RIA antes de inicias
objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA WHERE COD_USUARIO ="&Session("COD_USUARIO"))


'2) Seleciona todos os formandos de uma determinada produ��o
Set rs = objConexao.Execute("SELECT TB_FICHAS_IDENTIFICACAO.COD_FORMANDO,TB_FICHAS_IDENTIFICACAO.COD_CONTRATO,TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO,TB_PEDIDOS_FORMANDOS.COD_PEDIDO,TB_PEDIDOS.DATA_VENDA FROM TB_FICHAS_IDENTIFICACAO LEFT JOIN TB_PEDIDOS_FORMANDOS ON TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = TB_PEDIDOS_FORMANDOS.COD_FORMANDO LEFT JOIN TB_PEDIDOS ON TB_PEDIDOS_FORMANDOS.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO  WHERE TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO = '"&COD_PRODUCAO&"' GROUP BY COD_FORMANDO")
		
		Dim ROTINA_COD_FORMANDO, ROTINA_COD_CONTRATO, ROTINA_TIPO, ROTINA_BRINDE, ROTINA_COD_ULTIMO_VENDEDOR_FICHA, ROTINA_VENDIDO, ROTINA_COD_PEDIDO
		
		While not rs.Eof
		
		ROTINA_COD_FORMANDO = rs("COD_FORMANDO")
		ROTINA_COD_CONTRATO = rs("COD_CONTRATO")

		'====================================================
' GARANTE QUE MOVIMENTAÇÕES EXISTAM (TIPO 0, 1 e 2)
'====================================================
Dim existe

'--- TIPO 0 (ÁLBUM FÍSICO) ---
sql = "SELECT COD_MOVIMENTACAO FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES " &_
      "WHERE COD_FORMANDO=" & ROTINA_COD_FORMANDO & " AND TIPO=0 LIMIT 1"
Set rsCheck = objConexao.Execute(sql)

If rsCheck.Eof Then
    objConexao.Execute "INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES " & _
                       "(COD_FORMANDO, TIPO, CADASTRADO_POR, DATA_CADASTRO) VALUES (" & _
                       ROTINA_COD_FORMANDO & ", 0, " & Session("COD_USUARIO") & ", NOW())"
End If

rsCheck.Close

'--- TIPO 1 (FICHA) ---
sql = "SELECT COD_MOVIMENTACAO FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES " &_
      "WHERE COD_FORMANDO=" & ROTINA_COD_FORMANDO & " AND TIPO=1 LIMIT 1"
Set rsCheck = objConexao.Execute(sql)

If rsCheck.Eof Then
    objConexao.Execute "INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES " & _
                       "(COD_FORMANDO, TIPO, CADASTRADO_POR, DATA_CADASTRO) VALUES (" & _
                       ROTINA_COD_FORMANDO & ", 1, " & Session("COD_USUARIO") & ", NOW())"
End If

rsCheck.Close

'--- TIPO 2 (ÁLBUM DIGITAL) ---
sql = "SELECT COD_MOVIMENTACAO FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES " &_
      "WHERE COD_FORMANDO=" & ROTINA_COD_FORMANDO & " AND TIPO=2 LIMIT 1"
Set rsCheck = objConexao.Execute(sql)

If rsCheck.Eof Then
    objConexao.Execute "INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES " & _
                       "(COD_FORMANDO, TIPO, CADASTRADO_POR, DATA_CADASTRO) VALUES (" & _
                       ROTINA_COD_FORMANDO & ", 2, " & Session("COD_USUARIO") & ", NOW())"
End If

rsCheck.Close
Set rsCheck = Nothing
'====================================================

						
			'3) Verifica quem foi o �ltimo respons�vel pelo formando
			sql = " SELECT COD_VENDEDOR FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_FORMANDO ="&ROTINA_COD_FORMANDO&_
				  " AND TIPO = 1 ORDER BY DATA_SAIDA DESC LIMIT 1"
			Set rs2 = objConexao.Execute(sql)
						
						If not rs2.Eof Then
						
							ROTINA_COD_ULTIMO_VENDEDOR_FICHA = rs2("COD_VENDEDOR")
						
						Else
							'Se Não encontrar acusa falha
							Response.Redirect("adm_relatorios.asp?msg=FALHA AO CARREGAR RELATÓRIO, FORMANDO: "&ROTINA_COD_FORMANDO)
											
						End If
						
			rs2.Close
			Set rs2 = Nothing

			'4) Verifica quais eram os itens do formando			
				'Não � necess�rio a ficha, somente produtos
			sql = " SELECT DISTINCT (TIPO) FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_FORMANDO = "&rs("COD_FORMANDO")&_
				  " AND TIPO <> 1 ORDER BY TIPO"

			Set rs2 = objConexao.Execute(sql)
					
					While not rs2.Eof
					'Agora sabendo quais os produtos e o c�d. do �ltimo vendedor, verifica se foi vendido
					
					ROTINA_TIPO = formatNumber(0&rs2("TIPO"),0)
							
							If ROTINA_TIPO = 0 Then
							'5) Come�a pelo �lbum se est� vendido, �lbum Não ser� considerado como brinde, pois pode ser vendido parcialmente
									'Pega da pr�pria TB_FICHAS_IDENTIFICAO do formando j� aberta
									ROTINA_VENDIDO = 2
									ROTINA_COD_PEDIDO = 0
									Set rs3=objConexao.Execute("SELECT IFNULL(TB_PEDIDOS_FORMANDOS.COD_PEDIDO,0) AS VENDIDO FROM TB_FICHAS_IDENTIFICACAO LEFT JOIN TB_PEDIDOS_FORMANDOS ON TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = TB_PEDIDOS_FORMANDOS.COD_FORMANDO LEFT JOIN TB_PEDIDOS ON TB_PEDIDOS_FORMANDOS.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO WHERE TB_PEDIDOS.FINALIZA_VENDEDOR = 1 AND  TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = "&ROTINA_COD_FORMANDO)
										
										If not rs3.Eof Then
											
											If int(0&rs3("VENDIDO")) > 0 Then
												ROTINA_VENDIDO = 1
												ROTINA_COD_PEDIDO = rs3("VENDIDO")
											End IF
											
										End If
										
									rs3.Close
									Set rs3=Nothing
									
									ROTINA_BRINDE  = 2
																		
							Else
							'AGREGADOS SE HOUVER
							'//////////
							
							'6) Se tiver mais verifica junto a tabela 
								
							
							'7)Verifica se Não era de brinde o produto agregado
								
								'Não SER� VERIFICADO OS PRODUTOS AGREGADOS QUE ERAM DE COMISS�O, SEM ANTES
								'FINALIZAR A ROTINA COLOCANDO OS TIPOS NAS TBS QUE FALTAM
								
								ROTINA_BRINDE  = 2						
							'//////////
							End If
							
							'Ap�s isso tudo faz o INSERT NA TB TEMPOR�RIA
							sql = " INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA (COD_FORMANDO, COD_PRODUCAO, "&_
								  " COD_CONTRATO, TIPO, BRINDE, COD_ULTIMO_VENDEDOR_FICHA, VENDIDO, COD_USUARIO,COD_PEDIDO) VALUES ("&_
								  ROTINA_COD_FORMANDO&", "&COD_PRODUCAO&", "&sqlTexto(ROTINA_COD_CONTRATO)&", "&ROTINA_TIPO&", "&ROTINA_BRINDE&", "&_
								  ROTINA_COD_ULTIMO_VENDEDOR_FICHA&", "&ROTINA_VENDIDO&", "&Session("COD_USUARIO")&_
								  ", "&sqlTexto(ROTINA_COD_PEDIDO)&_								  
								  " ) "
							
							'rESPONSE.Write(sql)	   
							objConexao.Execute(sql)
							
					rs2.MoveNext()
					Wend
							
						
			rs2.Close
			Set rs2 = Nothing
					
				
				
		rs.MoveNext()
		Wend
		

rs.Close
Set rs = Nothing

Set rs=objConexao.Execute("Select descricao FROM TB_PRODUCAO WHERE COD_PRODUCAO = "&COD_PRODUCAO)
	If not rs.Eof Then
		DESCRICAO = rs(0)
	End If
rs.Close
Set rs = Nothing


FecharConexao()

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
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
	var text = "Se a janela Não abrir desbloquei os Pop-ups!";
	if(win == null) { alert(text); return; }
	win.moveTo(350,150);
}	
	</script>
    <style type="text/css">
<!--
.style1 {color: #006600}
.style2 {color: #CC0000}
.style3 {font-size: 10px}
-->
    </style>
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
  
   BTN_IMPRIMIR_LINK = "href=""javascript:popUp('adm_relatorios_extrato_vendedores_IMPRIMIR.asp?COD_VENDEDOR="&COD_VENDEDOR&"&DATA_INI="&sqlData(DATA_INI)&"&DATA_FIM="&sqlData(DATA_FIM)&"&FILTRO_ALBUNS="&FILTRO_ALBUNS&"&FILTRO_AGREGADOS="&FILTRO_AGREGADOS&"&CHEQ_DEVOLVIDOS="&CHEQ_DEVOLVIDOS&"')"""
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">
	
	  <h2>Relatórios - Extrato de Produções        (mesmo sem correção)<br />
	    
	    
          <br />
      </h2>
	  <div align="center" class="msg"><%=Request.QueryString("msg")%></div>
	  <%
		AbrirConexao()		
		Dim rs, sql, i, produtos(100,5)
		i = 0
		sql = "SELECT * FROM TB_PRODUCAO WHERE COD_PRODUCAO ="&COD_PRODUCAO
		Dim DESCRICAO, DATA_FINALIZADA, PRODUTOS_AGREDADOS
		
		
		Set rs = objConexao.Execute(sql, qtdRegistros)
		
		If not rs.Eof Then

		%>
        <span class="textoMaior"><strong><%=COD_PRODUCAO%> - <%=DESCRICAO%></strong></span>
      <table width="100%" 
		      border="1" align="center"
					cellpadding="5" 
					style="border-collapse:collapse">
<tr bgcolor="#DDDDDD">
                <td width="14%" height="28" rowspan="2"><div align="center"><strong>Contratos</strong></div></td>
                <td width="15%" rowspan="2"><div align="center"><strong>Data Finalizada / Lotes</strong></div></td>
                <td colspan="2"><div align="center"><strong>Qtd. Fotos - Produção</strong></div></td>
                <td colspan="2"><div align="center"><strong>Qtd. Fotos - Albuns</strong></div></td>                
                <td colspan="4"><div align="center"><strong>Álbuns</strong></div></td>
                <td width="32%" rowspan="2"><div align="center"><strong>Venda de Produtos       </strong></div>
  
                  <%
				  Set rs4=objConexao.Execute("SELECT TB_PEDIDOS_PRODUTOS.COD_PRODUTO, TB_PRODUTOS.NOME_PRODUTO, SUM( TB_PEDIDOS_PRODUTOS.QUANTIDADE ) AS QUANTIDADE FROM TB_PEDIDOS INNER JOIN TB_PEDIDOS_PRODUTOS ON ( TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO )INNER JOIN TB_PRODUTOS ON ( TB_PEDIDOS_PRODUTOS.COD_PRODUTO = TB_PRODUTOS.COD_PRODUTO ) INNER JOIN TB_FICHAS_IDENTIFICACAO ON ( TB_PEDIDOS_PRODUTOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO ) WHERE TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO ="&COD_PRODUCAO&" AND TB_PEDIDOS.FINALIZA_VENDEDOR = 1 GROUP BY TB_PEDIDOS_PRODUTOS.COD_PRODUTO",QTD_PRODUTOS)
				While not rs4.Eof
					
						
						PRODUTOS(i,0) = rs4("COD_PRODUTO")
						PRODUTOS(i,1) = rs4("NOME_PRODUTO")
						PRODUTOS(i,2) = rs4("QUANTIDADE")						
						i=i+1
					rs4.MoveNext()
					Wend
					
				  rs4.Close
				  Set rs4=Nothing
				  %>
&nbsp;  </td>
                <td width="6%" rowspan="2"><div align="center"><strong>Total vendas*</strong></div></td>
</tr>
<tr bgcolor="#DDDDDD">
  <td width="9%"><div align="center"><strong>Produzidas</strong></div></td>
  <td width="9%"><div align="center"><strong>Não Vendidas</strong></div></td>
  <td width="9%"><div align="center"><strong>N&atilde;o Vendidas</strong></div></td>
  <td width="6%"><div align="center"><strong>Vendidas</strong></div></td>
<%
Dim TOT_ALBUNS_TOTAL, TOT_ALBUNS_TOTAL_VENDIDO, TOT_ALBUNS_TOTAL_N_VENDIDO, TOT_ALBUNS_APROVEITAMENTO
Dim TOT_AGREGADOS_TOTAL, TOT_AGREGADOS_TOTAL_VENDIDO, TOT_AGREGADOS_TOTAL_N_VENDIDO, TOT_AGREGADOS_APROVEITAMENTO
Dim TOT_TOTAL_ARRECADADO, TOT_A_SER_ARRECADADO, TOT_APROVEITAMENTO_MEDIO
Dim TOTAL_GERAL_FOTOS_FATURADAS, TOTAL_GERAL_FOTOS_INICIAL, TOTAL_GERAL_FOTOS_APROVEITAMENTO
Dim TOTAL_GERAL_PAGAMENTO_EM_DINHEIRO, TOTAL_GERAL_XTOTAL_CHEQUE, TOTAL_GERAL_XTOTAL_CARTAO, TOTAL_GERAL

TOT_ALBUNS_TOTAL = 0
TOT_ALBUNS_TOTAL_VENDIDO = 0
TOT_ALBUNS_TOTAL_N_VENDIDO = 0
TOT_ALBUNS_APROVEITAMENTO = 0
TOT_AGREGADOS_TOTAL = 0
TOT_AGREGADOS_TOTAL_VENDIDO = 0
TOT_AGREGADOS_TOTAL_N_VENDIDO = 0
TOT_AGREGADOS_APROVEITAMENTO = 0
TOT_TOTAL_ARRECADADO = 0
TOT_A_SER_ARRECADADO = 0
TOT_APROVEITAMENTO_MEDIO = 0
TOTAL_GERAL_FOTOS_FATURADAS = 0
TOTAL_GERAL_FOTOS_INICIAL = 0
TOTAL_GERAL_FOTOS_APROVEITAMENTO = 0
%>
  <td width="4%"><div align="center" class="style3">Qtd. Total </div></td>
  <td width="5%"><div align="center" class="style3">Qtd. Vendida </div></td>
  <td width="5%"><div align="left" class="MensagemFicha style3">
    <div align="center">Qtd. Não Vendida </div>
  </div></td>
  <td width="4%"><div align="center" class="style3">Aprov. Álbuns</div></td>
  </tr>
		<%
		Dim ou, au
		ou = 0
		au = 0
		
		DESCRICAO			= rs("DESCRICAO")
		DATA_FINALIZADA		= rs("DATA_FINALIZADA")
		
		'Response.Write(DATA_FINALIZADA)
		PRODUTOS_AGREDADOS	= formatNumber(0&rs("PRODUTOS_AGREDADOS"),0)
		
		'Qtd total de Álbuns
		Set rs2 = objConexao.Execute("SELECT COD_PRODUCAO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = "&COD_PRODUCAO,QTD_ALBUNS)
		
		rs2.Close
		Set rs2 = Nothing
		
		'Qtd Álbuns Vendidos, nao precisa estar conferido
		Set rs3 = objConexao.Execute("SELECT TB_FICHAS_IDENTIFICACAO.COD_FORMANDO,TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO,TB_PEDIDOS_FORMANDOS.COD_PEDIDO,TB_PEDIDOS.DATA_VENDA,IF(TB_PEDIDOS_FORMANDOS.COD_PEDIDO is NULL,0,1) AS QTD_VENDIDO FROM TB_FICHAS_IDENTIFICACAO LEFT JOIN TB_PEDIDOS_FORMANDOS ON TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = TB_PEDIDOS_FORMANDOS.COD_FORMANDO LEFT JOIN TB_PEDIDOS ON TB_PEDIDOS_FORMANDOS.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO WHERE TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO = '"&COD_PRODUCAO&"' and IF(TB_PEDIDOS_FORMANDOS.COD_PEDIDO is NULL,0,1) = 1 GROUP BY COD_FORMANDO",QTD_ALBUNS_VENDIDOS)
				
		rs3.Close
		Set rs3 = Nothing
		
		Dim QTD_PRODUTOS_AGREGADOS_TOTAL, QTD_PRODUTOS_AGREGADOS_VENDIDOS, QTD_PRODUTOS_AGREGADOS_NAO_VENDIDOS
				
		
		Set rs5 = objConexao.Execute("SELECT * FROM TB_CACHE_VW_TOTAL_ARRECADADO_POR_PRODUCAO WHERE COD_PRODUCAO="&COD_PRODUCAO)
		Dim TOTAL_ARRECADADO
		TOTAL_ARRECADADO = 0			
					If not rs5.Eof Then
						TOTAL_ARRECADADO = formatNumber(0&rs5("TOTAL_ARRECADADO"))
					End If
					
		rs5.Close
		Set rs5 = Nothing
		
		'M�dia de Quanto Falta
		sql = "SELECT SUM(INICIAL) AS TOTAL_INICIAL FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = "&COD_PRODUCAO&" AND VENDIDO = 2 GROUP BY COD_PRODUCAO"
		
		Dim QTD_FOTOS_NAO_VENDIDAS
		QTD_FOTOS_NAO_VENDIDAS = 0
		Set rs6 = objConexao.Execute(sql)
			If not rs6.Eof Then
				QTD_FOTOS_NAO_VENDIDAS = formatNumber(0&rs6("TOTAL_INICIAL"),0)
			End If
		Set rs6 = Nothing

		
		
		Dim TOTAL_A_RECEBER_ALBUNS, TOTAL_A_RECEBER_PRODUTOS_AGREGADOS, TOTAL_NAO_ARRECADADO_GERAL
		TOTAL_A_RECEBER_ALBUNS				= 0
		TOTAL_A_RECEBER_PRODUTOS_AGREGADOS	= 0
		
		'1 de 2 totais a serem arrecadados - ALBUNS
		TOTAL_A_RECEBER_ALBUNS = QTD_FOTOS_NAO_VENDIDAS*23

		'2 de 2 totais arrecadados - PRODUTOS_AGREGAFDOS
		
		If PRODUTOS_AGREDADOS = 1 Then
					
			
					
			Set rs7 = objConexao.Execute("SELECT COD_PRODUTO, COUNT(*) AS QTD_PRODUTOS_AGREGADOS FROM TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS WHERE COD_PRODUCAO = "&COD_PRODUCAO&" AND VENDIDO <> 1 GROUP BY COD_PRODUTO, COD_PRODUCAO ")
			
				If not rs7.Eof Then
					Dim COD_PRODUTO_AGREGADO_1
					TOTAL_A_RECEBER_PRODUTOS_AGREGADOS = 0
					While not rs7.Eof
							
							COD_PRODUTO_AGREGADO_1 = formatNumber(rs7("COD_PRODUTO"),0)
							
								If COD_PRODUTO_AGREGADO_1 = 9 Then
									'Slide Show	
									TOTAL_A_RECEBER_PRODUTOS_AGREGADOS = TOTAL_A_RECEBER_PRODUTOS_AGREGADOS + (formatNumber(0&rs7("QTD_PRODUTOS_AGREGADOS"),0)*180)						
								Else
									'Outros
									TOTAL_A_RECEBER_PRODUTOS_AGREGADOS = TOTAL_A_RECEBER_PRODUTOS_AGREGADOS + (formatNumber(0&rs7("QTD_PRODUTOS_AGREGADOS"),0)*260)
								End If 
								
				
							
					rs7.MoveNext()
					Wend
					
					
					End If
			
			rs7.Close
			Set rs7 = Nothing
			
			'/////
			
			
			Set rs71 = objConexao.Execute("SELECT COUNT(*) AS QTD_PRODUTOS_AGREGADOS FROM TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS WHERE COD_PRODUCAO = "&COD_PRODUCAO&" AND VENDIDO = 2 GROUP BY COD_PRODUCAO")
					
					If not rs71.Eof Then
							
							QTD_NAO_VENDIDA_DEFINITIVAMENTE = formatNumber(rs71("QTD_PRODUTOS_AGREGADOS"),0)
							
					End If
					
					
			rs71.Close
			Set rs71 = Nothing
			'/////
			
		Else
			TOTAL_A_RECEBER_PRODUTOS_AGREGADOS = 0
		End if
				
			TOTAL_NAO_ARRECADADO_GERAL = (TOTAL_A_RECEBER_PRODUTOS_AGREGADOS+TOTAL_A_RECEBER_ALBUNS)/1
			
		Dim APROVEITAMENTO
		
		APROVEITAMENTO = formatNumber(100-((TOTAL_NAO_ARRECADADO_GERAL)*100/(TOTAL_ARRECADADO+TOTAL_NAO_ARRECADADO_GERAL)),1)
		
		'/////////////////////////////////////////////
		'FOTOS
		TOTAL_GERAL_FOTOS_FATURADAS = 0
		TOTAL_GERAL_FOTOS_INICIAL = 0
		TOTAL_GERAL_FOTOS_APROVEITAMENTO = 0

			Set rs5=objConexao.Execute("SELECT sum(INICIAL) AS INICIAL,SUM(FATURADO) AS FATURADO FROM VW_PEDIDOS_COMPLETO WHERE FINALIZA_VENDEDOR = 1 AND COD_PRODUCAO = "&COD_PRODUCAO)
				
				If not rs5.Eof Then
					
					TOTAL_GERAL_FOTOS_FATURADAS = formatNumber(0&rs5("FATURADO"),0)
					TOTAL_GERAL_FOTOS_INICIAL = formatNumber(0&rs5("INICIAL"),0)
					If TOTAL_GERAL_FOTOS_INICIAL > 0 Then
						TOTAL_GERAL_FOTOS_APROVEITAMENTO = clng(formatNumber(((100*TOTAL_GERAL_FOTOS_FATURADAS)/TOTAL_GERAL_FOTOS_INICIAL),1))
					Else
						TOTAL_GERAL_FOTOS_APROVEITAMENTO = 0	
					End If
				End If
				
			rs5.Close
			Set rs5=Nothing
		
		'/////////////////////////////////////////////
		'Total arrecadado
				TOTAL_GERAL_PAGAMENTO_EM_DINHEIRO = 0
				TOTAL_GERAL_XTOTAL_CHEQUE = 0
				TOTAL_GERAL_XTOTAL_CARTAO = 0
				TOTAL_GERAL = 0
				
				Set rs5=objConexao.Execute("SELECT SUM(TB_CACHE_VW_PEDIDOS_TOTAL_RECEBIDO.PAGAMENTO_EM_DINHEIRO) AS PAGAMENTO_EM_DINHEIRO, SUM(TB_CACHE_VW_PEDIDOS_TOTAL_RECEBIDO.XTOTAL_CHEQUE) AS XTOTAL_CHEQUE, SUM(TB_CACHE_VW_PEDIDOS_TOTAL_RECEBIDO.XTOTAL_CARTAO) AS XTOTAL_CARTAO FROM VW_PEDIDOS_COMPLETO INNER JOIN TB_CACHE_VW_PEDIDOS_TOTAL_RECEBIDO ON VW_PEDIDOS_COMPLETO.COD_PEDIDO = TB_CACHE_VW_PEDIDOS_TOTAL_RECEBIDO.COD_PEDIDO WHERE VW_PEDIDOS_COMPLETO.FINALIZA_VENDEDOR = 1 AND VW_PEDIDOS_COMPLETO.COD_PRODUCAO = "&COD_PRODUCAO)
				
				If not rs5.Eof Then
					
					TOTAL_GERAL_PAGAMENTO_EM_DINHEIRO = formatNumber(0&rs5("PAGAMENTO_EM_DINHEIRO"),2)/1
					TOTAL_GERAL_XTOTAL_CHEQUE = formatNumber(0&rs5("XTOTAL_CHEQUE"),2)/1
					TOTAL_GERAL_XTOTAL_CARTAO = formatNumber(0&rs5("XTOTAL_CARTAO"),2)/1
					TOTAL_GERAL = (TOTAL_GERAL_PAGAMENTO_EM_DINHEIRO+TOTAL_GERAL_XTOTAL_CHEQUE+TOTAL_GERAL_XTOTAL_CARTAO)/1
					
				End If
				
			rs5.Close
			Set rs5=Nothing
            
            TOTAL_FOTO_PRODUCAO = 0
            'Total de fotos na produ��o
				Set rs5=objConexao.Execute("SELECT SUM(INICIAL) AS INICIAL FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = "&COD_PRODUCAO)
				If not rs5.Eof Then
					TOTAL_FOTO_PRODUCAO = rs5("INICIAL")
                    TOTAL_FOTO_PRODUCAO_PARADA = clng(TOTAL_FOTO_PRODUCAO) - clng(TOTAL_GERAL_FOTOS_FATURADAS)
                    TOTAL_FOTO_PRODUCAO_PARADA =  formatNumber(0&TOTAL_FOTO_PRODUCAO_PARADA,0)
                    TOTAL_FOTO_PRODUCAO =  formatNumber(0&TOTAL_FOTO_PRODUCAO,0)                    
				End If
				
			rs5.Close
			Set rs5=Nothing
            
            
      

            
		
		%>     
				<tr bgcolor="whitesmoke">
		          <td style="border-bottom:#666666 solid 2px">
                    <div align="center">
                      <%
			Dim rsContratos, qtdContratos, CONTRATOS_ULT
			
			Set rsContratos = objConexao.Execute("SELECT DISTINCT(COD_CONTRATO) AS COD_CONTRATO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO="&COD_PRODUCAO&" ORDER BY COD_CONTRATO ASC",qtdContratos)
			
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
			
			'Dim CONTRATOS_TRATAMENTO
			'CONTRATOS_TRATAMENTO = Left(CONTRATOS_ULT,Len(CONTRATOS_ULT)-2)&Replace(Right(CONTRATOS_ULT,2),", ","")
			'Response.Write(CONTRATOS_TRATAMENTO)
			%>                  
                    </div></td>
                  <td  style="border-bottom:#666666 solid 2px"><div align="center"><%=ArrumaData(DATA_FINALIZADA)%></div></td>
                  <td  style="border-bottom:#666666 solid 2px"><div align="center"><%=TOTAL_FOTO_PRODUCAO%></div></td>
                  <td  style="border-bottom:#666666 solid 2px"><div align="center"><%=TOTAL_FOTO_PRODUCAO_PARADA%></div></td>
                  <td  style="border-bottom:#666666 solid 2px"><div align="center"><%=formatNumber(0&TOTAL_GERAL_FOTOS_INICIAL-TOTAL_GERAL_FOTOS_FATURADAS,0)%></div></td>
                  <td  style="border-bottom:#666666 solid 2px"><div align="center"><%=TOTAL_GERAL_FOTOS_FATURADAS%></div></td>
                  <td  style="border-bottom:#666666 solid 2px"><div align="center"><%=QTD_ALBUNS%></div></td>
                  <td  style="border-bottom:#666666 solid 2px">
                    <div align="center"><span class="style1"><%=QTD_ALBUNS_VENDIDOS%></span></div></td>
                  <td  style="border-bottom:#666666 solid 2px"><div align="center"><span class="style2"><%=QTD_ALBUNS-QTD_ALBUNS_VENDIDOS%></span></div></td>
                  <td  style="border-bottom:#666666 solid 2px"><div align="center">
				  <%
				  Dim APROVEITAMENTO_ALBUNS
				  		APROVEITAMENTO_ALBUNS = formatNumber(100-((QTD_ALBUNS-QTD_ALBUNS_VENDIDOS)*100/(QTD_ALBUNS_VENDIDOS+QTD_ALBUNS-QTD_ALBUNS_VENDIDOS)),0)
						Response.Write(APROVEITAMENTO_ALBUNS)
				  %>%
				  </div></td>
                  <td style="border-bottom:#666666 solid 2px">
                  
                  <%
				  If int(QTD_PRODUTOS) > 1 Then
				  %>
                  <table style="border-collapse:collapse; font-size:10px" width="100%" 
		      border="1" align="center"
					cellpadding="2" 
					>
                  <tr>
                  	<%For i=0 to QTD_PRODUTOS-1
					%>
                      <td><div align="center"><%=PRODUTOS(i,1)%></div></td>
                    <%
					Next
					%>
                    </tr>
                    <tr>
                  	<%For i=0 to QTD_PRODUTOS-1
					%>
                      <td><div align="center"><%=PRODUTOS(i,2)%></div></td>
                    <%
					Next
					%>
                    </tr>
                  </table>
                  <%End If%>&nbsp;                  </td>
				  <td style="border-bottom:#666666 solid 2px"><div align="right"><%=formatNumber(TOTAL_GERAL,2)%></div></td>
				</tr>
				
	<%
				sql79 = 	" SELECT TB_VENDEDORES.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR "&_
							" FROM TB_VENDEDORES "&_
							" INNER JOIN TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA ON "&_
							" TB_VENDEDORES.COD_VENDEDOR = TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA.COD_ULTIMO_VENDEDOR_FICHA "&_
							" WHERE COD_PRODUCAO ="&COD_PRODUCAO&_
							" GROUP BY COD_ULTIMO_VENDEDOR_FICHA ORDER BY TB_VENDEDORES.NOME_VENDEDOR"

				Set rs112 = objConexao.Execute(sql79)
				
				Dim TOTAL_QTD_ALBUNS
				TOTAL_QTD_ALBUNS = 0				
	
					'Totais
				
				Dim SUB_TOTAL_QTD_ALBUNS, SUB_TOTAL_QTD_ALBUNS_VENDIDOS, SUB_APROVEITAMENTO_ALBUNS
				Dim SUB_TOTAL_QTD_AGREGADOS, SUB_TOTAL_QTD_AGREGADOS_VENDIDOS, SUB_APROVEITAMENTO_AGREGADOS
				Dim SUB_APROVEITAMENTO_MEDIO
				Dim TOTAL_FOTOS_FATURADAS, TOTAL_FOTOS_INICIAL, TOTAL_FOTOS_APROVEITAMENTO	
				Dim TOTAL_PAGAMENTO_EM_DINHEIRO, TOTAL_XTOTAL_CHEQUE, TOTAL_XTOTAL_CARTAO, TOTAL					
								
				While not rs112.Eof
				
				SUB_TOTAL_QTD_ALBUNS = 0
				SUB_TOTAL_QTD_ALBUNS_VENDIDOS = 0
				SUB_APROVEITAMENTO_ALBUNS = 0
				SUB_TOTAL_QTD_AGREGADOS = 0
				SUB_TOTAL_QTD_AGREGADOS_VENDIDOS = 0
				SUB_APROVEITAMENTO_AGREGADOS = 0
				SUB_APROVEITAMENTO_MEDIO = 0
				
				
				
						sql = " SELECT COUNT(*) AS QTD_ALBUNS FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA WHERE "&_
							  " TIPO = 0 AND COD_ULTIMO_VENDEDOR_FICHA ="&rs112("COD_VENDEDOR")&" AND COD_USUARIO = '"&Session("COD_USUARIO")&"'"
						
						Set rsSubTotais = objConexao.Execute(sql)
								
								'Response.Write(sql&"<br />")
								
								If not rsSubTotais.Eof Then
								
									SUB_TOTAL_QTD_ALBUNS = formatNumber(0&rsSubTotais("QTD_ALBUNS"),0)
									'TOTAL_QTD_ALBUNS = TOTAL_QTD_ALBUNS + SUB_TOTAL_QTD_ALBUNS
									'Response.Write(rsSubTotais("QTD_ALBUNS")&"<br />")
								End If
								
						rsSubTotais.Close
						Set rsSubTotais = Nothing
						
						
						
						'//////////////
						
						sql = " SELECT COUNT(*) AS QTD_ALBUNS FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA WHERE "&_
							  " TIPO = 0 AND COD_ULTIMO_VENDEDOR_FICHA ="&rs112("COD_VENDEDOR")&" AND VENDIDO = 1 AND COD_USUARIO = '"&Session("COD_USUARIO")&"'"

						
						Set rsSubTotais = objConexao.Execute(sql)
								
								If not rsSubTotais.Eof Then
								
									SUB_TOTAL_QTD_ALBUNS_VENDIDOS = formatNumber(0&rsSubTotais("QTD_ALBUNS"),0)
									'TOTAL_QTD_ALBUNS = TOTAL_QTD_ALBUNS + SUB_TOTAL_QTD_ALBUNS_VENDIDOS
								End If
								
						rsSubTotais.Close
						Set rsSubTotais = Nothing
						
						
						'/////////////
						'Produtos Agregados Totais
						If PRODUTOS_AGREDADOS = 1 Then
						
							sql =	" SELECT COUNT(*) AS QTD_AGREGADOS FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA WHERE "&_
							  		" TIPO > 1 AND COD_ULTIMO_VENDEDOR_FICHA ="&rs112("COD_VENDEDOR")&" AND COD_USUARIO = '"&Session("COD_USUARIO")&"'"
						
							Set rsSubTotais = objConexao.Execute(sql)
								
								If not rsSubTotais.Eof Then
								
									SUB_TOTAL_QTD_AGREGADOS = formatNumber(0&rsSubTotais("QTD_AGREGADOS"),0)
									'TOTAL_QTD_ALBUNS = TOTAL_QTD_ALBUNS + SUB_TOTAL_QTD_AGREGADOS
								End If
								
							rsSubTotais.Close
							Set rsSubTotais = Nothing
							
							
							sql =	" SELECT COUNT(*) AS QTD_AGREGADOS FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA WHERE "&_
							  		" TIPO > 1 AND COD_ULTIMO_VENDEDOR_FICHA ="&rs112("COD_VENDEDOR")&" AND VENDIDO = 1 AND COD_USUARIO = '"&Session("COD_USUARIO")&"'"
						
							Set rsSubTotais = objConexao.Execute(sql)
								
								If not rsSubTotais.Eof Then
								
									SUB_TOTAL_QTD_AGREGADOS_VENDIDOS = formatNumber(0&rsSubTotais("QTD_AGREGADOS"),0)
									'TOTAL_QTD_ALBUNS = TOTAL_QTD_ALBUNS + SUB_TOTAL_QTD_AGREGADOS_VENDIDOS
								End If
								
							rsSubTotais.Close
							Set rsSubTotais = Nothing
						
						Else
						
							SUB_TOTAL_QTD_AGREGADOS  = "-"
							SUB_TOTAL_QTD_AGREGADOS_VENDIDOS = "-"
						
						End If
						
						
						'////////////////////				
						
						TOTAL_FOTOS_FATURADAS = 0
						TOTAL_FOTOS_INICIAL = 0
						TOTAL_FOTOS_APROVEITAMENTO = 0

						Set rs222=objConexao.Execute("SELECT sum(INICIAL) AS INICIAL,SUM(FATURADO) AS FATURADO FROM VW_PEDIDOS_COMPLETO WHERE VW_PEDIDOS_COMPLETO.FINALIZA_VENDEDOR = 1 AND COD_PRODUCAO = '"&COD_PRODUCAO&"' and COD_VENDEDOR = '"&rs112("COD_VENDEDOR")&"'")
							
							If not rs222.Eof Then
								
								TOTAL_FOTOS_FATURADAS = formatNumber(0&rs222("FATURADO"),0)
								TOTAL_FOTOS_INICIAL = formatNumber(0&rs222("INICIAL"),0)
								
								If int(TOTAL_FOTOS_FATURADAS) > 0 Then
									TOTAL_FOTOS_APROVEITAMENTO = formatNumber((100*TOTAL_FOTOS_FATURADAS)/TOTAL_FOTOS_INICIAL,1)
								End If
							End If
							
						rs222.Close
						Set rs222=Nothing
		
						
						Set rs222=objConexao.Execute("SELECT sum(INICIAL) AS INICIAL,SUM(FATURADO) AS FATURADO FROM VW_PEDIDOS_COMPLETO WHERE COD_PRODUCAO = '"&COD_PRODUCAO&"' and COD_VENDEDOR = '"&rs112("COD_VENDEDOR")&"' AND FINALIZA_VENDEDOR = 1")
							
							If not rs222.Eof Then
								
								TOTAL_ARRECADO_PRODUCAO = formatNumber(0&rs222("FATURADO"),0)
							End If
							
						rs222.Close
						Set rs222=Nothing
		
		'/////////////////////////////////////////////
		'Total arrecadado
				TOTAL_PAGAMENTO_EM_DINHEIRO = 0
				TOTAL_XTOTAL_CHEQUE = 0
				TOTAL_XTOTAL_CARTAO = 0
				TOTAL = 0
				
				Set rs5=objConexao.Execute("SELECT SUM(TB_CACHE_VW_PEDIDOS_TOTAL_RECEBIDO.PAGAMENTO_EM_DINHEIRO) AS PAGAMENTO_EM_DINHEIRO, SUM(TB_CACHE_VW_PEDIDOS_TOTAL_RECEBIDO.XTOTAL_CHEQUE) AS XTOTAL_CHEQUE, SUM(TB_CACHE_VW_PEDIDOS_TOTAL_RECEBIDO.XTOTAL_CARTAO) AS XTOTAL_CARTAO FROM VW_PEDIDOS_COMPLETO INNER JOIN TB_CACHE_VW_PEDIDOS_TOTAL_RECEBIDO ON VW_PEDIDOS_COMPLETO.COD_PEDIDO = TB_CACHE_VW_PEDIDOS_TOTAL_RECEBIDO.COD_PEDIDO WHERE VW_PEDIDOS_COMPLETO.COD_PRODUCAO = "&COD_PRODUCAO&" and VW_PEDIDOS_COMPLETO.COD_VENDEDOR = '"&rs112("COD_VENDEDOR")&"' AND VW_PEDIDOS_COMPLETO.FINALIZA_VENDEDOR = 1")
				
				If not rs5.Eof Then
					
					TOTAL_PAGAMENTO_EM_DINHEIRO = formatNumber(0&rs5("PAGAMENTO_EM_DINHEIRO"),2)/1
					TOTAL_XTOTAL_CHEQUE = formatNumber(0&rs5("XTOTAL_CHEQUE"),2)/1
					TOTAL_XTOTAL_CARTAO = formatNumber(0&rs5("XTOTAL_CARTAO"),2)/1
					TOTAL = (TOTAL_PAGAMENTO_EM_DINHEIRO+TOTAL_XTOTAL_CHEQUE+TOTAL_XTOTAL_CARTAO)/1
					
				End If
				
			rs5.Close
			Set rs5=Nothing	
			
			PEDIDOS = ""
			
			'Pedidos mostrar origem
			Set rs5=objConexao.Execute("SELECT COD_PEDIDO FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA WHERE COD_PRODUCAO = "&COD_PRODUCAO&" and COD_ULTIMO_VENDEDOR_FICHA = '"&rs112("COD_VENDEDOR")&"' AND COD_PEDIDO > 0 AND COD_USUARIO = "&Session("COD_USUARIO"))
				
				If not rs5.Eof Then
					While not rs5.Eof
					PEDIDOS = PEDIDOS&", "&rs5(0)
					rs5.MoveNext()
					Wend
				End If
				
			rs5.Close
			Set rs5=Nothing			
						
			If PEDIDOS = "" Then
				PEDIDOS = "Não houve pedidos deste vendedor"
			Else
				PEDIDOS = Mid(PEDIDOS,3,Len(PEDIDOS))
			End If			
						
				%>
				<tr>
                          <td>
                            <div align="right">
                            <strong><a href="javascript:alert('Pedidos:\n<%=PEDIDOS%>')"><%=Ucase(rs112("NOME_VENDEDOR"))%></a></strong></div></td>
                          <td colspan="5">
                            <div align="center">
                             <%
							Set rs8=objConexao.Execute("SELECT LOTE, DATA_SAIDA_INICIAL FROM TB_PRODUCAO_LOTES WHERE COD_PRODUCAO = "&COD_PRODUCAO&" AND COD_VENDEDOR_INICIAL = '"&rs112("COD_VENDEDOR")&"' order by DATA_SAIDA_INICIAL asc")
							
							If not rs8.Eof Then%>
                            <table width="100" border="0" cellspacing="0" cellpadding="2">
                          <tr>
                            <td><div align="center"><strong>Lote</strong></div></td>
                            <td><div align="center"><strong>Saída</strong></div></td>
                          </tr>

							

                            <%
							
								While not rs8.Eof
							%>
                           
                            <tr>
                            <td><div align="center"><%=rs8("LOTE")%></div></td>
                            <td><%=ArrumaData(rs8("DATA_SAIDA_INICIAL"))%></td>
                          </tr>
                            
                            <%
								rs8.MoveNext()
								Wend
							%>
							</table>
							<%
							Else%>
							-
                            
                            <%
							End If
							
							
							rs8.Close
							Set rs8=Nothing
							%>
                            </div></td>
                          <td><div align="center"><%=SUB_TOTAL_QTD_ALBUNS%></div></td>
                          <td><div align="center"><%=SUB_TOTAL_QTD_ALBUNS_VENDIDOS%></div></td>
                          <td><div align="center"><%=SUB_TOTAL_QTD_ALBUNS-SUB_TOTAL_QTD_ALBUNS_VENDIDOS%></div></td>
                          <td><div align="center">
                            <%
				  			SUB_APROVEITAMENTO_ALBUNS = formatNumber(100-((SUB_TOTAL_QTD_ALBUNS-SUB_TOTAL_QTD_ALBUNS_VENDIDOS)*100/(SUB_TOTAL_QTD_ALBUNS_VENDIDOS+(SUB_TOTAL_QTD_ALBUNS-SUB_TOTAL_QTD_ALBUNS_VENDIDOS))),0)
							Response.Write(SUB_APROVEITAMENTO_ALBUNS)
							%>%</div></td>
                          <td>                  <%
				  If int(QTD_PRODUTOS) > 1 Then
				  %>
                  <table style="border-collapse:collapse; font-size:10px" width="100%" 
		      border="1" align="center"
					cellpadding="2" 
					>
                    
                    <tr>
                  	<%For i=0 to QTD_PRODUTOS-1
					%>
                   
                      <td>
                     <div align="center" style="visibility:hidden; height:1px"><%=PRODUTOS(i,1)%></div>
                      <div align="center">
                      <%					  
					  Set rs4=objConexao.Execute("SELECT TB_PEDIDOS_PRODUTOS.COD_PRODUTO, TB_PRODUTOS.NOME_PRODUTO, SUM( TB_PEDIDOS_PRODUTOS.QUANTIDADE ) AS QUANTIDADE FROM TB_PEDIDOS INNER JOIN TB_PEDIDOS_PRODUTOS ON ( TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO )INNER JOIN TB_PRODUTOS ON ( TB_PEDIDOS_PRODUTOS.COD_PRODUTO = TB_PRODUTOS.COD_PRODUTO ) INNER JOIN TB_FICHAS_IDENTIFICACAO ON ( TB_PEDIDOS_PRODUTOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO ) WHERE TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO ="&COD_PRODUCAO&" AND TB_PEDIDOS.FINALIZA_VENDEDOR = 1 AND TB_PEDIDOS_PRODUTOS.COD_PRODUTO = "&PRODUTOS(i,0)&" AND TB_PEDIDOS.COD_VENDEDOR = "&rs112("COD_VENDEDOR")&" GROUP BY TB_PEDIDOS_PRODUTOS.COD_PRODUTO")
					  	
						If not rs4.Eof Then
							Response.Write(rs4("QUANTIDADE"))
						Else
							Response.Write("&nbsp;")
						End If
						
					  rs4.Close
					  Set rs4=Nothing
					  
					  %>
					  </div></td>
                    <%
					Next
					%>
                    </tr>
                  </table>
                  <%End If%>&nbsp;</td>
				          <td><div align="right"><%=formatNumber(TOTAL,2)%></div></td>
				          <%
				rs112.MoveNext()
				Wend
				
				
				rs112.Close
				Set rs112 = Nothing
				%> 
				        </table>
		
		  <p>* - O valor de arrecadação de produção Não é 100% exato, pois há alguns pedidos que possuem mais que um formando e de diferentes produções no mesmo pedido, nestes casos o sistema Não tem como desmembrar o valor por formando e por tanto considerado o valor total. </p>
		  <p>* - Os dados utilizados para gerar este relatório tem como base as informações registradas até as 8:00 AM da data de hoje. </p>
		  <p>Observa&ccedil;&atilde;o: Clique no nome do vendedor para detalhar o n&uacute;mero dos pedidos.</p>
		  <%Else%>
		   <div align="center">
           <br /><br /><br /><br />
		  <p><span class="textoMedio"><strong>Não há dados para os filtros selecionados!</strong></span></p>
          <br /><br /><br /><br />
		</div>
		    <%End If
			
			rs.Close
			Set rs = Nothing
			
			
			'Limpa novamente a TB tempor�ria
			'objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA WHERE COD_USUARIO ="&Session("COD_USUARIO"))
			FecharConexao()
			
			%>
		   
		    

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






