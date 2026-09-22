<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%




'EXCLUS�O DE FECHAMENTO COMPLETO
	 If Cstr(inStr(sub_validacoes_paginas,"#42")) <> "0" and Request.QueryString("comando") = "excluir" Then
		
		AbrirConexao()
		
		'1) Deleta o cabe�alho		
		sql = "DELETE FROM TB_FECHAMENTO_PAGAMENTO_CABECALHO WHERE VENDEDOR_RECEBE_EM = "&sqlData(Request.QueryString("VENDEDOR_RECEBE_EM"))
		
		'Response.Write(sql&"<br />")
		objConexao.Execute(sql)
		
		
		'2) Deleta Tabela de Fechamento	
		sql = "DELETE FROM TB_FECHAMENTO_PAGAMENTO WHERE VENDEDOR_RECEBE_EM = "&sqlData(Request.QueryString("VENDEDOR_RECEBE_EM"))
		
		'Response.Write(sql&"<br />")
		objConexao.Execute(sql)
		
		'3) Update em Pedidos como n�o processados
		sql = "UPDATE TB_PEDIDOS SET PAGAMENTO_PROCESSADO = 2 WHERE VENDEDOR_RECEBE_EM = "&sqlData(Request.QueryString("VENDEDOR_RECEBE_EM"))
		
		
		'Response.Write(sql&"<br />")
		objConexao.Execute(sql)
		
		'4) Ciclos de fechamentos
		sql = "UPDATE TB_FECHAMENTO_PAGAMENTO_PERIODOS SET data_processado_pagamento = NULL WHERE data_pagamento = "&sqlData(Request.QueryString("VENDEDOR_RECEBE_EM"))

		'Response.Write(sql&"<br />")
		objConexao.Execute(sql)



		objConexao.Execute("DELETE FROM TB_FECHAMENTO_PAGAMENTO_BOLETOS WHERE vendedor_recebe_em = "&sqlData(Request.QueryString("VENDEDOR_RECEBE_EM")) )
		objConexao.Execute("UPDATE TB_PEDIDOS_BOLETOS SET cod_fechamento_periodo = Null , pagamento_processado = 0 WHERE vendedor_recebe_em = "&sqlData(Request.QueryString("VENDEDOR_RECEBE_EM")) )

		FecharConexao()
		
		Response.Redirect("cap_fechamentos.asp?codmsg=3")
		
	 End If
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>


<title>Pagamento de Vendedores</title>
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
 


	<!--AJUDA F1 E CTRL PARA BUSCA RÁPIDA-->
	</script>
    
    <!--CONFIRMA��O SIM/N�O-->
<script language="javascript">

function confirmaExclusaoFechamento(vendedor_recebe_em,mes){
	if(YNconfirm('ATEN��O!!!!!!!!!!!!!!!!\nTEM CERTEZA QUE DESEJA EXCLUIR ESTE FECHAMENTO DE '+mes+'?'))
	{
		abre('cap_fechamentos.asp?comando=excluir&vendedor_recebe_em='+vendedor_recebe_em);
	}
}



function YNconfirm(txt){
return typeof(suporteVBscript)=="undefined"?confirm(txt):VBconfirm(txt)==6
}
</script>
<script language="vbscript">
suporteVBscript=1
function VBconfirm(mensagem)

 VBconfirm=msgbox(mensagem,vbyesno)
end function

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
<style type="text/css">
<!--
.menu_paginacao{
	float: left;
	width: 300px;
	position: relative;
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
	color: #990000;
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
  


	  <div style="width: 100%; text-align: center; margin-bottom: 15px;">
  <div style="background-color: #ff007f; color: white; padding: 10px 20px; border-radius: 8px; display: inline-flex; align-items: center; gap: 15px;">

    <strong style="font-size: 16px;">CAP - Pagamento de Vendedores</strong>

    <%
    AbrirConexao()
    Dim cod_fechamento_periodo, data_pagamento_nao_processado
    Set rs2=objConexao.Execute("SELECT cod_fechamento_periodo, data_pagamento FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE data_processado_pagamento IS NULL ORDER BY data_pagamento ASC ")
    
    If not rs2.Eof Then
        cod_fechamento_periodo = rs2(0)
        data_pagamento_nao_processado = formatDateTime(rs2(1))
    End If
    
    rs2.Close
    Set rs2=Nothing
    %>

    <label for="subMenuSelect" style="color: white; font-weight: bold;">| Sub-Menu:</label>

<select id="subMenuSelect" onchange="handleSubMenuChangeCustom(this)" 
        style="padding: 8px 12px; border-radius: 6px; border: 1px solid #ccc; background-color: white; color: #ff007f; font-weight: bold; cursor: pointer; min-width: 220px;">
    
    <option value="">Selecione uma opção</option>
    <option value="adm_rotina_fechamento_extrato_pagamento.asp?cod_fechamento_periodo=<%=cod_fechamento_periodo%>">
        Gerar novo Fechamento (<%=ArrumaData(data_pagamento_nao_processado)%>)
    </option>

    <option value="cap_fechamentos_periodos.asp">Períodos de Vendas</option>
    <option value="cap_fechamentos_tabelas_comissao_vendedores.asp">Tabela de Comissionamento de Vendedores</option>
    <option value="cap_fechamentos_comissionamento.asp">Tabela de Comissionamento de Coordenadores</option>
    <option value="cap_fechamentos_pendencias.asp">Pendências de Pagamento</option>
    <option value="ranking">Ranking de vendedores</option>
</select>

<!-- POPUP DE PERÍODO -->
<div id="popupRanking" style="display:none; position: fixed; top: 0; left: 0; width: 100%; height: 100%;
    background-color: rgba(0,0,0,0.6); z-index: 9999; justify-content: center; align-items: center;">
    <div style="background-color: white; padding: 25px; border-radius: 10px; width: 350px; text-align: center; position: relative;">
        <h3 style="color: #ff007f;">Selecione o Período</h3>
        <div style="margin-top: 10px;">
            <label>Data Início:</label><br>
            <input type="date" id="dataInicio" style="padding: 6px; margin-bottom: 10px; width: 80%;">
        </div>
        <div>
            <label>Data Fim:</label><br>
            <input type="date" id="dataFim" style="padding: 6px; margin-bottom: 15px; width: 80%;">
        </div>
        <div style="margin-top: 15px;">
            <button onclick="confirmarRanking()" style="background-color:#ff007f; color:white; border:none; border-radius:6px; padding:8px 15px; cursor:pointer;">Confirmar</button>
            <button onclick="fecharPopupRanking()" style="background-color:#ccc; color:black; border:none; border-radius:6px; padding:8px 15px; cursor:pointer; margin-left:10px;">Cancelar</button>
        </div>
    </div>
</div>

<script>
function handleSubMenuChangeCustom(select) {
    const valor = select.value;
    if (valor === "ranking") {
        document.getElementById("popupRanking").style.display = "flex";
        select.value = ""; // limpa a seleção
    } else if (valor) {
        // redirecionamentos normais
        if (valor.includes("adm_rotina_fechamento_extrato_pagamento")) {
            window.open(valor, "_blank");
        } else {
            window.location.href = valor;
        }
    }
}

function fecharPopupRanking() {
    document.getElementById("popupRanking").style.display = "none";
}

function confirmarRanking() {
    const inicio = document.getElementById("dataInicio").value;
    const fim = document.getElementById("dataFim").value;

    if (inicio === "" || fim === "") {
        alert("Por favor, selecione o período completo (início e fim).");
        return;
    }

    // redireciona com base nas datas
    window.location.href = `adm_relatorios_ranking_vendedores_completo.asp?data_venda_inicio=${inicio}&data_venda_fim=${fim}`;
}
</script>


  </div>
</div>

<script>
function handleSubMenuChange(select) {
    const url = select.value;
    if (url) {
        if (url.includes("adm_rotina_fechamento_extrato_pagamento")) {
            window.open(url, "_blank");
        } else {
            window.location.href = url;
        }
    }
}
</script>


</div>
	   
       <table style="border:white; padding:2px; background:#ff007f" width="99%" border="0" align="center" cellpadding="5" cellspacing="0">
	
		  <tr>
		    <td width="50%"><div align="right"><strong style="font-size: 16px; color: white">Relat&oacute;rio de Previs&atilde;o de Folha de Pagamento:</strong><img onclick="alert('Faz um c�lculo at� o momento da Folha de Pagamento dos per�odos de vendas que n�o foram processados')"  /></div></td>
		    <td width="81%">
				<form name="EXTRATO_APROVEITAMENTO" action="cap_fechamentos_IMPRESSAO_folha_geral.asp" target="_blank" method="get" style="margin:0px"> 
				<table width="340" border="0" cellspacing="0" cellpadding="3">
              <td width="70%">
  				<div style="text-align: center;">
    				<select name="cod_fechamento_periodo" class="inputTamanho2" style="padding: 8px; border-radius: 6px; border: 1px solid #ccc;">
				  <%
		AbrirConexao()
		Set rs2=objConexao.Execute("SELECT cod_fechamento_periodo, data_pagamento FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE data_processado_pagamento IS NULL ORDER BY data_pagamento ASC ")
								
		If not rs2.Eof Then
				While not rs2.Eof
					%>
					<option <%if int(cod_fechamento_periodo) = int(rs2(0)) Then%> selected="selected" <%End If%> value="<%=rs2(0)%>"><%=ArrumaData(rs2(1))%></option>
					<%
				rs2.MoveNext()
				Wend	
		End If
								
	rs2.Close
	Set rs2=Nothing
	FecharConexao()
	%>  
     </select>
                  </div></td>
                <td width="30%"><div align="center">
                  <button type="submit" style="
					background-color: white; 
					color: #ff007f ; 
					padding: 8px 16px; 
					border: none; 
					border-radius: 6px; 
					cursor: pointer; 
					font-weight: bold;">
    				BUSCAR
					</button>

                </div></td>
              </tr>
            </table>
				</form>			</td>
		    </tr>
        </table>
       <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
       <%
	Dim sucesso, codmsg
	codmsg	= CStr(""&Request.QueryString("codmsg"))
	sucesso = CStr(""&Request.QueryString("sucesso"))
		
	If sucesso = "" and codmsg <> "" Then%>
    	<div align="center" class="msgGrandeSucesso">
        	<%
			If codmsg = "1" Then
				Response.Write("Fechamento processado com sucesso, imprima os recibos!")
			Elseif codmsg = "2" Then

			Elseif codmsg = "3" Then	
				Response.Write("Fechamento exclu�do com sucesso!")
			End If	
			%>
            
        
        </div>
    <%ElseIf sucesso <> "" and codmsg <> "" Then%>
        <div align="center" class="msgGrandeErro">
			<%
			If codmsg = "41" Then
				Response.Write("Voc� est� tentando alterar/excluir uma pend�ncia de pagamento cujo pagamento do m�s de refer�ncia j� foi conclu�do e por tanto a pend�ncia n�o seria realizada.")	
			End If	
			%>
</div>
	<%End If%>
	  	  <%
		  AbrirConexao()
		  
		  
		  '////////////////////////////////////////////////////
		  'Valida��o de relat�rios
		  
		  Dim acesso_1, acesso_2, acesso_3, acesso_4
		  acesso_1 = false
		  acesso_2 = false
		  acesso_3 = false
		  acesso_4 = false
		  
		  Set rs=objConexao.Execute("SELECT * FROM TB_ADM_PERFIS_PAGINAS_HABILITADAS WHERE COD_PAGINA in(163,164,168,165) and COD_PERFIL = '"&Session("cod_perfil")&"'")
		  		
				While not rs.Eof
					
					If int(rs("COD_PAGINA")) = 163 Then
					'Impress�o de Extrato de Coordenadores
						acesso_1 = true
					Elseif int(rs("COD_PAGINA")) = 164 Then
					'Impress�o de Extrato de Vendedores
						acesso_2 = true
					Elseif int(rs("COD_PAGINA")) = 168 Then
					'Impress�o do Resumo da Folha de Pagamento
						acesso_3 = true
					Elseif int(rs("COD_PAGINA")) = 165 Then
					'Impress�o de Recibo
						acesso_4 = true
					End If
				
				rs.MoveNext()
				Wend
							  
		  rs.Close
		  Set rs=Nothing
		  '////////////////////////////////////////////////////
		  
		  Dim rs, sql
		  sql = "SELECT VENDEDOR_RECEBE_EM,DATA_PROCESSAMENTO, count(*) as QTD_PEDIDOS, COD_FECHAMENTO_PERIODO FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_FECHAMENTO_PERIODO > 0 group by VENDEDOR_RECEBE_EM order by VENDEDOR_RECEBE_EM desc"  
		  
		  
		  	Set rs=server.CreateObject("ADODB.recordset")
			rs.CursorLocation=3  'Cursor no cliente
			rs.PageSize=12
			rs.Open sql, objConexao		  
		  
		  If not rs.Eof then
		  
		  	Dim PAGINA_ATUAL
			PAGINA_ATUAL = formatNumber(0&Request.QueryString("pagina"),0)
			
			If PAGINA_ATUAL > 0 Then
					rs.AbsolutePage=PAGINA_ATUAL
			Else
				PAGINA_ATUAL = 1
				rs.AbsolutePage=1
			End If	
		  
		  %>
		  <h3 align="center">Rela&ccedil;&atilde;o de Fechamentos</h3>
		  <div id="comboTexto">
		  <table width="95%" border="0" align="center">
  <tr>
    <td width="47%"> 
P&aacute;gina <strong><%=rs.AbsolutePage%></strong> de <%=rs.PageCount%><br /></td>
    <td width="53%"><div align="right">Total <strong><%=rs.RecordCount%></strong> fechamentos</div></td>
  </tr>
</table>
		  
          <table width="95%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="18%" rowspan="2"><div align="center"><strong>Pagamento de</strong></div></td>
            <td width="18%" rowspan="2"><div align="center"><strong>Data do Processamento</strong> / <strong>Gerado por</strong></div></td>
            <td width="10%" rowspan="2"><div align="center"><strong>Qtd. Pedidos</strong></div></td>
            <td colspan="4"><div align="center"><strong>Recibos</strong></div></td>
            <td width="5%" rowspan="2"><p>&nbsp;</p>
              <p>&nbsp;</p></td>
</tr>
<tr bgcolor="whitesmoke">
  <td width="12%"><div align="center"><strong>1) Recibos</strong></div></td>
  <td width="20%"><div align="center"><strong>2) Extrato Vendedores</strong></div></td>
  <td width="10%"><div align="center"><strong>3) Folha Geral</strong></div></td>
  <td width="7%"><div align="center"><strong>4) Comissão por boleto</strong></div></td>
  </tr>
		  		 <%
	 
				 
				  
		 Dim x, primeiro
		 primeiro = true
		 
		 For x=1 To rs.PageSize
		    If Not rs.Eof Then
			
				Set rs2=objConexao.Execute("SELECT TB_FECHAMENTO_PAGAMENTO_CABECALHO.*, TB_ADM_USUARIOS.NOME FROM TB_ADM_USUARIOS INNER JOIN TB_FECHAMENTO_PAGAMENTO_CABECALHO ON TB_ADM_USUARIOS.COD_USUARIO = TB_FECHAMENTO_PAGAMENTO_CABECALHO.PAGAMENTO_GERADO_POR WHERE VENDEDOR_RECEBE_EM = "&sqlData(rs(0)))
				
					If not rs2.Eof Then
						
						DIA_VENCIMENTO = rs2("DIA_VENCIMENTO")
						PAGAMENTO_GERADO_POR = rs2("NOME")
						
					End If
				
				rs2.Close			
				Set rs2=Nothing
				%>
				<tr>
							<td class="textoMedio"><%=ArrumaData(rs("VENDEDOR_RECEBE_EM"))%></td>
							<td><div align="center"><%=ArrumaData(rs("DATA_PROCESSAMENTO"))%> / <%=PAGAMENTO_GERADO_POR%>
                            </div></td>
							<td><div align="center"><%=rs("QTD_PEDIDOS")%></div></td>
							<td><div align="center">
							<%If acesso_4 Then%>
                            	<a target="_blank" href="adm_rotina_fechamento_extrato_pagamento_part3_IMPRESSAO_recibo_extenso_agrupado.asp?COD_FECHAMENTO_PERIODO=<%=rs("COD_FECHAMENTO_PERIODO")%>">Recibos</a>
                            <%Else
							Response.Write("Acesso não permitido")
							
							End If
							%>
                                </div></td>
				            <td><div align="center">
				              <%If acesso_2 Then%>
				              <a target="_blank" href="adm_rotina_fechamento_extrato_pagamento_part3_IMPRESSAO_detalhamento_vendedores.asp?COD_FECHAMENTO_PERIODO=<%=rs("COD_FECHAMENTO_PERIODO")%>">Extrato Vendedores/Coordenadores</a>
				              <%Else
							Response.Write("Acesso não permitido")
							
							End If
							%>
				              </div></td>
				            <td><div align="center">
				              <%If acesso_3 Then%>
				              <a target="_blank" href="adm_rotina_fechamento_extrato_pagamento_part3_IMPRESSAO_folha_geral.asp?COD_FECHAMENTO_PERIODO=<%=rs("COD_FECHAMENTO_PERIODO")%>">Folha Geral</a><%Else
							Response.Write("Acesso não permitido")
							
							End If
							%></div></td>
				           <td>
							<div align="center">
								<% 
								' --- Mostra o link apenas para usuários autorizados ---
								 If Session("COD_USUARIO") = 57 Or Session("COD_USUARIO") = 16 Or Session("COD_USUARIO") = 22 Or Session("COD_USUARIO") = 42 Or Session("COD_USUARIO") = 53 Or Session("COD_USUARIO") = 1 Then
									Dim rsBoleto, temBoleto
									Set rsBoleto = objConexao.Execute("SELECT SUM(TOTAL_BOLETO) as TOTAL FROM TB_FECHAMENTO_PAGAMENTO WHERE VENDEDOR_RECEBE_EM = " & sqlData(rs("VENDEDOR_RECEBE_EM")) & " AND TOTAL_BOLETO > 0")
									temBoleto = False
									If Not rsBoleto.Eof Then
										If Not IsNull(rsBoleto("TOTAL")) Then
											If CDbl(rsBoleto("TOTAL")) > 0 Then temBoleto = True
										End If
									End If
									rsBoleto.Close
									Set rsBoleto = Nothing
									If temBoleto Then
								%>
									<a href="adm_controle_no_extrato_comissao_vendedores_boletos.asp?COD_FECHAMENTO_PERIODO=<%=rs("COD_FECHAMENTO_PERIODO")%>" 
									target="_blank">Boletos</a>
								<% Else %>
									<span style="color:#999999; font-size:11px;">- Sem pagamentos de boleto</span>
								<% 
									End If
							End If	
								%>
							</div>
							</td>
				            <td><div align="center">
                           <%
						   If Cstr(inStr(sub_validacoes_paginas,"#42")) <> "0" AND primeiro Then
						   		primeiro = false
						   %>
                            <a href="javascript:confirmaExclusaoFechamento('<%=rs(0)%>','<%=Ucase(MonthName(Month(rs("VENDEDOR_RECEBE_EM"))))%> / <%=Year(rs("VENDEDOR_RECEBE_EM"))%>')"><img src="images/icons/icon_excluir.png" width="16" height="16" border="0" /></a>	
                            <%End If%>&nbsp;
                            </div></td>
				</tr>
		        <%
				rs.MoveNext
				End If 
				Next	
      
		  %>
        </table>
		
		
		      <div align="left">
		        <%If rs.PageCount > 1 Then %>
					<br />
					<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td>
						<div class="menu_paginacao">
					Ir para a P&aacute;gina:
					
					<%
					Dim aae
					aae = 0
					For x=1 To rs.PageCount 
					aae = formatNumber(0&x,0)
					
						If int(PAGINA_ATUAL) <> int(aae) Then
						'Response.Write(PAGINA_ATUAL&":"&x&"<br />")
						%>
							<a href="?pagina=<%=x%>"><%=x%></a>
						<%Else%>
							<%=x%>
						<%End If%>
						
					<%Next%>
					</div>
					</td>
                      </tr>
                    </table>
					
					
					

				
				<%End if
		
		Else
		
		End If      
		
		
 		rs.Close()
		Set rs = Nothing
		FecharConexao()
		%>
		  <br /><br />
          <table width="95%" border="0" align="center" cellpadding="5" cellspacing="0" style="border:solid 1px #CCCCCC">
  <tr>
    <td bgcolor="#F4F4F4"><div align="center"><strong>Instru&ccedil;&otilde;es de Impress&atilde;o de Recibos</strong></div></td>
    </tr>
  <tr>
    <td><strong>1) Recibos</strong><br />
      <br />
      A impress&atilde;o &eacute; realizada pelo pr&oacute;prio navegador, mas antes configure a margem, cabe&ccedil;alho e rodap&eacute; para impress&atilde;o correta.<br />
      Para configurar no seu navegador v&aacute; at&eacute; o menu e clique em Arquivo / Configurar P&aacute;gina e deixe as op&ccedil;&atilde;es cabe&ccedil;alho e rodap&eacute; em branco e altere as margens para 1,5 cm em todos os lados.</td>
    </tr>
  <tr>
    <td><p><strong>2) Extrato de Vendedores</strong></p>
      <p>A impress&atilde;o dos extratos dependem do auxilio do Microsoft Word e um modelo b&aacute;sico de impress&atilde;o. Para baixar o modelo <a href="documentos/modelo_extrato_pagamento.doc" target="_blank">clique aqui</a>.<br />
        Ap&oacute;s gerar o recibo pressione as teclas CTRL + A (para selecionar todo o conte&uacute;do) e depois em CTRL + C (para copiar o conte&uacute;do).<br />Depois basta colar o conte&uacute;do no modelo de recibo de Vendedores e depois quebrar as p&aacute;ginas manualmente na divis&atilde;o de vendedores e onde achar conveniente.</p>
      <p>Aten&ccedil;&atilde;o utilize somente o navegador Internet Explorer para gerar os Extratos devido a compatibilidade com o Word.</p></td>
  </tr>
  <tr>
    <td><p><strong>3) Folha Geral</strong></p>
      <p>A impress&atilde;o &eacute; realizada pelo pr&oacute;prio navegador, mas antes configure a margem, cabe&ccedil;alho e rodap&eacute; para impress&atilde;o correta.<br />
        Para configurar no seu navegador v&aacute; at&eacute; o menu e clique em Arquivo / Configurar P&aacute;gina e deixe as op&ccedil;&otilde;es cabe&ccedil;alho e rodap&eacute; em branco e altere as margens para 1,5 cm em todos os lados.</p></td>
  </tr>
</table>

          
                    
		      </div>
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

		<script>
		function toggleSubMenu() {
			const menu = document.getElementById('subMenuDropdown');
			if (menu.style.display === 'none' || menu.style.display === '') {
				menu.style.display = 'block';
			} else {
				menu.style.display = 'none';
			}
		}
		</script>

</body>
</html>