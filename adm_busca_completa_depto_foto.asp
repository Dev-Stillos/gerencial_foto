<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%


'////////////////////////////////////////////////////////////////////
'Valida��o extra
BUSCA_35	= False 
BUSCA_36	= False 
BUSCA_37	= False 
BUSCA_38	= False 
BUSCA_39	= False 
BUSCA_40	= False 
BUSCA_90	= False

AbrirConexao()
Set rs = objConexao.Execute("SELECT COD_PERFIL, COD_SUB_VALIDACAO FROM TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES WHERE COD_PERFIL="&Session("COD_PERFIL")&" order by COD_SUB_VALIDACAO asc")
		
		While not rs.Eof
						
			If int(rs("COD_SUB_VALIDACAO")) = 35 Then
			'Busca por C�d. Pedido
				BUSCA_35 = True				
			End If	
			
			If int(rs("COD_SUB_VALIDACAO")) = 36 Then
			'Busca por Nome do Formando
				BUSCA_36 = True
			End If	
			
			If int(rs("COD_SUB_VALIDACAO")) = 37 Then
			'Busca por C�d. Formando Ficha
				BUSCA_37 = True
			End If	
			
			If int(rs("COD_SUB_VALIDACAO")) = 38 Then
			'Busca por N�mero do �lbum
				BUSCA_38 = True
			End If	
			
			If int(rs("COD_SUB_VALIDACAO")) = 39 Then
			'Busca por Contrato
				BUSCA_39 = True
			End If	

			If int(rs("COD_SUB_VALIDACAO")) = 40 Then
			'Busca por C�d. da Produ��o
				BUSCA_40 = True
			End If					
					
			If int(rs("COD_SUB_VALIDACAO")) = 90 Then
			'Busca por CPF
				BUSCA_90 = True
			End If	
		rs.MoveNext()
		Wend
		
rs.Close
Set rs=Nothing
FecharConexao()
'////////////////////////////////////////////////////////////////////






 Dim BUSCAR_TERMO, TIPO_BUSCA, SQL_TIPO, BUSCA_ALBUM
			BUSCAR_TERMO = TrocarAcento(""&Request.QueryString("BUSCAR_TERMO"))
			BUSCAR_TERMO = Replace(BUSCAR_TERMO,"-","")
			TIPO_BUSCA	= formatNumber(0&Request.QueryString("TIPO_BUSCA"),0)
			
BUSCA_ALBUM = 1			
		  
		  'Tipos de Busca
		  '1 - C�D. PEDIDO
		  '2 - NOME FORMANDO
		  '3 - N�M. �LBUM
		  '4 - NOME FACULDADE
		  '5 - CONTRATO
		  '6 - NOME PRODU��O
		  '7 - C�D. PRODU��O
		  '8 - CPF
		  
		  'SQL_TIPO
		  '1 - TB_PEDIDOS
		  '2 - TB_FICHAS_IDENTIFICACAO
		  
		  If TIPO_BUSCA = 0  Then
		  'Buscas iniciais	 
		 		  
			 If isNumeric(BUSCAR_TERMO) Then
			 	'Padr�o buscar pelo c�digo da venda
				TIPO_BUSCA = 1
			 Else
			 	'Padr�o buscar pelo nome do formando
				TIPO_BUSCA = 2
			 End If
			 
		  End If
		  
		  		If TIPO_BUSCA = 1 Then
				'1 - C�D. PEDIDO
				
					If not BUSCA_35 Then
					'não possui acesso a este tipo de busca
						Response.Redirect("adm_busca_completa_depto_foto.asp?msg=Você não possui acesso a este tipo de busca!")	
					End If
				
					If isNumeric(BUSCAR_TERMO) Then
					
					sql = "SELECT * FROM TB_PEDIDOS WHERE COD_PEDIDO = "&BUSCAR_TERMO&" AND FINALIZA_VENDEDOR=1"
					SQL_TIPO = 1
					Else
						Response.Redirect("adm_busca_completa_depto_foto.asp?msg=Busca incorreta!")						
					End If
					
				Elseif  TIPO_BUSCA = 2 Then
				'2 - NOME FORMANDO
				
					If not BUSCA_36 Then
					'não possui acesso a este tipo de busca
						Response.Redirect("adm_busca_completa_depto_foto.asp?msg=Você não possui acesso a este tipo de busca!")	
					End If
				
		  			sql = "SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE NOME_FORMANDO LIKE '%"&BUSCAR_TERMO&"%' ORDER BY NOME_FORMANDO"
					SQL_TIPO = 2
					
		  		Elseif  TIPO_BUSCA = 3 Then
				'3 - N�M. �LBUM
				
					If not BUSCA_38 Then
					'não possui acesso a este tipo de busca
						Response.Redirect("adm_busca_completa_depto_foto.asp?msg=Você não possui acesso a este tipo de busca!")	
					End If
					
					If isNumeric(BUSCAR_TERMO) Then
					
						Dim ANO_CONCLUSAO, SEMESTRE_CONCLUSAO
						ANO_CONCLUSAO		= ""&Request.QueryString("ANO_CONCLUSAO")
						SEMESTRE_CONCLUSAO	= ""&Request.QueryString("SEMESTRE_CONCLUSAO")
					
						If ANO_CONCLUSAO <> "" and SEMESTRE_CONCLUSAO <> "" Then
							sql = "SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE ALBUM = "&BUSCAR_TERMO&" AND SEMESTRE_CONCLUSAO ='"&SEMESTRE_CONCLUSAO&"' AND ANO_CONCLUSAO ="&ANO_CONCLUSAO&" ORDER BY CONVERT(ALBUM,SIGNED)"
							SQL_TIPO = 2
						Else
							BUSCA_ALBUM = 2
						End If	
							
					Else
						Response.Redirect("adm_busca_completa_depto_foto.asp?msg=Busca incorreta!")
					End If
					
					
				Elseif  TIPO_BUSCA = 5 Then
				'5 - CONTRATO
					
					If not BUSCA_39 Then
					'não possui acesso a este tipo de busca
						Response.Redirect("adm_busca_completa_depto_foto.asp?msg=Você não possui acesso a este tipo de busca!")	
					End If
				
						sql = "SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE COD_CONTRATO = '"&BUSCAR_TERMO&"' ORDER BY CONVERT(ALBUM,SIGNED)"
						SQL_TIPO = 2	
				Elseif  TIPO_BUSCA = 7 Then
				'7 - C�D. PRODU��O
				
					If not BUSCA_40 Then
					'não possui acesso a este tipo de busca
						Response.Redirect("adm_busca_completa_depto_foto.asp?msg=Você não possui acesso a este tipo de busca!")	
					End If
				
					If isNumeric(BUSCAR_TERMO) Then
						sql = "SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = "&BUSCAR_TERMO
						SQL_TIPO = 2
					Else
						Response.Redirect("adm_busca_completa_depto_foto.asp?msg=Busca incorreta!")					
					End if
				Elseif  TIPO_BUSCA = 8 Then
				'7 - C�D. FORMANDO DA FICHA
				
					If not BUSCA_37 Then
					'não possui acesso a este tipo de busca
						Response.Redirect("adm_busca_completa_depto_foto.asp?msg=Você não possui acesso a este tipo de busca!")	
					End If
				
					If isNumeric(BUSCAR_TERMO) Then
				
					If Int(Right(BUSCAR_TERMO,1)) = 1 Then
					sql = "SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO = "&Mid(BUSCAR_TERMO,1,Len(BUSCAR_TERMO)-1)&"  ORDER BY NOME_FORMANDO"
					SQL_TIPO = 2
					Else
					Response.Redirect("adm_busca_completa_depto_foto.asp?msg=Atenção insira o Cód. Formando da Ficha (com 1 no final)!")
					End If
					
					Else
						Response.Redirect("adm_busca_completa_depto_foto.asp?msg=Busca incorreta!")					
					End if
				Elseif TIPO_BUSCA = 9 Then
				'8 - CPF

					' Verifica se o usuário tem permissão
					If not BUSCA_90 Then ' ou o número novo que você cadastrou lá em cima
						Response.Redirect("adm_busca_completa_depto_foto.asp?msg=Você não possui acesso a este tipo de busca!")	
					End If

					' Valida formato
					BUSCAR_TERMO = Replace(BUSCAR_TERMO, ".", "")
					BUSCAR_TERMO = Replace(BUSCAR_TERMO, "-", "")

					If isNumeric(BUSCAR_TERMO) Then
						sql = "SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE CPF = '" & BUSCAR_TERMO & "'"
						SQL_TIPO = 2
					Else
						Response.Redirect("adm_busca_completa_depto_foto.asp?msg=CPF inválido!")					
					End If

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
	<script language="javascript">
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
	<script language="javascript">
function validaForm(cod_formando,nomeFormando){
	d = document.CAD_FORMANDOS
	var acao
	acao = "dip_formandos_cadastro.asp?cod_formando="+cod_formando+"&cadastro=importado"
		var valida = 'Tem certeza que deseja cadastrar o formando(a) '+nomeFormando
		
			if(YNconfirm(valida)){
				d.action = acao;
				d.submit();	
			}
	}
	
	function confirmarExclusaoAgendamento(){
		if(YNconfirm('Tem certeza que deseja limpar as etiquetas agendadas pra impressão\nCertifique-se que foram impressão com sucesso antes de continuar.')){
			abre('?limpar_agendadas=sim')
		}
	
		
	}
	function selecionaBusca(){
	d = document.busca_sub_menu
		d.BUSCAR_TERMO.select();
	}
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
</script>
</head>

<body onload="selecionaBusca()">

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

	  <h2>DIP/LCA - Busca</h2>
	   
	   <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
<%If TIPO_BUSCA=3 Then%>
<table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="100%" border="0" align="center" cellpadding="5" cellspacing="0">
		  <tr>
         
            <td width="96%">
              <form action="adm_busca_completa_depto_foto.asp" style="margin:0px">
            <table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td width="87%" class="textoMedio"><strong>Busca avan&ccedil;ada de &aacute;lbum:</strong></td>
                <td width="13%" rowspan="2"><div align="right">
                  <input class="btnFiltrarAmarelo" type="submit" tabindex="4" name="filtro1" value=" " />
                </div></td>
              </tr>
              <tr>
                <td><table width="675" border="0" cellpadding="2" cellspacing="0">
                  <tr>
                    <td width="102"><div align="right">N&uacute;mero do &Aacute;lbum:</div></td>
                    <td width="107"><label>
                      <input tabindex="1" type="text" value="<%=BUSCAR_TERMO%>" class="inputTamanho1" name="BUSCAR_TERMO" id="textfield" />
                    </label></td>
                    <td width="99"><div align="right">Ano de Conclus&atilde;o:</div></td>
                    <td width="102"><input tabindex="2" type="text" value="<%=ANO_CONCLUSAO%>" class="inputTamanho1" name="ANO_CONCLUSAO" /></td>
                    <td width="131"><div align="right">Semestre de Conclus&atilde;o:</div></td>
                    <td width="110">
                      <select tabindex="3" name="SEMESTRE_CONCLUSAO">
                        <option value="">Selecione</option>
                        <option <%If SEMESTRE_CONCLUSAO = "J" Then%> selected="selected" <%End If%> value="J">Julho</option>
                        <option <%If SEMESTRE_CONCLUSAO = "D" Then%> selected="selected" <%End If%> value="D">Dezembro</option></select>
            <input name="TIPO_BUSCA" value="3" type="hidden" />
                    </td>
                  </tr>
                </table></td>
              </tr>
            </table>
            </form>
            </td>

          </tr>
        </table>
<%
End If
%>
          
          
		  <%
		  If Request.QueryString("BUSCAR_TERMO") <> Empty AND BUSCA_ALBUM = 1 Then
		  %>
          <h3 align="center">Foram  encontrados os seguintes registros para a busca: </h3>
		  <%
		  AbrirConexao()
  		  'AbrirConexao2()
		  Set rs = objConexao.Execute(sql,qtdRegistrosBusca)		  
		 %>
		  
		  <%If not rs.Eof Then%>
		  <div id="comboTexto">
			<%
			 If SQL_TIPO = 1 Then
			%>
          		<table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="8%"><div align="center"><strong>C&oacute;d. Pedido </strong></div></td>
            <td width="14%"><div align="center"><strong>Formando</strong></div></td>
            <td width="16%"><div align="center"><strong>Institui&ccedil;&atilde;o de Ensino&nbsp;</strong></div></td>
            <td width="14%"><div align="center"><strong>Curso</strong></div></td>
            <td width="8%"><div align="center"><strong>&Aacute;lbum</strong></div></td>
            <td width="16%"><div align="center"><strong>Vendedor</strong></div></td>
            <td width="14%"><div align="center"><strong>Finaliza Depto. Foto </strong></div></td>
            <td width="10%"><div align="center"><strong>Finaliza Diretoria </strong></div></td>
            </tr>
		  <%
		  While not rs.Eof
		  
		  'TB_PEDIDOS		  
		  Dim COD_PEDIDO, FINALIZA_CORRECAO, FINALIZA_DIRETORIA, DATA_CORRIGIDO_DEP_FOTO, COD_VENDEDOR
		  
		  COD_PEDIDO				= rs("COD_PEDIDO")
		  FINALIZA_CORRECAO			= rs("FINALIZA_CORRECAO")
		  FINALIZA_DIRETORIA		= rs("FINALIZA_DIRETORIA")
		  DATA_CORRIGIDO_DEP_FOTO	= rs("DATA_CORRIGIDO_DEP_FOTO")
		  DATA_CORRIGIDO_DIRETORIA	= rs("DATA_CORRIGIDO_DIRETORIA")
		  COD_VENDEDOR				= rs("COD_VENDEDOR")
		   
		  
		  
	  	  %>
          <tr>
            <td><div align="center">
            <%IF int(FINALIZA_CORRECAO) = 1 Then%>
            <a href="lca_busca_detalhamento_pedido.asp?cod_pedido=<%=COD_PEDIDO%>"><%=COD_PEDIDO%></a>
            <%Else%>
            <%=COD_PEDIDO%><br /> <span class="textoMenor"><em>(N&atilde;o consulta <a href="javascript:alert('N&atilde;o &eacute; poss&iacute;vel consultar, pois o pedido n&atilde;o est&aacute; corrigido.')"><img src="images/icons/icon_help_mini.jpg"  border="0"/></a>)</em></span>
            <%End If%>
            </div></td>
            <td colspan="4"><div align="left">
			<%
			sql = 	"SELECT TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO,TB_FICHAS_IDENTIFICACAO.NOME_CURSO, TB_PEDIDOS_FORMANDOS.COD_FORMANDO,TB_FICHAS_IDENTIFICACAO.ENCADERNACOES_EXTRAS, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, TB_FICHAS_IDENTIFICACAO.ALBUM,  TB_FICHAS_IDENTIFICACAO.COD_CONTRATO "&_
					"FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON "&_
					"TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO "&_
					"WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&COD_PEDIDO&" ORDER BY ALBUM"
					
					'Response.Write(sql)
			Set rs3 = objConexao.Execute(sql)
				
		
				While not rs3.Eof
				
				%>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				  <tr>
					<!--<td width="23%"><div align="center"><%=digito(rs3("COD_FORMANDO"))%></div></td>-->
					<td width="27%"><div align="left"><%=rs3("NOME_FORMANDO")%></div></td>
					<td width="31%">					  
					  <div align="left">
					    <%
					Dim rs6
					Set rs6 = objConexao.Execute("SELECT DESCRICAO FROM TB_ENTIDADES WHERE COD_CONTRATO= "&sqlTexto(rs3("COD_CONTRATO")))
						
						If not rs6.Eof Then
							Response.Write(rs6("DESCRICAO"))
						End If
						
					rs6.Close
					Set rs6 = Nothing
					%>
					    </div></td><td width="28%"><div align="left"><%=rs3("NOME_CURSO")%></div></td>
				    <td width="14%"><div align="center"><%=rs3("ALBUM")%>
					<%
					ENCADERNACOES_EXTRAS = formatNumber(0&rs3("ENCADERNACOES_EXTRAS"),0)
					
					If ENCADERNACOES_EXTRAS = 1 Then
							Response.Write(" DUPLO")
					End If
					%></div>
				      <div align="center"></div></td>
				  </tr>
				</table>
				<%
				rs3.MoveNext()
				Wend
				
			rs3.Close
			Set rs3 = Nothing
			%>
			</div></td>
            <td>
			  <div align="center">
			  <%
			  Dim rs17
			  Set rs17 = objConexao.Execute("SELECT NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR="&rs("COD_VENDEDOR"))
			  	
				If not rs17.Eof Then
					Response.Write(rs17("NOME_VENDEDOR"))
				End If
				
			  
			  rs17.Close
			  Set rs17 = Nothing
			  %></div></td>
            <td>
			  <div align="center">
			    <%If FINALIZA_CORRECAO = 1 Then%>
			    <img src="images/icons/icon_balao_green.gif" /><span class="textoMenor">(<%=DATA_CORRIGIDO_DEP_FOTO%>)</span>
			    
			    <%Else%>
				<img src="images/icons/icon_balao_red.gif" />
		        <%End If%>			
			    </div></td>
            <td><div align="center">
              <%If FINALIZA_DIRETORIA = 1 Then%>
              <img src="images/icons/icon_balao_green.gif" /><span class="textoMenor">(<%=DATA_CORRIGIDO_DIRETORIA%>)</span> 
              <%Else%>
			  <img src="images/icons/icon_balao_red.gif" />
              <%End If%>
            </div></td>
            </tr>
          
		  <%
		  rs.MoveNext()
		  Wend
	      
		  rs.Close()
		  Set rs = Nothing
		  'FecharConexao2()
		  FecharConexao()
		  %>
		  <tr bgcolor="whitesmoke">
            <td colspan="8"><div align="center"><strong>Quantidade de Registros Encontrados: <%=qtdRegistrosBusca%></strong></div></td>
            </tr>
        </table>
		  </div>
		  	<%ElseIf SQL_TIPO = 2 Then%>
			<table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="10%"><div align="center"><strong>C&oacute;d. Formando </strong></div></td>
            <td width="21%"><div align="center"><strong>Nome </strong></div></td>
            <td width="10%"><div align="center"><strong>&Aacute;lbum</strong></div></td>
            <td width="10%"><div align="center"><strong>C&oacute;d. Produ&ccedil;&atilde;o</strong></div></td>
            <td width="26%"><div align="center"><strong>Institui&ccedil;&atilde;o de Ensino </strong></div></td>
            <td width="21%"><div align="center"><strong>Curso </strong></div></td>
            <td width="12%"><div align="center"><strong>Vendido</strong></div></td>
            </tr>
		  <%
		  While not rs.Eof
		  	  
		  'TB_FICHAS_IDENTIFICACAO
		  'Dim COD_CONTRATO, NOME_FORMANDO, NOME_CURSO, VENDIDO
		  
		  'TB_PEDIDOS		  
		  'Dim COD_PEDIDO, FINALIZA_CORRECAO, FINALIZA_DIRETORIA, DATA_CORRIGIDO_DEP_FOTO, COD_VENDEDOR
		  
	  	  %>
          <tr>
            <td><div align="center"><a href="dip_busca_detalhamento_formando.asp?cod_formando=<%=rs("COD_FORMANDO")%>" title="Clique aqui para detalhar o formando"><%=digito(rs("COD_FORMANDO"))%></a></div></td>
            <td><%=rs("NOME_FORMANDO")%></td>
            <td><div align="center"><%=rs("ALBUM")%>
            <%
					ENCADERNACOES_EXTRAS = formatNumber(0&rs("ENCADERNACOES_EXTRAS"),0)
					
					If ENCADERNACOES_EXTRAS = 1 Then
							Response.Write(" DUPLO")
					End If
					%></div></td>
            <td><div align="center"><%=rs("cod_producao")%></div></td>
            
            <td>
            <%
					Set rs6 = objConexao.Execute("SELECT DESCRICAO FROM TB_ENTIDADES WHERE  COD_CONTRATO= "&sqlTexto(rs("COD_CONTRATO")))
						
						If not rs6.Eof Then
							Response.Write(rs6("DESCRICAO"))
						End If
						
					rs6.Close
					Set rs6 = Nothing
					%></td>
            <td><%=rs("NOME_CURSO")%></td>
            <td>
			<%
			DIM PRE_VENDA
			Set rs6 = objConexao.Execute(	"SELECT TB_PEDIDOS.COD_PEDIDO, TB_PEDIDOS.FINALIZA_CORRECAO, TB_PEDIDOS.PRE_VENDA FROM TB_PEDIDOS INNER JOIN TB_PEDIDOS_FORMANDOS ON TB_PEDIDOS.COD_PEDIDO = "&_
											"TB_PEDIDOS_FORMANDOS.COD_PEDIDO WHERE TB_PEDIDOS_FORMANDOS.COD_FORMANDO = '"&rs("COD_FORMANDO")&"' AND FINALIZA_VENDEDOR = 1")
			
			
			If not rs6.Eof Then
				
				IF rs6("PRE_VENDA") = "S" Then 
				   PRE_VENDA = " É PRÉ-VENDA "
				Else
				   PRE_VENDA = ""
                END IF


				If int(rs6("FINALIZA_CORRECAO")) = 1 Then
				%> 	Sim, <%=PRE_VENDA%> (<a href="lca_busca_detalhamento_pedido.asp?cod_pedido=<%=rs6("COD_PEDIDO")%>" title="Clique aqui para detalhar o pedido"><%=rs6("COD_PEDIDO")%></a>) 	<%
				Else
				%>
				Sim (<%=rs6("COD_PEDIDO")%>, mas ainda n&atilde;o foi corrigido. ) <%=PRE_VENDA%>
				<%End If%>
				
			<%Else%>
			N&atilde;o
			<%
			End If
			
			rs6.Close
			Set rs6=Nothing
			%>			</td>
            </tr>
          
		  <%
		  rs.MoveNext()
		  Wend
	      
		  rs.Close()
		  Set rs = Nothing
		  'FecharConexao2()
		  FecharConexao()
		  %>
		  <tr bgcolor="whitesmoke">
            <td colspan="7"><div align="center"><strong>Quantidade de Registros Encontrados: <%=qtdRegistrosBusca%></strong></div></td>
            </tr>
        </table>
		  </div>
			<%End If%>
		 
		    
		      <%else
			  
			  %>
		      <div align="center">
			  <p><em>N&atilde;o foi encontrado registros para a busca informada!</em>	</p>
		    <p>&nbsp;</p>
		    <p>	 
			 </div>     
		      <%end if%>
		      </p>
		 
		  <%Else%>
		  <p>&nbsp;</p>
		  		  <p>&nbsp;</p>
				  		  <p align="center" class="textoMaior">Utilize a busca do menu no topo da p&aacute;gina!</p>
			  		    <p>&nbsp;</p>
								  		  <p>&nbsp;</p>
		                             
		 <%End If%>
	
		                                  
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