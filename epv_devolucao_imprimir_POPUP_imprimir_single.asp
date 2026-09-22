<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<%
Dim qtdItensDev
qtdItensDev = 0
'Para tirar a valida��o da p�gina basta tirar o "#" ou coloca-lo para habilitar antes do include
		AbrirConexao()
		'AbrirConexao2()


		
		'1) Verifica se h� pendencias de devolu��o
		'2) Verifica seleciona todas as fichas
		Dim rs, rs2, rsVendedorVerifica
		
		Set rs = objConexao.Execute("SELECT COD_ETIQUETA FROM TB_FICHAS_IDENTIFICACAO_DEVOLUCAO_TEMPORARIA WHERE COD_USUARIO="&Session("COD_USUARIO")&" GROUP BY COD_ETIQUETA")
		
		
		If not rs.Eof Then
			TIPO = Right(rs("COD_ETIQUETA"),1)				
			COD_FORMANDO = Mid(rs("COD_ETIQUETA"),1,Len(rs("COD_ETIQUETA"))-1)
			'Seleciona a �ltima devolu��o deste c�d_formando
			
			sql = "SELECT TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.TIPO, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_SAIDA, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_RETORNO FROM TB_VENDEDORES INNER JOIN TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ON TB_VENDEDORES.COD_VENDEDOR = TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR WHERE COD_FORMANDO = '"&COD_FORMANDO&"' AND TIPO = '"&TIPO&"' ORDER BY DATA_SAIDA DESC LIMIT 1"
			'Response.Write(sql)
			Set rsVendedorVerifica = objConexao.Execute(sql)
		
			If not rsVendedorVerifica.Eof Then
				Dim NOME_VENDEDOR, COD_VENDEDOR
				NOME_VENDEDOR 	= rsVendedorVerifica("NOME_VENDEDOR")
				COD_VENDEDOR 	= rsVendedorVerifica("COD_VENDEDOR")
			End if	
		
			rsVendedorVerifica.Close
			Set rsVendedorVerifica = Nothing


      sql = "SELECT * FROM TB_PRODUCAO WHERE COD_PRODUCAO IN (SELECT COD_PRODUCAO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO = "&sqlNumero(COD_FORMANDO)&")"
      set rs3 = objConexao.execute(sql)
      if not rs3.Eof Then
        descricoa_producao = rs3("cod_producao").value &" - "&rs3("descricao").value
      end if
      rs3.close()
      set rs3 = Nothing

		%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="print"></link>
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<style type="text/css" media="print">
<!--
.folha {
    page-break-after: always;
}
.btnFinalizaProducao{	background-image:url(images/layout/btnFinalizaProducao.jpg);
	cursor:pointer;
	width:140px;
	height:22px;
	border:0px;
	}
body,td,th {
	font-size: 10px;
}
.style1 {font-size: 12px}
-->
</style>

<title>StillosWEB - Sistema corporativo - Vers�o 1.2</title>
	<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>
	<script type="text/javascript">
function confirma(cod_formando,nomeFormando,comando){
var mensagem

	if (comando == "alterar"){
		mensagem = "ATEN��O ALTERA��O: \nDeseja resetar a ficha de: "+nomeFormando+" deseja continuar?\n*A ficha voltar� na forma original."	
	}	
	else {
		mensagem = "ATEN��O EXCLUS�O: \nSe voc� continuar poder� perder as altera��es feitas nesta p�gina,\nalem de resetar a ficha do formando: "+nomeFormando+"  deseja continuar?"
	}

	var resposta=confirm(mensagem);
		if (resposta==true) {
			
			if (comando == "excluir"){
				//Nova confirma��o, Exclui
				var resposta2=confirm("Deseja realmente EXCLUIR?")
					if (resposta2==true){
						document.producao["MODIFICACOES_"+cod_formando].value = "2";
						document.producao.submit();
					}	
			}
			else{
			//Altera
				document.producao["MODIFICACOES_"+cod_formando].value = "1"
				document.producao.submit();
			}
			
		}
}
</script>

	
	<script language="javascript">
		function popUp(campo,URL,campoSelect,msgErro){
			
			if(document.producao[campo].value==1){
				
				if (campoSelect !== 'no'){
					//Brinde = "Sim"
					document.producao[campoSelect].options[1].selected = "true";
				}
			
				win=window.open(URL,"","width=350,height=220,scrollbars=yes")
				var text = "Se a janela NÃO abrir desbloquei os Pop-ups!";
				if(win == null) { alert(text); return; }
				win.moveTo(350,150);
					
			}
			else {
				if (msgErro==1){alert('Para visualizar os brindes � necess�rio alterar a op��o Brinde para Sim.')}
			}
		
		}
		
		function fecharPopUp(){
			if(win){
				//win.close();
			}
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
	
	<script>

function mascara(o,f){
    v_obj=o
    v_fun=f
    setTimeout("execmascara()",1)
}

function execmascara(){
    v_obj.value=v_fun(v_obj.value)
}

function leech(v){
    v=v.replace(/o/gi,"0")
    v=v.replace(/i/gi,"1")
    v=v.replace(/z/gi,"2")
    v=v.replace(/e/gi,"3")
    v=v.replace(/a/gi,"4")
    v=v.replace(/s/gi,"5")
    v=v.replace(/t/gi,"7")
    return v
}

function soNumeros(v){
    return v.replace(/\D/g,"")
}

function telefone(v){
    v=v.replace(/\D/g,"")
    v=v.replace(/^(\d\d)(\d)/g,"($1) $2") 
    v=v.replace(/(\d{4})(\d)/,"$1-$2")   
    return v
}

function cpf(v){
    v=v.replace(/\D/g,"")                   
    v=v.replace(/(\d{3})(\d)/,"$1.$2")      
    v=v.replace(/(\d{3})(\d)/,"$1.$2")      
                                            
    v=v.replace(/(\d{3})(\d{1,2})$/,"$1-$2")
    return v
}

function cep(v){
    v.replace(/\D/g,"")               
    v=v.replace(/^(\d{5})(\d)/,"$1-$2") 
    return v
}

function cnpj(v){
    v=v.replace(/\D/g,"")                           
    v=v.replace(/^(\d{2})(\d)/,"$1.$2")             
    v=v.replace(/^(\d{2})\.(\d{3})(\d)/,"$1.$2.$3") 
    v=v.replace(/\.(\d{3})(\d)/,".$1/$2")         
    v=v.replace(/(\d{4})(\d)/,"$1-$2")              
    return v
}

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
	<!--
VB Confirm SIM/NÃO
-->

<script language=javascript>
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
function imprimir(){
	document.getElementById("imprimir").style.display = "none";
	document.getElementById("imprimir").style.visibility = "hidden";
	window.print();
	//var mensagem = pergunta = confirm('A impressão foi realizada com sucesso?')
	
	
	
	if(YNconfirm("A impressão foi realizada com sucesso?")){
				//Faz o reload na página anterior
				parent.location.href = 'dip_producoes_fichas_recibos_vendedores.asp?msg=Liberação concluída com sucesso!&cod_vend=<%=COD_VENDEDOR%>&cod_prod=<%=Request.QueryString("cod_prod")%>'
	}else{
			document.getElementById("imprimir").style.display = "block";
			document.getElementById("imprimir").style.visibility = "visible";			
			alert('NÃO FECHE ESTÁ TELA E CONTATE O DEPTO. DE T.I!');
			//Faz a barra de impressão novamente e o alert
			
	}
	
}
function fecharJanela(){
parent.location.href = 'dip_producoes_fichas_recibos_vendedores.asp?msg=2° Via impressa com sucesso!&cod_vend=<%=COD_VENDEDOR%>&cod_prod=<%=Request.QueryString("cod_prod")%>'
}
</script>
    <style type="text/css">
<!--
.style1 {font-weight: bold}
.style2 {
	font-size: 10px;
	font-weight: bold;
}
.style3 {font-size: 10px}
-->
    </style>
</head>
<script>
function imprimeFecha(){
	window.print();
	//window.close();
}
</script>
<body onload="imprimeFecha()" style="background:#FFFFFF">
<div class="folha">
<table width="100%" align="center" border="0" cellpadding="0" cellspacing="3">
<tr>
<td>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td valign="top"><h2 style="color:#000000; border-bottom:solid #000 3px; padding-bottom:2px">Devolução de Produtos para o Estoque </h2>
          <br />
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="90%" style="padding-bottom:10px"><table width="100%" border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td width="256" class="textoMedio">Data do retorno ao estoque:</td>
                  <td width="801"><strong class="textoMedio"><%=ArrumaData(Now())%> (1° Via Stillo's)</strong></td>
                </tr>
                <tr>
                  <td width="256" class="textoMedio">Produção</td>
                  <td width="801"><strong class="textoMedio"><%=descricoa_producao%></strong></td>
                </tr>
                <tr>
                  <td class="textoMedio">Representante Comercial:</td>
                  <td class="textoMedio"><strong><%=NOME_VENDEDOR%></strong></td>
                </tr>
              </table></td>
              <td width="10%"><div align="center"></div></td>
            </tr>
          </table></td>
        </tr>
      <tr>
        <td>   
		       <table width="100%" border="0" align="center" 	cellpadding="1" cellspacing="0">
			<tr bgcolor="whitesmoke">
       
            <td width="5%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style2">Retirada n° </div></td>
            <td width="4%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style2">Cód. Formando </div></td>
            <td width="5%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style2">Álbum</div></td>
            <td width="21%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style2">Nome</div></td>
            <td width="27%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style2">Faculdade</div></td>
            <td width="22%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style2">Curso</div></td>
            <td width="16%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style3"><div align="center"><strong>Item Devolvido </strong></div></div></td>
            </tr>
		
						
						<%

						While not rs.Eof
						'*Significa possui diversos formandos
								Dim COD_FORMANDO, TIPO, NOME_VENDEDOR_NEW, COD_VENDEDOR_NEW, DATA_RETORNO, COD_MOVIMENTACAO, NOME_CURSO, COD_CONTRATO, viaRetirada, ENCADERNACOES_EXTRAS
								viaRetirada = 0
								TIPO = Right(rs("COD_ETIQUETA"),1)
								COD_FORMANDO = Mid(rs("COD_ETIQUETA"),1,Len(rs("COD_ETIQUETA"))-1)
								
								sql = "SELECT * FROM VW_DEVOLUCAO_PRODUTOS WHERE COD_FORMANDO="&COD_FORMANDO&" AND TIPO="&TIPO&" ORDER BY TIPO, DATA_SAIDA DESC"
						
						Set rs2 = objConexao.Execute(sql,viaRetirada)
						
						If not rs2.Eof Then
							COD_MOVIMENTACAO		= rs2("COD_MOVIMENTACAO")
							COD_VENDEDOR_NEW		= rs2("COD_VENDEDOR")
							ALBUM					= rs2("ALBUM")
							DATA_RETORNO			= ""&rs2("DATA_RETORNO")
							NOME_FORMANDO			= rs2("NOME_FORMANDO")
							DATA_SAIDA				= rs2("DATA_SAIDA")	
							NOME_CURSO				= rs2("NOME_CURSO")
							COD_CONTRATO			= rs2("COD_CONTRATO")	
							ENCADERNACOES_EXTRAS	= formatNumber(0&rs2("ENCADERNACOES_EXTRAS"),0)
						
						qtdItensDev = qtdItensDev + 1
						%>
          					<tr>
            <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333; border-left:solid 0.25px #333333"><div align="center"><strong><%=viaRetirada%></strong></div></td>
            <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333"><div align="center" style="font-size:7px"><%=digito(COD_FORMANDO)%></div></td>
            <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333"><div align="center" style="font-size:7px"><%=ALBUM%></div></td>
            <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333"><div align="left" style="font-size:7px"><%=NOME_FORMANDO%></div></td>
            <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333">
              
              <div align="left" style="font-size:7px">
                <%
			Set rsTurma = objConexao.Execute("SELECT DESCRICAO FROM TB_ENTIDADES WHERE COD_CONTRATO="&sqlTexto(COD_CONTRATO))
				
				If not rsTurma.Eof Then
					Response.Write(rsTurma("DESCRICAO"))
				Else%>
                <span class="msg">Falha entidade de ensino não encontrada, informe ao depto. de T.I</span>
                    <%End If
			
			rsTurma.Close
			Set rsTurma = Nothing
			%>
              </div></td>
            <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333"><div align="left" style="font-size:7px"><%=NOME_CURSO%></div></td>
            <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333"><div align="center" style="font-size:7px">
			<%If TIPO = 0 Then%>
					<%IF ENCADERNACOES_EXTRAS = 1 Then%>
						<strong>ÁLBUM DUPLO</strong>			
					<%Else%>
						<strong>ÁLBUM</strong>
					<%End If%>
<%Elseif TIPO = 1 Then%>
<strong>FICHA</strong>
<%Else
			'Produto agregado, consultar qual �.
			sql = 	"SELECT TB_PRODUTOS.NOME_PRODUTO "&_
					"FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = "&_
					"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO WHERE "&_
					"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO = "&COD_FORMANDO&" AND "&_
					"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.TIPO = "&TIPO&" LIMIT 1"
			Set rsProdAgrePend = objConexao.Execute(sql)
				If not rsProdAgrePend.Eof Then
				%>
				<strong>				<%
						Dim NOME_PRODUTO_AGRE
						NOME_PRODUTO_AGRE = rsProdAgrePend("NOME_PRODUTO")
						NOME_PRODUTO_AGRE = Replace(NOME_PRODUTO_AGRE,"Trio (2 porta-retrados + 1 foto-Tela 20X30)","TRIO")
						Response.Write(UCASE(NOME_PRODUTO_AGRE))
				%></strong>					
			    <%
				End If
			rsProdAgrePend.Close
			Set rsProdAgrePend = Nothing
			End If%>
			</strong></div></td>
            </tr>
          				<%
						End If
							
						rs2.Close
						Set rs2 = Nothing
							
						rs.MoveNext()
						Wend
						%>
						<%
						'For i = 0 to 49-qtdItensDev
						%>
						<tr>
          					  <td></td>
          					  <td></td>
          					  <td></td>
          					  <td></td>
          					  <td></td>
          					  <td></td>
          					  <td></td>
   					  </tr>
					  	<%'Next%>
					  
					</table>
	      
	      <p>Qtd de Itens devolvidos: <strong><%=qtdItensDev%></strong><br />Obs.: Álbuns duplos constam apenas como um único álbum com a quantidade de fotos equivalente a soma dos dois álbuns do formando. </p>
	      <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
           <tr>
             <td style="border-bottom:solid 1px #000000">
               <p></p></td>
           </tr>
           
           <tr>
             <td><div align="center"><span class="textoMedio"><span class="style1">REPRESENTANTE COMERCIAL</span><br />
                   <span class="style1">  <strong><%=NOME_VENDEDOR%></strong></span></span></div></td>
           </tr>
         </table>
		 	      <p></p>
		   <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
           <tr>
             <td style="border-bottom:solid 1px #000000">
               <p></p></td>
           </tr>
           <tr>
             <td><div align="center"><span class="textoMedio style1">RECEBIDO POR<strong><br />
               <%=Ucase(Session("NOME"))%></strong></span></div></td>
           </tr>
         </table>
		<%
		rs.Close
		Set rs = Nothing
		%>
		<!-- #include file="includes/rodape_impressao.asp" -->
		<%End If%>
	
		 </td>
        </tr>


    </table></td>
  </tr>
</table>
</td>
</tr>
</table>
</div>
<%
qtdItensDev = 0
		
		'1) Verifica se h� pendencias de devolu��o
		'2) Verifica seleciona todas as fichas
		
		Set rs = objConexao.Execute("SELECT COD_ETIQUETA FROM TB_FICHAS_IDENTIFICACAO_DEVOLUCAO_TEMPORARIA WHERE COD_USUARIO="&Session("COD_USUARIO")&" GROUP BY COD_ETIQUETA")
		
		
		If not rs.Eof Then
			TIPO = Right(rs("COD_ETIQUETA"),1)				
			COD_FORMANDO = Mid(rs("COD_ETIQUETA"),1,Len(rs("COD_ETIQUETA"))-1)
			'Seleciona a �ltima devolu��o deste c�d_formando
			
			sql = "SELECT TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.TIPO, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_SAIDA, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_RETORNO FROM TB_VENDEDORES INNER JOIN TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ON TB_VENDEDORES.COD_VENDEDOR = TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR WHERE COD_FORMANDO = '"&COD_FORMANDO&"' AND TIPO = '"&TIPO&"' ORDER BY DATA_SAIDA DESC LIMIT 1"
		
		%>
<table width="100%" align="center" border="0" cellpadding="0" cellspacing="3">
<tr>
<td>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td valign="top"><h2 style="color:#000000; border-bottom:solid #000 3px; padding-bottom:2px">Devolução de Produtos para o Estoque </h2>
          <br />
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="90%"  style="padding-bottom:10px"><table width="100%" border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td width="256" class="textoMedio">Data do retorno ao estoque:</td>
                  <td width="801"><strong class="textoMedio"><%=ArrumaData(now())%> (2° Via)</strong></td>
                </tr>
                <tr>
                  <td width="256" class="textoMedio">Produção</td>
                  <td width="801"><strong class="textoMedio"><%=descricoa_producao%></strong></td>
                </tr>
                <tr>
                  <td class="textoMedio">Representante Comercial:</td>
                  <td class="textoMedio"><strong><%=NOME_VENDEDOR%></strong></td>
                </tr>
              </table></td>
              <td width="10%"><div align="center"></div></td>
            </tr>
          </table></td>
        </tr>
      <tr>
        <td>   
		       <table width="100%" border="0" align="center" 	cellpadding="1" cellspacing="0">
			<tr bgcolor="whitesmoke">
       
            <td width="5%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style2">Retirada n° </div></td>
            <td width="4%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style2">Cód. Formando </div></td>
            <td width="5%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style2">Álbum</div></td>
            <td width="21%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style2">Nome</div></td>
            <td width="27%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style2">Faculdade</div></td>
            <td width="22%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style2">Curso</div></td>
            <td width="16%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style3"><div align="center"><strong>Item Devolvido </strong></div></div></td>
            </tr>
		
						
						<%

						While not rs.Eof
						'*Significa possui diversos formandos
								
								viaRetirada = 0
								TIPO = Right(rs("COD_ETIQUETA"),1)
								COD_FORMANDO = Mid(rs("COD_ETIQUETA"),1,Len(rs("COD_ETIQUETA"))-1)
								
								sql = "SELECT * FROM VW_DEVOLUCAO_PRODUTOS WHERE COD_FORMANDO="&COD_FORMANDO&" AND TIPO="&TIPO&" ORDER BY TIPO, DATA_SAIDA DESC"
						
						Set rs2 = objConexao.Execute(sql,viaRetirada)
						
						If not rs2.Eof Then
							COD_MOVIMENTACAO		= rs2("COD_MOVIMENTACAO")
							COD_VENDEDOR_NEW		= rs2("COD_VENDEDOR")
							ALBUM					= rs2("ALBUM")
							DATA_RETORNO			= ""&rs2("DATA_RETORNO")
							NOME_FORMANDO			= rs2("NOME_FORMANDO")
							DATA_SAIDA				= rs2("DATA_SAIDA")	
							NOME_CURSO				= rs2("NOME_CURSO")
							COD_CONTRATO			= rs2("COD_CONTRATO")	
							ENCADERNACOES_EXTRAS	= formatNumber(0&rs2("ENCADERNACOES_EXTRAS"),0)
						
						qtdItensDev = qtdItensDev + 1
						%>
          					<tr>
            <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333; border-left:solid 0.25px #333333"><div align="center"><strong><%=viaRetirada%></strong></div></td>
            <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333"><div align="center" style="font-size:7px"><%=digito(COD_FORMANDO)%></div></td>
            <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333"><div align="center" style="font-size:7px"><%=ALBUM%></div></td>
            <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333"><div align="left" style="font-size:7px"><%=NOME_FORMANDO  %></div></td>
            <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333">
              
              <div align="left" style="font-size:7px">
                <%
			Set rsTurma = objConexao.Execute("SELECT DESCRICAO FROM TB_ENTIDADES WHERE COD_CONTRATO="&sqlTexto(COD_CONTRATO))
				
				If not rsTurma.Eof Then
					Response.Write(rsTurma("DESCRICAO"))
				Else%>
                <span class="msg">Falha entidade de ensino não encontrada, informe ao depto. de T.I</span>
                    <%End If
			
			rsTurma.Close
			Set rsTurma = Nothing
			%>
              </div></td>
            <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333"><div align="left" style="font-size:7px"><%=NOME_CURSO%></div></td>
            <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333"><div align="center" style="font-size:7px">
			<%If TIPO = 0 Then%>
					<%IF ENCADERNACOES_EXTRAS = 1 Then%>
						<strong>ÁLBUM DUPLO</strong>			
					<%Else%>
						<strong>ÁLBUM</strong>
					<%End If%>
<%Elseif TIPO = 1 Then%>
<strong>FICHA</strong>
<%Else
			'Produto agregado, consultar qual �.
			sql = 	"SELECT TB_PRODUTOS.NOME_PRODUTO "&_
					"FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = "&_
					"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO WHERE "&_
					"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO = "&COD_FORMANDO&" AND "&_
					"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.TIPO = "&TIPO&" LIMIT 1"
			Set rsProdAgrePend = objConexao.Execute(sql)
				If not rsProdAgrePend.Eof Then
				%>
				<strong>				<%
						
						NOME_PRODUTO_AGRE = rsProdAgrePend("NOME_PRODUTO")
						NOME_PRODUTO_AGRE = Replace(NOME_PRODUTO_AGRE,"Trio (2 porta-retrados + 1 foto-Tela 20X30)","TRIO")
						Response.Write(UCASE(NOME_PRODUTO_AGRE))
				%></strong>					
			    <%
				End If
			rsProdAgrePend.Close
			Set rsProdAgrePend = Nothing
			End If%>
			</strong></div></td>
            </tr>
          				<%
						End If
							
						rs2.Close
						Set rs2 = Nothing
							
						rs.MoveNext()
						Wend
						%>
						<%
						'For i = 0 to 49-qtdItensDev
						%>
						<tr>
          					  <td></td>
          					  <td></td>
          					  <td></td>
          					  <td></td>
          					  <td></td>
          					  <td></td>
          					  <td></td>
   					  </tr>
					  	<%'Next%>
					  
					</table>
	      
	      <p>Qtd de Itens devolvidos: <strong><%=qtdItensDev%></strong><br />Obs.: Álbuns duplos constam apenas como um único álbum com a quantidade de fotos equivalente a soma dos dois álbuns do formando. </p>
	      <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
           <tr>
             <td style="border-bottom:solid 1px #000000">
               <p></p></td>
           </tr>
           <tr>
             <td><div align="center"><span class="textoMedio"><span class="style1">REPRESENTANTE COMERCIAL</span><br />
                   <span class="style1">  <strong><%=NOME_VENDEDOR%></strong></span></span></div></td>
           </tr>
         </table>
		 	      <p></p>
		   <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
           <tr>
             <td style="border-bottom:solid 1px #000000">
               <p></p></td>
           </tr>
           <tr>
             <td><div align="center"><span class="textoMedio style1">RECEBIDO POR<strong><br />
               <%=Ucase(Session("NOME"))%></strong></span></div></td>
           </tr>
         </table>
		<%
		rs.Close
		Set rs = Nothing
		%>
		<!-- #include file="includes/rodape_impressao.asp" -->
		<%End If%>
	
		 </td>
        </tr>


    </table></td>
  </tr>
</table>
</td>
</tr>
</table>
</body>
</html>
<%
		'FecharConexao2()
		FecharConexao()
%>



