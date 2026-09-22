<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<%
Dim DATA_SAIDA, COD_FORNECEDOR_SAIDA, COD_FORNECEDOR, NOME_RESPONSAVEL_SAIDA, RG_RESPONSAVEL_SAIDA, NOME_FANTASIA
COD_FORNECEDOR_SAIDA = Request.QueryString("COD_FORNECEDOR_SAIDA")

AbrirConexao()

Set rs=objConexao.Execute("SELECT * FROM TB_FORNECEDORES_SAIDA WHERE COD_FORNECEDOR_SAIDA = "&COD_FORNECEDOR_SAIDA)
	
	If not rs.Eof Then
			COD_FORNECEDOR			= rs("COD_FORNECEDOR")
			NOME_RESPONSAVEL_SAIDA	= rs("NOME_RESPONSAVEL_SAIDA")
			RG_RESPONSAVEL_SAIDA	= rs("RG_RESPONSAVEL_SAIDA")
			DATA_SAIDA				= formatDateTime(rs("DATA_SAIDA"))
			
	End If
	
rs.Close
Set rs=Nothing

Set rs=objConexao.Execute("SELECT * FROM TB_FORNECEDORES WHERE COD_FORNECEDOR = "&COD_FORNECEDOR)
	
	If not rs.Eof Then
				RAZAO_SOCIAL 	= rs("RAZAO_SOCIAL")
				NOME_FANTASIA 	= rs("NOME_FANTASIA")
				CNPJ 			= rs("CNPJ")
				ENDERECO 		= rs("ENDERECO")
				BAIRRO 			= rs("BAIRRO")	
				CEP		 		= rs("CEP")
				CIDADE 			= rs("CIDADE")
				ESTADO 			= rs("ESTADO")
				TELEFONE 		= rs("TELEFONE")
	End If
	
rs.Close
Set rs=Nothing

	

FecharConexao()
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

<title>StillosWEB - Sistema corporativo - Versão 1.2</title>
	<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>
	<script type="text/javascript">
function confirma(cod_formando,nomeFormando,comando){
var mensagem

	if (comando == "alterar"){
		mensagem = "ATENÇÃO ALTERAÇÃO: \nDeseja resetar a ficha de: "+nomeFormando+" deseja continuar?\n*A ficha voltará na forma original."	
	}	
	else {
		mensagem = "ATENÇÃO EXCLUSÃO: \nSe você continuar poderá perder as alterações feitas nesta página,\nalem de resetar a ficha do formando: "+nomeFormando+"  deseja continuar?"
	}

	var resposta=confirm(mensagem);
		if (resposta==true) {
			
			if (comando == "excluir"){
				//Nova confirmação, Exclui
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
				if (msgErro==1){alert('Para visualizar os brindes é necessário alterar a opção Brinde para SIM.')}
			}
		
		}
		
		function fecharPopUp(){
			if(win){
				win.close();
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
</script>
    <style type="text/css">
<!--
.style1 {font-weight: bold}
.style2 {
	font-size: 12px;
	font-weight: bold;
}
.style3 {font-size: 10px}
.style4 {
	font-size: 15px;
	font-weight: bold;
}
.style5 {font-size: 20px}
-->
    </style>
</head>
<script>
function imprimeFecha(){
	window.print();
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
        <td valign="top"><h2 style="color:#000000; border-bottom:solid #000 3px; padding-bottom:2px">Saída de Produtos para Fornecedor</h2>
          <br />
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="90%" style="padding-bottom:10px"><table width="100%" border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td width="256" class="textoMedio">Data da Saída:</td>
                  <td width="801"><strong class="textoMedio"><%=ArrumaData(DATA_SAIDA)%> (1° Via Stillo's)</strong></td>
                </tr>
                <tr>
                  <td class="textoMedio">Razão social:</td>
                  <td class="textoMedio"><strong><%=RAZAO_SOCIAL%></strong></td>
                </tr>
                <tr>
                  <td class="textoMedio">Nome fantasia:</td>
                  <td class="textoMedio"><strong><%=NOME_FANTASIA%></strong></td>
                </tr>
                <tr>
                  <td class="textoMedio">CNPJ:</td>
                  <td class="textoMedio"><strong><%=CNPJ%></strong></td>
                </tr>
                <tr>
                  <td class="textoMedio">Endereço:</td>
                  <td class="textoMedio"><strong><%=ENDERECO%> - <%=BAIRRO%><BR /><%=CIDADE%>-<%=ESTADO%></strong></td>
                </tr>
                <tr>
                  <td class="textoMedio">Telefone:</td>
                  <td class="textoMedio"><strong><%=TELEFONE%></strong></td>
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
       
            <td style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style2">Descrição do Produto</div></td>
            <td style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style2">Quantidade</div></td>
            </tr>
		
						
						<%
						AbrirConexao()
						QTD_ITENS = 0
										
						Set rs =objConexao.Execute("SELECT COUNT(*) AS QTD_ITENS, COD_PRODUTO FROM TB_FORNECEDORES_SAIDA_ITENS WHERE  COD_FORNECEDOR_SAIDA = "&COD_FORNECEDOR_SAIDA&" GROUP BY COD_PRODUTO ORDER BY QTD_ITENS DESC")
						
						
						While not rs.Eof
						
						COD_PRODUTO_SELECIONADO = rs("COD_PRODUTO")
						QTD_ITENS				= QTD_ITENS + formatNumber(0&rs("QTD_ITENS"),0)
						
						%>
          					<tr>
                              <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333; border-left:solid 0.25px #333333"><div align="center">
                              <%
							  Set rs2=objConexao.Execute("SELECT * FROM TB_PRODUTOS WHERE COD_PRODUTO = "&COD_PRODUTO_SELECIONADO)
							  		
									If not rs2.Eof Then
											
											Response.Write(Ucase(rs2("NOME_PRODUTO")))
											
									End If
							  	
							  rs2.Close
							  Set rs2=Nothing
							  %>
                              </div></td>
                              <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333"><div align="center"><%=rs("QTD_ITENS")%></div></td>
            </tr>
          				<%	
						rs.MoveNext()
						Wend
						
						rs.Close
						Set rs = Nothing
						
						FecharConexao()
						%>
						<tr>
          					  <td colspan="2"><div align="center" class="style4">Quantidade total: <span class="style5"><%=QTD_ITENS%></span></div></td>
   					  </tr>
					</table>
	      
	      <p></p>
          <p></p>
          <p></p>
          <p></p>
	      <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
           <tr>
             <td style="border-bottom:solid 1px #000000">
               <p></p></td>
           </tr>
           <tr>
             <td><div align="center"><span class="textoMedio"><span class="style1">RESPONSÁVEL PELA RETIRADA</span><br />
                   <span class="style1">  <strong><%=NOME_RESPONSAVEL_SAIDA%><br />RG: <%=RG_RESPONSAVEL_SAIDA%></strong></span></span></div></td>
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
		<!-- #include file="includes/rodape_impressao.asp" -->

	
		 </td>
        </tr>


    </table></td>
  </tr>
</table>
</td>
</tr>
</table>
</div>
<table width="100%" align="center" border="0" cellpadding="0" cellspacing="3">
<tr>
<td>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td valign="top"><h2 style="color:#000000; border-bottom:solid #000 3px; padding-bottom:2px">Saída de Produtos para Fornecedor</h2>
          <br />
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="90%" style="padding-bottom:10px"><table width="100%" border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td width="256" class="textoMedio">Data da Saída:</td>
                  <td width="801"><strong class="textoMedio"><%=ArrumaData(DATA_SAIDA)%></strong></td>
                </tr>
                <tr>
                  <td class="textoMedio">Razão social:</td>
                  <td class="textoMedio"><strong><%=RAZAO_SOCIAL%></strong></td>
                </tr>
                <tr>
                  <td class="textoMedio">Nome fantasia:</td>
                  <td class="textoMedio"><strong><%=NOME_FANTASIA%></strong></td>
                </tr>
                <tr>
                  <td class="textoMedio">CNPJ:</td>
                  <td class="textoMedio"><strong><%=CNPJ%></strong></td>
                </tr>
                <tr>
                  <td class="textoMedio">Endereço:</td>
                  <td class="textoMedio"><strong><%=ENDERECO%> - <%=BAIRRO%><BR /><%=CIDADE%>-<%=ESTADO%></strong></td>
                </tr>
                <tr>
                  <td class="textoMedio">Telefone:</td>
                  <td class="textoMedio"><strong><%=TELEFONE%></strong></td>
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
       
            <td style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style2">Descrição do Produto</div></td>
            <td style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style2">Quantidade</div></td>
            </tr>
		
						
						<%
						AbrirConexao()
						QTD_ITENS = 0
										
						Set rs =objConexao.Execute("SELECT COUNT(*) AS QTD_ITENS, COD_PRODUTO FROM TB_FORNECEDORES_SAIDA_ITENS WHERE  COD_FORNECEDOR_SAIDA = "&COD_FORNECEDOR_SAIDA&" GROUP BY COD_PRODUTO ORDER BY QTD_ITENS DESC")
						
						
						While not rs.Eof
						
						COD_PRODUTO_SELECIONADO = rs("COD_PRODUTO")
						QTD_ITENS				= QTD_ITENS + formatNumber(0&rs("QTD_ITENS"),0)
						
						%>
          					<tr>
                              <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333; border-left:solid 0.25px #333333"><div align="center">
                              <%
							  Set rs2=objConexao.Execute("SELECT * FROM TB_PRODUTOS WHERE COD_PRODUTO = "&COD_PRODUTO_SELECIONADO)
							  		
									If not rs2.Eof Then
											
											Response.Write(Ucase(rs2("NOME_PRODUTO")))
											
									End If
							  	
							  rs2.Close
							  Set rs2=Nothing
							  %>
                              </div></td>
                              <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333"><div align="center"><%=rs("QTD_ITENS")%></div></td>
            </tr>
          				<%	
						rs.MoveNext()
						Wend
						
						rs.Close
						Set rs = Nothing
						
						FecharConexao()
						%>
						<tr>
          					  <td colspan="2"><div align="center" class="style4">Quantidade total: <span class="style5"><%=QTD_ITENS%></span></div></td>
   					  </tr>
					</table>
	      
	      <p></p>
          <p></p>
          <p></p>
          <p></p>
	      <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
           <tr>
             <td style="border-bottom:solid 1px #000000">
               <p></p></td>
           </tr>
           <tr>
             <td><div align="center"><span class="textoMedio"><span class="style1">RESPONSÁVEL PELA RETIRADA</span><br />
                   <span class="style1">  <strong><%=NOME_RESPONSAVEL_SAIDA%><br />RG: <%=RG_RESPONSAVEL_SAIDA%></strong></span></span></div></td>
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
		<!-- #include file="includes/rodape_impressao.asp" -->
	
	
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



