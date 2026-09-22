<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim COD_PRODUCAO
COD_PRODUCAO = Request.QueryString("COD_PRODUCAO")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="print"></link>
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<style type="text/css">
<!--
.btnFinalizaProducao{	background-image:url(images/layout/btnFinalizaProducao.jpg);
	cursor:pointer;
	width:140px;
	height:22px;
	border:0px;
}
.folha {
   page-break-after: always;
}
.style1 {color:#CC0000; font-weight: bold;}
-->
</style>

<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
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
				var text = "Se a janela não abrir desbloquei os Pop-ups!";
				if(win == null) { alert(text); return; }
				win.moveTo(350,150);
					
			}
			else {
				if (msgErro==1){alert('Para visualizar os brindes é necessário alterar a opção Brinde para Sim.')}
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
				parent.location.href = 'dip_producoes_fichas_recibos_vendedores.asp?msg=Liberação concluída com sucesso!&cod_prod=<%=COD_PRODUCAO%>'
	}else{
			document.getElementById("imprimir").style.display = "block";
			document.getElementById("imprimir").style.visibility = "visible";			
			alert('não FECHE ESTÁ TELA E CONTATE O DEPTO. DE T.I!');
			//Faz a barra de impressão novamente e o alert
			
	}
	
}
function fecharJanela(){
parent.location.href = 'dip_producoes_fichas_recibos_vendedores.asp?msg=2° Via impressa com sucesso!&cod_vend=<%=COD_VENDEDOR%>&cod_prod=<%=Request.QueryString("cod_prod")%>'
}
</script>
</head>

<body style="background:#FFFFFF" onload="window.print()">
<%
AbrirConexao()
					'//////////////////
					Dim QTD_BRINDES_EXTRAS
					QTD_BRINDES_EXTRAS = 0
					'Verifica se há brindes para está produção e vendedor
					sql = "SELECT COD_FORMANDO, BRINDE, COD_CONTRATO, ALBUM, NOME_FORMANDO, COD_PRODUCAO_desconto FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = "&COD_PRODUCAO&" AND BRINDE = 1 ORDER BY Convert(ALBUM,SIGNED)"
					Dim qtdBrindesTotaisSQL
					qtdBrindesTotaisSQL = formatNumber(0&qtdBrindesTotaisSQL,0)
					Set rsRelat = objConexao.Execute(sql,qtdBrindesTotaisSQL)
						
						classe = "class='folha'"
						While not rsRelat.Eof
						qtdBrindesTotaisSQL = qtdBrindesTotaisSQL -1
						
						If qtdBrindesTotaisSQL > 1 Then
							If classe = "class='folha'" Then
								classe = ""
							Else 
								classe = "class='folha'"	
							End If
						Else
							classe = ""
						End If	
						'Existem brindes para esta produção
							BRINDE			= rsRelat("BRINDE")
							COD_FORMANDO	= rsRelat("COD_FORMANDO")
													
					
					%>
					<table <%=classe%> width="100%" align="center" border="0" cellpadding="0" cellspacing="10">
					<tr>
					<td>
					<table width="100%" border="0" align="center" cellpadding="10" cellspacing="0" bgcolor="#FFFFFF">
					  <tr>
						<td><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0">
						  <tr>
							<td valign="top"><h2 style="color:#000000; border-bottom:solid #000 3px; padding-bottom:2px">Declara&ccedil;&atilde;o de Brinde - Contrato n&ordm;  <%=rsRelat("COD_CONTRATO")%> - &Aacute;lbum - <%=rsRelat("ALBUM")%></h2>
							  <br />
							  <table width="100%" border="0" cellspacing="0" cellpadding="0">
								<tr>
								  <td><p align="center" class="textoMedio">Eu, <strong><%=rsRelat("NOME_FORMANDO")%></strong>, estou recebendo da S&rsquo;tillo&rsquo;s Retortagens Fotogr&aacute;ficas o(s) seguinte(s) brindes:</p>
									<p><%
										
										sql = "SELECT TB_PRODUTOS.NOME_PRODUTO,TB_PRODUTOS.TIPO, Sum( TB_FICHAS_IDENTIFICACAO_BRINDES.QUANTIDADE ) AS QUANTIDADE"&_
																	" FROM TB_PRODUTOS"&_
																	" INNER JOIN TB_FICHAS_IDENTIFICACAO_BRINDES ON TB_PRODUTOS.COD_PRODUTO = TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO"&_
																	" WHERE TB_FICHAS_IDENTIFICACAO_BRINDES.COD_FORMANDO = "&COD_FORMANDO&_
																	" GROUP BY TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO ORDER BY TB_PRODUTOS.NOME_PRODUTO"
										
										
										Set rsBrinde = objConexao.Execute(sql)
					
										If not rsBrinde.Eof then%>
										<table width="400" border="1" align="center" 	cellpadding="2" style="border-collapse:collapse">
																	 <tr bgcolor="whitesmoke">
												<td width="14%"><div align="center"><strong>Qtd.</strong></div></td>
												<td width="86%"><div align="center"><strong>Produto</strong></div></td>
											  </tr>
												
											  <%While not rsBrinde.Eof%>
											  <tr>
												<td width="14%">
												  <div align="center">
													<%
												If rsBrinde("TIPO") = 1 Then
													Response.Write("1")
												Else
													Response.Write(rsBrinde("QUANTIDADE"))
												End If
												%>
												  </div></td>
												<td width="86%"><div align="left">
												<%If rsBrinde("TIPO") = 1 Then
													Response.Write("Álbum contendo "&rsBrinde("QUANTIDADE")&" fotos ("&rsBrinde("NOME_PRODUTO")&")")
												%>
												<%Else%>
													<%=rsBrinde("NOME_PRODUTO")%>
														<%If rsBrinde("NOME_PRODUTO") = "DVD" Then%> 
												(será entregue posteriormente conforme o prazo do pedido)
														<%End If%>
												<%End If%>
												</div></td>
											  </tr>
											  <%
											  rsBrinde.MoveNext()
											  Wend
											  %>
											</table>

											<%
											if rsRelat("COD_PRODUCAO_desconto") <> "" then
												set rsDesconto = objConexao.Execute("SELECT * FROM TB_PRODUCAO_DESCONTOS WHERE COD_PRODUCAO_desconto ="&rsRelat("COD_PRODUCAO_desconto"))
												if not rsDesconto.Eof Then
													Response.write "<center><strong>"&rsDesconto("DESCRICAO")&"</strong></center>"
												end if
											end if
											%>
										<div align="center">
										  <%Else

										if rsRelat("COD_PRODUCAO_desconto") <> "" then
											set rsDesconto = objConexao.Execute("SELECT * FROM TB_PRODUCAO_DESCONTOS WHERE COD_PRODUCAO_desconto ="&rsRelat("COD_PRODUCAO_desconto"))
											if not rsDesconto.Eof Then
												Response.write "<center><strong>"&rsDesconto("DESCRICAO")&"</strong></center>"
											end if

										else
										%>
											<span class="style1">Brinde não encontrado, contate o administrador do sistema!</span></span>
										<%
										end if
										  %>



											<%End if	
											
										rsBrinde.Close()		
										Set rsBrinde = Nothing
							
								%>
										  </p>
											</div>
							
									<p>&nbsp;</p>
									<p align="center">
									<table style="border-top:#000000 1px solid" width="300" border="0" align="center" cellpadding="4" cellspacing="0">
					  <tr>

						<td class="textoMedio"><div align="left">Nome: <strong><%=rsRelat("NOME_FORMANDO")%></strong><br />
						  RG:<br />
						  Data Recebida: 
						  </div></td>
					  </tr>
					</table>
					
									<p>&nbsp;</p>
									</p></td>
								  </tr>
							  </table></td>
							</tr>
						  <tr>
							<td>
						
							 </td>
							</tr>
					
					
						</table></td>
					  </tr>
					</table>
					</td>
					</tr>
					</table>
							<%
							
							
							
						
						rsRelat.MoveNext()
						Wend
						
						rsRelat.Close
						Set rsRelat = Nothing
						
					%>
						</td>
					  </tr>
					</table>
		<%
		'FecharConexao2()
		FecharConexao()
		%>	
</body>
</html>