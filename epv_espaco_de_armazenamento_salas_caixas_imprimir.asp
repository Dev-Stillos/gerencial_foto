<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<%
Dim COD_CAIXA_ARMAZENADO, NOME_SALA_ARMAZENADO, ANDAR, COD_SALA_ARMAZENADO, COLUNA, NUMERO_CAIXA
Dim COD_CAIXA_PADRAO
COD_CAIXA_ARMAZENADO = Request.QueryString("COD_CAIXA_ARMAZENADO")

AbrirConexao()

sql = "SELECT TB_LOCAL_ARMAZENAGEM_SALA.COD_SALA_ARMAZENADO,TB_LOCAL_ARMAZENAGEM_SALA_CAIXA.COLUNA, TB_LOCAL_ARMAZENAGEM_SALA.NOME_SALA_ARMAZENADO, TB_LOCAL_ARMAZENAGEM_SALA.ANDAR, TB_LOCAL_ARMAZENAGEM_SALA_CAIXA.NUMERO_CAIXA  FROM TB_LOCAL_ARMAZENAGEM_SALA INNER JOIN TB_LOCAL_ARMAZENAGEM_SALA_CAIXA ON (TB_LOCAL_ARMAZENAGEM_SALA.COD_SALA_ARMAZENADO = TB_LOCAL_ARMAZENAGEM_SALA_CAIXA.COD_SALA_ARMAZENADO) WHERE TB_LOCAL_ARMAZENAGEM_SALA_CAIXA.COD_CAIXA_ARMAZENADO = "&COD_CAIXA_ARMAZENADO

Set rs = objConexao.Execute(sql)
		
		If not rs.Eof Then
			COD_SALA_ARMAZENADO		= rs("COD_SALA_ARMAZENADO")
			NOME_SALA_ARMAZENADO	= rs("NOME_SALA_ARMAZENADO") 
			ANDAR					= formatNumber(0&rs("ANDAR"),0)
			COLUNA					= rs("COLUNA")
			NUMERO_CAIXA			= rs("NUMERO_CAIXA")
		End If
		
				 If ANDAR = 1 Then
				  	ANDAR = "(1) - Sub-Solo"
				  Elseif ANDAR = 2 Then
				  	ANDAR ="(2) - Térreo"
				  Elseif ANDAR = 3 Then
				  	ANDAR ="(3) - 1º Andar"
				  End If
		
rs.Close
Set rs = Nothing

QTD_ITENS = 0
QTD_FOTOS_TOTAL = 0
				  
				  
				  
				  
'////////////////////////////////////////////////

If Len(NUMERO_CAIXA) = 1 Then
	NUMERO_CAIXA = "000000"&NUMERO_CAIXA
Elseif Len(NUMERO_CAIXA) = 2 Then
	NUMERO_CAIXA = "00000"&NUMERO_CAIXA
Elseif Len(NUMERO_CAIXA) = 3 Then
	NUMERO_CAIXA = "0000"&NUMERO_CAIXA
Elseif Len(NUMERO_CAIXA) = 4 Then
	NUMERO_CAIXA = "000"&NUMERO_CAIXA
Elseif Len(NUMERO_CAIXA) = 5 Then
	NUMERO_CAIXA = "00"&NUMERO_CAIXA
Elseif Len(NUMERO_CAIXA) = 6 Then
	NUMERO_CAIXA = "0"&NUMERO_CAIXA
End If
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
-->
</style>

<title>StillosWEB - Sistema corporativo - Versão 1.2</title>
	<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>
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

<style type="text/css">
<!--
.style2 {
	font-size: 10px;
	font-weight: bold;
}
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
        <td valign="top"><h2 style="color:#000000; border-bottom:solid #000 3px; padding-bottom:2px">Relação de Itens na Caixa </h2>
          <br />
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="90%" style="padding-bottom:10px"><table width="100%" border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td width="256" class="textoMedio">Data da Impressão:</td>
                  <td width="801"><strong class="textoMedio"><%=ArrumaData(Now())%></strong></td>
                </tr>
                <tr>
                  <td class="textoMedio">Coluna / N&uacute;mero da Caixa:</td>
                  <td class="textoMedio"><strong><%=COLUNA%> / <%=NUMERO_CAIXA%></strong></td>
                </tr>
                <tr>
                  <td class="textoMedio">Sala / Andar:</td>
                  <td class="textoMedio"><strong>
				  <%Response.Write(COD_SALA_ARMAZENADO&" - "&NOME_SALA_ARMAZENADO&" / "&ANDAR)%>                  
				  </strong></td>
                </tr>
              </table></td>
              <td width="10%"><div align="center"></div></td>
            </tr>
          </table></td>
        </tr>
      <tr>
        <td> 
        <table width="100%" align="center">
        <tr>  
        <%		
		Set rs=objConexao.Execute("SELECT * FROM VW_TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_COMPLETA WHERE COD_CAIXA_ARMAZENADO = "&COD_CAIXA_ARMAZENADO&" AND DATA_RETORNO IS NOT NULL AND TIPO <> 1 AND TRANSFERIDO <> 4 ORDER BY COD_FORMANDO",QTD_ITENS)
		
		
		If not rs.Eof Then
		
		
		Dim COLUNAS, LINHAS, RESTO
		
		If QTD_ITENS > 10 Then
			COLUNAS = 3
		Elseif QTD_ITENS > 5 Then
			COLUNAS = 2
		Else
			COLUNAS = 1
		End If
		
		LINHAS = QTD_ITENS/COLUNAS
		RESTO = QTD_ITENS Mod COLUNAS
		
		For i=0 to COLUNAS-1
		%>
        <td style="padding:10px" valign="top">
		       <table width="320" border="0" align="center" 	cellpadding="1" cellspacing="0">
			<tr bgcolor="whitesmoke">
       
            <td width="19%" style="border-bottom:solid 2px #000"><div align="center" class="style2">Cód. Formando </div>              </td>
            <td width="11%" style="border-bottom:solid 2px #000"><div align="center">Álbum</div></td>
            <td width="11%" style="border-bottom:solid 2px #000"><div align="center">Qtd. Fotos Inicial</div></td>
            <td width="15%" style="border-bottom:solid 2px #000"><div align="center" class="style2">Contrato</div></td>
            <td width="55%" style="border-bottom:solid 2px #000"><div align="center" class="style2">Item</div><div align="center" class="style2"></div></td>
            </tr>
			<%
			If int(i)=(COLUNAS-1) Then
				LINHAS = LINHAS + RESTO
			End If
			
			For ii=0 to LINHAS-1
			%>							
								<%
						If not rs.Eof Then
						
						COD_FORMANDO		= rs("COD_FORMANDO")
						TIPO				= formatNumber(0&rs("TIPO"),0)
						
						Set rs2=objConexao.Execute("SELECT ENCADERNACOES_EXTRAS,INICIAL, COD_CONTRATO, ALBUM FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO = "&COD_FORMANDO)
						
							
							If not rs2.Eof Then
									ENCADERNACOES_EXTRAS	= formatNumber(0&rs2("ENCADERNACOES_EXTRAS"),0)
									COD_CONTRATO			= rs2("COD_CONTRATO")		
									ALBUM					= rs2("ALBUM")	
									INICIAL					= rs2("INICIAL")						
							End If
							
						rs2.Close
						Set rs2=Nothing

						QTD_FOTOS_TOTAL = QTD_FOTOS_TOTAL + INICIAL
						%>
          					<tr>
            <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333; border-left:solid 0.25px #333333"><%=digito(rs("COD_FORMANDO"))%></td>
            <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333; border-left:solid 0.25px #333333"><div align="center"><%=ALBUM%></div></td>
            <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333; border-left:solid 0.25px #333333"><div align="center"><%=INICIAL%></div></td>
            <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333; border-left:solid 0.25px #333333"><div align="center"><%=COD_CONTRATO%></div></td>
            <td style="border-bottom:solid 0.25px #333333; border-right:solid 0.25px #333333"><div align="center">
            
            <div align="center">
			<%If int(TIPO) = 0 Then%>
				<%If int(ENCADERNACOES_EXTRAS) = 1 Then%>
					Álbum Duplo
<%Else%>
Álbum
<%End If%>
			<%Else
			'Produto agregado, consultar qual é.
			Dim rsProdAgrePend
			sql = 	"SELECT TB_PRODUTOS.NOME_PRODUTO "&_
					"FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = "&_
					"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO WHERE "&_
					"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO = "&COD_FORMANDO&" AND "&_
					"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.TIPO = "&TIPO&" LIMIT 1"
			Set rsProdAgrePend = objConexao.Execute(sql)
				If not rsProdAgrePend.Eof Then
				%>
				<%=rsProdAgrePend("NOME_PRODUTO")%>
			    <%
				End If
			rsProdAgrePend.Close
			Set rsProdAgrePend = Nothing
			End If%>
			     </div></div></td>
            </tr>
						<%
						End If
						%>
			<%
			rs.MoveNext()
			Next
			%>
            </table>
                    
                    </td>
	    <%
		
		Next
		
		Else
			Response.Redirect("epv_espaco_de_armazenamento_salas.asp?msg=Não foi encontrado itens a serem impressos")
		End If
		%>  
          </tr>
          </table>
	      <p>
          Total Itens Armazenados na Caixa: <strong><%=QTD_ITENS%></strong><br>
          Total de Fotos Armazenadas na Caixa: <strong><%=QTD_FOTOS_TOTAL%></strong> 
          <br />
	        </p>
	      <p> </p>
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
</body>
</html>
<%
		FecharConexao()
%>



