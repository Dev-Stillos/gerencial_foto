<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Validação especial para verificar se possui sub-validação específica
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '45' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se não encontrar então redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
		
	rs.Close
	Set rs=Nothing
	
FecharConexao()


Dim sql, rs, rs2
Dim total_albuns, total_albuns_individual, total_arrecadado, total_arrecadado_individual, vendasCoordenador(4,2)
Dim VENDEDOR_RECEBE_EM, ANO, MES

ANO	= ""&Request.QueryString("ANO")
MES = ""&Request.QueryString("MES")

If ANO = "" AND MES = "" Then
	Response.Redirect("adm_relatorios.asp?msg=É necessário informar o mês/ano para geração do relatório.")
Else
	VENDEDOR_RECEBE_EM		= formatDateTime(ANO&"-"&MES&"-01")
End If

sql = "SELECT TB_PEDIDOS.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR, TB_PEDIDOS.COD_GRUPO_VENDEDORES, COUNT(VW_PEDIDOS_PRODUTOS_TOTAL_ALBUNS.XTOTAL_ALBUNS) AS TOTAL_ALBUNS, SUM(VW_PEDIDOS_TOTAL_RECEBIDO.XTOTAL) AS TOTAL_VENDAS  FROM TB_PEDIDOS INNER JOIN VW_PEDIDOS_TOTAL_RECEBIDO ON (TB_PEDIDOS.COD_PEDIDO = VW_PEDIDOS_TOTAL_RECEBIDO.COD_PEDIDO) INNER JOIN TB_VENDEDORES ON (TB_PEDIDOS.COD_VENDEDOR = TB_VENDEDORES.COD_VENDEDOR) INNER JOIN VW_PEDIDOS_PRODUTOS_TOTAL_ALBUNS ON (TB_PEDIDOS.COD_PEDIDO = VW_PEDIDOS_PRODUTOS_TOTAL_ALBUNS.COD_PEDIDO) WHERE (TB_PEDIDOS.VENDEDOR_RECEBE_EM = "&sqlData(VENDEDOR_RECEBE_EM)&") AND TB_PEDIDOS.COD_GRUPO_VENDEDORES > 0 AND TB_PEDIDOS.FINALIZA_CORRECAO = 1 GROUP BY TB_PEDIDOS.COD_VENDEDOR ORDER BY TOTAL_VENDAS DESC"

total_albuns=0
total_albuns_individual=0
total_arrecadado=0
total_arrecadado_individual=0

vendasCoordenador(1,0) = 0
vendasCoordenador(1,1) = 0
vendasCoordenador(2,0) = 0
vendasCoordenador(2,1) = 0
vendasCoordenador(3,0) = 0
vendasCoordenador(3,1) = 0

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
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

function focusFormando(){
	document.producao.<%="ALBUM_"&Request.QueryString("cod_form")%>.focus()
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
	<script language="javascript">
	function confirmaFinalizarProducao(){
	
	var pergunta = confirm('Tem certeza que deseja finalizar está produção?\nObs.: Esta produção não poderá ser alterada posteriormente somente com senha especial.')
		if(pergunta==true){
		abre('?cod_prod=<%=Request.QueryString("cod_prod")%>&comando=finalizar_producao')
		}
	}
	function confirmaExclusao(){
	var perguntaq = confirm('EXCLUIR PRODUÇÃO:\nAtenção caso queira continuar lembre-se que todas as fichas(se houver) serão excluídas completamente, deseja continuar?')
		if (perguntaq==true){
			var perguntadeNovo = confirm("CONFIRMA EXCLUSÃO DESTA PRODUÇÃO E SUAS RESPECTIVAS FICHAS?")
			if (perguntadeNovo==true){
			abre('?cod_prod=<%=Request.QueryString("cod_prod")%>&comando=excluir_producao')
			}
		}
	
	}
	</script>
</head>

<body>

<table cellpadding="9px" cellspacing="0" width="100%">
<tr>
<td>
  <!-- #include file="includes/topo.asp" -->
  <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td><table width="100%" style="border-bottom:#999999 solid dotted 1px" background="images/layout/fundo_menu_azul.jpg"border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
	<!-- #include file="includes/menu.asp" -->	  </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="10">
  <tr>
  <td>
   <%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 2
  BTN_IMPRIMIR = 2
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">
	
	  <h2>VENDAS DE REPASSES</h2>

		  
		<p> </p>
		
        
        <h3 align="center">Ranking Geral de Vendas do Grupo de Repasse<br />
          Dados com base no pagamento do m&ecirc;s de "<%=Ucase(MonthName(Month(VENDEDOR_RECEBE_EM)))%>/<%=Year(VENDEDOR_RECEBE_EM)%>"</h3>
       <%
	   AbrirConexao()
	   Set rs=objConexao.Execute(sql)
	   
	   If not rs.Eof Then
	   %>
	    <table width="90%" border="1" align="center" 	cellpadding="7" style="border-collapse:collapse">
			        <tr bgcolor="whitesmoke">
                    	<td width="7%"><div align="center"><strong>Posi&ccedil;&atilde;o</strong>
                              
                  	  </div></td>
                        <td width="29%"><div align="center"><strong>Vendedor Repasse</strong></div></td>
                        <td width="29%"><div align="center"><strong>Coordenador</strong></div></td>
			            <td width="17%"><div align="center"><strong>Total &Aacute;lbuns</strong></div></td>
			            <td width="18%"><div align="center"><strong>Total Arrecadado</strong> <strong>(R$)</strong></div></td>
			        </tr>
                    <%While not rs.Eof
						
					   i=i+1
					   total_albuns_individual = formatNumber(0&rs("TOTAL_ALBUNS"))
					   total_albuns = (total_albuns + total_albuns_individual)/1
					   
					   total_arrecadado_individual = formatNumber(0&rs("TOTAL_VENDAS"))
					   total_arrecadado = (total_arrecadado + total_arrecadado_individual)/1
					   
					  If int(rs("COD_GRUPO_VENDEDORES")) = 1 Then
					  	'Renato
					  	vendasCoordenador(1,0) = (vendasCoordenador(1,0)+total_albuns_individual)/1
						vendasCoordenador(1,1) = (vendasCoordenador(1,1)+total_arrecadado_individual)/1
					  Elseif int(rs("COD_GRUPO_VENDEDORES")) = 2 Then
					  	'Mário
						vendasCoordenador(2,0) = (vendasCoordenador(2,0)+total_albuns_individual)/1
						vendasCoordenador(2,1) = (vendasCoordenador(2,1)+total_arrecadado_individual)/1
					  Elseif int(rs("COD_GRUPO_VENDEDORES")) = 3 Then
					  	'Paulo Mantuam
						vendasCoordenador(3,0) = (vendasCoordenador(3,0)+total_albuns_individual)/1
						vendasCoordenador(3,1) = (vendasCoordenador(3,1)+total_arrecadado_individual)/1
            Elseif int(rs("COD_GRUPO_VENDEDORES")) = 4 Then
					  	'Vicente
						vendasCoordenador(4,0) = (vendasCoordenador(4,0)+total_albuns_individual)/1
						vendasCoordenador(4,1) = (vendasCoordenador(4,1)+total_arrecadado_individual)/1
					  End If
					%>
			        <tr>
			          <td><div align="center" class="textoMedio"><%=i%>º</div></td>
			          <td><div align="left"><%=rs("nome_vendedor")%></div></td>
			          <td><div align="left">
					  <%
					  If int(rs("COD_GRUPO_VENDEDORES")) = 1 Then
					  	Response.Write("RENATO LUÍZ ALBA")
					  Elseif int(rs("COD_GRUPO_VENDEDORES")) = 2 Then
					  	Response.Write("MÁRIO DE SANTIS")
					  Elseif int(rs("COD_GRUPO_VENDEDORES")) = 3 Then
					  	Response.Write("PAULO MANTUAM")
            Elseif int(rs("COD_GRUPO_VENDEDORES")) = 4 Then
					  	Response.Write("VICENTE")
					  End If
					  %></div></td>
			          <td><div align="center">
                      <%=formatNumber(total_albuns_individual,0)%>
                      </div></td>
			          <td><div align="right"><%=formatNumber(total_arrecadado_individual)%></div></td>
			          </tr>
			        
                    <%
					rs.MoveNext
					Wend
					
					%>
                    <tr bgcolor="whitesmoke">
			          <td colspan="3">&nbsp;</td>
			          <td><div align="center" class="textoMedio"><%=total_albuns%></div></td>
			          <td><div align="right" class="textoMedio"><%=formatNumber(total_arrecadado)%></div></td>
			          </tr>
        </table>
        <br />
         <table width="90%" border="1" align="center" 	cellpadding="7" style="border-collapse:collapse">
			        <tr bgcolor="whitesmoke">
                    	<td colspan="3"><div align="center"><strong class="textoMedio">Total Vendas por Coordenador:</strong></div></td>
                    </tr>
			        <tr bgcolor="whitesmoke">
			          <td width="40%"><div align="center"><strong>Coordenador</strong></div></td>
			          <td width="26%"><div align="center"><strong>Total &Aacute;lbuns</strong></div></td>
			          <td width="34%"><div align="center"><strong>Total Arrecadado</strong></div></td>
			        </tr>
			        <tr>
			          <td><div align="left"><strong>RENATO LU&Iacute;Z ALBA</strong></div></td>
			          <td><div align="center"><strong><%=vendasCoordenador(1,0)%></strong></div></td>
			          <td><div align="right"><strong><%=formatNumber(vendasCoordenador(1,1))%></strong></div></td>
			          </tr>
			        <tr>
			          <td><div align="left"><strong>MÁRIO DE SANTIS</strong></div></td>
			          <td><div align="center"><strong><%=vendasCoordenador(2,0)%></strong></div></td>
			          <td><div align="right"><strong><%=formatNumber(vendasCoordenador(2,1))%></strong></div></td>
			          </tr>
         </table>
      <div style="width:90%"></div>
       <%
	   Else
	   %>
     <div align="center" style="padding:50px"><strong> Não foram encontradas vendas para o período informado!</strong></div>
       <%End If
	   
	   FecharConexao()
	   %>
       <br /><br />
       Observação: Ranking gerado com base no mês de recebimento dos pedidos dos vendedores.
	    </td>
    </tr>
</table>

  
  
  
  </td>
    </tr>
</table>  </td>
  </tr>
</table>
<!-- #include file="includes/rodape.asp" -->

  </td>
</tr>
</table>
</body>
</html>



