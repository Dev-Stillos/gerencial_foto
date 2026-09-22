<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Valida somente o usu�rio
COD_USUARIO_INTERNO_LOGADO 	= Session("COD_USUARIO")
COD_PERFIL_LOGADO			= Session("COD_PERFIL")

'================ INÍCIO DO BLOCO DE FILTRO ========================
Dim MES_INICIO, ANO_INICIO, MES_FIM, ANO_FIM

If IsNumeric(Request("MES_INICIO")) And CInt(Request("MES_INICIO")) >= 1 And CInt(Request("MES_INICIO")) <= 12 Then
	MES_INICIO = CInt(Request("MES_INICIO"))
Else
	MES_INICIO = Month(Date)
End If

If IsNumeric(Request("ANO_INICIO")) Then
	ANO_INICIO = CInt(Request("ANO_INICIO"))
Else
	ANO_INICIO = Year(Date)
End If

If IsNumeric(Request("MES_FIM")) And CInt(Request("MES_FIM")) >= 1 And CInt(Request("MES_FIM")) <= 12 Then
	MES_FIM = CInt(Request("MES_FIM"))
Else
	MES_FIM = Month(Date)
End If

If IsNumeric(Request("ANO_FIM")) Then
	ANO_FIM = CInt(Request("ANO_FIM"))
Else
	ANO_FIM = Year(Date)
End If
'================ FIM DO BLOCO DE FILTRO ==========================


'Verifica se o usu�rio est� logado
if COD_USUARIO_INTERNO_LOGADO=Empty OR COD_PERFIL_LOGADO=Empty then
Session.Abandon()
Response.Redirect("index.asp")		
end if

'//ROTINA QUE VERIFICA SE H� PEDIDOS QUE N�O POSSUEM ORDEM DE SERVI�O GERADA POR ALGUM MOTIVO.
Server.Execute("adm_rotina_corrige_os_geracao.asp")

AbrirConexao()
'//ROTINA QUE VERIFICA SE H� OS FINALIZADAS, POR�M COM ETAPAS COM DATA EM ABERTO
objConexao.Execute("UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_termino = NOW( ) WHERE data_termino IS NULL AND cod_ordem_servico IN (SELECT cod_ordem_servico FROM TB_ORDENS_SERVICO WHERE status_ordem_servico = 3)")

'//ROTINA QUE CORRIGE SE H� FORMANDOS VENDIDOS E QUE CONSTAM COMO N�O ENTREGUES
Set rs=objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE TIPO = 0 AND DATA_RETORNO IS NULL AND COD_FORMANDO IN (SELECT COD_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE VENDIDO = 1)")

	While not rs.Eof		
		Set rs2=objConexao.Execute("SELECT DATA_VENDA FROM TB_PEDIDOS WHERE FINALIZA_CORRECAO = 1 AND COD_PEDIDO IN (SELECT cod_pedido FROM TB_PEDIDOS_FORMANDOS WHERE COD_FORMANDO = "&rs("COD_FORMANDO")&")")

			If not rs2.Eof Then
				sql = "UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET DATA_RETORNO = "&sqlData(rs2("DATA_VENDA"))&", COBRAR_DEVOLUCAO = 2 WHERE COD_MOVIMENTACAO = "&rs("COD_MOVIMENTACAO")
				objConexao.Execute(sql)
			End If
		rs2.Close	
		Set rs2=Nothing
	rs.MoveNext()
	Wend
	
rs.Close
Set rs=Nothing
'//


'//ROTINA QUE CORRIGE SE H� ITENS ANTIGOS QUE CONTAM EM ABERTO, POR�M O MESMO ITEM MAIS RECENTE EST� DEVOLVIDO. COMO � UM ALGORITMO MUITO IMPORTANTO OS REGISTROS ALTERADOS EST�O SENDO MARCARDOS.
AbrirConexao()

Set rs=objConexao.Execute("SELECT COD_MOVIMENTACAO AS a, COD_FORMANDO b, TIPO AS c, DATA_SAIDA AS d, DATA_RETORNO AS e FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES z WHERE COBRAR_DEVOLUCAO = 1 AND DATA_RETORNO IS NULL AND DATA_SAIDA < (SELECT MAX(DATA_SAIDA) FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_FORMANDO = z.COD_FORMANDO AND TIPO = z.TIPO) GROUP BY COD_FORMANDO, TIPO ORDER BY COD_FORMANDO ASC")

	While not rs.Eof
		
		'Atualiza o retorno com a pr�pria data de sa�da
		sql = "UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET DATA_RETORNO = "&sqlData(rs("d"))&", COBRAR_DEVOLUCAO = 2, CORRECAO_ALGORITMO_DATA = "&sqlData(now())&" WHERE COD_MOVIMENTACAO = "&rs("a")
		objConexao.Execute(sql)
		
	rs.MoveNext()
	Wend
	
rs.Close
Set rs=Nothing
'//


'//ROTINA QUE CORRIGE SE H� ITENS COMPRADOS, POR�M N�O H� MAIS OS PARA O ITEM E CANCELA
'//SOMENTE SE O PRODUTO N�O FOR ORIUNDO DE PEDIDO MANUAL'
AbrirConexao()

'Set rs=objConexao.Execute("SELECT a.cod_pedido_compra_produto, b.cod_ordem_servico FROM TB_PEDIDOS_COMPRAS_PRODUTOS a LEFT JOIN  TB_ORDENS_SERVICO_ETAPAS b ON (a.cod_ordem_servico_etapa = b.cod_ordem_servico_etapa) WHERE b.cod_ordem_servico IS NULL AND a.pedido_manual <> 'S' ")

''	While not rs.Eof

		'Atualiza o retorno com a pr�pria data de sa�da
''		sql = "UPDATE TB_PEDIDOS_COMPRAS_PRODUTOS SET status_produto = 'C', obs='Cancelado, pois OS n�o existe mais' WHERE cod_pedido_compra_produto = "&rs("cod_pedido_compra_produto")
''		objConexao.Execute(sql)
''		
		'N�o se preocupa com status principal, pois este dever� ser cancelado manualmente.
		
''	rs.MoveNext()
''	Wend
	
'rs.Close
'Set rs=Nothing

'//ROTINA QUE CORRIGE SE H� ITENS ANTIGOS QUE CONTAM EM ABERTO, POR�M O MESMO ITEM MAIS RECENTE EST� DEVOLVIDO. COMO � UM ALGORITMO MUITO IMPORTANTE OS REGISTROS ALTERADOS EST�O SENDO MARCARDOS.

AbrirConexao()

Set rs=objConexao.Execute("SELECT COD_MOVIMENTACAO AS a, COD_FORMANDO b, TIPO AS c, DATA_SAIDA AS d, DATA_RETORNO AS e FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES z WHERE COBRAR_DEVOLUCAO = 1 AND DATA_RETORNO IS NULL AND DATA_SAIDA < (SELECT MAX(DATA_SAIDA) FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_FORMANDO = z.COD_FORMANDO AND TIPO = z.TIPO) GROUP BY COD_FORMANDO, TIPO ORDER BY COD_FORMANDO ASC")

	While not rs.Eof
		
		'Atualiza o retorno com a pr�pria data de sa�da
		sql = "UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET DATA_RETORNO = "&sqlData(rs("d"))&", COBRAR_DEVOLUCAO = 2, CORRECAO_ALGORITMO_DATA = "&sqlData(now())&" WHERE COD_MOVIMENTACAO = "&rs("a")
		objConexao.Execute(sql)

	rs.MoveNext()
	Wend
	
rs.Close
Set rs=Nothing
'//


'//ROTINA QUE CORRIGE SE H� OS's PENDENTES, POR�M DEVERIAM ESTAR FINALIZADAS, UMA VEZ QUE AS TODAS AS ETAPAS J� EST�O FINALIZADAS.
AbrirConexao()

'Set rs=objConexao.Execute("SELECT cod_ordem_servico, (SELECT COUNT(*) as qtd_etapas FROM TB_ORDENS_SERVICO_ETAPAS b WHERE b.cod_ordem_servico = a.cod_ordem_servico  GROUP BY b.cod_ordem_servico) as qtd_etapas_total, (SELECT COUNT(*) as qtd_etapas FROM TB_ORDENS_SERVICO_ETAPAS c WHERE  c.cod_ordem_servico = a.cod_ordem_servico AND c.data_termino IS NOT NULL GROUP BY c.cod_ordem_servico) as qtd_etapas_pendentes  FROM TB_ORDENS_SERVICO a  WHERE a.data_finalizado IS NULL AND (SELECT COUNT(*) as qtd_etapas FROM TB_ORDENS_SERVICO_ETAPAS b WHERE b.cod_ordem_servico = a.cod_ordem_servico  GROUP BY b.cod_ordem_servico) = (SELECT COUNT(*) as qtd_etapas FROM TB_ORDENS_SERVICO_ETAPAS c WHERE  c.cod_ordem_servico = a.cod_ordem_servico AND c.data_termino IS NOT NULL GROUP BY c.cod_ordem_servico)")

	'While not rs.Eof
		
		'Atualiza o retorno com a pr�pria data de sa�da	
		'sql = "UPDATE TB_ORDENS_SERVICO SET status_ordem_servico = 2, data_finalizado = NOW(), cod_usuario_finalizou = "&Session("cod_usuario")&" WHERE cod_ordem_servico = "&sqlTexto(rs("cod_ordem_servico"))
		'objConexao.Execute(sql)
		
		
	'rs.MoveNext()
	'Wend
	
'rs.Close
'Set rs=Nothing
'//



'//ROTINA QUE VERIFICA SE H� PEDIDOS DE COMPRAS FINALIZADOS = 3, por�m na realidade possuem itens em aberto
Set rs=objConexao.Execute("SELECT cod_pedido_compra FROM VW_TB_PEDIDOS_COMPRAS_PRODUTOS WHERE status_produto = 'P' AND status_pedido_compra = 3 GROUP BY cod_pedido_compra")

	While not rs.Eof
		
		'Atualiza o retorno com a pr�pria data de sa�da	
		sql = "UPDATE TB_PEDIDOS_COMPRAS SET status_pedido_compra = 2 WHERE cod_pedido_compra = "&sqlTexto(rs("cod_pedido_compra"))
		objConexao.Execute(sql)		
		
	rs.MoveNext()
	Wend
	
rs.Close
Set rs=Nothing


'//ROTINA PARA CORRIGIR STATUS DE PEDIDOS DE COMPRAS
objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS_PRODUTOS SET status_produto = 'P' WHERE quantidade_entregue IS NULL AND (status_produto IS NULL OR status_produto = '0')")	
objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS_PRODUTOS SET status_produto = 'E' WHERE quantidade_entregue IS NOT NULL AND (status_produto IS NULL OR status_produto = '0') AND status_produto <> 'P'")

FecharConexao()


'Server.Execute("adm_ROTINA_importa_entidades.asp")


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<style>
/* Estilo base para links nos menus */
td a,
td a:link,
td a:visited {
  color: black !important;
  text-decoration: none !important;
  background-color: inherit !important;
}

/* Hover aplicado a qualquer link dentro de tabelas ou células com fundo rosa ou escuro */
td[style*="background-color: #2f2f44"] a:hover,
td[style*="background-color: #ff007f"] a:hover,
table[style*="background-color: #2f2f44"] a:hover,
table[style*="background-color: #ff007f"] a:hover {
  color: #ffe6f0 !important;
  background-color: #cc0066 !important;
  text-decoration: none !important;
}
</style>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<style>
table[style*="background-color: #2f2f44;"] a,
table[style*="background-color: #2f2f44;"] a:link,
table[style*="background-color: #2f2f44;"] a:visited,
table[style*="background-color: #2f2f44;"] strong,
table[style*="background-color: #2f2f44;"] td {
  color: white !important;
  background-color: #2f2f44 !important;
  text-decoration: none !important;
}

table[style*="background-color: #2f2f44;"] a:hover,
table[style*="background-color: #ff007f;"] a:hover {
  color: #ffe6f0 !important;
  background-color: #cc0066 !important;
  text-decoration: none !important;
}

</style>

<title>Painel Administrativo</title>
	<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>
	<script type="text/javascript">
function confirma(cod_formando,nomeFormando,comando){
var mensagem

	if (comando == "alterar"){
		mensagem = "ATENÇÃO ALTERAÇÃO: \nDeseja resetar a ficha de: "+nomeFormando+" deseja continuar?\n*A ficha voltar� na forma original."	
	}	
	else {
		mensagem = "ATENÇÃO EXCLUS�O: \nSe voc� continuar poder� perder as altera��es feitas nesta p�gina,\nalem de resetar a ficha do formando: "+nomeFormando+"  deseja continuar?"
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
				var text = "Se a janela n�o abrir desbloquei os Pop-ups!";
				if(win == null) { alert(text); return; }
				win.moveTo(350,150);
					
			}
			else {
				if (msgErro==1){alert('Para visualizar os brindes � necess�rio alterar a op��o Brinde para Sim.')}
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

function focusFormando(){
	document.producao.<%="ALBUM_"&Request.QueryString("cod_form")%>.focus()
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
	<script language="javascript">
	function confirmaFinalizarProducao(){
	
	var pergunta = confirm('Tem certeza que deseja finalizar est� produ��o?\nObs.: Esta produ��o n�o poder� ser alterada posteriormente somente com senha especial.')
		if(pergunta==true){
		abre('?cod_prod=<%=Request.QueryString("cod_prod")%>&comando=finalizar_producao')
		}
	}
	function confirmaExclusao(){
	var perguntaq = confirm('EXCLUIR PRODU��O:\nATENÇÃO caso queira continuar lembre-se que todas as fichas(se houver) ser�o exclu�das completamente, deseja continuar?')
		if (perguntaq==true){
			var perguntadeNovo = confirm("CONFIRMA EXCLUS�O DESTA PRODU��O E SUAS RESPECTIVAS FICHAS?")
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
    <td>
<table width="100%" style="border-bottom: dotted 1px #999999; background-color: #2f2f44;" border="0" cellspacing="0" cellpadding="0">
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
  BTN_VOLTAR = 2
  BTN_IMPRIMIR = 2
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">

		<!-- CÓDIGOS DE USUÁRIOS AUTORIZADOS -->
		<%If Session("COD_USUARIO") = 1 Or Session("COD_USUARIO") = 16 Then%>

	
	  <!-- CAIXA ROSA CENTRALIZADA -->
<div style="margin: 20px auto; display: flex; justify-content: center;">
  <div style="background-color: #ff007f; color: white; padding: 12px 24px; border-radius: 8px; font-weight: bold; text-align: center; font-size: 16px;">
    Painel Administrativo | Resumo de Vendas
  </div>
</div>

<%
' Parâmetros recebidos do filtro

MES_INICIO = Request("MES_INICIO") : If MES_INICIO = "" Then MES_INICIO = Month(Date)
ANO_INICIO = Request("ANO_INICIO") : If ANO_INICIO = "" Then ANO_INICIO = Year(Date)
MES_FIM = Request("MES_FIM") : If MES_FIM = "" Then MES_FIM = Month(Date)
ANO_FIM = Request("ANO_FIM") : If ANO_FIM = "" Then ANO_FIM = Year(Date)

Dim TOTAL_CARTAO, TOTAL_CHEQUE, TOTAL_BOLETO, TOTAL_DINHEIRO, TOTAL
TOTAL_CARTAO = 0 : TOTAL_CHEQUE = 0 : TOTAL_BOLETO = 0 : TOTAL_DINHEIRO = 0 : TOTAL = 0

AbrirConexao()

Dim PERIODO_INICIAL, PERIODO_FINAL
PERIODO_INICIAL = formatDateTime("01/" & MES_INICIO & "/" & ANO_INICIO)
PERIODO_FINAL = formatDateTime(QuantidadeDiasMes("01/" & MES_FIM & "/" & ANO_FIM) & "/" & MES_FIM & "/" & ANO_FIM)

sql = "SELECT SUM(IFNULL(PAGAMENTO_EM_DINHEIRO,0)) as dinheiro, SUM(IFNULL(XTOTAL_CHEQUE,0)) as cheque, SUM(IFNULL(XTOTAL_BOLETO,0)) as boleto, SUM(IFNULL (XTOTAL_CARTAO,0)) as cartao, SUM(IFNULL (XTOTAL,0)) as total FROM VW_PEDIDOS_TOTAL_RECEBIDO WHERE cod_pedido IN (SELECT cod_pedido FROM TB_PEDIDOS WHERE FINALIZA_CORRECAO = 1 AND DATA_VENDA BETWEEN " & sqlData(PERIODO_INICIAL) & " AND " & sqlData(PERIODO_FINAL) & ")"

Set rs = objConexao.Execute(sql)
If Not rs.EOF Then
    TOTAL_DINHEIRO = FormatNumber(0 & rs("dinheiro"), 2)
    TOTAL_CHEQUE = FormatNumber(0 & rs("cheque"), 2)
	TOTAL_BOLETO = FormatNumber(0 & rs("boleto"), 2)
    TOTAL_CARTAO = FormatNumber(0 & rs("cartao"), 2)
    TOTAL = FormatNumber(0 & rs("total"), 2)
End If
rs.Close
Set rs = Nothing
FecharConexao()
%>

<!-- CARD VISUAL -->
<div style="margin: 20px auto; width: 80%; background: #fff; border-radius: 8px; padding: 20px; box-shadow: 0 0 10px #ccc;">
  
<div style="display: flex; justify-content: space-between; align-items: center;">
    <h2 style="margin: 0; color: #000000 !important;">Resumo de Vendas (Painel)</h2>


<form method="get" action="adm_index.asp" style="display:flex; gap:6px; align-items:center;">
  <select name="MES_INICIO">
    <%For i = 1 To 12
      If MES_INICIO = i Then selected="selected" Else selected=""
    %>
      <option value="<%=i%>" <%=selected%>><%=UCase(MonthName(i))%></option>
    <%Next%>
  </select>

  <select name="ANO_INICIO">
    <%For i = 2008 To Year(Date)+1
      If ANO_INICIO = i Then selected="selected" Else selected=""
    %>
      <option value="<%=i%>" <%=selected%>><%=i%></option>
    <%Next%>
  </select>

  <span style="margin: 0 6px;">at&eacute;</span>

  <select name="MES_FIM">
    <%For i = 1 To 12
      If MES_FIM = i Then selected="selected" Else selected=""
    %>
      <option value="<%=i%>" <%=selected%>><%=UCase(MonthName(i))%></option>
    <%Next%>
  </select>

  <select name="ANO_FIM">
    <%For i = 2008 To Year(Date)+1
      If ANO_FIM = i Then selected="selected" Else selected=""
    %>
      <option value="<%=i%>" <%=selected%>><%=i%></option>
    <%Next%>
  </select>

  <input type="submit" value="Filtrar" style="padding: 4px 12px; border: 1px solid #ccc; background-color: #ff007f; color: white; font-weight: bold; border-radius: 4px;">
</form>

</div>


  <!-- Período atual -->
  <p style="text-align:center; color:#666; margin-top: 10px; font-weight: bold; font-size: 16px">
    Per&iacute;odo: <%=UCase(MonthName(MES_INICIO))%> de <%=ANO_INICIO%> &agrave; <%=UCase(MonthName(MES_FIM))%> de <%=ANO_FIM%>
  </p>

  <!-- Tabela de valores -->
  <table width="100%" cellpadding="10" cellspacing="0" style="border-collapse: collapse; font-size: 14px;">
    <tr style="background-color: #f2f2f2; font-weight: bold; text-align: center;">
      <td>Forma de Pagamento</td>
      <td>Valor</td>
    </tr>
	<tr>
      <td style="text-align: left;">Boleto</td>
      <td style="text-align: right; color: #2c3e50;"><%=TOTAL_BOLETO%></td>
    </tr>
    <tr>
      <td style="text-align: left;">Cart&atilde;o de Cr&eacute;dito</td>
      <td style="text-align: right; color: #2c3e50;"><%=TOTAL_CARTAO%></td>
    </tr>
    <tr>
      <td style="text-align: left;">Cheque</td>
      <td style="text-align: right; color: #2c3e50;"><%=TOTAL_CHEQUE%></td>
    </tr>
    <tr>
      <td style="text-align: left;">Dinheiro</td>
      <td style="text-align: right; color: #2c3e50;"><%=TOTAL_DINHEIRO%></td>
    </tr>
    <tr style="background-color: #ff007f; color: white; font-weight: bold;">
      <td style="text-align: left;">Total Geral</td>
      <td style="text-align: right;"><%=TOTAL%></td>
    </tr>
  </table>
</div>
<% End If ' <-- Fim da verificação do perfil %>

			
            <%If Request.QueryString("codmsg") = 1 Then%>
            	<div class="msgGrandeErro">Voc&ecirc; n&atilde;o possui acesso a altera&ccedil;&atilde;o de pedidos ap&oacute;s baixa da Diretoria. Solicite a Diretoria para alterar o pedido</div>
            <%End If%>
            
            
		<h4 align="center">Bem vindo(a) <%=Session("NOME")%>!<br />
        
        
		  <%=Session("DEPARTAMENTO")%></h4>
		  
           <%
            if Session("COD_PERFIL") = 12 then
        %>
            <style>
                .bt_projetos{
                    margin: auto;
                    width: 230px;
                    display: block;
                    text-align: center;
                    padding-top: 20px;
                    padding-bottom: 20px;
                    border: 1px solid #808080;
                    background-color: #B4D0F5;
                    font-size: 18px;
                    color: #383030;
                    text-decoration: none;
                }
                .bt_projetos:hover{
                    background-color: #88B6F1;
                }
            </style>
            <a href="ple_projetos.asp" class="bt_projetos"> Visualizar projetos - TI </a>
        <%
            end if
        %>
        
        
		<%If Session("ACESSO_EM") <> Empty then%>
		<p align="center">&Uacute;ltimo acesso em: <%=Session("ACESSO_EM")%>. <br />Total de acessos: <%=Session("QTDACESSOS")%></p>
		<%else%>
		<p align="center">Este &eacute; seu 1&ordm; acesso! A qualquer momento pressione F1 para ajuda nas telas. </p>
		<%end if%>
		<p>
	  </p>
	  		<p align="center"></p>
		<p align="center"></p>
		<p align="center"></p>
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






