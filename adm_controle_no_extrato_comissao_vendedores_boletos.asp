<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/topo.asp" -->
<!-- #include file="includes/sub_menu.asp" --> 
 <%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 2
  BTN_FONTE = 1
  BTN_IMPRIMIR_LINK = "#"
  %>
<%
'================= FILTROS =====================
Dim filtro_vendedor, filtro_status, filtro_cod_fechamento, sql, corStatus, rs

filtro_vendedor = Request("cod_vendedor")
filtro_status = Request("status")
filtro_cod_fechamento = Request("COD_FECHAMENTO_PERIODO")

AbrirConexao()

' ================== CARREGA VENDEDORES DO FECHAMENTO =====================
Dim sqlVend, rsVend
sqlVend = "SELECT DISTINCT F.COD_VENDEDOR, V.NOME_VENDEDOR " & _
          "FROM TB_FECHAMENTO_PAGAMENTO F " & _
          "INNER JOIN TB_VENDEDORES V ON F.COD_VENDEDOR = V.COD_VENDEDOR " & _
          "WHERE F.COD_FECHAMENTO_PERIODO = " & filtro_cod_fechamento & " " & _
          "ORDER BY V.NOME_VENDEDOR ASC"

Set rsVend = objConexao.Execute(sqlVend)

' ================== CONSULTA PRINCIPAL (com nome do vendedor e ordenada por nome) =====================
sql = "SELECT " & _ 
      "P.COD_PEDIDO, " & _
      "V.NOME_VENDEDOR, " & _
      "B.PARCELA, B.VALOR, B.DATA_VENCIMENTO, B.DATA_PAGAMENTO, B.VALOR_PAGO, B.valor_comissao_vendedor, " & _
      "FI.NOME_FORMANDO AS NOME_FORMANDO, " & _
      "FI.NOME_CURSO AS NOME_CURSO, " & _
      "E.DESCRICAO AS INSTITUICAO " & _
      "FROM TB_PEDIDOS P " & _
      "INNER JOIN TB_PEDIDOS_BOLETOS B ON P.COD_PEDIDO = B.COD_PEDIDO " & _
      "INNER JOIN TB_VENDEDORES V ON P.COD_VENDEDOR = V.COD_VENDEDOR " & _
      "INNER JOIN TB_PEDIDOS_FORMANDOS PF ON P.COD_PEDIDO = PF.COD_PEDIDO " & _
      "INNER JOIN TB_FICHAS_IDENTIFICACAO FI ON PF.COD_FORMANDO = FI.COD_FORMANDO " & _
      "INNER JOIN TB_ENTIDADES E ON FI.COD_CONTRATO = E.COD_CONTRATO " & _
      "WHERE 1=1 "


' 🔹 Filtra pelo COD_FECHAMENTO_PERIODO da URL
If filtro_cod_fechamento <> "" Then
    sql = sql & " AND P.COD_FECHAMENTO_PERIODO = " & filtro_cod_fechamento
End If

' 🔹 Filtra vendedor se informado
If filtro_vendedor <> "" Then
    sql = sql & " AND P.COD_VENDEDOR = " & filtro_vendedor
End If

' 🔹 Filtra status (Pago / Não pago)
If filtro_status = "PAGO" Then
    sql = sql & " AND B.VALOR_PAGO > 0 "
ElseIf filtro_status = "NAO_PAGO" Then
    sql = sql & " AND (B.VALOR_PAGO = 0 OR B.VALOR_PAGO IS NULL) "
End If

sql = sql & " ORDER BY V.NOME_VENDEDOR, P.COD_PEDIDO, B.PARCELA"

Set rs = objConexao.Execute(sql)
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" media="screen"></link>



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
	border: 1px solid white;
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


<title>Comissão Boletos</title>

<div align="center">
    <h2>Extrato de Comissão dos Vendedores</h2>

    <form method="get" action="adm_controle_no_extrato_comissao_vendedores_boletos.asp" style="margin-bottom:20px;">
        <input type="hidden" name="COD_FECHAMENTO_PERIODO" value="<%=filtro_cod_fechamento%>">
        <table border="0" cellpadding="5" cellspacing="0" style="background-color:#f8f8f8; border-radius:8px;">
            <tr>
                <td><strong>Vendedor:</strong></td>
<td>
    <select name="cod_vendedor" style="width:250px;">
        <option value="">Todos</option>
        <%
        If Not rsVend.EOF Then
            Do Until rsVend.EOF
                Dim codV, nomeV
                codV = rsVend("COD_VENDEDOR")
                nomeV = rsVend("NOME_VENDEDOR")
                If CStr(filtro_vendedor) = CStr(codV) Then
                    Response.Write("<option value='" & codV & "' selected>" & nomeV & "</option>")
                Else
                    Response.Write("<option value='" & codV & "'>" & nomeV & "</option>")
                End If
                rsVend.MoveNext
            Loop
        Else
            Response.Write("<option value=''>Nenhum vendedor neste fechamento</option>")
        End If
        rsVend.Close
        Set rsVend = Nothing
        %>
    </select>
</td>

                <td><strong>Status:</strong></td>
                <td>
                    <select name="status">
                        <option value="">Todos</option>
                        <option value="PAGO" <%If filtro_status="PAGO" Then Response.Write("selected")%>>Pago</option>
                        <option value="NAO_PAGO" <%If filtro_status="NAO_PAGO" Then Response.Write("selected")%>>Não Pago</option>
                    </select>
                </td>
                <td><input type="submit" value="Buscar" style="background-color:#E91E63;color:white;border:none;padding:5px 10px;border-radius:5px;cursor:pointer;"></td>
            </tr>
        </table>
    </form>

    <table border="1" cellspacing="0" cellpadding="4" width="95%" style="border-collapse:collapse; font-size:13px;">
        <tr style="background-color:#E91E63; color:white;">
            <th>Cód. Pedido</th>
            <th>Formando</th>
            <th>Instituição</th>
            <th>Curso</th>
            <th>Vendedor</th>
            <th>Parcela</th>
            <th>Valor Parcela (R$)</th>
            <th>Data Vencimento</th>
            <th>Valor Comissão (R$)</th>
            <th>Valor Pago (R$)</th>
            <th>Data Pagamento</th>
            <th>Editar ✍️</th>
            <th>Status</th>
        </tr>

        <%
        If Not rs.EOF Then
            Do Until rs.EOF
                Dim valorPago
                If IsNull(rs("VALOR_PAGO")) Or rs("VALOR_PAGO") = "" Then
                    valorPago = 0
                Else
                    valorPago = CDbl(rs("VALOR_PAGO"))
                End If

                If valorPago > 0 Then
                    corStatus = "#4CAF50" 'verde
                Else
                    corStatus = "#F44336" 'vermelho
                End If
        %>
                <tr>
                    <td align="center"><%=rs("COD_PEDIDO")%></td>
                    <td align="left"><%=rs("NOME_FORMANDO")%></td>
                    <td align="left"><%=rs("INSTITUICAO")%></td>
                    <td align="left"><%=rs("NOME_CURSO")%></td>
                    <td align="center"><%=rs("NOME_VENDEDOR")%></td>
                    <td align="center"><%=rs("PARCELA")%></td>
                    <td align="right"><%=FormatNumber(rs("VALOR"),2)%></td>
                    <td align="center"><%=rs("DATA_VENCIMENTO")%></td>
                    <td align="right"><%=FormatNumber(rs("valor_comissao_vendedor"),2)%></td>
                    <td align="right"><%=FormatNumber(valorPago,2)%></td>
                    <td align="center"><%=arrumaData(rs("DATA_PAGAMENTO"))%></td>
                    <td align="center">
                        <% 
						' --- Mostra o link apenas para usuários autorizados ---
						If Session("COD_USUARIO") = 1 Or Session("COD_USUARIO") = 42 Or Session("COD_USUARIO") = 53 Or Session("COD_VENDEDOR") = 280 Then
						%>
                        <a href="javascript:void(0);" 
                        onclick="abrirPopup('<%=rs("COD_PEDIDO")%>', '<%=rs("PARCELA")%>', '<%=Server.HTMLEncode(Replace(CStr(valorPago), " ", ""))%>')" 
                        style="color:#E91E63;text-decoration:none;">✍️ Editar valor pago</a>
                        <% 
						End If 
						%>

                </td>
                    <td align="center">
                        <div style="width:15px;height:15px;border-radius:50%;background-color:<%=corStatus%>"></div>
                    </td>
                </tr>
        <%
                rs.MoveNext
            Loop
        Else
        %>
            <tr><td colspan="7" align="center">Nenhum registro encontrado.</td></tr>
        <%
        End If
        rs.Close
        Set rs = Nothing
        FecharConexao()
        %>
    </table>
    <!-- Popup -->
<div id="popup-editar" style="display:none; position:fixed; top:0; left:0; width:100%; height:100%; 
  background-color:rgba(0,0,0,0.5); justify-content:center; align-items:center; z-index:9999;">
  <div style="background:white; padding:20px; border-radius:10px; width:320px; text-align:center;">
    <h3 style="margin-top:0;">Editar Valor Pago</h3>

    <form id="formEditarValor" method="post" action="update_valor_pago.asp" onsubmit="return validarForm();">
      <input type="hidden" name="COD_PEDIDO" id="COD_PEDIDO">
      <input type="hidden" name="PARCELA" id="PARCELA">
      <input type="hidden" name="COD_FECHAMENTO_PERIODO" id="COD_FECHAMENTO_PERIODO" value="<%= Server.HTMLEncode(filtro_cod_fechamento) %>">

      <div style="text-align:left;">
        <label><strong>Pedido:</strong></label> <span id="lblPedido"></span><br>
        <label><strong>Parcela:</strong></label> <span id="lblParcela"></span><br><br>

        <label for="VALOR_PAGO"><strong>Valor Pago (R$)</strong></label><br>
        <input type="text" name="VALOR_PAGO" id="VALOR_PAGO" 
               style="width:100%; padding:6px; box-sizing:border-box; text-align:right;"
               placeholder="ex: 1.234,56"
               oninput="formatarCampoValor(this)">
      </div>

      <div style="margin-top:12px;">
        <button type="submit" style="background:#4CAF50;color:white;border:none;padding:8px 12px;border-radius:5px;cursor:pointer;">Salvar</button>
        <button type="button" onclick="fecharPopup()" style="background:#F44336;color:white;border:none;padding:8px 12px;border-radius:5px;cursor:pointer;margin-left:8px;">Cancelar</button>
      </div>
    </form>
  </div>
</div>

<script>
function abrirPopup(codPedido, parcela, valorPago) {
  // garante strings
  codPedido = String(codPedido || "");
  parcela = String(parcela || "");
  valorPago = String(valorPago || "").trim();

  // corrige formatos vindos do ASP
  valorPago = valorPago.replace(/\./g, "").replace(",", ".");
  if (!isNaN(valorPago) && valorPago !== "") {
    valorPago = parseFloat(valorPago).toFixed(2);
  }

  // converte para BR (vírgula decimal)
  valorPago = valorPago.replace(".", ",");

  // preenche
  document.getElementById("COD_PEDIDO").value = codPedido;
  document.getElementById("PARCELA").value = parcela;
  document.getElementById("VALOR_PAGO").value = valorPago;
  document.getElementById("lblPedido").innerText = codPedido;
  document.getElementById("lblParcela").innerText = parcela;

  document.getElementById("popup-editar").style.display = "flex";
}

function fecharPopup() {
  document.getElementById("popup-editar").style.display = "none";
}

function validarForm() {
  var campo = document.getElementById("VALOR_PAGO");
  var valor = campo.value.trim();

  if (valor === "") return true; // permite vazio

  // normaliza para formato americano
  var normalizado = valor.replace(/\./g, "").replace(",", ".");
  if (isNaN(normalizado)) {
    alert("Digite um valor numérico válido, ex: 1234,56");
    campo.focus();
    return false;
  }

  // define o valor normalizado antes de enviar
  campo.value = parseFloat(normalizado).toFixed(2);
  return true;
}

// formata conforme digita (mantém vírgula como decimal)
function formatarCampoValor(input) {
  let valor = input.value;

  // remove caracteres inválidos
  valor = valor.replace(/[^0-9,]/g, "");

  // evita mais de uma vírgula
  const partes = valor.split(",");
  if (partes.length > 2) valor = partes[0] + "," + partes[1];

  // atualiza campo
  input.value = valor;
}
</script>


</div>

<!-- #include file="files/includes/miolo-inferior.asp" -->