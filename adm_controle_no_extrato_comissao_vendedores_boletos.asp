<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<%
'================= FILTROS =====================
Dim filtro_vendedor, filtro_status, filtro_cod_fechamento, sql, corStatus, rs

filtro_vendedor = Request("cod_vendedor")
filtro_status = Request("status")
filtro_cod_fechamento = Request("COD_FECHAMENTO_PERIODO")

Response.CodePage = 65001
Response.Charset = "utf-8"
Dim capBoletoCodigo, capBoletoUsuario
capBoletoUsuario = Session("COD_USUARIO") & ""
Select Case capBoletoUsuario
    Case "1", "16", "22", "42", "53", "57"
    Case Else
        If Len(capBoletoUsuario) = 0 Or (Session("COD_VENDEDOR") & "") <> "280" Then
            Response.Status = "403 Forbidden"
            Response.Write "Acesso restrito ao relat&oacute;rio de boletos."
            Response.End
        End If
End Select
Set capBoletoCodigo = New RegExp
capBoletoCodigo.Pattern = "^[0-9]{1,10}$"
If Not capBoletoCodigo.Test(filtro_cod_fechamento & "") Then
    Response.Status = "400 Bad Request"
    Response.Write "Informe um per&iacute;odo v&aacute;lido."
    Response.End
End If
If Len(filtro_vendedor & "") > 0 Then
    If Not capBoletoCodigo.Test(filtro_vendedor & "") Then
        Response.Status = "400 Bad Request"
        Response.Write "Informe um vendedor v&aacute;lido."
        Response.End
    End If
End If
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

<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Comiss&atilde;o por boleto | Stillos</title>
<style>
* { box-sizing:border-box; }
body { margin:0; background:#f6f7fb; color:#242738; font:14px/1.45 system-ui,Arial,sans-serif; }
.cap-document { max-width:1400px; margin:14px auto; padding:clamp(12px,3vw,26px); border:1px solid #e4dfeb; border-radius:12px; background:#fff; }
.cap-doc-header { display:flex; align-items:center; justify-content:space-between; gap:12px; padding-bottom:12px; margin-bottom:16px; border-bottom:2px solid #b20b61; }
.cap-doc-header h1 { margin:0; font:700 clamp(18px,3vw,27px)/1.2 system-ui,sans-serif; }
.cap-doc-logo { flex:none; width:clamp(95px,20vw,150px); height:auto; background:#292b40; border-radius:7px; }
.cap-doc-toolbar { display:flex; justify-content:flex-end; margin-bottom:10px; }
.cap-doc-button { display:inline-block; padding:8px 12px; border:1px solid #d8d0e1; border-radius:7px; background:#fff; color:#88134f; font:600 13px/1.3 system-ui,sans-serif; cursor:pointer; }
.cap-doc-primary { background:#b20b61; color:#fff; border-color:#b20b61; }
.cap-doc-scroll { overflow-x:auto; }
.cap-doc-scroll table { border-collapse:collapse; border:0; }
.cap-doc-scroll th, .cap-doc-scroll td { padding:6px 5px; border-color:#e4dfeb; font-size:inherit; }
.cap-doc-scroll th { background:#f3f0f7; color:#242738; }
.cap-doc-scroll .folha { width:100%; margin:12px 0; border:1px solid #e4dfeb; border-radius:8px; }
.cap-doc-scroll .folha td { padding:14px; }
.cap-doc-scroll .folha p { margin:8px 0; }
.cap-boleto-filter { display:flex; flex-wrap:wrap; gap:10px; align-items:flex-end; margin:0 0 14px; padding:12px; background:#f7f5fa; border-radius:8px; }
.cap-boleto-filter label { display:block; font-weight:600; }
.cap-boleto-filter select, .cap-boleto-edit input[type=text] { display:block; max-width:100%; margin-top:4px; padding:7px; border:1px solid #d8d0e1; border-radius:6px; background:#fff; color:inherit; }
.cap-boleto-table { width:100%; font-size:12px; }
.cap-boleto-table tbody tr:nth-child(even) { background:#faf9fc; }
.cap-boleto-status { display:inline-block; padding:3px 6px; border-radius:5px; white-space:nowrap; background:#f9e7ed; color:#86234a; }
.cap-boleto-status.is-paid { background:#e5f2ea; color:#275e42; }
.cap-boleto-edit { width:min(430px,calc(100% - 24px)); padding:20px; border:1px solid #e4dfeb; border-radius:12px; color:#242738; }
.cap-boleto-edit::backdrop { background:#24273899; }
.cap-boleto-edit h2 { margin-top:0; }
@media screen and (max-width:720px) {
  .cap-boleto-table, .cap-boleto-table tbody { display:block; }
  .cap-boleto-table thead { position:absolute; width:1px; height:1px; overflow:hidden; clip-path:inset(50%); }
  .cap-boleto-table tbody tr { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); margin-bottom:10px; border:1px solid #e4dfeb; border-radius:8px; }
  .cap-boleto-table td { display:block; text-align:left; overflow-wrap:anywhere; }
  .cap-boleto-table td::before { content:attr(data-label); display:block; font-weight:650; }
  .cap-boleto-table td[colspan] { grid-column:1/-1; }
}
@page { size:A4 landscape; margin:10mm; }
@media print {
  body { background:white; color:#111; font-size:9.5pt; }
  .cap-document { max-width:none; margin:0; padding:5mm 8mm; border:0; }
  .cap-doc-logo { width:32mm; background:transparent; filter:brightness(0); }
  .cap-doc-header h1 { font-size:15pt; }
  .cap-doc-toolbar, .cap-boleto-filter, .cap-boleto-edit, .cap-boleto-action { display:none !important; }
  .cap-doc-scroll { overflow:visible; }
  .cap-doc-scroll table { max-width:100%; }
  .cap-doc-scroll th, .cap-doc-scroll td { padding:4px 3px; }
  .cap-doc-scroll tr { break-inside:avoid; }
  .cap-doc-scroll thead { display:table-header-group; }
  .cap-doc-scroll .folha { break-after:page; }
  .cap-doc-scroll .folha:last-of-type { break-after:auto; }
}
</style>
</head>
<body>
<main class="cap-document" data-cap-report>
<%If Request.QueryString("modal") <> "1" Then%><div class="cap-doc-toolbar"><button class="cap-doc-button cap-doc-primary" type="button" onclick="window.print()">Imprimir / Salvar PDF</button></div><%End If%>
<header class="cap-doc-header"><h1>Comiss&atilde;o por boleto</h1><img class="cap-doc-logo" src="images/logoparafundopreto.png" alt="Stillos"></header>
<div class="cap-doc-scroll">

<form class="cap-boleto-filter" method="get" action="adm_controle_no_extrato_comissao_vendedores_boletos.asp">
  <input type="hidden" name="COD_FECHAMENTO_PERIODO" value="<%=Server.HTMLEncode(filtro_cod_fechamento & "")%>">
  <%If Request.QueryString("modal") = "1" Then%><input type="hidden" name="modal" value="1"><%End If%>
  <label>Vendedor<select name="cod_vendedor">
        <option value="">Todos</option>
        <%
        If Not rsVend.EOF Then
            Do Until rsVend.EOF
                Dim codV, nomeV
                codV = rsVend("COD_VENDEDOR")
                nomeV = rsVend("NOME_VENDEDOR")
                If CStr(filtro_vendedor) = CStr(codV) Then
                    Response.Write("<option value='" & codV & "' selected>" & Server.HTMLEncode(nomeV & "") & "</option>")
                Else
                    Response.Write("<option value='" & codV & "'>" & Server.HTMLEncode(nomeV & "") & "</option>")
                End If
                rsVend.MoveNext
            Loop
        Else
            Response.Write("<option value=''>Nenhum vendedor neste fechamento</option>")
        End If
        rsVend.Close
        Set rsVend = Nothing
        %>
    </select></label>
  <label>Status<select name="status">
                        <option value="">Todos</option>
                        <option value="PAGO" <%If filtro_status="PAGO" Then Response.Write("selected")%>>Pago</option>
                        <option value="NAO_PAGO" <%If filtro_status="NAO_PAGO" Then Response.Write("selected")%>>N&atilde;o pago</option>
                    </select></label>
  <button class="cap-doc-button cap-doc-primary" type="submit">Aplicar filtros</button>
</form>
    <table class="cap-boleto-table"><thead><tr>
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
            <th class="cap-boleto-action">A&ccedil;&atilde;o</th>
            <th>Status</th>
        </tr></thead><tbody>

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
                    <td data-label="Pedido"><%=rs("COD_PEDIDO")%></td>
                    <td data-label="Formando"><%=Server.HTMLEncode(rs("NOME_FORMANDO") & "")%></td>
                    <td data-label="Institui&ccedil;&atilde;o"><%=Server.HTMLEncode(rs("INSTITUICAO") & "")%></td>
                    <td data-label="Curso"><%=Server.HTMLEncode(rs("NOME_CURSO") & "")%></td>
                    <td data-label="Vendedor"><%=Server.HTMLEncode(rs("NOME_VENDEDOR") & "")%></td>
                    <td data-label="Parcela"><%=rs("PARCELA")%></td>
                    <td data-label="Valor da parcela"><%=FormatNumber(rs("VALOR"),2)%></td>
                    <td data-label="Vencimento"><%=rs("DATA_VENCIMENTO")%></td>
                    <td data-label="Comiss&atilde;o"><%=FormatNumber(rs("valor_comissao_vendedor"),2)%></td>
                    <td data-label="Valor pago"><%=FormatNumber(valorPago,2)%></td>
                    <td data-label="Data do pagamento"><%=arrumaData(rs("DATA_PAGAMENTO"))%></td>
                    <td data-label="A&ccedil;&atilde;o" class="cap-boleto-action">
                        <% 
						' --- Mostra o link apenas para usuários autorizados ---
						If Session("COD_USUARIO") = 1 Or Session("COD_USUARIO") = 42 Or Session("COD_USUARIO") = 53 Or Session("COD_VENDEDOR") = 280 Then
						%>
                        <a href="javascript:void(0);" 
                        onclick="abrirPopup('<%=rs("COD_PEDIDO")%>', '<%=rs("PARCELA")%>', '<%=Server.HTMLEncode(Replace(CStr(valorPago), " ", ""))%>')" 
                        style="color:#E91E63;text-decoration:none;">Editar valor pago</a>
                        <% 
						End If 
						%>

                </td>
                    <td data-label="Status">
                        <%If valorPago > 0 Then%><span class="cap-boleto-status is-paid">Pago</span><%Else%><span class="cap-boleto-status">N&atilde;o pago</span><%End If%>
                    </td>
                </tr>
        <%
                rs.MoveNext
            Loop
        Else
        %>
            <tr><td colspan="13" align="center">Nenhum registro encontrado.</td></tr>
        <%
        End If
        rs.Close
        Set rs = Nothing
        FecharConexao()
        %>
    </tbody></table>

</div></main>
<dialog id="popup-editar" class="cap-boleto-edit" aria-labelledby="cap-boleto-edit-title">
    <h2 id="cap-boleto-edit-title">Editar valor pago</h2>

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
        <button class="cap-doc-button" type="submit">Salvar</button>
        <button class="cap-doc-button" type="button" onclick="fecharPopup()">Cancelar</button>
      </div>
    </form>
</dialog>

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

  document.getElementById("popup-editar").showModal();
}

function fecharPopup() {
  document.getElementById("popup-editar").close();
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
</body></html>
