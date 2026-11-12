<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Validação especial para verificar se possui acesso a página mãe de formulário
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS where COD_PAGINA = '228' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se não encontrar então redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
		
	rs.Close
	Set rs=Nothing
	
FecharConexao()

'/////////////////////////////////////////////////////////////////////////////////////////////

'Declara variáveis para utilização do sistema
Dim cod_tabela_comissao_vendedores
cod_tabela_comissao_vendedores = Request.QueryString("cod_tabela_comissao_vendedores")

Dim commissionPattern, commissionName
Set commissionPattern = New RegExp
commissionPattern.Pattern = "^[1-9][0-9]{0,9}$"
If Not commissionPattern.Test(cod_tabela_comissao_vendedores) Then
    Response.Status = "400 Bad Request"
    Response.Write "Informe um c&oacute;digo de tabela v&aacute;lido."
    Response.End
End If
Set commissionPattern = Nothing

AbrirConexao()
Set rs = objConexao.Execute("SELECT descricao FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES WHERE cod_tabela_comissao_vendedores = " & cod_tabela_comissao_vendedores)
If rs.Eof Then
    rs.Close
    Set rs = Nothing
    FecharConexao()
    Response.Status = "404 Not Found"
    Response.Write "Tabela de comiss&otilde;es n&atilde;o encontrada."
    Response.End
End If
commissionName = rs("descricao") & ""
rs.Close
Set rs = Nothing

Function CommissionPrintValue(paymentType, paymentValue)
    Dim encodedValue
    encodedValue = Server.HTMLEncode(paymentValue & "")
    If Len(Trim(paymentValue & "")) = 0 Then
        CommissionPrintValue = "&mdash;"
    ElseIf CStr(paymentType & "") = "F" Then
        CommissionPrintValue = "R$ " & encodedValue
    Else
        CommissionPrintValue = encodedValue & "%"
    End If
End Function
%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Tabela de comiss&otilde;es | Portal Stillos</title>
<style>
* { box-sizing:border-box; }
body { margin:0; background:#f6f7fb; color:#242738; font:14px/1.45 system-ui,Arial,sans-serif; }
.cap-document { max-width:1150px; margin:16px auto; padding:clamp(12px,3vw,28px); border:1px solid #e4dfeb; border-radius:12px; background:#fff; }
.cap-doc-toolbar { text-align:right; margin:0 0 12px; }
.cap-doc-toolbar button { padding:8px 12px; border:0; border-radius:7px; background:#b20b61; color:white; font:inherit; cursor:pointer; }
.commission-header { display:flex; flex-wrap:wrap; align-items:center; justify-content:space-between; gap:12px; padding-bottom:12px; border-bottom:2px solid #b20b61; }
.commission-header h1 { margin:0; font-size:clamp(18px,3vw,26px); }
.commission-logo { width:clamp(100px,20vw,150px); height:auto; border-radius:7px; background:#292b40; }
.commission-scroll { overflow-x:auto; border:1px solid #e4dfeb; border-radius:8px; }
.commission-scroll:focus-visible { outline:2px solid #b20b61; }
.commission-table { width:100%; min-width:570px; border-collapse:collapse; font-variant-numeric:tabular-nums; }
.commission-table th, .commission-table td { padding:7px 9px; border-bottom:1px solid #e4dfeb; text-align:right; white-space:nowrap; }
.commission-table tr > :first-child { text-align:left; white-space:normal; }
.commission-table thead { background:#f0ebf5; }
.commission-table tbody tr:nth-child(even) { background:#faf9fc; }
.commission-table tr > :last-child { font-weight:700; }
.commission-note { margin:12px 0 0; color:#514960; }
@page { size:A4 landscape; margin:10mm; }
@media print {
  body { color:#111; background:#fff; font-size:9pt; line-height:1.25; }
  .cap-document { margin:0; max-width:none; padding:3mm 4mm; border:0; }
  .cap-doc-toolbar { display:none; }
  .commission-header { padding-bottom:2mm; border-color:#444; break-inside:avoid; }
  .commission-header h1 { font-size:14pt; }
  .commission-logo { width:30mm; background:transparent; filter:brightness(0); }
  .commission-scroll { overflow:visible; }
  .commission-table { min-width:0; }
  .commission-table th, .commission-table td { padding:1.2mm 2mm; }
  .commission-table thead { display:table-header-group; }
  .commission-table tr, .commission-note { break-inside:avoid; }
}

.commission-table thead th { text-align:center; }
.commission-table tbody th { font-weight:600; }
.commission-table tbody td { font-weight:400; }
.commission-table tbody td:nth-child(4),.commission-table thead tr:last-child th:nth-child(3) { border-left:1px solid #e4dfeb; }
.commission-table tbody th { width:46%; }
</style>
</head>
<body>
<main class="cap-document" data-commission-report>
  <%If Request.QueryString("modal") <> "1" Then%>
  <div class="cap-doc-toolbar"><button type="button" onclick="window.print()">Imprimir / Salvar PDF</button></div>
  <%End If%>
  <header class="commission-header">
    <div><h1>Tabela de comiss&otilde;es</h1><p><%=Server.HTMLEncode(commissionName)%> &middot; C&oacute;digo <%=cod_tabela_comissao_vendedores%></p></div>
    <img class="commission-logo" src="images/logoparafundopreto.png" alt="Stillos">
  </header>
  <p class="commission-note">Produtos ativos com comiss&atilde;o. O tra&ccedil;o indica um valor em branco.</p>
  <%
  Set rs=objConexao.Execute("SELECT cod_produto, nome_produto FROM stillos3.TB_PRODUTOS WHERE ativo = 1 and PAGAMENTO_COMISSAO = 'S' order by order_menu")
  If rs.Eof Then%>
    <p>Nenhum produto ativo com comiss&atilde;o encontrado.</p>
  <%Else%>
  <div class="commission-scroll" tabindex="0" role="region" aria-label="Valores por produto">
    <table class="commission-table">
      <thead>
        <tr><th scope="col" rowspan="2">Produto</th><th scope="colgroup" colspan="2">Vendas normais</th><th scope="colgroup" colspan="2">Vendas de repasse</th></tr>
        <tr><th scope="col">&Agrave; vista</th><th scope="col">A prazo</th><th scope="col">&Agrave; vista</th><th scope="col">A prazo</th></tr>
      </thead>
      <tbody>
      <%
      While Not rs.Eof
cod_produto = rs("cod_produto") 
					
					'///////////////////////////////////////////////////////////
					'Verifica se possui o produto cadastrado nesta tabela de venda
					Set rs2=objConexao.Execute("SELECT * FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES_PRODUTOS WHERE cod_tabela_comissao_vendedores = "&cod_tabela_comissao_vendedores&" and cod_produto ="&cod_produto)
				
						If not rs2.Eof Then
							tipo_pagamento_normal_avista=trim(rs2("tipo_pagamento_normal_avista"))
							valor_pagamento_normal_avista= ArrumaNumero(rs2("valor_pagamento_normal_avista"),"",2)
							tipo_pagamento_normal_aprazo=trim(rs2("tipo_pagamento_normal_aprazo"))
							valor_pagamento_normal_aprazo= ArrumaNumero(rs2("valor_pagamento_normal_aprazo"),"",2)
							tipo_pagamento_repasse_avista=trim(rs2("tipo_pagamento_repasse_avista"))
							valor_pagamento_repasse_avista= ArrumaNumero(rs2("valor_pagamento_repasse_avista"),"",2)
							tipo_pagamento_repasse_aprazo=trim(rs2("tipo_pagamento_repasse_aprazo"))
							valor_pagamento_repasse_aprazo= ArrumaNumero(rs2("valor_pagamento_repasse_aprazo"),"",2)

							
						Else
							tipo_pagamento_normal_avista=""
							valor_pagamento_normal_avista=""
							tipo_pagamento_normal_aprazo=""
							valor_pagamento_normal_aprazo=""
							tipo_pagamento_repasse_avista=""
							valor_pagamento_repasse_avista=""
							tipo_pagamento_repasse_aprazo=""
							valor_pagamento_repasse_aprazo=""					
						End If
					
					rs2.Close
					Set rs2=Nothing
      %>
        <tr>
          <th scope="row"><%=Server.HTMLEncode(rs("nome_produto") & "")%></th>
          <td><%=CommissionPrintValue(TIPO_PAGAMENTO_NORMAL_AVISTA, VALOR_PAGAMENTO_NORMAL_AVISTA)%></td>
          <td><%=CommissionPrintValue(TIPO_PAGAMENTO_NORMAL_APRAZO, VALOR_PAGAMENTO_NORMAL_APRAZO)%></td>
          <td><%=CommissionPrintValue(TIPO_PAGAMENTO_REPASSE_AVISTA, VALOR_PAGAMENTO_REPASSE_AVISTA)%></td>
          <td><%=CommissionPrintValue(TIPO_PAGAMENTO_REPASSE_APRAZO, VALOR_PAGAMENTO_REPASSE_APRAZO)%></td>
        </tr>
      <%rs.MoveNext
      Wend%>
      </tbody>
    </table>
  </div>
  <%End If
  rs.Close
  Set rs = Nothing
  FecharConexao()%>
</main>
</body>
</html>
