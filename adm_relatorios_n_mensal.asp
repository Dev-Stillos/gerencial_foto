<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Valida��o especial para verificar se possui sub-valida��o espec�fica
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '29' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se n�o encontrar ent�o redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
		
	rs.Close
	Set rs=Nothing
	
FecharConexao()








Dim notasMes, notasAno, notasValida, notasErro, DATA_INICIO, DATA_FIM, notasPeriodo
notasMes = Trim(Request.QueryString("MES_INICIO"))
notasAno = Trim(Request.QueryString("ANO_INICIO"))
Set notasValida = New RegExp
notasValida.Pattern = "^[0-9]{1,2}$"
If Not notasValida.Test(notasMes) Then notasErro = "Informe um m&ecirc;s v&aacute;lido."
notasValida.Pattern = "^[0-9]{4}$"
If Not notasValida.Test(notasAno) Then notasErro = "Informe um ano v&aacute;lido."
If Len(notasErro) = 0 Then
    If CInt(notasMes) < 1 Or CInt(notasMes) > 12 Or CInt(notasAno) < 1000 Or CInt(notasAno) > 9998 Then
        notasErro = "Informe um m&ecirc;s e ano v&aacute;lidos."
    Else
        DATA_INICIO = DateSerial(CInt(notasAno), CInt(notasMes), 1)
        DATA_FIM = DateAdd("m", 1, DATA_INICIO)
        notasPeriodo = MonthName(Month(DATA_INICIO)) & " de " & Year(DATA_INICIO)
    End If
End If
%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Relat&oacute;rio de notas | Portal Stillos</title>
<link rel="stylesheet" href="css/vendor/bootstrap.min.css">
<link rel="stylesheet" href="css/dashboard.css">
<style>
.notes-report { margin:24px auto; max-width:1550px; padding:clamp(16px,3vw,32px); background:white; border:1px solid #e4dfeb; border-radius:14px; }
.notes-tools,.notes-header { display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:12px; }
.notes-header { margin:18px 0; padding-bottom:12px; border-bottom:2px solid #b20b61; }
.notes-header p { margin:0; }
.notes-logo { width:150px; padding:10px; border-radius:8px; background:#292b40; }
.notes-orders { display:grid; gap:16px; }
.notes-card { border:1px solid #e4dfeb; border-radius:12px; background:#fff; }
.notes-card-head { display:flex; flex-wrap:wrap; justify-content:space-between; align-items:center; gap:8px 20px; padding:12px 18px; background:#f6f2f8; border-bottom:1px solid #e4dfeb; border-radius:11px 11px 0 0; }
.notes-card-head h2 { margin:0 0 3px; font-size:1.15em; }
.notes-card-head p { margin:0; color:#53566b; }
.notes-card-total { text-align:right; }
.notes-card-total strong { display:block; font-size:1.15em; color:#242738; }
.notes-card-body { display:grid; grid-template-columns:minmax(0,1fr) minmax(0,1.45fr) minmax(0,2fr); }
.notes-card-body > section { min-width:0; padding:14px 18px; }
.notes-card-body > section + section { border-left:1px solid #eeeaf2; }
.notes-card-body h3 { margin:0 0 8px; font-size:1em; }
.notes-pay,.notes-person dl { margin:0; }
.notes-pay div { display:flex; flex-wrap:wrap; justify-content:space-between; gap:1px 12px; padding:2px 0; }
.notes-pay dt,.notes-person dt,.notes-label { font-size:.9em; color:#53566b; font-weight:500; }
.notes-pay dd { margin:0; white-space:nowrap; font-variant-numeric:tabular-nums; }
.notes-items { width:100%; border-collapse:collapse; overflow-wrap:anywhere; }
.notes-items td { padding:3px 0; border-bottom:1px solid #eeeaf2; }
.notes-items td:last-child { width:1%; padding-left:10px; text-align:right; font-variant-numeric:tabular-nums; }
.notes-person { overflow-wrap:anywhere; }
.notes-person + .notes-person { margin-top:8px; padding-top:8px; border-top:1px solid #eeeaf2; }
.notes-person h4 { margin:0 0 5px; font-size:1em; font-weight:700; }
.notes-person dl { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:4px 10px; }
.notes-person .notes-address { grid-column:1 / -1; }
.notes-person dd { margin:0; }
.notes-person dd:empty::before { content:"\2014"; }
.notes-month-total { margin-top:16px; padding:12px 18px; background:#f6f2f8; border:1px solid #e4dfeb; border-radius:12px; }
.notes-month-total h2 { margin:0 0 8px; }
.notes-month-total dl { display:flex; flex-wrap:wrap; gap:4px 24px; }
.notes-month-total dl div { gap:6px; }
@media screen {
  .report-embedded .notes-report { margin:12px auto; }
}
@media screen and (max-width:780px) {
  .notes-card-body { grid-template-columns:minmax(0,1fr); }
  .notes-card-body > section + section { border-left:0; border-top:1px solid #eeeaf2; }
}
@page { size:A4 landscape; margin:11mm; }
@media print {
  body { margin:0; background:#fff; color:#111; font-size:9pt; line-height:1.2; }
  .notes-report { margin:0; padding:0; max-width:none; border:0; }
  .notes-tools { display:none; }
  .notes-header { margin:0 0 3mm; padding:0 0 2mm; border-color:#555; break-after:avoid; }
  .notes-header h1 { font-size:16pt; margin:0 0 1mm; }
  .notes-logo { width:30mm; padding:0; background:transparent; filter:brightness(0); }
  .notes-orders { display:block; }
  .notes-card,.notes-month-total { border:1px solid #aaa; border-radius:0; break-inside:avoid; }
  .notes-card + .notes-card { margin-top:2.5mm; }
  .notes-card-head { padding:1.2mm 2mm; background:#fff; border-color:#aaa; border-radius:0; break-after:avoid; }
  .notes-card-head h2,.notes-card-head p { display:inline; margin-right:2mm; }
  .notes-card-total strong { display:inline; margin-left:1mm; }
  .notes-card-body { grid-template-columns:minmax(0,19fr) minmax(0,32fr) minmax(0,49fr); }
  .notes-card-body > section { padding:1.4mm 2mm; }
  .notes-card-body h3,.notes-person h4 { margin:0 0 .6mm; break-after:avoid; }
  .notes-items td { padding:.35mm 0; }
  .notes-items tr,.notes-person { break-inside:avoid; }
  .notes-pay div { gap:0 2mm; padding:0; }
  .notes-person + .notes-person { margin-top:1mm; padding-top:1mm; }
  .notes-person dl { gap:.5mm 2mm; }
  .notes-person dt,.notes-person dd { display:inline; }
  .notes-person dt::after { content:": "; }
  .notes-month-total { margin-top:3mm; padding:1.4mm 2mm; background:#fff; }
  .notes-month-total h2 { font-size:10pt; margin-bottom:1mm; }
}
</style>
</head>
<body<%If Request.QueryString("modal") = "1" Then%> class="report-embedded"<%End If%>>
<main class="notes-report" data-report-preview-content>
<%If Request.QueryString("modal") <> "1" Then%>
<nav class="notes-tools" aria-label="A&ccedil;&otilde;es do relat&oacute;rio">
  <a class="btn btn-outline-secondary" href="adm_relatorios.asp">Voltar aos relat&oacute;rios</a>
  <%If Len(notasErro) = 0 Then%><button type="button" class="btn btn-primary" onclick="window.print()">Imprimir / Salvar PDF</button><%End If%>
</nav>
<%End If%>
<header class="notes-header">
  <div><h1>Relat&oacute;rio de notas</h1><p><%=Server.HTMLEncode(notasPeriodo & "")%></p></div>
  <img class="notes-logo" src="images/logoparafundopreto.png" alt="Stillos">
</header>
<%If Len(notasErro) > 0 Then%>
<p class="alert alert-warning" role="alert" data-report-error><%=notasErro%></p>
<%Else%>
<%
		AbrirConexao()
		AbrirConexaoSQL6()
		
		Dim rs, sql
			  
		sql = "SELECT COD_PEDIDO, XTOTAL_CHEQUE, XTOTAL_CARTAO, PAGAMENTO_EM_DINHEIRO, DATA_VENDA FROM VW_FECHAMENTO_GERAL WHERE DATA_VENDA >= "&sqlData(DATA_INICIO)&" AND DATA_VENDA < "&sqlData(DATA_FIM)&" ORDER BY DATA_VENDA, COD_PEDIDO"

		
		Set rs = objConexao.Execute(sql)
		
		If not rs.Eof Then

		%>
<div class="notes-orders" aria-label="Pedidos do per&iacute;odo">
<%
		Dim COD_PEDIDO
		Dim TOTAL_DINHEIRO, TOTAL_CHEQUE, TOTAL_CARTAO, TOTAL_FORMA_PAGAMENTO
		Dim TOTAL_GERAL_PAG
		Dim TOTAL_GERAL_DINHEIRO, TOTAL_GERAL_CHEQUE, TOTAL_GERAL_CARTAO
		
		TOTAL_GERAL_DINHEIRO	= 0	
		TOTAL_GERAL_CHEQUE		= 0
		TOTAL_GERAL_CARTAO		= 0
		
		While not rs.Eof
		
		
		COD_PEDIDO = rs("COD_PEDIDO")
		'Forma de Pagamento
				
					TOTAL_DINHEIRO	= CCur(0&rs("PAGAMENTO_EM_DINHEIRO"))
					TOTAL_CHEQUE	= CCur(0&rs("XTOTAL_CHEQUE"))
					TOTAL_CARTAO	= CCur(0&rs("XTOTAL_CARTAO"))
					TOTAL_FORMA_PAGAMENTO = (TOTAL_DINHEIRO + TOTAL_CHEQUE + TOTAL_CARTAO)
					TOTAL_GERAL_DINHEIRO = (TOTAL_GERAL_DINHEIRO + TOTAL_DINHEIRO)
					TOTAL_GERAL_CHEQUE = (TOTAL_GERAL_CHEQUE + TOTAL_CHEQUE)
					TOTAL_GERAL_CARTAO = (TOTAL_GERAL_CARTAO + TOTAL_CARTAO)
					TOTAL_GERAL_PAG = (TOTAL_GERAL_PAG + TOTAL_FORMA_PAGAMENTO)
		%>
        
					<article class="notes-card">
  <header class="notes-card-head">
    <div>
      <h2>Pedido <%=COD_PEDIDO%></h2>
      <p><%=Server.HTMLEncode(rs("DATA_VENDA") & "")%></p>
    </div>
    <div class="notes-card-total"><span class="notes-label">Total do pedido</span>
      <strong><%=FormatCurrency(TOTAL_FORMA_PAGAMENTO, 2)%></strong>
    </div>
  </header>
  <div class="notes-card-body">
    <section>
      <h3>Pagamento</h3>
      <dl class="notes-pay">
        <div><dt>Dinheiro</dt><dd><%=FormatCurrency(TOTAL_DINHEIRO, 2)%></dd></div>
        <div><dt>Cheque</dt><dd><%=FormatCurrency(TOTAL_CHEQUE, 2)%></dd></div>
        <div><dt>Cart&atilde;o</dt><dd><%=FormatCurrency(TOTAL_CARTAO, 2)%></dd></div>
      </dl>
    </section>
    <section class="notes-products">
      <h3>Produtos</h3>
      <table class="notes-items" aria-label="Produtos e quantidades">
        <%
				   '1) �lbum
				   '2) Prod. Agregado
				   '3) Produtos Extras
				   %>
                    
                    <%
					Set rs3=objConexao.Execute("SELECT COD_FORMANDO FROM TB_PEDIDOS_FORMANDOS WHERE COD_PEDIDO="&COD_PEDIDO)
					
					If not rs3.Eof Then
					
					While not rs3.Eof
					
						Set rs4=objConexao.Execute("SELECT FATURADO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO = "&rs3("COD_FORMANDO"))
						
						If not rs4.Eof Then
						
						Dim FATURADO
						FATURADO = formatNumber(0&rs4("FATURADO"),0)
						
								If FATURADO > 0 Then
						%>
        	            <tr>
            	          <td>Fotos</td>
                	      <td><%=FATURADO%></td>
                    	</tr>
	                    <%
						End If
						
								Set rsAgregados = objConexao.Execute(	" SELECT TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO, TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO_AGREGADO, TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.TIPO, "&_
												" Sum(TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.QUANTIDADE) as QUANTIDADE, "&_
												" TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.VENDIDO, TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.VALOR_UN, "&_
												" TB_PRODUTOS.NOME_PRODUTO FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS "&_
												" ON TB_PRODUTOS.COD_PRODUTO = TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO "&_
												" WHERE TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO="&rs3("COD_FORMANDO")&" AND TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.VENDIDO = 1"&_
												" GROUP BY TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO ")
								
								
								If not rsAgregados.Eof Then
								
								While not rsAgregados.Eof
								%>
								<tr>
								  <td><%=Server.HTMLEncode(rsAgregados("NOME_PRODUTO") & "")%></td>
								  <td><%=rsAgregados("QUANTIDADE")%></td>
								</tr>
								
								<%
								rsAgregados.MoveNext()
								Wend
								
								rsAgregados.Close
								Set rsAgregados = Nothing
								
								End If
						
								
								Set rsProdutosAparte = objConexao.Execute(" SELECT TB_PEDIDOS_PRODUTOS.COD_PEDIDO_PRODUTO, TB_PRODUTOS.NOME_PRODUTO, TB_PEDIDOS_PRODUTOS.VALOR_UN, "&_
										  " TB_PEDIDOS_PRODUTOS.QUANTIDADE FROM TB_PRODUTOS INNER JOIN "&_
										  " TB_PEDIDOS_PRODUTOS ON TB_PRODUTOS.COD_PRODUTO = TB_PEDIDOS_PRODUTOS.COD_PRODUTO"&_
										  " WHERE COD_PEDIDO="&COD_PEDIDO&" AND COD_FORMANDO="&rs3("COD_FORMANDO")&" ORDER BY NOME_PRODUTO")

								If not rsProdutosAparte.Eof Then
											
											While not rsProdutosAparte.Eof
												
												%>
                                                <tr>
												  <td><%=Server.HTMLEncode(rsProdutosAparte("NOME_PRODUTO") & "")%></td>
												  <td><%=rsProdutosAparte("QUANTIDADE")%></td>
												</tr>
                                                <%
												
											rsProdutosAparte.MoveNext()
											Wend
											
								
								rsProdutosAparte.Close
								Set rsProdutosAparte = Nothing

								End If
								
						End If
						
						
						rs4.Close
						Set rs4=Nothing
					
					
					rs3.MoveNext()
					Wend
					
					End If
					
					rs3.Close
					Set rs3=Nothing
					%>
                  </table>
    </section>
    <section class="notes-people">
      <h3>Formandos</h3>
<%
					Set rs2=objConexao.Execute("SELECT COD_FORMANDO FROM TB_PEDIDOS_FORMANDOS WHERE COD_PEDIDO="&COD_PEDIDO)
					
					Dim NOME_FORMANDO, ENDERECO, CEP, FONE_RES, BAIRRO, CIDADE, ESTADO
					
					While not rs2.Eof
                    NOME_FORMANDO = ""
                    ENDERECO = ""
                    CEP = ""
                    FONE_RES = ""
                    BAIRRO = ""
                    CIDADE = ""
                    ESTADO = ""
                    Set rs3=objConexao.Execute("SELECT NOME_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO ="&rs2("COD_FORMANDO"))
						
						If not rs3.Eof Then
							NOME_FORMANDO = rs3("NOME_FORMANDO")
						End If
							
					rs3.Close
					Set rs3=Nothing
					
					Set rs3=objConexao.Execute("SELECT * FROM  TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO WHERE COD_FORMANDO = "&rs2("COD_FORMANDO")&" ORDER BY DATA_CADASTRO DESC LIMIT 1")
						
						If not rs3.Eof Then
							ENDERECO		= rs3("ENDERECO")
							CEP				= rs3("CEP")
							FONE_RES		= rs3("FONE_RES")
							BAIRRO			= rs3("BAIRRO")
							CIDADE			= rs3("CIDADE")
							ESTADO			= rs3("ESTADO")
						End If
						
					rs3.Close
					Set rs3=Nothing
					%>
<section class="notes-person">
  <h4><%=Server.HTMLEncode(NOME_FORMANDO & "")%></h4>
  <dl>
    <div><dt>CPF</dt><dd><%
						Dim codFormando
						codFormando = Mid(rs2("COD_FORMANDO"), 1, Len(rs2("COD_FORMANDO")) - 1)

						If IsNumeric(codFormando) Then
						Set rs6 = objConexaoSQL6.Execute("SELECT FD_CPF_004 FROM TB_FORMANDOS WHERE FD_CODIG_004 = " & codFormando)

						If Not rs6.Eof Then
							Response.Write(Server.HTMLEncode(rs6("FD_CPF_004") & ""))
						End If

						rs6.Close
						Set rs6 = Nothing
						End If
						%></dd></div>
    <div><dt>Telefone</dt><dd><%=Server.HTMLEncode(FONE_RES & "")%></dd></div>
    <div class="notes-address"><dt>Endere&ccedil;o</dt><dd><%=Server.HTMLEncode(ENDERECO & "")%></dd></div>
    <div><dt>CEP</dt><dd><%=Server.HTMLEncode(CEP & "")%></dd></div>
    <div><dt>Bairro</dt><dd><%=Server.HTMLEncode(BAIRRO & "")%></dd></div>
    <div><dt>Cidade</dt><dd><%=Server.HTMLEncode(CIDADE & "")%></dd></div>
    <div><dt>UF</dt><dd><%=Server.HTMLEncode(ESTADO & "")%></dd></div>
  </dl>
</section>
<%
					rs2.MoveNext()
					Wend
					
					rs2.Close
					Set rs2=Nothing
					%>
                  </section>
  </div>
</article>
<%
			
			rs.MoveNext()
			Wend
			
			rs.Close
			Set rs = Nothing
			

			%>
</div>
<section class="notes-month-total">
  <h2>Totais do m&ecirc;s</h2>
  <dl class="notes-pay">
    <div><dt>Total</dt><dd><strong><%=FormatCurrency(TOTAL_GERAL_PAG, 2)%></strong></dd></div>
    <div><dt>Dinheiro</dt><dd><%=FormatCurrency(TOTAL_GERAL_DINHEIRO, 2)%></dd></div>
    <div><dt>Cheque</dt><dd><%=FormatCurrency(TOTAL_GERAL_CHEQUE, 2)%></dd></div>
    <div><dt>Cart&atilde;o</dt><dd><%=FormatCurrency(TOTAL_GERAL_CARTAO, 2)%></dd></div>
  </dl>
</section>
<%Else
    rs.Close
    Set rs = Nothing
%>
<p class="alert alert-info">Nenhum pedido encontrado para o m&ecirc;s selecionado.</p>
<%End If
FecharConexao()
FecharConexaoSQL6()
End If%>
</main>
</body>
</html>
