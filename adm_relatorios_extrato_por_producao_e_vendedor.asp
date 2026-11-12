<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Session.Timeout=5
'Valida??o especial para verificar se possui sub-valida??o espec?fica
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '31' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se Nao encontrar ent?o redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
		
	rs.Close
	Set rs=Nothing
	
FecharConexao()

Dim COD_PRODUCAO
COD_PRODUCAO = Trim(Request.QueryString("COD_PRODUCAO"))
Dim producaoCodigo
Set producaoCodigo = New RegExp
producaoCodigo.Pattern = "^[0-9]{1,10}$"
If Not producaoCodigo.Test(COD_PRODUCAO) Then
    Response.Write "<!doctype html><html lang='pt-BR'><meta charset='iso-8859-1'><main data-report-preview-content><p data-report-error>Selecione uma produ&ccedil;&atilde;o v&aacute;lida.</p><a href='adm_relatorios.asp'>Voltar aos relat&oacute;rios</a></main></html>"
    Response.End
End If

AbrirConexao()

'1) Deleta TB TEMPOR?RIA antes de inicias
objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA WHERE COD_USUARIO ="&Session("COD_USUARIO"))


'2) Seleciona todos os formandos de uma determinada produ??o
Set rs = objConexao.Execute("SELECT TB_FICHAS_IDENTIFICACAO.COD_FORMANDO,TB_FICHAS_IDENTIFICACAO.COD_CONTRATO,TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO,TB_PEDIDOS_FORMANDOS.COD_PEDIDO,TB_PEDIDOS.DATA_VENDA FROM TB_FICHAS_IDENTIFICACAO LEFT JOIN TB_PEDIDOS_FORMANDOS ON TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = TB_PEDIDOS_FORMANDOS.COD_FORMANDO LEFT JOIN TB_PEDIDOS ON TB_PEDIDOS_FORMANDOS.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO  WHERE TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO = '"&COD_PRODUCAO&"' GROUP BY COD_FORMANDO")
		
		Dim ROTINA_COD_FORMANDO, ROTINA_COD_CONTRATO, ROTINA_TIPO, ROTINA_BRINDE, ROTINA_COD_ULTIMO_VENDEDOR_FICHA, ROTINA_VENDIDO, ROTINA_COD_PEDIDO
		
		While not rs.Eof
		
		ROTINA_COD_FORMANDO = rs("COD_FORMANDO")
		ROTINA_COD_CONTRATO = rs("COD_CONTRATO")

		'====================================================
' GARANTE QUE MOVIMENTACOES EXISTAM (TIPO 0, 1 e 2)
'====================================================
Dim existe

'--- TIPO 0 (ALBUM FISICO) ---
sql = "SELECT COD_MOVIMENTACAO FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES " &_
      "WHERE COD_FORMANDO=" & ROTINA_COD_FORMANDO & " AND TIPO=0 LIMIT 1"
Set rsCheck = objConexao.Execute(sql)

If rsCheck.Eof Then
    objConexao.Execute "INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES " & _
                       "(COD_FORMANDO, TIPO, CADASTRADO_POR, DATA_CADASTRO) VALUES (" & _
                       ROTINA_COD_FORMANDO & ", 0, " & Session("COD_USUARIO") & ", NOW())"
End If

rsCheck.Close

'--- TIPO 1 (FICHA) ---
sql = "SELECT COD_MOVIMENTACAO FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES " &_
      "WHERE COD_FORMANDO=" & ROTINA_COD_FORMANDO & " AND TIPO=1 LIMIT 1"
Set rsCheck = objConexao.Execute(sql)

If rsCheck.Eof Then
    objConexao.Execute "INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES " & _
                       "(COD_FORMANDO, TIPO, CADASTRADO_POR, DATA_CADASTRO) VALUES (" & _
                       ROTINA_COD_FORMANDO & ", 1, " & Session("COD_USUARIO") & ", NOW())"
End If

rsCheck.Close

'--- TIPO 2 (ALBUM DIGITAL) ---
sql = "SELECT COD_MOVIMENTACAO FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES " &_
      "WHERE COD_FORMANDO=" & ROTINA_COD_FORMANDO & " AND TIPO=2 LIMIT 1"
Set rsCheck = objConexao.Execute(sql)

If rsCheck.Eof Then
    objConexao.Execute "INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES " & _
                       "(COD_FORMANDO, TIPO, CADASTRADO_POR, DATA_CADASTRO) VALUES (" & _
                       ROTINA_COD_FORMANDO & ", 2, " & Session("COD_USUARIO") & ", NOW())"
End If

rsCheck.Close
Set rsCheck = Nothing
'====================================================

						
			'3) Verifica quem foi o ?ltimo respons?vel pelo formando
			sql = " SELECT COD_VENDEDOR FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_FORMANDO ="&ROTINA_COD_FORMANDO&_
				  " AND TIPO = 1 ORDER BY DATA_SAIDA DESC LIMIT 1"
			Set rs2 = objConexao.Execute(sql)
						
						If not rs2.Eof Then
						
							ROTINA_COD_ULTIMO_VENDEDOR_FICHA = rs2("COD_VENDEDOR")
						
						Else
							'Se Nao encontrar acusa falha
							Response.Redirect("adm_relatorios.asp?msg=FALHA AO CARREGAR O EXTRATO, FORMANDO: "&ROTINA_COD_FORMANDO)
											
						End If
						
			rs2.Close
			Set rs2 = Nothing

			'4) Verifica quais eram os itens do formando			
				'Nao ? necess?rio a ficha, somente produtos
			sql = " SELECT DISTINCT (TIPO) FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_FORMANDO = "&rs("COD_FORMANDO")&_
				  " AND TIPO <> 1 ORDER BY TIPO"

			Set rs2 = objConexao.Execute(sql)
					
					While not rs2.Eof
					'Agora sabendo quais os produtos e o c?d. do ?ltimo vendedor, verifica se foi vendido
					
					ROTINA_TIPO = formatNumber(0&rs2("TIPO"),0)
							
							If ROTINA_TIPO = 0 Then
							'5) Come?a pelo ?lbum se est? vendido, ?lbum Nao ser? considerado como brinde, pois pode ser vendido parcialmente
									'Pega da pr?pria TB_FICHAS_IDENTIFICAO do formando j? aberta
									ROTINA_VENDIDO = 2
									ROTINA_COD_PEDIDO = 0
									Set rs3=objConexao.Execute("SELECT IFNULL(TB_PEDIDOS_FORMANDOS.COD_PEDIDO,0) AS VENDIDO FROM TB_FICHAS_IDENTIFICACAO LEFT JOIN TB_PEDIDOS_FORMANDOS ON TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = TB_PEDIDOS_FORMANDOS.COD_FORMANDO LEFT JOIN TB_PEDIDOS ON TB_PEDIDOS_FORMANDOS.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO WHERE TB_PEDIDOS.FINALIZA_VENDEDOR = 1 AND  TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = "&ROTINA_COD_FORMANDO)
										
										If not rs3.Eof Then
											
											If int(0&rs3("VENDIDO")) > 0 Then
												ROTINA_VENDIDO = 1
												ROTINA_COD_PEDIDO = rs3("VENDIDO")
											End IF
											
										End If
										
									rs3.Close
									Set rs3=Nothing
									
									ROTINA_BRINDE  = 2
																		
							Else
							'AGREGADOS SE HOUVER
							'//////////
							
							'6) Se tiver mais verifica junto a tabela 
								
							
							'7)Verifica se Nao era de brinde o produto agregado
								
								'Nao SER? VERIFICADO OS PRODUTOS AGREGADOS QUE ERAM DE COMISS?O, SEM ANTES
								'FINALIZAR A ROTINA COLOCANDO OS TIPOS NAS TBS QUE FALTAM
								
								ROTINA_BRINDE  = 2						
							'//////////
							End If
							
							'Ap?s isso tudo faz o INSERT NA TB TEMPOR?RIA
							sql = " INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA (COD_FORMANDO, COD_PRODUCAO, "&_
								  " COD_CONTRATO, TIPO, BRINDE, COD_ULTIMO_VENDEDOR_FICHA, VENDIDO, COD_USUARIO,COD_PEDIDO) VALUES ("&_
								  ROTINA_COD_FORMANDO&", "&COD_PRODUCAO&", "&sqlTexto(ROTINA_COD_CONTRATO)&", "&ROTINA_TIPO&", "&ROTINA_BRINDE&", "&_
								  ROTINA_COD_ULTIMO_VENDEDOR_FICHA&", "&ROTINA_VENDIDO&", "&Session("COD_USUARIO")&_
								  ", "&sqlTexto(ROTINA_COD_PEDIDO)&_								  
								  " ) "
							
							'rESPONSE.Write(sql)	   
							objConexao.Execute(sql)
							
					rs2.MoveNext()
					Wend
							
						
			rs2.Close
			Set rs2 = Nothing
					
				
				
		rs.MoveNext()
		Wend
		

rs.Close
Set rs = Nothing

Set rs=objConexao.Execute("Select descricao FROM TB_PRODUCAO WHERE COD_PRODUCAO = "&COD_PRODUCAO)
	If not rs.Eof Then
		DESCRICAO = rs(0)
	End If
rs.Close
Set rs = Nothing


FecharConexao()

%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Extrato de aproveitamento de produ&ccedil;&atilde;o | Portal Stillos</title>
<link rel="stylesheet" href="css/vendor/bootstrap.min.css">
<link rel="stylesheet" href="css/dashboard.css">
<style>
.month-report { max-width:1150px; margin:24px auto; padding:clamp(16px,4vw,36px); background:#fff; border:1px solid #e4dfeb; border-radius:14px; }
.month-tools,.month-header { display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:12px; }
.month-header { margin:20px 0; padding-bottom:14px; border-bottom:2px solid #b20b61; }
.month-header p { margin:0; }
.month-logo { width:150px; padding:10px; border-radius:8px; background:#292b40; }
.month-kpis { display:grid; grid-template-columns:repeat(auto-fit,minmax(min(100%,175px),1fr)); gap:12px; margin:12px 0 16px; }
.month-kpis > div { padding:12px; border:1px solid #e4dfeb; border-radius:10px; }
.month-kpis dd { margin:4px 0 0; font-weight:700; font-size:1.15em; }
.month-scroll { overflow:auto; }
.month-table { width:100%; min-width:520px; border-collapse:collapse; margin-top:10px; font-variant-numeric:tabular-nums; }
.month-table td,.month-table th { padding:8px; text-align:right; border-bottom:1px solid #e4dfeb; }
.month-table th:first-child { text-align:left; }
.month-table thead,.month-table tfoot { background:#f1edf6; }
.month-table tbody tr:nth-child(even) { background:#faf9fc; }
.month-note { margin:14px 0 0; color:#53566b; }
@media screen { .report-embedded .month-report { margin:12px auto; } }
.report-signatures { display:flex; flex-wrap:wrap; gap:24px; margin-top:28px; break-inside:avoid; }
.report-signatures span { flex:1 1 180px; border-top:1px solid #777; padding-top:4px; }
.month-table .report-text { text-align:left; }
.month-table th,.month-table td { overflow-wrap:anywhere; }
[hidden] { display:none !important; }
.month-report { max-width:1480px; }
.production-section { margin:20px 0; }
.production-section h2 { margin-bottom:10px; }
.production-table { min-width:1020px; }
.production-table td,.production-table th { text-align:center; }
.production-table td table { width:100%; border-collapse:collapse; }
.production-table td table td { padding:3px; }
.production-products th { font-weight:500; overflow-wrap:anywhere; }
.production-products td { width:1%; white-space:nowrap; }
.production-table thead td { background:#f1edf6; }
.production-table thead strong { font-weight:650; }
.production-rate { display:grid; gap:4px; }
.production-rate meter { width:100%; min-width:48px; height:12px; }
.production-filter { display:grid; gap:5px; width:min(100%,420px); margin:0 0 12px; }
.production-filter select { width:100%; min-width:0; }
.production-orders { color:#a31962; padding:0; border:0; background:none; font:inherit; text-decoration:underline; text-align:left; }
.production-dialog { width:min(480px,calc(100% - 24px)); padding:20px; border:1px solid #e4dfeb; border-radius:12px; }
.production-dialog::backdrop { background:#24273899; }
.production-dialog p { overflow-wrap:anywhere; }
.production-notes { color:#53566b; font-size:.9em; }
.production-cell-label { display:none; }
.production-table > thead :is(th,td) { overflow-wrap:normal; }
.production-table > tbody > tr > td { vertical-align:top; }
@media screen and (max-width:1150px) {
  .production-section .month-scroll { overflow:visible; }
  .production-table { display:block; min-width:0; }
  .production-table > thead { position:absolute; width:1px; height:1px; padding:0; overflow:hidden; clip-path:inset(50%); }
  .production-table > tbody { display:grid; gap:12px; }
  .production-table > tbody > tr { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); background:#fff; border:1px solid #e4dfeb; border-radius:10px; padding:4px 10px; }
  .production-table > tbody > tr > td { min-width:0; text-align:left; padding:9px 6px; }
  .production-table > tbody > tr > td > div[align] { text-align:left; }
  .production-table .production-wide { grid-column:1 / -1; }
  .production-cell-label { display:block; color:#53566b; font-size:.85em; font-weight:650; margin-bottom:4px; }
  .production-table .production-rate { max-width:160px; }
  .production-table .production-products { max-width:380px; }
  .production-table .production-products th { text-align:left; }
}
@page { size:A4 landscape; margin:10mm; }
@media print {
  body { margin:0; background:#fff; color:#111; font-size:10pt; line-height:1.3; }
  .month-report { margin:0; padding:6mm 7mm; max-width:none; border:0; border-radius:0; }
  .month-tools,.production-filter,.production-dialog { display:none; }
  .month-header { margin:0 0 5mm; padding-bottom:3mm; border-bottom:1px solid #555; break-after:avoid; }
  .month-header h1 { font-size:16pt; margin:0 0 1mm; }
  .month-logo { width:30mm; padding:0; background:transparent; filter:brightness(0); }
  .month-scroll { overflow:visible; }
  .production-orders { color:inherit; text-decoration:none; }
  .production-rate meter { display:none; }
  .production-section { margin:4mm 0; }
  .month-table { min-width:0; width:100%; table-layout:fixed; margin-top:0; }
  .month-table th,.month-table td { padding:1.8mm 1.5mm; border-bottom:1px solid #ccc; }
  .month-table thead { display:table-header-group; }
  .month-table thead th { border-top:1px solid #777; border-bottom:1px solid #777; }
  .month-table tfoot { display:table-row-group; }
  .month-table tr,.month-kpis,.report-signatures { break-inside:avoid; }
  .month-kpis { display:flex; flex-wrap:wrap; gap:2mm 6mm; margin:0 0 4mm; }
  .month-kpis > div { padding:0; border:0; flex:1 1 auto; }
  .month-kpis dt { font-size:9pt; }
  .month-kpis dd { font-size:10pt; margin:1mm 0 0; }
  .report-signatures { margin:12mm 0 0; gap:10mm; }
  .report-signatures span { flex:1 1 0; font-size:9pt; }
}
</style>
</head>
<body<%If Request.QueryString("modal") = "1" Then%> class="report-embedded"<%End If%>>
<main class="month-report" data-report-preview-content>
  <%If Request.QueryString("modal") <> "1" Then%>
  <nav class="month-tools" aria-label="A&ccedil;&otilde;es do relat&oacute;rio">
    <a class="btn btn-outline-secondary" href="adm_relatorios.asp">Voltar aos relat&oacute;rios</a>
    <button type="button" class="btn btn-primary" onclick="window.print()">Imprimir / Salvar PDF</button>
  </nav>
<%End If%>
  <header class="month-header">
    <div><h1>Extrato de aproveitamento de produ&ccedil;&atilde;o</h1><p><%=Server.HTMLEncode(COD_PRODUCAO & " - " & DESCRICAO)%></p></div>
    <img class="month-logo" src="images/logoparafundopreto.png" alt="Stillos">
  </header>
<%
		AbrirConexao()		
		Dim rs, sql, i, produtos()
		i = 0
		sql = "SELECT * FROM TB_PRODUCAO WHERE COD_PRODUCAO ="&COD_PRODUCAO
		Dim DESCRICAO, DATA_FINALIZADA, PRODUTOS_AGREDADOS
		
		
		Set rs = objConexao.Execute(sql, qtdRegistros)
		
		If not rs.Eof Then

		%>
        <section class="production-section"><h2>Resumo da produ&ccedil;&atilde;o</h2>
<div class="month-scroll" tabindex="0" role="group" aria-label="Resumo da produ&ccedil;&atilde;o"><table class="month-table production-table production-summary"><thead>
<tr bgcolor="#DDDDDD">
                <td width="14%" height="28" rowspan="2"><div align="center"><strong>Contratos</strong></div></td>
                <td width="15%" rowspan="2"><div align="center"><strong>Data Finalizada / Lotes</strong></div></td>
                <td colspan="2"><div align="center"><strong>Fotos da produ&ccedil;&atilde;o</strong></div></td>
                <td colspan="2"><div align="center"><strong>Fotos dos &aacute;lbuns</strong></div></td>                
                <td colspan="4"><div align="center"><strong>&Aacute;lbuns</strong></div></td>
                <td width="32%" rowspan="2"><div align="center"><strong>Venda de Produtos       </strong></div>
  
                  <%
				  Set rs4=objConexao.Execute("SELECT TB_PEDIDOS_PRODUTOS.COD_PRODUTO, TB_PRODUTOS.NOME_PRODUTO, SUM( TB_PEDIDOS_PRODUTOS.QUANTIDADE ) AS QUANTIDADE FROM TB_PEDIDOS INNER JOIN TB_PEDIDOS_PRODUTOS ON ( TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO )INNER JOIN TB_PRODUTOS ON ( TB_PEDIDOS_PRODUTOS.COD_PRODUTO = TB_PRODUTOS.COD_PRODUTO ) INNER JOIN TB_FICHAS_IDENTIFICACAO ON ( TB_PEDIDOS_PRODUTOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO ) WHERE TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO ="&COD_PRODUCAO&" AND TB_PEDIDOS.FINALIZA_VENDEDOR = 1 GROUP BY TB_PEDIDOS_PRODUTOS.COD_PRODUTO",QTD_PRODUTOS)
				While not rs4.Eof
					
						
						ReDim Preserve produtos(2,i)
                        PRODUTOS(0,i) = rs4("COD_PRODUTO")
						PRODUTOS(1,i) = rs4("NOME_PRODUTO")
						PRODUTOS(2,i) = rs4("QUANTIDADE")						
						i=i+1
					rs4.MoveNext()
					Wend
					
				  QTD_PRODUTOS = i
                  rs4.Close
				  Set rs4=Nothing
				  %>
&nbsp;  </td>
                <td width="6%" rowspan="2"><div align="center"><strong>Total vendas*</strong></div></td>
</tr>
<tr bgcolor="#DDDDDD">
  <td width="9%"><div align="center"><strong>Produzidas</strong></div></td>
  <td width="9%"><div align="center"><strong>N&atilde;o vendidas</strong></div></td>
  <td width="9%"><div align="center"><strong>N&atilde;o Vendidas</strong></div></td>
  <td width="6%"><div align="center"><strong>Vendidas</strong></div></td>
<%
Dim TOT_ALBUNS_TOTAL, TOT_ALBUNS_TOTAL_VENDIDO, TOT_ALBUNS_TOTAL_N_VENDIDO, TOT_ALBUNS_APROVEITAMENTO
Dim TOT_AGREGADOS_TOTAL, TOT_AGREGADOS_TOTAL_VENDIDO, TOT_AGREGADOS_TOTAL_N_VENDIDO, TOT_AGREGADOS_APROVEITAMENTO
Dim TOT_TOTAL_ARRECADADO, TOT_A_SER_ARRECADADO, TOT_APROVEITAMENTO_MEDIO
Dim TOTAL_GERAL_FOTOS_FATURADAS, TOTAL_GERAL_FOTOS_INICIAL, TOTAL_GERAL_FOTOS_APROVEITAMENTO
Dim TOTAL_GERAL_PAGAMENTO_EM_DINHEIRO, TOTAL_GERAL_XTOTAL_CHEQUE, TOTAL_GERAL_XTOTAL_CARTAO, TOTAL_GERAL

TOT_ALBUNS_TOTAL = 0
TOT_ALBUNS_TOTAL_VENDIDO = 0
TOT_ALBUNS_TOTAL_N_VENDIDO = 0
TOT_ALBUNS_APROVEITAMENTO = 0
TOT_AGREGADOS_TOTAL = 0
TOT_AGREGADOS_TOTAL_VENDIDO = 0
TOT_AGREGADOS_TOTAL_N_VENDIDO = 0
TOT_AGREGADOS_APROVEITAMENTO = 0
TOT_TOTAL_ARRECADADO = 0
TOT_A_SER_ARRECADADO = 0
TOT_APROVEITAMENTO_MEDIO = 0
TOTAL_GERAL_FOTOS_FATURADAS = 0
TOTAL_GERAL_FOTOS_INICIAL = 0
TOTAL_GERAL_FOTOS_APROVEITAMENTO = 0
%>
  <td width="4%"><div align="center" class="style3">Qtd. Total </div></td>
  <td width="5%"><div align="center" class="style3">Qtd. Vendida </div></td>
  <td width="5%"><div align="left" class="MensagemFicha style3">
    <div align="center">Qtd. n&atilde;o vendida </div>
  </div></td>
  <td width="4%"><div align="center" class="style3">Aproveitamento</div></td>
  </tr></thead><tbody>
		<%
		Dim ou, au
		ou = 0
		au = 0
		
		DESCRICAO			= rs("DESCRICAO")
		DATA_FINALIZADA		= rs("DATA_FINALIZADA")
		
		'Response.Write(DATA_FINALIZADA)
		PRODUTOS_AGREDADOS	= formatNumber(0&rs("PRODUTOS_AGREDADOS"),0)
		
		'Qtd total de Albuns
		Set rs2 = objConexao.Execute("SELECT COD_PRODUCAO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = "&COD_PRODUCAO,QTD_ALBUNS)
		
		rs2.Close
		Set rs2 = Nothing
		
		'Qtd Albuns Vendidos, nao precisa estar conferido
		Set rs3 = objConexao.Execute("SELECT TB_FICHAS_IDENTIFICACAO.COD_FORMANDO,TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO,TB_PEDIDOS_FORMANDOS.COD_PEDIDO,TB_PEDIDOS.DATA_VENDA,IF(TB_PEDIDOS_FORMANDOS.COD_PEDIDO is NULL,0,1) AS QTD_VENDIDO FROM TB_FICHAS_IDENTIFICACAO LEFT JOIN TB_PEDIDOS_FORMANDOS ON TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = TB_PEDIDOS_FORMANDOS.COD_FORMANDO LEFT JOIN TB_PEDIDOS ON TB_PEDIDOS_FORMANDOS.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO WHERE TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO = '"&COD_PRODUCAO&"' and IF(TB_PEDIDOS_FORMANDOS.COD_PEDIDO is NULL,0,1) = 1 GROUP BY COD_FORMANDO",QTD_ALBUNS_VENDIDOS)
				
		rs3.Close
		Set rs3 = Nothing
		
		Dim QTD_PRODUTOS_AGREGADOS_TOTAL, QTD_PRODUTOS_AGREGADOS_VENDIDOS, QTD_PRODUTOS_AGREGADOS_NAO_VENDIDOS
				
		
		Set rs5 = objConexao.Execute("SELECT * FROM TB_CACHE_VW_TOTAL_ARRECADADO_POR_PRODUCAO WHERE COD_PRODUCAO="&COD_PRODUCAO)
		Dim TOTAL_ARRECADADO
		TOTAL_ARRECADADO = 0			
					If not rs5.Eof Then
						TOTAL_ARRECADADO = formatNumber(0&rs5("TOTAL_ARRECADADO"))
					End If
					
		rs5.Close
		Set rs5 = Nothing
		
		'M?dia de Quanto Falta
		sql = "SELECT SUM(INICIAL) AS TOTAL_INICIAL FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = "&COD_PRODUCAO&" AND VENDIDO = 2 GROUP BY COD_PRODUCAO"
		
		Dim QTD_FOTOS_NAO_VENDIDAS
		QTD_FOTOS_NAO_VENDIDAS = 0
		Set rs6 = objConexao.Execute(sql)
			If not rs6.Eof Then
				QTD_FOTOS_NAO_VENDIDAS = formatNumber(0&rs6("TOTAL_INICIAL"),0)
			End If
		Set rs6 = Nothing

		
		
		Dim TOTAL_A_RECEBER_ALBUNS, TOTAL_A_RECEBER_PRODUTOS_AGREGADOS, TOTAL_NAO_ARRECADADO_GERAL
		TOTAL_A_RECEBER_ALBUNS				= 0
		TOTAL_A_RECEBER_PRODUTOS_AGREGADOS	= 0
		
		'1 de 2 totais a serem arrecadados - ALBUNS
		TOTAL_A_RECEBER_ALBUNS = QTD_FOTOS_NAO_VENDIDAS*23

		'2 de 2 totais arrecadados - PRODUTOS_AGREGAFDOS
		
		If PRODUTOS_AGREDADOS = 1 Then
					
			
					
			Set rs7 = objConexao.Execute("SELECT COD_PRODUTO, COUNT(*) AS QTD_PRODUTOS_AGREGADOS FROM TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS WHERE COD_PRODUCAO = "&COD_PRODUCAO&" AND VENDIDO <> 1 GROUP BY COD_PRODUTO, COD_PRODUCAO ")
			
				If not rs7.Eof Then
					Dim COD_PRODUTO_AGREGADO_1
					TOTAL_A_RECEBER_PRODUTOS_AGREGADOS = 0
					While not rs7.Eof
							
							COD_PRODUTO_AGREGADO_1 = formatNumber(rs7("COD_PRODUTO"),0)
							
								If COD_PRODUTO_AGREGADO_1 = 9 Then
									'Slide Show	
									TOTAL_A_RECEBER_PRODUTOS_AGREGADOS = TOTAL_A_RECEBER_PRODUTOS_AGREGADOS + (formatNumber(0&rs7("QTD_PRODUTOS_AGREGADOS"),0)*180)						
								Else
									'Outros
									TOTAL_A_RECEBER_PRODUTOS_AGREGADOS = TOTAL_A_RECEBER_PRODUTOS_AGREGADOS + (formatNumber(0&rs7("QTD_PRODUTOS_AGREGADOS"),0)*260)
								End If 
								
				
							
					rs7.MoveNext()
					Wend
					
					
					End If
			
			rs7.Close
			Set rs7 = Nothing
			
			'/////
			
			
			Set rs71 = objConexao.Execute("SELECT COUNT(*) AS QTD_PRODUTOS_AGREGADOS FROM TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS WHERE COD_PRODUCAO = "&COD_PRODUCAO&" AND VENDIDO = 2 GROUP BY COD_PRODUCAO")
					
					If not rs71.Eof Then
							
							QTD_NAO_VENDIDA_DEFINITIVAMENTE = formatNumber(rs71("QTD_PRODUTOS_AGREGADOS"),0)
							
					End If
					
					
			rs71.Close
			Set rs71 = Nothing
			'/////
			
		Else
			TOTAL_A_RECEBER_PRODUTOS_AGREGADOS = 0
		End if
				
			TOTAL_NAO_ARRECADADO_GERAL = (TOTAL_A_RECEBER_PRODUTOS_AGREGADOS+TOTAL_A_RECEBER_ALBUNS)/1
			
		Dim APROVEITAMENTO
		
		APROVEITAMENTO = 0
		If (TOTAL_ARRECADADO + TOTAL_NAO_ARRECADADO_GERAL) <> 0 Then APROVEITAMENTO = FormatNumber(100 * TOTAL_ARRECADADO / (TOTAL_ARRECADADO + TOTAL_NAO_ARRECADADO_GERAL), 1)
		
		'/////////////////////////////////////////////
		'FOTOS
		TOTAL_GERAL_FOTOS_FATURADAS = 0
		TOTAL_GERAL_FOTOS_INICIAL = 0
		TOTAL_GERAL_FOTOS_APROVEITAMENTO = 0

			Set rs5=objConexao.Execute("SELECT sum(INICIAL) AS INICIAL,SUM(FATURADO) AS FATURADO FROM VW_PEDIDOS_COMPLETO WHERE FINALIZA_VENDEDOR = 1 AND COD_PRODUCAO = "&COD_PRODUCAO)
				
				If not rs5.Eof Then
					
					TOTAL_GERAL_FOTOS_FATURADAS = formatNumber(0&rs5("FATURADO"),0)
					TOTAL_GERAL_FOTOS_INICIAL = formatNumber(0&rs5("INICIAL"),0)
					If TOTAL_GERAL_FOTOS_INICIAL > 0 Then
						TOTAL_GERAL_FOTOS_APROVEITAMENTO = clng(formatNumber(((100*TOTAL_GERAL_FOTOS_FATURADAS)/TOTAL_GERAL_FOTOS_INICIAL),1))
					Else
						TOTAL_GERAL_FOTOS_APROVEITAMENTO = 0	
					End If
				End If
				
			rs5.Close
			Set rs5=Nothing
		
		'/////////////////////////////////////////////
		'Total arrecadado
				TOTAL_GERAL_PAGAMENTO_EM_DINHEIRO = 0
				TOTAL_GERAL_XTOTAL_CHEQUE = 0
				TOTAL_GERAL_XTOTAL_CARTAO = 0
				TOTAL_GERAL = 0
				
				Set rs5=objConexao.Execute("SELECT SUM(TB_CACHE_VW_PEDIDOS_TOTAL_RECEBIDO.PAGAMENTO_EM_DINHEIRO) AS PAGAMENTO_EM_DINHEIRO, SUM(TB_CACHE_VW_PEDIDOS_TOTAL_RECEBIDO.XTOTAL_CHEQUE) AS XTOTAL_CHEQUE, SUM(TB_CACHE_VW_PEDIDOS_TOTAL_RECEBIDO.XTOTAL_CARTAO) AS XTOTAL_CARTAO FROM VW_PEDIDOS_COMPLETO INNER JOIN TB_CACHE_VW_PEDIDOS_TOTAL_RECEBIDO ON VW_PEDIDOS_COMPLETO.COD_PEDIDO = TB_CACHE_VW_PEDIDOS_TOTAL_RECEBIDO.COD_PEDIDO WHERE VW_PEDIDOS_COMPLETO.FINALIZA_VENDEDOR = 1 AND VW_PEDIDOS_COMPLETO.COD_PRODUCAO = "&COD_PRODUCAO)
				
				If not rs5.Eof Then
					
					TOTAL_GERAL_PAGAMENTO_EM_DINHEIRO = formatNumber(0&rs5("PAGAMENTO_EM_DINHEIRO"),2)/1
					TOTAL_GERAL_XTOTAL_CHEQUE = formatNumber(0&rs5("XTOTAL_CHEQUE"),2)/1
					TOTAL_GERAL_XTOTAL_CARTAO = formatNumber(0&rs5("XTOTAL_CARTAO"),2)/1
					TOTAL_GERAL = (TOTAL_GERAL_PAGAMENTO_EM_DINHEIRO+TOTAL_GERAL_XTOTAL_CHEQUE+TOTAL_GERAL_XTOTAL_CARTAO)/1
					
				End If
				
			rs5.Close
			Set rs5=Nothing
            
            TOTAL_FOTO_PRODUCAO = 0
            'Total de fotos na produ??o
				Set rs5=objConexao.Execute("SELECT SUM(INICIAL) AS INICIAL FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = "&COD_PRODUCAO)
				If not rs5.Eof Then
					TOTAL_FOTO_PRODUCAO = rs5("INICIAL")
                    TOTAL_FOTO_PRODUCAO_PARADA = clng(TOTAL_FOTO_PRODUCAO) - clng(TOTAL_GERAL_FOTOS_FATURADAS)
                    TOTAL_FOTO_PRODUCAO_PARADA =  formatNumber(0&TOTAL_FOTO_PRODUCAO_PARADA,0)
                    TOTAL_FOTO_PRODUCAO =  formatNumber(0&TOTAL_FOTO_PRODUCAO,0)                    
				End If
				
			rs5.Close
			Set rs5=Nothing
            
            
      

            
		
		%>     
				<tr bgcolor="whitesmoke">
		          <td><span class="production-cell-label" aria-hidden="true">Contratos</span>
                    <div align="center">
                      <%
			Dim rsContratos, qtdContratos, CONTRATOS_ULT
			
			Set rsContratos = objConexao.Execute("SELECT DISTINCT(COD_CONTRATO) AS COD_CONTRATO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO="&COD_PRODUCAO&" ORDER BY COD_CONTRATO ASC",qtdContratos)
			
			If not rsContratos.Eof then
			
				If qtdContratos > 1 then
			
					CONTRATOS_ULT = ""
					For i = 0 to qtdContratos-1
						
						If i = qtdContratos-1 then
						
							CONTRATOS_ULT = CONTRATOS_ULT&rsContratos("COD_CONTRATO")
						
						else
						
							CONTRATOS_ULT = CONTRATOS_ULT&rsContratos("COD_CONTRATO")&", "
						
						End If
					
					rsContratos.MoveNext()
					Next
					
					Response.Write(CONTRATOS_ULT)
					
				else
					
					Response.Write(rsContratos("COD_CONTRATO"))
				
				End If
				
					
					
			else
			
			Response.Write("-")
			
			End If
			
			rsContratos.Close()
			Set rsContratos = Nothing
			
			'Dim CONTRATOS_TRATAMENTO
			'CONTRATOS_TRATAMENTO = Left(CONTRATOS_ULT,Len(CONTRATOS_ULT)-2)&Replace(Right(CONTRATOS_ULT,2),", ","")
			'Response.Write(CONTRATOS_TRATAMENTO)
			%>                  
                    </div></td>
                  <td ><span class="production-cell-label" aria-hidden="true">Finaliza&ccedil;&atilde;o</span><div align="center"><%=ArrumaData(DATA_FINALIZADA)%></div></td>
                  <td ><span class="production-cell-label" aria-hidden="true">Fotos produzidas</span><div align="center"><%=TOTAL_FOTO_PRODUCAO%></div></td>
                  <td ><span class="production-cell-label" aria-hidden="true">Fotos da produ&ccedil;&atilde;o n&atilde;o vendidas</span><div align="center"><%=TOTAL_FOTO_PRODUCAO_PARADA%></div></td>
                  <td ><span class="production-cell-label" aria-hidden="true">Fotos dos &aacute;lbuns n&atilde;o vendidas</span><div align="center"><%=formatNumber(0&TOTAL_GERAL_FOTOS_INICIAL-TOTAL_GERAL_FOTOS_FATURADAS,0)%></div></td>
                  <td ><span class="production-cell-label" aria-hidden="true">Fotos dos &aacute;lbuns vendidas</span><div align="center"><%=TOTAL_GERAL_FOTOS_FATURADAS%></div></td>
                  <td ><span class="production-cell-label" aria-hidden="true">&Aacute;lbuns</span><div align="center"><%=QTD_ALBUNS%></div></td>
                  <td ><span class="production-cell-label" aria-hidden="true">&Aacute;lbuns vendidos</span>
                    <div align="center"><span class="style1"><%=QTD_ALBUNS_VENDIDOS%></span></div></td>
                  <td ><span class="production-cell-label" aria-hidden="true">&Aacute;lbuns n&atilde;o vendidos</span><div align="center"><span class="style2"><%=QTD_ALBUNS-QTD_ALBUNS_VENDIDOS%></span></div></td>
                  <td ><span class="production-cell-label" aria-hidden="true">Aproveitamento</span><div align="center">
				  <%
				  Dim APROVEITAMENTO_ALBUNS
				  		APROVEITAMENTO_ALBUNS = 0
                    If QTD_ALBUNS > 0 Then APROVEITAMENTO_ALBUNS = Int(100 * QTD_ALBUNS_VENDIDOS / QTD_ALBUNS + 0.5)
                    %><span class="production-rate"><%=FormatNumber(APROVEITAMENTO_ALBUNS, 0)%>%<meter min="0" max="100" value="<%= APROVEITAMENTO_ALBUNS %>" aria-label="Aproveitamento de &aacute;lbuns"></meter></span>
				  </div></td>
                  <td class="production-wide"><span class="production-cell-label" aria-hidden="true">Produtos vendidos</span>
                  
                  <%
				  If QTD_PRODUTOS > 0 Then
				  %>
                  <table class="production-products" aria-label="Produtos vendidos">
                    <%For i = 0 To QTD_PRODUTOS - 1%>
                      <tr><th scope="row"><%=Server.HTMLEncode(PRODUTOS(1,i) & "")%></th><td><%=PRODUTOS(2,i)%></td></tr>
                    <%Next%>
                  </table>
                  <%End If%>&nbsp;                  </td>
				  <td class="production-wide"><span class="production-cell-label" aria-hidden="true">Total de vendas estimado</span><div align="right"><%=formatNumber(TOTAL_GERAL,2)%></div></td>
				</tr>
				
	</tbody></table></div></section>
<dl class="month-kpis">
  <div><dt>&Aacute;lbuns vendidos</dt><dd><%=QTD_ALBUNS_VENDIDOS%> de <%=QTD_ALBUNS%></dd></div>
  <div><dt>Aproveitamento de &aacute;lbuns</dt><dd><%=FormatNumber(APROVEITAMENTO_ALBUNS, 0)%>%</dd></div>
  <div><dt>Total de vendas estimado</dt><dd><%=FormatCurrency(TOTAL_GERAL, 2)%></dd></div>
</dl>
<section class="production-section">
<h2>Desempenho por vendedor</h2>
<label class="production-filter" hidden>Vendedor
  <select id="production-seller" class="form-select"><option value="">Todos os vendedores</option></select>
</label>
<p id="production-filter-print" hidden></p>
<div class="month-scroll" tabindex="0" role="group" aria-label="Desempenho por vendedor">
<table class="month-table production-table" id="production-vendors">
<thead><tr><th scope="col">Vendedor e pedidos</th><th scope="col">Lotes</th><th scope="col">&Aacute;lbuns</th><th scope="col">Vendidos</th><th scope="col">N&atilde;o vendidos</th><th scope="col">Aproveitamento</th><th scope="col">Produtos vendidos</th><th scope="col">Total estimado</th></tr></thead>
<tbody>
<%
				sql79 = 	" SELECT TB_VENDEDORES.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR "&_
							" FROM TB_VENDEDORES "&_
							" INNER JOIN TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA ON "&_
							" TB_VENDEDORES.COD_VENDEDOR = TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA.COD_ULTIMO_VENDEDOR_FICHA "&_
							" WHERE COD_PRODUCAO ="&COD_PRODUCAO&_
							" GROUP BY COD_ULTIMO_VENDEDOR_FICHA ORDER BY TB_VENDEDORES.NOME_VENDEDOR"

				Set rs112 = objConexao.Execute(sql79)
				
				Dim TOTAL_QTD_ALBUNS
				TOTAL_QTD_ALBUNS = 0				
	
					'Totais
				
				Dim SUB_TOTAL_QTD_ALBUNS, SUB_TOTAL_QTD_ALBUNS_VENDIDOS, SUB_APROVEITAMENTO_ALBUNS
				Dim SUB_TOTAL_QTD_AGREGADOS, SUB_TOTAL_QTD_AGREGADOS_VENDIDOS, SUB_APROVEITAMENTO_AGREGADOS
				Dim SUB_APROVEITAMENTO_MEDIO
				Dim TOTAL_FOTOS_FATURADAS, TOTAL_FOTOS_INICIAL, TOTAL_FOTOS_APROVEITAMENTO	
				Dim TOTAL_PAGAMENTO_EM_DINHEIRO, TOTAL_XTOTAL_CHEQUE, TOTAL_XTOTAL_CARTAO, TOTAL					
								
				While not rs112.Eof
				
				SUB_TOTAL_QTD_ALBUNS = 0
				SUB_TOTAL_QTD_ALBUNS_VENDIDOS = 0
				SUB_APROVEITAMENTO_ALBUNS = 0
				SUB_TOTAL_QTD_AGREGADOS = 0
				SUB_TOTAL_QTD_AGREGADOS_VENDIDOS = 0
				SUB_APROVEITAMENTO_AGREGADOS = 0
				SUB_APROVEITAMENTO_MEDIO = 0
				
				
				
						sql = " SELECT COUNT(*) AS QTD_ALBUNS FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA WHERE "&_
							  " TIPO = 0 AND COD_ULTIMO_VENDEDOR_FICHA ="&rs112("COD_VENDEDOR")&" AND COD_USUARIO = '"&Session("COD_USUARIO")&"'"
						
						Set rsSubTotais = objConexao.Execute(sql)
								
								'Response.Write(sql&"<br />")
								
								If not rsSubTotais.Eof Then
								
									SUB_TOTAL_QTD_ALBUNS = formatNumber(0&rsSubTotais("QTD_ALBUNS"),0)
									'TOTAL_QTD_ALBUNS = TOTAL_QTD_ALBUNS + SUB_TOTAL_QTD_ALBUNS
									'Response.Write(rsSubTotais("QTD_ALBUNS")&"<br />")
								End If
								
						rsSubTotais.Close
						Set rsSubTotais = Nothing
						
						
						
						'//////////////
						
						sql = " SELECT COUNT(*) AS QTD_ALBUNS FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA WHERE "&_
							  " TIPO = 0 AND COD_ULTIMO_VENDEDOR_FICHA ="&rs112("COD_VENDEDOR")&" AND VENDIDO = 1 AND COD_USUARIO = '"&Session("COD_USUARIO")&"'"

						
						Set rsSubTotais = objConexao.Execute(sql)
								
								If not rsSubTotais.Eof Then
								
									SUB_TOTAL_QTD_ALBUNS_VENDIDOS = formatNumber(0&rsSubTotais("QTD_ALBUNS"),0)
									'TOTAL_QTD_ALBUNS = TOTAL_QTD_ALBUNS + SUB_TOTAL_QTD_ALBUNS_VENDIDOS
								End If
								
						rsSubTotais.Close
						Set rsSubTotais = Nothing
						
						
						'/////////////
						'Produtos Agregados Totais
						If PRODUTOS_AGREDADOS = 1 Then
						
							sql =	" SELECT COUNT(*) AS QTD_AGREGADOS FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA WHERE "&_
							  		" TIPO > 1 AND COD_ULTIMO_VENDEDOR_FICHA ="&rs112("COD_VENDEDOR")&" AND COD_USUARIO = '"&Session("COD_USUARIO")&"'"
						
							Set rsSubTotais = objConexao.Execute(sql)
								
								If not rsSubTotais.Eof Then
								
									SUB_TOTAL_QTD_AGREGADOS = formatNumber(0&rsSubTotais("QTD_AGREGADOS"),0)
									'TOTAL_QTD_ALBUNS = TOTAL_QTD_ALBUNS + SUB_TOTAL_QTD_AGREGADOS
								End If
								
							rsSubTotais.Close
							Set rsSubTotais = Nothing
							
							
							sql =	" SELECT COUNT(*) AS QTD_AGREGADOS FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA WHERE "&_
							  		" TIPO > 1 AND COD_ULTIMO_VENDEDOR_FICHA ="&rs112("COD_VENDEDOR")&" AND VENDIDO = 1 AND COD_USUARIO = '"&Session("COD_USUARIO")&"'"
						
							Set rsSubTotais = objConexao.Execute(sql)
								
								If not rsSubTotais.Eof Then
								
									SUB_TOTAL_QTD_AGREGADOS_VENDIDOS = formatNumber(0&rsSubTotais("QTD_AGREGADOS"),0)
									'TOTAL_QTD_ALBUNS = TOTAL_QTD_ALBUNS + SUB_TOTAL_QTD_AGREGADOS_VENDIDOS
								End If
								
							rsSubTotais.Close
							Set rsSubTotais = Nothing
						
						Else
						
							SUB_TOTAL_QTD_AGREGADOS  = "-"
							SUB_TOTAL_QTD_AGREGADOS_VENDIDOS = "-"
						
						End If
						
						
						'////////////////////				
						
						TOTAL_FOTOS_FATURADAS = 0
						TOTAL_FOTOS_INICIAL = 0
						TOTAL_FOTOS_APROVEITAMENTO = 0

						Set rs222=objConexao.Execute("SELECT sum(INICIAL) AS INICIAL,SUM(FATURADO) AS FATURADO FROM VW_PEDIDOS_COMPLETO WHERE VW_PEDIDOS_COMPLETO.FINALIZA_VENDEDOR = 1 AND COD_PRODUCAO = '"&COD_PRODUCAO&"' and COD_VENDEDOR = '"&rs112("COD_VENDEDOR")&"'")
							
							If not rs222.Eof Then
								
								TOTAL_FOTOS_FATURADAS = formatNumber(0&rs222("FATURADO"),0)
								TOTAL_FOTOS_INICIAL = formatNumber(0&rs222("INICIAL"),0)
								
								If int(TOTAL_FOTOS_FATURADAS) > 0 Then
									TOTAL_FOTOS_APROVEITAMENTO = formatNumber((100*TOTAL_FOTOS_FATURADAS)/TOTAL_FOTOS_INICIAL,1)
								End If
							End If
							
						rs222.Close
						Set rs222=Nothing
		
						
						Set rs222=objConexao.Execute("SELECT sum(INICIAL) AS INICIAL,SUM(FATURADO) AS FATURADO FROM VW_PEDIDOS_COMPLETO WHERE COD_PRODUCAO = '"&COD_PRODUCAO&"' and COD_VENDEDOR = '"&rs112("COD_VENDEDOR")&"' AND FINALIZA_VENDEDOR = 1")
							
							If not rs222.Eof Then
								
								TOTAL_ARRECADO_PRODUCAO = formatNumber(0&rs222("FATURADO"),0)
							End If
							
						rs222.Close
						Set rs222=Nothing
		
		'/////////////////////////////////////////////
		'Total arrecadado
				TOTAL_PAGAMENTO_EM_DINHEIRO = 0
				TOTAL_XTOTAL_CHEQUE = 0
				TOTAL_XTOTAL_CARTAO = 0
				TOTAL = 0
				
				Set rs5=objConexao.Execute("SELECT SUM(TB_CACHE_VW_PEDIDOS_TOTAL_RECEBIDO.PAGAMENTO_EM_DINHEIRO) AS PAGAMENTO_EM_DINHEIRO, SUM(TB_CACHE_VW_PEDIDOS_TOTAL_RECEBIDO.XTOTAL_CHEQUE) AS XTOTAL_CHEQUE, SUM(TB_CACHE_VW_PEDIDOS_TOTAL_RECEBIDO.XTOTAL_CARTAO) AS XTOTAL_CARTAO FROM VW_PEDIDOS_COMPLETO INNER JOIN TB_CACHE_VW_PEDIDOS_TOTAL_RECEBIDO ON VW_PEDIDOS_COMPLETO.COD_PEDIDO = TB_CACHE_VW_PEDIDOS_TOTAL_RECEBIDO.COD_PEDIDO WHERE VW_PEDIDOS_COMPLETO.COD_PRODUCAO = "&COD_PRODUCAO&" and VW_PEDIDOS_COMPLETO.COD_VENDEDOR = '"&rs112("COD_VENDEDOR")&"' AND VW_PEDIDOS_COMPLETO.FINALIZA_VENDEDOR = 1")
				
				If not rs5.Eof Then
					
					TOTAL_PAGAMENTO_EM_DINHEIRO = formatNumber(0&rs5("PAGAMENTO_EM_DINHEIRO"),2)/1
					TOTAL_XTOTAL_CHEQUE = formatNumber(0&rs5("XTOTAL_CHEQUE"),2)/1
					TOTAL_XTOTAL_CARTAO = formatNumber(0&rs5("XTOTAL_CARTAO"),2)/1
					TOTAL = (TOTAL_PAGAMENTO_EM_DINHEIRO+TOTAL_XTOTAL_CHEQUE+TOTAL_XTOTAL_CARTAO)/1
					
				End If
				
			rs5.Close
			Set rs5=Nothing	
			
			PEDIDOS = ""
			
			'Pedidos mostrar origem
			Set rs5=objConexao.Execute("SELECT COD_PEDIDO FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA WHERE COD_PRODUCAO = "&COD_PRODUCAO&" and COD_ULTIMO_VENDEDOR_FICHA = '"&rs112("COD_VENDEDOR")&"' AND COD_PEDIDO > 0 AND COD_USUARIO = "&Session("COD_USUARIO"))
				
				If not rs5.Eof Then
					While not rs5.Eof
					PEDIDOS = PEDIDOS&", "&rs5(0)
					rs5.MoveNext()
					Wend
				End If
				
			rs5.Close
			Set rs5=Nothing			
						
			If PEDIDOS = "" Then
				PEDIDOS = "Nenhum pedido deste vendedor"
			Else
				PEDIDOS = Mid(PEDIDOS,3,Len(PEDIDOS))
			End If			
						
				%>
				<tr data-production-seller="<%=rs112("COD_VENDEDOR")%>">
<td class="production-wide"><span class="production-cell-label" aria-hidden="true">Vendedor e pedidos</span><button type="button" class="production-orders" data-orders="<%=Server.HTMLEncode(PEDIDOS)%>"><%=Server.HTMLEncode(UCase(rs112("NOME_VENDEDOR") & ""))%></button></td>
<td class="production-wide"><span class="production-cell-label" aria-hidden="true">Lotes</span>
                            <div align="center">
                             <%
							Set rs8=objConexao.Execute("SELECT LOTE, DATA_SAIDA_INICIAL FROM TB_PRODUCAO_LOTES WHERE COD_PRODUCAO = "&COD_PRODUCAO&" AND COD_VENDEDOR_INICIAL = '"&rs112("COD_VENDEDOR")&"' order by DATA_SAIDA_INICIAL asc")
							
							If not rs8.Eof Then%>
                            <table width="100" border="0" cellspacing="0" cellpadding="2">
                          <tr>
                            <td><div align="center"><strong>Lote</strong></div></td>
                            <td><div align="center"><strong>Sa&iacute;da</strong></div></td>
                          </tr>

							

                            <%
							
								While not rs8.Eof
							%>
                           
                            <tr>
                            <td><div align="center"><%=rs8("LOTE")%></div></td>
                            <td><%=ArrumaData(rs8("DATA_SAIDA_INICIAL"))%></td>
                          </tr>
                            
                            <%
								rs8.MoveNext()
								Wend
							%>
							</table>
							<%
							Else%>
							-
                            
                            <%
							End If
							
							
							rs8.Close
							Set rs8=Nothing
							%>
                            </div></td>
                          <td><span class="production-cell-label" aria-hidden="true">&Aacute;lbuns</span><div align="center"><%=SUB_TOTAL_QTD_ALBUNS%></div></td>
                          <td><span class="production-cell-label" aria-hidden="true">Vendidos</span><div align="center"><%=SUB_TOTAL_QTD_ALBUNS_VENDIDOS%></div></td>
                          <td><span class="production-cell-label" aria-hidden="true">N&atilde;o vendidos</span><div align="center"><%=SUB_TOTAL_QTD_ALBUNS-SUB_TOTAL_QTD_ALBUNS_VENDIDOS%></div></td>
                          <td><span class="production-cell-label" aria-hidden="true">Aproveitamento</span><div align="center">
                            <%
				  			SUB_APROVEITAMENTO_ALBUNS = 0
                    If SUB_TOTAL_QTD_ALBUNS > 0 Then SUB_APROVEITAMENTO_ALBUNS = Int(100 * SUB_TOTAL_QTD_ALBUNS_VENDIDOS / SUB_TOTAL_QTD_ALBUNS + 0.5)
                    %><span class="production-rate"><%=FormatNumber(SUB_APROVEITAMENTO_ALBUNS, 0)%>%<meter min="0" max="100" value="<%= SUB_APROVEITAMENTO_ALBUNS %>" aria-label="Aproveitamento de &aacute;lbuns"></meter></span></div></td>
                          <td class="production-wide"><span class="production-cell-label" aria-hidden="true">Produtos vendidos</span>                  <%
				  If QTD_PRODUTOS > 0 Then
				  %>
                  <table class="production-products" aria-label="Produtos vendidos pelo vendedor">
                    <%For i = 0 To QTD_PRODUTOS - 1%>
                    <tr><th scope="row"><%=Server.HTMLEncode(PRODUTOS(1,i) & "")%></th><td><%					  
					  Set rs4=objConexao.Execute("SELECT TB_PEDIDOS_PRODUTOS.COD_PRODUTO, TB_PRODUTOS.NOME_PRODUTO, SUM( TB_PEDIDOS_PRODUTOS.QUANTIDADE ) AS QUANTIDADE FROM TB_PEDIDOS INNER JOIN TB_PEDIDOS_PRODUTOS ON ( TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_PRODUTOS.COD_PEDIDO )INNER JOIN TB_PRODUTOS ON ( TB_PEDIDOS_PRODUTOS.COD_PRODUTO = TB_PRODUTOS.COD_PRODUTO ) INNER JOIN TB_FICHAS_IDENTIFICACAO ON ( TB_PEDIDOS_PRODUTOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO ) WHERE TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO ="&COD_PRODUCAO&" AND TB_PEDIDOS.FINALIZA_VENDEDOR = 1 AND TB_PEDIDOS_PRODUTOS.COD_PRODUTO = "&PRODUTOS(0,i)&" AND TB_PEDIDOS.COD_VENDEDOR = "&rs112("COD_VENDEDOR")&" GROUP BY TB_PEDIDOS_PRODUTOS.COD_PRODUTO")
					  	
						If not rs4.Eof Then
							Response.Write(rs4("QUANTIDADE"))
						Else
							Response.Write("&mdash;")
						End If
						
					  rs4.Close
					  Set rs4=Nothing
					  
					  %>
					  </td></tr>
                    <%Next%>
                  </table>
                  <%End If%>&nbsp;</td>
				          <td class="production-wide"><span class="production-cell-label" aria-hidden="true">Total estimado</span><div align="right"><%=formatNumber(TOTAL,2)%></div></td></tr>
				          <%
				rs112.MoveNext()
				Wend
				
				
				rs112.Close
				Set rs112 = Nothing
				%> 
				        </tbody></table></div></section>
<section class="production-notes" aria-label="Notas sobre os dados">
  <p>O total de vendas &eacute; estimado. Pedidos com formandos de produ&ccedil;&otilde;es diferentes entram pelo valor integral.</p>
  <p>O relat&oacute;rio considera os dados registrados at&eacute; as 8h de hoje e inclui vendas ainda sem corre&ccedil;&atilde;o.</p>
</section>
<%Else%>
		   <p class="alert alert-info" role="status">Nenhum dado encontrado para esta produ&ccedil;&atilde;o.</p>
		    <%End If
			
			rs.Close
			Set rs = Nothing
			
			
			'Limpa novamente a TB tempor?ria
			'objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_TEMPORARIA WHERE COD_USUARIO ="&Session("COD_USUARIO"))
			FecharConexao()
			
			%>
		   
		    

		
<dialog class="production-dialog" id="production-orders-dialog" aria-labelledby="production-orders-title">
  <h2 id="production-orders-title">Pedidos do vendedor</h2>
  <p><strong id="production-orders-seller"></strong></p>
  <p id="production-orders-list"></p>
  <form method="dialog"><button class="btn btn-outline-secondary" autofocus>Fechar</button></form>
</dialog>
</main>
<script>
(function () {
    var table = document.getElementById('production-vendors');
    if (!table) return;
    var select = document.getElementById('production-seller');
    var scope = document.getElementById('production-filter-print');
    var rows = Array.from(table.querySelectorAll('[data-production-seller]'));
    var dialog = document.getElementById('production-orders-dialog');
    var seller = document.getElementById('production-orders-seller');
    var orders = document.getElementById('production-orders-list');

    rows.forEach(function (row) {
        var button = row.querySelector('.production-orders');
        select.add(new Option(button.textContent.trim(), row.dataset.productionSeller));
        button.addEventListener('click', function () {
            seller.textContent = button.textContent;
            orders.textContent = button.dataset.orders;
            dialog.showModal();
        });
    });
    function filterSellers() {
        rows.forEach(function (row) {
            row.hidden = Boolean(select.value && row.dataset.productionSeller !== select.value);
        });
        scope.hidden = !select.value;
        scope.textContent = select.value ? 'Vendedor: ' + select.options[select.selectedIndex].text : '';
    }
    if (rows.length) select.closest('label').hidden = false;
    select.addEventListener('change', filterSellers);
    window.addEventListener('pageshow', filterSellers);
    filterSellers();
})();
</script>
</body>
</html>
