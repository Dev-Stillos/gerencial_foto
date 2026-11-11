<%
Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
BTN_VOLTAR = 1
BTN_IMPRIMIR = 1
BTN_FONTE = 2

Dim cod_pedido
cod_pedido = Request.QueryString("cod_pedido")

BTN_IMPRIMIR_LINK = "href='lca_pedidos_pendentes_uniimpressao.asp?cod_pedido="&COD_PEDIDO&"' target='_blank'"
%>
<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim caminho_pagina, jquery
caminho_pagina = Replace(Request.ServerVariables("URL"),"/gerencial_foto/","")
jquery = True
%>
<!doctype html>
<html lang="pt-BR">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Detalhes do pedido | Portal Stillos</title>
<link rel="stylesheet" href="css/base.css" />
<link rel="stylesheet" href="css/chosen.min.css" />
<link rel="stylesheet" href="css/modal-message.css" />
<link rel="stylesheet" href="css/vendor/bootstrap.min.css" />
<link rel="stylesheet" href="css/dashboard.css" />
<!-- #include file="files/includes/scripts.asp" -->
<!-- #include file="files/includes/functions_local.asp" -->
<style>
.order-page .order-content { min-width:0; color:#292738; }
.order-page .order-heading { align-items:flex-start; }
.order-page .order-heading h1 { margin-bottom:0; }
.order-page .order-actions, .order-page .order-jump { display:flex; flex-wrap:wrap; gap:8px; }
.order-page .order-jump { margin:0 0 20px; }
.order-page .order-jump a { padding:6px 12px; border:1px solid #e4dfeb; border-radius:8px; background:#fff; color:#6d1749; font-weight:650; text-decoration:none; }
.order-page .order-content table[border="1"] { border:1px solid #e4dfeb; background:#fff; }
.order-page .order-content table[border="1"] > tbody > tr > td { padding:8px 10px; border-color:#eeeaf2; overflow-wrap:anywhere; }
.order-page .order-content tr[bgcolor="whitesmoke"] > td { background:#f1edf6; color:#3e344e; }
.order-page .order-content .order-summary { height:auto; margin:0; }
.order-page .order-summary h2, .order-page .order-section-heading { color:#292738; font:700 clamp(17px,2vw,23px)/1.3 system-ui,sans-serif; margin:0 0 12px; }
.order-page .order-section-heading { margin-top:28px; scroll-margin-top:18px; }
.order-page .order-summary table { width:100%; border-collapse:collapse; }
.order-page .order-summary table > tbody > tr > td:first-child:not([colspan]) { background:#f8f6fa; }
.order-page .order-scroll { max-width:100%; overflow-x:auto; }
.order-page .order-scroll table { min-width:1080px; }
.order-page .order-content :is(input,select,textarea,iframe,.btnGeral,table) { max-width:100%; }
.order-page .order-content .btnGeral { display:inline-block; padding:7px 10px; margin:3px 0; border:1px solid #ddd5e5; border-radius:7px; background:#fff; color:#7c174e; font-weight:650; white-space:normal; }
.order-page .order-content a:focus-visible { outline:2px solid #9e125c; outline-offset:2px; }
.order-page .dashboard-nav a.active:not([aria-current="page"]) { background:transparent; color:#d4d5e1; }
.order-page .dashboard-nav a[aria-current="page"] { background:#b20b61; color:#fff; font-weight:600; }
@media (max-width:650px) {
    .order-page .order-summary > table,
    .order-page .order-summary > table > tbody,
    .order-page .order-summary > table > tbody > tr,
    .order-page .order-summary > table > tbody > tr > td { display:block; width:100%; }
    .order-page .order-content { overflow-wrap:anywhere; }
}
</style>
</head>
<body class="order-page">
<div class="bloqueio_tela hidden"></div>
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<!-- #include file="includes/classe_fechamento.asp" -->
<%
Dim msg
Dim comando

msg = Request.QueryString("msg")
comando = Request.QueryString("comando")

Dim VALIDA_SUB21, VALIDA_SUB22, VALIDA_SUB32
VALIDA_SUB21 = formatNumber(0&inStr(sub_validacoes_paginas,"#21"),0)
VALIDA_SUB22 = formatNumber(0&inStr(sub_validacoes_paginas,"#22"),0)
VALIDA_SUB32 = formatNumber(0&inStr(sub_validacoes_paginas,"#32"),0)

AbrirConexao()

Dim NOME_VENDEDOR, DATA_VENDA, DATA_CADASTRO, FINALIZA_CORRECAO, FINALIZA_DIRETORIA, DATA_CORRIGIDO_DEP_FOTO, DATA_CHEGADA, DATA_DESCEU_DIRETORIA, DATA_CORRIGIDO_DIRETORIA, PEDIDO_EXTRA
Dim OBSERVACAO, VALOR_AJUDA_CUSTO, VENDEDOR_RECEBE_EM, PAGAMENTO_EM_DINHEIRO, ERRO_LISTA, FORMA_PAGAMENTO, TIPO_VENDA, PAGAMENTO_PROCESSADO, COD_GRUPO_VENDEDORES, CONFERIDO_POR_COD_USUARIO, NOME_USUARIO_CORRECAO_PEDIDO
Dim PRE_VENDA

sql = "SELECT "&_	
		"TB_VENDEDORES.NOME_VENDEDOR, "&_
		"TB_PEDIDOS.ERRO_LISTA, "&_
		"TB_PEDIDOS.FORMA_PAGAMENTO, "&_
		"TB_PEDIDOS.TIPO_VENDA, "&_
		"TB_PEDIDOS.PAGAMENTO_PROCESSADO, "&_
		"TB_PEDIDOS.COD_GRUPO_VENDEDORES, "&_
		"TB_PEDIDOS.PEDIDO_EXTRA, "&_
		"TB_PEDIDOS.DATA_VENDA, "&_
		"TB_PEDIDOS.CONFERIDO_POR_COD_USUARIO, "&_
		"TB_PEDIDOS.DATA_CHEGADA, "&_
		"TB_PEDIDOS.DATA_CADASTRO, "&_
		"TB_PEDIDOS.FINALIZA_CORRECAO, "&_
		"TB_PEDIDOS.FINALIZA_DIRETORIA, "&_
		"TB_PEDIDOS.DATA_CORRIGIDO_DEP_FOTO, "&_
		"TB_PEDIDOS.DATA_DESCEU_DIRETORIA, "&_
		"TB_PEDIDOS.DATA_CORRIGIDO_DIRETORIA, "&_
		"TB_PEDIDOS.VENDEDOR_RECEBE_EM, "&_
		"TB_PEDIDOS.OBSERVACAO, "&_
		"TB_PEDIDOS.VALOR_AJUDA_CUSTO, "&_
		"TB_PEDIDOS.PAGAMENTO_EM_DINHEIRO, "&_
		"TB_PEDIDOS.PRE_VENDA " &_
		"FROM TB_VENDEDORES INNER JOIN TB_PEDIDOS ON TB_VENDEDORES.COD_VENDEDOR = TB_PEDIDOS.COD_VENDEDOR "&_
		"WHERE COD_PEDIDO="&COD_PEDIDO
		
		Set rs=objConexao.Execute(sql)
		
		If not rs.Eof Then
		
			NOME_VENDEDOR				= rs("NOME_VENDEDOR")
			DATA_VENDA					= rs("DATA_VENDA")
			DATA_CADASTRO				= rs("DATA_CADASTRO")
			PEDIDO_EXTRA				= rs("PEDIDO_EXTRA")
			FINALIZA_CORRECAO			= rs("FINALIZA_CORRECAO")
			FINALIZA_DIRETORIA			= rs("FINALIZA_DIRETORIA")
			DATA_CORRIGIDO_DEP_FOTO		= ""&rs("DATA_CORRIGIDO_DEP_FOTO")
			DATA_DESCEU_DIRETORIA		= ""&rs("DATA_DESCEU_DIRETORIA")
			DATA_CORRIGIDO_DIRETORIA	= ""&rs("DATA_CORRIGIDO_DIRETORIA")
			DATA_DESCEU_DIRETORIA		= ""&rs("DATA_DESCEU_DIRETORIA")
			DATA_CHEGADA				= ""&rs("DATA_CHEGADA")
			OBSERVACAO					= ""&rs("OBSERVACAO")
			PRE_VENDA                   = rs("PRE_VENDA")
			CONFERIDO_POR_COD_USUARIO	= ""&rs("CONFERIDO_POR_COD_USUARIO")
			VALOR_AJUDA_CUSTO			= formatNumber(0&rs("VALOR_AJUDA_CUSTO"))
			PAGAMENTO_EM_DINHEIRO		= formatNumber(0&rs("PAGAMENTO_EM_DINHEIRO"))
			FORMA_PAGAMENTO				= formatNumber(0&rs("FORMA_PAGAMENTO"),0)
			TIPO_VENDA					= formatNumber(0&rs("TIPO_VENDA"),0)
			PAGAMENTO_PROCESSADO		= formatNumber(0&rs("PAGAMENTO_PROCESSADO"),0)
			COD_GRUPO_VENDEDORES		= formatNumber(0&rs("COD_GRUPO_VENDEDORES"),0)
			ERRO_LISTA					= ""&rs("ERRO_LISTA")
			VENDEDOR_RECEBE_EM			= rs("VENDEDOR_RECEBE_EM")
			
			'//Nome de quem conferiu a venda se j� foi conferida
			If int(FINALIZA_CORRECAO) = 1 Then
				Set rs2=objConexao.Execute("SELECT NOME FROM TB_ADM_USUARIOS WHERE COD_USUARIO = '"&CONFERIDO_POR_COD_USUARIO&"'")
					If not rs2.Eof Then
						NOME_USUARIO_CORRECAO_PEDIDO = rs2(0)
					End If
				rs2.Close
				Set rs2=Nothing
			End If
		End If

		rs.Close
		Set rs=Nothing
FecharConexao()

%> 

<a class="skip-link" href="#conteudo">Ir para o conte&uacute;do</a>
<div class="dashboard-layout">
  <aside class="dashboard-sidebar" id="menuPrincipal">
    <a class="dashboard-brand" href="adm_index.asp"><img src="images/logoparafundopreto.png" alt="Stillos" width="180" /><span>GERENCIAL FOTOS</span></a>
    <div class="nav-caption">&Aacute;REA DE TRABALHO</div>
    <!-- #include file="includes/dashboard_menu.asp" -->
  </aside>
  <div class="dashboard-workspace">
    <header class="dashboard-topbar">
      <button class="btn menu-toggle" id="ppvMenuToggle" type="button" aria-controls="menuPrincipal" aria-expanded="false">Menu</button>
      <div class="topbar-title">Portal Stillos <span>/ Detalhes do pedido</span></div>
      <div class="topbar-actions"><a class="user-name" href="adm_log_usuarios.asp"><%=Server.HTMLEncode(Session("NOME") & "")%></a><a href="manual_sistema.asp" target="_blank" rel="noopener">Manual</a><a class="btn btn-outline-secondary btn-sm" href="adm_sair.asp">Sair</a></div>
    </header>
    <main class="dashboard-main ppv-main" id="conteudo">
      <div class="order-content">
        <div class="dashboard-heading order-heading">
          <div>
            <p class="eyebrow">PEDIDOS</p>
            <h1>Pedido<%If int(PEDIDO_EXTRA) = 1 Then%> extra<%End If%> #<%=Server.HTMLEncode(COD_PEDIDO & "")%></h1>
          </div>
          <div class="order-actions">
            <a class="btn btn-outline-secondary" href="adm_index.asp" id="orderBack">Voltar</a>
            <a class="btn btn-primary" href="lca_pedidos_pendentes_uniimpressao.asp?cod_pedido=<%=Server.URLEncode(COD_PEDIDO & "")%>" target="_blank" rel="noopener">Imprimir pedido</a>
          </div>
        </div>
        <div class="search-panel"><div><strong>Busca no portal</strong></div><!-- #include file="includes/dashboard_busca.asp" --></div>
        <div class="session-notices"><!-- #include file="includes/dashboard_avisos.asp" --></div>
        <nav class="order-jump" aria-label="Se&ccedil;&otilde;es do pedido">
          <a href="#dados_pedido">Dados do pedido</a>
          <a href="#produtos_adquiridos">Produtos</a>
          <a href="#formas_pagamento">Pagamentos</a>
          <a href="#ordens_servico">Ordens de servi&ccedil;o</a>
        </nav>
        <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
        <section class="dashboard-panel order-summary" aria-labelledby="dados_pedido">
          <h2 id="dados_pedido">Dados do pedido</h2>
          <table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
 
  <tr>
            <td width="16%"><div align="left"><strong>Vendedor:</strong></div></td>
            <td width="84%">
			<%=NOME_VENDEDOR%></td>
            </tr>
			<tr>
			  <td><div align="left"><strong>Data prevista para pagamento: </strong></div></td>
			  <td>
			  <%If VENDEDOR_RECEBE_EM <> "" Then%>
			  <%=ArrumaData(VENDEDOR_RECEBE_EM)%>
              <%End If%>
              </td>
			</tr>
			<tr>
			  <td><div align="left"><strong>Data  da venda: </strong></div></td>
			  <td><%=ArrumaData(DATA_VENDA)%></td>
			  </tr>
              <tr>
			  <td><div align="left"><strong>Pagamento Processado: </strong></div></td>
			  	<td>
              	<%If int(PAGAMENTO_PROCESSADO) = 1 Then
					
					AbrirConexao()
						
						Set rs2=objConexao.Execute("SELECT DATA_PROCESSAMENTO FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_PEDIDO ="&COD_PEDIDO)
								
							If not rs2.Eof Then
									
								Response.Write(rs2(0))	
									
							Else
								Response.Write(Ucase(MonthName(Month(VENDEDOR_RECEBE_EM)))&" / "&Year(VENDEDOR_RECEBE_EM))
								'Response.Write("FALHA PAGAMENTO N�O PROCESSADO! CONTATE O ADMINISTRADOR DO SISTEMA!")		
							End If
						
						rs2.Close
						Set rs2=Nothing
					FecharConexao()
				%>
                
                	<%IF VALIDA_SUB32 <> 0 and Session("cod_usuario") = 16 Then%>
              		<br />
                    <div>ATEN&Ccedil;&Aacute;O O PAGAMENTO J&Aacute; FOI PROCESSADO, POR TANTO N&Atilde;O HAVER&Aacute; ALTERA&Ccedil;&Atilde;O NA COMISS&Atilde;O DO VENDEDOR. CUIDADO NAS ALTERA&Ccedil;&Otilde;ES CUJO  PAGAMENTO J&Aacute; FOI PROCESSADO. CASO EST&Aacute; CORRE&Ccedil;&Atilde;O GERAR ALTERA&Ccedil;&Atilde;O NA COMISS&Atilde;O DO VENDEDOR CADASTRE A DIVERG&Ecirc;NCIA PARA O PR&Oacute;XIMO PAGAMENTO.</div>
                    <a href="https://stillos.com.br/gerencial_foto/lca_pedidos_pendentes_alteracoes_part_1.asp?cod_pedido=<%=Server.URLEncode(COD_PEDIDO)%>">
                	<%End If%>
                
               <%Else%>
               N&Atilde;O.
			   <%If (VALIDA_SUB32 <> 0 and Cstr(""&DATA_DESCEU_DIRETORIA) = "") or (Cstr(""&DATA_DESCEU_DIRETORIA) <> "" And int(0&Session("cod_usuario")) = 16)  Then
			   
			   %> | &nbsp;<a href="https://stillos.com.br/gerencial_foto/lca_pedidos_pendentes_alteracoes_part_1.asp?cod_pedido=<%=COD_PEDIDO%>">Corrigir pedido.</a>
                   <%Else%>
                   		&nbsp;<a href="javascript:alert('Voc&ecirc; n&atilde;o possui acesso a corre&ccedil;&atilde;o de pedido e n&atilde;o pode corrigir um pedido que j&aacute; foi para diretoria. Se desejar alterar solicite a Diretoria/T.I com o login da mesma para clicar neste link.')">Corrigir pedido.</a>
				   <%End If%>
                <%End If%>              	
				</td>
			  </tr>
			<tr>
			  <td><div align="left"><strong>Data do lan&ccedil;amento da venda: </strong></div></td>
			  <td><%=ArrumaData(DATA_CADASTRO)%></td>
			  </tr>
              <tr>
			  <td><div align="left"><strong>Data da entrada na empresa: </strong></div></td>
			  <td><%=ArrumaData(DATA_CHEGADA)%></td>
			  </tr>
			<tr>
			  <td><div align="left"><strong>Data corrigido Depto. Foto </strong></div></td>
			  <td>
			  <%IF FINALIZA_CORRECAO = 1 Then%>
			  <%=ArrumaData(DATA_CORRIGIDO_DEP_FOTO)%> - <%=NOME_USUARIO_CORRECAO_PEDIDO%>
			  <%Else%>
			  N&atilde;o corrigido pelo departamento fotogr&aacute;fico
			  <%End If%>			  </td>
			  </tr>
			<tr>
			  <td><div align="left"><strong>Data que desceu para diretoria: </strong></div></td>
			  <td>
			  <%If(DATA_DESCEU_DIRETORIA) <> "" Then%>
				  	<%=ArrumaData(DATA_DESCEU_DIRETORIA)%>	  
			  <%Else%>
				  N&atilde;o desceu para diretoria
			  <%End IF%>			  </td>
			  </tr>
			<tr>
			  <td><div align="left"><strong>Data verificado pela Diretoria </strong></div></td>
			  <td>
			  <%If FINALIZA_DIRETORIA = 1 Then%>
			  	<%=DATA_CORRIGIDO_DIRETORIA%>
			  <%Else%>
			  	N&atilde;o corrigido pela diretoria
			  <%End If%>			  </td>
			  </tr>

			<tr>
			  <td><div align="left"><strong>Pr&eacute;-venda </strong></div></td>
			  <td>
			  <%If PRE_VENDA = "S" Then%>
			  	SIM
			  <%Else%>
			  	N&Atilde;O
			  <%End If%>			  </td>
			  </tr>



              <tr>
			  <td><div align="left"><strong>Tipo Venda</strong><a href="javascript:alert('Normal, Repasse ou Repasse Especial')">
<img src="images/icons/icon_help_mini.jpg" border="0"  />                    </a></div></td>
			  <td>
              
              
              <%If int(FINALIZA_CORRECAO) = 1 Then
			  			If int(TIPO_VENDA) = 1 Then
							Response.Write("Normal")
						Elseif int(TIPO_VENDA) = 2 Then
							Response.Write("Repasse")
						Elseif int(TIPO_VENDA) = 3 Then
							Response.Write("Repasse Especial")
						End If
				
				Else%>
              	N&atilde;o dispon&iacute;vel - somente quando o pedido for corrigido
              <%End If%>              </td>
			  </tr>
              <tr>
			  <td><div align="left"><strong>Forma de pagamento considerada</strong>:</div></td>
			  <td>
			
			  <%
					If int(FINALIZA_CORRECAO) = 1 Then
					
						If int(FORMA_PAGAMENTO) = 1 Then
							Response.Write("&Agrave; vista")
						Elseif int(FORMA_PAGAMENTO) = 2 Then
							Response.Write("&Agrave; prazo")					
						End If
					Else
					%>
                   N&atilde;o dispon&iacute;vel - somente quando o pedido for corrigido
                    <%
					End If
					%></td>
			  </tr>
              
              <%
			  If int(COD_GRUPO_VENDEDORES) > 0 Then
			  %>
              <tr>
			  <td><div align="left"><strong>Grupo Repasse</strong>:</div></td>
			  <td>
              <%
			  If int(COD_GRUPO_VENDEDORES)  =1 Then
			  %>
              RENATO
              <%
			  Elseif int(COD_GRUPO_VENDEDORES)  =2 Then
			  %>
              M&Aacute;RIO
              <%
			  Elseif int(COD_GRUPO_VENDEDORES)  =3 Then
			  %>
              PAULO MANTUAM
                        </td>
			  </tr>
			  <%
         Elseif int(COD_GRUPO_VENDEDORES)  =4 Then
			  %>
              VICENTE 
              <%End If%>              </td>
			  </tr>
			  <%
			  End If			  
			  %>
              
			  	<%If Len(OBSERVACAO) > 2 Then%>
			    	<tr>
						 <td><div align="left"><strong>Observa&ccedil;&atilde;o:</strong></div></td>
						 <td><div align="left"><%=OBSERVACAO%></div></td>
					</tr>
				<%End If%>
						 
			  
			   <%If VALIDA_SUB21 <> 0 and ERRO_LISTA <> "" Then%> 
				  <tr bgcolor="#F9D0D0">
						 <td><strong>Erro listado: </strong></td>
						 <td><%=ERRO_LISTA%></td>
						 </tr>
				  
			   <%End If%>
               <%
			   AbrirConexao()
			   
			   Set rs=objConexao.Execute("Select * From VW_TB_ADM_LOG_ALTERACOES where cod_pedido = '"&cod_pedido&"' ORDER BY DATA_CADASTRADO desc")
			   
				If not rs.Eof Then
			   %>
              		 <tr bgcolor="#F9D0D0">
				    <td><strong>Log de Altera&ccedil;&otilde;es:</strong></td>
				    <td>
                    	<table width="500" border="1" align="left" 	cellpadding="5" style="border-collapse:collapse">
  						<tr>
    					<td width="89"><div align="center"><strong>Usu&aacute;rio</strong></div></td>
    					<td width="276"><div align="center"><strong>Descri&ccedil;&atilde;o</strong></div></td>
                        <td width="89"><div align="center"><strong>Data da Alteração</strong></div></td>
  </tr>
  						<%While not rs.Eof%>
                        <tr>
  						  <td><div align="left"><%=rs("nome")%></div></td>
  						  <td><div align="left"><%=rs("DESCRICAO")%></div></td>
  						  <td><div align="center"><%=ArrumaData(rs("DATA_CADASTRADO"))%></div></td>
						  </tr>
                         <%
						 rs.Movenext()
						 Wend
						 %> 
</table>

                    
                    </td>
				    </tr>			
				<%
				End if
				
				rs.Close
				Set rs=Nothing
				FecharConexao()
				%>
            </table>
			</section>
            <h2 class="order-section-heading" id="produtos_adquiridos">Produtos adquiridos</h2>
				  <%
			  '1) Seleciona os c�digos dos formandos do pedido
			  '2) Seleciona os dados inseridos nas fichas
			  AbrirConexao()
			  'AbrirConexao2()
			  
			  Dim sql, rs
			  
		sql = 		"SELECT "&_
					"TB_PEDIDOS_FORMANDOS.COD_FORMANDO, "&_
					"TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, "&_
					"TB_FICHAS_IDENTIFICACAO.ALBUM, "&_
					"TB_FICHAS_IDENTIFICACAO.ENCADERNACOES_EXTRAS, "&_
					"TB_FICHAS_IDENTIFICACAO.COD_CONTRATO, "&_
					"TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO, "&_
					"TB_FICHAS_IDENTIFICACAO.NOME_CURSO, "&_
					"TB_FICHAS_IDENTIFICACAO.INICIAL, "&_
					"TB_FICHAS_IDENTIFICACAO.FATURADO, "&_
					"TB_FICHAS_IDENTIFICACAO.VALOR_UN, "&_
					"TB_FICHAS_IDENTIFICACAO.RETIRADO, "&_
					"TB_FICHAS_IDENTIFICACAO.COMISSAO, "&_								
					"TB_FICHAS_IDENTIFICACAO.BRINDE, "&_								
					"TB_FICHAS_IDENTIFICACAO.PRODUTO_AGREGADO, "&_								
					"TB_FICHAS_IDENTIFICACAO.COMISSAO "&_								
					"FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON "&_
					"TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO "&_
					"WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&COD_PEDIDO&" ORDER BY NOME_FORMANDO"
					
			  Set rs = objConexao.Execute(sql)
				Dim QTD_BRINDE, QTD_RETIRADA
				
													
					While not rs.Eof
					QTD_BRINDE = 0
					QTD_RETIRADA = 0
					
					Dim rsProducao, rsBrindesAlbum
				
				'Consulta a produ��o desta ficha para verificar qual o c�digo padr�o do �lbum
				Set rsProducao = objConexao.Execute("SELECT COD_FOTO_PADRAO FROM TB_PRODUCAO WHERE COD_PRODUCAO="&rs("COD_PRODUCAO"))
						If not rsProducao.Eof Then
						COD_FOTO_PADRAO = rsProducao("COD_FOTO_PADRAO")
						End If		
													
				rsProducao.Close()
				Set rsProducao = Nothing
				
				'Todos os pedidos que n�o tiverem foto padr�o a foto passar� a ser FOTO
				If COD_FOTO_PADRAO = Empty Then
					COD_FOTO_PADRAO = 2
				End If
				
				'Consulta a quantidade de brinde, agora com a foto padr�o do �lbum
				If COD_FOTO_PADRAO <> Empty Then
				Set rsBrindesAlbum = objConexao.Execute("SELECT Sum(QUANTIDADE) as QUANTIDADE FROM TB_FICHAS_IDENTIFICACAO_BRINDES WHERE COD_FORMANDO="&rs("COD_FORMANDO")&" AND COD_PRODUTO="&COD_FOTO_PADRAO&" GROUP BY COD_PRODUTO")
				
					If not rsBrindesAlbum.Eof then
						QTD_BRINDE = formatNumber(0&rsBrindesAlbum("QUANTIDADE"),0)
					Else
						QTD_BRINDE = 0
					End If
									
				rsBrindesAlbum.Close()
				Set rsBrindesAlbum = Nothing
				End If
				
			%>  
			<table width="100%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
			<tr bgcolor="whitesmoke">
    <td><div align="left" class="textoMedio"><strong><%=rs("NOME_FORMANDO")%>
	      <%If rs("COMISSAO") = 1 Then %>
	      - Membro de Comiss&atilde;o de Formatura
          <%end if%>
	    </strong></div></td>
    </tr>
<tr>
    <td width="100%">
	  
		<table width="98%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
  <tr>
    <td width="18%"><div align="center"><strong>C&oacute;digo Formando</strong></div></td>
                    <td width="26%"><div align="center"><strong>&Aacute;lbum</strong></div></td>
				    <td width="29%"><div align="center"><strong>Institui&ccedil;&atilde;o de Ensino </strong></div></td>
				    <td width="27%"><div align="center"><strong>Curso</strong></div></td>
                    </tr>
				
				  <tr>
				    <td><div align="center">
					<a href="dip_busca_detalhamento_formando.asp?cod_formando=<%=rs("COD_FORMANDO")%>">
					<%=digito(rs("COD_FORMANDO"))%>                    </a>
                    </div></td>
					  <td><div align="center"><%=rs("ALBUM")%> <% If int(rs("ENCADERNACOES_EXTRAS")) = 1 Then Response.Write(" DUPLO") End If%></div></td>
					  <td><div align="left">
					  <%
					Dim rs6
					Set rs6 = objConexao.Execute("SELECT DESCRICAO FROM TB_ENTIDADES WHERE  COD_CONTRATO= '"&rs("COD_CONTRATO")&"'")
						
						If not rs6.Eof Then
							Response.Write(rs6("DESCRICAO"))
						End If
						
					rs6.Close
					Set rs6 = Nothing
					%>
					  </div></td>
					  <td><div align="center"><%=rs("NOME_CURSO")%></div></td>
				      </tr>
				  </table>	  </td>
</tr>
<%If int(PEDIDO_EXTRA) = 2 Then%>
<tr bgcolor="whitesmoke">
    <td><div align="left"><strong>Dados do &Aacute;lbum </strong></div></td>
    </tr>
<tr>
  <td><table width="98%" border="1" align="center" cellpadding="2" cellspacing="2" style="border-collapse:collapse">
    <tr>
      <td width="15%"><div align="center"><strong>Qtd. Inicial</strong></div></td>
      <td width="13%"><div align="center"><strong>Qtd. Brinde</strong></div></td>
      <td width="15%"><div align="center"><strong>Qtd. Faturada</strong></div></td>
      <td width="14%"><div align="center"><strong>Qtd. Retirada </strong></div></td>
      <td width="15%"><div align="center"><strong>Valor Un.</strong></div></td>
      <td width="17%"><div align="center"><strong>Valor Total</strong></div></td>
      </tr>
    <tr>
      <td><div align="center"><%=rs("INICIAL")%></div></td>
      <td><div align="center"><%=QTD_BRINDE%></div></td>
      <td><div align="center">
	   <%
	  Dim  rsVALOR_UN, rsFATURADO, rsVALOR_TOTAL
	  rsVALOR_UN = formatNumber(0&rs("VALOR_UN"))
	  rsFATURADO = formatNumber(0&rs("FATURADO"),0)
	  rsVALOR_TOTAL = rsVALOR_UN * rsFATURADO
	  
	  QTD_RETIRADA = (0&rs("INICIAL")) - (QTD_BRINDE/1 + rsFATURADO/1)
	  %>
	  
      <%If rsFATURADO > 0 then %>
	  	<%=rsFATURADO%>
	  <%Else%>
	  	---
	  <%End If%>
	  </div></td>
      <td><div align="center"><%=QTD_RETIRADA%></div></td>
      <td><div align="center">
	  <%If rsVALOR_UN > 0 then %>
	  	<%=formatNumber(rsVALOR_UN)%>
	  <%Else%>
	  	---
	  <%End If%>
	  </div></td>
      <td>

        <div align="right">
          <%If rsVALOR_TOTAL > 0 then 
			
		%>
          <%=formatNumber(rsVALOR_TOTAL)%>
              <%Else%>
          ---
          <%End If%>
        </div></td>
      </tr>

  </table></td>
</tr>
<%
Dim rsBrindes


If COD_FOTO_PADRAO <> Empty Then
sql = " SELECT Sum(TB_FICHAS_IDENTIFICACAO_BRINDES.QUANTIDADE) as QUANTIDADE, "&_
									" TB_PRODUTOS.NOME_PRODUTO FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_BRINDES "&_
									" ON TB_PRODUTOS.COD_PRODUTO = TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO "&_
									" WHERE TB_FICHAS_IDENTIFICACAO_BRINDES.COD_FORMANDO="&rs("COD_FORMANDO")&" AND TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO <> "& COD_FOTO_PADRAO&_
									" GROUP BY TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO "
									
Else
sql = " SELECT Sum(TB_FICHAS_IDENTIFICACAO_BRINDES.QUANTIDADE) as QUANTIDADE, "&_
									" TB_PRODUTOS.NOME_PRODUTO FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_BRINDES "&_
									" ON TB_PRODUTOS.COD_PRODUTO = TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO "&_
									" WHERE TB_FICHAS_IDENTIFICACAO_BRINDES.COD_FORMANDO="&rs("COD_FORMANDO")&" GROUP BY TB_FICHAS_IDENTIFICACAO_BRINDES.COD_PRODUTO "
End If									

Set rsBrindes = objConexao.Execute(sql)
If not rsBrindes.Eof Then
%>
<tr bgcolor="whitesmoke">
  <td><div align="left"><strong>Outros Brindes</strong></div></td>
</tr>
<tr>
  <td><table width="98%" border="1" align="center" 	cellpadding="2" cellspacing="2" style="border-collapse:collapse">
    <tr>
      <td width="56%"><div align="center"><strong>Produto</strong></div></td>
      <td width="44%"><div align="center"><strong>Qtd</strong>.</div></td>
      </tr>
	   <%
	   While not rsBrindes.Eof
	   %>
    	<tr>
      <td><div align="left"><%=rsBrindes("NOME_PRODUTO")%></div></td>
      <td><div align="center"><%=rsBrindes("QUANTIDADE")%></div></td>
      </tr>
	   <%
	   rsBrindes.MoveNext()
	   Wend
	   %>
  </table></td>
</tr>
<%
End If

rsBrindes.Close()
Set rsBrindes = Nothing
%>
<%
Dim rsAgregados
sql =	" SELECT TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO_AGREGADO, "&_
		" Sum(TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.QUANTIDADE) as QUANTIDADE, "&_
		" TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.VENDIDO, TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.VALOR_UN, "&_
		" TB_PRODUTOS.NOME_PRODUTO FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS "&_
		" ON TB_PRODUTOS.COD_PRODUTO = TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO "&_
		" WHERE TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO="&rs("COD_FORMANDO")&" AND TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.VENDIDO = 1"&_
		" GROUP BY TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO "
'Response.Write(sql)

Set rsAgregados = objConexao.Execute(sql)

If not rsAgregados.Eof Then
%>
<tr bgcolor="whitesmoke">
  <td><div align="left"><strong>Produtos Agregados com sa&iacute;da de produ&ccedil;&atilde;o</strong><br />
  </div></td>
</tr>
<tr>  <td>	
		<form style="margin:0px" method="post" name="AGREGADOS_<%=rs("COD_FORMANDO")%>">
			<table width="98%" border="1" align="center" cellpadding="2" cellspacing="2" style="border-collapse:collapse">
    <tr>
      <td width="35%"><div align="center"><strong>Produto</strong></div></td>
      <td width="21%"><div align="center"><strong>Qtd</strong>.</div></td>
      <td width="17%"><div align="center"><strong>Vendido</strong></div></td>
      <td width="16%"><div align="center"><strong>Valor Unit&aacute;rio </strong></div></td>
      </tr>
	<%
	Dim contProdAgregado, qteCod, qtaUi
	contProdAgregado = 1
	qteCod = formatNumber(0&Request.QueryString("prod_agre_alterar"),0)
	qtaUi  = formatNumber(0&rsAgregados("COD_PRODUTO_AGREGADO"),0)
	While not rsAgregados.Eof
	
		contProdAgregado = contProdAgregado + 1
	%>
    <tr>
      <td><div align="left"><%=rsAgregados("NOME_PRODUTO")%></div></td>
      <td><div align="center"><%=rsAgregados("QUANTIDADE")%></div></td>
      <td><div align="center">
	  <%If rsAgregados("VENDIDO") = 3 then
	  CHAVE_AVANCAR = 2
	  %>
	  <%else%>
	  		<%If Request.QueryString("prod_agre_alterar") <> Empty Then
			CHAVE_AVANCAR = 2
			%>
	  		<%Else%>
				<%If rsAgregados("VENDIDO") = 1 Then%>
					Sim
				<%Else%>
					N&atilde;o
				<%End If%>	
	  		<%End If%>
	  <%End If%>
	  </div></td>
      <td>
	  
	    <div align="right">
	      <%If Request.QueryString("prod_agre_alterar") <> Empty Then%>
	          <%Else%>
	      <%If rsAgregados("VENDIDO") = 3 then%>
	      <%elseif rsAgregados("VENDIDO") = 2 then%>
	      ---
	      <%else%>
	        <%=formatNumber(0&rsAgregados("VALOR_UN"))%>
	        <%End If%>
	                  <%End If
	  
	  
	  %>
	        </div></td>
      </tr>
	<%
	rsAgregados.MoveNext()
	Wend
	%>
  </table>
  		</form>
  </td>
</tr>	
<%End If

rsAgregados.Close()
Set rsAgregados = Nothing

%>
<%End If%>

<%
Dim rsProdutosAparte
Set rsProdutosAparte = objConexao.Execute(" SELECT TB_PEDIDOS_PRODUTOS.COD_PEDIDO_PRODUTO, TB_PRODUTOS.NOME_PRODUTO, TB_PEDIDOS_PRODUTOS.VALOR_UN, TB_PRODUTOS.necessita_fotos, "&_
										  " TB_PEDIDOS_PRODUTOS.QUANTIDADE FROM TB_PRODUTOS INNER JOIN "&_
										  " TB_PEDIDOS_PRODUTOS ON TB_PRODUTOS.COD_PRODUTO = TB_PEDIDOS_PRODUTOS.COD_PRODUTO"&_
										  " WHERE COD_PEDIDO="&COD_PEDIDO&" AND COD_FORMANDO="&rs("COD_FORMANDO")&" ORDER BY NOME_PRODUTO")
If not rsProdutosAparte.Eof Then
%>
<tr bgcolor="whitesmoke">
  <td><div align="left"><strong><%If int(PEDIDO_EXTRA) = 2 Then%>Produtos vendidos &agrave; parte<%Else%>Produtos do pedido extra<%End If%></strong></div></td>
</tr>
<tr>
  <td><table width="98%" border="1" align="center" 	cellpadding="2" cellspacing="2" style="border-collapse:collapse">
    <tr>
      <td width="35%"><div align="center"><strong>Produto</strong></div></td>
      <td width="21%"><div align="center"><strong>Qtd.</strong></div></td>
      <td width="17%"><div align="center"><strong>Valor Unit&aacute;rio</strong></div></td>
      <td width="16%"><div align="center"><strong>Valor Total </strong></div></td>
      </tr>
    <%While not rsProdutosAparte.Eof%>
	<tr>
      <td><div align="left">
      	<%=rsProdutosAparte("NOME_PRODUTO")%>
      	<%
      		if rsProdutosAparte("necessita_fotos") = 1 then

      			set rsFotos = objConexao.Execute("SELECT * FROM TB_PEDIDOS_PRODUTOS_ARQUIVOS WHERE COD_PEDIDO_PRODUTO = "&rsProdutosAparte("COD_PEDIDO_PRODUTO"))

      			if not rsFotos.Eof then
      				Response.Write ("<br>Baixar arquivo(s)")
	      			While not rsFotos.Eof
	      				Response.Write ("<br><a href='lca_busca_detalhamento_pedido_download_arquivo.asp?p0="&rsFotos("id")&"' target='_blank'>-"&rsFotos("descricao")&"</a>")
	      			rsFotos.MoveNext
	      			Wend
	      		end if
      			rsFotos.close
      			set rsFotos = Nothing
      		end if
      	%>

      </div></td>
      <td><div align="center"><%=rsProdutosAparte("QUANTIDADE")%></div></td>
      <td><div align="center"><%=formatNumber(rsProdutosAparte("VALOR_UN"))%></div></td>
      <td>
	    <div align="right">
	      <%
	  Dim tot1, tot2
	 tot1 = formatNumber(0&rsProdutosAparte("QUANTIDADE"),0)
     tot2 = 0&rsProdutosAparte("VALOR_UN")
	  %>
	          <%=formatNumber(tot1*tot2)%></div></td>
      </tr>
	<%
	rsProdutosAparte.MoveNext()
	Wend
	%>
  </table></td>
</tr>
<%
End If

rsProdutosAparte.Close()
Set rsProdutosAparte = Nothing
%>
</table>
			<br />

			<%
			rs.MoveNext()
			Wend
				
			rs.Close()
			Set rs = Nothing
			'FecharConexao2()d
		  	FecharConexao()
  			%>
	  	      <h2 class="order-section-heading" id="formas_pagamento">Formas de pagamento</h2>
	  	     <table width="100%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
                <%if PAGAMENTO_EM_DINHEIRO > 0 then%>
			    <tr bgcolor="whitesmoke">
                  <td><strong>Dinheiro</strong></td>
                </tr>
                <tr>
                  <td>
				  <table width="98%" border="0" align="center" cellpadding="5">
						  <tr>
						    <td width="100%"><table width="210" border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65%"><div align="left" class="textoMedio">Total em Dinheiro:</div></td>
                                <td width="35%"><div align="right" class="textoMedio">
                                  <div align="right" class="textoMedio"><strong><%=PAGAMENTO_EM_DINHEIRO%></strong></div>
                                </div></td>
                              </tr>
                            </table></td>
						    </tr>
						</table>
						</td>
                </tr>
				<%End If%>
				<%
					Dim rsCartao
					
					AbrirConexao()
					
					Set rsCartao = objConexao.Execute("SELECT * FROM TB_PEDIDOS_CARTAO WHERE COD_PEDIDO="&COD_PEDIDO&" ORDER BY NOME_TITULAR") 
					if Not rsCartao.Eof then
					%>
                <tr bgcolor="whitesmoke">
                  <td><strong>Cart&atilde;o de Cr&eacute;dito</strong></td>
                </tr>
                <tr>
                  <td>
  <table width="98%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
                      <tr>
                        <td width="23%"><div align="center"><strong>Nome Titular </strong></div></td>
                        <td width="16%"><div align="center"><strong>Cart&atilde;o</strong></div></td>
                        <td width="16%"><div align="center"><strong>Autentica&ccedil;&atilde;o </strong></div></td>
                        <td width="15%"><div align="center"><strong>Documento</strong></div></td>
                        <td width="14%"><div align="center"><strong>Parcelas </strong></div></td>
                        <td width="16%"><div align="center"><strong>Valor Total </strong></div></td>
                        </tr>
                      <%
					  While not rsCartao.Eof
					  %>
                      <tr>
                        <td><div align="left"><%=rsCartao("NOME_TITULAR")%></div></td>
                        <td><div align="center"><%=rsCartao("CARTAO")%></div></td>
                        <td><div align="center"><%=rsCartao("AUTENTICACAO")%></div></td>
                        <td><div align="center"><%=rsCartao("DOCUMENTO")%></div></td>
                        <td><div align="center"><%=rsCartao("PARCELAS")%></div></td>
                        <td><div align="right"><%=formatNumber(rsCartao("VALOR_TOTAL"))%>
                              <%
						PARCIAL_CARTAO=formatNumber(rsCartao("VALOR_TOTAL"))
						TOTAL_CARTAO=(TOTAL_CARTAO+PARCIAL_CARTAO)/1
						PARCIAL_CARTAO=0						
						%>
                        </div></td>
                        </tr>
                      <%
					  rsCartao.MoveNext
					  Wend
					  
					  %>
                      <tr>
                        <td colspan="6">			
						<table width="210" border="0" cellspacing="0" cellpadding="0">
						  <tr>
						    <td width="62%"><div align="left" class="textoMedio">Total em Cart&atilde;o:</div></td>
						    <td width="38%"><div align="right" class="textoMedio"><strong><%=formatNumber(TOTAL_CARTAO)%></strong></div></td>
						  </tr>
						</table></td>
                        </tr>
                    </table>
	
					</td>
                </tr>
				<%
					end if
					
					rsCartao.Close
					Set rsCartao=Nothing

					Dim rsCheque
					
										
					Set rsCheque = objConexao.Execute("SELECT * FROM TB_PEDIDOS_CHEQUE WHERE COD_PEDIDO="&COD_PEDIDO&" ORDER BY BOM_PARA ASC") 
					
					if Not rsCheque.Eof then
					
					%>
                          <tr bgcolor="whitesmoke">

                  <td><strong>Cheque pr&eacute;-datado</strong></td>
                </tr>
                <tr>
                  <td><table width="98%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
                      <tr>
                        <td width="17%"><div align="center"><strong>Nome Titular </strong></div></td>
                        <td width="13%"><div align="center"><strong>CPF</strong></div></td>
                        <td width="9%"><div align="center"><strong>Banco</strong></div></td>
                        <td width="15%"><div align="center"><strong>C. Corrente </strong></div></td>
                        <td width="10%"><div align="center"><strong>N&ordm; Cheque </strong></div></td>
                        <td width="10%"><div align="center"><strong>Bom Para</strong></div></td>
                        <td width="10%"><div align="center"><strong>Se terceiro </strong></div></td>
                        <td width="10%"><div align="center"><strong>Historico</strong></div></td>
                        <td width="16%"><div align="center"><strong>Valor Cheque </strong></div></td>
                        </tr>
                      <%
					  TOTAL_CHEQUE = 0
					  Dim idParcial2, qtdChequesa
					  idParcial2 = 0
					  qtdChequesa = 0
					  While Not rsCheque.Eof
					  qtdChequesa = qtdChequesa + 1
					  idParcial2 = formatNumber(rsCheque("COD_PEDIDO_CHEQUE"))
					  %>
                      <tr>
                        <td><div align="left"><%=rsCheque("NOME")%></div></td>
                        <td><div align="center"><%=mascaraCPF(rsCheque("CPF"))%></div></td>
                        <td><div align="center"><%=rsCheque("BANCO")%></div></td>
                        <td><div align="center"><%=rsCheque("CONTA_CORRENTE")%></div></td>
                        <td><div align="center"><%=rsCheque("NUMERO_CHEQUE")%></div></td>
                        <td><div align="center"><%=ArrumaData(rsCheque("BOM_PARA"))%></div></td>
                        <td><div align="center"><%=rsCheque("SE_TERCEIRO")%></div></td>

                        <td><div align="center"><%
						Set rs3 = objConexao.Execute("SELECT count(*) FROM TB_PEDIDOS_CHEQUE_HISTORICO WHERE COD_PEDIDO_cheque = "&sqlNumero(rsCheque(0)))
						if not rs3.Eof Then
							qtdHistorico = rs3(0)
						End if
						rs3.Close
						set rs3 = Nothing

						Response.Write("<a href='javascript:historicoCheque("&rsCheque("COD_PEDIDO")&","&rsCheque("COD_PEDIDO_CHEQUE")&")'> Visualizar ("&qtdHistorico&") </a>")

                        %></div></td>

                        <td><div align="right"><%=FormatNumber(rsCheque("VALOR"))%><%
						PARCIAL_CHEQUE = formatNumber(rsCheque("VALOR"))
						TOTAL_CHEQUE=TOTAL_CHEQUE+formatNumber(PARCIAL_CHEQUE)
						PARCIAL_CHEQUE = ""
						%></div></td>
                        </tr>
                      <%
					  rsCheque.MoveNext
					  Wend
					  
					  %>
                      <tr>
                        <td colspan="9">
						<table width="400" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="33%"><div align="left" class="textoMedio">Total em Cheque:</div></td>
    <td width="20%"><div align="right" class="textoMedio"><strong><%=formatNumber(TOTAL_CHEQUE)%></strong></div></td>
    <td width="47%">&nbsp;&nbsp; |&nbsp; &nbsp;<span class="textoMedio">Qtd. cheque(s): <strong><%=qtdChequesa%></strong></span></td>
  </tr>
</table>						</td>
                        </tr>
                    </table></td>
                </tr>
				<%
					  rsCheque.Close
					  Set rsCheque = Nothing
					  end if
					  %>

	
				<!-- PAGAMENTO EM BOLETO -->
				<%
				Set rsBoleto = objConexao.Execute("SELECT * FROM TB_PEDIDOS_BOLETOS WHERE COD_PEDIDO = "&sqlNumero(COD_PEDIDO))
				if not rsBoleto.Eof then
					%>
					<tr bgcolor="whitesmoke">
						<td align="center" colspan="10"><strong>Pagamentos com Boleto</strong>
						
						<% if true then %>
							<a href='cob_pedidos_boletos_formulario.asp?comando=alterar&cod_pedido=<%=cod_pedido%>'>Alterar ✍️</a>
						<% end if %>

						</td>
					</tr>
					<tr>
					<td>
						<%
Function Nz(valor)
    If IsNull(valor) Or Trim(valor & "") = "" Then
        Nz = 0
    Else
        Nz = valor
    End If
End Function
%>

<table width="98%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
    <tr>
        <th>Parc.</th>
        <th>Vencimento</th>
        <th>Data Pagamento</th>
        <!--<th>Baixa</th>-->
        <th>Valor</th>
        <th>Valor Pago</th>
        <%if validaUsuario(1,360,0,true) then%>
            <th>Valor Comiss&atilde;o</th>
        <%end if%>
        <th>Recibos</th>
        <th>Status</th>
    </tr>

<%
totalBoleto = 0
while not rsBoleto.Eof
    valor = CDbl(Nz(rsBoleto("valor")))
    valorPago = CDbl(Nz(rsBoleto("valor_pago")))

    totalBoleto = totalBoleto + valor

    ' Define cor e texto do status
    If valorPago > 0 Then
        statusCor = "green"
        statusTexto = "Pago"
    Else
        statusCor = "red"
        statusTexto = "Pendente"
    End If
%>

    <tr>
        <td align="center"><%=rsBoleto("parcela")%></td>
        <td align="center"><%=arrumaData(rsBoleto("data_vencimento"))%></td>
        <td align="center"><%=arrumaData(rsBoleto("data_pagamento"))%></td>
        <!--<td align="center"><%=arrumaData(rsBoleto("data_baixa"))%></td>-->
        <td align="center">R$ <%=arrumaNumero(rsBoleto("valor"),"0,00",2)%></td>
        <td align="center">R$ <%=arrumaNumero(rsBoleto("valor_pago"),"0,00",2)%></td>
        <%if validaUsuario(1,360,0,true) then%>
            <td align="center">R$ <%=arrumaNumero(rsBoleto("valor_comissao_vendedor"),"0,00",2)%></td>
        <%end if%>
        <td align="center">
			<%if valorPago > 0 Then%>
            <a href="recibo_pedido_individual_boleto.asp?cod_pedido=<%=COD_PEDIDO%>&parcela=<%=rsBoleto("parcela")%>" target="_blank">
                Ver Recibo 🖨️
            </a>
			<%end if%>
        </td>
        <td align="center">
            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background-color:<%=statusCor%>;" title="<%=statusTexto%>"></span>
        </td>
    </tr>

<%
    rsBoleto.Movenext
Wend
%>

<tr bgcolor="whitesmoke">
    <td colspan="9" align="center" style="font-size: 15px;">
        <strong>TOTAL: R$ <%=arrumaNumero(totalBoleto,"0,00",2)%></strong>
    </td>
</tr>
					<%


					%> </table></td></tr> <%
				end if
				rsBoleto.close
				set rsBoleto = Nothing
				%>

              </table>







	  	      <p align="center">&nbsp;</p>
	  	     <%
				AbrirConexao()	
						
							'Dinheiro / Cart�o
							PAGAMENTO_EM_DINHEIRO = 0
							XTOTAL_CARTAO = 0
							XTOTAL_CHEQUE = 0
							
							Set rsVALOR_TOTAL = objConexao.Execute("SELECT * FROM VW_PEDIDOS_TOTAL_RECEBIDO WHERE COD_PEDIDO = "&COD_PEDIDO&" LIMIT 1")
							
								If not rsVALOR_TOTAL.Eof Then
									
									PAGAMENTO_EM_DINHEIRO = formatNumber(0&rsVALOR_TOTAL("PAGAMENTO_EM_DINHEIRO"))							
									XTOTAL_CARTAO = formatNumber(0&rsVALOR_TOTAL("XTOTAL_CARTAO"))
									XTOTAL_CHEQUE = formatNumber(0&rsVALOR_TOTAL("XTOTAL_CHEQUE"))
									XTOTAL_BOLETO = formatNumber(0&rsVALOR_TOTAL("XTOTAL_BOLETO"))
								End IF
								
							rsVALOR_TOTAL.Close()
							Set rsVALOR_TOTAL = Nothing

						FecharConexao()

						VALOR_TOTAL = (PAGAMENTO_EM_DINHEIRO/1 + XTOTAL_CHEQUE/1 +  XTOTAL_CARTAO/1 + XTOTAL_BOLETO/1)

						'//Fim total arrecadado
						If int(PEDIDO_EXTRA) = 2 Then%>
                   	    <table width="98%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr>
    <td width="463"><div align="left"><strong class="textoMaior">Total em Produtos:</strong></div></td>
    <td width="836"><div align="left"><span class="textoMaior"><strong>
    <%
						Dim rsProdutosTotal, QTD_EM_PRODUTOS, XTOTAL_AGREGADOS, XTOTAL_ALBUNS, XTOTAL_PRODUTOS_A_PARTE
						AbrirConexao()
						
						QTD_EM_PRODUTOS = 0
						XTOTAL_AGREGADOS = 0
						XTOTAL_ALBUNS = 0
						XTOTAL_PRODUTOS_A_PARTE = 0
						
							'VW_PEDIDOS_PRODUTOS_TOTAL_AGREGADOS
							Set rsProdutosTotal = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_AGREGADOS WHERE COD_PEDIDO="&COD_PEDIDO)
							
								if not rsProdutosTotal.Eof Then
									
									XTOTAL_AGREGADOS = formatNumber(0&rsProdutosTotal("XTOTAL_AGREGADOS"))/1
								
								End If

							rsProdutosTotal.Close()
							Set rsProdutosTotal = Nothing
							
							'VW_PEDIDOS_PRODUTOS_TOTAL_ALBUNS
							Set rsProdutosTotal = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_ALBUNS WHERE COD_PEDIDO="&COD_PEDIDO)
							
								if not rsProdutosTotal.Eof Then
									
									XTOTAL_ALBUNS = formatNumber(0&rsProdutosTotal("XTOTAL_ALBUNS"))/1
								
								End If
							
							rsProdutosTotal.Close()
							Set rsProdutosTotal = Nothing
							
							'VW_PEDIDOS_PRODUTOS_TOTAL_PRODUTOS_A_PARTE
							Set rsProdutosTotal = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_PRODUTOS_A_PARTE WHERE COD_PEDIDO="&COD_PEDIDO)
							
								if not rsProdutosTotal.Eof Then
									
									XTOTAL_PRODUTOS_A_PARTE = formatNumber(0&rsProdutosTotal("XTOTAL_PRODUTOS_A_PARTE"))/1
								
								End If
							
							
							rsProdutosTotal.Close()
							Set rsProdutosTotal = Nothing


							set rsProdutosTotal = objConexao.Execute("select SUM(valor_total_com_desconto) as valor_total_com_desconto from(SELECT CASE WHEN "&PEDIDO_EXTRA&" = 2 THEN SUM(faturado * (valor_un-(valor_un * (1-(ifnull(aliquota_desconto_foto,0)/100))))) ELSE '0' END as valor_total_com_desconto FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO in (SELECT cod_formando FROM TB_PEDIDOS_FORMANDOS WHERE cod_pedido =  "&cod_pedido&") union all SELECT SUM(quantidade * (valor_un-(valor_un * (1-(ifnull(aliquota,0)/100))))) as valor_total_com_desconto  FROM TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO = "&cod_pedido&") as total_desconto")
							if not rsProdutosTotal.Eof then
								XTOTAL_DESCONTO = formatCurrency(0&rsProdutosTotal(0))/1
							end if
							rsProdutosTotal.Close
							set rsProdutosTotal = Nothing
												
							QTD_EM_PRODUTOS = formatCurrency(XTOTAL_AGREGADOS + XTOTAL_ALBUNS + XTOTAL_PRODUTOS_A_PARTE - XTOTAL_DESCONTO)
						%>
      <%=QTD_EM_PRODUTOS%></strong></span></div></td>
    </tr>
  <tr>
    <td class="textoMaior"><div align="left"><strong>Total em Forma de Pagamento:</strong></div></td>
    <td class="textoMaior"><div align="left"><strong><%=formatCurrency(VALOR_TOTAL)%></strong></div></td>
    </tr>
  <%
  	'Cr�dito de formando, quando h� algum saldo residual por algum motivo vindo da formatura.
  	Dim TOTAL_CREDITO
  	sql =	" SELECT SUM(IFNULL(TB_FICHAS_IDENTIFICACAO.CREDITO_EM_VALOR,0)) AS TOTAL_CREDITO "&_
			" FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON "&_
			" TB_PEDIDOS_FORMANDOS.COD_FORMANDO =  "&_
			" TB_FICHAS_IDENTIFICACAO.COD_FORMANDO WHERE  "&_
			" TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&COD_PEDIDO&" GROUP BY TB_PEDIDOS_FORMANDOS.COD_PEDIDO"
	Set rsCredito = objConexao.Execute(sql) 
					
					If not rsCredito.Eof Then
						TOTAL_CREDITO = formatNumber(0&rsCredito("TOTAL_CREDITO"))
					End If
					
	rsCredito.Close
	Set rsCredito = Nothing
  
  If TOTAL_CREDITO > 0 Then %>
  <tr>
    <td class="textoMaior"><div align="left"><strong>Total Cr&eacute;dito da formatura:</strong></div></td>
    <td class="textoMaior"><div align="left"><strong><%=formatCurrency(TOTAL_CREDITO)%></strong></div></td>
    </tr>
  <tr bgcolor="#E2E2E2">
    <td bgcolor="#E2E2E2" class="textoMaior"><div align="left"><strong>Total em Pagamento + Cr&eacute;dito:</strong></div></td>
    <td class="textoMaior"><div align="left"><strong><%=formatCurrency(VALOR_TOTAL+TOTAL_CREDITO)%></strong></div></td>
    </tr>
  <%
  End If
  
FecharConexao()
  %>
</table>
					<%Else%>
                    	 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr>
    <td width="467" class="textoMaior"><div align="left"><strong>Total em Produtos:</strong></div></td>
    <td width="838" class="textoMaior">
      
        <div align="left">
          <%
						AbrirConexao()
						
						QTD_EM_PRODUTOS = 0
						
							Set rsProdutosTotal = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_PRODUTOS WHERE COD_PEDIDO="&COD_PEDIDO)
							
								if not rsProdutosTotal.Eof Then
									

									XTOTAL_PRODUTOS_A_PARTE = formatNumber(0&rsProdutosTotal("XTOTAL_PRODUTOS_A_PARTE"))/1
									
									QTD_EM_PRODUTOS = formatCurrency(XTOTAL_PRODUTOS_A_PARTE)
									
								Else
									
									QTD_EM_PRODUTOS = 0
									
								End If
							
								
						
							
							rsProdutosTotal.Close()
							Set rsProdutosTotal = Nothing

							set rsProdutosTotal = objConexao.Execute("select SUM(valor_total_com_desconto) as valor_total_com_desconto from(SELECT CASE WHEN "&PEDIDO_EXTRA&" = 2 THEN SUM(faturado * (valor_un-(valor_un * (1-(ifnull(aliquota_desconto_foto,0)/100))))) ELSE '0' END as valor_total_com_desconto FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO in (SELECT cod_formando FROM TB_PEDIDOS_FORMANDOS WHERE cod_pedido =  "&cod_pedido&") union all SELECT SUM(quantidade * (valor_un-(valor_un * (1-(ifnull(aliquota,0)/100))))) as valor_total_com_desconto  FROM TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO = "&cod_pedido&") as total_desconto")
							if not rsProdutosTotal.Eof then
								XTOTAL_DESCONTO = formatCurrency(0&rsProdutosTotal(0))/1
							end if
							rsProdutosTotal.Close
							set rsProdutosTotal = Nothing

							QTD_EM_PRODUTOS =  formatCurrency(QTD_EM_PRODUTOS - XTOTAL_DESCONTO)
						%>
          <strong><%=Trim(QTD_EM_PRODUTOS)%></strong></div></td>
    </tr>
  <tr>
    <td class="textoMaior"><div align="left"><strong>Total em Forma de Pagamento:</strong></div></td>
    <td class="textoMaior"><div align="left"><strong><%=formatCurrency(VALOR_TOTAL)%></strong></div></td>
    </tr>
 <%
  
FecharConexao()
  %>
</table>
<%End If%>
<br />
<hr />
<h2 class="order-section-heading" id="ordens_servico">Ordens de servi&ccedil;o da venda</h2>
       <%
	   AbrirConexao()
	   Dim pode_imprimir_2_via_despacho_procolo
	   pode_imprimir_2_via_despacho_procolo = true
	   
	   Set rs2=objConexao.Execute("SELECT * FROM VW_TB_ORDENS_SERVICO WHERE cod_pedido = "&COD_PEDIDO&" order by numero_ordem_servico desc, numero_ordem_servico_indice asc")
			if not rs2.Eof then

			If validaUsuario(2,,55,true) Then
			%>
			<script type="text/javascript">
			function cancelarOs()
			{
				var cods_os="";

				$("input[name=cod_ordem_servico_acao]:checked").each(function(){
					cods_os = cods_os + $(this).val()+",";
				});
				cods_os = cods_os.substr(0, cods_os.length-1);

				$("input[name=lote_os]").val(cods_os);

				if (cods_os.length > 0)
				{
					$("#cancelar_os").show();
					$("#sub-menu-os").hide();
				}
				else
				{
					$("#sub-menu-os").show();
					$("#cancelar_os").hide();
					alert("Selecione uma OS");
				}
			}
			function cancelarAcao()
			{	
				$("#cancelar_os").hide();
				$("textarea[name=MOTIVO_CANCELADA]").val('');
				$("#sub-menu-os").show();
			}
			</script>
			<div style="width:99%; float:left; position:relative;" id='sub-menu-os'>
				<a href="javascript:cancelarOs()" class="btnGeral" style="float:left; position:relative;">Cancelar Ordens de Servi&ccedil;o</a>
			</div>
			<form name="FORMULARIO_PADRAO">

			<div id='cancelar_os' style="width:430px; background-color:#ffffcc; padding:25px; margin:0px auto 10px; display:none; position:realtive; float:left;">
				<h1>Cancelar Ordens Servi&ccedil;o</h1>
				<textarea name="MOTIVO_CANCELADA" maxlength='255' class="inputTamanho3" rows="5"></textarea>
				<input type='hidden' name='lote_os' id='lote_os' value='' />
				<table>
				<tr>
					<td><a href="javascript:OSaction('cancelarVariasOs',0,'', $('#lote_os').val() )" class="btnGeral" style="float:left; position:relative;">Confirmar</a></td>
					<td><a href="javascript:cancelarOs()" class="btnGeral" style="float:left; position:relative;">Cancelar</a></td>
				</tr>
				</table>
			</div>


			</form>
			<%end if 'if validacao do usuario %>
		<div class="order-scroll" role="region" aria-label="Ordens de servi&ccedil;o da venda" tabindex="0">
<table width="99%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
  <tr bgcolor="whitesmoke">
  <td rowspan="2" width="3%"><div align="center"><strong> OS</strong></div></td>
  <td rowspan="2" width="12%"><div align="center"><strong>Formando</strong></div></td>
  <td width="10%" rowspan="2"><div align="center"><strong>Produto</strong></div></td>
  <td width="6%" rowspan="2"><div align="center"><strong>Qtd.</strong></div></td>
  <td colspan="5"><div align="center" style="font-weight: bold">
    <div align="center">Etapa Atual Informa&ccedil;&otilde;es</div>
  </div></td>
  <td width="9%" rowspan="2"><div align="center"><strong>Data Cadastro</strong></div></td>
  <td width="6%" rowspan="2"><div align="center"><strong>Prazo Geral</strong></div></td>
  <td width="5%" rowspan="2"><div align="center"><strong>Status</strong></div>  </td>
  <td width="3%" rowspan="2">&nbsp;</td>
  <td width="3%" rowspan="2"><p align="center"> </p>
    <p align="center"> </p></td>
    <td width="8%" rowspan="2" ><div align="center"><strong>&nbsp;</strong></div></td>
</tr>
  <tr bgcolor="whitesmoke">
    <td width="10%"><div align="center"><strong>Etapa / Total</strong></div></td>
    <td width="10%"><div align="center"><strong>Descri&ccedil;&atilde;o</strong></div></td>
    <td width="8%"><div align="center"><strong>Resp. Atual</strong></div></td>
    <td width="7%"><div align="center"><strong>Data In&iacute;cio</strong></div></td>
    <td width="8%"><div align="center"><strong>Prazo etapa</strong></div></td>
    
    </tr>                         
		<%
		While not rs2.Eof

		Set rs3=objConexao.Execute("Select count(*) as qtd_etapas From TB_ORDENS_SERVICO_ETAPAS where cod_ordem_servico = "&rs2("cod_ordem_servico")&" group by cod_ordem_servico")
		If not rs3.Eof Then
			qtd_etapas = formatNumber(0&rs3("qtd_etapas"),0)
		End If
		rs3.Close
		Set rs3=Nothing
		%>                   
            <tr>
                                        <td><div align="center"><strong><%=moduloPPVOSMascara(rs2("cod_ordem_servico"),"",true)%></strong></div></td>
                                        <td><%=rs2("nome_formando")%></td>
                            <td><div align="left">
                                <%=rs2("nome_produto")%>                                        
                                </div></td>
                            <td><div align="center"><%=rs2("quantidade")%> </div></td>
                                        <td><div align="center"><%=rs2("numero_etapa")%> / <%=qtd_etapas%></div></td>
                                        <td><div align="left"><%=rs2("descricao")%></div></td>
                                        <td><div align="left"><%=nomeUsuario(rs2("cod_usuario_responsavel"),true)%></div></td>
                                        <td><div align="center"><%=ArrumaData(rs2("data_inicio"))%></div></td>
                                        <td><div align="center"><%=ArrumaData(rs2("prazo_limite"))%></div></td>
                                        <td><div align="center"><%=ArrumaData(rs2("data_cadastrado"))%></div></td>
                                        <td><div align="center">
										<%=ArrumaData(DateAdd("d",rs2("prazo_cliente_total"),rs2("data_base_prazo_cliente_inicio")))%></div></td>
                                        <td><div align="center"><%=rs2("status_ordem_servico")%></div></td>
                                        <td><div align="center"><a href="ppv_ordens_servico_formulario.asp?comando=visualizar&cod_ordem_servico=<%=rs2("cod_ordem_servico")%>"><img src="images/icons/icon_visualizar.png" width="16" height="16" border="0" /></a></div></td>
                                        <td>
                                        <div align="center">
											<%If int(rs2("status_ordem_servico")) = 4 Then
											pode_imprimir_2_via_despacho_procolo = true
											%>
                                                <img src="images/icons/icon_green.gif" width="16" height="15" />
                                                <%ElseIf int(rs2("status_ordem_servico")) = 3 Then
											pode_imprimir_2_via_despacho_procolo = true
											%>
                                                <img src="images/icons/icon_entrega.png" width="16" height="16" />
                                            <%ElseIf int(rs2("status_ordem_servico")) = 2 Then
												pode_imprimir_2_via_despacho_procolo = false
											%> 
                                            	 <img src="images/icons/icon_aguardando.png" width='16px' height='16px' />
                                            <%Elseif  int(rs2("status_ordem_servico")) = 99  Then
												pode_imprimir_2_via_despacho_procolo = false
											%>
                                                <img src="images/icons/icon_red.gif" width="17" height="15" />
                                            <%Else%>
                                       	  <%
												pode_imprimir_2_via_despacho_procolo = false
												'Response.Write(rs2("diferenca_etapa"))
												diferenca = formatNumber(rs2("diferenca_etapa"))/1
												'Response.Write(diferenca)
												If diferenca > 0 Then%>
                                       	  <img src="images/icons/icon_balao_red.gif" alt="Etapa atrasada há <%=diferenca%> dia(s)." title="Etapa atrasada há <%=diferenca%> dia(s)." />
                                       	  <%Elseif diferenca = 0 Then%>
                                       	  <img src="images/icons/icon_balao_yellon.gif" alt="OS vencendo hoje." title="OS vencendo hoje" />
                                                <%Else%>
                                       	  <img src="images/icons/icon_balao_green.gif" alt="Faltam <%=(diferenca)*(-1)%> dia(s) para término do prazo da etapa." title="Faltam <%=(diferenca)*(-1)%> dia(s) para t�rmino do prazo da etapa." />                                                
                                                <%End If%>
                                            <%End If%>
                                        </div>                                        
									</td>
									<td>
									<% if int(rs2("status_ordem_servico")) < 4 then%>
										<input type="checkbox" name="cod_ordem_servico_acao" value="<%=rs2("cod_ordem_servico")%>" class='inputradio'>
									<%end if%>
									</td>
                            </tr>
                
       <%
	   rs2.MoveNext()
	   Wend
	   %>
        </table></div>
        <%If pode_imprimir_2_via_despacho_procolo Then%>
        
      <div style="position:relative; float:left; width: 100%; padding-top:30px;">
		<a href="ppv_ordens_servico_triagem_despacho_imprimir.asp?c=<%=cod_pedido%>" target="_blank" class="btnGeral" style="display:block;width:440px;">IMPRIMIR 2&ordm; VIA DE PROTOCOLO DESPACHO ( Todos os Itens )</a>
      <a href="ppv_ordens_servico_triagem_despacho_imprimir_entrega.asp?c=<%=cod_pedido%>" target="_blank" class="btnGeral" style="display:none;margin-top:5px;width:440px;">IMPRIMIR 2&ordm; VIA DE PROTOCOLO DESPACHO ( Somente produtos n&atilde;o entregues )</a>
      <span onclick="showListaItens()" class="btnGeral" style="display:block;margin-top:5px;width:440px;cursor:pointer;color:#1B5DE4;text-decoration: underline;">IMPRIMIR 2&ordm; VIA DE PROTOCOLO DESPACHO ( Selecionar Itens )</span> 
      <div id="ListaItens" style="display:none;">
      
    <form name="IMPRIMIRP_PROTOCOLO_CLIENTE" action="ppv_ordens_servico_triagem_despacho_imprimir_protocolo_selecionado.asp" method="post" style="margin:0px;margin-left:auto;">
       <table width="800px" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">     
        <tr bgcolor="whitesmoke">
		     <td width="70%"><div align="center"><strong>Produto</strong></div></td>

		     <td width="30%"><div align="center"><strong>Selecionar</strong></div></td>
        </tr>
      <%

		 Dim numero_etapa, descricao, cod_usuario_responsavel, nome_usuario_responsavel, prazo_em_dias, data_inicio, prazo_limite, diferenca, qtd_etapas
			
			Set rs2=objConexao.Execute("SELECT TB_ORDENS_SERVICO.* FROM TB_ORDENS_SERVICO WHERE cod_pedido = "&cod_pedido&" and data_finalizado IS NULL and status_ordem_servico NOT IN (0,1,2,99)")
				If not rs2.Eof Then
					permitido = false
				Else
					permitido = true							
				End If
			rs2.Close
			Set rs2=Nothing
          
			
           sql = "SELECT VW_TB_ORDENS_SERVICO.*,cast(TB_PEDIDOS_ITENS_ENTREGUES.data as date) AS data_entrega FROM VW_TB_ORDENS_SERVICO LEFT JOIN TB_PEDIDOS_ITENS_ENTREGUES ON TB_PEDIDOS_ITENS_ENTREGUES.cod_ordem_servico = VW_TB_ORDENS_SERVICO.cod_ordem_servico WHERE VW_TB_ORDENS_SERVICO.cod_pedido = "&cod_pedido&"  and VW_TB_ORDENS_SERVICO.status_ordem_servico NOT IN (0,1,2,99)  order by VW_TB_ORDENS_SERVICO.numero_ordem_servico desc, numero_ordem_servico_indice asc"
           Set rs2=objConexao.Execute(sql)
           
           While not rs2.Eof
           
                    
                    Set rs3=objConexao.Execute("Select count(*) as qtd_etapas From TB_ORDENS_SERVICO_ETAPAS where cod_ordem_servico = "&rs2("cod_ordem_servico")&" group by cod_ordem_servico")
                        
                        If not rs3.Eof Then
                            qtd_etapas = formatNumber(0&rs3("qtd_etapas"),0)
                        End If
                        
                    rs3.Close
                    Set rs3=Nothing
					
					data_finalizado = rs2("data_finalizado")
           
           %>                   
                <tr <%If not isDate(data_finalizado) Then%> bgcolor="#FFAAAA" <%End If%>>
                    <td><div align="left">
                    <%=rs2("nome_produto")%>                                        
                    </div></td>
                     <td >
                    	<div align="center">
                    		<% if rs2("data_entrega") <> "" then %>
								<input type="checkbox" value="<%=rs2("cod_ordem_servico")%>" name="produtos" class="inputRadio" />
							<% else	%>	
								 <input type="checkbox" value="<%=rs2("cod_ordem_servico")%>" name="produtos" class="inputRadio" />
							<% End If %>
           				</div>
       	 			</td>
                </tr>
           <%
           rs2.MoveNext()
           Wend
           rs2.Close
           Set rs2=Nothing
           %>
           
          <tr>
                  <td colspan="7"><div align="center"><input type="submit" style="width:500px;cursor:pointer" value="IMPRIMIR PROTOCOLO DESPACHO DOS ITENS SELECIONADOS"/></div></td>
           </tr>
         
          </table>
         </form>
      </div>
      
      </div>
        </div>
        <%End If%>
        
        		<%
			Else
				%>
<div style="padding:15px; text-align:center"><strong>N&atilde;o h&aacute; Ordens de Servi&ccedil;o para esta Venda</strong></div>
                <%
				
			end if

	   %>
       
       <%
	   FecharConexao()
	   %>
      </div>
<script type="text/javascript">

function historicoCheque(COD_PEDIDO,COD_PEDIDO_CHEQUE){
var caminho
caminho = "<table cellpadding=\'0\' cellspacing=\'0\' border=\'0\' width=\'100%\'><tr><td><table cellpadding=\'0\' cellspacing=\'0\' width=\'100%\'><tr><td width=\'70%\'><h1>Cheque pr&eacute;-datado</h1></td><td align=\'right\'><a href=\'#\' onclick=\'closeMessage();return false\'>Fechar Janela</a></td></tr></table></td></tr><tr><td><iframe frameborder=\'0\' scrolling=\'yes\' height=\'380px\' width=\'99%\' src=\'lca_pedidos_pendentes_alteracoes_part_3_POPUP_cheque_historico.asp?comando=visualizar&COD_PEDIDO="+COD_PEDIDO+"&cod_pedido_cheque="+COD_PEDIDO_CHEQUE+"\'></iframe></td></tr></table>"
displayStaticMessage(caminho,false,420)		
}

</script>
<!-- #include file="files/includes/scripts_rodape.asp" -->
    </main>
    <footer class="dashboard-footer">&copy; <%=Year(Date)%> Grupo Stillos <span>Sistema administrativo interno</span></footer>
  </div>
</div>
<script>
(function () {
    var toggle = document.getElementById('ppvMenuToggle');
    var menu = document.getElementById('menuPrincipal');
    var previous = menu.querySelector('[aria-current="page"]');
    var current = menu.querySelector('a[href="ppv_ordens_servico.asp"]');
    var back = document.getElementById('orderBack');
    back.addEventListener('click', function (event) {
        if (document.referrer && new URL(document.referrer).origin === window.location.origin) {
            event.preventDefault();
            window.history.back();
        }
    });
    if (previous) previous.removeAttribute('aria-current');
    if (current) current.setAttribute('aria-current', 'page');
    toggle.addEventListener('click', function () {
        var expanded = menu.classList.toggle('is-open');
        toggle.setAttribute('aria-expanded', String(expanded));
    });
    document.addEventListener('keydown', function (event) {
        if (event.key === 'Escape' && menu.classList.contains('is-open')) {
            menu.classList.remove('is-open');
            toggle.setAttribute('aria-expanded', 'false');
            toggle.focus();
        }
    });
})();
</script>
</body>
</html>
