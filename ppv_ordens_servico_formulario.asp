<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 2
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
  Dim titulo_pagina
  titulo_pagina = "Ordens SERVIÇO | Formulário"

'data in�cio formatado, data t�rmino, t�tulo da p�gina
'call manutencao(ArrumaData(date(),"")& "�s 08h00",ArrumaData(date(),"")& "�s 12h00","")
'//////////////////////////
Dim validacao_listar_os_todos_usuarios, complemento_validacao_usuario
Dim flag_controle
flag_controle = Cstr(""&Request.QueryString("abe"))
'Com Está flag � poss�vel visualizar todas as OS flag_controle
validacao_listar_os_todos_usuarios = validaUsuario(2,,57,false)
complemento_validacao_usuario = ""
If not validacao_listar_os_todos_usuarios and flag_controle <> "S" Then
	complemento_validacao_usuario = " AND COD_USUARIO_RESPONSAVEL = "&Session("cod_usuario")
End If
'//////////////////////////
Dim calendario_necessario : calendario_necessario = "NÃO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<%

'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 26/6/2012 10:10:55

'Declara vari�veis para utiliza��o do sistema
Dim cod_ordem_servico, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, actionformulario

redirecionar_incluir	= "ppv_ordens_servico.asp"
redirecionar_alterar	= "ppv_ordens_servico.asp"
redirecionar_excluir	= "ppv_ordens_servico.asp"
redirecionar_validacao	= "ppv_ordens_servico.asp"
msgAoExcluir			= "Tem certeza que deseja excluir esta Ordem de SERVIÇO?"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
'Declara vari�veis
Dim cod_pedido, cod_produto, cod_formando, numero_ordem_servico, numero_ordem_servico_indice, quantidade, cod_ordem_servico_etapa_atual, cod_usuario_processou, cod_usuario_finalizou, cod_usuario_despachou, cod_usuario_cancelou, prazo_cliente_total, data_base_prazo_cliente_inicio, data_finalizado, data_despachado, data_cancelado, motivo_gerado_manualmente, status_ordem_servico, data_cadastrado, nome_formando, motivo_cancelada

'P�gina para retornar
Dim pagina_retornar

'Secund�rias Etaps
Dim numero_etapa_atual, cod_usuario_responsavel_atual, data_inicio_atual, data_termino_atual, pedido_compra_gerado, necessita_gerar_pedido_compra, descricao, qtd_etapas

Dim insercao_manual
insercao_manual = false 'Manual somente se o comando for incluir

If comando = Empty Then
comando = "incluir"
insercao_manual = true
titulo_pagina = titulo_pagina&" / Inclusão"
End if

If comando = "incluir" and campo1 <> Empty Then
	'Incluir novo registro
	incluirRegistro(redirecionar_incluir)
	titulo_pagina = titulo_pagina&" / Inclusão"
Elseif comando = "alterar" and campo1 = Empty Then
	'Leitura inicial para poss�vel Alteração
	cod_ordem_servico = Request.QueryString("cod_ordem_servico")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "alterar" and campo1 <> Empty Then
	'Gravar Alteração
	cod_ordem_servico = Request.Form("cod_ordem_servico")
	alterarRegistro(redirecionar_alterar)
	titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 
	'Leitura inicial para poss�vel exclus�o
	cod_ordem_servico = Request.QueryString("cod_ordem_servico")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
	'Exclui registro
	cod_ordem_servico = Request.Form("cod_ordem_servico")
	excluirRegistro(redirecionar_excluir)
	titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
	'Visualiza registro
	cod_ordem_servico = Request.QueryString("cod_ordem_servico")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / Visualização"

Elseif comando = "trocarRespOS" or comando = "finalizarOS" or comando = "iniciarEtapa" or comando = "cancelarOS" or comando = "finalizarEtapa" or comando = "voltarEtapa" Then
	
	Dim parametro, parametro2, etapa
	parametro = Request.QueryString("parametro") 'Parametros diversos
	parametro2 = Ucase(Cstr(""&Request.QueryString("parametro2"))) 'Parametros diversos
	etapa = Request.QueryString("etapa") 'cod_ordem_servico_etapa
	cod_ordem_servico = Request.QueryString("cod_ordem_servico")
	pagina_retornar = Request.QueryString("pagina_retornar")
	pagina_retornar = Replace(pagina_retornar,"{","")
	pagina_retornar = Replace(pagina_retornar,"}","")
	pagina_retornar = Replace(pagina_retornar,"|","&")
	
	response.end()

End If

'Fun��es

Sub lerFormulario()

cod_ordem_servico = Request.Form("cod_ordem_servico")
cod_pedido = Request.Form("cod_pedido")
cod_produto = Request.Form("cod_produto")
cod_formando = Request.Form("cod_formando")
numero_ordem_servico = Request.Form("numero_ordem_servico")
numero_ordem_servico_indice = Request.Form("numero_ordem_servico_indice")
quantidade = Request.Form("quantidade")
cod_usuario_processou = Request.Form("cod_usuario_processou")
cod_usuario_finalizou = Request.Form("cod_usuario_finalizou")
cod_usuario_despachou = Request.Form("cod_usuario_despachou")
cod_usuario_cancelou = Request.Form("cod_usuario_cancelou")
prazo_cliente_total = Request.Form("prazo_cliente_total")
data_base_prazo_cliente_inicio = Request.Form("data_base_prazo_cliente_inicio")
data_finalizado = Request.Form("data_finalizado")
data_despachado = Request.Form("data_despachado")
data_cancelado = Request.Form("data_cancelado")
motivo_gerado_manualmente = Request.Form("motivo_gerado_manualmente")
status_ordem_servico = Request.Form("status_ordem_servico")
pagina_retornar = Request.Form("pagina_retornar")

End sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM VW_TB_ORDENS_SERVICO WHERE cod_ordem_servico = "&cod_ordem_servico&complemento_validacao_usuario)

				cod_pedido=trim(rs("cod_pedido"))
				cod_produto=trim(rs("cod_produto"))
				cod_formando=trim(rs("cod_formando"))
				numero_ordem_servico=trim(rs("numero_ordem_servico"))
				numero_ordem_servico_indice=trim(rs("numero_ordem_servico_indice"))
				quantidade=trim(rs("quantidade"))			
				cod_usuario_processou=trim(rs("cod_usuario_processou"))
				cod_usuario_finalizou=trim(rs("cod_usuario_finalizou"))
				cod_usuario_despachou=trim(rs("cod_usuario_despachou"))
				cod_usuario_cancelou=trim(rs("cod_usuario_cancelou"))
				prazo_cliente_total=trim(rs("prazo_cliente_total"))
				data_base_prazo_cliente_inicio=trim(rs("data_base_prazo_cliente_inicio"))
				data_finalizado=trim(rs("data_finalizado"))
				data_despachado=trim(rs("data_despachado"))
				data_cancelado=trim(rs("data_cancelado"))
				motivo_gerado_manualmente=trim(rs("motivo_gerado_manualmente"))
				status_ordem_servico=trim(rs("status_ordem_servico"))
				data_cadastrado=rs("data_cadastrado")
				nome_formando=rs("nome_formando")
				motivo_cancelada=rs("motivo_cancelada")
				
				'Informa��es da etapa atual
				'/////
				cod_ordem_servico_etapa_atual=trim(rs("cod_ordem_servico_etapa_atual"))
				numero_etapa_atual=formatNumber(0&rs("numero_etapa"),0)
				cod_usuario_responsavel_atual=trim(rs("cod_usuario_responsavel"))
				data_inicio_atual=rs("data_inicio")
				data_termino_atual=rs("data_termino")
				pedido_compra_gerado=Cstr(""&rs("pedido_compra_gerado"))
				necessita_gerar_pedido_compra=Cstr(""&rs("necessita_gerar_pedido_compra"))
				descricao=Lcase(Cstr(""&rs("descricao")))
				'/////
				
				'Conta a quantidade de etapas
				Set rs2=objConexao.Execute("Select count(*) as qtd_etapas From TB_ORDENS_SERVICO_ETAPAS where cod_ordem_servico = "&cod_ordem_servico&" group by cod_ordem_servico")
					
					If not rs2.Eof Then
						qtd_etapas = formatNumber(0&rs2("qtd_etapas"),0)
					End If
					
				rs2.Close
				Set rs2=Nothing
				'/////

		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("INSERT INTO TB_ORDENS_SERVICO (cod_pedido, cod_produto, cod_formando, numero_ordem_servico, numero_ordem_servico_indice, quantidade, cod_ordem_servico_etapa_atual, cod_usuario_processou, cod_usuario_finalizou, cod_usuario_despachou, cod_usuario_cancelou, prazo_cliente_total, data_base_prazo_cliente_inicio, data_finalizado, data_despachado, data_cancelado, motivo_gerado_manualmente, status_ordem_servico) VALUES ("&sqlNumero(cod_pedido)&", "&sqlTexto(cod_produto)&", "&sqlNumero(cod_formando)&", "&sqlNumero(numero_ordem_servico)&", "&sqlTexto(numero_ordem_servico_indice)&", "&sqlTexto(quantidade)&", "&sqlTexto(cod_ordem_servico_etapa_atual)&", "&sqlTexto(cod_usuario_processou)&", "&sqlTexto(cod_usuario_finalizou)&", "&sqlTexto(cod_usuario_despachou)&", "&sqlTexto(cod_usuario_cancelou)&", "&sqlTexto(prazo_cliente_total)&", "&sqlData(data_base_prazo_cliente_inicio)&", "&sqlData(data_finalizado)&", "&sqlData(data_despachado)&", "&sqlData(data_cancelado)&", "&sqlTexto(motivo_gerado_manualmente)&", "&sqlTexto(status_ordem_servico)&")")

		'Encontrar chave criada
		Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_ORDENS_SERVICO LIMIT 1")
			If not rs.Eof Then
				cod_ordem_servico = rs("ULTIMO_ID")
			End if 
		rs.Close
		Set rs=Nothing


	FecharConexao()
	Response.Redirect(pagina_retornar&"?cod_ordem_servico="&cod_ordem_servico&"&codmsg=1")
End sub

Sub alterarRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("UPDATE TB_ORDENS_SERVICO SET cod_pedido = "&sqlNumero(cod_pedido)&", cod_produto = "&sqlTexto(cod_produto)&", cod_formando = "&sqlNumero(cod_formando)&", numero_ordem_servico = "&sqlNumero(numero_ordem_servico)&", numero_ordem_servico_indice = "&sqlTexto(numero_ordem_servico_indice)&", quantidade = "&sqlTexto(quantidade)&", cod_ordem_servico_etapa_atual = "&sqlTexto(cod_ordem_servico_etapa_atual)&", cod_usuario_processou = "&sqlTexto(cod_usuario_processou)&", cod_usuario_finalizou = "&sqlTexto(cod_usuario_finalizou)&", cod_usuario_despachou = "&sqlTexto(cod_usuario_despachou)&", cod_usuario_cancelou = "&sqlTexto(cod_usuario_cancelou)&", prazo_cliente_total = "&sqlTexto(prazo_cliente_total)&", data_base_prazo_cliente_inicio = "&sqlData(data_base_prazo_cliente_inicio)&", data_finalizado = "&sqlData(data_finalizado)&", data_despachado = "&sqlData(data_despachado)&", data_cancelado = "&sqlData(data_cancelado)&", motivo_gerado_manualmente = "&sqlTexto(motivo_gerado_manualmente)&", status_ordem_servico = "&sqlTexto(status_ordem_servico)&" WHERE cod_ordem_servico = "&cod_ordem_servico)
	FecharConexao()
	Response.Redirect(pagina_retornar&"?cod_ordem_servico="&cod_ordem_servico&"&codmsg=2")
End sub

Sub excluirRegistro(redirecionar)
	'lerFormulario()
		'call verificaAntesExcluir(,cod_ordem_servico,"TABELASRELACIONADAS;TABELASRELACIONADAS","cod_ordem_servico",redirecionar_excluir&"?sucesso=2&codmsg=5")

	'AbrirConexao()
		'objConexao.Execute("DELETE FROM TB_ORDENS_SERVICO WHERE cod_ordem_servico = "&cod_ordem_servico)
	'FecharConexao()
	Response.Redirect(pagina_retornar&"?cod_ordem_servico="&cod_ordem_servico&"&codmsg=3")
End sub


'----------------------------------------------------------------------------------------
%>
<style>
.os-detail { box-sizing:border-box; width:100%; max-width:1440px; margin:0 auto; padding:clamp(12px,2.5vw,32px); color:#292738; background:#f6f5f9; border-radius:12px; font:14px/1.5 system-ui,sans-serif; }
.os-detail *, .os-detail *::before, .os-detail *::after { box-sizing:border-box; }
.os-detail .os-heading { display:flex; align-items:center; justify-content:space-between; gap:16px; flex-wrap:wrap; margin-bottom:20px; }
.os-detail .os-context { color:#9e125c; font-size:12px; font-weight:750; text-transform:uppercase; letter-spacing:.06em; }
.os-detail .os-heading h2 { color:#26243a; font:700 clamp(22px,3vw,32px)/1.2 system-ui,sans-serif; margin:5px 0 0; }
.os-detail .os-back { display:inline-block; padding:8px 12px; border:1px solid #ded9e6; background:#fff; border-radius:8px; color:#6d1749; font-weight:650; text-decoration:none; }
.os-detail .subMenu { margin-bottom:18px; padding:10px 12px; border:1px solid #e4dfeb; border-radius:10px; background:#fff; white-space:normal; line-height:2.3; }
.os-detail .subMenu a { display:inline-block; max-width:100%; overflow-wrap:anywhere; }
.os-detail fieldset.os-panel { min-width:0; margin:0 0 18px; padding:14px 16px 16px; border:1px solid #e4dfeb; border-radius:12px; background:#fff; }
.os-detail .os-panel legend { float:none; width:auto; margin:0; padding:0 8px; color:#292738; font:700 16px/1.5 system-ui,sans-serif; border:0; }
.os-detail .os-table { width:100%; margin:0; border:0; border-collapse:collapse; color:inherit; }
.os-detail .os-table > tbody > tr > td { padding:10px 12px; border:0; border-bottom:1px solid #eeeaf2; vertical-align:middle; overflow-wrap:anywhere; }
.os-detail .os-summary > tbody > tr > td:first-child:not([colspan]) { background:#f8f6fa; color:#494258; }
.os-detail .os-summary > tbody > tr > td[colspan] { border-bottom:0; padding:12px; }
.os-detail .os-stages > tbody > tr:first-child > td { background:#f1edf6; color:#443b56; }
.os-detail .os-stage-scroll { overflow:auto; }
.os-detail .os-stages { min-width:760px; }
.os-detail .os-stage-status, .os-detail .os-status { display:inline-block; border-radius:999px; background:#f0ecf7; color:#514369; padding:4px 8px; font-size:12px; font-weight:700; }
.os-detail .os-stage-status.done, .os-detail .os-status[data-status="4"] { background:#e1f3e9; color:#215b42; }
.os-detail .os-status[data-status="99"] { background:#fbe7ed; color:#8d2543; }
.os-detail .os-status[data-status="3"] { background:#e6effb; color:#254e82; }
.os-detail .os-status[data-status="2"] { background:#fff1d6; color:#785100; }
.os-detail .btnGeral { display:inline-block; max-width:100%; white-space:normal; padding:7px 11px; margin:3px; border:1px solid #ded5e4; border-radius:8px; color:#7c174e; background:#fff; text-decoration:none; font-weight:650; }
.os-detail .btnGeral.os-cancel { color:#8d2543; border-color:#edccd6; background:#fff2f5; }
.os-detail input, .os-detail select, .os-detail textarea, .os-detail table { max-width:100%; }
.os-detail .os-stage-scroll .os-stages { max-width:none; }
.os-detail iframe { display:block; width:100%; height:230px; margin-top:10px; border:1px solid #eeeaf2; border-radius:8px; }
.os-detail a:focus-visible { outline:2px solid #9e125c; outline-offset:2px; }
@media (max-width:650px) {
    .os-detail fieldset.os-panel { padding:10px; }
    .os-detail .os-summary, .os-detail .os-summary > tbody, .os-detail .os-summary > tbody > tr, .os-detail .os-summary > tbody > tr > td { display:block; width:100%; }
    .os-detail .os-summary > tbody > tr > td:first-child:not([colspan]) { padding-bottom:5px; border-bottom:0; }
}
</style>
<script>
if (!document.querySelector('meta[name="viewport"]')) {
    var osViewport = document.createElement('meta');
    osViewport.name = 'viewport';
    osViewport.content = 'width=device-width, initial-scale=1';
    document.head.appendChild(osViewport);
}
</script>
<section class="os-detail" aria-labelledby="os-detalhe-titulo">
    <header class="os-heading">
        <div>
            <div class="os-context">Produ&ccedil;&atilde;o p&oacute;s-venda</div>
            <h2 id="os-detalhe-titulo"><%If comando = "visualizar" Then%>Detalhes da ordem de servi&ccedil;o<%Else%><%=titulo_pagina%><%End If%></h2>
        </div>
        <a class="os-back" href="ppv_ordens_servico.asp">Voltar &agrave;s ordens</a>
    </header>
    <nav aria-label="Produ&ccedil;&atilde;o p&oacute;s-venda">
        <!-- #include file="includes/classe_modulo_ppv_menu.asp" -->
    </nav>
<%
Dim msg : msg = Request.QueryString("msg")

If msg <> Empty Then
	Response.Write("<div class='msg'>"&msg&"</div>")
End If
%>
<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post'>
    <fieldset class="os-panel">
    	<legend>Dados Básicos da OS</legend>
		<input type="hidden" value="<%=pagina_origem%>" name="pagina_retornar" />   
    	
    	<table class="os-table os-summary">
		<%If comando <> "incluir" Then %>
		<tr>
    		<td width='27%'><strong>Ordem de Servi&ccedil;o</strong></td>
    		<td width='73%'><strong><%=moduloPPVOSMascara(cod_ordem_servico,"",false)%></strong>
    			<%If isNumeric(numero_ordem_servico_indice) Then
				'Se NÃO for n�merico significa que foi inserido manualmente e o ind�ce � "U" %>
    			<br />
    				<a href="lca_busca_detalhamento_pedido.asp?cod_pedido=<%=cod_pedido%>#ordens_servico">Outras Ordens de SERVIÇO da venda <%=cod_pedido%></a>
				<%End If%>
    		</td>
    	</tr>
    	<%End If%>
    	<tr>
    		<td width='27%'><strong>Status</strong></td>
    		<td width='73%'><span class="os-status" data-status="<%=Server.HTMLEncode(status_ordem_servico & "")%>"><%
                Select Case CStr(status_ordem_servico & "")
                    Case "1": Response.Write "Em produ&ccedil;&atilde;o"
                    Case "2": Response.Write "Aguardando despacho"
                    Case "3": Response.Write "Despachada"
                    Case "4": Response.Write "Entregue"
                    Case "99": Response.Write "Cancelada"
                    Case Else: Response.Write moduloPPVStatus(status_ordem_servico)
                End Select
            %></span></td>
    	</tr>
    	<tr>
    		<td width='27%'><strong>Formando</strong></td>
    		<td width='73%'><%=nome_formando%> (<%=digito(cod_formando)%>)</td>
    	</tr>
    	<tr>
    		<td width='27%'><strong>Produto / Quantidade</strong></td>
    		<td width='73%'><div align='left'>
    		<%if insercao_manual then
    			response.write (padraoFormularioSelect("COD_PRODUTO","","inputTamanho3","","","Selecione",cod_produto,"SELECT COD_PRODUTO, NOME_PRODUTO FROM `TB_PRODUTOS` WHERE `ATIVO` = 1 ORDER BY NOME_PRODUTO ", comando, false, false, false))
    		else

        if (cod_produto = 34 or cod_produto = 241) and validaUsuario(2, 0, 87, false) then
        
          response.write (padraoFormularioSelect("COD_PRODUTO","","inputTamanho2","","","Selecione",cod_produto,"SELECT COD_PRODUTO, NOME_PRODUTO FROM `TB_PRODUTOS` WHERE `ATIVO` = 1 AND COD_PRODUTO IN (34, 241) ORDER BY NOME_PRODUTO ", "alterar", false, false, false)& " / " & quantidade)


          %>
            <a href="javascript:OSaction('alterarProduto','<%=cod_ordem_servico_etapa_atual%>', '', <%=cod_ordem_servico%>)" class="btnGeral">ALTERAR PRODUTO</a> 
          <%
          
        else
			  	response.write (padraoFormularioSelect("COD_PRODUTO","","inputTamanho3","","","Selecione",cod_produto,"SELECT COD_PRODUTO, NOME_PRODUTO FROM `TB_PRODUTOS` WHERE `ATIVO` = 1 ORDER BY NOME_PRODUTO ", "visualizar", false, false, false)& " / " & quantidade)    			
        end if


    		end if 
    		%>
		    </div></td>
    	</tr>
    	<%
		AbrirConexao()
		'///Abrir para verifica��o do nome do usu�rio
		%>
		<tr>
			<td width='27%'><strong>Data Cadastro / Resp.</strong></td>
    		<td width='73%'>
    			<%
    			response.write (ArrumaData(data_cadastrado))
    			If data_cadastrado <> Empty Then 
    				Response.Write(" / "&nomeUsuario(cod_usuario_processou,true)) 
    			End If%>&nbsp;
    		</td>
    	</tr>
    	<tr>
    		<td width='27%'><strong>Data Base In&iacute;cio do Prazo / Data Prazo Geral Limite</strong></td>
    		<td width='73%'>
    			<%=ArrumaData(data_base_prazo_cliente_inicio)& " / " & ArrumaData(DateAdd("d",prazo_cliente_total,data_base_prazo_cliente_inicio))%>&nbsp;</td>
    	</tr>
    	<tr>
    		<td width='27%'><strong>Data Finalizado / Resp.</strong></td>
    		<td width='73%'>
    			<%
				response.write(ArrumaData(data_finalizado))
				If data_finalizado <> Empty Then 
					Response.Write(" / "&nomeUsuario(cod_usuario_finalizou,true))
				End If
      			%>&nbsp;
      		</td>
    	</tr>
    	<tr>
    		<td width='27%'><strong>Data Despachado / Resp.</strong></td>
    		<td width='73%'>
    			<%
    			response.write(ArrumaData(data_despachado))
    			If data_despachado <> Empty Then 
    				Response.Write(" / "&nomeUsuario(cod_usuario_despachou,true)) 
    			End If
    			%>&nbsp;
    		</td>
    	</tr>
    	<tr>
    		<td width='27%'><strong>Data Cancelado / Motivo / Resp.</strong></td>
    		<td width='73%'>
				<%
				response.write (ArrumaData(data_cancelado))
      			If data_cancelado <> Empty Then
      				Response.Write(" / "&motivo_cancelada&" / "&nomeUsuario(cod_usuario_cancelou,true))
      			End If
      			%>&nbsp;
      		</td>
    	</tr>
    	<%
		'///----
		FecharConexao()
		
		If insercao_manual Then %>
		<tr>
			<td width='27%'><strong>Motivo da Geração manual</strong></td>
			<td width='73%'>
				<div align='left'>
          			<%If comando <> "visualizar" and comando <> "excluir"  Then%>
         				<table style='width:420px;' border='0' cellspacing='0' cellpadding='0'>
            				<tr>
              					<td width='57'>
              						<textarea rows='5' name='MOTIVO_GERADO_MANUALMENTE' class='inputTamanho3' onkeydown='textCounter(this.form.MOTIVO_GERADO_MANUALMENTE,this.form.charRestantes_MOTIVO_GERADO_MANUALMENTE,255);' onkeyup='textCounter(this.form.MOTIVO_GERADO_MANUALMENTE,this.form.charRestantes_MOTIVO_GERADO_MANUALMENTE,255);'><%=MOTIVO_GERADO_MANUALMENTE%></textarea>
              					</td>
							</tr>
            				<tr>
              					<td>
              						<table width='50' border='0' align='right' cellpadding='0' cellspacing='0'>
                						<tr>
                  							<td>
                  								<div align='center'>
                    								<input class='inputContaCaracteres' tabindex='10000' type='text' name='charRestantes_MOTIVO_GERADO_MANUALMENTE' value="<%=255- Len(MOTIVO_GERADO_MANUALMENTE)%>" />
                    							</div>
                    						</td>
										</tr>
                					</table>
                				</td>
              				</tr>
            			</table>
          			<%
          			Else
          				If not isNull(MOTIVO_GERADO_MANUALMENTE) Then 
          					 response.write(Replace(MOTIVO_GERADO_MANUALMENTE,Chr(13),"<br />"))
          				End If
          			End If
          			%>
          		</div>
          	</td>
    	</tr>
    	<%
    	End If

		If comando <> "visualizar" Then
		%>
    	<tr>
    		<td colspan='2' align='center'>
    			<div align='left'><strong>* - Campos obrigatórios</strong><br /></div>
    		</td>
    	</tr>
    	<% Else %>   
        <tr>
    		<td colspan='2' align='center'>
			<br />
			<%
			If int(status_ordem_servico) = 1 and int(cod_usuario_responsavel_atual) = int(Session("cod_usuario")) Then
			'OS em andamento e o usu�rio logado � o responsavel atual

				If isDate(data_inicio_atual) and numero_etapa_atual = qtd_etapas Then
				'�ltima etapa
					If pedido_compra_gerado = "" and necessita_gerar_pedido_compra = "S" Then
					'NÃO permite finalizar etapa
					%>
	                	<a href="javascript:alert('NÃO é permitido finalizar a Ordem de SERVIÇO, pois NÃO foi efetuada a Compra ao fornecedor.\nEstá Ordem de SERVIÇO será finalizada automaticamente com a baixa na entrada do item.')" class="btnGeral">FINALIZAR ORDEM DE SERVIÇO</a> 
					<%Else%>
                    	<a href="javascript:OSaction('finalizarOS','<%=cod_ordem_servico_etapa_atual%>','', <%=cod_ordem_servico%>)" class="btnGeral">FINALIZAR ORDEM DE SERVIÇO</a> 
                    <%End If
            		
                Elseif isDate(data_inicio_atual) <> Empty and int(numero_etapa_atual) <> int(qtd_etapas) Then
					'Etapas intermedi�rias
					If necessita_gerar_pedido_compra = "S" Then
					'NÃO permite finalizar etapa%>
	                    <a href="javascript:alert('NÃO é permitido finalizar etapa, pois NÃO foi efetuada a Compra ao fornecedor.\nEstá etapa será finalizada automaticamente com a baixa na entrada do item.')" class="btnGeral">FINALIZAR ETAPA</a> 
                    <%Elseif Cstr(descricao) = "triagem" Then 'Padronizado minusculo%>
	 					    <a href="javascript:alert('NÃO é permitido finalizar a etapa, pois a triagem NÃO foi realizada pelo responsável.')" class="btnGeral">FINALIZAR ETAPA</a> 
                    <%Else%>               
                        <a href="javascript:OSaction('finalizarEtapa','<%=cod_ordem_servico_etapa_atual%>','', <%=cod_ordem_servico%>)" class="btnGeral">FINALIZAR ETAPA</a>
                    <%End If
            	Else 'Etapa NÃO iniciada
				%>
                	<a href="javascript:OSaction('iniciarEtapa','<%=cod_ordem_servico_etapa_atual%>','',<%=cod_ordem_servico%>)" class="btnGeral">INICIAR ETAPA</a>	
                <%
				End If
			End If

			If validaUsuario(2,,56,false) Then
			%>
   	    		<a href="javascript:trocarRespEtapa(false)" class="btnGeral">TROCAR RESPONSABILIDADE</a>  
        	<%
			End If
			'--------------------------------------------------------------------------------------
			If validaUsuario(2,,55,false) Then
			%>
            	<a href="javascript:motivoCancelamento()" class="btnGeral os-cancel">CANCELAR ORDEM DE SERVIÇO</a>
                <div id="containerCancelamento" class="ocultarCompleto" style="padding:15px;">   
					<br />
					<br />
					<table width="500" border="0" bgcolor="#FFFFCC" cellspacing="0" cellpadding="10">
  						<tr>
    						<td>
    							<table align="center" style='width:420px;' border='0' cellspacing='0' cellpadding='0'>
      								<tr>
        								<td width='57'>
        									<textarea rows='5' name='MOTIVO_CANCELADA' class='inputTamanho3' onkeydown='textCounter(this.form.MOTIVO_CANCELADA,this.form.charRestantes_MOTIVO_CANCELADA,255);' onkeyup='textCounter(this.form.MOTIVO_CANCELADA,this.form.charRestantes_MOTIVO_CANCELADA,255);'><%=MOTIVO_CANCELADA%></textarea>
        								</td>
        							</tr>
      								<tr>
        								<td>
        									<table width='50' border='0' align='right' cellpadding='0' cellspacing='0'>
												<tr>
													<td>
														<div align='center'>
															<input class='inputContaCaracteres' style=" background-color:#FFFFCC" tabindex='10000' type='text' name='charRestantes_MOTIVO_CANCELADA' value="<%=255-Len(MOTIVO_CANCELADA)%>" />
														</div>
													</td>
												</tr>
          									</table>
          								</td>
        							</tr>
      								<tr>
        								<td>
          									<a href="javascript:OSaction('cancelarOS','<%=cod_ordem_servico_etapa_atual%>','', <%=cod_ordem_servico%>)" class="btnGeral">CONFIRMAR</a>&nbsp;&nbsp;<a href="javascript:motivoCancelamento();" class="btnGeral">VOLTAR</a></td>
        							</tr>
								</table>
							</td>
  						</tr>
					</table>
                </div>
    		<% End If %>
	            <div id="trocarResponsabilidade" style="padding:15px;" class="ocultarCompleto"><br />
					<table style="background-color:#FFFFCC" width="500" border="0" align="center" cellpadding="5" cellspacing="0">
	                    <tr>
							<td width="29%">Trocar responsável para:</td>
							<td width="71%">
								<%=padraoFormularioSelect("trocaResp_responsavel_cod_usuario","","inputTamanho2","","","SEM_SELECIONE","" ,"Select cod_usuario, nome FROM TB_ADM_USUARIOS where bloqueado = 2 and cod_usuario <> '"&cod_usuario_responsavel_atual&"' order by nome", "incluir", false, false, false)%>
							</td>
						</tr>
	                    <tr>
							<td colspan="2">Trocar responsabilidade de(a): 
<br>

              <input class="inputRadio" id='trocaResp_responsavel_cod_usuario_parametro2_1' name='trocaResp_responsavel_cod_usuario_parametro2' type='radio' value="S" /> 
              <label for='trocaResp_responsavel_cod_usuario_parametro2_1'> Todas as etapas em aberto <br>
              <input class="inputRadio" name='trocaResp_responsavel_cod_usuario_parametro2' id='trocaResp_responsavel_cod_usuario_parametro2_2' type='radio' value="N" /> <label for='trocaResp_responsavel_cod_usuario_parametro2_2'> somente etapa atual </label>
    
	                    </tr>
	                    <tr>
	                    	<td colspan="2"><div align="center"><br />
								<a href="javascript:OSaction('trocarRespOS','<%=cod_ordem_servico_etapa_atual%>','', <%=cod_ordem_servico%>)" class="btnGeral">CONFIRMAR</a> &nbsp;<a href="javascript:trocarRespEtapa(true)" class="btnGeral">CANCELAR</a>
							</div></td>                             
						</tr>
					</table>
					<br />
				</div>
	            <br />
				<br />
	            <br />&nbsp;
	            <a name="trocar_responsabilidade"></a> 
    		</td>
		</tr> 
    <%End If%>
    </table>
</fieldset>

<fieldset class="os-panel">
<legend>Descrição / Histórico</legend>
    <%If comando <> "incluir" Then%>
    	<div align="center"><a href="ppv_ordens_servico_formulario_consulta_historico.asp?COD_ORDEM_SERVICO=<%=COD_ORDEM_SERVICO%>" target="_blank">Abrir Histórico em nova janela</a></div>
      	<iframe title="Hist&oacute;rico da ordem de servi&ccedil;o" width="98%" height="220" src="ppv_ordens_servico_formulario_consulta_historico.asp?COD_ORDEM_SERVICO=<%=COD_ORDEM_SERVICO%>" frameborder="0"></iframe>
     <%End if%>
</fieldset>

<fieldset class="os-panel">
	<legend>Etapas</legend>
    	<div class="os-stage-scroll" tabindex="0" role="region" aria-label="Tabela de etapas">
        <table class="os-table os-stages">
		<tr bgcolor="whitesmoke">
    		<td width="7%"><div align="center" style="font-weight: bold">Status</div></td>
        	<td width="20%"><div align="center" style="font-weight: bold">Etapa</div></td>
        	<td width="21%"><div align="center" style="font-weight: bold">Data In&iacute;cio</div></td>
        	<td width="11%"><div align="center" style="font-weight: bold">Data T&eacute;rmino</div></td>
        	<td width="7%"><div align="center" style="font-weight: bold">Prazo em dias corridos</div></td>
        	<td width="8%"><div align="center"><strong>Necessita Gerar Pedido de Compra</strong></div></td>
        	<td width="17%"><div align="center" style="font-weight: bold">Respons&aacute;vel</div></td>
        	<td width="9%"><div align="center"></div></td>
		</tr>
		<%
		AbrirConexao()
		
		Set rs=objConexao.Execute("Select * From TB_ORDENS_SERVICO_ETAPAS where cod_ordem_servico = "&cod_ordem_servico&" order by numero_etapa asc")
		While not rs.Eof
		%>
        <tr>
    		<td width="7%"><div align="center">
        	<% If rs("data_termino") Then %>
        		<span class="os-stage-status done">Finalizada</span>
        	<% Else %>
       			<span class="os-stage-status">Pendente</span> 
        	<% End If %>
       		</div></td>
        	<td width="20%">
				<div align="left">
					<a href="javascript:info('Instruções desta etapa: <%=rs("instrucoes")%>')">
            			<%=rs("numero_etapa")%>&ordf;) <%=rs("descricao")%>
            		</a>
          		</div>
          	</td>
          	<td width="21%"><div align="center"><%=ArrumaData(rs("data_inicio"))%></div></td>
        	<td width="11%"><div align="center"><%=ArrumaData(rs("data_termino"))%></div></td>
        	<td><div align="center"><%=ArrumaNumero(rs("prazo_em_dias"),"-",0)%></div></td>
        	<td><div align="center"><%=booleano(rs("necessita_gerar_pedido_compra"))%>
          	<%
				If Cstr(""&rs("necessita_gerar_pedido_compra")) = "S" Then
					Response.Write(" (")
					If Cstr(""&rs("pedido_compra_gerado")) <> "" Then
						Set rs2=objConexao.Execute("SELECT numero_pedido_compra FROM TB_PEDIDOS_COMPRAS WHERE cod_pedido_compra = "&sqlTexto(rs("pedido_compra_gerado")))

							If not rs2.Eof Then
								Response.Write("Nº "&rs2(0))
							End If
						rs2.Close
						Set rs2=Nothing
					Else
						Response.Write("Pedido NÃO gerado")			
					End If
					Response.Write(")")
				End If
			%>
        	</div></td>
        	<td width="17%"><div align="center">
			<%
				If int(rs("cod_usuario_responsavel")) > 0 Then
					Response.Write(nomeUsuario(rs("cod_usuario_responsavel"),true))
				End If
			%>
        	</div></td>
        	<td width="9%"><div align="center">
        	<%
				'Op��o para voltar quando a etapa estiver finalizada E quando a OS estiver em andamento E se o usu�rio possuir acesso
				If rs("data_termino") AND Cstr(""&status_ordem_servico) = "1" and validaUsuario(2,,67,true)  Then
				%>
            		<a href="javascript:OSaction('voltarEtapa','<%=rs("cod_ordem_servico_etapa")%>','', <%=cod_ordem_servico%>)">Voltar para esta etapa</a></div>
            	<% 
				End If
			%>
            </td>
        </tr>
    	<%
		rs.Movenext()
		Wend
	rs.Close
	Set rs=Nothing
	FecharConexao()
	%>
	<tr bgcolor="whitesmoke">
    	<td colspan="8"><span style="font-weight: bold">
        	Etapa: <span class="textoMedio"><%=numero_etapa_atual%> de <%=qtd_etapas%></span></span></td>
    </tr>
	</table>
    </div>
</fieldset>
<input type='hidden' value='1' name='campo1'  />
<input type='hidden' name='COD_ORDEM_SERVICO' value="<%=COD_ORDEM_SERVICO%>" />
</form>
</section>
<script type='text/javascript'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('COD_PRODUTO#Cód. Produto#2|MOTIVO_GERADO_MANUALMENTE#Motivo da Geração manual#5',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>
<!-- #include file="files/includes/miolo-inferior.asp" -->