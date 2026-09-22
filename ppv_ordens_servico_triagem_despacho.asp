<%
Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
BTN_VOLTAR = 1
BTN_IMPRIMIR = 1
BTN_FONTE = 2
BTN_IMPRIMIR_LINK = "#"

Dim titulo_pagina
titulo_pagina = "Triagem Entrega | Transpostadora"

Dim calendario_necessario
calendario_necessario = "NAO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<%
Dim efetuar_despacho, despacho_transportador
efetuar_despacho = Cstr(""&Request.Form("campo1"))	
despacho_transportador = Cstr(""&Request.Form("despacho_transportador"))	

If efetuar_despacho = "S" And despacho_transportador <> "" Then

	AbrirConexao()
		'///
		'Limpar base tempor&aacuteria
			objConexao.Execute("DELETE FROM TB_ORDENS_SERVICO_TEMPORARIA WHERE cod_usuario = "&Session("cod_usuario"))
		'///
		'1) Selecionar todas OS para triagem
		Set rs=objConexao.Execute("SELECT cod_ordem_servico, cod_pedido FROM TB_ORDENS_SERVICO WHERE status_ordem_servico = 2 ORDER BY cod_pedido asc")
			If not rs.Eof then

				While not rs.Eof
					
					'3) Verifica se o respons&aacutevel foi selecionado
					cod_pedido = rs("cod_pedido")			
					cod_ordem_servico = rs("cod_ordem_servico")	
						
					despachar = Cstr(""&Request.Form("despachar_"&cod_pedido))
					despachar_item = Cstr(""&Request.Form("item_"&cod_ordem_servico))
					cod_ordem_servico = rs("cod_ordem_servico")
					
					If despachar_item = "S" Then
						
						'1) Altera status 
						sql = "UPDATE TB_ORDENS_SERVICO SET cod_usuario_despachou = "&Session("cod_usuario")&",	data_despachado = NOW(), despacho_transportador = "&sqlTexto(despacho_transportador)&", status_ordem_servico = 3 WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" AND status_ordem_servico = 2 AND data_despachado IS NULL"
						objConexao.Execute(sql)
						
						'2) Grava na base temporaria para emissao do recibo
						sql = "INSERT INTO TB_ORDENS_SERVICO_TEMPORARIA (cod_ordem_servico,cod_usuario) VALUES ('"&cod_ordem_servico&"', '"&Session("cod_usuario")&"')"
						objConexao.Execute(sql)						
					End If
					
								
				rs.MoveNext()
				Wend
				
			End if
			
		rs.close
		Set rs=nothing
		
	FecharConexao()
	
	Response.Redirect("ppv_ordens_servico.asp?codmsg=301")
	
End If
	
	
Dim busca_cod_ordem_servico, busca_nome_formando, busca_cod_contrato, busca_cod_formando, busca_finalizados_parcialmente, busca_cod_usuario_responsavel

Dim sql, sqlWhere, sqlOrderBy, sqlGroupBy

sqlOrderBy = "cod_contrato, nome_formando"
busca_cod_ordem_servico = Cstr(""&Request.QueryString("cod_ordem_servico")) 'N&atildeo est&aacute sendo utilizado
busca_nome_formando = Cstr(""&Request.QueryString("nome_formando"))
busca_cod_contrato = Cstr(""&Request.QueryString("cod_contrato"))
busca_cod_formando = Cstr(""&Request.QueryString("cod_formando"))
busca_finalizados_parcialmente = Cstr(""&Request.QueryString("finalizados_parcialmente"))
busca_cod_usuario_responsavel = Cstr(""&Request.QueryString("cod_usuario_responsavel"))

'Padr&atildeo
Dim validacao_listar_os_todos_usuarios, valida_para_despacho_qualquer_usuario
validacao_listar_os_todos_usuarios = validaUsuario(2,,57,false)
valida_para_despacho_qualquer_usuario = validaUsuario(2,,76,false)

If not valida_para_despacho_qualquer_usuario Then
	sqlWhere = sqlWhere&" and cod_usuario_responsavel = "&sqlTexto(Session("cod_usuario"))	
End If

'Por padr&atildeo lista somente os pedidos que est&atildeo totalmente prontos para despacho
If Request.QueryString("filtro1") <> Empty Then
		If busca_cod_ordem_servico <> "" Then
		
		Else

			If busca_finalizados_parcialmente <> "S" Then
			'Se a opcao de finalizados parcialmente estiver desabilidata ira filtrar somente TOTALMENTE finalizados para despacho
				sqlWhere = sqlWhere&" AND cod_ordem_servico NOT IN (SELECT cod_ordem_servico FROM TB_ORDENS_SERVICO WHERE status_ordem_servico IN(0,1))"
			End If
					
			If busca_cod_usuario_responsavel <> "" Then
				sqlWhere = sqlWhere&" and cod_usuario_responsavel = "&sqlTexto(busca_cod_usuario_responsavel)	
				sql_usuario = " and cod_usuario_responsavel = "&sqlTexto(busca_cod_usuario_responsavel)	
			End If
			
			If busca_nome_formando <> "" Then
				sqlWhere = sqlWhere&" and nome_formando LIKE '%"&busca_nome_formando&"%'"	
			End If

			If busca_cod_contrato <> "" Then
				sqlWhere = sqlWhere&" and cod_contrato = "&sqlTexto(busca_cod_contrato)			
			End If

			If busca_cod_formando <> "" Then
				temp = Right(busca_cod_formando,1)
				
				If Cstr(""&temp) <> "1" Then
					'Response.Redirect("ppv_ordens_servico_triagem_despacho.asp?codmsg=701&sucesso=2")
				Else
					sqlWhere = sqlWhere&" and cod_formando = "&sqlTexto(Mid(busca_cod_formando,1,Len(busca_cod_formando)-1))
				End if
		
			End If					
			
		End If
Else
	sqlWhere = sqlWhere&" AND cod_ordem_servico NOT IN (SELECT cod_ordem_servico FROM TB_ORDENS_SERVICO WHERE status_ordem_servico IN(0,1))"

End If

Dim acesso_inclusao_exclusao_os_manual
acesso_inclusao_exclusao_os_manual = validaUsuario(2,,51,false)
%>

<h2>PPV - Despacho para Cliente</h2>
<br />
<!-- #include file="includes/classe_modulo_ppv_menu.asp" -->
<br />
<br />
	<div class="msgGrande">
	<%=Response.Write("<a href='ppv_ordens_servico_triagem_despacho_imprimir.asp' target='_blank'>Imprimir o &uacuteltimo recibo novamente</a>")%>
	</div>

<form style="margin:0px">
	<table style="border:#999999 dashed 1px;  margin-bottom:7px;padding:2px; background:#FFFFCC" width="99%" border="0" align="center" cellpadding="5" cellspacing="0">
		<tr>
			<td width="150px"><strong>Nome do Formando</strong></td>
			<td width="300px">
				<input type="text" value="<%=busca_nome_formando%>" name="nome_formando" class="inputTamanho2" />
			</td>
			<td rowspan="5">
				<div align="left">
					<input class="btnFiltrarAmarelo" type="submit" name="filtro1" value=" " />
				</div>
			</td>
			<td rowspan="5">
				<div align="center">
					<div style="width:250px;" class='msgGrandeErro'>
						<em><strong>Listagem em Ordem crescente pelo contrato, nome do formando com todas OS, cujo os produtos n&atildeo foram despachados para o cliente.</strong></em><br />
					</div>
				</div>
			</td>
			
		</tr>
		<tr>
			<td>
				<strong>C&oacuted. Formando Ficha</strong>
				<a href="javascript:alert('(Cod. do formando acressido de \'1\' no final)')">[?]</a>
			</td>
			<td>
				<input type="text" value="<%=busca_cod_formando%>" name="cod_formando" class="inputTamanho2" />
			</td>
		</tr>
		<tr>
			<td>
				<strong>Cod. Contrato</strong>
			</td>
			<td>
				<input type="text" value="<%=busca_cod_contrato%>" name="cod_contrato" class="inputTamanho2" />
			</td>
		</tr>
		<tr>
			<td><strong>Respons&aacutevel  OS</strong></td>
			<td>
				<%If valida_para_despacho_qualquer_usuario Then%>
					<%=PadraoFormularioSelect("cod_usuario_responsavel","","inputTamanho2","","","Todos",busca_cod_usuario_responsavel,"SELECT cod_usuario, nome FROM TB_ADM_USUARIOS WHERE COD_USUARIO IN(SELECT cod_usuario_responsavel FROM VW_TB_ORDENS_SERVICO WHERE status_ordem_servico = 2 GROUP BY cod_usuario_responsavel) ORDER BY NOME", "incluir", false, false, false )%>
				<%Else%>
					<strong style="font-size:14px;"><%=nomeUsuario(Session("cod_usuario"),false)%></strong>
				<%End If%>
			</td>
		</tr>
		<tr>
			<td><div align="left"><strong> Filtrar pedidos que est&atilde;o finalizados parcialmente</strong></div></td>
			<td>
				<span style="padding-left:5px">
					<input type="checkbox" class="inputRadio" value="S" <% if finalizados_parcialmente = "S" then%> checked="checked"<%end if%> name="finalizados_parcialmente" />
				</span>
			</td>
		</tr>
	</table>
</form>

<div class="subMenu">
	<a href='ppv_ordens_servico_triagem_despacho_alterar.asp'>Alterar pedido j&aacute despachado</a>
</div>

<br />
<%						
sql = "SELECT cod_pedido FROM VW_TB_ORDENS_SERVICO WHERE status_ordem_servico = 2 "&sqlWhere&" GROUP BY cod_pedido ORDER BY cod_contrato, nome_formando LIMIT 100"

AbrirConexao()

Set rs=objConexao.Execute(sql)
If not rs.Eof Then

'Verifica a qtd de OS de acordo com o filtro
Set rs2=objConexao.Execute("SELECT COUNT(*) as qtd FROM VW_TB_ORDENS_SERVICO WHERE status_ordem_servico = 2 "&sqlWhere)
If not rs2.Eof Then
qtdOSs = rs2(0)
End If
rs2.Close
Set rs2=Nothing
%>

<form name="DESPACHO_CLIENTE" action="ppv_ordens_servico_triagem_despacho.asp" method="post" style="margin:0px">
	<input name="campo1" value="S" type="hidden" />
	<table width="99%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
		<tr bgcolor="#E1EDFF">
		<td colspan="2" style="font-size:16px;"><div align="center"><strong> Informe o Transportador</strong> *</div></td>
		<td colspan="5">
			<%=PadraoFormularioSelect("despacho_transportador","","inputTamanho3","","","Selecione",letra,"PADRAO:TRANSPORTADORA#TRANSPORTADORA|CORREIOS#CORREIOS|RETIRADA#RETIRADA|STILLOS#STILLOS|STILLOS#STILLOS|PANDEMIA#PANDEMIA|LOGISTICA#LOGISTICA|ESPECIAL#ESPECIAL", "incluir", false, false, false )%>
		</td>
		</tr>
			<tr bgcolor="whitesmoke">
			<td colspan="2"><div align="center"><strong> OS</strong></div></td>
			<td width="13%"><div align="center"><strong>Produto</strong></div></td>
			<td width="12%"><div align="center"><strong>Data Finalizada</strong></div></td>
			<td width="12%"><div align="center"><strong>Prazo Geral</strong></div></td>
			<td width="27%"><div align="center"><strong>Dias faltantes para prazo Geral</strong></div></td>
			<td width="11%"><div align="center"><strong>Despachar para cliente</strong></div></td>
		</tr>
		<%
		Dim numero_etapa, descricao, cod_usuario_responsavel, nome_usuario_responsavel, prazo_em_dias, data_inicio, prazo_limite, diferenca, qtd_etapas
		While not rs.Eof
			Set rs2=objConexao.Execute("SELECT * FROM TB_ORDENS_SERVICO WHERE cod_pedido = "&rs("cod_pedido")&" and data_finalizado IS NULL and status_ordem_servico NOT IN (0,99)")
			If not rs2.Eof Then
				permitido = false
			Else
				permitido = true							
			End If
			rs2.Close
			Set rs2=Nothing
			%>
			<tr>
				<td colspan="6">
					<table width="650" border="0" cellspacing="0" cellpadding="1">
					<tr>
						<td width="66" class="textoMedio"><strong><%=rs("cod_pedido")%></strong></td>
						<td width="580">
						<%
						sql = " SELECT "&_
						" TB_PEDIDOS_FORMANDOS.COD_FORMANDO, "&_
						" TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, "&_
						" TB_FICHAS_IDENTIFICACAO.COD_CONTRATO, "&_
						" TB_ENTIDADES.DESCRICAO, "&_
						" TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO, "&_
						" TB_FICHAS_IDENTIFICACAO.NOME_CURSO "&_
						" FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON "&_
						" TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO "&_
						" INNER JOIN TB_ENTIDADES ON(TB_FICHAS_IDENTIFICACAO.COD_CONTRATO = TB_ENTIDADES.COD_CONTRATO)"&_
						" WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&rs("cod_pedido")&" ORDER BY NOME_FORMANDO"

						Set rs2=objConexao.Execute(sql)
						response.Write("<table cellpadding='0' width='580' cellspacing='0'>")
						While not rs2.Eof
							response.write("<tr>")
							response.write("<td width='36%'>"&rs2("NOME_FORMANDO")&"</td>")
							response.write("<td width='64%'>"&rs2("DESCRICAO")&"</td>")
							response.write("</tr>")
						rs2.MoveNext()
						Wend
						response.Write("</table>")
						rs2.Close
						Set rs2=Nothing
						%>
						</td>
					</tr>
					</table>
				</td>
				<td class="textoMedio"><div style="display:none" align="center">
					<%If permitido Then%>
						<input type="checkbox" value="S" name="despachar_<%=rs("cod_pedido")%>" class="inputRadio" />
					<%Else%>
						N&atildeo permitido
					<%End If%>
				</div></td>
			</tr>
			<%
			sql = "SELECT VW_TB_ORDENS_SERVICO.*, A.data  FROM VW_TB_ORDENS_SERVICO LEFT JOIN TB_PEDIDOS_ITENS_ENTREGUES A ON A.cod_ordem_servico = VW_TB_ORDENS_SERVICO.cod_ordem_servico WHERE cod_pedido = "&rs("cod_pedido")&" and status_ordem_servico NOT IN (0,99) "&sql_usuario&" order by numero_ordem_servico desc, numero_ordem_servico_indice asc"
			'Response.Write(sql)
			Set rs2=objConexao.Execute(sql)
			While not rs2.Eof

				data_finalizado = rs2("data_finalizado")
			
				Set rs3=objConexao.Execute("Select count(*) as qtd_etapas From TB_ORDENS_SERVICO_ETAPAS where cod_ordem_servico = "&rs2("cod_ordem_servico")&" group by cod_ordem_servico")
				If not rs3.Eof Then
					qtd_etapas = formatNumber(0&rs3("qtd_etapas"),0)
				End If
				rs3.Close
				Set rs3=Nothing

			%>                   
			<tr <%If not isDate(data_finalizado) Then%> bgcolor="#FFAAAA" <%End If%>>
				<td width="3%">
					<div align="right">
						<img src="images/icons/seta_superior.gif" width="19" height="15" />
					</div>
				</td>
				<td width="22%">
					<div align="center">
						<%=moduloPPVOSMascara(rs2("cod_ordem_servico"),"",true)%>
					</div>
				</td>
				<td>
					<div align="left">
						<%=rs2("nome_produto")%>                                        
					</div>
				</td>
				<td>
					<div align="center">
						<%=ArrumaData(data_finalizado)%>
					</div>
				</td>
				<td>
					<div align="center">
						<%=ArrumaData(DateAdd("d",rs2("prazo_cliente_total"),rs2("data_base_prazo_cliente_inicio")))%>
					</div>
				</td>
				<td>
					<div align="center">
						<%=dateDiff("d",date(),DateAdd("d",rs2("prazo_cliente_total"),rs2("data_base_prazo_cliente_inicio")))%>
					</div>
				</td>
				<td>
					<div align="center">
						<%If len(rs2("data")) > 3 Then%>
							Entregue
						<%elseIf len(rs2("data_despachado")) > 3 Then%>
							Despachado
						<%Else%>
							<input type="checkbox" value="S" name="item_<%=rs2("cod_ordem_servico")%>" class="inputRadio" />
						<%End If%>
					</div>
				</td>
			</tr>

			<%
			rs2.MoveNext()
			Wend

			rs2.Close
			Set rs2=Nothing

		rs.MoveNext()
		Wend
	%>
<tr>
<td colspan="7"><div align="center"><a href="javascript:confirmaDespacho();"><img border="0" src="images/layout/btnGravar.jpg" /></a></div></td>
</tr>

</table>
</form>

<%
Else
%>
<div align="center" class="msgGrande">
Nenhuma OS encontrada para despacho!<br />
</div>
<%End if

rs.Close
Set rs=Nothing

FecharConexao()
%>
</p>
<br />
<br />

Observa&ccedil;&otilde;es:<br />
Linhas em vermelho significa que n&atilde;o ser&aacute; poss&iacute;vel despachar sem antes finalizar as outras OS da mesma venda
<br />

</td>
</tr>
</table>
<script type="text/javascript">
	function confirmaDespacho(){
		alert($("select[name='DESPACHO_TRANSPORTADOR']").val());
		if($("select[name='DESPACHO_TRANSPORTADOR']").val() !== ''){
			if(YNconfirm('Tem certeza que deseja efetuar o despacho das OSs selecionadas?')){
				document.DESPACHO_CLIENTE.submit();
			}	
		}else{
			alert('Informe o nome do transportador antes de continuar. Ex.: Correio.');
			document.DESPACHO_CLIENTE.despacho_transportador.focus();
		}
	}
</script>
<script type="text/javascript">

function alterarOSAnterior(acao)
{
	if(acao == "exibir")
	{
		tinybox.show
	}
}

</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->