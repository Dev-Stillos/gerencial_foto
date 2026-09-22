<%
Dim calendario_necessario
Calendario_necessario = "SIM"
 
Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 1
BTN_IMPRIMIR_LINK = "#"

titulo_pagina = "Compras - Listagem por Item Solicitados"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
  
    <%
	AbrirConexao()
	Dim acesso_entrada_pedido_compra, acesso_excluir_entrada_errada, validacao_para_entrada_qualquer_pedido
	acesso_entrada_pedido_compra = validaUsuario(2,,66,true)
	acesso_excluir_entrada_errada = validaUsuario(2,,69,true)
	validacao_para_entrada_qualquer_pedido = validaUsuario(2,,74,true)
	 
	
	Dim sql, sqlWhere, sqlOrderBy
	
	sql = "SELECT VW_TB_ORDENS_SERVICO.NOME_FORMANDO, VW_TB_ORDENS_SERVICO.ENTIDADE_ENSINO, VW_TB_PEDIDOS_COMPRAS_PRODUTOS.* FROM VW_TB_PEDIDOS_COMPRAS_PRODUTOS LEFT JOIN VW_TB_ORDENS_SERVICO on VW_TB_ORDENS_SERVICO.COD_ORDEM_SERVICO = VW_TB_PEDIDOS_COMPRAS_PRODUTOS.COD_ORDEM_SERVICO WHERE VW_TB_PEDIDOS_COMPRAS_PRODUTOS.status_pedido_compra <> 0"
	'sql = "SELECT * FROM VW_TB_PEDIDOS_COMPRAS_PRODUTOS WHERE status_pedido_compra <> 0 "
	sqlOrderBy =  " order by  VW_TB_PEDIDOS_COMPRAS_PRODUTOS.data_previsao_entrega asc"
	
	Dim status_produto, diferenca_prazo_entrega
	
	'/////////////
	'Filtro
		
		Dim filtro_status_produto
		Dim filtro_cod_fornecedor
		Dim filtro_NUMERO_PEDIDO_COMPRAS
		Dim filtro_COD_CONTRATO
		Dim filtro_ordernar_por
		Dim b_DATA_PREVISAO_ENTREGA_INICIO
		Dim b_DATA_PREVISAO_ENTREGA_FIM
		Dim FILTRO_COD_PRODUTO

		
		filtro_NUMERO_PEDIDO_COMPRAS	=	Cstr(""&Request.QueryString("NUMERO_PEDIDO_COMPRAS"))
		filtro_COD_CONTRATO				=	Cstr(""&Request.QueryString("COD_CONTRATO"))
		filtro_status_produto			=	Cstr(""&Request.QueryString("status_produto"))
		filtro_cod_fornecedor			=	Cstr(""&Request.QueryString("cod_fornecedor"))
		filtro_ordernar_por				=	Cstr(""&Request.QueryString("ordernar_por"))
		b_DATA_PREVISAO_ENTREGA_INICIO	=	Request.QueryString("DATA_PREVISAO_ENTREGA_INICIO")
		b_DATA_PREVISAO_ENTREGA_FIM		=	Request.QueryString("DATA_PREVISAO_ENTREGA_FIM")
		FILTRO_COD_PRODUTO				=	Cstr(""&Request.QueryString("COD_PRODUTO"))
		
		If Not validacao_para_entrada_qualquer_pedido Then
			sqlWhere = sqlWhere&" AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.cadastro_cod_usuario = '"&Session("cod_usuario")&"'"
		End If
		
		If Request.QueryString("filtro") <> "" Then
			
			If filtro_ordernar_por <> "" Then
				If filtro_ordernar_por = "1" Then
					sqlOrderBy =  " order by  VW_TB_PEDIDOS_COMPRAS_PRODUTOS.data_previsao_entrega asc"
				ElseIf filtro_ordernar_por = "2" Then
					sqlOrderBy =  " order by  VW_TB_ORDENS_SERVICO.NOME_FORMANDO asc"
				ElseIf filtro_ordernar_por = "3" Then
					sqlOrderBy =  " order by  VW_TB_ORDENS_SERVICO.ENTIDADE_ENSINO asc"
				End If
				
			End If
			
			If filtro_NUMERO_PEDIDO_COMPRAS <> "" Then
			'Filtro preferencial, limpa os demais
				
				filtro_cod_fornecedor = ""
				filtro_cadastro_cod_usuario = ""
				b_DATA_PREVISAO_ENTREGA_INICIO = ""
				b_DATA_PREVISAO_ENTREGA_FIM	= ""
				filtro_COD_CONTRATO = ""
				FILTRO_COD_PRODUTO = ""
				sqlWhere = " AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.numero_pedido_compra = '"&filtro_NUMERO_PEDIDO_COMPRAS&"'"
				
				If Not validacao_para_entrada_qualquer_pedido Then
					sqlWhere = sqlWhere&" AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.cadastro_cod_usuario = '"&Session("cod_usuario")&"'"
				End If
				
				If filtro_status_produto <> "" Then
					sqlWhere = sqlWhere&" AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.status_produto = '"&filtro_status_produto&"'"
				End If
			
			Else
			
				If filtro_COD_CONTRATO <> "" Then
					sqlWhere = sqlWhere&" AND VW_TB_ORDENS_SERVICO.COD_CONTRATO = '"&filtro_COD_CONTRATO&"'"
				End If
		
				If filtro_status_produto <> "" Then
					sqlWhere = sqlWhere&" AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.status_produto = '"&filtro_status_produto&"'"
				End If
				
				if filtro_cod_fornecedor <> "" Then
					sqlWhere = sqlWhere&" AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.cod_fornecedor = '"&filtro_cod_fornecedor&"'"
				End if
				
				if filtro_cadastro_cod_usuario <> "" Then
					sqlWhere = sqlWhere&" AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.cadastro_cod_usuario = '"&filtro_cadastro_cod_usuario&"'"
				End if	
	
				if FILTRO_COD_PRODUTO <> "" Then
					sqlWhere = sqlWhere&" AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.cod_produto = '"&FILTRO_COD_PRODUTO&"'"
				End if			
				
				
				If b_DATA_PREVISAO_ENTREGA_INICIO <> Empty or b_DATA_PREVISAO_ENTREGA_FIM <> Empty Then
							
							
							If b_DATA_PREVISAO_ENTREGA_INICIO <> Empty Then
								b_DATA_PREVISAO_ENTREGA_INICIO = Day(b_DATA_PREVISAO_ENTREGA_INICIO)&"/"&Month(b_DATA_PREVISAO_ENTREGA_INICIO)&"/"&Year(b_DATA_PREVISAO_ENTREGA_INICIO)&" 00:00:00"
							End If
		
							If b_DATA_PREVISAO_ENTREGA_FIM <> Empty Then
								b_DATA_PREVISAO_ENTREGA_FIM = Day(b_DATA_PREVISAO_ENTREGA_FIM)&"/"&Month(b_DATA_PREVISAO_ENTREGA_FIM)&"/"&Year(b_DATA_PREVISAO_ENTREGA_FIM)&"  23:59:59"
							End If					
																	
							If b_DATA_PREVISAO_ENTREGA_INICIO <> Empty and b_DATA_PREVISAO_ENTREGA_FIM <> Empty Then
								sqlWhere = sqlWhere&" AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.DATA_PREVISAO_ENTREGA BETWEEN "&sqlData(b_DATA_PREVISAO_ENTREGA_INICIO)&" AND "&sqlData(b_DATA_PREVISAO_ENTREGA_FIM)
							Elseif b_DATA_PREVISAO_ENTREGA_INICIO = Empty and b_DATA_PREVISAO_ENTREGA_FIM <> Empty Then
								sqlWhere = sqlWhere&" AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.DATA_PREVISAO_ENTREGA <= "&sqlData(b_DATA_PREVISAO_ENTREGA_FIM)
							Elseif b_DATA_PREVISAO_ENTREGA_INICIO <> Empty and b_DATA_PREVISAO_ENTREGA_FIM = Empty Then
								sqlWhere = sqlWhere&" AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.DATA_PREVISAO_ENTREGA >= "&sqlData(b_DATA_PREVISAO_ENTREGA_INICIO)
							End If
					
							If b_DATA_PREVISAO_ENTREGA_FIM <> Empty Then
							'Corrigi data
								b_DATA_PREVISAO_ENTREGA_FIM = Day(b_DATA_PREVISAO_ENTREGA_FIM)&"/"&Month(b_DATA_PREVISAO_ENTREGA_FIM)&"/"&Year(b_DATA_PREVISAO_ENTREGA_FIM)&"  00:00:00"
							End If	
					
				End If	
			
			
			End If

			
		Else
				'Padrão filtrado como somente pendentes		
				filtro_status_produto = "P"
				sqlWhere = " AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.status_produto = 'P' "
		End If	
		



'Junta tudo
sql = sql&sqlWhere&sqlOrderBy

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
If Request.Form("campo1") <> Empty and acesso_entrada_pedido_compra Then
		
	'Limpa a base temporária para o usuário
	objConexao.Execute("DELETE FROM TB_PEDIDOS_COMPRAS_PRODUTOS_TEMPORARIA WHERE cod_usuario = "&sqlTexto(Session("cod_usuario")))
	
	Set rs=objConexao.Execute("SELECT VW_TB_PEDIDOS_COMPRAS_PRODUTOS.* FROM VW_TB_ORDENS_SERVICO INNER JOIN VW_TB_PEDIDOS_COMPRAS_PRODUTOS on VW_TB_ORDENS_SERVICO.COD_ORDEM_SERVICO = VW_TB_PEDIDOS_COMPRAS_PRODUTOS.COD_ORDEM_SERVICO WHERE VW_TB_PEDIDOS_COMPRAS_PRODUTOS.status_pedido_compra <> 0 AND VW_TB_PEDIDOS_COMPRAS_PRODUTOS.status_produto = 'P'")
		
		While not rs.Eof
				
				cod_pedido_compra_produto = rs("cod_pedido_compra_produto")
				cod_ordem_servico_etapa = Cstr(""&rs("cod_ordem_servico_etapa"))
				cod_ordem_servico = Cstr(""&rs("cod_ordem_servico"))
				cod_pedido_compra = rs("cod_pedido_compra")
				
				If Cstr(""&Request.Form("cod_pedido_compra_produto_"&cod_pedido_compra_produto))  = "S" Then
					
					'//////
					'Insere na base de dados temporária para gerar o recibo futuramente
					objConexao.Execute("INSERT INTO TB_PEDIDOS_COMPRAS_PRODUTOS_TEMPORARIA (cod_pedido_compra_produto,cod_usuario) VALUES("&sqlTexto(cod_pedido_compra_produto)&","&sqlTexto(Session("cod_usuario"))&")")
					'//////
					
					
					'1) Baixa o item na compra
					objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS_PRODUTOS SET status_produto = 'E', quantidade_entregue = "&sqlNumero(rs("quantidade_solicitada"))&" WHERE cod_pedido_compra_produto = "&sqlNumero(cod_pedido_compra_produto)) 
					
					'2) Verifica se não há mais itens para serem baixados do pedido
					Set rs2=objConexao.Execute("SELECT cod_pedido_compra FROM TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra = "&sqlTexto(cod_pedido_compra)&" and status_produto = 'P'")
						
						If rs2.Eof Then
							'Finaliza o Pedido de Compra se não encontrar nenhum produto pendente
							objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS SET STATUS_PEDIDO_COMPRA = 3 WHERE cod_pedido_compra = "&sqlTexto(cod_pedido_compra))
														
						End If
						
					rs2.Close
					Set rs2=Nothing
					
					'3) Finaliza etapa da OS
					If cod_ordem_servico_etapa <> "" Then
						
						'Verifica a qtd_etapas
						Set rs3=objConexao.Execute("Select count(*) as qtd_etapas From TB_ORDENS_SERVICO_ETAPAS where cod_ordem_servico = "&cod_ordem_servico&" group by cod_ordem_servico")
													
							If not rs3.Eof Then
								qtd_etapas = formatNumber(0&rs3("qtd_etapas"),0)
							End If
													
						rs3.Close
						Set rs3=Nothing
						
						'Verifica qual o número da etapa atual
						Set rs3=objConexao.Execute("Select numero_etapa From VW_TB_ORDENS_SERVICO where cod_ordem_servico = "&cod_ordem_servico)
					
							If not rs3.Eof Then
								numero_etapa = formatNumber(0&rs3("numero_etapa"),0)
							End If
													
						rs3.Close
						Set rs3=Nothing

						If int(qtd_etapas) = int(numero_etapa) Then
						'Deverá finalizar OS, pois é a última etapa
							sql = "UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_termino = NOW() WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&sqlTexto(cod_ordem_servico_etapa)
							objConexao.Execute(sql)
							sql = "UPDATE TB_ORDENS_SERVICO SET status_ordem_servico = 2, data_finalizado = NOW(), cod_usuario_finalizou = "&Session("cod_usuario")&" WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)
							objConexao.Execute(sql)
						Else
						'Finaliza a etapa atual				
							sql = "UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_termino = NOW() WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&sqlTexto(cod_ordem_servico_etapa)		
							objConexao.Execute(sql)
							'Inicia a próxima Etapa
							Set rs3=objConexao.Execute("SELECT cod_ordem_servico_etapa FROM TB_ORDENS_SERVICO_ETAPAS WHERE cod_ordem_servico = "&cod_ordem_servico&" and numero_etapa IN(SELECT (numero_etapa+1) FROM VW_TB_ORDENS_SERVICO WHERE cod_ordem_servico = "&cod_ordem_servico&")")
												
									If not rs3.Eof Then
										sql = "UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_inicio = NOW() WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&rs3("cod_ordem_servico_etapa")
										objConexao.Execute(sql)
										sql = "UPDATE TB_ORDENS_SERVICO SET cod_ordem_servico_etapa_atual = "&rs3("cod_ordem_servico_etapa")&" WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)
										objConexao.Execute(sql)
									End If
												
							rs3.Close
							Set rs3=Nothing
										
						End If
						
					
					End If
					
					
				End If
				
		rs.MoveNext()
		Wend
		
	rs.Close
	Set rs=Nothing
	
	
	FecharConexao()
	
	Response.Redirect("ppv_compras_pedidos_produtos_solicitados.asp?codmsg=1")


End If
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


%>
<script type="text/javascript">
<%
'a = cod_pedido_compra (base64Encode)
'b = cod_pedido_compra_produto (base64Encode)
%>

function confirmaExclusaoEntradaErrada(b){
	if(YNconfirm('Tem certeza que deseja cancelar a entrada deste item?\nO item será retornado para antes como se não tivesse sido entregue.')){
		var caminho = 'ppv_compras_pedidos_ajax.asp?ts=cee&valor='+b;
		ajax_loadContent("containerProdutoPedidoCompra_"+b,caminho,false,false,false);
	}		
}
function CarregaAjax(div,ts,p0,p1,p2,p3){
	//Função padrao para utilização de ajax nesta página.
	var caminho = "ppv_compras_pedidos_ajax.asp?ts="+ts+"&valor="+p0+'&valor1='+p1+'&valor2='+p2+'&valor3='+p3;
	//alert(caminho);
	ajax_loadContent(div,caminho,false,false,false);	
}

function marcarLinhaPedidosCompra(selecionado,cod_pedido_compra_produto){
	
	if(selecionado){
		document.getElementById("linha_"+cod_pedido_compra_produto).style.backgroundColor = '#D6FED7';
	}else{
		document.getElementById("linha_"+cod_pedido_compra_produto).style.backgroundColor = '#FFFFFF';
	}
	
}

function confirmaEntrada(){
	if(YNconfirm('Tem certeza que deseja efetuar a entrada dos itens selecionados?')){
		document.FORMULARIO_PADRAO.submit();
	}		
}

</script>

<h2><%=titulo_pagina%></h2><br />
 <!-- #include file="includes/compras_sub_menu.asp" -->
<div class="subMenu" style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC">
<form action="<%=caminho_pagina%>" name="FORMULARIO_FILTRO" method="get">
<input name="" type="submit" style="display:none; visibility:hidden" />
<table width="483" border="0" cellpadding="0" cellspacing="0"><tr><td width="106" style="padding-left:10px;"><table width="483" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="32" rowspan="8"><span style="font-weight: bold">Filtrar</span></td>
    <td width="105" style="padding-left:10px; padding-top:5px;">N&ordm; de Pedido</td>
    <td width="105" style="padding-left:10px; padding-top:5px;">Contrato</td>
    <td style="padding-left:10px; padding-top:5px;">Ordenar por</td>
    <td width="124" style="padding-left:10px; padding-top:5px;">Status do Produto</td>
    <td colspan="2" style="padding-left:10px; padding-top:5px;">&nbsp;</td>
  </tr>
  <tr>
    <td width="105" style="padding-left:10px;"><input type='text' style="width:80px" class="inputTamanho1" value="<%=filtro_NUMERO_PEDIDO_COMPRAS%>" name='NUMERO_PEDIDO_COMPRAS' /></td>
    <td width="105" style="padding-left:10px;"><input type='text' class="inputTamanho1" style="width:80px" value="<%=filtro_COD_CONTRATO%>" name='COD_CONTRATO' /></td>
    <td style="padding-left:10px;"><select style="width:100%" name="ordernar_por" class="inputTamanho1">
      <option <%If Cstr(filtro_ordernar_por) = "1" Then%> selected="selected" <%End If%> value="1">Data de Entrega (Crescente)</option>
      <option <%If Cstr(filtro_ordernar_por) = "2" Then%> selected="selected" <%End If%> value="2">Formando</option>
      <option <%If Cstr(filtro_ordernar_por) = "3" Then%> selected="selected" <%End If%> value="3">Institui&ccedil;&atilde;o de Ensino</option>
    </select></td>
    <td width="124" style="padding-left:10px;"><select name="status_produto" style="width:100%" class="inputTamanho1">
      <option <%If Cstr(filtro_status_produto) = "" Then%> selected="selected" <%End If%> value="">Todos</option>
      <option <%If Cstr(filtro_status_produto) = "P" Then%> selected="selected" <%End If%> value="P">Pendentes</option>
      <option <%If Cstr(filtro_status_produto) = "E" Then%> selected="selected" <%End If%> value="E">Entregues</option>
      <option <%If Cstr(filtro_status_produto) = "C" Then%> selected="selected" <%End If%> value="C">Cancelados</option>
    </select></td>
    <td width="41" style="padding-left:10px;">&nbsp;</td>
    <td width="31" style="padding-left:10px;">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2"  style="padding-left:10px; padding-top:5px;">Respons&aacute;vel pelo pedido</td>
    <td colspan="2"  style="padding-left:10px; padding-top:5px;">Fornecedor</td>
    <td  style="padding-left:10px; padding-top:5px;">&nbsp;</td>
    <td  style="padding-left:10px; padding-top:5px;">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" style="padding-left:10px;"> <%If validacao_para_entrada_qualquer_pedido Then%>
    <select class="inputTamanho2" name="CADASTRO_COD_USUARIO">
      <option value="" <%If Cstr(filtro_cadastro_cod_usuario) = "" Then%> selected="selected" <%End If%>>Todos</option>
      <%                                          
                                            Set rs3=objConexao.Execute("SELECT COD_USUARIO, NOME FROM TB_ADM_USUARIOS WHERE COD_USUARIO IN(SELECT cadastro_cod_usuario FROM TB_PEDIDOS_COMPRAS WHERE status_pedido_compra > 0 GROUP BY cadastro_cod_usuario) ORDER BY NOME ASC")
 
 											If not rs3.Eof Then
 
                                            	While not rs3.Eof%>
      <option value="<%=rs3("COD_USUARIO")%>" <%If Cstr(filtro_cadastro_cod_usuario) = Cstr(rs3("cod_usuario")) Then%> selected="selected" <%End If%>><%=rs3("NOME")%></option>
      <%
                	                            rs3.MoveNext()
                    	                        Wend	
                                            
											End If
                                        
                                            rs3.Close
                                            Set rs3=Nothing
                                        %>
    </select>
    <%Else%>
    <strong style="font-size:14px;"><%=nomeUsuario(Session("cod_usuario"),true)%></strong>
    <%End If%></td>
    <td colspan="2" style="padding-left:10px;"><select  style="width:100%" class="inputTamanho2" name="COD_FORNECEDOR">
      <option value="" <%If Cstr(filtro_cod_fornecedor) = "" Then%> selected="selected" <%End If%>>Todos</option>
      <%                                          
                                            Set rs3=objConexao.Execute("SELECT cod_fornecedor, NOME_FANTASIA FROM TB_FORNECEDORES WHERE TB_FORNECEDORES.ATIVADO = 1 order by TB_FORNECEDORES.NOME_FANTASIA asc")
 
 											If not rs3.Eof Then
 
                                            	While not rs3.Eof%>
      <option value="<%=rs3("cod_fornecedor")%>" <%If Cstr(filtro_cod_fornecedor) = Cstr(rs3("cod_fornecedor")) Then%> selected="selected" <%End If%>><%=rs3("NOME_FANTASIA")%></option>
      <%
                	                            rs3.MoveNext()
                    	                        Wend	
                                            
											End If
                                        
                                            rs3.Close
                                            Set rs3=Nothing
                                        %>
    </select></td>
    <td style="padding-left:10px;">&nbsp;</td>
    <td style="padding-left:10px;">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" style="padding-left:10px; padding-top:5px;">Data Entrega</td>
    <td colspan="4" style="padding-left:10px; padding-top:5px;"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="32%">&nbsp;</td>
        <td width="68%">Produto</td>
        </tr>
    </table></td>
    </tr>
  <tr>
    <td colspan="6" style="padding-left:10px;"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="57%"><table width="300" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="135"><table width='135' border='0' align="left" cellpadding='0' cellspacing='0'>
              <tr>
                <td width='102'><input class='campoData' maxlength='10' onkeydown='mascara(this,data);' onkeypress='mascara(this,data);' onkeyup='mascara(this,data);' type='text' name='DATA_PREVISAO_ENTREGA_INICIO' value="<%=ArrumaData(b_DATA_PREVISAO_ENTREGA_INICIO)%>" onblur='validarData(this)' /></td>
                <td width='298'><div align='center'> <img onclick='pickDate(this,document.FORMULARIO_FILTRO.DATA_PREVISAO_ENTREGA_INICIO);' class='forcaBotao' src='images/icons/icon_calendario.jpg' width='26' height='22' /></div></td>
              </tr>
            </table></td>
            <td width="11"><div align="center">&agrave;</div></td>
            <td width="154"><table width='135' border='0' align="left" cellpadding='0' cellspacing='0'>
              <tr>
                <td width='102'><input class='campoData' maxlength='10' onkeydown='mascara(this,data);' onkeypress='mascara(this,data);' onkeyup='mascara(this,data);' type='text' name='DATA_PREVISAO_ENTREGA_FIM' value="<%=ArrumaData(b_DATA_PREVISAO_ENTREGA_FIM)%>" onblur='validarData(this)' /></td>
                <td width='298'><div align='center'> <img onclick='pickDate(this,document.FORMULARIO_FILTRO.DATA_PREVISAO_ENTREGA_FIM);' class='forcaBotao' src='images/icons/icon_calendario.jpg' width='26' height='22' /></div></td>
              </tr>
            </table></td>
          </tr>
        </table></td>
        <td width="43%"><select name='COD_PRODUTO' class='inputTamanho2'>
          <option value=''>Selecione</option>
          <%
Set rsE3=objConexao.Execute("SELECT COD_PRODUTO, NOME_PRODUTO FROM TB_PRODUTOS WHERE ATIVO = 1 AND gera_ordem_servico_prazo_cliente > 0 ORDER BY NOME_PRODUTO")
While not rsE3.Eof 
'Padrao
Dim selRelPadrao3, verificador3
verificador3 = ""&rsE3(0)

If CStr(verificador3) = CStr(filtro_COD_PRODUTO)  Then%>
          <option selected='selected' value="<%=rsE3(0)%>"><%=rsE3(1)%></option>
          <%Else%>
          <option value="<%=rsE3(0)%>"><%=rsE3(1)%></option>
          <%End if%>
          <%rsE3.MoveNext()
Wend
rsE3.Close
Set rsE3=Nothing
%>
        </select></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="6" style="padding-left:10px;"><div align="left" style="padding-top:10px">
      <input class="btnFiltrarAmarelo" type="submit" name="filtro1" value=" " />
    </div></td>
  </tr>
  </table>  <br /></td>
    <td width="106" style="padding-left:10px;">&nbsp;</td>
  </tr>
</table>
<input type="hidden" value="1" name="filtro" />
</form>
</div>

<%
'/////////////////////////////////////////////////////////////////////////////////
	'Verifica se o último recibo está disponível
	Dim ultimo_recibo_disponivel
	ultimo_recibo_disponivel = false
	
		Set rs2=objConexao.Execute("SELECT * FROM VW_TB_PEDIDOS_COMPRAS_PRODUTOS_TEMPORARIA WHERE cod_usuario_temporario = "&Session("cod_usuario")&" LIMIT 1")
			
			If not rs2.Eof Then
				ultimo_recibo_disponivel = true
			End If
			
		rs2.Close
		Set rs2=Nothing
	
		If ultimo_recibo_disponivel and Request.QueryString("codmsg") = Empty Then
		'Exibe somente se não tiver exibido a mensagem original de impressão.
			Response.Write("<div class='msgGrande'>Último recibo de entrega de fornecedor disponível para impressão: <a target='_blank' href='ppv_compras_pedidos_produtos_solicitados_recibo_entrega.asp'><strong>Clique aqui</strong></a></div>")
		End If
	
	'/////////////////////////////////////////////////////////////////////////////////
	
	Dim sucesso, codmsg
	codmsg	= CStr(""&Request.QueryString("codmsg"))
	sucesso = CStr(""&Request.QueryString("sucesso"))
		
			
	If sucesso = "" and codmsg <> "" Then%>
		<div align="center" class="msgGrandeSucesso">
		<%
			If codmsg = "1" Then
				Response.Write("Entrada de itens realizada com sucesso!<br />Para imprimir o recibo de entrega de fornecedor <a target='_blank' href='ppv_compras_pedidos_produtos_solicitados_recibo_entrega.asp'><strong>Clique aqui</strong></a>")			
			End If	
				%>           
				
	</div>
		<%ElseIf sucesso <> "" and codmsg <> "" Then%>
			<div align="center" class="msgGrandeErro">
	</div>
	<%End If%>
    
    <%
	If Request.QueryString("cod_fornecedor") <> Empty or Request.QueryString("NUMERO_PEDIDO_COMPRAS") <> Empty Then
	%>
    <%

	'Response.Write(sql)
	Set rs=server.CreateObject("ADODB.recordset")
	rs.CursorLocation=3  'Cursor no cliente
	rs.PageSize=200
	rs.Open sql, objConexao
	Dim paginacaocao
	paginacaocao = formatNumber(0&Request.QueryString("paginacao"),0)
	
	If not rs.Eof Then
	
		If paginacaocao > 0 Then
			rs.AbsolutePage=paginacaocao
		Else
			paginacaocao = 1
			rs.AbsolutePage=1
		End If	
%>

<table width="99%" border="0" align="center">
  <tr>
    <td width="47%"> 
Página <strong><%=rs.AbsolutePage%></strong> de <%=rs.PageCount%>
</td>
    <td width="53%"><div align="right">Total <strong><%=rs.RecordCount%></strong> Itens de Pedidos de Compras</div></td>
  </tr>
</table>
<form action="ppv_compras_pedidos_produtos_solicitados.asp" name="FORMULARIO_PADRAO" style="border:0px" method="post">
<input type="hidden" name="campo1" value="S" />
<table width="99%" border="1" align="center" cellpadding="5" style="border-collapse:collapse; text-align: right;">
<tr bgcolor="whitesmoke">
<td><div align="center" style="font-weight: bold">Efetuar Entrada</div></td>
<td colspan="2"><div align="center" style="font-weight: bold">Pedido Compra</div></td>
<td width="283"><div align="center" style="font-weight: bold">Fornecedor</div></td>
<td width="184"><div align="center"><strong>OS/Resp.</strong></div></td>
<td width="131"><div align="center"><strong>Formando</strong></div></td>
<td width="96"><div align="center"><strong>Institui&ccedil;&atilde;o de Ensino</strong></div></td>
<td width="187"><div align="center" style="font-weight: bold">
  <div align="center">Produto</div>
</div></td>
<td width="45"><div align="center" style="font-weight: bold">
  <div align="center">Qtd.</div>
</div></td>
<td width="88" align="center"><div align="center"><span style="font-weight: bold">Data Entrega</span></div></td>
<td width="46" align="center"><strong>Obs</strong></td>
              <td width="80"><div align="center"><strong>Status</strong></div></td>
</tr>
<%
Dim dias_atrasados, acesso_cancelamento
acesso_cancelamento = validaUsuario(2,,65,true)


Dim usuario_pode_imprimir
usuario_pode_imprimir = validaUsuario(2,,64,true)

Dim x
For x=1 To rs.PageSize
If Not rs.Eof Then
	status_produto = Cstr(""&rs("status_produto"))
	dias_atrasados = formatNumber(rs("dias_atrasados"))
%>
<tr id="linha_<%=rs("cod_pedido_compra_produto")%>">
  <td width="72"><div align="center">
    <%If status_produto = "P" and acesso_entrada_pedido_compra Then%>
    <input type="checkbox" onclick="marcarLinhaPedidosCompra(this.checked,'<%=rs("cod_pedido_compra_produto")%>')" name="cod_pedido_compra_produto_<%=rs("cod_pedido_compra_produto")%>" value="S" class="inputRadio" />
    <%Else%>
    -
  <%End If%>
  </div></td>
  <td width="31"><div align="center"><%=formatNumber(rs("numero_pedido_compra"),0)%></div></td>
  <td width="30"><div align="center">
  
  	<%If usuario_pode_imprimir Then%>
  	<a href="ppv_compras_pedidos_imprimir.asp?c=<%=rs("cod_pedido_compra")%>" target="_blank"><img src="images/icons/icon_imprimir.png" alt="" title="" border="0" /></a>
    <%Else%>
    -
    <%End If%>
    </div></td>
  <td><div align="left"><%=rs("razao_social")%></div></td>
  <td><div align="center"><%=moduloPPVOSMascara(rs("cod_ordem_servico"),"",true)%><%
	Set rs2=objConexao.Execute("SELECT cod_usuario_responsavel FROM  VW_TB_ORDENS_SERVICO WHERE cod_ordem_servico = "&sqlTexto(rs("cod_ordem_servico")))
		
		If not rs2.Eof Then
			Response.Write("<br />"&ucase(nomeUsuario(rs2("cod_usuario_responsavel"),true)))
		End If
		
	rs2.Close
	Set rs2=Nothing
	%>   
    </div></td>
  <td><div align="left"><%=rs("nome_formando")%></div></td>
  <td><div align="left"><span style="font-size:12px"><%=Ucase(rs("ENTIDADE_ENSINO"))%></span></div></td>
  <td><div align="left"><%=rs("nome_produto")%></div></td>
  <td><div align="center"><%=ArrumaNumero(rs("quantidade_solicitada"),"-",0)%></div></td>
  <td> <div align="center"><%=ArrumaData(rs("data_previsao_entrega"))%></div></td>
  <td><div align="left"><%=rs("obs")%></div></td>
<td align="center"  style="font-size:11px">
  <div id="containerProdutoPedidoCompra_<%=base64Encode(rs("cod_pedido_compra_produto"))%>">
    <%If status_produto = "E" Then%>
    <div class="contasPagarReceberStatusQ" style="text-decoration:none; cursor:auto;">Entregue</div>
    <%If acesso_excluir_entrada_errada Then%><a href="javascript:confirmaExclusaoEntradaErrada('<%=base64Encode(rs("cod_pedido_compra_produto"))%>')">Excluir entrada</a><%End If%>
    
    
    <%Elseif status_produto = "P" and dias_atrasados = 0 Then%>    
    <div class="contasPagarReceberStatusA_nodia" style="text-decoration:none; cursor:auto;">Pendente</div>
    <%Elseif status_produto = "P" and dias_atrasados < 1 Then%>    
    <div class="contasPagarReceberStatusA_nao_vencido" style="text-decoration:none; cursor:auto;">Pendente</div>
    <%Elseif status_produto = "P" and dias_atrasados > 0 Then%>
    <div class="contasPagarReceberStatusA_vencido" style="text-decoration:none; cursor:auto;">Pendente</div>
    <%Elseif status_produto = "C" Then%>
    <div class="contasPagarReceberStatusA_vencido" style="text-decoration:none; background-color:#999999; cursor:auto;">Cancelado</div>
    <%End If%>
  </div>    </td>
  </tr>

<%
rs.MoveNext()
End If
Next
%>
<%If acesso_entrada_pedido_compra Then%>
<tr bgcolor="whitesmoke">
      <td colspan="12">
        <div align="left">
          <table cellpadding="5" cellspacing="0">
            <tr>
              <th scope="col" width="22"><img src="images/icons/icon_setas.gif" border="0" height="15" width="22"></th>
              <th scope="col" width="366"><div align="left"><a href="javascript:confirmaEntrada()"><strong>Dar entrada nos itens selecionados acima</strong></a></div></th>
              </tr>
          </table>
        </div></td>
    </tr>
<%End If%>
<%If rs.PageCount > 1 Then %>                                  
     <tr bgcolor="whitesmoke">
      <td colspan="12">
      	   <%call paginacao(paginacaocao,rs.PageCount)%>      </td>
    </tr>   
<%End if%>

</table>
</form>

<br />
<p><strong>&nbsp;Legenda:</strong></p>
<table width="612" border="0" cellpadding="4" cellspacing="0">
  
  <tr>
    <td width="25"><div class="contasPagarReceberStatusA_nao_vencido" style="text-decoration:none; cursor:auto;">Pendente</div></td>
    <td width="579">Item pendente, por&eacute;m com o prazo em dia</td>
  </tr>
  <tr>
    <td><div class="contasPagarReceberStatusA_nodia" style="text-decoration:none; cursor:auto;">Pendente</div></td>
    <td>Item pendente, por&eacute;m com o prazo vencendo na data atual</td>
  </tr>
  <tr>
    <td><div class="contasPagarReceberStatusA_vencido" style="text-decoration:none; cursor:auto;">Pendente</div></td>
    <td>Item pendente, por&eacute;m com o prazo atrasado</td>
  </tr>
  <tr>
    <td><div class="contasPagarReceberStatusQ" style="text-decoration:none; cursor:auto;">Entregue</div></td>
    <td>Item entregue</td>
  </tr>
  <tr>
    <td><div class="contasPagarReceberStatusA_vencido" style="text-decoration:none; background-color:#999999; cursor:auto;">Cancelado</div></td>
    <td>item cancelado</td>
  </tr>
</table>

<%
	Else
	%>
    <div align="center" class="msgGrande">
    	Nenhum Item de Pedido de Compras encontrado!</div>
    <%End if
	
	rs.Close
	Set rs=Nothing
	
	FecharConexao()
	%>
    <%
	Else
	%>
    	<div class="msgGrande">Utilize o filtro para listar os itens, pelo menos filtre por fornecedor ou por n&uacute;mero do pedido de compras.</div>
    <%
	End If
	%>
    <div id="tsZXucAdca" class="ocultarCompleto"></div>
    <br />
	<br />
	<br />
<!-- #include file="files/includes/miolo-inferior.asp" -->
</body>
</html>