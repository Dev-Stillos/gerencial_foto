<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<%
Response.Buffer = False
Dim efetuar_transacao
efetuar_transacao = Cstr(""&Request.Form("campo1"))	

Dim busca_cod_produto

Dim sql, sqlWhere, sqlOrderBy, sqlGroupBy

'Padr�o Where
sqlWhere = sqlWhere&" and status_ordem_servico = '1' and necessita_gerar_pedido_compra = 'S' and pedido_compra_gerado IS NULL"

'Padr�o Order By 
sqlOrderBy = "cod_contrato, nome_formando, numero_ordem_servico desc, numero_ordem_servico_indice asc"

If Request.QueryString("filtro1") <> Empty Then
		
	Dim filtro_COD_CONTRATO
	Dim filtro_NOME_FORMANDO
	Dim filtro_COD_PRODUTO
	Dim filtro_ORDERNAR_POR
	
	filtro_COD_CONTRATO 	= Cstr(""&Request.QueryString("COD_CONTRATO"))
	filtro_NOME_FORMANDO 	= Cstr(""&Request.QueryString("NOME_FORMANDO"))
	filtro_COD_PRODUTO  	= Cstr(""&Request.QueryString("COD_PRODUTO"))
	filtro_ORDERNAR_POR  	= Cstr(""&Request.QueryString("ORDERNAR_POR"))
	
		If filtro_ordernar_por <> "" Then
			If filtro_ordernar_por = "2" Then
				sqlOrderBy = "nome_formando, numero_ordem_servico desc, numero_ordem_servico_indice asc"
			ElseIf filtro_ordernar_por = "3" Then				
				sqlOrderBy = "cod_contrato, nome_formando, numero_ordem_servico desc, numero_ordem_servico_indice asc"
			End If	
			
		End If
	
	If filtro_COD_PRODUTO <> "" Then
		sqlWhere = sqlWhere&" and cod_produto = "&filtro_COD_PRODUTO
	End If
	
	If filtro_COD_CONTRATO <> "" Then
			sqlWhere = sqlWhere&" AND COD_CONTRATO = '"&filtro_COD_CONTRATO&"'"
	End If
			
	If filtro_NOME_FORMANDO <> "" Then
			sqlWhere = sqlWhere&" AND NOME_FORMANDO LIKE '%"&filtro_NOME_FORMANDO&"%'"
	End If	
			

End If
'/////////////////////
'Padroniza sqlWhere

	If sqlWhere <> "" Then
		'Retirar o "AND "
		sqlWhere = Mid(sqlWhere,5,1000)
	End If


'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
If efetuar_transacao = "S" Then
	
	AbrirConexao()
		
		Dim cod_fornecedor, ultima_etapa, cod_ordem_servico, cod_ordem_servico_etapa
		
		'//////
		'Limpa base tempor�ria
		objConexao.Execute("DELETE FROM TB_PEDIDOS_COMPRAS_TEMPORARIA WHERE cod_usuario = "&sqlTexto(Session("cod_usuario")))	
		'//////		

		'1� ETAPA
		'/////////////////////////////////////////////////////////////////////		
		'1) Selecionar todas OS para triagem
		Set rs=objConexao.Execute("SELECT cod_produto, nome_produto, sum(quantidade) as qtd_necessario_comprar, nome_produto FROM VW_TB_ORDENS_SERVICO WHERE "&sqlWhere&" GROUP BY cod_produto order by nome_produto asc")
			If not rs.Eof then
				While not rs.Eof
						
					cod_produto = rs("cod_produto")				
					cod_fornecedor = Cstr(""&Request.Form("cod_fornecedor_"&cod_produto))
					
					'Response.Write(rs("nome_produto")&"<br />")
					
					'Verifica se foi selecionado algum fornecedor
					If cod_fornecedor <> "" Then
						
						sql = "SELECT * FROM VW_TB_ORDENS_SERVICO WHERE cod_produto = "&cod_produto&" and "&sqlWhere&" order by numero_ordem_servico desc, numero_ordem_servico_indice asc" 						
						'Seleciona todas as OS do produto atual
						'Response.Write(sql)
						Set rs2=objConexao.Execute(sql)
							
							If not rs2.Eof Then
								
								While not rs2.Eof
										
										cod_ordem_servico = rs2("cod_ordem_servico")
										ordem_servico_selecionada = Cstr(""&Request.Form("cod_ordem_servico_"&cod_ordem_servico))
										quantidade = rs2("quantidade")
										
										If ordem_servico_selecionada = "S" Then
											
											'Insere na base tempor�ria
											
												'Verifica se � �ltima etapa
												Set rs3=objConexao.Execute("Select count(*) as qtd_etapas From TB_ORDENS_SERVICO_ETAPAS where cod_ordem_servico = "&cod_ordem_servico&" group by cod_ordem_servico")
													
													If not rs3.Eof Then
														qtd_etapas = formatNumber(0&rs3("qtd_etapas"),0)
													End If
													
												rs3.Close
												Set rs3=Nothing
												
												if int(qtd_etapas) = int(rs2("numero_etapa")) Then
													ultima_etapa = "S"
												Else
													ultima_etapa = "N"
												End If
											
												sql = "INSERT INTO TB_PEDIDOS_COMPRAS_TEMPORARIA (cod_ordem_servico, cod_produto, cod_fornecedor, quantidade, cod_usuario,prazo_em_dias,cod_ordem_servico_etapa,ultima_etapa) VALUES ("&sqlTexto(cod_ordem_servico)&","&sqlTexto(cod_produto)&","&sqlTexto(cod_fornecedor)&","&sqlTexto(quantidade)&","&sqlTexto(Session("cod_usuario"))&","&rs2("prazo_em_dias")&","&rs2("cod_ordem_servico_etapa_atual")&","&sqlTexto(ultima_etapa)&")"
												
												objConexao.Execute(sql)
														
										End If
										
								rs2.MoveNext()
								Wend
								
							End If
							
						rs2.Close
						Set rs2=Nothing
											
					End If
									
				rs.MoveNext()
				Wend
				
			End if
			
		rs.close
		Set rs=nothing
		
		'2� ETAPA
		'/////////////////////////////////////////////////////////////////////
			
			Dim cod_pedido_compra, valor_custo_solicitado, data_previsao_entrega
			cod_pedido_compra = ""
			
			'Lista os fornecedores agrupados
			Set rs=objConexao.Execute("SELECT cod_fornecedor FROM TB_PEDIDOS_COMPRAS_TEMPORARIA WHERE cod_usuario = "&sqlTexto(Session("cod_usuario"))&" GROUP BY cod_fornecedor")
				If not rs.Eof then
				
					While not rs.Eof
					
						cod_fornecedor = rs("cod_fornecedor")		
						
						'1) Faz o insert do Pedido de Compras
						'//////
						objConexao.Execute("INSERT INTO TB_PEDIDOS_COMPRAS (cod_fornecedor,cadastro_cod_usuario,data_aprovado,aprovado_cod_usuario,status_pedido_compra,cod_usuario,modificado) VALUES ("&sqlTexto(cod_fornecedor)&","&sqlTexto(Session("cod_usuario"))&",NOW(),"&sqlTexto(Session("cod_usuario"))&",2,"&sqlTexto(Session("cod_usuario"))&",NOW())")
						'Obs: Status 2 = aprovado e n�o enviado para fornecedor
						'Busca a chave criada
						Set rs2=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_PEDIDOS_COMPRAS LIMIT 1")
							If not rs2.Eof Then
								cod_pedido_compra = rs2("ULTIMO_ID")
							End if 
						rs2.Close
						Set rs2=Nothing																
						
						'Atribui o n�mero do pedido de compras
						Set rs2=objConexao.Execute("SELECT IFNULL(MAX(numero_pedido_compra),0)+1 as numero_pedido_compra FROM TB_PEDIDOS_COMPRAS WHERE status_pedido_compra <> 0")
							If not rs2.Eof Then
								objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS SET numero_pedido_compra = "&sqlNumero(rs2(0))&" WHERE TB_PEDIDOS_COMPRAS.cod_pedido_compra ="&sqlTexto(cod_pedido_compra))
							End if 
						rs2.Close
						Set rs2=Nothing		
						'//////
						
						
						'###////////////////////////////////////////
						'### - N�O AGRUPA MAIS OS PRODUTOS
'						'2) Agrupa os produtos do fornededor
'						sql = "SELECT cod_produto, sum(quantidade) as quantidade, prazo_em_dias FROM TB_PEDIDOS_COMPRAS_TEMPORARIA WHERE cod_usuario = "&sqlTexto(Session("cod_usuario"))&" and cod_fornecedor = "&cod_fornecedor&" GROUP BY cod_produto"
'						'Response.Write(sql)
'						Set rs2=objConexao.Execute(sql)				
'								While not rs2.Eof
'									
'									cod_produto = rs2("cod_produto")
'									'3) Insere no pedido de compra produto
'										
'										'///
'										'Valor de custo do produto do fornecedor
'										valor_custo_solicitado  = 0
'										Set rs3=objConexao.Execute("SELECT valor_custo FROM TB_FORNECEDOR_PRODUTOS WHERE cod_fornecedor = "&sqlTexto(cod_fornecedor)&" AND cod_produto = "&sqlTexto(cod_produto))
'											
'											If not rs3.Eof Then
'												valor_custo_solicitado = rs3(0)
'											End If
'										
'										rs3.Close	
'										Set rs3=Nothing
'																			
'										'Previs�o de Entrega 
'										data_previsao_entrega = DateAdd("d",rs2("prazo_em_dias"),date())
'										'///
'									
'										objConexao.Execute("INSERT INTO TB_PEDIDOS_COMPRAS_PRODUTOS (cod_pedido_compra, cod_produto, quantidade_solicitada, valor_custo_solicitado, data_previsao_entrega, status_produto, cod_usuario, modificado) VALUES ("&sqlTexto(cod_pedido_compra)&",'"&cod_produto&"',"&sqlNumero(rs2("quantidade"))&","&sqlNumero(valor_custo_solicitado)&","&sqlData(data_previsao_entrega)&",'P',"&sqlTexto(Session("cod_usuario"))&",NOW())")
'									
'								rs2.MoveNext()
'								Wend
'						
'						rs2.Close
'						Set rs2=Nothing
						
						'###////////////////////////////////////////
						
						'4) Listar todos os produtos do fornecedor SEM Group By
						Set rs2=objConexao.Execute("SELECT * FROM TB_PEDIDOS_COMPRAS_TEMPORARIA WHERE cod_usuario = "&sqlTexto(Session("cod_usuario"))&" and cod_fornecedor = "&cod_fornecedor)				
								While not rs2.Eof
									
									cod_ordem_servico = rs2("cod_ordem_servico")
									cod_ordem_servico_etapa = rs2("cod_ordem_servico_etapa")									
									cod_produto = rs2("cod_produto")
										
'										'///
'										'Valor de custo do produto do fornecedor
										valor_custo_solicitado  = 0
										Set rs3=objConexao.Execute("SELECT valor_custo FROM TB_FORNECEDOR_PRODUTOS WHERE cod_fornecedor = "&sqlTexto(cod_fornecedor)&" AND cod_produto = "&sqlTexto(cod_produto))
											
											If not rs3.Eof Then
												valor_custo_solicitado = rs3(0)
											End If
										
										rs3.Close	
										Set rs3=Nothing
'																			
'										'Previs�o de Entrega 
										data_previsao_entrega = DateAdd("d",rs2("prazo_em_dias"),date())
										'///
									
										objConexao.Execute("INSERT INTO TB_PEDIDOS_COMPRAS_PRODUTOS (cod_pedido_compra, cod_produto, quantidade_solicitada, valor_custo_solicitado, data_previsao_entrega, cod_usuario, modificado,cod_ordem_servico_etapa) VALUES ("&sqlTexto(cod_pedido_compra)&",'"&cod_produto&"',"&sqlNumero(rs2("quantidade"))&","&sqlNumero(valor_custo_solicitado)&","&sqlData(data_previsao_entrega)&","&sqlTexto(Session("cod_usuario"))&",NOW(),"&sqlTexto(cod_ordem_servico_etapa)&")")
									
										objConexao.Execute("UPDATE TB_ORDENS_SERVICO_ETAPAS SET pedido_compra_gerado = "&sqlNumero(cod_pedido_compra)&" WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&sqlTexto(cod_ordem_servico_etapa))
									
									'////
									'N�O FINALIZA OS, SOMENTE QUANDO FOR ENTREGAR
									
									If Cstr(rs2("ultima_etapa")) = "S" Then
									'Dever� finalizar OS
											'sql = "UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_termino = NOW() WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&sqlTexto(cod_ordem_servico_etapa)
											'objConexao.Execute(sql)
											'sql = "UPDATE TB_ORDENS_SERVICO SET status_ordem_servico = 2, data_finalizado = NOW(), cod_usuario_finalizou = "&Session("cod_usuario")&" WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)
											'objConexao.Execute(sql)
									Else
										
											'Finaliza a etapa atual
											'sql = "UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_termino = NOW() WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&sqlTexto(cod_ordem_servico_etapa)		
											'objConexao.Execute(sql)
											'Inicia a pr�xima Etapa
											'Set rs3=objConexao.Execute("SELECT cod_ordem_servico_etapa FROM TB_ORDENS_SERVICO_ETAPAS WHERE cod_ordem_servico = "&cod_ordem_servico&" and numero_etapa IN(SELECT (numero_etapa+1) FROM VW_TB_ORDENS_SERVICO WHERE cod_ordem_servico = "&cod_ordem_servico&")")
												
											'	If not rs3.Eof Then
											'		sql = "UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_inicio = NOW() WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&rs3("cod_ordem_servico_etapa")
											'		objConexao.Execute(sql)
											'		sql = "UPDATE TB_ORDENS_SERVICO SET cod_ordem_servico_etapa_atual = "&rs3("cod_ordem_servico_etapa")&" WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)
											'		objConexao.Execute(sql)
											'	End If
												
											'rs3.Close
											'Set rs3=Nothing
										
									End If		
									
									'///							
									
								rs2.MoveNext()
								Wend
						
						rs2.Close
						Set rs2=Nothing						
			
					rs.MoveNext()
					Wend
				
				End if
			
			rs.close
			Set rs=nothing		
		
	FecharConexao()
	Response.Redirect("ppv_compras_pedidos.asp?codmsg=601")
	
End If
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
<!-- #include virtual="gerencial_foto/files/includes/scripts.asp" -->
<script language="javascript">
	function alterarQtdSolicitada(cod_produto,quantidade,campo,cod_ordem_servico){
		//var temp;
		//temp = new Number(document.FORMULARIO_PADRAO["qtd_solicitada_"+cod_produto].value);
		//var checked = campo.checked;
		var quantidade_total = 0, quantidade_total_nao_marcada = 0; qtd = 0;
		
		/*
		$.each($(".quantidade_por_os_"+cod_produto), function(index,value) {
			quantidade_total += parseInt($(value).val());
		});
		*/
		
		$.each($(".inputs_por_os_"+cod_produto), function(index,value) {
			if($(value).is(":checked")){
				qtd = parseInt($(value).attr("quantidade"));
				quantidade_total += qtd;
			}
			
		});		
							
		if(campo.checked){
			document.getElementById("linha_"+cod_ordem_servico).style.backgroundColor = '#D6FED7';
			//temp = temp+quantidade;
		}else{
			document.getElementById("linha_"+cod_ordem_servico).style.backgroundColor = '#FFFFFF';
			//temp = temp-quantidade;
		}
		
		document.getElementById("containnerSolicitar_"+cod_produto).innerHTML = quantidade_total;
		document.FORMULARIO_PADRAO["qtd_solicitada_"+cod_produto].value = quantidade_total;
	}
	
	function habilitaCampo(campo){
		
	}

	function confirmaTriagem(){
		if(YNconfirm('Tem certeza que deseja criar os pedidos de compra para as Ordens de Servi�o selecionadas?')){
			document.FORMULARIO_PADRAO.submit();
		}
	}
</script>  
</head>
<body>

<table cellpadding="9px" cellspacing="0" width="100%">
<tr>
<td>
<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td>
		<!-- #include file="includes/topo.asp" -->
<table width="100%" style="border-bottom:#999999 solid dotted 1px" background="images/layout/fundo_menu_azul.jpg"border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><!-- #include file="includes/menu.asp" --></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="10">
  <tr>
  <td>
   <%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 2
  BTN_FONTE = 1
  BTN_IMPRIMIR_LINK = "#"
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr> 
    <td valign="top" style="border:dotted 1px #999999">
	
	  <h2>PPV - Gera&ccedil;&atilde;o de Pedidos de Compras com Origem em Ordens de Servi&ccedil;o</h2>
         <br />
      <!-- include file="includes/classe_modulo_ppv_menu.asp" -->
	  <br />
   
      <form action="ppv_ordens_servico_facilitador_geracao_pedidos_compra.asp" style="margin:0px">
      	<input name="filtro1" value="1" type="hidden" />
	  <table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="99%" border="0" align="center" cellpadding="5" cellspacing="0">
		  <tr>
		    <td width="100%"><table width="483" border="0" cellpadding="2" cellspacing="0">
		      <tr>
		        <td width="32" rowspan="6"><span style="font-weight: bold">Filtrar</span></td>
		        <td width="210" style="padding-left:10px; padding-top:5px;">Produto</td>
		        <td style="padding-left:10px; padding-top:5px;">C&oacute;d. Contrato</td>
		        </tr>
		      <tr>
		        <td width="210" style="padding-left:10px;"><select name='COD_PRODUTO' class='inputTamanho2'>
		          <option value=''>Selecione</option>
		          <%
				  AbrirConexao()
				  
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

FecharConexao()
%>
		          </select></td>
		        <td style="padding-left:10px;"><input type='text' class="inputTamanho1" style="width:90%" value="<%=filtro_COD_CONTRATO%>" name='COD_CONTRATO' /></td>
		        </tr>
		      <tr>
		        <td  style="padding-left:10px; padding-top:5px;">Ordenar por</td>
		        <td  style="padding-left:10px; padding-top:5px;">Nome do Formando</td>
		        </tr>
		      <tr>
		        <td style="padding-left:10px;"><select style="width:100%" name="ordernar_por" class="inputTamanho1">
		          <option <%If Cstr(filtro_ordernar_por) = "2" Then%> selected="selected" <%End If%> value="2">Formando</option>
		          <option <%If Cstr(filtro_ordernar_por) = "3" or Cstr(filtro_ordernar_por) = "" Then%> selected="selected" <%End If%> value="3">Institui&ccedil;&atilde;o de Ensino</option>
		          </select></td>
		        <td style="padding-left:10px;"><input type='text' class="inputTamanho1" style="width:90%" value="<%=filtro_NOME_FORMANDO%>" name='NOME_FORMANDO' /></td>
		        </tr>
		      <tr>
		        <td colspan="2" style="padding-left:10px; padding-top:5px;"><div align="left" style="padding-top:10px">
		          <input class="btnFiltrarAmarelo" type="submit" name="filtro2" value=" " />
		          </div></td>
		        </tr>
		      </table></td>
		    </tr>
        </table>
      </form>
        <br />

        <em><strong>Listagem dos produtos cujas ordens de servi&ccedil;o est&atilde;o com a etapa atual marcada como: &quot;Necessita Gerar Pedido de Compra&quot; e que ainda n&atilde;o foram solicitados.</strong></em><br />
        <br />
        <%		
		Dim temp
		'Neste momento corrige o sqlWhere somente para o primeiro rs
		If sqlWhere <> "" Then
			temp = "WHERE "
		End If
		
		sql = "SELECT cod_produto, nome_produto, sum(quantidade) as qtd_necessario_comprar FROM VW_TB_ORDENS_SERVICO "&temp&sqlWhere&" GROUP BY cod_produto order by nome_produto asc"
		
		'Response.Write(sql)
		
		AbrirConexao()
		
		Set rs=objConexao.Execute(sql)
		
		If not rs.Eof Then
		%>
        <div class="msgGrande">Ap&oacute;s a gera&ccedil;&atilde;o da compra a etapa que est&aacute; pendente ser&aacute; finalizada das Ordens de Servi&ccedil;os marcadas.</div>
        <form name="FORMULARIO_PADRAO" action="ppv_ordens_servico_facilitador_geracao_pedidos_compra.asp" method="post" style="margin:0px">
        	<input name="campo1" value="S" type="hidden" />
            <table width="99%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
              <tr bgcolor="whitesmoke">
        <td><div align="center"><strong> Produto</strong></div></td>
        <td width="18%" colspan="-1"><div align="center"><strong>Qtd. Total</strong></div></td>
        <td width="16%" colspan="-1"><div align="center"><strong>Qtd. Solicitada Agora</strong></div></td>
        <td width="30%" colspan="-1"><div align="center"><strong>Fornecedor/Valor Custo Un.</strong></div></td>
        </tr>
      <%
    
    'Neste momento que j� passou o group by corrige o sqlWhere para o rs2
    If sqlWhere <> "" Then
        sqlWhere = " AND "&sqlWhere
    End If
    
    While not rs.Eof
    %>
 
    <tr>
                                <td class="textoMedio"><strong><%=Ucase(""&rs("nome_produto"))%></strong></td>
                                <td colspan="-1" class="textoMedio"><div align="center"><strong><%=rs("qtd_necessario_comprar")%></strong></div></td>
                                <td colspan="-1" class="textoMedio"><div style="font-weight:bold" align="center" id="containnerSolicitar_<%=rs("cod_produto")%>">0</div><input type="hidden" value="0" id="id_qtd_solicitada_<%=rs("cod_produto")%>" name="qtd_solicitada_<%=rs("cod_produto")%>" /></td>
                                <td colspan="-1"><div align="center">
                                  <select class="inputTamanho2" name="cod_fornecedor_<%=rs("cod_produto")%>">
                                    <%
                                            'Seleciona todos os funcion&aacute;rios ativos, exceto o usu&aacute;rio atual.
                                            Set rs3=objConexao.Execute("SELECT TB_FORNECEDORES.NOME_FANTASIA, TB_FORNECEDOR_PRODUTOS.* FROM  TB_FORNECEDOR_PRODUTOS INNER JOIN TB_FORNECEDORES ON  TB_FORNECEDOR_PRODUTOS.cod_fornecedor = TB_FORNECEDORES.cod_fornecedor WHERE TB_FORNECEDORES.ATIVADO = 1 and TB_FORNECEDOR_PRODUTOS.cod_produto = "&sqlTexto(rs("cod_produto"))&" order by TB_FORNECEDOR_PRODUTOS.valor_custo asc")
 
 											If not rs3.Eof Then
 
                                            	While not rs3.Eof%>
		    	                                <option value="<%=rs3("cod_fornecedor")%>"><%=rs3("NOME_FANTASIA")%> - <%=formatCurrency(rs3("valor_custo"))%></option>
        	                            		<%
                	                            rs3.MoveNext()
                    	                        Wend	
                                            
											Else		    
											%>
											<option value="">N�o h� fornecedor cadastrado</option>
											<%
											End If
                                        
                                            rs3.Close
                                            Set rs3=Nothing
                                        %>
                                  </select>
                                </div></td>
                            </tr>           
                <tr>
<td colspan="4"><div align="left">
  <table width="100%" border="0" cellspacing="0" cellpadding="2">
    <tr>
      <td width="29"  style="border-bottom:solid 1px #333">&nbsp;</td>
      <td width="20"  style="border-bottom:solid 1px #333">&nbsp;</td>
      <td width="96"  style="border-bottom:solid 1px #333"><div align="center">OS</div></td>
      <td width="449"  style="border-bottom:solid 1px #333"><div align="center">Formando</div></td>
      <td width="97"  style="border-bottom:solid 1px #333"><div align="center">Qtd.</div></td>
      <td width="113"  style="border-bottom:solid 1px #333"><div align="center">Data In&iacute;cio Etapa</div></td>
      <td width="113"  style="border-bottom:solid 1px #333"><div align="center">Prazo final</div></td>
      <td width="149" style="border-bottom:solid 1px #333"><div align="center">Resp.
        Etapa Atual</div></td>
    </tr>
                                                <%	 
		   sql = "SELECT * FROM VW_TB_ORDENS_SERVICO WHERE cod_produto = "&rs("cod_produto")&sqlWhere&" ORDER BY "&sqlOrderBy 
           'Response.Write(sql)
           Set rs2=objConexao.Execute(sql)
           
           While not rs2.Eof
           %>
           		<tr id="linha_<%=rs2("cod_ordem_servico")%>">
                	<td style="border-bottom:solid 1px #999"><div align="center"><img src="images/icons/seta_superior.gif" width="19" height="15" /></div></td>
                    <td style="border-bottom:solid 1px #999"><input type="checkbox" value="S" onclick="alterarQtdSolicitada(<%=rs("cod_produto")%>,<%=rs2("quantidade")%>,this,<%=rs2("cod_ordem_servico")%>)" class="inputRadio inputs_por_os_<%=rs("cod_produto")%>" name="cod_ordem_servico_<%=rs2("cod_ordem_servico")%>" quantidade="<%=rs2("quantidade")%>" /></td>
                    <td style="border-bottom:solid 1px #999"><div align="center"><%=moduloPPVOSMascara(rs2("cod_ordem_servico"),"",true)%></div></td>
                    <td style="border-bottom:solid 1px #999"><%
						sql = 	" SELECT "&_
								" TB_PEDIDOS_FORMANDOS.COD_FORMANDO, "&_
								" TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, "&_
								" TB_FICHAS_IDENTIFICACAO.COD_CONTRATO, "&_
								" TB_ENTIDADES.DESCRICAO, "&_
								" TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO, "&_
								" TB_FICHAS_IDENTIFICACAO.NOME_CURSO "&_
								" FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON "&_
								" TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO "&_
								" INNER JOIN TB_ENTIDADES ON(TB_FICHAS_IDENTIFICACAO.COD_CONTRATO = TB_ENTIDADES.COD_CONTRATO)"&_
								" WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&rs2("cod_pedido")&" ORDER BY NOME_FORMANDO"
											
						           Set rs3=objConexao.Execute(sql)
								   %>
											<table cellpadding="0" width="100%" cellspacing="0">
                                    	<%
                                    	While not rs3.Eof
										%>
                                        	<tr>
                                            <td width="42%"><%=rs3("NOME_FORMANDO")%></td>
                                            <td width="58%"><%=rs3("DESCRICAO")%></td>
                                            </tr>
                						<%
										rs3.MoveNext()
										Wend
										
										%>
                                        </table>
                                        <%
									
									rs3.Close
									Set rs3=Nothing
									%></td>
                              <td style="border-bottom:solid 1px #999"><div align="center">
							  <%=rs2("quantidade")%>
                              <input type="hidden" class="quantidade_por_os_<%=rs("cod_produto")%>" value="<%=rs2("quantidade")%>" />
                              
                              </div></td>
                              <td style="border-bottom:solid 1px #999"><div align="center"><%=ArrumaData(rs2("data_inicio"))%></div></td>
                              <td style="border-bottom:solid 1px #999">

                              <div align="center">
                               <% if DateDiff("d",now(),rs2("prazo_geral")) < 10 then%> 
                               <img src="images/icons/icon_balao_red.gif">
                              <%end if%>
                              	<%=ArrumaData(rs2("prazo_geral"))%>

                              </div></td>
                              <td style="border-bottom:solid 1px #999" align="center">
							  	<%=nomeUsuario(rs2("cod_usuario_responsavel"),true)%>
							  </td>
                            </tr>
           <%
           rs2.MoveNext()
           Wend
           
           rs2.Close
           Set rs2=Nothing
           %>
                                                
                                              </table>
                                            </div></td>
                            </tr>
               
                    
          
           
                    
    <%
    rs.MoveNext()
    Wend
    %>
          <tr>
                  <td colspan="4"><div align="center"><a href="javascript:confirmaTriagem();"><img border="0" src="images/layout/btnGravar.jpg" /></a></div></td>
                  </tr>
         
          </table>
      	</form>

<%
	Else
	%>
    <div align="center" class="msgGrande">
    Nenhuma OS encontrada com etapa atual em aberto para compra!<br />
</div>
    <%End if
	
	rs.Close
	Set rs=Nothing
	
	FecharConexao()
	%>
	  </p>
	   <br />
	   <br />

	    </td>
    </tr>
</table>

 
  
  </td>
    </tr>
</table>

  
  </td>
  </tr>
</table>
<!-- #include file="includes/rodape.asp" -->

  </td>
</tr>
</table>

</body>
</html>