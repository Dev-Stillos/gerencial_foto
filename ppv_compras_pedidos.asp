<%
Dim calendario_necessario, BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
calendario_necessario = "SIM"
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 1
BTN_IMPRIMIR_LINK = "#"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<h2>Compras</h2>
<!-- #include file="includes/compras_sub_menu.asp" -->
<%
	 
sub segundaViaRecibo(cod_pedido_compra, conexao_aberta)

	if not conexao_aberta Then
		AbrirConexao()
	end if

	objConexao.Execute("DELETE FROM TB_PEDIDOS_COMPRAS_PRODUTOS_TEMPORARIA WHERE cod_usuario = "&sqlNumero(session("cod_usuario")))

	objConexao.Execute("INSERT INTO TB_PEDIDOS_COMPRAS_PRODUTOS_TEMPORARIA (cod_pedido_compra_produto, cod_usuario) SELECT cod_pedido_compra_produto, "&session("cod_usuario")&" FROM TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra = "&sqlNumero(cod_pedido_compra) &" AND status_produto = 'E' " )

	if not conexao_aberta Then
		fecharConexao()
	end if


	Response.Redirect("ppv_compras_pedidos_produtos_solicitados_recibo_entrega.asp")


end sub

dim cod_pedido_compra_recibo, acao_recibo

cod_pedido_compra_recibo = Request.QueryString("cod_pedido_compra_recibo")
acao_recibo = Request.QueryString("acao_recibo")

if cod_pedido_compra_recibo <> "" and acao_recibo = "segunda_via_recibo" and validaUsuario(2,,83,true) then 
	call segundaViaRecibo(cod_pedido_compra_recibo, false)
end if

	AbrirConexao()
	
	Dim validacao_para_listar_qualquer_pedido
	validacao_para_listar_qualquer_pedido = validaUsuario(2,,73,true)
	'/////////////////////////////////////////////////////////////
	'Rotina de corre&ccedil;&atilde;o de pedidos pendentes
	
		'1) Lista dos os pedidos de compras pendentes
		Set rs=objConexao.Execute("SELECT * FROM TB_PEDIDOS_COMPRAS WHERE status_pedido_compra = 2")
			
			While not rs.Eof
				
				'2) Verifica se h&aacute; itens pendentes realmente
				Set rs2=objConexao.Execute("SELECT * FROM TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra = "&rs("cod_pedido_compra")&" AND status_produto = 'P'")
					
					If not rs2.Eof Then
						
						pode_finalizar = true
						
						While not rs2.Eof
							
							If Cstr(""&rs2("cod_ordem_servico_etapa")) <> "" Then
							
								'3) Verifica se o item pendente pode ser excluido, pois a OS foi excluida
								Set rs3=objConexao.Execute("SELECT * FROM TB_ORDENS_SERVICO_ETAPAS WHERE cod_ordem_servico_etapa = "&rs2("cod_ordem_servico_etapa"))
									
									If not rs3.Eof Then
									'Encontrou ent&atilde;o n&atilde;o pode excluir, pois esta pendente e h&aacute; relacionamento
										pode_finalizar = false
									Else
									'Cancela, pois OS n&atilde;o existe mais.
										objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS_PRODUTOS SET status_produto = 'C' WHERE cod_pedido_compra_produto = "&rs2("cod_pedido_compra_produto"))
									End If
								
								rs3.Close
								Set rs3=Nothing
							
							Else
								pode_finalizar = false
							End If
							
						rs2.MoveNext()
						Wend
						
						If pode_finalizar Then
							'Response.Write("Baixa o pedido de compra => "&rs("numero_pedido_compra")&" OS n&atilde;o existe mais<br />")		
							objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS set status_pedido_compra = 3 WHERE cod_pedido_compra = "&rs("cod_pedido_compra"))
							
							
						End If
					
					Else
						
						'Baixa o pedido de compra
						'Response.Write("Baixa o pedido de compra => "&rs("numero_pedido_compra")&"<br />")
						objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS set status_pedido_compra = 3 WHERE cod_pedido_compra = "&rs("cod_pedido_compra"))
						
					End If
				
				rs2.Close
				Set rs2=Nothing
				
			rs.MoveNext()
			Wend
		
		rs.Close
		Set rs=Nothing
	
	'/////////////////////////////////////////////////////////////
	
	Dim sql, sqlWhere, sqlOrderBy
	sql = "SELECT * FROM VW_TB_PEDIDOS_COMPRAS WHERE status_pedido_compra > 0 "
	sqlOrderBy = " order by numero_pedido_compra desc"
	
	Dim status_pedido_compra, diferenca_prazo_entrega

	'/////////////
	'Filtro
		Dim filtro_status_pedido_compra
		Dim filtro_cod_fornecedor
		Dim filtro_cadastro_cod_usuario
		Dim filtro_NUMERO_PEDIDO_COMPRAS
		Dim b_DATA_CADASTRO_INICIO
		Dim b_DATA_CADASTRO_FIM
		
		filtro_NUMERO_PEDIDO_COMPRAS	= Cstr(""&Request.QueryString("NUMERO_PEDIDO_COMPRAS"))
		filtro_status_pedido_compra		= Cstr(""&Request.QueryString("status_pedido_compra"))
		filtro_cod_fornecedor			= Cstr(""&Request.QueryString("cod_fornecedor"))
		filtro_cadastro_cod_usuario		= Cstr(""&Request.QueryString("cadastro_cod_usuario"))
		b_DATA_CADASTRO_INICIO			=	Request.QueryString("DATA_CADASTRO_INICIO")
		b_DATA_CADASTRO_FIM				=	Request.QueryString("DATA_CADASTRO_FIM")
		
		If Not validacao_para_listar_qualquer_pedido Then
			sqlWhere = sqlWhere&" AND cadastro_cod_usuario = '"&Session("cod_usuario")&"'"
		End If
		
		If Request.QueryString("filtro") <> "" Then
		
			If filtro_NUMERO_PEDIDO_COMPRAS <> "" Then
			'Filtro preferencial, limpa os demais
				
				filtro_cod_fornecedor = ""
				filtro_cadastro_cod_usuario = ""
				b_DATA_CADASTRO_INICIO = ""
				b_DATA_CADASTRO_FIM	= ""
				
					sqlWhere = sqlWhere&" AND numero_pedido_compra = '"&filtro_NUMERO_PEDIDO_COMPRAS&"'"
					
					If filtro_status_pedido_compra <> "" Then
						sqlWhere = sqlWhere&" AND status_pedido_compra = '"&filtro_status_pedido_compra&"'"
					End If
			
			Else
		
				If filtro_status_pedido_compra <> "" Then
					sqlWhere = sqlWhere&" AND status_pedido_compra = '"&filtro_status_pedido_compra&"'"
				End If
				
				if filtro_cod_fornecedor <> "" Then
					sqlWhere = sqlWhere&" AND cod_fornecedor = '"&filtro_cod_fornecedor&"'"
				End if
				
				if filtro_cadastro_cod_usuario <> "" Then
					sqlWhere = sqlWhere&" AND cadastro_cod_usuario = '"&filtro_cadastro_cod_usuario&"'"
				End if			
				
				
				If b_DATA_CADASTRO_INICIO <> Empty or b_DATA_CADASTRO_FIM <> Empty Then
							
							
							If b_DATA_CADASTRO_INICIO <> Empty Then
								b_DATA_CADASTRO_INICIO = Day(b_DATA_CADASTRO_INICIO)&"/"&Month(b_DATA_CADASTRO_INICIO)&"/"&Year(b_DATA_CADASTRO_INICIO)&" 00:00:00"
							End If
		
							If b_DATA_CADASTRO_FIM <> Empty Then
								b_DATA_CADASTRO_FIM = Day(b_DATA_CADASTRO_FIM)&"/"&Month(b_DATA_CADASTRO_FIM)&"/"&Year(b_DATA_CADASTRO_FIM)&"  23:59:59"
							End If					
																	
							If b_DATA_CADASTRO_INICIO <> Empty and b_DATA_CADASTRO_FIM <> Empty Then
								sqlWhere = sqlWhere&" AND DATA_CADASTRO BETWEEN "&sqlData(b_DATA_CADASTRO_INICIO)&" AND "&sqlData(b_DATA_CADASTRO_FIM)
							Elseif b_DATA_CADASTRO_INICIO = Empty and b_DATA_CADASTRO_FIM <> Empty Then
								sqlWhere = sqlWhere&" AND DATA_CADASTRO <= "&sqlData(b_DATA_CADASTRO_FIM)
							Elseif b_DATA_CADASTRO_INICIO <> Empty and b_DATA_CADASTRO_FIM = Empty Then
								sqlWhere = sqlWhere&" AND DATA_CADASTRO >= "&sqlData(b_DATA_CADASTRO_INICIO)
							End If
					
							If b_DATA_CADASTRO_FIM <> Empty Then
							'Corrigi data
								b_DATA_CADASTRO_FIM = Day(b_DATA_CADASTRO_FIM)&"/"&Month(b_DATA_CADASTRO_FIM)&"/"&Year(b_DATA_CADASTRO_FIM)&"  00:00:00"
							End If	
					
				End if
			
			
			End If
			
	End If	
	sql = sql&sqlWhere&sqlOrderBy
	%>
  <h2><%=titulo_pagina%></h2><br />
<div class="subMenu" style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC">
<form action="<%=caminho_pagina%>" name="FORMULARIO_FILTRO" method="get">
<input name="" type="submit" style="display:none; visibility:hidden" />
<table width="483" border="0" cellpadding="0" cellspacing="0"><tr>
  <td width="212" style="padding-left:10px;"><table width="483" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="43" rowspan="10"><span style="font-weight: bold">Filtrar</span></td>
    <td colspan="2" style="padding-left:10px; padding-top:5px;">N&uacute;mero do Pedido de Compras</td>
  </tr>
  <tr>
    <td colspan="2" style="padding-left:10px;"><input type='text' class="inputTamanho1" value="<%=filtro_NUMERO_PEDIDO_COMPRAS%>" name='NUMERO_PEDIDO_COMPRAS' /></td>
  </tr>
  <tr>
    <td width="240"  style="padding-left:10px; padding-top:5px;">Respons&aacute;vel pelo pedido</td>
    <td width="200"  style="padding-left:10px; padding-top:5px;">Status do Pedido de Compras</td>
  </tr>
  <tr>
    <td style="padding-left:10px;">
    <%If validacao_para_listar_qualquer_pedido Then%>
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
    <%End If%>
    </td>
    <td style="padding-left:10px;"><select name="status_pedido_compra" class="inputTamanho1">
      <option <%If Cstr(filtro_status_pedido_compra) = "" Then%> selected="selected" <%End If%> value="">Todos</option>
      <%
		Dim temp
		For i=1 to 4
			If i=4 Then
				temp = Cstr("99")
			Else
				temp = Cstr(i)
			End If

		pedidoCompraStatus(i)
		%>
      <option <%If Cstr(filtro_status_pedido_compra) = temp Then%> selected="selected" <%End If%> value="<%=temp%>"><%=pedidoCompraStatus(temp)%></option>
      <%

		Next
    	%>
    </select></td>
  </tr>
  <tr>
    <td colspan="2" style="padding-left:10px; padding-top:5px;">Fornecedor</td>
  </tr>
  <tr>
    <td colspan="2" style="padding-left:10px;"><select class="inputTamanho2" name="COD_FORNECEDOR">
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
  </tr>
  <tr>
    <td colspan="2" style="padding-left:10px; padding-top:5px;">Data Cadastro</td>
  </tr>
  <tr>
    <td colspan="2" style="padding-left:10px;"><table width="300" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="135"><table width='135' border='0' align="left" cellpadding='0' cellspacing='0'>
          <tr>
            <td width='102'><input class='campoData' maxlength='10' onkeydown='mascara(this,data);' onkeypress='mascara(this,data);' onkeyup='mascara(this,data);' type='text' name='DATA_CADASTRO_INICIO' value="<%=ArrumaData(b_DATA_CADASTRO_INICIO)%>" onblur='validarData(this)' /></td>
            <td width='298'><div align='center'> <img onclick='pickDate(this,document.FORMULARIO_FILTRO.DATA_CADASTRO_INICIO);' class='forcaBotao' src='images/icons/icon_calendario.jpg' width='26' height='22' /></div></td>
          </tr>
        </table></td>
        <td width="11"><div align="center">&agrave;</div></td>
        <td width="154"><table width='135' border='0' align="left" cellpadding='0' cellspacing='0'>
          <tr>
            <td width='102'><input class='campoData' maxlength='10' onkeydown='mascara(this,data);' onkeypress='mascara(this,data);' onkeyup='mascara(this,data);' type='text' name='DATA_CADASTRO_FIM' value="<%=ArrumaData(b_DATA_CADASTRO_FIM)%>" onblur='validarData(this)' /></td>
            <td width='298'><div align='center'> <img onclick='pickDate(this,document.FORMULARIO_FILTRO.DATA_CADASTRO_FIM);' class='forcaBotao' src='images/icons/icon_calendario.jpg' width='26' height='22' /></div></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="2" style="padding-left:10px;"><div align="left" style="padding-top:10px">
      <input class="btnFiltrarAmarelo" type="submit" name="filtro1" value=" " />
    </div></td>
  </tr>
  
  </table>    <br /></td>
  </tr>
</table>
<input type="hidden" value="1" name="filtro" />
</form>
</div>

<%
	
	Set rs=server.CreateObject("ADODB.recordset")
	rs.CursorLocation=3  'Cursor no cliente
	rs.PageSize=50
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
	
    <%
	Dim sucesso, codmsg
	codmsg	= CStr(""&Request.QueryString("codmsg"))
	sucesso = CStr(""&Request.QueryString("sucesso"))
		
	If sucesso = "" and codmsg <> "" Then%>
    	<div align="center" class="msgGrandeSucesso">
        	<%
			If codmsg = "1" Then
				Response.Write("Pedido de Compra inclu&iacute;do com sucesso!<br />ATEN&Ccedil;&Atilde;O IMPRIMA 3 C&oacute;PIAS, SENDO 1 PARA VOC�, OUTRA PARA O FORNECEDOR E A �LTIMA PARA SER ENTREGUE PARA O DPTO. FINANCEIRO JUNTO COM O BOLETO.")
			Elseif codmsg = "2" Then
				Response.Write("Pedido de Compra alterado com sucesso!")
			Elseif codmsg = "3" Then	
				Response.Write("Pedido de Compra exclu&iacute;do com sucesso!")
			Elseif codmsg = "202" or codmsg = "203" or codmsg = "2099" or codmsg = "2022" Then		
			'Mudan&ccedil;a de status com resposta dinamica
				Response.Write("Pedido de Compra alterado o status para """&pedidoCompraStatus(Mid(codmsg,3,Len(codmsg)))&""" com sucesso!")
			Elseif codmsg = "601" Then
				Response.Write("Pedidos de Compra gerados com sucesso!")
			End If	
			%>
</div>
    <%ElseIf sucesso <> "" and codmsg <> "" Then%>
        <div align="center" class="msgGrandeErro">
			<%
			If codmsg = "4" Then
				Response.Write("Pedido de Compra n&atilde;o alterado para n&atilde;o haver duplicidade de informa&ccedil;&atilde;o!<BR>O conte&uacute;do do campo """&Request.QueryString("chaveUnica")&""" j&aacute; possui uma informa&ccedil;&atilde;o igual em outro Pedido de Compra.<br />")
			Elseif codmsg = "5" Then
				Response.Write("Pedido de Compra n&atilde;o exclu&iacute;do, pois h&aacute; Or&ccedil;amentos/Pedidos/NF que dependem ou relacionam com este vendedor.<BR>Se n&atilde;o desejar ter mais este Pedido de Compra inative ao inv&eacute;s de excluir.<br />")
			
			Elseif codmsg = "300" Then
				Response.Write("Pedido de Compra N&atilde;O alterado status, <br />pois voc&ecirc; n&atilde;o possui acesso ou o pedido j&aacute; foi mudado para o status que voc&ecirc; est&aacute; tentando mudar.")
			
			Elseif codmsg = "900" Then
				Response.Write("Voc&ecirc; n&atilde;o possui acesso a VISUALIZA&ccedil;&atilde;O!")
			Elseif codmsg = "901" Then
				Response.Write("Voc&ecirc; n&atilde;o possui acesso a INCLUS&atilde;O!")
			Elseif codmsg = "902" Then
				Response.Write("Voc&ecirc; n&atilde;o possui acesso a ALTERA&ccedil;&atilde;O!")	
			Elseif codmsg = "903" Then
				Response.Write("Voc&ecirc; n&atilde;o possui acesso a EXCLUS&atilde;O!")
			End If					
			%>
</div>
<%End If%>
<table width="99%" border="0" align="center">
  <tr>
    <td width="47%"> 
P&aacute;gina <strong><%=rs.AbsolutePage%></strong> de <%=rs.PageCount%>
</td>
    <td width="53%"><div align="right">Total <strong><%=rs.RecordCount%></strong> pedidos de compra</div></td>
  </tr>
</table>
<table width="99%" border="1" align="center" cellpadding="5" style="border-collapse:collapse; text-align: right;">
<tr bgcolor="whitesmoke">
<td width="7%"><div align="center" style="font-weight: bold">C&oacute;d. Pedido Compra</div></td>
<td width="20%"><div align="center" style="font-weight: bold">Fornecedor</div></td>
<td width="14%"><div align="center"><strong>Produto</strong></div></td>
<td width="8%"><div align="center"><strong>Qtd. Pendente</strong></div></td>
<td width="10%" align="center"><div align="center"><strong>Qtd. Entregue</strong></div></td>
<td width="12%" align="center"><div align="center" style="font-weight: bold">Respons&aacute;vel pedido</div></td>
<td width="5%" align="center"><div align="center"><strong>Data Emiss&atilde;o</strong></div></td>
<td width="8%" align="center"><strong> Dias Atrasados</strong></td>
<td width="4%" align="center"><strong>Status</strong></td>
<td width="4%">&nbsp;</td>
<td width="8%">&nbsp;</td>
</tr>
<%
Dim qtd_pendente, qtd_entregue, qtd_atrasados
Dim usuario_pode_imprimir
usuario_pode_imprimir = validaUsuario(1,100,0,true)

Dim x
For x=1 To rs.PageSize
If Not rs.Eof Then
	
	status_pedido_compra = cstr(rs("status_pedido_compra"))
	status_pedido_compra = int(status_pedido_compra)
	qtd_pendente = pedidoCompraTotais(rs("cod_pedido_compra"),5,true)
	qtd_entregue = pedidoCompraTotais(rs("cod_pedido_compra"),4,true)
	qtd_atrasados= pedidoCompraTotais(rs("cod_pedido_compra"),7,true)
%>
<tr>
  <td><div align="center"><%=rs("numero_pedido_compra")%>
  <% if rs("pedido_manual") = "S" Then%>
  	(manual)
  <% end if %>
  </div></td>
  <td><div align="left"><%=rs("nome_fantasia")%></div></td>
  <td><div align="left">
  <%
  Set rs2=objConexao.Execute("SELECT nome_produto FROM VW_TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra = "&rs("cod_pedido_compra")&" GROUP BY nome_produto ORDER BY nome_produto ASC")
  	If not rs2.Eof Then
		While not rs2.Eof
			Response.Write(rs2("nome_produto")&"&nbsp;&nbsp;")
		rs2.MoveNext
		Wend
	End If
  rs2.Close
  Set rs2=Nothing
  %>
  </div></td>
  <td><div align="center"><%=qtd_pendente%></div></td>
  <td><div align="center"><%=qtd_entregue%></div></td>
  <td><div align="left"> <%=nomeUsuario(rs("cadastro_cod_usuario"),true)%> </div></td>
  <td><div align="center"><%=ArrumaData(rs("data_cadastro"))%></div></td>
  <td><div align="center">
  <%
  If qtd_atrasados > 0 Then
	  Set rs2=objConexao.Execute("SELECT MIN(data_previsao_entrega) AS data_previsao_entrega FROM VW_TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra = "&rs("cod_pedido_compra"))
		If not rs2.Eof Then
			Response.Write(DateDiff("d",rs2("data_previsao_entrega"),date()))
		End If
	  rs2.Close
	  Set rs2=Nothing
  Else
  	Response.Write("-")
  End If
  %>
  
  </div></td>
  <td align="center"><div align="center"><a href="javascript:alert('<%=pedidoCompraStatus(rs("status_pedido_compra"))%>')" style="font-weight: bold"><%=rs("status_pedido_compra")%></a></div></td>
  <td align="center">
  <%
  If qtd_pendente = 0 Then
  'N&atilde;o h&aacute; itens pendentes
  %>
  	<img title="Pedido de Compra com todos itens entregues" src="images/icons/icon_verificado.gif" width="20" height="19" />
  <%Elseif qtd_atrasados > 0 Then
  'H&aacute; itens pendentes e atrasados
  %>
  	<img title="H&aacute; itens pendentes e atrasados" src="images/icons/icon_balao_red.gif" />
  <%
  Else
  'H&aacute; itens pendentes e n�o atrasados
  %>
  	<img title="H&aacute; itens pendentes e n&atilde;o atrasados" src="images/icons/icon_balao_green.gif" />
  <%End If%>  </td>
  <td>
    
    <table border="0" align="center" cellpadding="3" cellspacing="0">
      <tr>
         <%
		 'Impress&atilde;o somente se o status do pedido de compra for Enviado para o fornecedor, pois neste caso j&aacute; foi aprovado e mudado o status para Enviado para fornecedor para for&ccedil;ar o usu&aacute;rio a mudar o status para Enviado para Fornecedor e ter como controlar a data que foi enviado.
		If status_pedido_compra > 1  Then%>
         <td><a href="ppv_compras_pedidos_imprimir.asp?c=<%=rs("cod_pedido_compra")%>" target="_blank">
          <img src="images/icons/icon_imprimir.png" alt="Segunda Via Pedido" title="Segunda Via Pedido" border="0" /></a></td>
          <%
		 End If

		 If status_pedido_compra > 2 and validaUsuario(2,,83,true) Then%>
         <td><a href="ppv_compras_pedidos.asp?cod_pedido_compra_recibo=<%=rs("cod_pedido_compra")%>&acao_recibo=segunda_via_recibo" target="_blank">
          <img src="images/icons/icon_recibo.png" alt="Segunda Via recibo" title="Segunda Via recibo" border="0" /></a></td>
          <%
		 End If
		  %>
          
       
        <td>    <a href="ppv_compras_pedidos_formulario.asp?comando=visualizar&cod_pedido_compra=<%=rs("cod_pedido_compra")%>">
          <img src="images/icons/icon_visualizar.png" alt="Visualizar Registro" title="Visualizar Registro" border="0" />
          </a></td>
        
        
        <%
  'Para alterar ou excluir somentes nos seguintes status
  'Se n&atilde;o estiver Enviado para fornecedor (status = 3)
  'Depois &eacute; permitido somente o cancelamento do pedido de compras completo ou cancelamento individual por item.
  %>
        
        <%If int(status_pedido_compra) < 3 and int(status_pedido_compra) <> 99 Then%>
			
                    <td>
                        <a href="ppv_compras_pedidos_formulario.asp?comando=alterar&cod_pedido_compra=<%=rs("cod_pedido_compra")%>">
              <img src="images/icons/icon_alterar.png" alt="Alterar Registro" title="Alterar Registro" width="16" height="16" border="0" /></a> 	 	</td>
       
			<%If  int(status_pedido_compra) <> 99 Then%>
            <td>  <a href="ppv_compras_pedidos_formulario.asp?comando=excluir&cod_pedido_compra=<%=rs("cod_pedido_compra")%>">
              <img src="images/icons/icon_excluir.png" alt="Excluir Registro" title="Excluir Registro" border="0" />
              </a></td>
          	 <%End If%>
        <%End If%>
        </tr>
</table>  </td>
  </tr>
<%
rs.MoveNext()
End If
Next
%>
<%If rs.PageCount > 1 Then %>                                  
     <tr bgcolor="whitesmoke">
      <td colspan="11">
      	   <%call paginacao(paginacaocao,rs.PageCount)%>      </td>
    </tr>   
<%End if%>
</table>
<br />

<strong>&nbsp;Legenda de Status:</strong><br />
<table width="612" border="0" cellpadding="2" cellspacing="0">
<%
'Criador dinamico de legenda de status de pedido de compra, por enquanto h&aacute; 7 status v&aacute;lidos (sem contas o status em cadastro)

For i=1 to 4%>
	 <tr>
		<%If i=4 Then%>
		   
				<td width="12" align="center">99</td>
	   <td width="284"><%=pedidoCompraStatus(99)%></td>
<%Else%>
				<td width="20" align="center"><%=i%></td>
	   <td width="280"><%=pedidoCompraStatus(i)%></td>
		<%End if%>
	</tr>

<%
Next
%>
</table>

<br />	
<p><strong>&nbsp;Legenda de &iacute;cones R&aacute;pidos:</strong><br />
</p>
<table width="612" border="0" cellpadding="2" cellspacing="0">
  <tr>
    <td width="25"><div align="center"><img src="images/icons/icon_alterar.png" width="16" height="16" border="0" /></div></td>
    <td width="579">Alterar Pedido de Compra</td>
  </tr>
  <tr>
    <td><div align="center"> <img src="images/icons/icon_excluir.png" width="16" height="16" border="0" /></div></td>
    <td>Excluir Pedido de Compra  </td>
  </tr>
  
  <tr>
  <td><div align="center"><img src="images/icons/icon_imprimir.png" alt="" title="" border="0" /></div></td>
  <td>Imprimir Pedido de Compras</td>
  </tr>
  <tr>
    <td><div align="center"><img src="images/icons/icon_visualizar.png" alt="Visualizar" width="16" height="16" border="0" /></div></td>
    <td>Visualizar Pedido de Compra</td>
  </tr>
  
  <tr>
    <td><div align="center"><img title="Pedido de Compra dentro do prazo" src="images/icons/icon_balao_green.gif" /></div></td>
    <td>Pedido de Compra com todos os itens dentro do prazo</td>
  </tr>
  <tr>
    <td><div align="center"><img title="Pedido de Compra vencido" src="images/icons/icon_balao_red.gif" /></div></td>
    <td>Pedido de Compra com item(ns) fora  do prazo para entrega</td>
  </tr>
<tr>
    <td><div align="center"><img title="Pedido de Compra Finalizado e Entregue" src="images/icons/icon_verificado.gif" width="20" height="19" /></div></td>
    <td>Pedido de Compra com todos os itens entregues</td>
  </tr>
</table>

<%
	Else
	%>
    <div align="center" class="msgGrande">
    	Nenhum Pedido de Compra encontrado!</div>
    <%End if
	
	rs.Close
	Set rs=Nothing
	
	FecharConexao()
	%>
    <br />
	<br />
	<br />
<!-- #include file="files/includes/miolo-inferior.asp" -->
</body>
</html>