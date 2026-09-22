<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 1
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
  Dim titulo_pagina
  titulo_pagina = "CAP - Formul&aacute;rio de Per&iacute;odos de Vendas"

'/////////////////////////////////////////////////////////////////////////////////////////////

'data in�cio formatado, data t�rmino, t�tulo da p�gina
'call manutencao(ArrumaData(date(),"")& "�s 08h00",ArrumaData(date(),"")& "�s 12h00","")

Dim calendario_necessario
calendario_necessario = "N&Atilde;O"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%

'Valida��o especial para verificar se possui acesso a p�gina m�e de formul�rio
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS where COD_PAGINA = '228' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se n�o encontrar ent�o redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
	rs.Close
	Set rs=Nothing
	
FecharConexao()


'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 11/1/2012 12:13:51

'Declara vari�veis para utiliza��o do sistema
Dim id, campo1, comando, cod_tabela_comissao_vendedores, actionformulario

actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
cod_tabela_comissao_vendedores = Request.QueryString("cod_tabela_comissao_vendedores")
'Declara vari�veis
Dim cod_produto, tipo_pagamento_normal_avista, valor_pagamento_normal_avista, tipo_pagamento_normal_aprazo, valor_pagamento_normal_aprazo, tipo_pagamento_repasse_avista, valor_pagamento_repasse_avista, tipo_pagamento_repasse_aprazo, valor_pagamento_repasse_aprazo, cod_usuario, modificado



If campo1 <> Empty Then
	
	'Padr�o
	cod_usuario = Session("cod_usuario")
	modificado = now()


	AbrirConexao()
				Set rs=objConexao.Execute("SELECT cod_produto, nome_produto FROM TB_PRODUTOS WHERE ativo = 1 and PAGAMENTO_COMISSAO = 'S' order by order_menu")
					
					If not rs.Eof Then
					
						While not rs.Eof
						
						cod_produto = rs("cod_produto") 
					
						'///////////////////////////////////////////////////////////							
								
						tipo_pagamento_normal_avista = Request.Form("tipo_pagamento_normal_avista_"&cod_produto)
						valor_pagamento_normal_avista = "0"&Request.Form("valor_pagamento_normal_avista_"&cod_produto)
						tipo_pagamento_normal_aprazo = Request.Form("tipo_pagamento_normal_aprazo_"&cod_produto)
						valor_pagamento_normal_aprazo = "0"&Request.Form("valor_pagamento_normal_aprazo_"&cod_produto)
						tipo_pagamento_repasse_avista = Request.Form("tipo_pagamento_repasse_avista_"&cod_produto)
						valor_pagamento_repasse_avista = "0"&Request.Form("valor_pagamento_repasse_avista_"&cod_produto)
						tipo_pagamento_repasse_aprazo = Request.Form("tipo_pagamento_repasse_aprazo_"&cod_produto)
						valor_pagamento_repasse_aprazo = "0"&Request.Form("valor_pagamento_repasse_aprazo_"&cod_produto)
						
						'///////////////////////////////////////////////////////////
						'Verifica se possui o produto cadastrado nesta tabela de venda
						'Para os comandos de incluir ou alterar obs.: N�O H� EXCLUIR!!! POIS PARA OS PRODUTOS QUE
						'FOREM INATIVADOS N�O SER�O LISTADOS E TAMB�M OS PRODUTOS EXCLUIDOS SER�O EXCLUIDOS NESTA
						'TABELA ABAIXO TAMB�M
						Set rs2=objConexao.Execute("SELECT * FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES_PRODUTOS WHERE cod_tabela_comissao_vendedores = "&cod_tabela_comissao_vendedores&" and cod_produto ="&cod_produto)
					
							If not rs2.Eof Then
								'J� existe por tanto altera!								
								objConexao.Execute("UPDATE TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES_PRODUTOS SET tipo_pagamento_normal_avista = "&sqlTexto(tipo_pagamento_normal_avista)&", valor_pagamento_normal_avista = '"&sqlNumero(valor_pagamento_normal_avista)&"', tipo_pagamento_normal_aprazo = "&sqlTexto(tipo_pagamento_normal_aprazo)&", valor_pagamento_normal_aprazo = '"&sqlNumero(valor_pagamento_normal_aprazo)&"', tipo_pagamento_repasse_avista = "&sqlTexto(tipo_pagamento_repasse_avista)&", valor_pagamento_repasse_avista = '"&sqlNumero(valor_pagamento_repasse_avista)&"', tipo_pagamento_repasse_aprazo = "&sqlTexto(tipo_pagamento_repasse_aprazo)&", valor_pagamento_repasse_aprazo = '"&sqlNumero(valor_pagamento_repasse_aprazo)&"', cod_usuario = "&sqlTexto(cod_usuario)&", modificado = "&sqlData(modificado)&" WHERE cod_produto = "&cod_produto&" and cod_tabela_comissao_vendedores = "&cod_tabela_comissao_vendedores)
								
								
							Else
								'N�o existe incluir								
								objConexao.Execute("INSERT INTO TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES_PRODUTOS (cod_produto, cod_tabela_comissao_vendedores, tipo_pagamento_normal_avista, valor_pagamento_normal_avista, tipo_pagamento_normal_aprazo, valor_pagamento_normal_aprazo, tipo_pagamento_repasse_avista, valor_pagamento_repasse_avista, tipo_pagamento_repasse_aprazo, valor_pagamento_repasse_aprazo, cod_usuario, modificado) VALUES ("&sqlTexto(cod_produto)&", "&sqlTexto(cod_tabela_comissao_vendedores)&", "&sqlTexto(tipo_pagamento_normal_avista)&", '"&sqlNumero(valor_pagamento_normal_avista)&"', "&sqlTexto(tipo_pagamento_normal_aprazo)&", '"&sqlNumero(valor_pagamento_normal_aprazo)&"', "&sqlTexto(tipo_pagamento_repasse_avista)&", '"&sqlNumero(valor_pagamento_repasse_avista)&"', "&sqlTexto(tipo_pagamento_repasse_aprazo)&", "&sqlNumero(valor_pagamento_repasse_aprazo)&", "&sqlTexto(cod_usuario)&", "&sqlData(modificado)&")")
													
							End If
						
						rs2.Close
						Set rs2=Nothing
						

						rs.MoveNext()
						Wend
						
						
						'Limpa todos os produtos inativos desta tabela ou que n�o � para pagar comiss�o
						objConexao.Execute("DELETE FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES_PRODUTOS WHERE cod_tabela_comissao_vendedores = "&cod_tabela_comissao_vendedores&" AND cod_produto IN(SELECT cod_produto FROM TB_PRODUTOS WHERE ativo = 2 or PAGAMENTO_COMISSAO = 'N')")					
					
					End If
				
				rs.Close
				Set rs=Nothing
	FecharConexao()

	Response.Redirect("cap_fechamentos_tabelas_comissao_vendedores.asp?codmsg=101")
	
End If

'----------------------------------------------------------------------------------------

%>

<h2><%=titulo_pagina%></h2>
<h3 align="center">Produtos da 
<%
	AbrirConexao()

			Set rs=objConexao.Execute("SELECT descricao FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES where cod_tabela_comissao_vendedores = "&cod_tabela_comissao_vendedores)
				
				If not rs.Eof Then
					Response.Write(rs(0))
				End If
			
			rs.Close
			Set rs=Nothing

	FecharConexao()

%>
</h3>
<script type="text/javascript">
function replicarValor(cod_produto){
	var d = document.FORMULARIO_PADRAO;
	var tipo_pagamento_normal_avista, valor_pagamento_normal_avista, tipo_pagamento_normal_aprazo, valor_pagamento_normal_aprazo, tipo_pagamento_repasse_avista, valor_pagamento_repasse_avista, tipo_pagamento_repasse_aprazo, valor_pagamento_repasse_aprazo;
	
	tipo_pagamento_normal_avista = d['TIPO_PAGAMENTO_NORMAL_AVISTA_'+cod_produto].value;
	valor_pagamento_normal_avista = d['VALOR_PAGAMENTO_NORMAL_AVISTA_'+cod_produto].value;
	tipo_pagamento_normal_aprazo = d['TIPO_PAGAMENTO_NORMAL_APRAZO_'+cod_produto].value;
	valor_pagamento_normal_aprazo = d['VALOR_PAGAMENTO_NORMAL_APRAZO_'+cod_produto].value;
	tipo_pagamento_repasse_avista = d['TIPO_PAGAMENTO_REPASSE_AVISTA_'+cod_produto].value;
	valor_pagamento_repasse_avista = d['VALOR_PAGAMENTO_REPASSE_AVISTA_'+cod_produto].value;
	tipo_pagamento_repasse_aprazo = d['TIPO_PAGAMENTO_REPASSE_APRAZO_'+cod_produto].value;
	valor_pagamento_repasse_aprazo = d['VALOR_PAGAMENTO_REPASSE_APRAZO_'+cod_produto].value;
	
	if(valor_pagamento_normal_avista !=='' && valor_pagamento_normal_avista !=='0,00'){
		
		if(valor_pagamento_normal_aprazo == ''){
		//Se n�o tiver valor altera para o valor � vista venda normal
			d['TIPO_PAGAMENTO_NORMAL_APRAZO_'+cod_produto].value = tipo_pagamento_normal_avista;
			d['VALOR_PAGAMENTO_NORMAL_APRAZO_'+cod_produto].value = valor_pagamento_normal_avista;
		}
		
		if(valor_pagamento_repasse_avista == ''){
		//Se n�o tiver valor altera para o valor � vista venda normal
			d['TIPO_PAGAMENTO_REPASSE_AVISTA_'+cod_produto].value = tipo_pagamento_normal_avista;
			d['VALOR_PAGAMENTO_REPASSE_AVISTA_'+cod_produto].value = valor_pagamento_normal_avista;	
		}
	
	
		if(valor_pagamento_repasse_aprazo == ''){
		//Se n�o tiver valor altera para o valor � vista venda normal
			d['TIPO_PAGAMENTO_REPASSE_APRAZO_'+cod_produto].value = tipo_pagamento_normal_avista;
			d['VALOR_PAGAMENTO_REPASSE_APRAZO_'+cod_produto].value = valor_pagamento_normal_avista;	
		}		

	}


	
}
</script>
<form action='<%=actionformulario%>?comando=<%=comando%>&cod_tabela_comissao_vendedores=<%=cod_tabela_comissao_vendedores%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post'>
<%
AbrirConexao()
			Set rs=objConexao.Execute("SELECT cod_produto, nome_produto FROM stillos3.TB_PRODUTOS WHERE ativo = 1 and PAGAMENTO_COMISSAO = 'S' order by order_menu")
				
				If not rs.Eof Then%>
				<table width='750' border='1' align='center' cellpadding='3' style='border-collapse:collapse'>
<tr>
  <td width='36%' rowspan="3"><div align="center"><strong>Produto</strong></div></td>
<td colspan="4"><div align="center"><strong>Vendas Normais</strong></div></td>
<td colspan="4"><div align="center"><strong>Vendas Repasse</strong></div></td>
</tr>
<tr>
  <td colspan="2"><div align="center"><strong>&Agrave; vista</strong></div></td>
  <td colspan="2"><div align="center"><strong>&Agrave; prazo</strong></div></td>
  <td colspan="2"><div align="center"><strong>&Agrave; vista</strong></div></td>
  <td colspan="2"><div align="center"><strong>&Agrave; prazo</strong></div></td>
</tr>
<tr>
  <td width="8%"><div align="center"><strong>Tipo</strong></div></td>
  <td width="8%"><div align="center"><strong>Valor</strong></div></td>
  <td width='8%'><div align="center"><strong>Tipo</strong></div></td>
  <td width='8%'><div align="center"><strong>Valor</strong></div></td>
  <td width='8%'><div align="center"><strong>Tipo</strong></div></td>
  <td width='8%'><div align="center"><strong>Valor </strong></div></td>
  <td width='8%'><div align="center"><strong>Tipo</strong></div></td>
  <td width='8%'><div align="center"><strong>Valor </strong></div></td>
</tr>
				<%
				Dim cor, caminho_validacao_completa
				
				cor = "#FFFFCC"
				caminho_validacao_completa = ""
				
					While not rs.Eof
					
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
					
					'//////////////////////////////////////////////////////////
					'Valida��o
						'caminho_validacao_completa = caminho_validacao_completa&"|TIPO_PAGAMENTO_NORMAL_AVISTA_"&cod_produto&"#Tipo de Pagamento para Vendas Normais#2|VALOR_PAGAMENTO_NORMAL_AVISTA_"&cod_produto&"#Valor Pagamento para Vendas Normais#1|TIPO_PAGAMENTO_NORMAL_APRAZO_"&cod_produto&"#Tipo de Pagamento para Vendas Normais#2|VALOR_PAGAMENTO_NORMAL_APRAZO_"&cod_produto&"#Valor Pagamento para Vendas Normais#1|TIPO_PAGAMENTO_REPASSE_AVISTA_"&cod_produto&"#Tipo de Pagamento para Vendas Repasse#2|VALOR_PAGAMENTO_REPASSE_AVISTA_"&cod_produto&"#Valor Pagamento para Vendas Repasse#1|TIPO_PAGAMENTO_REPASSE_APRAZO_"&cod_produto&"#Tipo de Pagamento para Vendas Repasse#2|VALOR_PAGAMENTO_REPASSE_APRAZO_"&cod_produto&"#Valor Pagamento para Vendas Repasse#1"
					
					'campos n�o sao mais obrigatorios'
					
					'///////////////////////////////////////////////////////////					
					If cor = "#FFFFCC" Then
						cor = ""
					Else
						cor = "#FFFFCC"
					End If
%>
<tr bgcolor="<%=cor%>">
  <td><div align="left"><strong><%=rs("nome_produto")%></strong></div></td>
  <td><select name='TIPO_PAGAMENTO_NORMAL_AVISTA_<%=cod_produto%>' style="width:50px" class='inputTamanho1'>

<%'Padrao
Dim selRelPadrao4
selRelPadrao4 = "F" 

If TIPO_PAGAMENTO_NORMAL_AVISTA = Empty or isNull(TIPO_PAGAMENTO_NORMAL_AVISTA) Then
TIPO_PAGAMENTO_NORMAL_AVISTA = selRelPadrao4
End if%>

<%If CStr(TIPO_PAGAMENTO_NORMAL_AVISTA) =  "F" Then
%><option selected='selected' value="F">R$</option>
<%Else%>
<option value="F">R$</option>
<%End if%>
<%If CStr(TIPO_PAGAMENTO_NORMAL_AVISTA) =  "V" Then
%><option selected='selected' value="V">%</option>
<%Else%>
<option value="V">%</option>
<%End if%>
</select></td>
  <td><input onblur="replicarValor('<%=cod_produto%>')" style="width:50px" name='VALOR_PAGAMENTO_NORMAL_AVISTA_<%=cod_produto%>' value="<%=VALOR_PAGAMENTO_NORMAL_AVISTA%>" class='inputTamanho1 padrao_valor' type='text' maxlength='200' /></td>
  <td><select style="width:50px" name='TIPO_PAGAMENTO_NORMAL_APRAZO_<%=cod_produto%>' class='inputTamanho1'>

<%'Padrao
Dim selRelPadrao6
selRelPadrao6 = "F" 

If TIPO_PAGAMENTO_NORMAL_APRAZO = Empty or isNull(TIPO_PAGAMENTO_NORMAL_APRAZO) Then
TIPO_PAGAMENTO_NORMAL_APRAZO = selRelPadrao6
End if%>

<%If CStr(TIPO_PAGAMENTO_NORMAL_APRAZO) =  "F" Then
%><option selected='selected' value="F">R$</option>
<%Else%>
<option value="F">R$</option>
<%End if%>
<%If CStr(TIPO_PAGAMENTO_NORMAL_APRAZO) =  "V" Then
%><option selected='selected' value="V">%</option>
<%Else%>
<option value="V">%</option>
<%End if%>
</select></td>
  <td><input style="width:50px" name='VALOR_PAGAMENTO_NORMAL_APRAZO_<%=cod_produto%>' value="<%=VALOR_PAGAMENTO_NORMAL_APRAZO%>" class='inputTamanho1 padrao_valor' type='text' maxlength='200' /></td>
  <td><select style="width:50px" name='TIPO_PAGAMENTO_REPASSE_AVISTA_<%=cod_produto%>' class='inputTamanho1'>

<%'Padrao
Dim selRelPadrao8
selRelPadrao8 = "F" 

If TIPO_PAGAMENTO_REPASSE_AVISTA = Empty or isNull(TIPO_PAGAMENTO_REPASSE_AVISTA) Then
TIPO_PAGAMENTO_REPASSE_AVISTA = selRelPadrao8
End if%>

<%If CStr(TIPO_PAGAMENTO_REPASSE_AVISTA) =  "F" Then
%><option selected='selected' value="F">R$</option>
<%Else%>
<option value="F">R$</option>
<%End if%>
<%If CStr(TIPO_PAGAMENTO_REPASSE_AVISTA) =  "V" Then
%><option selected='selected' value="V">%</option>
<%Else%>
<option value="V">%</option>
<%End if%>
</select></td>
  <td><input style="width:50px" name='VALOR_PAGAMENTO_REPASSE_AVISTA_<%=cod_produto%>' value="<%=VALOR_PAGAMENTO_REPASSE_AVISTA%>" class='inputTamanho1 padrao_valor' type='text' maxlength='200' /></td>
  <td><select style="width:50px" name='TIPO_PAGAMENTO_REPASSE_APRAZO_<%=cod_produto%>' class='inputTamanho1'>

<%'Padrao
Dim selRelPadrao10
selRelPadrao10 = "F" 

If TIPO_PAGAMENTO_REPASSE_APRAZO = Empty or isNull(TIPO_PAGAMENTO_REPASSE_APRAZO) Then
TIPO_PAGAMENTO_REPASSE_APRAZO = selRelPadrao10
End if%>

<%If CStr(TIPO_PAGAMENTO_REPASSE_APRAZO) =  "F" Then
%><option selected='selected' value="F">R$</option>
<%Else%>
<option value="F">R$</option>
<%End if%>
<%If CStr(TIPO_PAGAMENTO_REPASSE_APRAZO) =  "V" Then
%><option selected='selected' value="V">%</option>
<%Else%>
<option value="V">%</option>
<%End if%>
</select></td>
  <td><input style="width:50px" name='VALOR_PAGAMENTO_REPASSE_APRAZO_<%=cod_produto%>' value="<%=VALOR_PAGAMENTO_REPASSE_APRAZO%>" class='inputTamanho1 padrao_valor' type='text' maxlength='200' /></td>
</tr>

<%

					rs.MoveNext()
					Wend
					
					%>
                    <tr>
  <td colspan="9"><div align="center"><a href="javascript:salvar();"><img border="0" src="images/layout/btnGravar.jpg" /></a></div></td>
  </tr>
					</table>
				
					<%
				
				End If
			
			rs.Close
			Set rs=Nothing
FecharConexao()
%>
<input type='hidden' value='1' name='campo1'  />
</form>
<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
	
<%
'Corrige
caminho_validacao_completa = Mid(caminho_validacao_completa,2,Len(caminho_validacao_completa))
%>	
	
if(validaForm('<%=caminho_validacao_completa%>',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->


