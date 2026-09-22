<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 1
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
  Dim titulo_pagina
  titulo_pagina = "Produtos"
%>
<%
Dim calendario_necessario
calendario_necessario = "NÃO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<!-- #include file="includes/sub_menu_mck.asp" -->
<%

'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 10/10/2011 16:04:06

'Declara vari�veis para utiliza��o do sistema
Dim cod_mack_pedido, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

redirecionar_incluir	= "mck_pedidos.asp"
redirecionar_alterar	= "mck_pedidos.asp"
redirecionar_excluir	= "mck_pedidos.asp"
redirecionar_validacao	= "mck_pedidos.asp"
msgAoExcluir			= "Tem certeza que deseja excluir este formando?"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
'Declara vari�veis
Dim  cod_mack_formando, cod_mack_forma_pagamento, data_cadastro, cod_usuario, nome, tia, descricao_colacao, curso, email_1, qtd_parcelas, cpf, obs

If comando = Empty Then
comando = "incluir"
titulo_pagina = titulo_pagina&" / Incluir"
End if

If comando = "incluir" and campo1 <> Empty Then
	'Incluir novo registro
	incluirRegistro(redirecionar_incluir)
	titulo_pagina = titulo_pagina&" / Inclusão"
Elseif comando = "alterar" and campo1 = Empty Then
	'Leitura inicial para possível alteração
	cod_mack_pedido = Request.QueryString("cod_mack_pedido")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / alteração"
Elseif 	comando = "alterar" and campo1 <> Empty Then
	'Gravar alteração
	cod_mack_pedido = Request.Form("cod_mack_pedido")
	alterarRegistro(redirecionar_alterar)
	titulo_pagina = titulo_pagina&" / alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 
	'Leitura inicial para possível Exclusão
	cod_mack_pedido = Request.QueryString("cod_mack_pedido")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
	'Exclui registro
	cod_mack_pedido = Request.Form("cod_mack_pedido")
	excluirRegistro(redirecionar_excluir)
	titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
	'Visualiza registro
	cod_mack_pedido = Request.QueryString("cod_mack_pedido")
	lerRegistro()
	titulo_pagina = titulo_pagina&" / Visualização"
End If


'Fun��es

Sub lerFormulario()

	cod_mack_pedido = Request.Form("cod_mack_pedido")
	cod_mack_formando = Request.Form("cod_mack_formando_pedido")
	cod_mack_forma_pagamento  = Request.Form("cod_mack_forma_pagamento")
	qtd_parcelas = Request.Form("qtd_parcelas")
	cpf = Request.Form("cpf")
	obs = Request.Form("obs")

End sub


Sub lerRegistro()
	
	dim rs

	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM VW_TB_MACK_PEDIDOS WHERE cod_mack_pedido = "&cod_mack_pedido)
		if not rs.eof then
			cod_mack_pedido = rs("cod_mack_pedido")
			cod_mack_formando = rs("cod_mack_formando")
			cod_mack_forma_pagamento = rs("cod_mack_forma_pagamento")
			data_cadastro = rs("data_cadastro")
			cod_usuario = rs("cod_usuario")
			nome = rs("nome")
			tia = rs("tia")
			curso  = rs("curso")
			email_1 = rs("email_1")
			descricao_colacao = rs("descricao_colacao")
			qtd_parcelas = rs("qtd_parcelas")
			cpf = rs("cpf")
			obs = rs("obs")
		Else
			response.write "formando NÃO encontrado"
			response.end
		end if
		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)

	lerFormulario()

	AbrirConexao()

		if retirou_convites = "S" then
			data_retirou_convites = NOW()
		else
			data_retirou_convites = null
		end if

		'response.write "INSERT INTO TB_MACK_PEDIDOS (cod_mack_formando, cod_mack_forma_pagamento,  cod_usuario, qtd_parcelas, cpf, obs) VALUES ("&sqlNumero(cod_mack_formando)&", "&sqlNumero(cod_mack_forma_pagamento)&", "&sqlNumero(session("cod_usuario"))&", "&sqlNumero(qtd_parcelas)&", "&sqlTexto(cpf)&", "&sqltexto(obs)&" ) "
		dim desconto

		Set rs = objConexao.Execute("SELECT * FROM TB_MACK_FORMAS_PAGAMENTO  WHERE cod_mack_forma_pagamento = "&sqlNumero(cod_mack_forma_pagamento))
		if not rs.Eof  then
			desconto = formatNumber(rs("desconto_avista"))/1
		end if
		rs.Close
		set rs = Nothing

		objConexao.Execute("INSERT INTO TB_MACK_PEDIDOS (cod_mack_formando, cod_mack_forma_pagamento,  cod_usuario, qtd_parcelas, cpf, obs) VALUES ("&sqlNumero(cod_mack_formando)&", "&sqlNumero(cod_mack_forma_pagamento)&", "&sqlNumero(session("cod_usuario"))&", "&sqlNumero(qtd_parcelas)&", "&sqlTexto(cpf)&", "&sqltexto(obs)&" ) ")

		'Encontrar chave criada
		Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_MACK_PEDIDOS LIMIT 1")
			If not rs.Eof Then
				cod_mack_pedido = rs("ULTIMO_ID")
			End if 
		rs.Close
		Set rs=Nothing

		dim produtos, valor_prod
		produtos = request.form("Produtos")

		produtos = split(produtos, ";")

		for x = 0 to UBound(produtos)

			'buscamos o valor do produto na base
			set rs = objConexao.execute("SELECT valor from TB_MACK_PRODUTOS WHERE cod_mack_produto = "&sqlNumero(Produtos(i)))
			if not rs.Eof then
				valor_prod = formatNumber(rs(0))
			end If
			rs.Close
			set rs = Nothing

			if desconto > 0 and qtd_parcelas = 1 then
				valor_prod =valor_prod * (100-desconto)/100
			end if 


			objConexao.Execute("insert into TB_MACK_PEDIDOS_PRODUTOS (cod_mack_produto, valor_unitario, quantidade, cod_mack_pedido ) VALUES ("&sqlNumero(Produtos(i))&", "&sqlNumero(valor_prod)&", "&sqlNumero(request.form("produto_"&Produtos(i)))&", "&sqlNumero(cod_mack_pedido)&") ")

		next

	FecharConexao()
	Response.Redirect(redirecionar&"?cod_mack_pedido="&cod_mack_pedido&"&codmsg=1")
End sub

Sub alterarRegistro(redirecionar)
	
	lerFormulario()

	AbrirConexao()

		'if retirou_convites = "S" then
		''	data_retirou_convites = NOW()
		'else
		''	data_retirou_convites = null
		'end if

		'objConexao.execute("")

	FecharConexao()

	Response.Redirect(redirecionar&"?cod_mack_pedido="&cod_mack_pedido&"&codmsg=2")
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		if  estoqueVerificaManual(cod_mack_pedido, true) then
			objConexao.Execute("DELETE FROM TB_MACK_FORMANDOS WHERE cod_mack_pedido = "&cod_mack_pedido)
		end if
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_mack_pedido="&cod_mack_pedido&"&codmsg=3")
End sub

'----------------------------------------------------------------------------------------
%>

<h2><%=titulo_pagina%></h2><br />

<%If comando = "excluir" Then%>
	<div class='msgExcluir'><%=msgAoExcluir%></div>
<%End If%>

<style type="text/css">
	fieldset{
		width: 820px;
		margin:10px auto 5px auto ;
	}
</style>


<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post'>
	
	<fieldset>
		<legend>Dados do Pedido</legend>

		<% if comando = "incluir" then%>
			
			<table width='800px' border='1' align='center' cellpadding='5' style='border-collapse:collapse; margin-bottom: 10px;' id='table_busca'>
			<tr>
				<td width="50%">
					<table width="100%">
						<tr>
							<td>
								<h2 width='100%' style="text-align:center;">Código do Formando</h2>
							</td>
						</tr>
						<tr>
							<td align="center">
								<input type="text" class="inputTamanho" style="width:80%" id="cod_mack_formando">								
							</td>
						</tr>
					</table>
				</td>
				<td width="50%">
					<table width="100%">
						<tr>
							<td>
								<h2 width='100%' style="text-align:center;">TIA do formando</h2>
							</td>
						</tr>
						<tr>
							<td align="center">
								<input type="text" class="inputTamanho" style="width:80%" id="tia_formando">								
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td colspan="2" align="center">
					
					<button id='buscar_formando' class="btn btn-success">
						Buscar Formando
					</button>
				</td>
			</tr>
			</table>

		<% end if %>
		<table width='800px' id='table_formando' border='1' align='center' cellpadding='5' style='border-collapse:collapse; <% if comando = "incluir" then %> display: none;	<%end if%>'>
		<tr>
		<td width='27%'><strong>Cód. Formando </strong></td>
			<td width='73%'>
				<div align='left'>
				
					<span class='textoMaior'><strong id='ajax_cod'><%=cod_mack_formando%></strong></span>
				
				</div>
			</td>
		</tr>
		<tr>
			<td width='27%'><strong>Nome *</strong></td>
			<td width='73%'>
				<div align='left' id='ajax_nome'>
					<%=nome%>
				</div>
			</td>
		</tr>
		<tr>
			<td width='27%'><strong>Tia *</strong></td>
			<td width='73%'>
				<div align='left' id='ajax_tia'>
					<%=tia%>
				</div>
			</td>
		</tr>
		<tr>
			<td width='27%'><strong>Curso *</strong></td>
			<td width='73%'>
				<div align='left' id='ajax_curso'>
				<%=curso%>
				</div>
			</td>
		</tr>
		<tr>
			<td width='27%'><strong>CPF *</strong></td>
			<td width='73%'>
				<div align='left'>
				<%If comando <> "visualizar" and comando <> "excluir"  Then%>
					<input name='CPF' class='inputTamanho2 padrao_cpf' value="<%=CPF%>"  type='text' />
				<%
				Else
					If not isNull(cpf) Then
						response.write(cpf)
					End if
				End If %>
				</div>
			</td>
		</tr>
	
		<tr>
			<td width='27%'><strong>Forma de pagamento *</strong></td>
			<td width='73%'>
				<div align='left'>
				<%=padraoFormularioSelect("cod_mack_forma_pagamento", "", "inputTamanho2", "", "", "Selecione", cod_mack_forma_pagamento, "SELECT cod_mack_forma_pagamento, nome FROM TB_MACK_FORMAS_PAGAMENTO", comando, false, false, false)%>
				<em>
				<% 
				if retirou_convites = "S" then
					response.write(" Retirou em:"&arrumaData(data_retirou_convites," - "))
				end if
				%>
				</em>
				<input type="hidden" value="<%=data_retirou_convites%>" name="data_retirou_convites">
				</div>
			</td>
		</tr>

		<tr>
			<td width='27%'><strong>Parcelas *</strong></td>
			<td width='73%'>
				<div align='left'>
				<%=padraoFormularioSelect("qtd_parcelas","qtd_parcelas_id","inputTamanho3", "", "","SEM_SELECIONE", qtd_parcelas, "PADRAO:1#1|2#2", "incluir", false, false, false) %>
				</div>
			</td>
		</tr>
		<tr>
			<td width='27%'><strong>Observação *</strong></td>
			<td width='73%'>
				<div align='left'>
				<%If comando <> "visualizar" and comando <> "excluir"  Then%>
					<textarea name="obs" id="id_obs" class="inputTamanho2" rows='5'><%=obs%></textarea>
				<%
				Else
					If not isNull(obs) Then
						response.write(obs)
					End if
				End If 
				%>
				</div>
			</td>
		</tr>
	</table>
	
</fieldset>
	<div id='produtosInc'>
		<fieldset id='fieldset_prod' <% if comando = "incluir" then %> style='display: none;'	<%end if%>>
			<legend>Produtos</legend>
			
			<% if comando = "incluir" then  %>
			<table width="100%" border="0
			" align="center" cellpadding="3" style="border-collapse:collapse;border:#999999 dashed 1px; padding:2px; text-align:center; width: 800px;">
				<tr>
					<td align="center"><strong>Produto</strong></td>
					<td align="center"><strong>Quantidade</strong></td>
					<td align="center" rowspan="2">
						<img src="images/layout/btnIncluir.jpg" id='adicionar_produto' class='padrao_link' />
					</td>
				</tr>
				<tr>
					<td align="center">
						<%=padraoFormularioSelect("cod_mack_produto","cod_mack_produto_id","inputTamanho3", "", "","Selecione", cod_produto, "SELECT cod_mack_produto, NOME_PRODUTO FROM TB_MACK_PRODUTOS", "incluir", false, false, false) %>
					</td>
					<td align="center">
						<input type="text" class="padrao_numero inputTamanho1" id="quantidade_id">		
					</td>
				</tr>
			</table>
			<% end if %>
			<div id="lista" style="margin-top:10px;">
				<table width="800px" border="1" align="center" cellpadding="3" style="border-collapse:collapse">
					<tr>
						<td align="center"><strong>Produto</strong></td>
						<td align="center"><strong>Quantidade</strong></td>
						<td align="center"><strong>Valor Un.</strong></td>
						<td align="center"><strong>Valor total</strong></td>
						<td align="center"><strong></strong></td>
					</tr>
					<%
					if comando <> "incluir" then 
						AbrirConexao()
						set rs = objConexao.execute("SELECT *, quantidade * valor_unitario as valor_total FROM VW_TB_MACK_PEDIDOS_PRODUTOS WHERE cod_mack_pedido = "&sqlNumero(cod_mack_pedido))
						while not rs.Eof 
							%>
							<tr>
								<td align="center"><strong><%=rs("NOME_PRODUTO")%></strong></td>
								<td align="center"><strong><%=rs("quantidade")%></strong></td>
								<td align="center"><strong><%=rs("valor_unitario")%></strong></td>
								<td align="center"><strong><%=rs("valor_total")%></strong></td>
								<td align="center"><strong></strong></td>
							</tr>
							<%
						rs.movenext
						Wend
						rs.Close
						set rs = Nothing
						FecharConexao()
					else
					%>
					<tr id='aviso_vazio'>
						<td colspan="5" align="center"> 
							<strong>Sem produtos selecionados</strong>
						</td>
					</tr>
					<%end if%>
				</table>
			</div>
		</fieldset>
	</div>
	<% if comando = "incluir" then  %>
	<table width="90%" border="0" align="center" cellpadding="5" style='border-collapse:collapse; <% if comando = "incluir" then %> display: none;<%end if%>' id='table_gravar'>
			<tr>
				<td align="center">
			        <input class="btn btn-warning" type="submit" name="button" id="button" value="Finalizar Pedido" />   
				</td>
			</tr>
		</table>	
	<%end if%>
<input type='hidden' value='1' name='campo1'  />
<input type='hidden' value='<%=cod_mack_pedido%>' name='cod_mack_pedido'  />
<input type='hidden' value='<%=cod_mack_formando%>' name='cod_mack_formando_pedido' id='cod_mack_formando_pedido_id'  />
</form>



<script type="text/javascript">
	$(function(){
		function mostrarAviso()
		{
			if($("#lista > table > tbody > tr").length > 2)
				$("#aviso_vazio").css("display", "none");
			else
				$("#aviso_vazio").css("display", "table-row");

			return true;
		}

		$("#buscar_formando").click(function(event){
			event.preventDefault();

			if($("#cod_mack_formando").val() != '')
			{
				buscarFormando('1', $("#cod_mack_formando").val());
			}
			else
			{
				buscarFormando('2', $("#tia_formando").val());
			}
		});


		function buscarFormando(tipo, valor)
		{
			var request = $.ajax({
				url: "mck_formandos_formulario_ajax.asp",
				method: "POST",
				data: { t : "buscaFormando", p0 : tipo, p1 : valor },
				dataType: "html"
			});

			request.done(function( msg ) {

				if (msg.length > 0 ) {
				var retorno = msg.split(';');
					$("#ajax_nome").html(retorno[0]);
					$("#ajax_tia").html(retorno[1]);
					$("#ajax_curso").html(retorno[2]);
					$("#ajax_cod").html(retorno[3]);
					$("#cod_mack_formando_pedido_id").val(retorno[3]);




					$("#table_gravar").css("display", "table");
					$("#table_formando").css("display", "table");
					$("#fieldset_prod").css("display", "block");
					$("#table_busca").css("display", "none");
				}
				else
					alert('Formando NÃO encontrado');
			});

			request.fail(function( jqXHR, textStatus ) {
				alert( "Houve um erro ao buscar valor do produto. Informe o administrador do sistema" );
			});
		}

		$(document).on("click", ".remover_produto", function() {		
			$("#tr_"+$(this).attr('cp')).remove();
			mostrarAviso()
		});

 		$("#adicionar_produto").click(function(){
 		//verificar se o produto j� est� na lista se estiver exibe alerta

	 		var cp = $("#cod_mack_produto_id").val();
	 		var NOME_PRODUTO = $("#cod_mack_produto_id option[value="+cp+"]").text();
	 		var qtd = $("#quantidade_id").val();

	 		if( $("#tr_"+cp).length > 0)
	 		{
	 			alert('O produto selecionado já está na lista!');
	 		}
	 		else //se NÃO existir
	 		{
	 			var menuId = $( "ul.nav" ).first().attr( "id" );
				var request2 = $.ajax({
					url: "mck_formandos_formulario_ajax.asp",
					method: "POST",
					data: { t : "valorProduto", p0 : $("#cod_mack_produto_id").val() },
					dataType: "html"
				});

				request2.done(function( msg ) {
					var valor_un, valor_t, temp

					temp = msg;
					valor_un = temp;
					valor_t = temp;


					var ultimo = $("#lista > table > tbody:nth-last-child(1)");

		 			var html =  "<tr id='tr_"+cp+"'>"+
		 							"<td>"+NOME_PRODUTO+"<input type='hidden' value='"+cp+"' name='produtos' /></td>"+
		 							"<td  align='right'>"+qtd+"<input type='hidden' value='"+qtd+"' name='produto_"+cp+"' /></td>"+
		 							"<td align='right'>R$ "+valor_un+"</td>"+
		 							"<td align='right'>R$ "+valor_t*qtd+"</td>"+
		 							"<td>"+
		 								"<table width='100%'>"+
											"<tr>"+
												"<td align='center'>"+
													"<img src='images/icons/icon_excluir.png' alt='excluir Registro' title='excluir Registro' border='0' class='remover_produto' cp='"+cp+"' />"+
												"</td>"+
											"</tr>"+
										"</table>"+
		 							"</td>"+
		 						"</tr>";
		 			ultimo.append(html);	
		 			mostrarAviso();
				});

				request.fail(function( jqXHR, textStatus ) {
					alert( "Houve um erro ao buscar valor do produto. Informe o administrador do sistema" );
				});
			}
 		});
	});
</script>
<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>
<!-- #include file="files/includes/miolo-inferior.asp" -->