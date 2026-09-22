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
<%

'----------------------------------------------------------------------------------------

'Declara vari�veis para utiliza��o do sistema
Dim cod_movimentacao, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario, sucesso

actionformulario = caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.Form("comando")
sucesso = "N"
'Declara vari�veis
Dim tipo, cod_produto, quantidade, descricao

If comando = "saida" and campo1 = "1" Then
	call saida()
End if

'Fun��es
Sub saida()
	dim nome_vendedor, rs
	dim qtd_p
	'lemos o codigo do vendedor'
	cod_vendedor = Request.form("cod_vendedor")
	produtos = Request.form("produtos")
	produtos = split(produtos,",")

	abrirConexao()
	set rs = objConexao.execute("SELECT nome_vendedor FROM TB_VENDEDORES WHERE cod_vendedor = "&sqlNumero(cod_vendedor) )
	if not rs.Eof then
		nome_vendedor = rs("nome_vendedor")
	end if
	rs.Close
	set rs = nothing	
	fecharConexao()

	for each p in produtos
		qtd_p = Request.form("produto_"&trim(p))
		'Saida do estoque interno
		call estoqueMovimentacao("S", "vendedor", cod_vendedor, p, qtd_p, "Retirada de álbum vendedor: "&nome_vendedor, false)
		'log aqui

		'Entrada no estoque do  vendedor
		call estoqueMovimentacaoVendedor("E", cod_vendedor, p, qtd_p, 0, "Retirada de álbum vendedor: "&nome_vendedor, false)
		'Log aqui

		'Gerar recibo
	next
	sucesso = "S"
End sub


'----------------------------------------------------------------------------------------
%>

<h2><%=titulo_pagina%></h2>
<br />
<div class="subMenu">
  Sub-Menu: 
	  <a href="epv_devolucao.asp?comando=nova">Iniciar devolu&ccedil;&atilde;o de Vendedor</a> | 
	  <a href="epv_saida_inicio.asp?comando=nova">Sa&iacute;da de &Aacute;lbuns do Estoque</a> | 
	  <a href="epv_saida_fornecedores.asp">Saída do Estoque para Fornecedor</a> <br />
	  <a href="epv_saida_produtos.asp">Saída de itens do Estoque</a> | 
	  <a href="epv_devolucao_produtos.asp">Devolução de itens para o Estoque</a>
  </div>
  <br />
<%If sucesso = "S" Then%>
	<div class='mggGrandeSucesso'>
		Retirada efetuada com sucesso! 
	</div>
<%End If%>

<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post'>
	<fieldset>
		<legend>Informações Gerais</legend>
		<table width='90%' border='0' align='center' cellpadding='5' style='border-collapse:collapse'>
			<tr>
				<td align="center">
					<div align="center"><h3>Vendedor</h3></div>
				</td>
			</tr>
			<tr>
				<td align="center">
					<%=padraoFormularioSelect("cod_vendedor","","inputTamanho2", "", "","Selecione", cod_vendedor, "SELECT cod_vendedor, nome_vendedor FROM TB_VENDEDORES where ativado = 1 order by nome_vendedor", "incluir", false, false, false) %>
				</td>
			</tr>
		</table>
	</fieldset>
	<div id='produtosInc' style="display:none">
		<fieldset>
			<legend>Produtos</legend>
			<table width="100%" border="0
			" align="center" cellpadding="5" style="border-collapse:collapse;border:#999999 dashed 1px; padding:2px; text-align:center; background:#FFFFCC; width: 800px;">
				<tr>
					<td align="center"><strong>Produto</strong></td>
					<td align="center"><strong>Qunatidade</strong></td>
					<td align="center" rowspan="2">
						<img src="images/layout/btnIncluir.jpg" id='adicionar_produto' class='padrao_link' />
					</td>
				</tr>
				<tr>
					<td align="center">
						<%=padraoFormularioSelect("cod_produto","cod_produto_id","inputTamanho3", "", "","Selecione", cod_produto, "SELECT cod_produto, NOME_PRODUTO FROM TB_PRODUTOS where ativo = 1 and cod_categoria = 6", "incluir", false, false, false) %>
					</td>
					<td align="center">
						<input type="text" class="padrao_numero inputTamanho1" id="quantidade_id">		
					</td>
				</tr>
			</table>

			<div id="lista" style="margin-top:10px;">
				<table width="90%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
					<tr>
						<td align="center"><strong>Produto</strong></td>
						<td align="center"><strong>Quantidade</strong></td>
						<td align="center"><strong></strong></td>
					</tr>
					<tr id='aviso_vazio'>
						<td colspan="3" align="center"> 
							<strong>Sem produtos selecionados</strong>
						</td>
					</tr>
				</table>
			</div>
		</fieldset>
		<table width="90%" border="0" align="center" cellpadding="5" style="border-collapse:collapse">
			<tr>
				<td align="center">
			        <input class="btnGravar" type="submit" name="button" id="button" value=" " />   
				</td>
			</tr>
		</table>
	</div>

	<input type="hidden" name="campo1" value="1">
	<input type="hidden" name="comando" value="saida">
</form>

<script type="text/javascript">
	$(function(){
		$("select[name=COD_VENDEDOR]").val("");
		function mostrarAviso()
		{
			if($("#lista > table > tbody > tr").length > 2)
				$("#aviso_vazio").css("display", "none");
			else
				$("#aviso_vazio").css("display", "table-row");

			return true
		}
		
		$("select[name=COD_VENDEDOR]").change(function(){
			if($("select[name=COD_VENDEDOR]").val() == "")
				$("#produtosInc").css("display", "none");
			else
				$("#produtosInc").css("display", "block");

		});

		$(document).on("click", ".remover_produto", function() {		
			$("#tr_"+$(this).attr('cp')).remove();
			mostrarAviso()
		});

 		$("#adicionar_produto").click(function(){
 		//verificar se o produto j� est� na lista se estiver exibe alerta

 		var cp = $("#cod_produto_id").val();
 		var NOME_PRODUTO = $("#cod_produto_id option[value="+cp+"]").text();
 		var qtd = $("#quantidade_id").val();

 		if( $("#tr_"+cp).length > 0)
 		{
 			alert('O produto selecionado já está na lista!');
 		}
 		else //se NÃO existir
 		{
 			var ultimo = $("#lista > table > tbody:nth-last-child(1)");

 			var html =  "<tr id='tr_"+cp+"'>"+
 							"<td>"+NOME_PRODUTO+"<input type='hidden' value='"+cp+"' name='produtos' /></td>"+
 							"<td>"+qtd+"<input type='hidden' value='"+qtd+"' name='produto_"+cp+"' /></td>"+
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
 			mostrarAviso()
		}

 		});
	});
</script>

<script type='text/javascript'>
<%
If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('TIPO#Tipo#2|COD_PRODUTO#Produto#2|QUANTIDADE#Quantidade#1|DESCRICAO#Descrição#1',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
	return true
}
<%End If%>
</script>
<!-- #include file="files/includes/miolo-inferior.asp" -->