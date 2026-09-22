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
Dim cod_movimentacao, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

actionformulario = caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.Form("comando")
'Declara vari�veis
Dim tipo, cod_produto, quantidade, descricao, sucesso
sucesso = "N"

If comando = "devolucao" and campo1 = "1" Then
	call Devolucao()

End if

'Fun��es
Sub devolucao()
	
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

		'Entrada no estoque do  vendedor
		call estoqueMovimentacaoVendedor("S", cod_vendedor, p, qtd_p, 0, "Devolvido pelo vendedor: "&nome_vendedor, false)

		'Saida do estoque intern
		call estoqueMovimentacao("E", "vendedor", cod_vendedor, p, qtd_p, "Devolvido pelo vendedor: "&nome_vendedor, false)
		'log aqui

		'Gerar recibo ?
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
		Devolução efetuada com sucesso! 
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
					<%=padraoFormularioSelect("cod_vendedor","cod_vendedor_id","inputTamanho2", "", "","Selecione", cod_vendedor, "SELECT cod_vendedor, nome_vendedor FROM TB_VENDEDORES where ativado = 1 order by nome_vendedor", "incluir", false, false, false) %>
				</td>
			</tr>
		</table>
	</fieldset>
	<div id='container_pordutos' style="display: hidden;">
		<fieldset>
			<legend>Produtos Pendentes</legend>
			<div id="produtos_pendentes" style="margin-top:10px;"></div>
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
	<input type="hidden" name="comando" value="devolucao">
</form>

<script type="text/javascript">
	function carregandoRequisicao ()
	{
		$("#container_pordutos").css("display","hidden");
	}

	function requisicaoCarregada (){
		$("#container_pordutos").css("display","block");
	}

	$("#cod_vendedor_id").change(function(){
		if($("#cod_vendedor_id").val() != "")
		{
			var request = $.ajax({
			  url: "epv_devolucao_produtos_ajax.asp",
			  method: "POST",
			  data: { p0 : $("#cod_vendedor_id").val(), p1:"D" , t:"lista_produtos"},
			  dataType: "html",
			  beforesend: carregandoRequisicao()
			});

			request.done(function( msg ) {
				requisicaoCarregada ();
			  	$( "#produtos_pendentes" ).html( msg );
			});
			 
			request.fail(function( jqXHR, textStatus ) {
				requisicaoCarregada ();
			});
		}
	});
</script>
<!-- #include file="files/includes/miolo-inferior.asp" -->