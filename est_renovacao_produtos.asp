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
calendario_necessario = "N&AtildeO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%

'----------------------------------------------------------------------------------------

'Declara vari&aacuteveis para utiliza&ccedil&atildeo do sistema
Dim cod_movimentacao, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

actionformulario = caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.Form("comando")

'Declara vari&aacuteveis
Dim tipo, cod_produto, quantidade, descricao

If comando = "renovacao" and campo1 = "1" Then
	call renovacao()
End if

'Fun&ccedil&otildees
Sub renovacao()
	
	dim nome_vendedor, rs
	dim qtd_dias
	dim movimentacoes_renovacao
	'lemos o codigo do vendedor'
	cod_vendedor = Request.form("cod_vendedor")
	movimentacoes_renovacao = Request.form("movimentacoes_renovacao")
	movimentacoes_renovacao = split(movimentacoes_renovacao,",")

	abrirConexao()

	set rs = objConexao.execute("SELECT nome_vendedor FROM TB_VENDEDORES WHERE cod_vendedor = "&sqlNumero(cod_vendedor) )
	if not rs.Eof then
		nome_vendedor = rs("nome_vendedor")
	end if
	rs.Close
	set rs = nothing	
	fecharConexao()

	abrirConexao()

	for each r in movimentacoes_renovacao

		qtd_dias = request.form("renovacao_dias_"&triM(r))
		call estoqueMovimentacaoVendedorRenovacao(triM(r), qtd_dias, true)
		
	next
	
	fecharConexao()
	
End sub

'----------------------------------------------------------------------------------------
%>

<h2><%=titulo_pagina%></h2>
<br />
<div class="subMenu">
  Sub-Menu: 
	  <a href="epv_devolucao.asp?comando=nova">Iniciar devolu&ccedil&atildeo de Vendedor</a> | 
	  <a href="epv_saida_inicio.asp?comando=nova">Sa&iacuteda de &aacutelbuns do Estoque</a> | 
	  <a href="epv_saida_fornecedores.asp">Sa&iacuteda do Estoque para Fornecedor</a> <br />
	  <a href="epv_saida_produtos.asp">Sa&iacuteda de itens do Estoque</a> | 
	  <a href="epv_devolucao_produtos.asp">Devolu&ccedil&atildeo de itens para o Estoque</a>
  </div>
  <br />
<%If comando = "excluir" Then%>
	<div class='msgExcluir'><%=msgAoExcluir%></div>
<%End If%>

<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post'>
	<fieldset>
		<legend>Informa&ccedil&otildees Gerais</legend>
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
			  data: { p0 : $("#cod_vendedor_id").val(), p1:"R" , t:"lista_produtos"},
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



