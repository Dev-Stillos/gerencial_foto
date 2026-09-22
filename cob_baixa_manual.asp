<%
Dim calendario_necessario, BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
calendario_necessario = "NAO"
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 1
BTN_IMPRIMIR_LINK = "#"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->


<h2>COB - Baixa Manual</h2>
<br>
<!-- #include file="includes/classe_modulo_cob.asp" -->
<!-- #include file="includes/sub_menu_cob.asp" -->
<br />


<table width="800" border="0" align="center" cellpadding="5" style="border-collapse:collapse">
<tbody>
<tr>
	<td>
		<div align="center" style="border:1px solid #000;">
			<table  width="95%" style="margin:5px;">
				<tr>

					<th width="300px">Nosso Número</th>
					<th width="200px">Data Pago</th>
					<th width="200px">Valor Pago</th>

					<td rowspan="2">
						<input tabindex="4" type="button" value="Adicionar" class="btn btn-sm btn-primary" id='aaa'>
					</td>
				</tr>
				<tr>
					<td align="center"><input id='p0' tabindex="1" type="text" class="autotab inputTamanho2" name="" maxlength="8"></td>

					<td align="center"><input id='p2' tabindex="2" type="text" style="text-align:center;" maxlength="8" class="autotab inputTamanho1 padrao_data" name=""></td>

					<td align="center"><input id='p1' tabindex="3" type="text" style="text-align:center;" class="inputTamanho1 padrao_valor" name=""></td>
				</tr>
			</table>
		</div>

		<br />

		<div>
			<table width="100%" border="1" align="center" cellpadding="5" style="border-collapse:collapse" id='lista-baixas' class="hidden">
				<tr bgcolor="whitesmoke">
					<th rowspan="2">Formandos/Pedido</th>
					<th width="60px" rowspan="2">Parc.</th>
					<th colspan="2">Valor</th>
					<th colspan="2">Data</th>
					<!--<th width="70px" rowspan="2"> </th> -->
				</tr>
				<tr bgcolor="whitesmoke">				
					<th width="80px">A pagar</th>
					<th width="80px">Pago</th>
					<th width="90px">Vencimento</th>
					<th width="90px">Pagamento</th>				
				</tr>
			</table>
		</div>
	</td>
</tr>
</tbody>
</table>
 
<br />
<br />
<br />

<style>
.bloqueio_tela{
	width: 100%;
	height: 100%;
	background-color: rgba(0,0,0,0.3	);
	position: fixed;
	top: 0px;
	left: 0px;
	z-index: 1000;
	background-position:  center center;
	background-repeat: no-repeat;
	background-image: url( images/layout/preloader_grande.gif );
}
</style>

<script>
$(function(){

	$("#aaa").click(function(){
		baixaParcela($("#p0").val(), $("#p1").val(), $("#p2").val());
		exibirOcultarLista();
	});

	$("#p1").keyup(function(event){
		if ( event.which == 13 )
		{
			event.preventDefault();
			$("#aaa").trigger("click");
		}
	});

	$(".autotab").keyup(function(event){
		if ( event.which == 9 ) {
			event.preventDefault();
		}
		var a = $(this).val().replace(/\//gi,"").replace(/_/gi,"");
		if(a.length >= Number($(this).attr("maxlength")) && a != "__/__/____" )
			$("[tabindex="+(Number($(this).attr("tabindex"))+1)+"]").focus();			
	});

	function bloqueioTela(){
		$(".bloqueio_tela").removeClass("hidden");
		$("body").addClass("bloquedo");
	}

	function desbloqueioTela(){
		$(".bloqueio_tela").addClass("hidden");
		$("body").removeClass("bloquedo");
	}

	function limpaCampos()
	{
		$("input[type=text]").val("");
	}


	function baixaParcela(nossonumero, valor_pago, pagamento)
	{
		var request = $.ajax({
			url: "ajax_boletos.asp",
			type: "POST",
			data: { 
				tp : "baixarBoeltoManual", 
				p0 : nossonumero,//nosso numero
				p1 : valor_pago,//valor
				p2 : pagamento,//vencimento
			},
			beforeSend : function () {bloqueioTela();}, 
			dataType: "html"
		});
		request.done(function( msg ) {
			console.log(msg);
			x = msg.split("#");

			if(x[0].toLowerCase()    === 'true')  
			{
				alert(x[0]);
				p = x[1].split("|");
				addlinha(p[0], p[1], p[2], p[3], p[4], p[5], p[6]);
			}
			else
			{
				alert(x[1]);
			}
			limpaCampos();
		});
		request.fail(function( jqXHR, textStatus ) {
			alert(textStatus);
		});
		request.always(function(){
			$("#p0").focus();
			desbloqueioTela();
			limpacampos();

		});
	}

	function addlinha(nome, pedido, parcela, valor, valor_pago, vencimento, pagamento)
	{
		html =  '<tr>';
		html += '<td>'+nome+'<br><strong>Pedido:</strong>'+pedido+'</td>';
		html += '<td align="center">'+parcela+'</td>';
		html += '<td align="right">R$ '+valor+'</td>';
		html += '<td align="right">R$ '+valor_pago+'</td>';
		html += '<td align="center">'+vencimento+'</td>';
		html += '<td align="center">'+pagamento+'</td>';
		//html += '<td align="center">desfazer</td>';
		html += '</tr>';

		$("#lista-baixas").append(html);
	}

	function removerLinha()
	{

	}

	function exibirOcultarLista()
	{
		var total = $("#lista-baixas tr").size();

		if (total > 2)
		{
			$("#lista-baixas").removeClass("hidden");
		}
		else
		{
			$("#lista-baixas").addClass("hidden");
		}
	}
});

</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->


</body>
</html>