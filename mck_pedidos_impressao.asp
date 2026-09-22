<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%

abrirConexao()



Dim cod_mack_pedido
cod_mack_pedido = request.Querystring("c")


Dim  cod_mack_formando, cod_mack_forma_pagamento, data_cadastro, cod_usuario, qtd_parcelas, cpf, obs, nome, tia, curso,    descricao_colacao, email_1, endereco, cep, bairro, cidade, estado, forma_pagamento, ano_conclusao

Set rs = objConexao.Execute("SELECT * FROM VW_TB_MACK_PEDIDOS where cod_mack_pedido = "&sqlNumero(cod_mack_pedido))
if not rs.Eof then

	cod_mack_formando = rs("cod_mack_formando")
	cod_mack_forma_pagamento = rs("cod_mack_forma_pagamento")
	data_cadastro = rs("data_cadastro")
	cod_usuario = rs("cod_usuario")
	qtd_parcelas = rs("qtd_parcelas")
	cpf = rs("cpf")
	obs = rs("obs")
	nome = rs("nome")
	curso = rs("curso")
	ano_conclusao = rs("ano_conclusao")
	forma_pagamento = rs("forma_pagamento")
	email_1 = rs("email_1")
	endereco = rs("endereco")
	cep = rs("cep")
	bairro = rs("bairro")
	cidade = rs("cidade")
	estado = rs("estado")

end if
rs.close
set rs = nothing


function completa(str, caracter, quantidade, inicio_fim)

	dim retorno, i
	retorno = str

	if inicio_fim = "i" then
		while len(retorno) < quantidade
			retorno = caracter & retorno 
		wend
	else
		while len(retorno) < quantidade
			retorno = retorno & caracter
		wend
	end if

completa = retorno
end function



%>
<html>
<head>
	<title>Pedido </title>
	<style type="text/css">
	body {font-family:arial;}
	table{font-size: 12px; width: 99%; margin-bottom: 0px;}
	table th{font-size: 14px;}
	table tr{text-align: center;}
	table tr td, table tr th{padding: 3px;}
	table tr td h3{border-bottom: 1px solid #000; text-align: left; margin: 10px 0px 3px 0px; padding: 0px;}
	.tableCollapse{border-collapse: collapse;}
	fieldset {margin: 8px 0 8	px 0; color: #000; border: #000 solid 1px; }
	legend {background-color: #fff;}
	h4, h2{margin: 5px;}

	.alerta_tela{
		width: 100%;
		border: 2px dashed red;
		height: 40px;
		line-height: 38px;
		text-align: center;
		color: black;
		font-weight: bold;
		font-size: 16px;
		background-color: #FFE7E8;	
	}
@media print{
	.alerta_tela{
		display: none;
	}
}


	
	</style>

	<script>
	function imprime(){
		window.print();
	}
	</script>

</head>
<body>

	<div class='alerta_tela'>
		ATENÇÃO IMPRIMIR 2 VIAS <button onclick="imprime()"> IMPRIMIR </button>
	</div>


	<table border="0">
		<tr>
			<td width="100px"><img src="images/layout/logomarca_foto_reportagens.jpg" /></td>
			<td align="left">
				<strong>Stillo's Reportagens Fotográficas</strong>
				<br /> 
				
				Rua General Calado, 52 <br />
				Tatuapé - São Paulo, SP<br />
				Fone: 2672-9899 Cep: 03323-020<br />

			</td>
			<td style="border:1px solid #000;">
				<h1>Pedido de Venda
					<br>N°: 1215<%= completa(cod_mack_pedido, "0", 4, "i")%>   
				</h1>
				
			</td>
		</tr>
	</table>
	<fieldset>
		<legend>Resumo do Pedido</legend>
		<table border="0">
		<tr>
			<td align="right" width='160px'><strong>Vendedor:</strong></td>
			<td align="left" collspan="3"><%=nomeUsuario(session("cod_usuario"),true)%></td>
			<td align="right"><strong>Data do Pedido:</strong></td>
			<td align="left" ><%=data_cadastro%></td>
		</tr>
		<tr>
			<td align="right"><strong>Formade Pagamento:</strong></td>
			<td align="left" collspan="3"><%=forma_pagamento%></td>
			<td align="right"><strong>N° de Parcelas:</strong></td>
			<td align="left"><%=qtd_parcelas%> x</td>
		</tr>
		<tr>
			<td align="right"><strong>Observações:</strong></td>
			<td align="left" collspan="5">
				<%
				if obs <> "" then
					response.write(replace(obs,chr(13),"<br />"))
				else
					response.write("-")
				end if
				%>
			</td>
		</tr>
		<tr>
			<td align="left" colspan="6" style="padding-left:110px;"></td>
		</tr>
	</table>
	</fieldset>
	<fieldset>
		<legend>Informações do Cliente</legend>
		<table border="0">
			<tr>
				<td align="right" width="100px"><strong>Nome:</strong></td>
				<td align="left" collspan="3"><%=nome%></td>
				<td align="right"><strong>CPF:</strong></td>
				<td align="left"><%=cpf%></td>
			</tr>
			<tr>
				<td align="right"><strong>E-mail:</strong></td>
				<td align="left" collspan="3"><%=email_1%></td>
			</tr>
			<tr>
				<td align="right"><strong>Endereço:</strong></td>
				<td align="left" collspan="2"><%=endereco%></td>
				<td align="right"><strong>Bairro:</strong></td>
				<td align="left" collspan="2"><%=bairro%></td>
			</tr>
			<tr>
				<td align="right"><strong>Cidade/Estado:</strong></td>
				<td align="left" collspan="3"><%=cidade%> - <%=estado%></td>
				<td align="right"><strong>CEP:</strong></td>
				<td align="left"><%=cep%></td>
			</tr>
			<tr>
				<td align="right"><strong>Curso:</strong></td>
				<td align="left" collspan="3"><%=curso%></td>
				<td align="right"><strong>Conclusão:</strong></td>
				<td align="left"><%=ano_conclusao%></td>
			</tr>
		</table>
	</fieldset>
	<%


	set rs = objConexao.execute("SELECT *, quantidade * valor_unitario as valor_total FROM VW_TB_MACK_PEDIDOS_PRODUTOS WHERE cod_mack_pedido = "&sqlNumero(cod_mack_pedido))
	if not rs.Eof then	
	%>
	<fieldset>
		<legend>Produtos</legend>
		<table border="0" >
		
		<tr>
			<th>#</th>
			<th>Produto</th>
			<th>Quantidade</th>
			<th>Valor unitário</th>
			<th>Valor Total</th>
		</tr>
		<% While not rs.eof  %>
		<tr>
			<td>1</td>
			<td align='left'><%=rs("nome_produto")%></td>
			<td><%=rs("quantidade")%></td>
			<td align="right">R$ <%=arrumaNumero(rs("valor_unitario"),0,2)%></td>
			<td align="right">R$ <%=arrumaNumero(rs("valor_total"),0,2)%></td>
		</tr>
		<% 
		rs.movenext
		Wend
		%>
		<tr>
			<td colspan="5" align="right">
				<hr />
				<%
				Set rs2 = objConexao.execute ("SELEct SUM(quantidade * valor_unitario) as Total_Pedido FROM VW_TB_MACK_PEDIDOS_PRODUTOS WHERE cod_mack_pedido = "&sqlNumero(cod_mack_pedido))
				if not rs2.Eof then
					%>
					<h2>Total do Pedido: R$ <%=arrumaNumero(rs2(0),0,2)%></h2>
					<%
				end if
				rs2.close
				set rs2 = nothing
				%>
			</td>
		</tr>
	</table>
	</fieldset>
	<%
	else
		'response.redirect ("mck_pedidos.asp")
	end if
	rs.close 
	set rs = nothing
	%>

<!--
	<fieldset>
		<legend>Informações</legend>
		<table border="0" >
		
		<tr>
			<td>
			<ul style="text-align:left">
				<li>
					<h4>Entrega</h4>
						Quisque volutpat tortor turpis, id mollis ligula cursus at. Aliquam sagittis ultricies interdum. Ut at aliquet elit. Vivamus tempus dolor sit amet purus feugiat porta. Cras non semper eros. In porttitor ipsum elit, nec tincidunt mi sodales nec. Mauris non nunc lacinia, mollis neque a, fermentum dolor. Sed vel massa dictum, auctor velit sit amet, rutrum est. In at ante sit amet massa elementum semper.
				</li>
				<li>
					<h4>Pagamento</h4>
						Quisque volutpat tortor turpis, id mollis ligula cursus at. Aliquam sagittis ultricies interdum. Ut at aliquet elit. Vivamus tempus dolor sit amet purus feugiat porta. Cras non semper eros. In porttitor ipsum elit, nec tincidunt mi sodales nec. Mauris non nunc lacinia, mollis neque a, fermentum dolor. Sed vel massa dictum, auctor velit sit amet, rutrum est. In at ante sit amet massa elementum semper.
				</li>
			</ul>
				</td>
		</tr>
	</table>
	</fieldset>
-->
<br/><br/>
	<%if 1=2 then%>
		<table border="0">
		<tr>
			<th align="left">João Pedro Beltrame<br/><br/>____________________________</th>
		</tr>
	</table>
	<%end if%>
	<span style='font-size:10px; position:absolute; bottom:0px; right:0px;'><strong>Data e hora da impressão:</strong> <%=arrumaData(now())%></span>
</body>
</html>
<% fecharConexao() %>