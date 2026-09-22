<!-- #include virtual="gerencial_foto/includes/xqrtopqqetyynmkgfdconexao.asp" -->
<%
Dim cod_transferencia : cod_transferencia = request.querystring("cod_transferencia")

Dim cod_movimentacao_origem, cod_movimentacao_destino, local_armazenamento_origem, local_armazenamento_destino

Dim itens

abrirConexao

'BUSCAMOS OS CODIGOS DE MOVIMENTAÇÕES RELACIONADOS A ESTA TRANSFERENCIA
set rs = objConexao.execute("SELECT cod_movimentacao_origem, cod_movimentacao_destino, local_armazenamento_destino, local_armazenamento_origem FROM vw_est_movimentacao_transferencias WHERE cod_transferencia = "&sqlNumero(cod_transferencia))
if not rs.Eof then
	cod_movimentacao_destino = rs("cod_movimentacao_destino")
	local_armazenamento_origem = rs("local_armazenamento_origem")
	local_armazenamento_destino = rs("local_armazenamento_destino")
end if
rs.close
set rs = nothing

if local_armazenamento_origem = "Estoque Padrão" then
	local_armazenamento_origem = "Responsável"
end if

'BUSCAMOS OS ITEMS MOVIMENTADOS 
set rs = objConexao.execute("SELECT tb_est_itens_movimentacao.* , `TB_PRODUTOS`.`NOME_PRODUTO` FROM tb_est_itens_movimentacao JOIN `stillos3`.`TB_PRODUTOS` ON `TB_PRODUTOS`.`COD_PRODUTO` = `tb_est_itens_movimentacao`.`cod_item` WHERE tb_est_itens_movimentacao.cod_movimentacao = "&sqlNumero(cod_movimentacao_destino))
dim i : i = 0
while not rs.Eof 
	i = i + 1
	itens = itens & "<tr>"&_
						"<td align='center'>"&i&"</td>"&vbnewline&_
						"<td align='center'>"&rs("Quantidade")&"</td>"&vbnewline&_
						"<td align='left'>"&rs("NOME_PRODUTO")&"</td>"&vbnewline&_
					"</tr>"

rs.movenext()
Wend
rs.close
set rs = nothing


fecharConexao

%>
<html>
	<head>
		<title>Protocolo de transferência: </title>
		<!-- #include virtual="gerencial_foto/files/includes/scripts.asp" -->
		<!-- #include virtual="gerencial_foto/files/includes/functions_local.asp" -->
		<style type="text/css">
			body{
			    print-color-adjust:exact;
			    -webkit-print-color-adjust:exact;
			    -moz-print-color-adjust:exact;
			    -o-print-color-adjust:exact;
			    font-family: arial;
			}
			.via-2{margin-top:10px;}
			#container{	width: 950px;margin: 0px auto 0px auto;}
			.f_left{position:relative; float: left;}
			.outer-border{border: solid #000 5px; padding: 4px;}
			.inner-border{border: solid #CD0462 3px;}
			.header, .footer{ width: 922px;}
			.header .logo{margin: 7px; height: 60px; width: 140px; background-image: url(images/layout/logo_reportagens.png); background-position: center; background-repeat: no-repeat;}
			.header .title, .header .data-emissao{height: 62px; margin: 7px; font-size: 22px; text-decoration: underline;	text-align: center;	line-height: 60px; font-weight: bold;}
			.header .title{width: 595px;}
			.header .data-emissao{width: 120px;}
			.texto-padrao{margin: 7px; width: 916px;}
			.itens{height: 250px; width: 916px;	margin: 7px; border: 1px solid #000;}
			.itens > table{font-size: 12px;}
			.footer .ass, .footer .data-ass{ margin: 7px; margin-top: 30px; line-height: 20px; font-weight: bold;}
			.footer .ass {width: 447px;}
			.footer .data-ass{width: 160px; margin-top: 50px;}
			.footer .id {height: 15px; font-size: 10px; font-weight: bold; text-align: right; width: 916px;margin: 7px 7px 0px 7px;}

			.cut{width: 100%; height: 20px;}
			.cut .texto{
				width: 60px;
				height: 20px;
				font-size: 12px;
				margin: 8px 0px 8px 0px;
			}
			.cut .row{
				width: 890px;
				height: 16px;
				border-bottom: 1px dashed #000;

			}
		</style>
		<script type="text/javascript">
			$(function(){window.print()});
		</script>
	</head>
	<body>
		<div id="container">
			<div class='outer-border f_left via-1'>
				<div class='inner-border f_left'>
					<div class='header f_left'>
						<div class='logo f_left'></div>
						<div class='title f_left'>COMPROVANTE DE TRANSFERÊNCIA</div>
						<div class='data-emissao f_left'><%=cod_transferencia%>/<%=year(now())%></div>
					</div>
					<div class='texto-padrao f_left'>
						&nbsp;&nbsp;&nbsp;&nbsp;Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nam ac diam at nisl vestibulum consequat. Praesent nec augue purus. Morbi ornare, diam convallis ultricies cursus, lacus nunc viverra odio, et consequat sem mi et lectus. Vivamus vel purus ut odio vulputate sagittis. Nullam nec faucibus elit
					</div>
					<div class='itens f_left'>
						<table border="0" width="100%" cellpadding="0" cellspacing="0">
							<tr bgcolor="#F5F5F5">
								<th width="5%" align="center">#</th>
								<th width="20%" align="center">Quantidade</th>
								<th width="75%" align="left">Itens</th>
							</tr>
							<%=itens%>
						</table>
					</div>
					<div class='footer f_left'>
						<div class="ass f_left">
							<%=local_armazenamento_origem%>:<br><br><br>
							_____________________________________________
						</div>
						<div class="ass f_left">
							<%=local_armazenamento_destino%>:<br><br><br>
							_____________________________________________
						</div>
						<div class="data-ass f_left">
							Data:___/___/______
						</div>

						<div class='id f_left'>
							<!--Gerado por ******* as 99/99/9999 às 15h45m - IP 999.999.999.999-->
						</div>
					</div>
				</div>
			</div>

			<div class='cut f_left'>
				<div class='texto f_left'>Corte aqui</div>
				<div class='f_left row'></div>
			</div>
			<div class='outer-border f_left via-2'>
				<div class='inner-border f_left'>
					<div class='header f_left'>
						<div class='logo f_left'></div>
						<div class='title f_left'>COMPROVANTE DE TRANSFERÊNCIA</div>
						<div class='data-emissao f_left'><%=cod_transferencia%>/<%=year(now())%></div>
					</div>
					<div class='texto-padrao f_left'>
						&nbsp;&nbsp;&nbsp;&nbsp; Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nam ac diam at nisl vestibulum consequat. Praesent nec augue purus. Morbi ornare, diam convallis ultricies cursus, lacus nunc viverra odio, et consequat sem mi et lectus. Vivamus vel purus ut odio vulputate sagittis. Nullam nec faucibus elit
					</div>
					<div class='itens f_left'>
						<table border="0" width="100%" cellpadding="0" cellspacing="0">
							<tr bgcolor="#F5F5F5">
								<th width="5%" align="center">#</th>
								<th width="20%" align="center">Quantidade</th>
								<th width="75%" align="left">Itens</th>
							</tr>
							<%=itens%>
						</table>
					</div>
					<div class='footer f_left'>
						<div class="ass f_left">
							<%=local_armazenamento_origem%>:<br><br><br>
							_____________________________________________
						</div>
						<div class="ass f_left">
							<%=local_armazenamento_destino%>:<br><br><br>
							_____________________________________________
						</div>
						<div class="data-ass f_left">
							Data:___/___/______
						</div>

						<div class='id f_left'>
							<!--Gerado por ******* as 99/99/9999 às 15h45m - IP 999.999.999.999-->
						</div>
					</div>
				</div>
			</div>

		</div>
	</body>
</html>