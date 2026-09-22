<!-- #include virtual="gerencial_foto/includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="freeaspupload.asp" -->
<%

Dim mensagensPagina(5,1), msg

mensagensPagina(0,0) =  "success"
mensagensPagina(0,1) =  "Arquivo enviado com sucesso, feche a janela para continuar."

mensagensPagina(1,0) =  "info"
mensagensPagina(1,1) =  "Anexe a(s) foto(s) selecionada(s) pelo formando:"

mensagensPagina(2,0) =  "danger"
mensagensPagina(2,1) =  "Arquivo inv�lido foi encontrado, envie o arquivo em um dos formatos abaixo:<br /><ul><li>.jpg</li><li>.jpeg</li><li>.png</li></ul>"

mensagensPagina(3,0) =  "danger"
mensagensPagina(3,1) =  "Um erro inesperado ocorreu tente novamente mais tarde."

dim extencao, nome_arquivo	
Dim Upload, campo1, cod_pedido, cod_pedido_produto, qtd_arquivos, i 
Set Uploader = New FreeASPUpload

AbrirConexao()

if Upload.Form("campo1") <> "" then

	cod_pedido = Upload.form("cod_pedido")
	cod_pedido_produto = Upload.form("cod_pedido_produto")
	data_cadastro = now()
	cod_usuario = session("cod_usuario")

	if cod_pedido = "" or cod_pedido_produto = "" then
		msg = 3
	end if
	
	if isempty(msg) then
		'Buscamos todos os arquivos e enviamos
		Set rs = objConexao.Execute("SELECT id FROM TB_PEDIDOS_PRODUTOS_ARQUIVOS WHERE cod_pedido_produto = "&sqlNumero(cod_pedido_produto)&" AND nome_arquivo = '' ")
		while not rs.eof and msg = 0

			if Upload.UploadedFiles.Exists("arquivo_"&rs("id")) then
				extencao = lcase(replace(right(Upload.UploadedFiles.item("arquivo_"&rs("id")).fileName,4),".",""))

				if  extencao = "jpg" or extencao = "jpeg" or extencao = "png" then

					nome_arquivo = cod_pedido & "_" & cod_pedido_produto & "_" &rs("id")
					nome_original = Upload.UploadedFiles.item("arquivo_"&rs("id")).fileName

					'SALVAMOS O ARQUIVO 
					call Upload.SaveAs("arquivo_"&rs("id"), nome_arquivo&"."&extencao, caminho_dados&"pedidos_produtos_anexos\")

					'INCLUIMOS NA BASE
					sql = "UPDATE TB_PEDIDOS_PRODUTOS_ARQUIVOS SET nome_arquivo = "&sqlTexto(nome_arquivo)&", extencao = "&sqlTexto(extencao)&", data_cadastro = NOW(), cod_usuario = "&sqlNumero(session("COD_VENDEDOR"))&", nome_original = "&sqlTexto(nome_original)&" WHERE id = "&sqlNumero(rs("id")) 

					objConexao.Execute(sql)
					msg = 0
				else
					msg = 2
				end if
			end if

		rs.movenext
		wend
		rs.close

	end if

else
	cod_pedido = request.querystring("c")
	cod_pedido_produto = request.querystring("cp")
	msg = 1

end if

if cod_pedido = "" and cod_pedido_produto = "" then
	response.end()
end if

%>
<html>
	<head>
		<link rel="stylesheet" type="text/css" href="css/bootstrap.min.css">
		<script type="text/javascript" src='js/jquery-1.11.2.min.js'></script>
		<script type="text/javascript" src="js/bootstrap.min.js"></script>
		<script type="text/javascript"></script>
	</head>
	<body>
		<div class=''>
		<form method="post" enctype="multipart/form-data">
			<% if msg < 2 then %>
				<table border="0" cellpadding="5" width="100%" style="font-size:12px;">
					<%
					dim btn_gravar : btn_gravar = false
					i = 0 
					Set rs = objConexao.execute("SELECT * FROM VW_TB_PEDIDOS_PRODUTOS_ARQUIVOS WHERE cod_pedido_produto = "&sqlNumero(cod_pedido_produto)&" AND cod_pedido = "&sqlNumero(cod_pedido))
					if not rs.eof then

						if isempty(msg) then
							msg = 1
						end if%>
						<div class='alert alert-<%=mensagensPagina(msg,0) %>'>
							<strong>ATEN��O :</strong><br />
							<%=mensagensPagina(msg,1)%>
						</div>
						<%
						while not rs.Eof
						if i <> rs("ordem") and rs("cod_produto") = 7 then 
							i = rs("ordem")
							%>
							<tr style='font-size:16px;'>
								<td collspan="2">
									Trio (<%=i%>)
								</td>
							</tr>
						<%	
						elseif i <> rs("ordem") and rs("cod_produto") = 186 then
							i = rs("ordem")
						%>
							<tr style='font-size:16px;'>
								<td collspan="2">
									Mosaico (<%=i%>)
								</td>
							</tr>
						<%
						end if
						%>
							<tr>
								<td align="left">
									<%=rs("descricao")%>
								</td>
								<td align="center">
									<% if rs("nome_arquivo") = "" then %>
										<input type="file" name="arquivo_<%=rs("id")%>" style="width:300px;">
									<%
										btn_gravar = true
									else %>
										<a href="download.asp?t=0&p0=<%=rs("id")%>" target="_blank">Baixar</a>
									<% end if %>
								</td>
							</tr>
							<%
							'i = i + 1
						rs.movenext
						wend
					end if 

					rs.close
					set rs = nothing
					%>
					<tr>
						<td align="center" colspan="2" style="font-size:14px;">
							<%if btn_gravar then%>
								<input type="submit" vaule=" GRAVAR ">
							<%end if%>
							<input type="hidden" value="1" name="campo1">
							<input type="hidden" value="<%=comando%>" name="comando">
							<input type="hidden" value="<%=qtd_arquivos%>" name="qtd_arquivos">
							<input type="hidden" value="<%=cod_pedido%>" name="cod_pedido">
							<input type="hidden" value="<%=cod_pedido_produto%>" name="cod_pedido_produto">
						</td>
					</tr>
				</table>
			<%
			else
			%>
				<div class='alert alert-<%=mensagensPagina(msg,0) %>'>
					<strong>ATEN��O :</strong><br />
					<%=mensagensPagina(msg,1) %>
				</div>
			<%
			end if
			fecharConexao()
			%>
			</form>
		</div>
	</body>
</html>