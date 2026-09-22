<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/freeaspupload.asp" -->
<!-- #include virtual="gerencial_foto/files/includes/scripts.asp" -->
<!-- #include file="files/includes/functions_local.asp" -->
<%
Dim cod_mack_formando
Dim campo1, comando, comando_externo, status_produto
Dim cod_produto, obs, quantidade_solicitada
Dim pedido_manual, status_pedido_compra, cod_msg

Set Upload = New FreeASPUpload

campo1 = upload.form("campo1")

comando = upload.form("comando")
comando_externo = request.QueryString("comando_externo")
cod_mack_formando = Request.QueryString("cod_mack_formando")

redirecionar_incluir = "mck_formandos_formulario_upload.asp"
redirecionar_alterar = "mck_formandos_formulario_upload.asp"
redirecionar_excluir = "mck_formandos_formulario_upload.asp"


'Por 
if comando = "" and comando_externo <> "visualizar" and comando_externo <> "excluir" Then
	comando = "incluir"
elseif comando = "" and comando_externo = "visualizar" or comando_externo = "excluir" Then
	comando = "visualizar"
end if

'LISTAGEM DE COMANDO
If comando = "incluir" and campo1 = Empty Then
'.....
ElseIf comando = "incluir" and campo1 <> Empty Then
'Incluir novo registro
incluirRegistro(redirecionar_incluir)
titulo_pagina = titulo_pagina&" / Inclus�o"
Elseif comando = "alterar" and campo1 = Empty Then
'Leitura inicial para poss�vel altera��o
lerRegistro()
cod_mack_formando = Request.QueryString("cod_mack_formando")
titulo_pagina = titulo_pagina&" / Altera��o"
Elseif 	comando = "alterar" and campo1 <> Empty Then
'Gravar altera��o
cod_mack_formando = upload.Form("cod_mack_formando")
alterarRegistro(redirecionar_alterar)
titulo_pagina = titulo_pagina&" / Altera��o"
Elseif 	comando = "excluir" and campo1 = Empty Then 
'Leitura inicial para poss�vel exclus�o
cod_mack_formando = Request.QueryString("cod_mack_formando")
lerRegistro()
titulo_pagina = titulo_pagina&" / Exclus�o"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
'Exclui registro
cod_mack_formando = upload.Form("cod_mack_formando")
excluirRegistro(redirecionar_excluir)
titulo_pagina = titulo_pagina&" / Exclus�o"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
'Visualiza registro

titulo_pagina = titulo_pagina&" / Visualiza��o"
End If


sub lerFormulario()

	cod_mack_formando = upload.form("cod_mack_formando")

end sub


sub lerRegistro()
	


end sub


sub incluirRegistro (redirecionar)

	dim rs, nome_arquivo_original, inicio_ext, len_arquivo, extencao

	lerFormulario()

	AbrirConexao()

	if cod_mack_formando <> "" and Upload.UploadedFiles.Exists("arquivo")   Then

		'tratamos o nome do arquivo
		nome_arquivo_original = Upload.UploadedFiles.item("arquivo").fileName
		inicio_ext = instrrev(nome_arquivo_original,".")
		len_arquivo =  len(nome_arquivo_original)
		extencao = lcase(mid(nome_arquivo_original,inicio_ext+1,(len_arquivo-inicio_ext)))
		nome_arquivo_original = replace(nome_arquivo_original, "."&extencao,"")

		nome_arquivo_servidor = preencheCaracter(year(now()),"0",4)
		nome_arquivo_servidor = nome_arquivo_servidor & preencheCaracter(month(now()),"0",2)
		nome_arquivo_servidor = nome_arquivo_servidor & preencheCaracter(day(now()),"0",2)
		nome_arquivo_servidor = nome_arquivo_servidor & preencheCaracter(hour(now()),"0",2)
		nome_arquivo_servidor = nome_arquivo_servidor & preencheCaracter(minute(now()),"0",2)
		nome_arquivo_servidor = nome_arquivo_servidor & preencheCaracter(second(now()),"0",2)



		if  extencao = "jpg" or extencao = "jpeg" or extencao = "png" then

			'verificamos se j� tinha imagem relacionada
			Set rs = objConexao.Execute("SELECT nome_arquivo from TB_MACK_FORMANDOS where cod_mack_formando = "&sqlNumero(cod_mack_formando))
			if not rs.Eof Then
				'deletemos o arquivo'
				deletarArquivo(caminho_dados&"mck_identificacao\"&rs("nome_arquivo"))
			end If
			rs.close
			set rs =nothing


			'SALVAMOS O ARQUIVO 
			call Upload.SaveAs("arquivo", nome_arquivo_servidor&"."&extencao, caminho_dados&"mck_identificacao\")

			'inserimos no registo do formando
			objConexao.execute("UPDATE TB_MACK_FORMANDOS set nome_arquivo = "&sqlTexto(nome_arquivo_servidor&"."&extencao)&" where cod_mack_formando = "&sqlNumero(cod_mack_formando))

			'redirecionameos para a pagina de leitura do arquivo em PHP
			Response.redirect(redirecionar&"?cod_msg=1&comando_externo="&comando_externo&"&cod_mack_formando="&cod_mack_formando)

		else
		
			Response.redirect(redirecionar&"?cod_msg=2&comando_externo="&comando_externo&"&cod_mack_formando="&cod_mack_formando)
		end if

	end if

	FecharConexao()

	Response.redirect(redirecionar&"?cod_msg=3&comando_externo="&comando_externo&"&cod_mack_formando="&cod_mack_formando)
	
end sub
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link rel="stylesheet" type="text/css" href="css/bootstrap.min.css" />
<link type="text/css" rel="stylesheet" href="http://stillos.com.br/gerencial_foto/css/base.css" />
<link rel="stylesheet" media="all" type="text/css" href="http://stillos.com.br/gerencial_foto/css/modal-message.css">
<style>
	fieldset{
		width: 820px;
		margin:0 auto 0 auto;
	}
</style>
</head>
<body style="padding:10px; background-color:#FFF">
<%

	if cod_msg = "" then
		cod_msg = request.QueryString("cod_msg")
	end if

	if cod_msg = 1 then
	%>
		<div class="msgGrandeSucesso">
			Arquivo incluído com sucesso!			
		</div>
	<%
	elseif cod_msg = 2 then
	%>
		<div class="msgGrandeErro">
			Arquivo inválido.
		</div>
		<%
	elseif cod_msg = 3 then
	%>
		<div class="msgGrandeErro">
			Selecione um arquivo.
		</div>
	<%
	end if
%>
<fieldset>
		<legend>
			Foto Identificação
		</legend>
		<form name='FORMULARIO_PADRAO' action="?comando=<%=comando%>&comando_externo=<%=comando_externo%>" method="post" enctype="multipart/form-data">
		<table  width='800px' border='0' align='center' cellpadding='5' style='border-collapse:collapse' >
			<tr><td colspan="2">
				<div class="alert alert-info">
					<strong>
						Atenção somente serão aceitos os formatos de imagem listados abaixo:
					</strong>
					<br>
					<ul>
						<li>.jpg</li>
						<li>.jpeg</li>
						<li>.png</li>
					</ul>
				</div>
			</td></tr>
			<tr>
				<td><input type="file" name="arquivo" /></td>
				<td  align="right">
					<input type="submit" class='btn btn-sm btn-success' value="Salvar Imagem" />	
				</td>
			</tr>
			<% if cod_mack_formando <> Empty Then 
			AbrirConexao()

			set rs = objConexao.execute("SELECT * FROM TB_MACK_FORMANDOS where cod_mack_formando = "&sqlNumero(cod_mack_formando) )
			if not rs.Eof then

				if rs("nome_arquivo")&"" <> "" then
				%>
					<tr>
						<td colspan="2">
							<img src="mck_formandos_formulario_download.asp?c=<%=cod_mack_formando%>" width="800px">
						</td>
					</tr>
				<%
				end if
			end if 
			rs.close
			set rs =nothing

			FecharConexao()

			end if
			%>
		</table>
		<input type="hidden" name='cod_mack_formando' value="<%=cod_mack_formando%>" />
		<input type="hidden" name='comando' value="<%=comando%>" />
		<input type="hidden" name='campo1' value='1' />
		</form>
	</fieldset>
<!-- #include virtual="gerencial_foto/files/includes/scripts_rodape.asp" -->
</body>
</html>