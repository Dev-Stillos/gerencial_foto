<%
Dim calendario_necessario, BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
calendario_necessario = "NAO"
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 1
BTN_IMPRIMIR_LINK = "#"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<!-- #include file="includes/freeASPUpload.asp" -->


<%

cod_usuario = session("cod_usuario")

Set Upload = New FreeASPUpload

'Buscamos todos os arquivos e enviamos
if Upload.UploadedFiles.Exists("arquivo")  then
	
	'tratamos o nome do arquivo
	nome_arquivo_original = Upload.UploadedFiles.item("arquivo").fileName
	inicio_ext = instrrev(nome_arquivo_original,".")
	len_arquivo =  len(nome_arquivo_original)
	extencao = mid(nome_arquivo_original,inicio_ext+1,(len_arquivo-inicio_ext))
	nome_arquivo_original = replace(nome_arquivo_original, "."&extencao,"")

	'Tratamos o nome do arquivo
	nome_arquivo_servidor = "baixa_bancaria_"&preencheCaracter(year(now()),"0",4)
	nome_arquivo_servidor = nome_arquivo_servidor & preencheCaracter(month(now()),"0",2)
	nome_arquivo_servidor = nome_arquivo_servidor & preencheCaracter(day(now()),"0",2)
	nome_arquivo_servidor = nome_arquivo_servidor & preencheCaracter(hour(now()),"0",2)
	nome_arquivo_servidor = nome_arquivo_servidor & preencheCaracter(minute(now()),"0",2)
	nome_arquivo_servidor = nome_arquivo_servidor & preencheCaracter(second(now()),"0",2)

	' extencao = lcase(replace(right(Upload.UploadedFiles.item("arquivo")).fileName,4),".","")
	if  extencao = "xls" or extencao = "xlsx" then
		'SALVAMOS O ARQUIVO 
		call Upload.SaveAs("arquivo", nome_arquivo_servidor&"."&extencao, caminho_dados&"baixas\")

		'inserimos no sistema
		'objConexao.Execute("INSERT INTO TB_MACK_LISTAS_PROVAVEIS_FORMANDOS (cod_mack_temporada, nome_arquivo_original, nome_arquivo_servidor, extencao, data_cadastro, status, vigente, cod_usuario  ) VALUES ("&sqlNumero(cod_mack_temporada)&", "&sqlTexto(nome_arquivo_original)&", "&sqlTexto(nome_arquivo_servidor)&", "&sqlTexto(extencao)&", NOW(), 0, "&sqlTexto(definir_vigente)&", "&sqlNumero(cod_usuario)&")")

		'redirecionameos para a pagina de leitura do arquivo em PHP
		'response.redirect(redirecionar& "?l="&cod_mack_lista_provaveis_formandos)
	else
		msg = 2
	end if
else
	msg = 1 
end if



%>
<%=msg%>

<h2>COB - Processar arquivo</h2>
<br>
<!-- #include file="includes/classe_modulo_cob.asp" -->
<!-- #include file="includes/sub_menu_cob.asp" -->
<br />
<br />
<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' enctype="multipart/form-data" method='post'>
<table width="600" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
    <tbody><tr bgcolor="whitesmoke">
		<td>
			<h4>Instruções</h4>
			<ol start="1">
				<li>Impotação do arquivo retorno.</li>
				<li>Somente arquivos xls.</li>
			</ol>  
		</td>
	</tr>
	<tr>
		<td>
			<div align="center">
				<input type="file" class="inputTamanho3" name="arquivo">
			</div>
		</td>
	</tr>
	<tr>
		<td>
			<div align="center">
				<a href="javascript:document.FORMULARIO_PADRAO.submit();">
					<img src="images/layout/btnGravar.jpg">
				</a>
			</div>
		</td>
	</tr>
  </tbody></table>
  </form>
<br />
<br />
<br />
<!-- #include file="files/includes/miolo-inferior.asp" -->


</body>
</html>