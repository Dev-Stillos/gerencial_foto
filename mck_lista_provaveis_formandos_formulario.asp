<%
	Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
	BTN_VOLTAR = 1
	BTN_IMPRIMIR = 1
	BTN_FONTE = 2
	BTN_IMPRIMIR_LINK = "#"

	Dim titulo_pagina
	titulo_pagina = "MCK - Listas de Prováveis Formandos / Upload"

	Dim calendario_necessario : calendario_necessario = "NÃO"


	Function preencheCaracter(valor,caracter,qtd_caracteres)
		
		Dim retorna
		Dim diferenca
		
		retorna = valor
		
		diferenca = (qtd_caracteres-Len(valor))/1
		
			If diferenca > 0 Then
				
				For i=1 to diferenca
					retorna = caracter&retorna
				Next
				
			End If
		
		preencheCaracter = retorna
		
	End Function
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<!-- #include file="includes/freeASPUpload.asp" -->
<h2><%=titulo_pagina%></h2>
<br />
<!-- #include file="includes/sub_menu_mck.asp" -->
<br />
<%
'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 10/10/2011 16:04:06

'Instanciamos o a classe de upload
Set Upload = New FreeASPUpload
AbrirConexao()


dim msg_arr(4)
msg = 0
msg_arr(0) = "Arquivo salvo com sucesso!"
msg_arr(1) = "Formato inválido."
msg_arr(2) = "Nenhum arquivo enviado."
msg_arr(3) = ""
msg_arr(4) = ""


'Declara vari�veis para utiliza��o do sistema
Dim cod_mack_pedido, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

redirecionar_incluir	= "mck_lista_provaveis_formandos_formulario_processar.php"
redirecionar_alterar	= "mck_lista_provaveis_formandos.asp"
redirecionar_excluir	= "mck_lista_provaveis_formandos.asp"
redirecionar_validacao	= "mck_lista_provaveis_formandos.asp"
msgAoExcluir			= "Tem certeza que deseja excluir este formando?"
actionformulario		= caminho_pagina

campo1 = Upload.Form("campo1")
comando = Request.QueryString("comando")
'Declara vari�veis
Dim cod_mack_formando, cod_mack_forma_pagamento, data_cadastro, cod_usuario, nome, tia, descricao_colacao, curso, email_1, qtd_parcelas, cpf, obs

If comando = Empty Then
comando = "incluir"
titulo_pagina = titulo_pagina&" / Incluir"
End if

If comando = "incluir" and campo1 <> Empty Then
	'Incluir novo registro
	incluirRegistro(redirecionar_incluir)
	titulo_pagina = titulo_pagina&" / Inclusão"
Elseif 	comando = "download" and campo1 = Empty Then 
	'Leitura inicial para possível exclusão
	cod_mack_lista_provaveis_formandos = Request.QueryString("cod_mack_lista_provaveis_formandos")
	download()
End If


sub incluirRegistro(redirecionar)

cod_usuario = session("cod_usuario")
cod_mack_temporada = Upload.form("cod_mack_temporada")
definir_vigente = upload.form("definir_vigente")

if msg = 0 then
	'Buscamos todos os arquivos e enviamos
	if Upload.UploadedFiles.Exists("arquivo")  then
		
		'tratamos o nome do arquivo
		nome_arquivo_original = Upload.UploadedFiles.item("arquivo").fileName
		inicio_ext = instrrev(nome_arquivo_original,".")
		len_arquivo =  len(nome_arquivo_original)
		extencao = mid(nome_arquivo_original,inicio_ext+1,(len_arquivo-inicio_ext))
		nome_arquivo_original = replace(nome_arquivo_original, "."&extencao,"")


		'Tratamos o nome do arquivo
		nome_arquivo_servidor = "lista_provaveis_formandos"&preencheCaracter(year(now()),"0",4)
		nome_arquivo_servidor = nome_arquivo_servidor & preencheCaracter(month(now()),"0",2)
		nome_arquivo_servidor = nome_arquivo_servidor & preencheCaracter(day(now()),"0",2)
		nome_arquivo_servidor = nome_arquivo_servidor & preencheCaracter(hour(now()),"0",2)
		nome_arquivo_servidor = nome_arquivo_servidor & preencheCaracter(minute(now()),"0",2)
		nome_arquivo_servidor = nome_arquivo_servidor & preencheCaracter(second(now()),"0",2)

		' extencao = lcase(replace(right(Upload.UploadedFiles.item("arquivo")).fileName,4),".","")
		if  extencao = "xls" or extencao = "xlsx" then

			'SALVAMOS O ARQUIVO 
			call Upload.SaveAs("arquivo", nome_arquivo_servidor&"."&extencao, caminho_dados&"mck_listas\")

			'Antes de inserir alteramos as planilhas anteriores para NÃO vigetes
			if definir_vigente = "S" then
				objConexao.Execute("UPDATE TB_MACK_LISTAS_PROVAVEIS_FORMANDOS SET vigente = 'N' WHERE vigente = 'S'")
			end if

			'inserimos no sistema
			objConexao.Execute("INSERT INTO TB_MACK_LISTAS_PROVAVEIS_FORMANDOS (cod_mack_temporada, nome_arquivo_original, nome_arquivo_servidor, extencao, data_cadastro, status, vigente, cod_usuario  ) VALUES ("&sqlNumero(cod_mack_temporada)&", "&sqlTexto(nome_arquivo_original)&", "&sqlTexto(nome_arquivo_servidor)&", "&sqlTexto(extencao)&", NOW(), 0, "&sqlTexto(definir_vigente)&", "&sqlNumero(cod_usuario)&")")


			'Encontrar chave criada
		Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_MACK_LISTAS_PROVAVEIS_FORMANDOS LIMIT 1")
			If not rs.Eof Then
				cod_mack_lista_provaveis_formandos = rs("ULTIMO_ID")
			End if 
		rs.Close
		Set rs=Nothing



			'redirecionameos para a pagina de leitura do arquivo em PHP
			response.redirect(redirecionar& "?l="&cod_mack_lista_provaveis_formandos)

			msg = 0
		else
			msg = 2
		end if
	end if
end if'

end sub

'----------------------------------------------------------------------------------------
%>


<%If comando = "excluir" Then%>
	<div class='msgExcluir'><%=msgAoExcluir%></div>
<%End If%>

<style type="text/css">
	fieldset{
		width: 820px;
		margin:10px auto 5px auto ;
	}
</style>

<%
if msg > 0 then
	response.write msg_arr(msg)
end if
%>

<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' enctype="multipart/form-data"  onsubmit='return validaFormMaster()' method='post'>
	<fieldset>
		<legend>Infromções da Planilha</legend>
			<table width="100%" border="1" align="center" cellpadding="5" style="border-collapse:collapse" >
		 		<tr>
		            <td width="20%"><div align="left"><strong>Temporada:</strong></div></td>
		            <td width="80%">
		            	<%
		            	if cod_mack_temporada = "" and int(mackenzieTemporadaAtiva(false)) > 0 then
		            		cod_mack_temporada = int(mackenzieTemporadaAtiva(false))
		            	end if 
		            	%>
		            	<%=padraoFormularioSelect("cod_mack_temporada","cod_mack_temporada_id","inputTamanho2","","","Selecione", cod_mack_temporada,"SELECT cod_mack_temporada, concat( semestre_conclusao, '/' , ano_conclusao) as descricao FROM TB_MACK_TEMPORADAS order by cod_mack_temporada", "incluir", false, false, false)%>
		            </td>
		        </tr>
		        <tr>
		            <td width="20%"><div align="left"><strong>Planilha:</strong></div></td>
		            <td width="80%">
		            	<input type="file" name='arquivo' >		            	
		            </td>
		        </tr>
		         <tr>
		            <td width="20%"><div align="left"><strong>Definir como planilha vigente?:</strong></div></td>
		            <td width="80%">
		            	<table>
		            		<tr>
		            			<td><input type="radio" name="definir_vigente" value="S" id='definir_vigente_s' checked=""></td>
		            			<td><label for='definir_vigente_s'><strong>Sim</strong></label></td>
		            		</tr>
		            		<tr>
		            			<td><input type="radio" name="definir_vigente" value="N" id='definir_vigente_n'></td>
		            			<td><label for='definir_vigente_n'><strong>NÃO</strong></label></td>
		            		</tr>
		            	</table>

		            </td>
		        </tr>
		        <%if comando = "incluir" then %>
		        <tr>
					<td align="center" colspan="2">
				        <input class="btn btn-sm btn-warning" type="submit" name="button" id="button" value="Inserir Planilha" />   
					</td>
				</tr>
				<%end if%>

				<%if comando = "excluir" then %>
		        <tr>
					<td align="center" colspan="2">
				        <input class="btn btn-sm btn-Danger" type="submit" name="button" id="button" value="Exluir Planilha" />   
					</td>
				</tr>
				<%end if%>
			</table>
	</fieldset>	
	
<input type='hidden' value='1' name='campo1'  />

</form>

<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('COD_MACK_TEMPORADA#Temporada#2',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>
<!-- #include file="files/includes/miolo-inferior.asp" -->