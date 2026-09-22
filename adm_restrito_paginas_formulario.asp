<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 1
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
  Dim titulo_pagina
  titulo_pagina = "Formulário de Páginas"
%>
				
				
<%
Dim calendario_necessario
calendario_necessario = "NÃO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%

'Páginas especiais somente para o administrador master do sistema
'//////////////
If Session("cod_usuario") <> 1 Then
	Session.Abandon()
	Response.Redirect("index.asp")
End If
'//////////////

'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 27/7/2011 16:40:38

'Declara variáveis para utilização do sistema
Dim cod_pagina, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

redirecionar_incluir	= "adm_restrito_paginas.asp"
redirecionar_alterar	= "adm_restrito_paginas.asp"
redirecionar_excluir	= "adm_restrito_paginas.asp"
redirecionar_validacao	= "adm_restrito_paginas.asp"
msgAoExcluir			= "Tem certeza que deseja excluir esta Página inclusive no FTP?"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
'Declara variáveis
Dim NOME_PAGINA, COD_MODULO, CAMINHO_VALIDA, habilitada_para_todos, pag_filha_cod_grupo_acesso_pagina

If comando = Empty Then
comando = "incluir"
titulo_pagina = titulo_pagina&" / Inclusão"
End if

If comando = "incluir" and campo1 <> Empty Then
'Incluir novo registro
incluirRegistro(redirecionar_incluir)
titulo_pagina = titulo_pagina&" / Inclusão"
Elseif comando = "alterar" and campo1 = Empty Then
'Leitura inicial para possível alteração
cod_pagina = Request.QueryString("cod_pagina")
lerRegistro()
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "alterar" and campo1 <> Empty Then
'Gravar alteração
cod_pagina = Request.Form("cod_pagina")
alterarRegistro(redirecionar_alterar)
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 
'Leitura inicial para possível exclusão
cod_pagina = Request.QueryString("cod_pagina")
lerRegistro()
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
'Exclui registro
cod_pagina = Request.Form("cod_pagina")
excluirRegistro(redirecionar_excluir)
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
'Visualiza registro
cod_pagina = Request.QueryString("cod_pagina")
lerRegistro()
titulo_pagina = titulo_pagina&" / Visualização"
End If


'Funções

Sub lerFormulario()

COD_PAGINA = Request.Form("COD_PAGINA")
NOME_PAGINA = Request.Form("NOME_PAGINA")
CAMINHO_VALIDA = Request.Form("CAMINHO_VALIDA")
COD_MODULO = Request.Form("COD_MODULO")
habilitada_para_todos = Request.Form("habilitada_para_todos")
pag_filha_cod_grupo_acesso_pagina = Request.Form("pag_filha_cod_grupo_acesso_pagina")

End sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM TB_ADM_PAGINAS_HABILITADAS WHERE cod_pagina = "&cod_pagina)

				NOME_PAGINA=trim(rs("NOME_PAGINA"))
				CAMINHO_VALIDA=trim(rs("CAMINHO_VALIDA"))
				COD_MODULO=trim(rs("COD_MODULO"))
				habilitada_para_todos=trim(rs("habilitada_para_todos"))
				pag_filha_cod_grupo_acesso_pagina = trim(rs("pag_filha_cod_grupo_acesso_pagina"))

		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("INSERT INTO TB_ADM_PAGINAS_HABILITADAS (NOME_PAGINA, CAMINHO_VALIDA, COD_MODULO, habilitada_para_todos,pag_filha_cod_grupo_acesso_pagina) VALUES ("&sqlTexto(NOME_PAGINA)&", "&sqlTexto(CAMINHO_VALIDA)&", "&sqlTexto(COD_MODULO)&", "&sqlNumero(habilitada_para_todos)&",'"&pag_filha_cod_grupo_acesso_pagina&"')")

		'Encontrar chave criada
		Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_ADM_PAGINAS_HABILITADAS LIMIT 1")
			If not rs.Eof Then
				cod_pagina = rs("ULTIMO_ID")
			End if 
		rs.Close
		Set rs=Nothing


	FecharConexao()
	Response.Redirect(redirecionar&"?cod_pagina="&cod_pagina&"&codmsg=1")
End sub

Sub alterarRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("UPDATE TB_ADM_PAGINAS_HABILITADAS SET NOME_PAGINA = "&sqlTexto(NOME_PAGINA)&", CAMINHO_VALIDA = "&sqlTexto(CAMINHO_VALIDA)&", COD_MODULO = "&sqlTexto(COD_MODULO)&", habilitada_para_todos = "&sqlNumero(habilitada_para_todos)&", pag_filha_cod_grupo_acesso_pagina = '"&pag_filha_cod_grupo_acesso_pagina&"' WHERE cod_pagina = "&cod_pagina)
	FecharConexao()
	
	If int(Request.Form("alterar_codigo_fonte")) = 1 Then
		Dim codigo_fonte
		codigo_fonte = Request.Form("codigo_fonte")
		
		'Instância o objeto FSO. 
		Set Fso = CreateObject("Scripting.FileSystemObject") 
		 
		'Utilizando a propriedade CreateTextFile para criar o arquivo. 
		Set Cria = Fso.CreateTextFile("e:\home\stillos\web\gerencial_foto\"&CAMINHO_VALIDA,True)
		 
		
		arrayTexto = Split(codigo_fonte,Chr(13))
		
		'For i=0 to Ubound(arrayTexto)
			
			'Com o método WriteLine você escreve uma linha no arquivo criado acima. 
			'arrayTexto(i) = Replace(arrayTexto(i),Chr(13),"")
			'Cria.WriteLine arrayTexto(i)
			Cria.WriteLine codigo_fonte
			
			
		'Next
		
		'Fecha o arquivo criado. 
		Cria.Close 
		 
		'Destruíndo objetos 
		Set Fso = nothing 
		Set Cria = nothing 


		
		
		
	
	End If
	
	Response.Redirect(redirecionar&"?cod_pagina="&cod_pagina&"&codmsg=2")
End sub

%>

<h2><%=titulo_pagina%></h2><br />

<%If comando = "excluir" Then%>
<script type="text/javascript">
setTimeout("parent.submenu.trocarPorGravarporExclusao()",1)
</script>
<div class='msgExcluir'><%=msgAoExcluir%></div>
<%Elseif comando = "visualizar" Then%>
<script type="text/javascript">
setTimeout("parent.submenu.comandoVisualizar()",1)
</script>
<%End If%>

<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post'>
<table width='90%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
<%If comando <> "incluir" Then %><tr>
<td width='27%'><strong>Cód Página </strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<span class='textoMaior'><strong><%=COD_PAGINA%></strong></span><%Else%>
<%If not isNull(COD_PAGINA) Then%>
<%=COD_PAGINA%>
<%End If%>
<%End If%>
</div></td>
</tr>
<%End If%><tr>
<td width='27%'><strong>Descrição *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='NOME_PAGINA' value="<%=NOME_PAGINA%>" class='inputTamanho3' type='text' maxlength='250' />
<%Else%>
<%If not isNull(NOME_PAGINA) Then%>
<%=NOME_PAGINA%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>URL *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='CAMINHO_VALIDA' value="<%=CAMINHO_VALIDA%>" class='inputTamanho3' type='text' maxlength='250' />
<%Else%>
<%If not isNull(CAMINHO_VALIDA) Then%>
<input name='CAMINHO_VALIDA' value="<%=CAMINHO_VALIDA%>" class='inputTamanho3' type='hidden' maxlength='250' />
<%=CAMINHO_VALIDA%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Módulo *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<select name='COD_MODULO' class='inputTamanho3'>
<option value=''>Selecione</option>
<%AbrirConexao()
Set rsE4=objConexao.Execute("SELECT COD_MODULO, nome_modulo FROM TB_ADM_MODULOS ORDER BY descricao ASC")
While not rsE4.Eof 
'Padrao
Dim selRelPadrao4, verificador4
verificador4 = ""&rsE4(0)

selRelPadrao4 = "" 

If COD_MODULO = Empty or isNull(COD_MODULO) Then
COD_MODULO = selRelPadrao4
End if

If CStr(verificador4) = CStr(COD_MODULO)  Then%>
<option selected='selected' value="<%=rsE4(0)%>"><%=rsE4(1)%></option>
<%Else%>
<option value="<%=rsE4(0)%>"><%=rsE4(1)%></option>
<%End if%>
<%rsE4.MoveNext()
Wend
rsE4.Close
Set rsE4=Nothing
FecharConexao()%>
</select>
<%Else%>
<%If not isNull(COD_MODULO) Then%>
<%AbrirConexao()
Set rsE4=objConexao.Execute("SELECT COD_MODULO, nome_modulo FROM TB_ADM_MODULOS ORDER BY descricao ASC")
While not rsE4.Eof 
If CStr(rsE4(0)) = CStr(COD_MODULO)  Then%>
<%=rsE4(1)%><%End if%>
<%rsE4.MoveNext()
Wend
rsE4.Close
Set rsE4=Nothing
FecharConexao()%>

<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Pagina Filha de Outra P&aacute;gina *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<select name='pag_filha_cod_grupo_acesso_pagina' class='inputTamanho3'>
<option value='0'>Selecione</option>
<%AbrirConexao()
Set rsE5=objConexao.Execute("SELECT COD_PAGINA, CAMINHO_VALIDA FROM TB_ADM_PAGINAS_HABILITADAS where pag_filha_cod_grupo_acesso_pagina = 0 ORDER BY CAMINHO_VALIDA ASC")
While not rsE5.Eof 
'Padrao
Dim selRelPadrao5, verificador5
verificador5 = ""&rsE5(0)

selRelPadrao5 = "" 

If pag_filha_cod_grupo_acesso_pagina = Empty or isNull(pag_filha_cod_grupo_acesso_pagina) Then
pag_filha_cod_grupo_acesso_pagina = selRelPadrao5
End if

If CStr(verificador5) = CStr(pag_filha_cod_grupo_acesso_pagina)  Then%>
<option selected='selected' value="<%=rsE5(0)%>"><%=rsE5(1)%></option>
<%Else%>
<option value="<%=rsE5(0)%>"><%=rsE5(1)%></option>
<%End if%>
<%rsE5.MoveNext()
Wend
rsE5.Close
Set rsE5=Nothing
FecharConexao()%>
</select>
<%Else%>
<%If not isNull(pag_filha_cod_grupo_acesso_pagina) Then%>
<%AbrirConexao()
Set rsE5=objConexao.Execute("SELECT COD_PAGINA, CAMINHO_VALIDA FROM TB_ADM_PAGINAS_HABILITADAS where pag_filha_cod_grupo_acesso_pagina = 0 ORDER BY CAMINHO_VALIDA ASC")
While not rsE5.Eof 
If CStr(rsE5(0)) = CStr(pag_filha_cod_grupo_acesso_pagina)  Then%>
<%=rsE5(1)%><%End if%>
<%rsE5.MoveNext()
Wend
rsE5.Close
Set rsE5=Nothing
FecharConexao()%>

<%End If%>
<%End If%>
</div></td>
</tr>




<tr>
<td width='27%'><strong>Habilitada para todos </strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<select name='HABILITADA_PARA_TODOS' class='inputTamanho1'>
<option value=''>Selecione</option>
<%'Padrao
Dim selRelPadrao12
selRelPadrao12 = "1" 

If HABILITADA_PARA_TODOS = Empty or isNull(HABILITADA_PARA_TODOS) Then
HABILITADA_PARA_TODOS = selRelPadrao12
End if%>

<%If CStr(HABILITADA_PARA_TODOS) =  "1" Then
%><option value='1' selected='selected'>Sim</option>
<%Else%>
<option value='1'>Sim</option>
<%End if%>
<%If CStr(HABILITADA_PARA_TODOS) =  "2" Then
%><option value='2' selected='selected'>Não</option>
<%Else%>
<option value='2'>Não</option>
<%End if%>
</select>
<%Else%>
<%If not isNull(HABILITADA_PARA_TODOS) Then%>
<%If CStr(HABILITADA_PARA_TODOS) =  "1" Then
%>Sim
<%Elseif CStr(HABILITADA_PARA_TODOS) =  "2" Then%>
Não
<%End if%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Alterar C&oacute;digo Fonte:</strong></td>
<td width='73%'>
	<table width="100" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="25"><input type="checkbox" class="inputRadio" value="1" name="alterar_codigo_fonte" id="checkbox" /></td>
    <td width="75"><span style="font-weight: bold">Sim</span></td>
  </tr>
</table>

  </td>
</tr>
<%If comando <> "visualizar" Then%>
<tr>
<td colspan='2' align='center'><div align='left'><strong>* - Campos obrigatórios</strong><br /></div></td>
</tr>
<tr>
  <td colspan='2' align='center'><p><a href="javascript:salvar();"><img src="images/layout/btnGravar.jpg" /></a></p>
    <p>&nbsp;</p></td>
</tr>
<%End If%>
<tr>
<td colspan="2" align="center">
<textarea name="codigo_fonte" class="inputTamanho3" style="width:98%" cols="20" rows="20">
<%
'O parâmetro forReading é utilizado somente para leitura. 
Const forReading = 1, forWriting = 2, forAppending = 8 
 
'Instância o objeto FSO. 
Set Fso = CreateObject("Scripting.FileSystemObject") 
 
'O método GetFile retorna o arquivo existente. 
Set ObjFile = Fso.GetFile("e:\home\stillos\web\gerencial_foto\"&CAMINHO_VALIDA) 
 
'O método OpenAsTextStream abre o arquivo como sendo um fluxo de texto. 
Set objStream = ObjFile.OpenAsTextStream(forReading,false) 
 
'Looping para varrer todas as linha do arquivo. 

Dim numero_linha
numero_linha = 0

While Not ObjStream.AtEndOfStream 
 
numero_linha = numero_linha + 1 
  
'O método Readline lê uma linha de um arquivo 
Linha = Linha & ObjStream.Readline & Chr(13)
Wend 
 
'Mostra o conteúdo. 
Linha = Replace(Linha,"<","&lt;")
Linha = Replace(Linha,">","&gt;")

Response.write Linha 
 
'Fecha o arquivo. 
objStream.Close 
 
'Destruíndo objetos 
Set Fso = nothing 
Set objStream = nothing 
%>

</textarea></td>
</tr>

<tr>
<td colspan="2" align="center" bgcolor="#CCCCCC">
	
    <div style="height:250px; overflow:auto">
	<table align="center" style="font-size:12px" width="98%" border="0" cellspacing="0" cellpadding="3">
<%

 
'Instância o objeto FSO. 
Set Fso = CreateObject("Scripting.FileSystemObject") 
 
'O método GetFile retorna o arquivo existente. 
Set ObjFile = Fso.GetFile("e:\home\stillos\web\gerencial_foto\"&CAMINHO_VALIDA) 
 
'O método OpenAsTextStream abre o arquivo como sendo um fluxo de texto. 
Set objStream = ObjFile.OpenAsTextStream(forReading,false) 
 
'Looping para varrer todas as linha do arquivo. 
numero_linha = 0
Dim cor
cor = "#FFFCD9"

While Not ObjStream.AtEndOfStream 

	If cor = "#FFFCD9" Then
		cor = "#FFFFFF"
	Else
		cor = "#FFFCD9"
	End If
 
numero_linha = numero_linha + 1 
  
'O método Readline lê uma linha de um arquivo 
Linha = ObjStream.Readline
Linha = Replace(Linha,"<","&lt;")
Linha = Replace(Linha,">","&gt;")
%>

  <tr>
    <td width="36" style="border-bottom:solid 1px #EAEAEA; font-size:11px; color:#999999; background-color:#E8E8E8"><div align="center"><%=numero_linha%></div></td>
    <td width="753" bgcolor="<%=cor%>" style="border-bottom:solid 1px #EAEAEA"><%=Linha%></td>
  </tr>


<%
Wend  
 
'Fecha o arquivo. 
objStream.Close 
 
'Destruíndo objetos 
Set Fso = nothing 
Set objStream = nothing 
%>
	</table>
    </div>
</td>
</tr>



</table>
<input type='hidden' value='1' name='campo1'  />
<input type='hidden' name='COD_PAGINA' value="<%=COD_PAGINA%>" /> </form>
<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('NOME_PAGINA#Descrição#1|CAMINHO_VALIDA#URL#1|COD_MODULO#Módulo#2',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->
