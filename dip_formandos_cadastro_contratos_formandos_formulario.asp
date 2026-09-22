<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 1
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
%>
				
<%

Dim calendario_necessario
calendario_necessario = "NÃO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%

'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 9/8/2011 15:10:25

'Declara variáveis para utilização do sistema
Dim cod_contrato, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

redirecionar_incluir	= "dip_formandos_cadastro_contratos_formandos.asp"
redirecionar_alterar	= "dip_formandos_cadastro_contratos_formandos.asp"
redirecionar_excluir	= "dip_formandos_cadastro_contratos_formandos.asp"
redirecionar_validacao	= "dip_formandos_cadastro_contratos_formandos.asp"
msgAoExcluir			= "Tem certeza que deseja excluir este Contrato somente foto?<br /><br />Você poderá inativa-lo ao invés de exclui-lo. Se houver formandos com este Contrato somente foto, o mesmo não será excluído"
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
'Declara variáveis
Dim DESCRICAO, SEMESTRE, ANO, SOMENTE_FOTO

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
cod_contrato = Request.QueryString("cod_contrato")
lerRegistro()
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "alterar" and campo1 <> Empty Then
'Gravar alteração
cod_contrato = Request.Form("cod_contrato")
alterarRegistro(redirecionar_alterar)
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 
'Leitura inicial para possível exclusão
cod_contrato = Request.QueryString("cod_contrato")
lerRegistro()
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
'Exclui registro
cod_contrato = Request.Form("cod_contrato")
excluirRegistro(redirecionar_excluir)
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
'Visualiza registro
cod_contrato = Request.QueryString("cod_contrato")
lerRegistro()
titulo_pagina = titulo_pagina&" / Visualização"
End If


'Funções

Sub lerFormulario()

COD_CONTRATO = Request.Form("COD_CONTRATO")
DESCRICAO = Request.Form("DESCRICAO")
SEMESTRE = Request.Form("SEMESTRE")
ANO = Request.Form("ANO")
SOMENTE_FOTO = Request.Form("SOMENTE_FOTO")

End sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM TB_ENTIDADES WHERE cod_contrato = '"&cod_contrato&"'")

				DESCRICAO=Ucase(trim(rs("DESCRICAO")))
				SEMESTRE=Ucase(trim(rs("SEMESTRE")))
				ANO=trim(rs("ANO"))
				SOMENTE_FOTO=trim(rs("SOMENTE_FOTO"))

		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
	
		objConexao.Execute("INSERT INTO TB_ENTIDADES (COD_CONTRATO, DESCRICAO, SEMESTRE, ANO, SOMENTE_FOTO) VALUES ("&sqlTexto(COD_CONTRATO)&", "&sqlTexto(DESCRICAO)&", "&sqlTexto(SEMESTRE)&", "&sqlTexto(ANO)&",1)")

		'Encontrar chave criada
		Set rs=objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM TB_ENTIDADES LIMIT 1")
			If not rs.Eof Then
				cod_contrato = rs("ULTIMO_ID")
			End if 
		rs.Close
		Set rs=Nothing


	FecharConexao()
	Response.Redirect(redirecionar&"?cod_contrato="&cod_contrato&"&codmsg=1")
End sub

Sub alterarRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("UPDATE TB_ENTIDADES SET DESCRICAO = "&sqlTexto(DESCRICAO)&", SEMESTRE = "&sqlTexto(SEMESTRE)&", ANO = "&sqlTexto(ANO)&" WHERE cod_contrato = "&sqlTexto(cod_contrato))
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_contrato="&cod_contrato&"&codmsg=2")
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()
		call verificaAntesExcluir(cod_contrato,"TB_FICHAS_IDENTIFICACAO","cod_contrato",redirecionar_excluir&"?sucesso=2&codmsg=5")

	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_ENTIDADES WHERE cod_contrato = "&sqlTexto(cod_contrato))
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_contrato="&cod_contrato&"&codmsg=3")
End sub


'----------------------------------------------------------------------------------------

%>

<h2>DIP - Formandos Cadastro de Contratos somente Foto</h2>
<br /> 

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
<%If comando <> "incluir" Then %>
<tr>
<td width='27%'><strong>Cód. Contrato</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<span class='textoMaior'><strong><%=COD_CONTRATO%></strong></span><%Else%>
<%If not isNull(COD_CONTRATO) Then%>
<%=COD_CONTRATO%>
<%End If%>
<%End If%>
</div></td>
</tr>

<%Else%>

<tr>
<td width='27%'><strong>Cód. Contrato</strong></td>
<td width='73%'>
<div align='left'>
	Código de Contrato que será gerado: <span class='textoMaior'><strong>
    <%
	AbrirConexao()
	COD_CONTRATO = 0
	
	Set rs2=objConexao.Execute("SELECT Count(COD_CONTRATO) AS COD_CONTRATO FROM TB_ENTIDADES WHERE SOMENTE_FOTO = 1")
			
			If not rs2.Eof Then
				COD_CONTRATO = formatNumber(0&rs2(0),0)
				COD_CONTRATO = (COD_CONTRATO+1)/1
				COD_CONTRATO = Replace(COD_CONTRATO,",","")
				COD_CONTRATO = Replace(COD_CONTRATO,".","")
			Else
				COD_CONTRATO = 1
			End If
			
			COD_CONTRATO = COD_CONTRATO&"F"
		
		rs2.Close
		Set rs2=Nothing
	
	
	FecharConexao()
	
	Response.Write(COD_CONTRATO)
	
	DESCRICAO = COD_CONTRATO&" - "
	
	%>
    
    </strong></span>
</div></td>
</tr>



<%End If%><tr>
<td width='27%'><strong>Descrição *</strong></td>
<td width='73%'>
<div align='left'>
  <p>
    <%If comando <> "visualizar" and comando <> "excluir"  Then%>
      <input name='DESCRICAO' value="<%=DESCRICAO%>" class='inputTamanho3' type='text' maxlength='255' />
      <%Else%>
      <%If not isNull(DESCRICAO) Then%>
      <%=DESCRICAO%>
      <%End If%>
      <%End If%>
</p>
  Obs.: Na descri&ccedil;&atilde;o coloque o contrato e o semestre tamb&eacute;m exemplos:
  <ul>
  	<li>2706 - UNICSUL - 2012</li>
  	<li>1908 - MEDICINA ABC - JULHO/2010</li>
    <li>2717 - SAO JUDAS - 2012</li>
    <li>2738 - I T A - 2011</li>
  </ul>
  
  
 
  </div></td>
</tr>
<tr>
<td width='27%'><strong>Semestre de Conclusão *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<select name='SEMESTRE' class='inputTamanho2'>
<option value=''>Selecione</option>
<%'Padrao
Dim selRelPadrao3
selRelPadrao3 = "" 

If SEMESTRE = Empty or isNull(SEMESTRE) Then
SEMESTRE = selRelPadrao3
End if%>

<%If CStr(SEMESTRE) =  "J" Then
%><option selected='selected' value="J">Julho</option>
<%Else%>
<option value="J">Julho</option>
<%End if%>
<%If CStr(SEMESTRE) =  "D" Then
%><option selected='selected' value="D">Dezembro</option>
<%Else%>
<option value="D">Dezembro</option>
<%End if%>
</select>
<%Else%>
<%If not isNull(SEMESTRE) Then%>
<%If CStr(SEMESTRE) =  "J" Then%>
Julho<%End if%>
<%If CStr(SEMESTRE) =  "D" Then%>
Dezembro<%End if%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
<td width='27%'><strong>Ano de Conclusão *</strong></td>
<td width='73%'>
<div align='left'>
<%If comando <> "visualizar" and comando <> "excluir"  Then%>
<input name='ANO' maxlength='4' onkeypress='mascara(this,soNumeros)' class='inputTamanho1' value="<%=ANO%>" type='text' />
<%Else%>
<%If not isNull(ANO) Then%>
<%=ANO%>
<%End If%>
<%End If%>
</div></td>
</tr>
<tr>
  <td colspan='2' align='center'><a href="javascript:salvar();"><img src="images/layout/btnGravar.jpg" /></a></td>
</tr>
<%If comando <> "visualizar" Then%>
<tr>
<td colspan='2' align='center'><div align='left'><strong>* - Campos obrigatórios</strong><br /></div></td>
</tr>

<%End If%>
</table>
<input type='hidden' value='1' name='campo1'  />
<input type='hidden' name='COD_CONTRATO' value="<%=COD_CONTRATO%>" /> <input type='hidden' name='SOMENTE_FOTO' value="<%=SOMENTE_FOTO%>" /> </form>
<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
if(validaForm('DESCRICAO#Descrição#1|SEMESTRE#Semestre de Conclusão#2|ANO#Ano de Conclusão#1',document.FORMULARIO_PADRAO)){return true}else{return false}
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->
