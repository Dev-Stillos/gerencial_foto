<%
Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 2
BTN_IMPRIMIR_LINK = "#"
titulo_pagina = "PPV - Produtos - Etapas de Ordens de Serviços - Formulário"

Dim calendario_necessario
calendario_necessario = "NÃO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%
'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 5/5/2011 21:35:28

'Declara vari�veis para utiliza��o do sistema
Dim cod_etapa, nome_produto, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario

redirecionar_incluir	= "ppv_produtos_etapas_formulario.asp"
redirecionar_alterar	= "ppv_produtos_etapas_formulario.asp"
redirecionar_excluir	= "ppv_produtos_etapas_formulario.asp"
redirecionar_validacao	= "ppv_produtos_etapas_formulario.asp"
msgAoExcluir			= "Tem certeza que deseja excluir esta Etapa desta Categoria?<br />Se houver OS com esta etapa as mesmas serão mantidas, pois já foram geradas caso queira excluir neste caso exclua a Ordem de Serviço."
actionformulario		= caminho_pagina

campo1 = Request.Form("campo1")
comando = Request.QueryString("comando")
'Declara vari�veis
Dim cod_produto, numero_etapa, cod_usuario_responsavel, descricao, instrucoes, prazo_em_dias, necessita_gerar_pedido_compra

cod_produto = Request.QueryString("cod_produto")

If comando = Empty Then
comando = "incluir"
titulo_pagina = titulo_pagina&" / Inclusão"
End if

If comando = "incluir" and campo1 <> Empty Then
'Incluir novo registro
incluirRegistro(redirecionar_incluir)
titulo_pagina = titulo_pagina&" / inclusão"
Elseif comando = "alterar" and campo1 = Empty Then
'Leitura inicial para possível alteração
cod_etapa = Request.QueryString("cod_etapa")
lerRegistro()
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "alterar" and campo1 <> Empty Then
'Gravar altera��o
cod_etapa = Request.Form("cod_etapa")
alterarRegistro(redirecionar_alterar)
titulo_pagina = titulo_pagina&" / Alteração"
Elseif 	comando = "excluir" and campo1 = Empty Then 
'Leitura inicial para possível exclusão
cod_etapa = Request.QueryString("cod_etapa")
lerRegistro()
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "excluir" and campo1 <> Empty Then 
'Exclui registro
cod_etapa = Request.Form("cod_etapa")
excluirRegistro(redirecionar_excluir)
titulo_pagina = titulo_pagina&" / Exclusão"
Elseif 	comando = "visualizar" and campo1 = Empty Then 
'Visualiza registro
cod_etapa = Request.QueryString("cod_etapa")
lerRegistro()
titulo_pagina = titulo_pagina&" / Visualização"
End If


'Fun��es Extras

AbrirConexao()

	Set rs=objConexao.Execute("SELECT NOME_PRODUTO FROM TB_PRODUTOS WHERE COD_PRODUTO = "&COD_PRODUTO)
			if not rs.Eof then
				nome_produto = rs(0)
			end if
	rs.close
	Set rs=nothing


If comando = "incluir" and campo1 = Empty Then
	
	numero_etapa = 0
	
	'Verifica �ltima etapa v�lida
	Set rs=objConexao.Execute("SELECT COUNT(*) FROM TB_PRODUTOS_ETAPAS WHERE COD_PRODUTO = "&COD_PRODUTO)
			if not rs.Eof then
				numero_etapa = formatNumber(0&rs(0),0)
			end if
	rs.close
	Set rs=nothing
	
	numero_etapa = (numero_etapa+1)/1
	
End If

FecharConexao()


Sub lerFormulario()

cod_etapa = Request.Form("cod_etapa")
cod_produto = Request.Form("cod_produto")
numero_etapa = Request.Form("numero_etapa")
necessita_gerar_pedido_compra = Request.Form("necessita_gerar_pedido_compra")
cod_usuario_responsavel = Cstr(""&Request.Form("cod_usuario_responsavel"))
descricao = Request.Form("descricao")
instrucoes = Request.Form("instrucoes")
prazo_em_dias = Request.Form("prazo_em_dias")

End sub

Sub lerRegistro()
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM TB_PRODUTOS_ETAPAS WHERE cod_etapa = "&cod_etapa)

				numero_etapa=trim(rs("numero_etapa"))				
				descricao=trim(rs("descricao"))
				instrucoes=trim(rs("instrucoes"))
				necessita_gerar_pedido_compra=trim(rs("necessita_gerar_pedido_compra"))
				prazo_em_dias=trim(rs("prazo_em_dias"))
				cod_usuario_responsavel=trim(rs("cod_usuario_responsavel"))

		rs.Close
		Set rs=Nothing
	FecharConexao()
End sub

Sub incluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		sql = "INSERT INTO TB_PRODUTOS_ETAPAS (cod_produto, numero_etapa, cod_usuario_responsavel, descricao, instrucoes, prazo_em_dias, necessita_gerar_pedido_compra) VALUES ("&sqlNumero(cod_produto)&", "&sqlNumero(numero_etapa)&", "&sqlTexto(cod_usuario_responsavel)&", "&sqlTexto(descricao)&", "&sqlTexto(instrucoes)&", "&sqlTexto(prazo_em_dias)&", "&sqlTexto(necessita_gerar_pedido_compra)&")"
		Response.Write(sql)
		objConexao.Execute(sql)


	FecharConexao()
	Response.Redirect(redirecionar&"?cod_produto="&cod_produto&"&cod_etapa="&cod_etapa&"&codmsg=1")
End sub

Sub alterarRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		sql = "UPDATE TB_PRODUTOS_ETAPAS SET cod_usuario_responsavel = "&sqlTexto(cod_usuario_responsavel)&", descricao = "&sqlTexto(descricao)&", instrucoes = "&sqlTexto(instrucoes)&", prazo_em_dias = "&sqlTexto(prazo_em_dias)&", necessita_gerar_pedido_compra = "&sqlTexto(necessita_gerar_pedido_compra)&" WHERE cod_etapa = "&cod_etapa
		Response.Write(sql)
		
		objConexao.Execute(sql)
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_produto="&cod_produto&"&cod_etapa="&cod_etapa&"&codmsg=2")
End sub

Sub excluirRegistro(redirecionar)
	lerFormulario()

	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_PRODUTOS_ETAPAS WHERE cod_etapa = "&cod_etapa)
		
		'Reorganiza n�mero de etapas
		Dim i
		i=1
		
		Set rs=objConexao.Execute("SELECT cod_etapa FROM TB_PRODUTOS_ETAPAS WHERE cod_produto = "&cod_produto&" order by numero_etapa asc")
			if not rs.Eof then
				While not rs.Eof
					objConexao.Execute("UPDATE TB_PRODUTOS_ETAPAS SET numero_etapa = "&i&" WHERE cod_etapa = "&rs("cod_etapa"))
					i=i+1
				rs.MoveNext()
				Wend
			end if
		rs.close
		Set rs=nothing	
		
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_produto="&cod_produto&"&cod_etapa="&cod_etapa&"&codmsg=3")
End sub


'----------------------------------------------------------------------------------------

%>
<h2><%=titulo_pagina%></h2>
<br />

<!-- #include file="includes/classe_modulo_ppv_menu.asp" -->
<br />
<br />

<%
	Dim sucesso, codmsg
	codmsg	= CStr(""&Request.QueryString("codmsg"))
	sucesso = CStr(""&Request.QueryString("sucesso"))
		
	If sucesso = "" and codmsg <> "" Then%>
    	<div align="center" class="msgGrandeSucesso">
        	<%
			If codmsg = "1" Then
				Response.Write("Etapa incluída com sucesso!")
			Elseif codmsg = "2" Then
				Response.Write("Etapa alterada com sucesso!")
			Elseif codmsg = "3" Then	
				Response.Write("Etapa excluída com sucesso!")
			End If	
			%>
            
        
        </div>
    <%ElseIf sucesso <> "" and codmsg <> "" Then%>
        <div align="center" class="msgGrandeErro">
			<%
			If codmsg = "4" Then
				Response.Write("Contato NÃO alterado para NÃO haver duplicidade de informação!<BR>O conteúdo do campo """&Request.QueryString("chaveUnica")&""" já possui uma informação igual em outro Contato.<br />")
			Elseif codmsg = "5" Then
				Response.Write("Contato NÃO excluído, pois há Funções que dependem deste Contato.<BR>Se NÃO desejar ter mais esta Contato inative ao invés de excluir.<br />")
			
			'////////////////////////
			'Padr�o para Formulários
			Elseif codmsg = "900" Then
				Response.Write("Você NÃO possui acesso a VISUALIZAÇÃO!")
			Elseif codmsg = "901" Then
				Response.Write("Você NÃO possui acesso a inclusão!")
			Elseif codmsg = "902" Then
				Response.Write("Você NÃO possui acesso a ALTERAÇÃO!")	
			Elseif codmsg = "903" Then
				Response.Write("Você NÃO possui acesso a EXCLUSÃO!")
			End If	
			%>
</div>
	<%End If%>
<form action='<%=actionformulario%>?comando=<%=comando%>&COD_PRODUTO=<%=COD_PRODUTO%>' name='FORMULARIO_PADRAO' onsubmit='return validaFormMaster()' method='post'>
<%
AbrirConexao()
	'Contatos
	Set rs=objConexao.Execute("Select * From TB_PRODUTOS_ETAPAS Where COD_PRODUTO = '"&COD_PRODUTO&"' order by numero_etapa asc")%>
<table width="96%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
	<tr bgcolor="whitesmoke">
            <td colspan="7" class="textoMedio"><strong><%=nome_produto%>&nbsp;</strong></td>
      </tr>
  <tr bgcolor="whitesmoke">
            <td width="8%"><div align="center" style="font-weight: bold"><strong>Número da Etapa *</strong></div></td>
      <td width="22%"><div align="center" style="font-weight: bold"><strong>Título da Etapa *</strong></div></td>
      <td width="27%"><div align="center" style="font-weight: bold"><strong>Instruções </strong></div></td>
            <td width="21%"><div align="center"><strong>responsável</strong></div></td>
      <td width="6%"><div align="center" style="font-weight: bold">Prazo em dias corridos para finaliza&ccedil;&atilde;o</div></td>
      <td width="6%"><div align="center"><strong>Necessita gerar Pedido de Compra</strong></div></td>
      <td width="10%">&nbsp;</td>
    </tr>
            <tr>
              <td>
                <div align="center">
                  <input name='NUMERO_ETAPA' class='inputLarguraSmalInput' value="<%=NUMERO_ETAPA%>" type='hidden' />
                      <strong><%=NUMERO_ETAPA%></strong>
                  </div></td><td><div align='left'>
                <%If comando <> "visualizar" and comando <> "excluir"  Then%>
                <input name='DESCRICAO' style="width:96%" value="<%=DESCRICAO%>" class='inputTamanho3' type='text' maxlength='255' />
                (Utilize a palavra: "Triagem" para utilizar a opção de triagem válida para a 1° etapa.)
                
                <%Else%>
                <%If not isNull(DESCRICAO) Then%>
                <%=DESCRICAO%>
                <%End If%>
                <%End If%>
              </div></td>
              <td><div align='left'>
                <%If comando <> "visualizar" and comando <> "excluir"  Then%>
                <table width="232" border='0' align="center" cellpadding='0' cellspacing='0'>
                  <tr>
                    <td width='57'><textarea rows='5' name='INSTRUCOES' class='inputTamanho2' onkeydown='textCounter(this.form.INSTRUCOES,this.form.charRestantes_INSTRUCOES,255);' onkeyup='textCounter(this.form.INSTRUCOES,this.form.charRestantes_INSTRUCOES,255);'><%=INSTRUCOES%></textarea></td>
                  </tr>
                  <tr>
                    <td><table width='50' border='0' align='right' cellpadding='0' cellspacing='0'>
                        <tr>
                          <td><div align='center'>
                              <input class='inputContaCaracteres' tabindex='10000' type='text' name='charRestantes_INSTRUCOES' value="<%=255- Len(INSTRUCOES)%>" />
                          </div></td>
                        </tr>
                    </table></td>
                  </tr>
                </table>
                <%Else%>
                <%If not isNull(INSTRUCOES) Then%>
                <%=Replace(INSTRUCOES,Chr(13),"<br />")%>
                <%End If%>
                <%End If%>
              </div></td>
              <td><table width="180px" border="0" align="center" cellpadding="3" cellspacing="0">
                <tr>
                  <td><div align='left'>
                    <%If comando <> "visualizar" and comando <> "excluir" and numero_etapa = 1 Then%>
                    <select name='COD_USUARIO_RESPONSAVEL' style="width:180px;" class='inputTamanho3'>
                      <option value=''>Selecione</option>
                      <%
Set rsE5=objConexao.Execute("SELECT cod_usuario ,nome FROM TB_ADM_USUARIOS where bloqueado = '2' ORDER BY nome ASC")
While not rsE5.Eof 
'Padrao
Dim selRelPadrao5, verificador5
verificador5 = ""&rsE5(0)

selRelPadrao5 = "" 

If COD_USUARIO_RESPONSAVEL = Empty or isNull(COD_USUARIO_RESPONSAVEL) Then
COD_USUARIO_RESPONSAVEL = selRelPadrao5
End if

If CStr(verificador5) = CStr(COD_USUARIO_RESPONSAVEL)  Then%>
                      <option selected='selected' value="<%=rsE5(0)%>"><%=rsE5(1)%></option>
                      <%Else%>
                      <option value="<%=rsE5(0)%>"><%=rsE5(1)%></option>
                      <%End if%>
                      <%rsE5.MoveNext()
Wend
rsE5.Close
Set rsE5=Nothing
%>
                      </select>
                    <%Else%>
                    <%If not isNull(COD_USUARIO_RESPONSAVEL) Then%>
                    <%
Set rsE5=objConexao.Execute("SELECT cod_usuario, nome FROM TB_ADM_USUARIOS where bloqueado = '2' ORDER BY nome ASC")
While not rsE5.Eof 
If CStr(rsE5(0)) = CStr(COD_USUARIO_RESPONSAVEL)  Then%>
                    <%=rsE5(1)%>
                    <%End if%>
                    <%rsE5.MoveNext()
Wend
rsE5.Close
Set rsE5=Nothing
%>
                    <%End If%>
                    <%End If%>
                    </div></td>
                </tr>
              </table></td>
              <td><div align='left'>
                <%If comando <> "visualizar" and comando <> "excluir"  Then%>
                <input name='PRAZO_EM_DIAS' maxlength='10' onkeypress='mascara(this,soNumeros)' class='inputTamanho1' value="<%=PRAZO_EM_DIAS%>" type='text' />
                <%Else%>
                <%If not isNull(PRAZO_EM_DIAS) Then%>
                <%=PRAZO_EM_DIAS%>
                <%End If%>
                <%End If%>
              </div></td>
              <td><div align="center">
              <select name='necessita_gerar_pedido_compra' class='inputTamanho1'>
<%'Padrao
Dim selRelPadrao14
selRelPadrao14 = "N" 

If necessita_gerar_pedido_compra = Empty or isNull(necessita_gerar_pedido_compra) Then
necessita_gerar_pedido_compra = selRelPadrao14
End if%>

<%If CStr(necessita_gerar_pedido_compra) =  "S" Then
%><option selected='selected' value="S">Sim</option>
<%Else%>
<option value="S">Sim</option>
<%End if%>
<%If CStr(necessita_gerar_pedido_compra) =  "N" Then
%><option selected='selected' value="N">NÃO</option>
<%Else%>
<option value="N">NÃO</option>
<%End if%>
</select>
              </div></td>
              <td><div align="center">
                <a href="javascript:salvar();">
                  <%
				If comando = "excluir" Then
				%>
                  <img src="images/layout/btnExcluir.jpg" />
                  <%Else%>
                  <img src="images/layout/btnGravar.jpg" />
                  <%End If%>
                </a></div></td>
			</tr>
    
    <%
			While not rs.Eof%>
				
			<tr>
              <td><div align="left"><%=rs("numero_etapa")%></div></td>
              <td><div align="left"><%=rs("descricao")%></div></td>
              <td><div align="left"><%=Replace(rs("instrucoes"),Chr(13),"<br />")%></div></td>
              <td><div align="center"><%=nomeUsuario(rs("cod_usuario_responsavel"),true)%></div></td>
              <td><div align="center"><%=rs("PRAZO_EM_DIAS")%></div></td>
              <td><div align="center"><%=booleano(rs("necessita_gerar_pedido_compra"))%></div></td>
              <td><div align="center"><table border="0" align="center" cellpadding="3" cellspacing="0">
                <tr>
                  <td> 
                    <a href="javascript:void(0);" onclick="abre('ppv_produtos_etapas_formulario.asp?comando=alterar&cod_etapa=<%=rs(0)%>&cod_produto=<%=cod_produto%>')">
                      <img src="images/icons/icon_alterar.png" width="16" height="16" border="0" />  	</a>  </td>
                  <td> <a href="javascript:void(0);" onclick="abre('ppv_produtos_etapas_formulario.asp?comando=excluir&cod_etapa=<%=rs(0)%>&cod_produto=<%=cod_produto%>')">
                    <img src="images/icons/icon_excluir.png" width="16" height="16" border="0" />
                    </a></td>
                  </tr>
</table></div></td>
			</tr>	
				
			<%
			rs.MoveNext()
			Wend
			%>
            </table>
<br />
<br />
Obs.: Usuário responsável somente na 1° etapa.
        <%	
	
	rs.Close
	Set rs=Nothing
	
FecharConexao()
%>
<input type='hidden' value='1' name='campo1'  />
<input type='hidden' name='COD_ETAPA' value="<%=COD_ETAPA%>" /> <input type='hidden' name='COD_PRODUTO' value="<%=COD_PRODUTO%>" /></form>
<script type='text/javascript'>

<%If comando <> "visualizar" and comando <> "excluir"  Then%>
function validaFormMaster() {
	
	<%if numero_etapa = 10000 Then
	'Desabilitado validação para Usuário na 1° etapa. 
	%>
		if(validaForm('COD_USUARIO_RESPONSAVEL#Usuário responsável pela 1° etapa#1|DESCRICAO#Título da Etapa#1|PRAZO_EM_DIAS#Prazo estimado em dias corridos#1|necessita_gerar_pedido_compra#Necessita Gerar Pedido de Compra para Fornecedor',document.FORMULARIO_PADRAO)){return true}else{return false}
	<%Else%>
		if(validaForm('DESCRICAO#Título da Etapa#1|PRAZO_EM_DIAS#Prazo estimado em dias corridos#1|necessita_gerar_pedido_compra#Necessita Gerar Pedido de Compra para Fornecedor',document.FORMULARIO_PADRAO)){return true}else{return false}
	<%End If%>
}
<%Else%>
function validaFormMaster() {
return true
}
<%End If%>
</script>
<script type="text/javascript">
function salvar(){
	if(validaFormMaster()){
		
		if(document.FORMULARIO_PADRAO.RESPONSAVEL_COD_DEPARTAMENTO){
			if(document.FORMULARIO_PADRAO.RESPONSAVEL_COD_DEPARTAMENTO.value == '' && document.FORMULARIO_PADRAO.COD_USUARIO_RESPONSAVEL.value == ''){
			alert('Atenção selecione o responsável pela etapa Departamento ou Funcionário específico.');
			}else{
			document.FORMULARIO_PADRAO.submit();
			}
		}else{
			document.FORMULARIO_PADRAO.submit();
		}
	}
}
</script>
<!-- #include file="files/includes/miolo-inferior.asp" -->
