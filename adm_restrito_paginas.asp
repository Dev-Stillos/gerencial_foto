<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 1
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
  Dim titulo_pagina
  titulo_pagina = "Formulário de Páginas"

'Páginas especiais somente para o administrador master do sistema
'//////////////
If Session("cod_usuario") <> 1 Then
	Session.Abandon()
	Response.Redirect("index.asp")
End If
'//////////////
If Request.QueryString("comando") = "excluir" Then
	cod_pagina = Request.QueryString("cod_pagina")
	CAMINHO_VALIDA = Request.QueryString("CAMINHO_VALIDA")

	AbrirConexao()
		'Excluir na base de dados
		objConexao.Execute("DELETE FROM TB_ADM_PAGINAS_HABILITADAS WHERE cod_pagina = "&cod_pagina)
		objConexao.Execute("DELETE FROM TB_ADM_PERFIS_PAGINAS_HABILITADAS WHERE cod_pagina = "&cod_pagina)
		
		'Excluir no FTP
		dim caminho_completo
		caminho_completo = "e:\home\stillos\web\gerencial_foto\"&CAMINHO_VALIDA
		deletarArquivo(caminho_completo)
		
	FecharConexao()
	Response.Redirect(redirecionar&"?cod_pagina="&cod_pagina&"&codmsg=3")
End If
%>

				
				
<%
Dim calendario_necessario
calendario_necessario = "NÃO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<script type="text/javascript">
function confirmaExclusao(cod_pagina,caminho_valida){
	var msg = "Tem certeza que deseja excluir a página "+caminho_valida+"?";
	var caminho = "adm_restrito_paginas.asp?comando=excluir&cod_pagina="+cod_pagina+"&caminho_valida="+caminho_valida;
	if(YNconfirm(msg)){
		abre(caminho);	
	}
}

</script>
<h2>Páginas habilitadas do Sistema</h2>
<br />
<div class="subMenu">Sub-menu: <a href="adm_restrito_paginas.asp">Listar Páginas</a> | <a href="adm_modulos_paginas_habilitadas_cadastrar.asp">Cadastrar Páginas</a></div>
<br />
<%
AbrirConexao()
	'FILTRO POR SISTEMA!!!	
	sql = "SELECT TB_ADM_PAGINAS_HABILITADAS.*, TB_ADM_MODULOS.* from TB_ADM_MODULOS inner join TB_ADM_PAGINAS_HABILITADAS on TB_ADM_MODULOS.cod_modulo = TB_ADM_PAGINAS_HABILITADAS.cod_modulo  where pag_filha_cod_grupo_acesso_pagina = 0 order by TB_ADM_MODULOS.order_MENU,TB_ADM_PAGINAS_HABILITADAS.caminho_valida"
	
	'Response.Write(sql)
	
	Set rs=server.CreateObject("ADODB.recordset")
	rs.CursorLocation=3  'Cursor no cliente
	rs.PageSize=200
	rs.Open sql, objConexao
	
	Dim paginacaocao
	paginacaocao = formatNumber(0&Request.QueryString("paginacao"),0)
	
	If not rs.Eof Then
	
		If paginacaocao > 0 Then
			rs.AbsolutePage=paginacaocao
		Else
			paginacaocao = 1
			rs.AbsolutePage=1
		End If	
%>
	
    <%
	Dim sucesso, codmsg
	codmsg	= CStr(""&Request.QueryString("codmsg"))
	sucesso = CStr(""&Request.QueryString("sucesso"))
		
	If sucesso = "" and codmsg <> "" Then%>
    	<div align="center" class="msgGrandeSucesso">
        	<%
			If codmsg = "1" Then
				Response.Write("página incluída com sucesso!")
			Elseif codmsg = "2" Then
				Response.Write("página alterada com sucesso!")
			Elseif codmsg = "3" Then	
				Response.Write("página excluída com sucesso!")
			End If	
			%>
            
        
        </div>
    <%ElseIf sucesso <> "" and codmsg <> "" Then%>
        <div align="center" class="msgGrandeErro">
			<%
			If codmsg = "4" Then
				Response.Write("Vale Transporte NÃO alterada para NÃO haver duplicidade de informação!<BR>O conteúdo do campo """&Request.QueryString("chaveUnica")&""" já possui uma informação igual em outra Vale Transporte.<br />")
			Elseif codmsg = "5" Then
				Response.Write("Vale Transporte NÃO excluído, pois há funcionários que dependem deste Vale Transporte.<BR>Se NÃO desejar ter mais este Vale Transporte inative ao invés de excluir.<br />")
			
			'////////////////////////
			End If	
			%>
</div>
	<%End If%>
<table width="95%" border="0" align="center">
  <tr>
    <td width="47%"> 
página <strong><%=rs.AbsolutePage%></strong> de <%=rs.PageCount%>
</td>
    <td width="53%"><div align="right">Total <strong><%=rs.RecordCount%></strong> página(s)</div></td>
  </tr>
</table>
<table width="95%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
<td width="7%"><div align="center" style="font-weight: bold">Cód. página</div></td>
<td width="11%"><div align="center" style="font-weight: bold">Módulo</div></td>
<td width="23%"><div align="center" style="font-weight: bold">Nome página</div></td>
<td width="17%"><div align="center" style="font-weight: bold">URL</div></td>
<td width="11%"><div align="center" style="font-weight: bold">Habilitada para todos</div></td>
<td width="23%"><div align="center" style="font-weight: bold">Sub-Validações</div></td>
<td width="8%">Editar</td>
</tr>
<%
Dim x
For x=1 To rs.PageSize
If Not rs.Eof Then
%>
<tr <%If int(rs("habilitada_para_todos")) = 1 Then%> bgcolor="#CCCCCC" <%End If%>>
  <td><div align="center"><%=rs(0)%></div></td>
  <td><div align="center"><a href="javascript:alert('<%=rs("NOME_MODULO")%>')"><%=rs("SIGLA_MODULO")%></a></div></td>
  <td><div align="left"><%=rs("NOME_PAGINA")%></div></td>
  <td><div align="left"><a href="<%=rs("CAMINHO_VALIDA")%>" target="_blank"><%=rs("CAMINHO_VALIDA")%></a></div></td>
  <td><div align="center">
  <%=booleano(rs("habilitada_para_todos"))%>
  </div></td>
  <td><div align="center">
  <a href="adm_modulos_paginas_habilitadas_sub_validacoes.asp?cod_pagina=<%=rs("COD_PAGINA")%>&nome_pagina=<%=rs("NOME_PAGINA")%>">Alterar Validações</a><br />
              <%
			Set rsSubValidacoes = objConexao.Execute("SELECT * FROM TB_ADM_PAGINAS_HABILITADAS_SUB_VALIDACOES WHERE COD_PAGINA="&rs("COD_PAGINA"),qtdSubValidacoes)
				If not rsSubValidacoes.Eof Then%>
				<table width="100%" border="0" cellpadding="2" cellspacing="0">
                <%
				While not rsSubValidacoes.Eof
				%>
                <tr bgcolor="#D9E0FF">
                	<td width="15%"><div align="center"><%=rsSubValidacoes("COD_SUB_VALIDACAO")%></div></td>
                    <td width="85%"><div align="left"><%=rsSubValidacoes("NOME_SUB_VALIDACAO")%></div></td>
                </tr>
                <%
				rsSubValidacoes.MoveNext
				Wend%>
				</table>
				<%Else%>
				<div align="center">-</div>
				<%End If
				
            rsSubValidacoes.Close
			Set rsSubValidacoes = Nothing
			%>
            </div></td>
  <td>
  
  <table border="0" align="center" cellpadding="3" cellspacing="0">
  <tr>

    <td> 
    <a name="focus_pag_<%=rs(0)%>" href="adm_restrito_paginas_formulario.asp?comando=alterar&cod_pagina=<%=rs(0)%>">
  <img id="focus_pag_<%=rs(0)%>" src="images/icons/icon_alterar.png" width="16" height="16" border="0" />  	</a>  </td>
    <td><a href="javascript:confirmaExclusao('<%=rs(0)%>','<%=rs("CAMINHO_VALIDA")%>');">
  <img src="images/icons/icon_excluir.png" width="16" height="16" border="0" />
  </a></td>
  </tr>
</table>  </td>
  </tr>
  
  <%
  Set rsF=objConexao.Execute("SELECT TB_ADM_PAGINAS_HABILITADAS.*, TB_ADM_MODULOS.* from TB_ADM_MODULOS inner join TB_ADM_PAGINAS_HABILITADAS on TB_ADM_MODULOS.cod_modulo = TB_ADM_PAGINAS_HABILITADAS.cod_modulo  where pag_filha_cod_grupo_acesso_pagina = "&rs("COD_PAGINA")&" order by TB_ADM_MODULOS.order_MENU,TB_ADM_PAGINAS_HABILITADAS.caminho_valida")
  
  While not rsF.Eof
  %>
  <tr bgcolor="#99CCFF">
  <td><div align="center"><img src="images/layout/seta_superior.jpg" /><%=rsF(0)%></div></td>
  <td><div align="center"><a href="javascript:alert('<%=rsF("NOME_MODULO")%>')"><%=rsF("SIGLA_MODULO")%></a></div></td>
  <td><div align="left"><%=rsF("NOME_PAGINA")%></div></td>
  <td><div align="left"><a href="<%=rsF("CAMINHO_VALIDA")%>" target="_blank"><%=rsF("CAMINHO_VALIDA")%></a></div></td>
  <td><div align="center">
  <%=booleano(rsF("habilitada_para_todos"))%>
  </div></td>
  <td><div align="center">
  <a href="adm_modulos_paginas_habilitadas_sub_validacoes.asp?cod_pagina=<%=rsF("COD_PAGINA")%>&nome_pagina=<%=rsF("NOME_PAGINA")%>">Alterar Validações</a><br />
              <%
			Set rsSubValidacoes = objConexao.Execute("SELECT * FROM TB_ADM_PAGINAS_HABILITADAS_SUB_VALIDACOES WHERE COD_PAGINA="&rsF("COD_PAGINA"),qtdSubValidacoes)
				If not rsSubValidacoes.Eof Then%>
				<table width="100%" border="0" cellpadding="2" cellspacing="0">
                <%
				While not rsSubValidacoes.Eof
				%>
                <tr bgcolor="#D9E0FF">
                	<td width="15%"><div align="center"><%=rsSubValidacoes("COD_SUB_VALIDACAO")%></div></td>
                    <td width="85%"><div align="left"><%=rsSubValidacoes("NOME_SUB_VALIDACAO")%></div></td>
                </tr>
                <%
				rsSubValidacoes.MoveNext
				Wend%>
				</table>
				<%Else%>
				<div align="center">-</div>
				<%End If
				
            rsSubValidacoes.Close
			Set rsSubValidacoes = Nothing
			%>
            </div></td>
  <td>
  
  <table border="0" align="center" cellpadding="3" cellspacing="0">
  <tr>

    <td> 
    <a name="focus_pag_<%=rsF(0)%>" href="adm_restrito_paginas_formulario.asp?comando=alterar&cod_pagina=<%=rsF(0)%>">
  <img id="focus_pag_<%=rsF(0)%>" src="images/icons/icon_alterar.png" width="16" height="16" border="0" />  	</a>  </td>
    <td><a href="javascript:confirmaExclusao('<%=rsF(0)%>','<%=rsF("CAMINHO_VALIDA")%>');">
  <img src="images/icons/icon_excluir.png" width="16" height="16" border="0" />
  </a></td>
  </tr>
</table>  </td>
  </tr>
  
  <%
  rsF.MoveNext()
  Wend
  
  rsF.Close
  Set rsF=Nothing
  %>
  
<%
rs.MoveNext()
End If
Next
%>
	<%If rs.PageCount > 1 Then %>
                                        
            <tr bgcolor="whitesmoke">
      <td colspan="12">Ir para a página:<br /><div class="menu_paginacaocao">				
					<%
					Dim aae
					aae = 0
					For x=1 To rs.PageCount 
					aae = formatNumber(0&x,0)
					
						If int(paginacaocao) <> int(aae) Then
						%>
							<a href="?paginacao=<%=x%>"><%=x%></a>
	<%Else%>
							<span><%=x%></span>
						<%End If%>
						
					<%Next%>
					</div></td>
    </tr>
                        
                                
    <%End if%>
</table>	
<p><strong><br />
</strong></p>
<script type="text/javascript">
	function focusPagina(){
		<%If Request.QueryString("cod_pagina") <> Empty Then%>
			window.location.hash = "focus_pag_<%=Request.QueryString("cod_pagina")%>";
		<%End If%>
		
	}
	
	setTimeout("focusPagina()",1);
</script>
<%
	Else
	%>
    <div align="center" class="msgGrande">
    Nenhuma página encontrada! <br /></div>
    <%End if
	
	rs.Close
	Set rs=Nothing
	
	FecharConexao()
	%>

<!-- #include file="files/includes/miolo-inferior.asp" -->












