<%
TIPO_BUSCA	= formatNumber(0&Request.QueryString("TIPO_BUSCA"),0)


AbrirConexao()
Dim BUSCA_35, BUSCA_36, BUSCA_37, BUSCA_38, BUSCA_39, BUSCA_40, BUSCA_90, POSSUI_BUSCA
BUSCA_35	= False 
BUSCA_36	= False 
BUSCA_37	= False 
BUSCA_38	= False 
BUSCA_39	= False 
BUSCA_40	= False
BUSCA_90    = False 
POSSUI_BUSCA = False

If IsNumeric(Session("COD_PERFIL")) Then
    Set rs = objConexao.Execute("SELECT COD_PERFIL, COD_SUB_VALIDACAO FROM TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES WHERE COD_PERFIL=" & Session("COD_PERFIL") & " order by COD_SUB_VALIDACAO asc")
Else
    Response.Write "Erro: COD_PERFIL inv&aacute;lido ou n&atilde;o definido na sess&atilde;o."
    Response.End
End If

		
		While not rs.Eof
						
			If int(rs("COD_SUB_VALIDACAO")) = 35 Then
			'Busca por C?d. Pedido
				BUSCA_35 = True
				POSSUI_BUSCA = True
				
				'Padr?o para quem possui esta sub-valida??o
				If int(TIPO_BUSCA) = 0 Then
					TIPO_BUSCA = 1
				End If
				
			End If	
			
			If int(rs("COD_SUB_VALIDACAO")) = 36 Then
			'Busca por Nome do Formando
				BUSCA_36 = True
				POSSUI_BUSCA = True
			End If	
			
			If int(rs("COD_SUB_VALIDACAO")) = 37 Then
			'Busca por C?d. Formando Ficha
				BUSCA_37 = True
				POSSUI_BUSCA = True
				
				'Padr?o para quem possui esta sub-valida??o, lembrando que a prioridade ? c?d. pedido 
				If int(TIPO_BUSCA) = 0 Then
					TIPO_BUSCA = 8
				End If
				
				
			End If	
			
			If int(rs("COD_SUB_VALIDACAO")) = 38 Then
			'Busca por N?mero do ?lbum
				BUSCA_38 = True
				POSSUI_BUSCA = True
			End If	
			
			If int(rs("COD_SUB_VALIDACAO")) = 39 Then
			'Busca por Contrato
				BUSCA_39 = True
				POSSUI_BUSCA = True
			End If	

			If int(rs("COD_SUB_VALIDACAO")) = 40 Then
			'Busca por C?d. da Produ??o
				BUSCA_40 = True
				POSSUI_BUSCA = True
			End If					

            If int(rs("COD_SUB_VALIDACAO")) = 90 Then
            'Busca por CPF
                BUSCA_90 = True
                POSSUI_BUSCA = True

            'Se quiser deixar CPF como padrao ao abrir
                If int(TIPO_BUSCA) = 0 Then
                    TIPO_BUSCA = 9  ' Ou outro numero livre
                End If
            End if    
			
		rs.MoveNext()
		Wend
		
rs.Close
Set rs=Nothing
FecharConexao()
%>
<% If POSSUI_BUSCA = True Then %>
<form name="buscaForm" action="adm_busca_completa_depto_foto.asp" class="dashboard-search" role="search">

    <select aria-label="Tipo de busca" name="TIPO_BUSCA" class="form-select"
        style="background-color: white; color: black;">
        <% If BUSCA_35 = True Then %>
            <option <% If TIPO_BUSCA = 1 Then %> selected="selected" <% End If %> value="1">C&oacute;d. pedido</option>
        <% End If %>
        <% If BUSCA_36 = True Then %>
            <option <% If TIPO_BUSCA = 2 Then %> selected="selected" <% End If %> value="2">Nome do formando</option>
        <% End If %>
        <% If BUSCA_37 = True Then %>
            <option <% If TIPO_BUSCA = 8 Then %> selected="selected" <% End If %> value="8">C&oacute;d. formando (ficha)</option>
        <% End If %>
        <% If BUSCA_38 = True Then %>
            <option <% If TIPO_BUSCA = 3 Then %> selected="selected" <% End If %> value="3">N&uacute;m. &aacute;lbum</option>
        <% End If %>
        <% If BUSCA_39 = True Then %>
            <option <% If TIPO_BUSCA = 5 Then %> selected="selected" <% End If %> value="5">Contrato</option>
        <% End If %>
        <% If BUSCA_40 = True Then %>
            <option <% If TIPO_BUSCA = 7 Then %> selected="selected" <% End If %> value="7">C&oacute;d. produ&ccedil;&atilde;o</option>
        <% End If %>
        <% If BUSCA_90 = True Then %>
            <option <% If TIPO_BUSCA = 9 Then %> selected="selected" <% End If %> value="9">CPF</option>
        <% End If %>

    </select>

    <input name="BUSCAR_TERMO" value="<%=Server.HTMLEncode(BUSCAR_TERMO & "")%>" type="text" maxlength="100" class="form-control" aria-label="Termo da busca" placeholder='Digite aqui...'
        style="background-color: white; color: black;" />

    <input type="submit" value="Buscar" class="btn btn-primary"
        style="background-color: #ff007f; color: white;" />
</form>
<% End If %>

