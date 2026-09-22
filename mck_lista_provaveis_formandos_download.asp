<!-- #include virtual="gerencial_foto/includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include virtual="gerencial_foto/includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
abrirConexao()
cod_mack_lista_provaveis_formandos= request.queryString("cod_mack_lista_provaveis_formandos")

Set rs = objConexao.Execute("SELECT * FROM TB_MACK_LISTAS_PROVAVEIS_FORMANDOS WHERE cod_mack_lista_provaveis_formandos = "&sqlNumero(cod_mack_lista_provaveis_formandos))
if not rs.Eof then
	call downloadStreamRenomeia(caminho_dados&"mck_listas\"&rs("nome_arquivo_servidor")&"."&rs("extencao"), rs("nome_arquivo_original")&"."&rs("extencao"))
end if 
rs.close
set rs = nothing 

fecharConexao()
%>









