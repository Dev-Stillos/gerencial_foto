<!-- #include virtual="gerencial_foto/includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include virtual="gerencial_foto/includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
abrirConexao()
t = request.queryString("t")
c = request.queryString("c")

Set rs = objConexao.Execute("SELECT arquivo_"&t&" FROM TB_VENDEDORES WHERE cod_vendedor = "&sqlNumero(c))
if not rs.Eof then
	call donwloadStream(caminho_dados&"vendedores\"&rs(0))
end if 
rs.close
set rs = nothing 

fecharConexao()
%>