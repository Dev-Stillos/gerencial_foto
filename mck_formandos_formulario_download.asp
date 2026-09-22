<!-- #include virtual="gerencial_foto/includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include virtual="gerencial_foto/includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
abrirConexao()

dim c, t

dim rs, nome_arquivo_original, inicio_ext, len_arquivo, extencao, nome_arquivo, cod_mack_formando

c = request.queryString("c")


'response.write "SELECT frm.*, tmp.sequencia FROM TB_MACK_FORMANDOS  frm 	INNER JOIN TB_MACK_COLACOES col on frm.cod_mack_colacao = col.cod_mack_colacao	INNER JOIN TB_MACK_TEMPORADAS tmp on col.cod_mack_temporada = tmp.cod_mack_temporada where frm.cod_mack_formando = "&sqlNumero(cod_mack_formando)

Set rs = objConexao.Execute("SELECT frm.*, tmp.sequencia FROM TB_MACK_FORMANDOS  frm 	INNER JOIN TB_MACK_COLACOES col on frm.cod_mack_colacao = col.cod_mack_colacao	INNER JOIN TB_MACK_TEMPORADAS tmp on col.cod_mack_temporada = tmp.cod_mack_temporada where frm.cod_mack_formando = "&sqlNumero(c))


if not rs.Eof then


	nome_arquivo_original = rs("nome_arquivo")
	inicio_ext = instrrev(nome_arquivo_original,".")
	len_arquivo =  len(nome_arquivo_original)
	extencao = mid(nome_arquivo_original,inicio_ext+1,(len_arquivo-inicio_ext))


	nome_arquivo = rs("sequencia")& preencheCaracter(rs("numero_identificacao"),"0",6)&"."&extencao


	response.write caminho_dados&"mck_identificacao\"&nome_arquivo_original


	call downloadStreamRenomeia(caminho_dados&"mck_identificacao\"&nome_arquivo_original, nome_arquivo)

end if 
rs.close
set rs = nothing 

fecharConexao()
%>