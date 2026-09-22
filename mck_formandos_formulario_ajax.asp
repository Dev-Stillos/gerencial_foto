<% dim script_valida_form_master : script_valida_form_master = "NAO" %>
<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<!-- #include file="files/includes/functions_local.asp" -->

<%response.Charset="UTF-8"%>
<%

if session("cod_usuario") = "" then
	response.end
end if

dim retorno
dim t : t = request.form("t")
dim p(10)

for i = 0 to 10
	p(i) = request.form("p"&i)
next

if t = "buscaFormando" then
	'p(0) = tipo_dado
	'p(1) = valor
	retorno = buscaFormando(p(0), p(1), false)

elseif t = "valorProduto" then
	'p(0) = cod_mack_produto
	retorno = valorProduto(p(0), false)

elseif t = "selectColacoes" then
	'p(0) = cod_mack_temporada
	'p(0) = filtro_cod_mack_colacao
	retorno = selectColacoes(p(0), p(1))

elseif t = "validaTia" then
	'p(0) = tia
	retorno = validaTia(p(0), false)

end if

response.write(retorno)



'----------------------------------------------------------------------------------------------
function buscaFormando(tipo_busca, valor, conexao_aberta)

	Dim sql, retorno,rs

	if tipo_busca = "1" then
		sql = "SELECT nome, tia, curso, cod_mack_formando FROM TB_MACK_FORMANDOS WHERE cod_mack_colacao is not null and cod_mack_formando = "&sqlnumero(valor)
	elseif tipo_busca = "2" then
		sql = "SELECT nome, tia, curso, cod_mack_formando FROM TB_MACK_FORMANDOS WHERE cod_mack_colacao is not null and  tia = "&sqlnumero(valor)
	end if


	if not conexao_aberta then
		abrirConexao()
	end if

	Set rs = objConexao.execute(sql)
	if not rs.Eof then
		retorno = rs(0)&";"&rs(1)&";"&rs(2)&";"&rs(3)
	end if
	rs.close
	set rs = nothing

	if not conexao_aberta then
		fecharConexao()
	end if

buscaFormando = retorno
end function


function valorProduto(cod_mack_produto, conexao_aberta)
	
	if not conexao_aberta then
		abrirConexao()
	end if

	Dim rs, retorno

	Set rs = objConexao.execute("SELECT valor FROM TB_MACK_PRODUTOS WHERE cod_mack_produto = "&sqlnumero(cod_mack_produto))
	if not rs.Eof then
		retorno = rs(0)
	else
		response.end
	end if
	rs.close
	set rs = nothing

	if not conexao_aberta then
		fecharConexao()
	end if

valorProduto = retorno
end function



function selectColacoes(cod_mack_temporada, nome_campo)

dim classe, nome_campo_str

	if nome_campo = 0 then 
		classe = "inputTamanho2"
		nome_campo_str = "filtro_cod_mack_colacao"

	elseif nome_campo = 1 then
		classe = "inputTamanho3"
		nome_campo_str = "cod_mack_colacao"
	end if


if cod_mack_temporada <> "" then

	selectColacoes = padraoFormularioSelect(nome_campo_str,"id_"&nome_campo_str,classe,"","","Selecione", filtro_cod_mack_colacao,"SELECT cod_mack_colacao, concat(descricao ,' - ',cast(DATE_FORMAT(data_colacao,'%d/%m/%Y') as char)) FROM TB_MACK_COLACOES WHERE cod_mack_temporada = "&cod_mack_temporada&" ORDER BY descricao", "incluir", false, false, false)
	
else
	selectColacoes = "Selecione uma temporada"
end if


end function



function validaTia(tia, conexao_aberta)

	if not conexao_aberta then 
		abrirConexao()
	end if 


	Dim rs, retorno
	
	retorno = "false"

	if tia <> "" and len(tia) >= 6 then


		if tia = "000000" then
			retorno = "true"
		else
			set rs = objConexao.execute("SELECT cod_mack_formando FROM VW_TB_MACK_FORMANDOS WHERE tia = "&sqltexto(tia))
			if  rs.Eof then

				retorno = "true"

			end if
			rs.close
			set rs = nothing
		end if

	end if

	if not conexao_aberta then 
		fecharConexao()
	end if 

validaTia = retorno

end function




'------------------------------------------------------------------------------------------------------------
%>