<% dim script_valida_form_master : script_valida_form_master = "NAO" %>
<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<!-- #include file="files/includes/functions_local.asp" -->
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
'TESTE| LOG
'(OK) | (OK)  trocarRespOS(cod_ordem_servico, cod_usuario_resp, trocar_prox_etapas, conexao_aberta)
'(OK) | (OK)  finalizarOS(cod_ordem_servico,conexao_aberta)
'(??) | (OK)  iniciarEtapa(cod_ordem_servico, conexao_aberta)
'(OK) | (OK)  cancelarOS(cod_ordem_servico, conexao_aberta)
'(OK) | (OK)  finalizarEtapa(cod_ordem_servico, conexao_aberta)
'(OK) | (OK)  voltarEtapa(cod_ordem_servico, cod_ordem_servico_etapa, conexao_aberta)

if t = "trocarRespOS" then
	'p(0) = cod_ordem_servico
	'p(1) = cod_ordem_servico_etapa
	'p(2) = trocar_prox_etapas
	'p(3) = cod_usuario_resp
	retorno = trocarRespOS(p(0), p(1), p(3), p(2), false)

elseif t = "finalizarOS" then 
	'p(0) = cod_ordem_servico
	retorno = finalizarOS(p(0), p(1), false)

elseif t = "iniciarEtapa" then
	'p(0) = cod_ordem_servico
	'p(1) = cod_ordem_servico_etapa
	retorno = iniciarEtapa(p(0), p(1), false)

elseif t = "cancelarOS" then
	'p(0) = cod_ordem_servico
	'p(2) = motivo_cancelado
	retorno = cancelarOS(p(0), p(2), false)

elseif t = "finalizarEtapa" then
	'p(0) = cod_ordem_servico
	'p(1) = cod_ordem_servico_etapa
	retorno = finalizarEtapa(p(0), p(1), false)

elseif t = "voltarEtapa" then
	'p(0) = cod_ordem_servico
	'p(1) = cod_ordem_servico_etapa
	retorno = voltarEtapa(p(0), p(1), false)

elseif t = "cancelarVariasOs" then
	'p(0) = cod_ordem_servicos (separados por ",")
	'p(2) = motivo_cancelado
	retorno = cancelarVariasOs(p(0), p(2), false)

elseif t = "alterarProduto" then
	'p(0) = cod_ordem_servico
	'p(2) = cod_produto
	retorno = alterarProduto(p(0), p(1),  p(2), false)

end if

response.write(retorno)



'----------------------------------------------------------------------------------------------
function alterarProduto(cod_ordem_servico, cod_ordem_servico_etapa, cod_produto, conexao_aberta)
	
	Dim retorno, rs	

	if not conexao_aberta then
		abrirConexao()
	end if

  if cod_produto = 34 then 
    descricao = "Produto Alterado de Mini álbum 2 para Mini álbum"
  else
    descricao = "Produto Alterado de Mini álbum para Mini álbum 2"
  end if
  

  if not IsEmpty(cod_ordem_servico) and not IsEmpty(cod_produto)  then
    objConexao.Execute("UPDATE TB_ORDENS_SERVICO SET cod_produto = "&sqlNumero(cod_produto)&" WHERE cod_ordem_servico = "&sqlNumero(cod_ordem_servico) ) 

    objConexao.Execute("INSERT INTO TB_ORDENS_SERVICO_HISTORICO (cod_ordem_servico, descricao, cod_ordem_servico_etapa,cod_usuario) VALUES ("&sqlNumero(cod_ordem_servico)&", "&sqlTexto(descricao)&", "&sqlTexto(cod_ordem_servico_etapa)&", "&sqlNumero(session("cod_usuario"))&")")

    retorno = "0|Produto alterado com sucesso"
  else 
    retorno = "1|Produto ou OS inválidos"
  end if 

	if not conexao_aberta then
		fecharConexao()
	end if
alterarProduto = retorno
end function


function trocarRespOS(cod_ordem_servico, cod_ordem_servico_etapa, trocar_prox_etapas, cod_usuario_resp, conexao_aberta)
	
	Dim retorno, rs	

	if not conexao_aberta then
		abrirConexao()
	end if

	'Validações
	if  validaUsuario(2,,56,true) then

		'vefiricamos que ação o usuário selecionou
		if trocar_prox_etapas = "S" then
			'SELECIONAMOS TOADAS AS ETAPAS EM ABERTO 
			Set rs=objConexao.Execute("SELECT cod_ordem_servico_etapa, cod_usuario_responsavel FROM TB_ORDENS_SERVICO_ETAPAS WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and data_termino IS NULL")
			If not rs.Eof then
				While not rs.Eof
					'TRANSFERIMOS PARA O USUARIO SELECIONADO
					objConexao.Execute("UPDATE TB_ORDENS_SERVICO_ETAPAS SET cod_usuario_responsavel = "&sqlTexto(cod_usuario_resp)&" WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&sqlTexto(rs(0)))
					
					'GRAVAMOS O LOG DA OCORRENCIA
					call logEventos("[cod_ordem_servico:"&cod_ordem_servico&"][cod_ordem_servico_etapa:"&cod_ordem_servico_etapa&"][cod_usuario:"&session("cod_usuario")&"][cod_usuario_antigo:"&rs("cod_usuario_responsavel")&"][INFO: Responsavel da etapa da OS foi alterado para o usuário "&nomeUsuario(cod_usuario_resp,true)&" ("&cod_usuario_resp&") ]", true)
				rs.MoveNext()
				Wend
			end if
			rs.close
			Set rs=nothing

			retorno = "0|Etapas alteradas com sucesso."
		else
			'BUSCAMOS A ETAPA ATUAL DA OS
			if cod_ordem_servico_etapa > 0 and cod_ordem_servico_etapa <> "" then 
				'ALTERAMOS SOMENTE A ETAPA ATUAL
				objConexao.Execute("UPDATE TB_ORDENS_SERVICO_ETAPAS SET cod_usuario_responsavel = "&sqlTexto(cod_usuario_resp)&" WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&sqlTexto(cod_ordem_servico_etapa))

				'GRAVAMOS O LOG DA OCORRENCIA
				call logEventos("[cod_ordem_servico:"&cod_ordem_servico&"][cod_ordem_servico_etapa:"&cod_ordem_servico_etapa&"][cod_usuario:"&session("cod_usuario")&"][INFO: Responsavel da etapa da OS foi alterado para o usuário "&nomeUsuario(cod_usuario_resp,true)&" ("&cod_usuario_resp&") ]", true)
				
				retorno = "0|Etapa alterada com sucesso."
			else
				retorno = "1|Etapa inválida"
			end if

			retorno = "0|Etapa alteradas com sucesso."
		end if 
	Else
		retorno = "1|Você não possui validação para trocar a responsabilidade desta OS."
	End if
	
	if not conexao_aberta then
		fecharConexao()
	end if

trocarRespOS = retorno
end function

function finalizarOS(cod_ordem_servico, cod_ordem_servico_etapa,conexao_aberta)

	Dim retorno, rs, cod_produto, quantidade

	if not conexao_aberta then
		abrirConexao()
	end if

	'FINALIZAMOS A ETAPA ATUAL
	objConexao.Execute( "UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_termino = NOW() WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&sqlTexto(cod_ordem_servico_etapa))

	call logEventos("[cod_ordem_servico_etapa:"& cod_ordem_servico_etapa_atual &"][cod_ordem_servico_etapa:"&cod_ordem_servico_etapa&"][cod_usuario:"&session("COD_USUARIO")&"][info:Etapa finaliziada]" , true)

	'Verifica se há etapas em aberto se tiver, pausa e informa erro
	Set rs=objConexao.Execute("SELECT * FROM TB_ORDENS_SERVICO_ETAPAS WHERE data_termino IS NULL AND cod_ordem_servico = "&sqlTexto(cod_ordem_servico))
		If not rs.Eof Then
			retorno = "1|Ordem de serviço não finalziada pois exitem etapas em aberto!"
		End If
	rs.Close
	Set rs=Nothing

	if retorno = "" then
		objConexao.Execute("UPDATE TB_ORDENS_SERVICO SET status_ordem_servico = 2, data_finalizado = NOW(), cod_usuario_finalizou = "&Session("cod_usuario")&" WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico))

		call logEventos("[cod_ordem_servico:"& cod_ordem_servico &"][cod_ordem_servico_etapa:"&cod_ordem_servico_etapa&"][cod_usuario:"&session("COD_USUARIO")&"][info:OS finaliziada]" , true)


		'Buscamos o cod_Produto para baixarmos no estoque
		Set rs = objConexao.Execute("SELECT * FROM TB_ORDENS_SERVICO WHERE cod_ordem_servico = "&sqlNumero(cod_ordem_servico))
		if not rs.Eof then
			cod_produto = rs("cod_produto")
			quantidade = rs("quantidade")
		end if
		rs.close
		set rs = Nothing

		'verificamos se o produto tem baixa na os (atualmente somente para DVD e prt)
		if estoqueMomentoBaixa(cod_produto, true) = "orden_servico" then
			call estoqueMovimentacao("S", orden_servico, cod_ordem_servico, cod_produto, quantidade, "Produção de OS.", true)
		end if

		retorno = "0|Ordem de serviço finalizada com sucesso"
	end if

	if not conexao_aberta then
		fecharConexao()
	end if

finalizarOS = retorno
end function

function iniciarEtapa(cod_ordem_servico, cod_ordem_servico_etapa, conexao_aberta)

	Dim retorno, rs
	Dim cod_ordem_servico_etapa_atual

	if not conexao_aberta then
		fecharConexao()
	end if

	'cod_ordem_servico_etapa_atual = etapaOS(cod_ordem_servico, true)
	
	objConexao.Execute("UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_inicio = NOW() WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&sqlTexto(cod_ordem_servico_etapa_atual))
	
	call logEventos("[cod_ordem_servico:"& cod_ordem_servico &"][cod_ordem_servico_etapa:"&cod_ordem_servico_etapa&"][cod_usuario:"&session("COD_USUARIO")&"][info:Etapa Iniciada]" , true)


	retorno = "0|Etapa iniciada com sucesso!"

	if not conexao_aberta then
		fecharConexao()
	end if

iniciarEtapa = retorno
end function

function cancelarOS(cod_ordem_servico, motivo_cancelada, conexao_aberta)

	Dim retorno, rs
	Dim cod_ordem_servico_etapa_atual

	if not conexao_aberta then
		abrirConexao()
	end if

	if not validaUsuario(2,,55,false) then
		retorno = "1|Você não possui acesso a cancelamento."
	end if

	if motivo_cancelada = "" then
		retorno = "1|Informe o motivo dos  cancelamento."
	end if


	if retorno = "" then
		objConexao.Execute("UPDATE TB_ORDENS_SERVICO SET status_ordem_servico = 99, data_cancelado = NOW(), cod_usuario_cancelou = "&Session("cod_usuario")&", motivo_cancelada = "&sqlTexto(motivo_cancelada)&" WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico))

		call logEventos("[cod_ordem_servico:"& cod_ordem_servico &"][cod_ordem_servico_etapa:"&cod_ordem_servico_etapa&"][cod_usuario:"&session("COD_USUARIO")&"][info:OS Cancelada]" , true)
	
		retorno = "0|Ordem de serviço cancelada."
	end if

	if not conexao_aberta then
		fecharConexao()
	end if

cancelarOS = retorno
end function

function finalizarEtapa(cod_ordem_servico, cod_ordem_servico_etapa,conexao_aberta)
	
	Dim retorno, rs

	if not conexao_aberta then
		abrirConexao()
	end if

	'Finaliza a etapa atual
	objConexao.Execute("UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_termino = NOW() WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&sqlTexto(cod_ordem_servico_etapa))

	call logEventos("[cod_ordem_servico_etapa:"& cod_ordem_servico_etapa &"][cod_ordem_servico_etapa:"&cod_ordem_servico_etapa&"][cod_usuario:"&session("COD_USUARIO")&"][info:Etapa finaliziada]" , true)

	'BUSCAMOS A PROXIMA ETAPA
	Set rs=objConexao.Execute("SELECT cod_ordem_servico_etapa FROM TB_ORDENS_SERVICO_ETAPAS WHERE cod_ordem_servico = "&cod_ordem_servico&" and numero_etapa IN(SELECT (numero_etapa+1) FROM VW_TB_ORDENS_SERVICO WHERE cod_ordem_servico = "&cod_ordem_servico&")")
	If not rs.Eof Then
		'INICIAMOS A ETAPA ENCOTRADA
		objConexao.Execute("UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_inicio = NOW() WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&rs("cod_ordem_servico_etapa"))

		'ATUALIZAMOS A TABELA DE ORDENS DE SERVICO
		objConexao.Execute("UPDATE TB_ORDENS_SERVICO SET cod_ordem_servico_etapa_atual = "&rs("cod_ordem_servico_etapa")&" WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico))

		call logEventos("[cod_ordem_servico:"&cod_ordem_servico&"][cod_ordem_servico_etapa:"& rs("cod_ordem_servico_etapa")&"][cod_usuario_resp:"&session("COD_USUARIO")&"][info:Etapa Inciada]" , true)
	End If
	rs.Close
	Set rs=Nothing

	retorno = "0|Etapa finalizada com sucesso!"

	if not conexao_aberta then
		fecharConexao()
	end if

finalizarEtapa = retorno
end function


function voltarEtapa(cod_ordem_servico, cod_ordem_servico_etapa, conexao_aberta)
	
	Dim retorno, rs

	if not conexao_aberta then
		abrirConexao()
	end if

	if validaUsuario(2,,67,true) then

		'Verificamos se a etapa selecionada confere com a ordem de serviço
		Set rs = objConexao.Execute("SELECT cod_ordem_servico FROM TB_ORDENS_SERVICO_ETAPAS WHERE cod_ordem_servico_etapa = "&sqlNumero (cod_ordem_servico_etapa) &" AND cod_ordem_servico = "&sqlNumero(cod_ordem_servico) )
		if not rs.Eof then

			'1) Voltamos para a etapa selecionada;
			objConexao.Execute("UPDATE TB_ORDENS_SERVICO SET cod_ordem_servico_etapa_atual = "&sqlTexto(cod_ordem_servico_etapa)&" WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico))	
			
			'2) Limpa as etapas posteiores a esta 	
			objConexao.Execute("UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_termino = NULL WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa >= "&sqlTexto(cod_ordem_servico_etapa))

			call logEventos("[cod_ordem_servico:"&cod_ordem_servico&"][cod_ordem_servico_etapa:"&cod_ordem_servico_etapa&"][info:Etapa alterada]", true)

			retorno = "0|Etapa voltada com sucesso!"
		Else
			retorno = "1|A etapa que você está querendo voltar não pertence a esta OS."
		end if
	else
		retorno = "1|Você não tem acesso a opção de voltar uma etapa."
	end if

	if not conexao_aberta then
		fecharConexao()
	end if

voltarEtapa = retorno
end function

function cancelarVariasOs(codigos_os, motivo_cancelada, conexao_aberta)
'codigos_os: contem os codigos de OS separados por ","

'Regras:
'Para que possamos cancelar um lote de os todas devem ser do mesmo pedido

	Dim rs, retorno, aux, aux2, i, qtd_pedidos, erro, erro_msgs
	aux = split(codigos_os, ",")

	if not conexao_aberta then
		abrirConexao()
	end if

	'verificamos se todos os parametros sao numericos
	for i=0 to ubound(aux)
		if not isnumeric(aux(i)) then
			retorno = "1|Um erro inesperado ocorreu, tente novamente mais tarde"
		end if
	next

	if retorno = "" then
		call objConexao.Execute("SELECT cod_pedido FROM TB_ORDENS_SERVICO WHERE cod_ordem_servico in ("&codigos_os&") GROUP BY cod_pedido ", qtd_pedidos )
		

		if qtd_pedidos > 1 or qtd_pedidos = 0 then
			retorno = "1|Lote de cancelamento inválido, tente novamente mais tarde."
		else
			for i=0 to ubound(aux)
				aux2 = cancelarOS(aux(i), motivo_cancelada, true)

				if split(aux2,"|")(0) = "1" then
	 				erro = true
					erro_msgs = aux(i) &": "&split(aux2,"|")(1)
				end if

				if erro then
					retorno = "1|Uma ou mais OS não puderam ser canceladas." & erro_msgs
				else
					retorno = "0|Todas as OS foram canceladas."
				end if
			next
		end if
	end if


	if not conexao_aberta then
		fecharConexao()
	end if

cancelarVariasOs = retorno
end function

'------------------------------------------------------------------------------------------------------------
%>