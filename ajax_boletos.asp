<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/sha256.asp" -->
<!-- #include file="includes/classe_modulo_cob.asp" -->
<!-- #include file="includes/classe_fechamento.asp" -->
<%
if session("COD_VENDEDOR") = "" Then
	'Response.redirect("index.asp")
end if

dim p(15), tipo, retorno
tipo = request.form("tp")
for i = 0 to 15 
	p(i) = request.form("p"&i)
next


if (tipo = "valida_vencimento") then
	'vencimentoValido(dia, mes, ano)
	retorno = vencimentoValido(p(1), p(2), p(3))

elseif (tipo = "alteraParcelamento") then
	retorno = alteraParcelamento(p(1), p(2), p(3), p(4), p(5), false)

elseif (tipo = "excluirboletosSacados") then
	retorno = excluirboletosSacados(p(1), false)

elseif (tipo = "alteraSacado") then
' p1: cod_sacado 
' p2: cod_formando
' p3: nome_formando
' p4: cpf
' p5: email
' p6: cep
' p7: endereco
' p8: numero
' p9: bairro
' p10: cidade
' p11: estado
' p12: grau_parentesco
	retorno = alterasacado(p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), flase)

elseif (tipo = "incluirBoleto") then
	retorno = incluirBoleto(p(0), p(1), p(2), p(3), false)

elseif (tipo = "excluirBoleto") then
	retorno = excluirBoleto(p(0), false)

elseif (tipo = "baixarBoleto") then
	retorno = baixarBoleto(p(0), p(1), p(2), p(3), false)

elseif (tipo = "baixarBoeltoManual") then
	retorno = baixarBoeltoManual(p(0), p(1), p(2), false)	

end if


Response.write retorno


'----------------------------------------------------
function baixarBoeltoManual(nosso_numero_p, valor_pago_p, data_pagamento_p, conexao_aberta)
	
	dim retorno, rs

	dim nosso_numero_tratado, cod_pedido, parcela, vencimento

	if not conexao_aberta then
		abrirConexao()
	end if

	if len(nosso_numero_p) < 7 then
		baixarBoeltoManual = "false#nosso numero inválido"
		exit function
	end if 

	nosso_numero_tratado = preencheCaracter(nosso_numero_p,"0",8)
	cod_pedido = int(left(nosso_numero_tratado,6))
	parcela = right(nosso_numero_tratado,2)


	dim cod_boleto, cod_sacado, cod_vendedor, nosso_numero, hash, valor, valor_pago, valor_deveria_ser_pago, data_vencimento, data_pagamento, data_baixa, baixa_manual, baixa_manual_usuario, valor_comissao_vendedor, vendedor_recebe_em, cod_fechamento_periodo, pagamento_processado
	
	set rs = objConexao.Execute("SELECT * FROM TB_PEDIDOS_BOLETOS WHERE cod_pedido = "&sqlnumero(cod_pedido)&" AND parcela = "&sqlnumero(parcela))
	if not rs.Eof then
		'cod_pedido = rs("cod_pedido")
		'parcela = rs("parcela")
		cod_boleto = "" & rs("cod_boleto")
		cod_sacado = "" & rs("cod_sacado")
		cod_vendedor = "" & rs("cod_vendedor")
		nosso_numero = "" & rs("nosso_numero")
		hash = "" & rs("hash")
		valor = "" & rs("valor")
		valor_pago = "" & rs("valor_pago")
		valor_deveria_ser_pago = "" & rs("valor_deveria_ser_pago")
		data_vencimento = ""
		data_pagamento = "" & rs("data_pagamento")
		data_baixa = "" & rs("data_baixa")
		baixa_manual = "" & rs("baixa_manual")
		baixa_manual_usuario = "" & rs("baixa_manual_usuario")
		valor_comissao_vendedor = "" & rs("valor_comissao_vendedor")
		vendedor_recebe_em = "" & rs("vendedor_recebe_em")
		cod_fechamento_periodo = "" & rs("cod_fechamento_periodo")
		pagamento_processado = "" & rs("pagamento_processado")
	else
		baixarBoeltoManual = "false#Parcela não encontrada "&chr(10)&"Pedido:"&cod_pedido&chr(10)&"Parcela:"&parcela
		exit function

	end if
	rs.close
	set rs = nothing



	if data_pagamento <> "" and valor_pago <> ""  then
		baixarBoeltoManual = "false#Pagamento já processado"
		exit function
	end if 


	dim motivo_baixa
	data_pagamento = data_pagamento_p
	valor_pago = valor_pago_p
	motivo_baixa = "Baixa efetuada manualmente, através da página de baixa manual."


	'Se tudo estivet certo entao baixamos
	ret = baixarBoleto(cod_boleto, valor_pago, data_pagamento, motivo_baixa_1, true)

	if split(ret, "|")(0) = "true" then
		retorno = "true#1|"&cod_pedido&"|"&parcela&"|"&arrumaNumero(valor,"0,00",2)&"|"&arrumaNumero(valor_pago,"0,00",2)&"|"&arrumadata(vencimento)&"|"&arrumadata(data_pagamento)
	else
		retorno = replace(ret,"|", "#")
	end if


	if not conexao_aberta then
		fecharConexao()
	end if

baixarBoeltoManual = retorno
end function



function baixarBoleto(cod_boleto, valor_pago, data_pagamento, motivo_baixa, conexao_aberta)
	dim retorno, cod_formando, parcela, nosso_numero, cod_pedido

	if not conexao_aberta then
		abrirConexao()
	end if

	'1º verificamos se o motivo foi preechido
	'2º verificamos se o boleto já foi baixado

	'1º verificamos se o motivo foi preechido
	if motivo_baixa <> "" then
		baixarBoleto = "false|Motivo não informado"
		exit function
	end if

	'2º Verificamos se o boleto já foi baixado
	'Response.write "SELECT * FROM TB_PEDIDOS_BOLETOS WHERE cod_boleto = "&sqlNumero(cod_boleto)&" AND data_pagamento is not null <hr>" 
	set rs = objConexao.execute("SELECT * FROM TB_PEDIDOS_BOLETOS WHERE cod_boleto = "&sqlNumero(cod_boleto)&" AND data_pagamento is not null")
	if not rs.Eof then
		retorno = "false|Boleto já baixado."
	end if
	rs.close
	set rs = nothing

	if retorno <> "" then
		baixarBoleto = retorno
		exit function
	end if

	Set rs = objConexao.execute("SELECT min(data_pagamento) from TB_FECHAMENTO_PAGAMENTO_PERIODOS where data_pagamento not in (SELECT VENDEDOR_RECEBE_EM FROM TB_FECHAMENTO_PAGAMENTO_CABECALHO)")
	if not rs.Eof then
		vendedor_recebe_em = rs(0)
	end if
	rs.close
	set rs = nothing

		', data_pagamento, motivo_baixa
	objConexao.Execute("UPDATE TB_PEDIDOS_BOLETOS SET data_baixa = NOW(), valor_pago = "&sqlNumero(valor_pago)&",vendedor_recebe_em = "&sqlData(vendedor_recebe_em)&",  data_pagamento = "&sqlData(data_pagamento)&", baixa_manual = 'S', baixa_manual_usuario = "&sqlNumero(session("cod_usuario"))&", pagamento_processado = 0 where cod_boleto = "&sqlNumero(cod_boleto))

	retorno = "true|boleto baixado com sucesso."

	'buscamos infromacoes de log
	set rs = objConexao.Execute("SELECT parcela, nosso_numero, cod_pedido FROM TB_PEDIDOS_BOLETOS where cod_boleto = "&sqlnumero(cod_boleto))
	if not rs.Eof then
		parcela = rs("parcela")
		nosso_numero = rs("nosso_numero")
		cod_pedido = rs("cod_pedido")
	end if
	rs.close
	set rs = nothing

	call logEventos("[acao:baixa_boleto_manual][cod_pedido="&cod_pedido&"][motivo_baixa:"&motivo_baixa&"][data_pagamento:"&data_pagamento&"][valor_pago:"&valor_pago&"][cod_boleto:"&cod_boleto&"][parcela:"&parcela&"][nosso_numero:"&nosso_numero&"][cod_usuario:"&session("COD_USUARIO")&"]" , true)

	if not conexao_aberta then
		abrirConexao()
	end if

	baixarBoleto = retorno
end function


function excluirBoleto(cod_boleto, conexao_aberta)
	dim retorno, rs
	dim parcela, valor, data_vencimento, cod_formando, nosso_numero, cod_pedido
	retorno = ""

	if not conexao_aberta then
		abrirConexao()
	end if

	'1º Verificamos se já não foi baixada
	set rs = objConexao.execute("SELECT * FROM TB_PEDIDOS_BOLETOS WHERE cod_boleto = "&sqlNumero(cod_boleto)&" AND data_pagamento is not null")
	if not rs.Eof then
		retorno = "false|Não é possivel excluir parcelas já baixados."
	end if
	rs.close
	set rs = nothing

	if retorno <> "" then
		excluirBoleto = retorno
		exit function
	end if

	objConexao.Execute("DELETE FROM TB_PEDIDOS_BOLETOS WHERE cod_boleto = "&sqlNumero(cod_boleto)&" AND data_pagamento is null")
	retorno = "true|Parcela excluido com sucesso."


	set rs = objConexao.Execute("SELECT * FROM TB_PEDIDOS_BOLETOS WHERE cod_boleto = "&sqlnumero(cod_boleto))
	if not rs.Eof then
		parcela = rs("parcela")
		valor = rs("valor")
		data_vencimento = rs("data_vencimento")
		cod_boleto = rs("cod_boleto")
		cod_pedido = rs("cod_pedido")
		nosso_numero = rs("nosso_numero")
	end if
	rs.close
	set rs = nothing

	call logEventos("[acao:exclusao_boleto][cod_pedido:"&cod_pedido&"][valor:"&valor&"][data_vencimento:"&data_vencimento&"][cod_boleto:"&cod_boleto&"][parcela:"&parcela&"][nosso_numero:"&nosso_numero&"][cod_usuario:"&session("COD_USUARIO")&"]" , true)

	if not conexao_aberta then
		abrirConexao()
	end if

	excluirBoleto = retorno
end function


function incluirBoleto(cod_pedido, parcela, vencimento, valor, conexao_aberta)

	Dim rs, retorno
	if not conexao_aberta then
		abrirConexao()
	end if

	'1º Verificamos se o valor não irá ultrapassar o valor em produtos
	'2º verificamos se não existe parcela duplicada
	'3º Verificamos a comissao a ser paga
	'4º Buscamos sacado, vendedor, nosso_numero, hash


	'1º VERIFICACAO'
	Dim rsProdutosTotal, QTD_EM_PRODUTOS, XTOTAL_AGREGADOS, XTOTAL_ALBUNS, XTOTAL_PRODUTOS_A_PARTE, XTOTAL_DESCONTO, QTD_EM_PAGAMENTO

	QTD_EM_PRODUTOS = 0
	XTOTAL_AGREGADOS = 0
	XTOTAL_ALBUNS = 0
	XTOTAL_PRODUTOS_A_PARTE = 0
	XTOTAL_DESCONTO = 0
	QTD_EM_PAGAMENTO = 0
	
	Set rsProdutosTotal = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_AGREGADOS WHERE COD_PEDIDO="&COD_PEDIDO)
	if not rsProdutosTotal.Eof Then
		XTOTAL_AGREGADOS = cdbl(0&rsProdutosTotal("XTOTAL_AGREGADOS"))
	End If
	rsProdutosTotal.Close()
	Set rsProdutosTotal = Nothing
	
	Set rsProdutosTotal = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_ALBUNS WHERE COD_PEDIDO="&COD_PEDIDO)
	if not rsProdutosTotal.Eof Then
		XTOTAL_ALBUNS = cdbl(0&rsProdutosTotal("XTOTAL_ALBUNS"))
	End If
	rsProdutosTotal.Close()
	Set rsProdutosTotal = Nothing
	
	Set rsProdutosTotal = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_PRODUTOS_A_PARTE WHERE COD_PEDIDO="&COD_PEDIDO)
	if not rsProdutosTotal.Eof Then
		XTOTAL_PRODUTOS_A_PARTE = cdbl(0&rsProdutosTotal("XTOTAL_PRODUTOS_A_PARTE"))
	End If
	rsProdutosTotal.Close()
	Set rsProdutosTotal = Nothing

	set rsProdutosTotal = objConexao.Execute("select SUM(valor_total_com_desconto) as valor_total_com_desconto from(SELECT SUM(faturado * (valor_un-(valor_un * (1-(ifnull(aliquota_desconto_foto,0)/100))))) as valor_total_com_desconto FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO in (SELECT cod_formando FROM TB_PEDIDOS_FORMANDOS WHERE cod_pedido =  "&cod_pedido&") union all SELECT SUM(quantidade * (valor_un-(valor_un * (1-(ifnull(aliquota,0)/100))))) as valor_total_com_desconto  FROM TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO = "&cod_pedido&") as total_desconto")
	if not rsProdutosTotal.Eof then
		XTOTAL_DESCONTO = cdbl(0&rsProdutosTotal(0))
	end if
	rsProdutosTotal.Close
	set rsProdutosTotal = Nothing
	QTD_EM_PRODUTOS = cdbl(XTOTAL_AGREGADOS + XTOTAL_ALBUNS + XTOTAL_PRODUTOS_A_PARTE - XTOTAL_DESCONTO)

	set rs = objConexao.execute("SELECT ifnull(XTOTAL_BOLETO,0) + ifnull(XTOTAL_CHEQUE,0) + ifnull(PAGAMENTO_EM_DINHEIRO,0) + ifnull(XTOTAL_CARTAO,0) FROM VW_PEDIDOS_TOTAL_RECEBIDO WHERE COD_PEDIDO = "&sqlNumero(cod_pedido))
	if not rs.Eof then
		QTD_EM_PAGAMENTO = cdbl(0&rs(0))
	end if
	rs.Close
	set rs = Nothing


	if (QTD_EM_PAGAMENTO + valor) > QTD_EM_PRODUTOS then
		incluirBoleto = "false|Parcela não incluida pois irá ultrapassar o valor de produtos."
		exit function
	end if 
	'FIM 1º VERIFICACAO

	'2º VERIFICACAO'
	set rs = objConexao.Execute("SELECT * FROM TB_PEDIDOS_BOLETOS WHERE parcela = "&sqlnumero(parcela)&" AND cod_pedido = "&sqlNumero(cod_pedido))
	if not rs.Eof then
		retorno = "false|A parcela "&parcela&" já existe."
	end if
	rs.close
	set rs = nothing

	if retorno <> "" then
		incluirBoleto = retorno
		exit function
	end if 
	'FIM 2º VERIFICACAO
	
	'3º VERIFICACAO'
	dim representacao, valor_comissao_parcela, valor_total_boletos
 	valor_total_comissao = cdbl(fechamentoRetornaValorComissaoPedido(cod_pedido,1,true))
	representacao = valor / QTD_EM_PRODUTOS
	valor_comissao_parcela = representacao * valor_total_comissao
	valor_total_boletos = 0


	set rs = objConexao.execute("SELECT sum(valor_comissao_vendedor) FROM TB_PEDIDOS_BOLETOS WHERE COD_PEDIDO = "&sqlnumero(cod_pedido))
	if not rs.eof then
		valor_total_boletos = cdbl(0&rs(0))
	end if 
	rs.close()
	set rs = nothing



	diferenca = (valor_total_boletos + valor_comissao_parcela ) - valor_total_comissao


	if diferenca > 0 then
		valor_comissao_parcela =valor_comissao_parcela -  diferenca
	end if
	'FIM 3º VERIFICACAO

	'4º VERIFICACAO
	dim nosso_numero, hash, cod_vendedor_pedido, cod_sacado_pedido
	nosso_numero = cod_pedido & preencheCaracter(parcela, "0", 2)

	set rs = objConexao.execute("SELECT cod_sacado, cod_vendedor FROM TB_PEDIDOS WHERE cod_pedido = " & sqlnumero(cod_pedido))
	if not rs.Eof Then
		cod_vendedor_pedido = rs("cod_vendedor")
		cod_sacado_pedido = rs("cod_sacado")
	end if
	rs.close
	set rs = nothing

	hash = sha256(nosso_numero & hashkey)
	'FIM 4º VERIFICACAO

	objConexao.Execute("INSERT INTO TB_PEDIDOS_BOLETOS (cod_pedido, parcela, data_vencimento, valor, valor_comissao_vendedor, pagamento_processado, hash, cod_sacado, cod_vendedor, nosso_numero) VALUES ("&sqlnumero(cod_pedido)&", "&sqlNumero(parcela)&", "&sqlData(vencimento)&", "&sqlnumero(valor)&", "&sqlnumero(valor_comissao_parcela)&", 0, "&sqlTexto(hash)&","&sqlnumero(cod_sacado_pedido)&","&sqlnumero(cod_vendedor_pedido)&", "&sqlTexto(nosso_numero)&")")
	retorno = "true|Parcela incluida com sucesso!"

	call logEventos("[acao:inclusao_boleto][cod_pedido:"&cod_pedido&"][valor:"&valor&"][data_vencimento:"&data_vencimento&"][parcela:"&parcela&"][nosso_numero:"&nosso_numero&"][cod_sacado_pedido:"&cod_sacado_pedido&"][valor_comissao_parcela:"&valor_comissao_parcela&"][cod_vendedor_pedido:"&cod_vendedor_pedido&"][cod_usuario:"&session("COD_USUARIO")&"]" , true)

	if not conexao_aberta then
		abrirConexao()
	end if

	incluirBoleto = retorno
end function

'##########################################################################################
'##########################################################################################
'##########################################################################################

function vencimentoValido(dia, mes, ano)

	dim data_vencimento
	data_vencimento = CDate(ano&"-"&mes&"-"&dia)

 	dim diferenca 
 	diferenca = datediff("d", now(), data_vencimento)


 	if diferenca > 0 Then
 		vencimentoValido = true
 	else
 		vencimentoValido = false
 	end if

end function

function buscaFormando(cod_pedido, nome_formando)

Dim retorno, rs


set rs = objConexao.Execute("select * from tb_pedidos_formandos inner join tb_fichas_identificacao on tb_fichas_identificacao.cod_formando = tb_pedidos_formandos.cod_formando")
if not rs.Eof then

end if 
rs.close
set rs = nothing


buscaFormando = retorno
end function


function alteraParcelamento(cod_pedido, dia_vencimento, mes_vencimento, ano_vencimento, qtd_parcelas, conexao_aberta)

	dim retorno, rs, cod_sacado, cod_vendedor

	if not conexao_aberta then
		abrirConexao()
	end if

	'verificamos se o cod_do sacado consta no cabeçalho do pedido'
	set rs = objConexao.Execute("SELECT cod_sacado FROM TB_PEDIDOS WHERE cod_pedido = "&sqlNumero(cod_pedido))
	if not rs.eof then
		if cstr(""&rs("cod_sacado")) = "" then
			
			Set rs2 = objConexao.Execute("SELECT  cod_sacado FROM TB_PEDIDOS_BOLETOS WHERE cod_pedido = "&sqlNumero(cod_pedido)&" limit 1")
			if not rs2.eof then
				cod_sacado = rs2(0)
			end if
			rs2.close
			set rs2 = nothing
			objConexao.Execute("UPDATE TB_PEDIDOS SET cod_sacado = "&sqlNumero(cod_sacado)&" WHERE cod_pedido = "&sqlNumero(cod_pedido))
		end if
	end if
	rs.close()
	set rs = nothing

	'buscamos o cod_sacado
	set rs = objConexao.Execute("SELECT cod_sacado, cod_vendedor FROM TB_PEDIDOS WHERE cod_pedido = "&sqlNumero(cod_pedido))
	if not rs.Eof then
		cod_sacado = rs(0)
		cod_vendedor = rs(1)
	end if
	rs.close()
	set rs = nothing


	'alteraamos os boletos
	'deletamos todos
	objConexao.Execute("DELETE FROM TB_PEDIDOS_BOLETOS WHERE cod_pedido = "&sqlNumero(cod_pedido) &" and data_pagamento is null")

	'geramos novos
	'GERAMOS OS BOLETOS SEM A COMISSAO
	dim qtd_parcelas_pagas : qtd_parcelas_pagas = 0
	dim valor_pago : valor_pago = 0 
	'buscamos parcelas já pagas para deduzir se houver algo pago
	set rs = objConexao.Execute("SELECT count(parcela) as qtd_parcelas_pagas, sum(valor) as valor_pago FROM TB_PEDIDOS_BOLETOS WHERE data_pagamento is not null and cod_pedido = "&sqlNumero(cod_pedido))
	if not rs.Eof Then
		qtd_parcelas_pagas = int(0&rs("qtd_parcelas_pagas"))
		valor_pago  = cdbl(0&rs("valor_pago"))
	end if
	rs.Close
	set rs = Nothing

	dim qtd_parcelas_x
	qtd_parcelas_x = qtd_parcelas - qtd_parcelas_pagas
	valor_apagar = cdbl(parcelasBoleto(cod_pedido, "valor_apagar", true)) - valor_pago

	valor_parcela_ok = round(valor_apagar/qtd_parcelas_x,2)
	dim diferenca : diferenca = valor_apagar - (valor_parcela_ok*qtd_parcelas_x)

	dim ii 
	for ii = (1 + qtd_parcelas_pagas) to int(qtd_parcelas_x + qtd_parcelas_pagas)

		if (diferenca <> 0 and ii = (1 + qtd_parcelas_pagas)) then
		  valor_parcela = valor_parcela_ok + diferenca
		else
		  valor_parcela = valor_parcela_ok
		end If

		nosso_numero = COD_PEDIDO & preencheCaracter(ii,"0",2)

		'//data_vencimento(dia_vencimento, mes_vencimento, ano_vencimento, parcela)
		vencimento = data_vencimento(dia_vencimento, mes_vencimento, ano_vencimento, ii)
		hash = sha256(nosso_numero & hashkey)

		objConexao.Execute("INSERT INTO TB_PEDIDOS_BOLETOS (cod_pedido, cod_vendedor, cod_sacado, parcela, nosso_numero, hash, valor, data_vencimento) VALUES ("&_
		    sqlNumero(cod_pedido)&", "&_
		    sqlNumero(cod_vendedor)&", "&_
		    sqlNumero(cod_sacado)&", "&_
		    sqlNumero(ii)&", "&_
		    sqlTexto(nosso_numero)&", "&_
		    sqlTexto(hash)&", "&_
		    sqlNumero(valor_parcela)&", "&_
		    sqlData(vencimento)&""&_
		") ")

	next

	retorno = "ok"

	if not conexao_aberta then
		fecharConexao()
	end if

alteraParcelamento = retorno
end function

function excluirboletosSacados(cod_pedido, conexao_aberta)

	dim retorno, rs, cod_sacado

	if not conexao_aberta then
		abrirConexao()
	end if
	
	cod_sacado = ""
	
	set rs = objConexao.Execute("SELECT cod_sacado FROM TB_PEDIDOS_BOLETOS WHERE cod_pedido = "&sqlNumero(cod_pedido))
	if not rs.Eof then
		cod_sacado = rs(0)
	end if
	rs.close
	set rs = nothing

	if cod_sacado = "" then
		set rs = objConexao.Execute("SELECT cod_sacado FROM TB_PEDIDOS WHERE cod_pedido = "&sqlNumero(cod_pedido))
		if not rs.Eof then
			cod_sacado = rs(0)
		end if
		rs.close
		set rs = nothing
	end if

	'verificamos se temos boletos pagos, pois se tivermos então não apagamos o sacado nem as parcelas pagas
	set rs = objConexao.Execute("SELECT * FROM TB_PEDIDOS_BOLETOS WHERE data_vencimento is not null and cod_pedido = "&cod_pedido)
	if not rs.eof then 
		objConexao.execute("DELETE FROM TB_PEDIDOS_BOLETOS WHERE data_vencimento is null AND cod_pedido = "&sqlNumero(cod_pedido))

	else
		'Não encontramos nada
		objConexao.execute("DELETE FROM TB_SACADOS_ENDERECOS WHERE cod_sacado = "&sqlNumero(cod_sacado))
		objConexao.execute("DELETE FROM TB_SACADOS WHERE cod_sacado = "&sqlNumero(cod_sacado))
		objConexao.execute("DELETE FROM TB_PEDIDOS_BOLETOS WHERE cod_pedido = "&sqlNumero(cod_pedido))
	end if 
	rs.close
	set rs = nothing

	

	retorno = "Boletos Deletados com sucesso!"

	if not conexao_aberta then
		fecharConexao()
	end if

excluirboletosSacados = retorno
end function


function alterasacado(cod_sacado, cod_formando, nome_formando, cpf, email, cep, endereco, numero, bairro, cidade, estado, grau_parentesco, conexao_aberta)
'cod_formando
'nome_formando
'cpf
'email
'cep
'endereco
'numero
'bairro
'cidade
'estado
'grau_parentesco

dim def_cep, def_endereco, def_numero , def_bairro , def_cidade, def_estado , def_grau_parentesco, def_cpf, def_email, def_cod_formando, def_nome_formando

dim retorno, rs

if not conexao_aberta then
	abrirConexao()
end if

def_cpf = cpf
def_email = email

if cod_formando = "#outro" then
'usamos as informações enviadas 
	def_cod_formando = cod_formando
	def_nome_formando= nome_formando
	def_cep = cep
	def_endereco = endereco
	def_numero = numero
	def_bairro = bairro
	def_cidade = cidade 
	def_estado = estado
	def_grau_parentesco = grau_parentesco
else
'buscamos informações do formando'
	set rs = objConexao.Execute("SELECT COD_FORMANDO, NOME_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO = "&sqlNumero(cod_formando))
    if not rs.Eof then
      def_nome_formando = rs("NOME_FORMANDO")
    end if
    rs.Close
    set rs = Nothing

    set rs = objConexao.execute("SELECT * from TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO  WHERE  COD_FORMANDO = "&sqlTexto(cod_formando)&" AND DATA_CADASTRO = (SELECT MAX(DATA_CADASTRO) from TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO WHERE COD_FORMANDO = "&sqlTexto(cod_formando)&" ) ")
    if not rs.Eof Then
      def_grau_parentesco = ""
      def_cep = rs("cep")
	  def_endereco = rs("endereco")
	  def_numero = ""
	  def_complemento = ""
	  def_bairro = rs("bairro")
	  def_cidade = rs("cidade")
	  def_estado = rs("estado")
    end if 
    rs.Close
    Set rs = Nothing
end if




objConexao.Execute("UPDATE TB_SACADOS SET nome = "&sqlTexto(def_nome_formando)&", grau_parentesco = "&sqlTexto(def_grau_parentesco)&", email = "&sqlTexto(def_email)&", cpf = "&sqlTexto(def_cpf)&" Where cod_sacado = "&sqlNumero(cod_sacado))

objConexao.Execute("UPDATE TB_SACADOS_ENDERECOS SET "&_ 
	" endereco = "&sqlTexto(def_endereco)&_ 
	", numero = "&sqlTexto(def_numero)&_ 
	", complemento = "&sqlTexto(def_complemento)&_ 
	", bairro = "&sqlTexto(def_bairro)&_ 
	", cidade = "&sqlTexto(def_cidade)&_ 
	", estado = "&sqlTexto(def_estado)&_ 
	" WHERE cod_sacado = "&sqlNumero(cod_sacado))

if not conexao_aberta then
	fecharConexao()
end if

alterasacado = retorno
end function

%>