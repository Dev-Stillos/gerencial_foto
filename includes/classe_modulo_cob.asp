<%
'//////////////////////////////////////////////////////////////////////////
'ROTINA E FUNÇÕES PARA O MÓDULO COB - COBRANÇA


function baixaBoleto(cod_boleto_, valor_pago_, data_pagamento_)

	dim rs, retorno
	dim cod_boleto, valor_pago, data_pagamento

baixaBoleto = retorno
end function



'Chamada para  a página de impressao de boletos
'retorna caminho dos boleto gerados
function geraBoleto(cod, tipo, formato, conexao_aberta)

	dim retorno, sql,  i

	arquivos = ""

	if not conexao_aberta then
		AbrirConexao()
	end if

	if tipo = "boletos" then
		sql = "SELECT cod_pedido, cod_boleto, parcela, data_vencimento, valor, valor_pago, hash, data_pagamento FROM TB_PEDIDOS_BOLETOS WHERE cod_boleto in ("&cod&")"
	elseif tipo = "cod_pedido" then
		sql = "SELECT cod_pedido, cod_boleto, parcela, data_vencimento, valor, valor_pago, hash, data_pagamento FROM TB_PEDIDOS_BOLETOS WHERE cod_pedido = "&sqlNumero(cod)
	elseif tipo = "cod_sacado" then
		sql = "SELECT cod_pedido, cod_boleto, parcela, data_vencimento, valor, valor_pago, hash, data_pagamento FROM TB_PEDIDOS_BOLETOS WHERE cod_sacado = "&sqlNumero(cod)
	end if

	i = 0
	set rs = objConexao.Execute(sql)
	if not rs.Eof then
		while not rs.Eof

			'geramos o nome do arquivo
			nome_arquivo = "boleto_"&rs("cod_pedido")&"_"&preencheCaracter(rs("parcela"),"0",2)&".pdf"

			'call SaveFileFromUrl("https://stillos.com.br/gerencial_foto/gerarBoletoPdf.php?h="&rs("hash"), caminho_dados&"boletos\" & nome_arquivo)
			
			Set Pdf = Server.CreateObject("Persits.Pdf")
			Set Doc = Pdf.CreateDocument
			Doc.ImportFromUrl "https://stillos.com.br/gerencial_foto/xdwfsfdgn.asp?h="&rs("hash"), "scale=0.6; hyperlinks=true; drawbackground=true"
			Filename = Doc.Save( nome_arquivo, False )


			arquivos = arquivos & caminho_dados&"boletos\" & nome_arquivo & ";"
			i = i + 1

		rs.moveNext()
		wend
	end if 
	rs.Close
	set rs = Nothing

	retorno = mid(arquivos,1, len(arquivos) - 1)
	retorno = split(retorno, ";")

	if not conexao_aberta then
		FecharConexao()
	end if
geraBoleto = retorno
end function


function SaveFileFromUrl(Url, FileName)

	dim objXMLHTTP, objADOStream, objFSO

	Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP.3.0")
	'objXMLHTTP.setTimeouts 5000, 5000, 5000, 5000
	objXMLHTTP.open "GET", Url, false
	
	objXMLHTTP.send()
	

	If objXMLHTTP.Status = 200 Then 
		Set objADOStream = CreateObject("ADODB.Stream")
		objADOStream.Open
		objADOStream.Type = 1 'adTypeBinary

		objADOStream.Write objXMLHTTP.ResponseBody
		objADOStream.Position = 0 'Set the stream position to the start

		Set objFSO = Createobject("Scripting.FileSystemObject")
		If objFSO.Fileexists(FileName) Then objFSO.DeleteFile FileName
		Set objFSO = Nothing

		objADOStream.SaveToFile FileName
		objADOStream.Close
		Set objADOStream = Nothing

		SaveFileFromUrl = objXMLHTTP.getResponseHeader("Content-Type")
	else
		SaveFileFromUrl = ""
	End if

    Set objXMLHTTP = Nothing
end function



function enviaBoletos(cod_sacado, somente_avencer, conexao_aberta)

	dim rs, retorno(1), anexos
	retorno(0) = 0
	retorno(1) = ""

	if not conexao_aberta then
		AbrirConexao()
	end if

	if cod_sacado = "" then
		retorno(0) = -1
		retorno(1) = "Parametros inválidos"
	end if


	dim mensagem_email, assunto_email, destinatrio_email, hash

	assunto_email = "Stillo's Reportagens Fotográficas - Boletos"

	mensagem_email = "Prezado(a) [NOME], <br />"
	mensagem_email = mensagem_email & "" & "<br />"
	mensagem_email = mensagem_email & "Segue anexo os boletos referentes a sua compra." & " <br />"
	'mensagem_email = mensagem_email & "1) Acesse o link: <a href='http://stillos.com.br/boletosfoto/primeiro_acesso.asp?h=[HASH]' target='_blank'>http://stillos.com.br/boletosfoto/primeiro_acesso.asp?h=[HASH]</a> " & " ;<br />"
	'mensagem_email = mensagem_email & "2) Preencha o campo CPF com o CPF informando ao consultor;" & " <br />"
	'mensagem_email = mensagem_email & "3) Crie uma senha;" & " <br />"
	'mensagem_email = mensagem_email & "4) Acesse o painel com o e-mail informando e senha criada." & " <br /><br />"

	'mensagem_email = mensagem_email & "Caso tenha dificuldades no acesso entre em contato, pelo telefone: (11) 2672-9898 de Segunda a Sexta das 8:30 às 18:30 <br /><br />"

	mensagem_email = mensagem_email& "*****************************************"
	mensagem_email = mensagem_email& "ENVIO AUTOMÁTICO, NÃO RESPONDA ESTE E-MAIL"
	mensagem_email = mensagem_email& "*****************************************"


	if retorno(0) >= 0 then

		'buscamos o sacado
		set rs = objConexao.Execute("SELECT * FROM TB_SACADOS WHERE cod_sacado = "&sqlNumero(cod_sacado))
		if not rs.Eof then
			'mensagem_email = replace(mensagem_email, "[HASH]", ""&rs("hash"))
			mensagem_email = replace(mensagem_email, "[NOME]", ""&rs("nome"))
			destinatrio_email = "jpbeltrame@gmail.com"'rs("email")
		end if
		rs.Close
		set rs = Nothing
		anexos = geraBoleto(cod_sacado, "cod_sacado", "", true)

		a = enviaEmail(destinatrio_email,  assunto_email, mensagem_email,  anexos)
		retorno(0) = a(0) 
		retorno(1) = a(1) 
	end if

	if not conexao_aberta then
		FecharConexao()
	end if

enviaBoletos = retorno
end function


function enviaEmail(destinatrio,  assunto, mensagem, arrAnexos)

	on error resume next

	dim servidor_, usuario_, senha_, remetente_, porta_

	servidor_  = "smtplw.com.br"
	usuario_   = "stillosadm2"
	senha_     = "xwpdkwem4525"
	remetente_ = """Stillo's Reportagens Fotográficas"" <nao-responda@stillos.com.br>"
	porta_     = "587"

	set objMessage = createobject("cdo.message")
	set objConfig = createobject("cdo.configuration")
	Set Flds = objConfig.Fields

	Flds.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	Flds.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = servidor_
	Flds.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = porta_
	Flds.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = false
	Flds.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic (clear-text) authentication 
	Flds.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = usuario_
	Flds.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = senha_ 
	Flds.update

	Set objMessage.Configuration = objConfig 
	objMessage.To = destinatrio
	objMessage.From = remetente_
	objMessage.Subject = assunto
	objMessage.fields.update 
	objMessage.HTMLBody = mensagem
	
	If IsArray(arrAnexos) Then
		For I = 0 To UBound(arrAnexos)
			objMessage.AddAttachment arrAnexos(i)
		Next
	End If

	objMessage.Send

	set Flds       = Nothing
	set objMessage = Nothing
	set objConfig  = Nothing

	dim erro(1)
	erro(0) = Err.number	
	erro(1) = Err.description 

enviaEmail = erro 
end function

function nomeArquivo(caminho_completo)
dim retorno, tmp
	
	tmp = replace(tmp, "/", "\")
	tmp = split(tmp, "\")
	retorno = tmp(UBound(tmp))

nomeArquivo = retorno
end function


function parcelasBoleto(cod_pedido, tipo_retorno, conexao_aberta)
'valor Parcela 

  Dim config_valor_parcela_min : config_valor_parcela_min = 50
  Dim config_max_parcelas : config_max_parcelas = 15

  
  Dim XTOTAL_AGREGADOS, XTOTAL_ALBUNS, XTOTAL_PRODUTOS_A_PARTE
  Dim rs, retorno 
  Dim valor_pedido, valor_lancado, qtd_parcelas_pedido
  Dim i

  if not conexao_aberta then
    AbrirConexao()
  end if

  'buscamos as informações do pedido
  Set rs = objConexao.Execute("SELECT QTD_PARCELAS FROM TB_PEDIDOS WHERE COD_PEDIDO = "&COD_PEDIDO&" LIMIT 1")
    If not rs.Eof Then
      qtd_parcelas_pedido = rs(0)
    End IF              
  rs.Close()
  Set rs = Nothing

  'buscamos as informações do pedido
  Set rs = objConexao.Execute("SELECT * FROM VW_PEDIDOS_TOTAL_RECEBIDO WHERE COD_PEDIDO = "&COD_PEDIDO&" LIMIT 1")
    If not rs.Eof Then
      valor_lancado = CDBL("0"&rs("PAGAMENTO_EM_DINHEIRO")) + CDBL("0"&rs("XTOTAL_CARTAO")) + CDBL("0"&rs("XTOTAL_CHEQUE"))
    End IF              
  rs.Close()
  Set rs = Nothing

  '---------------------------------------------------------------'

  Set rs = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_AGREGADOS WHERE COD_PEDIDO="&COD_PEDIDO)
  if not rs.Eof Then
    XTOTAL_AGREGADOS = cdbl(0&rs("XTOTAL_AGREGADOS"))
  End If
  rs.Close()
  Set rs = Nothing

  Set rs = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_ALBUNS WHERE COD_PEDIDO="&COD_PEDIDO)
  if not rs.Eof Then
    XTOTAL_ALBUNS = cdbl(0&rs("XTOTAL_ALBUNS"))
  End If
  rs.Close()
  Set rs = Nothing

  Set rs = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_PRODUTOS WHERE COD_PEDIDO="&COD_PEDIDO)
  if not rs.Eof Then
    XTOTAL_PRODUTOS_A_PARTE  = cdbl(0&rs("XTOTAL_PRODUTOS_A_PARTE"))
  End If
  rs.Close()
  Set rs = Nothing

  'BUSCAMOS DESCONTOS'
  Set rsProdutosTotal = objConexao.Execute("select SUM(valor_total_com_desconto) as valor_total_com_desconto from(SELECT SUM(faturado * (valor_un-(valor_un * (1-(ifnull(aliquota_desconto_foto,0)/100))))) as valor_total_com_desconto FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO in (SELECT cod_formando FROM TB_PEDIDOS_FORMANDOS WHERE cod_pedido =  "&COD_PEDIDO&") union all SELECT SUM(quantidade * (valor_un-(valor_un * (1-(ifnull(aliquota,0)/100))))) as valor_total_com_desconto  FROM TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO = "&cod_pedido&") as total_desconto")
  if not rsProdutosTotal.Eof Then
    XTOTAL_DESCONTO = formatNumber(0&rsProdutosTotal(0))/1
  End If
  rsProdutosTotal.Close()
  Set rsProdutosTotal = Nothing


  valor_pedido = XTOTAL_PRODUTOS_A_PARTE + XTOTAL_ALBUNS + XTOTAL_AGREGADOS - XTOTAL_DESCONTO
  valor_restante  = valor_pedido - valor_lancado

  if valor_restante < 0 then
    valor_restante = 0
  end if
  
  if qtd_parcelas_pedido = 1 then  
    config_max_parcelas = 1   
  end if

  dim qtd_parcelas_pagas,valor_pago


  if tipo_retorno = "html" then
      
    if valor_restante < 50 then
      retorno = "Atenção, o valor minímo de parcela no boleto é R$ 50,00."
    else

      for i = 1 to config_max_parcelas 
        '
        valor_parcela = valor_restante/i
      
        if (valor_parcela >= 50) then
          retorno = retorno & "<option value='"&i&"'>"&i&"x "&formatCurrency(valor_parcela)&"</option>"
        else
          Exit For
        end if 

      next

      retorno = "<select name='qtd_parcelas' class='validacao_etapa_1 vencimento' nome-campo='Quant. Parcelas' required='required'><option value=''>Selecione</option>" & retorno & "</select>"
    end if
  elseif tipo_retorno = "valor_apagar" then
    retorno = valor_restante
  end if 





  if not conexao_aberta then
    FecharConexao()
  end if 

  parcelasBoleto = retorno

end function


function valorMultaJuros(data_vencimento_ , data_pagamento_ , valor_pago_ , valor_)

	dim rs, retorno, i, meses_atraso, valor, dias, hoje, data_vencimento
	dim multa : multa = 1.02 '2%
	dim juros : juros = 1.01 '1% Ao mes

	hoje = now()
	valor = cdbl(valor_)

	'verificamos de já foi pago
	if not isnull(data_pagamento_) then
		retorno =cdbl(valor_pago_)
		exit function 'finalizamos a função
	end if

	'verificamos a data de vencimento
	if isnull(data_vencimento_) then
		data_vencimento = now()
	else
		data_vencimento = data_vencimento_
	end if

	'verificamos se o vencimento caiu em um sabado ou domingo
	Select Case weekdaY(data_vencimento)
  		Case 1
			dias_extra = 1
  		Case 7
  			dias_extra = 2
  		case else
  			dias_extra = 0
  	end Select

  	'Adicionamos a variavel dias extras para não cosiderarmos juros
	diferenca = datediff("d",  dateadd("d",dias_extra,data_vencimento), hoje)

	'diferenca > 0 teve vencimento
	if diferenca > 0 then
		'multa
		retorno = valor * multa

		'juros
		meses_atraso = int(diferenca / 30)
		juros = juros ^ meses_atraso
		retorno = retorno * juros
	else
		retorno = cdbl(valor_)
	end if 


valorMultaJuros = cdbl(formatnumber(retorno,2))
end function

function vencimentoCorrigido(data_vencimento_)

	dim data_vencimento, retorno, diferenca, agora, dias_extra, hoje

	data_vencimento = data_vencimento_
	'hoje = mid(now(), 1, 10)

	Select Case weekdaY(data_vencimento)
  		Case 1
			dias_extra = 1
  		Case 7
  			dias_extra = 2
  		case else
  			dias_extra = 0
  	end Select
  	
  	'diferenca = datediff("d", dateadd("d",dias_extra,data_vencimento), hoje)

  	'if diferenca < 0 then
  		retorno = data_vencimento
  	'else
  	'	retorno = hoje
  	'end if

vencimentoCorrigido = retorno
end function



'copiar para o conexao.asp
function data_vencimento(dia_vencimento, mes_vencimento, ano_vencimento, parcela)
  dim retorno

  retorno = cdate(ano_vencimento&"-"&mes_vencimento&"-01")
  retorno = dateadd("m",parcela - 1,retorno)

  'ja temos o mes e ano de vencimento 
  if dia_vencimento = 30 and datepart("m", retorno) = 2 then
    retorno = dateadd("d", max_dias_mes(2, datepart("yyyy", retorno)) - 1, retorno) 
  else
    retorno = dateadd("d", dia_vencimento-1, retorno)
  end if

data_vencimento = retorno
end function

function anobissexto(ano)
    '//Se for divisível por 400, então É bissexto; 
    '//Senão, se for divisível por 100, então NÃO É bissexto; 
    '//Senão, se for divisível por 4, então É bissexto; 
    '//Senão, não é bissexto
    ano_ = int(ano) + 0

    retorno = ""

    if ( ano > 0 ) then
        if ( ano mod 400 = 0 ) then
            retorno = true
        else
            if ( ano mod 100 = 0) then
                retorno = false
            else
                if ( ano mod 4 = 0 ) then
                    retorno = true
                else
                    retorno = false
                end if
            end if
        end if
    else
        retorno =  false
    end if

    anobissexto = retorno

end function

function max_dias_mes(mes, ano) 

    dim retorno

    if (ano = 0) then
        ano_ = datepart("yyyy",now())
    else
        ano_ = ano
    end if


    if (mes = 0) then
        mes_ = datepart("m",now())
    else
        mes_ = mes
    end if

    Select  case mes_
        case 1
            retorno = 31
         
        case 2
            if anobissexto(ano_) then
                retorno = 29
            else
                retorno = 28
            end if

        case 3
          retorno = 31

        case 4
          retorno = 30

        case 5
          retorno = 31
          
        case 6
          retorno = 30
        
        case 7
          retorno = 31

        case 8
          retorno = 31
        
        case 9
          retorno = 30
        
        case 10
          retorno = 31
        
        case 11
          retorno = 30

        case 12
          retorno = 31

        case else
          retorno = 0
    end select

max_dias_mes = retorno
end function


%>