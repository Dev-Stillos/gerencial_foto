<%

'//////////////////////////////////////////////
'criptografia base64
Function MyASC(OneChar)
If OneChar = "" Then MyASC = 0 Else MyASC = Asc(OneChar)
End Function 

Function Base64Encode(inData)
  'rfc1521
  '2001 Antonin Foller, Motobit Software, http://Motobit.cz
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim cOut, sOut, I
  
  'For each group of 3 bytes
  For I = 1 To Len(inData) Step 3
    Dim nGroup, pOut, sGroup
    
    'Create one long from this 3 bytes.
    nGroup = &H10000 * Asc(Mid(inData, I, 1)) + _
      &H100 * MyASC(Mid(inData, I + 1, 1)) + MyASC(Mid(inData, I + 2, 1))
    
    'Oct splits the long To 8 groups with 3 bits
    nGroup = Oct(nGroup)
    
    'Add leading zeros
    nGroup = String(8 - Len(nGroup), "0") & nGroup
    
    'Convert To base64
    pOut = Mid(Base64, CLng("&o" & Mid(nGroup, 1, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 3, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 5, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 7, 2)) + 1, 1)
    
    'Add the part To OutPut string
    sOut = sOut + pOut
    
    'Add a new line For Each 76 chars In dest (76*3/4 = 57)
    'If (I + 2) Mod 57 = 0 Then sOut = sOut + vbCrLf
  Next
  Select Case Len(inData) Mod 3
    Case 1: '8 bit final
      sOut = Left(sOut, Len(sOut) - 2) + "=="
    Case 2: '16 bit final
      sOut = Left(sOut, Len(sOut) - 1) + "="
  End Select
  Base64Encode = sOut
End Function


Function geraAleatorio(menor,maior)

	Dim retorna

	Randomize	
	retorna = Int((maior-menor+1)*Rnd+menor)
	
	geraAleatorio = retorna
	
End Function


Function Base64Decode(ByVal base64String)
  'rfc1521
  '1999 Antonin Foller, Motobit Software, http://Motobit.cz
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim dataLength, sOut, groupBegin
  
  'remove white spaces, If any
  base64String = Replace(base64String, vbCrLf, "")
  base64String = Replace(base64String, vbTab, "")
  base64String = Replace(base64String, " ", "")
  
  'The source must consists from groups with Len of 4 chars
  dataLength = Len(base64String)
  If dataLength Mod 4 <> 0 Then
    Err.Raise 1, "Base64Decode", "Bad Base64 string."
    Exit Function
  End If

  
  ' Now decode each group:
  For groupBegin = 1 To dataLength Step 4
    Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
    ' Each data group encodes up To 3 actual bytes.
    numDataBytes = 3
    nGroup = 0

    For CharCounter = 0 To 3
      ' Convert each character into 6 bits of data, And add it To
      ' an integer For temporary storage.  If a character is a '=', there
      ' is one fewer data byte.  (There can only be a maximum of 2 '=' In
      ' the whole string.)

      thisChar = Mid(base64String, groupBegin + CharCounter, 1)

      If thisChar = "=" Then
        numDataBytes = numDataBytes - 1
        thisData = 0
      Else
        thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
      End If
      If thisData = -1 Then
        Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
        Exit Function
      End If

      nGroup = 64 * nGroup + thisData
    Next
    
    'Hex splits the long To 6 groups with 4 bits
    nGroup = Hex(nGroup)
    
    'Add leading zeros
    nGroup = String(6 - Len(nGroup), "0") & nGroup
    
    'Convert the 3 byte hex integer (6 chars) To 3 characters
    pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 5, 2)))
    
    'add numDataBytes characters To out string
    sOut = sOut & Left(pOut, numDataBytes)
  Next

  Base64Decode = sOut
End Function 

'FUNÇÕES DO MÓDULO DE COMPRAS
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Rotina criada para mudar o status do pedido de compra, retorna TRUE se foi realizado ou FALSE se não foi alterada
'Obs.: cancelado_motivo utilizado somente quando for para alterar para o status 99
Function pedidoCompraMudancaStatus(cod_pedido_compra,novo_status,cancelado_motivo,conexao_aberta)
	
	If not conexao_aberta Then
		AbrirConexao()
	End if	
	
	Dim rs, retorna, sql
		
		If int(novo_status) = 2 Then
		'Mudança de status para: Aprovado, mas não enviado para fornecedor
		
			'1) Verifica se já não foi aprovado antes
			Set rs=objConexao.Execute("SELECT cod_pedido_compra FROM TB_PEDIDOS_COMPRAS WHERE data_aprovado IS NULL AND cod_pedido_compra = '"&cod_pedido_compra&"'")
				
				If not rs.Eof Then
					objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS SET status_pedido_compra = "&novo_status&", data_aprovado = "&sqlData(now)&", aprovado_cod_usuario = "&Session("cod_usuario")&", modificado = "&sqlData(now)&", cod_usuario = "&Session("cod_usuario")&" WHERE cod_pedido_compra = '"&cod_pedido_compra&"'")
					retorna = True
				Else
					retorna = false
				End If
			
			rs.close()
			Set rs=Nothing
		
		Elseif int(novo_status) = 3 Then
		
			'1) Verifica se já não foi enviado para o fornecedor antes
			sql = "SELECT cod_pedido_compra FROM TB_PEDIDOS_COMPRAS WHERE data_enviado_fornecedor IS NULL AND cod_pedido_compra = '"&cod_pedido_compra&"'"
			
			
			Set rs=objConexao.Execute(sql)
	

				If not rs.Eof Then
					
					'1/1---------------------------------------------------------------------------------------
					'Atualiza, pois a data é nula para este pedido de compra
						objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS SET status_pedido_compra = "&novo_status&", data_enviado_fornecedor = "&sqlData(now)&", enviado_cod_usuario = "&Session("cod_usuario")&", modificado = "&sqlData(now)&", cod_usuario = "&Session("cod_usuario")&" WHERE cod_pedido_compra = '"&cod_pedido_compra&"'")

						objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS_PRODUTOS SET status_produto = 'E', quantidade_entregue = quantidade_solicitada, modificado = "&sqlData(now)&", cod_usuario = "&Session("cod_usuario")&" WHERE cod_pedido_compra = '"&cod_pedido_compra&"'")


						retorna = true

				Else
					retorna = false
				End If
			
			rs.close()
			Set rs=Nothing
			
			
		Elseif int(novo_status) = 22 Then	
		
				
			'1/1---------------------------------------------------------------------------------------
			'Volta o pedido com o status não enviado para o fornecedor
			'objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS SET status_pedido_compra = 2, data_enviado_fornecedor = NULL, enviado_cod_usuario = NULL, modificado = "&sqlData(now)&", cod_usuario = "&Session("cod_usuario")&" WHERE cod_pedido_compra = '"&cod_pedido_compra&"'")
			'retorna = true
				
			
		Elseif int(novo_status) = 99 Then
	
			'1) Verifica se já não foi cancelado anteriormente
			Set rs=objConexao.Execute("SELECT cod_pedido_compra FROM TB_PEDIDOS_COMPRAS WHERE status_pedido_compra <> 99 AND cod_pedido_compra = '"&cod_pedido_compra&"'")
				
				If not rs.Eof Then
				
					'1/2---------------------------------------------------------------------------------------
					'Atualiza, pois a data é nula para este pedido de compra
					objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS SET cancelado_motivo = "&sqlTexto(cancelado_motivo)&", status_pedido_compra = 99, modificado = "&sqlData(now)&", cod_usuario = "&Session("cod_usuario")&" WHERE cod_pedido_compra = '"&cod_pedido_compra&"'")
						
					'2/2---------------------------------------------------------------------------------------

					'Cancela todos os itens como cancelados
					objConexao.Execute("UPDATE TB_PEDIDOS_COMPRAS_PRODUTOS SET status_produto = 'C', modificado = "&sqlData(now)&", cod_usuario = "&Session("cod_usuario")&" WHERE cod_pedido_compra = '"&cod_pedido_compra&"'")	
							
					retorna = True
				Else
					retorna = false
				End If
			
			rs.close()
			Set rs=Nothing	
		
		End If
		
	If not conexao_aberta Then
		FecharConexao
	End if	
	
	pedidoCompraMudancaStatus = retorna
	
End Function


'Função que retorna o total de acordo com o tipo solicitado, informações sobre o significado de cada tipo na própria função
Function pedidoCompraTotais(cod_pedido_compra,tipo,conexao_aberta)

	If not conexao_aberta Then
		AbrirConexao(cod_base)
	End if	

	Dim rs, sql, casas_decimais
	
	If tipo = 1 Then
	'1 = Valor Total em Produtos Solicitados
		sql = "SELECT IFNULL(SUM(quantidade_solicitada*valor_custo_solicitado),0) as total FROM TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra = "&cod_pedido_compra
		casas_decimais = 2
	ElseIf tipo = 2 Then
	'2 = Valor Total em Produtos Entregues
		sql = "SELECT IFNULL(SUM(quantidade_entregue*valor_custo_entregue),0) as total FROM TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra = "&cod_pedido_compra
		casas_decimais = 2
	ElseIf tipo = 3 Then
	'3 = Qtd. Itens Solicitados (mesmo cancelados)
		sql = "SELECT IFNULL(COUNT(quantidade_solicitada),0) as total FROM TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra = "&cod_pedido_compra
		casas_decimais = 0
	ElseIf tipo = 4 Then
	'4 = Qtd. de Itens Entregues é a soma total 
		sql = "SELECT IFNULL(SUM(quantidade_solicitada),0) as total FROM TB_PEDIDOS_COMPRAS_PRODUTOS WHERE status_produto = 'E' and cod_pedido_compra = "&cod_pedido_compra
		casas_decimais = 0
	ElseIf tipo = 5 Then
	'5 = Qtd. de Itens Pendentes // Não é a soma total e sim a qtd. diferente de itens
		sql = "SELECT IFNULL(COUNT(quantidade_solicitada),0) as total FROM TB_PEDIDOS_COMPRAS_PRODUTOS WHERE status_produto = 'P' and cod_pedido_compra = "&cod_pedido_compra
		casas_decimais = 0
	ElseIf tipo = 6 Then
	'6 = Qtd. de Itens Cancelados // Não é a soma total e sim a qtd. diferente de itens
		sql = "SELECT IFNULL(COUNT(quantidade_solicitada),0) as total FROM TB_PEDIDOS_COMPRAS_PRODUTOS WHERE status_produto = 'C' and cod_pedido_compra = "&cod_pedido_compra
		casas_decimais = 0
	Elseif tipo = 7 Then
	'7 = Qtd. de Itens que estão atrasados
		sql = "SELECT IFNULL(COUNT(*),0) as total FROM TB_PEDIDOS_COMPRAS_PRODUTOS WHERE cod_pedido_compra = "&cod_pedido_compra&" and DATEDIFF(NOW(),data_previsao_entrega) > 0 and status_produto = 'P'"
		casas_decimais = 0
	End If

			Set rs=objConexao.Execute(sql)
				
				If not rs.Eof Then
					retorna = formatNumber(0&rs(0),casas_decimais)
				End If
			
			rs.close()
			Set rs=Nothing


	If not conexao_aberta Then
		FecharConexao(cod_base)
	End if	
	
	pedidoCompraTotais = retorna
	
End Function

Function pedidoCompraStatus(status_pedido_compra)
	
	Dim temp, retorna
	temp = formatNumber(0&status_pedido_compra,0)
	
		If temp = 0 Then
			retorna = "Em cadastro"
		Elseif temp = 1 Then
			retorna = "Aguardando aprova&ccedil;&atildeo"
		Elseif temp = 2 Then
			retorna = "Pendente"
		Elseif temp = 3 Then
			retorna = "Finalizado e entregue totalmente"
		Elseif temp = 99 Then
			retorna = "Cancelado"
		End If
	
	pedidoCompraStatus = retorna
	
End Function

'////////////////////////
'COMPRAS

'Função que retorna o total de acordo com o tipo solicitado, informações sobre o significado de cada tipo na própria função
Function compraTotais(cod_compra,tipo,conexao_aberta)

	If not conexao_aberta Then
		AbrirConexao()
	End if	

	Dim rs, sql, casas_decimais
	
	If tipo = 1 Then
	'1 = Valor Total em Produtos
		sql = "SELECT IFNULL(valor_total_produtos,0) as total FROM VW_TB_Compras WHERE cod_compra = "&cod_compra
		casas_decimais = 2
	ElseIf tipo = 2 Then
	'2 = Qtd. Itens
		sql = "SELECT IFNULL(COUNT(quantidade),0) as total FROM TB_ComprasProdutos WHERE cod_compra = "&cod_compra
		casas_decimais = 0
	Elseif tipo = 3 THen
	'3 - Valor Total Parcelas
		sql = "SELECT IFNULL(SUM(valor),'0') as total FROM TB_ComprasParcelas WHERE cod_compra = "&cod_compra
		casas_decimais = 2
	Elseif tipo = 4 THen
	'3 - Valor Total Parcelas
		sql = "SELECT IFNULL(SUM(valor_custo_total),'0') as total FROM VW_TB_ComprasProdutos WHERE cod_compra = "&cod_compra
		casas_decimais = 4
	End If
	

			Set rs=objConexao.Execute(sql)
				
				If not rs.Eof Then
					retorna = formatNumber(0&rs(0),casas_decimais)
				End If
			
			rs.close()
			Set rs=Nothing


	If not conexao_aberta Then
		FecharConexao()
	End if	
	
	compraTotais = retorna
	
End Function


'Fuction PadraoFormularioSelect
'	Responsalvel por criar selects simples(sem optiongroup)  
'PARAMETROS
'	nome			= propriedade NAME
'	id				= propriedade ID
'	classe			= propriedade CLASS
'	style			= propriedade STYLE
'	adicional		= parametro livre para iserção de propriedades como ONBLUR, ONCLICK....
'	optionPadrao	= parametro que define o nome do 1º option a ser listado Ex.: 'Selecione'
'	chaveSelecionada= paravetro que define se existe um campo option pré selecionado se vazio o campo padrão será o optionPadrao
'	query			= string SQL completa que gerara o select Ex.: SELECT cod_teste, descricao FROM TB_Teste WHERE cod_teste IN '2,3,45,6'
'	comando 		= comando da página
'	criptografar	= informa se a função de criptografia está habilitada (true/false) não está habilitada ainda
'	debug			= verifica se o modo debug esta ativado caso sim exibe response.Write adicionais.
'	conexao_aberta	= informa o estado da conexao true = aberta e false = fechada
'O CODIGO PODE SER USADO COM UM PADRAO NESSE CASO DEVEMOS USAR A SEGUINTE QUERY "PADRAO:valo1#texto1|valo2#texto2|valo3#texto3"
Function PadraoFormularioSelect(nome,id,classe,style,adicional,optionPadrao,chaveSelecionada,query,comando,criptografar,debug,conexao_aberta)
	Dim rs, retorno, temp, temp2, i
	'Verificamos se a conexao está aberta se não estiver abrimos
	If not conexao_aberta AND not  mid(query,1,6) = "PADRAO" Then
		AbrirConexao()

	End if
	
	If debug and session("cod_usuario")then
		response.Write("<pre>NOME CAMPO:" & trim(nome)& vbnewline & "CONSULTA:"&trim(query) & vbnewline& "COMANDO: "&trim(comando) &vbnewline & "CONEXAO ABERTA:"& trim(conexao_aberta))
	end if 
		
	'verificamos se todos os parametro principais foram preenchidos 
	If trim(nome) <> "" and trim(query) <> "" and trim(comando) <> "" and trim(conexao_aberta) <> "" Then
		If comando = "incluir" or comando = "alterar" Then
			retorno = retorno & "<select name='"&ucase(trim(nome))&"' "
			'veirificamos abaixo qual parametros estão sendo usados
			If trim(id) <> "" Then
				retorno = retorno & "id='"&trim(id)&"' "				
			End if
			If trim(classe) <> "" Then
				retorno = retorno & "class='"&trim(classe)&"' "
			End if
			If trim(style) <> "" Then
				retorno = retorno & "style='"&trim(style)&"' "
			End if
			If trim(adicional) <> "" Then
				retorno = retorno &" "&trim(adicional)&" "
			End if
			retorno = retorno & ">"						
			
			'Padroniza a chave selecionada e faz tratamento
			temp = Trim(Cstr(""&chaveSelecionada))
			
			'Verificamos o nome do 1º option casos o parametro optionPadrao seja vazio entao o padrao é 'Selecione'
			If optionPadrao <> "SEM_SELECIONE" Then
				If optionPadrao <> "" and temp = "" Then			
						retorno = retorno & "<option value='' selected='selected'>"&trim(optionPadrao)&"</option>"
				Elseif optionPadrao <> "" and temp <> "" Then
					retorno = retorno & "<option value=''>"&trim(optionPadrao)&"</option>"
				ElseIf optionPadrao = "" and temp = "" Then
					retorno = retorno & "<option value='' selected='selected'>Selecione</option>"
				ElseIf optionPadrao = "" and temp <> "" Then
					retorno = retorno & "<option value='' >Selecione</option>"
				End if
			End if		
			'Response.Write(query)
			If mid(query,1,6) = "PADRAO" Then
				i = 0
				temp2 = mid(query,8,len(query)-7)
				temp2 = split(temp2,"|")
				
				While i <= ubound(temp2)			
					retorno = retorno & "<option value='"&split(temp2(i),"#")(0)&"'"
					If Cstr(""&trim(split(temp2(i),"#")(0)&"")) = Cstr(""&trim(chaveSelecionada&"")) then
						retorno = retorno & " selected='selected'"
					End if
					retorno = retorno & ">"&split(temp2(i),"#")(1)&"</option>"
				i=i+1
				WEnd
			Else
                Response.Write(query)
				Set rs = ObjConexao.Execute(query)
					While not rs.Eof
						retorno = retorno & "<option value='"&rs(0)&"'"
						If trim(rs(0)&"") = trim(chaveSelecionada&"") then
							retorno = retorno & " selected='selected'"
						End if
						retorno = retorno & ">"&rs(1)&"</option>"
					rs.movenext()
					WEnd
				rs.close()
				Set rs = nothing
			End if
			retorno = retorno & "</select>"
			
		Elseif comando = "visualizar" or comando = "excluir" Then
		'Response.Write(query)
			If mid(query,1,6) = "PADRAO" Then
				i = 0
				temp2 = mid(query,8,len(query)-7)
				temp2 = split(temp2,"|")

				While i <= ubound(temp2)
					If trim(split(temp2(i),"#")(0)&"") = trim(chaveSelecionada&"") then
						retorno = retorno & ""&split(temp2(i),"#")(1)&"&nbsp;"
					End if
				i=i+1
				WEnd
				
			Else
				Set rs = ObjConexao.Execute(query)
					While not rs.Eof
						If trim(rs(0)&"") = trim(chaveSelecionada&"") then
							retorno = retorno & ""&rs(1)&"&nbsp;"
						End if
					rs.movenext()
					WEnd
				rs.close()
				Set rs = nothing
			End if
		Else
			retorno = retorno & "comando inválido"
		End if
	Else
		retorno = retorno & "Preencha os campos da função corretamente"
	End if
	
	If not conexao_aberta AND not  mid(query,1,6) = "PADRAO" Then
		FecharConexao()
	End if

	PadraoFormularioSelect = retorno
End Function


'GAMB PARA NAO JOGAR ESTE SCRIPT EM PAGINAS DE AJAX
if cstr(script_valida_form_master&"") = "" then
response.Write script_valida_form_master
%>
<script type="text/javascript">
function salvar(){
	if(validaFormMaster()){
		document.FORMULARIO_PADRAO.submit();
	}
}
</script>
<%end if%>