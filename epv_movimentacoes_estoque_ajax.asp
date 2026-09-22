<% Response.Charset="ISO-8859-1" %>
<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim tipo_solicitacao, retorna
Dim parametros(11)
tipo_solicitacao = Cstr(""&Request.QueryString("ts"))

For i=1 to 10
	parametros(i) = Request.QueryString("p"&i)
Next 

AbrirConexao()

Dim VALIDA_TIPO, VALIDA_COD_FORMANDO

If tipo_solicitacao = "incluir" Then
  'parametros(i) = TXT_DEVOLUCAO
	TXT_DEVOLUCAO = parametros(1)
	TXT_DEVOLUCAO = Replace(TXT_DEVOLUCAO," ","")
	TXT_DEVOLUCAO = Replace(TXT_DEVOLUCAO,".","")
	TXT_DEVOLUCAO = Replace(TXT_DEVOLUCAO,"-","")

	VALIDA_TIPO	= Right(TXT_DEVOLUCAO,1)
	VALIDA_COD_FORMANDO	= Mid(TXT_DEVOLUCAO,1,Len(TXT_DEVOLUCAO)-1)
  
  set config = getConfiguracoes(true, array("periodo_renovacao_padrao"))				
	
  sql = "SELECT * FROM VW_MOVIMENTACOES_COMPLETA WHERE COD_FORMANDO = '"&VALIDA_COD_FORMANDO&"' AND TIPO = '"&VALIDA_TIPO&"'"

	Set rsVerificaDig = objConexao.Execute(sql)
	If not rsVerificaDig.Eof Then
 
 	  'Verifica se j� n�o foi inserido na tabela tempor�ria
  	Set rs2=objConexao.Execute("SELECT COD_ETIQUETA FROM TB_FICHAS_IDENTIFICACAO_DEVOLUCAO_TEMPORARIA WHERE COD_USUARIO = "&Session("COD_USUARIO")&" AND COD_ETIQUETA = "&sqlTexto(TXT_DEVOLUCAO))
		If rs2.Eof Then 'Pode gravar, pois n�o foi inserido ainda	
			sql = "INSERT INTO TB_FICHAS_IDENTIFICACAO_DEVOLUCAO_TEMPORARIA (COD_USUARIO,COD_ETIQUETA,PRAZO,PRAZO_NECESSARIO, COD_FORMANDO) VALUES ('"&Session("COD_USUARIO")&"', "&TXT_DEVOLUCAO&", "&config("periodo_renovacao_padrao")&", '', '"&VALIDA_COD_FORMANDO&"')"

			objConexao.Execute(sql)
     
			temp = Cstr(""&rsVerificaDig("NOME_FORMANDO"))
			temp = Replace(temp,"'","")
			temp = Replace(temp,"""","")
				
			retorna = "<!--#true#carregaListagem();displayLeitura('sucesso','Item do formando inserido com sucesso: "&temp&"');#-->"				
		Else
			retorna = "<!--#true#displayLeitura('erro','Item já inserido: "&TXT_DEVOLUCAO&"');#-->"
		End If

    rs2.Close
		Set rs2=Nothing
	Else
		'N�o grava n�mero inv�lido
	
		retorna = "<!--#true#displayLeitura('erro','Número de etiqueta: "&TXT_DEVOLUCAO&" inválida');#-->"		
	End if	
		
  rsVerificaDig.Close
  Set rsVerificaDig = Nothing

Elseif tipo_solicitacao  = "listar" Then
	Set rs = objConexao.Execute("SELECT COD_ETIQUETA FROM TB_FICHAS_IDENTIFICACAO_DEVOLUCAO_TEMPORARIA WHERE COD_USUARIO="&Session("COD_USUARIO")&" GROUP BY COD_ETIQUETA ORDER BY DATA_CADASTRO DESC",qtdRegistros)
	If not rs.Eof Then

		TIPO = Right(rs("COD_ETIQUETA"),1)				
		COD_FORMANDO = Mid(rs("COD_ETIQUETA"),1,Len(rs("COD_ETIQUETA"))-1)
		'Seleciona a �ltima devolu��o deste c�d_formando
			
		sql = "SELECT TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.TIPO, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_SAIDA, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_RETORNO FROM TB_VENDEDORES INNER JOIN TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ON TB_VENDEDORES.COD_VENDEDOR = TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR WHERE COD_FORMANDO = '"&COD_FORMANDO&"' AND TIPO = '"&TIPO&"' ORDER BY DATA_SAIDA DESC LIMIT 1"

		Set rsVendedorVerifica = objConexao.Execute(sql)
	  If not rsVendedorVerifica.Eof Then
		  Dim NOME_VENDEDOR, COD_VENDEDOR
			NOME_VENDEDOR 	= rsVendedorVerifica("NOME_VENDEDOR")
			COD_VENDEDOR 	= rsVendedorVerifica("COD_VENDEDOR")
		End if	
		
		rsVendedorVerifica.Close
    Set rsVendedorVerifica = Nothing
		COD_FORMANDO = ""
    
    retorna = ""
    retorna = retorna & "<h3 align='center'>Devolução de Produtos do vendedor(a):<br />" & NOME_VENDEDOR & "</h3>"
    retorna = retorna & "<table width='95%' border='1' align='center' 	cellpadding='2' style='border-collapse:collapse'>"
		retorna = retorna & "<tr bgcolor='#D6E9FE'>"
    retorna = retorna & "<td colspan='8' class='textoMedio' style='padding:10px'><div align='center' id='qtdItens'><strong>Quantidade de Formandos:</strong><strong>" &qtdRegistros&"</strong></div></td>"
    retorna = retorna & "</tr>"

    'CABE�ALHO   					    
    retorna = retorna & "<tr bgcolor='whitesmoke'>"
    retorna = retorna & "<td width='6%'><div align='center'><strong>Ordem</strong></div></td>"
    retorna = retorna & "<td width='6%'><div align='center'><strong>C&oacute;d. Formando </strong></div></td>"
    retorna = retorna & "<td width='20%'><div align='center'><strong>Nome</strong></div></td>"
    retorna = retorna & "<td width='10%'><div align='center'><strong>&Aacute;lbum</strong></div></td>"
    retorna = retorna & "<td width='12%'><div align='center'> <div align='center'><strong>Item Pendente</strong></div></div></td>"
    retorna = retorna & "<td width='10%'><div align='center'><strong>Qtd Dias Fora</strong></div></td>"
    retorna = retorna & "<td width='14%'><div align='center'><strong>Status</strong></div></td>"
    retorna = retorna & "<td width='6%'><div align='center'><strong>Excluir da Devolu&ccedil;&atilde;o </strong></div></td>"
    retorna = retorna & "</tr>"

    'TABLE BODY
						
		Dim COD_FORMANDO, TIPO, NOME_VENDEDOR_NEW, COD_VENDEDOR_NEW, DATA_RETORNO, COD_MOVIMENTACAO, ENCADERNACOES_EXTRAS	
		i=qtdRegistros
		While not rs.Eof
			'*Significa possui diversos formandos						
			TIPO = Right(rs("COD_ETIQUETA"),1)
			COD_FORMANDO = Mid(rs("COD_ETIQUETA"),1,Len(rs("COD_ETIQUETA"))-1)
			
      Set rs2 = objConexao.Execute("SELECT * FROM VW_DEVOLUCAO_PRODUTOS WHERE COD_FORMANDO="&COD_FORMANDO&" AND TIPO=" & TIPO&" ORDER BY TIPO, DATA_SAIDA DESC LIMIT 1")
			
      DATA_RETORNO = ""
      
      If not rs2.Eof Then
        COD_MOVIMENTACAO = rs2("COD_MOVIMENTACAO")
        COD_VENDEDOR_NEW = rs2("COD_VENDEDOR")
        TIPO = CInt(rs2("TIPO"))
        DATA_RETORNO = ""&rs2("DATA_RETORNO")
        NOME_FORMANDO = rs2("NOME_FORMANDO")
        DATA_SAIDA = rs2("DATA_SAIDA")
        ENCADERNACOES_EXTRAS	= formatNumber(0&rs2("ENCADERNACOES_EXTRAS"),0)
        CONT_ITENS = CONT_ITENS + 1
        DATA_PREVISTA_RETORNO	= ""&rs2("DATA_PREVISTA_RETORNO")
        DIAS_ATRASADOS = DateDiff("d",rs2("DATA_PREVISTA_RETORNO"),Date())

        If int(DIAS_ATRASADOS) < 0 Then
          'N�o est� atrasado verifica se a ficha foi devolvida, pois se n�o for ir� barrar
        End If

        
        retorna = retorna & "<tr>"					
        retorna = retorna & "<td><div align='center'><strong>" & i & "°</strong></div></td>"
        retorna = retorna & "<td><div align='center'>" & digito(COD_FORMANDO) & "</div></td>"
        retorna = retorna & "<td><div align='left'>" & NOME_FORMANDO & "</div></td>"
        retorna = retorna & "<td><div align='center'>" & TIPO & "</div></td>"
        retorna = retorna & "<td><div align='center' class='textoMedio'><strong>"
        
        If TIPO = 0 Then
          If ENCADERNACOES_EXTRAS = 1 Then
            retorna = retorna & "Álbum Físico Duplo"
          Else
            retorna = retorna & "Álbum Físico"
          End If

        ElseIf TIPO = 1 Then
          retorna = retorna & "Ficha"

          motivo = ""
          'PEGA MOTIVO INFORMADO PELO VENDEDOR
          set rsMotivo = objConexao.execute("SELECT MOTIVO FROM TB_PRODUCAO_LOTES_FORMANDOS_MOTIVO_TEMPORARIA WHERE COD_FORMANDO = "&cod_formando&" AND COD_VENDEDOR = "&cod_vendedor)
          if not rsMotivo.Eof then
            motivo = rsMotivo("motivo").value
          end if
          rsMotivo.close()
          set rsMotivo = Nothing

          retorna = retorna & "<div style='font-size: 11px; font-weight: bold; background-color: #eee; border: 1px solid #000; padding: 0 3px 6px 3px;'> Motivo devolução"
          retorna = retorna & "<textarea name='motivo_"&COD_FORMANDO&"'>"&motivo&"</textarea>"
          retorna = retorna & "</div>"

        ElseIf TIPO = 2 Then
          retorna = retorna & "Álbum Digital"

        Else
          'Produto agregado, consultar qual é.
          Dim rsProdAgrePend
          sql = "SELECT TB_PRODUTOS.NOME_PRODUTO "&_
                "FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = "&_
                "TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO WHERE "&_
                "TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO = "&COD_FORMANDO&" AND "&_
                "TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.TIPO = "&TIPO&" LIMIT 1"
          Set rsProdAgrePend = objConexao.Execute(sql)
          If not rsProdAgrePend.Eof Then
            retorna = retorna & rsProdAgrePend("NOME_PRODUTO")
          End If
          rsProdAgrePend.Close
          Set rsProdAgrePend = Nothing
        End If

      
        retorna = retorna & "</strong></div></td>"
        retorna = retorna & "<td><div align='center'>"& DateDiff("d",DATA_SAIDA,Date()) &"</div></td>"
        retorna = retorna & "<td><table width='98%' border='0' cellspacing='0' cellpadding='0'>"

        If COD_VENDEDOR_NEW <> COD_VENDEDOR Then
          'Verifica se s�o c�digos diferentes de vendedor
          ERROR_LINHA = 2
          retorna = retorna & "<tr>"
          retorna = retorna & "<td width='13%'><img src='images/icons/icon_red.gif' width='17' height='15' /></td>"
          retorna = retorna & "<td width='87%'>Vendedores diferentes </td>"
          retorna = retorna & "</tr>"
        Elseif Len(DATA_RETORNO) > 3 Then
          ERROR_LINHA = 2
          retorna = retorna & "<tr>"
          retorna = retorna & "<td><img src='images/icons/icon_green.gif' width='17' height='15' /></td>"
          retorna = retorna & "<td align='center'> ITEM CONSTA COMO DEVOLVIDO! </td>"
          retorna = retorna & "<td><img src='images/icons/icon_green.gif' width='17' height='15' /></td>"
          retorna = retorna & "</tr>"
        Else
          retorna = retorna & "<tr>"
          retorna = retorna & "<td colspan='2'><div align='center'>"
          retorna = retorna & "<img src='images/icons/icon_green.gif' width='16' height='15' /></div></td>"
          retorna = retorna & "</tr>"
        End If

        retorna = retorna & "</table></td>"
        retorna = retorna & "<td><div align='center'>"
        retorna = retorna & "<a href='?comando=excluir&cod_etiqueta="& COD_FORMANDO & TIPO &"'>"
        retorna = retorna & "<img src='images/icons/icon_excluir.png' width='16' height='16' border='0' /></a>"
        retorna = retorna & "<input name='COD_MOVIMENTACAO_"& COD_FORMANDO &"_"& TIPO &"' value='"&COD_MOVIMENTACAO&"' type='hidden' />"
        retorna = retorna & "</div></td>"
        retorna = retorna & "</tr>"
      End If

      rs2.Close
      Set rs2 = Nothing	
      i=i-1
    rs.MoveNext()
    Wend
    retorna = retorna & "</table>"
  End If
		
  rs.Close
  Set rs=Nothing
          
  if qtdRegistros > 0 then
    retorna = retorna & "<div align='center'>"
    if ERROR_LINHA <> 2 Then
      retorna = retorna & "<table width='98%' border='0' cellspacing='0' cellpadding='5'>"
      retorna = retorna & "<tr><td><div align='center'>"
      retorna = retorna & "<input class='btnAvancar' name='btnAvancar' type='submit' value='' />"
      retorna = retorna & "<input name='COD_VENDEDOR' type='hidden' value='" & COD_VENDEDOR & "' />"
      retorna = retorna & "</div></td></tr></table>"
    Else
      retorna = retorna & "<span class='msg'>Atenção, não será possível avançar, pois h&aacute; itens com status em vermelho ❌ ou verde que está OK ✅.<br />Se o item está como devolvido volte para a tela inicial ou coloque outro código.</span>"
    End If
    retorna = retorna & "</div>"
  End If
End If

Response.Write(retorna)

FecharConexao()

%>