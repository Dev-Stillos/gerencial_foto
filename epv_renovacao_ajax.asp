<% Response.Charset="ISO-8859-1" %>
<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%

if not validaUsuario(1,388,0,false) then
  response.end()
end if

Dim tipo_solicitacao, retorna
Dim parametros(11)
tipo_solicitacao = Cstr(""&Request.QueryString("ts"))

For i=1 to 10
	parametros(i) = Request.QueryString("p"&i)
Next 

function listarRenovacoes() 
  AbrirConexao()

  dim sql : sql = " SELECT COD_USUARIO, "&_
                  " TB_FICHAS_IDENTIFICACAO_RENOVACAO_TEMPORARIA.COD_VENDEDOR,"&_
                  " TB_VENDEDORES.NOME_VENDEDOR,"&_
                  " TB_FICHAS_IDENTIFICACAO.COD_FORMANDO, "&_
                  " TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, "&_
                  " TB_FICHAS_IDENTIFICACAO_RENOVACAO_TEMPORARIA.COD_PRODUCAO, "&_
                  " TIPO, "&_
                  " MOTIVO, "&_
                  " PERIODO_RENOVACAO "&_
                  " FROM TB_FICHAS_IDENTIFICACAO_RENOVACAO_TEMPORARIA"&_
                  " INNER JOIN TB_VENDEDORES ON TB_VENDEDORES.COD_VENDEDOR = TB_FICHAS_IDENTIFICACAO_RENOVACAO_TEMPORARIA.COD_VENDEDOR"&_
                  " INNER JOIN TB_FICHAS_IDENTIFICACAO ON TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO_RENOVACAO_TEMPORARIA.COD_FORMANDO "&_
                  " WHERE COD_USUARIO = " & sqlNumero(Session("COD_USUARIO"))

  set rs = objConexao.execute(sql, qtdRegistros)
  if not rs.Eof then
    
    retorno = ""
    retorno = retorno & "<h3 align='center'>Devolução de Produtos do vendedor(a):<br />" & NOME_VENDEDOR & "</h3>"
    retorno = retorno & "<table width='95%' border='1' align='center' cellpadding='2' style='border-collapse:collapse'>"
		retorno = retorno & "<tr bgcolor='#D6E9FE'>"
    retorno = retorno & "<td colspan='10' class='textoMedio' style='padding:10px'><div align='center' id='qtdItens'><strong>Quantidade de Formandos:</strong><strong>" &qtdRegistros&"</strong></div></td>"
    retorno = retorno & "</tr>"

    'CABE�ALHO   					    
    retorno = retorno & "<tr bgcolor='whitesmoke'>"
    retorno = retorno & "<td width='6%'><div align='center'><strong>Ordem</strong></div></td>"
    retorno = retorno & "<td width='6%'><div align='center'><strong>Cód. Formando</strong></div></td>"
    retorno = retorno & "<td width='8%'><div align='center'><strong>Tipo</strong></div></td>"
    retorno = retorno & "<td width='14%'><div align='center'><strong>Nome Formando</strong></div></td>"
    retorno = retorno & "<td width='10%'><div align='center'><strong>Vendedor</strong></div></td>"
    retorno = retorno & "<td width='6%'><div align='center'><strong>Cód. Produção</strong></div></td>"
    retorno = retorno & "<td width='10%'><div align='center'><strong>Periodo</strong></div></td>"
    retorno = retorno & "<td width='10%'><div align='center'><strong>Motivo Renovação</strong></div></td>"
    retorno = retorno & "<td width='14%'><div align='center'><strong>Status</strong></div></td>"
    retorno = retorno & "<td width='6%'><div align='center'><strong>Remover</strong></div></td>"
    retorno = retorno & "</tr>"

    count = 1 
   While not rs.Eof
      retorno = retorno & "<tr bgcolor='whitesmoke'>"
      retorno = retorno & "<td><div align='center'><strong>"&count&"</strong></div></td>"
      retorno = retorno & "<td><div align='center'><strong>" & rs("cod_formando").value & "</strong></div></td>"
' <<< NOVO BLOCO: coluna TIPO
Dim tipoNome
Select Case rs("tipo").value
    Case 0
        tipoNome = "Álbum Físico"
    Case 1
        tipoNome = "Ficha"
    Case 2
        tipoNome = "Álbum Digital"
    Case Else
        tipoNome = "Desconhecido"
End Select
retorno = retorno & "<td><div align='center'><strong>" & tipoNome & "</strong></div></td>"
retorno = retorno & "<td><div align='center'><strong>"& rs("nome_formando").value &"</strong></div></td>"
retorno = retorno & "<td><div align='center'><strong>"& rs("nome_vendedor").value & "</strong></div></td>"

      
      retorno = retorno & "<td><div align='center'><strong>"& rs("cod_producao").value & "</strong></div></td>"
      retorno = retorno & "<td><div align='center'><strong>"& rs("PERIODO_RENOVACAO").value & "</strong></div></td>"
      retorno = retorno & "<td><div align='center'>"
      retorno = retorno & "<div style='font-size: 11px; font-weight: bold; background-color: #eee; border: 1px solid #000; padding: 3px 3px 6px 3px;'>"
      retorno = retorno & "<textarea name='motivo_"&rs("cod_formando").value&rs("tipo").value&"'>"& rs("motivo").value &"</textarea>"
      retorno = retorno & "</div></td>"
      retorno = retorno & "<td><div align='center'><strong>VERIFICAR</strong></div></td>"
      retorno = retorno & "<td><div align='center'><strong>"
      retorno = retorno & "<a href='javascript:removerItemRenovacao("& rs("cod_formando").value &");'>"
      retorno = retorno & "<img src='images/icons/icon_excluir.png' width='16' height='16' border='0' /></a>"
      retorno = retorno &"</strong></div></td>"
      retorno = retorno & "</tr>"
    count = count + 1
    rs.MoveNext
    Wend
  end if 
  rs.close
  set rs = nothing

  if qtdRegistros > 0 then
    retorno = retorno & "<div align='center'>"
    retorno = retorno & "<table width='98%' border='0' cellspacing='0' cellpadding='5'>"
    retorno = retorno & "<tr><td><div align='center'>"
    retorno = retorno & "<input class='btnAvancar' name='btnAvancar' type='submit' value='' />"
    retorno = retorno & "</div></td></tr></table>"
    retorno = retorno & "</div>"
  End If

  FecharConexao()

listarRenovacoes = retorno
end function

function removerRenovacao(cod_formando)

  AbrirConexao()

  cod_usuario = sqlNumero(session("cod_usuario"))
  objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_RENOVACAO_TEMPORARIA WHERE cod_formando = "&sqlNumero(cod_formando) & " AND cod_usuario = "& cod_usuario)

  FecharConexao()

removerRenovacao = "<!--#true#carregaListagem();displayLeitura('sucesso','Item removido com sucesso: "&cod_formando&"1');#-->"
end function

function incluirRenovacao(cod_item)
  AbrirConexao()

  dim config, cod_formando, tipo_produto, cod_usuario, erro, retorno

  retorno = ""
  erro = ""
  cod_formando = left(cod_item, len(cod_item)-1)
  tipo_produto = right(cod_item, 1)
  cod_usuario = Session("COD_USUARIO")
  set config = getConfiguracoes(true, array("periodo_renovacao_padrao"))

  ' Permite: Álbum físico (0), Ficha (1) e Álbum digital (2)
    if not (tipo_produto = 0 or tipo_produto = 1 or tipo_produto = 2) then
        erro = "Somente álbum físico (0), ficha (1) ou álbum digital (2) podem ser renovados. Demais produtos devem ser devolvidos e retirados posteriormente"
    end if


  '1) Pega o vendedor atrelado a renovacao
  if len(erro) = 0 then
    set rs = objConexao.execute("SELECT cod_vendedor, data_retorno FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_FORMANDO = "&sqlNumero(cod_formando) &" and tipo = "&sqlNumero(tipo_produto)&" ORDER BY cod_movimentacao DESC LIMIT 1")
    If not rs.Eof Then
      cod_vendedor = rs("cod_vendedor").value
    else
      erro = "Item não localizado"
    End If
    rs.close()
    Set rs = nothing
  end if

  '1) Pega o vendedor atrelado a renovacao
  if len(erro) = 0 then
    set rs = objConexao.execute("SELECT cod_vendedor, data_retorno FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_FORMANDO = "&sqlNumero(cod_formando) &" and tipo = "&sqlNumero(tipo_produto)&" AND data_retorno IS NULL ORDER BY cod_movimentacao DESC LIMIT 1")
    If rs.Eof Then
      erro = "Item consta como devolvido"
    End If
    rs.close()
    Set rs = nothing
  end if

  '2) verifica se exite outros vendedor na rodada de devolu��o atual
  if len(erro) = 0 then
    set rs = objConexao.execute("SELECT cod_vendedor FROM TB_FICHAS_IDENTIFICACAO_RENOVACAO_TEMPORARIA WHERE cod_vendedor = "&sqlNumero(cod_vendedor) &" and cod_usuario = "&sqlNumero(cod_usuario) & " AND cod_vendedor != "&sqlNumero(cod_vendedor) )
    If not rs.Eof Then
      erro = "Essa ficha está em posse de outro vendedor"
    End If
    rs.close()
    Set rs = nothing
  end if

  '3) valida se o item informado est� com o vendedor informado
 	if len(erro) = 0 then
    Set rs = objConexao.execute("SELECT cod_renovacao_temp FROM TB_FICHAS_IDENTIFICACAO_RENOVACAO_TEMPORARIA WHERE COD_USUARIO = "& sqlNumero(cod_usuario) &" AND COD_FORMANDO = " & sqlNumero(cod_formando) & " AND TIPO = "&sqlNumero(tipo_produto))
    If not rs.Eof Then 'Pode gravar, pois não foi inserido ainda	
      erro = "Item já inserido: "&cod_item
    End If
    rs.close()
    Set rs = nothing
  end if

  if len(erro) = 0 then
    dim motivo : motivo = ""
    'PEGA MOTIVO INFORMADO PELO VENDEDOR
    set rs = objConexao.execute("SELECT MOTIVO FROM TB_PRODUCAO_LOTES_FORMANDOS_MOTIVO_TEMPORARIA WHERE COD_FORMANDO = "&cod_formando&" AND COD_VENDEDOR = "&cod_vendedor)
    if (not rs.Eof) then
      motivo = rs("motivo").value
    end if
    rs.close()

    dim cod_producao : cod_producao = ""
    set rs = objConexao.execute("SELECT cod_producao FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO = "&sqlNumero(cod_formando))
    If not rs.Eof Then
      cod_producao = rs("cod_producao").value
    End If
    rs.close()
    Set rs = nothing
    
    sql = "INSERT INTO TB_FICHAS_IDENTIFICACAO_RENOVACAO_TEMPORARIA (COD_USUARIO, COD_VENDEDOR, COD_FORMANDO, TIPO, PERIODO_RENOVACAO, MOTIVO, COD_PRODUCAO ) VALUES ("&sqlNumero(cod_usuario)&", "&sqlNumero(cod_vendedor)&", "&sqlNumero(cod_formando)&", "&sqlNumero(tipo_produto)&", "&sqlNumero(config("periodo_renovacao_padrao"))&", "&sqlTexto(motivo)&", "&sqlNumero(cod_producao)&")"

	  objConexao.execute(sql)
    retorno = "<!--#true#carregaListagem();displayLeitura('sucesso','Item inserido com sucesso: "&cod_item&"');#-->"
  else
    retorno = "<!--#true#displayLeitura('erro','"&erro&"');#-->"
  end if

  set config = nothing

  FecharConexao()
incluirRenovacao = retorno
end function

dim retorno : retorno = ""
if tipo_solicitacao = "incluir" then
  retorno = incluirRenovacao(parametros(1))
elseif tipo_solicitacao = "listar" then
  retorno = listarRenovacoes() 
elseif tipo_solicitacao = "remover" then
  retorno = removerRenovacao(parametros(1))
end if

Response.Write(retorno)

%>