<% Response.Charset="ISO-8859-1" %>
<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%

if not validaUsuario(1,391,0,false) then
  response.end()
end if

Dim tipo_solicitacao, retorna
Dim parametros(11)
tipo_solicitacao = Cstr(""&Request.QueryString("ts"))

For i=1 to 10
	parametros(i) = Request.QueryString("p"&i)
Next 

function informacaoInventario(cod_item) 

  dim cod_formando, tipo_produto, item_values
  cod_formando = sqlNumero(left(cod_item, len(cod_item)-1))
  tipo_produto = sqlNumero(right(cod_item, 1))
  item_values = array("Albúm", "Ficha")
  AbrirConexao()

  'validações
  erro = ""
  'verifica se o item não foi vendido 
  set rs = objConexao.execute("SELECT COD_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE vendido = 1 and COD_FORMANDO = "&cod_formando)
  if not rs.Eof then
    erro = "Item consta como vendido"
  end if
  rs.close
  set rs = nothing

  if len(erro) = 0 then

    dim sql : sql = "SELECT b.cod_formando, b.nome_formando, c.DESCRICAO, b.nome_curso, b.cod_producao FROM TB_FICHAS_IDENTIFICACAO as b  INNER JOIN TB_ENTIDADES as c ON c.COD_CONTRATO = b.COD_CONTRATO WHERE cod_formando = "&cod_formando&" LIMIT 1"
    set rs = objConexao.execute(sql, qtdRegistros)
    if not rs.Eof then

      retorno = ""
      retorno = retorno & "<h3 align='center'>Novo lançamento</h3>"
      retorno = retorno & "<table width='800px' border='1' align='center' cellpadding='2' style='border-collapse:collapse'>"
      retorno = retorno & "<tr bgcolor='whitesmoke'>"
      retorno = retorno & "<td width='6%'><div align='center'><strong>Cód. Formando</strong></div></td>"
      retorno = retorno & "<td width='20%'><div align='center'><strong>Nome Formando</strong></div></td>"
      retorno = retorno & "<td width='20%'><div align='center'><strong>Item</strong></div></td>"
      retorno = retorno & "</tr>"

      While not rs.Eof
        retorno = retorno & "<tr bgcolor='whitesmoke'>"
        retorno = retorno & "<td width='6%'><div align='center'><strong>" & rs("cod_formando").value & "</strong></div></td>"
        retorno = retorno & "<td width='20%'><div align='center'><strong>"& rs("nome_formando").value &"</strong></div></td>"
        retorno = retorno & "<td width='10%'><div align='center'><strong>"& item_values(tipo_produto) & "</strong></div></td>"
        retorno = retorno & "</tr>"
      rs.MoveNext
      Wend

      retorno = retorno & "<div align='center'>"
      retorno = retorno & "<table width='98%' border='0' cellspacing='0' cellpadding='5'>"
      retorno = retorno & "<tr><td><div align='center'>"
      retorno = retorno & "<input name='cod_item' type='hidden' value='"&cod_item&"' />"
      retorno = retorno & "<input class='btnAvancar' name='btnAvancar' type='submit' value='' />"
      retorno = retorno & "</div></td></tr></table>"
      retorno = retorno & "</div>"
    else
      retorno = "<!--#true#displayLeitura('erro','Item não localizado');#-->"
    end if 
    rs.close
    set rs = nothing
  else
    retorno = "<!--#true#displayLeitura('erro','"&erro&"');#-->"
  end if 

  FecharConexao()

informacaoInventario = retorno
end function

dim retorno : retorno = ""
if tipo_solicitacao = "incluir" then
  retorno = informacaoInventario(parametros(1)) 
end if

Response.Write(retorno)

%>