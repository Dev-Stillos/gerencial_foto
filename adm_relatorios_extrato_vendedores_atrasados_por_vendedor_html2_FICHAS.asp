<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<%
If not validaUsuario(2,,15,false) Then
  Response.Redirect("inicio.asp")
End If

Dim cod_vendedor_bloqueado, com_copia, somente_vendedores_em_atraso
cod_vendedor_bloqueado = cint(Request.QueryString("COD_VENDEDOR_BLOQUEADO"))
com_copia	= cint(Request.QueryString("COM_COPIA"))
somente_vendedores_em_atraso = Cstr(Request.QueryString("SOMENTE_VENDEDORES_EM_ATRASO"))

If somente_vendedores_em_atraso = "" AND Request.QueryString("COD_VENDEDOR") <> Empty Then
	somente_vendedores_em_atraso = "N"
	cod_vendedor_bloqueado = Request.QueryString("COD_VENDEDOR")
End If

If cod_vendedor_bloqueado = 0 Then
	Response.Redirect("adm_relatorios_vendedores_bloqueados.asp")
End If

AbrirConexao()

Dim sql, sqlMedia, sqlWhere, registros, qtd_paginas, data_inicial, qtd_dias_devedores_padrao

sqlMedia = "SELECT AVG(DIAS_ATRASADOS) AS MEDIA FROM VW_MOVIMENTACOES_COMPLETA WHERE DATA_RETORNO IS NULL AND COD_VENDEDOR = " & COD_VENDEDOR_BLOQUEADO


if len(somente_vendedores_em_atraso) = 0 then
  sqlWhere = " AND DIAS_ATRASADOS >= 0 "
end if

sqlMedia = sqlMedia & sqlWhere
Set rs2 = objConexao.Execute(sqlMedia)
  If not rs2.Eof Then
    media_geral = cint(rs2("MEDIA").value)
  End If
rs2.Close
Set rs2 = Nothing

sql = "SELECT COD_FORMANDO, NOME_FORMANDO, NOME_VENDEDOR, TIPO, DATA_SAIDA, ENCADERNACOES_EXTRAS, DIAS_ATRASADOS, ALBUM, DESCRICAO, ENCADERNACOES_EXTRAS FROM VW_MOVIMENTACOES_COMPLETA WHERE COD_VENDEDOR = " & COD_VENDEDOR_BLOQUEADO & " AND DATA_SAIDA IS NOT NULL AND DATA_RETORNO IS NULL AND TIPO = 1 " & sqlWhere & " ORDER BY DESCRICAO, NOME_FORMANDO "

Set rs = objConexao.Execute(sql,registros)
if not rs.Eof then

  %>
    <html xmlns='http://www.w3.org/1999/xhtml'>
    <meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1' />
    <title>RELATÓRIO DE FICHAS EM POSSE DO VENDEDOR</title>
    <script type="text/javascript" src="js/jquery-1.11.2.min.js"></script>
    <body>
    <style>
    * {-webkit-print-color-adjust: exact; }
    .left { position: relative; float: left; }
    body { margin: 0; padding: 0; font-family: Arial, Helvetica, sans-serif; }
    table { width: 100%; border: 1px solid black; border-collapse: collapse; }
    .pagina { margin: 0; width: 210mm; height: 300mm; padding: 12px; overflow: hidden; page-break-after: always; }
    .barraCinza { width: 100%; height: 2px; background: #aeaeae; margin: 4px 0 4px 0; }
    .titulo { width: 100%; text-align: center; font-size: 22px; font-weight: bold; }
    .cabecalho { padding: 5px 0 5px 0; }
    .relatorio { font-family: Arial, Helvetica, sans-serif; font-size: 10px; margin: 4px 0 4px 0; }
    .campo { height: 30px; font-size: 16px; border-left: 1px solid #aeaeae; padding: 5px; font-weight: bold; }
    .campo div { font-weight: normal; font-size: 11px; }
    .campo:first-child { border: none; padding-left: 0; margin: 0; }
    .campo1 { width: 300px; }
    .campo2 { width: 150px; }
    .campo3 { width: 90px; }
    .campo4 { width: 90px; }
    .campo5 { width: 100px; }
    .data { position: absolute; bottom: 15px; height: 15px; left: 12px; }
    .paginacao { position: absolute; bottom: 15px; height: 15px; right: 12px; }
    .assinaturas { width: 600px; font-size: 16px; }
    .assinaturas div{ overflow: hidden; height: 16px; margin: 4px 0 4px 0; }
    .assinaturas div:nth-child(odd) { width: 220px; }
    .assinaturas div:nth-child(even) { border-bottom: 1px solid black; width: 60%; }
    </style>
  <%

  dim nome_vendedor, data_emissao, prazo, horas, minutos, itens_pagina, pagina_atual, tipo, rsProdAgrePend
  dim encadernacoes_extras
  dim itens_per_page : itens_per_page = 65
  qtd_paginas = int(registros/itens_per_page)
  resto = registros Mod itens_per_page
  nome_vendedor	= ucase(rs("NOME_VENDEDOR").value)
  horas	= Hour(Now())
  minutos	= Minute(Now())
	itens_pagina = 0
  count_total = 0
  pagina_atual = 1

  If horas < 10 Then
      horas = "0"&horas
  End If
  
  If minutos < 10 Then
      minutos = "0"&minutos
  End If

  data_emissao = ArrumaData(date()) & " " & horas &":"& minutos

  If resto > 0 Then
    qtd_paginas = qtd_paginas + 1
  End If

  while not rs.eof 
    if itens_pagina = 0 then
      %>
        <div class='pagina left'>
        <div class='titulo'>
        RELATÓRIO DE FICHAS EM POSSE DO VENDEDOR
        </div>
        <div class='barraCinza left'></div>
        <div class='cabecalho'>
        <div class='campo left campo1'>
        <div>Representante Comercial</div>
        <%=nome_vendedor%>
        </div>
        <div class='campo left campo2'>
        <div>Data emissão</div>
        <%=data_emissao%>
        </div>
        <div class='campo left campo3'>
        <div>Prazo normal</div>
        Variável
        </div>
        <div class='campo left campo4'>
        <div>Média de atraso</div>
        <%=media_geral%>
        </div>
        <div class='campo left campo5'>
        <div>Fichas em posse</div>
        <%=registros%>
        </div>
        </div>
        <div class='barraCinza left'></div>
        <table class='relatorio left' border='1'>
        <tr>
        <tr>
        <th>Cód. Form.</th>
        <th>Ficha</th>
        <th>Nome Formando</th>
        <th>Instituição de Ensino</th>
        <th>Data Saída</th>
        <th>Dias Excedentes</th>
        </tr>
        </tr>
      <%
    end if
    tipo = int(rs("tipo").value)
    encadernacoes_extras = int(rs("encadernacoes_extras").value)
    If TIPO = 0 Then					
      If encadernacoes_extras = 1 Then
        TIPO = "Álbum DUPLO"
      Else
        TIPO = "Álbum"
      End If				
		'Elseif TIPO = 1 Then
        'TIPO = "FICHA"
    Elseif TIPO > 1 Then
      'Verifica qual o produto agregado
        sql = 	"SELECT TB_PRODUTOS.NOME_PRODUTO "&_
            "FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = "&_
            "TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO WHERE "&_
            "TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO = "&rs("COD_FORMANDO")&" AND "&_
            "TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.TIPO = "&TIPO&" LIMIT 1"
        Set rsProdAgrePend = objConexao.Execute(sql)
          If not rsProdAgrePend.Eof Then
            TIPO = rsProdAgrePend("NOME_PRODUTO")
          End If
        rsProdAgrePend.Close
        Set rsProdAgrePend = Nothing
    End If

    %>
     <tr>
                <td><%=digito(rs("COD_FORMANDO").value)%></td>

        

        <% 'Coluna Digital %>
        <td>
          <% If rs("TIPO") = 1 Then %>
             Ficha
          <% End If %>
        </td>

        <td><%=ucase(left(rs("NOME_FORMANDO").value,30))%></td>
        <td><%=ucase(left(rs("DESCRICAO").value,35))%></td>
        <td><%=arrumadata(rs("DATA_SAIDA").value)%></td>
        <td><%=rs("DIAS_ATRASADOS").value%></td>

      </tr>
    <%
    count_total = count_total + 1
    itens_pagina = itens_pagina + 1

    if itens_pagina = itens_per_page or count_total = registros then
      itens_pagina = 0
      pagina_atual = pagina_atual + 1 
      %>
      </table>
      <div class='barraCinza left'></div>
      <div class='assinaturas'>
      <div class="left">Assinatura do Emissor</div>
      <div class="left"></div>
      <div class="left">Assinatura do Representante Comercial</div>
      <div class="left"></div>
      </div>
      <div class='barraCinza left'></div>
      <div class='data'><%=data_emissao%></div>
      <div class='paginacao'><%=pagina_atual%> de <%=qtd_paginas%></div>
      </div>
      <%
    end if
  rs.MoveNext()
  wend
  %>
    </body>
    <script>//$(function(){window.print();});</script>
    </html>
  <%
  end if
FecharConexao()	

%>