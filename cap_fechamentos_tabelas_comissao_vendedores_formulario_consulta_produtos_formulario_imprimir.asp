<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Validação especial para verificar se possui acesso a página mãe de formulário
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS where COD_PAGINA = '228' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se não encontrar então redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
		
	rs.Close
	Set rs=Nothing
	
FecharConexao()

'/////////////////////////////////////////////////////////////////////////////////////////////

'Declara variáveis para utilização do sistema
Dim cod_tabela_comissao_vendedores
cod_tabela_comissao_vendedores = Request.QueryString("cod_tabela_comissao_vendedores")
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="all"></link>
<title>Tabelas</title>
</head>
<body style="background:#FFF">
<!-- End code for the left panel -->
<table cellpadding="0px" cellspacing="0" bgcolor="#FFFFFF" width="100%">
<tr>
<td>
<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="10">
      <tr>
        <td>
          
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td>
<h2><%=titulo_pagina%></h2>
<h3 align="center">Produtos da 
<%
	AbrirConexao()

			Set rs=objConexao.Execute("SELECT descricao FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES where cod_tabela_comissao_vendedores = "&cod_tabela_comissao_vendedores)
				
				If not rs.Eof Then
					Response.Write(rs(0))
				End If
			
			rs.Close
			Set rs=Nothing

	FecharConexao()

%>
</h3>

<%
AbrirConexao()
			Set rs=objConexao.Execute("SELECT cod_produto, nome_produto FROM stillos3.TB_PRODUTOS WHERE ativo = 1 and PAGAMENTO_COMISSAO = 'S' order by order_menu")
				
				If not rs.Eof Then%>
				<table width='750' border='0' align='center' cellpadding='3' cellspacing="0" style='border-collapse:collapse'>
<tr>
  <td width='36%' rowspan="2"><div align="center"><strong>Produto</strong></div></td>
<td colspan="2"><div align="center"><strong>Vendas Normais</strong></div></td>
<td colspan="2"><div align="center"><strong>Vendas Repasse</strong></div></td>
</tr>
<tr>
  <td width="17%"><div align="center"><strong>&Agrave; vista</strong></div></td>
  <td width="15%"><div align="center"><strong>&Agrave; prazo</strong></div></td>
  <td><div align="center"><strong>&Agrave; vista</strong></div></td>
  <td><div align="center"><strong>&Agrave; prazo</strong></div></td>
</tr>
				<%
				Dim cor, caminho_validacao_completa
				
				cor = "#FFFFCC"
				caminho_validacao_completa = ""
				
					While not rs.Eof
					
					cod_produto = rs("cod_produto") 
					
					'///////////////////////////////////////////////////////////
					'Verifica se possui o produto cadastrado nesta tabela de venda
					Set rs2=objConexao.Execute("SELECT * FROM TB_FECHAMENTO_COMISSIONAMENTO_VENDEDORES_PRODUTOS WHERE cod_tabela_comissao_vendedores = "&cod_tabela_comissao_vendedores&" and cod_produto ="&cod_produto)
				
						If not rs2.Eof Then
							tipo_pagamento_normal_avista=trim(rs2("tipo_pagamento_normal_avista"))
							valor_pagamento_normal_avista= ArrumaNumero(rs2("valor_pagamento_normal_avista"),"",2)
							tipo_pagamento_normal_aprazo=trim(rs2("tipo_pagamento_normal_aprazo"))
							valor_pagamento_normal_aprazo= ArrumaNumero(rs2("valor_pagamento_normal_aprazo"),"",2)
							tipo_pagamento_repasse_avista=trim(rs2("tipo_pagamento_repasse_avista"))
							valor_pagamento_repasse_avista= ArrumaNumero(rs2("valor_pagamento_repasse_avista"),"",2)
							tipo_pagamento_repasse_aprazo=trim(rs2("tipo_pagamento_repasse_aprazo"))
							valor_pagamento_repasse_aprazo= ArrumaNumero(rs2("valor_pagamento_repasse_aprazo"),"",2)

							
						Else
							tipo_pagamento_normal_avista=""
							valor_pagamento_normal_avista=""
							tipo_pagamento_normal_aprazo=""
							valor_pagamento_normal_aprazo=""
							tipo_pagamento_repasse_avista=""
							valor_pagamento_repasse_avista=""
							tipo_pagamento_repasse_aprazo=""
							valor_pagamento_repasse_aprazo=""					
						End If
					
					rs2.Close
					Set rs2=Nothing
					
					'//////////////////////////////////////////////////////////
					'Validação
						caminho_validacao_completa = caminho_validacao_completa&"|TIPO_PAGAMENTO_NORMAL_AVISTA_"&cod_produto&"#Tipo de Pagamento para Vendas Normais#2|VALOR_PAGAMENTO_NORMAL_AVISTA_"&cod_produto&"#Valor Pagamento para Vendas Normais#1|TIPO_PAGAMENTO_NORMAL_APRAZO_"&cod_produto&"#Tipo de Pagamento para Vendas Normais#2|VALOR_PAGAMENTO_NORMAL_APRAZO_"&cod_produto&"#Valor Pagamento para Vendas Normais#1|TIPO_PAGAMENTO_REPASSE_AVISTA_"&cod_produto&"#Tipo de Pagamento para Vendas Repasse#2|VALOR_PAGAMENTO_REPASSE_AVISTA_"&cod_produto&"#Valor Pagamento para Vendas Repasse#1|TIPO_PAGAMENTO_REPASSE_APRAZO_"&cod_produto&"#Tipo de Pagamento para Vendas Repasse#2|VALOR_PAGAMENTO_REPASSE_APRAZO_"&cod_produto&"#Valor Pagamento para Vendas Repasse#1"
					
					
					
					'///////////////////////////////////////////////////////////					
					If cor = "#FFFFCC" Then
						cor = ""
					Else
						cor = "#FFFFCC"
					End If
%>
<tr bgcolor="<%=cor%>">
  <td style="border-bottom:solid 1px #CCC"><div align="left"><strong><%=rs("nome_produto")%></strong></div></td>
  <td style="border-bottom:solid 1px #CCC">
    <div align="right">
      <%If CStr(TIPO_PAGAMENTO_NORMAL_AVISTA) =  "F" Then%>
      R$ <%=VALOR_PAGAMENTO_NORMAL_AVISTA%>
      <%Else%>
      % <%=VALOR_PAGAMENTO_NORMAL_AVISTA%>
      <%End if%>
    </div></td>
  <td style="border-bottom:solid 1px #CCC">
    <div align="right">
      <%If CStr(TIPO_PAGAMENTO_NORMAL_APRAZO) =  "F" Then%>
      R$ <%=VALOR_PAGAMENTO_NORMAL_APRAZO%>
      <%Else%>
      % <%=VALOR_PAGAMENTO_NORMAL_APRAZO%>
      <%End if%>
    </div></td>
  <td style="border-bottom:solid 1px #CCC">  <div align="right">
    <%If CStr(TIPO_PAGAMENTO_REPASSE_AVISTA) =  "F" Then%>
    R$ <%=VALOR_PAGAMENTO_REPASSE_AVISTA%>
    <%Else%>
    % <%=VALOR_PAGAMENTO_REPASSE_AVISTA%>
    <%End if%>
  </div></td>
  <td style="border-bottom:solid 1px #CCC">  <div align="right">
    <%If CStr(TIPO_PAGAMENTO_REPASSE_APRAZO) =  "F" Then%>
    R$ <%=VALOR_PAGAMENTO_REPASSE_APRAZO%>
    <%Else%>
    % <%=VALOR_PAGAMENTO_REPASSE_APRAZO%>
    <%End if%>
  </div></td>
  </tr>

<%

					rs.MoveNext()
					Wend
					
					%>
					</table>
				
					<%
				
				End If
			
			rs.Close
			Set rs=Nothing
FecharConexao()
%>
              </td>
              
              </tr>
  </table>
          
          
          
          </td>
        </tr>
</table>

  
  </td>
  </tr>
</table></td>
</tr>
</table>
</body>
</html>