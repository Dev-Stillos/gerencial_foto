<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<%
Function ExtrairNumeros(strTexto)
	Dim objRegExp, strNumeros
	Set objRegExp = New RegExp
 
    objRegExp.Global = True ' Garante que a substituicao ocorra em toda a string
    objRegExp.Pattern = "[^\d]" ' Define o padrao para encontrar qualquer caractere que NAO seja um digito
 
    strNumeros = objRegExp.Replace(strTexto, "") ' Substitui os caracteres nao numericos por ""
 
    ExtrairNumeros = strNumeros
 
    Set objRegExp = Nothing
End Function
%>
 
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Versão 1.2</title>
</head>

<body>
<table cellpadding="9px" cellspacing="0" width="100%">
<tr>
<td>
<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="10">
  <tr>
  <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">

	  <h2> </h2>
	<%
		  AbrirConexao()
		  Dim rs, rs2, rs3
		  Dim QTD_ALBUNS, FICHAS_VENDA

		  Set rs = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO="&Request.Form("cod_prod")&" ORDER BY NOME_FORMANDO")
		  
		  If not rs.Eof then
		  %>
		  <h3 align="center">Tela de Finalização </h3>
		  <h3 align="center"><%=DESCRICAO%></h3>
		  <div id="comboTexto">
	
          <table width="98%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="26%"><div align="center"><strong>Pasta</strong></div></td>
            <td width="26%"><div align="center"><strong>Álbum</strong></div></td>
			<td width="10%"><div align="center"><strong>Álbum Digital</strong></div></td>
            <td width="25%"><div align="center"><strong>Inicial</strong></div></td>
            <td width="10%"><div align="center"><strong>Comissão</strong></div></td>
            <td width="11%"><div align="center"><strong>Brinde</strong></div></td>
            <td width="15%"><div align="center"><strong>Prod. Agreg. </strong></div></td>
            <td width="13%"><div align="center"><strong>Lote</strong></div></td>
</tr>
		  <%
		  Dim QTD_PREENCHIDA
		   QTD_PREENCHIDA = 0
		  While not rs.Eof
		  
		  	'Verifica se há não há erro, se não houver registra e deleta na tabela temporária
			If Request.Form("ERROR_FLAG_"&rs("COD_FORMANDO")) = 2 then
			
			'////////////////////////////////////////////////////////////////////////////////////////////////
			'MOMENTO CRUCIAL, NESTA ETAPA OS DADOS DA PÁGINA ANTERIOR SERÃO INCLUÍDAS NUMA TABELA TEMPORÁRIA
			'1) Guarda os valores em variáveis da pag. anterior
			Dim ALBUM_FORM, ALBUM_DIGITAL_FORM, INICIAL_FORM, COMISSAO_FORM, BRINDE_FORM, PROD_AGREG_FORM, VENDEDOR_FORM, MODIFICACOES, ENCADERNACOES_EXTRAS_FORM, LOTE_FORM
			
			ALBUM_FORM = CLng(0 & Request.Form("ALBUM_" & rs("COD_FORMANDO")))
			ALBUM_DIGITAL_FORM = CLng(0 & Request.Form("ALBUM_DIGITAL_" & rs("COD_FORMANDO"))) <!--Aqui o 0 & valor garante que, se vier vazio, será tratado como 0.-->
			PASTA_FORM = CLng(0 & Request.Form("PASTA_" & rs("COD_FORMANDO")))
			ENCADERNACOES_EXTRAS_FORM =		formatNumber(Request.Form("ENCADERNACOES_EXTRAS_"&rs("COD_FORMANDO")),0)
			INICIAL_FORM =  				ExtrairNumeros(formatNumber(Request.Form("INICIAL_"&rs("COD_FORMANDO")),0))
			COMISSAO_FORM =  				ExtrairNumeros(Request.Form("COMISSAO_"&rs("COD_FORMANDO")))
			BRINDE_FORM =  					ExtrairNumeros(Request.Form("BRINDE_"&rs("COD_FORMANDO")))
			PROD_AGREG_FORM = 				formatNumber(0&Request.Form("PRODUTO_AGREGADO_"&rs("COD_FORMANDO")),0)
			'VENDEDOR_FORM = 				Request.Form("COD_VENDEDOR_"&rs("COD_FORMANDO")) 'Não é mais necessário
			VENDEDOR_FORM =					0 'Nesta tela não será selecionado o lote para o vendedor.
			MODIFICACOES =					formatNumber(0&Request.Form("MODIFICACOES_"&rs("COD_FORMANDO")),0)
			LOTE_FORM = 					Request.Form("LOTE_"&rs("COD_FORMANDO"))
			
			'ATENÇÃO FAZ A ÚLTIMA VERIFICAÇÃO CASO 2 PESSOAS CADASTRAREM AO MESMO TEMPO
			Dim rsVerifica, ALBUM_VERIFICA, INICIAL_VERIFICA
			
			Set rsVerifica = objConexao.Execute("SELECT ALBUM, INICIAL FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO="&rs("COD_FORMANDO"))
			
				ALBUM_VERIFICA 		= 0&rsVerifica("ALBUM")
				INICIAL_VERIFICA	= 0&rsVerifica("INICIAL")
				
				ALBUM_VERIFICA		= formatNumber(ALBUM_VERIFICA,0)
				INICIAL_VERIFICA	= formatNumber(INICIAL_VERIFICA,0)
							
			rsVerifica.Close()
			Set rsVerifica = Nothing

			IF (ALBUM_VERIFICA = 0 AND INICIAL_VERIFICA = 0) or (ALBUM_DIGITAL_FORM > 0) THEN
			
				 QTD_PREENCHIDA = QTD_PREENCHIDA + 1

				 'verificamos se o item da ficha
				 if COMISSAO_FORM = "2" and brinde = "2" then
				 	aux = ", cod_producao_desconto = NULL "
				 else
				 	aux = " "
				 end if



				 '1) Faz o UPDATE apenas com os que estiverem CORRIGIDOS e depois deleta da tabela temporária
				 objConexao.Execute(" UPDATE TB_FICHAS_IDENTIFICACAO SET "&_
									" ALBUM = "&ALBUM_FORM&", "&_
									" ALBUM_DIGITAL = "&ALBUM_DIGITAL_FORM&", "&_
									" numero_pasta_gerencia = "&PASTA_FORM&", "&_
									" ENCADERNACOES_EXTRAS = "&ENCADERNACOES_EXTRAS_FORM&", "&_
									" INICIAL = "&INICIAL_FORM&", "&_
									" BRINDE = "&BRINDE_FORM&", "&_
									" COMISSAO = "&COMISSAO_FORM&", "&_
									" LOTE = "&LOTE_FORM&", "&_
									" PRODUTO_AGREGADO = "&PROD_AGREG_FORM&_
				 					aux &" WHERE COD_FORMANDO ="&rs("COD_FORMANDO")&" LIMIT 1 ")




				'Insere a 1ª movimentação, entretanto não coloca-se a data de saída somente na impressão dos "Quadradinhos", ou seja o
				'recibo dos vendedores antes de retirarem o álbum
				'/TIPO = 0 => Álbum
				'/////////////////////////////////////////
				'PRODUÇÃO POSTERIOR NÃO ALTERA MODIFICAÇÃO
				'///////////////////////////////////////// 
				
				 '/TIPO = 1 => Ficha
				'/////////////////////////////////////////
				'PRODUÇÃO POSTERIOR NÃO ALTERA MODIFICAÇÃO
				'///////////////////////////////////////// 
				If MODIFICACOES <> 9 Then
				 objConexao.Execute("INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ("&_
									"COD_FORMANDO ,"&_
									"TIPO ,"&_
									"COD_VENDEDOR ,"&_
									"CADASTRADO_POR, "&_
									"COBRAR_DEVOLUCAO "&_
									")"&_
									"VALUES ("&_
									rs("COD_FORMANDO")&",1, "&VENDEDOR_FORM&", '"&Session("COD_USUARIO")&"',2)")					
									
				'/////////////////////////////////				
				'Caso tiver produto agregado e o mesmo for igual a "SIM", busca o produto agregado e inclui na ficha
				End If
				
				'/////////////////////////////////////////
                '/TIPO = 2 => Álbum Digital
                '/////////////////////////////////////////
                If MODIFICACOES <> 9 Then
                    objConexao.Execute("INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ("&_
                                        "COD_FORMANDO ,"&_
                                        "TIPO ,"&_
                                        "COD_VENDEDOR ,"&_
                                        "CADASTRADO_POR, "&_
                                        "COBRAR_DEVOLUCAO "&_
                                        ") VALUES ("&_
                                        rs("COD_FORMANDO")&", 2, "&VENDEDOR_FORM&", '"&Session("COD_USUARIO")&"', 2)")
                End If

				If PROD_AGREG_FORM = 1 then
					Dim rsProdAgre
					
					Set rsProdAgre = objConexao.Execute("SELECT COD_PRODUCAO, COD_PRODUTO,  Sum(QUANTIDADE) as  QUANTIDADE  FROM TB_PRODUCAO_PRODUTOS_AGREGADOS WHERE COD_PRODUCAO="&Request.Form("cod_prod")&" GROUP BY COD_PRODUTO")
					
					If not rsProdAgre.Eof then
						Dim agregadoNumero
						agregadoNumero = 1
						While not rsProdAgre.Eof
						agregadoNumero = agregadoNumero + 1
						objConexao.Execute(	"INSERT INTO TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ( "&_
											"TIPO, "&_
											"COD_PRODUCAO , "&_
											"COD_FORMANDO , "&_
											"COD_PRODUTO , "&_
											"QUANTIDADE, CADASTRADO_POR, COBRAR_DEVOLUCAO) "&_
											"VALUES ("&_
											agregadoNumero&", "&Request.Form("cod_prod")&", "&rs("COD_FORMANDO")&", "&rsProdAgre("COD_PRODUTO")&", "&rsProdAgre("QUANTIDADE")&","&Session("COD_USUARIO")&",3)")
											
				'/////////////////////////////////////////
				'PRODUÇÃO POSTERIOR NÃO ALTERA MODIFICAÇÃO
				'/////////////////////////////////////////							
				 '/TIPO = 3 até 9 => agregados
				 If MODIFICACOES <> 9 Then
				 objConexao.Execute("INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ("&_
									"COD_FORMANDO ,"&_
									"TIPO ,"&_
									"COD_VENDEDOR ,"&_
									"CADASTRADO_POR, "&_
									"COBRAR_DEVOLUCAO "&_
									")"&_
									"VALUES ("&_
									rs("COD_FORMANDO")&","&agregadoNumero&", "&VENDEDOR_FORM&", '"&Session("COD_USUARIO")&"',3)")						
				 End If
					
						rsProdAgre.MoveNext()
						Wend
						
						
					else
						Response.Redirect("dip_producoes.asp?msg=Erro, foi escolhido produto agregado entretanto não há o mesmo cadastrado, registros não gravados.")
					End If
					
					
					rsProdAgre.Close()
					Set rsProdAgre = Nothing
					
				End If

				'2) Limpa o registro temporário da tabela temporária somente do formando selecionado
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_TEMPORARIA WHERE "&_
									"COD_FORMANDO = "&rs("COD_FORMANDO"))
				
				'3) Limpa caso for seleciona brinde "NÃO" e mesmo assim tiver brindes cadastrados, por tanto exclui os brindes
				If BRINDE_FORM = 2 Then
					
					objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_BRINDES WHERE COD_FORMANDO="&rs("COD_FORMANDO"))
									
				End If
			END IF
		  
		  %>
			<tr bgcolor="#FFFF99">
			<td><div align="center"><%=PASTA_FORM%></div></td>
            <td><div align="center"><%=ALBUM_FORM%></div></td>
			<td><div align="center"><%=ALBUM_DIGITAL_FORM%></div></td>
            <td><div align="center"><%=INICIAL_FORM%></div></td>
            <td><div align="center"><%=COMISSAO_FORM%></div></td>
            <td><div align="center"><%=BRINDE_FORM%></div></td>
            <td><div align="center"><%=PROD_AGREG_FORM%></div></td>
            <td><div align="center"><%=LOTE_FORM%></div></td>
          </tr>
          
		  <%
		  End If
		  
		  
		  rs.MoveNext()
		  Wend
	      
		  rs.Close()
		  Set rs = Nothing
		  FecharConexao()
		  
		  If QTD_PREENCHIDA = 0 then
		  	Response.Redirect("dip_producoes.asp?msg=Não foi incluída nenhuma ficha de identificação!")
		  Else
		  	Response.Redirect("dip_producoes.asp?msg="&QTD_PREENCHIDA&" ficha(s) de identificação incluída(s)!")
		  End If
		  
		  %>
		   <tr>
             <td colspan="6"> </td>
		   </tr>
        </table>

		</div>
		  <div align="center">
		    <p>
		      <%else%>
		      <em>Não há formando importados</em>!</p>
		    <p><a href="#">Deseja importar agora?</a></p>
		    <p> </p>
		    <p>	      
		      <%end if%>
		      </p>
		  </div>
		  <p align="center"> </p>		</td>
    </tr>
</table>

 
  
  </td>
    </tr>
</table>

  
  </td>
  </tr>
</table>
</td>
</tr>
</table>
</body>
</html>



