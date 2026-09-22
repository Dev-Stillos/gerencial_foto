<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!--
'*** Instruções
'1) Grava somente a se tiver sido lançado a ficha do formandos
'2) Após a gravação do cód. de identificação verifica o que o formando possui em aberto
'3) Verifica se o formando já não está na responsábilidade do vendedor que está retirando
-->
<%
Dim comando, COD_BARRAS, ERROR_LINHA, pronto, COD_VENDEDOR, CONT_ITENS
Dim COD_FORMANDO_COD_BARRAS, TIPO_COD_BARRAS
Dim redirecionar, redirecionar_msg
Dim COUNT_FICHAS, COUNT_AGREGADOS, COUNT_ALBUNS


redirecionar = "dip_repasse.asp"

'Contadores e verificador de erros
'///
CONT_ITENS 		= 0
ERROR_LINHA 	= 1
COUNT_FICHAS	= 0
COUNT_AGREGADOS	= 0
COUNT_ALBUNS	= 0
'///

comando = Request.QueryString("comando")
COD_BARRAS = ""&Request.Form("COD_BARRAS")
COD_BARRAS = Replace(COD_BARRAS," ","")
COD_VENDEDOR = Request.QueryString("COD_VENDEDOR")
Dim NOME_VENDEDOR

AbrirConexao()
	Set rs = objConexao.Execute("SELECT NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR = "&COD_VENDEDOR)
			
			If not rs.Eof Then
				
				NOME_VENDEDOR = rs("NOME_VENDEDOR")
				
			End If
						
	rs.Close
	Set rs = Nothing
FecharConexao()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="all"></link>
<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
        <style type="text/css" media="all">
<!--
.style1 {
	font-size: 7px;
	font-weight: bold;
}
body {
	background-color: #FFFFFF;
}
.style2 {font-size: 11px}
-->
        </style>
</head>
<body onload="window.print()">
<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
  <td>
   <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="top">
<h2 style="color:#000000;">Transferência de Itens</h2>

<table width="100%" border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td width="256" class="textoMedio">Data da transfer&ecirc;ncia:</td>
                  <td width="801"><strong class="textoMedio"><%=ArrumaData(Now())%></strong></td>
                </tr>
                <tr>
                  <td class="textoMedio">Representante Comercial:</td>
                  <td class="textoMedio"><strong><%=NOME_VENDEDOR%></strong></td>
                </tr>
              </table>
          <br /> <br />
      <%

AbrirConexao()

sql = "SELECT * FROM TB_FICHAS_IDENTIFICACAO_DEVOLUCAO_TEMPORARIA WHERE COD_USUARIO="&Session("COD_USUARIO")&" GROUP BY COD_ETIQUETA"
Set rs=objConexao.Execute(sql)

If not rs.Eof Then
%>
			<table width="98%" border="0" align="center" 	cellpadding="2" style="border-collapse:collapse">
			<tr bgcolor="whitesmoke">
       
            <td width="5%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style1">C&oacute;d. Formando </div></td>
            <td width="6%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style1">&Aacute;lbum</div></td>
            <td width="12%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style1">Nome</div></td>
            <td width="11%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style1">Institui&ccedil;&atilde;o de Ensino</div></td>
            <td width="13%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style1">Curso</div></td>
            <td width="6%" class="style1" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center">Conclus&atilde;o</div></td>
            <td width="3%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style1">Repasse*</div></td>
            <td width="4%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><div align="center" class="style1">Prazo</div></td>
            <td width="40%" style="border-bottom:solid 2px #000; border-top:solid 2px #000"><table width="100%" border="0" cellpadding="0" cellspacing="0">

              <tr>
                <td width="36%" class="style1"><div align="center">Item transferido</div></td>
                <td width="64%" class="style1"><div align="center">&Uacute;ltmo repres. comercial</div>
                    <div align="center"></div></td>
              </tr>
            </table></td>
			</tr>
			
			<%
			Dim COD_FORMANDO, NOME_FORMANDO, NOME_CURSO, SEMESTRE_CONCLUSAO, ANO_CONCLUSAO, ENCADERNACOES_EXTRAS, ALBUM, DESCRICAO, COD_CONTRATO
			Dim TRANS_DATA_CONCLUSAO, TRANS_DATA_VIRA_REPASSE, TRANS_DATA_ATUAL, DIFERENCA_DATA
			Dim TIPO, ITEM_TRANSFERENCIA, DATA_RETORNO
			Dim COD_VENDEDOR_FICHA, DATA_RETORNO_FICHA
			Dim PRAZO, PRAZO_NECESSARIO
			
			
			While not rs.Eof
				
				COD_FORMANDO = Mid(rs("COD_ETIQUETA"),1,Len(rs("COD_ETIQUETA"))-1)
								
				'/////////////////////////
				'Formandos que possuem lotes é necessário informar o prazo para prestação de contas
				PRAZO = formatNumber(0&rs("PRAZO"),0)
				PRAZO_NECESSARIO = Cstr(""&rs("PRAZO_NECESSARIO"))
				
				If int(PRAZO) = 0 Then
					PRAZO = 15
				End If
				
				if Cstr(PRAZO_NECESSARIO) = "S" Then
					
						objConexao.Execute("UPDATE TB_PRODUCAO_LOTES_FORMANDOS SET COD_VENDEDOR = '"&COD_VENDEDOR&"', DATA_SAIDA = "&sqlData(NOW)&" ,DATA_PRAZO = "&sqlData(DateAdd("d",PRAZO,Now()))&"  WHERE COD_FORMANDO = "&COD_FORMANDO&" limit 1")
				
				End If
				'/////////////////////////
				
				
				
				
				'1 - Verifica qual é o nome formando e nome do curso
				'///
				Set rs2 = objConexao.Execute("SELECT NOME_FORMANDO, NOME_CURSO, SEMESTRE_CONCLUSAO, ANO_CONCLUSAO, ENCADERNACOES_EXTRAS, ALBUM, COD_CONTRATO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO = "&COD_FORMANDO)	
					
						If not rs2.Eof Then
							ENCADERNACOES_EXTRAS= formatNumber(0&rs2("ENCADERNACOES_EXTRAS"),0)
							NOME_FORMANDO		= rs2("NOME_FORMANDO")
							NOME_CURSO			= rs2("NOME_CURSO")
							SEMESTRE_CONCLUSAO	= rs2("SEMESTRE_CONCLUSAO")
							ANO_CONCLUSAO		= rs2("ANO_CONCLUSAO")
							ALBUM				= rs2("ALBUM")
							COD_CONTRATO		= rs2("COD_CONTRATO")
							
							
							Set rs3=objConexao.Execute("Select DESCRICAO from TB_ENTIDADES where COD_CONTRATO = '"&COD_CONTRATO&"'")
									
									If not rs3.Eof Then
										DESCRICAO = rs3("DESCRICAO")
									End If
							
							rs3.Close
							Set rs3=Nothing							
							
							'Verifica repasse
							'////////////
							'***Baseado na data atual
							
							'Ex. 1 Repasse: Conclusão faculdade DEZEMBRO DE 2005 virará repasse em DEZEMBRO DE 2008.
							'Ex. 2 Repasse: Conclusão faculdade JULHO DE 2006 virará repasse em JULHO DE 2009.
					
							If SEMESTRE_CONCLUSAO = "J" Then
								SEMESTRE_CONCLUSAO = 7
							Elseif SEMESTRE_CONCLUSAO = "D" Then
								SEMESTRE_CONCLUSAO = 12
							End If
					
							TRANS_DATA_CONCLUSAO	= formatDateTime(ANO_CONCLUSAO&"-"&SEMESTRE_CONCLUSAO&"-1")
							TRANS_DATA_VIRA_REPASSE	= DateAdd("yyyy",3,TRANS_DATA_CONCLUSAO)
							TRANS_DATA_ATUAL 		= formatDateTime(Year(Now)&"-"&Month(Now)&"-1")
									
							'Para ser repasse tem de ser no mínimo 3 anos de diferença
							'Verifica se já ha tempo mínimo para virar repasse
							DIFERENCA_DATA = DateDiff("d",TRANS_DATA_VIRA_REPASSE,TRANS_DATA_ATUAL)
							'////////////
							
							
						End If
					
				rs2.Close
				Set rs2 = Nothing
				'///
			CONT_ITENS = CONT_ITENS + 1
			%>
			<tr>
            <td style="border-bottom:solid 1px #333333; border-right:solid 0.25px #333333"><div align="center" class="style1"><%=digito(COD_FORMANDO)%></div></td>
            <td style="border-bottom:solid 1px #333333; border-right:solid 0.25px #333333"><div class="style1">
              <div align="center"><%=ALBUM%></div>
            </div></td>
            <td style="border-bottom:solid 1px #333333; border-right:solid 0.25px #333333"><div class="style1"><%=NOME_FORMANDO%></div></td>
            <td style="border-bottom:solid 1px #333333; border-right:solid 0.25px #333333"><div class="style1"><%=DESCRICAO%></div></td>
            <td style="border-bottom:solid 1px #333333; border-right:solid 0.25px #333333"><div class="style1"><%=NOME_CURSO%></div></td>
            <td style="border-bottom:solid 1px #333333; border-right:solid 0.25px #333333"><div class="style1" align="center"><%=Left(Ucase(MonthName(SEMESTRE_CONCLUSAO)),3)%> / <%=ANO_CONCLUSAO%></div></td>
            <td style="border-bottom:solid 1px #333333; border-right:solid 0.25px #333333"><div align="center" class="style1">
            <%If DIFERENCA_DATA => 0 Then%>SIM<%Else%>N&Atilde;O<%End If%>
            </div></td>
            <td style="border-bottom:solid 1px #333333; border-right:solid 0.25px #333333">
            <div align="center" class="style1">
            <%
			If Cstr(PRAZO_NECESSARIO) = "S" Then
				Response.Write(PRAZO&" dia(s)")
			Else
				Response.Write("Padrão")
			End If
			%>
            </div></td>
            <td style="border-bottom:solid 1px #333333; border-right:solid 0.25px #333333"><table width="100%" border="0" cellpadding="2" cellspacing="0" style="border-collapse:collapse">
              <%
			'Transfer&ecirc;ncia 
			'/////////////////
				'1 Verifica quais o produtos que o formando possui
				'2 Verifica se h&aacute; algum produto em aberto, se tiver ser&aacute; transferido para o novo vendedor
				
				'Zerando variaveis
				ITEM_TRANSFERENCIA = ""
				COD_VENDEDOR_FICHA = 0
				DATA_RETORNO_FICHA = ""
				
				
				sql = "SELECT TIPO FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE  COD_FORMANDO = "&COD_FORMANDO&" GROUP BY TIPO ORDER BY TIPO"
				Set rs2 = objConexao.Execute(sql)
						

						While not rs2.Eof

							'Verifica dos tipo quais est&atilde;o em aberto	
							'////
								sql =	" SELECT TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_MOVIMENTACAO, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_SAIDA, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.DATA_RETORNO, TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR  FROM  "&_
										" TB_VENDEDORES INNER JOIN TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES ON (TB_VENDEDORES.COD_VENDEDOR = TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_VENDEDOR) "&_
										"  WHERE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.TIPO = "&rs2("TIPO")&" AND TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES.COD_FORMANDO = "&COD_FORMANDO&_
										"  ORDER BY COD_MOVIMENTACAO DESC LIMIT 1"
								Set rs3=objConexao.Execute(sql)
								
									If not rs3.Eof Then
									
									 
									'Exibe o item e &uacute;ltimo respos&aacute;vel e verifica qual &eacute; o tipo
										DATA_RETORNO = ""&rs3("DATA_RETORNO")
										
										TIPO = formatNumber(0&rs2("TIPO"),0)
										
										If TIPO = 0 and ENCADERNACOES_EXTRAS = 0 Then
										'&Aacute;lbum
											ITEM_TRANSFERENCIA = "&Aacute;LBUM"
											'Conta se não tiver sido devolvido
											'//
											If DATA_RETORNO = "" Then
												COUNT_ALBUNS = COUNT_ALBUNS + 1
											End If
											'//
										Elseif TIPO = 0 and ENCADERNACOES_EXTRAS = 1 Then
										'&Aacute;lbum Duplo
											ITEM_TRANSFERENCIA = "&Aacute;LBUM DUPLO"
											'Conta se não tiver sido devolvido
											'//
											If DATA_RETORNO = "" Then
												COUNT_ALBUNS = COUNT_ALBUNS + 1
											End If
											'//
										Elseif TIPO = 1 Then
										'Ficha de Identifica&ccedil;&atilde;o
											ITEM_TRANSFERENCIA = "FICHA"
											COD_VENDEDOR_FICHA = formatNumber(0&rs3("COD_VENDEDOR"),0)
											DATA_RETORNO_FICHA = rs3("DATA_RETORNO")
											'Sempre conta
											COUNT_FICHAS = COUNT_FICHAS+1
										Else
										'Produto agregado
										
											'Conta se não tiver sido devolvido
											'//
											If DATA_RETORNO = "" Then
												COUNT_AGREGADOS = COUNT_AGREGADOS+1
											End If
											'//											
										
											sql = 	"SELECT TB_PRODUTOS.NOME_PRODUTO "&_
													"FROM TB_PRODUTOS INNER JOIN TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS ON TB_PRODUTOS.COD_PRODUTO = "&_
													"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_PRODUTO WHERE "&_
													"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.COD_FORMANDO = "&COD_FORMANDO&" AND "&_
													"TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS.TIPO = "&TIPO&" LIMIT 1"
				
											Set rsProdAgrePend = objConexao.Execute(sql)
											If not rsProdAgrePend.Eof Then
											
												ITEM_TRANSFERENCIA = rsProdAgrePend("NOME_PRODUTO")					
											
											End If
											rsProdAgrePend.Close
											Set rsProdAgrePend = Nothing
										
											
										End If
									%>
              <% If int(TIPO) = 1 or DATA_RETORNO = "" Then
			  'So mostra itens que serão transferidos ou seja obrigatoriamente FICHAS + (ÁLBUNS + PROD. AGREGADOS que estiverem em ABERTO  para o vendedor antigo)
			  
			  '//////////////////////////////
			  'FAZ O INSERT DA TRANSFERÊNCIA
			  
			  		'A) SE ESTIVER EM ABERTO
						'1- Da baixa na ficha no registro anterior
						'2- Da um insert para o novo vendedor
					
					'B) SE ESTIVER DEVOLVIDA
						'1- Da um insert para o novo vendedor
						
					'C) OUTROS ITENS COMO ÁLBUM E PROD. AGREGADOS QUE ESTÃO EM ABERTO
					'   por tanto itens outros itens que estiverem no estoque permaneceram da mesma forma
					
						'1- Da baixa na item no registro anterior
						'2- Da um insert para o novo vendedor
			  
			  		If int(TIPO) = 1 and DATA_RETORNO = "" Then
					'A
						'1
						sql =	" UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET DATA_RETORNO = NOW( ), TRANSFERIDO = '1' "&_
								" WHERE COD_MOVIMENTACAO = "&rs3("COD_MOVIMENTACAO")&" LIMIT 1"
						
						objConexao.Execute(sql)
						
						'2
						sql =	" INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES (COD_FORMANDO, TIPO, COD_VENDEDOR, DATA_SAIDA, CADASTRADO_POR, TRANSFERIDO) "&_ 
								" VALUES ("&COD_FORMANDO&","&TIPO&","&COD_VENDEDOR&",NOW(),"&Session("COD_USUARIO")&",2)"
											
						objConexao.Execute(sql)
						
					Elseif int(TIPO) = 1 and DATA_RETORNO <> "" Then
					'B
						'1
						sql =	" INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES (COD_FORMANDO, TIPO, COD_VENDEDOR, DATA_SAIDA, CADASTRADO_POR, TRANSFERIDO) "&_ 
								" VALUES ("&COD_FORMANDO&","&TIPO&","&COD_VENDEDOR&",NOW(),"&Session("COD_USUARIO")&",2)"
											
						objConexao.Execute(sql)
						
					Elseif int(TIPO) <> 1 and DATA_RETORNO = "" Then
					'C
						'1
						sql =	" UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET DATA_RETORNO = NOW( ), TRANSFERIDO = '1' "&_
								" WHERE COD_MOVIMENTACAO = "&rs3("COD_MOVIMENTACAO")&" LIMIT 1"
											
						objConexao.Execute(sql)
						
						'2
						sql =	" INSERT INTO TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES (COD_FORMANDO, TIPO, COD_VENDEDOR, DATA_SAIDA, CADASTRADO_POR, TRANSFERIDO) "&_ 
								" VALUES ("&COD_FORMANDO&","&TIPO&","&COD_VENDEDOR&",NOW(),"&Session("COD_USUARIO")&",2)"
														
						objConexao.Execute(sql)						
						
						'PARA POR AQUI, OU SEJA OUTROS ITENS QUE ESTIVEREM DEVOLVIDOS NÃO SERÃO TRANSFERIDOS
						
					End If
			  		
			  '//////////////////////////////
			  
			  %>
              <tr>
                <td class="style1"><%=ITEM_TRANSFERENCIA%></td>
                <td width="64%" class="style1"><%=rs3("NOME_VENDEDOR")%></td>
              </tr>
              <%
			  End If
			  
									End If							
								Set rs3=Nothing
							'////	
							
							rs2.MoveNext()
							Wend				
										
				rs2.Close
				Set rs2 = Nothing
					
			'/////////////////		
			%>
            </table></td>
   	                </tr>
			<%
			rs.MoveNext()
			Wend
			%>                               
                                <tr>
          					  <td colspan="9" class="textoMedio" style="padding:10px"><div align="left">
          					    <table width="700" border="0" cellpadding="4">

                                  <tr>
                                    <td width="142" style="border-right: 1px #666666 solid"><strong class="textoMenor">Quantidade de Fichas: <%=COUNT_FICHAS%></strong></td>
                                    <td width="152" style="border-right: 1px #666666 solid"><span class="textoMenor"><strong>**Quantidade de &Aacute;lbuns: <%=COUNT_ALBUNS%></strong></span></td>
                                    <td width="374"><strong class="textoMenor">**Quantidade Produtos Agregados: <%=COUNT_AGREGADOS%></strong></td>
                                  </tr>
                                  <tr>
                                    <td colspan="3"><strong class="textoMenor style2">Total Itens Transferidos: <%=(COUNT_FICHAS+COUNT_ALBUNS+COUNT_AGREGADOS)/1%></strong></td>
                                  </tr>
                                  <tr>
                                    <td colspan="3"><strong><span class="textoMenor">Observa&ccedil;&otilde;es:</span></strong><span class="textoMenor"><br />
                                      *Repasse calculado com base na data atual<br />
                                      **Item que estava em posse do antigo repres. comercial e n&atilde;o est&aacute; no estoque. A partir de agora passa ser da responsabilidade do novo repres. comercial.</span></td>
                                  </tr>
                                </table>
          					    <br />
          				<table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
           <tr>
             <td style="border-bottom:solid 1px #000000">
               <p>&nbsp;</p></td>
           </tr>
           <tr>
             <td><div align="center"><span class="textoMedio"><span class="textoMenor">REPRESENTANTE COMERCIAL</span><br />
                   <span class="textoMenor">  <strong><%=NOME_VENDEDOR%></strong></span></span></div></td>
           </tr>
         </table>
          					  <br />
                              		 	      <p>&nbsp;</p>
		   <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
           <tr>
             <td style="border-bottom:solid 1px #000000">
               <p>&nbsp;</p></td>
           </tr>
           <tr>
             <td><div align="center"><span class="textoMenor">TRANSFERIDO POR <strong><br />
               <%=Ucase(Session("NOME"))%></strong></span></div></td>
           </tr>
         </table>
          					  </div></td>
       					    </tr>
						</table>			             
<%
End If

rs.Close
Set rs=Nothing

FecharConexao()
%></td>
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