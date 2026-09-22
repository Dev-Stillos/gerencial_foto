<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->


<%
Dim COD_PRODUCAO, comando
COD_PRODUCAO = Request.QueryString("cod_prod")

If Request.QueryString("comando") = "efetuar_alteracao" Then
	
	AbrirConexao()
	
	Dim PRAZO_DIAS_CORRIDOS_INICIAL, COD_VENDEDOR_INICIAL
	
		Set rs=objConexao.Execute("SELECT * FROM VW_TB_PRODUCAO_LOTES WHERE COD_PRODUCAO = "&COD_PRODUCAO)
			If not rs.Eof Then
				
				While not rs.Eof
					
					PRAZO_DIAS_CORRIDOS_INICIAL	= Request.Form("PRAZO_DIAS_CORRIDOS_INICIAL_"&rs("LOTE"))
					COD_VENDEDOR_INICIAL		= formatNumber(0&Request.Form("COD_VENDEDOR_INICIAL_"&rs("LOTE")),0)
					
					If int(COD_VENDEDOR_INICIAL) > 0 Then
					
					objConexao.Execute("UPDATE TB_PRODUCAO_LOTES SET COD_VENDEDOR_INICIAL = "&sqlNumero(COD_VENDEDOR_INICIAL)&",PRAZO_DIAS_CORRIDOS_INICIAL = '"&PRAZO_DIAS_CORRIDOS_INICIAL&"' WHERE COD_PRODUCAO ="&COD_PRODUCAO&" AND LOTE = "&rs("LOTE")&" AND (COD_VENDEDOR_INICIAL IS NULL or COD_VENDEDOR_INICIAL = 0)LIMIT 1")
					
					End If
					
				rs.MoveNext()
				Wend
	
			End If
		rs.Close
		Set rs=Nothing
		
		Response.Redirect("dip_producoes_fichas_lotes.asp?cod_prod="&COD_PRODUCAO&"&msg=Lotes separados para vendedores com sucesso!")
	
	FecharConexao()


ElseIf Request.QueryString("comando") = "excluir_de_lote" Then
	
	AbrirConexao()
	
		PRAZO_DIAS_CORRIDOS_INICIAL	= Request.Form("PRAZO_DIAS_CORRIDOS_INICIAL_"&Request.QueryString("lote"))
		COD_VENDEDOR_INICIAL		= formatNumber(0&Request.Form("COD_VENDEDOR_INICIAL_"&Request.QueryString("lote")),0)
				
		'Atualiza somente se a data de saída inicial for null.	
		objConexao.Execute("UPDATE TB_PRODUCAO_LOTES SET COD_VENDEDOR_INICIAL = NULL,PRAZO_DIAS_CORRIDOS_INICIAL = 0 WHERE COD_PRODUCAO ="&COD_PRODUCAO&" AND LOTE = "&Request.QueryString("lote")&" AND DATA_SAIDA_INICIAL IS NULL LIMIT 1")
					
		
		Response.Redirect("dip_producoes_fichas_lotes.asp?cod_prod="&COD_PRODUCAO&"&msg=Vendedor excluído do lote com sucesso!")
	
	FecharConexao()


End If

Function selectPadraoVendedores()

	Dim retorno, albuns_atrasados, autoriza_bloqueio, diferenca_dias
	Dim rsVendedor, rs, i, ii, temp, temp2
	Dim vendedores_desbloqueados, vendedores_desbloqueados_producoes, vendedor_liberado_manualmente

	'Solicitação realizada e produções tiveram saída
	'vendedores_desbloqueados_producoes = "394,395,396"
	'vendedores_desbloqueados = "104,4,11,41,10,18,96,17,21,8,98,14,48,105,12,7,23,106,75,47"
	
	'Solicitação dia 20/09/13
	vendedores_desbloqueados_producoes = "397,398,400"
	vendedores_desbloqueados = "1,4,7,8,10,11,12,14,16,17,18,19,21,23,27,37,41,47,48,75,96,97,98,104,105,106,108,109,110,111"
	
	vendedor_liberado_manualmente = false
	data_atual = date()
	diferenca_dias = DateDiff("d",date(),"2013-12-20")
	autoriza_bloqueio = false
	
	If diferenca_dias < 0 Then
		autoriza_bloqueio = true
	End If
	
	AbrirConexao()
	Set rsVendedor = objConexao.Execute("SELECT COD_VENDEDOR, NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR <> 1 AND ATIVADO = 1 AND LOTE <> 1 ORDER BY NOME_VENDEDOR")
		
	retorno = "<option selected='selected' value='0'>Selecione um vendedor para este lote</option>"
	
	While not rsVendedor.Eof
										
		albuns_atrasados = 0
		vendedor_liberado_manualmente = false
									
			Set rs2=objConexao.Execute("SELECT COUNT(*) AS ALBUNS_PENDENTES FROM VW_MOVIMENTACOES_COMPLETA WHERE TIPO <> 1 AND DATA_RETORNO IS NULL AND DIAS_ATRASADOS > 0 AND COD_VENDEDOR = "&sqlTexto(rsVendedor("COD_VENDEDOR")))
											
				If not rs2.Eof Then
					albuns_atrasados = int(0&rs2("ALBUNS_PENDENTES"))/1
				End If
									
			rs2.Close
			Set rs2=Nothing
			
			'///Verifica exceções
			If vendedores_desbloqueados_producoes <> "" And vendedores_desbloqueados <> "" Then
					
					temp = Split(vendedores_desbloqueados_producoes,",")
					For i=0 to Ubound(temp)
						If Cstr(""&temp(i)) = Cstr(""&cod_producao) Then							
							
							temp2=Split(vendedores_desbloqueados,",")
					
							For ii=0 to Ubound(temp2)
								
																
								'If rsVendedor("COD_VENDEDOR") = 34 Then
								'	Response.Write(Cstr(""&temp2(ii))&":"&Cstr(""&rsVendedor("COD_VENDEDOR"))&"<br />")
								'End If
								
								
								If Cstr(""&temp2(ii)) = Cstr(""&rsVendedor("COD_VENDEDOR")) Then	
									vendedor_liberado_manualmente = true
								End If
							Next
					
						End If
					Next
					
			End If
			
			
										
			If albuns_atrasados = 0 or Not autoriza_bloqueio or vendedor_liberado_manualmente Then
			'Disponível para novas produções.
			
				retorno = retorno & "<option value='"&rsVendedor("COD_VENDEDOR")&"'>"&Ucase(TrocarAcento(rsVendedor("NOME_VENDEDOR")))
				
				If vendedor_liberado_manualmente Then
					retorno = retorno & " (LIBERADO MANUALMENTE)"
				End If
				
				retorno = retorno &"</option>"
             
			Else
				
                retorno = retorno & "<option disabled='disabled' value='"&rsVendedor("COD_VENDEDOR")&"'>"&Ucase(TrocarAcento(rsVendedor("NOME_VENDEDOR")))&" (BLOQUEADO QTD.: "&albuns_atrasados&")</option>"
           
			End If
								
			rsVendedor.MoveNext()
			Wend

	FecharConexao()
	selectPadraoVendedores = retorno
	
End Function

Dim selectPadraoVendedoresVariavel
selectPadraoVendedoresVariavel = selectPadraoVendedores()
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css"></link>
<link rel="stylesheet" href="css/modal-message.css" type="text/css">
<title>StillosWEB - Sistema corporativo - Versão 1.2</title>

<script type="text/javascript" src="js/ajax.js"></script>
<script type="text/javascript" src="js/modal-message.js"></script>
<script type="text/javascript" src="js/ajax-dynamic-content.js"></script>
<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>

<!--INICIO VALIDAÇÃO SIM/NÃO-->

</head>

<body style='background-color:#fff !important;'>

<%
Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 1
BTN_IMPRIMIR_LINK = "#"

dim i
for i = 248 to 599
	COD_PRODUCAO = i

	Dim rs9, PRODUCAO_FINALIZADA, DESCRICAO, VERIFICA_FINALIZACAO

	VERIFICA_FINALIZACAO = 1

	AbrirConexao()
	Set rs9 = objConexao.Execute("SELECT PRODUCAO_FINALIZADA, DESCRICAO FROM TB_PRODUCAO WHERE COD_PRODUCAO="&COD_PRODUCAO)

	If not rs9.Eof then
		PRODUCAO_FINALIZADA = rs9("PRODUCAO_FINALIZADA")
		DESCRICAO = rs9("DESCRICAO") 
	End If					

	rs9.Close()
	Set rs9 = Nothing 

	%>
	<h2>DIP - Produções / Separação por Lotes</h2> 
	<div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	<h3 align="center">Lotes da Produção:<br />
	<%=COD_PRODUCAO%> - <%=DESCRICAO%></h3>
<%
Set rs=objConexao.Execute("SELECT * FROM VW_TB_PRODUCAO_LOTES WHERE COD_PRODUCAO = "&COD_PRODUCAO)
If not rs.Eof Then
				
%>
		  <form name="formulario_padrao" action="dip_producoes_fichas_lotes.asp?cod_prod=<%=COD_PRODUCAO%>&comando=efetuar_alteracao" method="post" style="margin:0px;">
          <table width="98%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="14%"><div align="center"><strong>Lote</strong></div></td>
            <td width="13%"><div align="center"><strong>Qtd. de Formandos</strong></div></td>
            <td width="31%"><div align="center"><strong> Vendedor</strong></div></td>
            <td width="14%"><div align="center"><strong>Data de Saída</strong></div></td>
            <td width="15%"><div align="center"><strong>Prazo em dias corridos</strong></div></td>
   
</tr>
            <%While not rs.Eof%>
            <tr>
              <td width="14%"><div align="center" class="textoMaior"><strong><%=rs("LOTE")%></strong></div></td>
              <td width="13%"><div align="center"><%=rs("QTD_FORMANDOS_INICIAL")%></div></td>
              <td width="31%">
              <div align="left">
              <%If int(rs("COD_VENDEDOR_INICIAL")) > 0 Then%>
			  	<%=rs("NOME_VENDEDOR")%>
              	<%If IsNull(rs("DATA_SAIDA_INICIAL")) Then %>
              	  <a href="javascript:confirmaExclusaoVendedor('<%=rs("nome_vendedor")%>',<%=rs("cod_vendedor_inicial")%>,'<%=rs("LOTE")%>');"><img border="0" src="images/icons/icon_excluir.png" /></a>
                <%End If%>
                
              <%Else%>
              		<select name="COD_VENDEDOR_INICIAL_<%=rs("LOTE")%>" style="font-size:12px">
						<%=selectPadraoVendedoresVariavel%>	
					</select>
              <%End If%>
              </div>              </td>
              <td width="14%"><div align="center"><%=ArrumaData(rs("DATA_SAIDA_INICIAL"))%></div></td>
              <td width="15%">
                            <div align="center">
                              <%
			  If int(rs("COD_VENDEDOR_INICIAL")) > 0 Then
			  
			  				If rs("DATA_SAIDA_INICIAL") <> Empty Then
			  %>
                                  <%=rs("PRAZO_DIAS_CORRIDOS_INICIAL")%> (Até <%=DateAdd("d",rs("PRAZO_DIAS_CORRIDOS_INICIAL"),rs("DATA_SAIDA_INICIAL"))%>)
                            <%Else%>
                            
                            <%=rs("PRAZO_DIAS_CORRIDOS_INICIAL")%>
                            
                            <%End If%>      
                                  
                              <%Else%>
                                  <input type="text" value="15" class="inputTamanho1" name="PRAZO_DIAS_CORRIDOS_INICIAL_<%=rs("LOTE")%>" />
<%End If%>
                                          </div></td>
            
            </tr>
			<%
			rs.MoveNext()
			wend
			%>
		
		  
        </table>
        </form>
		<br />
     <strong>   <span class="textoMedio">Observações:<br />
     </span></strong><span class="textoMedio"> Se você selecionar um vendedor para um lote e clicar em avançar, automaticamente todos os formandos deste lote serão da responsabilidade do vendedor.<br />
     Este recibo deverá ser utilizado somente a liberação inicial.<br />
     É possível alterar o vendedor do lote desde que o recibo não seja alterado. </span>
     <div align="center">
		    <p>
		      <%else%>
              <br /><br /><br />
		      <em>Não há lotes lançados para esta produção</em>  
               <br /><br /><br />   
		      <%end if%>
		      </p>
		  </div>
		<% fecharConexao() %>
		</td>
	</tr>
</table> 
<p style="page-break-after: always;">.</p>
<% next %>
  

</body>
</html>