<%

Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
BTN_VOLTAR = 2
BTN_IMPRIMIR = 2
BTN_FONTE = 2
BTN_IMPRIMIR_LINK = "#"


Dim calendario_necessario
calendario_necessario = "NÃO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%
Dim rs, sql, rs2, sql2
Dim sub_atrasados,sub_total,sub_prazo,sub_albuns_atrasados,total_atrasados,total_geral,total_prazo,total_albuns_atrasados
		
AbrirConexao()		  
 
sql = "SELECT TB_PRODUCAO_LOTES_FORMANDOS.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR FROM TB_PRODUCAO_LOTES_FORMANDOS LEFT JOIN TB_VENDEDORES ON TB_PRODUCAO_LOTES_FORMANDOS.COD_VENDEDOR = TB_VENDEDORES.COD_VENDEDOR WHERE TB_PRODUCAO_LOTES_FORMANDOS.STATUS =3 AND TB_VENDEDORES.NOME_VENDEDOR IS NOT NULL GROUP BY TB_PRODUCAO_LOTES_FORMANDOS.COD_VENDEDOR ORDER BY TB_VENDEDORES.NOME_VENDEDOR"
				
Set rs = objConexao.Execute(sql)
		
If not rs.Eof Then
%>
	<h2>Controle de Produ&ccedil;&otilde;es/Lotes</h2>
	<h3 align="center">Controle de Produ&ccedil;&otilde;es - Fichas e &Aacute;lb&uacute;ns n&atilde;o vendidos</h3>
	<table width="750" border="1" align="center" cellpadding="3" style="border-collapse:collapse">
        <tr bgcolor="whitesmoke">
            <td width="40%" rowspan="2"><div align="left" class="MensagemFicha" align="center"><strong>Consultor</strong></div></td>
            <td colspan="3"><div align="center"><strong>&Aacute;lbuns</strong> (entre outros itens)</div></td>
            <td colspan="3"><div align="center"><strong>Fichas sob a respons&aacute;bilidade</strong></div></td>
            <td colspan="3"><div align="center"><strong>Produtos</strong></div></td>
        </tr>
		<tr bgcolor="whitesmoke">
  			<td width="10%"><div align="center"><strong>Em atraso</strong></div></td>
  			<td width="10%"><div align="center"><strong>No prazo</strong></div></td>
  			<td width="10%"><div align="center"><strong>Total</strong></div></td>
            <td width="10%"><div align="center" class="MensagemFicha"><strong>Em atraso</strong></div></td>
            <td width="10%" class="MensagemFicha"><div align="center"><strong>No prazo</strong></div></td>
	        <td width="10%" class="MensagemFicha"><div align="center"><strong>Total</strong></div></td>
	        <td width="10%"><div align="center"><strong>Em atraso</strong></div></td>
  			<td width="10%"><div align="center"><strong>No prazo</strong></div></td>
  			<td width="10%"><div align="center"><strong>Total</strong></div></td>
        </tr>
		<%
		total_albuns_atrasados = 0
		total_albuns_prazo = 0
		total_albuns_total = 0
Dim total_prod_atraso, total_prod_prazo
		total_prod_atraso = 0
		total_prod_prazo = 0
		While not rs.Eof
			'//////////		
			sub_atrasados = 0
			sub_total = 0
			sub_prazo = 0
			sub_albuns_atrasados = 0
			sub_albuns_prazo = 0
			sub_albuns_total = 0
			
			Set rs2 = objConexao.Execute("SELECT SUM(IF(DATEDIFF(DATA_PRAZO,NOW()) < 0,1,0)) as ATRASADOS, COUNT(*) as TOTAL FROM TB_PRODUCAO_LOTES_FORMANDOS WHERE COD_VENDEDOR = '"&rs("COD_VENDEDOR")&"' AND STATUS = 3 GROUP BY COD_VENDEDOR")
					
			If Not rs2.Eof Then
				sub_atrasados = formatNumber(0&rs2("ATRASADOS"),0)
				sub_total = formatNumber(0&rs2("TOTAL"),0)
				sub_prazo = (sub_total-sub_atrasados)/1
				
				total_atrasados =	(total_atrasados + sub_atrasados)/1
				total_geral =		(total_geral + sub_total)/1
				total_prazo =		(total_prazo + sub_prazo)/1
			End If
			
			rs2.Close
			Set rs2 = Nothing
			
			Set rs2 = objConexao.Execute("SELECT COUNT(*) AS ALBUNS_PENDENTES FROM VW_MOVIMENTACOES_COMPLETA WHERE TIPO <> 1 AND DATA_RETORNO IS NULL AND DIAS_ATRASADOS > 0 AND COD_VENDEDOR = "&sqlTexto(rs("COD_VENDEDOR")))
			If Not rs2.Eof Then
				sub_albuns_atrasados = formatNumber(0&rs2("ALBUNS_PENDENTES"),0)/1
				total_albuns_atrasados =(total_albuns_atrasados + sub_albuns_atrasados)/1
			End If
			rs2.Close
			Set rs2 = Nothing
						
			Set rs2 = objConexao.Execute("SELECT COUNT(*) AS ALBUNS_PENDENTES FROM VW_MOVIMENTACOES_COMPLETA WHERE TIPO <> 1 AND DATA_RETORNO IS NULL AND DIAS_ATRASADOS < 0 AND COD_VENDEDOR = "&sqlTexto(rs("COD_VENDEDOR")))	
			If Not rs2.Eof Then
				sub_albuns_prazo = formatNumber(0&rs2("ALBUNS_PENDENTES"),0)/1
				total_albuns_prazo = (total_albuns_prazo + sub_albuns_prazo)/1
			End If
			rs2.Close
			Set rs2 = Nothing
			
			sub_albuns_total = (sub_albuns_prazo+sub_albuns_atrasados)/1
			total_albuns_total = (total_albuns_total+sub_albuns_total)/1

			Dim prod_atraso, prod_prazo
			prod_atraso = 0
			prod_prazo = 0


			'Produtos em atraso
			set rs2 = objConexao.Execute("SELECT SUm(quantidade-quantidade_saida) FROM VW_TB_ESTOQUE_MOVIMENTACOES_VENDEDORES WHERE cod_categoria = 6 and data_devolucao >= NOW() AND TIPO = 'E' AND COD_VENDEDOR = "&sqlNumero(rs("COD_VENDEDOR")))
			If not rs2.Eof then

				prod_atraso = formatNumber(0&rs2(0),0)
				total_prod_atraso = total_prod_atraso + prod_atraso
			else
				prod_atraso = 0
			End if
			rs2.Close
			Set rs2 = Nothing

			'Produtos em Dentro do prazo
			set rs2 = objConexao.Execute("SELECT SUm(quantidade-quantidade_saida) FROM VW_TB_ESTOQUE_MOVIMENTACOES_VENDEDORES WHERE cod_categoria = 6 and data_devolucao < NOW() AND TIPO = 'E' AND COD_VENDEDOR = "&sqlNumero(rs("COD_VENDEDOR")))
			If not rs2.Eof then

				prod_prazo =formatNumber(0&rs2(0),0)
				total_prod_prazo = total_prod_prazo + prod_prazo
			else
				prod_prazo = 0
			End if
			rs2.Close
			Set rs2 = Nothing

			%>
			<tr 
			<%
			Dim corr
			
			if corr = "#EAEAEA" then
			corr = "#FFFFFF"
			else
			corr = "#EAEAEA"
			end if
			%>
			bgcolor="<%=corr%>">
		    <td><div align="left"><a href="dip_controle_producoes_listagem_producoes.asp?COD_VENDEDOR=<%=rs("COD_VENDEDOR")%>"><%=Ucase(rs("NOME_VENDEDOR"))%></a></div></td>
		    <td><div align="center" style="height:15px; background-color:#FA9A9A; padding-top:2px"><span><%=sub_albuns_atrasados%></span></div></td>
		    <td><div align="center" style="height:15px; background-color:#00FF66; padding-top:2px"><span><%=sub_albuns_prazo%></span></div></td>
		    <td><div align="center"><strong><span><%=sub_albuns_total%></span></strong></div></td>
		    <td><div align="center" style="height:15px; background-color:#FA9A9A; padding-top:2px"><%=sub_atrasados%></div></td>
		    <td><div align="center" style="height:15px; background-color:#00FF66; padding-top:2px"><%=sub_prazo%></div></td>
		    <td><div align="center"><strong><%=sub_total%></strong></div></td>
			<td><div align="center" style="height:15px; background-color:#FA9A9A; padding-top:2px"><%=prod_atraso%></div></td>
		    <td><div align="center" style="height:15px; background-color:#00FF66; padding-top:2px"><%=prod_prazo%></div></td>
		    <td><div align="center"><strong><%=formatNumber(prod_atraso+prod_prazo,0)%></strong></div></td>
		</tr>
		<%
		rs.MoveNext()
		Wend			
		%>  
		<tr>
			<td>&nbsp;</td>
			<td><div align="center"><strong><%=total_albuns_atrasados%></strong></div></td>
			<td><div align="center"><strong><%=total_albuns_prazo%></strong></div></td>
			<td><div align="center" class="textoMedio"><strong><%=total_albuns_total%></strong></div></td>
			<td><div align="center"><strong><%=total_atrasados%></strong></div></td>
			<td><div align="center"><strong><%=total_prazo%></strong></div></td>
			<td><div align="center" class="textoMedio"><strong><%=total_geral%></strong></div></td>
			<td><div align="center"><strong><%=total_prod_atraso%></strong></div></td>
			<td><div align="center"><strong><%=total_prod_prazo%></strong></div></td>
			<td><div align="center" class="textoMedio"><strong><%=total_prod_atraso+total_prod_prazo%></strong></div></td>
		</tr>
    </table>
    <%Else%>
	    <br /><br /><br /><br />
	    <div align="center"><strong>Não há vendedores com fichas em posse.</strong></div>
	    <br /><br /><br /><br />
    <%
	End If
rs.Close
Set rs=Nothing
FecharConexao()
%>
  
<!-- #include file="files/includes/miolo-inferior.asp" -->