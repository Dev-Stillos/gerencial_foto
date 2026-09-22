<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" --> 

<%
Response.Buffer = True
%>

<%

abrirConexao()

dim cod_vendedor, cod_producao
cod_producao = request.queryString("p")
cod_vendedor = request.queryString("v")

' Monta condicionais apenas se houver valor
dim condProd, condProd2, condVend, condVend2, baseWhere, baseWhere2
condProd = ""
condProd2 = ""
condVend = ""
condVend2 = ""
baseWhere = " WHERE 1=1"
baseWhere2 = " where 1=1"

If Len(Trim(cod_producao)) > 0 Then
    If IsNumeric(Trim(cod_producao)) Then
        condProd = " AND P.COD_PRODUCAO = " & sqlNumero(cod_producao)
        condProd2 = " AND FI.COD_PRODUCAO = " & sqlNumero(cod_producao)
    End If
End If

If Len(Trim(cod_vendedor)) > 0 Then
    If IsNumeric(Trim(cod_vendedor)) Then
        condVend = " AND VND.COD_VENDEDOR = " & sqlNumero(cod_vendedor)
        condVend2 = " AND FIH.COD_VENDEDOR = " & sqlNumero(cod_vendedor)
    End If
End If


function statusPedidoFormando(cod_formando, conexao_aberta)

	if not conexao_aberta then
		abrirConexao()
	end if

	Dim rs, retorno

	Set rs = ObjConexao.Execute("SELECT * FROM TB_PEDIDOS_FORMANDOS WHERE cod_formando = "&sqlNumero(cod_formando))
	if not rs.Eof then
		retorno = "Sim"
	else
		retorno = "N�o"
	end if
	rs.close
	set rs = nothing

	if not conexao_aberta then
		feharConexao()
	end if 

statusPedidoFormando = retorno
end function

'Buscamos informa��es gerais

dim nome_vendedor, descricao_producao, numero_lote, qtd_lote

Set rs = ObjConexao.Execute("select P.DESCRICAO, PL.LOTE, PL.QTD_FORMANDOS_INICIAL, VND.NOME_VENDEDOR from TB_PRODUCAO as P inner join TB_PRODUCAO_LOTES as PL on P.COD_PRODUCAO = PL.COD_PRODUCAO inner join TB_VENDEDORES as VND on PL.COD_VENDEDOR_INICIAL = VND.COD_VENDEDOR" & baseWhere & condVend & condProd)
if not rs.Eof then
	nome_vendedor = rs("NOME_VENDEDOR")
	descricao_producao = rs("descricao")
	numero_lote = rs("LOTE")
	qtd_lote = rs("QTD_FORMANDOS_INICIAL")
else
	response.write("N�o encontramos as informa��es solicitadas")
	response.end()
end if
rs.close
set rs = nothing


%>

<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
		<title>Rel�torio de observa��es feitas por vendedores por Produ��o</title>
		<style>
			body{
				font-family: arial;
				font-size: 10px;
			}
			table {
				width: 99%;
				font-size: 12px;
			}
			.collapse{
				border: 1px solid #000;
				border-collapse: collapse;
			}
		</style>
	</head>
	<body>

	<h2 align="center">Rel�torio de observa��es feitas por vendedores por Produ��o</h2>
		<table>
			<tr>
				<th width="100px" align="left">Vendedor:</th>
				<td align="left"><%=nome_vendedor%></td>
			</tr>
			<tr>
				<th align="left">Produ��o:</th>
				<td align="left"><%=cod_producao%> - <%=descricao_producao%> / lote <%=numero_lote%> (<%=qtd_lote%> fichas)</td>
			</tr>
		</table>


		<%
		cod_formando = ""
		count = 0
			Set rs = ObjConexao.Execute( _
    "SELECT FIH.*, " & _
    "FI.NOME_FORMANDO, " & _
    "U.NOME, " & _
    "VND.NOME_VENDEDOR " & _
    "FROM TB_FICHAS_IDENTIFICACAO_HISTORICO FIH " & _
    "INNER JOIN TB_FICHAS_IDENTIFICACAO FI ON FI.COD_FORMANDO = FIH.COD_FORMANDO " & _
    "LEFT JOIN TB_ADM_USUARIOS U ON U.COD_USUARIO = FIH.GERADO_POR " & _
    "LEFT JOIN TB_VENDEDORES VND ON VND.COD_VENDEDOR = FIH.COD_VENDEDOR " & _
    "WHERE 1=1 " & _
    condVend2 & _
    condProd2 & _
    " ORDER BY FIH.COD_FORMANDO, FIH.ID DESC")

	If Not rs.EOF Then

    While Not rs.EOF

        count = count + 1

        If count Mod 100 = 0 Then
            Response.Flush
        End If

        If cod_formando <> rs("cod_formando") Then

            cod_formando = rs("cod_formando")

            If count > 1 Then
%>
                    </table>
                </td>
            </tr>
        </table>
<%
            End If
%>

<br>

<table border="1" class="collapse" cellpadding="5">
				<tr>
					<th align="left">Formando: <%=cod_formando%> - <%=rs("nome_formando")%> </th>
					<td width="70px" align="rigth"><b>Vendido:</b><%=statusPedidoFormando(cod_formando, true)%></td>
				</tr>
				<tr>
					<td align="center" colspan="2">
						<table border="1" class="collapse" cellpadding="3">
						<tr>
							<th width="75%">Mensagem</th>
							<th>Data</th>
						</tr>
				<%
	
				end if
			
			%>

<tr>
    <td><%=rs("mensagem")%></td>
    <td align="center"><%=arrumaData(rs("data_cadastro"))%></td>
</tr>

<%

        rs.MoveNext

    Wend

End If

rs.Close
Set rs = Nothing

If count > 0 Then
%>
                </table>
            </td>
        </tr>
    </table>
<%
End If
%>
	</body>
</html>