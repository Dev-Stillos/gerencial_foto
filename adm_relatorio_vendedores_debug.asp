<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%

dim campo_base
'campo_base = "DATA_CORRIGIDO_DEP_FOTO"
campo_base = "DATA_CADASTRO"

abrirConexao()

'resultados(i,0) = Qtd. de Álbs vendidos de todos os lotes retirados
'resultados(i,1) = Valor Arrecadado de todos os lotes retirados
'resultados(i,2) = Aproveitamento de albuns vendidos 
'resultados(i,3) = Qtd. Total de Álbs vendidos
'resultados(i,4) = Valor total arrecadado
'resultados(i,5) = qtd retirada

dim inicio, fim, rs, count

inicio = request.querystring("a")
fim = request.querystring("b")
com_valores = request.querystring("c")

if len(inicio) = 10 then
	inicio = inicio & " 00:00:00"
end if

if len(fim) = 10 then
	fim = fim & " 23:59:59"
end if


dim resultados(100,10)
dim matriz_ordena(100,1)
dim vendedores(100,3) 'armazena todos os vendedores do relatorio'

sub ordenaMatriz()
dim aux0, aux1, i, j

for i = 0 to ubound(matriz_ordena,1)
	for j = 0 to ubound(matriz_ordena,1)
        if cdbl(matriz_ordena(i,1)) > cdbl(matriz_ordena(j,1)) then
			'aqui acontece a troca, do maior cara  vaia para a direita e o menor para a esquerda
            aux0 = matriz_ordena(i,0)
            aux1 = matriz_ordena(i,1)
            
            matriz_ordena(i,0) = matriz_ordena(j,0)
            matriz_ordena(i,1) = matriz_ordena(j,1)
            
            matriz_ordena(j,0) = aux0
            matriz_ordena(j,1) = aux1
        end if
    next
next

end sub

sub limpaMatrizOrdena()
	Dim i :	i = 0

	for i = 0 to ubound(matriz_ordena,1)
		matriz_ordena(i,0) = ""
		matriz_ordena(i,1) = 0
	next
end sub

sub exibeMatriz()
	Dim i :	i = 0

	for i = 0 to ubound(matriz_ordena,1)
		response.write matriz_ordena(i,0) & " : " & matriz_ordena(i,1) & "<br>"
	next
end sub

sub exibeMatrizResultado()
	Dim i, x

	for i = 0 to ubound(resultados,1)
		for x = 0 to 10
			response.write resultados(i,x) & " : "
		next
		response.write "<br>"
	next
end sub



sub copiaMatriz(comValores, posicao)

dim i 

if comValores then
	for i= 0 to ubound(matriz_ordena,1)
		resultados(i,posicao)  = matriz_ordena(i,0) &" ("& matriz_ordena(i,1) &")"
	next
else
	for i= 0 to ubound(matriz_ordena,1)
		'resultados(i,posicao)  = matriz_ordena(i,0)
		resultados(i,posicao)  = matriz_ordena(i,0) &" " 
	next
end if

end sub


count = 0


'PRIMEIRO BUSCAMOS TODOS OS VENDEDORES Q RETIRARAM PRODUCAO NO PERIODO
sql = "SELECT TB_VENDEDORES.COD_VENDEDOR, NOME_VENDEDOR, SUM(QTD_FORMANDOS_INICIAL) AS QTD_RETIRADO FROM TB_PRODUCAO_LOTES INNER JOIN TB_VENDEDORES ON TB_VENDEDORES.COD_VENDEDOR = TB_PRODUCAO_LOTES.COD_VENDEDOR_INICIAL WHERE  (DATA_SAIDA_INICIAL >= "&sqlData(inicio)&" and  DATA_SAIDA_INICIAL <= "&sqlData(fim)&"  AND TB_VENDEDORES.COD_VENDEDOR <> 27) GROUP BY NOME_VENDEDOR union all SELECT COD_VENDEDOR, NOME_VENDEDOR, 0 as QTD_RETIRADO FROM TB_VENDEDORES where Ativado = 1 and cod_grupo_vendedores = 0 and COD_VENDEDOR not in ( SELECT COD_VENDEDOR FROM TB_PRODUCAO_LOTES INNER JOIN TB_VENDEDORES ON TB_VENDEDORES.COD_VENDEDOR = TB_PRODUCAO_LOTES.COD_VENDEDOR_INICIAL WHERE  (DATA_SAIDA_INICIAL >= "&sqlData(inicio)&" and  DATA_SAIDA_INICIAL <= "&sqlData(fim)&"  AND TB_VENDEDORES.COD_VENDEDOR <> 27) GROUP BY NOME_VENDEDOR) "
response.write sql
Set rs = objConexao.Execute(sql)
while not rs.Eof 
	
	vendedores(count,0) = rs("COD_VENDEDOR")
	vendedores(count,1) = rs("NOME_VENDEDOR")
	vendedores(count,2) = rs("QTD_RETIRADO")

	count = count + 1
rs.movenext
wend 
rs.close
set rs = nothing


call limpaMatrizOrdena()

'BUSCAMOS O QUANTO CADA VENDEDOR RECEBEU no periodo de acordo com os lotes retirados no periodo
for count = 0 to ubound(vendedores,1)
	if int("0"&vendedores(count,2)) > 0 then
		sql = "SELECT (SUM(IFNULL(PAGAMENTO_EM_DINHEIRO,0)) + SUM(IFNULL(XTOTAL_CHEQUE,0)) + SUM(IFNULL(XTOTAL_CARTAO,0))+ SUM(IFNULL(XTOTAL_BOLETO,0))) AS VALOR_VENDIDO FROM VW_EXTRATO_TOTAL_RECEBIDO_COMPLETO WHERE DATA_CORRIGIDO_DEP_FOTO is not null AND "&campo_base&" BETWEEN "&sqlData(inicio)&" AND "&sqlData(fim)&" AND COD_PEDIDO IN (SELECT COD_PEDIDO FROM TB_PEDIDOS_FORMANDOS WHERE COD_FORMANDO IN (SELECT COD_FORMANDO FROM TB_PRODUCAO_LOTES_FORMANDOS WHERE DATA_SAIDA BETWEEN "&sqlData(inicio)&" AND "&sqlData(fim)&" AND TB_PRODUCAO_LOTES_FORMANDOS.COD_VENDEDOR = COD_VENDEDOR)) AND COD_VENDEDOR = "&sqlNumero(vendedores(count,0))
		'response.write sql&"<hr>"

		set rs = objConexao.Execute(sql)
		if not rs.eof then
			matriz_ordena(count,0) = vendedores(count,1)
			matriz_ordena(count,1) = cdbl("0"&rs(0))
		end if
		rs.close
		set rs = nothing
	else
		matriz_ordena(count,0) = vendedores(count,1)
		matriz_ordena(count,1) = cdbl("0")
	end if
next

call ordenaMatriz()
call copiaMatriz(com_valores, 0)
call limpaMatrizOrdena()


for count = 0 to ubound(vendedores,1)
	if int("0"&vendedores(count,2)) > 0 then

		sql = "SELECT sum(VW_QTD_FORMANDOS_POR_PEDIDO.QTD_FORMANDOS) AS QTD_VENDIDO FROM VW_QTD_FORMANDOS_POR_PEDIDO INNER JOIN TB_PEDIDOS ON VW_QTD_FORMANDOS_POR_PEDIDO.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO WHERE DATA_CORRIGIDO_DEP_FOTO is not null AND "&campo_base&" BETWEEN "&sqlData(inicio)&" AND "&sqlData(fim)&" AND TB_PEDIDOS.COD_PEDIDO IN (SELECT COD_PEDIDO FROM TB_PEDIDOS_FORMANDOS WHERE COD_FORMANDO IN (SELECT COD_FORMANDO FROM TB_PRODUCAO_LOTES_FORMANDOS WHERE DATA_SAIDA BETWEEN "&sqlData(inicio)&" AND "&sqlData(fim)&"))  AND FINALIZA_CORRECAO = 1 AND COD_VENDEDOR = "&sqlNumero(vendedores(count,0))
		'response.write sql&"<hr>"

		set rs = objConexao.Execute(sql)
		if not rs.eof then
			matriz_ordena(count,0) = vendedores(count,1)
			matriz_ordena(count,1) = int("0"&rs(0))

			vendedores(count,3) = int("0"&rs(0))
		end if
		rs.close
		set rs = nothing
	else
		matriz_ordena(count,0) = vendedores(count,1)
		matriz_ordena(count,1) = cdbl("0")	

		vendedores(count,3) = int("0")
	end if
next

call ordenaMatriz()
call copiaMatriz(com_valores, 1)
call limpaMatrizOrdena()

for count = 0 to ubound(vendedores,1)

	if int("0"&vendedores(count,2)) > 0  then
		matriz_ordena(count,0)=vendedores(count,1)
		matriz_ordena(count,1)=round( (100*int("0"&vendedores(count,3)))/int("0"&vendedores(count,2)) , 2)
	else
		matriz_ordena(count,0)=vendedores(count,1)
		matriz_ordena(count,1)=0
	end if

next

call ordenaMatriz()
call copiaMatriz(com_valores, 2)
call limpaMatrizOrdena()

for count = 0 to ubound(vendedores,1)
	if int(vendedores(count,0)) > 0 then

		sql = "SELECT SUM( VW_QTD_FORMANDOS_POR_PEDIDO.QTD_FORMANDOS ) AS TOTAL_ALBUNS_GERAL FROM VW_QTD_FORMANDOS_POR_PEDIDO INNER JOIN TB_PEDIDOS ON VW_QTD_FORMANDOS_POR_PEDIDO.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO WHERE DATA_CORRIGIDO_DEP_FOTO is not null AND TB_PEDIDOS."&campo_base&" BETWEEN "&sqlData(inicio)&" AND "&sqlData(fim)&" AND FINALIZA_CORRECAO = 1 AND COD_VENDEDOR = "&sqlNumero(vendedores(count,0))
		response.write sql&"<hr>"
		set rs = objConexao.Execute(sql)
		if not rs.eof then
			matriz_ordena(count,0) = vendedores(count,1)
			matriz_ordena(count,1) = int("0"&rs(0))
		end if
		rs.close
		set rs = nothing
	end if
next

call ordenaMatriz()
call copiaMatriz(com_valores, 3)
call limpaMatrizOrdena()

for count = 0 to ubound(vendedores,1)
	if int(vendedores(count,0)) > 0 then

		sql = "SELECT (SUM(IFNULL(PAGAMENTO_EM_DINHEIRO,0)) + SUM(IFNULL(XTOTAL_CHEQUE,0)) + SUM(IFNULL(XTOTAL_CARTAO,0))) AS TOTAL_RECEBIDO_GERAL FROM VW_EXTRATO_TOTAL_RECEBIDO_COMPLETO WHERE DATA_CORRIGIDO_DEP_FOTO is not null AND "&campo_base&" BETWEEN "&sqlData(inicio)&" AND "&sqlData(fim)&" AND COD_VENDEDOR = "&sqlNumero(vendedores(count,0))
		response.write sql&"<hr>"

		set rs = objConexao.Execute(sql)
		if not rs.eof then
			matriz_ordena(count,0) = vendedores(count,1)
			matriz_ordena(count,1) = CDBL("0"&rs(0))
		end if
		rs.close
		set rs = nothing
	end if
next

call ordenaMatriz()
call copiaMatriz(com_valores, 4)
call limpaMatrizOrdena()

for count = 0 to ubound(vendedores,1)
	if int("0"&vendedores(count,0)) > 0  then
		matriz_ordena(count,0)=vendedores(count,1)
		matriz_ordena(count,1)=vendedores(count,2)
	end if
next

call ordenaMatriz()
call copiaMatriz(com_valores, 5)
call limpaMatrizOrdena()


function exibeNome(valor)

	dim retorno

	if inStr(1,valor,"(",vbTextCompare) > 0  then

		retorno = split(valor," ")(0) & " " & split(valor," ")(ubound(split(valor," ")))

	elseif ubound(split(valor," ")) >= 1 then
		retorno = split(valor," ")(0)
	end if

exibeNome = retorno
end function
%>

<html>
	<head>
		<title>Relatório de Ranking Geral</title>
		<style>
			body{
				font-family: arial;
				font-size: 12px;
			}
			h3{
				text-align: center;
			}
			table{
				border-collapse: collapse;
			}
			th {
				font-size: 13px;
			}
			td {
				font-size: 11px;
			}

			/*ALTERACOES CAUE*/
			.conteudo{
				overflow: auto;
				max-width: 100%;
				width: 29.7cm;
				margin: 0 auto; 
			}
			
			div#tabela-principal{
				margin: 5pt 0;
				padding: 0 2pt;
			}

			div#diretoria{
				width: 49%;
				float: right;
				padding: 2pt;
			}
			
			div#diretoria > table{
				float: right;
			}

			div#emitido{
				width: 49%;
				float: left;
				padding: 2pt;
			}

			@page {
				margin: 0.2cm;
				margin-top: 0.5cm;
				margin-bottom: 0.5cm;
				size: A4 landscape;
				max-height:100%;
				max-width:100%
			}
			@media print {
				* {
					background:transparent !important;
					text-shadow:none !important;
					filter:none !important;
					-ms-filter:none !important;
				}
				
				.page {
					margin: 0.2cm;
					margin-top: 0.5cm;
					margin-bottom: 0.5cm;
					border: initial;
					border-radius: initial;
					width: initial;
					min-height: initial;
					box-shadow: initial;
					background: initial;
					page-break-after: always;
				}
				#tabela-principal{
					color: red;
				}

			}

		</style>
	</head>
	<body>
		<div class="conteudo">

			<h3>RELATÓRIO PARA RETIRADA DE PRODUÇÃO (<%=inicio%> até <%=fim%>)</h3>

			<%
			dim inicio_t, fim_t
			inicio_t = mid(inicio,1,10)
			fim_t = mid(fim,1,10)
			%>
			<div id="tabela-principal">
				<!--<table width="98%" align="center" border="1" cellpadding="3">-->
				<table align="center" border="1" cellpadding="3">
					<tr style="background-color: whitesmoke;">
						<th width="50px">#</th>
						<th width="250px">
							Valor total arrecadado geral (<%=inicio_t%> até <%=fim_t%>)
						</th>
						<th width="250px">
							Qtd. Total de Álbs vendidos geral  (<%=inicio_t%> até <%=fim_t%>)
						</th>
						<th width="250px">
							Valor Arrecadado de todos os lotes retirados no periodo (<%=inicio_t%> até <%=fim_t%>)
						</th>
						<th width="250px">
							Qtd. de Álbs vendidos de todos os lotes retirados no periodo (<%=inicio_t%> até <%=fim_t%>)
						</th>
						<th width="250px">
							Aproveitamento de todos os lotes retirados no priodo (<%=inicio_t%> até <%=fim_t%>)
						</th>
						<th width="250px">
							Qtd. de Álbs retirados no periodo (<%=inicio_t%> até <%=fim_t%>)
						</th>
						<th width="250px">
							Colocação para retirada de nova produção
						</th>
						<th width="250px">
							Qtd. de Fichas
						</th>
					</tr>

					<% for count = 0 to ubound(resultados)
					if trim(exibeNome(resultados(count,4))) = "(0)" or trim(exibeNome(resultados(count,4))) = ""  then
						exit for
					end if
					%>
					<tr>
						<td align="center"><%=count+1%>º</td> <!-- Coluna # -->
						
						<td align="center" class="td-value">
							<%=exibeNome(resultados(count,4))%>
						</td> <!-- Coluna Valor total arrecadado geral -->
						
						<td align="center" class="td-value">
							<%=exibeNome(resultados(count,3))%>
						</td> <!-- Coluna Qtd. Total de Álbs vendidos geral  -->
						
						<td align="center" class="td-value">
							<%=exibeNome(resultados(count,0))%>
						</td><!-- Coluna Valor Arrecadado de todos os lotes retirados no periodo  -->
						
						<td align="center" class="td-value">
							<%=exibeNome(resultados(count,1))%>
						</td><!-- Coluna Qtd. de Álbs vendidos de todos os lotes retirados no periodo  -->
						
						<td align="center" class="td-value">
							<%=exibeNome(resultados(count,2))%>
						</td><!-- Coluna Aproveitamento de todos os lotes retirados no priodo  -->
						
						<td align="center" class="td-value">
							<%=exibeNome(resultados(count,5))%>
						</td><!-- Qtd. de Álbs retirados no periodo  -->
						
						<td align="center"></td><!-- Coluna Colocação para retirada de nova produção -->
						
						<td align="center"></td><!-- Coluna Qtd. de Fichas -->
					</tr>
					<% next %>
				</table>	
			</div>
			<!--<table style="position: absolute;  bottom: 30px; right: 30px; font-size: 14px;" id="table-diretoria"> -->
			<div id="diretoria">
				<table id="table-diretoria">
					<tr>
						<td align="center">_____________________________________</td>
					</tr>

					<tr>
						<td align="center">Diretoria</td>
					</tr>
					<tr>
						<td align="center">Data: ___/___/____</td>
					</tr>
				</table>	
			</div>
			
			<div id="emitido">
				<!--<table style="position: absolute;  bottom: 30px; left: 10px; font-size: 10px;" id="table-emitido">-->
				<table id="table-emitido">
					<tr>
						<td align="center">Emitido em: <%=arrumaData(now())%></td>
					</tr>
				</table>	
			</div>
			
		</div>
	</body>
</html>