<%
Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
BTN_VOLTAR = 1
BTN_IMPRIMIR = 1
BTN_FONTE = 2
BTN_IMPRIMIR_LINK = "#"

Dim titulo_pagina
titulo_pagina = "Triagem Entrega | Transpostadora"

Dim calendario_necessario
calendario_necessario = "N&AtildeO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<%
'Come&ccedilamos o cron&otildemetro!
starttime = Timer() 
Dim letra : letra = ucase(request.QueryString("letra"))

Dim efetuar_despacho, despacho_transportador
efetuar_despacho = Cstr(""&Request.Form("campo1"))	
despacho_transportador = Cstr(""&Request.Form("despacho_transportador"))	

AbrirConexao()

'Marcando produtos como entregue
produtos = Cstr(""&Request.Form("produtos"))	
produtos_array = split(produtos,",")
for each produto in produtos_array
	produto = Replace(produto," ","")
	data = Cstr(""&Request.Form("data_entrega_"&produto))
	quantidade_entregue = Cstr(""&Request.Form("quantidade_entregue_"&produto))
	if len(data)>9 then
		data_entrega_array = split(data,"/")

		data = "'"&data_entrega_array(2)&"-"&data_entrega_array(1)&"-"&data_entrega_array(0)&"'"
		sql = "INSERT TB_PEDIDOS_ITENS_ENTREGUES (cod_ordem_servico,data,data_lancamento,quantidade_entregue) VALUES ("&produto&","&data&",now(),"&sqlNumero(quantidade_entregue  )&")"

		objConexao.Execute(sql)
	end if
next

Dim total_os ,count,listar,limit_from,limit_to,order_nome,busca_cod_ordem_servico, busca_nome_formando, busca_cod_contrato, busca_cod_formando, busca_finalizados_parcialmente, busca_cod_usuario_responsavel, busca_despacho_transportador

Dim sql, sqlWhere, sqlOrderBy, sqlGroupBy

sqlOrderBy = "cod_contrato, nome_formando"
busca_nome_formando = Cstr(""&Request.QueryString("nome_formando"))
busca_cod_contrato = Cstr(""&Request.QueryString("cod_contrato"))
busca_cod_formando = Cstr(""&Request.QueryString("cod_formando"))
busca_finalizados_parcialmente = Cstr(""&Request.QueryString("finalizados_parcialmente"))
mostrar_entregues = Cstr(""&Request.QueryString("mostrar_entregues"))
ano = request.QueryString("ano")
busca_despacho_transportador = Cstr(""&request.QueryString("despacho_transportador"))

if ano = "" then
  ano = "maior_2016"
end if

busca_cod_usuario_responsavel = Cstr(""&Request.QueryString("cod_usuario_responsavel"))
listar= Cstr(""&Request.QueryString("listar"))

'Padr&atildeo
Dim validacao_listar_os_todos_usuarios, valida_para_despacho_qualquer_usuario
validacao_listar_os_todos_usuarios = validaUsuario(2,,57,false)
valida_para_despacho_qualquer_usuario = validaUsuario(2,,76,false)

If not valida_para_despacho_qualquer_usuario Then
	sqlWhere = sqlWhere&" and cod_usuario_responsavel = "&sqlTexto(Session("cod_usuario"))	
End If	

'Por padr&atildeo lista somente os pedidos que est&atildeo totalmente prontos para despacho
If Request.QueryString("filtro1") <> Empty Then					
	If busca_despacho_transportador <> "" then
    sqlWhere = sqlWhere&" AND despacho_transportador = "&sqlTexto(busca_despacho_transportador)
  else 
    sqlWhere = sqlWhere&" and despacho_transportador = 'TRANSPORTADORA' "
  end if
  
  'Busca prim&aacuteria
	If busca_finalizados_parcialmente = "S" Then
	'Se a op&ccedil&atildeo de finalizados parcialmente estiver desabilidata ir&aacute filtrar somente TOTALMENTE finalizados para despacho
		sqlWhere = sqlWhere&" AND VW_TB_ORDENS_SERVICO.cod_pedido NOT IN (SELECT TB_ORDENS_SERVICO.cod_pedido FROM TB_ORDENS_SERVICO WHERE status_ordem_servico IN(0,1))"
	End If

	If busca_cod_usuario_responsavel <> "" Then
		sqlWhere = sqlWhere&" and cod_usuario_responsavel = "&sqlTexto(busca_cod_usuario_responsavel)	
	End If
	
	If busca_nome_formando <> "" Then
		sqlWhere = sqlWhere&" and nome_formando LIKE '%"&busca_nome_formando&"%'"	
	End If

	If busca_cod_contrato <> "" Then
		sqlWhere = sqlWhere&" and cod_contrato = "&sqlTexto(busca_cod_contrato)			
	End If

  If ano = "maior_2016" Then
		sqlWhere = sqlWhere&" and ano > 2016 "
  else
  	sqlWhere = sqlWhere&" and ano <= 2016 "
	End If

	If busca_cod_formando <> "" Then
		temp = Right(busca_cod_formando,1)
		If not (Cstr(""&temp) <> "1") Then
			sqlWhere = sqlWhere&" and cod_formando = "&sqlTexto(Mid(busca_cod_formando,1,Len(busca_cod_formando)-1))
		End if
	End If
Else
	'sqlWhere = sqlWhere&" AND VW_TB_ORDENS_SERVICO.cod_pedido NOT IN (SELECT cod_pedido FROM TB_ORDENS_SERVICO WHERE status_ordem_servico IN(0,1))"
End If

Dim acesso_inclusao_exclusao_os_manual
acesso_inclusao_exclusao_os_manual = validaUsuario(2,,51,false)
%>
<h2>PPV - Confirma&ccedil&atildeo de Entrega</h2>
<br />
<!-- #include file="includes/classe_modulo_ppv_menu.asp" -->
<br />
<br />
<div class="msgGrande">
	Ap&oacutes confirmada a entrega dos itens a opera&ccedil&atildeo n&atildeo poder&aacute ser desfeita.
</div>

<form style="margin:0px">
	<table style="border:#999999 dashed 1px;  margin-bottom:7px;padding:2px; background:#FFFFCC" width="99%"  align="center" cellpadding="5" cellspacing="0">
		<tr>
			<td width="150px"><strong>Nome do Formando</strong></td>
			<td width="300px" colspan="4">
				<%=PadraoFormularioSelect("letra","","inputTamanho1","","","SEM_SELECIONE",letra,"PADRAO:A#A|B#B|C#C|D#D|E#E|F#F|G#G|H#H|I#I|J#J|K#K|L#L|M#M|N#N|O#O|P#P|Q#Q|R#R|S#S|T#T|U#U|V#V|W#W|X#X|Y#Y|Z#Z", "incluir", false, false, false )%>
			</td>
		</tr>
    <tr>
			<td width="150px"><strong>Tipo envio</strong></td>
			<td width="300px" colspan="4">
				<%=PadraoFormularioSelect("despacho_transportador","","inputTamanho2","","","SEM_SELECIONE",busca_despacho_transportador,"PADRAO:TRANSPORTADORA#Transportadora|CORREIOS#Correios|RETIRADA#Retirada|STILLOS#Stillos|PANDEMIA#Pandemia|LOGISTICA#Logistica|ESPECIAL#Especial|ANTONIO FELIPE#Antonio Felipe", "incluir", false, false, false )%>
			</td>
		</tr>
    <tr>
			<td width="150px"><strong>Ano</strong></td>
			<td width="300px" colspan="4">
				<%=PadraoFormularioSelect("ano","","inputTamanho2","","","SEM_SELECIONE",ano,"PADRAO:maior_2016#Maior 2016|menor_2016#Menor ou igual 2016", "incluir", false, false, false )%>
			</td>
		</tr>
		<tr>
			<td width="150px"><strong>Nome do Formando</strong></td>
			<td width="300px">
				<input type="text" value="<%=busca_nome_formando%>" name="nome_formando" class="inputTamanho2" />
			</td>
			<td rowspan="5">
				<div align="left">
					<input class="btnFiltrarAmarelo" type="submit" name="filtro1" value=" " />
				</div>
			</td>
			<td rowspan="5">
				<div align="center">
					<div style="width:250px;" class='msgGrandeErro'>
						<em><strong>Listagem em Ordem alfab&eacutetica pelo contrato, nome do formando com todas OS, cujo os produtos foram despachados para o cliente. <br> Para obter outros utilize os campos de pesquisa.</strong></em><br />
					</div>
				</div>
			</td>
		</tr>
		<tr>
			<td>
				<strong>C&oacuted. Formando Ficha</strong>
				<a href="javascript:alert('(C&oacuted. do formando acressido de \'1\' no final)')">[?]</a>
			</td>
			<td>
				<input type="text" value="<%=busca_cod_formando%>" name="cod_formando" class="inputTamanho2" />
			</td>
		</tr>
		<tr>
			<td>
				<strong>Cod. Contrato</strong>
			</td>
			<td>
				<input type="text" value="<%=busca_cod_contrato%>" name="cod_contrato" class="inputTamanho2" />
			</td>
		</tr>
		<tr>
			<td><strong>Respons&aacutevel  OS</strong></td>
			<td>
				<%=PadraoFormularioSelect("cod_usuario_responsavel","","inputTamanho2","","","Todos",busca_cod_usuario_responsavel,"SELECT cod_usuario, nome FROM TB_ADM_USUARIOS WHERE COD_USUARIO IN(SELECT cod_usuario_responsavel FROM VW_TB_ORDENS_SERVICO WHERE status_ordem_servico = 2 GROUP BY cod_usuario_responsavel) ORDER BY NOME", "incluir", false, false, false )%>
			</td>
		</tr>
	</table>
</form>
<div class="subMenu">
	<a href="javascript:inclusaoManual()">Entrega manual</a>
</div>
<style type="text/css">
	.tinner{
		padding: 0px;
		border-radius: 0px;
	}
</style>
<script type="text/javascript">
	function inclusaoManual()
	{
		TINY.box.show({iframe:'ppv_ordens_servico_triagem_entrega_manual.asp',width:350,height:215});
	}
</script>
<%				

If mostrar_entregues = "" Then
	'sqlEntregue = " where  TB_PEDIDOS_ITENS_ENTREGUES.data is null"
End If

count = 0


sql = "SELECT VW_TB_ORDENS_SERVICO.cod_pedido FROM VW_TB_ORDENS_SERVICO WHERE cod_ordem_servico not in (SELECT cod_ordem_servico FROM TB_PEDIDOS_ITENS_ENTREGUES) AND VW_TB_ORDENS_SERVICO.nome_formando like '"& replace(sqlTexto(letra),"'","") &"%' AND status_ordem_servico = 3  "&sqlWhere&"  GROUP BY VW_TB_ORDENS_SERVICO.cod_pedido  ORDER BY  nome_formando, cod_contrato desc "

AbrirConexao()

if letra <> "" then
	Set rs=server.CreateObject("ADODB.recordset")
	rs.CursorLocation = 3 'Cursor no cliente
	rs.CursorType = 0
	rs.PageSize = 500
	rs.LockType = 1
	rs.Open sql, objConexao

	'response.write sql

	Dim PAGINA_ATUAL
	PAGINA_ATUAL = formatNumber(0&Request.QueryString("pagina"),0)

	If not rs.Eof Then
		If PAGINA_ATUAL > 0 Then
			rs.AbsolutePage=PAGINA_ATUAL
		Else
			PAGINA_ATUAL = 1
			rs.AbsolutePage=1
		End If

%>
<form name="CONFIRMA_ENTREGA_CLIENTE" action="ppv_ordens_servico_triagem_entrega.asp?letra=<%=letra%>" method="post" style="margin:0px">
	<input name="campo1" value="S" type="hidden" />
	<table width="99%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
		<tr bgcolor="whitesmoke">
			<td colspan="2"><div align="center"><strong>OS</strong></div></td>
			<td width="23%"><div align="center"><strong>Produto</strong></div></td>
			<td width="12%"><div align="center"><strong>Data Finalizada</strong></div></td>
			<td width="12%"><div align="center"><strong>Prazo Geral</strong></div></td>
			<td width="11%"><div align="center"><strong>Marcar como entregue</strong></div></td>
		</tr>
<%
Dim numero_etapa, descricao, cod_usuario_responsavel, nome_usuario_responsavel, prazo_em_dias, data_inicio, prazo_limite, diferenca, qtd_etapas
Dim x

For x=1 To rs.PageSize
	If Not rs.Eof Then
		count = int(count) + 1
		
		Set rs2=objConexao.Execute("SELECT TB_ORDENS_SERVICO.* FROM TB_ORDENS_SERVICO WHERE cod_pedido = "&rs("cod_pedido")&" and data_finalizado IS NULL and status_ordem_servico NOT IN (0,1,2,99)")
		If not rs2.Eof Then
			permitido = false
		Else
			permitido = true							
		End If
		rs2.Close
		Set rs2=Nothing

		%>
		<tr>
			<td colspan="5">
			<table width="650" border="0" cellspacing="0" cellpadding="1">
			<tr>
			<td width="66" class="textoMedio"><strong><%=rs("cod_pedido")%></strong></td>
			<td width="580">
			<%
			sql = 	" SELECT "&_
					" TB_PEDIDOS_FORMANDOS.COD_FORMANDO, "&_
					" TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, "&_
					" TB_FICHAS_IDENTIFICACAO.COD_CONTRATO, "&_
					" TB_ENTIDADES.DESCRICAO, "&_
					" TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO, "&_
					" TB_FICHAS_IDENTIFICACAO.NOME_CURSO "&_
					" FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON "&_
					" TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO "&_
					" INNER JOIN TB_ENTIDADES ON(TB_FICHAS_IDENTIFICACAO.COD_CONTRATO = TB_ENTIDADES.COD_CONTRATO)"&_
					" WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&rs("cod_pedido")&" ORDER BY NOME_FORMANDO"

			Set rs2=objConexao.Execute(sql)
			%><table cellpadding="0" width="580" cellspacing="0"><%
				While not rs2.Eof
				%>
					<tr>
						<td width="36%"><%=rs2("NOME_FORMANDO")%></td>
						<td width="64%"><%=rs2("DESCRICAO")%></td>
					</tr>
				<%
				rs2.MoveNext()
				Wend
			%></table><%
			rs2.Close
			Set rs2=Nothing
			%>
			</td>
			</tr>
			</table>
			</td>
			<td class="textoMedio"><div align="center">
        <input type="text" name="data_entrega_geral_<%=rs("cod_pedido")%>" class='padrao_data' />
      </div></td>
		</tr>
		<%	   
		If mostrar_entregues <> "" Then
			'sqlWhereProduto = " and TB_PEDIDOS_ITENS_ENTREGUES.data is not null"
		else
			' sqlWhereProduto = " and TB_PEDIDOS_ITENS_ENTREGUES.data is null"
		End If

		sql = "SELECT VW_TB_ORDENS_SERVICO.*,cast(TB_PEDIDOS_ITENS_ENTREGUES.data as date) AS data_entrega FROM VW_TB_ORDENS_SERVICO LEFT JOIN TB_PEDIDOS_ITENS_ENTREGUES ON TB_PEDIDOS_ITENS_ENTREGUES.cod_ordem_servico = VW_TB_ORDENS_SERVICO.cod_ordem_servico WHERE VW_TB_ORDENS_SERVICO.cod_pedido = "&rs("cod_pedido")&"  and VW_TB_ORDENS_SERVICO.status_ordem_servico NOT IN (0,1,2,99) "&sqlWhereProduto&" order by VW_TB_ORDENS_SERVICO.numero_ordem_servico desc, numero_ordem_servico_indice asc"

		Set rs2=objConexao.Execute(sql)
		While not rs2.Eof
			Set rs3=objConexao.Execute("Select count(*) as qtd_etapas From TB_ORDENS_SERVICO_ETAPAS where cod_ordem_servico = "&rs2("cod_ordem_servico")&" group by cod_ordem_servico")
				If not rs3.Eof Then
					qtd_etapas = formatNumber(0&rs3("qtd_etapas"),0)
				End If
			rs3.Close
			Set rs3=Nothing
			data_finalizado = rs2("data_finalizado")
			%>
			<tr <%If not isDate(data_finalizado) Then%> bgcolor="#FFAAAA" <%End If%>>
			<td width="3%">
				<div align="right">
					<img src="images/icons/seta_superior.gif" width="19" height="15" />
				</div>
			</td>
			<td width="22%"><div align="center"><%=moduloPPVOSMascara(rs2("cod_ordem_servico"),"",true)%></div></td>
			<td><div align="left">(<input type="text" name="quantidade_entregue_<%=rs2("cod_ordem_servico")%>" value="<%=rs2("quantidade")%>" class="inputTamanho1 padrao_numero" style="width:25px; text-align:center;">)<%=rs2("nome_produto")%></div></td>
			<td><div align="center"><%=ArrumaData(data_finalizado)%></div></td>
			<td><div align="center"><%=ArrumaData(DateAdd("d",rs2("prazo_cliente_total"),rs2("data_base_prazo_cliente_inicio")))%></div></td>
			<td >
				<div align="center">
				<% if rs2("data_entrega") <> "" then %>
				Entregue (<%=ArrumaData(rs2("data_entrega"))%>) 
				<% else %>
					<table>
					<tr>
						<td>
							<input type="checkbox" value="<%=rs2("cod_ordem_servico")%>" pedido="<%=rs("cod_pedido")%>" name="produtos" class="inputRadio" />
						</td>
						<td>
							<input type="text" name="data_entrega_<%=rs2("cod_ordem_servico")%>" class="padrao_data pedido_<%=cod_pedido%>" />
							</div>
						</td>
					</tr>
					</table>
					<% End If %>
				</div>
			</td>
			</tr>
		<%
		rs2.MoveNext()
		Wend

		rs2.Close
		Set rs2=Nothing
	rs.MoveNext()
	end if
Next
%>
		<tr>
			<td colspan="7">
				<div align="center"><a href="javascript:confirmaEntrega();">
					<img border="0" src="images/layout/btnGravar.jpg" />
				</a></div>
			</td>
		</tr>
	</table>
</form>

<% Else %>
<div align="center" class="msgGrande">
	Nenhuma OS encontrada para confirmar entrega!<br />
</div>
<% End if %>

<p align="center"> 
<%If rs.PageCount > 1 Then %>
	<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
		<tr>
			<td> 
				<div class="menu_paginacao">
				Ir para a p&aacutegina:
				<% Dim aae : aae = 0
				For x=1 To rs.PageCount 
					aae = formatNumber(0&x,0)
					If int(PAGINA_ATUAL) <> int(aae) Then %>
						<a href="?pagina=<%=x%>&f=<%=FILTRO_F%>"><%=x%></a>
					<%Else
						Response.Write(x)
					End If
				Next%>
				</div>
			</td>
		</tr>
	</table>
<%End if

rs.Close
Set rs=Nothing

FecharConexao()

Else
%>
<div align="center" class="msgGrande">
	Selecione uma letra e clique em filtrar para carregar a listagem!<br />
</div>
<%
End if

%>
</p>
<script type="text/javascript">
	$(function(){
		$("input[name=produtos]").change(function(){
      var cod_pedido = $(this).attr('pedido');
      var cod_os = $(this).val();
			if ($(this)[0].checked) {
        if ($("input[name=data_entrega_"+cod_os+"]").val() == "") {
          var data_entrega_geral = $("input[name=data_entrega_geral_"+cod_pedido+"]").val();
          $("input[name=data_entrega_"+cod_os+"]").val(data_entrega_geral);
        }
			} else {
        $("input[name=data_entrega_"+cod_os+"]").val("");
      }
		});
	});
</script>
<script type="text/javascript">
function confirmaEntrega()
{
if(YNconfirm('Tem certeza que deseja confirmar a entrega dos itens selecionados?'))
document.CONFIRMA_ENTREGA_CLIENTE.submit();
}
</script>
<%

endtime = Timer() 
Response.Write "O carregamento se completou em " & endtime-starttime & " segundos = " 
Response.Write " (" & (endtime-starttime)*1000 & " mil&eacutesimos de segundos)."
%>
<!-- #include file="files/includes/miolo-inferior.asp" -->



