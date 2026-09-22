<%
Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 2
BTN_IMPRIMIR_LINK = "#"
titulo_pagina = "Triagem Entrega | Correios"

Server.ScriptTimeout = 1800

Dim calendario_necessario
calendario_necessario = "N&AtildeO"
dim msg : msg = request.QueryString("msg")
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<%


' Confirma&ccedil&atildeo de entrega aos correios
' Nesta pagina fazemos a leitura do arquivo  de retorno dos correios e 
'
'
dim caminho : caminho = "e:\home\stillos\dados\gerencial_foto\retorno_correios\"
Set Upload = Server.CreateObject("SoftArtisans.FileUp") 
	
dim campo1 : campo1 = Upload.form("campo1")
dim campo2 : campo2 = Upload.form("campo2")
comando = Request.QueryString("comando")

AbrirConexao()

if campo1 = "1" then

	Set Arquivo = Upload.form("arquivo")
		If Arquivo <> "" AND lcase(right(arquivo.ShortFilename,4)) = ".txt" then

			Arquivo.SaveAs (caminho&arquivo.ShortFilename)
			dim objFSO, txt, i

			set objFSO = Server.CreateObject("Scripting.FileSystemObject")
			set txt = objFSO.OpenTextFile(caminho & arquivo.ShortFilename, 1)
			i = 0

			dim lista : lista = txt.readAll
			lista = mid(lista, 1, len(lista)-2)
			lista = split(lista, chr(13)&chr(10))

			for i = 0 to ubound(lista)
				lista(i) = replace(lista(i),chr(34),"")
				lista(i) = split(lista(i) , ";")
			next

			txt.Close
			set txt = nothing
			set objFSO = nothing

		End If
	Set Arquivo = nothing
end if

if campo2 = "2" then
	
	dim pedidos : pedidos = replace(Upload.form("pedidos")," ","")
	pedidos = split(pedidos,",")

	for i = 0 to ubound(pedidos)
		temp = Upload.form(pedidos(i))
		data = split(temp,"#")(0)
		rastreio = split(temp,"#")(1)

		'Efetuamos o insert para todas as OS do pedido 
		set rsInsert = objConexao.Execute("SELECT cod_ordem_servico FROM TB_ORDENS_SERVICO WHERE COD_PEDIDO = "&sqlNumero(pedidos(i))&" AND status_ordem_servico = 3 AND ucase(despacho_transportador) <> 'TRANSPORTADORA' ")
		while not rsInsert.Eof

			objConexao.Execute("INSERT TB_PEDIDOS_ITENS_ENTREGUES (cod_ordem_servico,data,data_lancamento,codigo_rastreio) VALUES ("&rsInsert(0)&","&sqlData(data)&",now(),"&sqlTexto(rastreio)&")")

			rsInsert.MoveNext
		Wend
		rsInsert.Close
		set rsInsert = nothing
	next
response.redirect("ppv_ordens_servico_triagem_entrega_correios.asp?msg=Baixas realizadas com sucesso!&cod_masg=2")
end if

Set Upload = nothing

function relacionaFormandoCorreio(nome_formando, nome_campo)
	
	dim comp : comp = ""
	dim retorno : retorno = ""

	retorno = retorno & "<select name= '"&nome_campo&"' class='chosen-select' >"

	for i = 0 to ubound(lista)
		comp = ""
		if ucase(replace(nome_formando," ","")) = ucase(replace(lista(i)(3)," ","")) then
			comp = "selected='selected' "
		end if
		if i = 0 then
			retorno = retorno & "<option value=''>N&atildeo encontrado</option>"
		else
			retorno = retorno & "<option "&comp&"value='"&lista(i)(0) &"#"& lista(i)(9) &"'>"&lista(i)(3) & "   -   " & lista(i)(9) & "BR"&"</option>"
		end if
	next
	retorno = retorno & "</select>"

relacionaFormandoCorreio = retorno
end function

fecharConexao()
'----------------------------------------------------------------------------------------
%>
<h2><%=titulo_pagina%></h2>
<br />
<!-- #include file="includes/classe_modulo_ppv_menu.asp" -->
<br />
<br />
<form action='<%=actionformulario%>' name='FORMULARIO_PADRAO' method='post'  enctype="multipart/form-data">
  <table width="600" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
    <tr bgcolor="whitesmoke">
		<td>
			<h4>Instru&ccedil&otildees</h4>
			<ol start="1">
				<li>Importa&ccedil&atildeo do arquivo retorno dos correios.</li>
				<li>Para fazer baixar o arquivo </li>
			</ol>  
		</td>
	</tr>
	<tr>
		<td>
			<div align="center">
				<input type="file" class="inputTamanho3" name="arquivo" />
			</div>
		</td>
	</tr>
	<tr>
		<td>
			<div align="center">
				<a href="javascript:document.FORMULARIO_PADRAO.submit();">
					<img src="images/layout/btnGravar.jpg" />
				</a>
			</div>
		</td>
	</tr>
  </table>
<br />
<br />
<input type="hidden" name="campo1" value="1" />
</form>
<%				

If mostrar_entregues = "" Then
	sqlWhere = sqlWhere&" AND TB_PEDIDOS_ITENS_ENTREGUES.data is null"
End If

count = 0
sql = "SELECT VW_TB_ORDENS_SERVICO.cod_pedido FROM VW_TB_ORDENS_SERVICO LEFT JOIN TB_PEDIDOS_ITENS_ENTREGUES ON TB_PEDIDOS_ITENS_ENTREGUES.cod_ordem_servico = VW_TB_ORDENS_SERVICO.cod_ordem_servico WHERE status_ordem_servico = 3 AND ucase(despacho_transportador) <> 'TRANSPORTADORA' "&sqlWhere&" GROUP BY VW_TB_ORDENS_SERVICO.cod_pedido ORDER BY  nome_formando, cod_contrato desc LIMIT 10000"

'response.Write sql

AbrirConexao()
dim letra : letra = "a"
if campo1 = "1" then
	Set rs=server.CreateObject("ADODB.recordset")
	rs.CursorLocation = 3 'Cursor no cliente
	rs.CursorType = 0
	rs.PageSize = 1000
	rs.LockType = 1
	rs.Open sql, objConexao

	Dim PAGINA_ATUAL : PAGINA_ATUAL = formatNumber(0 & Request.QueryString("pagina"), 0)

	If not rs.Eof Then
		If PAGINA_ATUAL > 0 Then
			rs.AbsolutePage=PAGINA_ATUAL
		Else
			PAGINA_ATUAL = 1
			rs.AbsolutePage=1
		End If
%>
<form name="CONFIRMA_ENTREGA_CLIENTE" action="ppv_ordens_servico_triagem_entrega_correios.asp" enctype="multipart/form-data" method="post" style="margin:0px">
	<input name="campo2" value="2" type="hidden" />
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
		<tr id='tr_<%=rs("cod_pedido")%>' style='background-color:#F46F66;'>
			<td colspan="5">
			<table width="650" border="0" cellspacing="0" cellpadding="1">
			<tr class=''>
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

			dim nome_form 

			Set rs2=objConexao.Execute(sql)
			%><table cellpadding="0" width="580" cellspacing="0"><%
				nome_form =""
				While not rs2.Eof
					if nome_form = "" then
						nome_form = rs2("NOME_FORMANDO")
					end if 
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
			<td class="textoMedio">
				<%=relacionaFormandoCorreio(nome_form, rs("cod_pedido")) %>
				<input tabindex='5000<%=count%>' type="checkbox" name='pedidos' value='<%=rs("cod_pedido")%>' id='c_<%=rs("cod_pedido")%>' style='display:none; '>
				<%response.flush%>
			</td>
		</tr>
		<%	   
		If mostrar_entregues <> "" Then
			'sqlWhereProduto = " and TB_PEDIDOS_ITENS_ENTREGUES.data is not null"
		else
			sqlWhereProduto = " and TB_PEDIDOS_ITENS_ENTREGUES.data is null"
		End If

		sql = "SELECT VW_TB_ORDENS_SERVICO.*,cast(TB_PEDIDOS_ITENS_ENTREGUES.data as date) AS data_entrega FROM VW_TB_ORDENS_SERVICO LEFT JOIN TB_PEDIDOS_ITENS_ENTREGUES ON TB_PEDIDOS_ITENS_ENTREGUES.cod_ordem_servico = VW_TB_ORDENS_SERVICO.cod_ordem_servico WHERE VW_TB_ORDENS_SERVICO.cod_pedido = "&rs("cod_pedido")&"  and VW_TB_ORDENS_SERVICO.status_ordem_servico NOT IN (0,1,2,99) "&sqlWhereProduto&" order by VW_TB_ORDENS_SERVICO.numero_ordem_servico desc, numero_ordem_servico_indice asc"

		'n&atildeo estamos listando os intes, por padrao definimos como entregues todos os itens despachados pelo correio.
		if false then
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
			<td><div align="left"><%=rs2("nome_produto")%></div></td>
			<td><div align="center"><%=ArrumaData(data_finalizado)%></div></td>
			<td><div align="center"><%=ArrumaData(DateAdd("d",rs2("prazo_cliente_total"),rs2("data_base_prazo_cliente_inicio")))%></div></td>
			<td >
				<div align="center">
				<% if rs2("data_entrega") <> "" then %>
					Entregue (<%=ArrumaData(rs2("data_entrega"))%>) 
				<% else %>
					<input type="checkbox" value="<%=rs2("cod_ordem_servico")%>" name="produtos" class="inputRadio" />
				<% End If %>
				</div>
			</td>
			</tr>
		<%
		rs2.MoveNext()
		Wend

		rs2.Close
		Set rs2=Nothing

		end if
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
	if msg = "" then
%>
<div align="center" class="msgGrande">
	Selecione uma letra e clique em filtrar para carregar a listagem!<br />
</div>
<% 
	else
%>
<div align="center" class="msgGrandeSucesso">
	<%=msg%><br />
</div>
<% 
	end if
End if %>
</p>
<script type="text/javascript">
	$(function(){
		$(".chosen-select").chosen().change(function(){
			if ( $(this).val() != "" )
			{
				$("#c_"+$(this).attr('name'))[0].checked = true;
				$("#tr_"+$(this).attr('name')).css("background-color","#BDD7FC");
			}
			else
			{
				$("#c_"+$(this).attr('name'))[0].checked = false;
				$("#tr_"+$(this).attr('name')).css("background-color","#F46F66");
			}
		});

		$(".chosen-search > input[type=text]").focus(function(){
			$(this).parent().parent().parent().parent().parent().css("background-color","#cecece");
		});

		$(".chosen-search > input[type=text]").blur(function(){
			$tr = $(this).parent().parent().parent().parent().parent();
			if ($("select[name='"+$tr.attr('id').replace("tr_","")+"']").val() == "")
				var cor = "#F46F66";
			else
				var cor = "#BDD7FC";
			$(this).parent().parent().parent().parent().parent().css("background-color",cor);
		});

		$(".chosen-select").each(function()
		{
			if ( $(this).val() != "" )
			{
				$("#c_"+$(this).attr('name'))[0].checked = true;
				$("#tr_"+$(this).attr('name')).css("background-color","#BDD7FC");
			}
			else
			{
				$("#c_"+$(this).attr('name'))[0].checked = false;
				$("#tr_"+$(this).attr('name')).css("background-color","#F46F66");
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
<!-- #include file="files/includes/miolo-inferior.asp" -->