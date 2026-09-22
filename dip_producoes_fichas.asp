<% Server.ScriptTimeout = 300 %> 
<%
<!--starttime = Timer() 
Response.Buffer = False   ' <<< sempre True aqui-->
'//////////////////////////
'' Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 2
  BTN_FONTE = 1
  BTN_IMPRIMIR_LINK = "#"
Dim calendario_necessario : calendario_necessario = "NÃO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<!--BOTÃO DE SALVAR-->
<%
Dim mensagem : mensagem = ""

Function ToInt(v)
    v = Trim("" & v)
    If v <> "" And IsNumeric(v) Then
        ToInt = CLng(v)
    Else
        ToInt = 0
    End If
End Function

' === HANDLER DO BOTÃO SALVAR (ÚNICO) ===
If Request.ServerVariables("REQUEST_METHOD") = "POST" And Request.Form("acao") = "salvar" Then
    Dim COD_FORMANDO
    Dim ALBUM_FORM, ALBUM_DIGITAL_FORM, PASTA_FORM, ENCADERNACOES_EXTRAS_FORM
    Dim INICIAL_FORM, COMISSAO_FORM, BRINDE_FORM, PROD_AGREG_FORM, LOTE_FORM


    COD_FORMANDO             = ToInt(Request.Form("cod_formando"))
    ALBUM_FORM               = ToInt(Request.Form("ALBUM_" & COD_FORMANDO))
    ALBUM_DIGITAL_FORM       = ToInt(Request.Form("ALBUM_DIGITAL_" & COD_FORMANDO)) ' checkbox: 1/0
    PASTA_FORM               = ToInt(Request.Form("PASTA_" & COD_FORMANDO))
    ENCADERNACOES_EXTRAS_FORM= ToInt(Request.Form("ENCADERNACOES_EXTRAS_" & COD_FORMANDO)) ' checkbox: 1/0
    INICIAL_FORM             = ToInt(Request.Form("INICIAL_" & COD_FORMANDO))
    COMISSAO_FORM            = ToInt(Request.Form("COMISSAO_" & COD_FORMANDO))
    BRINDE_FORM              = ToInt(Request.Form("BRINDE_" & COD_FORMANDO))
    PROD_AGREG_FORM          = ToInt(Request.Form("PRODUTO_AGREGADO_" & COD_FORMANDO))
    LOTE_FORM                = ToInt(Request.Form("LOTE_" & COD_FORMANDO))

    ' === SALVAR NA TABELA TEMPORÁRIA ===
AbrirConexao()

' Se já existir um registro para este formando, apaga antes de inserir novamente
objConexao.Execute "DELETE FROM TB_FICHAS_IDENTIFICACAO_TEMPORARIA WHERE COD_FORMANDO=" & COD_FORMANDO

' Agora insere o novo registro
Dim sqlInsert
sqlInsert = "INSERT INTO TB_FICHAS_IDENTIFICACAO_TEMPORARIA " & _
            "(COD_FORMANDO, ALBUM, ALBUM_DIGITAL, numero_pasta_gerencia, " & _
            "ENCADERNACOES_EXTRAS, INICIAL, COMISSAO, BRINDE, PROD_AGREGADO, LOTE, DATA_GRAVADO_TEMPORARIO) VALUES (" & _
            COD_FORMANDO & ", " & ALBUM_FORM & ", " & ALBUM_DIGITAL_FORM & ", '" & PASTA_FORM & "', " & _
            ENCADERNACOES_EXTRAS_FORM & ", " & INICIAL_FORM & ", " & COMISSAO_FORM & ", " & _
            BRINDE_FORM & ", " & PROD_AGREG_FORM & ", " & LOTE_FORM & ", NOW())"


objConexao.Execute sqlInsert

FecharConexao()

	mensagem = "Formando " & COD_FORMANDO & " gravado na tabela temporária com sucesso."
End If
%>


<%
COD_FORMANDO_QS = Request.QueryString("COD_FORMANDO_QS")
NOME_FORMANDO_QS = Request.QueryString("NOME_FORMANDO_QS")

If Request.QueryString("comando") = "finalizar_producao" then
'FINALIZAR Produção
	AbrirConexao()
	Dim ver_LOTE
	'LanÃ§a o controle de lotes da Produção
	Set rs=objConexao.Execute("SELECT LOTE, COUNT(*) AS QTD_FORMANDOS FROM TB_FICHAS_IDENTIFICACAO WHERE cod_producao = "&Request.QueryString("cod_prod")&" group by lote order by lote")
	While not rs.Eof
		ver_LOTE = formatNumber(0&rs("LOTE"),0)

		If int(ver_LOTE) > 0 Then
			ver_LOTE = Replace(ver_LOTE,".","")
		
			'Insere os Lotes da Produção em uma tabela auxiliar de controle de lotes - CABE�ALHO
			'Estes dados sÃ£o os iniciais da Produção, podem ser alterados no futuro.
			objConexao.Execute("INSERT INTO TB_PRODUCAO_LOTES (COD_PRODUCAO, LOTE, QTD_FORMANDOS_INICIAL) VALUES ("&Request.QueryString("cod_prod")&","&ver_LOTE&",'"&rs("QTD_FORMANDOS")&"')")
		End If
	rs.MoveNext()
	Wend
		
	rs.Close
	Set rs=Nothing
	 
	'LanÃ§a o controle individual
	Set rs=objConexao.Execute("SELECT LOTE, COD_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE cod_producao = "&Request.QueryString("cod_prod")&" order by lote")
		
	While not rs.Eof
		
		ver_LOTE = formatNumber(0&rs("LOTE"),0)
		
		If int(ver_LOTE) > 0 Then
			ver_LOTE = Replace(ver_LOTE,".","")
		
			'deletamos antes so para ter certeza que nao vai dar erro'
			objConexao.Execute("DELETE FROM TB_PRODUCAO_LOTES_FORMANDOS where cod_formando = '"&rs("COD_FORMANDO")&"'")

			'Insere os Lotes da Produção em uma tabela auxiliar de controle de lotes, esta tabela ir� alterar com conforme as mudan�as de vendido, transferido e renovado e cada formando tera�o uma linha apenas desta tabela
			'Status padr�o 3 NÃO vendido e vendedor ainda NÃO selecionado
			objConexao.Execute("INSERT INTO TB_PRODUCAO_LOTES_FORMANDOS (COD_FORMANDO) VALUES ('"&rs("COD_FORMANDO")&"')")

		End If
	
	rs.MoveNext()
	Wend
		
	rs.Close
	Set rs=Nothing
	
	objConexao.Execute("UPDATE TB_PRODUCAO SET DATA_FINALIZADA = NOW(), PRODUCAO_FINALIZADA = '1' WHERE COD_PRODUCAO ="&Request.QueryString("cod_prod")&" LIMIT 1 ")
	
	objConexao.Execute("DELETE FROM TB_PRODUCAO_RECIBOS_TEMPORARIA WHERE COD_PRODUCAO="&Request.QueryString("cod_prod"))
		
	FecharConexao()
	
	<!--Response.Redirect("dip_producoes.asp?msg=Produção finalizada com sucesso!&cod_prod="&Request.QueryString("cod_prod"))-->
	
End If

If Request.QueryString("comando") = "excluir_producao" then
'Antes de excluir verifica novamente se a Produção não está finalizada para poder excluir

	Dim rsProdq, PRODUCAO_FINALIZADA_verf
	AbrirConexao()
	
	Set rsProdq = objConexao.Execute("SELECT PRODUCAO_FINALIZADA FROM TB_PRODUCAO WHERE COD_PRODUCAO="&Request.QueryString("cod_prod"))
	
		PRODUCAO_FINALIZADA_verf = rsProdq("PRODUCAO_FINALIZADA")
	
	rsProdq.Close()
	Set rsProdq = Nothing
	
	
	If PRODUCAO_FINALIZADA_verf = 2 then
		
		'
		Dim rsqq
		
		Set rsqq = objConexao.Execute("SELECT COD_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO="&Request.QueryString("cod_prod"))
		
		If not rsqq.Eof then
		'HÃ¡ diversas fichas
			
			While not rsqq.Eof
			'Iniciando dos auxiliares para tr�s, para finalizar excluindo as fichas e a Produção em si.
				'1) Exclui brindes se houver.
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_BRINDES WHERE COD_FORMANDO="&rsqq("COD_FORMANDO"))
				
				'2) Exclui movimentaÃ§Ã£o se houver.
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_FORMANDO="&rsqq("COD_FORMANDO"))
				
				'3) Exclui novo endereÃ§o se houver;
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO WHERE COD_FORMANDO="&rsqq("COD_FORMANDO"))
				
				'4) Exclui produtos agregados se houver;
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS WHERE COD_FORMANDO="&rsqq("COD_FORMANDO"))	
							
				'5) Exclui tabela temporÃ¡ria se houver;
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_TEMPORARIA WHERE COD_FORMANDO="&rsqq("COD_FORMANDO"))			
			
			rsqq.MoveNext()
			Wend
			
			rsqq.Close()
			Set rsqq = Nothing

			'6) Exclui todas as fichas
			objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO="&Request.QueryString("cod_prod"))

			'7) Exclui os produtos agregados da Produção
			objConexao.Execute("DELETE FROM TB_PRODUCAO_PRODUTOS_AGREGADOS WHERE COD_PRODUCAO="&Request.QueryString("cod_prod"))

			'8) Finalmente exclui a Produção
			objConexao.Execute("DELETE FROM TB_PRODUCAO WHERE COD_PRODUCAO="&Request.QueryString("cod_prod")&" LIMIT 1")

			'9) Por precauÃ§Ã£o exclui novamente a TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS, entretanto pelo cÃ³digo da Produção
			objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS WHERE COD_FORMANDO="&Request.QueryString("cod_prod"))	

			Response.Redirect("dip_producoes.asp?msg=Produção excluida com sucesso!")
				
		Else
			'não há fichas, por tanto basta excluir a Produção em si.
			objConexao.Execute("DELETE FROM TB_PRODUCAO WHERE COD_PRODUCAO = "&Request.QueryString("cod_prod")&" LIMIT 1")
			Response.Redirect("dip_producoes.asp?msg=Produção excluida com sucesso!")
		End If
					
	elseif PRODUCAO_FINALIZADA_verf = 1 then
	
		Response.Redirect("dip_producoes.asp?msg=Produção NÃO EXCLUIDA, pois a mesma j� est� finalizada e NÃO poder� ser excluida&cod_prod="&Request.QueryString("cod_prod"))

	End If

	FecharConexao()

End If
%>
<script language="javascript">
	<%If Request.QueryString("comando") = "focuBotao" Then%>
	function focoBotaoCarregar(){
		alterarFormandoFinalizado('<%=NOME_FORMANDO_QS%>','<%=COD_FORMANDO_QS%>')
	}
	<%End If%>
</script>

<h2>DIP - Produ&ccedil;&otilde;es / &Aacute;lbuns </h2>
<%
Dim rs9, PRODUCAO_FINALIZADA, DESCRICAO, VERIFICA_FINALIZACAO

VERIFICA_FINALIZACAO = 1

AbrirConexao()

Set rs9 = objConexao.Execute("SELECT PRODUCAO_FINALIZADA, DESCRICAO FROM TB_PRODUCAO WHERE COD_PRODUCAO="&Request.QueryString("cod_prod"))
If not rs9.Eof then
	PRODUCAO_FINALIZADA = rs9("PRODUCAO_FINALIZADA")
	DESCRICAO = rs9("DESCRICAO")
End If					
rs9.Close()
Set rs9 = Nothing 

'Calcula as quantidades de pré-venda e normais para o menu
	Dim QTD_PRE_VENDA_TOTAL, QTD_NORMAL_TOTAL, TOTAL_FORMANDOS_PROD
	QTD_PRE_VENDA_TOTAL = 0
	QTD_NORMAL_TOTAL = 0
	If PRODUCAO_FINALIZADA = 1 Then
		Set rsPre = objConexao.Execute("SELECT COUNT(DISTINCT F.COD_FORMANDO) as QTD FROM TB_FICHAS_IDENTIFICACAO F INNER JOIN TB_PEDIDOS_FORMANDOS PF ON F.COD_FORMANDO = PF.COD_FORMANDO INNER JOIN TB_PEDIDOS P ON PF.COD_PEDIDO = P.COD_PEDIDO WHERE F.COD_PRODUCAO = " & Request.QueryString("cod_prod") & " AND P.PRE_VENDA = 'S'")
		If Not rsPre.Eof Then
			QTD_PRE_VENDA_TOTAL = CLng(0&rsPre("QTD"))
		End If
		rsPre.Close
		Set rsPre = Nothing
		
		Set rsTotalCount = objConexao.Execute("SELECT COUNT(*) as QTD FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = " & Request.QueryString("cod_prod"))
		If Not rsTotalCount.Eof Then
			TOTAL_FORMANDOS_PROD = CLng(0&rsTotalCount("QTD"))
			QTD_NORMAL_TOTAL = TOTAL_FORMANDOS_PROD - QTD_PRE_VENDA_TOTAL
		End If
		rsTotalCount.Close
		Set rsTotalCount = Nothing
	End If

FecharConexao()
%>

<!-- SUB MENU -->
<div class="subMenu">
Sub-Menu:

<%If PRODUCAO_FINALIZADA  = 2 then%>
	<a href="javascript:modalDesconto(<%=Request.QueryString("cod_prod")%>)">Cadastrar desconto</a> | 
	<a href="dip_producoes_alteracoes.asp?cod_prod=<%=Request.QueryString("cod_prod")%>&comando=alterar">Alterar esta Produ&ccedil;&atilde;o</a> | 
	<a href="javascript:void(0)" onclick="confirmaExclusao()">Excluir esta Produ&ccedil;&atilde;o</a> |
	<a href="javascript:void(0)" onclick="ExclusaoMassa()">Exclus&atilde;o em massa</a> |
	<a href="dip_producoes_importar_formandos.asp?cod_prod=<%=Request.QueryString("cod_prod")%>">Importar Formandos</a> | 
	<a href="dip_formandos_cadastro_completo.asp?cod_prod=<%=Request.QueryString("cod_prod")%>&comando=novo">Cadastrar Formando</a>
<%
elseIf PRODUCAO_FINALIZADA  = 1 then %>
<a href="dip_producoes_fichas_lotes.asp?cod_prod=<%=Request.QueryString("cod_prod")%>">Lotes por vendedor</a> | 
			<% If Request.QueryString("pre_venda") = "S" Then %>
				<a href="dip_producoes_fichas.asp?cod_prod=<%=Request.QueryString("cod_prod")%>&prod_agre=<%=Request.QueryString("prod_agre")%>" style="color:red; font-weight:bold;">[X] Limpar Filtro Pré-venda (<%=QTD_PRE_VENDA_TOTAL%>)</a> | 
			<% Else %>
				<a href="dip_producoes_fichas.asp?cod_prod=<%=Request.QueryString("cod_prod")%>&prod_agre=<%=Request.QueryString("prod_agre")%>&pre_venda=S">Filtrar Pré-venda (<%=QTD_PRE_VENDA_TOTAL%>)</a> | 
			<% End If %>
			<span title="Formandos sem pedidos de pré-venda">Normais (<%=QTD_NORMAL_TOTAL%>)</span> |
			<a href="javascript:void(0)" onclick="imprimirSelecionados()">Imprimir DIGITAL 📸 e FICHAS 📔</a> | <a href="javascript:void(0)" onclick="imprimirSelecionadosOficial()">Imprimir ÁLBUM📒 E FICHAS📔</a> | <a href="javascript:void(0)" onclick="imprimirSelecionadosalbum1000()">Imprimir Álbuns 1000</a> | <a href="javascript:Void(0)" onclick="instrucoesEtiquetas()">Instru&ccedil;&otilde;es para colagem de Etiqueta</a>
<%End If%>
</div>
<!-- FIM SUB MENU -->

<% If PRODUCAO_FINALIZADA  = 1 then %>
<div>
	<table width="98%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td width="4%"><img src="images/icons/lista_exclamation.jpg" width="44" height="42" /></td>
			<td width="96%"><span class="msg"><strong class="msg">Aten&ccedil;&atilde;o</strong>!</span><br />
				Esta Produ&ccedil;&atilde;o est&aacute; finalizada poder&aacute; ser alterada somente com senha especial.
			</td>
		</tr>
	</table>
</div>
<%End If%> 
<div class="msg" align="center"><%=Request.QueryString("msg")%></div>
<%
	AbrirConexao()
	'AbrirConexao2()
		  
	Dim rs, rs2, rs3
	Dim QTD_ALBUNS, FICHAS_VENDA

  set listaInventario = Server.CreateObject("Scripting.Dictionary")
  'BUSCAMOS OS DADOS DA TABELA TEMPORARIA 
  Set rsInventario = objConexao.Execute("SELECT COD_FORMANDO, max(DATA) as data FROM TB_INVENTARIO WHERE COD_FORMANDO IN (SELECT COD_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = "&Request.QueryString("cod_prod")&") GROUP BY COD_FORMANDO")
  if not rsInventario.Eof then
    While not rsInventario.Eof
      call listaInventario.add(rsInventario("cod_formando").value, rsInventario("data").value)
    rsInventario.MoveNext
    Wend
  end if 
  rsInventario.close
  set rsInventario = nothing

	Dim sql_fichas
	If Request.QueryString("pre_venda") = "S" Then
		sql_fichas = "SELECT DISTINCT F.* FROM TB_FICHAS_IDENTIFICACAO F " & _
					 "INNER JOIN TB_PEDIDOS_FORMANDOS PF ON F.COD_FORMANDO = PF.COD_FORMANDO " & _
					 "INNER JOIN TB_PEDIDOS P ON PF.COD_PEDIDO = P.COD_PEDIDO " & _
					 "WHERE F.COD_PRODUCAO = " & Request.QueryString("cod_prod") & " " & _
					 "AND P.PRE_VENDA = 'S' " & _
					 "ORDER BY F.NOME_FORMANDO"
	Else
		sql_fichas = "SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO="&Request.QueryString("cod_prod")&" ORDER BY NOME_FORMANDO"
	End If

	Set rs = objConexao.Execute(sql_fichas)

	'Verifica se o filtro de pré-venda retornou resultados
	If Request.QueryString("pre_venda") = "S" And rs.Eof Then
		%>
		<script>
			alert("Nesta produ\u00e7\u00e3o n\u00e3o h\u00e1 pedidos marcados como Pr\u00e9-venda.");
			location.href = "dip_producoes_fichas.asp?cod_prod=<%=Request.QueryString("cod_prod")%>&prod_agre=<%=Request.QueryString("prod_agre")%>";
		</script>
		<%
		Response.End
	End If


  Set tmp_rs = rs
	If not rs.Eof then
	%>
	<h3 align="center">
		Formandos da Produ&ccedil;&atilde;o:<br /><%=Request.QueryString("cod_prod")%> - <%=DESCRICAO%>
	</h3>
	
	<form action="dip_producoes_fichas_confirma.asp?cod_prod=<%=Request.QueryString("cod_prod")%>&prod_agre=<%=Request.QueryString("prod_agre")%>" method="post" name="producao">

	<table width="98%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
		<tr style="background-color:#333353;; color:white !important;">
			<td width="20%"><div align="center"><strong>Nome Formando </strong></div></td>
			<td width="4%"><div align="center"><strong>C&oacute;d. Form. </strong></div></td>
			<td width="4%"><div align="center"><strong>Contrato</strong></div></td>
			<td width="4%"><div align="center"><strong>Pasta</strong></div></td>
			<td width="4%"><div align="center"><strong>Álbum físico</strong></div></td>
			<td width="4%"><div align="center"><strong>Álbum digital </strong></div></td>
			<td width="4%"><div align="center"><strong>Inicial</strong></div></td>
			<td width="4%"><div align="center"><strong>Comiss&atilde;o</strong></div></td>
			<td width="4%"><div align="center"><strong>Brinde</strong></div></td>
			<td width="4%"><div align="center"><strong>Prod. Agreg. </strong></div></td>
			<td width="4%"><div align="center"><strong>Lote</strong></div></td>
      		<!--<td width="6%"><div align="center"><strong>Data invent&aacute;rio.</strong></div></td>  NÃO ESTÁ MAIS EM USO 22/08/2025--> 
			<!--<% If PRODUCAO_FINALIZADA = 2 Then %>
			<td width="4%"><div align="center"><strong>Gravar (NÃO USE ❌ MANUTENÇÃO)</strong></div></td>
				<% End If %>-->
			<% If PRODUCAO_FINALIZADA = 1 Then %>
				<td width="4%"><div align="center"><strong>Imprimir Individuais</strong></div></td>
				<% End If %>
			<td width="4%"><div align= "center"><strong>Edição</strong></div></td>
		</tr>
		<%
		While not rs.Eof

		Dim INICIAL, COMISSAO, PRODUTO_AGREGADO, VENDIDO, ENCADERNACOES_EXTRAS, LOTE

		INICIAL = 0&rs("INICIAL")
		INICIAL = formatNumber(INICIAL,0)
		VENDIDO = formatNumber(0&rs("VENDIDO"),0)

		COMISSAO =  0&rs("COMISSAO")
		COMISSAO = formatNumber(COMISSAO,0)

		PRODUTO_AGREGADO = 0&rs("PRODUTO_AGREGADO")
		PRODUTO_AGREGADO = formatNumber(PRODUTO_AGREGADO,0)
		ENCADERNACOES_EXTRAS = formatNumber(0&rs("ENCADERNACOES_EXTRAS"),0)

		LOTE = rs("LOTE")

		'/VISUALIZA A GRAVA��O TEMPOR�RIA NA P�G. CONFIRMA��A NA PROX. TELA
		Dim rsTemporaria, rsEntidadeEnsino
		Dim TEMP_ALBUM, TEMP_ALBUM_DIGITAL, TEMP_INICIAL, TEMP_COMISSAO, TEMP_BRINDE, TEMP_COD_VENDEDOR, TEMP_PROD_AGREGADO, TEMP_ENCADERNACOES_EXTRAS, TEMP_LOTE

		TEMP_ALBUM = 0
		TEMP_ALBUM_DIGITAL = 0
		TEMP_ENCADERNACOES_EXTRAS = 0
		TEMP_INICIAL = 0
		TEMP_COMISSAO = 0
		TEMP_BRINDE = 0
		TEMP_PROD_AGREGADO = 0
		TEMP_LOTE = ""
		

		Set rsTemporaria = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO_TEMPORARIA WHERE COD_FORMANDO="&rs("COD_FORMANDO"))
		if not rsTemporaria.Eof then

			TEMP_ALBUM					= rsTemporaria("ALBUM")
			TEMP_ALBUM_DIGITAL			= rsTemporaria("ALBUM_DIGITAL")
			TEMP_INICIAL				= rsTemporaria("INICIAL")
			TEMP_COMISSAO 				= rsTemporaria("COMISSAO")
			TEMP_BRINDE				 	= rsTemporaria("BRINDE")
			TEMP_PROD_AGREGADO	 		= rsTemporaria("PROD_AGREGADO")
			TEMP_ENCADERNACOES_EXTRAS	= rsTemporaria("ENCADERNACOES_EXTRAS")
			TEMP_LOTE					= rsTemporaria("LOTE")
			TEMP_PASTA					= rsTemporaria("numero_pasta_gerencia")

			If int(TEMP_LOTE) = 0 Then
				TEMP_LOTE = ""
			End If
		End If
		rsTemporaria.Close()
		Set rsTemporaria = Nothing



		'Verificmaos se o usuairio solicitou altera��o deste formando
		Dim COD_FORM_STRING, COD_RS_FORMANDO
		COD_FORM_STRING = 0&Request.QueryString("cod_form")
		COD_FORM_STRING = formatNumber(COD_FORM_STRING,0)
		COD_RS_FORMANDO = 0&rs("COD_FORMANDO")
		COD_RS_FORMANDO = formatNumber(COD_RS_FORMANDO,0)

		if COD_FORM_STRING = COD_RS_FORMANDO then
			marcado_edicao =  "bgcolor='#FFFF99' "
		else
			marcado_edicao =  "bgcolor='#ffffff' "
		end if

		%>


			<!--PASTA-->
		<tr <%=marcado_edicao%>id="excluirCelula_<%=rs("COD_FORMANDO")%>">
			<td><%=rs("NOME_FORMANDO")%></td>
			<td><%=digito(rs("COD_FORMANDO"))%></td>
			<td><div align="center"><%=rs("COD_CONTRATO")%></div></td>
			<td><div align="center">
			<%If INICIAL > 0 then
				response.write(rs("numero_pasta_gerencia"))
			Else %>
				<input value="<%=TEMP_PASTA%>" onkeypress="mascara(this,soNumeros)" style="width:40px" type="text" name="PASTA_<%=rs("COD_FORMANDO")%>" />
			<%End If%>
			</div></td>
			<!-- ALBUM / DUPLO-->
			<td><div align="center">
			<%
			If INICIAL > 0 then
				tmp_comando = "visualizar"
				response.write(rs("ALBUM"))
				If ENCADERNACOES_EXTRAS = 1 Then %>
					<strong> DUPLO</strong>
				<%End If
			Else
				VERIFICA_FINALIZACAO = 2
				tmp_comando = "incluir"

				If  TEMP_ALBUM <= 0 Then
					TEMP_ALBUM = ""
				End If		

				IF TEMP_INICIAL <= 0 Then
					TEMP_INICIAL = ""
				End If
					
				IF PRODUCAO_FINALIZADA = 1 Then
					TEMP_ALBUM = ""
					TEMP_INICIAL = ""
				End If

				%>
	
		 		<input value="<%=TEMP_ALBUM%>" style="width:40px" onkeypress="mascara(this,soNumeros)" type="text" name="ALBUM_<%=rs("COD_FORMANDO")%>" />

           		<!--<input value="1" class="inputRadio" <%If TEMP_ENCADERNACOES_EXTRAS = 1 Then%> checked="checked" <%End If%> type="checkbox" name="ENCADERNACOES_EXTRAS_<%=rs("COD_FORMANDO")%>" />-->
			<% End If %>
            </div></td>

				<!-- ÁLBUM DIGITAL -->
					<td>
					<div align="center">
						<%
						If INICIAL > 0 Then
							ALBUM_DIGITAL = rs("ALBUM_DIGITAL")
							' Se já inicializado, mostra apenas SIM ou NÃO
							If ALBUM_DIGITAL = 1 Then
								response.write("SIM")
							Else
								response.write("NÃO")
							End If
						Else
							' Exibe a caixinha para marcar/desmarcar
							Dim checkedAlbum
							If TEMP_ALBUM_DIGITAL = 1 Then
								checkedAlbum = "checked"
							Else
								checkedAlbum = ""
							End If

							response.write("<input type='checkbox' name='ALBUM_DIGITAL_" & rs("COD_FORMANDO") & "' value='1' " & checkedAlbum & " codform='" & rs("COD_FORMANDO") & "'>")
						End If
						%>
					</div>
					</td>



			<!--INICIAL-->
            <td><div align="center">
             	<%If INICIAL > 0 then
					response.write(INICIAL)
				Else%>
              		<input value="<%=TEMP_INICIAL%>" onkeypress="mascara(this,soNumeros)" style="width:40px" type="text" name="INICIAL_<%=rs("COD_FORMANDO")%>" />
				<% End If %>
            </div></td>
			<!--COMISSÃO-->
            <td><div align="center">
            	<%
            	if tmp_comando = "" then
            		tmp_comando = "visualizar"
            	end if 

            	If INICIAL > 0 then
					If COMISSAO = 1 then
						response.write("SIM")
					Else
						response.write("N&Atilde;O")
					End If
				Else

            	response.write(padraoFormularioSelect("COMISSAO_"&rs("cod_formando"),"","modal_comissao","","codform='"&rs("COD_FORMANDO")&"'","SEM_SELECIONE", TEMP_COMISSAO, "PADRAO:2#N&Atilde;O|1#SIM" ,tmp_comando, false,false,false))
				end if
				%>
            </div>
            </td>
			<!--BRINDE-->
            <td><div align="center">
			<%
			If INICIAL > 0 then
				If COMISSAO = 1 then
					response.write("Sim")
				Else
					response.write("N&Atilde;O")
				End If
			Else
			%>
				<table border="0" align="center" cellpadding="0" cellspacing="0">
					<tr>
						<td width="68%">
							<select class='modal_brinde' codform='<%=rs("COD_FORMANDO")%>' name="BRINDE_<%=rs("COD_FORMANDO")%>" >
	                			<option value="2" <%If TEMP_BRINDE <> 1 then%>selected="selected"<%End If%>>N&Atilde;O</option>
	                			<option value="1" <%If TEMP_BRINDE = 1 then%>selected="selected"<%End If%>>SIM</option>
	              			</select>
	              		</td>
						<td width="32%">
							<div align="center" id="listaBrindes">
								<a href="#" codform='<%=rs("COD_FORMANDO")%>' class='btn_brinde' tabindex="9999">
									<img src="images/icons/icon_bloco_notas.gif" width="24" height="29" border="0" />
								</a>
							</div>
						</td>
					</tr>
				</table>
			<%End if%>
            </div></td>
			<!--PRODUTO AGREGADO-->
            <td><div align="center">
			<%
			If int(Request.QueryString("prod_agre")) = 1 then
				If int(INICIAL) > 0 then
					If int(PRODUTO_AGREGADO) = 1 then
						response.write("Sim")
					Else
						response.write("N&Atilde;O")
					End If
				Else %>
   					<select name="PRODUTO_AGREGADO_<%=rs("COD_FORMANDO")%>">
						<option value="2" <%If int(TEMP_PROD_AGREGADO) = 2 then%>selected="selected"<%End If%>>N&Atilde;O</option>
				  		<option value="1" <%If int(TEMP_PROD_AGREGADO) = 1 or TEMP_PROD_AGREGADO = 0 then%>selected="selected"<%End If%>>SIM</option>
				  	</select>
				<% End if
			else %>
				<input name="PRODUTO_AGREGADO_<%=rs("COD_FORMANDO")%>" type="hidden" value="2" />
				N&Atilde;O
			<%End If%>
            </div></td>
			<!--LOTE-->
            <td><div align="center">
             	<%If INICIAL > 0 then
					response.write(LOTE)
				Else %>
					<input value="<%=TEMP_LOTE%>" onkeypress="mascara(this,soNumeros)" style="width:40px" type="text" name="LOTE_<%=rs("COD_FORMANDO")%>" />
				<%End If%>
            </div></td>

			<!--DATA INVENTARIO
        <td><div align="center">
        <%
        if PRODUCAO_FINALIZADA = 1 and listaInventario.exists(rs("COD_FORMANDO").value) then
					response.write(listaInventario(rs("COD_FORMANDO").value))
        else
          response.write(" - ")
        end if
				%>
        </div></td> -->


<html>
<body>


		<!-- BOTÃO SALVAR POR FORMANDO
		 <% If PRODUCAO_FINALIZADA = 2 Then %>
			<td>
			<div align="center">
				<form method="post" action="" style="margin:0;" id="formSalvar_<%=rs("COD_FORMANDO")%>">
				<input type="hidden" name="acao" value="salvar">
				<input type="hidden" name="cod_formando" value="<%=rs("COD_FORMANDO")%>">
				<input type="hidden" name="PASTA_<%=rs("COD_FORMANDO")%>"               id="h_PASTA_<%=rs("COD_FORMANDO")%>">
				<input type="hidden" name="ALBUM_<%=rs("COD_FORMANDO")%>"               id="h_ALBUM_<%=rs("COD_FORMANDO")%>">
				<input type="hidden" name="ENCADERNACOES_EXTRAS_<%=rs("COD_FORMANDO")%>"id="h_ENC_<%=rs("COD_FORMANDO")%>">
				<input type="hidden" name="ALBUM_DIGITAL_<%=rs("COD_FORMANDO")%>"       id="h_DIG_<%=rs("COD_FORMANDO")%>">
				<input type="hidden" name="INICIAL_<%=rs("COD_FORMANDO")%>"             id="h_INICIAL_<%=rs("COD_FORMANDO")%>">
				<input type="hidden" name="COMISSAO_<%=rs("COD_FORMANDO")%>"            id="h_COM_<%=rs("COD_FORMANDO")%>">
				<input type="hidden" name="BRINDE_<%=rs("COD_FORMANDO")%>"              id="h_BRI_<%=rs("COD_FORMANDO")%>">
				<input type="hidden" name="PRODUTO_AGREGADO_<%=rs("COD_FORMANDO")%>"    id="h_AGR_<%=rs("COD_FORMANDO")%>">
				<input type="hidden" name="LOTE_<%=rs("COD_FORMANDO")%>"                id="h_LOT_<%=rs("COD_FORMANDO")%>">

				<input type="button" value="MANUTENÇÃO" class="btnSalvar"
						onclick="prepararESalvar(<%=rs("COD_FORMANDO")%>)">
				</form>
			</div>
			</td>
			<% End If %>
			<script>
function getByName(name) {
  var els = document.getElementsByName(name);
  return els && els.length ? els[0] : null;
}
function valOf(name) {
  var el = getByName(name);
  if (!el) return "";
  if (el.type === "checkbox") return el.checked ? "1" : "0";
  return el.value || "";
}
function prepararESalvar(cod) {
  // Lê dos inputs visíveis da linha
  document.getElementById("h_PASTA_" + cod).value = valOf("PASTA_" + cod);
  document.getElementById("h_ALBUM_" + cod).value = valOf("ALBUM_" + cod);
  document.getElementById("h_ENC_"   + cod).value = valOf("ENCADERNACOES_EXTRAS_" + cod);
  document.getElementById("h_DIG_"   + cod).value = valOf("ALBUM_DIGITAL_" + cod);
  document.getElementById("h_INICIAL_" + cod).value = valOf("INICIAL_" + cod);
  document.getElementById("h_COM_"     + cod).value = valOf("COMISSAO_" + cod);
  document.getElementById("h_BRI_"     + cod).value = valOf("BRINDE_" + cod);
  document.getElementById("h_AGR_"     + cod).value = valOf("PRODUTO_AGREGADO_" + cod);
  document.getElementById("h_LOT_"     + cod).value = valOf("LOTE_" + cod);

  // Envia apenas os dados dessa linha
  document.getElementById("formSalvar_" + cod).submit();
}
</script>
<style>
.btnSalvar {
  background-color: #ff0080; /* rosa forte */
  border: none;
  color: white;
  padding: 10px 10px;
  text-align: center;
  text-decoration: none;
  display: inline-block;
  font-size: 10px;
  font-weight: bold;
  border-radius: 8px; /* cantos arredondados */
  cursor: pointer;
  transition: background-color 0.3s ease;
}

.btnSalvar:hover {
  background-color: #e60073; /* tom mais escuro no hover */
}
</style>

MANUTENÇÃO 17/09/2025 -->
 

<!-- IMPRIMIR ETIQUETA -->
<% If PRODUCAO_FINALIZADA = 1 Then %>
<td>
<div align="center">
    <%
        response.write("<input type='checkbox' class='chkEtiqueta' " & _
                       "data-codform=" & rs("COD_FORMANDO") & " " & _
                       "value='1'>")
    %>
</div>
</td>
<% End If %>




			<!--EDICÃO-->
			<td><%If PRODUCAO_FINALIZADA = 2 THEN%>
			<table width="50" border="0" align="center" cellpadding="2" cellspacing="0">
				<tr>
					<td width="25"><div align="center">
					<%If INICIAL > 0 then%>
						<a href="javascript:alterarFormando('<%=rs("NOME_FORMANDO")%>','<%=rs("COD_FORMANDO")%>')" tabindex="98777"><img src="images/icons/icon_alterar.png" width="16" height="16" border="0" /></a>
					<%End If%>
					</div></td>
					<td width="25">
						<div align="center">
							<div style="border: solid 1px #999999; padding:2px">			  
					  			<a href="javascript:deletaFormando('<%=rs("NOME_FORMANDO")%>','<%=rs("COD_FORMANDO")%>')" tabindex="999999">
	                  				<img src="images/icons/icon_excluir.png" border="0" />
	                  			</a>
								<input name="EXCLUSAO_EM_MASSA_<%=rs("COD_FORMANDO")%>" onclick="javascript:exclusaoSeleciona('<%=rs("COD_FORMANDO")%>')" class="inputRadio" type="checkbox" value="1" />
							</div>
                		</div>
						<input style="width:35px" name="MODIFICACOES_<%=rs("COD_FORMANDO")%>" value="0" type="hidden" />
					</td>
				</tr>
			</table>	
			<%Elseif PRODUCAO_FINALIZADA = 1 AND VENDIDO = 2 AND INICIAL > 0 Then%>
				<div align="center">
					<a href="dip_busca_detalhamento_formando.asp?cod_formando=<%=rs("COD_FORMANDO")%>&msg=Para alterar o formando ap&oacute;s a finaliza&ccedil;&atilde;o da produ&ccedil;&atilde;o n&atilde;o necess&aacute;rio ter valida&ccedil;&atilde;o especial.<br />Se voc&ecirc; tiver a valdia&ccedil;&atilde;o especial a op&ccedil;&atilde;o Corre&ccedil;&atilde;o de Formando abaixo estar&aacute; dispon&iacute;vel." tabindex="98777">
						<img src="images/icons/icon_alterar_block.png" width="16" height="18" border="0"  />
					</a>            
				</div>
			    <div align="center">
				<%Elseif PRODUCAO_FINALIZADA = 1 AND VENDIDO = 2 AND INICIAL = 0 Then%>
					<input style="width:35px" name="MODIFICACOES_<%=rs("COD_FORMANDO")%>" value="9" type="hidden" />
				<%Elseif PRODUCAO_FINALIZADA = 1 AND VENDIDO = 1 Then%>
					&Aacute;lbum vendido e corrigido
				<%End If%>			
	            </div>
			</td>
			
			

			</tr>
			<%
			rs.MoveNext()
			Wend

			rs.Close()
			Set rs = Nothing
			%>
			<tr>
				<td colspan="9">
					<div align="center">
					<%If VERIFICA_FINALIZACAO = 1 and PRODUCAO_FINALIZADA = 2 then%>
						<input class="btnFinalizaProducao" onclick="confirmaFinalizarProducao()" type="button" name="Submit22" value=" " />
					<%Elseif VERIFICA_FINALIZACAO = 2 then%>
						<input class="btnAvancar" type="submit" name="Submit2" value=" " />
					<%End If%>
					</div>
				</td>
				<td colspan="4">
				<%If PRODUCAO_FINALIZADA = 2 Then%>
					<div align="center"><a href="javascript:excluirRegistrosemMassa()">Excluir Registros Selecionados ❌</a></div>
				<%End If%>
				</td>
			</tr>
		</table>
	</form>
<%else%>
	<div align="center">
		<p>
			<em>N&atilde;o h&aacute; formando importados!</em>
		<p>
			<a href="dip_producoes_importar_formandos.asp?cod_prod=<%=Request.QueryString("cod_prod")%>">Deseja importar agora?</a>
		</p>
	</div>    
<% end if

Dim RESUMO_QTD_FORMANDOS, RESUMO_QTD_COMISSOES, RESUMO_QTD_FOTOS
Dim RESUMO_QTD_FORMANDOS2, RESUMO_QTD_COMISSOES2, RESUMO_QTD_FOTOS2

%>
<table width="500" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
	<tr bgcolor="whitesmoke">
    	<td colspan="4">
    		<div align="center" class="textoMedio">
    			<strong>Resumo da divis&atilde;o por lotes</strong>
    		</div>
    	</td>
	</tr>
    <tr>
		<td width="98"><div align="center"><strong>Lote</strong></div></td>
		<td width="130"><div align="center"><strong>Qtd. de Formandos</strong></div></td>
		<td width="133"><div align="center"><strong>Qtd. Co</strong></div></td>
		<td width="79"><div align="center"><strong>Qtd. Fotos</strong></div></td>
	</tr>
	
	<%
	
	Dim total_fotos, total_formandos, total_i
	Dim array_formandos(), array_co(), array_fotos()
	total_fotos = 0
	total_formandos = 0
	total_i = 0

	For i=1 to 500 'Verifica em 500 lotes
		RESUMO_QTD_FORMANDOS = 0
		RESUMO_QTD_COMISSOES = 0
		RESUMO_QTD_FOTOS = 0
		RESUMO_QTD_FORMANDOS2 = 0
		RESUMO_QTD_COMISSOES2 = 0
		RESUMO_QTD_FOTOS2 = 0
					 
		'PARTE 1 TEMPOR�RIA comissao			  
		Set rs=objConexao.Execute("SELECT COUNT(*) as comissoes FROM TB_FICHAS_IDENTIFICACAO INNER JOIN  TB_FICHAS_IDENTIFICACAO_TEMPORARIA ON TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO_TEMPORARIA.COD_FORMANDO WHERE TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO = "&Request.QueryString("cod_prod")&" and TB_FICHAS_IDENTIFICACAO_TEMPORARIA.LOTE = '"&i&"' AND TB_FICHAS_IDENTIFICACAO_TEMPORARIA.COMISSAO = 1 GROUP BY TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO")
		If not rs.Eof Then
			RESUMO_QTD_COMISSOES =	(formatNumber(0&rs("comissoes"),0))/1
		End If
		rs.Close
		Set rs=Nothing

		'PARTE 1 TEMPOR�RIA SEM COMISS�O DE FORMATURA
		Set rs=objConexao.Execute("SELECT Sum(TB_FICHAS_IDENTIFICACAO_TEMPORARIA.INICIAL) as TEMP_QTD_FOTOS, COUNT(*) as TEMP_QTD_FORMANDOS FROM TB_FICHAS_IDENTIFICACAO INNER JOIN  TB_FICHAS_IDENTIFICACAO_TEMPORARIA ON TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO_TEMPORARIA.COD_FORMANDO WHERE TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO = "&Request.QueryString("cod_prod")&" and TB_FICHAS_IDENTIFICACAO_TEMPORARIA.LOTE = "&i&" GROUP BY TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO")
		If not rs.Eof Then

			RESUMO_QTD_FORMANDOS = (formatNumber(0&rs("TEMP_QTD_FORMANDOS"),0))/1
			RESUMO_QTD_FOTOS	 = (formatNumber(0&rs("TEMP_QTD_FOTOS"),0))/1
		End If
		rs.Close
		Set rs=Nothing
			  
		'PARTE 2 oficial sem duplicidade
		Set rs=objConexao.Execute("SELECT sum(inicial) as qtd_inicial,count(*) as qtd_formandos,sum(IF(COMISSAO=1,1,0)) as total_comissao FROM TB_FICHAS_IDENTIFICACAO WHERE cod_producao = "&Request.QueryString("cod_prod")&" and inicial > 0 and lote = "&i)
		
		
		
		If not rs.Eof Then

			RESUMO_QTD_FORMANDOS2 = (formatNumber(0&rs("qtd_formandos"),0))/1
			RESUMO_QTD_FOTOS2	  = (formatNumber(0&rs("qtd_inicial"),0))/1
			RESUMO_QTD_COMISSOES2 = (formatNumber(0&rs("total_comissao"),0))/1
		End If
		rs.Close
		Set rs=Nothing
			  
		RESUMO_QTD_COMISSOES = (RESUMO_QTD_COMISSOES + RESUMO_QTD_COMISSOES2)/1
		RESUMO_QTD_FORMANDOS = (RESUMO_QTD_FORMANDOS + RESUMO_QTD_FORMANDOS2)/1
		RESUMO_QTD_FOTOS	   = (RESUMO_QTD_FOTOS + RESUMO_QTD_FOTOS2)/1
			  
		if int(RESUMO_QTD_FORMANDOS) > 0 Then
		%>
			<tr>
				<td><div align="center" class="textoMedio"><%=i%></div></td>
				<td><div align="center"><%=RESUMO_QTD_FORMANDOS%></div></td>
				<td><div align="center"><%=RESUMO_QTD_COMISSOES%></div></td>
				<td><div align="center"><%=RESUMO_QTD_FOTOS%></div></td>
			</tr>
		<%
			total_formandos = total_formandos + RESUMO_QTD_FORMANDOS
			total_fotos = total_fotos + RESUMO_QTD_FOTOS
			total_i = total_i + i

			ReDim PRESERVE array_formandos(i)
			ReDim PRESERVE array_co(i)
			ReDim PRESERVE array_fotos(i)

			array_formandos(i) = RESUMO_QTD_FORMANDOS
			array_co(i) = RESUMO_QTD_COMISSOES
			array_fotos(i) = RESUMO_QTD_FOTOS

		End If
	Next
	FecharConexao()
	%>
	<tfoot style="background-color: #cecece;">
		<tr>
			<td colspan="2">
				Total de Formandos: <%= total_formandos %>
			</td>
			<td colspan="2">
				Total de Fotos: <%= total_fotos %>
			</td>
		</tr>
		<tr>
			<td colspan="4">
				<form action="dip_producoes_fichas_impressao.php" method="post" accept-charset="utf-8" target="_blank">
					<% 
						On Error Resume Next
						For i=1 to uBound(array_formandos) 
					%>
					<input type="hidden" name="formandos[<%= (i-1) %>]" value="<%= array_formandos(i) %>">
					<input type="hidden" name="co[<%= (i-1) %>]" value="<%= array_co(i) %>">
					<input type="hidden" name="fotos[<%= (i-1) %>]" value="<%= array_fotos(i) %>">
					<%
						Next
						If Err.Number <> 0 Then
							Response.write("Oops!, Erro " & Err.Description)
						End If
					%>
					<input type="hidden" name="nome" value="<%=Request.QueryString("cod_prod")%> - <%=DESCRICAO%>">
					<input type="hidden" name="total_formandos" value="<%= total_formandos %>">
					<input type="hidden" name="total_fotos" value="<%= total_fotos %>">
					<input type="submit" class="btn-sm btn-default btn-block text-center" value="Imprimir Resumo da divis&atilde;o por lotes">
				</form>

			</td>
		</tr>
	</tfoot>
</table>
<p>
	Observa&ccedil;&otilde;es:<br />
	Digito 0 = Formando importado da formatura<br />
	Digito 1 = Formando cadastrado pelo depto. de foto
</p>
<p align="center">
	<input class="btnTopo" onclick="abre('#')" type="button" name="btnTopo2" value=" " />
	<input type="hidden" id='cod_prod' value="<%=Request.QueryString("cod_prod")%>">
</p>

<!-- MODAL DE A�OES -->
<div class="modal fade" id='modalProducaoDiv'>
  <div class="modal-dialog">
    <div class="modal-content">
      <div class="modal-header">
        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
        <h4 class="modal-title">Cadastro de desconto</h4>
      </div>
	  <div class="modal-body">
	  <% if validaUsuario (2,,81,false) then%>
	    <iframe src="dip_producoes_descontos.asp" frameborder='0' width="100%" height="390px" id='iframeProducao'></iframe>
	  <% else %>
	   	 Voc&ecirc; n&atilde;o tem acesso a esta fun&ccedil;&atilde;o
	  <% end if %>
	  </div>
	  <div class="modal-footer"></div>
    </div><!-- /.modal-content -->
  </div><!-- /.modal-dialog -->
</div><!-- /.modal -->

<style type="text/css">
.btnFinalizaProducao{
	background-image:url(images/layout/btnFinalizaProducao.jpg);
	cursor:pointer;
	width:140px;
	height:22px;
	border:0px;
}
#mainContainer{
	width:500px;
	margin:0 auto;
	margin-top:10px;
	border:1px double #000;
	padding:3px;
}
.modalDialog_contentDiv_error{
	border:3px solid #FFF;	
	padding:2px;
	z-index:100;/* Always needed	*/
	position:absolute;	/* Always needed	*/
	background-color:#F00;	/* White background color for the message */
	color:#FFF;
}
.modalDialog_contentDiv_error a{
	color:#FFF;
}
h1{
	margin-top:0px;
}
#scroll {
	height:400px;
	background-color:#FFFFFF;
	overflow:auto;
	padding: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
	margin-left: 0px;
}
#exclusaoemMassa{
	display: none;
	visibility: hidden;
}

</style>

<%' SCRIPTS DA PAGINA %>
<script type="text/javascript">
	function modalDesconto(cod_producao)
	{
		$(".modal-title").html("Cadastro de desconto");
		$("#iframeProducao").prop("src", "dip_producoes_descontos.asp?cod_producao="+cod_producao);
		$("#modalProducaoDiv").modal('show');
	}
	function modalBrinde(cod_prod, cod_form, brinde)
	{
		$(".modal-title").html("DIP - Produções / Álbuns / Brindes");
		$("#iframeProducao").prop("src", "dip_producoes_fichas_POPUP_brindes.asp?cod_form="+cod_form+"&cod_prod="+cod_prod+"&brinde="+brinde);
		$("#modalProducaoDiv").modal('show');
	}
	$(function(){
		$(".modal-title").html("DIP - Produções / Álbuns / Brindes");
		$(".modal_brinde").on("change",function(){
			var cod_form = $(this).attr("codform")
			var cod_prod = $("#cod_prod").val()
			var brinde = $(this).val()

			if(brinde == '1')
				modalBrinde(cod_prod, cod_form, brinde);

		});
		$(".modal_comissao").on("change",function(){
			var cod_form = $(this).attr("codform")

			if($(this).val() == "1")
	 			$("select[name=BRINDE_"+cod_form+"]").val(1).trigger("change");
		});
		$(".btn_brinde").on("click", function(event){
			event.preventDefault()

			var cod_form = $(this).attr("codform")
			var cod_prod = $("#cod_prod").val()
			var brinde = $("select[name=BRINDE_"+cod_form+"]").val()

			if(brinde == '1')
				modalBrinde(cod_prod, cod_form, brinde);
			else
				alert('Para visualizar os brindes é necessário alterar a opção Brinde para Sim.');
		});
	});
</script>
	
<script language="javascript">
	function focusFormando(){
		document.producao.<%="ALBUM_"&Request.QueryString("cod_form")%>.focus();
		d = document.producao.<%="ALBUM_"&Request.QueryString("cod_form")%>;
	}
</script>
<script language="javascript">
	function confirmaFinalizarProducao(){
		var pergunta = confirm('Tem certeza que deseja finalizar esta Produção?\nObs: Esta Produção NÃO poderá ser alterada posteriormente somente com senha especial.')
		if(pergunta==true){
			abre('?cod_prod=<%=Request.QueryString("cod_prod")%>&comando=finalizar_producao')
		}
	}
</script>
<script language="javascript">
	function confirmaExclusao(){
		var perguntaq = confirm('EXCLUIR Produção:\Atenção caso queira continuar lembre-se que todas as fichas(se houver) serão excluidas completamente, deseja continuar?')
		if (perguntaq==true){
			var perguntadeNovo = confirm("CONFIRMA EXCLUSÃO DESTA Produção E SUAS RESPECTIVAS FICHAS?")
			if (perguntadeNovo==true){
			abre('?cod_prod=<%=Request.QueryString("cod_prod")%>&comando=excluir_producao')
			}
		}
	}
</script>
<script language="javascript">
	function deletaFormando(nomeFormando,codFormando){
		var mensagem = "Tem certeza que deseja excluir a ficha do(a) formando(a) "+nomeFormando+"?"

		if(YNconfirm(mensagem)){
			document.producao["MODIFICACOES_"+codFormando].value = "2";
			document.producao.submit();
		}
	}
</script>
<script language="javascript">
	function alterarFormando(nomeFormando,codFormando){
		var mensagem = "Tem certeza que deseja resetar a ficha do(a) formando(a) "+nomeFormando+"?\n*Desta maneira a ficha volta a forma inicial."

		if(YNconfirm(mensagem)){
			document.producao["MODIFICACOES_"+codFormando].value = "1";
			document.producao.submit();
		}
	}
</script>
<script language="javascript">
	function instrucoesEtiquetas(){
		var caminho
		caminho = "<table cellpadding=\'0\' cellspacing=\'0\' border=\'0\' width=\'100%\'><tr><td><table cellpadding=\'0\' cellspacing=\'0\' width=\'100%\'><tr><td width=\'100%\'><h1>Instru&ccedil;&otilde;es para colagem da Etiqueta</h1></td><td align=\'right\'><a href=\'#\' onclick=\'closeMessage();return false\'>Fechar Janela</a></td></tr></table></td></tr><tr><td><iframe frameborder=\'0\' scrolling=\'no\' height=\'380px\' width=\'99%\' src=\'dip_producoes_fichas_POPUP_instrucoes_etiqueta.asp?cod_prod=<%=Request.QueryString("cod_prod")%>\'></iframe></td></tr></table>"
		displayStaticMessage(caminho,false,420)		
	}  
</script>

<script>
function imprimirSelecionados() {
    var selecionados = [];
    document.querySelectorAll(".chkEtiqueta:checked").forEach(function(chk) {
        selecionados.push(chk.getAttribute("data-codform"));
    });

    

    var codProd = "<%=Request.QueryString("cod_prod")%>";
    var desc = "<%=Left(DESCRICAO,23)%>";
    var url = "teste_imprimir_etiquetas_individuais.asp?cod_prod=" + codProd +
              "&desc=" + desc +
              "&formandos=" + selecionados.join(",");

    abrir(url, "teste", "width=200,height=200,scrollbars=no");
}
</script>
<script>
function imprimirSelecionadosOficial() {
    var selecionados = [];
    document.querySelectorAll(".chkEtiqueta:checked").forEach(function(chk) {
        selecionados.push(chk.getAttribute("data-codform"));
    });

    var codProd = "<%=Request.QueryString("cod_prod")%>";
    var desc = "<%=Left(DESCRICAO,23)%>";

    var url = "dip_producoes_fichas_etiquetas2.asp?cod_prod=" + codProd +
              "&desc=" + desc +
              "&formandos=" + selecionados.join(",");

    abrir(url, "oficial", "width=200,height=200,scrollbars=no");
}
</script>
<script>
function imprimirSelecionadosalbum1000() {
    var selecionados = [];
    document.querySelectorAll(".chkEtiqueta:checked").forEach(function(chk) {
        selecionados.push(chk.getAttribute("data-codform"));
    });

    if (selecionados.length === 0) {
        alert("Selecione pelo menos um formando para imprimir!");
        return;
    }

    var codProd = "<%=Request.QueryString("cod_prod")%>";
    var desc = "<%=Left(DESCRICAO,23)%>";

    // Pergunta ao usuário o tipo de impressão
    var opcao = prompt("Digite a opção desejada:\n1 = Ficha + Digital\n2 = Ficha + Álbum");

    var tipo = "";
    if (opcao === "1") {
        tipo = "ficha_digital";
    } else if (opcao === "2") {
        tipo = "ficha_album";
    } else {
        alert("Opção inválida. Operação cancelada.");
        return;
    }

    var url = "dip_producoes_fichas_etiquetas_album1000.asp?cod_prod=" + codProd +
              "&desc=" + desc +
              "&formandos=" + selecionados.join(",") +
              "&tipo=" + tipo;

    abrir(url, "album1000", "width=200,height=200,scrollbars=no");
}
</script>





<script language="javascript">
	function alterarFormandoFinalizado(NOME_FORMANDO,COD_FORMANDO){
		var caminho = "<table cellpadding=\'0\' cellspacing=\'0\' border=\'0\' width=\'100%\'><tr><td><table cellpadding=\'0\' cellspacing=\'0\' width=\'100%\'><tr><td width=\'70%\'><h1>Senha especial para fazer a altera&ccedil;&atilde;o</h1></td><td align=\'right\'><a href=\'#\' onclick=\'closeMessage();return false\'>Fechar Janela</a></td></tr></table></td></tr><tr><td><iframe frameborder=\'0\' scrolling=\'no\' height=\'175px\' width=\'99%\' src=\'dip_producoes_fichas_POPUP_confirmacao_alteracao_senha_especial.asp?cod_producao=<%=Request.QueryString("cod_prod")%>&prod_agregado=<%=Request.QueryString("prod_agre")%>&cod_formando="+COD_FORMANDO+"&nome_formando="+NOME_FORMANDO+"\'></iframe></td></tr></table>"
		displayStaticMessage(caminho,false,200)
	}
</script>
<script language="javascript">
	function confirma(cod_formando,nomeFormando,comando){
	var mensagem

		if (comando == "alterar"){
			mensagem = "Atenção ALTERAÇÃO: \nDeseja resetar a ficha de: "+nomeFormando+" deseja continuar?\n*A ficha voltará na forma original."	
		}	
		else {
			mensagem = "Atenção EXCLUSÃO: \nSe vecê continuar poderá perder as alterações feitas nesta página,\nalém de resetar a ficha do formando: "+nomeFormando+"  deseja continuar?"
		}

		var resposta=confirm(mensagem);
			if (resposta==true) {
				
				if (comando == "excluir"){
					//Nova confirmaÃ§Ã£o, Exclui
					var resposta2=confirm("Deseja realmente EXCLUIR?")
						if (resposta2==true){
							document.producao["MODIFICACOES_"+cod_formando].value = "2";
							document.producao.submit();
						}	
				}
				else{
					//Altera
					document.producao["MODIFICACOES_"+cod_formando].value = "1"
					document.producao.submit();
				}
				
			}
	}
</script>
<script language="javascript">
	function confirmaFinalizado(cod_formando,nomeFormando){
		var mensagem = "ATEN&CCEDIL;&ATILDE;O EXCLUS&ATILDE;O AP&Aacute;S FINALIZA&ccedil;&Atilde;O: \nDeseja resetar a ficha de: "+nomeFormando+" e voltar a forma original?"
		if(YNconfirm(mensagem)){
			//location.href = 'dip_producoes.asp?cod_form='+
		}
	}
	</script>



<script language="javascript">
	function exclusaoSeleciona(COD_FORMANDO){

		var d = document.producao["EXCLUSAO_EM_MASSA_"+COD_FORMANDO]

		if(d.checked){
			document.getElementById('excluirCelula_'+COD_FORMANDO).style.backgroundColor ='#FFA6A6';
		}else {
			document.getElementById('excluirCelula_'+COD_FORMANDO).style.backgroundColor ='';
		}
	}
</script>
<script language="javascript">
	function excluirRegistrosemMassa(){
		if(YNconfirm('Aten&ccedil;&atilde;o tem certeza que deseja excluir os registros selecionados?\nSALVE OS REGISTROS ANTES DE CONTINUAR!')){
			document.producao.action='dip_producoes_fichas_exclusao_em_massa.asp?cod_prod=<%=Request.QueryString("cod_prod")%>&prod_agre=<%=Request.QueryString("prod_agre")%>';
			document.producao.submit();
		}
	}
</script>
<script language="javascript">
	function abrirEtiquetaIndividual() {
		var cod_prod = "<%=Request.QueryString("cod_prod")%>";
		
		var conteudo = "<table cellpadding='0' cellspacing='0' border='0' width='100%'>" +
			"<tr><td>" +
			"<table cellpadding='0' cellspacing='0' width='100%'>" +
			"<tr>" +
			"<td width='70%'><h1>Imprimir Etiquetas Individuais</h1></td>" +
			"<td align='right'><a href='#' onclick='closeMessage();return false'>Fechar Janela</a></td>" +
			"</tr>" +
			"</table>" +
			"</td></tr>" +
			"<tr><td style='padding: 20px'>" +
			"<label for='qtd_etiquetas'>Quantidade de etiquetas (1 a 2000):</label><br>" +
			"<input type='number' id='qtd_etiquetas' name='qtd_etiquetas' min='1' max='2000' value='1' style='width: 100px; padding: 5px; margin-top: 10px;'><br><br>" +
			"<button onclick='confirmarEtiquetaIndividual(" + cod_prod + ")' style='padding: 8px 15px;'>Imprimir</button>" +
			"</td></tr>" +
			"</table>";

		displayStaticMessage(conteudo, false, 300);
	}

	function confirmarEtiquetaIndividual(cod_prod) {
		var qtd = document.getElementById("qtd_etiquetas").value;
		qtd = parseInt(qtd);

		if (isNaN(qtd) || qtd < 1 || qtd > 2000) {
			alert("Digite um número válido entre 1 e 2000.");
			return;
		}

		'var url = "teste_imprimir_etiquetas_individuais.asp?imprimir=sim&quantidade=" + cod_prod + "&qtd=" + qtd;
		var cod = Request.QueryString("cod_prod");
		var desc = Left(DESCRICAO,23);
		'var url = "teste_imprimir_etiquetas_individuais.asp?cod_prod=" & cod & "
		window.open(url, "_blank");
		closeMessage();
	}
</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->
<%
endtime = Timer() 
Response.Write "O carregamento se completou em " & endtime-starttime & " segundos = " 
Response.Write " (" & (endtime-starttime)*1000 & " milésimos de segundos)."
%>