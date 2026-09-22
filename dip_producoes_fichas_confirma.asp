<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

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

<%
DIM ERROR_FLAG, QTD_ERROR, QTD_CORRETOS, VERIFICA_MODIFICACOES, COD_FORMANDO_ALTERADO, NOME_FORMANDO_ALTERADO, STATUS_MODIFICACAO
VERIFICA_MODIFICACOES = 0
QTD_ERROR = 0
QTD_CORRETOS = 0






AbrirConexao()
	'Busca a foto padr�o
	Dim rsPadrao, COD_FOTO_PADRAO, NOME_FOTO_PADRAO
				
	sql =  	"SELECT TB_PRODUTOS.NOME_PRODUTO, TB_PRODUCAO.COD_FOTO_PADRAO "&_
			"FROM TB_PRODUTOS "&_
			"INNER JOIN TB_PRODUCAO ON TB_PRODUTOS.COD_PRODUTO = TB_PRODUCAO.COD_FOTO_PADRAO "&_
			"WHERE TB_PRODUCAO.COD_PRODUCAO = "&Request.QueryString("cod_prod")
				
	Set rsPadrao = objConexao.Execute(sql)
				
		If not rsPadrao.Eof then
			COD_FOTO_PADRAO = formatNumber(0&rsPadrao("COD_FOTO_PADRAO"),0)
			NOME_FOTO_PADRAO =  rsPadrao("NOME_PRODUTO")
		End If
				
	rsPadrao.Close()
	Set rsPadrao = Nothing
	
FecharConexao()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
	<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>
	<script language="javascript">
		function mOvr(src, cor)
			{
			if (!src.contains(event.fromElement))
					{
						src.style.cursor = 'hand';
						src.bgColor = cor;
					}
			}
		
		function mOut(src, cor)
			{
				if (!src.contains(event.toElement))
					{
						src.style.cursor = 'default';
						src.bgColor = cor;
			
					}
			}
	</script>
	
	<script type="text/javascript" language="javascript">
	<!--AJUDA F1 E CTRL PARA BUSCA RÁPIDA-->
		
	function abrir(theURL,winName,features) 
	{ 
	 filho = window.open(theURL,winName,features);
	}
	
	function sair(){
		
		if (filho){
		alert('ok')
		}
	
	}
		
	function abre(url)
	{    
				//if(event.srcElement.tagName=='TD')
				//	{
				//		src.children.tags('A')[0].click(); 
				//	}
	location.href = url;

	}

	function buscaIni(){
	d = document.buscaForm
	d.busca.focus()
	d.busca.select();
	}
	function mudacorBusca(CAMPO){
	CAMPO.style.backgroundColor = '#F5ECBA'
	}
	function mudacorBusca2(CAMPO){
	CAMPO.style.backgroundColor = '#FFFFFF'
	}

 function Verificar()
 {
  var tecla=window.event.keyCode;
  if (tecla==17) {
    buscaIni();
	}
 }
 
	<!--AJUDA F1 E CTRL PARA BUSCA RÁPIDA-->
	</script>
</head>

<body onKeyDown="javascript:Verificar()">

<table cellpadding="9px" cellspacing="0" width="100%">
<tr>
<td>
<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td>
		<!-- #include file="includes/topo.asp" -->
<table width="100%" style="border-bottom:#999999 solid dotted 1px" background="images/layout/fundo_menu_azul.jpg"border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
	<!-- #include file="includes/menu.asp" -->
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="10">
  <tr>
  <td>
   <%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 2
  BTN_FONTE = 1
  BTN_IMPRIMIR_LINK = "#"
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">

	  <h2>DIP - Produ&ccedil;&otilde;es / &Aacute;lbuns / Confirma&ccedil;&atilde;o </h2>

	       <%
		  AbrirConexao()
		  Dim rs, rs2, rs3
		  Dim QTD_ALBUNS, FICHAS_VENDA

		  Set rs = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO="&Request.QueryString("cod_prod")&" ORDER BY NOME_FORMANDO")
		  
		  If not rs.Eof then
		  %>
		  <h3 align="center">Confirma&ccedil;&atilde;o antes da inclus&atilde;o:</h3>
		  <div id="comboTexto">
		  <form action="dip_producoes_fichas_confirma_finaliza.asp" method="post">
          <table width="98%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="27%"><div align="center"><strong>Nome Formando </strong></div></td>
            <td width="8%"><div align="center"><strong>C&oacute;d. Form. </strong></div></td>
            <td width="6%"><div align="center"><strong>Contrato </strong></div></td>
            <td width="6%"><div align="center"><strong>Ano / Sem.  </strong></div></td>
            <td width="6%"><div align="center"><strong>Pasta</strong></div></td>
            <td width="8%"><div align="center"><strong>&Aacute;lbum</strong><strong></strong></div></td>
			<td width="6%"><div align="center"><strong>Álbum Digital </strong></div></td>
            <td width="8%"><div align="center"><strong>Inicial</strong></div></td>
            <td width="8%"><div align="center"><strong>Comiss&atilde;o</strong></div></td>
            <td width="8%"><div align="center"><strong>Brinde</strong></div></td>
            <td width="8%"><div align="center"><strong>Prod. Agreg. </strong></div></td>
            <td width="12%"><div align="center"><strong>Lote</strong></div></td>
</tr>
		  <%
		  Dim QTD_PREENCHIDA
		   QTD_PREENCHIDA = 0
		  While not rs.Eof
		  
		  ERROR_FLAG = 2
		  Dim INICIAL, COMISSAO, PRODUTO_AGREGADO
		  
		  '--- INICIAL ---
				If IsNull(rs("INICIAL")) Or Not IsNumeric(rs("INICIAL")) Then
					INICIAL = 0
				Else
					INICIAL = CLng(rs("INICIAL"))
				End If

				'--- COMISSAO ---
				If IsNull(rs("COMISSAO")) Or Not IsNumeric(rs("COMISSAO")) Then
					COMISSAO = 0
				Else
					COMISSAO = CLng(rs("COMISSAO"))
				End If

				'--- PRODUTO_AGREGADO ---
				If IsNull(rs("PRODUTO_AGREGADO")) Or Not IsNumeric(rs("PRODUTO_AGREGADO")) Then
					PRODUTO_AGREGADO = 0
				Else
					PRODUTO_AGREGADO = CLng(rs("PRODUTO_AGREGADO"))
				End If
		  
		  
		  Dim CONF_INICIAL, MODIFICACOES
		 
		  	'Verifica se h� valor digitado na produção
		  	'--- CONF_INICIAL (do form) ---
			If IsNull(Request.Form("ALBUM_"&rs("COD_FORMANDO"))) Or _
			Trim(Request.Form("ALBUM_"&rs("COD_FORMANDO"))) = "" Or _
			Not IsNumeric(Request.Form("ALBUM_"&rs("COD_FORMANDO"))) Then
				CONF_INICIAL = 0
			Else
				CONF_INICIAL = CLng(Request.Form("ALBUM_"&rs("COD_FORMANDO")))
			End If
			MODIFICACOES = 0&Request.Form("MODIFICACOES_"&rs("COD_FORMANDO"))			
			MODIFICACOES = formatNumber(MODIFICACOES,0)
						
			'////////////////////////////////////////////////////////////////////////////////////////////////
			If MODIFICACOES > 0 or CONF_INICIAL > 0  then
			
			
			'SOMENTE FICHA ALTERADA/EXCLU�DA OU FICHAS MODIFICADAS
			
						'MOMENTO CRUCIAL, NESTA ETAPA OS DADOS DA P�GINA ANTERIOR seráO INCLU�DAS NUMA TABELA TEMPOR�RIA
			'1) Guarda os valores em vari�veis da pag. anterior
			Dim ALBUM_FORM, ALBUM_DIGITAL_FORM, INICIAL_FORM, COMISSAO_FORM, BRINDE_FORM, PROD_AGREG_FORM, VENDEDOR_FORM, ENCADERNACOES_EXTRAS_FORM, LOTE_FORM
			
			ALBUM_FORM = 				ExtrairNumeros(Request.Form("ALBUM_"&rs("COD_FORMANDO")))
			ALBUM_DIGITAL_FORM = 		ExtrairNumeros(Request.Form("ALBUM_DIGITAL_"&rs("COD_FORMANDO")))
			PASTA_FORM=					ExtrairNumeros(Request.Form("PASTA_"&rs("COD_FORMANDO")))
			INICIAL_FORM =  			formatNumber(0&Request.Form("INICIAL_"&rs("COD_FORMANDO")),0)/1
			COMISSAO_FORM =  			Request.Form("COMISSAO_"&rs("COD_FORMANDO"))
			BRINDE_FORM =  				Request.Form("BRINDE_"&rs("COD_FORMANDO"))
			'VENDEDOR_FORM = 			Request.Form("COD_VENDEDOR_"&rs("COD_FORMANDO")) 'Não é mais necessário
			PROD_AGREG_FORM = 			Request.Form("PRODUTO_AGREGADO_"&rs("COD_FORMANDO"))
			ENCADERNACOES_EXTRAS_FORM =	ExtrairNumeros(formatNumber(0&Request.Form("ENCADERNACOES_EXTRAS_"&rs("COD_FORMANDO")),0))
			LOTE_FORM =					formatNumber(0&Request.Form("LOTE_"&rs("COD_FORMANDO")),0)
			
				'S� grava na tempor�ria se Não tiver gravado definitivo na base de dados o álbum
					If rs("ALBUM") = 0 And rs("INICIAL") = 0 then
							'2) Limpa o registro tempor�rio da tabela tempor�ria somente do formando selecionado
			
			 objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_TEMPORARIA WHERE "&_
								"COD_FORMANDO = "&rs("COD_FORMANDO")&" LIMIT 1")
			
			'3) Grava na tempor�ria 
			'Caso for definitivamente gravada esta linha abaixo será apagada.

			IF Not IsNumeric(ALBUM_DIGITAL_FORM) then ALBUM_DIGITAL_FORM = 0
				
			sql = "INSERT INTO TB_FICHAS_IDENTIFICACAO_TEMPORARIA ("&_
                "COD_FORMANDO ,"&_
                "ALBUM ,"&_
                "ALBUM_DIGITAL ,"&_
                "numero_pasta_gerencia,"&_
                "ENCADERNACOES_EXTRAS ,"&_
                "INICIAL ,"&_
                "COMISSAO ,"&_
                "BRINDE ,"&_
                "LOTE, PROD_AGREGADO) "&_
                "VALUES ("&_
                rs("COD_FORMANDO") & ", " & _
                CLng(ALBUM_FORM) & ", " & _
                CLng(ALBUM_DIGITAL_FORM) & ", " & _
                CLng(PASTA_FORM) & ", " & _
                CLng(ENCADERNACOES_EXTRAS_FORM) & ", " & _
                CLng(INICIAL_FORM) & ", " & _
                CLng(COMISSAO_FORM) & ", " & _
                CLng(BRINDE_FORM) & ", " & _
                sqltexto(LOTE_FORM) & ", " & _
                CLng(PROD_AGREG_FORM) & ")"

								
							''	Response.Write(sql)

			 objConexao.Execute(sql)
								
			End If					
			
			End If	

			
			'5)Verifica se h� exclus�o ou alteração
			If MODIFICACOES = 1 then
			'Alterar, resetar todos os registros e voltar a etapa inicial
				VERIFICA_MODIFICACOES = MODIFICACOES
				COD_FORMANDO_ALTERADO = rs("COD_FORMANDO")
				NOME_FORMANDO_ALTERADO= rs("NOME_FORMANDO")
				STATUS_MODIFICACAO = "resetado"
				
				'1) Resetar Ficha de Identifica��o
				objConexao.Execute("UPDATE TB_FICHAS_IDENTIFICACAO SET "&_
									"ALBUM = 	0, "&_
									"ALBUM_DIGITAL = 0, "&_
									"INICIAL = 	0, "&_
									"ENCADERNACOES_EXTRAS = 0, "&_
									"BRINDE = 	2, "&_
									"COMISSAO = 2, "&_
									"PRODUTO_AGREGADO = 2, "&_
									"cod_producao_desconto = '' "&_
									"WHERE COD_FORMANDO = "&rs("COD_FORMANDO")&" LIMIT 1")
			
				'2) Exclui brindes se houver.
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_BRINDES WHERE COD_FORMANDO="&rs("COD_FORMANDO"))
				
				
				'3) Exclui movimenta��o se houver.
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_FORMANDO="&rs("COD_FORMANDO"))
								
				'5) Exclui produtos agregados se houver.
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS WHERE COD_FORMANDO="&rs("COD_FORMANDO"))	
				
				'6) Exclui tabela tempor�ria se houver.
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_TEMPORARIA WHERE COD_FORMANDO="&rs("COD_FORMANDO"))				 
					
					
			Elseif MODIFICACOES = 2 then
			'Exclui
				VERIFICA_MODIFICACOES = MODIFICACOES
				COD_FORMANDO_ALTERADO  = rs("COD_FORMANDO")
				NOME_FORMANDO_ALTERADO = rs("NOME_FORMANDO")
				STATUS_MODIFICACAO = "excluida"
				
				'1) Exclui Ficha de Identifica��o
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO="&rs("COD_FORMANDO")&" LIMIT 1")
			
				'2) Exclui brindes se houver.
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_BRINDES WHERE COD_FORMANDO="&rs("COD_FORMANDO"))
						
				'3) Exclui movimenta��o se houver.
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_FORMANDO="&rs("COD_FORMANDO"))
				
				'4) Exclui novo endere�o se houver;
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO  WHERE COD_FORMANDO="&rs("COD_FORMANDO"))
				
				'5) Exclui produtos agregados se houver;
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS WHERE COD_FORMANDO="&rs("COD_FORMANDO"))	
				
				'6) Exclui tabela tempor�ria se houver;
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_TEMPORARIA WHERE COD_FORMANDO="&rs("COD_FORMANDO"))	
				
				
			
			End If
			
			If CONF_INICIAL > 0 then
		  	QTD_PREENCHIDA = QTD_PREENCHIDA + 1
			
		  %>
         	
			<tr bgcolor="#FFFF99">
				<!--NOME FORMANDO-->
            <td><%=rs("NOME_FORMANDO")%></td>
				<!--CÓDIGO FORMANDO-->
            <td><%=digito(rs("COD_FORMANDO"))%></td>
				<!--CÓDIGO CONTRATO-->
            <td><%=rs("COD_CONTRATO")%>
			</td>
				<!--ANO CONCLUSÃO / SEMESTRE-->
            <td><%=rs("ANO_CONCLUSAO")%> / <%=rs("SEMESTRE_CONCLUSAO")%></td>
            
			<td> <!--PASTA-->
				<span class="textoMaior"><%=PASTA_FORM%></strong></span>
 				<input style="width:40px" type="hidden" name="PASTA_<%=rs("COD_FORMANDO")%>" value="<%=PASTA_FORM%>" />
			</td>
				<!--ALBUM / DUPLO-->
            <td><div align="center">
			<span class="textoMaior"><strong>
				<%= ExtrairNumeros(Request.Form("ALBUM_"&rs("COD_FORMANDO")))%>
				<%
				'If int(ENCADERNACOES_EXTRAS_FORM) = 1 Then
					'Response.Write(" - SIM")
				
				'End If
				%>
				
			</strong>			
			</span>
			
              <input style="width:40px" type="hidden" name="ALBUM_<%=rs("COD_FORMANDO")%>" value="<%= ALBUM_FORM %>" />
			  <input style="width:40px" type="hidden" name="ENCADERNACOES_EXTRAS_<%=rs("COD_FORMANDO")%>" value="<%= ENCADERNACOES_EXTRAS_FORM %>" />
            </div></td>

			<!--ALBUM DIGITAL-->
			<td><div align="center">
				<%If int(ALBUM_DIGITAL_FORM) = 1 then%>
					<span class="textoMaior"><strong>Sim</strong></span>
				<%Else%>
					<span class="textoMaior"><strong>Não</strong></span>
				<%End If%>

				<input name="ALBUM_DIGITAL_<%=rs("COD_FORMANDO")%>" type="hidden" value="<%=ALBUM_DIGITAL_FORM%>" /> 
			</div></td>


				<!--INICIAL-->
            <td><div align="center">
			
			<%If INICIAL_FORM <> 0 then%>
				<span class="textoMaior"><strong><%=INICIAL_FORM%></strong></span>
				<input style="width:40px" type="hidden" name="INICIAL_<%=rs("COD_FORMANDO")%>" value="<%=formatNumber(Request.Form("INICIAL_"&rs("COD_FORMANDO")),0)%>" />			
			<%else
			
				ERROR_FLAG = 1
			%>
				<span class="msg">Quantidade inicial Não cadastrada, esta ficha Não será cadastrada.</span>
			<%End If%>
			
              
            </div></td>
				<!--COMISSÃO-->
            <td><div align="center">

				<%If COMISSAO_FORM = 1 then%>
					
					<%
					'Verifica se foi selecionado o brinde da comissão
					If 	BRINDE_FORM = 2 then
						ERROR_FLAG = 1	
					%>
						
						<span class="msg">Você selecionou esta ficha como membro de comissão, entretanto Não informou se o mesmo tem brinde, por tanto Não será gravada está ficha.</span>
					    <%else%>
						
						<span class="textoMaior"><strong>Sim</strong></span>
						
						<%End If 
					Else%>
					<span class="textoMaior"><strong>Não</strong></span>
				<%End If%>
				<input name="COMISSAO_<%=rs("COD_FORMANDO")%>" type="hidden" value="<%=Request.Form("COMISSAO_"&rs("COD_FORMANDO"))%>" /> 
			  
            </div></td>
				<!--BRINDE-->
            <td><div align="center">
    
				<%If BRINDE_FORM = 1 then
					'Verifica se foi informado o brinde
					Dim rsBrinde
					Set rsBrinde = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO_BRINDES WHERE COD_FORMANDO="&rs("COD_FORMANDO"))
					
					If rsBrinde.Eof then
						
						set rsDesconto = objConexao.Execute("SELECT cod_producao_desconto FROM TB_FICHAS_IDENTIFICACAO WHERE cod_producao_desconto is not NULL AND COD_FORMANDO = "&sqlNumero(rs("COD_FORMANDO")))
						If rsDesconto.Eof Then
							ERROR_FLAG = 1	

							'Não h� brinde cadastrado%>
							<span class="msg">Você selecionou Brinde como "Sim", entretanto Não relacionou os brindes, 
						por tanto está ficha Não será cadastrada.</span>

						<% 
						end if
						rsDesconto.Close
						set rsDesconto = Nothing

					Else
		
					
						'Valida se a quantidade inicial � menor que a quantidade de brinde
						While not rsBrinde.Eof 
						
							Dim QTD_BRINDE_FOTO
							COD_PRODUTO_BRINDE_FOTO = formatNumber(0&rsBrinde("COD_PRODUTO"),0)
							QTD_BRINDE_FOTO = formatNumber(0&rsBrinde("QUANTIDADE"),0)/1
							
							

							If COD_PRODUTO_BRINDE_FOTO = COD_FOTO_PADRAO Then
							'Agora sim verifica se � menor a qtd inicial
						
									If INICIAL_FORM < QTD_BRINDE_FOTO Then
								
								'Bloqueia Inicial menor que brinde
									ERROR_FLAG = 1%>
									<span class="msg">Quantidade inicial do álbum (<strong><%=INICIAL_FORM%></strong>) é menor que a quantidade de brinde (<strong><%=QTD_BRINDE_FOTO%></strong>) , está ficha Não será gravada.</span>
								<%
								End If
								
							End If	
							
							
							
						rsBrinde.MoveNext()
						Wend
					End If
					
					
					rsBrinde.Close()
					Set rsBrinde = Nothing
					
					If ERROR_FLAG = 2 Then%>
						<span class="textoMaior"><strong>Sim</strong></span>
					<%End If%>
				<%Else%>
				<span class="textoMaior"><strong>Não</strong></span>
				<%End If%>
	<input name="BRINDE_<%=rs("COD_FORMANDO")%>" type="hidden" value="<%=BRINDE_FORM%>" /> 
	
            </div></td>
				<!--PRODUTO AGREGADO-->
            <td><div align="center">		
			<%If int(PROD_AGREG_FORM) = 1 then%>
					<span class="textoMaior"><strong>Sim</strong></span>
			<%Else%>
					<span class="textoMaior"><strong>Não</strong></span>
			<%End If%>
				   

			<input name="PRODUTO_AGREGADO_<%=rs("COD_FORMANDO")%>" type="hidden" value="<%=Request.Form("PRODUTO_AGREGADO_"&rs("COD_FORMANDO"))%>" /> 


			
            </div></td>
				<!--LOTE-->
            <td><div align="center">
		
						<%If LOTE_FORM <> 0 then%>
						  	<span class="textoMedio"><strong><%=LOTE_FORM%></strong></span>
						  <%						
						elseif MODIFICACOES  = 9 Then%>
							<strong>Devido a alteração for realizada após a finalização da produção a alteração de lote Não poderá ser feita.</strong><input style="width:35px" name="MODIFICACOES_<%=rs("COD_FORMANDO")%>" value="9" type="hidden" />
						    <%else
							ERROR_FLAG = 1
							%>
							<span class="msg">Lote Não informado, esta ficha Não será cadastrada!</span>
							<%End If%>
						
						<input name="LOTE_<%=rs("COD_FORMANDO")%>" type="hidden" value="<%=LOTE_FORM%>" /> 
						<input name="ERROR_FLAG_<%=rs("COD_FORMANDO")%>" type="hidden" value="<%=ERROR_FLAG%>" /> 
						 
            </div></td>
          </tr>
          
		  <%
		  	End If
		  
		  'Contador
		  If ERROR_FLAG = 1 then
		  
		  	QTD_ERROR = QTD_ERROR + 1	  
		  End If 
		  
		  rs.MoveNext()
		  Wend
	      
		  rs.Close()
		  Set rs = Nothing
		  
		  '//////////////////////////////////////////////////////////
		  'Redireciona se for da modifica��o
		  If VERIFICA_MODIFICACOES > 0 then
		  
		  	Response.Redirect("dip_producoes_fichas.asp?cod_prod="&Request.QueryString("cod_prod")&"&prod_agre="&Request.QueryString("prod_agre")&"&cod_form="&COD_FORMANDO_ALTERADO&_
			"&msg=Ficha do(a) formando(a): "&NOME_FORMANDO_ALTERADO&" foi "&STATUS_MODIFICACAO&" com sucesso!&s="&VERIFICA_MODIFICACOES)  
		  End If
		  
		  'Precisa fazer o script de focus na ficha alterada 
		  '//////////////////////////////////////////////////////////
		  
		  
		  
		  
		  %>
		   <tr>
            <td colspan="11"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td>Fichas Preenchidas:<strong>&nbsp; <%=QTD_PREENCHIDA%>&nbsp; |&nbsp; </strong>Fichas com erros: <strong class="msg"><%=QTD_ERROR%></strong>&nbsp; |&nbsp;&nbsp; Fichas Corretas: <strong class="msg2"><%=QTD_PREENCHIDA-QTD_ERROR%></strong> </td>
                </tr>
              <tr>
                <td>Lotes já definidos:<br />
                <%
				'Set rs=objConexao.Execute("SELECT LOTE FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO="&Request.QueryString("cod_prod")&" group by lote order by lote")
				
				
				'rs.Close
				'Set rs=Nothing
				%>                </td>
              </tr>
              <tr>
                <td>Lotes tempor&aacute;rios:</td>
              </tr>
            </table></td>
            </tr>
		   <tr>
		     <td colspan="12"><div align="center">
               <input class="btnVoltar" type="button" name="Submit22" onclick="history.go(-1)" value=" " />
  &nbsp;
  <input type="hidden" name="cod_prod" value="<%=Request.QueryString("cod_prod")%>"/>
  <input class="btnGravar" type="submit" name="Submit2" value=" " />
             </div></td>
		     </tr>
        </table>
		  </form>
		</div>
		  <div align="center">
		    <p>
		      <%else%>
		      <em>N&atilde;o h&aacute; formando importados</em>!</p>
		    <p><a href="#">Deseja importar agora?</a></p>
		    <p>&nbsp;</p>
		    <p>	      
		      <%end if%>
		      </p>
		  </div>
		  Observações:<br />
		  Digito 0 = Formando importado da formatura<br />
		  Digito 1 = Formando cadastrado pelo depto. de foto
		  <p align="center">
		  <input class="btnTopo" onclick="abre('#')" type="button" name="btnTopo2" value=" " />
		</p>		</td>
    </tr>
</table>

 
  
  </td>
    </tr>
</table>

  
  </td>
  </tr>
</table>
<!-- #include file="includes/rodape.asp" -->

  </td>
</tr>
</table>
</body>
</html>