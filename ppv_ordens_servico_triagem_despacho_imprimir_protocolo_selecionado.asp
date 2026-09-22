<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
    <%
	AbrirConexao()
	
	Dim cod_pedido
	Dim cod_pedido_inserir
	Dim resto
	Dim sqlWhere
	
	'Marcando produtos como entregue
	produtos = Cstr(""&Request.Form("produtos"))		

	If produtos <> "" Then
	'Inserir individualmente o pedido
	
		sqlWhere = " AND cod_ordem_servico IN ("&produtos&")"
		Set rs=objConexao.Execute("SELECT cod_ordem_servico, cod_pedido FROM TB_ORDENS_SERVICO WHERE cod_ordem_servico IN ("&produtos&")")
			If not rs.Eof Then
				While not rs.Eof
						cod_pedido_inserir = rs("cod_pedido")
					objConexao.Execute("INSERT INTO TB_ORDENS_SERVICO_TEMPORARIA (cod_ordem_servico,cod_usuario,cod_pedido) VALUES ('"&rs("cod_ordem_servico")&"', '"&Session("cod_usuario")&"','"&cod_pedido_inserir&"')")
				rs.MoveNext()
				Wend
			End If
	
		rs.Close
		Set rs=Nothing	
	End If
	
	
	Dim i, ii
	i	= 0
	ii	= 0 
		 
	sql = "SELECT TB_ORDENS_SERVICO.cod_pedido, TB_ORDENS_SERVICO.data_base_prazo_cliente_inicio FROM TB_ORDENS_SERVICO INNER JOIN TB_ORDENS_SERVICO_TEMPORARIA ON (TB_ORDENS_SERVICO.cod_ordem_servico = TB_ORDENS_SERVICO_TEMPORARIA .cod_ordem_servico) WHERE TB_ORDENS_SERVICO_TEMPORARIA.cod_usuario = "&Session("cod_usuario")&" AND TB_ORDENS_SERVICO.cod_ordem_servico IN ("&produtos&") GROUP BY TB_ORDENS_SERVICO.cod_pedido order by cod_pedido asc"
	
	
	Set rs=objConexao.Execute(sql,qtd_pedidos)
	
	
	If not rs.Eof Then
	
			
	'1º Parte Cria os arquivos vazios e da nome as futuros arquivos
	'//////////////////////////////////////////////////////////////////////////////////////////////////
	
		Dim COD_ALEATORIO
		COD_ALEATORIO			= Now()
		COD_ALEATORIO			= Replace(COD_ALEATORIO,":","")
		COD_ALEATORIO			= Replace(COD_ALEATORIO,",","")
		COD_ALEATORIO			= Replace(COD_ALEATORIO,"/","")
		COD_ALEATORIO			= Replace(COD_ALEATORIO," ","")
		NOME_ARQUIVO_MASTER 	= "master_pe_"&COD_ALEATORIO&"_"&Session("COD_USUARIO")
		NOME_ARQUIVO_INDIVIDUAL = "individual_pe_"&COD_ALEATORIO&"_"&Session("COD_USUARIO")
		NOME_ARQUIVO_HIPER		= "hiper_pe_"&COD_ALEATORIO&"_"&Session("COD_USUARIO")

		SET AspPdf = Server.CreateObject("Persits.Pdf")
		SET Doc = AspPdf.CreateDocument
		SET Page = Doc.Pages.Add(216, 280)
		'Define os parâmetros de alinhamento: X, Y do canto superior esquerdo ao lado inferior direito, altura, largura e tamanho da fonte.
		Filename = Doc.Save( "e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf") 
		SET AspPdf = Nothing
		SET Doc = Nothing
		SET Page = Nothing
		SET Font = Nothing
		SET Param = Nothing
		
		SET AspPdf = Server.CreateObject("Persits.Pdf")
		SET Doc = AspPdf.CreateDocument
		SET Page = Doc.Pages.Add( 216, 280 )
		'Define os parâmetros de alinhamento: X, Y do canto superior esquerdo ao lado inferior direito, altura, largura e tamanho da fonte.
		Filename = Doc.Save( "e:\home\stillos\web\gerencial_foto\documentos\hiper_pe.pdf")
		SET AspPdf = Nothing
		SET Doc = Nothing
		SET Page = Nothing
		SET Font = Nothing
		SET Param = Nothing
		
		
		'//////////////////////////////////////////////////////////////////////////////////////////////////	
		i = 0
		
		While not rs.Eof
 
 		i=i+1
 
		'Abri o PDF	
			'////////////////////////
			Set PDF = Server.CreateObject("Persits.PDF")
	
			' Open an existing document, form W-9
			Set Doc = PDF.OpenDocument("e:\home\stillos\web\gerencial_foto\documentos\protocolo_entrega.pdf")
	
			' Obtain Page object for page 1
			Set Page = Doc.Pages(1)
	
	
			' Create font object
			Set Font = Doc.Fonts("Arial")
			Set Param = PDF.CreateParam("html=true")
			
			'////////////////////////
			call impressaoProtocoloEntregaPDF(rs("COD_PEDIDO"),"A")
			call impressaoProtocoloEntregaPDF(rs("COD_PEDIDO"),"B")
			
				
			'Fecha o arquivo PDF duplo	
			'////////////////////////
			FileName = Doc.Save("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_INDIVIDUAL&".pdf",true)
			Set Page = Nothing
			Set Doc = Nothing
			Set Pdf = Nothing
			
			Dim DocMaster, DocIndividual
			Set PDF = Server.CreateObject("Persits.PDF")
					
			'Abri o master e o acabado de ser criado
			Set DocMaster 		= PDF.OpenDocument("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf")
			Set DocIndividual 	= PDF.OpenDocument("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_INDIVIDUAL&".pdf")
			Set Page = DocIndividual.Pages(1)
			'Concatena os 2 arquivos
			DocMaster.AppendDocument DocIndividual
				
			'Salva os arquivos
			FileName = DocMaster.Save("e:\home\stillos\web\gerencial_foto\documentos\hiper_pe.pdf",true)
			
			'Fechar os Arquivos
			Set DocMaster		= Nothing
			Set DocIndividual 	= Nothing
			Set PDF = Nothing
					
			'Deleta o Master
			deletarArquivo("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf")
			
				
			'Renomeia o Hiper para Master
			set fso = CreateObject("Scripting.FileSystemObject") 
			fso.MoveFile "e:\home\stillos\web\gerencial_foto\documentos\hiper_pe.pdf", "e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf"
			set fso = Nothing
					
	
		rs.MoveNext()
		Wend 
		
		'#######################################################################################
		
		'Retira a primeira página e exibe na tela
		Set Pdf = Server.CreateObject("Persits.Pdf") 
		Set Doc = Pdf.OpenDocument("e:\home\stillos\web\gerencial_foto\documentos\"&NOME_ARQUIVO_MASTER&".pdf")
			
		
		Doc.Pages.Remove 1
		Doc.SaveHttp "attachment;filename=protocolo_entrega.pdf"
					
		Set Doc = Nothing
		Set Pdf = Nothing
		'////////////////////////////
			
	
		End If
			
	
	rs.Close
	Set rs = Nothing

	If cod_pedido_inserir <> "" Then
	'Inserir individualmente o pedido
		Set rs=objConexao.Execute("SELECT cod_ordem_servico FROM TB_ORDENS_SERVICO WHERE COD_PEDIDO = "&cod_pedido_inserir&" AND data_despachado IS NOT NULL")
			If not rs.Eof Then
				While not rs.Eof
					objConexao.Execute("DELETE FROM TB_ORDENS_SERVICO_TEMPORARIA WHERE cod_ordem_servico = "&rs("cod_ordem_servico")&" AND cod_usuario = "&Session("cod_usuario"))
				rs.MoveNext()
				Wend
			End If
	
		rs.Close
		Set rs=Nothing	
	End If
	
FecharConexao()	

Sub impressaoProtocoloEntregaPDF(cod_pedido,IMPRESSAO_A_OU_B)


	'//////////////////////////////////////////////////////////////////////////////////////////////////
	'Variáveis
	'/////
	
	
	If IMPRESSAO_A_OU_B = "B" Then
		diferenca_y = 400
	Else
		diferenca = 0
	End If
	
	Dim sql, rs, rs2, rs3
	Dim i, ii, iii, temp, temp2, temp3
	
	Param.Add("x=390, y="&(814-diferenca_y)/1&", size=20")
	Page.Canvas.DrawText "<b>"&cod_pedido&"/"&year(now())&"</b>", Param, Font
		
	sql = 	" SELECT "&_
	
			" (SELECT concat(ADDRESS.ENDERECO,' , ', ADDRESS.CEP,' ,', ADDRESS.CIDADE,'-', ADDRESS.ESTADO) "&_	
			" FROM TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO ADDRESS WHERE TB_PEDIDOS_FORMANDOS.COD_FORMANDO = ADDRESS.COD_FORMANDO LIMIT 0,1) AS Endereco, "&_
			
			" TB_PEDIDOS_FORMANDOS.COD_FORMANDO, "&_
			" TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, "&_
			" TB_FICHAS_IDENTIFICACAO.COD_CONTRATO, "&_
			" TB_ENTIDADES.DESCRICAO, "&_
			" TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO, "&_
			" TB_FICHAS_IDENTIFICACAO.NOME_CURSO "&_
			" FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON "&_
			" TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO "&_
			" INNER JOIN TB_ENTIDADES ON(TB_FICHAS_IDENTIFICACAO.COD_CONTRATO = TB_ENTIDADES.COD_CONTRATO)"&_
			" WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&cod_pedido&" ORDER BY NOME_FORMANDO"
	response.write(sql)										
	Set rs2=objConexao.Execute(sql)
		
		i = 0
		While not rs2.Eof
			
			temp = (760-diferenca_y)/1
			temp = (temp-(i*12))/1
			
			Param.Add("x=30, y="&temp&", size=9")
			Page.Canvas.DrawText rs2("NOME_FORMANDO"), Param, Font
			
		
			Param.Add("x=220, y="&temp&", size=9")
			Page.Canvas.DrawText rs2("DESCRICAO"), Param, Font			
					
		
			Param.Add("x=430, y="&temp&", size=9")
			Page.Canvas.DrawText rs2("NOME_CURSO"), Param, Font		
			
			
			' Adicionando endereço
			temp = cint(cint(temp) - 15)
			Param.Add("x=30, y="&temp&", size=9")
			Page.Canvas.DrawText "Endereço: ", Param, Font
			
			temp = cint(cint(temp) - 12)
			Param.Add("x=30, y="&temp&", size=9")
			Page.Canvas.DrawText rs2("Endereco"), Param, Font		
			
			
			i=i+1
			
		rs2.MoveNext()
		Wend

		rs2.Close
		Set rs2=Nothing				
					
		
		
		sql = 	"SELECT VW_TB_ORDENS_SERVICO.* FROM VW_TB_ORDENS_SERVICO INNER JOIN TB_ORDENS_SERVICO_TEMPORARIA ON (VW_TB_ORDENS_SERVICO.cod_ordem_servico = TB_ORDENS_SERVICO_TEMPORARIA.cod_ordem_servico) WHERE TB_ORDENS_SERVICO_TEMPORARIA.cod_usuario = "&Session("cod_usuario")&" and VW_TB_ORDENS_SERVICO.cod_pedido="&cod_pedido&" order by VW_TB_ORDENS_SERVICO.numero_ordem_servico_indice asc"
		

	Set rs2=objConexao.Execute(sql)
		
		i = 0
		While not rs2.Eof


			temp = (655-diferenca_y)/1
			temp = (temp-(i*9))/1

			Param.Add("x=40, y="&temp&", size=9")
			Page.Canvas.DrawText moduloPPVOSMascara(rs2("cod_ordem_servico"),"",true), Param, Font		

			Param.Add("x=170, y="&temp&", size=9")
			Page.Canvas.DrawText rs2("quantidade"), Param, Font	
	
			Param.Add("x=225, y="&temp&", size=9")
			Page.Canvas.DrawText Ucase(rs2("nome_produto")), Param, Font			
			
		
		
		i=i+1
		rs2.MoveNext()
		Wend		
		
			
	rs2.Close
	Set rs2=Nothing				
						
End Sub

%>



