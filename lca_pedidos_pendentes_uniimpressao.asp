<%
Dim COD_PEDIDO_QUERY_STRING
COD_PEDIDO_QUERY_STRING = Request.QueryString("cod_pedido")
'CORRIGIR PROBLEMA QUANDO POSSUI UM PRODUTO COM SAÍDA DE PRODUÇÃO E O MESMO ESTÁ LANÇADO COMO BRINDE
%>
<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/impressao_pedido_PDF.asp" -->
<%
AbrirConexao()
	
		'Abri o PDF	
		'////////////////////////
		Set PDF = Server.CreateObject("Persits.PDF")

		' Open an existing document, form W-9
		Set Doc = PDF.OpenDocument("e:\home\stillos\web\gerencial_foto\documentos\impressao_pedidos_individual_v3.pdf")
		
		
		' Obtain Page object for page 1
		Set Page = Doc.Pages(1)


		' Create font object
		Set Font = Doc.Fonts("Arial")
		Set Param = PDF.CreateParam("html=true")
							
		'//////////////////////////////////////////////////////////////////////////////////////////////////////
		'Atenção padroniza somente o miolo, ou seja, se houver alguma alteração deverá ser feito no arquivo impressao_pedido_PDF.asp nos includes
		call impressaoPedidoPDF(COD_PEDIDO_QUERY_STRING,"A")									

		'////////////////////////
		Doc.SaveHttp "attachment;filename=pedido_individual_impressao_"&COD_PEDIDO_QUERY_STRING&".pdf"
		'FileName = Doc.SaveHttp("e:\home\stillos\web\gerencial_foto\documentos\get_impressao_pedido_individual.pdf")
		Set Page = Nothing
		Set Doc = Nothing
		Set Pdf = Nothing
	
FecharConexao()
%>