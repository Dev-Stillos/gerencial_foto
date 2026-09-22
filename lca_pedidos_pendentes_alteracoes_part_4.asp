<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%Dim msg_erro_readequacao_vendedor_recebe_em%>
<!-- #include file="includes/classe_fechamento.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<% 
'######################################################################
'INSTRU��ES CORRE��O DE PEDIDOS
'1) Lan�amento da DATA da VENDa // Muito importante, pois � calculado dados como quando o vendedores receber�
'2) Corre��o de Produtos Lan�ados e Calculo se o �lbum � repasse ou n�o, caso for ser� concedito um acr�scimo na comiss�o totalizando 10% do valor do recibo, independente se o pedido n�o conter somente o �lbum
'3) Corre��o de Forma de Pagamento
'4) Inser��o de Ajuda de Custo, Visualiza��o de Quando o vendedor receber�, Observa��o
'5) Conclus�o confirma��o e finaliza��o
'######################################################################

Dim COD_PEDIDO, VALIDADOR, QTD_PARCELAS, PEDIDO_EXTRA, motivo_solicitacao, NECESSARIO_AUTORIZACAO_ESPECIAL, PODE_FINALIZAR
Dim POSSUI_ERRO, cod_fechamento_periodo
COD_PEDIDO = Request.QueryString("cod_pedido")
VALIDADOR = Request.Form("VALIDADOR")

If VALIDADOR <> 1 Then
	Response.Redirect("lca_pedidos_pendentes_alteracoes_part_3.asp?cod_pedido="&COD_PEDIDO&"&msg=Verifique se h&aacute; alguma forma de pagamento com o status em vermelho")
End If

If Request.Form("btnFinalizaCorrecao") <> Empty Then

'1) Finaliza venda
'2) Finaliza na ficha como vendido
'3) Informa nas movimenta��es que est� vendido
'4) // Ainda n�o dispon�vel // Da um insert na TB venda de DVD para controle de DVD
	AbrirConexao()


	motivo_solicitacao	= Request.Form("motivo_solicitacao")
	NECESSARIO_AUTORIZACAO_ESPECIAL	= Request.Form("NECESSARIO_AUTORIZACAO_ESPECIAL")
	PODE_FINALIZAR = Request.Form("PODE_FINALIZAR")
	
	'////////////////////////////////////////////////
	If PODE_FINALIZAR = "N" Then
		
		msg_status = "Pedido n&atilde;o finalizado, aguardando aprova&ccedil;&atilde;o"	
		
		'Verifica se � necess�rio gravar autoriza��o especial
		If NECESSARIO_AUTORIZACAO_ESPECIAL = "S" Then
			'respons�vel �lcio
			objConexao.Execute("INSERT INTO TB_APROVACOES (cod_categoria, solicitante_cod_usuario, responsavel_cod_usuario, motivo_solicitacao, parametro) VALUES (1,"&Session("cod_usuario")&",9,"&sqlTexto(motivo_solicitacao)&","&sqlTexto(cod_pedido)&")")
		
		Else
			'Atualiza o motivo atual
			objConexao.Execute("UPDATE TB_APROVACOES SET solicitante_cod_usuario = '"&Session("cod_usuario")&"', motivo_solicitacao = "&sqlTexto(motivo_solicitacao)&" WHERE cod_categoria = 1 and parametro = "&sqlTexto(cod_pedido))
			
		End If
		
	Else
	
	Set rs=objConexao.Execute("SELECT PEDIDO_EXTRA FROM TB_PEDIDOS WHERE COD_PEDIDO = "&COD_PEDIDO)	
			
		If not rs.Eof Then
			PEDIDO_EXTRA = rs("PEDIDO_EXTRA")
		Else
			Response.Redirect("lca_pedidos_pendentes.asp?msg=Falha tente novamente!")
		End If
			
	rs.Close
	Set rs=Nothing
			
	If int(PEDIDO_EXTRA) = 2 Then
	
	'//////////////////////////////////////////////////
		'No exato momento da grava��o, verifica se o formando est� vendido, se estiver redireciona para a primeira p�g. informando que o formando consta como vendido no pedido tal caso o mesmo n�o for pedido extra.
		'Dim rsFormandoVendido
		
		'PERMITE CORRE��O EM CASOS QUE O PEDIDO NAO FOI BAIXADO PARA 
		Set rs = objConexao.Execute("SELECT COD_PEDIDO FROM TB_PEDIDOS WHERE COD_PEDIDO = '"&COD_PEDIDO&"' and  DATA_DESCEU_DIRETORIA IS NOT NULL")
	
			If not rs.Eof Then	
				Response.Redirect("lca_pedidos_pendentes.asp?cod_pedido="&COD_PEDIDO&"&msg=Pedido: "&rs("COD_PEDIDO")&" consta como j&aacute; desceu para diretoria e n&atilde;o poder&aacute; ser alterado!")
			End If

		rs.Close()
		Set rs = Nothing
		'//////////////////////////////////////////////////
		
		
		'Se chegou at� aqui � porque pode atualizar o pedido
		Dim rsFormandos
		
		Set rsFormandos = objConexao.Execute("SELECT * FROM TB_PEDIDOS_FORMANDOS WHERE COD_PEDIDO="&COD_PEDIDO)
			
			While not rsFormandos.Eof
			
				objConexao.Execute( " UPDATE TB_FICHAS_IDENTIFICACAO SET VENDIDO = '1' "&_
									" WHERE COD_FORMANDO= "&rsFormandos("COD_FORMANDO")&" LIMIT 1")
									
				objConexao.Execute( " UPDATE TB_PRODUCAO_LOTES_FORMANDOS SET STATUS	= '1' "&_
									" WHERE COD_FORMANDO= "&rsFormandos("COD_FORMANDO")&" LIMIT 1")				
									
				'### Limpa o motivo tempor�rio do formando
				objConexao.Execute("DELETE FROM TB_PRODUCAO_LOTES_FORMANDOS_MOTIVO_TEMPORARIA WHERE COD_FORMANDO = "&rsFormandos("COD_FORMANDO"))
			
			rsFormandos.MoveNext()
			Wend
		
		rsFormandos.Close
		Set rsFormandos = Nothing
		End If
		
		Dim VENDEDOR_RECEBE_EM
		VENDEDOR_RECEBE_EM = Request.Form("VENDEDOR_RECEBE_EM")
		cod_fechamento_periodo = Request.Form("cod_fechamento_periodo") 'cod_fechamento_periodo autom�tica vinda das fun��es
		
		'////////////////////////////////////////////////
		'Altera��o manual de data de recebimento
		
		Dim cod_fechamento_periodo_manual, cod_fechamento_periodo_manual_motivo
		Dim data_pagamento_manual
		
		cod_fechamento_periodo_manual = Request.Form("cod_fechamento_periodo_manual")
		cod_fechamento_periodo_manual_motivo = Request.Form("cod_fechamento_periodo_manual_motivo")
		
		
		If cod_fechamento_periodo_manual <> "" and cod_fechamento_periodo_manual_motivo <> "" Then
			
			'Busca a data do pagamento da sele��o feita pelo usu�rio manual
			Set rs=objConexao.Execute("SELECT data_pagamento FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE cod_fechamento_periodo = "&cod_fechamento_periodo_manual)
				
				If not rs.Eof Then
					data_pagamento_manual = rs("data_pagamento")
				End If
			
			rs.Close
			Set rs=Nothing
			
			'Altera a data de recebimento
			
			cod_fechamento_periodo_manual_motivo = "Pedido teve o m�s de pagamento alterado de "&VENDEDOR_RECEBE_EM&" para "&data_pagamento_manual&" Motivo:<br />"&Ucase(cod_fechamento_periodo_manual_motivo)
			
			'1) Insere no Log
			objConexao.Execute("INSERT INTO TB_ADM_LOG_ALTERACOES (COD_USUARIO,COD_PEDIDO,DESCRICAO)VALUES("&Session("COD_USUARIO")&","&COD_PEDIDO&","&sqlTexto(cod_fechamento_periodo_manual_motivo)&")")
			
			'2) Coloca o valor nas vari�veis oficiais
			cod_fechamento_periodo = cod_fechamento_periodo_manual
			VENDEDOR_RECEBE_EM = data_pagamento_manual
			
		End If

	'////////////////////////////////////////////////
	
	sql = 	" UPDATE TB_PEDIDOS SET FINALIZA_CORRECAO = '1', "&_							
			" FORMA_PAGAMENTO = "&Request.Form("FORMA_PAGAMENTO")&", "&_
			" TIPO_VENDA = "&Request.Form("TIPO_VENDA")&", "&_
			" COD_GRUPO_VENDEDORES = '"&int(Request.Form("COD_GRUPO_VENDEDORES"))&"', "&_
			" DATA_CORRIGIDO_DEP_FOTO = NOW(), "&_
			" OBSERVACAO="&sqlTexto(Request.Form("OBSERVACAO"))&", "&_
			" VALOR_AJUDA_CUSTO = '"&sqlNumero(Request.Form("VALOR_AJUDA_CUSTO"))&"',"&_
			" CONFERIDO_POR_COD_USUARIO = '"&Session("COD_USUARIO")&"',"&_
			" cod_fechamento_periodo = "&cod_fechamento_periodo&","&_
			" VENDEDOR_RECEBE_EM = "&sqlData(VENDEDOR_RECEBE_EM)&_
			" WHERE COD_PEDIDO = "&COD_PEDIDO&" LIMIT 1"

		objConexao.Execute(sql)

		'Verifica se h� erro de log, se tiver gravar na base
		Dim DESCRICAO_LOG
		
		DESCRICAO_LOG = ""&Request.Form("ERROR_LOG")
		
		If DESCRICAO_LOG <> "" Then
			sql = "INSERT INTO TB_ADM_LOG_ALTERACOES (COD_USUARIO,DESCRICAO,COD_PEDIDO)VALUES("&Session("COD_USUARIO")&","&sqlTexto(DESCRICAO_LOG)&","&COD_PEDIDO&")"
			objConexao.Execute(sql)
		End If
		
		'Verifica se existe endere�o extra e tempor�rio para este pedido
		
		Dim rsEnderecoExtra					
		Set rsEnderecoExtra = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO_TEMPORARIA WHERE COD_PEDIDO="&COD_PEDIDO)	
		
			While not rsEnderecoExtra.Eof
			
				sql = 	"INSERT INTO TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO ("&_
						"COD_FORMANDO ,"&_
						"ENDERECO ,"&_
						"CEP ,"&_
						"FONE_RES ,"&_
						"FONE_CEL ,"&_
						"BAIRRO ,"&_
						"CIDADE ,"&_
						"ESTADO ,"&_
						"CADASTRADO_POR "&_
						")"&_
						"VALUES ("&_
						rsEnderecoExtra("COD_FORMANDO")&", "&sqlTexto(rsEnderecoExtra("ENDERECO"))&", "&sqlTexto(rsEnderecoExtra("CEP"))&", "&sqlTexto(rsEnderecoExtra("FONE_RES"))&", "&sqlTexto(rsEnderecoExtra("FONE_CEL"))&", "&sqlTexto(rsEnderecoExtra("BAIRRO"))&", "&sqlTexto(rsEnderecoExtra("CIDADE"))&", "&sqlTexto(rsEnderecoExtra("ESTADO"))&", 9998" &_
						")"
						objConexao.Execute(sql)
			
			rsEnderecoExtra.MoveNext()
			Wend
		
		rsEnderecoExtra.Close
		Set rsEnderecoExtra = Nothing
		
		'Delete casa haja endere�o endere�o extra para este pedido
		objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO_TEMPORARIA WHERE COD_PEDIDO = "&COD_PEDIDO)
		
		
		'3 -  Baixa nas movimenta��es
		'///////////////////////////////////////////////////////////////////////////
	If int(PEDIDO_EXTRA) = 2 Then
	'Caso n�o for pedido extra, baixa as movimenta��es
	
			Set rs331 = objConexao.Execute("SELECT DATA_CADASTRO FROM TB_PEDIDOS WHERE COD_PEDIDO = "&COD_PEDIDO)
		
			If not rs331.Eof Then
				Dim DATA_CADASTRO
				DATA_CADASTRO = formatDateTime(rs331("DATA_CADASTRO"))
			End If
		
		rs331.Close
		Set rs331 = Nothing
					
		'3-Seleciona todos os formandos do pedido
		Set rs333 = objConexao.Execute("SELECT * FROM TB_PEDIDOS_FORMANDOS WHERE COD_PEDIDO="&COD_PEDIDO)
		
			While not rs333.Eof
			
				
					'Somente movimenta��es em aberto!					
				Set rs222 = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES WHERE COD_FORMANDO="&rs333("COD_FORMANDO")&" AND DATA_RETORNO IS NULL")								

			
					While not rs222.Eof
																
						'Agora limpa as movimenta��es que est�o em aberto, cujos pedidos foram finalizados
						Dim TIPO_RS
						TIPO_RS = formatNumber(0&rs222("TIPO"),0)
							'Baixa o �LBUM e a FICHA caso estiverem em aberto
							If TIPO_RS < 2 Then
								sql = "UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET DATA_RETORNO = "&sqlData(DATA_CADASTRO)&" WHERE COD_MOVIMENTACAO ="&rs222("COD_MOVIMENTACAO")&" LIMIT 1 "
								
								objConexao.Execute(sql)		
							
							Elseif TIPO_RS > 1 Then
								
								'Verifica qual � o produto agregado se este estiver como vendido baixa nas movimenta��es
								Set rs9877 = objConexao.Execute("SELECT VENDIDO FROM TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS WHERE TIPO="&TIPO_RS&" AND COD_FORMANDO="&rs333("COD_FORMANDO"))
									
									If not rs9877.Eof Then
										
										Dim VENDIDO_RS
											VENDIDO_RS = formatNumber(0&rs9877("VENDIDO"),0)
											
											'Baixa o produto agregado somente se estiver vendido
											If VENDIDO_RS = 1 Then
													
													sql = "UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET DATA_RETORNO = "&sqlData(DATA_CADASTRO)&" WHERE COD_MOVIMENTACAO ="&rs222("COD_MOVIMENTACAO")&" LIMIT 1 "
													
													objConexao.Execute(sql)
												
											End If
											
									End If
									
								rs9877.Close
								Set rs9877 = Nothing
							
							End If
																		
					rs222.MoveNext	
					Wend								
		
				rs222.Close
				Set rs222 = Nothing
				
				
			
			rs333.MoveNext()
			Wend
		
		rs333.Close
		Set rs333 = Nothing
		



		'///////////////////////////////////////////////////////////////////////////
	
	End If
		'buscamos o cod_vendedor 		
		retorna = moduloPPVgerarOrdensServico(COD_PEDIDO,true)	

		dim cod_vendedor_9
		set rs9912 = objConexao.Execute("SELECT cod_vendedor from TB_PEDIDOS where cod_pedido = "&sqlNumero(cod_pedido))
		if not rs9912.eof then
			cod_vendedor_9 = rs9912(0)
		end if
		rs9912.close
		set rs9912 = Nothing
		'MOVIMENTA��ES NO ESTPOQUE
		'SELECIONAMOS TODOS OS FORMANDOS DO PEDIDO E BAIXAMOS OS SEUS PRODUTOS
		Set rsEst1 =objConexao.Execute("SELECT * FROM TB_PEDIDOS_PRODUTOS inner join TB_PRODUTOS on TB_PEDIDOS_PRODUTOS.COD_PRODUTO = TB_PRODUTOS.COD_PRODUTO WHERE TB_PEDIDOS_PRODUTOS.cod_pedido = "&sqlNumero(cod_pedido)&" and TB_PRODUTOS.SAIDA_AUTOMATICA = 2")
		if not rsEst1.Eof then
			While not rsEst1.Eof
				'verificamos qual estoque deve ser baixado
				if ((rsEst1("estoque_saida")&"") = "") or (rsEst1("estoque_saida") = "p") then
					call estoqueMovimentacao("S", "pedido", cod_pedido, rsEst1("COD_PRODUTO"), rsEst1("quantidade"), "Saida autom&aacite;tica, pedido:" & cod_pedido &" formando:"&rsEst1("cod_formando"), true)
				else

					call estoqueMovimentacaoVendedor("S", cod_vendedor_9, rsEst1("COD_PRODUTO"), rsEst1("quantidade"), "", "Venda Pedido :#"&cod_pedido, true)
				end if
			rsEst1.MoveNext
			Wend
		end if
		rsEst1.close
		set rsEst1 = Nothing

		
		'buscamos parcelas j� pagas
		'Se houver 
		'-verificamos se exitem parcelas n�o pagas'
		'-Se houver
		'--Recalculamos a comissao a ser paga por parcela
		'--Lan�amos o valor da comissao a ser paga na parcelas n�o pagas
		'-se nao
		'--nao fazemos nada
		'-fim se
		'-Lan�amos a data do pagamento da comiss�o do vendedor
		'Se n�o
		'-lan�amos o valor da comissao da 1 parcela de das demais usando a fun��o de fechamento
		'Fim se
		dim valor_comissao_vendedor_paga, qtd_parcelas_pagas
		dim valor_comissao_vendedor, valor_comissao_vendedor_total
		dim valor_comissao_vendedor_parcelas, valor_comissao_vendedor_parcela1

		set rs = objConexao.Execute("SELECT sum(valor_comissao_vendedor) as valor_comissao_vendedor, count(*) as qtd_parcelas_pagas FROM TB_PEDIDOS_BOLETOS WHERE data_pagamento is not null and cod_pedido = "&sqlNumero(cod_pedido))
		if not rs.Eof Then
			qtd_parcelas_pagas = int(0&rs("qtd_parcelas_pagas"))
			valor_comissao_vendedor_paga  = cdbl(0&rs("valor_comissao_vendedor"))
			
		
			set rs2 = objConexao.Execute("SELECT count(*) as qtd_parcelas_pagas FROM TB_PEDIDOS_BOLETOS WHERE data_pagamento is null and cod_pedido = "&sqlNumero(cod_pedido))
			if not rs.Eof Then
				qtd_parcelas_nao_pagas = int(0&rs2("qtd_parcelas_pagas"))
			end if
			rs2.Close
			set rs2 = Nothing

			'//verificamos se exitem parcelas n�o pagas
			if qtd_parcelas_nao_pagas > 0 then
				
				'//Recalculamos a comissao a ser paga por parcela
				'valor_comissao_vendedor_total = cdbl(0&fechamentoRetornaValorComissaoPedidoFormaPagamento(cod_pedido, "boleto.comissao_total", false, true))
				'valor_comissao_vendedor = valor_comissao_vendedor_total - valor_comissao_vendedor_paga

				'valor_comissao_vendedor_parcelas = round(valor_comissao_vendedor / qtd_parcelas_nao_pagas,2)
				'valor_comissao_vendedor_parcela1 = valor_comissao_vendedor_parcelas + (valor_comissao_vendedor - (valor_comissao_vendedor_parcelas * qtd_parcelas_nao_pagas)  )
				'//Fim Recalculamos a comissao a ser paga por parcela
			
				'//Lan�amos o valor da comissao a ser paga na parcelas n�o pagas
				'objConexao.Execute("UPDATE TB_PEDIDOS_BOLETOS set valor_comissao_vendedor = "&sqlNumero(valor_comissao_vendedor_parcelas)&" WHERE cod_pedido = "&sqlNumero(cod_pedido)&" AND data_pagamento is null ")

				'/comissao da primeiroa carcela'
				set rs2 = objConexao.Execute("select max(parcela)+1 as parcela_1 from TB_PEDIDOS_BOLETOS where cod_pedido = "&sqlNumero(cod_pedido)&" and data_pagamento is not null")
				if not rs2.eof then
					'objConexao.Execute("UPDATE TB_PEDIDOS_BOLETOS set valor_comissao_vendedor = "&sqlNumero(valor_comissao_vendedor_parcela1)&" WHERE cod_pedido = "&sqlNumero(cod_pedido)&" AND data_pagamento is null and parcela = "&sqlNumero(0&rs2("parcela_1")))
				end if
				rs2.close
				set rs2 = Nothing
				'//Fim Lan�amos o valor da comissao a ser paga na parcelas n�o pagas


			end if
			'//fim verificamos se exitem parcelas n�o pagas

			'Lan�amos a data do pagamento da comiss�o do vendedor 
			set rsBol1 = objConexao.Execute("SELECT * FROM TB_PEDIDOS_BOLETOS WHERE cod_pedido = "&sqlNumero(cod_pedido)&" AND data_pagamento is not null")
			if not rsBol1.eof then
				'objConexao.Execute("UPDATE TB_PEDIDOS_BOLETOS SET cod_fechamento_periodo = "&sqlNumero(cod_fechamento_periodo)&", vendedor_recebe_em = "&sqlData(VENDEDOR_RECEBE_EM)&" WHERE cod_boleto = "&rsBol1("cod_boleto"))
			end if
			rsBol1.close
		else

			'valor_comissao_vendedor_parcelas = cdbl(0&fechamentoRetornaValorComissaoPedidoFormaPagamento(cod_pedido, "boleto.comissao_parcelas", false, true))
			'valor_comissao_vendedor_parcela1 = cdbl(0&fechamentoRetornaValorComissaoPedidoFormaPagamento(cod_pedido, "boleto.comissao_parcela1", false, true))

			'objConexao.Execute("UPDATE TB_PEDIDOS_BOLETOS set valor_comissao_vendedor = "&sqlNumero(valor_comissao_vendedor_parcelas)&" WHERE cod_pedido = "&sqlNumero(cod_pedido)&"  ")
			'objConexao.Execute("UPDATE TB_PEDIDOS_BOLETOS set valor_comissao_vendedor = "&sqlNumero(valor_comissao_vendedor_parcela1)&" WHERE cod_pedido = "&sqlNumero(cod_pedido)&" AND  parcela  = 1 ")

		end if
		rs.Close
		set rs = Nothing

		msg_status = "Pedido Finalizado com sucesso."
		
	End If '//Pode finalizar

	FecharConexao()
	Response.Redirect("lca_pedidos_pendentes.asp?cod_pedido="&COD_PEDIDO&"&msg="&msg_status)

End If

Dim rsVALOR_TOTAL, VALOR_TOTAL, PAGAMENTO_EM_DINHEIRO, XTOTAL_CHEQUE, XTOTAL_CARTAO
AbrirConexao()
	Set rsVALOR_TOTAL = objConexao.Execute("SELECT PAGAMENTO_EM_DINHEIRO, XTOTAL_CHEQUE, XTOTAL_CARTAO, XTOTAL_BOLETO FROM VW_PEDIDOS_TOTAL_RECEBIDO WHERE COD_PEDIDO = "&COD_PEDIDO&" LIMIT 1")
		If not rsVALOR_TOTAL.Eof Then
			PAGAMENTO_EM_DINHEIRO = formatNumber(0&rsVALOR_TOTAL("PAGAMENTO_EM_DINHEIRO"))
			XTOTAL_CHEQUE = formatNumber(0&rsVALOR_TOTAL("XTOTAL_CHEQUE"))
			XTOTAL_CARTAO = formatNumber(0&rsVALOR_TOTAL("XTOTAL_CARTAO"))
			XTOTAL_BOLETOS = formatNumber(0&rsVALOR_TOTAL("XTOTAL_BOLETO"))
			VALOR_TOTAL = (PAGAMENTO_EM_DINHEIRO/1 + XTOTAL_CHEQUE/1 +  XTOTAL_CARTAO/1 + XTOTAL_BOLETOS/1)
		End IF
	rsVALOR_TOTAL.Close()
	Set rsVALOR_TOTAL = Nothing
FecharConexao()
%>

<!-- include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Vers�o 1.2</title>
<style type="text/css">
	/* CSS for the demo. CSS needed for the scripts are loaded dynamically by the scripts */
	.btnFinalizaCorrecao{	background-image:url(images/layout/btnFinalizaCorrecao.jpg);
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
	
	/* Alternative layout for the message box - used by one of the messages */
	
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

	</style>	
	<script type="text/javascript">
	function verify(ver){
		
		if(ver){
			// Confirmed message, i.e. clicked on "Yes"
			alert('Message confirmed');
		}else{
			// Clicked on "No"
			alert('Message not confirmed');
		}
	}
	</script>
	<link rel="stylesheet" href="css/modal-message.css" type="text/css">
	
	
	
	<script type="text/javascript" src="js/ajax.js"></script>
	<script type="text/javascript" src="js/modal-message.js"></script>
	<script type="text/javascript" src="js/ajax-dynamic-content.js"></script>
	<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>
	
	
	
	<script type="text/javascript">
messageObj = new DHTML_modalMessage();	// We only create one object of this class
messageObj.setShadowOffset(5);	// Large shadow


function displayMessage(url)
{
	messageObj.setSource(url);
	messageObj.setCssClassMessageBox(false);
	messageObj.setSize(550,450);
	messageObj.setShadowDivVisible(true);	// Enable shadow for these boxes
	messageObj.display();
}

function displayStaticMessage(messageContent,cssClass,altura)
{
	messageObj.setHtmlContent(messageContent);
	messageObj.setSize(550,altura);
	messageObj.setCssClassMessageBox(cssClass);
	messageObj.setSource(false);	// no html source since we want to use a static message here.
	messageObj.setShadowDivVisible(false);	// Disable shadow for these boxes	
	messageObj.display();
	
	
}

function closeMessage()
{
	messageObj.close();	
}


</script>
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
	<!--AJUDA F1 E CTRL PARA BUSCA R�PIDA-->
		
	function abrir(theURL,winName,features) 
	{ 
	  window.open(theURL,winName,features);
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
		<!-- CALEND�RIO SCRIPTS
	<script type="text/javascript">
	DHTMLSuite.include("calendar");
	</script>
	<script type="text/javascript">
	var calendarObjForForm = new DHTMLSuite.calendar({minuteDropDownInterval:10,numberOfRowsInHourDropDown:5,callbackFunctionOnDayClick:'getDateFromCalendar',isDragable:true,displayTimeBar:false});
	
		 
	calendarObjForForm.setCallbackFunctionOnClose('myOtherFunction');
	
	function myOtherFunction()
	{
		
		
	}
	function pickDate(buttonObj,inputObject)
	{
		calendarObjForForm.setCalendarPositionByHTMLElement(inputObject,0,inputObject.offsetHeight+2);	// Position the calendar right below the form input
		//calendarObjForForm.setInitialDateFromInput(inputObject,'yyyy-mm-dd hh:ii');	// Specify that the calendar should set it's initial date from the value of the input field.
		
		calendarObjForForm.setInitialDateFromInput(inputObject,'dd/mm/yyyy');	// Specify that the calendar should set it's initial date from the value of the input field.
		
		calendarObjForForm.addHtmlElementReference('myDate',inputObject);	// Adding a reference to this element so that I can pick it up in the getDateFromCalendar below(myInput is a unique key)
		if(calendarObjForForm.isVisible()){
			calendarObjForForm.hide();
		}else{
			calendarObjForForm.resetViewDisplayedMonth();	// This line resets the view back to the inital display, i.e. it displays the inital month and not the month it displayed the last time it was open.
			calendarObjForForm.display();
		}		
	}	
	/* inputArray is an associative array with the properties
	year
	month
	day
	hour
	minute
	calendarRef - Reference to the DHTMLSuite.calendar object.
	*/
	function getDateFromCalendar(inputArray)
	{
		var references = calendarObjForForm.getHtmlElementReferences(); // Get back reference to form field.
		references.myDate.value = inputArray.day + '/' + inputArray.month + '/' + inputArray.year
//		references.myDate.value = inputArray.year + '-' + inputArray.month + '-' + inputArray.day + ' ' + inputArray.hour + ':' + inputArray.minute;
		calendarObjForForm.hide();	
		
	}	
	</script>
	-->
<script>
function confirma() {
	
	if(YNconfirm("Mensagem")){
			document.dados.submit();			
	}	
}
</script>
<script language="javascript">
function YNconfirm(txt){
return typeof(suporteVBscript)=="undefined"?confirm(txt):VBconfirm(txt)==6
}
</script>
<script language="vbscript">
suporteVBscript=1
function VBconfirm(mensagem)

 VBconfirm=msgbox(mensagem,vbyesno)
end function

</script>
<script>
function validaForm(){
	d = document.finalizacao;
	if (d.cod_fechamento_periodo_manual.value !== '' && d.cod_fechamento_periodo_manual_motivo.value == ''){
		alert('Voc&ecirc; selecionou um M&ecirc;s de pagamento manual, por&eacute;m n&atilde;o informou o motivo da altera&ccedil;&atilde;o manual.');
		d.cod_fechamento_periodo_manual.focus();
		return false;
	}else if (d.motivo_solicitacao.value == '' && d.NECESSARIO_AUTORIZACAO_ESPECIAL.value == 'S'){
		alert('Voc&ecirc; optou por finalizar uma venda que est&aacute; com os pre&ccedil;os incorretos, a mesma n&atilde;o ser&aacute; finalizada se a aprova&ccedil;&atilde;o n&atilde;o for realizada, informe o motivo dos pre&ccedil;os abaixo da tabela.');
		d.motivo_solicitacao.focus();
		return false;
	}else{
		return true;
	}
}
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  if (selObj.value !== ""){
  	eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  	if (restore) selObj.selectedIndex=0;
  }
}
</script>
<script>
function Limpar(valor, validos) {
// retira caracteres invalidos da string
var result = "";
var aux;
for (var i=0; i < valor.length; i++) {
aux = validos.indexOf(valor.substring(i, i+1));
if (aux>=0) {
result += aux;
}
}
return result;
}

//Formata n�mero tipo moeda usando o evento onKeyDown

function Formata(campo,tammax,teclapres,decimal) {
var tecla = teclapres.keyCode;
vr = Limpar(campo.value,"0123456789");
tam = vr.length;
dec=decimal

if (tam < tammax && tecla != 8){ tam = vr.length + 1 ; }

if (tecla == 8 )
{ tam = tam - 1 ; }

if ( tecla == 8 || tecla >= 48 && tecla <= 57 || tecla >= 96 && tecla <= 105 )
{

if ( tam <= dec )
{ campo.value = vr ; }

if ( (tam > dec) && (tam <= 5) ){
campo.value = vr.substr( 0, tam - 2 ) + "," + vr.substr( tam - dec, tam ) ; }
if ( (tam >= 6) && (tam <= 8) ){
campo.value = vr.substr( 0, tam - 5 ) + "." + vr.substr( tam - 5, 3 ) + "," + vr.substr( tam - dec, tam ) ; 
}
if ( (tam >= 9) && (tam <= 11) ){
campo.value = vr.substr( 0, tam - 8 ) + "." + vr.substr( tam - 8, 3 ) + "." + vr.substr( tam - 5, 3 ) + "," + vr.substr( tam - dec, tam ) ; }
if ( (tam >= 12) && (tam <= 14) ){
campo.value = vr.substr( 0, tam - 11 ) + "." + vr.substr( tam - 11, 3 ) + "." + vr.substr( tam - 8, 3 ) + "." + vr.substr( tam - 5, 3 ) + "," + vr.substr( tam - dec, tam ) ; }
if ( (tam >= 15) && (tam <= 17) ){
campo.value = vr.substr( 0, tam - 14 ) + "." + vr.substr( tam - 14, 3 ) + "." + vr.substr( tam - 11, 3 ) + "." + vr.substr( tam - 8, 3 ) + "." + vr.substr( tam - 5, 3 ) + "," + vr.substr( tam - 2, tam ) ;}
} 

}

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
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">

	  <h2>LCA - Pedidos Pendentes / Corre&ccedil;&atilde;o</h2>
	  <h3 align="center">Corre&ccedil;&atilde;o de Pedido - Finaliza&ccedil;&atilde;o</h3>
	  <form style="margin:0px" method="post" name="finalizacao" onsubmit="return validaForm()" action="lca_pedidos_pendentes_alteracoes_part_4.asp?cod_pedido=<%=COD_PEDIDO%>">
		  	<table width="98%" align="center" 	cellpadding="0" cellspacing="3">
			<tr>
            <td width="12%"><div align="left"><strong>Vendedor</strong>:</div></td>
            <td width="88%"><%
			AbrirConexao()
			'Nome do vendedor
			Dim rsVendedorQ, nome_vendedorq, COD_GRUPO_VENDEDORES
			Set rsVendedorQ = objConexao.Execute("SELECT TB_VENDEDORES.NOME_VENDEDOR, TB_VENDEDORES.COD_GRUPO_VENDEDORES, TB_PEDIDOS.PEDIDO_EXTRA FROM TB_VENDEDORES INNER JOIN TB_PEDIDOS ON TB_VENDEDORES.COD_VENDEDOR = TB_PEDIDOS.COD_VENDEDOR WHERE TB_PEDIDOS.COD_PEDIDO="&COD_PEDIDO&" LIMIT 1")
				
				If not rsVendedorQ.Eof Then
					Response.Write(rsVendedorQ("NOME_VENDEDOR"))
					nome_vendedorq = rsVendedorQ("NOME_VENDEDOR")
					PEDIDO_EXTRA = rsVendedorQ("PEDIDO_EXTRA")
					COD_GRUPO_VENDEDORES = rsVendedorQ("COD_GRUPO_VENDEDORES")
				End If
			
			rsVendedorQ.Close
			Set rsVendedorQ = Nothing
			
			%></td>
            </tr>
			<tr>
			  <td><strong>Formando(s):</strong></td>
			  <td><%
			sql = 	"SELECT TB_PEDIDOS_FORMANDOS.COD_FORMANDO,TB_PEDIDOS.DATA_VENDA, TB_PEDIDOS.QTD_PARCELAS,TB_FICHAS_IDENTIFICACAO.SEMESTRE_CONCLUSAO, TB_FICHAS_IDENTIFICACAO.ANO_CONCLUSAO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO "&_
					"FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON "&_
					"TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO "&_
					"INNER JOIN TB_PEDIDOS ON TB_PEDIDOS_FORMANDOS.COD_PEDIDO = TB_PEDIDOS.COD_PEDIDO "&_
					"WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&COD_PEDIDO&" ORDER BY NOME_FORMANDO"
			Set rs3 = objConexao.Execute(sql)
				
				While not rs3.Eof
					QTD_PARCELAS = rs3("QTD_PARCELAS")
					DATA_VENDA = rs3("DATA_VENDA")
				%>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				  <tr>
					<td width="46%"><div align="left"><%=rs3("NOME_FORMANDO")%> (<%=digito(rs3("COD_FORMANDO"))%>)</div></td>
					<td width="54%">Conclus&atilde;o em: <strong><%=rs3("SEMESTRE_CONCLUSAO")%>/<%=rs3("ANO_CONCLUSAO")%></strong></td>
				  </tr>
				</table>
				<%
				rs3.MoveNext()
				Wend
				
			rs3.Close
			Set rs3 = Nothing
			FecharConexao()
			%></td>
			  </tr>
			</table>
			<p align="center">
			
			<input name="VALIDADOR" value="1" type="hidden" />
			<table width="98%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">             
					<tr>
						<td width="23%"><div align="left"><strong>Data do pagamento:</strong></div></td>
						<td width="77%"><div align="left">
						  <p></p>
                          <%						  
						  cod_fechamento_periodo = fechamentoRetornaCodPeriodo(cod_pedido,false)
						  temp = fechamentoRetornaCodPeriodoVerificaErro(cod_fechamento_periodo)
						  If lcase(temp) = "true" Then
							  VENDEDOR_RECEBE_EM = fechamentoRetornaData(4,cod_fechamento_periodo,false)
						  	'N�o possui erro
						  Else
						  	POSSUI_ERRO = fechamentoRetornaCodPeriodoVerificaErro(cod_fechamento_periodo)
						  End If
						  
						  %>
                          
                          <strong><span class="textoMedio"><%=ArrumaData(VENDEDOR_RECEBE_EM)%></span><br /><span class="msg"><%=msg_erro_readequacao_vendedor_recebe_em%></span></strong>
                          <input type="hidden" name="cod_fechamento_periodo" value="<%=cod_fechamento_periodo%>" />
                          <input type="hidden" name="VENDEDOR_RECEBE_EM" value="<%=VENDEDOR_RECEBE_EM%>" />
                          <input type="hidden" name="ERROR_LOG" value="<%=msg_erro_readequacao_vendedor_recebe_em%>" />
			<%
				If  validaUsuario(2,"",79,false) or validaUsuario(2,"",85,false) Then	
			%>
			<table width="340" border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFCC">
                           <tr>
                              <td colspan="2"><div align="center">
                               <strong>Para alterar data de recebimento manual</strong>
                               <br />Obs: Se voc&ecirc; optar por alterar, ser&aacute; gravado no log de altera&ccedil;&otilde;es
                              </div></td>
                              </tr>
                             
                            <tr>
                              <td width="76">Pagamento *</td>
                              <td width="244">
							    <div align="left">                                
							<select name="cod_fechamento_periodo_manual" class="inputTamanho2">
                          	<option value=""></option>
        <%
		AbrirConexao()
		Set rs2=objConexao.Execute("SELECT cod_fechamento_periodo, data_pagamento FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE data_processado_pagamento IS NULL AND cod_fechamento_periodo <> '"&cod_fechamento_periodo&"' ORDER BY data_pagamento ASC ")
								
		If not rs2.Eof Then
				While not rs2.Eof
					%>
					<option <%if int(cod_fechamento_periodo) = int(rs2(0)) Then%> selected="selected" <%End If%> value="<%=rs2(0)%>"><%=ArrumaData(rs2(1))%></option>
					<%
				rs2.MoveNext()
				Wend	
		End If
								
	rs2.Close
	Set rs2=Nothing
	FecharConexao()
	%>
     </select>
	</div></td>
     </tr>
                            <tr>
                              <td>Motivo *</td>
                              <td><div align="left">
                                <textarea class="inputTamanho2" style="height:70px;" name="cod_fechamento_periodo_manual_motivo" id="textarea" cols="45" rows="5"></textarea>
                              </div></td>
                            </tr>
                          </table>
<%
	end if
%>
                    
						  <p>  </p>
						</div></td>
						</tr>
                        
                        <tr>
					  <td><div align="left"><strong>Valor estimado de comiss&atilde;o: <a href="javascript:alert('Valor estimado se o fechamento fosse hoje e n&atilde;o houvesse altera&ccedil;&otilde;es no pedido e na tabela de comissionamento.\nObs: Valor poder&aacute; ser alterado.')"><img src="images/icons/icon_help_mini.jpg" width="15" height="11" border="0" /></a> </strong></div></td>
					  <td><div align="left">Detalhamento: <br />
						<%
						'fechamentoRetornaValorComissaoPedidoFormaPagamento(cod_pedido, tipo, debug, conexao_aberta)
						%>
                      <%=fechamentoRetornaValorComissaoPedido(cod_pedido,2,false)%><br />
Total: <strong><%=formatCurrency(formatNumber(0&fechamentoRetornaValorComissaoPedido(cod_pedido,1,false)))%></strong>
		<br />
						Total (dinheiro, cart&atilde;o, cheque): <%=formatCurrency(fechamentoRetornaValorComissaoPedidoFormaPagamento(cod_pedido, "outros.comissao_total", false, false))%>
		<br />
						Total (boleto): <%=fechamentoRetornaValorComissaoPedidoFormaPagamento(cod_pedido, "boleto.descritivo", false, false)%>

                      </div></td>
					  </tr>
                    <tr>
					<tr>
					  <td><div align="left"><strong>Data da Venda: </strong></div></td>
					  <td><div align="left"><%=ArrumaData(DATA_VENDA)%> - <a href="lca_pedidos_pendentes_alteracoes_part_1.asp?cod_pedido=<%=COD_PEDIDO%>">Alterar Data da Venda</a></div></td>
					  </tr>
                    <tr>
					  <td><div align="left"><strong>Forma de Pagamento:<a href="javascript:alert('Op&ccedil;&otilde;es: &agrave; vista ou &agrave; prazo')"><img src="images/icons/icon_help_mini.jpg" width="15" height="11" border="0" /></a> </strong></div></td>
					  <td><div align="left">
					    <%
						FORMA_PAGAMENTO = fechamentoRetornaFormaPagamentoPedido(data_venda,COD_PEDIDO,false)
						
						If int(FORMA_PAGAMENTO) = 1 Then
							Response.Write("&Agrave; vista")
						Elseif int(FORMA_PAGAMENTO) = 2 Then
							Response.Write("&Agrave; prazo")
						End If
						%>
                        <input value="<%=FORMA_PAGAMENTO%>" name="FORMA_PAGAMENTO"  type="hidden" />
                        </div></td>
					  </tr>
                      <tr>
					  <td><div align="left"><strong>Qtd de Parcelas:<a href="javascript:alert('Op&ccedil;&otilde;es: Qtd de Parcelas totais')"><img src="images/icons/icon_help_mini.jpg" width="15" height="11" border="0" /></a> </strong></div></td>
					  <td><div align="left">
					    <%=QTD_PARCELAS%>
                        </div></td>
					  </tr>
                      <tr>
					  <td><div align="left"><strong>Produto com valor correto:<a href="javascript:alert('Op&ccedil;&otilde;es: Para os novos pedidos, h&aacute; valida&ccedil;&atilde;o de pre&ccedil;o se o pre&ccedil;o de algum item estiver abaixo do valor da tabela ou boletim informativo, ser&aacute; necess&aacute;rio aprova&ccedil;&atilde;o especial.')"><img src="images/icons/icon_help_mini.jpg" width="15" height="11" border="0" /></a> </strong></div></td>
					  <td><div align="left">
                      <%If pedidoProdutoValorCorreto(cod_pedido,false) Then
						Response.Write("Pre&ccedil;os corretos")
					  %>
							<input type="hidden" name="NECESSARIO_AUTORIZACAO_ESPECIAL" value="N" />
                            <input type="hidden" name="PODE_FINALIZAR" value="S" />
					   <%Else%>
                            <%
							'Verifica se este pedido possui aprova��o pendente
							AbrirConexao()
							
							Set rs=objConexao.Execute("SELECT * FROM TB_APROVACOES WHERE parametro = "&sqlTexto(cod_pedido))
								If not rs.Eof Then%>
                                	<input type="hidden" name="NECESSARIO_AUTORIZACAO_ESPECIAL" value="N" />
									<table width="100%" border="1" cellspacing="0" cellpadding="3">
                                      <tr>
                                        <td width="22%"><strong>Status</strong></td>
                                        <td width="78%">
                                        <%
                                        If int(rs("status")) = 1 Then
                                            Response.Write("Aprovado")%>
                                            <input type="hidden" name="PODE_FINALIZAR" value="S" />
                                        <%Elseif int(rs("status")) = 2 Then%>
                                        	<input type="hidden" name="PODE_FINALIZAR" value="N" />
                                        <%
                                            Response.Write("Reprovado")
                                        Else%>
                                        	<input type="hidden" name="PODE_FINALIZAR" value="N" />
                                        <%
                                            Response.Write("Aguardando aprova&ccedil;&atilde;o")	
                                        End If
                                        %>                                        </td>
                                      </tr>
                                      				<tr>
    								<td><strong>Solicitado por</strong></td>
    			                    <td>
                                    <%
									Set rs2=objConexao.Execute("Select nome FROM TB_ADM_USUARIOS WHERE COD_USUARIO = "&rs("solicitante_cod_usuario"))
										
										if not rs2.Eof Then
											Response.Write(rs2(0))
										End if
									
									rs2.Close
									Set rs2=Nothing
									%>                                    </td>
  </tr>
  									<tr>
    								<td><strong>Motivo da solicita&ccedil;&atilde;o</strong></td>
    			<td>
    	<%
		If int(rs("status")) = 0 Then
		'Pode alterar, pois n�o foi aprovado ainda
		%>
    		<textarea name="motivo_solicitacao" cols="40" rows="3"><%=rs("motivo_solicitacao")%></textarea>
        <%
		Else
			Response.Write(Replace(rs("motivo_solicitacao"),Chr(13),"<br />"))
        End If%>    </td>
  </tr>
  <tr>
    <td><strong>Data solicitado</strong></td>
    <td><%=ArrumaData(rs("data_cadastro"))%></td>
  </tr>
  <tr>
    <td><strong>Data aprovado/reprovado</strong></td>
    <td><%=ArrumaData(rs("data_finalizado"))%></td>
  </tr>
  <tr>
    <td><strong>Respons&aacute;vel aprova&ccedil;&atilde;o</strong></td>
    <td>&Eacute;lcio Freitas</td>
  </tr>
  <tr>
    <td><strong>Mensagem do respons&aacute;vel</strong></td>
    <td><%=rs("motivo_responsavel")%></td>
  </tr>
</table>
                                <%Else%>
                               	  <input type="hidden" name="PODE_FINALIZAR" value="N" />
                                	<input type="hidden" name="NECESSARIO_AUTORIZACAO_ESPECIAL" value="S" />
								  <div style=" color:#CC0000"><strong>Pre&ccedil;os Incorretos, corriga ou ser&aacute; necess&aacute;rio aprova&ccedil;&atilde;o especial para finalizar venda, informe o motivo abaixo dos valores abaixo da tabela</strong>.</div>
                                    <table width="100%" border="0" cellspacing="0" cellpadding="3">
  									<tr>
    								<td width="22%" bgcolor="#E6E6E6">Motivo da solicita&ccedil;&atilde;o</td>
    			<td width="78%"><textarea name="motivo_solicitacao" cols="40" rows="3"></textarea></td>
  </tr>
</table>
                                    
                                <%
								End if
								
							rs.Close
							Set rs=Nothing
							
							FecharConexao()
						
					   End If
					  %>
                      </div></td>
					  </tr>
                      
                      
                    <tr>
					  <td><div align="left"><strong>Tipo da Venda:<a href="javascript:alert('Op��es: Normal, Repasse, Repasse Especial')"><img src="images/icons/icon_help_mini.jpg" width="15" height="11" border="0" /></a> </strong></div></td>
					  <td><div align="left">
                     
                      <%

                      AbrirConexao()
					dim rsTipovenda
					set rsTipovenda = objConexao.Execute("select * from TB_PEDIDOS_FORMANDOS WHERE cod_pedido = "&cod_pedido, qtd_formandos)
						while not rsTipovenda.Eof


							 TIPO_VENDA = fechamentoRetornaTipoVenda(cod_pedido,true, rsTipovenda("cod_formando") )
					  
							  If int(TIPO_VENDA) = 1 Then
							  	Response.Write("Normal")
							  Elseif int(TIPO_VENDA) = 2  Then
							  	Response.Write("Repasse")
							  Elseif int(TIPO_VENDA) = 3  Then
							  	Response.Write("Repasse Especial")
							  End If

							  if qtd_formandos > 1 then
							  	response.write(" / ")
							  	qtd_formandos = qtd_formandos - 1
							  end if


						rsTipovenda.MoveNext
						Wend
					rsTipovenda.close
					set rsTipovenda = Nothing


					 FecharConexao()
					  %>
                       <input value="<%=TIPO_VENDA%>" name="TIPO_VENDA"  type="hidden" />
					  </div></td>
					  </tr>
                      <%
					  If int(COD_GRUPO_VENDEDORES) > 0 Then
					  %>
                   		<tr>
					  <td><div align="left"><strong>Grupo Repasse de: </strong></div></td>
					  <td><div align="left"><input value="<%=COD_GRUPO_VENDEDORES%>" name="COD_GRUPO_VENDEDORES"  type="hidden" /> 
					    <%=NOME_VENDEDOR_GRUPO_ESPECIAL%></div></td>
					  </tr> 
                      <%
					  End If
					  %>
					<tr>
					  <td><div align="left"><strong>Ajuda de Custo: </strong></div></td>
					  <td><div align="left">
					    <input name="VALOR_AJUDA_CUSTO" onKeydown="Formata(this,20,event,2)" class="inputTamanho1" type="text" size="40" maxlength="15" value="<%=formatarNumero(formatNumber(VALOR_AJUDA_CUSTO))%>" />
					  </div></td>
					  </tr>
					<tr>
					  <td valign="top"><div align="left"><strong>Observa&ccedil;&atilde;o:</strong></div></td>
					  <td><textarea name="OBSERVACAO" cols="40" rows="3"><%=OBSERVACAO%></textarea></td>
					  </tr>
                      <%If int(PEDIDO_EXTRA) = 2 Then%>
							<tr>
					  <td><div align="left"><strong>Total em Produtos:</strong></div></td>
					  <td><%
						Dim rsProdutosTotal, QTD_EM_PRODUTOS, XTOTAL_AGREGADOS, XTOTAL_ALBUNS, XTOTAL_PRODUTOS_A_PARTE, XTOTAL_DESCONTO
						AbrirConexao()
						
						QTD_EM_PRODUTOS = 0
						XTOTAL_AGREGADOS = 0
						XTOTAL_ALBUNS = 0
						XTOTAL_PRODUTOS_A_PARTE = 0
						XTOTAL_DESCONTO = 0
						
							'VW_PEDIDOS_PRODUTOS_TOTAL_AGREGADOS
							Set rsProdutosTotal = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_AGREGADOS WHERE COD_PEDIDO="&COD_PEDIDO)
							
								if not rsProdutosTotal.Eof Then
									
									XTOTAL_AGREGADOS = formatNumber(0&rsProdutosTotal("XTOTAL_AGREGADOS"))/1
								
								End If
							
							
							rsProdutosTotal.Close()
							Set rsProdutosTotal = Nothing
							
							'VW_PEDIDOS_PRODUTOS_TOTAL_ALBUNS
							Set rsProdutosTotal = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_ALBUNS WHERE COD_PEDIDO="&COD_PEDIDO)
							
								if not rsProdutosTotal.Eof Then
									
									XTOTAL_ALBUNS = formatNumber(0&rsProdutosTotal("XTOTAL_ALBUNS"))/1
								
								End If
							
							
							rsProdutosTotal.Close()
							Set rsProdutosTotal = Nothing
							
							'VW_PEDIDOS_PRODUTOS_TOTAL_PRODUTOS_A_PARTE
							Set rsProdutosTotal = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_PRODUTOS_A_PARTE WHERE COD_PEDIDO="&COD_PEDIDO)
							
								if not rsProdutosTotal.Eof Then
									
									XTOTAL_PRODUTOS_A_PARTE = formatNumber(0&rsProdutosTotal("XTOTAL_PRODUTOS_A_PARTE"))/1
								
								End If
							
							
							rsProdutosTotal.Close()
							Set rsProdutosTotal = Nothing
							
							set rsProdutosTotal = objConexao.Execute("select SUM(valor_total_com_desconto) as valor_total_com_desconto from(SELECT CASE WHEN "&PEDIDO_EXTRA&" = 2 THEN SUM(faturado * (valor_un-(valor_un * (1-(ifnull(aliquota_desconto_foto,0)/100))))) ELSE '0' END as valor_total_com_desconto FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO in (SELECT cod_formando FROM TB_PEDIDOS_FORMANDOS WHERE cod_pedido =  "&cod_pedido&") union all SELECT SUM(quantidade * (valor_un-(valor_un * (1-(ifnull(aliquota,0)/100))))) as valor_total_com_desconto  FROM TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO = "&cod_pedido&") as total_desconto")
							if not rsProdutosTotal.Eof then
								XTOTAL_DESCONTO = formatCurrency(0&rsProdutosTotal(0))/1
							end if
							rsProdutosTotal.Close
							set rsProdutosTotal = Nothing


												
							QTD_EM_PRODUTOS = formatCurrency(XTOTAL_AGREGADOS + XTOTAL_ALBUNS + XTOTAL_PRODUTOS_A_PARTE  - XTOTAL_DESCONTO)
						%>
      <%=QTD_EM_PRODUTOS%> - <a href="lca_pedidos_pendentes_alteracoes_part_2.asp?cod_pedido=<%=COD_PEDIDO%>">Alterar Produtos</a></td>
					  </tr>
					<tr>
					  <td><div align="left"><strong>Total em Forma de Pagamento:</strong></div></td>
					  <td>
						<div align="left"><%=formatCurrency(VALOR_TOTAL)%> - <a href="lca_pedidos_pendentes_alteracoes_part_3.asp?cod_pedido=<%=COD_PEDIDO%>">Alterar Forma de Pagamento</a></div></td>
					  </tr>
					  
	<%

					  
  	'Cr�dito de formando, quando h� algum saldo residual por algum motivo vindo da formatura.
  	Dim TOTAL_CREDITO
	TOTAL_CREDITO = 0
	
  	sql =	" SELECT SUM(IFNULL(TB_FICHAS_IDENTIFICACAO.CREDITO_EM_VALOR,0)) AS TOTAL_CREDITO "&_
			" FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON "&_
			" TB_PEDIDOS_FORMANDOS.COD_FORMANDO =  "&_
			" TB_FICHAS_IDENTIFICACAO.COD_FORMANDO WHERE  "&_
			" TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&COD_PEDIDO&" GROUP BY TB_PEDIDOS_FORMANDOS.COD_PEDIDO"
	
	Set rsCredito = objConexao.Execute(sql) 
					
					If not rsCredito.Eof Then
						TOTAL_CREDITO = formatNumber(0&rsCredito("TOTAL_CREDITO"))
					End If
					
	rsCredito.Close
	Set rsCredito = Nothing
  
  'SOMA DE QUALQUER FORMA E ALTERA O VALOR TOTAL (DEFAULT CR�DITO = "0")
  VALOR_TOTAL = (VALOR_TOTAL + TOTAL_CREDITO)/1
  
  
  						If TOTAL_CREDITO > 0 Then %>
					  <tr>
					  <td><div align="left"><strong>Total em Cr&eacute;dito:</strong></div></td>
					  <td>
						<div align="left"><%=formatCurrency(TOTAL_CREDITO)%></div></td>
					  </tr>
					  <tr bgcolor="#E2E2E2">
					  <td><div align="left"><strong>Total em Pagamento + Cr&eacute;dito:</strong></div></td>
					  <td>
						<div align="left"><%=formatCurrency(VALOR_TOTAL)%></div></td>
					  </tr>
                    	<%End If%>
                    
                      <%Else%>
							<tr>
					  <td><div align="left"><strong>Total em Produtos:</strong></div></td>
					  <td><%
						AbrirConexao()
						
						QTD_EM_PRODUTOS = 0
						
							Set rsProdutosTotal = objConexao.Execute("SELECT * FROM VW_PEDIDOS_PRODUTOS_TOTAL_PRODUTOS WHERE COD_PEDIDO="&COD_PEDIDO)
							
								if not rsProdutosTotal.Eof Then
									

									XTOTAL_PRODUTOS_A_PARTE = formatNumber(0&rsProdutosTotal("XTOTAL_PRODUTOS_A_PARTE"))/1
									
									QTD_EM_PRODUTOS = formatCurrency(XTOTAL_PRODUTOS_A_PARTE)
									
								Else
									
									QTD_EM_PRODUTOS = 0
									
								End If
							
							
							rsProdutosTotal.Close()
							Set rsProdutosTotal = Nothing
							
							set rsProdutosTotal = objConexao.Execute("select SUM(valor_total_com_desconto) as valor_total_com_desconto from(SELECT CASE WHEN "&PEDIDO_EXTRA&" = 2 THEN SUM(faturado * (valor_un-(valor_un * (1-(ifnull(aliquota_desconto_foto,0)/100))))) ELSE '0' END as valor_total_com_desconto FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO in (SELECT cod_formando FROM TB_PEDIDOS_FORMANDOS WHERE cod_pedido =  "&cod_pedido&") union all SELECT SUM(quantidade * (valor_un-(valor_un * (1-(ifnull(aliquota,0)/100))))) as valor_total_com_desconto  FROM TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO = "&cod_pedido&") as total_desconto")
							if not rsProdutosTotal.Eof then
								XTOTAL_DESCONTO = formatCurrency(0&rsProdutosTotal(0))/1
							end if
							rsProdutosTotal.Close
							set rsProdutosTotal = Nothing
					
 
					QTD_EM_PRODUTOS = formatCurrency(QTD_EM_PRODUTOS - XTOTAL_DESCONTO)
						%>
      <%=QTD_EM_PRODUTOS%> - <a href="lca_pedidos_pendentes_alteracoes_part_2.asp?cod_pedido=<%=COD_PEDIDO%>">Alterar Produtos</a></td>
					  </tr>
					<tr>
					  <td><div align="left"><strong>Total em Forma de Pagamento:</strong></div></td>
					  <td>
						<div align="left"><%=formatCurrency(VALOR_TOTAL)%> - <a href="lca_pedidos_pendentes_alteracoes_part_3.asp?cod_pedido=<%=COD_PEDIDO%>">Alterar Forma de Pagamento</a></div></td>
					  </tr>
					  
				  	<%End If%>
                  
					<tr>
					  <td colspan="2"><div align="center">  
					  
				<%
				temp = 0
				temp = formatNumber((QTD_EM_PRODUTOS-VALOR_TOTAL)/1,2)
				If temp < 0 Then
				%>
					<span class="msg">
						Aten&ccedil;&atilde;o o valor arrecadado &eacute; inferior do total em produtos, corriga antes de finalizar a corre&ccedil;&atilde;o<br />
						Diferen&ccedil;a faltante: <strong><%=formatNumber(QTD_EM_PRODUTOS-VALOR_TOTAL)%></strong>					</span>
				<%Elseif temp > 0  Then%>
						<span class="msg">
						Aten&ccedil;&atilde;o o valor total dos produtos dever&aacute; ser igual ao valor total na forma de pagamento, pois o pagamento dos consultores &eacute; feito com base nos produtos, por tanto corrija antes de continuar.<br />
						Diferen&ccedil;a: <strong><%=formatNumber(VALOR_TOTAL-QTD_EM_PRODUTOS)%></strong>					</span>
				<%Elseif Len(POSSUI_ERRO) > 3 Then%>
                	<span class="msg"><%=POSSUI_ERRO%></span>
                <%Else%>
					<input class="btnFinalizaCorrecao" type="submit" name="btnFinalizaCorrecao" value=" " />
				<%End If%>	  
				
	
					  </div></td>
					  </tr>
                    </table>
			</form>
</p>
		  </td>
  </tr>
</table>

 
  
  </td>
    </tr>
</table>

  
  </td>
  </tr>
</table>
	<%
FecharConexao()
%>
<!-- #include file="includes/rodape.asp" -->

  </td>
</tr>
</table>
</body>
</html>






