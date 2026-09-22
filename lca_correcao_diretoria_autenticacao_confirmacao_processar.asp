<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<%
Dim cod_fechamento_periodo, COD_VENDEDOR, NOME_VENDEDOR
Dim rs, rs3, rsNomeVend, sql, sqlBase

Dim PRODUTOS, DINHEIRO, CARTAO, CHEQUE
Dim SUB_TOTAL_EM_PRODUTOS, SUB_TOTAL_EM_DINHEIRO, SUB_TOTAL_EM_CARTAO, SUB_TOTAL_EM_CHEQUE
Dim TOTAL_EM_PRODUTOS, TOTAL_EM_DINHEIRO, TOTAL_EM_CARTAO, TOTAL_EM_CHEQUE

TOTAL_EM_PRODUTOS	= 0
TOTAL_EM_DINHEIRO	= 0
TOTAL_EM_CARTAO		= 0
TOTAL_EM_CHEQUE		= 0

COD_VENDEDOR	= Request.QueryString("COD_VENDEDOR")
cod_fechamento_periodo	= Request.QueryString("cod_fechamento_periodo")


AbrirConexao()

	Set rsNomeVend = objConexao.Execute("SELECT NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR="&COD_VENDEDOR)
	
		If not rsNomeVend.Eof Then
			NOME_VENDEDOR = rsNomeVend("NOME_VENDEDOR")
		End If
	
	rsNomeVend.Close
	Set rsNomeVend = Nothing
	
FecharConexao()

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
	<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>
	
	<script type="text/javascript" language="javascript">
<!--
<!--AJUDA F1 E CTRL PARA BUSCA RÁPIDA-->
		
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

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->

function cadastraVendedor(valor){
	alert(valor);
}
</script>

</head>

<body>
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

	  <h2>LCA - Corre&ccedil;&atilde;o Diretoria / Confirma&ccedil;&atilde;o de Baixa </h2>
	   <div class="subMenu">Sub-Menu: <a href="filtro_baixas_pag_erros.asp">Verificar fichas marcadas com erros</a> | <a href="relatorios_fichas_baixadas.asp?filtro1=+">&Uacute;ltimas fichas baixadas</a> </div>
	  
		<h3 align="center">Aguarde processando... </h3>
		<form name="vendas" method="post" action="lca_correcao_diretoria_autenticacao_confirmacao_processar.asp?COD_VENDEDOR=<%=COD_VENDEDOR%>&MES_VENDA=<%=MES_VENDA%>&ANO_VENDA=<%=ANO_VENDA%>">
          <%
		 Dim contaRegistrosConfirma
		 contaRegistrosConfirma = 0 
	sqlBase =	" SELECT TB_PEDIDOS.COD_PEDIDO, TB_PEDIDOS.DATA_VENDA, TB_PEDIDOS.ERRO_LISTA, TB_PEDIDOS_FORMANDOS.COD_FORMANDO, "&_
	  			" TB_PEDIDOS.FINALIZA_DIRETORIA, TB_PEDIDOS.DATA_CORRIGIDO_DIRETORIA, "&_
				" TB_FICHAS_IDENTIFICACAO.COD_CONTRATO "&_
				" FROM TB_PEDIDOS "&_
				" LEFT JOIN TB_PEDIDOS_FORMANDOS ON TB_PEDIDOS.COD_PEDIDO = TB_PEDIDOS_FORMANDOS.COD_PEDIDO "&_
				" LEFT JOIN TB_FICHAS_IDENTIFICACAO ON TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO WHERE "&_
				" TB_PEDIDOS.COD_VENDEDOR="&COD_VENDEDOR&" AND TB_PEDIDOS.cod_fechamento_periodo = '"&cod_fechamento_periodo&"' AND  TB_PEDIDOS.DATA_DESCEU_DIRETORIA IS NOT NULL"
		
		  AbrirConexao()
		  sql = sqlBase & " GROUP BY TB_PEDIDOS.COD_PEDIDO ORDER BY COD_CONTRATO"
		  'Response.Write(sql)
		  
		  Set rs = objConexao.Execute(sql)
		 %>
					
					
					<table width="70%" 
		      border="0" align="center"
					cellpadding="5" bgcolor="#FFFFFF" 
					style="border-collapse:collapse">
		
			<%
			While not rs.Eof

			If Request.Form("COMBO_"&rs("COD_PEDIDO")) = 1 Then

				'1) Verifica realmente se não estão baixadas
				'2) Se não estiverem baixa a ficha
					
				If rs("FINALIZA_DIRETORIA") = 2 Then

					contaRegistrosConfirma = contaRegistrosConfirma + 1
					objConexao.Execute("UPDATE TB_PEDIDOS SET FINALIZA_DIRETORIA = '1', DATA_CORRIGIDO_DIRETORIA = NOW( ) WHERE COD_PEDIDO = "&rs("COD_PEDIDO")&" LIMIT 1")
				
					'Tenta gerar OS
					'/////////////////////////////////
					'Deleta anteriores do mesmo produto que estejam no passo 1
					'objConexao.Execute("DELETE FROM TB_ORDENS_SERVICO_ETAPAS WHERE cod_ordem_servico IN(SELECT cod_ordem_servico FROM TB_ORDENS_SERVICO WHERE cod_pedido = "&rs("COD_PEDIDO")&")")		
					'objConexao.Execute("DELETE FROM TB_ORDENS_SERVICO_HISTORICO WHERE cod_ordem_servico IN(SELECT cod_ordem_servico FROM VW_TB_ORDENS_SERVICO WHERE cod_pedido = "&rs("COD_PEDIDO")&")")
					'objConexao.Execute("DELETE FROM TB_ORDENS_SERVICO WHERE status_ordem_servico = 1 AND cod_pedido = "&rs("COD_PEDIDO"))
					'objConexao.Execute("DELETE FROM TB_ORDENS_SERVICO_TEMPORARIA WHERE cod_ordem_servico IN(SELECT cod_ordem_servico FROM VW_TB_ORDENS_SERVICO WHERE cod_pedido = "&rs("COD_PEDIDO")&")")
					'////////////////////////////////
					
					'REMOVIDO POR JOAO, AGORA A ORDEM SERVICO SERA GERADA AO FINALIZAR A CORREÇÂO DA LIGIA.
					'retorna = moduloPPVgerarOrdensServico(rs("COD_PEDIDO"),true)
				End If
			End If

			rs.MoveNext()
		    Wend

			rs.Close
			Set rs=Nothing%>
		    <div align="center">
		<%
		FecharConexao()
		Response.Redirect("lca_correcao_diretoria_mensal.asp?cod_fechamento_periodo="&cod_fechamento_periodo&"&COD_VENDEDOR="&COD_VENDEDOR&"&msg="&contaRegistrosConfirma&" pedido(s) corrigido(s) com sucesso!")
		%> 
        <br />
        &nbsp;&nbsp;&nbsp;<br /><br />
		<br />
        </div>
		<br />
		<br />   
</form>
<p align="center">&nbsp;</p>		</td>
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