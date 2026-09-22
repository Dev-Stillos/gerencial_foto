<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>Baixar Pedidos para Diretoria</title>
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


function cadastraVendedor(valor){
	alert(valor);
}
function tirarSelecaodeTodasFichas(){
d = document.baixar_fichas
		<%
		AbrirConexao()
		sql = "SELECT COD_PEDIDO FROM TB_PEDIDOS WHERE DATA_DESCEU_DIRETORIA IS NULL AND FINALIZA_CORRECAO = 1"
		Set rsValidaCaoFichasJava = objConexao.Execute(sql)
		While not rsValidaCaoFichasJava.Eof
		%>
		d.PEDIDO_<%=rsValidaCaoFichasJava("COD_PEDIDO")%>.checked="";
		<%
		rsValidaCaoFichasJava.MoveNext()
		Wend
		
		rsValidaCaoFichasJava.Close()
		Set rsValidaCaoFichasJava=Nothing
		FecharConexao()
		%>
}
function selecionarTodasFichas(){
d = document.baixar_fichas
		<%
		AbrirConexao()
		sql = "SELECT COD_PEDIDO FROM TB_PEDIDOS WHERE DATA_DESCEU_DIRETORIA IS NULL AND FINALIZA_CORRECAO = 1"
		Set rsValidaCaoFichasJava = objConexao.Execute(sql)
		While not rsValidaCaoFichasJava.Eof
		%>
		d.PEDIDO_<%=rsValidaCaoFichasJava("COD_PEDIDO")%>.checked="checked";
		<%
		rsValidaCaoFichasJava.MoveNext()
		Wend
		
		rsValidaCaoFichasJava.Close()
		Set rsValidaCaoFichasJava=Nothing
		FecharConexao()
		%>
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

	  <h2>LCA - Pedidos Pendentes - Baixar Pedidos para Diretoria</h2>
      <div class="subMenu">
Sub-Menu: <a href="javascript:selecionarTodasFichas()">Selecionar todas as fichas</a> | <a href="javascript:tirarSelecaodeTodasFichas()">Tirar a seleção de todas as fichas</a></div>
	  <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	  
		  	<h3 align="center">Pedidos Pendentes Corrigidos </h3>
			<form name="baixar_fichas" method="post" action="lca_pedidos_pendentes_baixar_para_diretoria_processar.asp" style="margin:0px">
			<%
		  AbrirConexao()
		  Dim rs, rs2, rs3, rs4, sql
		  '1) Consulta todos os vendedores, por�m s� mostra na tela os que possuem pedidos pendentes
		  sql = "SELECT COD_VENDEDOR, NOME_VENDEDOR FROM TB_VENDEDORES ORDER BY NOME_VENDEDOR"
		   
		  
		  Set rs = objConexao.Execute(sql)
		  	  
		  	While not rs.Eof
		  
		 		 	'2) Verifica se h� registros pendentes do vendedor acima
			'		Filtros: Vendedor, N�o desceu para diretoria (Pendentes) e Pedidos Ativos	
			sql = "SELECT * FROM VW_PEDIDOS_COMPLETO WHERE COD_VENDEDOR="&rs("COD_VENDEDOR")&" AND DATA_DESCEU_DIRETORIA IS NULL AND FINALIZA_CORRECAO = 1 AND FINALIZA_VENDEDOR = 1 GROUP BY COD_PEDIDO ORDER BY NOME_FORMANDO"	
			
			
		  	
			Set rs2 = objConexao.Execute(sql)
			
			If not rs2.Eof Then%>
			
				<table style="margin-top:10px" cellpadding="5" cellspacing="0" align="center" width="100%">
			<tr>
			<td>
		      <div align="left"><span class="textoMedio"><strong><%=rs("NOME_VENDEDOR")%></strong></span></div></td>
			</tr>
			</table>
			<table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
            <td width="5%"><div align="center"><strong>Baixar</strong> </div></td>
            <td width="13%"><div align="center"><strong>Cód Pedido</strong> </div></td>
            <td width="35%"><div align="center"><strong>Cód. Formando(s) / Formando(s) </strong> </div></td>
            <td width="11%"><div align="center"><strong>Total Recebido </strong></div></td>
            <td width="10%"><div align="center"><strong>Lanç. do vend.  </strong></div></td>
            <td width="8%"><div align="center"><strong>Chegou em:</strong></div></td>
            <td width="18%"><div align="center"><strong>Vend. recebe em: </strong></div></td>
            </tr>
				<%
				Dim JAVA_CHECKED_TRUE, JAVA_CHECKED_FALSE
				JAVA_CHECKED_TRUE	= ""
				JAVA_CHECKED_FALSE	= ""
				
				While not rs2.Eof

				%>
					
					<tr>
            <td>
              <div align="center">
                <input class="inputRadio" value="1" type="checkbox" name="PEDIDO_<%=rs2("COD_PEDIDO")%>" id="checkbox" />
                </div>          </td>
            <td><div align="center"> <%=rs2("COD_PEDIDO")%> <%If int(rs2("PEDIDO_EXTRA")) = 1 Then%><em class="textoMenor">(Extra)</em><%End If%> </div></td>
            <td><div align="left">
			<%
			sql = 	"SELECT TB_PEDIDOS_FORMANDOS.COD_FORMANDO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO "&_
					"FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON "&_
					"TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO "&_
					"WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&rs2("COD_PEDIDO")&" ORDER BY NOME_FORMANDO"
			Set rs3 = objConexao.Execute(sql)
				
		
				While not rs3.Eof
				
				%>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				  <tr>
					<td width="23%"><div align="center">
					
					<%=digito(rs3("COD_FORMANDO"))%>
                    
                    </div></td>
					<td><div align="left"><%=rs3("NOME_FORMANDO")%></div></td>
				    </tr>
				</table>
				<%
				rs3.MoveNext()
				Wend
				
			rs3.Close
			Set rs3 = Nothing
			%>
			</div></td>
            <td><div align="right">
              <%
			Set rs4 = objConexao.Execute("SELECT * FROM VW_PEDIDOS_TOTAL_RECEBIDO WHERE COD_PEDIDO="&rs2("COD_PEDIDO"))
			Dim PAGAMENTO_EM_DINHEIRO, XTOTAL_CHEQUE, XTOTAL_CARTAO, XTOTAL_BOLETO, TOTAL_PAGO
			
			TOTAL_PAGO = 0
				
				If not rs4.Eof Then
				
					PAGAMENTO_EM_DINHEIRO	= formatNumber(0&rs4("PAGAMENTO_EM_DINHEIRO"))/1
					XTOTAL_CHEQUE			= formatNumber(0&rs4("XTOTAL_CHEQUE"))/1
					XTOTAL_CARTAO			= formatNumber(0&rs4("XTOTAL_CARTAO"))/1
					XTOTAL_BOLETO			= formatNumber(0&rs4("XTOTAL_BOLETO"))/1
					TOTAL_PAGO				= (PAGAMENTO_EM_DINHEIRO + XTOTAL_CHEQUE + XTOTAL_CARTAO + XTOTAL_BOLETO)/1
					
					Response.Write(formatNumber(TOTAL_PAGO))
					
				End If
			
			rs4.Close()
			Set rs4 = Nothing
			%>
            </div></td>
            <td><div align="center"><%=ArrumaData(rs2("DATA_CADASTRO"))%></div></td>
            <td><div align="center"><%=ArrumaData(rs2("DATA_CHEGADA"))%></div></td>
            <td>
			<div align="center">
			<%If rs2("VENDEDOR_RECEBE_EM") <> "" Then%>
				<%=MonthName(Month(rs2("VENDEDOR_RECEBE_EM")))&" / "&Year(rs2("VENDEDOR_RECEBE_EM"))%>
			<%Else%>
			-
			<%End iF%>
			</div>			</td>
            </tr>
					
						
				<%
				
				
				JAVA_CHECKED_TRUE	= JAVA_CHECKED_TRUE&" d.PEDIDO_"&rs2("COD_PEDIDO")&".checked=""checked""; "
				JAVA_CHECKED_FALSE	= JAVA_CHECKED_FALSE&" d.PEDIDO_"&rs2("COD_PEDIDO")&".checked=""""; "
				
				
				rs2.MoveNext()
				Wend
				%>
					<script language="javascript">
						function selecionartodasVendedor_<%=rs("COD_VENDEDOR")%>(){
							d = document.baixar_fichas
							<%=JAVA_CHECKED_TRUE%>
						}
						
						function tirarselecaodeTodasVendedor_<%=rs("COD_VENDEDOR")%>(){
							d = document.baixar_fichas
							<%=JAVA_CHECKED_FALSE%>
						}
					</script>
                <tr>
					  <td colspan="7"></span><a href="javascript:selecionartodasVendedor_<%=rs("COD_VENDEDOR")%>()">Selecionar todas deste vendedor</a><span class="msg2">|</span><a href="javascript:tirarselecaodeTodasVendedor_<%=rs("COD_VENDEDOR")%>()">Tirar seleção de todas deste vendedor</a></td>
					  </tr>
			</table>
	
			    <%End If
			
			rs2.Close
			Set rs2 = Nothing
			
		
		 	rs.MoveNext()
		  	Wend
	      
		  rs.Close()
		  Set rs = Nothing
		  	  
		  FecharConexao()
		  %>
          		<div align="center">
			  <p>
			    </p>
			  <p>
			    <input class="btnGravar" type="submit" name="btnTopo2" value=" " />
			    </p>
			  <p></p>
			</div>
			</form>
        
		
		</td>

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



