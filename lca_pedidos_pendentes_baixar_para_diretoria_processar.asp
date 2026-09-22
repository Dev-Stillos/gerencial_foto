<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/classe_fechamento.asp" -->
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
Sub-Menu: <a href="javascript:selecionarTodasFichas()">Selecionar todas as fichas</a> | <a href="javascript:tirarSelecaodeTodasFichas()">Tirar a sele&ccedil;&atilde;o de todas as fichas</a></div>
	  <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	  
		  	<h3 align="center">Pedidos Pendentes Corrigidos </h3>
			<form name="baixar_fichas" method="post" action="lca_pedidos_pendentes_baixar_para_diretoria_processar.asp" style="margin:0px">
			<%
		  AbrirConexao()
		  Dim rs, rs2, rs3, rs4, sql
		  '1) Consulta todos os vendedores, porém só mostra na tela os que possuem pedidos pendentes
		  sql = "SELECT COD_VENDEDOR, NOME_VENDEDOR FROM TB_VENDEDORES ORDER BY NOME_VENDEDOR"
		   
		  
		  Set rs = objConexao.Execute(sql)
		  	  
		  	While not rs.Eof
		  
		 		 	'2) Verifica se há registros pendentes do vendedor acima
			'		Filtros: Vendedor, Não desceu para diretoria (Pendentes) e Pedidos Ativos
			sql = "SELECT * FROM TB_PEDIDOS WHERE COD_VENDEDOR="&rs("COD_VENDEDOR")&" AND DATA_DESCEU_DIRETORIA IS NULL AND FINALIZA_CORRECAO = 1"
		  	
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
            <td width="13%"><div align="center"><strong>C&oacute;d Pedido</strong> </div></td>
            <td width="35%"><div align="center"><strong>C&oacute;d. Formando(s) / Formando(s) </strong> </div></td>
            <td width="13%"><div align="center"><strong>Total Recebido </strong></div></td>
            <td width="16%"><div align="center"><strong>Lan&ccedil;. do vend.  </strong></div></td>
            <td width="18%"><div align="center"><strong>Vend. recebe em: </strong></div></td>
            </tr>
				<%				
				While not rs2.Eof

				%>
					
					<tr>
            <td>
              <div align="center">
                <input class="inputRadio" value="1" type="checkbox" name="PEDIDO_<%=rs2("COD_PEDIDO")%>" id="checkbox" />
                <%
				Dim COD_PEDIDO, PEDIDO_VALIDADO
				COD_PEDIDO 		= rs2("COD_PEDIDO")
				PEDIDO_VALIDADO	= formatNumber(0&Request.Form("PEDIDO_"&COD_PEDIDO),0)
				
				If PEDIDO_VALIDADO = 1 Then
				'Baixar o pedido
					objConexao.Execute("UPDATE TB_PEDIDOS SET DATA_DESCEU_DIRETORIA = NOW( ) WHERE COD_PEDIDO ="&COD_PEDIDO&" LIMIT 1")
					
					
				End If
				%>
                
                </div>          </td>
            <td><div align="center"> <%=rs2("COD_PEDIDO")%> </div></td>
            <td colspan="4">&nbsp;</td>
            </tr>
					
						
				<%			
				
				rs2.MoveNext()
				Wend
				%>
					
			</table>
	
			    <%End If
			
			rs2.Close
			Set rs2 = Nothing
			
		
		 	rs.MoveNext()
		  	Wend
	      
		  rs.Close()
		  Set rs = Nothing
		  	  
		  FecharConexao()
		  Response.Redirect("lca_pedidos_pendentes.asp?msg=Pedidos baixados com sucesso!")
		  %>
          		<div align="center">
			  <p>
			    </p>
			  <p>
			    <input class="btnGravar" type="submit" name="btnTopo2" value=" " />
			    </p>
			  <p>&nbsp;</p>
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