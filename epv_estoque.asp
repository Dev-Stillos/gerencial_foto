<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!--
'*** Instru��es
'1) Relaciona todos os vendedores
'2) Mostra somentes os vendedores que possuem pedidos pendentes


-->

<%
'ROTINAS VERIFICAR CODE SMELL
'Server.Execute("adm_rotina_corrige_estoque_movimentacoes_prazo.asp")%>
<%'Server.Execute("adm_rotina_corrige_estoque_variados.asp")%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>Devolu&ccedil;&atilde;o e Sa&iacute;da</title>
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

<body onKeyDown="javascript:Verificar()">

<table cellpadding="9px" cellspacing="0" width="100%">
<tr>
<td>
<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td>
		<!-- #include file="includes/topo.asp" -->
<table width="100%" cellpadding="0" cellspacing="0" style="border-bottom: 1px dotted #999999; background-color: white;">
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
  
  <!-- Estilo compatível com o layout da imagem -->
<div style="display: flex; justify-content: center; margin-bottom: 20px;">
  <div style="background-color: #ff007f; padding: 10px 20px; border-radius: 8px; display: flex; align-items: center; gap: 10px; position: relative; z-index: 9;">

   <h2 style="color: white !important; margin: 0; font-size: 18px; background: #ff007f; z-index: 10; position: relative;">Devolu&ccedil;&atilde;o e Sa&iacute;da</h2>

    <button onclick="toggleDropdownEPV()" style="background: none; color: white; border: none; font-size: 14px; cursor: pointer;">
      | Sub-Menu <span style="font-size: 12px;">&#9660;</span>
    </button>

    <!-- SUBMENU DROPDOWN -->
    <div id="submenuDropdownEPV" style="display:none; position:absolute; top:100%; left:0; width:100%; background-color:#fff; border-radius:8px; box-shadow:0px 8px 16px rgba(0,0,0,0.2); z-index:1000; margin-top:5px;">
      <a href="epv_devolucao.asp?comando=nova" style="display:block; padding:10px; color:#ff007f; text-decoration:none;">Iniciar devolu&ccedil;&atilde;o de Vendedor</a>
      <a href="epv_saida_inicio.asp?comando=nova" style="display:block; padding:10px; color:#ff007f; text-decoration:none;">Sa&iacute;da de &Aacute;lbuns do Estoque</a>
      <a href="epv_saida_fornecedores.asp" style="display:block; padding:10px; color:#ff007f; text-decoration:none;">Sa&iacute;da do Estoque para Fornecedor</a>
      <a href="epv_saida_produtos.asp" style="display:block; padding:10px; color:#ff007f; text-decoration:none;">Sa&iacute;da de itens do Estoque</a>
      <a href="epv_devolucao_produtos.asp" style="display:block; padding:10px; color:#ff007f; text-decoration:none;">Devolu&ccedil;&atilde;o de itens para o Estoque</a>
    </div>
  </div>
</div>

<script>
  function toggleDropdownEPV() {
    var el = document.getElementById("submenuDropdownEPV");
    el.style.display = el.style.display === "none" || el.style.display === "" ? "block" : "none";
  }

  window.onclick = function(event) {
    if (!event.target.closest('.subMenuEPV') && !event.target.matches('button')) {
      var dropdown = document.getElementById("submenuDropdownEPV");
      if (dropdown && dropdown.style.display === "block") {
        dropdown.style.display = "none";
      }
    }
  }
</script>
  	  
	<%
		AbrirConexao()

		Dim data_inicial
		data_inicial = dateAdd("d",-45,Now())
		data_inicial = Year(data_inicial)&"-"&Month(data_inicial)&"-"&Day(data_inicial)
					
		sql = "SELECT " & _
      " COD_VENDEDOR, " & _
      " NOME_VENDEDOR, " & _
      " COUNT(*) AS ITENS_EM_POSSE, " & _
      " SUM(CASE WHEN TIPO = 0 AND NOW() > DATA_PREVISTA_RETORNO THEN 1 ELSE 0 END) AS ALBUM_FISICO, " & _
      " SUM(CASE WHEN TIPO = 1 AND NOW() > DATA_PREVISTA_RETORNO THEN 1 ELSE 0 END) AS FICHA, " & _
      " SUM(CASE WHEN TIPO = 2 THEN 1 ELSE 0 END) AS ALBUM_DIGITAL, " & _
      " ( " & _
      "   SUM(CASE WHEN TIPO = 0 AND NOW() > DATA_PREVISTA_RETORNO THEN 1 ELSE 0 END) + " & _
      "   SUM(CASE WHEN TIPO = 1 AND NOW() > DATA_PREVISTA_RETORNO THEN 1 ELSE 0 END) " & _
      " ) AS ITENS_PENDENTES " & _
      "FROM VW_MOVIMENTACOES_COMPLETA " & _
      "WHERE DATA_RETORNO IS NULL " & _   
      "AND TIPO IN (0,1,2) " & _
      "AND COD_VENDEDOR IN (SELECT COD_VENDEDOR FROM TB_VENDEDORES WHERE ATIVADO = 1) " & _
      "GROUP BY COD_VENDEDOR, NOME_VENDEDOR " & _
      "ORDER BY NOME_VENDEDOR"





		'Forma antiga que era 45 dias
		'sql = "SELECT COD_VENDEDOR, NOME_VENDEDOR, (SUM( (DATEDIFF( NOW( ) , DATA_SAIDA ) ) -45) / COUNT( * ) ) AS MEDIA_DIAS_FORA, COUNT( * ) AS ITENS_PENDENTES FROM VW_MOVIMENTACOES_COMPLETA WHERE DATA_RETORNO IS NULL AND TIPO <> 1 AND DATA_SAIDA <= '"&data_inicial&" 23:59:59' AND COD_VENDEDOR <> 27 GROUP BY COD_VENDEDOR ORDER BY NOME_VENDEDOR"

		'Elseif CHEQ_DEVOLVIDOS = 3 Then
		'Devolvidos e N�o devolvidos
			'N�o faz nada!!!

		Set rs = objConexao.Execute(sql)
		
		If not rs.Eof Then
		%>

	    <h3 align="center">Relat&oacute;rio de Representantes Comerciais com &Aacute;lbuns e demais itens em posse pendentes e atrasados.</h3>
	      <div class="msgGrande"><strong>Para consultar todos os itens em posse do Represente mesmo os itens que n&atilde;o est&atilde;o em atraso, utilize o menu: &quot;Relat&oacute;rios&quot; &gt; &quot;Relat&oacute;rio de Itens em posse do Vendedor&quot;.
          </strong></div>
	  
	  <table width="100%" 
		      		border="1" align="center"
					cellpadding="5" 
					style="border-collapse:collapse">
				<tr bgcolor="whitesmoke">
                <td width="15%" height="28"><div align="center"><span class="MensagemFicha"><strong>Representante Comercial </strong></span><strong><span class="MensagemFicha"></span></strong></div></td>
                <td colspan="3"><div align="center"><strong>Imprimir</strong></div></td>
                <td width="07%"><div align="center" class="textoMedio"><strong>Itens em Posse</strong></div></td>
				<td width="07%"><div align="center" class="textoMedio"><strong>Álbum Físico</strong></div></td>
				<td width="07%"><div align="center" class="textoMedio"><strong>Fichas</strong></div></td>
				<td width="07%"><div align="center" class="textoMedio"><strong>Álbum Digital</strong></div></td>
				<td width="07%"><div align="center" class="textoMedio"><strong>Itens Pendentes</strong></div></td>

                </tr>
		<%
		Dim QTD_GERAL, ATRASADOS
		QTD_GERAL = 0
		While not rs.Eof

			ATRASADOS = 0
			Dim EM_POSSE, PENDENTES
			EM_POSSE = formatNumber(0 & rs("ITENS_EM_POSSE"), 0)
			ALBUM_FISICO = formatNumber(0 & rs("ALBUM_FISICO"), 0)
			FICHA = formatNumber(0 & rs("FICHA"), 0)
			PENDENTES = formatNumber(0 & rs("ITENS_PENDENTES"), 0)
			QTD_GERAL = QTD_GERAL + PENDENTES


			'Mostrar todos, mesmo que tenha zero

			'Mostra somente se tiver pelo menos 1 item atrasado
			%>     
				<tr>
					<td><%=rs("NOME_VENDEDOR")%></td>
					<td width="8%"><div align="center"><a href="adm_relatorios_extrato_vendedores_atrasados_por_vendedor_html.asp?COD_VENDEDOR_BLOQUEADO=<%=rs("COD_VENDEDOR")%>&COM_COPIA=2" target="_blank"> 1° Via de Álbum Físico</a></div></td>
					<td width="7%"><div align="center"><a href="adm_relatorios_extrato_vendedores_atrasados_por_vendedor_html2.asp?COD_VENDEDOR_BLOQUEADO=<%=rs("COD_VENDEDOR")%>&COM_COPIA=1" target="_blank">1° Via de Ficha</a></div></td>
					<td width="7%"><div align="center"><a href="adm_relatorios_extrato_vendedores_atrasados_por_vendedor_html3_DIGITAL.asp?COD_VENDEDOR_BLOQUEADO=<%=rs("COD_VENDEDOR")%>&COM_COPIA=1" target="_blank">1° Via Álbum Digital</a></div></td>
					<td bgcolor="#DDFEFF"><div align="center" class="textoMedio"><strong><%=EM_POSSE%></strong></div></td>
					<td bgcolor="#DDFEFF"><div align="center" class="textoMedio"><strong><%=ALBUM_FISICO%></strong></div></td>
					<td bgcolor="#DDFEFF"><div align="center" class="textoMedio"><strong><%=FICHA%></strong></div></td>
					<td bgcolor="#ff007f"><div align="center" class="textoMedio" style="color:white;"><strong><%=formatNumber(0 & rs("ALBUM_DIGITAL"), 0)%></strong></div></td>
					<td bgcolor="#FDDBDB"><div align="center" class="textoMedio"><strong><%=PENDENTES%></strong></div></td>
				</tr>


			<%
			
			
			rs.MoveNext()
			Wend
			
			rs.Close
			Set rs = Nothing
			
			FecharConexao()
			%>		 
        
		  	<tr bgcolor="whitesmoke">
		  				    <td colspan="8"><div align="center"><strong>Quantidade Total de Itens Pendentes: <%=QTD_GERAL%></strong></div></td>
                        <tr>
                          <td colspan="8"><div align="center"><strong>Observa&ccedil;&otilde;es:</strong></div><br />
								Itens n&atilde;o devolvidos, mas j&aacute; lan&ccedil;ados pelo sistema pelo representante comercial<strong> constam</strong> como atrasados mesmo que n&atilde;o foram corrigidos pelo departamento fotogr&aacute;fico.<br />
								Caso o item for vendido sua data retorno ser&aacute; a mesma data do lan&ccedil;amento do sistema pelo representante comercial, entretanto esta data  &eacute; lan&ccedil;ada no final da corre&ccedil;&atilde;o efetuada pelo depto. de foto e reportagens.</td>
						</table>
		  <p>
           </p>
		    <%Else%>
		   <div align="center">
           <br /><br /><br /><br />
		  <p><span class="textoMedio"><strong>N&atilde;o h&aacute; dados para os filtros selecionados!</strong></span></p>
          <br /><br /><br /><br />
		</div>
		    <%End If%>
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