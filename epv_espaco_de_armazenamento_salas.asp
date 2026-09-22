<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim QTD_ALBUNS_TOTAL, QTD_ALBUNS_SALA
QTD_ALBUNS_TOTAL = 0
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
<!--
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
 


	<!--AJUDA F1 E CTRL PARA BUSCA R�PIDA-->

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->

function cadastraVendedor(valor){
	alert(valor);
}
</script>
<!--INICIO VALIDA��O SIM/N�O-->
<script language="javascript">
function confirmaExclusao(COD_SALA_ARMAZENADO,NOME_SALA_ARMAZENADO){	
	var mensagem = "Tem certeza que deseja excluir est� sala?\n"+NOME_SALA_ARMAZENADO
	if(YNconfirm(mensagem)){
		var urlDireciona = "epv_espaco_de_armazenamento_salas_sala.asp?comando=excluir&COD_SALA_ARMAZENADO="+COD_SALA_ARMAZENADO;
		document.location.href=urlDireciona;
	}
}
	
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
<!--FIM VALIDA��O SIM/N�O-->
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

	  <h2>EPV - Estoque</h2>
	  <div class="subMenu">
	  Sub-Menu: <a href="epv_estocagem_produtos.asp">Estocagem de produtos</a> | <a href="epv_espaco_de_armazenamento_salas_sala.asp?comando=incluir">Incluir Sala de Armazenagem</a>
    
     | <a href="epv_inventario.asp">Invent&aacute;rio</a>
    </div>
	  <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	 <h3 align="center">Salas de Armazenagem</h3>
		 <%
		 Dim rs, rs2, rs3, sql
		 AbrirConexao()
		 Set rs = objConexao.Execute("SELECT ANDAR FROM TB_LOCAL_ARMAZENAGEM_SALA GROUP BY ANDAR")
		 	While not rs.Eof
		 %>
		    <table width="100%" border="0" align="center" 	cellpadding="5">
                <tr>
                  <td><p class="textoMedio"><strong>
				  <%
				  If rs("ANDAR") = 1 Then
				  	Response.Write("(1) - Sub-Solo")
				  Elseif rs("ANDAR") = 2 Then
				  	Response.Write("(2) - T&eacute;rreo")
				  Elseif rs("ANDAR") = 3 Then
				  	Response.Write("(3) - 1&ordm; Andar")
				  End If
				  %>
                  </strong></p></td>
                  </tr>
              </table>
		    <table width="100%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
                <tr>
                  <td width="27%"><div align="center"><strong>Sala</strong></div></td>
                  <td><div align="center"><strong>Qtd. de Caixas</strong></div></td>
                  <td><div align="center"><strong>Caixas</strong></div></td>
                  <td width="8%"><div align="center"><strong>Alterar / Excluir</strong></div></td>
                </tr>
                <%
				Set rs2 = objConexao.Execute("SELECT * FROM TB_LOCAL_ARMAZENAGEM_SALA WHERE ANDAR='"&rs("ANDAR")&"' ORDER BY NOME_SALA_ARMAZENADO")
					While not rs2.Eof
				%>
                <tr>
                  <td><div align="left"><%=rs2("COD_SALA_ARMAZENADO")%> - <%=rs2("NOME_SALA_ARMAZENADO")%></div></td>
                  <td> <div align="center">
                    <%
				  Set rs3 = objConexao.Execute("SELECT COUNT(COD_CAIXA_ARMAZENADO) as QTD_CAIXAS FROM TB_LOCAL_ARMAZENAGEM_SALA_CAIXA WHERE COD_SALA_ARMAZENADO = "&rs2("COD_SALA_ARMAZENADO"))

						If not rs3.Eof Then
						
						Dim QTD_CAIXAS
						QTD_CAIXAS = formatNumber(0&rs3("QTD_CAIXAS"),0)
							If QTD_CAIXAS > 0 Then
								Response.Write(rs3("QTD_CAIXAS"))
							Else
								Response.Write("-")
							End If
						End If
						
				  rs3.Close
				  Set rs3 = Nothing
				  %>
                  </div></td>
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr>
    <td><div align="center"><a href="epv_espaco_de_armazenamento_salas_caixas.asp?COD_SALA_ARMAZENADO=<%=rs2("COD_SALA_ARMAZENADO")%>">Listar Caixas<br />
    </a></div></td>
  </tr>
  <tr>
    <td><div align="center"><a href="epv_espaco_de_armazenamento_salas_caixas_caixa.asp?COD_SALA_ARMAZENADO=<%=rs2("COD_SALA_ARMAZENADO")%>&comando=incluir">Incluir Caixa</a></div></td>
  </tr>
</table></td>
                  <td><div align="center"><a href="epv_espaco_de_armazenamento_salas_sala.asp?comando=alterar&COD_SALA_ARMAZENADO=<%=rs2("COD_SALA_ARMAZENADO")%>"><img src="images/icons/icon_alterar.png" width="16" height="16" border="0" /></a>&nbsp;&nbsp;<a href="javascript:confirmaExclusao(<%=rs2("COD_SALA_ARMAZENADO")%>,'<%=rs2("NOME_SALA_ARMAZENADO")%>')"><img src="images/icons/icon_excluir.png" width="16" height="16" border="0" /></a></div></td>
                </tr>
                <%
					rs2.MoveNext()
					Wend
				%>
              </table>
		   		<%
				rs2.Close
				Set rs2 = Nothing
				%>
         <%
		 	rs.MoveNext()
			Wend

		 rs.Close
		 Set rs = Nothing
		 %> 
         <br />
<br /> 
         </td>

    </tr>
</table>
<br />
<br />


 
  
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