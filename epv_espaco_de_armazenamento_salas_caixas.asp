<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim COD_SALA_ARMAZENADO, NOME_SALA_ARMAZENADO, ANDAR

COD_SALA_ARMAZENADO = Request.QueryString("COD_SALA_ARMAZENADO")

AbrirConexao()
Set rs = objConexao.Execute("SELECT * FROM TB_LOCAL_ARMAZENAGEM_SALA WHERE COD_SALA_ARMAZENADO="&COD_SALA_ARMAZENADO)
	
		If not rs.Eof Then
			NOME_SALA_ARMAZENADO=rs("NOME_SALA_ARMAZENADO")
			ANDAR				=rs("ANDAR")
		End If
	
rs.Close
Set rs = Nothing


 Dim rs, sql, sqlWhere, sqlOrderBy
 Dim NUMERO_CAIXA, COLUNA
 sql = "SELECT * FROM TB_LOCAL_ARMAZENAGEM_SALA_CAIXA WHERE COD_SALA_ARMAZENADO="&COD_SALA_ARMAZENADO
 sqlOrderBy =  " ORDER BY COLUNA, NUMERO_CAIXA"
 
 If Request.QueryString("filtro1") <> Empty Then
 	
	NUMERO_CAIXA = Cstr(""&Request.QueryString("NUMERO_CAIXA"))
	COLUNA = Cstr(""&Request.QueryString("COLUNA"))
	
	If NUMERO_CAIXA <> "" Then
		sqlWhere = " AND NUMERO_CAIXA = "&sqlTexto(NUMERO_CAIXA)
		COLUNA = "" 'Limpa campos não chave
	Else
		'Verifica coluna somente se não tiver preenchido campo chave NUMERO_CAIXA
		If COLUNA <> "" Then
			sqlWhere = " AND COLUNA = "&sqlTexto(COLUNA)
		End If	
	End If
	
 End If
 
'Junta tudo!!!
sql = sql&sqlWhere&sqlOrderBy
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
<!--INICIO VALIDAÇÃO SIM/NÃO-->
<script language="javascript"> 
function confirmaExclusao(NUMERO_CAIXA,COD_CAIXA_ARMAZENADO){	
	var mensagem = "Tem certeza que deseja excluir está caixa?\n"+NUMERO_CAIXA
	if(YNconfirm(mensagem)){
		var urlDireciona = "epv_espaco_de_armazenamento_salas_caixas_caixa.asp?comando=excluir&COD_CAIXA_ARMAZENADO="+COD_CAIXA_ARMAZENADO+"&COD_SALA_ARMAZENADO=<%=COD_SALA_ARMAZENADO%>";
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
<!--FIM VALIDAÇÃO SIM/NÃO-->
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

	  <h2>EPV - Estoque / Salas de Armazenagem</h2>
	  <div class="subMenu">
	  Sub-Menu:
      <a href="epv_espaco_de_armazenamento_salas.asp">Voltar para salas</a> | 
      <a href="epv_espaco_de_armazenamento_salas_caixas_caixa.asp?COD_SALA_ARMAZENADO=<%=COD_SALA_ARMAZENADO%>&comando=incluir">Incluir Caixa de Armazenagem nesta sala</a></div>
      <form action="?" style="margin:0px">
	  <table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="99%" border="0" align="center" cellpadding="5" cellspacing="0">
            <!--
            <tr>
		    <td><div align="right"><strong>N&uacute;mero Ordem de Servi&ccedil;o</strong></div></td>
		    <td><div align="left"><input type="text" value="<%=busca_cod_ordem_servico%>" name="cod_ordem_servico" class="inputTamanho2" /></div></td>
		    </tr>
            -->
            <tr>
		    <td width="17%"><div align="right"><strong>N&uacute;mero da Caixa</strong></div></td>
		    <td><table width="540" border="0" cellspacing="0" cellpadding="0">
		      <tr>
		        <td width="117"><div align="left">
		          <input type="text" value="<%=NUMERO_CAIXA%>" name="NUMERO_CAIXA" class="inputTamanho1" />
		          </div></td>
		        <td width="49"><div align="right"><strong>Coluna</strong></div></td>
		        <td width="374" style="padding-left:5px">
                <select name="COLUNA" id="select">
                  <option <%If COLUNA = "" Then%> selected="selected" <%End If%> value="">-</option>
                  <option <%If COLUNA = "A" Then%> selected="selected" <%End If%> value="A">A</option>
                  <option <%If COLUNA = "B" Then%> selected="selected" <%End If%> value="B">B</option>
                  <option <%If COLUNA = "C" Then%> selected="selected" <%End If%> value="C">C</option>
                  <option <%If COLUNA = "D" Then%> selected="selected" <%End If%> value="D">D</option>
                  <option <%If COLUNA = "E" Then%> selected="selected" <%End If%> value="E">E</option>
                  <option <%If COLUNA = "F" Then%> selected="selected" <%End If%> value="F">F</option>
                  <option <%If COLUNA = "G" Then%> selected="selected" <%End If%> value="G">G</option>
                  <option <%If COLUNA = "H" Then%> selected="selected" <%End If%> value="H">H</option>
	           </select></td>
		        </tr>
		      </table></td>
		    </tr>
		  <tr>
		    <td> </td>
		    <td><div align="left">
              <input type="hidden" name="COD_SALA_ARMAZENADO" value="<%=COD_SALA_ARMAZENADO%>" />
		      <input class="btnFiltrarAmarelo" type="submit" name="filtro1" value=" " />
		      </div></td>
		    </tr>
        </table>
      </form>
        <br />
      
	  <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
		 <%
		 AbrirConexao()
		 Set rs = objConexao.Execute(sql)
		 
		 
		 If not rs.Eof Then%>
          <h3 align="center">Caixas de Armazenagem da Sala &quot;<%=NOME_SALA_ARMAZENADO%>&quot;</h3>
          
		 <table width="70%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
                <tr>
                  <td width="20%"><div align="center"><strong>N&uacute;mero da Caixa</strong></div></td>
                  <td width="55%"><div align="center"><strong>Coluna</strong></div></td>
                  <td width="14%"><div align="center"><strong>Listar Itens Armazenados</strong></div></td>
                  <td width="11%"><div align="center"><strong>Alterar / Excluir</strong></div></td>
                </tr>
         <%While not rs.Eof
		 
		 COD_CAIXA_ARMAZENADO	= rs("COD_CAIXA_ARMAZENADO")
		 COLUNA					= rs("COLUNA")
		 
		 %>       
                <tr>
                  <td><div align="center"><%=rs("NUMERO_CAIXA")%></div></td>
                  <td><div align="center"><%=COLUNA%></div>                  </td>
                  <td><div align="center"> <a href="epv_espaco_de_armazenamento_salas_caixas_imprimir.asp?cod_caixa_armazenado=<%=COD_CAIXA_ARMAZENADO%>" target="_blank"><img src="images/icons/lista.png" alt="" width="26" height="29" border="0" /></a> </div></td>
                  <td><div align="center"><a href="epv_espaco_de_armazenamento_salas_caixas_caixa.asp?comando=alterar&COD_SALA_ARMAZENADO=<%=COD_SALA_ARMAZENADO%>&COD_CAIXA_ARMAZENADO=<%=COD_CAIXA_ARMAZENADO%>"><img src="images/icons/icon_alterar.png" width="16" height="16" border="0" /></a>&nbsp;&nbsp;<a href="javascript:confirmaExclusao(<%=rs("NUMERO_CAIXA")%>,<%=rs("COD_CAIXA_ARMAZENADO")%>)"><img src="images/icons/icon_excluir.png" width="16" height="16" border="0" /></a></div></td>
                </tr>
         <%
		 	rs.MoveNext()
			Wend
		 %>  
              </table>
         <p>
           <%
		 Else
			Response.Redirect("epv_espaco_de_armazenamento_salas_caixas_caixa.asp?comando=incluir&COD_SALA_ARMAZENADO="&COD_SALA_ARMAZENADO)
		 End If
		 
		 
		 rs.Close
		 Set rs = Nothing
		 FecharConexao()
		 %>  
         </p>
         <p>&nbsp;</p>
         <p>&nbsp;</p></td>

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