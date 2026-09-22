<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim COD_SALA_ARMAZENADO, NOME_SALA_ARMAZENADO, ANDAR
Dim comando, sql, rs

COD_SALA_ARMAZENADO	= Request.QueryString("COD_SALA_ARMAZENADO")
comando				= Request.QueryString("comando")
ANDAR				= 0

If comando = "alterar" and Request.Form("BOTAO") = Empty Then
		lerBase()			
Elseif comando = "alterar" and Request.Form("BOTAO") <> Empty Then
		alterar()
Elseif  comando = "excluir" Then
		excluir()		
Elseif comando = "incluir" and Request.Form("BOTAO") <> Empty  then
		incluir()
End If


Sub lerBase()
	
	AbrirConexao()
		
		Set rs=objConexao.Execute("SELECT * FROM TB_LOCAL_ARMAZENAGEM_SALA WHERE COD_SALA_ARMAZENADO = "&COD_SALA_ARMAZENADO)
				
				if not rs.Eof Then
					
					NOME_SALA_ARMAZENADO	= rs("NOME_SALA_ARMAZENADO")
					ANDAR					= formatNumber(0&rs("ANDAR"),0)
					
				End If
		
		rs.Close
		Set rs=Nothing
		
	FecharConexao()
	
End Sub

Sub lerFormulario()
	
	COD_SALA_ARMAZENADO		= Request.Form("COD_SALA_ARMAZENADO")
	NOME_SALA_ARMAZENADO	= Request.Form("NOME_SALA_ARMAZENADO")
	ANDAR					= Request.Form("ANDAR")
		
End Sub


Sub alterar()
	
	lerFormulario()
	
	AbrirConexao()
		sql = "UPDATE TB_LOCAL_ARMAZENAGEM_SALA SET NOME_SALA_ARMAZENADO = "&sqlTexto(NOME_SALA_ARMAZENADO)&", ANDAR = "&ANDAR&" WHERE COD_SALA_ARMAZENADO ="&COD_SALA_ARMAZENADO&" LIMIT 1"
		objConexao.Execute(sql)	
	FecharConexao()
	
	Response.Redirect("epv_espaco_de_armazenamento_salas.asp?msg=Sala alterada com sucesso!")
	
End Sub

Sub excluir()
	
	AbrirConexao()
		
		Set rs = objConexao.Execute("SELECT * FROM TB_LOCAL_ARMAZENAGEM_SALA_CAIXA WHERE COD_SALA_ARMAZENADO ="&COD_SALA_ARMAZENADO)	
			
				If not rs.Eof Then
					direcionar="epv_espaco_de_armazenamento_salas.asp?msg=Sala não pode ser excluída, pois há caixas utilizando está sala.<br />Por tanto exclua primeiro as caixas desta sala para depois exclui-la."
				Else
					sql = "DELETE FROM TB_LOCAL_ARMAZENAGEM_SALA WHERE COD_SALA_ARMAZENADO ="&COD_SALA_ARMAZENADO
					objConexao.Execute(sql)
					direcionar="epv_espaco_de_armazenamento_salas.asp?msg=Sala excluída com sucesso!"
					
				End If
			
		rs.Close
		Set rs = Nothing
		

	FecharConexao()
	Response.Redirect(direcionar)

End Sub

Sub incluir()
	
	lerFormulario()
		
	AbrirConexao()
		sql = "INSERT INTO TB_LOCAL_ARMAZENAGEM_SALA(NOME_SALA_ARMAZENADO, ANDAR) VALUES ("&sqlTexto(NOME_SALA_ARMAZENADO)&", "&ANDAR&")"
		objConexao.Execute(sql)	
	FecharConexao()
	
	Response.Redirect("epv_espaco_de_armazenamento_salas.asp?msg=Sala incluída com sucesso!")
	
End Sub
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
<script language="javascript">
function validaForm(){
	d=document.FORMULARIO
	if(d.ANDAR.value == 0){
		alert('Selecione o Andar!')
		d.ANDAR.focus();
		return false;
	}else if(d.NOME_SALA_ARMAZENADO.value == ""){
		alert('Preencha o campo Nome da Sala!')
		d.NOME_SALA_ARMAZENADO.focus();
		return false;
	}else{
		return true;
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
  BTN_FONTE = 1
  BTN_IMPRIMIR_LINK = "#"
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">

	  <h2>EPV - Estoque / Salas de Armazenagem</h2>
	  <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	 <h3 align="center">Sala de Armazenagem</h3>
		<form action="epv_espaco_de_armazenamento_salas_sala.asp?comando=<%=comando%>" method="post" style="margin:0px" name="FORMULARIO" onsubmit="return validaForm()">
         <table width="600" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
                <tr>
                  <td width="24%"><strong>Pr&eacute;dio:</strong></td>
                  <td width="76%">Antiga sede Stillo's</td>
                  </tr>
                <tr>
                  <td><strong>Nome da Sala:</strong></td>
                  <td><input class="inputTamanho2" value="<%=NOME_SALA_ARMAZENADO%>" type="text" name="NOME_SALA_ARMAZENADO" />                  </td>
                  </tr>
                <tr>
                  <td><strong>Andar:</strong></td>
                  <td>
                    <select name="ANDAR" id="select">
                    <option <%If int(ANDAR) = 0 Then%> selected="selected" <%End If%> value="0">Selecione o andar da Sala</option>
                    <option <%If int(ANDAR) = 1 Then%> selected="selected" <%End If%> value="1">(1) - Sub-Solo</option>
                    <option <%If int(ANDAR) = 2 Then%> selected="selected" <%End If%> value="2">(2) - Térreo</option>
                    <option <%If int(ANDAR) = 3 Then%> selected="selected" <%End If%> value="3">(3) - 1º Andar</option>                                        
                    </select>                  </td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td>
                    <input class="btnGravar" type="submit" name="BOTAO" id="button" value=" " />
                    <input value="<%=COD_SALA_ARMAZENADO%>" type="hidden" name="COD_SALA_ARMAZENADO" /> 
                  </td>
                </tr>
              </table>
        </form>     
         
		   		<br />
		   		<br /></td>

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