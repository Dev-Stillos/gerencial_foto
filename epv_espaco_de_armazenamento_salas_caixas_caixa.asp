<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim COD_SALA_ARMAZENADO, COD_CAIXA_ARMAZENADO, COD_TIPO_CAIXA, COLUNA, NUMERO_CAIXA
Dim comando, sql, rs

COD_SALA_ARMAZENADO	 = Request.QueryString("COD_SALA_ARMAZENADO")
COD_CAIXA_ARMAZENADO = Request.QueryString("COD_CAIXA_ARMAZENADO")
comando				 = Request.QueryString("comando")
COD_TIPO_CAIXA		 = 0

If comando = "alterar" and Request.Form("BOTAO") = Empty Then
		lerBase()			
Elseif comando = "alterar" and Request.Form("BOTAO") <> Empty Then
		alterar()
Elseif  comando = "excluir" Then
		excluir()		
Elseif comando = "incluir" and Request.Form("BOTAO") <> Empty  then
		incluir()
Elseif comando = "incluir" and Request.Form("BOTAO") = Empty  then

	'Verifica a última caixa cadastrada
	Dim ultima_caixa
	ultima_caixa = ""
	AbrirConexao()
	Set rs=objConexao.Execute("SELECT MAX(numero_caixa) as ultima_caixa FROM TB_LOCAL_ARMAZENAGEM_SALA_CAIXA",QTD_ITENS)
		
		If not rs.Eof Then
			ultima_caixa = Cstr(""&rs(0))
		End If
		
	rs.Close
	Set rs=Nothing
	FecharConexao()
	
End If


Sub lerBase()
	
	AbrirConexao()
		
		Set rs=objConexao.Execute("SELECT * FROM TB_LOCAL_ARMAZENAGEM_SALA_CAIXA WHERE COD_CAIXA_ARMAZENADO = "&COD_CAIXA_ARMAZENADO)
				
				if not rs.Eof Then
					
					COD_TIPO_CAIXA	= formatNumber(0&rs("COD_TIPO_CAIXA"),0)
					COLUNA			= rs("COLUNA")
					NUMERO_CAIXA	= Cstr(""&rs("NUMERO_CAIXA"))					
				End If
		
		rs.Close
		Set rs=Nothing
		
	FecharConexao()
	
End Sub

Sub lerFormulario()
	
	COD_CAIXA_ARMAZENADO	= Request.Form("COD_CAIXA_ARMAZENADO")
	COLUNA					= Request.Form("COLUNA")
	COD_TIPO_CAIXA			= Request.Form("COD_TIPO_CAIXA")
	NUMERO_CAIXA			= Request.Form("NUMERO_CAIXA")
		
End Sub


Sub alterar()
	
	lerFormulario()
	
	AbrirConexao()
		
		'Verifica se é numerico o campo
		If isNumeric(NUMERO_CAIXA) Then
			'Verifica se não está em uso
			Set rs=objConexao.Execute("SELECT COD_CAIXA_ARMAZENADO FROM TB_LOCAL_ARMAZENAGEM_SALA_CAIXA WHERE NUMERO_CAIXA = "&sqlTexto(NUMERO_CAIXA)&" AND COD_CAIXA_ARMAZENADO <> "&COD_CAIXA_ARMAZENADO)
				
				If rs.Eof Then
					sql = "UPDATE TB_LOCAL_ARMAZENAGEM_SALA_CAIXA SET COD_TIPO_CAIXA = "&COD_TIPO_CAIXA&", COLUNA='"&COLUNA&"', NUMERO_CAIXA = "&sqlTexto(NUMERO_CAIXA)&" WHERE COD_CAIXA_ARMAZENADO ="&COD_CAIXA_ARMAZENADO&"  LIMIT 1"
		objConexao.Execute(sql)	
				Else
					Response.Redirect("epv_espaco_de_armazenamento_salas_caixas.asp?COD_SALA_ARMAZENADO="&COD_SALA_ARMAZENADO&"&msg=Não pode ser alterada, pois há outra caixa utilizando o número de caixa : "&NUMERO_CAIXA)
				End If
				
			rs.Close
			Set rs=Nothing
		Else
			Response.Redirect("epv_espaco_de_armazenamento_salas_caixas.asp?COD_SALA_ARMAZENADO="&COD_SALA_ARMAZENADO&"&msg=Não pode ser alterada, pois o valor informado não é númerico no campo Número da Caixa")
		End If
		
		
	FecharConexao()
	
	Response.Redirect("epv_espaco_de_armazenamento_salas_caixas.asp?COD_SALA_ARMAZENADO="&COD_SALA_ARMAZENADO&"&msg=Caixa alterada com sucesso!")
	
End Sub

Sub excluir()
	
	AbrirConexao()
		
		Set rs = objConexao.Execute("SELECT * FROM VW_TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES_COMPLETA WHERE COD_CAIXA_ARMAZENADO ="&COD_CAIXA_ARMAZENADO)	
			
				If not rs.Eof Then
					direcionar="epv_espaco_de_armazenamento_salas_caixas.asp?msg=Caixa não pode ser excluída, pois há itens colocados na mesma.<br />Por tanto faça primeiro a transferência para posteriormente excluir esta caixa."
				Else
					sql = "DELETE FROM TB_LOCAL_ARMAZENAGEM_SALA_CAIXA WHERE COD_CAIXA_ARMAZENADO ="&COD_CAIXA_ARMAZENADO
					objConexao.Execute(sql)
					direcionar="epv_espaco_de_armazenamento_salas.asp?msg=Caixa excluída com sucesso!&cod_sala_armazenado="&COD_SALA_ARMAZENADO
				End If
			
		rs.Close
		Set rs = Nothing
		
	FecharConexao()
	Response.Redirect(direcionar)

End Sub

Sub incluir()
	
	lerFormulario()
		
	AbrirConexao()

	'Verifica se é numerico o campo
		If isNumeric(NUMERO_CAIXA) Then
			'Verifica se não está em uso
			Set rs=objConexao.Execute("SELECT COD_CAIXA_ARMAZENADO FROM TB_LOCAL_ARMAZENAGEM_SALA_CAIXA WHERE NUMERO_CAIXA = "&sqlTexto(NUMERO_CAIXA))
				
				If rs.Eof Then
					sql = "INSERT INTO TB_LOCAL_ARMAZENAGEM_SALA_CAIXA(COD_SALA_ARMAZENADO,COD_TIPO_CAIXA,COLUNA,NUMERO_CAIXA) VALUES ("&COD_SALA_ARMAZENADO&","&COD_TIPO_CAIXA&",'"&COLUNA&"',"&sqlTexto(NUMERO_CAIXA)&")"
		objConexao.Execute(sql)	
				Else
					Response.Redirect("epv_espaco_de_armazenamento_salas_caixas.asp?COD_SALA_ARMAZENADO="&COD_SALA_ARMAZENADO&"&msg=Não pode ser incluída, pois há outra caixa utilizando o número de caixa : "&NUMERO_CAIXA)
				End If
				
			rs.Close
			Set rs=Nothing
		Else
			Response.Redirect("epv_espaco_de_armazenamento_salas_caixas.asp?COD_SALA_ARMAZENADO="&COD_SALA_ARMAZENADO&"&msg=Não pode ser incluída, pois o valor informado não é númerico no campo Número da Caixa")
		End If
	
		
	FecharConexao()
	
	Response.Redirect("epv_espaco_de_armazenamento_salas.asp?COD_SALA_ARMAZENADO="&COD_SALA_ARMAZENADO&"&msg=Caixa incluída com sucesso!")
	
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
	
	if(d.NUMERO_CAIXA.value == 0){
		alert('Informe o número da caixa!')
		d.NUMERO_CAIXA.focus();
		return false;
	}else if(d.ANDAR.value == 0){
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

	  <h2>EPV - Estoque / Salas de Armazenagem</h2>
	  <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	 <h3 align="center">Caixa de Armazenagem</h3>
		<form action="epv_espaco_de_armazenamento_salas_caixas_caixa.asp?comando=<%=comando%>&COD_SALA_ARMAZENADO=<%=COD_SALA_ARMAZENADO%>" method="post" style="margin:0px" name="FORMULARIO" onsubmit="return validaForm()">
         <table width="600" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
        
                <tr>
                  <td width="24%"><strong>N&uacute;mero Caixa:</strong></td>
                  <td width="76%"><input type="text" maxlength="5" name="NUMERO_CAIXA" value="<%=NUMERO_CAIXA%>" /> 
                  <%If comando = "incluir" Then%>
                  (&uacute;ltima caixa cadastrada: <%=ultima_caixa%>)
                  <%End If%>
                  </td>
                  </tr>
                <tr>
                  <td><strong>Coluna:</strong></td>
                  <td>
               <select name="COLUNA" id="select">
                  <option <%If COLUNA = "A" Then%> selected="selected" <%End If%> value="A">A</option>
                  <option <%If COLUNA = "B" Then%> selected="selected" <%End If%> value="B">B</option>
                  <option <%If COLUNA = "C" Then%> selected="selected" <%End If%> value="C">C</option>
                  <option <%If COLUNA = "D" Then%> selected="selected" <%End If%> value="D">D</option>
                  <option <%If COLUNA = "E" Then%> selected="selected" <%End If%> value="E">E</option>
                  <option <%If COLUNA = "F" Then%> selected="selected" <%End If%> value="F">F</option>
                  <option <%If COLUNA = "G" Then%> selected="selected" <%End If%> value="G">G</option>
                  <option <%If COLUNA = "H" Then%> selected="selected" <%End If%> value="H">H</option>
	           </select>
               		</td>
                  </tr>
                <tr>
                  <td><strong>Tipo Caixa:</strong></td>
                  <td>
                  <select name="COD_TIPO_CAIXA">
				  <%
				  AbrirConexao()
				  
				  Set rs2 = objConexao.Execute("SELECT * FROM TB_LOCAL_ARMAZENAGEM_SALA_TIPOS_CAIXA")				  
				  If not rs2.Eof Then
				  
				  Dim COD_TIPO_CAIXA_QY
				  
				  
				  While not rs2.Eof
				  
				  COD_TIPO_CAIXA_QY = formatNumber(0&rs2("COD_TIPO_CAIXA"),0)
				  
				  %>
            	<option value="<%=rs2("COD_TIPO_CAIXA")%>" <%If COD_TIPO_CAIXA = COD_TIPO_CAIXA_QY Then%> selected="selected" <%End If%>><%=rs2("NOME_CAIXA")%></option>
                  <%
				  rs2.MoveNext()
				  Wend
				  
				  End If
				  
				  rs2.Close
				  Set rs2 = Nothing
				  
				  FecharConexao()
				  %>
                  </select>
                  </td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td>
                    <input class="btnGravar" type="submit" name="BOTAO" id="button" value=" " />
                    <input value="<%=COD_CAIXA_ARMAZENADO%>" type="hidden" name="COD_CAIXA_ARMAZENADO" /> 
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