<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Para tirar a validação da página basta tirar o "#" ou coloca-lo para habilitar antes do include
Dim COD_FORMANDO, CREDITO_EM_VALOR, CREDITO_MOTIVO, BOTAO, CARREGA_PAGINA

COD_FORMANDO	= Request.QueryString("cod_formando")
BOTAO			= Request.Form("btnGravar")

If BOTAO <> Empty Then

	COD_FORMANDO		= Request.Form("COD_FORMANDO")
	CREDITO_EM_VALOR	= Request.Form("CREDITO_EM_VALOR")
	CREDITO_MOTIVO		= Request.Form("CREDITO_MOTIVO")
	
	CREDITO_EM_VALOR	= sqlNumero(CREDITO_EM_VALOR)
	If Not isNumeric(CREDITO_EM_VALOR) Then
		Response.Redirect("dip_busca_detalhamento_formando_POPUP_credito_valor.asp?cod_formando="&COD_FORMANDO&"&msg=Valor do crédito não é numérico.")
	End If
	
	AbrirConexao()
		
		sql =	" UPDATE TB_FICHAS_IDENTIFICACAO SET CREDITO_DATA_INSERIDO ="&sqlData(Date())&", CREDITO_EM_VALOR = '"&CREDITO_EM_VALOR&"', CREDITO_MOTIVO = "&sqlTexto(CREDITO_MOTIVO)&_
				" WHERE COD_FORMANDO ="&COD_FORMANDO&" LIMIT 1"

		objConexao.Execute(sql)
		
	FecharConexao()

End If


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<style type="text/css">
<!--
body {
	background-color: #FFFFFF;
}
-->
</style>
<script language="javascript">
	function validaForm(){
		d = document.FORMULARIO
		if (d.CREDITO_EM_VALOR.value == "" || d.CREDITO_MOTIVO.value == ""){
			alert('Preencha todos os campos')
			d.CREDITO_EM_VALOR.focus();
			return false;
		}
	
	}
	function carregaCampo(NOME_CAMPO){
		document.FORMULARIO[NOME_CAMPO].focus();
		
	}
	function fecharPopUp(){
			//window.opener.document.getElementById("").style.color = '#FF0000'
			window.opener.location.reload();
			self.close();
	}
</script>

<%

	If BOTAO <> Empty Then
 		CARREGA_PAGINA = "onload=""fecharPopUp()"""
 	Else
 		CARREGA_PAGINA = "onload=""carregaCampo('CREDITO_EM_VALOR')"""
	End If
%>
<body <%=CARREGA_PAGINA%>>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td bgcolor="#FFFFFF"><h1 style="margin:0px; padding:0px">DIP - Inclus&atilde;o de Cr&eacute;dito em Valor </h1>
          <div align="left">C&oacute;d. Formando.: <strong><%=digito(COD_FORMANDO)%></strong></div>
		  <div class="msg">
		    <div align="center"><%=Request.QueryString("msg")%></div>
		  </div>
	<form name="FORMULARIO" method="post" action="dip_busca_detalhamento_formando_POPUP_credito_valor.asp" style="margin:0px" onsubmit="return validaForm()">
		   <table width="325" cellpadding="3" cellspacing="1" border="0" align="center">
		   	<tr>
               <td colspan="2"><div align="center" style="border-bottom: solid 1px #CCCCCC"><strong>Informe o valor de cr&eacute;dito.</strong></div></td>
               </tr>
             <tr>
               <td width="32%"><div align="right">Valor do Cr&eacute;dito:</div></td>
               <td width="68%"><input style="width:200px" type="text" name="CREDITO_EM_VALOR" /></td>
               </tr>
             <tr>
               <td valign="top"><div align="right">Motivo: </div></td>
               <td><textarea style="width:200px" name="CREDITO_MOTIVO" rows="3"></textarea></td>
               </tr>
             <tr>
               <td colspan="2"><div align="center">
                 <input type="hidden" value="<%=COD_FORMANDO%>" name="COD_FORMANDO" />
                 <input class="btnGravar" type="submit" name="btnGravar" value=" " />
               </div></td>
               </tr>
           </table>
	</form>
		 
	 </td>
      </tr>
    </table></td>
  </tr>
</table>


</body>
</html>
