<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%
Dim rs, botaoIncluir, botaoAlterar, ERRO_LISTA, COD_PEDIDO, TXT_REQUEST, msg, FORMANDO, FORMANDO2, FORMANDO3, FORMANDO4, erro_tipo

COD_PEDIDO = Request.QueryString("cod_pedido")

AbrirConexao()
Set rs = objConexao.Execute("SELECT ERRO_LISTA FROM TB_PEDIDOS WHERE COD_PEDIDO="&COD_PEDIDO)
	IF not rs.Eof then
		ERRO_LISTA=rs("ERRO_LISTA")
	end if
rs.Close
Set rs = Nothing
FecharConexao()

If Request.Form("Submit") = "Incluir" Then

AbrirConexao()

TXT_REQUEST = Request.Form("TXT_ERRO")

iF TXT_REQUEST = Empty then

	TXT_REQUEST = "Inclu�do em "&now()

end if

objConexao.Execute("UPDATE TB_PEDIDOS SET ERRO_LISTA = "&sqlTexto(TXT_REQUEST)&" WHERE COD_PEDIDO ="&COD_PEDIDO&" LIMIT 1")

FecharConexao()
msg = "Pend�ncia inclu�da na ficha!"
erro_tipo = 1


End If

If Request.Form("Submit3") = "Alterar" Then

AbrirConexao()
erro_tipo = 1
TXT_REQUEST = Request.Form("TXT_ERRO")

iF TXT_REQUEST = Empty then

	TXT_REQUEST = "Alterado em "&now()

end if

objConexao.Execute("UPDATE TB_PEDIDOS SET ERRO_LISTA = "&sqlTexto(TXT_REQUEST)&" WHERE COD_PEDIDO ="&COD_PEDIDO&" LIMIT 1")

FecharConexao()

msg = "Pend�ncia alterada da ficha!"

End If

If Request.Form("Submit2") = "Apagar" Then
erro_tipo = 2
AbrirConexao()

objConexao.Execute("UPDATE TB_PEDIDOS SET ERRO_LISTA = NULL WHERE COD_PEDIDO ="&COD_PEDIDO&" LIMIT 1")

FecharConexao()

msg = "Pend�ncia apagada da ficha!"

End If


%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script>
<!--
function fecharNovo1(){
	window.opener.document.getElementById("cor_<%=COD_PEDIDO%>").style.color = '#FF0000'
	//backgroundImage='url(icons/icon_green.jpg)'; 
	//window.opener.document.vendas.cor_[NUM_LINK].value='1';
	self.close();
}

function fecharNovo2(){
	window.opener.document.getElementById("cor_<%=COD_PEDIDO%>").style.color = '#000000'
	//backgroundImage='url(icons/icon_green.jpg)'; 
	//window.opener.document.vendas.cor_[NUM_LINK].value='1';
	self.close();
}
//-->
</script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Observações</title>
<style type="text/css">
<!--
body,td,th {
	font-family: Arial, Tahoma;
	font-size: 13px;
	color: #000000;
}
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style1 {
	font-size: 15px;
	font-weight: bold;
	color: #990000;
}
-->
</style></head>
<%if msg <> Empty then%>
<body onload="fecharNovo<%=erro_tipo%>()">
<%else%>
<body>
<%end if%>
<table width="329" height="180" border="0" cellspacing="0" cellpadding="5">
  <tr>
   <td valign="top"><div style="background-color:#EAEAEA">
     </div>
     <div align="center">
       <p>
         <%If msg <> Empty then%>
         <p><strong><%=msg%></strong></p></p>
      
	  	<input onclick="javascript:send()" type="button" value="Fechar" />
	  
         <%else%>
         
     </div>
     <form style="margin:0px" method="post">
         <div align="center">
           <p align="center"><strong>Observações:</strong><br />
             <textarea name="TXT_ERRO" cols="33" rows="3"><%=ERRO_LISTA%></textarea><br />
             &nbsp; 
<%             If IsNull(ERRO_LISTA) then%>
			 <input style="margin-top:10px" type="submit" name="Submit" value="Incluir" />  
             <%else%>
			 <input style="margin-top:10px" type="submit" name="Submit3" value="Alterar" />
          &nbsp; 
          <input style="margin-top:10px" type="submit" name="Submit2" value="Apagar" />
			 <%end if%>
           </p>
           <p>&nbsp;</p>
         </div>
	  </form>
<%end if%>

</td>
  </tr>
</table>
</body>
</html>
