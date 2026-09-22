<!-- #include virtual="gerencial_foto/includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include virtual="gerencial_foto/includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
Dim caminho_pagina
caminho_pagina = Request.ServerVariables("URL")
caminho_pagina = Replace(caminho_pagina,"/gerencial_foto/","")
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <style>
table[style*="background-color: #2f2f44;"] a,
table[style*="background-color: #ff007f;"] a:link,
table[style*="background-color: #2f2f44;"] a:visited,
table[style*="background-color: #2f2f44;"] strong,
table[style*="background-color: #2f2f44;"] td {
  color: white !important;
  background-color: #2f2f44 !important;
  text-decoration: none !important;
}

table[style*="background-color: #ff007f;"] a:hover {
  color: #ffe6f0 !important;
  background-color: #cc0066 !important;
  text-decoration: none !important;
}
</style>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link rel="stylesheet" type="text/css" href="css/bootstrap.min.css"></link>

<%
If Request.ServerVariables("HTTPS") = "on" then 
Response.Write("<link type='text/css' rel='stylesheet' href='http://stillos.com.br/gerencial_foto/css/base.css' media='screen'></link>")
else
Response.Write("<link type='text/css' rel='stylesheet' href='http://stillos.com.br/gerencial_foto/css/base.css' media='screen'></link>")
End if %>

<%
If Request.ServerVariables("HTTPS") = "on" then 
Response.Write("<link type='text/css' rel='stylesheet' href='http://stillos.com.br/gerencial_foto/css/chosen.min.css' media='screen'></link>")
else
Response.Write("<link type='text/css' rel='stylesheet' href='http://stillos.com.br/gerencial_foto/css/chosen.min.css' media='screen'></link>")
End if %>

<%
If Request.ServerVariables("HTTPS") = "on" then 
Response.Write("<link rel='stylesheet' media='all' type='text/css' href='http://stillos.com.br/gerencial_foto/css/modal-message.css'></link>")
else
Response.Write("<link rel='stylesheet' media='all' type='text/css' href='http://stillos.com.br/gerencial_foto/css/modal-message.css'></link>")
End if %>

<title>StillosWEB - Sistema corporativo - Vers&atilde;o 2.0</title>
<!-- #include virtual="gerencial_foto/files/includes/scripts.asp" -->
<!-- #include virtual="gerencial_foto/files/includes/functions_local.asp" -->

</head>

<body>
<div class="bloqueio_tela hidden">
</div>
<table cellpadding="9px" cellspacing="0" width="100%">
<tr>
<td>
<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td>
		<!-- #include virtual="gerencial_foto/includes/topo.asp" -->
<table width="100%" style="border-bottom: dotted 1px #999999; background-color: #2f2f44;" border="0" cellspacing="0" cellpadding="0">

  <tr>
    <td>
	<!-- #include virtual="gerencial_foto/includes/menu.asp" -->
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="10">
  <tr>
  <td>
  <% dim jquery : jquery = true %>
  <!-- #include virtual="gerencial_foto/includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">