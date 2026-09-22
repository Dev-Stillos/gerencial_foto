<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta http-equiv="cache-control" content="max-age=0" />
<meta http-equiv="cache-control" content="no-cache" />
<meta http-equiv="expires" content="0" />
<meta http-equiv="expires" content="Tue, 01 Jan 1980 1:00:00 GMT" />
<meta http-equiv="pragma" content="no-cache" />

<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>

<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
  <link rel="stylesheet" href="//code.jquery.com/ui/1.10.4/themes/smoothness/jquery-ui.css">
  <script src="js/jquery-1.8.3.js"></script>
  <script src="js/jquery-ui-1.9.2.custom.js"></script>
  <!-- Implementa&ccedil&atildeo : by Luiz -->    
  <link type="text/css" rel="stylesheet" href="includes/css/style.css" media="screen"></link>
  <script type="text/javascript" src="includes/js/validacao.js?ver=<%=now()%>"></script> 
  <script type="text/javascript" src="includes/js/est_script.js?ver=<%=now()%>"></script>
  <script type="text/javascript" src="includes/js/utils.js?ver=<%=now()%>"></script>
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
              	  <h2>EST -  Estoque Interno</h2>
              	  <%
                    ' Controle de permiss&atildeo de acesso 
                    ' Acesso completo
                    '   9  - &Aacutelcio Freitas  
                    '   12 - Perfil Diretoria 
                    '   25 - Perfil Assist&ecircncia Diretoria 
                    '   3  - Perfil Administrativo

                    ' Acesso limitado
                    '   Perfil Fotografia
                    cod_usuario = Session("COD_USUARIO")
                    cod_perfil = Session("COD_PERFIL")
                    if cod_perfil = 25 or cod_perfil = 12 or cod_perfil = 3 or cod_usuario = 9 then
                      perfil = 1
                    else
                      perfil = 2
                    end if
              	   %>
                  <!-- #include file="includes/adm_estoque_sub_menu.asp" -->
              		<div id="content">
                  </div>
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
