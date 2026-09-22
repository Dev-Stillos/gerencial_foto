<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>BackUP - Foto</title>
<style type="text/css">
<!--
.style1 {
	font-size: 36px;
	font-weight: bold;
}
.style4{
	color:#33CC00;
	font-weight: bold;
}
body,td,th {
	font-family: Arial, Helvetica, sans-serif;
}
body {
	background-color: #E9E9E9;
}
-->
</style>
</head>
<body>
<%
Dim COD_BACKUP, NOME_PASTA, NOME_BASE, QTD_LINHAS_POR_VEZ,TEMPO_REATUALIZAR
COD_BACKUP = formatNumber(0&Request.QueryString("COD_BACKUP"),0)
NOME_BASE  = Request.QueryString("NOME_BASE")


If COD_BACKUP = 0 Then
	Response.Redirect("http://stillos.com.br/gerencial_foto?msg=Falha no backup!")
End If

'////////////////////////////////////
'CONEXÃO COM A BASE AUXILIAR
Dim objConexaoBACK

Sub AbrirConexaoBACK()
Set objConexaoBACK = Server.CreateObject("ADODB.Connection")
	objConexaoBACK.Open("DRIVER={MySQL ODBC 5.1 Driver};SERVER=187.45.196.164;PORT=3306;DATABASE=stillos4;USER=stillos4;PASSWORD=fca123fca123@;")
End Sub


'Sub Rotina que Fecha a Conexao
Sub FecharConexaoBACK()
   objConexaoBACK.Close
   Set objConexaoBACK=nothing
End Sub


'////////////////////////////////////
				AbrirConexaoBACK()
				
				
				Dim DATA_TERMINO, TEMPO_TOTAL, DATA

				Set rsFim = objConexaoBACK.Execute("SELECT * FROM TB_BACKUPS WHERE COD_BACKUP = "&COD_BACKUP)
			
						If not rsFim.Eof Then
					
						DATA			= formatDateTime(rsFim("DATA"))			
						DATA_TERMINO	= formatDateTime(rsFim("DATA_TERMINO"))
					
				
						End If
										
				rsFim.Close
				Set rsFim = Nothing
				
				TEMPO_TOTAL		= DateDiff("n",DATA,DATA_TERMINO)
%>
				<table width="80%" border="0" align="center" cellpadding="4" cellspacing="0" style="border:dashed #999999 1px">
				  <tr>
					<td style="border-bottom:solid 3px #666666"><div align="center" class="style1">BACKUP <%=COD_BACKUP%> / <%=UCASE(NOME_BASE)%></div>      <div align="center"></div></td>
					</tr>
				
				
			  <tr>
				<td bgcolor="#FFFFFF" style="border-bottom: solid 2px #CCCCCC"><div align="center">
				  <p><br />
					<br />
					<strong>BACKUP FINALIZADO! </strong><br />
				  </p>
				  <table width="400" border="0" cellspacing="0" cellpadding="2">
					<tr>
					  <td width="202">
						<div align="right">Inicío em: </div></td>
					  <td width="198"><strong> <%=DATA%></strong></td>
					  </tr>
					<tr>
					  <td><div align="right">Finalizado em:</div></td>
					  <td><strong> <%=DATA_TERMINO%></strong></td>
					  </tr>
					<tr>
					  <td><div align="right">Tempo total em minutos: </div></td>
					  <td><strong> <%=TEMPO_TOTAL%></strong></td>
					  </tr>
				  </table>
				  <p>         </p>
				  <p><br />
					<br />
					<br />
				  </p>
				</div></td>
				</tr>
				</table>
<%
FecharConexaoBACK()
%>
</body>
</html>



