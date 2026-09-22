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
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

<%
If Request.ServerVariables("HTTPS") = "on" then 
Response.Write("<link type='text/css' rel='stylesheet' href='https://stillos.com.br/gerencial_foto/css/base.css' media='screen'></link>")
else
Response.Write("<link type='text/css' rel='stylesheet' href='https://stillos.com.br/gerencial_foto/css/base.css' media='screen'></link>")
End if %>

<%
If Request.ServerVariables("HTTPS") = "on" then 
Response.Write("<link rel='stylesheet' media='all' type='text/css' href='https://stillos.com.br/gerencial_foto/css/modal-message.css'>")
else
Response.Write("<link rel='stylesheet' media='all' type='text/css' href='https://stillos.com.br/gerencial_foto/css/modal-message.css'>")
End if %>

<title>Boletim Informativo</title>
<!-- #include virtual="gerencial_foto/files/includes/scripts.asp" -->
<style type="text/css" media="all">
<!--
body {
	background-color: #FFFFFF;
	margin-left: 10px;
	margin-top: 10px;
	margin-right: 10px;
	margin-bottom: 10px;
	font-family:Arial, Helvetica, sans-serif;
	font-size:font-size: 13px;
}
.style1 {font-size: 13px}
.style2 {font-size: 12px}
-->
</style></head>
<body>

  <%

'----------------------------------------------------------------------------------------
'Criado OXWebDeveloper em: 15/10/2011 15:55:46

'Declara vari�veis para utiliza��o do sistema
Dim cod_boletim_informativo, campo1, comando, redirecionar_incluir, redirecionar_alterar, redirecionar_excluir, redirecionar_valicacao, msg, actionformulario, comandoAuxiliar, data_cadastro

cod_boletim_informativo = Request.QueryString("cod_boletim_informativo")

Dim OBS, DATA_PROXIMA_ATUALIZACAO_MONETARIA, CADASTRADO_POR

	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM TB_BOLETIM_INFORMATIVO WHERE cod_boletim_informativo = "&cod_boletim_informativo)

				OBS=trim(rs("OBS"))
				OBS=Replace(OBS,Chr(13),"<br />")
				
				If OBS = "" Then
					OBS = "-"
				End If
				
				DATA_PROXIMA_ATUALIZACAO_MONETARIA=trim(rs("DATA_PROXIMA_ATUALIZACAO_MONETARIA"))
				CADASTRADO_POR=trim(rs("CADASTRADO_POR"))
				data_cadastro = rs("data_cadastro")

		rs.Close
		Set rs=Nothing
	FecharConexao()

'----------------------------------------------------------------------------------------

%>
<div align="center" style="font-size:22px; padding-bottom:10px"><strong>BOLETIM INFORMATIVO - <%=COD_BOLETIM_INFORMATIVO%>/<%=year(data_cadastro)%></strong></div>
<div>
  <span class="style1">Informamos a todos os representantes de álbuns que as instituições listadas abaixo deverão obedecer a preços específicos de venda. 
  </span><br /><br />
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="16%" valign="top"><span class="style2">Observações:</span></td>
        <td width="84%" valign="top"><span class="style2"><%=OBS%></span></td>
      </tr>
      <tr>
        <td valign="top"><span class="style2">Valor válidado até:</span></td>
        <td valign="top"><span class="style2"><%=ArrumaData(DATA_PROXIMA_ATUALIZACAO_MONETARIA )%>, após esta data serão corrigidos pelo IGPM / FGV e este<br /> boletim informativo perderá sua validade.</span></td>
      </tr>
  </table>
<br />

<%
AbrirConexao()

	
	Set rs=objConexao.Execute("SELECT * FROM TB_BOLETIM_INFORMATIVO_CONTRATOS WHERE COD_BOLETIM_INFORMATIVO = '"&COD_BOLETIM_INFORMATIVO&"'")
		If not rs.Eof Then

%>

  <table width='99%' border='0' align='center' cellpadding='3' cellspacing="0" style='border-collapse:collapse; font-size:12px'>
        <tr>
        <td style="border-bottom:solid 1px #999999"><strong>instituições de Ensino</strong> </td>
        </tr>
        <%
        
        %>
              <%While not rs.Eof%>
   
        <tr>
          <td><%
		  Set rs2=objConexao.Execute("SELECT DESCRICAO FROM TB_ENTIDADES WHERE COD_CONTRATO = '"&rs("COD_CONTRATO")&"'")
					
					If not rs2.Eof Then
						Response.Write(rs2(0))
					End If
					
				rs2.Close
				Set rs2=Nothing
		  %></td>
    </tr>
        <%
		rs.MoveNext()
		Wend
		
		%>
        
  </table>
    
    
    <%
		End If
	
	rs.Close
	Set rs=Nothing
	%>
<br />
<%
	Set rs=objConexao.Execute("SELECT * FROM VW_TB_BOLETIM_INFORMATIVO_PRODUTOS WHERE COD_BOLETIM_INFORMATIVO = '"&COD_BOLETIM_INFORMATIVO&"' ORDER BY NOME_PRODUTO ASC")
		If not rs.Eof Then
	%>

<table width='99%' border='0' align='center' cellpadding='3' cellspacing="0" style='border-collapse:collapse; font-size:11px; border-bottom:solid 1px #999999'>
<tr>
<td width='46%' style="border-bottom:solid 1px #999999"><div align="center" style="font-weight: bold">Produto</div></td>
<td width='18%' style="border-bottom:solid 1px #999999"><div align="center" style="font-weight: bold">
  <div align="center">Valor à vista</div>
</div></td>
<td width='18%' style="border-bottom:solid 1px #999999"><div align="center"><strong>Valor em 3X</strong></div></td>
<td width='18%' style="border-bottom:solid 1px #999999"><div align="center" style="font-weight: bold">
  <div align="center">Valor  de 4X à 12X</div>
</div></td>
</tr>
<%
While not rs.Eof
%>
<tr>
  <td style="border-right:solid 1px #999999; border-left:solid 1px #999999"><div align="left"><%=rs("NOME_PRODUTO")%></div></td>
  <td style="border-right:solid 1px #999999"><div align="right"><%=formatCurrency(rs("VALOR_A_VISTA"))%></div></td>
  <td style="border-right:solid 1px #999999"><div align="right"><%=formatCurrency(rs("VALOR_3X"))%></div></td>
  <td style="border-right:solid 1px #999999"><div align="right"><%=formatCurrency(rs("VALOR_PARCELADO"))%></div></td>
</tr>
<%
rs.MoveNext()
Wend
%>
</table>
    
<%
	End If
	
	rs.Close
	Set rs=Nothing
	%>
    <br />
  <div style="font-size:12px; width:99%; margin:auto; padding-bottom:10px">Assinam cientes:</div>
  
    <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
  	<%
	Set rs=objConexao.Execute("SELECT NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR NOT IN(1,27) and  ATIVADO = 1 and COD_GRUPO_VENDEDORES = 0 order by NOME_VENDEDOR",qtdVendedores)
	Dim i, ii, countadorA, countadorB, linhas, resto, colunaA(50), colunaB(50)
	linhas = qtdVendedores/2
	resto  = qtdVendedores mod 2
	
	if resto > 0 Then
		linhas = linhas+0.5
	End if
	
	countadorA = 0
	countadorB = 0
	
	ii = 0
		
	For ii=0 to 1
	
		For i=0 to linhas-1
			
			If ii=0 Then
				If not rs.Eof Then		
					colunaA(i) = Ucase(rs("NOME_VENDEDOR"))
				rs.MoveNext()
				End If
			Elseif ii=1 Then
				If not rs.Eof Then		
					colunaB(i) = Ucase(rs("NOME_VENDEDOR"))
				rs.MoveNext()
				End If
			End If	
				
		Next
	
	Next
	
	rs.Close
	Set rs=Nothing
	
	'/////////////////////////////////////////////

	countadorA = 0
	countadorB = 0
	
	For i=0 to linhas-1
	%>
    	<tr>	
			<%For ii=0 to 1%>
 
     				
                        <%If ii=0 Then%>
                        	<td align="left" style="padding:12px;font-size:10px;">
								<div style="border-bottom:solid 1px #666666;width:100%; float:left"><%Response.Write(colunaA(i))%></div>
                            </td>
						<%Else%>
                        	<td align="left" style="padding:12px;font-size:10px">
                                <div style="border-bottom:solid 1px #666666;width:100%; float:left"><%Response.Write(colunaB(i))%></div>
							</td>
                        <%End If%>
                    
            <%Next%>
       	</tr>
    <%
	Next
	
	%>
  </tr>
</table>

	
	<%
	FecharConexao()
	%>
</div>
</body>
</html>



