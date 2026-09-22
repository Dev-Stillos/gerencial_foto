<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'Valida��o especial para verificar se possui sub-valida��o espec�fica
'////////////////////////////////////////////////////////////////
AbrirConexao()
	
	Set rs=objConexao.Execute("Select COD_PERMISSAO_SUB_VALIDACAO From TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES where COD_SUB_VALIDACAO = '49' and cod_perfil = '"&Session("cod_perfil")&"'")
		
		If rs.Eof Then
		'Se n�o encontrar ent�o redireciona
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
		
	rs.Close
	Set rs=Nothing
	
FecharConexao()
'////////////////////////////////////////////////////////////////



Dim cod_producao, descricao
cod_producao = Request.QueryString("cod_producao")

	  AbrirConexao()
			  Set rs = objConexao.Execute("SELECT DESCRICAO FROM TB_PRODUCAO WHERE COD_PRODUCAO="&cod_producao&" and PRODUCAO_FINALIZADA = 1")
					If not rs.Eof then		
						DESCRICAO = rs("DESCRICAO")
					Else
						Response.Redirect("adm_relatorios.asp?msg=Produção não encontrada!")
					End If					
			  rs.Close()
	  		  Set rs = Nothing
	  FecharConexao()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="all"></link>
<title>Relatório de Formandos</title>

    <style type="text/css">
<!--
.style1 {color: #000000}
-->
    </style>
	<script language="javascript">
function imprimeFecha() {
window.print();
}
</script>
</head>

<body onload="imprimeFecha()">
<table cellpadding="0px" cellspacing="0" width="100%">
<tr>
<td>
<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
  <td>
  <%
	Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
	BTN_VOLTAR = 1
	BTN_IMPRIMIR = 1
	BTN_FONTE = 1
  %>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="top">
	
	  <h2 class="style1">Relatórios de Formandos por Produ&ccedil;&atilde;o</h2>
	  <p><span class="textoMedio"><strong>Produ&ccedil;&atilde;o: <%=COD_PRODUCAO%> - <%=DESCRICAO%>
	  </strong><br />
      <strong>Contratos: <%
	  AbrirConexao()
		Dim rsContratos, qtdContratos, CONTRATOS_ULT
		
		Set rsContratos = objConexao.Execute("SELECT DISTINCT(COD_CONTRATO) AS COD_CONTRATO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO="&COD_PRODUCAO&" ORDER BY COD_CONTRATO ASC",qtdContratos)
							
							If not rsContratos.Eof then
							
								If qtdContratos > 1 then
							
									CONTRATOS_ULT = ""
									For i = 0 to qtdContratos-1
										
										If i = qtdContratos-1 then
										
											CONTRATOS_ULT = CONTRATOS_ULT&rsContratos("COD_CONTRATO")
										
										else
										
											CONTRATOS_ULT = CONTRATOS_ULT&rsContratos("COD_CONTRATO")&", "
										
										End If
									
									rsContratos.MoveNext()
									Next
									
									Response.Write(CONTRATOS_ULT)
									
								else
									
									Response.Write(rsContratos("COD_CONTRATO"))
								
								End If
								
									
									
							else
							
							Response.Write("-")
							
							End If
							
							rsContratos.Close()
							Set rsContratos = Nothing							
							%></strong>
      </span>      
	  <br />
      
      <%
	  sql = "SELECT cod_formando, NOME_FORMANDO, NOME_CURSO, album, cod_contrato, vendido FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO="&cod_producao&" ORDER BY NOME_FORMANDO"
	  Set rs=objConexao.Execute(sql,qtdRegistros)
	  
	  If not rs.Eof Then
	  %>
	  <table width="100%" border="1" align="center" cellpadding="2" style="border-collapse:collapse"><tr bgcolor="whitesmoke"><td width="26%" height="28"><div align="center"><span class="MensagemFicha"><strong>Formando</strong></span><strong><span class="MensagemFicha"></span></strong></div></td><td width="15%"><div align="center"><strong>Curso</strong></div></td>
                <td width="15%"><div align="center"><strong>&Aacute;lbum</strong></div></td>
                <td width="31%"><div align="center"><strong>Institui&ccedil;&atilde;o de Ensino</strong></div></td>
                <td width="13%"><div align="center"><strong>Vendido</strong></div></td>
                </tr>
		<%While not rs.Eof%>     
				<tr>
		          <td><span class="textoMenor"><%=digito(rs("COD_FORMANDO"))%> - <%=rs("NOME_FORMANDO")%></span></td>
                  <td><span class="textoMenor"><%=rs("NOME_CURSO")%></span></td>
                  <td><div align="center" class="textoMenor"><%=rs("ALBUM")%></div></td>
                  <td><span class="textoMenor">
				  	<%
								Dim rsTurma
								Set rsTurma = objConexao.Execute("SELECT DESCRICAO FROM TB_ENTIDADES WHERE COD_CONTRATO="&sqlTexto(rs("COD_CONTRATO")))
									
									If not rsTurma.Eof Then
										Response.Write(rsTurma("DESCRICAO"))
									Else%>
								  <span class="style1">Falha entidade de ensino não encontrada, informe ao depto. de T.I</span></span>
								  <span class="textoMenor textoMenor">
									<%End If
								
								rsTurma.Close
								Set rsTurma = Nothing
								%>
                    </span></td>
                  <td>
                    <span class="textoMenor textoMenor">
                      <%
				  Dim VENDIDO
				  VENDIDO = formatNumber(0&rs("VENDIDO"),0)
				  %>
                      </span>
                    <div align="center" class="textoMenor textoMenor">
                      <%
					Set rsVera = objConexao.Execute("SELECT COD_PEDIDO, FINALIZA_CORRECAO FROM VW_PEDIDOS_COMPLETO WHERE COD_FORMANDO ="&rs("COD_FORMANDO")&" AND FINALIZA_CORRECAO = 1")
							
					If not rsVera.Eof Then
						VENDIDO = "Sim - Pedido: "&rsVera("COD_PEDIDO")	
					Else
						VENDIDO = "Não"
					End If
							
							
					rsVera.Close
					Set rsVera = Nothing
					  
					Response.Write(VENDIDO)	
						
						%>
                    </div></td>
                  </tr>
                  
					 
		<%
		rs.MoveNext()
		Wend
		%>
        <tr bgcolor="whitesmoke">
		  				    <td colspan="5"><div align="center"><strong>Quantidade Formandos: <%=qtdRegistros%></strong></div></td>
                        </tr>
      </table>
		
		  <p>
           </p>
		    <%Else%>
		   <div align="center">
           <br /><br /><br /><br />
		  <p><span class="textoMedio"><strong>N�o h� dados para os filtros selecionados!</strong></span></p>
          <br /><br /><br /><br />
		</div>
		    <%
			End If
			
			rs.Close
			Set rs=Nothing
			
			FecharConexao()
			%>
		<div align="center"><br />
            <br />
            <br />
            <br />
            <br />
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
</td>
</tr>
</table>
</body>
</html>