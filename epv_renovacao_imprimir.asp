<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
if not validaUsuario(1,388,0,false) then
  response.end()
end if

function retornaVendedor(cod_vendedor)
  dim rs, retorno
  set rs = objConexao.Execute("SELECT NOME_VENDEDOR FROM TB_VENDEDORES WHERE COD_VENDEDOR = "&sqlNumero(cod_vendedor) & " LIMIT 1")
  if not rs.Eof then
    retorno = rs("NOME_VENDEDOR").value
  end if 
  rs.close()
  set rs = nothing

retornaVendedor = retorno
end function

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="all"></link>
<title>StillosWEB - Sistema corporativo - Vers&atilde;o 2.0</title>
<style type="text/css" media="all">
  .style1 {
	  font-size: 7px;
	  font-weight: bold;
  }
  body {
	  background-color: #FFFFFF;
  }
  .style2 {
    font-size: 11px
  }
  .table-border {
    border-bottom:solid 2px #000; 
    border-top:solid 2px #000;
  }
  .folha {
    page-break-after: always;
  }
</style>
</head>
<body onload="window.print()">
<%
  AbrirConexao()
    
  cod_usuario = sqlNumero(Session("COD_USUARIO"))


  sql = "SELECT cod_producao FROM TB_FICHAS_IDENTIFICACAO_RENOVACAO_TEMPORARIA WHERE COD_USUARIO="&cod_usuario& " AND pode_imprimir = 1  GROUP BY COD_PRODUCAO"

  Set rs2=objConexao.Execute(sql)
  while not rs2.Eof 

  

  sql = "SELECT * FROM TB_FICHAS_IDENTIFICACAO_RENOVACAO_TEMPORARIA as temp INNER JOIN TB_FICHAS_IDENTIFICACAO as formandos ON formandos.cod_formando = temp.cod_formando INNER JOIN TB_ENTIDADES as contratos ON contratos.cod_contrato = formandos.cod_contrato WHERE COD_USUARIO="&cod_usuario& " AND pode_imprimir = 1 AND temp.COD_PRODUCAO = " &sqlNumero(rs2("cod_producao").value)

  Set rs=objConexao.Execute(sql)
  If not rs.Eof Then
    nome_vendedor = retornaVendedor(rs("cod_vendedor").value)
    simNao = array(null, "Sim", "Não")


    sql = "SELECT * FROM TB_PRODUCAO WHERE COD_PRODUCAO = " &sqlNumero(rs("cod_producao").value)
    set rs3 = objConexao.execute(sql)
    if not rs3.Eof Then
      descricoa_producao = rs3("cod_producao").value &" - "&rs3("descricao").value
    end if
    rs3.close()
    set rs3 = Nothing
%>
<table class="folha" width="100%" align="center" border="0" cellpadding="0" cellspacing="10">
<tr><td>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<tr><td>
      <table width="100%" border="0" align="center" cellpadding="5" cellspacing="0">
        <tr>
          <td valign="top">
            <h2 style="color:#000000; border-bottom:solid #000 3px; padding-bottom:2px">Renovação</h2>
            <br />
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="82%">
                  <table width="100%" border="0" cellspacing="0" cellpadding="2">
                    <tr>
                      <td class="textoMedio">Data e hora: <strong class="textoMedio"></strong></td>
                      <td class="textoMedio"><div align="left"><strong><%=Now()%></strong></div></td>
                    </tr>
                    <tr>
                      <td class="textoMedio">Repres. Comer.:</td>
                      <td class="textoMedio"><strong><%=Ucase(nome_vendedor)%></strong></td>
                    </tr>
                    <tr>
                      <td class="textoMedio">Produção</td>
                      <td class="textoMedio"><%=descricoa_producao%></td>
                    </tr>
                    <tr>
                      <td class="textoMedio">Gerado por:</td>
                      <td class="textoMedio"><strong><%=Ucase(Session("NOME"))%></strong></td>
                    </tr>
                  </table>
                </td>
                <td width="18%"><div align="center">
                  <img src="images/layout/logomarca_foto_reportagens.jpg" width="92" height="93" />
                </div></td>
              </tr>
            </table>
          </td>
        </tr>
				<tr>
				  <td>
            <table width="100%" border="1" align="center" 	cellpadding="1" cellspacing="0" style="border-collapse:collapse">
					    <tr bgcolor="whitesmoke">
                <td width="7%"><div align="center"><strong>Cód. Formando </strong></div></td>
                <td width="4%"><div align="center"><strong>Álbum</strong></div></td>
                <td width="27%"><div align="center"><strong>Nome</strong></div></td>
                <td width="21%"><div align="center"><strong>Entidade de Ensino </strong></div></td>
                <td width="7%"><div align="center"><strong>Comissão</strong></div></td>
                <td width="10%"><div align="center"><strong>Brinde</strong></div></td>
                <td width="10%"><div align="center"><strong>Periodo Renovação</strong></div></td>
				      </tr>
              <%

              while not rs.Eof
              %>
              <tr>
                <td><div align="center"><strong><%=rs("cod_formando").value%></strong></div></td>
                <td><div align="center"><strong><%=rs("album").value%></strong></div></td>
                <td><div align="center"><strong><%=rs("nome_formando").value%></strong></div></td>
                <td><div align="center"><strong><%=rs("descricao")%></strong></div></td>
                <td><div align="center"><strong><%=simNao(rs("comissao").value)%></strong></div></td>
                <td><div align="center"><strong><%=simNao(rs("brinde").value)%></strong></div></td>
                <td><div align="center"><strong><%=rs("periodo_renovacao").value%></strong></div></td>
              <tr>
              <%
              rs.moveNext()
              wend
              %>
            </table>
            <p align="justify">* Devolução de Fichas: É obrigatório a devolução do álbum e demais itens do formando no estoque.</p>
						<table width="350" border="0" align="center" cellpadding="0" cellspacing="0">
						  <tr>
								<td style="border-bottom:solid 1px #000000">
                  <p> </p>
								  <p> </p>
                  <br />
                </td>
							</tr>
							<tr>
							  <td><div align="center"><span class="textoMedio"><%=NOME_VENDEDOR%></span></div></td>
							</tr>
						</table>
          </td>
        </tr>
      </table>
  </td></tr>
  </table>
</td></tr>
</table>
<%
end if
rs.close()
set rs = nothing

rs2.moveNext()
wend
rs2.close
set rs2 = nothing
%>