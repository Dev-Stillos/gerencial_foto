<!-- #include file="includes/parametrosIniciais.asp" -->
<!-- #include file="restrito/includes/xqrtopqqetyynmkgfdconexao.asp" -->
<%
pagAtual = 5
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" media="all" type="text/css" href="estilo.css"/>
<!-- #include file="includes/metas.asp" -->
<title>Nante Indústria e Comércio - Orçamento</title>
<link rel="shortcut icon" href="http://www.oxnet.com.br/clientes/nante/images/layout/favicon.ico" />
<meta http-equiv="Content-Style-Type" content="text/css" />
<!-- #include file="includes/scripts.asp" -->
</head>
<body>
<%
Dim produtos_selecionados, quantidades_selecionadas, empresa,contato, telefone, email, mensagem, corpo_email, i

produtos_selecionados = split(request.Form("produtos_selecionados"),";")
quantidades_selecionadas = split(request.Form("quantidades_selecionadas"),";")
empresa = request.Form("NOME_EMPRESA")
contato = request.Form("NOME_CONTATO")
telefone = request.Form("telefone")
email = request.Form("email")
mensagem = request.Form("mensagem")


corpo_email = "<center><table style='border-collapse: collapse;' cellpadding='2' width='90%' border='1'>"
corpo_email = corpo_email & "<tr><td width='173'><div align='right'><font face='Arial, Helvetica, sans-serif' size='2'><strong>Empresa:</strong></font></div></td><td width='885'><div align='left'><font face='Arial, Helvetica, sans-serif' size='2'>"&empresa&"</font></div></td></tr>"
corpo_email = corpo_email & "<tr><td><div align='right'><font face='Arial, Helvetica, sans-serif' size='2'><strong>Contato:</strong></font></div></td><td><div align='left'><font face='Arial, Helvetica, sans-serif' size='2'>"&contato&"</font></div></td></tr>"
corpo_email = corpo_email & "<tr><td><div align='right'><font face='Arial, Helvetica, sans-serif' size='2'><strong>Telefone:</strong></font></div></td><td><div align='left'><font face='Arial, Helvetica, sans-serif' size='2'>"&telefone&"</font></div></td></tr>"
corpo_email = corpo_email & "<tr><td><div align='right'><font face='Arial, Helvetica, sans-serif' size='2'><strong>E-mail:</strong></font></div></td><td><div align='left'></div><font face='Arial, Helvetica, sans-serif' size='2'>"&email&"</font></td></tr>"
corpo_email = corpo_email & "<tr><td height='19'><div align='right'><font face='Arial, Helvetica, sans-serif' size='2'><strong>Mensagem:</strong></font></div></td><td><div align='left'><font face='Arial, Helvetica, sans-serif' size='2'>"&replace(mensagem,chr(10),"<br />")&"</font></div></td></tr>"
corpo_email = corpo_email & "</table><br />"

corpo_email = corpo_email & "<table width='90%' border='1' cellpadding='2' style='border-collapse: collapse;'>"
corpo_email = corpo_email & "<tr>"
corpo_email = corpo_email & "<td width='18%'><div align='center'><font face='Arial, Helvetica, sans-serif' size='2'><strong>Codigo</strong></font></div></td>"
corpo_email = corpo_email & "<td width='40%'><div align='center'><font face='Arial, Helvetica, sans-serif' size='2'><strong>Produto</strong></font></div></td>"
corpo_email = corpo_email & "<td width='10%'><div align='center'><font face='Arial, Helvetica, sans-serif' size='2'><strong>Grupo</strong></font></div></td>"
'corpo_email = corpo_email & "<td width='11%'><div align='center'><font face='Arial, Helvetica, sans-serif' size='2'><strong>Quantidade por <br />  embalagem</strong></font></div></td>"
corpo_email = corpo_email & "<td width='9%'><div align='center'><font face='Arial, Helvetica, sans-serif' size='2'><strong>Qtd.</strong></font></div></td>"
corpo_email = corpo_email & "<td width='11%'><div align='center'><font face='Arial, Helvetica, sans-serif' size='2'><strong>Unidade Medida</strong></font></div></td>"
corpo_email = corpo_email & "<td width='12%'><div align='center'><font face='Arial, Helvetica, sans-serif' size='2'><strong>Valor Total</strong></font></div></td>"
corpo_email = corpo_email & "</tr>"
i = 0
AbrirConexao()
While i <= Ubound(produtos_selecionados) 
	Set rs = Objconexao.Execute("SELECT tb_produtos.codigo_referencia, tb_produtos.cod_produto, tb_produtos.descricao, tb_produtos.unidade_medida, tb_produtos.data_cadastro, tb_produtos.quantidade_embalagem, tb_produtos.situacao, tb_produtos.cod_grupo, tb_produtos.modificado, tb_grupos.descricao AS grupo_descricao FROM tb_grupos INNER JOIN tb_produtos ON tb_grupos.cod_grupo = tb_produtos.cod_grupo WHERE cod_produto = "&sqlNUmero(produtos_selecionados(i)))
	If not rs.Eof Then
		corpo_email = corpo_email & "<tr>"
			corpo_email = corpo_email & "<td><div align='center'><font face='Arial, Helvetica, sans-serif' size='2'>"&rs("codigo_referencia")&"</font></div></td>"
			corpo_email = corpo_email & "<td><div align='left'><font face='Arial, Helvetica, sans-serif' size='2'>"&rs("descricao")&"</font></div></td>"
			corpo_email = corpo_email & "<td><div align='left'><font face='Arial, Helvetica, sans-serif' size='2'>"&rs("grupo_descricao")&"</font></div></td>"
'			corpo_email = corpo_email & "<td><div align='center'><font face='Arial, Helvetica, sans-serif' size='2'>"&rs("quantidade_embalagem")&"</font></div></td>"
			corpo_email = corpo_email & "<td><div align='center'><font face='Arial, Helvetica, sans-serif' size='2'>"&quantidades_selecionadas(i)&"</font></div></td>"
			corpo_email = corpo_email & "<td><div align='center'><font face='Arial, Helvetica, sans-serif' size='2'>"&rs("unidade_medida")&"</font></div></td>"
			corpo_email = corpo_email & "<td width='9%'><div align='center'><font face='Arial, Helvetica, sans-serif' size='2'>&nbsp;</font></div></td>"
	corpo_email = corpo_email & "</tr>"	
	End if
	rs.close()
	Set rs = nothing
i=i+1
WEnd
FecharConexao()
corpo_email = corpo_email & "</table></center>"
Dim emailEnviado
emailEnviado = EnviaEmailAutenticado(email, empresa&" - "&contato, "luiz@contapar-nante.com.br", "ORÇAMENTO FEITO PELO SITE", corpo_email)
'TESTE DE FUNCIONAMENTO
'emailEnviado = EnviaEmailAutenticado(email, empresa&" - "&contato, "jbeltrame@oxnet.com.br", "ORÇAMENTO FEITO PELO SITE", corpo_email)
%>
<div id="pagina">
    <!-- #include file="includes/topo.asp"-->
    <div id="bannerPrincial">
       	  <div class="conteudo">
       		<div id="bannerPrincialConteudo">
           <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td style="padding-top:20px; padding-left:45px">
      <h1>ORÇAMENTO</h1>
      <h2>SOLICITE UM ORÇAMENTO</h2></td>
    </tr>
</table>
	</div>    
      <!-- #include file="includes/menuprodutos.asp" -->
      </div>
      </div>
        
        <div id="corpointerno">
        	<div class="conteudo">
              <table width="926" border="0" align="center" cellpadding="2" cellspacing="0"  class="vcEstaAqui">
  <tr>
    <td width="7%" style="padding-top:2px"><img src="images/layout/voce_esta_aqui_text.jpg" alt="Você está aqui" /></td>
    <td width="79%">ORÇAMENTO</td>
    <td width="14%" class="alinharDireita"><a href="javascript:voltar()" onmouseout="rollImagem('btn_voltar','images/layout/btn_Voltar.jpg')" onmouseover="rollImagem('btn_voltar','images/layout/btn_VoltarOver.jpg')"><img src="images/layout/btn_Voltar.jpg" alt="Voltar" width="61" height="23" border="0" id="btn_voltar" /></a></td> 
  </tr>
</table>
        <div id="conteudoPaginasInternas">
          <div class="caixasinternas2">
            <div>
<h3>Passo 3 - Finalização</h3>
<table class="caixasinternas3" width="90%" border="0" align="center" cellpadding="10" cellspacing="0">

			        <tr>
			          <td>
                        <div align="center"><strong><br />
                        <%if emailEnviado Then%>
                        Orçamento enviado com sucesso!<br />
                        Um de nossos representantes entrará em contato o mais rápido o possível.<br />
                        <br />
                        <%else%>
                       	Houve um problema no evio de seu orçamento,<br />
						Tente novamente mais tarde<br />
                        <br />
                        <%end if%>
                        </strong>
                      </div></td>

		            </tr>

			        </table><p class="alinharCentro">&nbsp;</p></div>
				</div>
       	  </div>
       	  </div>
        </div> 


		<!-- #include file="includes/rodape.asp" -->
        
</div>
</body>
<!-- #include file="includes/googleAnalytics.asp" -->
</html>