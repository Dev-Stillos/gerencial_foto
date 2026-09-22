<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include virtual="gerencial_foto/files/includes/scripts.asp" -->
<!-- #include file="files/includes/functions_local.asp" -->
<%
function copiaFormandoProducaoPublica(id, conexao_aberta)

copiaFormandoProducaoPublica = ""

if not conexao_aberta then
  abrirConexao()
end if

'BUSCAMOS NA CONFIG
dim configs : set configs = getConfiguracoes(true, array("cod_producao_publica", "cod_contrato_publico"))

'BUSCAMOS INFO
set rs = objConexao.execute("SELECT nome, curso, cod_formando_gerado FROM TB_MACK_FORMANDOS where cod_mack_formando = " & int(id))
if not rs.Eof then

  if IsNull(rs("cod_formando_gerado")) or rs("cod_formando_gerado") = "" then
  
    'GERAMOS O CODIGO DO FORMANDO 
      objConexao.Execute("INSERT INTO TB_FICHAS_IDENTIFICACAO_CONTAGEM_NUMERO_FORMANDO_CADASTRADO (COD_FORMANDO_CADASTRADO)   VALUES (NULL)")
    dim rs1 : Set rs1 = objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM  TB_FICHAS_IDENTIFICACAO_CONTAGEM_NUMERO_FORMANDO_CADASTRADO LIMIT 1")
    if not rs1.Eof then
        cod_formando = sqlNumero( rs1("ULTIMO_ID") & 1)
      End if
    rs1.Close
    Set rs1 = Nothing

    dim cod_producao : cod_producao = int(configs("cod_producao_publica"))
    dim cod_contrato : cod_contrato = cstr(configs("cod_contrato_publico"))

    'INSERIMOS 
    objConexao.execute("INSERT INTO TB_FICHAS_IDENTIFICACAO (COD_FORMANDO, SISTEMA_NOVO, COD_PRODUCAO, COD_CONTRATO, COD_CURSO, NOME_CURSO, NOME_FORMANDO, ANO_CONCLUSAO, SEMESTRE_CONCLUSAO, ALBUM, CPF, FATURADO, COD_MOTIVO_INCLUSAO) VALUES ("&cod_formando&", 1, "&cod_producao&", '"&cod_contrato&"', 0, '"&rs("curso")&"', '"&rs("nome")&"', "&year(now())&", "&month(now())&", 0 , '', 2, 0) ")

    objConexao.execute("UPDATE TB_MACK_FORMANDOS SET cod_formando_gerado = "&int(cod_formando)&" WHERE cod_mack_formando = " & int(id))

    copiaFormandoProducaoPublica = cod_formando
  else 
    copiaFormandoProducaoPublica = rs("cod_formando_gerado")
  end if
end if
rs.close

if not conexao_aberta then
  fecharConexao()
end if

end function


Dim cod_mack_formando

campo1 = request.form("campo1")
comando = request.form("comando")
comando_externo = request.QueryString("comando_externo")
cod_mack_formando = Request.QueryString("cod_mack_formando")

redirecionar_incluir = "mck_formandos_formulario_upload.asp"
redirecionar_alterar = "mck_formandos_formulario_upload.asp"
redirecionar_excluir = "mck_formandos_formulario_upload.asp"

'Por 
if comando = "" and cod_mack_formando <> "" and comando_externo <> "visualizar" and comando_externo <> "excluir" Then
	comando = "incluir"
elseif comando = "" and cod_mack_formando <> "" and comando_externo = "visualizar" or comando_externo = "excluir"  Then
	comando = "visualizar"
else 
  response.end()
end if

'LISTAGEM DE COMANDO
cod_formando_venda = copiaFormandoProducaoPublica(cod_mack_formando, false)


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link rel="stylesheet" type="text/css" href="css/bootstrap.min.css" />
<link type="text/css" rel="stylesheet" href="http://stillos.com.br/gerencial_foto/css/base.css" />
<link rel="stylesheet" media="all" type="text/css" href="http://stillos.com.br/gerencial_foto/css/modal-message.css">
<style>
	fieldset{
		width: 820px;
		margin:0 auto 0 auto;
	}
</style>
</head>
<body style="padding:10px; background-color:#FFF">

<fieldset>
		<legend>
			Cód Formando para lancamento
		</legend>
		<form name='FORMULARIO_PADRAO' action="?comando=<%=comando%>&comando_externo=<%=comando_externo%>" method="post" enctype="multipart/form-data">

    <%if len(cod_formando_venda) > 0 then %>
      <h1 style='text-align:center;'><%=cod_formando_venda%>1</h1>
    <% else %>
      
    <% end if %>
		<input type="hidden" name='cod_mack_formando' value="<%=cod_mack_formando%>" />
		<input type="hidden" name='comando' value="<%=comando%>" />
		<input type="hidden" name='campo1' value='1' />
		</form>
	</fieldset>
<!-- #include virtual="gerencial_foto/files/includes/scripts_rodape.asp" -->
</body>
</html>