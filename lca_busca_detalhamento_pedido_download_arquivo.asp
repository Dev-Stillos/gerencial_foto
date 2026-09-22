<% dim script_valida_form_master : script_valida_form_master = "NAO" %>
<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<!-- #include file="files/includes/functions_local.asp" -->
<%
if session("cod_usuario") = "" then
	response.end
end if

dim retorno, i
dim t : t = request.querystring("t")
dim p(10)

for i = 0 to 10
	p(i) = request.querystring("p"&i)
next

abrirConexao()
	set rs = objConexao.Execute("SELECT * FROM TB_PEDIDOS_PRODUTOS_ARQUIVOS WHERE id = "&sqlNumero(p(0))   )
	if not rs.Eof then
		Dim path_file 
		path_file = caminho_dados& "pedidos_produtos_anexos\"&rs("nome_arquivo")&"."&rs("extencao")
%>
<!DOCTYPE html>
<html>
<head>
	<meta charset="utf-8">
	<meta http-equiv="X-UA-Compatible" content="IE=edge">
</head>
<body>
	<form action="http://stillos.com.br/gerencial_foto/lca_busca_detalhamento_pedido_download_arquivo.php" method="POST" accept-charset="utf-8" id="downloadFoto">
			<input type="hidden" name="pathFile" value='<%= path_file %>'>
	</form>
	<script src="js/jquery.min.js" type="text/javascript"></script>	
	<script type="text/javascript">
		$(document).ready(function(){
			$("#downloadFoto").trigger("submit");
			console.log('Enviei o formulário');

		});
	</script>	
</body>
</html>
<%
		'donwloadStream(caminho_dados& "pedidos_produtos_anexos\"&rs("nome_arquivo")&"."&rs("extencao"))
	else
		response.write("Arquivo não encontrado! 0x1")
	end if
	rs.close
	set rs = nothing
fecharConexao()

%>