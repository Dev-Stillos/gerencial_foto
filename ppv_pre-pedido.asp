<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!--
'*** Instruções
'1) 1 Grava na temporária os códigos informados e abaixo informa de quem corresponde
'2) Após informar todos clica em avançar para concluir devolução

-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
	<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>
	<script language="javascript">
		function mOvr(src, cor)
			{
			if (!src.contains(event.fromElement))
					{
						src.style.cursor = 'hand';
						src.bgColor = cor;
					}
			}
		
		function mOut(src, cor)
			{
				if (!src.contains(event.toElement))
					{
						src.style.cursor = 'default';
						src.bgColor = cor;
			
					}
			}
	</script>
	
	<script type="text/javascript" language="javascript">
<!--
<!--AJUDA F1 E CTRL PARA BUSCA RÁPIDA-->
		
	function abrir(theURL,winName,features) 
	{ 
	  window.open(theURL,winName,features);
	}

		
	function abre(url)
	{    
				//if(event.srcElement.tagName=='TD')
				//	{
				//		src.children.tags('A')[0].click(); 
				//	}
	location.href = url;

	}	
		
		
	

	function buscaIni(){
	d = document.buscaForm
	d.busca.focus()
	d.busca.select();
	}
	function mudacorBusca(CAMPO){
	CAMPO.style.backgroundColor = '#F5ECBA'
	}
	function mudacorBusca2(CAMPO){
	CAMPO.style.backgroundColor = '#FFFFFF'
	}

 function Verificar()
 {
  var tecla=window.event.keyCode;
  if (tecla==17) {
    buscaIni();
	}
 }
 


	<!--AJUDA F1 E CTRL PARA BUSCA RÁPIDA-->

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->

function cadastraVendedor(valor){
	alert(valor);
}
</script>
</head>
<script language="javascript">
function focoCx(campo){
	campo.style.backgroundColor = '#FFFF66'
}
function blurCx(campo){
	campo.style.backgroundColor = '';
}
function carregaCampo(){
document.DEVOLUCAO.TXT_SAIDA.focus();
}
</script>
<script>
function confirmaDevolucao() {
	var msg
	
	d = document.DEVOLUCAO_ITENS
	if (d.COD_VENDEDOR.value == ""){
		d.COD_VENDEDOR.focus();
		alert('Escolha o vendedor que está retirando os produtos!');
		return false;
	}else{
		msg = "Tem certeza que deseja gravar a saída destes itens do estoque para o vendedor selecionado ao lado?"
		if(YNconfirm(msg)){
				return true;
				//document.DEVOLUCAO.submit();			
		}else{
				return false;
		}
	
	}	
}
</script>
<script language="javascript">
function YNconfirm(txt){
return typeof(suporteVBscript)=="undefined"?confirm(txt):VBconfirm(txt)==6
}
</script>
<script language="vbscript">
suporteVBscript=1
function VBconfirm(mensagem)

 VBconfirm=msgbox(mensagem,vbyesno)
end function

</script>
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

													<h2>PPV - Pre Pedido</h2>
													<div class="col-md-12">
														<div class="embed-responsive container-fluid">
														  	<iframe class="embed-responsive-item" src="pre-pedido.php" style="width: 99.5%;"></iframe>
														</div>
													</div>
													<p align="center">&nbsp;</p>		
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