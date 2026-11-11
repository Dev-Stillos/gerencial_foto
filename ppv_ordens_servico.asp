<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->

<%
Dim busca_cod_ordem_servico, busca_status_ordem_servico, busca_cod_usuario_responsavel, busca_nome_formando, busca_cod_pedido, busca_prazo_geral_filtro, busca_data_venda_inicial, busca_data_venda_final
Dim busca_diferenca_etapa, busca_cod_formando

Dim sql, sqlWhere, sqlOrderBy, sqlGroupBy

sqlOrderBy = "cod_contrato, nome_formando"
busca_cod_ordem_servico = Cstr(""&Request.QueryString("cod_ordem_servico"))
busca_status_ordem_servico = Cstr(""&Request.QueryString("status_ordem_servico"))
busca_cod_usuario_responsavel = Cstr(""&Request.QueryString("cod_usuario_responsavel"))
busca_diferenca_etapa = Cstr(""&Request.QueryString("diferenca_etapa"))
busca_nome_formando = Cstr(""&Request.QueryString("nome_formando"))
busca_cod_pedido = Cstr(""&Request.QueryString("cod_pedido"))
busca_cod_formando = Cstr(""&Request.QueryString("cod_formando"))
busca_prazo_geral_filtro = Cstr(""&Request.QueryString("prazo_geral_filtro"))
busca_data_venda_inicial = Cstr(""&Request.QueryString("data_venda_inicial"))
busca_data_venda_final = Cstr(""&Request.QueryString("data_venda_final"))



Dim validacao_listar_os_todos_usuarios
validacao_listar_os_todos_usuarios = validaUsuario(2,,57,false)

If Request.QueryString("filtro1") <> Empty Then
		
		'Busca prim�ria
		If busca_cod_ordem_servico <> "" Then
			
			sqlWhere = sqlWhere&" and cod_ordem_servico = "&sqlTexto(moduloPPVOSMascara("",busca_cod_ordem_servico,false))					
			busca_status_ordem_servico = ""
			busca_cod_usuario_responsavel = ""
			busca_diferenca_etapa = ""
			busca_cod_formando = ""
			busca_data_venda_inicial = ""
			busca_data_venda_final = ""
		Else
		
			If busca_status_ordem_servico <> "" Then
				sqlWhere = sqlWhere&" and status_ordem_servico = "&sqlTexto(busca_status_ordem_servico)
				
				If busca_status_ordem_servico = "1" and busca_diferenca_etapa <> "" Then
				'Filtro adicional para OS em andamento
					If busca_diferenca_etapa = "1" Then
					'Somente OSs com etapa atual atrasada
						sqlWhere = sqlWhere&" and diferenca_etapa > 0"
					Elseif  busca_diferenca_etapa = "2" Then
					'Somente OSs com etapa atual vencendo na data atual
						sqlWhere = sqlWhere&" and diferenca_etapa = 0"
					Elseif  busca_diferenca_etapa = "3" Then
					'Somente OSs com etapa atual em dia
						sqlWhere = sqlWhere&" and diferenca_etapa < 0"
					End If
				End If
				
			End If
			
			If busca_cod_usuario_responsavel <> "" Then
				sqlWhere = sqlWhere&" and cod_usuario_responsavel = "&sqlTexto(busca_cod_usuario_responsavel)			
			End If

			If busca_nome_formando <> "" Then
				sqlWhere = sqlWhere&" and nome_formando LIKE '%"&busca_nome_formando&"%'"	
			End If

			If busca_cod_pedido <> "" Then
				sqlWhere = sqlWhere&" and cod_pedido = "&sqlTexto(busca_cod_pedido)
			End If
			
			If busca_prazo_geral_filtro = "1" Then
				sqlWhere = sqlWhere&" and dias_prazo_geral < 0 and dias_prazo_geral IS NOT NULL"
			Elseif busca_prazo_geral_filtro = "2" Then
				sqlWhere = sqlWhere&" and dias_prazo_geral >= 0 and dias_prazo_geral IS NOT NULL"
			End if

			
			If busca_cod_formando <> "" Then
				sqlWhere = sqlWhere&" and cod_formando = "&sqlTexto(Mid(busca_cod_formando,1,len(busca_cod_formando)-1))
			End If
			
			If busca_data_venda_inicial <> "" and busca_data_venda_final <> "" and IsDate(busca_data_venda_inicial) and isDate(busca_data_venda_final) Then
				
				'Data base prazo � a data da venda dos pedidos
				sqlWhere = sqlWhere&" and data_base_prazo_cliente_inicio BETWEEN  "&sqlData(busca_data_venda_inicial)&" AND "&sqlData(busca_data_venda_final)
			
			Else
				busca_data_venda_inicial = ""
				busca_data_venda_final  = ""
			
			End If
			
				
		End If

Else
	'Valores padr�es
	busca_status_ordem_servico  = "1"
	sqlWhere = sqlWhere&" and status_ordem_servico = 1"
End If


	If not validacao_listar_os_todos_usuarios Then
	'N�o possui acesso a listagem de todos usu�rios
		busca_cod_usuario_responsavel = Session("cod_usuario")
		sqlWhere = sqlWhere&" and cod_usuario_responsavel = "&busca_cod_usuario_responsavel
	End If


'/////////////////////
'Padroniza sqlWhere

	If sqlWhere <> "" Then
		'Retirar o "AND "
		sqlWhere = Mid(sqlWhere,5,1000)
	End If


Dim acesso_inclusao_exclusao_os_manual
acesso_inclusao_exclusao_os_manual = validaUsuario(2,,51,false)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Ordens de Servi&ccedil;o | Portal Stillos</title>
<!-- #include file="files/includes/scripts.asp" -->
	<script type="text/javascript" src="js/scrips_calendario/dhtmlSuite-common.js"></script>
<script language="JavaScript" type="text/javascript">
// JavaScript Document

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

//adiciona mascara de cnpj
function MascaraCNPJ(cnpj){
    if(mascaraInteiro(cnpj)==false){
        event.returnValue = false;
    }    
    return formataCampo(cnpj, '00.000.000/0000-00', event);
}

//adiciona mascara de cep
function MascaraCep(cep){
        if(mascaraInteiro(cep)==false){
        event.returnValue = false;
    }    
    return formataCampo(cep, '00.000-000', event);
}

//adiciona mascara de data
function MascaraData(data){
    if(mascaraInteiro(data)==false){
        event.returnValue = false;
    }    
    return formataCampo(data, '00/00/0000', event);
}

//adiciona mascara ao telefone
function MascaraTelefone(tel){    
    if(mascaraInteiro(tel)==false){
        event.returnValue = false;
    }    
    return formataCampo(tel, '(00) 0000-0000', event);
}

//adiciona mascara ao CPF
function MascaraCPF(cpf){
    if(mascaraInteiro(cpf)==false){
        event.returnValue = false;
    }    
    return formataCampo(cpf, '000.000.000-00', event);
}

//valida telefone
function ValidaTelefone(tel){
    exp = /\(\d{2}\)\ \d{4}\-\d{4}/
    if(!exp.test(tel.value))
        alert('Numero de Telefone Invalido!');
}

//valida CEP
function ValidaCep(cep){
    exp = /\d{2}\.\d{3}\-\d{3}/
    if(!exp.test(cep.value))
        alert('Numero de Cep Invalido!');        
}

//valida data
function validarData(campo){
var expReg = /^(([0-2]\d|[3][0-1])\/([0]\d|[1][0-2])\/[1-2][0-9]\d{2})$/;
var msgErro = 'Formato inv&aacute;lido de data.';
if ((campo.value.match(expReg)) && (campo.value!='')){
var dia = campo.value.substring(0,2);
var mes = campo.value.substring(3,5);
var ano = campo.value.substring(6,10);
if(mes==4 || mes==6 || mes==9 || mes==11 && dia > 30){
campo.focus();
alert("Dia incorreto ! O m&ecirc;s especificado cont&eacute;m no m&aacute;ximo 30 dias.");
return false;
} else{
if(ano%4!=0 && mes==2 && dia>28){
campo.focus();
alert("Data incorreta! O m&ecirc;s especificado cont&eacute;m no m&aacute;ximo 28 dias.");
return false;
} else{
if(ano%4==0 && mes==2 && dia>29){
campo.focus();
alert("Data incorreta! O m&ecirc;s especificado cont&eacute;m no m&aacute;ximo 29 dias.");
return false;
} else{ 
return true;
}}}} else {
alert(msgErro);
campo.focus();
return false;
}}


//valida o CPF digitado
function ValidarCPF(Objcpf){
    var cpf = Objcpf.value;
    exp = /\.|\-/g
    cpf = cpf.toString().replace( exp, "" ); 
    var digitoDigitado = eval(cpf.charAt(9)+cpf.charAt(10));
    var soma1=0, soma2=0;
    var vlr =11;
    
    for(i=0;i<9;i++){
        soma1+=eval(cpf.charAt(i)*(vlr-1));
        soma2+=eval(cpf.charAt(i)*vlr);
        vlr--;
    }    
    soma1 = (((soma1*10)%11)==10 ? 0:((soma1*10)%11));
    soma2=(((soma2+(2*soma1))*10)%11);
    
    var digitoGerado=(soma1*10)+soma2;
    if(digitoGerado!=digitoDigitado)    
        alert('CPF Invalido!');        
}

//valida numero inteiro com mascara
function mascaraInteiro(){
    if (event.keyCode < 48 || event.keyCode > 57){
        event.returnValue = false;
        return false;
    }
    return true;
}

//valida o CNPJ digitado
function ValidarCNPJ(ObjCnpj){
    var cnpj = ObjCnpj.value;
    var valida = new Array(6,5,4,3,2,9,8,7,6,5,4,3,2);
    var dig1= new Number;
    var dig2= new Number;
    
    exp = /\.|\-|\//g
    cnpj = cnpj.toString().replace( exp, "" ); 
    var digito = new Number(eval(cnpj.charAt(12)+cnpj.charAt(13)));
        
    for(i = 0; i<valida.length; i++){
        dig1 += (i>0? (cnpj.charAt(i-1)*valida[i]):0);    
        dig2 += cnpj.charAt(i)*valida[i];    
    }
    dig1 = (((dig1%11)<2)? 0:(11-(dig1%11)));
    dig2 = (((dig2%11)<2)? 0:(11-(dig2%11)));
    
    if(((dig1*10)+dig2) != digito)    
        alert('CNPJ Invalido!');
        
}

//formata de forma generica os campos
function formataCampo(campo, Mascara, evento) { 
    var boleanoMascara; 
    
    var Digitato = evento.keyCode;
    exp = /\-|\.|\/|\(|\)| /g
    campoSoNumeros = campo.value.toString().replace( exp, "" ); 
   
    var posicaoCampo = 0;     
    var NovoValorCampo="";
    var TamanhoMascara = campoSoNumeros.length;; 
    
    if (Digitato != 8) { // backspace 
        for(i=0; i<= TamanhoMascara; i++) { 
            boleanoMascara  = ((Mascara.charAt(i) == "-") || (Mascara.charAt(i) == ".")
                                || (Mascara.charAt(i) == "/")) 
            boleanoMascara  = boleanoMascara || ((Mascara.charAt(i) == "(") 
                                || (Mascara.charAt(i) == ")") || (Mascara.charAt(i) == " ")) 
            if (boleanoMascara) { 
                NovoValorCampo += Mascara.charAt(i); 
                  TamanhoMascara++;
            }else { 
                NovoValorCampo += campoSoNumeros.charAt(posicaoCampo); 
                posicaoCampo++; 
              }            
          }     
        campo.value = NovoValorCampo;
          return true; 
    }else { 
        return true; 
    }
}
</script> 

<script language="javascript">
//Formata n�mero tipo moeda usando o evento onKeyDown
function Limpar(valor, validos) {
// retira caracteres invalidos da string
var result = "";
var aux;
for (var i=0; i < valor.length; i++) {
aux = validos.indexOf(valor.substring(i, i+1));
if (aux>=0) {
result += aux;
}
}
return result;
}
function Formata(campo,tammax,teclapres,decimal) {
var tecla = teclapres.keyCode;
vr = Limpar(campo.value,"0123456789");
tam = vr.length;
dec=decimal

if (tam < tammax && tecla != 8){ tam = vr.length + 1 ; }

if (tecla == 8 )
{ tam = tam - 1 ; }

if ( tecla == 8 || tecla >= 48 && tecla <= 57 || tecla >= 96 && tecla <= 105 )
{

if ( tam <= dec )
{ campo.value = vr ; }

if ( (tam > dec) && (tam <= 5) ){
campo.value = vr.substr( 0, tam - 2 ) + "," + vr.substr( tam - dec, tam ) ; }
if ( (tam >= 6) && (tam <= 8) ){
campo.value = vr.substr( 0, tam - 5 ) + "." + vr.substr( tam - 5, 3 ) + "," + vr.substr( tam - dec, tam ) ; 
}
if ( (tam >= 9) && (tam <= 11) ){
campo.value = vr.substr( 0, tam - 8 ) + "." + vr.substr( tam - 8, 3 ) + "." + vr.substr( tam - 5, 3 ) + "," + vr.substr( tam - dec, tam ) ; }
if ( (tam >= 12) && (tam <= 14) ){
campo.value = vr.substr( 0, tam - 11 ) + "." + vr.substr( tam - 11, 3 ) + "." + vr.substr( tam - 8, 3 ) + "." + vr.substr( tam - 5, 3 ) + "," + vr.substr( tam - dec, tam ) ; }
if ( (tam >= 15) && (tam <= 17) ){
campo.value = vr.substr( 0, tam - 14 ) + "." + vr.substr( tam - 14, 3 ) + "." + vr.substr( tam - 11, 3 ) + "." + vr.substr( tam - 8, 3 ) + "." + vr.substr( tam - 5, 3 ) + "," + vr.substr( tam - 2, tam ) ;}
} 

}
</script>
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
	
	<script type="text/javascript">	
	
		function exibeOcultaDinamico(div)
		{
			var o = document.getElementById(div);
			
			
			if(o){
				if (o.style.display == "none" || o.style.display == "")
					{
						o.style.visibility = "visible";
						o.style.display = "inline";
					}
					else
					{
						o.style.visibility = "hidden";
						o.style.display = "none";
					}
			}	
	}
		
	function exibeOculta(div,acao){
	
		var o = document.getElementById(div);
		
		if(o){
			if(acao == "ocultar"){
				//Oculta
					o.style.visibility = "hidden";
					o.style.display = "none";
			}else{
				//Exibe
					o.style.visibility = "visible";
					o.style.display = "inline";		
			}
		}
	}
	function exibirBuscaComplemento(valor){
		if (valor == 1){	
			exibeOculta('comboComplementoBusca','exibir'); 
		}else{
			exibeOculta('comboComplementoBusca','ocultar'); 
		}
	}	
</script>
<style type="text/css">
<!--
.container_Faltam{
	display:block;
	width:80px;
	color:#000;
	font-weight:bold;
	padding:3px;
	border:solid 1px #FFF;
}
.container_Vencido{
	display:block;
	width:80px;
	background-color:#C00;
	color:#FFF;
	padding:3px;
	font-weight:bold;
	border:solid 1px #FFF;
}
.container_Vencido:hover,.container_Faltam:hover{
	cursor:pointer;
	border:solid 1px #CCC;
} 
-->
</style>    
<link rel="stylesheet" href="css/vendor/bootstrap.min.css" />
<link rel="stylesheet" href="css/dashboard.css" />
<style>
.ppv-page .ppv-main { max-width: 1680px; }
.ppv-page .ppv-tools { margin: 0 0 24px; padding: 12px 20px; background: #fff; border: 1px solid var(--line); border-radius: 12px; }
.ppv-page .ppv-tools table { max-width: 100%; }
.ppv-page .ppv-module { margin-bottom: 22px; }
.ppv-page .subMenu { background: #fff; border: 1px solid var(--line); border-radius: 12px; padding: 12px; line-height: 2.6; }
.ppv-page .subMenu a { display: inline-block; color: #53566b; text-decoration: none; padding: 0 8px; }
.ppv-page .subMenu a[href="ppv_ordens_servico.asp"] { background: #f9eaf2; color: #9e125c; border-radius: 6px; font-weight: 700; }
.ppv-page .ppv-filters { height: auto; margin-bottom: 24px; }
.ppv-page .ppv-fields { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 16px; margin: 18px 0; align-items: end; }
.ppv-page .ppv-fields label { display: block; font-weight: 600; margin-bottom: 6px; color: #44485b; }
.ppv-page .ppv-fields input, .ppv-page .ppv-fields select { width: 100%; min-width: 0; height: auto; }
.ppv-page .ppv-dates { display: flex; align-items: center; gap: 8px; }
.ppv-page .ppv-actions { display: flex; flex-wrap: wrap; gap: 8px; }
.ppv-page .ocultarCompleto { display: none; }
.ppv-page .ppv-results { background: #fff; border: 1px solid var(--line); border-radius: 12px; margin: 0 0 20px; overflow: hidden; }
.ppv-page .ppv-result-heading { display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px; padding: 18px 20px; }
.ppv-page .ppv-scroll { overflow-x: auto; }
.ppv-page .ppv-list { min-width: 1050px; width: 100%; border-collapse: collapse; margin: 0; }
.ppv-page .ppv-list > thead > tr > th { background: #f3f2f8; color: #42465a; font-weight: 650; }
.ppv-page .ppv-list > thead > tr > th, .ppv-page .ppv-list > tbody > tr > td { padding: 10px; vertical-align: middle; border: 0; border-bottom: 1px solid #ececf2; font-size: 12px; }
.ppv-page .ppv-list > tbody > tr.ppv-group > td { background: #f8f6fa; }
.ppv-page .ppv-list > tbody > tr:not(.ppv-group):hover > td { background: #faf9fc; }
.ppv-page .ppv-status, .ppv-page .container_Faltam, .ppv-page .container_Vencido { display: inline-block; width: auto; border: 0; border-radius: 6px; padding: 4px 7px; font-weight: 650; }
.ppv-page .ppv-status, .ppv-page .container_Faltam { background: #f1eff8; color: #494168; }
.ppv-page .container_Vencido { background: #fbe7ed; color: #9b2044; }
.ppv-page .ppv-guide { margin-top: 20px; height: auto; }
.ppv-page .ppv-guide summary { font-weight: 650; cursor: pointer; }
.ppv-page .ppv-guide table { width: auto; max-width: 100%; }
.ppv-page .dashboard-nav a.active { background: transparent; color: #d4d5e1; font-weight: 400; }
.ppv-page .dashboard-nav a[href="ppv_ordens_servico.asp"] { background: #b20b61; color: #fff; font-weight: 600; }
@media(max-width: 1250px) { .ppv-page .ppv-fields { grid-template-columns: repeat(2, minmax(0, 1fr)); } }
@media(max-width: 540px) { .ppv-page .ppv-fields { grid-template-columns: minmax(0, 1fr); } }
</style>
</head>
<body class="ppv-page">
<a class="skip-link" href="#conteudo">Ir para o conte&uacute;do</a>
<div class="dashboard-layout">
  <aside class="dashboard-sidebar" id="menuPrincipal">
    <a class="dashboard-brand" href="adm_index.asp"><img src="images/logoparafundopreto.png" alt="Stillos" width="180" /><span>GERENCIAL FOTOS</span></a>
    <div class="nav-caption">&Aacute;REA DE TRABALHO</div>
    <!-- #include file="includes/dashboard_menu.asp" -->
  </aside>
  <div class="dashboard-workspace">
    <header class="dashboard-topbar">
      <button class="btn menu-toggle" id="ppvMenuToggle" type="button" aria-controls="menuPrincipal" aria-expanded="false">Menu</button>
      <div class="topbar-title">Portal Stillos <span>/ Produ&ccedil;&atilde;o p&oacute;s-venda</span></div>
      <div class="topbar-actions"><a href="manual_sistema.asp" target="_blank" rel="noopener">Manual</a><a class="btn btn-outline-secondary btn-sm" href="adm_sair.asp">Sair</a></div>
    </header>
    <main class="dashboard-main ppv-main" id="conteudo">
      <div class="dashboard-heading">
        <div><p class="eyebrow">PRODU&Ccedil;&Atilde;O P&Oacute;S-VENDA</p><h1>Ordens de Servi&ccedil;o</h1><p>Acompanhe a produ&ccedil;&atilde;o, os respons&aacute;veis e os prazos de cada etapa.</p></div>
        <a href="adm_index.asp" class="btn btn-outline-secondary">Vis&atilde;o geral</a>
      </div>
      <div class="search-panel"><div><strong>Busca no portal</strong></div><!-- #include file="includes/dashboard_busca.asp" --></div>
      <div class="ppv-tools">
      <%
      Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
      BTN_VOLTAR = 1
      BTN_IMPRIMIR = 2
      BTN_FONTE = 1
      BTN_IMPRIMIR_LINK = "#"
      %>
      <!-- #include file="includes/sub_menu.asp" -->
      </div>
      <nav class="ppv-module" aria-label="Produ&ccedil;&atilde;o p&oacute;s-venda">
      <!-- #include file="includes/classe_modulo_ppv_menu.asp" -->
      </nav>
 <%
	Dim sucesso, codmsg
	codmsg	= CStr(""&Request.QueryString("codmsg"))
	sucesso = CStr(""&Request.QueryString("sucesso"))
		
	If sucesso = "" and codmsg <> "" Then%>
    	<div align="center" class="msgGrandeSucesso">
        	<%
			If codmsg = "1" Then
				Response.Write("Ordem de Servi&ccedil;o Manual inserida com sucesso!")
			Elseif codmsg = "2" Then
				Response.Write("Ordem de Servi&ccedil;o alterada com sucesso!")
			Elseif codmsg = "3" Then	
				Response.Write("Ordem de Servi&ccedil;o  exclu&iacute;da com sucesso!")
			Elseif codmsg = "101" Then	
				Response.Write("Responsabilidade da Ordem de Servi&ccedil;o alterada com sucesso!")
			Elseif codmsg = "102" Then	
				Response.Write("Ordem de Servi&ccedil;o Finalizada com sucesso!")
			Elseif codmsg = "103" Then	
				Response.Write("Ordem de Servi&ccedil;o com Etapa iniciada com sucesso!")
			Elseif codmsg = "104" Then	
				Response.Write("Ordem de Servi&ccedil;o Cancelada com sucesso!")
			Elseif codmsg = "105" Then	
				Response.Write("Etapa finaliza com sucesso!")
			Elseif codmsg = "106" Then	
				Response.Write("Etapa retornada com sucesso!")
			Elseif codmsg = "201" Then	
				Response.Write("Triagem inicial realizada com sucesso!")
			Elseif codmsg = "301" Then	
				Response.Write("Despacho realizado realizada com sucesso!<br /><a href='ppv_ordens_servico_triagem_despacho_imprimir.asp' target='_blank' class='linkMaster'>Imprimir Comprovantes</a>")
			End If	
			%>
            
        
        </div>
    <%ElseIf sucesso <> "" and codmsg <> "" Then%>
        <div align="center" class="msgGrandeErro">
			<%
			If codmsg = "4" Then
				Response.Write("Vendedor n&atilde;o alterado para n&atilde;o haver duplicidade de informa&ccedil;&atilde;o!<BR>O conte&uacute;do do campo """&Request.QueryString("chaveUnica")&""" j&aacute; possui uma informa&ccedil;&atilde;o igual em outro Vendedor.<br />")
			Elseif codmsg = "5" Then
				Response.Write("Vendedor n&atilde;o exclu&iacute;do, pois h&aacute; pedidos que dependem deste Vendedor.<BR>Se n&atilde;o desejar ter mais este Vendedor inative ao inv&eacute;s de excluir.<br />")
				
			'////
			'Usu�rios
			ElseIf codmsg = "44" Then
				Response.Write("Usu&aacute;rio n&atilde;o alterado para n&atilde;o haver duplicidade de informa&ccedil;&atilde;o!<BR>O conte&uacute;do do campo """&Request.QueryString("chaveUnica")&""" j&aacute; possui uma informa&ccedil;&atilde;o igual em outro Usu&aacute;rio.<br />")
				
			ElseIf codmsg = "701" Then
				Response.Write("Voc&ecirc; n&atilde;o digitou o c&oavute;d. formando ficha..<br />O C&oacute;digo formando ficha &eacute; o c&oacute;digo do formando com o final 1.")
			
			End If	
			%>
</div>
	<%End If

	abrirConexao()%>

      <section class="dashboard-panel ppv-filters" aria-labelledby="ppvFiltros">
        <h2 id="ppvFiltros">Filtrar ordens</h2>
        <p class="panel-description">Combine os filtros para encontrar as ordens que precisa acompanhar.</p>
        <form action="ppv_ordens_servico.asp" method="get">
          <input type="hidden" name="filtro1" value="1" />
          <div class="ppv-fields">
        <div><label for="ppv_status_ordem_servico">Status da OS</label>
<select id="ppv_status_ordem_servico" name="status_ordem_servico" onchange="exibirBuscaComplemento(this.value)" class="form-select">
		        <option value="" <%If busca_status_ordem_servico = "" Then%> selected="selected" <%End If%>>Todos</option>
		        <option value="1" <%If busca_status_ordem_servico = "1" Then%> selected="selected" <%End If%>>1 - Em andamento</option>
		        <option value="2" <%If busca_status_ordem_servico = "2" Then%> selected="selected" <%End If%>>2 - Finalizados e n&atilde;o despachados</option>
		        <option value="3" <%If busca_status_ordem_servico = "3" Then%> selected="selected" <%End If%>>3 - Despachados para o cliente</option>
            <option value="4" <%If busca_status_ordem_servico = "4" Then%> selected="selected" <%End If%>>4 - Entregue para o cliente</option>
		        <option value="99" <%If busca_status_ordem_servico = "99" Then%> selected="selected" <%End If%>>99 - Cancelados</option>
		       </select></div>
        <div><label for="ppv_prazo_geral_filtro">Prazo geral</label>
<select id="ppv_prazo_geral_filtro" name="prazo_geral_filtro" class="form-select">
		        <option value="" <%If busca_prazo_geral_filtro = "" Then%> selected="selected" <%End If%>>Todos</option>
		        <option value="1" <%If busca_prazo_geral_filtro = "1" Then%> selected="selected" <%End If%>>1 - Prazo geral Vencido</option>
		        <option value="2" <%If busca_prazo_geral_filtro = "2" Then%> selected="selected" <%End If%>>2 - Prazo geral N&atilde;o vencido</option>
		       </select></div>
        <div><label for="ppv_cod_usuario_responsavel">Respons&aacute;vel pela etapa atual</label>
<select id="ppv_cod_usuario_responsavel" name="cod_usuario_responsavel" class="form-select" <%If not validacao_listar_os_todos_usuarios Then%> disabled="disabled" <%End If%>>
		        <option value="" <%If Cstr(""&busca_cod_usuario_responsavel) = "" Then%> selected="selected" <%End If%>>Todos</option>
		        <%
				Set rs=objConexao.Execute("SELECT cod_usuario, nome FROM TB_ADM_USUARIOS WHERE BLOQUEADO = 2 AND COD_USUARIO <> 1 ORDER BY NOME")
					While not rs.Eof%>
		        <option  <%If Cstr(""&busca_cod_usuario_responsavel) = Cstr(rs("cod_usuario")) Then%> selected="selected" <%End If%> value="<%=rs("cod_usuario")%>"><%=Ucase(rs("nome"))%></option>
		        <%
					rs.MoveNext()
					Wend
				
				rs.Close
				Set rs=nothing
				%>
		        </select></div>
        <div id="comboComplementoBusca" <%If busca_status_ordem_servico <> "1" Then%>class="ocultarCompleto"<%End If%>>
          <label for="ppv_diferenca_etapa">Prazo da etapa atual</label>
<select id="ppv_diferenca_etapa" name="diferenca_etapa" class="form-select">
                <option value="" <%If busca_diferenca_etapa = "" Then%> selected="selected" <%End If%>>Todas</option>
                <option value="1" <%If busca_diferenca_etapa = "1" Then%> selected="selected" <%End If%>>Somente OSs com etapa atual atrasada</option>
                <option value="2" <%If busca_diferenca_etapa = "2" Then%> selected="selected" <%End If%>>Somente OSs com etapa atual vencendo na data atual</option>
                <option value="3" <%If busca_diferenca_etapa = "3" Then%> selected="selected" <%End If%>>Somente OSs com etapa atual em dia</option>
              </select>
        </div>
        <div><label for="ppv_nome_formando">Nome do formando</label>
<input id="ppv_nome_formando" type="text" name="nome_formando" value="<%=Server.HTMLEncode(busca_nome_formando)%>" class="form-control" /></div>
        <div><label for="ppv_cod_formando">C&oacute;d. do formando (ficha)</label>
<input id="ppv_cod_formando" type="text" name="cod_formando" value="<%=Server.HTMLEncode(busca_cod_formando)%>" class="form-control" /></div>
        <div><label for="ppv_cod_pedido">C&oacute;d. do pedido</label>
<input id="ppv_cod_pedido" type="text" name="cod_pedido" value="<%=Server.HTMLEncode(busca_cod_pedido)%>" class="form-control" /></div>
        <div><label for="ppv_DATA_VENDA_INICIAL">Data da venda</label><div class="ppv-dates">
          <input id="ppv_DATA_VENDA_INICIAL" type="text" name="DATA_VENDA_INICIAL" value="<%=Server.HTMLEncode(busca_data_venda_inicial)%>" class="form-control campoData" placeholder="dd/mm/aaaa" aria-label="Data inicial da venda" />
          <span>a</span>
          <input id="ppv_DATA_VENDA_FINAL" type="text" name="DATA_VENDA_FINAL" value="<%=Server.HTMLEncode(busca_data_venda_final)%>" class="form-control campoData" placeholder="dd/mm/aaaa" aria-label="Data final da venda" />
        </div></div>
          </div>
          <div class="ppv-actions">
            <button class="btn btn-primary" type="submit">Aplicar filtros</button>
            <a class="btn btn-outline-secondary" href="ppv_ordens_servico.asp">Limpar filtros</a>
          </div>
        </form>
      </section>
      <p class="panel-description mb-3">Ordenado pelo c&oacute;digo do contrato e pelo nome do formando.
      <%If Not validacao_listar_os_todos_usuarios Then%>
        A lista mostra ordens cuja etapa atual ou final est&aacute; sob sua responsabilidade.
      <%End If%></p>
        <%
		Dim temp
		'Neste momento corrige o sqlWhere somente para o primeiro rs
		If sqlWhere <> "" Then
			temp = "WHERE "
		End If
		
		sql = "SELECT cod_pedido FROM VW_TB_ORDENS_SERVICO "&temp&sqlWhere&" GROUP BY cod_pedido ORDER BY "&sqlOrderBy&" LIMIT 1000"
		


		Set rs=server.CreateObject("ADODB.recordset")
		rs.CursorLocation=3  'Cursor no cliente
		rs.PageSize=20
		rs.Open sql, objConexao
		
		Dim paginacaocao
		paginacaocao = formatNumber(0&Request.QueryString("paginacao"),0)
		
		If not rs.Eof Then
		
			If paginacaocao > 0 Then
				rs.AbsolutePage=paginacaocao
			Else
				paginacaocao = 1
				rs.AbsolutePage=1
			End If	
			
			
			'Verifica a qtd de OS de acordo com o filtro
			Set rs2=objConexao.Execute("SELECT COUNT(*) as qtd FROM VW_TB_ORDENS_SERVICO  "&temp&sqlWhere&" LIMIT 1000")
				If not rs2.Eof Then
					qtdOSs = rs2(0)
				End If
			rs2.Close
			Set rs2=Nothing
		%>
      <section class="ppv-results" aria-labelledby="ppvResultado">
        <div class="ppv-result-heading">
          <h2 id="ppvResultado">Ordens encontradas <span class="subtle-badge"><%=qtdOSs%> OS</span></h2>
          <span class="panel-description">P&aacute;gina <strong><%=rs.AbsolutePage%></strong> de <%=rs.PageCount%></span>
        </div>
        <div class="ppv-scroll" tabindex="0" role="region" aria-label="Listagem de ordens de servi&ccedil;o">
        <table class="ppv-list">
        <thead>
  <tr bgcolor="whitesmoke">
  <th colspan="2" rowspan="2"><div align="center"><strong> OS</strong></div></th>
  <th width="7%" rowspan="2"><div align="center"><strong>Produto</strong></div></th>
  <th width="4%" rowspan="2"><div align="center"><strong>Qtd.</strong></div></th>
  <th colspan="5"><div align="center" style="font-weight: bold">
    <div align="center">Etapa atual</div>
  </div></th>
  <th width="10%" rowspan="2"><div align="center"><strong>Data Cadastro</strong></div></th>
  <th width="8%" rowspan="2"><div align="center"><strong>Prazo Geral</strong></div></th>
  <th width="6%" rowspan="2"><div align="center"><strong>Status</strong></div></th>
  <th width="8%" rowspan="2">A&ccedil;&otilde;es</th>
  <th width="6%" rowspan="2">Sinaliza&ccedil;&atilde;o</th>
</tr>
  <tr bgcolor="whitesmoke">
    <th width="6%"><div align="center"><strong>Etapa / Total</strong></div></th>
    <th width="10%"><div align="center"><strong>Descri&ccedil;&atilde;o</strong></div></th>
    <th width="11%"><div align="center"><strong>Resp. Atual</strong></div></th>
    <th width="5%"><div align="center"><strong>Data In&iacute;cio</strong></div></th>
    <th width="7%"><div align="center"><strong>Prazo etapa</strong></div></th>
    </tr>
</thead><tbody>
<%

'Neste momento que j� passou o group by corrige o sqlWhere para o rs2
If sqlWhere <> "" Then
  	sqlWhere = " AND "&sqlWhere
End If

Dim numero_etapa, descricao, cod_usuario_responsavel, nome_usuario_responsavel, prazo_em_dias, data_inicio, prazo_limite, diferenca, qtd_etapas
Dim x
For x=1 To rs.PageSize
If Not rs.Eof Then
%>
    
<tr class="ppv-group">
							<td colspan="14"><table width="650" border="0" cellspacing="0" cellpadding="1">
                                  <tr>
                                    <td class="textoMedio"><strong>Pedido #<%=rs("cod_pedido")%></strong></td>
                                  </tr>
                                </table></td>
						  </tr>
<%
sql = 	" SELECT "&_
		" TB_PEDIDOS_FORMANDOS.COD_FORMANDO, "&_
		" TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, "&_
		" TB_FICHAS_IDENTIFICACAO.COD_CONTRATO, "&_
		" TB_ENTIDADES.DESCRICAO, "&_
		" TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO, "&_
		" TB_FICHAS_IDENTIFICACAO.NOME_CURSO "&_
		" FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON "&_
		" TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO "&_
		" INNER JOIN TB_ENTIDADES ON(TB_FICHAS_IDENTIFICACAO.COD_CONTRATO = TB_ENTIDADES.COD_CONTRATO)"&_
		" WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&rs("cod_pedido")&" ORDER BY NOME_FORMANDO"
											
		Set rs2=objConexao.Execute(sql)
		
			While not rs2.Eof
			
				sql = "SELECT * FROM VW_TB_ORDENS_SERVICO WHERE cod_formando = "&rs2("cod_formando")&" and cod_pedido = "&rs("cod_pedido")&sqlWhere&" order by numero_ordem_servico desc, numero_ordem_servico_indice asc"
			   'Response.Write(sql)
			   Set rs3=objConexao.Execute(sql)
			   	
				If not rs3.Eof Then
		%>
<tr class="ppv-group">
							<td colspan="14"><table width="650" border="0" cellspacing="0" cellpadding="1">
                                  <tr>
                                    <td width="21" class="textoMedio">&nbsp;</td>
                                    <td width="625">
                                   
				  <table cellpadding="0" width="580" cellspacing="0">	
           												
                                                  
                                                            	<tr>
                                                                <td width="36%"><%=rs2("NOME_FORMANDO")%></td>
                                                                <td width="15%"><%=rs2("COD_FORMANDO")%>-1</td>
                                                                <td width="64%"><%=rs2("DESCRICAO")%></td>
                                                                
                                                          
                                                            </table></td>
                                  </tr>
                                </table></td>
						  </tr>
  
<%
	   
	   While not rs3.Eof
	   
	   			
				Set rs4=objConexao.Execute("Select count(*) as qtd_etapas From TB_ORDENS_SERVICO_ETAPAS where cod_ordem_servico = "&rs3("cod_ordem_servico")&" group by cod_ordem_servico")
					
					If not rs4.Eof Then
						qtd_etapas = formatNumber(0&rs4("qtd_etapas"),0)
					End If
					
				rs4.Close
				Set rs4=Nothing  
                %>        
				<tr>
														<td width="4%"><div align="right"><img src="images/icons/seta_superior.gif" width="19" height="15" /></div></td>
											<td width="8%"><div align="center"><%=moduloPPVOSMascara(rs3("cod_ordem_servico"),"",true)%></div></td>
														<td><div align="left">
														<%=rs3("nome_produto")%>                                        
														</div></td>
														<td><div align="center"><%=rs3("quantidade")%></div></td>
														<td><div align="center"><%=rs3("numero_etapa")%> / <%=qtd_etapas%></div></td>
														<td><div align="left"><%=rs3("descricao")%></div></td>
														<td><div align="left"><%=nomeUsuario(rs3("cod_usuario_responsavel"),true)%></div></td>
														<td><div align="center"><%=ArrumaData(rs3("data_inicio"))%></div></td>
														<td><div align="center"><%=ArrumaData(rs3("prazo_limite"))%></div></td>
														<td><div align="center"><%=ArrumaData(rs3("data_cadastrado"))%></div></td>
														<td><div align="center">
														<%
														temp = Cstr(""&rs3("dias_prazo_geral"))
														temp2 = ""
														If temp <> "" Then
															temp = formatNumber(temp)
															
															If temp > 10 Then
																temp2 = "Faltam"
															Else if temp >=0 and temp <=10 Then
																temp2 = "Faltam"
															Else
																temp2 = "Vencido"
																temp = temp*-1
															End If
															%>
															
															<%End If%>
															<div class="container_<%=temp2%>" onclick="alert('<%=temp2%>:  <%=int(temp)%> dias');"><%=ArrumaData(rs3("prazo_geral"))%></div>
														<%Else%>
														-
														<%End If%>
														</div></td>
														<td><div align="center"><span class="ppv-status"><%
                    Select Case CStr(rs3("status_ordem_servico"))
                        Case "1": Response.Write "Em andamento"
                        Case "2": Response.Write "Aguardando despacho"
                        Case "3": Response.Write "Despachada"
                        Case "4": Response.Write "Entregue"
                        Case "99": Response.Write "Cancelada"
                        Case Else: Response.Write Server.HTMLEncode(rs3("status_ordem_servico") & "")
                    End Select
                    %></span></div></td>
														<td><div align="center">
																
																<table border="0" cellspacing="0" cellpadding="2">
				  <tr>
					<td><div align="center"><a href="ppv_ordens_servico_formulario.asp?comando=visualizar&cod_ordem_servico=<%=rs3("cod_ordem_servico")%>"><img src="images/icons/icon_visualizar.png" alt="Visualizar OS" title="Visualizar OS" width="16" height="16" border="0" /></a></div></td>
					<%If acesso_inclusao_exclusao_os_manual and int(rs3("status_ordem_servico")) = 1 and rs3("motivo_gerado_manualmente") <> Empty Then%>
					<td><a href="ppv_ordens_servico_formulario.asp?comando=alterar&cod_ordem_servico=<%=rs3("cod_ordem_servico")%>"><img src="images/icons/icon_alterar.png" alt="Alterar OS" title="Alterar OS" width="16" height="16" border="0" /></a></td>
					<td><a href="ppv_ordens_servico_formulario.asp?comando=excluir&cod_ordem_servico=<%=rs3("cod_ordem_servico")%>"><img src="images/icons/icon_excluir.png" alt="Excluir OS" title="Excluir OS" width="16" height="16" border="0" /></a></td>
					<%End If%>
				  </tr>                   
					                  
										  
							
				</table>
				
																	</div></td>
														<td>
														<div align="center">
															<%If int(rs3("status_ordem_servico")) = 4 Then%>
																<img src="images/icons/icon_green.gif" width="16" height="15" />
                              <%ElseIf int(rs3("status_ordem_servico")) = 3 Then%>
                                <img src="images/icons/icon_entrega.png" width="16" height="15" />
				                      <%ElseIf int(rs3("status_ordem_servico")) = 2 Then%> 
																<img src="images/icons/icon_aguardando.png" width='16px' height='16px' />
															<%Elseif  int(rs3("status_ordem_servico")) = 99  Then%>
																<img src="images/icons/icon_red.gif" width="17" height="15" />
															<%Else%>
														  <%
																
																'Response.Write(rs2("diferenca_etapa"))
																diferenca = formatNumber(rs3("diferenca_etapa"))/1
																'Response.Write(diferenca)
																If diferenca > 0 Then%>
														  <img src="images/icons/icon_balao_red.gif" alt="Etapa atrasada h&aacute; <%=diferenca%> dia(s)." title="Etapa atrasada h&aacute; <%=diferenca%> dia(s)." />
														  <%Elseif diferenca = 0 Then%>
														  <img src="images/icons/icon_balao_yellon.gif" alt="OS vencendo hoje." title="OS vencendo hoje" />
																<%Else%>
														  <img src="images/icons/icon_balao_green.gif" alt="Faltam <%=(diferenca)*(-1)%> dia(s) para t&eacute;rmino do prazo da etapa." title="Faltam <%=(diferenca)*(-1)%> dia(s) para t&eacute;rmino do prazo da etapa." />                                                
																<%End If%>
															<%End If%>
														</div>                                        </td>
											</tr>
                
       <% 
			
			rs3.MoveNext
			Wend			   
			
			End If
											
			rs3.Close
			Set rs3=Nothing
			   
	   rs2.MoveNext()
	   Wend
	   
	   rs2.Close
	   Set rs2=Nothing
	   %>
                
<%
rs.MoveNext()
End If
Next
%>
<%If rs.PageCount > 1 Then %>                                  
     <tr bgcolor="whitesmoke">
      <td colspan="14">
      	   <%call paginacao(paginacaocao,rs.PageCount)%>      </td>
    </tr>   
<%End if%>
        </tbody></table>
        </div></section>

<%
	Else
	%>
    <div class="alert alert-light border" role="status">Nenhuma ordem de servi&ccedil;o encontrada para os filtros selecionados.</div>
    <%End if
	
	rs.Close
	Set rs=Nothing
	
	FecharConexao()
	%>
      <details class="dashboard-panel ppv-guide"><summary>Legendas e orienta&ccedil;&otilde;es</summary><div class="mt-3">
<strong>Legenda de Status:</strong><br />
       1 - Em produ&ccedil;&atilde;o<br />
       2 - Finalizados e n&atilde;o despachados<br />
       3 - Despachado para cliente<br />
       4 - Entregue para o cliente<br />
       99 - Cancelado
<p align="left"><strong>Legenda de &Iacute;cones:</strong><br />
</p>
<table width="541" border="0" cellspacing="0" cellpadding="2">
  <tr>
    <td><div align="center"><img src="images/icons/icon_balao_green.gif" width="17" height="16" /></div></td>
    <td>Produto com etapa dentro do prazo (passe o mouse para saber qual o prazo que falta para etapa)</td>
  </tr>
  <tr>
    <td><div align="center"><img src="images/icons/icon_balao_yellon.gif" /></div></td>
    <td>Produto com etapa vencendo na data atual</td>
  </tr>
  <tr>
    <td width="17"><div align="center"><img src="images/icons/icon_balao_red.gif" width="17" height="16" /></div></td>
    <td width="516">Produto com etapa fora do prazo  (passe o mouse para saber quanto esta atrasado)</td>
  </tr>
  
  <tr>
    <td><img src="images/icons/icon_red.gif" width="17" height="15" /></td>
    <td>Produto cancelado</td>
  </tr>
  <tr>
    <td><div align="center"><img src="images/icons/icon_aguardando.png" width='16px' height='16px' /></div></td>
    <td><div align="left">Finalizado e aguardando despacho</div></td>
  </tr>
  <tr>
    <td><div align="center"><img src="images/icons/icon_entrega.png" /></div></td>
    <td><div align="left">Despachado para o cliente</div></td>
  </tr>
  <tr>
    <td><div align="center"><img src="images/icons/icon_green.gif" width="16" height="15" /></div></td>
    <td><div align="left">Produto finalizado e Entregue para cliente</div></td>
  </tr>
</table>
<p align="left"><strong>Observa&ccedil;&otilde;es:<br />
  </strong>1) &Eacute; poss&iacute;vel alterar a OS somente quem &eacute; o respons&aacute;vel atual pela OS<br />
2) &Eacute; poss&iacute;vel excluir a OS somente se a OS foi lan&ccedil;ada manualmente e n&atilde;o foi exclu&iacute;da, se n&atilde;o desejar ter mais a OS utilize o bot&atilde;o  Cancelar dispon&iacute;vel no modo de visualiza&ccedil;&atilde;o da OS<br />
3) Para trocar a responsabilidade da etapa utilize o modo de visualiza&ccedil;&atilde;o e em seguida o bot&atilde;o Trocar Responsabilidade (necess&aacute;rio valida&ccedil;&atilde;o especial)<br />
</p>
      </div></details>
    </main>
    <footer class="dashboard-footer">&copy; <%=Year(Date)%> Grupo Stillos <span>Sistema administrativo interno</span></footer>
  </div>
</div>
<script>
(function () {
    var toggle = document.getElementById('ppvMenuToggle');
    var menu = document.getElementById('menuPrincipal');
    var previous = menu.querySelector('[aria-current="page"]');
    var current = menu.querySelector('a[href="ppv_ordens_servico.asp"]');
    if (previous) previous.removeAttribute('aria-current');
    if (current) current.setAttribute('aria-current', 'page');
    toggle.addEventListener('click', function () {
        var expanded = menu.classList.toggle('is-open');
        toggle.setAttribute('aria-expanded', String(expanded));
    });
    document.addEventListener('keydown', function (event) {
        if (event.key === 'Escape' && menu.classList.contains('is-open')) {
            menu.classList.remove('is-open');
            toggle.setAttribute('aria-expanded', 'false');
            toggle.focus();
        }
    });
})();
</script>
</body>
</html>
