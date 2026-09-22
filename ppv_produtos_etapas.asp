<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>Produtos Etapas</title>
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
var msgErro = 'Formato inválido de data.';
if ((campo.value.match(expReg)) && (campo.value!='')){
var dia = campo.value.substring(0,2);
var mes = campo.value.substring(3,5);
var ano = campo.value.substring(6,10);
if(mes==4 || mes==6 || mes==9 || mes==11 && dia > 30){
campo.focus();
alert("Dia incorreto ! O mês especificado contém no máximo 30 dias.");
return false;
} else{
if(ano%4!=0 && mes==2 && dia>28){
campo.focus();
alert("Data incorreta! O mês especificado contém no máximo 28 dias.");
return false;
} else{
if(ano%4==0 && mes==2 && dia>29){
campo.focus();
alert("Data incorreta! O mês especificado contém no máximo 29 dias.");
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
    
    
    </head>
</head>

<body onKeyDown="javascript:Verificar()">

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
	
	  
 
 <h2>PPV - Produtos - Etapas de Ordens de Serviços</h2><br />
 <!-- #include file="includes/classe_modulo_ppv_menu.asp" -->
 <%
AbrirConexao()
	
	sql = "SELECT * FROM TB_PRODUTOS WHERE gera_ordem_servico_prazo_cliente > 0 and ATIVO = 1 order by nome_produto"
	

	Set rs=server.CreateObject("ADODB.recordset")
	rs.CursorLocation=3  'Cursor no cliente
	rs.PageSize=50
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
%>
	
    <%
	Dim sucesso, codmsg
	codmsg	= CStr(""&Request.QueryString("codmsg"))
	sucesso = CStr(""&Request.QueryString("sucesso"))
		
	If sucesso = "" and codmsg <> "" Then%>
    	<div align="center" class="msgGrandeSucesso">
        	<%
			If codmsg = "1" Then
				Response.Write("Etapas incluídas com sucesso!")
			Elseif codmsg = "2" Then
				Response.Write("Etapas alteradas com sucesso!")
			Elseif codmsg = "3" Then	
				Response.Write("Etapas excluídas com sucesso!")
			End If	
			%>
</div>
<%ElseIf sucesso <> "" and codmsg <> "" Then%>
        <div align="center" class="msgGrandeErro">
			<%
			If codmsg = "4" Then
				Response.Write("Etapas não alterado para não haver duplicidade de informação!<BR>O conteúdo do campo """&Request.QueryString("chaveUnica")&""" já possui uma informação igual em outro Etapas.<br />")
			Elseif codmsg = "5" Then
				Response.Write("Etapas não excluído, pois há Orçamentos/Pedidos/NF que dependem ou relacionam com este vendedor.<BR>Se não desejar ter mais este Etapas inative ao invés de excluir.<br />")
			
			Elseif codmsg = "900" Then
				Response.Write("Você não possui acesso a VISUALIZAÇÃO!")
			Elseif codmsg = "901" Then
				Response.Write("Você não possui acesso a INCLUSÃO!")
			Elseif codmsg = "902" Then
				Response.Write("Você não possui acesso a ALTERAÇÃO!")	
			Elseif codmsg = "903" Then
				Response.Write("Você não possui acesso a EXCLUSÃO!")
			End If					
			%>
</div>
<%End If%>
<br />

<table width="99%" border="0" align="center">
  <tr>
    <td width="47%"> 
Página <strong><%=rs.AbsolutePage%></strong> de <%=rs.PageCount%>
</td>
    <td width="53%"><div align="right">Total <strong><%=rs.RecordCount%></strong> Produtos com Prazo Geral para Cliente Cadastrado</div></td>
  </tr>
</table>
<table width="99%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
<tr bgcolor="whitesmoke">
  <td width="8%"><div align="center" style="font-weight: bold">C&oacute;d. Produto</div></td>
  <td width="25%"><div align="center" style="font-weight: bold">Produtos</div></td>
  <td width="20%"><div align="center"><strong>Etapas</strong></div></td>
  <td width="10%"><div align="center"><strong>Prazo Geral para Cliente</strong></div></td>
  
  <td width="3%"><div align="center"><strong>Editar ✍</strong></div></td>
</tr>
<%
Dim x
For x=1 To rs.PageSize
If Not rs.Eof Then
%>
<tr>
  <td><div align="center"><%=formatNumber(rs(0),0)%></div></td>
  <td><div align="center"><%=rs("nome_produto")%></div></td>
  <td><div align="center">
    <%
  Set rs2=objConexao.Execute("SELECT * FROM TB_PRODUTOS_ETAPAS WHERE cod_produto = "&rs("cod_produto"))
			if not rs2.Eof then
				While not rs2.Eof
					Response.Write(rs2("numero_etapa")&"º - "&rs2("descricao")&" ("&rs2("prazo_em_dias")&" dia(s)) <br />")
				rs2.MoveNext()
				Wend
			Else
				Response.Write("<strong>não há etapas cadastradas!</strong>")
			end if
	rs2.close
	Set rs2=nothing
  %>
    
  </div></td>
  <td><div align="center"><%=rs("gera_ordem_servico_prazo_cliente")%> dias</div></td>
  <td>
	<%
	If Session("cod_usuario") = 1 or Session("cod_usuario") = 16 or Session("cod_usuario") = 22 or Session("cod_usuario") = 87 Then
	%>    
    <table border="0" align="center" cellpadding="3" cellspacing="0"> 
      <tr>
        <td>
          <a href="ppv_produtos_etapas_formulario.asp?cod_produto=<%=rs("cod_produto")%>">
            <img src="images/icons/icon_alterar.png" alt="Alterar Registro" title="Alterar Registro" width="16" height="16" border="0" /></a>  </td>
        </tr>
</table>
	
    <%Else%>
		<a href="javascript:alert('Alteração permitida somente pela D. Andrea.');">Bloqueado</a>
	<%
	End If
	%>
  </td>
  </tr>
<%
rs.MoveNext()
End If
Next
%>
<%If rs.PageCount > 1 Then %>                                  
     <tr bgcolor="whitesmoke">
      <td colspan="6">
      	   <%call paginacao(paginacaocao,rs.PageCount)%>      </td>
    </tr>   
<%End if%>
</table>	
<%
	Else
	%>
    <div align="center" class="msgGrande">
    	Nenhuma Etapa encontrada!
    <br />
	</div>
    <%End if
	
	rs.Close
	Set rs=Nothing
	
	FecharConexao()
	%>
    <br />
<br />
    Obs: Somente produtos que possuem prazo para o cliente cadastrado no Cadastro de Produtos
    
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