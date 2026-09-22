<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<%
Dim efetuar_triagem
efetuar_triagem = Cstr(""&Request.Form("campo1"))	

	
Dim busca_cod_produto, busca_cod_contrato

Dim sql, sqlWhere, sqlOrderBy, sqlGroupBy

sqlWhere = sqlWhere&" and status_ordem_servico = 1 and descricao = 'triagem'"

If Request.QueryString("cod_contrato") <> Empty Then
	
	
	
	busca_cod_produto = Trim(Cstr(""&Request.QueryString("cod_produto")))
	busca_cod_contrato = Trim(Cstr(""&Request.QueryString("cod_contrato")))
	
	If busca_cod_produto <> "" Then
		busca_cod_produto = Replace(busca_cod_produto," ","")
		sqlWhere = sqlWhere&" and cod_produto = "&busca_cod_produto
	End If
			
	If busca_cod_contrato <> "" Then
		busca_cod_contrato = Replace(busca_cod_contrato," ","")
		sqlWhere = sqlWhere&" and cod_contrato = "&sqlTexto(busca_cod_contrato)

	End If
End If

'/////////////////////
'Padroniza sqlWhere

	If sqlWhere <> "" Then
		'Retirar o "AND "
		sqlWhere = Mid(sqlWhere,5,1000)
	End If



If efetuar_triagem = "S" Then
	
	AbrirConexao()
		
		'1) Selecionar todas OS para triagem
		sql = "SELECT cod_ordem_servico, cod_ordem_servico_etapa_atual FROM VW_TB_ORDENS_SERVICO WHERE "&sqlWhere
		
		Set rs=objConexao.Execute(sql)
			If not rs.Eof then
				While not rs.Eof
					
				
					'3) Verifica se o responsável foi selecionado
					cod_ordem_servico = rs("cod_ordem_servico")
					cod_ordem_servico_etapa = rs("cod_ordem_servico_etapa_atual")
					
					cod_usuario_responsavel = Cstr(""&Request.Form("cod_usuario_responsavel_"&cod_ordem_servico))
					
					If cod_usuario_responsavel <> "" Then
						
						'Finaliza a etapa atual
						sql = "UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_termino = NOW(), cod_usuario_responsavel = "&Session("cod_usuario")&" WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&sqlTexto(cod_ordem_servico_etapa)		
						objConexao.Execute(sql)
						'Inicia a próxima Etapa e as demais muda a reponsabilidade apenas
						Set rs2=objConexao.Execute("SELECT cod_ordem_servico_etapa, numero_etapa FROM TB_ORDENS_SERVICO_ETAPAS WHERE cod_ordem_servico = "&cod_ordem_servico&" and numero_etapa > 1")
							
							If not rs2.Eof Then
								
								While not rs2.Eof
							
									If int(rs2("numero_etapa")) = 2 Then								
									'Inicia a etapa seguinte
										sql = "UPDATE TB_ORDENS_SERVICO_ETAPAS SET data_inicio = NOW(), cod_usuario_responsavel = "&sqlTexto(cod_usuario_responsavel)&" WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&rs2("cod_ordem_servico_etapa")
										objConexao.Execute(sql)
										sql = "UPDATE TB_ORDENS_SERVICO SET cod_ordem_servico_etapa_atual = "&rs2("cod_ordem_servico_etapa")&" WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)
										'response.write(sql)
										objConexao.Execute(sql)
					
									Else
										sql = "UPDATE TB_ORDENS_SERVICO_ETAPAS SET cod_usuario_responsavel = "&sqlTexto(cod_usuario_responsavel)&" WHERE cod_ordem_servico = "&sqlTexto(cod_ordem_servico)&" and cod_ordem_servico_etapa = "&rs2("cod_ordem_servico_etapa")
										objConexao.Execute(sql)									
									
									End If
									
										COD_USUARIO = Session("COD_USUARIO")
										sql = "INSERT INTO TB_ADM_LOG_ALTERACOES (COD_USUARIO,DATA_CADASTRADO,DESCRICAO) VALUES ("&COD_USUARIO&", NOW(),'Responsavel da ordem de serviço "&cod_ordem_servico&" foi alterado para o usuario "&cod_usuario_responsavel&"')"
										objConexao.Execute(sql)		
										
									
								rs2.MoveNext()
								Wend
								
							End If
							
						rs2.Close
						Set rs2=Nothing
											
					End If
					
								
				rs.MoveNext()
				Wend
				
			End if
			
		rs.close
		Set rs=nothing
		
	FecharConexao()
	
	Response.Redirect("ppv_ordens_servico_triagem_inicial.asp?codmsg=201")
	
End If
	


Dim acesso_inclusao_exclusao_os_manual
acesso_inclusao_exclusao_os_manual = validaUsuario(2,,51,false)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>StillosWEB - Sistema corporativo - Vers&atilde;o 1.2</title>
<!-- #include virtual="gerencial_foto/files/includes/scripts.asp" -->
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
//Formata número tipo moeda usando o evento onKeyDown
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
	
	function confirmaTriagem(){
		if(YNconfirm('Tem certeza que deseja efetuar a triagem das OSs?')){
			document.TRIAGEM_INICIAL.submit();
		}
	}
	
	function selecionarVendedores(cod_pedido,cod_ordem_servico){
		var temp = "", a = "", cod_usuario_responsavel = "";
		var i=0
		temp = document.TRIAGEM_INICIAL["marcar_"+cod_pedido].value;
		cod_usuario_responsavel = document.TRIAGEM_INICIAL["cod_usuario_responsavel_"+cod_ordem_servico].value;
		a = temp.split(";");
		
		for(i=0;i<a.length;i++){
			document.TRIAGEM_INICIAL["cod_usuario_responsavel_"+a[i]].value = cod_usuario_responsavel;
		}		
	}
	function filtrarContrato(contrato){
		if(contrato !== ''){
			var d = document.FILTRO_PADRAO;
			d.COD_CONTRATO.value = contrato;
			d.submit();
		}
	}
	
	</script>
    
    
    </head>
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
    <td><!-- #include file="includes/menu.asp" --></td>
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
	
	  <h2>PPV - Triagem Inicial de Ordens de Servi&ccedil;o</h2>
         <br />
      <!-- #include file="includes/classe_modulo_ppv_menu.asp" -->
	  <br />
   
    <%
	Dim sucesso, codmsg
	codmsg	= CStr(""&Request.QueryString("codmsg"))
	sucesso = CStr(""&Request.QueryString("sucesso"))
		
	If sucesso = "" and codmsg <> "" Then%>
    	<div align="center" class="msgGrandeSucesso">
        	<%
			If codmsg = "201" Then	
				Response.Write("Triagem inicial realizada com sucesso!")
			End If	
			%>
            
        
        </div>
        
   <%End If %>


      <form action="ppv_ordens_servico_triagem_inicial.asp" name="FILTRO_PADRAO" style="margin:0px">
	  <table style="border:#999999 dashed 1px; padding:2px; background:#FFFFCC" width="99%" border="0" align="center" cellpadding="5" cellspacing="0">
      <!--
		  <tr>
		    <td width="19%"><div align="right"><strong>Produto</strong></div></td>
		    <td width="81%"><div align="left"><select name='COD_PRODUTO' class='inputTamanho3'>
    <option value=''>Todos</option>
    <%AbrirConexao()
	
    Set rsE3=objConexao.Execute("SELECT COD_PRODUTO, NOME_PRODUTO FROM TB_PRODUTOS WHERE ATIVO = 1 AND gera_ordem_servico_prazo_cliente > 0 ORDER BY NOME_PRODUTO ")
    While not rsE3.Eof 
    'Padrao
    Dim selRelPadrao3, verificador3
    verificador3 = ""&rsE3(0)
    
    selRelPadrao3 = "" 
    
    
    If CStr(verificador3) = CStr(busca_cod_produto)  Then%>
    <option selected='selected' value="<%=rsE3(0)%>"><%=rsE3(1)%></option>
    <%Else%>
    <option value="<%=rsE3(0)%>"><%=rsE3(1)%></option>
    <%End if%>
    <%rsE3.MoveNext()
    Wend
    rsE3.Close
    Set rsE3=Nothing
    FecharConexao()%>
    
    </select></div></td>
		    </tr>
       -->
            <tr>
		    <td width="19%"><div align="right"><strong>Contrato</strong></div></td>
		    <td width="81%"><div align="left"><input type="text" name="COD_CONTRATO" value="<%=busca_COD_CONTRATO%>" /></div></td>
		    </tr>
		  <tr>
		    <td> </td>
		    <td><div align="left">
		      <input class="btnFiltrarAmarelo" type="submit" name="filtro1" value=" " />
		      </div></td>
		    </tr>
        </table>
      </form>
        <br />

        <em><strong>Listagem em Ordem alfab&eacute;tica com todas OS, cuja as etapas iniciais estão em aberto e o nome da etapa é "Triagem".</strong></em><br />
        <br />
    	
        <%
		AbrirConexao()
		
		
		Dim qtd_geral_triagem
		Set rsCont = objConexao.Execute("SELECT COUNT(*) AS qtd FROM VW_TB_ORDENS_SERVICO WHERE descricao = 'triagem' AND status_ordem_servico = 1")
		
			If not rsCont.Eof Then
				qtd_geral_triagem = rsCont(0)
			End If
			
		rsCont.Close
		Set rsCont = Nothing
		%>
        
        	
            <div align="center" style="padding:10px; background-color:#F2F2F2">Qtd. Total Triagem Pendente Geral:  <strong><%=qtd_geral_triagem%></strong><br />Contratos: 
            <%		
			Dim contratos
			contratos = ""
			Set rsCont = objConexao.Execute("SELECT COD_CONTRATO FROM VW_TB_ORDENS_SERVICO WHERE descricao = 'triagem' AND status_ordem_servico = 1 GROUP BY COD_CONTRATO ORDER BY COD_CONTRATO ASC")
			While not rsCont.Eof
				If contratos="" Then
					contratos = "<a href=""javascript:filtrarContrato('"&rsCont("COD_CONTRATO")&"')"">"&rsCont("COD_CONTRATO")&"</a>"
				Else
					contratos = contratos&", <a href=""javascript:filtrarContrato('"&rsCont("COD_CONTRATO")&"')"">"&rsCont("COD_CONTRATO")&"</a>"
				End If
			rsCont.MoveNext()
			Wend
			
			rsCont.Close
			Set rsCont = Nothing
			
			If contratos = "" Then
				contratos = "NÃO HÁ ORDENS DE serviço PARA SEREM TRIADAS!"
			End If
			Response.Write(contratos)
			%>
            
            </div>
            <br />
<br />

        
        
        <%If Request.QueryString("filtro1") <> Empty  Then%>
            
        <%
		Dim temp
		'Neste momento corrige o sqlWhere somente para o primeiro rs
		If sqlWhere <> "" Then
			temp = "WHERE "
		End If
		
		sql = "SELECT cod_pedido, cod_formando FROM VW_TB_ORDENS_SERVICO "&temp&sqlWhere&" GROUP BY cod_pedido, cod_formando ORDER BY COD_CONTRATO, NOME_FORMANDO"
		Set rs=objConexao.Execute(sql)
		
		If not rs.Eof Then
			'Verifica a qtd de OS de acordo com o filtro
			Set rs2=objConexao.Execute("SELECT COUNT(*) as qtd FROM VW_TB_ORDENS_SERVICO "&temp&sqlWhere)
				If not rs2.Eof Then
					qtdOSs = rs2(0)
				End If
			rs2.Close
			Set rs2=Nothing
		%>
        <form name="TRIAGEM_INICIAL" action="ppv_ordens_servico_triagem_inicial.asp" method="post" style="margin:0px">
        	<input name="campo1" value="S" type="hidden" />
            <table width="99%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
              <tr bgcolor="whitesmoke">
        <td colspan="2"><div align="center"><strong> OS</strong></div></td>
        <td width="11%"><div align="center"><strong>Formando</strong></div></td>
        <td width="7%"><div align="center"><strong>Contrato</strong></div></td>
        <td width="20%"><div align="center"><strong>Produto</strong></div></td>
        <td width="13%"><div align="center"><strong>Data Cadastro</strong></div></td>
        <td width="17%"><div align="center"><strong>Prazo Geral</strong></div></td>
        <td width="24%"><div align="center"><strong>Respons&aacute;vel pela OS</strong></div></td>
        </tr>
      <%
    
    'Neste momento que já passou o group by corrige o sqlWhere para o rs2
    If sqlWhere <> "" Then
        sqlWhere = " AND "&sqlWhere
    End If
    
    Dim numero_etapa, descricao, cod_usuario_responsavel, nome_usuario_responsavel, prazo_em_dias, data_inicio, prazo_limite, diferenca, qtd_etapas,  ultimo_responsavel
    While not rs.Eof
    %>
        
    <tr>
                                <td colspan="8"><table width="650" border="0" cellspacing="0" cellpadding="1">
                                  <tr>
                                    <td class="textoMedio"><strong><%=rs("cod_pedido")%></strong>                                    </td>
                                  </tr>
                                </table></td>
                              </tr>
                              
           <%	   
           sql = "SELECT * FROM VW_TB_ORDENS_SERVICO WHERE cod_pedido = "&rs("cod_pedido")&sqlWhere&" order by numero_ordem_servico desc, numero_ordem_servico_indice asc"
           'Response.Write(sql)
           Set rs2=objConexao.Execute(sql,qtdEncontrado)
           cont = 0
		   ordens_servico = ""
		   
           While not rs2.Eof
           
                    cont = cont+1
					ordens_servico = ordens_servico&";"&rs2("cod_ordem_servico")
                    Set rs3=objConexao.Execute("Select count(*) as qtd_etapas From TB_ORDENS_SERVICO_ETAPAS where cod_ordem_servico = "&rs2("cod_ordem_servico")&" group by cod_ordem_servico")
                        
                        If not rs3.Eof Then
                            qtd_etapas = formatNumber(0&rs3("qtd_etapas"),0)
                        End If
                        
                    rs3.Close
                    Set rs3=Nothing
           
           %>                   
                              
                <tr>
                                            <td width="3%"><div align="right"><img src="images/icons/seta_superior.gif" width="19" height="15" /></div></td>
                              <td width="5%"><div align="center"><%=moduloPPVOSMascara(rs2("cod_ordem_servico"),"",true)%></div></td>
                              <td><%
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
											" WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&rs("cod_pedido")&" and TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = "&rs("cod_formando")&" ORDER BY NOME_FORMANDO"
											
											           Set rs5=objConexao.Execute(sql)
													   
													   		If not rs5.Eof Then
																NOME_FORMANDO = rs5("NOME_FORMANDO")
																COD_CONTRATO = rs5("COD_CONTRATO")
															End If
													   
													   	 Response.Write(NOME_FORMANDO)
													   %>
           													
                                                         
                                                        <%
														
														rs5.Close
														Set rs5=Nothing
														%></td>
                              <td><%=(COD_CONTRATO)%></td>
                                            <td><div align="left">
                                            <%=rs2("nome_produto")%>                                        
                                            </div></td>
                                            <td><div align="center"><%=ArrumaData(rs2("data_cadastrado"))%></div></td>
                                            <td><div align="center"> <%=ArrumaData(DateAdd("d",rs2("prazo_cliente_total"),rs2("data_base_prazo_cliente_inicio")))%></div></td>
                                            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                                              <tr>
                                                <td><div align="center">
                                                <%
												ultimo_responsavel = ""
												'Verifica último responsável, sem ser desta OS
												Set rs3=objConexao.Execute("SELECT cod_usuario_responsavel, nome_formando FROM `VW_TB_ORDENS_SERVICO` WHERE cod_contrato = "&sqlTexto(rs2("cod_contrato"))&" and cod_ordem_servico <> "&rs2("cod_ordem_servico")&" and cod_produto = "&rs2("cod_produto")&" LIMIT 1")
													If not rs3.Eof Then
														ultimo_responsavel = Cstr(""&rs3(0))
													End If
												rs3.Close
												Set rs3=Nothing
												
												%>
                                                
                                                
                                                  <select class="inputTamanho2" name="cod_usuario_responsavel_<%=rs2("cod_ordem_servico")%>">
                                                    <option value="" <%If ultimo_responsavel = "" or 1 = 1 Then%> selected="selected" <%End If%>></option>
                                                    <%
                                            
                                            'Seleciona todos os funcion&aacute;rios ativos, exceto o usu&aacute;rio atual.
                                            Set rs3=objConexao.Execute("Select cod_usuario, nome FROM TB_ADM_USUARIOS where bloqueado = 2 AND cod_usuario > 1 order by nome")
                                            While not rs3.Eof%>
                                                    <option value="<%=rs3(0)%>" <%If 1 = 2 and ultimo_responsavel = Cstr(""&rs3(0)) Then%> selected="selected" <%End If%>><%=rs3(1)%></option>
                                                    <%
                                            rs3.MoveNext()
                                            Wend	
                                                
                                        
                                            rs3.Close
                                            Set rs3=Nothing
                                        %>
                                                  </select>
                                                </div></td>
                                              </tr>
                                              <%If qtdEncontrado > 1 and cont = 1 Then%>
                                              <tr>
                                                <td><div align="center"><a href="javascript:selecionarVendedores(<%=rs("cod_pedido")%>,<%=rs2("cod_ordem_servico")%>);">Marcar todas desta venda com o resp. acima</a></div></td>
                                              </tr>
                                              <%End If%>
                                            </table></td>
                              </tr>
               
                    
           <%
           rs2.MoveNext()
           Wend
           
           rs2.Close
           Set rs2=Nothing
           
		   If ordens_servico <> "" Then
		   		ordens_servico = Mid(ordens_servico,2,Len(ordens_servico))
		   End If
		   %>
           
           <input type="hidden" name="marcar_<%=rs("cod_pedido")%>" value="<%=ordens_servico%>" />         
    <%
    rs.MoveNext()
    Wend
    %>
          <tr>
                  <td colspan="8"><div align="center"><a href="javascript:confirmaTriagem();"><img border="0" src="images/layout/btnGravar.jpg" /></a></div></td>
                  </tr>
         
          </table>
      	</form>


<%
	Else
	%>
    <div align="center" class="msgGrande">
    Nenhuma OS encontrada para triagem!<br />
</div>
    <%End if
	
	rs.Close
	Set rs=Nothing
	
	
	%>
    
    	<%Else%>
    	
        <div class="msgGrande" align="center">Utilize o filtro por contrato para iniciar a triagem</div>
        <%End if%>
	  </p>
	   <br />
	   <br />
       <strong>Observa&ccedil;&otilde;es:<br />
       </strong>1) Ap&oacute;s a triagem a etapa 1 ser&aacute; finalizada<br />2) Virá pré-selecionado o último responsável pelo CONTRATO e PRODUTO, caso o sistema não encontrar um último responsável não haverá pré-seleção
<p align="center"> </p>
	    </td>
    </tr>
</table>
<%
FecharConexao()
%>

 
  
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
