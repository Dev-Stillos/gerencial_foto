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
<%
Dim triageView
triageView = Request.QueryString("painel_triagem") & ""
If Request.ServerVariables("REQUEST_METHOD") = "GET" And (triageView = "pendentes" Or triageView = "concluidos" Or triageView = "historico") Then
    ExibirPainelTriagem triageView
    Response.End
End If

Sub ExibirPainelTriagem(ByVal modo)
    Dim digitosPagina, registros, consulta, filtro, codigo, termo, pagina, inicio, tamanho, exibidos, urlBase, condPendente, condConcluida, escopo, podeVerTodas, osAnterior, osAtual
    Response.CodePage = 65001
    Response.Charset = "utf-8"
    Response.CacheControl = "no-cache"
    Response.AddHeader "X-Triagem-Painel", "1"
    codigo = Left(Trim(Request.QueryString("contrato") & ""), 30)
    termo = Left(Trim(Request.QueryString("busca") & ""), 30)
    pagina = Request.QueryString("pagina") & ""
    Set digitosPagina = New RegExp
    digitosPagina.Pattern = "^[0-9]{1,5}$"
    If Not digitosPagina.Test(pagina) Then pagina = "1"
    pagina = CLng(pagina)
    If pagina < 1 Then pagina = 1
    tamanho = 30
    condPendente = "(o.status_ordem_servico = 1 AND o.descricao = 'triagem')"
    podeVerTodas = validaUsuario(2,,57,false)
    escopo = "1=1"
    If Not podeVerTodas Then escopo = "o.cod_usuario_responsavel = " & sqlNumero(Session("COD_USUARIO"))
    condConcluida = "(t.data_termino >= '1000-01-01' AND NOT " & condPendente & " AND " & escopo & ")"
    urlBase = "?painel_triagem=" & modo & "&contrato=" & Server.URLEncode(codigo) & "&busca=" & Server.URLEncode(termo)
    AbrirConexao()
    If modo = "historico" Then
        tamanho = 10
        inicio = (pagina - 1) * tamanho
        consulta = "SELECT ordens.cod_ordem_servico, e.numero_etapa, e.descricao, e.data_inicio, e.data_termino, u.nome AS responsavel " & _
            "FROM (SELECT o.cod_ordem_servico FROM VW_TB_ORDENS_SERVICO o WHERE o.cod_contrato = " & sqlTexto(codigo) & " AND " & escopo & _
            " AND EXISTS (SELECT 1 FROM TB_ORDENS_SERVICO_ETAPAS h WHERE h.cod_ordem_servico = o.cod_ordem_servico)" & _
            " GROUP BY o.cod_ordem_servico ORDER BY o.cod_ordem_servico DESC LIMIT " & inicio & "," & (tamanho + 1) & ") ordens " & _
            "INNER JOIN TB_ORDENS_SERVICO_ETAPAS e ON e.cod_ordem_servico = ordens.cod_ordem_servico " & _
            "LEFT JOIN TB_ADM_USUARIOS u ON u.cod_usuario = e.cod_usuario_responsavel ORDER BY ordens.cod_ordem_servico DESC, e.numero_etapa"
        Set registros = objConexao.Execute(consulta)
%>
        <h3 class="triage-history-title">Contrato <%=Server.HTMLEncode(codigo)%></h3>
        <p>In&iacute;cio, t&eacute;rmino e respons&aacute;vel registrado em cada etapa. As ordens seguem as permiss&otilde;es do seu perfil.</p>
        <% If registros.Eof Then %><p>Nenhuma etapa dispon&iacute;vel para exibi&ccedil;&atilde;o.</p><% End If %>
        <% osAnterior = ""
           exibidos = 0
           Do While Not registros.Eof
               osAtual = registros("cod_ordem_servico") & ""
               If osAtual <> osAnterior Then
                   If exibidos = tamanho Then Exit Do
                   If exibidos > 0 Then Response.Write "</ul></details>"
                   exibidos = exibidos + 1
                   osAnterior = osAtual
        %>
        <details class="triage-history-os"<%If exibidos = 1 Then%> open<%End If%>>
            <summary>OS <%=Server.HTMLEncode(moduloPPVOSMascara(osAtual,"",true) & "")%></summary>
            <a href="ppv_ordens_servico_formulario.asp?comando=visualizar&amp;cod_ordem_servico=<%=Server.URLEncode(osAtual)%>">Abrir ficha da OS e anota&ccedil;&otilde;es</a>
            <ul class="triage-history-steps">
        <% End If %>
            <li>
                <strong><%=Server.HTMLEncode(registros("numero_etapa") & ". " & registros("descricao"))%></strong>
                <span><b>In&iacute;cio</b><%=TextoPainelTriagem(registros("data_inicio"))%></span>
                <span><b>T&eacute;rmino</b><%=TextoPainelTriagem(registros("data_termino"))%></span>
                <span><b>Respons&aacute;vel</b><%=TextoPainelTriagem(registros("responsavel"))%></span>
            </li>
        <%     registros.MoveNext()
           Loop
           If exibidos > 0 Then Response.Write "</ul></details>"
        %>
<%
    Else
        inicio = (pagina - 1) * tamanho
        filtro = "o.cod_contrato IS NOT NULL AND o.cod_contrato <> ''"
        If termo <> "" Then filtro = filtro & " AND INSTR(UPPER(o.cod_contrato), " & sqlTexto(UCase(termo)) & ") > 0"
        If modo = "pendentes" Then
            consulta = "SELECT o.cod_contrato, COUNT(*) AS quantidade, NULL AS ultima FROM VW_TB_ORDENS_SERVICO o WHERE " & filtro & _
                " AND " & condPendente & " GROUP BY o.cod_contrato ORDER BY o.cod_contrato LIMIT " & inicio & "," & (tamanho + 1)
        Else
            consulta = "SELECT o.cod_contrato, COUNT(DISTINCT CASE WHEN " & condConcluida & " THEN o.cod_ordem_servico END) AS quantidade, " & _
                "MAX(CASE WHEN " & condConcluida & " THEN t.data_termino END) AS ultima " & _
                "FROM VW_TB_ORDENS_SERVICO o LEFT JOIN TB_ORDENS_SERVICO_ETAPAS t ON t.cod_ordem_servico = o.cod_ordem_servico AND t.numero_etapa = 1 AND t.descricao = 'triagem' " & _
                "WHERE " & filtro & " GROUP BY o.cod_contrato HAVING quantidade > 0 AND SUM(CASE WHEN " & condPendente & _
                " THEN 1 ELSE 0 END) = 0 ORDER BY ultima DESC, o.cod_contrato LIMIT " & inicio & "," & (tamanho + 1)
        End If
        Set registros = objConexao.Execute(consulta)
%>
        <form class="triage-filter triage-panel-search" method="get">
            <input type="hidden" name="painel_triagem" value="<%=modo%>">
            <div><label for="triage-modal-search">Buscar contrato</label><input class="form-control" id="triage-modal-search" name="busca" maxlength="30" value="<%=Server.HTMLEncode(termo)%>"></div>
            <button class="btn btn-primary" type="submit">Buscar</button>
        </form>
        <% If modo = "concluidos" Then %>
        <p>Contratos sem ordens aguardando triagem e com conclus&atilde;o registrada nas ordens que seu perfil pode consultar.</p>
        <% Else %>
        <p>A quantidade ao lado de cada contrato indica quantas ordens ainda aguardam triagem.</p>
        <% End If %>
        <% If registros.Eof Then %><p>Nenhum contrato encontrado.</p><% End If %>
        <ul class="triage-modal-list">
        <% exibidos = 0
           While Not registros.Eof And exibidos < tamanho
               codigo = registros("cod_contrato") & ""
        %>
            <li>
                <div><strong><%=Server.HTMLEncode(codigo)%></strong> <span class="triage-count"><%=registros("quantidade")%> OS</span>
                <% If modo = "concluidos" Then %><small>&Uacute;ltima triagem <%=TextoPainelTriagem(registros("ultima"))%></small><% End If %></div>
                <div class="triage-modal-actions">
                    <a class="btn btn-outline-secondary btn-sm" data-triage-load href="?painel_triagem=historico&amp;contrato=<%=Server.URLEncode(codigo)%>">Ver etapas</a>
                    <% If modo = "pendentes" Then %><a class="btn btn-primary btn-sm" href="?filtro1=1&amp;COD_CONTRATO=<%=Server.URLEncode(codigo)%>">Abrir triagem</a><% End If %>
                </div>
            </li>
        <%     exibidos = exibidos + 1
               registros.MoveNext()
           Wend %>
        </ul>
<%
    End If
%>
        <nav class="triage-modal-actions" aria-label="P&aacute;ginas dos resultados">
            <% If pagina > 1 Then %><a class="btn btn-outline-secondary btn-sm" data-triage-load href="<%=Server.HTMLEncode(urlBase & "&pagina=" & (pagina - 1))%>">Anterior</a><% End If %>
            <% If Not registros.Eof Then %><a class="btn btn-outline-secondary btn-sm" data-triage-load href="<%=Server.HTMLEncode(urlBase & "&pagina=" & (pagina + 1))%>">Pr&oacute;xima</a><% End If %>
        </nav>
<%
    registros.Close
    Set registros = Nothing
    FecharConexao()
End Sub

Function TextoPainelTriagem(ByVal valor)
    If Len(Trim(valor & "")) = 0 Then
        TextoPainelTriagem = "&mdash;"
    Else
        TextoPainelTriagem = Server.HTMLEncode(valor & "")
    End If
End Function
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Triagem inicial | Portal Stillos</title>
<link rel="stylesheet" href="css/vendor/bootstrap.min.css">
<link rel="stylesheet" href="css/dashboard.css">
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
	
	</script>
    
    
    
<style>
.triage-page .dashboard-main { min-width:0; }
.triage-page .dashboard-nav a.active:not([aria-current]) { background:transparent; color:#d4d5e1; }
.triage-page .dashboard-nav a[aria-current="page"] { background:#b20b61; color:#fff; font-weight:600; }
.triage-page .ppv-module { margin-bottom:20px; }
.triage-page .subMenu { padding:12px 16px; border:1px solid var(--line); border-radius:10px; background:white; line-height:2; }
.triage-page .subMenu [aria-current="page"] { font-weight:700; color:#88134f; }
.triage-page .triage-panel { height:auto; margin-bottom:18px; padding:clamp(16px, 3vw, 26px); }
.triage-heading { display:flex; align-items:center; flex-wrap:wrap; gap:10px; margin-bottom:16px; }
.triage-heading h2 { color:#242738; }
.triage-count { padding:4px 9px; border-radius:6px; background:#f3eaf3; color:#714064; font-weight:650; }
.triage-filter { display:flex; flex-wrap:wrap; align-items:flex-end; gap:8px; }
.triage-filter label { display:block; margin-bottom:5px; font-weight:600; }
.triage-dialog.triage-browser {
  box-sizing:border-box;
  width:min(1440px, calc(100% - 48px));
  max-width:none;
  height:calc(100vh - 48px);
  height:calc(100dvh - 48px);
  max-height:none;
  padding:0;
  overflow:hidden;
}
.triage-browser[open] { display:flex; flex-direction:column; }
.triage-browser-header { display:flex; justify-content:space-between; align-items:center; gap:8px; padding:16px 20px; border-bottom:1px solid var(--line); }
.triage-browser-content { flex:1; overflow:auto; overscroll-behavior:contain; min-height:0; padding:20px 24px; }
.triage-modal-actions { display:flex; flex-wrap:wrap; gap:8px; }
.triage-panel-search { margin-bottom:14px; }
.triage-modal-list, .triage-history-steps { list-style:none; padding:0; }
.triage-modal-list li { display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px; padding:10px 0; border-bottom:1px solid var(--line); }
.triage-modal-list small { display:block; margin-top:4px; color:#55475e; }
.triage-history-os { padding:12px; margin:10px 0; border:1px solid var(--line); border-radius:8px; }
.triage-history-os summary { font-weight:700; cursor:pointer; }
.triage-history-steps { margin:8px 0 0; }
.triage-history-title { margin:16px 0 6px; color:var(--ink); font:700 1.25rem/1.3 system-ui,sans-serif; }
.triage-history-steps li { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:8px 16px; padding:12px 0; border-top:1px solid var(--line); overflow-wrap:anywhere; }
.triage-history-steps strong { grid-column:1/-1; }
.triage-history-steps span b { display:block; font-size:.85em; color:#55475e; }
.triage-table { margin-bottom:0; }
.triage-table th, .triage-table td { padding:12px 8px; vertical-align:top; overflow-wrap:anywhere; }
.triage-table thead th, .triage-table .triage-order-group th { background:#f4f1f7; color:#42364e; }
.triage-table .form-select { min-width:140px; }
.triage-copy { margin-top:7px; white-space:normal; }
.triage-actions { display:flex; justify-content:flex-end; flex-wrap:wrap; gap:8px; margin-top:16px; }
.triage-dialog { width:min(450px, calc(100% - 24px)); padding:24px; border:1px solid var(--line); border-radius:14px; color:var(--ink); }
.triage-dialog::backdrop { background:#24273899; }
.triage-dialog p { margin:10px 0; }
@media(max-width:1050px) {
  .triage-table, .triage-table tbody { display:block; }
  .triage-table thead { position:absolute; width:1px; height:1px; overflow:hidden; clip-path:inset(50%); }
  .triage-order-group { display:block; margin-top:12px; }
  .triage-order-group th { display:block; border-radius:7px; }
  .triage-order { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); border-bottom:1px solid var(--line); }
  .triage-order td { display:block; border:0; }
  .triage-order td:last-child { grid-column:1/-1; }
  .triage-order td::before { content:attr(data-label); display:block; margin-bottom:4px; color:#66586e; font-size:.85em; font-weight:700; }
}
@media(max-width:640px) {
  .triage-dialog.triage-browser {
    width:calc(100% - 16px);
    height:calc(100vh - 16px);
    height:calc(100dvh - 16px);
  }
  .triage-browser-header, .triage-browser-content { padding:12px 14px; }
  .triage-browser-header h2 { min-width:0; overflow-wrap:anywhere; }
  .triage-browser .triage-modal-list li { flex-direction:column; align-items:stretch; }
  .triage-browser .triage-modal-list .triage-modal-actions > .btn { flex:1 1 130px; }
  .triage-browser .triage-history-steps li { grid-template-columns:minmax(0,1fr); }
}
@media(max-width:420px) {
  .triage-filter > div, .triage-filter > button { width:100%; }
}
</style>
</head>
<body class="triage-page">
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
      <div class="topbar-actions"><a class="user-name" href="adm_log_usuarios.asp"><%=Server.HTMLEncode(Session("NOME") & "")%></a><a href="manual_sistema.asp" target="_blank" rel="noopener">Manual</a><a class="btn btn-outline-secondary btn-sm" href="adm_sair.asp">Sair</a></div>
    </header>
    <main class="dashboard-main ppv-main" id="conteudo">
      <div class="dashboard-heading">
        <div><p class="eyebrow">PRODU&Ccedil;&Atilde;O P&Oacute;S-VENDA</p><h1>Triagem inicial</h1><p>Selecione um contrato e atribua respons&aacute;veis &agrave;s ordens de servi&ccedil;o.</p></div>
        <a href="ppv_ordens_servico.asp" class="btn btn-outline-secondary">Voltar &agrave;s ordens</a>
      </div>
      <div class="search-panel"><div><strong>Busca no portal</strong></div><!-- #include file="includes/dashboard_busca.asp" --></div>
      <div class="session-notices"><!-- #include file="includes/dashboard_avisos.asp" --></div>
      <nav class="ppv-module" aria-label="Produ&ccedil;&atilde;o p&oacute;s-venda">
        <!-- #include file="includes/classe_modulo_ppv_menu.asp" -->
      </nav>

      <%If CStr("" & Request.QueryString("sucesso")) = "" And CStr("" & Request.QueryString("codmsg")) = "201" Then%>
      <p class="alert alert-success" role="status">Triagem inicial realizada com sucesso.</p>
      <%End If%>
      <section class="dashboard-panel triage-panel" aria-labelledby="triage-filter-title">
        <div class="triage-heading"><h2 id="triage-filter-title">Filtrar por contrato</h2></div>
        <form class="triage-filter" action="ppv_ordens_servico_triagem_inicial.asp" method="get" name="FILTRO_PADRAO">
          <input type="hidden" name="filtro1" value="1">
          <div><label for="triage-contract">C&oacute;digo do contrato</label><input class="form-control" id="triage-contract" name="COD_CONTRATO" value="<%=Server.HTMLEncode(busca_COD_CONTRATO & "")%>"></div>
          <button class="btn btn-primary" type="submit">Filtrar ordens</button>
          <%If Len(busca_COD_CONTRATO & "") > 0 Then%><a class="btn btn-outline-secondary" href="ppv_ordens_servico_triagem_inicial.asp">Limpar filtro</a><%End If%>
        </form>
      </section>
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
        
        	

      <section class="dashboard-panel triage-panel" aria-labelledby="triage-pending">
        <div class="triage-heading">
          <h2 id="triage-pending">Contratos por situa&ccedil;&atilde;o</h2>
          <span class="triage-count">Total na fila: <%=qtd_geral_triagem%> OS</span>
        </div>
        <div class="triage-modal-actions">
          <button class="btn btn-primary" type="button" aria-haspopup="dialog" data-triage-mode="pendentes">Triagem pendente</button>
          <button class="btn btn-outline-secondary" type="button" aria-haspopup="dialog" data-triage-mode="concluidos">Triagem completa</button>
        </div>
      </section>
      <dialog class="triage-dialog triage-browser" id="triage-browser" aria-labelledby="triage-browser-title">
        <header class="triage-browser-header">
          <h2 id="triage-browser-title" tabindex="-1">Contratos</h2>
          <form method="dialog"><button class="btn btn-outline-secondary" autofocus>Fechar</button></form>
        </header>
        <div class="triage-browser-content">
          <button class="btn btn-outline-secondary btn-sm" id="triage-browser-back" type="button" hidden>Voltar aos contratos</button>
          <div id="triage-browser-results" aria-busy="false"></div>
        </div>
      </dialog>
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
        <section class="dashboard-panel triage-panel" aria-labelledby="triage-orders">
        <div class="triage-heading">
          <h2 id="triage-orders">Ordens para triagem</h2>
          <span class="triage-count"><%=qtdOSs%> OS</span>
        </div>
        <form id="triage-form" name="TRIAGEM_INICIAL" action="ppv_ordens_servico_triagem_inicial.asp" method="post" style="margin:0px">
        	<input name="campo1" value="S" type="hidden" />
            <table class="table triage-table" aria-labelledby="triage-orders" role="table">
              <thead role="rowgroup"><tr role="row">
                <th scope="col" role="columnheader">OS</th>
                <th scope="col" role="columnheader">Formando</th>
                <th scope="col" role="columnheader">Contrato</th>
                <th scope="col" role="columnheader">Produto</th>
                <th scope="col" role="columnheader">Cadastro</th>
                <th scope="col" role="columnheader">Prazo geral</th>
                <th scope="col" role="columnheader">Respons&aacute;vel pela OS</th>
              </tr></thead>
              <tbody role="rowgroup">
      <%
    
    'Neste momento que já passou o group by corrige o sqlWhere para o rs2
    If sqlWhere <> "" Then
        sqlWhere = " AND "&sqlWhere
    End If
    
    Dim numero_etapa, descricao, cod_usuario_responsavel, nome_usuario_responsavel, prazo_em_dias, data_inicio, prazo_limite, diferenca, qtd_etapas,  ultimo_responsavel
    While not rs.Eof
    %>
        
              <tr class="triage-order-group" role="row">
                <th colspan="7" role="cell">Pedido #<%=rs("cod_pedido")%></th>
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
                              
                <tr class="triage-order" role="row">
                              <td role="cell" data-label="OS"><strong><%=moduloPPVOSMascara(rs2("cod_ordem_servico"),"",true)%></strong></td>
                              <td role="cell" data-label="Formando"><%
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
                              <td role="cell" data-label="Contrato"><%=(COD_CONTRATO)%></td>
                                            <td role="cell" data-label="Produto"><%=rs2("nome_produto")%></td>
                                            <td role="cell" data-label="Cadastro"><%=ArrumaData(rs2("data_cadastrado"))%></td>
                                            <td role="cell" data-label="Prazo geral"><%=ArrumaData(DateAdd("d",rs2("prazo_cliente_total"),rs2("data_base_prazo_cliente_inicio")))%></td>
                                            <td role="cell" data-label="Respons&aacute;vel pela OS"><%
												ultimo_responsavel = ""
												'Verifica último responsável, sem ser desta OS
												Set rs3=objConexao.Execute("SELECT cod_usuario_responsavel, nome_formando FROM `VW_TB_ORDENS_SERVICO` WHERE cod_contrato = "&sqlTexto(rs2("cod_contrato"))&" and cod_ordem_servico <> "&rs2("cod_ordem_servico")&" and cod_produto = "&rs2("cod_produto")&" LIMIT 1")
													If not rs3.Eof Then
														ultimo_responsavel = Cstr(""&rs3(0))
													End If
												rs3.Close
												Set rs3=Nothing
												
												%>
                                                
                                                
                                                  <select class="form-select" aria-label="Respons&aacute;vel pela OS <%=rs2("cod_ordem_servico")%>" name="cod_usuario_responsavel_<%=rs2("cod_ordem_servico")%>">
                                                    <option value="" <%If ultimo_responsavel = "" or 1 = 1 Then%> selected="selected" <%End If%>>Selecione</option>
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
                                                <%If qtdEncontrado > 1 and cont = 1 Then%>
                                                <button type="button" class="btn btn-outline-secondary btn-sm triage-copy" onclick="selecionarVendedores(<%=rs("cod_pedido")%>,<%=rs2("cod_ordem_servico")%>)">Aplicar ao pedido</button>
                                                <%End If%>
                                            </td>
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
           
           <tr hidden><td colspan="7"><input type="hidden" name="marcar_<%=rs("cod_pedido")%>" value="<%=ordens_servico%>" /></td></tr>         
    <%
    rs.MoveNext()
    Wend
    %>
              </tbody>
            </table>
            <div class="triage-actions"><button type="button" class="btn btn-primary" id="triage-review" aria-haspopup="dialog">Concluir triagem</button></div>
        </form>
        <dialog class="triage-dialog" id="triage-confirm" aria-labelledby="triage-confirm-title" aria-describedby="triage-confirm-message">
          <h2 id="triage-confirm-title">Concluir triagem?</h2>
          <p id="triage-confirm-message">A etapa inicial ser&aacute; finalizada nas ordens com respons&aacute;vel selecionado.</p>
          <form method="dialog" class="triage-actions">
            <button class="btn btn-outline-secondary" value="cancelar" autofocus>Cancelar</button>
            <button class="btn btn-primary" value="confirmar">Confirmar triagem</button>
          </form>
        </dialog>
        </section>


<%
	Else
	%>
    <div class="dashboard-panel triage-panel" role="status">Nenhuma ordem encontrada para este filtro.</div>
    <%End if
	
	rs.Close
	Set rs=Nothing
	
	
	%>
    
    	<%Else%>
    	
        
        <%End if%>
      <%FecharConexao()%>
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

(function () {
    var submenuLink = document.querySelector('.ppv-module a[href="ppv_ordens_servico_triagem_inicial.asp"]');
    if (submenuLink) submenuLink.setAttribute('aria-current', 'page');

    var browser = document.getElementById('triage-browser');
    var title = document.getElementById('triage-browser-title');
    var results = document.getElementById('triage-browser-results');
    var back = document.getElementById('triage-browser-back');
    var lastList, controller;
    var titles = { pendentes: 'Triagem pendente', concluidos: 'Triagem completa', historico: 'Hist\u00f3rico de etapas' };

    async function loadPanel(url) {
        var mode = url.searchParams.get('painel_triagem');
        if (mode !== 'historico') lastList = url;
        if (controller) controller.abort();
        controller = new AbortController();
        var signal = controller.signal;
        title.textContent = titles[mode];
        back.hidden = mode !== 'historico';
        results.setAttribute('aria-busy', 'true');
        title.focus();
        results.textContent = 'Carregando\u2026';
        try {
            var response = await fetch(url, { signal: signal, cache: 'no-store' });
            if (!response.ok || response.headers.get('X-Triagem-Painel') !== '1') throw new Error();
            var html = await response.text();
            if (signal.aborted) return;
            results.innerHTML = html;
            results.parentNode.scrollTop = 0;
        } catch (error) {
            if (!signal.aborted) results.textContent = 'N\u00e3o foi poss\u00edvel carregar os dados. Feche o painel e tente novamente.';
        } finally {
            if (!signal.aborted) results.setAttribute('aria-busy', 'false');
        }
    }

    document.querySelectorAll('[data-triage-mode]').forEach(function (button) {
        button.addEventListener('click', function () {
            browser.showModal();
            loadPanel(new URL('?painel_triagem=' + button.dataset.triageMode, location.href));
        });
    });
    browser.addEventListener('close', function () {
        if (controller) controller.abort();
    });
    back.addEventListener('click', function () { loadPanel(lastList); });
    results.addEventListener('click', function (event) {
        var link = event.target.closest('a[data-triage-load]');
        if (!link || event.ctrlKey || event.metaKey || event.shiftKey || event.altKey) return;
        event.preventDefault();
        loadPanel(new URL(link.href));
    });
    results.addEventListener('submit', function (event) {
        event.preventDefault();
        var url = new URL(location.href);
        url.search = new URLSearchParams(new FormData(event.target)).toString();
        loadPanel(url);
    });

    var review = document.getElementById('triage-review');
    if (review) {
        var dialog = document.getElementById('triage-confirm');
        var form = document.getElementById('triage-form');
        review.addEventListener('click', function () {
            dialog.returnValue = '';
            dialog.showModal();
        });
        dialog.addEventListener('close', function () {
            if (dialog.returnValue === 'confirmar') form.submit();
        });
    }
})();
</script>
</body>
</html>
