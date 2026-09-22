<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'/////PÁGINA COM SUB VALIDAÇÃO
Dim VALIDA_SUB1, QUESTIONAR
QUESTIONAR = 2
'5 - Liberação de alteração após a produção estiver finalizada
'Obs.: Sempre maior que 0 (zero) significa que foi encontrada a validação, menor ou igual a zero não foi encontrada
'VALIDA_SUB1 = formatNumber(0&inStr(sub_validacoes_paginas,"#5"),0)
VALIDA_SUB1 = 0
'FAZER A PARTE DE SE O USUÁRIO POSSUIR JÁ ACESSO ESPECIAL, NÃO SERÁ NECESSÁRIO SENHA NOVAMENTE
'If 	VALIDA_SUB1 > 0 Then
'	HABILITADO = 1
'	QUESTIONAR = 1
'Else
'	HABILITADO = 2
'End if
%>


<%
Dim COD_PRODUCAO, PROD_AGREGADO, COD_FORMANDO, NOME_FORMANDO, msg
COD_PRODUCAO	= Request.QueryString("cod_producao")
PROD_AGREGADO	= Request.QueryString("prod_agregado")
COD_FORMANDO	= Request.QueryString("cod_formando")
NOME_FORMANDO	= Request.QueryString("nome_formando")
msg				= Request.QueryString("msg")

'Instruções
'1) Verificar se o usuário logado já e especial (Gerente) e possui habilitação para alteração, caso for apenas informa o do usuário ao invés do login e a opção avançar
'2) Caso o usuário logado não tiver senha especial para está pagina, solicita o login e senha de um gerente especial
'3) Se o login estiver correto reseta a ficha do formando e volta para a página dip_producoes_fichas, entretanto não solicita o nome do vendedor, pois para alterar vendedor deverá ser feito através da ENTRADA e SAíDA de estoque
'4) Mostra um botão gravar no lugar da alteração com senha
'5) Grava no log de usuários que este usuário especial alterou da seguinte forma posteriormente após as alterações fazer um string com o conteúdo completo do formando e depois comparar com o que foi alterado e salvar apenas o que foi modificado no log

If Request.Form("btnAvancar") <> Empty Then
	Dim US, SE, HABILITADO
	US = Request.Form("LOGIN")
	SE = Request.Form("SENHA")
	
	'1) Verifica se o login e senha estão corretos
	'2) Se estiverem verifica se este login informado possui validação para alteração
	
	AbrirConexao()
	Dim ERRO_NAO_LOGAR, COD_ERRO, COD_USUARIO
	ERRO_NAO_LOGAR = 2
		
			US = Replace(US,"'","''")
			US = Replace(US," ","")
			SE = Replace(SE,"'","''")
			SE = Replace(SE," ","")
			
			
			sql = " SELECT TB_ADM_USUARIOS.COD_PERFIL,TB_ADM_USUARIOS.BLOQUEADO, TB_ADM_USUARIOS.PERMITIR_USO_EXTERNO, TB_ADM_USUARIOS.COD_USUARIO, TB_ADM_USUARIOS.NOME, TB_ADM_USUARIOS.LOGIN, TB_ADM_USUARIOS.SENHA,"&_
				  "	TB_ADM_USUARIOS_DEPARTAMENTOS.NOME_DEPTO "&_
				  " FROM TB_ADM_USUARIOS_DEPARTAMENTOS "&_
				  " INNER JOIN TB_ADM_USUARIOS ON TB_ADM_USUARIOS_DEPARTAMENTOS.COD_DEPTO = TB_ADM_USUARIOS.COD_DEPTO"&_
				  " WHERE LOGIN = '"&US&"' AND SENHA = '"&SE&"' AND TB_ADM_USUARIOS.BLOQUEADO=2"
				  
				  
			Set rs = objConexao.execute(sql)
		
			If Not rs.eof Then

				COD_USUARIO	= rs("COD_USUARIO")
				COD_PERFIL	= rs("COD_PERFIL")
				NOME		= rs("NOME")
				
				
					sql="SELECT "&_
						"TB_ADM_PAGINAS_HABILITADAS_SUB_VALIDACOES.COD_SUB_VALIDACAO, "&_
						"TB_ADM_PAGINAS_HABILITADAS_SUB_VALIDACOES.NOME_SUB_VALIDACAO "&_
						"FROM TB_ADM_PAGINAS_HABILITADAS_SUB_VALIDACOES INNER JOIN "&_
						"TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES ON "&_
						"TB_ADM_PAGINAS_HABILITADAS_SUB_VALIDACOES.COD_SUB_VALIDACAO = TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES.COD_SUB_VALIDACAO WHERE "&_
						"TB_ADM_PERFIS_PAGINAS_HABILITADAS_SUB_VALIDACOES.COD_PERFIL="&COD_PERFIL&" AND TB_ADM_PAGINAS_HABILITADAS_SUB_VALIDACOES.COD_PAGINA="&COD_PAGINA_USUARIO
						
					sub_validacoes_paginas = 0
					Set rsAdmSubValidacoes = objConexao.Execute(sql)
					If not rsAdmSubValidacoes.Eof Then

						While not rsAdmSubValidacoes.Eof
						
							sub_validacoes_paginas = sub_validacoes_paginas&"#"&rsAdmSubValidacoes("COD_SUB_VALIDACAO")
		 
						rsAdmSubValidacoes.MoveNext()
						Wend
						
					End If
					
				 rsAdmSubValidacoes.Close
				 Set rsAdmSubValidacoes = Nothing
					
				'///////////////////
				
				
				VALIDA_SUB1 = formatNumber(0&inStr(sub_validacoes_paginas,"#5"),0)	
						'Response.Write("Validado ="&sub_validacoes_paginas)
						If 	VALIDA_SUB1 <> 0 Then
								HABILITADO = 1
						Elseif VALIDA_SUB1 = 0 Then
								ERRO_NAO_LOGAR = 1
						End If				
				
				
				'//////////////////	
					
					
					
					
					
					
								
			Else
					'1)
					ERRO_NAO_LOGAR = 1
					COD_ERRO 	   = 1
						 
			End If
				
				  
			
			rs.Close
			Set rs=Nothing 

			
			If ERRO_NAO_LOGAR = 2 Then
			'///////////////////////////////////////////////////
								 HABILITADO = 1
				'/Grava tudo o que tinha antes
				
				objConexao.Execute( "INSERT INTO TB_ADM_LOG_ALTERACOES ("&_
									"COD_USUARIO ,"&_
									"DESCRICAO "&_
									")"&_
									"VALUES ("&_
									Session("COD_USUARIO")&", 'Formando  ("&digito(COD_FORMANDO)&") "&NOME_FORMANDO&" foi alterado sob permissão do usuário: "&NOME&", posteriormente a produção finalizada.'"&_
									")" )
							
				
				'/Reseta formando
				'1) Resetar Ficha de Identificação
				objConexao.Execute("UPDATE TB_FICHAS_IDENTIFICACAO SET "&_
									"ALBUM = 	0, "&_
									"INICIAL = 	0, "&_
									"ENCADERNACOES_EXTRAS = 0, "&_
									"FATURADO = 0, "&_
									"RETIRADO = 0, "&_
									"VALOR_UN = 0, "&_
									"BRINDE = 	2, "&_
									"COMISSAO = 2, "&_
									"PRODUTO_AGREGADO = 2 "&_
									"WHERE COD_FORMANDO = "&COD_FORMANDO&" LIMIT 1")
			
				'2) Exclui brindes se houver.
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_BRINDES WHERE COD_FORMANDO="&COD_FORMANDO)
				
							
				'3) Exclui produtos agregados se houver;
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_PRODUTOS_AGREGADOS WHERE COD_FORMANDO="&COD_FORMANDO)	
				
				'4) Exclui tabela temporária se houver;
				objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_TEMPORARIA WHERE COD_FORMANDO="&COD_FORMANDO)	
			'///////////////////////////////////////////////////

			FecharConexao()	 
				 
			Else
				Response.Redirect("dip_producoes_fichas_POPUP_confirmacao_alteracao_senha_especial.asp?cod_producao="&COD_PRODUCAO&"&prod_agregado="&PROD_AGREGADO&"&cod_formando="&COD_FORMANDO&"&nome_formando="&NOME_FORMANDO&"&msg=1")
			End If
			

End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Cadastro de Produto em pedido</title>
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<style type="text/css">
<!--
body {
	background-color: #FFFFFF;
}
-->
</style></head>
<!--INICIO MASCARAS-->
<script language="JavaScript">
// JavaScript Document
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
function ValidaData(data){
    exp = /\d{2}\/\d{2}\/\d{4}/
    if(!exp.test(data.value))
        alert('Data Invalida!');            
}

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
<!--FIM MASCARAS-->
<script>

function mascara(o,f){
    v_obj=o
    v_fun=f
    setTimeout("execmascara()",1)
}

function execmascara(){
    v_obj.value=v_fun(v_obj.value)
}

function leech(v){
    v=v.replace(/o/gi,"0")
    v=v.replace(/i/gi,"1")
    v=v.replace(/z/gi,"2")
    v=v.replace(/e/gi,"3")
    v=v.replace(/a/gi,"4")
    v=v.replace(/s/gi,"5")
    v=v.replace(/t/gi,"7")
    return v
}

function soNumeros(v){
    return v.replace(/\D/g,"")
}

function telefone(v){
    v=v.replace(/\D/g,"")
    v=v.replace(/^(\d\d)(\d)/g,"($1) $2") 
    v=v.replace(/(\d{4})(\d)/,"$1-$2")   
    return v
}

function cpf(v){
    v=v.replace(/\D/g,"")                   
    v=v.replace(/(\d{3})(\d)/,"$1.$2")      
    v=v.replace(/(\d{3})(\d)/,"$1.$2")      
                                            
    v=v.replace(/(\d{3})(\d{1,2})$/,"$1-$2")
    return v
}

function cep(v){
    v.replace(/\D/g,"")               
    v=v.replace(/^(\d{5})(\d)/,"$1-$2") 
    return v
}

function cnpj(v){
    v=v.replace(/\D/g,"")                           
    v=v.replace(/^(\d{2})(\d)/,"$1.$2")             
    v=v.replace(/^(\d{2})\.(\d{3})(\d)/,"$1.$2.$3") 
    v=v.replace(/\.(\d{3})(\d)/,".$1/$2")         
    v=v.replace(/(\d{4})(\d)/,"$1-$2")              
    return v
}

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

<script language="javascript">
function focusIni(){
	document.validacao.LOGIN.focus();
}
function carregaFormulario(){

	document.VOLTAR.submit();
}
</script>

<body onload="<%If HABILITADO = 1 Then Response.Write("carregaFormulario()") Else Response.Write("focusIni()") End If%>">

  <%If HABILITADO = 1 Then%>
	<form method="post" name="VOLTAR" action="dip_producoes_fichas.asp?cod_prod=<%=COD_PRODUCAO%>&prod_agre=<%=PROD_AGREGADO%>&cod_form=<%=COD_FORMANDO%>&msg=Ficha do(a) formando(a): <%=NOME_FORMANDO%> foi resetada com sucesso!" target="_parent" style="margin:0px">
		
		  <input name="VALIDA_USER" value="1" type="hidden" />
		<div align="center">
		  <p>&nbsp;</p>
		  <p><strong>Aguarde processando...</strong></p>
		  <p><br />
		    <br />
	        </p>
		</div>
	</form>
  <%Else%>

<form name="validacao" style="margin:0px" action="?cod_producao=<%=COD_PRODUCAO%>&prod_agregado=<%=PROD_AGREGADO%>&cod_formando=<%=COD_FORMANDO%>&nome_formando=<%=NOME_FORMANDO%>" onsubmit="return verificaForm()" method="post">
<br />
<table width="402" border="0" align="center" cellpadding="2" cellspacing="2">
<tr>
<td colspan="2"><div align="center" class="msg">Aten&ccedil;&atilde;o ao avan&ccedil;ar ser&aacute; relacionada a modifica&ccedil;&atilde;o<br />no relat&oacute;rio de inconsist&ecirc;ncias.</div></td>
</tr>
		<%If VALIDA_SUB1 <= 0 Then%>
		
		 <tr>
              <td width="89">
        
           <div align="left"><strong>Login:</strong></div>             </td>
              <td width="299"><div align="left">
                <input name="LOGIN" type="text" size="40" maxlength="80" />
           </div></td>
    </tr>
			<tr>
              <td>
              <div align="left"><strong>Senha:</strong></div>              </td>
              <td><div align="left">
                <input name="SENHA" type="password" size="40" maxlength="15" />
              </div></td>
            </tr>
			
			
		<%Elseif VALIDA_SUB1 > 0 Then%>
		<%End If%>
 <tr>
              <td colspan="2"><div align="center">
                <p style="margin:0px"><%
				If Request.QueryString("msg") <> Empty Then
				Response.Write("Erro: usuário ou senha inválidas / usuário não tem <br />permissão de acesso para esta alteração")
				End if
				%></p>
                <p style="margin:0px">
                  <input name="btnAvancar" class="btnAvancar" type="submit" id="btnAvancar" value=" " />
                </p>
                <p style="margin:0px">&nbsp;</p>
              </div></td>
    </tr>
  </table>
</form>
<%End If%>


</body>
</html>
