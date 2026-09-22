<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<%
'######################################################################
'INSTRU��ES CORRE��O DE PEDIDOS
'1) Lan�amento da DATA da VENDa // Muito importante, pois � calculado dados como quando o vendedores receber�
'2) Corre��o de Produtos Lan�ados e Calculo se o �lbum � repasse ou n�o, caso for ser� concedito um acr�scimo na comiss�o totalizando 10% do valor do recibo, independente se o pedido n�o conter somente o �lbum
'3) Corre��o de Forma de Pagamento
'4) Inser��o de Ajuda de Custo, Visualiza��o de Quando o vendedor receber�, Observa��o
'5) Conclus�o confirma��o e finaliza��o
'######################################################################

'*** Instru��es
'1) Seleciona o vendedor e a data venda, ap�s a data da venda estar corrigida avan�a para efetuar os pr�ximos c�lculos.s
'2) Faz o uptade, caso a data estiver errada no sistema.


function excluirPedido(cod_pedido, conexao_aberta) 

  if not conexao_aberta then
    abrirConexao()
  end if 

  dim retorno : retorno = false
  dim pedido_extra, pagamento_processado

  PAGAMENTO_PROCESSADO = 2 
  Set rs = objConexao.Execute("SELECT PEDIDO_EXTRA, PAGAMENTO_PROCESSADO FROM TB_PEDIDOS WHERE COD_PEDIDO ="&COD_PEDIDO&temp)
  If not rs.Eof Then
    pedido_extra = rs("PEDIDO_EXTRA").value
    pagamento_processado = rs("PAGAMENTO_PROCESSADO").value
  End If
  rs.Close
  Set rs = Nothing

  if pagamento_processado = 2 then
    ' LOG_TB_PEDIDOScod_usuario_excluiu
    campos = " COD_PEDIDO, PEDIDO_EXTRA, SISTEMA_NOVO, DATA_VENDA, COD_VENDEDOR, COD_GRUPO_VENDEDORES, DATA_CADASTRO, DATA_CHEGADA, FINALIZA_VENDEDOR, FINALIZA_CORRECAO, FINALIZA_DIRETORIA, DATA_CORRIGIDO_DEP_FOTO, DATA_DESCEU_DIRETORIA, DATA_CORRIGIDO_DIRETORIA, OBSERVACAO, VALOR_AJUDA_CUSTO, VENDEDOR_RECEBE_EM, PAGAMENTO_EM_DINHEIRO, PAGAMENTO_PROCESSADO, FORMA_PAGAMENTO, QTD_PARCELAS, TIPO_VENDA, ERRO_LISTA, CONFERIDO_POR_COD_USUARIO, cod_fechamento_periodo, cod_sacado, boleto_enviado "
    objConexao.execute("INSERT INTO LOG_TB_PEDIDOS ("&campos&", cod_usuario_excluiu, data_excluido) SELECT "&campos&", "&Session("cod_usuario")&", now() FROM TB_PEDIDOS WHERE cod_pedido = "&sqlNumero(cod_pedido))

    ' LOG_TB_PEDIDOS_BOLETOS
    campos = "cod_boleto, cod_pedido, cod_sacado, cod_vendedor, parcela, nosso_numero, hash, valor, valor_pago, valor_deveria_ser_pago, data_vencimento, data_pagamento, data_baixa, baixa_manual, baixa_manual_usuario, valor_comissao_vendedor, vendedor_recebe_em, cod_fechamento_periodo, pagamento_processado "
    objConexao.execute("INSERT INTO LOG_TB_PEDIDOS_BOLETOS ("&campos&") SELECT "&campos&" FROM TB_PEDIDOS_BOLETOS WHERE cod_pedido = "&sqlNumero(cod_pedido))

    ' LOG_TB_PEDIDOS_CARTAO
    campos = "COD_PEDIDO_CARTAO, COD_PEDIDO, CARTAO, VALOR_TOTAL, PARCELAS, NOME_TITULAR, GRAU_PARENTESCO, AUTENTICACAO, DOCUMENTO, DATA_AUTENTICACAO, VERIFICADO, DESMEMBRADO, DESMEMBRADO_EM"
    objConexao.execute("INSERT INTO LOG_TB_PEDIDOS_CARTAO ("&campos&") SELECT "&campos&" FROM TB_PEDIDOS_CARTAO WHERE cod_pedido = "&sqlNumero(cod_pedido))

    ' LOG_TB_PEDIDOS_CARTAO_DESMEMBRADO
    campos = "ID, COD_PEDIDO_CARTAO, COD_PEDIDO, COD_CARTAO, DOCUMENTO, AUTENTICACAO, DATA_DOCUMENTO, PARCELA, VALOR_PARCELA, VALOR_PREVISTO, VALOR_CREDITO, DT_VENCIMENTO, DT_PAGAMENTO, TAXA, DATA_DESMEMBRADO"
    objConexao.execute("INSERT INTO LOG_TB_PEDIDOS_CARTAO_DESMEMBRADO ("&campos&") SELECT "&campos&" FROM TB_PEDIDOS_CARTAO_DESMEMBRADO WHERE cod_pedido = "&sqlNumero(cod_pedido))

    ' LOG_TB_PEDIDOS_CHEQUE
    campos = "COD_PEDIDO_CHEQUE, COD_PEDIDO, BANCO, AGENCIA, CONTA_CORRENTE, NUMERO_CHEQUE, VALOR, BOM_PARA, NOME, CPF, CNPJ, SE_TERCEIRO, FONE, VERIFICADO"
    objConexao.execute("INSERT INTO LOG_TB_PEDIDOS_CHEQUE ("&campos&") SELECT "&campos&" FROM TB_PEDIDOS_CHEQUE WHERE cod_pedido = "&sqlNumero(cod_pedido))

    ' LOG_TB_PEDIDOS_CHEQUE_HISTORICO
    campos = "cod_pedido_cheque_historico, cod_pedido_cheque, historico, data_cadastro, cod_usuario"
    objConexao.execute("INSERT INTO LOG_TB_PEDIDOS_CHEQUE_HISTORICO ("&campos&") SELECT "&campos&" FROM TB_PEDIDOS_CHEQUE_HISTORICO WHERE cod_pedido_cheque in ( select COD_PEDIDO_CHEQUE from LOG_TB_PEDIDOS_CHEQUE where cod_pedido = "&sqlNumero(cod_pedido)) &")"

    ' LOG_TB_PEDIDOS_FORMANDOS
    campos = "COD_PEDIDO, COD_FORMANDO"
    objConexao.execute("INSERT INTO LOG_TB_PEDIDOS_FORMANDOS ("&campos&") SELECT "&campos&" FROM TB_PEDIDOS_FORMANDOS WHERE cod_pedido = "&sqlNumero(cod_pedido))

    
    If int(PEDIDO_EXTRA) = 2 Then
      'Limpa as fichas dos formandos caso n�o for pedido extra		
      Set rsDel = objConexao.Execute("SELECT * FROM TB_PEDIDOS_FORMANDOS WHERE COD_PEDIDO ="&COD_PEDIDO)
      While not rsDel.Eof
        sql =	" UPDATE TB_FICHAS_IDENTIFICACAO SET FATURADO = '0',VALOR_UN = '0',RETIRADO = '0', VENDIDO = '2' WHERE COD_FORMANDO = "&rsDel ("COD_FORMANDO")&" LIMIT 1"	
        objConexao.Execute(sql)							
        objConexao.Execute("UPDATE TB_PRODUCAO_LOTES_FORMANDOS SET STATUS = 3 WHERE COD_FORMANDO = '"&rsDel("COD_FORMANDO")&"'")
      rsDel.MoveNext()
      Wend
      rsDel.Close
      Set rsDel = Nothing
    End If

    objConexao.Execute("DELETE FROM TB_ORDENS_SERVICO_ETAPAS WHERE cod_ordem_servico IN(SELECT cod_ordem_servico FROM TB_ORDENS_SERVICO WHERE cod_pedido = "&COD_PEDIDO&")")		
    objConexao.Execute("DELETE FROM TB_ORDENS_SERVICO_HISTORICO WHERE cod_ordem_servico IN(SELECT cod_ordem_servico FROM VW_TB_ORDENS_SERVICO WHERE cod_pedido = "&COD_PEDIDO&")")
    objConexao.Execute("DELETE FROM TB_ORDENS_SERVICO_TEMPORARIA WHERE cod_ordem_servico IN(SELECT cod_ordem_servico FROM VW_TB_ORDENS_SERVICO WHERE cod_pedido = "&COD_PEDIDO&")")
    objConexao.Execute("DELETE FROM TB_ORDENS_SERVICO WHERE cod_pedido = "&cod_pedido)

    objConexao.Execute("DELETE FROM TB_PEDIDOS WHERE COD_PEDIDO = "&COD_PEDIDO&" LIMIT 1")
    objConexao.Execute("DELETE FROM TB_PEDIDOS_CARTAO WHERE COD_PEDIDO = "&COD_PEDIDO)
    objConexao.Execute("DELETE FROM TB_PEDIDOS_CHEQUE WHERE COD_PEDIDO = "&COD_PEDIDO)
    objConexao.Execute("DELETE FROM TB_PEDIDOS_FORMANDOS WHERE COD_PEDIDO = "&COD_PEDIDO)
    objConexao.Execute("DELETE FROM TB_PEDIDOS_PRODUTOS WHERE COD_PEDIDO = "&COD_PEDIDO)
    objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO_TEMPORARIA WHERE COD_PEDIDO = "&COD_PEDIDO)

    retorno = true
  end if

  if not conexao_aberta then
    fecharConexao()
  end if 

excluirPedido = retorno
end function


Dim COD_PEDIDO, QTD_PARCELAS
COD_PEDIDO = Request.QueryString("cod_pedido")

'////////////////
If Request.QueryString("comando") = "excluir" Then

  pedido_excluido = excluirPedido(COD_PEDIDO, false)
	
	if (pedido_excluido) then
	  Response.Redirect("lca_pedidos_pendentes.asp?msg=Pedido excluído com sucesso!")
  else
    Response.Redirect("lca_pedidos_pendentes.asp?msg=Pedido não pode ser excluído")
  end if 
End If

Dim rs, rs2,rs3, sql
Dim DATA_VENDA, COD_VENDEDOR, DATA_CADASTRO, OBSERVACAO, VALOR_AJUDA_CUSTO, PAGAMENTO_EM_DINHEIRO,VENDEDOR_RECEBE_EM, NOME_VENDEDOR, DATA_DESCEU_DIRETORIA, PAGAMENTO_PROCESSADO
Dim DATA_REPASSE

sql = 	"SELECT "&_
		"TB_PEDIDOS.DATA_VENDA, "&_
		"TB_PEDIDOS.QTD_PARCELAS, "&_
		"TB_PEDIDOS.DATA_CADASTRO, "&_
		"TB_PEDIDOS.DATA_DESCEU_DIRETORIA, "&_
		"TB_PEDIDOS.FINALIZA_DIRETORIA, "&_
		"TB_PEDIDOS.COD_VENDEDOR, "&_
		"TB_PEDIDOS.OBSERVACAO, "&_
		"TB_PEDIDOS.PAGAMENTO_EM_DINHEIRO, "&_
		"TB_PEDIDOS.VALOR_AJUDA_CUSTO, "&_
		"TB_PEDIDOS.PAGAMENTO_PROCESSADO, "&_
		"TB_PEDIDOS.VENDEDOR_RECEBE_EM, "&_
		"TB_VENDEDORES.NOME_VENDEDOR "&_
		"FROM TB_VENDEDORES INNER JOIN TB_PEDIDOS ON TB_VENDEDORES.COD_VENDEDOR = TB_PEDIDOS.COD_VENDEDOR "&_
		"WHERE TB_PEDIDOS.COD_PEDIDO ="&COD_PEDIDO
		
		AbrirConexao()
		
		Set rs = objConexao.Execute(sql)
			
		  If not rs.Eof Then	
			DATA_VENDA			= rs("DATA_VENDA")
			COD_VENDEDOR		= rs("COD_VENDEDOR")
			DATA_CADASTRO		= rs("DATA_CADASTRO")
			OBSERVACAO			= rs("OBSERVACAO")
			VALOR_AJUDA_CUSTO	= rs("VALOR_AJUDA_CUSTO")
			VENDEDOR_RECEBE_EM	= rs("VENDEDOR_RECEBE_EM")
			NOME_VENDEDOR		= rs("NOME_VENDEDOR")
			PAGAMENTO_EM_DINHEIRO = rs("PAGAMENTO_EM_DINHEIRO")
			QTD_PARCELAS		= formatNumber(0&rs("QTD_PARCELAS"),0)
			DATA_DESCEU_DIRETORIA = Cstr(""&rs("DATA_DESCEU_DIRETORIA"))
			PAGAMENTO_PROCESSADO = int(rs("PAGAMENTO_PROCESSADO"))
			FINALIZA_DIRETORIA = int(0&rs("FINALIZA_DIRETORIA"))
			
			If int(QTD_PARCELAS) = 0 Then
				QTD_PARCELAS = ""
			End If
		  End If
		  
		rs.Close()
		Set rs = Nothing
		
		FecharConexao()

If DATA_DESCEU_DIRETORIA <> "" and Session("cod_usuario") <> 16 Then
	Response.Redirect("adm_index.asp?codmsg=1")
End If
	
'////////////////////////////////
Dim comando, botao, btnGravar
comando = Request.QueryString("comando")
botao   = Request.Form("btnAvancar1")
btnGravar = Request.Form("btnGravar")

If Request.Form("btnGravar") <> "" Then
  Dim CARTAO, PARCELAS, DATA_AUTENTICACAO
  CARTAO = Trim(Request.Form("CARTAO"))
  PARCELAS = Trim(Request.Form("PARCELAS"))
  DATA_AUTENTICACAO = Trim(Request.Form("DATA_AUTENTICACAO"))

  If CARTAO <> "" And PARCELAS <> "" And DATA_AUTENTICACAO <> "" Then
    AbrirConexao()

    sql = "UPDATE TB_PEDIDOS_CARTAO SET " & _
          "CARTAO = '" & Replace(CARTAO, "'", "''") & "', " & _
          "PARCELAS = " & sqlNumero(PARCELAS) & ", " & _
          "DATA_AUTENTICACAO = " & sqlData(DATA_AUTENTICACAO) & " " & _
          "WHERE COD_PEDIDO = " & sqlNumero(COD_PEDIDO)

    ' Debug opcional: descomente para ver o SQL gerado
    ' Response.Write "<!-- SQL DEBUG: " & sql & " -->" : Response.End

    objConexao.Execute sql
    FecharConexao()

    Response.Redirect "lca_pedidos_pendentes_alteracoes_part_1.asp?cod_pedido=" & COD_PEDIDO & "&msg=Cartão atualizado com sucesso!"
  Else
    Response.Redirect "lca_pedidos_pendentes_alteracoes_part_1.asp?cod_pedido=" & COD_PEDIDO & "&msg=Preencha todos os campos do cartão para gravar."
  End If
End If


If comando = "alterar" and botao <> Empty Then
'Faz o UPDATE
	Dim DATA_VENDA_VALIDA, DATA_ATUAL, QTD_DIAS
	DATA_VENDA_VALIDA = ""&Request.Form("DATA_VENDA")
	DATA_ATUAL = Date()
	
	QTD_DIAS = DateDiff("d",DATA_VENDA_VALIDA,DATA_ATUAL)
	
	If int(QTD_DIAS) => 0 Then
	
		AbrirConexao()
		objConexao.Execute("UPDATE TB_PEDIDOS SET DATA_VENDA = "&sqlData(Request.Form("DATA_VENDA"))&", QTD_PARCELAS = '"&Request.Form("QTD_PARCELAS")&"' WHERE COD_PEDIDO="&COD_PEDIDO&" LIMIT 1")
		FecharConexao()

		Response.Redirect("lca_pedidos_pendentes_alteracoes_part_2.asp?cod_pedido="&cod_pedido)
	
	Else
		
		Response.Redirect("lca_pedidos_pendentes_alteracoes_part_1.asp?cod_pedido="&COD_PEDIDO&"&msg=Aten&ccedil;&atilde;o a data da venda &eacute; superior a data atual, corriga antes de continuar.")
		
	End If

End If
		
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<title>Altera&ccedil;&otilde;es de Pedidos</title>
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
	<!--AJUDA F1 E CTRL PARA BUSCA R�PIDA-->
		
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
	</script>
	
	<script type="text/javascript">
	DHTMLSuite.include("calendar");
	</script>
	<script type="text/javascript">
	var calendarObjForForm = new DHTMLSuite.calendar({minuteDropDownInterval:10,numberOfRowsInHourDropDown:5,callbackFunctionOnDayClick:'getDateFromCalendar',isDragable:true,displayTimeBar:false});
	
		 
	calendarObjForForm.setCallbackFunctionOnClose('myOtherFunction');
	
	function myOtherFunction()
	{
		
		
	}
	function pickDate(buttonObj,inputObject)
	{
		calendarObjForForm.setCalendarPositionByHTMLElement(inputObject,0,inputObject.offsetHeight+2);	// Position the calendar right below the form input
		//calendarObjForForm.setInitialDateFromInput(inputObject,'yyyy-mm-dd hh:ii');	// Specify that the calendar should set it's initial date from the value of the input field.
		
		calendarObjForForm.setInitialDateFromInput(inputObject,'dd/mm/yyyy');	// Specify that the calendar should set it's initial date from the value of the input field.
		
		calendarObjForForm.addHtmlElementReference('myDate',inputObject);	// Adding a reference to this element so that I can pick it up in the getDateFromCalendar below(myInput is a unique key)
		if(calendarObjForForm.isVisible()){
			calendarObjForForm.hide();
		}else{
			calendarObjForForm.resetViewDisplayedMonth();	// This line resets the view back to the inital display, i.e. it displays the inital month and not the month it displayed the last time it was open.
			calendarObjForForm.display();
		}		
	}	
	/* inputArray is an associative array with the properties
	year
	month
	day
	hour
	minute
	calendarRef - Reference to the DHTMLSuite.calendar object.
	*/
	function getDateFromCalendar(inputArray)
	{
		var references = calendarObjForForm.getHtmlElementReferences(); // Get back reference to form field.
		references.myDate.value = inputArray.day + '/' + inputArray.month + '/' + inputArray.year
//		references.myDate.value = inputArray.year + '-' + inputArray.month + '-' + inputArray.day + ' ' + inputArray.hour + ':' + inputArray.minute;
		calendarObjForForm.hide();	
		
	}	
	</script>

<script>
function confirma() {
	
	<%
	autoriza_excluir = true
	
	If DATA_DESCEU_DIRETORIA <> "" and not validaUsuario(2,"",50,false) or PAGAMENTO_PROCESSADO = 1 Then
	'Já desceu para diretoria, por tanto verifica se pode o usuário poderá excluir
		autoriza_excluir = false
	End If
	%>
	
	<%If autoriza_excluir Then%>
		if(YNconfirm("Tem certeza que deseja excluir este pedido?")){
			abre('lca_pedidos_pendentes_alteracoes_part_1.asp?cod_pedido=<%=COD_PEDIDO%>&comando=excluir')		
		}
	<%Else%>
		alert('Pedido já foi baixado para diretoria e/ou já foi realizado o pagamento, para excluir é necessário ter a seguinte validação especial:\nExclusão de pedidos Enviados para Diretoria, porém se a venda já foi paga para o consultor não poderá ser excluída.')
	<%End If%>	
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
function validarData(campo){
var expReg = /^(([0-2]\d|[3][0-1])\/([0]\d|[1][0-2])\/[1-2][0-9]\d{2})$/;
var msgErro = 'Formato inv�lido de data.';
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
</script> 
<script type="text/javascript">
function validaForm(){
	
	var d=document.datasVenda;
	if(d.QTD_PARCELAS.value == '' || d.QTD_PARCELAS.value == '0'){
		alert('Informe a quantidade de parcelas antes de continuar.')
		d.QTD_PARCELAS.focus();
		return false
	}else{
		return true
	}
}


</script>
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
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  %>
  <!-- #include file="includes/sub_menu.asp" -->
  
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top" style="border:dotted 1px #999999">

	  <h2>LCA - Pedidos Pendentes / Corre&ccedil;&atilde;o </h2>
    <% If  validaUsuario(2,"",85,false) or validaUsuario(2,"",86,false) Then %>
	    <div class="subMenu">Sub-Menu: <a href="javascript:Void(0)" onclick="confirma()">Excluir Pedido</a></div>
    <%end if%>
	   <div class="msg" align="center"><%=Request.QueryString("msg")%></div>
	  	  
		  	<h3 align="center">Corre&ccedil;&atilde;o de Pedido - Dados B&aacute;sicos </h3>
			
			<table width="98%" align="center" 	cellpadding="0" cellspacing="3">
			<tr>
            <td width="12%"><div align="left"><strong>Vendedor</strong>:</div></td>
            <td width="88%"><%
			AbrirConexao()
			'Nome do vendedor
			Dim rsVendedorQ
			Set rsVendedorQ = objConexao.Execute("SELECT TB_VENDEDORES.NOME_VENDEDOR FROM TB_VENDEDORES INNER JOIN TB_PEDIDOS ON TB_VENDEDORES.COD_VENDEDOR = TB_PEDIDOS.COD_VENDEDOR WHERE TB_PEDIDOS.COD_PEDIDO="&COD_PEDIDO&" LIMIT 1")
				
				If not rsVendedorQ.Eof Then
					Response.Write(rsVendedorQ("NOME_VENDEDOR"))
				End If
			
			rsVendedorQ.Close
			Set rsVendedorQ = Nothing
			
			%></td>
            </tr>
			<tr>
			  <td><strong>Formando(s):</strong></td>
			  <td><%
			sql = 	"SELECT TB_PEDIDOS_FORMANDOS.COD_FORMANDO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO "&_
					"FROM TB_PEDIDOS_FORMANDOS INNER JOIN TB_FICHAS_IDENTIFICACAO ON "&_
					"TB_PEDIDOS_FORMANDOS.COD_FORMANDO = TB_FICHAS_IDENTIFICACAO.COD_FORMANDO "&_
					"WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&COD_PEDIDO&" ORDER BY NOME_FORMANDO"
			Set rs3 = objConexao.Execute(sql)
				
				While not rs3.Eof
				%>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				  <tr>
					<td><div align="left"><%=rs3("NOME_FORMANDO")%> (<%=digito(rs3("COD_FORMANDO"))%>) </div></td>
					</tr>
				</table>
				<%
				rs3.MoveNext()
				Wend
				
			rs3.Close
			Set rs3 = Nothing
			FecharConexao()
			%></td>
			  </tr>
			</table>
		  	<br />
			 <form onsubmit="return validaForm()" method="post" action="lca_pedidos_pendentes_alteracoes_part_1.asp?cod_pedido=<%=COD_PEDIDO%>&comando=alterar" style="margin:0px;" name="datasVenda">
			   <table width="98%" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
			 
			   <tr>
            <td width="16%"><div align="left"><strong>Data da Venda</strong></div></td>
            <td width="84%">

			<table width="98%" border="0" cellspacing="0" cellpadding="0">
			<% If validaUsuario(2,"",79,false) or validaUsuario(2,"",85,false) Then %>
			<tr>
				<td width="14%">
				<input class="campoData" type="text" onkeypress="MascaraData(datasVenda.DATA_VENDA);" name="DATA_VENDA" value="<%=ArrumaData(DATA_VENDA)%>"  maxlength="10" /></td>
				<td width="86%" class="textoMenor">(dd/mm/aaaa)</td>
			</tr>
			<% else %>
			<tr>
				<td width="14%">
				<span> <%=ArrumaData(DATA_VENDA)%></span>
				<input class="campoData" type="hidden"  name="DATA_VENDA" value="<%=ArrumaData(DATA_VENDA)%>"  maxlength="10" /></td>
				<td width="86%" class="textoMenor">(dd/mm/aaaa)</td>
			</tr>
			<% End If %>
            </table>
			
			</td>
			</tr>
            <tr>
            <td width="16%"><div align="left"><strong>Qtd de Parcelas</strong></div></td>
            <td>
			<% If  validaUsuario(2,"",85,false) Then %>
			<input class="inputTamanho1" onkeypress="mascara(this,soNumeros)" type="text" name="QTD_PARCELAS" value="<%=QTD_PARCELAS%>"  maxlength="10" />
			<% else %>
			<%=QTD_PARCELAS%>
			<input class="campoData" type="hidden"  name="QTD_PARCELAS" value="<%=QTD_PARCELAS%>"  maxlength="10" /></td>
			<% end if %>
			</td></tr>
			<!--CAMPO DE CARTÃO-->
			<tr>
			<td width="16%"><div align="left"><strong>Cartão</strong></div></td>
			<td>
				<% If  validaUsuario(2,"",85,false) Then %>
				<input class="inputTamanho1" type="text" name="CARTAO" value="<%=Request.Form("CARTAO")%>" maxlength="20" />
				<% Else %>
				<%=Request.Form("CARTAO")%>
				<input type="hidden" name="CARTAO" value="<%=Request.Form("CARTAO")%>" />
				<% End If %>
			</td>
			</tr>

			<tr>
  <td></td>
  <td>
    <div style="display: flex; gap: 10px; align-items: center;">
      <input class="btnAvancar" type="submit" name="btnAvancar1" value=" " />
      <input class="btnGravar" type="submit" name="btnGravar" value=" " style="background-image: url('images/layout/btnGravar.jpg'); background-repeat: no-repeat; padding-left: 18px; width: 80px; height: 28px;" />
    </div>
  </td>
</tr>
			</table>
			 </form>
			 <h3 align="center">Forma de Pagamento Apenas Para Consulta e verifica&ccedil;&atilde;o da Qtd. de Parcelas</h3>
			<table width="98%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
					<tr bgcolor="#E4E4E4">
                        <td><div align="center" class="textoMedio"><strong>Dinheiro</strong></div>                          
                          <div align="center"></div>                          <div align="center"></div>                          <div align="center"></div>                          <div align="center"></div>                          <div align="center"></div></td>
                      </tr>             
					<tr>
						<td><div align="left"><table width="210" border="0" cellspacing="0" cellpadding="0">
						  <tr>
							<td width="62%"><div align="left" class="textoMedio">Total em Dinheiro:</div></td>
							<td width="38%"><div align="right" class="textoMedio"><strong><%=formatNumber(PAGAMENTO_EM_DINHEIRO)%></strong></div></td>
						  </tr>
						</table>
						  </div></td>
						</tr>
                    </table>	
                    <%
					AbrirConexao()
					
					Set rsCartao = objConexao.Execute("SELECT * FROM TB_PEDIDOS_CARTAO WHERE COD_PEDIDO="&COD_PEDIDO&" ORDER BY NOME_TITULAR") 
					
					if Not rsCartao.Eof then
					%>
                    <table width="98%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
					<tr bgcolor="#E4E4E4">
                        <td colspan="6"><div align="center" class="textoMedio"><strong>Cart&atilde;o de Cr&eacute;dito/D&eacute;bito</strong></div>                          
                          <div align="center"></div>                          <div align="center"></div>                          <div align="center"></div>                          <div align="center"></div>                          <div align="center"></div></td>
                      </tr>
                      <tr>
                        <td width="25%"><div align="center"><strong>Nome Titular </strong></div></td>
                        <td width="15%"><div align="center"><strong>Cart&atilde;o</strong></div></td>
                        <td width="13%"><div align="center"><strong>Autentica&ccedil;&atilde;o </strong></div></td>
                        <td width="11%"><div align="center"><strong>Documento</strong></div></td>
                        <td width="10%"><div align="center"><strong>Parcelas </strong></div></td>
                        <td><div align="center"><strong>Valor Total </strong></div></td>
                        </tr>
                      <%
					  While not rsCartao.Eof
					  %>
                      <tr>
                        <td><div align="left"><%=rsCartao("NOME_TITULAR")%></div></td>
                        <td><div align="center"><%=rsCartao("CARTAO")%></div></td>
                        <td><div align="center"><%=rsCartao("AUTENTICACAO")%></div></td>
                        <td><div align="center"><%=rsCartao("DOCUMENTO")%></div></td>
                        <td><div align="center"><%=rsCartao("PARCELAS")%></div></td>
                        <td><div align="right"><%=formatNumber(rsCartao("VALOR_TOTAL"))%>
						<%
						PARCIAL_CARTAO=0
						PARCIAL_CARTAO=formatNumber(rsCartao("VALOR_TOTAL"))
						TOTAL_CARTAO=(TOTAL_CARTAO+PARCIAL_CARTAO)/1
						PARCIAL_CARTAO=0						
						%>
                        </div></td>
                        </tr>
                      <%
					  rsCartao.MoveNext
					  Wend
					  
					  %>
                      <tr>
                        <td colspan="6">			
						<div align="left">
						<table width="210" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="62%"><div align="left" class="textoMedio">Total em Cart&atilde;o:</div></td>
    <td width="38%"><div align="right" class="textoMedio"><strong><%=formatNumber(TOTAL_CARTAO)%></strong></div></td>
  </tr>
</table>  </div>                          </td>
                        </tr>
                    </table>
                    <%
					End If
					
					rsCartao.Close
					Set rsCartao = Nothing
					
					Set rsCheque = objConexao.Execute("SELECT * FROM TB_PEDIDOS_CHEQUE WHERE COD_PEDIDO="&COD_PEDIDO&" ORDER BY BOM_PARA") 
					
					if Not rsCheque.Eof then
					%>
                    <table width="98%" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
					<tr bgcolor="#E4E4E4">
                        <td colspan="6"><div align="center" class="textoMedio"><strong>Cheque</strong></div></td>
                      </tr>
                      <tr>
                        <td width="25%"><div align="center"><strong>Nome Titular </strong></div></td>
                        <td width="15%"><div align="center"><strong>CPF</strong></div></td>
                        <td width="13%"><div align="center"><strong>C. Corrente </strong></div></td>
                        <td width="11%"><div align="center"><strong>N&ordm; Cheque </strong></div></td>
                        <td width="10%"><div align="center"><strong>Bom Para</strong></div></td>
                        <td><div align="center"><strong>Valor Cheque </strong></div></td>
                        </tr>
                      <%
					  TOTAL_CHEQUE = 0
					  Dim idParcial2, qtdChequesa
					  idParcial2 = 0
					  qtdChequesa = 0
					  While Not rsCheque.Eof
					  qtdChequesa = qtdChequesa + 1
					  idParcial2 = formatNumber(rsCheque("COD_PEDIDO_CHEQUE"))
					  %>
                      <tr>
                        <td><div align="left"><%=rsCheque("NOME")%></div></td>
                        <td><div align="center"><%=mascaraCPF(rsCheque("CPF"))%></div></td>
                        <td><div align="center"><%=rsCheque("CONTA_CORRENTE")%></div></td>
                        <td><div align="center"><%=rsCheque("NUMERO_CHEQUE")%></div></td>
                        <td><div align="center"><%=ArrumaData(rsCheque("BOM_PARA"))%></div></td>
                        <td><div align="right"><%=FormatNumber(rsCheque("VALOR"))%><%
						PARCIAL_CHEQUE = formatNumber(rsCheque("VALOR"))
						TOTAL_CHEQUE=TOTAL_CHEQUE+formatNumber(PARCIAL_CHEQUE)
						PARCIAL_CHEQUE = ""
						%></div></td>
                        </tr>
                      <%
					  rsCheque.MoveNext
					  Wend
					  
					  %>
                      <tr>
                        <td colspan="6">
                        <div align="left">
						<table width="400" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="33%"><div align="left" class="textoMedio">Total em Cheque:</div></td>
    <td width="20%"><div align="right" class="textoMedio"><strong><%=formatNumber(TOTAL_CHEQUE)%></strong></div></td>
    <td width="47%">�� |� �<span class="textoMedio">Qtd. cheque(s): <strong><%=qtdChequesa%></strong></span></td>
  </tr>
</table>
						</div>						</td>
                        </tr>
                    </table>
                    
                    <%
					End if
					
					rsCheque.Close
					Set rsCheque = Nothing


					'PAGAMENTO EM BOLETO
					sql = "SELECT sum(valor) as valor,"&_
					" count(TB_PEDIDOS_BOLETOS.hash) as qtd_parcelas,"&_
					" min(valor) as valor_parcela,"&_
					" cod_pedido, "&_
					" TB_PEDIDOS_BOLETOS.cod_sacado, "&_
					" nome, "&_
					" cpf,"&_
					" email,"&_
					" TB_SACADOS_ENDERECOS.endereco,"&_
					" TB_SACADOS_ENDERECOS.numero,"&_
					" TB_SACADOS_ENDERECOS.complemento,"&_
					" '00000-000' as cep,"&_
					" TB_SACADOS_ENDERECOS.bairro,"&_
					" TB_SACADOS_ENDERECOS.cidade,"&_
					" TB_SACADOS_ENDERECOS.estado"&_
					" FROM stillos3.TB_PEDIDOS_BOLETOS"&_
					" inner join TB_SACADOS on TB_PEDIDOS_BOLETOS.cod_sacado = TB_SACADOS.cod_sacado"&_
					" inner join TB_SACADOS_ENDERECOS on TB_SACADOS.cod_sacado = TB_SACADOS_ENDERECOS.cod_sacado"&_
					" where cod_pedido = "&sqlTexto(cod_pedido)&_
					" group by cod_pedido, cod_sacado, nome, cpf, email, TB_SACADOS_ENDERECOS.endereco, TB_SACADOS_ENDERECOS.numero, TB_SACADOS_ENDERECOS.complemento, TB_SACADOS_ENDERECOS.bairro, TB_SACADOS_ENDERECOS.cidade, TB_SACADOS_ENDERECOS.estado"
					set rs = objConexao.Execute(sql)
					if not rs.Eof then
					%> 
						<table width="98%" border="1" align="center" cellpadding="5" style="border-collapse:collapse; margin-top:10px;">
							<tr bgcolor="#E4E4E4"><td colspan="6"><div align="center" class="textoMedio"><strong>Boleto</strong></div></td></tr>
							<tr>
								<td width="10%"><div align="center"><strong>Sacado:</strong></div></td>
								<td>
									<font face='Courier'>
									<strong>Nome:</strong><%=rs("nome")%> <strong>CPF:</strong><%=mascaraCPF(rs("cpf"))%> <br />
									<strong>E-mail:</strong><%=rs("email")%> <br />
									<strong>Endere&ccedil;o:</strong><%=rs("endereco")%>, <%=rs("Numero")%> - <%=rs("complemento")%> <br>
									<%=rs("bairro")%>, <%=rs("cidade")%> / <%=rs("estado")%>
									</font>
								</td>
							</tr>
							<tr>
								<td width="25%"><div align="center"><strong>Boletos:</strong></div></td>
								<td>
									<font face='Courier' size='2'>
										Total:
										<br>
										<%=rs("qtd_parcelas")%>x R$<%=arrumanumero(rs("valor_parcela"),"",2 )%> = R$<%=arrumanumero(rs("valor"),"",2 )%>
										<br>
										<br>
										Parcelas:
										<br>
										<%
											dim quebra :  quebra = round(int(rs("qtd_parcelas")&"" )/2,0)
											dim ct_linha : ct_linha = 0	

											set rs2 = objConexao.execute("SELECT * FROM TB_PEDIDOS_BOLETOS WHERE cod_pedido = "&sqlNumero(cod_pedido))
											While not rs2.Eof

												if ct_linha = 0 then
													%> <table width="100%"> <tr'><td width='200px;'  valign="top"><%
												end if

												if ct_linha = quebra then
													%></td><td valign="top"><%
												end if
													Response.write(preencheCaracter(rs2("parcela"),"0",2)&" "&ArrumaData(rs2("data_vencimento"))&" R$"&rs2("valor")&"<br>")



												if ct_linha = int(0&rs("qtd_parcelas")) then
													%></td></tr></table><%
												end if 

												ct_linha = ct_linha + 1

											rs2.MoveNext
											Wend
											rs2.Close
											set rs2 = Nothing

										%>
									</font>
								</td>
							</tr>
							
						</table>
					<%
					end if
					rs.Close
					set rs = Nothing
				FecharConexao()
					%>

			
					
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