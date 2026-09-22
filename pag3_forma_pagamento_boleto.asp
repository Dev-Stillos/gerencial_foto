<!-- #include virtual="gerencial_foto/includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include virtual="gerencial_foto/includes/sha256.asp" -->
<!-- #include virtual="gerencial_foto/includes/classe_modulo_cob.asp" -->
<!-- #include file="includes/funcoes.asp" -->

<%
Dim COD_VENDEDOR, COD_PEDIDO

COD_VENDEDOR = Session("COD_VENDEDOR")
TOTAL_PAGO_EM_PRODUTOS = 0
COD_PEDIDO = criptografar(Request.QueryString("cod_pedido"))

If Session("COD_VENDEDOR") = Empty Then
Session.Abandon()
Response.Redirect("index.asp")
End if


'Verifica se o pedido não está finalizado
AbrirConexao()

Dim rsPedido

	Set rsPedido = objConexao.Execute("SELECT FINALIZA_VENDEDOR, PAGAMENTO_EM_DINHEIRO FROM TB_PEDIDOS WHERE COD_PEDIDO="&COD_PEDIDO)
		If not rsPedido.Eof then
			If rsPedido("FINALIZA_VENDEDOR") = 1 Then
				Session.Abandon()
				Response.Redirect("index.asp")
			End If
			EM_DINHEIRO = formatNumber(0&rsPedido("PAGAMENTO_EM_DINHEIRO"))
		Else
			Session.Abandon()
			Response.Redirect("index.asp")
		End If
	rsPedido.Close()
	Set rsPedido = Nothing	

  

  'verificamos tb se já não existe boletos lançados
  set rs = objConexao.Execute("select * from TB_PEDIDOS_BOLETOS where cod_pedido = "&cod_pedido)
  if not rs.Eof then
    Response.Redirect("pag3_forma_pagamento.asp?cod_pedido="&criptografar(cod_pedido))
  end if
  rs.Close
  set rs = Nothing


FecharConexao()
'--------------------------------------
Dim chave

chave = Request.QueryString("z")

'TB_FOTO_VENDA_CHEQUES
Dim BANCO_1, AGENCIA_1, CONTA_CORRENTE_1, VALOR_1, BOM_PARA_1, NOME_1, CPF_1, SE_TERCEIRO_1, FONE_1



'///////////////////////////////////////////////
'Gravar compra através do cartão de crédito
If Request.Form("campo1") <> "" then
  AbrirConexao()
  dim dia_vencimento, mes_vencimento, ano_vencimento, qtd_parcelas
  dia_vencimento = request.form("dia_vencimento")
  mes_vencimento = request.form("mes_vencimento")
  ano_vencimento = request.form("ano_vencimento")
  qtd_parcelas = request.form("qtd_parcelas")

  dim nome_sacado, cpf, grau_parentesco, email
  'VERIFICAMOS QUAL CAMPO NOME DEVE SER USADO'
  if Request.form("NOME_TITULAR") = "#outro" then
    nome_sacado = request.form("NOME_OUTRO")
  else
    nome_sacado = Request.form("NOME_TITULAR")
  end if
  cpf =  mascaraCPF(request.form("cpf")&"-")
  grau_parentesco = Request.form("grau_parentesco")
  email = Request.form("email")

  dim cep, endereco, numero, complemento, bairro, cidade, estado
  cep = Request.form("cep")
  endereco = Request.form("endereco")
  numero = Request.form("numero")
  complemento = Request.form("complemento")
  bairro = Request.form("bairro")
  cidade = Request.form("cidade")
  estado = Request.form("estado")

  'SALVAMOS O SACADO 
  objConexao.Execute("INSERT INTO TB_SACADOS (nome, cpf, grau_parentesco, email) values ("&_
                      sqlTexto(nome_sacado)&","&_
                      sqlTexto(cpf)&","&_
                      sqlTexto(grau_parentesco)&","&_
                      sqlTexto(email)&""&_
                      ") ")

  dim cod_sacado

  set rs = objConexao.Execute("SELECT @@IDENTITY AS 'cod_sacado'")
  if not rs.Eof then
    cod_sacado = rs(0)
  end if
  rs.Close
  set rs = Nothing


  'SALVAMOS O ENDERECO DO SACADO
  objConexao.Execute("INSERT INTO TB_SACADOS_ENDERECOS (cod_sacado, endereco, numero, complemento, bairro, cidade, estado) values ("&_
                    sqlNumero(cod_sacado)&","&_
                    sqlTexto(endereco)&","&_
                    sqlNumero(numero)&","&_
                    sqlTexto(complemento)&","&_
                    sqlTexto(bairro)&","&_
                    sqlTexto(cidade)&","&_
                    sqlTexto(estado)&""&_
                    ") ")

  'GERAMOS OS BOLETOS SEM A COMISSAO
  valor_apagar = parcelasBoleto(cod_pedido, "valor_apagar", true)
  valor_parcela_ok = cdbl(FormatNumber(valor_apagar/qtd_parcelas,2))
  dim diferenca : diferenca = valor_apagar - (valor_parcela_ok*qtd_parcelas)



  for i = 1 to qtd_parcelas

    if (diferenca <> 0 and i = 1) then
      valor_parcela = valor_parcela_ok + diferenca
    else
      valor_parcela = valor_parcela_ok
    end If

    nosso_numero = COD_PEDIDO & preencheCaracter(i,"0",2)

    '//data_vencimento(dia_vencimento, mes_vencimento, ano_vencimento, parcela)
    vencimento = data_vencimento(dia_vencimento, mes_vencimento, ano_vencimento, i)
    hash = sha256(nosso_numero & hashkey)

    objConexao.Execute("INSERT INTO TB_PEDIDOS_BOLETOS (cod_pedido, cod_sacado, parcela, nosso_numero, hash, valor, data_vencimento) VALUES ("&_
        sqlNumero(cod_pedido)&", "&_
        sqlNumero(cod_sacado)&", "&_
        sqlNumero(i)&", "&_
        sqlTexto(nosso_numero)&", "&_
        sqlTexto(hash)&", "&_
        sqlNumero(valor_parcela)&", "&_
        sqlData(vencimento)&""&_
      ") ")

  next

  FecharConexao()
  Response.Redirect("pag3_forma_pagamento.asp?cod_pedido="&criptografar(cod_pedido))

End if

'----------------------------------------------------
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link type="text/css" rel="stylesheet" href="css/principal.css" media="screen"></link>
<script type="text/javascript" src='js/jquery-1.11.2.min.js'></script>
<script type="text/javascript" src='js/bootstrap.min.js'></script>
<script type="text/javascript" src='js/jquery-mask.js'></script>

<link rel="stylesheet" type="text/css" href="css/bootstrap.min.css" />
<link type="text/css" rel="stylesheet" href="css/principal.css"></link>

<link type="text/css" rel="stylesheet" href="css/font-awesome.min.css"></link>
<link type="text/css" rel="stylesheet" href="js/font-awesome-animation.min.css"></link>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Loja Virtual - Stillo's Produções e Eventos</title>

<body>
<table width="690" align="center" cellpadding="0" cellspacing="0" id="tb_p">
<tr>
<td width="688"><img src="imagens/cabecalho.jpg" width="688" height="109" /></td>
</tr>
<tr>
<td><img style="padding-top:0px" src="imagens/3_passo.jpg" width="688" /></td>
</tr>

<tr>
<td align="center">
<div id="tb_principal">
		<span class="style2">Adicionar Pagamento: Boleto 
</span>
		<div id="linha_preta"></div><br />
		<form method="post" name="addCartao">	
      <table width="95%" border="0" align="center" cellpadding="3" cellspacing="3">
  <tr>
    <td width="189"><div align="left">Primeiro Vencimento:</div></td>
    <td width="360">
      <div align="left">
        <select name='dia_vencimento' required="required" id='dia_vencimento_id' class='valida_vencimento validacao_etapa_1'>
          <option value=""></option>
          <option value="07">07</option>
          <option value="10">10</option>
          <option value="22">22</option>
          <option value="25">25</option>
        </select>
        /
        <select name='mes_vencimento' required="required" id='mes_vencimento_id' class='valida_vencimento validacao_etapa_1'>
          <option value=""></option>
          <option value="1">Janeiro</option>
          <option value="2">Fevereiro</option>
          <option value="3">Março</option>
          <option value="4">Abril</option>
          <option value="5">Maio</option>
          <option value="6">Junho</option>
          <option value="7">Julho</option>
          <option value="8">Agosto</option>
          <option value="9">Setembro</option>
          <option value="10">Outubro</option>
          <option value="11">Novembro</option>
          <option value="12">Dezembro</option>
        </select>
        /
        <select name='ano_vencimento' required="required" id='ano_vencimento_id' class='valida_vencimento validacao_etapa_1'>
          <option value=""></option>
          <option value="2016">2016</option>
          <option value="2017">2017</option>
        </select>
      </div>
    </td>
  </tr>
   <tr>
    <td width="189"><div align="left">Quant. Parcelas:</div></td>
    <td width="360"><div align="left"><%=parcelasBoleto(cod_pedido, "html", false)%></div></td>
  </tr>
  </table>

  <div id='Etapa2' class="">
    <hr />
  <table width="95%" border="0" align="center" cellpadding="3" cellspacing="3">
  <tr>
    <td><div align="left">Sacado:</div></td>
    <td>
      <div align="left">
        <select name="NOME_TITULAR" style="margin-bottom:5px" class="validacao_etapa_2" required="required">
          <option>Selecione o nome do Sacado</option>
          <%
          AbrirConexao()
          sql = "SELECT TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO, TB_FICHAS_IDENTIFICACAO.ALBUM, TB_FICHAS_IDENTIFICACAO.INICIAL, TB_FICHAS_IDENTIFICACAO.BRINDE, TB_FICHAS_IDENTIFICACAO.PRODUTO_AGREGADO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, TB_FICHAS_IDENTIFICACAO.COMISSAO, TB_PEDIDOS_FORMANDOS.COD_FORMANDO FROM TB_FICHAS_IDENTIFICACAO INNER JOIN TB_PEDIDOS_FORMANDOS ON TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = TB_PEDIDOS_FORMANDOS.COD_FORMANDO "&_
            "WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&COD_PEDIDO
          Set rsFormsCor = objConexao.Execute(sql)
          While not rsFormsCor.Eof
          %>
            <option value="<%=rsFormsCor("NOME_FORMANDO")%>"><%=rsFormsCor("NOME_FORMANDO")%></option>
          <%
          rsFormsCor.MoveNext()
          Wend
          rsFormsCor.Close
          Set rsFormsCor = Nothing
          FecharConexao()
          %>
          <option value="#outro">Outro...</option>
        </select>
        <input name="NOME_OUTRO" type="text" size="40" maxlength="80" class="hidden" />
      </div>
    </td>
  </tr>
  <tr>
    <td width="189"><div align="left">CPF*:</div></td>
    <td width="360"><div align="left"><input name="CPF" type="text" class="padrao_cpf" size="40" maxlength="80" class="validacao_etapa_2" required="required"/></div></td>
  </tr>
  
  <tr>
    <td><div align="left">Se for de terceiro informe o grau de parentesco:</div></td>
    <td><div align="left">
      <input name="GRAU_PARENTESCO" type="text" size="40" maxlength="50"  />
    </div></td>
  </tr>

  <tr>
    <td height="30"><div align="left">E-mail*:</div></td>
    <td><div align="left">
      
      <input name="email" type="email" required="required" size="40" maxlength="80"  class="validacao_etapa_2" />
    </div></td>
  </tr>

  </table>

  

  

<div id='Etapa3' class="">
<hr />
  <table width="95%" border="0" align="center" cellpadding="3" cellspacing="3">

  <tr>
    <td height="30"><div align="left">CEP:</div></td>
    <td>
      <input name="CEP" type="text" required="required" size="28" class="padrao_cep" id='cep_id' maxlength="80" class="validacao_etapa_3" />
<input type="button" id='buscar_cep_id' value="Busar CEP" />

      <input name="CEP" type="hidden" id='cep_ultima_pesquisa' maxlength="80"  />
      &nbsp;&nbsp;<i class="fa fa-refresh faa-spin animated hidden " id='carregandoa' style="font-size:16px;"></i>
    </td>
  </tr>

  <tr>
    <td width="189"><div align="left">Endereço*:</div></td>
    <td width="360"><input name="ENDERECO" type="text" required="required" size="40" maxlength="255" class="validacao_etapa_3" /></td>
  </tr>

  <tr>
    <td width="189"><div align="left">Número*:</div></td>
    <td width="360"><input name="NUMERO" type="number" required="required"  maxlength="80" min='1' class="validacao_etapa_3" /></td>
  </tr>


  <tr>
    <td width="189"><div align="left">Complemento:</div></td>
    <td width="360"><input name="COMPLEMENTO" type="text" size="40" maxlength="80" /></td>
  </tr>

  <tr>
    <td width="189"><div align="left">Bairro*:</div></td>
    <td width="360"><input name="BAIRRO" type="text" required="required" size="40" maxlength="80" class="validacao_etapa_3" /></td>
  </tr>

  <tr>
    <td width="189"><div align="left">Cidade*:</div></td>
    <td width="360"><input name="CIDADE" type="text" required="required" size="40" maxlength="80" class="validacao_etapa_3" /></td>
  </tr>

  <tr>
    <td width="189"><div align="left">Estado*:</div></td>
    <td width="360"><input name="ESTADO" type="text" required="required" size="40" maxlength="80" class="validacao_etapa_3" required="required"  /></td>
  </tr>
</table>

<div id='Etapa4' class="">

<table width="95%" border="0" align="center" cellpadding="3" cellspacing="3">
  <tr>
    <td colspan="2"><div align="center">
        <input type="hidden" name="campo1" value="1" />
      <input type="submit" onclick="return validaFormulario();" name="CARTAO_GRAVAR" value="Gravar" />
   &nbsp;&nbsp; <span style="padding-top:10px">
   <input name="button" type="button" onclick="history.go(-1)" value="Voltar" />
   </span></div></td>
    </tr>
</table>
</div><!-- etapa2 -->
</div><!-- etapa3 -->
</div><!-- etapa4 -->
		</form>
</p>
</div>



</td>
</tr>
</table>
<!-- #include file="rodape.asp" -->

<script type="text/javascript">

     function validaFormulario()
    {
      var validado = true;

      $("[required=required]").each(function(){
        if($(this).val() == "")
        {
          validado = false;
        }
      });

      return validado;
    }



  $(function(){

    function disableInputs(){
      $("input").prop('disabled', true);
      $("select").prop('disabled', true);
    }

    function enableInputs(){
      $("input").prop('disabled', false);
      $("select").prop('disabled', false);
    }
    
    //--------------------------------------------------------------
    // Validações de Etapas
    //--------------------------------------------------------------
      /*
    function validaEtapa(etapa)
    {
      var ok = true;
      $(".validacao_etapa_"+etapa+"[required=required]").each(function(){
        if($(this).val() == '')
        {
          ok = false;
        }
      });

      if (etapa == 1 )
      {
        if (ok)
          $("#Etapa2").removeClass("hidden");
        else
          $("#Etapa2").addClass("hidden");
      }

      if (etapa == 2 )
      {
        if (ok)
          $("#Etapa3").removeClass("hidden");
        else
          $("#Etapa3").addClass("hidden");
      }      

      if (etapa == 3 )
      {
        if (ok)
          $("#Etapa4").removeClass("hidden");
        else
          $("#Etapa4").addClass("hidden");
      }
    }
  
    $(".validacao_etapa_1").blur(function(){validaEtapa(1)});
    $(".validacao_etapa_1").change(function(){validaEtapa(1)});

    $(".validacao_etapa_2").blur(function(){validaEtapa(2)});
    $(".validacao_etapa_2").change(function(){validaEtapa(2)});

    $(".validacao_etapa_3").blur(function(){validaEtapa(3)});
    $(".validacao_etapa_3").change(function(){validaEtapa(3)});
   */

    //--------------------------------------------------------------
    // Data de Vencimento
    //--------------------------------------------------------------
    function validaVencimento(){
      var dia = $("#dia_vencimento_id").val();
      var mes = $("#mes_vencimento_id").val();
      var ano = $("#ano_vencimento_id").val();

      console.log('dia : '+ dia);
      console.log('mes : '+ mes);
      console.log('ano : '+ ano);

      if (dia != '' && mes != '' && ano != '')
      {
        //validamos via AJAX
        var request = $.ajax({
          url: "ajax.asp",
          method: "POST",
          data: { tp : 'valida_vencimento', p1 : dia, p2 : mes, p3 : ano },
          dataType: "text",
          beforeSend : function(){
            $("#carregando_vencimento").removeClass("hidden");
            disableInputs();
          }
        });

        request.done(function( msg ) {
          if(msg.toLowerCase() == 'false')
          { 
            alert('Data de vencimento inválida');
            $(".valida_vencimento").val('');
          }
          $("#carregando_vencimento").addClass("hidden");
          enableInputs();
          $("select[name=qtd_parcelas]").focus();
          $("#vencimento_validado").val(msg.toLowerCase());
        });

        request.fail(function( jqXHR, textStatus ) {
          $("#carregando_vencimento").addClass("hidden");
          enableInputs();
        });
      }
    }

    $(".valida_vencimento").blur(function(){validaVencimento()});
    

    //--------------------------------------------------------------
    // CEP
    //--------------------------------------------------------------
    $('#buscar_cep_id').click(function(){
      var cep_ultima_pesquisa = $("#cep_ultima_pesquisa").val();

      if (cep_ultima_pesquisa != $("#cep_id").val())
      {
        $("#cep_ultima_pesquisa").val($("#cep_id").val());
        var request = $.ajax({
          url: "busca_cep.php",
          method: "POST",
          data: { cep : $('#cep_id').val() },
          dataType: "text",
          beforeSend : function(){
            $("#carregandoa").removeClass("hidden");
            $("#buscar_cep_id").addClass("hidden");
            disableInputs();
          }
        });

        request.done(function( msg ) {
          if(msg == '')
          {
            alert('CEP Não localizado');
            $("#carregandoa").addClass("hidden");
            $("#buscar_cep_id").removeClass("hidden");
            enableInputs();
            $('input[name=ENDERECO]').focus();
          }
          else
          {
            $("#carregandoa").addClass("hidden");
            $("#buscar_cep_id").removeClass("hidden");
            enableInputs();
            cep = JSON.parse(msg);
            $('input[name=ENDERECO]').val(cep['logradouro']);
            $('input[name=BAIRRO]').val(cep['bairro']);
            $('input[name=CIDADE]').val(cep['cidade']);
            $('input[name=ESTADO]').val(cep['estado']);
            $('input[name=NUMERO]').focus();
          }
        });

        request.fail(function( jqXHR, textStatus ) {
          $("#carregandoa").addClass("hidden");
          enableInputs();
        });
      }
    });

    $("select[name=NOME_TITULAR]").change(function(){
      if ($(this).val() == "#outro" ) {
        $("input[name=NOME_OUTRO]").removeClass('hidden');
      }
      else
      {
        $("input[name=NOME_OUTRO]").addClass('hidden');
      }
    });
  


    //--------------------------------------------------------------
    //PADRAO
    //--------------------------------------------------------------
    $(".padrao_cep").mask("99999-999"); 
    $(".padrao_data").mask("99/99/9999" );  
    $(".padrao_cpf").mask("999.999.999-99" );

    console.log('mascaras carregadas');



    $('.padrao_cpf').blur(function(){
        var cpfv;
        var erro = new Array();
        var d = new Array();
        var s = new Number;
        var c = 10;
        var s = 0;
        
        cpfv = $(this).val();
        //cpfv = cpfv.replace(/\D/,'');
        cpfv = cpfv.replace('-','');
        cpfv = cpfv.replace('_','');
        cpfv = cpfv.replace('.','');
        cpfv = cpfv.replace('.','');
        
        
      if (cpfv !== '')
      {
        
        if (cpfv.length === 11)
        {
          if (cpfv == "00000000000" || cpfv == "11111111111" || cpfv == "22222222222" || cpfv == "33333333333" || cpfv == "44444444444" || cpfv == "55555555555" || cpfv == "66666666666" || cpfv == "77777777777" || cpfv == "88888888888" || cpfv == "99999999999") {
              erro[0] = true;
              erro[1] = "Número de CPF inválido!";
          }
          
          for(var i=0; i<9; i++){
            d[i] = cpfv.charAt(i);
            s += (d[i] * c);
            c--;
          }   
          
          x = s % 11
          if(x < 2){
            d[9] = 0
          }else{
            d[9] = 11-x
          }
          
          s = 0;
          c = 11;
          for(y=0; y<10; y++) 
          {
            s += (d[y] * c--);
          }
          
          x = s % 11      
          if(x < 2){
            d[10] = 0;
          }else{
            d[10] = 11-x;
          }
    
          if((cpfv.charAt(9) != d[9]) || (cpfv.charAt(10) != d[10])){
            erro[0] = true;
            erro[1] = "Número de CPF inválido!";
          }
          
          if (erro[0] === true)
          {
            alert(erro[1]);
            $(this).val('');    
          }
        }
        else
        {
          alert("Número de CPF inválido!");
        }
        
      }
      });

  });


  


</script>


</body>
</html>
