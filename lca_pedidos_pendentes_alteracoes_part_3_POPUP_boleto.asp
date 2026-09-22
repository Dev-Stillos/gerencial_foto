<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/classe_modulo_cob.asp" -->
<!-- #include file="includes/sha256.asp" -->
<%

Dim COD_PEDIDO, comando, sql, cod_formando
COD_PEDIDO      = Request.QueryString("cod_pedido")
comando       = Request.QueryString("comando")

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
    grau_parentesco = Request.form("grau_parentesco")
    cep = Request.form("cep")
    endereco = Request.form("endereco")
    numero = Request.form("numero")
    complemento = Request.form("complemento")
    bairro = Request.form("bairro")
    cidade = Request.form("cidade")
    estado = Request.form("estado")
  else
    cod_formando = Request.form("NOME_TITULAR") 'recebe o codigo do fromando

    set rs = objConexao.Execute("SELECT COD_FORMANDO, NOME_FORMANDO FROM TB_FICHAS_IDENTIFICACAO WHERE COD_FORMANDO = "&sqlNumero(cod_formando))
    if not rs.Eof then
      nome_sacado = rs("NOME_FORMANDO")
    end if
    rs.Close
    set rs = Nothing

    set rs = objConexao.execute("SELECT * from TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO  WHERE  COD_FORMANDO = "&sqlTexto(cod_formando)&" AND DATA_CADASTRO = (SELECT MAX(DATA_CADASTRO) from TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO WHERE COD_FORMANDO = "&sqlTexto(cod_formando)&" ) ")
    if not rs.Eof Then
      grau_parentesco = ""
      cep = rs("cep")
      endereco = rs("endereco")
      numero = ""
      complemento = ""
      bairro = rs("bairro")
      cidade = rs("cidade")
      estado = rs("estado")
    end if 
    rs.Close
    Set rs = Nothing

  end if

  cpf =  mascaraCPF(request.form("cpf")&"-")
  email = Request.form("email")

  dim cep, endereco, numero, complemento, bairro, cidade, estado

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

  objConexao.Execute("UPDATE TB_PEDIDOS SET cod_sacado = "&sqlNumero(cod_sacado)&" WHERE COD_PEDIDO = "&sqlNumero(COD_PEDIDO))


  'SALVAMOS O ENDERECO DO SACADO
  objConexao.Execute("INSERT INTO TB_SACADOS_ENDERECOS (cod_sacado, endereco, numero, complemento, bairro, cidade, estado) values ("&_
                    sqlNumero(cod_sacado)&","&_
                    sqlTexto(endereco)&","&_
                    sqlTexto(sqlNumero(numero))&","&_
                    sqlTexto(complemento)&","&_
                    sqlTexto(bairro)&","&_
                    sqlTexto(cidade)&","&_
                    sqlTexto(estado)&""&_
                    ") ")

  'GERAMOS OS BOLETOS SEM A COMISSAO
  valor_apagar = parcelasBoleto(cod_pedido, "valor_apagar", true)
  valor_parcela_ok = cdbl(FormatNumber(valor_apagar/qtd_parcelas,2))
  dim diferenca : diferenca = valor_apagar - (valor_parcela_ok*qtd_parcelas)

  dim ii
  for ii = 1 to int(qtd_parcelas)

    if (diferenca <> 0 and ii = 1) then
      valor_parcela = valor_parcela_ok + diferenca
    else
      valor_parcela = valor_parcela_ok
    end If

    nosso_numero = COD_PEDIDO & preencheCaracter(ii,"0",2)

    '//data_vencimento(dia_vencimento, mes_vencimento, ano_vencimento, parcela)
    vencimento = data_vencimento(dia_vencimento, mes_vencimento, ano_vencimento, ii)
    hash = sha256(nosso_numero & hashkey)

    objConexao.Execute("INSERT INTO TB_PEDIDOS_BOLETOS (cod_pedido, cod_sacado, parcela, nosso_numero, hash, valor, data_vencimento) VALUES ("&_
        sqlNumero(cod_pedido)&", "&_
        sqlNumero(cod_sacado)&", "&_
        sqlNumero(ii)&", "&_
        sqlTexto(nosso_numero)&", "&_
        sqlTexto(hash)&", "&_
        sqlNumero(valor_parcela)&", "&_
        sqlData(vencimento)&""&_
      ") ")

  next
  

  FecharConexao()
  'Response.Redirect("lca_pedidos_pendentes_alteracoes_part_3.asp?cod_pedido="&COD_PEDIDO&"&msg=Boleto alterado com sucesso!") 
  %>
    <script>
        window.parent.location.reload();
    </script>
  <%
  Response.end
End if

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Cadastro de Produto em pedido</title>
<link type="text/css" rel="stylesheet" href="css/base.css" media="screen"></link>
<script type="text/javascript" src="js/jquery-1.11.2.min.js"></script>
<script type="text/javascript" src='js/jquery-mask.js'></script>

<style type="text/css">
<!--
body {
	background-color: #FFFFFF;
}
-->
</style></head>
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





<script>
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

//Formata número tipo moeda usando o evento onKeyDown

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
<body>

<form method="post" name=""> 
      <table width="95%" border="0" align="center" cellpadding="3" cellspacing="3">
  <tr>
    <td width="189"><div align="left">Primeiro Vencimento:</div></td>
    <td width="360">
      <div align="left">
        <select name='dia_vencimento' required="required" id='dia_vencimento_id' class='valida_vencimento validacao_etapa_1' nome-campo='vencimento'>
          <option value=""></option>
          <option value="01">01</option>
          <option value="02">02</option>
          <option value="03">03</option>
          <option value="04">04</option>
          <option value="05">05</option>
          <option value="06">06</option>
          <option value="07">07</option>
          <option value="08">08</option>
          <option value="09">09</option>
          <option value="10">10</option>
          <option value="11">11</option>
          <option value="12">12</option>
          <option value="13">13</option>
          <option value="14">14</option>
          <option value="15">15</option>
          <option value="16">16</option>
          <option value="17">17</option>
          <option value="18">18</option>
          <option value="19">19</option>
          <option value="20">20</option>
          <option value="21">21</option>
          <option value="22">22</option>
          <option value="23">23</option>
          <option value="24">24</option>
          <option value="25">25</option>
          <option value="26">26</option>
          <option value="27">27</option>
          <option value="28">28</option>
          <option value="29">29</option>
          <option value="30">30</option>
        </select>
        /
        <select name='mes_vencimento' required="required" id='mes_vencimento_id' class='valida_vencimento validacao_etapa_1' nome-campo='vencimento' >
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
        <select name='ano_vencimento' required="required" id='ano_vencimento_id' class='valida_vencimento validacao_etapa_1' nome-campo='vencimento'>
          <option value=""></option>
          <option value="2020">2020</option>
          <option value="2021">2021</option>
          <option value="2022">2022</option>
          <option value="2023">2023</option>
          <option value="2024">2024</option>
          <option value="2025">2025</option>
          <option value="2025">2025</option>
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
        <select name="NOME_TITULAR" style="margin-bottom:5px; width:293px;" class="validacao_etapa_2" required="required"  nome-campo='Parcelas' >
          <option value="">Selecione o nome do Sacado</option>
          <%
          AbrirConexao()
          sql = "SELECT TB_FICHAS_IDENTIFICACAO.COD_PRODUCAO, TB_FICHAS_IDENTIFICACAO.ALBUM, TB_FICHAS_IDENTIFICACAO.INICIAL, TB_FICHAS_IDENTIFICACAO.BRINDE, TB_FICHAS_IDENTIFICACAO.PRODUTO_AGREGADO, TB_FICHAS_IDENTIFICACAO.NOME_FORMANDO, TB_FICHAS_IDENTIFICACAO.COMISSAO, TB_PEDIDOS_FORMANDOS.COD_FORMANDO FROM TB_FICHAS_IDENTIFICACAO INNER JOIN TB_PEDIDOS_FORMANDOS ON TB_FICHAS_IDENTIFICACAO.COD_FORMANDO = TB_PEDIDOS_FORMANDOS.COD_FORMANDO "&_
            "WHERE TB_PEDIDOS_FORMANDOS.COD_PEDIDO = "&COD_PEDIDO
          Set rsFormsCor = objConexao.Execute(sql)
          While not rsFormsCor.Eof
          %>
            <option value='<%=rsFormsCor("COD_FORMANDO")%>'><%=rsFormsCor("NOME_FORMANDO")%></option>
          <%
          rsFormsCor.MoveNext()
          Wend
          rsFormsCor.Close
          Set rsFormsCor = Nothing 
          FecharConexao()
          %>
          <option value="#outro">Outro...</option>
        </select>
        <input name="NOME_OUTRO" type="text" size="40" maxlength="80" class="hidden" class="required_form" nome-campo='Nome Formando' />
      </div>
    </td>
  </tr>
  <tr>
    <td width="189"><div align="left">CPF*:</div></td>
    <td width="360"><div align="left"><input name="CPF" type="text" class="padrao_cpf" size="40" maxlength="80" class="validacao_etapa_2" required="required"  nome-campo='CPF'/></div></td>
  </tr>
  
  <tr class='oculta_formando hidden'>
    <td><div align="left">Se for de terceiro informe o grau de parentesco:</div></td>
    <td><div align="left">
      <input name="GRAU_PARENTESCO" type="text" size="40" maxlength="50"  />
    </div></td>
  </tr>

  <tr></tr>
    <td height="30"><div align="left">E-mail*:</div></td>
    <td><div align="left">
      
      <input name="email" type="email" required="required" size="40" maxlength="80"  class="validacao_etapa_2" nome-campo='E-mail' />
    </div></td>
  </tr>
  </table>
<div id='Etapa3' class="">
<hr />
  <table width="95%" border="0" align="center" cellpadding="3" cellspacing="3"  class='oculta_formando hidden'>

  <tr>
    <td height="30"><div align="left">CEP:</div></td>
    <td>
      <input name="CEP" type="text" size="28" class="padrao_cep" id='cep_id' maxlength="80" class="validacao_etapa_3 required_form"  nome-campo='CPF' />
<input type="button" id='buscar_cep_id' value="Busar CEP" />

      <input name="CEP" type="hidden" id='cep_ultima_pesquisa' maxlength="80"  />
      &nbsp;&nbsp;<i class="fa fa-refresh faa-spin animated hidden " id='carregandoa' style="font-size:16px;"></i>
    </td>
  </tr>

  <tr>
    <td width="189"><div align="left">Endereço*:</div></td>
    <td width="360"><input name="ENDERECO" type="text" size="40" maxlength="255" class="validacao_etapa_3 required_form" nome-campo='Endereço'/></td>
  </tr>

  <tr>
    <td width="189"><div align="left">Número*:</div></td>
    <td width="360"><input name="NUMERO" type="number" maxlength="80" min='1' class="validacao_etapa_3 required_form"  nome-campo='Número' /></td>
  </tr>


  <tr>
    <td width="189"><div align="left">Complemento:</div></td>
    <td width="360"><input name="COMPLEMENTO" type="text" size="40" maxlength="80" nome-campo='Complemento' /></td>
  </tr>

  <tr>
    <td width="189"><div align="left">Bairro*:</div></td>
    <td width="360"><input name="BAIRRO" type="text" size="40" maxlength="80" class="validacao_etapa_3 required_form" nome-campo='Bairro' /></td>
  </tr>

  <tr>
    <td width="189"><div align="left">Cidade*:</div></td>
    <td width="360"><input name="CIDADE" type="text" size="40" maxlength="80" class="validacao_etapa_3 required_form" nome-campo='Cidade' /></td>
  </tr>

  <tr>
    <td width="189"><div align="left">Estado*:</div></td>
    <td width="360"><input name="ESTADO" type="text" size="40" maxlength="80" class="validacao_etapa_3 required_form" nome-campo='Estado' /></td>
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


<script>
  
  function validaFormulario()
    {
      var validado = true;
      var campos = '';
      $("[required=required]").each(function(){
        if($(this).val() == "")
        {
          validado = false;
          campos += $(this).attr("nome-campo") +   "\n";
        }
      }); 
      if (campos.length > 0) {
        alert(campos);
      }
      
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
          url: "ajax_boletos.asp",
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
          $("#buscar_cep_id").removeClass("hidden");
          $("#carregandoa").addClass("hidden");
          enableInputs();
        });
      }
    });

    $("select[name=NOME_TITULAR]").change(function(){
      if ($(this).val() == "#outro" ) {
        $("input[name=NOME_OUTRO]").removeClass('hidden');
        $(".oculta_formando").removeClass('hidden');

       $(".required_form").attr("required","");
      }
      else
      {
        $("input[name=NOME_OUTRO]").addClass('hidden');
        $(".oculta_formando").addClass('hidden');

        
         $(".required_form").removeAttr("required","");
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
