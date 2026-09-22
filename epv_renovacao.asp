
<%
Dim calendario_necessario, BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
calendario_necessario = "NÃO"
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 1
BTN_IMPRIMIR_LINK = "#"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%
Dim comando, TXT_DEVOLUCAO, ERROR_LINHA, pronto, CONT_ITENS
pronto = Request.QueryString("pronto")
ERROR_LINHA = 1
comando = Request.QueryString("comando")
TXT_DEVOLUCAO = Request.Form("TXT_DEVOLUCAO")
TXT_DEVOLUCAO = Replace(TXT_DEVOLUCAO," ","")
CONT_ITENS = 0

COD_USUARIO = sqlNumero(Session("COD_USUARIO"))

If comando = "nova" Then
'Apaga tudo da base do determinado usu�rio
AbrirConexao()
	objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_RENOVACAO_TEMPORARIA WHERE COD_USUARIO = "&COD_USUARIO)
FecharConexao()
End If

If comando = "efetuar_devolucao" Then
  AbrirConexao()
  dim devolucoes : devolucoes = ""
  Set rs = objConexao.Execute("SELECT * FROM TB_FICHAS_IDENTIFICACAO_RENOVACAO_TEMPORARIA WHERE COD_USUARIO="&COD_USUARIO)
  While not rs.Eof
    cod_renovacao = sqlNumero(rs("cod_renovacao_temp").value)
    cod_formando = sqlNumero(rs("cod_formando").value)
    cod_vendor = sqlNumero(rs("cod_vendedor").value)
    periodo_renovar = sqlNumero(rs("periodo_renovacao").value)
    novo_prazo = sqlData(dateAdd("d",periodo_renovar,Date()))
    motivo = sqlTexto(Request.form("motivo_"&rs("cod_formando").value&rs("tipo").value))
    tipo = sqlTexto(rs("tipo").value)
    renovacaoes = renovacaoes & cod_renovacao & "," 

    '1) Pega a data do prazo e aumenta para uma nova data com base no parametro de dias de renova��o e na data adual
    objConexao.Execute("UPDATE TB_PRODUCAO_LOTES_FORMANDOS SET DATA_PRAZO = "&novo_prazo&" WHERE COD_FORMANDO ="&sqlNumero(cod_formando)&" LIMIT 1")

    '2) Altera nas movimenta��es o novo prazo somente das movimenta��es que est�o em aberto
    objConexao.Execute("UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET DATA_PREVISTA_RETORNO = "&novo_prazo&" WHERE COD_FORMANDO ="&sqlNumero(cod_formando)&" AND DATA_RETORNO IS NULL AND TIPO = "&tipo)
            
    '3) Insere o motivo da renova��o na ficha do formando, sem necessidade de contato
    objConexao.Execute("INSERT INTO TB_FICHAS_IDENTIFICACAO_HISTORICO (COD_FORMANDO, GERADO_POR, MENSAGEM, COD_VENDEDOR) VALUES (" &cod_formando& ", "&cod_usuario&", "&motivo&", "&cod_vendor& ")")

  rs.MoveNext()
  Wend
  rs.Close()
  Set rs = Nothing

  if len(renovacaoes) > 0 Then
    renovacaoes = left(renovacaoes, len(renovacaoes) - 1)
    sql = "UPDATE TB_FICHAS_IDENTIFICACAO_RENOVACAO_TEMPORARIA SET PODE_IMPRIMIR = 1 WHERE COD_RENOVACAO_TEMP in ("&renovacaoes&")"
    objConexao.Execute(sql)
  end if
  FecharConexao()

	Response.Redirect("epv_renovacao_imprimir.asp")
End If
%>

<h2>EPV - Renova&ccedil;&atilde;o de Produtos </h2>
<div id="retorna_div"></div>
<div class="subMenu">
  Sub-Menu: 
  <a href="epv_devolucao.asp?comando=nova">Iniciar devolu&ccedil;&atilde;o de Vendedor</a> | 
  <a href="epv_saida_inicio.asp?comando=nova">Sa&iacute;da de &Aacute;lbuns do Estoque</a>
</div>
<div class="msg" align="center">
  <p><%=Request.QueryString("msg")%></p>
</div>

<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td style="padding:5px;">
      <div class="msgGrandeSucesso" style="display:none" id="sucesso_msg"></div>
      <div class="msgGrandeErro" style="display:none"  id="erro_msg"></div>
      <div align="center">
        <input onkeypress="return runScript(event,this);" maxlength="8" style="height:50px; width:200px; font-size:34px; border:" type="text" name="FORM_TXT_DEVOLUCAO" id="TXT_DEVOLUCAO" />
      </div>
    </td>
  </tr>
  <tr>
    <td>
      <div align="center">
        <input class="btnGravar " type="button" onclick="gravarItem()" id="botaoGravar" value=" " />
        <br /><br />
      </div>
    </td>
  </tr>
</table>

<form 
  action="?comando=efetuar_devolucao" 
  method="post" 
  name="DEVOLUCAO_ITENS" 
  style="margin:0px" 
  onsubmit="return confirmaDevolucao()">

  <div id="containerListagem"></div>
</form>

<div id="ajaxCarregandoContainer" class="ajaxCarregando" style="visibility:hidden">
	<table height="100%" width="100%" id="ajaxCarregando_loading">
   	<tr>
   	  <td align="center">
       	<strong>Aguarde carregando informações...</strong><br /><br />
        <table width="128" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><div id="ajaxCarregandoImg"></div></td>
          </tr>
      	</table>
      </td>
    </tr>
  </table>
</div>

<script type="text/javascript">
  $(document).ready(function() {
  	carregaListagem();
  });

  function carregaListagem() {
		var caminho = 'epv_renovacao_ajax.asp?ts=listar';
		ajax_loadContent('containerListagem',caminho,true,false,false);
		$("#TXT_DEVOLUCAO").val('');
		$("#TXT_DEVOLUCAO").focus();
	}

  function gravarItem(){
    var etiqueta = $("#TXT_DEVOLUCAO").val();
    if(etiqueta !== ''){
      var caminho = 'epv_renovacao_ajax.asp?ts=incluir&p1='+etiqueta;
    }
    ajax_loadContent('retorna_div',caminho,false,false,true);
  }

  function removerItemRenovacao(cod_formando) {
    var caminho = 'epv_renovacao_ajax.asp?ts=remover&p1='+cod_formando;
    ajax_loadContent('retorna_div',caminho,false,false,true);
  }

  function runScript(e,campo) {
    mascara(campo,soNumeros);

    if (e.keyCode == 13) {
      gravarItem();
      return false;
    }
  }

  function displayLeitura(tipo,msg) {
    $('#sucesso_msg').hide();
    $('#erro_msg').hide();

    if(tipo == 'erro'){
      $('#erro_msg').show();
      $('#erro_msg').html(msg);
    }else{
      $('#sucesso_msg').show();
      $('#sucesso_msg').html(msg);
    }
    $("#TXT_DEVOLUCAO").val('');		
  }

  function buscaIni() {
    d = document.buscaForm;
    d.busca.focus();
    d.busca.select();
  }

  function mudacorBusca(CAMPO) {
    CAMPO.style.backgroundColor = '#F5ECBA';
  }

  function mudacorBusca2(CAMPO) {
    CAMPO.style.backgroundColor = '#FFFFFF';
  }

  function Verificar() {
    var tecla=window.event.keyCode;
    if (tecla==17) {
      buscaIni();
    }
  }
  function confirmaDevolucao() {
    if (validaMotivos()) {
      var msg = "Tem certeza que deseja gravar a devolução destes itens?";
      if(YNconfirm(msg)){
          return true;	
      }else{
          return false;
      }	
    } else {
      alert('Preencha todos os campos de "motivo de devolução"');
      
      return false;
    }
  }

  function validaMotivos() {
    var motivos = document.DEVOLUCAO_ITENS.getElementsByTagName('textarea');
    
    var result = true;
    if (motivos.length > 0) {
      count = motivos.length;

      do {
        if (motivos[count-1].value == '') {
          result = false;
        }
        count --;
      } while (count > 0);
    }

    return result;
  }

  function YNconfirm(txt) {
    return confirm(txt);
  }
</script>
<!-- #include file="files/includes/miolo-inferior.asp" -->