
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

If comando = "nova" Then
'Apaga tudo da base do determinado usu�rio
AbrirConexao()
	objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_DEVOLUCAO_TEMPORARIA WHERE COD_USUARIO = "&Session("COD_USUARIO"))
FecharConexao()
End If

If comando = "excluir" Then
	AbrirConexao()
		objConexao.Execute("DELETE FROM TB_FICHAS_IDENTIFICACAO_DEVOLUCAO_TEMPORARIA WHERE COD_ETIQUETA = "&Request.QueryString("cod_etiqueta"))
		Response.Redirect("epv_devolucao.asp")
	FecharConexao()
End If

If comando = "efetuar_devolucao" Then
  AbrirConexao()

  Dim rsUp, COD_FORMANDO_UP, TIPO_UP
  Dim COD_MOVIMENTACAO_UP
  dim lista_cod_movimentacao : lista_cod_movimentacao = ""
  dim lista_producao : lista_producao = ""
  count = 0
  
  dim COD_VENDEDOR : COD_VENDEDOR = Request.form("COD_VENDEDOR")

  Set rsUp = objConexao.Execute("SELECT COD_ETIQUETA FROM TB_FICHAS_IDENTIFICACAO_DEVOLUCAO_TEMPORARIA WHERE COD_USUARIO="&Session("COD_USUARIO")&" GROUP BY COD_ETIQUETA")

  While not rsUP.Eof
    TIPO_UP	= Right(rsUP("COD_ETIQUETA"),1)
    COD_FORMANDO_UP = Mid(rsUP("COD_ETIQUETA"),1,Len(rsUP("COD_ETIQUETA"))-1)
    COD_MOVIMENTACAO_UP	= Request.Form("COD_MOVIMENTACAO_"&COD_FORMANDO_UP&"_"&TIPO_UP)

    'INFORMCOES DA TABELA TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES
    lista_cod_movimentacao  = lista_cod_movimentacao & " '"&COD_MOVIMENTACAO_UP&"',"
    count = count + 1

    'INFORMCOES DA TABELA TB_PRODUCAO_LOTES_FORMANDOS
    'ATUALIZAMOS SOMENTE SE FOR devolução DE FICHAS
    if (CStr(TIPO_UP) = "1") then
      lista_producao = lista_producao & sqlTexto(COD_FORMANDO_UP) &","

      motivo = motivo & "(" & sqlTexto(COD_FORMANDO_UP) &","& sqlTexto(Session("COD_USUARIO")) &", " & sqlTexto(Request.Form("motivo_"&COD_FORMANDO_UP)) & ", "& COD_VENDEDOR &"),"
    end if

  rsUp.MoveNext()
  Wend
  rsUp.Close()
  Set rsUp = Nothing

  'Executando o update - Adicionado pelo Luiz
  lista_cod_movimentacao = Left(lista_cod_movimentacao,Len(lista_cod_movimentacao)-1)
  sql = "UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET DATA_RETORNO = NOW() WHERE COD_MOVIMENTACAO IN ("&lista_cod_movimentacao&") LIMIT "&count

  'ATENCAO VERIFICAR SE � PARA DEVOLVER AUTOMATICAMENTE ALBUNS 
  'objConexao.Execute("UPDATE TB_FICHAS_IDENTIFICACAO_MOVIMENTACOES SET DATA_PREVISTA_RETORNO = NOW() WHERE COD_FORMANDO ='"&rs("COD_FORMANDO")&"' AND TIPO <> 1 AND DATA_RETORNO IS NULL")

  objConexao.Execute(sql)

  if len(lista_producao) > 0 then
    lista_producao = Left(lista_producao,Len(lista_producao)-1)
    sqlProducao = "UPDATE TB_PRODUCAO_LOTES_FORMANDOS SET COD_VENDEDOR = 0 WHERE STATUS = 3 AND COD_FORMANDO IN ("& lista_producao &") AND COD_VENDEDOR = "&sqlNumero(COD_VENDEDOR)

    objConexao.Execute(sqlProducao)
  end if

  if len(motivo)  > 0 then
    motivo = Left(motivo,Len(motivo)-1)
    sqlMotivosHist = "INSERT INTO TB_FICHAS_IDENTIFICACAO_HISTORICO (COD_FORMANDO, GERADO_POR, MENSAGEM, COD_VENDEDOR)  VALUES " & motivo

    objConexao.Execute(sqlMotivosHist)
  end if

  FecharConexao()
	Response.Redirect("epv_devolucao_imprimir.asp")
End If
%>
<h2>EPV - Devolu&ccedil;&atilde;o de Produtos </h2>
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
<%
AbrirConexao()
'1) Verifica se h� pendencias de devolução
'2) Verifica seleciona todas as fichas
%>
<form 
  action="epv_devolucao.asp?comando=efetuar_devolucao" 
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
		var caminho = 'epv_movimentacoes_estoque_ajax.asp?ts=listar';
		ajax_loadContent('containerListagem',caminho,true,false,false);
		$("#TXT_DEVOLUCAO").val('');
		$("#TXT_DEVOLUCAO").focus();
	}

  function gravarItem(){
    var etiqueta = $("#TXT_DEVOLUCAO").val();
    if(etiqueta !== ''){
      var caminho = 'epv_movimentacoes_estoque_ajax.asp?ts=incluir&p1='+etiqueta;
      ajax_loadContent('retorna_div',caminho,false,false,true);
    }			
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
    console.log('aaaaa');
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