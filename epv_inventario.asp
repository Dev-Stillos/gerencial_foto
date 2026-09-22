
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
Dim comando, TXT_DEVOLUCAO, CONT_ITENS
comando = Request.QueryString("comando")
TXT_DEVOLUCAO = Request.Form("TXT_DEVOLUCAO")
TXT_DEVOLUCAO = Replace(TXT_DEVOLUCAO," ","")
CONT_ITENS = 0

COD_USUARIO = sqlNumero(Session("COD_USUARIO"))

sub incluirRegistros()
  
  cod_item = Request.form("cod_item")
  cod_formando = sqlNumero(left(cod_item, len(cod_item)-1))
  tipo_produto = sqlNumero(right(cod_item, 1))

  sql = "INSERT INTO TB_INVENTARIO (COD_FORMANDO, ITEM, COD_USUARIO) VALUES ("&cod_formando&", "&tipo_produto&", "&cod_usuario&")" 
  abrirConexao()
    objConexao.execute(sql)
  fecharConexao()

  response.redirect("epv_inventario.asp")

end sub 

if comando = "incluir" then
  call incluirRegistros()
end if
%>

<h2>EPV - Inventário de Itens </h2>
<div id="retorna_div"></div>
<div class="subMenu">
  Sub-Menu: 
  <a href="epv_devolucao.asp?comando=nova">Iniciar devolução de Vendedor</a> | 
  <a href="epv_saida_inicio.asp?comando=nova">Saída de Álbuns do Estoque</a>
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
  action="?comando=incluir" 
  method="post" 
  name="IVENTARIO"
  onsubmit="return confirmaLancamento()">
  <div id="containerListagem">
    <br>
    <br>
    <br>
    <br>
  </div>
</form>
<h3 width="800px" align="center">Últimos Lançamentos</h3>
<table width="800px" border="1" align="center" 	cellpadding="5" style="border-collapse:collapse">
  <tr bgcolor="whitesmoke">
    <th>Cód. Formando</th>
    <th>Item</th>
    <th>Formando</th>
    <th>Instituição e Curso</th>
    <th>Produção</th>
    <th> </th>
  </tr>

  <%
  item_values = array("Albúm", "Ficha")
  abrirConexao()
  dim sql : sql = "SELECT  a.cod_formando, a.item, b.nome_formando, c.DESCRICAO, b.nome_curso, b.cod_producao FROM TB_INVENTARIO as a INNER JOIN TB_FICHAS_IDENTIFICACAO as b ON a.COD_FORMANDO = b.COD_FORMANDO INNER JOIN TB_ENTIDADES as c ON c.COD_CONTRATO = b.COD_CONTRATO WHERE COD_USUARIO = "&sqlNumero(cod_usuario)&" ORDER BY DATA DESC LIMIT 10"
  set rs = objConexao.execute(sql)
  if (not rs.Eof) then 
    while not rs.Eof
      %>              
        <tr>
          <td><%=rs("cod_formando").value%></td>
          <td><%=item_values(rs("item").value)%></td>
          <td><%=rs("nome_formando").value%></td>
          <td><%=rs("DESCRICAO").value%> / <%=rs("nome_curso").value%></td>
          <td><%=rs("COD_PRODUCAO").value%></td>
          <td> </td>
        </tr>
      <%
    rs.movenext()
    wend
  else
  %>
    <tr bgcolor="whitesmoke">
      <td colspan='6' style='text-align: center'>Este usuário não fez inventários</td>
    </tr>
  <%
  end if 
  rs.close()
  set rs = nothing
  fecharConexao()
  %>
</table>

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
  $(function(){
    $('#TXT_DEVOLUCAO').focus();
  });

  function runScript(e,campo) {
    mascara(campo,soNumeros);

    if (e.keyCode == 13) {
      gravarItem();
      return false;
    }
  }
  function gravarItem(){
    var etiqueta = $("#TXT_DEVOLUCAO").val();
    if(etiqueta !== ''){
      var caminho = 'epv_inventario_ajax.asp?ts=incluir&p1='+etiqueta;
    }
    ajax_loadContent('containerListagem',caminho,false,false,true);
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
  function confirmaLancamento() {
  
    var msg = "Tem certeza que deseja gravar o lançamento?";
    if(confirm(msg)){
        return true;	
    }else{
        return false;
    }	

  }

</script>
<!-- #include file="files/includes/miolo-inferior.asp" -->