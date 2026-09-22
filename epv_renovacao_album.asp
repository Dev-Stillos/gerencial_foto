
<%
Response.ContentType = "text/html"
Response.Charset = "UTF-8"
%>

 
<%
'=========================================
' EPV - Renovação de Álbuns Digitais e Fichas
'=========================================

Dim calendario_necessario, BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
calendario_necessario = "NÃO"
BTN_VOLTAR = 1
BTN_IMPRIMIR = 2
BTN_FONTE = 1
BTN_IMPRIMIR_LINK = "#"

' Include cabeçalho
<!-- #include file="includes/sub_menu.asp" -->
 <!-- #include file="includes/menu.asp" -->
  <!-- #include file="includes/topo.asp" -->

Dim comando, TXT_DEVOLUCAO, ERROR_LINHA, pronto, CONT_ITENS
pronto = Request.QueryString("pronto")
ERROR_LINHA = 1
comando = Request.QueryString("comando")
TXT_DEVOLUCAO = Request.Form("TXT_DEVOLUCAO")
TXT_DEVOLUCAO = Replace(TXT_DEVOLUCAO," ","")
CONT_ITENS = 0

Function sqlNumero(valor)
    If IsNumeric(valor) Then
        sqlNumero = CInt(valor)
    Else
        sqlNumero = 0
    End If
End Function


'--- Nova devolução: limpa temporários ---
If comando = "nova" Then
    
    objConexao.Execute("DELETE FROM TB_RENOVACAO_ITENS_DIGITAIS WHERE COD_USUARIO = "&COD_USUARIO)
    FecharConexao()
End If

'--- Efetivar devolução/renovação ---
If comando = "efetuar_devolucao" Then
    
    Dim renovacoes
    renovacoes = ""

    Set rs = objConexao.Execute("SELECT * FROM TB_RENOVACAO_ITENS_DIGITAIS WHERE COD_USUARIO=" & COD_USUARIO)
    While Not rs.Eof
        codRenovacao = sqlNumero(rs("COD_RENOVACAO_TEMP"))
        codFormando = sqlNumero(rs("COD_FORMANDO"))
        codVendedor = sqlNumero(rs("COD_VENDEDOR"))
        periodoRenovar = sqlNumero(rs("PERIODO_RENOVACAO"))
        novoPrazo = sqlData(DateAdd("d", periodoRenovar, Date()))
        motivo = sqlTexto(Request.Form("motivo_" & codFormando & rs("TIPO").value))
        tipo = sqlNumero(rs("TIPO"))

        ' Atualiza prazo e motivo
        objConexao.Execute "UPDATE TB_RENOVACAO_ITENS_DIGITAIS SET DATA_PRAZO=" & novoPrazo & ", MOTIVO=" & motivo & " WHERE COD_RENOVACAO_TEMP=" & codRenovacao
        
        ' Marca para impressão
        renovacoes = renovacoes & codRenovacao & ","
        
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing

    If Len(renovacoes) > 0 Then
        renovacoes = Left(renovacoes, Len(renovacoes)-1)
        objConexao.Execute "UPDATE TB_RENOVACAO_ITENS_DIGITAIS SET PODE_IMPRIMIR=1 WHERE COD_RENOVACAO_TEMP IN (" & renovacoes & ")"
    End If
    FecharConexao()

    Response.Redirect("epv_renovacao_imprimir.asp")
End If
%>

<h2>EPV - Renovação de Álbuns Digitais e Fichas</h2>

<div id="retorna_div"></div>
<div class="subMenu">
    Sub-Menu: 
    <a href="epv_renovacao_album.asp?comando=nova">Iniciar renovação/devolução</a>
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
            <input onkeypress="return runScript(event,this);" maxlength="8" style="height:50px; width:200px; font-size:34px;" type="text" name="FORM_TXT_DEVOLUCAO" id="TXT_DEVOLUCAO" placeholder="Código do item"/>
        </div>
    </td>
</tr>
<tr>
    <td>
        <div align="center">
            <input class="btnGravar" type="button" onclick="gravarItem()" id="botaoGravar" value="Gravar Item" />
            <br /><br />
        </div>
    </td>
</tr>
</table>

<form action="?comando=efetuar_devolucao" method="post" name="DEVOLUCAO_ITENS" style="margin:0px" onsubmit="return confirmaDevolucao()">
    <div id="containerListagem"></div>
</form>

<div id="ajaxCarregandoContainer" class="ajaxCarregando" style="visibility:hidden">
    <table height="100%" width="100%">
    <tr>
      <td align="center">
        <strong>Aguarde carregando informações...</strong><br /><br />
        <div id="ajaxCarregandoImg"></div>
      </td>
    </tr>
    </table>
</div>

<script type="text/javascript">
$(document).ready(function() {
    carregaListagem();
});

function carregaListagem() {
    var caminho = 'epv_renovacao_digital_ajax.asp?ts=listar';
    ajax_loadContent('containerListagem',caminho,true,false,false);
    $("#TXT_DEVOLUCAO").val('');
    $("#TXT_DEVOLUCAO").focus();
}

function gravarItem(){
    var etiqueta = $("#TXT_DEVOLUCAO").val();
    if(etiqueta !== ''){
        var caminho = 'epv_renovacao_digital_ajax.asp?ts=incluir&p1='+etiqueta;
        ajax_loadContent('retorna_div',caminho,false,false,true);
    }
}

function removerItemRenovacao(cod_formando) {
    var caminho = 'epv_renovacao_digital_ajax.asp?ts=remover&p1='+cod_formando;
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
      $('#erro_msg').show().html(msg);
    } else {
      $('#sucesso_msg').show().html(msg);
    }
    $("#TXT_DEVOLUCAO").val('');		
}

function confirmaDevolucao() {
    if(validaMotivos()) {
        return confirm("Tem certeza que deseja gravar a devolução destes itens?");
    } else {
        alert('Preencha todos os campos de "motivo de devolução"');
        return false;
    }
}

function validaMotivos() {
    var motivos = document.DEVOLUCAO_ITENS.getElementsByTagName('textarea');
    var result = true;
    for(var i=0;i<motivos.length;i++){
        if(motivos[i].value==''){ result=false; break; }
    }
    return result;
}
</script>

<!-- #include file="files/includes/miolo-inferior.asp" -->
