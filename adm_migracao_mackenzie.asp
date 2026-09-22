<%
response.Buffer = False
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 1
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
  Dim titulo_pagina : titulo_pagina = "Migração Mackenzie"
  Dim calendario_necessario : calendario_necessario = "NÃO"

%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%

'Declara variáveis para utilização do sistema
Dim campo1, comando, msg
Dim cod_colacao_importacao, ano_conclusao, semestre_conclusao, cod_producao, cod_contrato, cod_curso 

campo1 = Request.Form("campo1")
comando = Request.Form("comando")

If comando = Empty Then
  comando = "incluir"
  titulo_pagina = titulo_pagina&" / Inclusão"
End if

If comando = "incluir" and campo1 <> Empty Then

  'Gravar alteração
  migrarRegistro()

End If

Sub lerFormulario()

  cod_colacao_importacao = sqlNumero(Request.Form("cod_colacao"))
  ano_conclusao = sqlTexto(request.form("ano_conclusao"))
  semestre_conclusao = sqlTexto(request.form("semestre_conclusao"))
  cod_producao = sqlNumero(request.form("cod_producao"))
  cod_contrato = sqlTexto(request.form("cod_contrato"))
  cod_curso = sqlNumero(0)

End sub

Sub migrarRegistro()
	lerFormulario()

  AbrirConexao()

  dim album : album = "'0'"
  dim inicial : inicial = "'0'"
  dim sistema_novo : sistema_novo = 1
  dim count : count = 0
  

  dim rsMack, sql, sqlEndereco, rsCodNew
  Set rsMack = objConexao.Execute("SELECT * FROM stillos3.TB_MACK_FORMANDOS where numero_identificacao is not null and numero_identificacao <> ''  and cod_mack_colacao = "&cod_colacao_importacao)
  while not rsMack.Eof
	
	  'GERAMOS O CODIGO DO FORMANDO	
	  objConexao.Execute("INSERT INTO TB_FICHAS_IDENTIFICACAO_CONTAGEM_NUMERO_FORMANDO_CADASTRADO (COD_FORMANDO_CADASTRADO)   VALUES (NULL)")
	  Set rsCodNew = objConexao.Execute("SELECT LAST_INSERT_ID() AS ULTIMO_ID FROM  TB_FICHAS_IDENTIFICACAO_CONTAGEM_NUMERO_FORMANDO_CADASTRADO LIMIT 1")
	  if not rsCodNew.Eof then
	  	cod_formando = sqlNumero( rsCodNew("ULTIMO_ID") & 1)
	  End if
	  rsCodNew.Close
	  Set rsCodNew = Nothing

	  'FORMANDO
  	nome_formando = sqlTexto(replace(rsMack("nome"),"'"," "))
    nome_curso = sqlTexto(ucase(rsMack("curso")))
  	retirado = "''"
  	brinde = "''"
  	comissao = "''"
  	produto_agregado = "''"

    'INSERT DO FORMANDO
  	sql = "INSERT INTO TB_FICHAS_IDENTIFICACAO (COD_FORMANDO ,SISTEMA_NOVO ,COD_PRODUCAO ,COD_CONTRATO ,COD_CURSO ,NOME_CURSO ,NOME_FORMANDO ,ANO_CONCLUSAO ,SEMESTRE_CONCLUSAO ,ALBUM ,INICIAL ,RETIRADO ,BRINDE ,COMISSAO ,PRODUTO_AGREGADO )VALUES ("& COD_FORMANDO & ", "& sistema_novo &", "& cod_producao &", "& cod_contrato &", "& cod_curso &", "& nome_curso &", "& nome_formando &", "& ano_conclusao &", "& semestre_conclusao &", "& album &", "& inicial &", "& retirado &", "& brinde &", "& comissao &", "& produto_agregado &")"
    'response.write(sql & "<BR>"& "<BR>" )
    objConexao.Execute(sql)

  	'ENDERECO
  	ENDERECO = sqlTexto(replace(rsMack("residencia_endereco"),"'"," "))
  	CEP = sqlTexto(rsMack("residencia_cep"))
  	FONE_RES = sqlTexto(rsMack("telefone"))
  	FONE_CEL = sqlTexto(rsMack("celular"))
  	BAIRRO = sqlTexto(replace(rsMack("residencia_bairro"),"'"," "))
  	CIDADE = sqlTexto(replace(rsMack("residencia_cidade"),"'"," "))
  	ESTADO = sqlTexto(replace(rsMack("residencia_estado"),"'"," "))

  	'INSERT DO ENDERECO DO FORMANDO
  	sqlEndereco =	"INSERT INTO TB_FICHAS_IDENTIFICACAO_ENDERECO_FORMANDO(COD_FORMANDO, ENDERECO, CEP, FONE_RES, FONE_CEL, BAIRRO, CIDADE, ESTADO, CADASTRADO_POR)VALUES ("& COD_FORMANDO &", "& ENDERECO &", "& CEP &", "& FONE_RES &", "& FONE_CEL &", "& BAIRRO &", "& CIDADE &", "& ESTADO &" , 1)"
  	'response.write(sqlEndereco& "<BR>"& "<BR>")
    objConexao.Execute(sqlEndereco)

    count = count + 1
  rsMack.movenext()
  Wend
  rsMack.close
  Set rsMack = nothing

	FecharConexao()
End sub

'----------------------------------------------------------------------------------------
%>

<div style='style="border:dotted 1px #999999'>
  <h2><%=titulo_pagina%></h2>
  <br />

  <% if campo1 <> Empty then %>
    <div class='msg'>
      Registros importados: <%=count%>
    </div>
  <% end if %>

  <form method='POST' name='FORMULARIO_PADRAO'>
    <table width='80%' border='1'  cellpadding='5' align='center' style='border-collapse:collapse; margin-bottom: 10px;'>
      <tr>
        <td width="30%" class='text-bold text-right'>Colação a ser importada:</td>
        <td width="70%">
          <%=padraoFormularioSelect("COD_COLACAO", "", "inputTamanho3",""," required='required' ","Selecione", cod_mack_colacao,"SELECT cod_mack_colacao, concat(data_colacao, descricao) FROM TB_MACK_COLACOES ORDER BY cod_mack_colacao desc", "incluir", false, false, false) %>
        </td>
      </tr>
      <tr>
        <td width="30%" class='text-bold text-right'>Contrato Que receberá os formandos:</td>
        <td width="70%">
          <%=padraoFormularioSelect("COD_CONTRATO", "", "inputTamanho3",""," required='required' ","Selecione", cod_contrato,"SELECT cod_contrato, descricao FROM TB_ENTIDADES WHERE cod_contrato like '%f' ORDER BY cod_contrato", "incluir", false, false, false) %>
        </td>
      </tr>
      <tr>
        <td width="30%" class='text-bold text-right'>Produção que receberá os formandos:</td>
        <td width="70%">
          <%=padraoFormularioSelect("COD_PRODUCAO", "", "inputTamanho3",""," required='required' ","Selecione", cod_producao,"SELECT cod_producao, concat(cod_producao, ' - ',descricao) FROM TB_PRODUCAO ORDER BY cod_producao desc", "incluir", false, false, false) %>
        </td>
      </tr>
      <tr>
        <td class='text-bold text-right'>Ano Conclusão</td>
        <td>
          <input type='text' name='ano_conclusao' class='inputTamanho2' />
        </td>
      </tr>
      <tr>
        <td class='text-bold text-right'>Semestre Conclusão</td>
        <td>
          <input type='text' name='semestre_conclusao' class='inputTamanho2' />
        </td>
      </tr>
      <tr>
        <td colspan='2' align='center'><a href="javascript:salvar();"><img src="images/layout/btnGravar.jpg" /></a></td>
      </tr>
    </table>
    <input type='hidden' name='campo1' value='1' />
    <input type='hidden' name='comando' value='<%=comando%>' />
  </form>
  <script type='text/javascript'>
    function validaFormMaster() {
      if(validaForm('COD_COLACAO#Colação a ser importada#2|COD_CONTRATO#Cód. Contrato#2|COD_PRODUCAO#Cód. Produção#2|ano_conclusao#Ano Conclusão#1|semestre_conclusao#Semestre Conclusão#1',document.FORMULARIO_PADRAO)){return true}else{return false}
    }
</script>
</div>

<!-- #include file="files/includes/miolo-inferior.asp" -->

