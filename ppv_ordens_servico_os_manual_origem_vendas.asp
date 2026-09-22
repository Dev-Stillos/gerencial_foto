<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 2
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  titulo_pagina = "PPV - Produtos - Etapas de Ordens de Servi&ccedilos - Formul&aacuterio"
%>
				
<%
Dim calendario_necessario
calendario_necessario = "N&AtildeO"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<!-- #include file="includes/classe_modulo_ppv.asp" -->
<%
Dim cod_pedido, sucesso, codmsg
cod_pedido = Cstr(""&Request.Form("cod_pedido"))
cod_pedido = Replace(cod_pedido,".","")
cod_pedido = Replace(cod_pedido,"/","")
cod_pedido = Replace(cod_pedido," ","")
cod_pedido = Replace(cod_pedido,"-","")

	If cod_pedido <> "" Then
		
		AbrirConexao()
		
		Set rs=objConexao.Execute("SELECT cod_pedido FROM TB_PEDIDOS WHERE cod_pedido = "&sqlTexto(cod_pedido)&" and FINALIZA_CORRECAO = 1")
			if not rs.Eof then
				
				retorna = moduloPPVgerarOrdensServico(cod_pedido,true)
				'Response.Write(retorna)
				If int(retorna) = 1 Then
					sucesso = ""
					codmsg = 1	
				Else
					sucesso = 2
					codmsg = retorna
				End If				

			eLSE
				sucesso = 2
				codmsg = 100
			end if
		rs.close
		Set rs=nothing	
		
		FecharConexao()
		
		Response.Redirect("ppv_ordens_servico_os_manual_origem_vendas.asp?cod_pedido="&cod_pedido&"&codmsg="&codmsg&"&sucesso="&sucesso)
		
	End if

		
'----------------------------------------------------------------------------------------

%>
<h2><%=titulo_pagina%></h2>
<br />

<!-- #include file="includes/classe_modulo_ppv_menu.asp" -->
<br />
<br />
<%
	
	codmsg	= CStr(""&Request.QueryString("codmsg"))
	sucesso = CStr(""&Request.QueryString("sucesso"))
	cod_pedido = CStr(""&Request.QueryString("cod_pedido"))
		
	If sucesso = "" and codmsg <> "" Then%>
    	<div align="center" class="msgGrandeSucesso">
        	<%
			If codmsg = "1" Then
				Response.Write("Ordens de Servi&ccedilo geradas com sucesso para a Venda "&cod_pedido&"!<br / >OSs geradas de acordo com as regras informadas abaixo")
			End If	
			%>
            
        
        </div>
    <%ElseIf sucesso <> "" and codmsg <> "" Then%>
        <div align="center" class="msgGrandeErro">
			<%
			If codmsg = "100" Then
				Response.Write("Pedido "&cod_pedido&" n&atildeo encontrado e/ou n&atildeo est&aacute corrigido pelos respons&aacuteveis.")
			Elseif codmsg = "101" Then
				Response.Write("OS j&aacute gerada anteriormente para o pedido "&cod_pedido&". Para gerar novamente para o mesmo pedido exclua se poss&iacutevel ou cancele as OS da venda.")
			Else		
				Response.Write(Response.Write(moduloPPVgerarOrdensServicoMsgRetorno(codmsg,cod_pedido)))
			End If	
			%>
</div>
	<%End If%>

<form action='<%=actionformulario%>' name='FORMULARIO_PADRAO' method='post'>
  <table width="600" border="1" align="center" cellpadding="5" style="border-collapse:collapse">
    <tr bgcolor="whitesmoke">
    <td><strong>Informe o C&oacute;d. da Venda abaixo para Efetuar a gera&ccedil;&atilde;o Manual das Ordens de Servi&ccedil;o da Venda.<br />
      Regras:<br />
      1) 
      Se j&aacute; tiver OS criada anteriormente n&atilde;o ser&aacute; gerado, exceto a OS estiver com status Cancelada.
      <br />
      2) Pedido dever&aacute; ter sido j&aacute; corrigido pelas pessoas respons&aacute;veis.<br />
      3) Os geradas somente para produtos vendidos &agrave; parte com prazo do cliente cadastrado
    </strong></td>
    </tr>
            <tr>
              <td><div align="center">
                <input type="text" class="inputTamanho3" name="cod_pedido" />
              </div></td>
              </tr>
            <tr>
              <td><div align="center"> <a href="javascript:document.FORMULARIO_PADRAO.submit();">
                <img src="images/layout/btnGravar.jpg" />
              </a></div></td>
              </tr>
    
          </table>
<br />
<br /></form>
<!-- #include file="files/includes/miolo-inferior.asp" -->
