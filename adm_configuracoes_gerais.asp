<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 1
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
  Dim titulo_pagina
  titulo_pagina = "Configura&ccedil;&otilde;es Gerais"
%>
<%
Dim calendario_necessario
calendario_necessario = "N&Atilde;O"
%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%

'----------------------------------------------------------------------------------------
Dim qtdIPSHabilitados
qtdIPSHabilitados = 5 'Qtd Habilitada para este cliente


'Declara vari�veis para utiliza��o do sistema
Dim campo1, comando, redirecionar_alterar_gravar, redirecionar_excluir_registro, msg, actionformulario
Dim IP_FORMULARIO(20), descricao(20)

redirecionar_alterar_gravar		= "adm_configuracoes_gerais.asp"
redirecionar_excluir_registro	= "adm_configuracoes_gerais.asp"
actionformulario				= "adm_configuracoes_gerais.asp"

campo1 = Request.Form("campo1")
campo2 = Request.Form("campo2")
comando = "alterar"

If comando = Empty Then
comando = "alterar"
titulo_pagina = titulo_pagina
End if

If comando = "alterar" and campo1 = Empty Then
'Leitura inicial para poss�vel altera��o
lerRegistro()
titulo_pagina = titulo_pagina
Elseif 	comando = "alterar" and campo1 <> Empty Then
'Gravar altera��o
alterarRegistro(redirecionar_alterar_gravar)
titulo_pagina = titulo_pagina
End If


Sub lerFormulario()

	For i=0 to qtdIPSHabilitados-1
		IP_FORMULARIO(i)	= Request.Form("IP_FORMULARIO"&i)
		IP_FORMULARIO(i)	= Replace(IP_FORMULARIO(i)," ","")
		descricao(i)		= Request.Form("descricao"&i)
	Next
	
End sub

Sub lerRegistro()
	'Este cliente possui liberado 3 ips
	AbrirConexao()
		Set rs=objConexao.Execute("SELECT * FROM TB_SISTEMAS_IPS")
		
					For i=0 to qtdIPSHabilitados-1
					
						If not rs.Eof Then
							IP_FORMULARIO(i) = rs("ip_habilitado")
							descricao(i) = rs("descricao")
						rs.MoveNext()	
						End If
					Next
				
		rs.Close
		Set rs=Nothing
	FecharConexao()
	
End sub

Sub alterarRegistro(redirecionar)
	lerFormulario()
		
	AbrirConexao()
	
	'Excluir na base
	objConexao.Execute("DELETE FROM TB_SISTEMAS_IPS")
	
		For i=0 to qtdIPSHabilitados-1
		
			If IP_FORMULARIO(i) <> Empty Then
				objConexao.Execute("INSERT INTO TB_SISTEMAS_IPS(ip_habilitado,cadastrado_por,descricao) VALUES ("&sqlTexto(Trim(IP_FORMULARIO(i)))&",'',"&sqlTexto(descricao(i))&")")
			End if				
						
		Next
	
	FecharConexao()
	Response.Redirect(redirecionar&"?msg=Configura��es de IP alteradas com sucesso!")
End sub

'----------------------------------------------------------------------------------------

%>

<h2><%=titulo_pagina%></h2><br />
<%
	Dim msg_verifica
	msg_verifica = ""&Request.QueryString("msg")

		
	If msg_verifica <> "" Then%>
    	<div align="center" class="msgGrandeSucesso">
			<%=Response.Write(msg_verifica)%>          
        </div>
	<%End If%>
</div>
<script type="text/javascript">
function validaFormMaster(){
	return true;
}
</script>
<!-- #include file="includes/adm_usuarios_sub_menu.asp" -->
<br /><br />
<form action='<%=actionformulario%>?comando=<%=comando%>' name='FORMULARIO_PADRAO' method='post'>
<table width='98%' border='1' align='center' cellpadding='5' style='border-collapse:collapse'>
<tr>
<td colspan="4"><h3>Valida&ccedil;&atilde;o por IP</h3></td>
</tr>
<%
For i=0 to qtdIPSHabilitados-1
%>
<tr>
<td width='20%'><strong>IP Habilitado <%=i+1%>:</strong></td>
<td width='31%'>
<div align='left'>
<input type="text" class="inputTamanho2" name="IP_FORMULARIO<%=i%>" value="<%=IP_FORMULARIO(i)%>" />
</div></td>
<td width='17%'><span style="font-weight: bold">Descri&ccedil;&atilde;o do IP <strong><%=i+1%></strong>:</span></td>
<td width='32%'><div align='left'>
  <input type="text" class="inputTamanho2" name="descricao<%=i%>" value="<%=descricao(i)%>" />
</div></td>
</tr>
<%
Next
%>

<tr>
<td width='20%'><strong>IP do local que voc&ecirc; est&aacute; acessando agora:</strong></td>
<td colspan="3"><div align="left" style="font-weight: bold"><%=Request.ServerVariables("REMOTE_ADDR")%></div></td>
</tr>
<tr>
  <td colspan="4">Obs.: Se voc&ecirc; estiver acessando da rede da empresa e o n&uacute;mero de IP gravado for diferente do atual. Ent&atilde;o atualize para o n&uacute;mero correto para permitir acesso a todos os funcion&aacute;rios com acesso restrito a rede da empresa. Lembrando que voc&ecirc; possui espa&ccedil;o para <%=qtdIPSHabilitados%> IPs diferentes.</td>
  </tr>
<tr>
  <td colspan="4"><div align="center"><a href="javascript:salvar();"><img src="images/layout/btnGravar.jpg" /></a></div></td>
</tr>
</table>

<input type='hidden' value='1' name='campo1'  />
</form>
<br /><br />
<!-- #include file="files/includes/miolo-inferior.asp" -->
<script src='js/jquery-1.11.3.min.js'></script>
<script src='js/jQueryRotate.js'></script>


<style type="text/css">
@-moz-keyframes spin { 100% { -moz-transform: rotate(360deg); } }
@-webkit-keyframes spin { 100% { -webkit-transform: rotate(360deg); } }
@keyframes spin { 100% { -webkit-transform: rotate(360deg); transform:rotate(360deg); } }

</style>