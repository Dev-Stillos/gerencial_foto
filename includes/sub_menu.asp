<%
' a flag jquery informa que já temos importado o jquery'
if not jquery then %>
<!-- a-->
<script type="text/javascript" src="Scripts/jquery-1.8.3.min.js"></script>
<%end if%>
<script type="text/javascript">
$('.linkMenu').mouseover(function() {
	$(this).css('cursor','pointer');
});
$('.linkMenu').mouseout(function() {
	$(this).css('cursor','');
});
</script>


<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="31%">
	<div>
	<%If BTN_VOLTAR = 1 Then%>
		<a href="javascript:parent.history.go(-1)"><img alt="Voltar" class="botoes_icones" src="images/icons/icon_voltar.jpg" width="32" height="28" border="0" /></a>&nbsp;
	<%End If%>
	
	<%If BTN_IMPRIMIR = 1 Then%>
	<a <%=BTN_IMPRIMIR_LINK%>><img alt="Imprimir" src="images/icons/icon_print.jpg" class="botoes_icones" width="32" height="28" border="0" /></a>&nbsp;	
		
	<%End If%>

	  	
	</div>
	</td>
    <td width="69%"><div align="right" style="color:#333333"><table width="550" border="0" align="right" cellpadding="3" cellspacing="0">
          <tr>
            <td><div align="right">
            <%
			 AbrirConexao()
		If int(Session("COD_PERFIL")) = 5 Then	
			
		 
		  Dim rsRetContato, qtdRetornar
		  qtdRetornar= 0
		  
		  Set rsRetContato=objConexao.Execute("Select count(*) as qtd FROM TB_FICHAS_IDENTIFICACAO_RETORNAR_CONTATO where COD_USUARIO = '"&Session("cod_usuario")&"' or COD_USUARIO = 0")
		  		
				If not rsRetContato.Eof Then
					qtdRetornar = formatNumber(0&rsRetContato("qtd"),0)
					
					If qtdRetornar > 0 Then
					%>
                  	<a class="linkCinza" href="dip_formandos_entrar_em_contato.asp">Entrar em contato(<%=qtdRetornar%>)</a>
            &nbsp;&nbsp;&nbsp;
                      
                    
                <%
					End If
				End If
		  
		  rsRetContato.Close
		  Set rsRetContato=Nothing
		  
		  End If
		  
		  Dim rsEtiquetasAgendadas,qtdAgendada
		  qtdAgendada = 0
		  Set rsEtiquetasAgendadas = objConexao.Execute("SELECT COD_USUARIO FROM TB_FICHAS_IDENTIFICACAO_ETIQUETAS_TEMPORARIA WHERE COD_USUARIO="&Session("COD_USUARIO")&" AND AGENDADA=1",qtdAgendada)
			
			If int(qtdAgendada) > 0 Then
			%>
            <img src="images/icons/icon_cod_bar.jpg" width="15" height="11" />&nbsp;<a class="linkCinza" href="adm_etiquetas_agendadas.asp">Etiquetas Agendadas(<%=qtdAgendada%>)</a>
            &nbsp;&nbsp;&nbsp;
            <%End If
			
			rsEtiquetasAgendadas.Close
			 Set rsEtiquetasAgendadas = Nothing
			 
			 FecharConexao()
			%>
            
            <img src="images/icons/icon_help_mini.jpg" width="15" height="11" /> <a class="linkCinza" href="manual_sistema.asp" target="_blank">Manual</a>&nbsp;&nbsp;&nbsp;
            <img src="images/icons/icon_member.jpg" width="15" height="11" /> <a class="linkCinza" href="adm_log_usuarios.asp"><%=Session("NOME")%></a>&nbsp;&nbsp;&nbsp;<img src="images/icons/icon_abandon.jpg" width="12" height="11" /> <a class="linkCinza" href="adm_sair.asp">Sair</a></div></td>
          </tr>
          <tr>
            <td> <div align="right">
                <%
	  Dim DATA_NOME_DIA, DATA_NOME_MES
	  
	  DATA_NOME_DIA = WeekdayName(Weekday(now()))
	  DATA_NOME_MES = MonthName(Month(now()))
	  
	  DATA_NOME_DIA = Replace(Left(DATA_NOME_DIA,1),Left(DATA_NOME_DIA,1),UCase(Left(DATA_NOME_DIA,1))& Mid(DATA_NOME_DIA,2,50))
	  %>
	  
                <%Response.Write(DATA_NOME_DIA &", "&day(now())& " de " &DATA_NOME_MES &" de "& Year(now()))%>
            .</div></td>
          </tr>
        </table>
      <br />
      <br />
    </div></td>
  </tr>
</table>