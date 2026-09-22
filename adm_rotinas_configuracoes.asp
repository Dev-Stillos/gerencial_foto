<%
  Dim BTN_VOLTAR, BTN_IMPRIMIR, BTN_FONTE, BTN_IMPRIMIR_LINK
  BTN_VOLTAR = 1
  BTN_IMPRIMIR = 1
  BTN_FONTE = 2
  BTN_IMPRIMIR_LINK = "#"
  
  Dim titulo_pagina : titulo_pagina = "Aprovações Pendentes"

  Dim calendario_necessario : calendario_necessario = "NÃO"

%>
<!-- #include file="files/includes/miolo-superior.asp" -->
<%

'-------------------------------------------------------------------------------

%>
<div style='style="border:dotted 1px #999999'>
  <h2><%=titulo_pagina%></h2><br />

  <div class='subMenu'>
    <a href='adm_configuracoes.asp'> Configurações do sistema </a> | 
    <a href='adm_migracao_mackenzie.asp'> Migrar mackenzie </a>
  </div>

  <div>
  
  </div>
</div>
<!-- #include file="files/includes/miolo-inferior.asp" -->