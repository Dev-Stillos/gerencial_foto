<!-- upload_album_form.asp -->
<!-- #include file="files/includes/miolo-superior.asp" -->
<%
Dim cod_renovacao_temp
cod_renovacao_temp = Request.QueryString("cod_renovacao_temp")

If Not IsNumeric(cod_renovacao_temp) Then
  Response.Write("<div style='color:red;'>Código inválido!</div>")
  Response.End()
End If
%>

<h2>Upload de Foto do Álbum</h2>

<form action="epv_renovacao_album.asp" method="post" enctype="multipart/form-data">
  <input type="hidden" name="cod_renovacao_temp" value="<%=cod_renovacao_temp%>">
  <label for="foto_album">Selecione a imagem:</label><br>
  <input type="file" name="foto_album" accept=".jpg,.jpeg,.png,.gif" required><br><br>
  <input type="submit" value="Enviar">
</form>

<p><a href="epv_renovacao_album.asp">Voltar</a></p>

<!-- #include file="files/includes/miolo-inferior.asp" -->
