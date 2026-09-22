<%
Response.ContentType = "application/json"
Response.Write "{""sucesso"":true,""mensagem"":""Funcionando""}"
%>
<%
Response.ContentType = "application/json"
Response.CodePage = 65001

On Error Resume Next

Dim cod_formando, arquivos, file, fso, caminhoSalvar, nomeArquivo, i, IMG_PATH_WEB
cod_formando = Trim(Request.Form("cod_formando"))

If cod_formando = "" Then
    Response.Write "{""sucesso"":false,""mensagem"":""Código do formando é obrigatório.""}"
    Response.End
End If

Set fso = Server.CreateObject("Scripting.FileSystemObject")
caminhoSalvar = Server.MapPath("gerencial_foto/uploads")
If Not fso.FolderExists(caminhoSalvar) Then fso.CreateFolder(caminhoSalvar)

Set arquivos = Request.Files
If arquivos.Count = 0 Then
    Response.Write "{""sucesso"":false,""mensagem"":""Nenhuma foto foi enviada.""}"
    Response.End
End If

i = 0
Dim fotosSalvas: fotosSalvas = 0

For Each file In arquivos
    If i >= 300 Then Exit For
    If Not file Is Nothing Then
        If file.FileName <> "" Then
            nomeArquivo = cod_formando & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & i & "_" & Replace(file.FileName, " ", "_")
            file.SaveAs caminhoSalvar & nomeArquivo
            IMG_PATH_WEB = "gerencial_foto/uploads" & nomeArquivo

            ' Grava no banco
            AbrirConexao()
            sql = "INSERT INTO TB_RENOVACAO_DE_ALBUM_DIGITAL (COD_FORMANDO, IMG_ALBUM) VALUES (?, ?)"
            Set cmd = Server.CreateObject("ADODB.Command")
            With cmd
                .ActiveConnection = objConexao
                .CommandText = sql
                .Parameters.Append .CreateParameter("@formando", 200, 1, 50, cod_formando)
                .Parameters.Append .CreateParameter("@img", 200, 1, 255, IMG_PATH_WEB)
                .Execute
            End With
            Set cmd = Nothing
            FecharConexao()

            i = i + 1
            fotosSalvas = fotosSalvas + 1
        End If
    End If
Next

If fotosSalvas > 0 Then
    Response.Write "{""sucesso"":true,""mensagem"":""Upload concluído com " & fotosSalvas & " fotos.""""}"
Else
    Response.Write "{""sucesso"":false,""mensagem"":""Nenhuma foto foi salva no servidor.""}"
End If
%>
