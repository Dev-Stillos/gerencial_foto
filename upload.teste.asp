<!--#include file="clsUpload.asp" -->
<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<%
Response.Write "=== TESTE DE UPLOAD ===" & "<br>"
Response.Write "Content-Type: " & Request.ServerVariables("CONTENT_TYPE") & "<br>"
Response.Write "Content-Length: " & Request.ServerVariables("CONTENT_LENGTH") & "<br>"
Response.Write "Request Method: " & Request.ServerVariables("REQUEST_METHOD") & "<br>"
Response.Write "Total Bytes: " & Request.TotalBytes & "<br><br>"

' Debug adicional - verificar se há dados
If Request.TotalBytes = 0 Then
    Response.Write "❌ <strong>ERRO: Nenhum dado foi recebido (TotalBytes = 0)</strong><br>"
    Response.Write "Verifique se o formulário HTML tem enctype='multipart/form-data'<br><br>"
End If

' Verificar Content-Type
If InStr(Request.ServerVariables("CONTENT_TYPE"), "multipart/form-data") = 0 Then
    Response.Write "❌ <strong>ERRO: Content-Type não é multipart/form-data</strong><br>"
    Response.Write "Content-Type atual: " & Request.ServerVariables("CONTENT_TYPE") & "<br><br>"
End If

' Debug - mostrar primeiros bytes dos dados recebidos (se houver)
If Request.TotalBytes > 0 Then
    Response.Write "<strong>Primeiros 200 caracteres dos dados recebidos:</strong><br>"
    Response.Write "<pre style='background:#f0f0f0; padding:10px; font-size:11px;'>"
    
    Dim rawData, firstBytes, i
    rawData = Request.BinaryRead(Request.TotalBytes)
    
    ' Converter primeiros bytes para texto (para debug)
    firstBytes = ""
    For i = 1 To 200
        If i > LenB(rawData) Then Exit For
        firstBytes = firstBytes & Chr(AscB(MidB(rawData, i, 1)))
    Next
    
    Response.Write Server.HTMLEncode(firstBytes)
    Response.Write "</pre><br>"
    
    ' Recriar o uploader com os dados já lidos
    ' (Nota: Normalmente você não pode fazer Request.BinaryRead duas vezes,
    ' então vamos criar uma versão modificada)
End If

' Tentar criar o uploader
On Error Resume Next
Dim uploader
Set uploader = New FileUploader

If Err.Number <> 0 Then
    Response.Write "❌ <strong>ERRO ao criar FileUploader:</strong> " & Err.Description & "<br><br>"
    Err.Clear
End If
On Error Goto 0

If Not uploader Is Nothing Then
    ' Debug - mostrar todos os campos recebidos
    Response.Write "<strong>Campos do formulário recebidos:</strong><br>"
    If uploader.Form.Count > 0 Then
        Dim key
        For Each key In uploader.Form.Keys
            Response.Write key & " = [" & uploader.Form(key) & "]<br>"
        Next
    Else
        Response.Write "<em>Nenhum campo encontrado</em><br>"
    End If

    Response.Write "<br><strong>Arquivos recebidos:</strong><br>"
    If uploader.Files.Count > 0 Then
        For Each key In uploader.Files.Keys
            Response.Write key & " (arquivo: " & uploader.Files(key).FileName & ", tamanho: " & uploader.Files(key).GetSize() & " bytes)<br>"
        Next
    Else
        Response.Write "<em>Nenhum arquivo encontrado</em><br>"
    End If

    Response.Write "<br>Total de arquivos: " & uploader.Files.Count & "<br>"
    Response.Write "Total de campos: " & uploader.Form.Count & "<br><br>"

    ' Verifica se existem arquivos
    If uploader.Files.Count > 0 Then
        Response.Write "✓ Arquivos encontrados!<br>"
        
        ' Pegar o primeiro arquivo (independente do nome do campo)
        Dim arquivo, nomeArquivo, destino
        Set arquivo = Nothing
        
        ' Buscar por qualquer arquivo enviado
        For Each key In uploader.Files.Keys
            Set arquivo = uploader.Files(key)
            Exit For
        Next
        
        If Not arquivo Is Nothing Then
            nomeArquivo = arquivo.FileName
            Response.Write "Nome do arquivo: " & nomeArquivo & "<br>"
            Response.Write "Tamanho do arquivo: " & arquivo.GetSize() & " bytes<br>"
            Response.Write "Content-Type: " & arquivo.ContentType & "<br>"
            
            ' Verificar se o diretório uploads existe
            Dim fso, uploadsPath
            Set fso = Server.CreateObject("Scripting.FileSystemObject")
            uploadsPath = Server.MapPath("uploads")
            
            If Not fso.FolderExists(uploadsPath) Then
                Response.Write "⚠️ Diretório uploads não existe. Criando...<br>"
                fso.CreateFolder(uploadsPath)
            End If
            
            ' Criar nome único para evitar conflitos
            Dim extensao, nomeUnico
            extensao = ""
            If InStr(nomeArquivo, ".") > 0 Then
                extensao = Right(nomeArquivo, Len(nomeArquivo) - InStrRev(nomeArquivo, ".") + 1)
            End If
            nomeUnico = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & _
                       Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & "_" & nomeArquivo
            
            destino = Server.MapPath("uploads/" & nomeUnico)
            Response.Write "Destino: " & destino & "<br>"

            ' Tentar salvar o arquivo
            On Error Resume Next
            arquivo.SaveAs destino
            If Err.Number <> 0 Then
                Response.Write "❌ Erro ao salvar arquivo: " & Err.Description & " (Código: " & Err.Number & ")<br>"
            Else
                Response.Write "✅ Arquivo salvo com sucesso!<br>"
                
                ' Verificar se o arquivo foi realmente salvo
                If fso.FileExists(destino) Then
                    Response.Write "✅ Arquivo confirmado no disco. Tamanho: " & fso.GetFile(destino).Size & " bytes<br>"
                Else
                    Response.Write "❌ Arquivo não encontrado no disco após salvamento<br>"
                End If
            End If
            On Error Goto 0

            ' ========== Recuperar dados do formulário ==========
            Dim cod_formando, cod_vendedor, cod_contrato, status_venda, data_entrega, img_path
            
            cod_formando = Trim(uploader.Form("cod_formando"))
            cod_vendedor = Trim(uploader.Form("cod_vendedor"))
            cod_contrato = Trim(uploader.Form("cod_contrato"))
            status_venda = Trim(uploader.Form("status_venda"))
            'file_foto = Trim(uploader.Form("fotos")
            
            Response.Write "<br><strong>Dados do formulário:</strong><br>"
            Response.Write "Código Formando: [" & cod_formando & "]<br>"
            Response.Write "Código Vendedor: [" & cod_vendedor & "]<br>"
            Response.Write "Código Contrato: [" & cod_contrato & "]<br>"
            Response.Write "Status Venda: [" & status_venda & "]<br>"
            
            ' Validar campos obrigatórios
            If cod_formando = "" Or cod_vendedor = "" Or cod_contrato = "" Or status_venda = "" Then
                Response.Write "❌ Campos obrigatórios não preenchidos!<br>"
                Response.Write "Verifique se os nomes dos campos no formulário HTML estão corretos<br>"
            Else
                ' Converter status para número
                Dim status_num
                If LCase(status_venda) = "vendido" Then 
                    status_num = 1 
                Else 
                    status_num = 0 
                End If
                
                data_entrega = Now()
                img_path = "uploads/" & nomeUnico

                ' ========== Inserir no banco ==========
                Response.Write "<br><strong>Tentando inserir no banco...</strong><br>"
                
                On Error Resume Next
                AbrirConexao()
                If Err.Number <> 0 Then
                    Response.Write "❌ Erro ao abrir conexão: " & Err.Description & "<br>"
                Else
                    Response.Write "✓ Conexão aberta<br>"
                    
                    ' SQL com parâmetros seguros
                    Dim sql
                    sql = "INSERT INTO TB_RENOVACAO_DE_ALBUM_DIGITAL " & _
                        "(COD_FORMANDO, COD_VENDEDOR, COD_CONTRATO, STATUS_VENDA, DATA_ENTREGA, IMG_ALBUM, FOTO_FORMATO_PATH) VALUES (" & _
                        "'" & Replace(cod_formando, "'", "''") & "', " & _
                        cod_vendedor & ", '" & _
                        Replace(cod_contrato, "'", "''") & "', " & _
                        status_num & ", '" & _
                        data_entrega & "', '" & _
                        Replace(img_path, "'", "''") & "', '" & _
                        Replace(img_path, "'", "''") & "')"
                    
                    Response.Write "SQL: " & sql & "<br><br>"
                    
                    objConexao.Execute sql
                    If Err.Number <> 0 Then
                        Response.Write "❌ Erro ao executar SQL: " & Err.Description & " (Código: " & Err.Number & ")<br>"
                    Else
                        Response.Write "✅ Dados inseridos no banco com sucesso!<br>"
                        Response.Write "<script>setTimeout(function(){ alert('Upload realizado com sucesso!'); window.location.href='epv_renovacao_album.asp'; }, 2000);</script>"
                    End If
                    
                    FecharConexao()
                End If
                On Error Goto 0
            End If
            
            Set fso = Nothing
        Else
            Response.Write "❌ Nenhum arquivo válido encontrado<br>"
        End If
    Else
        Response.Write "❌ Nenhum arquivo foi enviado<br>"
        Response.Write "<br><strong>Possíveis causas:</strong><br>"
        Response.Write "1. Formulário HTML não tem enctype='multipart/form-data'<br>"
        Response.Write "2. Campo de arquivo não tem name definido<br>"
        Response.Write "3. Nenhum arquivo foi selecionado<br>"
        Response.Write "4. Problema na classe FileUploader<br>"
    End If

    Set uploader = Nothing
End If
%>