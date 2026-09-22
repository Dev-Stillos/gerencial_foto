<!--#include file="includes/xqrtopqqetyynmkgfdconexao.asp"-->
<%
Response.CodePage = 65001
Response.Charset = "UTF-8"
Session.CodePage = 65001

Function SplitB(bin, delim)
    Dim arrText, stream
    Set stream = Server.CreateObject("ADODB.Stream")
    stream.Type = 1 ' binário
    stream.Open
    stream.Write bin
    stream.Position = 0
    stream.Type = 2 ' texto
    stream.Charset = "iso-8859-1"
    arrText = stream.ReadText
    stream.Close
    Set stream = Nothing
    SplitB = Split(arrText, delim)
End Function

Function CStrB(bstr)
    Dim stream
    Set stream = Server.CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write bstr
    stream.Position = 0
    stream.Type = 2
    stream.Charset = "UTF-8"
    CStrB = stream.ReadText
    stream.Close
    Set stream = Nothing
End Function

Function GetFileName(contentDisposition)
    Dim matches, regex
    Set regex = New RegExp
    regex.Pattern = "filename=""(.+?)"""
    regex.IgnoreCase = True
    regex.Global = False
    Set matches = regex.Execute(contentDisposition)
    If matches.Count > 0 Then
        GetFileName = matches(0).SubMatches(0)
    Else
        GetFileName = ""
    End If
End Function

Dim binData, boundary, parts, i, headers, content, fieldName, fileName, formFields
Set formFields = Server.CreateObject("Scripting.Dictionary")

binData = Request.BinaryRead(Request.TotalBytes)
boundary = MidB(binData, 1, InStrB(binData, ChrB(13)) - 1)
parts = SplitB(binData, CStrB(boundary))

Set fso = Server.CreateObject("Scripting.FileSystemObject")
caminhoFisico = Server.MapPath("uploads/")
If Not fso.FolderExists(caminhoFisico) Then fso.CreateFolder(caminhoFisico)

For i = 1 To UBound(parts) - 1
    Dim item, headerBodySplit, headerRaw, nomeFinal, stream
    item = parts(i)
    headerBodySplit = SplitB(item, ChrB(13) & ChrB(10) & ChrB(13) & ChrB(10))

    If UBound(headerBodySplit) < 1 Then
        ' ignora parte inválida
    Else
        headerRaw = CStrB(headerBodySplit(0))
        content = MidB(headerBodySplit(1), 1, LenB(headerBodySplit(1)) - 2)
        fieldName = ""
        fileName = ""

        If InStr(headerRaw, "Content-Disposition") > 0 Then
            Dim nameStart, nameEnd
            nameStart = InStr(headerRaw, "name=""") + 6
            nameEnd = InStr(nameStart, headerRaw, """")
            fieldName = Mid(headerRaw, nameStart, nameEnd - nameStart)

            If InStr(headerRaw, "filename=""") > 0 Then
                fileName = GetFileName(headerRaw)
                If fileName <> "" Then
                    nomeFinal = Year(Now)&Month(Now)&Day(Now)&"_"&Hour(Now)&Minute(Now)&Second(Now)&"_"&Replace(fileName, " ", "_")
                    Set stream = Server.CreateObject("ADODB.Stream")
                    stream.Type = 1
                    stream.Open
                    stream.Write content
                    stream.SaveToFile caminhoFisico & "\" & nomeFinal, 2
                    stream.Close
                    Set stream = Nothing
                    formFields.Add "foto", "uploads/" & nomeFinal
                End If
            Else
                formFields.Add fieldName, CStrB(content)
            End If
        End If
    End If
Next

' ========== Recuperar dados ==========
cod_formando = formFields("cod_formando")
cod_vendedor = formFields("cod_vendedor")
cod_contrato = formFields("cod_contrato")
status_venda = formFields("status_venda")
If LCase(status_venda) = "vendido" Then status_venda = 1 Else status_venda = 0
data_entrega = Now()
img_path = formFields("foto")

' ========== Inserir no banco ==========
AbrirConexao()
sql = "INSERT INTO TB_RENOVACAO_DE_ALBUM_DIGITAL " & _
      "(COD_FORMANDO, COD_VENDEDOR, COD_CONTRATO, STATUS_VENDA, DATA_ENTREGA, IMG_ALBUM, FOTO_FORMATO_PATH) VALUES (" & _
      cod_formando & ", " & cod_vendedor & ", " & cod_contrato & ", " & status_venda & ", '" & _
      data_entrega & "', '" & img_path & "', '" & img_path & "')"
On Error Resume Next
objConexao.Execute sql
If Err.Number <> 0 Then
    Response.Write "<script>alert('Erro ao salvar no banco: " & Replace(Err.Description,"'","") & "'); history.back();</script>"
Else
    Response.Write "<script>alert('Cadastro realizado com sucesso!'); window.location.href='epv_renovacao_album.asp';</script>"
End If
FecharConexao()
%>
