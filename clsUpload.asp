<%
'==============================================================
' clsUpload.asp - Classe de Upload Corrigida para ASP Clássico
'==============================================================

Class FileUploader
    Public Form
    Public Files
    
    Private Sub Class_Initialize()
        Set Form = Server.CreateObject("Scripting.Dictionary")
        Set Files = Server.CreateObject("Scripting.Dictionary")
        
        If Request.TotalBytes > 0 Then
            ProcessUpload
        End If
    End Sub
    
    Private Sub ProcessUpload()
        Dim biData, sBoundary, vDataBounds, nDataBounds
        Dim wPosBegin, wPosEnd, wPos
        Dim sInputName, sFormValue, sFileName, sContentType
        Dim oFile
        
        ' Pegar os dados binários
        biData = Request.BinaryRead(Request.TotalBytes)
        
        ' Extrair boundary do Content-Type
        wPos = InStr(Request.ServerVariables("CONTENT_TYPE"), "boundary=")
        If wPos = 0 Then Exit Sub
        
        sBoundary = Mid(Request.ServerVariables("CONTENT_TYPE"), wPos + 9)
        sBoundary = "--" & sBoundary
        
        ' Converter boundary para binário
        Dim bBoundary
        bBoundary = StringToBinary(sBoundary)
        
        ' Encontrar primeira boundary
        nDataBounds = InStrB(1, biData, bBoundary)
        If nDataBounds = 0 Then Exit Sub
        
        ' Processar cada parte
        Do
            ' Encontrar próxima boundary
            wPosBegin = nDataBounds + LenB(bBoundary) + 2 ' +2 para CRLF
            wPosEnd = InStrB(wPosBegin, biData, bBoundary)
            
            If wPosEnd = 0 Then Exit Do
            
            ' Extrair dados da parte atual
            vDataBounds = MidB(biData, wPosBegin, wPosEnd - wPosBegin - 2) ' -2 para remover CRLF final
            
            ' Extrair headers
            wPos = InStrB(1, vDataBounds, StringToBinary(vbCrLf & vbCrLf))
            If wPos = 0 Then
                nDataBounds = wPosEnd
                ' CORREÇÃO: Substituir "Loop" por "GoTo NextBoundary" ou usar estrutura correta
                GoTo NextBoundary
            End If
            
            Dim sHeaders, sData
            sHeaders = BinaryToString(MidB(vDataBounds, 1, wPos - 1))
            sData = MidB(vDataBounds, wPos + 4) ' +4 para pular duplo CRLF
            
            ' Extrair nome do campo
            sInputName = ExtractFieldName(sHeaders)
            
            ' Verificar se é arquivo
            If InStr(sHeaders, "filename=") > 0 Then
                ' É um arquivo
                sFileName = ExtractFileName(sHeaders)
                sContentType = ExtractContentType(sHeaders)
                
                ' Criar objeto arquivo
                Set oFile = New UploadedFile
                oFile.FileName = sFileName
                oFile.FileData = sData
                oFile.ContentType = sContentType
                
                ' Adicionar ao dicionário de arquivos
                Set Files(sInputName) = oFile
            Else
                ' É um campo de formulário
                sFormValue = BinaryToString(sData)
                Form(sInputName) = sFormValue
            End If
            
NextBoundary:
            ' Próxima boundary
            nDataBounds = wPosEnd
        Loop
    End Sub
    
    ' Converter string para binário
    Private Function StringToBinary(sText)
        Dim i, sBinary
        For i = 1 To Len(sText)
            sBinary = sBinary & ChrB(Asc(Mid(sText, i, 1)))
        Next
        StringToBinary = sBinary
    End Function
    
    ' Converter binário para string
    Private Function BinaryToString(bData)
        Dim sResult, i
        For i = 1 To LenB(bData)
            sResult = sResult & Chr(AscB(MidB(bData, i, 1)))
        Next
        BinaryToString = sResult
    End Function
    
    ' Extrair nome do campo do header
    Private Function ExtractFieldName(sHeaders)
        Dim wPos, wPosEnd
        wPos = InStr(sHeaders, "name=""") + 6
        If wPos = 6 Then ExtractFieldName = "": Exit Function
        
        wPosEnd = InStr(wPos, sHeaders, """")
        ExtractFieldName = Mid(sHeaders, wPos, wPosEnd - wPos)
    End Function
    
    ' Extrair nome do arquivo do header
    Private Function ExtractFileName(sHeaders)
        Dim wPos, wPosEnd
        wPos = InStr(sHeaders, "filename=""") + 10
        If wPos = 10 Then ExtractFileName = "": Exit Function
        
        wPosEnd = InStr(wPos, sHeaders, """")
        ExtractFileName = Mid(sHeaders, wPos, wPosEnd - wPos)
    End Function
    
    ' Extrair content-type do header
    Private Function ExtractContentType(sHeaders)
        Dim wPos, wPosEnd
        wPos = InStr(sHeaders, "Content-Type: ") + 14
        If wPos = 14 Then ExtractContentType = "application/octet-stream": Exit Function
        
        wPosEnd = InStr(wPos, sHeaders, vbCrLf)
        If wPosEnd = 0 Then wPosEnd = Len(sHeaders) + 1
        
        ExtractContentType = Trim(Mid(sHeaders, wPos, wPosEnd - wPos))
    End Function
End Class

' Classe para arquivo enviado
Class UploadedFile
    Public FileName
    Public FileData
    Public ContentType
    
    Public Sub SaveAs(sPath)
        Dim oStream
        Set oStream = Server.CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1 'Binary
        oStream.Write FileData
        oStream.SaveToFile sPath, 2 'Overwrite
        oStream.Close
        Set oStream = Nothing
    End Sub
    
    Public Function GetSize()
        GetSize = LenB(FileData)
    End Function
End Class
%>