<%
Class ParseMultipartFormDataClass
    Public Form, Files

    Private Sub Class_Initialize()
        Set Form = Server.CreateObject("Scripting.Dictionary")
        Set Files = Server.CreateObject("Scripting.Dictionary")

        Dim Boundary, FullBody, PosIni, PosFim, BinData
        Dim ContentType, i

        ContentType = Request.ServerVariables("HTTP_CONTENT_TYPE")
        If InStr(ContentType, "multipart/form-data") = 0 Then Exit Sub

        Boundary = "--" & Mid(ContentType, InStr(ContentType, "boundary=") + 9)
        BinData = Request.BinaryRead(Request.TotalBytes)
        FullBody = BinaryToString(BinData)

        PosIni = InStr(1, FullBody, Boundary)
        Do While PosIni > 0
            PosFim = InStr(PosIni + 1, FullBody, Boundary)
            If PosFim = 0 Then Exit Do

            Dim Part, Disposition, FieldName, FileName, Value, ContentDisp, ContentTypeLine
            Part = Mid(FullBody, PosIni + Len(Boundary) + 2, PosFim - PosIni - Len(Boundary) - 4)
            ContentDisp = GetLine(Part, 1)
            ContentTypeLine = GetLine(Part, 2)

            FieldName = ParseFieldName(ContentDisp)
            FileName = ParseFileName(ContentDisp)

            If FileName <> "" Then
                ' É arquivo
                Dim BinPart, FileStart
                FileStart = InStr(Part, vbCrLf & vbCrLf) + 4
                BinPart = MidB(BinData, BytePos(FullBody, Part) + FileStart - 1, LenB(Mid(Part, FileStart)))

                Files.Add FileName, CreateUploadFile(FileName, BinPart)
            Else
                ' Campo comum
                Value = Mid(Part, InStr(Part, vbCrLf & vbCrLf) + 4)
                Value = Left(Value, Len(Value) - 2)
                Form.Add FieldName, Value
            End If

            PosIni = PosFim
        Loop
    End Sub

    Private Function GetLine(text, num)
        Dim lines, i
        lines = Split(text, vbCrLf)
        If num >= 0 And num < UBound(lines) Then
            GetLine = lines(num)
        Else
            GetLine = ""
        End If
    End Function

    Private Function ParseFieldName(header)
        Dim posIni, posFim
        posIni = InStr(header, "name=""") + 6
        posFim = InStr(posIni, header, """")
        ParseFieldName = Mid(header, posIni, posFim - posIni)
    End Function

    Private Function ParseFileName(header)
        Dim posIni, posFim
        posIni = InStr(header, "filename=""")
        If posIni = 0 Then
            ParseFileName = ""
        Else
            posIni = posIni + 10
            posFim = InStr(posIni, header, """")
            ParseFileName = Mid(header, posIni, posFim - posIni)
        End If
    End Function

    Private Function BytePos(strFull, strPart)
        BytePos = LenB(Left(strFull, InStr(strFull, strPart) - 1)) + 1
    End Function

    Private Function BinaryToString(BinData)
        Dim Stream
        Set Stream = Server.CreateObject("ADODB.Stream")
        Stream.Type = 1 'binary
        Stream.Open
        Stream.Write BinData
        Stream.Position = 0
        Stream.Type = 2 'text
        Stream.Charset = "UTF-8"
        BinaryToString = Stream.ReadText
        Stream.Close
        Set Stream = Nothing
    End Function

    Private Function CreateUploadFile(FileName, BinContent)
        Dim oFile
        Set oFile = New UploadFile
        oFile.FileName = Right(FileName, Len(FileName) - InStrRev(FileName, "\"))
        oFile.BinaryContent = BinContent
        Set CreateUploadFile = oFile
    End Function
End Class

Class UploadFile
    Public FileName
    Public BinaryContent

    Public Sub SaveAs(path)
        Dim fStream
        Set fStream = Server.CreateObject("ADODB.Stream")
        fStream.Type = 1 'binary
        fStream.Open
        fStream.Write BinaryContent
        fStream.SaveToFile path, 2
        fStream.Close
        Set fStream = Nothing
    End Sub
End Class

Function ParseMultipartFormData()
    Dim parser
    Set parser = New ParseMultipartFormDataClass
    Set ParseMultipartFormData = parser
End Function
%>
