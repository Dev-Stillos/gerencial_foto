<%


Function ConvertFromUTF8(sIn)

	CAcento = "àáâãäèéêëìíîïòóôõöùúûüÀÁÂÃÄÈÉÊËÌÍÎÒÓÔÕÖÙÚÛÜçÇñÑ"
	SAcento = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN"
	Texto = ""
	If sIn <> "" then
		     For X = 1 To Len(sIn)
		            Letra = Mid(sIn,X,1)
		            Pos_Acento = InStr(CAcento,Letra)
		            If Pos_Acento > 0 Then Letra = mid(SAcento,Pos_Acento,1)
		            Texto = Texto & Letra
		     Next
		     ConvertFromUTF8 = Texto
	End If
End Function


Function ConvertToUTF(sIn)
 texto  = Replace(sIn,"'","\'")
	Dim oIn: Set oIn = CreateObject("ADODB.Stream")

	oIn.Open
	oIn.CharSet = "ISO-8859-1"
	oIn.WriteText texto
	oIn.Position = 0
	oIn.CharSet = "UTF-8"
	ConvertToUTF = oIn.ReadText
	oIn.Close

End Function

Function ConvertFromISO(sIn)

	Dim oIn: Set oIn = CreateObject("ADODB.Stream")

	oIn.Open
	oIn.CharSet = "UTF-8"
	oIn.WriteText sIn
	oIn.Position = 0
	oIn.CharSet = "ISO-8859-1"
	ConvertFromISO = oIn.ReadText
	oIn.Close

End Function


%>
