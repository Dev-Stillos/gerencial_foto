<%
'EXEMPLO PADRรO DE UTILIZAวรO TROCAR O CONTEUDO "PASTA_PROJETO" pela pasta real

					Set Pdf = Server.CreateObject("Persits.Pdf")
					Set Doc = Pdf.OpenDocument("e:\home\stillos\Web\gerencial_foto\documentos\MANUAL_USUARIO.pdf")
					Doc.SaveHttp "attachment;filename=manual.pdf" 
					Set Doc = Nothing
					Set Pdf = Nothing
%>