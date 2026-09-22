<%
Function gerarQualquerNumeroAleatorio()
	gerarQualquerNumeroAleatorio = "12312312S09801928DQLLQW"& Replace(Replace(Replace(NOW(),":",""),"/","")," ","")
End Function

'FUNÇÃO QUE REMOVE TAGS DE UM CODIGO HTML ESPECIFICO
'PODE SER CASO NÃO SEJA INFORMANDA NENHUMA TAG A SER REMOVIDA SERÃO REMOVIDAS AS TAGS SCRIPT E STYLE
'ATENÇÃO NÃO REMOVE TAGS UNICAS COMO BR E IMG 
function trataHTML(codigo_html,tags_remover)

	dim retorno, tags, abertura_tag, fechamento_tag, texto_esquerda, texto_direita,  aux, fim, tags_unicas(10), i, ii

	tags_unicas(0) = "br"
	tags_unicas(1) = "img"
	tags_unicas(2) = "input"
	tags_unicas(3) = "area"
	tags_unicas(4) = "base"
	tags_unicas(5) = "basefont"
	tags_unicas(6) = "hr"
	tags_unicas(7) = "link"
	tags_unicas(8) = "meta"
	
	aux = codigo_html
	
	if tags_remover = "" then
		tags = split("script|style","|")
	else
		tags = split(tags_remover,"|")
	end if
	
	for i=0 to ubound(tags)
		abertura_tag = InStr(1, aux, "<"&tags(i), vbTextCompare)
		fechamento_tag = InStr(1, aux, "</"&tags(i)&">", vbTextCompare)
				
		if fechamento_tag > 0 then
			fechamento_tag = fechamento_tag + len(tags(i))+3
		end if
		if len(aux) < fechamento_tag then
			fim = fechamento_tag - len(aux) +1
		else
			fim = len(aux)  - fechamento_tag+1
		end if
		
		if 0 < abertura_tag then
			texto_esquerda = left(aux, abertura_tag-1)			
			texto_direita = mid(aux, fechamento_tag, fim)
			aux = texto_esquerda&texto_direita
		end if	

	next	
	retorno = trim(aux)
	
	trataHTML = retorno
end function


function alteraParametroHTML(codigo_html,parametros,novos_valores)

	dim retorno, tags, abertura_parametro, fechamento_parametro, texto_esquerda, texto_direita,  aux, fim, i, ii, parametros_, novos_valores_
	
	parametros_ = split(parametros,"|")
	novos_valores_ = split(novos_valores,"|")
	
	aux = codigo_html
	
	for i=0 to ubound(parametros_)	
		abertura_parametro = InStr(1, aux, parametros_(i)&"=""", vbTextCompare)+len(parametros_(i)&"=""")
		fechamento_parametro = InStr(abertura_parametro, aux, """", vbTextCompare)
		
		if len(aux) < fechamento_parametro then
			fim = fechamento_parametro - len(aux) +1
		else
			fim = len(aux)  - fechamento_parametro+1
		end if
		
		if 0 < abertura_parametro then
			texto_esquerda = left(aux, abertura_parametro-1)			
			texto_direita = mid(aux, fechamento_parametro, fim)
			aux = texto_esquerda&novos_valores_(i)&texto_direita
		end if	

	next

	retorno = trim(aux)
	
	alteraParametroHTML = retorno
end function

%>