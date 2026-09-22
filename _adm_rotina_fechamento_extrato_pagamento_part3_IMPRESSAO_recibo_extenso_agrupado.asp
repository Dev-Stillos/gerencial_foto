<!-- #include file="includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/osdifuqwslasdaplplasvalida_usuario.asp" -->
<!-- #include file="includes/classe_fechamento.asp" -->
<%
Dim DIA_VENCIMENTO, COD_FECHAMENTO_PERIODO
COD_FECHAMENTO_PERIODO = Request.QueryString("COD_FECHAMENTO_PERIODO")

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
Dim data_pagamento

AbrirConexao()

	Set rs=objConexao.Execute("Select * FROM TB_FECHAMENTO_PAGAMENTO_PERIODOS WHERE cod_fechamento_periodo = "&cod_fechamento_periodo)
		If not rs.Eof Then
			data_pagamento = ArrumaData(rs("data_pagamento"))
		End If
	rs.Close
	Set rs=Nothing

FecharConexao()
'////////////////////////////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Dim x_Centavos, x_I, x_J, x_K, x_Numero, x_QtdCentenas, x_TotCentenas, x_TxtExtenso( 900 ), x_TxtMoeda( 6 ), x_ValCentena( 6 ), x_Valor, x_ValSoma
' Matrizes de textos
x_TxtMoeda( 1 ) = "rea"
x_TxtMoeda( 2 ) = "mil"
x_TxtMoeda( 3 ) = "milh"
x_TxtMoeda( 4 ) = "bilh"
x_TxtMoeda( 5 ) = "trilh"

x_TxtExtenso( 1 ) = "um"
x_TxtExtenso( 2 ) = "dois"
x_TxtExtenso( 3 ) = "tres"
x_TxtExtenso( 4 ) = "quatro"
x_TxtExtenso( 5 ) = "cinco"
x_TxtExtenso( 6 ) = "seis"
x_TxtExtenso( 7 ) = "sete"
x_TxtExtenso( 8 ) = "oito"
x_TxtExtenso( 9 ) = "nove"
x_TxtExtenso( 10 ) = "dez"
x_TxtExtenso( 11 ) = "onze"
x_TxtExtenso( 12 ) = "doze"
x_TxtExtenso( 13 ) = "treze"
x_TxtExtenso( 14 ) = "quatorze"
x_TxtExtenso( 15 ) = "quinze"
x_TxtExtenso( 16 ) = "dezesseis"
x_TxtExtenso( 17 ) = "dezessete"
x_TxtExtenso( 18 ) = "dezoito"
x_TxtExtenso( 19 ) = "dezenove"
x_TxtExtenso( 20 ) = "vinte"
x_TxtExtenso( 30 ) = "trinta"
x_TxtExtenso( 40 ) = "quarenta"
x_TxtExtenso( 50 ) = "cinquenta"
x_TxtExtenso( 60 ) = "sessenta"
x_TxtExtenso( 70 ) = "setenta"
x_TxtExtenso( 80 ) = "oitenta"
x_TxtExtenso( 90 ) = "noventa"
x_TxtExtenso( 100 ) = "cento"
x_TxtExtenso( 200 ) = "duzentos"
x_TxtExtenso( 300 ) = "trezentos"
x_TxtExtenso( 400 ) = "quatrocentos"
x_TxtExtenso( 500 ) = "quinhentos"
x_TxtExtenso( 600 ) = "seiscentos"
x_TxtExtenso( 700 ) = "setentos"
x_TxtExtenso( 800 ) = "oitocentos"
x_TxtExtenso( 900 ) = "novecentos"

' Função Principal de Conversão de Valores em Extenso
Function Extenso( x_Numero )
    x_Numero = FormatNumber( x_Numero , 2 )
    x_Centavos = right( x_Numero , 2 )
    x_ValCentena( 0 ) = 0
    x_QtdCentenas = Int( ( Len( x_Numero ) + 1 ) / 4 )

    For x_I = 1 To x_QtdCentenas
        x_ValCentena( x_I ) = ""
    Next
    '
    x_I = 1
    x_J = 1
    For x_K = Len( x_Numero ) - 3 To 1 Step -1
        x_ValCentena( x_J ) = mid( x_Numero , x_K , 1 ) & x_ValCentena( x_J )
        If x_I / 3 = Int( x_I / 3 ) Then
            x_J = x_J + 1
            x_K = x_K - 1
        End If
        x_I = x_I + 1
    Next
    x_TotCentenas = 0
    Extenso = ""
    For x_I = x_QtdCentenas To 1 Step -1

        x_TotCentenas = x_TotCentenas + Int( x_ValCentena( x_I ) )

        If Int( x_ValCentena( x_I ) ) <> 0 Or ( Int( x_ValCentena( x_I ) ) = 0 And x_I = 1 )Then
            If Int( x_ValCentena( x_I ) = 0 And Int( x_ValCentena( x_I + 1 ) ) = 0 And x_I = 1 )Then
                Extenso = Extenso & ExtCentena( x_ValCentena( x_I ) , x_TotCentenas ) & " de " & x_TxtMoeda( x_I )
            Else
                Extenso = Extenso & ExtCentena( x_ValCentena( x_I ) , x_TotCentenas ) & " " & x_TxtMoeda( x_I )
            End If
            If Int( x_ValCentena( x_I ) ) <> 1 Or ( x_I = 1 And x_TotCentenas <> 1 ) Then
                Select Case x_I
                    Case 1
                        Extenso = Extenso & "is"
                    Case 3, 4, 5
                        Extenso = Extenso & "ões"
                End Select
            Else
                Select Case x_I
                    Case 1
                        Extenso = Extenso & "l"
                    Case 3, 4, 5
                        Extenso = Extenso & "ão"
                End Select
            End If
        End If
        If Int( x_ValCentena( x_I - 1 ) ) = 0 Then
            Extenso = Extenso
        Else
            If ( Int( x_ValCentena( x_I + 1 ) ) = 0 And ( x_I + 1 ) <= x_QtdCentenas ) Or ( x_I = 2 ) Then
                Extenso = Extenso & " e "
            Else
                Extenso = Extenso & ", "
            End If
        End If
    Next

    If x_Centavos > 0 Then
        If Int( x_Centavos ) = 1 Then
                Extenso = Extenso & " e " & ExtDezena( x_Centavos ) & " centavo"
        Else
                Extenso = Extenso &  " e " & ExtDezena( x_Centavos ) & " centavos"
        End If
    End If
    Extenso = UCase( Left( Extenso , 1 ) )&right( Extenso , Len( Extenso ) - 1 )
End Function

Function ExtDezena( x_Valor )
' Retorna o Valor por Extenso referente à DEZENA recebida
ExtDezena = ""
If Int( x_Valor ) > 0 Then
    If Int( x_Valor ) < 20 Then
        ExtDezena = x_TxtExtenso( Int( x_Valor ) )
    Else
        ExtDezena = x_TxtExtenso( Int( Int( x_Valor ) / 10 ) * 10 )
        If ( Int( x_Valor ) / 10 ) - Int( Int( x_Valor ) / 10 ) <> 0 Then
            ExtDezena = ExtDezena & " e " & x_TxtExtenso( Int( right( x_Valor , 1 ) ) )
        End If
    End If
End If
End Function

Function ExtCentena( x_Valor, x_ValSoma )
ExtCentena = ""

If Int( x_Valor ) > 0 Then

    If Int( x_Valor ) = 100 Then
        ExtCentena = "cem"
    Else
        If Int( x_Valor ) < 20 Then
            If Int( x_Valor ) = 1 Then
                If x_ValSoma - Int( x_Valor ) = 0 Then
                    ExtCentena = "hum"
                Else
                    ExtCentena = x_TxtExtenso( Int( x_Valor ) )
                End If
            Else
                    ExtCentena = x_TxtExtenso( Int( x_Valor ) )
            End If
        Else
            If Int( x_Valor ) < 100 Then
                ExtCentena = ExtDezena( right( x_Valor , 2 ) )
            Else
                ExtCentena = x_TxtExtenso( Int( Int( x_Valor ) / 100 )*100 )
                If ( Int( x_Valor ) / 100 ) - Int( Int( x_Valor ) / 100 ) <> 0 Then
                    ExtCentena = ExtCentena & " e " & ExtDezena( right( x_Valor , 2 ) )
                End If
            End If
        End If
    End If
End If
End Function
%>
<html>
<head>
<title>RECIBO_PAGAMENTO_<%=ArrumaData(data_pagamento)%></title>
</link>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css" media="all">
<!--
.style1 {
	font-size: 18px;
	font-weight: bold;
}
.folha {
    page-break-after: always;
	
}
body,td,th {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 18px;
}
.style2 {
	font-size: 27px;
	font-weight: bold;
}
-->
</style>
<style type="text/css">
<!--
.style2 {font-weight: bold}
.style6 {font-size: 11px; font-weight: bold; }
-->
</style>
</head>
<script language="javascript">
	

	function mClk2(src, url)
			{    
				//if(event.srcElement.tagName=='TD')
				//	{
				//		src.children.tags('A')[0].click(); 
				//	}
			location.href = url;

			}

</script>
<body onLoad="window.print()">
<%


AbrirConexao()

Dim rsTotos
Dim NOME_VENDEDOR, COD_VENDEDOR
Dim TOTAL_A_RECEBER_LIQUIDO, TODOS_PEDIDOS_POR_VENDEDOR, TODOS_PEDIDOS_BOLETOS_POR_VENDEDOR

'Obs.: Nos casos dos vendedores coordenadores 7 e 14 os mesmos deverão aparecer mesmo que não tenham venda

Set rsTodos = objConexao.Execute("SELECT "&_
								 "TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR ,"&_
								 "TB_FECHAMENTO_PAGAMENTO.COD_GRUPO_VENDEDORES,"&_
								 "TB_VENDEDORES.NOME_VENDEDOR "&_
								 "FROM TB_VENDEDORES INNER JOIN TB_FECHAMENTO_PAGAMENTO ON TB_VENDEDORES.COD_VENDEDOR = TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR "&_
								 "WHERE (TB_FECHAMENTO_PAGAMENTO.cod_fechamento_periodo = "&cod_fechamento_periodo&" AND TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR <> 27) OR TB_VENDEDORES.COD_VENDEDOR IN (SELECT cod_vendedor FROM TB_FECHAMENTO_PENDENCIAS_PAGAMENTO WHERE cod_fechamento_periodo = '"&COD_FECHAMENTO_PERIODO&"') or TB_VENDEDORES.COD_VENDEDOR IN(7,14,48)  "&_
								 "GROUP BY TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR ORDER BY NOME_VENDEDOR")

			While not rsTodos.Eof
			
				NOME_VENDEDOR				= Ucase(rsTodos("NOME_VENDEDOR"))
				COD_VENDEDOR 				= rsTodos("COD_VENDEDOR")
				TOTAL_A_RECEBER_LIQUIDO		= 0
				TODOS_PEDIDOS_POR_VENDEDOR  = ""
                TODOS_PEDIDOS_BOLETOS_POR_VENDEDOR = ""
				'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
				'TODOS_PEDIDOS_POR_VENDEDOR
				'seleciona todos os pedidos por vendedor
				
					qtdRegistros = 0
				
					Set rs=objConexao.Execute("SELECT COD_PEDIDO FROM TB_FECHAMENTO_PAGAMENTO WHERE COD_VENDEDOR = "&COD_VENDEDOR&" AND cod_fechamento_periodo ="&cod_fechamento_periodo&" ORDER BY COD_PEDIDO",qtdRegistros)
					
					While not rs.Eof
					
						qtdRegistros = 	qtdRegistros - 1
						
						If qtdRegistros = 0 then
							TODOS_PEDIDOS_POR_VENDEDOR = TODOS_PEDIDOS_POR_VENDEDOR & rs("COD_PEDIDO")
						else
							TODOS_PEDIDOS_POR_VENDEDOR = TODOS_PEDIDOS_POR_VENDEDOR & rs("COD_PEDIDO")&", "
						End If	
					
					rs.MoveNext()
					Wend
					
					rs.Close
					Set rs=Nothing


                    Set rs=objConexao.Execute("SELECT COD_PEDIDO, parcela FROM TB_FECHAMENTO_PAGAMENTO_BOLETOS WHERE COD_VENDEDOR = "&COD_VENDEDOR&" AND cod_fechamento_periodo ="&cod_fechamento_periodo&" ORDER BY COD_PEDIDO",qtdRegistros)
                    While not rs.Eof
                        qtdRegistros = qtdRegistros - 1
                        
                        If qtdRegistros = 0 then
                            TODOS_PEDIDOS_BOLETOS_POR_VENDEDOR = TODOS_PEDIDOS_BOLETOS_POR_VENDEDOR & rs("parcela") & "/" & rs("COD_PEDIDO")
                        else
                            TODOS_PEDIDOS_BOLETOS_POR_VENDEDOR = TODOS_PEDIDOS_BOLETOS_POR_VENDEDOR & rs("parcela") & "/" & rs("COD_PEDIDO")&", "
                        End If  
                    
                    rs.MoveNext()
                    Wend
                    
                    rs.Close
                    Set rs=Nothing
									  
'#################################################################################################################################				
				'1) TOTAL GERAL POR VENDEDOR LÍQUIDO
				TOTAL_A_RECEBER_LIQUIDO = fechamentoRetornaValorRecebidoVendedor(COD_VENDEDOR,cod_fechamento_periodo,true)
							
'#################################################################################################################################			
%>
    <table width="98%" border="0" align="center" cellpadding="20" cellspacing="0" class="folha">
      <tr>
      <td>
        <D><p> </p><p> </p><p> </p><p> </p><p> </p><p> </p>
        <p align="center" class="style1 style2" style="border-bottom: 1px #666666 solid; margin:15px">RECIBO</p>
        <p class="style1"></p>
        <p align="justify">Recebi da empresa: Stillo’s Reportagens  Fotográficas S/C Ltda., a importância de <strong><%=formatCurrency(TOTAL_A_RECEBER_LIQUIDO)%></strong> (<%=Extenso(TOTAL_A_RECEBER_LIQUIDO)%>) referente <strong>a comissão nas vendas </strong> dos pedidos: <%=TODOS_PEDIDOS_POR_VENDEDOR%> <% if TODOS_PEDIDOS_BOLETOS_POR_VENDEDOR <> "" then %> e também da compensasão dos boletos: <%= TODOS_PEDIDOS_BOLETOS_POR_VENDEDOR %>  <% end if %>

        <% If int(COD_VENDEDOR) = 14 or int(COD_VENDEDOR) = 7 or int(COD_VENDEDOR) = 48 Then %> e também referente comissão das vendas dos pedidos para os coordenadores de repasse<%End If%>.</p>
        <p> </p>
        <p align="left">Sem nada a reclamar, assino.</p>
        <p align="left"> </p>
       <div align="left"><strong>São  Paulo, <%=day(data_pagamento)%> de <%=MonthName(Month(data_pagamento))%> de <%=Year(data_pagamento)%>.</strong>
        <p align="left"> </p>
        <p align="left">-------------------------------------------</p>
        <div align="left"><em><%=NOME_VENDEDOR%></em>
        </div>
        <p> </p>
        <p> </p></TD>
      </tr>
    </table>

<%
rsTodos.MoveNext()
Wend
			
rsTodos.Close
Set rsTodos = Nothing
FecharConexao()
%>


</body>
</html>



