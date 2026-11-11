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
    
    if x_numero < 0 then
       EXIT Function
    END IF
    
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
                        Extenso = Extenso & "Yes"
                End Select
            Else
                Select Case x_I
                    Case 1
                        Extenso = Extenso & "l"
                    Case 3, 4, 5
                        Extenso = Extenso & "No"
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
' Retorna o Valor por Extenso referente � DEZENA recebida
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
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>RECIBO_PAGAMENTO_<%=ArrumaData(data_pagamento)%></title>

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
<style>
* { box-sizing:border-box; }
body { margin:0; background:#f6f7fb; color:#242738; font:14px/1.45 system-ui,Arial,sans-serif; }
.cap-document { max-width:1400px; margin:14px auto; padding:clamp(12px,3vw,26px); border:1px solid #e4dfeb; border-radius:12px; background:#fff; }
.cap-doc-header { display:flex; align-items:center; justify-content:space-between; gap:12px; padding-bottom:12px; margin-bottom:16px; border-bottom:2px solid #b20b61; }
.cap-doc-header h1 { margin:0; font:700 clamp(18px,3vw,27px)/1.2 system-ui,sans-serif; }
.cap-doc-logo { flex:none; width:clamp(95px,20vw,150px); height:auto; background:#292b40; border-radius:7px; }
.cap-doc-toolbar { display:flex; justify-content:flex-end; margin-bottom:10px; }
.cap-doc-button { display:inline-block; padding:8px 12px; border:1px solid #d8d0e1; border-radius:7px; background:#fff; color:#88134f; font:600 13px/1.3 system-ui,sans-serif; cursor:pointer; }
.cap-doc-primary { background:#b20b61; color:#fff; border-color:#b20b61; }
.cap-doc-scroll { overflow-x:auto; }
.cap-doc-scroll table { border-collapse:collapse; border:0; }
.cap-doc-scroll th, .cap-doc-scroll td { padding:6px 5px; border-color:#e4dfeb; font-size:inherit; }
.cap-doc-scroll th { background:#f3f0f7; color:#242738; }
.cap-doc-scroll .folha { width:100%; margin:12px 0; border:1px solid #e4dfeb; border-radius:8px; }
.cap-doc-scroll .folha td { padding:14px; }
.cap-doc-scroll .folha p { margin:8px 0; }
.cap-boleto-filter { display:flex; flex-wrap:wrap; gap:10px; align-items:flex-end; margin:0 0 14px; padding:12px; background:#f7f5fa; border-radius:8px; }
.cap-boleto-filter label { display:block; font-weight:600; }
.cap-boleto-filter select, .cap-boleto-edit input[type=text] { display:block; max-width:100%; margin-top:4px; padding:7px; border:1px solid #d8d0e1; border-radius:6px; background:#fff; color:inherit; }
.cap-boleto-table { width:100%; font-size:12px; }
.cap-boleto-table tbody tr:nth-child(even) { background:#faf9fc; }
.cap-boleto-status { display:inline-block; padding:3px 6px; border-radius:5px; white-space:nowrap; background:#f9e7ed; color:#86234a; }
.cap-boleto-status.is-paid { background:#e5f2ea; color:#275e42; }
.cap-boleto-edit { width:min(430px,calc(100% - 24px)); padding:20px; border:1px solid #e4dfeb; border-radius:12px; color:#242738; }
.cap-boleto-edit::backdrop { background:#24273899; }
.cap-boleto-edit h2 { margin-top:0; }
@media screen and (max-width:720px) {
  .cap-boleto-table, .cap-boleto-table tbody { display:block; }
  .cap-boleto-table thead { position:absolute; width:1px; height:1px; overflow:hidden; clip-path:inset(50%); }
  .cap-boleto-table tbody tr { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); margin-bottom:10px; border:1px solid #e4dfeb; border-radius:8px; }
  .cap-boleto-table td { display:block; text-align:left; overflow-wrap:anywhere; }
  .cap-boleto-table td::before { content:attr(data-label); display:block; font-weight:650; }
  .cap-boleto-table td[colspan] { grid-column:1/-1; }
}
@page { size:A4 portrait; margin:10mm; }
@media print {
  body { background:white; color:#111; font-size:9.5pt; }
  .cap-document { max-width:none; margin:0; padding:5mm 8mm; border:0; }
  .cap-doc-logo { width:32mm; background:transparent; filter:brightness(0); }
  .cap-doc-header h1 { font-size:15pt; }
  .cap-document > .cap-doc-header { display:none; }
  .cap-doc-toolbar, .cap-boleto-filter, .cap-boleto-edit, .cap-boleto-action { display:none !important; }
  .cap-doc-scroll { overflow:visible; }
  .cap-doc-scroll table { max-width:100%; }
  .cap-doc-scroll th, .cap-doc-scroll td { padding:4px 3px; }
  .cap-doc-scroll tr { break-inside:avoid; }
  .cap-doc-scroll thead { display:table-header-group; }
  .cap-doc-scroll .folha { break-after:page; }
  .cap-doc-scroll .folha:last-of-type { break-after:auto; }
}
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
<body class="cap-report-page">
<main class="cap-document" data-cap-report>
<%If Request.QueryString("modal") <> "1" Then%><div class="cap-doc-toolbar"><button class="cap-doc-button cap-doc-primary" type="button" onclick="window.print()">Imprimir / Salvar PDF</button></div><%End If%>
<header class="cap-doc-header"><h1>Recibos de pagamento</h1><img class="cap-doc-logo" src="images/logoparafundopreto.png" alt="Stillos"></header>
<div class="cap-doc-scroll">

<%


AbrirConexao()

Dim rsTotos
Dim NOME_VENDEDOR, COD_VENDEDOR
Dim TOTAL_A_RECEBER_LIQUIDO, TODOS_PEDIDOS_POR_VENDEDOR, TODOS_PEDIDOS_BOLETOS_POR_VENDEDOR

'Obs.: Nos casos dos vendedores coordenadores 7 e 14 os mesmos deverão aparecer mesmo que não tenham venda

Set rsTodos = objConexao.Execute("SELECT * FROM (SELECT "&_
	"TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR ,"&_
	"TB_VENDEDORES.NOME_VENDEDOR "&_
	"FROM TB_VENDEDORES INNER JOIN TB_FECHAMENTO_PAGAMENTO ON TB_VENDEDORES.COD_VENDEDOR = TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR "&_
	"WHERE (TB_FECHAMENTO_PAGAMENTO.COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&" AND TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR <> 27) OR TB_VENDEDORES.COD_VENDEDOR IN (SELECT cod_vendedor FROM TB_FECHAMENTO_PENDENCIAS_PAGAMENTO WHERE cod_fechamento_periodo = '"&COD_FECHAMENTO_PERIODO&"') or TB_VENDEDORES.COD_VENDEDOR IN(7,14,48) "&_
	"GROUP BY TB_FECHAMENTO_PAGAMENTO.COD_VENDEDOR UNION ALL "&_
	"SELECT TB_FECHAMENTO_PAGAMENTO_BOLETOS.COD_VENDEDOR, TB_VENDEDORES.NOME_VENDEDOR FROM TB_VENDEDORES INNER JOIN TB_FECHAMENTO_PAGAMENTO_BOLETOS ON TB_VENDEDORES.COD_VENDEDOR = TB_FECHAMENTO_PAGAMENTO_BOLETOS.COD_VENDEDOR WHERE (TB_FECHAMENTO_PAGAMENTO_BOLETOS.COD_FECHAMENTO_PERIODO = "&COD_FECHAMENTO_PERIODO&") GROUP BY TB_FECHAMENTO_PAGAMENTO_BOLETOS.COD_VENDEDOR ORDER BY NOME_VENDEDOR) AS teste GROUP BY teste.COD_VENDEDOR")

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
				'1) TOTAL GERAL POR VENDEDOR LíQUIDO - EXCLUINDO PRÉ-VENDA COM BOLETO
				Dim rsRecibo
				Set rsRecibo = objConexao.Execute("SELECT IFNULL(SUM(fp.VALOR_TOTAL_COMISSAO),0) AS TOTAL_A_RECEBER " &_
					"FROM TB_FECHAMENTO_PAGAMENTO fp " &_
					"LEFT JOIN TB_PEDIDOS p ON p.COD_PEDIDO = fp.COD_PEDIDO " &_
					"WHERE fp.COD_VENDEDOR=" & COD_VENDEDOR & " " &_
					"AND fp.COD_FECHAMENTO_PERIODO=" & COD_FECHAMENTO_PERIODO & " " &_
					"AND NOT (p.PRE_VENDA='S' AND fp.TOTAL_BOLETO > 0)")
				
				If Not rsRecibo.Eof Then
					TOTAL_A_RECEBER_LIQUIDO = rsRecibo("TOTAL_A_RECEBER")
				Else
					TOTAL_A_RECEBER_LIQUIDO = 0
				End If
				rsRecibo.Close
				Set rsRecibo = Nothing
				
				if ccur(TOTAL_A_RECEBER_LIQUIDO) < 0 then
                                   TOTAL_A_RECEBER_LIQUIDO   = 0
                            end if
'#################################################################################################################################			
%>
    <table width="98%" border="0" align="center" cellpadding="20" cellspacing="0" class="folha">
      <tr>
      <td>
        
        <div class="cap-doc-header"><strong>Recibo de pagamento</strong><img class="cap-doc-logo" src="images/logoparafundopreto.png" alt="Stillos"></div>
        <p class="style1"></p>
        <p align="justify">Recebi da empresa: Stillo's Reportagens  Fotográficas S/C Ltda., a importância de <strong><%=formatCurrency(TOTAL_A_RECEBER_LIQUIDO)%></strong>(<%=Extenso(TOTAL_A_RECEBER_LIQUIDO)%>) referente <strong>a comissão nas vendas </strong> dos pedidos:<%=TODOS_PEDIDOS_POR_VENDEDOR%> <% if TODOS_PEDIDOS_BOLETOS_POR_VENDEDOR <> "" then %> e também da compensasão dos boletos: <%= TODOS_PEDIDOS_BOLETOS_POR_VENDEDOR %>  <% end if %>

        <% If int(COD_VENDEDOR) = 14 or int(COD_VENDEDOR) = 7 or int(COD_VENDEDOR) = 48 Then %> e também referente comissão das vendas dos pedidos para os coordenadores de repasse<%End If%>.</p>
        <p></p>
        <p align="left">Sem nada a reclamar, assino.</p>
        <p align="left"></p>
       <div align="left"><strong>São  Paulo, <%=day(data_pagamento)%> de <%=MonthName(Month(data_pagamento))%> de <%=Year(data_pagamento)%>.</strong>
        <p align="left"></p>
        <p align="left">-------------------------------------------</p>
        <div align="left"><em><%=NOME_VENDEDOR%></em>
        </div>
        <p></p>
        <p></p></TD>
      </tr>
    </table>

<%
rsTodos.MoveNext()
Wend
			
rsTodos.Close
Set rsTodos = Nothing
FecharConexao()
%>


</div></main>
<%If Request.QueryString("modal") <> "1" Then%><script>window.addEventListener('load', function () { window.print(); });</script><%End If%>
</body>
</html>



