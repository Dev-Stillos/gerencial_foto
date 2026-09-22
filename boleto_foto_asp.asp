<!-- #include virtual="gerencial_foto/includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="includes/classe_modulo_cob.asp" -->
<%

dim hash_boleto : hash_boleto = request.queryString("h")
dim valor, nosso_numero, numero_documento, vencimento, parcela
dim nome_sacado, cpf_sacado
dim rs 

abrirConexao()

set rs = objconexao.Execute("SELECT * FROM TB_PEDIDOS_BOLETOS WHERE hash = "&sqlTexto(hash_boleto)&" limit 1")
if not rs.Eof then
  cod_sacado = rs("cod_sacado")
  valor = valorMultaJuros(rs("data_vencimento"), rs("data_pagamento"), rs("valor_pago"), rs("valor"))
  nosso_numero = rs("nosso_numero")
  numero_documento = rs("nosso_numero")
  vencimento = vencimentoCorrigido(rs("data_vencimento"))
  parcela = rs("parcela")
end if
rs.close
set rs = nothing

set rs = objconexao.execute("SELECT * FROM TB_SACADOS WHERE cod_sacado = "&sqlTexto(cod_sacado))
if not rs.Eof then
  nome_sacado = rs("nome")
  cpf_sacado = mascaraCPF(rs("cpf"))
end if
rs.close
set rs = nothing

fecharConexao()

'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'----------------------------------------------------------------- INICIO BOLETO ---------------------------------------------------------
'Desenvolvimento:
'Luciano Lima
'www.netdinamica.com.br/boleto
'Suporte: boleto@netdinamica.com.br

'Licença de uso:
'A licença deste script é para uso próprio em seu site.
'Não é permitida a revenda do mesmo, estando sujeito às penas prevista em lei.
'A netdinamica, não se responsabiliza por quaisquer prejuízos causados pelo uso deste script.
'Bem como imperícia ou alterações indevidas no código fonte do mesmo.
'O mesmo segue os padrões definidos por cada banco, cabendo ao comprador apenas a configuração e instalação do script em seu site, bem como fazer todos os testes, para verificação do correto funcionamento do sistema.
'Caso necessite da licença para revenda entre em contato.

'Instruções de Instalação do Script
'Leia o arquivo Boleto Bancário em PHP ou ASP que acompanha este script, antes de iniciar a instalação.
'=============================================================================================
'===================Dados Principias para gera o Boleto====================

x101    = arrumaNumero(valor,"0,00",2)    'Valor do Boleto
x81     = nosso_numero           'Nosso Numero 
x89     = numero_documento       'Numero do Documento
x33     = arrumaData(vencimento) 'Data Vencimento 
x97778  = arrumaData(mid(now(),1 , 10)) 'Data PRocessamento


'====Dados do Cliente  - Quem paga o Boleto====
x109 = Ucase(nome_sacado) ' Nome 
x70 = "CPF "&cpf_sacado ' CPF
x79 = " " ' Contrato / enderecod
x74 = " " ' Curso  / continuacao edereco

'==Dados do Cedente - Quem emiti o boleto (Titular da Conta no banco)
'==Este dados você pode deixar fixo, já que eles não mudam
x136 = "Stillo's Reportagens Fotograficas Ltda." ' Nome do Cedente (Emissor do Boleto)
x56  = "Rua General Calado, 52 - Tatuapé"
x52  = "São Paulo - SP - CEP 03334-060"
x49  = "67.636.886/0001-35" ' CPF ou CNPJ do Cedente
x10  = "16" ' Numero da Carteira
x83  = "0132" ' Numero da Agência sem o Digito
x121 = "5" ' Digito da Agência
x64  = "54627" ' Numero da Conta sem o digito
x72  = "5"  ' Digito da Conta

'===Instruções ao Caixa do Banco (Pox exemplo "Não Receber após o vencimento" ou "Cobrar multa de 2 % após o vencimento")
x80  = "NÃO CONSIDERAR CAMPOS DE DESCONTO" '  Campo opcional - Instrucoes 1
x171 = "NÃO DEFLACIONAR"                   '  Campo opcional - Instrucoes 2
x130 = "APÓS O VENCIMENTO COBRAR:"         '  Campo opcional - Instrucoes 3
x97  = "MULTA DE 2%, MAIS 1% AO MES"       '  Campo opcional - Instrucoes 4
x73  = ""                                  '  Campo opcional - Instrucoes 5

'=== Valores Padroes =======================================================
x163 = "" ' Uso do Banco
x98  = "" ' Quantidade
x34  = "R$" ' Especie
x174 = "NA" ' Aceite

'==Os valores Abaixo São Fixos para o Bradesco (Não mexa)
x128 = "237"  ' Banco
x19  = "9" ' Moeda 
x160 = "R$"
'===================================================================================================


'===================================================================================================


Function x32(x126, x13)
Dim x51, x27, x122, x150, x151, cnum
x126 = RTrim(LTrim(x126))
For x150 = 1 To Len(x126)
x151 = Mid(x126, x150, 1)
If IsNumeric(x151) Then
cnum = cnum & x151
End If
Next
x126 = cnum
x122 = "0000000000000000000000000000": x51 = CInt(x51)
If Len(x126) < x13 Then 
x51 = Abs(x13) - Abs(Len(x126)): x27 = Mid(x122, 1, x51) & CStr(x126)
x32 = x27
ElseIf Len(x126) > x13 Then
x32 = Right(x126, x13)
Else
x32 = x126
End If
End Function

Function x16(x15, x36, x94)
Dim x87, x69, x45
x45 = x36 + 1
x87 = ""
Do
If IsNumeric(Mid(x15, (x45), 1)) Then
x87 = x87 & Mid(x15, (x45), 1)
x45 = x45 + 1
End If
Loop While IsNumeric(Mid(x15, (x45), 1))
For x69 = 1 To x87
x94 = x94 & (x91(53))
Next
x36 = x36 + 2
x16 = x94
End Function

Function x14(x15)
Dim x36, x78, x60, x94
x94 = ""
For x36 = 1 To Len(x15)
x78 = 0
If Mid(x15, (x36), 1) = "¹" Then
x94 = x16(x15, x36, x94)
x78 = 0
End If
x78 = InStr(1, x135, Mid(x15, (x36), 1), vbBinaryCompare)
If x78 > 0 Then
   If x78 <= 113 Then
  x94 = x94 & (x91(x78 + 1))
   Else
  x94 = x94 & (x91(1))
  End If
End If
Next
x14 = x94 : response.write x14
End Function

Function x23(x15)
For x18 = 1 To Len(x15) Step 3
x78 = 0 : x78 = Mid(x15, (x18), 3) : x94 = x94 & x91(x78)
Next
x23 = x94 
End Function

Function x31(x36)
Dim x59, x30, x173, x68, x82, x120, x113
x68 = 2
For x59 = 1 To 43
x113 = Mid(Right(x36, x59), 1, 1)
If x68 > 9 Then
x68 = 2: x30 = 0
End If
x30 = x113 * x68: x173 = x173 + x30: x68 = x68 + 1
Next
x82 = x173 Mod 11
x120 = 11 - x82
If x120 = 0 Or x120 = 1 Or x120 > 9 Then
x31 = 1
Else
x31 = x120
End If
End Function

Function x102(x124)
Dim x59, x30, x173, x68, x82, x25
If Not IsNumeric(x124) Then
x102 = ""
Exit Function
End If
x68 = 2
For x59 = Len(x124) To 1 Step -1
x30 = CInt(Mid(x124, x59, 1)) * x68
If x30 > 9 Then
x30 = CInt(Left(x30, 1) + CInt(Right(x30, 1)))
End If
x173 = x173 + x30
If x68 = 2 Then
x68 = 1
Else
x68 = 2
End If
Next
x25 = (x173 / 10) * -1: x25 = Int([x25]) * -1: x25 = x25 * 10
x82 = x25 - x173: x102 = x82
End Function

Function x134(x124)
Dim x59, x30, x173, x68, x82, x170
If Not IsNumeric(x124) Then
x134 = ""
Exit Function
End If
x68 = 2
For x59 = Len(x124) To 1 Step -1
x30 = Mid(x124, x59, 1) * x68: x173 = x173 + x30
If x68 > 7 Then
x68 = 2
Else
x68 = x68 + 1
End If
If x68 = 8 Then x68 = 2
Next
x82 = x173 Mod 11
x134 = 11 - x82
If x82 = 1 Then
x134 = "P"
ElseIf x82 = 0 Then
x134 = "0"
End If
End Function

Function x77(x36, x66, x92, x142)
Dim x112, x55, x42, x65, x105, x87, x69, x45
x112 = Left(x36, 9): x55 = Mid(x36, 10, 10): x42 = Right(x36, 10)
x105 = CCur(x142): x65 = x32(x105, 10)
x87 = x102(x112): x69 = x102(x55): x45 = x102(x42)
x112 = Left(x112 & x87, 5) & "." & Right(x112 & x87, 5)
x55 = Left(x55 & x69, 5) & "." & Right(x55 & x69, 6)
x42 = Left(x42 & x45, 5) & "." & Right(x42 & x45, 6)
x77 = x112 & " &nbsp;" & x55 & " &nbsp;" & x42 & " &nbsp;" & x66 & " &nbsp;" & x92 & x65
End Function


Function x17(x108, x140, x106, x142, x58)
Dim x116, x146, x75, x46
If x106 <> "" Then
x75 = CDate("7/10/1997")
x146 = DateDiff("d", x75, CDate(x106))
Else
x146 = "0000"
End If
x142 = Int(x142 * 100)
x116 = x108 & x140 & x146 & x32(x142, 10) & x58: x46 = x31(x116)
x17 = (Left(x116, 4) & x46 & Right(x116, 39))
End Function

Function x147(x108, x140, x106, x142, x58)
Dim x116, x146, x75, x46
If x106 <> "" Then
x75 = CDate("7/10/1997")
x146 = DateDiff("d", x75, CDate(x106))
Else
x146 = "0000"
End If
x142 = Int(x142 * 100): x116 = x108 & x140 & x146 & x32(x142, 10) & x58
x46 = x31(x116): x147 = x77(x108 & x140 & x58, CStr(x46), x146, x32(x142, 10))
End Function

Function x138(x143, x36)
Dim x170 
x170 = x134(x32(x143, 2) & x32(x36, 11))
x138 = x32(x143, 2) & x32(x36, 11)
End Function

Function x148(x143, x36)
Dim x170 
x170 = x134(x32(x143, 2) & x32(x36, 11))
x148 = x32(x143, 2) & "/" & x32(x36, 11) & "-" & x170
End Function

Function x91(x78)
x91 = Mid(x135, x78, 1)
End Function

Function x4(x15)
For x18 = 1 To Len(x15) Step 3
x78 = 0 : x78 = Mid(x15, (x18), 3) : x94 = x94 & x91(x78)
Next
x4 = x94 : response.write x4
End Function
x11= x136
x81 = x32(x81, 11) 
x101 = FormatNumber((x101), 2,-2,-2,-2)
x83 = x32(x83, 4)
x121 = x32(x121, 1)
x103 = x83 & "-" & x121 
x64 = x32(x64, 7)
x72 = x32(x72, 1)
x100 = x64 & "-" & x72 
x127 = x103 & "/" & x100 : x119 = x138(x10, CStr(x81)) 
x5 = x83 & x119 & x64 & "0" : x123 = x17(x128, x19, CDate(x33), CCur(x101), x5)
x57 = x147(x128, x19, CDate(x33), CCur(x101), x5)
x104 = x148(x10, CStr(x81))

Dim x104_Esq, x104_Dir 
'MASCARA NO NOSSO NÚMERO

x104_Esq = Left(x104,5)
x104_Dir = Mid(x104,6,60)

x104 = x104_Esq&"/"&x104_Dir
%>
<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>
<HTML>
  <HEAD>
    <TITLE>Stillo's Produ&ccedil;&otilde;es e Eventos - Impress&atilde;o de Boleto Online</TITLE>
    <META http-equiv=Content-Type content=text/html; charset=windows-1252>
    <meta name=GENERATOR content=NetDinamica>
    <style type=text/css>
      <!--.cp {  font: bold 10px Arial; color: black}
        <!--.ti {  font: 9px Arial, Helvetica, sans-serif}
        <!--.ld { font: bold 13px Arial; color: #000000}
        <!--.ct { FONT: 9px "Arial Narrow"; COLOR: #000000}
        <!--.cn { FONT: 9px Arial; COLOR: black }
        <!--.bc { font: bold 22px Arial; text-align:center; color: #000000 }
        -->
    </style>
  </HEAD>
  <BODY text=#000000 bgColor=#ffffff topMargin=0 rightMargin=0>
    <table width=666 cellspacing=5 cellpadding=0 border=0 align=Default>
      <tr>
        <td width=41><img src="images/logo-stillos.jpg" width="110" height="86"></td>
        <td class=ti width=455 valign="middle">
          <%response.write x136 & "<br>" : response.write x56 & "<br>" : response.write x52 & "<br>": response.write x71 & "<br>" : response.write x168 %> 
        </td>
        <td align=RIGHT width=150 class=ti>&nbsp;</td>
      </tr>
    </table>
    <BR>
    <table cellspacing=0 cellpadding=0 width=666 border=0>
      <tr>
        <td class=cp>
          <table width="665" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td width="143"><span class="cp"><img src=images/logo-bradesco.jpg border=0 WIDTH="135" HEIGHT="36"></span></td>
              <td width="86" valign="bottom">
                <table width="80" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td valign="bottom"><img height=22 src=images/3.gif width=2 border=0></td>
                    <td valign="bottom"><div align=center><font class=bc>&nbsp;237-2</font></div></td>
                    <td valign="bottom"><img height=22 src=images/3.gif width=2 border=0></td>
                  </tr>
                </table>
              </td>
              <td width="436" valign="bottom">
                <div align="right"><span class=ld><%=x57%>&nbsp;</span></div>
              </td>
            </tr>
          </table>
        </td>
      </tr>
      <tbody>
        <tr>
          <td><img height=2 src=images/2.gif width=666 border=0></td>
        </tr>
      </tbody>
    </table>
    <table cellspacing=0 cellpadding=0 border=0>
      <tbody>
        <tr>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=298 height=13>Cedente</td>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=126 height=13>Agência/Código do Cedente</td>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=34 height=13>Espécie</td>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=53 height=13>Quantidade</td>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=120 height=13>Nosso número</td>
        </tr>
        <tr>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top width=298 height=12><%=x136%></td>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top width=126 height=12><%=x127%></td>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top  width=34 height=12><%=x160%></td>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top  width=53 height=12><%=x98%></td>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top align=right width=120 height=12><%=x104%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
        </tr>
        <tr>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=298 height=1><img height=1 src=images/2.gif width=298 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=126 height=1><img height=1 src=images/2.gif width=126 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=34 height=1><img height=1 src=images/2.gif width=34 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=53 height=1><img height=1 src=images/2.gif width=53 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=120 height=1><img height=1 src=images/2.gif width=120 border=0></td>
        </tr>
      </tbody>
    </table>
    <table cellspacing=0 cellpadding=0 border=0>
      <tbody>
        <tr>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top colspan=3 height=13>Número do documento</td>
          <td height=13 colspan="2" valign=top class=ct>&nbsp;</td>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=134 height=13>Vencimento</td>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=180 height=13>Valor documento</td>
        </tr>
        <tr>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top colspan=3 height=12><%=x89%></td>
          <td height=12 colspan="2" valign=top class=cp>&nbsp;</td>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top width=134 height=12><%=x33%></td>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top align=right width=180 height=12><%=x101%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
        </tr>
        <tr>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=113 height=1><img height=1 src=images/2.gif width=113 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=72 height=1><img height=1 src=images/2.gif width=72 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=132 height=1><img height=1 src=images/2.gif width=132 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=134 height=1><img height=1 src=images/2.gif width=134 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=180 height=1><img height=1 src=images/2.gif width=180 border=0></td>
        </tr>
      </tbody>
    </table>
    <table cellspacing=0 cellpadding=0 border=0>
      <tbody>
        <tr>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=113 height=13>(-) Desconto / Abatimentos</td>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=112 height=13>(-) Outras deduções</td>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=113 height=13>(+) Mora / Multa</td>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=113 height=13>(+) Outros acréscimos</td>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=180 height=13>(=) Valor cobrado</td>
        </tr>
        <tr>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top align=right width=113 height=12></td>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top align=right width=112 height=12></td>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top align=right width=113 height=12></td>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top align=right width=113 height=12></td>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top align=right width=180 height=12></td>
        </tr>
        <tr>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=113 height=1><img height=1 src=images/2.gif width=113 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=112 height=1><img height=1 src=images/2.gif width=112 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=113 height=1><img height=1 src=images/2.gif width=113 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=113 height=1><img height=1 src=images/2.gif width=113 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=180 height=1><img height=1 src=images/2.gif width=180 border=0></td>
        </tr>
      </tbody>
    </table>
    <table cellspacing=0 cellpadding=0 border=0>
      <tbody>
        <tr>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=659 height=13>Sacado</td>
        </tr>
        <tr>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top width=659 height=12><%=x109 %></td>
        </tr>
        <tr>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=659 height=1><img height=1 src=images/2.gif width=659 border=0></td>
        </tr>
      </tbody>
    </table>
    <table cellspacing=0 cellpadding=0 border=0>
      <tbody>
        <tr>
          <td class=ct width=7 height=12></td>
          <td class=ct width=511 >&nbsp;</td>
          <td class=ct width=10 height=12></td>
          <td class=ct width=138 >Autenticação mecânica</td>
        </tr>
        <tr>
          <td width=7 ></td>
          <td width=511 ></td>
          <td width=10 ></td>
          <td width=138 ></td>
        </tr>
      </tbody>
    </table>
    <table cellspacing=0 cellpadding=0 width=666 border=0>
      <tbody>
        <tr>
          <td width=7></td>
          <td width=500 class=cp>&nbsp;</td>
          <td width=159></td>
        </tr>
      </tbody>
    </table>
    <table cellspacing=0 cellpadding=0 width=666 border=0>
      <tr>
        <td class=ct width=666></td>
      </tr>
      <tbody>
        <tr>
          <td class=ct width=666>
            <div align=right>Corte na linha pontilhada
                <span class="cp">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
                <span class="cp">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
            </div>
          </td>
        </tr>
        <tr>
          <td class=ct width=666><img height=1 src=images/6.gif width=665 border=0></td>
        </tr>
      </tbody>
    </table>
    <br><br>
    <table cellspacing=0 cellpadding=0 width=666 border=0>
      <tr>
        <td class=cp>
          <span class=ld>
            <table width="665" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="143"><span class="cp"><img src=images/logo-bradesco.jpg border=0 WIDTH="135" HEIGHT="36"></span></td>
                <td width="86" valign="bottom">
                  <table width="80" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td valign="bottom"><img height=22 src=images/3.gif width=2 border=0></td>
                      <td valign="bottom"><div align=center><font class=bc>&nbsp;237-2</font></div></td>
                      <td valign="bottom"><img height=22 src=images/3.gif width=2 border=0></td>
                    </tr>
                  </table>
                </td>
                <td width="436" valign="bottom">
                  <div align="right"><span class=ld><%=x57%>&nbsp;</span></div>
                </td>
              </tr>
            </table>
          </span>
        </td>
      </tr>
      <tbody>
        <tr>
          <td><img height=2 src=images/2.gif width=666 border=0></td>
        </tr>
      </tbody>
    </table>
    <table cellspacing=0 cellpadding=0 border=0>
      <tbody>
        <tr>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=472 height=13>Local de pagamento</td>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=180 height=13>Vencimento</td>
        </tr>
        <tr>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top width=472 height=12>
            <%
              If INT(COD_INSTRUCAO) <> 99 Then
              %>
            Pagável em qualquer banco até o vencimento. Ap&oacute;s venc. somente em ag&ecirc;ncias Bradesco.
            <%Else%>
            Pagável em qualquer banco até o vencimento. NÃO ACEITAR APÓS VENCIMENTO.
            <%End If%>
          </td>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top align=right width=180 height=12><%=x33%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
        </tr>
        <tr>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=472 height=1><img height=1 src=images/2.gif width=472 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=180 height=1><img height=1 src=images/2.gif width=180 border=0></td>
        </tr>
      </tbody>
    </table>
    <table cellspacing=0 cellpadding=0 border=0>
      <tbody>
        <tr>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=472 height=13>Cedente</td>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=180 height=13>Agência/Código cedente</td>
        </tr>
        <tr>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top width=472 height=12><%=x136%></td>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top align=right width=180 height=12>0<%=x127%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
        </tr>
        <tr>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=472 height=1><img height=1 src=images/2.gif width=472 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=180 height=1><img height=1 src=images/2.gif width=180 border=0></td>
        </tr>
      </tbody>
    </table>
    <table cellspacing=0 cellpadding=0 border=0>
      <tbody>
        <tr>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=113 height=13>Data do documento</td>
          <td class=ct valign=top width=7 height=13> <img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=143 height=13>Número do documento</td>
          <td class=ct valign=top width=7 height=13> <img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=62 height=13>Espécie doc.</td>
          <td class=ct valign=top width=7 height=13> <img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=34 height=13>Aceite</td>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=9 height=13>Data processamento</td>
          <td class=ct valign=top width=7 height=13> <img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=180 height=13>Nosso número</td>
        </tr>
        <tr>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top  width=113 height=12><div align=left><%=formataData(x97778)%></div></td>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top width=143 height=12><%=x89%></td>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top  width=62 height=12><div align=left>OUTROS</div></td>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top  width=34 height=12><div align=left><%=x174%></div></td>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top  width=92 height=12><div align=left><%=formataData(x97778)%></div></td>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top width=180 height=12 align=right><%=x104%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
        </tr>
        <tr>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=113 height=1><img height=1 src=images/2.gif width=113 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=143 height=1><img height=1 src=images/2.gif width=143 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=62 height=1><img height=1 src=images/2.gif width=62 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=34 height=1><img height=1 src=images/2.gif width=34 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=92 height=1><img height=1 src=images/2.gif width=92 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=180 height=1><img height=1 src=images/2.gif width=180 border=0></td>
        </tr>
      </tbody>
    </table>
    <table cellspacing=0 cellpadding=0 border=0>
      <tbody>
        <tr>
          <td class=ct valign=top width=7 height=13> <img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top COLSPAN="3" height=13>Uso do banco</td>
          <td class=ct valign=top height=13 width=7> <img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=83 height=13>Carteira</td>
          <td class=ct valign=top height=13 width=7><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=53 height=13>Espécie</td>
          <td class=ct valign=top height=13 width=7><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=123 height=13>Quantidade</td>
          <td class=ct valign=top height=13 width=7><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=72 height=13>Valor</td>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=180 height=13>(=) Valor documento</td>
        </tr>
        <tr>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top height=12 COLSPAN="3"><div align=left><%=x163%></div></td>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top width=83><div align=left><%=x10%></div></td>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top width=53><div align=left><%=x160%></div></td>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top width=123><%=x98%></td>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top width=72><%=x101%></td>
          <td class=cp valign=top width=7 height=12> <img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top align=right width=180 height=12><%=x101%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
        </tr>
        <tr>
          <td valign=top width=7 height=1> <img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=75 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=31 height=1><img height=1 src=images/2.gif width=31 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=83 height=1><img height=1 src=images/2.gif width=83 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=53 height=1><img height=1 src=images/2.gif width=53 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=123 height=1><img height=1 src=images/2.gif width=123 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=72 height=1><img height=1 src=images/2.gif width=72 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=180 height=1><img height=1 src=images/2.gif width=180 border=0></td>
        </tr>
      </tbody>
    </table>
    <table cellspacing=0 cellpadding=0 width=666 border=0>
      <tbody>
        <tr>
          <td align=right width=7>
            <table cellspacing=0 cellpadding=0 border=0 align=left>
              <tbody>
                <tr>
                  <td class=ct valign=top width=10 height=13><img height=13 src=images/1.gif width=1 border=0></td>
                </tr>
                <tr>
                  <td class=cp valign=top width=10 height=12><img height=12 src=images/1.gif width=1 border=0></td>
                </tr>
                <tr>
                  <td valign=top width=10 height=1><img height=1 src=images/2.gif width=1 border=0></td>
                </tr>
              </tbody>
            </table>
          </td>
          <td valign=top width=468 rowspan=5><font class=ct>Instruções de responsabilidade do cedente.</font><br><span class=cp>
            <%
              'If INSTRUCOES_PARA_ATRASADOS = "diferente" Then
              'Diferente, pois o boleto está atrasado e segue um padrão de instruções para boletos atrasados
              ''  Response.Write( "- SAC: 2672-9899 ou www.stillos.com.br<br />"&_
              ''          " - NÃO CONSIDERAR CAMPOS DE DESCONTOS.<br />"&_
              ''          " - QUITAÇÃO SOMENTE APÓS A COMPENSAÇÃO DO CHEQUE.<br />"&_
              ''          " - NÃO RECEBER APÓS A DATA DE VENCIMENTO.")

             '' Else
              
                response.write x80 & "<br>" : response.write x171 & "<br>"
                response.write x130 & "<br>" : response.write x97 & "<br>" : response.write x73
                
              'End If
              %>
            </span>
          </td>
          <td align=right width=188>
            <table cellspacing=0 cellpadding=0 border=0>
              <tbody>
                <tr>
                  <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
                  <td class=ct valign=top width=180 height=13>(-) Desconto / Abatimentos</td>
                </tr>
                <tr>
                  <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
                  <td class=cp valign=top align=right width=180 height=12></td>
                </tr>
                <tr>
                  <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
                  <td valign=top width=180 height=1><img height=1 src=images/2.gif width=180 border=0></td>
                </tr>
              </tbody>
            </table>
          </td>
        </tr>
        <tr>
          <td align=right width=7>
            <table cellspacing=0 cellpadding=0 border=0 align=left>
              <tbody>
                <tr>
                  <td class=ct valign=top width=10 height=13><img height=13 src=images/1.gif width=1 border=0></td>
                </tr>
                <tr>
                  <td class=cp valign=top width=10 height=12><img height=12 src=images/1.gif width=1 border=0></td>
                </tr>
                <tr>
                  <td valign=top width=10 height=1><img height=1 src=images/2.gif width=1 border=0></td>
                </tr>
              </tbody>
            </table>
          </td>
          <td align=right width=188>
            <table cellspacing=0 cellpadding=0 border=0>
              <tbody>
                <tr>
                  <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
                  <td class=ct valign=top width=180 height=13>(-) Outras deduções</td>
                </tr>
                <tr>
                  <td class=cp valign=top width=7 height=12> <img height=12 src=images/1.gif width=1 border=0></td>
                  <td class=cp valign=top align=right width=180 height=12></td>
                </tr>
                <tr>
                  <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
                  <td valign=top width=180 height=1><img height=1 src=images/2.gif width=180 border=0></td>
                </tr>
              </tbody>
            </table>
          </td>
        </tr>
        <tr>
          <td align=right width=7>
            <table cellspacing=0 cellpadding=0 border=0 align=left>
              <tbody>
                <tr>
                  <td class=ct valign=top width=10 height=13><img height=13 src=images/1.gif width=1 border=0></td>
                </tr>
                <tr>
                  <td class=cp valign=top width=10 height=12><img height=12 src=images/1.gif width=1 border=0></td>
                </tr>
                <tr>
                  <td valign=top width=10 height=1><img height=1 src=images/2.gif width=1 border=0></td>
                </tr>
              </tbody>
            </table>
          </td>
          <td align=right width=188>
            <table cellspacing=0 cellpadding=0 border=0>
              <tbody>
                <tr>
                  <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
                  <td class=ct valign=top width=180 height=13>(+) Mora / Multa</td>
                </tr>
                <tr>
                  <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
                  <td class=cp valign=top align=right width=180 height=12></td>
                </tr>
                <tr>
                  <td valign=top width=7 height=1> <img height=1 src=images/2.gif width=7 border=0></td>
                  <td valign=top width=180 height=1><img height=1 src=images/2.gif width=180 border=0></td>
                </tr>
              </tbody>
            </table>
          </td>
        </tr>
        <tr>
          <td align=right width=7>
            <table cellspacing=0 cellpadding=0 border=0 align=left>
              <tbody>
                <tr>
                  <td class=ct valign=top width=10 height=13><img height=13 src=images/1.gif width=1 border=0></td>
                </tr>
                <tr>
                  <td class=cp valign=top width=10 height=12><img height=12 src=images/1.gif width=1 border=0></td>
                </tr>
                <tr>
                  <td valign=top width=10 height=1><img height=1 src=images/2.gif width=1 border=0></td>
                </tr>
              </tbody>
            </table>
          </td>
          <td align=right width=188>
            <table cellspacing=0 cellpadding=0 border=0>
              <tbody>
                <tr>
                  <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
                  <td class=ct valign=top width=180 height=13>(+) Outros acréscimos</td>
                </tr>
                <tr>
                  <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
                  <td class=cp valign=top align=right width=180 height=12></td>
                </tr>
                <tr>
                  <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
                  <td valign=top width=180 height=1><img height=1 src=images/2.gif width=180 border=0></td>
                </tr>
              </tbody>
            </table>
          </td>
        </tr>
        <tr>
          <td align=right width=7>
            <table cellspacing=0 cellpadding=0 border=0 align=left>
              <tbody>
                <tr>
                  <td class=ct valign=top width=10 height=13><img height=13 src=images/1.gif width=1 border=0></td>
                </tr>
                <tr>
                  <td class=cp valign=top width=10 height=12><img height=12 src=images/1.gif width=1 border=0></td>
                </tr>
              </tbody>
            </table>
          </td>
          <td align=right width=188>
            <table cellspacing=0 cellpadding=0 border=0>
              <tbody>
                <tr>
                  <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
                  <td class=ct valign=top width=180 height=13>(=) Valor cobrado</td>
                </tr>
                <tr>
                  <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
                  <td class=cp valign=top align=right width=180 height=12></td>
                </tr>
              </tbody>
            </table>
          </td>
        </tr>
      </tbody>
    </table>
    <table cellspacing=0 cellpadding=0 width=666 border=0>
      <tbody>
        <tr>
          <td valign=top width=666 height=1><img height=1 src=images/2.gif width=666 border=0></td>
        </tr>
      </tbody>
    </table>
    <table cellspacing=0 cellpadding=0 border=0>
      <tbody>
        <tr>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=659 height=13>Sacado</td>
        </tr>
        <tr>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top width=659 height=12><%=x109%> - <%=x70%></td>
        </tr>
      </tbody>
    </table>
    <table cellspacing=0 cellpadding=0 border=0>
      <tbody>
        <tr>
          <td class=cp valign=top width=7 height=12><img height=12 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top width=659 height=12><%=x79&" - "&x74%></td>
        </tr>
      </tbody>
    </table>
    <table cellspacing=0 cellpadding=0 border=0>
      <tbody>
        <tr>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=cp valign=top width=472 height=13></td>
          <td class=ct valign=top width=7 height=13><img height=13 src=images/1.gif width=1 border=0></td>
          <td class=ct valign=top width=180 height=13>Cód. baixa</td>
        </tr>
        <tr>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=472 height=1><img height=1 src=images/2.gif width=472 border=0></td>
          <td valign=top width=7 height=1><img height=1 src=images/2.gif width=7 border=0></td>
          <td valign=top width=180 height=1><img height=1 src=images/2.gif width=180 border=0></td>
        </tr>
      </tbody>
    </table>
    <TABLE cellSpacing=0 cellPadding=0 border=0 width=666>
      <TBODY>
        <TR>
          <TD class=ct  width=7 height=12></TD>
          <TD class=ct  width=409 >Sacador/Avalista</TD>
          <TD class=ct  width=250 >
            <div align=right>
                Autenticação mecânica - 
                <b class=cp>Ficha de Compensação&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b>
                <span class="cp">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
            </div>
          </TD>
        </TR>
        <TR>
          <TD class=ct  colspan=3 ></TD>
        </tr>
      </tbody>
    </table>
    <TABLE cellSpacing=0 cellPadding=0 width=666 border=0>
    <TBODY>
      <TR>
        <TD vAlign=bottom align=left height=50> 
<%x129( x123 )
Sub x129( x22 )
Dim x2(99)
Const x85 = 1 : Const x131 = 3 : Const x44 = 50
if isempty(x2(0)) then
x2(0) = "00110" : x2(1) = "10001" : x2(2) = "01001" : x2(3) = "11000"
x2(4) = "00101" : x2(5) = "10100" : x2(6) = "01100" : x2(7) = "00011"
x2(8) = "10010" : x2(9) = "01010" 
for x99 = 9 to 0 step -1
for x3 = 9 to 0 Step -1
x125 = x99 * 10 + x3 : x126 = ""
for x18 = 1 To 5
x126 = x126 & mid(x2(x99), x18, 1) + mid(x2(x3), x18, 1)
next
x2(x125) = x126
next
next
end if
%> <img src=images/p.gif width=<%=x85%> height=<%=x44%> border=0><img 
src=images/b.gif width=<%=x85%> height=<%=x44%> border=0><img 
src=images/p.gif width=<%=x85%> height=<%=x44%> border=0><img 
src=images/b.gif width=<%=x85%> height=<%=x44%> border=0><img 
<%
x126 = x22
if len(x126) mod 2 <> 0 then
x126 = "0" & x126
end if
do while len(x126) > 0
x18 = cint(left( x126, 2)) : x126 = right(x126, len(x126) - 2) : x125 = x2(x18)
for x18 = 1 to 10 step 2
if mid(x125, x18, 1) = "0" then
x99 = x85
else
x99 = x131
end if
%>
src=images/p.gif width=<%=x99%> height=<%=x44%> border=0><img 
<%if mid(x125, x18 + 1, 1) = "0" Then
x3 = x85
else
x3 = x131
end if%>
src=images/b.gif width=<%=x3%> height=<%=x44%> border=0><img 
<%next
loop%>
src=images/p.gif width=<%=x131%> height=<%=x44%> border=0><img src=images/b.gif width=<%=x85%> height=<%=x44%> border=0><img src=images/p.gif width=<%=1%> height=<%=x44%> border=0> 
<% end sub %> </TD></tr></tbody></table><TABLE cellSpacing=0 cellPadding=0 width=666 border=0><TR><TD class=ct width=666></TD></TR><TBODY><TR><TD class=ct width=666><div align=right>Corte 
na linha pontilhada<span class="cp">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span><span class="cp">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></div></TD></TR><TR><TD class=ct width=666><img height=1 src=images/6.gif width=665 border=0></TD></tr></tbody></table>
</BODY></HTML>