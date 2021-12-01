Attribute VB_Name = "LVBC"
Option Compare Database
Option Explicit

' ==========================================
' OBJETO - LVBC - TRATAMENTO EM C DE STRINGS
' ==========================================
' DATA - HISTÓRICO - TÉCNICO
' 23 MAI 2002 - IMPLEMENTAÇÃO CÓDIGO BASE

Function LItem(Texto As Variant, ITEM As Variant, Optional Limit As String) As Variant
On Error GoTo LItemErro
Dim CC As Integer, Var As String, LItemIni As Integer, LItemPos As Integer
Texto = Nz(Texto, "")
CC = 0
If Limit = "" Then
    If InStr(Texto, ";") Then
        Limit = ";"
    Else
        Limit = "."
    End If
End If
LItemIni = 1
Do While LItemIni <= Len(Texto)
    CC = CC + 1
    LItemPos = InStr(LItemIni, Texto, Limit)
    If LItemPos = 0 Then
        Var = Mid(Texto, LItemIni)
    Else
        Var = Mid(Texto, LItemIni, LItemPos - LItemIni)
    End If
    If VarType(ITEM) <> 8 Then
        If CC = ITEM Then
            LItem = Var
            Exit Function
        End If
    Else
        If Var = ITEM Then
            LItem = CC
            Exit Function
        End If
    End If
    If LItemPos = 0 Then
        LItemPos = Len(Texto)
    End If
    LItemIni = LItemPos + Len(Limit)
Loop
If VarType(ITEM) <> 8 Then
    If ITEM <> 0 Then
        LItem = ""
    Else
        LItem = CC
    End If
Else
    LItem = 0
End If

LItemSai:
Exit Function

LItemErro:
Dim xerr As Integer
xerr = LErro("LItem")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LItemSai
End Function



' CMNGSTRING - DEFINIÇÃO DE ITEM DENTRO DE UMA STRING
Function LItemDef(Texto As Variant, ITEM As Single, DEF As String, Optional Limit As String) As String
On Error GoTo LItemDefErro
Dim TOT As Integer, LItemIni As Integer, LItemPos As Integer, RESULT As String
Dim Ant As Integer, údo As String, INI As Integer, fim As Integer
Ant = Int(ITEM)
Texto = Nz(Texto, "")
TOT = 0
If Limit = "" Then
    If InStr(Texto, ";") Then
        Limit = ";"
    Else
        Limit = "."
    End If
End If
LItemIni = 1

Do While LItemIni <= Len(Texto)
    TOT = TOT + 1
    If TOT = Ant Then
        INI = LItemIni
    End If
    LItemPos = InStr(LItemIni, Texto, Limit)
    If LItemPos = 0 Then
        LItemPos = Len(Texto) + Len(Limit)
    End If
    If TOT = Ant Then
        fim = LItemPos - 1
    End If
    LItemIni = LItemPos + Len(Limit)
Loop

If ITEM < 0 Then
    ITEM = TOT + 1
End If
If ITEM > TOT Then
    RESULT = Texto & IIf(DEF <> "", IIf(TOT <> 0, Limit, "") & DEF, "")
ElseIf ITEM < 1 Then
    RESULT = IIf(DEF <> "", DEF & Limit, "") & Texto
Else
    LItemIni = INI
    LItemPos = fim
    If LItemPos = 0 Then
        LItemPos = Len(Texto) + Len(Limit)
    End If
    If Ant = ITEM Then
        LItemIni = LItemIni - 1
        LItemPos = LItemPos + 1
        ' Result = Left(Texto, IIf(LItemIni < 0, 0, LItemIni)) & IIf(LItemIni > 1 And Def <> "", Limit, "") & Def & IIf((LItemPos < Len(Texto) And Def <> "") Or (Def = "" And LItemIni > 1 And LItemPos < Len(Texto)), Limit, "") & Mid(Texto, LItemPos)
        RESULT = Left(Texto, IIf(LItemIni < 0, 0, LItemIni)) & DEF & Mid(Texto, LItemPos)
    Else
        RESULT = Left(Texto, LItemPos) & Limit & DEF & Mid(Texto, LItemPos + 1)
    End If
End If
LItemDef = RESULT

LItemDefSai:
Exit Function
LItemDefErro:
Dim xerr As Integer
xerr = LErro("LItemDef")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LItemDefSai
End Function



' CMNGSTRING - OBTEM UMA LINHA DE UMA STRING
Function LExtrai(TextoRec As Variant, Var As Variant, Optional SeparaLinha As String = "^m^j", Optional SeparaDef = ":") As Variant
On Error GoTo LExtraiErro
Dim Comando As String, INI As Integer, Última As Integer, Seq As Integer, pos As Integer
Dim LINHA As String, Conteudo As String, pos1 As Integer, Texto As String
INI = 1
Última = False

If SeparaLinha = "^m^j" Then
    SeparaLinha = Chr(13) & Chr(10)
End If

Texto = ""
On Error Resume Next
Texto = Nz(TextoRec, "")
On Error GoTo LExtraiErro
Seq = 0
While Última = False
    pos = InStr(INI, Texto, SeparaLinha)
    If pos = 0 Then
        Última = True
        LINHA = Mid$(Texto, INI)
    Else
        LINHA = Mid$(Texto, INI, pos - INI)
    End If
    Seq = Seq + 1
    If VarType(Var) <> 8 Then
        If Seq = Var Then
            LExtrai = LINHA
            Exit Function
        End If
    Else
        pos1 = InStr(LINHA, SeparaDef)
        If pos1 = 0 Then
            Comando = ""
            Conteudo = LINHA
        Else
            Comando = Left(LINHA, pos1 - 1)
            Conteudo = Mid$(LINHA, pos1 + 1)
        End If
        If Len(Var) = 0 Then
            If Len(Comando) = 0 Then
                LExtrai = RTrim(Conteudo)
                Exit Function
            End If
        Else
            If LItem(Comando, Var) <> 0 Then
                LExtrai = RTrim(Conteudo)
                Exit Function
            End If
        End If
    End If
    INI = pos + Len(SeparaLinha)
Wend
If VarType(Var) <> 8 Then
    If Var = 0 Then
        LExtrai = Seq
    End If
End If
LExtraiSai:
Exit Function
LExtraiErro:
Dim xerr
xerr = LErro("LExtrai")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LExtraiSai
End Function




' CMNGSTRING - INSERE UMA LINHA DENTRO DE UMA STRING
Function LInsere(Texto As Variant, Var As String, Conteúdo As Variant, Optional SeparaLinha As String = "^m^j", Optional SeparaDef As String = ":")
On Error GoTo LInsereErro
Dim INI As Integer, Última As Integer, pos As Integer
Dim LINHA As String, pos1 As Integer, Comando As String, údo As String
Dim Seq As Integer, TextoVar As String

INI = 1
Seq = 0
Última = False
Var = Trim(Var)
Texto = Nz(Texto, "")

If SeparaLinha = "^m^j" Then
    SeparaLinha = Chr(13) & Chr(10)
End If

While Última = False
BuscaLinha:
    pos = InStr(INI, Texto, SeparaLinha)
    If pos = 0 Then
        Última = True
        pos = Len(Texto) + 1
        LINHA = Mid$(Texto, INI)
    Else
        LINHA = Mid$(Texto, INI, pos - INI)
        If LINHA = "" Then
            Texto = Left(Texto, pos - 1) & Mid(Texto, pos + 2)
            GoTo BuscaLinha
        End If
    End If

    pos1 = InStr(LINHA, SeparaDef)
    If pos1 = 0 Then
        Comando = ""
        údo = LINHA
    Else
        Comando = Left(LINHA, pos1 - 1)
        údo = Mid$(LINHA, pos1 + 1)
    End If
    
    If Len(Var) = 0 Then
        If Len(Comando) = 0 Then
            GoTo AtualLinha
        End If
    Else
        If InStr(Comando, Var & ",") <> 0 Or InStr(Comando, Var & ";") <> 0 Or Comando = Var Then
            GoTo AtualLinha
        End If
    End If

    INI = pos + IIf(pos > Len(Texto), 0, Len(SeparaLinha))
Wend
Comando = Var

AtualLinha:
LINHA = IIf(Conteúdo <> "", Comando & IIf(Comando <> "", SeparaDef, "") & Conteúdo & SeparaLinha, "")
TextoVar = Left(Texto, INI - 1)
If TextoVar <> "" And Right(TextoVar, Len(SeparaLinha)) <> SeparaLinha Then
    TextoVar = TextoVar & SeparaLinha & LINHA
Else
    TextoVar = TextoVar & LINHA
End If
TextoVar = TextoVar & IIf(pos > Len(Texto), "", Mid(Texto, pos + Len(SeparaLinha)))
LInsere = TextoVar

LInsereSai:
Exit Function

LInsereErro:
Dim xerr As String
xerr = LErro("LInsere")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LInsereSai
End Function



' CMNGSTRING - CONVERTE UMA SEQUÊNCIA DE CARACTERES EM OUTRA DENTRO DE UMA STRING
Function LConvCarac(Texto As String, Carac As String, Conv As String, Optional VOLTA As Integer = True) As String
On Error GoTo LConvCaracErro
Dim pos, pos1 As Integer, Variav As String
Variav = Texto
pos = InStr(1, Variav, Carac)
Do While pos <> 0
    Variav = Mid(Variav, 1, pos - 1) & Conv & IIf(pos + Len(Carac) <= Len(Variav), Mid(Variav, pos + Len(Carac)), "")
    pos = InStr(IIf(VOLTA, 1, pos + Len(Conv)), Variav, Carac)
Loop
LConvCarac = RTrim(Variav)
LConvCaracSai:
Exit Function

LConvCaracErro:
Dim xerr As Integer
xerr = LErro("LConvCarac")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LConvCaracSai
End Function




' CMNGSTRING - FORMATA UM NÚMERO EM EXTENSO EM UM DETERMINADO IDIOMA
Function LExtenso(VV As Double, Ling As String) As String
On Error GoTo LExtensoErro
Dim StrUnid As String, StrDez As String, StrCem As String, StrMil As String, StrCent As String, MOEDA As String, SEP As String, StrPrimaDez As String
Dim StrVV As String, StrV0 As String, StrMoeda As String, StrM0 As String, Esp As String, SepDec As String
Dim pos As Integer, Z As Integer, ZZ As Integer, SepAndTrês As String
Dim SepMil As String
If Ling = "R$" Then
    StrUnid = "Um;Dois;Três;Quatro;Cinco;Seis;Sete;Oito;Nove"
    StrPrimaDez = "Onze;Doze;Treze;Quatorze;Quinze;Dezesseis;Dezesete;Dezoito;Dezenove"
    StrDez = "Dez;Vinte;Trinta;Quarenta;Cinquenta;Sessenta;Setenta;Oitenta;Noventa"
    StrCem = "Cento;Duzentos;Trezentos;Quatrocentos;Quinhentos;Seiscentos;Setecentos;Oitocentos;Novecentos"
    StrMil = "Mil.Mil;Milhão.Milhões;Bilhão.Bilhões;Trilhão.Trilhões"
    StrCent = "Centavo.Centavos"
    MOEDA = "Real.Reais.de Reais"
    SEP = " e "
    SepDec = " e "
    SepMil = ", "
    SepAndTrês = " e "
    Esp = " "
ElseIf Ling = "USD" Or Ling = "US$" Or Ling = "USD." Or Ling = "U.S.DLRS" Then
    StrUnid = "One;Two;Three;Four;Five;Six;Seven;Eight;Nine"
    StrPrimaDez = "Eleven;Twelve;Thirteen;Fourteen;Fifteen;Sixteen;Seventeen;Eighteen;Nineteen"
    StrDez = "Ten;Twenty;Thirty;Forty;Fifty;Sixty;Seventy;Eighty;Ninety"
    StrCem = "One Hundred;Two Hundred;Three Hundred;Four Hundred;Five Hundred;Six Hundred;Seven Hundred;Eight Hundred;Nine Hundred"
    StrMil = "Thousand.Thousand;Million.Millions;Billion.Billions;Trillion.Trillions"
    StrCent = "Cent.Cents"
    MOEDA = "U.S. Dollar;U.S. Dollars;U.S. Dollars"
    SEP = " "
    SepDec = " and "
    SepMil = ", "
    Esp = " "
    SepAndTrês = " and "
ElseIf Ling = "DM" Then
    StrUnid = "One;Two;Three;Four;Five;Six;Seven;Eight;Nine"
    StrPrimaDez = "Eleven;Twelve;Thirteen;Fourteen;Fifteen;Sixteen;Seventeen;Eighteen;Nineteen"
    StrDez = "Ten;Twenty;Thirty;Forty;Fifty;Sixty;Seventy;Eighty;Ninety"
    StrCem = "One Hundred;Two Hundred;Three Hundred;Four Hundred;Five Hundred;Six Hundred;Seven Hundred;Eight Hundred;Nine Hundred"
    StrMil = "Thousand.Thousand;Million.Millions;Billion.Billions;Trillion.Trillions"
    StrCent = "Cent.Cents"
    MOEDA = "Deutsche Mark;Deutsche Mark;Deutsche Mark"
    SEP = " "
    SepDec = " and "
    SepMil = ", "
    Esp = " "
    SepAndTrês = " and "
ElseIf Ling = "YEN" Then
    StrUnid = "One;Two;Three;Four;Five;Six;Seven;Eight;Nine"
    StrPrimaDez = "Eleven;Twelve;Thirteen;Fourteen;Fifteen;Sixteen;Seventeen;Eighteen;Nineteen"
    StrDez = "Ten;Twenty;Thirty;Forty;Fifty;Sixty;Seventy;Eighty;Ninety"
    StrCem = "One Hundred;Two Hundred;Three Hundred;Four Hundred;Five Hundred;Six Hundred;Seven Hundred;Eight Hundred;Nine Hundred"
    StrMil = "Thousand.Thousand;Million.Millions;Billion.Billions;Trillion.Trillions"
    StrCent = "Cent.Cents"
    MOEDA = "Yen;Yen;Yen"
    SEP = " "
    SepDec = " and "
    SepMil = ", "
    Esp = " "
    SepAndTrês = " and "
ElseIf Ling = "EURO" Then
    StrUnid = "One;Two;Three;Four;Five;Six;Seven;Eight;Nine"
    StrPrimaDez = "Eleven;Twelve;Thirteen;Fourteen;Fifteen;Sixteen;Seventeen;Eighteen;Nineteen"
    StrDez = "Ten;Twenty;Thirty;Forty;Fifty;Sixty;Seventy;Eighty;Ninety"
    StrCem = "One Hundred;Two Hundred;Three Hundred;Four Hundred;Five Hundred;Six Hundred;Seven Hundred;Eight Hundred;Nine Hundred"
    StrMil = "Thousand.Thousand;Million.Millions;Billion.Billions;Trillion.Trillions"
    StrCent = "Cent.Cents"
    MOEDA = "Euro;Euro;Euro"
    SEP = " "
    SepDec = " and "
    SepMil = ", "
    Esp = " "
    SepAndTrês = " and "
Else
    LExtenso = "#Erro"
    Exit Function
End If
StrVV = Format(VV, "000000000000000.00")
StrMoeda = ""

For Z = 1 To 6
    StrM0 = ""
    If Z <> 6 Then
        StrV0 = Mid(StrVV, Z * 3 - 2, 3)
        GoSub MontaCento
        If StrM0 <> "" Or Z = 5 Then
            If Z < 5 Then
                StrM0 = StrM0 & Esp & LItem(LItem(StrMil, 5 - Z), IIf(Val(StrV0) = 1, 1, 2))
            Else
                If Val(Mid(StrVV, 1, 15)) <> 0 Then
                    StrM0 = StrM0 & Esp & LItem(MOEDA, Switch(Mid(StrVV, 10, 6) = "000000", 3, Val(Mid(StrVV, 1, 15)) = 1, 1, Val(Mid(StrVV, 1, 15)) <> 1, 2))
                End If
            End If
        End If
    Else
        StrV0 = "0" & Mid(StrVV, 17, 2)
        GoSub MontaCento
        If StrM0 <> "" Then
            StrM0 = StrM0 & " " & LItem(StrCent, IIf(Val(StrV0) = 1, 1, 2))
        End If
    End If
    If StrM0 <> "" Then
        If Z = 6 Then
            StrMoeda = StrMoeda & IIf(StrMoeda <> "", SepDec, "") & StrM0
        Else
            StrMoeda = StrMoeda & IIf(StrMoeda <> "" And Val(StrV0) <> 0, IIf(Val(StrV0) < 101 Or Val(StrV0) Mod 100 = 0, SepAndTrês, SepMil), "") & StrM0
        End If
    End If
Next
LExtenso = StrMoeda
LExtensoSai:
Exit Function

MontaCento:
For ZZ = 1 To 3
    pos = Val(Mid(StrV0, ZZ, 1))
    If pos <> 0 Then
        Select Case ZZ
        Case 1
            StrM0 = StrM0 & LItem(StrCem, pos)
        Case 2
            If pos <> 1 Or Mid(StrV0, 3, 1) = "0" Then
                StrM0 = StrM0 & IIf(StrM0 <> "", SEP, "") & LItem(StrDez, pos)
            Else
                pos = Val(Mid(StrV0, 3, 1))
                StrM0 = StrM0 & IIf(StrM0 <> "", SEP, "") & LItem(StrPrimaDez, pos)
                ZZ = 3
            End If
        Case 3
            StrM0 = StrM0 & IIf(StrM0 <> "", SEP, "") & LItem(StrUnid, pos)
        End Select
    End If
Next
If StrM0 = "Cento" Then
    StrM0 = "Cem"
End If
Return

LExtensoErro:
Dim xerr
xerr = LErro("LExtenso")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LExtensoSai
End Function



' RETORNA UM CONTEÚDO EM UM DETERMINADO IDIOMA DE UMA CAMPO MULTILIGUE
' EX. DE CAMPO MULTILIGUE:
'        Bobina
'        I:Coil
'        C:Bobinex
Function LDescrLing(ByVal Campo As Variant, ByVal DOMÍNIO As String, ByVal Filtro As String, ByVal Ling As Variant) As String
On Error GoTo LDescrLingErro
Dim REC As DAO.Recordset, Ret As String, Palav As String
Static REF As Field
Ret = ""
Campo = IIf(IsNull(Campo), "", Campo)
Ling = Trim(IIf(IsNull(Ling), "", Ling))
If DOMÍNIO <> "" Then
    Set REC = CurrentDb.OpenRecordset(DOMÍNIO, dbOpenDynaset)
    REC.FindFirst Filtro
    If REC.NoMatch Then
        Ret = ""
    Else
LingNovamente:
        Ret = LExtrai(REC(Campo), IIf(Ling = "P", "", Ling))
        If Ret = "" Then
            If Ling Like "*+" Then
                Ling = Left(Ling, Len(Ling) - 1)
                GoTo LingNovamente
            End If
            Ret = LExtrai(REC(Campo), "")
        End If
    End If
Else
LingDireto:
    Ret = LExtrai(Campo, IIf(Ling = "P", "", Ling))
    If Ret = "" Then
        If Ling Like "*+" Then
            Ling = Left(Ling, Len(Ling) - 1)
            GoTo LingDireto
        End If
        Ret = LExtrai(Campo, "")
    End If
End If
LDescrLing = Ret
LDescrLingSai:
Exit Function

LDescrLingErro:
Dim xerr As Integer
xerr = LErro("LDescrLing")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LDescrLingSai
End Function




' FORMATA UMA STRING COM AS PRIMEIRAS LETRAS DE CADA PALAVRA EM MAIÚSCULO E O RESTO EM MINÚSCULO : NOME PRÓPRIO
Function LCorrigeNome(Nome As String)
On Error GoTo LCorrigeNomeErro
Dim Texto As String, RESULT As String, pos As Integer
Dim Prox As Integer, Palav As String
Texto = LCase(Nome)
RESULT = ""

pos = 1
Do While pos < Len(Texto)
    Prox = InStr(pos, Texto, " ")
    If Prox = 0 Then
        Prox = Len(Texto) + 1
    End If
    Palav = Mid(Texto, pos, Prox - pos)
    If Not Palav Like "d?" And Len(Palav) <> 1 Then
        Palav = UCase(Left(Palav, 1)) & Mid(Palav, 2)
    End If
    
    RESULT = RESULT & IIf(RESULT <> "", " ", "") & Palav
    pos = Prox + 1
Loop

LCorrigeNome = RESULT
LCorrigeNomeSai:
Exit Function

LCorrigeNomeErro:
Dim xerr
xerr = LErro("LCorrigeNome")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LCorrigeNomeSai
End Function




' LIMITA UM TEXTO EM UM NÚMERO DETERMINADO DE CARACTERES
Function LLimitText(ByVal Texto As Variant, Tam As Integer) As String
On Error GoTo LLimittextErro
Dim x, pos1, pos, Txt
Texto = Nz(Texto, "")
Texto = Texto & Space(Len(Texto) Mod Tam)
Txt = ""
pos = 1
Do While pos <= Len(Texto)
    pos1 = pos + Tam
    If Mid(Texto, pos1, 1) = " " Then
TrataLinha:
        Txt = RTrim(Txt & Mid(Texto, pos, pos1 - pos + 1) & Chr(13) & Chr(10))
        Do While Mid(Texto, pos1, 1) = " " And pos1 <= Len(Texto)
            pos1 = pos1 + 1
        Loop
        pos = pos1
    Else
        Do While Mid(Texto, pos1, 1) <> " " And pos1 > pos
            pos1 = pos1 - 1
        Loop
        If pos = pos1 Then
            pos1 = pos + Tam
        End If
        GoTo TrataLinha
    End If
Loop
If Txt = "" Then
    LLimitText = ""
Else
    LLimitText = Left(Txt, Len(Txt) - 2)
End If

LLimitTextsai:
Exit Function

LLimittextErro:
Dim xerr
xerr = LErro("LimitText")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LLimitTextsai
End Function



' MONTA UMA STRING NO FORMATO DE APRESENTAÇÃO COMUM PARA ENDEREÇAMENTO
Function LStrEnd(ByVal Ende As Variant, ByVal BAIRRO As Variant, ByVal cid As Variant, ByVal est As Variant, ByVal PAÍS As Variant, ByVal CEP As Variant, ByVal POBox As Variant, ByVal Líng As Variant, Optional ByVal Número As Variant = "", Optional ByVal Complemento As Variant = "") As String
On Error GoTo LStrEndErro
Dim Texto As String, pos As Integer, desc As String
Dim TROCA As Integer, DESCR As String
Ende = LDescrLing(IIf(Ende = ".", "", Ende), "", "", Líng)
BAIRRO = LDescrLing(IIf(BAIRRO = ".", "", BAIRRO), "", "", Líng)
cid = LDescrLing(IIf(cid = ".", "", cid), "", "", Líng)
est = LDescrLing(IIf(est = ".", "", est), "", "", Líng)

If LExists(CurrentDb.TableDefs, "PAÍS") Then
    DESCR = Nz(DLookup("Nome", "PAÍS", "[Cod] = '" & PAÍS & "'"), "")
Else
    DESCR = PAÍS
End If
PAÍS = LDescrLing(DESCR, "", "", Líng)

CEP = Nz(CEP, "")
POBox = Nz(POBox, "")

Número = Nz(Número, "")
Complemento = LConvCarac(Nz(Complemento, ""), ",", "|")

Texto = ";"
If Ende = "" Then
    Texto = Texto & IIf(POBox <> "", IIf(Líng = "I", "PO Box ", "Caixa Postal ") & POBox, "")
    Texto = Texto & ";" & cid & "," & est & "," & CEP & ";" & PAÍS
Else
    Texto = Texto & Ende & " " & Número & "|" & Complemento & "," & BAIRRO & ";" & cid & "," & est & "," & CEP & ";" & PAÍS
End If
Texto = Texto & ";"

Texto = LConvCarac(Texto, ",,", ",")
Texto = LConvCarac(Texto, ",;", ";")
Texto = LConvCarac(Texto, ";;", ";")
If Len(Texto) < 2 Then
    Texto = ""
Else
    Texto = Mid(Texto, 2, Len(Texto) - 2)
    Texto = Replace(Texto, ";", Chr(13) & Chr(10))
    Texto = Replace(Texto, ",", " - ")
End If
Texto = LConvCarac(Texto, "|", ", ")


LStrEnd = Texto
LStrEndSai:
Exit Function

LStrEndErro:
Dim xerr As Integer
xerr = LErro("LStrEnd")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LStrEndSai
End Function



' RETORNA UMA STRING COM UM NÚMERO FORMATADO EM UM DETERMINADO IDIOMA
Function LNumLing(NUM As Variant, Masc As String, Ling As Variant) As String
On Error GoTo LNumLingErro
Dim Texto As String
If IsNull(NUM) Then
    LNumLing = ""
Else
    Texto = ""
    If Not IsNull(NUM) Then
        Texto = Format(NUM, Masc)
        If Ling = "I" Then
            Texto = Replace(Texto, ",", "@")
            Texto = Replace(Texto, ".", ",")
            Texto = Replace(Texto, "@", ".")
        End If
    End If
    LNumLing = IIf(Texto = "", Null, Texto)
End If
LNumLingSai:
Exit Function

LNumLingErro:
Dim xerr As Integer
xerr = LErro("LNumLing")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LNumLingSai
End Function



' BUSCA ARQUIVOS CONFORME MÁSCARA NUM DETERMINADO DIRETÓRIO E SEUS FILHOS
Function LBuscaArquivos(Diretório As String, Máscara As String)
On Error GoTo LBuscaArquivosErro
Dim DD(5000) As String, SS As Integer, SS1 As Integer
Dim Arq As String, SS2 As Integer, Z As Integer, Txt As String

SS = 0
SS1 = SS

Arq = Diretório
If Arq <> "." And Arq <> ".." And GetAttr(Arq) = vbDirectory Then
    DD(SS) = Arq
    SS = SS + 1
End If

BUSCA_MAIS:
SS2 = SS

' BUSCA NÍVEIS POSTERIORES
For Z = SS1 To SS2 - 1
    Arq = dir(DD(Z) & "\*.", vbDirectory)
    Do While Arq <> ""
        If Arq <> "." And Arq <> ".." And GetAttr(DD(Z) & "\" & Arq) = vbDirectory Then
            DD(SS) = DD(Z) & "\" & Arq
            SS = SS + 1
        End If
        Arq = dir()
    Loop
Next

SS1 = SS2
If SS2 <> SS Then GoTo BUSCA_MAIS

' NOS DIRETÓRIOS ENCONTRADOS, BUSCA OS ARQUIVOS
Txt = ""
For Z = 0 To SS - 1
    Arq = dir(DD(Z) & "\" & Máscara)
    Do While Arq <> ""
        Txt = Txt & IIf(Txt <> "", Chr(13) & Chr(10), "") & DD(Z) & "\" & Arq
        Arq = dir()
    Loop
Next
LBuscaArquivos = Txt

LBuscaArquivosSai:
Exit Function

LBuscaArquivosErro:
Dim xerr
xerr = LErro("LBuscaArquivos")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LBuscaArquivosSai:
End Function





