Attribute VB_Name = "LVBV"
Option Compare Database
Option Explicit

' ==================================================
' OBJETO - LVBV - CHECAGENS E ESTRUTURAS NÃO STRINGS
' ==================================================
' DATA - HISTÓRICO - TÉCNICO
' 23 MAI 2002 - IMPLEMENTAÇÃO CÓDIGO BASE

' ARREDONDA UM VALOR EM UM NÚMERO DE CASAS DECIMAIS ESPECÍFICO
Function LNumDec(Nm As Variant, Dec As Integer) As Double
On Error GoTo LNumDecErro
Dim DD As Double, tt As Double, NUM As Double
If IsNull(Nm) Then
    LNumDec = 0
Else
    NUM = Nm
    DD = 10 ^ Dec
    tt = Val(Str(NUM * DD))
    LNumDec = Int(((tt - Int(tt)) * 2 + Int(tt))) / DD
End If

LNumDecSai:
Exit Function

LNumDecErro:
Dim xerr As Integer
xerr = LErro("NumDec")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LNumDecSai
End Function



' FORMATA DATA NUMA SEQUÊNCIA ESPECIAL
Function LData(Dat As String, Optional Fo As String) As String
On Error GoTo LdataErro
Dim MÊS As String, Cria As String
Dim Dia As Integer, Suf As Integer
MÊS = "Janeiro.January.Enero;Fevereiro.February.Febrero;Março.March.Marzo;Abril.April.Abril;Maio.May.Mayo;Junho.June.Junio;Julho.July.Julio;Agosto.August.Agosto;Setembro.September.Septiembre;Outubro.October.Octubre;Novembro.November.Noviembre;Dezembro.December.Diciembre"
Cria = ""
Dia = Day(Dat)
Suf = Dia Mod 10
If Trim(Fo) = "I" Then
    Cria = LItem(LItem(MÊS, Month(Dat)), 2) & " " & Dia & Switch(Dia > 10 And Dia < 14, "th", Suf = 1, "st", Suf = 2, "nd", Suf = 3, "rd", True, "th") & ", " & Year(Dat)
ElseIf Trim(Fo) = "A" Then
    Cria = Format(Dia, "00") & " " & UCase(Left(LItem(LItem(MÊS, Month(Dat)), 1), 3)) & " " & Format(Year(Dat), "0000")
ElseIf Trim(Fo) = "AI" Then
    Cria = Format(Dia, "00") & " " & UCase(Left(LItem(LItem(MÊS, Month(Dat)), 2), 3)) & " " & Format(Year(Dat), "0000")
ElseIf Trim(Fo) = "ORA" Then
    Cria = "TO_DATE('" & Format(Dat, "DD/MM/YYYY") & "', 'DD/MM/YYYY')"
ElseIf Trim(Fo) = "mmm dd, yyyy I" Then
    Cria = UCase(Left(LItem(LItem(MÊS, Month(Dat)), 2), 3)) & " " & Dia & ", " & Year(Dat)
ElseIf Trim(Fo) Like "mmm dd, yyyy*" Then
    Cria = UCase(Left(LItem(LItem(MÊS, Month(Dat)), 1), 3)) & " " & Dia & ", " & Year(Dat)
ElseIf Trim(Fo) Like "mmmm, yyyy I" Then
    Cria = LItem(LItem(MÊS, Month(Dat)), 2) & ", " & Format(Year(Dat), "0000")
ElseIf Trim(Fo) Like "mmmm, yyyy C" Then
    Cria = LItem(LItem(MÊS, Month(Dat)), 3) & ", " & Format(Year(Dat), "0000")
ElseIf Trim(Fo) = "C" Then
    Cria = Dia & " de " & LItem(LItem(MÊS, Month(Dat)), 3) & " de " & Year(Dat)
Else
    Cria = Dia & " de " & LItem(LItem(MÊS, Month(Dat)), 1) & " de " & Year(Dat)
End If
LData = Cria

LDataSai:
Exit Function

LdataErro:
Dim xerr
xerr = LErro("LData")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LDataSai
End Function



' FORMATA UM TIPO ESPECÍFICO EM STRING PARA SQL
Function LExprTipoVar(Param As Variant) As String
On Error GoTo LExprTipoVarErro
Dim tt As String, Texto As Variant, pos As Integer
If VarType(Param) = vbCurrency Or VarType(Param) = vbLong Or VarType(Param) = vbSingle Or VarType(Param) = vbDouble Or VarType(Param) = vbInteger Then
    LExprTipoVar = Str(Param)
ElseIf VarType(Param) = vbDate Then
    LExprTipoVar = "#" & Format(CVDate(Param), "mm/dd/yy") & "#"
ElseIf VarType(Param) = vbNull Then
    LExprTipoVar = "NULL"
ElseIf VarType(Param) = vbBoolean Then
    LExprTipoVar = IIf(Param, "TRUE", "FALSE")
Else
    Texto = Param
    pos = InStr(1, Texto, Chr(34))
    Do While pos <> 0
        Texto = Left(Texto, pos) & Chr(34) & Mid(Texto, pos + 1)
        pos = InStr(pos + 2, Texto, Chr(34))
    Loop
    LExprTipoVar = Chr(34) & Texto & Chr(34)
End If
LExprTipoVarSai:
Exit Function

LExprTipoVarErro:
Dim xerr As Integer
xerr = LErro("LExprTipoVar")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LExprTipoVarSai
End Function





' RETORNA O RESULTADO DA COMPARAÇÃO ENTRE DOIS PARÂMETROS QUAISQUER
Function LCompAtrStr(param1, param2)
On Error GoTo LcompAtrStrErro
LCompAtrStr = True
If Not IsNull(param1) Then
    If param1 <> "" Then
        If IsNull(param2) Then
            LCompAtrStr = False
        ElseIf param2 = "" Then
            LCompAtrStr = False
        ElseIf Not param1 = param2 Then
            LCompAtrStr = False
        End If
    End If
End If
LcompAtrStrSai:
Exit Function

LcompAtrStrErro:
Dim xerr As Integer
xerr = LErro("LCompAtrStr")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LcompAtrStrSai
End Function



' VERIFICA SE UMA DATA ESTÁ NUM INTERVALO ESPECÍFICO
Function LValidaData(ByVal DATAINI As Variant, ByVal DATAFIM As Variant, CUR As Variant) As Integer
On Error GoTo LValidaDataErro
LValidaData = False
If IsNull(CUR) And (IsNull(DATAINI) Or IsNull(DATAFIM)) Then
   LValidaData = True
   Exit Function
End If
If IsNull(DATAINI) Then
    DATAINI = #1/1/1900#
End If
If IsNull(DATAFIM) Then
    DATAFIM = #12/31/9999#
End If
If CUR >= DATAINI And CUR <= DATAFIM Then
    LValidaData = True
    Exit Function
End If
LValidaDataSai:
Exit Function
LValidaDataErro:
Dim xerr As Integer
xerr = LErro("LValidaData")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LValidaDataSai
End Function



' VERIFICA SE UM PERÍODO ESTÁ NUM INTERVALO ESPECÍFICO
Function LValidaPeríodo(ByVal DATAINI As Variant, ByVal DATAFIM As Variant, CUR As Variant) As Integer
On Error GoTo LValidaPeríodoErro
LValidaPeríodo = False
If IsNull(CUR) And (IsNull(DATAINI) Or IsNull(DATAFIM)) Then
   LValidaPeríodo = True
   Exit Function
End If
If IsNull(DATAINI) Then
    DATAINI = #1/1/1900#
End If
If IsNull(DATAFIM) Then
    DATAFIM = #12/31/9999#
End If
CUR = Format(CUR, "yy/mm")
DATAINI = Format(DATAINI, "yy/mm")
DATAFIM = Format(DATAFIM, "yy/mm")
If CUR >= DATAINI And CUR <= DATAFIM Then
    LValidaPeríodo = True
    Exit Function
End If
LValidaPeríodoSai:
Exit Function
LValidaPeríodoErro:
Dim xerr
xerr = LErro("LValidaPeríodo")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LValidaPeríodoSai
End Function




