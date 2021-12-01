Attribute VB_Name = "LVBE"
Option Compare Database
Option Explicit

' =========================================================
' OBJETO - LVBE - TRATAMENTO DE ERROS E FUNÇÕES DEPENDENTES
' =========================================================
' DATA - HISTÓRICO - TÉCNICO
' 23 MAI 2002 - IMPLEMENTAÇÃO CÓDIGO BASE

Global LErroRodando As Boolean



' TRATA ERRO - APRESENTA ERRO PARA O USUÁRIO E GRAVA EM LOG
Function LErro(Rotina, Optional ByVal DescrUsu As String = "") As Variant
Dim DESCR As String, ITEM As String, Z As Integer, bt As Integer
Dim NomUsu As String, NomApl As String, LogDir As String, LogArq As String
Dim ArqErr As String, FL As Integer, DescrMsg As String
LErroRodando = True
DESCR = "Erro no aplicativo."
 
If DescrUsu = "" Then
    ' salva condicionamento de erro
    If Err.Number <> 0 Then
        DESCR = "Err#" & Err.Description & "#" & Err.Number & "#" & Err.Source
    End If
    
    ' monta condicionamento caso probl em banco de dados
    If Err.Source = "DAO.Database" Or Err.Source = "DAO.Workspace" Or Err.Source = "DAO.Recordset" Or Err.Source = "DAO.QueryDef" Then
        For Z = 0 To DAO.Errors.Count - 1
            ITEM = Replace(DAO.Errors(Z).Description, Chr(10), "") & "#" & DAO.Errors(Z).Number & "#" & DAO.Errors(Z).Source
            If InStr(DESCR, ITEM) = 0 Then
                DESCR = DESCR & vbCrLf & Replace(Err.Source, "DAO.", "") & "#" & Replace(DAO.Errors(Z).Description, Chr(10), "") & "#" & DAO.Errors(Z).Number & "#" & DAO.Errors(Z).Source
            End If
        Next
    End If
    
    DescrUsu = DESCR
End If
 
' carrega indicadores diversos
NomUsu = LUsuário()
NomApl = LNomeApl()
LogDir = LConfig("ErrLocalArqsLog", , "C:\")
LogArq = LConfig("ErrArqLog", , NomApl & ".ERR")
ArqErr = Replace(LogDir & "\" & LogArq, "\\", "\")
 
' traduções de erros
DescrMsg = DescrUsu
If InStr(DescrMsg, "#2102#") <> 0 Then
    DescrMsg = "Recurso não disponível ou você não possui acesso. Verifique suas permissões para esse módulo ou contacte suporte especializado{OK}."
End If
 
' apresenta mensagem
If InStr(DescrMsg, "{OK}") Then
    bt = vbOKOnly
    DescrMsg = Replace(DescrMsg, "{OK}", "")
Else
    bt = vbAbortRetryIgnore
End If
 
LErro = MsgBox(DescrMsg, vbCritical + bt, NomApl & IIf(Rotina <> "", " (" & Rotina & ")", ""))
 
' registro em arquivo
FL = FreeFile()
Open ArqErr For Append As FL
Print #FL, ""
Print #FL, Format(Now(), "yyyy-mmm-dd hh:nn:ss ddd") & ", " & NomUsu & ", " & CurrentDb.Name
Print #FL, "Interno: " & DescrUsu
If DescrUsu <> DescrMsg Then
    Print #FL, "Usuário: " & DescrMsg
End If
Print #FL, "Resp:" & LErro & "-" & Switch(LErro = vbAbort, "ABORT", LErro = vbRetry, "RETRY", LErro = vbIgnore, "IGNORE", LErro = vbOK, "OK", True, "DESCONHECIDA")
Close #FL
LErroRodando = False
End Function


' RETORNA O NOME DO APLICATIVO - UTILIZADA POR TRATA ERRO
Function LNomeApl()
Dim PosIni As Integer, posfim As Integer
On Error GoTo LNomeAplErro
Dim Txt As String

LNomeApl = CurrentDb.Name

Txt = LConfig("Aplicação")
If Txt = "[APLICAÇÃO]" Then
    Txt = CurrentDb.Name
    posfim = InStrRev(Txt, ".")
    If posfim = 0 Then
        posfim = Len(Txt) + 1
    End If
    PosIni = InStrRev(Txt, "\", posfim - 1)
    Txt = Mid(Txt, PosIni + 1, posfim - PosIni - 1)
End If
LNomeApl = Txt

LNomeAplSai:
Exit Function

LNomeAplErro:
If Not LErroRodando Then
    Dim xerr
    xerr = LErro("LNomeApl")
    If xerr = 4 Then
        Resume 0
    ElseIf xerr = 5 Then
        Resume Next
    End If
End If
Resume LNomeAplSai
End Function



' RETORNA NOME DO USUÁRIO CORRENTE - UTILIZADA POR TRATA ERRO
' rotina criada para considerar possibilidade de controle
' de usuário pelo aplicativo e não pela system.mdw
Function LUsuário()
On Error GoTo LUsuárioErro
LUsuário = UCase(CurrentUser())

LUsuárioSai:
Exit Function

LUsuárioErro:
If Not LErroRodando Then
    Dim xerr
    xerr = LErro("LUsuário")
    If xerr = 4 Then
        Resume 0
    ElseIf xerr = 5 Then
        Resume Next
    End If
End If
Resume LUsuárioSai
End Function



' CONFIGURA OU RETORNA O CONTEÚDO DE UM PARÂMETRO LOCAL : REGISTRADO NA TABELA LCONFIGLOCAL
' UTILIZADA POR TRATA ERRO
' OPÇÕES CONHECIDAS
'
'    LOCAL_ARQS_LOG     localização de arquivos de log
'    ARQ_LOG_ERR        nome do arquivo de log de erros

Function LConfig(Campo As String, Optional Valor, Optional DEF)
On Error GoTo LConfigErro
Dim Conteúdo As Variant, REC As DAO.Recordset
If VarType(Valor) = 10 Then
    LConfig = IIf(IsMissing(DEF), "[" & Campo & "]", DEF)
    On Error Resume Next
    Err.Clear
    Conteúdo = DLookup("CONFIG", "SYS_CONFIG_LOCAL", "[PARAM] = '" & Campo & "'")
    If Err = 0 Then
        If Not IsNull(Conteúdo) Then
            LConfig = Conteúdo
        End If
    End If
Else
    LConfig = 0
    If VarType(Valor) = vbNull Then
        CurrentDb.Execute "DELETE * FROM SYS_CONFIG_LOCAL WHERE PARAM = '" & Campo & "';"
    Else
        Set REC = CurrentDb.OpenRecordset("SELECT * FROM SYS_CONFIG_LOCAL WHERE PARAM = '" & Campo & "';")
        If REC.RecordCount <> 0 Then
            REC.Edit
        Else
            REC.AddNew
            REC!Param = Campo
        End If
        REC!Config = Valor & ""
        REC.Update
    End If
End If
LConfigSai:
Exit Function

LConfigErro:
If Not LErroRodando Then
    Dim xerr
    xerr = LErro("LConfig")
    If xerr = 4 Then
        Resume 0
    ElseIf xerr = 5 Then
        Resume Next
    End If
End If
Resume LConfigSai
End Function




