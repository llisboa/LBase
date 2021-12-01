Attribute VB_Name = "LVBB"
Option Compare Database
Option Explicit

' =========================================
' OBJETO - LVBB - MANUTENÇÃO DE SOFT BÁSICO
' =========================================
' DATA - HISTÓRICO - TÉCNICO
' 23 MAI 2002 - NORMALIZAÇÃO DAS FUNÇÕES

' tipos de busca de janela
Public Enum wflags
    GW_HWNDFIRST
    GW_HWNDLAST
    GW_HWNDNEXT
    GW_HWNDPREV
    GW_OWNER
    GW_CHILD
End Enum

' declarações de bibliotecas externas
Declare Function LGetWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Declare Function LSetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long


' RETORNA CÓDIGO DE UM CONTROLE (CÓDIGO DO TIPO DE CRIAÇÃO)
Function LControlCode(CC As Control) As Integer
On Error GoTo LControlCodeErro
Dim CCode As Integer
If TypeOf CC Is Label Then
    CCode = 100
ElseIf TypeOf CC Is rectangle Then
    CCode = 101
ElseIf TypeOf CC Is Line Then
    CCode = 102
ElseIf TypeOf CC Is CommandButton Then
    CCode = 104
ElseIf TypeOf CC Is OptionButton Then
    CCode = 105
ElseIf TypeOf CC Is CheckBox Then
    CCode = 106
ElseIf TypeOf CC Is OptionGroup Then
    CCode = 107
ElseIf TypeOf CC Is BoundObjectFrame Then
    CCode = 108
ElseIf TypeOf CC Is TextBox Then
    CCode = 109
ElseIf TypeOf CC Is ListBox Then
    CCode = 110
ElseIf TypeOf CC Is ComboBox Then
    CCode = 111
ElseIf TypeOf CC Is SubForm Then
    CCode = 112
ElseIf TypeOf CC Is SubReport Then
    CCode = 112
ElseIf TypeOf CC Is ObjectFrame Then
    CCode = 114
ElseIf TypeOf CC Is PageBreak Then
    CCode = 118
ElseIf TypeOf CC Is ToggleButton Then
    CCode = 122
Else
    CCode = 0
End If
LControlCode = CCode

LControlCodeSai:
Exit Function
LControlCodeErro:
Dim xerr As Integer
xerr = LErro("LControlCode")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LControlCodeSai
End Function



' APRESENTA OBJETOS COMUNS PARA O DESENVOLVIMENTO
Function LManut()
On Error Resume Next
Dim Z As Integer
LSetWindowText Application.hWndAccessApp, LConfig("Aplicação") & " - " & CurrentDb.Name
LProperty CurrentDb, "StartUpShowDBWindow", True
LProperty CurrentDb, "AllowBuiltInToolbars", True
LProperty CurrentDb, "AllowToolbarChanges", True
Application.SetOption "Show Hidden Objects", True
Application.SetOption "Show System Objects", True
CommandBars("Visual Basic").Reset
For Z = 1 To Application.CommandBars.Count
    Application.CommandBars(Z).Enabled = True
Next
End Function


' COPIA MÓDULOS PARA OUTROS APLICATIVOS
Function LCopia_Módulos()
Dim RAIZ_FONTES As String, DESTINOS As String, Z As Integer, FOR_PRIMEIRO As Boolean
Dim ACC As Access.Application, TIPO As Integer, DESTDIR As String
Dim ELEMENTOS As String, PRIMEIROS As String, ZZ As Integer, Elem As String

DESTDIR = "D:\UNIF\"
DESTINOS = "CIEX_V04.06.mdb;ELOG_V02.02.mdb;EMAILLOG_V02.12.mdb;TAREFA_V04.03.mdb"

ELEMENTOS = _
"MLVBA;MLVBB;MLVBC;MLVBE;MLVBI;MLVBT;MLVBV;MLCDLG;" & _
"FLBuscaDinâmica;FLConfigDelete;FLConfigGlobal;" & _
"FLConfigHistórico;FLConfigLocal;FLConfigOcorrência;" & _
"FLConfigUsuário;FLConfigVídeo;FLCurtaMensagem;FLDataHora;" & _
"FLDicasDeTeclas;FLMensagem;FLVazio;FLAttach"

PRIMEIROS = _
"FLSobreSistema;FLTelaPrima;" & _
"TSYS_CONFIG_LOCAL;TSYS_CONFIG_GLOBAL;TSYS_GRUPO_REL;TSYS_OCORRÊNCIA;TSYS_USUÁRIO;" & _
"AAUTOEXEC;AAUTOKEYS;ATECLASATALHO"


For Z = 1 To LItem(DESTINOS, 0, ";")

    RAIZ_FONTES = LItem(DESTINOS, Z, ";")
    If dir(DESTDIR & RAIZ_FONTES) = "" Then
        FOR_PRIMEIRO = True
        DBEngine.CreateDatabase DESTDIR & RAIZ_FONTES, dbLangGeneral
    Else
        FOR_PRIMEIRO = False
    End If
    
    If (DESTDIR & RAIZ_FONTES) <> CurrentDb.Name Then

        FileCopy DESTDIR & RAIZ_FONTES, DESTDIR & RAIZ_FONTES & ".BKP"
        Set ACC = New Access.Application
        ACC.OpenCurrentDatabase DESTDIR & RAIZ_FONTES, True

        ' EXCLUI TUDO
        For ZZ = 1 To LItem(ELEMENTOS, 0)
            Elem = LItem(ELEMENTOS, ZZ)
            TIPO = Switch(Left(Elem, 1) = "M", acModule, Left(Elem, 1) = "T", acTable, Left(Elem, 1) = "F", acForm, Left(Elem, 1) = "O", acMacro)
            On Error Resume Next
            ACC.DoCmd.DeleteObject TIPO, Mid(Elem, 2)
            On Error GoTo 0
        Next
        
        If FOR_PRIMEIRO Then
            For ZZ = 1 To LItem(PRIMEIROS, 0)
                Elem = LItem(PRIMEIROS, ZZ)
                TIPO = Switch(Left(Elem, 1) = "M", acModule, Left(Elem, 1) = "T", acTable, Left(Elem, 1) = "F", acForm, Left(Elem, 1) = "O", acMacro)
                ACC.DoCmd.DeleteObject TIPO, Mid(Elem, 2)
                On Error GoTo 0
            Next
        End If
    
        ' LIBERA ARQUIVO
        ACC.CloseCurrentDatabase
        
        ' COPIA TUDO
        For ZZ = 1 To LItem(ELEMENTOS, 0)
            Elem = LItem(ELEMENTOS, ZZ)
            TIPO = Switch(Left(Elem, 1) = "M", acModule, Left(Elem, 1) = "T", acTable, Left(Elem, 1) = "F", acForm, Left(Elem, 1) = "O", acMacro)
            DoCmd.CopyObject DESTDIR & RAIZ_FONTES, Mid(Elem, 2), TIPO, Mid(Elem, 2)
        Next
        
        If FOR_PRIMEIRO Then
            For ZZ = 1 To LItem(PRIMEIROS, 0)
                Elem = LItem(PRIMEIROS, ZZ)
                TIPO = Switch(Left(Elem, 1) = "M", acModule, Left(Elem, 1) = "T", acTable, Left(Elem, 1) = "F", acForm, Left(Elem, 1) = "O", acMacro)
                DoCmd.CopyObject DESTDIR & RAIZ_FONTES, Mid(Elem, 2), TIPO, Mid(Elem, 2)
            Next
        End If

    End If
Next
End Function





