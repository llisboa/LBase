Attribute VB_Name = "LVBA"
Option Compare Database
Option Explicit

' ================================================
' OBJETO - LVBA - OBJETOS NATIVOS MICROSOFT ACCESS
' ================================================
' DATA - HISTÓRICO - TÉCNICO
' 23 MAI 2002 - NORMALIZAÇÃO DAS FUNÇÕES - LUCIANO
' 20 MAR 2003 - INCLUSÃO DE FUNÇÕES - LUCIANO
' 29 MAI 2003 - INCLUSÃO DE HOOK (FORM LINÍCIOAPL) EM LINÍCIOAPL - LUCIANO
' 05 AGO 2003 - ATUALIZAÇÃO DO LATTACH COMPATÍVEL COM ACCESS E ORACLE - LUCIANO
' 06 AGO 2003 - APAGADA A CLASSE LCDLG E ADICIONADA FUNÇÃO LABREARQUIVO PARA SUBSTITUIÇÃO - LUCIANO
' 07 AGO 2003 - CRIAÇÃO DA FUNÇÃO LITEMARQUIVO - CABRAL
' 03 MAI 2004 - INCLUSÃO DA ROTINA PROTECT NA LLIB - ABSTRACARAC RETORNANDO COD TB QUANDO CÓPIA XX É PASSADA - CABRAL
' 02 FEV 2005 - INCLUSÃO DA ROTINA PROTECT NOVO - 32 BITS - CABRAL
' 20 FEV 2005 - ALTERAÇÃO LMENSAGEM PARA CONSIDERAR MODAL - LATTACH PERMITINDO SENHA EM LIGAÇÃO COM ACCESS - LUCIANO
' 08 ABR 2005 - INCLUSÃO DE LATTACH DE TABELAS DO FIREBIRD - CABRAL
' 02 JUN 2005 - CRIAÇÃO DE NOVO LBASE COM FUNÇÕES LINÍCIO E LATTACH_OK ALTERADAS E CRIAÇÃO DE NOVO MDB - CABRAL
' 02 JUN 2005 - RETIRADA DO MÓDULO LVBP DAS PROTEÇÕES. AS FUNÇÕES ESTÃO NA DLL LLIB - CABRAL
' 02 JUN 2005 - EXCLUSÃO DO FORM LSYSCONFIGEMPRESA. O FORMULÁRIO DE LICENÇA ESTÁ NA DLL LLIB - CABRAL
' 03 JUN 2005 - FUNÇÕES DE MANIPULAÇÃO DO REGISTRO DO WINDOWS DO LLIB ADICIONADOS - CABRAL
' 17 JUN 2005 - FUNÇÃO LSOBRESISTEMA PARA ABRIR O FORM LSOBRESISTEMA - CABRAL
' 20 JUL 2005 - IMPLEMENTAÇÃO DO LBACKUP E LRESTORE - CABRAL
' 11 OUT 2005 - ALTERAÇÃO DE ESTRUTURA LATTACH E LINÍCIO - CABRAL

' definições de dll windows
Declare Function LGetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As Long
Declare Function LGetClientRect Lib "user32" Alias "GetClientRect" (ByVal hWin As Long, rectangle As Rect) As Long
Declare Function LGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hWin As Long, rectangle As Rect) As Long
Declare Function LShowWindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal i As Long) As Long
Declare Function LMoveWindow Lib "user32" Alias "MoveWindow" (ByVal hWin As Long, ByVal x As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal fRepaint As Long) As Long
Declare Function LFindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hParent As Long, ByVal hChildAfter As Long, ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long

Declare Function LGetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As String) As Long

Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegEnumValue Lib "advapi32" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

' declare base em funções dll llib versão
' >>>>>>>>>>
' OBS: quando for necessário mudar a versão da biblioteca, alterar também em linício a rotina de checagem
'          <<<<<<<<<<<
Declare Function LEncaixaTela Lib "LLib0224.dll" (ByVal Continente As Long, ByVal Conteúdo As Long) As Long
Declare Function LDesencaixaTela Lib "LLib0224.dll" (ByVal Continente As Long) As Long
Declare Function LDesencaixaTelas Lib "LLib0224.dll" () As Long
Declare Function LAbreArquivo Lib "LLib0224.dll" (ByVal Janela As Long, ByVal Titulo As String, ByVal Estilo As Long, ByVal Filtro As String, ByVal Arquivo As String) As Long
Declare Function LTrataSenha Lib "LLib0224.dll" (ByVal Senha As String) As Long
Declare Function LCheck_ Lib "LLib0224.dll" Alias "LProtCheck" (ByVal Fabric As String, ByVal Prod As String, ByVal Ver As String, Optional ByVal JanPai As Long = 0) As Long
Declare Function LAbstr_ Lib "LLib0224.dll" Alias "LAbstraCarac" (ByVal Texto As String) As Integer
Declare Function LExecShell Lib "LLib0224.dll" (ByVal Comando As String, ByVal Esperar As Long) As Long

Global LTelaPrima As String          ' nome da tela de fundo
Global LMensagemSaída As String      ' retorno do diálogo de lmensagem
Global StringConnect As String       ' estrutura de conexão
Global Const LRelBackColor = &HC0C0C0 ' cor de fundo

'Constantes para função DlgOpenFile
Global Const LLIB_OFN_CARREGAR_ARQ = &H2
Global Const LLIB_OFN_CARREGAR_ARQ_MULT = &H4
Global Const LLIB_OFN_SALVAR_ARQ = &H8
Global Const LLIB_OFN_CARREGAR_DIR = &H10

'Outros valores globais
Global Const MAX_PATH As Integer = 260
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const READ_CONTROL = &H20000
Global Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Global Const KEY_ENUMERATE_SUB_KEYS = &H8
Global Const KEY_NOTIFY = &H10
Global Const SYNCHRONIZE = &H100000
Global Const KEY_QUERY_VALUE = &H1
Global Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

' estrutura para manipulação de interfaces
Type Rect
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Dim dlgw As Long

' CÓDIGO SEMPRE EXECUTADO QUANDO SE INICIA O SISTEMA
Function LInício()
On Error GoTo LInícioErro
Dim Esquema As String, FABRICANTE As String, Aplicação As String, Versão As String, DescrApl As String, BaseDados As String, Usuário As String, Senha As String
Dim SS As Integer, SS_TOT As Integer, RR, CONN As Recordset, CONNSTR As String, Licença As String, DataLicença As Variant, x As Integer, EMPRESA As String, LocalInstalação As String, EmpresaContato As String, TelefoneContato As String
Dim Mark As Boolean, TBD As TableDef, Qry As QueryDef
Dim RT As Rect
Dim BackupDataDiff As Integer
Dim Wrk As Workspace, DB As Database
SS_TOT = 8

' simplifica tela, inicia progresso, ícone e variáveis
LBarras "Aplic_Menu;Aplic_Barra;Aplic_Acesso_Rápido"
LProperty CurrentDb, "AppIcon", Left(CurrentDb.Name, InStrRev(CurrentDb.Name, "\")) & LConfig("Ícone", , "")
DescrApl = LConfig("Aplicação") & " " & LConfig("Versão") & " - " & LConfig("DescriçãoAplicação", , "") & " - Intercraft Solutions"
LSetWindowText Application.hWndAccessApp, DescrApl
Application.RefreshTitleBar

' carrega tela de apresentação
If LConfig("LTelaApresenta", , "") <> "" Then
    DoCmd.Echo False
    DoCmd.OpenForm "LTelaApresenta"
    Forms("LTelaApresenta")("ImgTelaApres").PictureData = LConfig("LTelaApresenta")
    DoCmd.RepaintObject acForm, "LTelaApresenta"
    LPausa 3
    DoCmd.Echo True
    DoCmd.Close acForm, "LTelaApresenta"
End If

SS = LProgress(SS, LConfig("Aplicação") & " - Inicializando Aplicativo", SS_TOT, , "LCurtaMensagem") '0

' checa dll e inicia variáveis
On Error Resume Next
LTrataSenha "TESTE"
If Err <> 0 Then
    LErro "LInício", "{OK}Este programa necessita da biblioteca LLib0224.dll que não pode ser encontrada pelo sistema operacional. Reinstale o aplicativo ou faça atualização manual."
    Application.Quit acQuitSaveNone
End If

On Error GoTo LInícioErro
' verifica fabricante, produto e versão
FABRICANTE = LConfig("FABRICANTE", , "")
Aplicação = LConfig("APLICAÇÃO", , "")
Versão = LConfig("VERSÃO", , "")
If FABRICANTE = "" Or Aplicação = "" Or Versão = "" Then
    LErro "LInício", "{OK}Este programa necessita de definição do parâmetro 'fabaplver' pelo fornecedor. Contacte suporte autorizado."
    Application.Quit acQuitSaveNone
End If
SS = LProgress(SS, LConfig("Aplicação") & " - Estabelecendo conexão com a base de dados", SS_TOT, , "LCurtaMensagem") '1

' ativa conexão com banco de dados
If Not LExists(CurrentDb.TableDefs, "SYS_CONFIG_GLOBAL") Then
    If MsgBox("O aplicativo " & LConfig("Aplicação") & " não encontrou definições de conexão com a base de dados. Deseja configurar agora?", vbExclamation Or vbYesNo, LConfig("Aplicação")) = vbNo Then
        MsgBox "Configuração de conexão com o banco de dados cancelado pelo usuário. O aplicativo será fechado.", vbCritical, LConfig("Aplicação")
        Application.Quit acQuitSaveNone
    End If
    LAttach True
End If

BaseDados = LConfig("ConnBaseDados", , "")

connectnovamente:
On Error Resume Next
LLogin Usuário, Senha
Select Case BaseDados
    Case "Microsoft Access"
        StringConnect = "MS Access;DATABASE=" & LConfig("ConnArquivo") & ";UID=" & IIf(Usuário = "", "Admin", UCase(Usuário))
        If Senha <> "" Then StringConnect = StringConnect & ";PWD=" & Senha
    Case "Oracle"
        StringConnect = CurrentDb.TableDefs("SYS_CONFIG_GLOBAL").Connect
        StringConnect = LInsere(StringConnect, "DATABASE", "", ";", "=")
        StringConnect = StringConnect & ";UID=" & Usuário & ";PWD=" & Senha
    Case "Firebird"
        StringConnect = CurrentDb.TableDefs("SYS_CONFIG_GLOBAL").Connect
        StringConnect = StringConnect & ";UID=" & Usuário & ";PWD=" & Senha
End Select

Set Wrk = DBEngine.CreateWorkspace(LConfig("Aplicação") & "_logon_workspace", "Admin", "", IIf(BaseDados = "Microsoft Access", dbUseJet, dbUseODBC))
Set DB = Wrk.OpenDatabase("", dbDriverNoPrompt, False, StringConnect)
If Err <> 0 Then GoTo connecterro
DB.Close
Wrk.Close

CurrentDb.TableDefs.Append CurrentDb.CreateTableDef("SYS_CONN", 0, IIf(BaseDados = "Oracle", LConfig("ConnEsquema") & ".", "") & "SYS_CONFIG_GLOBAL", StringConnect)
If Err <> 0 Then GoTo connecterro
Set CONN = CurrentDb.OpenRecordset("SYS_CONN")
If Err <> 0 Then
connecterro:
    RR = LMensagem("Problemas ao tentar conectar com base de dados. {Sair}[Sair] do aplicativo, {Repetir}[Repetir] para tentar vincular novamente ou {Vincular}[Vincular] para estabelecer novo vínculo com base de dados{Modal} (" & Err.Description & ").")
    If RR = "Repetir" Then
        GoTo connectnovamente
    ElseIf RR = "Vincular" Then
connectlinktab:
        On Error GoTo LInícioErro
        DoCmd.OpenForm "LAttach"
        Do While LExists(Forms, "LAttach")
            DoEvents
        Loop
        LMensagem Null
        LProgress SS, LConfig("Aplicação") & " - Estabelecendo conexão com servidor de dados", SS_TOT, , "LCurtaMensagem"
        GoTo connectnovamente
    Else
        LErro "LInício", "{OK}Este programa necessita de vínculo ao banco de dados. Restabeleça-o antes de continuar."
        Application.Quit acQuitSaveNone
    End If
End If

CONN.Close
CurrentDb.TableDefs.Delete "SYS_CONN"

SS = LProgress(SS, LConfig("Aplicação") & " - Atualizando conexão dos objetos ao banco de dados", SS_TOT, , "LCurtaMensagem") '2
For Each TBD In CurrentDb.TableDefs
    If TBD.Connect <> "" Then
        TBD.Connect = StringConnect
    End If
Next

For Each Qry In CurrentDb.QueryDefs
    If Qry.Type = dbSQLPassThrough Then
        Qry.Connect = StringConnect
    End If
Next

On Error GoTo 0
SS = LProgress(SS, LConfig("Aplicação") & " - Verificando autenticidade do produto", SS_TOT, , "LCurtaMensagem") '3

' verifica se licença está válida
If LCheck_(LConfig("FABRICANTE"), LConfig("APLICAÇÃO"), LConfig("VERSÃO"), Application.hWndAccessApp) <> 0 Then
    LErro "LInício", "{OK}Este programa necessita de liberação. Contacte suporte autorizado."
    Application.Quit acQuitSaveNone
End If
SS = LProgress(SS, LConfig("Aplicação") & " - Realizando checagem de ambiente", SS_TOT, , "LCurtaMensagem") '4

'Verificação de backup
If LSistema("SuporteBackup", , "") <> "" Then
    If LConfig("LBackup_Data", , "") = "" Then
        LConfig "LBackup_Data", Format(Now, "dd/mm/yyyy")
    End If
    BackupDataDiff = DateDiff("d", LConfig("LBackup_Data"), Format(Now, "dd/mm/yyyy"))
    If BackupDataDiff > LConfig("LBackup_Intervalo", , 7) Then
        If MsgBox("O último backup do sistema foi realizado à " & BackupDataDiff & " dias. Deseja efetuar o backup agora?", vbQuestion Or vbYesNo Or vbApplicationModal, LConfig("Aplicação")) = vbYes Then
            LBackup
        End If
    End If
End If

' resolução de vídeo
SS = LProgress(SS) '5
LConfigVídeo False

' seta nome da janela novamente
LSetWindowText Application.hWndAccessApp, DescrApl

' condicionamento do ambiente 1
SS = LProgress(SS) '6
LProperty CurrentDb, "StartUpShowDBWindow", False

' Application.SetOption "Show Startup Dialog Box", False
Application.SetOption "Default Find/Replace Behavior", 1
Application.SetOption "Confirm Record Changes", True
Application.SetOption "Confirm Document Deletions", True
Application.SetOption "Confirm Action Queries", True
Application.SetOption "Default Record Locking", 2
Application.SetOption "Default Open Mode for Databases", 0
Application.SetOption "Number of Update Retries", 10

' condicionamento do ambiente 2
SS = LProgress(SS) '7
Application.SetOption "ODBC Refresh Interval (Sec)", Val(LConfig("Tempo_Atualização_ODBC", , 1))
Application.SetOption "Refresh Interval (Sec)", 60
Application.SetOption "Update Retry Interval (Msec)", 250
Application.SetOption "Show Status Bar", True
Application.SetOption "Show System Objects", False
Application.SetOption "Show Hidden Objects", False

' condicionamento do ambiente 3
SS = LProgress(SS) '8
LProperty CurrentDb, "StartUpShowDBWindow", False
LProperty CurrentDb, "StartUpShowStatusBar", True
LProperty CurrentDb, "AllowToolbarChanges", False

' abre form LInícioApl para configuração de inicialização local
If LExists(CurrentDb.Containers("Forms").Documents, "LInícioApl") Then
    DoCmd.OpenForm "LInícioApl", acNormal, , , , acHidden
End If

fim:

' finaliza linício
LInícioSai:
LProgress Null, , , , "LCurtaMensagem"
Exit Function

LInícioErro:
Dim xerr
xerr = LErro("LInício")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LInícioSai
End Function

Function LConvNull(Valor As Variant, Conv As Variant) As Variant
On Error Resume Next
LConvNull = Conv
LConvNull = IIf(IsNull(Valor), Conv, Valor)
End Function


'LÊ O CONTEÚDO DE UM ARQUIVO
Function LLerArquivo(ByVal Arquivo As String) As Variant
Dim Reg As String * 1000
Dim Txt As String
Dim Z As Long, Tam As Long
Dim TamBuffer As Integer
TamBuffer = 1000
'On Error GoTo LLerArquivo_Erro

Txt = ""
Open Arquivo For Binary As #1 Len = TamBuffer
For Z = 1 To (LOF(1) / TamBuffer) + 1
    If Len(Txt) = LOF(1) Then
        Exit For
    End If
    Get #1, Z, Reg
    Tam = LOF(1) - (Z - 1) * TamBuffer
    Txt = Txt & Left(Reg, IIf(Tam > TamBuffer, TamBuffer, Tam))
Next

LLerArquivo = Txt
Close #1

LLerArquivo_Sai:
Exit Function

LLerArquivo_Erro:
Dim xerr As Long
xerr = LErro("LLerArquivo")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If

Resume LLerArquivo_Sai
End Function

' ABRE UM FORMULÁRIO COMO UMA JANELA DE DIÁLOGO
Function LDiálogo(Optional ByVal Formulário As String, Optional Conteúd As String)
On Error GoTo LDiálogoErro
Dim RespForm As String
Static RESP As String

If Conteúd = "" Then
    RESP = LInsere(RESP, Formulário, "")
    DoCmd.OpenForm Formulário
    Do While LExists(Forms, Formulário)
        If Forms(Formulário).Visible = True Then
            DoEvents
        Else
            Exit Do
        End If
    Loop
    RespForm = LExtrai(RESP, Formulário)
    If RespForm <> "" Then
        LDiálogo = RespForm
    Else
        If LExists(Forms, Formulário) Then
            LDiálogo = True
        End If
    End If
Else
    If Formulário = "" Then
        If Application.CurrentObjectType = A_FORM Then
             Formulário = Application.CurrentObjectName
        End If
    End If
    RESP = LInsere(RESP, Formulário, Conteúd)
    DoCmd.Close acForm, Formulário, acSaveYes
End If
    
LDiálogoSai:
Exit Function

LDiálogoErro:
Dim xerr
xerr = LErro("LDiálogo")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LDiálogoSai
End Function



' VERIFICA SE UM OBJETO ESPECÍFICO EXISTE EM UMA COLEÇÃO
Function LExists(Colec As Object, OBJ As String)
On Error Resume Next
Dim oo As Object
LExists = False
Set oo = Colec(OBJ)
If Err = 0 Then
    LExists = True
End If
End Function


' ABRE UM FORMULÁRIO
Function LAbreForm(Nome As String, Optional Exibir As Variant = acNormal, Optional NomeFiltro As Variant = Null, Optional Condição As Variant = Null, Optional ModoDados As Variant = acFormPropertySettings, Optional ModoJanela As Variant = acWindowNormal) As Integer
On Error GoTo LAbreFormErro
 
'verifica se licença está válida
If LCheck_(LConfig("FABRICANTE"), LConfig("APLICAÇÃO"), LConfig("VERSÃO"), Application.hWndAccessApp) <> 0 Then
    LErro "LInício", "{OK}Este programa necessita de liberação. Contacte suporte autorizado."
    Exit Function
End If
 
DoCmd.OpenForm Nome, Exibir, NomeFiltro, Condição, ModoDados, ModoJanela
 
LAbreFormSai:
Exit Function
 
LAbreFormErro:
Dim xerr As Integer
xerr = LErro("LAbreForm")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LAbreFormSai
End Function


' FECHA UM FORMULÁRIO OU O CORRENTE CASO NADA SEJA INFORMADO
Function LFechaForm(stFrm As String) As Integer
On Error GoTo LFechaFormErro
If stFrm = "" Then
     If Application.CurrentObjectType = A_FORM Then
         On Error Resume Next
         DoCmd.Close A_FORM, Application.CurrentObjectName, acSaveYes
         On Error GoTo LFechaFormErro
     End If
Else
    If (LFormCarregado(stFrm)) Then
        On Error Resume Next
        DoCmd.Close A_FORM, stFrm, acSaveYes
        On Error GoTo LFechaFormErro
    End If
End If

LFechaFormSai:
Exit Function

LFechaFormErro:
Dim xerr
xerr = LErro("LFechaForm")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LFechaFormSai
End Function

Function LAbreRel(Nome As Variant, Optional Modo = acViewPreview, Optional NomeFiltro As Variant = Null, Optional Condicao As Variant = Null)
On Error GoTo LAbreRelErro
'verifica se licença está válida
If LCheck_(LConfig("FABRICANTE"), LConfig("APLICAÇÃO"), LConfig("VERSÃO"), Application.hWndAccessApp) <> 0 Then
    LErro "LInício", "{OK}Este programa necessita de liberação. Contacte suporte autorizado."
    Exit Function
End If

DoCmd.OpenReport Nome, Modo, NomeFiltro, Condicao

LAbreRelSai:
Exit Function

LAbreRelInexist:
    LErro "LAbreRel - " & Nome, "Recurso não implementado nesta versão. Entrar em contato com lucianoicraft@gmail.com e solicitar sua atualização."
Resume LAbreRelSai

LAbreRelErro:
Dim xerr As Integer
If Err = 2102 Then Resume LAbreRelInexist
If Err = 2501 Then Resume LAbreRelSai
xerr = LErro("LAbreRel")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LAbreRelSai
End Function

' VERIFICA SE FORMULÁRIO JÁ ESTÁ CARREGADO
Function LFormCarregado(Formula As Variant) As Integer
On Error Resume Next
LFormCarregado = False
LFormCarregado = Not IsNull(Forms(Formula).Name)
End Function

' VERIFICA SE RELATÓRIO JÁ ESTÁ CARREGADO
Function LRelCarregado(Rel As Variant) As Integer
On Error Resume Next
LRelCarregado = False
LRelCarregado = Not IsNull(Reports(Rel).Name)
End Function

' VERIFICA SE FORMULÁRIO OU RELATÓRIO JÁ ESTÁ CARREGADO
Function LCarregado(Nome As Variant) As Integer
On Error Resume Next
LCarregado = False
LCarregado = Not IsNull(Forms(Nome).Name)
If Not LCarregado Then
    LCarregado = Not IsNull(Reports(Nome).Name)
End If
End Function




' ESCONDE UM FORMULÁRIO OU FORMULÁRIO CORRENTE CASO NADA SEJA INFORMADO
Function LEscondeForm(Formula As String)
On Error GoTo LEscondeFormErro
If Formula = "" Then
     If Application.CurrentObjectType = A_FORM Then
         Forms(Application.CurrentObjectName).Visible = False
     End If
Else
    If LFormCarregado(Formula) Then
        Forms(Formula).Visible = False
    End If
End If

LEscondeFormSai:
Dim xerr
Exit Function
LEscondeFormErro:
xerr = LErro("LEscondeForm")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LEscondeFormSai
End Function



' APRESENTA UMA MENSAGEM NA TELA COM OU SEM BOTÃO PARA CONTINUAR
Sub LCurtaMensagem(Optional Texto As Variant = "", Optional NomeFont As Variant, Optional Font As Variant, Optional IndProgress, Optional TotProgress)
On Error GoTo LCurtaMensagemErro
LMensagem Texto, NomeFont, Font, IndProgress, TotProgress, "LCurtaMensagem"

LCurtaMensagemSai:
Exit Sub

LCurtaMensagemErro:
Dim xerr
xerr = LErro("LCurtaMensagem")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LCurtaMensagemSai
End Sub




' APRESENTA MENSAGEM COM OU SEM BOTÃO PARA UM TEXTO MAIOR
Function LMensagem(Optional Texto As Variant = "", Optional NomeFont As Variant, Optional Font As Variant, Optional IndProgress, Optional TotProgress, Optional FormName As String = "LMensagem")
Dim restrit As Boolean, bot As String

On Error GoTo LMensagemErro

restrit = False
LMensagemSaída = ""

If IsNull(Texto) Then
    If LExists(Forms, FormName) Then
        If LCarregado(FormName) Then
            DoCmd.Close acForm, FormName
        End If
    End If
    Exit Function
Else
    If Not LCarregado(FormName) Then
        DoCmd.OpenForm FormName
        Forms(FormName).Caption = LConfig("Aplicação")
    Else
        DoCmd.OpenForm FormName
    End If
        
    If Not IsMissing(NomeFont) Then
        If Not IsNull(NomeFont) Then
            Forms(FormName)!Texto.FontName = NomeFont
        End If
    End If
    If Not IsMissing(Font) Then
        If Not IsNull(Font) Then
            Forms(FormName)!Texto.FontSize = Font
        End If
    End If
    
    If Texto <> "" Then
        ' limpa texto e busca botões
        Dim pos As Long, posfim As Long, botões As String
        botões = ""
        pos = 1
        Do While pos <= Len(Texto)
            pos = InStr(pos, Texto, "{")
            If pos = 0 Then
                pos = Len(Texto) + 1
            Else
                posfim = InStr(pos + 1, Texto, "}")
                If posfim = 0 Then
                    posfim = Len(Texto) + 1
                End If
                bot = Mid(Texto, pos + 1, posfim - pos - 1)
                If bot = "MODAL" Then
                    restrit = True
                Else
                    botões = botões & Left(bot & String(20, " "), 20)
                End If
                Texto = Left(Texto, pos - 1) & Mid(Texto, posfim + 1)
            End If
        Loop
        
        ' apresenta botões
        If botões <> "" Then
            Dim tamv As Long, iniv As Long, Z As Integer, ctl As String
            tamv = Forms(FormName)("1").Width + Forms(FormName)("1").Width * 0.1
            iniv = (Forms(FormName).Width - tamv * Len(botões) / 20) / 2
            For Z = 1 To Len(botões) Step 20
                ctl = (Z + 19) / 20 & ""
                Forms(FormName)(ctl).Caption = "&" & Trim(Mid(botões, Z, 20))
                Forms(FormName)(ctl).Left = iniv
                iniv = iniv + tamv
                Forms(FormName)(ctl).Visible = True
            Next
        End If
            
        ' apresenta texto
        Forms(FormName)!Texto = Texto
        Forms(FormName).Repaint
    End If
End If

' apresenta progresso
If Not IsMissing(TotProgress) Then
    Forms(FormName).BProg.min = 0
    Forms(FormName).BProg.MAX = TotProgress
End If
If Not IsMissing(IndProgress) Then
    If IndProgress <= Forms(FormName).BProg.MAX Then
        Forms(FormName).BProg = IndProgress
    Else
        Forms(FormName).BProg = Forms(FormName).BProg.MAX
    End If
End If

' torna visível ou não o progresso
If Not IsMissing(TotProgress) Or Not IsMissing(IndProgress) Then
    Forms(FormName).BProg.Visible = True
    Forms(FormName).TProg.Visible = True
Else
    Forms(FormName).BProg.Visible = False
    Forms(FormName).TProg.Visible = False
End If

' bloqueia se for solicitada janela modal
If restrit Then
    Do While LMensagemSaída = ""
        DoEvents
    Loop
    LMensagem = LMensagemSaída
End If

LMensagemSai:
Exit Function

LMensagemErro:
Dim xerr
xerr = LErro("LMensagem")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LMensagemSai
End Function



' CRIA, DEFINE OU RETORNA UM CONTEÚDO DE UMA PROPRIEDADE DE UM OBJETO
Function LProperty(OBJ As Object, Prop As String, Optional Conteúd)
On Error Resume Next
Dim Prp As Property
If VarType(Conteúd) = 10 Then
    LProperty = OBJ.Properties(Prop)
    If Err <> 0 Then
        On Error GoTo LPropertyErro
        LProperty = Null
    End If
Else
    If IsNull(Conteúd) Then
        OBJ.Properties.Delete Prop
        LProperty = True
    Else
        On Error Resume Next
        OBJ.Properties(Prop) = Conteúd
        If Err <> 0 Then
            On Error GoTo LPropertyErro
            Set Prp = OBJ.CreateProperty(Prop, dbText, Conteúd)
            OBJ.Properties.Append Prp
            OBJ.Properties.Refresh
        End If
        On Error GoTo LPropertyErro
        LProperty = True
    End If
End If
LPropertySai:
Exit Function

LPropertyErro:
Dim xerr As Integer
xerr = LErro("LProperty")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LPropertySai
End Function



' RECONSULTA UMA LISTA DE CAMPOS OU O CAMPO ATUAL CASO NADA SEJA INFORMADO
Function LRequery(Optional CAMPOS As String)
On Error Resume Next
Dim Z As Integer, ITEM As String, FF As Form

If CAMPOS = "" Then
    Screen.ActiveControl.Requery
Else
    Set FF = Screen.ActiveControl.Parent
    For Z = 1 To LItem(CAMPOS, 0)
        ITEM = LItem(CAMPOS, Z)
        FF(ITEM).Requery
    Next
End If
End Function



' DESLOCA O FORMULÁRIO PARA UMA ETIQUETA ESPECÍFICA
Function LGotoEtiq(Nome As String)
On Error GoTo LGotoEtiqErro
Dim FF As String
If Application.CurrentObjectType = acForm Then
    FF = Application.CurrentObjectName
    Forms(FF).GoToPage 1, 0, Forms(FF)(Nome).Properties!Top - 94
End If
LGotoEtiqSai:
Exit Function

LGotoEtiqErro:
Dim xerr As Integer
xerr = LErro("LGoToEtiq")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LGotoEtiqSai
End Function




' CONCATENA O RESULTADO DE RECORDSET EM UMA SEQUENCIA DE CARACTERES DELIMITADA
Function LConcatSQL(DB As Database, SQL As String, Optional Delim As String = ";")
On Error GoTo LConcatSQLErro
Dim REC As DAO.Recordset, Ret As String, Palav As String

Ret = ""
Set REC = DB.OpenRecordset(SQL)
If REC.RecordCount > 0 Then
    REC.MoveFirst
    Do While Not REC.EOF
        Ret = Ret & IIf(Ret <> "", Delim, "") & REC(0)
        REC.MoveNext
    Loop
End If
REC.Close

LConcatSQL = Ret

LConcatSQLSai:
Exit Function

LConcatSQLErro:
Dim xerr As Integer
xerr = LErro("LConcatSQL")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LConcatSQLSai
End Function



' CONCATENA O RESULTADO DE RECORDSET EM UMA SEQUENCIA DE CARACTERES DELIMITADA
Function LConcatCamp(DOMÍNIO As String, Delimit As String) As String
On Error GoTo LConcatCampErro
Dim DB As Database, REC As Recordset, Ret As String, Palav As String, x As Variant
Static Reference As Field, NumIt As Integer

Ret = ""
SetBanco:
Set DB = CurrentDb
Set REC = DB.OpenRecordset(DOMÍNIO)
If REC.RecordCount > 0 Then
    REC.MoveFirst
    Do While Not REC.EOF
        Ret = Ret & IIf(Ret <> "", Delimit, "") & REC(0)
        REC.MoveNext
    Loop
End If
REC.Close

LConcatCamp = Ret

LConcatCampSai:
Exit Function

LConcatCampErro:
Dim xerr
xerr = LErro("LConcatCamp")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LConcatCampSai
End Function



' MOSTRA SOMENTE AS BARRAS DEFINIDAS PELO PARÂMETRO SEPARADO POR ";"
Sub LBarras(Barras As String)
On Error Resume Next
Dim Z As Integer
For Z = 1 To Application.CommandBars.Count
    If Application.CommandBars(Z).Protection = 0 Or Application.CommandBars(Z).Protection = 8 Then
        If LItem(Barras, Application.CommandBars(Z).Name) = 0 Then
            Application.CommandBars(Z).Enabled = False
        Else
            Application.CommandBars(Z).Enabled = True
            Application.CommandBars(Z).Visible = True
        End If
    End If
Next
End Sub



' DEFINE UM PARÂMETRO DE UM FORMULÁRIO
Function LFormParam(Formula As String, Param As String, Conteúdo As String)
On Error GoTo LFormParamErro

DoCmd.OpenForm Formula
Forms(Formula)(Param) = Conteúdo
If Forms(Formula)(Param).Enabled And Forms(Formula)(Param).Visible Then
    Forms(Formula)(Param).SetFocus
End If

LFormParamSai:
Exit Function

LFormParamErro:
Dim xerr As Integer
xerr = LErro("LFormParam")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LFormParamSai
End Function



' INSERE NULO EM UM CAMPO
Function LInsereNull(Optional ctl As Control)
On Error Resume Next
ctl = Null
If Err <> 0 Then
    If Application.CurrentObjectType = A_FORM Then
        Forms(Application.CurrentObjectName).SetFocus
        Screen.ActiveControl = Null
    End If
End If
End Function




' PERMITE INDICAÇÃO DE UM CAMINHO ODBC
Function LConnect(Optional CONN As String = "ODBC;")
On Error GoTo LConnectErro
LConnect = ""
Dim DB As Database
Set DB = DBEngine(0).OpenDatabase("", False, False, CONN)
LConnect = DB.Connect
DB.Close

LConnectSai:
Exit Function

LConnectErro:
Dim xerr As Integer
If Err <> 3059 Then
    xerr = LErro("LConnect")
    If xerr = 4 Then
        Resume 0
    ElseIf xerr = 5 Then
        Resume Next
    End If
End If
Resume LConnectSai
End Function



' CONFIGURA OU RECUPERA UM PARÂMETRO GLOBAL DO APLICATIVO : TABELA SYS_CONFIG_GLOBAL
Function LSistema(Campo As String, Optional Valor, Optional DEF)
On Error GoTo LSistemaErro
Dim Conteúdo As Variant, REC As DAO.Recordset
If VarType(Valor) = 10 Then
    LSistema = IIf(IsMissing(DEF), "[" & Campo & "]", DEF)
    On Error Resume Next
    Err.Clear
    Conteúdo = DLookup("CONFIG", "SYS_CONFIG_GLOBAL", "[PARAM] = '" & Campo & "'")
    If Err = 0 Then
        If Not IsNull(Conteúdo) Then
            LSistema = Conteúdo
        End If
    End If
Else
    LSistema = 0
    If VarType(Valor) = vbNull Then
        CurrentDb.Execute "DELETE * FROM SYS_CONFIG_GLOBAL WHERE PARAM = '" & Campo & "';"
    Else
        Set REC = CurrentDb.OpenRecordset("SELECT * FROM SYS_CONFIG_GLOBAL WHERE PARAM = '" & Campo & "';")
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

LSistemaSai:
Exit Function

LSistemaErro:
Dim xerr As Integer
xerr = LErro("LSistema")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LSistemaSai
End Function

' SE CERTIFICA QUE O USUÁRIO ESTÁ ADEQUADAMENTE CADASTRADO ALÉM DE DEFINIR OU RETORNAR CONTEÚDO DE UM PARÂMETRO
Function LConfigUsuário(Optional Campo As String, Optional Valor)
Dim Usu As String, REC As Recordset

LConfigUsuárioNovo:
On Error GoTo LConfigUsuárioErro
Set REC = CurrentDb.OpenRecordset("SELECT * FROM SYS_USUÁRIO WHERE USUÁRIO = '" & UCase(Application.CurrentUser) & "';", dbOpenDynaset)

If REC.RecordCount = 0 Then
    REC.AddNew
    REC!Usuário = UCase(CurrentUser())
    REC!Nome = LCorrigeNome(CurrentUser())
    REC.Update
    GoTo LConfigUsuárioNovo
End If

If Campo <> "" Then
    If VarType(Valor) <> 10 Then
        REC.Edit
        REC(Campo) = Valor
        REC.Update
    Else
        LConfigUsuário = REC(Campo)
    End If
End If

LconfigUsuárioSai:
Exit Function

LConfigUsuárioErro:
Dim xerr
xerr = LErro("LConfigUsuário")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LconfigUsuárioSai
End Function

' REALIZA BACKUP DO SISTEMA
Function LBackup()
Dim BaseDados As String
Dim AccessApp As String * 255

If LCheck_(LConfig("FABRICANTE"), LConfig("APLICAÇÃO"), LConfig("VERSÃO"), Application.hWndAccessApp) <> 0 Then
    LErro "LInício", "{OK}Este programa necessita de liberação. Contacte suporte autorizado."
    Exit Function
End If

Dim TamAccessApp As Long

BaseDados = LConfig("ConnBaseDados")

If BaseDados = "Microsoft Access" Then
    LBackup = LBackupAccess
    If LBackup = True Then
        LConfig "LBackup_Data", Format(Now, "dd/mm/yyyy")
    End If
ElseIf BaseDados = "Firebird" Then
    If LConfig("ConnDSN", , LConfig("ConnDSNPadrão", , "")) = "" Then
        LErro "LBackup", "{OK}Parâmetro de Fonte de Dados não foi encontrado. Execute o procedimento de configuração de base de dados para realizar o backup do sistema" & vbCrLf & vbCrLf & "Caso este problema persista, contate o suporte técnico Intercraft."
        Exit Function
    End If
    If MsgBox("O aplicativo deverá ser fechado para realização do backup. Deseja continuar?", vbYesNo Or vbQuestion, LConfig("Aplicação")) = vbNo Then
        Exit Function
    End If
    TamAccessApp = LGetModuleFileName(0&, AccessApp, 255)
    AccessApp = Left(AccessApp, TamAccessApp)
    LExecShell "icftbackup.exe " & Application.hWndAccessApp & " -b """ & Trim(AccessApp) & """ """ & CurrentDb.Name & """ """ & "DSN=" & LConfig("ConnDSN", , LConfig("ConnDSNPadrão")) & """", 0
    LConfig "LBackup_Data", Format(Now, "dd/mm/yyyy")
    Application.Quit acQuitSaveNone
Else
    LErro "LBackup", "{OK}Recurso ainda não implementado para o banco de dados " & LConfig("ConnBaseDados") & "! Será implementado nas próximas versões"
    LBackup = 0
End If
End Function


' COMPLEMENTO DO LBACKUP SE O BANCO DE DADOS FOR ACCESS
Function LBackupAccess()
Dim DBOrig As Database
Dim DBDest As Database
Dim NewDir As String
Dim NewDb As String
Dim CurDb As String
Dim Usu As String
Dim Pass As String
Dim Tbl As TableDef
Dim Qry As QueryDef
Dim Rel As Relation
Dim NewRel As Relation
Dim fld As Field
Dim NewFld As Field

CurDb = LConfig("ConnArquivo", , "")

If CurDb = "" Then
    LErro "LBackup", "{OK}O atributo Arquivo não foi encontrado! Execute o procedimento de configuração de base de dados para realizar o backup do sistema" & vbCrLf & vbCrLf & "Caso este problema persista, contate o suporte técnico Intercraft."
    LBackupAccess = False
    Exit Function
End If

NewDir = Left(CurrentDb.Name, InStrRev(CurrentDb.Name, "\")) & "Backup"

If dir(NewDir, vbDirectory) = "" Then
    MkDir NewDir
End If

NewDb = NewDir & Mid(CurDb, InStrRev(CurDb, "\"))

If dir(NewDb) <> "" Then
    Kill NewDb
End If

On Error GoTo LBackupAccess_Erro
With New Access.Application
    If LLogin(Usu, Pass) Then
        Set DBOrig = .DBEngine.OpenDatabase(CurDb, True, False, ";UID=" & Usu & ";PWD=" & Pass)
        .NewCurrentDatabase NewDb
        .CurrentDb.NewPassword "", Pass
        Usu = ""
        Pass = ""
        
        For Each Tbl In DBOrig.TableDefs
            If Not Tbl.Name Like "MSys*" Then
                .DoCmd.TransferDatabase acImport, "Microsoft Access", CurDb, acTable, Tbl.Name, Tbl.Name, False, True
            End If
        Next

        For Each Rel In DBOrig.Relations
            Set NewRel = .CurrentDb.CreateRelation(Rel.Name, Rel.Table, Rel.ForeignTable, Rel.Attributes)
            For Each fld In Rel.Fields
                Set NewFld = NewRel.CreateField(fld.Name)
                NewFld.ForeignName = fld.ForeignName
                NewRel.Fields.Append NewFld
            Next
            .CurrentDb.Relations.Append NewRel
        Next
        
        For Each Qry In DBOrig.QueryDefs
            .DoCmd.TransferDatabase acImport, "Microsoft Access", CurDb, acQuery, Qry.Name, Qry.Name, False, True
        Next
        
        .CloseCurrentDatabase
    End If
End With

LBackupAccess = True
LBackupAccess_Sai:
Exit Function

LBackupAccess_Erro:
LBackupAccess = False
Dim xerr As Integer
xerr = LErro(Err)

If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LBackupAccess_Sai
End Function

' REALIZA RESTAURAÇÃO DO SISTEMA
Function LRestore()
Dim BaseDados As String
Dim AccessApp As String * 255
Dim TamAccessApp As Integer

If LCheck_(LConfig("FABRICANTE"), LConfig("APLICAÇÃO"), LConfig("VERSÃO"), Application.hWndAccessApp) <> 0 Then
    LErro "LInício", "{OK}Este programa necessita de liberação. Contacte suporte autorizado."
    Exit Function
End If

BaseDados = LConfig("ConnBaseDados")

If BaseDados = "Microsoft Access" Then
    LRestore = LRestoreAccess
ElseIf BaseDados = "Firebird" Then
    If LConfig("ConnDSN", , LConfig("ConnDSNPadrão", , "")) = "" Then
        LErro "LRestore", "{OK}Parâmetro de Fonte de Dados não foi encontrado. Execute o procedimento de configuração de base de dados para realizar a restauração do sistema" & vbCrLf & vbCrLf & "Caso este problema persista, contate o suporte técnico Intercraft."
        Exit Function
    End If
    If MsgBox("O aplicativo deverá ser fechado para a restauração. Deseja continuar?", vbYesNo Or vbQuestion, LConfig("Aplicação")) = vbNo Then
        Exit Function
    End If
    TamAccessApp = LGetModuleFileName(0&, AccessApp, 255)
    AccessApp = Left(AccessApp, TamAccessApp)
    LExecShell "icftbackup.exe " & Application.hWndAccessApp & " -r """ & Trim(AccessApp) & """ """ & CurrentDb.Name & """ """ & "DSN=" & LConfig("ConnDSN", , LConfig("ConnDSNPadrão", , "")) & """", 0
    Application.Quit acQuitSaveNone
Else
    LErro "LRestore", "{OK}Recurso ainda não implementado para o banco de dados " & LConfig("ConnBaseDados") & "! Será implementado nas próximas versões"
    LRestore = 0
End If
End Function

Function LRestoreAccess()
Dim CurDb As String
Dim NewDir As String
Dim NewDb As String
Dim Usu As String
Dim Pass As String
Dim DB As Database
Dim Tbl As TableDef
Dim Qry As QueryDef
Dim Rel As Relation
Dim NewRel As Relation
Dim fld As Field
Dim NewFld As Field


CurDb = LConfig("ConnArquivo", , "")

If CurDb = "" Then
    LErro "LRestore", "{OK}O atributo Arquivo não foi encontrado! Execute o procedimento de configuração de base de dados para realizar o backup do sistema" & vbCrLf & vbCrLf & "Caso este problema persista, contate o suporte técnico Intercraft."
    LRestoreAccess = False
    Exit Function
End If

NewDir = Left(CurrentDb.Name, InStrRev(CurrentDb.Name, "\")) & "Backup"
NewDb = NewDir & Mid(CurDb, InStrRev(CurDb, "\"))

If dir(NewDb) = "" Then
    LErro "LRestore", "{OK}O arquivo de backup da base de dados não foi encontrado." & Chr(13) & Chr(10) & "Realize o backup do banco de dados antes de restaurá-lo."
    Exit Function
End If

On Error GoTo LRestoreAccess_Erro

With New Access.Application
    If LLogin(Usu, Pass) Then
        Set DB = .DBEngine.OpenDatabase(CurDb, True, False, ";UID=" & Usu & ";PWD=" & Pass)
        .OpenCurrentDatabase CurDb
        DB.Close
        Set DB = .DBEngine.OpenDatabase(NewDb, True, False, ";UID=" & Usu & ";PWD=" & Pass)
        Usu = ""
        Pass = ""
        
        For Each Rel In .CurrentDb.Relations
            .CurrentDb.Relations.Delete Rel.Name
        Next
        
        For Each Tbl In .CurrentDb.TableDefs
            If Not Tbl.Name Like "MSys*" Then
                .CurrentDb.TableDefs.Delete Tbl.Name
            End If
        Next
        
        For Each Tbl In DB.TableDefs
            If Not Tbl.Name Like "MSys*" Then
                .DoCmd.TransferDatabase acImport, "Microsoft Access", NewDb, acTable, Tbl.Name, Tbl.Name, False
            End If
        Next
        
        For Each Rel In DB.Relations
            Set NewRel = .CurrentDb.CreateRelation(Rel.Name, Rel.Table, Rel.ForeignTable, Rel.Attributes)
            For Each fld In Rel.Fields
                Set NewFld = NewRel.CreateField(fld.Name)
                NewFld.ForeignName = fld.ForeignName
                NewRel.Fields.Append NewFld
            Next
            .CurrentDb.Relations.Append NewRel
        Next
        
        For Each Qry In DB.QueryDefs
            .DoCmd.TransferDatabase acImport, "Microsoft Access", NewDb, acQuery, Qry.Name, Qry.Name, False, True
        Next
        
        DB.Close
        .CloseCurrentDatabase
    End If
End With

LRestoreAccess = True
LRestoreAccess_Sai:
Exit Function

LRestoreAccess_Erro:
LRestoreAccess = False
Dim xerr As Integer
xerr = LErro(Err)

If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LRestoreAccess_Sai
End Function

' APRESENTA UM DIÁLOGO QUE PERMITE SELECIONAR UM ARQUIVO
Function LDlgOpenFile(Optional Janela As Long = 0, Optional TÍTULO As String = "Abrir", Optional Estilo As String = "Carregar", Optional Filtro As String = "Todos os arquivos (*.*)|*.*|") As String
Dim Arq As String * 1000
Dim pos As Integer
Dim EstiloLng As Long

Arq = ""
If Janela = 0 Then
    Janela = Application.hWndAccessApp
End If

Filtro = Replace(Filtro, "|", Chr(0))

Select Case (Estilo)
    Case Is = "Carregar"
        EstiloLng = LLIB_OFN_CARREGAR_ARQ
    Case Is = "Carregar Múltiplos"
        EstiloLng = LLIB_OFN_CARREGAR_ARQ_MULT
    Case Is = "Salvar"
        EstiloLng = LLIB_OFN_SALVAR_ARQ
    Case Is = "Carregar Diretório"
        EstiloLng = LLIB_OFN_CARREGAR_DIR
    Case Else
        MsgBox "LDlgOpenFile - Estilo de Abertura de Janela não suportado!", vbCritical, LConfig("Aplicação")
        GoTo LDlgOpenFileSai
End Select

If LAbreArquivo(Janela, TÍTULO, EstiloLng, Filtro, Arq) = 0 Then
    GoTo LDlgOpenFileSai
End If

LDlgOpenFile = Left(Arq, InStr(1, Arq, Chr(0)) - 1)
On Error GoTo LDlgOpenFileErr

LDlgOpenFileSai:
Exit Function

LDlgOpenFileErr:
Dim xerr As Integer
xerr = LErro("LDlgOpenFile")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LDlgOpenFileSai
End Function

Function LItemArquivo(Arquivo As String, Optional Numero As Variant) As Variant
Dim pos, cont As Integer
Dim dir, Arq As String
cont = 0
pos = InStr(1, Arquivo, "\ ")
If pos = 0 Then
    Dim pos2, Temp As Integer
    Temp = 1
    While Not Temp = 0
        pos2 = Temp
        Temp = InStr(pos2 + 1, Arquivo, "\")
    Wend
    dir = Left(Arquivo, pos2)
    Arq = Mid(Arquivo, pos2 + 1, Len(Arquivo) - pos2)
    If IsMissing(Numero) Then
        LItemArquivo = 1
        Exit Function
    ElseIf Numero = 0 Then
        LItemArquivo = dir
        Exit Function
    Else
        LItemArquivo = Arq
        Exit Function
    End If
Else
    dir = Left(Arquivo, pos)
    Arq = Mid(Arquivo, pos + 2, Len(Arquivo) - pos + 2)
    If IsMissing(Numero) Then
        LItemArquivo = LItem(Arq, 0)
        Exit Function
    ElseIf Numero = 0 Then
        LItemArquivo = dir
        Exit Function
    Else
        LItemArquivo = LItem(Arq, Numero)
        Exit Function
    End If
End If
End Function

Function LListaAbstrCarac(Licença As String, DataLicença As String, Optional Seq As String = "")
On Error GoTo LListaAbstrCaracErr
Dim BUF As String * 200, Txt As String, NUM As Integer, RESULT As String

If Seq <> "" Then
    Txt = Licença & DataLicença & Seq
    LAbstr_ Txt
    RESULT = Left(Txt, 10)
Else
    Seq = "00"
    Do While Val(Seq) <= 99
        Txt = Licença & DataLicença & Seq
        LAbstr_ Txt
        RESULT = RESULT & IIf(RESULT <> "", Chr(13) & Chr(10), "") & Seq & "->" & Left(Txt, 10)
        Seq = Format(Val(Seq) + 1, "00")
    Loop
End If
LListaAbstrCarac = RESULT

LListaAbstrCaracSai:
Exit Function

LListaAbstrCaracErr:
Dim xerr
xerr = LErro("LListaAbstrCarac")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LListaAbstrCaracSai
End Function


Function LSQL(Param)
On Error GoTo LSqlErro
Dim Elem As Variant, Ret As Variant
If IsArray(Param) Then
    For Each Elem In Param
        Ret = Ret & IIf(Ret <> "", ",", "") & LExprTipoVar(Elem)
    Next
    LSQL = Ret
Else
    LSQL = LExprTipoVar(Param)
End If

LSqlSai:
Exit Function

LSqlErro:
Dim xerr
xerr = LErro("LSql")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LSqlSai

End Function

' PERMITE APRESENTAÇÃO DE UM INDICADOR DE PROGRESSO
Function LProgress(Optional Ind = Null, Optional Mens = Null, Optional TOT = Null, Optional Tela = True, Optional FormName As String)
On Error GoTo LProgressErro

Static AntTot As Long, AntInd As Long, AntMens As String, AntFormName As String
Dim Cancela As Boolean, Z As Long

' obtém parâmetros caso já tenham sido definidos previamente ou inicializa-os
If IsNull(Mens) Then
    Mens = IIf(AntMens <> "", AntMens, "Progresso")
End If
If IsNull(TOT) Then
    TOT = IIf(AntTot <> 0, AntTot, 100)
End If
If FormName = "" Then
    FormName = IIf(AntFormName <> "", AntFormName, "LMensagem")
End If

' verifica se é para cancelar
If IsMissing(Ind) Then
    Cancela = True
Else
    If IsNull(Ind) Then
        Cancela = True
    End If
End If

' cancela progresso
If Cancela Then
    If AntInd < AntTot Then
        SysCmd acSysCmdInitMeter, AntMens, AntTot
        If Tela Then
            LMensagem Mens, , , AntInd, AntTot, FormName
        End If
        For Z = AntInd To AntTot
            SysCmd acSysCmdUpdateMeter, Z
            If Tela Then
                LMensagem , , , Z, , FormName
            End If
        Next
    End If
    SysCmd acSysCmdRemoveMeter
    If Tela Then
        LMensagem Null, , , , , FormName
    End If
    
    ' inicializa variáveis anteriores
    AntInd = 0
    AntTot = 0
    AntMens = ""
    AntFormName = ""
    Exit Function
End If

' atualiza progresso
If Not IsMissing(Mens) Or Not IsMissing(TOT) Then
    SysCmd acSysCmdInitMeter, Mens, TOT
    LMensagem Mens, , , , TOT, FormName
End If
If Not IsMissing(Ind) Then
    SysCmd acSysCmdUpdateMeter, Ind
    LMensagem , , , Ind, , FormName
    LProgress = Ind + 1
End If

AntInd = Ind
AntTot = TOT
AntMens = Mens
AntFormName = FormName

LprogressSai:
Exit Function

LProgressErro:
Dim xerr
xerr = LErro("LProgress")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LprogressSai
End Function

'PAUSA A EXECUÇÃO DURANTE ALGUNS SEGUNDOS
Function LPausa(segs As Long)
Dim nn As Date
nn = Now()
Do While DateDiff("s", nn, Now()) < segs
    DoEvents
Loop
End Function

' CONFIGURA RESOLUÇÃO DE VÍDEO
Function LConfigVídeo(Optional Pergunta = 0)
On Error GoTo LConfigVídeoErro
Dim tipovídeo As String, pos As Long
Dim tamcol As Long, tamlin As Long, RT As Rect
Dim wincol As Long, winlin As Long
Dim Ret As String
Dim ccol As Long, llin As Long

If VarType(Pergunta) = vbInteger Or VarType(Pergunta) = vbBoolean Then
    ' mostra tela de pergunta sobre resolução
    If Pergunta Then
        Ret = LDiálogo("LConfigVídeo")
        If Ret = "" Then
            Exit Function
        Else
            tipovídeo = Ret
        End If
    Else
        tipovídeo = LConfig("Resolução", , "")
    End If
    
ElseIf VarType(Pergunta) = vbString Then
    ' interpreta parâmetro de pergunta como se fosse a própria resolução
    tipovídeo = Pergunta
End If

' se não tiver sido definida resolução, define 800x600
If tipovídeo = "" Then
    tipovídeo = "800x600"
Else
    tipovídeo = LCase(tipovídeo)
End If
LConfig "Resolução", tipovídeo

' descarrega tela prima
If LCarregado(LTelaPrima) Then
    Forms(LTelaPrima).OnUnload = ""
    tamcol = LDesencaixaTela(LGethWndMDIClient())
    DoCmd.Close A_FORM, LTelaPrima
End If
    
' converte em twip para pixel
pos = LGethWndDesktopClient()
LGetClientRect IIf(pos = 0, LGetDesktopWindow(), pos), RT
pos = InStr(tipovídeo, "x")
tamcol = Val(Left(tipovídeo, pos - 1))
tamlin = Val(Mid(tipovídeo, pos + 1))
wincol = RT.x2 - RT.x1
winlin = RT.y2 - RT.y1
If tamcol > wincol Then tamcol = wincol
If tamlin > winlin Then tamlin = winlin
ccol = (wincol - tamcol) / 2 - 8 + RT.x1
llin = (winlin - tamlin) / 2 - 8 + RT.y1
If ccol < -4 Then ccol = -4
If llin < -4 Then llin = -4

' define tamanho da janela do access
LShowWindow Application.hWndAccessApp, 9
LMoveWindow Application.hWndAccessApp, ccol, llin, tamcol + 6, tamlin + 8, True

' carrega tela prima adequada
Dim pict
LTelaPrima = "LTelaPrima"
DoCmd.OpenForm LTelaPrima, acNormal, , , acFormReadOnly, acWindowNormal

pict = LConfig("LTelaPrima" & tipovídeo, , "")
If pict <> "" Then
    Forms("LTelaPrima").PictureData = pict
    Forms("LTelaPrima").PictureTiling = IIf(LConfig("LTELAPRIMA_TELHA", , "N") = "S", True, False)
    Forms("LTelaPrima").PictureSizeMode = LConfig("LTELAPRIMA_MODO", , 0)
End If

LEncaixaTela LGethWndMDIClient(), Forms(LTelaPrima).hwnd

LconfigVídeoSai:
Exit Function

LConfigVídeoErro:
Dim xerr As Integer
xerr = LErro("LConfigVídeo")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LconfigVídeoSai
End Function

' CÓDIGO DE ABOUT DA APLICAÇÃO
Function LSobreSistema()
On Error Resume Next

DoCmd.OpenForm "LSobreSistema"
Forms("LSobreSistema")("ImgSobre").PictureData = LConfig("LSobreSistema")
End Function


' CÓDIGO QUE SEMPRE É EXECUTADO QUANDO SE SAI DO SISTEMA
Function LSaída(Optional Pergunta = True)
On Error Resume Next
If Pergunta Then
    If MsgBox("Deseja sair do sistema ?", vbQuestion + vbYesNo + vbDefaultButton2, LConfig("Aplicação")) <> vbYes Then
        LSaída = False
        Exit Function
    End If
End If

LProgress 0, "Aguarde que o sistema está encerrando suas atividades.", 100, , "LCurtaMensagem"
Forms(LTelaPrima).OnUnload = ""
LDesencaixaTela (LGethWndMDIClient())
DoCmd.Close acForm, LTelaPrima
LProgress Null, , , , "LCurtaMensagem"

Application.Quit IIf(Pergunta, acQuitSaveAll, acQuitSaveNone)
End Function


Function LLogin(ByRef Usuário As String, ByRef Senha As String) As Boolean
Dim REC As Recordset
DoCmd.OpenForm "LLogin"
While LCarregado("LLogin")
    DoEvents
Wend
If (Not LExists(CurrentDb.Properties, "Login")) And (Not LExists(CurrentDb.Properties, "Senha")) Then
    LLogin = False
Else
    Usuário = CurrentDb.Properties("Login").Value
    Senha = CurrentDb.Properties("Senha").Value
    CurrentDb.Properties.Delete "Login"
    CurrentDb.Properties.Delete "Senha"
    LLogin = True
End If
End Function

' PERMITE CONFIGURAR AS TABELAS VINCULADAS DO SISTEMA
Function LAttach(Pergunta As Boolean, Optional PREFIX As String)
DoCmd.OpenForm "LAttach"
Forms("LAttach").PREFIX = PREFIX
LDiálogo "LAttach"

If Not Pergunta Then
    LInício
End If
End Function

' CONFERE OU PERMITE CONFIGURAR AS TABELAS ATTACHADAS DO SISTEMA
'Function LAttach(Pergunta, Optional PREFIX As String)
'Dim Z As Integer, ZZ As Integer, DD As String, Usu As String
'On Error Resume Next
'DD = ""
'
'DD = CurrentDb.TableDefs("SYS_CONFIG_GLOBAL").Connect
'On Error GoTo LAttachErro
'Do While True
'    If Not Pergunta Then
'ConectNovamente:
'        On Error Resume Next
'        Err.Clear
'        Z = DCount("PARAM", "SYS_CONFIG_GLOBAL")
'        If Err = 0 Then
'            Z = DCount("USUÁRIO", "SYS_USUÁRIO")
'            If Err = 0 Then
'                If DD <> CurrentDb.TableDefs("SYS_CONFIG_GLOBAL").Connect Then
'                    LInício
'                End If
'                Exit Function
'            End If
'        End If
'        On Error GoTo LAttachErro
'        If MsgBox("Base de dados incorreta. Deseja configurar ?", vbQuestion + vbYesNo, LConfig("Aplicação")) = vbNo Then
'            MsgBox "Impossível continuar sem a base de dados configurada. O sistema foi abortado.", vbCritical + vbOKOnly, LConfig("Aplicação")
'            LSaída False
'        End If
'    End If
'    DoCmd.OpenForm "LATTACH"
'    Forms!LAttach.PREFIX = PREFIX
'    LDiálogo "LAttach"
'    Pergunta = False
'
'Loop
'LAttachSai:
'Exit Function
'
'LAttachErro:
'Dim xerr
'xerr = LErro("LAttach")
'If xerr = 4 Then
'    Resume 0
'ElseIf xerr = 5 Then
'    Resume Next
'End If
'Resume LAttachSai:
'End Function

'VINCULA TABELAS NA BASE ATUAL
Function LAttach_OK(Optional Janela As String, Optional PREFIX As String)
On Error GoTo lattach_OKErro
Dim Conexão As String, Esquema As String, Usuário As String, Servidor As String, Arquivo As String
Dim SS As Long, REC As Recordset, RS As Recordset, TBD As TableDef, Z As Integer, XX As String, QTDE As Integer, fld As Field, Prop As Property
Dim Ret As Long, hKey As Long
Dim RegValue As String, RegValueLen As Long
Dim SQL As String, i As Integer
Dim Qry As QueryDef

' inicia variáveis
Conexão = LConfig("ConnBaseDados", , "")
Servidor = LConfig("ConnDSN", , LConfig("ConnDSNPadrão", , ""))

If Conexão = "Oracle" Then
    Esquema = LConfig("ConnEsquema")
ElseIf Conexão = "Microsoft Access" Or Conexão = "Firebird" Then
    Arquivo = LConfig("ConnArquivo")
End If

Usuário = LConfig("ConnUsuário", , "")

' prepara janela caso esteja aberta
If Janela <> "" Then
    Forms(Janela)("TProg").Visible = True
    Forms(Janela)("BProg").Visible = True
End If

' carrega lista de tabelas a serem attachadas
SS = 0
LProgress SS, "Adquirindo lista de tabelas a serem anexadas", 1200, True, Janela
If LExists(CurrentDb.TableDefs, "SYS_USER_TABLES") Then
    DoCmd.DeleteObject acTable, "SYS_USER_TABLES"
End If
If LExists(CurrentDb.QueryDefs, "SYS_USER_TABLES") Then
    DoCmd.DeleteObject acQuery, "SYS_USER_TABLES"
End If
If Conexão = "Oracle" Then
    StringConnect = Forms("LAttach").PreencheODBCOracle
    SQL = "SELECT * FROM (SELECT OWNER, TABLE_NAME NOME, DECODE(OWNER,'" & Usuário & "',1,2) ORDEM, 'TB' TIPO FROM ALL_TABLES WHERE OWNER IN ('" & Esquema & "','" & Usuário & "') UNION "

    SQL = SQL & "SELECT OWNER, VIEW_NAME NOME, DECODE(OWNER,'" & Usuário & "',1,2) ORDEM, 'VW' TIPO FROM ALL_VIEWS WHERE OWNER IN ('" & Esquema & "','" & Usuário & "')) ORDER BY ORDEM, NOME"

    LQueryDef "SYS_USER_TABLES", SQL, StringConnect
    LQueryDef "SYS_USER_TABLES", SQL, StringConnect
    Set REC = CurrentDb.OpenRecordset("SELECT * FROM SYS_USER_TABLES")
ElseIf Conexão = "Firebird" Then
    StringConnect = "ODBC;DSN=" & Servidor & ";UID=" & Usuário & ";PWD=" & Forms("LAttach")("Senha") & ";DbName=" & Arquivo
    CurrentDb.TableDefs.Append CurrentDb.CreateTableDef("SYS_USER_TABLES", 0, "RDB$RELATIONS", StringConnect)
    CurrentDb.TableDefs.Append CurrentDb.CreateTableDef("SYS_USER_FIELDS", 0, "RDB$RELATION_FIELDS", StringConnect)
    Set REC = CurrentDb.OpenRecordset("SELECT * FROM SYS_USER_TABLES WHERE RDB$SYSTEM_FLAG = 0")
    Set RS = CurrentDb.OpenRecordset("SELECT * FROM SYS_USER_FIELDS WHERE RDB$SYSTEM_FLAG = 0 ORDER BY RDB$RELATION_NAME")
Else
    StringConnect = "MS Access;DATABASE=" & Arquivo & ";UID=" & IIf(Nz(Forms("LAttach")("Usuário"), "") = "", "Admin", Forms("LAttach")("Usuário"))
    If Forms("LAttach")("Senha") <> "" Then StringConnect = StringConnect & ";PWD=" & Forms("LAttach")("Senha")
    Set TBD = CurrentDb.CreateTableDef("SYS_USER_TABLES", 0, "MSysObjects", StringConnect)
    CurrentDb.TableDefs.Append TBD
    Set REC = CurrentDb.OpenRecordset("SELECT Name AS NOME FROM SYS_USER_TABLES WHERE ParentId in (SELECT SYS_USER_TABLES.Id FROM SYS_USER_TABLES WHERE (((SYS_USER_TABLES.Name) = 'Tables'));) and Type = 1 and not left(Name,4) = 'MSys';")
End If

' exclui tabelas antigas vinculadas
SS = LProgress(SS)
LProgress SS, "Apagando tabelas antigas", 1200, True, Janela
Z = 0
Do While Z < CurrentDb.TableDefs.Count
    If CurrentDb.TableDefs(Z).Connect <> "" And IIf(PREFIX <> "", CurrentDb.TableDefs(Z).Name Like PREFIX & "_*", True) Then
        DoCmd.DeleteObject acTable, CurrentDb.TableDefs(Z).Name
    Else
        Z = Z + 1
    End If
    SS = LProgress(SS)
Loop

'Atualiza as views do banco
For Each Qry In CurrentDb.QueryDefs
    If (Not (Qry.Name Like "~*" Or Qry.Name = "LConn")) And (Qry.Type = dbQSQLPassThrough) Then
        Qry.Connect = LInsere(StringConnect, "PWD", "", ";", "=")
    End If
Next

' vinculando tabelas remotas
SS = LProgress(SS)
LProgress SS, "Incluindo tabelas atualizadas", 1200, True, Janela
If Conexão = "Oracle" Then
    SS = LProgress(SS)
    Do While Not REC.EOF
        If Not LExists(CurrentDb.TableDefs, REC!Nome) Then
        
            Set TBD = CurrentDb.CreateTableDef(IIf(PREFIX <> "", PREFIX & "_", "") & REC!Nome, 0, REC!Owner & "." & REC!Nome, StringConnect)
            
            CurrentDb.TableDefs.Append TBD
            If LExists(CurrentDb.TableDefs, IIf(PREFIX <> "", PREFIX & "_", "") & REC!Nome) Then
                On Error GoTo lattach_OKErro
                
                ' cria visões
                If REC!TIPO = "VW" Then
                    XX = LExtrai(LConfig("VW_" & REC!Nome), "CHAVE", "|%|", ":")
                    If XX <> "" Then
                        CurrentDb.Execute "CREATE INDEX ID_" & REC!Nome & " ON " & REC!Nome & " (" & Replace(XX, ";", ",") & ")"

                    End If
                End If
            End If
            On Error GoTo lattach_OKErro
        
        Else
            If CurrentDb.TableDefs(REC!Nome).Connect = "" Then
                LErro "LAttach"
            End If
        End If
        SS = LProgress(SS)
        
        REC.MoveNext
    Loop
ElseIf Conexão = "Firebird" Then
    Do While Not REC.EOF
        
        If Nz(REC("RDB$VIEW_SOURCE"), "") = "" Then
            Set TBD = CurrentDb.CreateTableDef(IIf(PREFIX <> "", PREFIX & "_", "") & Trim(REC("RDB$RELATION_NAME")), 0, Trim(REC("RDB$RELATION_NAME")), StringConnect)
            On Error Resume Next
            CurrentDb.TableDefs.Append TBD
            
            CurrentDb.TableDefs(Trim(REC("RDB$RELATION_NAME"))).Properties.Append CurrentDb.TableDefs(Trim(REC("RDB$RELATION_NAME"))).CreateProperty("Description", dbText, Trim(REC("RDB$DESCRIPTION")))
            
            RS.FindFirst "RDB$RELATION_NAME = """ & Trim(REC("RDB$RELATION_NAME")) & """"
            Do While Not RS.EOF
                If (Trim(RS("RDB$RELATION_NAME")) <> Trim(REC("RDB$RELATION_NAME"))) Then
                    Exit Do
                End If
                CurrentDb.TableDefs(Trim(RS("RDB$RELATION_NAME"))).Fields(Trim(RS("RDB$FIELD_NAME"))).Properties.Append CurrentDb.TableDefs(Trim(RS("RDB$RELATION_NAME"))).Fields(Trim(RS("RDB$FIELD_NAME"))).CreateProperty("Description", dbText, Trim(RS("RDB$DESCRIPTION")))
                RS.MoveNext
            Loop
        Else
            If LExists(CurrentDb.QueryDefs, Trim(REC("RDB$RELATION_NAME"))) Then
                Set Qry = CurrentDb.QueryDefs(Trim(REC("RDB$RELATION_NAME")))
            Else
                Set Qry = CurrentDb.CreateQueryDef(Trim(REC("RDB$RELATION_NAME")))
            End If
            Qry.Connect = LInsere(StringConnect, "PWD", "", ";", "=")
            Qry.SQL = REC("RDB$VIEW_SOURCE")
            Qry.ReturnsRecords = True
        End If
        On Error GoTo lattach_OKErro
        
        SS = LProgress(SS)
        
        REC.MoveNext
    Loop
Else
    Do While Not REC.EOF
        Set TBD = CurrentDb.CreateTableDef(IIf(PREFIX <> "", PREFIX & "_", "") & REC!Nome, 0, REC!Nome, StringConnect)
        CurrentDb.TableDefs.Append TBD
        SS = LProgress(SS)
        
        REC.MoveNext
    Loop
End If
REC.Close

' exclusão objetos temporários
SS = LProgress(SS)
LProgress SS, "Apagando Objetos Temporários", 1200, True, Janela
If LExists(CurrentDb.TableDefs, "SYS_USER_TABLES") Then
    DoCmd.DeleteObject acTable, "SYS_USER_TABLES"
End If
If LExists(CurrentDb.TableDefs, "SYS_USER_FIELDS") Then
    DoCmd.DeleteObject acTable, "SYS_USER_FIELDS"
End If
If LExists(CurrentDb.QueryDefs, "SYS_USER_TABLES") Then
    DoCmd.DeleteObject acQuery, "SYS_USER_TABLES"
End If

LAttach_OKSai:
' finalização da rotina
LProgress Null
Exit Function



lattach_OKErro:
Dim xerr
xerr = LErro("LAttach_OK")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LAttach_OKSai
End Function



' REALIZA ENQUADRAMENTO DE TELA PRIMA EM FUNDO DE APLICATIVO
Function LEnquadra(Optional MM As String)
On Error GoTo LEnquadraErr
Dim RT As Rect, jan
If MM = "" Then
    MM = LTelaPrima
End If
LMoveWindow Forms(MM).hwnd, 0, 0, 0, 0, False

jan = LGethWndMDIClient()

' define o tamanho da janela
LGetWindowRect jan, RT
LMoveWindow Forms(MM).hwnd, 0, 0, RT.x2 - RT.x1 - 4, RT.y2 - RT.y1 - 4, True
' LEncaixaTela jan, Forms(MM).hwnd

LEnquadraSai:
Exit Function

LEnquadraErr:
Dim xerr
xerr = LErro("LEnquadra")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LEnquadraSai:
End Function

Function LGethWndDesktopClient()
On Error GoTo LGethWndDClientErr
Dim jan As Long
jan = LGetDesktopWindow()
jan = LFindWindowEx(jan, 0, "ProgMan", 0&)
jan = LFindWindowEx(jan, 0, "SHELLDLL_DefView", 0&)
jan = LFindWindowEx(jan, 0, "SysListView32", 0&)
LGethWndDesktopClient = jan

LGethWndDClientSai:
Exit Function

LGethWndDClientErr:
Dim xerr
xerr = LErro("LGethWndDesktopClient")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LGethWndDClientSai:

End Function


Function LGethWndMDIClient()
On Error GoTo LGethWndMDIClientErr
Dim jan As Long
' obtem mdi do access
jan = LGetWindow(Application.hWndAccessApp, GW_CHILD)
jan = LGetWindow(jan, GW_HWNDFIRST)
jan = LGetWindow(jan, GW_HWNDNEXT)
jan = LGetWindow(jan, GW_HWNDNEXT)
jan = LGetWindow(jan, GW_HWNDNEXT)
jan = LGetWindow(jan, GW_HWNDNEXT)
jan = LGetWindow(jan, GW_HWNDNEXT)
LGethWndMDIClient = jan


LGethWndMDIClientSai:
Exit Function

LGethWndMDIClientErr:
Dim xerr
xerr = LErro("LGethWndMDIClient")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LGethWndMDIClientSai:
End Function

' ATIVA FUNÇÕES RELACIONADAS COM TELA PRIMA DO APLICATIVO
Function LAtivaTelaPrima(FormName As String)
On Error GoTo LAtivaTelaPrimaErr
'Forms(FormName).hora = Format(Time, "hh:nn")
'Forms(FormName)!Título.Caption = LConfig("Aplicação")
'Forms(FormName)![_Título].Caption = LConfig("Aplicação")
'Forms(FormName)!Particular.Caption = LConfig("Descrição_Aplicação")
'Forms(FormName)![_Particular].Caption = LConfig("Descrição_Aplicação")
LEnquadra FormName

LAtivaTelaPrimaSai:
Exit Function

LAtivaTelaPrimaErr:
Dim xerr
xerr = LErro("LAtivaTelaPrima")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LAtivaTelaPrimaSai:
End Function


' LSERVIDOR
' OBTEM NOME DO SERVIDOR OU STRING DE CONEXÃO CASO O RECURSO LServidor NÃO ESTEJA DISPONÍVEL
' recurso LServidor deverá possuir a seguinte estrutura:
' SERVIDOR : método de obtenção de nome verdadeiro do servidor de banco de dados
Function LServidor(Optional CONN As String = "")
Dim SQL As String, REC As Recordset
On Error Resume Next
If CONN = "" Then
    CONN = LConfig("StringConnect")
End If

If LConfig("BaseDados") = "ORACLE" Then
    ' busca através de v$session
    SQL = "SELECT MACHINE FROM (SELECT MACHINE FROM V$SESSION WHERE NOT MACHINE IS NULL AND USER# = 0) WHERE ROWNUM = 1"
    LQueryDef "LServ", SQL, CONN
    Set REC = CurrentDb.OpenRecordset("select * from LServ")
    If Not (Err = 0 And REC.RecordCount <> 0) Then
        ' busca através de dbms_cx
        SQL = "SELECT " & LConfig("Esquema") & ".DBMS_LB.SERVIDOR MACHINE FROM DUAL"
        LQueryDef "LServ", SQL, CONN
        Set REC = CurrentDb.OpenRecordset("select * from LServ")
    End If
    
    If REC.RecordCount = 0 Then
        LServidor = "[indisponível]"
        Exit Function
    End If
Else
    If CurrentDb.TableDefs("SYS_CONFIG_GLOBAL").Connect = "" Then
        LServidor = CurrentDb.Name
    Else
        LServidor = LItem(CurrentDb.TableDefs("SYS_CONFIG_GLOBAL").Connect, 2, ";DATABASE=")
    End If
End If

LServidor = REC!MACHINE
End Function


' COPIA REGISTRO
Function LDuplica()
On Error Resume Next
Dim Txt As String

If Application.CurrentObjectType = acForm Then
    DoCmd.RunCommand acCmdSelectRecord
    Txt = ""
    If Err <> 0 Then
        Txt = "Não foi possível selecionar o registro atual para cópia"
    Else
        DoCmd.RunCommand acCmdCopy
        If Err <> 0 Then
           Txt = "Não foi possível realizar cópia do registro selecionado"
        Else
            DoCmd.RunCommand acCmdRecordsGoToNew
            If Err <> 0 Then
                Txt = "Não foi possível incluir novo registro durante procedimento de cópia"
            Else
                DoCmd.RunCommand acCmdSelectRecord
                If Err <> 0 Then
                    Txt = "Não foi possível colar conteúdo copiado no novo registro"
                Else
                    DoCmd.RunCommand acCmdPaste
                    If Err <> 0 Then
                        Txt = "Problemas ao tentar colar conteúdo copiado em novo registro"
                    End If
                End If
            End If
        End If
    End If
Else
    Txt = "Recurso de cópia somente permitido para formulários"
End If
If Txt <> "" Then
    MsgBox Txt, vbCritical + vbOKOnly, LNomeApl() & " (LDuplica)"
End If
End Function


' DEFINIÇÃO DE QUERYDEF NO SISTEMA
Sub LQueryDef(Nome As String, SQL, Optional CONN As String = "")
On Error GoTo LQueryDefErr
If Not IsNull(SQL) Then
    If Not LExists(CurrentDb.QueryDefs, Nome) Then
        CurrentDb.CreateQueryDef Nome
    End If
    CurrentDb.QueryDefs(Nome).SQL = SQL
    CurrentDb.QueryDefs(Nome).Connect = CONN
Else
    DoCmd.DeleteObject acQuery, Nome
End If

LQuerydefSai:
Exit Sub

LQueryDefErr:
Dim xerr
xerr = LErro("LQueryDef")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LQuerydefSai:
End Sub


' ABRE UM RELATÓRIO
Function LAbreReport(Nome As String, Optional Exibir As Variant = acViewPreview, Optional NomeFiltro As Variant = Null, Optional Condição As Variant = Null) As Integer
On Error GoTo LAbreReportErro

' verifica se licença está válida
If LCheck_(LConfig("FABRICANTE"), LConfig("APLICAÇÃO"), LConfig("VERSÃO")) <> 0 Then
    LErro "LInício", "{OK}Este programa necessita de liberação. Contacte suporte autorizado."
    Application.Quit acQuitSaveNone
End If

DoCmd.OpenReport Nome, Exibir, NomeFiltro, Condição

LAbreReportSai:
Exit Function

LAbreReportInexist:
    MsgBox "Recurso não implementado nesta versão. Entrar em contato com lucianoicraft@gmail.com e solicitar sua atualização.", vbExclamation + vbOKOnly, LConfig("Aplicação")
Resume LAbreReportSai

LAbreReportErro:
Dim xerr As Integer
If Err = 2103 Then Resume LAbreReportInexist
xerr = LErro("LAbreReport")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LAbreReportSai
End Function


'Apaga todas as tabelas vinculadas à aplicação, tanto por mdb como por ODBC.
Function LApagaTabelasVinculadas()
Dim REC As Recordset
If InStr(1, CurrentDb.Connect, ";DATABASE") <> 0 Then
    Set REC = CurrentDb.OpenRecordset("SELECT MSysObjects.Name, MSysObjects.Database FROM MSysObjects WHERE (((MSysObjects.Database)<>''));")
Else
    Set REC = CurrentDb.OpenRecordset("SELECT MSysObjects.name,MSysObjects.Connect FROM MSysObjects WHERE (((MSysObjects.Connect)<>''));")
End If
While Not REC.EOF
    DoCmd.DeleteObject acTable, REC("Name")
    REC.MoveNext
Wend
REC.Close
End Function


