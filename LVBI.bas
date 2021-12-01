Attribute VB_Name = "LVBI"
Option Compare Database
Option Explicit

' ======================================================
' OBJETO - LVBI - MANIPULAÇÃO DE ESTRUTURAS DE IMPRESSÃO
' ======================================================
' DATA - HISTÓRICO - TÉCNICO
' 23 MAI 2002 - IMPLEMENTAÇÃO CÓDIGO BASE

Type glr_tagDevMode
    DeviceName As String * 16
    SpecVersion As Integer
    DriverVersion As Integer
    Size As Integer
    DriverExtra As Integer
    Fields As Long
    Orientation As Integer
    PaperSize As Integer
    PaperLength As Integer
    PaperWidth As Integer
    Scale As Integer
    Copies As Integer
    DefaultSource As Integer
    PrintQuality As Integer
    Color As Integer
    Duplex As Integer
    Resolution As Integer
    TTOption As Integer
    Collate As Integer
    FormName As String * 16
    Pad As Long
    Bits As Long
    PW As Long
    PH As Long
    DFI As Long
    DFr As Long
End Type

Type glr_tagDevModeStr
    RGB As String * 96
End Type

Type glr_tagDevNames
    DriverPos As Integer
    DevicePos As Integer
    OutputPos As Integer
    Default As Integer
End Type

Type glr_tagDevNamesStr
    RGB As String * 8
End Type

Type glr_tagMarginInfo
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    DataOnly As Long
    Width As Long
    Height As Long
End Type

Type glr_tagMarginInfoStr
    RGB As String * 28
End Type



' CONFIGURA RELATÓRIOS E FORMULÁRIOS COM DEFINIÇÕES DA TABELA SYS_GRUPO_REL
Function LConfiguraImp()
On Error GoTo LConfiguraImpErro

Dim Z As Integer, Nome As String, mipstr As String, modstr As String, namstr As String
Dim pmod As String, pmip As String, pnam As String
Dim TIPO As String

TIPO = Application.CurrentObjectType
Nome = Application.CurrentObjectName

If TIPO = A_FORM And Not Nome Like "LTelaPrima*" Then
    On Error Resume Next
    DoCmd.SelectObject A_FORM, Nome
    DoCmd.DoMenuItem 0, 0, 7
    On Error GoTo LConfiguraImpErro
    Exit Function
ElseIf TIPO = A_REPORT Then
    On Error Resume Next
    DoCmd.SelectObject A_REPORT, Nome
    DoCmd.DoMenuItem 0, 0, 7
    On Error GoTo LConfiguraImpErro
    Exit Function
End If

If MsgBox("Todos os relatórios serão configurados com o padrão do aplicativo. Você confirma ?", 36 + 256, LConfig("Aplicação")) <> 6 Then
    Exit Function
End If

' reports
For Z = 0 To CurrentDb.Containers!Reports.Documents.Count - 1
    Nome = CurrentDb.Containers!Reports.Documents(Z).Name
    DoCmd.OpenReport Nome, acViewDesign
    
CONFIGREPORT:
    mipstr = Nz(Reports(Nome).PRTMIP, "")
    modstr = Nz(Reports(Nome).PrtDevMode, "")
    namstr = Nz(Reports(Nome).PrtDevNames, "")
    
    If pmod = "" Then
        On Error Resume Next
        pmod = ""
        DoCmd.DoMenuItem 0, 0, 7
        If Err = 0 Then
            pmod = Reports(Nome).PrtDevMode
            pmip = Reports(Nome).PRTMIP
            pnam = Reports(Nome).PrtDevNames
        End If
        On Error GoTo LConfiguraImpErro
        If pmod = "" Then
            DoCmd.Close acReport, Nome, acSaveNo
            MsgBox "Configuração de relatórios interrompida sem alterações.", vbInformation + vbOKOnly, LConfig("Aplicação")
            Exit Function
        End If
    End If
  
    LConfigImpItem Nome, mipstr, modstr, namstr, pmod, pmip, pnam
   
    On Error Resume Next
    Reports(Nome).PrtDevNames = pnam
    Reports(Nome).PrtDevMode = modstr
    Reports(Nome).PRTMIP = mipstr
    If Err <> 0 Then
        DoCmd.OpenForm Nome, acDesign
        DoCmd.DoMenuItem 0, 0, 7
        GoTo CONFIGREPORT
    End If
    On Error GoTo 0
    DoCmd.Close acReport, Nome, acSaveYes
Next
GoTo LConfiguraImpCont

' forms
For Z = 0 To CurrentDb.Containers!Forms.Documents.Count - 1
    Nome = CurrentDb.Containers!Forms.Documents(Z).Name
    DoCmd.OpenForm Nome, acDesign
    
CONFIGFORM:
    mipstr = Nz(Forms(Nome).PRTMIP, "")
    modstr = Nz(Forms(Nome).PrtDevMode, "")
    
    LConfigImpItem "FORMULÁRIO", mipstr, modstr, namstr, pmod, pmip, pnam
    
    Forms(Nome).PrtDevMode = modstr
    On Error Resume Next
    Forms(Nome).PRTMIP = mipstr
    If Err <> 0 Then
        DoCmd.OpenForm Nome, acDesign
        DoCmd.DoMenuItem 0, 0, 7
        GoTo CONFIGFORM
    End If
    On Error GoTo 0
    DoCmd.Close acForm, Nome, acSaveYes
Next

LConfiguraImpCont:
MsgBox "Todos os relatórios configurados com sucesso.", vbInformation + vbOKOnly, LConfig("Aplicação")

LConfiguraImpSai:
Exit Function

LConfiguraImpErro:
Dim xerr
xerr = LErro("LConfiguraImp")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LConfiguraImpSai

End Function




' TRABALHA COM A LCONFIGIMP DEFININDO OS PARÂMETROS PARA CADA OBJETO
Function LConfigImpItem(Nome As String, ByRef mipstr As String, ByRef modstr As String, ByRef namstr As String, pmod As String, pmip As String, pnam As String)
Dim PRTMIP As glr_tagMarginInfo, PRTMIPSTR As glr_tagMarginInfoStr
Dim PRTMOD As glr_tagDevMode, PRTMODSTR As glr_tagDevModeStr
Dim Orienta As Integer

Dim REC As Recordset

Set REC = CurrentDb.OpenRecordset("SELECT * from SYS_GRUPO_REL ORDER BY SEQ DESC")
Do While Not Nome Like REC!GRUPO
    REC.MoveNext
Loop

If Not REC.EOF Then

    ' prtmip
    mipstr = pmip
    PRTMIPSTR.RGB = Left(mipstr, 28)
    LSet PRTMIP = PRTMIPSTR
        
    PRTMIP.Top = REC!MARGEM_SUPERIOR * 567
    PRTMIP.Bottom = REC!MARGEM_INFERIOR * 567
    PRTMIP.Left = REC!MARGEM_ESQUERDA * 567
    PRTMIP.Right = REC!MARGEM_DIREITA * 567
    
    LSet PRTMIPSTR = PRTMIP
    mipstr = Left(PRTMIPSTR.RGB, 28) & Mid(mipstr, 29)
    
    ' prtmod
    PRTMODSTR.RGB = Left(modstr, 96)
    LSet PRTMOD = PRTMODSTR
    Orienta = PRTMOD.Orientation
    
    modstr = pmod
    PRTMODSTR.RGB = Left(modstr, 96)
    LSet PRTMOD = PRTMODSTR
        
    PRTMOD.PaperSize = REC!PAPEL_TIPO
    PRTMOD.Orientation = Orienta
    
    LSet PRTMODSTR = PRTMOD
    modstr = Left(PRTMODSTR.RGB, 96) & Mid(modstr, 97)
    
    ' prtnam
    'namstr = pnam
End If


LConfigImpItemSai:
Exit Function

LConfigImpItemErro:
Dim xerr
xerr = LErro("LConfigImpItem")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LConfigImpItemSai

End Function




Function LImprime()
On Error Resume Next
If Not Application.CurrentObjectName Like "LTelaPrima*" Then
    DoCmd.OpenForm "IMPRESSÃO RÁPIDA"
End If
End Function



Function LImpVisualiza()
On Error Resume Next
Dim Nome As String, FF As Form
Set FF = Forms(Application.CurrentObjectName)
If Not Application.CurrentObjectName Like "LTelaPrima*" Then
    MsgBox "limpvisualiza"
End If
End Function



Function LImpConfigura()
MsgBox "limpconfigura"
End Function





