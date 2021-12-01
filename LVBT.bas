Attribute VB_Name = "LVBT"
Option Compare Database
Option Explicit

' =====================================================
' OBJETO - LVBT - ROTINAS AINDA PENDENTES DE TRATAMENTO
' =====================================================
' DATA - HIST�RICO - T�CNICO
' 23 MAI 2002 - IMPLEMENTA��O C�DIGO BASE


' RECURSO DE POSICIONAMENTO DE REGISTRO ANTOM�TICO QUE POSICIONA NO �LTIMO REGISTRO MANIPULADO AO ABRIR UM FORMUL�RIO
' FF formul�rio que receber� o controle
' Chave campos concatenados com ; para busca de cada registro
' Conte�do posicionamento inicial do formul�rio
' SubFF Subformul�rio que ser� reposicionamento tamb�m conforme mudan�a de registro no formul�rio principal
' SubV�nculo Conte�do de chave vinculada para controle
' SubChave Conte�do inicial da busca no subformul�rio
Function LFormPos(ByVal FF As Form, Optional ByVal Chave As String = "", Optional ByVal Conte�do, Optional ByVal SubFF As Form, Optional ByVal SubV�nculo As String = "", Optional ByVal SubChave As String = "", Optional ByVal SubConte�do, Optional ByVal Filtro As String = "", Optional Auto As Integer = False)
On Error GoTo LFormPosErro
Dim Prop, SQL As String, pos As Integer, Txt As String
Dim Univ As String, Posic As String, Qry As QueryDef, Stab As String
 
' auto:
'       -1 registra e posiciona no primeiro
'        0 apenas registra
'        1 posiciona no primeiro
'        2 posiciona no anterior
'        3 posiciona no pr�ximo
'        4 posiciona no �ltimo
 
' Certeza de ter em FF o nome do form principal
 
If VarType(FF) <> 8 Then
    FF = FF.Name
End If
 
If Chave = "" Then
    
    ' posiciona o formul�rio no registro desejado
Posiciona:
    Prop = LProperty(CurrentDb, "Form_" & FF) & ""
    Chave = LExtrai(Prop, "Chave")
    Conte�do = LExtrai(Prop, "Conte�do")
    SubFF = LExtrai(Prop, "SubFF")
    SubV�nculo = LExtrai(Prop, "SubV�nculo")
    SubChave = LExtrai(Prop, "SubChave")
    SubConte�do = LExtrai(Prop, "SubConte�do")
    Filtro = LExtrai(Prop, "Filtro")
    Auto = 0
    
    DoCmd.Maximize
    
    ' busca registro no formul�rio principal
PosicionaDireto:
    If Conte�do <> "" Then
    
        ' busca atrav�s de pesquisa sequencial
        LCurtaMensagem "Realizando pesquisa."
        If Forms(FF).RecordsetClone.RecordCount <> 0 Then
            If Auto = 2 Then
                Forms(FF).RecordsetClone.FindPrevious Chave & " like " & Chr(34) & UCase(Conte�do) & Chr(34)
            ElseIf Auto = 3 Then
                Forms(FF).RecordsetClone.FindNext Chave & " like " & Chr(34) & UCase(Conte�do) & Chr(34)
            ElseIf Auto = 4 Then
                Forms(FF).RecordsetClone.FindLast Chave & " like " & Chr(34) & UCase(Conte�do) & Chr(34)
            Else
                pos = InStr(Chave, " & ")
                If pos <> 0 Then
                    Txt = Trim(Left(Chave, pos - 1))
                    Forms(FF).RecordsetClone.MoveFirst
                    Forms(FF).RecordsetClone.FindFirst Txt & " like Left(""" & Conte�do & """, len(" & Txt & "))"
                    If Not Forms(FF).RecordsetClone.NoMatch Then
                        If Forms(FF).RecordsetClone.AbsolutePosition > 0 Then
                            Forms(FF).RecordsetClone.MovePrevious
                            Forms(FF).RecordsetClone.FindNext Chave & " like " & Chr(34) & UCase(Conte�do) & Chr(34)
                        Else
                            Forms(FF).RecordsetClone.FindFirst Chave & " like " & Chr(34) & UCase(Conte�do) & Chr(34)
                        End If
                    End If
                Else
                    Forms(FF).RecordsetClone.FindFirst Chave & " like " & Chr(34) & UCase(Conte�do) & Chr(34)
                End If
            End If
        Else
            GoTo Conte�doN�oEncontrado
        End If
        
        LCurtaMensagem Null
        
        If Not Forms(FF).RecordsetClone.NoMatch Then
            Forms(FF).Bookmark = Forms(FF).RecordsetClone.Bookmark
            
            ' busca em subformul�rio caso exista
            If Not IsMissing(SubConte�do) Then
                If SubConte�do <> "" Then
                
                    ' certeza de SubFF ter o nome do subformul�rio
                    If VarType(SubFF) <> 8 Then
                        SubFF = SubFF.Name
                    End If
                    
                    If SubFF <> "" And SubV�nculo <> "" And SubChave <> "" Then
                        
                        pos = InStr(SubFF, "@")
                        If pos <> 0 Then
                            Stab = Mid(SubFF, pos + 1)
                                                    
                            ' busca atrav�s de pesquisa sequencial no subformul�rio
                            If Forms(FF)(SubFF).Form.RecordsetClone.RecordCount <> 0 Then
                                Forms(FF)(SubFF).Form.RecordsetClone.FindFirst SubChave & " like " & Chr(34) & UCase(SubConte�do) & Chr(34)
                                If Not Forms(FF)(SubFF).Form.RecordsetClone.NoMatch Then
                                    Forms(FF)(SubFF).Form.Bookmark = Forms(FF)(SubFF).Form.RecordsetClone.Bookmark
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Else
                        
Conte�doN�oEncontrado:
            ' posiciona em um novo registro caso n�o ache
 
            If Auto < 1 Then
                DoCmd.GoToRecord acDataForm, FF, acNewRec
            Else
                MsgBox "Conte�do n�o encontrado.", vbInformation + vbOKOnly, LConfig("Aplica��o")
            End If
        End If
        
'   -----------------------------
    Else
    
        ' posiciona em um novo registro
        DoCmd.GoToRecord acDataForm, FF, acNewRec
    End If
 
Else
 
    If Auto > 0 Then
        GoTo PosicionaDireto
    End If
    ' marca posi��o do formul�rio
    Prop = ""
    If Chave <> "" Then
        Prop = Prop & "Chave:" & Chave & Chr(13) & Chr(10)
    End If
    If Conte�do <> "" Then
        Prop = Prop & "Conte�do:" & Conte�do & Chr(13) & Chr(10)
    End If
    If Not IsMissing(SubFF) Then
        If SubFF <> "" Then
            Prop = Prop & "SubFF:" & SubFF & Chr(13) & Chr(10)
        End If
    End If
    If SubV�nculo <> "" Then
        Prop = Prop & "SubV�nculo:" & SubV�nculo & Chr(13) & Chr(10)
    End If
    If SubChave <> "" Then
        Prop = Prop & "SubChave:" & SubChave & Chr(13) & Chr(10)
    End If
    If Not IsMissing(SubFF) Then
        If SubConte�do <> "" Then
            Prop = Prop & "SubConte�do:" & SubConte�do & Chr(13) & Chr(10)
        End If
    End If
    If Filtro <> "" Then
        Prop = Prop & "Filtro:" & Filtro & Chr(13) & Chr(10)
    End If
    If Prop = "" Then
        Prop = Null
    End If
 
    LProperty CurrentDb, "Form_" & FF, Prop
    
    If Auto Then
        If LFormCarregado(FF) Then
            DoCmd.OpenForm FF, , , Filtro
            GoTo Posiciona
        Else
            DoCmd.OpenForm FF, , , Filtro
            If Filtro <> "" Then
                GoTo Posiciona
            End If
        End If
    End If
End If
 
LFormPosSai:
LCurtaMensagem Null
Exit Function
LFormPosErro:
LCurtaMensagem Null
Exit Function
Dim xerr
xerr = LErro("LFormPos")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LFormPosSai
End Function



' RECURSO LIMPLODE QUE PERMITE RETORNO DE UM FORMUL�RIO TIPO CADASTRO PARA O ANTERIOR COM UM VALOR
Function LImplode(FormN As String, CamposN As String, Optional SubFormN As String, Optional SubCamposN As String, Optional Form1 As String, Optional Campos1 As String, Optional SubForm1 As String, Optional SubCampos1 As String, Optional Volta1 As String, Optional Aberto As Integer)
On Error GoTo LImplodeErro
Dim NumCamposN As Integer, NumSubCamposN As Integer, NumCampos1 As Integer, NumSubCampos1 As Integer
Dim CampoN As String, SubCampoN As String
Dim Z As Integer, Conteu, ctl As Control, Txt As String, Txt1 As String
NumCamposN = LItem(CamposN, 0)
NumSubCamposN = LItem(SubCamposN, 0)
NumCampos1 = LItem(Campos1, 0)
NumSubCampos1 = LItem(SubCampos1, 0)
If Not NumCamposN + NumSubCamposN <> NumCampos1 + NumSubCampos1 Then
    For Z = 1 To NumCampos1 + NumSubCampos1
        If Z <= NumCampos1 Then
            Set ctl = Forms(Form1)(LItem(Campos1, Z))
        Else
            Set ctl = Forms(Form1)(SubForm1).Form(LItem(SubCampos1, Z - NumCampos1))
        End If
        If Z <= NumCamposN Then
            CampoN = LItem(CamposN, Z)
            If Not CampoN Like "{*}" Then
                If Forms(FormN)(CampoN) & "" = ctl.Value & "" Then
                Else
                    Forms(FormN)(CampoN) = ctl.Value
                    
                    On Error Resume Next
                    
                    'CASO PRECISE EXECUTAR ALGUMA FUN��O AO ATUALIZAR O CONTE�DO, CRIAR FUN��O LIMPLODERETORNA DENTRO DO FORM
                    Forms(FormN).LImplodeRetorna
                    On Error GoTo LImplodeErro
                End If
            Else
                If CampoN <> "{" & ctl.Value & "}" Then
                    MsgBox "Conte�do de retorno inv�lido. Escolha um registro onde " & ctl.Name & " seja " & Mid(CampoN, 2, Len(CampoN) - 2) & ".", vbCritical + vbOKOnly, LConfig("Aplica��o")
                    Exit Function
                End If
            End If
        Else
            SubCampoN = LItem(SubCamposN, Z - NumCamposN)
            If Not SubCampoN Like "{*}" Then
                If Forms(FormN)(SubFormN).Form(SubCampoN) & "" = ctl.Value & "" Then
                Else
                    Forms(FormN)(SubFormN).Form(SubCampoN) = ctl.Value
                    
                    On Error Resume Next
                    
                    'CASO PRECISE EXECUTAR ALGUMA FUN��O AO ATUALIZAR O CONTE�DO, CRIAR FUN��O LIMPLODERETORNA DENTRO DO FORM
                    Forms(FormN)(SubFormN).Form.LImplodeRetorna
                    On Error GoTo LImplodeErro
                End If
            Else
                If SubCampoN <> "{" & ctl.Value & "}" Then
                    MsgBox "Conte�do de retorno inv�lido. Escolha um registro onde " & ctl.Name & " seja " & Mid(SubCampoN, 2, Len(SubCampoN) - 2) & ".", vbCritical + vbOKOnly, LConfig("Aplica��o")
                    Exit Function
                End If
            End If
        End If
    Next
    If Not Aberto Then
        DoCmd.Close acForm, "Form1"
    End If
    If LItem(Volta1, 0) = 1 Then
        Forms(Form1)(Volta1).BorderStyle = 1
        Forms(Form1)(Volta1).SpecialEffect = 2
        Forms(Form1)(Volta1).OnDblClick = ""
    Else
        Txt = LItem(Volta1, 1)
        Txt1 = LItem(Volta1, 2)
        Forms(Form1)(Txt).Form(Txt1).BorderStyle = 1
        Forms(Form1)(Txt).Form(Txt1).SpecialEffect = 2
        Forms(Form1)(Txt).Form(Txt1).OnDblClick = ""
    End If
    DoCmd.OpenForm FormN
    If Not Aberto Then
        DoCmd.Close acForm, Form1
    End If
End If
LImplodeSai:
Exit Function

LImplodeErro:
Dim xerr
xerr = LErro("LImplode")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LImplodeSai
End Function



' RECURSO LEXPLODE QUE PERMITE ABRIR UM FORMUL�RIO TIPO CADASTRO A PARTIR DE UM CAMPO RELACIONADO
Function LExplode(FormN As String, CamposN As String, Optional SubFormN As String, Optional SubCamposN As String, Optional Form1 As String, Optional Campos1 As String, Optional SubForm1 As String, Optional SubCampos1 As String)
On Error GoTo LExplodeErro
Dim Conteu As String, ConteuN1 As String, SubConteuN1 As String, Expr1 As String, SubExpr1 As String, Z As Integer, CAMPO1 As String, ctl As Control
Dim NumCamposN As Integer, NumSubCamposN As Integer, NumCampos1 As Integer, NumSubCampos1 As Integer, Volta1 As String, Aberto As Integer, ITEM As String
Dim SubVolta1 As String
Dim Filtro As String, Txt
Expr1 = ""
SubExpr1 = ""
ConteuN1 = ""
SubConteuN1 = ""
Volta1 = ""
SubVolta1 = ""
Filtro = ""
NumCamposN = LItem(CamposN, 0)
NumSubCamposN = LItem(SubCamposN, 0)
NumCampos1 = LItem(Campos1, 0)
NumSubCampos1 = LItem(SubCampos1, 0)
If Not NumCamposN + NumSubCamposN <> NumCampos1 + NumSubCampos1 Then
    For Z = 1 To NumCamposN + NumSubCamposN
        If Z <= NumCamposN Then
            ITEM = LItem(CamposN, Z)
            If ITEM Like "{*}" Then
                Conteu = Mid(ITEM, 2, Len(ITEM) - 2)
            Else
                Conteu = Forms(FormN)(ITEM) & ""
            End If
        Else
            ITEM = LItem(SubCamposN, Z - NumCamposN)
            If ITEM Like "{*}" Then
                Conteu = Mid(ITEM, 2, Len(ITEM) - 2)
            Else
                Conteu = Forms(FormN)(SubFormN).Form(ITEM) & ""
            End If
        End If
        If Z <= NumCampos1 Then
            ConteuN1 = ConteuN1 & Conteu
            CAMPO1 = LItem(Campos1, Z)
            Expr1 = Expr1 & IIf(Expr1 <> "", " & ", "") & "[" & CAMPO1 & "]"
        Else
            SubConteuN1 = Conteu
            CAMPO1 = LItem(SubCampos1, Z - NumCampos1)
            SubExpr1 = SubExpr1 & IIf(SubExpr1 <> "", " & ", "") & "[" & CAMPO1 & "]"
        End If
        If ITEM Like "{*}" Then
            Filtro = Filtro & IIf(Filtro <> "", " AND ", "") & "[" & CAMPO1 & "] = " & Chr(34) & Conteu & Chr(34)
        Else
            Volta1 = CAMPO1
            If Z <= NumCampos1 Then
                SubVolta1 = ""
            Else
                SubVolta1 = SubForm1
            End If
        End If
    Next
    Aberto = LFormCarregado(Form1)
    LFormPos Form1, Expr1, ConteuN1, SubForm1, Expr1, SubExpr1, SubConteuN1, Filtro, True
    If SubVolta1 = "" Then
        Forms(Form1)(Volta1).SpecialEffect = 0
        Forms(Form1)(Volta1).BorderStyle = 4
        Forms(Form1)(Volta1).OnDblClick = "=LImplode(" & Chr(34) & FormN & Chr(34) & ", " & Chr(34) & CamposN & Chr(34) & ", " & Chr(34) & SubFormN & Chr(34) & ", " & Chr(34) & SubCamposN & Chr(34) & ", " & Chr(34) & Form1 & Chr(34) & ", " & Chr(34) & Campos1 & Chr(34) & ", " & Chr(34) & SubForm1 & Chr(34) & ", " & Chr(34) & SubCampos1 & Chr(34) & ", " & Chr(34) & Volta1 & Chr(34) & ", " & Aberto & ")"
    Else
        Forms(Form1)(SubVolta1).Form(Volta1).SpecialEffect = 0
        Forms(Form1)(SubVolta1).Form(Volta1).BorderStyle = 4
        Forms(Form1)(SubVolta1).Form(Volta1).OnDblClick = "=LImplode(" & Chr(34) & FormN & Chr(34) & ", " & Chr(34) & CamposN & Chr(34) & ", " & Chr(34) & SubFormN & Chr(34) & ", " & Chr(34) & SubCamposN & Chr(34) & ", " & Chr(34) & Form1 & Chr(34) & ", " & Chr(34) & Campos1 & Chr(34) & ", " & Chr(34) & SubForm1 & Chr(34) & ", " & Chr(34) & SubCampos1 & Chr(34) & ", " & Chr(34) & SubVolta1 & ";" & Volta1 & Chr(34) & ", " & Aberto & ")"
    End If
    
    For Z = 1 To NumCamposN + NumSubCamposN
        If Z <= NumCampos1 Then
            Set ctl = Forms(Form1)(LItem(Campos1, Z))
        Else
            Set ctl = Forms(Form1)(SubForm1).Form(LItem(SubCampos1, Z - NumCampos1))
        End If
        If Z <= NumCamposN Then
            ITEM = LItem(CamposN, Z)
            If ITEM Like "{*}" Then
                Conteu = Mid(ITEM, 2, Len(ITEM) - 2)
            Else
                Conteu = Forms(FormN)(ITEM) & ""
            End If
        Else
            ITEM = LItem(SubCamposN, Z - NumCamposN)
            If ITEM Like "{*}" Then
                Conteu = Mid(ITEM, 2, Len(ITEM) - 2)
            Else
                Conteu = Forms(FormN)(SubFormN).Form(ITEM) & ""
            End If
        End If
        If Not ITEM Like "{*}" Then
            If ctl.Value & "" = Conteu Then
            Else
                ctl.Value = Conteu
            End If
        Else
            ctl.Enabled = False
            ctl.DefaultValue = Chr(34) & Conteu & Chr(34)
        End If
    Next
End If
LExplodeSai:
Exit Function

LExplodeErro:
Dim xerr
xerr = LErro("LExplode")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LExplodeSai
End Function

' IMPLEMENTA O RECURSO DE BUSCA DIN�MICA
Function LBDina()
On Error GoTo LBDinaErro
Dim FF As Form, Param As String, REC As Recordset, DB As Database, tt
If Application.CurrentObjectType = A_FORM Then
    If Application.CurrentObjectName <> "LBuscaDin�mica" Then
        Set FF = Forms(Application.CurrentObjectName)
        tt = LExtrai(FF.Tag, "Tabela")
        If Nz(tt, "") = "" Then
            tt = FF.Name
        End If
        Set REC = CurrentDb.OpenRecordset("Select * From [SYS_Tabela] Where [Nome] = '" & tt & "';")
        If REC.RecordCount <> 0 Then
            REC.MoveFirst
            DoCmd.OpenForm "LBuscaDin�mica"
            
            Forms![LBuscaDin�mica]![In�cio].OnClick = "=LBusca(" & Chr(34) & FF.Name & Chr(34) & ", " & Chr(34) & REC![EXPR_PRIMARIA] & Chr(34) & ", [Busca], 1)"
            Forms![LBuscaDin�mica]![Anterior].OnClick = "=LBusca(" & Chr(34) & FF.Name & Chr(34) & ", " & Chr(34) & REC![EXPR_PRIMARIA] & Chr(34) & ", [Busca], 2)"
            Forms![LBuscaDin�mica]![Pr�ximo].OnClick = "=LBusca(" & Chr(34) & FF.Name & Chr(34) & ", " & Chr(34) & REC![EXPR_PRIMARIA] & Chr(34) & ", [Busca], 3)"
            Forms![LBuscaDin�mica]![Final].OnClick = "=LBusca(" & Chr(34) & FF.Name & Chr(34) & ", " & Chr(34) & REC![EXPR_PRIMARIA] & Chr(34) & ", [Busca], 4)"
            Forms![LBuscaDin�mica].Busca.StatusBarText = "Busca por " & Replace(CStr(REC!Chave), ";", " : ") & "."
            Forms![LBuscaDin�mica].OnClose = "=LBDin�mica(Forms('" & FF.Name & "'),Form.[Busca],Null,'Sa�da')"
            
        End If
    End If
End If

LBDinaSai:
Exit Function

LBDinaErro:
Dim xerr As Integer
xerr = LErro("LBDina")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LBDinaSai
End Function

' IMPLEMENTA O RECURSO DE BUSCA DIN�MICA
Function LBDinas()
On Error GoTo LBDinaErro
Dim FF As Form, Param As String, REC As Recordset, DB As Database, tt
If Application.CurrentObjectType = A_FORM Then
    If Application.CurrentObjectName <> "LBuscaDin�mica" Then
        Set FF = Forms(Application.CurrentObjectName)
        tt = LExtrai(FF.Tag, "Tabela")
        If Nz(tt, "") = "" Then
            tt = FF.Name
        End If
        Set REC = CurrentDb.OpenRecordset("Select * From [SYS_Tabela] Where [Nome] = '" & tt & "';")
        If REC.RecordCount <> 0 Then
            REC.MoveFirst
            DoCmd.OpenForm "LBuscaDin�mica"
            
            Forms![LBuscaDin�mica]![In�cio].OnClick = "=LBusca(" & Chr(34) & FF.Name & Chr(34) & ", " & Chr(34) & REC![EXPR_PRIMARIA] & Chr(34) & ", [Busca], 1)"
            Forms![LBuscaDin�mica]![Anterior].OnClick = "=LBusca(" & Chr(34) & FF.Name & Chr(34) & ", " & Chr(34) & REC![EXPR_PRIMARIA] & Chr(34) & ", [Busca], 2)"
            Forms![LBuscaDin�mica]![Pr�ximo].OnClick = "=LBusca(" & Chr(34) & FF.Name & Chr(34) & ", " & Chr(34) & REC![EXPR_PRIMARIA] & Chr(34) & ", [Busca], 3)"
            Forms![LBuscaDin�mica]![Final].OnClick = "=LBusca(" & Chr(34) & FF.Name & Chr(34) & ", " & Chr(34) & REC![EXPR_PRIMARIA] & Chr(34) & ", [Busca], 4)"
            Forms![LBuscaDin�mica].Busca.StatusBarText = "Busca por " & Replace(CStr(REC!Chave), ";", " : ") & "."
            Forms![LBuscaDin�mica].OnClose = "=LBDin�mica(Forms('" & FF.Name & "'),Form.[Busca],Null,'Sa�da')"
            
        End If
    End If
End If

LBDinaSai:
Exit Function

LBDinaErro:
Dim xerr As Integer
xerr = LErro("LBDina")
If xerr = 4 Then
    Resume 0
ElseIf xerr = 5 Then
    Resume Next
End If
Resume LBDinaSai
End Function




