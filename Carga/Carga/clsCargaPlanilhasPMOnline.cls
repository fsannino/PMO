VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCargaPMOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ImportarPlanilhasPMOnline(ByVal strDiretorio As String, _
                                          ByRef pProgressBarGeral As Object, _
                                          ByRef pProgressBarPlanilha As Object, _
                                          ByRef pCaptionFramePlanilha As Object) As String

    Dim strArqXLS As String
    
    With pProgressBarGeral
        .Min = 0
        .Value = 0
        .Max = 7
    End With
    
    With pProgressBarPlanilha
        .Min = 0
        .Value = 0
        .Max = 100
    End With
    
    strMensagem = "Nao foi possivel abrir a conexao com o banco."

    'Abre conexao com o Banco SQL-Server

    Call AbrirConexao

    strMensagem = "Nao foi possivel criar o log."
    
    strMensagem = "Nao foi possivel excluir os dados das tabelas."
    
    'Limpa a base de dados PMO

    Call Excluir_Dados_Pmo

    'Importa dados da planilha de Issues para base SQL-Server

    pProgressBarGeral.Value = 1
    pCaptionFramePlanilha.Caption = "IS.XLS"
    
    strArqXLS = strDiretorio & "\IS.XLS"

    If Not Incluir_Issues(strArqXLS, pProgressBarPlanilha) Then
        Call CriarArquivoLog(strMensagem)
        Call Excluir_Dados_Pmo
        ImportarPlanilhasPMOnline = strMensagem
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    'Importa dados da planilha de Risks para base SQL-Server

    pProgressBarGeral.Value = 2
    pCaptionFramePlanilha.Caption = "RK.XLS"
    
    strArqXLS = strDiretorio & "\RK.XLS"

    If Not Incluir_Risks(strArqXLS, pProgressBarPlanilha) Then
        Call CriarArquivoLog(strMensagem)
        Call Excluir_Dados_Pmo
        ImportarPlanilhasPMOnline = strMensagem
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    'Importa dados da planilha de Actions para base SQL-Server

    pProgressBarGeral.Value = 3
    pCaptionFramePlanilha.Caption = "AC.XLS"
    
    strArqXLS = strDiretorio & "\AC.XLS"

    If Not Incluir_Actions(strArqXLS, pProgressBarPlanilha) Then
        Call CriarArquivoLog(strMensagem)
        Call Excluir_Dados_Pmo
        ImportarPlanilhasPMOnline = strMensagem
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    'Importa dados da planilha de Chang_Requests para base SQL-Server

    pProgressBarGeral.Value = 4
    pCaptionFramePlanilha.Caption = "CR.XLS"
    
    strArqXLS = strDiretorio & "\CR.XLS"

    If Not Incluir_Chang_Requests(strArqXLS, pProgressBarPlanilha) Then
        Call CriarArquivoLog(strMensagem)
        Call Excluir_Dados_Pmo
        ImportarPlanilhasPMOnline = strMensagem
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    'Importa dados da planilha de Tasks para base SQL-Server

    pProgressBarGeral.Value = 5
    pCaptionFramePlanilha.Caption = "TK.XLS"
    
    strArqXLS = strDiretorio & "\TK.XLS"

    If Not Incluir_Tasks(strArqXLS, pProgressBarPlanilha) Then
        Call CriarArquivoLog(strMensagem)
        Call Excluir_Dados_Pmo
        ImportarPlanilhasPMOnline = strMensagem
        Screen.MousePointer = vbDefault
        Exit Function
    End If

'    'Importa dados da planilha de Meeting para base SQL-Server
'
'    pProgressBarGeral.Value = 6
'    pCaptionFramePlanilha.Caption = "MT.XLS"
'
'    strArqXLS = strDiretorio & "\MT.XLS"
'
'    If Not Incluir_Meeting(strArqXLS, pProgressBarPlanilha) Then
'        Call CriarArquivoLog(strMensagem)
'        Call Excluir_Dados_Pmo
'        ImportarPlanilhasPMOnline = strMensagem
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If

    'Importa dados da planilha de Traceability para base SQL-Server

    pProgressBarGeral.Value = 6
    pCaptionFramePlanilha.Caption = "TC.XLS"
    
    strArqXLS = strDiretorio & "\TC.XLS"

    If Not Incluir_Traceability(strArqXLS, pProgressBarPlanilha) Then
        Call CriarArquivoLog(strMensagem)
        Call Excluir_Dados_Pmo
        ImportarPlanilhasPMOnline = strMensagem
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    pProgressBarGeral.Value = 7
    pCaptionFramePlanilha.Caption = "Gerando informações Scorecard."

    strMensagem = "Nao foi possivel incluir Scorecard."
    
    Call Incluir_Scorecard

    strMensagem = "Nao foi possivel incluir Scorecard Risks."
    
    Call Incluir_Scorecard_Risks
    
    strMensagem = "Arquivos importados com sucesso."

    Call CriarArquivoLog(strMensagem)

    ImportarPlanilhasPMOnline = strMensagem
    
    Screen.MousePointer = vbDefault

    Exit Function

End Function

Private Function Incluir_Issues(ByVal strArqXLS As String, _
                                ByRef pProgressBarPlanilha As Object) As Boolean

'Variaveis Issues

Dim IS_Project_Code          As String
Dim IS_Sr_No                 As Integer
Dim IS_ID                    As String
Dim IS_Nome                  As String
Dim IS_Access                As String
Dim IS_Check_If_Resolved     As String
Dim IS_Aberto_em             As Date
Dim IS_Data_Limite           As Date
Dim IS_Data_Fechamento       As Date
Dim IS_Descricao             As String
Dim IS_Estimativa_Horas      As Integer
Dim IS_Criticidade_Impacto   As String
Dim IS_Tipo_Impacto          As String
Dim IS_Modulo                As String
Dim IS_Overall_Status        As String
Dim IS_Percent_Complete      As Double
Dim IS_Priority              As String
Dim IS_Resolution            As String
Dim IS_Equipe                As String
Dim IS_Usuario               As String
Dim IS_Workflow              As String
Dim IS_Identificador         As String
Dim IS_Designador            As String
Dim IS_Responsavel           As String
Dim IS_Revisor               As String
Dim IS_Aprovador             As String
Dim IS_Comments              As String
Dim IS_Traceability          As String
Dim IS_Activity_Log          As String
Dim IS_Causa                 As String
Dim IS_Comite                As String
Dim IS_Descricao_Impacto     As String

Dim XL As Excel.Application
Dim intContador As Integer
Dim strIndice As String
Dim strSql As String

Incluir_Issues = False

On Error GoTo ErrIssues


    strMensagem = "Issues - Nao foi possivel criar objeto do excel."

    Set XL = CreateObject("Excel.Application")
    
    strMensagem = "Issues - Nao foi possivel abrir arquivo do excel."
    
    Call XL.Workbooks.Open(strArqXLS)
    
    intContador = 2
    
    With pProgressBarPlanilha
        .Min = intContador
        .Value = intContador
        .Max = (XL.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row) + 1
    End With
   
    strIndice = "A" + Trim(CStr(intContador))
    
    ' Enquanto existirem dados na planilha
    Do While (XL.Range(strIndice).Value <> "")
       
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Project_Code = XL.Range(strIndice).Value
        
        strIndice = "B" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Sr_No = XL.Range(strIndice).Value
        
        strIndice = "C" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        If Len(Trim(XL.Range(strIndice).Value)) = 4 Then
            IS_ID = Left(Trim(XL.Range(strIndice).Value), 3) & "0" & Right(Trim(XL.Range(strIndice).Value), 1)
        Else
            IS_ID = XL.Range(strIndice).Value
        End If
        
        strIndice = "D" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Nome = XL.Range(strIndice).Value
        
        strIndice = "E" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Access = XL.Range(strIndice).Value
        
        strIndice = "F" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Check_If_Resolved = XL.Range(strIndice).Value
        
        strIndice = "G" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Aberto_em = XL.Range(strIndice).Value
        
        strIndice = "H" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Data_Limite = XL.Range(strIndice).Value
        
        strIndice = "I" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Data_Fechamento = XL.Range(strIndice).Value
        
        strIndice = "J" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        
        '*************************************
        IS_Descricao = XL.Range(strIndice).Value
        
        strIndice = "K" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Descricao = IS_Descricao & XL.Range(strIndice).Value
        
        strIndice = "L" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Descricao = IS_Descricao & XL.Range(strIndice).Value
        
        strIndice = "M" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Descricao = IS_Descricao & XL.Range(strIndice).Value
        
        IS_Descricao = TrocarCarEsp_Enter(IS_Descricao)

        
        '***************************************
        
        
        strIndice = "N" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Estimativa_Horas = XL.Range(strIndice).Value
        
        strIndice = "O" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Criticidade_Impacto = XL.Range(strIndice).Value
        
        strIndice = "P" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Tipo_Impacto = XL.Range(strIndice).Value
        
        strIndice = "Q" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Modulo = XL.Range(strIndice).Value
        
        strIndice = "R" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Overall_Status = XL.Range(strIndice).Value
        
        strIndice = "S" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Percent_Complete = XL.Range(strIndice).Value
        
        strIndice = "T" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Priority = XL.Range(strIndice).Value
        
        strIndice = "U" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        '*******************************************
        
        IS_Resolution = XL.Range(strIndice).Value
        
        strIndice = "V" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Resolution = IS_Resolution & XL.Range(strIndice).Value
        
        strIndice = "W" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Resolution = IS_Resolution & XL.Range(strIndice).Value
        
        strIndice = "X" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Resolution = IS_Resolution & XL.Range(strIndice).Value
        
        IS_Resolution = TrocarCarEsp_Enter(IS_Resolution)
        
        '******************************************
        
        
        strIndice = "Y" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Equipe = XL.Range(strIndice).Value
        
        strIndice = "Z" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Usuario = XL.Range(strIndice).Value
        
        strIndice = "AA" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Workflow = XL.Range(strIndice).Value
        
        strIndice = "AB" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Identificador = XL.Range(strIndice).Value
        
        strIndice = "AC" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Designador = XL.Range(strIndice).Value
        
        strIndice = "AD" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Responsavel = XL.Range(strIndice).Value
        
        strIndice = "AE" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Revisor = XL.Range(strIndice).Value
        
        strIndice = "AF" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Aprovador = XL.Range(strIndice).Value
        
        
        strIndice = "AG" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        
        '******************************************************
        
        IS_Comments = XL.Range(strIndice).Value
        
        strIndice = "AH" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Comments = IS_Comments & XL.Range(strIndice).Value
        
        strIndice = "AI" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Comments = IS_Comments & XL.Range(strIndice).Value
        
        strIndice = "AJ" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Comments = IS_Comments & XL.Range(strIndice).Value
        
        strIndice = "AK" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Comments = IS_Comments & XL.Range(strIndice).Value
        
        strIndice = "AL" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Comments = IS_Comments & XL.Range(strIndice).Value
        
        strIndice = "AM" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Comments = IS_Comments & XL.Range(strIndice).Value
        
        strIndice = "AN" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Comments = IS_Comments & XL.Range(strIndice).Value
        
        strIndice = "AO" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Comments = IS_Comments & XL.Range(strIndice).Value
        
        strIndice = "AP" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Comments = IS_Comments & XL.Range(strIndice).Value
        
        IS_Comments = TrocarCarEsp_Enter(IS_Comments)
        
        '**********************************************************
        
        
        strIndice = "AQ" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Traceability = XL.Range(strIndice).Value
        
        strIndice = "AR" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Traceability = IS_Traceability & XL.Range(strIndice).Value
        
        strIndice = "AS" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Traceability = IS_Traceability & XL.Range(strIndice).Value
        
        strIndice = "AT" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Traceability = IS_Traceability & XL.Range(strIndice).Value
        
        IS_Traceability = TrocarCarEsp_Enter(IS_Traceability)
        
        '****************************************************
        
        
        strIndice = "AU" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Activity_Log = XL.Range(strIndice).Value
        
        strIndice = "AV" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Activity_Log = IS_Activity_Log & XL.Range(strIndice).Value
        
        strIndice = "AW" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Activity_Log = IS_Activity_Log & XL.Range(strIndice).Value
        
        strIndice = "AX" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Activity_Log = IS_Activity_Log & XL.Range(strIndice).Value
        
        IS_Activity_Log = TrocarCarEsp_Enter(IS_Activity_Log)
        
        '***************************************************
                
        strIndice = "AY" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Causa = XL.Range(strIndice).Value
        
        strIndice = "AZ" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Comite = XL.Range(strIndice).Value
        
        strIndice = "BA" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Descricao_Impacto = XL.Range(strIndice).Value
       
       
        strMensagem = "Issues - Nao foi possivel incluir os dados no BD."
       
       
        strSql = "SP_INCLUIR_ISSUES "
        
        strSql = strSql & "'" & IS_Project_Code & "', "
        strSql = strSql & "" & IS_Sr_No & ", "
        strSql = strSql & "'" & IS_ID & "', "
        strSql = strSql & "'" & gsReplicaPlics(IS_Nome) & "', "
        strSql = strSql & "'" & IS_Access & "', "
        strSql = strSql & "'" & IS_Check_If_Resolved & "', "
        strSql = strSql & "'" & Mid(Format(IS_Aberto_em, "dd/mm/yyyy"), 7, 4) & Mid(Format(IS_Aberto_em, "dd/mm/yyyy"), 4, 2) & Mid(Format(IS_Aberto_em, "dd/mm/yyyy"), 1, 2) & "', "
        strSql = strSql & "'" & Mid(Format(IS_Data_Limite, "dd/mm/yyyy"), 7, 4) & Mid(Format(IS_Data_Limite, "dd/mm/yyyy"), 4, 2) & Mid(Format(IS_Data_Limite, "dd/mm/yyyy"), 1, 2) & "', "
        strSql = strSql & "'" & Mid(Format(IS_Data_Fechamento, "dd/mm/yyyy"), 7, 4) & Mid(Format(IS_Data_Fechamento, "dd/mm/yyyy"), 4, 2) & Mid(Format(IS_Data_Fechamento, "dd/mm/yyyy"), 1, 2) & "', "
        strSql = strSql & "'" & gsReplicaPlics(IS_Descricao) & "', "
        strSql = strSql & "" & IS_Estimativa_Horas & ", "
        strSql = strSql & "'" & IS_Criticidade_Impacto & "', "
        strSql = strSql & "'" & IS_Tipo_Impacto & "', "
        strSql = strSql & "'" & IS_Modulo & "', "
        strSql = strSql & "'" & IS_Overall_Status & "', "
        strSql = strSql & "" & IS_Percent_Complete & ", "
        strSql = strSql & "'" & IS_Priority & "', "
        strSql = strSql & "'" & gsReplicaPlics(IS_Resolution) & "', "
        strSql = strSql & "'" & IS_Equipe & "', "
        strSql = strSql & "'" & gsReplicaPlics(IS_Usuario) & "', "
        strSql = strSql & "'" & IS_Workflow & "', "
        strSql = strSql & "'" & gsReplicaPlics(IS_Identificador) & "', "
        strSql = strSql & "'" & gsReplicaPlics(IS_Designador) & "', "
        strSql = strSql & "'" & gsReplicaPlics(IS_Responsavel) & "', "
        strSql = strSql & "'" & gsReplicaPlics(IS_Revisor) & "', "
        strSql = strSql & "'" & gsReplicaPlics(IS_Aprovador) & "', "
        strSql = strSql & "'" & gsReplicaPlics(IS_Comments) & "', "
        strSql = strSql & "'" & gsReplicaPlics(IS_Traceability) & "', "
        strSql = strSql & "'" & gsReplicaPlics(IS_Activity_Log) & "', "
        strSql = strSql & "'" & gsReplicaPlics(IS_Causa) & "', "
        strSql = strSql & "'" & gsReplicaPlics(IS_Comite) & "', "
        strSql = strSql & "'" & gsReplicaPlics(IS_Descricao_Impacto) & "'"
        
        dbConexao.Execute strSql
        
        intContador = intContador + 1
    
        strIndice = "A" + Trim(CStr(intContador))
        
        pProgressBarPlanilha.Value = intContador
    
    Loop
   
    ' Fecha a planilha e libera o objeto
    Call XL.Workbooks.Close
    
    Set XL = Nothing
    
    Incluir_Issues = True

    Exit Function
    
ErrIssues:
    
   
    If Not (XL Is Nothing) Then
        Call XL.Workbooks.Close
        Set XL = Nothing
    End If
    
    Set XL = Nothing
    
    Incluir_Issues = False
    
End Function

Private Sub CriarArquivoLog(strMensagem As String)

Dim strNoArquivo As String
Dim bytArquivo As Byte
Dim Log As Log_TP

    bytArquivo = FreeFile(0)
    
    strNoArquivo = App.Path & "\Log.txt"
    
    With Log
        .Campo_1 = " "
        .Campo_2 = strMensagem
        .Campo_3 = vbCrLf
    End With
   
    If Dir(strNoArquivo) <> "" Then
       Kill strNoArquivo
    End If
    
    Open strNoArquivo For Random _
                    As bytArquivo Len = Len(Log)
    
    
    Put #bytArquivo, , Log
    
    Close
    
End Sub

Private Sub Excluir_Dados_Pmo()
       
        Dim strSql As String
        
        strSql = "SP_EXCLUIR_DADOS_PMO "
        
        dbConexao.Execute strSql

End Sub

Private Function Incluir_Actions(ByVal strArqXLS As String, _
                                 ByRef pProgressBarPlanilha As Object) As Boolean

'Variaveis Actions

Dim AT_Project_Code As String
Dim AT_Sr_No As Integer
Dim AT_ID As String
Dim AT_Name As String
Dim AT_Access As String
Dim AT_Aberto_em As Date
Dim AT_Data_Limite As String
Dim AT_Data_Fechamento As String
Dim AT_Descricao As String
Dim AT_Estimativa_Horas As Integer
Dim AT_Impacto As String
Dim AT_Motivo As String
Dim AT_Modulo As String
Dim AT_Overall_Status As String
Dim AT_Percent_Complete As Double
Dim AT_Priority As String
Dim AT_Equipe As String
Dim AT_Usuario As String
Dim AT_Workflow As String
Dim AT_Identificador As String
Dim AT_Responsavel As String
Dim AT_Revisor As String
Dim AT_Comments As String
Dim AT_Traceability As String
Dim AT_Activity_Log As String
Dim AT_Comite As String

Dim XL As Excel.Application
Dim intContador As Integer
Dim strIndice As String
Dim strSql As String

Incluir_Actions = False

On Error GoTo ErrActions

    strMensagem = "Actions - Nao foi possivel criar objeto do excel."

    Set XL = CreateObject("Excel.Application")

    strMensagem = "Actions - Nao foi possivel abrir arquivo do excel."

    Call XL.Workbooks.Open(strArqXLS)
   
    intContador = 2
    
    With pProgressBarPlanilha
        .Min = intContador
        .Value = intContador
        .Max = (XL.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row) + 1
    End With
   
    strIndice = "A" + Trim(CStr(intContador))
    
    ' Enquanto existirem dados na planilha
    Do While (XL.Range(strIndice).Value <> "")
      
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Project_Code = XL.Range(strIndice).Value
        
        strIndice = "B" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Sr_No = XL.Range(strIndice).Value
        
        strIndice = "C" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        If Len(Trim(XL.Range(strIndice).Value)) = 4 Then
            AT_ID = Left(Trim(XL.Range(strIndice).Value), 3) & "0" & Right(Trim(XL.Range(strIndice).Value), 1)
        Else
            AT_ID = XL.Range(strIndice).Value
        End If
        
        strIndice = "D" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Name = XL.Range(strIndice).Value
        
        strIndice = "E" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Access = XL.Range(strIndice).Value
        
        strIndice = "F" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Aberto_em = XL.Range(strIndice).Value
        
        strIndice = "G" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Data_Limite = XL.Range(strIndice).Value
        
        strIndice = "H" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Data_Fechamento = XL.Range(strIndice).Value
        
        strIndice = "I" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Descricao = XL.Range(strIndice).Value
        
        strIndice = "J" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Descricao = AT_Descricao & XL.Range(strIndice).Value
        
        strIndice = "K" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Descricao = AT_Descricao & XL.Range(strIndice).Value
        
        strIndice = "L" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Descricao = AT_Descricao & XL.Range(strIndice).Value
        
        AT_Descricao = TrocarCarEsp_Enter(AT_Descricao)
        
        strIndice = "M" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Estimativa_Horas = XL.Range(strIndice).Value
        
        strIndice = "N" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Impacto = XL.Range(strIndice).Value
        
        strIndice = "O" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Motivo = XL.Range(strIndice).Value
        
        strIndice = "P" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Modulo = XL.Range(strIndice).Value
        
        strIndice = "Q" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Overall_Status = XL.Range(strIndice).Value
        
        strIndice = "R" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Percent_Complete = XL.Range(strIndice).Value
        
        strIndice = "S" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Priority = XL.Range(strIndice).Value
        
        strIndice = "T" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Equipe = XL.Range(strIndice).Value
        
        strIndice = "U" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Usuario = XL.Range(strIndice).Value
        
        strIndice = "V" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Workflow = XL.Range(strIndice).Value
        
        strIndice = "W" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Identificador = XL.Range(strIndice).Value
        
        strIndice = "X" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Responsavel = XL.Range(strIndice).Value

        strIndice = "Y" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Revisor = XL.Range(strIndice).Value
        
        strIndice = "Z" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Comments = XL.Range(strIndice).Value
        
        strIndice = "AA" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Comments = AT_Comments & XL.Range(strIndice).Value
        
        strIndice = "AB" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Comments = AT_Comments & XL.Range(strIndice).Value
        
        strIndice = "AC" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Comments = AT_Comments & XL.Range(strIndice).Value
        
        strIndice = "AD" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Comments = AT_Comments & XL.Range(strIndice).Value
        
        strIndice = "AE" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Comments = AT_Comments & XL.Range(strIndice).Value
        
        strIndice = "AF" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Comments = AT_Comments & XL.Range(strIndice).Value
        
        strIndice = "AG" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Comments = AT_Comments & XL.Range(strIndice).Value
        
        strIndice = "AH" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Comments = AT_Comments & XL.Range(strIndice).Value
        
        strIndice = "AI" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Comments = AT_Comments & XL.Range(strIndice).Value
       
        AT_Comments = TrocarCarEsp_Enter(AT_Comments)
       
       
       
        strIndice = "AJ" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Traceability = XL.Range(strIndice).Value
       
        strIndice = "AK" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Traceability = AT_Traceability & XL.Range(strIndice).Value
       
        strIndice = "AL" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Traceability = AT_Traceability & XL.Range(strIndice).Value
       
        strIndice = "AM" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Traceability = AT_Traceability & XL.Range(strIndice).Value
        
        AT_Traceability = TrocarCarEsp_Enter(AT_Traceability)
        
        
        
        strIndice = "AN" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Activity_Log = XL.Range(strIndice).Value
        
        strIndice = "AO" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Activity_Log = AT_Activity_Log & XL.Range(strIndice).Value
        
        strIndice = "AP" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Activity_Log = AT_Activity_Log & XL.Range(strIndice).Value
        
        strIndice = "AQ" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Activity_Log = AT_Activity_Log & XL.Range(strIndice).Value
        
        AT_Activity_Log = TrocarCarEsp_Enter(AT_Activity_Log)
        
        
        
        strIndice = "AR" + Trim(CStr(intContador))
        strMensagem = "Actions - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        AT_Comite = XL.Range(strIndice).Value
       
        strMensagem = "Actions - Nao foi possivel incluir os dados no BD."
       
        strSql = "SP_INCLUIR_ACTIONS "
        
        strSql = strSql & "'" & AT_Project_Code & "', "
        strSql = strSql & "" & AT_Sr_No & ", "
        strSql = strSql & "'" & AT_ID & "', "
        strSql = strSql & "'" & gsReplicaPlics(AT_Name) & "', "
        strSql = strSql & "'" & AT_Access & "', "
        strSql = strSql & "'" & Mid(Format(AT_Aberto_em, "dd/mm/yyyy"), 7, 4) & Mid(Format(AT_Aberto_em, "dd/mm/yyyy"), 4, 2) & Mid(Format(AT_Aberto_em, "dd/mm/yyyy"), 1, 2) & "', "
        strSql = strSql & "'" & Mid(Format(AT_Data_Limite, "dd/mm/yyyy"), 7, 4) & Mid(Format(AT_Data_Limite, "dd/mm/yyyy"), 4, 2) & Mid(Format(AT_Data_Limite, "dd/mm/yyyy"), 1, 2) & "', "
        strSql = strSql & "'" & Mid(Format(AT_Data_Fechamento, "dd/mm/yyyy"), 7, 4) & Mid(Format(AT_Data_Fechamento, "dd/mm/yyyy"), 4, 2) & Mid(Format(AT_Data_Fechamento, "dd/mm/yyyy"), 1, 2) & "', "
        strSql = strSql & "'" & gsReplicaPlics(AT_Descricao) & "', "
        strSql = strSql & "" & AT_Estimativa_Horas & ", "
        strSql = strSql & "'" & AT_Impacto & "', "
        strSql = strSql & "'" & AT_Motivo & "', "
        strSql = strSql & "'" & AT_Modulo & "', "
        strSql = strSql & "'" & AT_Overall_Status & "', "
        strSql = strSql & "" & AT_Percent_Complete & ", "
        strSql = strSql & "'" & AT_Priority & "', "
        strSql = strSql & "'" & AT_Equipe & "', "
        strSql = strSql & "'" & gsReplicaPlics(AT_Usuario) & "', "
        strSql = strSql & "'" & AT_Workflow & "', "
        strSql = strSql & "'" & gsReplicaPlics(AT_Identificador) & "', "
        strSql = strSql & "'" & gsReplicaPlics(AT_Responsavel) & "', "
        strSql = strSql & "'" & gsReplicaPlics(AT_Revisor) & "', "
        strSql = strSql & "'" & gsReplicaPlics(AT_Comments) & "', "
        strSql = strSql & "'" & gsReplicaPlics(AT_Traceability) & "', "
        strSql = strSql & "'" & gsReplicaPlics(AT_Activity_Log) & "', "
        strSql = strSql & "'" & AT_Comite & "'"
        
        dbConexao.Execute strSql
        
        intContador = intContador + 1
        
        pProgressBarPlanilha.Value = intContador
    
        strIndice = "A" + Trim(CStr(intContador))
    
    Loop
   
    ' Fecha a planilha e libera o objeto
    Call XL.Workbooks.Close
    
    Set XL = Nothing
    
    Incluir_Actions = True

    Exit Function
    
ErrActions:
    
    If Not (XL Is Nothing) Then
        Call XL.Workbooks.Close
        Set XL = Nothing
    End If
    
    Set XL = Nothing

    Incluir_Actions = False

End Function
Private Function Incluir_Chang_Requests(ByVal strArqXLS As String, _
                                        ByRef pProgressBarPlanilha As Object) As Boolean

'Variaveis Chang_Requests

Dim CR_Project_Code As String
Dim CR_Sr_No As Integer
Dim CR_ID As String
Dim CR_Name As String
Dim CR_Aberto_em As Date
Dim CR_Access As String
Dim CR_Alteracao As String
Dim CR_Check_If_Resolved As String
Dim CR_Comite As String
Dim CR_Data_Limite As Date
Dim CR_Data_Fechamento As Date
Dim CR_Descricao As String
Dim CR_Equipe As String
Dim CR_Estimativa_Horas As Integer
Dim CR_Limite As String
Dim CR_Modulo As String
Dim CR_Overall_Status As String
Dim CR_Percent_Complete As Double
Dim CR_Priority As String
Dim CR_Release As String
Dim CR_Resolution As String
Dim CR_Revision_No As String
Dim CR_Site As String
Dim CR_Usuario As String
Dim CR_Workflow As String
Dim CR_Identificador As String
Dim CR_Requisitante As String
Dim CR_PMO As String
Dim CR_Analista As String
Dim CR_Revisor As String
Dim CR_Aprovador As String
Dim CR_Comments As String
Dim CR_Activity_Log As String
Dim CR_Traceability As String

Dim XL As Excel.Application
Dim intContador As Integer
Dim strIndice As String
Dim strSql As String

Incluir_Chang_Requests = False

On Error GoTo ErrChang_Requests

    strMensagem = "Chang Requests - Nao foi possivel criar objeto do excel."

    Set XL = CreateObject("Excel.Application")
    
    strMensagem = "Chang Requests - Nao foi possivel abrir arquivo do excel."

    Call XL.Workbooks.Open(strArqXLS)
   
    intContador = 2
    
    With pProgressBarPlanilha
        .Min = intContador
        .Value = intContador
        .Max = (XL.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row) + 1
    End With
   
    strIndice = "A" + Trim(CStr(intContador))
    
    ' Enquanto existirem dados na planilha
    Do While (XL.Range(strIndice).Value <> "")
       
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
       
        CR_Project_Code = XL.Range(strIndice).Value
        
        strIndice = "B" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Sr_No = XL.Range(strIndice).Value
        
        strIndice = "C" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        If Len(Trim(XL.Range(strIndice).Value)) = 4 Then
            CR_ID = Left(Trim(XL.Range(strIndice).Value), 3) & "0" & Right(Trim(XL.Range(strIndice).Value), 1)
        Else
            CR_ID = XL.Range(strIndice).Value
        End If
        
        strIndice = "D" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Name = XL.Range(strIndice).Value
        
        strIndice = "E" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Aberto_em = XL.Range(strIndice).Value
        
        strIndice = "F" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Access = XL.Range(strIndice).Value
        
        strIndice = "G" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Alteracao = XL.Range(strIndice).Value
        
        strIndice = "H" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Check_If_Resolved = XL.Range(strIndice).Value
        
        strIndice = "I" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Comite = XL.Range(strIndice).Value
        
        strIndice = "J" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Data_Limite = XL.Range(strIndice).Value
        
        strIndice = "K" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Data_Fechamento = XL.Range(strIndice).Value
        
        strIndice = "L" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Descricao = XL.Range(strIndice).Value
        
        strIndice = "M" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Descricao = CR_Descricao & XL.Range(strIndice).Value
        
        strIndice = "N" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Descricao = CR_Descricao & XL.Range(strIndice).Value
        
        strIndice = "O" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Descricao = CR_Descricao & XL.Range(strIndice).Value
        
        CR_Descricao = TrocarCarEsp_Enter(CR_Descricao)
        
        strIndice = "P" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Equipe = XL.Range(strIndice).Value
        
        strIndice = "Q" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Estimativa_Horas = XL.Range(strIndice).Value
        
        strIndice = "R" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Limite = XL.Range(strIndice).Value
        
        strIndice = "S" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Modulo = XL.Range(strIndice).Value
        
        strIndice = "T" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Overall_Status = XL.Range(strIndice).Value
        
        strIndice = "U" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Percent_Complete = XL.Range(strIndice).Value
        
        strIndice = "V" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Priority = XL.Range(strIndice).Value
        
        strIndice = "W" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Release = XL.Range(strIndice).Value
        
        strIndice = "X" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Resolution = XL.Range(strIndice).Value
        
        strIndice = "Y" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Resolution = CR_Resolution & XL.Range(strIndice).Value
        
        strIndice = "Z" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Resolution = CR_Resolution & XL.Range(strIndice).Value
        
        strIndice = "AA" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Resolution = CR_Resolution & XL.Range(strIndice).Value
        
        CR_Resolution = TrocarCarEsp_Enter(CR_Resolution)
        
        
        strIndice = "AB" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Revision_No = XL.Range(strIndice).Value
        
        strIndice = "AC" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Site = XL.Range(strIndice).Value
        
        strIndice = "AD" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Usuario = XL.Range(strIndice).Value
        
        strIndice = "AE" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Workflow = XL.Range(strIndice).Value
        
        strIndice = "AF" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Identificador = XL.Range(strIndice).Value
        
        strIndice = "AG" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Requisitante = XL.Range(strIndice).Value
        
        strIndice = "AH" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_PMO = XL.Range(strIndice).Value
        
        strIndice = "AI" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Analista = XL.Range(strIndice).Value
        
        strIndice = "AJ" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Revisor = XL.Range(strIndice).Value
        
        strIndice = "AK" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Aprovador = XL.Range(strIndice).Value
        
        strIndice = "AL" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Comments = XL.Range(strIndice).Value
       
        strIndice = "AM" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Comments = CR_Comments & XL.Range(strIndice).Value
       
        strIndice = "AN" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Comments = CR_Comments & XL.Range(strIndice).Value
       
        strIndice = "AO" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Comments = CR_Comments & XL.Range(strIndice).Value
       
        strIndice = "AP" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Comments = CR_Comments & XL.Range(strIndice).Value
       
        strIndice = "AQ" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Comments = CR_Comments & XL.Range(strIndice).Value
       
        strIndice = "AR" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Comments = CR_Comments & XL.Range(strIndice).Value
       
        strIndice = "AS" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Comments = CR_Comments & XL.Range(strIndice).Value
       
        strIndice = "AT" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Comments = CR_Comments & XL.Range(strIndice).Value
       
        strIndice = "AU" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Comments = CR_Comments & XL.Range(strIndice).Value
       
        CR_Comments = TrocarCarEsp_Enter(CR_Comments)

        strIndice = "AV" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Activity_Log = XL.Range(strIndice).Value
       
        strIndice = "AW" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Activity_Log = CR_Activity_Log & XL.Range(strIndice).Value
        
        strIndice = "AX" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Activity_Log = CR_Activity_Log & XL.Range(strIndice).Value
       
        strIndice = "AY" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Activity_Log = CR_Activity_Log & XL.Range(strIndice).Value
       
        CR_Activity_Log = TrocarCarEsp_Enter(CR_Activity_Log)
        
        strIndice = "AZ" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Traceability = XL.Range(strIndice).Value
       
        strIndice = "BA" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Traceability = CR_Traceability & XL.Range(strIndice).Value
       
        strIndice = "BB" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Traceability = CR_Traceability & XL.Range(strIndice).Value
       
        strIndice = "BC" + Trim(CStr(intContador))
        strMensagem = "Chang Requests - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        CR_Traceability = CR_Traceability & XL.Range(strIndice).Value
        
        CR_Traceability = TrocarCarEsp_Enter(CR_Traceability)
        
        strMensagem = "Chang Requests - Nao foi possivel incluir os dados no BD."
        
        strSql = "SP_INCLUIR_CHANG_REQUESTS "
        
        strSql = strSql & "'" & CR_Project_Code & "', "
        strSql = strSql & "" & CR_Sr_No & ", "
        strSql = strSql & "'" & CR_ID & "', "
        strSql = strSql & "'" & gsReplicaPlics(CR_Name) & "', "
        strSql = strSql & "'" & Mid(Format(CR_Aberto_em, "dd/mm/yyyy"), 7, 4) & Mid(Format(CR_Aberto_em, "dd/mm/yyyy"), 4, 2) & Mid(Format(CR_Aberto_em, "dd/mm/yyyy"), 1, 2) & "', "
        strSql = strSql & "'" & CR_Access & "', "
        strSql = strSql & "'" & CR_Alteracao & "', "
        strSql = strSql & "'" & CR_Check_If_Resolved & "', "
        strSql = strSql & "'" & CR_Comite & "', "
        strSql = strSql & "'" & Mid(Format(CR_Data_Limite, "dd/mm/yyyy"), 7, 4) & Mid(Format(CR_Data_Limite, "dd/mm/yyyy"), 4, 2) & Mid(Format(CR_Data_Limite, "dd/mm/yyyy"), 1, 2) & "', "
        strSql = strSql & "'" & Mid(Format(CR_Data_Fechamento, "dd/mm/yyyy"), 7, 4) & Mid(Format(CR_Data_Fechamento, "dd/mm/yyyy"), 4, 2) & Mid(Format(CR_Data_Fechamento, "dd/mm/yyyy"), 1, 2) & "', "
        strSql = strSql & "'" & gsReplicaPlics(CR_Descricao) & "', "
        strSql = strSql & "'" & CR_Equipe & "', "
        strSql = strSql & "" & CR_Estimativa_Horas & ", "
        strSql = strSql & "'" & CR_Limite & "', "
        strSql = strSql & "'" & CR_Modulo & "', "
        strSql = strSql & "'" & CR_Overall_Status & "', "
        strSql = strSql & "" & CR_Percent_Complete & ", "
        strSql = strSql & "'" & CR_Priority & "', "
        strSql = strSql & "'" & CR_Release & "', "
        strSql = strSql & "'" & gsReplicaPlics(CR_Resolution) & "', "
        strSql = strSql & "" & CR_Revision_No & ", "
        strSql = strSql & "'" & gsReplicaPlics(CR_Site) & "', "
        strSql = strSql & "'" & gsReplicaPlics(CR_Usuario) & "', "
        strSql = strSql & "'" & CR_Workflow & "', "
        strSql = strSql & "'" & gsReplicaPlics(CR_Identificador) & "', "
        strSql = strSql & "'" & gsReplicaPlics(CR_Requisitante) & "', "
        strSql = strSql & "'" & gsReplicaPlics(CR_PMO) & "', "
        strSql = strSql & "'" & gsReplicaPlics(CR_Analista) & "', "
        strSql = strSql & "'" & gsReplicaPlics(CR_Revisor) & "', "
        strSql = strSql & "'" & gsReplicaPlics(CR_Aprovador) & "',"
        strSql = strSql & "'" & gsReplicaPlics(CR_Comments) & "', "
        strSql = strSql & "'" & gsReplicaPlics(CR_Activity_Log) & "', "
        strSql = strSql & "'" & gsReplicaPlics(CR_Traceability) & "'"
        
        dbConexao.Execute strSql
        
        intContador = intContador + 1
        
        pProgressBarPlanilha.Value = intContador
    
        strIndice = "A" + Trim(CStr(intContador))
    
    Loop
   
    ' Fecha a planilha e libera o objeto
    Call XL.Workbooks.Close
    
    Set XL = Nothing

    Incluir_Chang_Requests = True
    
    Exit Function
    
ErrChang_Requests:
    
    If Not (XL Is Nothing) Then
        Call XL.Workbooks.Close
        Set XL = Nothing
    End If
    
    Set XL = Nothing
    
    Incluir_Chang_Requests = False

End Function




'Private Function Incluir_Meeting(ByVal strArqXLS As String, _
'                                 ByRef pProgressBarPlanilha As Object) As Boolean
'
''Variaveis Meeting
'
'Dim MT_Project_Code As String
'Dim MT_Sr_No As Integer
'Dim MT_ID As String
'Dim MT_Title As String
'Dim MT_Access As String
'Dim MT_Agenda_Reuniao As String
'Dim MT_Author As String
'Dim MT_Comite As String
'Dim MT_Data_Reuniao As Date
'Dim MT_Description As String
'Dim MT_Equipe As String
'Dim MT_Hora_Reuniao As Date
'Dim MT_Local As String
'Dim MT_Modulo As String
'Dim MT_Open_Items As String
'Dim MT_Priority As String
'Dim MT_Revision_No As Integer
'Dim MT_Total_Items As Integer
'Dim MT_Type As String
'Dim MT_Autor As String
'Dim MT_Aprovador As String
'Dim MT_Comments As String
'Dim MT_Activity_Log As String
'Dim MT_Traceability As String
'
'Dim XL As Excel.Application
'Dim intContador As Integer
'Dim strIndice As String
'Dim strSql As String
'
'Incluir_Meeting = False
'
'On Error GoTo ErrMeeting
'
'
'    strMensagem = "Meeting - Nao foi possivel criar objeto do excel."
'
'    Set XL = CreateObject("Excel.Application")
'
'    strMensagem = "Meeting - Nao foi possivel abrir arquivo do excel."
'
'    Call XL.Workbooks.Open(strArqXLS)
'
'    intContador = 2
'
'    With pProgressBarPlanilha
'        .Min = intContador
'        .Value = intContador
'        .Max = (XL.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row) + 1
'    End With
'
'    strIndice = "A" + Trim(CStr(intContador))
'
'    ' Enquanto existirem dados na planilha
'    Do While (XL.Range(strIndice).Value <> "")
'
'
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Project_Code = XL.Range(strIndice).Value
'
'        strIndice = "B" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Sr_No = XL.Range(strIndice).Value
'
'        strIndice = "C" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_ID = XL.Range(strIndice).Value
'
'        strIndice = "D" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Title = XL.Range(strIndice).Value
'
'        strIndice = "E" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Access = XL.Range(strIndice).Value
'
'        strIndice = "F" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
''**********************************************************
'
'        MT_Agenda_Reuniao = XL.Range(strIndice).Value
'
'        strIndice = "G" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Agenda_Reuniao = MT_Agenda_Reuniao & XL.Range(strIndice).Value
'
'        strIndice = "H" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Agenda_Reuniao = MT_Agenda_Reuniao & XL.Range(strIndice).Value
'
'        strIndice = "I" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Agenda_Reuniao = MT_Agenda_Reuniao & XL.Range(strIndice).Value
'
'        MT_Agenda_Reuniao = TrocarCarEsp_Enter(MT_Agenda_Reuniao)
'
''*******************************************************
'        strIndice = "J" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Author = XL.Range(strIndice).Value
'
'        strIndice = "K" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Comite = XL.Range(strIndice).Value
'
'        strIndice = "L" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Data_Reuniao = XL.Range(strIndice).Value
'
'        strIndice = "M" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'
''********************************************************
'
'        MT_Description = XL.Range(strIndice).Value
'
'        strIndice = "N" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Description = MT_Description & XL.Range(strIndice).Value
'
'        strIndice = "O" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Description = MT_Description & XL.Range(strIndice).Value
'
'        strIndice = "P" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Description = MT_Description & XL.Range(strIndice).Value
'
'        MT_Description = TrocarCarEsp_Enter(MT_Description)
'
''********************************************************
'        strIndice = "Q" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Equipe = XL.Range(strIndice).Value
'
'        strIndice = "R" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Hora_Reuniao = XL.Range(strIndice).Value
'
'        strIndice = "S" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Local = XL.Range(strIndice).Value
'
'        strIndice = "T" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Modulo = XL.Range(strIndice).Value
'
'        strIndice = "U" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Open_Items = XL.Range(strIndice).Value
'
'        strIndice = "V" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Priority = XL.Range(strIndice).Value
'
'        strIndice = "W" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Revision_No = XL.Range(strIndice).Value
'
'        strIndice = "X" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Total_Items = XL.Range(strIndice).Value
'
'        strIndice = "Y" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Type = XL.Range(strIndice).Value
'
'        strIndice = "Z" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Autor = XL.Range(strIndice).Value
'
'        strIndice = "AA" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Aprovador = XL.Range(strIndice).Value
'
'        strIndice = "AB" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
''**************************************************
'
'        MT_Comments = XL.Range(strIndice).Value
'
'        strIndice = "AC" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Comments = MT_Comments & XL.Range(strIndice).Value
'
'        strIndice = "AD" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Comments = MT_Comments & XL.Range(strIndice).Value
'
'        strIndice = "AE" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Comments = MT_Comments & XL.Range(strIndice).Value
'
'        strIndice = "AF" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Comments = MT_Comments & XL.Range(strIndice).Value
'
'        strIndice = "AG" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Comments = MT_Comments & XL.Range(strIndice).Value
'
'        strIndice = "AH" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Comments = MT_Comments & XL.Range(strIndice).Value
'
'        strIndice = "AI" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Comments = MT_Comments & XL.Range(strIndice).Value
'
'        strIndice = "AJ" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Comments = MT_Comments & XL.Range(strIndice).Value
'
'        strIndice = "AK" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Comments = MT_Comments & XL.Range(strIndice).Value
'
'        MT_Comments = TrocarCarEsp_Enter(MT_Comments)
'
'
'
''***************************************************************
'
'
'
'        strIndice = "AL" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Activity_Log = XL.Range(strIndice).Value
'
'        strIndice = "AM" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Activity_Log = MT_Activity_Log & XL.Range(strIndice).Value
'
'        strIndice = "AN" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Activity_Log = MT_Activity_Log & XL.Range(strIndice).Value
'
'        strIndice = "AO" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Activity_Log = MT_Activity_Log & XL.Range(strIndice).Value
'
'        MT_Activity_Log = TrocarCarEsp_Enter(MT_Activity_Log)
'
'
''******************************************************
'
'        strIndice = "AP" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Traceability = XL.Range(strIndice).Value
'
'        strIndice = "AQ" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Traceability = MT_Traceability & XL.Range(strIndice).Value
'
'        strIndice = "AR" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Traceability = MT_Traceability & XL.Range(strIndice).Value
'
'        strIndice = "AS" + Trim(CStr(intContador))
'        strMensagem = "Meeting - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
'
'        MT_Traceability = MT_Traceability & XL.Range(strIndice).Value
'
'        MT_Traceability = TrocarCarEsp_Enter(MT_Traceability)
'
'
''******************************************************
'
'        strMensagem = "Meeting - Nao foi possivel incluir os dados no BD."
'
'        strSql = "SP_INCLUIR_MEETING "
'
'        strSql = strSql & "'" & MT_Project_Code & "', "
'        strSql = strSql & "" & MT_Sr_No & ", "
'        strSql = strSql & "'" & MT_ID & "', "
'        strSql = strSql & "'" & gsReplicaPlics(MT_Title) & "', "
'        strSql = strSql & "'" & MT_Access & "', "
'        strSql = strSql & "'" & gsReplicaPlics(MT_Agenda_Reuniao) & "', "
'        strSql = strSql & "'" & gsReplicaPlics(MT_Author) & "', "
'        strSql = strSql & "'" & MT_Comite & "', "
'        strSql = strSql & "'" & Mid(Format(MT_Data_Reuniao, "dd/mm/yyyy"), 7, 4) & Mid(Format(MT_Data_Reuniao, "dd/mm/yyyy"), 4, 2) & Mid(Format(MT_Data_Reuniao, "dd/mm/yyyy"), 1, 2) & "', "
'        strSql = strSql & "'" & gsReplicaPlics(MT_Description) & "', "
'        strSql = strSql & "'" & MT_Equipe & "', "
'        strSql = strSql & "'" & MT_Hora_Reuniao & "', "
'        strSql = strSql & "'" & MT_Local & "', "
'        strSql = strSql & "'" & MT_Modulo & "', "
'        strSql = strSql & "" & MT_Open_Items & ", "
'        strSql = strSql & "'" & MT_Priority & "', "
'        strSql = strSql & "" & MT_Revision_No & ", "
'        strSql = strSql & "" & MT_Total_Items & ", "
'        strSql = strSql & "'" & MT_Type & "', "
'        strSql = strSql & "'" & gsReplicaPlics(MT_Autor) & "', "
'        strSql = strSql & "'" & gsReplicaPlics(MT_Aprovador) & "', "
'        strSql = strSql & "'" & gsReplicaPlics(MT_Comments) & "', "
'        strSql = strSql & "'" & gsReplicaPlics(MT_Activity_Log) & "',"
'        strSql = strSql & "'" & gsReplicaPlics(MT_Traceability) & "'"
'
'        dbConexao.Execute strSql
'
'        intContador = intContador + 1
'
'        pProgressBarPlanilha.Value = intContador
'
'        strIndice = "A" + Trim(CStr(intContador))
'
'    Loop
'
'    ' Fecha a planilha e libera o objeto
'    Call XL.Workbooks.Close
'
'    Set XL = Nothing
'
'    Incluir_Meeting = True
'
'    Exit Function
'
'ErrMeeting:
'
'    If Not (XL Is Nothing) Then
'        Call XL.Workbooks.Close
'        Set XL = Nothing
'    End If
'
'    Set XL = Nothing
'
'    Incluir_Meeting = False
'
'End Function

Private Function Incluir_Risks(ByVal strArqXLS As String, _
                               ByRef pProgressBarPlanilha As Object) As Boolean

'Variaveis Risks

Dim RK_Project_Code             As String
Dim RK_Sr_No                    As Integer
Dim RK_ID                       As String
Dim RK_Nome                     As String
Dim RK_Access                   As String
Dim RK_Area_Impacto             As String
Dim RK_Categoria                As String
Dim RK_Complexidade             As String
Dim RK_Aberto_em                As Date
Dim RK_Data_Limite              As Date
Dim RK_Data_Fechamento          As Date
Dim RK_Descricao                As String
Dim RK_Estimativa_Horas         As Integer
Dim RK_Fase_Gestão_Risco        As String
Dim RK_Impacto                  As String
Dim RK_Mitigation               As String
Dim RK_Mitigation_Completed     As String
Dim RK_Modulo                   As String
Dim RK_Overall_Status           As String
Dim RK_Percent_Complete         As Double
Dim RK_Priority                 As String
Dim RK_Probabilidade            As String
Dim RK_Ranking                  As String
Dim RK_Equipe                   As String
Dim RK_Usuario                  As String
Dim RK_Workflow                 As String
Dim RK_Identificador            As String
Dim RK_Designador               As String
Dim RK_Responsavel              As String
Dim RK_Integrador               As String
Dim RK_Revisor                  As String
Dim RK_Aprovador                As String
Dim RK_Comments                 As String
Dim RK_Traceability             As String
Dim RK_Activity_Log             As String
Dim RK_Comite                   As String
Dim RK_Fase_Gestao_Risco        As String
Dim RK_Tipo_Risco               As String
Dim RK_Dashboard                As String

Dim XL As Excel.Application
Dim intContador As Integer
Dim strIndice As String
Dim strSql As String

Incluir_Risks = False

On Error GoTo ErrRisks

    strMensagem = "Risks - Nao foi possivel criar objeto do excel."

    Set XL = CreateObject("Excel.Application")

    strMensagem = "Risks - Nao foi possivel abrir arquivo do excel."

    Call XL.Workbooks.Open(strArqXLS)
   
    intContador = 2
    
    With pProgressBarPlanilha
        .Min = intContador
        .Value = intContador
        .Max = (XL.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row) + 1
    End With
   
    strIndice = "A" + Trim(CStr(intContador))
    
    ' Enquanto existirem dados na planilha
    Do While (XL.Range(strIndice).Value <> "")
       
       
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
       
        RK_Project_Code = XL.Range(strIndice).Value
        
        strIndice = "B" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Sr_No = XL.Range(strIndice).Value
        
        strIndice = "C" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        If Len(Trim(XL.Range(strIndice).Value)) = 4 Then
            RK_ID = Left(Trim(XL.Range(strIndice).Value), 3) & "0" & Right(Trim(XL.Range(strIndice).Value), 1)
        Else
            RK_ID = XL.Range(strIndice).Value
        End If
        
        
        strIndice = "D" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Nome = XL.Range(strIndice).Value
        
        strIndice = "E" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Access = XL.Range(strIndice).Value
        
        strIndice = "F" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Area_Impacto = XL.Range(strIndice).Value
        
        strIndice = "G" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Categoria = XL.Range(strIndice).Value
        
        strIndice = "H" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Complexidade = XL.Range(strIndice).Value
        
        strIndice = "I" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Aberto_em = XL.Range(strIndice).Value
        
        strIndice = "J" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Data_Limite = XL.Range(strIndice).Value
        
        strIndice = "K" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Data_Fechamento = XL.Range(strIndice).Value
        
        strIndice = "L" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
'*****************************************************

        RK_Descricao = XL.Range(strIndice).Value
        
        strIndice = "M" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Descricao = RK_Descricao & XL.Range(strIndice).Value
        
        strIndice = "N" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Descricao = RK_Descricao & XL.Range(strIndice).Value
        
        strIndice = "O" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Descricao = RK_Descricao & XL.Range(strIndice).Value
        
        RK_Descricao = TrocarCarEsp_Enter(RK_Descricao)
        
        
'******************************************************
        
        strIndice = "P" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Estimativa_Horas = XL.Range(strIndice).Value
        
        strIndice = "Q" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Fase_Gestao_Risco = XL.Range(strIndice).Value
        
        strIndice = "R" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Impacto = XL.Range(strIndice).Value
        
        strIndice = "S" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."


'*************************************************
        
        RK_Mitigation = XL.Range(strIndice).Value
        
        strIndice = "T" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Mitigation = RK_Mitigation & XL.Range(strIndice).Value
        
        strIndice = "U" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Mitigation = RK_Mitigation & XL.Range(strIndice).Value
        
        strIndice = "V" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Mitigation = RK_Mitigation & XL.Range(strIndice).Value
        
        RK_Mitigation = TrocarCarEsp_Enter(RK_Mitigation)
        
        
'******************************************************
        
        
        strIndice = "W" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Mitigation_Completed = XL.Range(strIndice).Value
        
        strIndice = "X" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Modulo = XL.Range(strIndice).Value
        
        strIndice = "Y" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Overall_Status = XL.Range(strIndice).Value
        
        strIndice = "Z" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Percent_Complete = XL.Range(strIndice).Value
        
        strIndice = "AA" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Priority = XL.Range(strIndice).Value
        
        strIndice = "AB" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Probabilidade = XL.Range(strIndice).Value
        
        strIndice = "AC" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Ranking = XL.Range(strIndice).Value
        
        strIndice = "AD" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Equipe = XL.Range(strIndice).Value
        
        strIndice = "AE" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Usuario = XL.Range(strIndice).Value
        
        strIndice = "AF" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Workflow = XL.Range(strIndice).Value
        
        strIndice = "AG" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Identificador = XL.Range(strIndice).Value
        
        strIndice = "AH" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Designador = XL.Range(strIndice).Value
        
        strIndice = "AI" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Responsavel = XL.Range(strIndice).Value
        
        strIndice = "AJ" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Integrador = XL.Range(strIndice).Value
        
        strIndice = "AK" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Revisor = XL.Range(strIndice).Value
        
        strIndice = "AL" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Aprovador = XL.Range(strIndice).Value
       
        strIndice = "AM" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
'**************************************************


        RK_Comments = XL.Range(strIndice).Value
        
        strIndice = "AN" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Comments = RK_Comments & XL.Range(strIndice).Value
        
        strIndice = "AO" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Comments = RK_Comments & XL.Range(strIndice).Value
        
        strIndice = "AP" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Comments = RK_Comments & XL.Range(strIndice).Value
        
        strIndice = "AQ" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Comments = RK_Comments & XL.Range(strIndice).Value
        
        strIndice = "AR" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Comments = RK_Comments & XL.Range(strIndice).Value
        
        strIndice = "AS" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Comments = RK_Comments & XL.Range(strIndice).Value
        
        strIndice = "AT" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Comments = RK_Comments & XL.Range(strIndice).Value
        
        strIndice = "AU" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Comments = RK_Comments & XL.Range(strIndice).Value
        
        strIndice = "AV" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Comments = RK_Comments & XL.Range(strIndice).Value
        
        RK_Comments = TrocarCarEsp_Enter(RK_Comments)
        
        
'******************************************************
        
        
        
        strIndice = "AW" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Traceability = XL.Range(strIndice).Value
        
        strIndice = "AX" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Traceability = RK_Traceability & XL.Range(strIndice).Value
        
        strIndice = "AY" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Traceability = RK_Traceability & XL.Range(strIndice).Value
        
        strIndice = "AZ" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Traceability = RK_Traceability & XL.Range(strIndice).Value
        
        
        RK_Traceability = TrocarCarEsp_Enter(RK_Traceability)
        
        
'******************************************************
        
        
        
        strIndice = "BA" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Activity_Log = XL.Range(strIndice).Value
        
        strIndice = "BB" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Activity_Log = RK_Activity_Log & XL.Range(strIndice).Value
        
        strIndice = "BC" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Activity_Log = RK_Activity_Log & XL.Range(strIndice).Value
        
        strIndice = "BD" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Activity_Log = RK_Activity_Log & XL.Range(strIndice).Value
        
        
        RK_Activity_Log = TrocarCarEsp_Enter(RK_Activity_Log)
        
        
'******************************************************
        
        
        strIndice = "BE" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Comite = XL.Range(strIndice).Value
       
        strIndice = "BF" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Tipo_Risco = XL.Range(strIndice).Value
        
        strIndice = "BG" + Trim(CStr(intContador))
        strMensagem = "Risks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        RK_Dashboard = XL.Range(strIndice).Value
        
        
        strMensagem = "Risks - Nao foi possivel incluir os dados no BD."
        
        strSql = "SP_INCLUIR_RISKS "
        
        strSql = strSql & "'" & RK_Project_Code & "', "
        strSql = strSql & "" & RK_Sr_No & ", "
        strSql = strSql & "'" & RK_ID & "', "
        strSql = strSql & "'" & gsReplicaPlics(RK_Nome) & "', "
        strSql = strSql & "'" & RK_Access & "', "
        strSql = strSql & "'" & RK_Area_Impacto & "', "
        strSql = strSql & "'" & RK_Categoria & "', "
        strSql = strSql & "'" & RK_Complexidade & "', "
        strSql = strSql & "'" & Mid(Format(RK_Aberto_em, "dd/mm/yyyy"), 7, 4) & Mid(Format(RK_Aberto_em, "dd/mm/yyyy"), 4, 2) & Mid(Format(RK_Aberto_em, "dd/mm/yyyy"), 1, 2) & "', "
        strSql = strSql & "'" & Mid(Format(RK_Data_Limite, "dd/mm/yyyy"), 7, 4) & Mid(Format(RK_Data_Limite, "dd/mm/yyyy"), 4, 2) & Mid(Format(RK_Data_Limite, "dd/mm/yyyy"), 1, 2) & "', "
        strSql = strSql & "'" & Mid(Format(RK_Data_Fechamento, "dd/mm/yyyy"), 7, 4) & Mid(Format(RK_Data_Fechamento, "dd/mm/yyyy"), 4, 2) & Mid(Format(RK_Data_Fechamento, "dd/mm/yyyy"), 1, 2) & "', "
        strSql = strSql & "'" & gsReplicaPlics(RK_Descricao) & "', "
        strSql = strSql & "" & RK_Estimativa_Horas & ", "
        strSql = strSql & "'" & RK_Fase_Gestão_Risco & "', "
        strSql = strSql & "'" & RK_Impacto & "', "
        strSql = strSql & "'" & gsReplicaPlics(RK_Mitigation) & "', "
        strSql = strSql & "'" & RK_Mitigation_Completed & "', "
        strSql = strSql & "'" & RK_Modulo & "', "
        strSql = strSql & "'" & RK_Overall_Status & "', "
        strSql = strSql & "" & RK_Percent_Complete & ", "
        strSql = strSql & "'" & RK_Priority & "', "
        strSql = strSql & "'" & RK_Probabilidade & "', "
        strSql = strSql & "'" & RK_Ranking & "', "
        strSql = strSql & "'" & RK_Equipe & "', "
        strSql = strSql & "'" & gsReplicaPlics(RK_Usuario) & "', "
        strSql = strSql & "'" & RK_Workflow & "', "
        strSql = strSql & "'" & gsReplicaPlics(RK_Identificador) & "', "
        strSql = strSql & "'" & gsReplicaPlics(RK_Designador) & "', "
        strSql = strSql & "'" & gsReplicaPlics(RK_Responsavel) & "', "
        strSql = strSql & "'" & gsReplicaPlics(RK_Integrador) & "', "
        strSql = strSql & "'" & gsReplicaPlics(RK_Revisor) & "', "
        strSql = strSql & "'" & gsReplicaPlics(RK_Aprovador) & "', "
        strSql = strSql & "'" & gsReplicaPlics(RK_Comments) & "', "
        strSql = strSql & "'" & gsReplicaPlics(RK_Traceability) & "', "
        strSql = strSql & "'" & gsReplicaPlics(RK_Activity_Log) & "', "
        strSql = strSql & "'" & RK_Comite & "', "
        strSql = strSql & "'" & RK_Tipo_Risco & "', "
        strSql = strSql & "'" & RK_Dashboard & "'"
        
        dbConexao.Execute strSql
        
        intContador = intContador + 1
        
        pProgressBarPlanilha.Value = intContador
    
        strIndice = "A" + Trim(CStr(intContador))
    
    Loop
   
    ' Fecha a planilha e libera o objeto
    Call XL.Workbooks.Close
    
    Set XL = Nothing
    
    Incluir_Risks = True
    
    Exit Function
    
ErrRisks:

    
    If Not (XL Is Nothing) Then
        Call XL.Workbooks.Close
        Set XL = Nothing
    End If
    
    Set XL = Nothing

    Incluir_Risks = False


End Function



Private Function Incluir_Tasks(ByVal strArqXLS As String, _
                               ByRef pProgressBarPlanilha As Object) As Boolean

'Variaveis Tasks

Dim TS_Project_Code             As String
Dim TS_ID                       As String
Dim TS_Nome                     As String
Dim TS_Planned_Finish           As Date
Dim TS_Overall_Status           As String

Dim XL As Excel.Application
Dim intContador As Integer
Dim strIndice As String
Dim strSql As String

Incluir_Tasks = False

On Error GoTo ErrTasks

    strMensagem = "Tasks - Nao foi possivel criar objeto do excel."

    Set XL = CreateObject("Excel.Application")

    strMensagem = "Tasks - Nao foi possivel abrir arquivo do excel."

    Call XL.Workbooks.Open(strArqXLS)

    intContador = 2
    
    With pProgressBarPlanilha
        .Min = intContador
        .Value = intContador
        .Max = (XL.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row) + 1
    End With

    strIndice = "A" + Trim(CStr(intContador))
    
    ' Enquanto existirem dados na planilha
    Do While (XL.Range(strIndice).Value <> "")
     
        strMensagem = "Tasks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
     
        TS_Project_Code = XL.Range(strIndice).Value
        
        strIndice = "B" + Trim(CStr(intContador))
        strMensagem = "Tasks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        If Len(Trim(XL.Range(strIndice).Value)) = 4 Then
            TS_ID = Left(Trim(XL.Range(strIndice).Value), 3) & "0" & Right(Trim(XL.Range(strIndice).Value), 1)
        Else
            TS_ID = XL.Range(strIndice).Value
        End If
        
        strIndice = "C" + Trim(CStr(intContador))
        strMensagem = "Tasks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        TS_Nome = XL.Range(strIndice).Value
        
        strIndice = "D" + Trim(CStr(intContador))
        strMensagem = "Tasks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        TS_Planned_Finish = XL.Range(strIndice).Value
        
        strIndice = "E" + Trim(CStr(intContador))
        strMensagem = "Tasks - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        TS_Overall_Status = XL.Range(strIndice).Value
       
       
        strMensagem = "Tasks - Nao foi possivel incluir os dados no BD."
       
        strSql = "SP_INCLUIR_TASKS "
        strSql = strSql & "'" & TS_Project_Code & "', "
        strSql = strSql & "'" & TS_ID & "', "
        strSql = strSql & "'" & gsReplicaPlics(TS_Nome) & "', "
        strSql = strSql & "'" & Mid(Format(TS_Planned_Finish, "dd/mm/yyyy"), 7, 4) & Mid(Format(TS_Planned_Finish, "dd/mm/yyyy"), 4, 2) & Mid(Format(TS_Planned_Finish, "dd/mm/yyyy"), 1, 2) & "', "
        strSql = strSql & "'" & TS_Overall_Status & "'"
        
        dbConexao.Execute strSql
        
        intContador = intContador + 1
        
        pProgressBarPlanilha.Value = intContador
    
        strIndice = "A" + Trim(CStr(intContador))
    
    Loop
   
    ' Fecha a planilha e libera o objeto
    Call XL.Workbooks.Close

    Set XL = Nothing
 
    Incluir_Tasks = True

    Exit Function
    
ErrTasks:
    
    If Not (XL Is Nothing) Then
        Call XL.Workbooks.Close
        Set XL = Nothing
    End If
    
    Set XL = Nothing

    Incluir_Tasks = False
    
    
End Function

Private Function Incluir_Traceability(ByVal strArqXLS As String, _
                                      ByRef pProgressBarPlanilha As Object) As Boolean

'Variaveis Actions

Dim OR_Project_Code As String
Dim OR_ID As String
Dim DS_Project_Code As String
Dim DS_ID As String

Dim XL As Excel.Application
Dim intContador As Integer
Dim strIndice As String
Dim strSql As String

Incluir_Traceability = False

On Error GoTo ErrTraceability

    strMensagem = "Traceability - Nao foi possivel criar objeto do excel."

    Set XL = CreateObject("Excel.Application")

    strMensagem = "Traceability - Nao foi possivel abrir arquivo do excel."

    Call XL.Workbooks.Open(strArqXLS)
   
    intContador = 2
    
    With pProgressBarPlanilha
        .Min = intContador
        .Value = intContador
        .Max = (XL.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row) + 1
    End With
   
    strIndice = "A" + Trim(CStr(intContador))
    
    ' Enquanto existirem dados na planilha
    Do While (XL.Range(strIndice).Value <> "")
      
        strMensagem = "Traceability - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
      
        OR_Project_Code = XL.Range(strIndice).Value
        
        strIndice = "B" + Trim(CStr(intContador))
        strMensagem = "Traceability - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        If Len(Trim(XL.Range(strIndice).Value)) = 4 Then
            OR_ID = Left(Trim(XL.Range(strIndice).Value), 3) & "0" & Right(Trim(XL.Range(strIndice).Value), 1)
        Else
            OR_ID = XL.Range(strIndice).Value
        End If
        
        strIndice = "C" + Trim(CStr(intContador))
        strMensagem = "Traceability - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        DS_Project_Code = XL.Range(strIndice).Value
        
        strIndice = "D" + Trim(CStr(intContador))
        strMensagem = "Traceability - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
       
        If Len(Trim(XL.Range(strIndice).Value)) = 4 Then
            DS_ID = Left(Trim(XL.Range(strIndice).Value), 3) & "0" & Right(Trim(XL.Range(strIndice).Value), 1)
        Else
            DS_ID = XL.Range(strIndice).Value
        End If
       
        strMensagem = "Traceability - Nao foi possivel incluir os dados no BD."
        
        strSql = "SP_INCLUIR_TRACEABILITY "
        
        strSql = strSql & "'" & OR_Project_Code & "', "
        strSql = strSql & "'" & OR_ID & "', "
        strSql = strSql & "'" & DS_Project_Code & "', "
        strSql = strSql & "'" & DS_ID & "' "
        
        dbConexao.Execute strSql
        
        intContador = intContador + 1
        
        pProgressBarPlanilha.Value = intContador
    
        strIndice = "A" + Trim(CStr(intContador))
    
    Loop
   
    ' Fecha a planilha e libera o objeto
    Call XL.Workbooks.Close
    
    Set XL = Nothing
    
    Incluir_Traceability = True

    Exit Function
    
ErrTraceability:

    
    If Not (XL Is Nothing) Then
        Call XL.Workbooks.Close
        Set XL = Nothing
    End If
    
    Incluir_Traceability = False

End Function


Public Sub Incluir_Scorecard()
Dim strSql As String

    strSql = "EXEC SP_ATUALIZAR_DADOS_SCORECARD"
    
    dbConexao.Execute strSql

End Sub

Public Sub Incluir_Scorecard_Risks()
Dim strSql As String

    strSql = "EXEC SP_ATUALIZAR_DADOS_SCORECARD_RISKS"
    
    dbConexao.Execute strSql

End Sub

