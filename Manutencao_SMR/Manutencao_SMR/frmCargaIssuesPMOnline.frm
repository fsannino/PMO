VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCargaIssuesPMOnline 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga da planilha Issues"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "frmCargaIssuesPMOnline.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.DriveListBox drvDrivePlanilhasCarga 
      Height          =   315
      Left            =   135
      TabIndex        =   5
      Top             =   435
      Width           =   5025
   End
   Begin VB.DirListBox dirPastaPlanilhasCarga 
      Height          =   1665
      Left            =   135
      TabIndex        =   4
      Top             =   1245
      Width           =   5025
   End
   Begin VB.CommandButton cmdCancelarCarga 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   3735
      TabIndex        =   3
      Top             =   4185
      Width           =   1440
   End
   Begin VB.CommandButton cmdIniciarCarga 
      Caption         =   "Iniciar Carga"
      Height          =   345
      Left            =   2205
      TabIndex        =   2
      Top             =   4185
      Width           =   1440
   End
   Begin VB.Frame fraAndamentoPlanilha 
      Caption         =   "Importando planilha de "
      Height          =   900
      Left            =   90
      TabIndex        =   0
      Top             =   3135
      Width           =   5085
      Begin MSComctlLib.ProgressBar ProgressBarPlanilhas 
         Height          =   345
         Left            =   45
         TabIndex        =   1
         Top             =   345
         Width           =   4980
         _ExtentX        =   8784
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   1
         Min             =   1e-4
      End
   End
   Begin VB.Label lblPasta 
      AutoSize        =   -1  'True
      Caption         =   "Selecione a pasta onde se encontra a planilha:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   930
      Width           =   3330
   End
   Begin VB.Label lblDrive 
      AutoSize        =   -1  'True
      Caption         =   "Selecione o drive onde se encontra a planilha:"
      Height          =   195
      Left            =   105
      TabIndex        =   6
      Top             =   135
      Width           =   3285
   End
End
Attribute VB_Name = "frmCargaIssuesPMOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strMensagem As String

Private Sub mInicializaObjetosCarga()

    fraAndamentoPlanilha.Caption = ""
    
    With ProgressBarPlanilhas
        .Min = 0
        .Max = 100
        .Value = 0
    End With
    
    With drvDrivePlanilhasCarga
        .Drive = "P:\"
    End With
    
    With dirPastaPlanilhasCarga
        .Path = "P:\Petrobras\Projeto\Pmo\Operações\Planilhas PMOnline"
    End With
    
End Sub


Private Function mExistemPlanilhasPMOnline(ByVal pPasta As String) As Boolean

    Dim fso As FileSystemObject
    
    mExistemPlanilhasPMOnline = False
        
    Set fso = New FileSystemObject
    
    If fso.FileExists(pPasta & "IS.xls") Then
       
        mExistemPlanilhasPMOnline = True
        
    End If
    
    Set fso = Nothing

End Function

Private Sub cmdCancelarCarga_Click()

    Unload Me
    
End Sub


Private Sub cmdIniciarCarga_Click()

    Dim strPastaPlanilhas As String
    Dim strMensagem       As String

    On Error GoTo ErrcmdIniciarCarga_Click

    Screen.MousePointer = vbHourglass

    strPastaPlanilhas = gsTrataPath(dirPastaPlanilhasCarga.Path)

    If mExistemPlanilhasPMOnline(strPastaPlanilhas) Then
        strMensagem = ImportarPlanilhaIssuesPMOnline(strPastaPlanilhas, _
                                                    ProgressBarPlanilhas, _
                                                    fraAndamentoPlanilha)
        MsgBox strMensagem, vbInformation + vbOKOnly, App.Title
    Else
        MsgBox "Não foi encontrada pelo menos uma das planilhas necessárias a importação.", vbCritical + vbOKOnly, App.Title
    End If

    Screen.MousePointer = vbNormal

    Exit Sub

ErrcmdIniciarCarga_Click:

    Screen.MousePointer = vbNormal

    MsgBox "Ocorreu o seguinte erro na rotina cmdIniciarCarga_Click: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub drvDrivePlanilhasCarga_Change()

    On Error GoTo ErrdrvDrivePlanilhasCarga_Change
    
    With drvDrivePlanilhasCarga
        dirPastaPlanilhasCarga.Path = .Drive
    End With
    
    Exit Sub
    
ErrdrvDrivePlanilhasCarga_Change:

    MsgBox "Ocorreu o seguinte erro na rotina drvDrivePlanilhasCarga_Change: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
    
End Sub

Private Sub Form_Load()

    'Inicializa os objetos do form
    Call mInicializaObjetosCarga
    
End Sub

Private Function ImportarPlanilhaIssuesPMOnline(ByVal strDiretorio As String, _
                                                ByRef pProgressBarPlanilha As Object, _
                                                ByRef pCaptionFramePlanilha As Object) As String

    Dim strArqXLS As String
    
    On Error GoTo ErrIssues
    
    With pProgressBarPlanilha
        .Min = 0
        .Value = 0
        .Max = 100
    End With
    
    strMensagem = "Nao foi possivel abrir a conexao com o banco."

    'Abre conexao com o Banco SQL-Server

    strMensagem = "Nao foi possivel excluir os dados das tabelas."
    
    'Limpa a base de dados PMO

    Call Excluir_Dados_Issue_Pmo

    'Importa dados da planilha de Issues para base SQL-Server

    pCaptionFramePlanilha.Caption = "IS.XLS"
    
    strArqXLS = strDiretorio & "\IS.XLS"

    If Not Incluir_Issues(strArqXLS, pProgressBarPlanilha) Then
        Call Excluir_Dados_Issue_Pmo
        ImportarPlanilhaIssuesPMOnline = strMensagem
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    strMensagem = "Nao foi possivel incluir Scorecard."
    
    Call Incluir_Scorecard
    
    strMensagem = "Arquivos importados com sucesso."

    ImportarPlanilhaIssuesPMOnline = strMensagem
    
    Screen.MousePointer = vbDefault

    Exit Function

ErrIssues:

        ImportarPlanilhaIssuesPMOnline = strMensagem
        Screen.MousePointer = vbDefault

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
        
        IS_Descricao = Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "K" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Descricao = IS_Descricao & Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "L" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Descricao = IS_Descricao & Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "M" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Descricao = IS_Descricao & Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        
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
        
        IS_Resolution = Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "V" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Resolution = IS_Resolution & Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "W" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Resolution = IS_Resolution & Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "X" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Resolution = IS_Resolution & Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
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
        
        IS_Comments = Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "AH" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Comments = IS_Comments & Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "AI" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Comments = IS_Comments & Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "AJ" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Comments = IS_Comments & Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "AK" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Comments = IS_Comments & Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "AL" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Comments = IS_Comments & Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "AM" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Comments = IS_Comments & Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "AN" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Comments = IS_Comments & Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "AO" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Comments = IS_Comments & Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "AP" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Comments = IS_Comments & Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "AQ" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Traceability = Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "AR" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Traceability = IS_Traceability & Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "AS" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Traceability = IS_Traceability & Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "AT" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Traceability = IS_Traceability & Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "AU" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Activity_Log = Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "AV" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Activity_Log = IS_Activity_Log & Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "AW" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Activity_Log = IS_Activity_Log & Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
        strIndice = "AX" + Trim(CStr(intContador))
        strMensagem = "Issues - Nao foi possivel ler a coluna " & strIndice & " do arquivo excel."
        
        IS_Activity_Log = IS_Activity_Log & Replace(Replace(XL.Range(strIndice).Value, "**", Chr(10) & Chr(13)), "++", Chr(10) & Chr(13))
        
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
        
        dbConexaoSMR.Execute strSql
        
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


Private Sub Excluir_Dados_Issue_Pmo()
       
Dim strSql As String
    
    strSql = "SP_EXCLUIR_DADOS_ISSUES_PMO "
    
    dbConexaoSMR.Execute strSql

End Sub

Public Sub Incluir_Scorecard()
Dim strSql As String

    strSql = "EXEC SP_ATUALIZAR_DADOS_SCORECARD"
    
    dbConexaoSMR.Execute strSql

End Sub

