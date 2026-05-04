VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmExpCausasAtrasadas 
   Caption         =   "Geração do Arquivo de Causas Atrasadas"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   Icon            =   "frmExpCausasAtrasadas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdgGravarPlanilha 
      Left            =   1485
      Top             =   3150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraAndamentoProjetos 
      Caption         =   "Projeto"
      Height          =   915
      Left            =   120
      TabIndex        =   6
      Top             =   1050
      Width           =   5775
      Begin MSComctlLib.ProgressBar prbPlanilha 
         Height          =   390
         Left            =   105
         TabIndex        =   8
         Top             =   315
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   3135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Andamento da criação da planilha"
      Height          =   915
      Left            =   120
      TabIndex        =   5
      Top             =   2190
      Width           =   5775
      Begin MSComctlLib.ProgressBar prbProjetos 
         Height          =   390
         Left            =   105
         TabIndex        =   7
         Top             =   315
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Arquivo"
      Height          =   780
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   5775
      Begin VB.CommandButton cmdProcura 
         Caption         =   "Procurar"
         Height          =   345
         Left            =   4410
         TabIndex        =   4
         Top             =   255
         Width           =   1095
      End
      Begin VB.TextBox txtDiretorio 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdCancela 
      Caption         =   "Sair"
      Height          =   375
      Left            =   4545
      TabIndex        =   1
      Top             =   3225
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   3225
      Width           =   1230
   End
End
Attribute VB_Name = "frmExpCausasAtrasadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    Dim strArq As String
    Dim strDir As String
    
    On Error GoTo ErrcmdOK_Click

    strArq = txtDiretorio.Text

    strDir = Mid(strArq, 1, InStrRev(strArq, "\"))

    If Dir(strDir, vbDirectory) = "" Then
        MsgBox "Diretório não existe.", vbCritical + vbOKOnly, App.Title
        Exit Sub
    End If

    If UCase(Right(strArq, 4)) <> ".XLS" Then
        MsgBox "Arquivo inválido.", vbCritical + vbOKOnly, App.Title
        Exit Sub
    End If

    cmdOK.Enabled = False
    cmdProcura.Enabled = False
    cmdCancela.Enabled = False
    
    If ExportarArqCausasAtrasadas(strArq) Then
        MsgBox "Exportação gerada com sucesso.", vbInformation + vbOKOnly, App.Title
    Else
        MsgBox "Erro na geração do arquivo.", vbCritical + vbOKOnly, App.Title
    End If
    
    cmdOK.Enabled = True
    cmdProcura.Enabled = True
    cmdCancela.Enabled = True
    
    'Inicializando a ProgressBar
    With prbPlanilha
        .Min = 0
        .Value = 0
        .Max = 100
    End With
    
    With prbProjetos
        .Min = 0
        .Value = 0
        .Max = 100
    End With
    
    Exit Sub
    
ErrcmdOK_Click:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina cmdOK_Click: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
    
    'Inicializando a ProgressBar
    With prbPlanilha
        .Min = 0
        .Max = 100
        .Value = 0
    End With
    
    With prbProjetos
        .Min = 0
        .Max = 100
        .Value = 0
    End With
    
    cmdOK.Enabled = True
    cmdProcura.Enabled = True
    cmdCancela.Enabled = True
    
End Sub

Private Sub cmdProcura_Click()

    Dim strDirInicial As String

    On Error GoTo ErrcmdProcura

    strDirInicial = txtDiretorio.Text
    
    With CommonDialog1
    
        .DialogTitle = "Abrir Template Causas Atrasadas"
        .FileName = ""
        .Filter = "Arquivos do MS-Excel (*.xls)|*.xls"
        .CancelError = True
        .ShowOpen
        
        txtDiretorio.Text = .FileName
        
    End With

    Exit Sub

ErrcmdProcura:

    txtDiretorio.Text = strDirInicial

End Sub

Private Sub cmdCancela_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()

    txtDiretorio = gsTrataPath(App.Path) & "Analise dos Cronogramas.xls"
    
End Sub

Private Function ExportarArqCausasAtrasadas(ByVal strArquivo As String) As Boolean

    Dim objExcel                As Excel.Application
    Dim wrkSheetDados           As Excel.Worksheet
    Dim wrkBook                 As Excel.Workbook
    Dim rngSourceRange          As Excel.Range
    Dim rngTargetRange          As Excel.Range
    Dim rsProjetos              As ADODB.Recordset
    Dim rsTarefas               As ADODB.Recordset
    Dim cmdResultado            As ADODB.Command
    Dim strSql                  As String
    Dim strNomeWorkSheet        As String
    Dim strCelula               As String
    Dim intLinha                As Integer
    Dim intContaProjetos        As Integer
    Dim intContaLinhasPlanilha  As Integer
    Dim lngTotalLinhas          As Long
    
    On Error GoTo ErrExportarArqCausasAtrasadas

    ExportarArqCausasAtrasadas = False
    
    Screen.MousePointer = vbHourglass
        
    If Not (objExcel Is Nothing) Then

        'Restaurando o valor original da propriedade Interactive
        objExcel.Interactive = True

        'Fechando o Excel
        wrkBook.Close
        objExcel.Quit

        Set wrkSheetDados = Nothing
        Set wrkBook = Nothing
        Set objExcel = Nothing

        'Criando uma nova instância para o objeto
        Set objExcel = CreateObject("Excel.Application")
        
        'Abrindo o template
        Set wrkBook = objExcel.Workbooks.Open(strArquivo)

    Else

        'Criando uma nova instância para o objeto
        Set objExcel = CreateObject("Excel.Application")
        
        'Abrindo o template
        Set wrkBook = objExcel.Workbooks.Open(strArquivo)

    End If

    'Verificando se conseguiu inicializar o objeto
    If objExcel Is Nothing Then
        Screen.MousePointer = vbNormal
        MsgBox "O aplicativo Excel não pode ser inicializado." & vbCrLf, _
               vbCritical + vbOKOnly, _
               "ERRO"
        End
    End If

    'Vai inibir algumas mensagens enviadas pelo Excel
    objExcel.DisplayAlerts = False
    
    'Desabilita a emissão de sons
    objExcel.EnableSound = False
    
    'Desabilitando o modo interativo
    objExcel.Interactive = False
    
    intContaProjetos = 0
    
    'Recuperando os projetos existentes na base de dados
    strSql = "EXECUTE SP_LISTAR_PROJETOS_PROJECT"

    Set rsProjetos = New ADODB.Recordset

    rsProjetos.CursorLocation = adUseClient
    rsProjetos.Open strSql, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
    
    'Inicializando Progress Bar de Projetos
    
    With prbProjetos
        .Min = 0
        .Max = rsProjetos.RecordCount
        .Value = 0
    End With
    
    Do While Not rsProjetos.EOF
                
        intContaProjetos = intContaProjetos + 1
        
        'Recuperando as tarefas existentes na base de dados
        strSql = "EXECUTE SP_LISTAR_CAUSAS_ATRASADA " & rsProjetos("PROJ_ID")
        
        Set rsTarefas = New ADODB.Recordset
        
        rsTarefas.CursorLocation = adUseClient
        rsTarefas.Open strSql, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
        
        'Atualizando ProgressBar de Projetos
        prbProjetos.Value = intContaProjetos
        
        'Colocando o nome do projeto no frame
        fraAndamentoProjetos.Caption = Trim(rsProjetos("PROJ_NAME"))
        
        'Inicializando Progress Bar de Planilhas
        With prbPlanilha
            .Min = 0
            .Value = 0
            .Max = rsTarefas.RecordCount
        End With
        
        'Atualizando formulário
        Me.Refresh
        
        'Selecionando o Worksheet que será trabalhado
        strNomeWorkSheet = Trim(Replace(rsProjetos("PROJ_NAME"), "Cronograma", "", , , vbTextCompare))
        Set wrkSheetDados = wrkBook.Sheets(strNomeWorkSheet)
        
        'Inicializando contadores
        intLinha = 9
        intContaLinhasPlanilha = 0

        'Atualizando ProgressBar de Planilhas
        'prbPlanilha.Value = intLinha

        Do While Not rsTarefas.EOF
        
            If rsTarefas("TASK_IS_SUMMARY") Then
            
                'Células em negrito
                With wrkSheetDados
                    .Cells(intLinha, 1).Font.Bold = True
                    .Cells(intLinha, 2).Font.Bold = True
                    .Cells(intLinha, 3).Font.Bold = True
                    .Cells(intLinha, 4).Font.Bold = True
                    .Cells(intLinha, 5).Font.Bold = True
                    .Cells(intLinha, 6).Font.Bold = True
                    .Cells(intLinha, 7).Font.Bold = True
                    .Cells(intLinha, 8).Font.Bold = True
                End With
                
            Else
            
                'Células sem negrito
                With wrkSheetDados
                    .Cells(intLinha, 1).Font.Bold = False
                    .Cells(intLinha, 2).Font.Bold = False
                    .Cells(intLinha, 3).Font.Bold = False
                    .Cells(intLinha, 4).Font.Bold = False
                    .Cells(intLinha, 5).Font.Bold = False
                    .Cells(intLinha, 6).Font.Bold = False
                    .Cells(intLinha, 7).Font.Bold = False
                    .Cells(intLinha, 8).Font.Bold = False
                End With
                
            End If
        
            'Carregando a planilha com informações filtradas do BD
            With wrkSheetDados
                
                .Cells(intLinha, 1).Value = rsTarefas("TASK_UID")
                .Cells(intLinha, 2).Value = String(rsTarefas("TASK_OUTLINE_LEVEL"), " ") & rsTarefas("TASK_NAME")
                If rsTarefas("Duracao") > 1 Then
                    .Cells(intLinha, 3).Value = CStr(rsTarefas("Duracao")) & " dias"
                Else
                    .Cells(intLinha, 3).Value = CStr(rsTarefas("Duracao")) & " dia"
                End If
                .Cells(intLinha, 4).Value = CDate(CStr(rsTarefas("TASK_START_DATE")) & " 08:00:00")
                .Cells(intLinha, 5).Value = CDate(CStr(rsTarefas("TASK_FINISH_DATE")) & " 17:00:00")
                .Cells(intLinha, 6).Value = (rsTarefas("TASK_PCT_COMP") / 100)
                .Cells(intLinha, 8).Value = rsTarefas("Causa")
                
            End With
                        
            'Copiando fórmula
            If intLinha > 9 Then
                
                'Só faço a cópia da fórmula se já tiver passado da linha 9
                '(Linha 9 é a primeira linha de dados e nela contém a fórmula
                'que irei copiar para as demais linhas de dados da planilha
                
                strCelula = "G" & intLinha
                
                Set rngSourceRange = wrkSheetDados.Range("G9")
                Set rngTargetRange = wrkSheetDados.Range(strCelula)
                
                rngSourceRange.Copy rngTargetRange
                
                Set rngSourceRange = Nothing
                Set rngTargetRange = Nothing
                
            End If

            intLinha = intLinha + 1
            intContaLinhasPlanilha = intContaLinhasPlanilha + 1

            'Atualizando ProgressBar de Planilhas
            prbPlanilha.Value = intContaLinhasPlanilha

            rsTarefas.MoveNext
        
        Loop
        
        rsTarefas.Close
        Set rsTarefas = Nothing
        
        rsProjetos.MoveNext
        
    Loop
    
    'Selecionando a primeira célula da primeira Sheet
    objExcel.Sheets(1).Cells(1, 1).Select
    
    'Dialog Box para salvar arquivo
    With cdgGravarPlanilha
    
        .DialogTitle = "Salvar planilha como"
        .InitDir = App.Path
        .DefaultExt = ".xls"
        .Filter = "Arquivos do MS-Excel (*.xls)|*.xls"
        .CancelError = True
        
        strArquivo = Replace(strArquivo, ".xls", "", , , vbTextCompare)
        strArquivo = strArquivo & "_" & Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00")
        strArquivo = strArquivo & ".xls"
        .FileName = strArquivo
        
        .ShowSave
        
    End With
    
    'Deletando a planilha caso ela exista
    If Dir$(cdgGravarPlanilha.FileName) <> "" Then
        Kill strArquivo
    End If

    'Salvando a planilha
    wrkBook.SaveAs (cdgGravarPlanilha.FileName)
    
    'Restaurando o valor original da propriedade Interactive
    objExcel.Interactive = True
    
    wrkBook.Close
    objExcel.Quit
    
    Set wrkSheetDados = Nothing
    Set wrkBook = Nothing
    Set objExcel = Nothing
    
    'Inicializando a ProgressBar
    With prbPlanilha
        .Min = 0
        .Value = 0
        .Max = 100
    End With
    
    With prbProjetos
        .Min = 0
        .Value = 0
        .Max = 100
    End With
    
    Screen.MousePointer = vbNormal
    
    ExportarArqCausasAtrasadas = True
    
    Exit Function
        
ErrExportarArqCausasAtrasadas:
    
    Screen.MousePointer = vbNormal
    
    If Not (objExcel Is Nothing) Then
    
        'Restaurando o valor original da propriedade Interactive
        objExcel.Interactive = True
    
        wrkBook.Close
        objExcel.Quit
        
        Set wrkSheetDados = Nothing
        Set wrkBook = Nothing
        Set objExcel = Nothing
    
    End If
    
    If Not (rsProjetos Is Nothing) Then
        If rsProjetos.State = adStateOpen Then
            rsProjetos.Close
            Set rsProjetos = Nothing
        End If
    End If
    
    If Not (rsTarefas Is Nothing) Then
        If rsTarefas.State = adStateOpen Then
            rsTarefas.Close
            Set rsTarefas = Nothing
        End If
    End If
        
    ExportarArqCausasAtrasadas = False
    
End Function

Private Sub Form_Unload(Cancel As Integer)

    frmPrincipal.StatusBar.Panels(1).Text = ""
    
End Sub


