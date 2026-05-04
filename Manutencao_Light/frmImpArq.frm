VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImpArq 
   Caption         =   "Importação de Arquivos"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Opções"
      Height          =   645
      Left            =   120
      TabIndex        =   7
      Top             =   885
      Width           =   5775
      Begin VB.CheckBox chkDatas 
         Caption         =   "Datas"
         Height          =   300
         Left            =   4440
         TabIndex        =   10
         Top             =   225
         Width           =   1035
      End
      Begin VB.CheckBox chkComp 
         Caption         =   "% Complete"
         Height          =   255
         Left            =   2220
         TabIndex        =   9
         Top             =   255
         Width           =   1455
      End
      Begin VB.CheckBox chkNome 
         Caption         =   "Nome"
         Height          =   225
         Left            =   375
         TabIndex        =   8
         Top             =   270
         Width           =   1065
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   2325
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   120
      TabIndex        =   5
      Top             =   1575
      Width           =   5775
      Begin MSComctlLib.ProgressBar prbExp 
         Height          =   360
         Left            =   105
         TabIndex        =   6
         Top             =   195
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   635
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
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4545
      TabIndex        =   1
      Top             =   2325
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3225
      TabIndex        =   0
      Top             =   2340
      Width           =   1215
   End
End
Attribute VB_Name = "frmImpArq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
Dim strArq As String
Dim strDir As String

    strArq = txtDiretorio.Text

    If Dir(strArq) = "" Then
        MsgBox "Diretório ou arquivo não existe.", vbCritical + vbOKOnly, "Importação de Arquivos"
        Exit Sub
    End If

    If UCase(Right(strArq, 4)) <> ".XLS" Then
        MsgBox "Arquivo inválido.", vbCritical + vbOKOnly, "Importação de Arquivos"
        Exit Sub
    End If

    If chkNome.Value = 0 And chkComp.Value = 0 And chkDatas.Value = 0 Then
        MsgBox "Escolha uma opção.", vbCritical + vbOKOnly, "Importação de Arquivos"
        Exit Sub
    End If
    
    cmdOK.Enabled = False
    cmdProcura.Enabled = False
    cmdCancela.Enabled = False
    
    If ImportarArq(strArq) Then
        MsgBox "Importação de arquivos gerada com sucesso.", vbInformation + vbOKOnly, "Importação de Arquivos"
    Else
        MsgBox "Erro na Importação de arquivos.", vbCritical + vbOKOnly, "Importação de Arquivos"
    End If
    
    cmdOK.Enabled = True
    cmdProcura.Enabled = True
    cmdCancela.Enabled = True
    prbExp.Value = 1
    
End Sub

Private Sub cmdProcura_Click()
Dim strDirInicial As String

On Error GoTo ErrcmdProcura

    strDirInicial = txtDiretorio.Text
    
    CommonDialog1.DialogTitle = "Importação de Arquivos"
    CommonDialog1.FileName = ""
    CommonDialog1.Filter = "Arquivos do MS-Excel (*.xls)|*.xls"
    CommonDialog1.CancelError = True
    CommonDialog1.ShowOpen
    txtDiretorio.Text = CommonDialog1.FileName

Exit Sub

ErrcmdProcura:

txtDiretorio.Text = strDirInicial

End Sub

Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtDiretorio = "C:\Planilha.xls"
End Sub

Private Function ImportarArq(strArquivo As String) As Boolean

Dim objExcel          As Excel.Application
Dim wrkSheetDados     As Worksheet
Dim wrkBook           As Workbook

Dim objExcel1         As Excel.Application
Dim wrkSheetDados1    As Worksheet
Dim wrkBook1          As Workbook

Dim rs                As New ADODB.Recordset
Dim cmdAtualizar      As New ADODB.Command

Dim strSQL            As String
Dim intLinha          As Integer
Dim intLinhaAux       As Integer
Dim intTotal          As Integer
Dim strArqSaida       As String

' Variaveis de atualizacao
Dim intCod_Proj       As String
Dim intUID            As String
Dim strDescricao      As String
Dim intCompleteza     As Integer
Dim strDataIni        As String
Dim strDataFim        As String

On Error GoTo ErrImportarArq

        ImportarArq = False
        
        Screen.MousePointer = vbHourglass
        
        strArqSaida = Mid(strArquivo, 1, InStrRev(strArquivo, "\")) & "INC_IMP_ARQ" & Format(Year(Now()), "0000") & Format(Month(Now()), "00") & Format(Day(Now()), "00")
        
        If Not (objExcel Is Nothing) Then

            wrkBook.Close
            objExcel.Quit
            
            wrkBook1.Close
            objExcel1.Quit

            Set wrkSheetDados = Nothing
            Set wrkBook = Nothing
            Set objExcel = Nothing
            
            Set wrkSheetDados1 = Nothing
            Set wrkBook1 = Nothing
            Set objExcel1 = Nothing
        
        End If

        'Criando uma nova instância para o objeto
        Set objExcel = CreateObject("Excel.Application")
        'Adicionando um WorkBook
        Set wrkBook = objExcel.Workbooks.Open(strArquivo)
        'Setando um WorkSheet
        Set wrkSheetDados = wrkBook.Sheets(1)
        
        'Criando uma nova instância para o objeto
        Set objExcel1 = CreateObject("Excel.Application")
        'Adicionando um WorkBook
        Set wrkBook1 = objExcel1.Workbooks.Add
        'Setando um WorkSheet
        Set wrkSheetDados1 = wrkBook1.Sheets(1)

        'Deixando o Excel invisível
        objExcel.Visible = False
        
        'Vai inibir algumas mensagens enviadas pelo Excel
        objExcel.DisplayAlerts = False
        
        'Deixando o Excel invisível
        objExcel1.Visible = False
        
        'Vai inibir algumas mensagens enviadas pelo Excel
        objExcel1.DisplayAlerts = False
        
        
        wrkSheetDados1.Cells(1, 1).Value = "Proj_id"
        wrkSheetDados1.Cells(1, 2).Value = "UID"
        wrkSheetDados1.Cells(1, 3).Value = "Descrição"
        wrkSheetDados1.Cells(1, 4).Value = "Completeza Ant."
        wrkSheetDados1.Cells(1, 5).Value = "Completeza Nov."
        wrkSheetDados1.Cells(1, 6).Value = "Data Inicio"
        wrkSheetDados1.Cells(1, 7).Value = "Data Fim"
        
        intLinhaAux = 2
        
        intTotal = objExcel.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
        
        prbExp.Max = intTotal
        prbExp.Min = 1
        
        For intLinha = 2 To intTotal
            
            intCod_Proj = wrkSheetDados.Cells(intLinha, 1).Value
            intUID = wrkSheetDados.Cells(intLinha, 2).Value
            strDescricao = wrkSheetDados.Cells(intLinha, 3).Value
            intCompleteza = wrkSheetDados.Cells(intLinha, 4).Value
            strDataIni = wrkSheetDados.Cells(intLinha, 5).Value
            strDataFim = wrkSheetDados.Cells(intLinha, 6).Value
            
            strSQL = "SP_LISTAR_TAREFAS " & intCod_Proj & ", " & intUID & ""
            
            rs.Open strSQL, dbConexaoLight, adOpenStatic, adLockReadOnly, adCmdText
            
            If Not rs.EOF Then
                
                If chkComp = 1 Then
                                
                    If intCompleteza < rs("TASK_PCT_COMP") Then
                     
                        wrkSheetDados1.Cells(intLinhaAux, 1).Value = intCod_Proj
                        wrkSheetDados1.Cells(intLinhaAux, 2).Value = intUID
                        wrkSheetDados1.Cells(intLinhaAux, 3).Value = strDescricao
                        wrkSheetDados1.Cells(intLinhaAux, 4).Value = rs("TASK_PCT_COMP")
                        wrkSheetDados1.Cells(intLinhaAux, 5).Value = intCompleteza
                        wrkSheetDados1.Cells(intLinhaAux, 6).Value = strDataIni
                        wrkSheetDados1.Cells(intLinhaAux, 7).Value = strDataFim
                        
                        intLinhaAux = intLinhaAux + 1
                    
                    Else
    
                        With cmdAtualizar
                        
                            .ActiveConnection = dbConexaoLight
                            .CommandType = adCmdStoredProc
                            .CommandText = "SP_ATUALIZAR_TAREFAS_ARQ"
                            
                            .Parameters.Refresh
                            .Parameters(1).Value = intCod_Proj
                            .Parameters(2).Value = intUID
                        
                            If chkNome = 1 Then
                                .Parameters(3).Value = Trim(strDescricao)
                            End If
                            
                            If chkComp = 1 Then
                                .Parameters(4).Value = intCompleteza
                            End If
                            
                            If chkDatas = 1 Then
                                .Parameters(5).Value = strDataIni
                                .Parameters(6).Value = strDataFim
                            End If
                            
                            .Execute
                            
                        End With
                    
                    End If
                Else
                    With cmdAtualizar
                    
                        .ActiveConnection = dbConexaoLight
                        .CommandType = adCmdStoredProc
                        .CommandText = "SP_ATUALIZAR_TAREFAS_ARQ"
                        
                        .Parameters.Refresh
                        .Parameters(1).Value = intCod_Proj
                        .Parameters(2).Value = intUID
                    
                        If chkNome = 1 Then
                            .Parameters(3).Value = Trim(strDescricao)
                        End If
                        
                        If chkComp = 1 Then
                            .Parameters(4).Value = intCompleteza
                        End If
                        
                        If chkDatas = 1 Then
                            .Parameters(5).Value = strDataIni
                            .Parameters(6).Value = strDataFim
                        End If
                        
                        .Execute
                        
                    End With
                End If
            End If
                
            rs.Close
                            
            prbExp.Value = intLinha

        Next
        
        'Deletando a planilha caso ela exista
        If Dir$(strArqSaida) <> "" Then
            Kill strArqSaida
        End If

        'Salvando a planilha criada
        wrkBook1.SaveAs strArqSaida
        
        'Deixando o Excel visível
        objExcel.Visible = True

        wrkBook.Close
        objExcel.Quit
        
        Set wrkSheetDados = Nothing
        Set wrkBook = Nothing
        Set objExcel = Nothing
    
        'Deixando o Excel visível
        objExcel1.Visible = True

        wrkBook1.Close
        objExcel1.Quit
        
        Set wrkSheetDados1 = Nothing
        Set wrkBook1 = Nothing
        Set objExcel1 = Nothing
    
        Screen.MousePointer = vbNormal

        ImportarArq = True
        
        Exit Function
        
ErrImportarArq:
    
    Screen.MousePointer = vbNormal
    
    If Not (objExcel Is Nothing) Then
    
        'Deixando o Excel visível
        objExcel.Visible = True

        wrkBook.Close
        objExcel.Quit
        
        Set wrkSheetDados = Nothing
        Set wrkBook = Nothing
        Set objExcel = Nothing
    
    End If
        
    If Not (objExcel1 Is Nothing) Then
        
        'Deixando o Excel visível
        objExcel1.Visible = True

        wrkBook1.Close
        objExcel1.Quit
        
        Set wrkSheetDados1 = Nothing
        Set wrkBook1 = Nothing
        Set objExcel1 = Nothing
    
    End If
    
    ImportarArq = False
    
End Function



