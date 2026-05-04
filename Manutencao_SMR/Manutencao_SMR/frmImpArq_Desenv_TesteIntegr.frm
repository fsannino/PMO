VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmImpArq_Desenv_TesteIntegr 
   Caption         =   "Importação do Arquivo Desenvolvimento Teste Integrado"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   Icon            =   "frmImpArq_Desenv_TesteIntegr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   120
      TabIndex        =   5
      Top             =   870
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
      Top             =   1620
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3225
      TabIndex        =   0
      Top             =   1635
      Width           =   1215
   End
End
Attribute VB_Name = "frmImpArq_Desenv_TesteIntegr"
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
        MsgBox "Diretório ou arquivo não existe.", vbCritical + vbOKOnly, "Importação do Arquivo Desenvolvimento Teste Integrado"
        Exit Sub
    End If

    If UCase(Right(strArq, 4)) <> ".XLS" Then
        MsgBox "Arquivo inválido.", vbCritical + vbOKOnly, "Importação do Arquivo Desenvolvimento Teste Integrado"
        Exit Sub
    End If

    cmdOK.Enabled = False
    cmdProcura.Enabled = False
    cmdCancela.Enabled = False
    
    If ImportarArq(strArq) Then
        MsgBox "Importação do arquivo gerada com sucesso.", vbInformation + vbOKOnly, "Importação do Arquivo Desenvolvimento Teste Integrado"
    Else
        MsgBox "Erro na importação do arquivo.", vbCritical + vbOKOnly, "Importação do Arquivo Desenvolvimento Teste Integrado"
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
    
    CommonDialog1.DialogTitle = "Importação do Arquivo Desenvolvimento Teste Integrado"
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

Dim cmdAtualizar      As New ADODB.Command
Dim rs                As New ADODB.Recordset
Dim strSQL            As String
Dim intLinha          As Integer
Dim intLinha1         As Integer
Dim intLinha2         As Integer

Dim intTotal          As Integer
Dim strArqSaida1      As String
Dim strArqSaida2      As String

' Variaveis de atualizacao
Dim strCod_Desenv     As String
Dim strDescricao      As String
Dim dtDataInicio      As Date
Dim dtDataFim         As Date

On Error GoTo ErrImportarArq

        ImportarArq = False
        
        Call Excluir_Flag11
        
        Screen.MousePointer = vbHourglass
        
        strArqSaida1 = Mid(strArquivo, 1, InStrRev(strArquivo, "\")) & "DESENV_NAO_ENCOTRADOS" & Format(Year(Now()), "0000") & Format(Month(Now()), "00") & Format(Day(Now()), "00")
        
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
        
        wrkSheetDados1.Cells(1, 1).Value = "Cod_Desenv."
        wrkSheetDados1.Cells(1, 2).Value = "Tarefa"
        wrkSheetDados1.Cells(1, 3).Value = "Data Inicio"
        wrkSheetDados1.Cells(1, 4).Value = "Data Fim"
        
        intLinha1 = 2
        
        intTotal = objExcel.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
        
        prbExp.Max = intTotal
        prbExp.Min = 1
        
        For intLinha = 2 To intTotal
            
            strCod_Desenv = wrkSheetDados.Cells(intLinha, 1).Value
            strDescricao = wrkSheetDados.Cells(intLinha, 2).Value
            dtDataInicio = wrkSheetDados.Cells(intLinha, 3).Value
            dtDataFim = wrkSheetDados.Cells(intLinha, 4).Value
            
            strSQL = "SP_LISTAR_TAREFAS_DESENV_TIN '" & strCod_Desenv & "' "
            
            rs.Open strSQL, dbConexaoTIN, adOpenStatic, adLockReadOnly, adCmdText
            
            If rs.EOF Then
                    
                wrkSheetDados1.Cells(intLinha1, 1).Value = strCod_Desenv
                wrkSheetDados1.Cells(intLinha1, 2).Value = strDescricao
                wrkSheetDados1.Cells(intLinha1, 3).Value = dtDataInicio
                wrkSheetDados1.Cells(intLinha1, 4).Value = dtDataFim

                intLinha1 = intLinha1 + 1

            Else
            
                With cmdAtualizar
                
                    .ActiveConnection = dbConexaoTIN
                    .CommandType = 4
                    .CommandTimeout = 1200
                    .CommandText = "SP_ATUALIZAR_DESENV_TIN"
                    
                    .Parameters.Refresh
                    .Parameters(1).Value = rs("PROJ_ID")
                    .Parameters(2).Value = rs("TASK_UID")
                    .Parameters(3).Value = strDescricao
                    .Parameters(4).Value = dtDataInicio
                    .Parameters(5).Value = dtDataFim
                
                End With
            
                cmdAtualizar.Execute
            
            End If
                
            rs.Close
            prbExp.Value = intLinha

        Next
        
        'Deletando a planilha caso ela exista
        If Dir$(strArqSaida1) <> "" Then
            Kill strArqSaida1
        End If

        'Salvando a planilha criada
        wrkBook1.SaveAs strArqSaida1
        
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
    
        Call Gerar_Planulha_Nao_Atualizadas(strArquivo)
                
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

Private Sub Form_Unload(Cancel As Integer)

    frmPrincipal.StatusBar.Panels(1).Text = ""
    
End Sub


Public Sub Gerar_Planulha_Nao_Atualizadas(strArquivo As String)

Dim objExcel         As Excel.Application
Dim wrkSheetDados    As Worksheet
Dim wrkBook          As Workbook

Dim rs               As New ADODB.Recordset
Dim strSQL           As String
Dim intLinha         As Integer
Dim strArqSaida      As String

        strArqSaida = Mid(strArquivo, 1, InStrRev(strArquivo, "\")) & "DESENV_NAO_ATUALIZADOS" & Format(Year(Now()), "0000") & Format(Month(Now()), "00") & Format(Day(Now()), "00")

        If Not (objExcel Is Nothing) Then

            wrkBook.Close
            objExcel.Quit

            Set wrkSheetDados = Nothing
            Set wrkBook = Nothing
            Set objExcel = Nothing
        
        End If

        'Criando uma nova instância para o objeto
        Set objExcel = CreateObject("Excel.Application")
        'Adicionando um WorkBook
        Set wrkBook = objExcel.Workbooks.Add
        'Setando um WorkSheet
        Set wrkSheetDados = wrkBook.Sheets(1)

        'Deixando o Excel invisível
        objExcel.Visible = False
        
        'Vai inibir algumas mensagens enviadas pelo Excel
        objExcel.DisplayAlerts = False
        
        wrkSheetDados.Cells(1, 1).Value = "Cod_Proj"
        wrkSheetDados.Cells(1, 2).Value = "UID"
        wrkSheetDados.Cells(1, 3).Value = "Tarefa"
        wrkSheetDados.Cells(1, 4).Value = "Data Inicio"
        wrkSheetDados.Cells(1, 5).Value = "Data Fim"
        
        intLinha = 2

        strSQL = "SP_LISTAR_TAREFAS_DESENV_NAO_ATUALIZADAS"
        
        rs.Open strSQL, dbConexaoTIN, adOpenStatic, adLockReadOnly, adCmdText

        Do While Not rs.EOF

            If rs("TASK_FINISH_DATE") > Date Then
            
                wrkSheetDados.Cells(intLinha, 1).Value = rs("PROJ_ID")
                wrkSheetDados.Cells(intLinha, 2).Value = rs("TASK_UID")
                wrkSheetDados.Cells(intLinha, 3).Value = rs("TASK_NAME")
                wrkSheetDados.Cells(intLinha, 4).Value = rs("TASK_START_DATE")
                wrkSheetDados.Cells(intLinha, 5).Value = rs("TASK_FINISH_DATE")
    
                intLinha = intLinha + 1
            End If
            
            rs.MoveNext
            
        Loop
        
        rs.Close
        
        'Deletando a planilha caso ela exista
        If Dir$(strArqSaida) <> "" Then
            Kill strArqSaida
        End If

        'Salvando a planilha criada
        wrkBook.SaveAs strArqSaida
        
        'Deixando o Excel visível
        objExcel.Visible = True

        wrkBook.Close
        objExcel.Quit
        
        Set wrkSheetDados = Nothing
        Set wrkBook = Nothing
        Set objExcel = Nothing
        
        
End Sub

Public Sub Excluir_Flag11()
Dim strSQL           As String

    strSQL = "SP_EXCLUIR_FLAG11"

    dbConexaoTIN.Execute strSQL

End Sub
