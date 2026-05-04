VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmImpArq_H_Teste_Carga 
   Caption         =   "Importação do Arquivo da Ferramenta H - Teste de Carga"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   Icon            =   "frmImpArq_H_Teste_Carga.frx":0000
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
Attribute VB_Name = "frmImpArq_H_Teste_Carga"
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
        MsgBox "Diretório ou arquivo não existe.", vbCritical + vbOKOnly, "Importação do Arquivo da Ferramenta H - Teste de Carga"
        Exit Sub
    End If

    If UCase(Right(strArq, 4)) <> ".XLS" Then
        MsgBox "Arquivo inválido.", vbCritical + vbOKOnly, "Importação do Arquivo da Ferramenta H - Teste de Carga"
        Exit Sub
    End If

    cmdOK.Enabled = False
    cmdProcura.Enabled = False
    cmdCancela.Enabled = False
    
    If ImportarArq(strArq) Then
        MsgBox "Exportação gerada com sucesso.", vbInformation + vbOKOnly, "Geração do Arquivo para Ferramenta H - Teste de Carga"
    Else
        MsgBox "Erro na geração do arquivo.", vbCritical + vbOKOnly, "Geração do Arquivo para Ferramenta H - Teste de Carga"
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
    
    CommonDialog1.DialogTitle = "Importação do Arquivo F.H - Teste de Carga"
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

Dim objExcel2         As Excel.Application
Dim wrkSheetDados2    As Worksheet
Dim wrkBook2          As Workbook


Dim rs                As New ADODB.Recordset
Dim strSql            As String
Dim intLinha          As Integer
Dim intLinhaAux       As Integer
Dim intLinhaAux2      As Integer
Dim intTotal          As Integer
Dim strArqSaida       As String
Dim strArqSaida2      As String

' Variaveis de atualizacao
Dim intCod_Proj       As String
Dim intUID            As String
Dim strDescricao      As String
Dim intCompletezaE     As Integer
Dim intCompletezaV     As Integer
Dim strCODDES         As String
Dim strCiclo         As String

On Error GoTo ErrImportarArq

        ImportarArq = False
        
        Screen.MousePointer = vbHourglass
        
        strArqSaida = Mid(strArquivo, 1, InStrRev(strArquivo, "\")) & "Inc_FH_TC" & Format(Year(Now()), "0000") & Format(Month(Now()), "00") & Format(Day(Now()), "00")
        
        strArqSaida2 = Mid(strArquivo, 1, InStrRev(strArquivo, "\")) & "Nov_FH_TC" & Format(Year(Now()), "0000") & Format(Month(Now()), "00") & Format(Day(Now()), "00")
        
        If Not (objExcel Is Nothing) Then

            wrkBook.Close
            objExcel.Quit
            
            wrkBook1.Close
            objExcel1.Quit

            wrkBook2.Close
            objExcel2.Quit

            Set wrkSheetDados = Nothing
            Set wrkBook = Nothing
            Set objExcel = Nothing
            
            Set wrkSheetDados1 = Nothing
            Set wrkBook1 = Nothing
            Set objExcel1 = Nothing
        
            Set wrkSheetDados2 = Nothing
            Set wrkBook2 = Nothing
            Set objExcel2 = Nothing
        
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

        'Criando uma nova instância para o objeto
        Set objExcel2 = CreateObject("Excel.Application")
        'Adicionando um WorkBook
        Set wrkBook2 = objExcel2.Workbooks.Add
        'Setando um WorkSheet
        Set wrkSheetDados2 = wrkBook2.Sheets(1)

        'Verificando se conseguiu inicializar o objeto
        If objExcel Is Nothing Then
            Screen.MousePointer = vbNormal
            MsgBox "O aplicativo Excel não pode ser inicializado." & vbCrLf, _
                   vbCritical + vbOKOnly, _
                   "ERRO"
            End
        End If
        
        'Deixando o Excel invisível
        objExcel.Visible = False
        
        'Vai inibir algumas mensagens enviadas pelo Excel
        objExcel.DisplayAlerts = False
        
        'Deixando o Excel invisível
        objExcel1.Visible = False
        
        'Vai inibir algumas mensagens enviadas pelo Excel
        objExcel1.DisplayAlerts = False
        
        'Deixando o Excel invisível
        objExcel2.Visible = False
        
        'Vai inibir algumas mensagens enviadas pelo Excel
        objExcel2.DisplayAlerts = False
        
        wrkSheetDados1.Cells(1, 1).Value = "CODDES"
        wrkSheetDados1.Cells(1, 2).Value = "Descrição"
        wrkSheetDados1.Cells(1, 3).Value = "Ciclo"
        wrkSheetDados1.Cells(1, 4).Value = "Completeza E"
        wrkSheetDados1.Cells(1, 5).Value = "Completeza V"
        
        wrkSheetDados2.Cells(1, 1).Value = "CODDES"
        wrkSheetDados2.Cells(1, 2).Value = "Descrição"
        wrkSheetDados2.Cells(1, 3).Value = "Ciclo"
        wrkSheetDados2.Cells(1, 4).Value = "Completeza E"
        wrkSheetDados2.Cells(1, 5).Value = "Completeza V"
        
        intLinhaAux = 2
        
        intLinhaAux2 = 2
        
        intTotal = objExcel.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
        
        prbExp.Max = intTotal
        prbExp.Min = 1
        
        For intLinha = 2 To intTotal
            
            strCODDES = wrkSheetDados.Cells(intLinha, 1).Value
            strDescricao = wrkSheetDados.Cells(intLinha, 2).Value
            strCiclo = wrkSheetDados.Cells(intLinha, 3).Value
            intCompletezaE = wrkSheetDados.Cells(intLinha, 4).Value
            intCompletezaV = wrkSheetDados.Cells(intLinha, 5).Value
            
            
            If Trim(strCODDES) <> "" Then
            
                strSql = "SP_LISTAR_DESENVOLVIMENTO_TESTE_CARGA '" & strCODDES & "', 'E" & Trim(strCiclo) & "' "
                
                rs.Open strSql, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
            
                If Not rs.EOF Then
                    
                    intCod_Proj = rs("PROJ_ID")
                    intUID = rs("TASK_UID")
                    
                    If intCompletezaE < rs("TASK_PCT_COMP") Then
                        
                        wrkSheetDados1.Cells(intLinhaAux, 1).Value = strCODDES
                        wrkSheetDados1.Cells(intLinhaAux, 2).Value = strDescricao
                        wrkSheetDados1.Cells(intLinhaAux, 3).Value = strCiclo
                        wrkSheetDados1.Cells(intLinhaAux, 4).Value = intCompletezaE
                        wrkSheetDados1.Cells(intLinhaAux, 5).Value = intCompletezaV
                        
                        intLinhaAux = intLinhaAux + 1
                    
                    Else
                    
                       strSql = "EXEC SP_ATUALIZAR_DESENVOLVIMENTO " & intCod_Proj & ", " & intUID & ", Null, " & intCompletezaE & ", Null, '" & gsReplicaPlics(strDescricao) & "', Null, Null "
                    
                    End If
                    
                    dbConexaoSMR.Execute strSql
                
                Else
                
                    wrkSheetDados2.Cells(intLinhaAux2, 1).Value = strCODDES
                    wrkSheetDados2.Cells(intLinhaAux2, 2).Value = strDescricao
                    wrkSheetDados2.Cells(intLinhaAux2, 3).Value = strCiclo
                    wrkSheetDados2.Cells(intLinhaAux2, 4).Value = intCompletezaE
                    wrkSheetDados2.Cells(intLinhaAux2, 5).Value = intCompletezaV

                    intLinhaAux2 = intLinhaAux2 + 1
                
                End If
                
                rs.Close
                            
            
                strSql = "SP_LISTAR_DESENVOLVIMENTO_TESTE_CARGA '" & strCODDES & "', 'V" & Trim(strCiclo) & "' "
                
                rs.Open strSql, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
            
                If Not rs.EOF Then
                    
                    intCod_Proj = rs("PROJ_ID")
                    intUID = rs("TASK_UID")
                    
                    If intCompletezaV < rs("TASK_PCT_COMP") Then
                        
                        wrkSheetDados1.Cells(intLinhaAux, 1).Value = strCODDES
                        wrkSheetDados1.Cells(intLinhaAux, 2).Value = strDescricao
                        wrkSheetDados1.Cells(intLinhaAux, 3).Value = strCiclo
                        wrkSheetDados1.Cells(intLinhaAux, 4).Value = intCompletezaE
                        wrkSheetDados1.Cells(intLinhaAux, 5).Value = intCompletezaV
                        
                        intLinhaAux = intLinhaAux + 1
                    
                    Else
                    
                       strSql = "EXEC SP_ATUALIZAR_DESENVOLVIMENTO " & intCod_Proj & ", " & intUID & ", Null, " & intCompletezaV & ", Null, '" & gsReplicaPlics(strDescricao) & "', Null, Null "
                    
                    End If
                    
                    dbConexaoSMR.Execute strSql
                
                Else
                
                    wrkSheetDados2.Cells(intLinhaAux2, 1).Value = strCODDES
                    wrkSheetDados2.Cells(intLinhaAux2, 2).Value = strDescricao
                    wrkSheetDados2.Cells(intLinhaAux2, 3).Value = strCiclo
                    wrkSheetDados2.Cells(intLinhaAux2, 4).Value = intCompletezaE
                    wrkSheetDados2.Cells(intLinhaAux2, 5).Value = intCompletezaV

                    intLinhaAux2 = intLinhaAux2 + 1
                
                End If
                
                rs.Close
            
            End If
            
            prbExp.Value = intLinha

        Next
        
        'Deletando a planilha caso ela exista
        If Dir$(strArqSaida) <> "" Then
            Kill strArqSaida
        End If

        'Salvando a planilha criada
        wrkBook1.SaveAs strArqSaida
        
        'Deletando a planilha caso ela exista
        If Dir$(strArqSaida2) <> "" Then
            Kill strArqSaida2
        End If

        'Salvando a planilha criada
        wrkBook2.SaveAs strArqSaida2
        
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
    
        'Deixando o Excel visível
        objExcel2.Visible = True

        wrkBook2.Close
        objExcel2.Quit
        
        Set wrkSheetDados2 = Nothing
        Set wrkBook2 = Nothing
        Set objExcel2 = Nothing
    
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
    
    If Not (objExcel2 Is Nothing) Then
        
        'Deixando o Excel visível
        objExcel2.Visible = True

        wrkBook2.Close
        objExcel2.Quit
        
        Set wrkSheetDados2 = Nothing
        Set wrkBook2 = Nothing
        Set objExcel2 = Nothing
    
    End If
    
    ImportarArq = False
    
End Function

Private Sub Form_Unload(Cancel As Integer)

    frmPrincipal.StatusBar.Panels(1).Text = ""
    
End Sub


