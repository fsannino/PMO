VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAtualizarOperacoesCutOver 
   Caption         =   "Atualização das Operações Cut Over"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   Icon            =   "frmAtualizarOperacoesCutOver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   120
      TabIndex        =   2
      Top             =   105
      Width           =   5775
      Begin MSComctlLib.ProgressBar prbExp 
         Height          =   360
         Left            =   105
         TabIndex        =   3
         Top             =   195
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.CommandButton cmdCancela 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4545
      TabIndex        =   1
      Top             =   855
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3225
      TabIndex        =   0
      Top             =   870
      Width           =   1215
   End
End
Attribute VB_Name = "frmAtualizarOperacoesCutOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    cmdOK.Enabled = False
    cmdCancela.Enabled = False
   
    If AtualizarOperacoesCutOver() Then
        MsgBox "Atualização de exclusões de tarefas realizada com sucesso.", vbInformation + vbOKOnly, "Atualização de Exclusões de Tarefas"
    Else
        MsgBox "Erro na atualização de exclusões de tarefas.", vbCritical + vbOKOnly, "Atualização de Exclusões de Tarefas"
    End If
    
    cmdOK.Enabled = True
    cmdCancela.Enabled = True
    prbExp.Value = 0
    
End Sub


Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Function AtualizarOperacoesCutOver() As Boolean


Dim strSQL          As String
Dim rsProjeto       As New ADODB.Recordset
Dim MSProject       As Object
Dim intArq          As Long
Dim objTarefa       As Object
Dim cont            As Long
Dim cont2           As Long
Dim strNome         As String

On Error GoTo ErrAtualizarOperacoesCutOver
    
    
    strSQL = "select max(task_id) as task_id from MSP_TASKS Where PROJ_ID = 1"
    
    rsProjeto.Open strSQL, dbConexaoCUT, adOpenKeyset, adLockOptimistic
    
    cont = rsProjeto.Fields("task_id")
    
    rsProjeto.Close
    
    strSQL = "SP_LISTAR_OPERACOES_TAREFAS"
    
    rsProjeto.Open strSQL, dbConexaoCUT, adOpenStatic, adLockReadOnly, adCmdText
    
    If rsProjeto.EOF Then
        Call MsgBox("Não existem registros para serem excluidos")
        Exit Function
    Else
        
        prbExp.Max = rsProjeto.RecordCount
        prbExp.Min = 0
        prbExp.Value = 0

        Set MSProject = CreateObject("msproject.application")
        MSProject.Visible = True
        MSProject.DisplayAlerts = False
        
        MSProject.FileOpen Name:="<MASTER_CUT>\Cut Over - Petrobras", ReadOnly:=False, UserID:="MASTER_SMR", DatabasePassWord:="oivalf", FormatID:="MSProject.ODBC"

        MSProject.OutlineShowAllTasks
        MSProject.OutlineShowAllTasks
        MSProject.ViewApply Name:="Gantt Chart"
        MSProject.OptionsSchedule AutoSplit:=False, HonorConstraints:=False
        MSProject.SelectAll
    
        For Each objTarefa In MSProject.ActiveProject.Tasks
            
            If cont = rsProjeto.Fields("ID") Then
            
                If rsProjeto.Fields("Tipo") = "I" Then
                
                    cont2 = cont + 1
                    strNome = rsProjeto.Fields("Nome")
                    If MSProject.ActiveProject.Tasks(cont).UniqueID = rsProjeto.Fields("UID") Then
                        MSProject.EditGoto cont2
                        MSProject.ActiveProject.Tasks.Add strNome, cont2
                        MSProject.ActiveProject.Tasks(cont2).Start = Format(rsProjeto.Fields("DataIni"), "dd/mm/yyyy hh:mm:ss")
                        MSProject.ActiveProject.Tasks(cont2).Finish = Format(rsProjeto.Fields("DataFim"), "dd/mm/yyyy hh:mm:ss")
                        
                        MSProject.ActiveProject.Tasks(cont2).Type = 1
                        MSProject.ActiveProject.Saved

                    End If
                    
                    prbExp.Value = prbExp.Value + 1
                
                ElseIf rsProjeto.Fields("Tipo") = "E" Then
                
                    If MSProject.ActiveProject.Tasks(cont).UniqueID = rsProjeto.Fields("UID") Then
                    
                        MSProject.ActiveProject.Tasks(cont).PercentComplete = 0
                        MSProject.EditGoto cont
                    
                        MSProject.EditDelete
                        MSProject.ActiveProject.Saved
                    End If
                
                    prbExp.Value = prbExp.Value + 1
                
                ElseIf rsProjeto.Fields("Tipo") = "A" Then
                
                    If MSProject.ActiveProject.Tasks(cont).UniqueID = rsProjeto.Fields("UID") Then
                    
                        MSProject.ActiveProject.Tasks(cont).Start = Format(rsProjeto.Fields("DataIni"), "dd/mm/yyyy hh:mm:ss")
                        MSProject.ActiveProject.Tasks(cont).Finish = Format(rsProjeto.Fields("DataFim"), "dd/mm/yyyy hh:mm:ss")
                        MSProject.EditGoto cont
                        MSProject.ActiveProject.Saved
                    End If
                    
                    prbExp.Value = prbExp.Value + 1
                   
                End If
                
                rsProjeto.MoveNext
                
                If rsProjeto.EOF Then
                    Exit For
                End If
                
            ElseIf cont < rsProjeto.Fields("ID") Then
                rsProjeto.MoveNext
            ElseIf cont > rsProjeto.Fields("ID") Then
                cont = cont - 1
            End If
            
        Next
        
        MSProject.ActiveProject.Saved

        MSProject.Application.FileCloseAll 2
        
        Call Deletaregistro
    
    End If
        
    AtualizarOperacoesCutOver = True
    
    Exit Function
        
ErrAtualizarOperacoesCutOver:
    
    Screen.MousePointer = vbNormal
    AtualizarOperacoesCutOver = False
    
End Function


Sub Deletaregistro()
    
    Dim sSQL As String
    
    sSQL = "SP_EXCLUIR_OPERACOES"
    
    dbConexaoCUT.Execute sSQL
    
End Sub

