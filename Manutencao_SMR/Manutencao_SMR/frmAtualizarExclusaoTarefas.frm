VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAtualizarExclusaoTarefas 
   Caption         =   "Atualização de Exclusões de Tarefas"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   Icon            =   "frmAtualizarExclusaoTarefas.frx":0000
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
Attribute VB_Name = "frmAtualizarExclusaoTarefas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    cmdOK.Enabled = False
    cmdCancela.Enabled = False
   
    If AtualizarExclusoesTarefas() Then
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

Private Function AtualizarExclusoesTarefas() As Boolean


Dim strSQL          As String
Dim rsProjeto       As New ADODB.Recordset
Dim MSProject       As Object
Dim intArq          As Long
Dim objTarefa       As Object
Dim cont            As Long

On Error GoTo ErrAtualizarExclusoesTarefas
    
    
    strSQL = "select max(task_id) as task_id from MSP_TASKS Where PROJ_ID = 9"
    
    rsProjeto.Open strSQL, dbConexaoSMR, adOpenKeyset, adLockOptimistic
    
    cont = rsProjeto.Fields("task_id")
    
    rsProjeto.Close
    
    strSQL = "select * "
    strSQL = strSQL & "From TarefasExcluidas_GVI "
    strSQL = strSQL & "order by ID desc "
    
    rsProjeto.Open strSQL, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
    
    If rsProjeto.EOF Then
        Call MsgBox("Não existem registros para serem excluidos")
        Exit Function
    Else
        
        prbExp.Max = rsProjeto.RecordCount
        prbExp.Min = 0
        prbExp.Value = 0
        
        Set MSProject = CreateObject("msproject.application")
        MSProject.Visible = False
        MSProject.DisplayAlerts = False
        
        MSProject.FileOpen Name:="<MASTER_SMR>\8. Cronograma de Impactos - Gov. Integrada", ReadOnly:=False, UserID:="MASTER_SMR", DatabasePassWord:="oivalf", FormatID:="MSProject.ODBC"
        
        MSProject.OutlineShowAllTasks
        MSProject.OutlineShowAllTasks
        MSProject.ViewApply Name:="Gantt Chart"
        MSProject.OptionsSchedule AutoSplit:=False, HonorConstraints:=False
        MSProject.SelectAll
    
        For Each objTarefa In MSProject.ActiveProject.Tasks
            
            If cont = rsProjeto.Fields("ID") Then
            
                If MSProject.ActiveProject.Tasks(cont).UniqueID = rsProjeto.Fields("UID") Then
                
                    MSProject.ActiveProject.Tasks(cont).PercentComplete = 0
                    MSProject.EditGoto cont
                
                    MSProject.EditDelete
                    MSProject.ActiveProject.Saved
                
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
        
    AtualizarExclusoesTarefas = True
    
    Exit Function
        
ErrAtualizarExclusoesTarefas:
    
    Screen.MousePointer = vbNormal
    AtualizarExclusoesTarefas = False
    
End Function


Sub Deletaregistro()
    
    Dim sSQL As String
    
    sSQL = "DELETE FROM TarefasExcluidas_GVI"
    
    dbConexaoSMR.Execute sSQL
    
End Sub

