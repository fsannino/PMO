VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCargaPMOnline 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga das planilhas geradas pelo PMOnline"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "frmCargaPMOnline.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.DriveListBox drvDrivePlanilhasCarga 
      Height          =   315
      Left            =   135
      TabIndex        =   7
      Top             =   435
      Width           =   5025
   End
   Begin VB.DirListBox dirPastaPlanilhasCarga 
      Height          =   1665
      Left            =   135
      TabIndex        =   6
      Top             =   1245
      Width           =   5025
   End
   Begin VB.CommandButton cmdCancelarCarga 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   3735
      TabIndex        =   5
      Top             =   5610
      Width           =   1440
   End
   Begin VB.CommandButton cmdIniciarCarga 
      Caption         =   "Iniciar Carga"
      Height          =   345
      Left            =   2205
      TabIndex        =   4
      Top             =   5610
      Width           =   1440
   End
   Begin VB.Frame fraAndamentoImportacao 
      Caption         =   "Andamento geral da importação"
      Height          =   915
      Left            =   105
      TabIndex        =   2
      Top             =   4320
      Width           =   5070
      Begin MSComctlLib.ProgressBar ProgressBarGeral 
         Height          =   330
         Left            =   45
         TabIndex        =   3
         Top             =   375
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   1
      End
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
      Caption         =   "Selecione a pasta onde se encontram as planilhas:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   930
      Width           =   3600
   End
   Begin VB.Label lblDrive 
      AutoSize        =   -1  'True
      Caption         =   "Selecione o drive onde se encontram as planilhas:"
      Height          =   195
      Left            =   105
      TabIndex        =   8
      Top             =   135
      Width           =   3555
   End
End
Attribute VB_Name = "frmCargaPMOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub mInicializaObjetosCarga()

    fraAndamentoPlanilha.Caption = ""
    
    With ProgressBarPlanilhas
        .Min = 0
        .Max = 100
        .Value = 0
    End With
    
    With ProgressBarGeral
        .Min = 0
        .Max = 100
        .Value = 0
    End With
    
    With drvDrivePlanilhasCarga
        .Drive = "P:\"
    End With
    
    With dirPastaPlanilhasCarga
        .Path = "P:\Petrobras\Projeto\Pmo\Operacoes\Planilhas PMOnline"
    End With
    
End Sub


Private Function mExistemPlanilhasPMOnline(ByVal pPasta As String) As Boolean

    Dim fso As FileSystemObject
    
    mExistemPlanilhasPMOnline = False
        
    Set fso = New FileSystemObject
    
    If fso.FileExists(pPasta & "IS.xls") And _
       fso.FileExists(pPasta & "RK.xls") And _
       fso.FileExists(pPasta & "AC.xls") And _
       fso.FileExists(pPasta & "CR.xls") And _
       fso.FileExists(pPasta & "TK.xls") And _
       fso.FileExists(pPasta & "MT.xls") And _
       fso.FileExists(pPasta & "TC.xls") Then
       
        mExistemPlanilhasPMOnline = True
        
    End If
    
    Set fso = Nothing

End Function

Private Sub cmdCancelarCarga_Click()

    Unload Me
    
End Sub


Private Sub cmdIniciarCarga_Click()

    Dim objCargaPMOnline  As clsCargaPMOnline
    Dim strPastaPlanilhas As String
    Dim strMensagem       As String

    On Error GoTo ErrcmdIniciarCarga_Click

    Screen.MousePointer = vbHourglass

    Set objCargaPMOnline = New clsCargaPMOnline

    strPastaPlanilhas = gsTrataPath(dirPastaPlanilhasCarga.Path)

    If mExistemPlanilhasPMOnline(strPastaPlanilhas) Then
        strMensagem = objCargaPMOnline.ImportarPlanilhasPMOnline(strPastaPlanilhas, _
                                                                 ProgressBarGeral, _
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


