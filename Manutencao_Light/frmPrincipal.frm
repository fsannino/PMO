VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "Módulo de Manutenção do Banco de Dados Light"
   ClientHeight    =   6375
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9945
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   6150
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   397
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "17/11/2004"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "12:48 PM"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "Cadastros"
      Begin VB.Menu mnuCadPMOAuxUsuarios 
         Caption         =   "Usuários"
      End
      Begin VB.Menu mnuCadConfigFechamento 
         Caption         =   "Configuração do Fechamento"
      End
   End
   Begin VB.Menu mnuTransmissoes 
      Caption         =   "Transmissões"
      Begin VB.Menu mnuCargaPMOnline 
         Caption         =   "Carga PMOnline"
      End
      Begin VB.Menu mnuImportacaoPlanilha 
         Caption         =   "Importação de Planilha"
      End
   End
   Begin VB.Menu mnuSair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Activate()

    Dim dblLarguraPaineis2a5 As Double
    Dim dblLarguraPainel1    As Double
    
    With StatusBar
        'Calculando a largura do primeiro painel
        dblLarguraPainel1 = (Me.Width / 2)
        .Panels(1).Width = dblLarguraPainel1
        'Calculando as larguras dos outros painéis
        dblLarguraPaineis2a5 = ((Me.Width - dblLarguraPainel1) / 4)
        .Panels(2).Width = dblLarguraPaineis2a5
        .Panels(3).Width = dblLarguraPaineis2a5
        .Panels(4).Width = dblLarguraPaineis2a5
        .Panels(5).Width = dblLarguraPaineis2a5
        .Refresh
    End With
    
End Sub

Private Sub MDIForm_Load()

    On Error GoTo ErrMDIForm_Load
    
    Screen.MousePointer = vbHourglass
    
    'Abre conexao com o Banco SQL-Server
    Call gsAbrirConexaoBDLight
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrMDIForm_Load:
    
    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mPreencheListaPaginas: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
               
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Call gsFecharConexaoBDLight
End Sub

Private Sub mnuCadConfigFechamento_Click()
    frmCadConfigFechamento.Show vbModal
End Sub

Private Sub mnuCadPMOAuxUsuarios_Click()
    frmCadUsuarios.Show vbModal
End Sub

Private Sub mnuCargaPMOnline_Click()
    frmCargaPMOnline.Show vbModal
End Sub

Private Sub mnuImportacaoPlanilha_Click()
    frmImpArq.Show vbModal
End Sub

Private Sub mnuSair_Click()

    Unload Me
    
End Sub

