VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "Módulo de Manutenção do Banco de Dados PMO"
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
            TextSave        =   "9/9/2004"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "16:14"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "Cadastros"
      Begin VB.Menu mnuCadastrosDePara 
         Caption         =   "De/Para "
         Begin VB.Menu mnuCadastrosAtualizaSanSin 
            Caption         =   "Atualiza Saneamento -> Sinergia"
         End
         Begin VB.Menu mnuCadastrosAtualiza_EP_SMR 
            Caption         =   "Atualiza EP -> SMR"
         End
         Begin VB.Menu mnuCadastrosAtualizaPercDt_TIN_SMR 
            Caption         =   "Atualiza % Data TIN -> SMR"
         End
         Begin VB.Menu mnuCadastrosAtualizaCases_TIN_SMR 
            Caption         =   "Atualiza Cases TIN -> SMR"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuCadastrosAtualizaPerc_TCP_SMR 
            Caption         =   "Atualiza % TCP -> SMR"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuCadPMOAuxUsuarios 
         Caption         =   "Usuários"
      End
      Begin VB.Menu mnuCadUsuariosCutover 
         Caption         =   "Usuários Cutover"
      End
      Begin VB.Menu mnuCadPMOProjetos_PMO 
         Caption         =   "Projetos PMO"
      End
      Begin VB.Menu mnuCadPMOAuxFrentes 
         Caption         =   "Frentes"
      End
      Begin VB.Menu mnuCadPMOAuxEquipes 
         Caption         =   "Equipes"
      End
      Begin VB.Menu mnuCadPMOAuxGovernancas 
         Caption         =   "Governanças"
      End
      Begin VB.Menu mnuCadPMOAuxArea 
         Caption         =   "Areas"
      End
      Begin VB.Menu mnuCadPMOAuxUnidade 
         Caption         =   "Unidades"
      End
      Begin VB.Menu mnuCadConfigFechamento 
         Caption         =   "Configuração do Fechamento"
      End
      Begin VB.Menu mnuCadConfigFechamentoCutover 
         Caption         =   "Configuração do Fechamento Cutover"
      End
   End
   Begin VB.Menu mnuTransmissoes 
      Caption         =   "Transmissões"
      Begin VB.Menu mnuTransFerrH 
         Caption         =   "Ferramenta H"
         Begin VB.Menu mnuTransFerrHExportar 
            Caption         =   "Exportar"
         End
         Begin VB.Menu mnuTransFerrHImportar 
            Caption         =   "Importar"
         End
         Begin VB.Menu mnuTransFerrHImportarTesteCarga 
            Caption         =   "Importar Teste Carga"
         End
      End
      Begin VB.Menu mnuTransCarga 
         Caption         =   "Carga"
         Begin VB.Menu mnuTransCargaPMOnline 
            Caption         =   "PMOnline"
         End
         Begin VB.Menu mnuTransCargaIssuesPMOnline 
            Caption         =   "Issues PMOnline"
         End
      End
      Begin VB.Menu mnuTransCargaSMR 
         Caption         =   "Carga SMR"
         Begin VB.Menu mnuAtualizaPercSanSin 
            Caption         =   "Atualiza % Saneamento -> Sinergia"
         End
         Begin VB.Menu mnuAtualizaPercEPSMR 
            Caption         =   "Atualiza % . EP -> SMR."
         End
         Begin VB.Menu mnuAtualizaPercSCSMR 
            Caption         =   "Atualiza % . SC -> SMR."
         End
         Begin VB.Menu mnuAtualizaPercDtTinSin 
            Caption         =   "Atualiza % e Datas Teste Integrado -> Sinergia"
         End
         Begin VB.Menu mnuAtualizaFlagTinSin 
            Caption         =   "Atualiza Flag Teste Integrado -> Sinergia"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuAtualizaPercSMR_SMRTRAB 
            Caption         =   "Atualiza % SMR -> SMR Trabalho"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuAtualizaPercTcpSin 
            Caption         =   "Atualiza % Teste Campo -> Sinergia"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuAtualizaPercCutSin 
            Caption         =   "Atualiza % CutOver -> Sinergia"
         End
         Begin VB.Menu mnuAtualizaExclusaoTarefas 
            Caption         =   "Atualiza Exclusão de Tarefas"
         End
      End
      Begin VB.Menu mnuCargaTesteIntegrado 
         Caption         =   "Carga Teste Integrado"
         Begin VB.Menu mnuImportarPlanilhaDesenv 
            Caption         =   "Importar Planilha Desenvolvimento"
         End
         Begin VB.Menu mnuImportarPlanilha 
            Caption         =   "Importar Planilha %"
         End
         Begin VB.Menu mnuImportarPlanilhaCases 
            Caption         =   "Importar Planilha % Cases"
         End
         Begin VB.Menu mnuAtualizaProjeto 
            Caption         =   "Atualiza Projeto"
         End
         Begin VB.Menu mnuLimparFlagInclusao 
            Caption         =   "Limpa Flag Inclusão"
         End
      End
      Begin VB.Menu mnuCargaCutOver 
         Caption         =   "Carga Cut Over"
         Begin VB.Menu mnuAtualizaOperacoesCutOver 
            Caption         =   "Atualiza Operações Cut Over"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuExcluirOperacoesCutOver 
            Caption         =   "Excluir Operações Cut Over"
         End
      End
      Begin VB.Menu mnuCargaMatDidatico 
         Caption         =   "Carga Material Didático"
         Begin VB.Menu mnuImportarPlanilhaMatDid 
            Caption         =   "Importar Planilha %"
         End
      End
      Begin VB.Menu mnuTransCriacaoPlanilhas 
         Caption         =   "Criação Planilhas"
         Enabled         =   0   'False
         Begin VB.Menu mnuTransCriacaoPlanilhasCausasAtrasadas 
            Caption         =   "Causas Atrasadas"
            Enabled         =   0   'False
         End
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
    Call gsAbrirConexaoBDTIN
    
    'Abre conexao com o Banco SQL-Server
    Call gsAbrirConexaoBDTCP
    
    'Abre conexao com o Banco SQL-Server
    Call gsAbrirConexaoBDCUT

    'Abre conexao com o Banco SQL-Server
    Call gsAbrirConexaoBDEP
    
    'Abre conexao com o Banco SQL-Server
    Call gsAbrirConexaoBDSMR
    
    'Abre conexao com o Banco SQL-Server
    Call gsAbrirConexaoBDSMRTRAB
    
    'Abre conexao com o Banco SQL-Server
    Call gsAbrirConexaoBDSC
    
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

Private Sub mnuAtualizaExclusaoTarefas_Click()
    frmAtualizarExclusaoTarefas.Show vbModal
End Sub

Private Sub mnuAtualizaFlagTinSin_Click()

    frmAtualizarCases_TesteInt_Sin.Show vbModal

End Sub

Private Sub mnuAtualizaOperacoesCutOver_Click()
    frmAtualizarOperacoesCutOver.Show vbModal
End Sub

Private Sub mnuAtualizaPercCutSin_Click()

    frmAtualizarPerc_CutOver_Sin.Show vbModal

End Sub

Private Sub mnuAtualizaPercDtTinSin_Click()
    
    frmAtualizarPercDt_TesteInt_Sin.Show vbModal

End Sub


Private Sub mnuAtualizaPercEPSMR_Click()
    frmAtualizarPerc_EP_SMR.Show vbModal
End Sub

Private Sub mnuAtualizaPercSanSin_Click()
    frmAtualizarPercSanSin.Show vbModal
End Sub

Private Sub mnuAtualizaPercSCSMR_Click()
    frmAtualizarPerc_SC_SMR.Show vbModal
End Sub

Private Sub mnuAtualizaPercSMR_SMRTRAB_Click()
    frmAtualizarPerc_SMR_SMRTRAB.Show vbModal
End Sub

Private Sub mnuAtualizaPercTcpSin_Click()
    frmAtualizarPerc_TesteCampo_Sin.Show vbModal
End Sub

Private Sub mnuAtualizaProjeto_Click()

    Dim objAtualiza     As Atualizador.cAtualizador
    Dim intProjeto      As Integer
    Dim strProjeto      As String
    Dim blnSucesso      As String

    Set objAtualiza = New cAtualizador

    Screen.MousePointer = vbHourglass
    blnSucesso = objAtualiza.ExportaDados(intProjeto, strProjeto)
    Screen.MousePointer = vbDefault

End Sub


Private Sub mnuCadastrosAtualiza_EP_SMR_Click()
    frmCadAtualizaPerc_EP_SMR.Show vbModal
End Sub

Private Sub mnuCadastrosAtualizaCases_TIN_SMR_Click()
    frmCadAtualizaCases_TesteInt_Sin.Show vbModal
End Sub

Private Sub mnuCadastrosAtualizaPerc_TCP_SMR_Click()
    frmCadAtualizaPerc_TesteCampo_Sin.Show vbModal
End Sub

Private Sub mnuCadastrosAtualizaPercDt_TIN_SMR_Click()
    frmCadAtualizaPercDt_TesteInt_Sin.Show vbModal
End Sub

Private Sub mnuCadastrosAtualizaSanSin_Click()
    frmCadAtualizarPercSanSin.Show vbModal
End Sub

Private Sub mnuCadConfigFechamento_Click()
    frmCadConfigFechamento.Show vbModal
End Sub

Private Sub mnuCadConfigFechamentoCutover_Click()
    frmCadConfigFechamento_CUT.Show vbModal
End Sub

Private Sub mnuCadPMOAuxArea_Click()
    
    StatusBar.Panels(1).Text = Me.mnuCadastros.Caption & " -> " & mnuCadPMOAuxArea.Caption
    frmCadAuxArea.Show vbModal

End Sub

Private Sub mnuCadPMOAuxEquipes_Click()

    StatusBar.Panels(1).Text = Me.mnuCadastros.Caption & " -> " & mnuCadPMOAuxEquipes.Caption
    frmCadAuxEquipes.Show vbModal
    
End Sub

Private Sub mnuCadPMOAuxFrentes_Click()

    StatusBar.Panels(1).Text = Me.mnuCadastros.Caption & " -> " & mnuCadPMOAuxFrentes.Caption
    frmCadAuxFrentes.Show vbModal
    
End Sub

Private Sub mnuCadPMOAuxGovernancas_Click()

    StatusBar.Panels(1).Text = Me.mnuCadastros.Caption & " -> " & mnuCadPMOAuxGovernancas.Caption
    frmCadAuxGovernancas.Show vbModal
    
End Sub

Private Sub mnuCadPMOAuxUnidade_Click()
    
    StatusBar.Panels(1).Text = Me.mnuCadastros.Caption & " -> " & mnuCadPMOAuxUnidade.Caption
    frmCadAuxUnidade.Show vbModal

End Sub

Private Sub mnuCadPMOAuxUsuarios_Click()

    StatusBar.Panels(1).Text = Me.mnuCadastros.Caption & " -> " & mnuCadPMOAuxUsuarios.Caption
    frmCadAuxUsuarios.Show vbModal
    
End Sub



Private Sub mnuCadPMOProjetos_PMO_Click()

    StatusBar.Panels(1).Text = Me.mnuCadastros.Caption & " -> " & mnuCadPMOProjetos_PMO.Caption
    frmCadProjetosPMO.Show vbModal
    
End Sub


Private Sub mnuCadUsuariosCutover_Click()
    frmCadUsuarios_CUT.Show vbModal
End Sub

Private Sub mnuExcluirOperacoesCutOver_Click()
Dim strSql As String
On Error GoTo ErrExcluirOperacoesCutOver

    strSql = "EXEC SP_EXCLUIR_OPERACOES_CUTOVER"
                    
    dbConexaoCUT.Execute strSql

    Call MsgBox("Rotina executada com sucesso.")
    
Exit Sub

ErrExcluirOperacoesCutOver:

    Call MsgBox(Err.Description)
    
    
    
End Sub

Private Sub mnuImportarPlanilha_Click()
    frmImpArq_TesteIntegr.Show vbModal
End Sub

Private Sub mnuImportarPlanilhaCases_Click()
    frmImpArq_TesteIntegr_Case.Show vbModal
End Sub

Private Sub mnuImportarPlanilhaDesenv_Click()
    frmImpArq_Desenv_TesteIntegr.Show vbModal
End Sub

Private Sub mnuImportarPlanilhaMatDid_Click()
    frmImpArq_MatDidatico.Show vbModal
End Sub

Private Sub mnuLimparFlagInclusao_Click()
Dim strSql As String
On Error GoTo ErrFlagInclusao

    strSql = "EXEC SP_ATUALIZAR_FLAG_INCLUSAO"
                    
    dbConexaoTIN.Execute strSql

    Call MsgBox("Rotina executada com sucesso.")
    
Exit Sub

ErrFlagInclusao:

    Call MsgBox(Err.Description)

End Sub

Private Sub mnuSair_Click()

    Call gsFecharConexaoBDSMR
    Call gsFecharConexaoBDSMRTRAB
    Call gsFecharConexaoBDTIN
    Call gsFecharConexaoBDEP
    Call gsFecharConexaoBDTCP
    Call gsFecharConexaoBDCUT
    Call gsFecharConexaoBDSC
    
    End
    
End Sub


Private Sub mnuTransCargaIssuesPMOnline_Click()
    
    frmCargaIssuesPMOnline.Show vbModal

End Sub

Private Sub mnuTransCargaPMOnline_Click()

    frmCargaPMOnline.Show vbModal
    
End Sub

Private Sub mnuTransCriacaoPlanilhasCausasAtrasadas_Click()

'    frmExpCausasAtrasadas.Show vbModal
    
End Sub

Private Sub mnuTransFerrHExportar_Click()

    StatusBar.Panels(1).Text = Me.mnuTransmissoes.Caption & " -> " & Me.mnuTransFerrH.Caption & " -> " & Me.mnuTransFerrHExportar.Caption
    frmExpArq_H.Show vbModal
    
End Sub


Private Sub mnuTransFerrHImportar_Click()

    StatusBar.Panels(1).Text = Me.mnuTransmissoes.Caption & " -> " & Me.mnuTransFerrH.Caption & " -> " & Me.mnuTransFerrHImportar.Caption
    frmImpArq_H.Show vbModal
    
End Sub


Private Sub mnuTransFerrHImportarTesteCarga_Click()
    frmImpArq_H_Teste_Carga.Show vbModal
End Sub
