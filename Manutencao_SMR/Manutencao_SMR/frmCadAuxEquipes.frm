VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCadAuxEquipes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Equipes"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   Icon            =   "frmCadAuxEquipes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar tlbCadastroAuxEquipes 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Incluir"
            Object.Tag             =   "INCLUIR"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Alterar"
            Object.Tag             =   "ALTERAR"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Excluir"
            Object.Tag             =   "EXLCUIR"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Sair"
            Object.Tag             =   "SAIR"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraListaCadAuxEquipes 
      Caption         =   "Lista de equipes cadastradas"
      Height          =   2850
      Left            =   15
      TabIndex        =   2
      Top             =   525
      Width           =   10560
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8430
         Top             =   1035
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxEquipes.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxEquipes.frx":0554
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxEquipes.frx":0666
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxEquipes.frx":07D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxEquipes.frx":08E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxEquipes.frx":0D34
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwCadastroAuxEquipes 
         Height          =   2535
         Left            =   45
         TabIndex        =   3
         Top             =   255
         Width           =   10470
         _ExtentX        =   18468
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame fraDadosCadAuxEquipes 
      Caption         =   "Dados da equipe"
      Height          =   2835
      Left            =   15
      TabIndex        =   4
      Top             =   540
      Width           =   10575
      Begin VB.Frame Frame1 
         Height          =   885
         Left            =   120
         TabIndex        =   8
         Top             =   1830
         Width           =   10335
         Begin VB.CommandButton cmdGravar 
            Caption         =   "Gravar"
            Height          =   375
            Left            =   7950
            TabIndex        =   10
            Top             =   300
            Width           =   1020
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   9180
            TabIndex        =   9
            Top             =   300
            Width           =   1020
         End
      End
      Begin VB.TextBox txtSiglaEquipe 
         Height          =   345
         Left            =   2550
         MaxLength       =   50
         TabIndex        =   6
         Top             =   975
         Width           =   6675
      End
      Begin VB.TextBox txtDescricaoEquipe 
         Height          =   345
         Left            =   2565
         MaxLength       =   50
         TabIndex        =   1
         Top             =   495
         Width           =   6675
      End
      Begin VB.Label lblSiglaEquipe 
         AutoSize        =   -1  'True
         Caption         =   "Sigla da Equipe:"
         Height          =   195
         Left            =   945
         TabIndex        =   7
         Top             =   1050
         Width           =   1155
      End
      Begin VB.Label lblDescricaoEquipe 
         AutoSize        =   -1  'True
         Caption         =   "Descrição da Equipe:"
         Height          =   195
         Left            =   960
         TabIndex        =   5
         Top             =   570
         Width           =   1530
      End
   End
End
Attribute VB_Name = "frmCadAuxEquipes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub mDesabilitarBotoesToolBarEquipes()

    With tlbCadastroAuxEquipes
    
        .Buttons(1).Enabled = False
        .Buttons(2).Enabled = False
        .Buttons(3).Enabled = False
        
    End With
    
End Sub

Private Sub mHabilitarBotoesToolBarEquipes()

    With tlbCadastroAuxEquipes
    
        .Buttons(1).Enabled = True
        .Buttons(2).Enabled = True
        .Buttons(3).Enabled = True
        
    End With
    
End Sub

Private Sub mAlterarEquipe()

    Dim strSql            As String
    Dim lngLinhasAfetadas As Long
    
    On Error GoTo ErrmAlterarEquipe
    
    Screen.MousePointer = vbHourglass
    
    strSql = "SP_ALTERAR_EQUIPE_GVI " & CInt(lvwCadastroAuxEquipes.SelectedItem.Text) & ", " & _
             "'" & txtDescricaoEquipe.Text & "', '" & txtSiglaEquipe.Text & "'"
             
    dbConexaoSMR.Execute strSql, lngLinhasAfetadas
    
    If lngLinhasAfetadas > 0 Then
        MsgBox "Equipe alterada com sucesso ! ", vbInformation + vbOKOnly, App.Title
    End If
    
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxEquipes.Visible = False
    fraListaCadAuxEquipes.Visible = True
    
    'Alterando os dados na lista
    With lvwCadastroAuxEquipes.SelectedItem
        .SubItems(1) = txtDescricaoEquipe.Text
        .SubItems(2) = txtSiglaEquipe.Text
        .EnsureVisible
    End With
        
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmAlterarEquipe:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mAlterarEquipe: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mDadosEquipe()

    On Error GoTo ErrmDadosEquipe
    
    With lvwCadastroAuxEquipes.SelectedItem
        txtDescricaoEquipe.Text = .SubItems(1)
        txtSiglaEquipe.Text = .SubItems(2)
    End With
    
    Exit Sub
    
ErrmDadosEquipe:

    MsgBox "Ocorreu o seguinte erro na rotina mDadosEquipe: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub



Private Sub mExcluirEquipe()

    Dim strSql            As String
    Dim lngLinhasAfetadas As Long
    
    On Error GoTo ErrmExcluirEquipe
    
    Screen.MousePointer = vbHourglass
    
    strSql = "SP_EXCLUIR_EQUIPE_GVI " & CInt(lvwCadastroAuxEquipes.SelectedItem.Text)
             
    dbConexaoSMR.Execute strSql, lngLinhasAfetadas
    
    If lngLinhasAfetadas > 0 Then
        MsgBox "Equipe excluída com sucesso ! ", vbInformation + vbOKOnly, App.Title
    End If
    
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxEquipes.Visible = False
    fraListaCadAuxEquipes.Visible = True
    
    'Preencher lista
    Call mPreencheListaEquipes
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmExcluirEquipe:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mExcluirEquipe: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mHabilitarDesabilitarCamposCadastroEquipe(ByVal pAcao As String)

    fraDadosCadAuxEquipes.Tag = pAcao

    If UCase(Trim(pAcao)) = "I" Then
        
        txtDescricaoEquipe.Enabled = True
        txtSiglaEquipe.Enabled = True
        
        txtDescricaoEquipe.SetFocus
    
    ElseIf UCase(Trim(pAcao)) = "A" Then
        
        txtDescricaoEquipe.Enabled = True
        txtSiglaEquipe.Enabled = True
        
    End If
    
End Sub

Private Sub mIncluirEquipe()

    Dim strSql              As String
    Dim lngCodigoEquipe     As Long
    Dim lngLinhasAfetadas   As Long
    Dim itmEquipeInserida   As ListItem
    Dim cmdEquipe           As ADODB.Command
    
    On Error GoTo ErrmIncluirEquipe
    
    Screen.MousePointer = vbHourglass
    
    Set cmdEquipe = New ADODB.Command
    
    With cmdEquipe
    
        .ActiveConnection = dbConexaoSMR
        .CommandType = adCmdStoredProc
        .CommandText = "SP_INCLUIR_EQUIPE_GVI"
        
        .Parameters.Refresh
        .Parameters(1).Value = Trim(txtDescricaoEquipe.Text)
        .Parameters(2).Value = Trim(txtSiglaEquipe.Text)
        
        .Execute
        
        lngCodigoEquipe = .Parameters(3).Value
        
    End With
        
    If lngCodigoEquipe > 0 Then
        MsgBox "Equipe incluída com sucesso ! ", vbInformation + vbOKOnly, App.Title
    End If
    
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxEquipes.Visible = False
    fraListaCadAuxEquipes.Visible = True
    
    'Adicionando na lista
    Set itmEquipeInserida = lvwCadastroAuxEquipes.ListItems.Add(, , Format(lngCodigoEquipe, "000"))
    itmEquipeInserida.SubItems(1) = txtDescricaoEquipe.Text
    itmEquipeInserida.SubItems(2) = txtSiglaEquipe.Text
    
    itmEquipeInserida.Selected = True
    itmEquipeInserida.EnsureVisible
        
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmIncluirEquipe:
    
    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mIncluirEquipe: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mLimparDadosEquipe()

    txtDescricaoEquipe.Text = ""
    
End Sub


Private Sub mMontaCabecalhoListaEquipes()

    'Criando os cabeçalhos das colunas da lista
    With lvwCadastroAuxEquipes
        .ColumnHeaders.Add , , "Código", 1000
        .ColumnHeaders.Add , , "Descrição", 6870
        .ColumnHeaders.Add , , "Sigla", 2170
    End With
    
End Sub

Private Sub mPreencheListaEquipes()

    Dim rsEquipes    As ADODB.Recordset
    Dim itmEquipe    As ListItem
    Dim strSql       As String
    
    On Error GoTo ErrmPreencheListaEquipes
    
    Screen.MousePointer = vbHourglass
    
    With lvwCadastroAuxEquipes
    
        .ListItems.Clear
        .ColumnHeaders.Clear
        
        'Monta o cabeçalho da lista
        Call mMontaCabecalhoListaEquipes
        
        strSql = "EXECUTE SP_LISTAR_EQUIPES_GVI"
        
        Set rsEquipes = New ADODB.Recordset
        rsEquipes.Open strSql, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
        
        Do While Not rsEquipes.EOF
        
            Set itmEquipe = .ListItems.Add(, , Format(rsEquipes.Fields("Cod_Equipe").Value, "000"))
            itmEquipe.SubItems(1) = rsEquipes.Fields("Desc_Equipe").Value
            itmEquipe.SubItems(2) = rsEquipes.Fields("Sigla_Equipe").Value
              
            rsEquipes.MoveNext
            
        Loop
        
    End With
    
    rsEquipes.Close
    Set rsEquipes = Nothing
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmPreencheListaEquipes:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mPreencheListaEquipes: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
           
    If rsEquipes.State = adStateOpen Then
        rsEquipes.Close
        Set rsEquipes = Nothing
    End If

End Sub

Private Sub mTratarTelaAlterarEquipe()

    On Error GoTo ErrmTratarTelaAlterarEquipe
    
    'Desabilita os botões da ToolBar
    Call mDesabilitarBotoesToolBarEquipes
    
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadAuxEquipes.Visible = True
    fraListaCadAuxEquipes.Visible = False
        
    Exit Sub
    
ErrmTratarTelaAlterarEquipe:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaAlterarEquipe: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mTratarTelaExcluirEquipe()

    On Error GoTo ErrmTratarTelaExcluirEquipe
    
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadAuxEquipes.Visible = True
    fraListaCadAuxEquipes.Visible = False
    
    Exit Sub
    
ErrmTratarTelaExcluirEquipe:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaExcluirEquipe: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub


Private Sub mTratarTelaIncluirEquipe()

    On Error GoTo ErrmTratarTelaIncluirEquipe
    
    'Desabilita os botões da ToolBar
    Call mDesabilitarBotoesToolBarEquipes
    
    'Limpa dados
    Call mLimparDadosEquipe
    
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadAuxEquipes.Visible = True
    fraListaCadAuxEquipes.Visible = False
    
    Exit Sub
    
ErrmTratarTelaIncluirEquipe:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaIncluirEquipe: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
    
End Sub

Private Function mValidarCamposEquipe() As Boolean

    On Error GoTo ErrmValidarCamposEquipe
    
    If Trim(txtDescricaoEquipe.Text) = "" Then
        MsgBox "Descrição da equipe não foi preenchida !", vbInformation + vbOKOnly, App.Title
        txtDescricaoEquipe.SetFocus
        mValidarCamposEquipe = False
        Exit Function
    End If
        
    mValidarCamposEquipe = True
        
    Exit Function
    
ErrmValidarCamposEquipe:

    MsgBox "Ocorreu o seguinte erro na rotina mValidarCamposEquipe: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
    
End Function

Private Sub cmdCancelar_Click()

    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxEquipes.Visible = False
    fraListaCadAuxEquipes.Visible = True

    'Habilita os botões da ToolBar
    Call mHabilitarBotoesToolBarEquipes

End Sub

Private Sub cmdGravar_Click()
    
    On Error GoTo ErrcmdGravar_Click
    
    If mValidarCamposEquipe() Then
    
        If fraDadosCadAuxEquipes.Tag = "I" Then
        
            'Inclusao
            Call mIncluirEquipe
        
        ElseIf fraDadosCadAuxEquipes.Tag = "A" Then
            
            'Alteração
            Call mAlterarEquipe
        
        End If
        
        'Habilita os botões da ToolBar
        Call mHabilitarBotoesToolBarEquipes
        
    End If
    
    Exit Sub
    
ErrcmdGravar_Click:

    MsgBox "Ocorreu o seguinte erro na rotina cmdGravar_Click: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub Form_Load()
        
    'Habilita os botões da ToolBar
    Call mHabilitarBotoesToolBarEquipes
        
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxEquipes.Visible = False
    fraListaCadAuxEquipes.Visible = True

    'Preenche a lista de Equipes
    Call mPreencheListaEquipes
    
End Sub









Private Sub Form_Unload(Cancel As Integer)

    frmPrincipal.StatusBar.Panels(1).Text = ""
    
End Sub


Private Sub lvwCadastroAuxEquipes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    With lvwCadastroAuxEquipes
    
        .SortKey = ColumnHeader.Index - 1
        
        If .SortOrder = lvwAscending Then
           .SortOrder = lvwDescending
        Else
           .SortOrder = lvwAscending
        End If
        
        .Sorted = True
        
    End With

End Sub


Private Sub tlbCadastroAuxEquipes_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
    
        Case 1 'Incluir
        
            Call mTratarTelaIncluirEquipe
            Call mHabilitarDesabilitarCamposCadastroEquipe("I")
                    
        Case 2 'Alterar
        
            Call mTratarTelaAlterarEquipe
            Call mDadosEquipe
            Call mHabilitarDesabilitarCamposCadastroEquipe("A")
                        
        Case 3 'Excluir
        
            If MsgBox("Deseja realmente excluir este registro ?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                Call mExcluirEquipe
            End If
            
        Case 5 'Sair
            
            Unload Me
            
    End Select

End Sub


