VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCadAuxUnidade 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Unidades"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   Icon            =   "frmCadAuxUnidade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar tlbCadastroAuxUnidade 
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
   Begin VB.Frame fraListaCadAuxUnidade 
      Caption         =   "Lista de unidades cadastradas"
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
               Picture         =   "frmCadAuxUnidade.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxUnidade.frx":0554
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxUnidade.frx":0666
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxUnidade.frx":07D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxUnidade.frx":08E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxUnidade.frx":0D34
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwCadastroAuxUnidade 
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
   Begin VB.Frame fraDadosCadAuxUnidade 
      Caption         =   "Dados da unidade"
      Height          =   2835
      Left            =   15
      TabIndex        =   4
      Top             =   540
      Width           =   10575
      Begin VB.Frame Frame1 
         Height          =   915
         Left            =   120
         TabIndex        =   6
         Top             =   1785
         Width           =   10335
         Begin VB.CommandButton cmdGravar 
            Caption         =   "Gravar"
            Height          =   375
            Left            =   7875
            TabIndex        =   8
            Top             =   315
            Width           =   1020
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   9105
            TabIndex        =   7
            Top             =   315
            Width           =   1020
         End
      End
      Begin VB.TextBox txtDescricaoUnidade 
         Height          =   345
         Left            =   2925
         MaxLength       =   50
         TabIndex        =   1
         Top             =   495
         Width           =   6675
      End
      Begin VB.Label lblDescricaoUnidade 
         AutoSize        =   -1  'True
         Caption         =   "Descrição da Unidade:"
         Height          =   195
         Left            =   945
         TabIndex        =   5
         Top             =   570
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmCadAuxUnidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub mDesabilitarBotoesToolBarUnidade()

    With tlbCadastroAuxUnidade
    
        .Buttons(1).Enabled = False
        .Buttons(2).Enabled = False
        .Buttons(3).Enabled = False
        
    End With
    
End Sub

Private Sub mHabilitarBotoesToolBarUnidade()

    With tlbCadastroAuxUnidade
    
        .Buttons(1).Enabled = True
        .Buttons(2).Enabled = True
        .Buttons(3).Enabled = True
        
    End With
    
End Sub

Private Sub mAlterarUnidade()

    Dim strSql            As String
    Dim lngLinhasAfetadas As Long
    
    On Error GoTo ErrmAlterarUnidade
    
    Screen.MousePointer = vbHourglass
    
    strSql = "SP_ALTERAR_Unidade " & CInt(lvwCadastroAuxUnidade.SelectedItem.Text) & ", " & _
             "'" & txtDescricaoUnidade.Text & "'"
             
    dbConexaoSMR.Execute strSql, lngLinhasAfetadas
    
    If lngLinhasAfetadas > 0 Then
        MsgBox "Unidade alterada com sucesso ! ", vbInformation + vbOKOnly, App.Title
    End If
    
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxUnidade.Visible = False
    fraListaCadAuxUnidade.Visible = True
    
    'Alterando os dados na lista
    With lvwCadastroAuxUnidade.SelectedItem
        .SubItems(1) = txtDescricaoUnidade.Text
        .EnsureVisible
    End With
        
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmAlterarUnidade:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mAlterarUnidade: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mDadosUnidade()

    On Error GoTo ErrmDadosUnidade
    
    With lvwCadastroAuxUnidade.SelectedItem
        txtDescricaoUnidade.Text = .SubItems(1)
    End With
    
    Exit Sub
    
ErrmDadosUnidade:

    MsgBox "Ocorreu o seguinte erro na rotina mDadosUnidade: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub



Private Sub mExcluirUnidade()

    Dim strSql            As String
    Dim lngLinhasAfetadas As Long
    
    On Error GoTo ErrmExcluirUnidade
    
    Screen.MousePointer = vbHourglass
    
    strSql = "SP_EXCLUIR_Unidade " & CInt(lvwCadastroAuxUnidade.SelectedItem.Text)
             
    dbConexaoSMR.Execute strSql, lngLinhasAfetadas
    
    If lngLinhasAfetadas > 0 Then
        MsgBox "Unidade excluída com sucesso ! ", vbInformation + vbOKOnly, App.Title
    End If
    
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxUnidade.Visible = False
    fraListaCadAuxUnidade.Visible = True
    
    'Preencher lista
    Call mPreencheListaUnidade
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmExcluirUnidade:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mExcluirUnidade: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mHabilitarDesabilitarCamposCadastroUnidade(ByVal pAcao As String)

    fraDadosCadAuxUnidade.Tag = pAcao

    If UCase(Trim(pAcao)) = "I" Then
        
        txtDescricaoUnidade.Enabled = True
        
        txtDescricaoUnidade.SetFocus
    
    ElseIf UCase(Trim(pAcao)) = "A" Then
        
        txtDescricaoUnidade.Enabled = True
        
    End If
    
End Sub

Private Sub mIncluirUnidade()

    Dim strSql                  As String
    Dim lngCodigoUnidade     As Long
    Dim lngLinhasAfetadas       As Long
    Dim itmUnidadeInserida   As ListItem
    Dim cmdUnidade           As ADODB.Command
    
    On Error GoTo ErrmIncluirUnidade
    
    Screen.MousePointer = vbHourglass
    
    Set cmdUnidade = New ADODB.Command
    
    With cmdUnidade
    
        .ActiveConnection = dbConexaoSMR
        .CommandType = adCmdStoredProc
        .CommandText = "SP_INCLUIR_Unidade"
        
        .Parameters.Refresh
        .Parameters(1).Value = Trim(txtDescricaoUnidade.Text)
        
        .Execute
        
        lngCodigoUnidade = .Parameters(2).Value
        
    End With
        
    If lngCodigoUnidade > 0 Then
        MsgBox "Unidade incluída com sucesso ! ", vbInformation + vbOKOnly, App.Title
    End If
    
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxUnidade.Visible = False
    fraListaCadAuxUnidade.Visible = True
    
    'Adicionando na lista
    Set itmUnidadeInserida = lvwCadastroAuxUnidade.ListItems.Add(, , Format(lngCodigoUnidade, "000"))
    itmUnidadeInserida.SubItems(1) = txtDescricaoUnidade.Text
    
    itmUnidadeInserida.Selected = True
    itmUnidadeInserida.EnsureVisible
        
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmIncluirUnidade:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mIncluirUnidade: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mLimparDadosUnidade()

    txtDescricaoUnidade.Text = ""
    
End Sub


Private Sub mMontaCabecalhoListaUnidade()

    'Criando os cabeçalhos das colunas da lista
    With lvwCadastroAuxUnidade
        .ColumnHeaders.Add , , "Código", 1000
        .ColumnHeaders.Add , , "Descrição", 9370
    End With
    
End Sub

Private Sub mPreencheListaUnidade()

    Dim rsUnidade   As ADODB.Recordset
    Dim itmUnidade   As ListItem
    Dim strSql          As String
    
    On Error GoTo ErrmPreencheListaUnidade
    
    Screen.MousePointer = vbHourglass
    
    With lvwCadastroAuxUnidade
    
        .ListItems.Clear
        .ColumnHeaders.Clear
        
        'Monta o cabeçalho da lista
        Call mMontaCabecalhoListaUnidade
        
        strSql = "EXECUTE SP_LISTAR_Unidade"
        
        Set rsUnidade = New ADODB.Recordset
        rsUnidade.Open strSql, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
        
        Do While Not rsUnidade.EOF
        
            Set itmUnidade = .ListItems.Add(, , Format(rsUnidade.Fields("Cod_Unidade").Value, "000"))
            itmUnidade.SubItems(1) = rsUnidade.Fields("Desc_Unidade").Value
              
            rsUnidade.MoveNext
            
        Loop
        
    End With
    
    rsUnidade.Close
    Set rsUnidade = Nothing
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmPreencheListaUnidade:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mPreencheListaUnidade: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
           
    If rsUnidade.State = adStateOpen Then
        rsUnidade.Close
        Set rsUnidade = Nothing
    End If

End Sub

Private Sub mTratarTelaAlterarUnidade()

    On Error GoTo ErrmTratarTelaAlterarUnidade
    
    'Desabilita os botões da ToolBar
    Call mDesabilitarBotoesToolBarUnidade
    
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadAuxUnidade.Visible = True
    fraListaCadAuxUnidade.Visible = False
        
    Exit Sub
    
ErrmTratarTelaAlterarUnidade:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaAlterarUnidade: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mTratarTelaExcluirUnidade()

    On Error GoTo ErrmTratarTelaExcluirUnidade
    
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadAuxUnidade.Visible = True
    fraListaCadAuxUnidade.Visible = False
    
    Exit Sub
    
ErrmTratarTelaExcluirUnidade:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaExcluirUnidade: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub


Private Sub mTratarTelaIncluirUnidade()

    On Error GoTo ErrmTratarTelaIncluirUnidade
    
    'Desabilita os botões da ToolBar
    Call mDesabilitarBotoesToolBarUnidade
    
    'Limpa dados
    Call mLimparDadosUnidade
    
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadAuxUnidade.Visible = True
    fraListaCadAuxUnidade.Visible = False
    
    Exit Sub
    
ErrmTratarTelaIncluirUnidade:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaIncluirUnidade: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
    
End Sub

Private Function mValidarCamposUnidade() As Boolean

    On Error GoTo ErrmValidarCamposUnidade
    
    If Trim(txtDescricaoUnidade.Text) = "" Then
        MsgBox "Descrição da Unidade não foi preenchida !", vbInformation + vbOKOnly, App.Title
        txtDescricaoUnidade.SetFocus
        mValidarCamposUnidade = False
        Exit Function
    End If
        
    mValidarCamposUnidade = True
        
    Exit Function
    
ErrmValidarCamposUnidade:

    MsgBox "Ocorreu o seguinte erro na rotina mValidarCamposUnidade: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
    
End Function

Private Sub cmdCancelar_Click()

    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxUnidade.Visible = False
    fraListaCadAuxUnidade.Visible = True

    'Habilita os botões da ToolBar
    Call mHabilitarBotoesToolBarUnidade

End Sub

Private Sub cmdGravar_Click()
    
    On Error GoTo ErrcmdGravar_Click
    
    If mValidarCamposUnidade() Then
    
        If fraDadosCadAuxUnidade.Tag = "I" Then
        
            'Inclusao
            Call mIncluirUnidade
        
        ElseIf fraDadosCadAuxUnidade.Tag = "A" Then
            
            'Alteração
            Call mAlterarUnidade
        
        End If
        
        'Habilita os botões da ToolBar
        Call mHabilitarBotoesToolBarUnidade
        
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
    Call mHabilitarBotoesToolBarUnidade
        
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxUnidade.Visible = False
    fraListaCadAuxUnidade.Visible = True

    'Preenche a lista
    Call mPreencheListaUnidade
    
End Sub









Private Sub Form_Unload(Cancel As Integer)

    frmPrincipal.StatusBar.Panels(1).Text = ""
    
End Sub


Private Sub lvwCadastroAuxUnidade_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    With lvwCadastroAuxUnidade
    
        .SortKey = ColumnHeader.Index - 1
        
        If .SortOrder = lvwAscending Then
           .SortOrder = lvwDescending
        Else
           .SortOrder = lvwAscending
        End If
        
        .Sorted = True
        
    End With

End Sub


Private Sub tlbCadastroAuxUnidade_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
    
        Case 1 'Incluir
        
            Call mTratarTelaIncluirUnidade
            Call mHabilitarDesabilitarCamposCadastroUnidade("I")
                    
        Case 2 'Alterar
        
            Call mTratarTelaAlterarUnidade
            Call mDadosUnidade
            Call mHabilitarDesabilitarCamposCadastroUnidade("A")
                        
        Case 3 'Excluir
        
            If MsgBox("Deseja realmente excluir este registro ?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                Call mExcluirUnidade
            End If
            
        Case 5 'Sair
            
            Unload Me
            
    End Select

End Sub


